import logging
import secrets
from datetime import datetime
from pathlib import Path
from typing import Annotated

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import inspect, select, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session, selectinload
from starlette.middleware.sessions import SessionMiddleware

from app.config import BASE_DIR, settings, setup_logging
from app.database import Base, SessionLocal, engine, get_db
from app.models import AuditLog, InvitationCode, MonthlyConfirmation, Person, Role, StoredFile, User
from app.security import get_current_user, hash_password, verify_password
from app.services.excel import DEFAULT_EXPORT_FIELDS, FIELD_DEFINITIONS, create_template_xlsx, export_people_to_xlsx, import_people_from_xlsx
from app.services.ledger import (
    ACTIVE_STATUS,
    LEFT_STATUS,
    _all_people_query,
    annual_chart_data,
    current_period,
    dashboard_summary_from_people,
    confirmation_ids_for,
    normalize_period,
    normalize_status,
    period_add,
    profile_summary,
    valid_person_name,
    visible_people_as_of,
    year_periods,
)
from app.services.storage import ensure_storage_dirs, save_upload
from app.services.word import export_people_to_docx

setup_logging()
logger = logging.getLogger(__name__)

app = FastAPI(title=settings.app_name)
app.add_middleware(SessionMiddleware, secret_key=settings.secret_key, same_site="lax", https_only=settings.app_env == "production")
app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "app" / "templates")


def ensure_schema() -> None:
    Base.metadata.create_all(bind=engine)
    inspector = inspect(engine)
    table_columns = {table: {column["name"] for column in inspector.get_columns(table)} for table in inspector.get_table_names()}
    with engine.begin() as conn:
        people_columns = table_columns.get("people", set())
        if "monthly_confirmed" not in people_columns:
            conn.execute(text("ALTER TABLE people ADD COLUMN monthly_confirmed BOOLEAN DEFAULT 0 NOT NULL"))
        if "confirmed_at" not in people_columns:
            conn.execute(text("ALTER TABLE people ADD COLUMN confirmed_at DATETIME"))
        if "sfid" not in people_columns:
            conn.execute(text("ALTER TABLE people ADD COLUMN sfid VARCHAR(100)"))
        for column in [
            "sfid_expires_at", "disability_cert_id", "cert_issued_at", "household_address",
            "household_type", "contact_phone", "emergency_contact", "emergency_phone",
            "emergency_relation", "education", "marital_status", "bank_card_id", "bank_name",
            "disability_part", "disability_reason",
        ]:
            if column not in people_columns:
                conn.execute(text(f"ALTER TABLE people ADD COLUMN {column} VARCHAR(255)"))
        if "extra_fields" not in people_columns:
            conn.execute(text("ALTER TABLE people ADD COLUMN extra_fields TEXT"))
        user_columns = table_columns.get("users", set())
        if "default_service_fee" not in user_columns:
            conn.execute(text("ALTER TABLE users ADD COLUMN default_service_fee FLOAT DEFAULT 800"))
        if "invitation_code" not in user_columns:
            conn.execute(text("ALTER TABLE users ADD COLUMN invitation_code VARCHAR(64)"))


@app.on_event("startup")
def on_startup() -> None:
    ensure_storage_dirs()
    ensure_schema()
    with SessionLocal() as db:
        grandmaster = db.scalar(select(User).where(User.role == Role.GRANDMASTER.value))
        if not grandmaster:
            db.add(User(
                username=settings.grandmaster_username,
                display_name="总舵主",
                password_hash=hash_password(settings.grandmaster_password),
                role=Role.GRANDMASTER.value,
                default_service_fee=800,
            ))
            db.commit()
            logger.warning("已创建唯一总舵主账号：%s，请上线后立即修改默认密码", settings.grandmaster_username)


def render(request: Request, name: str, context: dict | None = None) -> HTMLResponse:
    payload = {
        "request": request,
        "app_name": settings.app_name,
        "message": request.query_params.get("message"),
        "error": request.query_params.get("error"),
    }
    payload.update(context or {})
    return templates.TemplateResponse(name, payload)


def redirect(url: str, **params) -> RedirectResponse:
    clean_params = {key: value for key, value in params.items() if value is not None}
    if clean_params:
        from urllib.parse import urlencode
        url = f"{url}?{urlencode(clean_params, doseq=True)}"
    return RedirectResponse(url, status_code=303)


def parse_optional_int(value: object | None) -> int | None:
    if value is None:
        return None
    text_value = str(value).strip()
    if not text_value or text_value.lower() == "none":
        return None
    return int(text_value)


def can_access_person(user: User, person: Person) -> bool:
    return user.role == Role.GRANDMASTER.value or person.owner_id == user.id


def require_grandmaster(user: User) -> None:
    if user.role != Role.GRANDMASTER.value:
        raise HTTPException(status_code=403, detail="只有总舵主可以操作")


def write_audit(db: Session, actor: User | None, action: str, target_type: str | None = None, target_id: int | None = None, detail: str | None = None) -> None:
    db.add(AuditLog(
        actor_id=actor.id if actor else None,
        action=action,
        target_type=target_type,
        target_id=target_id,
        detail=detail,
    ))


def distributors_for(db: Session, user: User):
    if user.role != Role.GRANDMASTER.value:
        return []
    return db.scalars(select(User).where(User.role == Role.DISTRIBUTOR.value).order_by(User.display_name)).all()


def owner_options_for(db: Session, user: User, owner_id: int | None = None) -> list[User]:
    if user.role != Role.GRANDMASTER.value:
        return [user]
    if owner_id:
        owner = db.get(User, owner_id)
        return [owner] if owner else []
    return [user, *distributors_for(db, user)]


def grouped_people_for_owners(owners: list[User], people: list[Person], confirmed_ids: set[int] | None = None, actor: User | None = None) -> list[dict]:
    confirmed_ids = confirmed_ids or set()
    groups = []
    for owner in owners:
        owner_people = [person for person in people if person.owner_id == owner.id]
        groups.append({
            "owner": owner,
            "label": "总舵组" if owner.role == Role.GRANDMASTER.value else owner.display_name,
            "people": owner_people,
            "active_people": [person for person in owner_people if normalize_status(person.employment_status) == ACTIVE_STATUS],
            "left_people": [person for person in owner_people if normalize_status(person.employment_status) == LEFT_STATUS],
            "summary": dashboard_summary_from_people(owner_people, confirmed_ids),
            "can_confirm": (actor.role != Role.GRANDMASTER.value if actor else True) or (actor is not None and owner.id == actor.id),
        })
    return groups


def duplicate_identifier(db: Session, field: str, value: str | None, exclude_id: int | None = None) -> bool:
    if not value:
        return False
    stmt = select(Person).where(getattr(Person, field) == value)
    if exclude_id:
        stmt = stmt.where(Person.id != exclude_id)
    return db.scalar(stmt) is not None


@app.get("/", response_class=HTMLResponse)
def home(request: Request, db: Annotated[Session, Depends(get_db)], user: Annotated[User, Depends(get_current_user)]):
    return dashboard(request, db, user)


@app.get("/auth/login", response_class=HTMLResponse)
def login_page(request: Request):
    return render(request, "login.html")


@app.post("/auth/login")
def login(request: Request, db: Annotated[Session, Depends(get_db)], username: str = Form(...), password: str = Form(...)):
    user = db.scalar(select(User).where(User.username == username))
    if not user or not verify_password(password, user.password_hash):
        return redirect("/auth/login", error="账号或密码不正确")
    if not user.is_active:
        return redirect("/auth/login", error="账号已停用")
    request.session.clear()
    request.session["user_id"] = user.id
    return redirect("/dashboard")


@app.get("/auth/register", response_class=HTMLResponse)
def register_page(request: Request, invite: str | None = None):
    return render(request, "register.html", {"invite": invite or ""})


@app.post("/auth/register")
def register(
    db: Annotated[Session, Depends(get_db)],
    username: str = Form(...),
    display_name: str = Form(...),
    password: str = Form(...),
    invite_code: str = Form(...),
):
    if len(password) < 8:
        return redirect("/auth/register", invite=invite_code, error="密码至少 8 位")
    code_text = invite_code.strip()
    invitation = db.scalar(select(InvitationCode).where(
        InvitationCode.code == code_text,
        InvitationCode.is_active.is_(True),
        InvitationCode.deleted_at.is_(None),
    ))
    if not invitation:
        return redirect("/auth/register", invite=code_text, error="邀请码无效或已作废")
    user = User(
        username=username.strip(),
        display_name=display_name.strip(),
        password_hash=hash_password(password),
        role=Role.DISTRIBUTOR.value,
        default_service_fee=800,
        invitation_code=invitation.code,
    )
    db.add(user)
    try:
        db.flush()
        if invitation.used_by_id is None:
            invitation.used_by_id = user.id
            invitation.used_at = datetime.utcnow()
        write_audit(db, user, "distributor.register", "user", user.id, f"邀请码 {invitation.code}")
        db.commit()
    except IntegrityError:
        db.rollback()
        return redirect("/auth/register", invite=code_text, error="用户名已存在")
    return redirect("/auth/login", message="注册成功，请登录")


@app.post("/auth/logout")
def logout(request: Request):
    request.session.clear()
    return redirect("/auth/login")


@app.get("/admin", response_class=HTMLResponse)
def admin_page(
    request: Request,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str | None = None,
):
    require_grandmaster(user)
    period = normalize_period(period or current_period())
    distributors = distributors_for(db, user)
    rows = []
    for distributor in distributors:
        people = visible_people_as_of(db, user, period, distributor.id)
        confirmed_ids = confirmation_ids_for(db, people, period)
        summary = dashboard_summary_from_people(people, confirmed_ids)
        rows.append({"distributor": distributor, "summary": summary})
    invitations = db.scalars(select(InvitationCode).order_by(InvitationCode.created_at.desc())).all()
    invitation_usage: dict[str, list[User]] = {}
    for distributor in distributors:
        if distributor.invitation_code:
            invitation_usage.setdefault(distributor.invitation_code, []).append(distributor)
    logs = db.scalars(select(AuditLog).order_by(AuditLog.created_at.desc()).limit(80)).all()
    return render(request, "admin.html", {
        "user": user,
        "period": period,
        "prev_period": period_add(period, -1),
        "next_period": period_add(period, 1),
        "quick_periods": year_periods(period),
        "rows": rows,
        "invitations": invitations,
        "invitation_usage": invitation_usage,
        "logs": logs,
    })


@app.post("/admin/invitations")
def create_invitation(
    request: Request,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    note: str | None = Form(None),
):
    require_grandmaster(user)
    code = secrets.token_urlsafe(8).replace("-", "").replace("_", "")[:10].upper()
    while db.scalar(select(InvitationCode).where(InvitationCode.code == code)):
        code = secrets.token_urlsafe(8).replace("-", "").replace("_", "")[:10].upper()
    invitation = InvitationCode(code=code, created_by_id=user.id, note=(note or "").strip() or None)
    db.add(invitation)
    write_audit(db, user, "invitation.create", "invitation", None, f"邀请码 {code}")
    db.commit()
    return redirect("/admin", message=f"邀请码已生成：{code}")


@app.post("/admin/invitations/{invite_id}/delete")
def delete_invitation(
    invite_id: int,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
):
    require_grandmaster(user)
    invitation = db.get(InvitationCode, invite_id)
    if not invitation:
        raise HTTPException(status_code=404, detail="邀请码不存在")
    invitation.is_active = False
    invitation.deleted_at = datetime.utcnow()
    write_audit(db, user, "invitation.delete", "invitation", invitation.id, f"邀请码 {invitation.code}")
    db.commit()
    return redirect("/admin", message="邀请码已作废")


@app.post("/admin/distributors/{distributor_id}")
def update_distributor(
    distributor_id: int,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    display_name: str = Form(...),
    default_service_fee: float = Form(800),
    is_active: str | None = Form(None),
    period: str = Form(...),
):
    require_grandmaster(user)
    distributor = db.get(User, distributor_id)
    if not distributor or distributor.role != Role.DISTRIBUTOR.value:
        raise HTTPException(status_code=404, detail="分销商不存在")
    distributor.display_name = display_name.strip()
    distributor.default_service_fee = default_service_fee
    distributor.is_active = is_active == "on"
    write_audit(db, user, "distributor.update", "user", distributor.id, f"编辑分销商 {distributor.username}")
    db.commit()
    return redirect("/admin", period=period, message="分销商信息已保存")


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(
    request: Request,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)] = None,
    period: str | None = None,
    owner_id: str | None = None,
):
    owner_id = parse_optional_int(owner_id)
    period = normalize_period(period or current_period())
    people = visible_people_as_of(db, user, period, owner_id)
    active_people = [person for person in people if normalize_status(person.employment_status) == ACTIVE_STATUS]
    left_people = [person for person in people if normalize_status(person.employment_status) == LEFT_STATUS]
    confirmed_ids = confirmation_ids_for(db, people, period)
    summary = dashboard_summary_from_people(people, confirmed_ids)
    owners = owner_options_for(db, user, owner_id)
    people_groups = grouped_people_for_owners(owners, people, confirmed_ids, actor=user)
    all_people_for_split = visible_people_as_of(db, user, period, None)
    all_confirmed_ids = confirmation_ids_for(db, all_people_for_split, period)
    own_people = [person for person in all_people_for_split if person.owner_id == user.id]
    distributor_people = [person for person in all_people_for_split if person.owner_id != user.id]
    own_summary = dashboard_summary_from_people(own_people, all_confirmed_ids)
    distributor_summary = dashboard_summary_from_people(distributor_people, all_confirmed_ids)
    confirmable_active_people = [person for person in active_people if user.role != Role.GRANDMASTER.value or person.owner_id == user.id]
    chart_data = annual_chart_data(db, user, period, owner_id)
    data_points = [item for item in chart_data if item["has_data"]]
    max_active_count = max([item["active_count"] for item in data_points] + [1])
    max_return_amount = max([item["return_amount"] for item in data_points] + [1])
    line_points = " ".join(
        f"{index * 21.8:.1f},{90 - (item['return_amount'] / max_return_amount * 72 if max_return_amount else 0):.1f}"
        for index, item in enumerate(chart_data)
        if item["has_data"]
    )
    return render(request, "dashboard.html", {
        "user": user,
        "period": period,
        "prev_period": period_add(period, -1),
        "next_period": period_add(period, 1),
        "quick_periods": year_periods(period),
        "people": people,
        "active_people": active_people,
        "confirmable_active_people": confirmable_active_people,
        "left_people": left_people,
        "people_groups": people_groups,
        "summary": summary,
        "own_summary": own_summary,
        "distributor_summary": distributor_summary,
        "distributor_count": len(distributors_for(db, user)) if user.role == Role.GRANDMASTER.value else 0,
        "confirmed_ids": confirmed_ids,
        "chart_data": chart_data,
        "line_points": line_points,
        "max_active_count": max_active_count,
        "max_return_amount": max_return_amount,
        "distributors": distributors_for(db, user),
        "owner_id": owner_id,
    })


@app.get("/me", response_class=HTMLResponse)
def me_page(
    request: Request,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)] = None,
    period: str | None = None,
    owner_id: str | None = None,
):
    owner_id = parse_optional_int(owner_id)
    period = normalize_period(period or current_period())
    people = db.scalars(_all_people_query(db, user, owner_id).options(selectinload(Person.owner))).all()
    summary_owner_id = owner_id if owner_id else (user.id if user.role == Role.GRANDMASTER.value else None)
    summary = profile_summary(db, user, period, summary_owner_id)
    people_groups = grouped_people_for_owners(owner_options_for(db, user, owner_id), people, actor=user)
    return render(request, "me.html", {
        "user": user,
        "period": period,
        "prev_period": period_add(period, -1),
        "next_period": period_add(period, 1),
        "quick_periods": year_periods(period),
        "summary": summary,
        "people": people,
        "people_groups": people_groups,
        "distributors": distributors_for(db, user),
        "owner_id": owner_id,
        "export_fields": [item for item in FIELD_DEFINITIONS if item[0] != "service_fee"],
        "default_export_fields": DEFAULT_EXPORT_FIELDS,
    })


@app.get("/export-center", response_class=HTMLResponse)
def export_center(
    request: Request,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)] = None,
    period: str | None = None,
    owner_id: str | None = None,
    download_id: int | None = None,
):
    owner_id = parse_optional_int(owner_id)
    period = normalize_period(period or current_period())
    people = visible_people_as_of(db, user, period, owner_id)
    download_file = db.get(StoredFile, download_id) if download_id else None
    if download_file and download_file.owner_id != user.id and user.role != Role.GRANDMASTER.value:
        download_file = None
    return render(request, "export_center.html", {
        "user": user,
        "period": period,
        "people": people,
        "distributors": distributors_for(db, user),
        "owner_id": owner_id,
        "export_fields": [item for item in FIELD_DEFINITIONS if item[0] != "service_fee"],
        "default_export_fields": DEFAULT_EXPORT_FIELDS,
        "download_file": download_file,
        "download_path": str(Path(download_file.path)) if download_file else None,
    })


@app.post("/settings/service-fee")
def update_service_fee(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    default_service_fee: float = Form(...),
    period: str = Form(...),
):
    if default_service_fee < 0:
        return redirect("/me", period=period, error="服务费不能小于 0")
    user.default_service_fee = default_service_fee
    write_audit(db, user, "settings.service_fee", "user", user.id, "修改默认服务费")
    db.commit()
    return redirect("/me", period=period, message="服务费设置已保存")


@app.post("/settings/password")
def update_password(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    current_password: str = Form(...),
    new_password: str = Form(...),
    confirm_password: str = Form(...),
    period: str = Form(...),
):
    if not verify_password(current_password, user.password_hash):
        return redirect("/me", period=period, error="当前密码不正确")
    if len(new_password) < 8:
        return redirect("/me", period=period, error="新密码至少 8 位")
    if new_password != confirm_password:
        return redirect("/me", period=period, error="两次输入的新密码不一致")
    if current_password == new_password:
        return redirect("/me", period=period, error="新密码不能和当前密码相同")
    user.password_hash = hash_password(new_password)
    write_audit(db, user, "settings.password", "user", user.id, "修改登录密码")
    db.commit()
    return redirect("/me", period=period, message="密码已修改，服务器已同步保存")


@app.get("/people/new", response_class=HTMLResponse)
def new_person_page(request: Request, db: Annotated[Session, Depends(get_db)], user: Annotated[User, Depends(get_current_user)] = None):
    return render(request, "person_form.html", {"user": user, "period": current_period(), "distributors": distributors_for(db, user), "person": None})


@app.get("/people/{person_id}/edit", response_class=HTMLResponse)
def edit_person_page(request: Request, person_id: int, db: Annotated[Session, Depends(get_db)], user: Annotated[User, Depends(get_current_user)] = None):
    person = db.get(Person, person_id)
    if not person or not can_access_person(user, person):
        raise HTTPException(status_code=404, detail="人员不存在")
    return render(request, "person_form.html", {"user": user, "period": person.settlement_period, "distributors": distributors_for(db, user), "person": person})


def person_payload(
    db: Session,
    user: User,
    serial_no: str | None,
    name: str,
    settlement_period: str,
    owner_id: int | None,
    sfid: str | None,
    sfid_expires_at: str | None,
    disability_cert_id: str | None,
    cert_issued_at: str | None,
    work_area: str | None,
    placement_period: str | None,
    salary_card: str | None,
    payroll_type: str | None,
    gross_pay: float,
    return_amount: float,
    employment_status: str,
    channel: str | None,
    household_address: str | None,
    household_type: str | None,
    contact_phone: str | None,
    emergency_contact: str | None,
    emergency_phone: str | None,
    emergency_relation: str | None,
    education: str | None,
    marital_status: str | None,
    bank_card_id: str | None,
    bank_name: str | None,
    disability_type1: str | None,
    disability_level1: str | None,
    disability_type2: str | None,
    disability_level2: str | None,
    entry_date: str | None,
    age: str | None,
    gender: str | None,
    disability_part: str | None,
    disability_reason: str | None,
    note: str | None,
    exclude_id: int | None = None,
) -> dict:
    if not valid_person_name(name):
        raise ValueError("姓名是必填项，不能填写合计行")
    target_owner_id = user.id
    if user.role == Role.GRANDMASTER.value and owner_id:
        target_owner_id = owner_id
    sfid = str(sfid or "").strip() or None
    disability_cert_id = str(disability_cert_id or "").strip() or None
    if duplicate_identifier(db, "sfid", sfid, exclude_id):
        raise ValueError("SFid 已存在于全局数据，请检查是否重复录入")
    if duplicate_identifier(db, "disability_cert_id", disability_cert_id, exclude_id):
        raise ValueError("残疾证ID 已存在于全局数据，请检查是否重复录入")
    owner = db.get(User, target_owner_id) or user
    return {
        "owner_id": target_owner_id,
        "serial_no": int(serial_no) if str(serial_no or "").strip().isdigit() else None,
        "name": name.strip(),
        "sfid": sfid,
        "sfid_expires_at": sfid_expires_at,
        "disability_cert_id": disability_cert_id,
        "cert_issued_at": cert_issued_at,
        "work_area": work_area,
        "placement_period": placement_period,
        "salary_card": salary_card,
        "service_fee": owner.default_service_fee or 0,
        "payroll_type": payroll_type,
        "gross_pay": gross_pay,
        "return_amount": return_amount,
        "settlement_period": normalize_period(settlement_period),
        "employment_status": normalize_status(employment_status),
        "channel": channel,
        "household_address": household_address,
        "household_type": household_type,
        "contact_phone": contact_phone,
        "emergency_contact": emergency_contact,
        "emergency_phone": emergency_phone,
        "emergency_relation": emergency_relation,
        "education": education,
        "marital_status": marital_status,
        "bank_card_id": bank_card_id,
        "bank_name": bank_name,
        "disability_type1": disability_type1,
        "disability_level1": disability_level1,
        "disability_type2": disability_type2,
        "disability_level2": disability_level2,
        "entry_date": entry_date,
        "age": age,
        "gender": gender,
        "disability_part": disability_part,
        "disability_reason": disability_reason,
        "note": note,
    }


@app.post("/people")
def create_person(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    serial_no: str | None = Form(None),
    name: str = Form(...),
    settlement_period: str = Form(...),
    owner_id: str | None = Form(None),
    sfid: str | None = Form(None),
    sfid_expires_at: str | None = Form(None),
    disability_cert_id: str | None = Form(None),
    cert_issued_at: str | None = Form(None),
    work_area: str | None = Form(None),
    placement_period: str | None = Form(None),
    salary_card: str | None = Form(None),
    payroll_type: str | None = Form(None),
    gross_pay: float = Form(0),
    return_amount: float = Form(0),
    employment_status: str = Form("在职"),
    channel: str | None = Form(None),
    household_address: str | None = Form(None),
    household_type: str | None = Form(None),
    contact_phone: str | None = Form(None),
    emergency_contact: str | None = Form(None),
    emergency_phone: str | None = Form(None),
    emergency_relation: str | None = Form(None),
    education: str | None = Form(None),
    marital_status: str | None = Form(None),
    bank_card_id: str | None = Form(None),
    bank_name: str | None = Form(None),
    disability_type1: str | None = Form(None),
    disability_level1: str | None = Form(None),
    disability_type2: str | None = Form(None),
    disability_level2: str | None = Form(None),
    entry_date: str | None = Form(None),
    age: str | None = Form(None),
    gender: str | None = Form(None),
    disability_part: str | None = Form(None),
    disability_reason: str | None = Form(None),
    note: str | None = Form(None),
):
    owner_id = parse_optional_int(owner_id)
    try:
        payload = person_payload(db, user, serial_no, name, settlement_period, owner_id, sfid, sfid_expires_at, disability_cert_id, cert_issued_at, work_area, placement_period, salary_card, payroll_type, gross_pay, return_amount, employment_status, channel, household_address, household_type, contact_phone, emergency_contact, emergency_phone, emergency_relation, education, marital_status, bank_card_id, bank_name, disability_type1, disability_level1, disability_type2, disability_level2, entry_date, age, gender, disability_part, disability_reason, note)
    except ValueError as exc:
        return redirect("/people/new", error=str(exc))
    person = Person(**payload)
    db.add(person)
    try:
        db.flush()
        write_audit(db, user, "person.create", "person", person.id, f"新增人员 {person.name}")
        db.commit()
    except IntegrityError:
        db.rollback()
        return redirect("/people/new", error="同一分销商、同一人员、同一结算月份已存在")
    except Exception:
        db.rollback()
        logger.exception("新增人员失败")
        return redirect("/people/new", error="新增失败，请查看日志")
    return redirect("/me", period=person.settlement_period, message="人员已添加")


@app.post("/people/{person_id}")
def update_person(
    person_id: int,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    serial_no: str | None = Form(None),
    name: str = Form(...),
    settlement_period: str = Form(...),
    owner_id: str | None = Form(None),
    sfid: str | None = Form(None),
    sfid_expires_at: str | None = Form(None),
    disability_cert_id: str | None = Form(None),
    cert_issued_at: str | None = Form(None),
    work_area: str | None = Form(None),
    placement_period: str | None = Form(None),
    salary_card: str | None = Form(None),
    payroll_type: str | None = Form(None),
    gross_pay: float = Form(0),
    return_amount: float = Form(0),
    employment_status: str = Form("在职"),
    channel: str | None = Form(None),
    household_address: str | None = Form(None),
    household_type: str | None = Form(None),
    contact_phone: str | None = Form(None),
    emergency_contact: str | None = Form(None),
    emergency_phone: str | None = Form(None),
    emergency_relation: str | None = Form(None),
    education: str | None = Form(None),
    marital_status: str | None = Form(None),
    bank_card_id: str | None = Form(None),
    bank_name: str | None = Form(None),
    disability_type1: str | None = Form(None),
    disability_level1: str | None = Form(None),
    disability_type2: str | None = Form(None),
    disability_level2: str | None = Form(None),
    entry_date: str | None = Form(None),
    age: str | None = Form(None),
    gender: str | None = Form(None),
    disability_part: str | None = Form(None),
    disability_reason: str | None = Form(None),
    note: str | None = Form(None),
):
    person = db.get(Person, person_id)
    if not person or not can_access_person(user, person):
        raise HTTPException(status_code=404, detail="人员不存在")
    try:
        owner_id = parse_optional_int(owner_id)
        payload = person_payload(db, user, serial_no, name, settlement_period, owner_id, sfid, sfid_expires_at, disability_cert_id, cert_issued_at, work_area, placement_period, salary_card, payroll_type, gross_pay, return_amount, employment_status, channel, household_address, household_type, contact_phone, emergency_contact, emergency_phone, emergency_relation, education, marital_status, bank_card_id, bank_name, disability_type1, disability_level1, disability_type2, disability_level2, entry_date, age, gender, disability_part, disability_reason, note, exclude_id=person_id)
    except ValueError as exc:
        return redirect(f"/people/{person_id}/edit", error=str(exc))
    for key, value in payload.items():
        setattr(person, key, value)
    person.monthly_confirmed = False
    person.confirmed_at = None
    try:
        write_audit(db, user, "person.update", "person", person.id, f"编辑人员 {person.name}")
        db.commit()
    except IntegrityError:
        db.rollback()
        return redirect(f"/people/{person_id}/edit", error="同一分销商、同一人员、同一结算月份已存在")
    except Exception:
        db.rollback()
        logger.exception("编辑人员失败")
        return redirect(f"/people/{person_id}/edit", error="保存失败，请查看日志")
    return redirect("/me", period=person.settlement_period, message="人员已更新")


@app.post("/people/{person_id}/status")
def update_status(person_id: int, db: Annotated[Session, Depends(get_db)], user: Annotated[User, Depends(get_current_user)], employment_status: str = Form(...), period: str = Form(...)):
    person = db.get(Person, person_id)
    if not person or not can_access_person(user, person):
        raise HTTPException(status_code=404, detail="人员不存在")
    person.employment_status = normalize_status(employment_status)
    person.monthly_confirmed = False
    person.confirmed_at = None
    write_audit(db, user, "person.status", "person", person.id, f"{person.name} 状态改为 {person.employment_status}")
    db.commit()
    return redirect("/dashboard", period=normalize_period(period), message="状态已更新")


@app.post("/people/{person_id}/delete")
def delete_person(
    person_id: int,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str = Form(...),
):
    person = db.get(Person, person_id)
    if not person or not can_access_person(user, person):
        raise HTTPException(status_code=404, detail="人员不存在")
    try:
        write_audit(db, user, "person.delete", "person", person.id, f"删除人员 {person.name}")
        db.delete(person)
        db.commit()
        logger.info("用户 %s 删除人员 id=%s name=%s", user.username, person_id, person.name)
        return redirect("/me", period=normalize_period(period), message="人员已删除")
    except Exception:
        db.rollback()
        logger.exception("删除人员失败")
        return redirect("/me", period=normalize_period(period), error="删除失败，请查看日志")


@app.post("/confirmations")
def batch_confirm(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str = Form(...),
    owner_id: str | None = Form(None),
    person_ids: list[int] = Form(default=[]),
):
    owner_id = parse_optional_int(owner_id)
    period = normalize_period(period)
    if not person_ids:
        return redirect("/dashboard", period=period, owner_id=owner_id, error="请先勾选要核准的在职人员")
    people = db.scalars(select(Person).where(Person.id.in_(person_ids))).all()
    visible_ids = {person.id for person in visible_people_as_of(db, user, period, owner_id)}
    changed = 0
    existing_ids = {
        item for item in db.scalars(
            select(MonthlyConfirmation.person_id).where(
                MonthlyConfirmation.period == period,
                MonthlyConfirmation.person_id.in_(person_ids),
            )
        ).all()
    }
    for person in people:
        if not can_access_person(user, person):
            continue
        if user.role == Role.GRANDMASTER.value and person.owner_id != user.id:
            continue
        if person.id not in visible_ids or normalize_status(person.employment_status) != ACTIVE_STATUS or not valid_person_name(person.name):
            continue
        if person.id in existing_ids:
            continue
        db.add(MonthlyConfirmation(person_id=person.id, period=period, confirmed_by_id=user.id, confirmed_at=datetime.utcnow()))
        changed += 1
    write_audit(db, user, "confirmation.batch", "monthly_confirmation", None, f"{period} 批量核准 {changed} 人")
    db.commit()
    return redirect("/dashboard", period=period, owner_id=owner_id, message=f"已批量核准 {changed} 人")


@app.get("/imports", response_class=HTMLResponse)
def import_page(request: Request, db: Annotated[Session, Depends(get_db)], user: Annotated[User, Depends(get_current_user)] = None):
    return render(request, "import.html", {"user": user, "distributors": distributors_for(db, user)})


@app.post("/imports")
async def import_excel(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    file: UploadFile = File(...),
    owner_id: str | None = Form(None),
):
    try:
        owner_id = parse_optional_int(owner_id)
        target_owner = user
        if user.role == Role.GRANDMASTER.value and owner_id:
            target_owner = db.get(User, owner_id)
            if not target_owner or target_owner.role != Role.DISTRIBUTOR.value:
                return redirect("/imports", error="请选择有效分销商")
        stored_name, path = await save_upload(file)
        db.add(StoredFile(owner_id=user.id, file_type="upload", original_name=file.filename, stored_name=stored_name, path=str(path)))
        db.commit()
        created = import_people_from_xlsx(db, path, target_owner, user)
        write_audit(db, user, "people.import", "user", target_owner.id, f"导入 {created} 条，文件 {file.filename}")
        db.commit()
        return redirect(f"/me?period={current_period()}&message=导入完成：{created} 条")
    except Exception as exc:
        logger.exception("上传导入失败")
        return redirect("/imports", error=str(exc))


@app.get("/exports")
def export_excel(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str | None = None,
    owner_id: str | None = None,
    fields: list[str] | None = Query(default=None),
    person_ids: list[int] | None = Query(default=None),
):
    try:
        owner_id = parse_optional_int(owner_id)
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/export-center", period=period, error="请先勾选需要导出的人员")
        visible_ids = {person.id for person in visible_people_as_of(db, user, period, owner_id)}
        people = [
            person for person in db.scalars(select(Person).where(Person.id.in_(person_ids)).options(selectinload(Person.owner))).all()
            if person.id in visible_ids and can_access_person(user, person)
        ]
        record = export_people_to_xlsx(db, people, user, period, fields)
        return FileResponse(record.path, filename=Path(record.path).name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        logger.exception("导出失败")
        raise HTTPException(status_code=500, detail="导出失败，请查看日志")


@app.get("/exports/create")
def create_excel_export(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str | None = None,
    owner_id: str | None = None,
    fields: list[str] | None = Query(default=None),
    person_ids: list[int] | None = Query(default=None),
):
    try:
        owner_id = parse_optional_int(owner_id)
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/me", period=period, error="请先勾选需要导出的人员")
        visible_ids = {person.id for person in visible_people_as_of(db, user, period, owner_id)}
        people = [
            person for person in db.scalars(select(Person).where(Person.id.in_(person_ids)).options(selectinload(Person.owner))).all()
            if person.id in visible_ids and can_access_person(user, person)
        ]
        if not people:
            return redirect("/export-center", period=period, error="没有可导出的人员")
        record = export_people_to_xlsx(db, people, user, period, fields)
        logger.info("用户 %s 生成 Excel 导出：file_id=%s people=%s", user.username, record.id, len(people))
        return redirect("/export-center", period=period, download_id=record.id, message=f"表格已生成：{len(people)} 人")
    except Exception:
        logger.exception("生成 Excel 导出失败")
        return redirect("/export-center", period=normalize_period(period or current_period()), error="生成表格失败，请查看日志")


@app.get("/exports/word")
def export_word(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str | None = None,
    owner_id: str | None = None,
    person_ids: list[int] | None = Query(default=None),
):
    try:
        owner_id = parse_optional_int(owner_id)
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/export-center", period=period, error="请先勾选需要导出的人员")
        visible_ids = {person.id for person in visible_people_as_of(db, user, period, owner_id)}
        people = [
            person for person in db.scalars(select(Person).where(Person.id.in_(person_ids)).options(selectinload(Person.owner))).all()
            if person.id in visible_ids and can_access_person(user, person)
        ]
        record = export_people_to_docx(db, people, user, period, None)
        return FileResponse(record.path, filename="个人信息.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception:
        logger.exception("导出 Word 失败")
        raise HTTPException(status_code=500, detail="导出 Word 失败，请查看日志")


@app.get("/exports/word/create")
def create_word_export(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str | None = None,
    owner_id: str | None = None,
    person_ids: list[int] | None = Query(default=None),
    fields: list[str] | None = Query(default=None),
):
    try:
        owner_id = parse_optional_int(owner_id)
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/me", period=period, error="请先勾选需要导出的人员")
        visible_ids = {person.id for person in visible_people_as_of(db, user, period, owner_id)}
        people = [
            person for person in db.scalars(select(Person).where(Person.id.in_(person_ids)).options(selectinload(Person.owner))).all()
            if person.id in visible_ids and can_access_person(user, person)
        ]
        if not people:
            return redirect("/export-center", period=period, error="没有可导出的人员")
        record = export_people_to_docx(db, people, user, period, fields)
        logger.info("用户 %s 生成 Word 导出：file_id=%s people=%s", user.username, record.id, len(people))
        return redirect("/export-center", period=period, download_id=record.id, message=f"Word 已生成：{len(people)} 人")
    except Exception:
        logger.exception("生成 Word 导出失败")
        return redirect("/export-center", period=normalize_period(period or current_period()), error="生成 Word 失败，请查看日志")


@app.get("/files/{file_id}")
def download_file(file_id: int, db: Annotated[Session, Depends(get_db)], user: Annotated[User, Depends(get_current_user)]):
    record = db.get(StoredFile, file_id)
    if not record or (record.owner_id != user.id and user.role != Role.GRANDMASTER.value):
        raise HTTPException(status_code=404, detail="文件不存在")
    path = Path(record.path)
    if not path.exists():
        raise HTTPException(status_code=404, detail="文件已不存在")
    if path.suffix.lower() == ".docx":
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        filename = record.stored_name or "个人信息.docx"
    else:
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = record.stored_name or path.name
    return FileResponse(
        path,
        filename=filename,
        media_type=media_type,
        headers={
            "X-Content-Type-Options": "nosniff",
            "Cache-Control": "private, no-store",
        },
    )


@app.get("/template.xlsx")
def download_template(user: Annotated[User, Depends(get_current_user)]):
    try:
        _, path = create_template_xlsx()
        return FileResponse(path, filename="人员导入模板.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        logger.exception("模板下载失败")
        raise HTTPException(status_code=500, detail="模板下载失败，请查看日志")


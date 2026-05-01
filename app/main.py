import logging
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
from app.models import Person, Role, StoredFile, User
from app.security import get_current_user, hash_password, verify_password
from app.services.excel import DEFAULT_EXPORT_FIELDS, FIELD_DEFINITIONS, create_template_xlsx, export_people_to_xlsx, import_people_from_xlsx
from app.services.ledger import current_period, dashboard_summary, normalize_period, normalize_status, profile_summary, valid_person_name, visible_people_query
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
    if params:
        from urllib.parse import urlencode
        url = f"{url}?{urlencode(params, doseq=True)}"
    return RedirectResponse(url, status_code=303)


def can_access_person(user: User, person: Person) -> bool:
    return user.role == Role.GRANDMASTER.value or person.owner_id == user.id


def distributors_for(db: Session, user: User):
    if user.role != Role.GRANDMASTER.value:
        return []
    return db.scalars(select(User).where(User.role == Role.DISTRIBUTOR.value).order_by(User.display_name)).all()


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
def register_page(request: Request):
    return render(request, "register.html")


@app.post("/auth/register")
def register(db: Annotated[Session, Depends(get_db)], username: str = Form(...), display_name: str = Form(...), password: str = Form(...)):
    if len(password) < 8:
        return redirect("/auth/register", error="密码至少 8 位")
    user = User(username=username.strip(), display_name=display_name.strip(), password_hash=hash_password(password), role=Role.DISTRIBUTOR.value, default_service_fee=800)
    db.add(user)
    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        return redirect("/auth/register", error="用户名已存在")
    return redirect("/auth/login", message="注册成功，请登录")


@app.post("/auth/logout")
def logout(request: Request):
    request.session.clear()
    return redirect("/auth/login")


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(
    request: Request,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)] = None,
    period: str | None = None,
    owner_id: int | None = None,
):
    period = normalize_period(period or current_period())
    people = db.scalars(visible_people_query(db, user, period, owner_id).options(selectinload(Person.owner))).all()
    active_people = [person for person in people if person.employment_status == "在职"]
    left_people = [person for person in people if person.employment_status == "离职"]
    summary = dashboard_summary(db, user, period, owner_id)
    return render(request, "dashboard.html", {
        "user": user,
        "period": period,
        "people": people,
        "active_people": active_people,
        "left_people": left_people,
        "summary": summary,
        "distributors": distributors_for(db, user),
        "owner_id": owner_id,
    })


@app.get("/me", response_class=HTMLResponse)
def me_page(
    request: Request,
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)] = None,
    period: str | None = None,
    owner_id: int | None = None,
):
    period = normalize_period(period or current_period())
    people = db.scalars(visible_people_query(db, user, period, owner_id).options(selectinload(Person.owner))).all()
    summary = profile_summary(db, user, period, owner_id)
    return render(request, "me.html", {
        "user": user,
        "period": period,
        "summary": summary,
        "people": people,
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
    owner_id: int | None = None,
    download_id: int | None = None,
):
    period = normalize_period(period or current_period())
    people = db.scalars(visible_people_query(db, user, period, owner_id).options(selectinload(Person.owner))).all()
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
    db.commit()
    return redirect("/me", period=period, message="服务费设置已保存")


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
    owner_id: int | None = Form(None),
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
    try:
        payload = person_payload(db, user, serial_no, name, settlement_period, owner_id, sfid, sfid_expires_at, disability_cert_id, cert_issued_at, work_area, placement_period, salary_card, payroll_type, gross_pay, return_amount, employment_status, channel, household_address, household_type, contact_phone, emergency_contact, emergency_phone, emergency_relation, education, marital_status, bank_card_id, bank_name, disability_type1, disability_level1, disability_type2, disability_level2, entry_date, age, gender, disability_part, disability_reason, note)
    except ValueError as exc:
        return redirect("/people/new", error=str(exc))
    person = Person(**payload)
    db.add(person)
    try:
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
    owner_id: int | None = Form(None),
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
        payload = person_payload(db, user, serial_no, name, settlement_period, owner_id, sfid, sfid_expires_at, disability_cert_id, cert_issued_at, work_area, placement_period, salary_card, payroll_type, gross_pay, return_amount, employment_status, channel, household_address, household_type, contact_phone, emergency_contact, emergency_phone, emergency_relation, education, marital_status, bank_card_id, bank_name, disability_type1, disability_level1, disability_type2, disability_level2, entry_date, age, gender, disability_part, disability_reason, note, exclude_id=person_id)
    except ValueError as exc:
        return redirect(f"/people/{person_id}/edit", error=str(exc))
    for key, value in payload.items():
        setattr(person, key, value)
    person.monthly_confirmed = False
    person.confirmed_at = None
    try:
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
    person_ids: list[int] = Form(default=[]),
):
    period = normalize_period(period)
    if not person_ids:
        return redirect("/dashboard", period=period, error="请先勾选要核准的在职人员")
    people = db.scalars(select(Person).where(Person.id.in_(person_ids))).all()
    changed = 0
    for person in people:
        if not can_access_person(user, person):
            continue
        if person.settlement_period != period or person.employment_status != "在职" or not valid_person_name(person.name):
            continue
        if not person.monthly_confirmed:
            person.monthly_confirmed = True
            person.confirmed_at = datetime.utcnow()
            changed += 1
    db.commit()
    return redirect("/dashboard", period=period, message=f"已批量核准 {changed} 人")


@app.get("/imports", response_class=HTMLResponse)
def import_page(request: Request, db: Annotated[Session, Depends(get_db)], user: Annotated[User, Depends(get_current_user)] = None):
    return render(request, "import.html", {"user": user, "distributors": distributors_for(db, user)})


@app.post("/imports")
async def import_excel(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    file: UploadFile = File(...),
    owner_id: int | None = Form(None),
):
    try:
        target_owner = user
        if user.role == Role.GRANDMASTER.value and owner_id:
            target_owner = db.get(User, owner_id)
            if not target_owner or target_owner.role != Role.DISTRIBUTOR.value:
                return redirect("/imports", error="请选择有效分销商")
        stored_name, path = await save_upload(file)
        db.add(StoredFile(owner_id=user.id, file_type="upload", original_name=file.filename, stored_name=stored_name, path=str(path)))
        db.commit()
        created = import_people_from_xlsx(db, path, target_owner, user)
        return redirect("/dashboard", message=f"导入完成：{created} 条")
    except Exception as exc:
        logger.exception("上传导入失败")
        return redirect("/imports", error=str(exc))


@app.get("/exports")
def export_excel(
    db: Annotated[Session, Depends(get_db)],
    user: Annotated[User, Depends(get_current_user)],
    period: str | None = None,
    owner_id: int | None = None,
    fields: list[str] | None = Query(default=None),
    person_ids: list[int] | None = Query(default=None),
):
    try:
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/export-center", period=period, error="请先勾选需要导出的人员")
        people = db.scalars(
            visible_people_query(db, user, period, owner_id)
            .where(Person.id.in_(person_ids))
            .options(selectinload(Person.owner))
        ).all()
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
    owner_id: int | None = None,
    fields: list[str] | None = Query(default=None),
    person_ids: list[int] | None = Query(default=None),
):
    try:
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/me", period=period, error="请先勾选需要导出的人员")
        people = db.scalars(
            visible_people_query(db, user, period, owner_id)
            .where(Person.id.in_(person_ids))
            .options(selectinload(Person.owner))
        ).all()
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
    owner_id: int | None = None,
    person_ids: list[int] | None = Query(default=None),
):
    try:
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/export-center", period=period, error="请先勾选需要导出的人员")
        people = db.scalars(
            visible_people_query(db, user, period, owner_id)
            .where(Person.id.in_(person_ids))
            .options(selectinload(Person.owner))
        ).all()
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
    owner_id: int | None = None,
    person_ids: list[int] | None = Query(default=None),
    fields: list[str] | None = Query(default=None),
):
    try:
        period = normalize_period(period or current_period())
        if not person_ids:
            return redirect("/me", period=period, error="请先勾选需要导出的人员")
        people = db.scalars(
            visible_people_query(db, user, period, owner_id)
            .where(Person.id.in_(person_ids))
            .options(selectinload(Person.owner))
        ).all()
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
        filename = "个人信息.docx"
    else:
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = path.name
    return FileResponse(path, filename=filename, media_type=media_type)


@app.get("/template.xlsx")
def download_template(user: Annotated[User, Depends(get_current_user)]):
    try:
        _, path = create_template_xlsx()
        return FileResponse(path, filename="人员导入模板.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        logger.exception("模板下载失败")
        raise HTTPException(status_code=500, detail="模板下载失败，请查看日志")

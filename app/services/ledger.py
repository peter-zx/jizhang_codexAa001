import logging
import re
from datetime import datetime

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app.models import Person, Role, User

logger = logging.getLogger(__name__)

_PERIOD_PATTERNS = [
    re.compile(r"(?P<year>20\d{2})\s*年\s*(?P<month>\d{1,2})\s*月"),
    re.compile(r"(?P<year>20\d{2})[-/.](?P<month>\d{1,2})"),
]
_INVALID_NAMES = ["-", "合计", "总计"]


def current_period() -> str:
    return datetime.now().strftime("%Y-%m")


def normalize_period(value: object | None) -> str:
    if value is None or value == "":
        return current_period()
    if isinstance(value, datetime):
        return value.strftime("%Y-%m")
    text = str(value).strip()
    for pattern in _PERIOD_PATTERNS:
        match = pattern.search(text)
        if match:
            return f"{int(match.group('year')):04d}-{int(match.group('month')):02d}"
    if re.fullmatch(r"\d{6}", text):
        return f"{text[:4]}-{text[4:]}"
    return text


def normalize_status(value: object | None) -> str:
    text = str(value or "").strip()
    if "离职" in text:
        return "离职"
    return "在职"


def valid_person_name(value: object | None) -> bool:
    text = str(value or "").strip()
    return bool(text) and text not in set(_INVALID_NAMES)


def _valid_people_filter(stmt):
    return stmt.where(Person.name.not_in(_INVALID_NAMES))


def _apply_scope(stmt, user: User, owner_id: int | None):
    if user.role == Role.DISTRIBUTOR.value:
        return stmt.where(Person.owner_id == user.id)
    if owner_id:
        return stmt.where(Person.owner_id == owner_id)
    return stmt


def visible_people_query(db: Session, user: User, period: str | None = None, owner_id: int | None = None):
    stmt = _valid_people_filter(select(Person))
    _apply_scope(stmt, user, owner_id)
    if period:
        stmt = stmt.where(Person.settlement_period == normalize_period(period))
    return stmt.order_by(Person.employment_status.desc(), Person.monthly_confirmed.asc(), Person.id.asc())


def _all_people_query(db: Session, user: User, owner_id: int | None = None):
    """All people regardless of period — for /me page listing."""
    stmt = _valid_people_filter(select(Person))
    _apply_scope(stmt, user, owner_id)
    return stmt.order_by(Person.employment_status.desc(), Person.monthly_confirmed.asc(), Person.id.asc())


def dashboard_summary(db: Session, user: User, period: str, owner_id: int | None = None) -> dict:
    period = normalize_period(period)
    active_base = _valid_people_filter(select(
        func.count(Person.id),
        func.coalesce(func.sum(Person.gross_pay), 0),
        func.coalesce(func.sum(Person.service_fee), 0),
        func.coalesce(func.sum(Person.return_amount), 0),
    )).where(Person.settlement_period == period, Person.employment_status == "在职")
    active_count, gross, service, returns = db.execute(_apply_scope(active_base, user, owner_id)).one()

    total = _valid_people_filter(select(func.count(Person.id))).where(Person.settlement_period == period)
    confirmed = _valid_people_filter(select(func.count(Person.id))).where(
        Person.settlement_period == period,
        Person.employment_status == "在职",
        Person.monthly_confirmed.is_(True),
    )
    total_count = db.scalar(_apply_scope(total, user, owner_id)) or 0
    confirmed_count = db.scalar(_apply_scope(confirmed, user, owner_id)) or 0
    pending_count = max((active_count or 0) - confirmed_count, 0)
    return {
        "count": active_count or 0,
        "total_count": total_count,
        "active_count": active_count or 0,
        "gross_pay": gross or 0,
        "service_fee": service or 0,
        "return_amount": returns or 0,
        "confirmed_count": confirmed_count,
        "pending_count": pending_count,
    }


def profile_summary(db: Session, user: User, period: str, owner_id: int | None = None) -> dict:
    period = normalize_period(period)
    active = _valid_people_filter(select(
        func.count(Person.id),
        func.coalesce(func.sum(Person.service_fee - 500), 0),
    )).where(Person.settlement_period == period, Person.employment_status == "在职")
    active_count, commission = db.execute(_apply_scope(active, user, owner_id)).one()

    total = _valid_people_filter(select(func.count(Person.id))).where(Person.settlement_period == period)
    left = _valid_people_filter(select(func.count(Person.id))).where(Person.settlement_period == period, Person.employment_status == "离职")
    return {
        "total_count": db.scalar(_apply_scope(total, user, owner_id)) or 0,
        "active_count": active_count or 0,
        "left_count": db.scalar(_apply_scope(left, user, owner_id)) or 0,
        "commission": commission or 0,
    }


def monthly_summary(db: Session, user: User, period: str, owner_id: int | None = None) -> dict:
    return dashboard_summary(db, user, period, owner_id)

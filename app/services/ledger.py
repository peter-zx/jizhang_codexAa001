import logging
import re
from datetime import datetime

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app.models import MonthlyConfirmation, Person, Role, User

logger = logging.getLogger(__name__)

_PERIOD_PATTERNS = [
    re.compile(r"(?P<year>20\d{2})\s*年\s*(?P<month>\d{1,2})\s*月?"),
    re.compile(r"(?P<year>20\d{2})[-/.](?P<month>\d{1,2})"),
]
_INVALID_NAMES = ["-", "合计", "总计"]
ACTIVE_STATUS = "在职"
LEFT_STATUS = "离职"


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
    if re.fullmatch(r"20\d{2}-\d{1,2}", text):
        year, month = text.split("-", 1)
        return f"{int(year):04d}-{int(month):02d}"
    return text


def period_add(period: str, months: int) -> str:
    year, month = [int(part) for part in normalize_period(period).split("-", 1)]
    month_index = year * 12 + month - 1 + months
    return f"{month_index // 12:04d}-{month_index % 12 + 1:02d}"


def recent_periods(period: str, count: int = 6) -> list[str]:
    period = normalize_period(period)
    return [period_add(period, offset) for offset in range(-(count - 1), 1)]


def year_periods(period: str) -> list[str]:
    year = int(normalize_period(period)[:4])
    return [f"{year:04d}-{month:02d}" for month in range(1, 13)]


def normalize_status(value: object | None) -> str:
    text = str(value or "").strip()
    if "离职" in text or "绂昏亴" in text:
        return LEFT_STATUS
    return ACTIVE_STATUS


def valid_person_name(value: object | None) -> bool:
    text = str(value or "").strip()
    return bool(text) and text not in set(_INVALID_NAMES)


def _valid_people_filter(stmt):
    return stmt.where(Person.name.not_in(_INVALID_NAMES))


def _apply_scope(stmt, user: User, owner_id: int | None):
    if user.role == Role.DISTRIBUTOR.value:
        return stmt.where(Person.owner_id == user.id)
    if user.role == Role.ASSISTANT.value:
        allowed_ids = []
        for raw_id in str(user.allowed_owner_ids or "").split(","):
            raw_id = raw_id.strip()
            if raw_id.isdigit():
                allowed_ids.append(int(raw_id))
        if owner_id and owner_id in allowed_ids:
            return stmt.where(Person.owner_id == owner_id)
        if allowed_ids:
            return stmt.where(Person.owner_id.in_(allowed_ids))
        return stmt.where(Person.owner_id == -1)
    if owner_id:
        return stmt.where(Person.owner_id == owner_id)
    return stmt


def person_start_period(person: Person) -> str:
    """The month when a person starts participating in monthly snapshots."""
    if person.entry_date:
        parsed = normalize_period(person.entry_date)
        if re.fullmatch(r"20\d{2}-\d{2}", parsed):
            return parsed
    if person.created_at:
        return person.created_at.strftime("%Y-%m")
    if person.settlement_period:
        return normalize_period(person.settlement_period)
    return current_period()


def _person_identity(person: Person) -> tuple[int, str, str]:
    if person.sfid and str(person.sfid).strip():
        return (person.owner_id, "sfid", str(person.sfid).strip())
    if person.disability_cert_id and str(person.disability_cert_id).strip():
        return (person.owner_id, "cert", str(person.disability_cert_id).strip())
    return (person.owner_id, "name", str(person.name).strip())


def visible_people_query(db: Session, user: User, period: str | None = None, owner_id: int | None = None):
    """Legacy row query, kept for forms/exports that still operate on stored rows."""
    stmt = _valid_people_filter(select(Person))
    stmt = _apply_scope(stmt, user, owner_id)
    if period:
        stmt = stmt.where(Person.settlement_period == normalize_period(period))
    return stmt.order_by(Person.employment_status.desc(), Person.monthly_confirmed.asc(), Person.id.asc())


def _all_people_query(db: Session, user: User, owner_id: int | None = None):
    """All people regardless of period for personnel management."""
    stmt = _valid_people_filter(select(Person))
    stmt = _apply_scope(stmt, user, owner_id)
    return stmt.order_by(Person.employment_status.desc(), Person.id.asc())


def visible_people_as_of(db: Session, user: User, period: str, owner_id: int | None = None) -> list[Person]:
    """Return one latest valid record per person as of the selected month.

    This treats settlement/entry month as the start month. Later monthly pages reuse
    the same person unless a newer imported/edited row exists for that identity.
    """
    target = normalize_period(period)
    stmt = _all_people_query(db, user, owner_id)
    rows = db.scalars(stmt).all()
    latest: dict[tuple[int, str, str], Person] = {}
    for person in rows:
        start = person_start_period(person)
        if start > target:
            continue
        key = _person_identity(person)
        current = latest.get(key)
        if current is None:
            latest[key] = person
            continue
        current_start = person_start_period(current)
        if (start, person.id) >= (current_start, current.id):
            latest[key] = person
    return sorted(
        latest.values(),
        key=lambda item: (item.employment_status != ACTIVE_STATUS, item.monthly_confirmed, item.id),
    )


def confirmation_ids_for(db: Session, people: list[Person], period: str) -> set[int]:
    period = normalize_period(period)
    person_ids = [person.id for person in people]
    if not person_ids:
        return set()
    confirmed_ids = set(db.scalars(
        select(MonthlyConfirmation.person_id).where(
            MonthlyConfirmation.period == period,
            MonthlyConfirmation.person_id.in_(person_ids),
        )
    ).all())
    for person in people:
        if person.monthly_confirmed and person_start_period(person) == period:
            confirmed_ids.add(person.id)
    return confirmed_ids


def dashboard_summary_from_people(people: list[Person], confirmed_ids: set[int] | None = None) -> dict:
    confirmed_ids = confirmed_ids or set()
    active_people = [person for person in people if normalize_status(person.employment_status) == ACTIVE_STATUS]
    left_people = [person for person in people if normalize_status(person.employment_status) == LEFT_STATUS]
    confirmed_count = sum(1 for person in active_people if person.id in confirmed_ids)
    active_count = len(active_people)
    return {
        "count": active_count,
        "total_count": len(people),
        "active_count": active_count,
        "left_count": len(left_people),
        "gross_pay": sum(person.gross_pay or 0 for person in active_people),
        "service_fee": sum(person.service_fee or 0 for person in active_people),
        "return_amount": sum(person.return_amount or 0 for person in active_people),
        "confirmed_count": confirmed_count,
        "pending_count": max(active_count - confirmed_count, 0),
    }


def dashboard_summary(db: Session, user: User, period: str, owner_id: int | None = None) -> dict:
    people = visible_people_as_of(db, user, period, owner_id)
    return dashboard_summary_from_people(people, confirmation_ids_for(db, people, period))


def profile_summary(db: Session, user: User, period: str, owner_id: int | None = None) -> dict:
    people = visible_people_as_of(db, user, period, owner_id)
    summary = dashboard_summary_from_people(people, confirmation_ids_for(db, people, period))
    active_people = [person for person in people if normalize_status(person.employment_status) == ACTIVE_STATUS]
    summary["commission"] = sum((person.service_fee or 0) - 500 for person in active_people)
    return summary


def monthly_chart_data(db: Session, user: User, period: str, owner_id: int | None = None, months: int = 6) -> list[dict]:
    data = []
    for item_period in recent_periods(period, months):
        summary = dashboard_summary(db, user, item_period, owner_id)
        data.append({
            "period": item_period,
            "label": item_period[5:],
            "total_count": summary["total_count"],
            "active_count": summary["active_count"],
            "left_count": summary["left_count"],
            "return_amount": float(summary["return_amount"] or 0),
            "confirmed_count": summary["confirmed_count"],
        })
    return data


def annual_chart_data(db: Session, user: User, period: str, owner_id: int | None = None) -> list[dict]:
    selected = normalize_period(period)
    selected_year = int(selected[:4])
    selected_month = int(selected[5:7])
    current = current_period()
    current_year = int(current[:4])
    current_month = int(current[5:7])
    if selected_year < current_year:
        cutoff_month = 12
    elif selected_year == current_year:
        cutoff_month = min(selected_month, current_month)
    else:
        cutoff_month = 0

    data = []
    for month in range(1, 13):
        item_period = f"{selected_year:04d}-{month:02d}"
        if month <= cutoff_month:
            summary = dashboard_summary(db, user, item_period, owner_id)
            has_data = summary["total_count"] > 0
            data.append({
                "period": item_period,
                "label": f"{month}月",
                "has_data": has_data,
                "total_count": summary["total_count"],
                "active_count": summary["active_count"],
                "left_count": summary["left_count"],
                "return_amount": float(summary["return_amount"] or 0),
                "confirmed_count": summary["confirmed_count"],
            })
        else:
            data.append({
                "period": item_period,
                "label": f"{month}月",
                "has_data": False,
                "total_count": 0,
                "active_count": 0,
                "left_count": 0,
                "return_amount": 0,
                "confirmed_count": 0,
            })
    return data


def monthly_summary(db: Session, user: User, period: str, owner_id: int | None = None) -> dict:
    return dashboard_summary(db, user, period, owner_id)

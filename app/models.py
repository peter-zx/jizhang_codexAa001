from datetime import datetime
from enum import Enum

from sqlalchemy import Boolean, DateTime, Float, ForeignKey, Integer, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.database import Base


class Role(str, Enum):
    GRANDMASTER = "grandmaster"
    DISTRIBUTOR = "distributor"


class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    username: Mapped[str] = mapped_column(String(64), unique=True, index=True, nullable=False)
    display_name: Mapped[str] = mapped_column(String(100), nullable=False)
    password_hash: Mapped[str] = mapped_column(String(255), nullable=False)
    role: Mapped[str] = mapped_column(String(32), default=Role.DISTRIBUTOR.value, index=True)
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)
    default_service_fee: Mapped[float] = mapped_column(Float, default=800)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    people: Mapped[list["Person"]] = relationship(back_populates="owner")


class Person(Base):
    __tablename__ = "people"
    __table_args__ = (UniqueConstraint("owner_id", "name", "settlement_period", name="uq_owner_name_period"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    owner_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True, nullable=False)
    serial_no: Mapped[int | None] = mapped_column(Integer)
    name: Mapped[str] = mapped_column(String(100), index=True, nullable=False)
    sfid: Mapped[str | None] = mapped_column(String(100), index=True)
    sfid_expires_at: Mapped[str | None] = mapped_column(String(50))
    disability_cert_id: Mapped[str | None] = mapped_column(String(100), index=True)
    cert_issued_at: Mapped[str | None] = mapped_column(String(50))
    work_area: Mapped[str | None] = mapped_column(String(100))
    placement_period: Mapped[str | None] = mapped_column(String(100))
    salary_card: Mapped[str | None] = mapped_column(String(100))
    service_fee: Mapped[float] = mapped_column(Float, default=0)
    payroll_type: Mapped[str | None] = mapped_column(String(50))
    gross_pay: Mapped[float] = mapped_column(Float, default=0)
    return_amount: Mapped[float] = mapped_column(Float, default=0)
    settlement_period: Mapped[str] = mapped_column(String(7), index=True, nullable=False)
    employment_status: Mapped[str] = mapped_column(String(20), default="在职")
    monthly_confirmed: Mapped[bool] = mapped_column(Boolean, default=False)
    confirmed_at: Mapped[datetime | None] = mapped_column(DateTime)
    note: Mapped[str | None] = mapped_column(Text)
    channel: Mapped[str | None] = mapped_column(String(100))
    household_address: Mapped[str | None] = mapped_column(String(255))
    household_type: Mapped[str | None] = mapped_column(String(100))
    contact_phone: Mapped[str | None] = mapped_column(String(100))
    emergency_contact: Mapped[str | None] = mapped_column(String(100))
    emergency_phone: Mapped[str | None] = mapped_column(String(100))
    emergency_relation: Mapped[str | None] = mapped_column(String(100))
    education: Mapped[str | None] = mapped_column(String(100))
    marital_status: Mapped[str | None] = mapped_column(String(100))
    bank_card_id: Mapped[str | None] = mapped_column(String(100))
    bank_name: Mapped[str | None] = mapped_column(String(100))
    disability_type1: Mapped[str | None] = mapped_column(String(100))
    disability_level1: Mapped[str | None] = mapped_column(String(100))
    disability_type2: Mapped[str | None] = mapped_column(String(100))
    disability_level2: Mapped[str | None] = mapped_column(String(100))
    entry_date: Mapped[str | None] = mapped_column(String(50))
    age: Mapped[str | None] = mapped_column(String(20))
    gender: Mapped[str | None] = mapped_column(String(20))
    disability_part: Mapped[str | None] = mapped_column(String(100))
    disability_reason: Mapped[str | None] = mapped_column(String(100))
    extra_fields: Mapped[str | None] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    owner: Mapped[User] = relationship(back_populates="people")


class StoredFile(Base):
    __tablename__ = "stored_files"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    owner_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True, nullable=False)
    file_type: Mapped[str] = mapped_column(String(20), index=True)
    original_name: Mapped[str | None] = mapped_column(String(255))
    stored_name: Mapped[str] = mapped_column(String(255), unique=True)
    path: Mapped[str] = mapped_column(String(500))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class MonthlyConfirmation(Base):
    __tablename__ = "monthly_confirmations"
    __table_args__ = (UniqueConstraint("person_id", "period", name="uq_person_period_confirmation"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    person_id: Mapped[int] = mapped_column(ForeignKey("people.id"), index=True, nullable=False)
    period: Mapped[str] = mapped_column(String(7), index=True, nullable=False)
    confirmed_by_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"), index=True)
    confirmed_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

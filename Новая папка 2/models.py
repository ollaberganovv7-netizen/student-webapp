from sqlalchemy import (
    Column, Integer, BigInteger, String, Boolean,
    DateTime, Text, JSON, ForeignKey
)
from sqlalchemy.orm import DeclarativeBase, relationship
from datetime import datetime


class Base(DeclarativeBase):
    pass


class User(Base):
    __tablename__ = "users"

    id = Column(BigInteger, primary_key=True)  # Telegram user ID
    username = Column(String(64), nullable=True)
    full_name = Column(String(128), nullable=True)
    language = Column(String(8), default="uz")
    free_used = Column(Boolean, default=False)
    is_premium = Column(Boolean, default=False) # Deprecated but kept for compatibility
    is_banned = Column(Boolean, default=False)
    premium_expires = Column(DateTime, nullable=True) # Deprecated
    balance = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.utcnow)
    last_active = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    payments = relationship("Payment", back_populates="user", lazy="select")
    requests = relationship("Request", back_populates="user", lazy="select")

    def __repr__(self):
        return f"<User id={self.id} username={self.username}>"


class Payment(Base):
    __tablename__ = "payments"

    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(BigInteger, ForeignKey("users.id"), nullable=False)
    amount = Column(Integer, nullable=False)
    package = Column(String(64), nullable=False)
    screenshot_file_id = Column(String(256), nullable=True)
    status = Column(String(16), default="pending")  # pending / approved / rejected
    admin_note = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    reviewed_at = Column(DateTime, nullable=True)

    user = relationship("User", back_populates="payments")

    def __repr__(self):
        return f"<Payment id={self.id} user={self.user_id} status={self.status}>"


class Request(Base):
    __tablename__ = "requests"

    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(BigInteger, ForeignKey("users.id"), nullable=False)
    service_type = Column(String(32), nullable=False)
    topic = Column(Text, nullable=False)
    options = Column(JSON, nullable=True)
    file_id = Column(String(256), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    user = relationship("User", back_populates="requests")

    def __repr__(self):
        return f"<Request id={self.id} type={self.service_type}>"

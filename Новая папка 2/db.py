from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession, async_sessionmaker
from sqlalchemy import select, update, func
from datetime import datetime, timedelta
from typing import Optional

from config import DATABASE_URL
from database.models import Base, User, Payment, Request

engine = create_async_engine(DATABASE_URL, echo=False)
async_session = async_sessionmaker(engine, expire_on_commit=False, class_=AsyncSession)


async def init_db():
    """Create all tables."""
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)


# ─── User helpers ───────────────────────────────────────────────────────────

async def get_or_create_user(
    user_id: int,
    username: Optional[str],
    full_name: Optional[str],
) -> User:
    async with async_session() as session:
        user = await session.get(User, user_id)
        if not user:
            user = User(
                id=user_id,
                username=username,
                full_name=full_name,
            )
            session.add(user)
            await session.commit()
            await session.refresh(user)
        else:
            # Update name/username if changed
            user.username = username
            user.full_name = full_name
            user.last_active = datetime.utcnow()
            await session.commit()
        return user


async def get_user(user_id: int) -> Optional[User]:
    async with async_session() as session:
        return await session.get(User, user_id)


async def mark_free_used(user_id: int):
    async with async_session() as session:
        await session.execute(
            update(User).where(User.id == user_id).values(free_used=True)
        )
        await session.commit()


async def activate_premium(user_id: int, days: int = 30):
    expires = datetime.utcnow() + timedelta(days=days)
    async with async_session() as session:
        await session.execute(
            update(User)
            .where(User.id == user_id)
            .values(is_premium=True, premium_expires=expires)
        )
        await session.commit()

async def add_balance(user_id: int, amount: int) -> Optional[User]:
    """Adds amount to user balance."""
    async with async_session() as session:
        user = await session.get(User, user_id)
        if user:
            user.balance = (user.balance or 0) + amount
            await session.commit()
            await session.refresh(user)
        return user

async def deduct_balance(user_id: int, amount: int) -> bool:
    """Deducts amount from user balance. Returns True if successful."""
    async with async_session() as session:
        user = await session.get(User, user_id)
        if user and (user.balance or 0) >= amount:
            user.balance -= amount
            await session.commit()
            return True
        return False


async def get_all_users() -> list[User]:
    async with async_session() as session:
        result = await session.execute(select(User))
        return result.scalars().all()


async def get_user_count() -> int:
    async with async_session() as session:
        result = await session.execute(select(func.count(User.id)))
        return result.scalar()


async def get_premium_count() -> int:
    async with async_session() as session:
        result = await session.execute(
            select(func.count(User.id)).where(User.is_premium == True)
        )
        return result.scalar()


# ─── Payment helpers ─────────────────────────────────────────────────────────

async def create_payment(
    user_id: int,
    amount: int,
    package: str,
    screenshot_file_id: str,
) -> Payment:
    async with async_session() as session:
        payment = Payment(
            user_id=user_id,
            amount=amount,
            package=package,
            screenshot_file_id=screenshot_file_id,
        )
        session.add(payment)
        await session.commit()
        await session.refresh(payment)
        return payment


async def get_pending_payments() -> list[Payment]:
    async with async_session() as session:
        result = await session.execute(
            select(Payment).where(Payment.status == "pending").order_by(Payment.created_at)
        )
        return result.scalars().all()


async def get_payment(payment_id: int) -> Optional[Payment]:
    async with async_session() as session:
        return await session.get(Payment, payment_id)


async def update_payment_status(
    payment_id: int,
    status: str,
    admin_note: Optional[str] = None,
) -> Optional[Payment]:
    async with async_session() as session:
        payment = await session.get(Payment, payment_id)
        if payment:
            payment.status = status
            payment.admin_note = admin_note
            payment.reviewed_at = datetime.utcnow()
            await session.commit()
            await session.refresh(payment)
        return payment


async def get_payment_count() -> int:
    async with async_session() as session:
        result = await session.execute(
            select(func.count(Payment.id)).where(Payment.status == "approved")
        )
        return result.scalar()


# ─── Request helpers ──────────────────────────────────────────────────────────

async def create_request(
    user_id: int,
    service_type: str,
    topic: str,
    options: Optional[dict] = None,
    file_id: Optional[str] = None,
) -> Request:
    async with async_session() as session:
        req = Request(
            user_id=user_id,
            service_type=service_type,
            topic=topic,
            options=options or {},
            file_id=file_id,
        )
        session.add(req)
        await session.commit()
        await session.refresh(req)
        return req


async def get_request_count() -> int:
    async with async_session() as session:
        result = await session.execute(select(func.count(Request.id)))
        return result.scalar()


async def get_recent_requests(limit: int = 10) -> list[Request]:
    async with async_session() as session:
        result = await session.execute(
            select(Request).order_by(Request.created_at.desc()).limit(limit)
        )
        return result.scalars().all()

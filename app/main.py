import asyncio
import io
import logging
import os
from contextlib import asynccontextmanager
from datetime import datetime, timezone

from fastapi import FastAPI
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    ChatMemberUpdated,
    BufferedInputFile,
)
from aiogram.enums import ChatMemberStatus
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

from sqlalchemy import (
    String,
    Integer,
    BigInteger,
    Boolean,
    DateTime,
    Text,
    select,
)
from sqlalchemy.ext.asyncio import (
    create_async_engine,
    async_sessionmaker,
)
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    mapped_column,
)

from openpyxl import Workbook


# ============================================================
# CONFIG
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

try:
    OWNER_ID = int(os.getenv("OWNER_ID", "0") or "0")
except ValueError:
    OWNER_ID = 0

CHANNEL_USERNAME = os.getenv(
    "CHANNEL_USERNAME",
    "@nemazing",
).strip()

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "sqlite+aiosqlite:///./data/nemazing.db",
).strip()

try:
    PORT = int(os.getenv("PORT", "10000") or "10000")
except ValueError:
    PORT = 10000


os.makedirs("data", exist_ok=True)
os.makedirs("exports", exist_ok=True)
os.makedirs("banners", exist_ok=True)


# PostgreSQL URL normalization
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace(
        "postgres://",
        "postgresql+asyncpg://",
        1,
    )

elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace(
        "postgresql://",
        "postgresql+asyncpg://",
        1,
    )


# ============================================================
# DATABASE
# ============================================================

engine = create_async_engine(
    DATABASE_URL,
    echo=False,
    pool_pre_ping=True,
)

SessionLocal = async_sessionmaker(
    engine,
    expire_on_commit=False,
)


class Base(DeclarativeBase):
    pass


# ============================================================
# MODELS
# ============================================================

class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
    )

    telegram_id: Mapped[int] = mapped_column(
        BigInteger,
        unique=True,
        index=True,
    )

    username: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    first_name: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    nickname: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    rank: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    position: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    department: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    organization: Mapped[str | None] = mapped_column(
        String(50),
        nullable=True,
    )

    role: Mapped[str] = mapped_column(
        String(30),
        default="PLAYER",
    )

    active: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
    )

    archived: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
    )

    joined_at: Mapped[datetime | None] = mapped_column(
        DateTime,
        nullable=True,
    )

    archived_at: Mapped[datetime | None] = mapped_column(
        DateTime,
        nullable=True,
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(timezone.utc),
    )


class Organization(Base):
    __tablename__ = "organizations"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
    )

    code: Mapped[str] = mapped_column(
        String(50),
        unique=True,
    )

    title: Mapped[str] = mapped_column(
        String(255),
    )

    chat_id: Mapped[int | None] = mapped_column(
        BigInteger,
        nullable=True,
    )

    invite_link: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    enabled: Mapped[bool] = mapped_column(
        Boolean,
        default=True,
    )

    auto_ban_on_leave: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
    )


class Application(Base):
    __tablename__ = "applications"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
    )

    telegram_id: Mapped[int] = mapped_column(
        BigInteger,
        index=True,
    )

    organization: Mapped[str] = mapped_column(
        String(50),
    )

    nickname: Mapped[str] = mapped_column(
        String(255),
    )

    rank: Mapped[str] = mapped_column(
        String(255),
    )

    position: Mapped[str] = mapped_column(
        String(255),
    )

    department: Mapped[str] = mapped_column(
        String(255),
    )

    status: Mapped[str] = mapped_column(
        String(30),
        default="PENDING",
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(timezone.utc),
    )

    processed_at: Mapped[datetime | None] = mapped_column(
        DateTime,
        nullable=True,
    )


class Criterion(Base):
    __tablename__ = "criteria"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
    )

    organization: Mapped[str] = mapped_column(
        String(50),
    )

    title: Mapped[str] = mapped_column(
        String(255),
    )

    description: Mapped[str] = mapped_column(
        Text,
    )

    enabled: Mapped[bool] = mapped_column(
        Boolean,
        default=True,
    )


class Report(Base):
    __tablename__ = "reports"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
    )

    telegram_id: Mapped[int] = mapped_column(
        BigInteger,
    )

    organization: Mapped[str] = mapped_column(
        String(50),
    )

    target: Mapped[str] = mapped_column(
        String(255),
    )

    author: Mapped[str] = mapped_column(
        String(255),
    )

    task_points: Mapped[str] = mapped_column(
        String(255),
    )

    evidence: Mapped[str] = mapped_column(
        Text,
    )

    signature: Mapped[str] = mapped_column(
        String(255),
    )

    status: Mapped[str] = mapped_column(
        String(30),
        default="PENDING",
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(timezone.utc),
    )

    processed_at: Mapped[datetime | None] = mapped_column(
        DateTime,
        nullable=True,
    )


class Banner(Base):
    __tablename__ = "banners"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
    )

    section: Mapped[str] = mapped_column(
        String(100),
        unique=True,
    )

    file_id: Mapped[str] = mapped_column(
        Text,
    )

    caption: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    enabled: Mapped[bool] = mapped_column(
        Boolean,
        default=True,
    )

    updated_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(timezone.utc),
    )


class AuditLog(Base):
    __tablename__ = "audit_logs"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
    )

    actor_id: Mapped[int] = mapped_column(
        BigInteger,
    )

    action: Mapped[str] = mapped_column(
        String(255),
    )

    target_id: Mapped[int | None] = mapped_column(
        BigInteger,
        nullable=True,
    )

    details: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=lambda: datetime.now(timezone.utc),
    )


# ============================================================
# FSM
# ============================================================

class ApplyStates(StatesGroup):
    organization = State()
    nickname = State()
    rank = State()
    position = State()
    department = State()
    confirmation = State()


class ReportStates(StatesGroup):
    target = State()
    author = State()
    points = State()
    evidence = State()
    signature = State()
    confirmation = State()


class CriterionStates(StatesGroup):
    organization = State()
    title = State()
    description = State()


class BannerStates(StatesGroup):
    section = State()
    photo = State()
    caption = State()


# ============================================================
# BOT
# ============================================================

dp = Dispatcher(storage=MemoryStorage())

bot = Bot(
    token=BOT_TOKEN,
)


# ============================================================
# HELPERS
# ============================================================

async def get_user(
    db,
    telegram_id: int,
):
    result = await db.execute(
        select(User).where(
            User.telegram_id == telegram_id
        )
    )

    return result.scalar_one_or_none()


async def is_owner(
    telegram_id: int,
) -> bool:
    return telegram_id == OWNER_ID


async def is_admin_user(
    db,
    telegram_id: int,
) -> bool:

    if telegram_id == OWNER_ID:
        return True

    user = await get_user(db, telegram_id)

    return bool(
        user
        and user.role in {"OWNER", "ADMIN"}
    )


async def audit(
    db,
    actor_id: int,
    action: str,
    target_id: int | None = None,
    details: str | None = None,
):
    db.add(
        AuditLog(
            actor_id=actor_id,
            action=action,
            target_id=target_id,
            details=details,
        )
    )


async def subscribed(
    telegram_id: int,
) -> bool:

    try:

        member = await bot.get_chat_member(
            CHANNEL_USERNAME,
            telegram_id,
        )

        return member.status in {
            ChatMemberStatus.MEMBER,
            ChatMemberStatus.ADMINISTRATOR,
            ChatMemberStatus.CREATOR,
        }

    except Exception as exc:

        logging.warning(
            "Subscription check error: %s",
            exc,
        )

        return False


def main_menu():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📝 Подать заявку",
                    callback_data="apply",
                )
            ],
            [
                InlineKeyboardButton(
                    text="👤 Профиль",
                    callback_data="profile",
                ),
                InlineKeyboardButton(
                    text="🏢 Моя организация",
                    callback_data="org",
                ),
            ],
            [
                InlineKeyboardButton(
                    text="📈 Критерии",
                    callback_data="criteria",
                ),
                InlineKeyboardButton(
                    text="📄 Рапорт",
                    callback_data="report",
                ),
            ],
            [
                InlineKeyboardButton(
                    text="🛠 Сервисы",
                    callback_data="services",
                ),
                InlineKeyboardButton(
                    text="ℹ️ Помощь",
                    callback_data="help",
                ),
            ],
        ]
    )


def admin_menu():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📥 Заявки",
                    callback_data="admin_apps",
                ),
                InlineKeyboardButton(
                    text="📄 Рапорты",
                    callback_data="admin_reports",
                ),
            ],
            [
                InlineKeyboardButton(
                    text="👥 Персонал",
                    callback_data="admin_people",
                ),
                InlineKeyboardButton(
                    text="📊 Excel",
                    callback_data="admin_excel",
                ),
            ],
            [
                InlineKeyboardButton(
                    text="🏢 Организации",
                    callback_data="admin_orgs",
                ),
            ],
            [
                InlineKeyboardButton(
                    text="📈 Критерии",
                    callback_data="admin_criteria",
                ),
                InlineKeyboardButton(
                    text="🖼 Баннер",
                    callback_data="admin_banner",
                ),
            ],
        ]
    )


def organization_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="ФСБ",
                    callback_data="apply_org:ФСБ",
                ),
                InlineKeyboardButton(
                    text="ВЧ",
                    callback_data="apply_org:ВЧ",
                ),
            ],
            [
                InlineKeyboardButton(
                    text="ЕСС",
                    callback_data="apply_org:ЕСС",
                ),
                InlineKeyboardButton(
                    text="УМВД",
                    callback_data="apply_org:УМВД",
                ),
            ],
        ]
    )


def application_confirm_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="✅ Отправить",
                    callback_data="apply_confirm",
                ),
                InlineKeyboardButton(
                    text="❌ Отмена",
                    callback_data="apply_cancel",
                ),
            ]
        ]
    )


def report_confirm_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="✅ Отправить рапорт",
                    callback_data="report_confirm",
                ),
                InlineKeyboardButton(
                    text="❌ Отмена",
                    callback_data="report_cancel",
                ),
            ]
        ]
    )


# ============================================================
# START
# ============================================================

@dp.message(Command("start"))
async def start(
    message: Message,
    state: FSMContext,
):

    await state.clear()

    if not await subscribed(
        message.from_user.id
    ):

        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="📢 Подписаться",
                        url=(
                            "https://t.me/"
                            + CHANNEL_USERNAME.lstrip("@")
                        ),
                    )
                ],
                [
                    InlineKeyboardButton(
                        text="✅ Я подписался",
                        callback_data="check_sub",
                    )
                ],
            ]
        )

        await message.answer(
            "🔐 <b>NEMAZING RP</b>\n\n"
            "Для входа необходимо подписаться "
            "на официальный канал проекта.",
            reply_markup=keyboard,
        )

        return

    async with SessionLocal() as db:

        user = await get_user(
            db,
            message.from_user.id,
        )

        if not user:

            user = User(
                telegram_id=message.from_user.id,
                username=message.from_user.username,
                first_name=message.from_user.first_name,
                role=(
                    "OWNER"
                    if message.from_user.id == OWNER_ID
                    else "PLAYER"
                ),
            )

            db.add(user)

        else:

            user.username = message.from_user.username
            user.first_name = message.from_user.first_name

        await db.commit()

    await message.answer(
        "🔥 <b>NEMAZING RP</b>\n\n"
        "Добро пожаловать в RP-проект.\n\n"
        "Используй меню ниже.",
        reply_markup=main_menu(),
    )


@dp.callback_query(F.data == "check_sub")
async def check_subscription(
    callback: CallbackQuery,
):

    if not await subscribed(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Подписка не найдена.",
            show_alert=True,
        )

        return

    await callback.message.edit_text(
        "✅ <b>Подписка подтверждена.</b>\n\n"
        "Добро пожаловать в NEMAZING RP.",
        reply_markup=main_menu(),
    )

    await callback.answer()


# ============================================================
# APPLICATION
# ============================================================

@dp.callback_query(F.data == "apply")
async def apply_start(
    callback: CallbackQuery,
    state: FSMContext,
):

    await state.clear()

    await state.set_state(
        ApplyStates.organization
    )

    await callback.message.answer(
        "📝 <b>Заявка в организацию</b>\n\n"
        "Выберите организацию:",
        reply_markup=organization_keyboard(),
    )

    await callback.answer()


@dp.callback_query(
    F.data.startswith("apply_org:")
)
async def application_organization(
    callback: CallbackQuery,
    state: FSMContext,
):

    organization = callback.data.split(
        ":",
        1,
    )[1]

    await state.update_data(
        organization=organization
    )

    await state.set_state(
        ApplyStates.nickname
    )

    await callback.message.answer(
        "👤 Введите ваш RP-никнейм:"
    )

    await callback.answer()


@dp.message(ApplyStates.nickname)
async def application_nickname(
    message: Message,
    state: FSMContext,
):

    if not message.text:
        await message.answer(
            "❌ Отправьте текст."
        )
        return

    await state.update_data(
        nickname=message.text.strip()
    )

    await state.set_state(
        ApplyStates.rank
    )

    await message.answer(
        "🎖 Введите звание:"
    )


@dp.message(ApplyStates.rank)
async def application_rank(
    message: Message,
    state: FSMContext,
):

    if not message.text:
        await message.answer(
            "❌ Отправьте текст."
        )
        return

    await state.update_data(
        rank=message.text.strip()
    )

    await state.set_state(
        ApplyStates.position
    )

    await message.answer(
        "💼 Введите должность:"
    )


@dp.message(ApplyStates.position)
async def application_position(
    message: Message,
    state: FSMContext,
):

    if not message.text:
        await message.answer(
            "❌ Отправьте текст."
        )
        return

    await state.update_data(
        position=message.text.strip()
    )

    await state.set_state(
        ApplyStates.department
    )

    await message.answer(
        "🏷 Введите отдел / подразделение:"
    )


@dp.message(ApplyStates.department)
async def application_department(
    message: Message,
    state: FSMContext,
):

    if not message.text:
        await message.answer(
            "❌ Отправьте текст."
        )
        return

    await state.update_data(
        department=message.text.strip()
    )

    data = await state.get_data()

    text = (
        "📋 <b>Проверьте заявку</b>\n\n"
        f"🏢 Организация: <b>{data['organization']}</b>\n"
        f"👤 Никнейм: <b>{data['nickname']}</b>\n"
        f"🎖 Звание: <b>{data['rank']}</b>\n"
        f"💼 Должность: <b>{data['position']}</b>\n"
        f"🏷 Отдел: <b>{data['department']}</b>\n\n"
        "Всё верно?"
    )

    await state.set_state(
        ApplyStates.confirmation
    )

    await message.answer(
        text,
        reply_markup=application_confirm_keyboard(),
    )


@dp.callback_query(
    F.data == "apply_cancel"
)
async def application_cancel(
    callback: CallbackQuery,
    state: FSMContext,
):

    await state.clear()

    await callback.message.answer(
        "❌ Заявка отменена.",
        reply_markup=main_menu(),
    )

    await callback.answer()


@dp.callback_query(
    F.data == "apply_confirm"
)
async def application_confirm(
    callback: CallbackQuery,
    state: FSMContext,
):

    data = await state.get_data()

    if not data:
        await callback.answer(
            "Сессия заявки истекла.",
            show_alert=True,
        )
        return

    async with SessionLocal() as db:

        existing = (
            await db.execute(
                select(Application).where(
                    Application.telegram_id
                    == callback.from_user.id,
                    Application.status == "PENDING",
                )
            )
        ).scalars().first()

        if existing:

            await state.clear()

            await callback.message.answer(
                f"⚠️ У вас уже есть активная заявка "
                f"#{existing.id}.",
                reply_markup=main_menu(),
            )

            await callback.answer()

            return

        application = Application(
            telegram_id=callback.from_user.id,
            organization=data["organization"],
            nickname=data["nickname"],
            rank=data["rank"],
            position=data["position"],
            department=data["department"],
        )

        db.add(application)

        await db.flush()

        application_id = application.id

        await audit(
            db,
            callback.from_user.id,
            "application_created",
            application_id,
            data["organization"],
        )

        await db.commit()

    await state.clear()

    await callback.message.answer(
        f"✅ <b>Заявка #{application_id} отправлена.</b>\n\n"
        "Ожидайте решения администрации.",
        reply_markup=main_menu(),
    )

    if OWNER_ID:

        admin_text = (
            f"🆕 <b>Новая заявка #{application_id}</b>\n\n"
            f"🏢 Организация: {data['organization']}\n"
            f"👤 Ник: {data['nickname']}\n"
            f"🎖 Звание: {data['rank']}\n"
            f"💼 Должность: {data['position']}\n"
            f"🏷 Отдел: {data['department']}\n"
            f"Telegram ID: <code>{callback.from_user.id}</code>\n"
            f"Username: @{callback.from_user.username}"
            if callback.from_user.username
            else
            f"🆕 <b>Новая заявка #{application_id}</b>\n\n"
            f"🏢 Организация: {data['organization']}\n"
            f"👤 Ник: {data['nickname']}\n"
            f"🎖 Звание: {data['rank']}\n"
            f"💼 Должность: {data['position']}\n"
            f"🏷 Отдел: {data['department']}\n"
            f"Telegram ID: <code>{callback.from_user.id}</code>"
        )

        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="✅ Одобрить",
                        callback_data=(
                            f"approve:{application_id}"
                        ),
                    ),
                    InlineKeyboardButton(
                        text="❌ Отклонить",
                        callback_data=(
                            f"reject:{application_id}"
                        ),
                    ),
                ]
            ]
        )

        await bot.send_message(
            OWNER_ID,
            admin_text,
            reply_markup=keyboard,
        )

    await callback.answer()


# ============================================================
# APPLICATION APPROVE / REJECT
# ============================================================

@dp.callback_query(
    F.data.startswith("approve:")
)
async def approve_application(
    callback: CallbackQuery,
):

    if not await is_owner(
        callback.from_user.id
    ):

        await callback.answer(
            "Нет доступа.",
            show_alert=True,
        )
        return

    application_id = int(
        callback.data.split(":")[1]
    )

    invite = None
    target_id = None

    async with SessionLocal() as db:

        application = await db.get(
            Application,
            application_id,
        )

        if (
            not application
            or application.status != "PENDING"
        ):

            await callback.answer(
                "Заявка уже обработана.",
                show_alert=True,
            )
            return

        target_id = application.telegram_id

        user = await get_user(
            db,
            application.telegram_id,
        )

        if not user:

            user = User(
                telegram_id=application.telegram_id,
                role="PLAYER",
            )

            db.add(user)

        user.nickname = application.nickname
        user.rank = application.rank
        user.position = application.position
        user.department = application.department
        user.organization = application.organization
        user.active = True
        user.archived = False
        user.archived_at = None
        user.joined_at = datetime.now(timezone.utc)

        application.status = "APPROVED"
        application.processed_at = datetime.now(
            timezone.utc
        )

        organization = (
            await db.execute(
                select(Organization).where(
                    Organization.code
                    == application.organization
                )
            )
        ).scalar_one_or_none()

        if organization:

            invite = organization.invite_link

        await audit(
            db,
            callback.from_user.id,
            "application_approved",
            application_id,
            application.organization,
        )

        await db.commit()

    if organization and organization.chat_id:

        try:

            created = await bot.create_chat_invite_link(
                chat_id=organization.chat_id,
                member_limit=1,
                name=(
                    f"NEMAZING-{target_id}"
                ),
            )

            invite = created.invite_link

        except Exception as exc:

            logging.warning(
                "Invite creation failed: %s",
                exc,
            )

    user_message = (
        "🎉 <b>Ваша заявка одобрена!</b>\n\n"
        f"🏢 Организация: {application.organization}\n"
        f"🎖 Звание: {application.rank}\n"
        f"💼 Должность: {application.position}\n"
        f"🏷 Отдел: {application.department}\n\n"
    )

    if invite:

        user_message += (
            f"🔗 <b>Вход в организацию:</b>\n"
            f"{invite}"
        )

    else:

        user_message += (
            "Свяжитесь с администрацией для "
            "получения ссылки."
        )

    try:

        await bot.send_message(
            target_id,
            user_message,
        )

    except Exception as exc:

        logging.warning(
            "Could not notify approved user: %s",
            exc,
        )

    try:

        await callback.message.edit_text(
            callback.message.text
            + "\n\n✅ <b>ОДОБРЕНО</b>"
        )

    except Exception:
        pass

    await callback.answer(
        "Заявка одобрена."
    )


@dp.callback_query(
    F.data.startswith("reject:")
)
async def reject_application(
    callback: CallbackQuery,
):

    if not await is_owner(
        callback.from_user.id
    ):

        await callback.answer(
            "Нет доступа.",
            show_alert=True,
        )

        return

    application_id = int(
        callback.data.split(":")[1]
    )

    telegram_id = None

    async with SessionLocal() as db:

        application = await db.get(
            Application,
            application_id,
        )

        if not application:

            await callback.answer(
                "Заявка не найдена.",
                show_alert=True,
            )

            return

        if application.status != "PENDING":

            await callback.answer(
                "Заявка уже обработана.",
                show_alert=True,
            )

            return

        application.status = "REJECTED"
        application.processed_at = datetime.now(
            timezone.utc
        )

        telegram_id = application.telegram_id

        await audit(
            db,
            callback.from_user.id,
            "application_rejected",
            application_id,
        )

        await db.commit()

    try:

        await bot.send_message(
            telegram_id,
            f"❌ <b>Заявка #{application_id} отклонена.</b>\n\n"
            "При необходимости обратитесь к администрации.",
        )

    except Exception:
        pass

    try:

        await callback.message.edit_text(
            callback.message.text
            + "\n\n❌ <b>ОТКЛОНЕНО</b>"
        )

    except Exception:
        pass

    await callback.answer(
        "Заявка отклонена."
    )


# ============================================================
# PROFILE
# ============================================================

@dp.callback_query(F.data == "profile")
async def profile(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        user = await get_user(
            db,
            callback.from_user.id,
        )

    if not user:

        await callback.answer(
            "Сначала выполните /start.",
            show_alert=True,
        )

        return

    username = (
        f"@{user.username}"
        if user.username
        else "—"
    )

    status = (
        "🟢 Активен"
        if user.active
        else "⚪ Неактивен"
    )

    archive = (
        "Да"
        if user.archived
        else "Нет"
    )

    text = (
        "👤 <b>ПРОФИЛЬ</b>\n\n"
        f"Telegram ID: <code>{user.telegram_id}</code>\n"
        f"Username: {username}\n"
        f"Имя: {user.first_name or '—'}\n\n"
        f"RP-ник: <b>{user.nickname or '—'}</b>\n"
        f"Организация: {user.organization or '—'}\n"
        f"Звание: {user.rank or '—'}\n"
        f"Должность: {user.position or '—'}\n"
        f"Отдел: {user.department or '—'}\n\n"
        f"Роль: <code>{user.role}</code>\n"
        f"Статус: {status}\n"
        f"В архиве: {archive}"
    )

    await callback.message.answer(
        text
    )

    await callback.answer()


# ============================================================
# ORGANIZATION
# ============================================================

@dp.callback_query(F.data == "org")
async def my_organization(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        user = await get_user(
            db,
            callback.from_user.id,
        )

        if (
            not user
            or not user.organization
        ):

            await callback.answer(
                "Организация не назначена.",
                show_alert=True,
            )

            return

        result = await db.execute(
            select(User).where(
                User.organization
                == user.organization,
                User.active.is_(True),
            )
        )

        employees = result.scalars().all()

    lines = []

    for employee in employees[:100]:

        lines.append(
            f"• {employee.nickname or employee.telegram_id}"
            f" — {employee.rank or '—'}"
            f" — {employee.position or '—'}"
        )

    text = (
        f"🏢 <b>{user.organization}</b>\n\n"
        f"Активных сотрудников: "
        f"<b>{len(employees)}</b>\n\n"
        + (
            "\n".join(lines)
            if lines
            else "Состав пока пуст."
        )
    )

    await callback.message.answer(
        text
    )

    await callback.answer()


# ============================================================
# CRITERIA
# ============================================================

@dp.callback_query(F.data == "criteria")
async def criteria(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        user = await get_user(
            db,
            callback.from_user.id,
        )

        if not user or not user.organization:

            await callback.message.answer(
                "Сначала вступите в организацию."
            )

            await callback.answer()

            return

        result = await db.execute(
            select(Criterion).where(
                Criterion.organization
                == user.organization,
                Criterion.enabled.is_(True),
            )
        )

        rows = result.scalars().all()

    if not rows:

        text = (
            "📈 <b>КРИТЕРИИ ПОВЫШЕНИЯ</b>\n\n"
            "Критерии для вашей организации "
            "пока не настроены."
        )

    else:

        parts = []

        for item in rows:

            parts.append(
                f"🔹 <b>{item.title}</b>\n"
                f"{item.description}"
            )

        text = (
            "📈 <b>КРИТЕРИИ ПОВЫШЕНИЯ</b>\n\n"
            + "\n\n".join(parts)
        )

    await callback.message.answer(
        text
    )

    await callback.answer()


# ============================================================
# REPORT
# ============================================================

@dp.callback_query(F.data == "report")
async def report_start(
    callback: CallbackQuery,
    state: FSMContext,
):

    async with SessionLocal() as db:

        user = await get_user(
            db,
            callback.from_user.id,
        )

    if (
        not user
        or not user.active
        or not user.organization
    ):

        await callback.answer(
            "Рапорт доступен только активному сотруднику.",
            show_alert=True,
        )

        return

    await state.clear()

    await state.update_data(
        organization=user.organization
    )

    await state.set_state(
        ReportStates.target
    )

    await callback.message.answer(
        "📄 <b>Новый рапорт</b>\n\n"
        "1/5\n"
        "Кому предназначен рапорт?"
    )

    await callback.answer()


@dp.message(ReportStates.target)
async def report_target(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        target=message.text.strip()
    )

    await state.set_state(
        ReportStates.author
    )

    await message.answer(
        "2/5\n"
        "От кого рапорт?\n"
        "Введите RP-ник автора."
    )


@dp.message(ReportStates.author)
async def report_author(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        author=message.text.strip()
    )

    await state.set_state(
        ReportStates.points
    )

    await message.answer(
        "3/5\n"
        "Укажите выполненные задачи / баллы:"
    )


@dp.message(ReportStates.points)
async def report_points(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        points=message.text.strip()
    )

    await state.set_state(
        ReportStates.evidence
    )

    await message.answer(
        "4/5\n"
        "Укажите доказательства, ссылки "
        "или описание выполненной работы:"
    )


@dp.message(ReportStates.evidence)
async def report_evidence(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        evidence=message.text.strip()
    )

    await state.set_state(
        ReportStates.signature
    )

    await message.answer(
        "5/5\n"
        "Введите подпись:"
    )


@dp.message(ReportStates.signature)
async def report_signature(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        signature=message.text.strip()
    )

    data = await state.get_data()

    text = (
        "📄 <b>ПРОВЕРКА РАПОРТА</b>\n\n"
        f"Кому: <b>{data['target']}</b>\n"
        f"От кого: <b>{data['author']}</b>\n"
        f"Задачи/баллы: <b>{data['points']}</b>\n"
        f"Доказательства:\n{data['evidence']}\n\n"
        f"Подпись: <b>{data['signature']}</b>\n"
        f"Организация: <b>{data['organization']}</b>\n\n"
        "Отправить рапорт администрации?"
    )

    await state.set_state(
        ReportStates.confirmation
    )

    await message.answer(
        text,
        reply_markup=report_confirm_keyboard(),
    )


@dp.callback_query(
    F.data == "report_cancel"
)
async def report_cancel(
    callback: CallbackQuery,
    state: FSMContext,
):

    await state.clear()

    await callback.message.answer(
        "❌ Рапорт отменён.",
        reply_markup=main_menu(),
    )

    await callback.answer()


@dp.callback_query(
    F.data == "report_confirm"
)
async def report_confirm(
    callback: CallbackQuery,
    state: FSMContext,
):

    data = await state.get_data()

    if not data:

        await callback.answer(
            "Сессия рапорта истекла.",
            show_alert=True,
        )

        return

    async with SessionLocal() as db:

        report = Report(
            telegram_id=callback.from_user.id,
            organization=data["organization"],
            target=data["target"],
            author=data["author"],
            task_points=data["points"],
            evidence=data["evidence"],
            signature=data["signature"],
        )

        db.add(report)

        await db.flush()

        report_id = report.id

        await audit(
            db,
            callback.from_user.id,
            "report_created",
            report_id,
            data["organization"],
        )

        await db.commit()

    await state.clear()

    await callback.message.answer(
        f"✅ <b>Рапорт #{report_id} отправлен.</b>\n\n"
        "Ожидайте решения администрации.",
        reply_markup=main_menu(),
    )

    if OWNER_ID:

        admin_text = (
            f"📄 <b>Новый рапорт #{report_id}</b>\n\n"
            f"Организация: {data['organization']}\n"
            f"Кому: {data['target']}\n"
            f"От кого: {data['author']}\n"
            f"Баллы: {data['points']}\n"
            f"Доказательства:\n{data['evidence']}\n\n"
            f"Подпись: {data['signature']}\n"
            f"Telegram ID: <code>{callback.from_user.id}</code>"
        )

        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="✅ Принять",
                        callback_data=(
                            f"report_approve:{report_id}"
                        ),
                    ),
                    InlineKeyboardButton(
                        text="❌ Отклонить",
                        callback_data=(
                            f"report_reject:{report_id}"
                        ),
                    ),
                ]
            ]
        )

        await bot.send_message(
            OWNER_ID,
            admin_text,
            reply_markup=keyboard,
        )

    await callback.answer()


# ============================================================
# REPORT APPROVAL
# ============================================================

@dp.callback_query(
    F.data.startswith("report_approve:")
)
async def report_approve(
    callback: CallbackQuery,
):

    if not await is_owner(
        callback.from_user.id
    ):

        await callback.answer(
            "Нет доступа.",
            show_alert=True,
        )

        return

    report_id = int(
        callback.data.split(":")[1]
    )

    telegram_id = None

    async with SessionLocal() as db:

        report = await db.get(
            Report,
            report_id,
        )

        if not report:

            await callback.answer(
                "Рапорт не найден.",
                show_alert=True,
            )

            return

        if report.status != "PENDING":

            await callback.answer(
                "Рапорт уже обработан.",
                show_alert=True,
            )

            return

        report.status = "APPROVED"
        report.processed_at = datetime.now(
            timezone.utc
        )

        telegram_id = report.telegram_id

        await audit(
            db,
            callback.from_user.id,
            "report_approved",
            report_id,
        )

        await db.commit()

    try:

        await bot.send_message(
            telegram_id,
            f"✅ <b>Рапорт #{report_id} принят.</b>",
        )

    except Exception:
        pass

    try:

        await callback.message.edit_text(
            callback.message.text
            + "\n\n✅ <b>ПРИНЯТ</b>"
        )

    except Exception:
        pass

    await callback.answer(
        "Рапорт принят."
    )


@dp.callback_query(
    F.data.startswith("report_reject:")
)
async def report_reject(
    callback: CallbackQuery,
):

    if not await is_owner(
        callback.from_user.id
    ):

        await callback.answer(
            "Нет доступа.",
            show_alert=True,
        )

        return

    report_id = int(
        callback.data.split(":")[1]
    )

    telegram_id = None

    async with SessionLocal() as db:

        report = await db.get(
            Report,
            report_id,
        )

        if not report:

            await callback.answer(
                "Рапорт не найден.",
                show_alert=True,
            )

            return

        if report.status != "PENDING":

            await callback.answer(
                "Рапорт уже обработан.",
                show_alert=True,
            )

            return

        report.status = "REJECTED"
        report.processed_at = datetime.now(
            timezone.utc
        )

        telegram_id = report.telegram_id

        await audit(
            db,
            callback.from_user.id,
            "report_rejected",
            report_id,
        )

        await db.commit()

    try:

        await bot.send_message(
            telegram_id,
            f"❌ <b>Рапорт #{report_id} отклонён.</b>",
        )

    except Exception:
        pass

    try:

        await callback.message.edit_text(
            callback.message.text
            + "\n\n❌ <b>ОТКЛОНЁН</b>"
        )

    except Exception:
        pass

    await callback.answer(
        "Рапорт отклонён."
    )


# ============================================================
# SERVICES
# ============================================================

@dp.callback_query(F.data == "services")
async def services(
    callback: CallbackQuery,
):

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🛡 Fenix VPN",
                    url="https://t.me/FenixVpNRobot",
                )
            ],
            [
                InlineKeyboardButton(
                    text="⭐ Buy Stars",
                    url="https://t.me/Fenix_stars_bot",
                )
            ],
        ]
    )

    await callback.message.answer(
        "🛠 <b>СЕРВИСЫ</b>\n\n"
        "Дополнительные сервисы проекта:",
        reply_markup=keyboard,
    )

    await callback.answer()


# ============================================================
# HELP
# ============================================================

@dp.callback_query(F.data == "help")
async def help_menu(
    callback: CallbackQuery,
):

    text = (
        "ℹ️ <b>NEMAZING RP — ПОМОЩЬ</b>\n\n"
        "/start — главное меню\n"
        "/cancel — отменить действие\n"
        "/admin — панель администратора\n"
        "/export — экспорт персонала\n\n"
        "📝 Подать заявку — вступление в организацию.\n"
        "📄 Рапорт — отправка рапорта администрации.\n"
        "📈 Критерии — требования для повышения.\n"
        "👤 Профиль — ваши данные."
    )

    await callback.message.answer(
        text
    )

    await callback.answer()


# ============================================================
# CANCEL
# ============================================================

@dp.message(Command("cancel"))
async def cancel(
    message: Message,
    state: FSMContext,
):

    await state.clear()

    await message.answer(
        "❌ Текущее действие отменено.",
        reply_markup=main_menu(),
    )


# ============================================================
# ADMIN
# ============================================================

@dp.message(Command("admin"))
async def admin(
    message: Message,
):

    async with SessionLocal() as db:

        allowed = await is_admin_user(
            db,
            message.from_user.id,
        )

    if not allowed:

        await message.answer(
            "⛔ Доступ запрещён."
        )

        return

    await message.answer(
        "🛠 <b>АДМИН-ПАНЕЛЬ NEMAZING RP</b>",
        reply_markup=admin_menu(),
    )


# ============================================================
# ADMIN APPLICATIONS
# ============================================================

@dp.callback_query(F.data == "admin_apps")
async def admin_apps(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

        result = await db.execute(
            select(Application)
            .order_by(
                Application.id.desc()
            )
            .limit(50)
        )

        rows = result.scalars().all()

    if not rows:

        await callback.message.answer(
            "📥 Заявок нет."
        )

        await callback.answer()

        return

    parts = []

    for item in rows:

        status_icon = {
            "PENDING": "🟡",
            "APPROVED": "🟢",
            "REJECTED": "🔴",
        }.get(
            item.status,
            "⚪",
        )

        parts.append(
            f"{status_icon} <b>#{item.id}</b>\n"
            f"🏢 {item.organization}\n"
            f"👤 {item.nickname}\n"
            f"🎖 {item.rank}\n"
            f"💼 {item.position}\n"
            f"🏷 {item.department}\n"
            f"ID: <code>{item.telegram_id}</code>"
        )

    await callback.message.answer(
        "📥 <b>ЗАЯВКИ</b>\n\n"
        + "\n\n".join(parts)
    )

    await callback.answer()


# ============================================================
# ADMIN REPORTS
# ============================================================

@dp.callback_query(F.data == "admin_reports")
async def admin_reports(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

        result = await db.execute(
            select(Report)
            .order_by(
                Report.id.desc()
            )
            .limit(50)
        )

        rows = result.scalars().all()

    if not rows:

        await callback.message.answer(
            "📄 Рапортов нет."
        )

        await callback.answer()

        return

    parts = []

    for item in rows:

        icon = {
            "PENDING": "🟡",
            "APPROVED": "🟢",
            "REJECTED": "🔴",
        }.get(
            item.status,
            "⚪",
        )

        parts.append(
            f"{icon} <b>#{item.id}</b>\n"
            f"🏢 {item.organization}\n"
            f"Кому: {item.target}\n"
            f"От: {item.author}\n"
            f"Баллы: {item.task_points}\n"
            f"Статус: {item.status}"
        )

    await callback.message.answer(
        "📄 <b>РАПОРТЫ</b>\n\n"
        + "\n\n".join(parts)
    )

    await callback.answer()


# ============================================================
# ADMIN PEOPLE
# ============================================================

@dp.callback_query(F.data == "admin_people")
async def admin_people(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

        result = await db.execute(
            select(User)
            .order_by(
                User.id.desc()
            )
            .limit(150)
        )

        rows = result.scalars().all()

    if not rows:

        await callback.message.answer(
            "👥 Персонал отсутствует."
        )

        await callback.answer()

        return

    parts = []

    for user in rows:

        status = (
            "🟢"
            if user.active
            else "⚪"
        )

        archive = (
            " 📦"
            if user.archived
            else ""
        )

        parts.append(
            f"{status}{archive} "
            f"<b>{user.nickname or user.telegram_id}</b>\n"
            f"{user.organization or '—'} | "
            f"{user.rank or '—'}\n"
            f"{user.position or '—'} | "
            f"{user.department or '—'}\n"
            f"ID: <code>{user.telegram_id}</code>\n"
            f"Роль: {user.role}"
        )

    await callback.message.answer(
        "👥 <b>ПЕРСОНАЛ</b>\n\n"
        + "\n\n".join(parts)
    )

    await callback.answer()


# ============================================================
# ADMIN ORGANIZATIONS
# ============================================================

@dp.callback_query(F.data == "admin_orgs")
async def admin_orgs(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

        result = await db.execute(
            select(Organization)
        )

        rows = result.scalars().all()

    parts = []

    for organization in rows:

        chat = (
            str(organization.chat_id)
            if organization.chat_id
            else "не задан"
        )

        invite = (
            "✅"
            if organization.invite_link
            else "❌"
        )

        auto_ban = (
            "ON"
            if organization.auto_ban_on_leave
            else "OFF"
        )

        parts.append(
            f"🏢 <b>{organization.code}</b>\n"
            f"Название: {organization.title}\n"
            f"Chat ID: <code>{chat}</code>\n"
            f"Инвайт: {invite}\n"
            f"Auto-ban: {auto_ban}\n"
            f"Включена: "
            f"{'Да' if organization.enabled else 'Нет'}"
        )

    await callback.message.answer(
        "🏢 <b>ОРГАНИЗАЦИИ</b>\n\n"
        + "\n\n".join(parts)
    )

    await callback.answer()


# ============================================================
# ADMIN CRITERIA
# ============================================================

@dp.callback_query(F.data == "admin_criteria")
async def admin_criteria(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

        result = await db.execute(
            select(Criterion)
            .order_by(
                Criterion.id.desc()
            )
            .limit(100)
        )

        rows = result.scalars().all()

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="➕ Добавить критерий",
                    callback_data="criterion_add",
                )
            ]
        ]
    )

    if not rows:

        text = (
            "📈 <b>КРИТЕРИИ</b>\n\n"
            "Критериев пока нет."
        )

    else:

        text = (
            "📈 <b>КРИТЕРИИ</b>\n\n"
            + "\n\n".join(
                f"#{x.id} "
                f"<b>{x.organization}</b> — "
                f"{x.title}\n"
                f"{x.description}\n"
                f"Статус: "
                f"{'ON' if x.enabled else 'OFF'}"
                for x in rows
            )
        )

    await callback.message.answer(
        text,
        reply_markup=keyboard,
    )

    await callback.answer()


@dp.callback_query(F.data == "criterion_add")
async def criterion_add(
    callback: CallbackQuery,
    state: FSMContext,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

    await state.clear()

    await state.set_state(
        CriterionStates.organization
    )

    await callback.message.answer(
        "🏢 Введите код организации:\n\n"
        "ФСБ / ВЧ / ЕСС / УМВД"
    )

    await callback.answer()


@dp.message(CriterionStates.organization)
async def criterion_organization(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        organization=message.text.strip()
    )

    await state.set_state(
        CriterionStates.title
    )

    await message.answer(
        "Название критерия:"
    )


@dp.message(CriterionStates.title)
async def criterion_title(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        title=message.text.strip()
    )

    await state.set_state(
        CriterionStates.description
    )

    await message.answer(
        "Описание критерия:"
    )


@dp.message(CriterionStates.description)
async def criterion_description(
    message: Message,
    state: FSMContext,
):

    data = await state.get_data()

    async with SessionLocal() as db:

        criterion = Criterion(
            organization=data["organization"],
            title=data["title"],
            description=message.text.strip(),
            enabled=True,
        )

        db.add(criterion)

        await audit(
            db,
            message.from_user.id,
            "criterion_created",
            details=data["organization"],
        )

        await db.commit()

    await state.clear()

    await message.answer(
        "✅ Критерий добавлен.",
        reply_markup=main_menu(),
    )


# ============================================================
# ADMIN BANNER
# ============================================================

@dp.callback_query(F.data == "admin_banner")
async def admin_banner(
    callback: CallbackQuery,
    state: FSMContext,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

    await state.clear()

    await state.set_state(
        BannerStates.section
    )

    await callback.message.answer(
        "🖼 <b>БАННЕР</b>\n\n"
        "Введите название раздела.\n\n"
        "Например:\n"
        "home\n"
        "profile\n"
        "organization\n"
        "services"
    )

    await callback.answer()


@dp.message(BannerStates.section)
async def banner_section(
    message: Message,
    state: FSMContext,
):

    await state.update_data(
        section=message.text.strip()
    )

    await state.set_state(
        BannerStates.photo
    )

    await message.answer(
        "📷 Теперь отправьте фотографию."
    )


@dp.message(BannerStates.photo)
async def banner_photo(
    message: Message,
    state: FSMContext,
):

    if not message.photo:

        await message.answer(
            "❌ Отправьте именно фотографию."
        )

        return

    file_id = message.photo[-1].file_id

    await state.update_data(
        file_id=file_id
    )

    await state.set_state(
        BannerStates.caption
    )

    await message.answer(
        "✏️ Введите подпись к баннеру "
        "или отправьте -"
    )


@dp.message(BannerStates.caption)
async def banner_caption(
    message: Message,
    state: FSMContext,
):

    data = await state.get_data()

    caption = message.text.strip()

    if caption == "-":
        caption = None

    async with SessionLocal() as db:

        result = await db.execute(
            select(Banner).where(
                Banner.section
                == data["section"]
            )
        )

        banner = result.scalar_one_or_none()

        if banner:

            banner.file_id = data["file_id"]
            banner.caption = caption
            banner.enabled = True
            banner.updated_at = datetime.now(
                timezone.utc
            )

        else:

            db.add(
                Banner(
                    section=data["section"],
                    file_id=data["file_id"],
                    caption=caption,
                    enabled=True,
                )
            )

        await audit(
            db,
            message.from_user.id,
            "banner_updated",
            details=data["section"],
        )

        await db.commit()

    await state.clear()

    await message.answer(
        "✅ Баннер сохранён.",
        reply_markup=main_menu(),
    )


# ============================================================
# EXCEL
# ============================================================

async def make_excel() -> bytes:

    async with SessionLocal() as db:

        users = (
            await db.execute(
                select(User)
                .order_by(User.id)
            )
        ).scalars().all()

        applications = (
            await db.execute(
                select(Application)
                .order_by(Application.id)
            )
        ).scalars().all()

        reports = (
            await db.execute(
                select(Report)
                .order_by(Report.id)
            )
        ).scalars().all()

    workbook = Workbook()

    # ---------------- PERSONNEL ----------------

    personnel = workbook.active
    personnel.title = "Personnel"

    personnel.append(
        [
            "ID",
            "Организация",
            "Звание",
            "Должность",
            "Отдел",
            "Никнейм",
            "Telegram ID",
            "Username",
            "Роль",
            "Активен",
            "Архив",
            "Дата вступления",
            "Дата архива",
        ]
    )

    for user in users:

        personnel.append(
            [
                user.id,
                user.organization,
                user.rank,
                user.position,
                user.department,
                user.nickname,
                user.telegram_id,
                user.username,
                user.role,
                "ДА" if user.active else "НЕТ",
                "ДА" if user.archived else "НЕТ",
                (
                    user.joined_at.isoformat()
                    if user.joined_at
                    else ""
                ),
                (
                    user.archived_at.isoformat()
                    if user.archived_at
                    else ""
                ),
            ]
        )

    # ---------------- APPLICATIONS ----------------

    applications_sheet = workbook.create_sheet(
        "Applications"
    )

    applications_sheet.append(
        [
            "ID",
            "Telegram ID",
            "Организация",
            "Никнейм",
            "Звание",
            "Должность",
            "Отдел",
            "Статус",
            "Создана",
            "Обработана",
        ]
    )

    for item in applications:

        applications_sheet.append(
            [
                item.id,
                item.telegram_id,
                item.organization,
                item.nickname,
                item.rank,
                item.position,
                item.department,
                item.status,
                (
                    item.created_at.isoformat()
                    if item.created_at
                    else ""
                ),
                (
                    item.processed_at.isoformat()
                    if item.processed_at
                    else ""
                ),
            ]
        )

    # ---------------- REPORTS ----------------

    reports_sheet = workbook.create_sheet(
        "Reports"
    )

    reports_sheet.append(
        [
            "ID",
            "Telegram ID",
            "Организация",
            "Кому",
            "Автор",
            "Задачи/баллы",
            "Доказательства",
            "Подпись",
            "Статус",
            "Создан",
            "Обработан",
        ]
    )

    for item in reports:

        reports_sheet.append(
            [
                item.id,
                item.telegram_id,
                item.organization,
                item.target,
                item.author,
                item.task_points,
                item.evidence,
                item.signature,
                item.status,
                (
                    item.created_at.isoformat()
                    if item.created_at
                    else ""
                ),
                (
                    item.processed_at.isoformat()
                    if item.processed_at
                    else ""
                ),
            ]
        )

    output = io.BytesIO()

    workbook.save(output)

    return output.getvalue()


@dp.callback_query(F.data == "admin_excel")
async def admin_excel(
    callback: CallbackQuery,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            callback.from_user.id,
        ):

            await callback.answer(
                "Нет доступа.",
                show_alert=True,
            )

            return

    data = await make_excel()

    await callback.message.answer_document(
        BufferedInputFile(
            data,
            filename="nemazing_rp.xlsx",
        )
    )

    await callback.answer(
        "Excel сформирован."
    )


@dp.message(Command("export"))
async def export(
    message: Message,
):

    async with SessionLocal() as db:

        if not await is_admin_user(
            db,
            message.from_user.id,
        ):

            await message.answer(
                "⛔ Нет доступа."
            )

            return

    data = await make_excel()

    await message.answer_document(
        BufferedInputFile(
            data,
            filename="nemazing_rp.xlsx",
        )
    )


# ============================================================
# ADMIN SET CHAT
# ============================================================

@dp.message(Command("setchat"))
async def setchat(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=2
    )

    if len(parts) != 3:

        await message.answer(
            "Использование:\n"
            "/setchat ФСБ -100123456789"
        )

        return

    code = parts[1].strip()

    try:
        chat_id = int(parts[2].strip())
    except ValueError:

        await message.answer(
            "❌ Chat ID должен быть числом."
        )

        return

    async with SessionLocal() as db:

        result = await db.execute(
            select(Organization).where(
                Organization.code == code
            )
        )

        organization = (
            result.scalar_one_or_none()
        )

        if not organization:

            organization = Organization(
                code=code,
                title=code,
            )

            db.add(organization)

        organization.chat_id = chat_id

        await audit(
            db,
            message.from_user.id,
            "organization_chat_set",
            details=f"{code}:{chat_id}",
        )

        await db.commit()

    await message.answer(
        "✅ Чат организации сохранён."
    )


# ============================================================
# ADMIN SET INVITE
# ============================================================

@dp.message(Command("setinvite"))
async def setinvite(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=2
    )

    if len(parts) != 3:

        await message.answer(
            "Использование:\n"
            "/setinvite ФСБ https://t.me/+..."
        )

        return

    code = parts[1].strip()
    invite = parts[2].strip()

    async with SessionLocal() as db:

        result = await db.execute(
            select(Organization).where(
                Organization.code == code
            )
        )

        organization = (
            result.scalar_one_or_none()
        )

        if not organization:

            organization = Organization(
                code=code,
                title=code,
            )

            db.add(organization)

        organization.invite_link = invite

        await audit(
            db,
            message.from_user.id,
            "organization_invite_set",
            details=code,
        )

        await db.commit()

    await message.answer(
        "✅ Инвайт сохранён."
    )


# ============================================================
# ADMIN SET ROLE
# ============================================================

@dp.message(Command("setrole"))
async def setrole(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=2
    )

    if len(parts) != 3:

        await message.answer(
            "Использование:\n"
            "/setrole TELEGRAM_ID PLAYER\n"
            "/setrole TELEGRAM_ID ADMIN"
        )

        return

    try:
        telegram_id = int(parts[1])
    except ValueError:

        await message.answer(
            "❌ Неверный Telegram ID."
        )

        return

    role = parts[2].upper()

    if role not in {
        "PLAYER",
        "ADMIN",
        "OWNER",
    }:

        await message.answer(
            "❌ Роль должна быть PLAYER, ADMIN или OWNER."
        )

        return

    async with SessionLocal() as db:

        user = await get_user(
            db,
            telegram_id,
        )

        if not user:

            user = User(
                telegram_id=telegram_id,
                role=role,
            )

            db.add(user)

        else:

            user.role = role

        await audit(
            db,
            message.from_user.id,
            "user_role_changed",
            telegram_id,
            role,
        )

        await db.commit()

    await message.answer(
        f"✅ Роль пользователя {telegram_id} "
        f"изменена на {role}."
    )


# ============================================================
# ADMIN SET RANK
# ============================================================

@dp.message(Command("setrank"))
async def setrank(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=2
    )

    if len(parts) != 3:

        await message.answer(
            "Использование:\n"
            "/setrank TELEGRAM_ID Звание"
        )

        return

    try:
        telegram_id = int(parts[1])
    except ValueError:

        await message.answer(
            "❌ Неверный Telegram ID."
        )

        return

    rank = parts[2].strip()

    async with SessionLocal() as db:

        user = await get_user(
            db,
            telegram_id,
        )

        if not user:

            await message.answer(
                "❌ Пользователь не найден."
            )

            return

        user.rank = rank

        await audit(
            db,
            message.from_user.id,
            "rank_changed",
            telegram_id,
            rank,
        )

        await db.commit()

    await message.answer(
        "✅ Звание изменено."
    )


# ============================================================
# ADMIN SET POSITION
# ============================================================

@dp.message(Command("setposition"))
async def setposition(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=2
    )

    if len(parts) != 3:

        await message.answer(
            "Использование:\n"
            "/setposition TELEGRAM_ID Должность"
        )

        return

    try:
        telegram_id = int(parts[1])
    except ValueError:

        await message.answer(
            "❌ Неверный Telegram ID."
        )

        return

    async with SessionLocal() as db:

        user = await get_user(
            db,
            telegram_id,
        )

        if not user:

            await message.answer(
                "❌ Пользователь не найден."
            )

            return

        user.position = parts[2].strip()

        await audit(
            db,
            message.from_user.id,
            "position_changed",
            telegram_id,
            user.position,
        )

        await db.commit()

    await message.answer(
        "✅ Должность изменена."
    )


# ============================================================
# ADMIN SET DEPARTMENT
# ============================================================

@dp.message(Command("setdepartment"))
async def setdepartment(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=2
    )

    if len(parts) != 3:

        await message.answer(
            "Использование:\n"
            "/setdepartment TELEGRAM_ID Отдел"
        )

        return

    try:
        telegram_id = int(parts[1])
    except ValueError:

        await message.answer(
            "❌ Неверный Telegram ID."
        )

        return

    async with SessionLocal() as db:

        user = await get_user(
            db,
            telegram_id,
        )

        if not user:

            await message.answer(
                "❌ Пользователь не найден."
            )

            return

        user.department = parts[2].strip()

        await audit(
            db,
            message.from_user.id,
            "department_changed",
            telegram_id,
            user.department,
        )

        await db.commit()

    await message.answer(
        "✅ Отдел изменён."
    )


# ============================================================
# ADMIN SET ORGANIZATION
# ============================================================

@dp.message(Command("setorg"))
async def setorg(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=2
    )

    if len(parts) != 3:

        await message.answer(
            "Использование:\n"
            "/setorg TELEGRAM_ID ФСБ"
        )

        return

    try:
        telegram_id = int(parts[1])
    except ValueError:

        await message.answer(
            "❌ Неверный Telegram ID."
        )

        return

    organization = parts[2].strip()

    async with SessionLocal() as db:

        user = await get_user(
            db,
            telegram_id,
        )

        if not user:

            await message.answer(
                "❌ Пользователь не найден."
            )

            return

        user.organization = organization
        user.active = True
        user.archived = False

        await audit(
            db,
            message.from_user.id,
            "organization_changed",
            telegram_id,
            organization,
        )

        await db.commit()

    await message.answer(
        "✅ Организация изменена."
    )


# ============================================================
# ARCHIVE
# ============================================================

@dp.message(Command("archive"))
async def archive_user(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=1
    )

    if len(parts) != 2:

        await message.answer(
            "Использование:\n"
            "/archive TELEGRAM_ID"
        )

        return

    try:
        telegram_id = int(parts[1])
    except ValueError:

        await message.answer(
            "❌ Неверный Telegram ID."
        )

        return

    async with SessionLocal() as db:

        user = await get_user(
            db,
            telegram_id,
        )

        if not user:

            await message.answer(
                "❌ Пользователь не найден."
            )

            return

        user.active = False
        user.archived = True
        user.archived_at = datetime.now(
            timezone.utc
        )

        await audit(
            db,
            message.from_user.id,
            "user_archived",
            telegram_id,
        )

        await db.commit()

    await message.answer(
        "📦 Пользователь отправлен в архив."
    )


# ============================================================
# RESTORE
# ============================================================

@dp.message(Command("restore"))
async def restore_user(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    parts = message.text.split(
        maxsplit=1
    )

    if len(parts) != 2:

        await message.answer(
            "Использование:\n"
            "/restore TELEGRAM_ID"
        )

        return

    try:
        telegram_id = int(parts[1])
    except ValueError:

        await message.answer(
            "❌ Неверный Telegram ID."
        )

        return

    async with SessionLocal() as db:

        user = await get_user(
            db,
            telegram_id,
        )

        if not user:

            await message.answer(
                "❌ Пользователь не найден."
            )

            return

        user.active = True
        user.archived = False
        user.archived_at = None

        await audit(
            db,
            message.from_user.id,
            "user_restored",
            telegram_id,
        )

        await db.commit()

    await message.answer(
        "♻️ Пользователь восстановлен."
    )


# ============================================================
# ADD CRITERION COMMAND
# ============================================================

@dp.message(Command("addcriterion"))
async def addcriterion(
    message: Message,
):

    if not await is_owner(
        message.from_user.id
    ):

        return

    raw = message.text.partition(" ")[2]

    parts = [
        item.strip()
        for item in raw.split("|")
    ]

    if len(parts) < 3:

        await message.answer(
            "Использование:\n"
            "/addcriterion ФСБ | Название | Описание"
        )

        return

    organization = parts[0]
    title = parts[1]
    description = " | ".join(
        parts[2:]
    )

    async with SessionLocal() as db:

        db.add(
            Criterion(
                organization=organization,
                title=title,
                description=description,
                enabled=True,
            )
        )

        await audit(
            db,
            message.from_user.id,
            "criterion_created",
            details=organization,
        )

        await db.commit()

    await message.answer(
        "✅ Критерий добавлен."
    )


# ============================================================
# AUTO ARCHIVE / BAN WHEN LEAVE
# ============================================================

@dp.chat_member()
async def member_update(
    event: ChatMemberUpdated,
):

    new_status = (
        event.new_chat_member.status
    )

    if new_status not in {
        ChatMemberStatus.LEFT,
        ChatMemberStatus.KICKED,
    }:

        return

    async with SessionLocal() as db:

        result = await db.execute(
            select(Organization).where(
                Organization.chat_id
                == event.chat.id
            )
        )

        organization = (
            result.scalar_one_or_none()
        )

        if not organization:
            return

        user = await get_user(
            db,
            event.new_chat_member.user.id,
        )

        if (
            not user
            or user.organization
            != organization.code
        ):

            return

        user.active = False
        user.archived = True
        user.archived_at = datetime.now(
            timezone.utc
        )

        await audit(
            db,
            event.from_user.id
            if event.from_user
            else 0,
            "employee_left_organization",
            event.new_chat_member.user.id,
            organization.code,
        )

        await db.commit()

        logging.info(
            "Archived %s from %s",
            user.telegram_id,
            organization.code,
        )

        if (
            organization.auto_ban_on_leave
            and new_status
            == ChatMemberStatus.LEFT
        ):

            try:

                await bot.ban_chat_member(
                    organization.chat_id,
                    user.telegram_id,
                )

            except Exception as exc:

                logging.warning(
                    "Auto-ban failed: %s",
                    exc,
                )


# ============================================================
# DATABASE INIT
# ============================================================

async def init_db():

    async with engine.begin() as connection:

        await connection.run_sync(
            Base.metadata.create_all
        )

    async with SessionLocal() as db:

        organizations = [
            (
                "ФСБ",
                "Федеральная служба безопасности",
            ),
            (
                "ВЧ",
                "Военная часть",
            ),
            (
                "ЕСС",
                "Единая служба спасения",
            ),
            (
                "УМВД",
                "Управление МВД",
            ),
        ]

        for code, title in organizations:

            result = await db.execute(
                select(Organization).where(
                    Organization.code == code
                )
            )

            existing = (
                result.scalar_one_or_none()
            )

            if not existing:

                db.add(
                    Organization(
                        code=code,
                        title=title,
                        enabled=True,
                        auto_ban_on_leave=False,
                    )
                )

        # Create owner user automatically
        if OWNER_ID:

            owner = await get_user(
                db,
                OWNER_ID,
            )

            if not owner:

                db.add(
                    User(
                        telegram_id=OWNER_ID,
                        role="OWNER",
                        active=True,
                    )
                )

            else:

                owner.role = "OWNER"

        await db.commit()


# ============================================================
# FASTAPI LIFESPAN
# ============================================================

@asynccontextmanager
async def lifespan(
    application: FastAPI,
):

    if not BOT_TOKEN:

        logging.error(
            "BOT_TOKEN is not configured."
        )

    await init_db()

    polling_task = asyncio.create_task(
        dp.start_polling(
            bot,
            allowed_updates=[
                "message",
                "callback_query",
                "chat_member",
            ],
        )
    )

    yield

    polling_task.cancel()

    try:

        await polling_task

    except asyncio.CancelledError:
        pass

    await bot.session.close()

    await engine.dispose()


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="NEMAZING RP",
    version="2.0.0",
    lifespan=lifespan,
)


@app.get("/")
async def root():

    return {
        "service": "NEMAZING RP",
        "status": "online",
        "version": "2.0.0",
        "health": "/health",
    }


@app.get("/health")
async def health():

    return {
        "status": "ok",
        "service": "nemazing-rp-bot",
    }


# ============================================================
# LOCAL START
# ============================================================

if __name__ == "__main__":

    import uvicorn

    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=PORT,
    )

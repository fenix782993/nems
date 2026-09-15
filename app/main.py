import asyncio
import io
import os
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from aiogram import Bot, Dispatcher, F
from aiogram.enums import ChatMemberStatus
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile,
    FSInputFile,
    CallbackQuery,
    ChatMemberUpdated,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from sqlalchemy import BigInteger, Boolean, DateTime, Integer, String, Text, select
from sqlalchemy.ext.asyncio import AsyncAttrs, AsyncSession, async_sessionmaker, create_async_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from openpyxl import Workbook

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent
(BASE_DIR / "data").mkdir(parents=True, exist_ok=True)
(BASE_DIR / "exports").mkdir(parents=True, exist_ok=True)
(BANNER_DIR := BASE_DIR / "баннеры").mkdir(parents=True, exist_ok=True)
(BASE_DIR / "banners").mkdir(parents=True, exist_ok=True)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
OWNER_ID = int(os.getenv("OWNER_ID", "0") or 0)
CHANNEL_USERNAME = os.getenv("CHANNEL_USERNAME", "@nemazing").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite+aiosqlite:///./data/nemazing.db").strip()

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)

if DATABASE_URL.startswith("sqlite"):
    engine = create_async_engine(DATABASE_URL, future=True)
else:
    engine = create_async_engine(DATABASE_URL, future=True, pool_pre_ping=True)

SessionLocal = async_sessionmaker(engine, class_=AsyncSession, expire_on_commit=False)


def utcnow() -> datetime:
    # PostgreSQL columns are TIMESTAMP WITHOUT TIME ZONE in this project.
    # Keep the datetime naive and consistently UTC.
    return datetime.utcnow()


class Base(AsyncAttrs, DeclarativeBase):
    pass


class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    telegram_id: Mapped[int] = mapped_column(BigInteger, unique=True, index=True)
    username: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    first_name: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    nickname: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    position: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    department: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    organization: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    role: Mapped[str] = mapped_column(String(50), default="PLAYER")
    active: Mapped[bool] = mapped_column(Boolean, default=False)
    archived: Mapped[bool] = mapped_column(Boolean, default=False)
    joined_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    archived_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class Organization(Base):
    __tablename__ = "organizations"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(255), unique=True)
    code: Mapped[str] = mapped_column(String(50), unique=True)
    chat_id: Mapped[Optional[int]] = mapped_column(BigInteger, nullable=True)
    invite_link: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class Application(Base):
    __tablename__ = "applications"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    organization: Mapped[str] = mapped_column(String(255))
    nickname: Mapped[str] = mapped_column(String(255))
    rank: Mapped[str] = mapped_column(String(255))
    position: Mapped[str] = mapped_column(String(255))
    department: Mapped[str] = mapped_column(String(255))
    status: Mapped[str] = mapped_column(String(50), default="PENDING")
    reject_reason: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    reviewed_by: Mapped[Optional[int]] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    reviewed_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)


class Criterion(Base):
    __tablename__ = "criteria"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    organization: Mapped[str] = mapped_column(String(255))
    rank: Mapped[str] = mapped_column(String(255))
    title: Mapped[str] = mapped_column(String(255))
    description: Mapped[str] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class Report(Base):
    __tablename__ = "reports"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    to_whom: Mapped[str] = mapped_column(String(255))
    from_whom: Mapped[str] = mapped_column(String(255))
    task_points: Mapped[str] = mapped_column(Text)
    evidence: Mapped[str] = mapped_column(Text)
    signature: Mapped[str] = mapped_column(String(255))
    status: Mapped[str] = mapped_column(String(50), default="PENDING")
    reject_reason: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    reviewed_by: Mapped[Optional[int]] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    reviewed_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)


class Banner(Base):
    __tablename__ = "banners"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    section: Mapped[str] = mapped_column(String(100), unique=True)
    file_id: Mapped[str] = mapped_column(Text)
    caption: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class ApplyStates(StatesGroup):
    organization = State()
    nickname = State()
    rank = State()
    position = State()
    department = State()
    confirm = State()


class ReportStates(StatesGroup):
    to_whom = State()
    from_whom = State()
    task_points = State()
    evidence = State()
    signature = State()
    confirm = State()


class AdminCriterionStates(StatesGroup):
    organization = State()
    rank = State()
    title = State()
    description = State()


class AdminChatStates(StatesGroup):
    organization = State()
    chat_id = State()


bot = Bot(BOT_TOKEN) if BOT_TOKEN else None
dp = Dispatcher()


def find_banner() -> Optional[Path]:
    # Automatically uses the first image placed in /баннеры or /banners.
    dirs = [BANNER_DIR, BASE_DIR / "banners"]
    exts = {".png", ".jpg", ".jpeg", ".webp"}
    for directory in dirs:
        if directory.exists():
            files = sorted([p for p in directory.iterdir() if p.is_file() and p.suffix.lower() in exts])
            if files:
                return files[0]
    return None


def page_keyboard(*rows):
    return InlineKeyboardMarkup(inline_keyboard=[*rows, [InlineKeyboardButton(text="🏠 Главное меню", callback_data="home")]])


async def edit_page(call: CallbackQuery, text: str, reply_markup: Optional[InlineKeyboardMarkup] = None):
    # Home can be a photo message. Telegram does not allow edit_text on photo messages.
    if call.message and call.message.photo:
        await call.message.edit_caption(caption=text, reply_markup=reply_markup)
    else:
        await call.message.edit_text(text, reply_markup=reply_markup)


def menu_keyboard(admin: bool = False) -> InlineKeyboardMarkup:
    rows = [
        [InlineKeyboardButton(text="📝  ПОДАТЬ ЗАЯВЛЕНИЕ", callback_data="apply")],
        [InlineKeyboardButton(text="👤 Профиль", callback_data="profile"), InlineKeyboardButton(text="🏛 Организация", callback_data="organization")],
        [InlineKeyboardButton(text="📋 Критерии", callback_data="criteria"), InlineKeyboardButton(text="📄 Рапорт", callback_data="report")],
        [InlineKeyboardButton(text="🛠 Сервисы", callback_data="services"), InlineKeyboardButton(text="❓ Помощь", callback_data="help")],
    ]
    if admin:
        rows.append([InlineKeyboardButton(text="⚙️  АДМИН-ПАНЕЛЬ", callback_data="admin")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

def organizations_keyboard(prefix: str = "org") -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🛡 ФСБ", callback_data=f"{prefix}:ФСБ"), InlineKeyboardButton(text="🎖 ВЧ", callback_data=f"{prefix}:ВЧ")],
        [InlineKeyboardButton(text="🚑 ЕСС", callback_data=f"{prefix}:ЕСС"), InlineKeyboardButton(text="🚔 УМВД", callback_data=f"{prefix}:УМВД")],
        [InlineKeyboardButton(text="🏠 Главное меню", callback_data="home")],
    ])

def admin_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📥 Заявления", callback_data="admin_apps"), InlineKeyboardButton(text="📄 Рапорты", callback_data="admin_reports")],
        [InlineKeyboardButton(text="👥 Личный состав", callback_data="admin_people"), InlineKeyboardButton(text="🏛 Организации", callback_data="admin_orgs")],
        [InlineKeyboardButton(text="📊 Скачать Excel", callback_data="admin_excel"), InlineKeyboardButton(text="📌 Критерии", callback_data="admin_criteria")],
        [InlineKeyboardButton(text="🖼 Показать баннер", callback_data="admin_banner")],
        [InlineKeyboardButton(text="🏠 Главное меню", callback_data="home")],
    ])

def is_admin(user_id: int) -> bool:
    return user_id == OWNER_ID


async def get_user(session: AsyncSession, telegram_id: int) -> Optional[User]:
    result = await session.execute(select(User).where(User.telegram_id == telegram_id))
    return result.scalar_one_or_none()


async def ensure_user(session: AsyncSession, telegram_id: int, username: Optional[str], first_name: Optional[str]) -> User:
    user = await get_user(session, telegram_id)
    if user is None:
        user = User(
            telegram_id=telegram_id,
            username=username,
            first_name=first_name,
            role="OWNER" if telegram_id == OWNER_ID else "PLAYER",
            active=(telegram_id == OWNER_ID),
            archived=False,
            created_at=utcnow(),
        )
        session.add(user)
    else:
        user.username = username
        user.first_name = first_name
        if telegram_id == OWNER_ID:
            user.role = "OWNER"
    await session.commit()
    await session.refresh(user)
    return user


async def subscribed(user_id: int) -> bool:
    if not bot:
        return False
    if not CHANNEL_USERNAME:
        return True
    try:
        member = await bot.get_chat_member(CHANNEL_USERNAME, user_id)
        return member.status in {
            ChatMemberStatus.MEMBER,
            ChatMemberStatus.ADMINISTRATOR,
            ChatMemberStatus.CREATOR,
        }
    except Exception:
        return False


async def send_home(message: Message):
    async with SessionLocal() as session:
        user = await ensure_user(session, message.from_user.id, message.from_user.username, message.from_user.first_name)
    text = (
        "NEMAZING RP\n"
        "━━━━━━━━━━━━━━━━━━\n"
        "Система управления RP-персоналом\n\n"
        "Выберите нужный раздел ниже."
    )
    banner = find_banner()
    if banner:
        try:
            await message.answer_photo(FSInputFile(banner), caption=text, reply_markup=menu_keyboard(is_admin(user.telegram_id)))
            return
        except Exception:
            pass
    await message.answer(text, reply_markup=menu_keyboard(is_admin(user.telegram_id)))

@dp.message(CommandStart())
async def start(message: Message):
    if not message.from_user:
        return
    async with SessionLocal() as session:
        user = await ensure_user(session, message.from_user.id, message.from_user.username, message.from_user.first_name)
    if not await subscribed(message.from_user.id) and not is_admin(message.from_user.id):
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="📢 Подписаться", url=f"https://t.me/{CHANNEL_USERNAME.lstrip('@')}")],
            [InlineKeyboardButton(text="✅ Проверить подписку", callback_data="check_sub")],
        ])
        await message.answer("Для входа в NEMAZING RP сначала подпишитесь на канал.", reply_markup=kb)
        return
    await send_home(message)


@dp.callback_query(F.data == "check_sub")
async def check_sub(call: CallbackQuery):
    if await subscribed(call.from_user.id) or is_admin(call.from_user.id):
        await call.answer("Подписка подтверждена")
        await edit_page(call, "NEMAZING RP\n\nДоступ открыт.", reply_markup=menu_keyboard(is_admin(call.from_user.id)))
    else:
        await call.answer("Подписка не найдена", show_alert=True)


@dp.callback_query(F.data == "home")
async def home(call: CallbackQuery, state: FSMContext):
    await state.clear()
    try:
        await call.message.delete()
    except Exception:
        pass
    await send_home(call.message)
    await call.answer()

@dp.callback_query(F.data == "apply")
async def apply_start(call: CallbackQuery, state: FSMContext):
    if not await subscribed(call.from_user.id) and not is_admin(call.from_user.id):
        await call.answer("Сначала подпишитесь на канал", show_alert=True)
        return
    await state.clear()
    await state.set_state(ApplyStates.organization)
    await edit_page(call, "Выберите организацию:", reply_markup=organizations_keyboard("applyorg"))
    await call.answer()


@dp.callback_query(F.data.startswith("applyorg:"), ApplyStates.organization)
async def apply_org(call: CallbackQuery, state: FSMContext):
    org = call.data.split(":", 1)[1]
    await state.update_data(organization=org)
    await state.set_state(ApplyStates.nickname)
    await edit_page(call, "Введите RP-никнейм:")
    await call.answer()


@dp.message(ApplyStates.nickname)
async def apply_nickname(message: Message, state: FSMContext):
    await state.update_data(nickname=message.text.strip())
    await state.set_state(ApplyStates.rank)
    await message.answer("Введите желаемое звание:")


@dp.message(ApplyStates.rank)
async def apply_rank(message: Message, state: FSMContext):
    await state.update_data(rank=message.text.strip())
    await state.set_state(ApplyStates.position)
    await message.answer("Введите должность:")


@dp.message(ApplyStates.position)
async def apply_position(message: Message, state: FSMContext):
    await state.update_data(position=message.text.strip())
    await state.set_state(ApplyStates.department)
    await message.answer("Введите отдел/подразделение:")


@dp.message(ApplyStates.department)
async def apply_department(message: Message, state: FSMContext):
    await state.update_data(department=message.text.strip())
    data = await state.get_data()
    await state.set_state(ApplyStates.confirm)
    text = (
        "Проверьте заявление\n\n"
        f"Организация: {data['organization']}\n"
        f"Ник: {data['nickname']}\n"
        f"Звание: {data['rank']}\n"
        f"Должность: {data['position']}\n"
        f"Отдел: {data['department']}"
    )
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Отправить", callback_data="apply_confirm")],
        [InlineKeyboardButton(text="❌ Отмена", callback_data="home")],
    ])
    await message.answer(text, reply_markup=kb)


@dp.callback_query(F.data == "apply_confirm", ApplyStates.confirm)
async def apply_confirm(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    async with SessionLocal() as session:
        user = await ensure_user(session, call.from_user.id, call.from_user.username, call.from_user.first_name)
        result = await session.execute(select(Application).where(Application.user_id == user.id, Application.status == "PENDING"))
        if result.scalar_one_or_none():
            await call.answer("У вас уже есть заявление на рассмотрении", show_alert=True)
            await state.clear()
            return
        app = Application(user_id=user.id, **data, status="PENDING", created_at=utcnow())
        session.add(app)
        await session.commit()
        app_id = app.id
    await state.clear()
    await edit_page(call, "✅ Заявление отправлено владельцу на рассмотрение.")
    if bot and OWNER_ID:
        await bot.send_message(
            OWNER_ID,
            f"Новое заявление #{app_id}\n\n"
            f"Пользователь: {call.from_user.id} @{call.from_user.username or 'нет'}\n"
            f"Организация: {data['organization']}\nНик: {data['nickname']}\nЗвание: {data['rank']}\n"
            f"Должность: {data['position']}\nОтдел: {data['department']}",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Одобрить", callback_data=f"approve_app:{app_id}"), InlineKeyboardButton(text="❌ Отклонить", callback_data=f"reject_app:{app_id}")]
            ])
        )
    await call.answer()


@dp.callback_query(F.data.startswith("approve_app:"))
async def approve_app(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    app_id = int(call.data.split(":")[1])
    async with SessionLocal() as session:
        app = await session.get(Application, app_id)
        if not app or app.status != "PENDING":
            await call.answer("Заявление уже обработано", show_alert=True)
            return
        user = await session.get(User, app.user_id)
        if not user:
            await call.answer("Пользователь не найден", show_alert=True)
            return
        user.nickname = app.nickname
        user.rank = app.rank
        user.position = app.position
        user.department = app.department
        user.organization = app.organization
        user.active = True
        user.archived = False
        user.joined_at = utcnow()
        app.status = "APPROVED"
        app.reviewed_by = call.from_user.id
        app.reviewed_at = utcnow()
        org_result = await session.execute(select(Organization).where(Organization.name == app.organization))
        org = org_result.scalar_one_or_none()
        invite = org.invite_link if org else None
        chat_id = org.chat_id if org else None
        await session.commit()
    await call.message.edit_reply_markup(reply_markup=None)
    await call.message.answer(f"✅ Заявление #{app_id} одобрено.")
    if bot:
        text = f"✅ Заявление одобрено\n\nВас приняли в {app.organization}."
        if invite:
            text += f"\n\n🔐 Ссылка в закрытый чат:\n{invite}"
        elif chat_id:
            text += "\n\nАдминистратор ещё не настроил персональную ссылку."
        try:
            await bot.send_message(user.telegram_id, text)
        except Exception:
            pass
    await call.answer("Одобрено")


@dp.callback_query(F.data.startswith("reject_app:"))
async def reject_app(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    app_id = int(call.data.split(":")[1])
    async with SessionLocal() as session:
        app = await session.get(Application, app_id)
        if not app or app.status != "PENDING":
            await call.answer("Заявление уже обработано", show_alert=True)
            return
        app.status = "REJECTED"
        app.reviewed_by = call.from_user.id
        app.reviewed_at = utcnow()
        user = await session.get(User, app.user_id)
        await session.commit()
    await call.message.edit_reply_markup(reply_markup=None)
    if user and bot:
        try:
            await bot.send_message(user.telegram_id, f"❌ Заявление #{app_id} отклонено владельцем.")
        except Exception:
            pass
    await call.answer("Отклонено")


@dp.callback_query(F.data == "profile")
async def profile(call: CallbackQuery):
    async with SessionLocal() as session:
        user = await ensure_user(session, call.from_user.id, call.from_user.username, call.from_user.first_name)
    await edit_page(call, 
        "👤 Профиль\n\n"
        f"ID: {user.telegram_id}\n"
        f"Ник: {user.nickname or '—'}\n"
        f"Организация: {user.organization or '—'}\n"
        f"Звание: {user.rank or '—'}\n"
        f"Должность: {user.position or '—'}\n"
        f"Отдел: {user.department or '—'}\n"
        f"Статус: {'АКТИВЕН' if user.active and not user.archived else 'АРХИВ'}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🏠 Назад", callback_data="home")]]),
    )
    await call.answer()


@dp.callback_query(F.data == "organization")
async def organization(call: CallbackQuery):
    async with SessionLocal() as session:
        user = await get_user(session, call.from_user.id)
        if not user or not user.organization:
            text = "Вы пока не состоите в организации."
        else:
            result = await session.execute(select(User).where(User.organization == user.organization, User.active == True, User.archived == False))
            people = result.scalars().all()
            lines = [f"🏛 {user.organization}", ""]
            for p in people[:80]:
                lines.append(f"• {p.nickname or p.first_name or 'Без имени'} — {p.rank or 'без звания'}")
            text = "\n".join(lines)
    await edit_page(call, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🏠 Назад", callback_data="home")]]))
    await call.answer()


@dp.callback_query(F.data == "criteria")
async def criteria(call: CallbackQuery):
    async with SessionLocal() as session:
        result = await session.execute(select(Criterion).order_by(Criterion.organization, Criterion.id))
        items = result.scalars().all()
    if not items:
        text = "📋 Критерии повышения\n\nКритерии пока не добавлены владельцем."
    else:
        parts = ["📋 Критерии повышения"]
        for c in items[:50]:
            parts.append(f"\n{c.organization} — {c.rank}\n{c.title}\n{c.description}")
        text = "\n".join(parts)
    await edit_page(call, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🏠 Назад", callback_data="home")]]))
    await call.answer()


@dp.callback_query(F.data == "report")
async def report_start(call: CallbackQuery, state: FSMContext):
    async with SessionLocal() as session:
        user = await get_user(session, call.from_user.id)
    if not user or not user.active or user.archived:
        await call.answer("Рапорт доступен только действующему составу", show_alert=True)
        return
    await state.clear()
    await state.set_state(ReportStates.to_whom)
    await edit_page(call, "Рапорт о повышении\n\nУкажите, кому рапорт:")
    await call.answer()


@dp.message(ReportStates.to_whom)
async def report_to(message: Message, state: FSMContext):
    await state.update_data(to_whom=message.text.strip())
    await state.set_state(ReportStates.from_whom)
    await message.answer("От кого рапорт:")


@dp.message(ReportStates.from_whom)
async def report_from(message: Message, state: FSMContext):
    await state.update_data(from_whom=message.text.strip())
    await state.set_state(ReportStates.task_points)
    await message.answer("Укажите выполненные задачи и баллы:")


@dp.message(ReportStates.task_points)
async def report_tasks(message: Message, state: FSMContext):
    await state.update_data(task_points=message.text.strip())
    await state.set_state(ReportStates.evidence)
    await message.answer("Прикрепите ссылки/доказательства текстом:")


@dp.message(ReportStates.evidence)
async def report_evidence(message: Message, state: FSMContext):
    await state.update_data(evidence=message.text.strip())
    await state.set_state(ReportStates.signature)
    await message.answer("Введите подпись:")


@dp.message(ReportStates.signature)
async def report_signature(message: Message, state: FSMContext):
    await state.update_data(signature=message.text.strip())
    data = await state.get_data()
    await state.set_state(ReportStates.confirm)
    await message.answer(
        "Проверьте рапорт\n\n"
        f"Кому: {data['to_whom']}\nОт кого: {data['from_whom']}\n"
        f"Задачи/баллы: {data['task_points']}\nДоказательства: {data['evidence']}\nПодпись: {data['signature']}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Отправить", callback_data="report_confirm")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="home")],
        ]),
    )


@dp.callback_query(F.data == "report_confirm", ReportStates.confirm)
async def report_confirm(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    async with SessionLocal() as session:
        user = await get_user(session, call.from_user.id)
        if not user:
            await call.answer("Пользователь не найден", show_alert=True)
            return
        report = Report(user_id=user.id, **data, status="PENDING", created_at=utcnow())
        session.add(report)
        await session.commit()
        report_id = report.id
    await state.clear()
    await edit_page(call, f"✅ Рапорт #{report_id} отправлен владельцу.")
    if bot and OWNER_ID:
        await bot.send_message(
            OWNER_ID,
            f"Новый рапорт #{report_id}\n\n"
            f"От: {call.from_user.id}\nКому: {data['to_whom']}\nОт кого: {data['from_whom']}\n"
            f"Задачи: {data['task_points']}\nДоказательства: {data['evidence']}\nПодпись: {data['signature']}",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Принять", callback_data=f"approve_report:{report_id}"), InlineKeyboardButton(text="❌ Отклонить", callback_data=f"reject_report:{report_id}")]
            ]),
        )
    await call.answer()


async def process_report(call: CallbackQuery, report_id: int, approved: bool):
    async with SessionLocal() as session:
        report = await session.get(Report, report_id)
        if not report or report.status != "PENDING":
            await call.answer("Рапорт уже обработан", show_alert=True)
            return
        report.status = "APPROVED" if approved else "REJECTED"
        report.reviewed_by = call.from_user.id
        report.reviewed_at = utcnow()
        user = await session.get(User, report.user_id)
        await session.commit()
    if user and bot:
        try:
            await bot.send_message(user.telegram_id, f"{'✅ Рапорт принят.' if approved else '❌ Рапорт отклонён.'}\nНомер: #{report_id}")
        except Exception:
            pass
    await call.message.edit_reply_markup(reply_markup=None)
    await call.answer("Принято" if approved else "Отклонено")


@dp.callback_query(F.data.startswith("approve_report:"))
async def approve_report(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    await process_report(call, int(call.data.split(":")[1]), True)


@dp.callback_query(F.data.startswith("reject_report:"))
async def reject_report(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    await process_report(call, int(call.data.split(":")[1]), False)


@dp.callback_query(F.data == "services")
async def services(call: CallbackQuery):
    await edit_page(call, 
        "🛠 Сервисы\n\n"
        "🔐 Fenix VPN — @FenixVpNRobot\n"
        "⭐ Buy Stars — @Fenix_stars_bot",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🏠 Назад", callback_data="home")]]),
    )
    await call.answer()


@dp.callback_query(F.data == "help")
async def help_page(call: CallbackQuery):
    await edit_page(call, 
        "❓ Помощь\n\n"
        "1. Подпишитесь на канал.\n"
        "2. Подайте заявление.\n"
        "3. Дождитесь решения владельца.\n"
        "4. После принятия получите доступ в организационный чат.\n"
        "5. Действующий сотрудник может отправлять рапорты на повышение.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🏠 Назад", callback_data="home")]]),
    )
    await call.answer()


@dp.callback_query(F.data == "admin")
async def admin_panel(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    await edit_page(call, "⚙️ Админ-панель NEMAZING RP", reply_markup=admin_keyboard())
    await call.answer()


@dp.callback_query(F.data == "admin_apps")
async def admin_apps(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        return
    async with SessionLocal() as session:
        result = await session.execute(select(Application).where(Application.status == "PENDING").order_by(Application.id.desc()).limit(20))
        apps = result.scalars().all()
    if not apps:
        text = "📥 Заявления\n\nНовых заявлений нет."
    else:
        lines = ["📥 Заявления"]
        for a in apps:
            lines.append(f"\n#{a.id} — {a.organization}\n{a.nickname} | {a.rank} | {a.position}")
        text = "\n".join(lines)
    await edit_page(call, text, reply_markup=admin_keyboard())
    await call.answer()


@dp.callback_query(F.data == "admin_reports")
async def admin_reports(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        return
    async with SessionLocal() as session:
        result = await session.execute(select(Report).where(Report.status == "PENDING").order_by(Report.id.desc()).limit(20))
        reports = result.scalars().all()
    if not reports:
        text = "📄 Рапорты\n\nНовых рапортов нет."
    else:
        text = "📄 Рапорты\n\n" + "\n".join(f"#{r.id} — {r.to_whom} — {r.from_whom}" for r in reports)
    await edit_page(call, text, reply_markup=admin_keyboard())
    await call.answer()


@dp.callback_query(F.data == "admin_people")
async def admin_people(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        return
    async with SessionLocal() as session:
        result = await session.execute(select(User).order_by(User.id.desc()).limit(50))
        users = result.scalars().all()
    lines = ["👥 Личный состав"]
    for u in users:
        lines.append(f"#{u.id} | {u.nickname or u.first_name or '—'} | {u.organization or '—'} | {u.role} | {'ACTIVE' if u.active else 'ARCHIVE'}")
    await edit_page(call, "\n".join(lines), reply_markup=admin_keyboard())
    await call.answer()


@dp.callback_query(F.data == "admin_orgs")
async def admin_orgs(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        return
    async with SessionLocal() as session:
        result = await session.execute(select(Organization).order_by(Organization.id))
        orgs = result.scalars().all()
    lines = ["🏛 Организации"]
    for o in orgs:
        lines.append(f"• {o.name} | chat: {o.chat_id or '—'} | invite: {'есть' if o.invite_link else 'нет'}")
    await edit_page(call, "\n".join(lines), reply_markup=admin_keyboard())
    await call.answer()


@dp.callback_query(F.data == "admin_criteria")
async def admin_criteria(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        return
    async with SessionLocal() as session:
        result = await session.execute(select(Criterion).order_by(Criterion.id.desc()).limit(30))
        criteria_items = result.scalars().all()
    text = "📌 Критерии\n\n" + ("\n".join(f"#{c.id} {c.organization} — {c.rank}: {c.title}" for c in criteria_items) if criteria_items else "Критерии отсутствуют.")
    await edit_page(call, text, reply_markup=admin_keyboard())
    await call.answer()


@dp.callback_query(F.data == "admin_excel")
async def admin_excel(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    async with SessionLocal() as session:
        users = (await session.execute(select(User).order_by(User.id))).scalars().all()
        apps = (await session.execute(select(Application).order_by(Application.id))).scalars().all()
        reports = (await session.execute(select(Report).order_by(Report.id))).scalars().all()
    wb = Workbook()
    ws = wb.active
    ws.title = "Personnel"
    ws.append(["ID", "Telegram ID", "Username", "Nickname", "Organization", "Rank", "Position", "Department", "Role", "Active", "Archived", "Joined"])
    for u in users:
        ws.append([u.id, u.telegram_id, u.username, u.nickname, u.organization, u.rank, u.position, u.department, u.role, u.active, u.archived, u.joined_at])
    wa = wb.create_sheet("Applications")
    wa.append(["ID", "User ID", "Organization", "Nickname", "Rank", "Position", "Department", "Status", "Created"])
    for a in apps:
        wa.append([a.id, a.user_id, a.organization, a.nickname, a.rank, a.position, a.department, a.status, a.created_at])
    wr = wb.create_sheet("Reports")
    wr.append(["ID", "User ID", "To", "From", "Tasks", "Evidence", "Signature", "Status", "Created"])
    for r in reports:
        wr.append([r.id, r.user_id, r.to_whom, r.from_whom, r.task_points, r.evidence, r.signature, r.status, r.created_at])
    stream = io.BytesIO()
    wb.save(stream)
    data_bytes = stream.getvalue()
    export_path = BASE_DIR / "exports" / "nemazing_personnel.xlsx"
    export_path.write_bytes(data_bytes)
    await call.message.answer_document(
        BufferedInputFile(data_bytes, filename="nemazing_personnel.xlsx"),
        caption="📊 Excel-выгрузка NEMAZING RP\nЛичный состав • заявления • рапорты"
    )
    await call.answer("Excel сформирован")


@dp.callback_query(F.data == "admin_banner")
async def admin_banner(call: CallbackQuery):
    if not is_admin(call.from_user.id):
        await call.answer("Нет доступа", show_alert=True)
        return
    banner = find_banner()
    if not banner:
        await call.answer("В папке баннеры нет PNG/JPG/WEBP", show_alert=True)
        return
    await call.message.answer_photo(FSInputFile(banner), caption="🖼 Баннер NEMAZING RP")
    await call.answer("Баннер отправлен")


@dp.message(Command("setchat"))
async def setchat(message: Message):
    if not is_admin(message.from_user.id):
        return
    parts = message.text.split(maxsplit=2)
    if len(parts) < 3:
        await message.answer("Использование: /setchat ФСБ -100123456789")
        return
    org_name, raw_chat = parts[1], parts[2]
    try:
        chat_id = int(raw_chat)
    except ValueError:
        await message.answer("chat_id должен быть числом")
        return
    async with SessionLocal() as session:
        result = await session.execute(select(Organization).where(Organization.name == org_name))
        org = result.scalar_one_or_none()
        if not org:
            await message.answer("Организация не найдена")
            return
        org.chat_id = chat_id
        await session.commit()
    await message.answer(f"✅ Чат {org_name} сохранён: {chat_id}")


@dp.message(Command("setinvite"))
async def setinvite(message: Message):
    if not is_admin(message.from_user.id):
        return
    parts = message.text.split(maxsplit=2)
    if len(parts) < 3:
        await message.answer("Использование: /setinvite ФСБ https://t.me/+...")
        return
    org_name, invite = parts[1], parts[2]
    async with SessionLocal() as session:
        result = await session.execute(select(Organization).where(Organization.name == org_name))
        org = result.scalar_one_or_none()
        if not org:
            await message.answer("Организация не найдена")
            return
        org.invite_link = invite
        await session.commit()
    await message.answer("✅ Ссылка сохранена")


@dp.message(Command("setrank"))
async def setrank(message: Message):
    if not is_admin(message.from_user.id):
        return
    parts = message.text.split(maxsplit=2)
    if len(parts) < 3:
        await message.answer("Использование: /setrank TELEGRAM_ID Новое звание")
        return
    try:
        tg_id = int(parts[1])
    except ValueError:
        await message.answer("Telegram ID должен быть числом")
        return
    async with SessionLocal() as session:
        user = await get_user(session, tg_id)
        if not user:
            await message.answer("Пользователь не найден")
            return
        user.rank = parts[2]
        await session.commit()
    await message.answer("✅ Звание изменено")


@dp.message(Command("setrole"))
async def setrole(message: Message):
    if not is_admin(message.from_user.id):
        return
    parts = message.text.split(maxsplit=2)
    if len(parts) < 3 or parts[2].upper() not in {"OWNER", "ADMIN", "PLAYER"}:
        await message.answer("Использование: /setrole TELEGRAM_ID OWNER|ADMIN|PLAYER")
        return
    try:
        tg_id = int(parts[1])
    except ValueError:
        await message.answer("Telegram ID должен быть числом")
        return
    async with SessionLocal() as session:
        user = await get_user(session, tg_id)
        if not user:
            await message.answer("Пользователь не найден")
            return
        user.role = parts[2].upper()
        await session.commit()
    await message.answer("✅ Роль изменена")


@dp.message(Command("addcriterion"))
async def addcriterion(message: Message):
    if not is_admin(message.from_user.id):
        return
    parts = message.text.split(maxsplit=4)
    if len(parts) < 5:
        await message.answer("Использование: /addcriterion ФСБ Капитан Название Описание")
        return
    async with SessionLocal() as session:
        session.add(Criterion(organization=parts[1], rank=parts[2], title=parts[3], description=parts[4], created_at=utcnow()))
        await session.commit()
    await message.answer("✅ Критерий добавлен")


@dp.message(Command("archive"))
async def archive_user(message: Message):
    if not is_admin(message.from_user.id):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /archive TELEGRAM_ID")
        return
    async with SessionLocal() as session:
        user = await get_user(session, int(parts[1]))
        if not user:
            await message.answer("Пользователь не найден")
            return
        user.active = False
        user.archived = True
        user.archived_at = utcnow()
        await session.commit()
    await message.answer("✅ Пользователь отправлен в архив")


@dp.message(Command("restore"))
async def restore_user(message: Message):
    if not is_admin(message.from_user.id):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /restore TELEGRAM_ID")
        return
    async with SessionLocal() as session:
        user = await get_user(session, int(parts[1]))
        if not user:
            await message.answer("Пользователь не найден")
            return
        user.active = True
        user.archived = False
        user.archived_at = None
        if not user.joined_at:
            user.joined_at = utcnow()
        await session.commit()
    await message.answer("✅ Пользователь восстановлен")


@dp.chat_member()
async def member_update(event: ChatMemberUpdated):
    if not bot:
        return
    old = event.old_chat_member.status
    new = event.new_chat_member.status
    leaving = old in {ChatMemberStatus.MEMBER, ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.RESTRICTED} and new in {ChatMemberStatus.LEFT, ChatMemberStatus.KICKED}
    if not leaving:
        return
    tg_id = event.from_user.id if event.new_chat_member.user.id == event.from_user.id else event.new_chat_member.user.id
    async with SessionLocal() as session:
        user = await get_user(session, tg_id)
        if not user:
            return
        if user.role == "OWNER":
            return
        user.active = False
        user.archived = True
        user.archived_at = utcnow()
        await session.commit()


async def _migrate_schema():
    """Small idempotent migration for existing Render/PostgreSQL databases.

    SQLAlchemy create_all() does not add missing columns to tables that already
    exist. The project previously created an organizations table with an older
    schema, so we explicitly add the current columns before ORM queries run.
    """
    if DATABASE_URL.startswith("sqlite"):
        return

    statements = [
        # users
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS username VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS first_name VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS nickname VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS rank VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS position VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS department VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS organization VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS role VARCHAR(50)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS active BOOLEAN DEFAULT FALSE",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS archived BOOLEAN DEFAULT FALSE",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS joined_at TIMESTAMP WITHOUT TIME ZONE",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS archived_at TIMESTAMP WITHOUT TIME ZONE",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITHOUT TIME ZONE",
        # organizations
        "ALTER TABLE organizations ADD COLUMN IF NOT EXISTS name VARCHAR(255)",
        "ALTER TABLE organizations ADD COLUMN IF NOT EXISTS code VARCHAR(50)",
        "ALTER TABLE organizations ADD COLUMN IF NOT EXISTS chat_id BIGINT",
        "ALTER TABLE organizations ADD COLUMN IF NOT EXISTS invite_link TEXT",
        "ALTER TABLE organizations ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITHOUT TIME ZONE",
        # applications
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS organization VARCHAR(255)",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS nickname VARCHAR(255)",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS rank VARCHAR(255)",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS position VARCHAR(255)",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS department VARCHAR(255)",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS status VARCHAR(50) DEFAULT 'PENDING'",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS reject_reason TEXT",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS reviewed_by BIGINT",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITHOUT TIME ZONE",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS reviewed_at TIMESTAMP WITHOUT TIME ZONE",
        # criteria
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS organization VARCHAR(255)",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS rank VARCHAR(255)",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS title VARCHAR(255)",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS description TEXT",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITHOUT TIME ZONE",
        # reports
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS user_id INTEGER",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS to_whom VARCHAR(255)",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS from_whom VARCHAR(255)",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS task_points TEXT",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS evidence TEXT",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS signature VARCHAR(255)",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS status VARCHAR(50) DEFAULT 'PENDING'",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS reject_reason TEXT",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS reviewed_by BIGINT",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITHOUT TIME ZONE",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS reviewed_at TIMESTAMP WITHOUT TIME ZONE",
        # banners table
        "ALTER TABLE banners ADD COLUMN IF NOT EXISTS section VARCHAR(100)",
        "ALTER TABLE banners ADD COLUMN IF NOT EXISTS file_id TEXT",
        "ALTER TABLE banners ADD COLUMN IF NOT EXISTS caption TEXT",
        "ALTER TABLE banners ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITHOUT TIME ZONE",
    ]
    async with engine.begin() as conn:
        for sql in statements:
            await conn.exec_driver_sql(sql)

        # Repair rows created by an older version. We only fill NULL values;
        # existing administrator data is preserved.
        await conn.exec_driver_sql("UPDATE users SET role = COALESCE(role, 'PLAYER')")
        await conn.exec_driver_sql("UPDATE users SET active = COALESCE(active, FALSE)")
        await conn.exec_driver_sql("UPDATE users SET archived = COALESCE(archived, FALSE)")
        await conn.exec_driver_sql("UPDATE users SET created_at = COALESCE(created_at, CURRENT_TIMESTAMP)")
        await conn.exec_driver_sql("UPDATE applications SET status = COALESCE(status, 'PENDING')")
        await conn.exec_driver_sql("UPDATE applications SET created_at = COALESCE(created_at, CURRENT_TIMESTAMP)")
        await conn.exec_driver_sql("UPDATE criteria SET created_at = COALESCE(created_at, CURRENT_TIMESTAMP)")
        await conn.exec_driver_sql("UPDATE reports SET status = COALESCE(status, 'PENDING')")
        await conn.exec_driver_sql("UPDATE reports SET created_at = COALESCE(created_at, CURRENT_TIMESTAMP)")

        await conn.exec_driver_sql(
            "UPDATE organizations SET name = COALESCE(name, 'Организация ' || id::text)"
        )
        await conn.exec_driver_sql(
            "UPDATE organizations SET code = COALESCE(code, 'ORG' || id::text)"
        )
        await conn.exec_driver_sql(
            "UPDATE organizations SET created_at = COALESCE(created_at, CURRENT_TIMESTAMP)"
        )


async def init_db():
    # First create any completely new tables.
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)

    # Then migrate old Render/PostgreSQL tables before selecting through ORM.
    await _migrate_schema()

    async with SessionLocal() as session:
        existing = (await session.execute(select(Organization))).scalars().all()
        by_name = {x.name for x in existing if x.name}
        defaults = [("ФСБ", "FSB"), ("ВЧ", "VC"), ("ЕСС", "ESS"), ("УМВД", "UMVD")]
        for name, code in defaults:
            if name not in by_name:
                session.add(Organization(name=name, code=code, created_at=utcnow()))

        if OWNER_ID:
            owner = await get_user(session, OWNER_ID)
            if owner is None:
                session.add(
                    User(
                        telegram_id=OWNER_ID,
                        role="OWNER",
                        active=True,
                        archived=False,
                        created_at=utcnow(),
                    )
                )
            else:
                owner.role = "OWNER"
                owner.active = True
                owner.archived = False

        await session.commit()

# ============================================================
# WEB SERVER / RENDER STARTUP
# ============================================================

import logging
from contextlib import suppress

import uvicorn

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(name)s | %(message)s")
logger = logging.getLogger("nemazing")
_polling_task: Optional[asyncio.Task] = None

async def _run_bot_polling():
    if not bot:
        logger.warning("BOT_TOKEN is not configured; Telegram polling is disabled.")
        return
    while True:
        try:
            logger.info("Starting Telegram bot polling...")
            await bot.delete_webhook(drop_pending_updates=False)
            await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Telegram polling crashed; retrying in 5 seconds.")
            await asyncio.sleep(5)

@asynccontextmanager
async def lifespan(app: FastAPI):
    global _polling_task
    logger.info("NEMAZING RP starting...")
    try:
        await init_db()
        logger.info("Database initialization completed.")
    except Exception:
        logger.exception("Database initialization failed.")
    if bot:
        _polling_task = asyncio.create_task(_run_bot_polling())
    yield
    if _polling_task:
        _polling_task.cancel()
        with suppress(asyncio.CancelledError):
            await _polling_task
    if bot:
        with suppress(Exception):
            await bot.session.close()
    with suppress(Exception):
        await engine.dispose()

app = FastAPI(title="NEMAZING RP", version="1.1.0", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

@app.get("/")
async def root():
    return {"service":"NEMAZING RP","version":"1.1.0","status":"online","health":"/health","api_health":"/api/health"}

@app.get("/health")
async def health():
    return {"status":"ok","service":"nemazing-rp"}

@app.get("/api/health")
async def api_health():
    return {"status":"ok","service":"nemazing-rp","bot_configured":bool(BOT_TOKEN),"database_configured":bool(DATABASE_URL)}

@app.get("/api/status")
async def api_status():
    return {"service":"NEMAZING RP","status":"online","bot_configured":bool(BOT_TOKEN),"polling_running":bool(_polling_task and not _polling_task.done()),"owner_configured":bool(OWNER_ID),"channel":CHANNEL_USERNAME}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", "10000")), log_level="info")

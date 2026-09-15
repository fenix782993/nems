import asyncio
import io
import os
import html
import logging
from contextlib import asynccontextmanager, suppress
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile, FSInputFile, CallbackQuery, ChatMemberUpdated,
    InlineKeyboardButton, InlineKeyboardMarkup, Message
)
from sqlalchemy import BigInteger, Boolean, DateTime, Integer, String, Text, select, func
from sqlalchemy.ext.asyncio import AsyncAttrs, AsyncSession, async_sessionmaker, create_async_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import uvicorn

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent
for d in ("data", "exports", "баннеры", "banners"):
    (BASE_DIR / d).mkdir(parents=True, exist_ok=True)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
OWNER_ID = int(os.getenv("OWNER_ID", "0") or 0)
CHANNEL_USERNAME = os.getenv("CHANNEL_USERNAME", "@nemazing").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite+aiosqlite:///./data/nemazing.db").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)

engine = create_async_engine(
    DATABASE_URL,
    future=True,
    pool_pre_ping=not DATABASE_URL.startswith("sqlite"),
)
SessionLocal = async_sessionmaker(engine, class_=AsyncSession, expire_on_commit=False)

def utcnow() -> datetime:
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
    callsign: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    role: Mapped[str] = mapped_column(String(50), default="PLAYER")
    active: Mapped[bool] = mapped_column(Boolean, default=False)
    archived: Mapped[bool] = mapped_column(Boolean, default=False)
    joined_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    archived_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)
    archived_organization: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    archived_rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    archived_position: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    archived_department: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    archived_callsign: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
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
    callsign: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    status: Mapped[str] = mapped_column(String(50), default="PENDING")
    reject_reason: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    reviewed_by: Mapped[Optional[int]] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    reviewed_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

class Criterion(Base):
    __tablename__ = "criteria"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    organization: Mapped[str] = mapped_column(String(255))
    rank: Mapped[str] = mapped_column(String(255), default="")
    from_rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    to_rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    title: Mapped[str] = mapped_column(String(255))
    description: Mapped[str] = mapped_column(Text)
    min_days: Mapped[int] = mapped_column(Integer, default=0)
    required_fixations: Mapped[int] = mapped_column(Integer, default=0)
    required_tasks: Mapped[int] = mapped_column(Integer, default=0)
    no_discipline: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)

class Report(Base):
    __tablename__ = "reports"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    source_organization: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    org_number: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    criterion_id: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    criterion_from_rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    criterion_to_rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    to_whom: Mapped[str] = mapped_column(String(255))
    from_whom: Mapped[str] = mapped_column(String(255))
    task_points: Mapped[str] = mapped_column(Text)
    evidence: Mapped[str] = mapped_column(Text, default="")
    signature: Mapped[str] = mapped_column(String(255))
    status: Mapped[str] = mapped_column(String(50), default="PENDING")
    reject_reason: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    reviewed_by: Mapped[Optional[int]] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    reviewed_at: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

class ReportAttachment(Base):
    __tablename__ = "report_attachments"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    report_id: Mapped[int] = mapped_column(Integer, index=True)
    kind: Mapped[str] = mapped_column(String(30))
    file_id: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    url: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    caption: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)

class PersonnelHistory(Base):
    __tablename__ = "personnel_history"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    action: Mapped[str] = mapped_column(String(50))
    from_organization: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    to_organization: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    from_rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    to_rank: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    from_position: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    to_position: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    note: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    actor_id: Mapped[Optional[int]] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)

class Banner(Base):
    __tablename__ = "banners"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    section: Mapped[str] = mapped_column(String(100), unique=True)
    file_id: Mapped[str] = mapped_column(Text)
    caption: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)

class Notification(Base):
    __tablename__ = "notifications"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    event_type: Mapped[str] = mapped_column(String(50))
    title: Mapped[str] = mapped_column(String(255))
    text: Mapped[str] = mapped_column(Text)
    related_user_id: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    related_id: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    read: Mapped[bool] = mapped_column(Boolean, default=False)

class AdminVote(Base):
    __tablename__ = "admin_votes"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    voter_id: Mapped[int] = mapped_column(BigInteger, index=True)
    admin_name: Mapped[str] = mapped_column(String(255))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)

class Discipline(Base):
    __tablename__ = "disciplines"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    kind: Mapped[str] = mapped_column(String(100))
    reason: Mapped[str] = mapped_column(Text)
    actor_id: Mapped[Optional[int]] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)

class DisciplineStates(StatesGroup):
    value = State()

class ApplyStates(StatesGroup):
    organization = State()
    nickname = State()
    rank = State()
    position = State()
    department = State()
    callsign = State()
    confirm = State()

class ReportStates(StatesGroup):
    organization = State()
    criterion = State()
    from_whom = State()
    task_points = State()
    attachments = State()
    signature = State()
    confirm = State()

class AdminCriterionStates(StatesGroup):
    organization = State()
    from_rank = State()
    to_rank = State()
    title = State()
    description = State()
    min_days = State()
    fixations = State()
    tasks = State()
    no_discipline = State()
    confirm = State()

class AdminActionStates(StatesGroup):
    value = State()
    transfer_org = State()
    transfer_position = State()

bot = Bot(BOT_TOKEN) if BOT_TOKEN else None
dp = Dispatcher()
_polling_task = None
logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(name)s | %(message)s")
logger = logging.getLogger("nemazing")

ORGS = [
    ("ФСБ", "FSB", "🛡"),
    ("ВЧ", "VC", "🎖"),
    ("УМВД", "UMVD", "🏛"),
    ("ДПС", "DPS", "🚓"),
    ("ЕСС", "ESS", "🚑"),
    ("Правительство", "GOV", "🏛"),
]
ORG_NAMES = [x[0] for x in ORGS]

def esc(v) -> str:
    return html.escape(str(v or "—"))

def find_banner() -> Optional[Path]:
    for directory in (BASE_DIR / "баннеры", BASE_DIR / "banners"):
        if directory.exists():
            for p in sorted(directory.iterdir()):
                if p.is_file() and p.suffix.lower() in {".png",".jpg",".jpeg",".webp"}:
                    return p
    return None

def org_icon(name: str) -> str:
    return next((i for n,_,i in ORGS if n == name), "🏢")

def is_admin(user_id: int) -> bool:
    return user_id == OWNER_ID

ADMIN_VOTE_NAMES = [
    "Fenix_Rubcov", "ALeksandr_Frolov", "Mikhail_Ofiserov",
    "Giovanni_Kravec", "Lorenzo_Tape",
]

async def create_notification(event_type: str, title: str, text: str, related_user_id=None, related_id=None):
    async with SessionLocal() as session:
        session.add(Notification(event_type=event_type, title=title, text=text,
                                 related_user_id=related_user_id, related_id=related_id, created_at=utcnow()))
        await session.commit()

async def record_history(session, user_id, action, actor_id=None, **kwargs):
    session.add(PersonnelHistory(user_id=user_id, action=action, actor_id=actor_id, created_at=utcnow(), **kwargs))

def page_kb(rows) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[*rows, [InlineKeyboardButton(text="🏠 Главное меню", callback_data="home")]])

def org_keyboard(prefix: str) -> InlineKeyboardMarkup:
    rows = []
    row = []
    for name, _, icon in ORGS:
        row.append(InlineKeyboardButton(text=f"{icon} {name}", callback_data=f"{prefix}:{name}"))
        if len(row) == 2:
            rows.append(row); row=[]
    if row: rows.append(row)
    rows.append([InlineKeyboardButton(text="🏠 Главное меню", callback_data="home")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

def menu_keyboard(admin=False):
    rows = [
        [InlineKeyboardButton(text="📝 ПОДАТЬ ЗАЯВЛЕНИЕ", callback_data="apply")],
        [InlineKeyboardButton(text="👤 Профиль", callback_data="profile"), InlineKeyboardButton(text="🏛 Организация", callback_data="organization")],
        [InlineKeyboardButton(text="📋 Критерии", callback_data="criteria"), InlineKeyboardButton(text="📄 Рапорт", callback_data="report")],
        [InlineKeyboardButton(text="🛠 Сервисы", callback_data="services"), InlineKeyboardButton(text="❓ Помощь", callback_data="help")],
        [InlineKeyboardButton(text="⭐ Ваш любимый администратор", callback_data="favorite_admin")],
    ]
    if admin:
        rows.append([InlineKeyboardButton(text="⚙️ АДМИН-ПАНЕЛЬ", callback_data="admin")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

def admin_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 Дашборд", callback_data="admin_dashboard")],
        [InlineKeyboardButton(text="🔔 Уведомления", callback_data="admin_notifications"), InlineKeyboardButton(text="📈 Статистика", callback_data="admin_stats")],
        [InlineKeyboardButton(text="👥 Личный состав", callback_data="admin_people_org")],
        [InlineKeyboardButton(text="🔄 Кадровые действия", callback_data="admin_actions")],
        [InlineKeyboardButton(text="📥 Заявления", callback_data="admin_apps"), InlineKeyboardButton(text="📄 Рапорты", callback_data="admin_reports")],
        [InlineKeyboardButton(text="📌 Критерии", callback_data="admin_criteria"), InlineKeyboardButton(text="📊 Excel", callback_data="admin_excel")],
        [InlineKeyboardButton(text="🏛 Организации", callback_data="admin_orgs"), InlineKeyboardButton(text="🖼 Баннер", callback_data="admin_banner")],
        [InlineKeyboardButton(text="🏠 Главное меню", callback_data="home")],
    ])

def personnel_actions_kb(user: User):
    rows = [
        [InlineKeyboardButton(text="🎖 Изменить звание", callback_data=f"act_rank:{user.id}")],
        [InlineKeyboardButton(text="💼 Изменить должность", callback_data=f"act_position:{user.id}")],
        [InlineKeyboardButton(text="🏛 Перевести", callback_data=f"act_transfer:{user.id}")],
        [InlineKeyboardButton(text="⬆️ Повысить", callback_data=f"act_promote:{user.id}")],
        [InlineKeyboardButton(text="🔽 Понизить", callback_data=f"act_demote:{user.id}")],
        [InlineKeyboardButton(text="⚠️ Взыскание", callback_data=f"act_discipline:{user.id}")],
        [InlineKeyboardButton(text="📦 В архив", callback_data=f"act_archive:{user.id}")],
        [InlineKeyboardButton(text="♻️ Восстановить", callback_data=f"act_restore:{user.id}")],
        [InlineKeyboardButton(text="📄 История", callback_data=f"act_history:{user.id}")],
        [InlineKeyboardButton(text="◀️ К составу", callback_data="admin_people_org")],
    ]
    return InlineKeyboardMarkup(inline_keyboard=rows)

async def edit_page(call: CallbackQuery, text: str, reply_markup=None):
    try:
        if call.message and call.message.photo:
            await call.message.edit_caption(caption=text, reply_markup=reply_markup)
        else:
            await call.message.edit_text(text, reply_markup=reply_markup)
    except Exception:
        if call.message:
            await call.message.answer(text, reply_markup=reply_markup)

async def get_user(session, telegram_id):
    return (await session.execute(select(User).where(User.telegram_id == telegram_id))).scalar_one_or_none()

async def ensure_user(session, tg_id, username, first_name):
    user = await get_user(session, tg_id)
    if not user:
        user = User(telegram_id=tg_id, username=username, first_name=first_name,
                    role="OWNER" if tg_id == OWNER_ID else "PLAYER",
                    active=tg_id == OWNER_ID, archived=False, created_at=utcnow())
        session.add(user)
    else:
        user.username, user.first_name = username, first_name
        if tg_id == OWNER_ID:
            user.role = "OWNER"
    await session.commit()
    await session.refresh(user)
    return user

async def send_home(message: Message):
    async with SessionLocal() as session:
        user = await ensure_user(session, message.from_user.id, message.from_user.username, message.from_user.first_name)
    text = "⚙️ NEMAZING RP\n━━━━━━━━━━━━━━━━━━\nКадровая система RP-проекта\n\nВыберите раздел."
    banner = find_banner()
    if banner:
        with suppress(Exception):
            await message.answer_photo(FSInputFile(banner), caption=text, reply_markup=menu_keyboard(is_admin(user.telegram_id)))
            return
    await message.answer(text, reply_markup=menu_keyboard(is_admin(user.telegram_id)))

@dp.message(CommandStart())
async def start(message: Message):
    """Открывает главное меню без обязательной подписки на канал."""
    if not message.from_user:
        return
    await send_home(message)


@dp.callback_query(F.data == "check_sub")
async def check_sub(call):
    # Оставлено для старых сообщений/кнопок, но подписка больше не требуется.
    await call.answer("Доступ открыт")
    await edit_page(call, "Доступ открыт.", menu_keyboard(is_admin(call.from_user.id)))


@dp.callback_query(F.data == "home")
async def home(call, state: FSMContext):
    await state.clear()
    with suppress(Exception): await call.message.delete()
    await send_home(call.message)
    await call.answer()

# -------------------- APPLICATION --------------------

@dp.callback_query(F.data == "apply")
async def apply_start(call, state):
    await state.clear(); await state.set_state(ApplyStates.organization)
    await edit_page(call, "📝 ПОДАЧА ЗАЯВЛЕНИЯ\n\nВыберите организацию:", org_keyboard("applyorg"))
    await call.answer()

@dp.callback_query(F.data.startswith("applyorg:"), ApplyStates.organization)
async def apply_org(call, state):
    org = call.data.split(":",1)[1]
    await state.update_data(organization=org); await state.set_state(ApplyStates.nickname)
    await edit_page(call, "Введите RP-никнейм:"); await call.answer()

@dp.message(ApplyStates.nickname)
async def apply_nick(message, state):
    await state.update_data(nickname=(message.text or "").strip()); await state.set_state(ApplyStates.rank)
    await message.answer("Введите звание в игре:")

@dp.message(ApplyStates.rank)
async def apply_rank(message, state):
    await state.update_data(rank=(message.text or "").strip()); await state.set_state(ApplyStates.position)
    await message.answer("Введите должность:")

@dp.message(ApplyStates.position)
async def apply_position(message, state):
    await state.update_data(position=(message.text or "").strip()); await state.set_state(ApplyStates.department)
    await message.answer("Введите отдел/подразделение:")

@dp.message(ApplyStates.department)
async def apply_department(message, state):
    await state.update_data(department=(message.text or "").strip()); await state.set_state(ApplyStates.callsign)
    data = await state.get_data()
    if data.get("organization") == "ФСБ":
        await message.answer("Введите позывной ФСБ:")
    else:
        await state.update_data(callsign="")
        await state.set_state(ApplyStates.confirm)
        await application_preview(message, state)

async def application_preview(message, state):
    d=await state.get_data()
    text=(f"📋 ПРОВЕРКА ЗАЯВЛЕНИЯ\n\n"
          f"Организация: {d['organization']}\nНик: {d['nickname']}\n"
          f"Звание: {d['rank']}\nДолжность: {d['position']}\nОтдел: {d['department']}")
    if d.get("organization")=="ФСБ": text += f"\nПозывной: {d.get('callsign') or '—'}"
    await message.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Отправить", callback_data="apply_confirm")],
        [InlineKeyboardButton(text="❌ Отмена", callback_data="home")]
    ]))

@dp.message(ApplyStates.callsign)
async def apply_callsign(message, state):
    await state.update_data(callsign=(message.text or "").strip())
    await state.set_state(ApplyStates.confirm)
    await application_preview(message, state)

@dp.callback_query(F.data == "apply_confirm")
async def apply_confirm(call, state):
    d=await state.get_data()
    if not d: await call.answer("Заявление устарело", show_alert=True); return
    async with SessionLocal() as session:
        user=await ensure_user(session, call.from_user.id, call.from_user.username, call.from_user.first_name)
        pending=(await session.execute(select(Application).where(Application.user_id==user.id, Application.status=="PENDING"))).scalar_one_or_none()
        if pending:
            await call.answer("У вас уже есть заявление на рассмотрении", show_alert=True); await state.clear(); return
        app=Application(user_id=user.id, organization=d["organization"], nickname=d["nickname"], rank=d["rank"],
                        position=d["position"], department=d["department"], callsign=d.get("callsign") or None,
                        status="PENDING", created_at=utcnow())
        session.add(app); await session.commit(); await session.refresh(app)
        app_id=app.id
    await create_notification("APPLICATION", "Новая анкета", f"Заявление #{app_id} • {d['organization']} • {d['nickname']}", user.id, app_id)
    await state.clear(); await edit_page(call, f"✅ Заявление #{app_id} отправлено владельцу.")
    if bot and OWNER_ID:
        await bot.send_message(OWNER_ID,
            f"📥 НОВОЕ ЗАЯВЛЕНИЕ #{app_id}\n\nПользователь: {call.from_user.id}\n"
            f"Организация: {d['organization']}\nНик: {d['nickname']}\nЗвание: {d['rank']}\n"
            f"Должность: {d['position']}\nОтдел: {d['department']}\nПозывной: {d.get('callsign') or '—'}",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Одобрить", callback_data=f"approve_app:{app_id}"),
                 InlineKeyboardButton(text="❌ Отклонить", callback_data=f"reject_app:{app_id}")]
            ]))
    await call.answer()

@dp.callback_query(F.data.startswith("approve_app:"))
async def approve_app(call):
    if not is_admin(call.from_user.id): await call.answer("Нет доступа", show_alert=True); return
    app_id=int(call.data.split(":")[1])
    async with SessionLocal() as session:
        app=await session.get(Application, app_id)
        if not app or app.status!="PENDING": await call.answer("Заявление уже обработано", show_alert=True); return
        user=await session.get(User, app.user_id)
        org=(await session.execute(select(Organization).where(Organization.name==app.organization))).scalar_one_or_none()
        if not user: await call.answer("Пользователь не найден", show_alert=True); return
        user.nickname, user.rank, user.position, user.department = app.nickname, app.rank, app.position, app.department
        user.organization, user.callsign = app.organization, app.callsign
        user.active, user.archived, user.joined_at, user.archived_at = True, False, utcnow(), None
        app.status="APPROVED"; app.reviewed_by=call.from_user.id; app.reviewed_at=utcnow()
        session.add(PersonnelHistory(user_id=user.id, action="ПРИЁМ", to_organization=app.organization,
                                     to_rank=app.rank, to_position=app.position, note=f"Заявление #{app_id}",
                                     actor_id=call.from_user.id, created_at=utcnow()))
        chat_id=org.chat_id if org else None
        await session.commit()
    with suppress(Exception): await call.message.edit_reply_markup(reply_markup=None)
    invite=None
    if bot and chat_id:
        try:
            invite=await bot.create_chat_invite_link(chat_id=chat_id, name=f"NEMAZING #{app_id}",
                expire_date=datetime.utcnow()+timedelta(hours=24), member_limit=1, creates_join_request=False)
        except Exception as e: logger.warning("invite error: %s", e)
    if bot:
        text=f"✅ Заявление #{app_id} одобрено.\n\nВас приняли в {app.organization}."
        if invite: text += f"\n\n🔐 Персональная ссылка в закрытый чат (1 вход, 24 часа):\n{invite.invite_link}"
        elif app.organization: text += "\n\n⚠️ Бот не смог создать персональную ссылку. Проверьте права администратора бота в чате."
        with suppress(Exception): await bot.send_message(user.telegram_id,text)
    await call.answer("Одобрено")

@dp.callback_query(F.data.startswith("reject_app:"))
async def reject_app(call):
    if not is_admin(call.from_user.id): await call.answer("Нет доступа", show_alert=True); return
    app_id=int(call.data.split(":")[1])
    async with SessionLocal() as session:
        app=await session.get(Application, app_id)
        if not app or app.status!="PENDING": await call.answer("Заявление уже обработано", show_alert=True); return
        app.status="REJECTED"; app.reviewed_by=call.from_user.id; app.reviewed_at=utcnow()
        user=await session.get(User,app.user_id); await session.commit()
    with suppress(Exception): await call.message.edit_reply_markup(reply_markup=None)
    if user and bot: 
        with suppress(Exception): await bot.send_message(user.telegram_id,f"❌ Заявление #{app_id} отклонено владельцем.")
    await call.answer("Отклонено")

# -------------------- PROFILE --------------------

@dp.callback_query(F.data=="profile")
async def profile(call):
    async with SessionLocal() as session:
        u=await ensure_user(session,call.from_user.id,call.from_user.username,call.from_user.first_name)
    text=f"👤 ПРОФИЛЬ\n\nID: {u.telegram_id}\nНик: {u.nickname or '—'}\n"
    if u.organization: text+=f"Организация: {u.organization}\nЗвание: {u.rank or '—'}\nДолжность: {u.position or '—'}\nОтдел: {u.department or '—'}\n"
    if u.organization=="ФСБ": text+=f"Позывной: {u.callsign or '—'}\n"
    text+=f"Статус: {'🟢 АКТИВЕН' if u.active and not u.archived else '📦 АРХИВ'}"
    await edit_page(call,text,page_kb([])); await call.answer()

@dp.callback_query(F.data=="organization")
async def organization(call):
    async with SessionLocal() as session:
        u=await get_user(session,call.from_user.id)
        if not u or not u.organization:
            text="🏛 Вы пока не состоите в организации."
        else:
            people=(await session.execute(select(User).where(User.organization==u.organization,User.active==True,User.archived==False).order_by(User.rank,User.nickname))).scalars().all()
            lines=[f"{org_icon(u.organization)} {u.organization}","",f"Сотрудников: {len(people)}",""]
            lines += [f"• {p.nickname or p.first_name or 'Без имени'} — {p.rank or 'без звания'}" for p in people[:100]]
            text="\n".join(lines)
    await edit_page(call,text,page_kb([])); await call.answer()

# -------------------- CRITERIA FILE MANAGER --------------------

@dp.callback_query(F.data=="criteria")
async def criteria_root(call):
    await edit_page(call,"📋 КРИТЕРИИ\n\nВыберите организацию:",org_keyboard("critorg")); await call.answer()

@dp.callback_query(F.data.startswith("critorg:"))
async def criteria_org(call):
    org=call.data.split(":",1)[1]
    async with SessionLocal() as session:
        items=(await session.execute(select(Criterion).where(Criterion.organization==org).order_by(Criterion.id))).scalars().all()
    rows=[]
    for c in items:
        fr=c.from_rank or c.rank or "—"; tr=c.to_rank or "—"
        rows.append([InlineKeyboardButton(text=f"📂 {fr} → {tr}",callback_data=f"crit:{c.id}")])
    if not rows: rows=[[InlineKeyboardButton(text="Пока нет критериев",callback_data="noop")]]
    rows.append([InlineKeyboardButton(text="◀️ Организации",callback_data="criteria")])
    await edit_page(call,f"{org_icon(org)} {org}\n\nВыберите переход:",InlineKeyboardMarkup(inline_keyboard=rows)); await call.answer()

@dp.callback_query(F.data.startswith("crit:"))
async def criteria_detail(call):
    cid=int(call.data.split(":")[1])
    async with SessionLocal() as session: c=await session.get(Criterion,cid)
    if not c: await call.answer("Критерий не найден",show_alert=True); return
    fr=c.from_rank or c.rank or "—"; tr=c.to_rank or "—"
    text=f"📂 {fr} → {tr}\n\n📌 Критерий:\n{c.title}\n\n📎 Необходимо:\n"
    text+=f"• Отслужить: {c.min_days} дней\n" if c.min_days else ""
    text+=f"• {c.required_fixations} фиксации\n" if c.required_fixations else ""
    text+=f"• {c.required_tasks} выполнения задач\n" if c.required_tasks else ""
    text+="• Отсутствие взысканий\n" if c.no_discipline else ""
    text+=f"\n📝 Подробности:\n{c.description}"
    kb=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 ПОДАТЬ РАПОРТ",callback_data=f"report_criterion:{c.id}")],
        [InlineKeyboardButton(text="◀️ Критерии",callback_data=f"critorg:{c.organization}")],
    ])
    await edit_page(call,text,kb); await call.answer()

@dp.callback_query(F.data=="noop")
async def noop(call): await call.answer("Здесь пока пусто")

# -------------------- REPORTS + REAL ATTACHMENTS --------------------

def attachment_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Добавить ещё",callback_data="attach_more"),
         InlineKeyboardButton(text="✅ Завершить",callback_data="attach_finish")],
        [InlineKeyboardButton(text="❌ Отмена",callback_data="home")]
    ])

async def report_begin(call,state,criterion_id=None):
    async with SessionLocal() as session:
        u=await get_user(session,call.from_user.id)
        if not u or not u.active or u.archived or not u.organization:
            await call.answer("Рапорт доступен только действующему составу",show_alert=True); return
        if criterion_id:
            c=await session.get(Criterion,criterion_id)
            if not c or c.organization!=u.organization:
                await call.answer("Этот критерий не относится к вашей организации",show_alert=True); return
            await state.update_data(organization=c.organization,criterion_id=c.id,criterion_from_rank=c.from_rank or c.rank,criterion_to_rank=c.to_rank or "")
            await state.set_state(ReportStates.from_whom)
            await edit_page(call,f"📄 РАПОРТ\n\n{c.from_rank or c.rank} → {c.to_rank or '—'}\n\nВведите, от кого рапорт:")
        else:
            await state.set_state(ReportStates.organization)
            await edit_page(call,"📄 РАПОРТ\n\nВыберите организацию:",org_keyboard("reportorg"))
    await call.answer()

@dp.callback_query(F.data=="report")
async def report_start(call,state): await report_begin(call,state)

@dp.callback_query(F.data.startswith("report_criterion:"))
async def report_from_criterion(call,state):
    await report_begin(call,state,int(call.data.split(":")[1]))

@dp.callback_query(F.data.startswith("reportorg:"),ReportStates.organization)
async def report_org(call,state):
    org=call.data.split(":",1)[1]
    async with SessionLocal() as session:
        u=await get_user(session,call.from_user.id)
        if not u or u.organization!=org:
            await call.answer("Можно подавать рапорт только по своей организации",show_alert=True); return
        items=(await session.execute(select(Criterion).where(Criterion.organization==org).order_by(Criterion.id))).scalars().all()
    rows=[[InlineKeyboardButton(text=f"📂 {c.from_rank or c.rank} → {c.to_rank or '—'}",callback_data=f"reportcrit:{c.id}")] for c in items]
    rows.append([InlineKeyboardButton(text="❌ Отмена",callback_data="home")])
    await state.set_state(ReportStates.criterion)
    await edit_page(call,f"{org_icon(org)} {org}\n\nВыберите критерий:",InlineKeyboardMarkup(inline_keyboard=rows)); await call.answer()

@dp.callback_query(F.data.startswith("reportcrit:"),ReportStates.criterion)
async def report_criterion(call,state):
    cid=int(call.data.split(":")[1])
    async with SessionLocal() as session:
        c=await session.get(Criterion,cid); u=await get_user(session,call.from_user.id)
    if not c or not u or c.organization!=u.organization:
        await call.answer("Критерий недоступен",show_alert=True); return
    expected = c.from_rank or c.rank
    if expected and u.rank and u.rank != expected:
        await call.answer(f"Критерий рассчитан на звание: {expected}",show_alert=True); return
    await state.update_data(organization=c.organization,criterion_id=c.id,criterion_from_rank=c.from_rank or c.rank,criterion_to_rank=c.to_rank or "")
    await state.set_state(ReportStates.from_whom)
    await edit_page(call,f"📂 {c.from_rank or c.rank} → {c.to_rank or '—'}\n\nВведите, от кого рапорт:")
    await call.answer()

@dp.message(ReportStates.from_whom)
async def report_from(message,state):
    d=await state.get_data()
    await state.update_data(from_whom=(message.text or "").strip(),to_whom=f"Руководителю {d.get('organization','')}")
    await state.set_state(ReportStates.task_points)
    await message.answer("Укажите выполненные задачи / баллы:")

@dp.message(ReportStates.task_points)
async def report_tasks(message,state):
    await state.update_data(task_points=(message.text or "").strip(),attachments=[])
    await state.set_state(ReportStates.attachments)
    await message.answer("📎 Фиксации\n\nОтправьте фото, видео, документ или ссылку.\nКогда всё добавите — нажмите «Завершить».",reply_markup=attachment_keyboard())

def extract_url(message):
    txt=(message.text or message.caption or "").strip()
    if txt.startswith(("http://","https://","t.me/")): return txt
    return None

@dp.message(ReportStates.attachments)
async def report_attachment(message,state):
    d=await state.get_data(); arr=d.get("attachments",[])
    item=None
    if message.photo:
        item={"kind":"photo","file_id":message.photo[-1].file_id,"caption":message.caption or ""}
    elif message.video:
        item={"kind":"video","file_id":message.video.file_id,"caption":message.caption or ""}
    elif message.document:
        item={"kind":"document","file_id":message.document.file_id,"caption":message.caption or ""}
    else:
        url=extract_url(message)
        if url: item={"kind":"link","url":url,"caption":message.text or url}
    if not item:
        await message.answer("Поддерживается: 📸 фото, 🎥 видео, 📄 документ, 🔗 ссылка.",reply_markup=attachment_keyboard()); return
    arr.append(item); await state.update_data(attachments=arr)
    counts={}
    for x in arr: counts[x["kind"]]=counts.get(x["kind"],0)+1
    text=f"📎 Фиксации: {len(arr)}\n" + "\n".join([f"📸 Фото — {counts.get('photo',0)}",f"🎥 Видео — {counts.get('video',0)}",f"📄 Документы — {counts.get('document',0)}",f"🔗 Ссылки — {counts.get('link',0)}"])
    await message.answer(text,reply_markup=attachment_keyboard())

@dp.callback_query(F.data=="attach_more",ReportStates.attachments)
async def attach_more(call): await call.answer("Отправляйте следующую фиксацию")

@dp.callback_query(F.data=="attach_finish",ReportStates.attachments)
async def attach_finish(call,state):
    d=await state.get_data()
    await state.set_state(ReportStates.signature)
    await edit_page(call,f"📎 Фиксаций собрано: {len(d.get('attachments',[]))}\n\nВведите подпись:")
    await call.answer()

@dp.message(ReportStates.signature)
async def report_signature(message,state):
    await state.update_data(signature=(message.text or "").strip())
    d=await state.get_data(); await state.set_state(ReportStates.confirm)
    ctext=f"{d.get('criterion_from_rank','—')} → {d.get('criterion_to_rank','—')}"
    await message.answer(
        f"📋 РАПОРТ №...\n\nОт: {d.get('from_whom')}\nОрганизация: {d.get('organization')}\n"
        f"Текущее звание: {d.get('criterion_from_rank')}\nПовышение: {d.get('criterion_to_rank')}\n"
        f"Критерий: {ctext}\nВыполненные задачи: {d.get('task_points')}\n"
        f"Фиксации: {len(d.get('attachments',[]))}\nПодпись: {d.get('signature')}\n"
        f"Дата: {datetime.utcnow().strftime('%d.%m.%Y')}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Отправить",callback_data="report_confirm")],
            [InlineKeyboardButton(text="❌ Отмена",callback_data="home")]
        ]))

@dp.callback_query(F.data=="report_confirm")
async def report_confirm(call,state):
    d=await state.get_data()
    async with SessionLocal() as session:
        u=await get_user(session,call.from_user.id)
        if not u or not u.active or u.archived: await call.answer("Пользователь не активен",show_alert=True); return
        last_num=await session.scalar(select(func.max(Report.org_number)).where(Report.source_organization==d.get("organization")))
        org_number=(last_num or 0)+1
        r=Report(user_id=u.id,source_organization=d.get("organization"),org_number=org_number,criterion_id=d.get("criterion_id"),
                 criterion_from_rank=d.get("criterion_from_rank"),criterion_to_rank=d.get("criterion_to_rank"),
                 to_whom=d.get("to_whom"),from_whom=d.get("from_whom"),task_points=d.get("task_points"),
                 evidence="",signature=d.get("signature"),status="PENDING",created_at=utcnow())
        session.add(r); await session.commit(); await session.refresh(r)
        for a in d.get("attachments",[]):
            session.add(ReportAttachment(report_id=r.id,**a,created_at=utcnow()))
        await session.commit(); rid=r.id
    await create_notification("REPORT", "Новый рапорт", f"Рапорт #{rid} • {d.get('organization')} • {d.get('criterion_from_rank')} → {d.get('criterion_to_rank')}", u.id, rid)
    await state.clear(); await edit_page(call,f"✅ Рапорт #{rid} отправлен владельцу на проверку.")
    if bot and OWNER_ID:
        txt=(f"📄 НОВЫЙ РАПОРТ #{d.get("organization")} №{org_number} (ID {rid})\n\n{d.get('organization')} | {d.get('criterion_from_rank')} → {d.get('criterion_to_rank')}\n"
             f"От: {d.get('from_whom')}\nТекущее звание: {d.get('criterion_from_rank')}\nПовышение: {d.get('criterion_to_rank')}\n"
             f"Критерий: {d.get('criterion_from_rank')} → {d.get('criterion_to_rank')}\nВыполненные задачи: {d.get('task_points')}\n"
             f"Фиксации: {len(d.get('attachments',[]))}\nПодпись: {d.get('signature')}\nДата: {datetime.utcnow().strftime('%d.%m.%Y')}\nTelegram ID: {call.from_user.id}")
        await bot.send_message(OWNER_ID,txt,reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Принять",callback_data=f"approve_report:{rid}"),
             InlineKeyboardButton(text="❌ Отклонить",callback_data=f"reject_report:{rid}")]
        ]))
        await send_report_attachments(OWNER_ID,d.get("attachments",[]))
    await call.answer()

async def send_report_attachments(chat_id, items):
    if not bot: return
    for a in items:
        try:
            if a["kind"]=="photo": await bot.send_photo(chat_id,a["file_id"],caption=a.get("caption") or None)
            elif a["kind"]=="video": await bot.send_video(chat_id,a["file_id"],caption=a.get("caption") or None)
            elif a["kind"]=="document": await bot.send_document(chat_id,a["file_id"],caption=a.get("caption") or None)
            elif a["kind"]=="link": await bot.send_message(chat_id,f"🔗 {a['url']}")
        except Exception as e: logger.warning("attachment send error: %s",e)

async def process_report(call,rid,approved):
    if not is_admin(call.from_user.id): await call.answer("Нет доступа",show_alert=True); return
    async with SessionLocal() as session:
        r=await session.get(Report,rid)
        if not r or r.status!="PENDING": await call.answer("Рапорт уже обработан",show_alert=True); return
        u=await session.get(User,r.user_id)
        r.status="APPROVED" if approved else "REJECTED"; r.reviewed_by=call.from_user.id; r.reviewed_at=utcnow()
        target_org=r.source_organization
        target_chat=None
        if approved and u and r.criterion_to_rank:
            old=u.rank; u.rank=r.criterion_to_rank
            session.add(PersonnelHistory(user_id=u.id,action="ПОВЫШЕНИЕ",from_organization=u.organization,to_organization=u.organization,
                from_rank=old,to_rank=u.rank,from_position=u.position,to_position=u.position,note=f"Рапорт #{rid}",
                actor_id=call.from_user.id,created_at=utcnow()))
        if target_org:
            org=(await session.execute(select(Organization).where(Organization.name==target_org))).scalar_one_or_none()
            target_chat=org.chat_id if org else None
        atts=(await session.execute(select(ReportAttachment).where(ReportAttachment.report_id==rid).order_by(ReportAttachment.id))).scalars().all()
        await session.commit()
    if approved and u and r.criterion_to_rank:
        await create_notification("PROMOTION","Сотрудник получил повышение",f"{u.nickname or u.first_name or u.telegram_id} • {r.source_organization}: {r.criterion_from_rank or '—'} → {r.criterion_to_rank}",u.id,rid)
    if u and bot:
        with suppress(Exception):
            await bot.send_message(u.telegram_id,("✅ Рапорт принят." if approved else "❌ Рапорт отклонён.")+f"\nНомер: #{rid}")
    if approved and bot and target_chat:
        with suppress(Exception):
            await bot.send_message(target_chat,f"📄 РАПОРТ №{r.org_number or rid} — ПРИНЯТ\n\n"
                f"Сотрудник: {u.nickname or u.first_name}\nПереход: {r.criterion_from_rank or '—'} → {r.criterion_to_rank or '—'}\n"
                f"Организация: {r.source_organization}")
        await send_report_attachments(target_chat,[{"kind":a.kind,"file_id":a.file_id,"url":a.url,"caption":a.caption} for a in atts])
    with suppress(Exception): await call.message.edit_reply_markup(reply_markup=None)
    await call.answer("Принято" if approved else "Отклонено")

@dp.callback_query(F.data.startswith("approve_report:"))
async def approve_report(call): await process_report(call,int(call.data.split(":")[1]),True)
@dp.callback_query(F.data.startswith("reject_report:"))
async def reject_report(call): await process_report(call,int(call.data.split(":")[1]),False)

# -------------------- OWNER NOTIFICATIONS / STATISTICS / VOTING --------------------

@dp.callback_query(F.data=="admin_notifications")
async def admin_notifications(call):
    if not is_admin(call.from_user.id): return
    async with SessionLocal() as s:
        items=(await s.execute(select(Notification).order_by(Notification.id.desc()).limit(30))).scalars().all()
    if not items:
        text="🔔 УВЕДОМЛЕНИЯ\n\nПока уведомлений нет."
    else:
        icons={"APPLICATION":"📥","REPORT":"📄","JOIN":"👤","LEAVE":"🚪","PROMOTION":"⬆️","DISCIPLINE":"⚠️"}
        lines=["🔔 УВЕДОМЛЕНИЯ","━━━━━━━━━━━━━━━━━━"]
        for n in items:
            dt=n.created_at.strftime("%d.%m %H:%M") if n.created_at else "—"
            lines.append(f"{icons.get(n.event_type,'🔔')} {n.title} • {dt}\n{n.text}")
        text="\n\n".join(lines)
    await edit_page(call,text,page_kb([])); await call.answer()

@dp.callback_query(F.data=="admin_stats")
async def admin_stats(call):
    if not is_admin(call.from_user.id): return
    month_start=datetime.utcnow().replace(day=1,hour=0,minute=0,second=0,microsecond=0)
    async with SessionLocal() as s:
        apps_total=await s.scalar(select(func.count(Application.id)).where(Application.created_at>=month_start)) or 0
        apps_ok=await s.scalar(select(func.count(Application.id)).where(Application.created_at>=month_start,Application.status=="APPROVED")) or 0
        apps_bad=await s.scalar(select(func.count(Application.id)).where(Application.created_at>=month_start,Application.status=="REJECTED")) or 0
        reps_total=await s.scalar(select(func.count(Report.id)).where(Report.created_at>=month_start)) or 0
        reps_ok=await s.scalar(select(func.count(Report.id)).where(Report.created_at>=month_start,Report.status=="APPROVED")) or 0
        reps_bad=await s.scalar(select(func.count(Report.id)).where(Report.created_at>=month_start,Report.status=="REJECTED")) or 0
        lines=["📊 СТАТИСТИКА","━━━━━━━━━━━━━━━━━━","За текущий месяц","",
               f"📥 Заявлений: {apps_total}",f"   ✅ Принято: {apps_ok}",f"   ❌ Отклонено: {apps_bad}","",
               f"📄 Рапортов: {reps_total}",f"   ✅ Одобрено: {reps_ok}",f"   ❌ Отклонено: {reps_bad}","","🏆 Больше всего повышений:"]
        promo_counts=[]
        for n,_,icon in ORGS:
            c=await s.scalar(select(func.count(PersonnelHistory.id)).where(PersonnelHistory.action.in_(["ПОВЫШЕНИЕ","РАПОРТ: ПОВЫШЕНИЕ"]),PersonnelHistory.to_organization==n,PersonnelHistory.created_at>=month_start)) or 0
            promo_counts.append((c,n,icon))
        promo_counts.sort(reverse=True)
        for i,(c,n,icon) in enumerate(promo_counts[:6],1): lines.append(f"{i}. {icon} {n} — {c}")
    await edit_page(call,"\n".join(lines),page_kb([])); await call.answer()

@dp.callback_query(F.data=="favorite_admin")
async def favorite_admin(call):
    rows=[[InlineKeyboardButton(text=f"👑 {name}",callback_data=f"vote_admin:{i}")] for i,name in enumerate(ADMIN_VOTE_NAMES)]
    rows.append([InlineKeyboardButton(text="📊 Результаты голосования",callback_data="vote_results")])
    rows.append([InlineKeyboardButton(text="🏠 Главное меню",callback_data="home")])
    await edit_page(call,"👑 ВАШ ЛЮБИМЫЙ АДМИНИСТРАТОР\n━━━━━━━━━━━━━━━━━━\nВыберите администратора и отдайте свой голос.\n\nОдин пользователь — один действующий голос.",InlineKeyboardMarkup(inline_keyboard=rows)); await call.answer()

@dp.callback_query(F.data.startswith("vote_admin:"))
async def vote_admin(call):
    idx=int(call.data.split(":")[1])
    if idx<0 or idx>=len(ADMIN_VOTE_NAMES): return
    name=ADMIN_VOTE_NAMES[idx]
    async with SessionLocal() as s:
        old=(await s.execute(select(AdminVote).where(AdminVote.voter_id==call.from_user.id))).scalars().first()
        if old: old.admin_name=name; old.created_at=utcnow()
        else: s.add(AdminVote(voter_id=call.from_user.id,admin_name=name,created_at=utcnow()))
        await s.commit()
    await call.answer(f"Голос отдан за {name}")
    await favorite_admin(call)

@dp.callback_query(F.data=="vote_results")
async def vote_results(call):
    async with SessionLocal() as s:
        rows=(await s.execute(select(AdminVote.admin_name,func.count(AdminVote.id)).group_by(AdminVote.admin_name))).all()
    counts={name:0 for name in ADMIN_VOTE_NAMES}
    for name,count in rows: counts[name]=count
    ordered=sorted(counts.items(),key=lambda x:(-x[1],x[0]))
    lines=["📊 ГОЛОСОВАНИЕ","━━━━━━━━━━━━━━━━━━"]+[f"{i}. 👑 {name} — {count}" for i,(name,count) in enumerate(ordered,1)]
    await edit_page(call,"\n".join(lines),InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Голосование",callback_data="favorite_admin")],[InlineKeyboardButton(text="🏠 Главное меню",callback_data="home")]])); await call.answer()

# -------------------- ADMIN DASHBOARD --------------------

@dp.callback_query(F.data=="admin")
async def admin_panel(call):
    if not is_admin(call.from_user.id): await call.answer("Нет доступа",show_alert=True); return
    await edit_page(call,"⚙️ NEMAZING RP\n\nАдминистративный центр",admin_keyboard()); await call.answer()

@dp.callback_query(F.data=="admin_dashboard")
async def admin_dashboard(call):
    if not is_admin(call.from_user.id): return
    async with SessionLocal() as s:
        total=await s.scalar(select(func.count(User.id)))
        active=await s.scalar(select(func.count(User.id)).where(User.active==True,User.archived==False))
        archive=await s.scalar(select(func.count(User.id)).where(User.archived==True))
        apps=await s.scalar(select(func.count(Application.id)).where(Application.status=="PENDING"))
        reports=await s.scalar(select(func.count(Report.id)).where(Report.status=="PENDING"))
        criteria_count=await s.scalar(select(func.count(Criterion.id)))
        counts={}
        for n,_,_ in ORGS:
            counts[n]=await s.scalar(select(func.count(User.id)).where(User.organization==n,User.active==True,User.archived==False))
    lines=[f"⚙️ NEMAZING RP\n",f"👥 Сотрудников: {total}",f"🟢 Активных: {active}",f"📦 В архиве: {archive}",
           "",f"📥 Заявлений: {apps}",f"📄 Рапортов: {reports}",f"📋 Критериев: {criteria_count}",""]
    lines += [f"{org_icon(n)} {n} — {counts[n]}" for n,_,_ in ORGS]
    await edit_page(call,"\n".join(lines),page_kb([])); await call.answer()

@dp.callback_query(F.data=="admin_people_org")
async def admin_people_org(call):
    if not is_admin(call.from_user.id): return
    await edit_page(call,"👥 ЛИЧНЫЙ СОСТАВ\n\nВыберите организацию:",org_keyboard("peopleorg")); await call.answer()

@dp.callback_query(F.data.startswith("peopleorg:"))
async def admin_people_list(call):
    if not is_admin(call.from_user.id): return
    org=call.data.split(":",1)[1]
    async with SessionLocal() as s:
        users=(await s.execute(select(User).where(User.organization==org,User.active==True,User.archived==False).order_by(User.nickname))).scalars().all()
    rows=[]
    for u in users:
        label=f"{u.nickname or u.first_name or 'Без имени'} — {u.rank or 'без звания'}"
        rows.append([InlineKeyboardButton(text=label[:60],callback_data=f"person:{u.id}")])
    rows.append([InlineKeyboardButton(text="📦 Архив",callback_data=f"archiveorg:{org}")])
    rows.append([InlineKeyboardButton(text="◀️ Организации",callback_data="admin_people_org")])
    await edit_page(call,f"{org_icon(org)} {org}\n\nАктивных: {len(users)}",InlineKeyboardMarkup(inline_keyboard=rows)); await call.answer()

@dp.callback_query(F.data.startswith("archiveorg:"))
async def admin_archive_list(call):
    if not is_admin(call.from_user.id): return
    org=call.data.split(":",1)[1]
    async with SessionLocal() as s:
        users=(await s.execute(select(User).where(User.archived==True,User.archived_organization==org).order_by(User.nickname))).scalars().all()
    rows=[[InlineKeyboardButton(text=f"📦 {u.nickname or u.first_name or 'Без имени'} — {u.archived_rank or '—'}",callback_data=f"person:{u.id}")] for u in users]
    rows.append([InlineKeyboardButton(text="◀️ Организация",callback_data=f"peopleorg:{org}")])
    await edit_page(call,f"📦 АРХИВ — {org}\n\nСотрудников: {len(users)}",InlineKeyboardMarkup(inline_keyboard=rows)); await call.answer()

@dp.callback_query(F.data.startswith("person:"))
async def person_card(call):
    if not is_admin(call.from_user.id): return
    uid=int(call.data.split(":")[1])
    async with SessionLocal() as s: u=await s.get(User,uid)
    if not u: await call.answer("Не найден",show_alert=True); return
    org=u.organization or u.archived_organization or "—"
    rank=u.rank or u.archived_rank or "—"
    pos=u.position or u.archived_position or "—"
    text=f"👤 {u.nickname or u.first_name or 'Без имени'}\n\nTelegram ID: {u.telegram_id}\nUsername: @{u.username or '—'}\nОрганизация: {org}\nЗвание: {rank}\nДолжность: {pos}\nОтдел: {u.department or u.archived_department or '—'}"
    if org=="ФСБ": text+=f"\nПозывной: {u.callsign or u.archived_callsign or '—'}"
    text+=f"\nСтатус: {'🟢 АКТИВЕН' if u.active and not u.archived else '📦 АРХИВ'}"
    await edit_page(call,text,personnel_actions_kb(u)); await call.answer()

@dp.callback_query(F.data=="admin_actions")
async def admin_actions(call):
    if not is_admin(call.from_user.id): return
    await edit_page(call,"🔄 КАДРОВЫЕ ДЕЙСТВИЯ\n\nВыберите организацию для работы с личным составом:",org_keyboard("peopleorg")); await call.answer()

async def begin_action(call,state,action,uid):
    async with SessionLocal() as s: u=await s.get(User,uid)
    if not u: await call.answer("Сотрудник не найден",show_alert=True); return
    await state.clear(); await state.update_data(user_id=uid,action=action)
    if action=="transfer":
        await state.set_state(AdminActionStates.transfer_org)
        await edit_page(call,f"🔄 Перевод\n\n{u.nickname or u.first_name}\nВыберите новую организацию:",org_keyboard("transferorg"))
    else:
        await state.set_state(AdminActionStates.value)
        prompt={"rank":"Введите новое звание:","position":"Введите новую должность:","promote":"Введите новое звание для повышения:",
                "demote":"Введите новое звание для понижения:"}.get(action,"Введите значение:")
        await edit_page(call,prompt)
    await call.answer()

@dp.callback_query(F.data.startswith("act_rank:"))
async def act_rank(call,state): await begin_action(call,state,"rank",int(call.data.split(":")[1]))
@dp.callback_query(F.data.startswith("act_position:"))
async def act_position(call,state): await begin_action(call,state,"position",int(call.data.split(":")[1]))
@dp.callback_query(F.data.startswith("act_promote:"))
async def act_promote(call,state): await begin_action(call,state,"promote",int(call.data.split(":")[1]))
@dp.callback_query(F.data.startswith("act_demote:"))
async def act_demote(call,state): await begin_action(call,state,"demote",int(call.data.split(":")[1]))

@dp.callback_query(F.data.startswith("act_transfer:"))
async def act_transfer(call,state): await begin_action(call,state,"transfer",int(call.data.split(":")[1]))

@dp.callback_query(F.data.startswith("act_archive:"))
async def act_archive(call):
    if not is_admin(call.from_user.id): return
    uid=int(call.data.split(":")[1])
    async with SessionLocal() as s:
        u=await s.get(User,uid)
        if not u: await call.answer("Не найден",show_alert=True); return
        u.archived_organization,u.archived_rank,u.archived_position,u.archived_department,u.archived_callsign=u.organization,u.rank,u.position,u.department,u.callsign
        u.active=False; u.archived=True; u.archived_at=utcnow(); u.organization=u.rank=u.position=u.department=u.callsign=None
        s.add(PersonnelHistory(user_id=uid,action="АРХИВ",from_organization=u.archived_organization,from_rank=u.archived_rank,from_position=u.archived_position,actor_id=call.from_user.id,created_at=utcnow()))
        await s.commit()
    await person_card(call); await call.answer("В архиве")

@dp.callback_query(F.data.startswith("act_restore:"))
async def act_restore(call):
    if not is_admin(call.from_user.id): return
    uid=int(call.data.split(":")[1])
    async with SessionLocal() as s:
        u=await s.get(User,uid)
        if not u: await call.answer("Не найден",show_alert=True); return
        u.organization=u.archived_organization; u.rank=u.archived_rank; u.position=u.archived_position; u.department=u.archived_department; u.callsign=u.archived_callsign
        u.active=True; u.archived=False; u.archived_at=None; u.joined_at=utcnow()
        s.add(PersonnelHistory(user_id=uid,action="ВОССТАНОВЛЕНИЕ",to_organization=u.organization,to_rank=u.rank,to_position=u.position,actor_id=call.from_user.id,created_at=utcnow()))
        await s.commit()
    await person_card(call); await call.answer("Восстановлен")

@dp.message(AdminActionStates.value)
async def admin_action_value(message,state):
    d=await state.get_data(); uid=d["user_id"]; action=d["action"]; val=(message.text or "").strip()
    async with SessionLocal() as s:
        u=await s.get(User,uid)
        if not u: await message.answer("Сотрудник не найден"); await state.clear(); return
        old_rank,old_pos=u.rank,u.position
        if action in ("rank","promote","demote"): u.rank=val
        elif action=="position": u.position=val
        hist_action={"rank":"ИЗМЕНЕНИЕ ЗВАНИЯ","promote":"ПОВЫШЕНИЕ","demote":"ПОНИЖЕНИЕ","position":"ИЗМЕНЕНИЕ ДОЛЖНОСТИ"}[action]
        s.add(PersonnelHistory(user_id=uid,action=hist_action,from_organization=u.organization,to_organization=u.organization,
            from_rank=old_rank,to_rank=u.rank,from_position=old_pos,to_position=u.position,actor_id=message.from_user.id,created_at=utcnow()))
        await s.commit()
        name=u.nickname or u.first_name or str(uid)
        org=u.organization
    if action in ("promote",):
        await create_notification("PROMOTION","Сотрудник получил повышение",f"{name} • {org}: {old_rank or '—'} → {val}",uid)
        if bot and OWNER_ID: pass
    await state.clear(); await message.answer("✅ Кадровое действие выполнено.")

@dp.callback_query(F.data.startswith("transferorg:"))
async def transfer_org(call,state):
    d=await state.get_data()
    if d.get("action")!="transfer": return
    new_org=call.data.split(":",1)[1]; uid=d["user_id"]
    async with SessionLocal() as s:
        u=await s.get(User,uid)
        if not u: await call.answer("Не найден",show_alert=True); return
        old_org=u.organization; u.organization=new_org; u.joined_at=utcnow()
        s.add(PersonnelHistory(user_id=uid,action="ПЕРЕВОД",from_organization=old_org,to_organization=new_org,
            from_rank=u.rank,to_rank=u.rank,from_position=u.position,to_position=u.position,actor_id=call.from_user.id,created_at=utcnow()))
        await s.commit()
    await state.clear(); await person_card(call); await call.answer("Перевод выполнен")

@dp.callback_query(F.data.startswith("act_discipline:"))
async def act_discipline(call,state):
    if not is_admin(call.from_user.id): return
    uid=int(call.data.split(":")[1])
    async with SessionLocal() as s: u=await s.get(User,uid)
    if not u: await call.answer("Не найден",show_alert=True); return
    await state.clear(); await state.update_data(user_id=uid)
    await state.set_state(DisciplineStates.value)
    await edit_page(call,f"⚠️ ВЗЫСКАНИЕ\n\nСотрудник: {u.nickname or u.first_name or uid}\n\nВведите вид и причину взыскания.\nНапример: Выговор — нарушение RP-правил.")
    await call.answer()

@dp.message(DisciplineStates.value)
async def discipline_value(message,state):
    d=await state.get_data(); uid=d.get("user_id"); text=(message.text or "").strip()
    if not text: await message.answer("Укажите причину взыскания."); return
    if "—" in text:
        kind,reason=[x.strip() for x in text.split("—",1)]
    elif "-" in text:
        kind,reason=[x.strip() for x in text.split("-",1)]
    else:
        kind,reason="Взыскание",text
    async with SessionLocal() as s:
        u=await s.get(User,uid)
        if not u: await state.clear(); await message.answer("Сотрудник не найден"); return
        s.add(Discipline(user_id=uid,kind=kind,reason=reason,actor_id=message.from_user.id,created_at=utcnow()))
        s.add(PersonnelHistory(user_id=uid,action="ВЗЫСКАНИЕ",from_organization=u.organization,to_organization=u.organization,from_rank=u.rank,to_rank=u.rank,note=f"{kind}: {reason}",actor_id=message.from_user.id,created_at=utcnow()))
        await s.commit()
        name=u.nickname or u.first_name or str(uid)
    await state.clear()
    await create_notification("DISCIPLINE","Сотрудник получил взыскание",f"{name} • {kind}: {reason}",uid)
    if bot:
        with suppress(Exception): await bot.send_message(uid,f"⚠️ Вам вынесено взыскание\n\n{kind}: {reason}")
    await message.answer("✅ Взыскание добавлено.")

@dp.callback_query(F.data.startswith("act_history:"))
async def act_history(call):
    if not is_admin(call.from_user.id): return
    uid=int(call.data.split(":")[1])
    async with SessionLocal() as s:
        u=await s.get(User,uid)
        items=(await s.execute(select(PersonnelHistory).where(PersonnelHistory.user_id==uid).order_by(PersonnelHistory.id.desc()).limit(50))).scalars().all()
    lines=[f"📄 ИСТОРИЯ — {u.nickname or u.first_name if u else uid}",""]
    for h in items:
        date=h.created_at.strftime("%d.%m.%Y %H:%M") if h.created_at else "—"
        lines.append(f"• {date} — {h.action}")
        if h.from_organization or h.to_organization: lines.append(f"  {h.from_organization or '—'} → {h.to_organization or '—'}")
        if h.from_rank or h.to_rank: lines.append(f"  {h.from_rank or '—'} → {h.to_rank or '—'}")
        if h.note: lines.append(f"  {h.note}")
    await edit_page(call,"\n".join(lines),InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Карточка",callback_data=f"person:{uid}")]])); await call.answer()

# -------------------- ADMIN APPLICATIONS / REPORTS / CRITERIA --------------------

@dp.callback_query(F.data=="admin_apps")
async def admin_apps(call):
    if not is_admin(call.from_user.id): return
    async with SessionLocal() as s: items=(await s.execute(select(Application).where(Application.status=="PENDING").order_by(Application.id.desc()).limit(30))).scalars().all()
    text="📥 ЗАЯВЛЕНИЯ\n\n"+("\n".join(f"#{a.id} — {a.organization} — {a.nickname} — {a.rank}" for a in items) if items else "Новых заявлений нет.")
    await edit_page(call,text,admin_keyboard()); await call.answer()

@dp.callback_query(F.data=="admin_reports")
async def admin_reports(call):
    if not is_admin(call.from_user.id): return
    async with SessionLocal() as s: items=(await s.execute(select(Report).where(Report.status=="PENDING").order_by(Report.id.desc()).limit(30))).scalars().all()
    text="📄 РАПОРТЫ\n\n"+("\n".join(f"#{r.id} — {r.source_organization or '—'} — {r.criterion_from_rank or '—'} → {r.criterion_to_rank or '—'}" for r in items) if items else "Новых рапортов нет.")
    await edit_page(call,text,admin_keyboard()); await call.answer()

@dp.callback_query(F.data=="admin_criteria")
async def admin_criteria(call):
    if not is_admin(call.from_user.id): return
    await edit_page(call,"📌 КРИТЕРИИ\n\nВыберите организацию:",org_keyboard("admincritorg")); await call.answer()

@dp.callback_query(F.data.startswith("admincritorg:"))
async def admin_criteria_org(call):
    if not is_admin(call.from_user.id): return
    org=call.data.split(":",1)[1]
    async with SessionLocal() as s: items=(await s.execute(select(Criterion).where(Criterion.organization==org).order_by(Criterion.id))).scalars().all()
    rows=[[InlineKeyboardButton(text=f"📂 {c.from_rank or c.rank} → {c.to_rank or '—'} | {c.title[:25]}",callback_data=f"admincrit:{c.id}")] for c in items]
    rows.append([InlineKeyboardButton(text="➕ Добавить критерий",callback_data=f"addcrit:{org}")])
    rows.append([InlineKeyboardButton(text="◀️ Критерии",callback_data="admin_criteria")])
    await edit_page(call,f"{org_icon(org)} {org}\n\nКритерии: {len(items)}",InlineKeyboardMarkup(inline_keyboard=rows)); await call.answer()

@dp.callback_query(F.data.startswith("addcrit:"))
async def addcrit_start(call,state):
    if not is_admin(call.from_user.id): return
    org=call.data.split(":",1)[1]; await state.clear(); await state.update_data(organization=org)
    await state.set_state(AdminCriterionStates.from_rank); await edit_page(call,f"➕ КРИТЕРИЙ — {org}\n\nВведите исходное звание:"); await call.answer()

@dp.message(AdminCriterionStates.from_rank)
async def ac_from(message,state):
    await state.update_data(from_rank=(message.text or "").strip()); await state.set_state(AdminCriterionStates.to_rank); await message.answer("Введите новое звание:")

@dp.message(AdminCriterionStates.to_rank)
async def ac_to(message,state):
    await state.update_data(to_rank=(message.text or "").strip()); await state.set_state(AdminCriterionStates.title); await message.answer("Название критерия:")

@dp.message(AdminCriterionStates.title)
async def ac_title(message,state):
    await state.update_data(title=(message.text or "").strip()); await state.set_state(AdminCriterionStates.description); await message.answer("Подробные требования:")

@dp.message(AdminCriterionStates.description)
async def ac_desc(message,state):
    await state.update_data(description=(message.text or "").strip()); await state.set_state(AdminCriterionStates.min_days); await message.answer("Минимум дней в звании/организации (0 если не требуется):")

@dp.message(AdminCriterionStates.min_days)
async def ac_days(message,state):
    try: n=max(0,int((message.text or "0").strip()))
    except: n=0
    await state.update_data(min_days=n); await state.set_state(AdminCriterionStates.fixations); await message.answer("Сколько фиксаций необходимо? (0 если нет)")

@dp.message(AdminCriterionStates.fixations)
async def ac_fix(message,state):
    try: n=max(0,int((message.text or "0").strip()))
    except: n=0
    await state.update_data(fixations=n); await state.set_state(AdminCriterionStates.tasks); await message.answer("Сколько выполнений задач необходимо? (0 если нет)")

@dp.message(AdminCriterionStates.tasks)
async def ac_tasks(message,state):
    try: n=max(0,int((message.text or "0").strip()))
    except: n=0
    await state.update_data(tasks=n); await state.set_state(AdminCriterionStates.no_discipline)
    await message.answer("Требовать отсутствие взысканий? Ответьте: да / нет")

@dp.message(AdminCriterionStates.no_discipline)
async def ac_disc(message,state):
    yes=(message.text or "").strip().lower() in {"да","д","yes","1","true"}
    await state.update_data(no_discipline=yes); await state.set_state(AdminCriterionStates.confirm)
    d=await state.get_data()
    text=(f"📋 ПРЕДПРОСМОТР\n\n{d['organization']}\n{d['from_rank']} → {d['to_rank']}\n"
          f"📌 {d['title']}\n📝 {d['description']}\n\nОтслужить: {d['min_days']} дн.\n"
          f"Фиксаций: {d['fixations']}\nЗадач: {d['tasks']}\nБез взысканий: {'да' if yes else 'нет'}")
    await message.answer(text,reply_markup=InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Сохранить",callback_data="savecrit"),InlineKeyboardButton(text="❌ Отмена",callback_data="home")]
    ]))

@dp.callback_query(F.data=="savecrit")
async def savecrit(call,state):
    if not is_admin(call.from_user.id): return
    d=await state.get_data()
    async with SessionLocal() as s:
        s.add(Criterion(organization=d["organization"],rank=d["from_rank"],from_rank=d["from_rank"],to_rank=d["to_rank"],
                        title=d["title"],description=d["description"],min_days=d["min_days"],required_fixations=d["fixations"],
                        required_tasks=d["tasks"],no_discipline=d["no_discipline"],created_at=utcnow()))
        await s.commit()
    await state.clear(); await edit_page(call,"✅ Критерий сохранён.",admin_keyboard()); await call.answer()

@dp.callback_query(F.data.startswith("admincrit:"))
async def admincrit_detail(call):
    if not is_admin(call.from_user.id): return
    cid=int(call.data.split(":")[1])
    async with SessionLocal() as s: c=await s.get(Criterion,cid)
    if not c: await call.answer("Не найден",show_alert=True); return
    text=f"📂 {c.from_rank or c.rank} → {c.to_rank}\n\n📌 {c.title}\n📝 {c.description}\n\nДни: {c.min_days}\nФиксации: {c.required_fixations}\nЗадачи: {c.required_tasks}\nБез взысканий: {'да' if c.no_discipline else 'нет'}"
    await edit_page(call,text,InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Назад",callback_data=f"admincritorg:{c.organization}")]])); await call.answer()

# -------------------- EXCEL / ORGS / BANNER / COMMANDS --------------------

def style_sheet(ws):
    for cell in ws[1]:
        cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="222222")
        cell.alignment=Alignment(horizontal="center",vertical="center")
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        letter=col[0].column_letter
        ws.column_dimensions[letter].width=min(max(max(len(str(c.value or "")) for c in col)+2,12),40)

@dp.callback_query(F.data=="admin_excel")
async def admin_excel(call):
    if not is_admin(call.from_user.id): return
    async with SessionLocal() as s:
        users=(await s.execute(select(User).order_by(User.organization,User.nickname))).scalars().all()
        apps=(await s.execute(select(Application).order_by(Application.id))).scalars().all()
        reports=(await s.execute(select(Report).order_by(Report.id))).scalars().all()
        hist=(await s.execute(select(PersonnelHistory).order_by(PersonnelHistory.id))).scalars().all()
    wb=Workbook()
    ws=wb.active; ws.title="Личный состав"
    ws.append(["ID","Telegram ID","Username","Имя","RP-ник","Организация","Звание","Должность","Отдел","Позывной","Роль","Статус","Дата вступления"])
    for u in users:
        if u.active and not u.archived:
            ws.append([u.id,u.telegram_id,u.username,u.first_name,u.nickname,u.organization,u.rank,u.position,u.department,u.callsign,u.role,"АКТИВЕН",u.joined_at])
    style_sheet(ws)
    wa=wb.create_sheet("Архив")
    wa.append(["ID","Telegram ID","RP-ник","Организация","Звание","Должность","Отдел","Позывной","Дата архива"])
    for u in users:
        if u.archived:
            wa.append([u.id,u.telegram_id,u.nickname,u.archived_organization,u.archived_rank,u.archived_position,u.archived_department,u.archived_callsign,u.archived_at])
    style_sheet(wa)
    wc=wb.create_sheet("Критерии")
    wc.append(["ID","Организация","С звания","На звание","Название","Требования","Дней","Фиксаций","Задач","Без взысканий"])
    async with SessionLocal() as s:
        cs=(await s.execute(select(Criterion).order_by(Criterion.organization,Criterion.id))).scalars().all()
    for c in cs: wc.append([c.id,c.organization,c.from_rank or c.rank,c.to_rank,c.title,c.description,c.min_days,c.required_fixations,c.required_tasks,"ДА" if c.no_discipline else "НЕТ"])
    style_sheet(wc)
    wa2=wb.create_sheet("Заявления")
    wa2.append(["ID","User ID","Организация","Ник","Звание","Должность","Отдел","Позывной","Статус","Создано"])
    for a in apps: wa2.append([a.id,a.user_id,a.organization,a.nickname,a.rank,a.position,a.department,a.callsign,a.status,a.created_at])
    style_sheet(wa2)
    wr=wb.create_sheet("Рапорты")
    wr.append(["ID","№ в организации","User ID","Организация","Критерий","С","На","Кому","От кого","Задачи","Статус","Создано"])
    for r in reports: wr.append([r.id,r.org_number,r.user_id,r.source_organization,r.criterion_id,r.criterion_from_rank,r.criterion_to_rank,r.to_whom,r.from_whom,r.task_points,r.status,r.created_at])
    style_sheet(wr)
    wh=wb.create_sheet("Кадровая история")
    wh.append(["ID","User ID","Действие","Из организации","В организацию","Из звания","В звание","Из должности","В должность","Примечание","Дата"])
    for h in hist: wh.append([h.id,h.user_id,h.action,h.from_organization,h.to_organization,h.from_rank,h.to_rank,h.from_position,h.to_position,h.note,h.created_at])
    style_sheet(wh)
    stream=io.BytesIO(); wb.save(stream); data=stream.getvalue()
    (BASE_DIR/"exports"/"nemazing_personnel_ru.xlsx").write_bytes(data)
    await call.message.answer_document(BufferedInputFile(data,filename="NEMAZING_RP_Личный_состав.xlsx"),
        caption="📊 Полная Excel-выгрузка NEMAZING RP\nЛичный состав всех организаций • архив • критерии • заявления • рапорты • кадровая история")
    await call.answer("Excel готов")

@dp.callback_query(F.data=="admin_orgs")
async def admin_orgs(call):
    if not is_admin(call.from_user.id): return
    async with SessionLocal() as s: orgs=(await s.execute(select(Organization).order_by(Organization.id))).scalars().all()
    text="🏛 ОРГАНИЗАЦИИ\n\n"+"\n".join(f"{org_icon(o.name)} {o.name} — чат: {o.chat_id or 'не настроен'}" for o in orgs)
    await edit_page(call,text,admin_keyboard()); await call.answer()

@dp.callback_query(F.data=="admin_banner")
async def admin_banner(call):
    if not is_admin(call.from_user.id): return
    p=find_banner()
    if not p: await call.answer("Положите PNG/JPG/WEBP в папку баннеры",show_alert=True); return
    await call.message.answer_photo(FSInputFile(p),caption="🖼 Баннер NEMAZING RP"); await call.answer("Готово")

@dp.message(Command("setchat"))
async def setchat(message):
    if not is_admin(message.from_user.id): return
    parts=(message.text or "").split(maxsplit=2)
    if len(parts)<3: await message.answer("Использование: /setchat ДПС -100123456789"); return
    org_name,raw=parts[1],parts[2]
    try: cid=int(raw)
    except: await message.answer("chat_id должен быть числом"); return
    async with SessionLocal() as s:
        o=(await s.execute(select(Organization).where(Organization.name==org_name))).scalar_one_or_none()
        if not o: await message.answer("Организация не найдена"); return
        o.chat_id=cid; await s.commit()
    await message.answer(f"✅ Чат {org_name} сохранён")

@dp.message(Command("setinvite"))
async def setinvite(message):
    if not is_admin(message.from_user.id): return
    parts=(message.text or "").split(maxsplit=2)
    if len(parts)<3: await message.answer("Использование: /setinvite ФСБ https://t.me/+..."); return
    async with SessionLocal() as s:
        o=(await s.execute(select(Organization).where(Organization.name==parts[1]))).scalar_one_or_none()
        if not o: await message.answer("Организация не найдена"); return
        o.invite_link=parts[2]; await s.commit()
    await message.answer("✅ Резервная ссылка сохранена")

@dp.message(Command("setrole"))
async def setrole(message):
    if not is_admin(message.from_user.id): return
    p=(message.text or "").split(maxsplit=2)
    if len(p)<3 or p[2].upper() not in {"OWNER","ADMIN","PLAYER"}:
        await message.answer("Использование: /setrole TELEGRAM_ID OWNER|ADMIN|PLAYER"); return
    async with SessionLocal() as s:
        u=await get_user(s,int(p[1]))
        if not u: await message.answer("Пользователь не найден"); return
        u.role=p[2].upper(); await s.commit()
    await message.answer("✅ Роль изменена")

@dp.message(Command("setrank"))
async def setrank(message):
    if not is_admin(message.from_user.id): return
    p=(message.text or "").split(maxsplit=2)
    if len(p)<3: await message.answer("Использование: /setrank TELEGRAM_ID Новое звание"); return
    async with SessionLocal() as s:
        u=await get_user(s,int(p[1]))
        if not u: await message.answer("Пользователь не найден"); return
        old=u.rank; u.rank=p[2]
        s.add(PersonnelHistory(user_id=u.id,action="ИЗМЕНЕНИЕ ЗВАНИЯ",from_organization=u.organization,to_organization=u.organization,from_rank=old,to_rank=u.rank,actor_id=message.from_user.id,created_at=utcnow()))
        await s.commit()
    await message.answer("✅ Звание изменено и записано в историю")

@dp.message(Command("archive"))
async def archive_cmd(message):
    if not is_admin(message.from_user.id): return
    p=(message.text or "").split(maxsplit=1)
    if len(p)<2: await message.answer("Использование: /archive TELEGRAM_ID"); return
    async with SessionLocal() as s:
        u=await get_user(s,int(p[1]))
        if not u: await message.answer("Пользователь не найден"); return
        u.archived_organization,u.archived_rank,u.archived_position,u.archived_department,u.archived_callsign=u.organization,u.rank,u.position,u.department,u.callsign
        u.active=False; u.archived=True; u.archived_at=utcnow(); u.organization=u.rank=u.position=u.department=u.callsign=None
        s.add(PersonnelHistory(user_id=u.id,action="АРХИВ",from_organization=u.archived_organization,from_rank=u.archived_rank,from_position=u.archived_position,actor_id=message.from_user.id,created_at=utcnow()))
        await s.commit()
    await message.answer("✅ Сотрудник в архиве")

@dp.message(Command("restore"))
async def restore_cmd(message):
    if not is_admin(message.from_user.id): return
    p=(message.text or "").split(maxsplit=1)
    if len(p)<2: await message.answer("Использование: /restore TELEGRAM_ID"); return
    async with SessionLocal() as s:
        u=await get_user(s,int(p[1]))
        if not u: await message.answer("Пользователь не найден"); return
        u.organization=u.archived_organization; u.rank=u.archived_rank; u.position=u.archived_position; u.department=u.archived_department; u.callsign=u.archived_callsign
        u.active=True; u.archived=False; u.archived_at=None; u.joined_at=utcnow()
        s.add(PersonnelHistory(user_id=u.id,action="ВОССТАНОВЛЕНИЕ",to_organization=u.organization,to_rank=u.rank,to_position=u.position,actor_id=message.from_user.id,created_at=utcnow()))
        await s.commit()
    await message.answer("✅ Сотрудник восстановлен")

# -------------------- CHAT MEMBER --------------------

@dp.chat_member()
async def member_update(event: ChatMemberUpdated):
    if not bot: return
    old,new=event.old_chat_member.status,event.new_chat_member.status
    tg_id=event.new_chat_member.user.id
    async with SessionLocal() as s:
        org=(await s.execute(select(Organization).where(Organization.chat_id==event.chat.id))).scalar_one_or_none()
        if not org: return
        user=await get_user(s,tg_id)
        joined = old in {ChatMemberStatus.LEFT,ChatMemberStatus.KICKED} and new in {ChatMemberStatus.MEMBER,ChatMemberStatus.ADMINISTRATOR,ChatMemberStatus.RESTRICTED}
        left = old in {ChatMemberStatus.MEMBER,ChatMemberStatus.ADMINISTRATOR,ChatMemberStatus.RESTRICTED} and new in {ChatMemberStatus.LEFT,ChatMemberStatus.KICKED}
        if joined:
            if user and user.active and not user.archived and user.organization==org.name:
                mention=f'<a href="tg://user?id={tg_id}">{esc(user.nickname or user.first_name or "Сотрудник")}</a>'
                text=f"🆕 НОВЫЙ УЧАСТНИК\n\n{mention}\nОрганизация: {esc(org.name)}\nЗвание: {esc(user.rank)}\nДолжность: {esc(user.position)}\nОтдел: {esc(user.department)}"
                if org.name=="ФСБ": text+=f"\nПозывной: {esc(user.callsign)}"
                await bot.send_message(event.chat.id,text,parse_mode="HTML")
                await create_notification("JOIN", "Участник вошёл в чат", f"{user.nickname or user.first_name or tg_id} вошёл в {org.name}.", user.id)
            else:
                await bot.send_message(event.chat.id,f"⚠️ В чат вошёл пользователь Telegram ID {tg_id}, который не числится действующим сотрудником {org.name}.")
        elif left and user and user.organization==org.name:
            user.archived_organization,user.archived_rank,user.archived_position,user.archived_department,user.archived_callsign=user.organization,user.rank,user.position,user.department,user.callsign
            user.active=False; user.archived=True; user.archived_at=utcnow()
            user.organization=user.rank=user.position=user.department=user.callsign=None
            s.add(PersonnelHistory(user_id=user.id,action="ВЫХОД ИЗ ЧАТА / АРХИВ",from_organization=org.name,from_rank=user.archived_rank,from_position=user.archived_position,actor_id=None,created_at=utcnow()))
            await s.commit()
            await bot.send_message(event.chat.id,f"📦 {user.nickname or user.first_name or tg_id} вышел из {org.name}.\nСотрудник автоматически перемещён в архив.")
            await create_notification("LEAVE", "Сотрудник покинул чат", f"{user.nickname or user.first_name or tg_id} покинул {org.name} и перемещён в архив.", user.id)

# -------------------- SERVICES / HELP --------------------

@dp.callback_query(F.data=="services")
async def services(call):
    await edit_page(call,"🛠 СЕРВИСЫ\n\n🔐 Fenix VPN — @FenixVpNRobot\n⭐ Buy Stars — @Fenix_stars_bot",page_kb([])); await call.answer()

@dp.callback_query(F.data=="help")
async def help_page(call):
    await edit_page(call,"❓ ПОМОЩЬ\n\n1. Откройте бота и выберите нужный раздел.\n2. Подайте заявление.\n3. Дождитесь решения владельца.\n4. После одобрения получите персональную ссылку в чат.\n5. Смотрите критерии как файловую структуру и подавайте рапорт.\n6. В рапорт можно приложить фото, видео, документы и ссылки.",page_kb([])); await call.answer()

# -------------------- DB MIGRATION --------------------

async def add_missing_columns():
    # create_all handles new tables; this section handles old Render databases.
    if DATABASE_URL.startswith("sqlite"):
        async with engine.begin() as conn:
            # SQLite: inspect table columns and add missing ones one-by-one.
            tables = {
                "users": {
                    "callsign":"VARCHAR(255)","archived_organization":"VARCHAR(255)","archived_rank":"VARCHAR(255)",
                    "archived_position":"VARCHAR(255)","archived_department":"VARCHAR(255)","archived_callsign":"VARCHAR(255)"
                },
                "applications":{"callsign":"VARCHAR(255)"},
                "criteria":{"from_rank":"VARCHAR(255)","to_rank":"VARCHAR(255)","min_days":"INTEGER DEFAULT 0",
                    "required_fixations":"INTEGER DEFAULT 0","required_tasks":"INTEGER DEFAULT 0","no_discipline":"BOOLEAN DEFAULT 0"},
                "reports":{"source_organization":"VARCHAR(255)","org_number":"INTEGER","criterion_id":"INTEGER","criterion_from_rank":"VARCHAR(255)","criterion_to_rank":"VARCHAR(255)"}
            }
            for table,cols in tables.items():
                existing={r[1] for r in (await conn.exec_driver_sql(f"PRAGMA table_info({table})")).all()}
                for col,typ in cols.items():
                    if col not in existing:
                        await conn.exec_driver_sql(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")
        return
    statements=[
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS callsign VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS archived_organization VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS archived_rank VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS archived_position VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS archived_department VARCHAR(255)",
        "ALTER TABLE users ADD COLUMN IF NOT EXISTS archived_callsign VARCHAR(255)",
        "ALTER TABLE applications ADD COLUMN IF NOT EXISTS callsign VARCHAR(255)",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS from_rank VARCHAR(255)",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS to_rank VARCHAR(255)",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS min_days INTEGER DEFAULT 0",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS required_fixations INTEGER DEFAULT 0",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS required_tasks INTEGER DEFAULT 0",
        "ALTER TABLE criteria ADD COLUMN IF NOT EXISTS no_discipline BOOLEAN DEFAULT FALSE",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS source_organization VARCHAR(255)",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS org_number INTEGER",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS criterion_id INTEGER",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS criterion_from_rank VARCHAR(255)",
        "ALTER TABLE reports ADD COLUMN IF NOT EXISTS criterion_to_rank VARCHAR(255)",
    ]
    async with engine.begin() as conn:
        for sql in statements:
            await conn.exec_driver_sql(sql)

async def init_db():
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)
    await add_missing_columns()
    async with SessionLocal() as s:
        existing=(await s.execute(select(Organization))).scalars().all()
        names={x.name for x in existing}
        for name,code,_ in ORGS:
            if name not in names: s.add(Organization(name=name,code=code,created_at=utcnow()))
        if OWNER_ID:
            owner=await get_user(s,OWNER_ID)
            if not owner:
                s.add(User(telegram_id=OWNER_ID,role="OWNER",active=True,archived=False,created_at=utcnow()))
            else:
                owner.role="OWNER"; owner.active=True; owner.archived=False
        await s.commit()

async def _run_polling():
    if not bot:
        logger.warning("BOT_TOKEN is not configured")
        return
    while True:
        try:
            await bot.delete_webhook(drop_pending_updates=False)
            await dp.start_polling(bot,allowed_updates=dp.resolve_used_update_types())
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Polling crashed; retry in 5 sec")
            await asyncio.sleep(5)

@asynccontextmanager
async def lifespan(app):
    global _polling_task
    await init_db()
    if bot: _polling_task=asyncio.create_task(_run_polling())
    yield
    if _polling_task:
        _polling_task.cancel()
        with suppress(asyncio.CancelledError): await _polling_task
    if bot:
        with suppress(Exception): await bot.session.close()
    with suppress(Exception): await engine.dispose()

app=FastAPI(title="NEMAZING RP",version="13.0.0",lifespan=lifespan)
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_credentials=True,allow_methods=["*"],allow_headers=["*"])

@app.get("/")
async def root():
    return {"service":"NEMAZING RP","version":"13.0.0","status":"online","health":"/health","api_health":"/api/health"}

@app.get("/health")
async def health(): return {"status":"ok","service":"nemazing-rp"}

@app.get("/api/health")
async def api_health(): return {"status":"ok","service":"nemazing-rp","bot_configured":bool(BOT_TOKEN),"database_configured":bool(DATABASE_URL)}

@app.get("/api/status")
async def api_status():
    return {"service":"NEMAZING RP","status":"online","version":"13.0.0","bot_configured":bool(BOT_TOKEN),
            "polling_running":bool(_polling_task and not _polling_task.done()),"owner_configured":bool(OWNER_ID),
            "subscription_required":False}

if __name__=="__main__":
    uvicorn.run(app,host="0.0.0.0",port=int(os.getenv("PORT","10000")),log_level="info")

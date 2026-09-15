import asyncio, io, logging, os
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from dotenv import load_dotenv
from fastapi import FastAPI
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, ChatMemberUpdated, BufferedInputFile
from aiogram.enums import ChatMemberStatus
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from sqlalchemy import String, Integer, BigInteger, Boolean, DateTime, Text, select
from sqlalchemy.ext.asyncio import create_async_engine, async_sessionmaker, AsyncSession
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from openpyxl import Workbook

load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)s | %(message)s')
BOT_TOKEN=os.getenv('BOT_TOKEN','')
OWNER_ID=int(os.getenv('OWNER_ID','0') or 0)
CHANNEL_USERNAME=os.getenv('CHANNEL_USERNAME','@nemazing')
DATABASE_URL=os.getenv('DATABASE_URL','sqlite+aiosqlite:///./data/nemazing.db')
PORT=int(os.getenv('PORT','10000'))
os.makedirs('data',exist_ok=True)
if DATABASE_URL.startswith('postgres://'): DATABASE_URL=DATABASE_URL.replace('postgres://','postgresql+asyncpg://',1)
elif DATABASE_URL.startswith('postgresql://'): DATABASE_URL=DATABASE_URL.replace('postgresql://','postgresql+asyncpg://',1)
engine=create_async_engine(DATABASE_URL,echo=False)
SessionLocal=async_sessionmaker(engine,expire_on_commit=False)
class Base(DeclarativeBase): pass
class User(Base):
    __tablename__='users'; id:Mapped[int]=mapped_column(Integer,primary_key=True); telegram_id:Mapped[int]=mapped_column(BigInteger,unique=True,index=True); username:Mapped[str|None]=mapped_column(String(255)); nickname:Mapped[str|None]=mapped_column(String(255)); rank:Mapped[str|None]=mapped_column(String(255)); position:Mapped[str|None]=mapped_column(String(255)); department:Mapped[str|None]=mapped_column(String(255)); organization:Mapped[str|None]=mapped_column(String(50)); role:Mapped[str]=mapped_column(String(30),default='PLAYER'); active:Mapped[bool]=mapped_column(Boolean,default=False); archived:Mapped[bool]=mapped_column(Boolean,default=False); joined_at:Mapped[datetime|None]=mapped_column(DateTime); archived_at:Mapped[datetime|None]=mapped_column(DateTime)
class Organization(Base):
    __tablename__='organizations'; id:Mapped[int]=mapped_column(Integer,primary_key=True); code:Mapped[str]=mapped_column(String(50),unique=True); title:Mapped[str]=mapped_column(String(255)); chat_id:Mapped[int|None]=mapped_column(BigInteger); invite_link:Mapped[str|None]=mapped_column(Text); enabled:Mapped[bool]=mapped_column(Boolean,default=True)
class Application(Base):
    __tablename__='applications'; id:Mapped[int]=mapped_column(Integer,primary_key=True); telegram_id:Mapped[int]=mapped_column(BigInteger,index=True); organization:Mapped[str]=mapped_column(String(50)); nickname:Mapped[str]=mapped_column(String(255)); rank:Mapped[str]=mapped_column(String(255)); position:Mapped[str]=mapped_column(String(255)); department:Mapped[str]=mapped_column(String(255)); status:Mapped[str]=mapped_column(String(30),default='PENDING'); created_at:Mapped[datetime]=mapped_column(DateTime,default=lambda:datetime.now(timezone.utc))
class Criterion(Base):
    __tablename__='criteria'; id:Mapped[int]=mapped_column(Integer,primary_key=True); organization:Mapped[str]=mapped_column(String(50)); title:Mapped[str]=mapped_column(String(255)); description:Mapped[str]=mapped_column(Text); enabled:Mapped[bool]=mapped_column(Boolean,default=True)
class Report(Base):
    __tablename__='reports'; id:Mapped[int]=mapped_column(Integer,primary_key=True); telegram_id:Mapped[int]=mapped_column(BigInteger); organization:Mapped[str]=mapped_column(String(50)); target:Mapped[str]=mapped_column(String(255)); author:Mapped[str]=mapped_column(String(255)); task_points:Mapped[str]=mapped_column(String(255)); evidence:Mapped[str]=mapped_column(Text); signature:Mapped[str]=mapped_column(String(255)); status:Mapped[str]=mapped_column(String(30),default='PENDING'); created_at:Mapped[datetime]=mapped_column(DateTime,default=lambda:datetime.now(timezone.utc))
class ApplyStates(StatesGroup): organization=State(); nickname=State(); rank=State(); position=State(); department=State()
class ReportStates(StatesGroup): target=State(); author=State(); points=State(); evidence=State(); signature=State()
dp=Dispatcher(storage=MemoryStorage()); bot=Bot(BOT_TOKEN)
def menu(): return InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='📝 Подать заявку',callback_data='apply')],[InlineKeyboardButton(text='👤 Профиль',callback_data='profile'),InlineKeyboardButton(text='🏢 Моя организация',callback_data='org')],[InlineKeyboardButton(text='📈 Критерии',callback_data='criteria'),InlineKeyboardButton(text='📄 Рапорт',callback_data='report')],[InlineKeyboardButton(text='🛠 Сервисы',callback_data='services'),InlineKeyboardButton(text='ℹ️ Помощь',callback_data='help')]])
def admin_menu(): return InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='📥 Заявки',callback_data='admin_apps'),InlineKeyboardButton(text='📄 Рапорты',callback_data='admin_reports')],[InlineKeyboardButton(text='👥 Персонал',callback_data='admin_people'),InlineKeyboardButton(text='📊 Excel',callback_data='admin_excel')],[InlineKeyboardButton(text='🏢 Организации',callback_data='admin_orgs')]])
async def get_user(db,tg): return (await db.execute(select(User).where(User.telegram_id==tg))).scalar_one_or_none()
async def subscribed(tg):
    try: return (await bot.get_chat_member(CHANNEL_USERNAME,tg)).status in {ChatMemberStatus.MEMBER,ChatMemberStatus.ADMINISTRATOR,ChatMemberStatus.CREATOR}
    except Exception as e: logging.warning('subscription check: %s',e); return False
@dp.message(Command('start'))
async def start(m:Message,state:FSMContext):
    await state.clear()
    if not await subscribed(m.from_user.id):
        kb=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='📢 Подписаться',url=f"https://t.me/{CHANNEL_USERNAME.lstrip('@')}")],[InlineKeyboardButton(text='✅ Я подписался',callback_data='check_sub')]])
        await m.answer('🔐 Для входа в NEMAZING RP сначала подпишись на канал.',reply_markup=kb); return
    async with SessionLocal() as db:
        u=await get_user(db,m.from_user.id)
        if not u: db.add(User(telegram_id=m.from_user.id,username=m.from_user.username)); await db.commit()
    await m.answer('🔥 <b>NEMAZING RP</b>\n\nДобро пожаловать.',reply_markup=menu())
@dp.callback_query(F.data=='check_sub')
async def check(c:CallbackQuery):
    if not await subscribed(c.from_user.id): await c.answer('Подписка не найдена.',show_alert=True); return
    await c.message.edit_text('✅ Подписка подтверждена.',reply_markup=menu()); await c.answer()
@dp.callback_query(F.data=='apply')
async def apply(c:CallbackQuery,state:FSMContext):
    await state.set_state(ApplyStates.organization); await c.message.answer('🏢 Выберите организацию:',reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='ФСБ',callback_data='org:ФСБ'),InlineKeyboardButton(text='ВЧ',callback_data='org:ВЧ')],[InlineKeyboardButton(text='ЕСС',callback_data='org:ЕСС'),InlineKeyboardButton(text='УМВД',callback_data='org:УМВД')]])); await c.answer()
@dp.callback_query(F.data.startswith('org:'))
async def org(c:CallbackQuery,state:FSMContext): await state.update_data(organization=c.data.split(':',1)[1]); await state.set_state(ApplyStates.nickname); await c.message.answer('👤 RP-никнейм:'); await c.answer()
@dp.message(ApplyStates.nickname)
async def an(m,state): await state.update_data(nickname=m.text); await state.set_state(ApplyStates.rank); await m.answer('🎖 Звание:')
@dp.message(ApplyStates.rank)
async def ar(m,state): await state.update_data(rank=m.text); await state.set_state(ApplyStates.position); await m.answer('💼 Должность:')
@dp.message(ApplyStates.position)
async def ap(m,state): await state.update_data(position=m.text); await state.set_state(ApplyStates.department); await m.answer('🏷 Отдел/подразделение:')
@dp.message(ApplyStates.department)
async def ad(m,state):
    d=await state.get_data(); d['department']=m.text
    async with SessionLocal() as db:
        a=Application(telegram_id=m.from_user.id,organization=d['organization'],nickname=d['nickname'],rank=d['rank'],position=d['position'],department=d['department']); db.add(a); await db.commit(); aid=a.id
    await state.clear(); await m.answer(f'✅ Заявка #{aid} отправлена владельцу.',reply_markup=menu())
    if OWNER_ID: await bot.send_message(OWNER_ID,f'🆕 <b>Заявка #{aid}</b>\nОрганизация: {d["organization"]}\nНик: {d["nickname"]}\nЗвание: {d["rank"]}\nДолжность: {d["position"]}\nОтдел: {d["department"]}\nTelegram ID: <code>{m.from_user.id}</code>',reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='✅ Одобрить',callback_data=f'approve:{aid}'),InlineKeyboardButton(text='❌ Отклонить',callback_data=f'reject:{aid}')]]))
@dp.callback_query(F.data.startswith('approve:'))
async def approve(c:CallbackQuery):
    if c.from_user.id!=OWNER_ID: await c.answer('Нет доступа.',show_alert=True); return
    aid=int(c.data.split(':')[1])
    async with SessionLocal() as db:
        a=await db.get(Application,aid)
        if not a or a.status!='PENDING': await c.answer('Уже обработано.',show_alert=True); return
        a.status='APPROVED'; u=await get_user(db,a.telegram_id)
        if not u: u=User(telegram_id=a.telegram_id); db.add(u)
        u.nickname=a.nickname;u.rank=a.rank;u.position=a.position;u.department=a.department;u.organization=a.organization;u.active=True;u.archived=False;u.joined_at=datetime.now(timezone.utc)
        o=(await db.execute(select(Organization).where(Organization.code==a.organization))).scalar_one_or_none(); invite=o.invite_link if o else None
        await db.commit()
        if o and o.chat_id:
            try: invite=(await bot.create_chat_invite_link(chat_id=o.chat_id,member_limit=1,name=f'NEMAZING-{a.telegram_id}')).invite_link
            except Exception: pass
    await bot.send_message(a.telegram_id,'🎉 <b>Заявка одобрена!</b>\n\n'+(f'🔗 Вход: {invite}' if invite else 'Свяжитесь с администрацией для получения доступа.')); await c.message.edit_text(c.message.text+'\n\n✅ ОДОБРЕНО'); await c.answer()
@dp.callback_query(F.data.startswith('reject:'))
async def reject(c:CallbackQuery):
    if c.from_user.id!=OWNER_ID:return
    aid=int(c.data.split(':')[1])
    async with SessionLocal() as db:
        a=await db.get(Application,aid)
        if not a:return
        a.status='REJECTED'; await db.commit(); tg=a.telegram_id
    await bot.send_message(tg,f'❌ Заявка #{aid} отклонена.'); await c.message.edit_text(c.message.text+'\n\n❌ ОТКЛОНЕНО'); await c.answer()
@dp.callback_query(F.data=='profile')
async def profile(c:CallbackQuery):
    async with SessionLocal() as db:u=await get_user(db,c.from_user.id)
    if not u: await c.answer('Сначала /start',show_alert=True);return
    await c.message.answer(f'👤 <b>Профиль</b>\n\nID: <code>{u.telegram_id}</code>\nНик: {u.nickname or "—"}\nОрганизация: {u.organization or "—"}\nЗвание: {u.rank or "—"}\nДолжность: {u.position or "—"}\nОтдел: {u.department or "—"}\nСтатус: {"🟢 Активен" if u.active else "⚪ Неактивен"}'); await c.answer()
@dp.callback_query(F.data=='org')
async def myorg(c:CallbackQuery):
    async with SessionLocal() as db:
        u=await get_user(db,c.from_user.id)
        if not u or not u.organization: await c.answer('Организация не назначена.',show_alert=True);return
        rows=(await db.execute(select(User).where(User.organization==u.organization,User.active==True))).scalars().all()
    await c.message.answer(f'🏢 <b>{u.organization}</b>\nАктивных: {len(rows)}\n\n'+'\n'.join(f'• {x.nickname or x.telegram_id} — {x.rank or "—"}' for x in rows[:50])); await c.answer()
@dp.callback_query(F.data=='criteria')
async def criteria(c:CallbackQuery):
    async with SessionLocal() as db:
        u=await get_user(db,c.from_user.id); rows=(await db.execute(select(Criterion).where(Criterion.organization==u.organization,Criterion.enabled==True))).scalars().all() if u and u.organization else []
    await c.message.answer('📈 <b>Критерии</b>\n\n'+('\n\n'.join(f'<b>{x.title}</b>\n{x.description}' for x in rows) if rows else 'Пока не настроены.')); await c.answer()
@dp.callback_query(F.data=='services')
async def services(c:CallbackQuery): await c.message.answer('🛠 <b>Сервисы</b>',reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='🛡 Fenix VPN',url='https://t.me/FenixVpNRobot')],[InlineKeyboardButton(text='⭐ Buy Stars',url='https://t.me/Fenix_stars_bot')]])); await c.answer()
@dp.callback_query(F.data=='help')
async def help_(c:CallbackQuery): await c.message.answer('ℹ️ /start — меню\n/cancel — отмена\n/admin — админ-панель\n/export — Excel'); await c.answer()
@dp.message(Command('cancel'))
async def cancel(m:Message,state:FSMContext): await state.clear(); await m.answer('❌ Отменено.',reply_markup=menu())
@dp.message(Command('admin'))
async def admin(m:Message):
    if m.from_user.id==OWNER_ID: await m.answer('🛠 <b>Админ-панель</b>',reply_markup=admin_menu())
@dp.callback_query(F.data=='admin_apps')
async def admin_apps(c:CallbackQuery):
    if c.from_user.id!=OWNER_ID:return
    async with SessionLocal() as db:r=(await db.execute(select(Application).where(Application.status=='PENDING').limit(30))).scalars().all()
    await c.message.answer('📥 Нет заявок.' if not r else '\n\n'.join(f'#{x.id} {x.organization} | {x.nickname}\n{x.rank} | {x.position}\nID {x.telegram_id}' for x in r)); await c.answer()
@dp.callback_query(F.data=='admin_reports')
async def admin_reports(c:CallbackQuery):
    if c.from_user.id==OWNER_ID: await c.message.answer('📄 Раздел рапортов готов к расширению.'); await c.answer()
@dp.callback_query(F.data=='admin_people')
async def admin_people(c:CallbackQuery):
    if c.from_user.id!=OWNER_ID:return
    async with SessionLocal() as db:r=(await db.execute(select(User).where(User.active==True))).scalars().all()
    await c.message.answer('👥 <b>Активный персонал</b>\n\n'+('\n'.join(f'{x.organization} | {x.nickname or x.telegram_id} | {x.rank or "—"}' for x in r[:100]) or 'Пусто.')); await c.answer()
@dp.callback_query(F.data=='admin_orgs')
async def admin_orgs(c:CallbackQuery):
    if c.from_user.id!=OWNER_ID:return
    async with SessionLocal() as db:r=(await db.execute(select(Organization))).scalars().all()
    await c.message.answer('🏢 <b>Организации</b>\n\n'+'\n'.join(f'{x.code} — chat: {x.chat_id or "не задан"}' for x in r)); await c.answer()
async def make_excel():
    async with SessionLocal() as db:r=(await db.execute(select(User))).scalars().all()
    wb=Workbook();ws=wb.active;ws.title='Персонал';ws.append(['Организация','Звание','Должность','Отдел','Никнейм','Telegram ID','Username','Статус','Дата'])
    for x in r:ws.append([x.organization,x.rank,x.position,x.department,x.nickname,x.telegram_id,x.username,'АКТИВЕН' if x.active else 'АРХИВ',x.joined_at.isoformat() if x.joined_at else ''])
    b=io.BytesIO();wb.save(b);return b.getvalue()
@dp.callback_query(F.data=='admin_excel')
async def admin_excel(c:CallbackQuery):
    if c.from_user.id!=OWNER_ID:return
    await c.message.answer_document(BufferedInputFile(await make_excel(),filename='nemazing_personnel.xlsx'));await c.answer()
@dp.message(Command('export'))
async def export(m:Message):
    if m.from_user.id==OWNER_ID:await m.answer_document(BufferedInputFile(await make_excel(),filename='nemazing_personnel.xlsx'))
@dp.message(Command('setchat'))
async def setchat(m:Message):
    if m.from_user.id!=OWNER_ID:return
    p=m.text.split(maxsplit=2)
    if len(p)!=3:return await m.answer('Использование: /setchat ФСБ -100123456789')
    async with SessionLocal() as db:
        o=(await db.execute(select(Organization).where(Organization.code==p[1]))).scalar_one_or_none()
        if not o:o=Organization(code=p[1],title=p[1]);db.add(o)
        o.chat_id=int(p[2]);await db.commit()
    await m.answer('✅ Чат сохранён.')
@dp.message(Command('setinvite'))
async def setinvite(m:Message):
    if m.from_user.id!=OWNER_ID:return
    p=m.text.split(maxsplit=2)
    if len(p)!=3:return await m.answer('Использование: /setinvite ФСБ https://t.me/+...')
    async with SessionLocal() as db:
        o=(await db.execute(select(Organization).where(Organization.code==p[1]))).scalar_one_or_none()
        if not o:o=Organization(code=p[1],title=p[1]);db.add(o)
        o.invite_link=p[2];await db.commit()
    await m.answer('✅ Инвайт сохранён.')
@dp.message(Command('setrank'))
async def setrank(m:Message):
    if m.from_user.id!=OWNER_ID:return
    p=m.text.split(maxsplit=2)
    if len(p)!=3:return await m.answer('Использование: /setrank TELEGRAM_ID Звание')
    async with SessionLocal() as db:
        u=await get_user(db,int(p[1]))
        if not u:return await m.answer('Пользователь не найден.')
        u.rank=p[2];await db.commit()
    await m.answer('✅ Звание изменено.')
@dp.message(Command('addcriterion'))
async def addcriterion(m:Message):
    if m.from_user.id!=OWNER_ID:return
    p=[x.strip() for x in m.text.partition(' ')[2].split('|')]
    if len(p)<3:return await m.answer('Использование: /addcriterion ФСБ | Название | Описание')
    async with SessionLocal() as db:db.add(Criterion(organization=p[0],title=p[1],description=' | '.join(p[2:])));await db.commit()
    await m.answer('✅ Критерий добавлен.')
@dp.chat_member()
async def member_update(e:ChatMemberUpdated):
    if e.new_chat_member.status not in {ChatMemberStatus.LEFT,ChatMemberStatus.KICKED}:return
    async with SessionLocal() as db:
        o=(await db.execute(select(Organization).where(Organization.chat_id==e.chat.id))).scalar_one_or_none()
        if not o:return
        u=await get_user(db,e.new_chat_member.user.id)
        if u and u.organization==o.code:u.active=False;u.archived=True;u.archived_at=datetime.now(timezone.utc);await db.commit();logging.info('Archived %s from %s',u.telegram_id,o.code)
async def init_db():
    async with engine.begin() as c:await c.run_sync(Base.metadata.create_all)
    async with SessionLocal() as db:
        for code,title in [('ФСБ','ФСБ'),('ВЧ','Военная часть'),('ЕСС','Единая служба спасения'),('УМВД','УМВД')]:
            if not (await db.execute(select(Organization).where(Organization.code==code))).scalar_one_or_none():db.add(Organization(code=code,title=title))
        await db.commit()
@asynccontextmanager
async def lifespan(app):
    await init_db(); task=asyncio.create_task(dp.start_polling(bot,allowed_updates=['message','callback_query','chat_member']))
    yield
    task.cancel()
    try:await task
    except asyncio.CancelledError:pass
    await bot.session.close()
app=FastAPI(title='NEMAZING RP',version='1.0.0',lifespan=lifespan)
@app.get('/')
async def root():return {'service':'NEMAZING RP','status':'online','version':'1.0.0','health':'/health'}
@app.get('/health')
async def health():return {'status':'ok','service':'nemazing-rp-bot'}
if __name__=='__main__':
    import uvicorn;uvicorn.run('app.main:app',host='0.0.0.0',port=PORT)

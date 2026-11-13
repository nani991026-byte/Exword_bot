import asyncio
import os
import re
import sqlite3
from datetime import datetime, date, timedelta
from io import BytesIO
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, Message, BufferedInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
BOT_TOKEN = os.getenv("BOT_TOKEN")
DB_PATH = BASE_DIR / "data.sqlite3"
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is not set in .env")

router = Router()

# ---------- DB ----------

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    with db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                owner_user_id TEXT NOT NULL DEFAULT '0',
                name TEXT NOT NULL,
                telegram TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                owner_user_id TEXT NOT NULL DEFAULT '0',
                name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(owner_user_id, name)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS appointments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER NOT NULL,
                file_id INTEGER NOT NULL,
                service TEXT NOT NULL,
                appointment_date TEXT NOT NULL,
                appointment_time TEXT NOT NULL,
                comment TEXT DEFAULT '',
                source TEXT DEFAULT '',
                contact TEXT DEFAULT '',
                pain TEXT DEFAULT '',
                paid_status TEXT DEFAULT 'Не указано',
                paid_amount TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE CASCADE,
                FOREIGN KEY(file_id) REFERENCES files(id) ON DELETE CASCADE
            )
        """)

        # Safe migration of old databases.
        client_cols = {r[1] for r in conn.execute("PRAGMA table_info(clients)")}
        if "owner_user_id" not in client_cols:
            conn.execute("ALTER TABLE clients ADD COLUMN owner_user_id TEXT NOT NULL DEFAULT '0'")

        file_cols = {r[1] for r in conn.execute("PRAGMA table_info(files)")}
        if "owner_user_id" not in file_cols:
            conn.execute("ALTER TABLE files ADD COLUMN owner_user_id TEXT NOT NULL DEFAULT '0'")
            # Old files had a global UNIQUE(name). Rebuild to allow the same name for different users.
            conn.execute("PRAGMA foreign_keys=OFF")
            conn.execute("""
                CREATE TABLE files_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    owner_user_id TEXT NOT NULL DEFAULT '0',
                    name TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    UNIQUE(owner_user_id, name)
                )
            """)
            conn.execute("INSERT INTO files_new(id, owner_user_id, name, created_at) SELECT id, owner_user_id, name, created_at FROM files")
            conn.execute("DROP TABLE files")
            conn.execute("ALTER TABLE files_new RENAME TO files")
            conn.execute("PRAGMA foreign_keys=ON")

        appt_cols = {r[1] for r in conn.execute("PRAGMA table_info(appointments)")}
        migrations = {
            "file_id": "ALTER TABLE appointments ADD COLUMN file_id INTEGER NOT NULL DEFAULT 1",
            "source": "ALTER TABLE appointments ADD COLUMN source TEXT DEFAULT ''",
            "contact": "ALTER TABLE appointments ADD COLUMN contact TEXT DEFAULT ''",
            "pain": "ALTER TABLE appointments ADD COLUMN pain TEXT DEFAULT ''",
            "paid_status": "ALTER TABLE appointments ADD COLUMN paid_status TEXT DEFAULT 'Не указано'",
            "paid_amount": "ALTER TABLE appointments ADD COLUMN paid_amount TEXT DEFAULT ''",
        }
        for col, sql in migrations.items():
            if col not in appt_cols:
                conn.execute(sql)

        conn.execute("CREATE INDEX IF NOT EXISTS idx_files_owner ON files(owner_user_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_clients_owner ON clients(owner_user_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_appointments_date ON appointments(appointment_date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_appointments_file ON appointments(file_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_appointments_client ON appointments(client_id)")


def ensure_user_data(user_id: str):
    """Claim legacy rows (owner=0) for the first user who opens the migrated bot."""
    with db() as conn:
        has_user_files = conn.execute("SELECT 1 FROM files WHERE owner_user_id=? LIMIT 1", (user_id,)).fetchone()
        has_user_clients = conn.execute("SELECT 1 FROM clients WHERE owner_user_id=? LIMIT 1", (user_id,)).fetchone()
        if not has_user_files:
            conn.execute("UPDATE files SET owner_user_id=? WHERE owner_user_id='0'", (user_id,))
        if not has_user_clients:
            conn.execute("UPDATE clients SET owner_user_id=? WHERE owner_user_id='0'", (user_id,))
        # Guarantee a default file for every user.
        exists = conn.execute("SELECT id FROM files WHERE owner_user_id=? LIMIT 1", (user_id,)).fetchone()
        if not exists:
            conn.execute("INSERT INTO files(owner_user_id,name,created_at) VALUES(?,?,?)", (user_id, "Основной", datetime.now().isoformat(timespec="seconds")))


def get_file(file_id: int, user_id: str):
    with db() as conn:
        return conn.execute("SELECT * FROM files WHERE id=? AND owner_user_id=?", (file_id, user_id)).fetchone()


def all_files(user_id: str):
    with db() as conn:
        return conn.execute("""
            SELECT f.*, COUNT(a.id) AS appointments_count
            FROM files f LEFT JOIN appointments a ON a.file_id=f.id
            WHERE f.owner_user_id=?
            GROUP BY f.id ORDER BY f.name COLLATE NOCASE
        """, (user_id,)).fetchall()


def find_file_by_name(name: str, user_id: str):
    with db() as conn:
        return conn.execute("SELECT * FROM files WHERE lower(name)=lower(?) AND owner_user_id=?", (name.strip(), user_id)).fetchone()


def create_file(name: str, user_id: str) -> int:
    with db() as conn:
        cur = conn.execute("INSERT INTO files(owner_user_id,name,created_at) VALUES(?,?,?)", (user_id, name.strip(), datetime.now().isoformat(timespec="seconds")))
        return cur.lastrowid


def update_file_name(file_id: int, name: str, user_id: str):
    with db() as conn:
        conn.execute("UPDATE files SET name=? WHERE id=? AND owner_user_id=?", (name.strip(), file_id, user_id))


def delete_file(file_id: int, user_id: str):
    with db() as conn:
        conn.execute("DELETE FROM appointments WHERE file_id=?", (file_id,))
        conn.execute("DELETE FROM files WHERE id=? AND owner_user_id=?", (file_id, user_id))


def get_client(client_id: int, user_id: str):
    with db() as conn:
        return conn.execute("SELECT * FROM clients WHERE id=? AND owner_user_id=?", (client_id, user_id)).fetchone()


def find_or_create_client(name: str, telegram: str, user_id: str) -> int:
    telegram = telegram.strip()
    if telegram and not telegram.startswith("@"):
        telegram = "@" + telegram
    with db() as conn:
        row = conn.execute("SELECT id FROM clients WHERE owner_user_id=? AND telegram=?", (user_id, telegram)).fetchone()
        if row:
            conn.execute("UPDATE clients SET name=? WHERE id=? AND owner_user_id=?", (name.strip(), row["id"], user_id))
            return row["id"]
        cur = conn.execute("INSERT INTO clients(owner_user_id,name,telegram,created_at) VALUES(?,?,?,?)", (user_id, name.strip(), telegram, datetime.now().isoformat(timespec="seconds")))
        return cur.lastrowid


def add_appointment(user_id: str, client_id: int, file_id: int, service: str, appt_date: str, appt_time: str,
                    comment: str = "", source: str = "", contact: str = "", pain: str = "",
                    paid_status: str = "Не указано", paid_amount: str = "") -> int:
    with db() as conn:
        # Verify ownership before insert.
        ok = conn.execute("SELECT 1 FROM files WHERE id=? AND owner_user_id=?", (file_id, user_id)).fetchone()
        if not ok:
            raise ValueError("File does not belong to user")
        ok = conn.execute("SELECT 1 FROM clients WHERE id=? AND owner_user_id=?", (client_id, user_id)).fetchone()
        if not ok:
            raise ValueError("Client does not belong to user")
        cur = conn.execute("""
            INSERT INTO appointments(client_id,file_id,service,appointment_date,appointment_time,comment,source,contact,pain,paid_status,paid_amount,created_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
        """, (client_id,file_id,service.strip(),appt_date,appt_time,comment.strip(),source.strip(),contact.strip(),pain.strip(),paid_status.strip(),paid_amount.strip(),datetime.now().isoformat(timespec="seconds")))
        return cur.lastrowid


def get_appointment(appt_id: int, user_id: str):
    with db() as conn:
        return conn.execute("""
            SELECT a.*, c.name, c.telegram, f.name AS file_name
            FROM appointments a JOIN clients c ON c.id=a.client_id
            JOIN files f ON f.id=a.file_id
            WHERE a.id=? AND f.owner_user_id=? AND c.owner_user_id=?
        """, (appt_id,user_id,user_id)).fetchone()


def update_appointment(appt_id: int, field: str, value: str, user_id: str):
    allowed = {"appointment_date","appointment_time","service","comment","source","contact","pain","paid_status","paid_amount"}
    if field not in allowed:
        raise ValueError("Invalid field")
    with db() as conn:
        conn.execute(f"""UPDATE appointments SET {field}=? WHERE id=? AND file_id IN (SELECT id FROM files WHERE owner_user_id=?)""", (value,appt_id,user_id))


def delete_appointment(appt_id: int, user_id: str):
    with db() as conn:
        conn.execute("DELETE FROM appointments WHERE id=? AND file_id IN (SELECT id FROM files WHERE owner_user_id=?)", (appt_id,user_id))


def appointments_for_date(appt_date: str, user_id: str, file_id=None):
    with db() as conn:
        if file_id is None:
            return conn.execute("""
                SELECT a.*, c.name, c.telegram, f.name AS file_name
                FROM appointments a JOIN clients c ON c.id=a.client_id JOIN files f ON f.id=a.file_id
                WHERE a.appointment_date=? AND f.owner_user_id=? ORDER BY a.appointment_time,a.id
            """, (appt_date,user_id)).fetchall()
        return conn.execute("""
            SELECT a.*, c.name, c.telegram, f.name AS file_name
            FROM appointments a JOIN clients c ON c.id=a.client_id JOIN files f ON f.id=a.file_id
            WHERE a.appointment_date=? AND a.file_id=? AND f.owner_user_id=? ORDER BY a.appointment_time,a.id
        """, (appt_date,file_id,user_id)).fetchall()


def all_appointments(user_id: str, file_id=None):
    with db() as conn:
        if file_id is None:
            return conn.execute("""
                SELECT a.*, c.name, c.telegram, f.name AS file_name
                FROM appointments a JOIN clients c ON c.id=a.client_id JOIN files f ON f.id=a.file_id
                WHERE f.owner_user_id=? ORDER BY a.appointment_date,a.appointment_time,a.id
            """, (user_id,)).fetchall()
        return conn.execute("""
            SELECT a.*, c.name, c.telegram, f.name AS file_name
            FROM appointments a JOIN clients c ON c.id=a.client_id JOIN files f ON f.id=a.file_id
            WHERE a.file_id=? AND f.owner_user_id=? ORDER BY a.appointment_date,a.appointment_time,a.id
        """, (file_id,user_id)).fetchall()


def all_clients(user_id: str):
    with db() as conn:
        return conn.execute("""
            SELECT c.*, COUNT(a.id) AS appointments_count
            FROM clients c LEFT JOIN appointments a ON a.client_id=c.id
            WHERE c.owner_user_id=? GROUP BY c.id ORDER BY c.name COLLATE NOCASE
        """, (user_id,)).fetchall()

# ---------- UI ----------

def bottom_nav():
    kb=InlineKeyboardBuilder(); kb.button(text="⌂ Главная",callback_data="home"); kb.button(text="📁 Файлы",callback_data="files"); kb.button(text="📅 Расписание",callback_data="schedule"); kb.button(text="👥 Клиенты",callback_data="clients"); kb.adjust(2,2); return kb.as_markup()

def main_menu():
    kb=InlineKeyboardBuilder()
    for text,data in [("➕ Новая запись","add"),("📅 Расписание","schedule"),("📁 Мои файлы","files"),("👥 Клиенты","clients"),("📊 Excel","excel_menu"),("📄 Word","word_menu"),("🔎 Найти файл","file_search"),("☀️ Сегодня",f"viewdate:{date.today().isoformat()}"),("🌙 Завтра",f"viewdate:{(date.today()+timedelta(days=1)).isoformat()}")]: kb.button(text=text,callback_data=data)
    kb.adjust(2,2,2,2,2); return kb.as_markup()

def file_keyboard(prefix="file", user_id=None):
    kb=InlineKeyboardBuilder()
    for f in all_files(user_id): kb.button(text=f"📁 {f['name']} · {f['appointments_count']}",callback_data=f"{prefix}:{f['id']}")
    kb.button(text="➕ Новый файл",callback_data="new_file"); kb.button(text="🔎 Найти по названию",callback_data="file_search"); kb.button(text="⌂ Главная",callback_data="home"); kb.adjust(1,2,2); return kb.as_markup()

def file_manage_keyboard(file_id):
    kb=InlineKeyboardBuilder()
    for text,data in [("➕ Добавить запись",f"addfile:{file_id}"),("📅 Расписание",f"file_schedule:{file_id}"),("📊 Excel",f"excel:{file_id}"),("📄 Word",f"word:{file_id}"),("✏️ Переименовать",f"rename_file:{file_id}"),("🗑 Удалить",f"delete_file:{file_id}"),("‹ К файлам","files"),("⌂ Главная","home")]: kb.button(text=text,callback_data=data)
    kb.adjust(1,1,2,2,2); return kb.as_markup()

def cancel_kb():
    kb=InlineKeyboardBuilder(); kb.button(text="✕ Отмена",callback_data="cancel"); return kb.as_markup()

def skip_kb(callback_data):
    kb=InlineKeyboardBuilder(); kb.button(text="Пропустить",callback_data=callback_data); kb.button(text="✕ Отмена",callback_data="cancel"); kb.adjust(1); return kb.as_markup()

def source_keyboard(prefix):
    kb=InlineKeyboardBuilder()
    for text,val in [("MAX","MAX"),("Авито","Авито"),("ВКонтакте","ВКонтакте"),("Telegram","Telegram"),("Другое","Другое")]: kb.button(text=text,callback_data=f"{prefix}:{val}")
    kb.button(text="✕ Отмена",callback_data="cancel"); kb.adjust(2,2,1,1); return kb.as_markup()

def payment_keyboard(prefix):
    kb=InlineKeyboardBuilder()
    for text,val in [("✅ Оплачено","Оплачено"),("❌ Не оплачено","Не оплачено"),("💳 Частично","Частично"),("➖ Не указано","Не указано")]: kb.button(text=text,callback_data=f"{prefix}:{val}")
    kb.button(text="✕ Отмена",callback_data="cancel"); kb.adjust(2,2,1); return kb.as_markup()

def date_keyboard(prefix):
    today=date.today(); kb=InlineKeyboardBuilder(); labels=["Сегодня","Завтра","Послезавтра"]
    for i in range(3): kb.button(text=f"{labels[i]} · {(today+timedelta(days=i)):%d.%m}",callback_data=f"{prefix}:{(today+timedelta(days=i)).isoformat()}")
    for i in range(3,17):
        d=today+timedelta(days=i); kb.button(text=d.strftime("%d.%m"),callback_data=f"{prefix}:{d.isoformat()}")
    kb.button(text="✕ Отмена",callback_data="cancel"); kb.adjust(1,1,1,4,4,4,1); return kb.as_markup()

def time_keyboard(prefix):
    kb=InlineKeyboardBuilder()
    for h in range(8,21):
        for m in (0,30): kb.button(text=f"{h:02d}:{m:02d}",callback_data=f"{prefix}:{h:02d}:{m:02d}")
    kb.button(text="✕ Отмена",callback_data="cancel"); kb.adjust(4); return kb.as_markup()

def appointment_kb(appt_id):
    kb=InlineKeyboardBuilder()
    for text,data in [("📅 Дата",f"edit_date:{appt_id}"),("🕐 Время",f"edit_time:{appt_id}"),("💅 Услуга",f"edit_service:{appt_id}"),("📍 Источник",f"edit_source:{appt_id}"),("📞 Контакт",f"edit_contact:{appt_id}"),("🎯 Запрос",f"edit_pain:{appt_id}"),("💰 Оплата",f"edit_paid:{appt_id}"),("📝 Комментарий",f"edit_comment:{appt_id}"),("🗑 Удалить",f"delete:{appt_id}"),("‹ К расписанию","schedule")]: kb.button(text=text,callback_data=data)
    kb.adjust(2,2,2,2,2); return kb.as_markup()

def comment_keyboard():
    kb=InlineKeyboardBuilder(); kb.button(text="Без комментария",callback_data="no_comment"); kb.button(text="✕ Отмена",callback_data="cancel"); kb.adjust(1); return kb.as_markup()

def format_appt(a):
    lines=[f"📁 {a['file_name']}",f"🕐 <b>{a['appointment_time']}</b> — {a['name']}",f"💅 {a['service']}",f"💬 {a['telegram']}"]
    if a["contact"]: lines.append(f"📞 {a['contact']}")
    if a["source"]: lines.append(f"📍 {a['source']}")
    if a["pain"]: lines.append(f"🎯 {a['pain']}")
    paid=a["paid_status"] or "Не указано"; paid += f" — {a['paid_amount']} ₽" if a["paid_amount"] else ""; lines.append(f"💰 {paid}")
    if a["comment"]: lines.append(f"📝 {a['comment']}")
    return "\n".join(lines)

def format_day_title(d):
    v=datetime.strptime(d,"%Y-%m-%d").date(); return f"{['Пн','Вт','Ср','Чт','Пт','Сб','Вс'][v.weekday()]}, {v:%d.%m.%Y}"

# ---------- FSM ----------
class AddAppointment(StatesGroup):
    name=State(); telegram=State(); contact=State(); source=State(); custom_source=State(); service=State(); pain=State(); date=State(); time=State(); payment=State(); paid_amount=State(); comment=State()
class EditText(StatesGroup): value=State()
class EditPaid(StatesGroup): amount=State()
class FileText(StatesGroup): name=State()
class FileSearch(StatesGroup): query=State()

# ---------- common ----------
def uid_from_event(event):
    return str(event.from_user.id)

async def ensure_event_user(event):
    uid=uid_from_event(event); ensure_user_data(uid); return uid

# ---------- handlers ----------
@router.message(CommandStart())
async def start(message: Message):
    await ensure_event_user(message)
    await message.answer("<b>📋 Клиенты & расписание</b>\n\nВыберите действие:", reply_markup=main_menu())

@router.callback_query(F.data=="cancel")
async def cancel(call: CallbackQuery,state:FSMContext):
    await ensure_event_user(call); await state.clear(); await call.message.edit_text("<b>Главное меню</b>\n\nВыберите действие:",reply_markup=main_menu()); await call.answer()

@router.callback_query(F.data=="add")
async def add_start(call: CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); await state.clear(); await state.update_data(owner_user_id=uid)
    await call.message.edit_text("<b>➕ Новая запись</b>\n\n📁 Выберите файл:",reply_markup=file_keyboard("addfile",uid)); await call.answer()

@router.callback_query(F.data.startswith("addfile:"))
async def add_file_selected(call: CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); file_id=int(call.data.split(":")[1]); f=get_file(file_id,uid)
    if not f: await call.answer("Файл не найден или недоступен",show_alert=True); return
    await state.update_data(owner_user_id=uid,file_id=file_id,file_name=f["name"]); await state.set_state(AddAppointment.name)
    await call.message.edit_text(f"<b>📁 {f['name']}</b>\n\n👤 Введите имя клиента:",reply_markup=cancel_kb()); await call.answer()

@router.message(AddAppointment.name)
async def add_name(message:Message,state:FSMContext):
    v=(message.text or "").strip()
    if not v: await message.answer("Введите имя клиента.",reply_markup=cancel_kb()); return
    await state.update_data(name=v); await state.set_state(AddAppointment.telegram); await message.answer("💬 Telegram-ник клиента:\nнапример, <code>@anna</code>",reply_markup=cancel_kb())

@router.message(AddAppointment.telegram)
async def add_telegram(message:Message,state:FSMContext):
    v=(message.text or "").strip()
    if not v: await message.answer("Введите Telegram-ник.",reply_markup=cancel_kb()); return
    await state.update_data(telegram=v); await state.set_state(AddAppointment.contact); await message.answer("📞 Контакт клиента\n\nТелефон, VK или другой контакт. Можно пропустить.",reply_markup=skip_kb("skip_contact"))

@router.callback_query(AddAppointment.contact,F.data=="skip_contact")
async def skip_contact(call:CallbackQuery,state:FSMContext):
    await state.update_data(contact=""); await state.set_state(AddAppointment.source); await call.message.edit_text("📍 Откуда пришёл клиент?",reply_markup=source_keyboard("addsource")); await call.answer()

@router.message(AddAppointment.contact)
async def add_contact(message:Message,state:FSMContext):
    await state.update_data(contact=(message.text or "").strip()); await state.set_state(AddAppointment.source); await message.answer("📍 Откуда пришёл клиент?",reply_markup=source_keyboard("addsource"))

@router.callback_query(AddAppointment.source,F.data.startswith("addsource:"))
async def add_source(call:CallbackQuery,state:FSMContext):
    source=call.data.split(":",1)[1]
    if source=="Другое": await state.set_state(AddAppointment.custom_source); await call.message.edit_text("📍 Введите источник клиента:",reply_markup=cancel_kb())
    else: await state.update_data(source=source); await state.set_state(AddAppointment.service); await call.message.edit_text("💅 Какая услуга?",reply_markup=cancel_kb())
    await call.answer()

@router.message(AddAppointment.custom_source)
async def add_custom_source(message:Message,state:FSMContext):
    v=(message.text or "").strip()
    if not v: await message.answer("Введите источник.",reply_markup=cancel_kb()); return
    await state.update_data(source=v); await state.set_state(AddAppointment.service); await message.answer("💅 Какая услуга?",reply_markup=cancel_kb())

@router.message(AddAppointment.service)
async def add_service(message:Message,state:FSMContext):
    v=(message.text or "").strip()
    if not v: await message.answer("Введите название услуги.",reply_markup=cancel_kb()); return
    await state.update_data(service=v); await state.set_state(AddAppointment.pain); await message.answer("🎯 Краткая боль / запрос клиента?\n\nМожно пропустить.",reply_markup=skip_kb("skip_pain"))

@router.callback_query(AddAppointment.pain,F.data=="skip_pain")
async def skip_pain(call:CallbackQuery,state:FSMContext):
    await state.update_data(pain=""); await state.set_state(AddAppointment.date); await call.message.edit_text("📅 Выберите дату:",reply_markup=date_keyboard("adddate")); await call.answer()

@router.message(AddAppointment.pain)
async def add_pain(message:Message,state:FSMContext):
    await state.update_data(pain=(message.text or "").strip()); await state.set_state(AddAppointment.date); await message.answer("📅 Выберите дату:",reply_markup=date_keyboard("adddate"))

@router.callback_query(AddAppointment.date,F.data.startswith("adddate:"))
async def add_date(call:CallbackQuery,state:FSMContext):
    await state.update_data(date=call.data.split(":",1)[1]); await state.set_state(AddAppointment.time); await call.message.edit_text("🕐 Выберите время:",reply_markup=time_keyboard("addtime")); await call.answer()

@router.callback_query(AddAppointment.time,F.data.startswith("addtime:"))
async def add_time(call:CallbackQuery,state:FSMContext):
    _,h,m=call.data.split(":"); await state.update_data(time=f"{h}:{m}"); await state.set_state(AddAppointment.payment); await call.message.edit_text("💰 Статус оплаты:",reply_markup=payment_keyboard("addpaid")); await call.answer()

@router.callback_query(AddAppointment.payment,F.data.startswith("addpaid:"))
async def add_payment(call:CallbackQuery,state:FSMContext):
    status=call.data.split(":",1)[1]; await state.update_data(paid_status=status)
    if status=="Частично": await state.set_state(AddAppointment.paid_amount); await call.message.edit_text("💳 Сколько оплачено?\n\nВведите сумму, например 1500",reply_markup=cancel_kb())
    else: await state.set_state(AddAppointment.comment); await call.message.edit_text("📝 Комментарий?\n\nНапишите текст или нажмите кнопку ниже.",reply_markup=comment_keyboard())
    await call.answer()

@router.message(AddAppointment.paid_amount)
async def add_paid_amount(message:Message,state:FSMContext):
    v=(message.text or "").strip().replace(",",".")
    if not v: await message.answer("Введите сумму.",reply_markup=cancel_kb()); return
    await state.update_data(paid_amount=v); await state.set_state(AddAppointment.comment); await message.answer("📝 Комментарий?\n\nНапишите текст или нажмите кнопку ниже.",reply_markup=comment_keyboard())

@router.callback_query(AddAppointment.comment,F.data=="no_comment")
async def no_comment(call:CallbackQuery,state:FSMContext):
    await save_appointment(call.message,state,"",True); await call.answer()

@router.message(AddAppointment.comment)
async def add_comment(message:Message,state:FSMContext):
    await save_appointment(message,state,(message.text or "").strip(),False)

async def save_appointment(target:Message,state:FSMContext,comment:str,edit:bool):
    data=await state.get_data(); uid=data.get("owner_user_id") or str(target.from_user.id)
    try:
        cid=find_or_create_client(data["name"],data["telegram"],uid)
        appt_id=add_appointment(uid,cid,int(data["file_id"]),data["service"],data["date"],data["time"],comment,data.get("source",""),data.get("contact",""),data.get("pain",""),data.get("paid_status","Не указано"),data.get("paid_amount",""))
    except Exception:
        await target.answer("❌ Не удалось сохранить запись. Попробуйте ещё раз.",reply_markup=main_menu()); await state.clear(); return
    await state.clear(); a=get_appointment(appt_id,uid); text="<b>✅ Запись сохранена</b>\n\n"+format_appt(a)
    if edit: await target.edit_text(text,reply_markup=appointment_kb(appt_id))
    else: await target.answer(text,reply_markup=appointment_kb(appt_id))

# Schedule
@router.callback_query(F.data=="schedule")
async def schedule(call:CallbackQuery):
    uid=await ensure_event_user(call); await call.message.edit_text("<b>📅 Расписание</b>\n\nВыберите файл:",reply_markup=file_keyboard("schedfile",uid)); await call.answer()

@router.callback_query(F.data.startswith("schedfile:"))
async def schedule_file(call:CallbackQuery):
    uid=await ensure_event_user(call); fid=int(call.data.split(":")[1]); f=get_file(fid,uid)
    if not f: await call.answer("Файл недоступен",show_alert=True); return
    await call.message.edit_text(f"<b>📅 {f['name']}</b>\n\nВыберите день:",reply_markup=date_keyboard(f"viewdate:{fid}")); await call.answer()

@router.callback_query(F.data.startswith("viewdate:"))
async def view_date(call:CallbackQuery):
    uid=await ensure_event_user(call); p=call.data.split(":")
    if len(p)==2: fid=None; d=p[1]
    else: fid=int(p[1]); d=p[2]
    if fid and not get_file(fid,uid): await call.answer("Файл недоступен",show_alert=True); return
    rows=appointments_for_date(d,uid,fid); title=format_day_title(d); fname=get_file(fid,uid)["name"] if fid else "Все файлы"
    text=f"<b>📅 {title}</b>\n📁 {fname} · {len(rows)} запис."
    if rows: text += "\n\n"+"\n\n".join(format_appt(a) for a in rows)
    else: text += "\n\nСвободный день ✨"
    kb=InlineKeyboardBuilder()
    for a in rows: kb.button(text=f"{a['appointment_time']} · {a['name']}",callback_data=f"open:{a['id']}")
    if fid: kb.button(text="➕ Добавить сюда",callback_data=f"addfile:{fid}"); kb.button(text="📁 Файл",callback_data=f"file:{fid}")
    else: kb.button(text="➕ Добавить запись",callback_data="add")
    kb.button(text="‹ Назад",callback_data=f"file_schedule:{fid}" if fid else "schedule"); kb.button(text="⌂ Главная",callback_data="home"); kb.adjust(1); await call.message.edit_text(text,reply_markup=kb.as_markup()); await call.answer()

@router.callback_query(F.data.startswith("file_schedule:"))
async def file_schedule(call:CallbackQuery):
    uid=await ensure_event_user(call); fid=int(call.data.split(":")[1]); f=get_file(fid,uid)
    if not f: await call.answer("Файл недоступен",show_alert=True); return
    await call.message.edit_text(f"<b>📅 {f['name']}</b>\n\nВыберите день:",reply_markup=date_keyboard(f"viewdate:{fid}")); await call.answer()

@router.callback_query(F.data.startswith("open:"))
async def open_appt(call:CallbackQuery):
    uid=await ensure_event_user(call); aid=int(call.data.split(":")[1]); a=get_appointment(aid,uid)
    if not a: await call.answer("Запись недоступна",show_alert=True); return
    await call.message.edit_text(f"<b>📋 Запись</b>\n\n📅 {format_day_title(a['appointment_date'])}\n\n{format_appt(a)}",reply_markup=appointment_kb(aid)); await call.answer()

@router.callback_query(F.data.startswith("delete:"))
async def delete_appt(call:CallbackQuery):
    uid=await ensure_event_user(call); aid=int(call.data.split(":")[1]); a=get_appointment(aid,uid)
    if not a: await call.answer("Запись недоступна",show_alert=True); return
    delete_appointment(aid,uid); await call.message.edit_text("<b>🗑 Запись удалена</b>",reply_markup=main_menu()); await call.answer()

# Edits
@router.callback_query(F.data.startswith("edit_date:"))
async def edit_date(call:CallbackQuery):
    uid=await ensure_event_user(call); aid=int(call.data.split(":")[1]);
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    await call.message.edit_text("<b>📅 Новая дата</b>",reply_markup=date_keyboard(f"setdate:{aid}")); await call.answer()

@router.callback_query(F.data.startswith("setdate:"))
async def set_date(call:CallbackQuery):
    uid=await ensure_event_user(call); _,aid,d=call.data.split(":"); aid=int(aid)
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    update_appointment(aid,"appointment_date",d,uid); a=get_appointment(aid,uid); await call.message.edit_text("<b>✅ Дата изменена</b>\n\n"+format_appt(a),reply_markup=appointment_kb(aid)); await call.answer()

@router.callback_query(F.data.startswith("edit_time:"))
async def edit_time(call:CallbackQuery):
    uid=await ensure_event_user(call); aid=int(call.data.split(":")[1]);
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    await call.message.edit_text("<b>🕐 Новое время</b>",reply_markup=time_keyboard(f"settime:{aid}")); await call.answer()

@router.callback_query(F.data.startswith("settime:"))
async def set_time(call:CallbackQuery):
    uid=await ensure_event_user(call); p=call.data.split(":"); aid=int(p[1]); val=f"{p[2]}:{p[3]}"
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    update_appointment(aid,"appointment_time",val,uid); a=get_appointment(aid,uid); await call.message.edit_text("<b>✅ Время изменено</b>\n\n"+format_appt(a),reply_markup=appointment_kb(aid)); await call.answer()

async def edit_text_start(call,state,field,prompt):
    uid=await ensure_event_user(call); aid=int(call.data.split(":")[1]);
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    await state.set_state(EditText.value); await state.update_data(appt_id=aid,field=field,owner_user_id=uid); await call.message.edit_text(prompt,reply_markup=cancel_kb()); await call.answer()

@router.callback_query(F.data.startswith("edit_service:"))
async def edit_service(call,state:FSMContext): await edit_text_start(call,state,"service","💅 Введите новую услугу:")
@router.callback_query(F.data.startswith("edit_comment:"))
async def edit_comment(call,state:FSMContext): await edit_text_start(call,state,"comment","📝 Введите новый комментарий:")
@router.callback_query(F.data.startswith("edit_contact:"))
async def edit_contact(call,state:FSMContext): await edit_text_start(call,state,"contact","📞 Введите новый контакт:")
@router.callback_query(F.data.startswith("edit_pain:"))
async def edit_pain(call,state:FSMContext): await edit_text_start(call,state,"pain","🎯 Введите новый запрос:")

@router.callback_query(F.data.startswith("edit_source:"))
async def edit_source(call:CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); aid=int(call.data.split(":")[1]);
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    await call.message.edit_text("📍 Выберите источник:",reply_markup=source_keyboard(f"setsource:{aid}")); await call.answer()

@router.callback_query(F.data.startswith("setsource:"))
async def set_source(call:CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); p=call.data.split(":",2); aid=int(p[1]); source=p[2]
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    if source=="Другое": await state.set_state(EditText.value); await state.update_data(appt_id=aid,field="source",owner_user_id=uid); await call.message.edit_text("📍 Введите источник:",reply_markup=cancel_kb())
    else: update_appointment(aid,"source",source,uid); a=get_appointment(aid,uid); await call.message.edit_text("<b>✅ Источник изменён</b>\n\n"+format_appt(a),reply_markup=appointment_kb(aid))
    await call.answer()

@router.callback_query(F.data.startswith("edit_paid:"))
async def edit_paid(call:CallbackQuery):
    uid=await ensure_event_user(call); aid=int(call.data.split(":")[1]);
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    await call.message.edit_text("💰 Новый статус оплаты:",reply_markup=payment_keyboard(f"setpaid:{aid}")); await call.answer()

@router.callback_query(F.data.startswith("setpaid:"))
async def set_paid(call:CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); p=call.data.split(":",2); aid=int(p[1]); status=p[2]
    if not get_appointment(aid,uid): await call.answer("Запись недоступна",show_alert=True); return
    update_appointment(aid,"paid_status",status,uid)
    if status=="Частично": await state.set_state(EditPaid.amount); await state.update_data(appt_id=aid,owner_user_id=uid); await call.message.edit_text("💳 Введите оплаченную сумму:",reply_markup=cancel_kb())
    else: update_appointment(aid,"paid_amount","",uid); a=get_appointment(aid,uid); await call.message.edit_text("<b>✅ Оплата изменена</b>\n\n"+format_appt(a),reply_markup=appointment_kb(aid))
    await call.answer()

@router.message(EditPaid.amount)
async def edit_paid_amount(message:Message,state:FSMContext):
    v=(message.text or "").strip().replace(",","."); data=await state.get_data(); uid=data.get("owner_user_id",str(message.from_user.id)); aid=data.get("appt_id")
    if not v: await message.answer("Введите сумму.",reply_markup=cancel_kb()); return
    if not get_appointment(aid,uid): await state.clear(); await message.answer("Запись недоступна.",reply_markup=main_menu()); return
    update_appointment(aid,"paid_amount",v,uid); await state.clear(); a=get_appointment(aid,uid); await message.answer("<b>✅ Оплата изменена</b>\n\n"+format_appt(a),reply_markup=appointment_kb(aid))

@router.message(EditText.value)
async def edit_text_value(message:Message,state:FSMContext):
    v=(message.text or "").strip(); data=await state.get_data(); uid=data.get("owner_user_id",str(message.from_user.id)); aid=data.get("appt_id"); field=data.get("field")
    if not v: await message.answer("Введите текст.",reply_markup=cancel_kb()); return
    if not get_appointment(aid,uid): await state.clear(); await message.answer("Запись недоступна.",reply_markup=main_menu()); return
    update_appointment(aid,field,v,uid); await state.clear(); a=get_appointment(aid,uid); await message.answer("<b>✅ Изменено</b>\n\n"+format_appt(a),reply_markup=appointment_kb(aid))

# Files
@router.callback_query(F.data=="files")
async def files_menu(call:CallbackQuery):
    uid=await ensure_event_user(call); await call.message.edit_text("<b>📁 Мои файлы</b>\n\nВыберите файл:",reply_markup=file_keyboard("file",uid)); await call.answer()

@router.callback_query(F.data=="new_file")
async def new_file_start(call:CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); await state.clear(); await state.set_state(FileText.name); await state.update_data(owner_user_id=uid); await call.message.edit_text("<b>➕ Новый файл</b>\n\nВведите название:",reply_markup=cancel_kb()); await call.answer()

@router.message(FileText.name)
async def new_file_name(message:Message,state:FSMContext):
    name=(message.text or "").strip(); data=await state.get_data(); uid=data.get("owner_user_id",str(message.from_user.id)); rename_id=data.get("rename_file_id")
    if not name: await message.answer("Введите название.",reply_markup=cancel_kb()); return
    if find_file_by_name(name,uid) and (not rename_id or int(find_file_by_name(name,uid)["id"])!=int(rename_id)):
        await message.answer("Файл с таким названием уже есть. Введите другое название.",reply_markup=cancel_kb()); return
    if rename_id: update_file_name(int(rename_id),name,uid); await state.clear(); await message.answer(f"<b>✅ Файл переименован</b>\n\n📁 {name}",reply_markup=file_manage_keyboard(int(rename_id)))
    else:
        fid=create_file(name,uid); await state.clear(); await message.answer(f"<b>✅ Файл создан</b>\n\n📁 {name}\n\nТеперь можно добавлять записи.",reply_markup=file_manage_keyboard(fid))

@router.callback_query(F.data.startswith("file:"))
async def file_card(call:CallbackQuery):
    uid=await ensure_event_user(call); fid=int(call.data.split(":")[1]); f=get_file(fid,uid)
    if not f: await call.answer("Файл недоступен",show_alert=True); return
    # refresh count
    with db() as conn: count=conn.execute("SELECT COUNT(*) FROM appointments WHERE file_id=?",(fid,)).fetchone()[0]
    await call.message.edit_text(f"<b>📁 {f['name']}</b>\n\n📋 Записей: {count}\n\nВыберите действие:",reply_markup=file_manage_keyboard(fid)); await call.answer()

@router.callback_query(F.data=="file_search")
async def file_search_start(call:CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); await state.clear(); await state.set_state(FileSearch.query); await state.update_data(owner_user_id=uid); await call.message.edit_text("<b>🔎 Найти файл</b>\n\nВведите название или его часть:",reply_markup=cancel_kb()); await call.answer()

@router.message(FileSearch.query)
async def file_search_query(message:Message,state:FSMContext):
    query=(message.text or "").strip().lower(); data=await state.get_data(); uid=data.get("owner_user_id",str(message.from_user.id))
    if not query: await message.answer("Введите название.",reply_markup=cancel_kb()); return
    with db() as conn:
        rows=conn.execute("""SELECT f.*,COUNT(a.id) AS appointments_count FROM files f LEFT JOIN appointments a ON a.file_id=f.id WHERE f.owner_user_id=? AND lower(f.name) LIKE ? GROUP BY f.id ORDER BY f.name COLLATE NOCASE""",(uid,f"%{query}%")).fetchall()
    await state.clear(); kb=InlineKeyboardBuilder()
    for f in rows: kb.button(text=f"📁 {f['name']} · {f['appointments_count']}",callback_data=f"file:{f['id']}")
    kb.button(text="🔎 Новый поиск",callback_data="file_search"); kb.button(text="‹ К файлам",callback_data="files"); kb.adjust(1)
    await message.answer(f"<b>🔎 Найдено: {len(rows)}</b>" if rows else "<b>Ничего не найдено.</b>",reply_markup=kb.as_markup())

@router.callback_query(F.data.startswith("rename_file:"))
async def rename_file_start(call:CallbackQuery,state:FSMContext):
    uid=await ensure_event_user(call); fid=int(call.data.split(":")[1]);
    if not get_file(fid,uid): await call.answer("Файл недоступен",show_alert=True); return
    await state.clear(); await state.set_state(FileText.name); await state.update_data(rename_file_id=fid,owner_user_id=uid); await call.message.edit_text("<b>✏️ Переименовать файл</b>\n\nВведите новое название:",reply_markup=cancel_kb()); await call.answer()

@router.callback_query(F.data.startswith("delete_file:"))
async def delete_file_confirm(call:CallbackQuery):
    uid=await ensure_event_user(call); fid=int(call.data.split(":")[1]); f=get_file(fid,uid)
    if not f: await call.answer("Файл недоступен",show_alert=True); return
    if f["name"]=="Основной": await call.answer("Основной файл нельзя удалить",show_alert=True); return
    kb=InlineKeyboardBuilder(); kb.button(text="Да, удалить",callback_data=f"confirm_delete_file:{fid}"); kb.button(text="Отмена",callback_data=f"file:{fid}"); kb.adjust(1)
    await call.message.edit_text(f"<b>🗑 Удалить «{f['name']}»?</b>\n\nЗаписи внутри тоже будут удалены.",reply_markup=kb.as_markup()); await call.answer()

@router.callback_query(F.data.startswith("confirm_delete_file:"))
async def delete_file_confirmed(call:CallbackQuery):
    uid=await ensure_event_user(call); fid=int(call.data.split(":")[1]); f=get_file(fid,uid)
    if not f: await call.answer("Файл уже удалён",show_alert=True); return
    delete_file(fid,uid); await call.message.edit_text("<b>🗑 Файл удалён</b>",reply_markup=file_keyboard("file",uid)); await call.answer()

# Clients
@router.callback_query(F.data=="clients")
async def clients(call:CallbackQuery):
    uid=await ensure_event_user(call); rows=all_clients(uid); kb=InlineKeyboardBuilder()
    for c in rows: kb.button(text=f"👤 {c['name']} · {c['appointments_count']}",callback_data=f"client:{c['id']}")
    kb.button(text="➕ Новая запись",callback_data="add"); kb.button(text="‹ Главная",callback_data="home"); kb.adjust(2,1,1)
    await call.message.edit_text("<b>👥 Клиенты</b>\n\n"+("Выберите клиента:" if rows else "Пока клиентов нет."),reply_markup=kb.as_markup()); await call.answer()

@router.callback_query(F.data.startswith("client:"))
async def client_card(call:CallbackQuery):
    uid=await ensure_event_user(call); cid=int(call.data.split(":")[1]); c=get_client(cid,uid)
    if not c: await call.answer("Клиент недоступен",show_alert=True); return
    kb=InlineKeyboardBuilder(); kb.button(text="➕ Новая запись",callback_data="add"); kb.button(text="‹ К клиентам",callback_data="clients"); kb.button(text="⌂ Главная",callback_data="home"); kb.adjust(1)
    await call.message.edit_text(f"<b>👤 {c['name']}</b>\n\n💬 {c['telegram']}\n📅 Записей: {c['appointments_count']}",reply_markup=kb.as_markup()); await call.answer()

# Export
@router.callback_query(F.data=="excel_menu")
async def excel_menu(call:CallbackQuery):
    uid=await ensure_event_user(call); await call.message.edit_text("<b>📊 Excel</b>\n\nВыберите файл:",reply_markup=file_keyboard("excelfile",uid)); await call.answer()
@router.callback_query(F.data.startswith("excelfile:"))
async def excel_file_selected(call:CallbackQuery): await excel_export(call,int(call.data.split(":")[1]))
@router.callback_query(F.data.startswith("excel:"))
async def excel_file_direct(call:CallbackQuery): await excel_export(call,int(call.data.split(":")[1]))

async def excel_export(call:CallbackQuery,file_id):
    uid=await ensure_event_user(call); f=get_file(file_id,uid)
    if not f: await call.answer("Файл недоступен",show_alert=True); return
    rows=all_appointments(uid,file_id)
    if not rows: await call.answer("В этом файле пока нет записей",show_alert=True); return
    wb=Workbook(); ws=wb.active; ws.title="Расписание"; ws["A1"]=f["name"]+" · КЛИЕНТЫ И РАСПИСАНИЕ"; ws.merge_cells("A1:H1"); ws["A1"].font=Font(bold=True,color="FFFFFF",size=16); ws["A1"].fill=PatternFill("solid",fgColor="6F4E37"); ws["A1"].alignment=Alignment(horizontal="left"); headers=["Дата","Время","Клиент","Контакт","Источник","Краткая боль / запрос","Оплачено","Комментарий"]; ws.append([]); ws.append(headers)
    for c in ws[3]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="6F4E37"); c.alignment=Alignment(horizontal="center",wrap_text=True)
    thin=Side(style="thin",color="D9D9D9")
    for r in rows:
        d=datetime.strptime(r["appointment_date"],"%Y-%m-%d").strftime("%d.%m.%Y"); paid=r["paid_status"] or "Не указано"; paid += f" ({r['paid_amount']} ₽)" if r["paid_amount"] else ""
        ws.append([d,r["appointment_time"],r["name"],r["contact"],r["source"],r["pain"],paid,r["comment"]])
    for row in ws.iter_rows(min_row=4):
        for c in row: c.border=Border(bottom=thin); c.alignment=Alignment(vertical="top",wrap_text=True)
    for i,w in enumerate([14,10,22,28,18,35,20,40],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A4"; ws.auto_filter.ref=ws.dimensions
    cws=wb.create_sheet("Клиенты"); cws.append(["Имя","Telegram"]); [setattr(c,"font",Font(bold=True,color="FFFFFF")) for c in cws[1]]; [setattr(c,"fill",PatternFill("solid",fgColor="6F4E37")) for c in cws[1]]
    for c in all_clients(uid): cws.append([c["name"],c["telegram"]])
    cws.column_dimensions["A"].width=25; cws.column_dimensions["B"].width=30
    bio=BytesIO(); wb.save(bio); bio.seek(0); fname=re.sub(r"[\\/:*?\"<>|]","-",f["name"])+".xlsx"
    msg=await call.message.answer_document(BufferedInputFile(bio.read(),filename=fname),caption="📊 <b>Excel готов</b>")
    await call.answer("Готово")

@router.callback_query(F.data=="word_menu")
async def word_menu(call:CallbackQuery):
    uid=await ensure_event_user(call); await call.message.edit_text("<b>📄 Word</b>\n\nВыберите файл:",reply_markup=file_keyboard("wordfile",uid)); await call.answer()
@router.callback_query(F.data.startswith("wordfile:"))
async def word_file_selected(call:CallbackQuery): await word_export(call,int(call.data.split(":")[1]))
@router.callback_query(F.data.startswith("word:"))
async def word_file_direct(call:CallbackQuery): await word_export(call,int(call.data.split(":")[1]))

async def word_export(call:CallbackQuery,file_id):
    uid=await ensure_event_user(call); f=get_file(file_id,uid)
    if not f: await call.answer("Файл недоступен",show_alert=True); return
    rows=all_appointments(uid,file_id)
    if not rows: await call.answer("В этом файле пока нет записей",show_alert=True); return
    doc=Document(); doc.add_heading(f["name"],0); doc.add_paragraph(f"Расписание · создано {datetime.now():%d.%m.%Y %H:%M}")
    table=doc.add_table(rows=1,cols=8); table.style="Table Grid"
    headers=["Дата","Время","Клиент","Контакт","Источник","Запрос","Оплачено","Комментарий"]
    for i,h in enumerate(headers): table.rows[0].cells[i].text=h
    for r in rows:
        cells=table.add_row().cells; paid=r["paid_status"] or "Не указано"; paid += f" ({r['paid_amount']} ₽)" if r["paid_amount"] else ""
        vals=[datetime.strptime(r["appointment_date"],"%Y-%m-%d").strftime("%d.%m.%Y"),r["appointment_time"],r["name"],r["contact"],r["source"],r["pain"],paid,r["comment"]]
        for i,v in enumerate(vals): cells[i].text=v or "—"
    for p in doc.paragraphs: p.style.font if False else None
    bio=BytesIO(); doc.save(bio); bio.seek(0); fname=re.sub(r"[\\/:*?\"<>|]","-",f["name"])+".docx"
    msg=await call.message.answer_document(BufferedInputFile(bio.read(),filename=fname),caption="📄 <b>Word готов</b>\nСообщение с файлом удалится через 4 минуты.")
    await call.answer("Готово")
    asyncio.create_task(delete_message_later(call.bot,msg.chat.id,msg.message_id,240))

async def delete_message_later(bot:Bot,chat_id:int,message_id:int,delay:int):
    await asyncio.sleep(delay)
    try: await bot.delete_message(chat_id,message_id)
    except Exception: pass

@router.callback_query(F.data=="home")
async def home(call:CallbackQuery):
    await ensure_event_user(call); await call.message.edit_text("<b>Главное меню</b>\n\nВыберите действие:",reply_markup=main_menu()); await call.answer()

async def main():
    init_db(); bot=Bot(token=BOT_TOKEN,default=DefaultBotProperties(parse_mode=ParseMode.HTML)); dp=Dispatcher(); dp.include_router(router); await dp.start_polling(bot)

if __name__=="__main__": asyncio.run(main())

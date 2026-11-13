import asyncio
import os
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

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
BOT_TOKEN = os.getenv("BOT_TOKEN")
DB_PATH = BASE_DIR / "data.sqlite3"

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is not set in .env")

router = Router()

# -------------------- DB --------------------

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                telegram TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS appointments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER NOT NULL,
                file_id INTEGER NOT NULL DEFAULT 1,
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
                FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE CASCADE
            )
        """)

        # Миграция существующей базы: старые записи сохраняются.
        existing = {row[1] for row in conn.execute("PRAGMA table_info(appointments)").fetchall()}
        conn.execute("INSERT OR IGNORE INTO files(name, created_at) VALUES(?,?)", ("Основной", datetime.now().isoformat(timespec="seconds")))
        migrations = {
            "file_id": "ALTER TABLE appointments ADD COLUMN file_id INTEGER NOT NULL DEFAULT 1",

            "source": "ALTER TABLE appointments ADD COLUMN source TEXT DEFAULT ''",
            "contact": "ALTER TABLE appointments ADD COLUMN contact TEXT DEFAULT ''",
            "pain": "ALTER TABLE appointments ADD COLUMN pain TEXT DEFAULT ''",
            "paid_status": "ALTER TABLE appointments ADD COLUMN paid_status TEXT DEFAULT 'Не указано'",
            "paid_amount": "ALTER TABLE appointments ADD COLUMN paid_amount TEXT DEFAULT ''",
        }
        for column, sql in migrations.items():
            if column not in existing:
                conn.execute(sql)

        conn.execute("CREATE INDEX IF NOT EXISTS idx_appointments_date ON appointments(appointment_date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_appointments_file ON appointments(file_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_clients_telegram ON clients(telegram)")


def get_file(file_id: int):
    with db() as conn:
        return conn.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()


def all_files():
    with db() as conn:
        return conn.execute("""
            SELECT f.*, COUNT(a.id) AS appointments_count
            FROM files f
            LEFT JOIN appointments a ON a.file_id=f.id
            GROUP BY f.id
            ORDER BY f.name COLLATE NOCASE
        """).fetchall()


def find_file_by_name(name: str):
    with db() as conn:
        return conn.execute("SELECT * FROM files WHERE lower(name)=lower(?)", (name.strip(),)).fetchone()


def create_file(name: str) -> int:
    with db() as conn:
        cur = conn.execute(
            "INSERT INTO files(name, created_at) VALUES(?,?)",
            (name.strip(), datetime.now().isoformat(timespec="seconds")),
        )
        return cur.lastrowid


def appointments_for_file(file_id: int):
    with db() as conn:
        return conn.execute("""
            SELECT a.*, c.name, c.telegram
            FROM appointments a JOIN clients c ON c.id=a.client_id
            WHERE a.file_id=?
            ORDER BY a.appointment_date, a.appointment_time, a.id
        """, (file_id,)).fetchall()


def update_file_name(file_id: int, name: str):
    with db() as conn:
        conn.execute("UPDATE files SET name=? WHERE id=?", (name.strip(), file_id))


def delete_file(file_id: int):
    with db() as conn:
        conn.execute("DELETE FROM appointments WHERE file_id=?", (file_id,))
        conn.execute("DELETE FROM files WHERE id=?", (file_id,))


def get_client(client_id: int):
    with db() as conn:
        return conn.execute("SELECT * FROM clients WHERE id=?", (client_id,)).fetchone()


def find_or_create_client(name: str, telegram: str) -> int:
    telegram = telegram.strip()
    if telegram and not telegram.startswith("@"):
        telegram = "@" + telegram
    with db() as conn:
        row = conn.execute("SELECT id FROM clients WHERE telegram=?", (telegram,)).fetchone()
        if row:
            conn.execute("UPDATE clients SET name=? WHERE id=?", (name.strip(), row["id"]))
            return row["id"]
        cur = conn.execute(
            "INSERT INTO clients(name, telegram, created_at) VALUES(?,?,?)",
            (name.strip(), telegram, datetime.now().isoformat(timespec="seconds")),
        )
        return cur.lastrowid


def add_appointment(
    client_id: int,
    file_id: int,
    service: str,
    appt_date: str,
    appt_time: str,
    comment: str = "",
    source: str = "",
    contact: str = "",
    pain: str = "",
    paid_status: str = "Не указано",
    paid_amount: str = "",
) -> int:
    with db() as conn:
        cur = conn.execute(
            """
            INSERT INTO appointments(
                client_id, file_id, service, appointment_date, appointment_time,
                comment, source, contact, pain, paid_status, paid_amount, created_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                client_id, file_id, service.strip(), appt_date, appt_time,
                comment.strip(), source.strip(), contact.strip(), pain.strip(),
                paid_status.strip(), paid_amount.strip(), datetime.now().isoformat(timespec="seconds")
            ),
        )
        return cur.lastrowid


def get_appointment(appt_id: int):
    with db() as conn:
        return conn.execute("""
            SELECT a.*, c.name, c.telegram, f.name AS file_name
            FROM appointments a JOIN clients c ON c.id=a.client_id
            JOIN files f ON f.id=a.file_id
            WHERE a.id=?
        """, (appt_id,)).fetchone()


def update_appointment(appt_id: int, field: str, value: str):
    allowed = {
        "appointment_date", "appointment_time", "service", "comment",
        "source", "contact", "pain", "paid_status", "paid_amount"
    }
    if field not in allowed:
        raise ValueError("Invalid field")
    with db() as conn:
        conn.execute(f"UPDATE appointments SET {field}=? WHERE id=?", (value, appt_id))


def delete_appointment(appt_id: int):
    with db() as conn:
        conn.execute("DELETE FROM appointments WHERE id=?", (appt_id,))


def appointments_for_date(appt_date: str, file_id=None):
    with db() as conn:
        if file_id is None:
            return conn.execute("""
                SELECT a.*, c.name, c.telegram, f.name AS file_name
                FROM appointments a JOIN clients c ON c.id=a.client_id
                JOIN files f ON f.id=a.file_id
                WHERE a.appointment_date=?
                ORDER BY a.appointment_time, a.id
            """, (appt_date,)).fetchall()
        return conn.execute("""
            SELECT a.*, c.name, c.telegram, f.name AS file_name
            FROM appointments a JOIN clients c ON c.id=a.client_id
            JOIN files f ON f.id=a.file_id
            WHERE a.appointment_date=? AND a.file_id=?
            ORDER BY a.appointment_time, a.id
        """, (appt_date, file_id)).fetchall()


def all_appointments(file_id=None):
    with db() as conn:
        if file_id is None:
            return conn.execute("""
                SELECT a.*, c.name, c.telegram, f.name AS file_name
                FROM appointments a
                JOIN clients c ON c.id=a.client_id
                JOIN files f ON f.id=a.file_id
                ORDER BY a.appointment_date, a.appointment_time, a.id
            """).fetchall()
        return conn.execute("""
            SELECT a.*, c.name, c.telegram, f.name AS file_name
            FROM appointments a
            JOIN clients c ON c.id=a.client_id
            JOIN files f ON f.id=a.file_id
            WHERE a.file_id=?
            ORDER BY a.appointment_date, a.appointment_time, a.id
        """, (file_id,)).fetchall()


def all_clients():
    with db() as conn:
        return conn.execute("SELECT * FROM clients ORDER BY name").fetchall()

# -------------------- UI --------------------

def bottom_nav():
    kb = InlineKeyboardBuilder()
    kb.button(text="⌂ Главная", callback_data="home")
    kb.button(text="📁 Файлы", callback_data="files")
    kb.button(text="📅 Расписание", callback_data="schedule")
    kb.button(text="👥 Клиенты", callback_data="clients")
    kb.adjust(4)
    return kb.as_markup()


def main_menu():
    kb = InlineKeyboardBuilder()
    kb.button(text="➕ Новая запись", callback_data="add")
    kb.button(text="📅 Расписание", callback_data="schedule")
    kb.button(text="📁 Мои файлы", callback_data="files")
    kb.button(text="👥 Клиенты", callback_data="clients")
    kb.button(text="📊 Excel", callback_data="excel_menu")
    kb.button(text="🔎 Найти файл", callback_data="file_search")
    kb.button(text="☀️ Сегодня", callback_data=f"viewdate:{date.today().isoformat()}")
    kb.button(text="🌙 Завтра", callback_data=f"viewdate:{(date.today() + timedelta(days=1)).isoformat()}")
    kb.adjust(2, 2, 2, 2)
    return kb.as_markup()


def file_keyboard(prefix="file"):
    kb = InlineKeyboardBuilder()
    for f in all_files():
        kb.button(text=f"📁 {f['name']} · {f['appointments_count']}", callback_data=f"{prefix}:{f['id']}")
    kb.button(text="➕ Новый файл", callback_data="new_file")
    kb.button(text="🔎 Найти по названию", callback_data="file_search")
    kb.button(text="⌂ Главная", callback_data="home")
    kb.button(text="📅 Расписание", callback_data="schedule")
    kb.button(text="👥 Клиенты", callback_data="clients")
    kb.adjust(1, 2, 3)
    return kb.as_markup()


def file_manage_keyboard(file_id: int):
    kb = InlineKeyboardBuilder()
    kb.button(text="➕ Добавить запись", callback_data=f"addfile:{file_id}")
    kb.button(text="📅 Открыть расписание", callback_data=f"file_schedule:{file_id}")
    kb.button(text="📊 Скачать Excel", callback_data=f"excel:{file_id}")
    kb.button(text="✏️ Переименовать", callback_data=f"rename_file:{file_id}")
    kb.button(text="🗑 Удалить файл", callback_data=f"delete_file:{file_id}")
    kb.button(text="‹ К файлам", callback_data="files")
    kb.button(text="⌂ Главная", callback_data="home")
    kb.adjust(1)
    return kb.as_markup()


def cancel_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="✕ Отмена", callback_data="cancel")
    return kb.as_markup()


def skip_kb(callback_data: str = "skip"):
    kb = InlineKeyboardBuilder()
    kb.button(text="Пропустить", callback_data=callback_data)
    kb.button(text="✕ Отмена", callback_data="cancel")
    kb.adjust(1)
    return kb.as_markup()


def source_keyboard(prefix: str = "source"):
    kb = InlineKeyboardBuilder()
    kb.button(text="MAX", callback_data=f"{prefix}:MAX")
    kb.button(text="Авито", callback_data=f"{prefix}:Авито")
    kb.button(text="ВКонтакте", callback_data=f"{prefix}:ВКонтакте")
    kb.button(text="Telegram", callback_data=f"{prefix}:Telegram")
    kb.button(text="Другое", callback_data=f"{prefix}:Другое")
    kb.button(text="✕ Отмена", callback_data="cancel")
    kb.adjust(2, 2, 1, 1)
    return kb.as_markup()


def payment_keyboard(prefix: str = "paid"):
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Оплачено", callback_data=f"{prefix}:Оплачено")
    kb.button(text="❌ Не оплачено", callback_data=f"{prefix}:Не оплачено")
    kb.button(text="💳 Частично", callback_data=f"{prefix}:Частично")
    kb.button(text="➖ Не указано", callback_data=f"{prefix}:Не указано")
    kb.button(text="✕ Отмена", callback_data="cancel")
    kb.adjust(2, 2, 1)
    return kb.as_markup()


def date_keyboard(prefix: str = "date"):
    today = date.today()
    kb = InlineKeyboardBuilder()
    labels = ["Сегодня", "Завтра", "Послезавтра"]
    for i in range(3):
        d = today + timedelta(days=i)
        kb.button(text=f"{labels[i]} · {d:%d.%m}", callback_data=f"{prefix}:{d.isoformat()}")
    for i in range(3, 17):
        d = today + timedelta(days=i)
        kb.button(text=d.strftime("%d.%m"), callback_data=f"{prefix}:{d.isoformat()}")
    kb.button(text="✕ Отмена", callback_data="cancel")
    kb.adjust(1, 1, 1, 4, 4, 4, 1)
    return kb.as_markup()


def time_keyboard(prefix: str = "time"):
    kb = InlineKeyboardBuilder()
    for h in range(8, 21):
        for m in (0, 30):
            kb.button(text=f"{h:02d}:{m:02d}", callback_data=f"{prefix}:{h:02d}:{m:02d}")
    kb.button(text="✕ Отмена", callback_data="cancel")
    kb.adjust(4)
    return kb.as_markup()


def appointment_kb(appt_id: int):
    kb = InlineKeyboardBuilder()
    kb.button(text="📅 Дата", callback_data=f"edit_date:{appt_id}")
    kb.button(text="🕐 Время", callback_data=f"edit_time:{appt_id}")
    kb.button(text="💅 Услуга", callback_data=f"edit_service:{appt_id}")
    kb.button(text="📍 Источник", callback_data=f"edit_source:{appt_id}")
    kb.button(text="💬 Контакт", callback_data=f"edit_contact:{appt_id}")
    kb.button(text="🎯 Запрос", callback_data=f"edit_pain:{appt_id}")
    kb.button(text="💰 Оплата", callback_data=f"edit_paid:{appt_id}")
    kb.button(text="📝 Комментарий", callback_data=f"edit_comment:{appt_id}")
    kb.button(text="🗑 Удалить", callback_data=f"delete:{appt_id}")
    kb.button(text="‹ К расписанию", callback_data="schedule")
    kb.adjust(2, 2, 2, 2, 2)
    return kb.as_markup()


def format_appt(a):
    lines = [
        f"📁 {a['file_name']}" if "file_name" in a.keys() else "",
        f"🕐 <b>{a['appointment_time']}</b> — {a['name']}",
        f"💅 {a['service']}",
        f"💬 {a['telegram']}",
    ]
    if a["contact"]:
        lines.append(f"📞 {a['contact']}")
    if a["source"]:
        lines.append(f"📍 {a['source']}")
    if a["pain"]:
        lines.append(f"🎯 {a['pain']}")
    paid = a["paid_status"] or "Не указано"
    if a["paid_amount"]:
        paid += f" — {a['paid_amount']} ₽"
    lines.append(f"💰 {paid}")
    if a["comment"]:
        lines.append(f"📝 {a['comment']}")
    return "\n".join(line for line in lines if line)


def format_day_title(d: str) -> str:
    value = datetime.strptime(d, "%Y-%m-%d").date()
    weekday = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"][value.weekday()]
    return f"{weekday}, {value:%d.%m.%Y}"

# -------------------- FSM --------------------

class AddAppointment(StatesGroup):
    file = State()
    name = State()
    telegram = State()
    contact = State()
    source = State()
    custom_source = State()
    service = State()
    pain = State()
    date = State()
    time = State()
    payment = State()
    paid_amount = State()
    comment = State()


class EditText(StatesGroup):
    value = State()


class EditPaid(StatesGroup):
    amount = State()


class FileText(StatesGroup):
    name = State()


class FileSearch(StatesGroup):
    query = State()

# -------------------- Handlers --------------------

@router.message(CommandStart())
async def start(message: Message):
    await message.answer(
        "<b>📋 Клиенты & расписание</b>\n\n"
        "Добавляйте клиентов, записи и выгружайте всё в Excel.",
        reply_markup=main_menu(),
    )


@router.callback_query(F.data == "cancel")
async def cancel(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await call.message.edit_text("<b>Главное меню</b>", reply_markup=main_menu())
    await call.answer()


@router.callback_query(F.data == "add")
async def add_start(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(AddAppointment.file)
    await call.message.edit_text("<b>➕ Новая запись</b>\n\n📁 Выберите файл, в который добавить запись:", reply_markup=file_keyboard("addfile"))
    await call.answer()


@router.callback_query(F.data.startswith("addfile:"))
async def add_file_selected(call: CallbackQuery, state: FSMContext):
    file_id = int(call.data.split(":")[1])
    f = get_file(file_id)
    if not f:
        await call.answer("Файл не найден", show_alert=True)
        return
    await state.update_data(file_id=file_id, file_name=f["name"])
    await state.set_state(AddAppointment.name)
    await call.message.edit_text(f"<b>📁 {f['name']}</b>\n\n👤 Введите имя клиента:", reply_markup=cancel_kb())
    await call.answer()


@router.message(AddAppointment.name)
async def add_name(message: Message, state: FSMContext):
    if not message.text or not message.text.strip():
        await message.answer("Введите имя клиента.", reply_markup=cancel_kb())
        return
    await state.update_data(name=message.text.strip())
    await state.set_state(AddAppointment.telegram)
    await message.answer("💬 Telegram-ник клиента:\nнапример, <code>@anna</code>", reply_markup=cancel_kb())


@router.message(AddAppointment.telegram)
async def add_telegram(message: Message, state: FSMContext):
    tg = (message.text or "").strip()
    if not tg:
        await message.answer("Введите Telegram-ник.", reply_markup=cancel_kb())
        return
    await state.update_data(telegram=tg)
    await state.set_state(AddAppointment.contact)
    await message.answer("📞 Контакт клиента\n\nТелефон, ссылка VK или другой контакт. Можно пропустить.", reply_markup=skip_kb("skip_contact"))


@router.callback_query(AddAppointment.contact, F.data == "skip_contact")
async def skip_contact(call: CallbackQuery, state: FSMContext):
    await state.update_data(contact="")
    await state.set_state(AddAppointment.source)
    await call.message.edit_text("📍 Откуда пришёл клиент?", reply_markup=source_keyboard("addsource"))
    await call.answer()


@router.message(AddAppointment.contact)
async def add_contact(message: Message, state: FSMContext):
    await state.update_data(contact=(message.text or "").strip())
    await state.set_state(AddAppointment.source)
    await message.answer("📍 Откуда пришёл клиент?", reply_markup=source_keyboard("addsource"))


@router.callback_query(AddAppointment.source, F.data.startswith("addsource:"))
async def add_source(call: CallbackQuery, state: FSMContext):
    source = call.data.split(":", 1)[1]
    if source == "Другое":
        await state.set_state(AddAppointment.custom_source)
        await call.message.edit_text("📍 Введите источник клиента:", reply_markup=cancel_kb())
    else:
        await state.update_data(source=source)
        await state.set_state(AddAppointment.service)
        await call.message.edit_text("💅 Какая услуга?", reply_markup=cancel_kb())
    await call.answer()


@router.message(AddAppointment.custom_source)
async def add_custom_source(message: Message, state: FSMContext):
    value = (message.text or "").strip()
    if not value:
        await message.answer("Введите источник или нажмите «Отмена».", reply_markup=cancel_kb())
        return
    await state.update_data(source=value)
    await state.set_state(AddAppointment.service)
    await message.answer("💅 Какая услуга?", reply_markup=cancel_kb())


@router.message(AddAppointment.service)
async def add_service(message: Message, state: FSMContext):
    if not message.text or not message.text.strip():
        await message.answer("Введите название услуги.", reply_markup=cancel_kb())
        return
    await state.update_data(service=message.text.strip())
    await state.set_state(AddAppointment.pain)
    await message.answer("🎯 Краткая боль / запрос клиента?\n\nНапример: «Деньги», «Нужна консультация». Можно пропустить.", reply_markup=skip_kb("skip_pain"))


@router.callback_query(AddAppointment.pain, F.data == "skip_pain")
async def skip_pain(call: CallbackQuery, state: FSMContext):
    await state.update_data(pain="")
    await state.set_state(AddAppointment.date)
    await call.message.edit_text("📅 Выберите дату:", reply_markup=date_keyboard("adddate"))
    await call.answer()


@router.message(AddAppointment.pain)
async def add_pain(message: Message, state: FSMContext):
    await state.update_data(pain=(message.text or "").strip())
    await state.set_state(AddAppointment.date)
    await message.answer("📅 Выберите дату:", reply_markup=date_keyboard("adddate"))


@router.callback_query(AddAppointment.date, F.data.startswith("adddate:"))
async def add_date(call: CallbackQuery, state: FSMContext):
    value = call.data.split(":", 1)[1]
    await state.update_data(date=value)
    await state.set_state(AddAppointment.time)
    await call.message.edit_text("🕐 Выберите время:", reply_markup=time_keyboard("addtime"))
    await call.answer()


@router.callback_query(AddAppointment.time, F.data.startswith("addtime:"))
async def add_time(call: CallbackQuery, state: FSMContext):
    _, h, m = call.data.split(":")
    await state.update_data(time=f"{h}:{m}")
    await state.set_state(AddAppointment.payment)
    await call.message.edit_text("💰 Статус оплаты:", reply_markup=payment_keyboard("addpaid"))
    await call.answer()


def comment_keyboard():
    kb = InlineKeyboardBuilder()
    kb.button(text="Без комментария", callback_data="no_comment")
    kb.button(text="✕ Отмена", callback_data="cancel")
    kb.adjust(1)
    return kb.as_markup()


@router.callback_query(AddAppointment.payment, F.data.startswith("addpaid:"))
async def add_payment(call: CallbackQuery, state: FSMContext):
    status = call.data.split(":", 1)[1]
    await state.update_data(paid_status=status)
    if status == "Частично":
        await state.set_state(AddAppointment.paid_amount)
        await call.message.edit_text("💳 Сколько оплачено?\n\nВведите сумму в рублях, например: 1500", reply_markup=cancel_kb())
    else:
        await state.set_state(AddAppointment.comment)
        await call.message.edit_text("📝 Комментарий?\n\nНапишите текст или нажмите «Без комментария». ", reply_markup=comment_keyboard())
    await call.answer()


@router.message(AddAppointment.paid_amount)
async def add_paid_amount(message: Message, state: FSMContext):
    amount = (message.text or "").strip().replace(",", ".")
    if not amount:
        await message.answer("Введите сумму.", reply_markup=cancel_kb())
        return
    await state.update_data(paid_amount=amount)
    await state.set_state(AddAppointment.comment)
    await message.answer("📝 Комментарий?\n\nНапишите текст или нажмите «Без комментария».", reply_markup=comment_keyboard())


@router.callback_query(AddAppointment.comment, F.data == "no_comment")
async def no_comment(call: CallbackQuery, state: FSMContext):
    await save_appointment(call.message, state, "", edit=True)
    await call.answer()


@router.message(AddAppointment.comment)
async def add_comment(message: Message, state: FSMContext):
    comment = (message.text or "").strip()
    if comment.lower() in {"нет", "-", "нету"}:
        comment = ""
    await save_appointment(message, state, comment, edit=False)


async def save_appointment(target: Message, state: FSMContext, comment: str, edit: bool):
    data = await state.get_data()
    client_id = find_or_create_client(data["name"], data["telegram"])
    appt_id = add_appointment(
        client_id,
        int(data.get("file_id", 1)),
        data["service"],
        data["date"],
        data["time"],
        comment,
        data.get("source", ""),
        data.get("contact", ""),
        data.get("pain", ""),
        data.get("paid_status", "Не указано"),
        data.get("paid_amount", ""),
    )
    await state.clear()
    a = get_appointment(appt_id)
    text = "<b>✅ Запись сохранена</b>\n\n" + format_appt(a)
    if edit:
        await target.edit_text(text, reply_markup=appointment_kb(appt_id))
    else:
        await target.answer(text, reply_markup=appointment_kb(appt_id))


@router.callback_query(F.data == "schedule")
async def schedule(call: CallbackQuery):
    await call.message.edit_text("<b>📅 Расписание</b>\n\nСначала выберите файл:", reply_markup=file_keyboard("schedfile"))
    await call.answer()


@router.callback_query(F.data.startswith("schedfile:"))
async def schedule_file(call: CallbackQuery):
    file_id = int(call.data.split(":")[1])
    f = get_file(file_id)
    if not f:
        await call.answer("Файл не найден", show_alert=True)
        return
    await call.message.edit_text(f"<b>📅 {f['name']}</b>\n\nВыберите день:", reply_markup=date_keyboard(f"viewdate:{file_id}"))
    await call.answer()


@router.callback_query(F.data.startswith("viewdate:"))
async def view_date(call: CallbackQuery):
    parts = call.data.split(":")
    if len(parts) == 2:
        file_id = None
        d = parts[1]
    else:
        file_id = int(parts[1])
        d = parts[2]
    rows = appointments_for_date(d, file_id)
    title = format_day_title(d)
    file_name = get_file(file_id)["name"] if file_id and get_file(file_id) else "Все файлы"
    count = len(rows)
    if not rows:
        text = f"<b>📅 {title}</b>\n📁 {file_name}\n\nСвободный день ✨"
    else:
        text = f"<b>📅 {title}</b>\n📁 {file_name} · {count} запис." + ("\n\n" + "\n\n".join(format_appt(a) for a in rows))
    kb = InlineKeyboardBuilder()
    for a in rows:
        kb.button(text=f"{a['appointment_time']} · {a['name']}", callback_data=f"open:{a['id']}")
    if file_id:
        kb.button(text="➕ Добавить сюда", callback_data=f"addfile:{file_id}")
        kb.button(text="‹ Другой день", callback_data=f"file_schedule:{file_id}")
        kb.button(text="📁 Файл", callback_data=f"file:{file_id}")
    else:
        kb.button(text="➕ Добавить запись", callback_data="add")
        kb.button(text="‹ Выбрать другой день", callback_data="schedule")
    kb.button(text="⌂ Главное меню", callback_data="home")
    kb.adjust(1)
    await call.message.edit_text(text, reply_markup=kb.as_markup())
    await call.answer()


@router.callback_query(F.data.startswith("file_schedule:"))
async def file_schedule(call: CallbackQuery):
    file_id = int(call.data.split(":")[1])
    f = get_file(file_id)
    if not f:
        await call.answer("Файл не найден", show_alert=True)
        return
    await call.message.edit_text(f"<b>📅 {f['name']}</b>\n\nВыберите день:", reply_markup=date_keyboard(f"viewdate:{file_id}"))
    await call.answer()


@router.callback_query(F.data.startswith("open:"))
async def open_appt(call: CallbackQuery):
    appt_id = int(call.data.split(":")[1])
    a = get_appointment(appt_id)
    if not a:
        await call.answer("Запись не найдена", show_alert=True)
        return
    d = format_day_title(a["appointment_date"])
    text = f"<b>📋 Запись</b>\n\n📅 {d}\n\n{format_appt(a)}"
    await call.message.edit_text(text, reply_markup=appointment_kb(appt_id))
    await call.answer()


@router.callback_query(F.data.startswith("delete:"))
async def delete_appt(call: CallbackQuery):
    appt_id = int(call.data.split(":")[1])
    delete_appointment(appt_id)
    await call.message.edit_text("<b>🗑 Запись удалена</b>", reply_markup=main_menu())
    await call.answer()


@router.callback_query(F.data.startswith("edit_date:"))
async def edit_date(call: CallbackQuery):
    appt_id = int(call.data.split(":")[1])
    await call.message.edit_text("<b>📅 Новая дата</b>", reply_markup=date_keyboard(f"setdate:{appt_id}"))
    await call.answer()


@router.callback_query(F.data.startswith("setdate:"))
async def set_date(call: CallbackQuery):
    _, appt_id, d = call.data.split(":")
    appt_id = int(appt_id)
    update_appointment(appt_id, "appointment_date", d)
    a = get_appointment(appt_id)
    await call.message.edit_text("<b>✅ Дата изменена</b>\n\n" + format_appt(a), reply_markup=appointment_kb(appt_id))
    await call.answer()


@router.callback_query(F.data.startswith("edit_time:"))
async def edit_time(call: CallbackQuery):
    appt_id = int(call.data.split(":")[1])
    await call.message.edit_text("<b>🕐 Новое время</b>", reply_markup=time_keyboard(f"settime:{appt_id}"))
    await call.answer()


@router.callback_query(F.data.startswith("settime:"))
async def set_time(call: CallbackQuery):
    parts = call.data.split(":")
    appt_id = int(parts[1])
    value = f"{parts[2]}:{parts[3]}"
    update_appointment(appt_id, "appointment_time", value)
    a = get_appointment(appt_id)
    await call.message.edit_text("<b>✅ Время изменено</b>\n\n" + format_appt(a), reply_markup=appointment_kb(appt_id))
    await call.answer()


async def edit_text_start(call: CallbackQuery, state: FSMContext, field: str, prompt: str):
    appt_id = int(call.data.split(":")[1])
    await state.set_state(EditText.value)
    await state.update_data(appt_id=appt_id, field=field)
    await call.message.edit_text(prompt, reply_markup=cancel_kb())
    await call.answer()


@router.callback_query(F.data.startswith("edit_service:"))
async def edit_service(call: CallbackQuery, state: FSMContext):
    await edit_text_start(call, state, "service", "💅 Введите новую услугу:")


@router.callback_query(F.data.startswith("edit_comment:"))
async def edit_comment(call: CallbackQuery, state: FSMContext):
    await edit_text_start(call, state, "comment", "📝 Введите новый комментарий:")


@router.callback_query(F.data.startswith("edit_contact:"))
async def edit_contact(call: CallbackQuery, state: FSMContext):
    await edit_text_start(call, state, "contact", "💬 Введите новый контакт:")


@router.callback_query(F.data.startswith("edit_pain:"))
async def edit_pain(call: CallbackQuery, state: FSMContext):
    await edit_text_start(call, state, "pain", "🎯 Введите новую боль / запрос клиента:")


@router.callback_query(F.data.startswith("edit_source:"))
async def edit_source(call: CallbackQuery, state: FSMContext):
    appt_id = int(call.data.split(":")[1])
    await state.clear()
    await call.message.edit_text("📍 Выберите новый источник:", reply_markup=source_keyboard(f"setsource:{appt_id}"))
    await call.answer()


@router.callback_query(F.data.startswith("setsource:"))
async def set_source(call: CallbackQuery, state: FSMContext):
    parts = call.data.split(":", 2)
    appt_id = int(parts[1])
    source = parts[2]
    if source == "Другое":
        await state.set_state(EditText.value)
        await state.update_data(appt_id=appt_id, field="source")
        await call.message.edit_text("📍 Введите новый источник:", reply_markup=cancel_kb())
        await call.answer()
        return
    update_appointment(appt_id, "source", source)
    a = get_appointment(appt_id)
    await call.message.edit_text("<b>✅ Источник изменён</b>\n\n" + format_appt(a), reply_markup=appointment_kb(appt_id))
    await call.answer()


@router.callback_query(F.data.startswith("edit_paid:"))
async def edit_paid(call: CallbackQuery):
    appt_id = int(call.data.split(":")[1])
    await call.message.edit_text("💰 Новый статус оплаты:", reply_markup=payment_keyboard(f"setpaid:{appt_id}"))
    await call.answer()


@router.callback_query(F.data.startswith("setpaid:"))
async def set_paid(call: CallbackQuery, state: FSMContext):
    parts = call.data.split(":", 2)
    appt_id = int(parts[1])
    status = parts[2]
    update_appointment(appt_id, "paid_status", status)
    if status == "Частично":
        await state.set_state(EditPaid.amount)
        await state.update_data(appt_id=appt_id)
        await call.message.edit_text("💳 Введите оплачeнную сумму в рублях:", reply_markup=cancel_kb())
    else:
        update_appointment(appt_id, "paid_amount", "")
        a = get_appointment(appt_id)
        await call.message.edit_text("<b>✅ Оплата изменена</b>\n\n" + format_appt(a), reply_markup=appointment_kb(appt_id))
    await call.answer()


@router.message(EditPaid.amount)
async def edit_paid_amount(message: Message, state: FSMContext):
    amount = (message.text or "").strip().replace(",", ".")
    if not amount:
        await message.answer("Введите сумму.", reply_markup=cancel_kb())
        return
    data = await state.get_data()
    appt_id = data["appt_id"]
    update_appointment(appt_id, "paid_amount", amount)
    a = get_appointment(appt_id)
    await state.clear()
    await message.answer("<b>✅ Оплата изменена</b>\n\n" + format_appt(a), reply_markup=appointment_kb(appt_id))


@router.message(EditText.value)
async def edit_text_value(message: Message, state: FSMContext):
    data = await state.get_data()
    value = (message.text or "").strip()
    field = data["field"]
    if field == "comment" and value.lower() in {"нет", "-", "нету"}:
        value = ""
    if not value and field in {"service", "source"}:
        await message.answer("Это поле не может быть пустым.", reply_markup=cancel_kb())
        return
    update_appointment(data["appt_id"], field, value)
    a = get_appointment(data["appt_id"])
    await state.clear()
    await message.answer("<b>✅ Изменено</b>\n\n" + format_appt(a), reply_markup=appointment_kb(data["appt_id"]))


@router.callback_query(F.data == "files")
async def files_menu(call: CallbackQuery):
    await call.message.edit_text("<b>📁 Мои файлы</b>\n\nКаждый файл — отдельное расписание. Выберите нужный:", reply_markup=file_keyboard("file"))
    await call.answer()


@router.callback_query(F.data == "new_file")
async def new_file_start(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(FileText.name)
    await call.message.edit_text("<b>📁 Новый файл</b>\n\nВведите название, например:\n<code>Ноябрь 2026</code>", reply_markup=cancel_kb())
    await call.answer()


@router.message(FileText.name)
async def new_file_name(message: Message, state: FSMContext):
    name = (message.text or "").strip()
    if not name:
        await message.answer("Введите название файла.", reply_markup=cancel_kb())
        return
    data = await state.get_data()
    rename_file_id = data.get("rename_file_id")
    existing = find_file_by_name(name)
    if existing and int(existing["id"]) != int(rename_file_id or 0):
        await message.answer("Файл с таким названием уже есть. Введите другое название.", reply_markup=cancel_kb())
        return
    if rename_file_id:
        update_file_name(int(rename_file_id), name)
        await state.clear()
        await message.answer(f"<b>✅ Файл переименован</b>\n\n📁 {name}", reply_markup=file_manage_keyboard(int(rename_file_id)))
        return
    file_id = create_file(name)
    await state.clear()
    await message.answer(f"<b>✅ Файл создан</b>\n\n📁 {name}\n\nТеперь можно добавлять в него записи.", reply_markup=file_manage_keyboard(file_id))


@router.callback_query(F.data.startswith("file:"))
async def file_card(call: CallbackQuery):
    file_id = int(call.data.split(":")[1])
    f = get_file(file_id)
    if not f:
        await call.answer("Файл не найден", show_alert=True)
        return
    await call.message.edit_text(
        f"<b>📁 {f['name']}</b>\n\n📋 Записей: {f['appointments_count']}",
        reply_markup=file_manage_keyboard(file_id),
    )
    await call.answer()


@router.callback_query(F.data == "file_search")
async def file_search_start(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await state.set_state(FileSearch.query)
    await call.message.edit_text("<b>🔎 Найти файл</b>\n\nВведите название или его часть:", reply_markup=cancel_kb())
    await call.answer()


@router.message(FileSearch.query)
async def file_search_query(message: Message, state: FSMContext):
    query = (message.text or "").strip().lower()
    if not query:
        await message.answer("Введите название.", reply_markup=cancel_kb())
        return
    with db() as conn:
        rows = conn.execute("""
            SELECT f.*, COUNT(a.id) AS appointments_count
            FROM files f LEFT JOIN appointments a ON a.file_id=f.id
            WHERE lower(f.name) LIKE ?
            GROUP BY f.id ORDER BY f.name COLLATE NOCASE
        """, (f"%{query}%",)).fetchall()
    await state.clear()
    if not rows:
        await message.answer("Ничего не найдено.", reply_markup=file_keyboard("file"))
        return
    kb = InlineKeyboardBuilder()
    for f in rows:
        kb.button(text=f"📁 {f['name']} · {f['appointments_count']}", callback_data=f"file:{f['id']}")
    kb.button(text="🔎 Новый поиск", callback_data="file_search")
    kb.button(text="‹ К файлам", callback_data="files")
    kb.adjust(1)
    await message.answer(f"<b>🔎 Найдено: {len(rows)}</b>", reply_markup=kb.as_markup())


@router.callback_query(F.data.startswith("rename_file:"))
async def rename_file_start(call: CallbackQuery, state: FSMContext):
    file_id = int(call.data.split(":")[1])
    if not get_file(file_id):
        await call.answer("Файл не найден", show_alert=True)
        return
    await state.clear()
    await state.set_state(FileText.name)
    await state.update_data(rename_file_id=file_id)
    await call.message.edit_text("<b>✏️ Переименовать файл</b>\n\nВведите новое название:", reply_markup=cancel_kb())
    await call.answer()


@router.callback_query(F.data.startswith("delete_file:"))
async def delete_file_confirm(call: CallbackQuery):
    file_id = int(call.data.split(":")[1])
    f = get_file(file_id)
    if not f:
        await call.answer("Файл не найден", show_alert=True)
        return
    if f["name"] == "Основной":
        await call.answer("Основной файл нельзя удалить", show_alert=True)
        return
    kb = InlineKeyboardBuilder()
    kb.button(text="Да, удалить", callback_data=f"confirm_delete_file:{file_id}")
    kb.button(text="Отмена", callback_data=f"file:{file_id}")
    kb.adjust(1)
    await call.message.edit_text(f"<b>🗑 Удалить «{f['name']}»?</b>\n\nВсе записи внутри этого файла тоже будут удалены.", reply_markup=kb.as_markup())
    await call.answer()


@router.callback_query(F.data.startswith("confirm_delete_file:"))
async def delete_file_confirmed(call: CallbackQuery):
    file_id = int(call.data.split(":")[1])
    f = get_file(file_id)
    if not f:
        await call.answer("Файл уже удалён", show_alert=True)
        return
    delete_file(file_id)
    await call.message.edit_text("<b>🗑 Файл удалён</b>", reply_markup=file_keyboard("file"))
    await call.answer()


@router.callback_query(F.data == "clients")
async def clients(call: CallbackQuery):
    rows = all_clients()
    if not rows:
        await call.message.edit_text("<b>👥 Клиенты</b>\n\nПока клиентов нет.", reply_markup=main_menu())
        await call.answer()
        return
    kb = InlineKeyboardBuilder()
    for c in rows:
        kb.button(text=f"👤 {c['name']}", callback_data=f"client:{c['id']}")
    kb.button(text="⌂ Главная", callback_data="home")
    kb.button(text="📁 Файлы", callback_data="files")
    kb.button(text="📅 Расписание", callback_data="schedule")
    kb.adjust(2, 2, 2)
    await call.message.edit_text("<b>👥 Клиенты</b>\n\nВыберите клиента:", reply_markup=kb.as_markup())
    await call.answer()


@router.callback_query(F.data.startswith("client:"))
async def client_card(call: CallbackQuery):
    cid = int(call.data.split(":")[1])
    c = get_client(cid)
    if not c:
        await call.answer("Клиент не найден", show_alert=True)
        return
    with db() as conn:
        count = conn.execute("SELECT COUNT(*) FROM appointments WHERE client_id=?", (cid,)).fetchone()[0]
    kb = InlineKeyboardBuilder()
    kb.button(text="➕ Новая запись", callback_data="add")
    kb.button(text="‹ К клиентам", callback_data="clients")
    kb.button(text="⌂ Главное меню", callback_data="home")
    kb.adjust(1)
    await call.message.edit_text(
        f"<b>👤 {c['name']}</b>\n\n💬 {c['telegram']}\n📅 Записей: {count}",
        reply_markup=kb.as_markup(),
    )
    await call.answer()


@router.callback_query(F.data == "excel_menu")
async def excel_menu(call: CallbackQuery):
    await call.message.edit_text("<b>📊 Excel</b>\n\nВыберите файл для выгрузки:", reply_markup=file_keyboard("excelfile"))
    await call.answer()


@router.callback_query(F.data.startswith("excelfile:"))
async def excel_file_selected(call: CallbackQuery):
    file_id = int(call.data.split(":")[1])
    await excel_export(call, file_id)


@router.callback_query(F.data.startswith("excel:"))
async def excel_file_direct(call: CallbackQuery):
    file_id = int(call.data.split(":")[1])
    await excel_export(call, file_id)


async def excel_export(call: CallbackQuery, file_id=None):
    if file_id is None:
        parts = call.data.split(":")
        file_id = int(parts[1]) if len(parts) > 1 and parts[1] else None
    rows = all_appointments(file_id)
    if not rows:
        await call.answer("В этом файле пока нет записей", show_alert=True)
        return
    selected_file = get_file(file_id) if file_id else None

    wb = Workbook()
    ws = wb.active
    ws.title = "Расписание"
    if selected_file:
        ws["A1"] = f"{selected_file['name']} · КЛИЕНТЫ И РАСПИСАНИЕ"
        ws.merge_cells("A1:H1")
        ws["A1"].font = Font(bold=True, color="FFFFFF", size=16)
        ws["A1"].fill = PatternFill("solid", fgColor="6F4E37")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30
        header_row = 3
    else:
        header_row = 1

    headers = [
        "Дата", "Время", "Клиент", "Контакт", "Источник",
        "Краткая боль / запрос", "Оплачено", "Комментарий"
    ]
    if selected_file:
        ws.append([None] * len(headers))
        ws.append(headers)
    else:
        ws.append(headers)
    header_fill = PatternFill("solid", fgColor="6F4E37")
    for cell in ws[header_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="D9D9D9")
    for r in rows:
        d = datetime.strptime(r["appointment_date"], "%Y-%m-%d").strftime("%d.%m.%Y")
        paid = r["paid_status"] or "Не указано"
        if r["paid_amount"]:
            paid += f" ({r['paid_amount']} ₽)"
        ws.append([
            d, r["appointment_time"], r["name"], r["contact"], r["source"],
            r["pain"], paid, r["comment"]
        ])
    for row in ws.iter_rows(min_row=header_row + 1):
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    widths = [14, 10, 22, 28, 18, 35, 20, 40]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = ws.dimensions
    if not selected_file:
        ws.row_dimensions[1].height = 32

    clients_ws = wb.create_sheet("Клиенты")
    clients_ws.append(["Имя", "Telegram"])
    for cell in clients_ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    for c in all_clients():
        clients_ws.append([c["name"], c["telegram"]])
    clients_ws.column_dimensions["A"].width = 25
    clients_ws.column_dimensions["B"].width = 30
    clients_ws.freeze_panes = "A2"

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    safe_name = (selected_file["name"] if selected_file else "Все_файлы").replace("/", "-")
    filename = f"{safe_name}_{date.today():%Y-%m-%d}.xlsx"
    await call.message.answer_document(
        BufferedInputFile(bio.read(), filename=filename),
        caption="📊 <b>Excel готов</b>\nВнутри: расписание и список клиентов.",
    )
    await call.answer()


@router.callback_query(F.data == "home")
async def home(call: CallbackQuery):
    await call.message.edit_text("<b>Главное меню</b>\n\nВыберите действие:", reply_markup=main_menu())
    await call.answer()


async def main():
    init_db()
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()
    dp.include_router(router)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

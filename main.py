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
            CREATE TABLE IF NOT EXISTS appointments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER NOT NULL,
                service TEXT NOT NULL,
                appointment_date TEXT NOT NULL,
                appointment_time TEXT NOT NULL,
                comment TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE CASCADE
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_appointments_date ON appointments(appointment_date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_clients_telegram ON clients(telegram)")


def get_client(client_id: int):
    with db() as conn:
        return conn.execute("SELECT * FROM clients WHERE id=?", (client_id,)).fetchone()


def find_or_create_client(name: str, telegram: str) -> int:
    telegram = telegram.strip()
    if not telegram.startswith("@"):
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


def add_appointment(client_id: int, service: str, appt_date: str, appt_time: str, comment: str = "") -> int:
    with db() as conn:
        cur = conn.execute(
            "INSERT INTO appointments(client_id,service,appointment_date,appointment_time,comment,created_at) VALUES(?,?,?,?,?,?)",
            (client_id, service.strip(), appt_date, appt_time, comment.strip(), datetime.now().isoformat(timespec="seconds")),
        )
        return cur.lastrowid


def get_appointment(appt_id: int):
    with db() as conn:
        return conn.execute("""
            SELECT a.*, c.name, c.telegram
            FROM appointments a JOIN clients c ON c.id=a.client_id
            WHERE a.id=?
        """, (appt_id,)).fetchone()


def update_appointment(appt_id: int, field: str, value: str):
    allowed = {"appointment_date", "appointment_time", "service", "comment"}
    if field not in allowed:
        raise ValueError("Invalid field")
    with db() as conn:
        conn.execute(f"UPDATE appointments SET {field}=? WHERE id=?", (value, appt_id))


def delete_appointment(appt_id: int):
    with db() as conn:
        conn.execute("DELETE FROM appointments WHERE id=?", (appt_id,))


def appointments_for_date(appt_date: str):
    with db() as conn:
        return conn.execute("""
            SELECT a.*, c.name, c.telegram
            FROM appointments a JOIN clients c ON c.id=a.client_id
            WHERE a.appointment_date=?
            ORDER BY a.appointment_time, a.id
        """, (appt_date,)).fetchall()


def all_appointments():
    with db() as conn:
        return conn.execute("""
            SELECT a.*, c.name, c.telegram
            FROM appointments a JOIN clients c ON c.id=a.client_id
            ORDER BY a.appointment_date, a.appointment_time, a.id
        """).fetchall()


def all_clients():
    with db() as conn:
        return conn.execute("SELECT * FROM clients ORDER BY name").fetchall()


def count_appointments(appt_date: str) -> int:
    with db() as conn:
        return conn.execute("SELECT COUNT(*) FROM appointments WHERE appointment_date=?", (appt_date,)).fetchone()[0]

# -------------------- UI --------------------

def main_menu():
    kb = InlineKeyboardBuilder()
    kb.button(text="➕ Новая запись", callback_data="add")
    kb.button(text="📅 Расписание", callback_data="schedule")
    kb.button(text="👥 Клиенты", callback_data="clients")
    kb.button(text="📊 Excel", callback_data="excel")
    kb.button(text="☀️ Сегодня", callback_data=f"viewdate:{date.today().isoformat()}")
    kb.button(text="🌙 Завтра", callback_data=f"viewdate:{(date.today() + timedelta(days=1)).isoformat()}")
    kb.adjust(2, 2, 2)
    return kb.as_markup()


def cancel_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="✕ Отмена", callback_data="cancel")
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
    kb.button(text="📝 Комментарий", callback_data=f"edit_comment:{appt_id}")
    kb.button(text="🗑 Удалить", callback_data=f"delete:{appt_id}")
    kb.button(text="‹ К расписанию", callback_data="schedule")
    kb.adjust(2, 2, 2)
    return kb.as_markup()


def format_appt(a):
    comment = f"\n📝 {a['comment']}" if a['comment'] else ""
    return f"🕐 <b>{a['appointment_time']}</b> — {a['name']}\n💅 {a['service']}\n💬 {a['telegram']}{comment}"


def format_day_title(d: str) -> str:
    value = datetime.strptime(d, "%Y-%m-%d").date()
    weekday = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"][value.weekday()]
    return f"{weekday}, {value:%d.%m.%Y}"

# -------------------- FSM --------------------

class AddAppointment(StatesGroup):
    name = State()
    telegram = State()
    service = State()
    date = State()
    time = State()
    comment = State()


class EditText(StatesGroup):
    value = State()

# -------------------- Handlers --------------------

@router.message(CommandStart())
async def start(message: Message):
    await message.answer(
        "<b>📋 Клиенты & расписание</b>\n\n"
        "Управляйте записями прямо в Telegram.\n"
        "Добавляйте клиентов, меняйте время и выгружайте всё в Excel.",
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
    await state.set_state(AddAppointment.name)
    await call.message.edit_text("<b>➕ Новая запись</b>\n\n👤 Введите имя клиента:", reply_markup=cancel_kb())
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
    await state.set_state(AddAppointment.service)
    await message.answer("💅 Какая услуга?", reply_markup=cancel_kb())


@router.message(AddAppointment.service)
async def add_service(message: Message, state: FSMContext):
    if not message.text or not message.text.strip():
        await message.answer("Введите название услуги.", reply_markup=cancel_kb())
        return
    await state.update_data(service=message.text.strip())
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
    await state.set_state(AddAppointment.comment)
    await call.message.edit_text("📝 Комментарий?\n\nНапишите текст или нажмите «Без комментария».", reply_markup=comment_keyboard())
    await call.answer()


def comment_keyboard():
    kb = InlineKeyboardBuilder()
    kb.button(text="Без комментария", callback_data="no_comment")
    kb.button(text="✕ Отмена", callback_data="cancel")
    kb.adjust(1)
    return kb.as_markup()


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
    appt_id = add_appointment(client_id, data["service"], data["date"], data["time"], comment)
    await state.clear()
    a = get_appointment(appt_id)
    text = "<b>✅ Запись сохранена</b>\n\n" + format_appt(a)
    if edit:
        await target.edit_text(text, reply_markup=appointment_kb(appt_id))
    else:
        await target.answer(text, reply_markup=appointment_kb(appt_id))


@router.callback_query(F.data == "schedule")
async def schedule(call: CallbackQuery):
    await call.message.edit_text("<b>📅 Расписание</b>\n\nВыберите день:", reply_markup=date_keyboard("viewdate"))
    await call.answer()


@router.callback_query(F.data.startswith("viewdate:"))
async def view_date(call: CallbackQuery):
    d = call.data.split(":", 1)[1]
    rows = appointments_for_date(d)
    title = format_day_title(d)
    count = len(rows)
    if not rows:
        text = f"<b>📅 {title}</b>\n\nСвободный день ✨"
    else:
        text = f"<b>📅 {title}</b>  ·  {count} запис." + ("\n\n" + "\n\n".join(format_appt(a) for a in rows))
    kb = InlineKeyboardBuilder()
    for a in rows:
        kb.button(text=f"{a['appointment_time']}  ·  {a['name']}", callback_data=f"open:{a['id']}")
    kb.button(text="➕ Добавить запись", callback_data="add")
    kb.button(text="‹ Выбрать другой день", callback_data="schedule")
    kb.button(text="⌂ Главное меню", callback_data="home")
    kb.adjust(1)
    await call.message.edit_text(text, reply_markup=kb.as_markup())
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


@router.message(EditText.value)
async def edit_text_value(message: Message, state: FSMContext):
    data = await state.get_data()
    value = (message.text or "").strip()
    if data["field"] == "comment" and value.lower() in {"нет", "-", "нету"}:
        value = ""
    if not value and data["field"] == "service":
        await message.answer("Услуга не может быть пустой.", reply_markup=cancel_kb())
        return
    update_appointment(data["appt_id"], data["field"], value)
    a = get_appointment(data["appt_id"])
    await state.clear()
    await message.answer("<b>✅ Изменено</b>\n\n" + format_appt(a), reply_markup=appointment_kb(data["appt_id"]))


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
    kb.button(text="⌂ Главное меню", callback_data="home")
    kb.adjust(2)
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


@router.callback_query(F.data == "excel")
async def excel(call: CallbackQuery):
    rows = all_appointments()
    if not rows:
        await call.answer("Пока нет записей", show_alert=True)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Расписание"
    headers = ["Дата", "Время", "Клиент", "Telegram", "Услуга", "Комментарий"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="6F4E37")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="D9D9D9")
    for r in rows:
        d = datetime.strptime(r["appointment_date"], "%Y-%m-%d").strftime("%d.%m.%Y")
        ws.append([d, r["appointment_time"], r["name"], r["telegram"], r["service"], r["comment"]])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    widths = [14, 10, 22, 22, 25, 40]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    clients_ws = wb.create_sheet("Клиенты")
    clients_ws.append(["Имя", "Telegram"])
    for cell in clients_ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="6F4E37")
    for c in all_clients():
        clients_ws.append([c["name"], c["telegram"]])
    clients_ws.column_dimensions["A"].width = 25
    clients_ws.column_dimensions["B"].width = 30
    clients_ws.freeze_panes = "A2"

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"raspisanie_{date.today():%Y-%m-%d}.xlsx"
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

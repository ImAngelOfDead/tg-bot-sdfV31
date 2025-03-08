import logging
import asyncio
import datetime
import psycopg2
import re

from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.filters import Command

DB_HOST = ""
DB_NAME = "mydatabase"
DB_USER = "myuser"
DB_PASSWORD = "mypassword"

TOKEN = ""
logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
dp = Dispatcher()

conn = psycopg2.connect(host=DB_HOST, database=DB_NAME, user=DB_USER, password=DB_PASSWORD)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    full_name VARCHAR,
    telegram_id VARCHAR UNIQUE,
    department VARCHAR,
    position VARCHAR,
    is_admin BOOLEAN DEFAULT FALSE,
    reminder VARCHAR
);
""")
cursor.execute("""
CREATE TABLE IF NOT EXISTS weekends (
    user_id INTEGER NOT NULL,
    date DATE NOT NULL,
    CONSTRAINT fk_weekends_user
        FOREIGN KEY (user_id)
        REFERENCES users (id)
        ON DELETE CASCADE
);
""")
cursor.execute("""
CREATE TABLE IF NOT EXISTS operations (
    user_id INTEGER NOT NULL,
    operation VARCHAR NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_operations_user
        FOREIGN KEY (user_id)
        REFERENCES users (id)
        ON DELETE CASCADE
);
""")
conn.commit()

def format_time(dt):
    return dt.strftime("%d.%m.%Y %H:%M:%S") if dt else ""

def get_or_create_user(telegram_id: str) -> int:
    cursor.execute("SELECT id FROM users WHERE telegram_id = %s", (telegram_id,))
    row = cursor.fetchone()
    if row:
        return row[0]
    cursor.execute("INSERT INTO users (telegram_id) VALUES (%s) RETURNING id", (telegram_id,))
    new_id = cursor.fetchone()[0]
    conn.commit()
    return new_id

def insert_operation(user_id: int, operation: str):
    cursor.execute("INSERT INTO operations (user_id, operation) VALUES (%s, %s)", (user_id, operation))
    conn.commit()

def get_last_operation_time(user_id: int, operation: str):
    cursor.execute("""
        SELECT created_at 
        FROM operations
        WHERE user_id = %s AND operation = %s
        ORDER BY created_at DESC
        LIMIT 1
    """, (user_id, operation))
    row = cursor.fetchone()
    return row[0] if row else None

def is_shift_active(user_id: int) -> bool:
    start_time = get_last_operation_time(user_id, "start_shift")
    if not start_time:
        return False
    cursor.execute("""
        SELECT COUNT(*) 
        FROM operations
        WHERE user_id = %s AND operation = 'end_shift' AND created_at > %s
    """, (user_id, start_time))
    return cursor.fetchone()[0] == 0

def is_break_active(user_id: int) -> bool:
    start_time = get_last_operation_time(user_id, "start_break")
    if not start_time:
        return False
    cursor.execute("""
        SELECT COUNT(*) 
        FROM operations
        WHERE user_id = %s AND operation = 'end_break' AND created_at > %s
    """, (user_id, start_time))
    return cursor.fetchone()[0] == 0

def get_last_shift_times(user_id: int):
    st_time = get_last_operation_time(user_id, "start_shift")
    if not st_time:
        return (None, None)
    cursor.execute("""
        SELECT created_at 
        FROM operations
        WHERE user_id = %s AND operation = 'end_shift' AND created_at > %s
        ORDER BY created_at ASC
        LIMIT 1
    """, (user_id, st_time))
    row = cursor.fetchone()
    return (st_time, row[0]) if row else (st_time, None)


user_dayoff_pages = {}

def build_day_off_reply_keyboard(page_start: datetime.date) -> ReplyKeyboardMarkup:
    today = datetime.date.today()
    min_date = today + datetime.timedelta(days=2)
    max_date = today + datetime.timedelta(days=30)
    if page_start < min_date:
        page_start = min_date

    keyboard_rows = []
    page_size = 5
    current_date = page_start
    for _ in range(page_size):
        if current_date > max_date:
            break

        keyboard_rows.append([KeyboardButton(text=current_date.strftime("%d.%m.%Y"))])
        current_date += datetime.timedelta(days=1)

    nav_row = []
    prev_page = page_start - datetime.timedelta(days=page_size)
    next_page = page_start + datetime.timedelta(days=page_size)
    if page_start > min_date:
        nav_row.append(KeyboardButton(text="←"))
    if current_date <= max_date:
        nav_row.append(KeyboardButton(text="→"))
    nav_row.append(KeyboardButton(text="Назад"))
    keyboard_rows.append(nav_row)

    return ReplyKeyboardMarkup(keyboard=keyboard_rows, resize_keyboard=True)

menu_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Время работы"), KeyboardButton(text="Поставить выходной")],
        [KeyboardButton(text="Начать смену"), KeyboardButton(text="Начать перерыв"),
         KeyboardButton(text="Закончить смену"), KeyboardButton(text="Закончить перерыв")]
    ],
    resize_keyboard=True
)

confirm_shift_keyboard = InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text="Подтвердить завершение смены", callback_data="confirm_end_shift")],
    [InlineKeyboardButton(text="Отмена", callback_data="cancel_end_shift")]
])

confirm_break_keyboard = InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text="Подтвердить завершение перерыва", callback_data="confirm_end_break")],
    [InlineKeyboardButton(text="Отмена", callback_data="cancel_end_break")]
])

@dp.message(Command("start"))
async def start_command(message: types.Message):
    await message.answer("Привет! Выбери действие:", reply_markup=menu_keyboard)

@dp.message(lambda msg: msg.text == "Начать смену")
async def start_shift(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))
    if is_shift_active(user_id):
        await message.answer("У вас уже есть активная смена. Завершите её.")
        return
    insert_operation(user_id, "start_shift")
    start_time = get_last_operation_time(user_id, "start_shift")
    await message.answer(f"Смена начата в {format_time(start_time)}. Пришли фото рабочего места, если требуется.")

@dp.message(lambda msg: msg.photo)
async def receive_photo(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))
    if not is_shift_active(user_id):
        await message.answer("Нет активной смены для фото.")
        return
    insert_operation(user_id, "photo_received")
    await message.answer("Фото принято. Хорошей смены!")

@dp.message(lambda msg: msg.text == "Начать перерыв")
async def start_break(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))
    if not is_shift_active(user_id):
        await message.answer("Сначала начните смену.")
        return
    if is_break_active(user_id):
        await message.answer("Перерыв уже идет. Завершите его.")
        return
    insert_operation(user_id, "start_break")
    start_time = get_last_operation_time(user_id, "start_break")
    await message.answer(f"Перерыв начат в {format_time(start_time)}.")

@dp.message(lambda msg: msg.text == "Закончить перерыв")
async def request_end_break(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))
    if not is_shift_active(user_id):
        await message.answer("Нет активной смены.")
        return
    if not is_break_active(user_id):
        await message.answer("Перерыв не начат или уже завершен.")
        return
    await message.answer("Завершить перерыв?", reply_markup=confirm_break_keyboard)

@dp.callback_query(lambda c: c.data == "confirm_end_break")
async def confirm_end_break(callback_query: types.CallbackQuery):
    user_id = get_or_create_user(str(callback_query.from_user.id))
    insert_operation(user_id, "end_break")
    end_time = get_last_operation_time(user_id, "end_break")
    await callback_query.answer()
    await callback_query.message.edit_text(f"Перерыв завершён в {format_time(end_time)}")
    await bot.send_message(callback_query.from_user.id, "Главное меню", reply_markup=menu_keyboard)

@dp.callback_query(lambda c: c.data == "cancel_end_break")
async def cancel_end_break(callback_query: types.CallbackQuery):
    await callback_query.answer()
    await callback_query.message.edit_text("Операция отменена.")
    await bot.send_message(callback_query.from_user.id, "Главное меню", reply_markup=menu_keyboard)

@dp.message(lambda msg: msg.text == "Закончить смену")
async def request_end_shift(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))
    if not is_shift_active(user_id):
        await message.answer("Нет активной смены.")
        return
    await message.answer("Завершить смену?", reply_markup=confirm_shift_keyboard)

@dp.callback_query(lambda c: c.data == "confirm_end_shift")
async def confirm_end_shift(callback_query: types.CallbackQuery):
    user_id = get_or_create_user(str(callback_query.from_user.id))
    insert_operation(user_id, "end_shift")
    end_time = get_last_operation_time(user_id, "end_shift")
    await callback_query.answer()
    await callback_query.message.edit_text(f"Смена завершена в {format_time(end_time)}")
    await bot.send_message(callback_query.from_user.id, "Главное меню", reply_markup=menu_keyboard)

@dp.callback_query(lambda c: c.data == "cancel_end_shift")
async def cancel_end_shift(callback_query: types.CallbackQuery):
    await callback_query.answer()
    await callback_query.message.edit_text("Операция отменена.")
    await bot.send_message(callback_query.from_user.id, "Главное меню", reply_markup=menu_keyboard)

@dp.message(lambda msg: msg.text == "Поставить выходной")
async def ask_day_off_date(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))
    today = datetime.date.today()
    min_date = today + datetime.timedelta(days=2)
    user_dayoff_pages[user_id] = min_date
    kb = build_day_off_reply_keyboard(min_date)
    await message.answer("Выберите дату для выходного:", reply_markup=kb)

@dp.message(lambda msg: msg.text 
            and get_or_create_user(str(msg.from_user.id)) in user_dayoff_pages 
            and (msg.text in ["Назад", "→", "←"] or re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", msg.text)))
async def day_off_selection(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))  # добавляем получение user_id
    page_start = user_dayoff_pages[user_id]
    today = datetime.date.today()
    min_date = today + datetime.timedelta(days=2)
    max_date = today + datetime.timedelta(days=30)
    page_size = 5

    if message.text == "Назад":
        user_dayoff_pages.pop(user_id, None)
        await message.answer("Главное меню", reply_markup=menu_keyboard)
        return
    elif message.text == "→":
        new_start = page_start + datetime.timedelta(days=page_size)
        if new_start > max_date:
            new_start = max_date
        user_dayoff_pages[user_id] = new_start
        kb = build_day_off_reply_keyboard(new_start)
        await message.answer("Выберите дату для выходного:", reply_markup=kb)
        return
    elif message.text == "←":
        new_start = page_start - datetime.timedelta(days=page_size)
        if new_start < min_date:
            new_start = min_date
        user_dayoff_pages[user_id] = new_start
        kb = build_day_off_reply_keyboard(new_start)
        await message.answer("Выберите дату для выходного:", reply_markup=kb)
        return
    else:
        try:
            selected_date = datetime.datetime.strptime(message.text, "%d.%m.%Y").date()
        except ValueError:
            return

        displayed_dates = []
        cur_date = page_start
        for _ in range(page_size):
            if cur_date > max_date:
                break
            displayed_dates.append(cur_date)
            cur_date += datetime.timedelta(days=1)

        if selected_date not in displayed_dates:
            return

        cursor.execute("SELECT COUNT(*) FROM weekends WHERE date = %s", (selected_date,))
        count = cursor.fetchone()[0]
        if count > 0:
            await message.answer("Этот день уже занят кем-то из отдела.")
        else:
            cursor.execute("INSERT INTO weekends (user_id, date) VALUES (%s, %s)", (user_id, selected_date))
            conn.commit()
            await message.answer(f"Выходной на {selected_date.strftime('%d.%m.%Y')} установлен.", reply_markup=menu_keyboard)
        user_dayoff_pages.pop(user_id, None)


@dp.message(lambda msg: msg.text == "Время работы")
async def work_time(message: types.Message):
    user_id = get_or_create_user(str(message.from_user.id))
    start_time, end_time = get_last_shift_times(user_id)
    if not start_time:
        await message.answer("Смена не начиналась.")
    else:
        if end_time:
            await message.answer(f"Смена:\nНачало: {format_time(start_time)}\nКонец: {format_time(end_time)}")
        else:
            await message.answer(f"Смена:\nНачало: {format_time(start_time)}\nНе завершена")



async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())

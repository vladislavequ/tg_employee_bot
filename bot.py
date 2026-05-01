import os
import logging
import asyncio
from datetime import datetime
from typing import Dict, Any

from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
    FSInputFile  # <-- Изменено: InputFile -> FSInputFile
)
from aiogram.filters import Command, StateFilter
from aiogram.enums import ParseMode  # <-- Изменен импорт ParseMode

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from aiohttp import web

# ==================================================================
# НАСТРОЙКИ
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise ValueError("BOT_TOKEN is not set!")

storage = MemoryStorage()
bot = Bot(token=BOT_TOKEN, parse_mode=ParseMode.HTML)  # parse_mode теперь из aiogram.enums
dp = Dispatcher(storage=storage)

# ==================================================================
class Form(StatesGroup):
    full_name = State()
    position = State()
    department = State()
    date = State()
    experience = State()
    supervisor = State()
    mission_goal = State()
    mission_impact = State()
    mission_absence = State()
    key_functions = State()
    competencies = State()
    direct_value = State()
    indirect_value = State()
    problems = State()
    improvements = State()
    goals = State()
    feedback = State()
    self_rating = State()
    self_strength = State()
    self_weakness = State()
    self_one_change = State()

user_data: Dict[int, Dict[str, Any]] = {}

# ==================================================================
# Клавиатуры
def get_main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🏁 Старт"), KeyboardButton(text="📝 Заполнить анкету")],
            [KeyboardButton(text="❌ Отмена"), KeyboardButton(text="ℹ️ Помощь")]
        ],
        resize_keyboard=True
    )

def get_cancel_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="❌ Отмена")]],
        resize_keyboard=True
    )

def get_rating_keyboard():
    kb = InlineKeyboardMarkup(row_width=5)
    for i in range(1, 11):
        kb.insert(InlineKeyboardButton(text=str(i), callback_data=f"rating_{i}"))
    return kb

def get_date_keyboard():
    today = datetime.now().strftime("%d.%m.%Y")
    kb = InlineKeyboardMarkup(row_width=2)
    kb.add(
        InlineKeyboardButton(text=f"Сегодня ({today})", callback_data="date_today"),
        InlineKeyboardButton(text="Ввести вручную", callback_data="date_manual")
    )
    return kb

# ==================================================================
# Хендлеры отмены и команд
@dp.message(Command("cancel"))
@dp.message(F.text == "❌ Отмена")
async def cmd_cancel(message: types.Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state is None:
        await message.answer("Нет активного опроса.", reply_markup=get_main_keyboard())
        return
    await state.clear()
    if message.from_user.id in user_data:
        del user_data[message.from_user.id]
    await message.answer("❌ Опрос отменён.", reply_markup=get_main_keyboard())

@dp.message(Command("start"))
@dp.message(F.text == "🏁 Старт")
async def cmd_start(message: types.Message):
    await message.answer(
        "🏢 <b>Бот для заполнения анкеты</b>\n\nНажмите «📝 Заполнить анкету».",
        reply_markup=get_main_keyboard()
    )

@dp.message(Command("help"))
@dp.message(F.text == "ℹ️ Помощь")
async def cmd_help(message: types.Message):
    await message.answer(
        "📌 <b>Доступные действия</b>:\n\n"
        "🏁 Старт – приветственное сообщение\n"
        "📝 Заполнить анкету – начать новый опрос\n"
        "❌ Отмена – прервать текущее анкетирование\n"
        "ℹ️ Помощь – это сообщение\n\n"
        "Для оценки (1-10) будут появляться кнопки.",
        reply_markup=get_main_keyboard()
    )

@dp.message(Command("fill"))
@dp.message(F.text == "📝 Заполнить анкету")
async def cmd_fill(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    if await state.get_state() is not None:
        await message.answer("Уже идет опрос. Отмените его.", reply_markup=get_cancel_keyboard())
        return
    user_data[user_id] = {}
    await message.answer("✏️ Введите ваше <b>ФИО</b>:", reply_markup=get_cancel_keyboard())
    await state.set_state(Form.full_name)

# ==================================================================
# Этот хендлер срабатывает, когда пользователь просто пишет боту вне сессии
@dp.message(StateFilter(None))
async def no_state_handler(message: types.Message):
    await message.answer("Нажмите «📝 Заполнить анкету», чтобы начать.", reply_markup=get_main_keyboard())

# ==================================================================
# Обработчики состояний — ОПРОС (здесь приведены для краткости только начало и конец)
@dp.message(Form.full_name)
async def process_full_name(message: types.Message, state: FSMContext):
    uid = message.from_user.id
    if uid not in user_data:
        await state.clear()
        await message.answer("❌ Ошибка: начните заново.", reply_markup=get_main_keyboard())
        return
    user_data[uid]['full_name'] = message.text.strip()
    await message.answer("Должность:")
    await state.set_state(Form.position)

# ... (остальные хендлеры состояний остаются без изменений) ...

# Последний хендлер для self_one_change
@dp.message(Form.self_one_change)
async def process_self_one_change(message: types.Message, state: FSMContext):
    uid = message.from_user.id
    if uid not in user_data:
        await state.clear()
        await message.answer("❌ Ошибка: начните анкету заново.", reply_markup=get_main_keyboard())
        return
    user_data[uid]['self_one_change'] = message.text.strip()
    await message.answer("Спасибо! Формирую документ...")

    try:
        filename = generate_docx(user_data[uid], uid)
        # Изменено: Используем FSInputFile вместо InputFile
        doc_file = FSInputFile(filename)
        await message.answer_document(
            document=doc_file,
            caption="✅ Ваша анкета готова!",
            reply_markup=get_main_keyboard()
        )
        os.remove(filename)
    except Exception as e:
        logger.exception("Ошибка при создании документа")
        await message.answer("❌ Ошибка при создании файла. Попробуйте еще раз.")

    if uid in user_data:
        del user_data[uid]
    await state.clear()

# ==================================================================
# Генерация .docx (функция без изменений)
def generate_docx(data: Dict[str, str], user_id: int) -> str:
    fio = data.get("full_name", "Сотрудник").replace(" ", "_")
    filename = f"Анкета_{fio}_{user_id}.docx"
    doc = Document()
    title = doc.add_heading("Анкета сотрудника", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fill_date = data.get("date", datetime.now().strftime("%d.%m.%Y"))
    doc.add_paragraph(f"Дата заполнения: {fill_date}")
    doc.add_paragraph()
    # Основная информация
    doc.add_heading("1. Основная информация", level=1)
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid'
    table.cell(0,0).text = "ФИО сотрудника"
    table.cell(0,1).text = data.get("full_name", "")
    table.cell(1,0).text = "Должность"
    table.cell(1,1).text = data.get("position", "")
    table.cell(2,0).text = "Отдел/Подразделение"
    table.cell(2,1).text = data.get("department", "")
    table.cell(3,0).text = "Дата заполнения"
    table.cell(3,1).text = fill_date
    table.cell(4,0).text = "Стаж в должности"
    table.cell(4,1).text = data.get("experience", "")
    table.cell(5,0).text = "Непосредственный руководитель"
    table.cell(5,1).text = data.get("supervisor", "")
    doc.add_heading("2. Миссия моей должности", level=1)
    doc.add_paragraph(f"Основная цель работы:\n{data.get('mission_goal', '')}")
    doc.add_paragraph(f"Влияние на успех компании:\n{data.get('mission_impact', '')}")
    doc.add_paragraph(f"Если бы должности не существовало:\n{data.get('mission_absence', '')}")
    doc.add_heading("3. Ключевые функции", level=1)
    doc.add_paragraph(data.get("key_functions", ""))
    doc.add_heading("4. Необходимые компетенции", level=1)
    doc.add_paragraph(data.get("competencies", ""))
    doc.add_heading("5. Ценность моей работы для компании", level=1)
    doc.add_paragraph("Прямая ценность:\n" + data.get("direct_value", ""))
    doc.add_paragraph("Косвенная ценность:\n" + data.get("indirect_value", ""))
    doc.add_heading("6. Проблемы и сложности", level=1)
    doc.add_paragraph(data.get("problems", ""))
    doc.add_heading("7. Идеи по улучшению", level=1)
    doc.add_paragraph(data.get("improvements", ""))
    doc.add_heading("8. Мои цели и развитие", level=1)
    doc.add_paragraph(data.get("goals", ""))
    doc.add_heading("9. Обратная связь и предложения", level=1)
    doc.add_paragraph(data.get("feedback", ""))
    doc.add_heading("10. Итоговая самооценка", level=1)
    doc.add_paragraph(f"Оценка эффективности (1-10):\n{data.get('self_rating', '')}")
    doc.add_paragraph(f"Главная сильная сторона:\n{data.get('self_strength', '')}")
    doc.add_paragraph(f"Над чем нужно работать:\n{data.get('self_weakness', '')}")
    doc.add_paragraph(f"Одно изменение:\n{data.get('self_one_change', '')}")
    doc.save(filename)
    return filename

# ==================================================================
# Health check сервер для Render
async def health_check(request):
    return web.Response(text="OK", status=200)

async def run_health_check():
    app = web.Application()
    app.router.add_get("/health", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", 8080)
    await site.start()
    logging.info("Health check server started on http://0.0.0.0:8080/health")

# ==================================================================
async def main():
    # Важно: сброс вебхука перед стартом
    await bot.delete_webhook(drop_pending_updates=True)
    logging.info("Webhook deleted, pending updates dropped.")

    asyncio.create_task(run_health_check())
    logging.info("Starting bot polling...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())

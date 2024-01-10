import pandas as pd
from aiogram import Bot, types, Dispatcher
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.utils import executor
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import json
import logging
import io
import tempfile

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

API_TOKEN = json.load(open("config.json"))["num"]
storage = MemoryStorage()
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=storage)
dp.middleware.setup(LoggingMiddleware())

file_checked = False

class AddExcelStates(StatesGroup):
    WaitingForFile = State()

@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await message.answer("Привет! Я бот, созданный для выдачи информации по группе студентов\nМне нужен файл Excel")
    await message.answer("Это можно сделать через команду /addExcel")

@dp.message_handler(commands=['addExcel'])
async def add_excel(message: types.Message):
    global file_checked
    await message.answer("Пожалуйста, отправьте файл Excel для добавления")
    await AddExcelStates.WaitingForFile.set()

@dp.message_handler(state=AddExcelStates.WaitingForFile, content_types=types.ContentType.DOCUMENT)
async def process_file(message: types.Message, state: FSMContext):
    try:
        file_info = message.document
        file_id = file_info.file_id

        # Изменено: используем bot.download_file_by_id
        file_content = await bot.download_file_by_id(file_id)

        logger.info("Файл успешно получен и обработан.")

        # Создаем временный файл и записываем в него содержимое BytesIO
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            temp_file.write(file_content.getvalue())

        # Читаем данные из временного файла
        df = pd.read_excel(temp_file.name)

        columns_to_check = ['Оценка', 'Сокращенная оценка', 'Период', 'Год', 'Семестр/Триместр', 'Курс', 'Часть года', 'Уровень контроля', 'Дисциплина', 'Личный номер студента', 'Группа', 'Факультет', 'Программа', 'Форма обучения', 'Тип финансирования']
        missing_columns = set(columns_to_check) - set(df.columns)

        if missing_columns:
            await message.reply(f"Отсутствующие столбцы: {', '.join(missing_columns)}")
        else:
            await message.reply("Файл соответствует ожидаемым столбцам!")
            file_checked = True


        await message.reply("Файл успешно обработан!")
        file_checked = True
    except Exception as e:
        logger.error(f"Произошла ошибка при чтении файла: {e}")
        await message.reply(f"Произошла ошибка при чтении файла: {e}")
    finally:
        await state.finish()

if __name__ == '__main__':
    from aiogram import executor
    executor.start_polling(dp, skip_updates=True)

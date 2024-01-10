import pandas as pd
from aiogram import Bot, types, Dispatcher
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import json


API_TOKEN = json.load(open("config.json"))["num"]

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)


class AddExcel(StatesGroup):
    WaitingForDocument = State()


@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await message.answer(
        "Привет! Я бот созданный для выдачи информации по группе студенте\nДля начала введите /addExcel 'номер группы'. Например /addExcel ПИ101")


@dp.message_handler(commands=['addExcel'])
async def add_excel(message: types.Message):
    global group_id
    command_parts = message.text.split()
    if len(command_parts) == 2:
        group_id = message.text.split(' ', 1)[1]
        await message.reply("Теперь отправьте файл Excel")
        await AddExcel.WaitingForDocument.set()
    else:
        await message.reply("Неверный формат команды. Введите /addExcel 'номер группы'. Например /addExcel ПИ101")


@dp.message_handler(content_types=types.ContentType.DOCUMENT, state=AddExcel.WaitingForDocument)
async def process_document(message: types.Message, state: FSMContext):
    document_id = message.document.file_id
    file_info = await bot.get_file(document_id)
    fi = file_info.file_path
    try:
        await message.reply("Обработка...")
        df_raw = pd.read_excel(f'https://api.telegram.org/file/bot{API_TOKEN}/{fi}')
        df = pd.DataFrame(df_raw)

        pi101_row = df[df['Группа'] == group_id]
        unique_student_row = pi101_row['Личный номер студента'].unique()
        unique_student_row_str = [str(number) for number in unique_student_row]
        unique_student_row_str = ', '.join(unique_student_row_str)
        year_control = pi101_row['Год'].unique()
        year_control_row_str = [str(number) for number in year_control]
        year_control_row_str = ', '.join(year_control_row_str)
        type_control = pi101_row['Уровень контроля'].unique()
        type_control_str = [str(number) for number in type_control]
        type_control_str = ', '.join(type_control_str)

        await message.reply(
            f'В исходном датасете содержалось {len(df)} оценок, из них {len(pi101_row)}\n оценок относятся к группе ПИ101\n В датасете находятся оценки {len(unique_student_row)} студентов со следующими личными номерами: {unique_student_row_str} \nИспользуемые формы контроля: {type_control_str}\nДанные представлены по следующим учебным годам: {year_control_row_str}')
    except Exception as e:
        await message.reply(f"read error: {e}")
    await state.finish()


if __name__ == '__main__':
    from aiogram import executor

    executor.start_polling(dp, skip_updates=True)


from aiogram import types, executor, Dispatcher, Bot
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from auth_data import token
from aiogram.dispatcher.filters import Text
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from docxtpl import DocxTemplate


# инициализация бота и диспетчера
storage = MemoryStorage()
bot = Bot(token=token)
dp = Dispatcher(bot, storage=storage)

# создание кнопки
but_1 = KeyboardButton('Запрос учредителю')
main_menu = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(but_1)


# классы для записи получаемых данных
class FSMUser(StatesGroup):
    type_of_manager = State()
    organization_name = State()
    date = State()
    name_of_director = State()
    index_address_director = State()
    decision = State()
    organization_info = State()
    application = State()

# начальний хендлер для старта записи данных
@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await bot.send_message(message.chat.id, 'Приветствую, выберете необходимый файл:', reply_markup=main_menu)


# далее идут хендлеры которые получают на хвод сообщение сохраняют в память и выдают сообщение о следующих данных
@dp.message_handler(state=None)
async def manager(message: types.Message):
    if message.text == 'Запрос учредителю':
        await FSMUser.type_of_manager.set()
        await message.reply('Выбери тип Арбитражного управляющего: конкурсный, временный и т.д.')
    else:
        await message.reply('Неизвестная команда')


@dp.message_handler(state='*', commands='отмена')
@dp.message_handler(Text(equals='отмена', ignore_case=True), state='*')
async def cancel_handler(massage: types.Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state is None:
        return
    await state.finish()
    await massage.reply('ОК')


@dp.message_handler(state=FSMUser.type_of_manager)
async def get_organization(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['type_of_manager'] = message.text
    await FSMUser.next()
    await message.reply('Введи название ООО')


@dp.message_handler(state=FSMUser.organization_name)
async def get_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['organization_name'] = message.text
    await FSMUser.next()
    await message.reply('Введи дату в виде: Исх. № 0903 от «09» марта 2022г.')


@dp.message_handler(state=FSMUser.date)
async def get_director_name(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['date'] = message.text
    await FSMUser.next()
    await message.reply('Введи ФИО учредителя')


@dp.message_handler(state=FSMUser.name_of_director)
async def get_director_address(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['name_of_director'] = message.text
    await FSMUser.next()
    await message.reply('Введи индекс и адрес учредителя в виде: 249500, Калужская обл., '
                        'Куйбышевский рн, д. Садовище, д.48, кв. 1')


@dp.message_handler(state=FSMUser.index_address_director)
async def get_decision(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['index_address_director'] = message.text
    await FSMUser.next()
    await message.reply('Введи решение суда в виде: Арбитражный суд Калужской области по делу № '
                        'А23-7211/2021 от «18» февраля 2022г. ')


@dp.message_handler(state=FSMUser.decision)
async def get_organization_info(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['decision'] = message.text
    await FSMUser.next()
    await message.reply('Введи ОГРН, ИНН и адрес организации в виде: ОГРН 1174027001610, '
                        'ИНН 4027131752; 248033, ОБЛАСТЬ КАЛУЖСКАЯ, Г. КАЛУГА, '
                        'ПР-Д 1-Й АКАДЕМИЧЕСКИЙ, Д. 10, ПОМ. 1')


@dp.message_handler(state=FSMUser.organization_info)
async def get_application(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['organization_info'] = message.text
    await FSMUser.next()
    await message.reply('Введи приложение в виде: Копия решения Арбитражного суда Калужской '
                        'области от «18» февраля 2022 г. по делу № А23-7211/2021.')


# в последнем хеднлере все данные подставляются в шаблон, после сохраняют в новый файл
@dp.message_handler(state=FSMUser.application)
async def closed(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['application'] = message.text
        doc = DocxTemplate("Запрос учредителю.docx")
        doc.render(data)
        doc.save("generate_doc.docx")

    # новый файл передается в клиенту и данные что были в памяти удаляются
    await message.reply_document(open("generate_doc.docx", 'rb'))
    await state.finish()

if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)

import requests
import os
from urllib.parse import urljoin
from bs4 import BeautifulSoup
from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from openpyxl import load_workbook

from config import TOKEN

bot = Bot(token=TOKEN)
dp = Dispatcher(bot)

URL = 'https://crc-reg.com/about/our-clients/'

r = requests.get(URL)
soup = BeautifulSoup(r.text, 'html.parser')
table1 = soup.find('table')

clients = soup.find_all('td', class_='clients-table-col-2')
inns = soup.find_all('td', class_='clients-table-col-3')
ogrns = soup.find_all('td', class_='clients-table-col-4')

ao = []
for client in clients:
    ao.append(client.get_text(strip=True))

inem = []
for inn in inns:
    inem.append("; ИНН: " + inn.get_text(strip=True))

ogern = []
for ogrn in ogrns:
    ogern.append("; ОГРН: " + ogrn.get_text(strip=True) + '\n\n')

allinfo = zip(ao, inem, ogern)
allinfolist = list(allinfo)

aostr = ''.join(''.join(elems) for elems in allinfolist)

URLCON = 'https://crc-reg.com/contacts/'

rcon = requests.get(URLCON)
soupcon = BeautifulSoup(rcon.text, 'html.parser')
contacts = soupcon.find_all('div', class_='contacts__sub-ttl')
for contact in contacts:
    contstr = contact.text

bankreks = soupcon.find_all('div', class_='banks')
for bank in bankreks:
    bankstr = bank.text

button_rstr = KeyboardButton('Реестры')

button_vernkank = KeyboardButton('Вернуться к разделу "Анкеты"')
button_vernkrasp = KeyboardButton('Вернуться к разделу "Распоряжения"')
button_vernkpp = KeyboardButton('Вернуться к разделу "Порядок предоставления документов"')
button_vernkols = KeyboardButton('Вернуться к разделу "Документы для открытия лицевого счета"')
button_vernkvzi = KeyboardButton('Вернуться к разделу "Документы для внесения записи в информацию ЛС о ЗЛ"')
button_vernkprskur = KeyboardButton('Вернуться к разделу "Прейскуранты"')

button_vglav = KeyboardButton('В главное меню')

button_pzip = KeyboardButton('Порядок заполнения и предоставления анкет, распоряжений и запросов')
button_ddpo = KeyboardButton('Документы для проведения операций')

button_prscur = KeyboardButton('Прейскуранты')

inline_kb_full = InlineKeyboardMarkup(row_width=1)
inline_btn_1 = InlineKeyboardButton('Прейскурант услуг Эмитентам', callback_data='btn1')
inline_btn_2 = InlineKeyboardButton('Прейскурант услуг зарегистрированным лицам', callback_data='btn2')
inline_btn_3 = InlineKeyboardButton('Прейскурант доп.услуг зарегистрированным лицам', callback_data='btn3')
inline_kb_full.add(inline_btn_1, inline_btn_2, inline_btn_3)

button_cont = KeyboardButton('Контакты')

button_pzipa = KeyboardButton('Анкеты')
button_pzipr = KeyboardButton('Распоряжения')
button_pzippp = KeyboardButton('Порядок предоставления документов')
button_ddols = KeyboardButton('Документы для открытия лицевого счета')
button_ddvz = KeyboardButton('Документы для внесения записи в информацию ЛС о ЗЛ')

inline_kb_full2 = InlineKeyboardMarkup(row_width=1)
inline_btn_4 = InlineKeyboardButton('Анкета Эмитента', callback_data='btn4')
inline_btn_5 = InlineKeyboardButton('Анкета физического лица', callback_data='btn5')
inline_btn_6 = InlineKeyboardButton('Анкета для юридического лица', callback_data='btn6')
inline_btn_7 = InlineKeyboardButton('Анкета нотариуса', callback_data='btn7')
inline_btn_8 = InlineKeyboardButton('Анкета Уполномоченного органа', callback_data='btn8')
inline_btn_9 = InlineKeyboardButton('Анкета доверительного управляющего', callback_data='btn9')
inline_kb_full2.add(inline_btn_4, inline_btn_5, inline_btn_6, inline_btn_7, inline_btn_8, inline_btn_9)

inline_kb_full3 = InlineKeyboardMarkup(row_width=1)
inline_btn_10 = InlineKeyboardButton('Распоряжение о совершении операции', callback_data='btn10')
inline_btn_11 = InlineKeyboardButton('Залоговое распоряжение', callback_data='btn11')
inline_btn_12 = InlineKeyboardButton('Распоряжение о внесении изменений о заложенных ЦБ и условиях залога',
                                     callback_data='btn12')
inline_btn_13 = InlineKeyboardButton('Распоряжение о передаче прав залога', callback_data='btn13')
inline_btn_14 = InlineKeyboardButton('Распоряжение о прекращении залога', callback_data='btn14')
inline_btn_15 = InlineKeyboardButton('Распоряжение на предоставление информации из реестра', callback_data='btn15')
inline_btn_16 = InlineKeyboardButton('Распоряжение на снятие факта ограничения операций с ЦБ', callback_data='btn16')
inline_btn_17 = InlineKeyboardButton('Распоряжение владельца о передаче ЦБ на депозитный ЛС', callback_data='btn17')
inline_kb_full3.add(inline_btn_10, inline_btn_11, inline_btn_12, inline_btn_13, inline_btn_14, inline_btn_15,
                    inline_btn_16, inline_btn_17)

inline_kb_full4 = InlineKeyboardMarkup(row_width=1)
inline_btn_18 = InlineKeyboardButton('Порядок предоставления документов для открытия ЛС', callback_data='btn18')
inline_btn_19 = InlineKeyboardButton('Порядок предоставления документов для совершения операции и\n для предоставления \
информации', callback_data='btn19')
inline_kb_full4.add(inline_btn_18, inline_btn_19)

inline_kb_full5 = InlineKeyboardMarkup(row_width=1)
inline_btn_20 = InlineKeyboardButton('Документы для физического лица', callback_data='btn20')
inline_btn_21 = InlineKeyboardButton('Документы для нотариуса', callback_data='btn21')
inline_btn_22 = InlineKeyboardButton('Документы ЮЛ, являющемуся резидентом РФ', callback_data='btn22')
inline_btn_23 = InlineKeyboardButton('Документы для Уполномоченного органа', callback_data='btn23')
inline_btn_24 = InlineKeyboardButton('Документы для ЮЛ-нерезидента', callback_data='btn24')
inline_btn_25 = InlineKeyboardButton('Документы для открытия казначейского ЛС эмитенту', callback_data='btn25')
inline_btn_26 = InlineKeyboardButton('Документы для открытия эмиссионного счета', callback_data='btn26')
inline_kb_full5.add(inline_btn_20, inline_btn_21, inline_btn_22, inline_btn_23, inline_btn_24, inline_btn_25,
                    inline_btn_26)

inline_kb_full6 = InlineKeyboardMarkup(row_width=1)
inline_btn_27 = InlineKeyboardButton('Документы для физического лица', callback_data='btn27')
inline_btn_28 = InlineKeyboardButton('Документы для юридического лица', callback_data='btn28')
inline_btn_29 = InlineKeyboardButton('Документы для передачи ЦБ при совершении сделки', callback_data='btn29')
inline_btn_30 = InlineKeyboardButton('Документы для передачи ЦБ при наследовании', callback_data='btn30')
inline_btn_31 = InlineKeyboardButton('Документы для исполнения судебных актов', callback_data='btn31')
inline_btn_32 = InlineKeyboardButton('Документы для передачи ЦБ при реорганизации', callback_data='btn32')
inline_btn_33 = InlineKeyboardButton('Документы для передачи ЦБ при ликвидации ЮЛ', callback_data='btn33')
inline_btn_34 = InlineKeyboardButton('Документы для передачи ЦБ при приватизации', callback_data='btn34')
inline_btn_35 = InlineKeyboardButton('Для внесения записи о прекращении залога и передаче ЦБ в связи с обращением на \
них взыскания без решения суда', callback_data='btn35')
inline_btn_36 = InlineKeyboardButton('Для внесения записи о факте фиксации ограничения операций с ЦБ по их полной \
оплате', callback_data='btn36')
inline_btn_37 = InlineKeyboardButton('Для внесения записи о фиксации/снятии факта ограничения операций с ЦБ по ЛС ЗЛ',
                                     callback_data='btn37')
inline_btn_38 = InlineKeyboardButton('Для внесения записи о зачислении/списании ЦБ со счета НД', callback_data='btn38')
inline_btn_39 = InlineKeyboardButton('Для зачисления на ЛС НД заложенных ЦБ', callback_data='btn39')
inline_btn_40 = InlineKeyboardButton('Особенности проведения операций по ЛС ДУ', callback_data='btn40')
inline_btn_41 = InlineKeyboardButton('Особенности проведения операций по депозитному ЛС', callback_data='btn41')
inline_btn_42 = InlineKeyboardButton('Объединение ЛС в реестре', callback_data='btn42')
inline_btn_43 = InlineKeyboardButton('Закрытие ЛС', callback_data='btn43')
inline_kb_full6.add(inline_btn_27, inline_btn_28, inline_btn_29, inline_btn_30, inline_btn_31, inline_btn_32,
                    inline_btn_33, inline_btn_34, inline_btn_35, inline_btn_36, inline_btn_37, inline_btn_38,
                    inline_btn_39,inline_btn_40,inline_btn_41,inline_btn_42,inline_btn_43)

greet_kb = ReplyKeyboardMarkup(row_width=1)
greet_kb.add(button_rstr, button_pzip, button_ddpo, button_prscur, button_cont)

greet_kb2 = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb2.add(button_vglav)

greet_kb3 = ReplyKeyboardMarkup(row_width=1)
greet_kb3.add(button_pzipa, button_pzipr, button_pzippp, button_vglav)

greet_kb4 = ReplyKeyboardMarkup(row_width=1)
greet_kb4.add(button_ddols, button_ddvz, button_vglav)

greet_kb5 = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb5.add(button_vernkank, button_vglav)

greet_kb6 = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb6.add(button_vernkrasp, button_vglav)

greet_kb7 = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb7.add(button_vernkpp, button_vglav)

greet_kb8 = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb8.add(button_vernkols, button_vglav)

greet_kb9 = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb9.add(button_vernkvzi, button_vglav)

greet_kb10 = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb10.add(button_vernkprskur, button_vglav)


@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    await message.answer('Выбери нужную кнопку в меню', reply_markup=greet_kb)


@dp.message_handler(lambda message: message.text == 'В главное меню')
async def glavmenu(message: types.Message):
    await message.answer('Выберите нужную кнопку в меню', reply_markup=greet_kb)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Анкеты"')
async def vernka(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full2)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Распоряжения"')
async def vernkr(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full3)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Порядок предоставления документов"')
async def vernkr(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full4)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Документы для открытия лицевого счета"')
async def vernkr(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full5)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Документы для внесения записи в информацию \
ЛС о ЗЛ"')
async def vernkr(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full6)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Прейскуранты"')
async def vernkr(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full)


@dp.message_handler(lambda message: message.text == 'Реестры')
async def allnashirstr(message: types.Message):
    await message.answer(aostr, reply_markup=greet_kb2)


@dp.message_handler(lambda message: message.text == 'Контакты')
async def allnashirstr(message: types.Message):
    await message.answer(contstr + bankstr, reply_markup=greet_kb2)


@dp.message_handler(lambda message: message.text == 'Порядок заполнения и предоставления анкет, распоряжений и \
запросов')
async def porzap(message: types.Message):
    await message.answer("Выберите документ, который Вас интересует", reply_markup=greet_kb3)


@dp.message_handler(lambda message: message.text == 'Документы для проведения операций')
async def porzap(message: types.Message):
    await message.answer("Выберите операцию, которая Вас интересует", reply_markup=greet_kb4)


@dp.message_handler(lambda message: message.text == 'Документы для открытия лицевого счета')
async def ankets(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full5)


@dp.message_handler(lambda message: message.text == 'Документы для внесения записи в информацию ЛС о ЗЛ')
async def ankets(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full6)


@dp.message_handler(lambda message: message.text == 'Прейскуранты')
async def ankets(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full)


link = 'https://crc-reg.com/for-issuers/price-lists/'
folder_location = r''
response = requests.get(link)
soup2 = BeautifulSoup(response.text, "html.parser")

findhref = soup2.select_one("td>a[href$='.pdf']")
filename = os.path.join(folder_location, findhref['href'].split('/')[-1])
with open(filename, 'wb') as f:
    f.write(requests.get(urljoin(link, findhref['href'])).content)


@dp.callback_query_handler(text='btn1')
async def price(call: types.CallbackQuery):
    with open(filename, 'rb') as file:
        await call.message.answer_document(file, caption='Прейскурант услуг, предоставляемых Эмитенту ценных бумаг',
                                           reply_markup=greet_kb10)
        file.close()


link2 = 'https://crc-reg.com/for-shareholders/price-lists/'
folder_location2 = r''
response2 = requests.get(link2)
soup3 = BeautifulSoup(response2.text, "html.parser")

findhref2 = soup3.select("a[href$='.pdf']")[1]
filename2 = os.path.join(folder_location2, findhref2['href'].split('/')[-1])
with open(filename2, 'wb') as f:
    f.write(requests.get(urljoin(link2, findhref2['href'])).content)


@dp.callback_query_handler(text='btn2')
async def price(call: types.CallbackQuery):
    with open(filename2, 'rb') as file:
        await call.message.answer_document(file, caption='Прейскурант услуг, предоставляемых зарегистрированному лицу',
                                           reply_markup=greet_kb10)
        file.close()


findhref3 = soup3.select("a[href$='.pdf']")[2]
filename3 = os.path.join(folder_location2, findhref3['href'].split('/')[-1])
with open(filename3, 'wb') as f:
    f.write(requests.get(urljoin(link2, findhref3['href'])).content)


@dp.callback_query_handler(text='btn3')
async def price(call: types.CallbackQuery):
    with open(filename3, 'rb') as file:
        await call.message.answer_document(file, caption='Прейскурант дополнительных услуг, предоставляемых \
        зарегистрированному лицу', reply_markup=greet_kb10)
        file.close()


@dp.message_handler(content_types=['text'])
async def ankets(message: types.Message):
    lookfor = message.text

    wb = load_workbook(filename='file.xlsx')
    ws = wb.active

    for rr in range(1, 2):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await message.answer(text=axa[x222:x222 + 4096], reply_markup=inline_kb_full2)
            else:
                await message.answer(text=axa, reply_markup=inline_kb_full2)

    for rr in range(8, 9):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await message.answer(text=axa[x222:x222 + 4096], reply_markup=inline_kb_full3)
            else:
                await message.answer(text=axa, reply_markup=inline_kb_full3)

    for rr in range(17, 18):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await message.answer(text=axa[x222:x222 + 4096], reply_markup=inline_kb_full4)
            else:
                await message.answer(text=axa, reply_markup=inline_kb_full4)

    wb.close()


@dp.callback_query_handler(lambda c: c.data and c.data.startswith('btn'))
async def price(call: types.CallbackQuery):
    lookfor = call.data

    wb = load_workbook(filename='file.xlsx')
    ws = wb.active

    for rr in range(2, 8):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await call.message.answer(text=axa[x222:x222 + 4096], reply_markup=greet_kb5)
            else:
                await call.message.answer(text=axa, reply_markup=greet_kb5)
    wb.close()

    for rr in range(9, 17):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await call.message.answer(text=axa[x222:x222 + 4096], reply_markup=greet_kb6)
            else:
                await call.message.answer(text=axa, reply_markup=greet_kb6)
    wb.close()

    for rr in range(18, 20):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await call.message.answer(text=axa[x222:x222 + 4096], reply_markup=greet_kb7)
            else:
                await call.message.answer(text=axa, reply_markup=greet_kb7)
    wb.close()

    for rr in range(20, 27):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await call.message.answer(text=axa[x222:x222 + 4096], reply_markup=greet_kb8)
            else:
                await call.message.answer(text=axa, reply_markup=greet_kb8)
    wb.close()

    for rr in range(27, 44):
        vl = ws.cell(row=rr, column=1).value
        if vl == lookfor:
            axa = ws.cell(row=rr, column=2).value
            if len(axa) > 4096:
                for x222 in range(0, len(axa), 4096):
                    await call.message.answer(text=axa[x222:x222 + 4096], reply_markup=greet_kb9)
            else:
                await call.message.answer(text=axa, reply_markup=greet_kb9)
    wb.close()


if __name__ == '__main__':
    executor.start_polling(dp)

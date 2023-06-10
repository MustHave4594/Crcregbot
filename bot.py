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

URL_RU = 'https://crc-reg.com/about/our-clients/'

r_url_ru = requests.get(URL_RU)
soup_url_ru = BeautifulSoup(r_url_ru.text, 'html.parser')
table_ru = soup_url_ru.find('table')

clients = soup_url_ru.find_all('td', class_='clients-table-col-2')
inns = soup_url_ru.find_all('td', class_='clients-table-col-3')
ogrns = soup_url_ru.find_all('td', class_='clients-table-col-4')

ao_ru = []
for client in clients:
    ao_ru.append(client.get_text(strip=True))

inn_ru = []
for inn in inns:
    inn_ru.append("; ИНН: " + inn.get_text(strip=True))

ogrn_ru = []
for ogrn in ogrns:
    ogrn_ru.append("; ОГРН: " + ogrn.get_text(strip=True) + '\n\n')

all_info = zip(ao_ru, inn_ru, ogrn_ru)
all_info_list = list(all_info)

ao_ru_str = ''.join(''.join(elems) for elems in all_info_list)

URL_CONTACTS_ru = 'https://crc-reg.com/contacts/'

r_contact_ru = requests.get(URL_CONTACTS_ru)
soup_contact_ru = BeautifulSoup(r_contact_ru.text, 'html.parser')
contacts_ru = soup_contact_ru.find_all('div', class_='contacts__sub-ttl')
contact_str_ru = ' '
for contact_ru in contacts_ru:
    contact_str_ru = contact_str_ru.join(contact_ru.text.split(' '))
contact_str_ru_replace = contact_str_ru.replace('Телеграм-бот','')

bank_info_ru = soup_contact_ru.find_all('div', class_='banks')
for bank in bank_info_ru:
    bank_str_ru = bank.text

btn_ru = KeyboardButton('Русский язык')
btn_en = KeyboardButton('English language')

btn_list_ru = KeyboardButton('Реестры')

btn_return_A_f_ru = KeyboardButton('Вернуться к разделу "Анкеты"')
btn_return_orders_ru = KeyboardButton('Вернуться к разделу "Распоряжения"')
btn_return_p_s_d_ru = KeyboardButton('Вернуться к разделу "Порядок предоставления документов"')
btn_return_d_o_a_ru = KeyboardButton('Вернуться к разделу "Документы для открытия лицевого счета"')
btn_return_d_t_p_i_ru = KeyboardButton('Вернуться к разделу "Документы для внесения записи в информацию ЛС о ЗЛ"')
btn_return_price_ru = KeyboardButton('Вернуться к разделу "Прейскуранты"')

btn_main_menu_ru = KeyboardButton('В главное меню')

btn_p_f_s_ru = KeyboardButton('Порядок заполнения и предоставления анкет, распоряжений и запросов')
btn_d_c_t_ru = KeyboardButton('Документы для проведения операций')

btn_price_ru = KeyboardButton('Прейскуранты')

btn_contacts_ru = KeyboardButton('Контакты')

btn_A_f_ru = KeyboardButton('Анкеты')
btn_orders_ru = KeyboardButton('Распоряжения')
btn_p_s_d_ru = KeyboardButton('Порядок предоставления документов')
btn_d_o_a_ru = KeyboardButton('Документы для открытия лицевого счета')
btn_d_t_p_i_ru = KeyboardButton('Документы для внесения записи в информацию ЛС о ЗЛ')


inline_kb_full_ru = InlineKeyboardMarkup(row_width=1)
inline_btn_1_ru = InlineKeyboardButton('Прейскурант услуг Эмитентам', callback_data='btn1_ru')
inline_btn_2_ru = InlineKeyboardButton('Прейскурант услуг зарегистрированным лицам', callback_data='btn2_ru')
inline_btn_3_ru = InlineKeyboardButton('Прейскурант доп.услуг зарегистрированным лицам', callback_data='btn3_ru')
inline_kb_full_ru.add(inline_btn_1_ru, inline_btn_2_ru, inline_btn_3_ru)

inline_kb_full2_ru = InlineKeyboardMarkup(row_width=1)
inline_btn_4_ru = InlineKeyboardButton('Анкета Эмитента', callback_data='btn4_ru')
inline_btn_5_ru = InlineKeyboardButton('Заявление-анкета физического лица', callback_data='btn5_ru')
inline_btn_6_ru = InlineKeyboardButton('Заявление-анкета юридического лица', callback_data='btn6_ru')
inline_btn_7_ru = InlineKeyboardButton('Заявление-анкета нотариуса', callback_data='btn7_ru')
inline_btn_8_ru = InlineKeyboardButton('Заявление-анкета органа государственной власти (органа местного самоуправления)', callback_data='btn8_ru')
inline_btn_9_ru = InlineKeyboardButton('Заявление-анкета доверительного управляющего', callback_data='btn9_ru')
inline_btn_10_ru = InlineKeyboardButton('Заявление-анкета эскроу-агента', callback_data='btn10_ru')
inline_btn_11_ru = InlineKeyboardButton('Заявление-анкета инвестиционного товарищества', callback_data='btn11_ru')
inline_btn_12_ru = InlineKeyboardButton('Заявление-анкета. Цифровые финансовые активы', callback_data='btn12_ru')
inline_btn_13_ru = InlineKeyboardButton('Заявление-анкета иностранной структуры без образования юридического лица', callback_data='btn13_ru')
inline_btn_14_ru = InlineKeyboardButton('Заявление-анкета залогодержателя (физическое лицо)', callback_data='btn14_ru')
inline_btn_15_ru = InlineKeyboardButton('Заявление-анкета залогодержателя (юридическое лицо)', callback_data='btn15_ru')
inline_kb_full2_ru.add(inline_btn_4_ru, inline_btn_5_ru, inline_btn_6_ru, inline_btn_7_ru, inline_btn_8_ru, inline_btn_9_ru, inline_btn_10_ru, inline_btn_11_ru, inline_btn_12_ru, inline_btn_13_ru, inline_btn_14_ru, inline_btn_15_ru)

inline_kb_full3_ru = InlineKeyboardMarkup(row_width=1)
inline_btn_16_ru = InlineKeyboardButton('Распоряжение о совершении операции', callback_data='btn16_ru')
inline_btn_17_ru = InlineKeyboardButton('Залоговое распоряжение', callback_data='btn17_ru')
inline_btn_18_ru = InlineKeyboardButton('Распоряжение о внесении изменений о заложенных ЦБ и условиях залога', callback_data='btn18_ru')
inline_btn_19_ru = InlineKeyboardButton('Распоряжение о передаче прав залога', callback_data='btn19_ru')
inline_btn_20_ru = InlineKeyboardButton('Распоряжение о прекращении залога', callback_data='btn20_ru')
inline_btn_21_ru = InlineKeyboardButton('Распоряжение на предоставление информации из реестра', callback_data='btn21_ru')
inline_btn_22_ru = InlineKeyboardButton('Распоряжение на снятие факта ограничения операций с ЦБ', callback_data='btn22_ru')
inline_btn_23_ru = InlineKeyboardButton('Распоряжение владельца о передаче ЦБ на депозитный ЛС', callback_data='btn23_ru')
inline_kb_full3_ru.add(inline_btn_16_ru, inline_btn_17_ru, inline_btn_18_ru, inline_btn_19_ru, inline_btn_20_ru, inline_btn_21_ru, inline_btn_22_ru, inline_btn_23_ru)

inline_kb_full4_ru = InlineKeyboardMarkup(row_width=1)
inline_btn_24_ru = InlineKeyboardButton('Порядок предоставления документов для открытия ЛС', callback_data='btn24_ru')
inline_btn_25_ru = InlineKeyboardButton('Порядок предоставления документов для совершения операции и\n для предоставления информации', callback_data='btn25_ru')
inline_kb_full4_ru.add(inline_btn_24_ru, inline_btn_25_ru)

inline_kb_full5_ru = InlineKeyboardMarkup(row_width=1)
inline_btn_26_ru = InlineKeyboardButton('Документы для физического лица', callback_data='btn26_ru')
inline_btn_27_ru = InlineKeyboardButton('Документы для нотариуса', callback_data='btn27_ru')
inline_btn_28_ru = InlineKeyboardButton('Документы ЮЛ, являющемуся резидентом РФ', callback_data='btn28_ru')
inline_btn_29_ru = InlineKeyboardButton('Документы для органа государственной власти (органа местного самоуправления)', callback_data='btn29_ru')
inline_btn_30_ru = InlineKeyboardButton('Документы для ЮЛ-нерезидента', callback_data='btn30_ru')
inline_btn_31_ru = InlineKeyboardButton('Документы для открытия казначейского ЛС эмитенту', callback_data='btn31_ru')
inline_kb_full5_ru.add(inline_btn_26_ru, inline_btn_27_ru, inline_btn_28_ru, inline_btn_29_ru, inline_btn_30_ru, inline_btn_31_ru)

inline_kb_full6_ru = InlineKeyboardMarkup(row_width=1)
inline_btn_32_ru = InlineKeyboardButton('Документы для физического лица', callback_data='btn32_ru')
inline_btn_33_ru = InlineKeyboardButton('Документы для юридического лица', callback_data='btn33_ru')
inline_btn_34_ru = InlineKeyboardButton('Документы для передачи ЦБ при совершении сделки', callback_data='btn34_ru')
inline_btn_35_ru = InlineKeyboardButton('Документы для передачи ЦБ при наследовании', callback_data='btn35_ru')
inline_btn_36_ru = InlineKeyboardButton('Документы для исполнения судебных актов', callback_data='btn36_ru')
inline_btn_37_ru = InlineKeyboardButton('Документы для передачи ЦБ при реорганизации', callback_data='btn37_ru')
inline_btn_38_ru = InlineKeyboardButton('Документы для передачи ЦБ при ликвидации ЮЛ', callback_data='btn38_ru')
inline_btn_39_ru = InlineKeyboardButton('Документы для передачи ЦБ при приватизации', callback_data='btn39_ru')
inline_btn_40_ru = InlineKeyboardButton('Для внесения записи о фиксации/прекращении права залога (последующего залога) ЦБ', callback_data='btn40_ru')
inline_btn_41_ru = InlineKeyboardButton('Для внесения записи о факте фиксации ограничения операций с ЦБ по их полной оплате', callback_data='btn41_ru')
inline_btn_42_ru = InlineKeyboardButton('Для внесения записи о фиксации/снятии факта ограничения операций с ЦБ по ЛС ЗЛ', callback_data='btn42_ru')
inline_btn_43_ru = InlineKeyboardButton('Для внесения записи о зачислении/списании ЦБ со счета НД', callback_data='btn43_ru')
inline_btn_44_ru = InlineKeyboardButton('Для зачисления на ЛС НД заложенных ЦБ', callback_data='btn44_ru')
inline_btn_45_ru = InlineKeyboardButton('Особенности проведения операций по ЛС ДУ', callback_data='btn45_ru')
inline_btn_46_ru = InlineKeyboardButton('Особенности проведения операций по депозитному ЛС', callback_data='btn46_ru')
inline_btn_47_ru = InlineKeyboardButton('Объединение ЛС в реестре', callback_data='btn47_ru')
inline_btn_48_ru = InlineKeyboardButton('Закрытие ЛС', callback_data='btn48_ru')
inline_kb_full6_ru.add(inline_btn_32_ru, inline_btn_33_ru, inline_btn_34_ru, inline_btn_35_ru, inline_btn_36_ru, inline_btn_37_ru, inline_btn_38_ru, inline_btn_39_ru, inline_btn_40_ru, inline_btn_41_ru, inline_btn_42_ru ,inline_btn_43_ru, inline_btn_44_ru, inline_btn_45_ru, inline_btn_46_ru, inline_btn_47_ru, inline_btn_48_ru)


greet_kb = ReplyKeyboardMarkup(row_width=1)
greet_kb.add(btn_ru, btn_en)

greet_kb_ru = ReplyKeyboardMarkup(row_width=1)
greet_kb_ru.add(btn_list_ru, btn_p_f_s_ru, btn_d_c_t_ru, btn_price_ru, btn_contacts_ru)

greet_kb2_ru = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb2_ru.add(btn_main_menu_ru)

greet_kb3_ru = ReplyKeyboardMarkup(row_width=1)
greet_kb3_ru.add(btn_A_f_ru, btn_orders_ru, btn_p_s_d_ru, btn_main_menu_ru)

greet_kb4_ru = ReplyKeyboardMarkup(row_width=1)
greet_kb4_ru.add(btn_d_o_a_ru, btn_d_t_p_i_ru, btn_main_menu_ru)

greet_kb5_ru = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb5_ru.add(btn_return_A_f_ru, btn_main_menu_ru)

greet_kb6_ru = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb6_ru.add(btn_return_orders_ru, btn_main_menu_ru)

greet_kb7_ru = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb7_ru.add(btn_return_p_s_d_ru, btn_main_menu_ru)

greet_kb8_ru = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb8_ru.add(btn_return_d_o_a_ru, btn_main_menu_ru)

greet_kb9_ru = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb9_ru.add(btn_return_d_t_p_i_ru, btn_main_menu_ru)

greet_kb10_ru = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb10_ru.add(btn_return_price_ru, btn_main_menu_ru)

btn_return_A_f_en = KeyboardButton('Go back to the section "Application forms"')
btn_return_orders_en = KeyboardButton('Go back to the section "Orders"')
btn_return_p_s_d_en = KeyboardButton('Go back to the section "Procedure for submitting documents"')
btn_return_d_o_a_en = KeyboardButton('Go back to the section "Procedure for submitting documents to open a personal account"')
btn_return_d_t_p_i_en = KeyboardButton('Go back to the section "Procedure for submitting documents for a transaction and for providing information"')
btn_main_menu_en = KeyboardButton('Go back to the section "To the main menu"')
btn_p_f_s_en = KeyboardButton('Procedure for filling out and submitting Application forms, Orders and Requests')
btn_d_c_t_en = KeyboardButton('Documents for conducting transactions')

btn_contacts_en = KeyboardButton('Contacts')

btn_A_f_en = KeyboardButton('Application forms')
btn_orders_en = KeyboardButton('Orders')
btn_p_s_d_en = KeyboardButton('Procedure for submitting documents')
btn_d_o_a_en = KeyboardButton('Procedure for submitting documents to open a personal account')
btn_d_t_p_i_en = KeyboardButton('Procedure for submitting documents for a transaction and for providing information')

inline_kb_full_en = InlineKeyboardMarkup(row_width=1)
inline_btn_1_en = InlineKeyboardButton('Issuer Application Form', callback_data ='btn1_en')
inline_btn_2_en = InlineKeyboardButton('Individual Application Form', callback_data ='btn2_en')
inline_btn_3_en = InlineKeyboardButton('Legal Entity Application Form', callback_data ='btn3_en')
inline_btn_4_en = InlineKeyboardButton('Notary Application Form', callback_data ='btn4_en')
inline_btn_5_en = InlineKeyboardButton('Authorised Body Application Form', callback_data ='btn5_en')
inline_btn_6_en = InlineKeyboardButton('Trustee Application Form', callback_data ='btn6_en')
inline_btn_7_en = InlineKeyboardButton('Pledgee(individual) Application Form', callback_data ='btn7_en')
inline_btn_8_en = InlineKeyboardButton('Pledgee(legal entity) Application Form', callback_data ='btn8_en')
inline_kb_full_en.add(inline_btn_1_en, inline_btn_2_en, inline_btn_3_en, inline_btn_4_en, inline_btn_5_en,
                      inline_btn_6_en, inline_btn_7_en, inline_btn_8_en)

inline_kb_full_2_en = InlineKeyboardMarkup(row_width=1)
inline_btn_9_en = InlineKeyboardButton('Transaction Order', callback_data ='btn9_en')
inline_btn_10_en = InlineKeyboardButton('Pledge Order', callback_data ='btn10_en')
inline_btn_11_en = InlineKeyboardButton('Order to amend the pledged securities and the terms and conditions of the pledge', callback_data ='btn11_en')
inline_btn_12_en = InlineKeyboardButton('Order to assign pledge rights', callback_data ='btn12_en')
inline_btn_13_en = InlineKeyboardButton('Order to terminate pledge', callback_data ='btn13_en')
inline_btn_14_en = InlineKeyboardButton('Order to provide information from the registry', callback_data ='btn14_en')
inline_btn_15_en = InlineKeyboardButton('Order to lift restrictions on transactions with securities', callback_data='btn15_en')
inline_btn_16_en = InlineKeyboardButton("Owner's order to transfer securities to the deposit account", callback_data='btn16_en')
inline_kb_full_2_en.add(inline_btn_9_en, inline_btn_10_en, inline_btn_11_en, inline_btn_12_en, inline_btn_13_en,
                        inline_btn_14_en, inline_btn_15_en, inline_btn_16_en)

inline_kb_full_3_en = InlineKeyboardMarkup(row_width=1)
inline_btn_17_en = InlineKeyboardButton('Procedure for submitting documents to open a personal account', callback_data='btn17_en')
inline_btn_18_en = InlineKeyboardButton('Procedure for submitting documents for a transaction and for providing information', callback_data='btn18_en')
inline_kb_full_3_en.add(inline_btn_17_en, inline_btn_18_en)

inline_kb_full_4_en = InlineKeyboardMarkup(row_width=1)
inline_btn_19_en = InlineKeyboardButton('Documents for an individual', callback_data='btn19_en')
inline_btn_20_en = InlineKeyboardButton('Documents for the notary', callback_data='btn20_en')
inline_btn_21_en = InlineKeyboardButton('Documents of a legal entity being a resident of the Russian Federation', callback_data='btn21_en')
inline_btn_22_en = InlineKeyboardButton('Documents for the Authorised Body', callback_data='btn22_en')
inline_btn_23_en = InlineKeyboardButton('Documents for non-resident legal entity', callback_data='btn23_en')
inline_btn_24_en = InlineKeyboardButton('Documents to open a treasury account for the issuer', callback_data='btn24_en')
inline_btn_25_en = InlineKeyboardButton('Documents to open a personal account for common hared ownership', callback_data='btn25_en')
inline_kb_full_4_en.add(inline_btn_19_en, inline_btn_20_en, inline_btn_21_en, inline_btn_22_en, inline_btn_23_en, inline_btn_24_en, inline_btn_25_en)

inline_kb_full_5_en = InlineKeyboardMarkup(row_width=1)
inline_btn_26_en = InlineKeyboardButton('Documents for an individual', callback_data='btn26_en')
inline_btn_27_en = InlineKeyboardButton('Documents for a legal entity', callback_data='btn27_en')
inline_btn_28_en = InlineKeyboardButton('Documents to transfer securities upon execution of a transaction', callback_data='btn28_en')
inline_btn_29_en = InlineKeyboardButton('Documents to transfer securities upon inheritance', callback_data='btn29_en')
inline_btn_30_en = InlineKeyboardButton('Documents to enforce court orders', callback_data='btn30_en')
inline_btn_31_en = InlineKeyboardButton('Documents to transfer securities upon reorganisation', callback_data='btn31_en')
inline_btn_32_en = InlineKeyboardButton('Documents to transfer securities upon liquidation of a legal entity', callback_data='btn32_en')
inline_btn_33_en = InlineKeyboardButton('Documents to transfer securities upon privatisation', callback_data='btn33_en')
inline_btn_34_en = InlineKeyboardButton('In order to make an entry on the recording / termination of a pledge (subsequent pledge) of securities', callback_data='btn34_en')
inline_btn_35_en = InlineKeyboardButton('To make an entry on committing restrictions on securities transactions to pay them in full', callback_data='btn35_en')
inline_btn_36_en = InlineKeyboardButton('To make an entry on committing / removing restrictions on securities transactions on the personal account of a registered person', callback_data='btn36_en')
inline_btn_37_en = InlineKeyboardButton('To make an entry on deposition / withdrawal of securities to / from an account of a nominee holder', callback_data='btn37_en')
inline_btn_38_en = InlineKeyboardButton('To deposit pledged securities to an account of a nominee holder', callback_data='btn38_en')
inline_btn_39_en = InlineKeyboardButton("Specifics of transactions in the trustee's personal account", callback_data='btn39_en')
inline_btn_40_en = InlineKeyboardButton('Specifics of deposit personal account transactions', callback_data='btn40_en')
inline_btn_41_en = InlineKeyboardButton('Pooling of personal accounts in the register', callback_data='btn41_en')
inline_btn_42_en = InlineKeyboardButton('Closure of a personal account', callback_data='btn42_en')
inline_kb_full_5_en.add(inline_btn_26_en, inline_btn_27_en, inline_btn_28_en, inline_btn_29_en, inline_btn_30_en, inline_btn_31_en, inline_btn_32_en, inline_btn_33_en, inline_btn_34_en, inline_btn_35_en, inline_btn_36_en, inline_btn_37_en, inline_btn_38_en, inline_btn_39_en, inline_btn_40_en, inline_btn_41_en, inline_btn_42_en)

greet_kb_en = ReplyKeyboardMarkup(row_width=1)
greet_kb_en.add(btn_p_f_s_en, btn_d_c_t_en, btn_contacts_en)

greet_kb_2_en = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb_2_en.add(btn_main_menu_en)

greet_kb_3_en = ReplyKeyboardMarkup(row_width=1)
greet_kb_3_en.add(btn_A_f_en, btn_orders_en, btn_p_s_d_en, btn_main_menu_en)

greet_kb_4_en = ReplyKeyboardMarkup(row_width=1)
greet_kb_4_en.add(btn_d_o_a_en, btn_d_t_p_i_en, btn_main_menu_en)

greet_kb_5_en = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb_5_en.add(btn_return_A_f_en, btn_main_menu_en)

greet_kb_6_en = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb_6_en.add(btn_return_orders_en, btn_main_menu_en)

greet_kb_7_en = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb_7_en.add(btn_return_p_s_d_en, btn_main_menu_en)

greet_kb_8_en = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb_8_en.add(btn_return_d_o_a_en, btn_main_menu_en)

greet_kb_9_en = ReplyKeyboardMarkup(resize_keyboard=True)
greet_kb_9_en.add(btn_return_d_t_p_i_en, btn_main_menu_en)

@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    await message.answer('Выберите язык / Select a language', reply_markup=greet_kb)

@dp.message_handler(lambda message: message.text == 'Русский язык')
async def ru_language(message: types.Message):
    await message.answer('Выбран русский язык.\n Выбери нужную кнопку в меню', reply_markup=greet_kb_ru)

@dp.message_handler(lambda message: message.text == 'В главное меню')
async def main_menu_ru (message: types.Message):
    await message.answer('Выберите нужную кнопку в меню', reply_markup=greet_kb_ru)

@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Анкеты"')
async def back_a_f_ru (message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full2_ru)

@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Распоряжения"')
async def back_orders_ru (message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full3_ru)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Порядок предоставления документов"')
async def back_p_s_d_ru (message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full4_ru)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Документы для открытия лицевого счета"')
async def back_p_p_a_ru (message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full5_ru)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Документы для внесения записи в информацию ЛС о ЗЛ"')
async def back_p_s_t_p_ru (message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full6_ru)


@dp.message_handler(lambda message: message.text == 'Вернуться к разделу "Прейскуранты"')
async def back_price_ru (message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full_ru)


@dp.message_handler(lambda message: message.text == 'Реестры')
async def list_ru(message: types.Message):
    await message.answer(ao_ru_str)
    await message.answer("Если хотите передать реестр на обслуживание в Регистратор, то пишите нам на электронную почту почту info@crc-reg.com ", reply_markup=greet_kb2_ru)

@dp.message_handler(lambda message: message.text == 'Контакты')
async def contacts_ru(message: types.Message):
    await message.answer(contact_str_ru_replace + bank_str_ru, reply_markup=greet_kb2_ru)


@dp.message_handler(lambda message: message.text == 'Порядок заполнения и предоставления анкет, распоряжений и запросов')
async def p_f_s_ru (message: types.Message):
    await message.answer("Выберите документ, который Вас интересует", reply_markup=greet_kb3_ru)


@dp.message_handler(lambda message: message.text == 'Документы для проведения операций')
async def d_c_t_ru (message: types.Message):
    await message.answer("Выберите операцию, которая Вас интересует", reply_markup=greet_kb4_ru)


@dp.message_handler(lambda message: message.text == 'Документы для открытия лицевого счета')
async def d_o_a_ru (message: types.Message):
    await message.answer(message.text)
    await message.answer("Если Вы не нашли нужную Вам операцию, то пишите нам на электронную почту info@crc-reg.com", reply_markup=inline_kb_full5_ru)


@dp.message_handler(lambda message: message.text == 'Документы для внесения записи в информацию ЛС о ЗЛ')
async def d_t_p_i_ru (message: types.Message):
    await message.answer(message.text)
    await message.answer("Если Вы не нашли нужную Вам операцию, то пишите нам на электронную почту info@crc-reg.com", reply_markup=inline_kb_full6_ru)


@dp.message_handler(lambda message: message.text == 'Прейскуранты')
async def price_ru(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full_ru)


link_issuers = 'https://crc-reg.com/for-issuers/price-lists/'
folder_location_issuers = r''
response_issuers = requests.get(link_issuers)
soup_price_issuers_ru = BeautifulSoup(response_issuers.text, "html.parser")

find_href_issuers = soup_price_issuers_ru.select_one("td>a[href$='.pdf']")
filename_issuers = os.path.join(folder_location_issuers, find_href_issuers['href'].split('/')[-1])
with open(filename_issuers, 'wb') as f:
    f.write(requests.get(urljoin(link_issuers, find_href_issuers['href'])).content)


@dp.callback_query_handler(text='btn1_ru')
async def price_ru(call: types.CallbackQuery):
    with open(filename_issuers, 'rb') as file:
        await call.message.answer_document(file, caption='Прейскурант услуг, предоставляемых Эмитенту ценных бумаг', reply_markup=greet_kb10_ru)
        file.close()


link_shareholders = 'https://crc-reg.com/for-shareholders/price-lists/'
folder_location_shareholders = r''
response_shareholders = requests.get(link_shareholders)
soup_shareholders = BeautifulSoup(response_shareholders.text, "html.parser")

find_href_shareholders = soup_shareholders.select("a[href$='.pdf']")[1]
filename_shareholders = os.path.join(folder_location_shareholders, find_href_shareholders ['href'].split('/')[-1])
with open(filename_shareholders, 'wb') as f:
    f.write(requests.get(urljoin(link_shareholders, find_href_shareholders['href'])).content)


@dp.callback_query_handler(text='btn2_ru')
async def price_ru(call: types.CallbackQuery):
    with open(filename_shareholders, 'rb') as file:
        await call.message.answer_document(file, caption='Прейскурант услуг, предоставляемых зарегистрированному лицу', reply_markup=greet_kb10_ru)
        file.close()


find_href_shareholders_2 = soup_shareholders.select("a[href$='.pdf']")[2]
filename_shareholders_2 = os.path.join(folder_location_shareholders, find_href_shareholders_2['href'].split('/')[-1])
with open(filename_shareholders_2, 'wb') as f:
    f.write(requests.get(urljoin(link_shareholders, find_href_shareholders_2['href'])).content)


@dp.callback_query_handler(text='btn3_ru')
async def price_ru(call: types.CallbackQuery):
    with open(filename_shareholders_2, 'rb') as file:
        await call.message.answer_document(file, caption='Прейскурант дополнительных услуг, предоставляемых зарегистрированному лицу', reply_markup=greet_kb10_ru)
        file.close()



@dp.message_handler(lambda message: message.text == 'English language')
async def en_language(message: types.Message):
    await message.answer('English is selected.\n Select the desired button in the menu', reply_markup=greet_kb_en)


@dp.message_handler(lambda message: message.text == 'To the main menu')
async def main_menu_en(message: types.Message):
    await message.answer('Select the desired button in the menu', reply_markup = greet_kb_en)

@dp.message_handler(lambda message: message.text == 'Go back to the section "Application forms"')
async def back_a_f_en(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full_en)

@dp.message_handler(lambda message: message.text == 'Go back to the section "Orders"')
async def back_orders_en(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full_2_en)

@dp.message_handler(lambda message: message.text == 'Go back to the section "Procedure for submitting documents"')
async def back_p_s_d_en(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full_3_en)

@dp.message_handler(lambda message: message.text == 'Go back to the section "Procedure for submitting documents to open a personal account"')
async def back_p_p_a_en(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full_4_en)

@dp.message_handler(lambda message: message.text == 'Go back to the section "Procedure for submitting documents for a transaction and for providing information"')
async def back_p_s_t_p_en(message: types.Message):
    await message.answer(message.text, reply_markup=inline_kb_full_5_en)

@dp.message_handler(lambda message: message.text == 'Contacts')
async def contacts_en(message: types.Message):
    await message.answer(contact_str_en_replace, reply_markup=greet_kb_2_en)

@dp.message_handler(lambda message: message.text == 'Go back to the section "To the main menu"')
async def back_p_s_t_p_en(message: types.Message):
    await message.answer(message.text, reply_markup=greet_kb_en)

@dp.message_handler(lambda message: message.text == 'Procedure for filling out and submitting Application forms, Orders and Requests')
async def p_f_s_en(message: types.Message):
    await message.answer('Select the document you are interested in', reply_markup = greet_kb_3_en)

@dp.message_handler(lambda message: message.text == 'Documents for conducting transactions')
async def d_c_t_en(message: types.Message):
    await message.answer('Select the operation that interests you', reply_markup = greet_kb_4_en)

@dp.message_handler(lambda message: message.text == 'Procedure for submitting documents to open a personal account')
async def p_p_a_en(message: types.Message):
    await message.answer(message.text)
    await message.answer('If you have not found the operation you need, then write to us by e - mail info @ crc - reg.com', reply_markup = inline_kb_full_4_en)

@dp.message_handler(lambda message: message.text == 'Procedure for submitting documents for a transaction and for providing information')
async def p_s_t_p_en(message: types.Message):
    await message.answer(message.text)
    await message.answer('If you have not found the operation you need, then write to us by e - mail info @ crc - reg.com', reply_markup = inline_kb_full_5_en)




URL_CONTACTS_en = 'https://crc-reg.com/en/about/contacts/'

r_contact_en = requests.get(URL_CONTACTS_en)
soup_contact_en = BeautifulSoup(r_contact_en.text, 'html.parser')
contacts_en = soup_contact_en.find_all('div', class_='cnt')
contact_str_en = ' '
for contact_en in contacts_en:
    contact_str_en = contact_str_en.join(contact_en.text.split(' '))
contact_str_en_replace = contact_str_en.replace('Contacts','')


@dp.message_handler(content_types=['text'])
async def from_file(message: types.Message):
    look_for = message.text

    wb = load_workbook(filename='file.xlsx')
    ws = wb.active

    for cell_ru in range(1, 2):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await message.answer(text=value_ru[len_ru:len_ru + 4096])
                    await message.answer("Если Вы не нашли нужную Вам анкету, то пишите нам на электронную почту info@crc-reg.com",
                                         reply_markup=inline_kb_full2_ru)
            else:
                await message.answer(text=value_ru)
                await message.answer("Если Вы не нашли нужную Вам анкету, то пишите нам на электронную почту info@crc-reg.com",
                                     reply_markup=inline_kb_full2_ru)

    for cell_ru in range(14, 15):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await message.answer(text=value_ru[len_ru:len_ru + 4096])
                    await message.answer(
                        "Если Вы не нашли нужное Вам Распоряжение, то пишите нам на электронную почту info@crc-reg.com",
                        reply_markup=inline_kb_full3_ru)
            else:
                await message.answer(text=value_ru)
                await message.answer(
                    "Если Вы не нашли нужное Вам Распоряжение, то пишите нам на электронную почту info@crc-reg.com",
                    reply_markup=inline_kb_full3_ru)

    for cell_ru in range(23, 24):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await message.answer(text=value_ru[len_ru:len_ru + 4096],reply_markup=inline_kb_full4_ru)
            else:
                await message.answer(text=value_ru,reply_markup=inline_kb_full4_ru)


    for cell_en in range(49, 50):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await message.answer(text=value_en[len_en:len_en + 4096])
                    await message.answer('If you have not found the Application form you need, then write to us by e - mail info @ crc - reg.com', reply_markup = inline_kb_full_en)
            else:
                    await message.answer(text=value_en)
                    await message.answer('If you have not found the Application form you need, then write to us by e - mail info @ crc - reg.com', reply_markup = inline_kb_full_en)


    for cell_en in range(58, 59):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await message.answer(text=value_en[len_en:len_en + 4096])
                    await message.answer('If you have not found the Order you need, then write to us by e-mail info@crc-reg.com', reply_markup = inline_kb_full_2_en)
            else:
                    await message.answer(text=value_en)
                    await message.answer('If you have not found the Order you need, then write to us by e-mail info@crc-reg.com', reply_markup = inline_kb_full_2_en)

    for cell_en in range(67, 68):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await message.answer(text=value_en[len_en:len_en + 4096], reply_markup=inline_kb_full_3_en)
            else:
                    await message.answer(text=value_en, reply_markup=inline_kb_full_3_en)
    wb.close()


@dp.callback_query_handler(lambda c: c.data and c.data.startswith('btn'))
async def from_file(call: types.CallbackQuery):
    look_for = call.data

    wb = load_workbook(filename='file.xlsx')
    ws = wb.active

    for cell_ru in range(2, 14):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await call.message.answer(text=value_ru[len_ru:len_ru + 4096], reply_markup=greet_kb5_ru)
            else:
                await call.message.answer(text=value_ru, reply_markup=greet_kb5_ru)
    wb.close()


    for cell_ru in range(15, 23):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await call.message.answer(text=value_ru[len_ru:len_ru + 4096], reply_markup=greet_kb6_ru)
            else:
                await call.message.answer(text=value_ru, reply_markup=greet_kb6_ru)
    wb.close()


    for cell_ru in range(24, 26):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await call.message.answer(text=value_ru[len_ru:len_ru + 4096], reply_markup=greet_kb7_ru)
            else:
                await call.message.answer(text=value_ru, reply_markup=greet_kb7_ru)
    wb.close()


    for cell_ru in range(26, 32):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await call.message.answer(text=value_ru[len_ru:len_ru + 4096], reply_markup=greet_kb8_ru)
            else:
                await call.message.answer(text=value_ru, reply_markup=greet_kb8_ru)
    wb.close()


    for cell_ru in range(32, 49):
        key_value_ru = ws.cell(row=cell_ru, column=1).value
        if key_value_ru == look_for:
            value_ru = ws.cell(row=cell_ru, column=2).value
            if len(value_ru) > 4096:
                for len_ru in range(0, len(value_ru), 4096):
                    await call.message.answer(text=value_ru[len_ru:len_ru + 4096], reply_markup=greet_kb9_ru)
            else:
                await call.message.answer(text=value_ru, reply_markup=greet_kb9_ru)
    wb.close()



    for cell_en in range(50, 58):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await call.message.answer(text=value_en[len_en:len_en + 4096], reply_markup=greet_kb_5_en)
            else:
                await call.message.answer(text=value_en, reply_markup=greet_kb_5_en)
    wb.close()

    for cell_en in range(59, 67):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await call.message.answer(text=value_en[len_en:len_en + 4096], reply_markup=greet_kb_6_en)
            else:
                await call.message.answer(text=value_en, reply_markup=greet_kb_6_en)
    wb.close()

    for cell_en in range(68, 60):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await call.message.answer(text=value_en[len_en:len_en + 4096], reply_markup=greet_kb_7_en)
            else:
                await call.message.answer(text=value_en, reply_markup=greet_kb_7_en)
    wb.close()

    for cell_en in range(70, 77):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await call.message.answer(text=value_en[len_en:len_en + 4096], reply_markup=greet_kb_8_en)
            else:
                await call.message.answer(text=value_en, reply_markup=greet_kb_8_en)
    wb.close()

    for cell_en in range(77, 94):
        key_value_en = ws.cell(row=cell_en, column=1).value
        if key_value_en == look_for:
            value_en = ws.cell(row=cell_en, column=2).value
            if len(value_en) > 4096:
                for len_en in range(0, len(value_en), 4096):
                    await call.message.answer(text=value_en[len_en:len_en + 4096], reply_markup=greet_kb_9_en)
            else:
                await call.message.answer(text=value_en, reply_markup=greet_kb_9_en)
    wb.close()


if __name__ == '__main__':
    executor.start_polling(dp)

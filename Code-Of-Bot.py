#Бот расписатор
#Открытая версия дли GitHub. Проверьте токены, пароли, и другие настройки в кавычках
# -*- coding: utf-8 -*-

#Старт импорта
import shutil
from aiogram import Bot, Dispatcher, types, executor
from datetime import datetime
from pathlib import Path
import openpyxl
from aiogram.types import ReplyKeyboardMarkup, ReplyKeyboardRemove, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
import random
import requests
import os
#Конец импорта


root_path = Path(__file__).resolve().parents[0] #определяем корневую папку (исп. для серверов)
print(root_path)

global sob
sob = ''
code = 'Пароль от грядущих событий'

daylist = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'] #список дней и их ряды в таблице
stolblist = [2, 10, 18, 26, 34]





klassovindex = 0
cur_d=0
needtime = 0

x = 1 #счетчик

b_tod = KeyboardButton('Уроки сегодня') #описание кнопок быстрой клавиатуры
b_mor = KeyboardButton('Уроки завтра')
b_ch = KeyboardButton('Сменить класс')
b_day = KeyboardButton('Расписание по дням')
b_zam = KeyboardButton('Инфо')
b_sob = KeyboardButton('Грядущие события')
b_men = KeyboardButton('Меню')

kb_client = ReplyKeyboardMarkup()

kb_client.row(b_tod, b_mor).row(b_zam, b_sob).row(b_ch, b_day).row(b_men)  #расположение кнопок

def settime():
    global fact_d
    global cur_d
    global d_tab
    global d_tab_tom
    fact_d=cur_d #определяем время, день недели
    if fact_d >4:
        d_tom = 0
    else:
        d_tom=fact_d+1
    cur_t = datetime.now().time()
    cur_d = datetime.now().weekday()
    d_tab = 8*fact_d+2
    d_tab_tom = 8*d_tom+2

settime()

col = int(0)

bot = Bot('Код из BotFather')
dp = Dispatcher(bot)
admin_id = 'вставьте сюда свой id из телеграма. чтобы его найти, включите функцию с доп. информацией в настройках для разработчика десктопа тг'

book = openpyxl.open('rasp3kv1.xlsx', read_only=True) #открываем таблицу с расписанием

sheet = book.active


klassusers = [] #пользователи
klasscol = [] #колонки классов в соответствии с номером пользователя в предыдущем списке
listklass=[] #список классов
idklass=[] #id классов

val = '0'
i=1

while val != None: #определяем классы
    val = sheet.cell(1, i).value
    print(val)
    if val!='Дни' and val!='Уроки' and val!='Каб.':
        listklass.append(val)
        idklass.append(i-1)
    i+=1

def func1():
    global d_tab
    global x
    global mes
    mes = d_tab + x
    x = x+1

def func2():
    global d_tab_tom
    global x
    global mes_t
    mes_t = d_tab_tom + x
    x = x+1


@dp.message_handler(commands=['start']) #вывод при комм. /start
async def echo(message: types.Message):
    await bot.send_message(message.from_user.id, "Привет! Этот бот поможет тебе не копаться в таблицах с расписанием! В помощь тебе, или облегчить жизнь ленивцам!) "
                                                 "Теперь напиши свой класс. Пример: 8г1", reply_markup=kb_client)




@dp.message_handler() #обработка сообщений
async def rasp_tod(message: types.Message):
    global col
    global x
    global klassovindex
    settime()
    print(f'Получено сообщение: {message.text}')
    print(f'Сообщение выше отправил:{message.from_user.id}')


    #Выбор класса
    if str(message.text) in listklass:
        klassovindex = listklass.index(message.text)
        if (int(message.from_user.id)) in klassusers:
            if klassusers.index(message.from_user.id) <= len(klasscol):
                klasscol[klassusers.index(message.from_user.id)] = idklass[klassovindex]
            else:
                await bot.send_message(message.from_user.id, 'Произошла ошибка. Код ошибки - KChs1')
                return
        else:
            klassusers.append(message.from_user.id)
            klasscol.append(idklass[klassovindex])
        col = idklass[klassovindex]
        col = int(col)
        await bot.send_message(message.from_user.id, 'Принято. Вы выбрали класс '+ message.text +'. Если это не нужный вам класс, то пропишите /start, или нажмите кнопку Сменить Класс')
        print("выполнил")
        print(klassusers)
        print(klasscol)


    #Гайд по выбору класса
    elif message.text == 'Сменить класс':
        await bot.send_message(message.from_user.id, 'Вы начали процесс смены класса. Чтобы его сменить, напишите свой класс. Пример: 6в1')


    #Гайд по расписанию по дням
    elif message.text == 'Расписание по дням':
        await bot.send_message(message.from_user.id,'Чтобы получить расписание на конкретный день, напишите его краткое название, пр.: Пятница - Пт (Раскладка важна) ') #задел


    #Уроки сегодня
    elif message.text == 'Уроки сегодня':
        if int(message.from_user.id) not in klassusers: #еcли у пользователя не выбран класс
            await bot.send_message(message.from_user.id,'У вас не выбран класс! Нажмите на кнопку "Сменить класс"!')
        elif cur_d>4 and cur_d<7: #если сегодня выходной
            await bot.send_message(message.from_user.id, 'Сегодня выходной! Воспользуйтесь функцией "Расписание по дням"')
        else: #если все нормально
            await bot.send_message(message.from_user.id,
                                   f'*Хорошо. Вот расписание на сегодня для класса {(listklass[klasscol[int(klassusers.index(message.from_user.id))] - 2])}:*',
                                   parse_mode='Markdown')
            x = 0 #далее перебираем уроки в нужный день нужного класса, отсылаем если соответствующие ячейки не пусты
            while x < 8:
                func1()
                if sheet[mes][int(klasscol[int(klassusers.index(message.from_user.id))])].value == None and x > 3:
                    x = 8
                elif sheet[mes][int(klasscol[int(klassusers.index(message.from_user.id))])].value != None:
                    await bot.send_message(message.from_user.id, f'Урок {x}: ' + str(
                        sheet[mes][int(klasscol[int(klassusers.index(message.from_user.id))])].value))
        print(d_tab)
        print(fact_d)
        print(klassovindex)



    #Уроки завтра
    elif message.text == 'Уроки завтра': #далее все аналогично Урокам Сегодня
        if int(message.from_user.id) not in klassusers:
            await bot.send_message(message.from_user.id,'У вас не выбран класс! Нажмите на кнопку "Сменить класс"!')
        elif cur_d>3 and cur_d<6:
            await bot.send_message(message.from_user.id, 'Завтра будет выходной! Воспользуйтесь функцией "Расписание по дням"')
        else:
            await bot.send_message(message.from_user.id,
                                   f'*Хорошо. Вот расписание на завтра для класса {(listklass[klasscol[int(klassusers.index(message.from_user.id))] - 2])}:*',
                                   parse_mode='Markdown')
            x = 0
            while x < 8:
                func2()
                if sheet[mes_t][int(klasscol[int(klassusers.index(message.from_user.id))])].value == None and x>3:
                    x = 8
                else:
                    await bot.send_message(message.from_user.id, str(f'Урок {x}: ' + sheet[mes_t][int(klasscol[int(klassusers.index(message.from_user.id))])].value))
            print(d_tab_tom)
            print(cur_d)
            print(klassusers)
            print(klasscol)


    #Уроки по дням
    elif message.text in daylist and message.text!='Сб' and message.text!='Вс': #если в сообщении код рабочего дня
        if int(message.from_user.id) not in klassusers: #если у пользователя не выбран класс
            await bot.send_message(message.from_user.id,'У вас не выбран класс! Нажмите на кнопку "*Сменить класс*"!', parse_mode='Markdown')
        else: #аналогично с уроками по дням все
            await bot.send_message(message.from_user.id, f'*Хорошо. Вот расписание на {message.text} для класса {(listklass[klasscol[int(klassusers.index(message.from_user.id))]-2])}:*', parse_mode='Markdown')
            x=0
            while x <8:
                if sheet[stolblist[daylist.index(message.text)]+x][int(klasscol[int(klassusers.index(message.from_user.id))])].value == None and x>3:
                    x = 8
                elif sheet[stolblist[daylist.index(message.text)]+x][int(klasscol[int(klassusers.index(message.from_user.id))])].value != None:
                    await bot.send_message(message.from_user.id, f'Урок {x+1}: '+ str(sheet[stolblist[daylist.index(message.text)]+x][int(klasscol[int(klassusers.index(message.from_user.id))])].value))
                x=x+1


    #Задаем события
    elif message.text[:6]==code and message.from_user.id==admin_id: #вместо 6 поставьте свою длину пароля.
        global sob
        sob=message.text[6:] #аналогично
        await bot.send_message(message.from_user.id,'Принято. Новый текст грядущих событий -'+sob)


    #Выдаем дату
    elif message.text=='Дата':
        await bot.send_message(message.from_user.id, 'Сегодня ' + str(datetime.today()))
        await bot.send_message(message.from_user.id, 'День недели: '+ daylist[cur_d])


    #Бросаем монетку (если вы это увидели, поздравляем - это пасхалка :DD)
    elif message.text=='Монетка':
        await bot.send_message(message.from_user.id, random.choice(['Орёл', 'Решка']))


    #Выдаем таблицу
    elif message.text=='Таблица':
        await bot.send_document(message.from_user.id, open('rasp3kv1.xlsx', 'rb'))


    #Выдаем события
    elif message.text=='Грядущие события' and sob!='':
        await bot.send_message(message.from_user.id,sob)


    #Если в событиях пусто
    elif message.text=='Грядущие события' and sob=='':

        await bot.send_message(message.from_user.id, 'Пока тут пусто. Если считаете, что пропущено какое-то событие - обратитесь к @imonbog')


    #Выдаем справку
    elif message.text=='Инфо':
        await bot.send_message(message.from_user.id,'Это бот расписатор - @BauRasp_bot \n \n Что умеет этот бот? \n •Выдавать расписание на сегодня (команда Уроки сегодня) \n •Выдавать расписание на завтра (команда Уроки завтра) \n •Выдавать расписание на конкретный день (команда Расписание по дням) \n •Грядущие события. Отправит грядущие события корпуса. Будьте внимательны, они заполняются вручную \n \n Лайфхаки и скрытые функции: \n •Для выбора класса необязательно использовать команду, можно просто написать класс \n •Для получения расписания по дням необязательно использовать команду, можно просто написать краткое название дня \n •Команда Дата отправит вам сегодняшнее число и день недели \n •Команда Монетка подбросит монетку. Почему бы и нет:D \n •Команда Таблица отправит расписание в формате xlsx. \n \n ⚠️ Будьте внимательны - раскладка всегда важна \n \n Владелец бота - @imonbog. Все вопросы и предложения направляйте к нему \n \n Бот создан учеником 1580 и для учеников 1580 :)')


    #Выдаем меню
    elif message.text=='Меню':
        if os.path.exists('menu.pdf'): #обновляем меню
            os.remove('menu.pdf')
        response = requests.get('https://prikaz.1580.ru/sites/prikaz.1580.ru/files/food/15.pdf') #вставьте здесь в кавычки адрес меню. в новом типе сайтов московских школ он статичен
        with open("menu.pdf", "wb") as file:
            file.write(response.content)
        await bot.send_document(message.from_user.id, open('menu.pdf', 'rb'))

    #Разбираемся с таблицей
    elif 'Таблица' in message.text:
        if message.text[7:] in listklass:
            tablkl = message.text[7:]
            print(tablkl)
            if os.path.exists(f'{root_path}/{tablkl}.xlsx'): #если таблица уже есть в корневой
                await bot.send_message(message.from_user.id,'Таблица найдена, отправляю...')
                print('2')
                await bot.send_document(message.from_user.id, open(f'{tablkl}.xlsx', 'rb'))

            else: #а если нет...
                await bot.send_message(message.from_user.id, 'Таблица не найдена, копирую шаблон...')
                shutil.copyfile(f'{root_path}/shablon.xlsx', f'{root_path}/{tablkl}.xlsx') #копируем шаблон
                tabl = openpyxl.open(f'{root_path}/{tablkl}.xlsx', read_only=False) #открываем шаблон
                sh = tabl.active
                print(idklass)
                print(listklass)
                print(f'ВтЗн = {int(idklass[listklass.index(tablkl)])}')
                await bot.send_message(message.from_user.id, f'Генерирую таблицу с нужными уроками класса {tablkl}...')
                for i in range(5): #для каждых 5 рабочих дней...
                    for k in range(8): #для каждых 8 уроков... заполнение аналогично урокам сегодня
                        print(f'i = {i*8+1+k}, k = {k}')
                        if sheet[i*8+k+2][int(idklass[listklass.index(tablkl)])].value!=None:
                            print(sheet[i*8+k+2][int(idklass[listklass.index(tablkl)])].value)
                            sh.cell(row=k+2, column=i+2).value = (sheet[i*8+k+2][int(idklass[listklass.index(tablkl)])].value)
                        print('1')
                sh['A1'] = tablkl
                await bot.send_message(message.from_user.id, 'Генерация завершена, отправляю...')
                tabl.save(f'{root_path}/{tablkl}.xlsx') #сохраняемся чтобы после не регенерить
                await bot.send_document(message.from_user.id, open(f'{tablkl}.xlsx', 'rb')) #отсылаем

    else: #если текст сообщения не подошел ни под одну команду
        await bot.send_message(message.from_user.id,'Неопознанная команда! Пожалуйста, воспользуйтесь кнопками на специальной клавиатуре, или пропишите команду "Инфо" без кавычек! (Чтобы открыть клавиатуру нажмите на значок 4х точек в квадрате ниже)  \n\n ⚠️Для всех команд должен быть соблюдён регистр')

    print(d_tab)
    print(cur_d)
    x = 1

if __name__ == '__main__': 
    executor.start_polling(dp, skip_updates=False)
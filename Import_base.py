# -*- coding: utf-8 -*-
import pyodbc #подключаем библиотеку для работы с ODBC
import xlrd #подключаем библиотку для работы с XLSX
import pandas as pd #подключаем библиотеку для работы с данными

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
cursor.execute("""
CREATE TABLE [dbo].[Base](
    Ag_name nvarchar(50),
    Rat_name nvarchar(50),
    Grade nvarchar(50), 
    Outlook nvarchar(50),
    Change nvarchar(50),
    Date_str nvarchar(50),
    Ent_name nvarchar(50),
    Okpo float,
    Ogrn float,
    Inn float,
    Finst nvarchar(50),
    Ag_id nvarchar(50),
    Rat_sctr nvarchar(50),
    Rat_type nvarchar(50),
    Horizon nvarchar(50),
    Scale_typer nvarchar(50),
    Currency nvarchar(50),
    Backed_flag nvarchar(50))""") #создаем таблицу с необходимыми полями
connection.commit() #сохраняем изменения

data = pd.read_excel('Base.xlsx') #читаем книгу
data = data.rename(columns={'Ag_name': 'Ag_name',
                            'Rat_name': 'Rat_name',
                            'Grade': 'Grade',
                            'Outlook': 'Outlook',
                            'Change': 'Change',
                            'Date_str': 'Date_str',
                            'Ent_name': 'Ent_name',
                            'Okpo': 'Okpo',
                            'Ogrn': 'Ogrn',
                            'Inn': 'Inn',
                            'Finst': 'Finst',
                            'Ag_id': 'Ag_id',
                            'Rat_sctr': 'Rat_sctr',
                            'Rat_type': 'Rat_type',
                            'Horizon': 'Horizon',
                            'Scale_typer': 'Scale_typer',
                            'Currency': 'Currency',
                            'Backed_flag': 'Backed_flag'}) #избавляемся от пробелов

data.to_excel('Base.xlsx', index=False) #экспорт данных

book = xlrd.open_workbook('Base.xlsx') #открываем книгу
sheet = book.sheet_by_name('Sheet1') #открываем лист

query = ("""
INSERT INTO [Rates].[dbo].[Base](
    Ag_name,
    Rat_name,
    Grade, 
    Outlook,
    Change,
    Date_str,
    Ent_name,
    Okpo,
    Ogrn,
    Inn,
    Finst,
    Ag_id,
    Rat_sctr,
    Rat_type,
    Horizon,
    Scale_typer,
    Currency,
    Backed_flag) 
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""")

for r in range(1, sheet.nrows): #цикл построчного считывания данных
    Ag_name = sheet.cell(r,0).value
    Rat_name = sheet.cell(r,1).value
    Grade = sheet.cell(r,2).value
    Outlook = sheet.cell(r,3).value
    Change = sheet.cell(r,4).value
    Date_str = sheet.cell(r,5).value
    Ent_name = sheet.cell(r,6).value
    Okpo = sheet.cell(r,7).value
    Ogrn = sheet.cell(r,8).value
    Inn = sheet.cell(r,9).value
    Finst = sheet.cell(r,10).value
    Ag_id = sheet.cell(r,11).value
    Rat_sctr = sheet.cell(r,12).value
    Rat_type = sheet.cell(r,13).value
    Horizon = sheet.cell(r,14).value
    Scale_typer = sheet.cell(r,15).value
    Currency = sheet.cell(r,16).value
    Backed_flag = sheet.cell(r,17).value
    
    values = (Ag_name,
              Rat_name,
              Grade, 
              Outlook,
              Change,
              Date_str,
              Ent_name,
              Okpo,
              Ogrn,
              Inn,
              Finst,
              Ag_id,
              Rat_sctr,
              Rat_type,
              Horizon,
              Scale_typer,
              Currency,
              Backed_flag)
    cursor.execute(query, values) #назначаем значения сторка/столбец
connection.commit() #сохраняем изменения
if cursor.rowcount > 0: #проверка импорта данных
    print('Import successful')
else:
    print('Import failed')
cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение
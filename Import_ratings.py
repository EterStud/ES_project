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
CREATE TABLE [dbo].[Ratings](
    ID int,
    Code_Name nvarchar(50),
    Agency nvarchar(50), 
    Full_Name_Rus nvarchar(100),
    Full_Name_Eng nvarchar(100),
    For_Instrument nvarchar(50),
    For_Company nvarchar(50))""") #создаем таблицу с необходимыми полями
connection.commit() #сохраняем изменения

data = pd.read_excel('Ratings.xlsx') #читаем книгу
data = data.rename(columns={'ID': 'ID',
                            'Code_Name': 'Code_Name',
                            'Agency': 'Agency',
                            'Full_Name_Rus': 'Full_Name_Rus',
                            'Full_Name_Eng': 'Full_Name_Eng',
                            'For_Instrument': 'For_Instrument',
                            'For_Company': 'For_Company'}) #избавляемся от пробелов

data.to_excel('Ratings.xlsx', index=False) #экспорт данных

book = xlrd.open_workbook('Ratings.xlsx') #открываем книгу
sheet = book.sheet_by_name('Sheet1') #открываем лист

query = ("""
INSERT INTO [Rates].[dbo].[Ratings](
    ID,
    Code_Name, 
    Agency,
    Full_Name_Rus,
    Full_Name_Eng,
    For_Instrument, 
    For_Company) 
VALUES (?,?,?,?,?,?,?)""")

for r in range(1, sheet.nrows): #цикл построчного считывания данных
    ID = sheet.cell(r,0).value
    Code_Name = sheet.cell(r,1).value
    Agency = sheet.cell(r,2).value
    Full_Name_Rus = sheet.cell(r,3).value
    Full_Name_Eng = sheet.cell(r,4).value
    For_Instrument = sheet.cell(r,5).value
    For_Company = sheet.cell(r,6).value
    
    values = (ID,
              Code_Name, 
              Agency,
              Full_Name_Rus,
              Full_Name_Eng,
              For_Instrument, 
              For_Company)
    cursor.execute(query, values) #назначаем значения сторка/столбец
connection.commit() #сохраняем изменения
if cursor.rowcount > 0: #проверка импорта данных
    print('Import successful')
else:
    print('Import failed')
cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение
# -*- coding: utf-8 -*-
import pyodbc #подключаем библиотеку для работы с ODBC

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'TestBase' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД			  
cursor.execute('SELECT * FROM dbo.Table_1') #создаем запрос всех строк
for row in cursor:
    print('row = %r' % (row,)) #выводим все строки таблицы
cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение
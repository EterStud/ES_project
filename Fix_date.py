# -*- coding: utf-8 -*-
import pyodbc

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
cursor.execute("""UPDATE [dbo].[Base]
SET Date_str = CONVERT (datetime,convert(int,Date_str))""") 
cursor.commit()
cursor.execute("""UPDATE [dbo].[Base]
SET Date_str = CONVERT (date,convert(datetime,Date_str), 102)""")
cursor.commit()
cursor.close()
connection.close()
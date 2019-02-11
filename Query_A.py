# -*- coding: utf-8 -*-
import pyodbc #подключаем библиотеку для работы с ODBC

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
cursor.execute("""SELECT * INTO [dbo].[Query_Aa] 
               FROM [dbo].[Base] WHERE Date_str > N'2010-07-01'""") #выбираем поля в отдельную таблицу
connection.commit() #сохраняем изменения
cursor.execute("""SELECT * INTO [dbo].[Query_Ab] 
               FROM [dbo].[Base] WHERE Date_str <= '2010-07-01'
               AND Change !='cнят' AND Change !='Снят'""")
connection.commit() #сохраняем изменения
cursor.execute("""SELECT * INTO [dbo].[Query_A] FROM [dbo].[Query_Aa]
UNION ALL SELECT * FROM [dbo].[Query_Ab]""")
connection.commit() #сохраняем изменения
cursor.execute("""DROP TABLE [dbo].[Query_Aa]""")
connection.commit() #сохраняем изменения
cursor.execute("""DROP TABLE [dbo].[Query_Ab]""")
connection.commit() #сохраняем изменения
cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение
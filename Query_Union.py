# -*- coding: utf-8 -*-
import pyodbc #подключаем библиотеку для работы с ODBC

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
cursor.execute("""SELECT * INTO [dbo].[Query_Union] 
               FROM [dbo].[Query_A] WHERE 
               Ag_name <> N'АМБест Компани' AND 
               Ag_name <> N'Интерфакс' AND
               Rat_sctr = N'B' AND
               Scale_typer <> N'Isc' AND 
               Scale_typer <> N'UAsc' AND 
               Scale_typer <> N'KZsc'""") #выбираем поля в отдельную таблицу
connection.commit() #сохраняем изменения
cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение
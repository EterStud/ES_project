# -*- coding: utf-8 -*-
import pyodbc #подключаем библиотеку для работы с ODBC

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'TestBase' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes',
                            autocommit = True) #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД			  

cursor.execute('CREATE DATABASE Rates') #создаем запрос на создание БД
cursor.commit() #сохраняем изменения

cursor.execute('SELECT name FROM master.dbo.sysdatabases') #выполняем проверку
for db in cursor:
  print(db) #выводим список всех БД на сервере
cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение
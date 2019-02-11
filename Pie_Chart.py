# -*- coding: utf-8 -*-
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
user_input = input('Enter date in numeric value (YYYY-MM-DD)\n')
date = [user_input]
cursor.execute("""SELECT * INTO [dbo].[Data_{0}] 
               FROM [dbo].[Query_Union]
               WHERE Date_str = N'{0}'""".format(*date))
connection.commit()
cursor.execute("""Select Ag_name, COUNT(*) as 
               Count INTO [dbo].[Pie_{0}] 
               FROM [dbo].[Data_{0}] GROUP BY Ag_name""".format(*date))
connection.commit()
query = """SELECT * FROM [dbo].[Pie_{0}]""".format(*date)
data = pd.read_sql(query,connection)
cursor.execute(query)

df = pd.DataFrame.from_records(cursor.fetchall(), columns = [desc[0] for desc in cursor.description])
Data = pd.DataFrame(data)
print(data)
plt.figure(figsize=plt.figaspect(1))
labels=df['Ag_name']
sizes=df['Count']

def make_autopct(sizes):
    def my_autopct(pct):
        total = sum(sizes)
        val = int(round(pct*total/100.0))
        return '{p:.2f}%  ({v:d})'.format(p=pct,v=val)
    return my_autopct

plt.pie(sizes, counterclock=False, shadow = True, autopct=make_autopct(sizes), wedgeprops = { 'linewidth': 1.5, "edgecolor" :"k" })
plt.title("Данные рейтингов на {0}".format(*date))
plt.legend(labels,loc=1, bbox_to_anchor=(1.4, 1))
plt.show()

cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
ag_name = input('Enter Agency\n')
date_input = input('Enter date\n')
cursor.execute("""SELECT Ag_name, Date_str, Grade INTO [dbo].[Query_{1}] 
               FROM [dbo].[Base] WHERE Ag_name = N'{0}' 
               AND Date_str = N'{1}'""".format(ag_name, date_input))
connection.commit()
cursor.execute("""Select Ag_name, Grade, COUNT(*) as Count INTO [dbo].[Bar_{0}] FROM [dbo].[Query_{0}] GROUP BY Grade, Ag_name""".format(date_input))
connection.commit()
Query = """SELECT * FROM [dbo].[Bar_{0}]""".format(date_input)
data = pd.read_sql(Query,connection)
cursor.execute(Query)
df = pd.DataFrame.from_records(cursor.fetchall(), columns = [desc[0] for desc in cursor.description])
Data = pd.DataFrame(data)
x = df['Grade']
y = df['Count']
l = df['Ag_name']
x_pos = [i for i, _ in enumerate(x)]
plt.bar(x_pos, y)
plt.title("Оценки на {0}".format(date_input))
plt.legend(l, loc='best')
plt.xticks(x_pos, x)
plt.show()
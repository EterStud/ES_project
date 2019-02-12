import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
import os
import sys

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
ag_name = input('Enter Agency\n')
date_input = input('Enter date\n')
cursor.execute("""SELECT Ag_name, Date_str, Grade INTO [dbo].[Query_{1}] 
               FROM [dbo].[Query_Union] WHERE Ag_name = N'{0}' 
               AND Date_str <= N'{1}' 
               AND Change <>'cнят' AND Change <>'Снят'""".format(ag_name, date_input))
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
plt.title("Распределение рейтингов на {0}".format(date_input))
plt.legend(l, loc='best')
plt.xticks(x_pos, x)
print("\nDo you want to save bar chart in PNG format?")
decision_6 = input("Type Yes or No\n")
if decision_6.lower() == 'yes':
    print("\nDo you want to change image name?")
    decision_7 = input("Type Yes or No\n")
    print("\nDo you want to change image path?")
    decision_8 = input("Type Yes or No\n")
    if (decision_7.lower() == 'yes' and decision_8.lower() == 'yes'):
        bar_name = input("Enter image name\n")
        bar_path = input("Enter image path\n")
        assert os.path.exists(bar_path)
        plt.savefig((bar_path +"{0}.png".format(bar_name)), bbox_inches='tight')   
    elif (decision_7.lower() == 'yes' and decision_8.lower() == 'no'):
        bar_name = input("Enter image name\n")
        print("\nDefault image path - folder with scripts")
        plt.savefig("{0}.png".format(bar_name), bbox_inches='tight')   
    elif (decision_7.lower() == 'no' and decision_8.lower() == 'yes'):
        print("\nImage name as table name")
        bar_path = input("Enter new image path\n")
        assert os.path.exists(bar_path)
        plt.savefig((bar_path +"Bar_{0}.png".format(date_input)), bbox_inches='tight')    
    elif (decision_7.lower() == 'no' and decision_8.lower() == 'no'):
        print("\nImage name as table name")
        print("\nDefault image path - folder with scripts")
        plt.savefig("Bar_{0}.png".format(date_input), bbox_inches='tight')
    cursor.close()
    connection.close()
elif decision_6.lower() == 'no':
    sys.exit("Nothing to create as Image")
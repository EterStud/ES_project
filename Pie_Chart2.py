# -*- coding: utf-8 -*-
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
user_input = input('Enter date (YYYY-MM-DD)\n')
date = [user_input]
cursor.execute("""SELECT * INTO [dbo].[Data_{0}] 
               FROM [dbo].[Query_Union]
               WHERE Date_str <= N'{0}' 
               AND Change <>'cнят' AND Change <>'Снят'""".format(*date))
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

print("\nDo you want to save pie chart in PNG format?")
decision_3 = input("Type Yes or No\n")
if decision_3.lower() == 'yes':
    print("\nDo you want to change image name?")
    decision_4 = input("Type Yes or No\n")
    print("\nDo you want to change image path?")
    decision_5 = input("Type Yes or No\n")
    if (decision_4.lower() == 'yes' and decision_5.lower() == 'yes'):
        pie_name = input("Enter image name\n")
        pie_path = input("Enter image path\n")
        assert os.path.exists(pie_path)
        plt.savefig((pie_path +"{0}.png".format(pie_name)), bbox_inches='tight')   
    elif (decision_4.lower() == 'yes' and decision_5.lower() == 'no'):
        pie_name = input("Enter image name\n")
        print("\nDefault image path - folder with scripts")
        plt.savefig("{0}.png".format(pie_name), bbox_inches='tight')   
    elif (decision_4.lower() == 'no' and decision_5.lower() == 'yes'):
        print("\nImage name as table name")
        pie_path = input("Enter new image path\n")
        assert os.path.exists(pie_path)
        plt.savefig((pie_path +"Pie_{0}.png".format(*date)), bbox_inches='tight')    
    elif (decision_4.lower() == 'no' and decision_5.lower() == 'no'):
        print("\nImage name as table name")
        print("\nDefault image path - folder with scripts")
        plt.savefig("Pie_{0}.png".format(*date), bbox_inches='tight')
    cursor.close()
    connection.close()
elif decision_3.lower() == 'no':
    sys.exit("Nothing to create as Image")
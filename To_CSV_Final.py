# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
import pyodbc #подключаем библиотеку для работы с ODBC
import csv #подключаем библиотеку для работы с CSV
import sys #системный выход
import os #абсолютный путь к файлам

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения
print("List of tables in {0}:\n".format(database))
db_list = cursor.execute("""SELECT Distinct TABLE_NAME FROM information_schema.TABLES""")
table_list = [x[0] for x in db_list]
format_table_list = ', '.join([x[0] for x in db_list])
print(table_list)
print("\nDo you want to save data in CSV format?")
decision_1 = input("Type Yes or No\n")
if decision_1.lower() == 'yes':
    table_name = input("Enter table name for import\n")
    cursor.execute("""SELECT * FROM [dbo].[{0}]""".format(table_name))
    print("\nDo you want to change out file name?")
    decision_2 = input("Type Yes or No\n")
    print("\nDo you want to change file path?")
    decision_3 = input("Type Yes or No\n")
    if (decision_2.lower() == 'yes' and decision_3.lower() == 'yes'):
        csv_name = input("Enter file name\n")
        csv_path = input("Enter new file path\n")
        assert os.path.exists(csv_path)
        with open((csv_path +"{0}.csv".format(csv_name)), "w", newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)
    elif (decision_2.lower() == 'yes' and decision_3.lower() == 'no'):
        csv_name = input("Enter file name\n")
        print("\nDefault file path - folder with scripts")
        with open(("{0}.csv".format(csv_name)), "w", newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)
    elif (decision_2.lower() == 'no' and decision_3.lower() == 'yes'):
        print("\nFile name as table name")
        csv_path = input("Enter new file path\n")
        assert os.path.exists(csv_path)
        with open((csv_path +"{0}.csv".format(table_name)), "w", newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)
    elif (decision_2.lower() == 'no' and decision_3.lower() == 'no'):
        print("\nFile name as table name")
        print("\nDefault file path - folder with scripts")
        with open("{0}.csv".format(table_name), "w", newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)
    connection.commit()
elif decision_1.lower() == 'no':
    sys.exit("Nothing to create as CSV")
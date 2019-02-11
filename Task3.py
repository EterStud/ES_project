# -*- coding: utf-8 -*-
import pyodbc #подключаем библиотеку для работы с ODBC
import csv #подключаем библиотеку для работы с CSV
import pandas as pd
import matplotlib.pyplot as plt
import sys
from fpdf import FPDF
from pandas.tools.plotting import table

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
    print("Do you want to change out file name?")
    decision_2 = input("Type Yes or No\n")
    print("Do you want to change file path?")
    decision_3 = input("Type Yes or No\n")
    if decision_2.lower() == 'yes':
        if decision_3.lower() == 'yes':
            out_file_name = input("Enter file name\n")
            path_out_file = input("Enter new file path\n")
            with open(path_out_file +"{0}.csv".format(out_file_name), "w", newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow([i[0] for i in cursor.description])
                csv_writer.writerows(cursor)
        elif decision_3.lower() == 'no':
            print("Default file path - folder with scripts")
            with open("{0}.csv".format(out_file_name), "w", newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow([i[0] for i in cursor.description])
                csv_writer.writerows(cursor)
    elif decision_2.lower() == 'no':
            print("File name as table name")
            with open("{0}.csv".format(table_name), "w", newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow([i[0] for i in cursor.description])
                csv_writer.writerows(cursor)
    connection.commit()
elif decision_1.lower() == 'no':
    print("Do you want to plot in PNG format?")
    print("You need Pie_ or Bar_ table to use data")    
    decision_4 = input("Type Yes or No\n")
    if decision_4.lower() == 'yes':
        decision_5 = input("Type Pie or Bar\n")
        if decision_5.lower() == 'pie':
            date_input = input('Enter date Pie table\n')
            date = [date_input]
            query = """SELECT * FROM [dbo].[Pie_{0}]""".format(*date)
            data = pd.read_sql(query,connection)
            cursor.execute(query)
            df = pd.DataFrame.from_records(cursor.fetchall(), columns = [desc[0] for desc in cursor.description])
            Data = pd.DataFrame(data)
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
            print ("Do you want to change image name?")
            decision_6 = input("Type Yes or No\n")
            print("Do you want to change image path?")
            decision_7 = input("Type Yes or No\n")
            if decision_6.lower() == 'yes':
                if decision_7.lower() == 'yes':
                    pie_name = input('Enter image name\n')
                    pie_path = input('Enter image path\n')
                    plt.savefig(pie_path + "{0}.png".format(pie_name))
                elif decision_7.lower() == 'no':
                    print("Default file path - folder with scripts")
                    plt.savefig("{0}.png".format(pie_name))
            elif decision_6.lower() == 'no':        
                print("IMG file named by date")
                plt.savefig("{0}.png".format(date))
        elif decision_5.lower() == 'bar':
            date_input = input('Enter date Bar table\n')
            date = [date_input]
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
            print ("Do you want to change image name?")
            decision_8 = input("Type Yes or No\n")
            print("Do you want to change image path?")
            decision_9 = input("Type Yes or No\n")
            if decision_8.lower() == 'yes':
                if decision_9.lower() == 'yes':
                    bar_name = input('Enter image name\n')
                    bar_path = input('Enter image path\n')
                    plt.savefig(bar_path + "{0}.png".format(bar_name))
                elif decision_7.lower() == 'no':
                    print("Default file path - folder with scripts")
                    plt.savefig("{0}.png".format(bar_name))
            elif decision_6.lower() == 'no':        
                print("IMG file named by date")
                plt.savefig("{0}.png".format(date))
    if decision_4.lower() == 'no':
        sys.exit("Import/Saving data aborted!")
else:    
    print("\nDo you want to create a PDF file?")
    decision_n = input("Type Yes or No\n")
    if decision_n.lower() == 'yes':
        print ("Add table?\n")
        decision_10 = input('Type Yes or No\n')
    if decision_10.lower() == 'yes':
        Query = """SELECT * FROM [dbo].[{0}]""".format(table_name)
        data = pd.read_sql(Query,connection)
        cursor.execute(Query)
        df = pd.DataFrame.from_records(cursor.fetchall(), columns = [desc[0] for desc in cursor.description])
        Data = pd.DataFrame(data)
        ax2 = plt.subplot(222)
        plt.axis('off')
        tb = table(ax2, df, loc='center')
        tb.auto_set_font_size(True)
        plt.savefig("{0}.png".format(table_name))
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.add_page()
        pdf.add_font('FreeSans', '', 'FreeSans.ttf', uni=True)
        pdf.set_font("FreeSans")
        pdf.cell(200, 10, txt="Отчет", ln=1, align="C")
        pdf.ln(10)
        pdf.cell(200, 10, txt="""Данный отчет содержит данные об актуальных рейтингах на {0}""".format(*date), ln=1, align="C")
        pdf.ln(10)
        pdf.image("{0}.png".format(table_name), x = 80, y = 50, w = 60, h = 60, type = '', link = '')
        decision_11 = input("pie or bar?\n")
        if decision_11.lower() == 'pie':
            pdf.image("{0}.png".format(pie_name), x = 80, y = 50, w = 60, h = 60, type = '', link = '')
            pdf.ln(40)
            pdf.cell(200, 10, txt="""На графике представлены актуальные на {0} рейтинги по агенствам""".format(*date), ln=1, align="C")
            pdf.output("simple_demo.pdf")
        elif decision_11.lower() == 'bar':
            pdf.image("{0}.png".format(bar_name), x = 80, y = 50, w = 60, h = 60, type = '', link = '')
            pdf.ln(40)
            pdf.cell(200, 10, txt="""На графике представлены актуальные оценки на {0} по агенствам""".format(*date), ln=1, align="C")
            pdf.output("simple_demo.pdf")
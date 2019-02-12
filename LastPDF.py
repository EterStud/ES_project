# -*- coding: utf-8 -*-
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import sys
import os
import numpy as np
import six 

server = 'DESKTOP-HP138CN\SQLEXPRESS' #имя сервера
database = 'Rates' #имя базы данных


connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            +' ;Trusted_Connection=yes') #создание нового подключения
cursor = connection.cursor() #создаем специальный объект для выполнения запросов к БД
print ("Do you want to create PDF for SQL Query?")
decision_1 = input('Type Yes or No\n')
if decision_1.lower() == 'yes':
    print("List of tables in {0}:\n".format(database))
    db_list = cursor.execute("""SELECT Distinct TABLE_NAME FROM information_schema.TABLES""")
    table_list = [x[0] for x in db_list]
    format_table_list = ', '.join([x[0] for x in db_list])
    print(table_list)
    table_name = input("Enter table name for import\n")
    Query = """SELECT * FROM [dbo].[{0}]""".format(table_name)
    data = pd.read_sql(Query,connection)
    cursor.execute(Query)
    df = pd.DataFrame.from_records(cursor.fetchall(), columns = [desc[0] for desc in cursor.description])
    Data = pd.DataFrame(data)
    print(data)
    def render_mpl_table(data, col_width=3.0, row_height=0.625, font_size=14,
                         header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w',
                         bbox=[0, 0, 1, 1], header_columns=0,
                         ax=None, **kwargs):
        if ax is None:
            size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
            fig, ax = plt.subplots(figsize=size)
            ax.axis('off')   
        mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs) 
        mpl_table.auto_set_font_size(False)
        mpl_table.set_fontsize(font_size)     
        for k, cell in  six.iteritems(mpl_table._cells):
            cell.set_edgecolor(edge_color)
            if k[0] == 0 or k[1] < header_columns:
                cell.set_text_props(weight='bold', color='w')
                cell.set_facecolor(header_color)
            else:
                cell.set_facecolor(row_colors[k[0]%len(row_colors) ])
        return ax
    render_mpl_table(df, header_columns=0, col_width=2.0)
    
    print("\nDo you want to change out file name?")
    decision_2 = input("Type Yes or No\n")
    print("\nDo you want to change file path?")
    decision_3 = input("Type Yes or No\n")
    if (decision_2.lower() == 'yes' and decision_3.lower() == 'yes'):
        tb_name = input("Enter table name\n")
        tb_path = input("Enter table path\n")
        assert os.path.exists(tb_path)
        plt.savefig((tb_path +"Tb_{0}.png".format(tb_name)), bbox_inches='tight')   
    elif (decision_2.lower() == 'yes' and decision_3.lower() == 'no'):
        tb_name = input("Enter image name\n")
        print("\nDefault image path - folder with scripts")
        plt.savefig("Tb_{0}.png".format(tb_name), bbox_inches='tight')   
    elif (decision_2.lower() == 'no' and decision_3.lower() == 'yes'):
        print("\nTable name as table name in DB")
        tb_path = input("Enter new image path\n")
        assert os.path.exists(tb_path)
        plt.savefig((tb_path +"Tb_{0}.png".format(table_name)), bbox_inches='tight')    
    elif (decision_2.lower() == 'no' and decision_3.lower() == 'no'):
        print("\nTable name as table name in DB")
        print("\nDefault image path - folder with scripts")
        plt.savefig("Tb_{0}.png".format(table_name), bbox_inches='tight')
    print("Pie or Bar")
    decision_4 = input('Type Yes or No\n')
    if decision_4.lower() == 'pie':
        user_input = input('Enter date (YYYY-MM-DD)\n')
        date = [user_input]
        cursor.execute("""SELECT * INTO [dbo].[Data_{0}] 
                       FROM [dbo].[Query_A]
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
        
        print("\nDo you want to save pie chart in PNG format?")
        decision_6 = input("Type Yes or No\n")
        if decision_6.lower() == 'yes':
            print("\nDo you want to change image name?")
            decision_7 = input("Type Yes or No\n")
            print("\nDo you want to change image path?")
            decision_8 = input("Type Yes or No\n")
            if (decision_7.lower() == 'yes' and decision_8.lower() == 'yes'):
                pie_name = input("Enter image name\n")
                pie_path = input("Enter image path\n")
                assert os.path.exists(pie_path)
                plt.savefig((pie_path +"{0}.png".format(pie_name)), bbox_inches='tight')   
            elif (decision_7.lower() == 'yes' and decision_8.lower() == 'no'):
                pie_name = input("Enter image name\n")
                print("\nDefault image path - folder with scripts")
                plt.savefig("{0}.png".format(pie_name), bbox_inches='tight')   
            elif (decision_7.lower() == 'no' and decision_8.lower() == 'yes'):
                print("\nImage name as table name")
                pie_path = input("Enter new image path\n")
                assert os.path.exists(pie_path)
                plt.savefig((pie_path +"Pie_{0}.png".format(*date)), bbox_inches='tight')    
            elif (decision_7.lower() == 'no' and decision_8.lower() == 'no'):
                print("\nImage name as table name")
                print("\nDefault image path - folder with scripts")
                plt.savefig("Pie_{0}.png".format(*date), bbox_inches='tight')
    
        elif decision_3.lower() == 'no':
            sys.exit("Nothing to create as Image")
            
    elif decision_4.lower() == 'bar':
        ag_name = input('Enter Agency\n')
        date_input = input('Enter date\n')
        cursor.execute("""SELECT Ag_name, Date_str, Grade INTO [dbo].[Query_{1}] 
                       FROM [dbo].[Query_Union] WHERE Ag_name = N'{0}' 
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
        plt.title("Распределение рейтингов на {0}".format(date_input))
        plt.legend(l, loc='best')
        plt.xticks(x_pos, x)
        print("\nDo you want to save bar chart in PNG format?")
        decision_9 = input("Type Yes or No\n")
        if decision_9.lower() == 'yes':
            print("\nDo you want to change image name?")
            decision_10 = input("Type Yes or No\n")
            print("\nDo you want to change image path?")
            decision_11 = input("Type Yes or No\n")
            if (decision_10.lower() == 'yes' and decision_11.lower() == 'yes'):
                bar_name = input("Enter image name\n")
                bar_path = input("Enter image path\n")
                assert os.path.exists(bar_path)
                plt.savefig((bar_path +"{0}.png".format(bar_name)), bbox_inches='tight')   
            elif (decision_10.lower() == 'yes' and decision_11.lower() == 'no'):
                bar_name = input("Enter image name\n")
                print("\nDefault image path - folder with scripts")
                plt.savefig("{0}.png".format(bar_name), bbox_inches='tight')   
            elif (decision_10.lower() == 'no' and decision_11.lower() == 'yes'):
                print("\nImage name as table name")
                bar_path = input("Enter new image path\n")
                assert os.path.exists(bar_path)
                plt.savefig((bar_path +"Bar_{0}.png".format(date_input)), bbox_inches='tight')    
            elif (decision_10.lower() == 'no' and decision_11.lower() == 'no'):
                print("\nImage name as table name")
                print("\nDefault image path - folder with scripts")
                plt.savefig("Bar_{0}.png".format(date_input), bbox_inches='tight')   
    
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.add_font('FreeSans', '', 'FreeSans.ttf', uni=True)
    pdf.set_font("FreeSans")
    pdf.cell(200, 10, txt="Отчет", ln=1, align="C")
    pdf.ln(10)
    pdf.cell(200, 10, txt="""Данный отчет содержит данные об актуальных рейтингах""", ln=1, align="C")
    pdf.ln(40)
    
    if (decision_2.lower() == 'yes' and decision_3.lower() == 'yes'):
        pdf.image((tb_path +"Tb_{0}.png".format(tb_name)), x = 60, y = 50, w = 80, h = 40, type = '', link = '')    
    elif (decision_2.lower() == 'yes' and decision_3.lower() == 'no'):
        pdf.image("Tb_{0}.png".format(tb_name), x = 60, y = 50, w = 80, h = 40, type = '', link = '')
    elif (decision_2.lower() == 'no' and decision_3.lower() == 'yes'):
        pdf.image((tb_path +"Tb_{0}.png".format(table_name)), x = 60, y = 50, w = 80, h = 40, type = '', link = '')   
    elif (decision_2.lower() == 'no' and decision_3.lower() == 'no'):
        pdf.image("Tb_{0}.png".format(table_name), x = 60, y = 50, w = 80, h = 40, type = '', link = '')
    
    pdf.ln(10)
    if decision_4.lower() == 'pie':     
        pdf.cell(200, 10, txt=("""На диаграмме изображено распределение рейтингов, актуальных на {0}""".format(*date)), ln=1, align="C")
        pdf.ln(10)
        if (decision_7.lower() == 'yes' and decision_8.lower() == 'yes'):
            pdf.image((pie_path +"{0}.png".format(pie_name)), x = 80, y = 120, w = 60, h = 60, type = '', link = '')
        elif (decision_7.lower() == 'yes' and decision_8.lower() == 'no'):
            pdf.image("{0}.png".format(pie_name), x = 80, y = 120, w = 60, h = 60, type = '', link = '') 
        elif (decision_7.lower() == 'no' and decision_8.lower() == 'yes'):
            pdf.image((pie_path +"Pie_{0}.png".format(*date)), x = 80, y = 120, w = 60, h = 60, type = '', link = '') 
        elif (decision_7.lower() == 'no' and decision_8.lower() == 'no'):
            pdf.image("Pie_{0}.png".format(*date), x = 80, y = 120, w = 60, h = 60, type = '', link = '') 
    elif decision_4.lower() == 'bar': 
        pdf.cell(200, 10, txt=("""На графике изображено распределение рейтинга агентства, актуальных на {0}""".format(*date)), ln=1, align="C")
        pdf.ln(10)
        if (decision_10.lower() == 'yes' and decision_11.lower() == 'yes'):
            pdf.image((bar_path +"{0}.png".format(bar_name)), x = 80, y = 120, w = 60, h = 60, type = '', link = '')
        elif (decision_10.lower() == 'yes' and decision_11.lower() == 'no'):
            pdf.image("{0}.png".format(bar_name), x = 80, y = 120, w = 60, h = 60, type = '', link = '')
        elif (decision_10.lower() == 'no' and decision_11.lower() == 'yes'):
            pdf.image((bar_path +"Bar_{0}.png".format(date_input)), x = 80, y = 120, w = 60, h = 60, type = '', link = '')  
        elif (decision_10.lower() == 'no' and decision_11.lower() == 'no'):
            pdf.image("Bar_{0}.png".format(date_input), x = 80, y = 120, w = 60, h = 60, type = '', link = '')
    pdf.output("simple_demo.pdf")
elif decision_1.lower() == 'no':
    sys.exit("Stop creating PDF")
cursor.close() #удаляем специальный объект
connection.close() #закрываем соединение

        
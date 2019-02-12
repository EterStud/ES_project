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
        pdf.output("simple_demo.pdf")

    elif (decision_2.lower() == 'yes' and decision_3.lower() == 'no'):
        pdf.image("Tb_{0}.png".format(tb_name), x = 60, y = 50, w = 80, h = 40, type = '', link = '')
        pdf.output("simple_demo.pdf")

    elif (decision_2.lower() == 'no' and decision_3.lower() == 'yes'):
        pdf.image((tb_path +"Tb_{0}.png".format(table_name)), x = 60, y = 50, w = 80, h = 40, type = '', link = '')   
        pdf.output("simple_demo.pdf")

    elif (decision_2.lower() == 'no' and decision_3.lower() == 'no'):
        pdf.image("Tb_{0}.png".format(table_name), x = 60, y = 50, w = 80, h = 40, type = '', link = '')    
        pdf.output("simple_demo.pdf")
# -----------------------------------------------------------------------
# pip install xlsxwriter, openpyxl, pandas
# -----------------------------------------------------------------------

import os
import csv
import glob    

import pandas as pd
from pandas.io.excel import ExcelWriter
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
#import xlrd

#all_files = ["csv1.csv", "csv5.csv"];
all_files = ["csv3.csv", "csv4.csv"];

df_from_each_file = (pd.read_csv(f) for f in all_files)

writer = pd.ExcelWriter("result.xlsx", engine='xlsxwriter');
for idx, df in enumerate(df_from_each_file):
	df.to_excel(writer, sheet_name='Sheet {0}'.format(idx), index=False, encoding='utf-8')
writer.save()



book = load_workbook('result.xlsx')
for i in range(2, book["Sheet 0"].max_row+1):
#	print(".");
	book["Sheet 0"]['C'+str(i)] = i
	book["Sheet 0"]['D'+str(i)] = i+1
	book["Sheet 0"]['E'+str(i)] = '=SUM(C'+str(i)+',D'+str(i)+')'

dv = DataValidation(type="list", formula1='"Dog,Cat,Bat"', allow_blank=True)
book["Sheet 0"].add_data_validation(dv)
dv.add('F2:F'+str(book["Sheet 0"].max_row+1));

book["Sheet 0"].column_dimensions['F'].width = 100

book.save('result.xlsx')




import pandas as pd
from datetime import date, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
import pyarrow as pa
import os
from src.DailyReports.uitls import PosCash, eMO, ePost, write_file

#file_path = r'C:/Users/Administrator/Downloads/EXPORT (8).xlsx'
file_path =  str(input("Enter file_path: "))
#data = pd.read_excel(file_path)
#data, file_name = PosCash(file_path)
data, file_name = eMO(file_path)
#data, file_name = ePost(file_path)

write_file(data, file_path = 'Reports', file_name=file_name)
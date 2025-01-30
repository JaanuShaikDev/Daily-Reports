import pandas as pd
from datetime import date, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
import jinja2
import os
from src.DailyReports.uitls import PosCash, eMO, ePost, write_file

#file_path =  str(input("Enter file_path: "))
data, file_name = PosCash()
data, file_name = eMO()
data, file_name = ePost()

#write_file(data, file_path = 'Reports', file_name=file_name)
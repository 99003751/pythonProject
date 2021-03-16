import os
import pandas as pd
from xlsxwriter import Workbook


xls = pd.ExcelFile('Employee_data1.xlsx')
df = pd.read_excel('Employee_data1.xlsx', 'Sheet1')
print(df)

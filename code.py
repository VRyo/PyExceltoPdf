# Import Library
import os
from win32com import client

# Opening Microsoft Excel
excel = client.Dispatch("Excel.Application")

# Read Excel File
data = os.listdir('sources')

for i in data:
    sheets = excel.Workbooks.Open(os.path.abspath('sources/' + i))
    work_sheets = sheets.Worksheets[0]

    # Converting into PDF File
    work_sheets.ExportAsFixedFormat(0, os.path.abspath('output/' + os.path.splitext(i)[0] + '.pdf'))
import win32com.client
from pywintypes import com_error
from tkinter.filedialog import askopenfilenames 
from pathlib import PurePath

o = win32com.client.Dispatch("Excel.Application")
wb_path = askopenfilenames(title = "Select Excel file")
o.Visible = False
i = 0

try:
    wb = o.Workbooks.Open(PurePath(wb_path[i]))

    ws_index_list = [1,2,3] # Select the sheet that you want to print
    path_to_pdf = (r"c:/Users/cmeneses/Desktop/sample.pdf")
    wb.WorkSheets(ws_index_list).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)

    ws_index_list = [4,5,6] # Select the sheet that you want to print
    path_to_pdf = (r"c:/Users/cmeneses/Desktop/sample1.pdf")
    wb.WorkSheets(ws_index_list).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)

    i += i
except com_error as e:
    print('failed.')
else:
    print('Succeeded.')
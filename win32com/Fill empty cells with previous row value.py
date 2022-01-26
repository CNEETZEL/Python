import win32com.client as win32
import os
win32c = win32.constants

fname = r"C:\folder\file.xls"

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
ws = wb.ActiveSheet

begrow = 2
endrow = ws.UsedRange.Rows.Count
for row in range(begrow,endrow+1): 
  if ws.Range('A{}'.format(row)).Value is None:
    ws.Range('A{}'.format(row)).Value = excel.ActiveCell.Offset(2,1).Value

wb.SaveAs(fname.rsplit('.', 1)[0]+".xlsx", FileFormat:=61)
wb.Close()
excel.Application.Quit()

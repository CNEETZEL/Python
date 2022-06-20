import win32com
import win32com.client as win32
win32c = win32.constants
import os

Input = os.path.split(os.path.split(os.path.realpath(__file__))[0])[-2] + '\\Input\\CDC.xls'

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(Input)
ws = wb.ActiveSheet

#Last Row from indicated column
#Ultima fila indicando la columna donde se contara
def LastRow(WS,COL):
  return WS.Cells(WS.Rows.Count, COL).End(win32c.xlShiftUp).Row
print(LastRow(ws,"A"))


#Last Column whole range
#Ultima columna en todo el rango
def LastCol(WS):
  return WS.UsedRange.Columns.Count
print(LastCol(ws))

#Last Column from indicated range
#Ultima columna indicando el rango donde se comienza a contar
def LC(WS,RNG):
  return WS.Range(RNG).End(win32c.xlToRight).Column
print(LC(ws,"A2"))
 
#Last Column letter from indicated range
#Letra de la última columna indicando el rango donde se comienza a contar
def LCL(WS,RNG):
  LastColumn = WS.Range(RNG).End(win32c.xlToRight).Column
  ADDRESS = WS.Cells(1,LastColumn).Address 
  start = '$'
  end = '$'
  return ADDRESS.split(start)[1].split(end)[0]
print(LCL(ws,"A2"))

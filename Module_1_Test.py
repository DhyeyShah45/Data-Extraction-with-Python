import pytest
from  Module1 import *
wb = openpyxl.load_workbook("C:/Users/manoj/PycharmProjects/Performance/Mark_sheet.xlsx")
def test_readdata():
    assert ReadExcelData().readData(wb,5,2,'Applied_SDLC')==['Venkata sai yamini  Thirluka',50.60000000000001,2]
    assert ReadExcelData().readData(wb,10,3,'Adv_Python')==[99003183,74,3]
def test_searchbyname():
    assert SearchExcelData().searchByName(wb,'Sushma  S  M','Adv_Python')==[26,3,'Adv_Python']
    assert SearchExcelData().searchByName(wb, 'Rithesh R Prabhu', 'MBSE') == [15, 3, 'MBSE']
def test_searchbypsno():
    assert SearchExcelData().searchByPsno(wb,99003186,'Adv_Python')==[13,2,'Adv_Python']
    assert SearchExcelData().searchByPsno(wb, 99003184, 'Applied_SDLC') == [11, 2, 'Applied_SDLC']
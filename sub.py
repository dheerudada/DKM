import os
import openpyxl
import xlrd
import pyexcel as p
from cpy_pst import copyRange, pasteRange
from openpyxl.chart import LineChart,Reference, ScatterChart, Series


def where(mask):
    msk=mask
    if mask[0:2]==('GT'):
        tec ='FGT'
    elif mask[0:2]==('GQ'):
        tec ='FGQ'
    elif mask[0:2]==('GF'):
        tec ='FPN'
    else:
        technology ='FGN'
    return tec

def clearsheet():
    #Clear results file
    template = openpyxl.load_workbook(f'C:/Users/dmohata/PycharmProjects/HelloWorld/tes.xlsx')
    getsheets=template.worksheets
    print(getsheets)
    #print(template.get_sheet_names())

    std=template['Sheet2']
    template.remove(std)
    #print(template.get_sheet_names())
    template.save(f'C:/Users/dmohata/PycharmProjects/HelloWorld/tes.xlsx')
    #####clear results file done
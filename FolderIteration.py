import os
from os import listdir
import numpy as np
import pandas as dp
from xlrd import open_workbook
import xlwt, xlrd
from xlutils.copy import copy





def get_sd(path):
    return filter(os.path.isdir, [os.path.join(path,f) for f in os.listdir(path)])

def createExcelSheet(fileName, destinationPath):
    newXLfile = xlwt.Workbook(encoding="utf-8")
    sheet = newXLfile.add_sheet("Finger Measurement Data")
    compiledDataDestination = str(destinationPath) + '\\' + str(fileName) + '.xls'
    newXLfile.save(compiledDataDestination)
    return compiledDataDestination

def rev(s):
    return ''.join(reversed(s))
    

def excelOperations(filePath, subfolder, compiledDataDestination, row):
    try:        
        wb = xlrd.open_workbook(filePath)
        sheetNamesArray = wb.sheet_names()
        workingSheet = wb.sheet_by_name(sheetNamesArray[0])


        index = [[workingSheet.cell_value(ri,ci) for ri in range(1,12)] for ci in range(1,2)]
        middle = [[workingSheet.cell_value(rm, cm) for rm in range(1,12)] for cm in range(2,3)]
        ring = [[workingSheet.cell_value(rr,cr) for rr in range(1,12)] for cr in range(3,4)]
        little = [[workingSheet.cell_value(rl, cl) for rl in range(1,12)] for cl in range(4,5)]

        

        subfolderName = ''
        for c in reversed(subfolder):
            if (c == '\\'):
                break
            else:
                subfolderName = subfolderName + c
        subfolderName = rev(subfolderName)




        savedFile = open_workbook(compiledDataDestination, "rb")
        savedFile_sheet = savedFile.sheet_by_index(0)
        writingFile = copy(savedFile)
        writingFile_sheet = writingFile.get_sheet(0)
        writingFile_sheet.write(row,0, subfolderName)


        writingFile_sheet.write_merge(0,0,2,13, 'Index', xlwt.easyxf("align: horiz center"))

        for c in range(2,13):
            data = index[0][c-2]
            writingFile_sheet.write(row,c, data )
            

        writingFile_sheet.write_merge(0,0,14,25, 'Middle', xlwt.easyxf("align:horiz center"))
        for d in range(14,25):
            data = middle[0][d-14]
            writingFile_sheet.write(row,d,data)


        writingFile_sheet.write_merge(0,0,26,37, 'Ring' , xlwt.easyxf("align:horiz center"))
        for e in range(26,37):
            data = ring[0][e-26]
            writingFile_sheet.write(row,e,data)

        writingFile_sheet.write_merge(0,0,38,49, 'Little', xlwt.easyxf("align:horiz center"))
        for f in range(38, 49):
            data = little[0][f-38]
            writingFile_sheet.write(row,f,data)


        
        writingFile.save(compiledDataDestination)
        
        
        
    except FileNotFoundError:
        print("That is an invalid path")
    

def main():
    cPath = os.path.dirname(os.path.abspath(__file__))
    row = 1
    compiledDataDestination = createExcelSheet('dataFinal', cPath)
    subdir = get_sd(cPath)
    print (cPath)
    print(compiledDataDestination)
    for folder in subdir:
        subsubdir = get_sd(folder)

        for subfolder in subsubdir:
            directory = os.listdir(subfolder)
            for files in directory:
                if files.endswith('.xls') or files.endswith('.xlsx'):
                    filePath = str(subfolder) + "\\" + str(files)
                    excelOperations(filePath, subfolder, compiledDataDestination, row)
                    row = row+1
            

main()

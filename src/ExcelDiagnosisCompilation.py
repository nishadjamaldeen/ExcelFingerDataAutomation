import xlrd, xlwt, os
from os import listdir
from xlrd import open_workbook
from openpyxl import load_workbook, Workbook
from xlutils.copy import copy as copyxl
import pandas as pd


newFileWB = Workbook()




def main():

    cPath = os.path.dirname(os.path.abspath(__file__))
    sheet = 1
    compiledDataDestination = createExcelSheet('diagnosisCompilation', cPath)
    newFileWB.save(compiledDataDestination)

    
    subdir = get_sd(cPath)

    for subfolder in subdir:
        directory = os.listdir(subfolder)
        for files in directory:
            if files.endswith('_diagnosis.xls') or files.endswith('_diagnosis.xlsx'):
                diagnosis_filePath = str(subfolder) + "\\" + str(files)
				    
        excelOperations(diagnosis_filePath, subfolder, compiledDataDestination, sheet)
        sheet = sheet + 1

def get_sd(path):
    return filter(os.path.isdir, [os.path.join(path,f) for f in os.listdir(path)])
	
def rev(s):
    return ''.join(reversed(s))

def createExcelSheet(fileName, destinationPath):

    compiledDataDestination = str(destinationPath) + '\\' + str(fileName) + '.xlsx'
    return compiledDataDestination
	
	
def excelOperations(diagnosis_filePath, subfolder, compiledDataDestination, sheet):
    try:

        wb = xlrd.open_workbook(diagnosis_filePath)
        sheetNamesArray = wb.sheet_names()
        workingSheet = wb.sheet_by_name(sheetNamesArray[0])
		
        data = pd.read_excel(diagnosis_filePath)
        data = data.drop([0])
        dataMatrix = data.as_matrix()
        
		
		
        subfolderName = ''
        for c in reversed(subfolder):
            if (c == '\\'):
                break
            else:
                subfolderName = subfolderName + c
        
        subfolderName = rev(subfolderName)


        savedFile = load_workbook(compiledDataDestination)
        savedFile_sheet = savedFile.create_sheet(subfolderName)


        for r in range(0,8):
            for c in range(0,5):
                savedFile_sheet.cell(column = c+1, row = r+1, value = dataMatrix[r][c])
            
        

        savedFile.save(compiledDataDestination)
        
        
    except FileNotFoundError:
        print("That is an invalid path")

main()
    

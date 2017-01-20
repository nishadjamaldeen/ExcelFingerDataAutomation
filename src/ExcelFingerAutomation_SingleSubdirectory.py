#####################################################################
#Program Name: Excel Data Automation (for Hand Arthritis Algorithm) #
#Author: Nishad KeyError                                            #   
#Location: eTreatMD Downtown                                        #
#Date Created: 18 January                                           #   
#Date Mofified: 19 January                                          #
#Description: Searches through a folder tree for ourput data from   #
#   the MATLAB algorithm, compiling all avaialible information      #
#   into a single file for analysis and presentation                #
#####################################################################



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

    measurementHeadings = ['Pip Joint Angle', 'Dip Joint Angle', 'Finger Length', 'Distal Midpoint Thickness',
                            'Dip Joint Thickness', 'Middle Midpoint Thickness', 'Pip Joint Thickness',
                            'Proximal Midpoint Thickness', 'Proximal Length', 'Middle Length', 'Distal Length']
    deformityHeadings = ['Index Dip', 'Index Pip', 'Middle Dip', 'Middle Pip', 'Ring Dip', 'Ring Pip',
                            'Little Dip', 'Little Pip']

    newXLfile = xlwt.Workbook(encoding="utf-8")
    sheet = newXLfile.add_sheet("Finger Measurement Data")
    fingerCount = 4;
    range_low = 2;
    range_high = 13;

    while fingerCount > 0:
        for c in range(range_low, range_high):
            sheet.write(1,c,measurementHeadings[c-range_low])
        fingerCount-=1
        range_low +=12
        range_high +=12

    

    for c in range(50,58):
        sheet.write(1,c,deformityHeadings[c-50])
   
    compiledDataDestination = str(destinationPath) + '\\' + str(fileName) + '.xls'
    newXLfile.save(compiledDataDestination)
    return compiledDataDestination


def rev(s):
    return ''.join(reversed(s))
    


def excelOperations(measurementFilePath, deformityFilePath, subfolder, compiledDataDestination, row):
    try:

        wb = xlrd.open_workbook(measurementFilePath)
        sheetNamesArray = wb.sheet_names()
        workingSheet = wb.sheet_by_name(sheetNamesArray[0])


        index = [[workingSheet.cell_value(ri,ci) for ri in range(1,12)] for ci in range(1,2)]
        middle = [[workingSheet.cell_value(rm, cm) for rm in range(1,12)] for cm in range(2,3)]
        ring = [[workingSheet.cell_value(rr,cr) for rr in range(1,12)] for cr in range(3,4)]
        little = [[workingSheet.cell_value(rl, cl) for rl in range(1,12)] for cl in range(4,5)]
		
        wb_deformity = xlrd.open_workbook(deformityFilePath)
        sheetNamesArray_deformity = wb.sheet_names()
        deformityWorkingSheet = wb_deformity.sheet_by_name(sheetNamesArray_deformity[0])

        deformityMatrix = [[deformityWorkingSheet.cell_value(ri,ci) for ci in range(0,1)] for ri in range(0,8)]
        print(deformityMatrix)
        print(deformityMatrix[1][0])
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
        writingFile_sheet.write(row+1,0, subfolderName)


        
        for c in range(2,13):
            data = index[0][c-2]
            writingFile_sheet.write(row+1,c, data )
            
        for d in range(14,25):
            data = middle[0][d-14]
            writingFile_sheet.write(row+1,d,data)
        
        for e in range(26,37):
            data = ring[0][e-26]
            writingFile_sheet.write(row+1,e,data)

        for f in range(38, 49):
            data = little[0][f-38]
            writingFile_sheet.write(row+1,f,data)	
        
        for g in range(50,58):
            data = deformityMatrix[g-50][0]
            writingFile_sheet.write(row+1,g,data)

        writingFile_sheet.write_merge(0,0,2,13,'Index', xlwt.easyxf("align: horiz center"))
        writingFile_sheet.write_merge(0,0,14,25,'Middle', xlwt.easyxf("align: horiz center"))
        writingFile_sheet.write_merge(0,0,26,37,'Ring', xlwt.easyxf("align: horiz center"))
        writingFile_sheet.write_merge(0,0,38,49,'Little', xlwt.easyxf("align: horiz center"))
        writingFile_sheet.write_merge(0,0,50,58, 'Deformities', xlwt.easyxf("align:horiz center"))

        writingFile.save(compiledDataDestination)
        
        
    except FileNotFoundError:
        print("That is an invalid path")
    

def main():
    cPath = os.path.dirname(os.path.abspath(__file__))
    row = 1
    compiledDataDestination = createExcelSheet('dataFinal', cPath)
    
    print (cPath)
    print(compiledDataDestination)
    
    subdir = get_sd(cPath)

    for subfolder in subdir:
        directory = os.listdir(subfolder)
        for files in directory:
            if files.endswith('fingerMeasurements.xls') or files.endswith('fingerMeasurements.xlsx'):
                measurement_filePath = str(subfolder) + "\\" + str(files)
                
            elif files.endswith('deformity.xls') or files.endswith('deformity.xlsx'):
                deformity_filePath = str(subfolder) + '\\' + str(files)
				    
        excelOperations(measurement_filePath, deformity_filePath, subfolder, compiledDataDestination, row)
        row = row+1

main()

#https://automatetheboringstuff.com
import openpyxl
import xlsxwriter   #https://pypi.python.org/pypi/XlsxWriter / http://xlsxwriter.readthedocs.org


colNum = 0   #column a ltter of spreadsheet
rowNum = 0   #row number of spreadsheet
playerFirstName = ''
playerLastName = ''
fdFPPG = 0   #FD fanatsy points per game
fdSalary = 0   #FD salary amount
projectedValue = 0  #FD projected value based on FPPG / Salary

fdWB = openpyxl.load_workbook('FDStats.xlsx')   #read FD player stat spreadsheet
fdSheet = fdWB.get_active_sheet()
playerFirstName = fdSheet['C2']
playerLastName = fdSheet['D2']
fdFPPG = fdSheet['E2']
fdSalary = fdSheet['G2']
projectedValue = (fdFPPG.value / fdSalary.value) * 1000

dfsWB = xlsxwriter.Workbook('DFStatsTest.xlsx')   #create xlsx file
dfsSheet = dfsWB.add_worksheet('DFSTest')
dfsSheet.write(rowNum, colNum, projectedValue)



dfsWB.close()

#OLD CODE, HOLDING JUST IN CASE
"""
import webbrowser
import requests
import linecache   #https://docs.python.org/2/library/linecache.html
import bs4
from selenium import webdriver
lineNum = 594  #line number of where player stats starts
lineEndNum = 969   #line where the player stats end
resSite = requests.get('http://rotoguru1.com/cgi-bin/fyday.pl?week=5&game=fd&scsv=1')   #open/download website text info
dfFile = open('DFStats.txt', 'wb')   #open text file in binary write mode
for chunk in resSite.iter_content(10000):   #write text info to file
    dfFile.write(chunk)
dfFile.close()
while lineNum <= lineEndNum:
    testLine = linecache.getline('DFStats.txt', lineNum).replace(';', ' ').split()  #get starting line for player stats and store in list, replace ; and then split the white space.
    print(testLine)
    lineNum += 1
    for i in range(len(testLine)):
        sheet.write(rowNum, colNum, testLine[i])   #write data from testLine into correct column
        colNum += 1
    colNum = 0   # reset colNum on newline
    rowNum +=1
"""

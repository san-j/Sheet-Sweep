# -*- coding: utf-8 -*-
"""
Created on Fri Oct  4 21:43:41 2019
SheetSweep: a Python application for reading csv data and editing Google Sheets,
primarily for HC Ally attendance management
@author: Sanjay Goyal
"""
import pygsheets
import csv

def markCell(rowNum, column, wks):
    cellID = column + str(rowNum)
    wks.update_value(cellID, 'x')
    
client = pygsheets.authorize(service_file='credentials.json')
sheet = client.open("Copy of FYM Attendance Record")
wks = sheet.worksheet('index',0)

names = client.get_range("1TQ0siChZ7kG3K6vDWLr_l1epbU9i_GD5d2YQCi_hrsQ", "A2:A60", "ROWS")

file = input("Enter filepath to Excel file: ")
column = input("Enter the column of the Drive sheet to edit: ")
with open(file, 'r') as csvfile:
    reader = csv.reader(csvfile)
    
    #Skip past header information
    for x in range(6):
        next(reader)
    
    rowNum = 0
    for row in reader:
        s = row[0] + " " + row[1]
        if s == 'Megan Maniar':
            rowNum = 39
        elif s == 'Nivedita Rajiv':
            rowNum = 45
        elif s == 'Shanrui Yu':
            rowNum = 53
        elif s == 'Tianqi Li':
            rowNum = 57
        elif [s] in names:
            rowNum = names.index([s])+2
        markCell(rowNum,column, wks)
        
        

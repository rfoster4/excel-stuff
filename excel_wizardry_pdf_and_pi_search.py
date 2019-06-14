# -*- coding: utf-8 -*-
#!/usr/bin/python3
"""
Created on Mon Jun 10 11:37:13 2019

@author: rkfoster
"""
import os
import openpyxl
from openpyxl import Workbook
#import re
#rootdir = r'S:\Photo\Don Early\00000_Summer\Richard'
#print(rootdir)
book = Workbook
pattern = '_pi_'
### Find the folder name from the Excel file...
i = 3
wb = openpyxl.load_workbook('C:\Temp\Index_Mosaic.xlsx')
sheet = wb.active
#sheet['A4']
#sheet['A4'].value
for i in range(3,1664):
    cellz = i + 1
    folderName = 'A{}'.format(cellz)
    searchFolder = sheet[folderName].value
    print(searchFolder)

### Search directory for folder...
    myDir = "S:\\GeospatialData\\PhotoDIL\\Digital\\Index_Mosaic\\" + searchFolder
    print(myDir)
    for filename in os.listdir(myDir):
        if pattern in filename and filename.endswith('.dgn'):
            writeInput = 'F{}'.format(cellz)
            sheet[writeInput].value = 'yes'
            print(filename)
        
        
### Take the file name and write it to the correct position in Excel



    
#end
        
wb.save('C:\Temp\Index_Mosaic.xlsx')

     
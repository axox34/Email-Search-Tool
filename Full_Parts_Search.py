"""
Given customer part list, does full database cross reference check and saves the results in a text file.
Works as a standalone for large queries
"""


import pyodbc
import pandas as pd
from _sqlite3 import Row
import io

import pyodbc
import pandas as pd
from _sqlite3 import Row
import io

def getpartinfo(partno):
    buffer1 = io.StringIO()
    cnxn = pyodbc.connect(DSN="SAGE100", UID='MAS_REPORTS', PWD='Reporting1')
    
        
    cursor = cnxn.cursor()
    
    '''
    cursor.execute("SELECT DISTINCT ItemCodeDesc FROM CI_Item ")
    rows = cursor.fetchall()
    for row in rows:
    
        y = row
        print(y[0])
        
        '''
   # itemcode = partno.get()
    #aliasitemno = partno.get()
    #rint ("TEST", aliasitemno)

    searchstr = partno
    cursor.execute("SELECT DISTINCT CI_Item.ItemCode, CI_Item.ItemCodeDesc, IM_AliasItem.AliasItemNo, CI_Item.UDF_ITEMDESC, CI_Item.UDF_IM_DESC_2 \
    FROM CI_Item FULL OUTER JOIN IM_AliasItem ON CI_Item.ItemCode=IM_AliasItem.ItemCode \
    WHERE AliasItemNo  like ? GROUP BY IM_AliasItem.AliasItemNo, CI_Item.ItemCode, CI_Item.ItemCodeDesc, CI_Item.UDF_ITEMDESC, CI_Item.UDF_IM_DESC_2",(searchstr+'%'))
    rows = cursor.fetchall()
    for row in rows:
    
        y = row
        results =  (y[2],y[0], y[1], y[3], y[4])
        results = '\t'.join([str(x) for x in y])
        return results

 
import openpyxl
 
#Prepare the spreadsheets to copy from and paste too.

    
#Copy range of cells as a nested list
#Takes: start cell, end cell, and sheet you want to copy from.
def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
 
    return rangeSelected

#File to be copied
wb = openpyxl.load_workbook("POR.xlsx") #Add file name
sheet = wb["Sheet1"] #Add Sheet name


if __name__ == "__main__":
    test = getpartinfo("853-285045-003")
    print(test)
    
    
    
    df = pd.read_excel ('POR.xlsx') #for an earlier version of Excel, you may need to use the file extension of 'xls'
    x = copyRange(1, 546, 1, 854, sheet)
    with open('meow.txt', 'w') as f:
        for item in range (0,len(x)):
            print(x[item])
            str1 = ",".join(str(bit) for bit in x[item])
            y= getpartinfo(str1)
            print(y)
            if y:
                f.write("%s\n" % y)
            else:
                f.write("not found \n")
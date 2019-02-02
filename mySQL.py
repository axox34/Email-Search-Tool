"""
Probably should save account information somewhere else.
Fetches the most important searches to display in emailsearchGUI.
The database has unknown structure so possibly there are more powerful tables that need to be explored.

"""
import pyodbc
import pandas as pd
from _sqlite3 import Row
import io

def getpartinfo(partno):
    buffer1 = io.StringIO()
    cnxn = pyodbc.connect(DSN="######", UID='######', PWD='####')
    
        
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

    searchstr = partno.get()
    cursor.execute("SELECT DISTINCT CI_Item.ItemCode, CI_Item.ItemCodeDesc, IM_AliasItem.AliasItemNo FROM CI_Item FULL OUTER JOIN IM_AliasItem ON CI_Item.ItemCode=IM_AliasItem.ItemCode WHERE AliasItemNo  like ? GROUP BY CI_Item.ItemCode, CI_Item.ItemCodeDesc, IM_AliasItem.AliasItemNo",(searchstr+'%'))
    rows = cursor.fetchall()
    if rows:
        print("[][][][][][][][][][][][][][][][][][][][][][][][][][][]\n", file = buffer1)
        print("------RESULTS BY ALIAS NUMBER-------", file = buffer1)
    for row in rows:
    
        y = row
        print ("%-30s    Item Code: %-6s    Desc: %-35s" % (y[2],y[0], y[1]), file = buffer1)
     
    cursor.execute("SELECT DISTINCT CI_Item.ItemCode, CI_Item.ItemCodeDesc, IM_AliasItem.AliasItemNo FROM CI_Item FULL OUTER JOIN IM_AliasItem ON CI_Item.ItemCode=IM_AliasItem.ItemCode WHERE CI_Item.ItemCode  like ? GROUP BY CI_Item.ItemCode, CI_Item.ItemCodeDesc, IM_AliasItem.AliasItemNo",(searchstr+'%'))
    rows = cursor.fetchall()
    if rows:
        print("[][][][][][][][][][][][][][][][][][][][][][][][][][][]\n", file = buffer1)
        print("--------RESULTS BY ITEM CODE--------",file = buffer1 )
    for row in rows:
    
        y = row
        print ("%-6s    ALIAS: %-30s    Desc: %-35s" % (y[0],y[2], y[1]), file = buffer1)

    
    cursor.execute("SELECT DISTINCT CI_Item.ItemCode, CI_Item.ItemCodeDesc, IM_AliasItem.AliasItemNo FROM CI_Item FULL OUTER JOIN IM_AliasItem ON CI_Item.ItemCode=IM_AliasItem.ItemCode WHERE CI_Item.ItemCodeDesc like ? GROUP BY CI_Item.ItemCode, CI_Item.ItemCodeDesc, IM_AliasItem.AliasItemNo",('%'+searchstr+'%'))
    rows = cursor.fetchall()
    if rows:
        print("[][][][][][][][][][][][][][][][][][][][][][][][][][][]\n", file = buffer1)
        print("---------RESULTS BY ITEM DESCRIPTION---------", file = buffer1 )
    for row in rows:
    
        y = row
        print ("%-35s    Item Code: %-6s    Alias: %-30s" % (y[1],y[0], y[2]), file = buffer1)
        #print(y[1],"    Item Code: ",y[0], "    ALIAS: ",y[2],file = buffer1 )

        

    print (buffer1.getvalue())
    return buffer1
    print("done")

if __name__ == "__main__":
    getpartinfo("853-285045")

# coding: utf-8

import os, sys
from shared_code.WappHandler import ParseTabText, colnum2Letter, MEXCEL
from shared_code.AzureHelper import XEDB, dict2obj,loads,dumps,CONT, hasattr, obj2dict
from openpyxl.worksheet.table import Table, TableStyleInfo

INPUTDIR="/Temp/80_SNr-Reporting"

import glob

inpfiles=glob.glob(INPUTDIR+"/10_Input/Input Fremdlisten aktuell/Digital Trucks/*.csv")


import csv

for inpfile in inpfiles:
    xls=MEXCEL("templates/empty.xlsx")
    ws=xls.asheet
    print("reading "+inpfile)
    inpfshort=inpfile.split("\\")[-1]
    #inpf=open(inpfile,"r",encoding="utf8")
    #content=inpf.read()
    if 0:
        inpf=open(inpfile,"rb")
        lines=[]
        line=""
        linenr=1
        while True:
            cread=inpf.read(1)
            if not cread:break
            try:
                charread=cread.decode()
                if charread in ["\n","\r"]:
                    if line:
                        #print(linenr,line)
                        lines.append(line)
                        line=""
                        linenr+=1
                else:
                    line+=charread
            except:
                pass
            print
    if 0:
        with open(inpfile, newline='') as csvfile:

            spamreader = csv.reader(csvfile, delimiter=' ', quotechar='|')

            mrow=0
            for row in spamreader:
                mrow+=1
                #print(mrow,row)

    #ptt=ParseTabText(None,1,",",lines=lines)
    ptt=ParseTabText(inpfile,1,",")
    ws.append(ptt.attl)
    for o in ptt.objects():
        toapp=[]
        for att in ptt.attl:
            toapp.append(getattr(o,att))
        #print(ptt._row,toapp)
        ws.append(toapp)
        #break
    tref="A1:%s%ld"%(colnum2Letter(len(ptt.attl)),ptt._row-1)#int(table.row)-1)
    tab = Table(displayName="Table1", ref=tref)
    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                        showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    '''
    Table must be added using ws.add_table() method to avoid duplicate names.
    Using this method ensures table name is unque through out defined names and all other table name. 
    '''
    ws.add_table(tab)
    outfname="output/%s.xlsx"%inpfshort[:-4]
    print("writing "+outfname)
    xls.save(outfname)
    

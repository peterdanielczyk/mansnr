# coding: utf-8

import os, sys, glob
from shared_code.MANSNR import ParseTabText, colnum2Letter, MEXCEL, matchall, checkInputFiles, intersect, extendUnique, getCurrentTimestamp, getOrAdd, TStamp, ECONT
from shared_code.AzureHelper import XEDB, dict2obj,loads,dumps,CONT, hasattr, obj2dict
#from openpyxl.worksheet.table import Table, TableStyleInfo

#Loggs to both, the terminal and a log file
Logger=XEDB.Logger
#collects all records written records to the target excel
WRITTEN={}

#read the config
CONFIG=loads(open("config.json").read())

"""wenn Spaltenindex_quelle = Source:
    dann Source-Wert von Quelle vor Ist-Eintrag von Zielzelle ergänzt [voher "Digital59", nachher "Digital62;Digital59")
wenn Spaltenindex_quelle = structure_path
    dann baue Zeilenübergreifend den Strukture-path auf (z. B. 1_2088106_2154891_)
wenn Spaltenidex_quelle = ....
    dann sonstige Rechenoperation
else
    copy paste mit Überschreiben des vorherigen Wertes in der Zielzelle"""
def Funktion_Zellwerte_kopieren(attribute,src_object, tgt_object, envi):
    if attribute=="lfdnr":
        tgt_object.lfdnr=XEDB.LFDNR
    elif attribute=="source":
        tgt_object.source=envi.inp.SOURCE
    elif attribute=="prio":
        tgt_object.prio=envi.inp.PRIO
    elif attribute=="structure_path":
        tgt_object.structure_path=getStructurePath(src_object)
    else:
        src_val=getattr(src_object,attribute)
        #if src_val:
        setattr(tgt_object,attribute,src_val)

def getStructurePath(src_object):
    try:
        rv2level=int(src_object.rv2level)
        STRUCTURE_PATH[rv2level]=src_object.rv2_id
        spath=""
        for i in range(0,rv2level+1):
            spath+=STRUCTURE_PATH.get(i)+"_"
        return spath
    except:return ""

def handleCollision(newo):
    #get object already in target
    if newo.source=="CombinedBOM":
        print
    allin=WRITTEN.get(newo.rv2_id)
    if not allin:
        #if in "other", skip the record
        if XEDB.input_name=="Input_other":return 1
        WRITTEN[newo.rv2_id]=newo
        return 0
    if newo.source not in allin.source:
        allin.setValue("source",allin.source+"; "+newo.source)
    for att in newo.__dict__:
        if att.startswith("_cfe_"):continue
        if att in EXTRA_COLS:continue
        allin_val=getattr(allin,att)
        if allin_val and isinstance(allin_val, str):
            allin_val=allin_val.strip()
            allin_val_cmp=allin_val.lower()
        else:allin_val_cmp=allin_val
        newo_val=getattr(newo,att)
        if not newo_val:continue
        if newo_val and isinstance(newo_val, str):
            newo_val=newo_val.strip()
            newo_val_cmp=newo_val.lower()
        else:newo_val_cmp=newo_val
        if allin_val_cmp==newo_val_cmp:continue
        if 1 or att=="rv2description":# and allin.rv2_id==1392830:
            pass
        else:continue

        if not allin_val:
            allin.setValue(att,newo_val)
            if 0:
                cmt="%s %s:''\n"%(allin.prio,allin.source) 
                allin.setComment(att,cmt+"%s %s:%s"%(newo.prio,newo.source,repr(newo_val)) )
            continue
        if 0 and att=="rv2designgroup":
            Logger.print("old: %s new:%s"%(allin_val,newo_val))
        oldcmt=allin.getComment(att)
        if oldcmt:
            oldcmt+="\n"
        allin.setComment(att,oldcmt+"%s %s:%s"%(newo.prio,newo.source,repr(newo_val) ))
    return 1

tstamp=TStamp()

XEDB.ALLOWED_FTYPES=[".csv",".xlsx"]
XEDB.LFDNR=1
STRUCTURE_PATH={}

xls=MEXCEL(CONFIG.template_path+"target_template.xlsx")
target=xls.getTable()

ATTS_TO_FILL=list(target.COLNAMES.values())
EXTRA_COLS=["lfdnr","prio","structure_path","source"]
ATTS_TO_READ=[]
for att in ATTS_TO_FILL:
    if att not in EXTRA_COLS:
        ATTS_TO_READ.append(att)
srco_cnt_total=0

"""Pfad dieser Datei finden
Liste der Dateiein in Pfad\Input_scope sowie Pfad\Input_other
Liste aufsteigend sortieren

Optional!!!
    User-Meldung auf Basis der Liste:
        folgende Dateien wurden gefunden und erfüllen folgende Q-Kriterien:
            Prio
            Source
            Doppelt-Optional: Spalte rv2._id (Unique-ID) wurde gefunden
    weiter oder Abbruch"""

for XEDB.input_name in ["Input_scope_path","Input_other_path","Output_path"]:
    input_dir=getattr(CONFIG,XEDB.input_name)
    Logger.print("======CHECKING %s (%s)=========\n"%(XEDB.input_name,input_dir))
        
    #checks, if the filenames contain Prio&Source and if the rv2._id attribute is found
    #otherwise, those files are skipped. Everything is logged out
    #Logger.print("Checking files in folder at %s"%(getCurrentTimestamp()))
    inpfiles=checkInputFiles(input_dir)

    Logger.print("\n====  Starting file processing for %s at %s ===="%(XEDB.input_name, getCurrentTimestamp()))

    for inp in inpfiles:
        #list of attributes in both: the source and the target
        attrl=intersect(inp.ptt.attl, ATTS_TO_FILL)
        #extend the list with extra attributes only in the target
        extendUnique(attrl,EXTRA_COLS)
        envi=CONT()
        envi.inp=inp
        srco_cnt=0
        for src_object in inp.ptt.objects(ATTS_TO_READ,preread=1):
            srco_cnt+=1
            srco_cnt_total+=1
            #initialize empty target object to be filled, this is a special excel object which can be used later on to add comments, change the value tranparently etc.
            tgt_object=ECONT()
            for att in attrl:
                Funktion_Zellwerte_kopieren(att,src_object,tgt_object, envi)
            #add the target object to the target table, if there is no collision - otherwise update the existing record
            if not handleCollision(tgt_object):
                target.setObject(tgt_object)
                XEDB.LFDNR+=1
            #if extra logging is required
            if XEDB.verbose:
                Logger.print(XEDB.LFDNR,dumps(src_object))

        Logger.print("Input file processed at %s: %s, %ld records"%(getCurrentTimestamp(),inp.inpfshort,srco_cnt))
    

#resize the target table properly
target.table.ref="A1:%s%ld"%(colnum2Letter(target.endColumn),XEDB.LFDNR)
outfname="target_%s_%s.xlsx"%(getCurrentTimestamp().replace(":","-").replace(" ","_"), os.getlogin())
Logger.print("%ld records read - %ld target rows written"%(srco_cnt_total,XEDB.LFDNR-1))
Logger.print("writing "+CONFIG.Output_path+outfname)
xls.save(CONFIG.Output_path+outfname)
Logger.print("Time needed for allover processing: %ld seconds"%int(tstamp.gap()/1000))
    

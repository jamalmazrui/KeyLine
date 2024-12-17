import os, sys
import db

iArgCount = len(sys.argv)
if iArgCount < 2: print("Syntax: ini2csv FileIni FileCsv"); sys.exit()
sFileIni = sys.argv[1]
if iArgCount >2: sFileCsv = sys.argv[2]
else: sFileCsv = sFileIni[0:-3] + "csv"

lHeaders = []
lRows = []
dIni = db.readConfig(sFileIni)
for sSection in dIni.keys():
	dSection = dIni[sSection]
	for sKey in dSection.keys():
		if sKey not in lHeaders: lHeaders.append(sKey)
	dRow = {k: (v.strip() + "\n" if "\n" in v.strip() else v.strip()) for k, v in dSection.items()}
	lRows.append(dRow)

db.saveCsv(sFileCsv, lHeaders, lRows)

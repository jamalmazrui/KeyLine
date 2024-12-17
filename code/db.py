# August 2, 2020
import csv, os, re, sys, sqlite3
from collections import OrderedDict
from pprint import pprint
from configobj import ConfigObj

def saveCsv(sFileCsv=None, lHeaders=None, lRows=None):
	# f = open(sFileCsv, "w", encoding="utf-8", newline="")
	f = open(sFileCsv, "w", encoding="utf-8-sig", newline="")
	writer = csv.DictWriter(f, fieldnames=lHeaders)
	writer.writeheader()
	for row in lRows: writer.writerow(dict(row))
	f.close()

def lister(sText): return sText.split() 

def readConfig(sIniFile):
	config = ConfigObj(infile=sIniFile, encoding="utf-8", interpolation=False, list_values=False, stringify=True, indent_type="", default_encoding="utf-8", write_empty_values=True, )
	config.bom = True
	return config

def commaJoin(lTerms): return ", ".join(lTerms)

def commaSplit(sTerms): return [s.strip() for s in sTerms.split(",")]

def beginTransaction(): return con.execute("begin transaction")

def getTables():
	rows = fetchRows(lTables=["sqlite_master"], lSelectFields=["name"], dWhereFields={"type": "table"}, lOrderFields = "1")
	lNames = [row[0] for row in rows]
	return lNames

def getViews():
	rows = fetchRows(lTables=["sqlite_master"], lSelectFields=["name"], dWhereFields={"type": "view"}, lOrderFields = ["1"])
	lViews = [row[0] for row in rows]
	return lViews

def drop(sType, sName):
	sCommand = "drop " + sType + " if exists " + sName
	return cursor.execute(sCommand)

def makeViewCommand(sView, lTables=None, lSelectFields=None, dWhereFields=None, lOrderFields=None):
	if lTables: sTables = commaJoin(lTables)
	if lSelectFields == None: sSelectFields = "*"
	else: sSelectFields = commaJoin(lSelectFields)
	if lOrderFields != None: sOrderFields = commaJoin(lOrderFields)

	sCommand = "create view if not exists " + sView + " "
	# sCommand += makeSelectCommand(sTables, sSelectFields, dWhereFields, sMisc)
	sCommand += makeSelectCommand(lTables, lSelectFields, dWhereFields, lOrderFields)
	return sCommand

def createView(sView, lTables=None, lSelectFields=None, dWhereFields=None, lOrderFields = None):
	# if lSelectFields == None: sSelectFields = "*"
	# else: sSelectFields = commaJoin(lSelectFields)
	sCommand = makeViewCommand(sView, lTables, dWhereFields, lOrderFields)
	return cursor.execute(sCommand, dWhereFields)

def getTableInfo(sTable):
	sCommand = "pragma table_info(" + sTable + ")"
	rows = cursor.execute(sCommand).fetchall()
	return rows

def updateRow(sTable, dUpdateFields, dWhereFields):
	sCommand = makeupdateCommand(sTable, dUpdateFields, dWhereFields)
	dCombinedFields = {}
	# for sField in dUpdateFields.keys(): dCombinedFields[":" + sField] = dUpdateFields[sField]
	for k, v in dUpdateFields.items(): dCombinedFields[k] = v
	for k, v in dWhereFields.items(): dCombinedFields[k] = v
	print("sCommand", sCommand)
	pprint(dCombinedFields)
	return cursor.execute(sCommand, dCombinedFields)

def makeupdateCommand(sTable, dUpdateFields=None, dWhereFields=None):
	sCommand = "update " + sTable + " "
	if dUpdateFields:
		for iField, sField in enumerate(dUpdateFields.keys()):
			if iField == 0: sCommand += "set "
			else: sCommand += ", "
			sCommand += sField + " = " + ":" + sField
	if dWhereFields:
		sCommand += " where "
		for iField, sField in enumerate(dWhereFields.keys()):
			if iField > 0: sCommand += " and "
			sCommand += sField + " == :" + sField
	return sCommand

def old_makeupdateCommand(sTable, dUpdateFields, dWhereFields):
	sCommand = "update " + sTable + " "
	for iField, sField in enumerate(dUpdateFields.keys()):
		sCommand += "set " + sField + " = " + repr(dUpdateFields[sField])
	if dWhereFields:
		sCommand += " where "
		for iField, sField in enumerate(dWhereFields.keys()):
			if iField > 0: sCommand += " and "
			sCommand += sField + " == :" + sField
	return sCommand

def fetchId(sTable, dWhereFields):
	sField = getPrimaryKeyField(sTable)
	return fetchValue([sTable], sField, dWhereFields)

def deleteRows(sTable, dWhereFields):
	sCommand = makeDeleteCommand(sTable, dWhereFields)
	return cursor.execute(sCommand, dWhereFields)

def makeDeleteCommand(sTable, dWhereFields):
	sCommand = "delete from " + sTable + " "
	if dWhereFields:
		sCommand += " where "
		for iField, sField in enumerate(dWhereFields.keys()):
			if iField > 0: sCommand += " and "
			sCommand += sField + " == :" + sField
	return sCommand

def makeCreateTableCommand(sTable, dFieldTypes):
	sCommand = "create table if not exists " + sTable + " ("
	for iField, sField in enumerate(dFieldTypes):
		if iField > 0: sCommand += ", "
		sCommand += sField + " " + dFieldTypes[sField]

	sCommand += ")"
	return sCommand

def createTable(sTable, dFieldTypes):
	sCommand = makeCreateTableCommand(sTable, dFieldTypes)
	return cursor.execute(sCommand, dFieldTypes)

def getDistinctFields(sTable):
	lDrop = "added updated observed notes tags marked".split(" ")
	lDrop.append(getPrimaryKeyField(sTable))
	lFields = getFieldTypes(sTable).keys()
	for s in lDrop:
		if s in lFields: lFields.remove(s)
	return lFields

def getPrimaryKeyField(sTable):
	sField = None
	rows = getTableInfo(sTable)
	for row in rows:
		if row["pk"]:
			sField = row["name"]
			break
	return sField

def old_getPrimaryKeyField(sTable):
	dWhereFields = {"type": "table", "name": sTable}
	sSql = fetchValue("sqlite_master", "sql", dWhereFields)
	sRegex = r"\bPRIMARY KEY *\( *(.*?) *\)"
	m1 = re.search(sRegex, sSql, re.I)
	sRegex = r"(\w+) +(\w+) +PRIMARY KEY"
	m2 = re.search(sRegex, sSql, re.I)

	if m1: sField = m1.group(1)
	elif m2: sField = m2.group(1)
	else: sField = None
	return sField

def getEmptyUniqueConstraintFields(sTable):
	lFields = getUniqueConstraintFields(sTable)
	dReturn = {sField: None for sField in lFields}
	return dReturn

def getUniqueConstraintFields(sTable):
	dWhereFields = {"type": "table", "name": sTable}
	sSql = fetchValue(["sqlite_master"], "sql", dWhereFields)
	sRegex = r"CONSTRAINT +\w+ +UNIQUE *\((.*?)\)"
	print("sRegex", sRegex)
	print("sSql", sSql)
	m = re.search(sRegex, sSql)
	if m:
		sFields = m.group(1)
		lFields = sFields.split(",")
		lFields = [s.strip() for s in lFields]
	else: lFields = []
	return lFields

def getFieldTypes(sTable):
	sCommand = "pragma table_info(" + sTable + ")"
	rows = cursor.execute(sCommand).fetchall()
	dReturn = {}
	for row in rows: dReturn[row[1]] = row[2]
	return dReturn

def openDb(sDbFile, bDictRow=True):
	global con, cursor
	# con = sqlite3.connect(sDbFile, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
	con = sqlite3.connect(sDbFile)
	if bDictRow: con.row_factory = sqlite3.Row
	cursor = con.cursor()
	return con, cursor

def stringQuote(sText): return '"' + sText + '"'

# def upsertRow(sTable, dInsertFields=None): insertRow(sTable, dInsertFields, bUpsert=True)
def upsertRow(sTable, dInsertFields=None):
	# print("sTable=", sTable)
	pprint(dInsertFields)
	iId = fetchId(sTable, dInsertFields)
	# print("iId", iId)
	# if iId: insertRow(sTable, dInsertFields, bUpsert=True); return iId 
	# if iId: dInsertFields[getPrimaryKeyField(sTable)] = iId; insertRow(sTable, dInsertFields, bUpsert=True); return iId 
	if iId:
		dWhereFields = {getPrimaryKeyField(sTable): iId}
		updateRow(sTable, dInsertFields, dWhereFields)
	else: return insertRow(sTable, dInsertFields)

def insertRow(sTable, dInsertFields=None, bUpsert=False):
	# print("bUpsert", bUpsert)
	sCommand = makeInsertCommand(sTable, dInsertFields, bUpsert)
	# print("sCommand", sCommand)
	# print("dInsertFields", dInsertFields)
	cursor.execute(sCommand, dInsertFields)
	# print("lastrowid", cursor.lastrowid)
	return cursor.lastrowid

def makeInsertCommand(sTable, dInsertFields=None, bUpsert=False):
	if bUpsert: sCommand = "insert or replace into " + sTable
	else: sCommand = "insert into " + sTable
	if dInsertFields:
		sCommand += " ("
		for iField, sField in enumerate(dInsertFields.keys()):
			if iField > 0: sCommand += ", "
			sCommand += sField
		sCommand += ") values ("
		for iField, sField in enumerate(dInsertFields.keys()):
			if iField > 0: sCommand += ", "
			sCommand += ":" + sField
		sCommand += ")" 
	return sCommand

def makeInsertTupleCommand(sTable, lInsertFields):
	sCommand = "insert into " + sTable + " ("
	for iField, sField in enumerate(lInsertFields):
		if iField > 0: sCommand += ", "
		sCommand += sField
	sCommand += ") values ("
	for iField, sField in enumerate(lInsertFields):
		if iField > 0: sCommand += ", "
		sCommand += "?"
	sCommand += ")" 
	return sCommand

def makeSelectCommand(lTables=None, lSelectFields=None, dWhereFields=None, lOrderFields = None):
	if lTables != None: sTables = commaJoin(lTables)
	if lSelectFields == None: sSelectFields = "*"
	else: sSelectFields = commaJoin(lSelectFields)
	if lOrderFields != None: sOrderFields = commaJoin(lOrderFields)

	# sFields = sSelectFields
	# if not sFields and dWhereFields: sFields = commaJoin(dWhereFields.keys())
	# sCommand = "select " + sFields
	sCommand = "select " + sSelectFields
	if lTables: sCommand += " from " + sTables
	if dWhereFields:
		sCommand += " where "
		for iField, sField in enumerate(dWhereFields.keys()):
			if iField > 0: sCommand += " and "
			sCommand += sField + " == :" + sField
			if lOrderFields: sCommand += " order by " + sOrderFields
	return sCommand

def fetchValue(lTables=None, sField=None, dWhereFields=None, lOrderFields=None):
	# print("sField", sField)
	# print("lTables", lTables)
	sCommand = makeSelectCommand(lTables, [sField], dWhereFields, lOrderFields)
	print("sCommand", sCommand)
	pprint(dWhereFields)
	row = cursor.execute(sCommand, dWhereFields).fetchone()
	xValue = (row[sField] if row else None)
	# print("xValue", xValue)
	return xValue

def fetchRows(lTables=None, lSelectFields=None, dWhereFields=None, lOrderFields=None):
# if lSelectFields == None: sSelectFields = "*"
	# else: sSelectFields = commaJoin(lSelectFields)
	sCommand = makeSelectCommand(lTables, lSelectFields, dWhereFields, lOrderFields)
	if dWhereFields == None: dWhereFields = {}
	rows = cursor.execute(sCommand, dWhereFields).fetchall()
	# rows = cursor.execute(sCommand, dWhereFields.values()).fetchall()
	return rows

def getUniqueId(sTable, sField, dWhereFields):
	xValue = fetchValue(sTable, sField, dWhereFields)
	if not sValue:
		sCommand = makeInsertCommand(sTable, dWhereFields)
		cursor.execute(sCommand, dWhereFields)
		sValue = cursor.lastrowid
	return sValue

"""
# main
sDbFile = r"C:\Kaxe\AccAudit.db"
sTable = "apps"
dWhereFields = {"name": "Kindle", "os": "ios", "variations": None}
sCommand = makeInsertCommand(sTable, dWhereFields)

sCommand = makeSelectCommand(sTables, "*", dWhereFields)

con,cursor = openDb(sDbFile)
d = getFieldTypes("apps")

d = getUniqueConstraintFields("apps")

d = getEmptyUniqueConstraintFields("problems")

s = getPrimaryKeyField("types")



l = getDistinctFields("problems")



d = {"person_id": "integer primary key not null", "first": "varchar", "last": "varchar"}
s = "persons"
o = createTable(s, d)

i = addRow("persons", {"first": "Susan", "last": "Mazrui"})
j = fetchId("persons", {"first": "Jamal", "last": "Mazrui"})
# deleteRows("persons", {"person_id": j})
updateRows("persons", {"first": "Fred"}, {"person_id": j})

sField = getPrimaryKeyField("lookups")

drop("table", "persons")
con.commit()
con.close()


sTable = "test"
# dFields = {"id": "int", "added": "DateTime", "name": "str"}
# dFields = OrderedDict({"id": "int", "added": "DateTime", "name": "str"})
dFields = OrderedDict([("id", "int"), ("added", "DateTime"), ("name", "str")])
print makeCreateTableCommand(sTable, dFields)
"""


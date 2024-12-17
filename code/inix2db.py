print("starting")
import os, sys
import db

iArgCount = len(sys.argv)
sInix, sDb = "", ""
if iArgCount > 1: sInix = sys.argv[1]
if iArgCount > 2: sDb = sys.argv[2]
if not sInix: sInix = os.getenv("AccAuditInix")
if not sDb: sDb = os.getenv("AccAuditDb")

print("sInix", sInix)
print("sDb", sDb)
dPrimaryKeys = {}
con, cursor = db.openDb(sDb)
dInix = db.readConfig(sInix)
for sSection in dInix.keys():
	dSection = dInix[sSection]
	sTable = dSection["table"]
	del(dSection["table"])
	dFieldTypes = db.getFieldTypes(sTable)
	dSection = {k: v for k, v in dSection.items() if k in dFieldTypes.keys()}
	sPrimaryKeyName = db.getPrimaryKeyField(sTable)
	lUniqueFields = db.getUniqueConstraintFields(sTable)
	for sKey in dPrimaryKeys.keys():
		# if sKey != sPrimaryKeyName and not dSection.has_key(sKey) and sKey in dFieldTypes.keys(): dSection[sKey] = dPrimaryKeys[sKey]
		if sKey != sPrimaryKeyName and not sKey in dSection.keys() and sKey in dFieldTypes.keys(): dSection[sKey] = dPrimaryKeys[sKey]
	# iPrimaryKeyValue = db.insertRow(sTable, dSection)
	iPrimaryKeyValue = db.upsertRow(sTable, dSection)
	print(sPrimaryKeyName, iPrimaryKeyValue)
	dPrimaryKeys[sPrimaryKeyName] = iPrimaryKeyValue


	



con.commit()
con.close()

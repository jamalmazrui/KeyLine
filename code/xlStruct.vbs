' xlProperties.vbs - Analyze and output the structure and attributes of an Excel workbook

' Constant definitions
const iErrorInvalidArguments = 1
const iErrorOpeningWorkbook = 2
const bIncludeDataRegions = true
const bIncludeBuiltinProperties = false
const bIncludeRecentFiles = false

' Variable definitions
dim bEnclosed
dim dFilteredRegions, dProperties, dRegionAddresses, dRegions
dim iSheetCount
dim oCell, oCurrentRegion, oExcel, oOtherRegion, oRegion, oSheet, oWorkbook
dim sOtherRegion, sRegionName, sRegion, sWorkbookPath
dim vValue

' Ensure a single argument is provided
if wscript.arguments.count <> 1 then
    wscript.echo "Usage: xlWorkbookOverview.vbs ""<WorkbookPath>"""
    wscript.quit iErrorInvalidArguments
end if

' Get the workbook path from the command-line arguments
sWorkbookPath = wscript.arguments(0)

' Create Excel application object
set oExcel = createObject("Excel.Application")
oExcel.visible = false
on error resume next
set oWorkbook = oExcel.workbooks.open(sWorkbookPath)
if err.number <> 0 then
    wscript.echo "Error: Unable to open workbook at " & sWorkbookPath
    oExcel.quit
    set oExcel = nothing
    wscript.quit iErrorOpeningWorkbook
end if
on error goto 0

' Output workbook properties
wscript.echo "Workbook: " & oWorkbook.fullname
if bIncludeBuiltinProperties then
wscript.echo "Builtin Properties:"
for each dProperties in oWorkbook.BuiltinDocumentProperties
on error resume next
'     if isEmpty(dProperties.value) = false then
    if trim(dProperties.value) <> "" then
        wscript.echo "" & dProperties.name & ": " & dProperties.value
on error goto 0
    end if
next
end if

' Include recent files if specified
if bIncludeRecentFiles then
    wscript.echo vbnewline & "Recent Files:"
    for each vValue in oExcel.recentFiles
        wscript.echo "" & vValue.Path
    next
end if

' Count and summarize sheets
iSheetCount = oWorkbook.sheets.count
wscript.echo vbnewline & "Number of Sheets: " & iSheetCount
for each oSheet in oWorkbook.sheets
wscript.echo
    wscript.echo "Sheet: " & oSheet.name
    wscript.echo "UsedRange: " & oSheet.usedRange.address
    wscript.echo "Rows: " & oSheet.usedRange.rows.count
    wscript.echo "Columns: " & oSheet.usedRange.columns.count

    ' Analyze data regions
    if bIncludeDataRegions then
        set dRegionAddresses = createObject("Scripting.Dictionary")

        for each oCell in oSheet.usedRange
            if not isEmpty(oCell.value) then
                set oRegion = oCell.currentRegion
                if not dRegionAddresses.exists(oRegion.address) then
                    dRegionAddresses.add oRegion.address, oRegion
                end if
            end if
        next
wscript.echo "Regions: " & dRegionAddresses.count

        ' Filter enclosed regions using row/column comparisons
        set dFilteredRegions = createObject("Scripting.Dictionary")
        for each sRegion in dRegionAddresses.keys
            set oCurrentRegion = dRegionAddresses(sRegion)
            bEnclosed = false
'             for each oOtherRegion in dRegionAddresses
            for each sOtherRegion in dRegionAddresses.keys
                set oOtherRegion = dRegionAddresses(sOtherRegion)
                if oCurrentRegion.address <> oOtherRegion.address then
                    if isRegionEnclosed(oCurrentRegion, oOtherRegion) then
                        bEnclosed = true
                        exit for
                    end if
                end if
            next
            if not bEnclosed then dFilteredRegions.add oCurrentRegion.address, oCurrentRegion
        next

        ' Output region details
        for each sRegion in dFilteredRegions
            set oRegion = dFilteredRegions(sRegion)
wscript.echo
sRegionName = "unnamed"
on error resume next
' sRegionName = oRegion.name
sRegionName = oRegion.name.name
' if isEmpty(sRegionName) then sRegionName = oRegion.rows(1).cells(1).name
if isEmpty(sRegionName) then sRegionName = oRegion.rows(1).cells(1).name.name
on error goto 0
            wscript.echo "Region: " & sRegionName
            wscript.echo "Address: " & oRegion.address
            wscript.echo "Rows: " & oRegion.rows.count
            wscript.echo "Columns: " & oRegion.columns.count
            wscript.echo "Formulas: " & bContainsFormulas(oRegion)
            wscript.echo "Comments: " & bContainsComments(oRegion)
            wscript.echo "Hyperlinks: " & bContainsHyperlinks(oRegion)
        next
    end if
next

' Helper functions
function bContainsFormulas(oRange)
    dim oCell
    bContainsFormulas = false
    for each oCell in oRange
        if oCell.hasFormula then
            bContainsFormulas = true
            exit function
        end if
    next
end function

function bContainsComments(oRange)
    dim oCell
    bContainsComments = false
    for each oCell in oRange
        if not isEmpty(oCell.comment) then
            bContainsComments = true
            exit function
        end if
    next
end function

function bContainsHyperlinks(oRange)
    dim oCell
    bContainsHyperlinks = false
    for each oCell in oRange
        if not isEmpty(oCell.hyperlinks) then
            bContainsHyperlinks = true
            exit function
        end if
    next
end function

function isRegionEnclosed(oInner, oOuter)
    isRegionEnclosed = false
    if oInner.rows(1).row >= oOuter.rows(1).row and _
        oInner.rows(oInner.rows.count).row <= oOuter.rows(oOuter.rows.count).row and _
       oInner.columns(1).column >= oOuter.columns(1).column and _
       oInner.columns(oInner.columns.count).column <= oOuter.columns(oOuter.columns.count).column then
        isRegionEnclosed = true
    end if
end function

' Cleanup
oWorkbook.close false
oExcel.quit
set oWorkbook = nothing
set oExcel = nothing

Option Explicit
WScript.Echo "Starting SaveAppointments"

Dim dExtensions, dFolders
Dim iAppointment, iMailBox, iMessage, iMessageCount, iAttachment, iAttachmentCount
Dim oApp, oExplorer, oMessages, oMessage, oAttachments, oAttachment, oNameSpace, oFolder
Dim sExtensions, sFilter, sIcalFile, sMessage, sRoot, sExt, sAction, sFolder, sBaseName, sFileName, sDir

'olFolder enumeration

Const olFolderCalendar = 9 ' The Calendar folder.
Const olFolderConflicts = 19 ' The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
Const olFolderContacts = 10 ' The Contacts folder.
Const olFolderDeletedItems = 3 ' The Deleted Items folder.
Const olFolderDrafts = 16 ' The Drafts folder.
Const olFolderInbox = 6 ' The Inbox folder.
Const olFolderJournal = 11 ' The Journal folder.
Const olFolderJunk = 23 ' The Junk E-Mail folder.
Const olFolderLocalFailures = 21 ' The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
Const olFolderManagedEmail = 29 ' The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
Const olFolderNotes = 12 ' The Notes folder.
Const olFolderOutbox = 4 ' The Outbox folder.
Const olFolderSentMail = 5 ' The Sent Mail folder.
Const olFolderServerFailures = 22 ' The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
Const olFolderSuggestedContacts = 30 ' The Suggested Contacts folder.
Const olFolderSyncIssues = 20 ' The Sync Issues folder. Only available for an Exchange account.
Const olFolderTasks = 13 ' The Tasks folder.
Const olFolderToDo = 28 ' The To Do folder.
Const olPublicFoldersAllPublicFolders = 18 ' The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
Const olFolderRssFeeds = 25 ' The RSS Feeds folder.

' Outlook SaveAs enumeration
Const olDoc = 4
Const olHTML = 5
Const olICal = 8
Const olMSG = 3
Const olMSGUnicode = 9
Const olRTF = 1
Const olTemplate = 2
Const olTXT = 0
Const olVCal = 7
Const olVCard = 6

Dim a, aStyles, aIni
Dim bLoop, bBackupDocx, bLogActions, bValue, bFound, bAddToRecentFiles, bConfirmConversions, bIncludePageNumbers, bHidePageNumbersInWeb, bRightAlignPageNumbers, bUseFields, bUseHeadingStyles, bUseHyperlinks, bUseOutlineLevels, bReadOnly
Dim bFormat, bForward, bMatchAlefHamza, bMatchAllWordForms, bMatchCase, bMatchControl, bMatchDiacritics, bMatchKashida, bMatchSoundsLike, bMatchWholeWord, bMatchWildcards
Dim d, dHeadingStyles, dStyle, dIni, dSourceIni, dSection
Dim iValue, i, iLevel, iReplaceCount, iTableId, iReplace, iWrap, iForward, iArgCount, iCount, iLowerHeadingLevel, iUpperHeadingLevel
Dim oFindFormat, oFindFont, oReplaceFormat, oReplaceFont, oSystem, oFile, oParagraph, oField, oAddedStyles, oData, oDoc, oDocs, oFind, oFont, oFormat, oProperty, oRange, oReplace, oStyle, oStyles, oToc, oTocs
Dim nValue
Dim sBackupDocx, sTargetLog, sScriptVbs, sHomerLibVbs, sCode, sFindStyle, sReplaceStyle, sKey, sFind, sFindText, sReplaceText, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni, sSection

' wdStyleType enumeration
Const wdStyleTypeParagraph = 1
Const wdStyleTypeCharacter = 2
Const wdStyleTypeTable = 3
Const wdStyleTypeList = 4

' wdOrganizerObject enumeration
Const wdOrganizerObjectStyles = 0
Const wdOrganizerObjectAutoText = 1
Const wdOrganizerObjectCommandBars = 2
Const wdOrganizerObjectProjectItems = 3

Const wdOutlineLevel1 = 1 ' Outline level 1
Const wdOutlineLevelBodyText = 10 'No outline level

Const wdRDIAll = 99 
' Removes all document information.
Const wdRDIComments = 1 
' Removes document comments.
Const wdRDIContentType = 16 
' Removes content type information.
Const wdRDIDocumentManagementPolicy = 15 
' Removes document management policy information.
Const wdRDIDocumentProperties = 8 
' Removes document properties.
Const wdRDIDocumentServerProperties = 14 
' Removes document server properties.
Const wdRDIDocumentWorkspace = 10 
' Removes document workspace information.
Const wdRDIEmailHeader = 5 
' Removes e-mail header information.
Const wdRDIInkAnnotations = 11 
' Removes ink annotations.
Const wdRDIRemovePersonalInformation = 4 
' Removes personal information.
Const wdRDIRevisions = 2 
' Removes revision marks.
Const wdRDIRoutingSlip = 6 
' Removes routing slip information.
Const wdRDISendForReview = 7 
' Removes information stored when sending a document for review.
Const wdRDITaskpaneWebExtensions = 17 
' Removes taskpane web extensions information.
Const wdRDITemplate = 9 
' Removes template information.
Const wdRDIVersions = 3 
' Removes document version information.

Const wdNoProtection = -1
 
Const WdCollapseEnd = 0
Const WdDoNotSaveChanges = 0

Const wdFindContinue = 1

Const wdReplaceOne = 1
Const wdReplaceAll = 2

Function FileInclude(sFile)
With CreateObject("Scripting.FileSystemObject")
ExecuteGlobal .openTextFile(sFile).readAll()
End With

' executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

' Main
sScriptVbs = WScript.ScriptFullName
Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.GetFile(sScriptVbs)
sDir = oSystem.GetParentFolderName(oFile) 
sHomerLibVbs = sDir & "\" & "HomerLib.vbs"
FileInclude sHomerLibVbs
' FileInclude "HomerLib.vbs"

iArgCount = WScript.Arguments.Count

' ' If iArgCount < 2 Then Quit "Specify a source .docx file as the first parameter an a configuration .ini file as the second parameter."
sAction = "save"
sFolder = "Calendar"
sExtensions = ""
sFilter = ""

If iArgCount > 0 Then sAction = LCase(WScript.Arguments(0))
If iArgCount > 1 Then sFolder = lCase(WScript.Arguments(1))
If iArgCount > 2 Then sExtensions = lCase(WScript.Arguments(2))
If iArgCount > 3 Then sFilter = lCase(WScript.Arguments(3))

' bBackupDocx = GetGlobalValue(dSourceIni, "BackupDocx", True)
' bLogActions = GetGlobalValue(dSourceIni, "LogActions", True)

sTargetLog = PathCombine(PathGetCurrentDirectory(), "Appointments.log")

sDir = PathGetCurrentDirectory()

Set dFolders = CreateDictionary
dFolders.Add "calendar", 9
dFolders.Add "contacts", 10
dFolders.Add "deleted", 3
dFolders.Add "drafts", 16
dFolders.Add "inbox", 6
dFolders.Add "journal", 11
dFolders.Add "junk", 23
dFolders.Add "notes", 12
dFolders.Add "outbox", 4
dFolders.Add "sent", 5
dFolders.Add "tasks", 13
' dFolders.Add "todo", 28

If not (sAction = "remove" or sAction = "list" or sAction = "save") Then Quit "No " & sAction & " action"
If Not dFolders.Exists(sFolder) Then Quit "No " & sFolder & "folder"
iMailBox = dFolders(sFolder)

Set oApp = CreateObject("Outlook.Application")
Set oNameSpace = oApp.GetNamespace("MAPI")
Set oFolder = oNamespace.GetDefaultFolder(iMailbox)
Set oMessages= oFolder.Items
If Len(sFilter) >0 Then oMessages = oMessages.Restrict(sFilter) ' "[UnRead] = True"
iMessageCount = oMessages.Count
printBlank
print "Saving " & StringPlural("appointment", iMessageCount)

For iMessage = 1 to iMessageCount
Set oMessage = oMessages(iMessage)
' sMessage = Trim(oMessage.Subject)
sMessage = Trim(oMessage.Subject & " " & FormatShortDateTime(oMessage.Start))
print iMessage & ". " & sMessage

sRoot = sMessage
' sExt = "ical"
sExt = "ics"
' sExt = "vcal"
sIcalFile = PathGetUnique(sDir, sRoot, sExt)
' print "target=" & sIcalFile
On Error Resume Next
oMessage.SaveAs sIcalFile, olICal
' oMessage.SaveAs sIcalFile, olVCal
On Error GoTo 0
sIcalFile = PathChangeExtension(sIcalFile, "txt")
oMessage.SaveAs sIcalFile, olTXT    
Next
printBlank

printBlank
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf

oApp.Quit

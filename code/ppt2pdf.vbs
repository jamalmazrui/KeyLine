Option Explicit

Function GetOutputPath(inputPath, extension)
   Dim file
   Dim basename
   Dim foldername
   
   basename = FSO.GetBaseName( inputPath )
   set file = FSO.GetFile( inputPath )
   foldername = file.ParentFolder
   GetOutputPath = foldername & "\" & basename & extension
End Function

Sub ConvertToPDF(ppt, inputpath, outputpath)
   Dim presentation
   Dim printoptions

   ppt.Presentations.Open inputpath

   set presentation = ppt.ActivePresentation
   set printoptions = presentation.PrintOptions

   printoptions.Ranges.Add 1,presentation.Slides.Count
   printoptions.RangeType = 1 ' Show all.

   const ppFixedFormatTypePDF = 2
   const ppFixedFormatIntentScreen = 1
   const msoFalse = 0
   const msoTrue = -1
   const ppPrintHandoutHorizontalFirst = 2
   const ppPrintOutputSlides = 1
   const ppPrintAll = 1

   presentation.ExportAsFixedFormat outputpath, ppFixedFormatTypePDF, ppFixedFormatIntentScreen, msoTrue, ppPrintHandoutHorizontalFirst, ppPrintOutputSlides, msoFalse, printoptions.Ranges(1), ppPrintAll, inputFile, False, False, False, False, False

   presentation.Close
End Sub

Dim FSO
Dim ppt

set FSO = CreateObject("Scripting.FileSystemObject")
set ppt = CreateObject("PowerPoint.Application")
ppt.Visible = True

If WScript.Arguments.Length > 0 Then
  Dim inputFile
  Dim inputpath
  Dim outputpath

  inputFile = WScript.Arguments(0)

  If Not FSO.FileExists( inputFile ) Then
     WScript.Stdout.Writeline "File not found: " & inputFile
  End If

  inputpath = FSO.GetAbsolutePathName( inputFile )
  WScript.Stdout.WriteLine "Full path: " & inputpath

  outputpath = GetOutputPath( inputpath, ".pdf" )

  WScript.Stdout.WriteLine "Output path: " & outputpath

  ConvertToPdf ppt, inputpath, outputpath

Else 
   Dim RootFolder
   Dim folder 
   Dim fld
   Dim folder_stack 
   Dim file_stack
   Dim file

   set folder_stack = CreateObject("System.Collections.Stack")
   set file_stack = CreateObject("System.Collections.Stack")
   set RootFolder = FSO.GetFolder(".")
   folder_stack.push(RootFolder)
   While folder_stack.Count > 0 
      set folder = folder_stack.pop()
	  WScript.Stdout.WriteLine "Processing " & folder.Name 
	  If Folder.SubFolders.Count > 0  Then
         For Each fld in folder.SubFolders
		    folder_stack.push(fld)
		 Next
	  Else
	     For Each file in folder.Files
		   Dim extension
		   extension = FSO.GetExtensionName(file.path)
		   If extension = "pptx" or extension = "ppt" Then
			  WScript.Stdout.WriteLine "Converting " & file.Path
			  outputpath = GetOutputPath( file.Path, ".pdf" )
			  ConvertToPDF ppt, file.Path, outputpath
		   End If
		 Next
	  End If
   Wend
   
   ppt.Quit
End If
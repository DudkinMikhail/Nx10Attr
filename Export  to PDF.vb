Option Strict Off
Imports System
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms
Imports System.Windows.Forms.MessageBox
Imports NXOpen
Imports NXOpen.UF
 
Module NXJournal
 
Dim theSession As Session = Session.GetSession()
Dim workPart As Part = theSession.Parts.Work
Dim displayPart As Part = theSession.Parts.Display

Function GetFileName()
	Dim strPath as String
	Dim strPart as String
	Dim pos as Integer
	
	'get the full file path
	strPath = displayPart.fullpath
	'get the part file name
	pos = InStrRev(strPath, "\")
	strPart = Mid(strPath, pos + 1)
	
	strPath = Left(strPath, pos)
	'strip off the ".prt" extension
	strPart = Left(strPart, Len(strPart) - 4)
	
	GetFileName = strPart
End Function


Function GetFilePath()
	Dim strPath as String
	Dim strPart as String
	Dim pos as Integer
	
	'get the full file path
	strPath = displayPart.fullpath
	'get the part file name
	pos = InStrRev(strPath, "\")
	strPart = Mid(strPath, pos + 1)
	
	strPath = Left(strPath, pos)
	'strip off the ".prt" extension
	strPart = Left(strPart, Len(strPart) - 4)
	
	GetFilePath = strPath
End Function


Sub Main

	Dim dwgs As Drawings.DrawingSheetCollection
	dwgs = workPart.DrawingSheets
	Dim sheet As Drawings.DrawingSheet
	Dim pdfFile As String
	Dim currentPath As String
	Dim currentFile As String
	Dim exportFile As String
	Dim partUnits As Integer
	Dim strOutputFolder As String
	Dim strRevision As String
	Dim rspFileExists

	 
	 
	Dim shts As New ArrayList()
	For Each sheet in dwgs
	shts.Add(sheet.Name)
	Next
	shts.Sort()
	 
	Dim sht As String
	For Each sht in shts
	 
	   For Each sheet in dwgs
		   If sheet.name = sht Then

				pdfFile=GetFilePath & "\Out\" & GetFileName() & ".pdf"
'				theSession.Parts.Work.DraftingViews.UpdateViews(Drawings.DraftingViewCollection.ViewUpdateOption.OutOfDate, sheet)
	 
				Try
				   ExportPDF(sheet, pdfFile, partUnits)
				Catch ex As exception
				   msgbox("Error occurred in PDF export" & vbcrlf & ex.message & vbcrlf & "journal exiting", vbcritical + vbokonly, "Error")
				   Exit Sub
			   End Try
			   Exit For
		   End If
	   Next
	 
	Next
	 

End sub

Sub ExportPDF(dwg As Drawings.DrawingSheet, outputFile As String, units As Integer)
 
   Dim printPDFBuilder1 As PrintPDFBuilder
 
   printPDFBuilder1 = workPart.PlotManager.CreatePrintPdfbuilder()
   printPDFBuilder1.Scale = 1.0
   printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Native
   printPDFBuilder1.Colors = PrintPDFBuilder.Color.BlackOnWhite
   printPDFBuilder1.Size = PrintPDFBuilder.SizeOption.ScaleFactor
   If units = 0 Then
       printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.English
   Else
       printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.Metric
   End If
   printPDFBuilder1.XDimension = dwg.height
   printPDFBuilder1.YDimension = dwg.length
   printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Polylines
   printPDFBuilder1.RasterImages = True
   printPDFBuilder1.ImageResolution = PrintPDFBuilder.ImageResolutionOption.Medium
   printPDFBuilder1.Append = True
 
   Dim sheets1(0) As NXObject
   Dim drawingSheet1 As Drawings.DrawingSheet = CType(dwg, Drawings.DrawingSheet)
 
   sheets1(0) = drawingSheet1
   printPDFBuilder1.SourceBuilder.SetSheets(sheets1)
 
   printPDFBuilder1.Filename = outputFile
 
   Dim nXObject1 As NXObject
   nXObject1 = printPDFBuilder1.Commit()
 
   printPDFBuilder1.Destroy()
 
End Sub

End module
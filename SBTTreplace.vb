Option Strict Off
Imports System
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms
Imports System.Windows.Forms.MessageBox
Imports NXOpen
Imports NXOpen.UF


Module NXJournal
	Sub Main(ByVal args() As String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		For Each foundFile As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "*000.000 СБ.prt")
			SBTTreplace(foundFile)
		Next
	End sub


	Sub SBTTreplace(ByVal SBPath as String)
		'Открываем файл сборки
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim basePart1 As NXOpen.BasePart
		Dim partLoadStatus1 As NXOpen.PartLoadStatus
		basePart1 = theSession.Parts.OpenBaseDisplay(SBPath, partLoadStatus1)
		Dim workPart As NXOpen.Part = theSession.Parts.Work
		Dim displayPart As NXOpen.Part = theSession.Parts.Display

		Dim NewSBTT as string
		NewSBTT = InputBox(SBPath, "1 - 31, 4 - 54", "1")
		
			if NewSBTT = "1" then
				'Замена ТТ
				Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ СБ IP31.txt")
				Dim TTFileContent As String
				Dim TTFileContentX As String
					For Each TTFileContent In readText
						TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
					next
					
					'Добавляем аттрибут IP
					Dim objects(0) As NXObject
					objects(0) = workPart
					Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
					attributePropertiesBuilder1 = workpart.PropertiesManager.CreateAttributePropertiesBuilder(objects)
					attributePropertiesBuilder1.ObjectPicker = AttributePropertiesBaseBuilder.ObjectOptions.ComponentAsPartAttribute
					attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.String
					attributePropertiesBuilder1.Title = "IP"
					attributePropertiesBuilder1.IsArray = False
					attributePropertiesBuilder1.StringValue = "31"
					Dim nXObject1 As NXObject
					nXObject1 = attributePropertiesBuilder1.Commit()
					attributePropertiesBuilder1.Destroy()

				workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
			end if
			if NewSBTT = "4" then
				Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ СБ IP54.txt")
				Dim TTFileContent As String
				Dim TTFileContentX As String
					For Each TTFileContent In readText
						TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
					next
					
					'Добавляем аттрибут IP
					Dim objects(0) As NXObject
					objects(0) = workPart
					Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
					attributePropertiesBuilder1 = workpart.PropertiesManager.CreateAttributePropertiesBuilder(objects)
					attributePropertiesBuilder1.ObjectPicker = AttributePropertiesBaseBuilder.ObjectOptions.ComponentAsPartAttribute
					attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.String
					attributePropertiesBuilder1.Title = "IP"
					attributePropertiesBuilder1.IsArray = False
					attributePropertiesBuilder1.StringValue = "54"
					Dim nXObject1 As NXObject
					nXObject1 = attributePropertiesBuilder1.Commit()
					attributePropertiesBuilder1.Destroy()

				workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
			end if


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



		'Сохранение 
		Dim partSaveStatus1 As NXOpen.PartSaveStatus
		partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.False)
		partSaveStatus1.Dispose()



		'Закрытие 
		Dim partCloseResponses1 As NXOpen.PartCloseResponses
		partCloseResponses1 = theSession.Parts.NewPartCloseResponses()
		workPart.Close(NXOpen.BasePart.CloseWholeTree.True, NXOpen.BasePart.CloseModified.UseResponses, partCloseResponses1)
		workPart = Nothing
		displayPart = Nothing
		partCloseResponses1.Dispose()
	End Sub


Sub ExportPDF(dwg As Drawings.DrawingSheet, outputFile As String, units As Integer)
Dim theSession As Session = Session.GetSession()
Dim workPart As Part = theSession.Parts.Work
Dim displayPart As Part = theSession.Parts.Display


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



Function GetFileName()
Dim theSession As Session = Session.GetSession()
Dim workPart As Part = theSession.Parts.Work
Dim displayPart As Part = theSession.Parts.Display

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
Dim theSession As Session = Session.GetSession()
Dim workPart As Part = theSession.Parts.Work
Dim displayPart As Part = theSession.Parts.Display

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



End module
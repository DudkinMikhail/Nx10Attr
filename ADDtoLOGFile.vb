Option Strict Off
Imports System
Imports System.Text
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms
Imports System.Windows.Forms.MessageBox
Imports NXOpen
Imports NXOpen.UF

Module NXJournal

	Sub Main(ByVal args() As String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		
		' Путь к файлу со списком
		Dim PathToIdList As String = "F:\VB.NET\idList.txt"
		
		' Количество строк/кодов 1С
		Dim IdCodeCount = File.ReadAllLines(PathToIdList).Length
		
		' Создание массива со списком того что следует поменять
		Dim IdListArray(IdCodeCount-1) As String
		
		' Количество артикулов для изменения, начинается с НУЛЯ
		Dim SRi As Double = 0
		
		' Открываем консоль
		Dim lw As ListingWindow = theSession.ListingWindow
		lw.Open()
		
		' Открываем файл со списком кодов 1С, считываем построчно и заполняем массив
		Dim SR As StreamReader = New StreamReader(PathToIdList)
				Do While SR.Peek() >= 0
			IdListArray(SRi) = SR.ReadLine()
			SRi=SRi+1
		loop
			
		
		' Обнуляем счетчик, теперь он пригодится для подсчета количества найденых файлов и задания размера будущего массива
		SRi=0
		For Each foundFile As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "*100.001 Стенка задняя.prt")
			SRi=SRi+1
		Next

		' Объявление массива со списком найденного
		Dim foundFileArrayId(SRi) As String
		Dim foundFileArrayPath(SRi) As String

		' Путь к файлу id.txt.txt
		Dim Path1 As String
		Dim SR1 As StreamReader

		' Еще раз ищем стеночки
		SRi=0
		For Each foundFile1 As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "*100.001 Стенка задняя.prt")
			
			Path1 = Path.GetDirectoryName(foundFile1) & "\id.txt.txt"
			SR1 = My.Computer.FileSystem.OpenTextFileReader(Path1, Encoding.Default)
			
			
			foundFileArrayId(SRi) = SR1.ReadLine()
			foundFileArrayPath(SRi) = foundFile1
			
			SRi=SRi+1
		Next
		
		
		' Ищем одно в другом
		SRi=0
		dim a as double
		Dim ids as string
		For Each ids in IdListArray
			a=Array.IndexOf(foundFileArrayId, IdListArray(SRi))
			if a=-1 then goto Line2
				lw.Writeline(a)
				lw.Writeline("Искалось: " & IdListArray(SRi))
				lw.Writeline("Нашлось: " & foundFileArrayId(a))
				lw.Writeline("Стеночка: " & foundFileArrayPath(a))	
				BackPlateEditMotherFucker(foundFileArrayPath(a))
				lw.Writeline(Path.GetDirectoryName(foundFileArrayPath(a)))
			Line2:
				SRi=SRi+1	
		Next
		
	
	End sub


	Sub BackPlateEditMotherFucker(ByVal BackPlatePath as String)
		' Создаем попку null для модели
		Directory.CreateDirectory(Path.GetDirectoryName(BackPlatePath) & "\" & "null")
		
		' Создаем попку null для PDF
		Directory.CreateDirectory(Path.GetDirectoryName(BackPlatePath) & "\Out\" & "null")
		
		' Копируем в нее уже старый файл модели
		File.Copy(BackPlatePath, Path.GetDirectoryName(BackPlatePath) & "\null\OLD_" & Format(Date.Now, "d.M.yy Hmm") & "_" & Path.GetFileName(BackPlatePath))
		
		' Копируем в нее уже старый файл PDF
		File.Copy(Path.GetDirectoryName(BackPlatePath) & "\Out\" & Path.GetFileNameWithoutExtension(BackPlatePath) & ".pdf", Path.GetDirectoryName(BackPlatePath) & "\Out\null\OLD_" & Format(Date.Now, "d.M.yy Hmm") & "_" & Path.GetFileNameWithoutExtension(BackPlatePath) & ".pdf")
		
		' Удаляем ставший уже ненужным PDF файл, ой опасно, удаляется НАСОВСЕМ
		My.Computer.FileSystem.DeleteFile(Path.GetDirectoryName(BackPlatePath) & "\Out\" & Path.GetFileNameWithoutExtension(BackPlatePath) & ".pdf")
		
		' Создаем файл Log, если уже есть добавляем в него инфу о изменении
		Dim Logpath As String = Path.GetDirectoryName(BackPlatePath) & "\log.txt"
        Dim fs As FileStream = File.Create(Logpath)

        ' Пишем текст
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(Format(Date.Now, "d.M.yy Hmm") & " - Высота и ширина уменьшены на 1 мм")
        fs.Write(info, 0, info.Length)
        fs.Close()		


		'Открываем заднюю стенку
		Dim partLoadStatus As NXOpen.PartLoadStatus
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		theSession.Parts.OpenBaseDisplay(BackPlatePath, partLoadStatus)

        Dim workPart As Part = theSession.Parts.Work
        Dim displayPart As Part = theSession.Parts.Display

        Dim expression1 As Expression = CType(workPart.Expressions.FindObject("H"), Expression)
        Dim expression2 As Expression = CType(workPart.Expressions.FindObject("W"), Expression)
        Dim expression33 As Expression = CType(workPart.Expressions.FindObject("Sheet_Metal_Material_Thickness"), Expression)

		Dim NewBoxH as integer
		Dim NewBoxW as integer
		
		NewBoxH=CInt(expression1.Value.ToString)
		NewBoxW=CInt(expression2.Value.ToString)
		
		
		
		
		'Переключение в режим листового металла
		theSession.ApplicationSwitchImmediate("UG_APP_SBSM")
		Dim expression3 As NXOpen.Expression = CType(workPart.Expressions.FindObject("H"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression3, unit1, NewBoxH-1)
		Dim expression4 As NXOpen.Expression = CType(workPart.Expressions.FindObject("W"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression4, unit2, NewBoxW-1)
		workPart.Preferences.PartSheetmetal.Commit()
		theSession.Preferences.Modeling.UpdatePending = True
	
		Dim MATPARAMSattributeInfo As NXObject.AttributeInformation
		Dim RollMaterialValue As String

		RollMaterialValue = "Рулон <R" & expression33.Value.ToString & "х" & NewBoxW-1 & " ГОСТ 19904-90!08пс ГОСТ 16523-97>"
		workPart.SetUserAttribute("MAT_MARKA", -1, RollMaterialValue, Update.Option.Now)

	
		' Переключение в режим чертежа	
		theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING")

		' Одновление всех видов на чертеже
        For Each tempSheet As Drawings.DrawingSheet In workPart.DrawingSheets
            For Each tempView As Drawings.DraftingView In tempSheet.GetDraftingViews
                If tempView.IsOutOfDate Then
                    tempView.Update()
                End If
            Next
        Next


		' Сохраняем PDF
		GetPDFBitch()
		
		
		'Сохраняем
		Dim partSaveStatus1 As NXOpen.PartSaveStatus
		partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.False)
		partSaveStatus1.Dispose()


		'Закрываем
		Dim partCloseResponses1 As NXOpen.PartCloseResponses
		partCloseResponses1 = theSession.Parts.NewPartCloseResponses()
		workPart.Close(NXOpen.BasePart.CloseWholeTree.True, NXOpen.BasePart.CloseModified.UseResponses, partCloseResponses1)
		workPart = Nothing
		displayPart = Nothing
		partCloseResponses1.Dispose()

	End sub




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


	Sub GetPDFBitch
	Dim theSession As Session = Session.GetSession()
	Dim workPart As Part = theSession.Parts.Work
	Dim displayPart As Part = theSession.Parts.Display

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

	Public Function GetUnloadOption(ByVal dummy As String) As Integer
	GetUnloadOption = NXOpen.Session.LibraryUnloadOption.AtTermination
	End Function

End Module
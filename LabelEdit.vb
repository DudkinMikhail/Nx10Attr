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
		For Each foundFile As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "*001 Этикетка.prt")
			LabelParamEdit(foundFile)
		Next
End sub


Sub LabelParamEdit(ByVal SBPath as String)
		'Открываем файл этикетки
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()

		
		Dim basePart1 As NXOpen.BasePart
		Dim partLoadStatus1 As NXOpen.PartLoadStatus
		basePart1 = theSession.Parts.OpenBaseDisplay(SBPath, partLoadStatus1)
		
		
		Dim workPart As NXOpen.Part = theSession.Parts.Work
		Dim displayPart As NXOpen.Part = theSession.Parts.Display
		
		Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("p6"), NXOpen.Expression)
		Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression1, unit1, "75")

		Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("p7"), NXOpen.Expression)
		Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
		workPart.Expressions.EditWithUnits(expression2, unit2, "105")


		theSession.Preferences.Modeling.UpdatePending = False




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

End module
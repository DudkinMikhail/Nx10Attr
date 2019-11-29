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
		Dim lw As ListingWindow = theSession.ListingWindow
		lw.Open()

		For Each foundFile As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "*100.001 Стенка задняя.prt")
			SquaringTheCircle(foundFile)
		Next
	End sub

	Sub SquaringTheCircle(ByVal SBPath as String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim basePart1 As NXOpen.BasePart
		Dim partLoadStatus1 As NXOpen.PartLoadStatus
		basePart1 = theSession.Parts.OpenBaseDisplay(SBPath, partLoadStatus1)
		Dim workPart As NXOpen.Part = theSession.Parts.Work
		Dim displayPart As NXOpen.Part = theSession.Parts.Display

	Try
		Dim features1(0) As NXOpen.Features.Feature
		Dim extrude1 As NXOpen.Features.Extrude = CType(workPart.Features.FindObject("EXTRUDE(8)"), NXOpen.Features.Extrude)
		features1(0) = extrude1
		workPart.Features.SuppressFeatures(features1)


		Dim features2(0) As NXOpen.Features.Feature
		Dim dimple1 As NXOpen.Features.Dimple = CType(workPart.Features.FindObject("SB_DIMPLE(6)"), NXOpen.Features.Dimple)
		features2(0) = dimple1
		workPart.Features.SuppressFeatures(features2)


		Dim features3(0) As NXOpen.Features.Feature
		Dim dimple2 As NXOpen.Features.Dimple = CType(workPart.Features.FindObject("SB_DIMPLE(7)"), NXOpen.Features.Dimple)
		features3(0) = dimple2
		workPart.Features.UnsuppressFeatures(features3)


		' Изменяем расстояние выдавки
		
		Dim expNameToFind As String = "p259"
        Dim myExp As Expression
 
        myExp = workPart.Expressions.FindObject(expNameToFind)
		if myExp.RightHandSide = 35 then 
			Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("p259"), NXOpen.Expression)
			Dim unit1 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
			workPart.Expressions.EditWithUnits(expression1, unit1, "0")
			theSession.Preferences.Modeling.UpdatePending = False
			
			Dim expression2 As NXOpen.Expression = CType(workPart.Expressions.FindObject("p2138"), NXOpen.Expression)
			Dim unit2 As NXOpen.Unit = CType(workPart.UnitCollection.FindObject("MilliMeter"), NXOpen.Unit)
			workPart.Expressions.EditWithUnits(expression2, unit2, "0")
			theSession.Preferences.Modeling.UpdatePending = False
		end if
		

		Dim features4(0) As NXOpen.Features.Feature
		Dim extrude4 As NXOpen.Features.Extrude = CType(workPart.Features.FindObject("EXTRUDE(8)"), NXOpen.Features.Extrude)
		features4(0) = extrude4
		workPart.Features.UnsuppressFeatures(features4)
		

		'Сохранение 
		Dim partSaveStatus1 As NXOpen.PartSaveStatus
		partSaveStatus1 = workPart.Save(NXOpen.BasePart.SaveComponents.True, NXOpen.BasePart.CloseAfterSave.False)
		partSaveStatus1.Dispose()
		
	Catch e As Exception
		theSession.ListingWindow.WriteLine("Failed: " & e.ToString)
		theSession.ListingWindow.WriteLine("#########: " & SBPath)
	End Try

		
		'Закрытие 
		Dim partCloseResponses1 As NXOpen.PartCloseResponses
		partCloseResponses1 = theSession.Parts.NewPartCloseResponses()
		workPart.Close(NXOpen.BasePart.CloseWholeTree.True, NXOpen.BasePart.CloseModified.UseResponses, partCloseResponses1)
		workPart = Nothing
		displayPart = Nothing
		partCloseResponses1.Dispose()
	End sub

End Module
Option Strict Off
Imports System
Imports System.IO
Imports System.Collections
Imports NXOpen
Imports NXOpen.UF
Imports Microsoft.VisualBasic


Module NXJournal

	Sub Main(ByVal args() As String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim workPart As NXOpen.Part = theSession.Parts.Work
		Dim displayPart As NXOpen.Part = theSession.Parts.Display
		
		For Each foundFile As String In My.Computer.FileSystem.GetFiles("K:\Industrial\КД\Ящики злектрические",FileIO.SearchOption.SearchAllSubDirectories, "*.000.000 СБ.prt")
			DeleteOldVP(foundFile)
			OpenAssemblyAddShield(foundFile)
			'OpenVP(foundFile)
		Next

		
		

	End sub
	
	
	Sub DeleteOldVP(ByVal SBPath as String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim A As String = Path.GetFileNameWithoutExtension(SBPath)
		My.Computer.FileSystem.DeleteFile(Path.GetDirectoryName(SBPath) & "\Out\" & Left(A, A.Length-2) & "ВП.pdf")
	End sub
	
	
	Sub OpenAssemblyAddShield(ByVal SBPath as String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim basePart1 As NXOpen.BasePart
		Dim partLoadStatus1 As NXOpen.PartLoadStatus
		basePart1 = theSession.Parts.OpenBaseDisplay(SBPath, partLoadStatus1)
		Dim workPart As NXOpen.Part = theSession.Parts.Work
		Dim displayPart As NXOpen.Part = theSession.Parts.Display
	

		Dim basePoint1 As NXOpen.Point3d = New NXOpen.Point3d(100, -200, 600)
		Dim orientation1 As NXOpen.Matrix3x3
		orientation1.Xx = 0.0
		orientation1.Xy = -2.32329036961177e-015
		orientation1.Xz = -1.0
		orientation1.Yx = 0.0
		orientation1.Yy = 1.0
		orientation1.Yz = -2.32329036961177e-015
		orientation1.Zx = 1.0
		orientation1.Zy = 0.0
		orientation1.Zz = 0.0
		Dim component1 As NXOpen.Assemblies.Component
		
		component1 = workPart.ComponentAssembly.AddComponent("K:\Industrial\Библиотеки\1 Заимствованные\12 Детали\121 Металл\1211 Плоские\ЭРА.Б1211.006 Товарный знак.prt", "MODEL", "ЭРА.Б1211.006 ТОВАРНЫЙ ЗНАК", basePoint1, orientation1, -1, partLoadStatus1, True)

		partLoadStatus1.Dispose()

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

	End sub
	
	
	Sub OpenVP(ByVal SBPath as String)
		Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
		Dim basePart1 As NXOpen.BasePart
		Dim partLoadStatus1 As NXOpen.PartLoadStatus
		basePart1 = theSession.Parts.OpenBaseDisplay(SBPath, partLoadStatus1)
		Dim workPart As NXOpen.Part = theSession.Parts.Work
		Dim displayPart As NXOpen.Part = theSession.Parts.Display
		theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING")
	End sub


End module
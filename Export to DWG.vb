Option Strict Off
Imports System
Imports NXOpen
Imports System.IO
Imports System.Windows.Forms

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
	Dim dxfdwgCreator1 As NXOpen.DxfdwgCreator
	dxfdwgCreator1 = theSession.DexManager.CreateDxfdwgCreator()
	dxfdwgCreator1.ExportData = NXOpen.DxfdwgCreator.ExportDataOption.Drawing
	dxfdwgCreator1.ViewEditMode = True
	dxfdwgCreator1.FlattenAssembly = True
	dxfdwgCreator1.SettingsFile = "C:\Program Files\Siemens\NX 10.0\dxfdwg\dxfdwg.def"
	dxfdwgCreator1.OutputFileType = NXOpen.DxfdwgCreator.OutputFileTypeOption.Dwg
	dxfdwgCreator1.OutputFile = GetFilePath & "\Out\" & GetFileName() & ".dwg"
	dxfdwgCreator1.ExportData = NXOpen.DxfdwgCreator.ExportDataOption.Modeling
	dxfdwgCreator1.ObjectTypes.Curves = True
	dxfdwgCreator1.ObjectTypes.Annotations = True
	dxfdwgCreator1.ObjectTypes.Structures = True
	dxfdwgCreator1.AutoCADRevision = NXOpen.DxfdwgCreator.AutoCADRevisionOptions.R2010
	dxfdwgCreator1.FlattenAssembly = False
	dxfdwgCreator1.InputFile = GetFilePath & GetFileName() & ".prt"
	dxfdwgCreator1.WidthFactorMode = NXOpen.DxfdwgCreator.WidthfactorMethodOptions.AutomaticCalculation
	dxfdwgCreator1.LayerMask = "1-256"
	dxfdwgCreator1.DrawingList = ""
	dxfdwgCreator1.ViewList = "FLAT-PATTERN#1"

	Dim nXObject1 As NXOpen.NXObject
	nXObject1 = dxfdwgCreator1.Commit()
	dxfdwgCreator1.Destroy()
End Sub

Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
End Function

End Module
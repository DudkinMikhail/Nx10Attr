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

	Dim NewBoxSbody as string
	NewBoxSbody = workPart.Preferences.PartSheetmetal.GetThickness.Value()
	
		if NewBoxSbody = "0,5" then
			'Замена ТТ
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.5.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		if NewBoxSbody = "0,6" then
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.6.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if
		if NewBoxSbody = "0,8" then
			Dim readText() As String = File.ReadAllLines("C:\SWPlus\NX10\Тектовые шаблоны\ТТ Дверка 0.8.txt")
			Dim TTFileContent As String
			Dim TTFileContentX As String
				For Each TTFileContent In readText
					TTFileContentX = TTFileContentX & TTFileContent & Chr(13)
				next
			workPart.SetUserAttribute("TT", -1, TTFileContentX, Update.Option.Now)
		end if


End sub
End module
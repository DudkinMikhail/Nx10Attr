Option Strict Off
Imports System
Imports System.IO
Imports System.Collections
Imports NXOpen
Imports NXOpen.UF
Imports Microsoft.VisualBasic

Module NXJournal

	Sub Main()
	
	Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
	Dim workPart As NXOpen.Part = theSession.Parts.Work
	Dim displayPart As NXOpen.Part = theSession.Parts.Display
	
	Dim NewBoxName, Message1, Message2, Title, DefaultValue, RollMaterialValue, RollMaterialThin, RollMaterialWidth As String
	
	Dim MATPARAMSattributeInfo As NXObject.AttributeInformation
	
	'Постоянные переменные
	Message1 = "Ширина рулона"
	Message2 = "Толщина листа"
	Title = "Do barrel roll"
	DefaultValue = ""

	
	MATPARAMSattributeInfo = workPart.GetUserAttribute("MAT_PARAMS", NXObject.AttributeType.String, -1)
	
	'Ввод толщины
	RollMaterialThin = InputBox(message2, title, MATPARAMSattributeInfo.StringValue)
	
	'Ввод ширины
	RollMaterialWidth = InputBox(message1, title, DefaultValue)
	
	RollMaterialValue = "Рулон <R" & RollMaterialThin & "х" & RollMaterialWidth & " ГОСТ 19904-90!08пс ГОСТ 16523-97>"
	

	workPart.SetUserAttribute("MAT_MARKA", -1, RollMaterialValue, Update.Option.Now)
	
	
	

	End sub

End Module
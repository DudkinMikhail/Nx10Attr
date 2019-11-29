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
	
	Dim Oboznachenie As NXObject.AttributeInformation
	Dim Oboznachenie1 As String
	Dim ItogoVGrammah As String

	
	Oboznachenie = workPart.GetUserAttribute("OBOZNACHENIE", NXObject.AttributeType.String, -1)
	Oboznachenie1 = Oboznachenie.StringValue
	ItogoVGrammah = Right(Oboznachenie1, Oboznachenie1.Length-4)
	

	workPart.SetUserAttribute("OBOZNACHENIE", -1, "ЭРА.SIMPLE " & ItogoVGrammah, Update.Option.Now)
	
	
	

	End sub

End Module
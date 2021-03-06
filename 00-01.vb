Option Strict Off
Imports System
Imports NXOpen

Module NXJournal

Dim theSession As Session = Session.GetSession()
Dim workPart As Part = theSession.Parts.Work
Dim displayPart As Part = theSession.Parts.Display

	Sub Main

	Dim expression0 As NXOpen.Expression = CType(workPart.Expressions.FindObject("Isp_00"), NXOpen.Expression)
	Dim expression1 As NXOpen.Expression = CType(workPart.Expressions.FindObject("Isp_01"), NXOpen.Expression)
	
	Dim nullNXOpen_Unit As NXOpen.Unit = Nothing
		

	if expression0.Value.ToString = "1" then
		workPart.Expressions.EditWithUnits(expression0, nullNXOpen_Unit, "0")
		workPart.Expressions.EditWithUnits(expression1, nullNXOpen_Unit, "1")
	else
		workPart.Expressions.EditWithUnits(expression0, nullNXOpen_Unit, "1")
		workPart.Expressions.EditWithUnits(expression1, nullNXOpen_Unit, "0")
	end if
	
		theSession.Preferences.Modeling.UpdatePending = False
		Dim markId1 As NXOpen.Session.UndoMarkId = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Update Expression Data")
		Dim nErrs1 As Integer = theSession.UpdateManager.DoUpdate(markId1)

	End Sub

End Module
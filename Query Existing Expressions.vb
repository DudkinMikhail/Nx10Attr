Option Strict Off
Imports System
Imports NXOpen
 
Module NXJ_Exp_1
 
    Sub Main()
 
        Dim theSession As Session = Session.GetSession()
        If IsNothing(theSession.Parts.Work) Then
            'active part required
            Return
        End If
 
        Dim workPart As Part = theSession.Parts.Work
        Dim lw As ListingWindow = theSession.ListingWindow
        lw.Open()
 
        Const undoMarkName As String = "NXJ query expressions"
        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, undoMarkName)
 
        For Each temp As Expression In workPart.Expressions
            lw.WriteLine("name: " & temp.Name)
            lw.WriteLine("type: " & temp.Type)
            lw.WriteLine("description: " & temp.Description)
            lw.WriteLine("descriptor: " & temp.GetDescriptor)
            lw.WriteLine("journal identifier: " & temp.JournalIdentifier)
            lw.WriteLine("measurement expression? " & temp.IsMeasurementExpression.ToString)
            lw.WriteLine("geometric expression? " & temp.IsGeometricExpression.ToString)
            lw.WriteLine("edit locked? " & temp.IsNoEdit.ToString)
            lw.WriteLine("locked by user? " & temp.IsUserLocked.ToString)
            lw.WriteLine("right hand side: " & temp.RightHandSide)
            lw.WriteLine("equation: " & temp.Equation)
 
            If temp.RightHandSide.Contains("//") Then
                'expression contains comment
                Dim expComment As String = ""
                expComment = temp.RightHandSide.Substring(temp.RightHandSide.IndexOf("//") + 2)
                lw.WriteLine("comment: " & expComment)
            End If
 
            Select Case temp.Type
                Case Is = "Number"
                    lw.WriteLine("value (base part units): " & temp.Value.ToString)
                    Try
                        lw.WriteLine("expression units: " & temp.Units.Name & " (" & temp.Units.Abbreviation & ")")
                        lw.WriteLine("value in expression units: " & temp.GetValueUsingUnits(Expression.UnitsOption.Expression).ToString)
                    Catch ex As NullReferenceException
                        lw.WriteLine("expression is constant (unitless)")
                    Catch ex2 As Exception
                        lw.WriteLine("!! error: " & ex2.Message)
                    End Try
 
                Case Is = "String"
                    lw.WriteLine("string value: " & temp.StringValue)
 
                Case Is = "Integer"
                    lw.WriteLine("integer value: " & temp.IntegerValue.ToString)
 
                Case Is = "Boolean"
                    lw.WriteLine("boolean value: " & temp.BooleanValue.ToString)
 
                Case Is = "Vector"
                    lw.WriteLine("vector value: " & temp.VectorValue.ToString)
 
                Case Is = "Point"
                    lw.WriteLine("point value: " & temp.PointValue.ToString)
 
                Case Is = "List"
                    'returns a generic object, not especially useful
                    'lw.WriteLine("list value: " & temp.GetListValue.ToString)
 
                    'parse the right hand side to get the list items
                    Dim theList As String = temp.RightHandSide
                    theList = theList.Replace("{", "")
                    theList = theList.Replace("}", "")
                    Dim theItems() As String = theList.Split(",")
                    For Each item As String In theItems
                        lw.WriteLine(item.Trim)
                    Next
 
                Case Else
                    lw.WriteLine("Type: " & temp.Type & " is not handled by this journal")
 
            End Select
 
            Dim theOwningFeature As Features.Feature = temp.GetOwningFeature
            If Not IsNothing(theOwningFeature) Then
                lw.WriteLine("owning feature: " & theOwningFeature.GetFeatureName)
            End If
 
            Dim theOwningRpoFeature As Features.Feature = temp.GetOwningRpoFeature
            If Not IsNothing(theOwningRpoFeature) Then
                lw.WriteLine("owning rpo feature: " & theOwningRpoFeature.GetFeatureName)
            End If
 
            Dim referencingExps() As Expression = temp.GetReferencingExpressions
            If referencingExps.Length > 0 Then
                lw.WriteLine(" $$ used by expression(s):")
                For Each refTemp As Expression In referencingExps
                    lw.WriteLine("    " & refTemp.Name)
                Next
            End If
 
            Dim referencingFeats() As Features.Feature = temp.GetUsingFeatures
            If referencingFeats.Length > 0 Then
                lw.WriteLine(" %% used by feature(s):")
                For Each featTemp As Features.Feature In referencingFeats
                    lw.WriteLine("    " & featTemp.GetFeatureName)
                Next
            End If
 
            lw.WriteLine("")
        Next
 
        lw.Close()
 
    End Sub
 
 
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
 
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
 
    End Function
 
End Module
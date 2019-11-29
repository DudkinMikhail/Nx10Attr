Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.UI
Imports NXOpen.Annotations

Module EditTolerance
    Dim s As Session = Session.GetSession()
    Dim ui As UI = UI.GetUI()
    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim workPart As Part = s.Parts.Work
	
    Sub Main()
        Dim selecteddim As NXObject = Nothing
        Dim response1 As Selection.Response = Selection.Response.Cancel
start1:
        response1 = Select_dim(selecteddim)
        If response1 = Selection.Response.Cancel Then GoTo end1
        EditDimensionTolerance(selecteddim)
        GoTo start1
end1:
    End Sub
    ' ----------------------------------------------
    '   sub to edit tolerance
    ' ----------------------------------------------
    Sub EditDimensionTolerance(ByVal selecteddim As NXObject)
        Dim dimname1 As String = selecteddim.ToString
            Dim linearDimensionBuilder1 As Annotations.LinearDimensionBuilder
            linearDimensionBuilder1 = workPart.Dimensions.CreateLinearDimensionBuilder(selecteddim)
			linearDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 0
			linearDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 0
            Dim nXObject1 As NXObject
            nXObject1 = linearDimensionBuilder1.Commit()    
    End Sub
    ' ----------------------------------------------
    '   function to select dimensions
    ' ----------------------------------------------
    Function Select_dim(ByRef obj As NXObject) As Selection.Response
        Dim resp As Selection.Response = Selection.Response.Cancel
        Dim prompt As String = "Select dimensions"
        Dim message As String = "Select dimensions"
        Dim title As String = "Selection"
        Dim scope As Selection.SelectionScope = Selection.SelectionScope.WorkPart
        Dim selAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific
        Dim selectionMask_array(0) As Selection.MaskTriple
        Dim includeFeatures As Boolean = False
        Dim keepHighlighted As Boolean = False
        With selectionMask_array(0)
            .Type = UFConstants.UF_dimension_type
            .Subtype = 0
            .SolidBodySubtype = 0
        End With
        Dim cursor As Point3d = New Point3d
        resp = ui.SelectionManager.SelectTaggedObject(prompt, message, _
            scope, selAction, includeFeatures, keepHighlighted, selectionMask_array, obj, cursor)
        If resp = Selection.Response.ObjectSelected Or _
           resp = Selection.Response.ObjectSelectedByName Then
            Return Selection.Response.Ok
        Else
            Return Selection.Response.Cancel
        End If
    End Function

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function

End Module 
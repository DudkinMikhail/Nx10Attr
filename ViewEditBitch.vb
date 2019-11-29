Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.UI
Imports NXOpen.Annotations
Imports NXOpen.View

Module EditTolerance
    Dim s As Session = Session.GetSession()
    Dim ui As UI = UI.GetUI()
    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim workPart As Part = s.Parts.Work
	
    Sub Main()
        Dim SelectedVIEW As NXObject = Nothing
        Dim response1 As Selection.Response = Selection.Response.Cancel
start1:
        response1 = Select_VIEW(SelectedVIEW)
        If response1 = Selection.Response.Cancel Then GoTo end1
        EditVIEWProp(SelectedVIEW)
        GoTo start1
end1:
    End Sub
    ' ----------------------------------------------
    '   sub to edit VIEW
    ' ----------------------------------------------
    Sub EditVIEWProp(ByVal SelectedVIEW As NXObject)
        Dim dimname1 As String = SelectedVIEW.ToString
            Dim linearDimensionBuilder1 As Annotations.LinearDimensionBuilder
            linearDimensionBuilder1 = workPart.Dimensions.CreateLinearDimensionBuilder(SelectedVIEW)
            linearDimensionBuilder1.Style.DimensionStyle.ToleranceType = Annotations.ToleranceType.BilateralOneLine
			linearDimensionBuilder1.Style.DimensionStyle.UpperToleranceMetric = 3.0
			linearDimensionBuilder1.Style.DimensionStyle.AngularDimensionValuePrecision = 0
			linearDimensionBuilder1.Style.DimensionStyle.DimensionValuePrecision = 0
            Dim nXObject1 As NXObject
            nXObject1 = linearDimensionBuilder1.Commit()    
    End Sub
    ' ----------------------------------------------
    '   function to select VIEW
    ' ----------------------------------------------
    Function Select_VIEW(ByRef obj As NXObject) As Selection.Response
        Dim resp As Selection.Response = Selection.Response.Cancel
        Dim prompt As String = "Select VIEW"
        Dim message As String = "Select VIEW"
        Dim title As String = "Selection"
        Dim scope As Selection.SelectionScope = Selection.SelectionScope.WorkPart
        Dim selAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific
        Dim selectionMask_array(0) As Selection.MaskTriple
        Dim includeFeatures As Boolean = False
        Dim keepHighlighted As Boolean = False
        With selectionMask_array(0)
            .Type = UFConstants.UF_view_type
            .Subtype = UFConstants.UF_all_subtype
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
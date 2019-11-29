Option Strict Off
Imports System
Imports NXOpen
 
Module Module2
 
    Sub Main()
 
        Dim theSession As Session = Session.GetSession()
        Dim workpart As Part = theSession.Parts.Work
        Dim lw As ListingWindow = theSession.ListingWindow
        lw.Open()
 
        For Each dwg As Drawings.DrawingSheet In workpart.DrawingSheets
            dwg.Open()
            workpart.DeleteRetainedDraftingObjectsInCurrentLayout()
            'other code as necessary
        Next
 
        lw.Close()
 
    End Sub
 
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
 
        'Unloads the image when the NX session terminates
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.AtTermination
 
    End Function
 
End Module
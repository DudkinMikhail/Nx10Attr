		' Одновление всех видов на чертеже
        For Each tempSheet As Drawings.DrawingSheet In workPart.DrawingSheets
            For Each tempView As Drawings.DraftingView In tempSheet.GetDraftingViews
                If tempView.IsOutOfDate Then
                    tempView.Update()
                End If
            Next
        Next

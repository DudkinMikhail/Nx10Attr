' NX 10.0.0.24
' Journal created by dudkin_m on Thu Oct 10 12:55:00 2019 RTZ 2 (зима)
'
Option Strict Off
Imports System
Imports NXOpen

Module NXJournal
Sub Main (ByVal args() As String) 

Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
Dim workPart As NXOpen.Part = theSession.Parts.Work

Dim displayPart As NXOpen.Part = theSession.Parts.Display

' ----------------------------------------------
'   Меню: Вставить->Аннотация->Примечание...
' ----------------------------------------------
Dim markId1 As NXOpen.Session.UndoMarkId
markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Начало")

Dim nullNXOpen_Annotations_SimpleDraftingAid As NXOpen.Annotations.SimpleDraftingAid = Nothing

Dim draftingNoteBuilder1 As NXOpen.Annotations.DraftingNoteBuilder
draftingNoteBuilder1 = workPart.Annotations.CreateDraftingNoteBuilder(nullNXOpen_Annotations_SimpleDraftingAid)

draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

draftingNoteBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter

Dim text1(0) As String
text1(0) = "<$LT> направление движения ленты"
draftingNoteBuilder1.Text.TextBlock.SetText(text1)

draftingNoteBuilder1.TextAlignment = NXOpen.Annotations.DraftingNoteBuilder.TextAlign.BelowTop

theSession.SetUndoMarkName(markId1, "Диалоговое окно Замечание")

draftingNoteBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane

draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

Dim leaderData1 As NXOpen.Annotations.LeaderData
leaderData1 = workPart.Annotations.CreateLeaderData()

leaderData1.StubSize = 1.0

leaderData1.Arrowhead = NXOpen.Annotations.LeaderData.ArrowheadType.FilledArrow

leaderData1.VerticalAttachment = NXOpen.Annotations.LeaderVerticalAttachment.Center

draftingNoteBuilder1.Leader.Leaders.Append(leaderData1)

leaderData1.Arrowhead = NXOpen.Annotations.LeaderData.ArrowheadType.None

leaderData1.StubSide = NXOpen.Annotations.LeaderSide.Inferred

leaderData1.StubSize = 3.0

Dim symbolscale1 As Double
symbolscale1 = draftingNoteBuilder1.Text.TextBlock.SymbolScale

Dim symbolaspectratio1 As Double
symbolaspectratio1 = draftingNoteBuilder1.Text.TextBlock.SymbolAspectRatio

draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

Dim markId2 As NXOpen.Session.UndoMarkId
markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Начало")

theSession.SetUndoMarkName(markId2, "Диалоговое окно Настройки")

theSession.SetUndoMarkVisibility(markId2, Nothing, NXOpen.Session.MarkVisibility.Visible)

' ----------------------------------------------
'   Начало меню Настройки
' ----------------------------------------------
Dim markId3 As NXOpen.Session.UndoMarkId
markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Настройки")

theSession.DeleteUndoMark(markId3, Nothing)

Dim markId4 As NXOpen.Session.UndoMarkId
markId4 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Настройки")

draftingNoteBuilder1.Style.LetteringStyle.Angle = 90.0

theSession.SetUndoMarkName(markId4, "Настройки - Угол текста")

theSession.SetUndoMarkVisibility(markId4, Nothing, NXOpen.Session.MarkVisibility.Visible)

theSession.SetUndoMarkVisibility(markId2, Nothing, NXOpen.Session.MarkVisibility.Invisible)

Dim markId5 As NXOpen.Session.UndoMarkId
markId5 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Настройки")

Dim markId6 As NXOpen.Session.UndoMarkId
markId6 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Настройки")

theSession.DeleteUndoMark(markId6, Nothing)

theSession.SetUndoMarkName(markId2, "Настройки")

theSession.DeleteUndoMark(markId5, Nothing)

theSession.SetUndoMarkVisibility(markId2, Nothing, NXOpen.Session.MarkVisibility.Visible)

theSession.DeleteUndoMark(markId4, Nothing)

theSession.DeleteUndoMark(markId2, Nothing)

Dim assocOrigin1 As NXOpen.Annotations.Annotation.AssociativeOriginData
assocOrigin1.OriginType = NXOpen.Annotations.AssociativeOriginType.Drag
Dim nullNXOpen_View As NXOpen.View = Nothing

assocOrigin1.View = nullNXOpen_View
assocOrigin1.ViewOfGeometry = nullNXOpen_View
Dim nullNXOpen_Point As NXOpen.Point = Nothing

assocOrigin1.PointOnGeometry = nullNXOpen_Point
Dim nullNXOpen_Annotations_Annotation As NXOpen.Annotations.Annotation = Nothing

assocOrigin1.VertAnnotation = nullNXOpen_Annotations_Annotation
assocOrigin1.VertAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
assocOrigin1.HorizAnnotation = nullNXOpen_Annotations_Annotation
assocOrigin1.HorizAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
assocOrigin1.AlignedAnnotation = nullNXOpen_Annotations_Annotation
assocOrigin1.DimensionLine = 0
assocOrigin1.AssociatedView = nullNXOpen_View
assocOrigin1.AssociatedPoint = nullNXOpen_Point
assocOrigin1.OffsetAnnotation = nullNXOpen_Annotations_Annotation
assocOrigin1.OffsetAlignmentPosition = NXOpen.Annotations.AlignmentPosition.TopLeft
assocOrigin1.XOffsetFactor = 0.0
assocOrigin1.YOffsetFactor = 0.0
assocOrigin1.StackAlignmentPosition = NXOpen.Annotations.StackAlignmentPosition.Above
draftingNoteBuilder1.Origin.SetAssociativeOrigin(assocOrigin1)

Dim point1 As NXOpen.Point3d = New NXOpen.Point3d(572.626410531972, 254.76182160129, 0.0)
draftingNoteBuilder1.Origin.Origin.SetValue(Nothing, nullNXOpen_View, point1)

draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(True)

Dim markId7 As NXOpen.Session.UndoMarkId
markId7 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Замечание")

Dim nXObject1 As NXOpen.NXObject
nXObject1 = draftingNoteBuilder1.Commit()

theSession.DeleteUndoMark(markId7, Nothing)

theSession.SetUndoMarkName(markId1, "Замечание")

theSession.SetUndoMarkVisibility(markId1, Nothing, NXOpen.Session.MarkVisibility.Visible)

draftingNoteBuilder1.Destroy()

Dim markId8 As NXOpen.Session.UndoMarkId
markId8 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Start")

Dim draftingNoteBuilder2 As NXOpen.Annotations.DraftingNoteBuilder
draftingNoteBuilder2 = workPart.Annotations.CreateDraftingNoteBuilder(nullNXOpen_Annotations_SimpleDraftingAid)

draftingNoteBuilder2.Origin.SetInferRelativeToGeometry(True)

draftingNoteBuilder2.Origin.SetInferRelativeToGeometry(True)

draftingNoteBuilder2.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter

Dim text2(0) As String
text2(0) = "<$LT> направление движения ленты"
draftingNoteBuilder2.Text.TextBlock.SetText(text2)

draftingNoteBuilder2.Style.LetteringStyle.Angle = 90.0

draftingNoteBuilder2.TextAlignment = NXOpen.Annotations.DraftingNoteBuilder.TextAlign.BelowTop

theSession.SetUndoMarkName(markId8, "Диалоговое окно Замечание")

draftingNoteBuilder2.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane

draftingNoteBuilder2.Origin.SetInferRelativeToGeometry(True)

Dim leaderData2 As NXOpen.Annotations.LeaderData
leaderData2 = workPart.Annotations.CreateLeaderData()

leaderData2.StubSize = 1.0

leaderData2.Arrowhead = NXOpen.Annotations.LeaderData.ArrowheadType.FilledArrow

leaderData2.VerticalAttachment = NXOpen.Annotations.LeaderVerticalAttachment.Center

draftingNoteBuilder2.Leader.Leaders.Append(leaderData2)

leaderData2.Arrowhead = NXOpen.Annotations.LeaderData.ArrowheadType.None

leaderData2.StubSide = NXOpen.Annotations.LeaderSide.Inferred

leaderData2.StubSize = 3.0

Dim symbolscale2 As Double
symbolscale2 = draftingNoteBuilder2.Text.TextBlock.SymbolScale

Dim symbolaspectratio2 As Double
symbolaspectratio2 = draftingNoteBuilder2.Text.TextBlock.SymbolAspectRatio

draftingNoteBuilder2.Origin.SetInferRelativeToGeometry(True)

draftingNoteBuilder2.Origin.SetInferRelativeToGeometry(True)

draftingNoteBuilder2.Destroy()

theSession.UndoToMark(markId8, Nothing)

theSession.DeleteUndoMark(markId8, Nothing)

' ----------------------------------------------
'   Меню: Инструменты->Журнал->Остановка записи
' ----------------------------------------------

End Sub
End Module
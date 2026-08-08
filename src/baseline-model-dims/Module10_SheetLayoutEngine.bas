Option Explicit

' A7/C4 sheet-aware final placement.
' Every coordinate in this module is in metres in the drawing sheet frame.
' The engine first tries the approved scales unchanged. If the complete set of
' measured envelopes cannot be placed, it reduces every view by the same
' factor, rebuilds, remeasures, and retries. A failed retry restores the
' original scales and positions.

' swZoneMargin_e, SOLIDWORKS API Help / MCP checked 2026-08-08.
Private Const swZoneTopMargin As Long = 0
Private Const swZoneBottomMargin As Long = 1
Private Const swZoneRightMargin As Long = 2
Private Const swZoneLeftMargin As Long = 3

' swSketchSegments_e, SOLIDWORKS API Help / MCP checked 2026-08-08.
Private Const swSketchLINE As Long = 0

Private Const DEFAULT_BORDER_MARGIN_M As Double = 0.008
Private Const VIEW_CLEARANCE_M As Double = 0.006
Private Const PROTECTED_CLEARANCE_M As Double = 0.004
' Planning reserve only. Final validation still enforces the exact clearances
' above. The r42 six-view run proved that a rebuild can grow a measured
' annotation envelope by 0.191 mm after placement; exact-edge plans therefore
' need explicit headroom instead of relying on POSITION_TOLERANCE_M.
Private Const PLACEMENT_STABILITY_BUFFER_M As Double = 0.001
Private Const TEXT_WIDTH_FACTOR As Double = 0.65
Private Const MIN_TEXT_HEIGHT_M As Double = 0.0025
Private Const SCALE_TOLERANCE As Double = 0.0000001
Private Const POSITION_TOLERANCE_M As Double = 0.00001
Private Const TEMPLATE_GEOMETRY_TOLERANCE_M As Double = 0.00001
Private Const MAX_PROTECTED_RECTS As Long = 256
Private Const SCALE_REDUCTION_FACTOR As Double = 0.9
Private Const MIN_VIEW_SCALE_DECIMAL As Double = 0.25
Private Const MAX_SCALE_ATTEMPTS As Long = 8

Private Type LayoutRect
    LeftX As Double
    BottomY As Double
    RightX As Double
    TopY As Double
    Seeded As Boolean
End Type

Private Type ProtectedRect
    Name As String
    Bounds As LayoutRect
End Type

Public LastLayoutPassed As Boolean
Public LastLayoutReport As String

Public Function ArrangeDrawingContent( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc) As Boolean

    On Error GoTo Failed

    ResetLayoutEvidence

    If swDrawModel Is Nothing Or swDraw Is Nothing Then
        RecordFailure "DrawingUnavailable"
        Exit Function
    End If

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then
        RecordFailure "CurrentSheetUnavailable"
        Exit Function
    End If

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    If Not ReadSheetSize(swSheet, sheetWidth, sheetHeight) Then
        RecordFailure "SheetSizeUnavailable"
        Exit Function
    End If

    Dim usable As LayoutRect
    Dim marginSource As String
    If Not MeasureUsableRectangle( _
        swSheet, sheetWidth, sheetHeight, usable, marginSource) Then
        RecordFailure "UsableRectangleUnavailable"
        Exit Function
    End If

    Dim protected(1 To MAX_PROTECTED_RECTS) As ProtectedRect
    Dim protectedCount As Long
    Dim titleSource As String
    If Not AddMeasuredTitleBlock( _
        swSheet, sheetWidth, sheetHeight, protected, protectedCount, _
        titleSource) Then
        RecordFailure "TitleBlockExtentUnavailable"
        Exit Function
    End If

    Dim swSheetView As SldWorks.View
    Set swSheetView = swDraw.GetFirstView
    If swSheetView Is Nothing Then
        RecordFailure "SheetViewUnavailable"
        Exit Function
    End If

    If Not AddSheetNoteObstacles( _
        swSheetView, protected(1).Bounds, protected, protectedCount) Then
        RecordFailure "SheetNoteExtentUnavailable"
        Exit Function
    End If

    AppendReport "LAYOUT_SHEET|size=" & _
        Format$(sheetWidth, "0.000000") & "," & _
        Format$(sheetHeight, "0.000000") & _
        "|usable=" & FormatRect(usable) & _
        "|margins=" & marginSource & _
        "|title=" & titleSource

    Dim protectedIndex As Long
    For protectedIndex = 1 To protectedCount
        AppendReport "LAYOUT_PROTECTED|index=" & CStr(protectedIndex) & _
            "|name=" & CleanToken(protected(protectedIndex).Name) & _
            "|bounds=" & FormatRect(protected(protectedIndex).Bounds)
    Next protectedIndex

    Dim views As Collection
    Set views = New Collection

    Dim swView As SldWorks.View
    Set swView = swSheetView.GetNextView
    Do While Not swView Is Nothing
        views.Add swView
        Set swView = swView.GetNextView
    Loop

    If views.Count = 0 Then
        RecordFailure "NoDrawingViews"
        Exit Function
    End If

    Dim bounds() As LayoutRect
    Dim planned() As LayoutRect
    Dim scales() As Double
    Dim plannedScales() As Double
    Dim originalPositions() As Variant
    Dim roles() As Long
    ReDim bounds(1 To views.Count)
    ReDim planned(1 To views.Count)
    ReDim scales(1 To views.Count)
    ReDim plannedScales(1 To views.Count)
    ReDim originalPositions(1 To views.Count)
    ReDim roles(1 To views.Count)

    Dim i As Long
    For i = 1 To views.Count
        Set swView = views(i)
        If Not BuildViewEnvelope(swView, bounds(i)) Then
            RecordFailure "ViewEnvelopeUnavailable:" & SafeViewName(swView)
            Exit Function
        End If

        scales(i) = CDbl(swView.ScaleDecimal)
        plannedScales(i) = scales(i)
        originalPositions(i) = swView.Position
        roles(i) = Module8_ViewClassifier.ClassifyView(swView)

        AppendReport "LAYOUT_INPUT|view=" & CleanToken(SafeViewName(swView)) & _
            "|role=" & Module8_ViewClassifier.RoleName(roles(i)) & _
            "|bounds=" & FormatRect(bounds(i)) & _
            "|width=" & Format$(RectWidth(bounds(i)), "0.000000") & _
            "|height=" & Format$(RectHeight(bounds(i)), "0.000000") & _
            "|scale=" & Format$(scales(i), "0.000000")

        Dim isLocked As Boolean
        isLocked = swView.PositionLocked
        If Not (isLocked = False) Then
            RecordFailure "ViewPositionLocked:" & SafeViewName(swView)
            Exit Function
        End If
    Next i

    Dim scaleFactor As Double
    Dim scaleChanged As Boolean
    If Not PlanWithScaleFallback( _
        swDrawModel, views, bounds, roles, scales, usable, protected, _
        protectedCount, planned, plannedScales, scaleFactor, _
        scaleChanged) Then
        RecordFailure "NoCollisionFreePlacementAtMinimumScale"
        Exit Function
    End If

    ' Every position is planned before the first move. Scale fallback can
    ' rebuild while searching; any failed search or final validation restores
    ' the captured scales and positions.
    For i = 1 To views.Count
        Set swView = views(i)
        If Not MoveViewByEnvelopeDelta(swView, bounds(i), planned(i)) Then
            RestoreLayoutState swDrawModel, views, scales, originalPositions
            RecordFailure "ViewMoveFailed:" & SafeViewName(swView)
            Exit Function
        End If
    Next i

    Dim rebuildSucceeded As Boolean
    rebuildSucceeded = swDrawModel.ForceRebuild3(False)
    If rebuildSucceeded = False Then
        RestoreLayoutState swDrawModel, views, scales, originalPositions
        RecordFailure "ForceRebuild3ReturnedFalse"
        Exit Function
    End If

    If Not ValidateFinalLayout( _
        views, usable, protected, protectedCount, plannedScales) Then
        RestoreLayoutState swDrawModel, views, scales, originalPositions
        Exit Function
    End If

    LastLayoutPassed = True
    ArrangeDrawingContent = True
    AppendReport "SHEET_LAYOUT|status=PASS|views=" & CStr(views.Count) & _
        "|protected=" & CStr(protectedCount) & _
        "|margins=" & marginSource & _
        "|title=" & titleSource & _
        "|scaleChanged=" & CStr(scaleChanged) & _
        "|scaleFactor=" & Format$(scaleFactor, "0.000000")
    Exit Function

Failed:
    RecordFailure "APIError:" & CStr(Err.Number) & ":" & CleanToken(Err.Description)
End Function

Public Function DescribeLayout() As String
    DescribeLayout = LastLayoutReport
End Function

Private Sub ResetLayoutEvidence()
    LastLayoutPassed = False
    LastLayoutReport = vbNullString
End Sub

Private Sub RecordFailure(ByVal reason As String)
    LastLayoutPassed = False
    AppendReport "SHEET_LAYOUT|status=FAIL|reason=" & CleanToken(reason)
End Sub

Private Sub AppendReport(ByVal message As String)
    Debug.Print message
    If Len(LastLayoutReport) > 0 Then
        LastLayoutReport = LastLayoutReport & vbCrLf
    End If
    LastLayoutReport = LastLayoutReport & message
End Sub

Private Function ReadSheetSize( _
    ByRef swSheet As SldWorks.Sheet, _
    ByRef sheetWidth As Double, _
    ByRef sheetHeight As Double) As Boolean

    On Error GoTo Failed
    swSheet.GetSize sheetWidth, sheetHeight
    ReadSheetSize = (sheetWidth > 0# And sheetHeight > 0#)
    Exit Function

Failed:
    ReadSheetSize = False
End Function

Private Function MeasureUsableRectangle( _
    ByRef swSheet As SldWorks.Sheet, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef usable As LayoutRect, _
    ByRef source As String) As Boolean

    On Error GoTo FallbackMargins

    Dim topMargin As Double
    Dim bottomMargin As Double
    Dim rightMargin As Double
    Dim leftMargin As Double

    topMargin = swSheet.GetZoneMargin(swZoneTopMargin)
    bottomMargin = swSheet.GetZoneMargin(swZoneBottomMargin)
    rightMargin = swSheet.GetZoneMargin(swZoneRightMargin)
    leftMargin = swSheet.GetZoneMargin(swZoneLeftMargin)

    If topMargin <= 0# Or bottomMargin <= 0# Or _
       rightMargin <= 0# Or leftMargin <= 0# Then
        GoTo FallbackMargins
    End If

    SetRect usable, leftMargin, bottomMargin, _
        sheetWidth - rightMargin, sheetHeight - topMargin
    source = "ISheet.GetZoneMargin"
    MeasureUsableRectangle = RectIsValid(usable)
    Exit Function

FallbackMargins:
    Err.Clear
    SetRect usable, DEFAULT_BORDER_MARGIN_M, DEFAULT_BORDER_MARGIN_M, _
        sheetWidth - DEFAULT_BORDER_MARGIN_M, _
        sheetHeight - DEFAULT_BORDER_MARGIN_M
    source = "Conservative8mmFallback"
    MeasureUsableRectangle = RectIsValid(usable)
End Function

Private Function AddMeasuredTitleBlock( _
    ByRef swSheet As SldWorks.Sheet, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef protected() As ProtectedRect, _
    ByRef protectedCount As Long, _
    ByRef source As String) As Boolean

    On Error GoTo TryTemplateSketch

    Dim titleBlock As SldWorks.TitleBlock
    Set titleBlock = swSheet.TitleBlock
    If titleBlock Is Nothing Then GoTo TryTemplateSketch

    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    titleBlock.GetExtents x1, y1, x2, y2

    Dim titleBounds As LayoutRect
    SetRect titleBounds, MinDouble(x1, x2), MinDouble(y1, y2), _
        MaxDouble(x1, x2), MaxDouble(y1, y2)
    If Not RectIsValid(titleBounds) Then GoTo TryTemplateSketch

    AddProtectedRect protected, protectedCount, "TitleBlock", titleBounds
    source = "ITitleBlock.GetExtents"
    AddMeasuredTitleBlock = True
    Exit Function

TryTemplateSketch:
    Err.Clear
    Dim measured As LayoutRect
    If MeasureTemplateTitleBlock( _
        swSheet, sheetWidth, sheetHeight, measured) Then
        AddProtectedRect protected, protectedCount, _
            "TitleBlockTemplateSketch", measured
        source = "ISheet.GetTemplateSketch measured lines"
        AddMeasuredTitleBlock = True
    End If
End Function

' Some controlled templates expose no native ITitleBlock. In that case the
' sheet-format sketch is measured, not guessed. Only line segments whose two
' endpoints are in the lower-right title-block region are accepted, and the
' resulting rectangle must have enough contributing geometry and remain
' anchored to the lower-right sheet corner.
Private Function MeasureTemplateTitleBlock( _
    ByRef swSheet As SldWorks.Sheet, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef measured As LayoutRect) As Boolean

    On Error GoTo Failed

    Dim templateSketch As SldWorks.Sketch
    Set templateSketch = swSheet.GetTemplateSketch
    If templateSketch Is Nothing Then Exit Function

    Dim segments As Variant
    segments = templateSketch.GetSketchSegments
    If Not IsArray(segments) Then Exit Function

    Dim zoneLeft As Double
    Dim zoneRight As Double
    Dim zoneBottom As Double
    Dim zoneTop As Double
    zoneLeft = sheetWidth * 0.55
    zoneRight = sheetWidth * 0.99
    zoneBottom = sheetHeight * 0.015
    zoneTop = sheetHeight * 0.35

    Dim accepted As Long
    Dim bottomBoundaryProved As Boolean
    Dim topBoundaryProved As Boolean
    Dim leftBoundaryProved As Boolean
    Dim rightBoundaryProved As Boolean
    Dim i As Long
    For i = LBound(segments) To UBound(segments)
        Dim segment As SldWorks.SketchSegment
        Set segment = segments(i)
        If segment Is Nothing Then GoTo ContinueSegment
        If segment.GetType <> swSketchLINE Then GoTo ContinueSegment

        ' GetSketchSegments returns ISketchSegment. GetStartPoint2 and
        ' GetEndPoint2 belong to ISketchLine, so the type check and cast are
        ' both required. Calling the endpoints late-bound on ISketchSegment
        ' was the r34 defect and accepted zero template lines.
        Dim sketchLine As SldWorks.SketchLine
        Set sketchLine = segment

        Dim startPoint As SldWorks.SketchPoint
        Dim endPoint As SldWorks.SketchPoint
        Set startPoint = sketchLine.GetStartPoint2
        Set endPoint = sketchLine.GetEndPoint2

        If startPoint Is Nothing Or endPoint Is Nothing Then _
            GoTo ContinueSegment

        Dim startX As Double
        Dim startY As Double
        Dim endX As Double
        Dim endY As Double
        startX = CDbl(startPoint.X)
        startY = CDbl(startPoint.Y)
        endX = CDbl(endPoint.X)
        endY = CDbl(endPoint.Y)

        If startX >= zoneLeft - TEMPLATE_GEOMETRY_TOLERANCE_M And _
           endX >= zoneLeft - TEMPLATE_GEOMETRY_TOLERANCE_M And _
           startX <= zoneRight + TEMPLATE_GEOMETRY_TOLERANCE_M And _
           endX <= zoneRight + TEMPLATE_GEOMETRY_TOLERANCE_M And _
           startY >= zoneBottom - TEMPLATE_GEOMETRY_TOLERANCE_M And _
           endY >= zoneBottom - TEMPLATE_GEOMETRY_TOLERANCE_M And _
           startY <= zoneTop + TEMPLATE_GEOMETRY_TOLERANCE_M And _
           endY <= zoneTop + TEMPLATE_GEOMETRY_TOLERANCE_M Then
            AddPoint measured, startX, startY
            AddPoint measured, endX, endY
            accepted = accepted + 1

            Dim lineLength As Double
            lineLength = Sqr((endX - startX) * (endX - startX) + _
                (endY - startY) * (endY - startY))

            If Abs(endY - startY) <= TEMPLATE_GEOMETRY_TOLERANCE_M And _
               lineLength >= sheetWidth * 0.24 Then
                If startY <= sheetHeight * 0.1 Then
                    bottomBoundaryProved = True
                ElseIf startY >= sheetHeight * 0.15 Then
                    topBoundaryProved = True
                End If
            End If

            If Abs(endX - startX) <= TEMPLATE_GEOMETRY_TOLERANCE_M And _
               lineLength >= sheetHeight * 0.13 Then
                If startX <= sheetWidth * 0.75 Then
                    leftBoundaryProved = True
                ElseIf startX >= sheetWidth * 0.95 Then
                    rightBoundaryProved = True
                End If
            End If
        End If

ContinueSegment:
    Next i

    If accepted < 10 Or Not RectIsValid(measured) Or _
       Not bottomBoundaryProved Or Not topBoundaryProved Or _
       Not leftBoundaryProved Or Not rightBoundaryProved Or _
       RectWidth(measured) < sheetWidth * 0.22 Or _
       RectWidth(measured) > sheetWidth * 0.5 Or _
       RectHeight(measured) < sheetHeight * 0.12 Or _
       RectHeight(measured) > sheetHeight * 0.35 Or _
       measured.RightX < sheetWidth * 0.95 Or _
       measured.BottomY > sheetHeight * 0.1 Then

        AppendReport "TITLE_BLOCK_MEASURE|status=REJECT|segments=" & _
            CStr(accepted) & "|bounds=" & FormatRect(measured) & _
            "|bottomEdge=" & CStr(bottomBoundaryProved) & _
            "|topEdge=" & CStr(topBoundaryProved) & _
            "|leftEdge=" & CStr(leftBoundaryProved) & _
            "|rightEdge=" & CStr(rightBoundaryProved)
        Exit Function
    End If

    AppendReport "TITLE_BLOCK_MEASURE|status=PASS|segments=" & _
        CStr(accepted) & "|bounds=" & FormatRect(measured) & _
        "|structuralEdges=True"

    MeasureTemplateTitleBlock = True
    Exit Function

Failed:
    AppendReport "TITLE_BLOCK_MEASURE|status=ERROR|error=" & _
        CStr(Err.Number) & ":" & CleanToken(Err.Description)
    MeasureTemplateTitleBlock = False
End Function

Private Function AddSheetNoteObstacles( _
    ByRef swSheetView As SldWorks.View, _
    ByRef titleBounds As LayoutRect, _
    ByRef protected() As ProtectedRect, _
    ByRef protectedCount As Long) As Boolean

    On Error GoTo Failed

    Dim notes As Variant
    notes = swSheetView.GetNotes
    If IsEmpty(notes) Then
        AddSheetNoteObstacles = True
        Exit Function
    End If
    If Not IsArray(notes) Then Exit Function

    Dim i As Long
    For i = LBound(notes) To UBound(notes)
        Dim swNote As SldWorks.Note
        Set swNote = notes(i)
        If swNote Is Nothing Then GoTo ContinueNote

        Dim noteBounds As LayoutRect
        If ReadNoteExtent(swNote, noteBounds) Then
            AddProtectedRect protected, protectedCount, _
                "SheetNote" & CStr(i), noteBounds
        Else
            Dim fallbackStatus As String
            If Not MeasureSheetNoteFallback( _
                swNote, titleBounds, noteBounds, fallbackStatus) Then
                AppendReport "SHEET_NOTE_MEASURE|index=" & CStr(i) & _
                    "|status=REJECT|reason=" & fallbackStatus
                Exit Function
            End If

            If noteBounds.Seeded Then
                AddProtectedRect protected, protectedCount, _
                    "SheetNoteFallback" & CStr(i), noteBounds
            End If

            AppendReport "SHEET_NOTE_MEASURE|index=" & CStr(i) & _
                "|status=" & fallbackStatus & _
                "|bounds=" & FormatRect(noteBounds)
        End If

ContinueNote:
    Next i

    AddSheetNoteObstacles = True
    Exit Function

Failed:
    AppendReport "SHEET_NOTE_MEASURE|status=ERROR|error=" & _
        CStr(Err.Number) & ":" & CleanToken(Err.Description)
    AddSheetNoteObstacles = False
End Function

' INote.GetExtent can decline for sheet-format notes. Those notes are safe to
' skip only when their documented annotation position lies inside the already
' measured title-block rectangle. Any other note receives a conservative
' text rectangle from its position, rendered text, and ITextFormat.CharHeight.
Private Function MeasureSheetNoteFallback( _
    ByRef swNote As SldWorks.Note, _
    ByRef titleBounds As LayoutRect, _
    ByRef bounds As LayoutRect, _
    ByRef status As String) As Boolean

    On Error GoTo Failed

    Dim annotation As SldWorks.Annotation
    Set annotation = swNote.GetAnnotation
    If annotation Is Nothing Then
        status = "AnnotationUnavailable"
        Exit Function
    End If

    Dim position As Variant
    position = annotation.GetPosition
    If Not IsArray(position) Then
        status = "PositionUnavailable"
        Exit Function
    End If
    If UBound(position) - LBound(position) + 1 < 2 Then
        status = "PositionTooShort"
        Exit Function
    End If

    Dim positionX As Double
    Dim positionY As Double
    positionX = CDbl(position(LBound(position)))
    positionY = CDbl(position(LBound(position) + 1))

    If PointInsideRect(positionX, positionY, titleBounds) Then
        status = "COVERED_BY_TITLE_BLOCK"
        MeasureSheetNoteFallback = True
        Exit Function
    End If

    Dim noteText As String
    noteText = CStr(swNote.GetText)
    If Len(noteText) = 0 Then
        status = "TextUnavailable"
        Exit Function
    End If

    Dim charHeight As Double
    charHeight = MIN_TEXT_HEIGHT_M

    On Error Resume Next
    Dim textFormat As SldWorks.TextFormat
    Set textFormat = annotation.GetTextFormat(0)
    If Not textFormat Is Nothing Then
        If CDbl(textFormat.CharHeight) > 0# Then
            charHeight = CDbl(textFormat.CharHeight)
        End If
    End If
    Err.Clear
    On Error GoTo Failed

    Dim normalized As String
    normalized = Replace(noteText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(normalized, vbLf)

    Dim lineCount As Long
    Dim maxCharacters As Long
    lineCount = UBound(lines) - LBound(lines) + 1

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Len(CStr(lines(i))) > maxCharacters Then
            maxCharacters = Len(CStr(lines(i)))
        End If
    Next i

    Dim estimatedWidth As Double
    Dim estimatedHeight As Double
    estimatedWidth = CDbl(maxCharacters) * charHeight * TEXT_WIDTH_FACTOR
    estimatedHeight = CDbl(lineCount) * charHeight * 1.25
    If estimatedWidth < charHeight Then estimatedWidth = charHeight
    If estimatedHeight < charHeight Then estimatedHeight = charHeight

    ' Annotation anchor semantics vary with note justification. Expand by the
    ' full estimated size in every direction instead of assuming one corner.
    SetRect bounds, positionX - estimatedWidth, _
        positionY - estimatedHeight, positionX + estimatedWidth, _
        positionY + estimatedHeight

    status = "TEXT_FALLBACK"
    MeasureSheetNoteFallback = True
    Exit Function

Failed:
    status = "APIError:" & CStr(Err.Number) & ":" & _
        CleanToken(Err.Description)
End Function

Private Function BuildViewEnvelope( _
    ByRef swView As SldWorks.View, _
    ByRef bounds As LayoutRect) As Boolean

    On Error GoTo Failed

    Dim outline As Variant
    outline = swView.GetOutline
    If Not IsArray(outline) Then Exit Function
    If UBound(outline) - LBound(outline) + 1 < 4 Then Exit Function

    Dim base As Long
    base = LBound(outline)
    SetRect bounds, CDbl(outline(base)), CDbl(outline(base + 1)), _
        CDbl(outline(base + 2)), CDbl(outline(base + 3))
    If Not RectIsValid(bounds) Then Exit Function

    If Not AddDisplayDimensionsToEnvelope(swView, bounds) Then Exit Function
    If Not AddViewNotesToEnvelope(swView, bounds) Then Exit Function

    BuildViewEnvelope = RectIsValid(bounds)
    Exit Function

Failed:
    BuildViewEnvelope = False
End Function

Private Function AddDisplayDimensionsToEnvelope( _
    ByRef swView As SldWorks.View, _
    ByRef bounds As LayoutRect) As Boolean

    On Error GoTo Failed

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions
    If IsEmpty(dimensions) Then
        AddDisplayDimensionsToEnvelope = True
        Exit Function
    End If
    If Not IsArray(dimensions) Then Exit Function

    Dim i As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)
        If displayDimension Is Nothing Then GoTo ContinueDimension

        Dim annotation As SldWorks.Annotation
        Set annotation = displayDimension.GetAnnotation
        If annotation Is Nothing Then Exit Function

        Dim originX As Double
        Dim originY As Double
        If Not AddAnnotationGeometry(annotation, bounds, originX, originY) Then
            Exit Function
        End If

        Dim displayData As SldWorks.DisplayData
        Set displayData = displayDimension.GetDisplayData
        If displayData Is Nothing Then Exit Function
        If Not AddDisplayTextGeometry( _
            displayData, originX, originY, bounds) Then Exit Function

ContinueDimension:
    Next i

    AddDisplayDimensionsToEnvelope = True
    Exit Function

Failed:
    AddDisplayDimensionsToEnvelope = False
End Function

Private Function AddViewNotesToEnvelope( _
    ByRef swView As SldWorks.View, _
    ByRef bounds As LayoutRect) As Boolean

    On Error GoTo Failed

    Dim notes As Variant
    notes = swView.GetNotes
    If IsEmpty(notes) Then
        AddViewNotesToEnvelope = True
        Exit Function
    End If
    If Not IsArray(notes) Then Exit Function

    Dim i As Long
    For i = LBound(notes) To UBound(notes)
        Dim swNote As SldWorks.Note
        Set swNote = notes(i)
        If swNote Is Nothing Then GoTo ContinueNote

        Dim noteBounds As LayoutRect
        If Not ReadNoteExtent(swNote, noteBounds) Then Exit Function
        AddRect bounds, noteBounds

        Dim annotation As SldWorks.Annotation
        Set annotation = swNote.GetAnnotation
        If Not annotation Is Nothing Then
            Dim originX As Double
            Dim originY As Double
            If Not AddAnnotationGeometry( _
                annotation, bounds, originX, originY) Then Exit Function
        End If

ContinueNote:
    Next i

    AddViewNotesToEnvelope = True
    Exit Function

Failed:
    AddViewNotesToEnvelope = False
End Function

Private Function ReadNoteExtent( _
    ByRef swNote As SldWorks.Note, _
    ByRef bounds As LayoutRect) As Boolean

    On Error GoTo Failed

    Dim extent As Variant
    extent = swNote.GetExtent
    If Not IsArray(extent) Then Exit Function
    If UBound(extent) - LBound(extent) + 1 < 6 Then Exit Function

    Dim base As Long
    base = LBound(extent)
    SetRect bounds, CDbl(extent(base)), CDbl(extent(base + 1)), _
        CDbl(extent(base + 3)), CDbl(extent(base + 4))
    ReadNoteExtent = RectIsValid(bounds)
    Exit Function

Failed:
    ReadNoteExtent = False
End Function

Private Function AddAnnotationGeometry( _
    ByRef annotation As SldWorks.Annotation, _
    ByRef bounds As LayoutRect, _
    ByRef originX As Double, _
    ByRef originY As Double) As Boolean

    On Error GoTo Failed

    Dim position As Variant
    position = annotation.GetPosition
    If Not IsArray(position) Then Exit Function
    If UBound(position) - LBound(position) + 1 < 2 Then Exit Function

    originX = CDbl(position(LBound(position)))
    originY = CDbl(position(LBound(position) + 1))
    AddPoint bounds, originX, originY

    AddAnnotationGeometry = True
    Exit Function

Failed:
    AddAnnotationGeometry = False
End Function

' Live r38 proved that GetTextPositionAtIndex cannot be added to
' IAnnotation.GetPosition as though both share one origin: doing so put the
' front envelope at X=0.598795 and the section at X=0.901541 on a 0.42 m
' sheet, while the full-sheet screenshot showed neither there. The documented
' annotation position is used as the text anchor. Text content and height only
' determine a conservative box around it. Display-data line and leader-point
' coordinates are also excluded because their drawing frame is not stated;
' the union of view outline and text/note endpoints already bounds each leader.
Private Function AddDisplayTextGeometry( _
    ByRef displayData As SldWorks.DisplayData, _
    ByVal originX As Double, _
    ByVal originY As Double, _
    ByRef bounds As LayoutRect) As Boolean

    On Error GoTo Failed

    Dim textCount As Long
    textCount = CLng(displayData.GetTextCount)

    Dim i As Long
    For i = 0 To textCount - 1
        Dim textHeight As Double
        textHeight = CDbl(displayData.GetTextHeightAtIndex(i))
        If textHeight <= 0# Then textHeight = MIN_TEXT_HEIGHT_M

        Dim displayText As String
        displayText = CStr(displayData.GetTextAtIndex(i))

        Dim textWidth As Double
        textWidth = CDbl(Len(displayText)) * textHeight * TEXT_WIDTH_FACTOR
        If textWidth < textHeight Then textWidth = textHeight

        AddPoint bounds, originX - textWidth / 2#, _
            originY - textHeight / 2#
        AddPoint bounds, originX + textWidth / 2#, _
            originY + textHeight / 2#
    Next i

    AddDisplayTextGeometry = True
    Exit Function

Failed:
    AddDisplayTextGeometry = False
End Function

Private Function PlanWithScaleFallback( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef views As Collection, _
    ByRef bounds() As LayoutRect, _
    ByRef roles() As Long, _
    ByRef originalScales() As Double, _
    ByRef usable As LayoutRect, _
    ByRef protected() As ProtectedRect, _
    ByVal protectedCount As Long, _
    ByRef planned() As LayoutRect, _
    ByRef plannedScales() As Double, _
    ByRef acceptedScaleFactor As Double, _
    ByRef scaleChanged As Boolean) As Boolean

    On Error GoTo Failed

    Dim attempt As Long
    Dim factor As Double
    factor = 1#

    For attempt = 0 To MAX_SCALE_ATTEMPTS
        If attempt > 0 Then
            factor = factor * SCALE_REDUCTION_FACTOR
            If Not ScaleFactorIsPermitted(originalScales, factor) Then _
                Exit For

            If Not ApplyUniformScale( _
                views, originalScales, factor, plannedScales) Then
                RestoreViewScales swDrawModel, views, originalScales
                Exit Function
            End If

            Dim rebuildSucceeded As Boolean
            rebuildSucceeded = swDrawModel.ForceRebuild3(False)
            If rebuildSucceeded = False Then
                RestoreViewScales swDrawModel, views, originalScales
                Exit Function
            End If

            If Not RefreshViewEnvelopes(views, bounds) Then
                RestoreViewScales swDrawModel, views, originalScales
                Exit Function
            End If
        Else
            Dim scaleIndex As Long
            For scaleIndex = 1 To views.Count
                plannedScales(scaleIndex) = originalScales(scaleIndex)
            Next scaleIndex
        End If

        AppendReport "LAYOUT_SCALE_ATTEMPT|attempt=" & CStr(attempt + 1) & _
            "|factor=" & Format$(factor, "0.000000")

        Dim rejectedIndex As Long
        Dim placedCount As Long
        If TryPlanCurrentEnvelopes( _
            views, bounds, roles, usable, protected, protectedCount, _
            planned, rejectedIndex, placedCount) Then
            Dim scaleReadbackView As SldWorks.View
            For scaleIndex = 1 To views.Count
                Set scaleReadbackView = views(scaleIndex)
                plannedScales(scaleIndex) = _
                    CDbl(scaleReadbackView.ScaleDecimal)
            Next scaleIndex

            acceptedScaleFactor = factor
            scaleChanged = (Abs(factor - 1#) > SCALE_TOLERANCE)
            PlanWithScaleFallback = True
            Exit Function
        End If

        Dim rejectedView As SldWorks.View
        Set rejectedView = views(rejectedIndex)
        AppendReport "LAYOUT_REJECT|view=" & _
            CleanToken(SafeViewName(rejectedView)) & _
            "|attempt=" & CStr(attempt + 1) & _
            "|factor=" & Format$(factor, "0.000000") & _
            "|usable=" & FormatRect(usable) & _
            "|envelope=" & FormatRect(bounds(rejectedIndex)) & _
            "|protected=" & CStr(protectedCount) & _
            "|placed=" & CStr(placedCount)
    Next attempt

    If Abs(factor - 1#) > SCALE_TOLERANCE Then
        RestoreViewScales swDrawModel, views, originalScales
    End If
    Exit Function

Failed:
    On Error Resume Next
    RestoreViewScales swDrawModel, views, originalScales
    On Error GoTo 0
End Function

Private Function TryPlanCurrentEnvelopes( _
    ByRef views As Collection, _
    ByRef bounds() As LayoutRect, _
    ByRef roles() As Long, _
    ByRef usable As LayoutRect, _
    ByRef protected() As ProtectedRect, _
    ByVal protectedCount As Long, _
    ByRef planned() As LayoutRect, _
    ByRef rejectedIndex As Long, _
    ByRef rejectedPlacedCount As Long) As Boolean

    Dim order() As Long
    Dim placedOrder() As Long
    ReDim order(1 To views.Count)
    ReDim placedOrder(1 To views.Count)

    Dim i As Long
    For i = 1 To views.Count
        order(i) = i
    Next i

    SortPlacementOrder order, roles, bounds, views.Count

    Dim placedCount As Long
    For i = 1 To views.Count
        Dim viewIndex As Long
        viewIndex = order(i)

        If Not FindPlacement( _
            bounds(viewIndex), roles(viewIndex), usable, protected, _
            protectedCount, planned, placedOrder, placedCount, _
            planned(viewIndex)) Then
            rejectedIndex = viewIndex
            rejectedPlacedCount = placedCount
            Exit Function
        End If

        Dim plannedView As SldWorks.View
        Set plannedView = views(viewIndex)
        AppendReport "LAYOUT_PLAN|view=" & _
            CleanToken(SafeViewName(plannedView)) & _
            "|bounds=" & FormatRect(planned(viewIndex)) & _
            "|stabilityBuffer=" & _
            Format$(PLACEMENT_STABILITY_BUFFER_M, "0.000000")

        placedCount = placedCount + 1
        placedOrder(placedCount) = viewIndex
    Next i

    TryPlanCurrentEnvelopes = True
End Function

Private Function ScaleFactorIsPermitted( _
    ByRef originalScales() As Double, _
    ByVal factor As Double) As Boolean

    Dim i As Long
    For i = LBound(originalScales) To UBound(originalScales)
        Dim minimumScale As Double
        minimumScale = MIN_VIEW_SCALE_DECIMAL
        If originalScales(i) < minimumScale Then _
            minimumScale = originalScales(i)

        If originalScales(i) * factor < minimumScale - _
           SCALE_TOLERANCE Then Exit Function
    Next i

    ScaleFactorIsPermitted = True
End Function

Private Function ApplyUniformScale( _
    ByRef views As Collection, _
    ByRef originalScales() As Double, _
    ByVal factor As Double, _
    ByRef plannedScales() As Double) As Boolean

    On Error GoTo Failed

    Dim i As Long
    Dim swView As SldWorks.View
    For i = 1 To views.Count
        Set swView = views(i)
        plannedScales(i) = originalScales(i) * factor
        swView.ScaleDecimal = plannedScales(i)
        AppendReport "LAYOUT_SCALE|view=" & _
            CleanToken(SafeViewName(swView)) & _
            "|from=" & Format$(originalScales(i), "0.000000") & _
            "|to=" & Format$(plannedScales(i), "0.000000") & _
            "|factor=" & Format$(factor, "0.000000")
    Next i

    ApplyUniformScale = True
    Exit Function

Failed:
    AppendReport "LAYOUT_SCALE_ERROR|view=" & _
        CleanToken(SafeViewName(swView)) & _
        "|error=" & CStr(Err.Number) & ":" & CleanToken(Err.Description)
End Function

Private Function RefreshViewEnvelopes( _
    ByRef views As Collection, _
    ByRef bounds() As LayoutRect) As Boolean

    Dim i As Long
    Dim swView As SldWorks.View
    For i = 1 To views.Count
        Set swView = views(i)
        If Not BuildViewEnvelope(swView, bounds(i)) Then Exit Function
        AppendReport "LAYOUT_REMEASURE|view=" & _
            CleanToken(SafeViewName(swView)) & _
            "|bounds=" & FormatRect(bounds(i)) & _
            "|scale=" & Format$(CDbl(swView.ScaleDecimal), "0.000000")
    Next i

    RefreshViewEnvelopes = True
End Function

Private Sub RestoreViewScales( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef views As Collection, _
    ByRef originalScales() As Double)

    On Error Resume Next

    Dim i As Long
    Dim swView As SldWorks.View
    For i = 1 To views.Count
        Set swView = views(i)
        swView.ScaleDecimal = originalScales(i)
    Next i

    Dim rebuildSucceeded As Boolean
    rebuildSucceeded = swDrawModel.ForceRebuild3(False)
    AppendReport "LAYOUT_RESTORE|scope=Scales|rebuild=" & _
        CStr(rebuildSucceeded)
    On Error GoTo 0
End Sub

Private Sub RestoreLayoutState( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef views As Collection, _
    ByRef originalScales() As Double, _
    ByRef originalPositions() As Variant)

    On Error Resume Next

    Dim i As Long
    Dim swView As SldWorks.View
    For i = 1 To views.Count
        Set swView = views(i)
        swView.ScaleDecimal = originalScales(i)
        swView.Position = originalPositions(i)
    Next i

    Dim rebuildSucceeded As Boolean
    rebuildSucceeded = swDrawModel.ForceRebuild3(False)
    AppendReport "LAYOUT_RESTORE|scope=ScalesAndPositions|rebuild=" & _
        CStr(rebuildSucceeded)
    On Error GoTo 0
End Sub

Private Sub SortPlacementOrder( _
    ByRef order() As Long, _
    ByRef roles() As Long, _
    ByRef bounds() As LayoutRect, _
    ByVal count As Long)

    Dim i As Long
    Dim j As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            Dim shouldSwap As Boolean
            shouldSwap = _
                (RolePriority(roles(order(j))) < _
                 RolePriority(roles(order(i))))

            If Not shouldSwap And _
               RolePriority(roles(order(j))) = _
               RolePriority(roles(order(i))) Then
                shouldSwap = _
                    (RectWidth(bounds(order(j))) * _
                     RectHeight(bounds(order(j))) > _
                     RectWidth(bounds(order(i))) * _
                     RectHeight(bounds(order(i))))
            End If

            If shouldSwap Then
                Dim swapIndex As Long
                swapIndex = order(i)
                order(i) = order(j)
                order(j) = swapIndex
            End If
        Next j
    Next i
End Sub

Private Function RolePriority(ByVal role As Long) As Long
    Select Case role
        Case Module8_ViewClassifier.VIEW_ROLE_FRONT
            RolePriority = 10
        Case Module8_ViewClassifier.VIEW_ROLE_SECTION
            RolePriority = 20
        Case Module8_ViewClassifier.VIEW_ROLE_TOP, _
             Module8_ViewClassifier.VIEW_ROLE_BOTTOM, _
             Module8_ViewClassifier.VIEW_ROLE_RIGHT, _
             Module8_ViewClassifier.VIEW_ROLE_LEFT, _
             Module8_ViewClassifier.VIEW_ROLE_BACK, _
             Module8_ViewClassifier.VIEW_ROLE_PROJECTED, _
             Module8_ViewClassifier.VIEW_ROLE_DETAIL
            RolePriority = 30
        Case Module8_ViewClassifier.VIEW_ROLE_UNKNOWN
            RolePriority = 40
        Case Module8_ViewClassifier.VIEW_ROLE_PICTORIAL
            RolePriority = 50
        Case Else
            RolePriority = 60
    End Select
End Function

Private Function FindPlacement( _
    ByRef original As LayoutRect, _
    ByVal role As Long, _
    ByRef usable As LayoutRect, _
    ByRef protected() As ProtectedRect, _
    ByVal protectedCount As Long, _
    ByRef planned() As LayoutRect, _
    ByRef placedOrder() As Long, _
    ByVal placedCount As Long, _
    ByRef result As LayoutRect) As Boolean

    Dim width As Double
    Dim height As Double
    width = RectWidth(original)
    height = RectHeight(original)

    Dim planningUsable As LayoutRect
    SetRect planningUsable, _
        usable.LeftX + PLACEMENT_STABILITY_BUFFER_M, _
        usable.BottomY + PLACEMENT_STABILITY_BUFFER_M, _
        usable.RightX - PLACEMENT_STABILITY_BUFFER_M, _
        usable.TopY - PLACEMENT_STABILITY_BUFFER_M

    If Not RectIsValid(planningUsable) Then Exit Function
    If width > RectWidth(planningUsable) Or _
       height > RectHeight(planningUsable) Then _
        Exit Function

    Dim xCandidates() As Double
    Dim yCandidates() As Double
    Dim xCount As Long
    Dim yCount As Long

    AddCandidate xCandidates, xCount, planningUsable.LeftX
    AddCandidate xCandidates, xCount, planningUsable.RightX - width
    AddCandidate yCandidates, yCount, planningUsable.BottomY
    AddCandidate yCandidates, yCount, planningUsable.TopY - height

    Dim targetLeft As Double
    Dim targetBottom As Double
    GetRoleTarget role, planningUsable, width, height, _
        targetLeft, targetBottom
    AddCandidate xCandidates, xCount, targetLeft
    AddCandidate yCandidates, yCount, targetBottom

    Dim i As Long
    For i = 1 To protectedCount
        AddRectEdgeCandidates protected(i).Bounds, width, height, _
            PROTECTED_CLEARANCE_M + PLACEMENT_STABILITY_BUFFER_M, _
            xCandidates, xCount, _
            yCandidates, yCount
    Next i

    For i = 1 To placedCount
        AddRectEdgeCandidates planned(placedOrder(i)), width, height, _
            VIEW_CLEARANCE_M + PLACEMENT_STABILITY_BUFFER_M, _
            xCandidates, xCount, yCandidates, yCount
    Next i

    Dim bestScore As Double
    bestScore = 1E+30

    Dim xi As Long
    Dim yi As Long
    For xi = 1 To xCount
        For yi = 1 To yCount
            Dim candidate As LayoutRect
            SetRect candidate, xCandidates(xi), yCandidates(yi), _
                xCandidates(xi) + width, yCandidates(yi) + height

            If RectInside(candidate, planningUsable) Then
                If CandidateIsClear( _
                    candidate, protected, protectedCount, planned, _
                    placedOrder, placedCount) Then
                    Dim score As Double
                    score = (candidate.LeftX - targetLeft) ^ 2 + _
                        (candidate.BottomY - targetBottom) ^ 2
                    If score < bestScore Then
                        bestScore = score
                        CopyRect result, candidate
                        FindPlacement = True
                    End If
                End If
            End If
        Next yi
    Next xi
End Function

Private Sub GetRoleTarget( _
    ByVal role As Long, _
    ByRef usable As LayoutRect, _
    ByVal width As Double, _
    ByVal height As Double, _
    ByRef targetLeft As Double, _
    ByRef targetBottom As Double)

    targetLeft = usable.LeftX
    targetBottom = usable.BottomY

    Select Case role
        Case Module8_ViewClassifier.VIEW_ROLE_SECTION, _
             Module8_ViewClassifier.VIEW_ROLE_TOP
            targetBottom = usable.TopY - height
        Case Module8_ViewClassifier.VIEW_ROLE_PICTORIAL
            targetLeft = usable.RightX - width
            targetBottom = usable.TopY - height
        Case Module8_ViewClassifier.VIEW_ROLE_RIGHT, _
             Module8_ViewClassifier.VIEW_ROLE_LEFT, _
             Module8_ViewClassifier.VIEW_ROLE_BACK
            targetLeft = usable.RightX - width
        Case Module8_ViewClassifier.VIEW_ROLE_PROJECTED, _
             Module8_ViewClassifier.VIEW_ROLE_DETAIL
            targetBottom = usable.TopY - height
    End Select
End Sub

Private Sub AddRectEdgeCandidates( _
    ByRef obstacle As LayoutRect, _
    ByVal width As Double, _
    ByVal height As Double, _
    ByVal clearance As Double, _
    ByRef xCandidates() As Double, _
    ByRef xCount As Long, _
    ByRef yCandidates() As Double, _
    ByRef yCount As Long)

    AddCandidate xCandidates, xCount, obstacle.LeftX - clearance - width
    AddCandidate xCandidates, xCount, obstacle.RightX + clearance
    AddCandidate yCandidates, yCount, obstacle.BottomY - clearance - height
    AddCandidate yCandidates, yCount, obstacle.TopY + clearance
End Sub

Private Sub AddCandidate( _
    ByRef values() As Double, _
    ByRef count As Long, _
    ByVal value As Double)

    Dim i As Long
    For i = 1 To count
        If Abs(values(i) - value) <= POSITION_TOLERANCE_M Then Exit Sub
    Next i

    count = count + 1
    ReDim Preserve values(1 To count)
    values(count) = value
End Sub

Private Function CandidateIsClear( _
    ByRef candidate As LayoutRect, _
    ByRef protected() As ProtectedRect, _
    ByVal protectedCount As Long, _
    ByRef planned() As LayoutRect, _
    ByRef placedOrder() As Long, _
    ByVal placedCount As Long) As Boolean

    Dim i As Long
    For i = 1 To protectedCount
        If RectsConflict( _
            candidate, protected(i).Bounds, PROTECTED_CLEARANCE_M) Then
            Exit Function
        End If
    Next i

    For i = 1 To placedCount
        If RectsConflict( _
            candidate, planned(placedOrder(i)), VIEW_CLEARANCE_M) Then
            Exit Function
        End If
    Next i

    CandidateIsClear = True
End Function

Private Function MoveViewByEnvelopeDelta( _
    ByRef swView As SldWorks.View, _
    ByRef currentBounds As LayoutRect, _
    ByRef plannedBounds As LayoutRect) As Boolean

    On Error GoTo Failed

    Dim position As Variant
    position = swView.Position
    If Not IsArray(position) Then Exit Function
    If UBound(position) - LBound(position) + 1 < 2 Then Exit Function

    Dim target(0 To 1) As Double
    target(0) = CDbl(position(LBound(position))) + _
        plannedBounds.LeftX - currentBounds.LeftX
    target(1) = CDbl(position(LBound(position) + 1)) + _
        plannedBounds.BottomY - currentBounds.BottomY

    swView.Position = target
    AppendReport "LAYOUT_VIEW|name=" & CleanToken(SafeViewName(swView)) & _
        "|operation=Move|dx=" & Format$( _
            plannedBounds.LeftX - currentBounds.LeftX, "0.000000") & _
        "|dy=" & Format$( _
            plannedBounds.BottomY - currentBounds.BottomY, "0.000000")
    MoveViewByEnvelopeDelta = True
    Exit Function

Failed:
    MoveViewByEnvelopeDelta = False
End Function

Private Function ValidateFinalLayout( _
    ByRef views As Collection, _
    ByRef usable As LayoutRect, _
    ByRef protected() As ProtectedRect, _
    ByVal protectedCount As Long, _
    ByRef expectedScales() As Double) As Boolean

    On Error GoTo Failed

    Dim finalBounds() As LayoutRect
    ReDim finalBounds(1 To views.Count)

    Dim i As Long
    Dim j As Long
    Dim swView As SldWorks.View
    For i = 1 To views.Count
        Set swView = views(i)
        If Abs(CDbl(swView.ScaleDecimal) - expectedScales(i)) > _
           SCALE_TOLERANCE Then
            RecordFailure "ScaleReadbackMismatch:" & SafeViewName(swView) & _
                ":expected=" & Format$(expectedScales(i), "0.000000") & _
                ":actual=" & Format$(CDbl(swView.ScaleDecimal), "0.000000")
            Exit Function
        End If

        If Not BuildViewEnvelope(swView, finalBounds(i)) Then
            RecordFailure "FinalEnvelopeUnavailable:" & SafeViewName(swView)
            Exit Function
        End If

        AppendReport "LAYOUT_FINAL|view=" & _
            CleanToken(SafeViewName(swView)) & _
            "|bounds=" & FormatRect(finalBounds(i)) & _
            "|scale=" & Format$(CDbl(swView.ScaleDecimal), "0.000000")

        If Not RectInside(finalBounds(i), usable) Then
            RecordFailure "ContentOutsideUsableSheet:" & SafeViewName(swView)
            Exit Function
        End If

        For j = 1 To protectedCount
            If RectsConflict( _
                finalBounds(i), protected(j).Bounds, _
                PROTECTED_CLEARANCE_M) Then
                RecordFailure "ProtectedRegionCollision:" & _
                    SafeViewName(swView) & ":" & protected(j).Name
                Exit Function
            End If
        Next j
    Next i

    For i = 1 To views.Count - 1
        For j = i + 1 To views.Count
            If RectsConflict( _
                finalBounds(i), finalBounds(j), VIEW_CLEARANCE_M) Then
                Dim firstView As SldWorks.View
                Dim secondView As SldWorks.View
                Set firstView = views(i)
                Set secondView = views(j)
                RecordFailure "ViewEnvelopeCollision:" & _
                    SafeViewName(firstView) & ":" & _
                    SafeViewName(secondView) & _
                    ":first=" & FormatRect(finalBounds(i)) & _
                    ":second=" & FormatRect(finalBounds(j)) & _
                    ":requiredClearance=" & _
                    Format$(VIEW_CLEARANCE_M, "0.000000")
                Exit Function
            End If
        Next j
    Next i

    ValidateFinalLayout = True
    Exit Function

Failed:
    RecordFailure "FinalValidationError:" & CStr(Err.Number)
End Function

Private Sub AddProtectedRect( _
    ByRef protected() As ProtectedRect, _
    ByRef count As Long, _
    ByVal name As String, _
    ByRef bounds As LayoutRect)

    If count >= MAX_PROTECTED_RECTS Then
        Err.Raise vbObjectError + 1034, , "ProtectedRectangleCapacityExceeded"
    End If

    count = count + 1
    protected(count).Name = name
    CopyRect protected(count).Bounds, bounds
End Sub

Private Sub CopyRect( _
    ByRef target As LayoutRect, _
    ByRef source As LayoutRect)

    target.LeftX = source.LeftX
    target.BottomY = source.BottomY
    target.RightX = source.RightX
    target.TopY = source.TopY
    target.Seeded = source.Seeded
End Sub

Private Sub SetRect( _
    ByRef bounds As LayoutRect, _
    ByVal leftX As Double, _
    ByVal bottomY As Double, _
    ByVal rightX As Double, _
    ByVal topY As Double)

    bounds.LeftX = leftX
    bounds.BottomY = bottomY
    bounds.RightX = rightX
    bounds.TopY = topY
    bounds.Seeded = True
End Sub

Private Sub AddPoint( _
    ByRef bounds As LayoutRect, _
    ByVal x As Double, _
    ByVal y As Double)

    If Not bounds.Seeded Then
        SetRect bounds, x, y, x, y
        Exit Sub
    End If

    If x < bounds.LeftX Then bounds.LeftX = x
    If x > bounds.RightX Then bounds.RightX = x
    If y < bounds.BottomY Then bounds.BottomY = y
    If y > bounds.TopY Then bounds.TopY = y
End Sub

Private Sub AddRect( _
    ByRef target As LayoutRect, _
    ByRef source As LayoutRect)

    If Not source.Seeded Then Exit Sub
    AddPoint target, source.LeftX, source.BottomY
    AddPoint target, source.RightX, source.TopY
End Sub

Private Function RectIsValid(ByRef bounds As LayoutRect) As Boolean
    RectIsValid = bounds.Seeded And _
        bounds.RightX > bounds.LeftX And _
        bounds.TopY > bounds.BottomY
End Function

Private Function RectWidth(ByRef bounds As LayoutRect) As Double
    RectWidth = bounds.RightX - bounds.LeftX
End Function

Private Function RectHeight(ByRef bounds As LayoutRect) As Double
    RectHeight = bounds.TopY - bounds.BottomY
End Function

Private Function RectInside( _
    ByRef inner As LayoutRect, _
    ByRef outer As LayoutRect) As Boolean

    RectInside = inner.LeftX >= outer.LeftX - POSITION_TOLERANCE_M And _
        inner.RightX <= outer.RightX + POSITION_TOLERANCE_M And _
        inner.BottomY >= outer.BottomY - POSITION_TOLERANCE_M And _
        inner.TopY <= outer.TopY + POSITION_TOLERANCE_M
End Function

Private Function PointInsideRect( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByRef bounds As LayoutRect) As Boolean

    PointInsideRect = bounds.Seeded And _
        x >= bounds.LeftX - POSITION_TOLERANCE_M And _
        x <= bounds.RightX + POSITION_TOLERANCE_M And _
        y >= bounds.BottomY - POSITION_TOLERANCE_M And _
        y <= bounds.TopY + POSITION_TOLERANCE_M
End Function

Private Function RectsConflict( _
    ByRef first As LayoutRect, _
    ByRef second As LayoutRect, _
    ByVal clearance As Double) As Boolean

    RectsConflict = Not ( _
        first.RightX + clearance <= _
            second.LeftX + POSITION_TOLERANCE_M Or _
        second.RightX + clearance <= _
            first.LeftX + POSITION_TOLERANCE_M Or _
        first.TopY + clearance <= _
            second.BottomY + POSITION_TOLERANCE_M Or _
        second.TopY + clearance <= _
            first.BottomY + POSITION_TOLERANCE_M)
End Function

Private Function SafeViewName(ByRef swView As SldWorks.View) As String
    On Error Resume Next
    SafeViewName = swView.Name
    If Len(SafeViewName) = 0 Then SafeViewName = "(unnamed)"
    On Error GoTo 0
End Function

Private Function CleanToken(ByVal value As String) As String
    Dim result As String
    result = Replace(value, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, "|", "/")
    CleanToken = Trim$(result)
End Function

Private Function FormatRect(ByRef bounds As LayoutRect) As String
    If Not bounds.Seeded Then
        FormatRect = "unseeded"
    Else
        FormatRect = Format$(bounds.LeftX, "0.000000") & "," & _
            Format$(bounds.BottomY, "0.000000") & "," & _
            Format$(bounds.RightX, "0.000000") & "," & _
            Format$(bounds.TopY, "0.000000")
    End If
End Function

Private Function MinDouble(ByVal first As Double, ByVal second As Double) As Double
    If first < second Then
        MinDouble = first
    Else
        MinDouble = second
    End If
End Function

Private Function MaxDouble(ByVal first As Double, ByVal second As Double) As Double
    If first > second Then
        MaxDouble = first
    Else
        MaxDouble = second
    End If
End Function

' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub

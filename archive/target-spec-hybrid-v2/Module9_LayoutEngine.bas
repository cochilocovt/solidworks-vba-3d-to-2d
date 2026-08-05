Option Explicit

Private Const swDrawingSectionView As Long = 2
Private Const swDrawingDetailView As Long = 3
Private Const VIEW_GAP_M As Double = 0.012
Private Const LAYOUT_COMPARISON_TOLERANCE_M As Double = 0.000001

' R23-812. Two view scales are the same ratio when they agree to this.
' Adjacent approved scales differ by a factor, never by a millionth, so
' this absorbs double representation and discriminates nothing real.
Private Const VIEW_SCALE_TOLERANCE As Double = 0.000001
Private Const TITLE_BLOCK_RESERVED_TOKEN As String = _
    "titleBlockIsReservedRectangle=True"

' isFinalStructuralPass distinguishes the two legitimate placements.
' The rough pass runs at pipeline step 3, before the section and isometric
' views exist, purely so the views resolved in step 4 do not overlap. It
' cannot satisfy a fixture reference view set and therefore never writes the
' LAYOUT verdict: MarkStageFailed is permanent, so a rough failure would
' pre-empt the real pass. The final pass runs once every required view exists
' and is the only one that proves or fails the stage.
Public Function ArrangeViewsInMeasuredGrid( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal layoutPass As String, _
    ByVal isFinalStructuralPass As Boolean, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    ' Registered on both passes. If the final pass never runs the stage seals
    ' as NOT_STARTED rather than disappearing from the required set.
    evidence.RequireStage "LAYOUT"

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        evidence.AddFailure "Layout: current sheet is Nothing."
        evidence.MarkStageFailed "LAYOUT", "current sheet is Nothing"
        Exit Function
    End If

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim paperSize As Long
    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)

    If sheetWidth <= 0# Or sheetHeight <= 0# Then
        evidence.AddFailure "Layout: sheet size is invalid."
        evidence.MarkStageFailed "LAYOUT", "invalid sheet size"
        Exit Function
    End If

    Dim usedDiagnosticBounds As Boolean
    usedDiagnosticBounds = False

    If Not evidence.LayoutBoundariesProven Then
        If Not Module1_Main.DIAGNOSTIC_DRAWING_MODE Then
            evidence.AddFailure _
                "Layout: measured border/title-block bounds are not proven."
            evidence.MarkStageFailed "LAYOUT", _
                "controlled sheet boundaries are not proved"
            Exit Function
        End If

        usedDiagnosticBounds = True
        evidence.UsableLeft = Module8_RuntimeSupport.LAYOUT_MARGIN_M
        evidence.UsableRight = sheetWidth - _
            Module8_RuntimeSupport.LAYOUT_MARGIN_M
        evidence.UsableBottom = Module8_RuntimeSupport.LAYOUT_MARGIN_M
        evidence.UsableTop = sheetHeight - _
            Module8_RuntimeSupport.LAYOUT_MARGIN_M
        evidence.ContentBorderLeft = 0#
        evidence.ContentBorderBottom = 0#
        evidence.ContentBorderRight = sheetWidth
        evidence.ContentBorderTop = sheetHeight
        evidence.TitleBlockLeft = sheetWidth * 0.68
        evidence.TitleBlockRight = sheetWidth
        evidence.TitleBlockBottom = 0#
        evidence.TitleBlockTop = sheetHeight * 0.3
        evidence.AddWarning "LAYOUT_BOUNDS|source=DiagnosticSheetReserve|" & _
            "measured=False|acceptance=False|" & _
            TITLE_BLOCK_RESERVED_TOKEN
    End If

    Dim views As Collection
    Set views = GetModelViews(swDraw)

    If views.Count = 0 Then
        evidence.AddFailure "Layout: drawing contains no model views."
        evidence.MarkStageFailed "LAYOUT", "no model views"
        Exit Function
    End If

    If usedDiagnosticBounds Then
        ArrangeViewsInMeasuredGrid = ArrangeDiagnosticRows( _
            swDrawModel, swDraw, views, evidence.UsableLeft, _
                evidence.UsableBottom, _
            evidence.UsableRight, evidence.UsableTop, _
            isFinalStructuralPass, evidence)

        If Not isFinalStructuralPass Then
            evidence.AddInfo "LAYOUT_ROUGH|pass=" & layoutPass & _
                "|bounds=Diagnostic|placed=" & _
                CStr(ArrangeViewsInMeasuredGrid) & _
                "|verdictDeferredTo=FinalStructuralPass"
            Exit Function
        End If

        If ArrangeViewsInMeasuredGrid Then
            evidence.MarkStageFailed "LAYOUT", _
                "diagnostic zone-aware bounds were used; visual output only"
        Else
            evidence.MarkStageFailed "LAYOUT", _
                "diagnostic zone-aware layout did not fit"
        End If
        Exit Function
    End If

    Dim leftBoundary As Double
    Dim rightBoundary As Double
    Dim bottomBoundary As Double
    Dim topBoundary As Double

    leftBoundary = evidence.UsableLeft
    rightBoundary = evidence.UsableRight
    bottomBoundary = evidence.TitleBlockBottom + _
        Module8_RuntimeSupport.LAYOUT_MARGIN_M
    topBoundary = evidence.UsableTop

    ArrangeViewsInMeasuredGrid = ArrangeZoneAwareViews( _
        swDrawModel, swDraw, views, leftBoundary, bottomBoundary, _
        rightBoundary, topBoundary, _
        "MeasuredZoneAware." & layoutPass, isFinalStructuralPass, evidence)

    If Not isFinalStructuralPass Then
        evidence.AddInfo "LAYOUT_ROUGH|pass=" & layoutPass & _
            "|bounds=Measured|placed=" & _
            CStr(ArrangeViewsInMeasuredGrid) & _
            "|verdictDeferredTo=FinalStructuralPass"
        Exit Function
    End If

    If ArrangeViewsInMeasuredGrid And Not usedDiagnosticBounds Then
        evidence.MarkStageProved "LAYOUT", _
            "all view outlines fit measured title-block zones without " & _
                "independent scale changes"
    ElseIf ArrangeViewsInMeasuredGrid Then
        evidence.MarkStageFailed "LAYOUT", _
            "diagnostic fallback bounds were used; visual output only"
    Else
        evidence.MarkStageFailed "LAYOUT", _
            "boundary, collision, or scale validation failed"
    End If
    Exit Function

Failed:
    evidence.AddFailure "Layout error " & CStr(Err.Number) & _
        ": " & Err.Description
    evidence.MarkStageFailed "LAYOUT", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function ArrangeDiagnosticRows( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef views As Collection, _
    ByVal leftBoundary As Double, _
    ByVal bottomBoundary As Double, _
    ByVal rightBoundary As Double, _
    ByVal topBoundary As Double, _
    ByVal isFinalStructuralPass As Boolean, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim viewCount As Long
    viewCount = views.Count

    ArrangeDiagnosticRows = ArrangeZoneAwareViews( _
        swDrawModel, swDraw, views, leftBoundary, bottomBoundary, _
        rightBoundary, topBoundary, "DiagnosticZoneAware", _
        isFinalStructuralPass, evidence)

    If ArrangeDiagnosticRows Then
        evidence.AddWarning "LAYOUT_RESULT|mode=DiagnosticZoneAware|" & _
            "acceptance=False|viewCount=" & CStr(viewCount)
    End If
    Exit Function

Failed:
    evidence.AddFailure "Diagnostic zone-aware layout error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function ArrangeZoneAwareViews( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef views As Collection, _
    ByVal leftBoundary As Double, _
    ByVal bottomBoundary As Double, _
    ByVal rightBoundary As Double, _
    ByVal topBoundary As Double, _
    ByVal layoutMode As String, _
    ByVal isFinalStructuralPass As Boolean, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If views.Count < 1 Then Exit Function
    If evidence.TitleBlockRight <= evidence.TitleBlockLeft Or _
       evidence.TitleBlockTop <= evidence.TitleBlockBottom Then
        evidence.AddFailure "Layout: title-block rectangle is invalid."
        Exit Function
    End If

    ' The fixture reference-zone arrangement asserts the complete approved
    ' view set (primary, side, J-J section, isometric). Only the final
    ' structural pass sees that set; running it on the rough pass produced the
    ' unconditional "requires exactly one primary, side, J-J section, and
    ' isometric view" failure in the 2026-08-04 18:45 run, which left the
    ' drawing with zero view moves. The rough pass uses the generic
    ' zone-aware placement below, which accepts any view count.
    If isFinalStructuralPass And _
       Module1_Main.GetFixtureKey(evidence.PartPath) = "P-0251-14A-001" Then

        ArrangeZoneAwareViews = ArrangeP0251ReferenceZones( _
            swDrawModel, swDraw, views, leftBoundary, bottomBoundary, _
            rightBoundary, topBoundary, layoutMode, evidence)
        Exit Function
    End If

    ' A title block occupies a rectangle, not the entire lower sheet.  Its
    ' bottom edge is the best available lower-border proof in this module;
    ' preserve an additional layout margin before using the lower-left zone.
    Dim zoneBottom As Double
    zoneBottom = evidence.TitleBlockBottom + _
        Module8_RuntimeSupport.LAYOUT_MARGIN_M
    If zoneBottom < bottomBoundary Then zoneBottom = bottomBoundary

    Dim leftZoneRight As Double
    leftZoneRight = evidence.TitleBlockLeft - VIEW_GAP_M
    If leftZoneRight <= leftBoundary Or topBoundary <= zoneBottom Then
        evidence.AddFailure _
            "Layout: title-block rectangle leaves no safe left zone."
        Exit Function
    End If

    Dim primaryView As SldWorks.View
    Set primaryView = views(1)

    Dim primaryOutline As Variant
    primaryOutline = primaryView.GetOutline
    Dim primaryWidth As Double
    Dim primaryHeight As Double
    primaryWidth = CDbl(primaryOutline(2)) - CDbl(primaryOutline(0))
    primaryHeight = CDbl(primaryOutline(3)) - CDbl(primaryOutline(1))

    If primaryWidth <= 0# Or primaryHeight <= 0# Then
        evidence.AddFailure "Layout: primary view has an invalid outline."
        Exit Function
    End If
    If primaryWidth > leftZoneRight - leftBoundary Or _
       primaryHeight > topBoundary - zoneBottom Then
        evidence.AddFailure _
            "Layout: primary view does not fit the safe left zone " & _
            "without an unapproved scale change."
        Exit Function
    End If

    Dim i As Long
    Dim remainingWidth As Double
    Dim remainingHeight As Double
    Dim remainingTotalWidth As Double
    remainingWidth = rightBoundary - (leftBoundary + primaryWidth + VIEW_GAP_M)

    For i = 2 To views.Count
        Dim candidateView As SldWorks.View
        Set candidateView = views(i)

        Dim candidateOutline As Variant
        candidateOutline = candidateView.GetOutline
        Dim candidateWidth As Double
        Dim candidateHeight As Double
        candidateWidth = CDbl(candidateOutline(2)) - CDbl(candidateOutline(0))
        candidateHeight = CDbl(candidateOutline(3)) - CDbl(candidateOutline(1))

        If candidateWidth <= 0# Or candidateHeight <= 0# Then
            evidence.AddFailure "Layout: invalid outline for '" & _
                Module8_RuntimeSupport.GetViewName(candidateView) & "'."
            Exit Function
        End If

        remainingTotalWidth = remainingTotalWidth + candidateWidth
        If candidateHeight > remainingHeight Then remainingHeight = _
            candidateHeight
    Next i

    If views.Count > 2 Then
        remainingTotalWidth = remainingTotalWidth + VIEW_GAP_M * _
            CDbl(views.Count - 2)
    End If

    If views.Count > 1 Then
        If remainingWidth <= 0# Or _
           remainingTotalWidth > remainingWidth Or _
           remainingHeight > topBoundary - evidence.TitleBlockTop - _
               VIEW_GAP_M Then

            evidence.AddFailure _
                "Layout: remaining views do not fit the safe " & _
                "band above the title-block rectangle without a scale change."
            Exit Function
        End If
    End If

    Dim primaryPosition(0 To 1) As Double
    primaryPosition(0) = leftBoundary + (leftZoneRight - leftBoundary) / 2#
    primaryPosition(1) = zoneBottom + (topBoundary - zoneBottom) / 2#
    evidence.RecordSolidWorksMutation "IView.Position:" & _
        Module8_RuntimeSupport.GetViewName(primaryView)
    primaryView.Position = primaryPosition
    evidence.LayoutMoves = evidence.LayoutMoves + 1

    If views.Count > 1 Then
        Dim remainingCursor As Double
        remainingCursor = leftBoundary + primaryWidth + VIEW_GAP_M + _
            (remainingWidth - remainingTotalWidth) / 2#

        For i = 2 To views.Count
            Set candidateView = views(i)
            candidateOutline = candidateView.GetOutline
            candidateWidth = CDbl(candidateOutline(2)) - _
                CDbl(candidateOutline(0))
            candidateHeight = CDbl(candidateOutline(3)) - _
                CDbl(candidateOutline(1))

            Dim candidatePosition(0 To 1) As Double
            candidatePosition(0) = remainingCursor + candidateWidth / 2#
            candidatePosition(1) = topBoundary - remainingHeight / 2#
            evidence.RecordSolidWorksMutation "IView.Position:" & _
                Module8_RuntimeSupport.GetViewName(candidateView)
            candidateView.Position = candidatePosition
            evidence.LayoutMoves = evidence.LayoutMoves + 1
            remainingCursor = remainingCursor + candidateWidth + VIEW_GAP_M
        Next i
    End If

    If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, layoutMode, evidence) Then Exit Function

    ArrangeZoneAwareViews = ValidateLayout( _
        swDraw, views, leftBoundary, zoneBottom, rightBoundary, topBoundary, _
        layoutMode, evidence)

    If ArrangeZoneAwareViews Then
        evidence.AddInfo "LAYOUT_RESULT|mode=" & layoutMode & _
            "|titleBlockReservedAsRectangle=True|lowerLeftZoneUsed=True"
    End If
    Exit Function

Failed:
    evidence.AddFailure "Zone-aware layout error " & CStr(Err.Number) & _
        ": " & Err.Description
End Function

Private Function ArrangeP0251ReferenceZones( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef views As Collection, _
    ByVal leftBoundary As Double, _
    ByVal bottomBoundary As Double, _
    ByVal rightBoundary As Double, _
    ByVal topBoundary As Double, _
    ByVal layoutMode As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim primaryView As SldWorks.View
    Dim sideView As SldWorks.View
    Dim sectionView As SldWorks.View
    Dim isometricView As SldWorks.View

    Set primaryView = views(1)

    Dim i As Long
    For i = 2 To views.Count
        Dim candidateView As SldWorks.View
        Set candidateView = views(i)

        If IsIsometricView(candidateView) Then
            If Not isometricView Is Nothing Then
                evidence.AddFailure _
                    "P-0251 layout found more than one isometric view."
                Exit Function
            End If
            Set isometricView = candidateView
        ElseIf candidateView.Type = swDrawingSectionView Then
            If Not sectionView Is Nothing Then
                evidence.AddFailure _
                    "P-0251 layout found more than one section view."
                Exit Function
            End If
            Set sectionView = candidateView
        Else
            If Not sideView Is Nothing Then
                evidence.AddFailure _
                    "P-0251 layout found an unexpected extra " & _
                    "non-isometric model view."
                Exit Function
            End If
            Set sideView = candidateView
        End If
    Next i

    If primaryView Is Nothing Or sideView Is Nothing Or _
       sectionView Is Nothing Or isometricView Is Nothing Then

        evidence.AddFailure "P-0251 reference layout requires exactly one " & _
            "primary, side, J-J section, and isometric view."
        Exit Function
    End If

    Dim zoneBottom As Double
    zoneBottom = evidence.TitleBlockBottom + _
        Module8_RuntimeSupport.LAYOUT_MARGIN_M
    If zoneBottom < bottomBoundary Then zoneBottom = bottomBoundary

    Dim leftZoneRight As Double
    leftZoneRight = evidence.TitleBlockLeft - VIEW_GAP_M
    If leftZoneRight <= leftBoundary Or topBoundary <= zoneBottom Then
        evidence.AddFailure "P-0251 layout has no safe left view zone."
        Exit Function
    End If

    Dim primaryOutline As Variant
    Dim sideOutline As Variant
    Dim sectionOutline As Variant
    Dim isoOutline As Variant
    primaryOutline = primaryView.GetOutline
    sideOutline = sideView.GetOutline
    sectionOutline = sectionView.GetOutline
    isoOutline = isometricView.GetOutline

    Dim primaryWidth As Double
    Dim primaryHeight As Double
    Dim sideWidth As Double
    Dim sideHeight As Double
    Dim sectionWidth As Double
    Dim sectionHeight As Double
    Dim isoWidth As Double
    Dim isoHeight As Double
    primaryWidth = OutlineWidth(primaryOutline)
    primaryHeight = OutlineHeight(primaryOutline)
    sideWidth = OutlineWidth(sideOutline)
    sideHeight = OutlineHeight(sideOutline)
    sectionWidth = OutlineWidth(sectionOutline)
    sectionHeight = OutlineHeight(sectionOutline)
    isoWidth = OutlineWidth(isoOutline)
    isoHeight = OutlineHeight(isoOutline)

    If primaryWidth <= 0# Or primaryHeight <= 0# Or _
       sideWidth <= 0# Or sideHeight <= 0# Or _
       sectionWidth <= 0# Or sectionHeight <= 0# Or _
       isoWidth <= 0# Or isoHeight <= 0# Then

        evidence.AddFailure "P-0251 layout found an invalid view outline."
        Exit Function
    End If

    Dim rowWidth As Double
    Dim rowHeight As Double
    rowWidth = primaryWidth + sideWidth + sectionWidth + (2# * VIEW_GAP_M)
    rowHeight = primaryHeight
    If sideHeight > rowHeight Then rowHeight = sideHeight
    If sectionHeight > rowHeight Then rowHeight = sectionHeight

    Dim leftZoneWidth As Double
    Dim leftZoneHeight As Double
    leftZoneWidth = leftZoneRight - leftBoundary
    leftZoneHeight = topBoundary - zoneBottom

    If rowWidth > leftZoneWidth Or rowHeight > leftZoneHeight Then
        evidence.AddFailure _
            "P-0251 primary/side/section row does not fit the " & _
            "title-clear left zone without an unapproved scale change."
        Exit Function
    End If

    Dim rightZoneLeft As Double
    Dim rightZoneBottom As Double
    rightZoneLeft = evidence.TitleBlockLeft + VIEW_GAP_M
    rightZoneBottom = evidence.TitleBlockTop + VIEW_GAP_M

    If rightZoneLeft >= rightBoundary Or rightZoneBottom >= topBoundary Or _
       isoWidth > rightBoundary - rightZoneLeft Or _
       isoHeight > topBoundary - rightZoneBottom Then

        evidence.AddFailure "P-0251 isometric view does not fit the " & _
            "title-clear upper-right zone without an unapproved scale change."
        Exit Function
    End If

    Dim rowCursor As Double
    Dim rowCenterY As Double
    rowCursor = leftBoundary + (leftZoneWidth - rowWidth) / 2#
    ' Bias the P-0251 source row upward.  The J-J cut line belongs to the
    ' primary view and its lower marker otherwise intrudes into the controlled
    ' part-identification band even when the model-view outline itself fits.
    rowCenterY = topBoundary - rowHeight / 2# - VIEW_GAP_M / 2#

    If rowCenterY - rowHeight / 2# < zoneBottom Then
        evidence.AddFailure _
            "P-0251 upward-biased row cannot clear the lower " & _
            "controlled content while remaining inside the measured border."
        Exit Function
    End If

    MoveViewOutlineCenter primaryView, _
        rowCursor + primaryWidth / 2#, rowCenterY, evidence
    rowCursor = rowCursor + primaryWidth + VIEW_GAP_M
    MoveViewOutlineCenter sideView, _
        rowCursor + sideWidth / 2#, rowCenterY, evidence
    rowCursor = rowCursor + sideWidth + VIEW_GAP_M
    MoveViewOutlineCenter sectionView, _
        rowCursor + sectionWidth / 2#, rowCenterY, evidence

    MoveViewOutlineCenter isometricView, _
        rightZoneLeft + (rightBoundary - rightZoneLeft) / 2#, _
        rightZoneBottom + (topBoundary - rightZoneBottom) / 2#, evidence

    If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, layoutMode & ".P0251ReferenceZones", evidence) Then

        Exit Function
    End If

    ArrangeP0251ReferenceZones = ValidateLayout( _
        swDraw, views, leftBoundary, zoneBottom, rightBoundary, topBoundary, _
        layoutMode, evidence)

    If ArrangeP0251ReferenceZones Then
        evidence.AddInfo "LAYOUT_RESULT|mode=" & layoutMode & _
            "|profile=P0251ReferenceZones" & _
            "|titleBlockReservedAsRectangle=True|lowerLeftZoneUsed=True"
    End If
    Exit Function

Failed:
    evidence.AddFailure "P-0251 reference-zone layout error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub MoveViewOutlineCenter( _
    ByRef swView As SldWorks.View, _
    ByVal targetCenterX As Double, _
    ByVal targetCenterY As Double, _
    ByRef evidence As CRunEvidence)

    Dim outline As Variant
    outline = swView.GetOutline

    Dim currentCenterX As Double
    Dim currentCenterY As Double
    currentCenterX = (CDbl(outline(0)) + CDbl(outline(2))) / 2#
    currentCenterY = (CDbl(outline(1)) + CDbl(outline(3))) / 2#

    Dim currentPosition As Variant
    currentPosition = swView.Position

    Dim targetPosition(0 To 1) As Double
    targetPosition(0) = CDbl(currentPosition(0)) + _
        (targetCenterX - currentCenterX)
    targetPosition(1) = CDbl(currentPosition(1)) + _
        (targetCenterY - currentCenterY)

    evidence.RecordSolidWorksMutation "IView.Position:" & _
        Module8_RuntimeSupport.GetViewName(swView)
    swView.Position = targetPosition
    evidence.LayoutMoves = evidence.LayoutMoves + 1

    Dim readbackOutline As Variant
    readbackOutline = swView.GetOutline
    evidence.AddInfo "LAYOUT_MOVE|view=" & _
        Module8_RuntimeSupport.GetViewName(swView) & _
        "|requestedCenter=" & Format$(targetCenterX, "0.000000") & "," & _
        Format$(targetCenterY, "0.000000") & _
        "|readbackOutline=" & Format$(CDbl(readbackOutline(0)), "0.000000") & _
        "," & Format$(CDbl(readbackOutline(1)), "0.000000") & "," & _
        Format$(CDbl(readbackOutline(2)), "0.000000") & "," & _
        Format$(CDbl(readbackOutline(3)), "0.000000")
End Sub

Private Function OutlineWidth(ByVal outline As Variant) As Double
    If Not IsArray(outline) Then Exit Function
    If UBound(outline) - LBound(outline) + 1 < 4 Then Exit Function
    OutlineWidth = CDbl(outline(LBound(outline) + 2)) - _
        CDbl(outline(LBound(outline)))
End Function

Private Function OutlineHeight(ByVal outline As Variant) As Double
    If Not IsArray(outline) Then Exit Function
    If UBound(outline) - LBound(outline) + 1 < 4 Then Exit Function
    OutlineHeight = CDbl(outline(LBound(outline) + 3)) - _
        CDbl(outline(LBound(outline) + 1))
End Function

Private Function GetModelViews( _
    ByRef swDraw As SldWorks.DrawingDoc) As Collection

    Dim results As New Collection
    Set GetModelViews = results

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        results.Add swView
        Set swView = swView.GetNextView
    Loop
End Function

Private Function FitViewToCell( _
    ByRef swView As SldWorks.View, _
    ByVal cellWidth As Double, _
    ByVal cellHeight As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim outline As Variant
    outline = swView.GetOutline

    Dim viewWidth As Double
    Dim viewHeight As Double
    viewWidth = CDbl(outline(2)) - CDbl(outline(0))
    viewHeight = CDbl(outline(3)) - CDbl(outline(1))

    If viewWidth <= 0# Or viewHeight <= 0# Then
        evidence.AddFailure "Layout: invalid outline for '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    Dim allowedWidth As Double
    Dim allowedHeight As Double
    allowedWidth = cellWidth - (2# * Module8_RuntimeSupport.LAYOUT_MARGIN_M)
    allowedHeight = cellHeight - (2# * Module8_RuntimeSupport.LAYOUT_MARGIN_M)

    If allowedWidth <= 0# Or allowedHeight <= 0# Then
        evidence.AddFailure _
            "Layout: usable cell is smaller than the required margin for '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    Dim factor As Double
    factor = 1#

    If viewWidth > allowedWidth Then factor = allowedWidth / viewWidth
    If viewHeight * factor > allowedHeight Then factor = allowedHeight / _
        viewHeight

    If factor < 0.999 Then
        evidence.AddFailure _
            "Layout would require an unapproved independent scale for '" & _
            Module8_RuntimeSupport.GetViewName(swView) & _
            "'; select an approved sheet scale or sheet size instead."
        Exit Function
    End If

    FitViewToCell = True
    Exit Function

Failed:
    evidence.AddFailure "Layout fit error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
    FitViewToCell = False
End Function

' R23-812. True only when the view is demonstrably DRAWN at the proved
' sheet ratio, whatever its UseSheetScale flag says.
'
' Fails closed in both directions that matter. An unproved sheet scale is
' not a match, because then there is no ratio to compare against and
' accepting the view would be accepting a scale nobody measured. An
' unreadable view scale is not a match for the same reason. Only an actual
' numeric agreement returns True.
Private Function ViewScaleMatchesSheet( _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If swView Is Nothing Or evidence Is Nothing Then Exit Function
    If Not evidence.SheetScaleReadbackProven Then Exit Function
    If evidence.ActualScaleDenominator = 0# Then Exit Function

    Dim sheetDecimal As Double
    sheetDecimal = evidence.ActualScaleNumerator / _
        evidence.ActualScaleDenominator

    ViewScaleMatchesSheet = _
        (Abs(swView.ScaleDecimal - sheetDecimal) <= VIEW_SCALE_TOLERANCE)
    Exit Function

Failed:
    ViewScaleMatchesSheet = False
End Function

' The numbers behind a scale refusal, so the failure text says what was
' compared instead of only that something did not match.
Private Function DescribeViewScale( _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As String

    On Error Resume Next

    Dim viewScale As Double
    Dim flagValue As Long
    Dim sheetDecimal As Double
    Dim proven As Boolean

    viewScale = swView.ScaleDecimal
    flagValue = swView.UseSheetScale

    If Not evidence Is Nothing Then
        proven = evidence.SheetScaleReadbackProven
        If evidence.ActualScaleDenominator <> 0# Then
            sheetDecimal = evidence.ActualScaleNumerator / _
                evidence.ActualScaleDenominator
        End If
    End If

    DescribeViewScale = "useSheetScale=" & CStr(flagValue) & _
        "/viewScale=" & Format$(viewScale, "0.000000") & _
        "/sheetScale=" & Format$(sheetDecimal, "0.000000") & _
        "/sheetScaleProven=" & CStr(proven)
End Function

Private Function ValidateLayout( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef views As Collection, _
    ByVal leftBoundary As Double, _
    ByVal bottomBoundary As Double, _
    ByVal rightBoundary As Double, _
    ByVal topBoundary As Double, _
    ByVal layoutMode As String, _
    ByRef evidence As CRunEvidence) As Boolean

    Dim valid As Boolean
    valid = True

    Dim i As Long
    For i = 1 To views.Count
        Dim firstView As SldWorks.View
        Set firstView = views(i)

        Dim firstOutline As Variant
        firstOutline = firstView.GetOutline

        evidence.AddInfo "LAYOUT_READBACK|mode=" & layoutMode & _
            "|view=" & Module8_RuntimeSupport.GetViewName(firstView) & _
            "|outline=" & Format$(CDbl(firstOutline(0)), "0.000000") & "," & _
            Format$(CDbl(firstOutline(1)), "0.000000") & "," & _
            Format$(CDbl(firstOutline(2)), "0.000000") & "," & _
            Format$(CDbl(firstOutline(3)), "0.000000")

        ' R23-812. Every view's scale flag AND ratio, so the two can never
        ' again be confused for each other.
        evidence.AddInfo "VIEW_SCALE_READBACK|mode=" & layoutMode & _
            "|view=" & Module8_RuntimeSupport.GetViewName(firstView) & _
            "|type=" & CStr(firstView.Type) & _
            "|isIsometric=" & CStr(IsIsometricView(firstView)) & _
            "|useSheetScale=" & CStr(firstView.UseSheetScale) & _
            "|scaleDecimal=" & Format$(firstView.ScaleDecimal, _
                "0.000000")

        If firstView.Type = swDrawingDetailView Then
            If firstView.UseSheetScale <> 0 Or _
               Abs(firstView.ScaleDecimal - 3#) > 0.000001 Then

                evidence.AddFailure _
                    "Detail view is not using its approved 3:1 scale: '" & _
                    Module8_RuntimeSupport.GetViewName(firstView) & "'."
                valid = False
            End If
        ElseIf Not IsIsometricView(firstView) Then
            If firstView.UseSheetScale <> 1 Then
                ' R23-812. The flag is not the ratio. SOLIDWORKS 2025 Help,
                ' IView::UseSheetScale: "If the property is 0, then it is
                ' possible that the view scale is the same as the sheet
                ' scale." A section view is tied to its PARENT view through
                ' IView::UseParentScale, so it reads 0 here while being
                ' drawn at exactly the sheet ratio - which is what failed
                ' the whole LAYOUT stage for Section View J-J in the r49
                ' run. What this stage actually needs to forbid is a view
                ' DRAWN at an unapproved scale, so compare the ratio.
                If Not ViewScaleMatchesSheet(firstView, evidence) Then
                    evidence.AddFailure _
                        "Non-isometric view is not drawn at the proved " & _
                        "sheet scale: '" & _
                        Module8_RuntimeSupport.GetViewName(firstView) & _
                        "' (" & DescribeViewScale(firstView, evidence) & _
                        ")."
                    valid = False
                End If
            End If
        End If

        If CDbl(firstOutline(0)) < leftBoundary Or _
           CDbl(firstOutline(1)) < bottomBoundary Or _
           CDbl(firstOutline(2)) > rightBoundary Or _
           CDbl(firstOutline(3)) > topBoundary Then

            evidence.AddFailure "Layout boundary violation: '" & _
                Module8_RuntimeSupport.GetViewName(firstView) & "'."
            valid = False
        End If

        If OutlineIntersectsTitleBlock(firstOutline, evidence) Then
            evidence.AddFailure "Layout title-block intrusion: '" & _
                Module8_RuntimeSupport.GetViewName(firstView) & "'."
            valid = False
        End If

        If Not swDraw Is Nothing Then
            If Not ViewClearsAllNotes(swDraw, firstOutline, evidence) Then
                evidence.AddFailure "Layout note intrusion: '" & _
                    Module8_RuntimeSupport.GetViewName(firstView) & "'."
                valid = False
            End If
        End If

        Dim j As Long
        For j = i + 1 To views.Count
            Dim secondView As SldWorks.View
            Set secondView = views(j)

            Dim secondOutline As Variant
            secondOutline = secondView.GetOutline

            evidence.AddInfo "LAYOUT_PAIR_CLEARANCE|mode=" & layoutMode & _
                "|first=" & Module8_RuntimeSupport.GetViewName(firstView) & _
                "|second=" & Module8_RuntimeSupport.GetViewName(secondView) & _
                "|horizontalGap=" & Format$( _
                    AxisGap(CDbl(firstOutline(0)), CDbl(firstOutline(2)), _
                            CDbl(secondOutline(0)), CDbl(secondOutline(2))), _
                    "0.000000") & _
                "|verticalGap=" & Format$( _
                    AxisGap(CDbl(firstOutline(1)), CDbl(firstOutline(3)), _
                            CDbl(secondOutline(1)), CDbl(secondOutline(3))), _
                    "0.000000") & _
                "|requiredClearance=" & Format$( _
                    Module8_RuntimeSupport.LAYOUT_MARGIN_M / 2#, "0.000000")

            If OutlinesOverlap(firstOutline, secondOutline) Then
                evidence.AddFailure "View collision [" & layoutMode & _
                    "]: '" & _
                    Module8_RuntimeSupport.GetViewName(firstView) & _
                    "' and '" & _
                    Module8_RuntimeSupport.GetViewName(secondView) & "'."
                valid = False
            End If
        Next j
    Next i

    ValidateLayout = valid
End Function

Private Function OutlineIntersectsTitleBlock( _
    ByVal viewOutline As Variant, _
    ByRef evidence As CRunEvidence) As Boolean

    Dim clearance As Double
    clearance = Module8_RuntimeSupport.LAYOUT_MARGIN_M / 2#

    OutlineIntersectsTitleBlock = Not ( _
        CDbl(viewOutline(2)) + clearance <= evidence.TitleBlockLeft Or _
        evidence.TitleBlockRight + clearance <= CDbl(viewOutline(0)) Or _
        CDbl(viewOutline(3)) + clearance <= evidence.TitleBlockBottom Or _
        evidence.TitleBlockTop + clearance <= CDbl(viewOutline(1)))
End Function

Private Function ViewClearsAllNotes( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal viewOutline As Variant, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim ownerView As SldWorks.View
    Set ownerView = swDraw.GetFirstView

    Do While Not ownerView Is Nothing
        Dim note As SldWorks.Note
        Set note = ownerView.GetFirstNote

        Do While Not note Is Nothing
            Dim renderedNoteText As String
            If Not TryReadRenderedNoteText(note, renderedNoteText) Then
                evidence.AddFailure _
                    "Layout cannot read note text; refusing " & _
                    "to classify its visible extent."
                Exit Function
            End If

            If Len(renderedNoteText) = 0 Then
                evidence.AddInfo "NON_RENDERED_NOTE_SKIPPED|context=Layout"
                GoTo NextNote
            End If

            Dim noteExtent As Variant
            noteExtent = note.GetExtent
            If Not IsArray(noteExtent) Or _
               UBound(noteExtent) - LBound(noteExtent) + 1 < 6 Then

                If Module1_Main.DIAGNOSTIC_DRAWING_MODE And _
                   Not evidence.LayoutBoundariesProven Then

                    evidence.AddWarning _
                        "NOTE_EXTENT_SKIPPED|context=Layout" & _
                        "|reason=UnavailableOrIncomplete|acceptance=False"
                    GoTo NextNote
                End If

                evidence.AddFailure _
                    "Layout cannot read a note extent; refusing " & _
                    "to place a view across an unproved note region."
                Exit Function
            End If

            Dim baseIndex As Long
            baseIndex = LBound(noteExtent)
            Dim noteLeft As Double
            Dim noteBottom As Double
            Dim noteRight As Double
            Dim noteTop As Double
            noteLeft = CDbl(noteExtent(baseIndex))
            noteBottom = CDbl(noteExtent(baseIndex + 1))
            noteRight = CDbl(noteExtent(baseIndex + 3))
            noteTop = CDbl(noteExtent(baseIndex + 4))

            If noteRight < noteLeft Then SwapValues noteLeft, noteRight
            If noteTop < noteBottom Then SwapValues noteBottom, noteTop

            If noteRight <= noteLeft Or noteTop <= noteBottom Then
                If Module1_Main.DIAGNOSTIC_DRAWING_MODE And _
                   Not evidence.LayoutBoundariesProven Then

                    evidence.AddWarning _
                        "NOTE_EXTENT_SKIPPED|context=Layout" & _
                        "|reason=Invalid|acceptance=False"
                    GoTo NextNote
                End If

                evidence.AddFailure _
                    "Layout found an invalid note extent; refusing " & _
                    "to place views around it."
                Exit Function
            End If

            If OutlineIntersectsRectangle( _
                viewOutline, noteLeft, noteBottom, noteRight, noteTop) Then
                Exit Function
            End If

NextNote:
            Set note = note.GetNext
        Loop

        Set ownerView = ownerView.GetNextView
    Loop

    ViewClearsAllNotes = True
    Exit Function

Failed:
    evidence.AddFailure "Layout note-extent inspection error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function TryReadRenderedNoteText( _
    ByRef note As SldWorks.Note, _
    ByRef renderedText As String) As Boolean

    On Error GoTo Failed

    renderedText = note.GetText
    renderedText = Replace$(renderedText, vbCr, vbNullString)
    renderedText = Replace$(renderedText, vbLf, vbNullString)
    renderedText = Trim$(renderedText)
    TryReadRenderedNoteText = True
    Exit Function

Failed:
    renderedText = vbNullString
End Function

Private Function OutlineIntersectsRectangle( _
    ByVal viewOutline As Variant, _
    ByVal rectLeft As Double, _
    ByVal rectBottom As Double, _
    ByVal rectRight As Double, _
    ByVal rectTop As Double) As Boolean

    Dim clearance As Double
    clearance = Module8_RuntimeSupport.LAYOUT_MARGIN_M / 2#

    OutlineIntersectsRectangle = Not ( _
        CDbl(viewOutline(2)) + clearance <= rectLeft Or _
        rectRight + clearance <= CDbl(viewOutline(0)) Or _
        CDbl(viewOutline(3)) + clearance <= rectBottom Or _
        rectTop + clearance <= CDbl(viewOutline(1)))
End Function

Private Sub SwapValues(ByRef firstValue As Double, ByRef secondValue As Double)
    Dim temporaryValue As Double
    temporaryValue = firstValue
    firstValue = secondValue
    secondValue = temporaryValue
End Sub

Private Function IsIsometricView( _
    ByRef swView As SldWorks.View) As Boolean

    On Error GoTo Failed

    Dim orientationName As String
    orientationName = UCase$(Trim$(swView.GetOrientationName))

    IsIsometricView = _
        (InStr(orientationName, "ISOMETRIC") > 0) Or _
        (InStr(orientationName, "TRIMETRIC") > 0) Or _
        (InStr(orientationName, "DIMETRIC") > 0)
    Exit Function

Failed:
    IsIsometricView = False
End Function

Private Function AxisGap( _
    ByVal firstMinimum As Double, _
    ByVal firstMaximum As Double, _
    ByVal secondMinimum As Double, _
    ByVal secondMaximum As Double) As Double

    Dim secondAfterFirst As Double
    Dim firstAfterSecond As Double
    secondAfterFirst = secondMinimum - firstMaximum
    firstAfterSecond = firstMinimum - secondMaximum

    If secondAfterFirst > firstAfterSecond Then
        AxisGap = secondAfterFirst
    Else
        AxisGap = firstAfterSecond
    End If
End Function

Private Function OutlinesOverlap( _
    ByVal firstOutline As Variant, _
    ByVal secondOutline As Variant) As Boolean

    Dim clearance As Double
    clearance = Module8_RuntimeSupport.LAYOUT_MARGIN_M / 2#

    OutlinesOverlap = Not ( _
        CDbl(firstOutline(2)) + clearance <= _
            CDbl(secondOutline(0)) + LAYOUT_COMPARISON_TOLERANCE_M Or _
        CDbl(secondOutline(2)) + clearance <= _
            CDbl(firstOutline(0)) + LAYOUT_COMPARISON_TOLERANCE_M Or _
        CDbl(firstOutline(3)) + clearance <= _
            CDbl(secondOutline(1)) + LAYOUT_COMPARISON_TOLERANCE_M Or _
        CDbl(secondOutline(3)) + clearance <= _
            CDbl(firstOutline(1)) + LAYOUT_COMPARISON_TOLERANCE_M)
End Function

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub

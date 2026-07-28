Option Explicit

' Explicit design tolerances. These values are emitted to QA and must be
' confirmed by the authorized-fixture runtime evidence before acceptance.
Public Const GEOMETRY_TOLERANCE_M As Double = 0.0000001
Public Const PROJECTED_TOLERANCE_M As Double = 0.00001
Public Const AXIS_NORMAL_MIN_COS As Double = 0.999847695
Public Const LAYOUT_MARGIN_M As Double = 0.012

Private Const swRebuildActiveDoc As Long = 2
Private Const swObjectSame As Long = 1
Private Const swModelRebuildStatus_FullyRebuilt As Long = 0

Private Const swDrawingProjectedView As Long = 4
Private Const swDrawingStandardView As Long = 6
Private Const swDrawingNamedView As Long = 7
Private Const swSketchLINE As Long = 0
Private Const swZoneTopMargin As Long = 0
Private Const swZoneBottomMargin As Long = 1
Private Const swZoneRightMargin As Long = 2
Private Const swZoneLeftMargin As Long = 3

Private Const CONTROLLED_LEGACY_FORMAT_TOKEN As String = _
    "veemap drawing.slddrt"
Private Const TEMPLATE_GEOMETRY_TOLERANCE_M As Double = 0.00001

Private mProvenOrdinateViews As Object

Public Sub ResetProvenViewRegistry()
    Set mProvenOrdinateViews = CreateObject("Scripting.Dictionary")
End Sub

Public Sub RegisterProvenOrdinateView( _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence)

    If mProvenOrdinateViews Is Nothing Then ResetProvenViewRegistry

    Dim viewName As String
    viewName = LCase$(GetViewName(swView))

    If Len(viewName) = 0 Then
        evidence.AddFailure "Cannot register an unnamed ordinate-eligible view."
        Exit Sub
    End If

    If Not mProvenOrdinateViews.Exists(viewName) Then
        mProvenOrdinateViews.Add viewName, True
    End If

    evidence.AddInfo "Registered created orthographic view as proven for ordinate " & _
        "policy: '" & GetViewName(swView) & "'."
End Sub

Public Function ActivateDrawingDocument( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim errors As Long
    Dim activated As SldWorks.ModelDoc2

    Set activated = swApp.ActivateDoc3( _
        swDrawModel.GetTitle, False, swRebuildActiveDoc, errors)

    If activated Is Nothing Then
        evidence.AddFailure "ActivateDoc3 returned Nothing; error=" & CStr(errors)
        Exit Function
    End If

    If Not ObjectsAreSame(swApp, activated, swDrawModel) Then
        evidence.AddFailure "ActivateDoc3 activated a different document."
        Exit Function
    End If

    ActivateDrawingDocument = True
    Exit Function

Failed:
    evidence.AddFailure "ActivateDoc3 error " & CStr(Err.Number) & _
        ": " & Err.Description
End Function

Public Function ActivateDrawingView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence, _
    ByVal operationName As String) As Boolean

    On Error GoTo Failed

    If swView Is Nothing Then
        evidence.AddFailure operationName & ": view is Nothing."
        Exit Function
    End If

    Dim viewName As String
    viewName = GetViewName(swView)

    If Len(viewName) = 0 Then
        evidence.AddFailure operationName & ": view has no displayed name."
        Exit Function
    End If

    Dim activateResult As Boolean
    activateResult = swDraw.ActivateView(viewName)

    Dim activeView As SldWorks.View
    Set activeView = swDraw.ActiveDrawingView

    Dim activeViewMatches As Boolean
    activeViewMatches = False
    If Not activeView Is Nothing Then
        activeViewMatches = (StrComp( _
            GetViewName(activeView), viewName, vbTextCompare) = 0)
    End If

    If Not activeViewMatches Then

        swDrawModel.ClearSelection2 True

        Dim selectedView As Boolean
        selectedView = swDrawModel.Extension.SelectByID2( _
            viewName, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0)

        If selectedView Then
            activateResult = swDraw.ActivateView(viewName)
            Set activeView = swDraw.ActiveDrawingView
        End If

        swDrawModel.ClearSelection2 True
        evidence.AddInfo "ACTIVATE_VIEW_RETRY|operation=" & operationName & _
            "|view=" & viewName & _
            "|selectedByID=" & CStr(selectedView) & _
            "|setterResult=" & CStr(activateResult)
    End If

    If activeView Is Nothing Then
        evidence.AddFailure operationName & _
            ": ActivateView returned " & CStr(activateResult) & _
            " and no drawing view is active for '" & viewName & "'."
        Exit Function
    End If

    If StrComp(GetViewName(activeView), viewName, vbTextCompare) <> 0 Then
        evidence.AddFailure operationName & _
            ": active-view readback mismatch; requested='" & viewName & _
            "', actual='" & GetViewName(activeView) & "'."
        Exit Function
    End If

    If Not activateResult Then
        evidence.AddWarning "ACTIVATE_VIEW|operation=" & operationName & _
            "|view=" & viewName & _
            "|setterResult=False|readbackMatched=True"
    End If

    ActivateDrawingView = True
    Exit Function

Failed:
    evidence.AddFailure operationName & ": ActivateView error " & _
        CStr(Err.Number) & " - " & Err.Description
End Function

Public Sub RestoreSheetContext( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc)

    On Error Resume Next
    If Not Module1_Main.GlobalEvidence Is Nothing Then
        Module1_Main.GlobalEvidence.RecordSolidWorksMutation _
            "RestoreSheetContext"
    End If

    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If Not swSheet Is Nothing Then
        swDraw.ActivateSheet swSheet.GetName
    End If
    On Error GoTo 0
End Sub

Public Function GetViewName(ByRef swView As SldWorks.View) As String
    On Error Resume Next
    GetViewName = swView.GetName2
    On Error GoTo 0
End Function

Public Function IsOrdinateEligibleView( _
    ByRef swView As SldWorks.View) As Boolean

    On Error GoTo NotEligible

    Dim viewType As Long
    viewType = swView.Type

    If viewType <> swDrawingProjectedView And _
       viewType <> swDrawingStandardView And _
       viewType <> swDrawingNamedView Then Exit Function

    If viewType = swDrawingNamedView Then
        If mProvenOrdinateViews Is Nothing Then Exit Function
        If Not mProvenOrdinateViews.Exists( _
            LCase$(GetViewName(swView))) Then Exit Function
    End If

    Dim orientationName As String
    orientationName = UCase$(Trim$(swView.GetOrientationName))

    If InStr(orientationName, "ISOMETRIC") > 0 Then Exit Function
    If InStr(orientationName, "TRIMETRIC") > 0 Then Exit Function
    If InStr(orientationName, "DIMETRIC") > 0 Then Exit Function

    Select Case orientationName
        Case "*FRONT", "FRONT", "*BACK", "BACK", _
             "*LEFT", "LEFT", "*RIGHT", "RIGHT", _
             "*TOP", "TOP", "*BOTTOM", "BOTTOM"
            IsOrdinateEligibleView = True
    End Select
    Exit Function

NotEligible:
    IsOrdinateEligibleView = False
End Function

Public Function MeasureControlledSheetRegions( _
    ByRef swSheet As SldWorks.Sheet, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed
    evidence.RequireStage "CONTROLLED_SHEET"
    evidence.LayoutBoundariesProven = False

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim paperSize As Long

    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)
    evidence.SheetWidth = sheetWidth
    evidence.SheetHeight = sheetHeight

    If evidence.SheetWidth <= 0# Or evidence.SheetHeight <= 0# Then
        evidence.AddFailure "Controlled sheet has invalid physical dimensions."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "invalid physical dimensions"
        Exit Function
    End If

    evidence.SheetTemplateName = Trim$(swSheet.GetTemplateName)
    evidence.SheetFormatName = Trim$(swSheet.GetSheetFormatName)

    If Len(evidence.SheetTemplateName) = 0 Then
        evidence.AddFailure "Controlled sheet has no template path/name."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "missing template path or name"
        Exit Function
    End If

    If Len(evidence.SheetFormatName) = 0 Then
        evidence.AddFailure "Controlled sheet has no sheet format."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "missing sheet format"
        Exit Function
    End If

    Dim formatVisible As Boolean
    formatVisible = swSheet.SheetFormatVisible

    If Not formatVisible Then
        evidence.RecordSolidWorksMutation "EnsureSheetFormatVisible"
        swSheet.SheetFormatVisible = True
        formatVisible = swSheet.SheetFormatVisible
    End If

    If Not formatVisible Then
        evidence.AddWarning _
            "ISheet.SheetFormatVisible remained False; continuing to structural " & _
            "template, title-block, margin, and usable-area proof."
    End If

    Dim titleBlock As SldWorks.TitleBlock
    Set titleBlock = swSheet.TitleBlock

    Dim titleX1 As Double
    Dim titleY1 As Double
    Dim titleX2 As Double
    Dim titleY2 As Double
    Dim titleBoundsSource As String

    If titleBlock Is Nothing Then
        If Not TryMeasureLegacyControlledTitleBlock( _
            swSheet, sheetWidth, sheetHeight, titleX1, titleY1, _
            titleX2, titleY2, evidence) Then

            evidence.AddFailure "Controlled sheet has neither an ITitleBlock " & _
                "definition nor a proved legacy title-block rectangle."
            evidence.MarkStageFailed "CONTROLLED_SHEET", _
                "formal and legacy title-block proofs failed"
            Exit Function
        End If

        titleBoundsSource = _
            "ISheet.GetTemplateSketch/ISketch.GetSketchSegments"
    Else
        titleBlock.GetExtents titleX1, titleY1, titleX2, titleY2
        titleBoundsSource = "ITitleBlock.GetExtents"
    End If

    evidence.TitleBlockLeft = IIf(titleX1 < titleX2, titleX1, titleX2)
    evidence.TitleBlockRight = IIf(titleX1 > titleX2, titleX1, titleX2)
    evidence.TitleBlockBottom = IIf(titleY1 < titleY2, titleY1, titleY2)
    evidence.TitleBlockTop = IIf(titleY1 > titleY2, titleY1, titleY2)

    If evidence.TitleBlockRight <= evidence.TitleBlockLeft Or _
       evidence.TitleBlockTop <= evidence.TitleBlockBottom Then

        evidence.AddFailure "Controlled title-block proof returned invalid bounds."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "invalid title-block extents"
        Exit Function
    End If

    If evidence.TitleBlockLeft < 0# Or _
       evidence.TitleBlockBottom < 0# Or _
       evidence.TitleBlockRight > evidence.SheetWidth Or _
       evidence.TitleBlockTop > evidence.SheetHeight Then

        evidence.AddFailure "Measured title-block bounds lie outside the sheet."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "title-block extents outside sheet"
        Exit Function
    End If

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

        evidence.AddFailure "Controlled zoned-border margins are unavailable or zero."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "one or more zone margins are zero"
        Exit Function
    End If

    evidence.ContentBorderLeft = leftMargin
    evidence.ContentBorderBottom = bottomMargin
    evidence.ContentBorderRight = evidence.SheetWidth - rightMargin
    evidence.ContentBorderTop = evidence.SheetHeight - topMargin

    If evidence.ContentBorderRight <= evidence.ContentBorderLeft Or _
       evidence.ContentBorderTop <= evidence.ContentBorderBottom Then

        evidence.AddFailure "Controlled zoned-border margins produce invalid content bounds."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "zone-margin content bounds are invalid"
        Exit Function
    End If

    evidence.UsableLeft = leftMargin + LAYOUT_MARGIN_M
    evidence.UsableRight = evidence.SheetWidth - rightMargin - LAYOUT_MARGIN_M
    evidence.UsableTop = evidence.SheetHeight - topMargin - LAYOUT_MARGIN_M
    evidence.UsableBottom = evidence.TitleBlockTop + LAYOUT_MARGIN_M

    If evidence.UsableBottom < bottomMargin + LAYOUT_MARGIN_M Then
        evidence.UsableBottom = bottomMargin + LAYOUT_MARGIN_M
    End If

    If evidence.UsableRight <= evidence.UsableLeft Or _
       evidence.UsableTop <= evidence.UsableBottom Then

        evidence.AddFailure "Measured border/title-block reserve leaves no usable view area."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "no usable view area remains"
        Exit Function
    End If

    evidence.LayoutBoundariesProven = True
    evidence.AddInfo "CONTROLLED_REGIONS|content=" & _
        Format$(evidence.ContentBorderLeft, "0.000000") & "," & _
        Format$(evidence.ContentBorderBottom, "0.000000") & "," & _
        Format$(evidence.ContentBorderRight, "0.000000") & "," & _
        Format$(evidence.ContentBorderTop, "0.000000") & _
        "|viewUsable=" & Format$(evidence.UsableLeft, "0.000000") & "," & _
        Format$(evidence.UsableBottom, "0.000000") & "," & _
        Format$(evidence.UsableRight, "0.000000") & "," & _
        Format$(evidence.UsableTop, "0.000000") & _
        "|titleSource=" & titleBoundsSource
    evidence.MarkStageProved "CONTROLLED_SHEET", _
        "template, format geometry, zone margins, title block, and usable bounds proved"
    MeasureControlledSheetRegions = True
    Exit Function

Failed:
    evidence.AddFailure "Controlled sheet measurement error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed "CONTROLLED_SHEET", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function TryMeasureLegacyControlledTitleBlock( _
    ByRef swSheet As SldWorks.Sheet, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef titleLeft As Double, _
    ByRef titleBottom As Double, _
    ByRef titleRight As Double, _
    ByRef titleTop As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim templateName As String
    templateName = LCase$(Trim$(swSheet.GetTemplateName))

    If InStr(1, templateName, CONTROLLED_LEGACY_FORMAT_TOKEN, _
        vbTextCompare) = 0 Then Exit Function

    If Abs(sheetWidth - 0.42) > TEMPLATE_GEOMETRY_TOLERANCE_M Or _
       Abs(sheetHeight - 0.297) > TEMPLATE_GEOMETRY_TOLERANCE_M Then

        evidence.AddFailure "Legacy controlled-format proof rejected an " & _
            "unexpected sheet size."
        Exit Function
    End If

    Dim templateSketch As SldWorks.Sketch
    Set templateSketch = swSheet.GetTemplateSketch
    If templateSketch Is Nothing Then Exit Function

    Dim segments As Variant
    segments = templateSketch.GetSketchSegments
    If IsEmpty(segments) Or Not IsArray(segments) Then Exit Function

    Dim zoneLeft As Double
    Dim zoneRight As Double
    Dim zoneBottom As Double
    Dim zoneTop As Double
    zoneLeft = sheetWidth * 0.63
    zoneRight = sheetWidth * 0.98
    ' zoneBottom must sit below the smallest realistic inner-border offset.  At
    ' 0.03 the accepted band for the title-block bottom rule was only 8.9-17.8mm
    ' on A3, which silently excluded any format drawn against a 20mm border.
    zoneBottom = sheetHeight * 0.015
    zoneTop = sheetHeight * 0.27

    titleLeft = sheetWidth
    titleBottom = sheetHeight
    titleRight = 0#
    titleTop = 0#

    Dim candidateCount As Long
    Dim bottomBoundaryProved As Boolean
    Dim topBoundaryProved As Boolean
    Dim leftBoundaryProved As Boolean
    Dim rightBoundaryProved As Boolean

    Dim i As Long
    For i = LBound(segments) To UBound(segments)
        Dim segment As SldWorks.SketchSegment
        Set segment = segments(i)

        If Not segment Is Nothing Then
            If segment.GetType = swSketchLINE Then
                Dim sketchLine As SldWorks.SketchLine
                Set sketchLine = segment

                Dim startPoint As SldWorks.SketchPoint
                Dim endPoint As SldWorks.SketchPoint
                Set startPoint = sketchLine.GetStartPoint2
                Set endPoint = sketchLine.GetEndPoint2

                If Not startPoint Is Nothing And Not endPoint Is Nothing Then
                    Dim x1 As Double
                    Dim y1 As Double
                    Dim x2 As Double
                    Dim y2 As Double
                    x1 = startPoint.X
                    y1 = startPoint.Y
                    x2 = endPoint.X
                    y2 = endPoint.Y

                    If x1 >= zoneLeft - TEMPLATE_GEOMETRY_TOLERANCE_M And _
                       x2 >= zoneLeft - TEMPLATE_GEOMETRY_TOLERANCE_M And _
                       x1 <= zoneRight + TEMPLATE_GEOMETRY_TOLERANCE_M And _
                       x2 <= zoneRight + TEMPLATE_GEOMETRY_TOLERANCE_M And _
                       y1 >= zoneBottom - TEMPLATE_GEOMETRY_TOLERANCE_M And _
                       y2 >= zoneBottom - TEMPLATE_GEOMETRY_TOLERANCE_M And _
                       y1 <= zoneTop + TEMPLATE_GEOMETRY_TOLERANCE_M And _
                       y2 <= zoneTop + TEMPLATE_GEOMETRY_TOLERANCE_M Then

                        candidateCount = candidateCount + 1
                        If x1 < titleLeft Then titleLeft = x1
                        If x2 < titleLeft Then titleLeft = x2
                        If x1 > titleRight Then titleRight = x1
                        If x2 > titleRight Then titleRight = x2
                        If y1 < titleBottom Then titleBottom = y1
                        If y2 < titleBottom Then titleBottom = y2
                        If y1 > titleTop Then titleTop = y1
                        If y2 > titleTop Then titleTop = y2

                        Dim lineLength As Double
                        lineLength = Sqr((x2 - x1) * (x2 - x1) + _
                                         (y2 - y1) * (y2 - y1))

                        If Abs(y2 - y1) <= _
                           TEMPLATE_GEOMETRY_TOLERANCE_M And _
                           lineLength >= 0.1 Then

                            If y1 <= sheetHeight * 0.09 Then
                                bottomBoundaryProved = True
                            ElseIf y1 >= sheetHeight * 0.15 Then
                                topBoundaryProved = True
                            End If
                        End If

                        If Abs(x2 - x1) <= _
                           TEMPLATE_GEOMETRY_TOLERANCE_M And _
                           lineLength >= 0.04 Then

                            If x1 <= sheetWidth * 0.75 Then
                                leftBoundaryProved = True
                            ElseIf x1 >= sheetWidth * 0.95 Then
                                rightBoundaryProved = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    Dim titleWidth As Double
    Dim titleHeight As Double
    titleWidth = titleRight - titleLeft
    titleHeight = titleTop - titleBottom

    ' titleLeft/titleTop are accumulated only from segments that already passed
    ' the zoneLeft/zoneTop candidate filter, so comparing them back against
    ' those same bounds proves nothing.  Anchor the rectangle instead: it must
    ' reach the sheet's right edge and sit against the bottom of the sheet.
    If candidateCount < 10 Or _
       Not bottomBoundaryProved Or Not topBoundaryProved Or _
       Not leftBoundaryProved Or Not rightBoundaryProved Or _
       titleWidth < 0.1 Or titleWidth > 0.18 Or _
       titleHeight < 0.04 Or titleHeight > 0.09 Or _
       titleRight < sheetWidth * 0.95 Or _
       titleBottom > sheetHeight * 0.1 Then

        evidence.AddFailure "Legacy title-block sketch geometry did not " & _
            "satisfy the controlled lower-right rectangle contract."
        evidence.AddInfo "TITLE_BOUNDS_REJECTED|segments=" & _
            CStr(candidateCount) & _
            "|left=" & Format$(titleLeft, "0.000000") & _
            "|bottom=" & Format$(titleBottom, "0.000000") & _
            "|right=" & Format$(titleRight, "0.000000") & _
            "|top=" & Format$(titleTop, "0.000000") & _
            "|bottomEdge=" & CStr(bottomBoundaryProved) & _
            "|topEdge=" & CStr(topBoundaryProved) & _
            "|leftEdge=" & CStr(leftBoundaryProved) & _
            "|rightEdge=" & CStr(rightBoundaryProved)
        Exit Function
    End If

    evidence.AddInfo "TITLE_BOUNDS|source=LegacyTemplateSketch" & _
        "|segments=" & CStr(candidateCount) & _
        "|left=" & Format$(titleLeft, "0.000000") & _
        "|bottom=" & Format$(titleBottom, "0.000000") & _
        "|right=" & Format$(titleRight, "0.000000") & _
        "|top=" & Format$(titleTop, "0.000000") & _
        "|structuralEdgesProved=True"

    TryMeasureLegacyControlledTitleBlock = True
    Exit Function

Failed:
    evidence.AddFailure "Legacy title-block measurement error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub FinalizeSelectionAndSheetState( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed
    evidence.RequireStage "FINAL_CLEANUP"
    evidence.FinalCleanupVerified = False
    evidence.RecordSolidWorksMutation "FinalizeSelectionAndSheetState"

    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        evidence.AddFailure "Final cleanup could not obtain the current sheet."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "current sheet is Nothing"
        Exit Sub
    End If

    Dim activateResult As Boolean
    activateResult = swDraw.ActivateSheet(swSheet.GetName)

    If Not activateResult Then
        Dim currentSheet As SldWorks.Sheet
        Set currentSheet = swDraw.GetCurrentSheet

        Dim currentNameMatches As Boolean
        currentNameMatches = False
        If Not currentSheet Is Nothing Then
            currentNameMatches = (StrComp(currentSheet.GetName, _
                swSheet.GetName, vbTextCompare) = 0)
        End If

        If currentNameMatches And swDraw.ActiveDrawingView Is Nothing Then
            evidence.AddWarning "FINAL_SHEET_CONTEXT|ActivateSheet=False|" & _
                "readbackMatched=True|sheet=" & swSheet.GetName
        Else
            evidence.AddFailure "Final ActivateSheet returned False and sheet-context " & _
                "readback did not prove the requested sheet."
            evidence.MarkStageFailed "FINAL_CLEANUP", _
                "ActivateSheet false with mismatched context"
            Exit Sub
        End If
    End If

    VerifyFinalCleanupState swDrawModel, swDraw, evidence
    Exit Sub

Failed:
    evidence.AddFailure "Final selection/sheet cleanup error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.FinalCleanupVerified = False
    evidence.MarkStageFailed "FINAL_CLEANUP", _
        "cleanup API error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Public Function VerifyFinalCleanupState( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed
    evidence.RequireStage "FINAL_CLEANUP"

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        evidence.AddFailure "Final cleanup verification has no current sheet."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "current sheet is Nothing"
        Exit Function
    End If

    evidence.FinalSelectionCount = _
        swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)
    evidence.SheetContextRestored = (swDraw.ActiveDrawingView Is Nothing)

    If evidence.FinalSelectionCount <> 0 Then
        evidence.AddFailure "Final selection list is not empty."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "selection count=" & CStr(evidence.FinalSelectionCount)
        Exit Function
    End If

    If Not evidence.SheetContextRestored Then
        evidence.AddFailure "Final active context is not the drawing sheet."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "active drawing view is not Nothing"
        Exit Function
    End If

    evidence.FinalCleanupVerified = True
    evidence.FinalCleanupInvalidated = False
    evidence.MarkStageProved "FINAL_CLEANUP", _
        "normal pick mode requested, sheet active, selection count zero"
    evidence.AddInfo "Final SetPickMode invoked; selection count=0."
    VerifyFinalCleanupState = True
    Exit Function

Failed:
    evidence.FinalCleanupVerified = False
    evidence.AddFailure "Final cleanup verification error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed "FINAL_CLEANUP", _
        "verification API error " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function ObjectsAreSame( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal firstObject As Object, _
    ByVal secondObject As Object) As Boolean

    If firstObject Is Nothing Then Exit Function
    If secondObject Is Nothing Then Exit Function

    On Error Resume Next
    ObjectsAreSame = (swApp.IsSame(firstObject, secondObject) = swObjectSame)
    On Error GoTo 0
End Function

Public Function CountVariantItems(ByVal items As Variant) As Long
    On Error GoTo Failed

    If IsEmpty(items) Or IsNull(items) Then Exit Function

    If IsArray(items) Then
        CountVariantItems = UBound(items) - LBound(items) + 1
    Else
        CountVariantItems = 1
    End If
    Exit Function

Failed:
    CountVariantItems = 0
End Function

Public Function SetConfiguredSheetScale( _
    ByRef swSheet As SldWorks.Sheet, _
    ByVal decimalScale As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    evidence.RequireStage "SHEET_SCALE"
    evidence.SheetScaleReadbackProven = False

    If decimalScale <= 0# Then
        evidence.AddFailure "Configured sheet scale is not positive."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "configured scale is not positive"
        Exit Function
    End If

    Dim numerator As Long
    Dim denominator As Long
    DecimalToRatio decimalScale, numerator, denominator
    evidence.RequestedScaleNumerator = numerator
    evidence.RequestedScaleDenominator = denominator

    On Error GoTo Failed
    evidence.RecordSolidWorksMutation "ISheet.SetScale"
    Dim setterResult As Boolean
    setterResult = swSheet.SetScale( _
        CDbl(numerator), CDbl(denominator), True, False)

    Dim sheetProperties As Variant
    sheetProperties = swSheet.GetProperties2

    If Not IsArray(sheetProperties) Then
        evidence.AddFailure "ISheet.GetProperties2 did not return a scale array."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "scale readback was not an array"
        SetConfiguredSheetScale = False
        Exit Function
    End If

    If UBound(sheetProperties) < 3 Then
        evidence.AddFailure "ISheet.GetProperties2 scale array is incomplete."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "scale readback contained fewer than four values"
        SetConfiguredSheetScale = False
        Exit Function
    End If

    evidence.ActualScaleNumerator = CDbl(sheetProperties(2))
    evidence.ActualScaleDenominator = CDbl(sheetProperties(3))

    If evidence.ActualScaleNumerator <= 0# Or _
       evidence.ActualScaleDenominator <= 0# Then

        evidence.AddFailure "Actual sheet-scale readback is not positive."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "nonpositive scale readback"
        SetConfiguredSheetScale = False
        Exit Function
    End If

    Dim actualDecimalScale As Double
    actualDecimalScale = evidence.ActualScaleNumerator / _
        evidence.ActualScaleDenominator

    If Abs(actualDecimalScale - decimalScale) > 0.000000001 Then
        evidence.AddFailure "Actual sheet scale does not match the requested scale."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "requested=" & Format$(decimalScale, "0.000000000") & _
            "; actual=" & Format$(actualDecimalScale, "0.000000000")
        SetConfiguredSheetScale = False
        Exit Function
    End If

    evidence.SheetScaleReadbackProven = True
    SetConfiguredSheetScale = True
    evidence.MarkStageProved "SHEET_SCALE", _
        "requested and GetProperties2 scale ratios match"
    If Not setterResult Then
        evidence.AddWarning "SHEET_SCALE|setterResult=False|" & _
            "readbackMatched=True|ratio=" & _
            Format$(evidence.ActualScaleNumerator, "0.###") & ":" & _
            Format$(evidence.ActualScaleDenominator, "0.###")
    End If
    evidence.AddInfo "Sheet scale set and read back as " & _
        Format$(evidence.ActualScaleNumerator, "0.###") & ":" & _
        Format$(evidence.ActualScaleDenominator, "0.###") & "."
    Exit Function

Failed:
    evidence.AddFailure "ISheet.SetScale error " & CStr(Err.Number) & _
        ": " & Err.Description
    evidence.SheetScaleReadbackProven = False
    evidence.MarkStageFailed "SHEET_SCALE", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
    SetConfiguredSheetScale = False
End Function

Public Function RebuildDocumentVerified( _
    ByRef swModel As SldWorks.ModelDoc2, _
    ByVal operationName As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    evidence.RecordSolidWorksMutation operationName & " EditRebuild3"

    Dim rebuildResult As Boolean
    Dim rebuildStatus As Long
    rebuildResult = swModel.EditRebuild3
    rebuildStatus = swModel.Extension.NeedsRebuild2

    If rebuildStatus <> swModelRebuildStatus_FullyRebuilt Then
        evidence.AddFailure "REBUILD|operation=" & operationName & _
            "|setterResult=" & CStr(rebuildResult) & _
            "|status=" & CStr(rebuildStatus)
        Exit Function
    End If

    If Not rebuildResult Then
        evidence.AddWarning "REBUILD|operation=" & operationName & _
            "|setterResult=False|readbackFullyRebuilt=True"
    End If

    RebuildDocumentVerified = True
    Exit Function

Failed:
    evidence.AddFailure "Rebuild verification error for " & operationName & _
        ": " & CStr(Err.Number) & " - " & Err.Description
End Function

Private Sub DecimalToRatio( _
    ByVal value As Double, _
    ByRef numerator As Long, _
    ByRef denominator As Long)

    denominator = 1000
    numerator = CLng(value * CDbl(denominator) + 0.5)

    Dim divisor As Long
    divisor = GreatestCommonDivisor(Abs(numerator), denominator)

    If divisor > 0 Then
        numerator = numerator \ divisor
        denominator = denominator \ divisor
    End If

    If numerator < 1 Then numerator = 1
    If denominator < 1 Then denominator = 1
End Sub

Private Function GreatestCommonDivisor( _
    ByVal firstValue As Long, _
    ByVal secondValue As Long) As Long

    Dim remainder As Long
    Do While secondValue <> 0
        remainder = firstValue Mod secondValue
        firstValue = secondValue
        secondValue = remainder
    Loop

    GreatestCommonDivisor = firstValue
End Function

Public Function TransformPointToView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal z As Double, _
    ByRef viewX As Double, _
    ByRef viewY As Double, _
    ByRef viewZ As Double, _
    ByRef coordinateFrameProof As String, _
    ByVal requireOutlineContainment As Boolean) As Boolean

    On Error GoTo Failed

    coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=unknown" & _
        "|reason=TransformNotAttempted"

    If swApp Is Nothing Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=unknown" & _
            "|reason=ApplicationUnavailable"
        Exit Function
    End If

    If swView Is Nothing Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=unknown" & _
            "|reason=ViewUnavailable"
        Exit Function
    End If

    Dim mathUtil As SldWorks.MathUtility
    Set mathUtil = swApp.GetMathUtility
    If mathUtil Is Nothing Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & _
            TransformTokenValue(GetViewName(swView)) & _
            "|reason=MathUtilityUnavailable"
        Exit Function
    End If

    Dim pointData(0 To 2) As Double
    pointData(0) = x
    pointData(1) = y
    pointData(2) = z

    Dim mathPoint As SldWorks.MathPoint
    Set mathPoint = mathUtil.CreatePoint(pointData)
    If mathPoint Is Nothing Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & _
            TransformTokenValue(GetViewName(swView)) & _
            "|reason=MathPointUnavailable"
        Exit Function
    End If

    Dim modelToView As SldWorks.MathTransform
    Set modelToView = swView.ModelToViewTransform
    If modelToView Is Nothing Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & _
            TransformTokenValue(GetViewName(swView)) & _
            "|reason=ModelToViewTransformUnavailable"
        Exit Function
    End If

    Set mathPoint = mathPoint.MultiplyTransform(modelToView)
    If mathPoint Is Nothing Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & _
            TransformTokenValue(GetViewName(swView)) & _
            "|reason=ModelToViewTransformReturnedNothing"
        Exit Function
    End If

    Dim result As Variant
    result = mathPoint.ArrayData
    If Not IsArray(result) Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & _
            TransformTokenValue(GetViewName(swView)) & _
            "|reason=TransformArrayUnavailable"
        Exit Function
    End If

    viewX = CDbl(result(0))
    viewY = CDbl(result(1))
    viewZ = CDbl(result(2))

    If Not ProvePageCoordinateAgainstViewOutline( _
        swView, viewX, viewY, requireOutlineContainment, _
        coordinateFrameProof) Then Exit Function

    TransformPointToView = True
    Exit Function

Failed:
    Dim transformErrorNumber As Long
    transformErrorNumber = Err.Number
    coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & _
        TransformTokenValue(SafeTransformViewName(swView)) & _
        "|reason=APIError:" & CStr(transformErrorNumber)
    TransformPointToView = False
End Function

' requireOutlineContainment must be True for points that are claimed to lie on
' visible geometry (hole centres, mapped model vertices).  It must be False for
' reference points such as the projected model origin, which legitimately fall
' outside the view outline whenever the part origin sits off the solid.
Private Function ProvePageCoordinateAgainstViewOutline( _
    ByRef swView As SldWorks.View, _
    ByVal pageX As Double, _
    ByVal pageY As Double, _
    ByVal requireOutlineContainment As Boolean, _
    ByRef coordinateFrameProof As String) As Boolean

    On Error GoTo Failed

    Dim viewName As String
    viewName = TransformTokenValue(SafeTransformViewName(swView))

    Dim outline As Variant
    outline = swView.GetOutline

    If Not IsArray(outline) Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & viewName & _
            "|reason=ViewOutlineUnavailable"
        Exit Function
    End If

    Dim outlineLowerBound As Long
    Dim outlineUpperBound As Long
    outlineLowerBound = LBound(outline)
    outlineUpperBound = UBound(outline)

    If outlineUpperBound - outlineLowerBound < 3 Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & viewName & _
            "|reason=ViewOutlineMalformed"
        Exit Function
    End If

    Dim minimumX As Double
    Dim minimumY As Double
    Dim maximumX As Double
    Dim maximumY As Double
    minimumX = CDbl(outline(outlineLowerBound))
    minimumY = CDbl(outline(outlineLowerBound + 1))
    maximumX = CDbl(outline(outlineLowerBound + 2))
    maximumY = CDbl(outline(outlineLowerBound + 3))

    If minimumX <> minimumX Or minimumY <> minimumY Or _
       maximumX <> maximumX Or maximumY <> maximumY Then

        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & viewName & _
            "|reason=ViewOutlineNotFinite"
        Exit Function
    End If

    If maximumX <= minimumX Or maximumY <= minimumY Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & viewName & _
            "|reason=ViewOutlineDegenerate" & _
            "|outline=" & TransformOutlineToken( _
                minimumX, minimumY, maximumX, maximumY)
        Exit Function
    End If

    If pageX <> pageX Or pageY <> pageY Then
        coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & viewName & _
            "|reason=TransformedCoordinateNotFinite"
        Exit Function
    End If

    If pageX < minimumX - PROJECTED_TOLERANCE_M Or _
       pageX > maximumX + PROJECTED_TOLERANCE_M Or _
       pageY < minimumY - PROJECTED_TOLERANCE_M Or _
       pageY > maximumY + PROJECTED_TOLERANCE_M Then

        If requireOutlineContainment Then
            coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & viewName & _
                "|reason=OutsideCurrentViewOutline" & _
                "|pageXY=" & TransformCoordinateToken(pageX) & "," & _
                    TransformCoordinateToken(pageY) & _
                "|outline=" & TransformOutlineToken( _
                    minimumX, minimumY, maximumX, maximumY) & _
                "|tolerance=" & TransformCoordinateToken(PROJECTED_TOLERANCE_M)
            Exit Function
        End If

        coordinateFrameProof = "TRANSFORM_PAGE_PROOF|view=" & viewName & _
            "|source=ModelToViewTransform" & _
            "|frame=DrawingPage" & _
            "|basis=IView.GetOutline" & _
            "|containment=NotRequired" & _
            "|pageXY=" & TransformCoordinateToken(pageX) & "," & _
                TransformCoordinateToken(pageY) & _
            "|outline=" & TransformOutlineToken( _
                minimumX, minimumY, maximumX, maximumY) & _
            "|tolerance=" & TransformCoordinateToken(PROJECTED_TOLERANCE_M)

        ProvePageCoordinateAgainstViewOutline = True
        Exit Function
    End If

    coordinateFrameProof = "TRANSFORM_PAGE_PROOF|view=" & viewName & _
        "|source=ModelToViewTransform" & _
        "|frame=DrawingPage" & _
        "|basis=IView.GetOutline" & _
        "|pageXY=" & TransformCoordinateToken(pageX) & "," & _
            TransformCoordinateToken(pageY) & _
        "|outline=" & TransformOutlineToken( _
            minimumX, minimumY, maximumX, maximumY) & _
        "|tolerance=" & TransformCoordinateToken(PROJECTED_TOLERANCE_M)

    ProvePageCoordinateAgainstViewOutline = True
    Exit Function

Failed:
    Dim outlineErrorNumber As Long
    outlineErrorNumber = Err.Number
    coordinateFrameProof = "TRANSFORM_PAGE_REJECT|view=" & _
        TransformTokenValue(SafeTransformViewName(swView)) & _
        "|reason=OutlineProofError:" & CStr(outlineErrorNumber)
End Function

Private Function TransformOutlineToken( _
    ByVal minimumX As Double, _
    ByVal minimumY As Double, _
    ByVal maximumX As Double, _
    ByVal maximumY As Double) As String

    TransformOutlineToken = _
        TransformCoordinateToken(minimumX) & "," & _
        TransformCoordinateToken(minimumY) & "," & _
        TransformCoordinateToken(maximumX) & "," & _
        TransformCoordinateToken(maximumY)
End Function

Private Function TransformCoordinateToken( _
    ByVal coordinateValue As Double) As String

    TransformCoordinateToken = Format$(coordinateValue, "0.000000000")
End Function

Private Function SafeTransformViewName( _
    ByRef swView As SldWorks.View) As String

    SafeTransformViewName = "unknown"
    If swView Is Nothing Then Exit Function

    On Error GoTo Failed
    SafeTransformViewName = GetViewName(swView)
    If Len(Trim$(SafeTransformViewName)) = 0 Then _
        SafeTransformViewName = "unnamed"
    Exit Function

Failed:
    SafeTransformViewName = "unknown"
End Function

Private Function TransformTokenValue(ByVal value As String) As String
    TransformTokenValue = Replace$(Trim$(value), "|", "/")
    TransformTokenValue = Replace$(TransformTokenValue, vbCr, " ")
    TransformTokenValue = Replace$(TransformTokenValue, vbLf, " ")
End Function

Public Function TransformVectorToView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal z As Double, _
    ByRef viewX As Double, _
    ByRef viewY As Double, _
    ByRef viewZ As Double) As Boolean

    On Error GoTo Failed

    Dim mathUtil As SldWorks.MathUtility
    Set mathUtil = swApp.GetMathUtility

    Dim vectorData(0 To 2) As Double
    vectorData(0) = x
    vectorData(1) = y
    vectorData(2) = z

    Dim mathVector As SldWorks.MathVector
    Set mathVector = mathUtil.CreateVector(vectorData)
    Set mathVector = mathVector.MultiplyTransform(swView.ModelToViewTransform)

    Dim result As Variant
    result = mathVector.ArrayData

    viewX = CDbl(result(0))
    viewY = CDbl(result(1))
    viewZ = CDbl(result(2))
    TransformVectorToView = True
    Exit Function

Failed:
    TransformVectorToView = False
End Function

' ViewToSheetCoordinates was removed in r20.  ModelToViewTransform already
' supplies drawing-page coordinates, so callers assign SheetX/SheetY directly
' from ViewX/ViewY.  Keeping an identity-valued helper here would silently
' mis-place any future caller that genuinely held view-local coordinates.

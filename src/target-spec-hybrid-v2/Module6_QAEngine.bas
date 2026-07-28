Option Explicit

Private Const swDrawingSectionView As Long = 2
Private Const swDrawingDetailView As Long = 3
Private Const swDatumTag As Long = 2
Private Const swDatumTargetSym As Long = 3
Private Const swDisplayDimension As Long = 4
Private Const swGTol As Long = 5
Private Const swNote As Long = 6
Private Const swSFSymbol As Long = 7
Private Const swWeldSymbol As Long = 8
Private Const swTableAnnotation As Long = 14
Private Const swOrdinateDimension As Long = 1
Private Const swHorOrdinateDimension As Long = 7
Private Const swVertOrdinateDimension As Long = 8
Private Const swAngularOrdinateDimension As Long = 16
Private Const BOUNDARY_TOLERANCE_M As Double = 0.000001

Public Sub PerformFinalDrawingChecks( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed
    RequireQaStages evidence

    Dim finalQaValid As Boolean
    Dim viewPolicyValid As Boolean
    Dim annotationExtentsValid As Boolean
    Dim modelAnnotationsValid As Boolean
    Dim physicalLocationsValid As Boolean
    Dim ordinateCoverageValid As Boolean

    finalQaValid = True
    viewPolicyValid = True
    annotationExtentsValid = True
    modelAnnotationsValid = True
    physicalLocationsValid = True
    ordinateCoverageValid = True

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        evidence.AddFailure "QA: current sheet is Nothing."
        evidence.MarkStageFailed "FINAL_QA", "current sheet is Nothing"
        evidence.MarkStageFailed "ANNOTATION_EXTENTS", _
            "current sheet is Nothing"
        evidence.MarkStageFailed "VIEW_DIMENSION_POLICY", _
            "current sheet is Nothing"
        Exit Sub
    End If

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim paperSize As Long
    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)

    If sheetWidth <= 0# Or sheetHeight <= 0# Then
        evidence.AddFailure "QA: current sheet dimensions are invalid."
        finalQaValid = False
        annotationExtentsValid = False
    End If

    Dim realViewCount As Long
    Dim isometricCount As Long
    Dim sectionCount As Long
    Dim detailCount As Long
    Dim ordinateEligibleCount As Long
    Dim currentDimensionCount As Long
    Dim currentOrdinateCount As Long

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        realViewCount = realViewCount + 1

        If swView.Type = swDrawingSectionView Then
            sectionCount = sectionCount + 1
        End If
        If swView.Type = swDrawingDetailView Then
            detailCount = detailCount + 1
        End If

        Dim ordinateEligible As Boolean
        ordinateEligible = _
            Module8_RuntimeSupport.IsOrdinateEligibleView(swView)
        If ordinateEligible Then
            ordinateEligibleCount = ordinateEligibleCount + 1
        End If

        Dim isometricView As Boolean
        isometricView = IsIsometricView(swView)
        If isometricView Then isometricCount = isometricCount + 1

        Dim viewDimensionCount As Long
        Dim viewOrdinateCount As Long
        viewDimensionCount = 0
        viewOrdinateCount = 0
        CountViewDimensions swView, viewDimensionCount, viewOrdinateCount

        currentDimensionCount = currentDimensionCount + viewDimensionCount
        currentOrdinateCount = currentOrdinateCount + viewOrdinateCount

        If isometricView And viewDimensionCount > 0 Then
            evidence.AddFailure "QA: isometric/axonometric view '" & _
                Module8_RuntimeSupport.GetViewName(swView) & _
                "' contains " & CStr(viewDimensionCount) & _
                " display dimension(s)."
            viewPolicyValid = False
            finalQaValid = False
        End If

        If Not ordinateEligible And viewOrdinateCount > 0 Then
            evidence.AddFailure "QA: unsupported view '" & _
                Module8_RuntimeSupport.GetViewName(swView) & _
                "' contains " & CStr(viewOrdinateCount) & _
                " ordinate display dimension(s)."
            viewPolicyValid = False
            finalQaValid = False
        End If

        If Not CheckViewAnnotations( _
            swView, sheetWidth, sheetHeight, evidence) Then

            annotationExtentsValid = False
            finalQaValid = False
        End If

        Set swView = swView.GetNextView
    Loop

    If Not CheckSectionLineClearance(swDraw, evidence) Then
        annotationExtentsValid = False
        finalQaValid = False
    End If

    If realViewCount = 0 Then
        evidence.AddFailure "QA: no model views exist."
        finalQaValid = False
        viewPolicyValid = False
    End If

    If Module1_Main.GlobalConfig.CreateIso And isometricCount = 0 Then
        evidence.AddFailure "QA: requested isometric view is missing."
        finalQaValid = False
        viewPolicyValid = False
    End If

    If currentDimensionCount = 0 Then
        evidence.AddFailure "QA: drawing contains zero visible display dimensions."
        finalQaValid = False
    End If

    If evidence.ImportedAnnotations = 0 Then
        evidence.AddFailure "QA: required model-annotation import returned zero."
        modelAnnotationsValid = False
        finalQaValid = False
    End If

    If evidence.CandidatesAccepted = 0 Then
        evidence.AddFailure "QA: no ownership-proven hole-location candidates were accepted."
        physicalLocationsValid = False
        finalQaValid = False
    End If

    If evidence.UniquePhysicalLocationCount = 0 Then
        evidence.AddFailure "QA: no canonical physical hole/location records were registered."
        physicalLocationsValid = False
        ordinateCoverageValid = False
        finalQaValid = False
    End If

    If evidence.UniquePhysicalLocationCount > 0 Then
        If evidence.ProjectionLocationRecordCount < _
           evidence.UniquePhysicalLocationCount Then

            evidence.AddFailure "QA: projected-location evidence covers fewer records " & _
                "than the canonical physical-location ledger."
            ordinateCoverageValid = False
            finalQaValid = False
        End If

        If evidence.HorizontalOrdinateGroups + _
           evidence.VerticalOrdinateGroups = 0 Then

            evidence.AddFailure "QA: owned physical locations exist but no " & _
                "ordinate groups were created."
            ordinateCoverageValid = False
            finalQaValid = False
        End If

        If evidence.OrdinateEntitiesSelected < _
           evidence.UniquePhysicalLocationCount Then

            evidence.AddFailure "QA: ordinate selections cover fewer entities " & _
                "than the unique physical-location ledger."
            ordinateCoverageValid = False
            finalQaValid = False
        End If
    End If

    If Module1_Main.GlobalConfig.PopulateTitle And _
       evidence.TitlePropertiesWritten = 0 Then

        evidence.AddFailure "QA: title population requested but no drawing " & _
            "properties were written."
        finalQaValid = False
    End If

    If Not evidence.LayoutBoundariesProven Then
        evidence.AddFailure "QA: border/title-block layout boundaries were not proven."
        finalQaValid = False
        annotationExtentsValid = False
    End If

    If Not evidence.SheetScaleReadbackProven Then
        evidence.AddFailure "QA: actual sheet scale was not read back and proved."
        finalQaValid = False
    End If

    If Not evidence.FinalCleanupVerified Or _
       evidence.FinalCleanupInvalidated Or _
       evidence.FinalSelectionCount <> 0 Or _
       Not evidence.SheetContextRestored Then

        evidence.AddFailure "QA: final selection/sheet context is not clean."
        finalQaValid = False
    End If

    If Not ApplyFixtureSpecificChecks( _
        swDraw, sectionCount, detailCount, ordinateEligibleCount, _
        isometricCount, evidence) Then

        finalQaValid = False
    End If

    If modelAnnotationsValid Then
        evidence.MarkStageProved "MODEL_ANNOTATIONS", _
            "nonzero imported annotation result recorded"
    Else
        evidence.MarkStageFailed "MODEL_ANNOTATIONS", _
            "required import returned zero"
    End If

    If physicalLocationsValid Then
        evidence.MarkStageProved "PHYSICAL_LOCATIONS", _
            "ownership candidates and canonical unique records are nonzero"
    Else
        evidence.MarkStageFailed "PHYSICAL_LOCATIONS", _
            "ownership or canonical unique record evidence is missing"
    End If

    If ordinateCoverageValid Then
        evidence.MarkStageProved "ORDINATE_COVERAGE", _
            "projection records, ordinate groups, and selections cover unique locations"
    Else
        evidence.MarkStageFailed "ORDINATE_COVERAGE", _
            "projection or ordinate coverage is incomplete"
    End If

    If viewPolicyValid Then
        evidence.MarkStageProved "VIEW_DIMENSION_POLICY", _
            "isometric views have no dimensions and unsupported views have no ordinates"
    Else
        evidence.MarkStageFailed "VIEW_DIMENSION_POLICY", _
            "unsupported-view dimension policy failed"
    End If

    If annotationExtentsValid Then
        evidence.MarkStageProved "ANNOTATION_EXTENTS", _
            "position-capable annotation origins, measurable note extents, and leader geometry passed"
    Else
        evidence.MarkStageFailed "ANNOTATION_EXTENTS", _
            "one or more annotation origins, measurable note extents, or leader paths were invalid"
    End If

    If finalQaValid Then
        evidence.MarkStageProved "FINAL_QA", _
            "read-only semantic and boundary checks passed"
    Else
        evidence.MarkStageFailed "FINAL_QA", _
            "one or more read-only final checks failed"
    End If

    evidence.AddWarning "Exact display-dimension glyph collision and manufacturing " & _
        "completeness still require visual comparison with the manual reference."
    evidence.AddInfo "Final drawing view count=" & CStr(realViewCount) & _
        ", isometric count=" & CStr(isometricCount) & _
        ", section count=" & CStr(sectionCount) & _
        ", eligible orthographic count=" & CStr(ordinateEligibleCount) & _
        ", display dimension count=" & CStr(currentDimensionCount) & _
        ", ordinate dimension count=" & CStr(currentOrdinateCount) & "."
    Exit Sub

Failed:
    evidence.AddFailure "Final QA error " & CStr(Err.Number) & _
        ": " & Err.Description
    evidence.MarkStageFailed "FINAL_QA", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Private Sub RequireQaStages(ByRef evidence As CRunEvidence)
    evidence.RequireStage "MODEL_ANNOTATIONS"
    evidence.RequireStage "PHYSICAL_LOCATIONS"
    evidence.RequireStage "ORDINATE_COVERAGE"
    evidence.RequireStage "VIEW_DIMENSION_POLICY"
    evidence.RequireStage "ANNOTATION_EXTENTS"
    evidence.RequireStage "FINAL_QA"
End Sub

Private Sub CountViewDimensions( _
    ByRef swView As SldWorks.View, _
    ByRef dimensionCount As Long, _
    ByRef ordinateCount As Long)

    On Error GoTo Failed
    dimensionCount = 0
    ordinateCount = 0

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If IsEmpty(dimensions) Or Not IsArray(dimensions) Then Exit Sub

    Dim i As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)

        If Not displayDimension Is Nothing Then
            dimensionCount = dimensionCount + 1
            If IsOrdinateDimensionType(displayDimension.Type2) Then
                ordinateCount = ordinateCount + 1
            End If
        End If
    Next i
    Exit Sub

Failed:
    Err.Raise Err.Number, "CountViewDimensions", Err.Description
End Sub

Private Function IsOrdinateDimensionType(ByVal dimensionType As Long) As Boolean
    Select Case dimensionType
        Case swOrdinateDimension, swHorOrdinateDimension, _
             swVertOrdinateDimension, swAngularOrdinateDimension

            IsOrdinateDimensionType = True
    End Select
End Function

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

Private Function CheckViewAnnotations( _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed
    CheckViewAnnotations = True

    Dim annotations As Variant
    annotations = swView.GetAnnotations

    If IsArray(annotations) Then
        Dim i As Long
        For i = LBound(annotations) To UBound(annotations)
            Dim annotation As SldWorks.Annotation
            Set annotation = annotations(i)

            If Not annotation Is Nothing Then
                If AnnotationSupportsPosition(annotation.GetType) Then
                    If Not CheckAnnotationGeometry( _
                        annotation, swView, sheetWidth, sheetHeight, _
                        evidence) Then CheckViewAnnotations = False
                End If
            End If
        Next i
    End If

    Dim note As SldWorks.Note
    Set note = swView.GetFirstNote

    Do While Not note Is Nothing
        If Not CheckModelViewNoteExtent( _
            note, swView, sheetWidth, sheetHeight, evidence) Then

            CheckViewAnnotations = False
        End If
        Set note = note.GetNext
    Loop
    Exit Function

Failed:
    evidence.AddFailure "Annotation-boundary QA could not inspect '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
    CheckViewAnnotations = False
End Function

Private Function AnnotationSupportsPosition( _
    ByVal annotationType As Long) As Boolean

    Select Case annotationType
        Case swDatumTag, swDatumTargetSym, swDisplayDimension, swGTol, _
             swNote, swSFSymbol, swWeldSymbol, swTableAnnotation
            AnnotationSupportsPosition = True
    End Select
End Function

Private Function AnnotationTypeLabel(ByVal annotationType As Long) As String
    Select Case annotationType
        Case swDatumTag: AnnotationTypeLabel = "DatumFeature"
        Case swDatumTargetSym: AnnotationTypeLabel = "DatumTarget"
        Case swDisplayDimension: AnnotationTypeLabel = "DisplayDimension"
        Case swGTol: AnnotationTypeLabel = "GeometricTolerance"
        Case swNote: AnnotationTypeLabel = "Note"
        Case swSFSymbol: AnnotationTypeLabel = "SurfaceFinish"
        Case swWeldSymbol: AnnotationTypeLabel = "WeldSymbol"
        Case swTableAnnotation: AnnotationTypeLabel = "Table"
        Case Else: AnnotationTypeLabel = "Type" & CStr(annotationType)
    End Select
End Function

Private Function SafeAnnotationName( _
    ByRef annotation As SldWorks.Annotation) As String

    If annotation Is Nothing Then
        SafeAnnotationName = "(nothing)"
        Exit Function
    End If

    On Error GoTo Unavailable
    SafeAnnotationName = Trim$(annotation.GetName)
    If Len(SafeAnnotationName) = 0 Then SafeAnnotationName = "(unnamed)"
    Exit Function

Unavailable:
    SafeAnnotationName = "(name-unavailable)"
End Function

Private Function EvidenceValue(ByVal value As String) As String
    EvidenceValue = Replace$(Trim$(value), "|", "/")
    EvidenceValue = Replace$(EvidenceValue, vbCr, " ")
    EvidenceValue = Replace$(EvidenceValue, vbLf, " ")
End Function

Private Function CheckAnnotationGeometry( _
    ByRef annotation As SldWorks.Annotation, _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed
    CheckAnnotationGeometry = True

    Dim position As Variant
    position = annotation.GetPosition
    Dim annotationType As Long
    Dim annotationName As String
    annotationType = annotation.GetType
    annotationName = SafeAnnotationName(annotation)

    If Not IsArray(position) Then
        evidence.AddFailure "Annotation position is unavailable in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        CheckAnnotationGeometry = False
    ElseIf UBound(position) - LBound(position) + 1 < 2 Then
        evidence.AddFailure "Annotation position is incomplete in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        CheckAnnotationGeometry = False
    Else
        Dim positionIndex As Long
        Dim positionX As Double
        Dim positionY As Double
        Dim violationReason As String
        positionIndex = LBound(position)
        positionX = CDbl(position(positionIndex))
        positionY = CDbl(position(positionIndex + 1))

        evidence.AddInfo "ANNOTATION_GEOMETRY|view=" & _
            EvidenceValue(Module8_RuntimeSupport.GetViewName(swView)) & _
            "|type=" & EvidenceValue(AnnotationTypeLabel(annotationType)) & _
            "|name=" & EvidenceValue(annotationName) & _
            "|x=" & Format$(positionX, "0.000000") & _
            "|y=" & Format$(positionY, "0.000000")

        If PointViolatesControlledRegions( _
            positionX, positionY, sheetWidth, sheetHeight, _
            evidence, violationReason) Then

            evidence.AddFailure "Annotation origin violation in '" & _
                Module8_RuntimeSupport.GetViewName(swView) & _
                "': type=" & AnnotationTypeLabel(annotationType) & _
                ", name=" & annotationName & _
                ", x=" & Format$(positionX, "0.000000") & _
                ", y=" & Format$(positionY, "0.000000") & _
                ", region=" & violationReason & "."
            CheckAnnotationGeometry = False
        End If
    End If

    Dim leaderCount As Long
    leaderCount = annotation.GetLeaderCount
    If leaderCount < 0 Then
        evidence.AddFailure "Annotation leader count is invalid in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        CheckAnnotationGeometry = False
        Exit Function
    End If

    Dim leaderIndex As Long
    For leaderIndex = 0 To leaderCount - 1
        Dim leaderPoints As Variant
        leaderPoints = annotation.GetLeaderPointsAtIndex(leaderIndex)

        If Not CheckLeaderPoints( _
            leaderPoints, swView, sheetWidth, sheetHeight, evidence) Then

            CheckAnnotationGeometry = False
        End If
    Next leaderIndex
    Exit Function

Failed:
    evidence.AddFailure "Annotation geometry API error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
    CheckAnnotationGeometry = False
End Function

Private Function CheckLeaderPoints( _
    ByVal leaderPoints As Variant, _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    CheckLeaderPoints = True

    If Not IsArray(leaderPoints) Then
        evidence.AddFailure "Leader points are unavailable in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        CheckLeaderPoints = False
        Exit Function
    End If

    Dim pointValueCount As Long
    pointValueCount = UBound(leaderPoints) - LBound(leaderPoints) + 1

    If pointValueCount < 6 Or pointValueCount Mod 3 <> 0 Then
        evidence.AddFailure "Leader point array is incomplete in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        CheckLeaderPoints = False
        Exit Function
    End If

    Dim firstIndex As Long
    firstIndex = LBound(leaderPoints)

    Dim pointIndex As Long
    Dim previousX As Double
    Dim previousY As Double
    Dim hasPrevious As Boolean
    Dim violationReason As String

    For pointIndex = firstIndex To UBound(leaderPoints) Step 3
        Dim currentX As Double
        Dim currentY As Double
        currentX = CDbl(leaderPoints(pointIndex))
        currentY = CDbl(leaderPoints(pointIndex + 1))

        If PointViolatesControlledRegions( _
            currentX, currentY, sheetWidth, sheetHeight, _
            evidence, violationReason) Then

            evidence.AddFailure "Leader point violation in '" & _
                Module8_RuntimeSupport.GetViewName(swView) & _
                "': x=" & Format$(currentX, "0.000000") & _
                ", y=" & Format$(currentY, "0.000000") & _
                ", region=" & violationReason & "."
            CheckLeaderPoints = False
        End If

        If hasPrevious Then
            If SegmentIntersectsRectangle( _
                previousX, previousY, currentX, currentY, _
                evidence.TitleBlockLeft, evidence.TitleBlockBottom, _
                evidence.TitleBlockRight, evidence.TitleBlockTop) Then

                evidence.AddFailure "Leader segment crosses the controlled " & _
                    "title block in '" & _
                    Module8_RuntimeSupport.GetViewName(swView) & "'."
                CheckLeaderPoints = False
            End If
        End If

        previousX = currentX
        previousY = currentY
        hasPrevious = True
    Next pointIndex
End Function

Private Function CheckModelViewNoteExtent( _
    ByRef note As SldWorks.Note, _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim renderedText As String
    If Not TryReadRenderedNoteText(note, renderedText) Then
        evidence.AddFailure "Mandatory note text could not be read in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    If Len(renderedText) = 0 Then
        evidence.AddInfo "NON_RENDERED_NOTE_SKIPPED|context=FinalQA|view=" & _
            Module8_RuntimeSupport.GetViewName(swView)
        CheckModelViewNoteExtent = True
        Exit Function
    End If

    Dim extent As Variant
    extent = note.GetExtent

    If Not IsArray(extent) Then
        If Module1_Main.DIAGNOSTIC_DRAWING_MODE And _
           Not evidence.LayoutBoundariesProven Then

            evidence.AddWarning "NOTE_EXTENT_SKIPPED|context=FinalQA" & _
                "|view=" & Module8_RuntimeSupport.GetViewName(swView) & _
                "|reason=Unavailable|acceptance=False"
            CheckModelViewNoteExtent = True
            Exit Function
        End If

        evidence.AddFailure "Mandatory note extent is unavailable in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    If UBound(extent) - LBound(extent) + 1 < 6 Then
        If Module1_Main.DIAGNOSTIC_DRAWING_MODE And _
           Not evidence.LayoutBoundariesProven Then

            evidence.AddWarning "NOTE_EXTENT_SKIPPED|context=FinalQA" & _
                "|view=" & Module8_RuntimeSupport.GetViewName(swView) & _
                "|reason=Incomplete|acceptance=False"
            CheckModelViewNoteExtent = True
            Exit Function
        End If

        evidence.AddFailure "Mandatory note extent is incomplete in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    Dim index As Long
    index = LBound(extent)

    Dim extentLeft As Double
    Dim extentBottom As Double
    Dim extentRight As Double
    Dim extentTop As Double
    extentLeft = CDbl(extent(index))
    extentBottom = CDbl(extent(index + 1))
    extentRight = CDbl(extent(index + 3))
    extentTop = CDbl(extent(index + 4))

    If extentRight < extentLeft Then SwapDouble extentLeft, extentRight
    If extentTop < extentBottom Then SwapDouble extentBottom, extentTop

    If extentRight <= extentLeft Or extentTop <= extentBottom Then
        If Module1_Main.DIAGNOSTIC_DRAWING_MODE And _
           Not evidence.LayoutBoundariesProven Then

            evidence.AddWarning "NOTE_EXTENT_SKIPPED|context=FinalQA" & _
                "|view=" & Module8_RuntimeSupport.GetViewName(swView) & _
                "|reason=Invalid|acceptance=False"
            CheckModelViewNoteExtent = True
            Exit Function
        End If

        evidence.AddFailure "Mandatory note extent is invalid in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    If extentLeft < evidence.ContentBorderLeft - BOUNDARY_TOLERANCE_M Or _
       extentBottom < evidence.ContentBorderBottom - BOUNDARY_TOLERANCE_M Or _
       extentRight > evidence.ContentBorderRight + BOUNDARY_TOLERANCE_M Or _
       extentTop > evidence.ContentBorderTop + BOUNDARY_TOLERANCE_M Then

        evidence.AddFailure "Note extent crosses the measured zoned border in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "': " & _
            Format$(extentLeft, "0.000000") & "," & _
            Format$(extentBottom, "0.000000") & " to " & _
            Format$(extentRight, "0.000000") & "," & _
            Format$(extentTop, "0.000000") & "."
        Exit Function
    End If

    If RectanglesOverlap( _
            extentLeft, extentBottom, extentRight, extentTop, _
            evidence.TitleBlockLeft, evidence.TitleBlockBottom, _
            evidence.TitleBlockRight, evidence.TitleBlockTop) Then

        evidence.AddFailure "Note extent intrudes into the measured title-block " & _
            "rectangle in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    If evidence.PartIdentificationBoundsProven And _
       InStr(1, renderedText, _
             "*" & Module1_Main.GetFixtureKey(evidence.PartPath) & "*", _
             vbTextCompare) = 0 Then

        If RectanglesOverlap( _
                extentLeft, extentBottom, extentRight, extentTop, _
                evidence.PartIdentificationLeft, _
                evidence.PartIdentificationBottom, _
                evidence.PartIdentificationRight, _
                evidence.PartIdentificationTop) Then

            evidence.AddFailure "Note extent intrudes into the measured " & _
                "part-identification rectangle in '" & _
                Module8_RuntimeSupport.GetViewName(swView) & "'."
            Exit Function
        End If
    End If

    CheckModelViewNoteExtent = True
    Exit Function

Failed:
    evidence.AddFailure "Note extent API error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
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

Private Function PointViolatesControlledRegions( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef violationReason As String) As Boolean

    violationReason = vbNullString

    If x < 0# Or x > sheetWidth Or y < 0# Or y > sheetHeight Then
        violationReason = "PhysicalSheet"
        PointViolatesControlledRegions = True
        Exit Function
    End If

    If x < evidence.ContentBorderLeft - BOUNDARY_TOLERANCE_M Or _
       x > evidence.ContentBorderRight + BOUNDARY_TOLERANCE_M Or _
       y < evidence.ContentBorderBottom - BOUNDARY_TOLERANCE_M Or _
       y > evidence.ContentBorderTop + BOUNDARY_TOLERANCE_M Then

        violationReason = "ZonedBorder"
        PointViolatesControlledRegions = True
        Exit Function
    End If

    PointViolatesControlledRegions = _
        PointInsideRectangle( _
            x, y, evidence.TitleBlockLeft, evidence.TitleBlockBottom, _
            evidence.TitleBlockRight, evidence.TitleBlockTop)
    If PointViolatesControlledRegions Then
        violationReason = "TitleBlock"
    End If
End Function

Private Function PointInsideRectangle( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal leftValue As Double, _
    ByVal bottomValue As Double, _
    ByVal rightValue As Double, _
    ByVal topValue As Double) As Boolean

    PointInsideRectangle = _
        (x >= leftValue) And (x <= rightValue) And _
        (y >= bottomValue) And (y <= topValue)
End Function

Private Function RectanglesOverlap( _
    ByVal firstLeft As Double, _
    ByVal firstBottom As Double, _
    ByVal firstRight As Double, _
    ByVal firstTop As Double, _
    ByVal secondLeft As Double, _
    ByVal secondBottom As Double, _
    ByVal secondRight As Double, _
    ByVal secondTop As Double) As Boolean

    RectanglesOverlap = Not ( _
        firstRight <= secondLeft Or secondRight <= firstLeft Or _
        firstTop <= secondBottom Or secondTop <= firstBottom)
End Function

Private Function SegmentIntersectsRectangle( _
    ByVal x1 As Double, _
    ByVal y1 As Double, _
    ByVal x2 As Double, _
    ByVal y2 As Double, _
    ByVal leftValue As Double, _
    ByVal bottomValue As Double, _
    ByVal rightValue As Double, _
    ByVal topValue As Double) As Boolean

    If PointInsideRectangle( _
        x1, y1, leftValue, bottomValue, rightValue, topValue) Or _
       PointInsideRectangle( _
        x2, y2, leftValue, bottomValue, rightValue, topValue) Then

        SegmentIntersectsRectangle = True
        Exit Function
    End If

    SegmentIntersectsRectangle = _
        SegmentsIntersect(x1, y1, x2, y2, _
            leftValue, bottomValue, rightValue, bottomValue) Or _
        SegmentsIntersect(x1, y1, x2, y2, _
            rightValue, bottomValue, rightValue, topValue) Or _
        SegmentsIntersect(x1, y1, x2, y2, _
            rightValue, topValue, leftValue, topValue) Or _
        SegmentsIntersect(x1, y1, x2, y2, _
            leftValue, topValue, leftValue, bottomValue)
End Function

Private Function SegmentsIntersect( _
    ByVal ax As Double, _
    ByVal ay As Double, _
    ByVal bx As Double, _
    ByVal by As Double, _
    ByVal cx As Double, _
    ByVal cy As Double, _
    ByVal dx As Double, _
    ByVal dy As Double) As Boolean

    Dim firstSide As Double
    Dim secondSide As Double
    Dim thirdSide As Double
    Dim fourthSide As Double

    firstSide = CrossProduct(ax, ay, bx, by, cx, cy)
    secondSide = CrossProduct(ax, ay, bx, by, dx, dy)
    thirdSide = CrossProduct(cx, cy, dx, dy, ax, ay)
    fourthSide = CrossProduct(cx, cy, dx, dy, bx, by)

    Const tolerance As Double = 0.000000000001
    SegmentsIntersect = _
        ((firstSide > tolerance And secondSide < -tolerance) Or _
         (firstSide < -tolerance And secondSide > tolerance)) And _
        ((thirdSide > tolerance And fourthSide < -tolerance) Or _
         (thirdSide < -tolerance And fourthSide > tolerance))
End Function

Private Function CrossProduct( _
    ByVal ax As Double, _
    ByVal ay As Double, _
    ByVal bx As Double, _
    ByVal by As Double, _
    ByVal px As Double, _
    ByVal py As Double) As Double

    CrossProduct = (bx - ax) * (py - ay) - _
        (by - ay) * (px - ax)
End Function

Private Function CheckSectionLineClearance( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence) As Boolean

    CheckSectionLineClearance = True
    If Module1_Main.GetFixtureKey(evidence.PartPath) <> _
       "P-0251-14A-001" Then Exit Function

    On Error GoTo Failed
    evidence.RequireStage "SECTION_CLEARANCE"

    If Not evidence.PartIdentificationBoundsProven Then
        evidence.AddFailure "P-0251 section clearance cannot be checked because " & _
            "the part-identification note extent is unproved."
        GoTo ClearanceFailed
    End If

    Dim sectionLineCount As Long
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim viewSectionLineCount As Long
        Dim sectionLineInfoSize As Long
        viewSectionLineCount = _
            swView.GetSectionLineCount2(sectionLineInfoSize)

        If viewSectionLineCount > 0 Then
            If sectionLineInfoSize < 1 Then
                evidence.AddFailure "P-0251 section-line count is nonzero but " & _
                    "IView.GetSectionLineCount2 returned an invalid data size " & _
                    "for '" & Module8_RuntimeSupport.GetViewName(swView) & "'."
                GoTo ClearanceFailed
            End If

            Dim sectionInfo As Variant
            sectionInfo = swView.GetSectionLineInfo2

            If IsArray(sectionInfo) Then
                Dim sectionInfoItemCount As Long
                sectionInfoItemCount = _
                    Module8_RuntimeSupport.CountVariantItems(sectionInfo)

                If sectionInfoItemCount <> sectionLineInfoSize Then
                    evidence.AddFailure "P-0251 section-line data-size mismatch " & _
                        "in '" & Module8_RuntimeSupport.GetViewName(swView) & _
                        "': expected=" & CStr(sectionLineInfoSize) & _
                        ", actual=" & CStr(sectionInfoItemCount) & "."
                    GoTo ClearanceFailed
                End If

                evidence.AddInfo "SECTION_LINE_READBACK|view=" & _
                    EvidenceValue( _
                        Module8_RuntimeSupport.GetViewName(swView)) & _
                    "|count=" & CStr(viewSectionLineCount) & _
                    "|size=" & CStr(sectionLineInfoSize)

                If Not ValidateSectionLineInfo( _
                    sectionInfo, swView, evidence, sectionLineCount) Then

                    GoTo ClearanceFailed
                End If
            Else
                evidence.AddFailure "P-0251 section-line count is nonzero but " & _
                    "IView.GetSectionLineInfo2 returned no array in '" & _
                    Module8_RuntimeSupport.GetViewName(swView) & "'."
                GoTo ClearanceFailed
            End If
        End If

        Set swView = swView.GetNextView
    Loop

    If sectionLineCount = 0 Then
        evidence.AddFailure "P-0251 J-J section-line geometry readback is empty."
        GoTo ClearanceFailed
    End If

    evidence.MarkStageProved "SECTION_CLEARANCE", _
        CStr(sectionLineCount) & _
        " section line(s) proved clear of the measured part-identification extent"
    Exit Function

ClearanceFailed:
    evidence.MarkStageFailed "SECTION_CLEARANCE", _
        "section segment, arrow, or label geometry was unavailable or intersected the part-identification extent"
    CheckSectionLineClearance = False
    Exit Function

Failed:
    evidence.AddFailure "P-0251 section-line clearance API error " & _
        CStr(Err.Number) & ": " & Err.Description
    Resume ClearanceFailed
End Function

Private Function ValidateSectionLineInfo( _
    ByVal sectionInfo As Variant, _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence, _
    ByRef accumulatedLineCount As Long) As Boolean

    On Error GoTo Failed

    Dim cursor As Long
    Dim upperIndex As Long
    cursor = LBound(sectionInfo)
    upperIndex = UBound(sectionInfo)
    If cursor > upperIndex Then Exit Function

    Dim lineCount As Long
    lineCount = CLng(sectionInfo(cursor))
    cursor = cursor + 1
    If lineCount < 0 Then Exit Function
    If lineCount = 0 Then
        ValidateSectionLineInfo = True
        Exit Function
    End If

    Dim lineIndex As Long
    For lineIndex = 1 To lineCount
        If cursor + 1 > upperIndex Then Exit Function

        Dim layerValue As Double
        Dim segmentCount As Long
        layerValue = CDbl(sectionInfo(cursor))
        cursor = cursor + 1
        segmentCount = CLng(sectionInfo(cursor))
        cursor = cursor + 1
        If segmentCount < 1 Then Exit Function

        Dim segmentIndex As Long
        For segmentIndex = 1 To segmentCount
            If cursor + 6 > upperIndex Then Exit Function

            Dim startX As Double
            Dim startY As Double
            Dim endX As Double
            Dim endY As Double
            startX = CDbl(sectionInfo(cursor + 1))
            startY = CDbl(sectionInfo(cursor + 2))
            endX = CDbl(sectionInfo(cursor + 4))
            endY = CDbl(sectionInfo(cursor + 5))

            If SectionSegmentTouchesPartIdentification( _
                startX, startY, endX, endY, evidence) Then

                evidence.AddFailure "P-0251 J-J section segment intersects " & _
                    "the measured part-identification extent in '" & _
                    Module8_RuntimeSupport.GetViewName(swView) & "'."
                Exit Function
            End If
            cursor = cursor + 7
        Next segmentIndex

        If cursor + 24 > upperIndex Then Exit Function

        Dim arrow1StartX As Double
        Dim arrow1StartY As Double
        Dim arrow1EndX As Double
        Dim arrow1EndY As Double
        arrow1StartX = CDbl(sectionInfo(cursor))
        arrow1StartY = CDbl(sectionInfo(cursor + 1))
        arrow1EndX = CDbl(sectionInfo(cursor + 3))
        arrow1EndY = CDbl(sectionInfo(cursor + 4))
        cursor = cursor + 9

        Dim arrow2StartX As Double
        Dim arrow2StartY As Double
        Dim arrow2EndX As Double
        Dim arrow2EndY As Double
        arrow2StartX = CDbl(sectionInfo(cursor))
        arrow2StartY = CDbl(sectionInfo(cursor + 1))
        arrow2EndX = CDbl(sectionInfo(cursor + 3))
        arrow2EndY = CDbl(sectionInfo(cursor + 4))
        cursor = cursor + 9

        If SectionSegmentTouchesPartIdentification( _
               arrow1StartX, arrow1StartY, arrow1EndX, arrow1EndY, evidence) Or _
           SectionSegmentTouchesPartIdentification( _
               arrow2StartX, arrow2StartY, arrow2EndX, arrow2EndY, evidence) Then

            evidence.AddFailure "P-0251 J-J arrow geometry intersects the " & _
                "measured part-identification extent in '" & _
                Module8_RuntimeSupport.GetViewName(swView) & "'."
            Exit Function
        End If

        Dim text1X As Double
        Dim text1Y As Double
        Dim text2X As Double
        Dim text2Y As Double
        Dim textHeight As Double
        text1X = CDbl(sectionInfo(cursor))
        text1Y = CDbl(sectionInfo(cursor + 1))
        text2X = CDbl(sectionInfo(cursor + 3))
        text2Y = CDbl(sectionInfo(cursor + 4))
        textHeight = Abs(CDbl(sectionInfo(cursor + 6)))
        cursor = cursor + 7

        Dim textClearance As Double
        textClearance = textHeight * 2#
        If textClearance < 0.002 Then textClearance = 0.002

        If PointTouchesPartIdentification( _
               text1X, text1Y, textClearance, evidence) Or _
           PointTouchesPartIdentification( _
               text2X, text2Y, textClearance, evidence) Then

            evidence.AddFailure "P-0251 J-J label position intersects the " & _
                "measured part-identification extent in '" & _
                Module8_RuntimeSupport.GetViewName(swView) & "'."
            Exit Function
        End If

        evidence.AddInfo "SECTION_LINE_GEOMETRY|view=" & _
            EvidenceValue(Module8_RuntimeSupport.GetViewName(swView)) & _
            "|line=" & CStr(lineIndex) & _
            "|layer=" & Format$(layerValue, "0.###") & _
            "|segments=" & CStr(segmentCount) & _
            "|text1=" & Format$(text1X, "0.000000") & "," & _
                Format$(text1Y, "0.000000") & _
            "|text2=" & Format$(text2X, "0.000000") & "," & _
                Format$(text2Y, "0.000000") & _
            "|textHeight=" & Format$(textHeight, "0.000000")
    Next lineIndex

    accumulatedLineCount = accumulatedLineCount + lineCount
    ValidateSectionLineInfo = True
    Exit Function

Failed:
    ValidateSectionLineInfo = False
End Function

Private Function SectionSegmentTouchesPartIdentification( _
    ByVal startX As Double, _
    ByVal startY As Double, _
    ByVal endX As Double, _
    ByVal endY As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    Const clearance As Double = 0.002
    SectionSegmentTouchesPartIdentification = SegmentIntersectsRectangle( _
        startX, startY, endX, endY, _
        evidence.PartIdentificationLeft - clearance, _
        evidence.PartIdentificationBottom - clearance, _
        evidence.PartIdentificationRight + clearance, _
        evidence.PartIdentificationTop + clearance)
End Function

Private Function PointTouchesPartIdentification( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal clearance As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    PointTouchesPartIdentification = _
        (x >= evidence.PartIdentificationLeft - clearance) And _
        (x <= evidence.PartIdentificationRight + clearance) And _
        (y >= evidence.PartIdentificationBottom - clearance) And _
        (y <= evidence.PartIdentificationTop + clearance)
End Function

Private Sub SwapDouble(ByRef firstValue As Double, ByRef secondValue As Double)
    Dim temporaryValue As Double
    temporaryValue = firstValue
    firstValue = secondValue
    secondValue = temporaryValue
End Sub

Private Function ApplyFixtureSpecificChecks( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal sectionCount As Long, _
    ByVal detailCount As Long, _
    ByVal ordinateEligibleCount As Long, _
    ByVal isometricCount As Long, _
    ByRef evidence As CRunEvidence) As Boolean

    ApplyFixtureSpecificChecks = True

    Dim fixtureName As String
    fixtureName = UCase$(FileNameWithoutExtension(evidence.PartPath))

    If ordinateEligibleCount < 2 Then
        evidence.AddFailure "Fixture QA: fewer than two supported orthographic views."
        ApplyFixtureSpecificChecks = False
    End If

    If isometricCount < 1 Then
        evidence.AddFailure "Fixture QA: required undimensioned isometric view is missing."
        ApplyFixtureSpecificChecks = False
    End If

    Select Case fixtureName
        Case "P-0251-14A-001"
            If sectionCount <> 1 Then
                evidence.AddFailure "P-0251 acceptance requires exactly one J-J section."
                ApplyFixtureSpecificChecks = False
            End If

            If Not evidence.StageIsProved("MANUFACTURING_DEFINITION") Or _
               Not P0251ManufacturingDefinitionIsVisible(swDraw) Then

                evidence.AddFailure "P-0251 acceptance requires the visible " & _
                    "stepped-bore, six-hole counterbore, and four-hole tapped definitions."
                ApplyFixtureSpecificChecks = False
            End If

            If evidence.UniquePhysicalLocationCount < 10 Then
                evidence.AddFailure "P-0251 canonical evidence contains fewer than " & _
                    "the ten reference hole locations."
                ApplyFixtureSpecificChecks = False
            End If

        Case "P-0252-01-001"
            If sectionCount > 0 Then
                evidence.AddWarning "P-0252-01-001 reference has no default section; " & _
                    "the configured section requires explicit review."
            End If

        Case "P-0252-01-013"
            If sectionCount <> 1 Then
                evidence.AddFailure "P-0252-01-013 acceptance requires exactly one B-B section."
                ApplyFixtureSpecificChecks = False
            End If

            If detailCount <> 2 Then
                evidence.AddFailure _
                    "P-0252-01-013 acceptance requires Details C and D."
                ApplyFixtureSpecificChecks = False
            End If

            If CountAnnotationsOfType(swDraw, swDatumTag) = 0 Then
                evidence.AddFailure "P-0252-01-013 has no visible imported datum feature symbol."
                ApplyFixtureSpecificChecks = False
            End If
    End Select
End Function

Private Function P0251ManufacturingDefinitionIsVisible( _
    ByRef swDraw As SldWorks.DrawingDoc) As Boolean

    P0251ManufacturingDefinitionIsVisible = _
        AssociatedCalloutContainsTokens( _
            swDraw, Array( _
                "DIAMETER:47 H7 (+0.025/+0.000)", _
                "DIAMETER:40")) And _
        AssociatedCalloutContainsTokens( _
            swDraw, Array( _
                "6X DIAMETER:6.6 THRU", _
                "C'BORE DIAMETER:11 x 6 DEEP")) And _
        AssociatedCalloutContainsTokens( _
            swDraw, Array( _
                "4X DIAMETER:4.2 x 12.4 DEEP", _
                "TAP M5x0.8-6H x 10 DEEP"))
End Function

Private Function AssociatedCalloutContainsTokens( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal requiredTokens As Variant) As Boolean

    On Error GoTo Failed

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing
        Dim note As SldWorks.Note
        Set note = swView.GetFirstNote

        Do While Not note Is Nothing
            Dim renderedText As String
            renderedText = NormalizeDiameterSyntax(note.GetText)

            Dim allTokensPresent As Boolean
            allTokensPresent = True

            Dim tokenIndex As Long
            For tokenIndex = LBound(requiredTokens) To UBound(requiredTokens)
                If InStr(1, renderedText, CStr(requiredTokens(tokenIndex)), _
                    vbTextCompare) = 0 Then

                    allTokensPresent = False
                    Exit For
                End If
            Next tokenIndex

            If allTokensPresent Then
                Dim annotation As SldWorks.Annotation
                Set annotation = note.GetAnnotation

                If Not annotation Is Nothing Then
                    Dim attached As Variant
                    attached = annotation.GetAttachedEntities3

                    If annotation.GetLeaderCount > 0 And IsArray(attached) Then
                        If Module8_RuntimeSupport.CountVariantItems(attached) > 0 Then
                            AssociatedCalloutContainsTokens = True
                            Exit Function
                        End If
                    End If
                End If
            End If

            Set note = note.GetNext
        Loop

        Set swView = swView.GetNextView
    Loop
    Exit Function

Failed:
    AssociatedCalloutContainsTokens = False
End Function

Private Function NormalizeDiameterSyntax(ByVal text As String) As String
    NormalizeDiameterSyntax = Replace$( _
        text, "<MOD-DIAM>", "DIAMETER:", 1, -1, vbTextCompare)
    NormalizeDiameterSyntax = Replace$( _
        NormalizeDiameterSyntax, ChrW$(&H2300), "DIAMETER:")
    NormalizeDiameterSyntax = Replace$( _
        NormalizeDiameterSyntax, ChrW$(&HD8), "DIAMETER:")
    NormalizeDiameterSyntax = Replace$(NormalizeDiameterSyntax, vbCr, vbNullString)
    NormalizeDiameterSyntax = Replace$(NormalizeDiameterSyntax, vbLf, " ")
    NormalizeDiameterSyntax = Trim$(NormalizeDiameterSyntax)
End Function

Private Function CountAnnotationsOfType( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal annotationType As Long) As Long

    On Error GoTo Failed

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing
        Dim annotations As Variant
        annotations = swView.GetAnnotations

        If IsArray(annotations) Then
            Dim i As Long
            For i = LBound(annotations) To UBound(annotations)
                Dim annotation As SldWorks.Annotation
                Set annotation = annotations(i)

                If Not annotation Is Nothing Then
                    If annotation.GetType = annotationType Then
                        CountAnnotationsOfType = CountAnnotationsOfType + 1
                    End If
                End If
            Next i
        End If

        Set swView = swView.GetNextView
    Loop
    Exit Function

Failed:
    CountAnnotationsOfType = 0
End Function

Public Sub EmitRunEvidence(ByRef evidence As CRunEvidence)
    RequireCoreStages evidence

    Dim summary As String
    Dim writeSucceeded As Boolean
    writeSucceeded = WriteEvidenceFileAtomically(evidence, summary)

    Debug.Print summary

    If Module1_Main.DIAGNOSTIC_DRAWING_MODE And _
       evidence.HasFailures And evidence.ViewsCreated > 0 Then

        MsgBox "Diagnostic drawing generation completed with " & _
            CStr(evidence.ViewsCreated) & " created view(s)." & vbCrLf & vbCrLf & _
            "Production QA remains FAIL. Inspect the drawing and send the " & _
            "Immediate Window log or QA report for the remaining feature-level " & _
            "diagnosis." & vbCrLf & vbCrLf & _
            evidence.EvidenceReportPath, vbExclamation, _
            "Target-Spec Hybrid V2 - Diagnostic Output Ready"
    ElseIf evidence.HasFailures Or Not writeSucceeded Then
        MsgBox summary, vbCritical, _
            "Target-Spec Hybrid V2 - Failed Closed"
    Else
        MsgBox summary, vbInformation, _
            "Target-Spec Hybrid V2 - Review Required"
    End If
End Sub

Private Sub RequireCoreStages(ByRef evidence As CRunEvidence)
    evidence.RequireStage "CONTROLLED_SHEET"
    evidence.RequireStage "SHEET_SCALE"
    evidence.RequireStage "LAYOUT"
    evidence.RequireStage "MODEL_ANNOTATIONS"
    evidence.RequireStage "PHYSICAL_LOCATIONS"
    evidence.RequireStage "ORDINATE_COVERAGE"
    evidence.RequireStage "TITLE_PROPERTIES"
    evidence.RequireStage "GENERAL_NOTES"
    evidence.RequireStage "PART_IDENTIFICATION"
    evidence.RequireStage "VIEW_DIMENSION_POLICY"
    evidence.RequireStage "ANNOTATION_EXTENTS"
    evidence.RequireStage "FINAL_QA"
    evidence.RequireStage "FINAL_CLEANUP"
    evidence.RequireStage "EVIDENCE_WRITE"
End Sub

Private Function WriteEvidenceFileAtomically( _
    ByRef evidence As CRunEvidence, _
    ByRef summary As String) As Boolean

    On Error GoTo Failed

    evidence.EvidenceWriteAttempted = True

    Dim reportFolder As String
    reportFolder = CreateUniqueReportFolder(evidence)
    If Len(reportFolder) = 0 Then
        Err.Raise vbObjectError + 1601, _
            "WriteEvidenceFileAtomically", _
            "Could not resolve or create the controlled report folder."
    End If

    Dim reportPath As String
    Dim temporaryPath As String
    reportPath = reportFolder & "\QA_REPORT.txt"
    temporaryPath = reportFolder & "\QA_REPORT.tmp"

    evidence.EvidenceReportPath = reportPath
    evidence.EvidenceWriteSucceeded = True
    evidence.EvidenceWriteVerified = True
    evidence.MarkStageProved "EVIDENCE_WRITE", _
        "Unicode temporary write, exact readback, atomic rename, and final readback proved"
    evidence.SealRequiredStages
    summary = evidence.BuildText

    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    Dim writer As Object
    Set writer = fileSystem.CreateTextFile(temporaryPath, True, True)
    writer.Write summary
    writer.Close
    Set writer = Nothing

    Dim reader As Object
    Set reader = fileSystem.OpenTextFile(temporaryPath, 1, False, -1)

    Dim readback As String
    readback = reader.ReadAll
    reader.Close
    Set reader = Nothing

    If StrComp(readback, summary, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 1602, _
            "WriteEvidenceFileAtomically", _
            "Temporary evidence readback did not exactly match the report text."
    End If

    Name temporaryPath As reportPath

    Set reader = fileSystem.OpenTextFile(reportPath, 1, False, -1)
    readback = reader.ReadAll
    reader.Close
    Set reader = Nothing

    If StrComp(readback, summary, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 1603, _
            "WriteEvidenceFileAtomically", _
            "Final evidence readback did not exactly match the report text."
    End If

    WriteEvidenceFileAtomically = True
    Exit Function

Failed:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description

    On Error Resume Next
    If Not (writer Is Nothing) Then writer.Close
    If Not (reader Is Nothing) Then reader.Close

    If Len(reportPath) > 0 And DirectoryEntryExists(reportPath) Then
        Name reportPath As reportPath & ".unverified"
    End If
    On Error GoTo 0

    evidence.EvidenceWriteSucceeded = False
    evidence.EvidenceWriteVerified = False
    evidence.MarkStageFailed "EVIDENCE_WRITE", _
        "error " & CStr(failureNumber) & ": " & failureDescription
    evidence.AddFailure "Evidence write/readback failed " & _
        CStr(failureNumber) & ": " & failureDescription
    evidence.SealRequiredStages
    summary = evidence.BuildText
    WriteEvidenceFileAtomically = False
End Function

Private Function CreateUniqueReportFolder( _
    ByRef evidence As CRunEvidence) As String

    On Error GoTo Failed

    Dim marker As String
    marker = "\test_assets\models\"

    Dim markerPosition As Long
    markerPosition = InStr(1, evidence.PartPath, marker, vbTextCompare)
    If markerPosition = 0 Then Exit Function

    Dim testAssetsPath As String
    testAssetsPath = Left$(evidence.PartPath, markerPosition - 1) & _
        "\test_assets"
    If Not DirectoryExists(testAssetsPath) Then Exit Function

    Dim evidenceRoot As String
    evidenceRoot = testAssetsPath & "\iteration_evidence"
    If Not DirectoryExists(evidenceRoot) Then MkDir evidenceRoot

    Dim macroQaRoot As String
    macroQaRoot = evidenceRoot & "\macro_qa"
    If Not DirectoryExists(macroQaRoot) Then MkDir macroQaRoot

    Dim baseFolder As String
    baseFolder = macroQaRoot & "\" & evidence.RunId & "_" & _
        FileNameWithoutExtension(evidence.PartPath)

    Dim reportFolder As String
    reportFolder = baseFolder

    Dim suffix As Long
    Do While DirectoryExists(reportFolder)
        suffix = suffix + 1
        reportFolder = baseFolder & "_" & Format$(suffix, "000")
    Loop

    MkDir reportFolder
    CreateUniqueReportFolder = reportFolder
    Exit Function

Failed:
    CreateUniqueReportFolder = vbNullString
End Function

Private Function DirectoryExists(ByVal path As String) As Boolean
    On Error Resume Next
    DirectoryExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Function DirectoryEntryExists(ByVal path As String) As Boolean
    On Error Resume Next
    DirectoryEntryExists = (Len(Dir$(path, vbNormal Or vbHidden Or vbSystem)) > 0)
    On Error GoTo 0
End Function

Private Function FileNameWithoutExtension(ByVal path As String) As String
    Dim slashPosition As Long
    slashPosition = InStrRev(path, "\")

    Dim nameOnly As String
    If slashPosition > 0 Then
        nameOnly = Mid$(path, slashPosition + 1)
    Else
        nameOnly = path
    End If

    Dim dotPosition As Long
    dotPosition = InStrRev(nameOnly, ".")

    If dotPosition > 1 Then
        FileNameWithoutExtension = Left$(nameOnly, dotPosition - 1)
    Else
        FileNameWithoutExtension = nameOnly
    End If
End Function

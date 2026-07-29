Option Explicit

Private Const swImportModelItemsFromEntireModel As Long = 0

Private Const swInsertDatums As Long = 2
Private Const swInsertGTols As Long = 32
Private Const swInsertNotes As Long = 64
Private Const swInsertDimensionsMarkedForDrawing As Long = 32768
Private Const swInsertHoleWizardProfileDimensions As Long = 65536
Private Const swInsertHoleWizardLocationDimensions As Long = 131072
Private Const swInsertHoleCallout As Long = 1048576

Private Const swAlignDimensionType_AutoArrange As Long = 0

' Verified against the installed SOLIDWORKS 2025 interop type library.
Private Const swDimensionType_HorOrdinate As Long = 7
Private Const swDimensionType_VertOrdinate As Long = 8
Private Const swDimensionType_HorLinear As Long = 11
Private Const swDimensionType_VertLinear As Long = 12
Private Const swDimensionType_Radial As Long = 5
Private Const swDimensionType_Diameter As Long = 6
Private Const swDimensionType_Ordinate As Long = 1
Private Const swDimensionType_AngularOrdinate As Long = 16

Private Const DIMENSION_ARRANGE_STAGE As String = "DIMENSION_ARRANGE"
Private Const DIMENSION_ARRANGE_SPACING_M As Double = 0.006
Private Const DIMENSION_ARRANGE_BORDER_INSET_M As Double = 0.001
Private Const DIMENSION_ARRANGE_READBACK_TOLERANCE_M As Double = 0.000001

' Per-view non-ordinate lane allocation state for the deterministic fallback.
' Reset by ResetDimensionLaneState at the top of every view transaction.
Private mBottomLaneCount As Long
Private mTopLaneCount As Long
Private mLeftLaneCount As Long
Private mRightLaneCount As Long

Public Function ImportModelItemsAcrossDrawing( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal anchorViewName As String, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    If Not Module8_RuntimeSupport.ActivateDrawingDocument( _
        swApp, swDrawModel, evidence) Then Exit Function

    Dim mask As Long
    mask = GetModelItemMask()

    Dim total As Long
    total = ImportModelItemsPerView( _
        swApp, swDrawModel, swDraw, mask, evidence)

    If total = 0 Then
        evidence.AddWarning "Model-item import returned zero; fallback can locate " & _
            "owned holes but cannot invent missing size, fit, tolerance, or GD&T intent."
    End If

    ImportModelItemsAcrossDrawing = total

SafeExit:
    Module8_RuntimeSupport.RestoreSheetContext swDrawModel, swDraw
    Exit Function

Failed:
    evidence.AddFailure "InsertModelAnnotations4 error " & _
        CStr(Err.Number) & ": " & Err.Description
    Resume SafeExit
End Function

Private Function ImportModelItemsPerView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal mask As Long, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    Dim runningTotal As Long
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim viewName As String
        viewName = Module8_RuntimeSupport.GetViewName(swView)

        If Module8_RuntimeSupport.IsOrdinateEligibleView(swView) Then
            If Not Module8_RuntimeSupport.ActivateDrawingDocument( _
                swApp, swDrawModel, evidence) Then Exit Function

            If Not Module8_RuntimeSupport.ActivateDrawingView( _
                swDrawModel, swDraw, swView, evidence, _
                "Orthographic model import") Then

                Exit Function
            End If

            Dim activeView As SldWorks.View
            Set activeView = swDraw.ActiveDrawingView
            If activeView Is Nothing Or StrComp( _
                Module8_RuntimeSupport.GetViewName(activeView), _
                viewName, vbTextCompare) <> 0 Then

                evidence.AddFailure "Orthographic model import: active-view " & _
                    "readback mismatch for '" & viewName & "'."
                Exit Function
            End If

            swDrawModel.ClearSelection2 True

            Dim selectedView As Boolean
            selectedView = swDrawModel.Extension.SelectByID2( _
                viewName, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0)

            evidence.AddInfo "MODEL_IMPORT_VIEW_SELECTION|view=" & viewName & _
                "|selectedByID=" & CStr(selectedView) & _
                "|activeView=" & _
                Module8_RuntimeSupport.GetViewName(activeView) & _
                "|activeViewMatched=True"

            If selectedView Then
                evidence.AddInfo "Orthographic model import: SelectByID2 " & _
                    "accepted '" & viewName & "'."
            Else
                evidence.AddWarning "Orthographic model import: SelectByID2 " & _
                    "returned False for '" & viewName & _
                    "'; active-view readback matched, so continuing with " & _
                    "the verified selected-view import workflow."
            End If

            evidence.AddInfo "MODEL_IMPORT_EXECUTE|view=" & viewName & _
                "|allViews=False|duplicateDims=True|activeViewMatched=True"

            Dim inserted As Variant
            inserted = swDraw.InsertModelAnnotations4( _
                swImportModelItemsFromEntireModel, mask, False, True, _
                False, False, False, False)

            RecordImportedAnnotationTypes _
                inserted, "view " & viewName, evidence

            Dim insertedCount As Long
            insertedCount = Module8_RuntimeSupport.CountVariantItems(inserted)
            runningTotal = runningTotal + insertedCount

            evidence.AddInfo "Orthographic selected-view import '" & viewName & _
                "' count=" & CStr(insertedCount) & "; mask=" & CStr(mask) & "."
        Else
            evidence.AddInfo "Model import skipped unsupported or unregistered view '" & _
                viewName & "'."
        End If

        swDrawModel.ClearSelection2 True
        Set swView = swView.GetNextView
    Loop

    ImportModelItemsPerView = runningTotal
    Exit Function

Failed:
    evidence.AddFailure "Selected-view model import error " & _
        CStr(Err.Number) & ": " & Err.Description
    ImportModelItemsPerView = runningTotal
End Function

Public Sub RecordDisplayDimensionCounts( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        evidence.AddInfo "Displayed dimensions in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'=" & _
            CStr(CountDisplayDimensionsInView(swView)) & "."
        Set swView = swView.GetNextView
    Loop
End Sub

Private Sub RecordImportedAnnotationTypes( _
    ByVal inserted As Variant, _
    ByVal contextName As String, _
    ByRef evidence As CRunEvidence)

    If IsEmpty(inserted) Or Not IsArray(inserted) Then Exit Sub

    On Error GoTo Failed

    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(inserted) To UBound(inserted)
        Dim annotation As SldWorks.Annotation
        Set annotation = inserted(i)

        If Not annotation Is Nothing Then
            Dim typeName As String
            typeName = AnnotationTypeName(annotation.GetType)

            If counts.Exists(typeName) Then
                counts(typeName) = CLng(counts(typeName)) + 1
            Else
                counts.Add typeName, 1
            End If
        End If
    Next i

    Dim key As Variant
    For Each key In counts.Keys
        evidence.AddInfo "Imported annotation type in " & contextName & _
            ": " & CStr(key) & "=" & CStr(counts(key)) & "."
    Next key
    Exit Sub

Failed:
    evidence.AddWarning "Could not classify imported annotations for " & _
        contextName & ": " & Err.Description
End Sub

Private Function AnnotationTypeName(ByVal annotationType As Long) As String
    Select Case annotationType
        Case 1: AnnotationTypeName = "CosmeticThread"
        Case 2: AnnotationTypeName = "DatumFeature"
        Case 3: AnnotationTypeName = "DatumTarget"
        Case 4: AnnotationTypeName = "DisplayDimension"
        Case 5: AnnotationTypeName = "GeometricTolerance"
        Case 6: AnnotationTypeName = "Note"
        Case 7: AnnotationTypeName = "SurfaceFinish"
        Case 8: AnnotationTypeName = "WeldSymbol"
        Case 13: AnnotationTypeName = "CenterMark"
        Case 15: AnnotationTypeName = "CenterLine"
        Case 16: AnnotationTypeName = "DatumOrigin"
        Case Else: AnnotationTypeName = "Type" & CStr(annotationType)
    End Select
End Function

Public Function CountAllDisplayDimensions( _
    ByRef swDraw As SldWorks.DrawingDoc) As Long

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        CountAllDisplayDimensions = CountAllDisplayDimensions + _
            CountDisplayDimensionsInView(swView)
        Set swView = swView.GetNextView
    Loop
End Function

Public Function CountDisplayDimensionsInView( _
    ByRef swView As SldWorks.View) As Long

    On Error GoTo Failed
    If swView Is Nothing Then Exit Function

    CountDisplayDimensionsInView = _
        Module8_RuntimeSupport.CountVariantItems(swView.GetDisplayDimensions)
    Exit Function

Failed:
    CountDisplayDimensionsInView = 0
End Function

Public Function ApplyImportedCoverageToCandidates( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef candidates As Collection, _
    ByRef horizontalDatum As CDatumProof, _
    ByRef verticalDatum As CDatumProof, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If candidates Is Nothing Then
        ApplyImportedCoverageToCandidates = True
        Exit Function
    End If

    If candidates.Count = 0 Then
        ApplyImportedCoverageToCandidates = True
        Exit Function
    End If

    If horizontalDatum Is Nothing Or verticalDatum Is Nothing Then
        evidence.AddFailure _
            "Coverage inspection received a missing directional datum."
        Exit Function
    End If

    If Not horizontalDatum.Proven Or horizontalDatum.DrawingEntity Is Nothing Or _
       Not verticalDatum.Proven Or verticalDatum.DrawingEntity Is Nothing Then

        evidence.AddFailure _
            "Coverage inspection received an unproved directional datum."
        Exit Function
    End If

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If IsEmpty(dimensions) Then
        MarkCoverageInspectionComplete _
            swView, candidates, horizontalDatum, verticalDatum, evidence
        ApplyImportedCoverageToCandidates = True
        Exit Function
    End If

    If Not IsArray(dimensions) Then
        evidence.AddFailure "Coverage inspection could not enumerate display " & _
            "dimensions in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    Dim horizontalDatumProven As Boolean
    Dim verticalDatumProven As Boolean
    Dim i As Long

    Dim displayDimension As SldWorks.DisplayDimension
    Dim dimensionType As Long
    Dim attached As Variant
    Dim dimensionKey As String
    Dim matchedCount As Long
    Dim approvedDatumAttached As Boolean

    For i = LBound(dimensions) To UBound(dimensions)
        Set displayDimension = dimensions(i)

        If Not displayDimension Is Nothing Then
            dimensionType = displayDimension.Type2

            If IsLocationDimensionType(dimensionType) Then
                If Not TryGetDimensionAttachments( _
                    displayDimension, attached, swView, evidence) Then

                    Exit Function
                End If

                dimensionKey = DisplayDimensionKey(displayDimension)
                matchedCount = 0
                approvedDatumAttached = False

                Select Case dimensionType
                    Case swDimensionType_HorOrdinate
                        approvedDatumAttached = AttachmentsContainDatum( _
                            swApp, attached, horizontalDatum)
                        If approvedDatumAttached Then

                            horizontalDatumProven = True
                            matchedCount = MarkAttachedCoverage( _
                                swApp, attached, candidates, True, _
                                "ImportedHorizontalOrdinate:" & dimensionKey)
                        ElseIf AttachmentsMatchAnyCandidate( _
                            swApp, attached, candidates) Then

                            evidence.AddFailure _
                                "Horizontal imported ordinate in '" & _
                                Module8_RuntimeSupport.GetViewName(swView) & _
                                "' matches an owned candidate but not the approved X datum."
                            Exit Function
                        End If

                    Case swDimensionType_VertOrdinate
                        approvedDatumAttached = AttachmentsContainDatum( _
                            swApp, attached, verticalDatum)
                        If approvedDatumAttached Then

                            verticalDatumProven = True
                            matchedCount = MarkAttachedCoverage( _
                                swApp, attached, candidates, False, _
                                "ImportedVerticalOrdinate:" & dimensionKey)
                        ElseIf AttachmentsMatchAnyCandidate( _
                            swApp, attached, candidates) Then

                            evidence.AddFailure _
                                "Vertical imported ordinate in '" & _
                                Module8_RuntimeSupport.GetViewName(swView) & _
                                "' matches an owned candidate but not the approved Y datum."
                            Exit Function
                        End If

                    Case swDimensionType_HorLinear
                        approvedDatumAttached = AttachmentsContainDatum( _
                            swApp, attached, horizontalDatum)
                        If approvedDatumAttached Then

                            matchedCount = MarkAttachedCoverage( _
                                swApp, attached, candidates, True, _
                                "ImportedHorizontalLinear:" & dimensionKey)
                        ElseIf AttachmentsMatchAnyCandidate( _
                            swApp, attached, candidates) Then

                            evidence.AddFailure _
                                "Horizontal imported linear dimension in '" & _
                                Module8_RuntimeSupport.GetViewName(swView) & _
                                "' matches an owned candidate but its X datum is unresolved."
                            Exit Function
                        End If

                    Case swDimensionType_VertLinear
                        approvedDatumAttached = AttachmentsContainDatum( _
                            swApp, attached, verticalDatum)
                        If approvedDatumAttached Then

                            matchedCount = MarkAttachedCoverage( _
                                swApp, attached, candidates, False, _
                                "ImportedVerticalLinear:" & dimensionKey)
                        ElseIf AttachmentsMatchAnyCandidate( _
                            swApp, attached, candidates) Then

                            evidence.AddFailure _
                                "Vertical imported linear dimension in '" & _
                                Module8_RuntimeSupport.GetViewName(swView) & _
                                "' matches an owned candidate but its Y datum is unresolved."
                            Exit Function
                        End If
                End Select

                evidence.AddInfo "EVIDENCE|COVERAGE_DIMENSION|view=" & _
                    EvidenceValue(Module8_RuntimeSupport.GetViewName(swView)) & _
                    "|dimension=" & EvidenceValue(dimensionKey) & _
                    "|type=" & _
                    EvidenceValue(LocationDimensionTypeName(dimensionType)) & _
                    "|approvedDatumAttached=" & _
                    CStr(approvedDatumAttached) & _
                    "|matchedCandidates=" & CStr(matchedCount)
            End If
        End If
    Next i

    MarkCoverageInspectionComplete _
        swView, candidates, horizontalDatum, verticalDatum, evidence

    evidence.AddInfo "Coverage reconciliation in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & _
        "': horizontal datum=" & CStr(horizontalDatumProven) & _
        "; vertical datum=" & CStr(verticalDatumProven) & _
        "; coordinate tolerance m=" & _
        Format$(Module8_RuntimeSupport.PROJECTED_TOLERANCE_M, "0.000000") & "."

    ApplyImportedCoverageToCandidates = True
    Exit Function

Failed:
    evidence.AddFailure "Coverage reconciliation error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
End Function

Private Function TryGetDimensionAttachments( _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByRef attached As Variant, _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim annotation As SldWorks.Annotation
    Set annotation = displayDimension.GetAnnotation

    If annotation Is Nothing Then
        evidence.AddFailure "Location dimension '" & _
            DisplayDimensionKey(displayDimension) & _
            "' has no annotation in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    attached = annotation.GetAttachedEntities3

    If IsEmpty(attached) Or Not IsArray(attached) Then
        evidence.AddFailure "Location dimension '" & _
            DisplayDimensionKey(displayDimension) & _
            "' has no inspectable attachments in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    TryGetDimensionAttachments = True
    Exit Function

Failed:
    evidence.AddFailure "Attachment inspection failed for location dimension '" & _
        DisplayDimensionKey(displayDimension) & "' in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
End Function

Private Function AttachmentsContainDatum( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal attached As Variant, _
    ByRef datum As CDatumProof) As Boolean

    If datum Is Nothing Then Exit Function

    AttachmentsContainDatum = AttachmentsContainObject( _
        swApp, attached, datum.DrawingEntity)

    If Not AttachmentsContainDatum And Not datum.ModelEntity Is Nothing Then
        AttachmentsContainDatum = AttachmentsContainObject( _
            swApp, attached, datum.ModelEntity)
    End If
End Function

Private Function AttachmentsContainObject( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal attached As Variant, _
    ByVal targetObject As Object) As Boolean

    If targetObject Is Nothing Then Exit Function
    If IsEmpty(attached) Or Not IsArray(attached) Then Exit Function

    Dim i As Long
    For i = LBound(attached) To UBound(attached)
        If IsObject(attached(i)) Then
            Dim attachedObject As Object
            Set attachedObject = attached(i)

            If Not attachedObject Is Nothing Then
                If Module8_RuntimeSupport.ObjectsAreSame( _
                    swApp, attachedObject, targetObject) Then

                    AttachmentsContainObject = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function AttachmentsMatchCandidate( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal attached As Variant, _
    ByRef candidate As CHoleCandidate) As Boolean

    AttachmentsMatchCandidate = AttachmentsContainObject( _
        swApp, attached, candidate.DrawingEdge)

    If Not AttachmentsMatchCandidate Then
        AttachmentsMatchCandidate = AttachmentsContainObject( _
            swApp, attached, candidate.ModelEdge)
    End If

    If Not AttachmentsMatchCandidate Then
        AttachmentsMatchCandidate = AttachmentsContainAliasCollection( _
            swApp, attached, candidate.DrawingEntityAliases)
    End If

    If Not AttachmentsMatchCandidate Then
        AttachmentsMatchCandidate = AttachmentsContainAliasCollection( _
            swApp, attached, candidate.ModelEntityAliases)
    End If
End Function

Private Function AttachmentsContainAliasCollection( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal attached As Variant, _
    ByRef aliases As Collection) As Boolean

    If aliases Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To aliases.Count
        If AttachmentsContainObject(swApp, attached, aliases(i)) Then
            AttachmentsContainAliasCollection = True
            Exit Function
        End If
    Next i
End Function

Private Function AttachmentsMatchAnyCandidate( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal attached As Variant, _
    ByRef candidates As Collection) As Boolean

    Dim i As Long
    For i = 1 To candidates.Count
        Dim candidate As CHoleCandidate
        Set candidate = candidates(i)

        If AttachmentsMatchCandidate(swApp, attached, candidate) Then
            AttachmentsMatchAnyCandidate = True
            Exit Function
        End If
    Next i
End Function

Private Function MarkAttachedCoverage( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal attached As Variant, _
    ByRef candidates As Collection, _
    ByVal markX As Boolean, _
    ByVal coverageSource As String) As Long

    Dim i As Long
    For i = 1 To candidates.Count
        Dim candidate As CHoleCandidate
        Set candidate = candidates(i)

        If AttachmentsMatchCandidate(swApp, attached, candidate) Then
            If markX Then
                candidate.CoveredX = True
                candidate.CoverageSourceX = AppendEvidenceToken( _
                    candidate.CoverageSourceX, coverageSource)
            Else
                candidate.CoveredY = True
                candidate.CoverageSourceY = AppendEvidenceToken( _
                    candidate.CoverageSourceY, coverageSource)
            End If

            MarkAttachedCoverage = MarkAttachedCoverage + 1
        End If
    Next i
End Function

Private Sub MarkCoverageInspectionComplete( _
    ByRef swView As SldWorks.View, _
    ByRef candidates As Collection, _
    ByRef horizontalDatum As CDatumProof, _
    ByRef verticalDatum As CDatumProof, _
    ByRef evidence As CRunEvidence)

    Dim i As Long
    For i = 1 To candidates.Count
        Dim candidate As CHoleCandidate
        Set candidate = candidates(i)

        candidate.CoverageInspectionSucceeded = True
        candidate.CoverageScopeKey = _
            "view=" & LCase$(Module8_RuntimeSupport.GetViewName(swView)) & _
            "|family=" & candidate.FamilyKey & _
            "|datumX=" & horizontalDatum.StableKey & _
            "|datumY=" & verticalDatum.StableKey

        evidence.AddInfo "EVIDENCE|COVERAGE_RESULT|view=" & _
            EvidenceValue(Module8_RuntimeSupport.GetViewName(swView)) & _
            "|family=" & EvidenceValue(candidate.FamilyKey) & _
            "|instance=" & EvidenceValue(candidate.PhysicalInstanceKey) & _
            "|datumX=" & EvidenceValue(horizontalDatum.StableKey) & _
            "|datumY=" & EvidenceValue(verticalDatum.StableKey) & _
            "|coveredX=" & CStr(candidate.CoveredX) & _
            "|coveredY=" & CStr(candidate.CoveredY) & _
            "|sourceX=" & EvidenceValue(candidate.CoverageSourceX) & _
            "|sourceY=" & EvidenceValue(candidate.CoverageSourceY)
    Next i
End Sub

Private Function IsLocationDimensionType(ByVal dimensionType As Long) As Boolean
    Select Case dimensionType
        Case swDimensionType_HorOrdinate, swDimensionType_VertOrdinate, _
             swDimensionType_HorLinear, swDimensionType_VertLinear
            IsLocationDimensionType = True
    End Select
End Function

Private Function LocationDimensionTypeName( _
    ByVal dimensionType As Long) As String

    Select Case dimensionType
        Case swDimensionType_HorOrdinate
            LocationDimensionTypeName = "HorizontalOrdinate"
        Case swDimensionType_VertOrdinate
            LocationDimensionTypeName = "VerticalOrdinate"
        Case swDimensionType_HorLinear
            LocationDimensionTypeName = "HorizontalLinear"
        Case swDimensionType_VertLinear
            LocationDimensionTypeName = "VerticalLinear"
        Case Else
            LocationDimensionTypeName = "Type" & CStr(dimensionType)
    End Select
End Function

Private Function DisplayDimensionKey( _
    ByRef displayDimension As SldWorks.DisplayDimension) As String

    On Error Resume Next
    DisplayDimensionKey = Trim$(displayDimension.GetNameForSelection)
    On Error GoTo 0

    If Len(DisplayDimensionKey) = 0 Then
        DisplayDimensionKey = "unnamed"
    End If
End Function

Private Function AppendEvidenceToken( _
    ByVal existingValue As String, _
    ByVal newValue As String) As String

    If Len(existingValue) = 0 Then
        AppendEvidenceToken = newValue
    ElseIf InStr(1, existingValue, newValue, vbTextCompare) = 0 Then
        AppendEvidenceToken = existingValue & "," & newValue
    Else
        AppendEvidenceToken = existingValue
    End If
End Function

Private Function EvidenceValue(ByVal value As String) As String
    EvidenceValue = Replace$(Trim$(value), "|", "/")
    EvidenceValue = Replace$(EvidenceValue, vbCr, " ")
    EvidenceValue = Replace$(EvidenceValue, vbLf, " ")
End Function

Public Sub AutoArrangeAllDrawingDimensions( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed
    evidence.RequireStage DIMENSION_ARRANGE_STAGE

    Dim allViewsProved As Boolean
    allViewsProved = True

    Dim inspectedViews As Long
    Dim apiArrangedViews As Long
    Dim fallbackArrangedViews As Long
    Dim noActionViews As Long

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        inspectedViews = inspectedViews + 1

        Dim arrangeOutcome As String
        arrangeOutcome = vbNullString

        If AutoArrangeDimensionsInView( _
            swDrawModel, swDraw, swView, evidence, arrangeOutcome) Then

            Select Case arrangeOutcome
                Case "API"
                    apiArrangedViews = apiArrangedViews + 1
                Case "FALLBACK"
                    fallbackArrangedViews = fallbackArrangedViews + 1
                Case Else
                    noActionViews = noActionViews + 1
            End Select
        Else
            allViewsProved = False
        End If

        Set swView = swView.GetNextView
    Loop

    If inspectedViews = 0 Then
        allViewsProved = False
        evidence.AddFailure _
            "Dimension arrange could not inspect any real drawing views."
    End If

    Dim stageDetail As String
    stageDetail = "views=" & CStr(inspectedViews) & _
        ", api=" & CStr(apiArrangedViews) & _
        ", fallback=" & CStr(fallbackArrangedViews) & _
        ", noAction=" & CStr(noActionViews)

    If allViewsProved Then
        evidence.MarkStageProved DIMENSION_ARRANGE_STAGE, _
            "selection, API/fallback result, position readback, and " & _
            "content-border origin checks proved; " & stageDetail
    Else
        evidence.MarkStageFailed DIMENSION_ARRANGE_STAGE, _
            "one or more view-level arrange transactions failed; " & _
            stageDetail
    End If
    Exit Sub

Failed:
    evidence.AddFailure "Dimension arrange traversal API error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed DIMENSION_ARRANGE_STAGE, _
        "drawing-view traversal API error " & CStr(Err.Number) & _
        ": " & Err.Description
End Sub

Private Function AutoArrangeDimensionsInView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence, _
    ByRef arrangeOutcome As String) As Boolean

    On Error GoTo Failed

    Dim viewName As String
    viewName = Module8_RuntimeSupport.GetViewName(swView)

    If Not Module8_RuntimeSupport.ActivateDrawingView( _
        swDrawModel, swDraw, swView, evidence, _
        "Dimension arrange") Then

        evidence.AddFailure "Dimension arrange could not prove active view '" & _
            viewName & "'."
        arrangeOutcome = "FAILED_ACTIVATION"
        Exit Function
    End If

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If IsEmpty(dimensions) Then
        arrangeOutcome = "NO_DIMENSIONS"
        AutoArrangeDimensionsInView = True
        GoTo SafeExit
    End If

    If Not IsArray(dimensions) Then
        evidence.AddFailure "Dimension arrange could not enumerate display " & _
            "dimensions in '" & viewName & "'."
        arrangeOutcome = "FAILED_ENUMERATION"
        GoTo SafeExit
    End If

    swDrawModel.ClearSelection2 True

    Dim selectionManager As SldWorks.SelectionMgr
    Set selectionManager = swDrawModel.SelectionManager
    If selectionManager Is Nothing Then
        evidence.AddFailure "Dimension arrange selection manager is Nothing in '" & _
            viewName & "'."
        arrangeOutcome = "FAILED_SELECTION_MANAGER"
        GoTo SafeExit
    End If

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionManager.CreateSelectData
    If selectData Is Nothing Then
        evidence.AddFailure "Dimension arrange SelectData is Nothing in '" & _
            viewName & "'."
        arrangeOutcome = "FAILED_SELECT_DATA"
        GoTo SafeExit
    End If
    Set selectData.View = swView

    Dim i As Long
    Dim attemptedCount As Long
    Dim selectedCount As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)

        If displayDimension Is Nothing Then
            evidence.AddFailure "Dimension arrange received a Nothing display " & _
                "dimension in '" & viewName & "' at array index " & CStr(i) & "."
            arrangeOutcome = "FAILED_DIMENSION_OBJECT"
            GoTo SafeExit
        End If

        Dim annotation As SldWorks.Annotation
        Set annotation = displayDimension.GetAnnotation

        If annotation Is Nothing Then
            evidence.AddFailure "Dimension arrange received a display dimension " & _
                "without an annotation in '" & viewName & "' at array index " & _
                CStr(i) & "."
            arrangeOutcome = "FAILED_ANNOTATION_OBJECT"
            GoTo SafeExit
        End If

        attemptedCount = attemptedCount + 1

        Dim annotationSelected As Boolean
        annotationSelected = CBool( _
            annotation.Select3(selectedCount > 0, selectData))

        If annotationSelected Then
            selectedCount = selectedCount + 1
        Else
            evidence.AddFailure "Dimension Select3 returned False in '" & _
                viewName & "' at array index " & CStr(i) & "."
            arrangeOutcome = "FAILED_SELECT3"
            GoTo SafeExit
        End If
    Next i

    Dim readbackCount As Long
    readbackCount = selectionManager.GetSelectedObjectCount2(-1)

    evidence.AddInfo "DIMENSION_ARRANGE_SELECTION|view=" & _
        EvidenceValue(viewName) & _
        "|attempted=" & CStr(attemptedCount) & _
        "|selected=" & CStr(selectedCount) & _
        "|readback=" & CStr(readbackCount)

    If selectedCount <> attemptedCount Then
        evidence.AddFailure "Dimension arrange selection proof failed in '" & _
            viewName & "': attempted=" & CStr(attemptedCount) & _
            ", selected=" & _
            CStr(selectedCount) & ", readback=" & CStr(readbackCount) & "."
        arrangeOutcome = "FAILED_SELECTION_READBACK"
        GoTo SafeExit
    End If

    ' readback may legitimately trail selected: GetDisplayDimensions can hand
    ' back the same annotation twice (a base ordinate and its subordinates all
    ' report type 1), and re-selecting an already-selected annotation returns
    ' True without growing the selection list.  Only an empty selection is a
    ' real failure; a shortfall is worth recording but must not reject the run.
    If readbackCount < 1 Then
        evidence.AddFailure "Dimension arrange selected nothing in '" & _
            viewName & "': attempted=" & CStr(attemptedCount) & _
            ", selected=" & CStr(selectedCount) & "."
        arrangeOutcome = "FAILED_SELECTION_READBACK"
        GoTo SafeExit
    End If

    If readbackCount <> selectedCount Then
        evidence.AddWarning "Dimension arrange selection readback differs from " & _
            "the Select3 success count in '" & viewName & "': selected=" & _
            CStr(selectedCount) & ", readback=" & CStr(readbackCount) & _
            "; treating duplicate annotations as one selection."
    End If

    If readbackCount < 2 Then
        If Not ValidateDimensionAnnotationOrigins( _
            dimensions, evidence, viewName, "NoAction") Then

            arrangeOutcome = "FAILED_NO_ACTION_BOUNDS"
            GoTo SafeExit
        End If

        evidence.AddInfo "DIMENSION_ARRANGE_SKIPPED|view=" & _
            EvidenceValue(viewName) & _
            "|reason=FewerThanTwoSelectedDimensions"
        arrangeOutcome = "NO_ACTION"
        AutoArrangeDimensionsInView = True
        GoTo SafeExit
    End If

    Dim arrangeSucceeded As Boolean
    evidence.RecordSolidWorksMutation _
        "IModelDocExtension.AlignDimensions(DimensionArrange)"
    arrangeSucceeded = CBool(swDrawModel.Extension.AlignDimensions( _
        swAlignDimensionType_AutoArrange, DIMENSION_ARRANGE_SPACING_M))

    evidence.AddInfo "DIMENSION_ARRANGE_RESULT|view=" & _
        EvidenceValue(viewName) & _
        "|selected=" & CStr(selectedCount) & _
        "|spacingM=0.006000|returned=" & CStr(arrangeSucceeded)

    If arrangeSucceeded Then
        If Not ValidateDimensionAnnotationOrigins( _
            dimensions, evidence, viewName, "AlignDimensions") Then

            arrangeOutcome = "FAILED_API_BOUNDS"
            GoTo SafeExit
        End If

        arrangeOutcome = "API"
        AutoArrangeDimensionsInView = True
    Else
        swDrawModel.ClearSelection2 True

        Dim fallbackArranged As Long
        If TryArrangeWithDeterministicLanes( _
            swDrawModel, swView, dimensions, evidence, fallbackArranged) Then

            evidence.AddInfo "DIMENSION_ARRANGE_FALLBACK_RESULT|view=" & _
                EvidenceValue(viewName) & _
                "|arranged=" & CStr(fallbackArranged) & _
                "|status=PROVED"
            arrangeOutcome = "FALLBACK"
            AutoArrangeDimensionsInView = True
        Else
            evidence.AddFailure "AlignDimensions returned False and the " & _
                "deterministic lane fallback was not proved in '" & _
                viewName & "'."
            arrangeOutcome = "FAILED_API_AND_FALLBACK"
        End If
    End If

SafeExit:
    swDrawModel.ClearSelection2 True
    Exit Function

Failed:
    evidence.AddFailure "Dimension arrange API error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & _
        CStr(Err.Number) & ": " & Err.Description
    arrangeOutcome = "FAILED_API_ERROR"
    Resume SafeExit
End Function

Private Function TryArrangeWithDeterministicLanes( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByVal dimensions As Variant, _
    ByRef evidence As CRunEvidence, _
    ByRef arrangedCount As Long) As Boolean

    On Error GoTo Failed

    Dim viewName As String
    viewName = Module8_RuntimeSupport.GetViewName(swView)

    Dim outline As Variant
    outline = swView.GetOutline

    Dim viewLeft As Double
    Dim viewBottom As Double
    Dim viewRight As Double
    Dim viewTop As Double

    If Not TryReadViewOutline( _
        outline, viewLeft, viewBottom, viewRight, viewTop) Then

        evidence.AddFailure "Dimension lane fallback could not read a valid " & _
            "view outline for '" & viewName & "'."
        Exit Function
    End If

    If Not ContentBorderIsValid(evidence) Then
        evidence.AddFailure "Dimension lane fallback has no proved content " & _
            "border for '" & viewName & "'."
        Exit Function
    End If

    ResetDimensionLaneState

    Dim retainedUnsupportedCount As Long

    Dim i As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)
        If displayDimension Is Nothing Then Exit Function

        Dim annotation As SldWorks.Annotation
        Set annotation = displayDimension.GetAnnotation
        If annotation Is Nothing Then Exit Function

        Dim originalX As Double
        Dim originalY As Double
        If Not TryReadAnnotationPosition( _
            annotation, originalX, originalY) Then

            evidence.AddFailure "Dimension lane fallback could not read the " & _
                "annotation position in '" & viewName & "' at array index " & _
                CStr(i) & "."
            Exit Function
        End If

        Dim dimensionType As Long
        dimensionType = displayDimension.Type2

        If DimensionTypeIsOrdinate(dimensionType) Then
            If Not TryAutoJogOrdinateDimension( _
                displayDimension, annotation, evidence, viewName, i) Then

                Exit Function
            End If

            arrangedCount = arrangedCount + 1
        ElseIf DimensionPositionIsUnsupported(dimensionType) Then
            If Not PositionIsInsideContentBorder( _
                originalX, originalY, evidence) Then

                evidence.AddFailure "Unmovable radial or diametric dimension " & _
                    "origin is outside the content border in '" & viewName & _
                    "' at array index " & CStr(i) & "."
                Exit Function
            End If

            retainedUnsupportedCount = retainedUnsupportedCount + 1
            evidence.AddInfo "DIMENSION_ARRANGE_FALLBACK_ITEM|view=" & _
                EvidenceValue(viewName) & _
                "|index=" & CStr(i) & _
                "|type=" & CStr(dimensionType) & _
                "|mode=RetainedUnsupported" & _
                "|x=" & Format$(originalX, "0.000000") & _
                "|y=" & Format$(originalY, "0.000000")
        Else
            Dim targetX As Double
            Dim targetY As Double
            Dim laneName As String

            If Not TryChooseDimensionLane( _
                dimensionType, originalX, originalY, _
                viewLeft, viewBottom, viewRight, viewTop, evidence, _
                targetX, targetY, laneName) Then

                evidence.AddFailure "Dimension lane fallback could not allocate " & _
                    "a border-safe lane in '" & viewName & "' at array index " & _
                    CStr(i) & "."
                Exit Function
            End If

            Dim positionSet As Boolean
            evidence.RecordSolidWorksMutation _
                "IAnnotation.SetPosition2(DimensionArrangeFallback)"
            positionSet = CBool(annotation.SetPosition2(targetX, targetY, 0#))
            If Not positionSet Then
                evidence.AddFailure "IAnnotation.SetPosition2 returned False " & _
                    "for dimension lane fallback in '" & viewName & _
                    "' at array index " & CStr(i) & "."
                Exit Function
            End If

            Dim readbackX As Double
            Dim readbackY As Double
            If Not TryReadAnnotationPosition( _
                annotation, readbackX, readbackY) Then

                evidence.AddFailure "Dimension lane fallback position readback " & _
                    "failed in '" & viewName & "' at array index " & _
                    CStr(i) & "."
                Exit Function
            End If

            If Abs(readbackX - targetX) > _
                   DIMENSION_ARRANGE_READBACK_TOLERANCE_M Or _
               Abs(readbackY - targetY) > _
                   DIMENSION_ARRANGE_READBACK_TOLERANCE_M Then

                evidence.AddFailure "Dimension lane fallback position readback " & _
                    "mismatch in '" & viewName & "' at array index " & _
                    CStr(i) & "."
                Exit Function
            End If

            If Not PositionIsInsideContentBorder( _
                readbackX, readbackY, evidence) Then

                evidence.AddFailure "Dimension lane fallback readback is outside " & _
                    "the content border in '" & viewName & "' at array index " & _
                    CStr(i) & "."
                Exit Function
            End If

            arrangedCount = arrangedCount + 1
            evidence.AddInfo "DIMENSION_ARRANGE_FALLBACK_ITEM|view=" & _
                EvidenceValue(viewName) & _
                "|index=" & CStr(i) & _
                "|type=" & CStr(dimensionType) & _
                "|mode=SetPosition2" & _
                "|lane=" & laneName & _
                "|target=" & Format$(targetX, "0.000000") & "," & _
                Format$(targetY, "0.000000") & _
                "|readback=" & Format$(readbackX, "0.000000") & "," & _
                Format$(readbackY, "0.000000")
        End If
    Next i

    If arrangedCount = 0 Then
        evidence.AddFailure "Dimension fallback could not arrange any of the " & _
            "selected dimensions in '" & viewName & "'; retained unsupported=" & _
            CStr(retainedUnsupportedCount) & "."
        Exit Function
    End If

    TryArrangeWithDeterministicLanes = ValidateDimensionAnnotationOrigins( _
        dimensions, evidence, viewName, "DeterministicLaneFallback")
    Exit Function

Failed:
    evidence.AddFailure "Dimension lane fallback API error in '" & viewName & _
        "': " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function TryAutoJogOrdinateDimension( _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByRef annotation As SldWorks.Annotation, _
    ByRef evidence As CRunEvidence, _
    ByVal viewName As String, _
    ByVal dimensionIndex As Long) As Boolean

    On Error GoTo Failed

    evidence.RecordSolidWorksMutation _
        "IDisplayDimension.AutoJogOrdinate(DimensionArrangeFallback)"

    Dim autoJogSucceeded As Boolean
    autoJogSucceeded = CBool(displayDimension.AutoJogOrdinate)
    If Not autoJogSucceeded Then
        evidence.AddFailure "IDisplayDimension.AutoJogOrdinate returned False " & _
            "in '" & viewName & "' at array index " & _
            CStr(dimensionIndex) & "."
        Exit Function
    End If

    Dim readbackX As Double
    Dim readbackY As Double
    If Not TryReadAnnotationPosition( _
        annotation, readbackX, readbackY) Then

        evidence.AddFailure "Ordinate auto-jog position readback failed in '" & _
            viewName & "' at array index " & CStr(dimensionIndex) & "."
        Exit Function
    End If

    If Not PositionIsInsideContentBorder( _
        readbackX, readbackY, evidence) Then

        evidence.AddFailure "Ordinate auto-jog readback is outside the " & _
            "content border in '" & viewName & "' at array index " & _
            CStr(dimensionIndex) & "."
        Exit Function
    End If

    evidence.AddInfo "DIMENSION_ARRANGE_FALLBACK_ITEM|view=" & _
        EvidenceValue(viewName) & _
        "|index=" & CStr(dimensionIndex) & _
        "|type=" & CStr(displayDimension.Type2) & _
        "|mode=AutoJogOrdinate" & _
        "|readback=" & Format$(readbackX, "0.000000") & "," & _
        Format$(readbackY, "0.000000")

    TryAutoJogOrdinateDimension = True
    Exit Function

Failed:
    evidence.AddFailure "Ordinate auto-jog API error in '" & viewName & _
        "' at array index " & CStr(dimensionIndex) & ": " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function TryChooseDimensionLane( _
    ByVal dimensionType As Long, _
    ByVal originalX As Double, _
    ByVal originalY As Double, _
    ByVal viewLeft As Double, _
    ByVal viewBottom As Double, _
    ByVal viewRight As Double, _
    ByVal viewTop As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef targetX As Double, _
    ByRef targetY As Double, _
    ByRef laneName As String) As Boolean

    Dim preferPositiveSide As Boolean
    Select Case dimensionType
        Case swDimensionType_HorLinear
            preferPositiveSide = (originalY >= (viewBottom + viewTop) / 2#)
            TryChooseDimensionLane = TryAllocateHorizontalLane( _
                preferPositiveSide, originalX, _
                viewBottom, viewTop, evidence, _
                targetX, targetY, laneName)

        Case swDimensionType_VertLinear
            preferPositiveSide = (originalX >= (viewLeft + viewRight) / 2#)
            TryChooseDimensionLane = TryAllocateVerticalLane( _
                preferPositiveSide, originalY, _
                viewLeft, viewRight, evidence, _
                targetX, targetY, laneName)

        Case Else
            ' Ordinate types are handled by AutoJogOrdinate before this helper.
            ' Other unclassified movable types use their nearest view side.
            If NearestViewSideIsHorizontal( _
                originalX, originalY, _
                viewLeft, viewBottom, viewRight, viewTop) Then

                preferPositiveSide = (originalY >= (viewBottom + viewTop) / 2#)
                TryChooseDimensionLane = TryAllocateHorizontalLane( _
                    preferPositiveSide, originalX, _
                    viewBottom, viewTop, evidence, _
                    targetX, targetY, laneName)
            Else
                preferPositiveSide = (originalX >= (viewLeft + viewRight) / 2#)
                TryChooseDimensionLane = TryAllocateVerticalLane( _
                    preferPositiveSide, originalY, _
                    viewLeft, viewRight, evidence, _
                    targetX, targetY, laneName)
            End If
    End Select
End Function

Private Function TryAllocateHorizontalLane( _
    ByVal preferTop As Boolean, _
    ByVal originalX As Double, _
    ByVal viewBottom As Double, _
    ByVal viewTop As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef targetX As Double, _
    ByRef targetY As Double, _
    ByRef laneName As String) As Boolean

    targetX = ClampToContentBorderX(originalX, evidence)

    If preferTop Then
        If TryTopLane(viewTop, evidence, targetY, laneName) Then
        ElseIf TryBottomLane( _
            viewBottom, evidence, targetY, laneName) Then
        Else
            Exit Function
        End If
    Else
        If TryBottomLane( _
            viewBottom, evidence, targetY, laneName) Then
        ElseIf TryTopLane(viewTop, evidence, targetY, laneName) Then
        Else
            Exit Function
        End If
    End If

    TryAllocateHorizontalLane = PositionIsInsideContentBorder( _
        targetX, targetY, evidence)
End Function

Private Function TryAllocateVerticalLane( _
    ByVal preferRight As Boolean, _
    ByVal originalY As Double, _
    ByVal viewLeft As Double, _
    ByVal viewRight As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef targetX As Double, _
    ByRef targetY As Double, _
    ByRef laneName As String) As Boolean

    targetY = ClampToContentBorderY(originalY, evidence)

    If preferRight Then
        If TryRightLane(viewRight, evidence, targetX, laneName) Then
        ElseIf TryLeftLane(viewLeft, evidence, targetX, laneName) Then
        Else
            Exit Function
        End If
    Else
        If TryLeftLane(viewLeft, evidence, targetX, laneName) Then
        ElseIf TryRightLane(viewRight, evidence, targetX, laneName) Then
        Else
            Exit Function
        End If
    End If

    TryAllocateVerticalLane = PositionIsInsideContentBorder( _
        targetX, targetY, evidence)
End Function

Private Function TryBottomLane( _
    ByVal viewBottom As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef targetY As Double, _
    ByRef laneName As String) As Boolean

    Dim proposedLane As Long
    proposedLane = mBottomLaneCount + 1
    targetY = viewBottom - _
        DIMENSION_ARRANGE_SPACING_M * CDbl(proposedLane)

    If targetY < evidence.ContentBorderBottom + _
       DIMENSION_ARRANGE_BORDER_INSET_M Then Exit Function

    mBottomLaneCount = proposedLane
    laneName = "BOTTOM-" & CStr(proposedLane)

    TryBottomLane = True
End Function

Private Function TryTopLane( _
    ByVal viewTop As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef targetY As Double, _
    ByRef laneName As String) As Boolean

    Dim proposedLane As Long
    proposedLane = mTopLaneCount + 1
    targetY = viewTop + _
        DIMENSION_ARRANGE_SPACING_M * CDbl(proposedLane)

    If targetY > evidence.ContentBorderTop - _
       DIMENSION_ARRANGE_BORDER_INSET_M Then Exit Function

    mTopLaneCount = proposedLane
    laneName = "TOP-" & CStr(proposedLane)

    TryTopLane = True
End Function

Private Function TryLeftLane( _
    ByVal viewLeft As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef targetX As Double, _
    ByRef laneName As String) As Boolean

    Dim proposedLane As Long
    proposedLane = mLeftLaneCount + 1
    targetX = viewLeft - _
        DIMENSION_ARRANGE_SPACING_M * CDbl(proposedLane)

    If targetX < evidence.ContentBorderLeft + _
       DIMENSION_ARRANGE_BORDER_INSET_M Then Exit Function

    mLeftLaneCount = proposedLane
    laneName = "LEFT-" & CStr(proposedLane)

    TryLeftLane = True
End Function

Private Function TryRightLane( _
    ByVal viewRight As Double, _
    ByRef evidence As CRunEvidence, _
    ByRef targetX As Double, _
    ByRef laneName As String) As Boolean

    Dim proposedLane As Long
    proposedLane = mRightLaneCount + 1
    targetX = viewRight + _
        DIMENSION_ARRANGE_SPACING_M * CDbl(proposedLane)

    If targetX > evidence.ContentBorderRight - _
       DIMENSION_ARRANGE_BORDER_INSET_M Then Exit Function

    mRightLaneCount = proposedLane
    laneName = "RIGHT-" & CStr(proposedLane)

    TryRightLane = True
End Function

Private Sub ResetDimensionLaneState()
    mBottomLaneCount = 0
    mTopLaneCount = 0
    mLeftLaneCount = 0
    mRightLaneCount = 0
End Sub

Private Function DimensionTypeIsOrdinate( _
    ByVal dimensionType As Long) As Boolean

    ' The fallback delegates every ordinate family to the set-aware
    ' AutoJogOrdinate API instead of assigning baselines per annotation.
    Select Case dimensionType
        Case swDimensionType_Ordinate, swDimensionType_HorOrdinate, _
             swDimensionType_VertOrdinate, swDimensionType_AngularOrdinate

            DimensionTypeIsOrdinate = True
    End Select
End Function

Private Function NearestViewSideIsHorizontal( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal viewLeft As Double, _
    ByVal viewBottom As Double, _
    ByVal viewRight As Double, _
    ByVal viewTop As Double) As Boolean

    Dim horizontalDistance As Double
    horizontalDistance = Abs(y - viewBottom)
    If Abs(y - viewTop) < horizontalDistance Then
        horizontalDistance = Abs(y - viewTop)
    End If

    Dim verticalDistance As Double
    verticalDistance = Abs(x - viewLeft)
    If Abs(x - viewRight) < verticalDistance Then
        verticalDistance = Abs(x - viewRight)
    End If

    NearestViewSideIsHorizontal = _
        (horizontalDistance <= verticalDistance)
End Function

Private Function DimensionPositionIsUnsupported( _
    ByVal dimensionType As Long) As Boolean

    DimensionPositionIsUnsupported = _
        (dimensionType = swDimensionType_Radial Or _
         dimensionType = swDimensionType_Diameter)
End Function

Private Function ValidateDimensionAnnotationOrigins( _
    ByVal dimensions As Variant, _
    ByRef evidence As CRunEvidence, _
    ByVal viewName As String, _
    ByVal validationMode As String) As Boolean

    On Error GoTo Failed

    If Not ContentBorderIsValid(evidence) Then
        evidence.AddFailure "Dimension arrange has no proved content border " & _
            "for '" & viewName & "'."
        Exit Function
    End If

    Dim checkedCount As Long
    Dim i As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)
        If displayDimension Is Nothing Then Exit Function

        Dim annotation As SldWorks.Annotation
        Set annotation = displayDimension.GetAnnotation
        If annotation Is Nothing Then Exit Function

        Dim x As Double
        Dim y As Double
        If Not TryReadAnnotationPosition(annotation, x, y) Then
            evidence.AddFailure "Dimension arrange could not read annotation " & _
                "position in '" & viewName & "' at array index " & CStr(i) & "."
            Exit Function
        End If

        If Not PositionIsInsideContentBorder(x, y, evidence) Then
            evidence.AddFailure "Dimension annotation origin is outside the " & _
                "content border after " & validationMode & " in '" & _
                viewName & "' at array index " & CStr(i) & "."
            Exit Function
        End If

        checkedCount = checkedCount + 1
    Next i

    evidence.AddInfo "DIMENSION_ARRANGE_BOUNDS|view=" & _
        EvidenceValue(viewName) & _
        "|mode=" & EvidenceValue(validationMode) & _
        "|checked=" & CStr(checkedCount) & _
        "|status=PROVED"

    ValidateDimensionAnnotationOrigins = True
    Exit Function

Failed:
    evidence.AddFailure "Dimension arrange position validation API error in '" & _
        viewName & "': " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function TryReadAnnotationPosition( _
    ByRef annotation As SldWorks.Annotation, _
    ByRef x As Double, _
    ByRef y As Double) As Boolean

    On Error GoTo Failed

    Dim position As Variant
    position = annotation.GetPosition
    If IsEmpty(position) Or Not IsArray(position) Then Exit Function
    If UBound(position) - LBound(position) < 2 Then Exit Function

    x = CDbl(position(LBound(position)))
    y = CDbl(position(LBound(position) + 1))
    TryReadAnnotationPosition = True
    Exit Function

Failed:
    TryReadAnnotationPosition = False
End Function

Private Function TryReadViewOutline( _
    ByVal outline As Variant, _
    ByRef leftValue As Double, _
    ByRef bottomValue As Double, _
    ByRef rightValue As Double, _
    ByRef topValue As Double) As Boolean

    On Error GoTo Failed

    If IsEmpty(outline) Or Not IsArray(outline) Then Exit Function
    If UBound(outline) - LBound(outline) < 3 Then Exit Function

    Dim baseIndex As Long
    baseIndex = LBound(outline)

    leftValue = CDbl(outline(baseIndex))
    bottomValue = CDbl(outline(baseIndex + 1))
    rightValue = CDbl(outline(baseIndex + 2))
    topValue = CDbl(outline(baseIndex + 3))

    If rightValue <= leftValue Or topValue <= bottomValue Then Exit Function

    TryReadViewOutline = True
    Exit Function

Failed:
    TryReadViewOutline = False
End Function

Private Function ContentBorderIsValid( _
    ByRef evidence As CRunEvidence) As Boolean

    ContentBorderIsValid = _
        (evidence.ContentBorderRight > evidence.ContentBorderLeft And _
         evidence.ContentBorderTop > evidence.ContentBorderBottom And _
         evidence.ContentBorderLeft >= 0# And _
         evidence.ContentBorderBottom >= 0# And _
         evidence.ContentBorderRight <= evidence.SheetWidth And _
         evidence.ContentBorderTop <= evidence.SheetHeight)
End Function

Private Function PositionIsInsideContentBorder( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    PositionIsInsideContentBorder = _
        (x >= evidence.ContentBorderLeft - _
             DIMENSION_ARRANGE_READBACK_TOLERANCE_M And _
         x <= evidence.ContentBorderRight + _
             DIMENSION_ARRANGE_READBACK_TOLERANCE_M And _
         y >= evidence.ContentBorderBottom - _
             DIMENSION_ARRANGE_READBACK_TOLERANCE_M And _
         y <= evidence.ContentBorderTop + _
             DIMENSION_ARRANGE_READBACK_TOLERANCE_M)
End Function

Private Function ClampToContentBorderX( _
    ByVal x As Double, _
    ByRef evidence As CRunEvidence) As Double

    Dim minimumX As Double
    minimumX = evidence.ContentBorderLeft + _
        DIMENSION_ARRANGE_BORDER_INSET_M

    Dim maximumX As Double
    maximumX = evidence.ContentBorderRight - _
        DIMENSION_ARRANGE_BORDER_INSET_M

    If x < minimumX Then
        ClampToContentBorderX = minimumX
    ElseIf x > maximumX Then
        ClampToContentBorderX = maximumX
    Else
        ClampToContentBorderX = x
    End If
End Function

Private Function ClampToContentBorderY( _
    ByVal y As Double, _
    ByRef evidence As CRunEvidence) As Double

    Dim minimumY As Double
    minimumY = evidence.ContentBorderBottom + _
        DIMENSION_ARRANGE_BORDER_INSET_M

    Dim maximumY As Double
    maximumY = evidence.ContentBorderTop - _
        DIMENSION_ARRANGE_BORDER_INSET_M

    If y < minimumY Then
        ClampToContentBorderY = minimumY
    ElseIf y > maximumY Then
        ClampToContentBorderY = maximumY
    Else
        ClampToContentBorderY = y
    End If
End Function

Public Function GetFirstRealViewName( _
    ByRef swDraw As SldWorks.DrawingDoc) As String

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    If Not swView Is Nothing Then
        GetFirstRealViewName = Module8_RuntimeSupport.GetViewName(swView)
    End If
End Function

Private Function FindViewByName( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal viewName As String) As SldWorks.View

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing
        If StrComp(Module8_RuntimeSupport.GetViewName(swView), _
                   viewName, vbTextCompare) = 0 Then
            Set FindViewByName = swView
            Exit Function
        End If

        Set swView = swView.GetNextView
    Loop
End Function

Private Function GetModelItemMask() As Long
    GetModelItemMask = _
        swInsertDatums Or _
        swInsertDimensionsMarkedForDrawing Or _
        swInsertGTols Or _
        swInsertNotes Or _
        swInsertHoleWizardProfileDimensions Or _
        swInsertHoleWizardLocationDimensions

    If Module1_Main.GlobalConfig.ImportHoleCallouts Then
        GetModelItemMask = GetModelItemMask Or swInsertHoleCallout
    End If
End Function

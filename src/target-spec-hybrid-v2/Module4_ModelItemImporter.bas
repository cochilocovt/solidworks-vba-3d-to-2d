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

Private Const swDimensionType_HorOrdinate As Long = 7
Private Const swDimensionType_VertOrdinate As Long = 8
Private Const swDimensionType_HorLinear As Long = 11
Private Const swDimensionType_VertLinear As Long = 12

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

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        AutoArrangeDimensionsInView swDrawModel, swDraw, swView, evidence
        Set swView = swView.GetNextView
    Loop
End Sub

Private Sub AutoArrangeDimensionsInView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    If Not Module8_RuntimeSupport.ActivateDrawingView( _
        swDrawModel, swDraw, swView, evidence, _
        "Dimension arrange") Then Exit Sub

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If IsEmpty(dimensions) Or Not IsArray(dimensions) Then GoTo SafeExit

    swDrawModel.ClearSelection2 True

    Dim i As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)

        If Not displayDimension Is Nothing Then
            Dim annotation As SldWorks.Annotation
            Set annotation = displayDimension.GetAnnotation

            If Not annotation Is Nothing Then
                annotation.Select3 True, Nothing
            End If
        End If
    Next i

    If Not swDrawModel.Extension.AlignDimensions( _
        swAlignDimensionType_AutoArrange, 0.006) Then

        evidence.AddWarning "AlignDimensions returned False in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
    End If

SafeExit:
    swDrawModel.ClearSelection2 True
    Exit Sub

Failed:
    evidence.AddWarning "Dimension arrange error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
    Resume SafeExit
End Sub

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

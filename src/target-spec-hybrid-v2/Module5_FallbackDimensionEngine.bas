Option Explicit

Private Const PI As Double = 3.14159265358979

Private Const swDocPART As Long = 1
Private Const swSolidBody As Long = 0

Private Const swVerticalOrdinate As Long = 2
Private Const swHorizontalOrdinate As Long = 3

Private Const swCreateOrdDimErr_Success As Long = 0

Public Sub AddMissingOrdinateDimensions( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef modelHoleFeatures As Collection, _
    ByVal requestedOrigin As String, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        If Module8_RuntimeSupport.IsOrdinateEligibleView(swView) Then
            AddMissingOrdinateDimensionsToView _
                swApp, swDrawModel, swDraw, swView, modelHoleFeatures, _
                requestedOrigin, evidence
        Else
            evidence.AddInfo "Ordinate policy skipped unsupported view '" & _
                Module8_RuntimeSupport.GetViewName(swView) & "'."
        End If

        Set swView = swView.GetNextView
    Loop

    Module8_RuntimeSupport.RestoreSheetContext swDrawModel, swDraw
    Exit Sub

Failed:
    evidence.AddFailure "Fallback dimension engine error " & _
        CStr(Err.Number) & ": " & Err.Description
    Module8_RuntimeSupport.RestoreSheetContext swDrawModel, swDraw
End Sub

Private Sub AddMissingOrdinateDimensionsToView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef modelHoleFeatures As Collection, _
    ByVal requestedOrigin As String, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    If Not Module8_RuntimeSupport.ActivateDrawingDocument( _
        swApp, swDrawModel, evidence) Then GoTo SafeExit

    If Not Module8_RuntimeSupport.ActivateDrawingView( _
        swDrawModel, swDraw, swView, evidence, _
        "Ordinate candidate collection") Then GoTo SafeExit

    Dim rejected As Object
    Set rejected = CreateObject("Scripting.Dictionary")

    Dim candidates As Collection
    Set candidates = CollectOwnedCandidates( _
        swApp, swView, modelHoleFeatures, rejected, evidence)

    EmitRejectionSummary swView, rejected, evidence

    If candidates.Count = 0 Then
        evidence.AddInfo "No ownership-proven normal-axis hole centres in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        GoTo SafeExit
    End If

    Dim horizontalOrigin As String
    Dim verticalOrigin As String
    horizontalOrigin = Module1_Main.GetDatumOriginForDirection( _
        evidence.PartPath, requestedOrigin, swHorizontalOrdinate)
    verticalOrigin = Module1_Main.GetDatumOriginForDirection( _
        evidence.PartPath, requestedOrigin, swVerticalOrdinate)

    Dim horizontalDatum As CDatumProof
    Dim verticalDatum As CDatumProof
    Set horizontalDatum = ProveDatum( _
        swApp, swDrawModel, swView, candidates, horizontalOrigin, evidence)
    Set verticalDatum = ProveDatum( _
        swApp, swDrawModel, swView, candidates, verticalOrigin, evidence)

    If Not DatumIsUsable(horizontalDatum) Then
        RecordDatumFailure swView, "X", horizontalOrigin, horizontalDatum, evidence
        GoTo SafeExit
    End If

    If Not DatumIsUsable(verticalDatum) Then
        RecordDatumFailure swView, "Y", verticalOrigin, verticalDatum, evidence
        GoTo SafeExit
    End If

    RecordDatumEvidence swView, "X", horizontalDatum, evidence
    RecordDatumEvidence swView, "Y", verticalDatum, evidence

    If Not Module4_ModelItemImporter.ApplyImportedCoverageToCandidates( _
        swApp, swView, candidates, horizontalDatum, verticalDatum, evidence) Then

        evidence.AddFailure "CoverageInspectionFailed in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & _
            "'; no fallback ordinate was created."
        GoTo SafeExit
    End If

    evidence.HorizontalOrdinateGroups = _
        evidence.HorizontalOrdinateGroups + _
        CreateDirectionOrdinateGroups( _
            swApp, swDrawModel, swDraw, swView, horizontalDatum, _
            candidates, True, evidence)

    evidence.VerticalOrdinateGroups = _
        evidence.VerticalOrdinateGroups + _
        CreateDirectionOrdinateGroups( _
            swApp, swDrawModel, swDraw, swView, verticalDatum, _
            candidates, False, evidence)

SafeExit:
    Module8_RuntimeSupport.RestoreSheetContext swDrawModel, swDraw
    Exit Sub

Failed:
    evidence.AddFailure "Ordinate view error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
    Resume SafeExit
End Sub

Private Function CollectOwnedCandidates( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef modelHoleFeatures As Collection, _
    ByRef rejected As Object, _
    ByRef evidence As CRunEvidence) As Collection

    Dim results As New Collection
    Set CollectOwnedCandidates = results

    On Error GoTo Failed

    Dim componentFailure As String
    Dim component As SldWorks.Component2
    Set component = GetSingleVisibleComponent(swView, componentFailure)

    If component Is Nothing Then
        RecordRejection rejected, componentFailure
        Exit Function
    End If

    Dim referencedModel As SldWorks.ModelDoc2
    Set referencedModel = swView.ReferencedDocument

    If referencedModel Is Nothing Then
        RecordRejection rejected, "ReferencedDocumentUnavailable"
        Exit Function
    End If

    If referencedModel.GetType <> swDocPART Then
        RecordRejection rejected, "ReferencedDocumentIsNotPart"
        Exit Function
    End If

    If modelHoleFeatures Is Nothing Then
        RecordRejection rejected, "ModelHoleFeatureAuditUnavailable"
        Exit Function
    End If

    Dim featureCount As Long
    Dim faceCount As Long
    Dim cylinderFaceCount As Long
    Dim circularEdgeCount As Long
    Dim mappedEdgeCount As Long

    Dim featureIndex As Long
    For featureIndex = 1 To modelHoleFeatures.Count
        Dim ownerFeature As SldWorks.Feature
        Set ownerFeature = modelHoleFeatures(featureIndex)

        If Not ownerFeature Is Nothing Then
            featureCount = featureCount + 1

            Dim featureFaces As Variant
            featureFaces = ownerFeature.GetFaces

            If IsArray(featureFaces) Then
                Dim faceIndex As Long
                For faceIndex = LBound(featureFaces) To UBound(featureFaces)
                    Dim cylinderFace As SldWorks.Face2
                    Set cylinderFace = featureFaces(faceIndex)

                    If Not cylinderFace Is Nothing Then
                        faceCount = faceCount + 1

                        If Module3_ModelAudit.IsInternalCylindricalFace( _
                            cylinderFace) Then

                            cylinderFaceCount = cylinderFaceCount + 1

                            Dim modelEdges As Variant
                            modelEdges = cylinderFace.GetEdges

                            If IsArray(modelEdges) Then
                                Dim edgeIndex As Long
                                For edgeIndex = LBound(modelEdges) To UBound(modelEdges)
                                    Dim modelEdge As SldWorks.Edge
                                    Set modelEdge = modelEdges(edgeIndex)

                                    If ModelEdgeIsCircular(modelEdge) Then
                                        circularEdgeCount = circularEdgeCount + 1

                                        Dim drawingEdge As SldWorks.Edge
                                        Set drawingEdge = _
                                            GetDrawingEdgeForModelEdge( _
                                                swView, modelEdge)

                                        If drawingEdge Is Nothing Then
                                            RecordRejection rejected, _
                                                "NoDrawingViewCorrespondence"
                                        Else
                                            mappedEdgeCount = mappedEdgeCount + 1

                                            Dim candidate As CHoleCandidate
                                            Set candidate = BuildOwnedCandidate( _
                                                swApp, swView, component, _
                                                drawingEdge, modelEdge, _
                                                cylinderFace, ownerFeature)

                                            If candidate.Accepted Then
                                                AddOrMergeCandidate results, candidate
                                            Else
                                                RecordRejection rejected, _
                                                    candidate.RejectionReason
                                            End If
                                        End If
                                    End If
                                Next edgeIndex
                            End If
                        End If
                    End If
                Next faceIndex
            End If
        End If
    Next featureIndex

    evidence.AddInfo "ENTITY_SOURCE|view=" & _
        EvidenceValue(Module8_RuntimeSupport.GetViewName(swView)) & _
        "|component=" & EvidenceValue(component.Name2) & _
        "|source=AuditedFeature.GetFaces/Face2.GetEdges" & _
        "|mapping=IView.GetCorrespondingEntity" & _
        "|features=" & CStr(featureCount) & _
        "|faces=" & CStr(faceCount) & _
        "|internalCylinders=" & CStr(cylinderFaceCount) & _
        "|circularEdges=" & CStr(circularEdgeCount) & _
        "|mappedEdges=" & CStr(mappedEdgeCount)

    evidence.CandidatesAccepted = _
        evidence.CandidatesAccepted + results.Count

    Dim acceptedIndex As Long
    For acceptedIndex = 1 To results.Count
        Dim acceptedCandidate As CHoleCandidate
        Set acceptedCandidate = results(acceptedIndex)

        If Not evidence.RegisterPhysicalLocation( _
            acceptedCandidate.PhysicalInstanceKey, _
            acceptedCandidate.ComponentName, _
            acceptedCandidate.FeatureName, _
            acceptedCandidate.ConfigurationName, _
            BuildCanonicalModelCentreKey(acceptedCandidate), _
            BuildCanonicalModelAxisKey(acceptedCandidate), _
            acceptedCandidate.ViewName) Then

            evidence.AddFailure _
                "Physical-location registration failed for candidate '" & _
                acceptedCandidate.PhysicalInstanceKey & "'."
        End If

        evidence.AddProjectionLocationRecord _
            acceptedCandidate.PhysicalInstanceKey, _
            acceptedCandidate.ViewName, _
            acceptedCandidate.ViewX, _
            acceptedCandidate.ViewY, _
            "circular-edge-centre"

        evidence.AddInfo "Accepted candidate in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "': " & _
            acceptedCandidate.ComponentName & "/" & _
            acceptedCandidate.FeatureName & _
            "; type=" & acceptedCandidate.FeatureType & _
            "; config=" & acceptedCandidate.ConfigurationName & _
            "; ownership=" & acceptedCandidate.OwnershipProof & _
            "; viewXY=" & Format$(acceptedCandidate.ViewX, "0.000000") & "," & _
            Format$(acceptedCandidate.ViewY, "0.000000") & _
            "; modelAxis=" & Format$(acceptedCandidate.ModelAxisX, "0.000000") & _
            "," & Format$(acceptedCandidate.ModelAxisY, "0.000000") & _
            "," & Format$(acceptedCandidate.ModelAxisZ, "0.000000") & _
            "; axisZ=" & Format$(acceptedCandidate.AxisZ, "0.000000") & _
            "; radii=" & acceptedCandidate.RadiiStack & _
            "; semantics=" & acceptedCandidate.SemanticsSummary & _
            "; familyKey=" & acceptedCandidate.FamilyKey & _
            "; instanceKey=" & acceptedCandidate.PhysicalInstanceKey & "."

        evidence.AddInfo "EVIDENCE|CANDIDATE_ACCEPTED|view=" & _
            EvidenceValue(acceptedCandidate.ViewName) & _
            "|family=" & EvidenceValue(acceptedCandidate.FamilyKey) & _
            "|instance=" & EvidenceValue(acceptedCandidate.PhysicalInstanceKey) & _
            "|ownership=" & EvidenceValue(acceptedCandidate.OwnershipProof) & _
            "|config=" & EvidenceValue(acceptedCandidate.ConfigurationName)
    Next acceptedIndex

    Dim rejectedCount As Long
    rejectedCount = DictionaryTotal(rejected)
    evidence.CandidatesRejected = evidence.CandidatesRejected + rejectedCount
    Exit Function

Failed:
    evidence.AddFailure "Candidate collection error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
End Function

Private Function GetSingleVisibleComponent( _
    ByRef swView As SldWorks.View, _
    ByRef failureReason As String) As SldWorks.Component2

    On Error GoTo Failed

    Dim visibleComponents As Variant
    visibleComponents = swView.GetVisibleComponents

    If IsEmpty(visibleComponents) Or Not IsArray(visibleComponents) Then
        failureReason = "NoVisibleComponents"
        Exit Function
    End If

    Dim componentCount As Long
    Dim i As Long
    For i = LBound(visibleComponents) To UBound(visibleComponents)
        Dim candidateComponent As SldWorks.Component2
        Set candidateComponent = visibleComponents(i)

        If Not candidateComponent Is Nothing Then
            componentCount = componentCount + 1
            If componentCount = 1 Then
                Set GetSingleVisibleComponent = candidateComponent
            End If
        End If
    Next i

    If componentCount = 1 Then Exit Function

    Set GetSingleVisibleComponent = Nothing
    If componentCount = 0 Then
        failureReason = "NoVisibleComponents"
    Else
        failureReason = "MultipleVisibleComponentsUnsupported"
    End If
    Exit Function

Failed:
    Set GetSingleVisibleComponent = Nothing
    failureReason = "VisibleComponentReadError:" & CStr(Err.Number)
End Function

Private Function ModelEdgeIsCircular( _
    ByRef modelEdge As SldWorks.Edge) As Boolean

    If modelEdge Is Nothing Then Exit Function
    On Error GoTo Failed

    Dim curve As SldWorks.Curve
    Set curve = modelEdge.GetCurve
    If curve Is Nothing Then Exit Function

    ModelEdgeIsCircular = curve.IsCircle
    Exit Function

Failed:
    ModelEdgeIsCircular = False
End Function

Private Function GetDrawingEdgeForModelEdge( _
    ByRef swView As SldWorks.View, _
    ByRef modelEdge As SldWorks.Edge) As SldWorks.Edge

    If swView Is Nothing Or modelEdge Is Nothing Then Exit Function
    On Error GoTo Failed

    Dim mappedObject As Object
    Set mappedObject = swView.GetCorrespondingEntity(modelEdge)
    If mappedObject Is Nothing Then Exit Function

    Set GetDrawingEdgeForModelEdge = mappedObject
    Exit Function

Failed:
    Set GetDrawingEdgeForModelEdge = Nothing
End Function

Private Function BuildOwnedCandidate( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByRef drawingEdge As SldWorks.Edge, _
    ByRef modelEdge As SldWorks.Edge, _
    ByRef cylinderFace As SldWorks.Face2, _
    ByRef ownerFeature As SldWorks.Feature) As CHoleCandidate

    Dim candidate As New CHoleCandidate
    Set BuildOwnedCandidate = candidate

    Set candidate.DrawingEdge = drawingEdge
    candidate.DrawingEntityAliases.Add drawingEdge
    Set candidate.Component = component
    Set candidate.DrawingView = swView
    candidate.ViewName = Module8_RuntimeSupport.GetViewName(swView)
    Set candidate.ModelEdge = modelEdge
    candidate.ModelEntityAliases.Add modelEdge
    Set candidate.OwnerFeature = ownerFeature
    candidate.OwnershipProof = "AuditedFeatureFaceSet"

    On Error GoTo RejectedByError

    candidate.ComponentName = component.Name2
    candidate.ConfigurationName = Trim$(swView.ReferencedConfiguration)

    If Len(candidate.ConfigurationName) = 0 Then
        candidate.RejectionReason = "ReferencedConfigurationUnavailable"
        Exit Function
    End If

    If modelEdge Is Nothing Then
        candidate.RejectionReason = "ModelEdgeUnavailable"
        Exit Function
    End If

    If cylinderFace Is Nothing Then
        candidate.RejectionReason = "CylinderFaceUnavailable"
        Exit Function
    End If

    If ownerFeature Is Nothing Then
        candidate.RejectionReason = "OwnerFeatureUnavailable"
        Exit Function
    End If

    If Not Module3_ModelAudit.FeatureContainsFace( _
        ownerFeature, cylinderFace) Then

        candidate.RejectionReason = "AuditedFeatureDoesNotContainCylinderFace"
        Exit Function
    End If

    If Not Module3_ModelAudit.IsInternalCylindricalFace(cylinderFace) Then
        candidate.RejectionReason = "CylinderFaceIsNotInternal"
        Exit Function
    End If

    Dim curve As SldWorks.Curve
    Set curve = modelEdge.GetCurve

    If curve Is Nothing Then
        candidate.RejectionReason = "NoUnderlyingCurve"
        Exit Function
    End If

    If Not curve.IsCircle Then
        candidate.RejectionReason = "NotCircular"
        Exit Function
    End If

    Dim parameterData As Object
    Set parameterData = modelEdge.GetCurveParams3

    If parameterData Is Nothing Then
        candidate.RejectionReason = "NoCurveParams3"
        Exit Function
    End If

    If Not IsClosedCurveParameterization(parameterData) Then
        candidate.RejectionReason = "CircularArcNotFullCircle"
        Exit Function
    End If

    Dim circleData As Variant
    circleData = curve.CircleParams

    If Not IsArray(circleData) Then
        candidate.RejectionReason = "NoCircleParameters"
        Exit Function
    End If

    candidate.ModelX = CDbl(circleData(0))
    candidate.ModelY = CDbl(circleData(1))
    candidate.ModelZ = CDbl(circleData(2))
    candidate.Radius = Abs(CDbl(circleData(6)))

    If Not CylinderFaceMatchesRadius(cylinderFace, candidate.Radius) Then
        candidate.RejectionReason = "CylinderRadiusDoesNotMatchCircularEdge"
        Exit Function
    End If

    If Not Module3_ModelAudit.IsOwnedHoleFeature(ownerFeature) Then
        candidate.RejectionReason = "RejectedFeatureType:" & _
            Module3_ModelAudit.DescribeFeatureType(ownerFeature)
        Exit Function
    End If

    If Not Module3_ModelAudit.IsFeatureActiveInConfiguration( _
        ownerFeature, candidate.ConfigurationName) Then

        candidate.RejectionReason = _
            "FeatureSuppressedOrUnprovedInReferencedConfiguration"
        Exit Function
    End If

    candidate.FeatureName = ownerFeature.Name
    candidate.FeatureType = Module3_ModelAudit.DescribeFeatureType(ownerFeature)

    Dim featureTypeUpper As String
    featureTypeUpper = UCase$(candidate.FeatureType)

    If featureTypeUpper <> "HOLEWZD" And _
       featureTypeUpper <> "ADVHOLEWZD" And _
       featureTypeUpper <> "SKETCHHOLE" Then

        If InStr(featureTypeUpper, "CUT") = 0 Or _
           Not Module3_ModelAudit.IsInternalCylindricalFace(cylinderFace) Then

            candidate.RejectionReason = "GenericCutMatchedCylinderNotProvenInternal"
            Exit Function
        End If
    End If

    Dim modelAxisMagnitude As Double
    modelAxisMagnitude = Sqr(CDbl(circleData(3)) * CDbl(circleData(3)) + _
                             CDbl(circleData(4)) * CDbl(circleData(4)) + _
                             CDbl(circleData(5)) * CDbl(circleData(5)))

    If modelAxisMagnitude <= Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then
        candidate.RejectionReason = "DegenerateModelAxis"
        Exit Function
    End If

    candidate.ModelAxisX = CDbl(circleData(3)) / modelAxisMagnitude
    candidate.ModelAxisY = CDbl(circleData(4)) / modelAxisMagnitude
    candidate.ModelAxisZ = CDbl(circleData(5)) / modelAxisMagnitude
    CanonicalizeAxis candidate.ModelAxisX, _
        candidate.ModelAxisY, candidate.ModelAxisZ

    Dim transformedAxisX As Double
    Dim transformedAxisY As Double
    Dim transformedAxisZ As Double

    If Not Module8_RuntimeSupport.TransformVectorToView( _
        swApp, swView, candidate.ModelAxisX, candidate.ModelAxisY, _
        candidate.ModelAxisZ, transformedAxisX, transformedAxisY, _
        transformedAxisZ) Then

        candidate.RejectionReason = "AxisTransformFailed"
        Exit Function
    End If

    Dim axisMagnitude As Double
    axisMagnitude = Sqr(transformedAxisX * transformedAxisX + _
                        transformedAxisY * transformedAxisY + _
                        transformedAxisZ * transformedAxisZ)

    If axisMagnitude <= Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then
        candidate.RejectionReason = "DegenerateAxis"
        Exit Function
    End If

    candidate.AxisX = transformedAxisX / axisMagnitude
    candidate.AxisY = transformedAxisY / axisMagnitude
    candidate.AxisZ = transformedAxisZ / axisMagnitude

    If Abs(candidate.AxisZ) < Module8_RuntimeSupport.AXIS_NORMAL_MIN_COS Then
        candidate.RejectionReason = "AxisNotNormalToView"
        Exit Function
    End If

    Dim viewZ As Double
    If Not Module8_RuntimeSupport.TransformPointToView( _
        swApp, swView, candidate.ModelX, candidate.ModelY, candidate.ModelZ, _
        candidate.ViewX, candidate.ViewY, viewZ) Then

        candidate.RejectionReason = "CentreTransformFailed"
        Exit Function
    End If

    Module8_RuntimeSupport.ViewToSheetCoordinates _
        swView, candidate.ViewX, candidate.ViewY, _
        candidate.SheetX, candidate.SheetY

    CaptureFeatureSemantics candidate, cylinderFace
    candidate.FamilyKey = BuildFamilyKey(candidate)
    candidate.PhysicalInstanceKey = BuildPhysicalInstanceKey(candidate)
    candidate.SemanticKey = BuildSemanticKey(candidate)
    candidate.Accepted = True
    Exit Function

RejectedByError:
    candidate.Accepted = False
    candidate.RejectionReason = "CandidateError:" & CStr(Err.Number)
End Function

Private Sub CaptureFeatureSemantics( _
    ByRef candidate As CHoleCandidate, _
    ByRef cylinderFace As SldWorks.Face2)

    candidate.RadiiStack = Format$(candidate.Radius, "0.000000")

    If cylinderFace.FaceInSurfaceSense Then
        candidate.MachiningSide = "internal-cylinder"
    Else
        candidate.MachiningSide = "feature-owned-cylinder"
    End If

    candidate.SemanticsSummary = _
        "feature=" & candidate.FeatureType & _
        ";side=" & candidate.MachiningSide

    On Error GoTo SafeExit

    Dim definition As Object
    Set definition = candidate.OwnerFeature.GetDefinition
    If definition Is Nothing Then GoTo SafeExit

    Select Case UCase$(candidate.FeatureType)
        Case "HOLEWZD"
            AppendSemanticProperty candidate, definition, "HoleDiameter"
            AppendSemanticProperty candidate, definition, "HoleDepth"
            AppendSemanticProperty candidate, definition, "EndCondition"
            AppendSemanticProperty candidate, definition, "CounterBoreDiameter"
            AppendSemanticProperty candidate, definition, "CounterBoreDepth"
            AppendSemanticProperty candidate, definition, "CounterSinkDiameter"
            AppendSemanticProperty candidate, definition, "CounterSinkAngle"
            AppendSemanticProperty candidate, definition, "ThreadDiameter"
            AppendSemanticProperty candidate, definition, "ThreadDepth"
            AppendSemanticProperty candidate, definition, "ThreadEndCondition"
            AppendSemanticProperty candidate, definition, "FastenerSize"
            AppendSemanticProperty candidate, definition, "Standard2"

        Case "SKETCHHOLE"
            AppendSemanticProperty candidate, definition, "Diameter"
            AppendSemanticProperty candidate, definition, "Depth"
            AppendSemanticProperty candidate, definition, "Type"

        Case "ADVHOLEWZD"
            AppendSemanticProperty candidate, definition, "CustomizeCallout"
            AppendSemanticProperty candidate, definition, "NearSideElementsCount"
            AppendSemanticProperty candidate, definition, "FarSideElementsCount"
    End Select

SafeExit:
End Sub

Private Sub AppendSemanticProperty( _
    ByRef candidate As CHoleCandidate, _
    ByVal definition As Object, _
    ByVal propertyName As String)

    Dim value As String
    value = ReadDefinitionProperty(definition, propertyName)

    If Len(value) > 0 Then
        candidate.SemanticsSummary = candidate.SemanticsSummary & _
            ";" & propertyName & "=" & value
    End If
End Sub

Private Function ReadDefinitionProperty( _
    ByVal definition As Object, _
    ByVal propertyName As String) As String

    On Error GoTo Failed

    Dim value As Variant
    value = CallByName(definition, propertyName, VbGet)

    If IsEmpty(value) Or IsNull(value) Then Exit Function
    If IsObject(value) Or IsArray(value) Then Exit Function

    If VarType(value) = vbDouble Or VarType(value) = vbSingle Or _
       VarType(value) = vbDecimal Then

        ReadDefinitionProperty = Format$(CDbl(value), "0.000000")
    Else
        ReadDefinitionProperty = CStr(value)
    End If
    Exit Function

Failed:
    ReadDefinitionProperty = vbNullString
End Function

Private Function IsClosedCurveParameterization( _
    ByVal parameterData As Object) As Boolean

    On Error GoTo Failed

    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = parameterData.StartPoint
    endPoint = parameterData.EndPoint

    Dim distance As Double
    distance = Sqr( _
        (CDbl(startPoint(0)) - CDbl(endPoint(0))) ^ 2 + _
        (CDbl(startPoint(1)) - CDbl(endPoint(1))) ^ 2 + _
        (CDbl(startPoint(2)) - CDbl(endPoint(2))) ^ 2)

    If distance > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then Exit Function

    If Abs(CDbl(parameterData.UMaxValue) - _
           CDbl(parameterData.UMinValue)) <= _
           Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then Exit Function

    IsClosedCurveParameterization = True
    Exit Function

Failed:
    IsClosedCurveParameterization = False
End Function

Private Function CylinderFaceMatchesRadius( _
    ByRef cylinderFace As SldWorks.Face2, _
    ByVal circleRadius As Double) As Boolean

    On Error GoTo Failed

    If cylinderFace Is Nothing Then Exit Function

    Dim surface As SldWorks.Surface
    Set surface = cylinderFace.GetSurface

    If surface Is Nothing Then Exit Function
    If Not surface.IsCylinder Then Exit Function

    Dim cylinderData As Variant
    cylinderData = surface.CylinderParams
    If Not IsArray(cylinderData) Then Exit Function

    CylinderFaceMatchesRadius = _
        (Abs(Abs(CDbl(cylinderData(6))) - circleRadius) <= _
         Module8_RuntimeSupport.PROJECTED_TOLERANCE_M)
    Exit Function

Failed:
    CylinderFaceMatchesRadius = False
End Function

Private Sub AddOrMergeCandidate( _
    ByRef results As Collection, _
    ByRef candidate As CHoleCandidate)

    Dim i As Long
    For i = 1 To results.Count
        Dim existing As CHoleCandidate
        Set existing = results(i)

        If IsSameSemanticInstance(existing, candidate) Then
            existing.DrawingEntityAliases.Add candidate.DrawingEdge
            existing.ModelEntityAliases.Add candidate.ModelEdge

            If candidate.Radius < existing.Radius Then
                Set existing.DrawingEdge = candidate.DrawingEdge
                Set existing.ModelEdge = candidate.ModelEdge
                existing.Radius = candidate.Radius
                existing.SemanticKey = candidate.SemanticKey
            End If

            If InStr(1, existing.RadiiStack, _
                Format$(candidate.Radius, "0.000000"), vbTextCompare) = 0 Then

                existing.RadiiStack = existing.RadiiStack & "," & _
                    Format$(candidate.Radius, "0.000000")
            End If
            Exit Sub
        End If
    Next i

    results.Add candidate
End Sub

Private Function IsSameSemanticInstance( _
    ByRef firstCandidate As CHoleCandidate, _
    ByRef secondCandidate As CHoleCandidate) As Boolean

    If Len(firstCandidate.PhysicalInstanceKey) > 0 And _
       Len(secondCandidate.PhysicalInstanceKey) > 0 Then

        IsSameSemanticInstance = _
            (StrComp(firstCandidate.PhysicalInstanceKey, _
                     secondCandidate.PhysicalInstanceKey, _
                     vbTextCompare) = 0)
        Exit Function
    End If

    If StrComp(firstCandidate.ComponentName, _
               secondCandidate.ComponentName, vbTextCompare) <> 0 Then Exit Function

    If StrComp(firstCandidate.ConfigurationName, _
               secondCandidate.ConfigurationName, vbTextCompare) <> 0 Then Exit Function

    If StrComp(firstCandidate.FeatureName, _
               secondCandidate.FeatureName, vbTextCompare) <> 0 Then Exit Function

    If Abs(firstCandidate.ViewX - secondCandidate.ViewX) > _
       Module8_RuntimeSupport.PROJECTED_TOLERANCE_M Then Exit Function

    If Abs(firstCandidate.ViewY - secondCandidate.ViewY) > _
       Module8_RuntimeSupport.PROJECTED_TOLERANCE_M Then Exit Function

    Dim axisDot As Double
    axisDot = firstCandidate.AxisX * secondCandidate.AxisX + _
              firstCandidate.AxisY * secondCandidate.AxisY + _
              firstCandidate.AxisZ * secondCandidate.AxisZ

    If Abs(axisDot) < Module8_RuntimeSupport.AXIS_NORMAL_MIN_COS Then Exit Function

    IsSameSemanticInstance = True
End Function

Private Function BuildSemanticKey( _
    ByRef candidate As CHoleCandidate) As String

    BuildSemanticKey = candidate.PhysicalInstanceKey & _
        "|r=" & QuantizedValue(candidate.Radius, _
                                 Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M)
End Function

Private Function BuildFamilyKey( _
    ByRef candidate As CHoleCandidate) As String

    BuildFamilyKey = _
        "cfg=" & LCase$(candidate.ConfigurationName) & _
        "|component=" & LCase$(candidate.ComponentName) & _
        "|componentId=" & SafeComponentId(candidate.Component) & _
        "|feature=" & LCase$(candidate.FeatureName) & _
        "|featureId=" & SafeFeatureId(candidate.OwnerFeature) & _
        "|type=" & LCase$(candidate.FeatureType)
End Function

Private Function BuildPhysicalInstanceKey( _
    ByRef candidate As CHoleCandidate) As String

    Dim momentX As Double
    Dim momentY As Double
    Dim momentZ As Double
    momentX = candidate.ModelY * candidate.ModelAxisZ - _
              candidate.ModelZ * candidate.ModelAxisY
    momentY = candidate.ModelZ * candidate.ModelAxisX - _
              candidate.ModelX * candidate.ModelAxisZ
    momentZ = candidate.ModelX * candidate.ModelAxisY - _
              candidate.ModelY * candidate.ModelAxisX

    BuildPhysicalInstanceKey = candidate.FamilyKey & _
        "|axis=" & QuantizedValue(candidate.ModelAxisX, _
                                    Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M) & _
        "," & QuantizedValue(candidate.ModelAxisY, _
                              Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M) & _
        "," & QuantizedValue(candidate.ModelAxisZ, _
                              Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M) & _
        "|moment=" & QuantizedValue(momentX, _
                                      Module8_RuntimeSupport.PROJECTED_TOLERANCE_M) & _
        "," & QuantizedValue(momentY, _
                              Module8_RuntimeSupport.PROJECTED_TOLERANCE_M) & _
        "," & QuantizedValue(momentZ, _
                               Module8_RuntimeSupport.PROJECTED_TOLERANCE_M)
End Function

Private Function BuildCanonicalModelCentreKey( _
    ByRef candidate As CHoleCandidate) As String

    Dim momentX As Double
    Dim momentY As Double
    Dim momentZ As Double
    momentX = candidate.ModelY * candidate.ModelAxisZ - _
              candidate.ModelZ * candidate.ModelAxisY
    momentY = candidate.ModelZ * candidate.ModelAxisX - _
              candidate.ModelX * candidate.ModelAxisZ
    momentZ = candidate.ModelX * candidate.ModelAxisY - _
              candidate.ModelY * candidate.ModelAxisX

    Dim canonicalX As Double
    Dim canonicalY As Double
    Dim canonicalZ As Double
    canonicalX = candidate.ModelAxisY * momentZ - _
                 candidate.ModelAxisZ * momentY
    canonicalY = candidate.ModelAxisZ * momentX - _
                 candidate.ModelAxisX * momentZ
    canonicalZ = candidate.ModelAxisX * momentY - _
                 candidate.ModelAxisY * momentX

    BuildCanonicalModelCentreKey = _
        QuantizedValue(canonicalX, _
                       Module8_RuntimeSupport.PROJECTED_TOLERANCE_M) & "," & _
        QuantizedValue(canonicalY, _
                       Module8_RuntimeSupport.PROJECTED_TOLERANCE_M) & "," & _
        QuantizedValue(canonicalZ, _
                       Module8_RuntimeSupport.PROJECTED_TOLERANCE_M)
End Function

Private Function BuildCanonicalModelAxisKey( _
    ByRef candidate As CHoleCandidate) As String

    BuildCanonicalModelAxisKey = _
        QuantizedValue(candidate.ModelAxisX, _
                       Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M) & "," & _
        QuantizedValue(candidate.ModelAxisY, _
                       Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M) & "," & _
        QuantizedValue(candidate.ModelAxisZ, _
                       Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M)
End Function

Private Function SafeComponentId( _
    ByRef component As SldWorks.Component2) As String

    If component Is Nothing Then
        SafeComponentId = "none"
        Exit Function
    End If

    On Error GoTo Failed
    SafeComponentId = CStr(component.GetID)
    Exit Function

Failed:
    SafeComponentId = "unavailable"
End Function

Private Function SafeFeatureId( _
    ByRef swFeature As SldWorks.Feature) As String

    If swFeature Is Nothing Then
        SafeFeatureId = "none"
        Exit Function
    End If

    On Error GoTo Failed
    SafeFeatureId = CStr(swFeature.GetID)
    Exit Function

Failed:
    SafeFeatureId = "unavailable"
End Function

Private Function QuantizedValue( _
    ByVal value As Double, _
    ByVal tolerance As Double) As String

    If tolerance <= 0# Then
        QuantizedValue = Format$(value, "0.000000000")
        Exit Function
    End If

    Dim scaled As Double
    scaled = value / tolerance

    If scaled >= 0# Then
        QuantizedValue = Format$(Fix(scaled + 0.5), "0")
    Else
        QuantizedValue = Format$(Fix(scaled - 0.5), "0")
    End If
End Function

Private Sub CanonicalizeAxis( _
    ByRef axisX As Double, _
    ByRef axisY As Double, _
    ByRef axisZ As Double)

    Dim reverseAxis As Boolean
    If Abs(axisX) > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then
        reverseAxis = (axisX < 0#)
    ElseIf Abs(axisY) > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then
        reverseAxis = (axisY < 0#)
    Else
        reverseAxis = (axisZ < 0#)
    End If

    If reverseAxis Then
        axisX = -axisX
        axisY = -axisY
        axisZ = -axisZ
    End If
End Sub

Private Function ProveDatum( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef candidates As Collection, _
    ByVal requestedOrigin As String, _
    ByRef evidence As CRunEvidence) As CDatumProof

    Dim proof As New CDatumProof
    Set ProveDatum = proof

    Set proof.DrawingView = swView
    proof.RequestedOrigin = requestedOrigin
    proof.ViewName = Module8_RuntimeSupport.GetViewName(swView)

    On Error Resume Next
    proof.ConfigurationName = Trim$(swView.ReferencedConfiguration)
    On Error GoTo 0

    If Len(proof.ConfigurationName) = 0 Then
        proof.FailureReason = _
            "Drawing view referenced configuration is unavailable."
        Exit Function
    End If

    Select Case UCase$(Trim$(requestedOrigin))
        Case "BOTTOM-LEFT"
            ProveCornerDatum swApp, swDrawModel, swView, False, proof

        Case "TOP-LEFT"
            ProveCornerDatum swApp, swDrawModel, swView, True, proof

        Case "CENTER"
            ProveCenterDatum swApp, swDrawModel, swView, candidates, proof

        Case Else
            proof.FailureReason = "Unsupported datum option '" & requestedOrigin & "'."
    End Select
End Function

Private Function DatumIsUsable( _
    ByRef datum As CDatumProof) As Boolean

    If datum Is Nothing Then Exit Function
    If Not datum.Proven Then Exit Function
    If datum.DrawingEntity Is Nothing Then Exit Function
    If datum.ModelEntity Is Nothing Then Exit Function
    If datum.Component Is Nothing Then Exit Function
    If Len(datum.ConfigurationName) = 0 Then Exit Function
    If Len(datum.EntityKind) = 0 Then Exit Function
    If Len(datum.StableKey) = 0 Then Exit Function
    DatumIsUsable = True
End Function

Private Sub RecordDatumFailure( _
    ByRef swView As SldWorks.View, _
    ByVal directionName As String, _
    ByVal requestedOrigin As String, _
    ByRef datum As CDatumProof, _
    ByRef evidence As CRunEvidence)

    Dim reason As String
    reason = "Datum proof returned Nothing."
    If Not datum Is Nothing Then
        If Len(datum.FailureReason) > 0 Then reason = datum.FailureReason
    End If

    evidence.AddFailure "DatumNotSelectable in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & _
        "' for " & directionName & " direction and origin '" & _
        requestedOrigin & "': " & reason
    evidence.AddInfo "EVIDENCE|DATUM_PROOF|view=" & _
        EvidenceValue(Module8_RuntimeSupport.GetViewName(swView)) & _
        "|direction=" & directionName & _
        "|requested=" & EvidenceValue(requestedOrigin) & _
        "|proven=False|reason=" & EvidenceValue(reason)
End Sub

Private Sub RecordDatumEvidence( _
    ByRef swView As SldWorks.View, _
    ByVal directionName As String, _
    ByRef datum As CDatumProof, _
    ByRef evidence As CRunEvidence)

    evidence.AddInfo "Datum proof in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "' for " & _
        directionName & ": " & datum.ProofSource & _
        "; entity=" & datum.EntityKind & _
        "; modelXYZ=" & Format$(datum.ModelX, "0.000000") & "," & _
        Format$(datum.ModelY, "0.000000") & "," & _
        Format$(datum.ModelZ, "0.000000") & _
        "; viewXY=" & Format$(datum.ViewX, "0.000000") & "," & _
        Format$(datum.ViewY, "0.000000") & _
        "; sheetXY=" & Format$(datum.SheetX, "0.000000") & "," & _
        Format$(datum.SheetY, "0.000000") & _
        "; key=" & datum.StableKey & "."

    evidence.AddInfo "EVIDENCE|DATUM_PROOF|view=" & _
        EvidenceValue(datum.ViewName) & _
        "|direction=" & directionName & _
        "|requested=" & EvidenceValue(datum.RequestedOrigin) & _
        "|proven=True|entity=" & EvidenceValue(datum.EntityKind) & _
        "|component=" & EvidenceValue(datum.ComponentName) & _
        "|modelX=" & Format$(datum.ModelX, "0.000000000") & _
        "|modelY=" & Format$(datum.ModelY, "0.000000000") & _
        "|modelZ=" & Format$(datum.ModelZ, "0.000000000") & _
        "|viewX=" & Format$(datum.ViewX, "0.000000000") & _
        "|viewY=" & Format$(datum.ViewY, "0.000000000") & _
        "|sheetX=" & Format$(datum.SheetX, "0.000000000") & _
        "|sheetY=" & Format$(datum.SheetY, "0.000000000") & _
        "|key=" & EvidenceValue(datum.StableKey)
End Sub

Private Function BuildDatumStableKey( _
    ByRef proof As CDatumProof) As String

    BuildDatumStableKey = _
        "cfg=" & LCase$(proof.ConfigurationName) & _
        "|origin=" & LCase$(proof.RequestedOrigin) & _
        "|component=" & LCase$(proof.ComponentName) & _
        "|componentId=" & SafeComponentId(proof.Component) & _
        "|entity=" & LCase$(proof.EntityKind) & _
        "|model=" & QuantizedValue(proof.ModelX, _
                                    Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M) & _
        "," & QuantizedValue(proof.ModelY, _
                              Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M) & _
        "," & QuantizedValue(proof.ModelZ, _
                              Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M)
End Function

Private Function AsEntity( _
    ByVal sourceObject As Object) As SldWorks.Entity

    If sourceObject Is Nothing Then Exit Function
    On Error Resume Next
    Set AsEntity = sourceObject
    On Error GoTo 0
End Function

Private Function EvidenceValue(ByVal value As String) As String
    EvidenceValue = Replace$(Trim$(value), "|", "/")
    EvidenceValue = Replace$(EvidenceValue, vbCr, " ")
    EvidenceValue = Replace$(EvidenceValue, vbLf, " ")
End Function

Private Sub ProveCornerDatum( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByVal useTop As Boolean, _
    ByRef proof As CDatumProof)

    On Error GoTo Failed

    Dim entities As New Collection
    Dim modelEntities As New Collection
    Dim components As New Collection
    Dim modelXValues As New Collection
    Dim modelYValues As New Collection
    Dim modelZValues As New Collection
    Dim xValues As New Collection
    Dim yValues As New Collection

    CollectModelFirstVertexCoordinates _
        swApp, swView, entities, modelEntities, components, _
        modelXValues, modelYValues, modelZValues, xValues, yValues

    If entities.Count = 0 Then
        proof.FailureReason = "No corresponding visible model vertices."
        Exit Sub
    End If

    Dim minimumX As Double
    Dim targetY As Double
    minimumX = CDbl(xValues(1))
    targetY = CDbl(yValues(1))

    Dim i As Long
    For i = 2 To entities.Count
        If CDbl(xValues(i)) < minimumX Then minimumX = CDbl(xValues(i))

        If useTop Then
            If CDbl(yValues(i)) > targetY Then targetY = CDbl(yValues(i))
        Else
            If CDbl(yValues(i)) < targetY Then targetY = CDbl(yValues(i))
        End If
    Next i

    For i = 1 To entities.Count
        If Abs(CDbl(xValues(i)) - minimumX) <= _
           Module8_RuntimeSupport.PROJECTED_TOLERANCE_M And _
           Abs(CDbl(yValues(i)) - targetY) <= _
           Module8_RuntimeSupport.PROJECTED_TOLERANCE_M Then

            Set proof.DrawingEntity = entities(i)
            Set proof.ModelEntity = modelEntities(i)
            Set proof.Component = components(i)
            proof.ComponentName = proof.Component.Name2
            proof.EntityKind = "Vertex"
            proof.ModelX = CDbl(modelXValues(i))
            proof.ModelY = CDbl(modelYValues(i))
            proof.ModelZ = CDbl(modelZValues(i))
            proof.ViewX = CDbl(xValues(i))
            proof.ViewY = CDbl(yValues(i))
            Module8_RuntimeSupport.ViewToSheetCoordinates _
                swView, proof.ViewX, proof.ViewY, proof.SheetX, proof.SheetY
            proof.Proven = True
            proof.ProofSource = IIf(useTop, _
                "visible model vertex at simultaneous minimum-X/maximum-Y", _
                "visible model vertex at simultaneous minimum-X/minimum-Y")
            proof.StableKey = BuildDatumStableKey(proof)
            Exit Sub
        End If
    Next i

    proof.FailureReason = _
        "No visible vertex lies at both requested projected extrema."
    Exit Sub

Failed:
    proof.FailureReason = "Corner proof error: " & Err.Description
End Sub

Private Sub ProveCenterDatum( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef candidates As Collection, _
    ByRef proof As CDatumProof)

    On Error GoTo Failed

    Dim entities As New Collection
    Dim modelEntities As New Collection
    Dim components As New Collection
    Dim modelXValues As New Collection
    Dim modelYValues As New Collection
    Dim modelZValues As New Collection
    Dim xValues As New Collection
    Dim yValues As New Collection

    CollectModelFirstVertexCoordinates _
        swApp, swView, entities, modelEntities, components, _
        modelXValues, modelYValues, modelZValues, xValues, yValues

    Dim i As Long
    For i = 1 To entities.Count
        If Abs(CDbl(xValues(i))) <= _
           Module8_RuntimeSupport.PROJECTED_TOLERANCE_M And _
           Abs(CDbl(yValues(i))) <= _
           Module8_RuntimeSupport.PROJECTED_TOLERANCE_M Then

            Set proof.DrawingEntity = entities(i)
            Set proof.ModelEntity = modelEntities(i)
            Set proof.Component = components(i)
            proof.ComponentName = proof.Component.Name2
            proof.EntityKind = "Vertex"
            proof.ModelX = CDbl(modelXValues(i))
            proof.ModelY = CDbl(modelYValues(i))
            proof.ModelZ = CDbl(modelZValues(i))
            proof.ViewX = CDbl(xValues(i))
            proof.ViewY = CDbl(yValues(i))
            Module8_RuntimeSupport.ViewToSheetCoordinates _
                swView, proof.ViewX, proof.ViewY, proof.SheetX, proof.SheetY
            proof.Proven = True
            proof.ProofSource = "visible model vertex at projected model origin"
            proof.StableKey = BuildDatumStableKey(proof)
            Exit Sub
        End If
    Next i

    For i = 1 To candidates.Count
        Dim candidate As CHoleCandidate
        Set candidate = candidates(i)

        If Abs(candidate.ViewX) <= _
           Module8_RuntimeSupport.PROJECTED_TOLERANCE_M And _
           Abs(candidate.ViewY) <= _
           Module8_RuntimeSupport.PROJECTED_TOLERANCE_M Then

            Set proof.DrawingEntity = AsEntity(candidate.DrawingEdge)
            Set proof.ModelEntity = AsEntity(candidate.ModelEdge)
            Set proof.Component = candidate.Component
            proof.ComponentName = candidate.ComponentName
            proof.EntityKind = "CircularEdge"
            proof.ModelX = candidate.ModelX
            proof.ModelY = candidate.ModelY
            proof.ModelZ = candidate.ModelZ
            proof.ViewX = candidate.ViewX
            proof.ViewY = candidate.ViewY
            proof.SheetX = candidate.SheetX
            proof.SheetY = candidate.SheetY
            proof.Proven = True
            proof.ProofSource = _
                "ownership-proven circular edge centered at projected model origin"
            proof.StableKey = BuildDatumStableKey(proof)

            If proof.DrawingEntity Is Nothing Or proof.ModelEntity Is Nothing Then
                proof.Proven = False
                proof.StableKey = vbNullString
                proof.FailureReason = _
                    "Circular datum edge/model correspondence does not expose IEntity."
            End If
            Exit Sub
        End If
    Next i

    proof.FailureReason = _
        "No selectable visible entity is exactly at the projected model origin."
    Exit Sub

Failed:
    proof.FailureReason = "Center proof error: " & Err.Description
End Sub

Private Sub CollectModelFirstVertexCoordinates( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef entities As Collection, _
    ByRef modelEntities As Collection, _
    ByRef components As Collection, _
    ByRef modelXValues As Collection, _
    ByRef modelYValues As Collection, _
    ByRef modelZValues As Collection, _
    ByRef xValues As Collection, _
    ByRef yValues As Collection)

    On Error GoTo Failed

    Dim componentFailure As String
    Dim component As SldWorks.Component2
    Set component = GetSingleVisibleComponent(swView, componentFailure)
    If component Is Nothing Then Exit Sub

    Dim referencedModel As SldWorks.ModelDoc2
    Set referencedModel = swView.ReferencedDocument
    If referencedModel Is Nothing Then Exit Sub
    If referencedModel.GetType <> swDocPART Then Exit Sub

    Dim swPart As SldWorks.PartDoc
    Set swPart = referencedModel
    If swPart Is Nothing Then Exit Sub

    Dim bodies As Variant
    bodies = swPart.GetBodies2(swSolidBody, True)
    If IsEmpty(bodies) Or Not IsArray(bodies) Then Exit Sub

    Dim bodyIndex As Long
    For bodyIndex = LBound(bodies) To UBound(bodies)
        Dim body As SldWorks.Body2
        Set body = bodies(bodyIndex)

        If Not body Is Nothing Then
            Dim bodyEdges As Variant
            bodyEdges = body.GetEdges

            If IsArray(bodyEdges) Then
                Dim edgeIndex As Long
                For edgeIndex = LBound(bodyEdges) To UBound(bodyEdges)
                    Dim modelEdge As SldWorks.Edge
                    Set modelEdge = bodyEdges(edgeIndex)

                    If Not modelEdge Is Nothing Then
                        Dim startVertex As SldWorks.Vertex
                        Dim endVertex As SldWorks.Vertex
                        Set startVertex = modelEdge.GetStartVertex
                        Set endVertex = modelEdge.GetEndVertex

                        AddMappedModelVertex swApp, swView, component, _
                            startVertex, entities, modelEntities, components, _
                            modelXValues, modelYValues, modelZValues, _
                            xValues, yValues
                        AddMappedModelVertex swApp, swView, component, _
                            endVertex, entities, modelEntities, components, _
                            modelXValues, modelYValues, modelZValues, _
                            xValues, yValues
                    End If
                Next edgeIndex
            End If
        End If
    Next bodyIndex
    Exit Sub

Failed:
    ' The caller fails closed when no selectable mapped vertex is available.
End Sub

Private Sub AddMappedModelVertex( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByRef modelVertex As SldWorks.Vertex, _
    ByRef entities As Collection, _
    ByRef modelEntities As Collection, _
    ByRef components As Collection, _
    ByRef modelXValues As Collection, _
    ByRef modelYValues As Collection, _
    ByRef modelZValues As Collection, _
    ByRef xValues As Collection, _
    ByRef yValues As Collection)

    If modelVertex Is Nothing Then Exit Sub
    On Error GoTo Failed

    Dim mappedObject As Object
    Set mappedObject = swView.GetCorrespondingEntity(modelVertex)
    If mappedObject Is Nothing Then Exit Sub

    Dim drawingVertex As SldWorks.Vertex
    Set drawingVertex = mappedObject
    If drawingVertex Is Nothing Then Exit Sub

    Dim pointData As Variant
    pointData = modelVertex.GetPoint
    If Not IsArray(pointData) Then Exit Sub

    Dim viewX As Double
    Dim viewY As Double
    Dim viewZ As Double
    If Not Module8_RuntimeSupport.TransformPointToView( _
        swApp, swView, CDbl(pointData(0)), CDbl(pointData(1)), _
        CDbl(pointData(2)), viewX, viewY, viewZ) Then Exit Sub

    If MappedVertexAlreadyCollected( _
        modelXValues, modelYValues, modelZValues, _
        CDbl(pointData(0)), CDbl(pointData(1)), CDbl(pointData(2))) Then _
        Exit Sub

    Dim drawingEntity As SldWorks.Entity
    Dim modelEntity As SldWorks.Entity
    Set drawingEntity = AsEntity(drawingVertex)
    Set modelEntity = AsEntity(modelVertex)

    If drawingEntity Is Nothing Or modelEntity Is Nothing Then Exit Sub

    entities.Add drawingEntity
    modelEntities.Add modelEntity
    components.Add component
    modelXValues.Add CDbl(pointData(0))
    modelYValues.Add CDbl(pointData(1))
    modelZValues.Add CDbl(pointData(2))
    xValues.Add viewX
    yValues.Add viewY
    Exit Sub

Failed:
    ' One unmappable vertex does not invalidate other model-first vertices.
End Sub

Private Function MappedVertexAlreadyCollected( _
    ByRef modelXValues As Collection, _
    ByRef modelYValues As Collection, _
    ByRef modelZValues As Collection, _
    ByVal modelX As Double, _
    ByVal modelY As Double, _
    ByVal modelZ As Double) As Boolean

    Dim i As Long
    For i = 1 To modelXValues.Count
        If Abs(CDbl(modelXValues(i)) - modelX) <= _
           Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M And _
           Abs(CDbl(modelYValues(i)) - modelY) <= _
           Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M And _
           Abs(CDbl(modelZValues(i)) - modelZ) <= _
           Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then

            MappedVertexAlreadyCollected = True
            Exit Function
        End If
    Next i
End Function

Private Function CreateDirectionOrdinateGroups( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef datum As CDatumProof, _
    ByRef candidates As Collection, _
    ByVal useX As Boolean, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To candidates.Count
        Dim candidate As CHoleCandidate
        Set candidate = candidates(i)

        If Not candidate.CoverageInspectionSucceeded Then
            evidence.AddFailure "Coverage was not proved for candidate '" & _
                candidate.PhysicalInstanceKey & "'."
            Exit Function
        End If

        Dim alreadyCovered As Boolean
        If useX Then
            alreadyCovered = candidate.CoveredX
        Else
            alreadyCovered = candidate.CoveredY
        End If

        If Not alreadyCovered Then
            If Module8_RuntimeSupport.ObjectsAreSame( _
                swApp, candidate.DrawingEdge, datum.DrawingEntity) Then

                MarkCandidateDirectionCoverage candidate, useX, _
                    "ProvenDatumZero:" & datum.StableKey
                evidence.AddInfo "EVIDENCE|DATUM_ZERO_COVERAGE|group=" & _
                    EvidenceValue(BuildOrdinateGroupKey( _
                        candidate, swView, datum, useX)) & _
                    "|instance=" & _
                    EvidenceValue(candidate.PhysicalInstanceKey) & _
                    "|direction=" & IIf(useX, "X", "Y")
                alreadyCovered = True
            End If
        End If

        If Not alreadyCovered Then
            If Len(candidate.FamilyKey) = 0 Or _
               Len(candidate.PhysicalInstanceKey) = 0 Then

                evidence.AddFailure _
                    "Ordinate grouping received a candidate without a stable family/instance key."
                Exit Function
            End If

            Dim groupKey As String
            groupKey = BuildOrdinateGroupKey(candidate, swView, datum, useX)

            Dim groupCandidates As Collection
            If groups.Exists(groupKey) Then
                Set groupCandidates = groups(groupKey)
            Else
                Set groupCandidates = New Collection
                groups.Add groupKey, groupCandidates
            End If

            AddUniqueCoordinateToGroup _
                groupCandidates, candidate, useX, groupKey, evidence
        End If
    Next i

    If groups.Count = 0 Then
        evidence.AddInfo IIf(useX, "X", "Y") & _
            " coverage already satisfied in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    Dim key As Variant
    For Each key In groups.Keys
        Set groupCandidates = groups(key)

        If CreateOneOrdinateGroup( _
            swApp, swDrawModel, swDraw, swView, datum, groupCandidates, _
            useX, CStr(key), evidence) Then

            MarkCreatedCoordinateCoverage _
                candidates, groupCandidates, swView, datum, useX, _
                CStr(key), evidence
            CreateDirectionOrdinateGroups = _
                CreateDirectionOrdinateGroups + 1
        End If
    Next key
    Exit Function

Failed:
    evidence.AddFailure "Direction-group planning error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
End Function

Private Function BuildOrdinateGroupKey( _
    ByRef candidate As CHoleCandidate, _
    ByRef swView As SldWorks.View, _
    ByRef datum As CDatumProof, _
    ByVal useX As Boolean) As String

    BuildOrdinateGroupKey = candidate.FamilyKey & _
        "|view=" & LCase$(Module8_RuntimeSupport.GetViewName(swView)) & _
        "|datum=" & datum.StableKey & _
        "|direction=" & IIf(useX, "X", "Y")
End Function

Private Sub AddUniqueCoordinateToGroup( _
    ByRef groupCandidates As Collection, _
    ByRef candidate As CHoleCandidate, _
    ByVal useX As Boolean, _
    ByVal groupKey As String, _
    ByRef evidence As CRunEvidence)

    Dim i As Long
    For i = 1 To groupCandidates.Count
        Dim existing As CHoleCandidate
        Set existing = groupCandidates(i)

        Dim coordinateMatches As Boolean
        If useX Then
            coordinateMatches = _
                Abs(existing.ViewX - candidate.ViewX) <= _
                Module8_RuntimeSupport.PROJECTED_TOLERANCE_M
        Else
            coordinateMatches = _
                Abs(existing.ViewY - candidate.ViewY) <= _
                Module8_RuntimeSupport.PROJECTED_TOLERANCE_M
        End If

        If coordinateMatches Then
            evidence.AddInfo "EVIDENCE|COORDINATE_SUPPRESSED|group=" & _
                EvidenceValue(groupKey) & _
                "|keptInstance=" & _
                EvidenceValue(existing.PhysicalInstanceKey) & _
                "|suppressedInstance=" & _
                EvidenceValue(candidate.PhysicalInstanceKey) & _
                "|coordinate=" & Format$( _
                    IIf(useX, candidate.ViewX, candidate.ViewY), "0.000000000") & _
                "|tolerance=" & _
                Format$(Module8_RuntimeSupport.PROJECTED_TOLERANCE_M, _
                        "0.000000000")
            Exit Sub
        End If
    Next i

    groupCandidates.Add candidate
End Sub

Private Sub MarkCreatedCoordinateCoverage( _
    ByRef allCandidates As Collection, _
    ByRef selectedCoordinates As Collection, _
    ByRef swView As SldWorks.View, _
    ByRef datum As CDatumProof, _
    ByVal useX As Boolean, _
    ByVal groupKey As String, _
    ByRef evidence As CRunEvidence)

    Dim candidateIndex As Long
    For candidateIndex = 1 To allCandidates.Count
        Dim candidate As CHoleCandidate
        Set candidate = allCandidates(candidateIndex)

        Dim alreadyCovered As Boolean
        If useX Then
            alreadyCovered = candidate.CoveredX
        Else
            alreadyCovered = candidate.CoveredY
        End If

        If Not alreadyCovered Then
            If StrComp(BuildOrdinateGroupKey( _
                candidate, swView, datum, useX), groupKey, _
                vbTextCompare) = 0 Then

                Dim selectedIndex As Long
                For selectedIndex = 1 To selectedCoordinates.Count
                    Dim representative As CHoleCandidate
                    Set representative = selectedCoordinates(selectedIndex)

                    Dim coordinateMatches As Boolean
                    If useX Then
                        coordinateMatches = _
                            Abs(representative.ViewX - candidate.ViewX) <= _
                            Module8_RuntimeSupport.PROJECTED_TOLERANCE_M
                    Else
                        coordinateMatches = _
                            Abs(representative.ViewY - candidate.ViewY) <= _
                            Module8_RuntimeSupport.PROJECTED_TOLERANCE_M
                    End If

                    If coordinateMatches Then
                        MarkCandidateDirectionCoverage candidate, useX, _
                            "CreatedOrdinateGroup:" & groupKey
                        evidence.AddInfo _
                            "EVIDENCE|CREATED_COORDINATE_COVERAGE|group=" & _
                            EvidenceValue(groupKey) & _
                            "|instance=" & _
                            EvidenceValue(candidate.PhysicalInstanceKey) & _
                            "|representative=" & _
                            EvidenceValue(representative.PhysicalInstanceKey) & _
                            "|direction=" & IIf(useX, "X", "Y")
                        Exit For
                    End If
                Next selectedIndex
            End If
        End If
    Next candidateIndex
End Sub

Private Sub MarkCandidateDirectionCoverage( _
    ByRef candidate As CHoleCandidate, _
    ByVal useX As Boolean, _
    ByVal coverageSource As String)

    If useX Then
        candidate.CoveredX = True
        candidate.CoverageSourceX = coverageSource
    Else
        candidate.CoveredY = True
        candidate.CoverageSourceY = coverageSource
    End If
End Sub

Private Function CreateOneOrdinateGroup( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef datum As CDatumProof, _
    ByRef candidates As Collection, _
    ByVal horizontalDirection As Boolean, _
    ByVal groupKey As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim outcome As String
    Dim datumSelected As Boolean
    Dim selectedByMulti As Long
    Dim finalSelectionCount As Long
    Dim resultCode As Long
    Dim cleanupSelectionCount As Long
    outcome = "NotStarted"
    selectedByMulti = -1
    finalSelectionCount = -1
    resultCode = -9999
    cleanupSelectionCount = -1

    If candidates.Count = 0 Then Exit Function

    Dim selectableCandidates As New Collection
    Dim candidateIndex As Long

    For candidateIndex = 1 To candidates.Count
        Dim sourceCandidate As CHoleCandidate
        Set sourceCandidate = candidates(candidateIndex)

        If Not Module8_RuntimeSupport.ObjectsAreSame( _
            swApp, sourceCandidate.DrawingEdge, datum.DrawingEntity) Then
            selectableCandidates.Add sourceCandidate
        End If
    Next candidateIndex

    If selectableCandidates.Count = 0 Then
        evidence.AddInfo "Ordinate direction contains only the proven zero entity in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        Exit Function
    End If

    If Not Module8_RuntimeSupport.ActivateDrawingView( _
        swDrawModel, swDraw, swView, evidence, _
        "Ordinate creation") Then

        outcome = "ViewActivationFailed"
        GoTo SafeExit
    End If

    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True

    Dim selectData As SldWorks.SelectData
    Set selectData = swDrawModel.SelectionManager.CreateSelectData
    Set selectData.View = swView

    datum.SelectionAttempts = datum.SelectionAttempts + 1
    datumSelected = datum.DrawingEntity.Select4(False, selectData)
    datum.SelectionSucceeded = datum.SelectionSucceeded Or datumSelected

    If Not datumSelected Then
        evidence.AddFailure "Datum Select4 returned False in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        outcome = "DatumSelectionFailed"
        GoTo SafeExit
    End If

    evidence.AddInfo "EVIDENCE|ORDINATE_SELECTION_ITEM|group=" & _
        EvidenceValue(groupKey) & _
        "|order=0|role=Datum|datumKey=" & EvidenceValue(datum.StableKey)

    Dim selectableObjects() As Object
    ReDim selectableObjects(0 To selectableCandidates.Count - 1)

    Dim i As Long
    For i = 1 To selectableCandidates.Count
        Dim candidate As CHoleCandidate
        Set candidate = selectableCandidates(i)
        Set selectableObjects(i - 1) = candidate.DrawingEdge

        evidence.AddInfo "EVIDENCE|ORDINATE_SELECTION_ITEM|group=" & _
            EvidenceValue(groupKey) & _
            "|order=" & CStr(i) & _
            "|role=Feature|instance=" & _
            EvidenceValue(candidate.PhysicalInstanceKey)
    Next i

    selectedByMulti = swDrawModel.Extension.MultiSelect2( _
        selectableObjects, True, selectData)

    If selectedByMulti <> selectableCandidates.Count Then
        evidence.AddFailure "MultiSelect2 selected " & CStr(selectedByMulti) & _
            " of " & CStr(selectableCandidates.Count) & " feature entities in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        outcome = "MultiSelectMismatch"
        GoTo SafeExit
    End If

    finalSelectionCount = _
        swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)

    If finalSelectionCount <> selectableCandidates.Count + 1 Then
        evidence.AddFailure "Final ordinate selection count=" & _
            CStr(finalSelectionCount) & ", expected=" & _
            CStr(selectableCandidates.Count + 1) & "."
        outcome = "FinalSelectionMismatch"
        GoTo SafeExit
    End If

    Dim outline As Variant
    outline = swView.GetOutline

    Dim dimensionType As Long
    Dim locationX As Double
    Dim locationY As Double

    If horizontalDirection Then
        dimensionType = swHorizontalOrdinate
        locationX = (CDbl(outline(0)) + CDbl(outline(2))) / 2#
        locationY = CDbl(outline(1)) - 0.012
    Else
        dimensionType = swVerticalOrdinate
        locationX = CDbl(outline(0)) - 0.012
        locationY = (CDbl(outline(1)) + CDbl(outline(3))) / 2#
    End If

    resultCode = swDrawModel.Extension.AddOrdinateDimension( _
        dimensionType, locationX, locationY, 0#)

    If resultCode <> swCreateOrdDimErr_Success Then
        evidence.AddFailure "AddOrdinateDimension returned " & _
            DecodeOrdinateResult(resultCode) & " in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
        outcome = "AddOrdinateFailed"
        GoTo SafeExit
    End If

    evidence.OrdinateEntitiesSelected = _
        evidence.OrdinateEntitiesSelected + selectableCandidates.Count
    evidence.AddInfo IIf(horizontalDirection, "Horizontal", "Vertical") & _
        " ordinate group created in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & _
        "' with " & CStr(selectableCandidates.Count) & " unique coordinates."

    outcome = "Created"
    CreateOneOrdinateGroup = True

SafeExit:
    On Error Resume Next
    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True
    cleanupSelectionCount = _
        swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)
    On Error GoTo 0

    If cleanupSelectionCount <> 0 Then
        evidence.AddFailure "Ordinate cleanup left " & _
            CStr(cleanupSelectionCount) & " selected objects in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "'."
    End If

    evidence.AddInfo "EVIDENCE|ORDINATE_GROUP|view=" & _
        EvidenceValue(Module8_RuntimeSupport.GetViewName(swView)) & _
        "|group=" & EvidenceValue(groupKey) & _
        "|direction=" & IIf(horizontalDirection, "X", "Y") & _
        "|datumKey=" & EvidenceValue(datum.StableKey) & _
        "|datumSelect=" & CStr(datumSelected) & _
        "|multiSelected=" & CStr(selectedByMulti) & _
        "|expectedCandidates=" & CStr(selectableCandidates.Count) & _
        "|finalSelection=" & CStr(finalSelectionCount) & _
        "|expectedFinal=" & CStr(selectableCandidates.Count + 1) & _
        "|resultCode=" & CStr(resultCode) & _
        "|result=" & EvidenceValue(DecodeOrdinateResult(resultCode)) & _
        "|cleanupSelection=" & CStr(cleanupSelectionCount) & _
        "|outcome=" & EvidenceValue(outcome)
    Exit Function

Failed:
    evidence.AddFailure "Ordinate group error in '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
    outcome = "RuntimeError:" & CStr(Err.Number)
    Resume SafeExit
End Function

Private Function DecodeOrdinateResult(ByVal resultCode As Long) As String
    Select Case resultCode
        Case 0: DecodeOrdinateResult = "0 Success"
        Case 1: DecodeOrdinateResult = "1 GeneralFailure"
        Case 2: DecodeOrdinateResult = "2 NoInternalDimensions"
        Case 3: DecodeOrdinateResult = "3 BadSelection"
        Case 4: DecodeOrdinateResult = "4 ModelNotLoaded"
        Case 5: DecodeOrdinateResult = "5 SamePartOnly"
        Case 6: DecodeOrdinateResult = "6 ExtraSelection"
        Case 7: DecodeOrdinateResult = "7 OrdinateFailure"
        Case 8: DecodeOrdinateResult = "8 Duplicate"
        Case 9: DecodeOrdinateResult = "9 BadDirection"
        Case -1: DecodeOrdinateResult = "-1 Undefined"
        Case -9999: DecodeOrdinateResult = "NotCalled"
        Case Else: DecodeOrdinateResult = CStr(resultCode) & " Unknown"
    End Select
End Function

Private Sub RecordRejection( _
    ByRef rejected As Object, _
    ByVal reason As String)

    If Len(reason) = 0 Then reason = "Unspecified"

    If rejected.Exists(reason) Then
        rejected(reason) = CLng(rejected(reason)) + 1
    Else
        rejected.Add reason, 1
    End If
End Sub

Private Function DictionaryTotal(ByRef dictionary As Object) As Long
    Dim key As Variant
    For Each key In dictionary.Keys
        DictionaryTotal = DictionaryTotal + CLng(dictionary(key))
    Next key
End Function

Private Sub EmitRejectionSummary( _
    ByRef swView As SldWorks.View, _
    ByRef rejected As Object, _
    ByRef evidence As CRunEvidence)

    Dim key As Variant
    For Each key In rejected.Keys
        evidence.AddInfo "Rejected in '" & _
            Module8_RuntimeSupport.GetViewName(swView) & "': " & _
            CStr(key) & "=" & CStr(rejected(key)) & "."
    Next key
End Sub

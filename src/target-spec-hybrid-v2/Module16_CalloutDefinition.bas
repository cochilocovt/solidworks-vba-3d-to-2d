Option Explicit

' R23 Phase 6. Native hole-callout reconciliation and controlled fallback.
'
' SAFETY BOUNDARY. Exactly one procedure changes a drawing:
' CreateNativeCalloutForFamily, which refuses unless passed an explicit
' allowMutation argument. R23_ProbeCalloutDefinition never passes it and
' contains no AddHoleCallout2 call, so it can be run against the manual
' reference drawing.
'
' R23-609. There is NO hardcoded callout text anywhere in this module: no
' part number, no "6X", no "M5x0.8", no diameter literal, and no scoring by
' feature name or by proximity to an expected radius. Every field comes from
' either the native callout's own variables or the typed feature data Phase
' 2 proved. The legacy path in Module7_TitleBlockEngine.bas still contains
' those literals; see the note on R23-609 in the plan for why removing them
' belongs to the phase that switches the pipeline over rather than to this
' one.

' A native hole callout reports Type2 = 6 (swDiameterDimension), but so does
' an ordinary diameter dimension, so type is never the classifier here.
' IDisplayDimension.IsHoleCallout is, and no dimension-type constant is
' declared in this module precisely so none can be reached for.
'
' swEndConditions_e 0 is swEndCondBlind. The completeness rule in
' CCalloutDefinition reads the stored code rather than a feature name, so a
' through hole is not asked for a depth it does not have.

' Definition sources.
Public Const DEFINITION_NATIVE As String = "NativeCallout"
Public Const DEFINITION_FALLBACK As String = "ControlledFallback"

Private mEmitDiagnostics As Boolean

Private Sub EmitInfo( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddInfo message
    If mEmitDiagnostics Then Debug.Print message
End Sub

Private Sub EmitWarning( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddWarning message
    If mEmitDiagnostics Then Debug.Print message
End Sub

Private Sub EmitFailure( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddFailure message
    If mEmitDiagnostics Then Debug.Print message
End Sub

' R23-600. The hole-callout bit must be requested, not hoped for. Phase 4
' verified swInsertholeCallout = 1048576 member by member against the 2025
' table; this asserts the mask R23 actually passes still carries it.
Public Function CalloutImportRequested() As Boolean
    CalloutImportRequested = _
        ((Module14_AnnotationImport.IMPORT_MASK_FULL And 1048576) <> 0)
End Function

' R23-601. Native hole-callout classification.
'
' IsHoleCallout is the ONLY test. A native callout reports Type2 = 6
' (swDiameterDimension), but so does an ordinary diameter dimension, so type
' can be recorded and never relied on. The raw return goes through
' NormalizeSwBoolean because a SOLIDWORKS COM Boolean is not reliably -1/0.
Public Function IsNativeHoleCallout( _
    ByRef swDispDim As SldWorks.DisplayDimension) As Boolean

    On Error GoTo Failed

    If swDispDim Is Nothing Then Exit Function

    IsNativeHoleCallout = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDispDim.IsHoleCallout)
    Exit Function

Failed:
    IsNativeHoleCallout = False
End Function

' R23-602. Attaches a native callout to exactly one semantic family, by the
' identity route Phase 4 proved: the callout's attached drawing entities are
' compared against the drawing entities each projection owns.
'
' Returns the family key, or an empty string when the callout cannot be
' attributed. It is never attributed by nearest hole or by radius.
Public Function MatchCalloutToFamily( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef graph As CLocationGraph, _
    ByRef swDispDim As SldWorks.DisplayDimension, _
    ByVal viewName As String, _
    ByRef diagnostics As String) As String

    On Error GoTo Failed

    diagnostics = "match=NotAttempted"

    If swDispDim Is Nothing Then Exit Function
    If graph Is Nothing Then Exit Function

    Dim swAnnotation As SldWorks.Annotation
    Set swAnnotation = swDispDim.GetAnnotation
    If swAnnotation Is Nothing Then
        diagnostics = "match=NoAnnotation"
        Exit Function
    End If

    Dim attachments As Variant
    attachments = swAnnotation.GetAttachedEntities3

    If VariantItemCount(attachments) = 0 Then
        diagnostics = "match=NoAttachedEntities"
        Exit Function
    End If

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(viewName)

    Dim matchedKeys As String
    Dim familyKey As String

    Dim p As Long
    For p = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(p)

        If projection Is Nothing Then GoTo ContinueProjection
        If projection.PhysicalLocation Is Nothing Then GoTo ContinueProjection

        Dim a As Long
        For a = LBound(attachments) To UBound(attachments)
            Dim attached As Object
            Set attached = Nothing
            On Error Resume Next
            Set attached = attachments(a)
            On Error GoTo Failed

            If attached Is Nothing Then GoTo ContinueAttachment

            If ProjectionOwnsDrawingEntity( _
                swApp, projection, attached) Then

                Dim candidateKey As String
                candidateKey = _
                    projection.PhysicalLocation.SemanticFamilyKey

                If Len(familyKey) = 0 Then
                    familyKey = candidateKey
                ElseIf StrComp(familyKey, candidateKey, _
                    vbBinaryCompare) <> 0 Then

                    ' R23-602 says ONE family. A callout that resolves to
                    ' two is a real ambiguity, not something to break ties
                    ' on.
                    diagnostics = "match=AmbiguousFamilies" & _
                        "|first=" & familyKey & _
                        "|second=" & candidateKey
                    MatchCalloutToFamily = vbNullString
                    Exit Function
                End If

                matchedKeys = matchedKeys & "1"
            End If

ContinueAttachment:
        Next a

ContinueProjection:
    Next p

    If Len(familyKey) = 0 Then
        diagnostics = "match=NoOwningProjection" & _
            "|attachments=" & CStr(VariantItemCount(attachments)) & _
            "|projectionsInView=" & CStr(projections.Count)
        Exit Function
    End If

    diagnostics = "match=Family" & _
        "|attachments=" & CStr(VariantItemCount(attachments)) & _
        "|ownedMatches=" & CStr(Len(matchedKeys)) & _
        "|identity=ISldWorks.IsSame"

    MatchCalloutToFamily = familyKey
    Exit Function

Failed:
    diagnostics = "match=Error:" & CStr(Err.Number)
    MatchCalloutToFamily = vbNullString
End Function

' Identity only, against every drawing entity the projection owns rather
' than only its chosen anchor. Phase 4 established that the anchor tier
' deliberately prefers the through hole while a counterbore callout attaches
' to the wider mouth, so anchor-only comparison misses the callout.
Private Function ProjectionOwnsDrawingEntity( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef projection As CViewHoleProjection, _
    ByRef candidate As Object) As Boolean

    On Error GoTo Failed

    If projection Is Nothing Then Exit Function
    If candidate Is Nothing Then Exit Function

    If Not projection.PrimaryAnchor Is Nothing Then
        If ObjectsAreSame(swApp, projection.PrimaryAnchor, candidate) Then
            ProjectionOwnsDrawingEntity = True
            Exit Function
        End If
    End If

    Dim i As Long
    For i = 1 To projection.DrawingEntityAliases.Count
        Dim alternate As Object
        Set alternate = Nothing
        On Error Resume Next
        Set alternate = projection.DrawingEntityAliases(i)
        On Error GoTo Failed

        If Not alternate Is Nothing Then
            If ObjectsAreSame(swApp, alternate, candidate) Then
                ProjectionOwnsDrawingEntity = True
                Exit Function
            End If
        End If
    Next i

    Exit Function

Failed:
    ProjectionOwnsDrawingEntity = False
End Function

' ISldWorks.IsSame returns swObjectEquality, NOT a Boolean: 0 NotSame,
' 1 Same, 2 Unsupported. Only 1 is a match.
Private Function ObjectsAreSame( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef first As Object, _
    ByRef second As Object) As Boolean

    On Error GoTo Failed

    If first Is Nothing Or second Is Nothing Then Exit Function

    ObjectsAreSame = (CLng(swApp.IsSame(first, second)) = 1)
    Exit Function

Failed:
    ObjectsAreSame = False
End Function

' R23-603. Reads a native callout's own variables.
'
' IDisplayDimension.GetHoleCalloutVariables returns an array of
' ICalloutVariable, each exposing HoleFit, ShaftFit, ToleranceType,
' ToleranceMin and ToleranceMax. Reading the callout's rendered text instead
' would give a string that cannot be validated field by field, which is what
' R23-603 exists to prevent.
Public Sub ReadNativeCalloutFields( _
    ByRef swDispDim As SldWorks.DisplayDimension, _
    ByRef definition As CCalloutDefinition, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim variables As Variant
    variables = swDispDim.GetHoleCalloutVariables

    Dim variableCount As Long
    variableCount = VariantItemCount(variables)
    definition.NativeVariableCount = variableCount

    If variableCount = 0 Then
        definition.ToleranceProofSource = "Unproven"
        Exit Sub
    End If

    Dim i As Long
    For i = LBound(variables) To UBound(variables)
        Dim calloutVariable As Object
        Set calloutVariable = Nothing
        On Error Resume Next
        Set calloutVariable = variables(i)
        On Error GoTo Failed

        If calloutVariable Is Nothing Then GoTo ContinueVariable

        Dim holeFit As String
        Dim shaftFit As String
        Dim toleranceType As Long
        Dim toleranceMin As Double
        Dim toleranceMax As Double
        Dim variableName As String

        holeFit = vbNullString
        shaftFit = vbNullString
        toleranceType = -1
        toleranceMin = 0#
        toleranceMax = 0#
        variableName = vbNullString

        On Error Resume Next
        variableName = calloutVariable.UserReadableVariableName
        holeFit = calloutVariable.HoleFit
        shaftFit = calloutVariable.ShaftFit
        toleranceType = CLng(calloutVariable.ToleranceType)
        toleranceMin = CDbl(calloutVariable.ToleranceMin)
        toleranceMax = CDbl(calloutVariable.ToleranceMax)
        On Error GoTo Failed

        EmitInfo evidence, "CALLOUT_VARIABLE|family=" & _
            definition.FamilyKey & _
            "|index=" & CStr(i) & _
            "|name=" & variableName & _
            "|holeFit=" & holeFit & _
            "|shaftFit=" & shaftFit & _
            "|toleranceType=" & CStr(toleranceType) & _
            "|toleranceMin=" & Format$(toleranceMin, "0.000000000") & _
            "|toleranceMax=" & Format$(toleranceMax, "0.000000000") & _
            "|source=ICalloutVariable"

        ' The first variable carrying fit or tolerance data defines the
        ' callout's fit fields. Later variables are reported but do not
        ' overwrite a value already proved.
        If StrComp(definition.ToleranceProofSource, "Unproven", _
            vbBinaryCompare) = 0 Then

            If Len(Trim$(holeFit)) > 0 Or Len(Trim$(shaftFit)) > 0 Or _
                toleranceType > 0 Then

                definition.HoleFit = holeFit
                definition.ShaftFit = shaftFit
                definition.ToleranceType = toleranceType
                definition.ToleranceMinM = toleranceMin
                definition.ToleranceMaxM = toleranceMax
                definition.ToleranceProofSource = _
                    "IDisplayDimension.GetHoleCalloutVariables" & _
                    ".ICalloutVariable"
            End If
        End If

ContinueVariable:
    Next i

    Exit Sub

Failed:
    EmitWarning evidence, "CALLOUT_VARIABLE_ERROR|family=" & _
        definition.FamilyKey & "|error=" & CStr(Err.Number)
End Sub

' R23-605, R23-606, R23-607 and R23-608. Builds the definition for one
' family from typed feature data.
'
' This runs for every family, so that a native callout can be COMPARED
' against what the model actually says rather than trusted blindly. R23-605
' governs which one is retained, not which one is computed.
Public Function BuildDefinitionFromTypedData( _
    ByRef graph As CLocationGraph, _
    ByVal familyKey As String, _
    ByRef evidence As CRunEvidence) As CCalloutDefinition

    On Error GoTo Failed

    Dim definition As CCalloutDefinition
    Set definition = New CCalloutDefinition
    Set BuildDefinitionFromTypedData = definition

    definition.FamilyKey = familyKey
    definition.DefinitionSource = DEFINITION_FALLBACK

    Dim locations As Collection
    Set locations = graph.LocationsForFamily(familyKey)

    ' R23-606. Quantity is the number of unique physical locations. Not the
    ' feature count: one Hole Wizard feature plus a mirror produces many
    ' holes. Not the edge count: a counterbore contributes several edges per
    ' hole.
    definition.Quantity = locations.Count
    definition.QuantityProofSource = _
        "CLocationGraph.LocationsForFamily.UniquePhysicalLocations"

    If locations.Count = 0 Then
        definition.RejectionReason = "NoPhysicalLocationsInFamily"
        Exit Function
    End If

    Dim firstLocation As CPhysicalHoleLocation
    Set firstLocation = locations(1)

    definition.MachiningFaceKey = "axis=" & _
        Module11_GeometryIdentity.CanonicalAxisToken( _
            firstLocation.AxisX, firstLocation.AxisY, firstLocation.AxisZ)

    ' R23-607 and R23-608. Every remaining field comes from the typed
    ' feature definitions in the location's own stack, each carrying the
    ' proof source Phase 2 recorded. Nothing is parsed from a feature name.
    Dim s As Long
    For s = 1 To firstLocation.StackFeatures.Count
        Dim featureDefinition As CFeatureDefinition
        Set featureDefinition = firstLocation.StackFeatures(s)
        If featureDefinition Is Nothing Then GoTo ContinueFeature

        ApplyTypedFeatureFields definition, featureDefinition

ContinueFeature:
    Next s

    Exit Function

Failed:
    EmitFailure evidence, "CALLOUT_FALLBACK_ERROR|family=" & familyKey & _
        "|error=" & CStr(Err.Number)
End Function

' Copies one typed feature's proved fields onto the definition. A field is
' only taken when the feature actually proved it, so an unproven field stays
' unproven rather than being overwritten with a zero.
Private Sub ApplyTypedFeatureFields( _
    ByRef definition As CCalloutDefinition, _
    ByRef featureDefinition As CFeatureDefinition)

    On Error Resume Next

    If featureDefinition.NominalDiameterM > 0# And _
        definition.NominalDiameterM <= 0# Then

        definition.NominalDiameterM = featureDefinition.NominalDiameterM
        definition.DiameterProofSource = _
            featureDefinition.DiameterProofSource
    End If

    If featureDefinition.DepthM > 0# And definition.DepthM <= 0# Then
        definition.DepthM = featureDefinition.DepthM
        definition.DepthProofSource = featureDefinition.DepthProofSource
    End If

    If StrComp(definition.EndConditionProofSource, "Unproven", _
        vbBinaryCompare) = 0 Then

        If Len(Trim$(featureDefinition.EndConditionProofSource)) > 0 And _
            StrComp(featureDefinition.EndConditionProofSource, _
                "Unproven", vbBinaryCompare) <> 0 Then

            definition.EndConditionCode = featureDefinition.EndConditionCode
            definition.EndConditionProofSource = _
                featureDefinition.EndConditionProofSource
        End If
    End If

    ' R23-607.
    If featureDefinition.CounterBoreDiameterM > 0# And _
        definition.CounterBoreDiameterM <= 0# Then

        definition.CounterBoreDiameterM = _
            featureDefinition.CounterBoreDiameterM
        definition.CounterBoreDepthM = featureDefinition.CounterBoreDepthM
        definition.CounterBoreProofSource = _
            featureDefinition.CounterBoreProofSource
    End If

    If Len(Trim$(featureDefinition.ThreadDescription)) > 0 And _
        Len(Trim$(definition.ThreadDescription)) = 0 Then

        definition.ThreadDescription = featureDefinition.ThreadDescription
        definition.ThreadDepthM = featureDefinition.ThreadDepthM
        definition.ThreadProofSource = featureDefinition.ThreadProofSource
    End If

    ' R23-608. Fit and tolerance come only from source model dimension data.
    ' Phase 4's standing instruction stands: a tolerance read off a drawing
    ' is evidence that a designer typed a number, and must never be promoted
    ' into a manufacturing requirement here.
    If StrComp(definition.ToleranceProofSource, "Unproven", _
        vbBinaryCompare) = 0 Then

        If Len(Trim$(featureDefinition.FitDescription)) > 0 Then
            definition.HoleFit = featureDefinition.FitDescription
            definition.ToleranceMinM = featureDefinition.ToleranceMinM
            definition.ToleranceMaxM = featureDefinition.ToleranceMaxM
            definition.ToleranceProofSource = _
                featureDefinition.FitProofSource
        End If
    End If
End Sub

' R23-604 and R23-605. Chooses which definition is retained for a family.
'
' A COMPLETE native callout always wins: it is associative, so it tracks the
' model. The controlled fallback is used only when no complete native
' definition exists, and the reason the native one was rejected is recorded
' so "we built our own" is never an unexplained outcome.
Public Function RetainDefinitionForFamily( _
    ByRef nativeDefinition As CCalloutDefinition, _
    ByRef fallbackDefinition As CCalloutDefinition, _
    ByRef decisionProof As String) As CCalloutDefinition

    On Error GoTo Failed

    If Not nativeDefinition Is Nothing Then
        If nativeDefinition.IsComplete() Then
            decisionProof = "retained=NativeCallout" & _
                "|reason=CompleteAssociativeDefinitionAvailable"
            Set RetainDefinitionForFamily = nativeDefinition
            Exit Function
        End If

        decisionProof = "retained=ControlledFallback" & _
            "|reason=NativeIncomplete" & _
            "|nativeMissing=" & nativeDefinition.CompletenessFailureReason()
    Else
        decisionProof = "retained=ControlledFallback" & _
            "|reason=NoNativeCalloutAttributedToFamily"
    End If

    Set RetainDefinitionForFamily = fallbackDefinition
    Exit Function

Failed:
    decisionProof = "retained=None|reason=Error:" & CStr(Err.Number)
    Set RetainDefinitionForFamily = Nothing
End Function

' R23-610. Reports a field-specific reason rather than a bare failure, so a
' missing thread depth and a missing attachment are never the same result.
Public Function VerifyManufacturingDefinitions( _
    ByRef definitions As Collection) As String

    On Error GoTo Failed

    Dim failures As String

    Dim i As Long
    For i = 1 To definitions.Count
        Dim definition As CCalloutDefinition
        Set definition = definitions(i)

        Dim reason As String
        reason = definition.CompletenessFailureReason()

        If StrComp(reason, "None", vbBinaryCompare) <> 0 Then
            failures = AppendFailure(failures, _
                definition.FamilyKey & ":" & reason)
        End If
    Next i

    If Len(failures) = 0 Then failures = "None"

    VerifyManufacturingDefinitions = "definitions=" & _
        CStr(definitions.Count) & _
        "|definitionFailures=" & failures
    Exit Function

Failed:
    VerifyManufacturingDefinitions = _
        "definitionFailures=Error:" & CStr(Err.Number)
End Function

' R23-611. Required definitions are stated as SHAPES rather than as part
' numbers: one multi-hole counterbored family and one multi-hole threaded
' family. P-0251 satisfies this with its six counterbores and four tapped
' side holes, but nothing here is keyed to that part.
Public Function VerifyRequiredDefinitionShapes( _
    ByRef definitions As Collection) As String

    On Error GoTo Failed

    Dim counterboreFamilies As Long
    Dim threadedFamilies As Long
    Dim failures As String

    Dim i As Long
    For i = 1 To definitions.Count
        Dim definition As CCalloutDefinition
        Set definition = definitions(i)

        If definition.Quantity < 2 Then GoTo ContinueDefinition
        If Not definition.IsComplete() Then GoTo ContinueDefinition

        If definition.CounterBoreDiameterM > 0# Then
            counterboreFamilies = counterboreFamilies + 1
        End If

        If Len(Trim$(definition.ThreadDescription)) > 0 Then
            threadedFamilies = threadedFamilies + 1
        End If

ContinueDefinition:
    Next i

    If counterboreFamilies = 0 Then
        failures = AppendFailure(failures, _
            "NoCompleteCounterboredFamilyDefinition")
    End If

    If threadedFamilies = 0 Then
        failures = AppendFailure(failures, _
            "NoCompleteThreadedFamilyDefinition")
    End If

    If Len(failures) = 0 Then failures = "None"

    VerifyRequiredDefinitionShapes = _
        "counterboredFamilies=" & CStr(counterboreFamilies) & _
        "|threadedFamilies=" & CStr(threadedFamilies) & _
        "|shapeFailures=" & failures
    Exit Function

Failed:
    VerifyRequiredDefinitionShapes = "shapeFailures=Error:" & CStr(Err.Number)
End Function

' R23-604. MUTATES THE DRAWING.
'
' Refuses unless allowMutation is True. Selects the family's proven drawing
' entity, then calls IDrawingDoc.AddHoleCallout2, whose Remarks require the
' hole's edge to be selected first. The created callout is associative,
' which is the whole reason a native callout is preferred over text.
'
' Returns the created DisplayDimension, or Nothing on refusal or failure.
Public Function CreateNativeCalloutForFamily( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef projection As CViewHoleProjection, _
    ByVal placeX As Double, _
    ByVal placeY As Double, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As SldWorks.DisplayDimension

    On Error GoTo Failed

    If Not allowMutation Then
        EmitWarning evidence, "CALLOUT_CREATE_REFUSED" & _
            "|reason=MutationNotAuthorized"
        Exit Function
    End If

    If swDraw Is Nothing Then Exit Function
    If projection Is Nothing Then Exit Function

    If Not projection.HasSelectableAnchor() Then
        EmitFailure evidence, "CALLOUT_CREATE_REFUSED" & _
            "|reason=NoProvenAnchor" & _
            "|physical=" & projection.PhysicalInstanceKey
        Exit Function
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    If Not Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDrawing.ActivateView(projection.ViewName)) Then

        EmitFailure evidence, "CALLOUT_CREATE_REFUSED" & _
            "|reason=ViewActivationFailed|view=" & projection.ViewName
        Exit Function
    End If

    swDraw.ClearSelection2 True

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDraw.SelectionManager

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionMgr.CreateSelectData

    Dim viewBinding As String
    viewBinding = TryBindSelectDataView(selectData, projection.DrawingView)

    Dim selected As Boolean
    selected = Module11_GeometryIdentity.NormalizeSwBoolean( _
        projection.PrimaryAnchor.Select4(False, selectData))

    Dim selectedCount As Long
    selectedCount = selectionMgr.GetSelectedObjectCount2(-1)

    If Not selected Or selectedCount <> 1 Then
        EmitFailure evidence, "CALLOUT_CREATE_REFUSED" & _
            "|reason=AnchorSelectionFailed" & _
            "|selected=" & CStr(selected) & _
            "|selectedCount=" & CStr(selectedCount)
        swDraw.ClearSelection2 True
        Exit Function
    End If

    Dim created As Object
    Set created = swDrawing.AddHoleCallout2(placeX, placeY, 0#)

    swDraw.ClearSelection2 True

    If created Is Nothing Then
        EmitFailure evidence, "CALLOUT_CREATE_FAILED" & _
            "|reason=AddHoleCallout2ReturnedNothing" & _
            "|view=" & projection.ViewName
        Exit Function
    End If

    Set CreateNativeCalloutForFamily = created

    EmitInfo evidence, "CALLOUT_CREATED|view=" & projection.ViewName & _
        "|physical=" & projection.PhysicalInstanceKey & _
        "|viewBinding=" & viewBinding & _
        "|isHoleCallout=" & CStr(IsNativeHoleCallout(created)) & _
        "|source=IDrawingDoc.AddHoleCallout2"
    Exit Function

Failed:
    EmitFailure evidence, "CALLOUT_CREATE_ERROR|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    swDraw.ClearSelection2 True
    Set CreateNativeCalloutForFamily = Nothing
End Function

Private Function TryBindSelectDataView( _
    ByRef selectData As SldWorks.SelectData, _
    ByRef swView As SldWorks.View) As String

    On Error GoTo Failed
    Set selectData.View = swView
    TryBindSelectDataView = "Bound"
    Exit Function
Failed:
    TryBindSelectDataView = "UnboundAfterError:" & CStr(Err.Number)
End Function

Private Function AppendFailure( _
    ByVal existing As String, _
    ByVal reason As String) As String

    If Len(existing) = 0 Then
        AppendFailure = reason
    Else
        AppendFailure = existing & ";" & reason
    End If
End Function

Private Function VariantItemCount(ByVal value As Variant) As Long
    On Error GoTo Failed

    If IsEmpty(value) Then Exit Function
    If Not IsArray(value) Then Exit Function

    VariantItemCount = UBound(value) - LBound(value) + 1
    Exit Function

Failed:
    VariantItemCount = 0
End Function

Private Function SafeViewName( _
    ByRef swView As SldWorks.View) As String

    On Error Resume Next
    If swView Is Nothing Then
        SafeViewName = "Nothing"
        Exit Function
    End If

    SafeViewName = "Unnamed"
    SafeViewName = swView.GetName2
End Function

' R23-600 to R23-611 evidence entry point. STRICTLY READ-ONLY: it contains
' no call to CreateNativeCalloutForFamily and no AddHoleCallout2 call, so it
' cannot create an annotation.
Public Sub R23_ProbeCalloutDefinition()
    On Error GoTo Failed

    mEmitDiagnostics = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Debug.Print "R23_CALLOUT_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Debug.Print "R23_CALLOUT_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Debug.Print "R23_CALLOUT_FATAL|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Debug.Print "R23_CALLOUT_FATAL|reason=NoViewsOnSheet"
        Exit Sub
    End If

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = Nothing

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim candidateView As SldWorks.View
        Set candidateView = views(i)
        If Not candidateView Is Nothing Then
            If Not candidateView.ReferencedDocument Is Nothing Then
                Set swPart = candidateView.ReferencedDocument
                Exit For
            End If
        End If
    Next i

    If swPart Is Nothing Then
        Debug.Print "R23_CALLOUT_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Debug.Print "R23_CALLOUT_FATAL|reason=UnauthorizedFixture" & _
            "|path=" & partPath
        Exit Sub
    End If

    Dim drawingSaveBefore As Boolean
    drawingSaveBefore = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Dim initialSelectionCount As Long
    initialSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim evidence As CRunEvidence
    Set evidence = New CRunEvidence

    Dim graph As CLocationGraph
    Set graph = New CLocationGraph

    Debug.Print "R23_CALLOUT_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0" & _
        "|calloutImportRequested=" & CStr(CalloutImportRequested())

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        Debug.Print "R23_CALLOUT_FATAL|reason=CatalogUnavailable"
        Exit Sub
    End If

    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueProjectionView

        Module13_ProjectionResolution.BuildViewProjections _
            swApp, swDraw, swView, graph, evidence

ContinueProjectionView:
    Next i

    Debug.Print "R23_CALLOUT_GRAPH|" & graph.GraphSummary()

    ' Pass one. Every native hole callout on the sheet, attributed to at
    ' most one family.
    Dim nativeByFamily As Object
    Set nativeByFamily = CreateObject("Scripting.Dictionary")

    Dim nativeCount As Long
    Dim attributedCount As Long

    For i = LBound(views) To UBound(views)
        Dim calloutView As SldWorks.View
        Set calloutView = views(i)
        If calloutView Is Nothing Then GoTo ContinueCalloutView

        Dim viewName As String
        viewName = SafeViewName(calloutView)

        Dim dimensions As Variant
        dimensions = calloutView.GetDisplayDimensions

        If VariantItemCount(dimensions) = 0 Then GoTo ContinueCalloutView

        Dim d As Long
        For d = LBound(dimensions) To UBound(dimensions)
            Dim swDispDim As SldWorks.DisplayDimension
            Set swDispDim = Nothing
            On Error Resume Next
            Set swDispDim = dimensions(d)
            On Error GoTo Failed

            If swDispDim Is Nothing Then GoTo ContinueDimension
            If Not IsNativeHoleCallout(swDispDim) Then GoTo ContinueDimension

            nativeCount = nativeCount + 1

            Dim matchDiagnostics As String
            Dim familyKey As String
            familyKey = MatchCalloutToFamily( _
                swApp, graph, swDispDim, viewName, matchDiagnostics)

            If Len(familyKey) = 0 Then
                Debug.Print "QA INFO: CALLOUT_UNATTRIBUTED|view=" & _
                    viewName & "|" & matchDiagnostics
                GoTo ContinueDimension
            End If

            attributedCount = attributedCount + 1

            Dim nativeDefinition As CCalloutDefinition
            Set nativeDefinition = BuildDefinitionFromTypedData( _
                graph, familyKey, evidence)

            nativeDefinition.DefinitionSource = DEFINITION_NATIVE
            nativeDefinition.OwnerViewName = viewName
            Set nativeDefinition.NativeAnnotation = swDispDim
            nativeDefinition.AttachmentProven = True
            nativeDefinition.AttachmentProofSource = matchDiagnostics

            ReadNativeCalloutFields swDispDim, nativeDefinition, evidence

            If Not nativeByFamily.Exists(familyKey) Then
                nativeByFamily.Add familyKey, nativeDefinition
            End If

            Debug.Print "QA INFO: CALLOUT_NATIVE|view=" & viewName & _
                "|" & nativeDefinition.Summary()

ContinueDimension:
        Next d

ContinueCalloutView:
    Next i

    Debug.Print "R23_CALLOUT_NATIVE_SUMMARY|nativeCallouts=" & _
        CStr(nativeCount) & _
        "|attributedToFamily=" & CStr(attributedCount) & _
        "|distinctFamilies=" & CStr(nativeByFamily.Count) & _
        "|classifier=IDisplayDimension.IsHoleCallout"

    ' Pass two. One retained definition per family in the graph.
    Dim retained As Collection
    Set retained = New Collection

    Dim familyKeys As Collection
    Set familyKeys = graph.FamilyKeys()

    Dim f As Long
    For f = 1 To familyKeys.Count
        Dim key As String
        key = CStr(familyKeys(f))

        Dim fallbackDefinition As CCalloutDefinition
        Set fallbackDefinition = BuildDefinitionFromTypedData( _
            graph, key, evidence)

        Dim nativeCandidate As CCalloutDefinition
        Set nativeCandidate = Nothing
        If nativeByFamily.Exists(key) Then
            Set nativeCandidate = nativeByFamily(key)
        End If

        Dim decisionProof As String
        Dim chosen As CCalloutDefinition
        Set chosen = RetainDefinitionForFamily( _
            nativeCandidate, fallbackDefinition, decisionProof)

        If chosen Is Nothing Then GoTo ContinueFamily

        retained.Add chosen

        Debug.Print "QA INFO: CALLOUT_DEFINITION|" & decisionProof & _
            "|" & chosen.Summary()

ContinueFamily:
    Next f

    Dim definitionVerdict As String
    definitionVerdict = VerifyManufacturingDefinitions(retained)

    Dim shapeVerdict As String
    shapeVerdict = VerifyRequiredDefinitionShapes(retained)

    Dim finalSelectionCount As Long
    finalSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim drawingSaveAfter As Boolean
    drawingSaveAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Debug.Print "R23_CALLOUT_END|" & definitionVerdict & _
        "|" & shapeVerdict & _
        "|nativeCallouts=" & CStr(nativeCount) & _
        "|creations=0" & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & _
        CStr(drawingSaveBefore = drawingSaveAfter)
    Exit Sub

Failed:
    Debug.Print "R23_CALLOUT_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub

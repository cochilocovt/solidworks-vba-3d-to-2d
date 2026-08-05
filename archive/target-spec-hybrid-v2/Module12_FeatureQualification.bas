Option Explicit

' R23 Phase 2 feature qualification.
' Walks the model feature tree, resolves each feature's effective type,
' proves its suppression state against an exact configuration, reads its
' typed definition, and turns feature-owned cylindrical geometry into
' physical locations in a CLocationGraph.
' Boundaries held by this module:
' - it never calls ModifyDefinition and never saves the model;
' - every successful AccessSelections is paired with
' ReleaseSelectionAccess on success, rejection, error and early exit;
' - ownership comes from IFeature.GetFaces, never from IFace2.GetFeature
' alone, because that reports only the oldest owner;
' - IFace2.FaceInSurfaceSense is supplementary orientation evidence only
' and never decides whether a face bounds an internal hole;
' - only exact effective "CUT" and "CUTTHIN" reach the extrude route; a
' substring match would accept unrelated feature types.
' The module changes no existing pipeline. It is consumed by the Phase 11
' reorder; until then R23_ProbeFeatureCatalog is the evidence entry point.

' Verified against the local solidworks-api MCP and installed SOLIDWORKS 2025
' SP1.2 interop.
Private Const swSpecifyConfiguration As Long = 3
Private Const swDocPART As Long = 1

' Exact IFeature.GetTypeName2 literals. Feature names are never matched.
Private Const TYPE_ICE As String = "ICE"
Private Const TYPE_HOLE_WIZARD As String = "HOLEWZD"
Private Const TYPE_ADVANCED_HOLE As String = "ADVHOLEWZD"
Private Const TYPE_SKETCH_HOLE As String = "SKETCHHOLE"
Private Const TYPE_CUT As String = "CUT"
Private Const TYPE_CUT_THIN As String = "CUTTHIN"
Private Const TYPE_COSMETIC_THREAD As String = "COSMETICTHREAD"

' Candidate members that can carry a Hole Wizard's nominal size or depth.
' Order is reporting order only; it implies no preference. CounterBore* are
' included deliberately as the control - they are already known to read
' correctly, so a dump where they are also zero would mean the probe itself
' failed rather than the members being empty.
Private Const HOLE_WIZARD_PROBE_MEMBERS As String = _
    "Type,HoleDiameter,HoleDepth,Diameter,Depth," & _
    "ThruHoleDiameter,ThruHoleDepth," & _
    "TapDrillDiameter,TapDrillDepth," & _
    "ThruTapDrillDiameter,ThruTapDrillDepth," & _
    "MajorDiameter,MinorDiameter,ThreadDiameter,ThreadDepth," & _
    "CounterBoreDiameter,CounterBoreDepth," & _
    "Standard,FastenerType,FastenerSize,HoleFit,EndCondition"

Private Const MAX_SEED_CHAIN_DEPTH As Long = 16
Private Const CURVE_ROLE_COUNTERBORE As Long = 1
Private Const CURVE_ROLE_TAPPED As Long = 2
Private Const CURVE_ROLE_MIRRORED As Long = 3
Private Const CURVE_ROLE_EXTRUDED_CUT As Long = 4
Private Const CURVE_ORDER_TOLERANCE_M As Double = 0.000000001

' When True, every recorded line is also echoed to the Immediate Window.
' CRunEvidence deliberately exposes no read-back accessor for its info list,
' so the diagnostic entry point echoes here rather than replaying evidence.
Private mEmitDiagnostics As Boolean

Private Sub EmitInfo( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddInfo message
    If mEmitDiagnostics Then
        Module21_EvidenceSink.LogLine message
    End If
End Sub

Private Sub EmitWarning( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddWarning message
    If mEmitDiagnostics Then
        Module21_EvidenceSink.LogLine message
    End If
End Sub

Private Sub EmitFailure( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddFailure message
    If mEmitDiagnostics Then
        Module21_EvidenceSink.LogLine message
    End If
End Sub

' Builds the feature catalog for one part into the supplied graph.
' referencedConfiguration must be the exact configuration the consuming
' drawing view references. Suppression that cannot be proved for that
' configuration fails the feature closed rather than assuming the active one.
Public Function BuildFeatureCatalog( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal referencedConfiguration As String, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If swPart Is Nothing Or graph Is Nothing Then Exit Function

    Dim activeConfiguration As String
    activeConfiguration = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    Dim visitedKeys As Object
    Set visitedKeys = CreateObject("Scripting.Dictionary")

    Dim visitedCount As Long
    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Do While Not swFeature Is Nothing
        VisitFeature swApp, swPart, swFeature, _
            referencedConfiguration, activeConfiguration, _
            graph, evidence, visitedKeys, visitedCount, 0

        Set swFeature = swFeature.GetNextFeature
    Loop

    EmitInfo evidence, "FEATURE_CATALOG|visitedFeatures=" & _
            CStr(visitedCount) & _
            "|configuration=" & referencedConfiguration & _
            "|activeConfiguration=" & activeConfiguration & _
            "|" & graph.GraphSummary()

    BuildFeatureCatalog = True
    Exit Function

Failed:
    EmitFailure evidence, "Feature catalog error " & _
            CStr(Err.Number) & ": " & Err.Description
End Function

' Depth-first visit with a composite cycle guard. ObjPtr alone is unsafe in
' this VBA host: transient COM wrappers reuse addresses, which silently
' suppressed 32 of 47 features in an earlier probe.
Private Sub VisitFeature( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal referencedConfiguration As String, _
    ByVal activeConfiguration As String, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence, _
    ByRef visitedKeys As Object, _
    ByRef visitedCount As Long, _
    ByVal depth As Long)

    On Error GoTo Failed

    If swFeature Is Nothing Then Exit Sub
    If depth > MAX_SEED_CHAIN_DEPTH Then Exit Sub

    Dim traversalKey As String
    traversalKey = BuildTraversalKey(swFeature)

    If visitedKeys.Exists(traversalKey) Then Exit Sub
    visitedKeys.Add traversalKey, True
    visitedCount = visitedCount + 1

    QualifyFeature swApp, swPart, swFeature, _
        referencedConfiguration, activeConfiguration, _
        graph, evidence

    Dim subFeature As SldWorks.Feature
    Set subFeature = swFeature.GetFirstSubFeature

    Do While Not subFeature Is Nothing
        VisitFeature swApp, swPart, subFeature, _
            referencedConfiguration, activeConfiguration, _
            graph, evidence, visitedKeys, visitedCount, depth + 1

        Set subFeature = subFeature.GetNextSubFeature
    Loop
    Exit Sub

Failed:
    EmitWarning evidence, "FEATURE_VISIT_ERROR|error=" & _
            CStr(Err.Number) & "|description=" & Err.Description
End Sub

Private Function BuildTraversalKey( _
    ByRef swFeature As SldWorks.Feature) As String

    On Error Resume Next

    Dim featureName As String
    Dim typeName2 As String
    featureName = swFeature.Name
    typeName2 = swFeature.GetTypeName2

    ' ObjPtr is deliberately NOT part of the key. Two live catalog runs over
    ' an unchanged P-0251 visited 47 then 46 features, and the sketches that
    ' were visited twice differed between runs: the same underlying feature
    ' reached as a tree entry and as a consuming feature's subfeature can
    ' arrive through two COM wrappers at different addresses, which splits
    ' one feature into two keys and defeats the guard. Feature names are
    ' unique within a part document, so name plus type is the exact-once
    ' key. The address is used only when the name cannot be read, where it
    ' is better to risk a repeat visit than to collapse several unnamed
    ' features into one key and skip them.
    If Len(featureName) = 0 Then
        BuildTraversalKey = "ptr=" & CStr(ObjPtr(swFeature))
        On Error GoTo 0
        Exit Function
    End If

    BuildTraversalKey = _
        Module11_GeometryIdentity.IdentityToken(featureName) & "|" & _
        Module11_GeometryIdentity.IdentityToken(typeName2)

    On Error GoTo 0
End Function

' Resolves one feature into a CFeatureDefinition and, when it carries
' manufacturing geometry, into physical locations.
Private Sub QualifyFeature( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal referencedConfiguration As String, _
    ByVal activeConfiguration As String, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim definition As CFeatureDefinition
    Set definition = New CFeatureDefinition

    Set definition.ModelFeature = swFeature
    definition.FeatureName = SafeFeatureName(swFeature)
    definition.ConfigurationName = referencedConfiguration

    Dim typeFailure As String
    If Not ResolveEffectiveType(swFeature, definition, typeFailure) Then
        RecordRejectedFeature definition, typeFailure, graph, evidence
        Exit Sub
    End If

    If Not ProveSuppressionState( _
        swFeature, definition, referencedConfiguration, _
        activeConfiguration) Then

        RecordRejectedFeature definition, _
            definition.SuppressionState, graph, evidence
        Exit Sub
    End If

    CollectOwnedFaces swFeature, definition

    ' Type routing uses the normalized effective type only.
    Select Case definition.EffectiveType
        Case TYPE_HOLE_WIZARD
            ReadHoleWizardDefinition swPart, swFeature, definition
        Case TYPE_ADVANCED_HOLE
            ReadAdvancedHoleDefinition swPart, swFeature, definition
        Case TYPE_SKETCH_HOLE
            ReadSimpleHoleDefinition swPart, swFeature, definition
        Case TYPE_CUT, TYPE_CUT_THIN
            ReadExtrudeCutDefinition swPart, swFeature, definition
        Case TYPE_COSMETIC_THREAD
            ReadCosmeticThreadDefinition swPart, swFeature, definition
        Case Else
            If Not IsSupportedPatternType(definition.EffectiveType) Then
                RecordRejectedFeature definition, _
                    "UnsupportedFeatureType", graph, evidence
                Exit Sub
            End If

            ' The seed-chain reason codes are only worth computing if they
            ' are enforced; the accept path below clears RejectionReason, so
            ' an unchecked failure here would be silently accepted.
            If Not ResolvePatternSeed( _
                swPart, swFeature, definition, evidence) Then

                RecordRejectedFeature definition, _
                    definition.RejectionReason, graph, evidence
                Exit Sub
            End If
    End Select

    If definition.OperationKind = "Boss" Then
        RecordRejectedFeature definition, "BossFeatureRejected", _
            graph, evidence
        Exit Sub
    End If

    If Not definition.HasOwnedFaces() Then
        RecordRejectedFeature definition, "NoOwnedGeometry", _
            graph, evidence
        Exit Sub
    End If

    definition.Accepted = True
    definition.RejectionReason = vbNullString
    graph.AddFeatureDefinition definition

    BuildPhysicalLocations swPart, definition, graph, evidence

    EmitInfo evidence, "FEATURE_ACCEPTED|name=" & _
            definition.FeatureName & _
            "|rawType2=" & definition.RawTypeName2 & _
            "|rawType1=" & definition.RawTypeName & _
            "|effectiveType=" & definition.EffectiveType & _
            "|seed=" & definition.SeedFeatureName & _
            "|readStatus=" & definition.DefinitionReadStatus & _
            "|" & definition.SemanticSummary()
    Exit Sub

Failed:
    EmitWarning evidence, "FEATURE_QUALIFY_ERROR|feature=" & _
            SafeFeatureName(swFeature) & _
            "|error=" & CStr(Err.Number) & _
            "|description=" & Err.Description
End Sub

' R23-200 and R23-201. Three fields are always recorded: the raw
' GetTypeName2, the raw GetTypeName, and the normalized effective type.
' GetTypeName2 of "ICE" is an Instant3D wrapper whose real type only
' GetTypeName supplies.
Private Function ResolveEffectiveType( _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition, _
    ByRef failureReason As String) As Boolean

    On Error GoTo Failed

    definition.RawTypeName2 = Trim$(swFeature.GetTypeName2)
    definition.RawTypeName = Trim$(swFeature.GetTypeName)

    Dim normalizedType2 As String
    normalizedType2 = UCase$(definition.RawTypeName2)

    If Len(normalizedType2) = 0 Then
        failureReason = "FeatureTypeUnresolved"
        definition.TypeResolutionSource = "Unresolved"
        Exit Function
    End If

    If normalizedType2 = TYPE_ICE Then
        Dim normalizedType1 As String
        normalizedType1 = UCase$(Trim$(definition.RawTypeName))

        If Len(normalizedType1) = 0 Then
            failureReason = "IceUnderlyingTypeMissing"
            definition.TypeResolutionSource = "IceUnresolved"
            Exit Function
        End If

        definition.EffectiveType = normalizedType1
        definition.TypeResolutionSource = "GetTypeName(ICE)"
    Else
        definition.EffectiveType = normalizedType2
        definition.TypeResolutionSource = "GetTypeName2"
    End If

    ResolveEffectiveType = True
    Exit Function

Failed:
    failureReason = "FeatureTypeReadError:" & CStr(Err.Number)
    definition.TypeResolutionSource = "ReadError"
End Function

' R23-203. Suppression is evaluated against the exact referenced
' configuration.
' On this installed build IFeature.IsSuppressed2 returned Empty for every
' feature, so a fallback is required. The fallback is only sound when the
' requested configuration IS the active one; otherwise the state is
' genuinely unknown and the feature fails closed.
Private Function ProveSuppressionState( _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition, _
    ByVal referencedConfiguration As String, _
    ByVal activeConfiguration As String) As Boolean

    On Error GoTo Failed

    Dim rawResult As Variant
    rawResult = swFeature.IsSuppressed2( _
        swSpecifyConfiguration, Array(referencedConfiguration))

    Dim suppressed As Boolean
    Dim resolved As Boolean

    If IsArray(rawResult) Then
        If UBound(rawResult) >= LBound(rawResult) Then
            suppressed = Module11_GeometryIdentity.NormalizeSwBoolean( _
                rawResult(LBound(rawResult)))
            resolved = True
            definition.SuppressionState = _
                IIf(suppressed, "Suppressed", "Active") & ":SpecifyArray"
        End If
    ElseIf Not IsEmpty(rawResult) And Not IsNull(rawResult) Then
        suppressed = Module11_GeometryIdentity.NormalizeSwBoolean(rawResult)
        resolved = True
        definition.SuppressionState = _
            IIf(suppressed, "Suppressed", "Active") & ":SpecifyScalar"
    End If

    If Not resolved Then
        ' Equivalence fallback, permitted only when the requested and active
        ' configurations are the same name.
        If StrComp( _
            Module11_GeometryIdentity.IdentityToken(referencedConfiguration), _
            Module11_GeometryIdentity.IdentityToken(activeConfiguration), _
            vbBinaryCompare) <> 0 Then

            definition.SuppressionProven = False
            definition.SuppressionState = "ConfigurationSuppressionUnproven"
            Exit Function
        End If

        suppressed = Module11_GeometryIdentity.NormalizeSwBoolean( _
            swFeature.IsSuppressed)
        definition.SuppressionState = _
            IIf(suppressed, "Suppressed", "Active") & _
            ":ActiveConfigurationEquivalenceFallback"
    End If

    definition.SuppressionProven = True

    If suppressed Then
        definition.SuppressionState = "FeatureSuppressedInConfiguration"
        Exit Function
    End If

    ProveSuppressionState = True
    Exit Function

Failed:
    definition.SuppressionProven = False
    definition.SuppressionState = _
        "ConfigurationSuppressionUnproven:" & CStr(Err.Number)
End Function

' R23-207. Ownership comes from the feature's own face set.
Private Sub CollectOwnedFaces( _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition)

    On Error Resume Next

    Dim faces As Variant
    faces = swFeature.GetFaces

    If IsArray(faces) Then
        Dim i As Long
        For i = LBound(faces) To UBound(faces)
            If Not faces(i) Is Nothing Then
                definition.OwnedFaces.Add faces(i)
            End If
        Next i
    End If

    On Error GoTo 0
End Sub

' Returns the first candidate member that reads a non-zero value, and names
' it in memberRoute. Each read is individually guarded because the Help
' documents several of these as relevant only for particular
' swWzdHoleTypes_e values, and an irrelevant member may raise rather than
' return zero. When every candidate declines, the route is reported as
' AllCandidatesZero so a genuinely empty read stays distinguishable from a
' successful one - the failure this whole change exists to stop hiding.
Private Function FirstNonZeroHoleMember( _
    ByRef holeData As SldWorks.WizardHoleFeatureData2, _
    ByVal memberNames As String, _
    ByRef memberRoute As String) As Double

    memberRoute = "Unavailable"
    If holeData Is Nothing Then Exit Function

    Dim members As Variant
    members = Split(memberNames, ",")

    Dim attempted As String

    Dim i As Long
    For i = LBound(members) To UBound(members)
        Dim memberName As String
        memberName = Trim$(CStr(members(i)))

        Dim raw As Variant
        Dim readError As Long
        On Error Resume Next
        raw = CallByName(holeData, memberName, VbGet)
        readError = Err.Number
        Err.Clear
        On Error GoTo 0

        If Len(attempted) > 0 Then attempted = attempted & "/"
        attempted = attempted & memberName

        If readError = 0 Then
            If Not (IsEmpty(raw) Or IsNull(raw)) Then
                If IsNumeric(raw) Then
                    If CDbl(raw) <> 0# Then
                        memberRoute = memberName
                        FirstNonZeroHoleMember = CDbl(raw)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    memberRoute = "AllCandidatesZero:" & attempted
End Function

' Reports each candidate member once. Every read is individually guarded:
' the Help states several of these are relevant only for particular
' swWzdGeneralHoleTypes_e values, and an irrelevant member may raise rather
' than return zero. A raise is recorded as its error number so "member does
' not apply here" stays distinguishable from "member returned nothing".
Private Sub EmitHoleWizardMemberProbe( _
    ByRef holeData As SldWorks.WizardHoleFeatureData2, _
    ByRef definition As CFeatureDefinition)

    On Error GoTo Failed

    If holeData Is Nothing Then Exit Sub

    Dim members As Variant
    members = Split(HOLE_WIZARD_PROBE_MEMBERS, ",")

    Dim rendered As String
    rendered = "HOLEWIZARD_MEMBER_PROBE|feature=" & _
        Module16_CalloutDefinition.EvidenceToken(definition.FeatureName) & _
        "|accessGranted=" & CStr(definition.SelectionAccessGranted)

    Dim i As Long
    For i = LBound(members) To UBound(members)
        Dim memberName As String
        memberName = Trim$(CStr(members(i)))

        Dim raw As Variant
        Dim readError As Long
        On Error Resume Next
        raw = CallByName(holeData, memberName, VbGet)
        readError = Err.Number
        Err.Clear
        On Error GoTo Failed

        Dim renderedValue As String
        If readError <> 0 Then
            renderedValue = "Error:" & CStr(readError)
        ElseIf IsEmpty(raw) Or IsNull(raw) Then
            renderedValue = "Empty"
        ElseIf IsNumeric(raw) Then
            renderedValue = Format$(CDbl(raw), "0.000000000")
        Else
            renderedValue = Module16_CalloutDefinition.EvidenceToken( _
                CStr(raw))
        End If

        rendered = rendered & "|" & memberName & "=" & renderedValue
    Next i

    If Not Module1_Main.GlobalEvidence Is Nothing Then
        Module1_Main.GlobalEvidence.AddInfo rendered
    End If
    Module21_EvidenceSink.LogLine rendered
    Exit Sub

Failed:
    If Not Module1_Main.GlobalEvidence Is Nothing Then
        Module1_Main.GlobalEvidence.AddWarning _
            "HOLEWIZARD_MEMBER_PROBE|status=Unavailable|error=" & _
            CStr(Err.Number)
    End If
End Sub

Private Sub ReadHoleWizardDefinition( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition)

    Dim accessGranted As Boolean
    Dim holeData As SldWorks.WizardHoleFeatureData2

    On Error GoTo Failed

    Set holeData = swFeature.GetDefinition
    If holeData Is Nothing Then
        definition.DefinitionReadStatus = "DefinitionUnavailable"
        Exit Sub
    End If

    definition.DefinitionInterfaceName = "IWizardHoleFeatureData2"
    Set definition.DefinitionObject = holeData

    accessGranted = Module11_GeometryIdentity.NormalizeSwBoolean( _
        holeData.AccessSelections(swPart, Nothing))
    definition.SelectionAccessGranted = accessGranted

    definition.OperationKind = "HoleWizard"

    ' Measured, run macro_qa/20260805_001154_P-0251-14A-001. HoleDiameter and
    ' HoleDepth return 0 for both P-0251 Hole Wizard features while the
    ' type-specific members carry the real values, so reading them was the
    ' whole of the all-zero defect. IWizardHoleFeatureData2.Type decides
    ' which member applies (swWzdHoleTypes_e):
    '   Type 14 swCounterBoreThru       ThruHoleDiameter 0.0066 / Depth 0.018
    '   Type 46 swTapBlindCosmeticThread ThreadDiameter  0.0050 / Depth 0.010
    ' 6.60 THRU and 11.00 by 6.00 are exactly what the reference callout
    ' shows, so this is the drawing's own nominal, not a derived guess.
    ' The chain is ordered so the two observed types are disjoint - a
    ' counterbore has no ThreadDiameter and a tapped hole has no
    ' ThruHoleDiameter - and the member that supplied the value is always
    ' named in the proof, so a future hole type that lands on a different
    ' member is visible in evidence rather than silently accepted.
    Dim diameterRoute As String
    definition.NominalDiameterM = FirstNonZeroHoleMember(holeData, _
        "ThruHoleDiameter,ThreadDiameter,TapDrillDiameter," & _
        "HoleDiameter,Diameter", diameterRoute)
    definition.DiameterProofSource = _
        "IWizardHoleFeatureData2." & diameterRoute

    Dim depthRoute As String
    definition.DepthM = FirstNonZeroHoleMember(holeData, _
        "ThruHoleDepth,ThreadDepth,TapDrillDepth," & _
        "HoleDepth,Depth", depthRoute)
    definition.DepthProofSource = "IWizardHoleFeatureData2." & depthRoute
    definition.CounterBoreDiameterM = holeData.CounterBoreDiameter
    definition.CounterBoreDepthM = holeData.CounterBoreDepth
    definition.CounterBoreProofSource = _
        "IWizardHoleFeatureData2.CounterBore*"
    definition.ThreadDescription = Trim$(CStr(holeData.FastenerSize))
    definition.ThreadDepthM = holeData.ThreadDepth
    definition.ThreadProofSource = _
        "IWizardHoleFeatureData2.FastenerSize/ThreadDepth"
    ' IWizardHoleFeatureData2.HoleFit returns swWzdHoleScrewClearanceTypes_e
    ' and the 2025 Help states it "applies to counterbore and countersink
    ' Hole Wizard features only". A tapped hole returns -1, which is outside
    ' the enum and means not applicable. Publishing it as a fit value would
    ' put a bogus clearance on a tapped-hole callout, so anything outside the
    ' enum is recorded as absent rather than as data.
    definition.FitDescription = _
        ScrewClearanceText(CLng(holeData.HoleFit))
    If Len(definition.FitDescription) > 0 Then
        definition.FitProofSource = _
            "IWizardHoleFeatureData2.HoleFit=" & CStr(holeData.HoleFit)
    End If
    definition.EndConditionCode = holeData.EndCondition
    definition.EndConditionProofSource = _
        "IWizardHoleFeatureData2.EndCondition"

    ' A read that returns nothing for every dimensional member is not a
    ' successful read. Until r46 it was recorded as "Read" and the defect
    ' only surfaced two stages later as an incomplete callout.
    If definition.NominalDiameterM = 0# And definition.DepthM = 0# And _
       definition.CounterBoreDiameterM = 0# And _
       definition.CounterBoreDepthM = 0# Then

        definition.DefinitionReadStatus = "ReadAllZeroValues:" & _
            diameterRoute & ";" & depthRoute
    Else
        definition.DefinitionReadStatus = "Read"
    End If

    ' R23 Root-2 diagnosis, read-only. HoleDiameter and HoleDepth return 0 for
    ' both P-0251 Hole Wizard features while CounterBoreDiameter and
    ' CounterBoreDepth return correct values (0.011 / 0.006), so
    ' AccessSelections plainly worked and the zero is member selection, not
    ' access. The 2025 Help makes which member carries the value depend on
    ' IWizardHoleFeatureData2.Type: TapDrillDiameter is documented as
    ' relevant only for the swTap* types, ThruHoleDiameter only for
    ' non-tapered holes. Rather than guess which one holds the 6.60 mm the
    ' drawing callout already shows, every candidate is reported once and the
    ' evidence decides. Nothing below changes what the definition publishes.
    EmitHoleWizardMemberProbe holeData, definition

    ReleaseAccess holeData, accessGranted, definition
    Exit Sub

Failed:
    definition.DefinitionReadStatus = _
        "DefinitionReadError:" & CStr(Err.Number)
    ReleaseAccess holeData, accessGranted, definition
End Sub

Private Sub ReadAdvancedHoleDefinition( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition)

    On Error GoTo Failed

    Dim advancedData As Object
    Set advancedData = swFeature.GetDefinition

    If advancedData Is Nothing Then
        definition.DefinitionReadStatus = "DefinitionUnavailable"
        Exit Sub
    End If

    definition.DefinitionInterfaceName = "IAdvancedHoleFeatureData"
    Set definition.DefinitionObject = advancedData
    definition.OperationKind = "AdvancedHole"

    ' Advanced holes expose their sizes through element collections whose
    ' contract is not yet runtime-proved on this build. The typed object is
    ' retained, but no semantic field is claimed without a proof source.
    definition.DefinitionReadStatus = "ReadPendingElementContract"
    Exit Sub

Failed:
    definition.DefinitionReadStatus = _
        "DefinitionReadError:" & CStr(Err.Number)
End Sub

Private Sub ReadSimpleHoleDefinition( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition)

    Dim accessGranted As Boolean
    Dim simpleData As SldWorks.SimpleHoleFeatureData2

    On Error GoTo Failed

    Set simpleData = swFeature.GetDefinition
    If simpleData Is Nothing Then
        definition.DefinitionReadStatus = "DefinitionUnavailable"
        Exit Sub
    End If

    definition.DefinitionInterfaceName = "ISimpleHoleFeatureData2"
    Set definition.DefinitionObject = simpleData

    accessGranted = Module11_GeometryIdentity.NormalizeSwBoolean( _
        simpleData.AccessSelections(swPart, Nothing))
    definition.SelectionAccessGranted = accessGranted

    definition.OperationKind = "SimpleHole"
    definition.NominalDiameterM = simpleData.Diameter
    definition.DiameterProofSource = "ISimpleHoleFeatureData2.Diameter"
    definition.DepthM = simpleData.Depth
    definition.DepthProofSource = "ISimpleHoleFeatureData2.Depth"
    definition.EndConditionCode = simpleData.Type
    definition.EndConditionProofSource = "ISimpleHoleFeatureData2.Type"
    definition.DefinitionReadStatus = "Read"

    ReleaseAccess simpleData, accessGranted, definition
    Exit Sub

Failed:
    definition.DefinitionReadStatus = _
        "DefinitionReadError:" & CStr(Err.Number)
    ReleaseAccess simpleData, accessGranted, definition
End Sub

' R23-209. Only exact CUT and CUTTHIN reach this route, and a cut is only
' accepted as a manufacturing cut when IsBossFeature is False.
Private Sub ReadExtrudeCutDefinition( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition)

    Dim accessGranted As Boolean
    Dim extrudeData As SldWorks.ExtrudeFeatureData2

    On Error GoTo Failed

    Set extrudeData = swFeature.GetDefinition
    If extrudeData Is Nothing Then
        definition.DefinitionReadStatus = "DefinitionUnavailable"
        Exit Sub
    End If

    definition.DefinitionInterfaceName = "IExtrudeFeatureData2"
    Set definition.DefinitionObject = extrudeData

    accessGranted = Module11_GeometryIdentity.NormalizeSwBoolean( _
        extrudeData.AccessSelections(swPart, Nothing))
    definition.SelectionAccessGranted = accessGranted

    definition.IsBossFeature = _
        Module11_GeometryIdentity.NormalizeSwBoolean( _
            extrudeData.IsBossFeature)
    definition.BossProofSource = "IExtrudeFeatureData2.IsBossFeature"

    If definition.IsBossFeature Then
        definition.OperationKind = "Boss"
        definition.DefinitionReadStatus = "ReadRejectedAsBoss"
        ReleaseAccess extrudeData, accessGranted, definition
        Exit Sub
    End If

    definition.OperationKind = "ExtrudedCut"
    definition.DepthM = extrudeData.GetDepth(True)
    definition.DepthProofSource = "IExtrudeFeatureData2.GetDepth(True)"
    definition.EndConditionCode = extrudeData.GetEndCondition(True)
    definition.EndConditionProofSource = _
        "IExtrudeFeatureData2.GetEndCondition(True)"
    definition.DefinitionReadStatus = "Read"

    ReleaseAccess extrudeData, accessGranted, definition
    Exit Sub

Failed:
    definition.DefinitionReadStatus = _
        "DefinitionReadError:" & CStr(Err.Number)
    ReleaseAccess extrudeData, accessGranted, definition
End Sub

Private Sub ReadCosmeticThreadDefinition( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition)

    Dim accessGranted As Boolean
    Dim threadData As SldWorks.CosmeticThreadFeatureData

    On Error GoTo Failed

    Set threadData = swFeature.GetDefinition
    If threadData Is Nothing Then
        definition.DefinitionReadStatus = "DefinitionUnavailable"
        Exit Sub
    End If

    definition.DefinitionInterfaceName = "ICosmeticThreadFeatureData"
    Set definition.DefinitionObject = threadData

    accessGranted = Module11_GeometryIdentity.NormalizeSwBoolean( _
        threadData.AccessSelections(swPart, Nothing))
    definition.SelectionAccessGranted = accessGranted

    definition.OperationKind = "CosmeticThread"
    definition.NominalDiameterM = threadData.Diameter
    definition.DiameterProofSource = "ICosmeticThreadFeatureData.Diameter"
    definition.ThreadDescription = Trim$(CStr(threadData.ThreadCallout))
    definition.ThreadDepthM = threadData.BlindDepth
    definition.ThreadProofSource = _
        "ICosmeticThreadFeatureData.ThreadCallout/BlindDepth"
    definition.EndConditionCode = threadData.EndCondition
    definition.EndConditionProofSource = _
        "ICosmeticThreadFeatureData.EndCondition"
    definition.DefinitionReadStatus = "Read"

    ReleaseAccess threadData, accessGranted, definition
    Exit Sub

Failed:
    definition.DefinitionReadStatus = _
        "DefinitionReadError:" & CStr(Err.Number)
    ReleaseAccess threadData, accessGranted, definition
End Sub

' R23-205. One release path used by every reader, so a successful access is
' released on success, rejection and error alike.
Private Sub ReleaseAccess( _
    ByVal featureData As Object, _
    ByVal accessGranted As Boolean, _
    ByRef definition As CFeatureDefinition)

    If Not accessGranted Then Exit Sub
    If featureData Is Nothing Then Exit Sub

    On Error Resume Next
    featureData.ReleaseSelectionAccess
    definition.SelectionAccessReleased = (Err.Number = 0)
    On Error GoTo 0
End Sub

' R23-211. Pattern and mirror instances resolve through their seed feature.
Private Function ResolvePatternSeed( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As CFeatureDefinition, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    definition.OperationKind = "PatternInstance"

    Dim resolvedSeedName As String
    Dim resolvedCount As Long
    Dim i As Long

    For i = 1 To definition.OwnedFaces.Count
        Dim ownedFace As SldWorks.Face2
        Set ownedFace = definition.OwnedFaces(i)
        If ownedFace Is Nothing Then GoTo ContinueFace

        Dim seedFeature As SldWorks.Feature
        Set seedFeature = Nothing
        On Error Resume Next
        Set seedFeature = ownedFace.GetSeedFeature
        On Error GoTo Failed

        If Not seedFeature Is Nothing Then
            Dim seedName As String
            seedName = SafeFeatureName(seedFeature)

            If Len(resolvedSeedName) = 0 Then
                resolvedSeedName = seedName
                Set definition.SeedFeature = seedFeature
                definition.SeedFeatureName = seedName
                resolvedCount = 1
            ElseIf StrComp(resolvedSeedName, seedName, _
                vbTextCompare) <> 0 Then

                resolvedCount = resolvedCount + 1
            End If
        End If

ContinueFace:
    Next i

    If resolvedCount = 0 Then
        definition.RejectionReason = "SeedChainUnresolved"
        Exit Function
    End If

    If resolvedCount > 1 Then
        definition.RejectionReason = "SeedChainMultiplyResolved"
        Exit Function
    End If

    If StrComp(definition.SeedFeatureName, definition.FeatureName, _
        vbTextCompare) = 0 Then

        definition.RejectionReason = "SeedChainCircular"
        Exit Function
    End If

    definition.DefinitionReadStatus = "SeedResolved"

    ' The instance has no manufacturing definition of its own; it is the
    ' seed's. Without inheriting it the instance forms its own empty
    ' PatternInstance family, which is what split the four M5 tapped holes
    ' into two families of two on the first catalog run.
    If Not InheritSeedSemantics(swPart, definition, evidence) Then
        Exit Function
    End If

    ResolvePatternSeed = True
    Exit Function

Failed:
    definition.RejectionReason = _
        "SeedChainReadError:" & CStr(Err.Number)
End Function

' R23-211. Reads the seed through the same typed readers that qualify it in
' its own right, so the inherited values carry the seed's proof sources
' rather than newly asserted ones. A seed that is itself a pattern is
' refused: following the chain needs recursion whose termination this build
' gives no evidence for, and P-0251 has no nested pattern.
Private Function InheritSeedSemantics( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef definition As CFeatureDefinition, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If definition.SeedFeature Is Nothing Then
        definition.RejectionReason = "SeedFeatureUnavailable"
        Exit Function
    End If

    Dim seedDefinition As CFeatureDefinition
    Set seedDefinition = New CFeatureDefinition

    Set seedDefinition.ModelFeature = definition.SeedFeature
    seedDefinition.FeatureName = definition.SeedFeatureName
    seedDefinition.ConfigurationName = definition.ConfigurationName

    Dim typeFailure As String
    If Not ResolveEffectiveType( _
        definition.SeedFeature, seedDefinition, typeFailure) Then

        definition.RejectionReason = "SeedTypeUnresolved:" & typeFailure
        Exit Function
    End If

    Select Case seedDefinition.EffectiveType
        Case TYPE_HOLE_WIZARD
            ReadHoleWizardDefinition swPart, _
                definition.SeedFeature, seedDefinition
        Case TYPE_ADVANCED_HOLE
            ReadAdvancedHoleDefinition swPart, _
                definition.SeedFeature, seedDefinition
        Case TYPE_SKETCH_HOLE
            ReadSimpleHoleDefinition swPart, _
                definition.SeedFeature, seedDefinition
        Case TYPE_CUT, TYPE_CUT_THIN
            ReadExtrudeCutDefinition swPart, _
                definition.SeedFeature, seedDefinition
        Case TYPE_COSMETIC_THREAD
            ReadCosmeticThreadDefinition swPart, _
                definition.SeedFeature, seedDefinition
        Case Else
            If IsSupportedPatternType(seedDefinition.EffectiveType) Then
                definition.RejectionReason = "SeedIsPatternChainUnsupported"
            Else
                definition.RejectionReason = "SeedTypeUnsupported:" & _
                    seedDefinition.EffectiveType
            End If
            Exit Function
    End Select

    If seedDefinition.OperationKind = "Boss" Then
        definition.RejectionReason = "SeedIsBossFeature"
        Exit Function
    End If

    CopySeedSemantics seedDefinition, definition
    InheritSeedSemantics = True
    Exit Function

Failed:
    definition.RejectionReason = _
        "SeedSemanticsReadError:" & CStr(Err.Number)
End Function

' Copies the seed's typed manufacturing values onto the instance so both fall
' in one semantic family. Proof sources are carried across prefixed with the
' seed name, never invented: a field the seed could not prove stays unproven
' on the instance.
Private Sub CopySeedSemantics( _
    ByRef seedDefinition As CFeatureDefinition, _
    ByRef definition As CFeatureDefinition)

    Dim provenance As String
    provenance = "SeedInherited(" & seedDefinition.FeatureName & "):"

    definition.OperationKind = seedDefinition.OperationKind
    definition.DefinitionInterfaceName = _
        seedDefinition.DefinitionInterfaceName
    definition.DefinitionReadStatus = "SeedSemanticsInherited:" & _
        seedDefinition.DefinitionReadStatus

    definition.NominalDiameterM = seedDefinition.NominalDiameterM
    definition.DepthM = seedDefinition.DepthM
    definition.EndConditionCode = seedDefinition.EndConditionCode
    definition.CounterBoreDiameterM = seedDefinition.CounterBoreDiameterM
    definition.CounterBoreDepthM = seedDefinition.CounterBoreDepthM
    definition.ThreadDescription = seedDefinition.ThreadDescription
    definition.ThreadDepthM = seedDefinition.ThreadDepthM
    definition.FitDescription = seedDefinition.FitDescription
    definition.ToleranceMinM = seedDefinition.ToleranceMinM
    definition.ToleranceMaxM = seedDefinition.ToleranceMaxM

    definition.DiameterProofSource = _
        PrefixProofSource(provenance, seedDefinition.DiameterProofSource)
    definition.DepthProofSource = _
        PrefixProofSource(provenance, seedDefinition.DepthProofSource)
    definition.EndConditionProofSource = PrefixProofSource( _
        provenance, seedDefinition.EndConditionProofSource)
    definition.CounterBoreProofSource = PrefixProofSource( _
        provenance, seedDefinition.CounterBoreProofSource)
    definition.ThreadProofSource = _
        PrefixProofSource(provenance, seedDefinition.ThreadProofSource)
    definition.FitProofSource = _
        PrefixProofSource(provenance, seedDefinition.FitProofSource)
    definition.ToleranceProofSource = PrefixProofSource( _
        provenance, seedDefinition.ToleranceProofSource)
End Sub

Private Function PrefixProofSource( _
    ByVal provenance As String, _
    ByVal proofSource As String) As String

    If Len(Trim$(proofSource)) = 0 Then Exit Function
    PrefixProofSource = provenance & proofSource
End Function

' Maps swWzdHoleScrewClearanceTypes_e to text. Values outside the enum mean
' the property does not apply to this hole type and yield an empty string,
' which the caller records as absent rather than as a fit value.
Private Function ScrewClearanceText(ByVal fitCode As Long) As String
    Select Case fitCode
        Case 0
            ScrewClearanceText = "Close"
        Case 1
            ScrewClearanceText = "Normal"
        Case 2
            ScrewClearanceText = "Loose"
    End Select
End Function

Private Function IsSupportedPatternType( _
    ByVal effectiveType As String) As Boolean

    ' Exact GetTypeName2 pattern literals confirmed for SOLIDWORKS 2025.
    Select Case effectiveType
        Case "MIRRORPATTERN", "MIRRORSOLID", "LPATTERN", "CIRPATTERN", _
             "CURVEPATTERN", "TABLEPATTERN", "APATTERN", "DIMPATTERN", _
             "SKETCHPATTERN", "DERIVEDCIRPATTERN", "DERIVEDHOLEPATTERN", _
             "DERIVEDLPATTERN", "LOCALCHAINPATTERN", "LOCALCIRPATTERN", _
             "LOCALCURVEPATTERN", "LOCALLPATTERN", "LOCALSKETCHPATTERN"
            IsSupportedPatternType = True
    End Select
End Function

' R23-212. Every feature-owned cylindrical face becomes a location
' candidate; the graph consolidates coaxial candidates whose axial intervals
' meet, which is what turns a counterbore plus its through hole, or two
' separate cut extrudes, into one stepped-bore stack.
Private Sub BuildPhysicalLocations( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef definition As CFeatureDefinition, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim i As Long
    For i = 1 To definition.OwnedFaces.Count
        Dim ownedFace As SldWorks.Face2
        Set ownedFace = definition.OwnedFaces(i)
        If ownedFace Is Nothing Then GoTo ContinueFace

        Dim originX As Double
        Dim originY As Double
        Dim originZ As Double
        Dim axisX As Double
        Dim axisY As Double
        Dim axisZ As Double
        Dim radiusM As Double
        Dim axialMinM As Double
        Dim axialMaxM As Double
        Dim geometryFailure As String

        If Not ReadCylindricalFaceGeometry( _
            ownedFace, originX, originY, originZ, _
            axisX, axisY, axisZ, radiusM, _
            axialMinM, axialMaxM, geometryFailure) Then

            GoTo ContinueFace
        End If

        Dim candidate As CPhysicalHoleLocation
        Set candidate = New CPhysicalHoleLocation
        candidate.ConfigurationName = definition.ConfigurationName
        candidate.ComponentName = vbNullString
        candidate.ModelPathName = swPart.GetPathName

        If candidate.SetAxisFromSample( _
            originX, originY, originZ, axisX, axisY, axisZ) Then

            candidate.AddStackMember definition, radiusM, _
                axialMinM, axialMaxM
            candidate.AddSourceFace ownedFace
            candidate.RecomputeKeys

            Dim resolved As CPhysicalHoleLocation
            Set resolved = graph.ResolveOrCreatePhysicalLocation(candidate)

            If Not resolved Is Nothing Then
                graph.RegisterFaceOwner _
                    BuildCylindricalFaceKey(resolved.LineKey, radiusM), _
                    definition
            End If
        End If

ContinueFace:
    Next i
    Exit Sub

Failed:
    EmitWarning evidence, "PHYSICAL_LOCATION_ERROR|feature=" & _
            definition.FeatureName & _
            "|error=" & CStr(Err.Number) & _
            "|description=" & Err.Description
End Sub

Public Function BuildCylindricalFaceKey( _
    ByVal lineKey As String, _
    ByVal radiusM As Double) As String

    BuildCylindricalFaceKey = lineKey & "|radius=" & _
        Module11_GeometryIdentity.CanonicalRadiusToken(radiusM)
End Function

' Reads a cylindrical face's axis, radius and axial extent. The extent comes
' from the face's own boundary edges, so a blind hole and an opposite blind
' hole on the same infinite line produce disjoint intervals.
Private Function ReadCylindricalFaceGeometry( _
    ByRef swFace As SldWorks.Face2, _
    ByRef originX As Double, _
    ByRef originY As Double, _
    ByRef originZ As Double, _
    ByRef axisX As Double, _
    ByRef axisY As Double, _
    ByRef axisZ As Double, _
    ByRef radiusM As Double, _
    ByRef axialMinM As Double, _
    ByRef axialMaxM As Double, _
    ByRef failureReason As String) As Boolean

    On Error GoTo Failed

    failureReason = "NotEvaluated"

    Dim surface As SldWorks.surface
    Set surface = swFace.GetSurface
    If surface Is Nothing Then
        failureReason = "SurfaceUnavailable"
        Exit Function
    End If

    If Not Module11_GeometryIdentity.NormalizeSwBoolean( _
        surface.IsCylinder) Then

        failureReason = "NotCylindrical"
        Exit Function
    End If

    Dim cylinderData As Variant
    cylinderData = surface.CylinderParams

    If Not IsArray(cylinderData) Then
        failureReason = "CylinderParamsUnavailable"
        Exit Function
    End If

    Dim baseIndex As Long
    baseIndex = LBound(cylinderData)

    If (UBound(cylinderData) - baseIndex + 1) <> 7 Then
        failureReason = "CylinderParamsWrongSize"
        Exit Function
    End If

    originX = CDbl(cylinderData(baseIndex))
    originY = CDbl(cylinderData(baseIndex + 1))
    originZ = CDbl(cylinderData(baseIndex + 2))
    radiusM = Abs(CDbl(cylinderData(baseIndex + 6)))

    If Not Module11_GeometryIdentity.NormalizeAxisDirection( _
        CDbl(cylinderData(baseIndex + 3)), _
        CDbl(cylinderData(baseIndex + 4)), _
        CDbl(cylinderData(baseIndex + 5)), _
        axisX, axisY, axisZ) Then

        failureReason = "DegenerateCylinderAxis"
        Exit Function
    End If

    If Not ComputeFaceAxialInterval( _
        swFace, axisX, axisY, axisZ, axialMinM, axialMaxM) Then

        failureReason = "AxialIntervalUnavailable"
        Exit Function
    End If

    failureReason = vbNullString
    ReadCylindricalFaceGeometry = True
    Exit Function

Failed:
    failureReason = "CylinderReadError:" & CStr(Err.Number)
End Function

Private Function ComputeFaceAxialInterval( _
    ByRef swFace As SldWorks.Face2, _
    ByVal axisX As Double, _
    ByVal axisY As Double, _
    ByVal axisZ As Double, _
    ByRef axialMinM As Double, _
    ByRef axialMaxM As Double) As Boolean

    On Error GoTo Failed

    Dim edges As Variant
    edges = swFace.GetEdges
    If Not IsArray(edges) Then Exit Function

    Dim initialized As Boolean
    Dim i As Long

    For i = LBound(edges) To UBound(edges)
        Dim swEdge As SldWorks.Edge
        Set swEdge = edges(i)
        If swEdge Is Nothing Then GoTo ContinueEdge

        Dim curveParameters As SldWorks.CurveParamData
        Set curveParameters = Nothing
        On Error Resume Next
        Set curveParameters = swEdge.GetCurveParams3
        On Error GoTo Failed
        If curveParameters Is Nothing Then GoTo ContinueEdge

        AccumulateAxialPoint curveParameters.StartPoint, _
            axisX, axisY, axisZ, axialMinM, axialMaxM, initialized
        AccumulateAxialPoint curveParameters.EndPoint, _
            axisX, axisY, axisZ, axialMinM, axialMaxM, initialized

ContinueEdge:
    Next i

    ComputeFaceAxialInterval = initialized
    Exit Function

Failed:
    ComputeFaceAxialInterval = False
End Function

Private Sub AccumulateAxialPoint( _
    ByVal pointData As Variant, _
    ByVal axisX As Double, _
    ByVal axisY As Double, _
    ByVal axisZ As Double, _
    ByRef axialMinM As Double, _
    ByRef axialMaxM As Double, _
    ByRef initialized As Boolean)

    If Not IsArray(pointData) Then Exit Sub

    Dim baseIndex As Long
    baseIndex = LBound(pointData)
    If (UBound(pointData) - baseIndex + 1) < 3 Then Exit Sub

    Dim axialValue As Double
    axialValue = Module11_GeometryIdentity.AxialParameter( _
        CDbl(pointData(baseIndex)), _
        CDbl(pointData(baseIndex + 1)), _
        CDbl(pointData(baseIndex + 2)), _
        axisX, axisY, axisZ)

    If Not initialized Then
        axialMinM = axialValue
        axialMaxM = axialValue
        initialized = True
    Else
        If axialValue < axialMinM Then axialMinM = axialValue
        If axialValue > axialMaxM Then axialMaxM = axialValue
    End If
End Sub

Private Sub RecordRejectedFeature( _
    ByRef definition As CFeatureDefinition, _
    ByVal reasonCode As String, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence)

    definition.Accepted = False
    definition.RejectionReason = reasonCode

    EmitInfo evidence, "FEATURE_REJECTED|name=" & _
            definition.FeatureName & _
            "|rawType2=" & definition.RawTypeName2 & _
            "|rawType1=" & definition.RawTypeName & _
            "|effectiveType=" & definition.EffectiveType & _
            "|reason=" & reasonCode
End Sub

Private Function SafeFeatureName( _
    ByRef swFeature As SldWorks.Feature) As String

    On Error Resume Next
    If swFeature Is Nothing Then
        SafeFeatureName = "Nothing"
    Else
        SafeFeatureName = swFeature.Name
    End If
    On Error GoTo 0
End Function

' R23-213. Confirms the expected P-0251 catalog shape. Returns a failure
' summary; an empty result means every expectation held.
Public Function VerifyExpectedCatalog( _
    ByRef graph As CLocationGraph, _
    ByVal expectedSmallHoleFamilies As String) As String

    Dim failures As String

    If graph Is Nothing Then
        VerifyExpectedCatalog = "GraphMissing"
        Exit Function
    End If

    Dim sixCount As Long
    Dim fourCount As Long
    Dim steppedStacks As Long

    Dim familyKeys As Collection
    Set familyKeys = graph.FamilyKeys()

    Dim i As Long
    For i = 1 To familyKeys.Count
        Dim familyBucket As Collection
        Set familyBucket = graph.LocationsForFamily(CStr(familyKeys(i)))

        If familyBucket.Count = 6 Then sixCount = sixCount + 1
        If familyBucket.Count = 4 Then fourCount = fourCount + 1
    Next i

    Dim locations As Collection
    Set locations = graph.Locations()

    For i = 1 To locations.Count
        Dim location As CPhysicalHoleLocation
        Set location = locations(i)
        If location.StackMemberCount() >= 2 Then
            If Not Module11_GeometryIdentity.RadiiMatch( _
                location.PrimaryRadiusM, location.MaximumRadiusM) Then

                steppedStacks = steppedStacks + 1
            End If
        End If
    Next i

    If InStr(1, expectedSmallHoleFamilies, "6", vbTextCompare) > 0 Then
        If sixCount < 1 Then
            failures = AppendFailure(failures, "NoSixLocationFamily")
        End If
    End If

    If InStr(1, expectedSmallHoleFamilies, "4", vbTextCompare) > 0 Then
        If fourCount < 1 Then
            failures = AppendFailure(failures, "NoFourLocationFamily")
        End If
    End If

    If steppedStacks < 1 Then
        failures = AppendFailure(failures, "NoSteppedBoreStack")
    End If

    VerifyExpectedCatalog = failures
End Function

Private Function AppendFailure( _
    ByVal current As String, _
    ByVal reasonCode As String) As String

    If Len(current) > 0 Then
        AppendFailure = current & "," & reasonCode
    Else
        AppendFailure = reasonCode
    End If
End Function

' R23-006 read-only confirmation. The production code has one order, but the
' Phase 0 gate must prove that both historical orders return equivalent data
' on one owned circular boundary for each representative feature role.
Private Function ProbeCurveReadOrders( _
    ByRef graph As CLocationGraph) As String

    Dim failures As String
    Dim roleFailure As String

    roleFailure = ProbeCurveRole(graph, CURVE_ROLE_COUNTERBORE)
    If Len(roleFailure) > 0 Then failures = _
        AppendFailure(failures, roleFailure)

    roleFailure = ProbeCurveRole(graph, CURVE_ROLE_TAPPED)
    If Len(roleFailure) > 0 Then failures = _
        AppendFailure(failures, roleFailure)

    roleFailure = ProbeCurveRole(graph, CURVE_ROLE_MIRRORED)
    If Len(roleFailure) > 0 Then failures = _
        AppendFailure(failures, roleFailure)

    roleFailure = ProbeCurveRole(graph, CURVE_ROLE_EXTRUDED_CUT)
    If Len(roleFailure) > 0 Then failures = _
        AppendFailure(failures, roleFailure)

    Module21_EvidenceSink.LogLine "R23_CURVE_ORDER_END|failures=" & _
        IIf(Len(failures) = 0, "None", failures)
    ProbeCurveReadOrders = failures
End Function

Private Function ProbeCurveRole( _
    ByRef graph As CLocationGraph, _
    ByVal roleCode As Long) As String

    Dim roleName As String
    roleName = CurveRoleName(roleCode)

    Dim features As Collection
    Set features = graph.Features()

    Dim i As Long
    For i = 1 To features.Count
        Dim definition As CFeatureDefinition
        Set definition = features(i)
        If definition Is Nothing Then GoTo ContinueFeature
        If Not DefinitionMatchesCurveRole(definition, roleCode) Then
            GoTo ContinueFeature
        End If

        Dim representativeFound As Boolean
        Dim orderMismatch As Boolean
        representativeFound = False
        orderMismatch = False

        If ProbeFeatureCircularEdge( _
            definition, roleName, representativeFound, orderMismatch) Then

            Exit Function
        End If

        If orderMismatch Then
            ProbeCurveRole = "CurveOrderMismatch:" & roleName
            Module21_EvidenceSink.LogLine "R23_CURVE_ORDER|role=" & _
                roleName & "|status=Fail|reason=" & ProbeCurveRole
            Exit Function
        End If

ContinueFeature:
    Next i

    ProbeCurveRole = "NoCircularRepresentative:" & roleName
    Module21_EvidenceSink.LogLine "R23_CURVE_ORDER|role=" & roleName & _
        "|status=Fail|reason=" & ProbeCurveRole
End Function

Private Function CurveRoleName(ByVal roleCode As Long) As String
    Select Case roleCode
        Case CURVE_ROLE_COUNTERBORE
            CurveRoleName = "Counterbore"
        Case CURVE_ROLE_TAPPED
            CurveRoleName = "Tapped"
        Case CURVE_ROLE_MIRRORED
            CurveRoleName = "Mirrored"
        Case CURVE_ROLE_EXTRUDED_CUT
            CurveRoleName = "ExtrudedCut"
        Case Else
            CurveRoleName = "Unknown"
    End Select
End Function

Private Function DefinitionMatchesCurveRole( _
    ByRef definition As CFeatureDefinition, _
    ByVal roleCode As Long) As Boolean

    If definition Is Nothing Then Exit Function

    Select Case roleCode
        Case CURVE_ROLE_COUNTERBORE
            DefinitionMatchesCurveRole = _
                definition.CounterBoreDiameterM > 0#
        Case CURVE_ROLE_TAPPED
            DefinitionMatchesCurveRole = _
                Len(Trim$(definition.ThreadDescription)) > 0 And _
                definition.CounterBoreDiameterM <= 0# And _
                StrComp(definition.EffectiveType, "MIRRORPATTERN", _
                    vbTextCompare) <> 0
        Case CURVE_ROLE_MIRRORED
            DefinitionMatchesCurveRole = _
                StrComp(definition.EffectiveType, "MIRRORPATTERN", _
                    vbTextCompare) = 0
        Case CURVE_ROLE_EXTRUDED_CUT
            DefinitionMatchesCurveRole = _
                StrComp(definition.OperationKind, "ExtrudedCut", _
                    vbTextCompare) = 0
    End Select
End Function

Private Function ProbeFeatureCircularEdge( _
    ByRef definition As CFeatureDefinition, _
    ByVal roleName As String, _
    ByRef representativeFound As Boolean, _
    ByRef orderMismatch As Boolean) As Boolean

    representativeFound = False
    orderMismatch = False

    Dim faceIndex As Long
    For faceIndex = 1 To definition.OwnedFaces.Count
        Dim swFace As SldWorks.Face2
        Set swFace = definition.OwnedFaces(faceIndex)
        If swFace Is Nothing Then GoTo ContinueFace

        Dim edges As Variant
        edges = swFace.GetEdges
        If Not IsArray(edges) Then GoTo ContinueFace

        Dim edgeIndex As Long
        For edgeIndex = LBound(edges) To UBound(edges)
            Dim swEdge As SldWorks.Edge
            Set swEdge = edges(edgeIndex)
            If swEdge Is Nothing Then GoTo ContinueEdge

            Dim r22Read As Boolean
            Dim retainedRead As Boolean
            Dim r22Circle As Boolean
            Dim retainedCircle As Boolean
            Dim r22Raw As String
            Dim retainedRaw As String
            Dim r22Items As Long
            Dim retainedItems As Long
            Dim r22RadiusM As Double
            Dim retainedRadiusM As Double
            Dim r22ClosureM As Double
            Dim retainedClosureM As Double
            Dim r22Reason As String
            Dim retainedReason As String

            r22Read = ReadCurveInOrder( _
                swEdge, True, r22Circle, r22Raw, r22Items, r22RadiusM, _
                r22ClosureM, r22Reason)
            retainedRead = ReadCurveInOrder( _
                swEdge, False, retainedCircle, retainedRaw, retainedItems, _
                retainedRadiusM, retainedClosureM, retainedReason)

            If Not r22Circle And Not retainedCircle Then GoTo ContinueEdge

            representativeFound = True

            Dim matched As Boolean
            matched = CurveOrderResultsMatch( _
                r22Read, retainedRead, r22Circle, retainedCircle, _
                r22Items, retainedItems, r22RadiusM, retainedRadiusM, _
                r22ClosureM, retainedClosureM)

            Module21_EvidenceSink.LogLine "R23_CURVE_ORDER|role=" & _
                roleName & "|feature=" & CurveOrderToken( _
                    definition.FeatureName) & _
                "|face=" & CStr(faceIndex) & _
                "|edge=" & CStr(edgeIndex) & _
                "|status=" & IIf(matched, "Pass", "Fail") & _
                "|r22Raw=" & r22Raw & _
                "|retainedRaw=" & retainedRaw & _
                "|r22Items=" & CStr(r22Items) & _
                "|retainedItems=" & CStr(retainedItems) & _
                "|r22RadiusM=" & Format$(r22RadiusM, "0.000000000") & _
                "|retainedRadiusM=" & _
                    Format$(retainedRadiusM, "0.000000000") & _
                "|r22ClosureM=" & Format$(r22ClosureM, "0.000000000") & _
                "|retainedClosureM=" & _
                    Format$(retainedClosureM, "0.000000000") & _
                "|r22Reason=" & r22Reason & _
                "|retainedReason=" & retainedReason

            If Not matched Then orderMismatch = True
            ProbeFeatureCircularEdge = matched
            Exit Function

ContinueEdge:
        Next edgeIndex

ContinueFace:
    Next faceIndex
End Function

Private Function ReadCurveInOrder( _
    ByRef swEdge As SldWorks.Edge, _
    ByVal r22Order As Boolean, _
    ByRef isCircle As Boolean, _
    ByRef rawCircle As String, _
    ByRef circleItems As Long, _
    ByRef radiusM As Double, _
    ByRef closureM As Double, _
    ByRef reason As String) As Boolean

    On Error GoTo Failed

    rawCircle = "Unread"
    circleItems = -1
    radiusM = -1#
    closureM = -1#
    reason = "None"

    Dim swCurve As SldWorks.Curve
    Set swCurve = swEdge.GetCurve
    If swCurve Is Nothing Then
        reason = "CurveUnavailable"
        Exit Function
    End If

    Dim curveParams As SldWorks.CurveParamData
    If r22Order Then Set curveParams = swEdge.GetCurveParams3

    Dim rawIsCircle As Variant
    rawIsCircle = swCurve.IsCircle
    rawCircle = CStr(rawIsCircle)
    isCircle = Module11_GeometryIdentity.NormalizeSwBoolean(rawIsCircle)

    Dim circleData As Variant
    circleData = swCurve.CircleParams
    circleItems = Module8_RuntimeSupport.CountVariantItems(circleData)
    If circleItems = 7 Then
        radiusM = Abs(CDbl(circleData(LBound(circleData) + 6)))
    End If

    If Not r22Order Then Set curveParams = swEdge.GetCurveParams3
    If curveParams Is Nothing Then
        reason = "CurveParamsNothing"
        ReadCurveInOrder = True
        Exit Function
    End If

    closureM = CurveParamClosureM(curveParams)
    If closureM < 0# Then reason = "ClosureUnavailable"
    ReadCurveInOrder = True
    Exit Function

Failed:
    reason = "ReadError:" & CStr(Err.Number)
End Function

Private Function CurveParamClosureM( _
    ByRef curveParams As SldWorks.CurveParamData) As Double

    On Error GoTo Failed

    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = curveParams.StartPoint
    endPoint = curveParams.EndPoint

    If Module8_RuntimeSupport.CountVariantItems( _
        startPoint) < 3 Then GoTo Failed
    If Module8_RuntimeSupport.CountVariantItems(endPoint) < 3 Then GoTo Failed

    Dim startIndex As Long
    Dim endIndex As Long
    startIndex = LBound(startPoint)
    endIndex = LBound(endPoint)

    CurveParamClosureM = Sqr( _
        (CDbl(startPoint(startIndex)) - CDbl(endPoint(endIndex))) ^ 2 + _
        (CDbl(startPoint(startIndex + 1)) - _
            CDbl(endPoint(endIndex + 1))) ^ 2 + _
        (CDbl(startPoint(startIndex + 2)) - _
            CDbl(endPoint(endIndex + 2))) ^ 2)
    Exit Function

Failed:
    CurveParamClosureM = -1#
End Function

Private Function CurveOrderResultsMatch( _
    ByVal r22Read As Boolean, _
    ByVal retainedRead As Boolean, _
    ByVal r22Circle As Boolean, _
    ByVal retainedCircle As Boolean, _
    ByVal r22Items As Long, _
    ByVal retainedItems As Long, _
    ByVal r22RadiusM As Double, _
    ByVal retainedRadiusM As Double, _
    ByVal r22ClosureM As Double, _
    ByVal retainedClosureM As Double) As Boolean

    If Not r22Read Or Not retainedRead Then Exit Function
    If Not r22Circle Or Not retainedCircle Then Exit Function
    If r22Items <> 7 Or retainedItems <> 7 Then Exit Function
    If Abs(r22RadiusM - retainedRadiusM) > _
       CURVE_ORDER_TOLERANCE_M Then Exit Function
    If r22ClosureM < 0# Or retainedClosureM < 0# Then Exit Function
    If Abs(r22ClosureM - retainedClosureM) > _
       CURVE_ORDER_TOLERANCE_M Then Exit Function

    CurveOrderResultsMatch = True
End Function

Private Function CurveOrderToken(ByVal value As String) As String
    value = Replace$(value, "|", "/")
    value = Replace$(value, "=", ":")
    CurveOrderToken = value
End Function

' Read-only evidence entry point for the Phase 2 gate.
' Builds the catalog against the active authorized part and prints the
' result. It creates no drawing, mutates nothing and never saves, so it can
' be run before the Phase 11 pipeline reorder consumes BuildFeatureCatalog.
Public Sub R23_ProbeFeatureCatalog()
    On Error GoTo Failed

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_CATALOG_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = swApp.ActiveDoc
    If swPart Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_CATALOG_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swPart.GetType <> swDocPART Then
        Module21_EvidenceSink.LogLine _
            "R23_CATALOG_FATAL|reason=ActiveDocumentNotPart"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Module21_EvidenceSink.LogLine _
            "R23_CATALOG_FATAL|reason=UnauthorizedFixture" & _
            "|path=" & partPath
        Exit Sub
    End If

    Dim saveFlagBefore As Boolean
    saveFlagBefore = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swPart.GetSaveFlag)

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    Dim evidence As CRunEvidence
    Set evidence = New CRunEvidence

    Dim graph As CLocationGraph
    Set graph = New CLocationGraph

    Module21_EvidenceSink.LogLine _
        "R23_CATALOG_BEGIN|part=" & partPath & _
        "|configuration=" & configurationName & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath)

    Dim built As Boolean
    built = BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence)

    ' Diagnostics are already streamed to the Immediate window as they are
    ' emitted: CRunEvidence.AddInfo/AddWarning/AddFailure each Debug.Print,
    ' and EmitInfo/EmitWarning/EmitFailure print again under
    ' mEmitDiagnostics. The collections themselves are Private with no
    ' public accessor, so only the tallies are reported here.
    Module21_EvidenceSink.LogLine _
        "R23_CATALOG_EVIDENCE|warnings=" & _
        CStr(evidence.WarningCount) & _
        "|failures=" & CStr(evidence.FailureCount)

    Dim i As Long
    Dim locations As Collection
    Set locations = graph.Locations()

    For i = 1 To locations.Count
        Dim location As CPhysicalHoleLocation
        Set location = locations(i)
        Module21_EvidenceSink.LogLine _
            "R23_CATALOG_LOCATION|index=" & CStr(i) & _
            "|" & location.AggregateDefinitionSummary & _
            "|physicalKey=" & location.PhysicalInstanceKey
    Next i

    Dim familyKeys As Collection
    Set familyKeys = graph.FamilyKeys()

    For i = 1 To familyKeys.Count
        Module21_EvidenceSink.LogLine _
            "R23_CATALOG_FAMILY|key=" & CStr(familyKeys(i)) & _
            "|locations=" & CStr( _
                graph.LocationsForFamily(CStr(familyKeys(i))).Count)
    Next i

    Dim catalogFailures As String
    catalogFailures = VerifyExpectedCatalog(graph, "6,4")

    Dim curveOrderFailures As String
    curveOrderFailures = ProbeCurveReadOrders(graph)
    If Len(curveOrderFailures) > 0 Then
        catalogFailures = AppendFailure( _
            catalogFailures, curveOrderFailures)
    End If

    Dim saveFlagAfter As Boolean
    saveFlagAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swPart.GetSaveFlag)

    Module21_EvidenceSink.LogLine _
        "R23_CATALOG_END|built=" & CStr(built) & _
        "|" & graph.GraphSummary() & _
        "|catalogFailures=" & _
            IIf(Len(catalogFailures) = 0, "None", catalogFailures) & _
        "|modelUnchanged=" & CStr(saveFlagBefore = saveFlagAfter)
    Exit Sub

Failed:
    Module21_EvidenceSink.LogLine _
        "R23_CATALOG_FATAL|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description
End Sub

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub

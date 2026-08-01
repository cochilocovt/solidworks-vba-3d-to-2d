Option Explicit

' R23 Phase 5. Required location schemes and the ordinate transaction.
'
' SAFETY BOUNDARY. Exactly one procedure in this module changes a drawing:
' CreateOrdinateGroup, which refuses unless passed an explicit allowMutation
' argument. R23_ProbeOrdinateScheme never passes it. The probe resolves
' schemes, datums and coordinate buckets, proves every anchor is selectable,
' and reports required coverage without creating anything, so it can be run
' against the manual reference drawing.
'
' Selection state is not document content, but the probe still restores it:
' every path clears selections on exit and the entry point reports the
' selection count before and after.

' swAddOrdinateDims_e, verified against the 2025 table:
'   1 swOrdinate         direction inferred from the selected points
'   2 swVerticalOrdinate
'   3 swHorizontalOrdinate
'   4 swAngularOrdinate
'
' R23 never uses swOrdinate. Letting SOLIDWORKS infer the direction from the
' selection would make the created dimension depend on selection order
' rather than on the scheme, which is the opposite of what R23-500 is for.
Private Const ORD_INFERRED As Long = 1
Public Const ORD_VERTICAL As Long = 2
Public Const ORD_HORIZONTAL As Long = 3
Private Const ORD_ANGULAR As Long = 4

' swCreateOrdDimError_e, verified against the 2025 table. Every member is
' decoded by name in evidence: R23-508 requires the complete result, and a
' bare integer is not a result anyone can act on.
Private Const ORD_ERR_UNDEFINED As Long = -1
Private Const ORD_ERR_SUCCESS As Long = 0
Private Const ORD_ERR_ORD_FAILURE As Long = 1
Private Const ORD_ERR_NO_INTERNAL_DIMS As Long = 2
Private Const ORD_ERR_BAD_SEL As Long = 3
Private Const ORD_ERR_NEED_MODEL_LOADED As Long = 4
Private Const ORD_ERR_SAME_PART_ONLY As Long = 5
Private Const ORD_ERR_EXTRA_SELECTION As Long = 6
Private Const ORD_ERR_GEN_FAILURE As Long = 7
Private Const ORD_ERR_DUP_IN_GROUP As Long = 8
Private Const ORD_ERR_BAD_DIR As Long = 9

' swDrawingViewTypes_e.
Private Const VIEW_TYPE_SECTION As Long = 2
Private Const VIEW_TYPE_DETAIL As Long = 3
Private Const VIEW_TYPE_PROJECTED As Long = 4
Private Const VIEW_TYPE_AUXILIARY As Long = 5
Private Const VIEW_TYPE_NAMED As Long = 7

' R23-505. Two page coordinates are the same ordinate coordinate when they
' agree to this quantum. It is deliberately the same order as the Phase 3
' coincidence quantum that proved two coaxial holes share one drawing
' entity, because that is the physical situation this bucketing exists to
' handle.
Public Const ORDINATE_COORDINATE_QUANTUM_M As Double = 0.000001

' R23-501 and R23-502. Datum policy identifiers. The policy is part of the
' scheme key, so a change of datum produces a different scheme rather than
' silently re-basing an existing one.
Public Const DATUM_POLICY_X As String = "CentreBoreProjectedCentre"
Public Const DATUM_POLICY_Y As String = "LowestProjectedGeometry"

' Datum kinds, which say what the datum actually IS rather than which
' policy asked for it. ProjectionDerived means the datum is one of the
' part's own hole projections; OutlineDerived would mean a profile edge.
Private Const DATUM_KIND_PROJECTION As String = "ProjectionDerived"
Private Const DATUM_KIND_UNRESOLVED As String = "Unresolved"

' R23-507. A location enters the small-hole ledger when it belongs to a
' repeated family, and the profile/reference ledger when it stands alone.
' This is measured from the graph rather than from a radius threshold: a
' magic millimetre value would silently reclassify a different part, and
' P-0251's stepped bore is excluded because it is a singleton, not because
' it is large.
Private Const MIN_FAMILY_SIZE_FOR_SMALL_HOLE As Long = 2

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

' R23-500. Builds every ordinate scheme the graph implies, for one view.
'
' Returns a Collection of COrdinateScheme, normally two per eligible view -
' one horizontal, one vertical - per machining face present in that view.
Public Function BuildSchemesForView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Collection

    On Error GoTo Failed

    Dim schemes As Collection
    Set schemes = New Collection
    Set BuildSchemesForView = schemes

    If swView Is Nothing Then Exit Function
    If graph Is Nothing Then Exit Function

    Dim viewName As String
    viewName = SafeViewName(swView)

    Dim viewType As Long
    viewType = SafeViewType(swView)

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(viewName)

    If projections.Count = 0 Then
        EmitInfo evidence, "ORDINATE_SCHEME_SKIPPED|view=" & viewName & _
            "|reason=NoProjectionsInView"
        Exit Function
    End If

    Dim directions(1 To 2) As Long
    directions(1) = ORD_HORIZONTAL
    directions(2) = ORD_VERTICAL

    Dim d As Long
    For d = 1 To 2
        Dim faceKeys As Collection
        Set faceKeys = DistinctMachiningFaces(projections)

        Dim f As Long
        For f = 1 To faceKeys.Count
            Dim scheme As COrdinateScheme
            Set scheme = New COrdinateScheme

            scheme.ViewName = viewName
            scheme.ViewType = viewType
            scheme.ViewRole = ClassifyViewRole(swView, graph)
            scheme.MachiningFaceKey = CStr(faceKeys(f))
            scheme.Direction = directions(d)
            scheme.DirectionName = DirectionName(directions(d))
            scheme.DatumPolicyId = DatumPolicyIdFor(directions(d))
            Set scheme.DrawingView = swView

            PopulateSchemeLedger swApp, graph, scheme, projections, evidence

            If scheme.BucketCount() > 0 Then
                ResolveSchemeDatum swDraw, scheme, evidence
                schemes.Add scheme
            End If
        Next f
    Next d

    Exit Function

Failed:
    EmitFailure evidence, "ORDINATE_SCHEME_ERROR|error=" & _
        CStr(Err.Number) & "|description=" & Err.Description
End Function

' Distinct machining faces present among a view's projections, taken from
' each physical location's sign-normalized axis. A machining face is a
' property of the hole, not of the view, so it survives reorientation.
Private Function DistinctMachiningFaces( _
    ByRef projections As Collection) As Collection

    Dim result As Collection
    Set result = New Collection
    Set DistinctMachiningFaces = result

    Dim i As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)

        Dim faceKey As String
        faceKey = MachiningFaceKey(projection)
        If Len(faceKey) = 0 Then GoTo ContinueProjection

        Dim found As Boolean
        found = False

        Dim j As Long
        For j = 1 To result.Count
            If StrComp(CStr(result(j)), faceKey, vbBinaryCompare) = 0 Then
                found = True
                Exit For
            End If
        Next j

        If Not found Then result.Add faceKey

ContinueProjection:
    Next i
End Function

' The machining face a hole is produced from, keyed by its sign-normalized
' axis. Two holes drilled from opposite faces have opposite raw axes and the
' same normalized axis, so the axial interval sign is carried too.
Public Function MachiningFaceKey( _
    ByRef projection As CViewHoleProjection) As String

    On Error GoTo Failed

    If projection Is Nothing Then Exit Function
    If projection.PhysicalLocation Is Nothing Then Exit Function

    Dim location As CPhysicalHoleLocation
    Set location = projection.PhysicalLocation

    MachiningFaceKey = "axis=" & _
        Module11_GeometryIdentity.CanonicalAxisToken( _
            location.AxisX, location.AxisY, location.AxisZ)
    Exit Function

Failed:
    MachiningFaceKey = vbNullString
End Function

' R23-403 and R23-404 already measure what a view is FOR. The role reuses
' those measurements rather than reading a view name, so a renamed view is
' still classified correctly.
Public Function ClassifyViewRole( _
    ByRef swView As SldWorks.View, _
    ByRef graph As CLocationGraph) As String

    On Error GoTo Failed

    If swView Is Nothing Then
        ClassifyViewRole = "Unknown"
        Exit Function
    End If

    If SafeViewType(swView) = VIEW_TYPE_SECTION Then
        ClassifyViewRole = "Section"
        Exit Function
    End If

    If Module14_AnnotationImport.IsDeferredCreationView( _
        graph, swView) Then

        ClassifyViewRole = "DeferredNoNormalAxis"
        Exit Function
    End If

    Dim eligibilityReason As String
    If Module14_AnnotationImport.IsOrdinateEligibleView( _
        graph, swView, eligibilityReason) Then

        ClassifyViewRole = "OrdinateBearing"
        Exit Function
    End If

    ClassifyViewRole = "NonOrdinate"
    Exit Function

Failed:
    ClassifyViewRole = "Unknown"
End Function

Public Function DirectionName(ByVal direction As Long) As String
    Select Case direction
        Case ORD_HORIZONTAL
            DirectionName = "Horizontal"
        Case ORD_VERTICAL
            DirectionName = "Vertical"
        Case ORD_ANGULAR
            DirectionName = "Angular"
        Case ORD_INFERRED
            DirectionName = "InferredNotUsedByR23"
        Case Else
            DirectionName = "Unknown:" & CStr(direction)
    End Select
End Function

Public Function DatumPolicyIdFor(ByVal direction As Long) As String
    If direction = ORD_HORIZONTAL Then
        DatumPolicyIdFor = DATUM_POLICY_X
    ElseIf direction = ORD_VERTICAL Then
        DatumPolicyIdFor = DATUM_POLICY_Y
    Else
        DatumPolicyIdFor = "Unsupported"
    End If
End Function

' R23-503, R23-504, R23-505 and R23-507. Fills one scheme's coordinate
' buckets from the view's projections.
Private Sub PopulateSchemeLedger( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef graph As CLocationGraph, _
    ByRef scheme As COrdinateScheme, _
    ByRef projections As Collection, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim i As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)

        If projection Is Nothing Then GoTo ContinueProjection
        If projection.PhysicalLocation Is Nothing Then GoTo ContinueProjection

        If StrComp(MachiningFaceKey(projection), _
            scheme.MachiningFaceKey, vbBinaryCompare) <> 0 Then

            GoTo ContinueProjection
        End If

        ' R23-507. Singleton-family locations are profile/reference content
        ' and are held apart from the small-hole ledger entirely.
        If Not IsSmallHoleLocation(graph, projection.PhysicalLocation) Then
            scheme.ProfileProjections.Add projection
            GoTo ContinueProjection
        End If

        ' Only projections Phase 3 accepted can carry an ordinate. An
        ' unaccepted projection is still credited to nothing and reported,
        ' never silently dropped.
        If Not projection.Accepted Then
            EmitInfo evidence, "ORDINATE_LEDGER_EXCLUDED|" & _
                scheme.SchemeKey() & _
                "|physical=" & projection.PhysicalInstanceKey & _
                "|reason=" & projection.RejectionReason
            GoTo ContinueProjection
        End If

        Dim coordinateKey As String
        coordinateKey = CoordinateKeyFor(projection, scheme.Direction)

        Dim bucket As COrdinateBucket
        Set bucket = scheme.BucketByKey(coordinateKey)

        If bucket Is Nothing Then
            Set bucket = New COrdinateBucket
            bucket.CoordinateKey = coordinateKey
            bucket.Direction = scheme.Direction
            bucket.DirectionName = scheme.DirectionName
            bucket.PageX = projection.PageX
            bucket.PageY = projection.PageY
            bucket.AnchorSourceKey = projection.PhysicalInstanceKey
            bucket.AnchorProof = projection.AnchorProofSource

            If projection.HasSelectableAnchor() Then
                Set bucket.Anchor = projection.PrimaryAnchor
            End If

            scheme.AddBucket bucket
        End If

        ' R23-505. Every physical location the coordinate represents is
        ' credited, including ones whose projection coincides with another.
        bucket.CreditLocation projection.PhysicalInstanceKey, projection

        If Len(projection.CoincidentWithAnchoredKey) > 0 Then
            CreditCoincidentPartner _
                graph, scheme, bucket, projection, evidence
        End If

ContinueProjection:
    Next i

    Exit Sub

Failed:
    EmitFailure evidence, "ORDINATE_LEDGER_ERROR|" & scheme.SchemeKey() & _
        "|error=" & CStr(Err.Number)
End Sub

' Phase 3 proved live that two coaxial holes viewed along their shared axis
' produce ONE drawing entity. The partner location has no anchor of its own
' and would otherwise vanish from coverage, so it is credited to the bucket
' its twin occupies.
Private Sub CreditCoincidentPartner( _
    ByRef graph As CLocationGraph, _
    ByRef scheme As COrdinateScheme, _
    ByRef bucket As COrdinateBucket, _
    ByRef projection As CViewHoleProjection, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim partnerKey As String
    partnerKey = projection.CoincidentWithAnchoredKey
    If Len(Trim$(partnerKey)) = 0 Then Exit Sub

    If bucket.AlreadyCredited(partnerKey) Then Exit Sub

    bucket.CreditLocation partnerKey, Nothing

    EmitInfo evidence, "ORDINATE_COINCIDENT_CREDIT|" & _
        scheme.SchemeKey() & _
        "|coordinateKey=" & bucket.CoordinateKey & _
        "|anchoredBy=" & projection.PhysicalInstanceKey & _
        "|alsoCredits=" & partnerKey & _
        "|reason=OneDrawingEntityForTwoCoaxialHoles"
    Exit Sub

Failed:
    EmitWarning evidence, "ORDINATE_COINCIDENT_CREDIT_ERROR|error=" & _
        CStr(Err.Number)
End Sub

' R23-507. Small-hole membership is family size, measured from the graph.
Public Function IsSmallHoleLocation( _
    ByRef graph As CLocationGraph, _
    ByRef location As CPhysicalHoleLocation) As Boolean

    On Error GoTo Failed

    If graph Is Nothing Then Exit Function
    If location Is Nothing Then Exit Function
    If Len(Trim$(location.SemanticFamilyKey)) = 0 Then Exit Function

    Dim familyMembers As Collection
    Set familyMembers = graph.LocationsForFamily(location.SemanticFamilyKey)

    IsSmallHoleLocation = _
        (familyMembers.Count >= MIN_FAMILY_SIZE_FOR_SMALL_HOLE)
    Exit Function

Failed:
    IsSmallHoleLocation = False
End Function

' R23-505. The bucket key is the quantized page coordinate in the direction
' being ordinated. A horizontal ordinate measures X, so two holes in the
' same column share a bucket regardless of their Y.
Public Function CoordinateKeyFor( _
    ByRef projection As CViewHoleProjection, _
    ByVal direction As Long) As String

    If projection Is Nothing Then Exit Function

    If direction = ORD_HORIZONTAL Then
        CoordinateKeyFor = "x=" & QuantizeCoordinate(projection.PageX)
    ElseIf direction = ORD_VERTICAL Then
        CoordinateKeyFor = "y=" & QuantizeCoordinate(projection.PageY)
    Else
        CoordinateKeyFor = "unsupportedDirection=" & CStr(direction)
    End If
End Function

Public Function QuantizeCoordinate(ByVal value As Double) As String
    Dim quantized As Double
    quantized = _
        Int(value / ORDINATE_COORDINATE_QUANTUM_M + 0.5) * _
        ORDINATE_COORDINATE_QUANTUM_M

    QuantizeCoordinate = Format$(quantized, "0.000000000")
End Function

' R23-501 and R23-502. Chooses the datum bucket for the scheme's direction
' and proves it can actually be selected in this view.
'
' Horizontal: the centre bore's projected centre - the largest-radius
' location present in the view whose axis Phase 3 measured as normal to it.
' Vertical: the lowest page coordinate in the ledger.
'
' Both datums are PROJECTION-DERIVED, recorded as such. A profile-outline
' datum is a different kind and is not claimed here.
Private Sub ResolveSchemeDatum( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef scheme As COrdinateScheme, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    scheme.DatumKind = DATUM_KIND_UNRESOLVED
    scheme.DatumProof = "datum=NotResolved"

    Dim chosen As COrdinateBucket
    Set chosen = Nothing

    If scheme.Direction = ORD_HORIZONTAL Then
        Set chosen = ResolveCentreBoreDatum(scheme)
        If chosen Is Nothing Then Set chosen = LowestBucket(scheme)
    Else
        Set chosen = LowestBucket(scheme)
    End If

    If chosen Is Nothing Then
        EmitWarning evidence, "ORDINATE_DATUM|" & scheme.SchemeKey() & _
            "|resolved=False|reason=NoBucketAvailable"
        Exit Sub
    End If

    Set scheme.DatumBucket = chosen
    chosen.IsDatum = True
    scheme.DatumKind = DATUM_KIND_PROJECTION

    Dim selectionProof As String
    Dim selectable As Boolean
    selectable = ProveBucketSelectable( _
        swDraw, scheme, chosen, selectionProof)

    scheme.DatumResolved = selectable
    scheme.DatumProof = "datumKind=" & scheme.DatumKind & _
        "|coordinateKey=" & chosen.CoordinateKey & _
        "|anchorSource=" & chosen.AnchorSourceKey & _
        "|" & selectionProof

    EmitInfo evidence, "ORDINATE_DATUM|" & scheme.SchemeKey() & _
        "|resolved=" & CStr(selectable) & "|" & scheme.DatumProof
    Exit Sub

Failed:
    scheme.DatumResolved = False
    scheme.DatumProof = "datum=Error:" & CStr(Err.Number)
    EmitFailure evidence, "ORDINATE_DATUM_ERROR|" & scheme.SchemeKey() & _
        "|error=" & CStr(Err.Number)
End Sub

' The centre bore for X: the scheme's own machining face has already been
' matched, so this looks among the view's profile entries - where the
' singleton stepped bore lives - for the largest radius with a usable
' anchor, and returns a bucket wrapping it.
Private Function ResolveCentreBoreDatum( _
    ByRef scheme As COrdinateScheme) As COrdinateBucket

    On Error GoTo Failed

    Dim best As CViewHoleProjection
    Set best = Nothing

    Dim i As Long
    For i = 1 To scheme.ProfileProjections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = scheme.ProfileProjections(i)

        If candidate Is Nothing Then GoTo ContinueCandidate
        If Not candidate.Accepted Then GoTo ContinueCandidate
        If Not candidate.HasSelectableAnchor() Then GoTo ContinueCandidate
        If candidate.PhysicalLocation Is Nothing Then GoTo ContinueCandidate

        If best Is Nothing Then
            Set best = candidate
        ElseIf candidate.PhysicalLocation.MaximumRadiusM > _
            best.PhysicalLocation.MaximumRadiusM Then

            Set best = candidate
        End If

ContinueCandidate:
    Next i

    If best Is Nothing Then Exit Function

    Dim bucket As COrdinateBucket
    Set bucket = New COrdinateBucket
    bucket.CoordinateKey = CoordinateKeyFor(best, scheme.Direction)
    bucket.Direction = scheme.Direction
    bucket.DirectionName = scheme.DirectionName
    bucket.PageX = best.PageX
    bucket.PageY = best.PageY
    bucket.AnchorSourceKey = best.PhysicalInstanceKey
    bucket.AnchorProof = best.AnchorProofSource
    Set bucket.Anchor = best.PrimaryAnchor
    bucket.CreditLocation best.PhysicalInstanceKey, best

    Set ResolveCentreBoreDatum = bucket
    Exit Function

Failed:
    Set ResolveCentreBoreDatum = Nothing
End Function

' The lowest page coordinate in the scheme's direction, among buckets that
' actually carry an anchor.
Private Function LowestBucket( _
    ByRef scheme As COrdinateScheme) As COrdinateBucket

    On Error GoTo Failed

    Dim best As COrdinateBucket
    Set best = Nothing

    Dim i As Long
    For i = 1 To scheme.Buckets.Count
        Dim candidate As COrdinateBucket
        Set candidate = scheme.Buckets(i)

        If Not candidate.HasSelectableAnchor() Then GoTo ContinueBucket

        If best Is Nothing Then
            Set best = candidate
        ElseIf BucketOrdinateValue(candidate, scheme.Direction) < _
            BucketOrdinateValue(best, scheme.Direction) Then

            Set best = candidate
        End If

ContinueBucket:
    Next i

    Set LowestBucket = best
    Exit Function

Failed:
    Set LowestBucket = Nothing
End Function

Private Function BucketOrdinateValue( _
    ByRef bucket As COrdinateBucket, _
    ByVal direction As Long) As Double

    If direction = ORD_HORIZONTAL Then
        BucketOrdinateValue = bucket.PageX
    Else
        BucketOrdinateValue = bucket.PageY
    End If
End Function

' Selects the bucket's anchor, proves the selection landed in the intended
' view, then clears. Ownership is proved AFTER the fact through
' ISelectionMgr.GetSelectedObjectsDrawingView2 rather than assumed, because
' ISelectData.View is documented get/set but raises runtime error 91 on this
' build, which Module2_DrawingPipeline already records.
Public Function ProveBucketSelectable( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef scheme As COrdinateScheme, _
    ByRef bucket As COrdinateBucket, _
    ByRef selectionProof As String) As Boolean

    On Error GoTo Failed

    selectionProof = "selection=NotAttempted"

    If swDraw Is Nothing Then Exit Function
    If bucket Is Nothing Then Exit Function

    If Not bucket.HasSelectableAnchor() Then
        selectionProof = "selection=Reject|reason=NoAnchor"
        Exit Function
    End If

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDraw.SelectionManager
    If selectionMgr Is Nothing Then
        selectionProof = "selection=Reject|reason=NoSelectionManager"
        Exit Function
    End If

    swDraw.ClearSelection2 True

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionMgr.CreateSelectData

    Dim viewBinding As String
    viewBinding = TryBindSelectDataView(selectData, scheme.DrawingView)

    Dim selected As Boolean
    selected = Module11_GeometryIdentity.NormalizeSwBoolean( _
        bucket.Anchor.Select4(False, selectData))

    Dim selectedCount As Long
    selectedCount = selectionMgr.GetSelectedObjectCount2(-1)

    Dim owningViewName As String
    owningViewName = "Unproven"

    If selected And selectedCount = 1 Then
        Dim owningView As SldWorks.View
        Set owningView = Nothing
        On Error Resume Next
        Set owningView = selectionMgr.GetSelectedObjectsDrawingView2(1, -1)
        On Error GoTo Failed

        If Not owningView Is Nothing Then
            owningViewName = SafeViewName(owningView)
        End If
    End If

    Dim ownershipProven As Boolean
    ownershipProven = (StrComp( _
        Module11_GeometryIdentity.IdentityToken(owningViewName), _
        Module11_GeometryIdentity.IdentityToken(scheme.ViewName), _
        vbBinaryCompare) = 0)

    selectionProof = "selection=" & CStr(selected) & _
        "|viewBinding=" & viewBinding & _
        "|selectedCount=" & CStr(selectedCount) & _
        "|owningView=" & owningViewName & _
        "|ownershipProven=" & CStr(ownershipProven) & _
        "|ownershipSource=ISelectionMgr.GetSelectedObjectsDrawingView2"

    swDraw.ClearSelection2 True

    ProveBucketSelectable = _
        (selected And selectedCount = 1 And ownershipProven)
    Exit Function

Failed:
    selectionProof = "selection=Reject|reason=Error:" & CStr(Err.Number)
    On Error Resume Next
    swDraw.ClearSelection2 True
    ProveBucketSelectable = False
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

' R23-506 and R23-508. MUTATES THE DRAWING.
'
' Refuses unless allowMutation is True. Performs the ordinate transaction in
' the documented order: activate and verify the view, bind ISelectData.View,
' select the datum FIRST, append each remaining bucket explicitly in
' deterministic coordinate order, verify the selection count at every step,
' call AddOrdinateDimension, decode the complete result, call SetPickMode to
' close the group, clear selections on every exit, then read back what was
' actually created.
'
' Returns the number of display dimensions read back, or -1 on refusal or
' failure. A positive return is never assumed from the API result alone:
' R23-508 requires the read-back.
Public Function CreateOrdinateGroup( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef scheme As COrdinateScheme, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    CreateOrdinateGroup = -1

    If Not allowMutation Then
        scheme.TransactionProof = "refused=MutationNotAuthorized"
        EmitWarning evidence, "ORDINATE_CREATE_REFUSED|" & _
            scheme.SchemeKey() & "|reason=MutationNotAuthorized"
        Exit Function
    End If

    If swDraw Is Nothing Then Exit Function
    If scheme Is Nothing Then Exit Function

    If Not scheme.DatumResolved Then
        scheme.TransactionProof = "refused=DatumNotProven"
        EmitFailure evidence, "ORDINATE_CREATE_REFUSED|" & _
            scheme.SchemeKey() & "|reason=DatumNotProven"
        Exit Function
    End If

    scheme.TransactionAttempted = True

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim activated As Boolean
    activated = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDrawing.ActivateView(scheme.ViewName))

    If Not activated Then
        scheme.TransactionProof = "refused=ViewActivationFailed"
        EmitFailure evidence, "ORDINATE_CREATE_REFUSED|" & _
            scheme.SchemeKey() & "|reason=ViewActivationFailed"
        Exit Function
    End If

    swDraw.ClearSelection2 True

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDraw.SelectionManager

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionMgr.CreateSelectData

    Dim viewBinding As String
    viewBinding = TryBindSelectDataView(selectData, scheme.DrawingView)

    ' Taken before anything is created, so the read-back can name what is
    ' new by identity rather than inferring it from a count.
    Dim before As Collection
    Set before = SnapshotViewDimensions(scheme.DrawingView)

    ' Datum first. The Remarks for AddOrdinateDimension are explicit that
    ' the base entity acting as the datum must be selected before anything
    ' else, and everything downstream is measured from it.
    Dim expectedCount As Long
    expectedCount = 0

    If Not AppendBucketSelection( _
        swDraw, selectionMgr, selectData, scheme.DatumBucket, _
        expectedCount, evidence, scheme) Then

        scheme.TransactionProof = "refused=DatumSelectionFailed"
        swDraw.ClearSelection2 True
        Exit Function
    End If

    Dim ordered As Collection
    Set ordered = BucketsInDeterministicOrder(scheme)

    Dim i As Long
    For i = 1 To ordered.Count
        Dim bucket As COrdinateBucket
        Set bucket = ordered(i)

        If bucket Is scheme.DatumBucket Then GoTo ContinueBucket

        If Not AppendBucketSelection( _
            swDraw, selectionMgr, selectData, bucket, _
            expectedCount, evidence, scheme) Then

            scheme.TransactionProof = "refused=AppendSelectionFailed" & _
                "|coordinateKey=" & bucket.CoordinateKey
            swDraw.ClearSelection2 True
            Exit Function
        End If

ContinueBucket:
    Next i

    Dim placeX As Double
    Dim placeY As Double
    placeX = scheme.DatumBucket.PageX
    placeY = scheme.DatumBucket.PageY

    Dim result As Long
    result = swDraw.Extension.AddOrdinateDimension( _
        scheme.Direction, placeX, placeY, 0#)

    ' Close the group. The Remarks are explicit that selections made after
    ' the call keep appending to it until SetPickMode is called, so this
    ' runs whatever the result was.
    swDraw.SetPickMode
    swDraw.ClearSelection2 True

    Dim createdCount As Long
    createdCount = ReadBackCreatedOrdinates( _
        swApp, scheme, before, evidence)

    scheme.CreatedDimensionCount = createdCount
    scheme.TransactionProof = "viewActivated=True" & _
        "|viewBinding=" & viewBinding & _
        "|selectionsAppended=" & CStr(expectedCount) & _
        "|expectedSelections=" & CStr(scheme.AnchoredBucketCount()) & _
        "|apiResult=" & CStr(result) & _
        "|apiResultName=" & DecodeOrdinateError(result) & _
        "|createdReadBack=" & CStr(createdCount) & _
        "|pickModeRestored=True"

    EmitInfo evidence, "ORDINATE_CREATE|" & scheme.SchemeKey() & _
        "|" & scheme.TransactionProof

    If result <> ORD_ERR_SUCCESS Then
        EmitFailure evidence, "ORDINATE_CREATE_FAILED|" & _
            scheme.SchemeKey() & _
            "|apiResult=" & CStr(result) & _
            "|apiResultName=" & DecodeOrdinateError(result)
        CreateOrdinateGroup = -1
        Exit Function
    End If

    CreateOrdinateGroup = createdCount
    Exit Function

Failed:
    scheme.TransactionProof = "error=" & CStr(Err.Number)
    EmitFailure evidence, "ORDINATE_CREATE_ERROR|" & scheme.SchemeKey() & _
        "|error=" & CStr(Err.Number) & "|description=" & Err.Description

    On Error Resume Next
    swDraw.SetPickMode
    swDraw.ClearSelection2 True
    CreateOrdinateGroup = -1
End Function

' Appends one bucket's anchor to the running selection and verifies the
' count advanced by exactly one. R23-508 requires the count to be checked at
' every step, not only at the end: a silently ignored selection would
' otherwise produce a group that is short by one hole and look successful.
Private Function AppendBucketSelection( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef selectionMgr As SldWorks.SelectionMgr, _
    ByRef selectData As SldWorks.SelectData, _
    ByRef bucket As COrdinateBucket, _
    ByRef expectedCount As Long, _
    ByRef evidence As CRunEvidence, _
    ByRef scheme As COrdinateScheme) As Boolean

    On Error GoTo Failed

    If bucket Is Nothing Then Exit Function
    If Not bucket.HasSelectableAnchor() Then
        EmitFailure evidence, "ORDINATE_APPEND_FAILED|" & _
            scheme.SchemeKey() & _
            "|coordinateKey=" & bucket.CoordinateKey & _
            "|reason=NoAnchor"
        Exit Function
    End If

    Dim appended As Boolean
    appended = Module11_GeometryIdentity.NormalizeSwBoolean( _
        bucket.Anchor.Select4(True, selectData))

    Dim actualCount As Long
    actualCount = selectionMgr.GetSelectedObjectCount2(-1)

    If Not appended Or actualCount <> expectedCount + 1 Then
        EmitFailure evidence, "ORDINATE_APPEND_FAILED|" & _
            scheme.SchemeKey() & _
            "|coordinateKey=" & bucket.CoordinateKey & _
            "|appended=" & CStr(appended) & _
            "|expectedCount=" & CStr(expectedCount + 1) & _
            "|actualCount=" & CStr(actualCount)
        Exit Function
    End If

    expectedCount = actualCount
    AppendBucketSelection = True
    Exit Function

Failed:
    EmitFailure evidence, "ORDINATE_APPEND_ERROR|" & scheme.SchemeKey() & _
        "|error=" & CStr(Err.Number)
    AppendBucketSelection = False
End Function

' Deterministic append order: ascending ordinate coordinate. Order must not
' depend on graph traversal, or two runs over the same drawing could produce
' differently ordered groups from identical geometry.
Public Function BucketsInDeterministicOrder( _
    ByRef scheme As COrdinateScheme) As Collection

    Dim ordered As Collection
    Set ordered = New Collection
    Set BucketsInDeterministicOrder = ordered

    Dim taken() As Boolean
    Dim total As Long
    total = scheme.Buckets.Count
    If total = 0 Then Exit Function

    ReDim taken(1 To total)

    Dim placed As Long
    For placed = 1 To total
        Dim bestIndex As Long
        bestIndex = 0

        Dim i As Long
        For i = 1 To total
            If taken(i) Then GoTo ContinueCandidate

            If bestIndex = 0 Then
                bestIndex = i
            ElseIf BucketOrdinateValue( _
                scheme.Buckets(i), scheme.Direction) < _
                BucketOrdinateValue( _
                    scheme.Buckets(bestIndex), scheme.Direction) Then

                bestIndex = i
            End If

ContinueCandidate:
        Next i

        If bestIndex = 0 Then Exit For
        taken(bestIndex) = True
        ordered.Add scheme.Buckets(bestIndex)
    Next placed
End Function

' Snapshot of a view's display dimensions.
'
' IView.GetDisplayDimensions is scoped to THIS view and returns the whole
' array at once. IView.GetFirstDisplayDimension5 is not used: it is marked
' obsolete in favour of GetFirstDisplayDimension6, and its own Remarks say
' the GetNext5 walk covers the drawing SHEET, so a read-back built on it
' would count dimensions belonging to other views as this scheme's output.
Private Function SnapshotViewDimensions( _
    ByRef swView As SldWorks.View) As Collection

    On Error GoTo Failed

    Dim snapshot As Collection
    Set snapshot = New Collection
    Set SnapshotViewDimensions = snapshot

    If swView Is Nothing Then Exit Function

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If VariantItemCount(dimensions) = 0 Then Exit Function

    Dim i As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim candidate As Object
        Set candidate = Nothing
        On Error Resume Next
        Set candidate = dimensions(i)
        On Error GoTo Failed

        If Not candidate Is Nothing Then snapshot.Add candidate
    Next i

    Exit Function

Failed:
    Set SnapshotViewDimensions = snapshot
End Function

' R23-508. Reports the display dimensions that exist in the view AFTER the
' transaction and were not in the before-snapshot, matched by COM identity.
'
' A count difference alone would not do: an unrelated dimension appearing
' for any other reason would inflate it, and the whole point of the
' read-back is that the API's own result is a claim rather than evidence.
Private Function ReadBackCreatedOrdinates( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef scheme As COrdinateScheme, _
    ByRef before As Collection, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    Dim found As Long

    Dim after As Collection
    Set after = SnapshotViewDimensions(scheme.DrawingView)

    Dim i As Long
    For i = 1 To after.Count
        Dim candidate As Object
        Set candidate = after(i)

        If CollectionContainsObject(swApp, before, candidate) Then
            GoTo ContinueDimension
        End If

        Dim dimensionType As Long
        dimensionType = -1
        On Error Resume Next
        dimensionType = candidate.Type2
        On Error GoTo Failed

        found = found + 1

        EmitInfo evidence, "ORDINATE_READBACK|" & scheme.SchemeKey() & _
            "|index=" & CStr(found) & _
            "|type2=" & CStr(dimensionType) & _
            "|isOrdinateType=" & _
            CStr(IsOrdinateDimensionType(dimensionType)) & _
            "|source=IView.GetDisplayDimensions" & _
            "|identity=ISldWorks.IsSame"

ContinueDimension:
    Next i

    EmitInfo evidence, "ORDINATE_READBACK_SUMMARY|" & _
        scheme.SchemeKey() & _
        "|before=" & CStr(before.Count) & _
        "|after=" & CStr(after.Count) & _
        "|new=" & CStr(found)

    ReadBackCreatedOrdinates = found
    Exit Function

Failed:
    EmitWarning evidence, "ORDINATE_READBACK_ERROR|" & _
        scheme.SchemeKey() & "|error=" & CStr(Err.Number)
    ReadBackCreatedOrdinates = found
End Function

' ISldWorks.IsSame returns swObjectEquality, not a Boolean: 0 NotSame,
' 1 Same, 2 Unsupported. Only 1 is a match.
Private Function CollectionContainsObject( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef items As Collection, _
    ByRef candidate As Object) As Boolean

    On Error GoTo Failed

    If items Is Nothing Then Exit Function
    If candidate Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To items.Count
        Dim equality As Long
        equality = -1
        On Error Resume Next
        equality = CLng(swApp.IsSame(items(i), candidate))
        On Error GoTo Failed

        If equality = 1 Then
            CollectionContainsObject = True
            Exit Function
        End If
    Next i

    Exit Function

Failed:
    CollectionContainsObject = False
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

' Ordinate display-dimension types as proved live in Phase 0: created
' ordinates carried Type2 of 1 and 7, and the reference drawing's own
' ordinates read as 1, 7 and 8.
Public Function IsOrdinateDimensionType( _
    ByVal dimensionType As Long) As Boolean

    IsOrdinateDimensionType = _
        (dimensionType = 1 Or dimensionType = 7 Or dimensionType = 8)
End Function

' R23-508. Every swCreateOrdDimError_e member decoded by name.
Public Function DecodeOrdinateError(ByVal result As Long) As String
    Select Case result
        Case ORD_ERR_SUCCESS
            DecodeOrdinateError = "swCreateOrdDimErr_Success"
        Case ORD_ERR_ORD_FAILURE
            DecodeOrdinateError = "swCreateOrdDimErr_OrdFailure"
        Case ORD_ERR_NO_INTERNAL_DIMS
            DecodeOrdinateError = "swCreateOrdDimErr_GenNoInternalDims"
        Case ORD_ERR_BAD_SEL
            DecodeOrdinateError = "swCreateOrdDimErr_GenBadSel"
        Case ORD_ERR_NEED_MODEL_LOADED
            DecodeOrdinateError = "swCreateOrdDimErr_GenNeedModelLoaded"
        Case ORD_ERR_SAME_PART_ONLY
            DecodeOrdinateError = "swCreateOrdDimErr_GenSamePartOnly"
        Case ORD_ERR_EXTRA_SELECTION
            DecodeOrdinateError = "swCreateOrdDimErr_GenExtraSelection"
        Case ORD_ERR_GEN_FAILURE
            DecodeOrdinateError = "swCreateOrdDimErr_GenFailure"
        Case ORD_ERR_DUP_IN_GROUP
            DecodeOrdinateError = "swCreateOrdDimErr_OrdDupInGroup"
        Case ORD_ERR_BAD_DIR
            DecodeOrdinateError = "swCreateOrdDimErr_OrdBadDir"
        Case ORD_ERR_UNDEFINED
            DecodeOrdinateError = "swCreateOrdDimErr_Undefined"
        Case Else
            DecodeOrdinateError = "Unmapped:" & CStr(result)
    End Select
End Function

' R23-509. Reports what is missing rather than a bare count.
'
' The carried Phase 3 finding is enforced here: required coverage is counted
' per DISTINCT PAGE POSITION per view, because coaxial holes collapse to one
' drawing entity and demanding one dimension per physical location would be
' unsatisfiable by construction. Locations are still counted, but as
' CREDITS against those positions.
Public Function VerifyDirectionalCoverage( _
    ByRef schemes As Collection, _
    ByVal expectedSmallHoleLocations As Long) As String

    On Error GoTo Failed

    Dim failures As String
    Dim creditedKeys As Collection
    Set creditedKeys = New Collection

    Dim horizontalSchemes As Long
    Dim verticalSchemes As Long

    Dim i As Long
    For i = 1 To schemes.Count
        Dim scheme As COrdinateScheme
        Set scheme = schemes(i)

        If scheme.Direction = ORD_HORIZONTAL Then
            horizontalSchemes = horizontalSchemes + 1
        ElseIf scheme.Direction = ORD_VERTICAL Then
            verticalSchemes = verticalSchemes + 1
        End If

        If Not scheme.DatumResolved Then
            failures = AppendFailure(failures, "DatumUnproven:" & _
                scheme.SchemeKey())
        End If

        If scheme.AnchoredBucketCount() <> scheme.BucketCount() Then
            failures = AppendFailure(failures, "UnanchoredBucket:" & _
                scheme.SchemeKey())
        End If

        Dim b As Long
        For b = 1 To scheme.Buckets.Count
            Dim bucket As COrdinateBucket
            Set bucket = scheme.Buckets(b)

            Dim k As Long
            For k = 1 To bucket.RepresentedLocationKeys.Count
                AddDistinctKey creditedKeys, _
                    CStr(bucket.RepresentedLocationKeys(k))
            Next k
        Next b
    Next i

    If horizontalSchemes = 0 Then
        failures = AppendFailure(failures, "NoHorizontalScheme")
    End If

    If verticalSchemes = 0 Then
        failures = AppendFailure(failures, "NoVerticalScheme")
    End If

    If expectedSmallHoleLocations > 0 And _
        creditedKeys.Count <> expectedSmallHoleLocations Then

        failures = AppendFailure(failures, _
            "SmallHoleLocationCredit:expected=" & _
            CStr(expectedSmallHoleLocations) & _
            ",credited=" & CStr(creditedKeys.Count))
    End If

    If Len(failures) = 0 Then failures = "None"

    VerifyDirectionalCoverage = "schemes=" & CStr(schemes.Count) & _
        "|horizontalSchemes=" & CStr(horizontalSchemes) & _
        "|verticalSchemes=" & CStr(verticalSchemes) & _
        "|creditedLocations=" & CStr(creditedKeys.Count) & _
        "|expectedLocations=" & CStr(expectedSmallHoleLocations) & _
        "|coverageFailures=" & failures
    Exit Function

Failed:
    VerifyDirectionalCoverage = "coverageFailures=Error:" & CStr(Err.Number)
End Function

' R23-509. Counts the graph's small-hole locations independently of any
' view, so a view that fails to project them is visible as a shortfall
' rather than redefining the target.
Public Function CountSmallHoleLocations( _
    ByRef graph As CLocationGraph) As Long

    On Error GoTo Failed

    Dim total As Long
    Dim locations As Collection
    Set locations = graph.Locations()

    Dim i As Long
    For i = 1 To locations.Count
        Dim location As CPhysicalHoleLocation
        Set location = locations(i)
        If IsSmallHoleLocation(graph, location) Then total = total + 1
    Next i

    CountSmallHoleLocations = total
    Exit Function

Failed:
    CountSmallHoleLocations = -1
End Function

Private Sub AddDistinctKey( _
    ByRef keys As Collection, _
    ByVal value As String)

    If Len(Trim$(value)) = 0 Then Exit Sub

    Dim i As Long
    For i = 1 To keys.Count
        If StrComp(CStr(keys(i)), value, vbBinaryCompare) = 0 Then Exit Sub
    Next i

    keys.Add value
End Sub

Private Function AppendFailure( _
    ByVal existing As String, _
    ByVal reason As String) As String

    If Len(existing) = 0 Then
        AppendFailure = reason
    Else
        AppendFailure = existing & ";" & reason
    End If
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

Private Function SafeViewType( _
    ByRef swView As SldWorks.View) As Long

    On Error Resume Next
    SafeViewType = -1
    If swView Is Nothing Then Exit Function
    SafeViewType = swView.Type
End Function

' R23-500 to R23-509 evidence entry point. STRICTLY READ-ONLY: it contains
' no call to CreateOrdinateGroup and therefore cannot create a dimension.
Public Sub R23_ProbeOrdinateScheme()
    On Error GoTo Failed

    mEmitDiagnostics = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Debug.Print "R23_ORDINATE_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Debug.Print "R23_ORDINATE_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Debug.Print "R23_ORDINATE_FATAL|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Debug.Print "R23_ORDINATE_FATAL|reason=NoViewsOnSheet"
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
        Debug.Print "R23_ORDINATE_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Debug.Print "R23_ORDINATE_FATAL|reason=UnauthorizedFixture" & _
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

    Debug.Print "R23_ORDINATE_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0"

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        Debug.Print "R23_ORDINATE_FATAL|reason=CatalogUnavailable"
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

    Debug.Print "R23_ORDINATE_GRAPH|" & graph.GraphSummary()

    Dim smallHoleLocations As Long
    smallHoleLocations = CountSmallHoleLocations(graph)

    Debug.Print "R23_ORDINATE_LEDGER|smallHoleLocations=" & _
        CStr(smallHoleLocations) & _
        "|totalLocations=" & CStr(graph.LocationCount()) & _
        "|families=" & CStr(graph.FamilyKeys().Count) & _
        "|rule=FamilySize>=" & CStr(MIN_FAMILY_SIZE_FOR_SMALL_HOLE)

    Dim allSchemes As Collection
    Set allSchemes = New Collection

    For i = LBound(views) To UBound(views)
        Dim schemeView As SldWorks.View
        Set schemeView = views(i)
        If schemeView Is Nothing Then GoTo ContinueSchemeView

        Dim viewSchemes As Collection
        Set viewSchemes = BuildSchemesForView( _
            swApp, swDraw, schemeView, graph, evidence)

        Dim s As Long
        For s = 1 To viewSchemes.Count
            Dim scheme As COrdinateScheme
            Set scheme = viewSchemes(s)
            allSchemes.Add scheme

            Debug.Print "R23_ORDINATE_SCHEME|" & scheme.Summary()

            Dim b As Long
            For b = 1 To scheme.Buckets.Count
                Dim bucket As COrdinateBucket
                Set bucket = scheme.Buckets(b)

                Dim bucketProof As String
                bucketProof = "selection=NotAttempted"

                If bucket.HasSelectableAnchor() Then
                    ProveBucketSelectable _
                        swDraw, scheme, bucket, bucketProof
                End If

                Debug.Print "QA INFO: ORDINATE_BUCKET|" & _
                    scheme.SchemeKey() & "|" & bucket.Summary() & _
                    "|credits=" & bucket.CreditedKeyList() & _
                    "|" & bucketProof
            Next b

            Debug.Print "QA INFO: ORDINATE_DATUM_PROOF|" & _
                scheme.SchemeKey() & "|" & scheme.DatumProof
        Next s

ContinueSchemeView:
    Next i

    Dim coverage As String
    coverage = VerifyDirectionalCoverage(allSchemes, smallHoleLocations)

    Dim finalSelectionCount As Long
    finalSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim drawingSaveAfter As Boolean
    drawingSaveAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Debug.Print "R23_ORDINATE_END|" & coverage & _
        "|creations=0" & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & _
        CStr(drawingSaveBefore = drawingSaveAfter)

    ' Deliberately NOT Module6_QAEngine.EmitRunEvidence. That is the
    ' production gate: RequireCoreStages demands fourteen pipeline stages -
    ' CONTROLLED_SHEET, LAYOUT, TITLE_PROPERTIES, FINAL_QA and the rest -
    ' that a read-only probe never runs, so it fails closed and reports
    ' RESULT: FAIL for a probe that in fact succeeded. The Phase 3 and
    ' Phase 4 probes do not call it either. Evidence still reaches the
    ' Immediate window: CRunEvidence.AddInfo/AddWarning/AddFailure each
    ' Debug.Print their own line.
    Exit Sub

Failed:
    Debug.Print "R23_ORDINATE_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub

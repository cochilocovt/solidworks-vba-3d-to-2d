Option Explicit

' R23 Phase 7. Section-line path built from model intent.
' SAFETY BOUNDARY. Exactly one procedure changes a drawing:
' CreateSectionFromPath, which refuses unless passed an explicit
' allowMutation argument and refuses again unless the path's crossings are
' proved. R23_ProbeSectionPath never passes it and contains no CreateLine or
' CreateSectionViewAt5 call.
' R23-700. This module consumes the location graph and never touches the
' ordinate engine: section construction reads proved locations, it does not
' wait on or interact with dimension creation. There is deliberately no
' reference to Module15_OrdinateScheme anywhere below.
' R23-704. None of the disproved strategy survives here: no `extension`, no
' `topY + extension`, no `bottomY - extension`, and no fraction of the view
' outline. Every waypoint is a proved projection's page coordinate. The old
' approach put the upper label in the zone region and the lower arrow in the
' part-identification band, because a percentage of an outline knows nothing
' about where the holes are. Those literals still exist in
' Module2_DrawingPipeline.bas, which is the reachable production path until
' the pipeline switches over.

' R23-703. A path segment counts as crossing a hole when it passes within
' the hole's own projected radius of its centre. The radius comes from the
' projection, so a large bore and a small tapped hole are judged on their
' own size rather than against a shared tolerance.
' This small slack absorbs the page-coordinate quantum only; it is not a
' search radius. Widening it would let a segment "cross" a hole it misses.
Public Const CROSSING_SLACK_M As Double = 0.000001

' Waypoints are compared for distinctness at the same quantum Phase 5 used
' for ordinate coordinates, so a column and a row mean the same thing in
' both phases.
Public Const PATH_COORDINATE_QUANTUM_M As Double = 0.000001

' R23-815. MCP corpus value for
' swCreateSectionViewAtOptions_e.swCreateSectionView_OffsetSection; verify
' in the SW2025 Object Browser. Documented as: "If set, then an aligned
' section view is created (two lines at an angle); if not set, a normal
' projection section view is created."
'
' CreateSectionViewAt5 was called with Options=0 until r53, so SOLIDWORKS
' built a NORMAL PROJECTION section from this module's three-segment jogged
' line and cut at a single offset. Measured in run
' macro_qa/20260805_050411_P-0251-14A-001, where GetSectionLineInfo2
' returned exactly the requested path -
'   seg1 (-0.102,0.000)->(0.008,0.000)
'   seg2 (0.008,0.000)->(0.008,-0.015)
'   seg3 (0.008,-0.015)->(0.088,-0.015)
' - while the section view held the counterbore-column features and no bore
' at all. The bore is at transverse 0.000 and the counterbores at -0.015,
' so the cut in use was segment 3's alone. That is also why r51 could
' lengthen segment 1 by 40 mm and produce a byte-for-byte identical section
' view: a segment that is not cut along cannot change the cut.
Private Const SECTION_OPTION_OFFSET As Long = 2

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

' R23-701 and R23-702. Builds the path for one view.
' Returns a CSectionPath whose Resolved flag says whether every input was
' proved. An unresolved path carries the specific reason and is never
' partially usable.
Public Function ResolveSectionPath( _
    ByRef graph As CLocationGraph, _
    ByRef swView As SldWorks.View, _
    ByVal sectionLabel As String, _
    ByRef evidence As CRunEvidence) As CSectionPath

    On Error GoTo Failed

    Dim path As CSectionPath
    Set path = New CSectionPath
    Set ResolveSectionPath = path

    path.SectionLabel = sectionLabel
    Set path.SourceView = swView
    path.SourceViewName = SafeViewName(swView)

    If graph Is Nothing Or swView Is Nothing Then
        path.RejectionReason = "NoGraphOrView"
        Exit Function
    End If

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(path.SourceViewName)

    If projections.Count = 0 Then
        path.RejectionReason = "NoProjectionsInView"
        Exit Function
    End If

    ' R23-701, first half. The stepped-bore centre is the projection of the
    ' largest singleton-family location: singleton because a bore is not one
    ' of a repeated family, largest because the section is cut through the
    ' principal bore. Family size comes from the graph, so this is not a
    ' radius threshold. The projection must carry a proved page position;
    ' see ResolveBoreProjection for why a selectable anchor is the wrong
    ' requirement here.
    Dim boreBasis As String
    Set path.BoreProjection = ResolveBoreProjection( _
        graph, projections, boreBasis)
    path.BoreProjectionBasis = boreBasis

    If path.BoreProjection Is Nothing Then
        path.RejectionReason = "NoUsableSingletonBoreProjection"
        Exit Function
    End If

    ' A position-proved bore is a deliberate, narrower proof than an accepted
    ' one, so it is stated in evidence every time rather than being silently
    ' equivalent.
    If StrComp(Left$(boreBasis, 14), "PositionProved", _
        vbBinaryCompare) = 0 Then

        EmitInfo evidence, "SECTION_PATH_BORE_BASIS|view=" & _
            path.SourceViewName & _
            "|physical=" & path.BoreProjection.PhysicalInstanceKey & _
            "|basis=" & boreBasis & _
            "|anchor=" & CStr(path.BoreProjection.HasSelectableAnchor()) & _
            "|frame=" & path.BoreProjection.CoordinateFrameProof & _
            "|axisNormal=" & CStr(path.BoreProjection.AxisNormalToView) & _
            "|use=WaypointsOnly"
    End If

    ' R23-701, second half. The face-hole family shares the bore's machining
    ' face and is a repeated family. Its projections must form a grid, which
    ' the column and row counts below prove rather than assume.
    Dim faceHoles As Collection
    Set faceHoles = CollectFaceHoleProjections( _
        graph, projections, path.BoreProjection)

    If faceHoles.Count = 0 Then
        path.RejectionReason = "NoAcceptedFaceHoleProjections"
        Exit Function
    End If

    Dim columns As Collection
    Dim rows As Collection
    Set columns = DistinctCoordinates(faceHoles, True)
    Set rows = DistinctCoordinates(faceHoles, False)

    path.DistinctColumns = columns.Count
    path.DistinctRows = rows.Count

    If columns.Count < 2 Then
        path.RejectionReason = "FewerThanTwoFaceHoleColumns"
        Exit Function
    End If

    If rows.Count < 2 Then
        path.RejectionReason = "FewerThanTwoFaceHoleRows"
        Exit Function
    End If

    ' R23-702. Four waypoints, in the order the reference approves.
    Dim highestRowY As Double
    Dim lowestRowY As Double
    highestRowY = ExtremeValue(rows, True)
    lowestRowY = ExtremeValue(rows, False)

    Dim minimumColumnX As Double
    minimumColumnX = ExtremeValue(columns, False)

    ' R23-813. Waypoint 1 starts BEYOND the bore, not at its centre.
    '
    ' Measured in run macro_qa/20260805_041027_P-0251-14A-001: Section View
    ' J-J contained no bore geometry at all. Its only arc radii were
    ' 0.002100 (the M5 tap drill) and 0.035/0.036 (the plate's rounded
    ' top); neither 0.020 nor 0.023500 appeared, and no 0.040 or 0.047 span
    ' existed on either axis. The bore's cut walls would sit at
    ' Y = 0.062 +/- 0.023500 and the only nearby coordinate was 0.062
    ' itself - the centre, which is where waypoint 1 used to sit. A cut
    ' that starts at the centre removes half the bore and shows no opening
    ' to dimension, which is why INNER_BORE_D40 and FIT_BORE_D47_H7 had no
    ' candidate geometry.
    '
    ' The overshoot is the bore's OWN size: twice the projected radius from
    ' the centre, so the line ends one full radius past the far wall. Same
    ' principle as the crossing slack in R23-703 - a large bore and a small
    ' tapped hole are judged on their own size, never against a shared
    ' literal. No fraction of the view outline is involved.
    '
    ' Direction comes from the geometry rather than an assumption about
    ' which way up the part sits: away from the face-hole rows, which are
    ' where the rest of the path goes.
    path.BoreOvershootM = 2# * path.BoreProjection.ProjectedRadiusM

    Dim overshootDirection As String
    path.W1X = path.BoreProjection.PageX

    If path.BoreProjection.PageY >= highestRowY Then
        path.W1Y = path.BoreProjection.PageY + path.BoreOvershootM
        overshootDirection = "AwayFromRows:PositiveY"
    Else
        path.W1Y = path.BoreProjection.PageY - path.BoreOvershootM
        overshootDirection = "AwayFromRows:NegativeY"
    End If

    path.W2X = path.BoreProjection.PageX
    path.W2Y = highestRowY

    EmitInfo evidence, "SECTION_PATH_BORE_OVERSHOOT|view=" & _
        path.SourceViewName & _
        "|centreY=" & Format$(path.BoreProjection.PageY, "0.000000000") & _
        "|projectedRadiusM=" & _
            Format$(path.BoreProjection.ProjectedRadiusM, "0.000000000") & _
        "|overshootM=" & Format$(path.BoreOvershootM, "0.000000000") & _
        "|w1Y=" & Format$(path.W1Y, "0.000000000") & _
        "|direction=" & overshootDirection & _
        "|reason=CutMustCrossWholeBore"
    path.W3X = minimumColumnX
    path.W3Y = highestRowY
    path.W4X = minimumColumnX
    path.W4Y = lowestRowY

    CollectColumnProjections faceHoles, minimumColumnX, path

    ' R23-703.
    ProvePathCrossings path, evidence

    If Len(path.CrossingFailures) > 0 And _
        StrComp(path.CrossingFailures, "None", vbBinaryCompare) <> 0 Then

        path.RejectionReason = "CrossingUnproven:" & path.CrossingFailures
        Exit Function
    End If

    path.Resolved = True
    path.RejectionReason = "None"

    EmitInfo evidence, "SECTION_PATH|" & path.Summary()
    Exit Function

Failed:
    path.RejectionReason = "Error:" & CStr(Err.Number)
    EmitFailure evidence, "SECTION_PATH_ERROR|view=" & _
        path.SourceViewName & "|error=" & CStr(Err.Number)
End Function

' The projection whose location stands alone in its family and has the
' largest radius. Singleton family membership is read from the graph.
'
' The path reads PageX, PageY and ProjectedRadiusM from this projection and
' never selects it, so a proved page POSITION is the correct requirement here
' - not a selectable anchor. Requiring Accepted made the section
' unreachable for the one part it exists to draw: run
' macro_qa/20260805_033146_P-0251-14A-001 proved the stepped bore is
' completely obscured in every orthographic view, so it can never carry an
' orthographic anchor, and the section that would finally show it demanded
' that anchor. Circular. User decision, 2026-08-05: the section path takes
' position-proved projections; the dimensioning, callout and ordinate paths
' keep requiring CViewHoleProjection.QualificationFailureReason.
'
' Largest radius still wins over acceptance state, because "the principal
' bore" is a fact about the part. Preferring a smaller accepted bore would
' cut the section through the wrong feature. basis records which proof the
' chosen projection actually carried.
Private Function ResolveBoreProjection( _
    ByRef graph As CLocationGraph, _
    ByRef projections As Collection, _
    ByRef basis As String) As CViewHoleProjection

    On Error GoTo Failed

    basis = "None"

    Dim best As CViewHoleProjection
    Set best = Nothing

    Dim i As Long
    For i = 1 To projections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = projections(i)

        If candidate Is Nothing Then GoTo ContinueCandidate
        If candidate.PhysicalLocation Is Nothing Then GoTo ContinueCandidate

        If Not candidate.Accepted Then
            If Not candidate.HasProvedPosition() Then GoTo ContinueCandidate
        End If

        Dim family As Collection
        Set family = graph.LocationsForFamily( _
            candidate.PhysicalLocation.SemanticFamilyKey)

        If family.Count <> 1 Then GoTo ContinueCandidate

        If best Is Nothing Then
            Set best = candidate
        ElseIf candidate.PhysicalLocation.MaximumRadiusM > _
            best.PhysicalLocation.MaximumRadiusM Then

            Set best = candidate
        End If

ContinueCandidate:
    Next i

    If Not best Is Nothing Then
        If best.Accepted Then
            basis = "Accepted"
        Else
            basis = "PositionProved:" & best.RejectionReason
        End If
    End If

    Set ResolveBoreProjection = best
    Exit Function

Failed:
    basis = "Error:" & CStr(Err.Number)
    Set ResolveBoreProjection = Nothing
End Function

' Accepted projections of repeated families sharing the bore's machining
' face. Sharing the face is what makes them face holes rather than side
' holes, and it is read from the location's own axis.
Private Function CollectFaceHoleProjections( _
    ByRef graph As CLocationGraph, _
    ByRef projections As Collection, _
    ByRef boreProjection As CViewHoleProjection) As Collection

    Dim result As Collection
    Set result = New Collection
    Set CollectFaceHoleProjections = result

    On Error GoTo Failed

    Dim boreFace As String
    boreFace = MachiningFaceToken(boreProjection)

    Dim i As Long
    For i = 1 To projections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = projections(i)

        If candidate Is Nothing Then GoTo ContinueCandidate
        If Not candidate.Accepted Then GoTo ContinueCandidate
        If candidate.PhysicalLocation Is Nothing Then GoTo ContinueCandidate

        If StrComp(MachiningFaceToken(candidate), boreFace, _
            vbBinaryCompare) <> 0 Then

            GoTo ContinueCandidate
        End If

        Dim family As Collection
        Set family = graph.LocationsForFamily( _
            candidate.PhysicalLocation.SemanticFamilyKey)

        If family.Count < 2 Then GoTo ContinueCandidate

        result.Add candidate

ContinueCandidate:
    Next i

    Exit Function

Failed:
    Set CollectFaceHoleProjections = result
End Function

Private Function MachiningFaceToken( _
    ByRef projection As CViewHoleProjection) As String

    On Error GoTo Failed

    If projection Is Nothing Then Exit Function
    If projection.PhysicalLocation Is Nothing Then Exit Function

    MachiningFaceToken = _
        Module11_GeometryIdentity.CanonicalAxisToken( _
            projection.PhysicalLocation.AxisX, _
            projection.PhysicalLocation.AxisY, _
            projection.PhysicalLocation.AxisZ)
    Exit Function

Failed:
    MachiningFaceToken = vbNullString
End Function

' Distinct quantized X (columns) or Y (rows) among the given projections.
Private Function DistinctCoordinates( _
    ByRef projections As Collection, _
    ByVal useX As Boolean) As Collection

    Dim result As Collection
    Set result = New Collection
    Set DistinctCoordinates = result

    Dim i As Long
    For i = 1 To projections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = projections(i)

        Dim value As Double
        If useX Then
            value = candidate.PageX
        Else
            value = candidate.PageY
        End If

        Dim token As String
        token = QuantizeCoordinate(value)

        Dim found As Boolean
        found = False

        Dim j As Long
        For j = 1 To result.Count
            If StrComp(QuantizeCoordinate(CDbl(result(j))), token, _
                vbBinaryCompare) = 0 Then

                found = True
                Exit For
            End If
        Next j

        If Not found Then result.Add value
    Next i
End Function

Private Function ExtremeValue( _
    ByRef values As Collection, _
    ByVal wantMaximum As Boolean) As Double

    Dim best As Double
    Dim seeded As Boolean

    Dim i As Long
    For i = 1 To values.Count
        Dim candidate As Double
        candidate = CDbl(values(i))

        If Not seeded Then
            best = candidate
            seeded = True
        ElseIf wantMaximum Then
            If candidate > best Then best = candidate
        Else
            If candidate < best Then best = candidate
        End If
    Next i

    ExtremeValue = best
End Function

Public Function QuantizeCoordinate(ByVal value As Double) As String
    Dim quantized As Double
    quantized = _
        Int(value / PATH_COORDINATE_QUANTUM_M + 0.5) * _
        PATH_COORDINATE_QUANTUM_M

    QuantizeCoordinate = Format$(quantized, "0.000000000")
End Function

' Every face-hole projection sitting on the chosen column, which is the set
' the final segment must cross.
Private Sub CollectColumnProjections( _
    ByRef faceHoles As Collection, _
    ByVal columnX As Double, _
    ByRef path As CSectionPath)

    Dim columnToken As String
    columnToken = QuantizeCoordinate(columnX)

    Dim i As Long
    For i = 1 To faceHoles.Count
        Dim candidate As CViewHoleProjection
        Set candidate = faceHoles(i)

        If StrComp(QuantizeCoordinate(candidate.PageX), columnToken, _
            vbBinaryCompare) = 0 Then

            path.ColumnProjections.Add candidate
        End If
    Next i
End Sub

' R23-703. Proves the path crosses the stepped bore and every hole in the
' chosen column, reporting each one that it misses.
Private Sub ProvePathCrossings( _
    ByRef path As CSectionPath, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim failures As String

    ' R23-813. The bore is held to the stronger predicate: the cut must
    ' pass ALL THE WAY THROUGH it, entering one wall and leaving the other
    ' inside the same segment. PathCrossesCircle only asks whether a
    ' segment comes within the radius, which a segment starting at the
    ' centre satisfies trivially - and did, reporting
    ' crossingsProven=4|crossingFailures=None for a cut that removed half
    ' the bore and produced a section with no bore in it at all
    ' (macro_qa/20260805_041027_P-0251-14A-001). A predicate that passes
    ' the case it exists to reject is worse than no predicate.
    '
    ' Column holes keep the weaker test on purpose: the first and last hole
    ' on the chosen column sit AT the segment's endpoints, so requiring a
    ' full crossing of those would refuse the correct path. They need to be
    ' on the cut; only the bore needs its whole opening shown.
    If Not PathFullyCrossesCircle(path, path.BoreProjection) Then
        failures = AppendFailure(failures, "BoreNotFullyCrossed")
    Else
        path.CrossingProofs.Add "bore:" & _
            path.BoreProjection.PhysicalInstanceKey
    End If

    Dim i As Long
    For i = 1 To path.ColumnProjections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = path.ColumnProjections(i)

        If PathCrossesCircle(path, candidate) Then
            path.CrossingProofs.Add "columnHole:" & _
                candidate.PhysicalInstanceKey
        Else
            failures = AppendFailure(failures, _
                "ColumnHoleNotCrossed:" & candidate.PhysicalInstanceKey)
        End If
    Next i

    If path.ColumnProjections.Count = 0 Then
        failures = AppendFailure(failures, "NoHolesOnChosenColumn")
    End If

    If Len(failures) = 0 Then failures = "None"
    path.CrossingFailures = failures

    EmitInfo evidence, "SECTION_CROSSING|view=" & path.SourceViewName & _
        "|proven=" & CStr(path.CrossingProofs.Count) & _
        "|columnHoles=" & CStr(path.ColumnProjections.Count) & _
        "|failures=" & failures
    Exit Sub

Failed:
    path.CrossingFailures = "Error:" & CStr(Err.Number)
End Sub

' True when any of the three segments passes within the projection's own
' projected radius of its centre.
Private Function PathCrossesCircle( _
    ByRef path As CSectionPath, _
    ByRef projection As CViewHoleProjection) As Boolean

    On Error GoTo Failed

    If projection Is Nothing Then Exit Function

    Dim radius As Double
    radius = projection.ProjectedRadiusM + CROSSING_SLACK_M

    If SegmentReaches(path.W1X, path.W1Y, path.W2X, path.W2Y, _
        projection.PageX, projection.PageY, radius) Then

        PathCrossesCircle = True
        Exit Function
    End If

    If SegmentReaches(path.W2X, path.W2Y, path.W3X, path.W3Y, _
        projection.PageX, projection.PageY, radius) Then

        PathCrossesCircle = True
        Exit Function
    End If

    If SegmentReaches(path.W3X, path.W3Y, path.W4X, path.W4Y, _
        projection.PageX, projection.PageY, radius) Then

        PathCrossesCircle = True
        Exit Function
    End If

    Exit Function

Failed:
    PathCrossesCircle = False
End Function

' R23-813. True when one segment passes entirely through the circle: it
' enters one wall and leaves the other, both inside the segment. This is
' what "the section shows the whole bore" means geometrically, and it is
' strictly stronger than PathCrossesCircle above.
Private Function PathFullyCrossesCircle( _
    ByRef path As CSectionPath, _
    ByRef projection As CViewHoleProjection) As Boolean

    On Error GoTo Failed

    If projection Is Nothing Then Exit Function

    Dim radius As Double
    radius = projection.ProjectedRadiusM + CROSSING_SLACK_M

    If SegmentSpansCircle(path.W1X, path.W1Y, path.W2X, path.W2Y, _
        projection.PageX, projection.PageY, radius) Then

        PathFullyCrossesCircle = True
        Exit Function
    End If

    If SegmentSpansCircle(path.W2X, path.W2Y, path.W3X, path.W3Y, _
        projection.PageX, projection.PageY, radius) Then

        PathFullyCrossesCircle = True
        Exit Function
    End If

    If SegmentSpansCircle(path.W3X, path.W3Y, path.W4X, path.W4Y, _
        projection.PageX, projection.PageY, radius) Then

        PathFullyCrossesCircle = True
        Exit Function
    End If

    Exit Function

Failed:
    PathFullyCrossesCircle = False
End Function

' Both intersections of the infinite line with the circle must lie inside
' the finite segment. The foot of the perpendicular is at distance
' footDistance along the segment and the two intersections sit half a chord
' either side of it, so the test is exactly whether that interval fits.
Private Function SegmentSpansCircle( _
    ByVal startX As Double, _
    ByVal startY As Double, _
    ByVal endX As Double, _
    ByVal endY As Double, _
    ByVal centreX As Double, _
    ByVal centreY As Double, _
    ByVal radius As Double) As Boolean

    Dim dx As Double
    Dim dy As Double
    dx = endX - startX
    dy = endY - startY

    Dim lengthSquared As Double
    lengthSquared = dx * dx + dy * dy
    If lengthSquared <= 0# Then Exit Function

    Dim segmentLength As Double
    segmentLength = Sqr(lengthSquared)

    Dim t As Double
    t = ((centreX - startX) * dx + (centreY - startY) * dy) / lengthSquared

    Dim footX As Double
    Dim footY As Double
    footX = startX + (t * dx)
    footY = startY + (t * dy)

    Dim perpendicular As Double
    perpendicular = Sqr(((centreX - footX) ^ 2) + ((centreY - footY) ^ 2))

    ' The line misses the circle entirely.
    If perpendicular > radius Then Exit Function

    Dim halfChord As Double
    halfChord = Sqr((radius * radius) - (perpendicular * perpendicular))

    Dim footDistance As Double
    footDistance = t * segmentLength

    SegmentSpansCircle = _
        (footDistance - halfChord >= 0#) And _
        (footDistance + halfChord <= segmentLength)
End Function

' Shortest distance from a point to a finite segment, compared to a radius.
' The parameter is clamped to the segment, so a circle beyond an endpoint is
' correctly reported as not crossed even when the infinite line would pass
' through it.
Private Function SegmentReaches( _
    ByVal startX As Double, _
    ByVal startY As Double, _
    ByVal endX As Double, _
    ByVal endY As Double, _
    ByVal pointX As Double, _
    ByVal pointY As Double, _
    ByVal radius As Double) As Boolean

    Dim dx As Double
    Dim dy As Double
    dx = endX - startX
    dy = endY - startY

    Dim lengthSquared As Double
    lengthSquared = dx * dx + dy * dy

    Dim t As Double
    If lengthSquared <= 0# Then
        t = 0#
    Else
        t = ((pointX - startX) * dx + (pointY - startY) * dy) / lengthSquared
        If t < 0# Then t = 0#
        If t > 1# Then t = 1#
    End If

    Dim nearestX As Double
    Dim nearestY As Double
    nearestX = startX + t * dx
    nearestY = startY + t * dy

    Dim separationX As Double
    Dim separationY As Double
    separationX = pointX - nearestX
    separationY = pointY - nearestY

    SegmentReaches = _
        ((separationX * separationX + separationY * separationY) <= _
         (radius * radius))
End Function

' R23-705. Converts one PAGE point into the source view's sketch frame.
' Called exactly once per waypoint, immediately before CreateLine, and never
' anywhere else. IView.GetXform returns the view's sheet origin and scale;
' IView.Angle returns its rotation. Converting twice - or converting a value
' that is already in the view frame - is the mixed-frame defect this project
' has hit before, which is why nothing upstream of here holds view
' coordinates at all.
Public Function PageToViewSketch( _
    ByRef swView As SldWorks.View, _
    ByVal pageX As Double, _
    ByVal pageY As Double, _
    ByRef viewX As Double, _
    ByRef viewY As Double, _
    ByRef frameProof As String) As Boolean

    On Error GoTo Failed

    frameProof = "frame=Unproven"

    If swView Is Nothing Then Exit Function

    Dim xform As Variant
    xform = swView.GetXform

    If IsEmpty(xform) Or Not IsArray(xform) Then
        frameProof = "frame=Reject|reason=NoXform"
        Exit Function
    End If

    If (UBound(xform) - LBound(xform) + 1) < 3 Then
        frameProof = "frame=Reject|reason=XformTooShort"
        Exit Function
    End If

    Dim baseIndex As Long
    baseIndex = LBound(xform)

    Dim originX As Double
    Dim originY As Double
    Dim viewScale As Double
    originX = CDbl(xform(baseIndex))
    originY = CDbl(xform(baseIndex + 1))
    viewScale = CDbl(xform(baseIndex + 2))

    If viewScale <= 0# Then
        frameProof = "frame=Reject|reason=NonPositiveScale"
        Exit Function
    End If

    Dim deltaX As Double
    Dim deltaY As Double
    deltaX = (pageX - originX) / viewScale
    deltaY = (pageY - originY) / viewScale

    Dim viewAngle As Double
    viewAngle = swView.Angle

    viewX = deltaX * Cos(viewAngle) + deltaY * Sin(viewAngle)
    viewY = -deltaX * Sin(viewAngle) + deltaY * Cos(viewAngle)

    frameProof = "frame=ViewSketch" & _
        "|source=IView.GetXform+IView.Angle" & _
        "|originX=" & Format$(originX, "0.000000000") & _
        "|originY=" & Format$(originY, "0.000000000") & _
        "|scale=" & Format$(viewScale, "0.000000000") & _
        "|angle=" & Format$(viewAngle, "0.000000000") & _
        "|conversions=1"

    PageToViewSketch = True
    Exit Function

Failed:
    frameProof = "frame=Reject|reason=Error:" & CStr(Err.Number)
    PageToViewSketch = False
End Function

' R23-706 and R23-707. MUTATES THE DRAWING.
' Refuses unless allowMutation is True AND the path resolved with its
' crossings proved. Creates exactly three view-owned sketch segments,
' verifies the selection count before calling CreateSectionViewAt5, then
' reads GetSectionLineInfo2 back.
' placeX and placeY are the SHEET position for the CENTRE of the new section
' view, per the CreateSectionViewAt5 contract. They are caller arguments and
' deliberately not derived here: choosing where a view sits is layout, which
' is a later phase's responsibility and needs the full annotation envelopes
' this module cannot see. Defaulting them to a point on the source view -
' the obvious shortcut - would stack the section on top of the view it was
' cut from.
' Returns the created section view, or Nothing on refusal or failure.
Public Function CreateSectionFromPath( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef path As CSectionPath, _
    ByVal placeX As Double, _
    ByVal placeY As Double, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As SldWorks.View

    On Error GoTo Failed

    If Not allowMutation Then
        EmitWarning evidence, "SECTION_CREATE_REFUSED" & _
            "|reason=MutationNotAuthorized"
        Exit Function
    End If

    If path Is Nothing Then Exit Function

    ' R23-708. An unproved path is never approximated into existence.
    If Not path.Resolved Then
        EmitFailure evidence, "SECTION_CREATE_REFUSED" & _
            "|reason=PathUnresolved|detail=" & path.RejectionReason
        Exit Function
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    ' The raw IDrawingDoc.ActivateView result is a false negative on this
    ' build. Gating the cut on it would refuse a section whose source view had
    ' in fact activated, so activation is proved by active-view readback.
    If path.SourceView Is Nothing Then
        EmitFailure evidence, "SECTION_CREATE_REFUSED" & _
            "|reason=SourceViewUnavailable|view=" & path.SourceViewName
        Exit Function
    End If

    If Not Module8_RuntimeSupport.ActivateDrawingView( _
        swDraw, swDrawing, path.SourceView, evidence, _
        "Section cut from '" & path.SourceViewName & "'") Then

        EmitFailure evidence, "SECTION_CREATE_REFUSED" & _
            "|reason=ViewActivationFailed|view=" & path.SourceViewName
        Exit Function
    End If

    swDraw.ClearSelection2 True

    Dim waypointX(1 To 4) As Double
    Dim waypointY(1 To 4) As Double
    waypointX(1) = path.W1X
    waypointY(1) = path.W1Y
    waypointX(2) = path.W2X
    waypointY(2) = path.W2Y
    waypointX(3) = path.W3X
    waypointY(3) = path.W3Y
    waypointX(4) = path.W4X
    waypointY(4) = path.W4Y

    ' R23-705. The single conversion, here and nowhere else.
    Dim sketchX(1 To 4) As Double
    Dim sketchY(1 To 4) As Double

    Dim w As Long
    For w = 1 To 4
        Dim frameProof As String

        If Not PageToViewSketch(path.SourceView, _
            waypointX(w), waypointY(w), _
            sketchX(w), sketchY(w), frameProof) Then

            EmitFailure evidence, "SECTION_CREATE_REFUSED" & _
                "|reason=FrameConversionFailed" & _
                "|waypoint=" & CStr(w) & "|" & frameProof
            swDraw.ClearSelection2 True
            Exit Function
        End If

        EmitInfo evidence, "SECTION_WAYPOINT|view=" & _
            path.SourceViewName & _
            "|index=" & CStr(w) & _
            "|pageX=" & Format$(waypointX(w), "0.000000000") & _
            "|pageY=" & Format$(waypointY(w), "0.000000000") & _
            "|viewX=" & Format$(sketchX(w), "0.000000000") & _
            "|viewY=" & Format$(sketchY(w), "0.000000000") & _
            "|" & frameProof
    Next w

    Dim segments As Collection
    Set segments = New Collection

    Dim s As Long
    For s = 1 To 3
        Dim segment As SldWorks.SketchSegment
        Set segment = swDraw.SketchManager.CreateLine( _
            sketchX(s), sketchY(s), 0#, _
            sketchX(s + 1), sketchY(s + 1), 0#)

        If segment Is Nothing Then
            EmitFailure evidence, "SECTION_CREATE_FAILED" & _
                "|reason=CreateLineReturnedNothing|segment=" & CStr(s)
            swDraw.ClearSelection2 True
            Exit Function
        End If

        segments.Add segment
    Next s

    path.SegmentsCreated = segments.Count

    If segments.Count <> path.SegmentCount() Then
        EmitFailure evidence, "SECTION_CREATE_FAILED" & _
            "|reason=SegmentCountMismatch" & _
            "|created=" & CStr(segments.Count) & _
            "|expected=" & CStr(path.SegmentCount())
        swDraw.ClearSelection2 True
        Exit Function
    End If

    ' R23-706. Selection order is deterministic and verified before the
    ' section is cut: SOLIDWORKS reads the segments in selection order, so
    ' an unverified order produces a differently shaped cut.
    swDraw.ClearSelection2 True

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDraw.SelectionManager

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionMgr.CreateSelectData

    Dim expectedCount As Long
    expectedCount = 0

    Dim i As Long
    For i = 1 To segments.Count
        Dim toSelect As SldWorks.SketchSegment
        Set toSelect = segments(i)

        Dim appended As Boolean
        appended = Module11_GeometryIdentity.NormalizeSwBoolean( _
            toSelect.Select4(True, selectData))

        Dim actualCount As Long
        actualCount = selectionMgr.GetSelectedObjectCount2(-1)

        If Not appended Or actualCount <> expectedCount + 1 Then
            EmitFailure evidence, "SECTION_CREATE_FAILED" & _
                "|reason=SegmentSelectionOrderUnverified" & _
                "|segment=" & CStr(i) & _
                "|appended=" & CStr(appended) & _
                "|expectedCount=" & CStr(expectedCount + 1) & _
                "|actualCount=" & CStr(actualCount)
            swDraw.ClearSelection2 True
            Exit Function
        End If

        expectedCount = actualCount
    Next i

    ' CreateSectionViewAt5(X, Y, Z, SectionLabel, Options,
    ' ExcludedComponents, SectionDepth). Verified against the 2025 Help,
    ' whose Remarks require the section line to be selected first - which
    ' the verified append loop above has just done.
    '
    ' R23-815. Options carries swCreateSectionView_OffsetSection because
    ' this module always builds a jogged line; see the constant for the
    ' measurement that showed Options=0 discarding every segment but one.
    ' The option is recorded in evidence rather than left implicit, because
    ' it changes what the cut contains and nothing else in the report would
    ' show it.
    Dim sectionOptions As Long
    sectionOptions = SECTION_OPTION_OFFSET

    EmitInfo evidence, "SECTION_CREATE_OPTIONS|view=" & _
        path.SourceViewName & _
        "|options=" & CStr(sectionOptions) & _
        "|offsetSection=True" & _
        "|segments=" & CStr(segments.Count) & _
        "|source=swCreateSectionViewAtOptions_e"

    Dim sectionView As SldWorks.View
    Set sectionView = swDrawing.CreateSectionViewAt5( _
        placeX, placeY, 0#, path.SectionLabel, sectionOptions, Nothing, 0#)

    swDraw.ClearSelection2 True

    If sectionView Is Nothing Then
        EmitFailure evidence, "SECTION_CREATE_FAILED" & _
            "|reason=CreateSectionViewAt5ReturnedNothing"
        Exit Function
    End If

    path.SectionViewName = SafeViewName(sectionView)
    Set CreateSectionFromPath = sectionView

    ' R23-707.
    path.SectionLineInfoProof = ReadSectionLineInfo(path.SourceView)

    EmitInfo evidence, "SECTION_CREATED|view=" & path.SourceViewName & _
        "|sectionView=" & path.SectionViewName & _
        "|segments=" & CStr(path.SegmentsCreated) & _
        "|selectionsVerified=" & CStr(expectedCount) & _
        "|" & path.SectionLineInfoProof
    Exit Function

Failed:
    EmitFailure evidence, "SECTION_CREATE_ERROR|error=" & _
        CStr(Err.Number) & "|description=" & Err.Description

    On Error Resume Next
    swDraw.ClearSelection2 True
    Set CreateSectionFromPath = Nothing
End Function

' R23-707. Reads the section-line geometry back from the SOURCE view. Must
' be called after creation and again after every later view move, because a
' moved view re-lays the section line and the previously read geometry stops
' describing the sheet.
Public Function ReadSectionLineInfo( _
    ByRef sourceView As SldWorks.View) As String

    On Error GoTo Failed

    If sourceView Is Nothing Then
        ReadSectionLineInfo = "sectionLine=NoSourceView"
        Exit Function
    End If

    Dim info As Variant
    info = sourceView.GetSectionLineInfo2

    If IsEmpty(info) Or Not IsArray(info) Then
        ReadSectionLineInfo = "sectionLine=NoGeometryReturned"
        Exit Function
    End If

    ReadSectionLineInfo = "sectionLine=Read" & _
        "|values=" & CStr(UBound(info) - LBound(info) + 1) & _
        "|source=IView.GetSectionLineInfo2"
    Exit Function

Failed:
    ReadSectionLineInfo = "sectionLine=Error:" & CStr(Err.Number)
End Function

' R23-814. READ-ONLY. Decodes what SOLIDWORKS actually did with the section
' line, as opposed to what the path asked for.
'
' Every run since the section first existed has read GetSectionLineInfo2 and
' logged only its element count (values=49, r50 and r51 alike). r51 moved
' waypoint 1 by 40 mm, the sheet shows a visibly longer line, and the
' resulting section view was byte-for-byte identical to r50's - same 38
' records, same 1056 doubles, same seven X and seven Y coordinates. Our
' intent has been verified on every run; the drawing's result never has.
'
' Documented layout, SOLIDWORKS 2025 Help, IView::GetSectionLineInfo2:
'   [numSectionLines, layer,
'     per line: numSegments,
'       per segment: lineType, startPt[3], endPt[3],
'       arrowStart1[3], arrowEnd1[3], arrowWidth1, arrowHeight1,
'       arrowStyle1, arrowStart2[3], arrowEnd2[3], arrowWidth2,
'       arrowHeight2, arrowStyle2, textPt1[3], textPt2[3], textHeight]
'
' Three segments under that layout account for 53 doubles, and the live
' array holds 49, so the tail does not match the documentation on this
' build. The segment block sits at the FRONT of the array and is unaffected
' by whatever the tail turns out to be, so the segments are decoded and the
' discrepancy is reported rather than hidden. The raw array is dumped as
' well: 49 doubles is small, and a dump cannot be wrong about the thing a
' structured decode might be wrong about.
Public Sub EmitSectionLineDecode( _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    If swView Is Nothing Then Exit Sub

    Dim viewName As String
    viewName = SafeViewName(swView)

    Dim info As Variant
    Dim readError As Long
    On Error Resume Next
    info = swView.GetSectionLineInfo2
    readError = Err.Number
    Err.Clear
    On Error GoTo Failed

    If IsEmpty(info) Or Not IsArray(info) Then
        EmitWarning evidence, "SECTION_LINE_DECODE|view=" & viewName & _
            "|status=NoArray|error=" & CStr(readError)
        Exit Sub
    End If

    Dim lower As Long
    Dim upper As Long
    lower = LBound(info)
    upper = UBound(info)

    Dim count As Long
    count = upper - lower + 1

    ' Raw first. Six per line keeps each evidence line readable.
    Dim chunk As String
    Dim chunkStart As Long
    Dim i As Long

    chunkStart = 0
    chunk = vbNullString

    For i = 0 To count - 1
        If Len(chunk) > 0 Then chunk = chunk & ";"
        chunk = chunk & Format$(CDbl(info(lower + i)), "0.000000000")

        If ((i + 1) Mod 6 = 0) Or (i = count - 1) Then
            EmitInfo evidence, "SECTION_LINE_RAW|view=" & viewName & _
                "|from=" & CStr(chunkStart) & "|values=" & chunk
            chunkStart = i + 1
            chunk = vbNullString
        End If
    Next i

    If count < 3 Then
        EmitWarning evidence, "SECTION_LINE_DECODE|view=" & viewName & _
            "|status=TooShort|count=" & CStr(count)
        Exit Sub
    End If

    Dim lineCount As Long
    Dim layerValue As Double
    Dim segmentCount As Long
    lineCount = CLng(info(lower))
    layerValue = CDbl(info(lower + 1))
    segmentCount = CLng(info(lower + 2))

    Dim segmentsDecoded As Long
    Dim consumed As Long
    consumed = 3

    If segmentCount > 0 And segmentCount < 64 Then
        For i = 0 To segmentCount - 1
            Dim base As Long
            base = lower + consumed + (i * 7)
            If base + 6 > upper Then Exit For

            EmitInfo evidence, "SECTION_LINE_SEGMENT|view=" & viewName & _
                "|index=" & CStr(i + 1) & _
                "|lineType=" & Format$(CDbl(info(base)), "0.###") & _
                "|start=" & Format$(CDbl(info(base + 1)), "0.000000000") & _
                    "," & Format$(CDbl(info(base + 2)), "0.000000000") & _
                    "," & Format$(CDbl(info(base + 3)), "0.000000000") & _
                "|end=" & Format$(CDbl(info(base + 4)), "0.000000000") & _
                    "," & Format$(CDbl(info(base + 5)), "0.000000000") & _
                    "," & Format$(CDbl(info(base + 6)), "0.000000000") & _
                "|frame=AsReturned"

            segmentsDecoded = segmentsDecoded + 1
        Next i

        consumed = consumed + (segmentsDecoded * 7)
    End If

    ' 29 doubles of arrow and text data are documented after the segments.
    Dim documentedTotal As Long
    documentedTotal = 3 + (segmentCount * 7) + 29

    EmitInfo evidence, "SECTION_LINE_DECODE|view=" & viewName & _
        "|numSectionLines=" & CStr(lineCount) & _
        "|layer=" & Format$(layerValue, "0.###") & _
        "|numSegments=" & CStr(segmentCount) & _
        "|segmentsDecoded=" & CStr(segmentsDecoded) & _
        "|count=" & CStr(count) & _
        "|documentedTotal=" & CStr(documentedTotal) & _
        "|tailMatchesDocumented=" & CStr(documentedTotal = count) & _
        "|error=" & CStr(readError)
    Exit Sub

Failed:
    Dim decodeErrorNumber As Long
    decodeErrorNumber = Err.Number

    EmitWarning evidence, "SECTION_LINE_DECODE|view=" & _
        SafeViewName(swView) & "|status=Error:" & CStr(decodeErrorNumber)
End Sub

' R23-708. Reports why the cut cannot be made rather than approximating it.
Public Function VerifySectionGeometry( _
    ByRef path As CSectionPath) As String

    On Error GoTo Failed

    Dim failures As String

    If path Is Nothing Then
        VerifySectionGeometry = "sectionFailures=NoPath"
        Exit Function
    End If

    If Not path.Resolved Then
        failures = AppendFailure(failures, path.RejectionReason)
    End If

    If path.BoreProjection Is Nothing Then
        failures = AppendFailure(failures, "NoBoreProjection")
    End If

    If path.ColumnProjections.Count = 0 Then
        failures = AppendFailure(failures, "NoColumnHoles")
    End If

    ' "NotAttempted" is the initial STATE of the crossing proof, not a
    ' failure of it. A path rejected before crossings could be tested - no
    ' bore, too few columns - already reported why; appending NotAttempted
    ' alongside those reasons dilutes them, which is the opposite of what
    ' R23-610-style field-specific reporting is for.
    If StrComp(path.CrossingFailures, "None", vbBinaryCompare) <> 0 And _
        StrComp(path.CrossingFailures, "NotAttempted", _
            vbBinaryCompare) <> 0 Then

        failures = AppendFailure(failures, path.CrossingFailures)
    End If

    If Len(failures) = 0 Then failures = "None"

    VerifySectionGeometry = "segments=" & CStr(path.SegmentCount()) & _
        "|columnHoles=" & CStr(path.ColumnProjections.Count) & _
        "|crossingsProven=" & CStr(path.CrossingProofs.Count) & _
        "|sectionFailures=" & failures
    Exit Function

Failed:
    VerifySectionGeometry = "sectionFailures=Error:" & CStr(Err.Number)
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

' R23-700 to R23-708 evidence entry point. STRICTLY READ-ONLY: it contains
' no CreateSectionFromPath call, no CreateLine and no CreateSectionViewAt5.
Public Sub R23_ProbeSectionPath()
    On Error GoTo Failed

    mEmitDiagnostics = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTION_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTION_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTION_FATAL|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTION_FATAL|reason=NoViewsOnSheet"
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
        Module21_EvidenceSink.LogLine _
            "R23_SECTION_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTION_FATAL|reason=UnauthorizedFixture" & _
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

    Module21_EvidenceSink.LogLine _
        "R23_SECTION_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0"

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        Module21_EvidenceSink.LogLine _
            "R23_SECTION_FATAL|reason=CatalogUnavailable"
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

    Module21_EvidenceSink.LogLine _
        "R23_SECTION_GRAPH|" & graph.GraphSummary()

    Dim sectionLabel As String
    sectionLabel = Module1_Main.GetSectionLabelOrDefault(1)

    Dim resolvedPaths As Long
    Dim bestVerdict As String
    bestVerdict = "sectionFailures=NoViewAttempted"

    For i = LBound(views) To UBound(views)
        Dim pathView As SldWorks.View
        Set pathView = views(i)
        If pathView Is Nothing Then GoTo ContinuePathView

        Dim path As CSectionPath
        Set path = ResolveSectionPath( _
            graph, pathView, sectionLabel, evidence)

        If StrComp(path.RejectionReason, "NoProjectionsInView", _
            vbBinaryCompare) = 0 Then

            GoTo ContinuePathView
        End If

        Module21_EvidenceSink.LogLine _
            "QA INFO: SECTION_PATH_CANDIDATE|" & path.Summary()

        Dim verdict As String
        verdict = VerifySectionGeometry(path)

        Module21_EvidenceSink.LogLine _
            "QA INFO: SECTION_GEOMETRY|view=" & _
            path.SourceViewName & "|" & verdict

        If path.Resolved Then
            resolvedPaths = resolvedPaths + 1
            bestVerdict = verdict

            ' The frame conversion is exercised read-only so the transform
            ' is proved before anything is created from it.
            Dim probeViewX As Double
            Dim probeViewY As Double
            Dim frameProof As String

            If PageToViewSketch(path.SourceView, path.W1X, path.W1Y, _
                probeViewX, probeViewY, frameProof) Then

                Module21_EvidenceSink.LogLine _
                    "QA INFO: SECTION_FRAME_PROBE|view=" & _
                    path.SourceViewName & _
                    "|pageX=" & Format$(path.W1X, "0.000000000") & _
                    "|pageY=" & Format$(path.W1Y, "0.000000000") & _
                    "|viewX=" & Format$(probeViewX, "0.000000000") & _
                    "|viewY=" & Format$(probeViewY, "0.000000000") & _
                    "|" & frameProof
            Else
                Module21_EvidenceSink.LogLine _
                    "QA INFO: SECTION_FRAME_PROBE|view=" & _
                    path.SourceViewName & "|" & frameProof
            End If
        End If

ContinuePathView:
    Next i

    Dim finalSelectionCount As Long
    finalSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim drawingSaveAfter As Boolean
    drawingSaveAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Module21_EvidenceSink.LogLine _
        "R23_SECTION_END|resolvedPaths=" & CStr(resolvedPaths) & _
        "|" & bestVerdict & _
        "|creations=0" & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & _
        CStr(drawingSaveBefore = drawingSaveAfter)
    Exit Sub

Failed:
    Module21_EvidenceSink.LogLine _
        "R23_SECTION_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub

Option Explicit

' R23 Phase 7. Section-line path built from model intent.
'
' SAFETY BOUNDARY. Exactly one procedure changes a drawing:
' CreateSectionFromPath, which refuses unless passed an explicit
' allowMutation argument and refuses again unless the path's crossings are
' proved. R23_ProbeSectionPath never passes it and contains no CreateLine or
' CreateSectionViewAt5 call.
'
' R23-700. This module consumes the location graph and never touches the
' ordinate engine: section construction reads proved locations, it does not
' wait on or interact with dimension creation. There is deliberately no
' reference to Module15_OrdinateScheme anywhere below.
'
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
'
' This small slack absorbs the page-coordinate quantum only; it is not a
' search radius. Widening it would let a segment "cross" a hole it misses.
Public Const CROSSING_SLACK_M As Double = 0.000001

' Waypoints are compared for distinctness at the same quantum Phase 5 used
' for ordinate coordinates, so a column and a row mean the same thing in
' both phases.
Public Const PATH_COORDINATE_QUANTUM_M As Double = 0.000001

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

' R23-701 and R23-702. Builds the path for one view.
'
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

    ' R23-701, first half. The stepped-bore centre is the accepted
    ' projection of the largest singleton-family location: singleton because
    ' a bore is not one of a repeated family, largest because the section is
    ' cut through the principal bore. Family size comes from the graph, so
    ' this is not a radius threshold.
    Set path.BoreProjection = ResolveBoreProjection(graph, projections)

    If path.BoreProjection Is Nothing Then
        path.RejectionReason = "NoAcceptedSingletonBoreProjection"
        Exit Function
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

    path.W1X = path.BoreProjection.PageX
    path.W1Y = path.BoreProjection.PageY
    path.W2X = path.BoreProjection.PageX
    path.W2Y = highestRowY
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

' The accepted projection whose location stands alone in its family and has
' the largest radius. Singleton family membership is read from the graph.
Private Function ResolveBoreProjection( _
    ByRef graph As CLocationGraph, _
    ByRef projections As Collection) As CViewHoleProjection

    On Error GoTo Failed

    Dim best As CViewHoleProjection
    Set best = Nothing

    Dim i As Long
    For i = 1 To projections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = projections(i)

        If candidate Is Nothing Then GoTo ContinueCandidate
        If Not candidate.Accepted Then GoTo ContinueCandidate
        If candidate.PhysicalLocation Is Nothing Then GoTo ContinueCandidate

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

    Set ResolveBoreProjection = best
    Exit Function

Failed:
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

    If Not PathCrossesCircle(path, path.BoreProjection) Then
        failures = AppendFailure(failures, "BoreNotCrossed")
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
'
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
'
' Refuses unless allowMutation is True AND the path resolved with its
' crossings proved. Creates exactly three view-owned sketch segments,
' verifies the selection count before calling CreateSectionViewAt5, then
' reads GetSectionLineInfo2 back.
'
' placeX and placeY are the SHEET position for the CENTRE of the new section
' view, per the CreateSectionViewAt5 contract. They are caller arguments and
' deliberately not derived here: choosing where a view sits is layout, which
' is a later phase's responsibility and needs the full annotation envelopes
' this module cannot see. Defaulting them to a point on the source view -
' the obvious shortcut - would stack the section on top of the view it was
' cut from.
'
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

    If Not Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDrawing.ActivateView(path.SourceViewName)) Then

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
    Dim sectionView As SldWorks.View
    Set sectionView = swDrawing.CreateSectionViewAt5( _
        placeX, placeY, 0#, path.SectionLabel, 0, Nothing, 0#)

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

    If StrComp(path.CrossingFailures, "None", vbBinaryCompare) <> 0 Then
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
        Debug.Print "R23_SECTION_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Debug.Print "R23_SECTION_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Debug.Print "R23_SECTION_FATAL|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Debug.Print "R23_SECTION_FATAL|reason=NoViewsOnSheet"
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
        Debug.Print "R23_SECTION_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Debug.Print "R23_SECTION_FATAL|reason=UnauthorizedFixture" & _
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

    Debug.Print "R23_SECTION_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0"

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        Debug.Print "R23_SECTION_FATAL|reason=CatalogUnavailable"
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

    Debug.Print "R23_SECTION_GRAPH|" & graph.GraphSummary()

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

        Debug.Print "QA INFO: SECTION_PATH_CANDIDATE|" & path.Summary()

        Dim verdict As String
        verdict = VerifySectionGeometry(path)

        Debug.Print "QA INFO: SECTION_GEOMETRY|view=" & _
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

                Debug.Print "QA INFO: SECTION_FRAME_PROBE|view=" & _
                    path.SourceViewName & _
                    "|pageX=" & Format$(path.W1X, "0.000000000") & _
                    "|pageY=" & Format$(path.W1Y, "0.000000000") & _
                    "|viewX=" & Format$(probeViewX, "0.000000000") & _
                    "|viewY=" & Format$(probeViewY, "0.000000000") & _
                    "|" & frameProof
            Else
                Debug.Print "QA INFO: SECTION_FRAME_PROBE|view=" & _
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

    Debug.Print "R23_SECTION_END|resolvedPaths=" & CStr(resolvedPaths) & _
        "|" & bestVerdict & _
        "|creations=0" & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & _
        CStr(drawingSaveBefore = drawingSaveAfter)
    Exit Sub

Failed:
    Debug.Print "R23_SECTION_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub

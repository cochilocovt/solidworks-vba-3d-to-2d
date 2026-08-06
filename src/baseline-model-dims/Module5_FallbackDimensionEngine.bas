Option Explicit

' MCP-confirmed 2026-08-05 against swAddOrdinateDims_e and
' swCreateOrdDimError_e. Verify in SW2025 Object Browser before acceptance.
Private Const swHorizontalOrdinate As Long = 3
Private Const swVerticalOrdinate As Long = 2
Private Const swCreateOrdDimErr_Success As Long = 0

' swViewEntityType_e, MCP-confirmed 2026-08-06.
Private Const swViewEntityType_Edge As Long = 1

' swInConfigurationOpts_e, MCP-confirmed 2026-08-06. Used by the dimension
' readback in DescribeDimensionReadback.
Private Const swThisConfiguration As Long = 1

' Outcome codes returned by CreateOneOrdinateChain in addition to the
' swCreateOrdDimError_e values, which are all >= -1.
Private Const CHAIN_SKIPPED_TOO_FEW As Long = -100
Private Const CHAIN_SELECTION_FAILED As Long = -101

' Sentinel for LastFailureCode. 0 is swCreateOrdDimErr_Success, so an
' unset code of 0 reads as a successful chain in the report. The r3 run
' reported "Last code=0 (Unrecognised code)" for exactly that reason.
Public Const ORDINATE_CODE_UNSET As Long = -999

' Drawing-space tolerances, metres.
'
' AXIS_PARALLEL_TOL is how far a line's two endpoints may differ on an axis
' before it stops counting as parallel to the other one. 0.05 mm in view
' space, well below anything a drawing distinguishes.
'
' MIN_EDGE_LENGTH_M rejects slivers: a 1 mm chamfer face edge is axis-parallel
' and useless as an ordinate station.
'
' COORD_DEDUP_TOL_M is the r7 value, retained. It is still an unjustified
' magic number (gap A9) and should be derived from drawing resolution.
Private Const AXIS_PARALLEL_TOL_M As Double = 0.00005
Private Const MIN_EDGE_LENGTH_M As Double = 0.002
Private Const COORD_DEDUP_TOL_M As Double = 0.0015

' Candidate stations for ONE axis.
'
' Gap A2: the r7 engine resolved a single 2-D datum index and used it for both
' chains. X and Y are independent problems with different answers - on the
' reference drawing the short axis is measured from the part centreline and
' the long axis from an end - so they now carry independent candidate sets and
' independent datums.
'
' Gap A3/A4: IsEdgeKind distinguishes a straight model edge from a hole
' centre. Both are admitted. The r7 engine admitted circles only, which is why
' its chain terminated on the outermost hole (23.60) instead of the part
' silhouette (36.00).
' Alt holds a second entity that sits at the same station, kept because a
' circular edge contributes to BOTH axes: the bore's edge is a station on the
' X chain and on the Y chain. The r10 sheet rendered exactly three ordinates
' as 0.00 in the dangling colour while the report listed their true offsets,
' and all three are entities the two chains share. An alternate entity at the
' same coordinate lets the second chain attach to something the first has not
' already consumed, without losing the station.
Private Type AxisCandidates
    Objs() As Object
    Alt() As Object
    Coord() As Double
    IsEdgeKind() As Boolean
    Count As Long
    HoleCount As Long
    EdgeCount As Long
    DatumIndex As Long
    DatumBasis As String
End Type

' Structured ordinate outcome, so a failed chain cannot be masked in QA by a
' healthy imported-model-item count.
Public Type OrdinateRunStatus
    ViewsProcessed As Long
    ViewsWithTooFewCandidates As Long
    ViewsErrored As Long
    ChainsAttempted As Long
    ChainsCreated As Long
    ChainsFailed As Long
    LastFailureCode As Long
    LastFailureView As String
    LastErrorStage As String
    LastErrorNumber As Long
    LastErrorText As String
    EdgeRoute As String
    LastActivateResult As String
    SelDataScope As String
    SelDataViewError As Long
    CandidateEdgesSeen As Long
    CircularEdgesSeen As Long
    LinearEdgesSeen As Long
    ArcEdgesSeen As Long
    AxisRoleBasis As String
    XAxisReport As String
    YAxisReport As String
    HolesOnlyRetries As Long
    SharedEntitySubstitutions As Long
    SharedEntityUnresolved As Long
    ViewScale As Double
    DanglingFound As Long
    DanglingPruned As Long
    PruneSelectFailed As Long
    PruneDeleteFailed As Long
End Type

Public Sub CreateHoleOrdinateDims( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByVal datumType As String, _
    ByRef status As OrdinateRunStatus)

    ' Staged so a thrown error can name where it happened. The r3 run failed
    ' on all four views and the handler recorded nothing but a count, which
    ' proved only that something was wrong.
    Dim stage As String

    On Error GoTo Failed

    status.ViewsProcessed = status.ViewsProcessed + 1

    stage = "BindModel"
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swDraw

    Dim swModelExt As SldWorks.ModelDocExtension
    Set swModelExt = swModel.Extension

    ' Agents.md: establish the correct active view before view-scoped
    ' selection. The baseline never activated, which is the documented
    ' difference from the active-ordinate engine. The r4 run raised error 91
    ' in this block on all six views.
    ' src/active-ordinate clears the selection before activating, and carries
    ' an explicit comment that this ordering was "the error in the prior
    ' rewrite". The baseline cleared nothing here.
    stage = "ClearSelection"
    swModel.ClearSelection2 True

    stage = "ActivateView"
    Dim activated As Boolean
    activated = swDraw.ActivateView(swView.Name)
    status.LastActivateResult = CStr(activated)

    ' Sub-staged so a repeat failure names the exact statement rather than
    ' the block.
    stage = "SelectionManager"
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    If swSelMgr Is Nothing Then
        Err.Raise vbObjectError + 7200, "CreateHoleOrdinateDims", _
            "ModelDoc2.SelectionManager returned Nothing for view " & _
            swView.Name
    End If

    stage = "CreateSelectData"
    Dim swSelData As SldWorks.SelectData
    Set swSelData = swSelMgr.CreateSelectData
    If swSelData Is Nothing Then
        Err.Raise vbObjectError + 7201, "CreateHoleOrdinateDims", _
            "SelectionMgr.CreateSelectData returned Nothing for view " & _
            swView.Name
    End If

    ' Non-fatal. r5 proved this assignment raises 91 with both operands
    ' valid, which killed the whole engine before a single edge was read.
    ' The view is activated by this point, so an unscoped SelectData still
    ' selects into it; scoping is a correctness guard, not a precondition.
    ' Record which route was taken so the report says whether the drawing
    ' was built with scoped or unscoped selection.
    ' ISelectData.View is documented as "Gets or sets the drawing view that
    ' contains the selected object", with no error condition (MCP,
    ' 2026-08-06). Error 91 from "Set swSelData.View = swView", with both
    ' operands proved non-Nothing, is therefore not the API refusing - it is
    ' VBA. A property exposed as propertyput rather than propertyputref takes
    ' a plain assignment; using Set against it fails.
    '
    ' So try the let-assignment first and keep Set as the fallback, recording
    ' which one the build accepts. This matters beyond tidiness: scoping the
    ' SelectData to the view is the pattern
    ' docs/CODESTACK_DRAWING_API_COVERAGE.md rows 17 and 31 prescribe for
    ' selecting entities returned by GetVisibleEntities2, and every run since
    ' r5 has been selecting unscoped instead.
    stage = "SetSelectDataView"
    On Error Resume Next

    swSelData.View = swView
    If Err.Number = 0 Then
        status.SelDataScope = "ScopedToView(Let)"
    Else
        Err.Clear
        Set swSelData.View = swView

        If Err.Number = 0 Then
            status.SelDataScope = "ScopedToView(Set)"
        Else
            status.SelDataViewError = Err.Number
            status.SelDataScope = "Unscoped(err=" & CStr(Err.Number) & ")"
            Err.Clear
        End If
    End If

    On Error GoTo Failed

    stage = "GetOutline"
    Dim viewOutline As Variant
    viewOutline = swView.GetOutline

    ' IView.GetVisibleEntities2 declares LpViewComponent as Component2 (MCP,
    ' 2026-08-05). The baseline passed Nothing. For a part drawing view the
    ' component must come from IView.GetVisibleComponents; passing Nothing is
    ' an unresolved failure mode recorded in the gap analysis. Try the
    ' resolved component first and fall back to Nothing, recording which
    ' route produced entities.
    stage = "ResolveViewComponent"
    Dim swViewComponent As Object
    Set swViewComponent = ResolveFirstVisibleComponent(swView)

    stage = "GetVisibleEntities2"
    Dim edges As Variant
    edges = GetVisibleEdges(swView, swViewComponent, status)

    stage = "CheckEdges"
    If IsEmpty(edges) Then Exit Sub
    If IsNull(edges) Then Exit Sub
    If Not IsArray(edges) Then Exit Sub

    stage = "MathUtility"
    Dim swMathUtil As SldWorks.MathUtility
    Set swMathUtil = Application.SldWorks.GetMathUtility

    Dim swTransform As SldWorks.MathTransform
    Set swTransform = swView.ModelToViewTransform

    ' ModelToViewTransform carries the view scale. r12 ran at 1:2 and every
    ' station offset came out at exactly half its r10 value while the
    ' dimensions SOLIDWORKS rendered stayed true, which proved two defects at
    ' once: the reported offsets were view-space units mislabelled as
    ' millimetres, and COORD_DEDUP_TOL_M was being applied in view space, so
    ' stations merged differently at different scales - the 43 mm and 45 mm
    ' stations collapsed into one at 1:2. Dividing by the scale puts every
    ' coordinate, tolerance and report back into model units.
    stage = "ReadViewScale"
    Dim viewScale As Double
    viewScale = ReadViewScale(swView)
    status.ViewScale = viewScale

    stage = "ScanEdges"
    Dim xs As AxisCandidates
    Dim ys As AxisCandidates
    CollectAxisCandidates edges, swMathUtil, swTransform, viewScale, _
        xs, ys, status

    If xs.Count < 2 And ys.Count < 2 Then
        status.ViewsWithTooFewCandidates = _
            status.ViewsWithTooFewCandidates + 1
        Exit Sub
    End If

    stage = "ResolveDatums"
    ResolveAxisDatums xs, ys, datumType, status

    status.XAxisReport = DescribeAxis("X", xs)
    status.YAxisReport = DescribeAxis("Y", ys)

    stage = "CreateChains"

    ' One entity may belong to only one chain. A circular edge is a station on
    ' both axes, so the X chain would otherwise consume the bore edge as its
    ' datum and leave the Y chain re-selecting the same entity for its 36.00
    ' station - which is what the r10 sheet showed as a dangling 0.00.
    Dim usedObjs() As Object
    Dim usedCount As Long
    usedCount = 0

    RecordChainOutcome status, swView.Name, _
        CreateChainWithFallback(swModel, swModelExt, swView, swSelData, _
            xs, True, viewOutline, usedObjs, usedCount, status)
    RecordChainOutcome status, swView.Name, _
        CreateChainWithFallback(swModel, swModelExt, swView, swSelData, _
            ys, False, viewOutline, usedObjs, usedCount, status)

    ' Pruning does NOT happen here. r14 proved why: this ran immediately after
    ' creation and found zero dangling annotations, while the QA readback -
    ' after the pipeline's ForceRebuild3 - found two in the same view. The
    ' dangling flag is not set until a rebuild, so the prune has to run after
    ' one. Module2_DrawingPipeline calls PruneDanglingAcrossDrawing there.
    Exit Sub

Failed:
    ' A thrown view is not a failed chain. Counting it as one is what made
    ' the r3 report say "0 of 0 attempted" and "4 chains failed" at once.
    status.ViewsErrored = status.ViewsErrored + 1
    status.LastErrorStage = stage
    status.LastErrorNumber = Err.Number
    status.LastErrorText = Err.Description
    status.LastFailureView = swView.Name
    Debug.Print "Fallback ordinate error in view " & swView.Name & _
        " at stage " & stage & ": " & CStr(Err.Number) & " " & Err.Description
End Sub

' Walks the visible edges once and feeds both axis candidate sets.
'
' A circle contributes its centre to BOTH axes: a hole is a station on the X
' chain and on the Y chain. A straight line contributes to ONE axis only - the
' axis it is perpendicular to - because a line parallel to the X axis has no
' single X to dimension.
Private Sub CollectAxisCandidates( _
    ByRef edges As Variant, _
    ByRef swMathUtil As SldWorks.MathUtility, _
    ByRef swTransform As SldWorks.MathTransform, _
    ByVal viewScale As Double, _
    ByRef xs As AxisCandidates, _
    ByRef ys As AxisCandidates, _
    ByRef status As OrdinateRunStatus)

    Dim i As Long
    For i = LBound(edges) To UBound(edges)
        Dim swEdge As SldWorks.Edge
        Set swEdge = edges(i)
        If swEdge Is Nothing Then GoTo NextEdge

        status.CandidateEdgesSeen = status.CandidateEdgesSeen + 1

        Dim swCurve As SldWorks.Curve
        Set swCurve = swEdge.GetCurve
        If swCurve Is Nothing Then GoTo NextEdge

        ' ONLY "= False" is safe against a SOLIDWORKS COM Boolean here, and
        ' "= True" is NOT its mirror image.
        '
        ' Two live findings, both on this build. r6: "If Not <comBool>" yields
        ' -2, which VBA treats as True, so it rejected every edge - 349 edges,
        ' 0 circles on a part with 12 holes. r9: rewriting the fixed
        ' "IsCircle = False" test as "IsCircle = True" took the circular count
        ' from 35 to 0 on the identical 64 edges. VBA True is -1; the value
        ' the call returns is truthy but not -1, so "= True" never matches.
        '
        ' Comparing against False and negating the resulting VBA Boolean is
        ' the only form proved to work. Keep it.
        Dim isCircle As Boolean
        Dim isLine As Boolean
        isCircle = Not (swCurve.IsCircle = False)
        isLine = Not (swCurve.IsLine = False)

        If isCircle Then
            ' ICurve.IsCircle is true for an ARC as well as a full circle -
            ' its Remarks say to use IEdge::GetCurveParams2 to tell them
            ' apart (MCP, 2026-08-06). The trunk did not, so P-0251's rounded
            ' end was treated as a hole and its arc CENTRE became the X datum.
            ' That is the 60 mm offset between this chain and the reference.
            '
            ' An arc is KEPT as a station. r18 excluded them and that was an
            ' over-correction: it fixed the X datum and broke the Y one,
            ' because P-0251's rounded end has its arc centre exactly on the
            ' part's centreline and was the only entity there. The Y chain
            ' went from 36/15/0/15/35 to 21/0/30/50 with the datum 14.5 mm off
            ' centre.
            '
            ' An arc's centre is a true, dimensionable coordinate - it is what
            ' an ordinate attached to that edge reads, and for a rounded end
            ' it is exactly the axis of symmetry. The r17 defect was never
            ' that the centre is wrong; it was that the centre was mistaken
            ' for the part's EXTREME. That is fixed in ResolveOneDatum by
            ' requiring an end datum to be a straight edge, which is
            ' sufficient on its own.
            If IsFullCircle(swEdge) = False Then
                status.ArcEdgesSeen = status.ArcEdgesSeen + 1
            End If

            status.CircularEdgesSeen = status.CircularEdgesSeen + 1

            Dim cp As Variant
            cp = swCurve.CircleParams

            Dim centre As Variant
            centre = ToModelUnitPoint(swMathUtil, swTransform, viewScale, _
                cp(0), cp(1), cp(2))

            AddAxisCandidate xs, swEdge, centre(0), False
            AddAxisCandidate ys, swEdge, centre(1), False
            GoTo NextEdge
        End If

        If isLine Then
            Dim p1 As Variant
            Dim p2 As Variant
            If EdgeEndpoints(swEdge, p1, p2) = False Then GoTo NextEdge

            Dim v1 As Variant
            Dim v2 As Variant
            v1 = ToModelUnitPoint(swMathUtil, swTransform, viewScale, _
                p1(0), p1(1), p1(2))
            v2 = ToModelUnitPoint(swMathUtil, swTransform, viewScale, _
                p2(0), p2(1), p2(2))

            Dim dx As Double
            Dim dy As Double
            dx = Abs(v1(0) - v2(0))
            dy = Abs(v1(1) - v2(1))

            status.LinearEdgesSeen = status.LinearEdgesSeen + 1

            If dx < AXIS_PARALLEL_TOL_M And dy >= MIN_EDGE_LENGTH_M Then
                ' Runs vertically in the view: one constant X.
                AddAxisCandidate xs, swEdge, v1(0), True
            ElseIf dy < AXIS_PARALLEL_TOL_M And dx >= MIN_EDGE_LENGTH_M Then
                ' Runs horizontally in the view: one constant Y.
                AddAxisCandidate ys, swEdge, v1(1), True
            End If
        End If
NextEdge:
    Next i
End Sub

' Adds one station, deduplicating in ONE dimension.
'
' Gap A8: the r7 engine deduplicated with a 2-D proximity test, which is the
' wrong question for a 1-D chain. Two holes sharing an X but 40 mm apart in Y
' are two points to a 2-D test and one station to the X chain.
Private Sub AddAxisCandidate( _
    ByRef axis As AxisCandidates, _
    ByRef obj As Object, _
    ByVal coord As Double, _
    ByVal isEdgeKind As Boolean)

    Dim i As Long
    For i = 0 To axis.Count - 1
        If Abs(axis.Coord(i) - coord) < COORD_DEDUP_TOL_M Then
            ' A straight edge at the same station displaces a hole. The
            ' reference drawing's chains terminate on the silhouette, so when
            ' both exist the edge is the one worth keeping. The displaced
            ' entity is retained as the alternate rather than dropped.
            If isEdgeKind = True And axis.IsEdgeKind(i) = False Then
                If axis.Alt(i) Is Nothing Then Set axis.Alt(i) = axis.Objs(i)
                Set axis.Objs(i) = obj
                axis.IsEdgeKind(i) = True
                axis.HoleCount = axis.HoleCount - 1
                axis.EdgeCount = axis.EdgeCount + 1
            ElseIf axis.Alt(i) Is Nothing Then
                If Not obj Is axis.Objs(i) Then Set axis.Alt(i) = obj
            End If
            Exit Sub
        End If
    Next i

    ReDim Preserve axis.Objs(0 To axis.Count)
    ReDim Preserve axis.Alt(0 To axis.Count)
    ReDim Preserve axis.Coord(0 To axis.Count)
    ReDim Preserve axis.IsEdgeKind(0 To axis.Count)

    Set axis.Objs(axis.Count) = obj
    Set axis.Alt(axis.Count) = Nothing
    axis.Coord(axis.Count) = coord
    axis.IsEdgeKind(axis.Count) = isEdgeKind
    axis.Count = axis.Count + 1

    If isEdgeKind = True Then
        axis.EdgeCount = axis.EdgeCount + 1
    Else
        axis.HoleCount = axis.HoleCount + 1
    End If
End Sub

' Per-axis datum resolution (gap A2).
'
' The reference drawing measures its LONG axis from one end (0, 10, 50, 90,
' 160) and its SHORT axis from the part centreline (36, 15, 0, 15, 36). That
' asymmetry is the contract implemented here for the default "Center" origin:
' the axis with the larger candidate span is measured from its minimum
' station, the other from its midpoint.
'
' Whether front-view ordinates measured this way are a general rule or an
' instance of one drawing is open question 2 in
' docs/BASELINE_TO_REFERENCE_DRAWING_GAP.md. The basis is recorded in the QA
' report on every run so the choice is visible rather than implicit.
'
' "Top-Left" and "Bottom-Left" keep their corner meaning on both axes.
Private Sub ResolveAxisDatums( _
    ByRef xs As AxisCandidates, _
    ByRef ys As AxisCandidates, _
    ByVal datumType As String, _
    ByRef status As OrdinateRunStatus)

    Dim xMin As Double, xMax As Double
    Dim yMin As Double, yMax As Double
    AxisRange xs, xMin, xMax
    AxisRange ys, yMin, yMax

    Dim xTarget As Double
    Dim yTarget As Double
    Dim xFromExtreme As Boolean
    Dim yFromExtreme As Boolean
    Dim basis As String

    If datumType = "Top-Left" Then
        xTarget = xMin
        yTarget = yMax
        xFromExtreme = True
        yFromExtreme = True
        basis = "CornerTopLeft"
    ElseIf datumType = "Bottom-Left" Then
        xTarget = xMin
        yTarget = yMin
        xFromExtreme = True
        yFromExtreme = True
        basis = "CornerBottomLeft"
    ElseIf (xMax - xMin) >= (yMax - yMin) Then
        xTarget = xMin
        yTarget = (yMin + yMax) / 2#
        xFromExtreme = True
        yFromExtreme = False
        basis = "LongAxis=X:FromMinEnd;ShortAxis=Y:FromCentreline"
    Else
        xTarget = (xMin + xMax) / 2#
        yTarget = yMin
        xFromExtreme = False
        yFromExtreme = True
        basis = "LongAxis=Y:FromMinEnd;ShortAxis=X:FromCentreline"
    End If

    status.AxisRoleBasis = basis

    ' An "from the end" datum is a datum FACE - a straight edge at an extreme
    ' of the part - not whichever feature happens to sit nearest the extreme.
    ' On P-0251 the nearest thing to xMin was the rounded end's arc centre,
    ' which is why the long chain started 60 mm inside the part. A centreline
    ' datum takes no such restriction: it is normally a hole or an axis.
    ResolveOneDatum xs, xTarget, xFromExtreme
    ResolveOneDatum ys, yTarget, yFromExtreme
End Sub

Private Sub AxisRange( _
    ByRef axis As AxisCandidates, _
    ByRef minCoord As Double, _
    ByRef maxCoord As Double)

    minCoord = 0#
    maxCoord = 0#
    If axis.Count < 1 Then Exit Sub

    minCoord = axis.Coord(0)
    maxCoord = axis.Coord(0)

    Dim i As Long
    For i = 1 To axis.Count - 1
        If axis.Coord(i) < minCoord Then minCoord = axis.Coord(i)
        If axis.Coord(i) > maxCoord Then maxCoord = axis.Coord(i)
    Next i
End Sub

' Picks the station nearest the target. A straight edge wins a tie against a
' hole, because gap A3 records that neither datum on the reference drawing is
' a hole. Whatever is chosen, its kind and its distance from the ideal target
' go into the report - a datum 12 mm from the centreline is a finding, not a
' detail.
Private Sub ResolveOneDatum( _
    ByRef axis As AxisCandidates, _
    ByVal target As Double, _
    ByVal requireStraightEdge As Boolean)

    axis.DatumIndex = 0
    axis.DatumBasis = "NoCandidates"
    If axis.Count < 1 Then Exit Sub

    ' A datum face is a straight edge. Restrict to those when the datum is an
    ' end datum and any exist; fall back to the whole set otherwise, so a part
    ' with no straight edge on that axis still gets a chain.
    Dim edgesOnly As Boolean
    edgesOnly = (requireStraightEdge = True) And (axis.EdgeCount > 0)

    Dim bestIdx As Long
    Dim bestDist As Double
    bestIdx = -1
    bestDist = 0#

    Dim i As Long
    For i = 0 To axis.Count - 1
        If edgesOnly = False Or axis.IsEdgeKind(i) = True Then
            Dim dist As Double
            dist = Abs(axis.Coord(i) - target)

            If bestIdx < 0 Then
                bestIdx = i
                bestDist = dist
            ElseIf dist < bestDist - AXIS_PARALLEL_TOL_M Then
                bestDist = dist
                bestIdx = i
            ElseIf Abs(dist - bestDist) <= AXIS_PARALLEL_TOL_M Then
                If axis.IsEdgeKind(i) = True And _
                    axis.IsEdgeKind(bestIdx) = False Then
                    bestDist = dist
                    bestIdx = i
                End If
            End If
        End If
    Next i

    If bestIdx < 0 Then bestIdx = 0
    axis.DatumIndex = bestIdx

    Dim kind As String
    If axis.IsEdgeKind(bestIdx) = True Then
        kind = "Edge"
    Else
        kind = "Hole"
    End If

    axis.DatumBasis = kind & ":offsetFromTarget=" & _
        Format$(bestDist * 1000#, "0.00") & "mm"
End Sub

Private Function DescribeAxis( _
    ByVal axisName As String, _
    ByRef axis As AxisCandidates) As String

    DescribeAxis = axisName & ": " & axis.Count & " stations (holes=" & _
        axis.HoleCount & ", edges=" & axis.EdgeCount & "), datum " & _
        axis.DatumBasis

    If axis.Count < 1 Then Exit Function

    DescribeAxis = DescribeAxis & vbCrLf & "    offsets(mm): " & _
        DescribeAxisOffsets(axis)
End Function

' The ordinate values the chain should render, in millimetres from the datum.
' The r8 sheet showed a horizontal chain reading 0 / 0.00 / 70.00 / 110.00 /
' 0.00 / 150.00 and nothing in the report could confirm or deny it. Printing
' the intended stations makes the next sheet checkable against the code
' instead of against a screenshot.
Private Function DescribeAxisOffsets(ByRef axis As AxisCandidates) As String
    Dim sorted() As Double
    ReDim sorted(0 To axis.Count - 1)

    Dim datumCoord As Double
    datumCoord = axis.Coord(axis.DatumIndex)

    Dim i As Long
    For i = 0 To axis.Count - 1
        sorted(i) = (axis.Coord(i) - datumCoord) * 1000#
    Next i

    Dim j As Long
    For i = 1 To axis.Count - 1
        Dim key As Double
        key = sorted(i)
        j = i - 1
        Do While j >= 0
            If sorted(j) <= key Then Exit Do
            sorted(j + 1) = sorted(j)
            j = j - 1
        Loop
        sorted(j + 1) = key
    Next i

    For i = 0 To axis.Count - 1
        If i > 0 Then DescribeAxisOffsets = DescribeAxisOffsets & ", "
        DescribeAxisOffsets = DescribeAxisOffsets & _
            Format$(sorted(i), "0.00")
    Next i
End Function

' IView.GetVisibleComponents returns the drawing-context components for the
' view. A part drawing view has exactly one. Returns Nothing when the call
' is unavailable or empty, so the caller can fall back.
Private Function ResolveFirstVisibleComponent( _
    ByRef swView As SldWorks.View) As Object

    On Error GoTo Failed

    Dim components As Variant
    components = swView.GetVisibleComponents

    If IsEmpty(components) Then Exit Function
    If IsNull(components) Then Exit Function
    If Not IsArray(components) Then Exit Function
    If UBound(components) < LBound(components) Then Exit Function

    Set ResolveFirstVisibleComponent = components(LBound(components))
    Exit Function

Failed:
    Set ResolveFirstVisibleComponent = Nothing
End Function

' Asks for visible edges with the resolved component first, then with
' Nothing. Records which route answered so the QA report can say whether the
' component argument was the problem.
Private Function GetVisibleEdges( _
    ByRef swView As SldWorks.View, _
    ByRef swViewComponent As Object, _
    ByRef status As OrdinateRunStatus) As Variant

    Dim result As Variant

    If Not swViewComponent Is Nothing Then
        On Error Resume Next
        result = swView.GetVisibleEntities2(swViewComponent, _
            swViewEntityType_Edge)
        On Error GoTo 0

        If IsArray(result) Then
            status.EdgeRoute = "ViaComponent"
            GetVisibleEdges = result
            Exit Function
        End If
    End If

    On Error Resume Next
    result = swView.GetVisibleEntities2(Nothing, swViewEntityType_Edge)
    On Error GoTo 0

    If IsArray(result) Then
        status.EdgeRoute = "ViaNothing"
    ElseIf swViewComponent Is Nothing Then
        status.EdgeRoute = "NoComponentAndNothingFailed"
    Else
        status.EdgeRoute = "BothRoutesFailed"
    End If

    GetVisibleEdges = result
End Function

' Transforms a model point into the view frame and removes the view scale, so
' every station, tolerance and reported offset is in model units regardless of
' the scale the operator picked. Only differences between these coordinates
' are ever used, so dividing the translation component by the same factor is
' harmless.
Private Function ToModelUnitPoint( _
    ByRef swMathUtil As SldWorks.MathUtility, _
    ByRef swTransform As SldWorks.MathTransform, _
    ByVal viewScale As Double, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal z As Double) As Variant

    Dim raw(2) As Double
    raw(0) = x
    raw(1) = y
    raw(2) = z

    Dim swPt As SldWorks.MathPoint
    Set swPt = swMathUtil.CreatePoint(raw)
    Set swPt = swPt.MultiplyTransform(swTransform)

    Dim arr As Variant
    arr = swPt.ArrayData

    Dim scaled(2) As Double
    scaled(0) = arr(0) / viewScale
    scaled(1) = arr(1) / viewScale
    scaled(2) = arr(2) / viewScale

    ToModelUnitPoint = scaled
End Function

' IView.ScaleDecimal. Falls back to 1 rather than 0: a zero divisor would make
' every station infinite, and an unscaled reading is a visible wrong answer
' instead of a crash.
Private Function ReadViewScale(ByRef swView As SldWorks.View) As Double
    On Error GoTo Failed

    ReadViewScale = swView.ScaleDecimal
    If ReadViewScale <= 0 Then ReadViewScale = 1#
    Exit Function

Failed:
    ReadViewScale = 1#
End Function

' IEdge.GetStartVertex returns null for an edge with no vertex (the Help gives
' a newly created cylinder as the case), so both are checked. Direction is not
' needed - the caller compares the two transformed endpoints - which avoids
' transforming a direction vector through IView.ModelToViewTransform, a
' contract this project has not established.
' True when a circular edge closes on itself, false when it is an arc.
'
' IEdge.GetStartVertex Remarks: it and GetEndVertex "return distinct vertices,
' unless the edge is closed", and return null when the edge has no vertex at
' all. Both of those are the closed case; two distinct vertices is an arc.
' Fails closed to True, because misreading an arc as a hole is the defect this
' exists to prevent and misreading a hole as an arc only loses a station.
Private Function IsFullCircle(ByRef swEdge As SldWorks.Edge) As Boolean
    On Error GoTo Failed

    Dim v1 As Object
    Dim v2 As Object
    Set v1 = swEdge.GetStartVertex
    Set v2 = swEdge.GetEndVertex

    If v1 Is Nothing Then
        IsFullCircle = True
        Exit Function
    End If
    If v2 Is Nothing Then
        IsFullCircle = True
        Exit Function
    End If

    Dim p1 As Variant
    Dim p2 As Variant
    p1 = v1.GetPoint
    p2 = v2.GetPoint

    If Not IsArray(p1) Then GoTo Failed
    If Not IsArray(p2) Then GoTo Failed

    IsFullCircle = _
        (Abs(p1(0) - p2(0)) < AXIS_PARALLEL_TOL_M) And _
        (Abs(p1(1) - p2(1)) < AXIS_PARALLEL_TOL_M) And _
        (Abs(p1(2) - p2(2)) < AXIS_PARALLEL_TOL_M)
    Exit Function

Failed:
    IsFullCircle = True
End Function

Private Function EdgeEndpoints( _
    ByRef swEdge As SldWorks.Edge, _
    ByRef p1 As Variant, _
    ByRef p2 As Variant) As Boolean

    On Error GoTo Failed

    EdgeEndpoints = False

    Dim v1 As Object
    Dim v2 As Object
    Set v1 = swEdge.GetStartVertex
    Set v2 = swEdge.GetEndVertex

    If v1 Is Nothing Then Exit Function
    If v2 Is Nothing Then Exit Function

    p1 = v1.GetPoint
    p2 = v2.GetPoint

    If Not IsArray(p1) Then Exit Function
    If Not IsArray(p2) Then Exit Function

    EdgeEndpoints = True
    Exit Function

Failed:
    EdgeEndpoints = False
End Function

' Creates one chain, and on failure retries once with holes only.
'
' Admitting straight edges is new at r9. Whether AddOrdinateDimension accepts
' a linear edge as its base entity is NOT established on this build - the Help
' says only "select the base entity to act as the datum point". If the richer
' candidate set is rejected, the run falls back to the r8 behaviour rather
' than losing the chain, and the retry is counted so the report says which
' route produced the drawing.
Private Function CreateChainWithFallback( _
    ByRef swModel As SldWorks.ModelDoc2, _
    ByRef swModelExt As SldWorks.ModelDocExtension, _
    ByRef swView As SldWorks.View, _
    ByRef swSelData As SldWorks.SelectData, _
    ByRef axis As AxisCandidates, _
    ByVal isHorizontal As Boolean, _
    ByVal viewOutline As Variant, _
    ByRef usedObjs() As Object, _
    ByRef usedCount As Long, _
    ByRef status As OrdinateRunStatus) As Long

    ' A failed attempt must not leave its entities claimed, or the retry
    ' substitutes alternates for stations that were never actually drawn.
    Dim usedBefore As Long
    usedBefore = usedCount

    Dim rc As Long
    rc = CreateOneOrdinateChain(swModel, swModelExt, swView, swSelData, _
        axis, isHorizontal, viewOutline, usedObjs, usedCount, status)

    If rc = swCreateOrdDimErr_Success Then
        CreateChainWithFallback = rc
        Exit Function
    End If

    If rc = CHAIN_SKIPPED_TOO_FEW Then
        CreateChainWithFallback = rc
        Exit Function
    End If

    If axis.EdgeCount < 1 Then
        CreateChainWithFallback = rc
        Exit Function
    End If

    Dim holesOnly As AxisCandidates
    holesOnly = HolesOnlySubset(axis)
    If holesOnly.Count < 2 Then
        CreateChainWithFallback = rc
        Exit Function
    End If

    ' Re-resolve the datum inside the reduced set: the edge that was the datum
    ' is no longer present.
    ' False: the holes-only set has no straight edges left to require.
    ResolveOneDatum holesOnly, axis.Coord(axis.DatumIndex), False

    status.HolesOnlyRetries = status.HolesOnlyRetries + 1
    usedCount = usedBefore

    Dim retryRc As Long
    retryRc = CreateOneOrdinateChain(swModel, swModelExt, swView, swSelData, _
        holesOnly, isHorizontal, viewOutline, usedObjs, usedCount, status)

    If isHorizontal Then
        status.XAxisReport = status.XAxisReport & vbCrLf & _
            "    RETRY holes-only after code " & rc & ": code " & retryRc
    Else
        status.YAxisReport = status.YAxisReport & vbCrLf & _
            "    RETRY holes-only after code " & rc & ": code " & retryRc
    End If

    CreateChainWithFallback = retryRc
End Function

Private Function HolesOnlySubset( _
    ByRef axis As AxisCandidates) As AxisCandidates

    Dim result As AxisCandidates

    Dim i As Long
    For i = 0 To axis.Count - 1
        If axis.IsEdgeKind(i) = False Then
            AddAxisCandidate result, axis.Objs(i), axis.Coord(i), False
            ' Same coordinate, so this lands as the station's alternate.
            If Not axis.Alt(i) Is Nothing Then
                AddAxisCandidate result, axis.Alt(i), axis.Coord(i), False
            End If
        End If
    Next i

    HolesOnlySubset = result
End Function

' Hands back an entity for this station that no earlier chain has consumed.
'
' The r10 sheet rendered three ordinates as 0.00 in the dangling colour while
' the report listed their true offsets, so creation placed them correctly and
' the attachment failed. All three are stations whose entity the other chain
' had already taken - a circular edge is a station on both axes, and the bore
' edge served as the X datum and as the Y 36.00 station.
'
' This is a hypothesis with a measurement attached, not a certainty: the run
' reports SharedEntitySubstitutions and SharedEntityUnresolved, so the next
' sheet either loses the dangling values with a non-zero substitution count -
' confirming it - or keeps them, refuting it and pointing elsewhere.
Private Function ClaimEntity( _
    ByRef axis As AxisCandidates, _
    ByVal index As Long, _
    ByRef usedObjs() As Object, _
    ByRef usedCount As Long, _
    ByRef status As OrdinateRunStatus) As Object

    Dim primary As Object
    Set primary = axis.Objs(index)

    If IsEntityUsed(usedObjs, usedCount, primary) = False Then
        MarkEntityUsed usedObjs, usedCount, primary
        Set ClaimEntity = primary
        Exit Function
    End If

    Dim alternate As Object
    Set alternate = axis.Alt(index)

    If Not alternate Is Nothing Then
        If IsEntityUsed(usedObjs, usedCount, alternate) = False Then
            MarkEntityUsed usedObjs, usedCount, alternate
            status.SharedEntitySubstitutions = _
                status.SharedEntitySubstitutions + 1
            Set ClaimEntity = alternate
            Exit Function
        End If
    End If

    ' Nothing distinct is available at this station. Keep the station rather
    ' than dropping it, and record that it is still shared.
    status.SharedEntityUnresolved = status.SharedEntityUnresolved + 1
    Set ClaimEntity = primary
End Function

Private Function IsEntityUsed( _
    ByRef usedObjs() As Object, _
    ByVal usedCount As Long, _
    ByRef obj As Object) As Boolean

    IsEntityUsed = False
    If obj Is Nothing Then Exit Function

    Dim i As Long
    For i = 0 To usedCount - 1
        If usedObjs(i) Is obj Then
            IsEntityUsed = True
            Exit Function
        End If
    Next i
End Function

Private Sub MarkEntityUsed( _
    ByRef usedObjs() As Object, _
    ByRef usedCount As Long, _
    ByRef obj As Object)

    If obj Is Nothing Then Exit Sub

    ReDim Preserve usedObjs(0 To usedCount)
    Set usedObjs(usedCount) = obj
    usedCount = usedCount + 1
End Sub

Private Sub RecordChainOutcome( _
    ByRef status As OrdinateRunStatus, _
    ByVal viewName As String, _
    ByVal outcome As Long)

    If outcome = CHAIN_SKIPPED_TOO_FEW Then Exit Sub

    status.ChainsAttempted = status.ChainsAttempted + 1

    If outcome = swCreateOrdDimErr_Success Then
        status.ChainsCreated = status.ChainsCreated + 1
    Else
        status.ChainsFailed = status.ChainsFailed + 1
        status.LastFailureCode = outcome
        status.LastFailureView = viewName
    End If
End Sub

' Renders an OrdinateRunStatus for the QA report. Kept here so the outcome
' vocabulary lives beside the code that produces it.
Public Function DescribeOrdinateStatus( _
    ByRef status As OrdinateRunStatus) As String

    Dim text As String
    text = "Ordinate views processed: " & status.ViewsProcessed & vbCrLf
    text = text & "Ordinate edges seen: " & status.CandidateEdgesSeen & _
        " (circular: " & status.CircularEdgesSeen & _
        ", of which arcs: " & status.ArcEdgesSeen & _
        ", linear: " & status.LinearEdgesSeen & ")" & vbCrLf

    If Len(status.EdgeRoute) > 0 Then
        text = text & "Ordinate edge route: " & status.EdgeRoute & vbCrLf
    End If

    If Len(status.LastActivateResult) > 0 Then
        text = text & "Last ActivateView result: " & _
            status.LastActivateResult & vbCrLf
    End If

    If Len(status.SelDataScope) > 0 Then
        text = text & "Selection scope: " & status.SelDataScope & vbCrLf
    End If

    If status.ViewScale > 0 Then
        text = text & "View scale: " & Format$(status.ViewScale, "0.####") & _
            " (stations below are model mm, scale-normalised)" & vbCrLf
    End If

    If Len(status.AxisRoleBasis) > 0 Then
        text = text & "Datum contract: " & status.AxisRoleBasis & vbCrLf
    End If

    If Len(status.XAxisReport) > 0 Then
        text = text & "  " & status.XAxisReport & vbCrLf
    End If

    If Len(status.YAxisReport) > 0 Then
        text = text & "  " & status.YAxisReport & vbCrLf
    End If

    text = text & "Ordinate chains created: " & status.ChainsCreated & _
        " of " & status.ChainsAttempted & " attempted" & vbCrLf

    text = text & "Shared entities: " & status.SharedEntitySubstitutions & _
        " substituted, " & status.SharedEntityUnresolved & _
        " still shared" & vbCrLf

    If status.DanglingFound > 0 Then
        text = text & "Dangling ordinates: " & status.DanglingFound & _
            " found, " & status.DanglingPruned & " deleted (" & _
            status.PruneSelectFailed & " select-refused, " & _
            status.PruneDeleteFailed & " delete-refused)" & vbCrLf
    End If

    If status.DanglingPruned > 0 Then
        text = text & "WARNING: " & status.DanglingPruned & _
            " dangling ordinate(s) deleted. Those stations are NOT on the " & _
            "drawing; root cause is still open." & vbCrLf
    End If

    If status.SharedEntityUnresolved > 0 Then
        text = text & "WARNING: " & status.SharedEntityUnresolved & _
            " station(s) reuse an entity another chain already took. " & _
            "Expect that many dangling 0.00 ordinates on the sheet." & vbCrLf
    End If

    If status.HolesOnlyRetries > 0 Then
        text = text & "NOTE: " & status.HolesOnlyRetries & _
            " chain(s) fell back to holes-only candidates. A linear edge " & _
            "was rejected as an ordinate entity on this build." & vbCrLf
    End If

    If status.ViewsWithTooFewCandidates > 0 Then
        text = text & "NOTE: " & status.ViewsWithTooFewCandidates & _
            " view(s) had fewer than two ordinate candidates." & vbCrLf
    End If

    ' A thrown view and a failed chain are different faults with different
    ' fixes; report them separately.
    If status.ViewsErrored > 0 Then
        text = text & "WARNING: " & status.ViewsErrored & _
            " view(s) raised a VBA error. Last: stage=" & _
            status.LastErrorStage & " err=" & status.LastErrorNumber & _
            " """ & status.LastErrorText & """ in view " & _
            status.LastFailureView & "." & vbCrLf
    End If

    If status.ChainsFailed > 0 Then
        text = text & "WARNING: " & status.ChainsFailed & _
            " ordinate chain(s) failed. Last code=" & _
            status.LastFailureCode & " (" & _
            DescribeChainFailure(status.LastFailureCode) & ") in view " & _
            status.LastFailureView & "." & vbCrLf
    End If

    DescribeOrdinateStatus = text
End Function

' Removes ordinates whose attachment did not survive creation.
'
' A dangling ordinate renders 0.00 at a geometrically correct position, which
' on a manufacturing drawing is worse than a missing station: it reads as a
' real dimension of zero. The root cause for the two X stations on P-0251 is
' still unknown, so this does not claim to fix it - it removes the wrong value
' and counts what it removed, so the loss stays visible in the report rather
' than being papered over.
' Must be called after a rebuild. See PruneDanglingOrdinates.
Public Sub PruneDanglingAcrossDrawing( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef status As OrdinateRunStatus)

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        PruneDanglingOrdinates swDrawModel, swView, status
        Set swView = swView.GetNextView
    Loop
End Sub

Private Sub PruneDanglingOrdinates( _
    ByRef swModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef status As OrdinateRunStatus)

    On Error GoTo Failed

    Dim vDims As Variant
    vDims = swView.GetDisplayDimensions
    If Not IsArray(vDims) Then Exit Sub

    ' Collect first, delete after. Deleting while walking the array
    ' SOLIDWORKS returned is what invalidates the rest of it.
    Dim doomed() As Object
    Dim doomedCount As Long
    doomedCount = 0

    Dim i As Long
    For i = LBound(vDims) To UBound(vDims)
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = vDims(i)
        If swDispDim Is Nothing Then GoTo NextDim

        Dim swAnn As SldWorks.Annotation
        Set swAnn = swDispDim.GetAnnotation
        If swAnn Is Nothing Then GoTo NextDim

        ' COM Boolean: compare against False only.
        If Not (swAnn.IsDangling = False) Then
            ReDim Preserve doomed(0 To doomedCount)
            Set doomed(doomedCount) = swAnn
            doomedCount = doomedCount + 1
        End If
NextDim:
    Next i

    status.DanglingFound = status.DanglingFound + doomedCount
    If doomedCount < 1 Then Exit Sub

    Dim swModelExt As SldWorks.ModelDocExtension
    Set swModelExt = swModel.Extension

    ' r13 pruned nothing and could not say which call declined. Annotation
    ' selection is view-scoped, and the ordinate work leaves the pick mode and
    ' the active view wherever the last chain left them, so re-establish the
    ' view first and record each failure separately.
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    swModel.SetPickMode
    swModel.ClearSelection2 True
    swDraw.ActivateView swView.Name

    For i = 0 To doomedCount - 1
        Dim swDoomed As SldWorks.Annotation
        Set swDoomed = doomed(i)

        swModel.ClearSelection2 True
        If Not (swDoomed.Select3(False, Nothing) = False) Then
            If Not (swModelExt.DeleteSelection2(0) = False) Then
                status.DanglingPruned = status.DanglingPruned + 1
            Else
                status.PruneDeleteFailed = status.PruneDeleteFailed + 1
            End If
        Else
            status.PruneSelectFailed = status.PruneSelectFailed + 1
        End If
    Next i

    swModel.ClearSelection2 True
    Exit Sub

Failed:
    swModel.ClearSelection2 True
End Sub

' Reads back what was actually created, rather than what was intended.
'
' Two diagnoses of the r10 dangling ordinates were inferred from station
' counts and sheet positions, and both were wrong: r11 substituted three
' shared entities and the two dangling values at X=55.00 and X=135.00 survived
' unchanged. Inference has now failed twice, so this measures instead.
'
' IAnnotation.IsDangling (MCP, 2026-08-06) reports the state directly, and
' IDimension.GetValue3 gives the value the sheet renders. Together they name
' which stations failed without needing a screenshot.
Public Function DescribeDimensionReadback( _
    ByRef swView As SldWorks.View) As String

    On Error GoTo Failed

    Dim vDims As Variant
    vDims = swView.GetDisplayDimensions
    If Not IsArray(vDims) Then Exit Function

    Dim total As Long
    Dim danglingCount As Long
    Dim danglingDetail As String
    Dim liveDetail As String

    Dim i As Long
    For i = LBound(vDims) To UBound(vDims)
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = vDims(i)
        If swDispDim Is Nothing Then GoTo NextDim

        total = total + 1

        Dim shownValue As String
        shownValue = ReadDimensionValue(swDispDim)

        ' COM Boolean: only "= False" is reliable on this build, so the flag
        ' is read by comparing against False and negating a real VBA Boolean.
        Dim isDangling As Boolean
        isDangling = False

        Dim swAnn As SldWorks.Annotation
        Set swAnn = swDispDim.GetAnnotation
        If Not swAnn Is Nothing Then
            isDangling = Not (swAnn.IsDangling = False)
        End If

        If isDangling Then
            danglingCount = danglingCount + 1
            If Len(danglingDetail) > 0 Then _
                danglingDetail = danglingDetail & ", "
            danglingDetail = danglingDetail & shownValue
        Else
            If Len(liveDetail) > 0 Then liveDetail = liveDetail & ", "
            liveDetail = liveDetail & shownValue
        End If
NextDim:
    Next i

    DescribeDimensionReadback = "readback: " & total & " dims, " & _
        danglingCount & " dangling"

    If danglingCount > 0 Then
        DescribeDimensionReadback = DescribeDimensionReadback & _
            vbCrLf & "      DANGLING values: " & danglingDetail
    End If

    If Len(liveDetail) > 0 Then
        DescribeDimensionReadback = DescribeDimensionReadback & _
            vbCrLf & "      live values: " & liveDetail
    End If

    Exit Function

Failed:
    DescribeDimensionReadback = "readback unavailable (err=" & _
        CStr(Err.Number) & ")"
End Function

Private Function ReadDimensionValue( _
    ByRef swDispDim As SldWorks.DisplayDimension) As String

    On Error GoTo Failed

    ReadDimensionValue = "?"

    Dim swDim As SldWorks.Dimension
    Set swDim = swDispDim.GetDimension2(0)
    If swDim Is Nothing Then Exit Function

    Dim vals As Variant
    vals = swDim.GetValue3(swThisConfiguration, Nothing)
    If Not IsArray(vals) Then Exit Function

    ReadDimensionValue = Format$(vals(LBound(vals)), "0.00")
    Exit Function

Failed:
    ReadDimensionValue = "err" & CStr(Err.Number)
End Function

' swCreateOrdDimError_e members, MCP-confirmed 2026-08-05.
Private Function DescribeChainFailure(ByVal code As Long) As String
    Select Case code
        Case ORDINATE_CODE_UNSET
            DescribeChainFailure = "no code recorded"
        Case CHAIN_SELECTION_FAILED
            DescribeChainFailure = "MultiSelect2 selected the wrong count"
        Case -1
            DescribeChainFailure = "swCreateOrdDimErr_Undefined"
        Case 1
            DescribeChainFailure = "swCreateOrdDimErr_OrdFailure"
        Case 2
            DescribeChainFailure = "swCreateOrdDimErr_GenNoInternalDims"
        Case 3
            DescribeChainFailure = "swCreateOrdDimErr_GenBadSel"
        Case 4
            DescribeChainFailure = "swCreateOrdDimErr_GenNeedModelLoaded"
        Case 5
            DescribeChainFailure = "swCreateOrdDimErr_GenSamePartOnly"
        Case 6
            DescribeChainFailure = "swCreateOrdDimErr_GenExtraSelection"
        Case 7
            DescribeChainFailure = "swCreateOrdDimErr_GenFailure"
        Case 8
            DescribeChainFailure = "swCreateOrdDimErr_OrdDupInGroup"
        Case 9
            DescribeChainFailure = "swCreateOrdDimErr_OrdBadDir"
        Case Else
            DescribeChainFailure = "Unrecognised code"
    End Select
End Function

Public Sub InsertHoleCalloutsForView(ByRef swDraw As SldWorks.DrawingDoc, ByRef swView As SldWorks.View)
    On Error GoTo Failed

    Dim edges As Variant
    edges = swView.GetVisibleEntities2(Nothing, swViewEntityType_Edge)
    If IsEmpty(edges) Then Exit Sub

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swDraw

    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager

    Dim swSelData As SldWorks.SelectData
    Set swSelData = swSelMgr.CreateSelectData
    Set swSelData.View = swView

    Dim swModelExt As SldWorks.ModelDocExtension
    Set swModelExt = swModel.Extension

    Dim i As Long
    For i = LBound(edges) To UBound(edges)
        Dim swEdge As SldWorks.Edge
        Set swEdge = edges(i)
        If swEdge Is Nothing Then GoTo NextCallout

        Dim swCurve As SldWorks.Curve
        Set swCurve = swEdge.GetCurve
        If swCurve Is Nothing Then GoTo NextCallout
        ' Same SOLIDWORKS COM Boolean contract as CreateHoleOrdinateDims.
        If swCurve.IsCircle = False Then GoTo NextCallout

        Dim cp As Variant
        cp = swCurve.CircleParams

        swModel.ClearSelection2 True
        Dim objs(0 To 0) As Object
        Set objs(0) = swEdge
        If swModelExt.MultiSelect2(objs, False, swSelData) <> 1 Then GoTo NextCallout

        On Error Resume Next
        swDraw.AddHoleCallout2 cp(0) + 0.01, cp(1) + 0.01, 0
        If Err.Number <> 0 Then
            Err.Clear
            swModelExt.AddDimension2 cp(0) + 0.01, cp(1) + 0.01, 0
        End If
        On Error GoTo Failed

NextCallout:
        swModel.ClearSelection2 True
    Next i
    Exit Sub

Failed:
    Debug.Print "Hole callout fallback warning in view " & swView.Name & ": " & Err.Description
    swModel.ClearSelection2 True
End Sub

' Creates one ordinate group and returns its outcome: a swCreateOrdDimError_e
' value, or one of the CHAIN_* codes above.
'
' IModelDocExtension.AddOrdinateDimension Remarks (MCP, 2026-08-05):
'   "Before using this method, select the base entity to act as the datum
'    point for the ordinate dimension and any additional entities to include
'    in the group of ordinate dimensions."
'   "Selections made immediately after calling this method continue to add
'    ordinate dimensions to the group of ordinate dimensions. When you finish
'    adding ordinate dimensions to the group, use IModelDoc2::SetPickMode to
'    return to the default selection mode."
'
' The datum is therefore selection index 0, which is why the datum entity is
' written into the array first rather than in coordinate order.
'
' The horizontal chain runs first, so without SetPickMode the vertical
' chain's MultiSelect2 is a selection made after the call and appends to the
' still-open horizontal group instead of starting a vertical one. Every later
' selection in the run - auto-arrange, title block - inherits the same mode.
' ClearSelection2 empties the selection list; it does not close the group.
Private Function CreateOneOrdinateChain( _
    ByRef swModel As SldWorks.ModelDoc2, _
    ByRef swModelExt As SldWorks.ModelDocExtension, _
    ByRef swView As SldWorks.View, _
    ByRef swSelData As SldWorks.SelectData, _
    ByRef axis As AxisCandidates, _
    ByVal isHorizontal As Boolean, _
    ByVal viewOutline As Variant, _
    ByRef usedObjs() As Object, _
    ByRef usedCount As Long, _
    ByRef status As OrdinateRunStatus) As Long

    ' A single-entity chain is the datum alone: nothing to dimension against.
    If axis.Count < 2 Then
        CreateOneOrdinateChain = CHAIN_SKIPPED_TOO_FEW
        Exit Function
    End If

    ' Stations are already deduplicated per axis by AddAxisCandidate, so the
    ' selection is the datum followed by every other station in order.
    Dim selObjs() As Object
    ReDim selObjs(0 To axis.Count - 1)

    Set selObjs(0) = ClaimEntity(axis, axis.DatumIndex, usedObjs, _
        usedCount, status)

    Dim selCnt As Long
    selCnt = 1

    Dim i As Long
    For i = 0 To axis.Count - 1
        If i <> axis.DatumIndex Then
            Set selObjs(selCnt) = ClaimEntity(axis, i, usedObjs, _
                usedCount, status)
            selCnt = selCnt + 1
        End If
    Next i

    swModel.ClearSelection2 True
    If swModelExt.MultiSelect2(selObjs, False, swSelData) <> selCnt Then
        CreateOneOrdinateChain = CHAIN_SELECTION_FAILED
        swModel.ClearSelection2 True
        Exit Function
    End If

    Dim rc As Long
    If isHorizontal Then
        rc = swModelExt.AddOrdinateDimension(swHorizontalOrdinate, viewOutline(0), viewOutline(3) + 0.015, 0)
    Else
        rc = swModelExt.AddOrdinateDimension(swVerticalOrdinate, viewOutline(2) + 0.015, viewOutline(1), 0)
    End If

    ' Close the ordinate group before anything else selects. Unconditional:
    ' the group can be left open on a failure path too.
    swModel.SetPickMode

    If rc <> swCreateOrdDimErr_Success Then
        Debug.Print "Ordinate add failed in view " & swView.Name & ": code=" & rc
    End If

    swModel.ClearSelection2 True
    CreateOneOrdinateChain = rc
End Function

' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub

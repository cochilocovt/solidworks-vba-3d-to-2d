Option Explicit

' MCP-confirmed 2026-08-05 against swAddOrdinateDims_e and
' swCreateOrdDimError_e. Verify in SW2025 Object Browser before acceptance.
Private Const swHorizontalOrdinate As Long = 3
Private Const swVerticalOrdinate As Long = 2
Private Const swCreateOrdDimErr_Success As Long = 0

' swViewEntityType_e, MCP-confirmed 2026-08-05. The baseline passed a bare 1.
' swViewEntityType_SilhouetteEdge = 4 is the route to the reference drawing's
' +/-36 silhouette ordinates; not used yet (gap A4).
Private Const swViewEntityType_Edge As Long = 1

' Outcome codes returned by CreateOneOrdinateChain in addition to the
' swCreateOrdDimError_e values, which are all >= -1.
Private Const CHAIN_SKIPPED_TOO_FEW As Long = -100
Private Const CHAIN_SELECTION_FAILED As Long = -101

' Sentinel for LastFailureCode. 0 is swCreateOrdDimErr_Success, so an
' unset code of 0 reads as a successful chain in the report. The r3 run
' reported "Last code=0 (Unrecognised code)" for exactly that reason.
Public Const ORDINATE_CODE_UNSET As Long = -999

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
    stage = "SetSelectDataView"
    On Error Resume Next
    Set swSelData.View = swView

    If Err.Number <> 0 Then
        status.SelDataViewError = Err.Number
        status.SelDataScope = "Unscoped(err=" & CStr(Err.Number) & ")"
        Err.Clear
    Else
        status.SelDataScope = "ScopedToView"
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

    stage = "ScanEdges"

    Dim objs() As Object
    Dim px() As Double
    Dim py() As Double
    Dim cnt As Long
    cnt = 0

    Dim i As Long
    For i = LBound(edges) To UBound(edges)
        Dim swEdge As SldWorks.Edge
        Set swEdge = edges(i)
        If swEdge Is Nothing Then GoTo NextEdge

        status.CandidateEdgesSeen = status.CandidateEdgesSeen + 1

        Dim swCurve As SldWorks.Curve
        Set swCurve = swEdge.GetCurve
        If swCurve Is Nothing Then GoTo NextEdge

        ' "If Not <SOLIDWORKS COM Boolean>" is unsafe on this build. The
        ' contract recorded in docs/SOLIDWORKS_API_VALIDATION.md (2026-07-31
        ' third run) is that Not yields -2, which VBA treats as True, so this
        ' test rejected every edge: r6 read 349 edges and found 0 circles on
        ' a part with 12 holes. That table also records CBool(comCall) then
        ' Not failing specifically for ICurve.IsCircle, so "= False" is the
        ' safe form here, not a CBool round-trip.
        If swCurve.IsCircle = False Then GoTo NextEdge

        status.CircularEdgesSeen = status.CircularEdgesSeen + 1

        Dim cp As Variant
        cp = swCurve.CircleParams

        Dim ctr(2) As Double
        ctr(0) = cp(0)
        ctr(1) = cp(1)
        ctr(2) = cp(2)

        Dim swPt As SldWorks.MathPoint
        Set swPt = swMathUtil.CreatePoint(ctr)
        Set swPt = swPt.MultiplyTransform(swTransform)

        Dim arr As Variant
        arr = swPt.ArrayData

        If Not HasNearbyPoint(px, py, cnt, arr(0), arr(1)) Then
            ReDim Preserve objs(0 To cnt)
            ReDim Preserve px(0 To cnt)
            ReDim Preserve py(0 To cnt)
            Set objs(cnt) = swEdge
            px(cnt) = arr(0)
            py(cnt) = arr(1)
            cnt = cnt + 1
        End If
NextEdge:
    Next i

    If cnt < 2 Then
        status.ViewsWithTooFewCandidates = _
            status.ViewsWithTooFewCandidates + 1
        Exit Sub
    End If

    Dim datumIdx As Long
    datumIdx = ResolveDatumIndex(swView, viewOutline, px, py, cnt, datumType)

    RecordChainOutcome status, swView.Name, _
        CreateOneOrdinateChain(swModel, swModelExt, swView, swSelData, objs, px, cnt, datumIdx, True, viewOutline)
    RecordChainOutcome status, swView.Name, _
        CreateOneOrdinateChain(swModel, swModelExt, swView, swSelData, objs, py, cnt, datumIdx, False, viewOutline)
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
        " (circular: " & status.CircularEdgesSeen & ")" & vbCrLf

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

    text = text & "Ordinate chains created: " & status.ChainsCreated & _
        " of " & status.ChainsAttempted & " attempted" & vbCrLf

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
    edges = swView.GetVisibleEntities2(Nothing, 1)
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

Private Function HasNearbyPoint(ByRef px() As Double, ByRef py() As Double, ByVal cnt As Long, ByVal x As Double, ByVal y As Double) As Boolean
    Dim i As Long
    For i = 0 To cnt - 1
        If Abs(px(i) - x) < 0.0015 And Abs(py(i) - y) < 0.0015 Then
            HasNearbyPoint = True
            Exit Function
        End If
    Next i
End Function

Private Function ResolveDatumIndex(ByRef swView As SldWorks.View, ByVal viewOutline As Variant, ByRef ptX() As Double, ByRef ptY() As Double, ByVal ptsCount As Long, ByVal datumType As String) As Long
    Dim tx As Double, ty As Double

    If datumType = "Center" Then
        tx = (viewOutline(0) + viewOutline(2)) / 2 - swView.Position(0)
        ty = (viewOutline(1) + viewOutline(3)) / 2 - swView.Position(1)
    ElseIf datumType = "Top-Left" Then
        tx = viewOutline(0) - swView.Position(0)
        ty = viewOutline(3) - swView.Position(1)
    Else
        tx = viewOutline(0) - swView.Position(0)
        ty = viewOutline(1) - swView.Position(1)
    End If

    Dim i As Long
    Dim minDist As Double
    minDist = 1E+30
    ResolveDatumIndex = 0

    For i = 0 To ptsCount - 1
        Dim dist As Double
        dist = Sqr((ptX(i) - tx) ^ 2 + (ptY(i) - ty) ^ 2)
        If dist < minDist Then
            minDist = dist
            ResolveDatumIndex = i
        End If
    Next i
End Function

' Creates one ordinate group and returns its outcome: a swCreateOrdDimError_e
' value, or one of the CHAIN_* codes above.
'
' IModelDocExtension.AddOrdinateDimension Remarks (MCP, 2026-08-05):
'   "Selections made immediately after calling this method continue to add
'    ordinate dimensions to the group of ordinate dimensions. When you finish
'    adding ordinate dimensions to the group, use IModelDoc2::SetPickMode to
'    return to the default selection mode."
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
    ByRef objs() As Object, _
    ByRef coords() As Double, _
    ByVal cnt As Long, _
    ByVal datumIdx As Long, _
    ByVal isHorizontal As Boolean, _
    ByVal viewOutline As Variant) As Long

    Dim uniq() As Double
    Dim uniqCnt As Long
    Dim selObjs() As Object
    Dim selCnt As Long

    ReDim uniq(0 To 0)
    ReDim selObjs(0 To 0)

    uniq(0) = coords(datumIdx)
    Set selObjs(0) = objs(datumIdx)
    uniqCnt = 1
    selCnt = 1

    Dim i As Long, j As Long
    For i = 0 To cnt - 1
        If i <> datumIdx Then
            Dim isDup As Boolean
            isDup = False
            For j = 0 To uniqCnt - 1
                If Abs(coords(i) - uniq(j)) < 0.0015 Then
                    isDup = True
                    Exit For
                End If
            Next j

            If Not isDup Then
                ReDim Preserve uniq(0 To uniqCnt)
                ReDim Preserve selObjs(0 To selCnt)
                uniq(uniqCnt) = coords(i)
                Set selObjs(selCnt) = objs(i)
                uniqCnt = uniqCnt + 1
                selCnt = selCnt + 1
            End If
        End If
    Next i

    ' A single-entity chain is the datum alone: nothing to dimension against.
    If selCnt < 2 Then
        CreateOneOrdinateChain = CHAIN_SKIPPED_TOO_FEW
        Exit Function
    End If

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

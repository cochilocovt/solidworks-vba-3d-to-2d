Attribute VB_Name = "Module5_FallbackDimensionEngine"
Option Explicit

Private Const swHorizontalOrdinate As Long = 3
Private Const swVerticalOrdinate As Long = 2
Private Const swCreateOrdDimErr_Success As Long = 0

Public Sub CreateHoleOrdinateDims(ByRef swDraw As SldWorks.DrawingDoc, ByRef swView As SldWorks.View, ByVal datumType As String)
    On Error GoTo Failed

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swDraw

    Dim swModelExt As SldWorks.ModelDocExtension
    Set swModelExt = swModel.Extension

    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager

    Dim swSelData As SldWorks.SelectData
    Set swSelData = swSelMgr.CreateSelectData
    Set swSelData.View = swView

    Dim viewOutline As Variant
    viewOutline = swView.GetOutline

    Dim edges As Variant
    edges = swView.GetVisibleEntities2(Nothing, 1)
    If IsEmpty(edges) Then Exit Sub

    Dim swMathUtil As SldWorks.MathUtility
    Set swMathUtil = Application.SldWorks.GetMathUtility

    Dim swTransform As SldWorks.MathTransform
    Set swTransform = swView.ModelToViewTransform

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

        Dim swCurve As SldWorks.Curve
        Set swCurve = swEdge.GetCurve
        If swCurve Is Nothing Then GoTo NextEdge
        If Not swCurve.IsCircle Then GoTo NextEdge

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

    If cnt < 2 Then Exit Sub

    Dim datumIdx As Long
    datumIdx = ResolveDatumIndex(swView, viewOutline, px, py, cnt, datumType)

    CreateOneOrdinateChain swModel, swModelExt, swView, swSelData, objs, px, cnt, datumIdx, True, viewOutline
    CreateOneOrdinateChain swModel, swModelExt, swView, swSelData, objs, py, cnt, datumIdx, False, viewOutline
    Exit Sub

Failed:
    Debug.Print "Fallback ordinate error in view " & swView.Name & ": " & Err.Description
End Sub

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
        If Not swCurve.IsCircle Then GoTo NextCallout

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

Private Sub CreateOneOrdinateChain( _
    ByRef swModel As SldWorks.ModelDoc2, _
    ByRef swModelExt As SldWorks.ModelDocExtension, _
    ByRef swView As SldWorks.View, _
    ByRef swSelData As SldWorks.SelectData, _
    ByRef objs() As Object, _
    ByRef coords() As Double, _
    ByVal cnt As Long, _
    ByVal datumIdx As Long, _
    ByVal isHorizontal As Boolean, _
    ByVal viewOutline As Variant)

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

    swModel.ClearSelection2 True
    If swModelExt.MultiSelect2(selObjs, False, swSelData) <> selCnt Then Exit Sub

    Dim rc As Long
    If isHorizontal Then
        rc = swModelExt.AddOrdinateDimension(swHorizontalOrdinate, viewOutline(0), viewOutline(3) + 0.015, 0)
    Else
        rc = swModelExt.AddOrdinateDimension(swVerticalOrdinate, viewOutline(2) + 0.015, viewOutline(1), 0)
    End If

    If rc <> swCreateOrdDimErr_Success Then
        Debug.Print "Ordinate add failed in view " & swView.Name & ": code=" & rc
    End If

    swModel.ClearSelection2 True
End Sub

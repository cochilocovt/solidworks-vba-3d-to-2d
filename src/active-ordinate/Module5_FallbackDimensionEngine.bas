Attribute VB_Name = "Module5_FallbackDimensionEngine"

Option Explicit

' ============================================================================
' SOLIDWORKS 2025 fallback ordinate-dimension engine
'
' This module:
' - Activates each drawing view before creating selection data.
' - Uses view-scoped SelectData before selecting visible model edges.
' - Enumerates GetVisibleComponents and supplies each Component2 to
'   GetVisibleEntities2, as required by the SOLIDWORKS 2025 API.
' - Does NOT pass a Variant-held Edge ByRef.
' - Deduplicates concentric circular edges by projected circle centre.
' ============================================================================

Private Const swHorizontalOrdinate As Long = 3
Private Const swVerticalOrdinate As Long = 2
Private Const swCreateOrdDimErrSuccess As Long = 0
Private Const swViewEntityTypeEdge As Long = 1

' Verified swDrawingViewTypes_e members used by the strict view policy.
Private Const swDrawingSectionView As Long = 2
Private Const swDrawingDetailView As Long = 3
Private Const swDrawingProjectedView As Long = 4
Private Const swDrawingStandardView As Long = 6
Private Const swDrawingNamedView As Long = 7

' SOLIDWORKS API length unit = metre.
Private Const HoleCentreTolerance As Double = 0.0015
Private Const OrdinateCoordinateTolerance As Double = 0.0015

' ============================================================================
' PUBLIC ENTRY POINT USED BY Module2_DrawingPipeline
' ============================================================================
Public Sub CreateHoleOrdinateDims( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByVal datumType As String)

    On Error GoTo Failed

    Dim swModel As SldWorks.ModelDoc2
    Dim swModelExt As SldWorks.ModelDocExtension
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim swSelData As SldWorks.SelectData

    If swDraw Is Nothing Then
        Debug.Print "Module5: swDraw is Nothing."
        Exit Sub
    End If

    If swView Is Nothing Then
        Debug.Print "Module5: swView is Nothing."
        Exit Sub
    End If

    If Not IsOrdinateEligibleView(swView) Then
        Debug.Print "Module5: skipped unsupported ordinate view '" & swView.Name & "'."
        Exit Sub
    End If

    Set swModel = swDraw

    If swModel Is Nothing Then
        Debug.Print "Module5: swModel could not be obtained from swDraw."
        Exit Sub
    End If

    ' IMPORTANT:
    ' Activate the drawing view BEFORE requesting SelectData or assigning
    ' the View property. This is the ordering error in the prior rewrite.
    swModel.ClearSelection2 True

    Dim isActivated As Boolean
    isActivated = swDraw.ActivateView(swView.Name)

    If Not isActivated Then
        Debug.Print "Module5: failed to activate ordinate view '" & swView.Name & "'."
        GoTo SafeExit
    End If

    Set swModelExt = swModel.extension

    If swModelExt Is Nothing Then
        Debug.Print "Module5: ModelDocExtension is Nothing in '" & swView.Name & "'."
        GoTo SafeExit
    End If

    Set swSelMgr = swModel.SelectionManager

    If swSelMgr Is Nothing Then
        Debug.Print "Module5: SelectionManager is Nothing in '" & swView.Name & "'."
        GoTo SafeExit
    End If

    Set swSelData = swSelMgr.CreateSelectData

    If swSelData Is Nothing Then
        Debug.Print "Module5: CreateSelectData returned Nothing in '" & swView.Name & "'."
        GoTo SafeExit
    End If

    Set swSelData.View = swView

    Debug.Print String$(72, "-")
    Debug.Print "Module5: processing view '" & swView.Name & "'"
    Debug.Print "Module5: active view successfully assigned to SelectData."

    Dim viewOutline As Variant
    viewOutline = swView.GetOutline

    If IsEmpty(viewOutline) Then
        Debug.Print "Module5: GetOutline returned Empty for '" & swView.Name & "'."
        GoTo SafeExit
    End If

    Dim edges As Collection
    Set edges = GetVisibleEdgesForView(swView)

    If edges Is Nothing Then
        Debug.Print "Module5: visible-edge collection failed for '" & _
                    swView.Name & "'."
        GoTo SafeExit
    End If

    If edges.Count = 0 Then
        Debug.Print "Module5: GetVisibleEntities2 returned no edges for '" & _
                    swView.Name & "'."
        GoTo SafeExit
    End If

    Debug.Print "Module5: raw visible edge count = " & CStr(edges.Count)

    Dim swApp As SldWorks.SldWorks
    Dim swMathUtil As SldWorks.MathUtility
    Dim swTransform As SldWorks.MathTransform

    ' Use the active document's SW application reference; do not depend on
    ' Application.SldWorks, which can be unset depending on macro context.
    Set swApp = swModel.GetSwApp

    If swApp Is Nothing Then
        Debug.Print "Module5: GetSwApp returned Nothing."
        GoTo SafeExit
    End If

    Set swMathUtil = swApp.GetMathUtility

    If swMathUtil Is Nothing Then
        Debug.Print "Module5: GetMathUtility returned Nothing."
        GoTo SafeExit
    End If

    Set swTransform = swView.ModelToViewTransform

    If swTransform Is Nothing Then
        Debug.Print "Module5: ModelToViewTransform is Nothing for '" & _
                    swView.Name & "'."
        GoTo SafeExit
    End If

    Dim holeEdges() As Object
    Dim holeX() As Double
    Dim holeY() As Double
    Dim holeDia() As Double
    Dim holeCount As Long

    Dim i As Long
    Dim swEdge As SldWorks.Edge
    Dim swCurve As SldWorks.Curve
    Dim circleParams As Variant
    Dim centre(0 To 2) As Double
    Dim swPoint As SldWorks.MathPoint
    Dim pointData As Variant
    Dim x As Double
    Dim y As Double
    Dim dia As Double
    Dim existingIndex As Long
    Dim circularEdgeCount As Long

    holeCount = 0
    circularEdgeCount = 0

    For i = 1 To edges.Count
        Set swEdge = Nothing
        Set swCurve = Nothing
        Set swPoint = Nothing

        Set swEdge = edges.Item(i)
        If swEdge Is Nothing Then GoTo NextEdge

        Set swCurve = swEdge.GetCurve
        If swCurve Is Nothing Then GoTo NextEdge

        If Not swCurve.IsCircle Then GoTo NextEdge

        circleParams = swCurve.circleParams
        If IsEmpty(circleParams) Then GoTo NextEdge

        centre(0) = CDbl(circleParams(0))
        centre(1) = CDbl(circleParams(1))
        centre(2) = CDbl(circleParams(2))

        Set swPoint = swMathUtil.CreatePoint(centre)

        If swPoint Is Nothing Then GoTo NextEdge

        Set swPoint = swPoint.MultiplyTransform(swTransform)

        If swPoint Is Nothing Then GoTo NextEdge

        pointData = swPoint.ArrayData

        If IsEmpty(pointData) Then GoTo NextEdge

        x = CDbl(pointData(0))
        y = CDbl(pointData(1))
        dia = CDbl(circleParams(6)) * 2#

        circularEdgeCount = circularEdgeCount + 1

        existingIndex = FindHoleCentre(holeX, holeY, holeCount, x, y)

        If existingIndex = -1 Then
            AddHoleLocation holeEdges, holeX, holeY, holeDia, holeCount, _
                            swEdge, x, y, dia

        ElseIf dia < holeDia(existingIndex) Then
            Set holeEdges(existingIndex) = swEdge
            holeDia(existingIndex) = dia
        End If

NextEdge:
    Next i

    Debug.Print "Module5: circular edges = " & CStr(circularEdgeCount)
    Debug.Print "Module5: unique circular centres = " & CStr(holeCount)

    If holeCount < 2 Then
        Debug.Print "Module5: fewer than two unique circle centres; no dimensions created."
        GoTo SafeExit
    End If

    Dim datumIndex As Long

    datumIndex = ResolveDatumIndex( _
        swView, viewOutline, holeX, holeY, holeCount, datumType)

    Debug.Print "Module5: datum origin = '" & datumType & "'"
    Debug.Print "Module5: selected datum index = " & CStr(datumIndex + 1)

    CreateOneOrdinateChain swModel, swModelExt, swView, swSelData, _
                           holeEdges, holeX, holeCount, datumIndex, _
                           True, viewOutline

    CreateOneOrdinateChain swModel, swModelExt, swView, swSelData, _
                           holeEdges, holeY, holeCount, datumIndex, _
                           False, viewOutline

SafeExit:
    On Error Resume Next
    swModel.SetPickMode
    swModel.ClearSelection2 True
    swDraw.ActivateView ""
    On Error GoTo 0
    Exit Sub

Failed:
    Debug.Print "Module5 CreateHoleOrdinateDims error in '" & _
                swView.Name & "': " & Err.Number & " - " & Err.Description
    Resume SafeExit
End Sub

' ============================================================================
' OPTIONAL: PROCESS ALL NON-ISOMETRIC MODEL VIEWS
' ============================================================================
Public Sub AddOrdinateDimensionsToAllViews( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal datumType As String)

    On Error GoTo Failed

    Dim swView As SldWorks.View

    Set swView = swDraw.GetFirstView

    ' The first item is the sheet-format view.
    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    Do While Not swView Is Nothing
        If IsOrdinateEligibleView(swView) Then
            CreateHoleOrdinateDims swDraw, swView, datumType
        Else
            Debug.Print "Module5: skipped unsupported ordinate view '" & swView.Name & "'."
        End If

        Set swView = swView.GetNextView
    Loop

    Exit Sub

Failed:
    Debug.Print "Module5 AddOrdinateDimensionsToAllViews error: " & _
                Err.Number & " - " & Err.Description
End Sub

' ============================================================================
' VISIBLE EDGE ENUMERATION
' ============================================================================
Private Function GetVisibleEdgesForView( _
    ByRef swView As SldWorks.View) As Collection

    On Error GoTo Failed

    Dim visibleEdges As Collection
    Set visibleEdges = New Collection

    If swView Is Nothing Then
        Set GetVisibleEdgesForView = visibleEdges
        Exit Function
    End If

    Dim visibleComponents As Variant
    visibleComponents = swView.GetVisibleComponents

    If IsEmpty(visibleComponents) Then
        Set GetVisibleEdgesForView = visibleEdges
        Exit Function
    End If

    If IsNull(visibleComponents) Then
        Set GetVisibleEdgesForView = visibleEdges
        Exit Function
    End If

    If Not IsArray(visibleComponents) Then
        Debug.Print "Module5: GetVisibleComponents returned a non-array value for '" & _
                    swView.Name & "'."
        Set GetVisibleEdgesForView = visibleEdges
        Exit Function
    End If

    Dim i As Long
    Dim swComponent As SldWorks.Component2

    Debug.Print "Module5: visible component count = " & _
                CStr(UBound(visibleComponents) - LBound(visibleComponents) + 1)

    For i = LBound(visibleComponents) To UBound(visibleComponents)
        Set swComponent = Nothing
        Set swComponent = visibleComponents(i)

        If Not swComponent Is Nothing Then
            AppendVisibleComponentEdges swView, swComponent, visibleEdges
        End If
    Next i

    Set GetVisibleEdgesForView = visibleEdges
    Exit Function

Failed:
    Debug.Print "Module5 GetVisibleEdgesForView error in '" & swView.Name & _
                "': " & Err.Number & " - " & Err.Description
    Set GetVisibleEdgesForView = visibleEdges
End Function

Private Sub AppendVisibleComponentEdges( _
    ByRef swView As SldWorks.View, _
    ByRef swComponent As SldWorks.Component2, _
    ByRef visibleEdges As Collection)

    On Error GoTo Failed

    Dim componentEdges As Variant
    componentEdges = swView.GetVisibleEntities2( _
        swComponent, _
        swViewEntityTypeEdge)

    If IsEmpty(componentEdges) Then Exit Sub
    If IsNull(componentEdges) Then Exit Sub
    If Not IsArray(componentEdges) Then Exit Sub

    Dim i As Long
    Dim swEdge As SldWorks.Edge

    For i = LBound(componentEdges) To UBound(componentEdges)
        Set swEdge = Nothing
        Set swEdge = componentEdges(i)

        If Not swEdge Is Nothing Then
            visibleEdges.Add swEdge
        End If
    Next i

    Exit Sub

Failed:
    Debug.Print "Module5 AppendVisibleComponentEdges error in '" & swView.Name & _
                "': " & Err.Number & " - " & Err.Description
End Sub

' ============================================================================
' HOLE-CENTRE STORAGE
' ============================================================================
Private Function FindHoleCentre( _
    ByRef holeX() As Double, _
    ByRef holeY() As Double, _
    ByVal holeCount As Long, _
    ByVal candidateX As Double, _
    ByVal candidateY As Double) As Long

    Dim i As Long

    FindHoleCentre = -1

    For i = 0 To holeCount - 1
        If Abs(holeX(i) - candidateX) <= HoleCentreTolerance And _
           Abs(holeY(i) - candidateY) <= HoleCentreTolerance Then

            FindHoleCentre = i
            Exit Function
        End If
    Next i
End Function

Private Sub AddHoleLocation( _
    ByRef holeEdges() As Object, _
    ByRef holeX() As Double, _
    ByRef holeY() As Double, _
    ByRef holeDia() As Double, _
    ByRef holeCount As Long, _
    ByRef swEdge As SldWorks.Edge, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal dia As Double)

    If holeCount = 0 Then
        ReDim holeEdges(0 To 0)
        ReDim holeX(0 To 0)
        ReDim holeY(0 To 0)
        ReDim holeDia(0 To 0)
    Else
        ReDim Preserve holeEdges(0 To holeCount)
        ReDim Preserve holeX(0 To holeCount)
        ReDim Preserve holeY(0 To holeCount)
        ReDim Preserve holeDia(0 To holeCount)
    End If

    Set holeEdges(holeCount) = swEdge
    holeX(holeCount) = x
    holeY(holeCount) = y
    holeDia(holeCount) = dia

    holeCount = holeCount + 1
End Sub

' ============================================================================
' DATUM SELECTION
' ============================================================================
Private Function ResolveDatumIndex( _
    ByRef swView As SldWorks.View, _
    ByVal viewOutline As Variant, _
    ByRef pointX() As Double, _
    ByRef pointY() As Double, _
    ByVal pointCount As Long, _
    ByVal datumType As String) As Long

    Dim viewPosition As Variant
    Dim targetX As Double
    Dim targetY As Double

    viewPosition = swView.Position

    Select Case UCase$(Trim$(datumType))
        Case "CENTER"
            targetX = ((CDbl(viewOutline(0)) + CDbl(viewOutline(2))) / 2#) - _
                      CDbl(viewPosition(0))

            targetY = ((CDbl(viewOutline(1)) + CDbl(viewOutline(3))) / 2#) - _
                      CDbl(viewPosition(1))

        Case "TOP-LEFT"
            targetX = CDbl(viewOutline(0)) - CDbl(viewPosition(0))
            targetY = CDbl(viewOutline(3)) - CDbl(viewPosition(1))

        Case Else
            ' Bottom-Left.
            targetX = CDbl(viewOutline(0)) - CDbl(viewPosition(0))
            targetY = CDbl(viewOutline(1)) - CDbl(viewPosition(1))
    End Select

    Dim i As Long
    Dim minDistance As Double
    Dim currentDistance As Double

    minDistance = 1E+30
    ResolveDatumIndex = 0

    For i = 0 To pointCount - 1
        currentDistance = Sqr( _
            (pointX(i) - targetX) ^ 2 + _
            (pointY(i) - targetY) ^ 2)

        If currentDistance < minDistance Then
            minDistance = currentDistance
            ResolveDatumIndex = i
        End If
    Next i
End Function

' ============================================================================
' CREATE HORIZONTAL OR VERTICAL ORDINATE CHAIN
' ============================================================================
Private Sub CreateOneOrdinateChain( _
    ByRef swModel As SldWorks.ModelDoc2, _
    ByRef swModelExt As SldWorks.ModelDocExtension, _
    ByRef swView As SldWorks.View, _
    ByRef swSelData As SldWorks.SelectData, _
    ByRef holeEdges() As Object, _
    ByRef coordinates() As Double, _
    ByVal holeCount As Long, _
    ByVal datumIndex As Long, _
    ByVal isHorizontal As Boolean, _
    ByVal viewOutline As Variant)

    On Error GoTo Failed

    If swModel Is Nothing Then Exit Sub
    If swModelExt Is Nothing Then Exit Sub
    If swView Is Nothing Then Exit Sub
    If swSelData Is Nothing Then Exit Sub

    Dim uniqueCoordinates() As Double
    Dim selectedObjects() As Object
    Dim selectedCount As Long

    ReDim uniqueCoordinates(0 To 0)
    ReDim selectedObjects(0 To 0)

    uniqueCoordinates(0) = coordinates(datumIndex)
    Set selectedObjects(0) = holeEdges(datumIndex)
    selectedCount = 1

    Dim i As Long
    Dim j As Long
    Dim isDuplicate As Boolean

    For i = 0 To holeCount - 1
        If i <> datumIndex Then
            isDuplicate = False

            For j = 0 To selectedCount - 1
                If Abs(coordinates(i) - uniqueCoordinates(j)) <= _
                   OrdinateCoordinateTolerance Then

                    isDuplicate = True
                    Exit For
                End If
            Next j

            If Not isDuplicate Then
                ReDim Preserve uniqueCoordinates(0 To selectedCount)
                ReDim Preserve selectedObjects(0 To selectedCount)

                uniqueCoordinates(selectedCount) = coordinates(i)
                Set selectedObjects(selectedCount) = holeEdges(i)

                selectedCount = selectedCount + 1
            End If
        End If
    Next i

    If selectedCount < 2 Then
        Debug.Print "Module5: " & DirectionName(isHorizontal) & _
                    " chain skipped in '" & swView.Name & _
                    "'; fewer than two unique coordinates."
        GoTo SafeExit
    End If

    swModel.ClearSelection2 True

    Dim swDatumEntity As SldWorks.Entity
    Set swDatumEntity = selectedObjects(0)

    If swDatumEntity Is Nothing Then
        Debug.Print "Module5: datum entity is Nothing; ordinate chain not created."
        GoTo SafeExit
    End If

    If Not swDatumEntity.Select4(False, swSelData) Then
        Debug.Print "Module5: datum-first selection failed; ordinate chain not created."
        GoTo SafeExit
    End If

    Dim additionalObjects() As Object
    ReDim additionalObjects(0 To selectedCount - 2)

    For i = 1 To selectedCount - 1
        Set additionalObjects(i - 1) = selectedObjects(i)
    Next i

    Dim appendedSelections As Long
    appendedSelections = swModelExt.MultiSelect2( _
        additionalObjects, True, swSelData)

    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager

    Dim actualSelections As Long
    actualSelections = 0

    If Not swSelMgr Is Nothing Then
        actualSelections = swSelMgr.GetSelectedObjectCount2(-1)
    End If

    Debug.Print "Module5: " & DirectionName(isHorizontal) & _
                " selected = " & CStr(actualSelections) & " / " & _
                CStr(selectedCount) & " in '" & swView.Name & "'"

    If appendedSelections <> selectedCount - 1 Or _
       actualSelections <> selectedCount Then
        Debug.Print "Module5: " & DirectionName(isHorizontal) & _
                    " selection failed; ordinate chain not created."
        GoTo SafeExit
    End If

    Dim resultCode As Long
    Dim margin As Double

    margin = 0.015

    If isHorizontal Then
        resultCode = swModelExt.AddOrdinateDimension( _
            swHorizontalOrdinate, _
            CDbl(viewOutline(0)), _
            CDbl(viewOutline(3)) + margin, _
            0#)
    Else
        resultCode = swModelExt.AddOrdinateDimension( _
            swVerticalOrdinate, _
            CDbl(viewOutline(2)) + margin, _
            CDbl(viewOutline(1)), _
            0#)
    End If

    Debug.Print "Module5: " & DirectionName(isHorizontal) & _
                " AddOrdinateDimension return code = " & CStr(resultCode)

    If resultCode <> swCreateOrdDimErrSuccess Then
        Debug.Print "Module5: API rejected the " & _
                    DirectionName(isHorizontal) & " ordinate chain."
    End If

SafeExit:
    On Error Resume Next
    swModel.SetPickMode
    swModel.ClearSelection2 True
    On Error GoTo 0
    Exit Sub

Failed:
    Debug.Print "Module5 CreateOneOrdinateChain error in '" & _
                swView.Name & "': " & Err.Number & " - " & Err.Description
    Resume SafeExit
End Sub

' ============================================================================
' GENERAL HELPERS
' ============================================================================
Private Function DirectionName(ByVal isHorizontal As Boolean) As String
    If isHorizontal Then
        DirectionName = "horizontal"
    Else
        DirectionName = "vertical"
    End If
End Function

Private Function IsOrdinateEligibleView( _
    ByRef swView As SldWorks.View) As Boolean

    On Error GoTo NotEligible

    If swView Is Nothing Then Exit Function

    Dim viewType As Long
    viewType = swView.Type

    If viewType = swDrawingSectionView Or _
       viewType = swDrawingDetailView Then
        Exit Function
    End If

    If viewType <> swDrawingProjectedView And _
       viewType <> swDrawingStandardView And _
       viewType <> swDrawingNamedView Then
        Exit Function
    End If

    Dim orientationName As String
    orientationName = UCase$(Trim$(swView.GetOrientationName))

    Select Case orientationName
        Case "*FRONT", "FRONT", _
             "*BACK", "BACK", _
             "*LEFT", "LEFT", _
             "*RIGHT", "RIGHT", _
             "*TOP", "TOP", _
             "*BOTTOM", "BOTTOM"
            IsOrdinateEligibleView = True
    End Select

    Exit Function

NotEligible:
    IsOrdinateEligibleView = False
End Function


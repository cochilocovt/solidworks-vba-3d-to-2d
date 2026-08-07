Option Explicit

Private Const swImportModelItemsFromEntireModel As Long = 0

Private Const swInsertDimensions As Long = 8
Private Const swInsertGTols As Long = 32
Private Const swInsertDimensionsMarkedForDrawing As Long = 32768
Private Const swInsertHoleWizardProfileDimensions As Long = 65536
Private Const swInsertHoleWizardLocationDimensions As Long = 131072
Private Const swInsertDimensionsNotMarkedForDrawing As Long = 524288
Private Const swInsertholeCallout As Long = 1048576

Private Const swAlignDimensionType_AutoArrange As Long = 0

Public Function ImportModelItemsAcrossDrawing( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal anchorViewName As String) As Long

    On Error GoTo Failed

    If swDrawModel Is Nothing Then Exit Function
    If swDraw Is Nothing Then Exit Function

    Dim mask As Long
    mask = GetModelItemMask()

    Dim total As Long
    total = 0

    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""

    ' Parameter 4 is DuplicateDims. Its documented contract reads backwards
    ' from its name: "True to eliminate duplicate dimensions, false to allow
    ' duplicate dimensions" (IDrawingDoc.InsertModelAnnotations4, MCP
    ' 2026-08-05). Passing False alongside AllViews=True inserted the same
    ' annotation into every view, which is the repeated-hole-callout defect.
    Dim inserted As Variant
    inserted = swDraw.InsertModelAnnotations4( _
                    swImportModelItemsFromEntireModel, _
                    mask, _
                    True, _
                    True, _
                    False, _
                    False, _
                    False, _
                    False)

    total = CountVariantItems(inserted)

    If total = 0 Then
        total = ImportModelItemsPerView(swDrawModel, swDraw, mask)
    End If

    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""

    ImportModelItemsAcrossDrawing = total
    Exit Function

Failed:
    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""
    ImportModelItemsAcrossDrawing = 0
End Function

Private Function ImportModelItemsPerView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal mask As Long) As Long

    On Error GoTo Failed

    Dim runningTotal As Long
    runningTotal = 0

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim viewName As String
        viewName = swView.Name

        swDrawModel.ClearSelection2 True
        swDraw.ActivateView viewName

        Dim ok As Boolean
        ok = swDrawModel.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)

        If ok Then
            ' DuplicateDims=True, same contract as the whole-drawing call.
            Dim inserted As Variant
            inserted = swDraw.InsertModelAnnotations4( _
                            swImportModelItemsFromEntireModel, _
                            mask, _
                            False, _
                            True, _
                            False, _
                            False, _
                            False, _
                            False)

            runningTotal = runningTotal + CountVariantItems(inserted)
        End If

        swDrawModel.ClearSelection2 True
        swDraw.ActivateView ""

        Set swView = swView.GetNextView
    Loop

    ImportModelItemsPerView = runningTotal
    Exit Function

Failed:
    ImportModelItemsPerView = 0
End Function

Public Function CountDisplayDimensionsInView(ByRef swView As SldWorks.View) As Long
    On Error GoTo Failed

    Dim vDims As Variant
    vDims = swView.GetDisplayDimensions
    CountDisplayDimensionsInView = CountVariantItems(vDims)
    Exit Function

Failed:
    CountDisplayDimensionsInView = 0
End Function

' Section 7 of BASELINE_TO_REFERENCE_DRAWING_GAP.md: whether InsertModelAnnotations4's
' swInsertholeCallout bit (gated on GlobalConfig.ImportHoleCallouts, default
' True) actually produces any IDisplayDimension.IsHoleCallout=True entries on
' this fixture has never been checked -- QA has only ever counted dimensions
' generically. IDisplayDimension.IsHoleCallout MCP-confirmed 2026-08-06:
' "Gets whether this display dimension is a hole callout", returns Boolean.
Public Function CountHoleCalloutsInView(ByRef swView As SldWorks.View) As Long
    On Error GoTo Failed

    Dim vDims As Variant
    vDims = swView.GetDisplayDimensions
    If IsEmpty(vDims) Then Exit Function

    Dim i As Long
    For i = LBound(vDims) To UBound(vDims)
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = vDims(i)
        If Not swDispDim Is Nothing Then
            ' Same SOLIDWORKS COM Boolean contract as the rest of this trunk:
            ' only "= False" is reliable.
            If Not (swDispDim.IsHoleCallout = False) Then
                CountHoleCalloutsInView = CountHoleCalloutsInView + 1
            End If
        End If
    Next i
    Exit Function

Failed:
    CountHoleCalloutsInView = 0
End Function

Public Sub AutoArrangeAllDrawingDimensions(ByRef swDrawModel As SldWorks.ModelDoc2, ByRef swDraw As SldWorks.DrawingDoc)
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        AutoArrangeDimensionsInView swDrawModel, swView
        Set swView = swView.GetNextView
    Loop
End Sub

Public Sub AutoArrangeDimensionsInView(ByRef swDrawModel As SldWorks.ModelDoc2, ByRef swView As SldWorks.View)
    On Error GoTo SafeExit

    If swView Is Nothing Then Exit Sub
    swDrawModel.ClearSelection2 True

    Dim vDims As Variant
    vDims = swView.GetDisplayDimensions
    If IsEmpty(vDims) Then GoTo SafeExit

    ' Scope the selection to this view. Passing Nothing leaves the selection
    ' unscoped, so in a multi-view drawing an annotation can be picked up
    ' outside the view being arranged.
    Dim swSelData As SldWorks.SelectData
    Set swSelData = swDrawModel.SelectionManager.CreateSelectData
    Set swSelData.View = swView

    Dim i As Long
    For i = LBound(vDims) To UBound(vDims)
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = vDims(i)

        If Not swDispDim Is Nothing Then
            Dim swAnn As SldWorks.Annotation
            Set swAnn = swDispDim.GetAnnotation
            If Not swAnn Is Nothing Then swAnn.Select3 True, swSelData
        End If
    Next i

    swDrawModel.Extension.AlignDimensions swAlignDimensionType_AutoArrange, 0.06

SafeExit:
    swDrawModel.ClearSelection2 True
End Sub

Public Function GetFirstRealViewName(ByRef swDraw As SldWorks.DrawingDoc) As String
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    If Not swView Is Nothing Then GetFirstRealViewName = swView.Name
End Function

' Every member below was confirmed against swInsertAnnotation_e via MCP on
' 2026-08-05. Verify in the SW2025 Object Browser before acceptance.
'
' swInsertholeCallout is now gated on the form's ImportHoleCallouts checkbox.
' It was previously OR-ed in unconditionally, which left that control dead.
Private Function GetModelItemMask() As Long
    Dim mask As Long
    mask = swInsertDimensions Or _
            swInsertGTols Or _
            swInsertDimensionsMarkedForDrawing Or _
            swInsertHoleWizardProfileDimensions Or _
            swInsertHoleWizardLocationDimensions

    If Module1_Main.GlobalConfig.ImportHoleCallouts Then
        mask = mask Or swInsertholeCallout
    End If

    GetModelItemMask = mask
End Function

Private Function CountVariantItems(ByVal vItems As Variant) As Long
    On Error GoTo Failed

    If IsEmpty(vItems) Then Exit Function

    If IsArray(vItems) Then
        CountVariantItems = UBound(vItems) - LBound(vItems) + 1
    Else
        CountVariantItems = 1
    End If

    Exit Function

Failed:
    CountVariantItems = 0
End Function


' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub

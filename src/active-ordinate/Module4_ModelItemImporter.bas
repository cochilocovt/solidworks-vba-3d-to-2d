Attribute VB_Name = "Module4_ModelItemImporter"
Option Explicit

' Verified against the installed SOLIDWORKS 2025 SP1.2 type library.
Private Const swImportModelItemsFromEntireModel As Long = 0

Private Const swInsertGTols As Long = 32
Private Const swInsertNotes As Long = 64
Private Const swInsertDimensionsMarkedForDrawing As Long = 32768
Private Const swInsertHoleWizardProfileDimensions As Long = 65536
Private Const swInsertHoleWizardLocationDimensions As Long = 131072
Private Const swInsertholeCallout As Long = 1048576

Private Const swAlignDimensionType_AutoArrange As Long = 0

Public Function ImportModelItemsAcrossDrawing( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal anchorViewName As String) As Long

    On Error GoTo Failed

    If swDrawModel Is Nothing Then Exit Function
    If swDraw Is Nothing Then Exit Function

    If Len(Trim$(anchorViewName)) = 0 Then
        anchorViewName = GetFirstRealViewName(swDraw)
    End If

    If Len(Trim$(anchorViewName)) = 0 Then
        Debug.Print "Module4: no real drawing view is available for model-item import."
        Exit Function
    End If

    Dim mask As Long
    mask = GetModelItemMask()

    Dim isSelected As Boolean
    Dim isActivated As Boolean

    swDrawModel.ClearSelection2 True

    isSelected = swDrawModel.Extension.SelectByID2( _
        anchorViewName, _
        "DRAWINGVIEW", _
        0#, _
        0#, _
        0#, _
        False, _
        0, _
        Nothing, _
        0)

    isActivated = swDraw.ActivateView(anchorViewName)

    If Not isActivated Then
        Debug.Print "Module4: failed to activate anchor view '" & anchorViewName & "'."
        GoTo SafeExit
    End If

    If Not isSelected Then
        Debug.Print "Module4: SelectByID2 returned False for anchor view '" & _
                    anchorViewName & "'; continuing with the activated view."
    End If

    ' SOLIDWORKS' 2025 example clears the selection after activating the view;
    ' InsertModelAnnotations4 uses the active drawing-view context.
    swDrawModel.ClearSelection2 True

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

    Dim total As Long
    total = CountVariantItems(inserted)

    Debug.Print "Module4: whole-drawing InsertModelAnnotations4 count = " & CStr(total)

    If total = 0 Then
        total = ImportModelItemsPerView(swDrawModel, swDraw, mask)
    End If

    ImportModelItemsAcrossDrawing = total

SafeExit:
    On Error Resume Next
    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""
    On Error GoTo 0
    Exit Function

Failed:
    Dim errorNumber As Long
    Dim errorDescription As String

    errorNumber = Err.Number
    errorDescription = Err.Description

    Debug.Print "Module4 ImportModelItemsAcrossDrawing error: " & _
                CStr(errorNumber) & " - " & errorDescription
    ImportModelItemsAcrossDrawing = 0

    On Error Resume Next
    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""
    On Error GoTo 0
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

    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    Do While Not swView Is Nothing
        Dim viewName As String
        viewName = swView.Name

        Dim isSelected As Boolean
        Dim isActivated As Boolean

        swDrawModel.ClearSelection2 True

        isSelected = swDrawModel.Extension.SelectByID2( _
            viewName, _
            "DRAWINGVIEW", _
            0#, _
            0#, _
            0#, _
            False, _
            0, _
            Nothing, _
            0)

        isActivated = swDraw.ActivateView(viewName)

        If isActivated Then
            If Not isSelected Then
                Debug.Print "Module4: SelectByID2 returned False for retry view '" & _
                            viewName & "'; continuing with the activated view."
            End If

            swDrawModel.ClearSelection2 True

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

            Dim insertedCount As Long
            insertedCount = CountVariantItems(inserted)
            runningTotal = runningTotal + insertedCount

            Debug.Print "Module4: selected-view import '" & viewName & _
                        "' count = " & CStr(insertedCount)
        Else
            Debug.Print "Module4: selected-view retry skipped; activation failed for '" & _
                        viewName & "'."
        End If

        swDrawModel.ClearSelection2 True
        swDraw.ActivateView ""
        Set swView = swView.GetNextView
    Loop

    ImportModelItemsPerView = runningTotal
    Exit Function

Failed:
    Debug.Print "Module4 ImportModelItemsPerView error: " & _
                CStr(Err.Number) & " - " & Err.Description
    ImportModelItemsPerView = runningTotal
End Function

Public Function CountDisplayDimensionsInView( _
    ByRef swView As SldWorks.View) As Long

    On Error GoTo Failed

    If swView Is Nothing Then Exit Function

    Dim vDims As Variant
    vDims = swView.GetDisplayDimensions

    CountDisplayDimensionsInView = CountVariantItems(vDims)
    Exit Function

Failed:
    CountDisplayDimensionsInView = 0
End Function

Public Sub AutoArrangeAllDrawingDimensions( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc)

    If swDrawModel Is Nothing Then Exit Sub
    If swDraw Is Nothing Then Exit Sub

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    Do While Not swView Is Nothing
        AutoArrangeDimensionsInView swDrawModel, swView
        Set swView = swView.GetNextView
    Loop
End Sub

Public Sub AutoArrangeDimensionsInView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View)

    On Error GoTo SafeExit

    If swDrawModel Is Nothing Then Exit Sub
    If swView Is Nothing Then Exit Sub

    swDrawModel.ClearSelection2 True

    Dim vDims As Variant
    vDims = swView.GetDisplayDimensions

    If IsEmpty(vDims) Then GoTo SafeExit
    If Not IsArray(vDims) Then GoTo SafeExit

    Dim i As Long
    For i = LBound(vDims) To UBound(vDims)
        Dim swDispDim As SldWorks.DisplayDimension
        Set swDispDim = vDims(i)

        If Not swDispDim Is Nothing Then
            Dim swAnn As SldWorks.Annotation
            Set swAnn = swDispDim.GetAnnotation

            If Not swAnn Is Nothing Then
                swAnn.Select3 True, Nothing
            End If
        End If
    Next i

    Dim aligned As Boolean
    aligned = swDrawModel.Extension.AlignDimensions( _
        swAlignDimensionType_AutoArrange, _
        0.06)

    If Not aligned Then
        Debug.Print "Module4: AlignDimensions failed in view '" & swView.Name & "'."
    End If

SafeExit:
    On Error Resume Next
    swDrawModel.ClearSelection2 True
    On Error GoTo 0
End Sub

Public Function GetFirstRealViewName( _
    ByRef swDraw As SldWorks.DrawingDoc) As String

    If swDraw Is Nothing Then Exit Function

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    If Not swView Is Nothing Then
        GetFirstRealViewName = swView.Name
    End If
End Function

Private Function GetModelItemMask() As Long
    GetModelItemMask = _
        swInsertDimensionsMarkedForDrawing Or _
        swInsertGTols Or _
        swInsertNotes Or _
        swInsertHoleWizardProfileDimensions Or _
        swInsertHoleWizardLocationDimensions

    If Module1_Main.GlobalConfig.ImportHoleCallouts Then
        GetModelItemMask = GetModelItemMask Or swInsertholeCallout
    End If
End Function

Private Function CountVariantItems(ByVal vItems As Variant) As Long
    On Error GoTo Failed

    If IsEmpty(vItems) Then Exit Function
    If IsNull(vItems) Then Exit Function

    If IsArray(vItems) Then
        CountVariantItems = UBound(vItems) - LBound(vItems) + 1
    Else
        CountVariantItems = 1
    End If

    Exit Function

Failed:
    CountVariantItems = 0
End Function

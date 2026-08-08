Option Explicit

Private Const swImportModelItemsFromEntireModel As Long = 0

Private Const swInsertDimensions As Long = 8
Private Const swInsertGTols As Long = 32
Private Const swInsertDimensionsMarkedForDrawing As Long = 32768
Private Const swInsertHoleWizardProfileDimensions As Long = 65536
Private Const swInsertHoleWizardLocationDimensions As Long = 131072
Private Const swInsertDimensionsNotMarkedForDrawing As Long = 524288
Private Const swInsertholeCallout As Long = 1048576

' Installed SW2025 Phase 0 import logs recorded selectionType=12 for each
' selected drawing view. MCP swSelectType_e confirms swSelDRAWINGVIEWS=12.
Private Const swSelDRAWINGVIEWS As Long = 12

Private Const swAlignDimensionType_AutoArrange As Long = 0

Private mLastImportDiagnostics As String

Public Function ImportModelItemsAcrossDrawing( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal anchorViewName As String) As Long

    On Error GoTo Failed

    If swDrawModel Is Nothing Then Exit Function
    If swDraw Is Nothing Then Exit Function

    Dim mask As Long
    mask = GetModelItemMask()

    If Len(Trim$(anchorViewName)) = 0 Then
        anchorViewName = GetFirstRealViewName(swDraw)
    End If

    mLastImportDiagnostics = "Model item import transactions:" & vbCrLf & _
        "  strategy=SelectedViews; anchor=" & SafeText(anchorViewName) & _
        "; mask=" & CStr(mask) & _
        "; allViews=False; duplicateDims=True" & vbCrLf

    ' IDrawingDoc.InsertModelAnnotations4 inserts into the currently selected
    ' drawing view. The previous whole-drawing call cleared selection, activated
    ' sheet context, and then used AllViews=True. Because it returned ordinary
    ' dimensions, the selected-view retry was never reached even when the
    ' hole-callout category stayed at zero. Selected-view import is now the only
    ' transaction, with the anchor first and every other eligible view second.
    Dim total As Long
    total = ImportModelItemsPerView( _
        swDrawModel, swDraw, mask, anchorViewName)

    swDrawModel.ClearSelection2 True

    ImportModelItemsAcrossDrawing = total
    Exit Function

Failed:
    AppendImportDiagnostic "  IMPORT_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & SafeText(Err.Description)
    swDrawModel.ClearSelection2 True
    ImportModelItemsAcrossDrawing = 0
End Function

Private Function ImportModelItemsPerView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal mask As Long, _
    ByVal anchorViewName As String) As Long

    On Error GoTo Failed

    Dim runningTotal As Long
    runningTotal = 0

    Dim anchorView As SldWorks.View
    Set anchorView = FindRealViewByName(swDraw, anchorViewName)

    If Not anchorView Is Nothing Then
        If AllowsModelItemImport(anchorView) Then
            runningTotal = runningTotal + ImportModelItemsForView( _
                swDrawModel, swDraw, anchorView, mask)
        Else
            AppendImportDiagnostic "  IMPORT_SKIP|view=" & _
                SafeText(anchorViewName) & "|reason=AnchorNotEligible"
        End If
    Else
        AppendImportDiagnostic "  IMPORT_WARNING|view=" & _
            SafeText(anchorViewName) & "|reason=AnchorNotFound"
    End If

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim viewName As String
        viewName = swView.Name

        If StrComp(viewName, anchorViewName, vbBinaryCompare) <> 0 Then
            If AllowsModelItemImport(swView) Then
                runningTotal = runningTotal + ImportModelItemsForView( _
                    swDrawModel, swDraw, swView, mask)
            Else
                AppendImportDiagnostic "  IMPORT_SKIP|view=" & _
                    SafeText(viewName) & "|reason=PictorialOrSheet"
            End If
        End If

        Set swView = swView.GetNextView
    Loop

    ImportModelItemsPerView = runningTotal
    Exit Function

Failed:
    AppendImportDiagnostic "  IMPORT_LOOP_ERROR|error=" & _
        CStr(Err.Number) & "|description=" & SafeText(Err.Description)
    ImportModelItemsPerView = runningTotal
End Function

Private Function ImportModelItemsForView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByVal mask As Long) As Long

    On Error GoTo Failed

    If swView Is Nothing Then Exit Function

    Dim viewName As String
    viewName = swView.Name

    Dim dimsBefore As Long
    Dim calloutsBefore As Long
    dimsBefore = CountDisplayDimensionsInView(swView)
    calloutsBefore = CountHoleCalloutsInView(swView)

    swDrawModel.ClearSelection2 True

    Dim activated As Boolean
    activated = swDraw.ActivateView(viewName)

    Dim activeMatched As Boolean
    activeMatched = ActiveViewMatches(swDraw, viewName)

    If activated = False And activeMatched = False Then
        AppendImportDiagnostic "  IMPORT_SKIP|view=" & SafeText(viewName) & _
            "|reason=ActivationFailed"
        GoTo CleanExit
    End If

    Dim selectedByID As Boolean
    selectedByID = swDrawModel.Extension.SelectByID2( _
        viewName, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0)

    Dim selectionCount As Long
    selectionCount = swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)

    Dim selectionType As Long
    selectionType = 0
    If selectionCount = 1 Then
        selectionType = swDrawModel.SelectionManager.GetSelectedObjectType3(1, -1)
    End If

    If activeMatched = False Or selectionCount <> 1 Or _
        selectionType <> swSelDRAWINGVIEWS Then

        AppendImportDiagnostic "  IMPORT_SKIP|view=" & SafeText(viewName) & _
            "|reason=SelectionContractFailed" & _
            "|activated=" & CStr(activated) & _
            "|activeMatched=" & CStr(activeMatched) & _
            "|selectedByID=" & CStr(selectedByID) & _
            "|selectionCount=" & CStr(selectionCount) & _
            "|selectionType=" & CStr(selectionType)
        GoTo CleanExit
    End If

    ' DuplicateDims=True means eliminate duplicates. AllViews=False binds the
    ' insertion to the selected drawing view, which is verified immediately
    ' above by active-view and selection-list readback.
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

    Dim returnedCount As Long
    Dim dimsAfter As Long
    Dim calloutsAfter As Long
    returnedCount = CountVariantItems(inserted)
    dimsAfter = CountDisplayDimensionsInView(swView)
    calloutsAfter = CountHoleCalloutsInView(swView)

    AppendImportDiagnostic "  IMPORT_VIEW|view=" & SafeText(viewName) & _
        "|activated=" & CStr(activated) & _
        "|activeMatched=" & CStr(activeMatched) & _
        "|selectedByID=" & CStr(selectedByID) & _
        "|selectionCount=" & CStr(selectionCount) & _
        "|selectionType=" & CStr(selectionType) & _
        "|returned=" & CStr(returnedCount) & _
        "|dimsBefore=" & CStr(dimsBefore) & _
        "|dimsAfter=" & CStr(dimsAfter) & _
        "|calloutsBefore=" & CStr(calloutsBefore) & _
        "|calloutsAfter=" & CStr(calloutsAfter)

    ImportModelItemsForView = returnedCount

CleanExit:
    swDrawModel.ClearSelection2 True
    Exit Function

Failed:
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    AppendImportDiagnostic "  IMPORT_VIEW_ERROR|view=" & SafeText(viewName) & _
        "|error=" & CStr(errorNumber) & _
        "|description=" & SafeText(errorDescription)
    Resume CleanExit
End Function

Private Function FindRealViewByName( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal viewName As String) As SldWorks.View

    On Error GoTo SafeExit

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        If StrComp(swView.Name, viewName, vbBinaryCompare) = 0 Then
            Set FindRealViewByName = swView
            Exit Function
        End If
        Set swView = swView.GetNextView
    Loop

SafeExit:
End Function

Private Function AllowsModelItemImport(ByRef swView As SldWorks.View) As Boolean
    If swView Is Nothing Then Exit Function

    Dim role As Long
    role = Module8_ViewClassifier.ClassifyView(swView)

    Select Case role
        Case Module8_ViewClassifier.VIEW_ROLE_PICTORIAL, _
             Module8_ViewClassifier.VIEW_ROLE_SHEET
            AllowsModelItemImport = False
        Case Else
            AllowsModelItemImport = True
    End Select
End Function

Private Function ActiveViewMatches( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal expectedViewName As String) As Boolean

    On Error GoTo SafeExit

    Dim activeViewObject As Object
    Set activeViewObject = swDraw.ActiveDrawingView
    If activeViewObject Is Nothing Then Exit Function

    Dim activeView As SldWorks.View
    Set activeView = activeViewObject

    ActiveViewMatches = (StrComp( _
        activeView.Name, expectedViewName, vbBinaryCompare) = 0)

SafeExit:
End Function

Private Sub AppendImportDiagnostic(ByVal text As String)
    mLastImportDiagnostics = mLastImportDiagnostics & text & vbCrLf
End Sub

Private Function SafeText(ByVal text As String) As String
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    text = Replace(text, "|", "/")
    SafeText = text
End Function

Public Function DescribeLastImportTransactions() As String
    If Len(mLastImportDiagnostics) = 0 Then
        DescribeLastImportTransactions = _
            "Model item import transactions: not run"
    Else
        DescribeLastImportTransactions = mLastImportDiagnostics
    End If
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

' Direct callout readback. Live r25/r26 proved that the earlier sheet-context
' AllViews=True transaction requested swInsertholeCallout but produced zero
' IDisplayDimension.IsHoleCallout=True entries. The selected-view transaction
' now records the same count before and after every individual view import.
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

' What the run ASKED FOR, as distinct from what it produced.
'
' Every QA number before this reported created output only. That made two
' different failures indistinguishable in the run record: "SOLIDWORKS was
' asked and declined" versus "nothing ever asked". r25 hit this twice in one
' run - zero hole callouts with no way to tell whether the mask bit was set,
' and a missing isometric view that turned out to be an unticked, registry-
' persisted checkbox rather than any creation failure.
'
' UserForm1 seeds every one of these from SaveSetting/ReadBoolSetting and
' writes them back on OK, so GlobalConfig at report time is the operator's
' actual answer for this run - which is exactly the value that was never
' recorded anywhere. Same root cause as A11.
Public Function DescribeRequestedConfig() As String
    Dim text As String

    text = "Requested config (what the run asked for, not what it produced):" & vbCrLf

    text = text & "  Model item mask: " & GetModelItemMask() & _
        " (holeCallout bit " & OnOff(Module1_Main.GlobalConfig.ImportHoleCallouts) & ")" & vbCrLf

    text = text & "  Producers: modelDims=" & OnOff(Module1_Main.GlobalConfig.UseModelDimensions) & _
        " ordinates=" & OnOff(Module1_Main.GlobalConfig.UseOrdinateDims) & _
        " autoArrange=" & OnOff(Module1_Main.GlobalConfig.AutoArrange) & vbCrLf

    text = text & "  Views requested: " & _
        "front=" & OnOff(Module1_Main.GlobalConfig.CreateFront) & _
        " top=" & OnOff(Module1_Main.GlobalConfig.CreateTop) & _
        " bottom=" & OnOff(Module1_Main.GlobalConfig.CreateBottom) & _
        " left=" & OnOff(Module1_Main.GlobalConfig.CreateLeft) & _
        " right=" & OnOff(Module1_Main.GlobalConfig.CreateRight) & _
        " back=" & OnOff(Module1_Main.GlobalConfig.CreateBack) & _
        " iso=" & OnOff(Module1_Main.GlobalConfig.CreateIso) & _
        " section=" & OnOff(Module1_Main.GlobalConfig.CreateSection) & vbCrLf

    text = text & "  Display: HLR=" & OnOff(Module1_Main.GlobalConfig.UseHLR) & _
        " (ordinate harvest forces HLR regardless - see HarvestDisplayMode)" & vbCrLf

    DescribeRequestedConfig = text
End Function

Private Function OnOff(ByVal flag As Boolean) As String
    If flag Then
        OnOff = "ON"
    Else
        OnOff = "off"
    End If
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

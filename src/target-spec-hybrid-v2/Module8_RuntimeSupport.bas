Option Explicit

' Explicit design tolerances. These values are emitted to QA and must be
' confirmed by the authorized-fixture runtime evidence before acceptance.
Public Const GEOMETRY_TOLERANCE_M As Double = 0.0000001
Public Const PROJECTED_TOLERANCE_M As Double = 0.00001
Public Const AXIS_NORMAL_MIN_COS As Double = 0.999847695
Public Const LAYOUT_MARGIN_M As Double = 0.012

Private Const swRebuildActiveDoc As Long = 2
Private Const swObjectSame As Long = 1
Private Const swModelRebuildStatus_FullyRebuilt As Long = 0

Private Const swDrawingProjectedView As Long = 4
Private Const swDrawingStandardView As Long = 6
Private Const swDrawingNamedView As Long = 7
Private Const swZoneTopMargin As Long = 0
Private Const swZoneBottomMargin As Long = 1
Private Const swZoneRightMargin As Long = 2
Private Const swZoneLeftMargin As Long = 3

Private mProvenOrdinateViews As Object

Public Sub ResetProvenViewRegistry()
    Set mProvenOrdinateViews = CreateObject("Scripting.Dictionary")
End Sub

Public Sub RegisterProvenOrdinateView( _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence)

    If mProvenOrdinateViews Is Nothing Then ResetProvenViewRegistry

    Dim viewName As String
    viewName = LCase$(GetViewName(swView))

    If Len(viewName) = 0 Then
        evidence.AddFailure "Cannot register an unnamed ordinate-eligible view."
        Exit Sub
    End If

    If Not mProvenOrdinateViews.Exists(viewName) Then
        mProvenOrdinateViews.Add viewName, True
    End If

    evidence.AddInfo "Registered created orthographic view as proven for ordinate " & _
        "policy: '" & GetViewName(swView) & "'."
End Sub

Public Function ActivateDrawingDocument( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim errors As Long
    Dim activated As SldWorks.ModelDoc2

    Set activated = swApp.ActivateDoc3( _
        swDrawModel.GetTitle, False, swRebuildActiveDoc, errors)

    If activated Is Nothing Then
        evidence.AddFailure "ActivateDoc3 returned Nothing; error=" & CStr(errors)
        Exit Function
    End If

    If Not ObjectsAreSame(swApp, activated, swDrawModel) Then
        evidence.AddFailure "ActivateDoc3 activated a different document."
        Exit Function
    End If

    ActivateDrawingDocument = True
    Exit Function

Failed:
    evidence.AddFailure "ActivateDoc3 error " & CStr(Err.Number) & _
        ": " & Err.Description
End Function

Public Function ActivateDrawingView( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence, _
    ByVal operationName As String) As Boolean

    On Error GoTo Failed

    If swView Is Nothing Then
        evidence.AddFailure operationName & ": view is Nothing."
        Exit Function
    End If

    Dim viewName As String
    viewName = GetViewName(swView)

    If Len(viewName) = 0 Then
        evidence.AddFailure operationName & ": view has no displayed name."
        Exit Function
    End If

    Dim activateResult As Boolean
    activateResult = swDraw.ActivateView(viewName)

    Dim activeView As SldWorks.View
    Set activeView = swDraw.ActiveDrawingView

    Dim activeViewMatches As Boolean
    activeViewMatches = False
    If Not activeView Is Nothing Then
        activeViewMatches = (StrComp( _
            GetViewName(activeView), viewName, vbTextCompare) = 0)
    End If

    If Not activeViewMatches Then

        swDrawModel.ClearSelection2 True

        Dim selectedView As Boolean
        selectedView = swDrawModel.Extension.SelectByID2( _
            viewName, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0)

        If selectedView Then
            activateResult = swDraw.ActivateView(viewName)
            Set activeView = swDraw.ActiveDrawingView
        End If

        swDrawModel.ClearSelection2 True
        evidence.AddInfo "ACTIVATE_VIEW_RETRY|operation=" & operationName & _
            "|view=" & viewName & _
            "|selectedByID=" & CStr(selectedView) & _
            "|setterResult=" & CStr(activateResult)
    End If

    If activeView Is Nothing Then
        evidence.AddFailure operationName & _
            ": ActivateView returned " & CStr(activateResult) & _
            " and no drawing view is active for '" & viewName & "'."
        Exit Function
    End If

    If StrComp(GetViewName(activeView), viewName, vbTextCompare) <> 0 Then
        evidence.AddFailure operationName & _
            ": active-view readback mismatch; requested='" & viewName & _
            "', actual='" & GetViewName(activeView) & "'."
        Exit Function
    End If

    If Not activateResult Then
        evidence.AddWarning "ACTIVATE_VIEW|operation=" & operationName & _
            "|view=" & viewName & _
            "|setterResult=False|readbackMatched=True"
    End If

    ActivateDrawingView = True
    Exit Function

Failed:
    evidence.AddFailure operationName & ": ActivateView error " & _
        CStr(Err.Number) & " - " & Err.Description
End Function

Public Sub RestoreSheetContext( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc)

    On Error Resume Next
    If Not Module1_Main.GlobalEvidence Is Nothing Then
        Module1_Main.GlobalEvidence.RecordSolidWorksMutation _
            "RestoreSheetContext"
    End If

    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If Not swSheet Is Nothing Then
        swDraw.ActivateSheet swSheet.GetName
    End If
    On Error GoTo 0
End Sub

Public Function GetViewName(ByRef swView As SldWorks.View) As String
    On Error Resume Next
    GetViewName = swView.GetName2
    On Error GoTo 0
End Function

Public Function IsOrdinateEligibleView( _
    ByRef swView As SldWorks.View) As Boolean

    On Error GoTo NotEligible

    Dim viewType As Long
    viewType = swView.Type

    If viewType <> swDrawingProjectedView And _
       viewType <> swDrawingStandardView And _
       viewType <> swDrawingNamedView Then Exit Function

    If viewType = swDrawingNamedView Then
        If mProvenOrdinateViews Is Nothing Then Exit Function
        If Not mProvenOrdinateViews.Exists( _
            LCase$(GetViewName(swView))) Then Exit Function
    End If

    Dim orientationName As String
    orientationName = UCase$(Trim$(swView.GetOrientationName))

    If InStr(orientationName, "ISOMETRIC") > 0 Then Exit Function
    If InStr(orientationName, "TRIMETRIC") > 0 Then Exit Function
    If InStr(orientationName, "DIMETRIC") > 0 Then Exit Function

    Select Case orientationName
        Case "*FRONT", "FRONT", "*BACK", "BACK", _
             "*LEFT", "LEFT", "*RIGHT", "RIGHT", _
             "*TOP", "TOP", "*BOTTOM", "BOTTOM"
            IsOrdinateEligibleView = True
    End Select
    Exit Function

NotEligible:
    IsOrdinateEligibleView = False
End Function

Public Function MeasureControlledSheetRegions( _
    ByRef swSheet As SldWorks.Sheet, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed
    evidence.RequireStage "CONTROLLED_SHEET"
    evidence.LayoutBoundariesProven = False

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim paperSize As Long

    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)
    evidence.SheetWidth = sheetWidth
    evidence.SheetHeight = sheetHeight

    If evidence.SheetWidth <= 0# Or evidence.SheetHeight <= 0# Then
        evidence.AddFailure "Controlled sheet has invalid physical dimensions."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "invalid physical dimensions"
        Exit Function
    End If

    evidence.SheetTemplateName = Trim$(swSheet.GetTemplateName)
    evidence.SheetFormatName = Trim$(swSheet.GetSheetFormatName)

    If Len(evidence.SheetTemplateName) = 0 Then
        evidence.AddFailure "Controlled sheet has no template path/name."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "missing template path or name"
        Exit Function
    End If

    If Len(evidence.SheetFormatName) = 0 Then
        evidence.AddFailure "Controlled sheet has no sheet format."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "missing sheet format"
        Exit Function
    End If

    Dim formatVisible As Boolean
    formatVisible = swSheet.SheetFormatVisible

    If Not formatVisible Then
        evidence.RecordSolidWorksMutation "EnsureSheetFormatVisible"
        swSheet.SheetFormatVisible = True
        formatVisible = swSheet.SheetFormatVisible
    End If

    If Not formatVisible Then
        evidence.AddWarning _
            "ISheet.SheetFormatVisible remained False; continuing to structural " & _
            "template, title-block, margin, and usable-area proof."
    End If

    Dim titleBlock As SldWorks.TitleBlock
    Set titleBlock = swSheet.TitleBlock

    If titleBlock Is Nothing Then
        evidence.AddFailure "Controlled sheet has no ITitleBlock definition."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "missing ITitleBlock"
        Exit Function
    End If

    Dim titleX1 As Double
    Dim titleY1 As Double
    Dim titleX2 As Double
    Dim titleY2 As Double
    titleBlock.GetExtents titleX1, titleY1, titleX2, titleY2

    evidence.TitleBlockLeft = IIf(titleX1 < titleX2, titleX1, titleX2)
    evidence.TitleBlockRight = IIf(titleX1 > titleX2, titleX1, titleX2)
    evidence.TitleBlockBottom = IIf(titleY1 < titleY2, titleY1, titleY2)
    evidence.TitleBlockTop = IIf(titleY1 > titleY2, titleY1, titleY2)

    If evidence.TitleBlockRight <= evidence.TitleBlockLeft Or _
       evidence.TitleBlockTop <= evidence.TitleBlockBottom Then

        evidence.AddFailure "ITitleBlock.GetExtents returned invalid bounds."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "invalid title-block extents"
        Exit Function
    End If

    If evidence.TitleBlockLeft < 0# Or _
       evidence.TitleBlockBottom < 0# Or _
       evidence.TitleBlockRight > evidence.SheetWidth Or _
       evidence.TitleBlockTop > evidence.SheetHeight Then

        evidence.AddFailure "Measured title-block bounds lie outside the sheet."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "title-block extents outside sheet"
        Exit Function
    End If

    Dim topMargin As Double
    Dim bottomMargin As Double
    Dim rightMargin As Double
    Dim leftMargin As Double

    topMargin = swSheet.GetZoneMargin(swZoneTopMargin)
    bottomMargin = swSheet.GetZoneMargin(swZoneBottomMargin)
    rightMargin = swSheet.GetZoneMargin(swZoneRightMargin)
    leftMargin = swSheet.GetZoneMargin(swZoneLeftMargin)

    If topMargin <= 0# Or bottomMargin <= 0# Or _
       rightMargin <= 0# Or leftMargin <= 0# Then

        evidence.AddFailure "Controlled zoned-border margins are unavailable or zero."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "one or more zone margins are zero"
        Exit Function
    End If

    evidence.UsableLeft = leftMargin + LAYOUT_MARGIN_M
    evidence.UsableRight = evidence.SheetWidth - rightMargin - LAYOUT_MARGIN_M
    evidence.UsableTop = evidence.SheetHeight - topMargin - LAYOUT_MARGIN_M
    evidence.UsableBottom = evidence.TitleBlockTop + LAYOUT_MARGIN_M

    If evidence.UsableBottom < bottomMargin + LAYOUT_MARGIN_M Then
        evidence.UsableBottom = bottomMargin + LAYOUT_MARGIN_M
    End If

    If evidence.UsableRight <= evidence.UsableLeft Or _
       evidence.UsableTop <= evidence.UsableBottom Then

        evidence.AddFailure "Measured border/title-block reserve leaves no usable view area."
        evidence.MarkStageFailed "CONTROLLED_SHEET", _
            "no usable view area remains"
        Exit Function
    End If

    evidence.LayoutBoundariesProven = True
    evidence.AddInfo "Measured sheet/border/title-block regions from ISheet zone margins " & _
        "and ITitleBlock.GetExtents."
    evidence.MarkStageProved "CONTROLLED_SHEET", _
        "template, visible format, zone margins, title block, and usable bounds proved"
    MeasureControlledSheetRegions = True
    Exit Function

Failed:
    evidence.AddFailure "Controlled sheet measurement error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed "CONTROLLED_SHEET", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub FinalizeSelectionAndSheetState( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed
    evidence.RequireStage "FINAL_CLEANUP"
    evidence.FinalCleanupVerified = False
    evidence.RecordSolidWorksMutation "FinalizeSelectionAndSheetState"

    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        evidence.AddFailure "Final cleanup could not obtain the current sheet."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "current sheet is Nothing"
        Exit Sub
    End If

    Dim activateResult As Boolean
    activateResult = swDraw.ActivateSheet(swSheet.GetName)

    If Not activateResult Then
        Dim currentSheet As SldWorks.Sheet
        Set currentSheet = swDraw.GetCurrentSheet

        Dim currentNameMatches As Boolean
        currentNameMatches = False
        If Not currentSheet Is Nothing Then
            currentNameMatches = (StrComp(currentSheet.GetName, _
                swSheet.GetName, vbTextCompare) = 0)
        End If

        If currentNameMatches And swDraw.ActiveDrawingView Is Nothing Then
            evidence.AddWarning "FINAL_SHEET_CONTEXT|ActivateSheet=False|" & _
                "readbackMatched=True|sheet=" & swSheet.GetName
        Else
            evidence.AddFailure "Final ActivateSheet returned False and sheet-context " & _
                "readback did not prove the requested sheet."
            evidence.MarkStageFailed "FINAL_CLEANUP", _
                "ActivateSheet false with mismatched context"
            Exit Sub
        End If
    End If

    VerifyFinalCleanupState swDrawModel, swDraw, evidence
    Exit Sub

Failed:
    evidence.AddFailure "Final selection/sheet cleanup error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.FinalCleanupVerified = False
    evidence.MarkStageFailed "FINAL_CLEANUP", _
        "cleanup API error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Public Function VerifyFinalCleanupState( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed
    evidence.RequireStage "FINAL_CLEANUP"

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        evidence.AddFailure "Final cleanup verification has no current sheet."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "current sheet is Nothing"
        Exit Function
    End If

    evidence.FinalSelectionCount = _
        swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)
    evidence.SheetContextRestored = (swDraw.ActiveDrawingView Is Nothing)

    If evidence.FinalSelectionCount <> 0 Then
        evidence.AddFailure "Final selection list is not empty."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "selection count=" & CStr(evidence.FinalSelectionCount)
        Exit Function
    End If

    If Not evidence.SheetContextRestored Then
        evidence.AddFailure "Final active context is not the drawing sheet."
        evidence.MarkStageFailed "FINAL_CLEANUP", _
            "active drawing view is not Nothing"
        Exit Function
    End If

    evidence.FinalCleanupVerified = True
    evidence.FinalCleanupInvalidated = False
    evidence.MarkStageProved "FINAL_CLEANUP", _
        "normal pick mode requested, sheet active, selection count zero"
    evidence.AddInfo "Final SetPickMode invoked; selection count=0."
    VerifyFinalCleanupState = True
    Exit Function

Failed:
    evidence.FinalCleanupVerified = False
    evidence.AddFailure "Final cleanup verification error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed "FINAL_CLEANUP", _
        "verification API error " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function ObjectsAreSame( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal firstObject As Object, _
    ByVal secondObject As Object) As Boolean

    If firstObject Is Nothing Then Exit Function
    If secondObject Is Nothing Then Exit Function

    On Error Resume Next
    ObjectsAreSame = (swApp.IsSame(firstObject, secondObject) = swObjectSame)
    On Error GoTo 0
End Function

Public Function CountVariantItems(ByVal items As Variant) As Long
    On Error GoTo Failed

    If IsEmpty(items) Or IsNull(items) Then Exit Function

    If IsArray(items) Then
        CountVariantItems = UBound(items) - LBound(items) + 1
    Else
        CountVariantItems = 1
    End If
    Exit Function

Failed:
    CountVariantItems = 0
End Function

Public Function SetConfiguredSheetScale( _
    ByRef swSheet As SldWorks.Sheet, _
    ByVal decimalScale As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    evidence.RequireStage "SHEET_SCALE"
    evidence.SheetScaleReadbackProven = False

    If decimalScale <= 0# Then
        evidence.AddFailure "Configured sheet scale is not positive."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "configured scale is not positive"
        Exit Function
    End If

    Dim numerator As Long
    Dim denominator As Long
    DecimalToRatio decimalScale, numerator, denominator
    evidence.RequestedScaleNumerator = numerator
    evidence.RequestedScaleDenominator = denominator

    On Error GoTo Failed
    evidence.RecordSolidWorksMutation "ISheet.SetScale"
    Dim setterResult As Boolean
    setterResult = swSheet.SetScale( _
        CDbl(numerator), CDbl(denominator), True, False)

    Dim sheetProperties As Variant
    sheetProperties = swSheet.GetProperties2

    If Not IsArray(sheetProperties) Then
        evidence.AddFailure "ISheet.GetProperties2 did not return a scale array."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "scale readback was not an array"
        SetConfiguredSheetScale = False
        Exit Function
    End If

    If UBound(sheetProperties) < 3 Then
        evidence.AddFailure "ISheet.GetProperties2 scale array is incomplete."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "scale readback contained fewer than four values"
        SetConfiguredSheetScale = False
        Exit Function
    End If

    evidence.ActualScaleNumerator = CDbl(sheetProperties(2))
    evidence.ActualScaleDenominator = CDbl(sheetProperties(3))

    If evidence.ActualScaleNumerator <= 0# Or _
       evidence.ActualScaleDenominator <= 0# Then

        evidence.AddFailure "Actual sheet-scale readback is not positive."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "nonpositive scale readback"
        SetConfiguredSheetScale = False
        Exit Function
    End If

    Dim actualDecimalScale As Double
    actualDecimalScale = evidence.ActualScaleNumerator / _
        evidence.ActualScaleDenominator

    If Abs(actualDecimalScale - decimalScale) > 0.000000001 Then
        evidence.AddFailure "Actual sheet scale does not match the requested scale."
        evidence.MarkStageFailed "SHEET_SCALE", _
            "requested=" & Format$(decimalScale, "0.000000000") & _
            "; actual=" & Format$(actualDecimalScale, "0.000000000")
        SetConfiguredSheetScale = False
        Exit Function
    End If

    evidence.SheetScaleReadbackProven = True
    SetConfiguredSheetScale = True
    evidence.MarkStageProved "SHEET_SCALE", _
        "requested and GetProperties2 scale ratios match"
    If Not setterResult Then
        evidence.AddWarning "SHEET_SCALE|setterResult=False|" & _
            "readbackMatched=True|ratio=" & _
            Format$(evidence.ActualScaleNumerator, "0.###") & ":" & _
            Format$(evidence.ActualScaleDenominator, "0.###")
    End If
    evidence.AddInfo "Sheet scale set and read back as " & _
        Format$(evidence.ActualScaleNumerator, "0.###") & ":" & _
        Format$(evidence.ActualScaleDenominator, "0.###") & "."
    Exit Function

Failed:
    evidence.AddFailure "ISheet.SetScale error " & CStr(Err.Number) & _
        ": " & Err.Description
    evidence.SheetScaleReadbackProven = False
    evidence.MarkStageFailed "SHEET_SCALE", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
    SetConfiguredSheetScale = False
End Function

Public Function RebuildDocumentVerified( _
    ByRef swModel As SldWorks.ModelDoc2, _
    ByVal operationName As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    evidence.RecordSolidWorksMutation operationName & " EditRebuild3"

    Dim rebuildResult As Boolean
    Dim rebuildStatus As Long
    rebuildResult = swModel.EditRebuild3
    rebuildStatus = swModel.Extension.NeedsRebuild2

    If rebuildStatus <> swModelRebuildStatus_FullyRebuilt Then
        evidence.AddFailure "REBUILD|operation=" & operationName & _
            "|setterResult=" & CStr(rebuildResult) & _
            "|status=" & CStr(rebuildStatus)
        Exit Function
    End If

    If Not rebuildResult Then
        evidence.AddWarning "REBUILD|operation=" & operationName & _
            "|setterResult=False|readbackFullyRebuilt=True"
    End If

    RebuildDocumentVerified = True
    Exit Function

Failed:
    evidence.AddFailure "Rebuild verification error for " & operationName & _
        ": " & CStr(Err.Number) & " - " & Err.Description
End Function

Private Sub DecimalToRatio( _
    ByVal value As Double, _
    ByRef numerator As Long, _
    ByRef denominator As Long)

    denominator = 1000
    numerator = CLng(value * CDbl(denominator) + 0.5)

    Dim divisor As Long
    divisor = GreatestCommonDivisor(Abs(numerator), denominator)

    If divisor > 0 Then
        numerator = numerator \ divisor
        denominator = denominator \ divisor
    End If

    If numerator < 1 Then numerator = 1
    If denominator < 1 Then denominator = 1
End Sub

Private Function GreatestCommonDivisor( _
    ByVal firstValue As Long, _
    ByVal secondValue As Long) As Long

    Dim remainder As Long
    Do While secondValue <> 0
        remainder = firstValue Mod secondValue
        firstValue = secondValue
        secondValue = remainder
    Loop

    GreatestCommonDivisor = firstValue
End Function

Public Function TransformPointToView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal z As Double, _
    ByRef viewX As Double, _
    ByRef viewY As Double, _
    ByRef viewZ As Double) As Boolean

    On Error GoTo Failed

    Dim mathUtil As SldWorks.MathUtility
    Set mathUtil = swApp.GetMathUtility

    Dim pointData(0 To 2) As Double
    pointData(0) = x
    pointData(1) = y
    pointData(2) = z

    Dim mathPoint As SldWorks.MathPoint
    Set mathPoint = mathUtil.CreatePoint(pointData)
    Set mathPoint = mathPoint.MultiplyTransform(swView.ModelToViewTransform)

    Dim result As Variant
    result = mathPoint.ArrayData

    viewX = CDbl(result(0))
    viewY = CDbl(result(1))
    viewZ = CDbl(result(2))
    TransformPointToView = True
    Exit Function

Failed:
    TransformPointToView = False
End Function

Public Function TransformVectorToView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal z As Double, _
    ByRef viewX As Double, _
    ByRef viewY As Double, _
    ByRef viewZ As Double) As Boolean

    On Error GoTo Failed

    Dim mathUtil As SldWorks.MathUtility
    Set mathUtil = swApp.GetMathUtility

    Dim vectorData(0 To 2) As Double
    vectorData(0) = x
    vectorData(1) = y
    vectorData(2) = z

    Dim mathVector As SldWorks.MathVector
    Set mathVector = mathUtil.CreateVector(vectorData)
    Set mathVector = mathVector.MultiplyTransform(swView.ModelToViewTransform)

    Dim result As Variant
    result = mathVector.ArrayData

    viewX = CDbl(result(0))
    viewY = CDbl(result(1))
    viewZ = CDbl(result(2))
    TransformVectorToView = True
    Exit Function

Failed:
    TransformVectorToView = False
End Function

Public Sub ViewToSheetCoordinates( _
    ByRef swView As SldWorks.View, _
    ByVal viewX As Double, _
    ByVal viewY As Double, _
    ByRef sheetX As Double, _
    ByRef sheetY As Double)

    Dim position As Variant
    position = swView.Position

    sheetX = CDbl(position(0)) + viewX
    sheetY = CDbl(position(1)) + viewY
End Sub

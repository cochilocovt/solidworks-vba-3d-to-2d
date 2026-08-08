
Option Explicit

Public Function BuildRunSummary( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal holes As Collection, _
    ByVal importedModelItems As Long, _
    ByRef ordinateStatus As Module5_FallbackDimensionEngine.OrdinateRunStatus) As String

    Dim report As String
    Dim totalDims As Long
    totalDims = CountAllViewDimensions(swDraw)

    Dim totalCallouts As Long
    totalCallouts = CountAllViewHoleCallouts(swDraw)

    Dim holeWizardCount As Long
    holeWizardCount = Module3_ModelAudit.CountHoleWizardItems(holes)

    Dim controlledCallouts As Long
    controlledCallouts = _
        Module9_HoleCalloutEngine.CreatedControlledCalloutCount()

    Dim nativeCoveredFamilies As Long
    nativeCoveredFamilies = _
        Module9_HoleCalloutEngine.NativeCoveredFamilyCount()

    Dim requiredCalloutFamilies As Long
    requiredCalloutFamilies = Module9_HoleCalloutEngine.RequiredFamilyCount()

    report = "Drawing QA Summary" & vbCrLf & String(26, "-") & vbCrLf

    ' Requested config first, before any produced number. A report that opens
    ' with output invites reading a zero as a failure when it may only be an
    ' unticked box.
    report = report & Module4_ModelItemImporter.DescribeRequestedConfig() & vbCrLf
    report = report & Module4_ModelItemImporter.DescribeLastImportTransactions() & vbCrLf
    If Module1_Main.GlobalConfig.UseModelDimensions And _
       Module1_Main.GlobalConfig.ImportHoleCallouts Then
        report = report & _
            Module9_HoleCalloutEngine.DescribeControlledCallouts() & vbCrLf
    End If
    report = report & Module10_SheetLayoutEngine.DescribeLayout() & vbCrLf

    report = report & "Detected hole-like features: " & Module3_ModelAudit.CountHoles(holes) & _
        " (Hole Wizard: " & holeWizardCount & ", plain cut: " & _
        (Module3_ModelAudit.CountHoles(holes) - holeWizardCount) & ")" & vbCrLf
    report = report & "Imported model items: " & importedModelItems & vbCrLf
    report = report & "Total drawing view dimensions: " & totalDims & vbCrLf
    report = report & "Hole callouts (IsHoleCallout=True) across drawing: " & _
        totalCallouts & " of " & totalDims & " dimensions" & vbCrLf
    report = report & "Controlled attached hole-callout notes: " & _
        controlledCallouts & " (required Hole Wizard families: " & _
        requiredCalloutFamilies & ")" & vbCrLf
    report = report & "Native hole-callout families proved by attached " & _
        "owned geometry: " & nativeCoveredFamilies & vbCrLf
    report = report & BuildPerViewSummary(swDraw)

    ' Ordinate outcome is reported separately from the dimension total, so a
    ' healthy imported-model-item count cannot mask a fully failed ordinate
    ' run. Previously chain failures reached only Debug.Print.
    If Module1_Main.GlobalConfig.UseOrdinateDims Then
        report = report & _
            Module5_FallbackDimensionEngine.DescribeOrdinateStatus(ordinateStatus)
    End If

    ' Only a warning when import was actually asked for. The r3 run reported
    ' "import returned zero items" while running in ordinate-only mode, where
    ' the importer is never called.
    If Not Module1_Main.GlobalConfig.UseModelDimensions Then
        report = report & "Model item import: skipped (not requested)." & vbCrLf
    ElseIf importedModelItems = 0 Then
        report = report & "WARNING: Model item import returned zero items." & vbCrLf
    End If

    ' Fail closed: ordinates were requested, chains were attempted, and none
    ' were created. That is a failed run whatever the dimension total says.
    Dim ordinatesRequestedButAbsent As Boolean
    ordinatesRequestedButAbsent = _
        Module1_Main.GlobalConfig.UseOrdinateDims And _
        (ordinateStatus.ChainsAttempted > 0) And _
        (ordinateStatus.ChainsCreated = 0)

    Dim calloutCoverageFailed As Boolean
    calloutCoverageFailed = _
        Module1_Main.GlobalConfig.UseModelDimensions And _
        Module1_Main.GlobalConfig.ImportHoleCallouts And _
        (holeWizardCount > 0) And _
        ((requiredCalloutFamilies <= 0) Or _
         ((nativeCoveredFamilies + controlledCallouts) < _
            requiredCalloutFamilies) Or _
         (Module9_HoleCalloutEngine.ControlledCalloutFailureCount() > 0))

    Dim sheetLayoutFailed As Boolean
    sheetLayoutFailed = Not Module10_SheetLayoutEngine.LastLayoutPassed

    ' Now that requested config is reported, a zero-callout run can be
    ' classified rather than merely observed. Only the first branch is a real
    ' API finding; the second is an operator setting and must not be reported
    ' as a defect.
    If totalCallouts = 0 And holeWizardCount > 0 Then
        If Module1_Main.GlobalConfig.ImportHoleCallouts Then
            If controlledCallouts > 0 Then
                report = report & "FINDING: InsertModelAnnotations4 produced " & _
                    "no native IsHoleCallout dimension; the typed, " & _
                    "geometry-attached controlled fallback created " & _
                    CStr(controlledCallouts) & " family callout(s)." & vbCrLf
            Else
                report = report & "FINDING: " & holeWizardCount & _
                    " Hole Wizard item(s) present and the native callout " & _
                    "bit WAS requested, yet neither a native nor a " & _
                    "controlled callout was proved. Not an operator " & _
                    "setting." & vbCrLf
            End If
        Else
            report = report & "NOTE: no hole callouts, but ImportHoleCallouts was off. " & _
                "Not evidence of an API problem." & vbCrLf
        End If
    End If

    If sheetLayoutFailed Then
        report = report & "FAIL: Sheet-aware placement did not prove " & _
            "content bounds, protected-region clearance, view-to-view " & _
            "clearance, and accepted-scale readback." & vbCrLf
    ElseIf totalDims = 0 Then
        report = report & "WARNING: No drawing dimensions were found. Review model-item import and fallback logic." & vbCrLf
    ElseIf ordinatesRequestedButAbsent Then
        report = report & "FAIL: Ordinate dimensions were requested but every chain failed. Review the ordinate warning above." & vbCrLf
    ElseIf calloutCoverageFailed Then
        report = report & "FAIL: Hole callout family coverage is incomplete " & _
            "or its typed definition/geometry attachment was not proved." & _
            vbCrLf
    ElseIf totalDims < 6 Then
        report = report & "WARNING: Drawing has very few dimensions. Manual review required." & vbCrLf
    Else
        report = report & "PASS: Drawing contains dimensions." & vbCrLf
    End If

    BuildRunSummary = report
End Function

Public Function CountAllViewDimensions(ByRef swDraw As SldWorks.DrawingDoc) As Long
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        CountAllViewDimensions = CountAllViewDimensions + Module4_ModelItemImporter.CountDisplayDimensionsInView(swView)
        Set swView = swView.GetNextView
    Loop
End Function

Public Function CountAllViewHoleCallouts(ByRef swDraw As SldWorks.DrawingDoc) As Long
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        CountAllViewHoleCallouts = CountAllViewHoleCallouts + Module4_ModelItemImporter.CountHoleCalloutsInView(swView)
        Set swView = swView.GetNextView
    Loop
End Function

' Per-view roster. Carries the raw IView.Type and IView.GetOrientationName
' returns, not just the classification, so the run doubles as the probe for
' those two members -- neither has been characterised on this build.
Private Function BuildPerViewSummary(ByRef swDraw As SldWorks.DrawingDoc) As String
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    BuildPerViewSummary = "View roster:" & vbCrLf

    Do While Not swView Is Nothing
        Dim role As Long
        role = Module8_ViewClassifier.ClassifyView(swView)

        Dim ordinatePolicy As String
        If Module8_ViewClassifier.AllowsOrdinateDimensions(role) Then
            ordinatePolicy = "ordinates=allowed"
        Else
            ordinatePolicy = "ordinates=skipped"
        End If

        Dim viewCalloutCount As Long
        viewCalloutCount = Module4_ModelItemImporter.CountHoleCalloutsInView(swView)

        BuildPerViewSummary = BuildPerViewSummary & "  " & _
            Module8_ViewClassifier.DescribeView(swView) & _
            " | " & ordinatePolicy & _
            " | " & _
            Module4_ModelItemImporter.CountDisplayDimensionsInView(swView) & _
            " dims, " & viewCalloutCount & " callouts" & vbCrLf

        ' Only where dimensions exist: a readback line per empty view is
        ' noise, and every failure so far has been in the ordinated view.
        If Module4_ModelItemImporter.CountDisplayDimensionsInView(swView) > 0 _
            Then
            BuildPerViewSummary = BuildPerViewSummary & "    " & _
                Module5_FallbackDimensionEngine.DescribeDimensionReadback( _
                    swView) & vbCrLf
        End If

        Set swView = swView.GetNextView
    Loop
End Function



' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub

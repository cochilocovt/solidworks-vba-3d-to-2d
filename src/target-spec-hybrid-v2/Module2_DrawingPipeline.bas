Option Explicit

Private Const swCreateSectionView_NotAligned As Long = 1
Private Const swDocDRAWING As Long = 3
Private Const swDrawingDetailView As Long = 3
Private Const swDrawingStandardView As Long = 6
Private Const swDrawingNamedView As Long = 7
Private Const swDetViewSTANDARD As Long = 0
Private Const swDetCircleCIRCLE As Long = 1
Private Const swSketchARC As Long = 1
Private Const swHIDDEN_GREYED As Long = 1
Private Const swHIDDEN As Long = 2
Private Const swSHADED As Long = 3

Public Sub CreateDrawing( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String, _
    ByRef evidence As CRunEvidence)

    RunDrawingPipeline swApp, swPart, partPath, templatePath, evidence
End Sub

Public Sub RunDrawingPipeline( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String, _
    ByRef evidence As CRunEvidence)

    On Error GoTo FailRun

    Dim finalizationStarted As Boolean
    Dim pipelineErrorNumber As Long
    Dim pipelineErrorDescription As String

    Dim modelHoleFeatures As Collection
    Set modelHoleFeatures = Module3_ModelAudit.GetAllHoleLikeFeatures(swPart)
    evidence.AddInfo "Model audit accepted " & _
        CStr(Module3_ModelAudit.CountHoles(modelHoleFeatures)) & _
        " semantic hole-producing features; sketch-point counts were not used."

    Dim swDrawModel As SldWorks.ModelDoc2
    Set swDrawModel = swApp.NewDocument(templatePath, 0, 0#, 0#)

    If swDrawModel Is Nothing Then
        evidence.AddFailure "NewDocument returned Nothing."
        GoTo FinishRun
    End If

    If swDrawModel.GetType <> swDocDRAWING Then
        evidence.AddFailure "NewDocument did not return a drawing document."
        GoTo FinishRun
    End If

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swDrawModel

    If swDraw Is Nothing Then
        evidence.AddFailure "The drawing-document interface could not be acquired."
        GoTo FinishRun
    End If

    If Not Module8_RuntimeSupport.ActivateDrawingDocument( _
        swApp, swDrawModel, evidence) Then GoTo FinishRun

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        evidence.AddFailure "New drawing has no current sheet."
        GoTo FinishRun
    End If

    If Not Module8_RuntimeSupport.MeasureControlledSheetRegions( _
        swSheet, evidence) Then

        If Not Module1_Main.DIAGNOSTIC_DRAWING_MODE Then GoTo FinishRun

        evidence.AddWarning _
            "DIAGNOSTIC_DRAWING_MODE: controlled-sheet preflight failed; " & _
            "continuing to create views for inspection. This run is not an " & _
            "acceptance result."
    End If

    If Not Module8_RuntimeSupport.SetConfiguredSheetScale( _
        swSheet, Module1_Main.GlobalConfig.SheetScale, evidence) Then

        If Not Module1_Main.DIAGNOSTIC_DRAWING_MODE Then GoTo FinishRun

        evidence.AddWarning _
            "DIAGNOSTIC_DRAWING_MODE: sheet-scale transaction failed; " & _
            "continuing with the template's existing scale for view inspection."
    End If

    Dim primaryView As SldWorks.View
    Dim detailSourceView As SldWorks.View
    Set primaryView = CreateConfiguredViews( _
        swDrawModel, swDraw, partPath, detailSourceView, evidence)

    If primaryView Is Nothing Then GoTo FinishRun

    If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, "Initial drawing", evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "initial drawing rebuild", evidence) Then GoTo FinishRun
    End If

    If Module1_Main.GlobalConfig.CreateSection Then
        Dim sectionCreated As Boolean
        sectionCreated = CreatePrimarySection( _
            swDrawModel, swDraw, primaryView, evidence)

        If Not sectionCreated Then
            If Not ContinueDiagnosticPipeline( _
                "primary section creation", evidence) Then GoTo FinishRun
        ElseIf Not Module8_RuntimeSupport.RebuildDocumentVerified( _
            swDrawModel, "Post-section", evidence) Then

            If Not ContinueDiagnosticPipeline( _
                "post-section rebuild", evidence) Then GoTo FinishRun
        End If
    End If

    If Not CreateRequiredDetails( _
        swApp, swPart, swDrawModel, swDraw, detailSourceView, _
        evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "required detail creation", evidence) Then GoTo FinishRun
    End If

    If Module1_Main.GetFixtureKey(partPath) = "P-0252-01-013" Then
        If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
            swDrawModel, "Post-detail", evidence) Then

            If Not ContinueDiagnosticPipeline( _
                "post-detail rebuild", evidence) Then GoTo FinishRun
        End If
    End If

    If Module1_Main.GlobalConfig.CreateIso Then
        If Not CreateIsometricView(swDraw, partPath, evidence) Then
            If Not ContinueDiagnosticPipeline( _
                "isometric view creation", evidence) Then GoTo FinishRun
        End If
    Else
        evidence.AddFailure "The fixed acceptance workflow requires an isometric view."
        GoTo FinishRun
    End If

    If Not Module9_LayoutEngine.ArrangeViewsInMeasuredGrid( _
        swDrawModel, swDraw, "Initial", evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "initial layout", evidence) Then GoTo FinishRun
    End If

    Dim drawingContextReady As Boolean
    drawingContextReady = Module8_RuntimeSupport.ActivateDrawingDocument( _
        swApp, swDrawModel, evidence)

    If drawingContextReady Then
        evidence.ImportedAnnotations = _
            Module4_ModelItemImporter.ImportModelItemsAcrossDrawing( _
                swApp, swDrawModel, swDraw, _
                Module8_RuntimeSupport.GetViewName(primaryView), evidence)

        evidence.ImportedDisplayDimensions = _
            Module4_ModelItemImporter.CountAllDisplayDimensions(swDraw)
        Module4_ModelItemImporter.RecordDisplayDimensionCounts swDraw, evidence

        Module5_FallbackDimensionEngine.AddMissingOrdinateDimensions _
            swApp, swDrawModel, swDraw, _
            modelHoleFeatures, Module1_Main.GlobalConfig.DatumOrigin, evidence
    Else
        If Not ContinueDiagnosticPipeline( _
            "drawing activation before dimensions", evidence) Then GoTo FinishRun
        evidence.AddWarning "DIAGNOSTIC_DRAWING_MODE: dimension stages were " & _
            "skipped because the drawing document context was not proved."
    End If

    If Module1_Main.GlobalConfig.AutoArrange Then
        Module4_ModelItemImporter.AutoArrangeAllDrawingDimensions _
            swDrawModel, swDraw, evidence
    End If

    Module7_TitleBlockEngine.PopulateTitleBlock _
        swPart, swDrawModel, swDraw, evidence

    If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, "Final content", evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "final content rebuild", evidence) Then GoTo FinishRun
    End If

    If Not Module9_LayoutEngine.ArrangeViewsInMeasuredGrid( _
        swDrawModel, swDraw, "Final", evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "final layout", evidence) Then GoTo FinishRun
    End If

    If Not Module7_TitleBlockEngine.AddRequiredManufacturingDefinitions( _
        swDrawModel, swDraw, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "manufacturing definitions", evidence) Then GoTo FinishRun
    End If

    If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, "Post-layout", evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "post-layout rebuild", evidence) Then GoTo FinishRun
    End If

FinishRun:
    On Error Resume Next
    FinalizeRunOnce swDrawModel, swDraw, evidence, finalizationStarted
    If Err.Number <> 0 Then
        MsgBox "Finalization failed before complete evidence could be emitted: " & _
            CStr(Err.Number) & ": " & Err.Description, _
            vbCritical, "Target-Spec Hybrid V2"
        Err.Clear
    End If
    On Error GoTo 0
    Exit Sub

FailRun:
    pipelineErrorNumber = Err.Number
    pipelineErrorDescription = Err.Description

    On Error Resume Next
    evidence.AddFailure "Pipeline error " & CStr(pipelineErrorNumber) & _
        ": " & pipelineErrorDescription
    On Error GoTo 0
    GoTo FinishRun
End Sub

Private Function ContinueDiagnosticPipeline( _
    ByVal failedOperation As String, _
    ByRef evidence As CRunEvidence) As Boolean

    If Not Module1_Main.DIAGNOSTIC_DRAWING_MODE Then Exit Function

    evidence.AddWarning "DIAGNOSTIC_DRAWING_MODE: " & failedOperation & _
        " failed; continuing independent generation stages for visual inspection."
    ContinueDiagnosticPipeline = True
End Function

Private Sub FinalizeRunOnce( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence, _
    ByRef finalizationStarted As Boolean)

    If finalizationStarted Then Exit Sub
    finalizationStarted = True

    If swDrawModel Is Nothing Then
        evidence.AddFailure "Finalization skipped: no drawing document exists."
    ElseIf swDraw Is Nothing Then
        evidence.AddFailure _
            "Finalization skipped: the drawing interface is unavailable."
    Else
        TryZoomDrawing swDrawModel, evidence
        TryFinalizeDrawingState swDrawModel, swDraw, evidence
        TryRunReadOnlyQa swDrawModel, swDraw, evidence
    End If

    TryEmitEvidence evidence
End Sub

Private Sub TryZoomDrawing( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    evidence.RecordSolidWorksMutation "ViewZoomtofit2"
    swDrawModel.ViewZoomtofit2
    Exit Sub

Failed:
    evidence.AddFailure "Final zoom error " & CStr(Err.Number) & _
        ": " & Err.Description
End Sub

Private Sub TryFinalizeDrawingState( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Module8_RuntimeSupport.FinalizeSelectionAndSheetState _
        swDrawModel, swDraw, evidence
    Exit Sub

Failed:
    evidence.AddFailure "Final selection/sheet cleanup error " & _
        CStr(Err.Number) & ": " & Err.Description
End Sub

Private Sub TryRunReadOnlyQa( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Module6_QAEngine.PerformFinalDrawingChecks _
        swDrawModel, swDraw, evidence
    Exit Sub

Failed:
    evidence.AddFailure "Read-only final QA error " & CStr(Err.Number) & _
        ": " & Err.Description
End Sub

Private Sub TryEmitEvidence(ByRef evidence As CRunEvidence)
    Dim reportErrorNumber As Long
    Dim reportErrorDescription As String

    On Error GoTo Failed

    Module6_QAEngine.EmitRunEvidence evidence
    Exit Sub

Failed:
    reportErrorNumber = Err.Number
    reportErrorDescription = Err.Description

    On Error Resume Next
    evidence.AddFailure "Evidence emission error " & _
        CStr(reportErrorNumber) & ": " & reportErrorDescription
    Debug.Print evidence.BuildText
    MsgBox "Evidence emission failed: " & CStr(reportErrorNumber) & _
        ": " & reportErrorDescription, vbCritical, _
        "Target-Spec Hybrid V2"
    On Error GoTo 0
End Sub

Private Function CreateConfiguredViews( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal partPath As String, _
    ByRef detailSourceView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As SldWorks.View

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim paperSize As Long
    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)

    Dim orthoMode As Long
    If Module1_Main.GlobalConfig.UseHLR Then
        orthoMode = swHIDDEN
    Else
        orthoMode = swHIDDEN_GREYED
    End If

    Dim primaryView As SldWorks.View
    Set primaryView = swDraw.CreateDrawViewFromModelView3( _
        partPath, "*Front", sheetWidth * 0.32, sheetHeight * 0.62, 0#)

    If primaryView Is Nothing Then
        evidence.AddFailure "Failed to create the primary *Front view."
        Exit Function
    End If

    If Not ConfigureOrthographicView(primaryView, orthoMode, evidence) Then
        Exit Function
    End If

    If Not ApplyFixturePrimaryViewRotation( _
        primaryView, Module1_Main.GetFixtureKey(partPath), evidence) Then
        Exit Function
    End If

    evidence.ViewsCreated = evidence.ViewsCreated + 1
    evidence.AddInfo "VIEW|role=Primary|orientation=*Front|eligibleOrdinate=True"

    Dim slot As Long
    Dim otherCreatedView As SldWorks.View
    slot = 1

    If Module1_Main.GlobalConfig.CreateTop Then
        If Not CreateNamedView(swDraw, partPath, "*Top", slot, _
            sheetWidth, sheetHeight, orthoMode, False, _
            otherCreatedView, evidence) Then Exit Function
        slot = slot + 1
    End If

    If Module1_Main.GlobalConfig.CreateBottom Then
        If Not CreateNamedView(swDraw, partPath, "*Bottom", slot, _
            sheetWidth, sheetHeight, orthoMode, False, _
            detailSourceView, evidence) Then Exit Function
        slot = slot + 1
    End If

    If Module1_Main.GlobalConfig.CreateRight Then
        If Not CreateNamedView(swDraw, partPath, "*Right", slot, _
            sheetWidth, sheetHeight, orthoMode, False, _
            otherCreatedView, evidence) Then Exit Function
        slot = slot + 1
    End If

    If Module1_Main.GlobalConfig.CreateLeft Then
        If Not CreateNamedView(swDraw, partPath, "*Left", slot, _
            sheetWidth, sheetHeight, orthoMode, False, _
            otherCreatedView, evidence) Then Exit Function
        slot = slot + 1
    End If

    If Module1_Main.GlobalConfig.CreateBack Then
        If Not CreateNamedView(swDraw, partPath, "*Back", slot, _
            sheetWidth, sheetHeight, orthoMode, False, _
            otherCreatedView, evidence) Then Exit Function
        slot = slot + 1
    End If

    evidence.AddInfo "VIEW_PLAN|fixture=" & _
        Module1_Main.GetFixtureKey(partPath) & _
        "|orthographicCount=" & CStr(evidence.ViewsCreated) & _
        "|rotationPolicy=P0251PrimaryClockwise90Only"

    Set CreateConfiguredViews = primaryView
    Exit Function

Failed:
    evidence.AddFailure "View creation error " & CStr(Err.Number) & _
        ": " & Err.Description
End Function

Private Function CreateNamedView( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal partPath As String, _
    ByVal modelViewName As String, _
    ByVal slot As Long, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByVal displayMode As Long, _
    ByVal isIsometric As Boolean, _
    ByRef createdView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As Boolean

    Set createdView = Nothing

    Dim xPos As Double
    Dim yPos As Double
    xPos = sheetWidth * (0.25 + (0.2 * CDbl(slot Mod 3)))
    yPos = sheetHeight * (0.68 - (0.22 * CDbl(slot \ 3)))

    Dim swView As SldWorks.View
    Set swView = swDraw.CreateDrawViewFromModelView3( _
        partPath, modelViewName, xPos, yPos, 0#)

    If swView Is Nothing Then
        evidence.AddFailure "Failed to create model view " & modelViewName & "."
        Exit Function
    End If

    If isIsometric Then
        If Not ConfigureIsometricView(swView, displayMode, evidence) Then
            Exit Function
        End If
    Else
        If Not ConfigureOrthographicView(swView, displayMode, evidence) Then
            Exit Function
        End If
    End If

    evidence.ViewsCreated = evidence.ViewsCreated + 1
    evidence.AddInfo "VIEW|role=" & IIf(isIsometric, "OrientationAid", "Orthographic") & _
        "|orientation=" & modelViewName & _
        "|eligibleOrdinate=" & CStr(Not isIsometric)
    Set createdView = swView
    CreateNamedView = True
End Function

Private Function CreateIsometricView( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal partPath As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then
        evidence.AddFailure "Cannot create isometric view without a current sheet."
        Exit Function
    End If

    Dim sheetWidth As Double
    Dim sheetHeight As Double
    Dim paperSize As Long
    paperSize = swSheet.GetSize(sheetWidth, sheetHeight)

    Dim isometricView As SldWorks.View

    CreateIsometricView = CreateNamedView( _
        swDraw, partPath, "*Isometric", 5, sheetWidth, sheetHeight, _
        swSHADED, True, isometricView, evidence)

    If CreateIsometricView Then
        If Module1_Main.GetFixtureKey(partPath) = "P-0251-14A-001" Then
            isometricView.UseSheetScale = 0
            isometricView.ScaleDecimal = 0.5

            If isometricView.UseSheetScale <> 0 Or _
               Abs(isometricView.ScaleDecimal - 0.5) > 0.000001 Then

                evidence.AddFailure "P-0251 isometric view rejected its approved " & _
                    "orientation-aid scale of 1:2."
                CreateIsometricView = False
                Exit Function
            End If

            evidence.AddInfo "VIEW_SCALE|role=OrientationAid|fixture=" & _
                "P-0251-14A-001|ratio=1:2|independent=True"
        End If

        evidence.AddInfo "VIEW_POLICY|isometricCreatedBeforeSelectedViewImport=True"
    End If
    Exit Function

Failed:
    evidence.AddFailure "Isometric view creation error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function ConfigureOrthographicView( _
    ByRef swView As SldWorks.View, _
    ByVal displayMode As Long, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    swView.UseSheetScale = 1
    If swView.UseSheetScale <> 1 Then
        evidence.AddFailure "Orthographic view rejected sheet-scale mode."
        Exit Function
    End If

    If Not ApplyDisplayModeWithReadback(swView, displayMode, _
        "orthographic", evidence) Then Exit Function

    Module8_RuntimeSupport.RegisterProvenOrdinateView swView, evidence
    ConfigureOrthographicView = True
    Exit Function

Failed:
    evidence.AddFailure "Orthographic view configuration error: " & Err.Description
End Function

Private Function ApplyFixturePrimaryViewRotation( _
    ByRef primaryView As SldWorks.View, _
    ByVal fixtureKey As String, _
    ByRef evidence As CRunEvidence) As Boolean

    Const P0251_PRIMARY_CLOCKWISE_90_RAD As Double = -1.5707963267949
    Const ANGLE_READBACK_TOLERANCE_RAD As Double = 0.000001

    On Error GoTo Failed

    If primaryView Is Nothing Then
        evidence.AddFailure "Primary-view rotation was requested without a view."
        Exit Function
    End If

    If fixtureKey <> "P-0251-14A-001" Then
        evidence.AddInfo "VIEW_ROTATION|fixture=" & fixtureKey & _
            "|role=Primary|requested=None|readback=Unchanged"
        ApplyFixturePrimaryViewRotation = True
        Exit Function
    End If

    Dim requestedAngle As Double
    requestedAngle = P0251_PRIMARY_CLOCKWISE_90_RAD

    evidence.AddInfo "VIEW_ROTATION|fixture=P-0251-14A-001" & _
        "|role=Primary|phase=BeforeSetter|requestedRad=" & _
        Format$(requestedAngle, "0.000000000")

    primaryView.Angle = requestedAngle

    Dim actualAngle As Double
    actualAngle = primaryView.Angle
    evidence.AddInfo "VIEW_ROTATION|fixture=P-0251-14A-001" & _
        "|role=Primary|phase=AfterSetter|requestedRad=" & _
        Format$(requestedAngle, "0.000000000") & _
        "|actualRad=" & Format$(actualAngle, "0.000000000")

    If Abs(actualAngle - requestedAngle) > ANGLE_READBACK_TOLERANCE_RAD Then
        evidence.AddFailure "P-0251 primary-view 90-degree rotation readback " & _
            "mismatch: requested=" & Format$(requestedAngle, "0.000000000") & _
            ", actual=" & Format$(actualAngle, "0.000000000") & "."
        Exit Function
    End If

    ApplyFixturePrimaryViewRotation = True
    Exit Function

Failed:
    evidence.AddFailure "P-0251 primary-view rotation error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function ConfigureIsometricView( _
    ByRef swView As SldWorks.View, _
    ByVal displayMode As Long, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    swView.UseSheetScale = 1
    If swView.UseSheetScale <> 1 Then
        evidence.AddFailure "Isometric view rejected sheet-scale mode."
        Exit Function
    End If

    If Not ApplyDisplayModeWithReadback(swView, displayMode, _
        "isometric", evidence) Then Exit Function
    ConfigureIsometricView = True
    Exit Function

Failed:
    evidence.AddFailure "Isometric view configuration error: " & Err.Description
End Function

Private Function ApplyDisplayModeWithReadback( _
    ByRef swView As SldWorks.View, _
    ByVal requestedMode As Long, _
    ByVal viewRole As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim setterResult As Boolean
    Dim actualMode As Long
    setterResult = swView.SetDisplayMode4( _
        False, requestedMode, False, True, True)
    actualMode = swView.GetDisplayMode2

    If actualMode <> requestedMode Then
        evidence.AddFailure "DISPLAY_MODE|view=" & _
            Module8_RuntimeSupport.GetViewName(swView) & _
            "|role=" & viewRole & _
            "|setterResult=" & CStr(setterResult) & _
            "|requested=" & CStr(requestedMode) & _
            "|actual=" & CStr(actualMode)
        Exit Function
    End If

    If Not setterResult Then
        evidence.AddWarning "DISPLAY_MODE|view=" & _
            Module8_RuntimeSupport.GetViewName(swView) & _
            "|role=" & viewRole & _
            "|setterResult=False|readbackMatched=True|mode=" & _
            CStr(actualMode)
    End If

    ApplyDisplayModeWithReadback = True
    Exit Function

Failed:
    evidence.AddFailure "Display-mode verification error for '" & _
        Module8_RuntimeSupport.GetViewName(swView) & "': " & Err.Description
End Function

Private Function CreateRequiredDetails( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef sourceView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If Module1_Main.GetFixtureKey(evidence.PartPath) <> "P-0252-01-013" Then
        evidence.MarkStageNotApplicable "REQUIRED_DETAILS_STRUCTURE", _
            "the authorized fixture reference does not require detail views"
        evidence.AddInfo "REQUIRED_DETAILS_GEOMETRY|status=NOT_APPLICABLE"
        evidence.AddInfo "REQUIRED_DETAILS_LEGIBILITY|status=NOT_APPLICABLE"
        CreateRequiredDetails = True
        Exit Function
    End If

    evidence.RequireStage "REQUIRED_DETAILS_STRUCTURE"

    If sourceView Is Nothing Then
        evidence.AddFailure _
            "Pump Holder Details C/D require the exact created *Bottom view."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "the created *Bottom role was not retained"
        Exit Function
    End If

    Dim sourceViewType As Long
    sourceViewType = sourceView.Type
    If sourceViewType <> swDrawingStandardView And _
       sourceViewType <> swDrawingNamedView Then

        evidence.AddFailure "Pump Holder detail source is not a supported " & _
            "standard or named orthographic view."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created *Bottom role has unsupported type=" & CStr(sourceViewType)
        Exit Function
    End If

    If StrComp( _
        Trim$(sourceView.GetOrientationName), "*Bottom", vbTextCompare) <> 0 Then

        evidence.AddFailure "Pump Holder detail source orientation is not " & _
            "exactly *Bottom."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created detail-source orientation mismatch"
        Exit Function
    End If

    Dim referencedPart As SldWorks.ModelDoc2
    Set referencedPart = sourceView.ReferencedDocument
    If referencedPart Is Nothing Then
        evidence.AddFailure "Pump Holder detail source has no referenced part."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created *Bottom view has no ReferencedDocument"
        Exit Function
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, referencedPart, swPart) Then

        evidence.AddFailure "Pump Holder detail source references the wrong model."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created *Bottom view does not reference the authorized part object"
        Exit Function
    End If

    If StrComp( _
        Trim$(sourceView.ReferencedConfiguration), _
        Trim$(evidence.ConfigurationName), _
        vbTextCompare) <> 0 Then

        evidence.AddFailure "Pump Holder detail source configuration mismatch."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created *Bottom view does not reference the approved configuration"
        Exit Function
    End If

    Dim currentSheet As SldWorks.Sheet
    Dim sourceSheet As SldWorks.Sheet
    Set currentSheet = swDraw.GetCurrentSheet
    Set sourceSheet = sourceView.Sheet
    If currentSheet Is Nothing Or sourceSheet Is Nothing Then
        evidence.AddFailure "Pump Holder detail source sheet ownership is missing."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "current or source sheet readback is Nothing"
        Exit Function
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, currentSheet, sourceSheet) Then

        evidence.AddFailure "Pump Holder detail source belongs to another sheet."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created *Bottom view sheet identity mismatch"
        Exit Function
    End If

    evidence.AddInfo "DETAIL_SOURCE_PROOF|role=CreatedBottom" & _
        "|orientation=*Bottom|view=" & _
        Module8_RuntimeSupport.GetViewName(sourceView) & _
        "|type=" & CStr(sourceViewType) & _
        "|configuration=" & sourceView.ReferencedConfiguration & _
        "|model=" & sourceView.GetReferencedModelName

    Dim leftX As Double
    Dim bottomY As Double
    Dim rightX As Double
    Dim topY As Double
    If Not TryGetViewOutline( _
        sourceView, leftX, bottomY, rightX, topY, _
        "DETAIL_SOURCE", evidence) Then

        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "source-view outline was not proved"
        Exit Function
    End If

    Dim sourceWidth As Double
    Dim sourceHeight As Double
    sourceWidth = rightX - leftX
    sourceHeight = topY - bottomY

    Dim profileRadiusSheet As Double
    profileRadiusSheet = sourceWidth * (5# / 62.85)

    Dim detailCX As Double
    Dim detailCY As Double
    Dim detailDX As Double
    Dim detailDY As Double
    detailCX = leftX + sourceWidth * (4# / 62.85)
    detailCY = bottomY + sourceHeight * (29.75 / 31#)
    detailDX = leftX + sourceWidth * (35.3 / 62.85)
    detailDY = bottomY + sourceHeight * (8.5 / 31#)

    Dim usableWidth As Double
    Dim usableHeight As Double
    usableWidth = evidence.UsableRight - evidence.UsableLeft
    usableHeight = evidence.UsableTop - evidence.UsableBottom

    evidence.AddInfo "DETAIL_PLAN|fixture=P-0252-01-013|source=*Bottom" & _
        "|C_reference_mm=4.00,29.75|D_reference_mm=35.30,8.50" & _
        "|profileRadiusReference_mm=5.00|scale=3:1"

    If Not CreateOneDetail( _
        swApp, swDrawModel, swDraw, sourceView, "C", detailCX, detailCY, _
        profileRadiusSheet, evidence.UsableLeft + usableWidth * 0.68, _
        evidence.UsableBottom + usableHeight * 0.72, evidence) Then

        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "Detail C creation or readback failed"
        Exit Function
    End If

    If Not CreateOneDetail( _
        swApp, swDrawModel, swDraw, sourceView, "D", detailDX, detailDY, _
        profileRadiusSheet, evidence.UsableLeft + usableWidth * 0.72, _
        evidence.UsableBottom + usableHeight * 0.45, evidence) Then

        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "Detail D creation or readback failed"
        Exit Function
    End If

    evidence.MarkStageProved "REQUIRED_DETAILS_STRUCTURE", _
        "Details C and D created from reference-led *Bottom profiles at independent 3:1 scale"
    evidence.AddWarning "REQUIRED_DETAILS_GEOMETRY|status=PENDING" & _
        "|evidenceLevel=E4|profileCoordinateAndFeatureContainment=Unproved"
    evidence.AddWarning "REQUIRED_DETAILS_LEGIBILITY|status=PENDING" & _
        "|evidenceLevel=E6_E7|7x4AndC0.5DimensionsCollisionsReadability=Unproved"
    CreateRequiredDetails = True
    Exit Function

Failed:
    evidence.AddFailure "Required-detail creation error " & _
        CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CreateOneDetail( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef sourceView As SldWorks.View, _
    ByVal detailLabel As String, _
    ByVal centreSheetX As Double, _
    ByVal centreSheetY As Double, _
    ByVal radiusSheet As Double, _
    ByVal placementSheetX As Double, _
    ByVal placementSheetY As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim profileWasCreated As Boolean
    Dim detailWasCreated As Boolean
    Dim structuralPostconditionsPassed As Boolean
    Dim acceptedProfileItemCount As Long

    If Not Module8_RuntimeSupport.ActivateDrawingView( _
        swDrawModel, swDraw, sourceView, evidence, _
        "Detail " & detailLabel & " source") Then
        GoTo SafeExit
    End If

    Dim activeSourceView As SldWorks.View
    Set activeSourceView = swDraw.ActiveDrawingView
    If activeSourceView Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " source activation read back no active drawing view."
        GoTo SafeExit
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, activeSourceView, sourceView) Then

        evidence.AddFailure "Detail " & detailLabel & _
            " source activation read back a different drawing view."
        GoTo SafeExit
    End If
    If StrComp( _
        Trim$(sourceView.GetOrientationName), "*Bottom", vbTextCompare) <> 0 Then

        evidence.AddFailure "Detail " & detailLabel & _
            " source orientation is not exactly *Bottom."
        GoTo SafeExit
    End If

    swDrawModel.ClearSelection2 True

    Dim initialSelectionManager As SldWorks.SelectionMgr
    Set initialSelectionManager = swDrawModel.SelectionManager
    If initialSelectionManager Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " initial selection manager is Nothing."
        GoTo SafeExit
    End If
    If initialSelectionManager.GetSelectedObjectCount2(-1) <> 0 Then
        evidence.AddFailure "Detail " & detailLabel & _
            " source activation cleanup left selected objects."
        GoTo SafeExit
    End If

    Dim centreViewX As Double
    Dim centreViewY As Double
    If Not SheetToViewSketchCoordinates( _
        sourceView, centreSheetX, centreSheetY, _
        centreViewX, centreViewY, "DETAIL_" & detailLabel, _
        evidence) Then GoTo SafeExit

    Dim xform As Variant
    xform = sourceView.GetXform
    If IsEmpty(xform) Or Not IsArray(xform) Then
        evidence.AddFailure "Detail " & detailLabel & _
            " source GetXform returned no scale."
        GoTo SafeExit
    End If
    If UBound(xform) - LBound(xform) + 1 < 3 Then
        evidence.AddFailure "Detail " & detailLabel & _
            " source GetXform returned fewer than three values."
        GoTo SafeExit
    End If

    Dim viewScale As Double
    viewScale = CDbl(xform(LBound(xform) + 2))
    If Not IsFiniteSheetValue(viewScale) Or viewScale <= 0# Then
        evidence.AddFailure "Detail " & detailLabel & _
            " source view scale is non-finite or non-positive."
        GoTo SafeExit
    End If

    Dim profileSegment As SldWorks.SketchSegment
    Set profileSegment = swDrawModel.SketchManager.CreateCircleByRadius( _
        centreViewX, centreViewY, 0#, radiusSheet / viewScale)
    If profileSegment Is Nothing Then
        evidence.AddFailure "CreateCircleByRadius returned Nothing for Detail " & _
            detailLabel & "."
        GoTo SafeExit
    End If
    profileWasCreated = True

    Dim requestedViewRadius As Double
    requestedViewRadius = radiusSheet / viewScale
    If Not ValidateCircularSketchProfile( _
        profileSegment, centreViewX, centreViewY, requestedViewRadius, _
        "Detail " & detailLabel & " created profile", evidence) Then

        GoTo SafeExit
    End If

    Dim selectionManager As SldWorks.SelectionMgr
    Set selectionManager = swDrawModel.SelectionManager
    If selectionManager Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " selection manager is Nothing."
        GoTo SafeExit
    End If

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionManager.CreateSelectData
    If selectData Is Nothing Then
        evidence.AddFailure "CreateSelectData returned Nothing for Detail " & _
            detailLabel & "."
        GoTo SafeExit
    End If
    Set selectData.View = sourceView

    Dim profileSelected As Boolean
    profileSelected = CBool(profileSegment.Select4(False, selectData))
    If profileSelected = False Then
        evidence.AddFailure "Detail " & detailLabel & _
            " profile Select4 returned False."
        GoTo SafeExit
    End If

    If selectionManager.GetSelectedObjectCount2(-1) <> 1 Then
        evidence.AddFailure "Detail " & detailLabel & _
            " profile selection count is not one."
        GoTo SafeExit
    End If

    Dim selectedObject As Object
    Set selectedObject = selectionManager.GetSelectedObject6(1, -1)
    If selectedObject Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " selected profile readback is Nothing."
        GoTo SafeExit
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, selectedObject, profileSegment) Then

        evidence.AddFailure "Detail " & detailLabel & _
            " selected profile readback does not match the created circle."
        GoTo SafeExit
    End If
    If selectionManager.GetSelectedObjectMark(1) <> 0 Then
        evidence.AddFailure "Detail " & detailLabel & _
            " selected profile has an unexpected selection mark."
        GoTo SafeExit
    End If

    Dim selectedView As SldWorks.View
    Set selectedView = selectionManager.GetSelectedObjectsDrawingView2(1, -1)
    If selectedView Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " profile has no drawing-view selection context."
        GoTo SafeExit
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, selectedView, sourceView) Then

        evidence.AddFailure "Detail " & detailLabel & _
            " profile belongs to the wrong source view."
        GoTo SafeExit
    End If

    Dim createdObject As Object
    Set createdObject = swDraw.CreateDetailViewAt4( _
        placementSheetX, placementSheetY, 0#, _
        swDetViewSTANDARD, 3#, 1#, detailLabel, _
        swDetCircleCIRCLE, True, False, False, 1)

    If createdObject Is Nothing Then
        evidence.AddFailure "CreateDetailViewAt4 returned Nothing for Detail " & _
            detailLabel & "."
        GoTo SafeExit
    End If
    detailWasCreated = True

    Dim detailView As SldWorks.View
    Set detailView = createdObject

    detailView.UseParentScale = False
    detailView.UseSheetScale = 0
    detailView.ScaleDecimal = 3#

    If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, "Detail " & detailLabel & " structural readback", _
        evidence) Then
        GoTo SafeExit
    End If

    If detailView.Type <> swDrawingDetailView Then
        evidence.AddFailure "Detail " & detailLabel & _
            " did not read back as a detail drawing-view type."
        GoTo SafeExit
    End If

    If Not ValidateDetailScale(detailView, detailLabel, evidence) Then
        GoTo SafeExit
    End If

    If Not ValidateDetailPlacement( _
        detailView, detailLabel, placementSheetX, placementSheetY, evidence) Then
        GoTo SafeExit
    End If

    Dim detailCircle As SldWorks.DetailCircle
    Set detailCircle = detailView.GetDetail
    If detailCircle Is Nothing Then
        evidence.AddFailure "IView.GetDetail returned Nothing for Detail " & _
            detailLabel & "."
        GoTo SafeExit
    End If

    If StrComp( _
        Trim$(detailCircle.GetLabel), detailLabel, vbTextCompare) <> 0 Then

        Dim labelSet As Boolean
        labelSet = CBool(detailCircle.SetLabel(detailLabel))
        If labelSet = False Then
            evidence.AddFailure "IDetailCircle.SetLabel returned False for Detail " & _
                detailLabel & "."
            GoTo SafeExit
        End If

        If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
            swDrawModel, "Detail " & detailLabel & " label readback", _
            evidence) Then
            GoTo SafeExit
        End If
    End If

    If detailCircle.GetStyle <> swDetViewSTANDARD Then
        evidence.AddFailure "Detail " & detailLabel & _
            " style readback is not standard."
        GoTo SafeExit
    End If
    If detailCircle.GetDisplay <> swDetCircleCIRCLE Then
        evidence.AddFailure "Detail " & detailLabel & _
            " profile display readback is not circular."
        GoTo SafeExit
    End If
    If detailCircle.NoOutline Then
        evidence.AddFailure "Detail " & detailLabel & _
            " unexpectedly suppresses its detail-view outline."
        GoTo SafeExit
    End If
    Dim hasFullOutline As Boolean
    hasFullOutline = CBool(detailCircle.HasFullOutline)
    If hasFullOutline = False Then
        evidence.AddFailure "Detail " & detailLabel & _
            " did not read back with a full outline."
        GoTo SafeExit
    End If
    If detailCircle.JaggedOutline Then
        evidence.AddFailure "Detail " & detailLabel & _
            " unexpectedly read back with a jagged outline."
        GoTo SafeExit
    End If

    If StrComp( _
        Trim$(detailCircle.GetLabel), detailLabel, vbTextCompare) <> 0 Then

        evidence.AddFailure "Detail label readback mismatch for '" & _
            detailLabel & "'."
        GoTo SafeExit
    End If

    Dim parentView As SldWorks.View
    Set parentView = detailCircle.GetView
    If parentView Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " has no source-view readback."
        GoTo SafeExit
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, parentView, sourceView) Then

        evidence.AddFailure "Detail " & detailLabel & _
            " source-view readback mismatch."
        GoTo SafeExit
    End If

    Dim detailViewReadback As SldWorks.View
    Set detailViewReadback = detailCircle.GetDetailView
    If detailViewReadback Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " generated-view readback is Nothing."
        GoTo SafeExit
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, detailViewReadback, detailView) Then

        evidence.AddFailure "Detail " & detailLabel & _
            " generated-view identity readback mismatch."
        GoTo SafeExit
    End If

    If detailCircle.GetProfileItemsCount <> 1 Then
        evidence.AddFailure "Detail " & detailLabel & _
            " profile count readback is not exactly one."
        GoTo SafeExit
    End If

    Dim profileItems As Variant
    profileItems = detailCircle.GetProfileItems
    If Not IsArray(profileItems) Or _
       Module8_RuntimeSupport.CountVariantItems(profileItems) <> 1 Then

        evidence.AddFailure "Detail " & detailLabel & _
            " profile-item readback is not a one-item array."
        GoTo SafeExit
    End If

    Dim profileReadback As SldWorks.SketchSegment
    Set profileReadback = profileItems(LBound(profileItems))
    If profileReadback Is Nothing Then
        evidence.AddFailure "Detail " & detailLabel & _
            " profile-item readback is Nothing."
        GoTo SafeExit
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, profileReadback, profileSegment) Then

        evidence.AddFailure "Detail " & detailLabel & _
            " profile-item identity readback mismatch."
        GoTo SafeExit
    End If
    If Not ValidateCircularSketchProfile( _
        profileReadback, centreViewX, centreViewY, requestedViewRadius, _
        "Detail " & detailLabel & " consumed profile", evidence) Then

        GoTo SafeExit
    End If

    acceptedProfileItemCount = _
        Module8_RuntimeSupport.CountVariantItems(profileItems)
    structuralPostconditionsPassed = True

SafeExit:
    Module8_RuntimeSupport.RestoreSheetContext swDrawModel, swDraw

    Dim cleanupSelectionManager As SldWorks.SelectionMgr
    Dim cleanupActiveView As SldWorks.View
    Dim cleanupErrorNumber As Long
    Dim cleanupErrorDescription As String
    On Error Resume Next
    Set cleanupSelectionManager = swDrawModel.SelectionManager
    If cleanupSelectionManager Is Nothing Then
        cleanupErrorNumber = -1
        cleanupErrorDescription = "selection manager is Nothing"
    ElseIf cleanupSelectionManager.GetSelectedObjectCount2(-1) <> 0 Then
        cleanupErrorNumber = -2
        cleanupErrorDescription = "selection count is not zero"
    End If
    Set cleanupActiveView = swDraw.ActiveDrawingView
    If Not cleanupActiveView Is Nothing Then
        cleanupErrorNumber = -3
        cleanupErrorDescription = "a drawing view remains active"
    End If
    If Err.Number <> 0 Then
        cleanupErrorNumber = Err.Number
        cleanupErrorDescription = Err.Description
    End If
    Err.Clear
    On Error GoTo 0

    If cleanupErrorNumber <> 0 Then
        evidence.AddFailure "Detail " & detailLabel & _
            " cleanup postcondition failed: " & cleanupErrorDescription & "."
    End If

    If structuralPostconditionsPassed And cleanupErrorNumber = 0 Then
        evidence.AddInfo "DETAIL_RESULT|label=" & detailLabel & _
            "|created=True|sourceView=" & _
            Module8_RuntimeSupport.GetViewName(sourceView) & _
            "|scale=3:1|profileItemCount=" & _
            CStr(acceptedProfileItemCount) & "|cleanup=True"
        evidence.ViewsCreated = evidence.ViewsCreated + 1
        CreateOneDetail = True
    ElseIf profileWasCreated And Not detailWasCreated Then
        evidence.AddFailure "DETAIL_ORPHAN_PROFILE|label=" & detailLabel & _
            "|drawingRejected=True|automaticRollback=NotLiveProved"
    ElseIf detailWasCreated Then
        evidence.AddFailure "DETAIL_REJECTED_VIEW|label=" & detailLabel & _
            "|drawingRejected=True|reason=PostconditionOrCleanupFailed"
    End If
    Exit Function

Failed:
    evidence.AddFailure "Detail " & detailLabel & " creation error " & _
        CStr(Err.Number) & ": " & Err.Description
    Resume SafeExit
End Function

Private Function ValidateCircularSketchProfile( _
    ByRef profileSegment As SldWorks.SketchSegment, _
    ByVal expectedX As Double, _
    ByVal expectedY As Double, _
    ByVal expectedRadius As Double, _
    ByVal contextName As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If profileSegment.GetType <> swSketchARC Then
        evidence.AddFailure contextName & _
            " is not a sketch-arc segment."
        Exit Function
    End If

    Dim profileArc As SldWorks.SketchArc
    Set profileArc = profileSegment
    If profileArc Is Nothing Then
        evidence.AddFailure contextName & _
            " could not be read through ISketchArc."
        Exit Function
    End If
    If profileArc.IsCircle <> 1 Then
        evidence.AddFailure contextName & _
            " is a partial circular arc rather than a complete circle."
        Exit Function
    End If

    Dim profileCurve As SldWorks.Curve
    Set profileCurve = profileSegment.GetCurve
    If profileCurve Is Nothing Then
        evidence.AddFailure contextName & " has no underlying curve."
        Exit Function
    End If
    Dim profileIsCircular As Boolean
    profileIsCircular = CBool(profileCurve.IsCircle)
    If profileIsCircular = False Then
        evidence.AddFailure contextName & " is not circular."
        Exit Function
    End If

    Dim circleParameters As Variant
    circleParameters = profileCurve.CircleParams
    If Not IsArray(circleParameters) Then
        evidence.AddFailure contextName & " returned no CircleParams array."
        Exit Function
    End If
    If Module8_RuntimeSupport.CountVariantItems(circleParameters) <> 7 Then
        evidence.AddFailure contextName & _
            " CircleParams does not contain exactly seven values."
        Exit Function
    End If

    Dim firstParameter As Long
    firstParameter = LBound(circleParameters)

    Dim axisX As Double
    Dim axisY As Double
    Dim axisZ As Double
    Dim axisLength As Double
    axisX = CDbl(circleParameters(firstParameter + 3))
    axisY = CDbl(circleParameters(firstParameter + 4))
    axisZ = CDbl(circleParameters(firstParameter + 5))
    axisLength = Sqr(axisX * axisX + axisY * axisY + axisZ * axisZ)

    If Not IsFiniteSheetValue(axisLength) Or _
       Abs(axisLength - 1#) > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Or _
       Abs(axisX) > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Or _
       Abs(axisY) > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Or _
       Abs(Abs(axisZ) - 1#) > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then

        evidence.AddFailure contextName & _
            " circle axis is not a unit normal to the view sketch plane."
        Exit Function
    End If

    If Abs(CDbl(circleParameters(firstParameter)) - expectedX) > _
       Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Or _
       Abs(CDbl(circleParameters(firstParameter + 1)) - expectedY) > _
       Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Or _
       Abs(CDbl(circleParameters(firstParameter + 2))) > _
       Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Or _
       Abs(Abs(CDbl(circleParameters(firstParameter + 6))) - _
           expectedRadius) > Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M Then

        evidence.AddFailure contextName & _
            " CircleParams centre or radius readback mismatch."
        Exit Function
    End If

    If expectedRadius <= 0# Then
        evidence.AddFailure contextName & " requested radius is not positive."
        Exit Function
    End If

    ValidateCircularSketchProfile = True
    Exit Function

Failed:
    evidence.AddFailure contextName & " geometry-readback error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function ValidateDetailScale( _
    ByRef detailView As SldWorks.View, _
    ByVal detailLabel As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If detailView.UseParentScale Then
        evidence.AddFailure "Detail " & detailLabel & _
            " still inherits its parent scale."
        Exit Function
    End If
    If detailView.UseSheetScale <> 0 Or _
       Abs(detailView.ScaleDecimal - 3#) > 0.000001 Then

        evidence.AddFailure "Detail " & detailLabel & _
            " did not read back at independent 3:1 scale."
        Exit Function
    End If

    Dim scaleRatio As Variant
    scaleRatio = detailView.ScaleRatio
    If Not IsArray(scaleRatio) Or _
       Module8_RuntimeSupport.CountVariantItems(scaleRatio) <> 2 Then

        evidence.AddFailure "Detail " & detailLabel & _
            " ScaleRatio readback is not a two-value array."
        Exit Function
    End If

    Dim ratioBase As Long
    ratioBase = LBound(scaleRatio)
    If Abs(CDbl(scaleRatio(ratioBase)) - 3#) > 0.000001 Or _
       Abs(CDbl(scaleRatio(ratioBase + 1)) - 1#) > 0.000001 Then

        evidence.AddFailure "Detail " & detailLabel & _
            " ScaleRatio readback is not exactly 3:1."
        Exit Function
    End If

    ValidateDetailScale = True
    Exit Function

Failed:
    evidence.AddFailure "Detail " & detailLabel & _
        " scale-readback error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function ValidateDetailPlacement( _
    ByRef detailView As SldWorks.View, _
    ByVal detailLabel As String, _
    ByVal expectedSheetX As Double, _
    ByVal expectedSheetY As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim positionData As Variant
    positionData = detailView.Position
    If Not IsArray(positionData) Or _
       Module8_RuntimeSupport.CountVariantItems(positionData) <> 2 Then

        evidence.AddFailure "Detail " & detailLabel & _
            " Position readback is not a two-value array."
        Exit Function
    End If

    Dim positionBase As Long
    positionBase = LBound(positionData)
    If Abs(CDbl(positionData(positionBase)) - expectedSheetX) > _
       Module8_RuntimeSupport.PROJECTED_TOLERANCE_M Or _
       Abs(CDbl(positionData(positionBase + 1)) - expectedSheetY) > _
       Module8_RuntimeSupport.PROJECTED_TOLERANCE_M Then

        evidence.AddFailure "Detail " & detailLabel & _
            " initial sheet-position readback mismatch."
        Exit Function
    End If

    ValidateDetailPlacement = True
    Exit Function

Failed:
    evidence.AddFailure "Detail " & detailLabel & _
        " position-readback error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CreatePrimarySection( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef primaryView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim sectionStep As String
    sectionStep = "Entry"

    If swDrawModel Is Nothing Or swDraw Is Nothing Or primaryView Is Nothing Then
        evidence.AddFailure "SECTION_STEP|step=" & sectionStep & _
            "|status=Failed|reason=RequiredDrawingContextIsNothing"
        Exit Function
    End If

    If Module1_Main.GlobalSectionCount <> 1 Then
        evidence.AddFailure "Exactly one primary section configuration was expected."
        Exit Function
    End If

    If Not Module8_RuntimeSupport.ActivateDrawingView( _
        swDrawModel, swDraw, primaryView, evidence, _
        "Section source") Then Exit Function

    swDrawModel.ClearSelection2 True

    Dim leftX As Double
    Dim bottomY As Double
    Dim rightX As Double
    Dim topY As Double

    If Not TryGetViewOutline( _
        primaryView, leftX, bottomY, rightX, topY, _
        "SECTION_SOURCE", evidence) Then
        GoTo SafeExit
    End If

    Dim width As Double
    Dim height As Double
    width = rightX - leftX
    height = topY - bottomY

    If width <= 0# Or height <= 0# Then
        evidence.AddFailure "Section source view has an invalid outline."
        GoTo SafeExit
    End If

    Dim sectionLabel As String
    sectionLabel = Module1_Main.GetSectionLabelOrDefault(1)

    Dim fixtureKey As String
    fixtureKey = Module1_Main.GetFixtureKey(evidence.PartPath)

    Select Case fixtureKey
        Case "P-0251-14A-001"
            sectionLabel = "J"
        Case "P-0252-01-013"
            sectionLabel = "B"
        Case Else
            evidence.AddFailure "No approved primary-section plan exists for " & _
                fixtureKey & "."
            GoTo SafeExit
    End Select

    Dim extension As Double
    extension = 0.1 * IIf(width > height, width, height)

    Dim sectionSegments As Collection
    Set sectionSegments = New Collection

    Dim targetX As Double
    Dim targetY As Double
    Dim offsetX As Double

    If fixtureKey = "P-0251-14A-001" Then
        targetX = leftX + width * 0.5
        offsetX = leftX + width * (15# / 72#)
        targetY = bottomY + height * (90# / 196#)

        If Not AddSectionSegment(swDrawModel, primaryView, sectionSegments, _
            targetX, topY + extension, targetX, targetY, evidence) Then GoTo SafeExit
        If Not AddSectionSegment(swDrawModel, primaryView, sectionSegments, _
            targetX, targetY, offsetX, targetY, evidence) Then GoTo SafeExit
        If Not AddSectionSegment(swDrawModel, primaryView, sectionSegments, _
            offsetX, targetY, offsetX, bottomY - extension, evidence) Then GoTo SafeExit

        evidence.AddInfo "SECTION_PLAN|label=J|intent=SteppedBoreAndLeftFaceHoleColumn" & _
            "|segmentCount=3|boreX=" & Format$(targetX, "0.000000") & _
            "|offsetX=" & Format$(offsetX, "0.000000") & _
            "|jogY=" & Format$(targetY, "0.000000")
    Else
        targetY = bottomY + height * (15.84 / 24#)

        If Not AddSectionSegment(swDrawModel, primaryView, sectionSegments, _
            leftX - extension, targetY, rightX + extension, targetY, evidence) Then
            GoTo SafeExit
        End If

        evidence.AddInfo "SECTION_PLAN|label=B|intent=UpperEarThreadedHoles" & _
            "|segmentCount=1|cutY=" & Format$(targetY, "0.000000")
    End If

    swDrawModel.ClearSelection2 True

    Dim selectionManager As SldWorks.SelectionMgr
    sectionStep = "SelectionManager.Acquire.Before"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep
    Set selectionManager = swDrawModel.SelectionManager
    sectionStep = "SelectionManager.Acquire.After"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep & _
        "|isNothing=" & CStr(selectionManager Is Nothing)
    If selectionManager Is Nothing Then
        evidence.AddFailure "Section selection manager is Nothing."
        GoTo SafeExit
    End If

    Dim selectData As SldWorks.SelectData
    sectionStep = "CreateSelectData.Before"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep
    Set selectData = selectionManager.CreateSelectData
    sectionStep = "CreateSelectData.After"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep & _
        "|isNothing=" & CStr(selectData Is Nothing)
    If selectData Is Nothing Then
        evidence.AddFailure "CreateSelectData returned Nothing for section selection."
        GoTo SafeExit
    End If

    sectionStep = "SelectData.ViewAssignment.Before"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep
    ' The source drawing view is already active before the sketch segments are
    ' created.  Although ISelectData.View is a get/set property in the API,
    ' assigning it to these newly-created drawing-view sketch segments raises
    ' runtime error 91 in the installed VBA host.  Keep the selection data
    ' unbound and prove ownership after Select4 instead.
    sectionStep = "SelectData.ViewAssignment.Skipped"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep & _
        "|reason=ActiveSourceViewOwnsNewSketchSegments"

    Dim segmentIndex As Long
    Dim sectionSegment As SldWorks.SketchSegment
    Dim sectionSegmentSelected As Boolean
    For segmentIndex = 1 To sectionSegments.Count
        Set sectionSegment = sectionSegments(segmentIndex)

        If sectionSegment Is Nothing Then
            evidence.AddFailure "Section segment is Nothing at index " & _
                CStr(segmentIndex) & "."
            GoTo SafeExit
        End If

        sectionStep = "Select4.Before.Index=" & CStr(segmentIndex)
        evidence.AddInfo "SECTION_STEP|step=" & sectionStep
        sectionSegmentSelected = CBool( _
            sectionSegment.Select4(segmentIndex > 1, selectData))
        If sectionSegmentSelected = False Then
            sectionStep = "Select4.After.Index=" & CStr(segmentIndex) & _
                ".False"
            evidence.AddInfo "SECTION_STEP|step=" & sectionStep
            evidence.AddFailure "Section segment Select4 returned False at index " & _
                CStr(segmentIndex) & "."
            GoTo SafeExit
        End If
        sectionStep = "Select4.After.Index=" & CStr(segmentIndex) & ".True"
        evidence.AddInfo "SECTION_STEP|step=" & sectionStep
    Next segmentIndex

    sectionStep = "VerifySectionSelection.Before"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep
    If Not VerifySectionSelection( _
        selectionManager, sectionSegments, primaryView, evidence) Then
        sectionStep = "VerifySectionSelection.After.False"
        evidence.AddInfo "SECTION_STEP|step=" & sectionStep
        GoTo SafeExit
    End If
    sectionStep = "VerifySectionSelection.After.True"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep

    Dim sectionView As SldWorks.View
    sectionStep = "CreateSectionViewAt5.Before"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep
    Set sectionView = swDraw.CreateSectionViewAt5( _
        rightX + width * 0.75, targetY, 0#, sectionLabel, _
        swCreateSectionView_NotAligned, Empty, 0#)
    sectionStep = "CreateSectionViewAt5.After"
    evidence.AddInfo "SECTION_STEP|step=" & sectionStep & _
        "|isNothing=" & CStr(sectionView Is Nothing)

    If sectionView Is Nothing Then
        evidence.AddFailure "CreateSectionViewAt5 returned Nothing."
        GoTo SafeExit
    End If

    sectionView.UseSheetScale = 1
    If sectionView.UseSheetScale <> 1 Then
        evidence.AddFailure "Created section view rejected sheet-scale mode."
        GoTo SafeExit
    End If

    Dim sectionData As SldWorks.DrSection
    Set sectionData = sectionView.GetSection
    If sectionData Is Nothing Then
        evidence.AddFailure "Created section view did not expose section data."
        GoTo SafeExit
    End If

    Dim labelStatus As Long
    labelStatus = sectionData.SetLabel2(sectionLabel)
    If labelStatus <> 0 Then
        evidence.AddFailure "IDrSection.SetLabel2 returned status " & _
            CStr(labelStatus) & " for label '" & sectionLabel & "'."
        GoTo SafeExit
    End If

    Dim actualLabel As String
    actualLabel = Trim$(sectionData.GetLabel)
    If StrComp(actualLabel, sectionLabel, vbTextCompare) <> 0 Then
        evidence.AddFailure "Section label readback mismatch: requested='" & _
            sectionLabel & "', actual='" & actualLabel & "'."
        GoTo SafeExit
    End If

    Dim sectionLineInfoSize As Long
    Dim sectionLineCount As Long
    sectionLineCount = primaryView.GetSectionLineCount2(sectionLineInfoSize)
    If sectionLineCount < 1 Or sectionLineInfoSize < 1 Then
        evidence.AddFailure "The source view did not report a section line."
        GoTo SafeExit
    End If

    Dim sectionLineInfo As Variant
    sectionLineInfo = primaryView.GetSectionLineInfo2
    If IsEmpty(sectionLineInfo) Or Not IsArray(sectionLineInfo) Then
        evidence.AddFailure "GetSectionLineInfo2 returned no section geometry."
        GoTo SafeExit
    End If

    evidence.AddInfo "SECTION_RESULT|label=" & sectionLabel & _
        "|created=True|sourceView=" & _
        Module8_RuntimeSupport.GetViewName(primaryView) & _
        "|view=" & Module8_RuntimeSupport.GetViewName(sectionView) & _
        "|selectionCount=" & CStr(sectionSegments.Count) & _
        "|sectionLineCount=" & CStr(sectionLineCount) & _
        "|sectionLineInfoSize=" & CStr(sectionLineInfoSize) & _
        "|postconditions=LabelAndSectionLineInfoReadback"
    evidence.AddWarning "SECTION_TRANSACTION|status=PENDING" & _
        "|evidenceLevel=E4_E6|coordinateSelectionCreationReadback=Unproved"
    evidence.AddWarning "SECTION_VISUAL|status=PENDING" & _
        "|evidenceLevel=E7|featureIntersectionArrowsHatchLegibility=Unproved"

    evidence.ViewsCreated = evidence.ViewsCreated + 1
    CreatePrimarySection = True

SafeExit:
    Module8_RuntimeSupport.RestoreSheetContext swDrawModel, swDraw
    Exit Function

Failed:
    evidence.AddFailure "Section creation error " & CStr(Err.Number) & _
        ": " & Err.Description & "|SECTION_STEP=" & sectionStep
    Resume SafeExit
End Function

Private Function TryGetViewOutline( _
    ByRef sourceView As SldWorks.View, _
    ByRef leftX As Double, _
    ByRef bottomY As Double, _
    ByRef rightX As Double, _
    ByRef topY As Double, _
    ByVal evidenceContext As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim evidencePrefix As String
    evidencePrefix = UCase$(Trim$(evidenceContext))
    If Len(evidencePrefix) = 0 Then evidencePrefix = "VIEW_SOURCE"

    Dim outline As Variant
    outline = sourceView.GetOutline

    If IsEmpty(outline) Or Not IsArray(outline) Then
        evidence.AddFailure evidencePrefix & _
            ": IView.GetOutline returned no source bounds."
        Exit Function
    End If

    Dim lowerIndex As Long
    lowerIndex = LBound(outline)

    If UBound(outline) - lowerIndex + 1 <> 4 Then
        evidence.AddFailure evidencePrefix & _
            ": IView.GetOutline did not return exactly four values."
        Exit Function
    End If

    leftX = CDbl(outline(lowerIndex))
    bottomY = CDbl(outline(lowerIndex + 1))
    rightX = CDbl(outline(lowerIndex + 2))
    topY = CDbl(outline(lowerIndex + 3))

    If Not IsFiniteSheetValue(leftX) Or _
       Not IsFiniteSheetValue(bottomY) Or _
       Not IsFiniteSheetValue(rightX) Or _
       Not IsFiniteSheetValue(topY) Then

        evidence.AddFailure evidencePrefix & _
            ": IView.GetOutline returned non-finite bounds."
        Exit Function
    End If

    evidence.AddInfo evidencePrefix & "_OUTLINE|left=" & _
        Format$(leftX, "0.000000000") & _
        "|bottom=" & Format$(bottomY, "0.000000000") & _
        "|right=" & Format$(rightX, "0.000000000") & _
        "|top=" & Format$(topY, "0.000000000")

    TryGetViewOutline = True
    Exit Function

Failed:
    evidence.AddFailure evidencePrefix & " outline readback error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function VerifySectionSelection( _
    ByRef selectionManager As SldWorks.SelectionMgr, _
    ByRef sectionSegments As Collection, _
    ByRef sourceView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim selectedCount As Long
    selectedCount = selectionManager.GetSelectedObjectCount2(-1)
    If selectedCount <> sectionSegments.Count Then
        evidence.AddFailure "Section selection count mismatch: expected=" & _
            CStr(sectionSegments.Count) & ", actual=" & _
            CStr(selectedCount) & "."
        Exit Function
    End If

    Dim expectedViewName As String
    expectedViewName = Module8_RuntimeSupport.GetViewName(sourceView)

    Dim selectionIndex As Long
    For selectionIndex = 1 To selectedCount
        Dim selectedObject As Object
        Set selectedObject = _
            selectionManager.GetSelectedObject6(selectionIndex, -1)

        If selectedObject Is Nothing Then
            evidence.AddFailure "Section selection contains Nothing at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        Dim expectedSegment As SldWorks.SketchSegment
        Set expectedSegment = sectionSegments(selectionIndex)
        If Not selectedObject Is expectedSegment Then
            evidence.AddFailure "Section selection order mismatch at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        Dim selectedView As SldWorks.View
        Set selectedView = selectionManager.GetSelectedObjectsDrawingView2( _
            selectionIndex, -1)

        If selectedView Is Nothing Then
            evidence.AddFailure "Section selection has no drawing-view owner at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        If StrComp( _
            Module8_RuntimeSupport.GetViewName(selectedView), _
            expectedViewName, vbTextCompare) <> 0 Then

            evidence.AddFailure "Section selection belongs to the wrong view at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        If selectionManager.GetSelectedObjectMark(selectionIndex) <> 0 Then
            evidence.AddFailure "Section selection has an unexpected mark at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If
    Next selectionIndex

    evidence.AddInfo "SECTION_SELECTION|count=" & CStr(selectedCount) & _
        "|orderVerified=True|view=" & expectedViewName & _
        "|marks=0"
    VerifySectionSelection = True
    Exit Function

Failed:
    evidence.AddFailure "Section selection verification error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function IsFiniteSheetValue(ByVal value As Double) As Boolean
    On Error GoTo Failed

    If value <> value Then Exit Function
    If Abs(value) > 1000000# Then Exit Function

    IsFiniteSheetValue = True
    Exit Function

Failed:
    IsFiniteSheetValue = False
End Function

Private Function AddSectionSegment( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef sourceView As SldWorks.View, _
    ByRef sectionSegments As Collection, _
    ByVal startX As Double, _
    ByVal startY As Double, _
    ByVal endX As Double, _
    ByVal endY As Double, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim startViewX As Double
    Dim startViewY As Double
    Dim endViewX As Double
    Dim endViewY As Double

    If Not SheetToViewSketchCoordinates( _
        sourceView, startX, startY, startViewX, startViewY, _
        "SECTION", evidence) Then
        Exit Function
    End If

    If Not SheetToViewSketchCoordinates( _
        sourceView, endX, endY, endViewX, endViewY, _
        "SECTION", evidence) Then
        Exit Function
    End If

    Dim sectionSegment As SldWorks.SketchSegment
    Set sectionSegment = swDrawModel.SketchManager.CreateLine( _
        startViewX, startViewY, 0#, endViewX, endViewY, 0#)

    If sectionSegment Is Nothing Then
        evidence.AddFailure "SketchManager.CreateLine returned Nothing for section."
        Exit Function
    End If

    sectionSegments.Add sectionSegment
    AddSectionSegment = True
    Exit Function

Failed:
    evidence.AddFailure "Section-segment creation error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function SheetToViewSketchCoordinates( _
    ByRef sourceView As SldWorks.View, _
    ByVal sheetX As Double, _
    ByVal sheetY As Double, _
    ByRef viewX As Double, _
    ByRef viewY As Double, _
    ByVal evidenceContext As String, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim evidencePrefix As String
    evidencePrefix = UCase$(Trim$(evidenceContext))
    If Len(evidencePrefix) = 0 Then evidencePrefix = "VIEW"

    Dim xform As Variant
    xform = sourceView.GetXform
    If IsEmpty(xform) Or Not IsArray(xform) Then
        evidence.AddFailure evidencePrefix & _
            ": IView.GetXform returned no coordinate data."
        Exit Function
    End If
    If UBound(xform) - LBound(xform) + 1 < 3 Then
        evidence.AddFailure evidencePrefix & _
            ": IView.GetXform returned fewer than three values."
        Exit Function
    End If

    Dim viewScale As Double
    viewScale = CDbl(xform(LBound(xform) + 2))
    If Not IsFiniteSheetValue(viewScale) Or viewScale <= 0# Then
        evidence.AddFailure evidencePrefix & _
            ": IView.GetXform returned a non-finite or non-positive scale."
        Exit Function
    End If

    If Not IsFiniteSheetValue(CDbl(xform(LBound(xform)))) Or _
       Not IsFiniteSheetValue(CDbl(xform(LBound(xform) + 1))) Then

        evidence.AddFailure evidencePrefix & _
            ": IView.GetXform returned a non-finite sheet origin."
        Exit Function
    End If

    Dim deltaX As Double
    Dim deltaY As Double
    deltaX = (sheetX - CDbl(xform(LBound(xform)))) / viewScale
    deltaY = (sheetY - CDbl(xform(LBound(xform) + 1))) / viewScale

    Dim viewAngle As Double
    viewAngle = sourceView.Angle
    If Not IsFiniteSheetValue(viewAngle) Then
        evidence.AddFailure evidencePrefix & _
            ": source-view angle is non-finite."
        Exit Function
    End If

    viewX = deltaX * Cos(viewAngle) + deltaY * Sin(viewAngle)
    viewY = -deltaX * Sin(viewAngle) + deltaY * Cos(viewAngle)

    If Not IsFiniteSheetValue(viewX) Or Not IsFiniteSheetValue(viewY) Then
        evidence.AddFailure evidencePrefix & _
            ": computed view coordinates are non-finite."
        Exit Function
    End If

    evidence.AddInfo evidencePrefix & "_COORDINATE|sheetX=" & _
        Format$(sheetX, "0.000000") & _
        "|sheetY=" & Format$(sheetY, "0.000000") & _
        "|viewX=" & Format$(viewX, "0.000000") & _
        "|viewY=" & Format$(viewY, "0.000000") & _
        "|xformSheetX=" & _
        Format$(CDbl(xform(LBound(xform))), "0.000000") & _
        "|xformSheetY=" & _
        Format$(CDbl(xform(LBound(xform) + 1)), "0.000000") & _
        "|scale=" & Format$(viewScale, "0.000000") & _
        "|angle=" & Format$(viewAngle, "0.000000") & _
        "|liveVerification=Required"

    SheetToViewSketchCoordinates = True
    Exit Function

Failed:
    evidence.AddFailure evidencePrefix & " coordinate transform error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function FileNameWithoutExtension(ByVal path As String) As String
    Dim slashPosition As Long
    slashPosition = InStrRev(path, "\")

    Dim nameOnly As String
    If slashPosition > 0 Then
        nameOnly = Mid$(path, slashPosition + 1)
    Else
        nameOnly = path
    End If

    Dim dotPosition As Long
    dotPosition = InStrRev(nameOnly, ".")

    If dotPosition > 1 Then
        FileNameWithoutExtension = Left$(nameOnly, dotPosition - 1)
    Else
        FileNameWithoutExtension = nameOnly
    End If
End Function

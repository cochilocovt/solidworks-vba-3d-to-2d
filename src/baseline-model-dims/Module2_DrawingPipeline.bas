Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const swCreateSectionView_NotAligned As Long = 1

' swDisplayMode_e, MCP-confirmed 2026-08-06. All three were wrong before r16
' and none of them raised: a wrong display mode renders a plausible view.
'
'   was 3 for HLR   -> 3 is swSHADED. HLR is swHIDDEN = 2.
'   was 1 for HLV   -> correct; swHIDDEN_GREYED = 1 is "Hidden Lines Visible".
'   was 6 for shaded-with-edges -> 6 is swFACETED_HIDDEN, and
'       IView.SetDisplayMode3 Remarks state that any swFACETED_* passed in
'       Mode is silently treated as its non-faceted equivalent. The isometric
'       view was therefore rendered hidden-lines-removed, not shaded.
'
' Shaded-with-edges is swSHADED plus the method's own Edges argument, which
' ConfigureView already passes as True. The Remarks also offer the
' swDrawingsDefaultDisplayTypeHLREdgesWhenShaded user preference; that route
' is not used here because mutating a user preference to render one view is a
' side effect on the operator's installation.
Private Const swDisplayMode_HiddenLinesRemoved As Long = 2
Private Const swDisplayMode_HiddenLinesVisible As Long = 1
Private Const swDisplayMode_ShadedWithEdges As Long = 3

Public Sub CreateDrawing( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String)

    RunDrawingPipeline swApp, swPart, partPath, templatePath
End Sub

Public Sub RunDrawingPipeline( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String)

    On Error GoTo FailRun

    Dim holes As Collection
    Set holes = Module3_ModelAudit.GetAllHoleLikeFeatures(swPart)

    Dim swDrawModel As SldWorks.ModelDoc2
    Set swDrawModel = swApp.NewDocument(templatePath, 0, 0, 0)
    If swDrawModel Is Nothing Then
        Err.Raise vbObjectError + 701, , "Failed to create drawing document"
    End If

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swDrawModel

    Dim swFrontView As SldWorks.View
    CreateViews swPart, swDrawModel, swDraw, partPath, swFrontView

    swDrawModel.ForceRebuild3 False
    DoEvents
    Sleep 300

    ' ORDINATES FIRST, import second. Reordered 2026-08-06 (r22).
    '
    ' r21 was the first run with both producers active, and the Y ordinate
    ' chain failed with swCreateOrdDimErr_OrdFailure, then returned success on
    ' the holes-only retry while creating nothing at all. Candidate collection
    ' was byte-identical to r20's successful run - same 39 edges, same 5+5
    ' stations, same datums - so nothing about the geometry changed. The only
    ' difference was that 19 imported dimensions already existed in the view,
    ' including a `72.00` spanning exactly the two +/-36 silhouette edges the
    ' Y chain needs.
    '
    ' Whether a dimensioned entity can still anchor an ordinate is NOT
    ' documented: neither AddOrdinateDimension nor InsertModelAnnotations4
    ' says anything about it (MCP, 2026-08-06). Running the ordinate engine
    ' against a clean view removes the interaction entirely, and is the right
    ' order regardless - import is additive and has no such precondition.
    '
    ' This ordering is also the experiment. If the Y chain now succeeds, the
    ' imported dimensions were the cause; if it still fails, they were not.
    Dim ordinateStatus As Module5_FallbackDimensionEngine.OrdinateRunStatus
    ordinateStatus.LastFailureCode = _
        Module5_FallbackDimensionEngine.ORDINATE_CODE_UNSET

    If Module1_Main.GlobalConfig.UseOrdinateDims Then
        AddFallbackOrdinateDimensions swDrawModel, swDraw, ordinateStatus
    End If

    Dim importedModelItems As Long
    importedModelItems = 0

    If Module1_Main.GlobalConfig.UseModelDimensions Then
        importedModelItems = Module4_ModelItemImporter.ImportModelItemsAcrossDrawing(swDrawModel, swDraw, swFrontView.Name)
    End If

    If Module1_Main.GlobalConfig.AutoArrange Then
        Module4_ModelItemImporter.AutoArrangeAllDrawingDimensions swDrawModel, swDraw
    End If

    If Module1_Main.GlobalConfig.PopulateTitle Then
        Module7_TitleBlockEngine.PopulateTitleBlock swPart, swDrawModel, swDraw
    End If

    swDrawModel.ForceRebuild3 False

    ' After the rebuild, not before. IAnnotation.IsDangling reads False on a
    ' freshly created ordinate and only becomes True once the drawing has
    ' rebuilt, so a prune inside the ordinate stage sees nothing (r14).
    If Module1_Main.GlobalConfig.UseOrdinateDims Then
        Module5_FallbackDimensionEngine.PruneDanglingAcrossDrawing _
            swDrawModel, swDraw, ordinateStatus
    End If

    swDrawModel.ViewZoomtofit2

    Dim qaReport As String
    qaReport = Module6_QAEngine.BuildRunSummary( _
        swDrawModel, swDraw, holes, importedModelItems, ordinateStatus)

    ' Write before showing. A MsgBox is not a record: the first live trunk
    ' run lost its ordinate evidence the moment the operator clicked OK.
    Dim qaReportPath As String
    qaReportPath = Module21_EvidenceSink.WriteQaReport( _
        swPart.GetPathName, qaReport)

    If Len(qaReportPath) > 0 Then
        qaReport = qaReport & vbCrLf & "Report written to:" & vbCrLf & _
            qaReportPath & vbCrLf
    Else
        qaReport = qaReport & vbCrLf & _
            "WARNING: QA report could not be written to disk." & vbCrLf
    End If

    MsgBox qaReport, vbInformation, "Drawing QA Summary"
    Exit Sub

FailRun:
    MsgBox "Drawing pipeline failed: " & Err.Description, vbCritical, "Pipeline Error"
End Sub

Private Sub AddFallbackOrdinateDimensions( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef status As Module5_FallbackDimensionEngine.OrdinateRunStatus)

    On Error GoTo SafeExit

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        ' Was a substring test for "ISO" in IView.Name. SOLIDWORKS auto-names
        ' these views "Drawing View1".."Drawing View5", so it never matched and
        ' the r7 run ordinated the isometric view. Module8 classifies from
        ' IView.Type and IView.GetOrientationName instead.
        If Module8_ViewClassifier.AllowsOrdinateDimensions( _
            Module8_ViewClassifier.ClassifyView(swView)) Then
            ' HLR is a precondition of the ordinate engine, not a preference.
            ' Harvest under HLR whatever the operator asked for, then restore.
            Dim restoreMode As Long
            restoreMode = ForceHlrForHarvest(swView, status)

            Module5_FallbackDimensionEngine.CreateHoleOrdinateDims swDraw, swView, Module1_Main.GlobalConfig.DatumOrigin, status

            RestoreDisplayMode swView, restoreMode
        End If
        Set swView = swView.GetNextView
    Loop

SafeExit:
End Sub

' Why the ordinate engine cannot run under HLV
' ---------------------------------------------
' GetVisibleEntities2 returns hidden-line edges under HLV, and those edges
' select, return swCreateOrdDimErr_Success, and then dangle (r17). r22 showed
' a second way it degrades: HLV raised the candidate pool from 39 edges to 64,
' which handed ResolveOneDatum a nearer straight edge, so the X end datum
' landed 43 mm inside the part instead of on the end face. Same run under HLR
' (r20) produced the reference chain exactly.
'
' r17 made HLR the default in ResetGlobalConfig - but that is the no-form
' fallback. UserForm1 defaults chkHLR to False and persists the operator's
' last answer, so every form run silently reverted to HLV. Fixing the form
' default would only move the dependency; a geometry precondition of the
' engine should not be reachable from a checkbox at all.
'
' GetDisplayMode2 / SetDisplayMode3 / UpdateViewDisplayGeometry all MCP-checked
' 2026-08-06. UpdateViewDisplayGeometry Remarks name this exact case: it gives
' immediate access to the new geometry after an HLR/HLV switch, without waiting
' for Windows to repaint.
'
' Returns the mode to restore, or -1 when no change was made.
Private Function ForceHlrForHarvest( _
    ByRef swView As SldWorks.View, _
    ByRef status As Module5_FallbackDimensionEngine.OrdinateRunStatus) As Long

    ForceHlrForHarvest = -1
    On Error GoTo SafeExit

    Dim currentMode As Long
    currentMode = swView.GetDisplayMode2

    If currentMode = swDisplayMode_HiddenLinesRemoved Then
        status.HarvestDisplayMode = "HLR (already)"
        Exit Function
    End If

    ' UseParent False so the view keeps its own local setting. Facetted False
    ' keeps precision quality - a faceted view is the draft-quality tessellation
    ' and is not what the engine measures against.
    If swView.SetDisplayMode3(False, swDisplayMode_HiddenLinesRemoved, False, True) = False Then
        status.HarvestDisplayMode = "HLR forced: REFUSED (was " & currentMode & ")"
        Exit Function
    End If

    swView.UpdateViewDisplayGeometry

    status.HarvestDisplayMode = "HLR forced (was " & currentMode & ")"
    ForceHlrForHarvest = currentMode
    Exit Function

SafeExit:
    status.HarvestDisplayMode = "HLR force errored: " & Err.Number & " " & Err.Description
    Err.Clear
End Function

Private Sub RestoreDisplayMode( _
    ByRef swView As SldWorks.View, _
    ByVal restoreMode As Long)

    On Error Resume Next
    If restoreMode < 0 Then Exit Sub
    swView.SetDisplayMode3 False, restoreMode, False, True
    swView.UpdateViewDisplayGeometry
    Err.Clear
End Sub

Private Sub CreateViews( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal partPath As String, _
    ByRef swFrontView As SldWorks.View)

    On Error GoTo FailCreateViews

    Dim swPartDoc As SldWorks.PartDoc
    Set swPartDoc = swPart

    Dim bbox As Variant
    bbox = swPartDoc.GetPartBox(True)

    Dim partW As Double
    Dim partH As Double
    partW = Abs(bbox(3) - bbox(0))
    partH = Abs(bbox(4) - bbox(1))

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    Dim sheetW As Double
    Dim sheetH As Double
    swSheet.GetSize sheetW, sheetH

    Dim scaleVal As Double
    scaleVal = Module1_Main.GlobalConfig.SheetScale
    If scaleVal <= 0 Then scaleVal = 1#

    Dim displayMode As Long
    If Module1_Main.GlobalConfig.UseHLR Then
        displayMode = swDisplayMode_HiddenLinesRemoved
    Else
        displayMode = swDisplayMode_HiddenLinesVisible
    End If

    Dim xFront As Double
    Dim yFront As Double
    xFront = sheetW * 0.42
    yFront = sheetH * 0.48

    Set swFrontView = swDraw.CreateDrawViewFromModelView3(partPath, "*Front", xFront, yFront, 0)
    If swFrontView Is Nothing Then
        Err.Raise vbObjectError + 711, , "Failed to create front view"
    End If
    ConfigureView swFrontView, displayMode, scaleVal

    If Module1_Main.GlobalConfig.CreateTop Then
        CreateNamedView swDraw, partPath, "*Top", xFront, yFront + (partH * scaleVal) * 0.75, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateBottom Then
        CreateNamedView swDraw, partPath, "*Bottom", xFront, yFront - (partH * scaleVal) * 0.75, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateRight Then
        CreateNamedView swDraw, partPath, "*Right", xFront + (partW * scaleVal) * 0.55, yFront, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateLeft Then
        CreateNamedView swDraw, partPath, "*Left", xFront - (partW * scaleVal) * 0.55, yFront, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateBack Then
        CreateNamedView swDraw, partPath, "*Back", xFront + (partW * scaleVal) * 0.85, yFront, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateIso Then
        CreateNamedView swDraw, partPath, "*Isometric", sheetW * 0.8, sheetH * 0.72, swDisplayMode_ShadedWithEdges, scaleVal * 0.55
    End If

    swDrawModel.ForceRebuild3 False
    DoEvents
    Sleep 300

    If Module1_Main.GlobalConfig.CreateSection Then
        CreateSectionFromConfig swDrawModel, swDraw, bbox, swFrontView, displayMode, scaleVal
    End If

    Exit Sub

FailCreateViews:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub CreateNamedView( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal partPath As String, _
    ByVal modelViewName As String, _
    ByVal xPos As Double, _
    ByVal yPos As Double, _
    ByVal displayMode As Long, _
    ByVal scaleVal As Double)

    Dim swView As SldWorks.View
    Set swView = swDraw.CreateDrawViewFromModelView3(partPath, modelViewName, xPos, yPos, 0)

    If Not swView Is Nothing Then
        ConfigureView swView, displayMode, scaleVal
    End If
End Sub

Private Sub ConfigureView(ByRef swView As SldWorks.View, ByVal displayMode As Long, ByVal scaleVal As Double)
    On Error Resume Next

    swView.ScaleDecimal = scaleVal
    swView.SetDisplayMode3 False, displayMode, False, True
    swView.SetLightweightToResolved
End Sub

Private Sub CreateSectionFromConfig( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal bbox As Variant, _
    ByRef swFrontView As SldWorks.View, _
    ByVal displayMode As Long, _
    ByVal scaleVal As Double)

    On Error GoTo SafeExit

    If swFrontView Is Nothing Then Exit Sub

    swDraw.ActivateView swFrontView.Name

    Dim midX As Double
    Dim midY As Double
    midX = (bbox(0) + bbox(3)) / 2#
    midY = (bbox(1) + bbox(4)) / 2#

    Dim spanX As Double
    Dim spanY As Double
    spanX = Abs(bbox(3) - bbox(0))
    spanY = Abs(bbox(4) - bbox(1))

    Dim secLabel As String
    Dim secVertical As Boolean
    GetPrimarySectionSettings secLabel, secVertical

    Dim swLine As SldWorks.SketchLine
    If secVertical Then
        Set swLine = swDrawModel.SketchManager.CreateLine(midX, bbox(1) - spanY * 0.15, 0, midX, bbox(4) + spanY * 0.15, 0)
    Else
        Set swLine = swDrawModel.SketchManager.CreateLine(bbox(0) - spanX * 0.15, midY, 0, bbox(3) + spanX * 0.15, midY, 0)
    End If

    If swLine Is Nothing Then GoTo SafeExit

    swDrawModel.ClearSelection2 True
    swLine.Select4 True, Nothing

    Dim frontPos As Variant
    frontPos = swFrontView.Position

    Dim targetX As Double
    Dim targetY As Double
    targetX = frontPos(0) + 0.18
    targetY = frontPos(1)

    Dim swSectionView As SldWorks.View
    Set swSectionView = swDraw.CreateSectionViewAt5(targetX, targetY, 0, secLabel, swCreateSectionView_NotAligned, Nothing, 0)

    If Not swSectionView Is Nothing Then
        ConfigureView swSectionView, displayMode, scaleVal
    End If

SafeExit:
    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""
End Sub

Private Sub GetPrimarySectionSettings(ByRef secLabel As String, ByRef secVertical As Boolean)
    secLabel = vbNullString
    secVertical = False

    If Module1_Main.GlobalSectionCount > 0 Then
        secLabel = Module1_Main.GetSectionLabelOrDefault(1)
        secVertical = Module1_Main.GlobalSections(1).Vertical
    Else
        secLabel = Trim$(Module1_Main.GlobalConfig.SectionLabel)
        If Len(secLabel) = 0 Then secLabel = "J"
        secVertical = Module1_Main.GlobalConfig.SectionVertical
    End If
End Sub

' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub

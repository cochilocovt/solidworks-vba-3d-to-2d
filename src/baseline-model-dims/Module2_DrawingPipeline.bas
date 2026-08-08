Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const swCreateSectionView_NotAligned As Long = 1

' Section-cut evidence for the QA report. Module-level rather than threaded
' through BuildRunSummary, which would mean changing Module6's signature for
' two strings. Reset at the start of every pipeline run.
Public LastSectionPlan As String
Public LastSectionReadback As String

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

    LastSectionPlan = vbNullString
    LastSectionReadback = vbNullString

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

    ' InsertModelAnnotations4 still returned zero native hole callouts in the
    ' r27 selected-view run. The controlled fallback reads type-specific Hole
    ' Wizard data, maps an edge owned by that feature into an orthographic
    ' drawing view, and inserts an attached leader note. It is deliberately
    ' separate from the ordinary dimension count and is proved again by QA.
    If Module1_Main.GlobalConfig.UseModelDimensions And _
       Module1_Main.GlobalConfig.ImportHoleCallouts Then

        Module9_HoleCalloutEngine.EnsureControlledHoleCallouts _
            swApp, swPart, swDrawModel, swDraw, holes
    End If

    If Module1_Main.GlobalConfig.PopulateTitle Then
        Module7_TitleBlockEngine.PopulateTitleBlock swPart, swDrawModel, swDraw
    End If

    swDrawModel.ForceRebuild3 False

    ' A7/C4 final placement runs only after dimensions, controlled callouts,
    ' and title-block content exist. Module10 measures the complete travelling
    ' envelope of each view, plans every move before mutating, preserves scale,
    ' then rebuilds and fails closed on any remaining collision or boundary
    ' violation. Its outcome is included in the final QA verdict.
    Dim layoutPassed As Boolean
    layoutPassed = Module10_SheetLayoutEngine.ArrangeDrawingContent( _
        swDrawModel, swDraw)

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

    If layoutPassed = False Then
        qaReport = "RUN STATUS: FAIL - SHEET LAYOUT" & vbCrLf & _
            "The layout engine could not place every measured view " & _
            "envelope. Original view state was restored; inspect the " & _
            "LAYOUT_REJECT evidence below." & vbCrLf & vbCrLf & qaReport
    End If

    If Len(LastSectionPlan) > 0 Then
        qaReport = qaReport & LastSectionPlan & vbCrLf
    End If

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

    If layoutPassed = False Then
        MsgBox qaReport, vbCritical, "Drawing QA FAILED - Sheet Layout"
    Else
        MsgBox qaReport, vbInformation, "Drawing QA Summary"
    End If
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
        CreateSectionFromConfig swPart, swDrawModel, swDraw, bbox, swFrontView, displayMode, scaleVal
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

' The reference drawing's section J-J is a stepped cut, not a centreline cut
' -----------------------------------------------------------------------
' Read off test_assets/reference_drawings/P-0251-14A-001.PNG: the section line
' runs along the part's long axis through the centre of the R36 end bore, jogs
' sideways once, and continues through one row of counterbored holes. That is
' what puts both the 040/047 bore profile AND the counterbore steps in the same
' section. Cutting straight down the centreline - what the trunk did through
' r23 - catches the bore and misses all six holes.
'
' A stepped section is not swCreateSectionView_OffsetSection. That flag is
' documented as "an aligned section view is created (two lines at an angle)",
' which is the revolved case. CreateSectionViewAt5 Remarks say to select "the
' section line or lines" - plural - so a stepped cut is an ordinary NotAligned
' section over a multi-segment sketch (MCP, 2026-08-06).
'
' Coordinate frame, and why the old code proved nothing about it
' --------------------------------------------------------------
' CodeStack row 9 warns that view-owned sketch entities live in the view's own
' frame and says to verify the transform direction on SW2025. The pre-r24 line
' was drawn at midY, and for this fixture the model is symmetric about Y=0, so
' midY was 0 - and 0 is 0 under any scale factor. The cut landing correctly was
' therefore no evidence at all about whether one sketch unit is one model metre
' or one sheet metre.
'
' Two facts settle the origin: the line spanned bbox(0)..bbox(3), roughly
' 0..0.196, while the front view sits at sheet position ~(0.176, 0.143). Had
' the frame been sheet coordinates that line would have run along the bottom
' edge of the sheet, intersected no view, and CreateSectionViewAt5 would have
' had nothing to cut. It produced a correct section, so the frame is the view's
' local frame with the model origin at its origin.
'
' Scale is still not proven. CodeStack row 9's "move, scale, and rotate with
' the view" implies the sketch is stored at model size and the view scale is
' applied on display, which is the assumption used here. This is the first
' revision to place a section line at a NONZERO offset, so it is also the first
' run that can tell the difference: the readback below reports the endpoints
' the sketch actually kept. If one unit were a sheet metre the reported jog
' would come back at 0.6667 of what was asked for.
Private Sub CreateSectionFromConfig( _
    ByRef swPart As SldWorks.ModelDoc2, _
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

    ' Along = the axis the section line runs down. Across = the axis it jogs on.
    Dim boreAcross As Double
    Dim rowAcross As Double
    Dim jogAlong As Double
    Dim hasJog As Boolean
    hasJog = PlanSteppedCut(swPart, secVertical, midX, midY, _
                            boreAcross, rowAcross, jogAlong)

    Dim startAlong As Double
    Dim endAlong As Double
    If secVertical Then
        startAlong = bbox(1) - spanY * 0.15
        endAlong = bbox(4) + spanY * 0.15
    Else
        startAlong = bbox(0) - spanX * 0.15
        endAlong = bbox(3) + spanX * 0.15
    End If

    swDrawModel.ClearSelection2 True

    Dim segments As Long
    segments = 0

    If hasJog Then
        ' Bore leg, the jog itself, then the hole-row leg.
        If AddCutSegment(swDrawModel, secVertical, startAlong, boreAcross, jogAlong, boreAcross) Then segments = segments + 1
        If AddCutSegment(swDrawModel, secVertical, jogAlong, boreAcross, jogAlong, rowAcross) Then segments = segments + 1
        If AddCutSegment(swDrawModel, secVertical, jogAlong, rowAcross, endAlong, rowAcross) Then segments = segments + 1
    End If

    If segments < 3 Then
        ' Either no distinct hole row was found or a segment failed. Fall back
        ' to the straight centreline cut rather than leaving a partial line -
        ' a half-drawn stepped cut is worse than the r23 behaviour.
        swDrawModel.ClearSelection2 True
        hasJog = False
        segments = 0
        If AddCutSegment(swDrawModel, secVertical, startAlong, midAcross(secVertical, midX, midY), _
                         endAlong, midAcross(secVertical, midX, midY)) Then segments = 1
    End If

    LastSectionPlan = DescribeSectionPlan(hasJog, secVertical, boreAcross, rowAcross, _
                                          jogAlong, segments, scaleVal)

    If segments = 0 Then GoTo SafeExit

    Dim frontPos As Variant
    frontPos = swFrontView.Position

    ' Was frontPos(0) + 0.18 - a fixed 180 mm nudge that took no account of how
    ' wide either view actually is, which is why the section overflowed the
    ' sheet border on the r23 sheet. Half the front view plus half the section
    ' (the section is as long as the part's long axis) plus a 15 mm gutter.
    Dim halfFront As Double
    Dim halfSection As Double
    halfFront = (spanX * scaleVal) / 2#
    If secVertical Then
        halfSection = (spanX * scaleVal) / 2#
    Else
        halfSection = (spanY * scaleVal) / 2#
    End If

    Dim targetX As Double
    Dim targetY As Double
    targetX = frontPos(0) + halfFront + halfSection + 0.015
    targetY = frontPos(1)

    Dim swSectionView As SldWorks.View
    Set swSectionView = swDraw.CreateSectionViewAt5(targetX, targetY, 0, secLabel, swCreateSectionView_NotAligned, Nothing, 0)

    If Not swSectionView Is Nothing Then
        ConfigureView swSectionView, displayMode, scaleVal
    Else
        LastSectionPlan = LastSectionPlan & vbCrLf & _
            "  CreateSectionViewAt5 returned Nothing"
    End If

SafeExit:
    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""
End Sub

Private Function midAcross(ByVal secVertical As Boolean, _
                           ByVal midX As Double, ByVal midY As Double) As Double
    If secVertical Then
        midAcross = midX
    Else
        midAcross = midY
    End If
End Function

' Creates one leg of the section line and appends it to the selection. Reads
' the endpoints back off the sketch segment rather than trusting the arguments,
' because the sketch frame's scale is the open question this revision answers.
Private Function AddCutSegment( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByVal secVertical As Boolean, _
    ByVal along1 As Double, ByVal across1 As Double, _
    ByVal along2 As Double, ByVal across2 As Double) As Boolean

    On Error GoTo Failed

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    If secVertical Then
        x1 = across1: y1 = along1: x2 = across2: y2 = along2
    Else
        x1 = along1: y1 = across1: x2 = along2: y2 = across2
    End If

    Dim swLine As SldWorks.SketchLine
    Set swLine = swDrawModel.SketchManager.CreateLine(x1, y1, 0, x2, y2, 0)
    If swLine Is Nothing Then GoTo Failed

    ' Append rather than replace: CreateSectionViewAt5 wants every leg selected
    ' at once ("the section line or lines").
    swLine.Select4 True, Nothing

    Dim swStart As SldWorks.SketchPoint
    Set swStart = swLine.GetStartPoint2
    If Not swStart Is Nothing Then
        LastSectionReadback = LastSectionReadback & _
            Format$(swStart.X * 1000#, "0.###") & "," & _
            Format$(swStart.Y * 1000#, "0.###") & " "
    End If

    AddCutSegment = True
    Exit Function

Failed:
    Err.Clear
    AddCutSegment = False
End Function

' Picks the two across-axis coordinates the stepped cut needs: the bore's, and
' the busiest row of ordinary holes. Deliberately derived from the model rather
' than hardcoded to this fixture - the widest hole-like feature is the bore on
' any part where a bore is the reason a section exists at all.
Private Function PlanSteppedCut( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal secVertical As Boolean, _
    ByVal midX As Double, ByVal midY As Double, _
    ByRef boreAcross As Double, ByRef rowAcross As Double, _
    ByRef jogAlong As Double) As Boolean

    On Error GoTo Failed

    boreAcross = midAcross(secVertical, midX, midY)
    PlanSteppedCut = False

    Dim holes As Collection
    Set holes = Module3_ModelAudit.GetAllHoleLikeFeatures(swPart)
    If holes Is Nothing Then Exit Function
    If holes.Count < 2 Then Exit Function

    ' Widest feature is the bore. Everything within 60% of its diameter counts
    ' as "the bore" too, so a counterbore's own ring does not become the row.
    Dim biggest As Double
    Dim boreAlong As Double
    Dim i As Long
    biggest = 0#
    For i = 1 To holes.Count
        Dim d As Double
        d = Module3_ModelAudit.GetHoleDiameter(holes(i))
        If d > biggest Then
            biggest = d
            boreAcross = AcrossOf(secVertical, holes(i))
            boreAlong = AlongOf(secVertical, holes(i))
        End If
    Next i

    If biggest <= 0# Then Exit Function

    ' Busiest across-coordinate among the non-bore holes, and the one of those
    ' furthest along the axis from the bore - that is where the cut has to end
    ' up to pass through the row.
    Dim bestCount As Long
    Dim bestAcross As Double
    Dim nearestAlong As Double
    Dim haveRow As Boolean
    bestCount = 0
    haveRow = False

    For i = 1 To holes.Count
        If Module3_ModelAudit.GetHoleDiameter(holes(i)) < biggest * 0.6 Then
            Dim candAcross As Double
            candAcross = AcrossOf(secVertical, holes(i))

            Dim n As Long
            Dim closestAlong As Double
            n = 0
            closestAlong = 0#

            Dim j As Long
            For j = 1 To holes.Count
                If Module3_ModelAudit.GetHoleDiameter(holes(j)) < biggest * 0.6 Then
                    If Abs(AcrossOf(secVertical, holes(j)) - candAcross) <= 0.0005 Then
                        n = n + 1
                        Dim a As Double
                        a = AlongOf(secVertical, holes(j))
                        If n = 1 Then
                            closestAlong = a
                        ElseIf Abs(a - boreAlong) < Abs(closestAlong - boreAlong) Then
                            closestAlong = a
                        End If
                    End If
                End If
            Next j

            If n > bestCount Then
                bestCount = n
                bestAcross = candAcross
                nearestAlong = closestAlong
                haveRow = True
            End If
        End If
    Next i

    If Not haveRow Then Exit Function
    If bestCount < 2 Then Exit Function
    If Abs(bestAcross - boreAcross) <= 0.0005 Then Exit Function

    rowAcross = bestAcross
    ' Jog midway between the bore centre and the first hole of the row, so
    ' neither feature is clipped by the step.
    jogAlong = (boreAlong + nearestAlong) / 2#

    PlanSteppedCut = True
    Exit Function

Failed:
    Err.Clear
    PlanSteppedCut = False
End Function

Private Function AlongOf(ByVal secVertical As Boolean, ByVal holeInfo As Variant) As Double
    If secVertical Then
        AlongOf = Module3_ModelAudit.GetHoleY(holeInfo)
    Else
        AlongOf = Module3_ModelAudit.GetHoleX(holeInfo)
    End If
End Function

Private Function AcrossOf(ByVal secVertical As Boolean, ByVal holeInfo As Variant) As Double
    If secVertical Then
        AcrossOf = Module3_ModelAudit.GetHoleX(holeInfo)
    Else
        AcrossOf = Module3_ModelAudit.GetHoleY(holeInfo)
    End If
End Function

Private Function DescribeSectionPlan( _
    ByVal hasJog As Boolean, ByVal secVertical As Boolean, _
    ByVal boreAcross As Double, ByVal rowAcross As Double, _
    ByVal jogAlong As Double, ByVal segments As Long, _
    ByVal scaleVal As Double) As String

    Dim text As String
    text = "Section cut: "
    If hasJog Then
        text = text & "stepped, " & segments & " segments" & vbCrLf & _
            "  bore leg at " & Format$(boreAcross * 1000#, "0.###") & "mm, " & _
            "row leg at " & Format$(rowAcross * 1000#, "0.###") & "mm, " & _
            "jog at " & Format$(jogAlong * 1000#, "0.###") & "mm"
    Else
        text = text & "straight centreline (no distinct hole row found), " & _
            segments & " segment(s)"
    End If

    If Len(LastSectionReadback) > 0 Then
        text = text & vbCrLf & "  sketch readback (mm): " & Trim$(LastSectionReadback) & vbCrLf & _
            "  view scale " & Format$(scaleVal, "0.####") & _
            " - readback matching the requested values proves the sketch frame is model-scale"
    End If

    DescribeSectionPlan = text
End Function

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

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

' Retired Phase 9 mutation entry point. User accepted the existing layout on
' 2026-08-04, so this entry point must never reposition or rescale a drawing.
Public Sub R23_ApplyContentLayoutToScratch()
    Module21_EvidenceSink.LogLine _
        "R23_LAYOUT_MUTATION_SKIPPED|reason=UserAcceptedLayoutAsIs" & _
        "|mutations=0"
End Sub

Private Function TryGetAuthorizedScratchPart( _
    ByRef swDrawing As SldWorks.DrawingDoc, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef partPath As String, _
    ByRef failure As String) As Boolean

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet
    If swSheet Is Nothing Then
        failure = "NoCurrentSheet"
        Exit Function
    End If

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then
        failure = "NoViewsOnSheet"
        Exit Function
    End If

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If Not swView Is Nothing Then
            If Not swView.ReferencedDocument Is Nothing Then
                Set swPart = swView.ReferencedDocument
                Exit For
            End If
        End If
    Next i

    If swPart Is Nothing Then
        failure = "NoReferencedDocument"
        Exit Function
    End If

    partPath = swPart.GetPathName
    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        failure = "UnauthorizedFixture|path=" & partPath
        Exit Function
    End If

    TryGetAuthorizedScratchPart = True
End Function

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

    ' R23-1100. All later drawing decisions consume one typed graph. The
    ' legacy feature collection had no physical identity, view projection or
    ' source-backed manufacturing definition, so it cannot remain the route
    ' that drives dimension creation.
    Dim graph As CLocationGraph
    Set graph = New CLocationGraph

    Dim configurationName As String
    configurationName = swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        evidence.AddFailure _
            "R23 pipeline stopped: feature catalog unavailable."
        GoTo FinishRun
    End If

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
        evidence.AddFailure _
            "The drawing-document interface could not be acquired."
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
            "continuing to create views for inspection. This run is not " & _
                "an " & _
            "acceptance result."
    End If

    If Not Module8_RuntimeSupport.SetConfiguredSheetScale( _
        swSheet, Module1_Main.GlobalConfig.SheetScale, evidence) Then

        If Not Module1_Main.DIAGNOSTIC_DRAWING_MODE Then GoTo FinishRun

        evidence.AddWarning _
            "DIAGNOSTIC_DRAWING_MODE: sheet-scale transaction failed; " & _
            "continuing with the template's existing scale for view " & _
                "inspection."
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

    ' R23-1101 step 3. This is only a geometry-first rough placement. The
    ' finished annotation envelopes are not known until all dimensions and
    ' definitions exist, so the final content-aware layout is deliberately
    ' later in this sequence.
    If Not Module9_LayoutEngine.ArrangeViewsInMeasuredGrid( _
        swDrawModel, swDraw, "R23Rough", False, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "rough geometry layout", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 step 4. View projections are complete before a section,
    ' annotation or ordinate can be created from them.
    If Not BuildAllViewProjections( _
        swApp, swDrawModel, swDraw, graph, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "view projection resolution", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 step 5. The J-J line comes from the graph, never from a view
    ' outline percentage. It is created before imported dimensions so the
    ' section is an explicit import destination.
    If Module1_Main.GlobalConfig.CreateSection Then
        If Not CreateSemanticPrimarySection( _
            swApp, swDrawModel, swDraw, graph, evidence) Then

            If Not ContinueDiagnosticPipeline( _
                "semantic section creation", evidence) Then GoTo FinishRun
        ElseIf Not Module8_RuntimeSupport.RebuildDocumentVerified( _
            swDrawModel, "R23 post-section", evidence) Then

            If Not ContinueDiagnosticPipeline( _
                "semantic post-section rebuild", evidence) Then GoTo FinishRun
        End If
    End If

    ' The new section adds a view. Resolve it as well before importing into
    ' selected views, so it has its own graph records and coverage ledger.
    If Not BuildAllViewProjections( _
        swApp, swDrawModel, swDraw, graph, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "post-section projection resolution", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 steps 6 and 7. Annotation inventory is independent from the
    ' insert result and is reconciled by entity identity before new ordinates
    ' are considered.
    If Not ImportAndReconcileR23Annotations( _
        swApp, swDrawModel, swDraw, primaryView, graph, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "semantic annotation import", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 step 8. Each group is a typed graph scheme with a proved
    ' datum and selection contract; Module5's feature-list fallback is no
    ' longer called from the production route.
    If Not CreateR23OrdinateGroups( _
        swApp, swDrawModel, swDraw, graph, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "R23 ordinate creation", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 step 9. Existing imported section dimensions are reconciled
    ' before any creation. A missing dimension remains a named failure until
    ' its required drawing entities are proved and selected by location.
    If Not ReconcileR23SectionDimensions(swDraw, evidence) Then
        If Not ContinueDiagnosticPipeline( _
            "section dimension reconciliation", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 step 10. Native associative callouts are created only for a
    ' family whose current retained definition is not already native.
    If Not CreateMissingR23Callouts( _
        swApp, swDrawModel, swDraw, graph, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "native callout completion", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 step 11. Isometric comes after all dimension-bearing views,
    ' and the importer never receives it as an eligible target.
    If Module1_Main.GlobalConfig.CreateIso Then
        If Not CreateIsometricView(swDraw, partPath, evidence) Then
            If Not ContinueDiagnosticPipeline( _
                "isometric view creation", evidence) Then GoTo FinishRun
        End If
    Else
        evidence.AddFailure _
            "The fixed acceptance workflow requires an isometric view."
        GoTo FinishRun
    End If

    ' R23-1101 step 11b. The approved structural grid, run once, now that
    ' every required view exists. The step-3 pass could only pre-place the
    ' two orthographic views it could see. This is the placement the LAYOUT
    ' stage judges, and it is the same geometry-first grid the 2026-08-04
    ' user decision preserves - it is not the content-envelope repositioning
    ' that decision retired, which remains uncalled.
    If Not Module9_LayoutEngine.ArrangeViewsInMeasuredGrid( _
        swDrawModel, swDraw, "R23Structural", True, evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "final structural layout", evidence) Then GoTo FinishRun
    End If

    ' R23-707. A moved source view re-lays its section line, so the recorded
    ' geometry proof stops describing the sheet until it is read again.
    RecordSectionLineAfterLayout swDraw, evidence

    ' R23-1101 step 12. Title content precedes arrangement and the final
    ' envelope measurement, so it cannot invalidate a completed layout.
    Module7_TitleBlockEngine.PopulateTitleBlock _
        swPart, swDrawModel, swDraw, evidence

    If Module1_Main.GlobalConfig.AutoArrange Then
        Module4_ModelItemImporter.AutoArrangeAllDrawingDimensions _
            swDrawModel, swDraw, evidence
    End If

    ' R23-823 step 14b. LAST, because every earlier step can move annotation
    ' text and a correction applied before them goes stale. The structural
    ' grid moved the section view after its dimensions were created, which
    ' is how r61's RD2 ended in the zoned border despite being placed at
    ' exactly UsableTop. Annotation origins only - no view moves.
    ClampR23SectionAnnotations swDraw, evidence

    If Not Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, "R23 final content", evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "final content rebuild", evidence) Then GoTo FinishRun
    End If

    ' User decision, 2026-08-04. Final geometry-first grid placement remains
    ' intact, but no content-envelope repositioning or rescaling may occur.
    If Not RecordR23UserAcceptedLayout(evidence) Then

        If Not ContinueDiagnosticPipeline( _
            "user-accepted final layout record", evidence) Then GoTo FinishRun
    End If

    ' R23-1101 step 15. The semantic judge is the production gate; Module6
    ' now supplies only remaining structural/view-boundary checks.
    Module19_SemanticQA.EvaluateSemanticDrawing _
        swApp, swDrawModel, swDraw, graph, evidence

FinishRun:
    On Error Resume Next
    FinalizeRunOnce swDrawModel, swDraw, evidence, finalizationStarted
    If Err.Number <> 0 Then
        MsgBox _
            "Finalization failed before complete evidence could be " & _
            "emitted: " & _
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
        " failed; continuing independent generation stages for visual " & _
            "inspection."
    ContinueDiagnosticPipeline = True
End Function

' R23-1101 step 4. A view is resolved once per graph. Re-running the mapper
' would append a second projection record for the same physical/view key and
' turn a retry into a duplicate-identity defect.
Private Function BuildAllViewProjections( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Function

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then Exit Function

    ' Route D refuses to select while anything is already selected, because
    ' IView.SelectEntity with AppendFlag=False would destroy an interactive
    ' caller's selection. View creation leaves one object selected, so the
    ' r47 run - the first in which Route D was eligible at all - refused it
    ' for every location in the sheet with
    ' D:RefusedPreexistingSelection:count1, and the stepped bore still had no
    ' mapped edge. One stale selection blocked all eleven.
    '
    ' Same correction as the ordinate stage at r44, at a different point in
    ' the pipeline, for the same reason: this is the production route, which
    ' owns the document it just created. The guard in Module13 is untouched
    ' and still fails closed for the read-only probes, which must never
    ' silently discard a user's selection.
    Dim preexistingSelection As Long
    preexistingSelection = _
        swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)

    If preexistingSelection > 0 Then
        evidence.RecordSolidWorksMutation "ClearSelection2:BeforeProjections"
        swDrawModel.ClearSelection2 True
    End If

    evidence.AddInfo "R23_PROJECTION_SELECTION_PRECONDITION" & _
        "|preexisting=" & CStr(preexistingSelection) & _
        "|cleared=" & CStr(preexistingSelection > 0) & _
        "|remaining=" & _
        CStr(swDrawModel.SelectionManager.GetSelectedObjectCount2(-1))

    Dim attempted As Long
    Dim complete As Boolean
    complete = True

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        If Module18_ContentEnvelope.IsTemplateOrientationView(swView) Then
            GoTo ContinueView
        End If

        Dim viewName As String
        viewName = Module8_RuntimeSupport.GetViewName(swView)
        If graph.ProjectionsForView(viewName).Count > 0 Then GoTo ContinueView

        attempted = attempted + 1
        If Not Module13_ProjectionResolution.BuildViewProjections( _
            swApp, swDrawModel, swView, graph, evidence) Then

            complete = False
        End If

ContinueView:
    Next i

    evidence.AddInfo "R23_PROJECTION_PASS|attemptedViews=" & _
        CStr(attempted) & "|projections=" & _
        CStr(graph.ProjectionCount()) & "|complete=" & CStr(complete)

    ' A post-section pass legitimately sees no new view when section creation
    ' was disabled or failed in diagnostic mode.  Existing graph coverage is
    ' judged later; this pass only reports whether every attempted map worked.
    BuildAllViewProjections = complete
    Exit Function

Failed:
    evidence.AddFailure "R23 projection pass error " & CStr(Err.Number) & _
        ": " & Err.Description
End Function

' R23-1101 step 5. The present semantic path intentionally implements the
' approved stepped-bore/grid construction. Other configured section shapes
' fail closed until they have their own graph-backed path rather than falling
' through to the retired outline-percentage section routine.
Private Function CreateSemanticPrimarySection( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If Module1_Main.GetFixtureKey(evidence.PartPath) <> _
        "P-0251-14A-001" Then

        evidence.AddFailure _
            "R23 section path is not yet defined for fixture " & _
            Module1_Main.GetFixtureKey(evidence.PartPath) & "."
        Exit Function
    End If

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Function

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then Exit Function

    Dim sectionLabel As String
    sectionLabel = Module1_Main.GetSectionLabelOrDefault(1)

    ' ResolveSectionPath emits evidence only when it succeeds, so the 18:45
    ' and 23:24 production runs recorded the bare sentence below and nothing
    ' about WHY. Every candidate view is now reported with its own rejection
    ' reason, and the reasons are carried into the failure text. A view with
    ' no projections at all is the uninteresting case and is summarised
    ' rather than listed.
    Dim path As CSectionPath
    Dim resolvedPath As CSectionPath
    Dim rejectionSummary As String
    Dim viewsWithoutProjections As Long

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim sourceView As SldWorks.View
        Set sourceView = views(i)
        If sourceView Is Nothing Then GoTo ContinueView

        Set path = Module17_SectionPath.ResolveSectionPath( _
            graph, sourceView, sectionLabel, evidence)

        If path Is Nothing Then GoTo ContinueView

        If StrComp(path.RejectionReason, "NoProjectionsInView", _
            vbBinaryCompare) = 0 Then

            viewsWithoutProjections = viewsWithoutProjections + 1
            GoTo ContinueView
        End If

        evidence.AddInfo "R23_SECTION_PATH_CANDIDATE|" & path.Summary()

        If path.Resolved Then
            Set resolvedPath = path
            Exit For
        End If

        If Len(rejectionSummary) > 0 Then
            rejectionSummary = rejectionSummary & ";"
        End If
        rejectionSummary = rejectionSummary & path.SourceViewName & ":" & _
            path.RejectionReason

ContinueView:
    Next i

    If resolvedPath Is Nothing Then
        If Len(rejectionSummary) = 0 Then
            rejectionSummary = "NoCandidateViewHadProjections"
        End If

        evidence.AddFailure "R23 semantic section path was not resolved: " & _
            rejectionSummary & _
            " (viewsWithoutProjections=" & _
            CStr(viewsWithoutProjections) & ")."
        Exit Function
    End If

    Set path = resolvedPath

    Dim placeX As Double
    Dim placeY As Double
    placeX = evidence.UsableLeft + _
        (evidence.UsableRight - evidence.UsableLeft) * 0.75
    placeY = evidence.UsableBottom + _
        (evidence.UsableTop - evidence.UsableBottom) * 0.5

    Dim sectionView As SldWorks.View
    Set sectionView = Module17_SectionPath.CreateSectionFromPath( _
        swApp, swDrawModel, path, placeX, placeY, True, evidence)

    If sectionView Is Nothing Then Exit Function

    evidence.ViewsCreated = evidence.ViewsCreated + 1
    evidence.AddInfo "R23_SECTION_CREATE|source=" & path.SourceViewName & _
        "|section=" & path.SectionViewName & _
        "|placement=ProvisionalBeforeFinalEnvelopeLayout"
    CreateSemanticPrimarySection = True
    Exit Function

Failed:
    evidence.AddFailure "R23 semantic section error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

' R23-707. Read-only. The section line belongs to the SOURCE view, so moving
' that view re-lays it on the sheet and any geometry read before the move
' stops describing the drawing. Called immediately after the final structural
' layout so the recorded proof matches the placed sheet. It reports what
' GetSectionLineInfo2 returns and never repairs or re-cuts anything.
Private Sub RecordSectionLineAfterLayout( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Sub

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then Exit Sub

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        Dim proof As String
        proof = Module17_SectionPath.ReadSectionLineInfo(swView)

        ' Corrected after the 23:24 run: a view with no cut does NOT answer
        ' NoGeometryReturned as first assumed. All three views returned
        ' sectionLine=Read|values=0, so the value count is what distinguishes
        ' a real section line from an empty read.
        If InStr(1, proof, "sectionLine=Read", vbBinaryCompare) = 1 And _
           InStr(1, proof, "|values=0", vbBinaryCompare) = 0 Then
            evidence.AddInfo "R23_SECTION_LINE_POSTLAYOUT|view=" & _
                Module8_RuntimeSupport.GetViewName(swView) & "|" & proof

            ' R23-814. Read-only. The count alone has never said where
            ' SOLIDWORKS actually put the cut, which is why r51 could move
            ' the line 40 mm and produce a byte-identical section view
            ' without anything in evidence contradicting the intent.
            Module17_SectionPath.EmitSectionLineDecode swView, evidence
        End If

ContinueView:
    Next i
    Exit Sub

Failed:
    evidence.AddWarning "R23_SECTION_LINE_POSTLAYOUT|status=Unavailable" & _
        "|error=" & CStr(Err.Number)
End Sub

' R23-1101 steps 6 and 7. The import target order is section, then the
' non-primary orthographic views, then primary. It retains the Phase 0
' selected-view outcome and never offers the isometric as an import target.
Private Function ImportAndReconcileR23Annotations( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef primaryView As SldWorks.View, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim orderedViews As Collection
    Set orderedViews = CollectR23ImportViews(swDraw, primaryView)
    If orderedViews.Count = 0 Then
        evidence.AddFailure "R23 import has no eligible drawing views."
        Exit Function
    End If

    evidence.ImportedAnnotations = _
        Module14_AnnotationImport.ImportModelAnnotations( _
            swDrawModel, orderedViews, True, graph, evidence)

    If evidence.ImportedAnnotations < 0 Then Exit Function

    If Not Module14_AnnotationImport.BuildAnnotationInventory( _
        swDrawModel, graph, evidence) Then Exit Function

    Dim reconciled As Long
    reconciled = Module14_AnnotationImport.ReconcileWithLocationGraph( _
        swApp, swDrawModel, graph, evidence)

    Dim coverage As String
    coverage = Module14_AnnotationImport.VerifyRequiredCoverage(graph, _
        evidence)
    evidence.AddInfo "R23_IMPORT_RECONCILE|imported=" & _
        CStr(evidence.ImportedAnnotations) & "|reconciled=" & _
        CStr(reconciled) & "|coverage=" & coverage

    ' Coverage is intentionally observed before the creation stages.  Missing
    ' ordinates and native callouts are then created from the graph; making
    ' this pre-creation observation a hard stop would prevent that repair.
    ImportAndReconcileR23Annotations = True
    Exit Function

Failed:
    evidence.AddFailure "R23 annotation import error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CollectR23ImportViews( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef primaryView As SldWorks.View) As Collection

    Dim result As Collection
    Set result = New Collection
    Set CollectR23ImportViews = result

    Dim seenViews As Object
    Set seenViews = CreateObject("Scripting.Dictionary")

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Function

    Dim primaryName As String
    primaryName = Module8_RuntimeSupport.GetViewName(primaryView)

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then Exit Function

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueSection

        Dim reason As String
        If Module14_AnnotationImport.IsModelImportEligibleView( _
            swView, reason) And swView.Type = 2 Then

            AddR23ImportViewOnce result, seenViews, swView
        End If

ContinueSection:
    Next i

    For i = LBound(views) To UBound(views)
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueOther

        If StrComp(Module8_RuntimeSupport.GetViewName(swView), primaryName, _
            vbTextCompare) = 0 Then GoTo ContinueOther

        If swView.Type = 2 Then GoTo ContinueOther

        If Module14_AnnotationImport.IsModelImportEligibleView( _
            swView, reason) Then

            AddR23ImportViewOnce result, seenViews, swView
        End If

ContinueOther:
    Next i

    If Not primaryView Is Nothing Then
        If Module14_AnnotationImport.IsModelImportEligibleView( _
            primaryView, reason) Then

            AddR23ImportViewOnce result, seenViews, primaryView
        End If
    End If
    Exit Function

Failed:
    Set CollectR23ImportViews = result
End Function

' View names are unique on a drawing sheet. InsertModelAnnotations4 must see
' each selected view once, even when a primary view is itself a section.
Private Sub AddR23ImportViewOnce( _
    ByRef orderedViews As Collection, _
    ByRef seenViews As Object, _
    ByRef swView As SldWorks.View)

    If orderedViews Is Nothing Then Exit Sub
    If seenViews Is Nothing Then Exit Sub
    If swView Is Nothing Then Exit Sub

    Dim viewName As String
    viewName = Module8_RuntimeSupport.GetViewName(swView)
    If Len(Trim$(viewName)) = 0 Then Exit Sub
    If seenViews.Exists(viewName) Then Exit Sub

    seenViews.Add viewName, True
    orderedViews.Add swView
End Sub

' R23-506 and R23-508. Imported linear locations never suppress a required
' ordinate: location coverage and ordinate datum coverage are independent.
Private Function CreateR23OrdinateGroups( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Function

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then Exit Function

    ' The datum proof refuses to select while anything is already selected -
    ' deliberately, because IView.SelectEntity with AppendFlag=False would
    ' silently destroy an interactive caller's selection. Annotation import
    ' runs immediately before this stage and leaves its last inserted
    ' annotation selected, so every scheme in the r43 run was refused with
    ' selection=Reject|reason=PreexistingSelection|initialSelectionCount=1
    ' and both datums reported DatumNotProven - while the same datums
    ' resolved cleanly in the later QA pass, where nothing was selected.
    ' Clearing here is safe and is not a weakening of that guard: this is
    ' the production pipeline, which owns the document it just created. The
    ' guard itself is untouched and still protects the read-only probes.
    Dim preexistingSelection As Long
    preexistingSelection = swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)

    If preexistingSelection > 0 Then
        evidence.RecordSolidWorksMutation "ClearSelection2:BeforeOrdinates"
        swDrawModel.ClearSelection2 True
    End If

    evidence.AddInfo "R23_ORDINATE_SELECTION_PRECONDITION" & _
        "|preexisting=" & CStr(preexistingSelection) & _
        "|cleared=" & CStr(preexistingSelection > 0) & _
        "|remaining=" & _
        CStr(swDrawModel.SelectionManager.GetSelectedObjectCount2(-1))

    Dim createdGroups As Long
    Dim failedGroups As Long
    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        Dim schemes As Collection
        Set schemes = Module15_OrdinateScheme.BuildSchemesForView( _
            swApp, swDrawModel, swView, graph, evidence)

        Dim s As Long
        For s = 1 To schemes.Count
            Dim scheme As COrdinateScheme
            Set scheme = schemes(s)

            Dim created As Long
            created = Module15_OrdinateScheme.CreateOrdinateGroup( _
                swApp, swDrawModel, scheme, True, evidence)

            If created < 0 Then
                failedGroups = failedGroups + 1
            Else
                createdGroups = createdGroups + 1
                If scheme.Direction = _
                    Module15_OrdinateScheme.ORD_HORIZONTAL Then
                    evidence.HorizontalOrdinateGroups = _
                        evidence.HorizontalOrdinateGroups + 1
                ElseIf scheme.Direction = _
                    Module15_OrdinateScheme.ORD_VERTICAL Then

                    evidence.VerticalOrdinateGroups = _
                        evidence.VerticalOrdinateGroups + 1
                End If
                evidence.OrdinateEntitiesSelected = _
                    evidence.OrdinateEntitiesSelected + _
                        scheme.CreditedLocationCount()
            End If
        Next s

ContinueView:
    Next i

    evidence.AddInfo "R23_ORDINATE_CREATE|groups=" & _
        CStr(createdGroups) & "|failed=" & CStr(failedGroups)
    ' A view set can legitimately yield no ordinate scheme.  The semantic
    ' final gate decides whether required locations remain uncovered.
    CreateR23OrdinateGroups = (failedGroups = 0)
    Exit Function

Failed:
    evidence.AddFailure "R23 ordinate creation error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

' R23-803 and R23-811. The P-0251 section requirements are resolved by
' semantic identity before creation. This helper does not fabricate entity
' selections; a missing requirement remains a fail-closed production result.
Private Function ReconcileR23SectionDimensions( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If Module1_Main.GetFixtureKey(evidence.PartPath) <> _
        "P-0251-14A-001" Then

        ReconcileR23SectionDimensions = True
        Exit Function
    End If

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Function

    Dim views As Variant
    views = swSheet.GetViews

    Dim sectionViews As Collection
    Set sectionViews = _
        Module10_SectionDimensionEngine.CollectSectionViews(views)
    If sectionViews.Count <> 1 Then
        evidence.AddFailure _
            "R23 section-dimension reconciliation requires " & _
            "exactly one section view."
        Exit Function
    End If

    ' R23-812 evidence pass. Read-only. The section view has existed only
    ' since r49 and nothing has measured what geometry it exposes, so
    ' SECTION_DIMENSIONS reporting 7 missing / 0 satisfied says nothing
    ' about whether the geometry to satisfy them is even there.
    ' The drawing document is the same COM object as the DrawingDoc; the
    ' selection manager lives on the ModelDoc2 interface.
    Dim swDrawModel As SldWorks.ModelDoc2
    Set swDrawModel = swDraw

    ' The requirements come back carrying the section geometry that would
    ' measure each one, and the entity array those records index into.
    Dim requirements As Collection
    Dim sectionEntities As Variant
    Module10_SectionDimensionEngine.InventorySectionGeometry _
        swDrawModel, sectionViews(1), evidence, requirements, _
        sectionEntities

    If requirements Is Nothing Then
        Set requirements = _
            Module10_SectionDimensionEngine.BuildSectionRequirements()
    End If

    Dim sectionDimensions As Collection
    Set sectionDimensions = _
        Module10_SectionDimensionEngine.InventorySectionDimensions( _
            sectionViews(1), evidence)

    ' R23-802. Reconciliation runs BEFORE creation, always: creating a
    ' second dimension for a requirement an imported one already satisfies
    ' is the defect this ordering exists to prevent.
    Module10_SectionDimensionEngine.ReconcileSectionDimensions _
        requirements, sectionDimensions, evidence

    ' R23-821. Create only what reconciliation left missing and the
    ' inventory proved both present and selectable.
    Module10_SectionDimensionEngine.CreateResolvedSectionDimensions _
        swDrawModel, sectionViews(1), sectionEntities, requirements, _
        True, evidence

    ' R23-822. Re-read the view's own dimensions AND re-derive the
    ' requirement state from scratch. Re-reading the dimensions alone was not
    ' enough: reconciliation had already written "NoImportedDimension" into
    ' each requirement's Failures before creation ran, and creation sets
    ' Matched without clearing it, so VerifySectionDimensions counted every
    ' created dimension as RequirementFlagged. r60 reported satisfied=0 for
    ' five dimensions the QA stage independently proved satisfied. Module19
    ' reconciles fresh requirements against the finished drawing; that is the
    ' honest reading, and it is taken here too.
    Set sectionDimensions = _
        Module10_SectionDimensionEngine.InventorySectionDimensions( _
            sectionViews(1), evidence)

    Dim finalRequirements As Collection
    Set finalRequirements = _
        Module10_SectionDimensionEngine.BuildSectionRequirements()

    Module10_SectionDimensionEngine.ReconcileSectionDimensions _
        finalRequirements, sectionDimensions, evidence

    Dim verdict As String
    verdict = Module10_SectionDimensionEngine.VerifySectionDimensions( _
        finalRequirements, sectionDimensions)
    evidence.AddInfo "R23_SECTION_DIMENSIONS|" & verdict

    ReconcileR23SectionDimensions = _
        (InStr(1, verdict, "requirementFailures=None", _
            vbBinaryCompare) > 0)
    Exit Function

Failed:
    evidence.AddFailure "R23 section-dimension reconciliation error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

' R23-823. Runs the post-layout annotation clamp on the one section view,
' for the same fixture whose section dimensions this pipeline creates.
' Scoped that narrowly on purpose: the only annotations in that view are the
' ones this run created, and widening it would move imported dimensions the
' 2026-08-04 user decision preserves.
Private Sub ClampR23SectionAnnotations( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    If Module1_Main.GetFixtureKey(evidence.PartPath) <> _
        "P-0251-14A-001" Then Exit Sub

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Sub

    Dim views As Variant
    views = swSheet.GetViews

    Dim sectionViews As Collection
    Set sectionViews = _
        Module10_SectionDimensionEngine.CollectSectionViews(views)
    If sectionViews.Count <> 1 Then Exit Sub

    Module10_SectionDimensionEngine.ClampSectionAnnotationsIntoUsableArea _
        sectionViews(1), True, evidence
    Exit Sub

Failed:
    evidence.AddFailure "R23 section annotation clamp error " & _
        CStr(Err.Number) & ": " & Err.Description
End Sub

' R23-604. A native definition wins only when it is complete and attached.
' The fallback record names an incomplete native result; only then does the
' pipeline create one at a graph-proved anchor.
Private Function CreateMissingR23Callouts( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Function

    Dim views As Variant
    views = swSheet.GetViews

    Dim definitions As Collection
    Set definitions = Module19_SemanticQA.CollectRetainedDefinitions( _
        swApp, swDrawModel, views, graph, evidence)

    Dim requested As Long
    Dim created As Long
    Dim failed As Long
    Dim i As Long
    For i = 1 To definitions.Count
        Dim definition As CCalloutDefinition
        Set definition = definitions(i)
        If definition Is Nothing Then GoTo ContinueDefinition
        If definition.IsNative() Then GoTo ContinueDefinition
        ' CompletenessFailureReason already names every missing field, but
        ' only IsComplete() was consulted here, so the r46 QA report said
        ' "incomplete" for the M5 family without saying which field was
        ' missing. The reason is now carried into the failure text and into
        ' its own evidence line, the same correction the section path needed.
        If Not definition.IsComplete() Then
            failed = failed + 1

            Dim completenessReason As String
            completenessReason = definition.CompletenessFailureReason()

            evidence.AddFailure "R23 callout creation refused: incomplete " & _
                "source-backed definition for " & definition.FamilyKey & _
                "; missing=" & completenessReason
            evidence.AddInfo "R23_CALLOUT_INCOMPLETE|family=" & _
                Module16_CalloutDefinition.EvidenceToken( _
                    definition.FamilyKey) & _
                "|missing=" & _
                Module16_CalloutDefinition.EvidenceToken(completenessReason) & _
                "|" & definition.Summary()
            GoTo ContinueDefinition
        End If

        requested = requested + 1

        Dim projection As CViewHoleProjection
        Set projection = FirstAnchoredProjectionForFamily(graph, _
            definition.FamilyKey)
        If projection Is Nothing Then
            failed = failed + 1
            evidence.AddFailure _
                "R23 callout creation refused: no anchor for " & _
                definition.FamilyKey
            GoTo ContinueDefinition
        End If

        Dim callout As SldWorks.DisplayDimension
        Set callout = _
            Module16_CalloutDefinition.CreateNativeCalloutForFamily( _
            swApp, swDrawModel, projection, projection.PageX + 0.015, _
            projection.PageY + 0.015, True, evidence)

        If callout Is Nothing Then
            failed = failed + 1
        Else
            created = created + 1
        End If

ContinueDefinition:
    Next i

    evidence.AddInfo "R23_CALLOUT_CREATE|requested=" & CStr(requested) & _
        "|created=" & CStr(created) & "|failed=" & CStr(failed)
    CreateMissingR23Callouts = (failed = 0)
    Exit Function

Failed:
    evidence.AddFailure "R23 callout creation error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function FirstAnchoredProjectionForFamily( _
    ByRef graph As CLocationGraph, _
    ByVal familyKey As String) As CViewHoleProjection

    Dim projections As Collection
    Set projections = graph.Projections()

    Dim i As Long
    For i = 1 To projections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = projections(i)
        If candidate Is Nothing Then GoTo ContinueProjection
        If candidate.PhysicalLocation Is Nothing Then GoTo ContinueProjection
        If Not candidate.HasSelectableAnchor() Then GoTo ContinueProjection

        If StrComp(candidate.PhysicalLocation.SemanticFamilyKey, _
            familyKey, vbBinaryCompare) = 0 Then

            Set FirstAnchoredProjectionForFamily = candidate
            Exit Function
        End If

ContinueProjection:
    Next i
End Function

' Retains the automatic Phase 9 implementation for read-only diagnostics and
' history. R23_LAYOUT_USER_ACCEPTED_AS_IS prevents the production route and
' public mutation entry point from calling this procedure.
Private Function ApplyR23ContentLayout( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    If swSheet Is Nothing Then Exit Function

    Dim proof As String
    If Not Module18_ContentEnvelope.MeasureSheetRegions( _
        swSheet, evidence, proof) Then

        evidence.AddFailure "R23 final layout sheet measurement failed: " & _
            proof
        Exit Function
    End If

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then Exit Function

    Dim layoutViews As Collection
    Dim envelopes As Collection
    Set layoutViews = New Collection
    Set envelopes = New Collection

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView
        If Module18_ContentEnvelope.IsTemplateOrientationView(swView) Then
            GoTo ContinueView
        End If

        Dim envelope As CContentEnvelope
        Set envelope = Module18_ContentEnvelope.BuildViewEnvelope( _
            swView, evidence.SheetWidth, evidence.SheetHeight, evidence)
        If Not envelope.Seeded Then
            evidence.AddFailure _
                "R23 final layout has an unseeded envelope: " & _
                envelope.Name
            Exit Function
        End If

        layoutViews.Add swView
        envelopes.Add envelope

ContinueView:
    Next i

    If layoutViews.Count = 0 Then Exit Function

    Dim targetCentres As Collection
    Dim planProof As String
    Dim suggestedScale As Double
    Set targetCentres = Module18_ContentEnvelope.PlanPlacement( _
        envelopes, evidence, planProof, suggestedScale)

    Dim layoutProof As String
    layoutProof = Module18_ContentEnvelope.ApplyPlacementPlan( _
        swDrawModel, layoutViews, envelopes, targetCentres, _
        evidence.SheetWidth, evidence.SheetHeight, True, evidence)

    evidence.AddInfo "R23_FINAL_LAYOUT|" & layoutProof
    If InStr(1, layoutProof, "layout=Applied", vbBinaryCompare) = 0 Then
        Exit Function
    End If

    Dim sealedProof As String
    sealedProof = Module18_ContentEnvelope.VerifyNothingCreatedAfterLayout( _
        evidence)
    evidence.AddInfo "R23_LAYOUT_SEAL|" & sealedProof
    ApplyR23ContentLayout = _
        (InStr(1, sealedProof, "postLayout=True", vbBinaryCompare) > 0)
    Exit Function

Failed:
    evidence.AddFailure "R23 final layout error " & CStr(Err.Number) & _
        ": " & Err.Description
End Function

' The user-approved as-is policy replaces automatic envelope placement. The
' mutation ledger is sealed after all drawing definitions, making a later
' unexpected creation visible without claiming automatic clearance.
Private Function RecordR23UserAcceptedLayout( _
    ByRef evidence As CRunEvidence) As Boolean

    Module18_ContentEnvelope.SealLayout evidence

    Dim sealedProof As String
    sealedProof = Module18_ContentEnvelope.VerifyNothingCreatedAfterLayout( _
        evidence)

    ' These two counters name the retired content-envelope pass only. The
    ' approved geometry-first structural grid runs at step 11b and its moves
    ' are reported separately as evidence.LayoutMoves; conflating the two
    ' would make this line contradict the run's own COUNTS block.
    evidence.AddWarning "R23_FINAL_LAYOUT|layout=UserAcceptedAsIs" & _
        "|automaticClearance=DeferredByUser" & _
        "|automaticScaleChanges=0|automaticViewMoves=0" & _
        "|structuralGridMoves=" & CStr(evidence.LayoutMoves)
    evidence.AddInfo "R23_LAYOUT_SEAL|" & sealedProof

    RecordR23UserAcceptedLayout = _
        (InStr(1, sealedProof, "postLayout=True", vbBinaryCompare) > 0)
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
    evidence.AddInfo _
        "VIEW|role=Primary|orientation=*Front|eligibleOrdinate=True"

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
        evidence.AddFailure "Failed to create model view " & modelViewName _
            & "."
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
    evidence.AddInfo "VIEW|role=" & IIf(isIsometric, "OrientationAid", _
        "Orthographic") & _
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
        evidence.AddFailure _
            "Cannot create isometric view without a current sheet."
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

                evidence.AddFailure _
                    "P-0251 isometric view rejected its approved " & _
                    "orientation-aid scale of 1:2."
                CreateIsometricView = False
                Exit Function
            End If

            evidence.AddInfo "VIEW_SCALE|role=OrientationAid|fixture=" & _
                "P-0251-14A-001|ratio=1:2|independent=True"
        End If

        evidence.AddInfo _
            "VIEW_POLICY|isometricCreatedBeforeSelectedViewImport=True"
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
    evidence.AddFailure "Orthographic view configuration error: " & _
        Err.Description
End Function

Private Function ApplyFixturePrimaryViewRotation( _
    ByRef primaryView As SldWorks.View, _
    ByVal fixtureKey As String, _
    ByRef evidence As CRunEvidence) As Boolean

    Const P0251_PRIMARY_CLOCKWISE_90_RAD As Double = -1.5707963267949
    Const ANGLE_READBACK_TOLERANCE_RAD As Double = 0.000001

    On Error GoTo Failed

    If primaryView Is Nothing Then
        evidence.AddFailure _
            "Primary-view rotation was requested without a view."
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
        evidence.AddFailure _
            "P-0251 primary-view 90-degree rotation readback " & _
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
    evidence.AddFailure "Isometric view configuration error: " & _
        Err.Description
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
        Trim$(sourceView.GetOrientationName), "*Bottom", vbTextCompare) <> _
            0 Then

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

        evidence.AddFailure _
            "Pump Holder detail source references the wrong model."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created *Bottom view does not reference the authorized part " & _
                "object"
        Exit Function
    End If

    If StrComp( _
        Trim$(sourceView.ReferencedConfiguration), _
        Trim$(evidence.ConfigurationName), _
        vbTextCompare) <> 0 Then

        evidence.AddFailure "Pump Holder detail source configuration mismatch."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "created *Bottom view does not reference the approved " & _
                "configuration"
        Exit Function
    End If

    Dim currentSheet As SldWorks.Sheet
    Dim sourceSheet As SldWorks.Sheet
    Set currentSheet = swDraw.GetCurrentSheet
    Set sourceSheet = sourceView.Sheet
    If currentSheet Is Nothing Or sourceSheet Is Nothing Then
        evidence.AddFailure _
            "Pump Holder detail source sheet ownership is missing."
        evidence.MarkStageFailed "REQUIRED_DETAILS_STRUCTURE", _
            "current or source sheet readback is Nothing"
        Exit Function
    End If
    If Not Module8_RuntimeSupport.ObjectsAreSame( _
        swApp, currentSheet, sourceSheet) Then

        evidence.AddFailure _
            "Pump Holder detail source belongs to another sheet."
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
        "Details C and D created from reference-led *Bottom profiles at " & _
            "independent 3:1 scale"
    evidence.AddWarning "REQUIRED_DETAILS_GEOMETRY|status=PENDING" & _
        "|evidenceLevel=E4|profileCoordinateAndFeatureContainment=Unproved"
    evidence.AddWarning "REQUIRED_DETAILS_LEGIBILITY|status=PENDING" & _
        "|evidenceLevel=E6_E7|7x4AndC0.5DimensionsCollisionsReadability=Un" & _
            "proved"
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
        Trim$(sourceView.GetOrientationName), "*Bottom", vbTextCompare) <> _
            0 Then

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
        evidence.AddFailure _
            "CreateCircleByRadius returned Nothing for Detail " & _
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
        evidence.AddFailure _
            "CreateDetailViewAt4 returned Nothing for Detail " & _
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
        detailView, detailLabel, placementSheetX, placementSheetY, _
            evidence) Then
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
            evidence.AddFailure _
                "IDetailCircle.SetLabel returned False for Detail " & _
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

    If swDrawModel Is Nothing Or swDraw Is Nothing Or primaryView Is _
        Nothing Then
        evidence.AddFailure "SECTION_STEP|step=" & sectionStep & _
            "|status=Failed|reason=RequiredDrawingContextIsNothing"
        Exit Function
    End If

    If Module1_Main.GlobalSectionCount <> 1 Then
        evidence.AddFailure _
            "Exactly one primary section configuration was expected."
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
            evidence.AddFailure _
                "No approved primary-section plan exists for " & _
                fixtureKey & "."
            GoTo SafeExit
    End Select

    Dim sectionSegments As Collection
    Set sectionSegments = New Collection

    Dim targetX As Double
    Dim targetY As Double
    Dim offsetX As Double

    If fixtureKey = "P-0251-14A-001" Then
        evidence.AddFailure "Retired section path invoked for P-0251. " & _
            "Use Module17_SectionPath through the R23 pipeline."
        GoTo SafeExit
    Else
        Dim outsideMargin As Double
        outsideMargin = 0.1 * IIf(width > height, width, height)
        targetY = bottomY + height * (15.84 / 24#)

        If Not AddSectionSegment(swDrawModel, primaryView, sectionSegments, _
            leftX - outsideMargin, targetY, rightX + outsideMargin, targetY, _
            evidence) Then
            GoTo SafeExit
        End If

        evidence.AddInfo _
            "SECTION_PLAN|label=B|intent=UpperEarThreadedHoles" & _
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
        evidence.AddFailure _
            "CreateSelectData returned Nothing for section selection."
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
            evidence.AddFailure _
                "Section segment Select4 returned False at index " & _
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
            evidence.AddFailure _
                "Section selection contains Nothing at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        Dim expectedSegment As SldWorks.SketchSegment
        Set expectedSegment = sectionSegments(selectionIndex)
        If Not selectedObject Is expectedSegment Then
            evidence.AddFailure _
                "Section selection order mismatch at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        Dim selectedView As SldWorks.View
        Set selectedView = selectionManager.GetSelectedObjectsDrawingView2( _
            selectionIndex, -1)

        If selectedView Is Nothing Then
            evidence.AddFailure _
                "Section selection has no drawing-view owner at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        If StrComp( _
            Module8_RuntimeSupport.GetViewName(selectedView), _
            expectedViewName, vbTextCompare) <> 0 Then

            evidence.AddFailure _
                "Section selection belongs to the wrong view at index " & _
                CStr(selectionIndex) & "."
            Exit Function
        End If

        If selectionManager.GetSelectedObjectMark(selectionIndex) <> 0 Then
            evidence.AddFailure _
                "Section selection has an unexpected mark at index " & _
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
        evidence.AddFailure _
            "SketchManager.CreateLine returned Nothing for section."
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

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub

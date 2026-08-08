Option Explicit

' Programmatic compile gate for the trunk.
'
' Ported 2026-08-05 from archive/target-spec-hybrid-v2/Module20_ProbeRunner.bas
' when the baseline snapshot became the trunk. Only the infrastructure that
' tools/probe-runner and tools/production-runner actually invoke was carried
' across: the VBE compile route, the production pre-flight, and their helpers.
'
' Deliberately NOT ported: R23_RunAllProbes and the nine R23_Probe* entry
' points. Those called Module10..Module19, which were part of the archived
' target-spec work and do not exist here. Reintroduce a probe entry point only
' when there is a read-only question the trunk can actually answer.
'
' This module never mutates a drawing. It contains no allowMutation argument
' and no call to Module1_Main.main; the production runner invokes main as a
' separate transaction, and only after reading verdict=Clean out of this log.
'
' The project must reference "Microsoft Visual Basic for Applications
' Extensibility 5.3" so the intrinsic VBE object is available. That reference
' and the VBE.CommandBars route are already proved in this repository by
' tools/swp-deploy/Module0_SourceDeployment.bas.

Private Const MAX_CONTROL_WALK_DEPTH As Long = 6

' Verified against the installed SOLIDWORKS 2025 SP1.2 type library.
Private Const swDocPART As Long = 1
Private Const swRebuildActiveDoc As Long = 2

' Enumerates every VBE command bar control's Id and Caption. Diagnostic only;
' run once to identify the compile control by evidence rather than by a
' remembered ID, per the solidworks-api-lookup finding that VBIDE CommandBars
' control IDs are outside the SOLIDWORKS API corpus.
Public Sub R23_EnumerateVbeControls()
    On Error GoTo Failed

    ' Self-contained when run standalone: open the log here if none is
    ' already open, and close only the log this call opened.
    Dim openedHere As Boolean
    If Not Module21_EvidenceSink.IsOpen() Then
        Dim swApp As SldWorks.SldWorks
        Set swApp = Application.SldWorks

        Module21_EvidenceSink.OpenLog ResolveLogAnchorPath(swApp)
        openedHere = True
    End If

    Module21_EvidenceSink.LogLine "R23_VBE_ENUM_BEGIN"

    Dim vbeInstance As Object
    Set vbeInstance = VBE

    Dim dummyFound As Object
    Dim barIndex As Long
    Dim barCount As Long
    barCount = vbeInstance.CommandBars.Count

    For barIndex = 1 To barCount
        Dim bar As Object
        Set bar = vbeInstance.CommandBars(barIndex)

        Module21_EvidenceSink.LogLine _
            "R23_VBE_BAR|index=" & CStr(barIndex) & _
            "|name=" & CleanControlText(CStr(bar.Name))

        WalkVbeControls bar.Controls, 0, True, vbNullString, dummyFound
    Next barIndex

    Module21_EvidenceSink.LogLine "R23_VBE_ENUM_END|bars=" & _
        CStr(barCount)
    If openedHere Then Module21_EvidenceSink.CloseLog
    Exit Sub

Failed:
    Module21_EvidenceSink.LogLine "R23_VBE_ENUM_FATAL|error=" & _
        CStr(Err.Number) & "|description=" & Err.Description
    If openedHere Then Module21_EvidenceSink.CloseLog
End Sub

' Resolves the VBE Compile control by caption, executes it, and reports a
' verdict derived from the control's enabled state. VBE disables Compile once
' the project is fully compiled, so still-enabled after Execute means the
' project is not clean. Never guesses a control ID: an unresolved control
' fails named after a full enumeration dump.
Public Function R23_CompileProject() As String
    On Error GoTo Failed

    Dim openedHere As Boolean
    If Not Module21_EvidenceSink.IsOpen() Then
        Dim swApp As SldWorks.SldWorks
        Set swApp = Application.SldWorks

        Module21_EvidenceSink.OpenLog ResolveLogAnchorPath(swApp)
        openedHere = True
    End If

    Dim compileControl As Object
    Set compileControl = FindVbeControlByCaption("compile")

    If compileControl Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_COMPILE_FATAL|reason=ControlNotFound"
        R23_EnumerateVbeControls
        R23_CompileProject = "ControlNotFound"
        If openedHere Then Module21_EvidenceSink.CloseLog
        Exit Function
    End If

    Dim resolvedId As Long
    Dim resolvedCaption As String
    resolvedId = compileControl.Id
    resolvedCaption = CleanControlText(CStr(compileControl.Caption))

    Dim enabledBefore As Boolean
    enabledBefore = CBool(compileControl.Enabled)

    Module21_EvidenceSink.LogLine _
        "R23_COMPILE_CONTROL|id=" & CStr(resolvedId) & _
        "|caption=" & resolvedCaption & _
        "|enabledBefore=" & CStr(enabledBefore)

    compileControl.Execute
    DoEvents

    Dim compileControlAfter As Object
    Set compileControlAfter = FindVbeControlByCaption("compile")

    Dim enabledAfter As Boolean
    Dim verdict As String

    If compileControlAfter Is Nothing Then
        verdict = "Unknown:ControlUnresolvedAfterExecute"
    Else
        enabledAfter = CBool(compileControlAfter.Enabled)
        If enabledAfter Then
            verdict = "NotClean"
        Else
            verdict = "Clean"
        End If
    End If

    Module21_EvidenceSink.LogLine _
        "R23_COMPILE_VERDICT|id=" & CStr(resolvedId) & _
        "|caption=" & resolvedCaption & _
        "|enabledBefore=" & CStr(enabledBefore) & _
        "|enabledAfter=" & CStr(enabledAfter) & _
        "|verdict=" & verdict

    R23_CompileProject = verdict
    If openedHere Then Module21_EvidenceSink.CloseLog
    Exit Function

Failed:
    Module21_EvidenceSink.LogLine "R23_COMPILE_FATAL|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & "|description=" & Err.Description
    R23_CompileProject = "Error:" & CStr(Err.Number)
    If openedHere Then Module21_EvidenceSink.CloseLog
End Function

Private Function FindVbeControlByCaption( _
    ByVal captionFilter As String) As Object

    On Error GoTo Failed

    Dim vbeInstance As Object
    Set vbeInstance = VBE

    Dim foundControl As Object
    Dim barIndex As Long
    For barIndex = 1 To vbeInstance.CommandBars.Count
        Dim bar As Object
        Set bar = vbeInstance.CommandBars(barIndex)

        WalkVbeControls bar.Controls, 0, False, captionFilter, foundControl
        If Not foundControl Is Nothing Then Exit For
    Next barIndex

    Set FindVbeControlByCaption = foundControl
    Exit Function

Failed:
    Set FindVbeControlByCaption = Nothing
End Function

' Recursively walks one CommandBar controls collection. When logAll is True
' every visited control is logged as evidence (used by the enumeration probe).
' When captionFilter is non-empty, the first control whose caption contains it
' (case-insensitive) is returned in foundControl.
Private Sub WalkVbeControls( _
    ByRef controls As Object, _
    ByVal depth As Long, _
    ByVal logAll As Boolean, _
    ByVal captionFilter As String, _
    ByRef foundControl As Object)

    If depth > MAX_CONTROL_WALK_DEPTH Then Exit Sub

    On Error Resume Next
    Dim controlCount As Long
    controlCount = controls.Count
    On Error GoTo 0

    Dim i As Long
    For i = 1 To controlCount
        On Error Resume Next
        Dim ctrl As Object
        Set ctrl = Nothing
        Set ctrl = controls(i)
        On Error GoTo 0

        If Not ctrl Is Nothing Then
            Dim ctrlId As Long
            Dim ctrlCaption As String
            ctrlId = 0
            ctrlCaption = vbNullString

            On Error Resume Next
            ctrlId = ctrl.Id
            ctrlCaption = CStr(ctrl.Caption)
            On Error GoTo 0

            ' Live finding 2026-08-04: VBIDE CommandBarControl.Caption
            ' includes the raw "&" accelerator-key marker (e.g.
            ' "Compi&le Fable"), so the caption match must run against the
            ' cleaned text -- matching the raw string silently never matches
            ' any caption whose accelerator letter falls inside the filter.
            Dim cleanedCaption As String
            cleanedCaption = CleanControlText(ctrlCaption)

            If logAll Then
                Module21_EvidenceSink.LogLine _
                    "R23_VBE_CONTROL|depth=" & CStr(depth) & _
                    "|id=" & CStr(ctrlId) & _
                    "|caption=" & cleanedCaption
            End If

            If foundControl Is Nothing And Len(captionFilter) > 0 Then
                If InStr(1, cleanedCaption, captionFilter, _
                    vbTextCompare) > 0 Then Set foundControl = ctrl
            End If

            Dim childControls As Object
            Set childControls = Nothing
            On Error Resume Next
            Set childControls = ctrl.Controls
            On Error GoTo 0

            If Not childControls Is Nothing Then
                WalkVbeControls childControls, depth + 1, logAll, _
                    captionFilter, foundControl
            End If
        End If
    Next i
End Sub

Private Function CleanControlText(ByVal rawText As String) As String
    Dim cleaned As String
    cleaned = Replace$(rawText, "|", "/")
    cleaned = Replace$(cleaned, vbCr, " ")
    cleaned = Replace$(cleaned, vbLf, " ")
    cleaned = Replace$(cleaned, "&", vbNullString)
    CleanControlText = Trim$(cleaned)
End Function

' The standalone diagnostic entry points need a path under \test_assets\models\
' to derive the evidence log location, but ActiveDoc is whatever the user last
' activated -- frequently a drawing, which lives outside test_assets entirely.
' Prefer an already-open authorized part; fall back to ActiveDoc only when no
' authorized part is open.
Private Function ResolveLogAnchorPath( _
    ByRef swApp As SldWorks.SldWorks) As String

    If swApp Is Nothing Then Exit Function

    Dim authorizedPart As SldWorks.ModelDoc2
    Set authorizedPart = FindOpenAuthorizedPart(swApp)
    If Not authorizedPart Is Nothing Then
        ResolveLogAnchorPath = authorizedPart.GetPathName
        Exit Function
    End If

    If Not swApp.ActiveDoc Is Nothing Then
        ResolveLogAnchorPath = swApp.ActiveDoc.GetPathName
    End If
End Function

Private Function FindOpenAuthorizedPart( _
    ByRef swApp As SldWorks.SldWorks) As SldWorks.ModelDoc2

    On Error GoTo Failed

    Dim doc As Object
    Set doc = swApp.GetFirstDocument

    Do While Not doc Is Nothing
        If doc.GetType = swDocPART Then
            If Module1_Main.IsAuthorizedFixture(doc.GetPathName) Then
                Set FindOpenAuthorizedPart = doc
                Exit Function
            End If
        End If
        Set doc = doc.GetNext
    Loop
    Exit Function

Failed:
    Set FindOpenAuthorizedPart = Nothing
End Function

Private Function ActivateDocumentByTitle( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal documentTitle As String) As Boolean

    On Error GoTo Failed

    Dim activateErrors As Long
    Dim activated As Object
    Set activated = swApp.ActivateDoc3( _
        documentTitle, False, swRebuildActiveDoc, activateErrors)

    ActivateDocumentByTitle = _
        (Not activated Is Nothing) And (activateErrors = 0)

    Dim activeTitle As String
    activeTitle = "None"
    If Not swApp.ActiveDoc Is Nothing Then
        activeTitle = swApp.ActiveDoc.GetTitle
    End If

    If StrComp(activeTitle, documentTitle, vbTextCompare) <> 0 Then
        ActivateDocumentByTitle = False
    End If

    Module21_EvidenceSink.LogLine _
        "R23_RUN_ACTIVATE|title=" & CleanControlText(documentTitle) & _
        "|activeTitle=" & CleanControlText(activeTitle) & _
        "|errors=" & CStr(activateErrors) & _
        "|succeeded=" & CStr(ActivateDocumentByTitle)
    Exit Function

Failed:
    Module21_EvidenceSink.LogLine _
        "R23_RUN_ACTIVATE|title=" & CleanControlText(documentTitle) & _
        "|error=" & CStr(Err.Number) & "|succeeded=False"
    ActivateDocumentByTitle = False
End Function

' Calls one no-op per deployed standard module in the manifest. VBA compiles
' at module granularity, so a module that loads this Sub has compiled. Returns
' an empty string when every module loaded, otherwise the name of the first
' module that failed.
'
' Note the limit recorded in Agents.md: a genuine VBA compile error is raised
' before VBA enters any error handler, so this cannot localise one. It catches
' load failures only. R23_CompileProject's verdict is the real gate.
Public Function R23_TouchAllModules() As String
    On Error GoTo M1Failed
    Module1_Main.R23_CompileTouch
    On Error GoTo M2Failed
    Module2_DrawingPipeline.R23_CompileTouch
    On Error GoTo M3Failed
    Module3_ModelAudit.R23_CompileTouch
    On Error GoTo M4Failed
    Module4_ModelItemImporter.R23_CompileTouch
    On Error GoTo M5Failed
    Module5_FallbackDimensionEngine.R23_CompileTouch
    On Error GoTo M6Failed
    Module6_QAEngine.R23_CompileTouch
    On Error GoTo M7Failed
    Module7_TitleBlockEngine.R23_CompileTouch
    On Error GoTo M8Failed
    Module8_ViewClassifier.R23_CompileTouch
    On Error GoTo M9Failed
    Module9_HoleCalloutEngine.R23_CompileTouch
    On Error GoTo M10Failed
    Module10_SheetLayoutEngine.R23_CompileTouch
    On Error GoTo M20Failed
    Module20_ProbeRunner.R23_CompileTouch
    On Error GoTo M21Failed
    Module21_EvidenceSink.R23_CompileTouch

    On Error GoTo 0
    R23_TouchAllModules = vbNullString
    Exit Function

M1Failed: R23_TouchAllModules = "Module1_Main": Exit Function
M2Failed: R23_TouchAllModules = "Module2_DrawingPipeline": Exit Function
M3Failed: R23_TouchAllModules = "Module3_ModelAudit": Exit Function
M4Failed: R23_TouchAllModules = "Module4_ModelItemImporter": Exit Function
M5Failed:
    R23_TouchAllModules = "Module5_FallbackDimensionEngine"
    Exit Function
M6Failed: R23_TouchAllModules = "Module6_QAEngine": Exit Function
M7Failed: R23_TouchAllModules = "Module7_TitleBlockEngine": Exit Function
M8Failed: R23_TouchAllModules = "Module8_ViewClassifier": Exit Function
M9Failed: R23_TouchAllModules = "Module9_HoleCalloutEngine": Exit Function
M10Failed: R23_TouchAllModules = "Module10_SheetLayoutEngine": Exit Function
M20Failed: R23_TouchAllModules = "Module20_ProbeRunner": Exit Function
M21Failed: R23_TouchAllModules = "Module21_EvidenceSink": Exit Function
End Function

' Pre-flight for an authorized mutating production run. STRICTLY READ-ONLY:
' it contains no call to Module1_Main.main and creates nothing.
'
' The manual "Debug > Compile Project" gate that used to precede a mutating
' run exists for one reason: VBA compiles lazily, so a module with an error
' can sit undetected until the moment it is first called - which may be after
' several views and dozens of dimensions already exist. That leaves a
' half-built drawing and an unreadable failure.
'
' This performs the same check programmatically. The verdict is not a claim:
' it is the VBE Compile control's enabled-state flip. It then activates the
' authorized part so main() finds it as ActiveDoc.
'
' main() is deliberately NOT called from here. It is invoked as a separate
' transaction, and only after the caller has read verdict=Clean out of this
' log, so a dirty compile can never reach a mutating run.
Public Sub R23_PrepareProductionRun()
    On Error GoTo PreflightFatal

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    Dim swPartDoc As SldWorks.ModelDoc2
    If Not swApp Is Nothing Then
        Set swPartDoc = FindOpenAuthorizedPart(swApp)
    End If

    Dim anchorPath As String
    If Not swPartDoc Is Nothing Then anchorPath = swPartDoc.GetPathName

    Dim logPath As String
    logPath = Module21_EvidenceSink.OpenLog(anchorPath)

    Module21_EvidenceSink.LogLine "R23_PREFLIGHT_BEGIN|logPath=" & logPath & _
        "|revision=" & Module1_Main.MACRO_SOURCE_REVISION & _
        "|mode=ReadOnly|creations=0"

    Dim compileVerdict As String
    compileVerdict = R23_CompileProject()

    Module21_EvidenceSink.LogLine _
        "R23_PREFLIGHT_COMPILE|verdict=" & compileVerdict

    If StrComp(compileVerdict, "Clean", vbTextCompare) <> 0 Then
        ' A compile error is raised before VBA enters any error handler, so
        ' no touch routine can truthfully name its module. The VBE dialog and
        ' its highlighted line are the only diagnostic evidence.
        Module21_EvidenceSink.LogLine _
            "R23_PREFLIGHT_COMPILE_GUIDANCE" & _
            "|action=ReadVbeDialogAndHighlightedLine"
        Module21_EvidenceSink.LogLine _
            "R23_PREFLIGHT_END|ready=False|reason=CompileNotClean" & _
            "|verdict=" & compileVerdict
        Module21_EvidenceSink.CloseLog
        Exit Sub
    End If

    If swPartDoc Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_PREFLIGHT_END|ready=False|reason=NoOpenAuthorizedPart"
        Module21_EvidenceSink.CloseLog
        Exit Sub
    End If

    Dim activated As Boolean
    activated = ActivateDocumentByTitle(swApp, swPartDoc.GetTitle)

    If Not activated Then
        Module21_EvidenceSink.LogLine _
            "R23_PREFLIGHT_END|ready=False|reason=PartActivationFailed" & _
            "|part=" & swPartDoc.GetPathName
        Module21_EvidenceSink.CloseLog
        Exit Sub
    End If

    ' main() refuses anything that is not one of the three authorized
    ' fixtures. Reporting the same fact here means the caller never invokes a
    ' mutating run that would only fail-closed inside a message box.
    Module21_EvidenceSink.LogLine _
        "R23_PREFLIGHT_END|ready=True|verdict=Clean" & _
        "|activePart=" & swPartDoc.GetPathName & _
        "|authorized=" & _
            CStr(Module1_Main.IsAuthorizedFixture(swPartDoc.GetPathName))
    Module21_EvidenceSink.CloseLog
    Exit Sub

PreflightFatal:
    Module21_EvidenceSink.LogLine _
        "R23_PREFLIGHT_END|ready=False|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & "|description=" & Err.Description
    Module21_EvidenceSink.CloseLog
End Sub

' Compile-failure localisation no-op. VBA compiles at module granularity, so
' a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub

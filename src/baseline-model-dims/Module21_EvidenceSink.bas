Option Explicit

' R23 probe-automation evidence sink.
' Debug.Print writes only to the VBE Immediate Window, and nothing outside
' the VBA host can read that buffer. OpenLog/LogLine/CloseLog tee every
' diagnostic line to a plain-text file under test_assets/iteration_evidence/
' so an external process can read the evidence of a probe run without a
' manual copy-paste. LogLine always keeps Debug.Print alongside the file
' write; behaviour in the VBA editor is unchanged.
' The file handle is opened, written and closed on every single LogLine
' call rather than held open for the run. A held handle risks losing
' buffered lines if a modal dialog or a crash interrupts the run; a run is
' at most a few hundred lines, so the per-line open/close cost is
' irrelevant next to losing the evidence.

Private mLogFilePath As String
Private mLogOpen As Boolean

' Opens a new probe-run log derived from an authorized fixture path.
' anchorPath must contain the "\test_assets\models\" marker (an authorized
' fixture's full path satisfies this). The log is created at
' test_assets/iteration_evidence/probe_runs/<yyyyMMdd_HHmmss>/probe_log.txt
' and the returned path is empty if the marker is absent or any directory
' could not be created.
Public Function OpenLog(ByVal anchorPath As String) As String
    On Error GoTo Failed

    Dim marker As String
    marker = "\test_assets\models\"

    Dim markerPosition As Long
    markerPosition = InStr(1, anchorPath, marker, vbTextCompare)
    If markerPosition = 0 Then Exit Function

    Dim testAssetsPath As String
    testAssetsPath = Left$(anchorPath, markerPosition - 1) & _
        "\test_assets"
    If Not DirectoryExists(testAssetsPath) Then Exit Function

    Dim evidenceRoot As String
    evidenceRoot = testAssetsPath & "\iteration_evidence"
    If Not DirectoryExists(evidenceRoot) Then MkDir evidenceRoot

    Dim probeRunsRoot As String
    probeRunsRoot = evidenceRoot & "\probe_runs"
    If Not DirectoryExists(probeRunsRoot) Then MkDir probeRunsRoot

    Dim runFolder As String
    runFolder = probeRunsRoot & "\" & Format$(Now, "yyyymmdd_hhmmss")
    If Not DirectoryExists(runFolder) Then MkDir runFolder

    mLogFilePath = runFolder & "\probe_log.txt"
    mLogOpen = True
    OpenLog = mLogFilePath
    Exit Function

Failed:
    mLogFilePath = vbNullString
    mLogOpen = False
    OpenLog = vbNullString
End Function

' Writes one line to the open log file and to the Immediate Window.
' Degrades to Debug.Print alone when no log is open, so every caller can
' route through LogLine unconditionally.
Public Sub LogLine(ByVal text As String)
    If mLogOpen Then
        Dim fileNum As Integer
        fileNum = FreeFile

        On Error Resume Next
        Open mLogFilePath For Append As #fileNum
        Print #fileNum, text
        Close #fileNum
        On Error GoTo 0
    End If

    Debug.Print text
End Sub

' Writes the QA summary to disk beside the probe logs.
'
' Added 2026-08-05 (r3). The trunk previously showed BuildRunSummary in a
' MsgBox and nowhere else, so the ordinate chain counts and any decoded
' swCreateOrdDimError_e code existed only until the operator clicked OK.
' The first live trunk run lost exactly that evidence. A MsgBox is not a
' record.
'
' The folder shape matches what tools/production-runner expects:
' test_assets/iteration_evidence/macro_qa/<yyyymmdd_hhmmss>_<part>/QA_REPORT.txt
' Returns the written path, or an empty string if it could not be written.
Public Function WriteQaReport( _
    ByVal anchorPath As String, _
    ByVal reportText As String) As String

    On Error GoTo Failed

    Dim marker As String
    marker = "\test_assets\models\"

    Dim markerPosition As Long
    markerPosition = InStr(1, anchorPath, marker, vbTextCompare)
    If markerPosition = 0 Then Exit Function

    Dim testAssetsPath As String
    testAssetsPath = Left$(anchorPath, markerPosition - 1) & "\test_assets"
    If Not DirectoryExists(testAssetsPath) Then Exit Function

    Dim evidenceRoot As String
    evidenceRoot = testAssetsPath & "\iteration_evidence"
    If Not DirectoryExists(evidenceRoot) Then MkDir evidenceRoot

    Dim qaRoot As String
    qaRoot = evidenceRoot & "\macro_qa"
    If Not DirectoryExists(qaRoot) Then MkDir qaRoot

    Dim runFolder As String
    runFolder = qaRoot & "\" & Format$(Now, "yyyymmdd_hhmmss") & "_" & _
        BaseNameWithoutExtension(anchorPath)
    If Not DirectoryExists(runFolder) Then MkDir runFolder

    Dim reportPath As String
    reportPath = runFolder & "\QA_REPORT.txt"

    Dim fileNum As Integer
    fileNum = FreeFile
    Open reportPath For Output As #fileNum
    Print #fileNum, reportText
    Close #fileNum

    WriteQaReport = reportPath
    Exit Function

Failed:
    WriteQaReport = vbNullString
End Function

Private Function BaseNameWithoutExtension(ByVal path As String) As String
    Dim fileNameOnly As String
    fileNameOnly = path

    Dim slashPosition As Long
    slashPosition = InStrRev(path, "\")
    If slashPosition > 0 Then fileNameOnly = Mid$(path, slashPosition + 1)

    Dim dotPosition As Long
    dotPosition = InStrRev(fileNameOnly, ".")
    If dotPosition > 1 Then fileNameOnly = Left$(fileNameOnly, dotPosition - 1)

    BaseNameWithoutExtension = fileNameOnly
End Function

Public Sub CloseLog()
    mLogFilePath = vbNullString
    mLogOpen = False
End Sub

Public Function IsOpen() As Boolean
    IsOpen = mLogOpen
End Function

Public Function LogFilePath() As String
    LogFilePath = mLogFilePath
End Function

Private Function DirectoryExists(ByVal path As String) As Boolean
    On Error Resume Next
    DirectoryExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub

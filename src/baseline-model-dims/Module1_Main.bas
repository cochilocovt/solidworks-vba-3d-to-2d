Option Explicit

Public Type SectionConfig
    Label As String
    Vertical As Boolean
End Type

Public Type DrawingConfig
    CreateFront As Boolean
    CreateTop As Boolean
    CreateBottom As Boolean
    CreateRight As Boolean
    CreateLeft As Boolean
    CreateBack As Boolean
    CreateIso As Boolean

    CreateSection As Boolean
    SectionCount As Long
    SectionLabel As String
    SectionVertical As Boolean

    UseModelDimensions As Boolean
    UseOrdinateDims As Boolean
    RunHybridStrategy As Boolean
    ImportHoleCallouts As Boolean
    DatumOrigin As String

    PopulateTitle As Boolean
    InsertBarcode As Boolean
    InsertNotes As Boolean
    AutoArrange As Boolean

    SheetScale As Double
    CustomScaleText As String
    UseHLR As Boolean
    ShowLayoutPreview As Boolean
    TotalCostManual As String

    GenerateQAReport As Boolean
    Cancelled As Boolean
End Type

Public GlobalConfig As DrawingConfig
Public GlobalSections(1 To 5) As SectionConfig
Public GlobalSectionCount As Long

' Single version source for the trunk. tools/swp-deploy regenerates
' deployment-request.txt from this; never hand-edit that file. Bump whenever
' deployable behaviour changes.
Public Const MACRO_SOURCE_REVISION As String = _
    "trunk-2026-08-06-r21"

' User-confirmed 2026-08-05. The baseline snapshot carried "V:\SW_data\..."
' with the VEEMAP segment missing, so the controlled template never resolved
' and GetValidDrawingTemplatePath silently fell through to the SOLIDWORKS
' default drawing template.
Private Const TEMPLATE_PATH As String = _
    "V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT"
Private Const swDocPART As Long = 1
Private Const swDefaultTemplateDrawing As Long = 10

' The three authorized fixtures. tools/production-runner and
' tools/probe-runner mirror this list and never widen it.
Private Const FIXTURE_1 As String = _
    "C:\Users\V.T\Documents\VBA 3D TO " & _
        "2D\test_assets\models\P-0251-14A-001.SLDPRT"
Private Const FIXTURE_2 As String = _
    "C:\Users\V.T\Documents\VBA 3D TO " & _
        "2D\test_assets\models\P-0252-01-001.SLDPRT"
Private Const FIXTURE_3 As String = _
    "C:\Users\V.T\Documents\VBA 3D TO " & _
        "2D\test_assets\models\P-0252-01-013.SLDPRT"

Public Sub main()
    On Error GoTo Failed

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "Could not connect to SOLIDWORKS.", vbCritical, "Macro Error"
        Exit Sub
    End If

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = swApp.ActiveDoc
    If swPart Is Nothing Then
        MsgBox "No active document. Please open a part file first.", vbExclamation, "Macro Error"
        Exit Sub
    End If

    If swPart.GetType <> swDocPART Then
        MsgBox "Please activate a Part document (.sldprt) and run again.", vbExclamation, "Macro Error"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then
        MsgBox "Please save the part before running the macro.", vbExclamation, "Macro Error"
        Exit Sub
    End If

    If Not IsAuthorizedFixture(partPath) Then
        MsgBox _
            "Fail-closed: this build runs only against the three " & _
            "authorized fixtures." & vbCrLf & partPath, _
            vbCritical, "Unauthorized Model"
        Exit Sub
    End If

    ResetGlobalConfig
    ResetGlobalSections

    UserForm1.Show

    If GlobalConfig.Cancelled Then Exit Sub

    Dim finalTemplatePath As String
    finalTemplatePath = GetValidDrawingTemplatePath(swApp)
    If Len(finalTemplatePath) = 0 Then
        MsgBox "No valid drawing template could be found.", vbCritical, "Template Error"
        Exit Sub
    End If

    Module2_DrawingPipeline.CreateDrawing swApp, swPart, partPath, finalTemplatePath
    Exit Sub

Failed:
    MsgBox "Drawing automation failed: " & Err.Description, vbCritical, "Macro Error"
End Sub

Public Sub ResetGlobalConfig()
    With GlobalConfig
        .CreateFront = True
        .CreateTop = False
        .CreateBottom = False
        .CreateRight = True
        .CreateLeft = False
        .CreateBack = False
        .CreateIso = True

        .CreateSection = False
        .SectionCount = 0
        .SectionLabel = vbNullString
        .SectionVertical = False

        ' Both producers always run. The form's mutually exclusive
        ' "Import Model Dims" / "Ordinate Dims" option group was removed on
        ' 2026-08-06: it made the two impossible to combine, and the
        ' reference drawing needs both at once.
        .UseModelDimensions = True
        .UseOrdinateDims = True
        .RunHybridStrategy = True
        .ImportHoleCallouts = True
        .DatumOrigin = "Center"

        .PopulateTitle = True
        .InsertBarcode = True
        .InsertNotes = True
        .AutoArrange = True

        .SheetScale = 1#
        .CustomScaleText = vbNullString
        ' NOTE: this value does NOT reach the operator. UserForm1 seeds its
        ' checkbox from a saved registry setting -
        ' chkHLR.Value = ReadBoolSetting("UseHLR", False) - and writes it back
        ' over GlobalConfig on OK, so the form wins for every run that shows
        ' it, which is every run. The same applies to the other fields below.
        ' ResetGlobalConfig is the no-form fallback, not the user-visible
        ' default. Changing what the operator sees means changing the form,
        ' which lives outside the deployment manifest.
        '
        ' HLR is still the right fallback. Evidence, 2026-08-06: run twice on the same r16
        ' binary, HLV against HLR, nothing else changed.
        '   HLV  64 edges -> X 9 stations, Y 7, TWO dangling ordinates
        '   HLR  39 edges -> X 5 stations, Y 5, zero dangling
        ' An edge drawn in hidden-line font is returned by
        ' GetVisibleEntities2, selects, and yields
        ' swCreateOrdDimErr_Success - and the ordinate dangles. Under HLR
        ' those edges are not in the view, so the failure cannot arise, and
        ' the station counts match the reference drawing exactly.
        '
        ' Tradeoff, stated rather than hidden: HLR also removes the stepped
        ' bore from the front view. That is what the reference drawing does
        ' too - it dimensions that bore in SECTION J-J, not in the front
        ' view. Untick HLR on the form to restore hidden-line display.
        .UseHLR = True
        .ShowLayoutPreview = True
        .TotalCostManual = vbNullString

        .GenerateQAReport = True
        .Cancelled = True
    End With
End Sub

Public Sub ResetGlobalSections()
    Dim i As Long

    For i = 1 To 5
        GlobalSections(i).Label = vbNullString
        GlobalSections(i).Vertical = False
    Next i

    GlobalSectionCount = 0
End Sub

Public Sub SyncSectionCompatibilityFields()
    GlobalConfig.SectionCount = GlobalSectionCount
    GlobalConfig.CreateSection = (GlobalSectionCount > 0)

    If GlobalSectionCount > 0 Then
        GlobalConfig.SectionLabel = GetSectionLabelOrDefault(1)
        GlobalConfig.SectionVertical = GlobalSections(1).Vertical
    Else
        GlobalConfig.SectionLabel = vbNullString
        GlobalConfig.SectionVertical = False
    End If
End Sub

Public Function GetSectionLabelOrDefault(ByVal index As Long) As String
    If index < 1 Or index > 5 Then
        GetSectionLabelOrDefault = "J"
        Exit Function
    End If

    GetSectionLabelOrDefault = UCase$(Left$(Trim$(GlobalSections(index).Label), 2))
    If Len(GetSectionLabelOrDefault) = 0 Then
        GetSectionLabelOrDefault = Chr$(Asc("J") + index - 1)
    End If
End Function

Public Function IsAuthorizedFixture(ByVal partPath As String) As Boolean
    Dim normalized As String
    normalized = LCase$(Replace$(Trim$(partPath), "/", "\"))

    IsAuthorizedFixture = _
        (normalized = LCase$(FIXTURE_1)) Or _
        (normalized = LCase$(FIXTURE_2)) Or _
        (normalized = LCase$(FIXTURE_3))
End Function

' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub

Private Function GetValidDrawingTemplatePath(ByRef swApp As SldWorks.SldWorks) As String
    Dim candidate As String

    candidate = TEMPLATE_PATH
    If Len(candidate) > 0 Then
        If Dir$(candidate) <> vbNullString Then
            GetValidDrawingTemplatePath = candidate
            Exit Function
        End If
    End If

    candidate = swApp.GetUserPreferenceStringValue(swDefaultTemplateDrawing)
    If Len(candidate) > 0 Then
        If Dir$(candidate) <> vbNullString Then
            GetValidDrawingTemplatePath = candidate
        End If
    End If
End Function

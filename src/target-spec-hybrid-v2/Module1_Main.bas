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
Public GlobalEvidence As CRunEvidence

Public Const MACRO_SOURCE_REVISION As String = _
    "target-spec-hybrid-v2-2026-07-29-r22"

' Temporary inspection mode. Keep False for production acceptance.
Public Const DIAGNOSTIC_DRAWING_MODE As Boolean = True

Private Const CONTROLLED_TEMPLATE_PATH As String = _
    "V:\VEEMAP\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT"

Private Const FIXTURE_1 As String = _
    "C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\models\P-0251-14A-001.SLDPRT"
Private Const FIXTURE_2 As String = _
    "C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\models\P-0252-01-001.SLDPRT"
Private Const FIXTURE_3 As String = _
    "C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\models\P-0252-01-013.SLDPRT"

Private Const swDocPART As Long = 1
Private Const swVerticalOrdinate As Long = 2
Private Const swHorizontalOrdinate As Long = 3

Public Sub main()
    Set GlobalEvidence = Nothing

    On Error GoTo Failed

    Dim mainErrorNumber As Long
    Dim mainErrorDescription As String

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        MsgBox "Could not connect to SOLIDWORKS.", vbCritical, "Macro Error"
        Exit Sub
    End If

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = swApp.ActiveDoc

    If swPart Is Nothing Then
        MsgBox "Open and activate one authorized test part first.", _
               vbExclamation, "Target-Spec Hybrid V2"
        Exit Sub
    End If

    If swPart.GetType <> swDocPART Then
        MsgBox "The active document must be a saved part.", _
               vbExclamation, "Target-Spec Hybrid V2"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Len(partPath) = 0 Then
        MsgBox "Save the authorized part before running the macro.", _
               vbExclamation, "Target-Spec Hybrid V2"
        Exit Sub
    End If

    If Not IsAuthorizedFixture(partPath) Then
        MsgBox "Fail-closed: this build runs only against the three authorized fixtures." & _
               vbCrLf & partPath, vbCritical, "Unauthorized Model"
        Exit Sub
    End If

    Set GlobalEvidence = New CRunEvidence
    InitializeEvidenceIdentity _
        swApp, swPart, partPath, CONTROLLED_TEMPLATE_PATH

    swPart.SetPickMode
    swPart.ClearSelection2 True

    ResetGlobalConfig
    ResetGlobalSections
    Module8_RuntimeSupport.ResetProvenViewRegistry
    UserForm1.Show

    If GlobalConfig.Cancelled Then Exit Sub

    ApplyFixtureAcceptanceProfile partPath

    InitializeEvidence swApp, swPart, partPath, CONTROLLED_TEMPLATE_PATH

    Dim templatePath As String
    templatePath = GetValidDrawingTemplatePath(swApp)

    If Len(templatePath) = 0 Then
        GlobalEvidence.AddFailure "Controlled drawing template is missing: " & _
            CONTROLLED_TEMPLATE_PATH
        Module6_QAEngine.EmitRunEvidence GlobalEvidence

        MsgBox "Fail-closed: no existing drawing template was found. " & _
               "The macro will not invent a sheet format.", _
               vbCritical, "Template Error"
        Exit Sub
    End If

    GlobalEvidence.TemplatePath = templatePath

    Module2_DrawingPipeline.CreateDrawing _
        swApp, swPart, partPath, templatePath, GlobalEvidence
    Exit Sub

Failed:
    mainErrorNumber = Err.Number
    mainErrorDescription = Err.Description

    On Error Resume Next
    If Not swPart Is Nothing Then
        swPart.SetPickMode
        swPart.ClearSelection2 True
    End If

    If Not GlobalEvidence Is Nothing Then
        GlobalEvidence.AddFailure "Unhandled main error " & _
            CStr(mainErrorNumber) & ": " & mainErrorDescription
        Module6_QAEngine.EmitRunEvidence GlobalEvidence
    End If

    MsgBox "Drawing automation failed: " & mainErrorDescription, _
           vbCritical, "Target-Spec Hybrid V2"
    On Error GoTo 0
End Sub

Public Function IsAuthorizedFixture(ByVal partPath As String) As Boolean
    Dim normalized As String
    normalized = LCase$(Replace$(Trim$(partPath), "/", "\"))

    IsAuthorizedFixture = _
        (normalized = LCase$(FIXTURE_1)) Or _
        (normalized = LCase$(FIXTURE_2)) Or _
        (normalized = LCase$(FIXTURE_3))
End Function

Public Function GetFixtureKey(ByVal partPath As String) As String
    Dim normalized As String
    normalized = LCase$(Replace$(Trim$(partPath), "/", "\"))

    If normalized = LCase$(FIXTURE_1) Then
        GetFixtureKey = "P-0251-14A-001"
    ElseIf normalized = LCase$(FIXTURE_2) Then
        GetFixtureKey = "P-0252-01-001"
    ElseIf normalized = LCase$(FIXTURE_3) Then
        GetFixtureKey = "P-0252-01-013"
    End If
End Function

Public Sub ApplyFixtureAcceptanceProfile(ByVal partPath As String)
    GlobalConfig.CreateFront = True
    GlobalConfig.CreateTop = False
    GlobalConfig.CreateBottom = False
    GlobalConfig.CreateRight = False
    GlobalConfig.CreateLeft = False
    GlobalConfig.CreateBack = False
    GlobalConfig.CreateIso = True

    GlobalConfig.UseModelDimensions = True
    GlobalConfig.UseOrdinateDims = True
    GlobalConfig.RunHybridStrategy = True
    GlobalConfig.PopulateTitle = True
    GlobalConfig.InsertBarcode = True
    GlobalConfig.InsertNotes = True
    GlobalConfig.GenerateQAReport = True

    ResetGlobalSections

    Select Case GetFixtureKey(partPath)
        Case "P-0251-14A-001"
            GlobalConfig.CreateLeft = True
            ConfigureRequiredSection "J", False

        Case "P-0252-01-001"
            GlobalConfig.CreateRight = True
            SyncSectionCompatibilityFields

        Case "P-0252-01-013"
            GlobalConfig.CreateBottom = True
            GlobalConfig.CreateLeft = True
            GlobalConfig.CreateRight = True
            ConfigureRequiredSection "B", False
    End Select
End Sub

Private Sub ConfigureRequiredSection( _
    ByVal sectionLabel As String, _
    ByVal sectionVertical As Boolean)

    GlobalSectionCount = 1
    GlobalSections(1).Label = sectionLabel
    GlobalSections(1).Vertical = sectionVertical
    SyncSectionCompatibilityFields
End Sub

Public Function GetDatumOriginForDirection( _
    ByVal partPath As String, _
    ByVal configuredOrigin As String, _
    ByVal ordinateDirection As Long) As String

    Dim requested As String
    requested = Trim$(configuredOrigin)
    If Len(requested) = 0 Then requested = "Bottom-Left"

    If GetFixtureKey(partPath) = "P-0251-14A-001" And _
       StrComp(requested, "Bottom-Left", vbTextCompare) = 0 Then

        If ordinateDirection = swHorizontalOrdinate Then
            GetDatumOriginForDirection = "Center"
        ElseIf ordinateDirection = swVerticalOrdinate Then
            GetDatumOriginForDirection = "Bottom-Left"
        Else
            GetDatumOriginForDirection = requested
        End If
    Else
        GetDatumOriginForDirection = requested
    End If
End Function

Private Sub InitializeEvidence( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String)

    With GlobalEvidence
        .RunId = Format$(Now, "yyyymmdd_hhnnss")
        .PartPath = partPath
        .TemplatePath = templatePath
        .DatumOrigin = GlobalConfig.DatumOrigin
        .SheetScale = GlobalConfig.SheetScale
        .SolidWorksRevision = swApp.RevisionNumber
        .MacroRevision = MACRO_SOURCE_REVISION
        .ConfigurationName = _
            swPart.ConfigurationManager.ActiveConfiguration.Name
        .SettingsSummary = BuildSettingsSummary()
        .AddInfo "Fixed workflow: model annotations, ownership-proven fallback ordinates, layout, title data, QA."
        .AddInfo "Feature tolerances are recorded in Module8_RuntimeSupport and require runtime evidence."
    End With
End Sub

Private Sub InitializeEvidenceIdentity( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String)

    With GlobalEvidence
        .RunId = Format$(Now, "yyyymmdd_hhnnss")
        .PartPath = partPath
        .TemplatePath = templatePath
        .SolidWorksRevision = swApp.RevisionNumber
        .MacroRevision = MACRO_SOURCE_REVISION
        .ConfigurationName = _
            swPart.ConfigurationManager.ActiveConfiguration.Name
    End With
End Sub

Private Function BuildSettingsSummary() As String
    BuildSettingsSummary = _
        "Front=True" & _
        ";Top=" & CStr(GlobalConfig.CreateTop) & _
        ";Bottom=" & CStr(GlobalConfig.CreateBottom) & _
        ";Right=" & CStr(GlobalConfig.CreateRight) & _
        ";Left=" & CStr(GlobalConfig.CreateLeft) & _
        ";Back=" & CStr(GlobalConfig.CreateBack) & _
        ";Iso=" & CStr(GlobalConfig.CreateIso) & _
        ";Section=" & CStr(GlobalConfig.CreateSection) & _
        ";SectionCount=" & CStr(GlobalSectionCount) & _
        ";Callouts=" & CStr(GlobalConfig.ImportHoleCallouts) & _
        ";Datum=" & GlobalConfig.DatumOrigin & _
        ";Scale=" & Format$(GlobalConfig.SheetScale, "0.###") & _
        ";HLR=" & CStr(GlobalConfig.UseHLR) & _
        ";Arrange=" & CStr(GlobalConfig.AutoArrange) & _
        ";Title=" & CStr(GlobalConfig.PopulateTitle) & _
        ";Notes=" & CStr(GlobalConfig.InsertNotes) & _
        ";PartId=" & CStr(GlobalConfig.InsertBarcode) & _
        ";Preview=" & CStr(GlobalConfig.ShowLayoutPreview)
End Function

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

        .UseModelDimensions = True
        .UseOrdinateDims = True
        .RunHybridStrategy = True
        .ImportHoleCallouts = True
        .DatumOrigin = "Bottom-Left"

        .PopulateTitle = True
        .InsertBarcode = True
        .InsertNotes = True
        .AutoArrange = True

        .SheetScale = 1#
        .CustomScaleText = vbNullString
        .UseHLR = False
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
    SyncSectionCompatibilityFields
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

    GetSectionLabelOrDefault = _
        UCase$(Left$(Trim$(GlobalSections(index).Label), 2))

    If Len(GetSectionLabelOrDefault) = 0 Then
        GetSectionLabelOrDefault = Chr$(Asc("J") + index - 1)
    End If
End Function

Public Function GetValidDrawingTemplatePath( _
    ByRef swApp As SldWorks.SldWorks) As String

    If FileExists(CONTROLLED_TEMPLATE_PATH) Then
        GetValidDrawingTemplatePath = CONTROLLED_TEMPLATE_PATH
    End If
End Function

Private Function FileExists(ByVal path As String) As Boolean
    If Len(Trim$(path)) = 0 Then Exit Function

    On Error Resume Next
    FileExists = (Len(Dir$(path, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0)
    On Error GoTo 0
End Function

Attribute VB_Name = "Module1_Main"
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

Private Const TEMPLATE_PATH As String = "V:\SW_data\Custom Templates\VEEMAP DRAWING.DRWDOT"

Private Const swDocPART As Long = 1
Private Const swDefaultTemplateDrawing As Long = 10

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

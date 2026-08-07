Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{CB612155-0A53-4C54-BBB0-A018142FB585}{FFA92355-C917-4226-BAA2-3FD9E9D37697}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private chkFront As Object
Private chkTop As Object
Private chkBottom As Object
Private chkRight As Object
Private chkLeft As Object
Private chkBack As Object
Private chkIso As Object

Private lstSections As Object
Private cmdAddSection As Object
Private cmdRemoveSection As Object

' The Import Model Dims / Ordinate Dims option group was removed 2026-08-06.
' They were mutually exclusive members of the "DimMode" group, so the two
' producers could never run in the same drawing - and the reference drawing
' P-0251-14A-001 needs both at once: native import supplies R36, the section
' diameters and the hole callouts, the ordinate engine supplies the two
' position chains. Both now always run; see Module1_Main.ResetGlobalConfig.
Private cmbDatum As Object
Private chkHoleCallouts As Object
Private lblHoleCount As Object

Private chkPopulateTitle As Object
Private chkBarcode As Object
Private chkNotes As Object
Private chkAutoArrange As Object
Private txtTotalCost As Object

Private cmbScale As Object
Private txtCustomScale As Object
Private chkHLR As Object
Private chkPreview As Object

Private cmdCreate As Object
Private cmdCancel As Object

Private hCreate As BtnHandler
Private hCancel As BtnHandler
Private hAddSection As SectionBtnHandler
Private hRemoveSection As SectionBtnHandler

Private mSections(1 To 5) As SectionConfig
Private mSectionCount As Long

Private Sub UserForm_Initialize()
    Me.caption = "Auto Drawing Settings"
    Me.Width = 390
    Me.Height = 470

    BuildUI
    LoadPreferences
    RefreshSectionList
    RefreshHoleCount
End Sub

Private Sub BuildUI()
    Dim f1 As Object
    Dim f2 As Object
    Dim f3 As Object
    Dim f4 As Object
    Dim f5 As Object

    Set f1 = Me.Controls.Add("Forms.Frame.1", "frmViews", True)
    f1.caption = "Drawing Views"
    f1.Left = 8
    f1.Top = 10
    f1.Width = 160
    f1.Height = 150

    Set chkFront = AddCheckBox(f1, "Front View (Primary)", 10, 15, 135, True, False)
    Set chkTop = AddCheckBox(f1, "Top View", 10, 35, 120, False, True)
    Set chkBottom = AddCheckBox(f1, "Bottom View", 10, 55, 120, False, True)
    Set chkRight = AddCheckBox(f1, "Right View", 10, 75, 120, True, True)
    Set chkLeft = AddCheckBox(f1, "Left View", 10, 95, 120, False, True)
    Set chkBack = AddCheckBox(f1, "Back View", 10, 115, 120, False, True)
    Set chkIso = AddCheckBox(f1, "Isometric View", 10, 135, 120, True, True)

    Set f2 = Me.Controls.Add("Forms.Frame.1", "frmSections", True)
    f2.caption = "Section Views (configurable, 0-5)"
    f2.Left = 178
    f2.Top = 10
    f2.Width = 192
    f2.Height = 150

    Set lstSections = f2.Controls.Add("Forms.ListBox.1", "lstSections", True)
    lstSections.Left = 10
    lstSections.Top = 18
    lstSections.Width = 170
    lstSections.Height = 82

    Set cmdAddSection = f2.Controls.Add("Forms.CommandButton.1", "cmdAddSection", True)
    cmdAddSection.caption = "Add Section..."
    cmdAddSection.Left = 10
    cmdAddSection.Top = 108
    cmdAddSection.Width = 78

    Set cmdRemoveSection = f2.Controls.Add("Forms.CommandButton.1", "cmdRemoveSection", True)
    cmdRemoveSection.caption = "Remove"
    cmdRemoveSection.Left = 102
    cmdRemoveSection.Top = 108
    cmdRemoveSection.Width = 78

    Set f3 = Me.Controls.Add("Forms.Frame.1", "frmDims", True)
    f3.caption = "Dimension Settings"
    f3.Left = 8
    f3.Top = 170
    f3.Width = 160
    f3.Height = 128

    AddLabel f3, "Model dims + ordinate chains both applied.", 10, 14, 145

    AddLabel f3, "Ordinate Datum Origin:", 10, 36, 120
    Set cmbDatum = f3.Controls.Add("Forms.ComboBox.1", "cmbDatum", True)
    cmbDatum.Left = 10
    cmbDatum.Top = 52
    cmbDatum.Width = 120
    cmbDatum.Style = 2
    cmbDatum.AddItem "Bottom-Left"
    cmbDatum.AddItem "Center"
    cmbDatum.AddItem "Top-Left"
    cmbDatum.ListIndex = 1

    Set chkHoleCallouts = AddCheckBox(f3, "Insert Hole Wizard Callouts", 10, 74, 140, True, True)
    Set lblHoleCount = AddLabel(f3, "Hole-like features detected: 0", 10, 90, 145)

    Set f4 = Me.Controls.Add("Forms.Frame.1", "frmOther", True)
    f4.caption = "Other Settings"
    f4.Left = 178
    f4.Top = 170
    f4.Width = 192
    f4.Height = 128

    Set chkPopulateTitle = AddCheckBox(f4, "Populate Title Block", 10, 16, 140, True, True)
    Set chkBarcode = AddCheckBox(f4, "Insert Barcode", 10, 36, 140, True, True)
    Set chkNotes = AddCheckBox(f4, "Insert General Notes", 10, 56, 140, True, True)
    Set chkAutoArrange = AddCheckBox(f4, "Auto-Arrange Dims", 10, 76, 140, True, True)

    AddLabel f4, "Total Cost (manual):", 10, 100, 90
    Set txtTotalCost = f4.Controls.Add("Forms.TextBox.1", "txtTotalCost", True)
    txtTotalCost.Left = 105
    txtTotalCost.Top = 96
    txtTotalCost.Width = 70

    Set f5 = Me.Controls.Add("Forms.Frame.1", "frmSheet", True)
    f5.caption = "Sheet Settings"
    f5.Left = 8
    f5.Top = 308
    f5.Width = 362
    f5.Height = 78

    AddLabel f5, "Sheet Scale:", 10, 18, 60
    Set cmbScale = f5.Controls.Add("Forms.ComboBox.1", "cmbScale", True)
    cmbScale.Left = 72
    cmbScale.Top = 14
    cmbScale.Width = 68
    cmbScale.Style = 2
    cmbScale.AddItem "1:1"
    cmbScale.AddItem "1:2"
    cmbScale.AddItem "1:5"
    cmbScale.AddItem "2:1"
    cmbScale.AddItem "1:10"
    cmbScale.AddItem "Custom"

    AddLabel f5, "Custom:", 150, 18, 45
    Set txtCustomScale = f5.Controls.Add("Forms.TextBox.1", "txtCustomScale", True)
    txtCustomScale.Left = 198
    txtCustomScale.Top = 14
    txtCustomScale.Width = 60

    Set chkHLR = AddCheckBox(f5, "Use Hidden Lines Removed (HLR)", 10, 42, 170, False, True)
    Set chkPreview = AddCheckBox(f5, "Show Layout Preview Before Generating", 190, 42, 165, True, True)

    Set cmdCreate = Me.Controls.Add("Forms.CommandButton.1", "cmdCreate", True)
    cmdCreate.caption = "Create Drawing"
    cmdCreate.Left = 95
    cmdCreate.Top = 400
    cmdCreate.Width = 90
    cmdCreate.Height = 24

    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    cmdCancel.caption = "Cancel"
    cmdCancel.Left = 205
    cmdCancel.Top = 400
    cmdCancel.Width = 90
    cmdCancel.Height = 24

    Set hCreate = New BtnHandler
    Set hCreate.cmdBtn = cmdCreate
    Set hCreate.FormRef = Me
    hCreate.IsCreate = True

    Set hCancel = New BtnHandler
    Set hCancel.cmdBtn = cmdCancel
    Set hCancel.FormRef = Me
    hCancel.IsCreate = False

    Set hAddSection = New SectionBtnHandler
    Set hAddSection.cmdBtn = cmdAddSection
    Set hAddSection.FormRef = Me
    hAddSection.IsAdd = True

    Set hRemoveSection = New SectionBtnHandler
    Set hRemoveSection.cmdBtn = cmdRemoveSection
    Set hRemoveSection.FormRef = Me
    hRemoveSection.IsAdd = False
End Sub

Private Function AddCheckBox(ByRef parent As Object, ByVal caption As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal defVal As Boolean, ByVal enabledState As Boolean) As Object
    Set AddCheckBox = parent.Controls.Add("Forms.CheckBox.1", "chk" & CStr(parent.Controls.Count + 1), True)
    AddCheckBox.caption = caption
    AddCheckBox.Left = l
    AddCheckBox.Top = t
    AddCheckBox.Width = w
    AddCheckBox.Value = defVal
    AddCheckBox.Enabled = enabledState
End Function

Private Function AddOptionButton(ByRef parent As Object, ByVal caption As String, ByVal grp As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal defVal As Boolean) As Object
    Set AddOptionButton = parent.Controls.Add("Forms.OptionButton.1", "opt" & CStr(parent.Controls.Count + 1), True)
    AddOptionButton.caption = caption
    AddOptionButton.Left = l
    AddOptionButton.Top = t
    AddOptionButton.Width = w
    AddOptionButton.GroupName = grp
    AddOptionButton.Value = defVal
End Function

Private Function AddLabel(ByRef parent As Object, ByVal caption As String, ByVal l As Single, ByVal t As Single, ByVal w As Single) As Object
    Set AddLabel = parent.Controls.Add("Forms.Label.1", "lbl" & CStr(parent.Controls.Count + 1), True)
    AddLabel.caption = caption
    AddLabel.Left = l
    AddLabel.Top = t
    AddLabel.Width = w
End Function

Private Sub LoadPreferences()
    chkTop.Value = ReadBoolSetting("TopView", False)
    chkBottom.Value = ReadBoolSetting("BottomView", False)
    chkRight.Value = ReadBoolSetting("RightView", True)
    chkLeft.Value = ReadBoolSetting("LeftView", False)
    chkBack.Value = ReadBoolSetting("BackView", False)
    chkIso.Value = ReadBoolSetting("IsoView", True)

    cmbDatum.Text = ReadTextSetting("DatumOrigin", "Center")
    chkHoleCallouts.Value = ReadBoolSetting("ImportHoleCallouts", True)

    chkPopulateTitle.Value = ReadBoolSetting("PopulateTitle", True)
    chkBarcode.Value = ReadBoolSetting("Barcode", True)
    chkNotes.Value = ReadBoolSetting("Notes", True)
    chkAutoArrange.Value = ReadBoolSetting("AutoArrange", True)
    txtTotalCost.Text = ReadTextSetting("TotalCostManual", "")

    cmbScale.Text = ReadTextSetting("SheetScale", "1:1")
    txtCustomScale.Text = ReadTextSetting("CustomScale", "")
    chkHLR.Value = ReadBoolSetting("UseHLR", False)
    chkPreview.Value = ReadBoolSetting("ShowPreview", True)

    If Len(Trim$(cmbScale.Text)) = 0 Then cmbScale.Text = "1:1"

    LoadSectionsFromRegistry
End Sub

Private Sub LoadSectionsFromRegistry()
    Dim i As Long
    Dim defaultLabel As String
    Dim sLabel As String

    mSectionCount = CLng(Val(ReadTextSetting("SectionCount", "0")))
    If mSectionCount < 0 Then mSectionCount = 0
    If mSectionCount > 5 Then mSectionCount = 5

    For i = 1 To 5
        mSections(i).Label = vbNullString
        mSections(i).Vertical = False
    Next i

    For i = 1 To mSectionCount
        defaultLabel = Chr$(Asc("J") + i - 1)
        sLabel = Trim$(ReadTextSetting("Section" & i & "Label", defaultLabel))
        If Len(sLabel) = 0 Then sLabel = defaultLabel

        mSections(i).Label = UCase$(Left$(sLabel, 2))
        mSections(i).Vertical = ReadBoolSetting("Section" & i & "Vertical", False)
    Next i
End Sub

Private Function ReadTextSetting(ByVal keyName As String, ByVal defaultValue As String) As String
    ReadTextSetting = GetSetting("VeemapDrawingMacro", "Settings", keyName, defaultValue)
End Function

Private Function ReadBoolSetting(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    Dim s As String
    s = LCase$(Trim$(GetSetting("VeemapDrawingMacro", "Settings", keyName, IIf(defaultValue, "True", "False"))))
    ReadBoolSetting = (s = "true" Or s = "1" Or s = "yes")
End Function

Private Sub SavePreferences()
    SaveSetting "VeemapDrawingMacro", "Settings", "TopView", CStr(chkTop.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "BottomView", CStr(chkBottom.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "RightView", CStr(chkRight.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "LeftView", CStr(chkLeft.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "BackView", CStr(chkBack.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "IsoView", CStr(chkIso.Value)

    SaveSetting "VeemapDrawingMacro", "Settings", "DatumOrigin", cmbDatum.Text
    SaveSetting "VeemapDrawingMacro", "Settings", "ImportHoleCallouts", CStr(chkHoleCallouts.Value)

    SaveSetting "VeemapDrawingMacro", "Settings", "PopulateTitle", CStr(chkPopulateTitle.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "Barcode", CStr(chkBarcode.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "Notes", CStr(chkNotes.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "AutoArrange", CStr(chkAutoArrange.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "TotalCostManual", Trim$(txtTotalCost.Text)

    SaveSetting "VeemapDrawingMacro", "Settings", "SheetScale", cmbScale.Text
    SaveSetting "VeemapDrawingMacro", "Settings", "CustomScale", Trim$(txtCustomScale.Text)
    SaveSetting "VeemapDrawingMacro", "Settings", "UseHLR", CStr(chkHLR.Value)
    SaveSetting "VeemapDrawingMacro", "Settings", "ShowPreview", CStr(chkPreview.Value)

    Dim i As Long
    SaveSetting "VeemapDrawingMacro", "Settings", "SectionCount", CStr(mSectionCount)
    For i = 1 To 5
        SaveSetting "VeemapDrawingMacro", "Settings", "Section" & i & "Label", mSections(i).Label
        SaveSetting "VeemapDrawingMacro", "Settings", "Section" & i & "Vertical", CStr(mSections(i).Vertical)
    Next i
End Sub

Private Sub RefreshSectionList()
    Dim i As Long
    Dim rowText As String

    lstSections.Clear

    For i = 1 To mSectionCount
        rowText = mSections(i).Label & " | " & IIf(mSections(i).Vertical, "Vertical", "Horizontal")
        If i = 1 Then rowText = rowText & " | Primary"
        lstSections.AddItem rowText
    Next i
End Sub

Private Sub RefreshHoleCount()
    On Error Resume Next

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then Exit Sub

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = swApp.ActiveDoc
    If swPart Is Nothing Then Exit Sub
    If swPart.GetType <> 1 Then Exit Sub

    Dim holes As Variant
    holes = Module3_ModelAudit.GetAllHoleLikeFeatures(swPart)
    lblHoleCount.caption = "Hole-like features detected: " & Module3_ModelAudit.CountHoles(holes)
End Sub

Public Sub DoAddSection()
    Dim newLabel As String

    If mSectionCount >= 5 Then
        MsgBox "Maximum of 5 sections supported.", vbExclamation, "Sections"
        Exit Sub
    End If

    UserFormSection.SetupDefaults SuggestedSectionLabel
    UserFormSection.Show

    ' UserFormSection now hides rather than unloading, so these three reads
    ' still see what DoOk assigned. Copy everything out before unloading:
    ' the Unload below is what resets the instance for the next Add Section,
    ' and after it these properties are back to their type defaults.
    Dim wasCancelled As Boolean
    Dim pickedVertical As Boolean
    wasCancelled = UserFormSection.Cancelled
    newLabel = UCase$(Left$(Trim$(UserFormSection.SectionLabel), 2))
    pickedVertical = UserFormSection.SectionVertical
    Unload UserFormSection

    If wasCancelled Then Exit Sub
    If Len(newLabel) = 0 Then Exit Sub

    If SectionLabelExists(newLabel) Then
        MsgBox "Section label '" & newLabel & "' is already in use.", vbExclamation, "Sections"
        Exit Sub
    End If

    mSectionCount = mSectionCount + 1
    mSections(mSectionCount).Label = newLabel
    mSections(mSectionCount).Vertical = pickedVertical

    RefreshSectionList
End Sub

Public Sub DoRemoveSection()
    Dim idx As Long
    Dim i As Long

    idx = lstSections.ListIndex + 1
    If idx <= 0 Or idx > mSectionCount Then Exit Sub

    For i = idx To mSectionCount - 1
        mSections(i) = mSections(i + 1)
    Next i

    mSections(mSectionCount).Label = vbNullString
    mSections(mSectionCount).Vertical = False

    mSectionCount = mSectionCount - 1
    RefreshSectionList
End Sub

Private Function SectionLabelExists(ByVal testLabel As String) As Boolean
    Dim i As Long

    For i = 1 To mSectionCount
        If StrComp(Trim$(mSections(i).Label), Trim$(testLabel), vbTextCompare) = 0 Then
            SectionLabelExists = True
            Exit Function
        End If
    Next i
End Function

Private Function SuggestedSectionLabel() As String
    SuggestedSectionLabel = Chr$(Asc("J") + mSectionCount)
End Function

Private Sub CopySectionsToGlobalState()
    Dim i As Long

    Module1_Main.ResetGlobalSections
    Module1_Main.GlobalSectionCount = mSectionCount

    For i = 1 To mSectionCount
        Module1_Main.GlobalSections(i).Label = Trim$(mSections(i).Label)
        Module1_Main.GlobalSections(i).Vertical = mSections(i).Vertical
    Next i

    Module1_Main.SyncSectionCompatibilityFields
End Sub

Public Sub DoCreate()
    SavePreferences

    Module1_Main.GlobalConfig.CreateFront = True
    Module1_Main.GlobalConfig.CreateTop = chkTop.Value
    Module1_Main.GlobalConfig.CreateBottom = chkBottom.Value
    Module1_Main.GlobalConfig.CreateRight = chkRight.Value
    Module1_Main.GlobalConfig.CreateLeft = chkLeft.Value
    Module1_Main.GlobalConfig.CreateBack = chkBack.Value
    Module1_Main.GlobalConfig.CreateIso = chkIso.Value

    ' Both producers, always. The reference drawing needs native import for
    ' R36 / the section diameters / the hole callouts AND the ordinate engine
    ' for the two position chains; the old option group made that impossible.
    Module1_Main.GlobalConfig.UseModelDimensions = True
    Module1_Main.GlobalConfig.UseOrdinateDims = True
    Module1_Main.GlobalConfig.RunHybridStrategy = True
    Module1_Main.GlobalConfig.ImportHoleCallouts = chkHoleCallouts.Value
    Module1_Main.GlobalConfig.DatumOrigin = cmbDatum.Text

    Module1_Main.GlobalConfig.PopulateTitle = chkPopulateTitle.Value
    Module1_Main.GlobalConfig.InsertBarcode = chkBarcode.Value
    Module1_Main.GlobalConfig.InsertNotes = chkNotes.Value
    Module1_Main.GlobalConfig.AutoArrange = chkAutoArrange.Value

    Module1_Main.GlobalConfig.SheetScale = ResolveScaleValue(cmbScale.Text, txtCustomScale.Text)
    Module1_Main.GlobalConfig.CustomScaleText = Trim$(txtCustomScale.Text)
    Module1_Main.GlobalConfig.UseHLR = chkHLR.Value
    Module1_Main.GlobalConfig.ShowLayoutPreview = chkPreview.Value
    Module1_Main.GlobalConfig.TotalCostManual = Trim$(txtTotalCost.Text)
    Module1_Main.GlobalConfig.GenerateQAReport = True
    Module1_Main.GlobalConfig.Cancelled = False

    CopySectionsToGlobalState

    Unload Me
End Sub

Private Function ResolveScaleValue(ByVal scaleText As String, ByVal customText As String) As Double
    Select Case Trim$(scaleText)
        Case "1:1": ResolveScaleValue = 1#
        Case "1:2": ResolveScaleValue = 0.5
        Case "1:5": ResolveScaleValue = 0.2
        Case "2:1": ResolveScaleValue = 2#
        Case "1:10": ResolveScaleValue = 0.1
        Case "Custom": ResolveScaleValue = ParseCustomScale(customText)
        Case Else: ResolveScaleValue = 1#
    End Select
End Function

Private Function ParseCustomScale(ByVal customText As String) As Double
    Dim s As String
    Dim a() As String

    s = Replace(Trim$(customText), " ", "")

    If InStr(1, s, ":") > 0 Then
        a = Split(s, ":")
        If UBound(a) = 1 Then
            If Val(a(1)) <> 0 Then ParseCustomScale = Val(a(0)) / Val(a(1))
        End If
    ElseIf Val(s) > 0 Then
        ParseCustomScale = Val(s)
    End If

    If ParseCustomScale <= 0 Then ParseCustomScale = 1#
End Function

Public Sub DoCancel()
    Module1_Main.GlobalConfig.Cancelled = True
    Unload Me
End Sub

Private Sub UserForm_QueryUnload(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Module1_Main.GlobalConfig.Cancelled = True
End Sub

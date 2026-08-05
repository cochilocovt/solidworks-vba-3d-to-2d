Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{CAEA2A24-62AC-4DD4-8B65-EB771354E048}{7E608A52-3539-44B8-A593-42277EABED84}"
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
    Me.Caption = "Auto Drawing Settings"
    Me.width = 390
    Me.height = 470

    BuildUI
    LoadPreferences
    ApplyFixtureDefaults
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
    f1.Caption = "Fixture View Plan (locked)"
    f1.Left = 8
    f1.Top = 10
    f1.width = 160
    f1.height = 150

    Set chkFront = AddCheckBox(f1, "Front View (Primary)", 10, 15, 135, True, False)
    Set chkTop = AddCheckBox(f1, "Top View", 10, 35, 120, False, True)
    Set chkBottom = AddCheckBox(f1, "Bottom View", 10, 55, 120, False, True)
    Set chkRight = AddCheckBox(f1, "Right View", 10, 75, 120, True, True)
    Set chkLeft = AddCheckBox(f1, "Left View", 10, 95, 120, False, True)
    Set chkBack = AddCheckBox(f1, "Back View", 10, 115, 120, False, True)
    Set chkIso = AddCheckBox(f1, "Isometric View", 10, 135, 120, True, True)

    Set f2 = Me.Controls.Add("Forms.Frame.1", "frmSections", True)
    f2.Caption = "Fixture Primary Section (locked)"
    f2.Left = 178
    f2.Top = 10
    f2.width = 192
    f2.height = 150

    Set lstSections = f2.Controls.Add("Forms.ListBox.1", "lstSections", True)
    lstSections.Left = 10
    lstSections.Top = 18
    lstSections.width = 170
    lstSections.height = 82

    Set cmdAddSection = f2.Controls.Add("Forms.CommandButton.1", "cmdAddSection", True)
    cmdAddSection.Caption = "Add Section..."
    cmdAddSection.Left = 10
    cmdAddSection.Top = 108
    cmdAddSection.width = 78

    Set cmdRemoveSection = f2.Controls.Add("Forms.CommandButton.1", "cmdRemoveSection", True)
    cmdRemoveSection.Caption = "Remove"
    cmdRemoveSection.Left = 102
    cmdRemoveSection.Top = 108
    cmdRemoveSection.width = 78

    Set f3 = Me.Controls.Add("Forms.Frame.1", "frmDims", True)
    f3.Caption = "Dimension Settings (Model Dims + Ordinate, Automatic)"
    f3.Left = 8
    f3.Top = 170
    f3.width = 160
    f3.height = 98

    AddLabel f3, "Ordinate Datum Origin:", 10, 16, 140

    Set cmbDatum = f3.Controls.Add("Forms.ComboBox.1", "cmbDatum", True)
    cmbDatum.Left = 10
    cmbDatum.Top = 32
    cmbDatum.width = 120
    cmbDatum.Style = 2
    cmbDatum.AddItem "Bottom-Left"
    cmbDatum.AddItem "Center"
    cmbDatum.AddItem "Top-Left"
    cmbDatum.ListIndex = 0

    Set chkHoleCallouts = AddCheckBox( _
        f3, _
        "Insert Hole Wizard Callouts", _
        10, _
        56, _
        140, _
        True, _
        True)

    Set lblHoleCount = AddLabel( _
        f3, _
        "Semantic hole-producing features: 0", _
        10, _
        76, _
        145)

    Set f4 = Me.Controls.Add("Forms.Frame.1", "frmOther", True)
    f4.Caption = "Other Settings"
    f4.Left = 178
    f4.Top = 170
    f4.width = 192
    f4.height = 122

    Set chkPopulateTitle = AddCheckBox(f4, "Populate Title Block (required)", 10, 16, 170, True, False)
    Set chkBarcode = AddCheckBox(f4, "Insert Part Identification (required)", 10, 36, 180, True, False)
    Set chkNotes = AddCheckBox(f4, "Insert General Notes (required)", 10, 56, 170, True, False)
    Set chkAutoArrange = AddCheckBox(f4, "Auto-Arrange Dims", 10, 76, 140, True, True)

    AddLabel f4, "Total Cost (manual):", 10, 98, 90

    Set txtTotalCost = f4.Controls.Add("Forms.TextBox.1", "txtTotalCost", True)
    txtTotalCost.Left = 105
    txtTotalCost.Top = 94
    txtTotalCost.width = 70

    Set f5 = Me.Controls.Add("Forms.Frame.1", "frmSheet", True)
    f5.Caption = "Sheet Settings"
    f5.Left = 8
    f5.Top = 302
    f5.width = 362
    f5.height = 78

    AddLabel f5, "Sheet Scale:", 10, 18, 60

    Set cmbScale = f5.Controls.Add("Forms.ComboBox.1", "cmbScale", True)
    cmbScale.Left = 72
    cmbScale.Top = 14
    cmbScale.width = 68
    cmbScale.Style = 2
    cmbScale.AddItem "1:1"
    cmbScale.AddItem "1:2"
    cmbScale.AddItem "1:5"
    cmbScale.AddItem "1:6"
    cmbScale.AddItem "1.5:1"
    cmbScale.AddItem "2:1"
    cmbScale.AddItem "1:10"
    cmbScale.AddItem "Custom"

    AddLabel f5, "Custom:", 150, 18, 45

    Set txtCustomScale = f5.Controls.Add("Forms.TextBox.1", "txtCustomScale", True)
    txtCustomScale.Left = 198
    txtCustomScale.Top = 14
    txtCustomScale.width = 60

    Set chkHLR = AddCheckBox( _
        f5, _
        "Use Hidden Lines Removed (HLR)", _
        10, _
        42, _
        170, _
        False, _
        True)

    Set chkPreview = AddCheckBox( _
        f5, _
        "Layout Preview (not available)", _
        190, _
        42, _
        165, _
        False, _
        False)

    Set cmdCreate = Me.Controls.Add("Forms.CommandButton.1", "cmdCreate", True)
    cmdCreate.Caption = "Create Drawing"
    cmdCreate.Left = 95
    cmdCreate.Top = 398
    cmdCreate.width = 90
    cmdCreate.height = 24

    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    cmdCancel.Caption = "Cancel"
    cmdCancel.Left = 205
    cmdCancel.Top = 398
    cmdCancel.width = 90
    cmdCancel.height = 24

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

Private Function AddCheckBox( _
    ByRef parent As Object, _
    ByVal Caption As String, _
    ByVal l As Single, _
    ByVal t As Single, _
    ByVal w As Single, _
    ByVal defVal As Boolean, _
    ByVal enabledState As Boolean) As Object

    Set AddCheckBox = parent.Controls.Add( _
        "Forms.CheckBox.1", _
        "chk" & CStr(parent.Controls.Count + 1), _
        True)

    AddCheckBox.Caption = Caption
    AddCheckBox.Left = l
    AddCheckBox.Top = t
    AddCheckBox.width = w
    AddCheckBox.Value = defVal
    AddCheckBox.Enabled = enabledState
End Function

Private Function AddLabel( _
    ByRef parent As Object, _
    ByVal Caption As String, _
    ByVal l As Single, _
    ByVal t As Single, _
    ByVal w As Single) As Object

    Set AddLabel = parent.Controls.Add( _
        "Forms.Label.1", _
        "lbl" & CStr(parent.Controls.Count + 1), _
        True)

    AddLabel.Caption = Caption
    AddLabel.Left = l
    AddLabel.Top = t
    AddLabel.width = w
End Function

Private Sub SetComboChoice( _
    ByRef combo As Object, _
    ByVal requestedValue As String, _
    ByVal fallbackValue As String)

    Dim itemIndex As Long
    For itemIndex = 0 To combo.ListCount - 1
        If StrComp( _
            CStr(combo.List(itemIndex)), requestedValue, vbTextCompare) = 0 Then

            combo.ListIndex = itemIndex
            Exit Sub
        End If
    Next itemIndex

    For itemIndex = 0 To combo.ListCount - 1
        If StrComp( _
            CStr(combo.List(itemIndex)), fallbackValue, vbTextCompare) = 0 Then

            combo.ListIndex = itemIndex
            Exit Sub
        End If
    Next itemIndex

    If combo.ListCount > 0 Then combo.ListIndex = 0
End Sub

Private Sub LoadPreferences()
    chkTop.Value = ReadBoolSetting("TopView", False)
    chkBottom.Value = ReadBoolSetting("BottomView", False)
    chkRight.Value = ReadBoolSetting("RightView", True)
    chkLeft.Value = ReadBoolSetting("LeftView", False)
    chkBack.Value = ReadBoolSetting("BackView", False)
    chkIso.Value = ReadBoolSetting("IsoView", True)

    SetComboChoice cmbDatum, _
        ReadTextSetting("DatumOrigin", "Bottom-Left"), "Bottom-Left"
    chkHoleCallouts.Value = ReadBoolSetting("ImportHoleCallouts", True)

    chkPopulateTitle.Value = ReadBoolSetting("PopulateTitle", True)
    chkBarcode.Value = ReadBoolSetting("Barcode", True)
    chkNotes.Value = ReadBoolSetting("Notes", True)
    chkAutoArrange.Value = ReadBoolSetting("AutoArrange", True)
    txtTotalCost.Text = ReadTextSetting("TotalCostManual", "")

    SetComboChoice cmbScale, _
        ReadTextSetting("SheetScale", "1:1"), "1:1"
    txtCustomScale.Text = ReadTextSetting("CustomScale", "")
    chkHLR.Value = ReadBoolSetting("UseHLR", False)
    chkPreview.Value = False

    LoadSectionsFromRegistry
End Sub

Private Sub ApplyFixtureDefaults()
    On Error GoTo SafeExit

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then Exit Sub

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = swApp.ActiveDoc
    If swPart Is Nothing Then Exit Sub

    Dim fixtureKey As String
    fixtureKey = Module1_Main.GetFixtureKey(swPart.GetPathName)
    If Len(fixtureKey) = 0 Then Exit Sub

    chkTop.Value = False
    chkBottom.Value = False
    chkRight.Value = False
    chkLeft.Value = False
    chkBack.Value = False
    chkIso.Value = True
    chkTop.Enabled = False
    chkBottom.Enabled = False
    chkRight.Enabled = False
    chkLeft.Enabled = False
    chkBack.Enabled = False
    chkIso.Enabled = False
    lstSections.Enabled = False
    cmdAddSection.Enabled = False
    cmdRemoveSection.Enabled = False
    chkPreview.Value = False

    mSectionCount = 0
    mSections(1).Label = vbNullString
    mSections(1).Vertical = False

    Select Case fixtureKey
        Case "P-0251-14A-001"
            chkLeft.Value = True
            SetComboChoice cmbScale, "1:1", "1:1"
            mSectionCount = 1
            mSections(1).Label = "J"

        Case "P-0252-01-001"
            chkRight.Value = True
            SetComboChoice cmbScale, "1:6", "1:1"

        Case "P-0252-01-013"
            chkBottom.Value = True
            chkLeft.Value = True
            chkRight.Value = True
            SetComboChoice cmbScale, "1.5:1", "1:1"
            mSectionCount = 1
            mSections(1).Label = "B"
    End Select

SafeExit:
End Sub

Private Sub LoadSectionsFromRegistry()
    Dim i As Long
    Dim defaultLabel As String
    Dim sLabel As String

    mSectionCount = CLng(Val(ReadTextSetting("SectionCount", "0")))

    If mSectionCount < 0 Then mSectionCount = 0
    If mSectionCount > 1 Then mSectionCount = 1

    For i = 1 To 5
        mSections(i).Label = vbNullString
        mSections(i).Vertical = False
    Next i

    For i = 1 To mSectionCount
        defaultLabel = Chr$(Asc("J") + i - 1)
        sLabel = Trim$(ReadTextSetting("Section" & i & "Label", defaultLabel))

        If Len(sLabel) = 0 Then sLabel = defaultLabel

        mSections(i).Label = UCase$(Left$(sLabel, 2))
        mSections(i).Vertical = ReadBoolSetting( _
            "Section" & i & "Vertical", _
            False)
    Next i
End Sub

Private Function ReadTextSetting( _
    ByVal keyName As String, _
    ByVal defaultValue As String) As String

    ReadTextSetting = GetSetting( _
        "VeemapDrawingMacro", _
        "Settings", _
        keyName, _
        defaultValue)
End Function

Private Function ReadBoolSetting( _
    ByVal keyName As String, _
    ByVal defaultValue As Boolean) As Boolean

    Dim s As String

    s = LCase$(Trim$(GetSetting( _
        "VeemapDrawingMacro", _
        "Settings", _
        keyName, _
        IIf(defaultValue, "True", "False"))))

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
        SaveSetting "VeemapDrawingMacro", "Settings", _
                    "Section" & i & "Label", _
                    mSections(i).Label

        SaveSetting "VeemapDrawingMacro", "Settings", _
                    "Section" & i & "Vertical", _
                    CStr(mSections(i).Vertical)
    Next i
End Sub

Private Sub RefreshSectionList()
    Dim i As Long
    Dim rowText As String

    lstSections.Clear

    For i = 1 To mSectionCount
        rowText = mSections(i).Label & " | " & _
                  IIf(mSections(i).Vertical, "Vertical", "Horizontal")

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

    Dim holes As Collection
    Set holes = Module3_ModelAudit.GetAllHoleLikeFeatures(swPart)

    lblHoleCount.Caption = _
        "Semantic hole-producing features: " & _
        Module3_ModelAudit.CountHoles(holes)
End Sub

Public Sub DoAddSection()
    Dim newLabel As String

    If mSectionCount >= 1 Then
        MsgBox "The current production data model supports one primary section.", _
               vbExclamation, _
               "Sections"
        Exit Sub
    End If

    Dim sectionForm As UserFormSection
    Set sectionForm = New UserFormSection
    sectionForm.SetupDefaults SuggestedSectionLabel()
    sectionForm.Show

    If sectionForm.Cancelled Then
        Unload sectionForm
        Exit Sub
    End If

    newLabel = UCase$(Left$(Trim$(sectionForm.SectionLabel), 2))

    If Len(newLabel) = 0 Then
        Unload sectionForm
        Exit Sub
    End If

    If SectionLabelExists(newLabel) Then
        MsgBox "Section label '" & newLabel & "' is already in use.", _
               vbExclamation, _
               "Sections"
        Unload sectionForm
        Exit Sub
    End If

    mSectionCount = mSectionCount + 1
    mSections(mSectionCount).Label = newLabel
    mSections(mSectionCount).Vertical = sectionForm.SectionVertical

    Unload sectionForm

    RefreshSectionList
End Sub

Public Sub DoRemoveSection()
    Dim idx As Long
    Dim i As Long

    idx = lstSections.ListIndex + 1

    If idx <= 0 Or idx > mSectionCount Then Exit Sub

    For i = idx To mSectionCount - 1
        mSections(i).Label = mSections(i + 1).Label
        mSections(i).Vertical = mSections(i + 1).Vertical
    Next i

    mSections(mSectionCount).Label = vbNullString
    mSections(mSectionCount).Vertical = False

    mSectionCount = mSectionCount - 1

    RefreshSectionList
End Sub

Private Function SectionLabelExists(ByVal testLabel As String) As Boolean
    Dim i As Long

    For i = 1 To mSectionCount
        If StrComp( _
            Trim$(mSections(i).Label), _
            Trim$(testLabel), _
            vbTextCompare) = 0 Then

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
    Dim resolvedScale As Double
    Dim scaleError As String

    If Not TryResolveScaleValue( _
        cmbScale.Text, txtCustomScale.Text, resolvedScale, scaleError) Then

        MsgBox scaleError, vbExclamation, "Sheet Scale"
        Exit Sub
    End If

    SavePreferences

    Module1_Main.GlobalConfig.CreateFront = True
    Module1_Main.GlobalConfig.CreateTop = chkTop.Value
    Module1_Main.GlobalConfig.CreateBottom = chkBottom.Value
    Module1_Main.GlobalConfig.CreateRight = chkRight.Value
    Module1_Main.GlobalConfig.CreateLeft = chkLeft.Value
    Module1_Main.GlobalConfig.CreateBack = chkBack.Value
    Module1_Main.GlobalConfig.CreateIso = chkIso.Value

    Module1_Main.GlobalConfig.UseModelDimensions = True
    Module1_Main.GlobalConfig.UseOrdinateDims = True
    Module1_Main.GlobalConfig.RunHybridStrategy = True

    Module1_Main.GlobalConfig.ImportHoleCallouts = chkHoleCallouts.Value
    Module1_Main.GlobalConfig.DatumOrigin = cmbDatum.Text

    Module1_Main.GlobalConfig.PopulateTitle = True
    Module1_Main.GlobalConfig.InsertBarcode = True
    Module1_Main.GlobalConfig.InsertNotes = True
    Module1_Main.GlobalConfig.AutoArrange = chkAutoArrange.Value

    Module1_Main.GlobalConfig.SheetScale = resolvedScale

    Module1_Main.GlobalConfig.CustomScaleText = Trim$(txtCustomScale.Text)
    Module1_Main.GlobalConfig.UseHLR = chkHLR.Value
    Module1_Main.GlobalConfig.ShowLayoutPreview = False
    Module1_Main.GlobalConfig.TotalCostManual = Trim$(txtTotalCost.Text)

    Module1_Main.GlobalConfig.GenerateQAReport = True
    Module1_Main.GlobalConfig.Cancelled = False

    CopySectionsToGlobalState

    Unload Me
End Sub

Private Function TryResolveScaleValue( _
    ByVal scaleText As String, _
    ByVal customText As String, _
    ByRef scaleValue As Double, _
    ByRef errorText As String) As Boolean

    Select Case Trim$(scaleText)
        Case "1:1"
            scaleValue = 1#

        Case "1:2"
            scaleValue = 0.5

        Case "1:5"
            scaleValue = 0.2

        Case "1:6"
            scaleValue = 1# / 6#

        Case "1.5:1"
            scaleValue = 1.5

        Case "2:1"
            scaleValue = 2#

        Case "1:10"
            scaleValue = 0.1

        Case "Custom"
            If Not TryParseCustomScale(customText, scaleValue) Then
                errorText = _
                    "Enter a positive scale as N:D or a positive decimal value."
                Exit Function
            End If

        Case Else
            errorText = "The selected sheet scale is not supported."
            Exit Function
    End Select

    TryResolveScaleValue = True
End Function

Private Function TryParseCustomScale( _
    ByVal customText As String, _
    ByRef scaleValue As Double) As Boolean

    Dim s As String
    Dim a() As String

    s = Replace(Trim$(customText), " ", "")

    If InStr(1, s, ":") > 0 Then
        a = Split(s, ":")

        If UBound(a) = 1 Then
            If IsNumeric(a(0)) And IsNumeric(a(1)) Then
                If CDbl(a(0)) > 0# And CDbl(a(1)) > 0# Then
                    scaleValue = CDbl(a(0)) / CDbl(a(1))
                    TryParseCustomScale = True
                End If
            End If
        End If

    ElseIf IsNumeric(s) Then
        If CDbl(s) > 0# Then
            scaleValue = CDbl(s)
            TryParseCustomScale = True
        End If
    End If
End Function

Public Sub DoCancel()
    Module1_Main.GlobalConfig.Cancelled = True
    Unload Me
End Sub

Private Sub UserForm_QueryUnload( _
    Cancel As Integer, _
    CloseMode As Integer)

    If CloseMode = 0 Then
        Module1_Main.GlobalConfig.Cancelled = True
    End If
End Sub


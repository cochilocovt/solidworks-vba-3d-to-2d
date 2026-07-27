Attribute VB_Name = "UserFormSection"
Attribute VB_Base = "0{EAD39048-A2A0-4E7B-8AF1-684127D52F80}{152E07B4-9077-4762-A230-8E140C9793A1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private txtLabel As Object
Private optVertical As Object
Private optHorizontal As Object
Private cmdOK As Object
Private cmdCancel As Object

Private hOk As SectionDlgBtnHandler
Private hCancel As SectionDlgBtnHandler

Public SectionLabel As String
Public SectionVertical As Boolean
Public Cancelled As Boolean

Private mDefaultLabel As String

Public Sub SetupDefaults(ByVal defaultLabel As String)
    mDefaultLabel = UCase$(Left$(Trim$(defaultLabel), 2))
    Cancelled = True

    On Error Resume Next
    If Not txtLabel Is Nothing Then
        If Len(mDefaultLabel) > 0 Then
            txtLabel.Text = mDefaultLabel
        Else
            txtLabel.Text = "J"
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    Cancelled = True
    SectionLabel = vbNullString
    SectionVertical = False

    Me.Caption = "Section Settings"
    Me.width = 230
    Me.height = 150

    Dim lbl1 As Object
    Set lbl1 = Me.Controls.Add("Forms.Label.1", "lbl1", True)
    lbl1.Caption = "Section Label:"
    lbl1.Left = 12
    lbl1.Top = 14
    lbl1.width = 70

    Set txtLabel = Me.Controls.Add("Forms.TextBox.1", "txtLabel", True)
    txtLabel.Left = 90
    txtLabel.Top = 10
    txtLabel.width = 45
    If Len(mDefaultLabel) > 0 Then
        txtLabel.Text = mDefaultLabel
    Else
        txtLabel.Text = "J"
    End If

    Set optVertical = Me.Controls.Add("Forms.OptionButton.1", "optVertical", True)
    optVertical.Caption = "Vertical Cut Line"
    optVertical.Left = 12
    optVertical.Top = 42
    optVertical.width = 120
    optVertical.GroupName = "SectionDir"

    Set optHorizontal = Me.Controls.Add("Forms.OptionButton.1", "optHorizontal", True)
    optHorizontal.Caption = "Horizontal Cut Line"
    optHorizontal.Left = 12
    optHorizontal.Top = 62
    optHorizontal.width = 130
    optHorizontal.GroupName = "SectionDir"
    optHorizontal.Value = True

    Set cmdOK = Me.Controls.Add("Forms.CommandButton.1", "cmdOK", True)
    cmdOK.Caption = "OK"
    cmdOK.Left = 38
    cmdOK.Top = 92
    cmdOK.width = 60

    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    cmdCancel.Caption = "Cancel"
    cmdCancel.Left = 112
    cmdCancel.Top = 92
    cmdCancel.width = 60

    Set hOk = New SectionDlgBtnHandler
    Set hOk.cmdBtn = cmdOK
    Set hOk.FormRef = Me
    hOk.IsOk = True

    Set hCancel = New SectionDlgBtnHandler
    Set hCancel.cmdBtn = cmdCancel
    Set hCancel.FormRef = Me
    hCancel.IsOk = False
End Sub

Public Sub DoOk()
    Dim cleanLabel As String

    cleanLabel = UCase$(Left$(Trim$(txtLabel.Text), 2))
    If Len(cleanLabel) = 0 Then
        MsgBox "Please enter a section label.", vbExclamation, "Section"
        Exit Sub
    End If

    SectionLabel = cleanLabel
    SectionVertical = optVertical.Value
    Cancelled = False
    Unload Me
End Sub

Public Sub DoCancel()
    Cancelled = True
    Unload Me
End Sub

Private Sub UserForm_QueryUnload(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancelled = True
End Sub

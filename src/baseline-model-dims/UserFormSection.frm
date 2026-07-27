Attribute VB_Name = "UserFormSection"
Attribute VB_Base = "0{E0A2079B-A743-4DD2-BDB5-A11777DD65E2}{EED94989-72DA-406F-8A5A-5E20DA97FCA7}"
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

    Me.caption = "Section Settings"
    Me.Width = 230
    Me.Height = 150

    Dim lbl1 As Object
    Set lbl1 = Me.Controls.Add("Forms.Label.1", "lbl1", True)
    lbl1.caption = "Section Label:"
    lbl1.Left = 12
    lbl1.Top = 14
    lbl1.Width = 70

    Set txtLabel = Me.Controls.Add("Forms.TextBox.1", "txtLabel", True)
    txtLabel.Left = 90
    txtLabel.Top = 10
    txtLabel.Width = 45
    If Len(mDefaultLabel) > 0 Then
        txtLabel.Text = mDefaultLabel
    Else
        txtLabel.Text = "J"
    End If

    Set optVertical = Me.Controls.Add("Forms.OptionButton.1", "optVertical", True)
    optVertical.caption = "Vertical Cut Line"
    optVertical.Left = 12
    optVertical.Top = 42
    optVertical.Width = 120
    optVertical.GroupName = "SectionDir"

    Set optHorizontal = Me.Controls.Add("Forms.OptionButton.1", "optHorizontal", True)
    optHorizontal.caption = "Horizontal Cut Line"
    optHorizontal.Left = 12
    optHorizontal.Top = 62
    optHorizontal.Width = 130
    optHorizontal.GroupName = "SectionDir"
    optHorizontal.Value = True

    Set cmdOK = Me.Controls.Add("Forms.CommandButton.1", "cmdOK", True)
    cmdOK.caption = "OK"
    cmdOK.Left = 38
    cmdOK.Top = 92
    cmdOK.Width = 60

    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    cmdCancel.caption = "Cancel"
    cmdCancel.Left = 112
    cmdCancel.Top = 92
    cmdCancel.Width = 60

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

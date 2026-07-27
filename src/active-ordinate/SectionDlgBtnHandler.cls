Attribute VB_Name = "SectionDlgBtnHandler"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Public WithEvents cmdBtn As MSForms.CommandButton
Attribute cmdBtn.VB_VarHelpID = -1
Public FormRef As Object
Public IsOk As Boolean

Private Sub cmdBtn_Click()
    If IsOk Then
        FormRef.DoOk
    Else
        FormRef.DoCancel
    End If
End Sub


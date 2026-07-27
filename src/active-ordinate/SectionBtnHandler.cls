Attribute VB_Name = "SectionBtnHandler"
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
Public IsAdd As Boolean

Private Sub cmdBtn_Click()
    If IsAdd Then
        FormRef.DoAddSection
    Else
        FormRef.DoRemoveSection
    End If
End Sub


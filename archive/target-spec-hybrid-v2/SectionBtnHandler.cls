Option Explicit

Public WithEvents cmdBtn As MSForms.CommandButton
Public FormRef As Object
Public IsAdd As Boolean

Private Sub cmdBtn_Click()
    If IsAdd Then
        FormRef.DoAddSection
    Else
        FormRef.DoRemoveSection
    End If
End Sub


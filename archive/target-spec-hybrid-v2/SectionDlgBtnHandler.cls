Option Explicit

Public WithEvents cmdBtn As MSForms.CommandButton
Public FormRef As Object
Public IsOk As Boolean

Private Sub cmdBtn_Click()
    If IsOk Then
        FormRef.DoOk
    Else
        FormRef.DoCancel
    End If
End Sub


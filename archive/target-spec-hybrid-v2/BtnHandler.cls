Option Explicit

Public WithEvents cmdBtn As MSForms.CommandButton
Public FormRef As Object
Public IsCreate As Boolean

Private Sub cmdBtn_Click()
    If IsCreate Then
        FormRef.DoCreate
    Else
        FormRef.DoCancel
    End If
End Sub


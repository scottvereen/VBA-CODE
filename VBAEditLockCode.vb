Option Compare Database
Option Explicit

Private Sub cmdEdit_Click()

If Me.cmdEdit.Caption = "Edit Locked" Then
    Me.cmdEdit.Caption = "Edit Unlocked"
    Me.cmdEdit.BackColor = vbYellow
    Me.txtMaterialTransaction.Locked = False
    Me.OutlineBox.BorderColor = vbYellow
    
Else
    Me.cmdEdit.Caption = "Edit Locked"
    Me.cmdEdit.BackColor = 13487553
    Me.txtMaterialTransaction.Locked = True
    Me.OutlineBox.BorderColor = vbBlack
End If

    
End Sub

Private Sub Form_AfterUpdate()
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "QryAppendMaterialChangeLog"
    DoCmd.SetWarnings True
    
End Sub

Private Sub Form_Load()
DoCmd.GoToRecord , , acNewRec
Me.txtMaterialTransaction.Locked = True

End Sub


Option Compare database
Friend Sub cmdEdit_Click()
'code in here

End Sub 

Private Sub cmdEdit_Click()

	If Me.cmdEdit.Caption = "Edit Locked" Then
    				Me.cmdEdit.Caption = "Edit Unlocked"
    				'Me.cmdEdit.BackColor = vbYellow
    				'Me.txtMaterialTransaction.Locked = False
				Me.AsofPSR.Locked = False
				Me.FundingSent = False
				Me.RDSIIncurred = False
    				'Me.OutlineBox.BorderColor = vbYellow
Else

    				'Me.cmdEdit.Caption = "Edit Locked"
    				'Me.cmdEdit.BackColor = 13487553
    				'Me.txtMaterialTransaction.Locked = True
				Me.AsofPSR.Locked = True
     			' Me.OutlineBox.BorderColor = vbBlack
				Me.FundingSent =True
				Me.RDSIIncurred= True
			
End If

End Sub

Private Sub Form_AfterUpdate()
       		 	DoCmd.SetWarnings False
        			DoCmd.OpenQuery "QFL"
        			DoCmd.SetWarnings True
End Sub

Private Sub Form_Load()
    				'Me.txtMaterialTransaction.Locked = True
				Me.AsofPSR.locked=True
				Me.RDSIIncurred=True
				Me.FundingSent=True
    				DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub MaterialTransaction_AfterUpdate()

    DoCmd.SetWarnings False
    DoCmd.OpenQuery "QFL"
    DoCmd.SetWarnings True

End Sub



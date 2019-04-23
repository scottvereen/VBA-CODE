'Combo Box Search drop down
Private Sub cbo_TIAfterUpdate()
Dim myTI As String
myTI="Select * from tblFunding where ([TI_ID]=" & Me.cboTI &")"
Me.SubFrmFundingSearch.Form.RecordSource=myTI
Me.SubFrmFundingSearch.Form.Requery
End Sub

Dim myContract As String
myContract="Select * from tblFunding where ([ContractNumber_ID]=" & Me.cboContract &")"
Me.SubFrmFundingSearch.Form.RecordSource=myContract
Me.SubFrmFundingSearch.Form.Requery


Dim myProject As String
myProject="Select * from tblFunding where ([Project/ProgramName_ID]="&Me.cboProject &")"
Me.SubFrmFundingSearch.Form.RecordSource=myProject
Me.SubFrmFundingSearch.Form.Requery

'List box search

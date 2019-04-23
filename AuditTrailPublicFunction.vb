Option Compare Database


Public Function AuditChanges(RecordID As String, UserAction As String)
On Error Goto AuditErr
'Dim cnn As ADODB.Connection
'Dim rst As ADODB.Recordset
Dim DB As Database
Dim rst As Recordset
Dim clt As Control
Dim UserLogin As String

Set DB = CurrentDb
Set rst = DB.OpenRecordset("SELECT * FROM tbl_audittrail",adOpenDynamic)
UserLogin = enriron("Username")
Select Case UserAction
	Case "new"
		With rst
				.AddNew
				![DateTime]= Now()
				!UserName = UserLogin
				!FormName = Screen.ActiveForm.Name
				!Action = UserAction
				!RecordID = Screen.ActiveForm.Controls(RecordID).Value
				.Update
				
		End With
				
	Case "Delete"
		With rst
				.AddNew
				![DateTime]= Now()
				!UserName = UserLogin
				!FormName = Screen.ActiveForm.Name
				!Action = UserAction
				!RecordID = Screen.ActiveForm.Controls(RecordID).Value
				.Update
				
		End With
	
	Case "Edit"
		For Each clt In Screen.ActiveForm.Controls
			If(clt.ControlType = acTextBox_
			Or clt.ControlType = acComboBox) Then
			If Nz(clt.Value)<> Nz(clt.OldValue) Then
				With rst
					.AddNew
						![DateTime] = Now()
						!UserName = UserLogin
						!FormName = Screen.ActiveForm.Name
						!Action = UserAction
						!RecordID = Screen.ActiveForm.Controls(RecordID).Value
						!fieldname = ctl.ControlSource
						!OldValue = clt.OldValue
						!newvalue = ctl.Value
						
						
			End With
		End If
	End If
	
	Next clt
		
End Select
				
rst.Close
DB.Close
Set rst = Nothing
Set DB = Nothing

auditerr:
	MsgBox Err.Number & "   :   " & Err.Description, vbCritical, "Error"
		Exit Function
		
	End Function
	




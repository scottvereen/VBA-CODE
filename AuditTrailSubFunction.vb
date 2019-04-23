Sub AuditChanges (IDField As String, UserAction As String)
On Error Goto AuditChanges_Err
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim ctl As Control
Dim datTimeCheck As Date
Dim strUserID As String

Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset

rst.Open “Select * From tblAuditTrail”, cnn, adOpenDynamic, adLockOptimistic
datTimeCheck = Now()
strUserID = Environ(“USERNAME”)
Select Case useraction
Case “EDIT”
For Each ctl In Screen.ActiveForm.Controls
If ctl.Tag = “Audit” Then
If Nz(ctl.Value) <> Nz(ctl.OldValue) Then
With rst
.AddNew
![FormName] = Screen.ActiveForm.Name
![RecordID] = Screen.ActiveForm.Controls(IDField).Value
![FieldName] = ctl.ControlSource
![OldValue] = ctl.OldValue
![NewValue] = ctl.Value
![UserID] = strUserID
![DateTime] = datTimeCheck
.Update
End With
End If
End If
Next ctl
Case Else
With rst
.AddNew
![DateTime] = datTimeCheck
![UserID] = strUserID
![FormName] = Screen.ActiveForm.Name
![Action] = useraction
![RecordID] = Screen.ActiveForm.Controls(IDField).Value
.Update
End With
End Select
AuditChanges_Exit:
On Error Resume Next
rst.Close
cnn.Close
Set rst = Nothing
Set cnn = Nothing
Exit Sub
AuditChanges_Err:
MsgBox Err.Description, vbCritical, “Error!”
Resume AuditChanges_Exit
End Sub
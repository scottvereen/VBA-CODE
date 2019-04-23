Sub AuditFunding(IDField As String)
    On Error GoTo AuditFunding_Err
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim ctl As Control
    Dim datTimeCheck As Date
    Dim strUserID As String
    Set cnn = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM tblFundingTrail", cnn, adOpenDynamic, adLockOptimistic
    datTimeCheck = Now()
    strUserID = Environ("USERNAME")
    For Each ctl In Screen.ActiveForm.Controls
        If ctl.Tag = "Audit" Then
            If Nz(ctl.Value) <> Nz(ctl.OldValue) Then
                With rst
                    .AddNew
                    ![DateTime] = datTimeCheck
                    ![UserName] = strUserID
                    ![FormName] = Screen.ActiveForm.Name
                    ![RecordID] = Screen.ActiveForm.Controls(IDField).Value
                    ![FieldName] = ctl.ControlSource
                    ![OldValue] = ctl.OldValue
                    ![NewValue] = ctl.Value
                    .Update
                End With
            End If
        End If
    Next ctl
AuditFunding_Exit:
    On Error Resume Next
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    Exit Sub
AuditFunding_Err:
    MsgBox Err.Description, vbCritical, "ERROR!"
    Resume AuditFunding_Exit
End Sub

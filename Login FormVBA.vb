Private Sub btnOk_Click()
Dim User As String
Dim UserLevel As Integer
Dim TempPass As String
Dim ID As Integer
Dim workerName As String
Dim TempLoginID As String


If IsNull(Me.txtUserName) Then
	MsgBox "Please Enter UserName", vbInformation, "Username Required"
	Me.txtUserName.SetFocus
Else If IsNull(Me.txtPassword) Then
	MsgBox "Please Enter UserName", vbInformation, "Username Required"
	Me.txtPassword.SetFocus
Else
	If (IsNull(DLookup("userID", "tbluser", "UserName= ' " & Me.txtUserName.Value & " ' And UserPassword = ' " & Me.txtPassword.Value &  " ' " ) ) ) Then
		MsgBox "Invalid UserName or Password!"
	Else
				TempLogID = Me.txtUserName.Value
				workerName = DLook("[Employeename]", "tbluser", "[UserName] = ' " & Me.txtUserName.Value & " ' ")
				UserLevel = DLookup ("[UserType]", "tbluser", "[UserName]= ' " & Me.txtUserName.Value & " ' ")
				TempPass = DLookup("[password]", "tbluser", "[UserName]= ' " & Me.txtUserName.Value & " ' ")
				ID = DLookup("[Userid]" , "tbluser", "[UserName] = ' " & Me.txtUserName.Value & " ' ")
				DoCmd.Close
				If (TempPass = "password") Then
					MsgBox "Please change password", vbInformation, "New password requeired"
					DoCmd.OpenForm "frmworkerinfo" , , , "[userid] = " & ID
				Else
					'Open different form according to user level
						If UserLevel = 1 Then 'for admin
							DoCmd.OpenForm "Navigation Form"
							Forms![Navigation Form] ! [txtLogin] = TempLoginID
							Forms![Navigation Form] ! [yxyUser] - workerName
						DoCmd.BrowserTo acBrowserToForm, "frmFirstPage" , "Navigaton Form.NavigationSubForm" , , , acFormEdit
						Else 
						DoCmd.OpenForm "Navigation Form"
							Forms![Navigation Form] ! [txtLogin] = TempLoginID
							Forms![Navigation Form]!NavigationButton13.Enableed = False
							Forms![Navigation Form] ! [yxyUser] - workerName
						DoCmd.BrowserTo acBrowserToForm, "frmFirstPage" , "Navigaton Form.NavigationSubForm" , , , acFormEdit
						End If
						
						End If
					End If
			End If
			
End Sub

Private Sub cmd_login_Click()

  Dim db As DAO.Database
  Dim rst As DAO.Recordset
  Dim strSQL As String
 
  If Trim(Me.txt_username.Value & vbNullString) = vbNullString Then
    MsgBox prompt:="Username should not be left blank.", buttons:=vbInformation, title:="Username Required"
    Me.txt_username.SetFocus
    Exit Sub
  End If
 
  If Trim(Me.txt_password.Value & vbNullString) = vbNullString Then
    MsgBox prompt:="Password should not be left blank.", buttons:=vbInformation, title:="Password Required"
    Me.txt_password.SetFocus
    Exit Sub
  End If
 
  'query to check if login details are correct
  strSQL = "SELECT FirstName FROM tbl_login WHERE Username = """ & Me.txt_username.Value & """ AND Password = """ & Me.txt_password.Value & """"
 
  Set db = CurrentDb
  Set rst = db.OpenRecordset(strSQL)
  If rst.EOF Then
    MsgBox prompt:="Incorrect username/password. Try again.", buttons:=vbCritical, title:="Login Error"
    Me.txt_username.SetFocus
  Else
    MsgBox prompt:="Hello, " & rst.Fields(0).Value & ".", buttons:=vbOKOnly, title:="Login Successful"
    DoCmd.Close acForm, "frm_login", acSaveYes
  End If
 
 Set db = Nothing
 Set rst = Nothing

End Sub
Private Sub cmd_cancel_Click()
   DoCmd.Quit acQuitSaveAll
End Sub


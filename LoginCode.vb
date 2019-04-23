' Login button code variables are as follows
' txtbox calls large box at bottom of form
' login_id field name in TableLogin

Private Sub btnLogin_Click()

	Dim blRet As Boolean
	
	Login_ID.SetFocus
	txtbox.RowSource = "Select*from TableLogin where Login_ID = ' " & Login_ID.Value & " ' and password = ' " & Password.value & " ' "
	If txtbox.ListCount = 0 Then
		msgbox  " Wrong Login and or Password", vbCritical
	Else
		Modulel.variable_username = txtbox.Column(2, 1)
		Modulel.variable_role = txtbox.Column(4, 1)
		
		
			DoCmd.Close
			
			If Modulel.variable_role = "admin" Then
				stDocName = "menu_admin"
				DoCmd.OpenForm stDocName, , , stLinkCriteria
			End If
			
			If Modulel.variable_role = "read" Then
				stDocName = "menu_read"
				DoCmd.OpenForm stDocName, , , stLinkCriteria
			End If
			
	End If
	


End Sub

'Code for Modulel "Module"

Option Compare Database

Public variable_username
Public variable_role
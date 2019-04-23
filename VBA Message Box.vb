

'Message Box VBA code
'Simple Yes/No

 If MsgBox("Do you want to close the database?", vbQuestion + vbYesNo, "Exit Database") = vbYes Then
        DoCmd.RunCommand acCmdExit
    Else
        MsgBox "Function aborted"
    End If
	
'Message full Yes no cancel	
	Private Sub bSave_Click()
On Error Goto Err_bSave_Click

    Me.tbHidden.SetFocus

    If IsNull(tbFirstName) Or IsNull(tbLastName) Then
        Beep
        MsgBox "All required fields must be completed before you can save a record.", vbCritical, "Invalid Save"
        Exit Sub
    End If

    Beep
    Select Case MsgBox("Do you want to save your changes to the current record?" & vbCrLf & vbLf & "  Yes:         Saves Changes" & vbCrLf & "  No:          Does NOT Save Changes" & vbCrLf & "  Cancel:    Reset (Undo) Changes" & vbCrLf, vbYesNoCancel + vbQuestion, "Save Current Record?")
        Case vbYes: 'Save the changes
            DoCmd.RunCommand acCmdSaveRecord

        Case vbNo: 'Do not save or undo
            'Do nothing

        Case vbCancel: 'Undo the changes
            DoCmd.RunCommand acCmdUndo

        Case Else: 'Default case to trap any errors
            'Do nothing
    End Select

Exit_bSave_Click:
    Exit Sub

Err_bSave_Click:
    If Err = 2046 Then 'The command or action Undo is not available now
        Exit Sub
    Else
        MsgBox Err.Number & " - " & Err.Description
        Resume Exit_bSave_Click
    End If
    
End Sub

	
	
	

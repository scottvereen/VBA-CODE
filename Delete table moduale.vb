Option Compare Database
Option Explicit

Public Function delete_table(tblNTIRAInstallPlan_xlsx As String) As Boolean
'returns true if the table is deleted or if the table does not exists
On Error GoTo errHandler
    If table_exists(tblNTIRAInstallPlan_xlsx) Then
        DoCmd.DeleteObject acTable, tblNTIRAInstallPlan_xlsx
    End If
    delete_table = True
exitSuccess:
    Exit Function
errHandler:
    Call display_error
    Resume exitSuccess
End Function



CurrentDb.TableDefs("' sheet1 $'_ImportErrors").Name= "Error"


' Sub display_error()
' 'display error code number and description in the immediate window
    ' Debug.Print Err.Number, Err.Description
' End Sub


Public Sub test()
'test functions
    
    Const MY_TABLE = "NTIRAInstallPlan_xlsx"
    DoCmd.SetWarnings False
    Debug.Print
    Debug.Print delete_table(MY_TABLE)
    DoCmd.SetWarnings True
    
End Sub


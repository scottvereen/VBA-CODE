Option Compare Database
Option Explicit
Global LSQL As String


Public Function delete_table_records(tblNTIRA_xlsx As String) As Boolean
    'This Function deletes all records in a table by its name main table name is NTRIABase
    
    On Error GoTo errHandler
        If table_exists(tblNTIRA) Then
            LSQL = "DELETE * FROM(" & tblNTIRA_xlsx & ")"
                DoCmd.RunSQL LSQL
        End If
        
        delete_table_records = True
        
exitSuccess:
        Exit Function
errHandler:
        Debug.Print Err.Number, Err.Description
        Resume exitSuccess
        

End Function

Public Function table_exists(tblNTIRA_xlsx As String) As Boolean
    'Returns a true if table exist
    Dim tdf As TableDef
    For Each tdf In CurrentDb.TableDefs
               
        If StrComp(tblNTIRA_xlsx, tdf.NTRI_xlsx) = 0 Then
            table_exists = True
            Exit For
        End If
    Next tdf
End Function


Public Function delete_table(tblNTIRA_xlsx As String) As Boolean

    'Returns a true if the table is deleted or if table does not exist
    On Error GoTo errHandler
        If table_exists(tblNTIRA_xlsx) Then
            DoCmd.DeleteObject acTable, tblNTIRA_xlsx
        
        End If
        delete_table = True
exitSuccess:
        Exit Function
errHandler:
    Debug.Print Err.Number, Err.Description
    Resume exitSuccess

End Function
 Sub testdelete()
 
    Const MY_TABLE = "NTIRA_xlsx"
    DoCmd.SetWarnings False
    Debug.Print
    Debug.Print delete_table(MY_TABLE)
    DoCmd.SetWarnings True
 
 End Sub

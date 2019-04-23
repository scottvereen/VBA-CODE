'Delete Tables Buttons

Private Sub Command9_Click()
Dim tbl As AccessObject

For Each tbl In CurrentData.AllTables
If tbl.Name Like "*_ImportError*" Then
DoCmd.DeleteObject acTable, tbl.Name
End If
Next tbl
End Sub
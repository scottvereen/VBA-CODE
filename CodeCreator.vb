Expr1: [PO_Number] & [TechnicaInstructionNumber]

Like "*" & [Forms]![FormSearchMain]![search].[Text] & "*"

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Not Me.NewRecord Then Call AuditChanges("CustomerID")
End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)
    Call AuditChanges("EmployeeID")
End Sub
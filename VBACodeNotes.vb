
'Clean form load
Private Sub Form_Load()
'This is it
	DoCmd.GoToRecord , , acNewRec
End Sub

'?Other
<>Nz([Form]![ID],0)

DoCmd.GoToRecord , , acNewRec
=Nz([Parts_Description]![Description],"Untitled")


'Duplicate Check
=IIf(DCount("*","[Employees Extended]","[ID]<>" & Nz([ID],0) & " And [Employee Name] = '" & Replace(Nz([Employee Name]),"'","''") & "'")>0,"Possible Duplicate","")

"Requested";"Submitted";"Ordered";"Received";"Transfer"



'Adding a button that will upload Excel spread sheet

Private Sub btnUpload_Click()
'Set the file path to a string
Dim filepath As String

'Choose correct file path

filepath = "C:\User\Scott\Upload\UploadSheet.xlsx"
'If Statement Only if you with to try his code
If FileExist(filepath) Then
'This is One sheet only
DoCmd.TransferSpreasheet acImport, , "ImprtFromExcel" , filepath, True

'Several Sheet select the one you want
DoCmd.TransferSpreadsheet acImport,,"ImportFromExcel",filepath, True, "Sheet1"﻿

'If Else statment Optional
Else 
	MsgBox "File Not Found. Please Check File Name or File Location"
'After Update Code
Private Sub PurchaseRequisitionNumber_AfterUpdate()
End If
End Sub

'Using the DIR function may not be the best way to determine
'if a file exists, especially if you are using VB to compare two
'dictionarys, and take actions when the files do not match

Function FileExist(sTestFile As String) As Boolean


'this function does not use DIR since it is possible that you might have
'been in the middle of running DIR against another directory in
'an attemplt to match one directory against another
'It does not handle wildcard characters
Dim 1Size As Long
On Error Resume Next
'Preset lenght to -1 because files can be zero bytes in lenght
1Size = -1
'Get the lenght of the file
1Size = FileLen(sTestFile)
If 1Size > -1 Then
FileExist = True
Else
FileExist = False
End If
End Function


'As i looked up i found it changed to this
'This example checks to see whether or not the file Check.txt exists and supplies the information in a message box.
If My.Computer.FileSystem.FileExists("c:\Check.txt") Then
    MsgBox("File found.")
Else
    MsgBox("File not found.")
End If


'Turning information into a string
	Dim NewPurchaseRequisitionNumber As String
	Dim stLinkCriteria A String
	
	
	NewPurchaseRequisitionNumber = Me.PurchaseRequisitionNumber.Value
	stLinkCriteria = "[PurchaseRequisitionNumber]=" & "'" & NewPurchaseRequisitionNumber & "'"
	If Me.PurchaseRequisitionNumber = dlookup("[PurchaseRequisitionNumber], tbl3Document",stLinkCriteria) Then
		MsgBox "This Purchase Request Number, "& NewPurchaseRequisitionNumber &" Is already Being Used" _
			& vbCr & vbCr & "Please Check",vbInformation, "Duplicate Information"
			
		Me.Undo
	
	End If
	



End Sub

=If(MAX(COUNTIF(A2:A623,A2:A623))>1,"Duplicate","No Duplocate")



=IIf(DCount("*","[[tbl4Parts]![PartNumber]]","[PartName]<>" & Nz([ID],0) & " And [PartName] = '" & Replace(Nz([PartName]),"'","''") & "'")>0,"Possible Duplicate","")

=IIf(DCount("*","[Employees Extended]","[ID]<>" & Nz([ID],0) & " And [Employee Name] = '" & Replace(Nz([Employee Name]),"'","''") & "'")>0,"Possible Duplicate","")

<>Nz([Form]![ID],0)

=If(MAX(COUNTIF(A2:A423,A2:A423))>1,"Duplicate","No Duplocate")



TechnicalIID: Nz([tblTI].[TechnicalIID],"")


Like "*" & [Forms]![FormCDRLsSearch]![txtTechnicalIID] & "*"


"Received";"Transferred";"Ordered";"Requested
";"Not Ordered
";"Shipped"


'Note search query allowing all variables
TechnicalIID: Nz([tblTI].[TechnicalIID],"")
Like "*" & [Forms]![FormCDRLsSearch]![txtTechnicalIID] & "*"




=IIf(DCount("*","[[tbl3Employee]![ID]]","[ID]<>" & Nz([ID],0) & " And [LastName] = '" & Replace(Nz([LastName]),"'","''") & "'")>0,"Possible Duplicate","")

'Importing Information From One Sheet In Excel to another




=OFFSET(Sheet1!$A$1,(ROW()-1)*7,0)
=OFFSET(Sheet1!A$1,(ROW()-1)*7,0)


=IFERROR(INDEX(Task_Definitions!$A:$BA,$E10,MATCH(H$9,Task_Definitions!$1:$1,0)),"")
=IFERROR(INDEX(Report!$A:$BA,$A1,MATCH(H$9,Task_Definitions!$1:$1,0)),"")



=MATCH($A$2,'41133007.01.F8340.000.000.CUST'!$A:$A,0)
=MATCH($A$2,INDIRECT("'41133007.01.F8340.000.000.CUST'!$A"&A3+1&":$A10000"),0)+$A3
=MATCH($H$2,INDIRECT("'41133007.01.F8340.000.000.CUST'!$A"&H3+1&":$A10000"),0)+$H3


=MATCH($A$2,'Report!$A:$A,0)




=INDEX('41133007.01.F8340.000.000.CUST'!$A:$Z,$A3,B1)
=INDEX('41133007.01.F8340.000.000.CUST'!$A:$Z,$A4,2)

 =Mid(CELL("filename",O2),FIND("]",CELL("filename",O2))+1,10)
 
 =Excel.CurrentWorkbook()







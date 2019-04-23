File As: IIf(IsNull([Last Name]),IIf(IsNull([First Name]),[Company],[First Name]),IIf(IsNull([First Name]),[Last Name],[Last Name] & ", " & [First Name]))
File As: IIf(IsNull([Last Name]),IIf(IsNull([First Name]),[Company],[First Name]),IIf(IsNull([First Name]),[Last Name],[Last Name] & ", " & [First Name]))

Supplier Name: IIf(IsNull([Last Name]),IIf(IsNull([First Name]),[Company],[First Name]),IIf(IsNull([First Name]),[Last Name],[First Name] & " " & [Last Name]))
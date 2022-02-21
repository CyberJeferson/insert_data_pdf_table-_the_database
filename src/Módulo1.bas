Attribute VB_Name = "Módulo1"
Sub insert_db()
Dim cont As Integer

cont = 1
While (cont < 1919)
If (Cells.Item(cont, 1) = "") Then
Else
Cells.Item(cont, 3) = "insert into u_names(name,total) values('" & Cells.Item(cont, 1) & "', " & Cells.Item(cont, 2) & ");"
End If
cont = cont + 1
Wend
End Sub

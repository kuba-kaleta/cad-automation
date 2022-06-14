Function table_test()

Dim varData(3) As Variant
varData(0) = "Claudia Bendel"
varData(1) = "4242 Maple Blvd"
varData(2) = 38
varData(3) = Format("06-09-1952", "General Date")

table_test = varData

End Function


Sub main()

Dim varData1 As Variant
varData1 = table_test()
MsgBox Join(varData1, vbCrLf)

End Sub

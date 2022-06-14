Function remove_pl(str)
If True Then ' wlaczenie polskich znakow
str = Replace(str, "ą", "a")
str = Replace(str, "ż", "z")
str = Replace(str, "ź", "z")
str = Replace(str, "ę", "e")
str = Replace(str, "ć", "c")
str = Replace(str, "ł", "l")
str = Replace(str, "ń", "n")
str = Replace(str, "ó", "o")
str = Replace(str, "-", " ")
End If
remove_pl = str
End Function

Sub create_sql()

Dim Counter
Counter = 10

Dim str
str1 = "$kod_arr = array("
str2 = "$nazwa_arr = array("
str3 = "$opis_arr = array("
str4 = "$stanowisko_arr = array("
str5 = "$stawka_arr = array("
str6 = "$czas_arr = array("
str7 = "$skl_sumy = array("

While Cells(Counter + 1, 1).Value <> ""

str1 = str1 & "'" & remove_pl(Cells(Counter, 1).Value) & "'" & ", " ' wlaczenie polskich znakow w funkcji remove_pl
str2 = str2 & "'" & remove_pl(Cells(Counter, 2).Value) & "'" & ", "
str3 = str3 & "'" & remove_pl(Cells(Counter, 3).Value) & "'" & ", "
str4 = str4 & "'" & remove_pl(Cells(Counter, 4).Value) & "'" & ", "
str5 = str5 & Cells(Counter, 5).Text & ", "
str6 = str6 & Cells(Counter, 6).Text & ", "
str7 = str7 & "'" & remove_pl(Cells(Counter, 7).Value) & "'" & ", "

Counter = Counter + 1
Wend

str1 = str1 & "'" & remove_pl(Cells(Counter, 1).Value) & "'" & ")" & ";"
str2 = str2 & "'" & remove_pl(Cells(Counter, 2).Value) & "'" & ")" & ";"
str3 = str3 & "'" & remove_pl(Cells(Counter, 3).Value) & "'" & ")" & ";"
str4 = str4 & "'" & remove_pl(Cells(Counter, 4).Value) & "'" & ")" & ";"
str5 = str5 & Cells(Counter, 5).Text & ")" & ";"
str6 = str6 & Cells(Counter, 6).Text & ")" & ";"
str7 = str7 & "'" & remove_pl(Cells(Counter, 7).Value) & "'" & ")" & ";"

Cells(1, 10).Value = str1
Cells(2, 10).Value = str2
Cells(3, 10).Value = str3
Cells(4, 10).Value = str4
Cells(5, 10).Value = str5
Cells(6, 10).Value = str6
Cells(7, 10).Value = str7

End Sub



Sub prepare_w2_table()

    ' Cells(27, 6).Value = "ok"

    Dim Licznik
    Licznik = 0
    Dim Counter

    For I = 1 To 200
        
        Cells(I + 22, 10).Value = ""
        Cells(I + 22, 11).Value = ""
        Cells(I + 22, 14).Value = ""
    
    Next I

    For I = 5 To 25
        Counter = 0

        While Cells(I, 2 * (Counter + 1) + 1).Value <> "" And Cells(I, 2 * (Counter + 1) + 1).Value <> "0"
            Cells(I + 23 + Counter + Licznik, 10).Value = Cells(I, 2 * (Counter + 1))
            Cells(I + 23 + Counter + Licznik, 11).Value = Cells(I, 1)
            Cells(I + 23 + Counter + Licznik, 14).Value = Cells(I, 2 * (Counter + 1) + 1) / 60
            Counter = Counter + 1
        Wend

        Licznik = Licznik + Counter - 1

    Next I

End Sub

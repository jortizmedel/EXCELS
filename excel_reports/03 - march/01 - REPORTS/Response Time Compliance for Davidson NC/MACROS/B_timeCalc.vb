Sub B_timeCalc()
    Dim y As Long
    Dim valor As Variant
    
    Sheets("DATA").Select
    y = Sheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    
    For i = 2 To y
        valor = Cells(i, 8).Value - Cells(i, 7).Value
        Cells(i, 9).Value = Format(valor, "h:mm:ss")
    Next i
    
    For i = 2 To y
        valor = Cells(i, 10).Value - Cells(i, 8).Value
        Cells(i, 11).Value = Format(valor, "h:mm:ss")
    Next i
    
    For i = 2 To y
        valor = Cells(i, 10).Value - Cells(i, 7).Value
        Cells(i, 12).Value = Format(valor, "h:mm:ss")
    Next i
End Sub
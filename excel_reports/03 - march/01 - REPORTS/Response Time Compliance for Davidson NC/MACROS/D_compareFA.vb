Sub D_compareFA()
    Dim term As Variant
    Dim check As Boolean
    Dim y As Long
    Dim valor As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    Sheets("DATA").Select
    Range("M2").Value = "1st"
    Range("N2").Value = "2nd"
    Range("O2").Value = "3rd"
    Range("P2").Value = "4th"
    
    y = Sheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    
    '1st
    For i = 3 To y
        cont = WorksheetFunction.CountIf(Range("C:C"), Cells(i, 3).Value)
        fechaIn = Format(Cells(i, 12).Value, "h:mm:ss")
        fechafin = "0:06:20"
        check = False
        If fechaIn <= fechafin Then check = True
        
        '1 apparatus
        If check = True And cont = 1 Then Cells(i, 13).Value = "100%"
        If check = False And cont = 1 Then Cells(i, 13).Value = "0%"
        '2+ apparatus
        If cont > 1 Then
            min = MinInc(Range(Cells(2, 12), Cells(y, 12)), Range(Cells(2, 4), Cells(y, 4)), Cells(i, 4).Value, Range(Cells(2, 3), Cells(y, 3)), Cells(i, 3).Value)
            If Cells(i, 12).Value = min Then
                fechaIn = Format(min, "h:mm:ss")
                If fechaIn <= fechafin Then
                    Cells(i, 13).Value = "100%"
                Else
                    Cells(i, 13).Value = "0%"
                End If
            Else
                Cells(i, 13).Value = ""
            End If
        End If
        For Z = 3 To y
            count = WorksheetFunction.CountIfs(Range(Cells(3, 12), Cells(Z, 12)), Cells(Z, 12).Value, Range(Cells(3, 4), Cells(Z, 4)), Cells(Z, 4).Value, Range(Cells(3, 3), Cells(Z, 3)), Cells(Z, 3).Value)
            If count <> 1 Then Cells(Z, 13).Value = ""
        Next
    Next

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
End Sub
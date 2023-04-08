Sub E_compareSA()
    Dim term As Variant
    Dim check As Boolean
    Dim y As Long
    Dim valor As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    y = Sheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    
    '2nd
    For q = 3 To y
        fechaIn = Format(Cells(q, 12).Value, "h:mm:ss")
        fechafin = "0:08:20"
        check = False
        If fechaIn <= fechafin Then check = True
        If Cells(q, 13).Value <> "" Then Cells(q, 14).Value = ""
        If Cells(q, 13).Value = "" Then
            min1 = MinInc(Range(Cells(3, 12), Cells(y, 12)), Range(Cells(3, 4), Cells(y, 4)), Cells(q, 4).Value, Range(Cells(3, 3), Cells(y, 3)), Cells(q, 3).Value)
            If Cells(q, 12).Value = min1 Then
                fechaIn = Format(min1, "h:mm:ss")
                If fechaIn <= fechafin Then Cells(q, 14).Value = "100%"
                If fechaIn > fechafin Then Cells(q, 14).Value = "0%"
            Else
                min2 = MinIncSecond(Range(Cells(3, 12), Cells(y, 12)), Range(Cells(3, 4), Cells(y, 4)), Cells(q, 4).Value, Range(Cells(3, 3), Cells(y, 3)), Cells(q, 3).Value, min1)
                If Cells(q, 12).Value = min2 Then
                    fechaIn = Format(min2, "h:mm:ss")
                    If fechaIn <= fechafin Then Cells(q, 14).Value = "100%"
                    If fechaIn > fechafin Then Cells(q, 14).Value = "0%"
                Else
                    Cells(q, 14).Value = ""
                End If
            End If
        End If
    Next

    For a = y To 3 Step -1
        If Cells(a, 14).Value <> "" And Cells(a - 1, 14).Value <> "" And Cells(a - 2, 13) <> "" Then Cells(a, 14).Value = ""
        If Cells(a, 14).Value <> "" And Cells(a - 1, 14).Value <> "" And Cells(a - 1, 13).Value = "" Then Cells(a, 14).Value = ""
        If Cells(a, 14).Value <> "" And Cells(a - 1, 14).Value <> "" Then Cells(a, 14).Value = ""
        If Cells(a, 14).Value <> "" And Cells(a, 13).Value <> "" Then Cells(a, 14).Value = ""
    Next a

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
End Sub
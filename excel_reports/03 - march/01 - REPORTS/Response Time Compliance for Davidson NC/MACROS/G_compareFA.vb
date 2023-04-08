Sub G_compareFA()
    Dim term As Variant
    Dim check As Boolean
    Dim y As Long
    Dim valor As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    y = Sheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    
    '4th
    For q = 3 To y
        fechaIn = Format(Cells(q, 12).Value, "h:mm:ss")
        fechafin = "0:08:20"
        check = False
        If fechaIn <= fechafin Then check = True
        If Cells(q, 15).Value <> "" Then Cells(q, 16).Value = ""
        If Cells(q, 13).Value = "" And Cells(q, 14).Value = "" And Cells(q, 15) = "" Then
            If fechaIn <= fechafin Then Cells(q, 16).Value = "100%"
            If fechaIn > fechafin Then Cells(q, 16).Value = "0%"
        End If
    Next

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
End Sub
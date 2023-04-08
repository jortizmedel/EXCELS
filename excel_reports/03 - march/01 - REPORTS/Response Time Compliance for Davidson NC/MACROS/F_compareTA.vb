Sub F_compareTA()
    Dim term As Variant
    Dim check As Boolean
    Dim y As Long
    Dim valor As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    y = Sheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    
    '3rd
    For q = 3 To y
        fechaIn = Format(Cells(q, 12).Value, "h:mm:ss")
        fechafin = "0:08:20"
        check = False
        If fechaIn <= fechafin Then check = True
        If Cells(q, 14).Value <> "" Then Cells(q, 15).Value = ""
        If Cells(q, 13).Value = "" And Cells(q, 14) = "" Then
            If fechaIn <= fechafin Then Cells(q, 15).Value = "100%"
            If fechaIn > fechafin Then Cells(q, 15).Value = "0%"
        End If
    Next

    For a = y To 3 Step -1
        If Cells(a, 15).Value <> "" And Cells(a - 1, 15).Value <> "" And Cells(a - 2, 13) <> "" Then Cells(a, 15).Value = ""
        If Cells(a, 15).Value <> "" And Cells(a - 1, 15).Value <> "" And Cells(a - 1, 13).Value = "" Then Cells(a, 15).Value = ""
        If Cells(a, 15).Value <> "" And Cells(a - 1, 15).Value <> "" Then Cells(a, 15).Value = ""
        If Cells(a, 15).Value <> "" And Cells(a, 13).Value <> "" Then Cells(a, 15).Value = ""
    Next a

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.CutCopyMode = False
    
End Sub
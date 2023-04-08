Sub A_copyData()
    Dim x As Long
    Dim Address As String
    Dim city As String
    Dim state As String
    Dim zipCode As String
    
    Sheets("REPORT").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    Sheets("DATA").Select
    Range("A1").Select
    
    x = Sheets("Response Time Compliance - Raw ").Cells(Rows.count, "A").End(xlUp).Row

    Sheets("Response Time Compliance - Raw ").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Range("A1").Value = "Incident Type"
     
    Sheets("Response Time Compliance - Raw ").Select
    Columns("H:H").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("B:B").Select
    ActiveSheet.Paste
    Columns("B:B").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("B1").Value = "Incident Date"
    
    For i = 2 To x + 1
        Sheets("Response Time Compliance - Raw ").Select
        Address = Cells(i, 3).Value
        city = Cells(i, 4).Value
        state = Cells(i, 5).Value
        zipCode = Cells(i, 6).Value
        Sheets("DATA").Select
        Cells(i, 4).Value = Address & " " & city & ", " & state & " " & zipCode
    Next
    
    Sheets("Response Time Compliance - Raw ").Select
    Columns("B:B").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("C:C").Select
    ActiveSheet.Paste
    Range("C1").Value = "Incident #"

    Sheets("Response Time Compliance - Raw ").Select
    Columns("G:G").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("E:E").Select
    ActiveSheet.Paste
    Range("E1").Value = "Apparatus"
    
    Sheets("DATA").Select
    Columns("B:B").Select
    Selection.Copy
    Columns("F:F").Select
    ActiveSheet.Paste
    Range("F1").Value = "Alarm Time"
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "h:mm:ss"
    
    Sheets("Response Time Compliance - Raw ").Select
    Columns("I:I").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("G:G").Select
    ActiveSheet.Paste
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "h:mm:ss"
    Range("G1").Value = "Dispatch Time"
    
    Sheets("Response Time Compliance - Raw ").Select
    Columns("J:J").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("H:H").Select
    ActiveSheet.Paste
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "h:mm:ss"
    Range("H1").Value = "Enroute Time"

    Sheets("Response Time Compliance - Raw ").Select
    Columns("K:K").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("J:J").Select
    ActiveSheet.Paste
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "h:mm:ss"
    Range("J1").Value = "Arrival Time"
End Sub
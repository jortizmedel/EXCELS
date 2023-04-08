Sub K_DeleteSheets()
    Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.name <> "REPORT" And xWs.name <> "INSTRUCTIONS" And xWs.name <> "Response Time Compliance - Raw " And xWs.name <> "DATA" Then
            xWs.Delete
        End If
    Next
    
    Sheets("DATA").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Value = "Incident Type"
    Range("B1").Value = "Incident Date"
    Range("C1").Value = "Incident #"
    Range("D1").Value = "Address"
    Range("E1").Value = "Apparatus"
    Range("F1").Value = "Alarm Time"
    Range("G1").Value = "Dispatch Time"
    Range("H1").Value = "Enroute Time"
    Range("I1").Value = "Turnout Time"
    Range("J1").Value = "Arrival Time"
    Range("K1").Value = "Travel Time"
    Range("L1").Value = "Total Time"
    Range("M1").Value = "Arriving Apparatus"
    Range("N1").Value = "Arriving Apparatus"
    Range("O1").Value = "Arriving Apparatus"
    Range("P1").Value = "Arriving Apparatus"
    Range("A2").Select
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
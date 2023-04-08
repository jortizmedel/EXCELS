Function MinInc(ToTime As Range, ListAdd As Range, Address As Variant, ListInc As Range, Inc As Variant)
    Dim c As Range
    Dim min As Variant
    Dim check As Boolean
    check = False
    counter = 1
    For Each c In ToTime
        If ListAdd(counter) = Address And ListInc(counter) = Inc Then
            If check = False Then
                min = c.Value
                check = True
            Else
                If c.Value < min Then min = c.Value
            End If
        End If
        counter = counter + 1
    Next
    MinInc = min
End Function

Function MinIncSecond(ToTime As Range, ListAdd As Range, Address As Variant, ListInc As Range, Inc As Variant, Exc As Variant)
    Dim c As Range
    Dim min As Variant
    Dim check As Boolean
    check = False
    counter = 1
    For Each c In ToTime
        If ListAdd(counter) = Address And ListInc(counter) = Inc And c.Value <> Exc Then
            If check = False Then
                min = c.Value
                check = True
            Else
                If c.Value < min Then min = c.Value
            End If
        End If
        counter = counter + 1
    Next
    MinIncSecond = min
End Function

Function contaMerge(celda) As Integer
    dire = celda.Address
    Set ma = Range(dire).MergeArea
    contaMerge = ma.count
End Function

Function moveInfo(Hoja As Worksheet) As Boolean
    Dim celdaObj As Range
    Dim lastRow As Long
    Set celdaObj = Range("A1")
    Hoja.Activate
    moveInfo = True
    Hoja.Cells(celdaObj.Row, celdaObj.Column).CurrentRegion.Select
    Selection.Copy
    lastRow = Sheets("REPORT").Cells(Rows.count, 1).End(xlUp).Row
    Sheets("REPORT").Activate
    Sheets("REPORT").Cells(lastRow + 1, 1).Select
    Sheets("REPORT").Paste
    Application.CutCopyMode = False
End Function

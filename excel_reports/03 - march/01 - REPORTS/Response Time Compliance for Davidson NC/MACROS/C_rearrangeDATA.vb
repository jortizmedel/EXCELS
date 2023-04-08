Sub C_rearrangeDATA()
    Sheets("DATA").Select
    
    x = Sheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("DATA").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DATA").Sort.SortFields.Add2 Key:=Range(Cells(2, 1), Cells(x, 1)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("DATA").Sort.SortFields.Add2 Key:=Range(Cells(2, 2), Cells(x, 2)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("DATA").Sort.SortFields.Add2 Key:=Range(Cells(2, 3), Cells(x, 3)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("DATA").Sort.SortFields.Add2 Key:=Range(Cells(2, 12), Cells(x, 12)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("DATA").Sort.SortFields.Add2 Key:=Range(Cells(2, 5), Cells(x, 5)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA").Sort
        '.SetRange Range("A1:L269")
		.SetRange Range(Cells(1, 1), Cells(x, 12))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
End Sub
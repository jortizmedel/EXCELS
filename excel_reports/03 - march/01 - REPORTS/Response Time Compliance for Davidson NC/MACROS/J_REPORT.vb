Sub J_REPORT()
    Dim Hoja As Worksheet
    Dim fc As Long
    Dim x As Boolean
    Dim cant100 As Long
    Dim cant0 As Long
    Dim i As Long
    Dim suma100 As Long
    Dim suma0 As Long
    Dim porcent As Long
    Dim porcent1 As Long
    
    cant100 = 0
    cant0 = 0
    For Each Hoja In ThisWorkbook.Sheets
        If Hoja.name <> "DATA" And Hoja.name <> "REPORT" And Hoja.name <> "Response Time Compliance - Raw " And Hoja.name <> "STORE" Then
            x = moveInfo(Hoja)
        End If
    Next Hoja

    Columns("B:B").ColumnWidth = 10.57
    Columns("A:A").ColumnWidth = 9.71
    Columns("C:C").ColumnWidth = 16.14
    Columns("D:D").ColumnWidth = 8.14
    Columns("E:G").ColumnWidth = 7.71
    Columns("I:I").ColumnWidth = 7.71
    Columns("K:K").ColumnWidth = 6.57
    Columns("H:H").ColumnWidth = 6.57
    Columns("J:J").ColumnWidth = 6.57
    Columns("L:O").ColumnWidth = 8.4
    Cells.Select
    Cells.EntireRow.AutoFit
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    fc = Worksheets("REPORT").Cells(Rows.count, "L").End(xlUp).Row
    
    Cells(fc + 2, 11).Select
    Cells(fc + 2, 11).Value = "Overall Response Compliance by Apparatus"
    Cells(fc + 3, 11).Value = "Overall Response Compliance"
    Cells(fc + 4, 11).Value = "Overall Response Percentage by Apparatus"
    Cells(fc + 5, 11).Value = "Overall Response Percentage"
   
    For i = 1 To fc
        If Cells(i, 12).Text = "100%" And Cells(i, 11).MergeCells = False Then cant100 = cant100 + 1
        If Cells(i, 12).Text = "0%" And Cells(i, 11).MergeCells = False Then cant0 = cant0 + 1
    Next
    
    If cant100 <> 0 Then porcent = (cant100 / (cant100 + cant0)) * 100
    Cells(fc + 2, 12).Value = cant100 & " of " & cant100 + cant0
    If porcent <> 0 Then Cells(fc + 4, 12).Value = porcent & "%"

    porcent = 0
    suma100 = suma100 + cant100
    suma0 = suma0 + cant0
    cant100 = 0
    cant0 = 0
    
    For i = 1 To fc
        If Cells(i, 13).Text = "100%" And Cells(i, 11).MergeCells = False Then cant100 = cant100 + 1
        If Cells(i, 13).Text = "0%" And Cells(i, 11).MergeCells = False Then cant0 = cant0 + 1
    Next
    
    If cant100 <> 0 Then porcent = (cant100 / (cant100 + cant0)) * 100
    Cells(fc + 2, 13).Value = cant100 & " of " & cant100 + cant0
    If porcent <> 0 Then Cells(fc + 4, 13).Value = porcent & "%"

    porcent = 0
    suma100 = suma100 + cant100
    suma0 = suma0 + cant0
    cant100 = 0
    cant0 = 0
    
    For i = 1 To fc
        If Cells(i, 14).Text = "100%" And Cells(i, 11).MergeCells = False Then cant100 = cant100 + 1
        If Cells(i, 14).Text = "0%" And Cells(i, 11).MergeCells = False Then cant0 = cant0 + 1
    Next
    
    If cant100 <> 0 Then porcent = (cant100 / (cant100 + cant0)) * 100
    Cells(fc + 2, 14).Value = cant100 & " of " & cant100 + cant0
    If porcent <> 0 Then Cells(fc + 4, 14).Value = porcent & "%"

    porcent = 0
    suma100 = suma100 + cant100
    suma0 = suma0 + cant0
    cant100 = 0
    cant0 = 0
    
    For i = 1 To fc
        If Cells(i, 15).Text = "100%" And Cells(i, 11).MergeCells = False Then cant100 = cant100 + 1
        If Cells(i, 15).Text = "0%" And Cells(i, 11).MergeCells = False Then cant0 = cant0 + 1
    Next
    
    If cant100 <> 0 Then porcent = (cant100 / (cant100 + cant0)) * 100
    Cells(fc + 2, 15).Value = cant100 & " of " & cant100 + cant0
    If porcent <> 0 Then Cells(fc + 4, 15).Value = porcent & "%"

    porcent = 0
    suma100 = suma100 + cant100
    suma0 = suma0 + cant0
    cant100 = 0
    cant0 = 0
    
    If suma100 <> 0 Then porcent1 = (suma100 / (suma100 + suma0)) * 100
    Cells(fc + 3, 12).Value = suma100 & " of " & suma100 + suma0
    If porcent1 <> 0 Then Cells(fc + 5, 12).Value = porcent1 & "%"
    
    Cells(fc + 2, 11).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Cells(fc + 3, 11).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Cells(fc + 4, 11).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Cells(fc + 5, 11).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    Range(Cells(fc + 3, 12), Cells(fc + 3, 15)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range(Cells(fc + 5, 12), Cells(fc + 5, 15)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range(Cells(fc + 2, 12), Cells(fc + 5, 15)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range(Selection, Selection.End(xlToLeft)).Select

    Range(Cells(fc + 2, 1), Cells(fc + 5, 15)).Select
    Cells(fc + 2, 15).Activate
    With Selection.Font
        .name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 95
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Rows(fc + 1).Select
    Selection.Delete Shift:=xlUp
    
    Columns("L:O").Select
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3Arrows)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValuePercent
        .Value = 50
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValuePercent
        .Value = 70
        .Operator = 7
    End With
    ActiveWindow.SmallScroll Down:=3
    Range("A1").Select
    
    Application.PrintCommunication = True
End Sub
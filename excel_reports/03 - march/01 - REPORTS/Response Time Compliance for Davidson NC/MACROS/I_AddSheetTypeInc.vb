Sub I_AddSheetTypeInc()
    Dim fc As Long
    Dim est As Long
    Dim x As Integer

    Dim name As String
    Dim rangoCelda As String

    Application.ScreenUpdating = Fals
    Application.DisplayAlerts = False

    Worksheets("DATA").Select
    fc = Worksheets("DATA").Cells(Rows.count, "B").End(xlUp).Row

    For Each celda In Worksheets("DATA").Range(Cells(3, 1), Cells(fc, 1))
        name = celda.Value
        If celda.Value <> "" Then
            Worksheets("DATA").Select
            Rows("1:2").Select
            Selection.Copy
            On Error Resume Next
            Worksheets.Add(after:=Worksheets(Worksheets.count)).name = Mid(name, 1, 3)
            Worksheets("Mid(name, 1, 3)").Select
            ActiveSheet.Paste
            Range("A1").Select
            
            'SI LAS CELDAS SON COMBINADAS PARA 1 INCIDENTE
            If celda.MergeCells = True Then
                Worksheets("DATA").Select
                rangoCelda = celda.Address
                nm = Mid(rangoCelda, 4)
                g = contaMerge(Range(rangoCelda))
                Range(Cells(nm, 1), Cells(nm + g - 1, 43)).Select
                Selection.Copy
                Worksheets(Mid(name, 1, 3)).Select
                Range("A3").Select
                ActiveSheet.Paste
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                 :=False, Transpose:=False
                
                x = contaMerge(Range("A3"))
                Range(Cells(3, 1), Cells(x, 1)).Select
                With Selection
                    .HorizontalAlignment = xlGeneral
                    .VerticalAlignment = xlCenter
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                Selection.UnMerge
                
                Range("A3").Select
                Selection.Copy
                Range("B2").Select
                ActiveSheet.Paste
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                Columns("A:A").Select
                Application.CutCopyMode = False
                Selection.Delete Shift:=xlToLeft
                
                Range("A2:K2").Select
                Selection.Merge
                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                
                Range("L1:O1").Select
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
                
                est = Worksheets(Mid(name, 1, 3)).Cells(Rows.count, "D").End(xlUp).Row
                Cells(est + 1, 11).Value = "Compliant Calls for" & " " & Mid(name, 1, 3)
                Cells(est + 2, 11).Value = "Compliant Apparatuses For Call Type" & " " & Mid(name, 1, 3)
                
                Dim cant100 As Long
                Dim cant0 As Long
                Dim porcent As Long
                cant100 = 0
                cant0 = 0
                porcent = 0
                
                For j = 3 To est
                    If Cells(j, 12).Text = "100%" Then cant100 = cant100 + 1
                    If Cells(j, 12).Text = "0%" Then cant0 = cant0 + 1
                Next
                
                porcent = (cant100 / (cant100 + cant0)) * 100
                Cells(est + 1, 12).Value = cant100 & " of " & cant100 + cant0
                Cells(est + 2, 12).Value = porcent & "%"

                cant100 = 0
                cant0 = 0
                porcent = 0

                For j = 3 To est
                    If Cells(j, 13).Text = "100%" Then cant100 = cant100 + 1
                    If Cells(j, 13).Text = "0%" Then cant0 = cant0 + 1
                Next
                
                porcent = (cant100 / (cant100 + cant0)) * 100
                Cells(est + 1, 13).Value = cant100 & " of " & cant100 + cant0
                Cells(est + 2, 13).Value = porcent & "%"
                
                cant100 = 0
                cant0 = 0
                porcent = 0
                
                For j = 3 To est
                    If Cells(j, 14).Text = "100%" Then cant100 = cant100 + 1
                    If Cells(j, 14).Text = "0%" Then cant0 = cant0 + 1
                Next
                
                porcent = (cant100 / (cant100 + cant0)) * 100
                Cells(est + 1, 14).Value = cant100 & " of " & cant100 + cant0
                Cells(est + 2, 14).Value = porcent & "%"
                
                cant100 = 0
                cant0 = 0
                porcent = 0
                
                For j = 3 To est
                    If Cells(j, 15).Text = "100%" Then cant100 = cant100 + 1
                    If Cells(j, 15).Text = "0%" Then cant0 = cant0 + 1
                Next
                
                porcent = (cant100 / (cant100 + cant0)) * 100
                Cells(est + 1, 15).Value = cant100 & " of " & cant100 + cant0
                Cells(est + 2, 15).Value = porcent & "%"
                
                If Cells(est + 1, 12).Value = "0 of 0" Then Cells(est + 2, 12).Value = ""
                If Cells(est + 1, 13).Value = "0 of 0" Then Cells(est + 2, 13).Value = ""
                If Cells(est + 1, 14).Value = "0 of 0" Then Cells(est + 2, 14).Value = ""
                If Cells(est + 1, 15).Value = "0 of 0" Then Cells(est + 2, 15).Value = ""
                
                Range(Cells(est + 1, 11), Cells(est + 2, 15)).Select
                With Selection
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                
                Columns("L:O").Select
                Selection.ColumnWidth = 8.4
                Range("A:A,B:B,C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O").Select
                With Selection.Font
                    .name = "Calibri"
                    .Size = 10
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
                
                Columns("D:D").ColumnWidth = 8.43
                Columns("E:E").ColumnWidth = 7.14
                Columns("F:F").ColumnWidth = 7.14
                Columns("G:G").ColumnWidth = 7.14
                Columns("H:H").ColumnWidth = 7.14
                Columns("I:I").ColumnWidth = 7.14
                Columns("J:J").ColumnWidth = 7.14
                Columns("K:K").ColumnWidth = 7.14
                Columns("A:A").ColumnWidth = 9.43
                Columns("A:A").ColumnWidth = 9.71
                Columns("B:B").ColumnWidth = 8.86
                Columns("C:C").ColumnWidth = 14.71
                
                Range(Cells(1, 1), Cells(est, 3)).Select
                Range("A3").Activate
                With Selection
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
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
                
                Range(Cells(1, 1), Cells(est + 2, 15)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                End With
                
                Range("A2:K2").Select
                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                
                Range(Cells(est + 1, 1), Cells(est + 1, 11)).Select
                Cells(est + 1, 11).Activate
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
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
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                Range(Cells(est + 2, 1), Cells(est + 2, 11)).Select
                Cells(est + 2, 11).Activate
                 With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
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
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                Range(Cells(1, 1), Cells(est + 2, 15)).Select
                Cells(est + 2, 15).Activate
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
                
                With Selection
                    .VerticalAlignment = xlCenter
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                End With
                
                Range("A1:O1").Select
                Range("L1").Activate
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                Range("A2:O2").Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                Range(Cells(3, 1), Cells(est + 2, 15)).Select
                Cells(est + 2, 15).Activate
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 15528702
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("A1:O2").Select
                Range("O2").Activate
                Selection.Font.Bold = True
                Range(Cells(est + 1, 1), Cells(est + 2, 11)).Select
                Cells(est + 2, 11).Activate
                Selection.Font.Bold = True
                Range(Cells(est + 1, 12), Cells(est + 2, 15)).Select
                Cells(est + 2, 15).Activate
                Selection.Font.Bold = True


                Range(Cells(est + 1, 1), Cells(est + 1, 15)).Select
                Cells(est + 1, 15).Activate
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                Range(Cells(est + 2, 1), Cells(est + 2, 15)).Select
                Cells(est + 2, 15).Activate
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                
                Range("M3").Select
                Rows("2:2").Select
                Selection.AutoFilter
                Worksheets(Mid(name, 1, 3)).AutoFilter.Sort.SortFields.Clear
                Worksheets(Mid(name, 1, 3)).AutoFilter.Sort.SortFields.Add2 Key:=Range( _
                    "A2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
                    xlSortNormal
                With Worksheets(Mid(name, 1, 3)).AutoFilter.Sort
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Selection.AutoFilter
                Range("A1:K1").Select
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
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                
                Range("A3").Select
                Do While ActiveCell.Value <> ""
                    valor = ActiveCell
                    contarTI = Application.WorksheetFunction.CountIf(Columns(1), valor)
                    If contarTI > 0 Then
                        Range(Cells(ActiveCell.Row, 1), Cells(ActiveCell.Row + contarTI - 1, 1)).Merge
                    End If
                    ActiveCell.Offset(1, 0).Select
               Loop
               ActiveCell.Offset(1, 0).Select
               
               Range("B3").Select
                Do While ActiveCell.Value <> ""
                    valor = ActiveCell
                    contarTI = Application.WorksheetFunction.CountIf(Columns(2), valor)
                    If contarTI > 0 Then
                        Range(Cells(ActiveCell.Row, 2), Cells(ActiveCell.Row + contarTI - 1, 2)).Merge
                    End If
                    ActiveCell.Offset(1, 0).Select
               Loop
               ActiveCell.Offset(1, 0).Select
                
            End If
            'SI CELDA UNICA PARA 1 INCIDENTE
            If celda.MergeCells = False Then
                Worksheets("DATA").Select
                rangoCelda = celda.Address
                nm = Mid(rangoCelda, 4)
                Range(Cells(nm, 1), Cells(nm, 43)).Select
                Selection.Copy
                Worksheets(Mid(name, 1, 3)).Select
                Range("A3").Select
                ActiveSheet.Paste
                
                Range("A3").Select
                Selection.Copy
                Range("B2").Select
                ActiveSheet.Paste
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                Columns("A:A").Select
                Application.CutCopyMode = False
                Selection.Delete Shift:=xlToLeft
                
                Range("L1:O1").Select
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
                
                Range("K4").Value = "Compliant Calls for" & " " & Mid(name, 1, 3)
                Range("K5").Value = "Compliant Apparatuses For Call Type" & " " & Mid(name, 1, 3)
                If Range("L3").Text = "100%" Then
                    Range("L4").Value = "1 of 1"
                    Range("M4").Value = "0 of 0"
                    Range("N4").Value = "0 of 0"
                    Range("O4").Value = "0 of 0"
                    Range("L5").Value = "100%"
                End If
                If Range("L3").Text = "0%" Then
                    Range("L4").Value = "0 of 0"
                    Range("M4").Value = "0 of 0"
                    Range("N4").Value = "0 of 0"
                    Range("O4").Value = "0 of 0"
                    Range("L5").Value = "0%"
                End If
                
                Range("K4:O5").Select
                With Selection
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                
                Columns("L:O").Select
                Selection.ColumnWidth = 8.4
                Range("A:A,B:B,C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O").Select
                With Selection.Font
                    .name = "Calibri"
                    .Size = 10
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
                
                Columns("D:D").ColumnWidth = 8.43
                Columns("E:E").ColumnWidth = 7.14
                Columns("F:F").ColumnWidth = 7.14
                Columns("G:G").ColumnWidth = 7.14
                Columns("H:H").ColumnWidth = 7.14
                Columns("I:I").ColumnWidth = 7.14
                Columns("J:J").ColumnWidth = 7.14
                Columns("K:K").ColumnWidth = 7.14
                Columns("A:A").ColumnWidth = 9.43
                Columns("A:A").ColumnWidth = 9.71
                Columns("B:B").ColumnWidth = 8.86
                Columns("C:C").ColumnWidth = 14.71
                Range("1:1,3:3").Select
                Range("A3").Activate
                With Selection
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
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
                
                Range("A1:O5").Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                End With
                
                Range("A2:K2").Select
                Selection.Merge
                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                
                Range("A4:K4").Select
                Range("K4").Activate
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
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
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                
                Range("A5:K5").Select
                Range("K5").Activate
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
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
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                
                Range("A1:O5").Select
                Range("O5").Activate
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
                Range("A1:O1").Select
                Range("L1").Activate
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                Range("A2:O2").Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                Range("A3:O5").Select
                Range("O5").Activate
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 15528702
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("A1:O2,A4:K4,A5:K5,L5,L4,M4,N4,O4,O5,N5,M5").Select
                Range("M5").Activate
                Selection.Font.Bold = True
                Range("A4:O5").Select
                Range("O5").Activate
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                
                Rows("1:5").Select
                With Selection
                    .VerticalAlignment = xlCenter
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                End With
                
                Range("M3").Select
                ActiveCell.Offset(1, 0).Select
                
            End If
        End If
    Next
    
    Worksheets("DATA").Select
    Range("A1").Select
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
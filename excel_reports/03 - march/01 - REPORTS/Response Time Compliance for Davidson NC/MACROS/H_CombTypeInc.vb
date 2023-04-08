Sub H_CombTypeInc()
    Sheets("DATA").Select
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Range("A3").Select
    
    Do While ActiveCell.Value <> ""
        valor = ActiveCell
        contarTI = Application.WorksheetFunction.CountIf(Columns(1), valor)
        If contarTI > 0 Then
            Range(Cells(ActiveCell.Row, 1), Cells(ActiveCell.Row + contarTI - 1, 1)).Merge
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
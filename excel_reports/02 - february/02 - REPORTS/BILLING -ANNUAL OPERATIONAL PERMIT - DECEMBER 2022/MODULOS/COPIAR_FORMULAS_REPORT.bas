Attribute VB_Name = "Módulo9"
Sub CopiarFormulasReport()
   
    'Set hoja = ThisWorkbook.Sheets("DATA")
    Sheets("Output Report").Select
    Range("A2", "A" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("A2").Formula
    Range("D2", "D" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("D2").Formula
    Range("F2", "F" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("F2").Formula
    Range("G2", "G" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("G2").Formula
    Range("H2", "H" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("H2").Formula
    Range("I2", "I" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("I2").Formula
    Range("K2", "K" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("K2").Formula
    Range("L2", "L" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("L2").Formula
    Range("N2", "N" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("N2").Formula
    Range("O2", "O" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("O2").Formula
    Range("P2", "P" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("P2").Formula
    Range("R2", "R" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("R2").Formula
    Range("V2", "V" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("V2").Formula
    Range("Z2", "Z" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("Z2").Formula
    Range("AB2", "AB" & Sheets("DATA").Cells(Rows.count, 1).End(xlUp).Row) = Range("AB2").Formula
    
  
End Sub


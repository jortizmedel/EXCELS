Attribute VB_Name = "Módulo7"
Sub CopiarFormulas()
   
    Sheets("DATA").Select
    Range("K2", "K" & Cells(Rows.count, 1).End(xlUp).Row) = Range("K2").Formula
    Range("L2", "L" & Cells(Rows.count, 1).End(xlUp).Row) = Range("L2").Formula
    Range("M2", "M" & Cells(Rows.count, 1).End(xlUp).Row) = Range("M2").Formula
    Range("N2", "N" & Cells(Rows.count, 1).End(xlUp).Row) = Range("N2").Formula
    Range("O2", "O" & Cells(Rows.count, 1).End(xlUp).Row) = Range("O2").Formula
    Range("P2", "P" & Cells(Rows.count, 1).End(xlUp).Row) = Range("P2").Formula
    Range("Q2", "Q" & Cells(Rows.count, 1).End(xlUp).Row) = Range("Q2").Formula
    Range("R2", "R" & Cells(Rows.count, 1).End(xlUp).Row) = Range("R2").Formula
    Range("S2", "S" & Cells(Rows.count, 1).End(xlUp).Row) = Range("S2").Formula
    Range("T2", "T" & Cells(Rows.count, 1).End(xlUp).Row) = Range("T2").Formula
    Range("U2", "U" & Cells(Rows.count, 1).End(xlUp).Row) = Range("U2").Formula
    Range("V2", "V" & Cells(Rows.count, 1).End(xlUp).Row) = Range("V2").Formula
    Range("W2", "W" & Cells(Rows.count, 1).End(xlUp).Row) = Range("W2").Formula
    Range("X2", "X" & Cells(Rows.count, 1).End(xlUp).Row) = Range("X2").Formula
    Range("Y2", "Y" & Cells(Rows.count, 1).End(xlUp).Row) = Range("Y2").Formula
    Range("Z2", "Z" & Cells(Rows.count, 1).End(xlUp).Row) = Range("Z2").Formula
  
End Sub

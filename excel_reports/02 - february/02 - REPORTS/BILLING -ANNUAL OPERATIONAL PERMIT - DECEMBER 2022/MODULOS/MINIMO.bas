Attribute VB_Name = "Módulo13"
Sub sacarMinVBA()
    Dim numero As Double
    Dim comparar As Long
    Dim min As Long
    Dim fC As Long
    Dim i As Long
    Dim RJ As Variant
    Dim RA As Variant
    
    
    fC = Worksheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    RJ = "J" & CStr(fC)
    
    
    
    For i = 2 To fC
      
      comparar = Worksheets("DATA").Cells(i, "A").Value
      numero = WorksheetFunction.MinIfs(Range(Cells(2, 10), Cells(fC, 10)), Range(Cells(2, 1), Cells(fC, 1)), comparar)
      Worksheets("DATA").Cells(i, "AA").Value = numero
    
    Next i
      
    
End Sub

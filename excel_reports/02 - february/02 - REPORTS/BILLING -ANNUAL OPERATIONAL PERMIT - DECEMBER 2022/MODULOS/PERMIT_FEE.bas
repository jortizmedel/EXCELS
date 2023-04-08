Attribute VB_Name = "Módulo11"
Sub PermitFee()
    Dim suma As Long
    Dim menor As Long
    Dim invoiceID As Long
    Dim count As Long
    Dim fC As Long
    Dim i As Long
    Dim y As Long
    
    fC = Worksheets("DATA").Cells(Rows.count, "A").End(xlUp).Row
    
    For i = 2 To fC
      count = 0
      suma = 0
      invoiceID = Worksheets("DATA").Cells(i, "A").Value
      For y = 2 To fC 'para sacar contador
        If Worksheets("DATA").Cells(y, "A").Value = invoiceID Then
            count = count + 1
            suma = suma + Worksheets("DATA").Cells(y, "J").Value
            menor = Worksheets("DATA").Cells(i, "AA").Value
        End If
      Next y
      If count = 1 Then
        Worksheets("DATA").Cells(i, "AB").Value = suma
      End If
      If count = 2 Then
        Worksheets("DATA").Cells(i, "AB").Value = (suma + (menor * 0.8)) - menor
      End If
      If count > 2 Then
        Worksheets("DATA").Cells(i, "AB").Value = (suma + (menor * 0.7)) - menor
      End If
      
    Next i
End Sub

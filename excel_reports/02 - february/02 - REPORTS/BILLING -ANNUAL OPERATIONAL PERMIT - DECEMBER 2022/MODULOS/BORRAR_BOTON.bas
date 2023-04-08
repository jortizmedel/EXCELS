Attribute VB_Name = "Módulo10"
Sub COPY_DATA()
'
' COPY_DATA Macro
'

'
    Application.Run _
        "'Report.xlsm'!CopiarHoja"
    Range("A2").Select
    Sheets("Output Report").Select
    Range("A3").Select
    Application.CutCopyMode = False
    Application.Run _
        "'Report.xlsm'!CopiarFormulas"
    Range("A2").Select
    Sheets("Output Report").Select
    Range("A3").Select
End Sub



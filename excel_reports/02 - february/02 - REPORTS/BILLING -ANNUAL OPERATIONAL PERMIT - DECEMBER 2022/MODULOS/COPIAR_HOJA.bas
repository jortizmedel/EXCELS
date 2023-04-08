Attribute VB_Name = "Módulo6"
Sub CopiarHoja()
'
' CopiarHoja Macro
'

'
    Sheets("Input Report").Select
    Columns("A:J").Select
    Selection.Copy
    Sheets("DATA").Select
    Columns("A:A").Select
    ActiveSheet.Paste
End Sub

Attribute VB_Name = "M�dulo4"
Sub BorrarHojas()
'
' BorrarHojas Macro
'

'
    Sheets("DATA").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Instructions").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Input Report").Select
    ActiveWindow.SelectedSheets.Delete
End Sub

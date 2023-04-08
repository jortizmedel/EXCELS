Attribute VB_Name = "Módulo12"
Sub RUN_REPORT()
'
' RUN_REPORT Macro
'

'
    Sheets("DATA").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 20
    Application.Run "Report.xlsm!sacarMinVBA"
    Application.Run "Report.xlsm!PermitFee"
    Sheets("Output Report").Select
    Application.Run "Report.xlsm!CopiarFormulasReport"
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Application.Run "Report.xlsm!QuitarFormulas"
    ActiveWindow.LargeScroll ToRight:=-1
    Range("A3").Select
    Application.CutCopyMode = False
    Application.Run "Report.xlsm!EliminarFilasVacias"
    Range("A2").Select
    Application.Run "Report.xlsm!BorrarHojas"
    Application.Run "Report.xlsm!BorrarBoton"
End Sub


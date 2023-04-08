Sub RUN_REPORT()
    Dim joja As Object
    Dim nombre As String
    
    Application.CutCopyMode = False
    Application.Run "'FOR MACROS.xltm'!A_copyData"
    Application.Run "'FOR MACROS.xltm'!B_timeCalc"
    Application.Run "'FOR MACROS.xltm'!C_rearrangeDATA"
    Application.Run "'FOR MACROS.xltm'!D_compareFA"
    Application.Run "'FOR MACROS.xltm'!E_compareSA"
    Application.Run "'FOR MACROS.xltm'!F_compareTA"
    Application.Run "'FOR MACROS.xltm'!G_compareFA"
    Application.Run "'FOR MACROS.xltm'!H_CombTypeInc"
    Application.Run "'FOR MACROS.xltm'!I_AddSheetTypeInc"
    Application.Run "'FOR MACROS.xltm'!J_REPORT"
    Application.Run "'FOR MACROS.xltm'!K_DeleteSheets"
    
    'Worksheets("REPORT").Select
    'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=ThisWorkbook.Path & "\Response Time Compliance_Report " & Format(Now(), "YYYY-MM-DD"), Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    Worksheets("Response Time Compliance - Raw ").Select
    Cells.Clear
    Range("A1").Select
    Worksheets("REPORT").Select
    Range("A1").Select
End Sub
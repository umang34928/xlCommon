Attribute VB_Name = "sPDF"
Function D0(SheetsToSaveName As String, FileSaveAsName As String)
'sPDF.D0("Sheet,Sheet1,Sheet2","PDF Name")
Dim SavePath As String

Application.DisplayAlerts = False
ActiveWorkbook.Save
SavePath = ActiveWorkbook.path ' Required as new sheet does not have a path yet

Sheets(Array(SheetsToSaveName)).Select
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=SavePath & "\" & SheetsToSaveName, _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

MsgBox ("Sheet(s) " & SheetToSaveName & "Saved as " & vbNewLine & SavePath & "\" & FileSaveAsName & ".pdf")

Application.DisplayAlerts = True

End Function


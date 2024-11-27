Attribute VB_Name = "sCSV"
Function D0(SheetToSaveName As String, FileSaveAsName As String, Optional message As String)
'call sCSV.D0("Sheet","Csv Name") 'Creates/Saves "Sheet" as "Sheet Name" .csv
Dim SavePath As String

Application.DisplayAlerts = False
ActiveWorkbook.Save
SavePath = ActiveWorkbook.Path ' Required as new sheet does not have a path yet

ActiveWorkbook.Sheets(SheetToSaveName).Copy
ActiveWorkbook.SaveAs Filename:=SavePath & "\" & FileSaveAsName, FileFormat:=xlCSV, CreateBackup:=True
ActiveWorkbook.Close

MsgBox ("Sheet " & SheetToSaveName & "Saved as " & vbNewLine & SavePath & "\" & FileSaveAsName & ".csv" & vbNewLine & message)

Application.DisplayAlerts = True

End Function


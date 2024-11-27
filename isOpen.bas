Attribute VB_Name = "isOpen"
Function isOpen(strBookName As String) As Boolean
' e.Info ("Workbook Status Rajput Calculation " & e.isOpen("xlCommon.xlsm")) Ensure not to use full path of workbook

    Dim oBook As Workbook
    On Error Resume Next
    
    Set oBook = Workbooks(strBookName)

        If oBook Is Nothing Then
            isOpen = False
        Else
            isOpen = True
        End If

End Function

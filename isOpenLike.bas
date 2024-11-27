Attribute VB_Name = "isOpenLike"
Function isOpenLike(strBookName As String) As Boolean
' e.Info ("Workbook Status Rajput Calculation " & e.isOpenLike("xlCommon")) Ensure not to use full path of workbook

    Dim oBook As Workbook
    On Error Resume Next
        isOpenLike = False
        
    For Each oBook In Application.Workbooks
        If oBook.Name Like strBookName & "*" Then
            Set oBook = Workbooks(oBook)
              isOpenLike = True
            Exit Function
        End If
    Next oBook

End Function

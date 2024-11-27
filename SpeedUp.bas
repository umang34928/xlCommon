Attribute VB_Name = "SpeedUp"
Function D0(FValue As Boolean) As String
'Call SpeedUp.Do (True) 'Apply Calculations/Enable Screen update
'Call SpeedUp.Do (False) ' Disable Screen refresh/Disable calculation to speed up

Application.ScreenUpdating = FValue
Application.DisplayStatusBar = FValue

    If FValue Then
    Application.Calculation = xlCalculationAutomatic
    Else
    Application.Calculation = xlCalculationManual
    End If

End Function

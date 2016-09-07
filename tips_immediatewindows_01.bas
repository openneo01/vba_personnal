Sub immediate_windows_clear_01()
'Desrc : Clear immediate windows
'src = https://social.msdn.microsoft.com/Forums/office/en-US/da87e63f-676b-4505-adeb-564257a56cfe/vba-to-clear-the-immediate-window?forum=exceldev
'Clear 
Dim x As Long
For x = 1 To 10
Debug.Print x
Next

Debug.Print Now
Application.SendKeys "^g ^a {DEL}"
End Sub


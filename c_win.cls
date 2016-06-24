'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_win (Module de classe)
' Project   : VBAProject
' Author    : GRT GAZ
' Date      : 02/05/2016
' Purpose   : **MANDATORY**
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1
Private cWinFile As c_win_file

'***************************
'*******INI / TERM *********
'***************************

Private Sub Class_Initialize()
    Set cWinFile = New c_win_file
End Sub

Private Sub Class_Terminate()
    Set cWinFile = Nothing

End Sub


'********************************
'********** SUB CLASS ***********
'********************************


'***************************
'******* PROPERTIES *********
'***************************

Public Property Get oWinFile() As c_win_file
    Set oWinFile = cWinFile
End Property


'********************************
'********** FUNCTIONS ***********
'********************************




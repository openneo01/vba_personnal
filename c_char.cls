'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_char (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 28/06/2016
' Purpose   : Char manipulation
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1


'********************************
'********* INIT / TERM **********
'********************************
Private Sub Class_Initialize()
Debug.Print "Load Module de classe : Classe1"

End Sub

Private Sub Class_Terminate()
Debug.Print "Unload Module de classe : Classe1"

End Sub

'********************************
'********** SUB CLASS ***********
'********************************




'********************************
'********* PROPERTIES ***********
'********************************


'********************************
'********** FUNCTIONS ***********
'********************************

Public Function char_occurence(ByVal char_string As String, ByVal char_search As String) As Long
char_occurence = Len(char_string) - Len(Replace(char_string, char_search, ""))
End Function

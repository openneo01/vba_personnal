'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_win_file (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 02/05/2016
' Purpose   : win file functions related
'
'///////////////////////////////////////////////////////////////////////////////////////
'

Option Explicit
Option Base 1

'***************************
'*******INI / TERM *********
'***************************

Private Sub Class_Initialize()

End Sub

'***************************
'******* FUNCTIONS **********
'***************************

Public Function file_dir(ByRef Folderpath As String) As Boolean
'
'=======================================================================================
' Procedure : file_dir (Function)
' Module    : c_win_file (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 25/04/2016
' Comments  : Check target dir access
' Unit Test : (MISTERNEO) 25/04/2016 16:37 | Description [OK]
' Arg./i    :
'           - [Folderpath] folder to process
'           -
'           -
' Arg./o    : Boolean(b)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'02/05/2016         GRT GAZ             Initiate
'
'=======================================================================================
'

Dim Fso As Object

Set Fso = CreateObject("scripting.filesystemobject")


'If Right(Folderpath, 1) <> "\" Then
    'Folderpath = Folderpath & "\"
'End If

If Fso.FolderExists(Folderpath) = True Then file_dir = True

End Function


Public Function file_exist(ByVal filePath As String) As Boolean
If Dir(filePath) <> "" Then file_exist = True
End Function

Public Function file_read(ByVal filePath As String, Optional splitLineCar As Integer) As String()
'
'=======================================================================================
' Procedure : file_read (Function)
' Module    : c_win_file (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 17/05/2016
' Comments  : fast read file
' Unit Test : (MISTERNEO) 17/05/2016 14:22 | Description [OK]
' Arg./i    :
'           - [filePath] File (fulle path) to process
'           - [splitLineCar] char identification for new line in file
'           -
' Arg./o    : Variant(v)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'17/05/2016         GRT / GAZ               Initiate
'
'=======================================================================================
'

Dim nFile As Integer
Dim buffer As String

nFile = FreeFile

Open filePath For Binary As #nFile
    buffer = String$(LOF(nFile), " ")
    Get #nFile, , buffer
Close #nFile
        
'Note : split line into array
file_read = Split(buffer, splitLineCar)

End Function


Public Function file_transform_dictionary(ByVal data_array_string, ByVal data_header As Collection, ByVal asHeader As Boolean, Optional ByVal splitDataChar As String = ";") As Scripting.Dictionary
'
'=======================================================================================
' Procedure : file_transform_dictionary (Function)
' Module    : c_win_file (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 24/06/2016
' Comments  : **MANDATORY**
' Unit Test : (MISTERNEO) 24/06/2016 15:43 | Description [OK/NOK]
' Arg./i    :
'           - [**MANDATORY**] / [NO PARAM]
'           -
'           -
' Arg./o    : Scripting.Dictionary(cl)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'24/06/2016         GRT / GAZ               Initiate
'
'=======================================================================================
'
Dim oDatas As New Scripting.Dictionary
Dim uKey
Dim ODataLine() As String
Dim i As Long, j As Byte, l As Long

If asHeader = True Then j = 1 Else j = 0
For i = LBound(data_array_string) + j To UBound(data_array_string)
    ODataLine = Split(data_array_string(i), splitDataChar)
    For Each uKey In data_header
        Debug.Print uKey
    Next
    'oDatas.Add
Next i

End Function





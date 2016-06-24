'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_error (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 02/05/2016
' Purpose   : error management
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1


Public Enum err_level
    errl_info = 0
    errl_warning = 1
    err_fatal = 2
End Enum

Public Enum err_type
    errt_process = 0
    errt_datasource = 1
    errt_xlsobj = 2
    errt_datatype = 3
    errt_writeread = 4
    errt_calcul = 5
    errt_system = 6
    errt_registry = 7
    errt_dll = 8
    errt_undefined = 9
End Enum

Private m_serrPath As String

'***************************
'*******INI / TERM *********
'***************************
Private Sub Class_Initialize()
Debug.Print "Load Module de classe : c_error"
End Sub

Private Sub Class_Terminate()
Debug.Print "Unload Module de classe : c_error"

End Sub


'****************************
'******* PROPERTIES *********
'****************************

Public Property Get errPath() As String
    errPath = m_serrPath
End Property

Public Property Let errPath(ByVal serrPath As String)
    m_serrPath = serrPath
End Property


'****************************
'******* FUNCTIONS  *********
'****************************

Public Function log(ByVal errLevel As err_level, ByVal errType As err_type, ByVal errLog As String)
'
'=======================================================================================
' Procedure : log (Function)
' Module    : c_error (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 01/06/2016
' Comments  : log error from code
' Unit Test : (MISTERNEO) 01/06/2016 11:15 | Description [OK]
' Arg./i    :
'           - [errLevel] err level 3 = end the code
'           - [errType] type of error for analysis
'           - [errLog] log message
'           -
' Arg./o    : Variant(v)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'01/06/2016         MISTERNEO               Initiate
'
'=======================================================================================
'
    Dim owkb As Workbook
    Dim FileNum As Single
    
    FileNum = FreeFile
    
    Open errPath For Append As #FileNum
        Print #FileNum, Now & ";" & errLevel & ";" & errType & ";" & errLog
    Close #FileNum
    
    'Note : test if code need to be stopped
    continue errLevel, errType
    
End Function

Private Function continue(errLevel As err_level, ByVal errType As err_type)
'
'=======================================================================================
' Procedure : continue (Function)
' Module    : c_error (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 01/06/2016
' Comments  : Test if code need to continue or not
' Unit Test : (MISTERNEO) 01/06/2016 13:42 | Description [OK]
' Arg./i    :
'           - [errLevel] err level 3 = end the code
'           - [errType] type of error for analysis
'           -
' Arg./o    : Variant(v)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'01/06/2016         MISTERNEO               Initiate
'
'=======================================================================================
'
Select Case errLevel
    Case err_level.errl_info
        Application.StatusBar = "Small errors detected with no impact on datas... Check log for further details :-|"
        
    Case err_level.errl_warning
        Application.StatusBar = "Warning. App still running... Check log for further details :-("
        
    Case err_level.err_fatal
        Application.StatusBar = "fatal error. App stop... Check log for further details :-O"
        End
End Select
End Function


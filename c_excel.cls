'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_excel (Module de classe)
' Project   : VBAProject
' Author    : misterneo
' Date      : 02/05/2016
' Purpose   : ms excel FUNCTIONS
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1

Private cIde As c_excel_ide
Private cSheet As c_excel_sheet
Private cData As c_excel_data

Private m_clApp As Workbook
Private m_oclErr As c_error


'***************************
'*******INI / TERM *********
'***************************

Private Sub Class_Initialize()
    Debug.Print "Load Module de classe : c_excel"
    Set cIde = New c_excel_ide
    Set cSheet = New c_excel_sheet
    Set cData = New c_excel_data
End Sub

Private Sub Class_Terminate()
    Debug.Print "Unload Module de classe : c_excel"
    Set cIde = Nothing
    Set cSheet = Nothing
    Set m_oclErr = Nothing
    Set m_clApp = Nothing
End Sub

'***************************
'******* SUB CLASS *********
'***************************

Public Property Get clErr() As c_error
    Set clErr = m_oclErr
End Property

Public Property Set clErr(oclErr As c_error)
    Set m_oclErr = oclErr
End Property


Public Property Get oIde() As c_excel_ide
    Set oIde = cIde
End Property

Public Property Get oSheet() As c_excel_sheet
    Set oSheet = cSheet
End Property

Public Property Get oData() As c_excel_data
    Set oData = cData
End Property


'***************************
'******* PROPERTIES *********
'***************************

Public Property Get App() As Workbook

    If m_clApp Is Nothing Then
        Set App = ThisWorkbook
    Else
        Set App = m_clApp
    End If
    
End Property

Public Property Set App(clApp As Workbook)
    Set m_clApp = clApp
End Property

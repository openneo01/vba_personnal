'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_excel_sheet (Module de classe)
' Project   : VBAProject
' Author    : misterneo
' Date      : 26/05/2016
' Purpose   : excel sheets related functions
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1



Private m_ocListObject As c_excel_sheet_listobject
Private m_clworksheet As Worksheet

'***************************
'*******INI / TERM *********
'***************************

Private Sub Class_Initialize()
    Debug.Print "Load Module de classe : c_excel_sheet"

    Set m_ocListObject = New c_excel_sheet_listobject
End Sub

Private Sub Class_Terminate()
    Debug.Print "Unload Module de classe : c_excel_sheet"

    Set m_ocListObject = Nothing
    Set m_clworksheet = Nothing
End Sub


'***************************
'******* SUBCLASS *********
'***************************

Public Property Get cllistobject() As c_excel_sheet_listobject
    Set cllistobject = m_ocListObject
End Property

Public Property Set cllistobject(olistobject As c_excel_sheet_listobject)
    Set m_ocListObject = olistobject
End Property

'***************************
'******* PROPERTIES *********
'***************************

Public Property Get Worksheet() As Worksheet
    Set Worksheet = m_clworksheet
End Property

Public Property Set Worksheet(clworksheet As Worksheet)
    Set m_clworksheet = clworksheet
End Property


'********************************
'********** FUNCTIONS ***********
'********************************

Public Function sheet_clear()

With Worksheet
    .Cells.Clear
End With

End Function
Public Function sheet_exist(ByVal worksheet_name As String, Optional workbook_target As Excel.Workbook)
Dim oWks As Excel.Worksheet

On Error Resume Next

If IsMissing(workbook_target) Then workbook_target = ThisWorkbook
Set oWks = workbook_target.Sheets(worksheet_name)
On Error GoTo 0
sheet_exist = Not oWks Is Nothing
End Function

Public Function sheet_listobjects_delete(ByVal worksheet_name As String, Optional workbook_target As Excel.Workbook)
Dim oList As listobject

If IsMissing(workbook_target) Then workbook_target = ThisWorkbook
Set Worksheet = workbook_target.Sheets(worksheet_name)

With cllistobject
    For Each oList In Worksheet.ListObjects
        Set .listobject = oList: .delete_
    Next
    
End With

End Function


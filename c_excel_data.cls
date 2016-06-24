'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_excel_data (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 17/05/2016
' Purpose   : Excel data functions
'
'///////////////////////////////////////////////////////////////////////////////////////
'

Option Explicit
Option Base 1


Private m_clErr As c_error



'********************************
'********* INIT / TERM **********
'********************************


Private Sub Class_Initialize()
Debug.Print "Load Module de classe : c_excel_data"

End Sub

Private Sub Class_Terminate()
Debug.Print "Unload Module de classe : c_excel_data"

    Set m_clErr = Nothing
End Sub


'********************************
'********** SUB CLASS ***********
'********************************
Public Property Get cErr() As c_error
    Set cErr = m_clErr
End Property

Public Property Set cErr(oErr As c_error)
    Set m_clErr = cErr
End Property


'********************************
'********* PROPERTIES ***********
'********************************


'********************************
'********** FUNCTIONS ***********
'********************************

Public Function read_xls_odbc(ByVal xls_file_path As String, ByVal xls_file_version As String, ByVal SheetName As String, ByVal rangeRef As String) As Variant
      '
      '=======================================================================================
      ' Procedure : EXTRACT_DATA_FROM_XLS (Function)
      ' Module    : cls_WORKBOOK (Module de classe)
      ' Project   : VBAProject
      ' Author    : MISTERNEO
      ' Date      : 19/01/2015
      ' Comments  : READ DATA FROM CLOSE EXCEL WORKBOOK AND POPULATE ARRAY
      '             CAN'T READ DATA IF WORKBOOK IS ALREADY OPEN !
            'I * Visual Basic For Applications
            'I * Microsoft Excel 12.0 ObjectLibrary
            'I * Microsoft ADO Ext. 6.0 for DDL and Security
            'I * Microsoft ActiveX Data Objects 6.1 Library
            'I * Microsoft AcitveX Data Objects Recordset 6.0 Library
            '   Warning : Text column limitatiion = max 255 caracteres

      ' Arg./i    :
      '           - [TARGET_FILE_NAME] FULL PATH TO EXCEL WORKBOOK
      '           - [TARGET_SHEET_NAME] NAME OF TARGET SHEET WHERE TO FIND DATA
      '           - [TARGET_RANGE] TARGET RANGE ex : A1:D45
      '           - [TARGET_XLS_VERSION] VERSION OF EXCEL WORKBOOK (default : 2007)
      '           -
      ' Arg./o    : Variant(v)
      '
      'Changes--------------------------------------------------------------------------------
      'Date               Programmer                      Change
      '19/01/2015         MISTERNEO               Initiate
      '
      '=======================================================================================
      '
10    On Error GoTo Err_Handler

      Dim sCnx As String
      Dim oRst As ADODB.Recordset
      Dim oRcd As ADODB.Record
      Dim oFld As Field

      Dim sql As String
      Dim oDatas() As Variant
      Dim orng() As Variant
      Dim i As Long

20    Select Case xls_file_version
          Case "2007"
30            sCnx = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source=" & xls_file_path & ";" & _
                 "Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
                 sCnx = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & xls_file_path & "';Extended Properties='Excel 12.0;IMEX=1;HDR=YES';"
                 
40        Case "2010"
50            sCnx = "Provider=Microsoft.ACE.OLEDB.14.0;" & _
                "Data Source=" & xls_file_path & ";" & _
                "';Extended Properties='Excel 14.0;IMEX=1;HDR=YES';"
          
60        Case "2013"
70            sCnx = "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                 "Data Source=" & xls_file_path & ";" & _
                 "';Extended Properties='Excel 16.0;IMEX=1;HDR=YES';"
          
80    End Select
Debug.Print sCnx
      'Note : Set SQL command
90    sql = "SELECT * FROM " & "[" & SheetName & "$" & rangeRef & "] WHERE Identifiant <>''"

      'Note : Ini recordset and connect
100   Set oRst = New ADODB.Recordset
110   oRst.Open sql, sCnx, adOpenStatic, adLockReadOnly

      'Note : Loop throught recordset
      Dim fld As Field
120   With oRst
        
130       oDatas = .GetRows
140   End With
    
      'Note : close recordset
150   oRst.Close

Err_Exit:
160       read_xls_odbc = oDatas
          'Note : Delete object
170       Set oRst = Nothing
          'Note : Exit
180       Exit Function

Err_Handler:
          'Note : Exit Function
190       cErr.log 3, errt_datasource, Err.Number & "(" & Erl & ");" & Err.Description & ";VBAProject;c_excel_data (Module de classe);read_xls_odbc (Function))"
200       GoTo Err_Exit
End Function

Public Function read_xls(ByVal xls_file_path As String, ByVal xls_file_version As String, ByVal SheetName As String, ByVal rangeRef As String) As Variant
      '
      '=======================================================================================
      ' Procedure : read_xls (Function)
      ' Module    : c_excel_data (Module de classe)
      ' Project   : VBAProject
      ' Author    : MISTERNEO
      ' Date      : 24/06/2016
      ' Comments  : read xls file
      ' Unit Test : (MISTERNEO) 24/06/2016 11:52 | Description [OK]
      ' Arg./i    :
            '           - [TARGET_FILE_NAME] FULL PATH TO EXCEL WORKBOOK
            '           - [TARGET_SHEET_NAME] NAME OF TARGET SHEET WHERE TO FIND DATA
            '           - [TARGET_RANGE] TARGET RANGE ex : A1:D45
            '           - [TARGET_XLS_VERSION] VERSION OF EXCEL WORKBOOK (default : 2007)
      '           -
      ' Arg./o    : Variant(v)
      '
      'Changes--------------------------------------------------------------------------------
      'Date               Programmer                      Change
      '24/06/2016         MISTERNEO               Initiate
      '
      '=======================================================================================
      '
      Dim owkb As Workbook, owkb_current As Workbook
      Dim orng As Range
      Dim data As Variant

10    On Error GoTo Err_Handler

20    Set owkb_current = ThisWorkbook
30    Set owkb = Workbooks.Open(xls_file_path, False, True)

40    With owkb
50        data = .Worksheets(SheetName).Range("A1").CurrentRegion
60        .Close SaveChanges:=False
70    End With

80    read_xls = data
Err_Exit:
90        owkb_current.Activate
100       Set owkb_current = Nothing
110       Set owkb = Nothing
120       Exit Function

Err_Handler:
          'Note : Log Error
130       cErr.log 3, errt_datasource, Err.Number & "(" & Erl & ");" & Err.Description & ";VBAProject;c_excel_data (Module de classe);read_xls (Function))"
140       GoTo Err_Exit
End Function




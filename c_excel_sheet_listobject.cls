'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_excel_sheet_listobject (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 26/05/2016
' Purpose   : excel sheet list object related functions
'
'///////////////////////////////////////////////////////////////////////////////////////
'

Option Explicit
Option Base 1

Private m_cllistobject As listobject
Private m_oclErr As c_error



'********************************
'********* INIT / TERM **********
'********************************

Private Sub Class_Initialize()
Debug.Print "Load Module de classe : c_excel_sheet_listobject"

End Sub

Private Sub Class_Terminate()
Debug.Print "Unload Module de classe : c_excel_sheet_listobject"

    Set m_cllistobject = Nothing
    Set m_oclErr = Nothing
End Sub


'********************************
'********** SUB CLASS ***********
'********************************


Public Property Get clErr() As c_error
    Set clErr = m_oclErr
End Property

Public Property Set clErr(oclErr As c_error)
    Set m_oclErr = oclErr
End Property




'********************************
'********* PROPERTIES ***********
'********************************

Public Property Get listobject() As listobject
    Set listobject = m_cllistobject
End Property

Public Property Set listobject(cllistobject As listobject)
    Set m_cllistobject = cllistobject
End Property


'********************************
'********** FUNCTIONS ***********
'********************************

Public Function delete_()
With listobject
    .delete
End With
End Function



Public Function sort_with_lists(ParamArray olists())
      '
      '=======================================================================================
      ' Procedure : sort_with_lists (Function)
      ' Module    : c_excel_sheet_listobject (Module de classe)
      ' Project   : VBAProject
      ' Author    : MISTERNEO
      ' Date      : 24/06/2016
      ' Comments  : Multi list object sort
      ' Unit Test : (MISTERNEO) 24/06/2016 15:29 | Description [OK]
      ' Arg./i    :
      '           - [olists] Lists of sorting with column index (sortinglist1,index1, sortinglist2,index2)
      '           -
      '           -
      ' Arg./o    : Variant(v)
      '
      'Changes--------------------------------------------------------------------------------
      'Date               Programmer                      Change
      '24/06/2016         MISTERNEO               Initiate
      '
      '=======================================================================================
      '
      Dim oApp As Application
      Dim orng As Range, orng_dt As Range
      Dim i As Integer

10    On Error GoTo Err_Handler

20    With listobject
          'Note : clear previous filter
30        .Sort.SortFields.Clear
          'Note : custom sorting list
40        For i = LBound(olists) To UBound(olists) Step 2
50            Set orng = .ListColumns(olists(i + 1)).Range
60            .Sort.SortFields.Add Key:=orng, SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=Join(olists(i), ","), DataOption:=xlSortNormal
70        Next i
          
          'Note : specific standard sort list
80        Set orng_dt = .ListColumns(1).Range
90        .Sort.SortFields.Add Key:=orng_dt, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
100       Set orng_dt = .ListColumns(2).Range
110       .Sort.SortFields.Add Key:=orng_dt, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
          
120       With .Sort
130           .Header = xlYes
140           .MatchCase = False
150           .Orientation = xlTopToBottom
160           .SortMethod = xlPinYin
170           .Apply
180       End With
          
190   End With


Err_Exit:
200       Set orng = Nothing
210       Exit Function

Err_Handler:
          'Note : Log Error
220       clErr.log 1, errt_process, Err.Number & "(" & Erl & ");" & Err.Description & ";VBAProject;c_excel_sheet_listobject (Module de classe);sort_with_list (Function))"
230       GoTo Err_Exit

End Function

Public Function sort_with_list(oList As Variant, ByVal field_index As Integer)
'
'=======================================================================================
' Procedure : sort_with_list (Function)
' Module    : c_excel_sheet_listobject (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 24/06/2016
' Comments  : sort listobject (standard)
' Unit Test : (MISTERNEO) 24/06/2016 15:31 | Description [OK]
' Arg./i    :
'           - [olists] Lists of sorting
'           - [Field index] column index
'           -
' Arg./o    : Variant(v)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'24/06/2016         MISTERNEO               Initiate
'
'=======================================================================================
'
      Dim oApp As Application
      Dim orng As Range

10    On Error GoTo Err_Handler

20    With listobject
30       Set orng = .ListColumns(field_index).Range
40        .Sort.SortFields.Clear
50        .Sort.SortFields.Add Key:=orng, SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=Join(oList, ","), DataOption:=xlSortNormal
60        With .Sort
70            .Header = xlYes
80            .MatchCase = False
90            .Orientation = xlTopToBottom
100           .SortMethod = xlPinYin
110           .Apply
120       End With
          
130   End With


Err_Exit:
140       Set orng = Nothing
150       Exit Function

Err_Handler:

          'Note : Log Error
160       clErr.log 1, errt_process, Err.Number & "(" & Erl & ");" & Err.Description & ";VBAProject;c_excel_sheet_listobject (Module de classe);sort_with_list (Function))"
170       GoTo Err_Exit

End Function


Public Function erase_rows()

With listobject
    If .ListRows.count > 0 Then .DataBodyRange.delete
End With
End Function

Public Function erase_row(row_to_delete As Long) As String
With listobject
    .ListRows(row_to_delete).delete
End With
End Function

Public Function row_add(oVariant As Variant) As String
Dim oRow As ListRow

With listobject
    Set oRow = .ListRows.Add
    oRow.Range = oVariant
End With

End Function

Public Function rows_add(oDictionary As Scripting.Dictionary) As String
Dim oRow As ListRow
Dim uKey

With listobject
    For Each uKey In oDictionary.Keys
        Set oRow = .ListRows.Add
        oRow.Range = oDictionary.Items(uKey)
    Next
End With
End Function
Public Function count(Optional ByVal RC As String = "R") As Variant
    With listobject
        Select Case RC
            Case "R"
                count = .ListRows.count
            Case "C"
                count = .ListColumns.count
            Case Else
                count = 0
        End Select
    End With
End Function

Public Function create(oWorksheet As Worksheet, ByVal sName As String, oCell() As Variant, oHeader() As String, Optional theme As Variant)
      '
      '=======================================================================================
      ' Procedure : create (Function)
      ' Module    : c_excel_sheet_listobject (Module de classe)
      ' Project   : VBAProject
      ' Author    : MISTERNEO
      ' Date      : 24/06/2016
      ' Comments  : create listobject
      ' Unit Test : (MISTERNEO) 24/06/2016 15:32 | Description [OK]
      ' Arg./i    :
      '           - [oWorksheet] target worksheet
      '           - [sname] listobject name
      '           - [oCell] target range (start)
      '           - [oHeader] target range (start)
      '           -
      ' Arg./o    : Variant(v)
      '
      'Changes--------------------------------------------------------------------------------
      'Date               Programmer                      Change
      '24/06/2016         MISTERNEO               Initiate
      '
      '=======================================================================================
      '
      Dim myRng As Range
      Dim myCell As Range
      Dim oStyle As TableStyle

10    On Error GoTo Err_Handler

20    With oWorksheet
30        Set myCell = .Cells(oCell(1), oCell(2))
40        myCell.Select
50        Set myRng = .Range(Cells(oCell(1), oCell(2)), Cells(oCell(1), UBound(oHeader)))
60        myRng.Select
70        myRng = oHeader
80        Set listobject = .ListObjects.Add(xlSrcRange, myRng, , xlYes, myCell)
90        With listobject
100           .Name = sName
110           If IsMissing(theme) Then theme = "TableStyleMedium2"
120           On Error Resume Next
130           .TableStyle = theme
140       End With
150   End With

160   create = True

Err_Exit:
170       Set myCell = Nothing
180       Set listobject = Nothing
190       Exit Function

Err_Handler:
          'Note : Log Error
200       clErr.log 3, errt_process, Err.Number & "(" & Erl & ");" & Err.Description & ";VBAProject;c_excel_sheet_listobject (Module de classe);create (Function))"
210       GoTo Err_Exit
End Function


Public Function read() As Variant
'
'=======================================================================================
' Procedure : read (Function)
' Module    : c_excel_sheet_listobject (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 24/06/2016
' Comments  : read list object
' Unit Test : (MISTERNEO) 24/06/2016 15:36 | Description [OK]
' Arg./i    :
'           - [NO PARAM]
'           -
'           -
' Arg./o    : Variant(v)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'24/06/2016         MISTERNEO               Initiate
'
'=======================================================================================
'
10    On Error GoTo Err_Handler

      Dim oDatas() As Variant

20    With listobject
30        oDatas = .DataBodyRange
40    End With

50    read = oDatas

Err_Exit:
60        Exit Function

Err_Handler:
          'Note : Log Error
70        clErr.log 1, errt_datasource, Err.Number & "(" & Erl & ");" & Err.Description & ";VBAProject;c_excel_sheet_listobject (Module de classe);read (Function))"
80        GoTo Err_Exit

End Function


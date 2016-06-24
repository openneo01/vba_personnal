'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_excel_xml (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 24/06/2016
' Purpose   : ms excel xml related functions
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1

'********************************
'********* INIT / TERM **********
'********************************

Private Sub Class_Initialize()
Debug.Print "Load Module de classe : c_excel_xml"

End Sub

Private Sub Class_Terminate()
Debug.Print "Unload Module de classe : c_excel_xml"

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


Sub Create_XSD2()
'
'=======================================================================================
' Procedure : Create_XSD2 (Sub)
' Module    : c_excel_xml (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 24/06/2016
' Comments  : Create xsd
' Unit Test : (MISTERNEO) 24/06/2016 15:38 | Description [OK]
' Arg./i    :
'           - [NO PARAM]
'           -
'           -
' Arg./o    : ()
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'24/06/2016         MISTERNEO               Initiate
'
'=======================================================================================
'
   Dim StrMyXml As String, MyMap As XmlMap
   Dim StrMySchema As String

   StrMyXml = ThisWorkbook.Path & "\lo_xsd.xml"

   ' Turn off async loading.
   Application.DisplayAlerts = False
   ' Add the string to the XmlMaps collection.
   Set MyMap = ThisWorkbook.XmlMaps.Add(StrMyXml)
   Application.DisplayAlerts = True

   ' Create an empty file and output the schema.
   StrMySchema = ThisWorkbook.XmlMaps(1).Schemas(1).XML
   Open ThisWorkbook.Path & "\export_xsd.xsd" For Output As #1
   Print #1, StrMySchema
   Close #1
End Sub



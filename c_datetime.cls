'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_datetime (Module de classe)
' Project   : VBAProject
' Author    : misterneo
' Date      : 24/06/2016
' Purpose   : datetime functions
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1


Public Function IsoWeekNumber(InDate As Date) As Long
'
'=======================================================================================
' Procedure : IsoWeekNumber (Function)
' Module    : c_datetime (Module de classe)
' Project   : VBAProject
' Author    : GRT / GAZ
' Date      : 26/05/2016
' Comments  : This method uses the International Standards Organization (ISO) definition of a week number.
'           An ISO week always begins on a Monday, and ISO week 1 begins on the Monday of the first week of the year that contains a Thursday.
'           Said differently, ISO week 1 is the first week (beginning on Monday) that has at least four days.
' Unit Test : (GRT / GAZ) 26/05/2016 17:59 | Description [OK]
' Arg./i    :
'           - [InDate] Datetime to process
'           -
'           -
' Arg./o    : Long(l)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'26/05/2016         GRT / GAZ               Initiate
'
'=======================================================================================
'
    IsoWeekNumber = DatePart("ww", InDate, vbMonday, vbFirstFourDays)
End Function

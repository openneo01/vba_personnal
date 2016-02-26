Function test_missing_param(Optional dt As Double)
'
'=======================================================================================
' Procedure : test_missing_param (Function)
' Module    : Module1 (Module)
' Project   :
' Author    : GRT / GAZ
' Date      : 26/02/2016
' Comments  : **MANDATORY**
' Unit Test : (MISTERNEO) 26/02/2016 10:00 | Description [OK]
' Arg./i    :
'           - [dt] Any value (WARNING: optional parameter should be variant !! otherwise return 0
'           -
'           -
' Arg./o    : Variant(v)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'26/02/2016         MISTERNEO               Initiate
'
'=======================================================================================
'
If IsMissing(dt) Then MsgBox "Missing param" Else MsgBox dt

End Function

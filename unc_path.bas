Function Get_Local_Path_From_Unc(UncPath As String) As String '
'=======================================================================================
' Procedure : Get_Local_Path_From_Unc (Function)
' Module    : Module1 (Module)
' Project   : VBAProject
' Author    : misterneo
' Date      : 07/10/2015
' Comments  : retrieve Local Mapped Letter drive from Unc Path ' Unit Test : 07/10/2015 17:34 | Description [NOK]
' Arg./i    :
'           - [UncPath] : UNC Path
'           -
'           -
' Arg./o    : String(s)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'07/10/2015         misterneo               Initiate
'
'=======================================================================================
'
On Error GoTo Err_Handler

Dim objWMI As Object
Dim disks As Object
Dim disk As Object

Set objWMI = GetWMIService
Set disks = objWMI.ExecQuery("Select * from Win32_MappedLogicalDisk")

For Each disk In disks
    If UncPath Like disk.ProviderName & "*" Then
        'Note : Return Function value
        Get_Local_Path_From_Unc = Replace(UncPath, disk.ProviderName, disk.Name)
        Exit Function
    End If
Next disk

Err_Exit:
    'Note : Delete object
    Set disks = Nothing
    Set objWMI = Nothing
    'Note : Exit
    Exit Function

Err_Handler:
    'Note : Return Function value if error
    Get_Local_Path_From_Unc = UncPath
    'Note : Exit Function
    GoTo Err_Exit
 
End Function
 
Function GetWMIService() As Object
' http://msdn.microsoft.com/en-us/library/aa394586(VS.85).aspx
Dim strComputer As String
  strComputer = "."
  Set GetWMIService = GetObject("winmgmts:" _
                              & "{impersonationLevel=impersonate}!\\" _
                              & strComputer & "\root\cimv2") 
End Function




'
'///////////////////////////////////////////////////////////////////////////////////////
' Module    : c_outlook (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 24/05/2016
' Purpose   : **MANDATORY**
'
'///////////////////////////////////////////////////////////////////////////////////////
'
Option Explicit
Option Base 1

Private m_oclErr As c_Error

Private m_cloApp As Outlook.Application
Private m_clNamespace As Outlook.Namespace
Private m_clRecipient As Outlook.Recipient

Private m_oclOutlook_itm_mail As c_outlook_itm_mail
Private m_oclOutlook_itm_task As c_outlook_itm_task


'********************************
'********* INIT / TERM **********
'********************************

Private Sub Class_Initialize()
Debug.Print "Load Module de classe : c_outlook"

Set m_oclOutlook_itm_mail = New c_outlook_itm_mail
Set m_oclOutlook_itm_task = New c_outlook_itm_task

load_outlook

End Sub

Private Sub Class_Terminate()
Debug.Print "Unload Module de classe : c_outlook"

    Set m_oclErr = Nothing
    Set m_oclOutlook_itm_mail = Nothing
    Set m_oclOutlook_itm_task = Nothing
    Set m_cloApp = Nothing
    Set m_clNamespace = Nothing
    Set m_clRecipient = Nothing
End Sub



'********************************
'********** SUB CLASS ***********
'********************************
Public Property Get clErr() As c_Error
    Set clErr = m_oclErr
End Property

Public Property Set clErr(oclErr As c_Error)
    Set m_oclErr = oclErr
    Set m_oclOutlook_itm_mail.clErr = oclErr
    Set m_oclOutlook_itm_task.clErr = oclErr
End Property

Public Property Get clOutlook_itm_mail() As c_outlook_itm_mail
    Set clOutlook_itm_mail = m_oclOutlook_itm_mail
End Property

Public Property Get clOutlook_itm_task() As c_outlook_itm_task
    Set clOutlook_itm_task = m_oclOutlook_itm_task
End Property

'********************************
'********* PROPERTIES ***********
'********************************

Public Property Get oApp() As Outlook.Application
    Set oApp = m_cloApp
End Property

Public Property Set oApp(cloApp As Outlook.Application)
    Set m_cloApp = cloApp
End Property


Public Property Get Namespace() As Outlook.Namespace
    Set Namespace = m_clNamespace
End Property

Public Property Set Namespace(clNamespace As Outlook.Namespace)
    Set m_clNamespace = clNamespace
End Property

Public Property Get Recipient() As Outlook.Recipient
    Set Recipient = m_clRecipient
End Property

Public Property Set Recipient(clRecipient As Outlook.Recipient)
    Set m_clRecipient = clRecipient
End Property

'********************************
'********** FUNCTIONS ***********
'********************************

Public Function load_outlook() As Boolean
'
'=======================================================================================
' Procedure : load_outlook (Function)
' Module    : c_outlook (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 28/06/2016
' Comments  : test if oulook is installed and loaded
' Unit Test : (MISTERNEO) 28/06/2016 10:23 | Description [OK]
' Arg./i    :
'           - [NO PARAM]
'           -
'           -
' Arg./o    : Boolean(b)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'28/06/2016         GRT / GAZ               Initiate
'
'=======================================================================================
'
10    On Error GoTo Err_Handler

      Dim oApp As Object

20    Set oApp = CreateObject("Outlook.Application")
30    If oApp Is Nothing Then GoTo Err_Handler

40    load_outlook = True

Err_Exit:
60        Exit Function

Err_Handler:
70        load_outlook = False
80        clErr.log 3, etype_object, Err.Number & "(" & Erl & ");" & Err.Description & ";VBAProject;c_outlook (Module de classe);load_outlook (Function))"
90        GoTo Err_Exit
End Function


Public Function set_recipient(ByVal recipient_current As String)
'
'=======================================================================================
' Procedure : set_recipient (Function)
' Module    : c_outlook (Module de classe)
' Project   : VBAProject
' Author    : MISTERNEO
' Date      : 28/06/2016
' Comments  : set recipient (mailbox)
' Unit Test : (MISTERNEO) 28/06/2016 10:24 | Description [OK]
' Arg./i    :
'           - [recipient_current] email
'           -
'           -
' Arg./o    : Variant(v)
'
'Changes--------------------------------------------------------------------------------
'Date               Programmer                      Change
'28/06/2016         GRT / GAZ               Initiate
'
'=======================================================================================
'

With oApp
    Set Namespace = oApp.GetNamespace("MAPI")
    Set Recipient = Namespace.CreateRecipient(recipient_current)
End With
End Function








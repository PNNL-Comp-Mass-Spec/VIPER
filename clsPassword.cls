VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CheckPassword(strInfo As String, strPassword As String) As Boolean
    
    Dim objPWForm As frmPassword
    
    Set objPWForm = New frmPassword
    
    objPWForm.Initialize strInfo, strPassword
    
    objPWForm.Show
    Do
        DoEvents
        Sleep 50
    Loop While Not objPWForm.ProceedAndCloseForm()

    CheckPassword = objPWForm.PasswordWasValidated()

    Set objPWForm = Nothing
    
End Function

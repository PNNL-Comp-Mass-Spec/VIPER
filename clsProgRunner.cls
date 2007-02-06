VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class will run an application and monitor its progress


' The WaitForSingleObject function returns when one of the following occurs:
' - The specified object is in the signaled state.
' - The time-out interval elapses.
'
' The dwMilliseconds parameter specifies the time-out interval, in milliseconds.
' The function returns if the interval elapses, even if the object's state is
' nonsignaled. If dwMilliseconds is zero, the function tests the object's state
' and returns immediately. If dwMilliseconds is INFINITE, the function's time-out
' interval never elapses.
'
' This example waits an INFINITE amount of time for the process to end. As a
' result this process will be frozen until the shelled process terminates. The
' down side is that if the shelled process hangs, so will this one.
'
' A better approach is to wait a specific amount of time. Once the time-out
' interval expires, test the return value. If it is WAIT_TIMEOUT, the process
' is still not signaled. Then you can either wait again or continue with your
' processing.
'
' DOS Applications:
' Waiting for a DOS application is tricky because the DOS window never goes
' away when the application is done. To get around this, prefix the app that
' you are shelling to with "command.com /c".
'
' For example: lPid = Shell("command.com /c " & txtApp.Text, vbNormalFocus)


Private Const SYNCHRONIZE As Long = &H100000
Private Const WAIT_TIME_MSEC As Long = 100
 
' The state of the specified object is signaled
Private Const WAIT_OBJECT_0 = 0
 
' Code indicating the time-out interval elapsed & the object's state is nonsignaled.
Private Const WAIT_TIMEOUT As Long = &H102
 
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Private Type LUID
    lowpart As Long
    highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type



Private mApplicationPath As String
Private mAppRunning As Boolean
Private mAppStartSuccess As Boolean         ' True if the application successfully starts
Private mAbortProcessing As Boolean

Private mAppHandle As Long
Private mAppProcessID As Long

Public Property Get AppRunning() As Integer
    UpdateStatus
    AppRunning = mAppRunning
End Property

Public Property Get AppStartSuccess() As Boolean
    AppStartSuccess = mAppStartSuccess
End Property

Public Function AbortProcessing() As Boolean
    ' Returns True if the process if successfully killed
    ' Returns false if mAppHandle is 0 or if the process cannot be killed
    
    Dim blnKilled As Boolean
    
    blnKilled = False
    If mAppHandle > 0 Then
        blnKilled = KillProcess(mAppProcessID, 0)
        
        CloseAppHandle
    End If
    
    AbortProcessing = blnKilled
    
End Function

Private Sub CloseAppHandle()
    On Error Resume Next
    
    If mAppHandle <> 0 Then
        CloseHandle (mAppHandle)
    End If
    
    mAppHandle = 0
    mAppProcessID = 0
End Sub

Private Sub ResetStatus(ByVal strApplicationPath As String)
    mApplicationPath = strApplicationPath
    mAppRunning = False
    mAppStartSuccess = False
    mAbortProcessing = False
    
    mAppHandle = 0
    mAppProcessID = 0
End Sub

Private Sub Class_Initialize()
    ResetStatus ""
End Sub

Function KillProcess(ByVal hProcessID As Long, Optional ByVal exitCode As Long) As Boolean
    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES
    
    Dim blnSuccess As Boolean
    blnSuccess = False
    
    If GetVersion() >= 0 Then
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
            If hToken Then CloseHandle hToken
            KillProcess = False
            Exit Function
        End If
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            If hToken Then CloseHandle hToken
            KillProcess = False
            Exit Function
        End If
        
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
            If hToken Then CloseHandle hToken
            KillProcess = False
            Exit Function
        End If
    End If
        
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If hProcess Then
        blnSuccess = (TerminateProcess(hProcess, exitCode) <> 0)
        CloseHandle hProcess
    End If
        
    If GetVersion() >= 0 Then
        tp.Attributes = 0
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
        
        If hToken Then CloseHandle hToken
    End If
    
    KillProcess = blnSuccess
End Function

Public Function StartProgram(ByVal strApplicationPath As String, ByVal strArguments As String, ByVal eWindowStyle As VbAppWinStyle) As Boolean
    ' Runs strApplicationPath using the Shell command
    ' Returns True if the application starts; false if it doesn't
    ' To monitor the application's status use properties ApplicationRunning and ApplicationStatus

    Dim fso As New FileSystemObject
    Dim strCommandLine As String

On Error GoTo StartProgramErrorHandler

    ResetStatus strApplicationPath

    If fso.FileExists(strApplicationPath) Then
        strCommandLine = strApplicationPath
        
        If InStr(strCommandLine, " ") > 0 Then
            If Left(strCommandLine, 1) <> """" Then
                strCommandLine = """" & strCommandLine & """"
            End If
        End If
        
        If Len(Trim(strArguments)) > 0 Then
            strCommandLine = strCommandLine & " " & strArguments
        End If
        
        mAppProcessID = Shell(strCommandLine, eWindowStyle)

        If mAppProcessID > 0 Then
            mAppStartSuccess = True
            mAppRunning = True

            ' Get a handle to the shelled process.
            mAppHandle = OpenProcess(SYNCHRONIZE, 0, mAppProcessID)
        End If
    End If

    Set fso = Nothing

    StartProgram = mAppStartSuccess
    Exit Function

StartProgramErrorHandler:
    Debug.Print "Error in StartProgram: " & Err.Description
    Debug.Assert False

    StartProgram = False

End Function

Private Sub UpdateStatus()
    Dim lngReturn As Long

On Error GoTo UpdateStatusErrorHandler

    If mAppHandle <> 0 Then
        lngReturn = WaitForSingleObject(mAppHandle, WAIT_TIME_MSEC)

        If lngReturn <> WAIT_TIMEOUT Then
            mAppRunning = False
            CloseAppHandle
        End If
    Else
        mAppRunning = False
    End If

    Exit Sub

UpdateStatusErrorHandler:
    Debug.Assert False
End Sub


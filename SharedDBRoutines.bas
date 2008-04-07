Attribute VB_Name = "modSharedDBRoutines"
Option Explicit

' Written by Matthew Monroe, PNNL
' Started January 2, 2003
'
' Routines for accessing databases
'
' Last modified June 25, 2003

' Constants
Public DB_CONNECTION_TIMEOUT As Long        ' Default 120
Private Const SECONDS_PER_DAY = 86400


Public Function EstablishConnection(cnnConnection As ADODB.Connection, strConnectionString As String, Optional blnUseFormProgress As Boolean = True, Optional lngTimeoutOverrideSeconds As Long = 0) As Boolean
    ' If lngTimeoutOverrideSeconds is > 0, then that value overrides .DBConnectionTimeoutSeconds
    
    Dim lngSecElapsed As Long
    Dim blnProgressShown As Boolean
    
On Error GoTo EstablishConnectionResume

    TraceLog 3, "EstablishConnection", "Set cnnConnection = New ADODB.Connection"

    Set cnnConnection = New ADODB.Connection
    
    If lngTimeoutOverrideSeconds > 0 Then
        DB_CONNECTION_TIMEOUT = lngTimeoutOverrideSeconds
    Else
        DB_CONNECTION_TIMEOUT = glbPreferencesExpanded.AutoAnalysisOptions.DBConnectionTimeoutSeconds
        If DB_CONNECTION_TIMEOUT = 0 Then DB_CONNECTION_TIMEOUT = 300
    End If
    
    TraceLog 3, "EstablishConnection", "Set Timeout = " & DB_CONNECTION_TIMEOUT
    cnnConnection.ConnectionTimeout = DB_CONNECTION_TIMEOUT
    
    TraceLog 3, "EstablishConnection", "cnnConnection.Open"
    cnnConnection.Open strConnectionString, , , adAsyncConnect
    
    Dim StartTime As Date
    StartTime = Now()
    With cnnConnection
        Do While .STATE <> adStateOpen And (Now() - StartTime) * SECONDS_PER_DAY < DB_CONNECTION_TIMEOUT
            Sleep 50
            lngSecElapsed = CLng((Now() - StartTime) * SECONDS_PER_DAY)
            If lngSecElapsed >= 1 And blnUseFormProgress Then
                If Not blnProgressShown Then
                    ' Show progress bar, since process is taking over 1 second
                    frmProgress.InitializeSubtask "Connecting to database", 0, DB_CONNECTION_TIMEOUT
                    blnProgressShown = True
                Else
                    ' Update progress bar
                    frmProgress.UpdateSubtaskProgressBar lngSecElapsed
                    If KeyPressAbortProcess > 1 Then
                        Exit Do
                    End If
                End If
            End If
            DoEvents
        Loop
    End With

EstablishConnectionResume:
    If Err.Number <> 0 Then
        TraceLog 10, "EstablishConnection", "Error occurred: " & Err.Description
    End If
    
    On Error Resume Next
    If cnnConnection.STATE = adStateOpen Then
        TraceLog 3, "EstablishConnection", "Connection established"
        EstablishConnection = True
        Exit Function
    Else
        TraceLog 10, "EstablishConnection", "Connection failed"
        If blnUseFormProgress Then frmProgress.UpdateCurrentSubTask "Connection failed; releasing memory." & vbCrLf & "(may take up to " & Trim(Str(DB_CONNECTION_TIMEOUT)) & " seconds)"
        Set cnnConnection = Nothing
    End If

    EstablishConnection = False

End Function
    
Public Function FixNull(vPossiblyNullValue As Variant) As String
    
    On Error Resume Next
    If VarType(vPossiblyNullValue) = 1 Then
        FixNull = ""
    Else
        FixNull = vPossiblyNullValue
    End If

End Function

Public Function FixNullInt(vPossiblyNullValue As Variant, Optional intValueToReturnIfNull As Integer = 0) As Integer
    
    On Error Resume Next
    If VarType(vPossiblyNullValue) = 1 Then
        FixNullInt = intValueToReturnIfNull
    Else
        FixNullInt = vPossiblyNullValue
    End If

End Function

Public Function FixNullLng(vPossiblyNullValue As Variant, Optional lngValueToReturnIfNull As Long = 0) As Long
    
    On Error Resume Next
    If VarType(vPossiblyNullValue) = 1 Then
        FixNullLng = lngValueToReturnIfNull
    Else
        FixNullLng = vPossiblyNullValue
    End If

End Function

Public Function FixNullDbl(vPossiblyNullValue As Variant, Optional dblValueToReturnIfNull As Double = 0#) As Double
    
    On Error Resume Next
    If VarType(vPossiblyNullValue) = 1 Then
        FixNullDbl = dblValueToReturnIfNull
    Else
        FixNullDbl = vPossiblyNullValue
    End If

End Function



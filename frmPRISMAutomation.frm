VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPRISMAutomation 
   Caption         =   "VIPER -- PRISM Automation"
   ClientHeight    =   6984
   ClientLeft      =   60
   ClientTop       =   372
   ClientWidth     =   6792
   LinkTopic       =   "Form1"
   ScaleHeight     =   6984
   ScaleWidth      =   6792
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOptions 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6495
      Begin VB.TextBox txtServerForPreferredMTDB 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   1260
         Width           =   1935
      End
      Begin VB.CheckBox chkShowDebugPrompts 
         Caption         =   "Show debug info"
         Height          =   615
         Left            =   5400
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtPriorityMax 
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Text            =   "5"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPriorityMin 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPreferredMTDB 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkExclusiveDatabaseProcessing 
         Caption         =   "Exclusively process this DB"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtQueryInterval 
         Height          =   310
         Left            =   1440
         TabIndex        =   3
         Text            =   "60"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblServerForPreferredMTDB 
         Caption         =   "Server for MTDB:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1280
         Width           =   1455
      End
      Begin VB.Label lblPriorityRangeToProcessCoupler 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Preferred MTDB:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   980
         Width           =   1455
      End
      Begin VB.Label lblPriorityRangeToProcess 
         Caption         =   "Priority range to process"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   1815
      End
      Begin VB.Label lblTimeToNextQuery 
         Caption         =   "Time to next query: 0 seconds"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   270
         Width           =   2900
      End
      Begin VB.Label lblQueryInterval 
         Caption         =   "Query Interval:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblQueryIntervalUnits 
         Caption         =   "seconds"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame fraControls 
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   6495
      Begin VB.CommandButton cmdQueryNow 
         Caption         =   "&Query Now"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkLockoutControls 
         Caption         =   "Lockout Controls"
         Height          =   495
         Left            =   5280
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPauseQuery 
         Caption         =   "&Pause Querying"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdExitAutomation 
         Caption         =   "E&xit Automation Mode ..."
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdExitProgram 
         Caption         =   "Exit Program ..."
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Timer tmrPRISMQueryDelay 
      Interval        =   1000
      Left            =   6240
      Top             =   0
   End
   Begin RichTextLib.RichTextBox rtbPRISMAutomationLog 
      Height          =   2535
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   6495
      _ExtentX        =   11451
      _ExtentY        =   4466
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPRISMAutomation.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblExplanation 
      Caption         =   "Explanation"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmPRISMAutomation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
    
Private Const TIMER_DELAY_MSEC = 1000
Private Const MAX_LOG_ENTRIES = 500              ' maximum number of log entries; a given entry can have multiple lines

Private Const PRISM_AUTO_ANALYSIS_LOGFILE_NAME = "PrismAutoAnalysis.Log"
Private Const PRISM_AUTO_ANALYSIS_LOGFILE_MAXSIZE_BYTES = 7943757     ' 7.5 MB, so that when decreasing by 33%, will end up being 5 MB
Private Const STOP_VIPER_TEXTFILE_NAME = "StopViper.txt"
Private Const QUERY_PRISM_TEXTFILE_NAME = "QueryPrism.txt"
Private Const RESTART_VIPER_TEXTFILE_NAME = "RestartViper.txt"

Private mPRISMQueryInterval As Long     ' In Seconds
Private mControlsLocked As Boolean

Private LastUpdateTime As Date
Private mPaused As Boolean
Private mViperLocalOnServer As Boolean
Private mDebug As Boolean
Private mInitiatedViaCommandLine As Boolean
Private mExitAutomationWhenIdle As Boolean

' mForcePRISMQueryNow is used to force a re-query of PRISM directly after a job has completed processing
Private mForcePRISMQueryNow As Boolean

Private mHistoryLog(MAX_LOG_ENTRIES) As String            ' 0-based array; Circular buffer holding log lines; start of buffer is at mHistoryLogStartIndex
Private mHistoryLogCount As Long
' Note: After mHistoryLog() gets filled, then mHistoryLogStartIndex starts getting incremented, and we start
'       overwriting the earlier entries in mHistoryLog
Private mHistoryLogStartIndex As Long
'

Private Sub AddToPrismAutoAnalysisLog(strTextToAdd As String)
    
    ' Note that strTextToAdd can have embedded Carriage Returns (vbCrLf)
    
    If mHistoryLogCount < MAX_LOG_ENTRIES Then
        mHistoryLog(mHistoryLogCount) = strTextToAdd
        mHistoryLogCount = mHistoryLogCount + 1
    Else
        mHistoryLog(mHistoryLogStartIndex) = strTextToAdd
        mHistoryLogStartIndex = mHistoryLogStartIndex + 1
        If mHistoryLogStartIndex >= MAX_LOG_ENTRIES Then mHistoryLogStartIndex = 0
    End If
    
    DisplayHistoryLog
    
    WriteLatestLogEntryToDisk strTextToAdd
End Sub

Private Sub CheckElapsedIterations()
    Static intCheckForControlFile As Integer
    Dim strControlFilePath As String
    
    Dim SecondsElapsed As Long
    
On Error GoTo CheckElapsedIterationsErrorHandler

    If LastUpdateTime = 0 Then
        LastUpdateTime = Now()
    End If
    
    If LastUpdateTime > Now() Then
        ' This will happen when user pressed Query Now while paused
        LastUpdateTime = Now() - mPRISMQueryInterval / 86400!
    End If
    
    SecondsElapsed = ((Now()) - LastUpdateTime) * 86400!
    
    If SecondsElapsed >= mPRISMQueryInterval Or mForcePRISMQueryNow Then
        mForcePRISMQueryNow = False
        
        ' Set intCheckForControlFile to 10 so that we check for STOP_VIPER_TEXTFILE_NAME after we finish processing this job
        intCheckForControlFile = 10
        
        LastUpdateTime = Now()
        If mPRISMQueryInterval < 10 Then
            mPRISMQueryInterval = 60
            txtQueryInterval = Trim(mPRISMQueryInterval)
        End If
        
        lblTimeToNextQuery = "Time to next query: 0 seconds"
        
        QueryPRISM
    Else
        lblTimeToNextQuery = "Time to next query: " & CStr(mPRISMQueryInterval - SecondsElapsed) & " seconds"
    End If
    
    If intCheckForControlFile >= 10 Then
        strControlFilePath = AppendToPath(App.Path, STOP_VIPER_TEXTFILE_NAME)
        If FileExists(strControlFilePath) Then
            AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > StopViper.Txt file was found; shutting down VIPER"
            
            RenameControlFile strControlFilePath
            
            ExitProgramQueryUser False, False
        End If
        
        strControlFilePath = AppendToPath(App.Path, QUERY_PRISM_TEXTFILE_NAME)
        If FileExists(strControlFilePath) Then
            mForcePRISMQueryNow = True
            
            RenameControlFile strControlFilePath
            
            CheckElapsedIterations
        End If
        intCheckForControlFile = 0
    Else
        intCheckForControlFile = intCheckForControlFile + 1
    End If
    
    Exit Sub
    
CheckElapsedIterationsErrorHandler:
    Debug.Assert False
    tmrPRISMQueryDelay.Enabled = True
    LastUpdateTime = Now()
    
End Sub

Private Sub CheckForRestartViperFile()
    ' Looks for file RESTART_VIPER_TEXTFILE_NAME
    ' If found, then renames to RESTART_VIPER_TEXTFILE_NAME.Done
    
    Dim strControlFilePath  As String
    
    On Error Resume Next
    
    strControlFilePath = AppendToPath(App.Path, RESTART_VIPER_TEXTFILE_NAME)
    
    If FileExists(strControlFilePath) Then
        ' Rename file RESTART_VIPER_TEXTFILE_NAME to RESTART_VIPER_TEXTFILE_NAME.Done
        RenameControlFile strControlFilePath
    End If
End Sub

Private Sub CreateControlFile(ByVal strControlFilePath As String, ByVal strRestartReason As String)
    
    Dim strOldControlFilePath As String
    Dim fso As FileSystemObject
    Dim tsControlFile As TextStream
    
    ' Creates file strControlFilePath, storing the current date and time in the file
    ' In addition, deletes file strControlFilePath.Done if it exists
    
    Set fso = New FileSystemObject
    
    On Error Resume Next
    
    strOldControlFilePath = strControlFilePath & ".Done"
    If fso.FileExists(strOldControlFilePath) Then
        fso.DeleteFile strOldControlFilePath, True
        Sleep 100
    End If
    
    Set tsControlFile = fso.CreateTextFile(strControlFilePath, True)
    If Len(strRestartReason) = 0 Then
        strRestartReason = "Unknown restart reason"
    End If
    
    tsControlFile.WriteLine ("Restart Viper: " & GetCurrentTimeStamp & " - " & strRestartReason)
    tsControlFile.Close
    Set tsControlFile = Nothing

    Set fso = Nothing

End Sub

Private Sub DisplayHistoryLog(Optional strLineDelimeter As String = vbCrLf)
    
    Const CUMULATIVE_CHUNK_SIZE = 500
    
    Dim lngFillStringCount As Long
    Dim lngSrcIndex As Long
    Dim FillStringArray() As String
    Dim FillStringCumulative As String
    Dim lngStartIndex As Long, lngEndIndex As Long
    Dim lngSelStartSaved As Long
    Dim blnPlaceCursorAtEnd As Boolean
    
On Error GoTo DisplayHistoryLogErrorHandler

    ReDim FillStringArray(CLng(mHistoryLogCount / CUMULATIVE_CHUNK_SIZE) + 5)
    
    lngStartIndex = mHistoryLogStartIndex
    lngEndIndex = mHistoryLogCount - 1
    
    lngFillStringCount = 0
    For lngSrcIndex = lngStartIndex To lngEndIndex
        If lngSrcIndex Mod CUMULATIVE_CHUNK_SIZE = 0 Then
            If lngSrcIndex <> lngStartIndex Then
                lngFillStringCount = lngFillStringCount + 1
            End If
        End If
        
        FillStringArray(lngFillStringCount) = FillStringArray(lngFillStringCount) & mHistoryLog(lngSrcIndex) & strLineDelimeter
    
    Next lngSrcIndex
    
    ' If the History Log Start Index is greater than 0, then now need to append the log entries from index 0 to index mHistoryLogStartIndex-1
    If mHistoryLogStartIndex > 0 Then
        lngStartIndex = 0
        lngEndIndex = mHistoryLogStartIndex - 1
        For lngSrcIndex = lngStartIndex To lngEndIndex
            If lngSrcIndex Mod CUMULATIVE_CHUNK_SIZE = 0 Then
                If lngSrcIndex <> lngStartIndex Then
                    lngFillStringCount = lngFillStringCount + 1
                End If
            End If
            
            FillStringArray(lngFillStringCount) = FillStringArray(lngFillStringCount) & mHistoryLog(lngSrcIndex) & strLineDelimeter
        
        Next lngSrcIndex
    End If
    
    For lngSrcIndex = 0 To lngFillStringCount
        FillStringCumulative = FillStringCumulative & FillStringArray(lngSrcIndex)
    Next lngSrcIndex
    
    With rtbPRISMAutomationLog
        lngSelStartSaved = .SelStart
        If lngSelStartSaved > Len(rtbPRISMAutomationLog.Text) - 1000 Then blnPlaceCursorAtEnd = True
        .Text = FillStringCumulative
        .SelStart = lngSelStartSaved
        If blnPlaceCursorAtEnd Then
            .SelStart = Len(rtbPRISMAutomationLog.Text)
        End If
    End With
    
    Exit Sub
    
DisplayHistoryLogErrorHandler:
    Debug.Assert False
    tmrPRISMQueryDelay.Enabled = True
    
End Sub

Public Sub ExitAutomationQueryUser()
    Dim eResponse As VbMsgBoxResult
    
    eResponse = MsgBox("Really exit Automated Mode?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Exit")
    If eResponse = vbYes Then
        If mInitiatedViaCommandLine Then
            AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > PRISM Automation stopped by user"
        End If
        
        tmrPRISMQueryDelay.Enabled = False
        MDIForm1.RestoreMenus
        MDIForm1.Show
        Unload Me
    End If
    
End Sub

Private Sub ExitProgramQueryUser(Optional blnQueryUser As Boolean = True, Optional blnCreateRestartViperFile As Boolean = False, Optional strRestartReason As String = "")
    Dim eResponse As VbMsgBoxResult
    
    If blnQueryUser Then
        eResponse = MsgBox("Really exit Automated Mode and close the program?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Exit")
    End If
    
    If eResponse = vbYes Or Not blnQueryUser Then
        If mInitiatedViaCommandLine And blnQueryUser Then
            AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > PRISM Automation stopped by user"
        End If
        
        If blnCreateRestartViperFile Then
            CreateControlFile RESTART_VIPER_TEXTFILE_NAME, strRestartReason
        End If
        
        tmrPRISMQueryDelay.Enabled = False
        MDIForm1.RestoreMenus
        MDIForm1.Show
        MDIForm1.UnloadPRISMAutomationForm
    End If
    
End Sub

Private Sub ForceQueryUpdateNow()
    mForcePRISMQueryNow = True
End Sub

Private Function GetCurrentTimeStamp() As String
    GetCurrentTimeStamp = Format(Now(), "yyyy.mm.dd Hh:Nn:Ss")
End Function

Public Sub InitiateFromCommandLine(ByVal blnExitAutomationWhenIdle As Boolean)
    AddToPrismAutoAnalysisLog vbCrLf & vbCrLf
    AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > Automated PRISM analysis started"
    mInitiatedViaCommandLine = True
    mExitAutomationWhenIdle = blnExitAutomationWhenIdle
    
    ToggleLockoutControls True
    ForceQueryUpdateNow
End Sub

Public Sub InitializeControls()
    With glbPreferencesExpanded.AutoQueryPRISMOptions
        SetPRISMQueryInterval .QueryIntervalSeconds
        
        txtPriorityMin = .MinimumPriorityToProcess
        txtPriorityMax = .MaximumPriorityToProcess
        
        txtPreferredMTDB = .PreferredDatabaseToProcess
        txtServerForPreferredMTDB = .ServerForPreferredDatabase
        
        SetCheckBox chkExclusiveDatabaseProcessing, .ExclusivelyUseThisDatabase
    End With
    
End Sub

Private Sub PositionControls()
    Const MIN_LOG_BOX_WIDTH = 4000
    Const MIN_LOG_BOX_HEIGHT = 2000
    
    Dim lngDesiredValue As Long
    
    With rtbPRISMAutomationLog
        .Top = fraOptions.Top + fraOptions.Height + 120
        .Left = fraOptions.Left
        
        lngDesiredValue = Me.width - .Left - 360
        If lngDesiredValue < MIN_LOG_BOX_WIDTH Then lngDesiredValue = MIN_LOG_BOX_WIDTH
        .width = lngDesiredValue
    
        lngDesiredValue = Me.Height - .Top - fraControls.Height - 760
        If lngDesiredValue < MIN_LOG_BOX_HEIGHT Then lngDesiredValue = MIN_LOG_BOX_HEIGHT
        .Height = lngDesiredValue
    
        fraControls.Top = .Top + .Height + 120
        fraControls.Left = .Left
    End With
    
    
End Sub

Private Sub PostLogEntryToDB(strLogEntryType As String, strLogMessage As String)
    ' Posts an entry to T_Log_Entries in PRISM_RPT
    
    Dim strConnectionString As String
    Dim strPostLogEntrySPName As String
    Dim strComputerName As String
    Dim strPostedBy As String
    
    Dim cnnConnection As ADODB.Connection
    Dim cmdPostLogEntry As New ADODB.Command
    Dim prmType As New ADODB.Parameter
    Dim prmMessage As New ADODB.Parameter
    Dim prmPostedBy As New ADODB.Parameter
    
On Error GoTo PostLogEntryToDBErrorHandler

    'Get the computer name
    strComputerName = String(255, Chr(0))
    GetComputerName strComputerName, 255
    strComputerName = Left(strComputerName, InStr(strComputerName, Chr(0)) - 1)
    If Len(strComputerName) = 0 Then strComputerName = "UnknownViperAnalysisComputer"


    With glbPreferencesExpanded.AutoQueryPRISMOptions
        strConnectionString = .ConnectionStringQueryDB
        strPostLogEntrySPName = .PostLogEntrySPName
    End With
    

    ' Call a stored procedure to post an entry to the log
    ' Establish the connection
    If Not EstablishConnection(cnnConnection, strConnectionString) Then
        AddToPrismAutoAnalysisLog "Unable to connect to PRISM to post an entry to the DB log: ConnectionString = " & strConnectionString
        AddToPrismAutoAnalysisLog "Will try the default connection string instead (" & PRISM_AUTOMATION_CONNECTION_STRING_DEFAULT & ")" & vbCrLf
        strConnectionString = PRISM_AUTOMATION_CONNECTION_STRING_DEFAULT
        If Not EstablishConnection(cnnConnection, strConnectionString) Then
            AddToPrismAutoAnalysisLog "Unable to connect to PRISM using the default connection string either" & vbCrLf
            Exit Sub
        End If
    End If

    If Len(strPostLogEntrySPName) = 0 Then strPostLogEntrySPName = PRISM_AUTOMATION_SP_POST_LOG_ENTRY_DEFAULT

    ' Initialize the SP parameters
    InitializeSPCommand cmdPostLogEntry, cnnConnection, strPostLogEntrySPName
    
    Set prmType = cmdPostLogEntry.CreateParameter("type", adVarChar, adParamInput, 50, strLogEntryType)
    cmdPostLogEntry.Parameters.Append prmType
    
    Set prmMessage = cmdPostLogEntry.CreateParameter("message", adVarChar, adParamInput, 500, strLogMessage)
    cmdPostLogEntry.Parameters.Append prmMessage
    
    strPostedBy = "VIPER: " & strComputerName
    If Len(strPostedBy) > 50 Then
        strPostedBy = Left(strPostedBy, 50)
    End If

    Set prmPostedBy = cmdPostLogEntry.CreateParameter("postedBy", adVarChar, adParamInput, 50, strPostedBy)
    cmdPostLogEntry.Parameters.Append prmPostedBy

    ' Execute the SP
    cmdPostLogEntry.Execute

    Set cmdPostLogEntry.ActiveConnection = Nothing
    Set cmdPostLogEntry = Nothing
 
    If cnnConnection.STATE <> adStateClosed Then cnnConnection.Close
    Set cnnConnection = Nothing

    Exit Sub

PostLogEntryToDBErrorHandler:
    LogErrors Err.Number, "frmPRISMAutomation->PostLogEntryToDB"
    Debug.Assert False

End Sub

Private Function ProcessJob(udtAutoParams As udtAutoAnalysisParametersType) As Boolean
    ' Returns True if success, False if not
    
    Dim blnSuccess As Boolean
    Dim strMemoryLog As String
    
    ' Call AutoAnalysisStart to do all the work
    blnSuccess = AutoAnalysisStart(udtAutoParams, True, mDebug)
    
    MDIForm1.Hide
    
    ' Add the new log lines to mHistoryLog
    strMemoryLog = Trim(AutoAnalysisMemoryLogGet())
    Do While Right(strMemoryLog, 2) = vbCrLf
        strMemoryLog = Left(strMemoryLog, Len(strMemoryLog) - 2)
    Loop
    
    AddToPrismAutoAnalysisLog strMemoryLog
    
    AddToPrismAutoAnalysisLog "--------------------"
    If blnSuccess Then
        AddToPrismAutoAnalysisLog "Analysis Successful"
    Else
        If Len(udtAutoParams.FilePaths.LogFilePath) > 0 Then
            AddToPrismAutoAnalysisLog "Error: Analysis Failed; see " & udtAutoParams.FilePaths.LogFilePath & " for details"
        Else
            AddToPrismAutoAnalysisLog "Error: Analysis Failed"
        End If
    End If
    AddToPrismAutoAnalysisLog "====================" & vbCrLf
    
    ProcessJob = blnSuccess
    
End Function

Private Sub QueryPRISM()
    Const MAX_RETRY_COUNT As Integer = 5
    
    Static blnWorking As Boolean
    
    Dim udtAutoParams As udtAutoAnalysisParametersType
    
    Dim cnnConnection As ADODB.Connection
    
    Dim strConnectionString As String
    
    Dim fso As FileSystemObject
    
    Dim lngAvailableJobID As Long
    Dim blnSuccess As Boolean
    Dim intCallCount  As Integer
    
    Dim intMinPriority As Integer, intMaxPriority As Integer
    Dim strPreferredDatabaseToProcess As String
    Dim strServerForPreferredDatabase As String
    Dim strComputerName As String
    Dim strAppFolderName As String
    
    Dim intClientPerspective As Integer                     ' 0 if VIPER is running locally on the Database server computer; 1 otherwise (and thus running remotely)
    Dim intExclusivelyProcessThisDatabase As Integer        ' 0 = False, 1 = True
    
    Dim strLastGoodLocation As String
    
    Dim strRequestTaskSPName As String
    Dim strSetTaskCompleteSPName As String
    Dim strLogEntryType As String
    
    Dim cmdGetPMTask As New ADODB.Command
    Dim prmProcessorName As New ADODB.Parameter
    Dim prmClientPerspective As New ADODB.Parameter
    Dim prmPriorityMin As New ADODB.Parameter
    Dim prmPriorityMax As New ADODB.Parameter
    Dim prmRestrictToMtdbName As New ADODB.Parameter
    Dim prmTaskID As New ADODB.Parameter
    Dim prmTaskPriority As New ADODB.Parameter
    Dim prmAnalysisJob As New ADODB.Parameter
    Dim prmAnalysisResultsFolderPath As New ADODB.Parameter
    Dim prmServerName As New ADODB.Parameter
    Dim prmMtdbName As New ADODB.Parameter
    Dim prmAMTsOnly As New ADODB.Parameter
    Dim prmConfirmedOnly As New ADODB.Parameter
    Dim prmLockersOnly As New ADODB.Parameter
    Dim prmLimitToPMTsFromDataset As New ADODB.Parameter
    
    Dim prmMTsubsetID As New ADODB.Parameter
    Dim prmModList As New ADODB.Parameter
    
    Dim prmMinimumHighNormalizedScore As New ADODB.Parameter
    Dim prmMinimumHighDiscriminantScore As New ADODB.Parameter
    Dim prmMinimumPMTQualityScore As New ADODB.Parameter
    Dim prmExperimentInclusionFilter As New ADODB.Parameter
    Dim prmExperimentExclusionFilter As New ADODB.Parameter
    Dim prmInternalStdExplicit As New ADODB.Parameter

    Dim prmNETValueType As New ADODB.Parameter
    Dim prmIniFilePath As New ADODB.Parameter
    Dim prmOutputFolderPath As New ADODB.Parameter
    Dim prmLogFilePath As New ADODB.Parameter
    Dim prmTaskAvailable As New ADODB.Parameter
    Dim prmMessage As New ADODB.Parameter
    Dim prmDBSchemaVersion As New ADODB.Parameter
    Dim prmToolVersion As New ADODB.Parameter
    Dim prmMinimumPeptideProphetProbability As New ADODB.Parameter
    
    Dim strMessage As String
    
    ' Do not allow recursive calls to this sub
    ' If a recursive call does occur, make sure the Timer is disabled
    If blnWorking Then
        tmrPRISMQueryDelay.Enabled = False
        Exit Sub
    End If
    
    blnWorking = True
    
On Error GoTo QueryPRISMErrorHandler

    tmrPRISMQueryDelay.Enabled = False
    Me.MousePointer = vbHourglass
    
    AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > Querying PRISM"
    
    strLastGoodLocation = "Get auto query options"
    If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    
    ' Get the current Auto Query options
    With glbPreferencesExpanded.AutoQueryPRISMOptions
        strConnectionString = .ConnectionStringQueryDB
        strRequestTaskSPName = .RequestTaskSPName
        
        intMinPriority = .MinimumPriorityToProcess
        If .MaximumPriorityToProcess <= 0 Then
            intMaxPriority = 200                    ' Set to 200, allowing for 201 to 255 to be used for debugging purposes
        Else
            intMaxPriority = .MaximumPriorityToProcess
        End If
        
        If intMaxPriority < intMinPriority Then
            If intMinPriority <= 0 Then
                intMinPriority = 0
                intMaxPriority = 1
            Else
                intMaxPriority = intMinPriority
            End If
            txtPriorityMin = intMinPriority
            txtPriorityMax = intMaxPriority
        End If
        
        If Len(.PreferredDatabaseToProcess) > 0 Or Len(.ServerForPreferredDatabase) > 0 Then
            strPreferredDatabaseToProcess = .PreferredDatabaseToProcess
            strServerForPreferredDatabase = .ServerForPreferredDatabase
        Else
            strPreferredDatabaseToProcess = ExtractDBNameFromConnectionString(CurrMTDatabase)
            strServerForPreferredDatabase = ""
        End If
        
        If .ExclusivelyUseThisDatabase Then
            intExclusivelyProcessThisDatabase = 1
        Else
            intExclusivelyProcessThisDatabase = 0
        End If
    End With
    
    If intExclusivelyProcessThisDatabase = 1 And (Len(strPreferredDatabaseToProcess) = 0 Or Len(strServerForPreferredDatabase) = 0) Then
        ' User specified to exclusively process a database, but no database is defined
        ' Thus, do not query the database
        AddToPrismAutoAnalysisLog "Error: Please specify both a database name and server name when 'Exclusively process this DB' is enabled" & vbCrLf
        
        blnWorking = False
        If Not mPaused Then tmrPRISMQueryDelay.Enabled = True
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    strLastGoodLocation = "Call Get Compute Name"
    If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    
    'Get the computer name
    strComputerName = String(255, Chr(0))
    GetComputerName strComputerName, 255
    strComputerName = Left(strComputerName, InStr(strComputerName, Chr(0)) - 1)
    If Len(strComputerName) = 0 Then strComputerName = "UnknownViperAnalysisComputer"
    
    ' Get the name of the folder containing Viper
    ' If the folder is not simply VIPER, then we'll append the extra text (or the full name) to the computer name
    ' This is done in case multiple copies of VIPER are running on the same computer.
    strAppFolderName = App.Path
    
    Set fso = New FileSystemObject
    strAppFolderName = fso.GetBaseName(strAppFolderName)
    If UCase(strAppFolderName) <> "VIPER" Then
        If UCase(Left(strAppFolderName, 5)) = "VIPER" Then
            strAppFolderName = Mid(strAppFolderName, 6)
        End If
        If Left(strAppFolderName, 1) = "_" Then strAppFolderName = Mid(strAppFolderName, 2)
        strComputerName = strComputerName & "_" & strAppFolderName
    End If
    
    If mViperLocalOnServer Then
        intClientPerspective = 0            ' Running on server
    Else
        intClientPerspective = 1            ' Running on client
    End If
    
    strLastGoodLocation = "Trying to connect to Prism"
    If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    
    ' Call a stored procedure to populate lngAvailableJobID; return -1 if none available
    ' Establish the connection
    If Not EstablishConnection(cnnConnection, strConnectionString) Then
        AddToPrismAutoAnalysisLog "Unable to connect to PRISM: ConnectionString = " & strConnectionString
        AddToPrismAutoAnalysisLog "Will try the default connection string instead (" & PRISM_AUTOMATION_CONNECTION_STRING_DEFAULT & ")" & vbCrLf
        strConnectionString = PRISM_AUTOMATION_CONNECTION_STRING_DEFAULT
        If Not EstablishConnection(cnnConnection, strConnectionString) Then
            AddToPrismAutoAnalysisLog "Unable to connect to PRISM using the default connection string either" & vbCrLf
            blnWorking = False
            If Not mPaused Then tmrPRISMQueryDelay.Enabled = True
            Me.MousePointer = vbDefault
            Set fso = Nothing
            Exit Sub
        End If
    End If

    strLastGoodLocation = "Connected to Prism"
 If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    
    If Len(strRequestTaskSPName) = 0 Then strRequestTaskSPName = PRISM_AUTOMATION_SP_REQUEST_TASK_DEFAULT

    strLastGoodLocation = "Set cmdGetPMTask.ActiveConnection"
    If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    
    ' Initialize the SP parameters
    InitializeSPCommand cmdGetPMTask, cnnConnection, strRequestTaskSPName
    
    strLastGoodLocation = "Set prmProcessorName = cmdGetPMTask.CreateParameter(" & Chr(34) & "processorName" & Chr(34) & ")"
 If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    Set prmProcessorName = cmdGetPMTask.CreateParameter("processorName", adVarChar, adParamInput, 128, strComputerName)
    
    strLastGoodLocation = "cmdGetPMTask.Parameters.Append prmProcessorName"
 If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    cmdGetPMTask.Parameters.Append prmProcessorName
    
    Set prmClientPerspective = cmdGetPMTask.CreateParameter("clientPerspective", adTinyInt, adParamInput, , intClientPerspective)
    cmdGetPMTask.Parameters.Append prmClientPerspective

    Set prmPriorityMin = cmdGetPMTask.CreateParameter("priorityMin", adTinyInt, adParamInput, , intMinPriority)
    cmdGetPMTask.Parameters.Append prmPriorityMin
    Set prmPriorityMax = cmdGetPMTask.CreateParameter("priorityMax", adTinyInt, adParamInput, , intMaxPriority)
    cmdGetPMTask.Parameters.Append prmPriorityMax

    Set prmRestrictToMtdbName = cmdGetPMTask.CreateParameter("restrictToMtdbName", adTinyInt, adParamInput, , intExclusivelyProcessThisDatabase)
    cmdGetPMTask.Parameters.Append prmRestrictToMtdbName

    Set prmTaskID = cmdGetPMTask.CreateParameter("taskID", adInteger, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmTaskID
    Set prmTaskPriority = cmdGetPMTask.CreateParameter("taskPriority", adTinyInt, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmTaskPriority

    Set prmAnalysisJob = cmdGetPMTask.CreateParameter("analysisJob", adInteger, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmAnalysisJob
    Set prmAnalysisResultsFolderPath = cmdGetPMTask.CreateParameter("analysisResultsFolderPath", adVarChar, adParamOutput, 256, "")
    cmdGetPMTask.Parameters.Append prmAnalysisResultsFolderPath
    
    Set prmServerName = cmdGetPMTask.CreateParameter("ServerName", adVarChar, adParamInputOutput, 128, strServerForPreferredDatabase)
    cmdGetPMTask.Parameters.Append prmServerName
    
    Set prmMtdbName = cmdGetPMTask.CreateParameter("mtdbName", adVarChar, adParamInputOutput, 128, strPreferredDatabaseToProcess)
    cmdGetPMTask.Parameters.Append prmMtdbName

    Set prmAMTsOnly = cmdGetPMTask.CreateParameter("amtsOnly", adTinyInt, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmAMTsOnly
    Set prmConfirmedOnly = cmdGetPMTask.CreateParameter("confirmedOnly", adTinyInt, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmConfirmedOnly
    Set prmLockersOnly = cmdGetPMTask.CreateParameter("lockersOnly", adTinyInt, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmLockersOnly
    Set prmLimitToPMTsFromDataset = cmdGetPMTask.CreateParameter("LimitToPMTsFromDataset", adTinyInt, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmLimitToPMTsFromDataset
   
    Set prmMTsubsetID = cmdGetPMTask.CreateParameter("mtsubsetID", adInteger, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmMTsubsetID

    Set prmModList = cmdGetPMTask.CreateParameter("modList", adVarChar, adParamOutput, 128, "")
    cmdGetPMTask.Parameters.Append prmModList
    
    Set prmMinimumHighNormalizedScore = cmdGetPMTask.CreateParameter("MinimumHighNormalizedScore", adSingle, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmMinimumHighNormalizedScore
    
    Set prmMinimumHighDiscriminantScore = cmdGetPMTask.CreateParameter("MinimumHighDiscriminantScore", adSingle, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmMinimumHighDiscriminantScore
    
    Set prmMinimumPMTQualityScore = cmdGetPMTask.CreateParameter("MinimumPMTQualityScore", adSingle, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmMinimumPMTQualityScore
    
    Set prmExperimentInclusionFilter = cmdGetPMTask.CreateParameter("ExperimentFilter", adVarChar, adParamOutput, 64, "")
    cmdGetPMTask.Parameters.Append prmExperimentInclusionFilter
    
    Set prmExperimentExclusionFilter = cmdGetPMTask.CreateParameter("ExperimentExclusionFilter", adVarChar, adParamOutput, 64, "")
    cmdGetPMTask.Parameters.Append prmExperimentExclusionFilter
    
    Set prmInternalStdExplicit = cmdGetPMTask.CreateParameter("InternalStdExplicit", adVarChar, adParamOutput, 255, "")
    cmdGetPMTask.Parameters.Append prmInternalStdExplicit
    
    Set prmNETValueType = cmdGetPMTask.CreateParameter("NETValueType", adTinyInt, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmNETValueType
    
    Set prmIniFilePath = cmdGetPMTask.CreateParameter("iniFilePath", adVarChar, adParamOutput, 255, "")
    cmdGetPMTask.Parameters.Append prmIniFilePath
    Set prmOutputFolderPath = cmdGetPMTask.CreateParameter("outputFolderPath", adVarChar, adParamOutput, 255, "")
    cmdGetPMTask.Parameters.Append prmOutputFolderPath
    Set prmLogFilePath = cmdGetPMTask.CreateParameter("logFilePath", adVarChar, adParamOutput, 255, "")
    cmdGetPMTask.Parameters.Append prmLogFilePath

    Set prmTaskAvailable = cmdGetPMTask.CreateParameter("taskAvailable", adTinyInt, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmTaskAvailable
    Set prmMessage = cmdGetPMTask.CreateParameter("message", adVarChar, adParamOutput, 512, "")
    cmdGetPMTask.Parameters.Append prmMessage

    Set prmDBSchemaVersion = cmdGetPMTask.CreateParameter("DBSchemaVersion", adSingle, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmDBSchemaVersion
    
    Set prmToolVersion = cmdGetPMTask.CreateParameter("toolVersion", adVarChar, adParamInput, 128, GetMyNameVersion(True, True))
    cmdGetPMTask.Parameters.Append prmToolVersion
    
    Set prmMinimumPeptideProphetProbability = cmdGetPMTask.CreateParameter("MinimumPeptideProphetProbability", adSingle, adParamOutput, , 0)
    cmdGetPMTask.Parameters.Append prmMinimumPeptideProphetProbability
    
    
    strLastGoodLocation = "Execute Stored Procedure: " & strRequestTaskSPName
    If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
    
    ' Execute the SP
    cmdGetPMTask.Execute

    If Not (IsNull(prmMessage.Value)) Then
        strMessage = CStr(prmMessage)
    Else
        strMessage = ""
    End If
    
    If (IsNull(prmTaskAvailable.Value)) Then
        strLastGoodLocation = "Execute done: No available jobs"
        If Len(strMessage) > 0 Then
            strLastGoodLocation = strLastGoodLocation & vbCrLf & "  -> " & strMessage
        End If
        
        If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation & vbCrLf
        lngAvailableJobID = -1
    ElseIf CLngSafe(prmTaskAvailable.Value) = 0 Then
        strLastGoodLocation = "Execute done: No available jobs"
        If Len(strMessage) > 0 Then
            strLastGoodLocation = strLastGoodLocation & vbCrLf & "  -> " & strMessage
        End If
        
        If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation & vbCrLf
        lngAvailableJobID = -1
    Else
        strLastGoodLocation = "Execute done: Job found"
        If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation
        
        InitializeAutoAnalysisParameters udtAutoParams
        With udtAutoParams
            With .FilePaths
                
                .ResultsFolder = fso.GetFileName(CStrSafe(prmAnalysisResultsFolderPath.Value))
                
                ' Parse out the Dataset Folder
                .DatasetFolder = fso.GetParentFolderName(CStrSafe(prmAnalysisResultsFolderPath.Value))
                .DatasetFolder = fso.GetFileName(.DatasetFolder)
                
                .InputFilePath = AppendToPath(CStrSafe(prmAnalysisResultsFolderPath.Value), "*")
            
                .OutputFolderPath = CStrSafe(prmOutputFolderPath.Value)
                .LogFilePath = CStrSafe(prmLogFilePath.Value)
                
                .IniFilePath = Trim(CStrSafe(prmIniFilePath.Value))
                
                ' .IniFilePath sometimes contains 1 or more carriage returns
                ' Check for and remove them
                If InStr(.IniFilePath, vbCrLf) > 0 Then
                    .IniFilePath = Replace(.IniFilePath, vbCrLf, "")
                End If
            End With
            
            .FullyAutomatedPRISMMode = True
            
            .JobNumber = CLngSafe(prmAnalysisJob.Value)
            
            With .MTDBOverride
                .Enabled = True
                .DBSchemaVersion = prmDBSchemaVersion.Value
                If .DBSchemaVersion < 1 Then .DBSchemaVersion = 1

                .ServerName = CStrSafe(prmServerName.Value)
                .MTDBName = CStrSafe(prmMtdbName.Value)
                
                .ConnectionString = ConstructConnectionString(.ServerName, .MTDBName, strConnectionString)
                .AMTsOnly = CBoolSafe(prmAMTsOnly.Value)
                .ConfirmedOnly = CBoolSafe(prmConfirmedOnly.Value)
                .LockersOnly = CBoolSafe(prmLockersOnly.Value)
                .LimitToPMTsFromDataset = CBoolSafe(prmLimitToPMTsFromDataset.Value)
                
                .MinimumHighNormalizedScore = CSngSafe(prmMinimumHighNormalizedScore.Value)
                .MinimumHighDiscriminantScore = CSngSafe(prmMinimumHighDiscriminantScore.Value)
                .MinimumPeptideProphetProbability = CSngSafe(prmMinimumPeptideProphetProbability.Value)
                .MinimumPMTQualityScore = CSngSafe(prmMinimumPMTQualityScore.Value)
                
                .ExperimentInclusionFilter = CStrSafe(prmExperimentInclusionFilter.Value)
                .ExperimentExclusionFilter = CStrSafe(prmExperimentExclusionFilter.Value)
                .InternalStandardExplicit = CStrSafe(prmInternalStdExplicit.Value)
                
                .NETValueType = CIntSafe(prmNETValueType.Value)
                .MTSubsetID = CLngSafe(prmMTsubsetID.Value)
                .ModList = Trim(CStrSafe(prmModList.Value))
                
                ' A blank mod list is signified by "-1" in schema version 1
                ' Use "" in higher schema versions
                If .ModList = "" Then
                    If .DBSchemaVersion < 2 Then .ModList = "-1"
                End If
                
                ' Also check for a mod list of "none"
                ' Change this to -1 if schema version is < 2, otherwise change to ""
                If LCase(.ModList) = "none" Then
                    If .DBSchemaVersion < 2 Then .ModList = "-1" Else .ModList = ""
                End If
                
                .PeakMatchingTaskID = CLngSafe(prmTaskID.Value)
            End With
            
            lngAvailableJobID = udtAutoParams.JobNumber
            
            .AutoCloseFileWhenDone = True
            
            .ComputerName = strComputerName
        End With
        
    End If
    
    Set fso = Nothing
    
    Set cmdGetPMTask.ActiveConnection = Nothing
    Set cmdGetPMTask = Nothing
    
    If lngAvailableJobID >= 0 Then
        AddToPrismAutoAnalysisLog vbCrLf & "===================="
        AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > Processing Job " & Trim(lngAvailableJobID)

        strLastGoodLocation = "Process job " & Trim(lngAvailableJobID)
        If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation

        ' Process the job
        blnSuccess = ProcessJob(udtAutoParams)

        If Not blnSuccess And udtAutoParams.ErrorBits = 0 Then
            udtAutoParams.ErrorBits = -1     ' Unknown Error
        End If

        ' Call SetPeakMatchingTaskComplete to mark the task as complete (or SetPeakMatchingTaskToRestart to reset it)
        ' Network errors sometimes cause this call to fail, so we'll attempt the call at most 5 times
        intCallCount = 0
        Do
            If udtAutoParams.ExitViperASAP And udtAutoParams.RestartAfterExit Then
                strSetTaskCompleteSPName = glbPreferencesExpanded.AutoQueryPRISMOptions.SetTaskToRestartSPName
            Else
                strSetTaskCompleteSPName = glbPreferencesExpanded.AutoQueryPRISMOptions.SetTaskCompleteSPName
            End If
            
            strLastGoodLocation = "Execute stored procedure: " & strSetTaskCompleteSPName & " (intCallCount = " & CStr(intCallCount) & ")"
            If mDebug Then AddToPrismAutoAnalysisLog strLastGoodLocation

            If udtAutoParams.ExitViperASAP And udtAutoParams.RestartAfterExit Then
                blnSuccess = SetPeakMatchingTaskToRestart(cnnConnection, udtAutoParams)
            Else
                blnSuccess = SetPeakMatchingTaskComplete(cnnConnection, udtAutoParams)
            End If
            
            If Not blnSuccess Then
                ' Sleep for 10 seconds
                Sleep 10000
                
                ' Just to be safe, re-establish the connection (it's often in a closed state if this error occurs)
                If cnnConnection.STATE <> adStateClosed Then cnnConnection.Close
                If Not EstablishConnection(cnnConnection, strConnectionString) Then
                    AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > Error re-establishing connection after failed call to " & strSetTaskCompleteSPName
                End If
            End If

            intCallCount = intCallCount + 1
        Loop While Not blnSuccess And intCallCount < MAX_RETRY_COUNT

        If Not blnSuccess Then
            AddToPrismAutoAnalysisLog "Error: Call to SetPeakMatchingTaskComplete failed " & CStr(MAX_RETRY_COUNT) & " times; aborting" & vbCrLf
        Else
            If intCallCount > 1 Then
                AddToPrismAutoAnalysisLog "Call to SetPeakMatchingTaskComplete was successful" & vbCrLf
            End If
            
            If Not udtAutoParams.ExitViperASAP Then
                mForcePRISMQueryNow = True
            End If
        End If
    End If
    
QueryPRISMCleanup:
    On Error Resume Next

    If cnnConnection.STATE <> adStateClosed Then cnnConnection.Close
    Set cnnConnection = Nothing

    blnWorking = False
    
    If udtAutoParams.ExitViperASAP Then
        ' A fatal error has occurred; exit Viper
        AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > Fatal error (" & udtAutoParams.ExitViperReason & "); shutting down VIPER"
        
        If InStr(udtAutoParams.ExitViperReason, "Failed to load control 'ctl2DHeatMap'") > 0 Then
            strLogEntryType = "ErrorIgnore"
        Else
            strLogEntryType = "Error"
        End If
        PostLogEntryToDB strLogEntryType, "Fatal error (" & udtAutoParams.ExitViperReason & "); shutting down VIPER"
        
        ExitProgramQueryUser False, udtAutoParams.RestartAfterExit, udtAutoParams.ExitViperReason
    Else
        If Not mPaused Then
            If mExitAutomationWhenIdle And lngAvailableJobID < 0 Then
                ExitProgramQueryUser False, False, "No available tasks"
            End If
            
            tmrPRISMQueryDelay.Enabled = True
        End If
    End If
    
    Me.MousePointer = vbDefault
    Exit Sub
    
QueryPRISMErrorHandler:
    Debug.Print Err.Description
    AddToPrismAutoAnalysisLog "Error while querying PRISM (Last Good Location = " & strLastGoodLocation & ": " & Err.Description & vbCrLf
    
    Resume QueryPRISMCleanup
    
End Sub

Private Function SetPeakMatchingTaskComplete(cnnConnection As ADODB.Connection, udtAutoParams As udtAutoAnalysisParametersType) As Boolean
    ' Call SetPeakMatchingTaskCompleteMaster SP to mark task as complete
    ' Returns True if successfully called, False if an error occurs
    
    Dim strSetPMTaskComplete As String
    
    Dim cmdSetPMTaskComplete As New ADODB.Command
        
    Dim prmTaskID As New ADODB.Parameter
    Dim prmServerName As New ADODB.Parameter
    Dim prmMtdbName As New ADODB.Parameter
    Dim prmMessage As New ADODB.Parameter
    
    Dim prmErrorCode As New ADODB.Parameter
    Dim prmWarningCode As New ADODB.Parameter
    Dim prmMDID As New ADODB.Parameter

On Error GoTo SetPeakMatchingTaskCompleteErrorHandler
        
    With glbPreferencesExpanded.AutoQueryPRISMOptions
        strSetPMTaskComplete = .SetTaskCompleteSPName
    End With
    
    If Len(strSetPMTaskComplete) = 0 Then strSetPMTaskComplete = PRISM_AUTOMATION_SP_SET_COMPLETE_DEFAULT
    
    ' Initialize the SP parameters
    InitializeSPCommand cmdSetPMTaskComplete, cnnConnection, strSetPMTaskComplete
    
    Set prmTaskID = Nothing
    Set prmTaskID = cmdSetPMTaskComplete.CreateParameter("taskID", adInteger, adParamInput, , udtAutoParams.MTDBOverride.PeakMatchingTaskID)
    cmdSetPMTaskComplete.Parameters.Append prmTaskID
        
    Set prmServerName = Nothing
    Set prmServerName = cmdSetPMTaskComplete.CreateParameter("ServerName", adVarChar, adParamInput, 128, udtAutoParams.MTDBOverride.ServerName)
    cmdSetPMTaskComplete.Parameters.Append prmServerName
    
    Set prmMtdbName = Nothing
    Set prmMtdbName = cmdSetPMTaskComplete.CreateParameter("mtdbName", adVarChar, adParamInput, 128, udtAutoParams.MTDBOverride.MTDBName)
    cmdSetPMTaskComplete.Parameters.Append prmMtdbName
    
    Set prmErrorCode = cmdSetPMTaskComplete.CreateParameter("errorCode", adInteger, adParamInput, , udtAutoParams.ErrorBits)
    cmdSetPMTaskComplete.Parameters.Append prmErrorCode
    
    Set prmWarningCode = cmdSetPMTaskComplete.CreateParameter("warningCode", adInteger, adParamInput, , udtAutoParams.WarningBits)
    cmdSetPMTaskComplete.Parameters.Append prmWarningCode
    
    If udtAutoParams.MDID >= 0 Then
        Set prmMDID = cmdSetPMTaskComplete.CreateParameter("MDID", adInteger, adParamInput, , udtAutoParams.MDID)
    Else
        Set prmMDID = cmdSetPMTaskComplete.CreateParameter("MDID", adInteger, adParamInput)
    End If
    cmdSetPMTaskComplete.Parameters.Append prmMDID
    
    Set prmMessage = Nothing
    Set prmMessage = cmdSetPMTaskComplete.CreateParameter("message", adVarChar, adParamOutput, 512, "")
    cmdSetPMTaskComplete.Parameters.Append prmMessage
    
    ' Execute the SP
    cmdSetPMTaskComplete.Execute
    
    If Len(prmMessage.Value) > 0 Then
        AddToPrismAutoAnalysisLog "Error calling SP " & strSetPMTaskComplete & ": " & CStrSafe(prmMessage.Value)
    End If
    
    Set cmdSetPMTaskComplete.ActiveConnection = Nothing
    Set cmdSetPMTaskComplete = Nothing
    
    SetPeakMatchingTaskComplete = True
    Exit Function

SetPeakMatchingTaskCompleteErrorHandler:
    Debug.Print Err.Description
    AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > Error while calling SP " & strSetPMTaskComplete & ": " & Err.Description
    
    SetPeakMatchingTaskComplete = False

End Function

Private Function SetPeakMatchingTaskToRestart(cnnConnection As ADODB.Connection, udtAutoParams As udtAutoAnalysisParametersType) As Boolean
    ' Call SetPeakMatchingTaskToRestartMaster SP to mark task as complete
    ' Returns True if successfully called, False if an error occurs
    
    Dim strSetPeakMatchingTaskToRestart As String
    
    Dim cmdSetPeakMatchingTaskToRestart As New ADODB.Command
        
    Dim prmTaskID As New ADODB.Parameter
    Dim prmServerName As New ADODB.Parameter
    Dim prmMtdbName As New ADODB.Parameter
    Dim prmMessage As New ADODB.Parameter
    
On Error GoTo SetPeakMatchingTaskToRestartErrorHandler
        
    With glbPreferencesExpanded.AutoQueryPRISMOptions
        strSetPeakMatchingTaskToRestart = .SetTaskToRestartSPName
    End With
    
    If Len(strSetPeakMatchingTaskToRestart) = 0 Then strSetPeakMatchingTaskToRestart = PRISM_AUTOMATION_SP_RESTART_TASK_DEFAULT
    
    ' Initialize the SP parameters
    InitializeSPCommand cmdSetPeakMatchingTaskToRestart, cnnConnection, strSetPeakMatchingTaskToRestart
    
    Set prmTaskID = Nothing
    Set prmTaskID = cmdSetPeakMatchingTaskToRestart.CreateParameter("taskID", adInteger, adParamInput, , udtAutoParams.MTDBOverride.PeakMatchingTaskID)
    cmdSetPeakMatchingTaskToRestart.Parameters.Append prmTaskID
        
    Set prmServerName = Nothing
    Set prmServerName = cmdSetPeakMatchingTaskToRestart.CreateParameter("ServerName", adVarChar, adParamInput, 128, udtAutoParams.MTDBOverride.ServerName)
    cmdSetPeakMatchingTaskToRestart.Parameters.Append prmServerName
    
    Set prmMtdbName = Nothing
    Set prmMtdbName = cmdSetPeakMatchingTaskToRestart.CreateParameter("mtdbName", adVarChar, adParamInput, 128, udtAutoParams.MTDBOverride.MTDBName)
    cmdSetPeakMatchingTaskToRestart.Parameters.Append prmMtdbName
    
    Set prmMessage = Nothing
    Set prmMessage = cmdSetPeakMatchingTaskToRestart.CreateParameter("message", adVarChar, adParamOutput, 512, "")
    cmdSetPeakMatchingTaskToRestart.Parameters.Append prmMessage
    
    ' Execute the SP
    cmdSetPeakMatchingTaskToRestart.Execute
    
    If Len(prmMessage.Value) > 0 Then
        AddToPrismAutoAnalysisLog "Error calling SP " & strSetPeakMatchingTaskToRestart & ": " & CStrSafe(prmMessage.Value)
    End If
    
    Set cmdSetPeakMatchingTaskToRestart.ActiveConnection = Nothing
    Set cmdSetPeakMatchingTaskToRestart = Nothing
    
    SetPeakMatchingTaskToRestart = True
    Exit Function

SetPeakMatchingTaskToRestartErrorHandler:
    Debug.Print Err.Description
    AddToPrismAutoAnalysisLog GetCurrentTimeStamp() & " > Error while calling SP " & strSetPeakMatchingTaskToRestart & ": " & Err.Description
    
    SetPeakMatchingTaskToRestart = False

End Function

Private Sub RenameControlFile(ByVal strControlFilePath As String)
    
    Dim strNewControlFilePath As String
    Dim fso As FileSystemObject
    
    ' Rename the file to strControlFilePath.Done
    Set fso = New FileSystemObject
    
    On Error Resume Next
    If fso.FileExists(strControlFilePath) Then
        strNewControlFilePath = strControlFilePath & ".Done"
        If fso.FileExists(strNewControlFilePath) Then
            fso.DeleteFile strNewControlFilePath, True
            Sleep 100
        End If
        
        fso.MoveFile strControlFilePath, strNewControlFilePath
    End If
    
    Set fso = Nothing

End Sub

Private Sub SetPRISMQueryInterval(Optional lngSecondsBetweenQueries As Long = 60)
    Static blnUpdating As Boolean
    
    If blnUpdating Then Exit Sub
    blnUpdating = True
    
    If lngSecondsBetweenQueries < 1 Then lngSecondsBetweenQueries = 1
    mPRISMQueryInterval = lngSecondsBetweenQueries
    
    glbPreferencesExpanded.AutoQueryPRISMOptions.QueryIntervalSeconds = mPRISMQueryInterval
    
    UpdateStatus
    
    txtQueryInterval = Trim(mPRISMQueryInterval)
    
    blnUpdating = False
End Sub

Private Sub ToggleLockoutControls(blnLockoutEnabled As Boolean)
    mControlsLocked = blnLockoutEnabled
    
    chkExclusiveDatabaseProcessing.Enabled = Not mControlsLocked
    chkShowDebugPrompts.Enabled = Not mControlsLocked
    
    txtPreferredMTDB.Enabled = Not mControlsLocked
    txtServerForPreferredMTDB.Enabled = Not mControlsLocked
    txtQueryInterval.Enabled = Not mControlsLocked
    txtPriorityMin.Enabled = Not mControlsLocked
    txtPriorityMax.Enabled = Not mControlsLocked
    
    cmdExitAutomation.Enabled = Not mControlsLocked
    cmdExitProgram.Enabled = Not mControlsLocked
    cmdPauseQuery.Enabled = Not mControlsLocked
    
    SetCheckBox chkLockoutControls, mControlsLocked
    
End Sub

Public Sub TogglePause(blnPauseEnabled As Boolean)
    Static PauseStartTime As Date
    
    mPaused = blnPauseEnabled
    
    If blnPauseEnabled Then
        ' Just enabled pausing
        tmrPRISMQueryDelay.Enabled = False
        lblTimeToNextQuery = "Querying paused"
        cmdPauseQuery.Caption = "&Unpause"
        UpdateStatus
        PauseStartTime = Now()
    Else
        ' Just unpaused
        LastUpdateTime = LastUpdateTime + (Now() - PauseStartTime)
        cmdPauseQuery.Caption = "&Pause Querying"
        tmrPRISMQueryDelay.Enabled = True
        UpdateStatus
        CheckElapsedIterations
    End If
    
End Sub

Private Sub UpdateStatus()
    If mPaused Then
        lblExplanation = "VIPER is currently paused.  Click Unpause to enable fully automated processing"
    Else
        lblExplanation = "VIPER is now running in fully automated mode.  It is querying PRISM every " & mPRISMQueryInterval & " seconds to see if any jobs are available for processing.  The listing below shows the log file entries generated during the processing."
        lblExplanation = lblExplanation & "  Note: Create a file named StopViper.txt in the program directory (" & App.Path & ") to automatically close VIPER."
        
    End If
    
End Sub

Private Sub WriteLatestLogEntryToDisk(strTextToAdd As String)
    Dim strPrismLogFilePath As String, strPrismLogTrimmedFilePath As String
    Dim fso As New FileSystemObject
    Dim tsLogFile As TextStream, tsLogFileTrimmed As TextStream
    Dim lngFileSizeBytes As Long, lngBytesToRemove As Long, lngLinesRead As Long
    Dim strLineIn As String, lngBytesRead As Long
    Dim blnSizeChecked As Boolean
    
On Error GoTo WriteLatestLogEntryToDiskErrorHandler

    strPrismLogFilePath = fso.BuildPath(App.Path, PRISM_AUTO_ANALYSIS_LOGFILE_NAME)
    
    If fso.FileExists(strPrismLogFilePath) Then
        ' See if file is too large
        lngFileSizeBytes = FileLen(strPrismLogFilePath)
        lngBytesToRemove = 0.33 * lngFileSizeBytes
        blnSizeChecked = True
        
        If lngFileSizeBytes > PRISM_AUTO_ANALYSIS_LOGFILE_MAXSIZE_BYTES Then
        
            frmProgress.InitializeForm "Reducing PRISM logfile size by 33%", 0, lngFileSizeBytes, False, False, False, Me
            Me.MousePointer = vbHourglass
            
            ' Decrease the file size by 33%
            strPrismLogTrimmedFilePath = fso.BuildPath(App.Path, PRISM_AUTO_ANALYSIS_LOGFILE_NAME) & ".trimmed"
            Set tsLogFileTrimmed = fso.CreateTextFile(strPrismLogTrimmedFilePath, True, False)
            
            Set tsLogFile = fso.OpenTextFile(strPrismLogFilePath, ForReading, True)
            
            lngBytesRead = 0
            Do While Not tsLogFile.AtEndOfStream And lngBytesRead < lngBytesToRemove
                strLineIn = tsLogFile.ReadLine()
                lngBytesRead = lngBytesRead + Len(strLineIn) + 2
                lngLinesRead = lngLinesRead + 1
                If lngLinesRead Mod 100 = 0 Then frmProgress.UpdateProgressBar lngBytesRead
            Loop
            
            Do While Not tsLogFile.AtEndOfStream
                strLineIn = tsLogFile.ReadLine()
                tsLogFileTrimmed.WriteLine (strLineIn)
                lngBytesRead = lngBytesRead + Len(strLineIn) + 2
                lngLinesRead = lngLinesRead + 1
                If lngLinesRead Mod 100 = 0 Then frmProgress.UpdateProgressBar lngBytesRead
            Loop
            
            tsLogFile.Close
            tsLogFileTrimmed.Close
            
            Set tsLogFile = Nothing
            Set tsLogFileTrimmed = Nothing
            
            Sleep 100
            fso.CopyFile strPrismLogTrimmedFilePath, strPrismLogFilePath, True
            fso.DeleteFile strPrismLogTrimmedFilePath
        
            frmProgress.HideForm
            Me.MousePointer = vbDefault
        End If
        
    End If
    
    Set tsLogFile = fso.OpenTextFile(strPrismLogFilePath, ForAppending, True)
    tsLogFile.WriteLine strTextToAdd

CleanUp:
    On Error Resume Next
    If Not tsLogFile Is Nothing Then
        tsLogFile.Close
        Set tsLogFile = Nothing
    End If
    If Not tsLogFileTrimmed Is Nothing Then
        tsLogFileTrimmed.Close
        Set tsLogFileTrimmed = Nothing
    End If
    Set fso = Nothing
    
    Exit Sub

WriteLatestLogEntryToDiskErrorHandler:
    Debug.Print "Error in WriteLatestLogEntryToDisk: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "WriteLatestLogEntryToDisk", Err.Description, 0, False
    If blnSizeChecked Then
        Resume CleanUp
    Else
        Resume Next
    End If
    
End Sub

Private Sub chkExclusiveDatabaseProcessing_Click()
    glbPreferencesExpanded.AutoQueryPRISMOptions.ExclusivelyUseThisDatabase = cChkBox(chkExclusiveDatabaseProcessing)
End Sub

Private Sub chkLockoutControls_Click()
    ToggleLockoutControls cChkBox(chkLockoutControls)
End Sub

Private Sub chkShowDebugPrompts_Click()
    mDebug = cChkBox(chkShowDebugPrompts)
End Sub

Private Sub cmdExitAutomation_Click()
    ExitAutomationQueryUser
End Sub

Private Sub cmdExitProgram_Click()
    ExitProgramQueryUser
End Sub

Private Sub cmdPauseQuery_Click()
    TogglePause Not mPaused
End Sub

Private Sub cmdQueryNow_Click()
    mForcePRISMQueryNow = True
    CheckElapsedIterations
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowUpperThird, 7000, 8000
    
    If mPRISMQueryInterval < 1 Then mPRISMQueryInterval = 60
    
    CheckForRestartViperFile
    
    LastUpdateTime = Now()
    
    tmrPRISMQueryDelay.Interval = TIMER_DELAY_MSEC
    tmrPRISMQueryDelay.Enabled = True
    mViperLocalOnServer = False
    
    InitializeControls
    
    PositionControls
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub tmrPRISMQueryDelay_Timer()
    CheckElapsedIterations
End Sub

Private Sub txtPreferredMTDB_Change()
    glbPreferencesExpanded.AutoQueryPRISMOptions.PreferredDatabaseToProcess = txtPreferredMTDB
End Sub

Private Sub txtPreferredMTDB_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPreferredMTDB, KeyAscii, True, True, True, True, True, True, True, True, False
End Sub

Private Sub txtPriorityMax_Change()
    If IsNumeric(txtPriorityMax) Then
        glbPreferencesExpanded.AutoQueryPRISMOptions.MaximumPriorityToProcess = CIntSafe(txtPriorityMax)
    Else
        glbPreferencesExpanded.AutoQueryPRISMOptions.MaximumPriorityToProcess = 0
    End If
End Sub

Private Sub txtPriorityMax_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPriorityMax, KeyAscii, True, False
End Sub

Private Sub txtPriorityMax_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtPriorityMax, 0, 255, 5
End Sub

Private Sub txtPriorityMin_Change()
    If IsNumeric(txtPriorityMin) Then
        glbPreferencesExpanded.AutoQueryPRISMOptions.MinimumPriorityToProcess = CIntSafe(txtPriorityMin)
    Else
        glbPreferencesExpanded.AutoQueryPRISMOptions.MinimumPriorityToProcess = 0
    End If
End Sub

Private Sub txtPriorityMin_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPriorityMin, KeyAscii, True, False
End Sub

Private Sub txtPriorityMin_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtPriorityMin, 0, 255, 1
End Sub

Private Sub txtQueryInterval_Change()
    If IsNumeric(txtQueryInterval) Then
        SetPRISMQueryInterval CLngSafe(txtQueryInterval)
        If Not mPaused Then tmrPRISMQueryDelay.Enabled = True
    End If
End Sub

Private Sub txtServerForPreferredMTDB_Change()
    glbPreferencesExpanded.AutoQueryPRISMOptions.ServerForPreferredDatabase = txtServerForPreferredMTDB
End Sub

VERSION 5.00
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "VIPER"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8820
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      Begin VB.CommandButton cmdVFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   6720
         Picture         =   "MDIForm1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move down"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdVFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   6360
         Picture         =   "MDIForm1.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move up"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdHFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   4440
         Picture         =   "MDIForm1.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Move right"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdHFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   4080
         Picture         =   "MDIForm1.frx":0600
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Move left"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdVFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   6000
         Picture         =   "MDIForm1.frx":06F2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Expand bottom"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdVFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   5640
         Picture         =   "MDIForm1.frx":07E4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Shrink bottom"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdVFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   5280
         Picture         =   "MDIForm1.frx":08D6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Shrink top"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdVFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   4920
         Picture         =   "MDIForm1.frx":09C8
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Expand top"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdHFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   3720
         Picture         =   "MDIForm1.frx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Expand right"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdHFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   3360
         Picture         =   "MDIForm1.frx":0BAC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Shrink right"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdHFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   3000
         Picture         =   "MDIForm1.frx":0C9E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Shrink left"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdHFineTune 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2640
         Picture         =   "MDIForm1.frx":0D90
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Expand left"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdICR2LS 
         Height          =   330
         Left            =   2160
         Picture         =   "MDIForm1.frx":0E82
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Launch ICR-2LS application"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdLast 
         Height          =   330
         Left            =   7080
         Picture         =   "MDIForm1.frx":0F84
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdCopy 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1680
         Picture         =   "MDIForm1.frx":14B6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Copy gel on Clipboard"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1200
         Picture         =   "MDIForm1.frx":19E8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print gel"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   330
         Left            =   720
         Picture         =   "MDIForm1.frx":1F1A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdNew 
         Height          =   330
         Left            =   360
         Picture         =   "MDIForm1.frx":244C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "New .GEL file"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   330
         Left            =   0
         Picture         =   "MDIForm1.frx":297E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Open .GEL file"
         Top             =   0
         Width           =   330
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "0%"
         Height          =   195
         Left            =   7560
         TabIndex        =   3
         Top             =   45
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New (Load peak list file)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuNewAnalysis 
         Caption         =   "New Anal&ysis (Choose from DMS)"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuNewAutoAnalysis 
         Caption         =   "New Automatic Analysis (Choose manually)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open .Gel file"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveLoadEditAnalysisSettings 
         Caption         =   "Save/Load/Edit Analysis Settings"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile8"
         Index           =   8
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPEK 
         Caption         =   "&PEK Functions"
         Begin VB.Menu mnuPEKMerge 
            Caption         =   "&Merge PEK Files"
         End
         Begin VB.Menu mnuPEKSplit 
            Caption         =   "&Split PEK File"
         End
         Begin VB.Menu mnuPEKFilter 
            Caption         =   "&Filter PEK File"
         End
         Begin VB.Menu mnuPEKModify 
            Caption         =   "&Modify Data"
         End
      End
      Begin VB.Menu mnu3DErrorViewer 
         Caption         =   "3D Error Viewer (Load from File)"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuToolsSepAutomation 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsGenerateIndexHtmlFiles 
         Caption         =   "Generate Index.HTML files"
      End
      Begin VB.Menu mnuInitiatePRISMAutomation 
         Caption         =   "Initiate Automated PRISM Analysis"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuV2DGelHelp 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu mnuHelpSetTraceLogLevel 
         Caption         =   "Set Trace Log Level"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuEndAutomation 
      Caption         =   "End Automation"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Last modified: 02/11/2003 nt
'-------------------------------------------------
Option Explicit

Public WithEvents MyAnalysisInit As AnalysisInitiator
Attribute MyAnalysisInit.VB_VarHelpID = -1

Public Sub InitiatePRISMAutomation(Optional blnInitiatedViaCommandLine As Boolean = False)
    
    Dim strMessage As String
    
    ' If the software was not initiated by the command line, then prompt for the auto-analysis password
    If Not blnInitiatedViaCommandLine Then
        ' Query user for Auto Analysis Password
        strMessage = "Automated Prism Analysis is an advanced feature that should only be used when running VIPER automatically on an analysis server.  Please enter the password for initiating PRISM automation:"
        If Not QueryUserForExportToDBPassword(strMessage) Then
            Exit Sub
        End If
    End If
    
    ShowHideAutomationMenus True
    frmPRISMAutomation.InitializeControls
    
    If blnInitiatedViaCommandLine Then
        frmPRISMAutomation.InitiateFromCommandLine
    Else
        frmPRISMAutomation.TogglePause True
    End If
    
    frmPRISMAutomation.Show vbModeless
    Me.Hide
End Sub

Public Sub RestoreMenus()
    ShowHideAutomationMenus False
End Sub

Private Sub ShowHideAutomationMenus(blnHideMenus As Boolean)
    ' Hide all the menus, then show mnuEndAutomation
    Dim ctlThisControl As VB.Control
    
    On Error Resume Next
    For Each ctlThisControl In Me.Controls
        If TypeOf ctlThisControl Is Menu Then
            If ctlThisControl.Name = "mnuRecentFiles" And Not blnHideMenus Then
                If Len(ctlThisControl.Caption) > 0 Then
                    ctlThisControl.Visible = True
                End If
            Else
                ctlThisControl.Visible = Not blnHideMenus
            End If
        End If
    Next ctlThisControl
    
    picToolbar.Visible = Not blnHideMenus
    
    mnuEndAutomation.Visible = blnHideMenus
End Sub

Public Sub UnloadPRISMAutomationForm(Optional blnCloseProgram As Boolean = True)
    Unload frmPRISMAutomation
    If blnCloseProgram Then
        Unload Me
    End If
End Sub

Private Sub cmdCopy_Click()
Dim Ind As Long
On Error Resume Next
If Not MDIForm1.ActiveForm Is Nothing Then
   Ind = CLng(MDIForm1.ActiveForm.Tag)
   GelDrawMetafile Ind, False, "", False
End If
End Sub

Private Sub cmdHFineTune_Click(Index As Integer)
Dim gbInd As Long
Dim nDirection As Integer
On Error GoTo err_hfinetune
Select Case Index
Case 0, 3
     nDirection = glExpand
Case 1, 2
     nDirection = glShrink
End Select
If Not MDIForm1.ActiveForm Is Nothing Then
   gbInd = CLng(MDIForm1.ActiveForm.Tag)
   With GelBody(gbInd).csMyCooSys
        Select Case Index
        Case 0, 1
            If .csXOrient = glNormal Then
               .FineTuneH nDirection, glTuneLT
            Else
               .FineTuneH nDirection, glTuneRB
            End If
        Case 2, 3
            If .csXOrient = glNormal Then
               .FineTuneH nDirection, glTuneRB
            Else
               .FineTuneH nDirection, glTuneLT
            End If
        Case 4
            .FineTuneH nDirection, glTuneMoveL
        Case 5
            .FineTuneH nDirection, glTuneMoveR
        End Select
   End With
End If
err_hfinetune:
End Sub

Private Sub cmdICR2LS_Click()
LaunchICR2LS
End Sub

Private Sub cmdVFineTune_Click(Index As Integer)
Dim gbInd As Long
Dim nDirection As Integer
On Error GoTo err_vfinetune
Select Case Index
Case 0, 3
     nDirection = glExpand
Case 1, 2
     nDirection = glShrink
End Select
If Not MDIForm1.ActiveForm Is Nothing Then
   gbInd = CLng(MDIForm1.ActiveForm.Tag)
   With GelBody(gbInd).csMyCooSys
        Select Case Index
        Case 0, 1
             If .csYOrient = glNormal Then
                .FineTuneV nDirection, glTuneRB
             Else
                .FineTuneV nDirection, glTuneLT
             End If
        Case 2, 3
             If .csYOrient = glNormal Then
                .FineTuneV nDirection, glTuneLT
             Else
                .FineTuneV nDirection, glTuneRB
             End If
        Case 4
             .FineTuneV nDirection, glTuneMoveU
        Case 5
             .FineTuneV nDirection, glTuneMoveD
        End Select
   End With
End If
err_vfinetune:
End Sub

Private Sub cmdNew_Click()
mnuNew_Click
End Sub

Private Sub cmdOpen_Click()
mnuOpen_Click
End Sub

Private Sub cmdPrint_Click()
Dim Ind As Long
If Not MDIForm1.ActiveForm Is Nothing Then
   Ind = CLng(MDIForm1.ActiveForm.Tag)
   If GoPrint(Ind) < 0 Then MsgBox "Error printing.", vbOKOnly
End If
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If Not MDIForm1.ActiveForm Is Nothing Then
   MDIForm1.ActiveForm.mnuSave_Click
End If
End Sub

Private Sub MDIForm_Load()

''
''Dim objTest2 As New Dictionary
''
''objTest2.add 15, 1
''objTest2.add 20, 2
''objTest2.add 22, 3
''
''Dim lngItemsToKeep() As Long
''ReDim lngItemsToKeep(99)
''Dim lngItemCount As Long
''
''lngItemCount = 0
''
''lngItemsToKeep(lngItemCount) = 15
''lngItemCount = lngItemCount + 1
''
''If UBound(lngItemCount) <= lngItemCount Then
''    ReDim Preserve lngItemCount(UBound(lngItemCount) * 2 - 1)
''End If
''
''lngItemsToKeep(lngItemCount) = 20
''lngItemCount = lngItemCount + 1
''
''lngItemsToKeep(lngItemCount) = 22
''lngItemCount = lngItemCount + 1
''
''
''ShellSortLong lngItemsToKeep, 0, lngItemCount - 1
''lngIndex = BinarySearchLng(lngItemsToKeep, 15, 0, lngItemCount - 1)
''
''Debug.Print "done"

On Error Resume Next
ChDir App.Path
ChDrive App.Path
Debug.Assert LCase(CurDir()) = LCase(App.Path)

On Error GoTo MDIFormLoadErrorHandler

Randomize Timer

mnuEndAutomation.Visible = False

ShowHidePNNLMenus

SizeAndCenterWindow Me, cWindowTopLeft, 800 * Screen.TwipsPerPixelX, 600 * Screen.TwipsPerPixelY

Me.Show
Me.MousePointer = vbHourglass

ReDim GelBody(0)
ReDim GelStatus(0)
ReDim GelData(0)
lblStatus.Left = picToolbar.width - lblStatus.width - 100
Initialize
Me.MousePointer = vbDefault


Exit Sub

MDIFormLoadErrorHandler:
Debug.Print "Error in MDIForm_Load (or one of its subroutines): " & Err.Description
Debug.Assert False
LogErrors Err.Number, "MDIForm1->MDIForm_Load (or one of its subroutines)"
Resume Next

End Sub

'Private Sub BenchmarkDictionary()
'    Dim lngIndex As Long
'    Dim lngDictionarySize As Long
'    Dim lngDataArraySize As Long
'
'    Dim htKeys As clsHashTable
'    Dim objKeyValPairs As clsParallelLngArrays
'
'    Dim lngKeys() As Long
'    Dim lngItems() As Long
'    Dim varItem As Variant
'
'    Dim lngDataArray() As Long
'    Dim lngValue As Long
'    Dim lngMatchCount As Long
'    Dim lngIndexMatch As Long
'
'    Dim blnMatchFound As Boolean
'
'    Dim dtStartTime As Date
'    Const MAX_WAIT_TIME_SECONDS As Integer = 30
'
'    lngDictionarySize = 1000000
'    lngDataArraySize = 5000000
'
''    Set htKeys = New Dictionary
''    Debug.Print Now() & "; Populate htKeys"
''    dtStartTime = Now()
''
''    ' Populate the dictionary
''    For lngIndex = 0 To lngDictionarySize - 1
''        htKeys.Add lngIndex * 3, 1234
''        If lngIndex Mod 25000 = 0 Then
''            Debug.Print Now() & "; Populating hashtable: " & lngIndex
''            If DateDiff("s", dtStartTime, Now()) >= MAX_WAIT_TIME_SECONDS Then Exit For
''        End If
''    Next lngIndex
''
''    Debug.Print Now() & "; Processing " & lngDataArraySize & " data rows against the hash table (contains " & htKeys.Count & " keys)"
''    dtStartTime = Now()
''
''    lngMatchCount = 0
''    For lngIndex = 0 To lngDataArraySize
''        If htKeys.Exists(lngIndex) Then
''            lngMatchCount = lngMatchCount + 1
''            lngValue = htKeys(lngIndex)
''        End If
''        If lngIndex Mod 10000 = 0 Then
''            Debug.Print Now() & "; Testing hashtable: " & lngIndex
''            If DateDiff("s", dtStartTime, Now()) >= MAX_WAIT_TIME_SECONDS Then Exit For
''        End If
''    Next lngIndex
''
''    Debug.Print Now() & "; Found " & lngMatchCount & " matches against the hash table"
''
''
''    Set htKeys = New clsHashTable
''
''    ' Can turn off Ignore Case since our keys are numbers
''    htKeys.IgnoreCase = False
''    htKeys.SetSize lngDictionarySize
''
''    Debug.Print Now() & "; Populate htKeys (using clsHashTable)"
''    dtStartTime = Now()
''
''    ' Populate the dictionary
''    For lngIndex = 0 To lngDictionarySize - 1
''        htKeys.Add lngIndex * 3, 1234
''        If lngIndex Mod 25000 = 0 Then
''            Debug.Print Now() & "; Populating hashtable: " & lngIndex
''            If DateDiff("s", dtStartTime, Now()) >= MAX_WAIT_TIME_SECONDS Then Exit For
''        End If
''    Next lngIndex
''
''    Debug.Print Now() & "; Processing " & lngDataArraySize & " data rows against the hash table (contains " & htKeys.Count & " keys)"
''    dtStartTime = Now()
''
''    lngMatchCount = 0
''    For lngIndex = 0 To lngDataArraySize
''        varItem = htKeys.Item(CStr(lngIndex))
''        If Not IsEmpty(varItem) Then
''            lngMatchCount = lngMatchCount + 1
''            lngValue = CStr(varItem)
''        End If
''        If lngIndex Mod 100000 = 0 Then
''            Debug.Print Now() & "; Testing hashtable: " & lngIndex
''            If DateDiff("s", dtStartTime, Now()) >= MAX_WAIT_TIME_SECONDS Then Exit For
''        End If
''    Next lngIndex
''
''    Debug.Print Now() & "; Found " & lngMatchCount & " matches against the hash table"
'
'
'    Set objKeyValPairs = New clsParallelLngArrays
'
'    objKeyValPairs.SetSize lngDictionarySize
'    objKeyValPairs.PreventDuplicateKeys = False
'
'    Debug.Print Now() & "; Populate objKeyValPairs (using clsParallelLngArrays)"
'    dtStartTime = Now()
'
'    ' Populate the dictionary
'    For lngIndex = 0 To lngDictionarySize - 1
'        objKeyValPairs.Add lngIndex * 3, 1234
'        If lngIndex Mod 25000 = 0 Then
'            Debug.Print Now() & "; Populating hashtable: " & lngIndex
'            If DateDiff("s", dtStartTime, Now()) >= MAX_WAIT_TIME_SECONDS Then Exit For
'        End If
'    Next lngIndex
'
'    Debug.Print Now() & "; Sorting the keys in objKeyValPairs"
'    objKeyValPairs.SortNow
'
'    Debug.Print Now() & "; Processing " & lngDataArraySize & " data rows against the hash table (contains " & objKeyValPairs.Count & " keys)"
'    dtStartTime = Now()
'
'    lngMatchCount = 0
'    For lngIndex = 0 To lngDataArraySize
'        lngValue = objKeyValPairs.GetItemForKey(lngIndex, blnMatchFound)
'
'        If blnMatchFound Then
'            lngMatchCount = lngMatchCount + 1
'        End If
'        If lngIndex Mod 100000 = 0 Then
'            Debug.Print Now() & "; Testing hashtable: " & lngIndex
'            If DateDiff("s", dtStartTime, Now()) >= MAX_WAIT_TIME_SECONDS Then Exit For
'        End If
'    Next lngIndex
'
'    Debug.Print Now() & "; Found " & lngMatchCount & " matches against the hash table"
'
'
'
'    Debug.Print Now() & "; Populate lngKeys() and lngItems()"
'
'    ReDim lngKeys(lngDictionarySize - 1)
'    ReDim lngItems(lngDictionarySize - 1)
'
'    For lngIndex = 0 To lngDictionarySize - 1
'        lngKeys(lngIndex) = (lngDictionarySize - 1 - lngIndex) * 3
'        lngItems(lngIndex) = 1234
'        If lngIndex Mod 25000 = 0 Then Debug.Print Now() & "; Populating parallel arrays: " & lngIndex
'    Next lngIndex
'
'    Debug.Print Now() & "; Sort lngKeys() and lngItems()"
'    ShellSortLongWithParallelLong lngKeys, lngItems, 0, lngDictionarySize - 1
'
'    Debug.Print Now() & "; Processing " & lngDataArraySize & " data rows against the parallel arrays"
'
'    lngMatchCount = 0
'    For lngIndex = 0 To lngDataArraySize
'        lngIndexMatch = BinarySearchLng(lngKeys, lngIndex)
'
'        If lngIndexMatch >= 0 Then
'            lngValue = lngItems(lngIndexMatch)
'            lngMatchCount = lngMatchCount + 1
'        End If
'
'        If lngIndex Mod 100000 = 0 Then Debug.Print Now() & "; Testing parallel arrays: " & lngIndex
'    Next lngIndex
'
'    Debug.Print Now() & "; Found " & lngMatchCount & " matches against the parallel arrays"
'
'    Debug.Assert False
'
'End Sub

Private Sub MDIForm_Resize()
If IsWinLoaded(TrackerCaption) Then
   If Me.WindowState = vbMinimized Then
      frmTracker.Hide
   Else
      lblStatus.Left = picToolbar.width - lblStatus.width - 100
      ' Need Resume Next error handling in case this function gets called when a form is shown Modally
      On Error Resume Next
      frmTracker.Show
   End If
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'If the Unload event was not canceled from any of the child forms
'then unload ICR-2LS(if it was not loaded on start) and end application
On Error Resume Next
Me.MousePointer = vbHourglass
DoEvents

Erase GelBody
If OlyCnt > 0 Then Erase Oly  'this should not happen but just in case
CleanICR2LS
CloseConnections
IniFileSaveSettings glbPreferencesExpanded, UMCDef, UMCIonNetDef, UMCNetAdjDef, UMCInternalStandards, samtDef, glPreferences
Unload frmProgress
DestroyDrawingObjects
DestroyMTSearchObject
If Not MyAnalysisInit Is Nothing Then Set MyAnalysisInit = Nothing

''' MonroeMod
''If Not ORFViewerLoader Is Nothing Then Set ORFViewerLoader = Nothing

' This End command is included because sometimes some of the classes don't unload and Viper thus keeps running
End
End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuToolsGenerateIndexHtmlFiles.Visible = blnVisible
    mnuInitiatePRISMAutomation.Visible = blnVisible
    mnuToolsSepAutomation.Visible = blnVisible
    
    mnuNewAnalysis.Visible = Not APP_BUILD_DISABLE_MTS
End Sub

Private Sub mnu3DErrorViewer_Click()
frmErrorDistribution3DFromFile.Show
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, MDIForm1
End Sub

Private Sub mnuEndAutomation_Click()
    frmPRISMAutomation.ExitAutomationQueryUser
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelpSetTraceLogLevel_Click()
    SetTraceLogLevel 0, True
End Sub

Private Sub mnuInitiatePRISMAutomation_Click()
    InitiatePRISMAutomation
End Sub

Private Sub mnuNew_Click()
MDIStatus True, "Loading ... "
FileNew (Me.hwnd)
End Sub

Private Sub mnuNewAnalysis_Click()
    On Error GoTo NewAnalysisErrorHandler
    Set MyAnalysisInit = New AnalysisInitiator
    MyAnalysisInit.GetNewAnalysisDialog glInitFile

    Exit Sub
    
NewAnalysisErrorHandler:
    MsgBox "Error initiating a new analysis of a file on DMS (" & Err.Number & "):" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub mnuNewAutoAnalysis_Click()
Dim udtAutoParams As udtAutoAnalysisParametersType

InitializeAutoAnalysisParameters udtAutoParams
udtAutoParams.ShowMessages = True

AutoAnalysisStart udtAutoParams
End Sub

Private Sub mnuOpen_Click()
MDIStatus True, "Loading ... "
FileOpenProc (Me.hwnd)
End Sub

Private Sub mnuOptions_Click()
On Error Resume Next
frmOptions.Tag = "MDI"
frmOptions.Show vbModal
If Abs(vWhatever) > 0 Then   'colors changed
   SetBackForeColorObjects
   SetCSIsoColorObjects
   SetDDRColorObjects
   SetSelColorObjects
End If
End Sub

Private Sub mnuPEKFilter_Click()
frmPEKFilter.Show vbModal
End Sub

Private Sub mnuPEKMerge_Click()
frmPEKMerge.Show vbModal
End Sub

Private Sub mnuPEKModify_Click()
frmPEKScrambler.Show vbModal
End Sub

Private Sub mnuPEKSplit_Click()
frmPEKSplit.Show vbModal
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
    ' MonroeMod
    Dim strFullFilePath As String
    
    strFullFilePath = RecentFileLookUpFullPath(mnuRecentFiles(Index).Caption)
    If Len(strFullFilePath) > 0 Then
        MDIStatus True, "Loading ... please be patient"
        ReadGelFile strFullFilePath
    End If
End Sub

Private Sub mnuSaveLoadEditAnalysisSettings_Click()
    frmEditAnalysisSettings.SetCallerID 0
    frmEditAnalysisSettings.Show vbModeless, Me
End Sub

Private Sub mnuToolsGenerateIndexHtmlFiles_Click()
    Me.MousePointer = vbHourglass
    DoEvents
    
    GenerateAutoAnalysisHtmlFiles
    frmProgress.HideForm
    
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuV2DGelHelp_Click()
    MsgBox "Please see the Powerpoint Help file (e.g. VIPER_HelpFile_v3.20.ppt).  If inside PNNL, you can find this file at \\floyd\software\VIPER\ .  If outside PNNL, please visit http://omics.pnl.gov/software/ for help on using this software.", vbInformation + vbOKOnly, glFGTU
End Sub

Public Sub ProperToolbar(ByVal bAnyChild As Boolean)
Dim i As Long
For i = 0 To 5
    cmdHFineTune(i).Enabled = bAnyChild
    cmdVFineTune(i).Enabled = bAnyChild
Next i
cmdCopy.Enabled = bAnyChild
cmdPrint.Enabled = bAnyChild
cmdSave.Enabled = bAnyChild
End Sub

Private Sub MyAnalysisInit_DialogClosed()
Dim Ind As Long
Dim sFileName As String
Dim strExtension As String

Dim OpenResult As Long

' MonroeMod
Dim udtDBSettings As udtDBSettingsType
Dim eResponse As VbMsgBoxResult
Dim fso As New FileSystemObject

On Error GoTo err_NewAnalysis

If Not MyAnalysisInit.NewAnalysis Is Nothing Then
    Ind = FindFreeIndex()
    Set GelAnalysis(Ind) = MyAnalysisInit.NewAnalysis
    
    ' MonroeMod Begin
    FillDBSettingsUsingAnalysisObject udtDBSettings, GelAnalysis(Ind)
    sFileName = fso.BuildPath(GelAnalysis(Ind).Desc_DataFolder, GelAnalysis(Ind).MD_file)
    
    If Len(sFileName) = 0 Then
        ' User cancelled or chose database with no files
        MsgBox "Valid file not selected", vbOKOnly, glFGTU
        MDIStatus False, "Done"
        GelStatus(Ind).Deleted = True
    Else
    
        ' Do not update this here since can take awhile: udtDBSettings.SelectedMassTagCount
        IniFileUpdateRecentDatabaseConnectionInfo udtDBSettings
        AddToAnalysisHistory Ind, "Database connection defined: " & udtDBSettings.DatabaseName & "; " & udtDBSettings.SelectedMassTagCount & " MT tags"
    
        If Not FileExists(sFileName) Then
            OpenResult = -6
        Else
            GelData(Ind).Certificate = glCERT2003
            GelData(Ind).Comment = glCOMMENT_CREATED & Now & vbCrLf & glCOMMENT_USER & UserName
         
            ' MonroeMod
            AddToAnalysisHistory Ind, "New gel created (user " & UserName & ")"
        
            GelData(Ind).pICooSysEnabled = False
            GelData(Ind).PathtoDatabase = glbPreferencesExpanded.LegacyAMTDBPath
            
            If Ind > glMaxGels Then
               MsgBox "Command aborted. Too many open files.", vbOKOnly, glFGTU
               Exit Sub
            End If
            With GelData(Ind)  'save parameters for this doc
              ' The full path to the .Pek, .CSV, .mzXML, or .mzData file
              .FileName = sFileName
              .Fileinfo = GetFileInfo(sFileName)
              ResetDataFilters Ind, glPreferences
            
              ' MonroeMod
              AddToAnalysisHistory Ind, "Loading File; " & .Fileinfo
            End With
            
            eResponse = MsgBox("Auto-analyze file?  If you choose Yes, you will need to select a .Ini file that contains the auto analysis settings.", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Auto Analyze")
            If eResponse = vbYes Then
                
                GelAnalysis(Ind).MD_Parameters = GelAnalysis(Ind).GetParameters
        
                Dim udtAutoParams As udtAutoAnalysisParametersType
        
                InitializeAutoAnalysisParameters udtAutoParams
                With udtAutoParams
                  .ShowMessages = True
                  .FilePaths.InputFilePath = sFileName
                  .GelIndexToForce = Ind
                  .AutoDMSAnalysisManuallyInitiated = True
                End With
                
                If AutoAnalysisStart(udtAutoParams, True, False) Then
                  OpenResult = 1
                Else
                  OpenResult = -50
                End If
            Else
                OpenResult = LoadNewData(fso, sFileName, Ind, True, "")
            End If
        End If
        
        strExtension = GetFileExtension(sFileName)
        
        Select Case OpenResult
        Case 0      'success
           GelStatus(Ind).Dirty = True
           GelBody(Ind).Tag = Ind
           GelBody(Ind).Caption = "Untitled:" & Ind & " (" & fso.GetBaseName(sFileName) & ")"
           GelStatus(Ind).GelFilePathFull = App.Path
           GelBody(Ind).Show
        Case 1      'auto analysis success
           ' Will already have been set to "Dirty" by the auto-analysis process and changed to not Dirty if successfully saved
           GelBody(Ind).Tag = Ind
           GelBody(Ind).Show
        Case -1     'user canceled load of large data set
           MDIStatus False, "Done"
           GelStatus(Ind).Deleted = True
        Case -2     'data sets too large
        Case -3     'data structure problem
           MDIStatus False, "Done"
           MsgBox "Scan numbers in the " & strExtension & " file must be in ascending order." & vbCrLf & _
                  "Open " & strExtension & " file with any text editor and make changes.", vbOKOnly, glFGTU
           GelStatus(Ind).Deleted = True
        Case -4     'no valid data
           MDIStatus False, "Done"
           MsgBox "No valid data found in file: " & sFileName, vbOKOnly, glFGTU
           GelStatus(Ind).Deleted = True
        Case -5     ' User Cancelled load in the middle of loading (or post-load processing)
           MDIStatus False, "Done"
           MsgBox "Load cancelled.", vbOKOnly, glFGTU
           GelStatus(Ind).Deleted = True
        Case -6, -7    ' File not found
           MDIStatus False, "Done"
           If OpenResult = -6 Then
              MsgBox "File not found:" & vbCrLf & sFileName, vbOKOnly, glFGTU
           End If
           GelStatus(Ind).Deleted = True
        Case -50     ' Error during auto-analysis
           MDIStatus False, "Done"
           If GelData(Ind).IsoLines > 0 Or GelData(Ind).CSLines > 0 Then
              ' If data is in memory, then show the graph
              GelBody(Ind).Tag = Ind
              GelBody(Ind).Show
           Else
              ' No data in memory; mark as deleted
              MDIStatus False, "Done"
              GelStatus(Ind).Deleted = True
           End If
        Case Else   'some other error
           MDIStatus False, "Done"
           MsgBox "Error loading data from file " & sFileName & "." & vbCrLf & "File maybe contains no data or structure does not match expected.", vbOKOnly, glFGTU
           GelStatus(Ind).Deleted = True
        End Select
        
        If GelStatus(Ind).Deleted Then
             SetGelStateToDeleted Ind
        End If
    End If
    GelAnalysis(Ind).MD_Parameters = GelAnalysis(Ind).GetParameters
End If

exit_NewAnalysis:
Set MyAnalysisInit = Nothing
frmProgress.HideForm
Screen.MousePointer = vbDefault
Exit Sub

err_NewAnalysis:
LogErrors Err.Number, "MDIForm.MyAnalysisInit_DialogClosed"
MsgBox "Error initiating new analysis.", vbOKOnly, glFGTU
Resume exit_NewAnalysis
End Sub

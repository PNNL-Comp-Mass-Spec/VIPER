VERSION 5.00
Begin VB.Form frmNewAnalysis 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Analysis to Load"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   Icon            =   "frmNewAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleMode       =   0  'User
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGANETInfo 
      Caption         =   "&GANET Info"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "TIC Alignment Info"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdExperimentInfo 
      Caption         =   "&Experiment"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      ToolTipText     =   "Experiment Info (Browser)"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnalysisInfo 
      Caption         =   "&Analysis"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      ToolTipText     =   "Analysis Job Info (Browser)"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDatasetInfo 
      Caption         =   "&Dataset"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "Dataset Info (Browser)"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraStage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Analysis Parameters"
      Height          =   4335
      Index           =   3
      Left            =   2400
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton cmdResetParameters 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   5880
         TabIndex        =   37
         ToolTipText     =   "Reset parameters"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtParameters 
         Height          =   3015
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   960
         Width           =   6735
      End
      Begin VB.ComboBox cmbAnalysisType 
         Height          =   315
         ItemData        =   "frmNewAnalysis.frx":030A
         Left            =   1920
         List            =   "frmNewAnalysis.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Analysis type:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.Frame fraStage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selection of Analysis Result File"
      Height          =   4335
      Index           =   2
      Left            =   1800
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox txtShowTextFile 
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2640
         Width           =   6735
      End
      Begin VB.ComboBox cmbFileType 
         Height          =   315
         ItemData        =   "frmNewAnalysis.frx":0353
         Left            =   1920
         List            =   "frmNewAnalysis.frx":0355
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   360
         Width           =   1935
      End
      Begin VB.FileListBox lbFileList 
         DragIcon        =   "frmNewAnalysis.frx":0357
         Height          =   1455
         Left            =   240
         Pattern         =   "*.dll"
         TabIndex        =   32
         Top             =   960
         Width           =   6735
      End
      Begin VB.Label lblSelectedFile 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   3840
         Width           =   6735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "List files of type:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.Frame fraStage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selection of ICR-2LS Analysis Result File"
      Height          =   4335
      Index           =   1
      Left            =   1080
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton cmdSelectDatasetAnalysis 
         Caption         =   "Select"
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "&Info"
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdShowAll 
         Caption         =   "Show &All"
         Height          =   285
         Left            =   3840
         TabIndex        =   20
         ToolTipText     =   "Shows All ICR2LS Analysis Results"
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox lstICR2LSJobs 
         Height          =   2790
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   6735
      End
      Begin VB.CommandButton cmdShowNew 
         Caption         =   "Show Ne&w"
         Height          =   285
         Left            =   4920
         TabIndex        =   21
         ToolTipText     =   "Shows All ICR2LS Analysis Results With State=New"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         ItemData        =   "frmNewAnalysis.frx":0799
         Left            =   6120
         List            =   "frmNewAnalysis.frx":07BE
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3570
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4200
         TabIndex        =   25
         Top             =   3540
         Width           =   1095
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Text            =   "Summus Deus"
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dataset/Analysis Folder:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblSelectionDatasetAnalysis 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   28
         ToolTipText     =   "Reference Job ID/ Datasets Name"
         Top             =   3960
         Width           =   6615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   26
         Top             =   3660
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search for:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   3660
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   5040
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   5040
      Width           =   900
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   4560
      Width           =   900
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   4560
      Width           =   900
   End
   Begin VB.Frame fraStage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Organism Mass Tag Database/Process Type"
      Height          =   4335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdSearchForDB 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4080
         TabIndex        =   42
         Top             =   3270
         Width           =   1095
      End
      Begin VB.TextBox txtSearchForDB 
         Height          =   285
         Left            =   1200
         TabIndex        =   41
         Top             =   3320
         Width           =   2655
      End
      Begin VB.CheckBox chkShowUnusedDBs 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Show Unused Databases"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ListBox lstOrgMTDBNames 
         Height          =   2205
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4455
      End
      Begin VB.CheckBox chkShowFrozenDBs 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Show Frozen Databases"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdSelectMT 
         Caption         =   "Select Mass Tags"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         ToolTipText     =   "Select which mass tags will be used for search"
         Top             =   1200
         Width           =   1900
      End
      Begin VB.CommandButton cmdConfigure 
         Caption         =   "Co&nfigure DB"
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   600
         Width           =   1900
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search for:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   3345
         Width           =   855
      End
      Begin VB.Label lblMTDBDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "No directory data found; Server might be down!"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   3650
         Width           =   6735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Organism Mass Tag Database"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Establishing connection with the server..."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   5400
      Width           =   7215
   End
End
Attribute VB_Name = "frmNewAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'form that controls initialization of new FTICR analysis
'from the selected Organism Mass Tag database
'-------------------------------------------------------
'last modified: 05/23/2002 nt
'last modified: 01/06/2004 by mem
'-------------------------------------------------------
'NOTE: cmdDatasetInfo and cmdAnalysisInfo should be
'replaced with hyperlink controls when done!
'NOTE: previous version included two options
'optProcessing that would allow to set processing type
'for fAnalysis to fptMassLock or fptMassMatch. This is
'now always set to fptMassMatch but it might come in
'use again
'-------------------------------------------------------
Option Explicit

Const SMALL_FILE_MAX_SIZE = 30 * 1024

Const FILE_TYPE_PEK = 0
Const FILE_TYPE_CSV = 1
Const FILE_TYPE_HDF = 2
Const FILE_TYPE_ANY = 3

Const MAX_ROWS = 500        'maximum number of rows in list

Const STAGE_COUNT = 4
Const SECTION_MTDB = "[Organism Mass Tag Database]"

'initialization file entries containing data access queries
'dataset/analysis list, analysis type table
Const INIT_Fill_All = "sql_GET_ICR2LS_Analyses_List_All"
Const INIT_Fill_New = "sql_GET_ICR2LS_Analyses_List_New"
Const INIT_Fill_Year = "sql_GET_ICR2LS_Analyses_List_Search_Year"
Const INIT_Fill_Pattern = "sql_GET_ICR2LS_Analyses_List_Search_Name"

Const INIT_t_Analysis_Type = "Analysis Type Table"
Const INIT_Get_DB_Schema_Version = "sp_GetDBSchemaVersion"

'provides URL to DMS browser-report functions
Const INIT_Dataset_Info_URL = "URL_Dataset_Details"
Const INIT_Analysis_Info_URL = "URL_Analysis_Details"
Const INIT_Experiment_Info_URL = "URL_Experiment_Details"

Dim FirstFill As String         'initial fill method
Dim sqlShowAll As String
Dim sqlShowNew As String
Dim sqlShowYear As String
Dim sqlShowPattern As String

Dim AnalysisTypeTable As String

Dim CurrStage As Long              'current visible stage

Dim MyInit As New InitFile         'object that deals with initialization file

Public InitFileName As String       'full path to init file
Public fAnalysis As FTICRAnalysis   'FTICR Analysis object
Attribute fAnalysis.VB_VarHelpID = -1

Public Event AnalysisDialogClose()

Private MTDBInd As Long                     'index of currently selected db
Private MTDBCnt As Long                     'count of mass tag databases
Private MTDBInfo() As udtMTDBInfoType       ' 0-based array

Private MTDBCntVisible As Long
Private MTDBNameListPointers() As Long      ' Used to display database names sorted properly; 0-based array

Dim TxtViewMax As Boolean               'controls the size of small text file viewer
Dim TxtViewOrigHeight As Long
Dim TxtViewOrigTop As Long

'URLs to access DMS browser-reports
'add name of dataset/analysis and ShellExecute it with browser
Dim BaseURLDataset As String
Dim BaseURLAnalysis As String
Dim BaseURLExperiment As String

Private Sub HighlightDBByName(ByVal strTextToFind As String, ByVal intIndexStart)
    
    Dim i As Integer
    Dim intCharLoc As Integer
    
    If Len(strTextToFind) > 0 And lstOrgMTDBNames.ListCount > 0 Then
        ' Step through lstOrgMTDBNames and find the first to contain strTextToFind (starting at index intIndexStart)
            
        strTextToFind = LCase(strTextToFind)
        
        If intIndexStart < 0 Then
            intIndexStart = lstOrgMTDBNames.ListCount - 1
        End If

        i = intIndexStart
        Do
            i = i + 1
            If i > lstOrgMTDBNames.ListCount - 1 Then
                i = 0
            End If

            intCharLoc = InStr(LCase(lstOrgMTDBNames.List(i)), strTextToFind)
            
            If intCharLoc > 0 Then
                lstOrgMTDBNames.ListIndex = i
                Exit Do
            End If

        Loop While i <> intIndexStart
        
    End If
    
End Sub

Private Sub PopulateDatabaseCombobox()
    Dim i As Long
    Dim blnShowFrozenDBs As Boolean
    Dim blnShowUnusedDBs As Boolean     ' Forced to False if blnShowFrozenDBs = False
    Dim strDatabaseNameSaved As String
    
On Error GoTo PopulateDatabaseComboboxErrorHandler

    If chkShowFrozenDBs.Value = vbChecked Then
        blnShowFrozenDBs = True
        blnShowUnusedDBs = (chkShowUnusedDBs.Value = vbChecked)
    Else
        blnShowFrozenDBs = False
        blnShowUnusedDBs = False
    End If
    
    If lstOrgMTDBNames.ListIndex >= 0 Then
        strDatabaseNameSaved = lstOrgMTDBNames.List(lstOrgMTDBNames.ListIndex)
    End If
    
    lstOrgMTDBNames.Clear
    
    MTDBCnt = UBound(MTDBInfo) + 1
    If MTDBCnt > 0 Then
        SortMTDBNameList MTDBInfo(), MTDBNameListPointers(), 0, MTDBCnt - 1, blnShowFrozenDBs, blnShowUnusedDBs
        
        MTDBCntVisible = UBound(MTDBNameListPointers) + 1
        
        For i = 0 To MTDBCntVisible - 1
            lstOrgMTDBNames.AddItem MTDBInfo(MTDBNameListPointers(i)).Name
            If MTDBInfo(MTDBNameListPointers(i)).Name = strDatabaseNameSaved Then
                lstOrgMTDBNames.ListIndex = lstOrgMTDBNames.ListCount - 1
            End If
        Next i
        
        If lstOrgMTDBNames.ListIndex < 0 And lstOrgMTDBNames.ListCount > 0 Then
            lstOrgMTDBNames.ListIndex = 0
        End If
        
        'set default options
        CurrStage = 0
        cmbFileType.ListIndex = FILE_TYPE_ANY
    End If
    
    Exit Sub

PopulateDatabaseComboboxErrorHandler:
Debug.Assert False

End Sub

Private Sub chkShowFrozenDBs_Click()
    EnableDisableControls
    PopulateDatabaseCombobox
End Sub

Private Sub chkShowUnusedDBs_Click()
    PopulateDatabaseCombobox
End Sub

Private Sub cmbAnalysisType_Click()
On Error Resume Next
With cmbAnalysisType
  fAnalysis.MD_Type = .ItemData(.ListIndex)
  fAnalysis.Desc_Type = .List(.ListIndex)
End With
End Sub


Private Sub cmbFileType_Click()
With cmbFileType
    Select Case .ListIndex
    Case FILE_TYPE_PEK
        lbFileList.Pattern = "*.pek*"
    Case FILE_TYPE_CSV
        lbFileList.Pattern = "*.csv*"
    Case FILE_TYPE_HDF
        lbFileList.Pattern = "*.h5"
    Case FILE_TYPE_ANY
        lbFileList.Pattern = "*.*"
    End Select
End With
End Sub

Private Sub cmdSearchForDB_Click()
    HighlightDBByName txtSearchForDB, lstOrgMTDBNames.ListIndex
End Sub

Private Sub lstOrgMTDBNames_Click()
'-------------------------------------------------
'initialize database connection with new selection
'of Mass Tags database
'-------------------------------------------------
UpdateInfoForSelectedDB
End Sub


Private Sub cmbYear_Click()
'--------------------------------------------------
'fills analysis list with files from specific year
'that of course depends on what is in sqlShowYear
'--------------------------------------------------
Dim sqlFull
Dim WhichYear As String
On Error Resume Next
WhichYear = cmbYear.Text
If Len(sqlShowYear) > 0 And Len(WhichYear) > 0 Then
   sqlFull = sqlShowYear & WhichYear
   FillAnalysisList sqlFull
End If
End Sub

Private Sub cmdAnalysisInfo_Click()
Dim AnalysisURL As String
On Error Resume Next
If fAnalysis.Job > 0 Then
   AnalysisURL = BaseURLAnalysis & fAnalysis.Job
   RunShellExecute "open", AnalysisURL, 0&, 0&, SW_SHOWNORMAL
Else
   MsgBox "Select ICR-2LS job first!", vbOKOnly, MyName
End If
End Sub

Private Sub cmdCancel_Click()
Set fAnalysis = Nothing
Me.Hide
RaiseEvent AnalysisDialogClose
End Sub


Private Sub cmdConfigure_Click()
'----------------------------------------
'displays advanced DB settings form
'----------------------------------------
Dim MyAdvancedDBForm As New frmAdvancedDB
On Error GoTo err_cmdConfigure
With MyAdvancedDBForm
    .Caption = "Organism Mass Tag Database"
    Set .ThisCN = fAnalysis.MTDB.cn
    CopyPairsCollection fAnalysis.MTDB.DBStuff, .ThisStuff
    .Show vbModal
    If .AcceptChanges Then Set fAnalysis.MTDB.cn = .ThisCN
End With
Unload MyAdvancedDBForm
Set MyAdvancedDBForm = Nothing
Exit Sub

err_cmdConfigure:
MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbOKOnly, MyName
End Sub


Private Sub cmdDatasetInfo_Click()
Dim DatasetURL As String
On Error Resume Next
If Len(fAnalysis.Dataset) > 0 Then
   DatasetURL = BaseURLDataset & fAnalysis.Dataset & "&" & "database_name=" & MTDBInfo(MTDBInd).Name
   RunShellExecute "open", DatasetURL, 0&, 0&, SW_SHOWNORMAL
Else
   MsgBox "Select ICR-2LS job first!", vbOKOnly, MyName
End If
End Sub

Private Sub cmdExperimentInfo_Click()
Dim ExperimentURL As String
On Error Resume Next
If fAnalysis.Job > 0 Then
   ExperimentURL = BaseURLExperiment & fAnalysis.Experiment
   RunShellExecute "open", ExperimentURL, 0&, 0&, SW_SHOWNORMAL
Else
   MsgBox "Select ICR-2LS job first!", vbOKOnly, MyName
End If
End Sub

Private Sub cmdGANETInfo_Click()
On Error Resume Next
If fAnalysis.Job > 0 Then
   MsgBox fAnalysis.GetJobInfoGANET, vbOKOnly, MyName
Else
   MsgBox "Select ICR-2LS job first!", vbOKOnly, MyName
End If
End Sub

Private Sub cmdInfo_Click()
'-----------------------------------------------
'display info about currently selected job
'-----------------------------------------------
If fAnalysis.Job <> 0 Then
   MsgBox fAnalysis.GetJobInfo, vbOKOnly, MyName
Else
   MsgBox "No item selected!", vbOKOnly, MyName
End If
End Sub


Private Sub cmdNext_Click()
'-----------------------------------------------
'controls stage changes
'-----------------------------------------------
'make current stage invisible
fraStage(CurrStage).Visible = False
CurrStage = (CurrStage + 1) Mod STAGE_COUNT
'make current stage visible
fraStage(CurrStage).Visible = True
NextPreviousOnOff

If CurrStage = 1 Then
    ' Populate lstICR2LSJobs
    FillDBInfoAndJobs
End If

'display current parameters if in right stage
If CurrStage = 3 Then txtParameters = fAnalysis.GetParameters
End Sub

Private Sub cmdOK_Click()
'pick current parameters
FillParametersFromTextBox
Me.Hide
RaiseEvent AnalysisDialogClose
End Sub

Private Sub cmdPrevious_Click()
'-----------------------------------------------
'controls stage changes
'-----------------------------------------------
If CurrStage = 3 Then FillParametersFromTextBox
If CurrStage < 1 Then CurrStage = 1

'make current stage invisible
fraStage(CurrStage).Visible = False
CurrStage = (CurrStage - 1) Mod STAGE_COUNT
'make current stage visible
fraStage(CurrStage).Visible = True
NextPreviousOnOff
End Sub


''Private Sub cmdPrintList_Click()
'''--------------------------------------------------
'''prints list of ICR-2LS analysis from the list
'''--------------------------------------------------
''Dim sList As String
''Dim sNow As String
''Dim sCalLine As String
''Dim i As Long
''On Error Resume Next
''
''With lstICR2LSJobs
''    If .ListCount > 0 Then
''       For i = 0 To .ListCount - 1
''           sList = sList & .ItemData(i) & Chr$(9) & .List(i) & vbCrLf
''       Next i
''       sCalLine = "---0000--( () )--0000------------------"
''       sNow = CStr(Now())
''       Printer.Print " "
''       Printer.Font.Name = "Courier New"
''       Printer.Font.Size = 9
''       Printer.Print " "
''       Printer.Print "         <____>                 "
''       Printer.Print "          o  o                  "
''       Printer.Print sCalLine
''       Printer.Print "|          UU                         |"
''       Printer.Print "|" & String$(Len(sCalLine) - 2, " ") & "|"
''       Printer.Print "|     ICR-2LS Analysis List           |"
''       Printer.Print "|" & String$(5, " ") & sNow & String(Len(sCalLine) - Len(sNow) - 7, " ") & "|"
''       Printer.Print "|" & Space(sCalLine - 2) & "|"
''       Printer.Print "---------------------------------------"
''       Printer.Print "        | |  | | "
''       Printer.Print "       0000  0000 "
''       Printer.Print
''       Printer.Print "Organism Mass Tags Database: " & MTDBInfo(MTDBInd), Name
''       Printer.Print "Description: " & MTDBInfo(MTDBInd).Description
''       Printer.Print
''       Printer.Print
''       Printer.Print "Job ID" & Chr$(9) & "Analysis Folder"
''       Printer.Print sList
''       Printer.EndDoc
''    Else
''       MsgBox "No items found in list!", vbOKOnly, MyName
''    End If
''End With
''End Sub

Private Sub cmdResetParameters_Click()
'-----------------------------------------------
'reset parameters settings
'-----------------------------------------------
txtParameters.Text = fAnalysis.GetParameters
End Sub

Private Sub cmdSearch_Click()
'--------------------------------------------------------
'fills analysis list with files whose name fits specified
'pattern; that depends on what is in sqlShowPattern
'--------------------------------------------------------
Dim sqlFull
Dim SearchWhat As String
On Error Resume Next
SearchWhat = Trim$(txtSearch.Text)
If Len(sqlShowPattern) > 0 Then
    sqlFull = sqlShowPattern
    If Len(SearchWhat) > 0 Then
        sqlFull = Replace(sqlShowPattern, "'%%'", "'%" & SearchWhat & "%'")
    End If
    FillAnalysisList sqlFull
End If
End Sub

Private Sub cmdSelectDatasetAnalysis_Click()
'-----------------------------------------------
'set stage for file selection
'-----------------------------------------------
UpdateDatasetPath
End Sub

Private Sub cmdSelectLockers_Click()
fAnalysis.MTDB.SelectLockers
End Sub

Private Sub cmdSelectMT_Click()
fAnalysis.MTDB.SelectMassTags InitFileName
End Sub

Private Sub cmdShowAll_Click()
If Len(sqlShowAll) > 0 Then FillAnalysisList sqlShowAll
End Sub

Private Sub cmdShowNew_Click()
If Len(sqlShowNew) > 0 Then FillAnalysisList sqlShowNew
End Sub

Private Sub cmdTICAlignment_Click()
On Error Resume Next
If fAnalysis.Job > 0 Then
   MsgBox fAnalysis.GetJobInfoTIC, vbOKOnly, MyName
Else
   MsgBox "Select ICR-2LS job first!", vbOKOnly, MyName
End If
End Sub

Private Sub lbFileList_Click()
fAnalysis.MD_File = lbFileList.FileName
lblSelectedFile.Caption = lbFileList.FileName
End Sub


Private Sub lbFileList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbFileList.Drag vbBeginDrag
End Sub

Private Sub lbFileList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbFileList.Drag vbEndDrag
End Sub

Private Sub Form_Load()
'-----------------------------------------------------------
'loads information from initialization file
'-----------------------------------------------------------
Dim Res As Long
Dim intIndex As Integer
On Error Resume Next

' Position the Frames
With fraStage(0)
    For intIndex = 1 To STAGE_COUNT - 1
        fraStage(intIndex).Left = .Left
        fraStage(intIndex).Top = .Top
        fraStage(intIndex).Visible = False
    Next intIndex
End With

' Populate the FileTypes listbox
With cmbFileType
    .Clear
    .AddItem "PEK files (*.pek*)", FILE_TYPE_PEK
    .AddItem "CSV files (*.csv)", FILE_TYPE_CSV
    .AddItem "HDF5 files (*.h5)", FILE_TYPE_HDF
    .AddItem "All files (*.*)", FILE_TYPE_ANY
End With

EnableDisableControls

'allow form to present itself
Me.Show
DoEvents

TxtViewOrigTop = txtShowTextFile.Top
TxtViewOrigHeight = txtShowTextFile.Height

Set fAnalysis = New FTICRAnalysis
fAnalysis.ProcessingType = fptMassMatch
MTDBInd = -1
' 12/12/2004 mem - Switched from using MT_Main to MTS_Master to retrieve DB info
Res = GetMTSMasterDirectoryData(InitFileName, MTDBInfo)
If Res <> 0 Then
   Set fAnalysis = Nothing
   MsgBox "Error loading Mass Tag databases information!" & vbCrLf & "Error: " & Res & " - " & Error(Res), vbOKOnly
   Me.Hide
   RaiseEvent AnalysisDialogClose
Else
    GetMTDBSchema
    PopulateDatabaseCombobox
End If

Call GetDMSURLs
End Sub


Private Function InitDBConnection() As Boolean
'--------------------------------------------------------
'initializes MT database connection from InitFileName
'returns True if everything cool; False otherwise
'--------------------------------------------------------
Dim Arg() As String
Dim ArgCnt As Long
Dim MTName As String, MTValue As String
Dim MTValuePos As Long
Dim nv As NameValue
Dim spGetDBSchemaVersion As String
Dim i As Long
On Error GoTo err_InitDBConnection

lblMTDBDesc.Caption = ""
If MTDBInd < 0 Then Exit Function
lblMTDBDesc.Caption = MTDBInfo(MTDBInd).Description & vbCrLf & "State: " & MTDBInfo(MTDBInd).DBState & vbCrLf & "Server: " & MTDBInfo(MTDBInd).Server

With fAnalysis.MTDB.cn
   If .State <> adStateClosed Then .Close
   .ConnectionString = MTDBInfo(MTDBInd).CnStr
   .ConnectionTimeout = 30
End With

With fAnalysis
    If MTDBInfo(MTDBInd).DBSchemaVersion = 0 Then
        On Error Resume Next
        spGetDBSchemaVersion = .MTDB.DBStuff.Item(INIT_Get_DB_Schema_Version).Value
        If Err Then
            spGetDBSchemaVersion = "GetDBSchemaVersion"
            Err.Clear
        End If
        
        MTDBInfo(MTDBInd).DBSchemaVersion = GetDBSchemaVersion(MTDBInfo(MTDBInd).CnStr, spGetDBSchemaVersion)
    End If
    .DB_Schema_Version = MTDBInfo(MTDBInd).DBSchemaVersion
End With

On Error GoTo err_InitDBConnection

If CurrStage > 1 Then
    FillDBInfoAndJobs
End If

lblStatus.Caption = "Database: " & MTDBInfo(MTDBInd).Name & " initialized."
InitDBConnection = True
Exit Function

err_InitDBConnection:
End Function

Private Sub lstICR2LSJobs_Click()
With lstICR2LSJobs
    fAnalysis.Job = .ItemData(.ListIndex)
    fAnalysis.MD_Reference_Job = fAnalysis.Job
    lblSelectionDatasetAnalysis.Caption = .List(.ListIndex)
    'fill all records for selected job
    fAnalysis.FillFADRecord
    fAnalysis.MD_File = ""
End With
UpdateDatasetPath
End Sub

Private Sub NextPreviousOnOff()
'----------------------------------------------------
'disable Previous if first stage; disable Nex if last
'----------------------------------------------------
If CurrStage = STAGE_COUNT - 1 Then
   cmdNext.Enabled = False
Else
   cmdNext.Enabled = True
End If
If CurrStage = 0 Then
   cmdPrevious.Enabled = False
Else
   cmdPrevious.Enabled = True
End If
If CurrStage > 0 Then
    cmdGANETInfo.Visible = True
Else
    cmdGANETInfo.Visible = False
End If

End Sub

Private Function FillAnalysisList(ByVal FillSQL As String) As Boolean
'--------------------------------------------------------------------
'retrieves rows and fills list; this procedures expects recordset
'with two fields; 1st field is PrimaryKey; 2nd is name to go in list
'--------------------------------------------------------------------
Dim rsDatasets As ADODB.Recordset
Dim ListRows() As Variant
Dim i As Long
Dim lngMaxIndex As Long
On Error GoTo err_FillAnalysisList

'retrieve records
fAnalysis.MTDB.cn.Open
Set rsDatasets = New ADODB.Recordset
Set rsDatasets.ActiveConnection = fAnalysis.MTDB.cn
rsDatasets.CursorLocation = adUseClient
rsDatasets.CursorType = adOpenForwardOnly
Set rsDatasets = fAnalysis.MTDB.cn.Execute(FillSQL)

' Need On Error Resume Next in case no rows were returned via the above .Execute() statement
On Error Resume Next
ListRows = rsDatasets.GetRows
lngMaxIndex = -1
lngMaxIndex = UBound(ListRows, 2)

On Error GoTo err_FillAnalysisList
rsDatasets.Close
fAnalysis.MTDB.cn.Close
'fill the list(GetRows returns transposed rows/columns)
lstICR2LSJobs.Clear
With lstICR2LSJobs
   For i = 0 To lngMaxIndex
     .AddItem ListRows(1, i)
     .ItemData(.NewIndex) = ListRows(0, i)
   Next i
End With
If lstICR2LSJobs.ListCount >= 1 Then
    lstICR2LSJobs.Selected(0) = True
Else
    fAnalysis.Job = 0
    fAnalysis.MD_Reference_Job = 0
    lblSelectionDatasetAnalysis.Caption = ""
    'fill all records for selected job
    fAnalysis.FillFADRecord
    fAnalysis.MD_File = ""
    UpdateDatasetPath
End If
FillAnalysisList = True

exit_FillAnalysisList:
Set rsDatasets = Nothing
With fAnalysis.MTDB.cn
    If .State <> adStateClosed Then .Close
End With
Exit Function

err_FillAnalysisList:
LogMessages "Error in FillAnalysisList: " & Err.Number & " - " & Err.Description
Resume exit_FillAnalysisList:
End Function

Private Function FillAnalysisTypesCombo() As Boolean
'-----------------------------------------------------------------------
'retrieves rows and fills list; this procedures expects recordset
'with two fields; 1st field is PrimaryKey; 2nd is name to go in the list
'-----------------------------------------------------------------------
Dim rsDatasets As ADODB.Recordset
Dim ListRows()
Dim i As Long
Dim FillSQL As String
On Error GoTo err_FillAnalysisTypesCombo

'retrieve records
FillSQL = "SELECT * FROM " & AnalysisTypeTable
fAnalysis.MTDB.cn.Open
Set rsDatasets = New ADODB.Recordset
Set rsDatasets.ActiveConnection = fAnalysis.MTDB.cn
rsDatasets.CursorLocation = adUseClient
rsDatasets.CursorType = adOpenForwardOnly
Set rsDatasets = fAnalysis.MTDB.cn.Execute(FillSQL)
ListRows = rsDatasets.GetRows
'fill combo(GetRows returns transposed rows/columns)
cmbAnalysisType.Clear
With cmbAnalysisType
   For i = 0 To UBound(ListRows, 2)
     .AddItem ListRows(1, i)
     .ItemData(.NewIndex) = ListRows(0, i)
   Next i
   If .ListCount > 0 Then .ListIndex = 0
End With
FillAnalysisTypesCombo = True

exit_FillAnalysisTypesCombo:
Set rsDatasets = Nothing
With fAnalysis.MTDB.cn
    If .State <> adStateClosed Then .Close
End With
Exit Function

err_FillAnalysisTypesCombo:
Resume exit_FillAnalysisTypesCombo
End Function

Private Sub FillDBInfoAndJobs()
    Dim blnSuccess As Boolean
    
    'can not proceed with no data access statement
    If Len(FirstFill) <= 0 Then Exit Sub
    AnalysisTypeTable = fAnalysis.MTDB.DBStuff.Item(INIT_t_Analysis_Type).Value
    
    'can not proceede without analysis type table
    If Len(AnalysisTypeTable) <= 0 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    blnSuccess = FillAnalysisTypesCombo()
    If blnSuccess Then
        blnSuccess = FillAnalysisList(FirstFill)
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub FillParametersFromFile()
'---------------------------------------------------------
'retrieves parameters from parameters section of init file
'---------------------------------------------------------
Dim SecCnt As Long
Dim SecNames() As String
Dim Pars() As String
Dim ParsCnt As Long
Dim i As Long
On Error Resume Next

Select Case fAnalysis.ProcessingType
Case fptMassLock
  ParsCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_Parameters_Lock, Pars())
  If ParsCnt > 0 Then FillParameters Pars(), False
Case fptMassMatch
  SecCnt = MyInit.GetSectionNames(InitFileName, SecNames())
  For i = 0 To SecCnt - 1
    If InStr(1, SecNames(i), MyGl.SECTION_Parameters_Match_Any) > 0 Then
       ParsCnt = MyInit.GetSection(InitFileName, SecNames(i), Pars())
       If ParsCnt > 0 Then FillParameters Pars(), False
    End If
  Next i
End Select
End Sub

Private Sub FillParametersFromTextBox()
'---------------------------------------------
'retrieves parameters from parameters text box
'---------------------------------------------
Dim Pars() As String
Dim ParsCnt As Long
On Error Resume Next

Pars() = Split(Trim$(txtParameters.Text), vbCrLf)
ParsCnt = UBound(Pars) + 1
If ParsCnt > 0 Then FillParameters Pars(), True
End Sub


Private Sub FillParameters(Pars() As String, _
                           ByVal Destroy As Boolean)
'---------------------------------------------------
'fills parameters from array of strings
'---------------------------------------------------
Dim ParsCnt As Long
Dim ValuePos As Long
Dim ParName As String
Dim ParValue As String
Dim nv As NameValue
Dim i As Long
On Error Resume Next
ParsCnt = UBound(Pars) + 1
If ParsCnt > 0 Then
   With fAnalysis
     If Destroy Then .DestroyParameters
     For i = 0 To ParsCnt - 1
       If Len(Trim$(Pars(i))) > 0 Then
         ValuePos = InStr(1, Pars(i), MyGl.INIT_Value)
         If ValuePos > 0 Then
            ParName = Trim(Left$(Pars(i), ValuePos - 1))
            ParValue = Trim$(Right$(Pars(i), Len(Pars(i)) - ValuePos))
         Else     'everything is a name
            ParName = Trim(Pars(i))
            ParValue = ""
         End If
         Set nv = New NameValue
         nv.Name = ParName
         nv.Value = ParValue
         .Parameters.Add nv, nv.Name
       End If
     Next i
   End With
End If
End Sub


Private Sub GetMTDBSchema()
'-----------------------------------------------------
'retrieves schema information from initialization file
'this is same for all Mass Tags databases
'also enables user to select subset of all mass tags
'to work with
'-----------------------------------------------------
Dim Arg() As String
Dim ArgCnt As Long
Dim MTName As String, MTValue As String
Dim MTValuePos As Long
Dim nv As NameValue
Dim i As Long
On Error Resume Next

ArgCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_MTDB_Schema, Arg())
If ArgCnt > 0 Then
   For i = 0 To ArgCnt - 1
       MTValuePos = InStr(1, Arg(i), MyGl.INIT_Value)
       If MTValuePos > 0 Then
          MTName = Trim(Left$(Arg(i), MTValuePos - 1))
          MTValue = Trim$(Right$(Arg(i), Len(Arg(i)) - MTValuePos))
       Else     'everything is a name
          MTName = Trim(Arg(i))
          MTValue = ""
       End If
       If Len(MTName) > 0 Then
          Set nv = New NameValue
          nv.Name = MTName
          nv.Value = MTValue
          fAnalysis.MTDB.DBStuff.Add nv, nv.Name
       End If   'do nothing if name is missing
   Next i
End If
'do some checks and enable/disable some functions
'check what list filling procedures are available
With fAnalysis.MTDB.DBStuff
   sqlShowAll = .Item(INIT_Fill_All).Value
   If Len(sqlShowAll) > 0 Then
      cmdShowAll.Enabled = True
      FirstFill = sqlShowAll
   Else
      cmdShowAll.Enabled = False
   End If
   sqlShowNew = .Item(INIT_Fill_New).Value
   If Len(sqlShowNew) > 0 Then
      cmdShowNew.Enabled = True
      If Len(FirstFill) <= 0 Then FirstFill = sqlShowNew
   Else
      cmdShowNew.Enabled = False
   End If
   sqlShowYear = .Item(INIT_Fill_Year).Value
   If Len(sqlShowYear) > 0 Then
      cmbYear.Enabled = True
      If Len(FirstFill) <= 0 Then FirstFill = sqlShowYear & " = " & Year(Now())
   Else
      cmbYear.Enabled = False
   End If
   sqlShowPattern = .Item(INIT_Fill_Pattern).Value
   If Len(sqlShowPattern) > 0 Then
      cmdSearch.Enabled = True
      If Len(FirstFill) <= 0 Then FirstFill = sqlShowPattern & Chr$(34) & "%" & Chr$(34)
   Else
      cmdSearch.Enabled = False
   End If
End With
'retrieve parameter information
FillParametersFromFile
End Sub

Private Sub txtShowTextFile_DblClick()
'-----------------------------------------------------------------
'toggle between larger and smaller view for small text file viewer
'-----------------------------------------------------------------
If TxtViewMax Then      'shrink it
   txtShowTextFile.Top = TxtViewOrigTop
   txtShowTextFile.Height = TxtViewOrigHeight
Else                    'enlarge it
   txtShowTextFile.Top = cmbFileType.Top
   txtShowTextFile.Height = TxtViewOrigHeight + (TxtViewOrigTop - cmbFileType.Top)
End If
TxtViewMax = Not TxtViewMax
End Sub

Private Sub txtShowTextFile_DragDrop(Source As Control, X As Single, Y As Single)
Dim FName As String
Dim fso As New FileSystemObject
On Error GoTo err_DD
If TypeOf Source Is FileListBox Then
   FName = fAnalysis.Desc_DataFolder & Source.FileName
   If FileLen(FName) <= SMALL_FILE_MAX_SIZE Then
      txtShowTextFile.Text = fso.OpenTextFile(FName, ForReading).ReadAll
   Else
      MsgBox "Only small files (up to 30KB) could be viewed in this panel!", vbOKOnly, App.Title
   End If
End If
Exit Sub

err_DD:
MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, App.Title
End Sub

Private Sub EnableDisableControls()
    chkShowUnusedDBs.Enabled = (chkShowFrozenDBs.Value = vbChecked)
End Sub

Public Sub GetDMSURLs()
'------------------------------------------------------
'extracts URLs information from the initialization file
'------------------------------------------------------
Dim Arg() As String
Dim ArgCnt As Long
Dim MTName As String, MTValue As String
Dim MTValuePos As Long
Dim i As Long
On Error Resume Next

ArgCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_URL, Arg())
If ArgCnt > 0 Then
   For i = 0 To ArgCnt - 1
       MTValuePos = InStr(1, Arg(i), MyGl.INIT_Value)
       If MTValuePos > 0 Then
          MTName = Trim(Left$(Arg(i), MTValuePos - 1))
          MTValue = Trim$(Right$(Arg(i), Len(Arg(i)) - MTValuePos))
       Else     'everything is a name
          MTName = Trim(Arg(i))
          MTValue = ""
       End If
       Select Case MTName
       Case INIT_Dataset_Info_URL
            BaseURLDataset = MTValue
       Case INIT_Analysis_Info_URL
            BaseURLAnalysis = MTValue
       Case INIT_Experiment_Info_URL
            BaseURLExperiment = MTValue
       End Select
   Next i
End If
If Len(BaseURLDataset) > 0 Then
   cmdDatasetInfo.Enabled = True
End If
If Len(BaseURLAnalysis) > 0 Then
   cmdAnalysisInfo.Enabled = True
End If
If Len(BaseURLExperiment) > 0 Then
   cmdExperimentInfo.Enabled = True
End If
End Sub

Private Sub UpdateDatasetPath()
    Dim intExtensionIndex As Integer
    Dim intFileIndex As Integer
    Dim intCompareLen As Integer
    Dim intBestIndex As Integer
    
    Dim strTestExt As String
    Dim strTestFilename As String
    
    Dim intPreferredExtensionCount As Integer
    Dim strPreferredExtensions(5) As String
    
    intPreferredExtensionCount = 5
    strPreferredExtensions(0) = "_isos.csv"
    strPreferredExtensions(1) = "_ic.pek"
    strPreferredExtensions(2) = "_s.pek"
    strPreferredExtensions(3) = ".pek"
    strPreferredExtensions(4) = "DeCal.pek-3"
    strPreferredExtensions(5) = ".pek-3"

On Error GoTo PathSetErrorHandler
    ' Update the path for the File list control
    lbFileList.Path = fAnalysis.Desc_DataFolder
    
    ' Now highlight the most logical .PEK or .CSV file
    With lbFileList
        If .ListCount > 0 Then
            intBestIndex = -1
            For intExtensionIndex = 0 To intPreferredExtensionCount - 1
                strTestExt = LCase(strPreferredExtensions(intExtensionIndex))
                intCompareLen = Len(strTestExt)
                For intFileIndex = 0 To .ListCount - 1
                    strTestFilename = .List(intFileIndex)
                    If Len(strTestFilename) > intCompareLen Then
                        If Right(LCase(strTestFilename), intCompareLen) = strTestExt Then
                            intBestIndex = intFileIndex
                            Exit For
                        End If
                    End If
                Next intFileIndex
                If intBestIndex >= 0 Then Exit For
            Next intExtensionIndex
        
            .ListIndex = intBestIndex
        End If
    End With
    Exit Sub

PathSetErrorHandler:
    If Err.Number = 76 Then
        MsgBox "Be sure you have read-access to this server and folder, and that the folder exists:" & vbCrLf & fAnalysis.Desc_DataFolder, vbExclamation + vbOKOnly, "Path/file access error"
    Else
        MsgBox "Error occurred while setting the dataset folder path: " & vbCrLf & fAnalysis.Desc_DataFolder & vbCrLf & "Error: " & Err.Description, vbExclamation + vbOKOnly, "Path/file access error"
    End If

End Sub

Private Sub UpdateInfoForSelectedDB()
    
    If lstOrgMTDBNames.ListIndex < 0 Then Exit Sub
    
On Error GoTo UpdateInfoForSelectedDBErrorHandler
    
    MTDBInd = MTDBNameListPointers(lstOrgMTDBNames.ListIndex)
    Me.MousePointer = vbHourglass
    cmdNext.Enabled = False
    cmdSelectMT.Enabled = False
    cmdOK.Enabled = False
    
    If Not InitDBConnection() Then
       lblStatus.Caption = "Error retrieving database connection information!"
    End If

ExitUpdateInfoForSelectedDB:
    Me.MousePointer = vbDefault
    cmdNext.Enabled = True
    cmdSelectMT.Enabled = True
    cmdOK.Enabled = True
    
    Exit Sub
    
UpdateInfoForSelectedDBErrorHandler:
    Debug.Print "Error in UpdateInfoForSelectedDB: " & Err.Description
    Debug.Assert False
    Resume ExitUpdateInfoForSelectedDB
    
End Sub

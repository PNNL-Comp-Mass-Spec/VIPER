VERSION 5.00
Begin VB.Form frmSearchMT_UMC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Mass Tags DB - UMC"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMods 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modifications"
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   6735
      Begin VB.CheckBox chkPEO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PEO"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkICATLt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICAT d0"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkICATHv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICAT d8"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkAlkylation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alkylation"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         ToolTipText     =   "Check to add the alkylation mass correction below to all mass tag masses (added to each cys residue)"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtAlkylationMWCorrection 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Text            =   "57.0215"
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame fraOptionFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Index           =   49
         Left            =   4460
         TabIndex        =   31
         Top             =   360
         Width           =   1095
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fixed"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   750
         End
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dynamic"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   525
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Mod Type:"
            Height          =   255
            Index           =   100
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.Frame fraOptionFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Index           =   47
         Left            =   5760
         TabIndex        =   35
         Top             =   360
         Width           =   800
         Begin VB.OptionButton optN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N15"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   38
            Top             =   525
            Width           =   700
         End
         Begin VB.OptionButton optN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N14"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   240
            Value           =   -1  'True
            Width           =   700
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "N Type:"
            Height          =   255
            Index           =   103
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.TextBox txtResidueToModifyMass 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   30
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cboResidueToModify 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alkylation mass:"
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   720
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1320
         X2              =   1320
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4440
         X2              =   4440
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5640
         X2              =   5640
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mass (Da):"
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Residue to modify:"
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSearchAllUMCs 
      Caption         =   "Search All UMC's"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Search UMC's"
      Top             =   560
      Width           =   1815
   End
   Begin VB.CheckBox chkUpdateGelDataWithSearchResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update data in current file with results of search"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   560
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame fraMWField 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Molecular Mass Field"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A&verage"
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   80
         TabIndex        =   6
         Top             =   540
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   80
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame fraNET 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NET  Calculation"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   4455
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "First choice"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Use NET calculated only from Sequest ""first choice"" peptides"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All results"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Use NET calculated from all peptides of mass tags"
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   660
         Width           =   2775
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   19
         Text            =   "0.1"
         Top             =   1040
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "T&olerance"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   1060
         Width           =   855
      End
   End
   Begin VB.Frame fraMWTolerance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1215
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   2175
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   160
         TabIndex        =   10
         Text            =   "10"
         Top             =   640
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerance"
         Height          =   255
         Left            =   160
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   560
      Width           =   1815
   End
   Begin VB.Label lblETType 
      BackStyle       =   0  'Transparent
      Caption         =   "Generic NET"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   475
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Status of the Mass Tag database"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuF 
      Caption         =   "&Function"
      Begin VB.Menu mnuFSearchAll 
         Caption         =   "Search All UMCs"
      End
      Begin VB.Menu mnuFSearchPaired 
         Caption         =   "Search Paired UMCs"
      End
      Begin VB.Menu mnuFSearchNonPaired 
         Caption         =   "Search Non-paired UMCs"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFReport 
         Caption         =   "Report Results (Old Format)"
      End
      Begin VB.Menu mnuFReportWithORFs 
         Caption         =   "Report Results (New Format)"
      End
      Begin VB.Menu mnuFReportIncludeORFs 
         Caption         =   "Include ORFs in Report"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExpMTDB 
         Caption         =   "Export Results To &Mass Tags DB"
      End
      Begin VB.Menu mnuFSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuMT 
      Caption         =   "&Mass Tags"
      Begin VB.Menu mnuMTLoadMT 
         Caption         =   "Load Mass Tags DB"
      End
      Begin VB.Menu mnuMTLoadLegacy 
         Caption         =   "Load Legacy MT DB"
      End
      Begin VB.Menu mnuMTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTStatus 
         Caption         =   "Mass Tags Status"
      End
   End
   Begin VB.Menu mnuETDummy 
      Caption         =   "&Elution Time"
      Begin VB.Menu mnuET 
         Caption         =   "&Generic NET"
         Index           =   0
      End
      Begin VB.Menu mnuET 
         Caption         =   "&TIC Fit NET"
         Index           =   1
      End
      Begin VB.Menu mnuET 
         Caption         =   "G&ANET"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSearchMT_UMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is UMC identification - pairs are here just to distinguish
'which UMC to include in search
'---------------------------------------------------------------
'Elution is not corrected for N15 versions of peptides (???)
'When looking for N14; UMCs that are heavy members of pairs only
'are not search; neither are UMCs light only pair members when
'N15 search is performed
'---------------------------------------------------------------
'created: 10/10/2002 nt
'last modified: 10/17/2002 nt
'---------------------------------------------------------------
Option Explicit

Private Const MASS_PRECISION = 6
Private Const FIT_PRECISION = 3
Private Const NET_PRECISION = 5

Const MOD_TKN_NONE = "none"
Const MOD_TKN_PEO = "PEO"
Const MOD_TKN_ICAT_D0 = "ICAT_D0"
Const MOD_TKN_ICAT_D8 = "ICAT_D8"
Const MOD_TKN_ALK = "ALK"
Const MOD_TKN_N14 = "N14"
Const MOD_TKN_N15 = "N15"
Const MOD_TKN_RES_MOD = "RES_MOD"
Const MOD_TKN_MT_MOD = "MT_MOD"

Const SEARCH_N14 = 0
Const SEARCH_N15 = 1

Const MODS_FIXED = 0
Const MODS_DYNAMIC = 1

Const SEARCH_ALL = 0
Const SEARCH_PAIRED = 1
Const SEARCH_NON_PAIRED = 2

'if called with any positive number add that many points
Const MNG_RESET = 0
Const MNG_ERASE = -1
Const MNG_TRIM = -2
Const MNG_ADD_START_SIZE = -3

Const MNG_START_SIZE = 500

'in this case CallerID is a public property
Public CallerID As Long

Private bLoading As Boolean

Private OldSearchFlag As Long

'for faster search mass array will be sorted; therefore all other arrays
'has to be addressed indirectly (mMTNET(mMTInd(i))
Private mMTCnt                  'count of masses to search
Private mMTInd() As Long        'index(unique key)              ' 0-based array
Private mMTOrInd() As Long      'index of original mass tag (in AMT array)
Private mMTMWN14() As Double    'mass to look for N14
Private mMTMWN15() As Double    'mass to look for N15
Private mMTNET() As Double      'NET value
Private mMTMods() As String     'modification description

Private MWFastSearch As MWUtil

Private AlkMWCorrection As Double
Private N14N15 As Long                  'N14 0; N15 1
Private SearchType As Long

Private LastSearchTypeN14N15 As Long
Private NTypeStr As String

'following arrays are parallel to the UMCs
Private ClsCnt As Long              'this is not actually neccessary except
Private ClsStat() As Double         'to create nice reports
Private eClsPaired() As umcpUMCPairMembershipConstants      ' Keeps track of whether UMC is member of 1 or more pairs
                                
'mUMCMatchStats contains all possible identifications for all UMCs with scores
'as count of each identification hits within the UMC
Dim mMatchStatsCount As Long                                'count of UMC-ID matches
Dim mUMCMatchStats() As udtUMCMassTagMatchStats           ' 0-based array

' The following hold match stats for each individual UMC
Dim mgCurrIDCnt As Long
Dim mgCurrIDInd() As Long
Dim mgCurrIDScore() As Double
'NOTE: mg stands for "mare nostrum", which is the ancient latin name for the Mediterranean Sea

'Expression Evaluator variables for elution time calculation
Private MyExprEva As ExprEvaluator
Private VarVals()
Private MinFN As Long
Private MaxFN As Long

Private ExpAnalysisSPName As String             ' Stored procedure AddMatchMaking
Private ExpPeakSPName As String                 ' Stored procedure AddFTICRPeak
Private ExpUmcSPName As String                  ' Stored procedure AddFTICRUmc
Private ExpUmcMatchSPName As String             ' Stored procedure AddFTICRUmcMatch
Private ExpQuantitationDescription As String    ' Stored procedure AddQuantitationDescription

Private mKeyPressAbortProcess As Integer
Private mUsingDefaultGANET As Boolean

Private objMTDBNameLookupClass As mtdbMTNames
'

Private Function InitializeORFInfo(blnForceDataReload As Boolean) As Boolean
    ' Initializes objMTDBNameLookupClass
    ' Returns True if success, False if failure
    ' If the class has already been initialized, then does nothing, unless blnForceDataReload = True
    Dim blnSuccess As Boolean
    
    If Not objMTDBNameLookupClass Is Nothing Then
        If Not blnForceDataReload Then
            If objMTDBNameLookupClass.DataStatus = dsLoaded Then
                InitializeORFInfo = True
                Exit Function
            End If
        End If
        
        objMTDBNameLookupClass.DeleteData
        Set objMTDBNameLookupClass = Nothing
    End If
    
    Set objMTDBNameLookupClass = New mtdbMTNames
    
    With objMTDBNameLookupClass
        'loading mass tag names
        UpdateStatus "Loading ORF info"
        
        Me.MousePointer = vbHourglass
        
        .DBConnectionString = GelAnalysis(CallerID).MTDB.cn.ConnectionString
        .RetrieveSQL = GelAnalysis(CallerID).MTDB.DBStuff(SQL_GET_MT_Names).value
        If .FillData(Me) Then
           If .DataStatus = dsLoaded Then
                blnSuccess = True
            End If
        End If
        Me.MousePointer = vbDefault
    End With
    
    InitializeORFInfo = blnSuccess
End Function

Public Sub InitializeSearch()
'------------------------------------------------------------------------------------
'load mass tags database data if neccessary
'if CallerID is associated with mass tags database load that db if not already loaded
'if CallerID is not associated with mass tags database load legacy database
'------------------------------------------------------------------------------------
Dim blnSuccess As Boolean
Dim eResponse As VbMsgBoxResult

On Error Resume Next
Me.MousePointer = vbHourglass
If bLoading Then
   If GelAnalysis(CallerID) Is Nothing Then
      If AMTCnt > 0 Then    'something is loaded
         If Len(CurrMTDatabase) > 0 And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            'mass tags data; we dont know is it appropriate; warn user
            MsgBox "Current display is not associated with any Mass Tags database!" & vbCrLf _
                 & "However, mass tags are loaded from the Mass Tags database!" & vbCrLf _
                 & "If search should be performed on different Mass Tags DB you" & vbCrLf _
                 & "should close this dialog and establish link with other DB" & vbCrLf _
                 & "using Gel Parameters function from the Edit menu or select" & vbCrLf _
                 & "Mass Tags->Load Legacy MT DB on this dialog to load" & vbCrLf _
                 & "data from legacy database!", vbOKOnly, glFGTU
         End If
         lblMTStatus.Caption = "Mass tags count: " & AMTCnt
         
         ' Initialize the MT search object
         If Not CreateNewMTSearchObject() Then
            lblMTStatus.Caption = "Error creating search object!"
         End If
      
      Else                  'nothing is loaded
         If Len(glbPreferencesExpanded.LegacyAMTDBPath) > 0 And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            eResponse = MsgBox("Current display is not associated with any Mass Tags database.  Do you want to load the mass tags from the defined legacy mass tag database?" & vbCrLf & glbPreferencesExpanded.LegacyAMTDBPath, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load Legacy Mass Tags")
         Else
            eResponse = vbNo
         End If
         
         If eResponse = vbYes Then
            LoadLegacyMassTags
         Else
            Call Info_NoMTDBLink
            lblMTStatus.Caption = "No mass tags loaded"
         End If
      End If
   Else         'have to have mass tags database loaded
      Call LoadMTDB
   End If
   UpdateStatus "Generating UMC statistic ..."
   ClsCnt = UMCStatistics1(CallerID, ClsStat())
   UpdateStatus "Pairs Count: " & GelP_D_L(CallerID).PCnt
   Call mnuET_Click(etGANET)
   UpdateStatus "UMCs pairing status ..."
   blnSuccess = UpdateUMCsPairingStatus(CallerID, eClsPaired())
   UpdateStatus "Ready"
   'memorize number of scans (to be used with elution)
   MinFN = GelData(CallerID).DFFN(1)
   MaxFN = GelData(CallerID).DFFN(UBound(GelData(CallerID).DFFN))
   bLoading = False
End If
Me.MousePointer = vbDefault
End Sub

Private Sub LoadLegacyMassTags()

    '------------------------------------------------------------
    'load/reload mass tags
    '------------------------------------------------------------
    Dim eResponse As VbMsgBoxResult
    On Error Resume Next
    'ask user if it wants to replace legitimate Mass Tags DB with legacy DB
    If Not GelAnalysis(CallerID) Is Nothing Then
       eResponse = MsgBox("Current display is associated with Mass Tags database!" & vbCrLf _
                    & "Are you sure you want to use a legacy database for search?", vbYesNoCancel, glFGTU)
       If eResponse <> vbYes Then Exit Sub
    End If
    Me.MousePointer = vbHourglass
    If Len(glbPreferencesExpanded.LegacyAMTDBPath) > 0 Then
       If ConnectToAMT(False) Then
          If CreateNewMTSearchObject() Then
             lblMTStatus.Caption = "Loaded; Mass Tags Count: " & AMTCnt
          Else
             lblMTStatus.Caption = "Error creating search object!"
          End If
       Else
          lblMTStatus.Caption = "Error loading mass tags!"
       End If
    Else
       MsgBox "Path to legacy mass tags database not found!" & vbCrLf _
            & "In the main window, use Tools->Options, then go to the Miscellaneous tab and define 'AMT Database Location'.", vbOKOnly, glFGTU
    End If
    Me.MousePointer = vbDefault

End Sub

Public Sub SetAlkylationMWCorrection(dblMass As Double)
    txtAlkylationMWCorrection = dblMass
    AlkMWCorrection = dblMass
End Sub

Private Sub SetDBSearchModType(blnDynamicMods As Boolean)
    If blnDynamicMods Then
        optDBSearchModType(MODS_DYNAMIC).value = True
    Else
        optDBSearchModType(MODS_FIXED).value = True
    End If
    GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods = optDBSearchModType(MODS_DYNAMIC).value
End Sub

Public Sub SetDBSearchNType(blnUseN15 As Boolean)
    If blnUseN15 Then
        optN(1).value = True
        N14N15 = SEARCH_N15
    Else
        optN(0).value = True
        N14N15 = SEARCH_N14
    End If
End Sub

Private Function ShowOrSaveResultsOld(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True) As Long
'-------------------------------------
'report identified unique mass classes
' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
'
' Returns 0 if no error, the error number if an error
'-------------------------------------
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim FName As String
Dim i As Long
Dim strSepChar As String
Dim ErrCnt As Long
On Error GoTo err_mnuFReport_Click
If mMatchStatsCount > 0 Then
   UpdateStatus "Preparing results: 0 / " & Trim(mMatchStatsCount)
   mKeyPressAbortProcess = 0
   Me.MousePointer = vbHourglass
   Select Case LastSearchTypeN14N15
   Case SEARCH_N14
        NTypeStr = MOD_TKN_N14
   Case SEARCH_N15
        NTypeStr = MOD_TKN_N15
   End Select
   
   strSepChar = LookupDefaultSeparationCharacter()
   
   'temporary file for results output
   FName = GetTempFolder() & RawDataTmpFile
   If Len(strOutputFilePath) > 0 Then FName = strOutputFilePath
   Set ts = fso.OpenTextFile(FName, ForWriting, True)
   WriteReportHeader ts, strSepChar
   For i = 0 To mMatchStatsCount - 1
       ReportUMCID ts, i, strSepChar
       If i Mod 10 = 0 Then
          UpdateStatus "Preparing results: " & Trim(i) & " / " & Trim(mMatchStatsCount)
          If mKeyPressAbortProcess > 1 Then Exit For
       End If
   Next i
   ts.Close
   Set ts = Nothing
   If Len(strOutputFilePath) > 0 Then
       AddToAnalysisHistory CallerID, "Saved search results to disk: " & strOutputFilePath
   End If
   Me.MousePointer = vbDefault
   UpdateStatus ""
   If blnDisplayResults Then
        frmDataInfo.Tag = "UMC_MTID"
        frmDataInfo.Show vbModal
   End If
Else
   If blnDisplayResults Then MsgBox "No identification hits found!", vbOKOnly, glFGTU
End If

Set fso = Nothing
Exit Function

err_mnuFReport_Click:
If Err.Number = 70 And ErrCnt = 0 Then       'maybe we on a drive we don't have permission to write
   ChDir App.Path                            'change to local drive (if we run application from local)
   ChDrive App.Path
   ErrCnt = ErrCnt + 1                       'don't do this more than once
   Resume
Else
   LogErrors Err.Number, "frmSearchMT_UMC.mnuFReport_Click"
   ShowOrSaveResultsOld = Err.Number
End If
Set fso = Nothing
End Function

Public Function ShowOrSaveResultsByUMC(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True, Optional ByVal blnIncludeORFInfo As Boolean = True) As Long
'-------------------------------------
'report identified unique mass classes
' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
' If blnIncludeORFInfo = True, then attempts to connect to the database and retrieve the ORF information for each mass tag
'
' Returns 0 if no error, the error number if an error
'-------------------------------------
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim strLineOut As String
Dim FName As String
Dim mgInd As Long
Dim lngUMCIndexOriginal As Long                     'absolute index of UMC
Dim lngMassTagIndexPointer As Long                  'absolute index in mMT... arrays
Dim lngMassTagIndexOriginal As Long                 'absolute index in AMT... arrays
Dim strSepChar As String
Dim dblAMTMass As Double
Dim dblMassErrorPPM As Double
Dim lngAvgScanNumber As Long
Dim dblAvgGANET As Double
Dim dblGANETError As Double
Dim objORFNameFastSearch As New FastSearchArrayLong
Dim blnSuccess As Boolean

Dim lngPairIndex As Long
Dim dblExpressionRatio As Double
Dim dblExpressionRatioStDev As Double
Dim intExpressionRatioChargeStateBasisCount As Integer
Dim lngExpressionRatioMemberBasisCount As Long

Dim objP1IndFastSearch As FastSearchArrayLong
Dim objP2IndFastSearch As FastSearchArrayLong
Dim blnPairsPresent As Boolean

On Error GoTo ShowOrSaveResultsByUMCErrorHandler

If blnIncludeORFInfo Then
    UpdateStatus "Sorting ORF lookup arrays"
    If MTtoORFMapCount = 0 Then
        blnIncludeORFInfo = InitializeORFInfo(False)
    Else
        ' We can use MTIDMap(), ORFIDMap(), and ORFRefNames() to get the ORF name
        blnSuccess = objORFNameFastSearch.Fill(MTIDMap())
        Debug.Assert blnSuccess
    End If
End If

UpdateStatus "Preparing results: 0 / " & Trim(mMatchStatsCount)
mKeyPressAbortProcess = 0
Me.MousePointer = vbHourglass

'temporary file for results output
FName = GetTempFolder() & RawDataTmpFile
If Len(strOutputFilePath) > 0 Then FName = strOutputFilePath
Set ts = fso.OpenTextFile(FName, ForWriting, True)

Select Case LastSearchTypeN14N15
Case SEARCH_N14
     NTypeStr = MOD_TKN_N14
Case SEARCH_N15
     NTypeStr = MOD_TKN_N15
End Select

' Initialize the PairIndex lookup objects
blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)

strSepChar = LookupDefaultSeparationCharacter()

' UMCIndex; ScanStart; ScanEnd; AvgGANET; UMCMonoMW; UMCMWStDev; UMCMWMin; UMCMWMax; UMCAbundance; ClassStatsChargeBasis; ChargeStateMin; ChargeStateMax; UMCMemberCount; UMCMemberCountUsedForAbu; UMCAverageFit; PairIndex; ExpressionRatio; MultiMassTagHitCount; MassTagID; MassTagMonoMW; MassTagMods; MemberCountMatchingMassTag; MassErrorPPM; GANETError
strLineOut = "UMCIndex" & strSepChar & "ScanStart" & strSepChar & "ScanEnd" & strSepChar & "AvgGANET" & strSepChar & "UMCMonoMW" & strSepChar & "UMCMWStDev" & strSepChar & "UMCMWMin" & strSepChar & "UMCMWMax" & strSepChar & "UMCAbundance" & strSepChar
strLineOut = strLineOut & "ClassStatsChargeBasis" & strSepChar & "ChargeStateMin" & strSepChar & "ChargeStateMax" & strSepChar & "UMCMemberCount" & strSepChar & "UMCMemberCountUsedForAbu" & strSepChar & "UMCAverageFit" & strSepChar & "PairIndex" & strSepChar
strLineOut = strLineOut & "ExpressionRatio" & strSepChar & "ExpressionRatioStDev" & strSepChar & "ExpressionRatioChargeStateBasisCount" & strSepChar & "ExpressionRatioMemberBasisCount" & strSepChar
strLineOut = strLineOut & "MultiMassTagHitCount" & strSepChar & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagMods" & strSepChar & "MemberCountMatchingMassTag" & strSepChar & "MassErrorPPM" & strSepChar & "GANETError"
If blnIncludeORFInfo Then strLineOut = strLineOut & strSepChar & "MultiORFCount" & strSepChar & "ORFName"
ts.WriteLine strLineOut

For mgInd = 0 To mMatchStatsCount - 1
    lngUMCIndexOriginal = mUMCMatchStats(mgInd).UMCIndex
    
    If LastSearchTypeN14N15 = SEARCH_N14 Then
        ' N14
        dblAMTMass = mMTMWN14(mUMCMatchStats(mgInd).IDIndex)
    Else
        ' N15
        dblAMTMass = mMTMWN15(mUMCMatchStats(mgInd).IDIndex)
    End If
    lngMassTagIndexPointer = mMTInd(mUMCMatchStats(mgInd).IDIndex)
    lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
    
    With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
        lngAvgScanNumber = (.MinScan + .MaxScan) / 2
        
        dblAvgGANET = GelBody(CallerID).FNtoNET(lngAvgScanNumber)
        
        lngPairIndex = -1
        dblExpressionRatio = 0
        dblExpressionRatioStDev = 0
        intExpressionRatioChargeStateBasisCount = 0
        lngExpressionRatioMemberBasisCount = 0
        If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
            lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, objP1IndFastSearch, objP2IndFastSearch, (LastSearchTypeN14N15 = SEARCH_N15), dblExpressionRatio, dblExpressionRatioStDev, intExpressionRatioChargeStateBasisCount, lngExpressionRatioMemberBasisCount)
        End If
        
        strLineOut = lngUMCIndexOriginal & strSepChar & .MinScan & strSepChar & .MaxScan & strSepChar & Format(dblAvgGANET, "0.0000") & strSepChar & Round(.ClassMW, 6) & strSepChar
        strLineOut = strLineOut & Round(.ClassMWStD, 6) & strSepChar & .MinMW & strSepChar & .MaxMW & strSepChar & .ClassAbundance & strSepChar
        If GelUMC(CallerID).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
            strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge
        Else
            strLineOut = strLineOut & "0"
        End If
        strLineOut = strLineOut & strSepChar & ClsStat(lngUMCIndexOriginal, ustChargeMin) & strSepChar & ClsStat(lngUMCIndexOriginal, ustChargeMax) & strSepChar
        strLineOut = strLineOut & .ClassCount & strSepChar
        
        ' Include UMCMemberCountUsedForAbu
        If GelUMC(CallerID).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
            strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Count & strSepChar
        Else
            strLineOut = strLineOut & .ClassCount & strSepChar
        End If
        
        strLineOut = strLineOut & Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), 3) & strSepChar & Trim(lngPairIndex) & strSepChar
        strLineOut = strLineOut & Trim(dblExpressionRatio) & strSepChar & Trim(dblExpressionRatioStDev) & strSepChar & Trim(intExpressionRatioChargeStateBasisCount) & strSepChar & Trim(lngExpressionRatioMemberBasisCount) & strSepChar
        strLineOut = strLineOut & mUMCMatchStats(mgInd).MultiAMTHitCount & strSepChar
    
        dblMassErrorPPM = MassToPPM(.ClassMW - dblAMTMass, .ClassMW)
        dblGANETError = dblAvgGANET - AMTNET(lngMassTagIndexOriginal)
    End With
    
    strLineOut = strLineOut & AMTID(lngMassTagIndexOriginal) & strSepChar & Round(dblAMTMass, 6) & strSepChar & NTypeStr
    If Len(mMTMods(lngMassTagIndexPointer)) > 0 Then
        strLineOut = strLineOut & " " & mMTMods(lngMassTagIndexPointer)
    End If
    strLineOut = strLineOut & strSepChar & mUMCMatchStats(mgInd).MemberHitCount & strSepChar & Round(dblMassErrorPPM, 4) & strSepChar & Round(dblGANETError, NET_PRECISION)
    
    If Not blnIncludeORFInfo Then
        ts.WriteLine strLineOut
    Else
        WriteORFResults ts, strLineOut, CLngSafe(AMTID(lngMassTagIndexOriginal)), objORFNameFastSearch, strSepChar
    End If
    
    If mgInd Mod 25 = 0 Then
        UpdateStatus "Preparing results: " & Trim(mgInd) & " / " & Trim(mMatchStatsCount)
        If mKeyPressAbortProcess > 1 Then Exit For
    End If
Next mgInd
ts.Close

If Len(strOutputFilePath) > 0 Then
    AddToAnalysisHistory CallerID, "Saved search results to disk: " & strOutputFilePath
End If

Me.MousePointer = vbDefault
UpdateStatus ""
If blnDisplayResults Then
     frmDataInfo.Tag = "UMC_MTID"
     frmDataInfo.Show vbModal
End If

Set ts = Nothing
Set fso = Nothing
Set objORFNameFastSearch = Nothing
Exit Function

ShowOrSaveResultsByUMCErrorHandler:
ShowOrSaveResultsByUMC = Err.Number
LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.ShowOrSaveResultsByUMC"
Set fso = Nothing
End Function

Public Function StartSearchAll(Optional blnShowMessages As Boolean = True) As Long
' Returns the number of hits
Dim i As Long
Dim eResponse As VbMsgBoxResult
Dim strModMassDescription As String

On Error Resume Next

mKeyPressAbortProcess = 0
If mMatchStatsCount > 0 Then    'something already identified
   If blnShowMessages Then
      eResponse = MsgBox("Identification results already exist! If you continue current findings will be lost! Continue?", vbOKCancel, glFGTU)
   Else
      eResponse = vbOK
   End If
   If eResponse <> vbOK Then Exit Function
   Call DestroyIDStructures
End If
cmdSearchAllUMCs.Visible = False
SearchType = SEARCH_ALL
CheckNETEquationStatus

If PrepareMTArrays() Then
   For i = 0 To ClsCnt - 1
       If i Mod 25 = 0 Then
          UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
          If mKeyPressAbortProcess > 1 Then Exit For
       End If
       Call SearchUMC(i)
   Next i
   LastSearchTypeN14N15 = N14N15
   
   UpdateStatus "UMC - Mass Tag ID Cnt: " & mMatchStatsCount
    
    With GelSearchDef(CallerID).AMTSearchMassMods
        If .PEO Then
            GelAnalysis(CallerID).MD_Type = stLabeledPEO
        ElseIf .ICATd0 Then
            GelAnalysis(CallerID).MD_Type = stLabeledICATD0
        ElseIf .ICATd8 Then
            GelAnalysis(CallerID).MD_Type = stLabeledICATD8
        Else
            GelAnalysis(CallerID).MD_Type = stStandardIndividual
        End If
    End With
    
    If mKeyPressAbortProcess <= 1 Then
        'MonroeMod
        GelSearchDef(CallerID).AMTSearchOnUMCs = samtDef
        AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched all UMC's for mass tags (searched ion by ion, only examining ions belonging to UMC's; however, all members of a UMC are assigned all matches found for any member of the UMC)", mMatchStatsCount, 0, samtDef, False)
    
        strModMassDescription = ConstructMassTagModMassDescription(GelSearchDef(CallerID).AMTSearchMassMods)
        If Len(strModMassDescription) > 0 Then
            AddToAnalysisHistory CallerID, strModMassDescription
        End If
    End If
Else
   UpdateStatus "Error searching for matches"
End If

If mKeyPressAbortProcess > 1 Then
    UpdateStatus "Search aborted."
Else
    If chkUpdateGelDataWithSearchResults Then
        ' Store the search results in the gel data
        If mMatchStatsCount > 0 Then RecordSearchResultsInData
        UpdateStatus "UMC - Mass Tag ID Cnt: " & mMatchStatsCount
    End If
End If
cmdSearchAllUMCs.Visible = True
StartSearchAll = mMatchStatsCount
End Function

Public Function StartSearchPaired(Optional blnShowMessages As Boolean = True) As Long
' Returns the number of hits
Dim i As Long
Dim eResponse As VbMsgBoxResult
Dim strModMassDescription As String

On Error Resume Next

mKeyPressAbortProcess = 0
If mMatchStatsCount > 0 Then    'something already identified
   If blnShowMessages Then
      eResponse = MsgBox("Identification results already exist! If you continue current findings will be lost! Continue?", vbOKCancel, glFGTU)
   Else
      eResponse = vbOK
   End If
   If eResponse <> vbOK Then Exit Function
   Call DestroyIDStructures
End If
cmdSearchAllUMCs.Visible = False
SearchType = SEARCH_PAIRED
CheckNETEquationStatus

If PrepareMTArrays() Then
   For i = 0 To ClsCnt - 1
       If eClsPaired(i) <> umcpNone Then
          ' MonroeMod: Added i Mod 25
          If i Mod 25 = 0 Then
              UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
              If mKeyPressAbortProcess > 1 Then Exit For
          End If
          Call SearchUMC(i)
       End If
   Next i
   LastSearchTypeN14N15 = N14N15
   
   UpdateStatus "UMC - Mass Tag ID Cnt: " & mMatchStatsCount
    
    If GelAnalysis(CallerID).MD_Type = stNotDefined Or GelAnalysis(CallerID).MD_Type = stStandardIndividual Then
        ' Only update MD_Type if it is currently stStandardIndividual
        GelAnalysis(CallerID).MD_Type = stPairsN14N15
    End If
    
    If mKeyPressAbortProcess <= 1 Then
        'MonroeMod
        GelSearchDef(CallerID).AMTSearchOnUMCs = samtDef
        AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched paired UMC's for mass tags", mMatchStatsCount, 0, samtDef, False)
    
        strModMassDescription = ConstructMassTagModMassDescription(GelSearchDef(CallerID).AMTSearchMassMods)
        If Len(strModMassDescription) > 0 Then
            AddToAnalysisHistory CallerID, strModMassDescription
        End If
    End If
Else
   UpdateStatus "Error searching for matches"
End If

If mKeyPressAbortProcess > 1 Then
    UpdateStatus "Search aborted."
Else
    If chkUpdateGelDataWithSearchResults Then
        ' Store the search results in the gel data
        If mMatchStatsCount > 0 Then RecordSearchResultsInData
        UpdateStatus "UMC - Mass Tag ID Cnt: " & mMatchStatsCount
    End If
End If
cmdSearchAllUMCs.Visible = True
StartSearchPaired = mMatchStatsCount
End Function

Public Function StartSearchNonPaired(Optional blnShowMessages As Boolean = True) As Long
' Returns the number of hits
Dim i As Long
Dim eResponse As VbMsgBoxResult
Dim strModMassDescription As String

On Error Resume Next

mKeyPressAbortProcess = 0
If mMatchStatsCount > 0 Then    'something already identified
   If blnShowMessages Then
      eResponse = MsgBox("Identification results already exist! If you continue current findings will be lost! Continue?", vbOKCancel, glFGTU)
   Else
      eResponse = vbOK
   End If
   If eResponse <> vbOK Then Exit Function
   Call DestroyIDStructures
End If
cmdSearchAllUMCs.Visible = False
SearchType = SEARCH_NON_PAIRED
CheckNETEquationStatus

If PrepareMTArrays() Then
   For i = 0 To ClsCnt - 1
       If eClsPaired(i) = umcpNone Then
          ' MonroeMod: Added i Mod 25
          If i Mod 25 = 0 Then
             UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
             If mKeyPressAbortProcess > 1 Then Exit For
          End If
          Call SearchUMC(i)
       End If
   Next i
   LastSearchTypeN14N15 = N14N15
   
   UpdateStatus "UMC - Mass Tag ID Cnt: " & mMatchStatsCount
    
    With GelSearchDef(CallerID).AMTSearchMassMods
        If .PEO Then
            GelAnalysis(CallerID).MD_Type = stLabeledPEO
        ElseIf .ICATd0 Then
            GelAnalysis(CallerID).MD_Type = stLabeledICATD0
        ElseIf .ICATd8 Then
            GelAnalysis(CallerID).MD_Type = stLabeledICATD8
        Else
            GelAnalysis(CallerID).MD_Type = stStandardIndividual
        End If
    End With
    
    If mKeyPressAbortProcess <= 1 Then
        'MonroeMod
        GelSearchDef(CallerID).AMTSearchOnUMCs = samtDef
        AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched non-paired UMC's for mass tags", mMatchStatsCount, 0, samtDef, False)
    
        strModMassDescription = ConstructMassTagModMassDescription(GelSearchDef(CallerID).AMTSearchMassMods)
        If Len(strModMassDescription) > 0 Then
            AddToAnalysisHistory CallerID, strModMassDescription
        End If
    End If
Else
   UpdateStatus "Error searching for matches"
End If

If mKeyPressAbortProcess > 1 Then
    UpdateStatus "Search aborted."
Else
    If chkUpdateGelDataWithSearchResults Then
        ' Store the search results in the gel data
        If mMatchStatsCount > 0 Then RecordSearchResultsInData
        UpdateStatus "UMC - Mass Tag ID Cnt: " & mMatchStatsCount
    End If
End If
cmdSearchAllUMCs.Visible = True
StartSearchNonPaired = mMatchStatsCount
End Function

Private Sub cboResidueToModify_Click()
    If cboResidueToModify.List(cboResidueToModify.ListIndex) = glPHOSPHORYLATION Then
        txtResidueToModifyMass = Trim(glPHOSPHORYLATION_Mass)
    Else
        ' For safety reasons, reset txtResidueToModifyMass to "0"
        txtResidueToModifyMass = "0"
    End If
End Sub

Private Sub chkAlkylation_Click()
    If cChkBox(chkAlkylation) And CDblSafe(txtAlkylationMWCorrection) <= 0 Then
        txtAlkylationMWCorrection = glALKYLATION
        AlkMWCorrection = glALKYLATION
    End If
End Sub

Private Sub cmdCancel_Click()
    mKeyPressAbortProcess = 2
    KeyPressAbortProcess = 2
End Sub

Private Sub cmdSearchAllUMCs_Click()
StartSearchAll
End Sub

Private Sub Form_Activate()
InitializeSearch
End Sub

Private Sub Form_Load()
'----------------------------------------------------
'load search settings and initializes controls
'----------------------------------------------------

Dim intIndex As Integer

On Error GoTo FormLoadErrorHandler

bLoading = True
If IsWinLoaded(TrackerCaption) Then Unload frmTracker
If CallerID >= 1 And CallerID <= UBound(GelBody) Then samtDef = GelSearchDef(CallerID).AMTSearchOnUMCs
'set current Search Definition values
With samtDef
    txtMWTol.Text = .MWTol
    optMWField(.MWField - MW_FIELD_OFFSET).value = True
    optNETorRT(.NETorRT).value = True
    Select Case .TolType
    Case gltPPM
      optTolType(0).value = True
    Case gltABS
      optTolType(1).value = True
    Case Else
      Debug.Assert False
    End Select
    'save old value and set search on "search all"
    OldSearchFlag = .SearchFlag
    .SearchFlag = 0         'search all
    'NETTol is used both for NET and RT
    If .NETTol >= 0 Then
       txtNETTol.Text = .NETTol
       txtNETTol_Validate False
    Else
       txtNETTol.Text = ""
    End If
End With

With GelSearchDef(CallerID).AMTSearchMassMods
    SetCheckBox chkPEO, .PEO
    SetCheckBox chkICATLt, .ICATd0
    SetCheckBox chkICATHv, .ICATd8
    SetCheckBox chkAlkylation, .Alkylation
    txtAlkylationMWCorrection = .AlkylationMass
    
    PopulateComboBoxes
    
    cboResidueToModify.ListIndex = 0
    If Len(.ResidueToModify) >= 1 Then
        For intIndex = 0 To cboResidueToModify.ListCount - 1
            If UCase(cboResidueToModify.List(intIndex)) = UCase(.ResidueToModify) Then
                cboResidueToModify.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    txtResidueToModifyMass = Round(.ResidueMassModification, 5)
    
    SetAlkylationMWCorrection .AlkylationMass
    SetDBSearchModType .DynamicMods
    SetDBSearchNType .N15InsteadOfN14
End With

If Not GelAnalysis(CallerID) Is Nothing Then
    ExpAnalysisSPName = GelAnalysis(CallerID).MTDB.DBStuff(PUT_NEW_ANALYSIS).value
    ExpPeakSPName = GelAnalysis(CallerID).MTDB.DBStuff(PUT_FTICR_PEAK).value
    ExpUmcSPName = GelAnalysis(CallerID).MTDB.DBStuff(PUT_FTICR_UMC).value
    ExpUmcMatchSPName = GelAnalysis(CallerID).MTDB.DBStuff(PUT_FTICR_UMC_MATCH).value
    ExpQuantitationDescription = GelAnalysis(CallerID).MTDB.DBStuff(PUT_QUANTITATION_DESCRIPTION).value
End If

If Len(ExpUmcSPName) = 0 Then
    ExpUmcSPName = "AddFTICRUmc"
End If
Debug.Assert ExpUmcSPName = "AddFTICRUmc"

If Len(ExpUmcMatchSPName) = 0 Then
    ExpUmcMatchSPName = "AddFTICRUmcMatch"
End If
Debug.Assert ExpUmcMatchSPName = "AddFTICRUmcMatch"

If Len(ExpQuantitationDescription) = 0 Then
    ExpQuantitationDescription = "AddQuantitationDescription"
End If
Debug.Assert ExpQuantitationDescription = "AddQuantitationDescription"

If Len(ExpAnalysisSPName) = 0 Then
    ExpAnalysisSPName = "AddMatchMaking"
End If
Debug.Assert ExpAnalysisSPName = "AddMatchMaking"

If Len(ExpPeakSPName) = 0 Then
    ExpPeakSPName = "AddFTICRPeak"
End If
Debug.Assert ExpPeakSPName = "AddFTICRPeak"

' Possibly add a checkmark to the mnuFReportIncludeORFs menu
mnuFReportIncludeORFs.Checked = glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput
Exit Sub

FormLoadErrorHandler:
LogErrors Err.Number, "frmSearchMT_UMC.Form_Load"
Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
' Restore .SearchFlag using the saved value
samtDef.SearchFlag = OldSearchFlag
If Not objMTDBNameLookupClass Is Nothing Then
    objMTDBNameLookupClass.DeleteData
    Set objMTDBNameLookupClass = Nothing
End If
End Sub

Private Sub mnuET_Click(Index As Integer)
Dim i As Long
On Error Resume Next
If GelAnalysis(CallerID) Is Nothing Then Index = etGenericNET
Select Case Index
Case etGenericNET
  txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
Case etTICFitNET
  With GelAnalysis(CallerID)
    If .NET_Slope <> 0 Then
       txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
    Else
       txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
    End If
  End With
  If Err Then
     MsgBox "Make sure display is loaded as analysis! Use New Analysis command from the File menu!", vbOKOnly, glFGTU
     Exit Sub
  End If
Case etGANET
  With GelAnalysis(CallerID)
    If .GANET_Slope <> 0 Then
       txtNETFormula.Text = ConstructNETFormula(.GANET_Slope, .GANET_Intercept)
    Else
       txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
    End If
  End With
  If Err Then
     MsgBox "Make sure display is loaded as analysis! Use New Analysis command from the File menu!", vbOKOnly, glFGTU
     Exit Sub
  End If
End Select
For i = mnuET.LBound To mnuET.UBound
    If i = Index Then
       mnuET(i).Checked = True
       lblETType.Caption = "ET: " & mnuET(i).Caption
    Else
       mnuET(i).Checked = False
    End If
Next i
Call txtNETFormula_LostFocus        'make sure expression evaluator is
                                    'initialized for this formula
End Sub

Private Sub mnuETDummy_Click()
Call PickParameters
End Sub

Private Sub mnuFExpMTDB_Click()
Dim eResponse As VbMsgBoxResult
Dim strStatus As String
Dim strUMCSearchMode As String

If mMatchStatsCount = 0 And Not glbPreferencesExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches Then
    MsgBox "Search results not found in memory.", vbInformation + vbOKOnly, "Nothing to Export"
Else
    eResponse = MsgBox("Proceed with exporting of the search results to the database?  This is an advanced feature that should normally only be performed during VIPER Automated PRISM Analysis Mode.  If you continue, you will be prompted for a password.", vbQuestion + vbYesNo + vbDefaultButton1, "Export Results")
    If eResponse = vbYes Then
        If QueryUserForExportToDBPassword(, False) Then
            ' Update the text in MD_Parameters
            strUMCSearchMode = FindSettingInAnalysisHistory(CallerID, UMC_SEARCH_MODE_SETTING_TEXT, , True, ":", ";")
            If Right(strUMCSearchMode, 1) = ")" Then strUMCSearchMode = Left(strUMCSearchMode, Len(strUMCSearchMode) - 1)
            GelAnalysis(CallerID).MD_Parameters = ConstructAnalysisParametersText(CallerID, strUMCSearchMode, AUTO_SEARCH_UMC_HERETIC)
                
            strStatus = ExportMTDBbyUMC(True)
            MsgBox strStatus, vbInformation + vbOKOnly, glFGTU
        Else
            MsgBox "Invalid password, export aborted.", vbExclamation Or vbOKOnly, "Invalid"
        End If
    End If
End If
End Sub

Private Sub mnuF_Click()
Call PickParameters
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFReport_Click()
ShowOrSaveResultsOld ""
End Sub

Private Sub mnuFReportIncludeORFs_Click()
    mnuFReportIncludeORFs.Checked = Not mnuFReportIncludeORFs.Checked
    glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput = mnuFReportIncludeORFs.Checked
End Sub

Private Sub mnuFReportWithORFs_Click()
ShowOrSaveResultsByUMC ""
End Sub

Private Sub mnuFSearchAll_Click()
StartSearchAll
End Sub

Private Sub mnuFSearchNonPaired_Click()
StartSearchNonPaired
End Sub

Private Sub mnuFSearchPaired_Click()
StartSearchPaired
End Sub

Private Sub mnuMT_Click()
Call PickParameters
End Sub

Private Sub mnuMTLoadLegacy_Click()
    LoadLegacyMassTags
End Sub

Private Sub mnuMTLoadMT_Click()
'------------------------------------------------------------
'load/reload mass tags
'------------------------------------------------------------
If Not GelAnalysis(CallerID) Is Nothing Then
   Call LoadMTDB(True)
Else
   Call Info_NoMTDBLink
   lblMTStatus.Caption = "No mass tags loaded"
End If
End Sub

Private Sub mnuMTStatus_Click()
'----------------------------------------------
'displays short mass tags statistics, it might
'help with determining problems with mass tags
'----------------------------------------------
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly
Me.MousePointer = vbDefault
End Sub

Private Sub optMWField_Click(Index As Integer)
samtDef.MWField = 6 + Index
End Sub

Private Sub optN_Click(Index As Integer)
N14N15 = Index
End Sub

Private Sub optNETorRT_Click(Index As Integer)
samtDef.NETorRT = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   samtDef.TolType = gltPPM
Else
   samtDef.TolType = gltABS
End If
End Sub

Private Sub txtAlkylationMWCorrection_LostFocus()
If IsNumeric(txtAlkylationMWCorrection.Text) Then
   AlkMWCorrection = CDbl(txtAlkylationMWCorrection.Text)
Else
   MsgBox "This argument should be numeric!", vbOKOnly, glFGTU
   txtAlkylationMWCorrection.SetFocus
End If
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   samtDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "Molecular Mass Tolerance should be numeric value!", vbOKOnly
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNETFormula_LostFocus()
'------------------------------------------------
'initialize new expression evaluator
'------------------------------------------------
If Not InitExprEvaluator(txtNETFormula.Text) Then
   MsgBox "Error in elution calculation formula!", vbOKOnly, glFGTU
   txtNETFormula.SetFocus
Else
   samtDef.Formula = txtNETFormula.Text
End If
End Sub

Private Sub txtNETTol_LostFocus()
If IsNumeric(txtNETTol.Text) Then
   samtDef.NETTol = CDbl(txtNETTol.Text)
Else
   If Len(Trim(txtNETTol.Text)) > 0 Then
      MsgBox "NET Tolerance should be number between 0 and 1!", vbOKOnly
      txtNETTol.SetFocus
   Else
      samtDef.NETTol = -1   'do not consider NET when searching
   End If
End If
End Sub

Private Sub txtNETTol_Validate(Cancel As Boolean)
    TextBoxLimitNumberLength txtNETTol, 12
End Sub

Private Function IsValidMatch(CurrMW As Double, AbsMWErr As Double, CurrScan As Long, lngMassTagIndexOriginal As Long, dblAMTMass As Double) As Boolean
    ' Checks if CurrMW is within tolerance of the given mass tag
    ' Also checks if the NET equivalent of CurrScan is within tolerance of the NET value for the given mass tag
    ' Returns True if both are within tolerance, false otherwise
    
    Dim InvalidMatch As Boolean
    
    ' If CurrMW is not within AbsMWErr of dblAMTMass then this match is inherited
    If Abs(CurrMW - dblAMTMass) > AbsMWErr Then
        InvalidMatch = True
    Else
        ' If CurrScan is not within .NETTol of mMTNET() then this match is inherited
        If samtDef.NETTol >= 0 Then
            If Abs(ConvertScanToNET(CurrScan) - AMTNET(lngMassTagIndexOriginal)) > samtDef.NETTol Then
                InvalidMatch = True
            End If
        End If
    End If
    
    IsValidMatch = Not InvalidMatch
End Function

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean

    If ConfirmMassTagsAndNetAdjLockersLoaded(Me, CallerID, True, 0, blnForceReload, True, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTStatus.Caption = "Mass tags count: " & AMTCnt
    
        If Not CreateNewMTSearchObject() Then
           lblMTStatus.Caption = "Error creating search object!"
        End If
    Else
        If blnDBConnectionError Then
            lblMTStatus.Caption = "Error loading mass tags: database connection error!"
        Else
            lblMTStatus.Caption = "Error loading mass tags: no valid mass tags were found (possibly missing NET values)"
        End If
    End If

End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub


Private Sub ScoreCurrIDs()
'---------------------------------------------------------------
'does unique count of identifications found in array mgCurrIDInd
'after this procedure mgCurrIDInd array will contain only unique
'identifications
'After procedure unique IDs are ordered ascending on scores
'NOTE: score is just unique count of each identification
'NOTE: results are returned in the same container arrays
'NOTE: this procedure is called only for mgCurrIDCnt>0
'---------------------------------------------------------------
Dim UnqCnt As Long
Dim TmpID() As Long
Dim CurrID As Long
Dim i As Long, j As Long
Dim qsd As New QSDouble
On Error Resume Next

TmpID = mgCurrIDInd     'copy mgCurrIDInd data to temporary array
UnqCnt = 0
'zero scores and unique identifications
ReDim mgCurrIDInd(mgCurrIDCnt - 1)
ReDim mgCurrIDScore(mgCurrIDCnt - 1)

For i = 0 To mgCurrIDCnt - 1
    CurrID = TmpID(i)
    For j = 0 To UnqCnt - 1
        If CurrID = mgCurrIDInd(j) Then
           mgCurrIDScore(j) = mgCurrIDScore(j) + 1
           Exit For
        End If
    Next j
    If j > UnqCnt - 1 Then          'CurrID not found among unique
       UnqCnt = UnqCnt + 1          'IDs - add it
       mgCurrIDInd(UnqCnt - 1) = CurrID
       mgCurrIDScore(UnqCnt - 1) = 1
    End If
Next i
'truncate the unique counts
If UnqCnt > 0 Then
   mgCurrIDCnt = UnqCnt
   Call ManageCurrID(MNG_TRIM)
   ReDim Preserve mgCurrIDInd(UnqCnt - 1)
   ReDim Preserve mgCurrIDScore(UnqCnt - 1)
   If mgCurrIDCnt > 1 Then
      Call qsd.QSAsc(mgCurrIDScore(), mgCurrIDInd())
   End If
Else                            'should not happen but...
   mgCurrIDCnt = -1
   Call ManageCurrID(MNG_ERASE)
End If
End Sub

Private Sub WriteORFResults(ts As TextStream, strLineOutPrefix As String, lngMassTagID As Long, objORFNameFastSearch As FastSearchArrayLong, Optional strSepChar As String = glARG_SEP)
    
    Dim ORFNames() As String            ' 0-based array
    Dim lngORFNamesCount As Long
    Dim lngORFNameIndex As Long

    If MTtoORFMapCount = 0 Then
        lngORFNamesCount = LookupORFNamesForMTIDusingMTDBNamer(objMTDBNameLookupClass, lngMassTagID, ORFNames())
    Else
        lngORFNamesCount = LookupORFNamesForMTIDusingMTtoORFMapOptimized(lngMassTagID, ORFNames(), objORFNameFastSearch)
    End If
    
    If lngORFNamesCount > 0 Then
        For lngORFNameIndex = 0 To lngORFNamesCount - 1
            ts.WriteLine strLineOutPrefix & strSepChar & lngORFNamesCount & strSepChar & ORFNames(lngORFNameIndex)
        Next lngORFNameIndex
    Else
        ts.WriteLine strLineOutPrefix & strSepChar & "0" & strSepChar & "UnknownORF"
    End If

End Sub

Private Sub WriteReportHeader(ts As TextStream, Optional strSepChar As String = glARG_SEP)
'--------------------------------------------------------------------
'write report header block
'--------------------------------------------------------------------
ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
ts.WriteLine "Gel File: " & GelBody(CallerID).Caption
ts.WriteLine "Reporting identification for UMC search"
ts.WriteLine "Total Hits: " & Trim(mMatchStatsCount)
ts.WriteLine
ts.WriteLine "UMC_Ind" & strSepChar & "UMC_MW" & strSepChar & "UMC_Abu" & strSepChar _
        & "UMC_Cnt" & strSepChar & "FN1" & strSepChar & "FN2" & strSepChar & "ID" _
        & strSepChar & "ID MW" & strSepChar & "ID Mods" & strSepChar & "ID Score"
End Sub

Private Sub RecordSearchResultsInData()
    ' Step through mUMCMatchStats() and add the ID's for each UMC to all of the members of each UMC
    
    Dim lngIndex As Long, lngMemberIndex As Long
    Dim lngUMCIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long                  'absolute index in mMT... arrays
    Dim lngMassTagIndexOriginal As Long                 'absolute index in AMT... arrays
    Dim lngIonIndexOriginal As Long
    Dim blnAddAMTRef As Boolean
    Dim lngIonCountUpdated As Long
    Dim blnStoreAbsoluteValueOfError As Boolean
    
    Dim AMTRef As String
    Dim dblAMTMass As Double
    Dim CurrMW As Double, AbsMWErr As Double
    Dim CurrScan As Long
    
    ' Need to remove any existing search results before adding these new ones
    RemoveAMT CallerID, glScope.glSc_All
    
    GelStatus(CallerID).Dirty = True
    AddToAnalysisHistory CallerID, "Deleted mass tag search results from ions"
    
    'always reinitialize statistics arrays
    InitAMTStat
    
    KeyPressAbortProcess = 0
    
    ' Note, we are no longer storing the absolute value of errors in the AMT Ref for the data
    blnStoreAbsoluteValueOfError = False
    
    CheckNETEquationStatus
    
On Error GoTo RecordSearchResultsInDataErrorHandler

    With GelData(CallerID)
        For lngIndex = 0 To mMatchStatsCount - 1
            If lngIndex Mod 25 = 0 Then
                UpdateStatus "Storing results: " & Trim(lngIndex) & " / " & Trim(mMatchStatsCount)
                If KeyPressAbortProcess > 1 Then Exit For
            End If
            
            lngUMCIndexOriginal = mUMCMatchStats(lngIndex).UMCIndex
            lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngIndex).IDIndex)
            lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            
            If LastSearchTypeN14N15 = SEARCH_N14 Then
                ' N14
                dblAMTMass = mMTMWN14(mUMCMatchStats(lngIndex).IDIndex)
            Else
                ' N15
                dblAMTMass = mMTMWN15(mUMCMatchStats(lngIndex).IDIndex)
            End If
            
            For lngMemberIndex = 0 To GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassCount - 1
                lngIonIndexOriginal = GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMInd(lngMemberIndex)
                blnAddAMTRef = False
                
                Select Case GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMType(lngMemberIndex)
                Case glCSType
                    CurrMW = .CSNum(lngIonIndexOriginal, csfMW)
                    CurrScan = .CSNum(lngIonIndexOriginal, csfScan)
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = CurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select
                    
                    AMTRef = ConstructAMTReference(.CSNum(lngIonIndexOriginal, csfMW), ConvertScanToNET(CLng(.CSNum(lngIonIndexOriginal, csfScan))), 0, lngMassTagIndexOriginal, dblAMTMass, blnStoreAbsoluteValueOfError)
                    If IsNull(InStr(.CSVar(lngIonIndexOriginal, csvfMTID), AMTRef)) Then
                        blnAddAMTRef = True
                    ElseIf InStr(.CSVar(lngIonIndexOriginal, csvfMTID), AMTRef) <= 0 Then
                        blnAddAMTRef = True
                    End If
                    
                    If blnAddAMTRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        If Not IsValidMatch(CurrMW, AbsMWErr, CurrScan, lngMassTagIndexOriginal, dblAMTMass) Then
                            AMTRef = Trim(AMTRef)
                            If Right(AMTRef, 1) = glARG_SEP Then
                                AMTRef = Left(AMTRef, Len(AMTRef) - 1)
                            End If
                            AMTRef = AMTRef & AMTMatchInheritedMark
                        End If
                        
                        InsertBefore .CSVar(lngIonIndexOriginal, csvfMTID), AMTRef
                    End If
                Case glIsoType
                    CurrMW = .IsoNum(lngIonIndexOriginal, samtDef.MWField)
                    CurrScan = .IsoNum(lngIonIndexOriginal, isfScan)
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = CurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select

                    AMTRef = ConstructAMTReference(.IsoNum(lngIonIndexOriginal, samtDef.MWField), ConvertScanToNET(CLng(.IsoNum(lngIonIndexOriginal, isfScan))), 0, lngMassTagIndexOriginal, dblAMTMass, blnStoreAbsoluteValueOfError)
                    If IsNull(.IsoVar(lngIonIndexOriginal, isvfMTID)) Then
                        blnAddAMTRef = True
                    ElseIf InStr(.IsoVar(lngIonIndexOriginal, isvfMTID), AMTRef) <= 0 Then
                        blnAddAMTRef = True
                    End If
                    
                    If blnAddAMTRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        If Not IsValidMatch(CurrMW, AbsMWErr, CurrScan, lngMassTagIndexOriginal, dblAMTMass) Then
                            AMTRef = Trim(AMTRef)
                            If Right(AMTRef, 1) = glARG_SEP Then
                                AMTRef = Left(AMTRef, Len(AMTRef) - 1)
                            End If
                            AMTRef = AMTRef & AMTMatchInheritedMark
                        End If
                        
                        InsertBefore .IsoVar(lngIonIndexOriginal, isvfMTID), AMTRef
                    End If
                End Select
            Next lngMemberIndex
        Next lngIndex
    End With
    
    If KeyPressAbortProcess <= 1 Then
        AddToAnalysisHistory CallerID, "Stored search results in ions; recorded all mass tag hits for each UMC in all members of the UMC; total ions updated = " & Trim(lngIonCountUpdated)
    End If
        
    Exit Sub

RecordSearchResultsInDataErrorHandler:
    LogErrors Err.Number, "frmSearchMT_UMC->RecordSearchResultsInData"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured while storing the search results in the data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub ReportUMCID(ts As TextStream, ByVal mgInd As Long, Optional strSepChar As String = glARG_SEP)
'---------------------------------------------------------------------
'writes lines of report(all identifications) for index in mg... arrays
'---------------------------------------------------------------------
Dim SP As String                    'pair part of line
Dim UMCInd As Long                  'absolute index of UMC
Dim tMTInd As Long                  'absolute index in mMT... arrays
Dim tAMTInd As Long                 'absolute index in AMT... arrays
On Error Resume Next
UMCInd = mUMCMatchStats(mgInd).UMCIndex
With GelUMC(CallerID).UMCs(UMCInd)
   SP = mUMCMatchStats(mgInd).UMCIndex & strSepChar & .ClassMW & strSepChar & .ClassAbundance _
      & strSepChar & .ClassCount & strSepChar & .MinScan & strSepChar _
      & .MaxScan & strSepChar
End With
tMTInd = mMTInd(mUMCMatchStats(mgInd).IDIndex)
tAMTInd = mMTOrInd(tMTInd)
SP = SP & AMTID(tAMTInd) & strSepChar & AMTMW(tAMTInd) & strSepChar _
     & NTypeStr & Chr$(32) & mMTMods(tMTInd) & strSepChar & mUMCMatchStats(mgInd).MemberHitCount
ts.WriteLine SP
End Sub

Private Function InitExprEvaluator(ByVal sExpr As String) As Boolean
'-------------------------------------------------------------------
'initializes expression evaluator for elution time
'-------------------------------------------------------------------
On Error Resume Next
Set MyExprEva = New ExprEvaluator
With MyExprEva
    .Vars.add 1, "FN"
    .Vars.add 2, "MinFN"
    .Vars.add 3, "MaxFN"
    .Expr = sExpr
    InitExprEvaluator = .IsExprValid
    ReDim VarVals(1 To 3)
End With
End Function


Private Function Elution(FN, MinFN, MaxFN)
'---------------------------------------------------
'this function does not care are we using NET or RT
'---------------------------------------------------
VarVals(1) = FN
VarVals(2) = MinFN
VarVals(3) = MaxFN
Elution = MyExprEva.ExprVal(VarVals())
End Function

Public Function ExportMTDBbyUMC(Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional strIniFileName As String = "", Optional ByRef lngErrorNumber As Long, Optional ByRef lngMDID As Long = -1) As String
'--------------------------------------------------------------------------------
' This function exports data to both T_FTICR_Peak_Results and T_FTICR_UMC_Results (plus T_FTICR_UMC_ResultDetails)
' Optionally returns the error number in lngErrorNumber
' Optionally returns the MD_ID value in lngMDID
'--------------------------------------------------------------------------------
    
    Dim strStatus As String
    Dim eResponse As VbMsgBoxResult
    Dim blnAddQuantitationEntry As Boolean
    Dim blnExportUMCsWithNoMatches As Boolean
    
    lngMDID = -1
    cmdSearchAllUMCs.Visible = False
        
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        With glbPreferencesExpanded.AutoAnalysisOptions
            blnAddQuantitationEntry = .AddQuantitationDescriptionEntry
            blnExportUMCsWithNoMatches = .ExportUMCsWithNoMatches
        End With
    Else
        eResponse = MsgBox("Export UMC's that do not have any database matches?", vbQuestion + vbYesNo + vbDefaultButton2, "Export Non-Matching UMC's")
        blnExportUMCsWithNoMatches = (eResponse = vbYes)
    End If
    
    ' Note: The following function call will create a new entry in T_Match_Making_Description
    strStatus = ExportMTDBbyUMCToPeakResultsTable(lngMDID, blnUpdateGANETForAnalysisInDB, lngErrorNumber)
    
    If lngErrorNumber = 0 And lngMDID >= 0 Then
        strStatus = strStatus & vbCrLf & ExportMTDBbyUMCToUMCResultsTable(lngMDID, False, False, lngErrorNumber, blnAddQuantitationEntry, blnExportUMCsWithNoMatches, strIniFileName)
    End If
    
    cmdSearchAllUMCs.Visible = True
    ExportMTDBbyUMC = strStatus
    
End Function

Private Function ExportMTDBbyUMCToPeakResultsTable(ByRef lngMDID As Long, Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByRef lngErrorNumber As Long) As String
'---------------------------------------------------
'this is simple but long procedure of exporting data
'results to Organism Mass Tags database associated with gel
'
'We're currently writing the results to the T_Match_Making_Description table and T_FTICR_Peak_Results
'These tables are designed to hold search results from an ion-by-ion search (either using all ions or using UMC ions only)
'Since this form uses a UMC by UMC search, and we assign all matches for a UMC to all ions for the UMC, we'll
'  only export the search results for the class representative ion for each UMC (typically the most abundant ion)
'
'Returns a status message
'lngErrorNumber will contain the error number, if an error occurs
'lngMDID contains the new MMD_ID value
'---------------------------------------------------
Dim mgInd As Long
Dim lngUMCIndexOriginal As Long, lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long
Dim ExpCnt As Long
Dim strCaptionSaved As String
Dim strExportStatus As String

Dim lngPairIndex As Long
Dim dblExpressionRatio As Double

Dim objP1IndFastSearch As FastSearchArrayLong
Dim objP2IndFastSearch As FastSearchArrayLong
Dim blnPairsPresent As Boolean

'ADO objects for stored procedure adding Match Making row
Dim cnNew As New ADODB.Connection

'ADO objects for stored procedure that adds FTICR peak rows
Dim cmdPutNewPeak As New ADODB.Command
Dim prmMMDID As New ADODB.Parameter
Dim prmFTICRID As New ADODB.Parameter
Dim prmFTICRType As New ADODB.Parameter
Dim prmScanNumber As New ADODB.Parameter
Dim prmChargeState As New ADODB.Parameter
Dim prmMonoisotopicMass As New ADODB.Parameter
Dim prmAbundance As New ADODB.Parameter
Dim prmFit As New ADODB.Parameter
Dim prmExpressionRatio As New ADODB.Parameter
Dim prmLckID As New ADODB.Parameter
Dim prmFreqShift As New ADODB.Parameter
Dim prmMassCorrection As New ADODB.Parameter
Dim prmMassTagID As New ADODB.Parameter
Dim prmResType As New ADODB.Parameter
Dim prmHitsCount As New ADODB.Parameter
Dim prmUMCInd As New ADODB.Parameter
Dim prmUMCFirstScan As New ADODB.Parameter
Dim prmUMCLastScan As New ADODB.Parameter
Dim prmUMCCount As New ADODB.Parameter
Dim prmUMCAbundance As New ADODB.Parameter
Dim prmUMCBestFit As New ADODB.Parameter
Dim prmUMCAvgMW As New ADODB.Parameter
Dim prmPairInd As New ADODB.Parameter

On Error GoTo err_ExportMTDBbyUMC

strCaptionSaved = Me.Caption

' Connect to the database
Me.Caption = "Connecting to the database"
If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
    Debug.Assert False
    lngErrorNumber = -1
    Me.Caption = strCaptionSaved
    ExportMTDBbyUMCToPeakResultsTable = "Error: Unable to establish a connection to the database"
    Exit Function
End If

'first write new analysis in T_Match_Making_Description table
AddEntryToMatchMakingDescriptionTable cnNew, lngMDID, ExpAnalysisSPName, CallerID, mMatchStatsCount, True
AddToAnalysisHistory CallerID, "Exported UMC Identification results (UMC based ion-by-ion search) to Peak Results table in database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file

'nothing to export
If mMatchStatsCount <= 0 Then
    cnNew.Close
    Me.Caption = strCaptionSaved
    Exit Function
End If

' Initialize the SP
InitializeSPCommand cmdPutNewPeak, cnNew, ExpPeakSPName

Set prmMMDID = cmdPutNewPeak.CreateParameter("MMDID", adInteger, adParamInput, , lngMDID)
cmdPutNewPeak.Parameters.Append prmMMDID
Set prmFTICRID = cmdPutNewPeak.CreateParameter("FTICRID", adVarChar, adParamInput, 50, Null)
cmdPutNewPeak.Parameters.Append prmFTICRID
Set prmFTICRType = cmdPutNewPeak.CreateParameter("FTICRType", adTinyInt, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmFTICRType
Set prmScanNumber = cmdPutNewPeak.CreateParameter("ScanNumber", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmScanNumber
Set prmChargeState = cmdPutNewPeak.CreateParameter("ChargeState", adSmallInt, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmChargeState
Set prmMonoisotopicMass = cmdPutNewPeak.CreateParameter("MonoisotopicMass", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmMonoisotopicMass
Set prmAbundance = cmdPutNewPeak.CreateParameter("Abundance", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmAbundance
Set prmFit = cmdPutNewPeak.CreateParameter("Fit", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmFit
Set prmExpressionRatio = cmdPutNewPeak.CreateParameter("ExpressionRatio", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmExpressionRatio
Set prmLckID = cmdPutNewPeak.CreateParameter("LckID", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmLckID
Set prmFreqShift = cmdPutNewPeak.CreateParameter("FreqShift", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmFreqShift
Set prmMassCorrection = cmdPutNewPeak.CreateParameter("MassCorrection", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmMassCorrection
Set prmMassTagID = cmdPutNewPeak.CreateParameter("MassTagID", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmMassTagID
Set prmResType = cmdPutNewPeak.CreateParameter("Type", adInteger, adParamInput, , FPR_Type_Standard)
cmdPutNewPeak.Parameters.Append prmResType
Set prmHitsCount = cmdPutNewPeak.CreateParameter("HitCount", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmHitsCount
Set prmUMCInd = cmdPutNewPeak.CreateParameter("UMCInd", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmUMCInd
Set prmUMCFirstScan = cmdPutNewPeak.CreateParameter("UMCFirstScan", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmUMCFirstScan
Set prmUMCLastScan = cmdPutNewPeak.CreateParameter("UMCLastScan", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmUMCLastScan
Set prmUMCCount = cmdPutNewPeak.CreateParameter("UMCCount", adInteger, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmUMCCount
Set prmUMCAbundance = cmdPutNewPeak.CreateParameter("UMCAbundance", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmUMCAbundance
Set prmUMCBestFit = cmdPutNewPeak.CreateParameter("UMCBestFit", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmUMCBestFit
Set prmUMCAvgMW = cmdPutNewPeak.CreateParameter("UMCAvgMW", adDouble, adParamInput, , 0)
cmdPutNewPeak.Parameters.Append prmUMCAvgMW
Set prmPairInd = cmdPutNewPeak.CreateParameter("PairInd", adInteger, adParamInput, , -1)
cmdPutNewPeak.Parameters.Append prmPairInd

' Initialize the PairIndex lookup objects
blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)

Me.Caption = "Exporting peaks to DB: 0 / " & Trim(mMatchStatsCount)

'now export data
ExpCnt = 0
With GelData(CallerID)
    ' Step through the UMC hits and export information on each hit
    ' Since the target table is an ion-based table, will use the index and info of the class representative ion
    For mgInd = 0 To mMatchStatsCount - 1
        If mgInd Mod 25 = 0 Then
            Me.Caption = "Exporting peaks to DB: " & Trim(mgInd) & " / " & Trim(mMatchStatsCount)
            DoEvents
        End If
        
        lngUMCIndexOriginal = mUMCMatchStats(mgInd).UMCIndex
        
        With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
            prmFTICRID.value = .ClassRepInd
            prmFTICRType.value = .ClassRepType
            
            Select Case .ClassRepType
            Case glCSType
                prmScanNumber.value = GelData(CallerID).CSNum(.ClassRepInd, csfScan)
                prmChargeState.value = GelData(CallerID).CSNum(.ClassRepInd, csfFirstCS)
                prmMonoisotopicMass.value = GelData(CallerID).CSNum(.ClassRepInd, csfMW)
                prmAbundance.value = GelData(CallerID).CSNum(.ClassRepInd, csfAbu)
                prmFit.value = GelData(CallerID).CSNum(.ClassRepInd, csfStD)     'standard deviation
                If GelLM(CallerID).CSCnt > 0 Then
                  prmLckID.value = GelLM(CallerID).CSLckID(.ClassRepInd)
                  prmFreqShift.value = GelLM(CallerID).CSFreqShift(.ClassRepInd)
                  prmMassCorrection.value = GelLM(CallerID).CSMassCorrection(.ClassRepInd)
                End If
            Case glIsoType
                prmScanNumber.value = GelData(CallerID).IsoNum(.ClassRepInd, isfScan)
                prmChargeState.value = GelData(CallerID).IsoNum(.ClassRepInd, isfCS)
                prmMonoisotopicMass.value = GelData(CallerID).IsoNum(.ClassRepInd, isfMWMono)
                prmAbundance.value = GelData(CallerID).IsoNum(.ClassRepInd, isfAbu)
                prmFit.value = GelData(CallerID).IsoNum(.ClassRepInd, isfFit)
                If GelLM(CallerID).IsoCnt > 0 Then
                  prmLckID.value = GelLM(CallerID).IsoLckID(.ClassRepInd)
                  prmFreqShift.value = GelLM(CallerID).IsoFreqShift(.ClassRepInd)
                  prmMassCorrection.value = GelLM(CallerID).IsoMassCorrection(.ClassRepInd)
                End If
            End Select
            
            ' Note: The multi-hit count value for the UMC is the same as that for the class representative, and can thus be placed in prmHitsCount
            prmHitsCount.value = mUMCMatchStats(mgInd).MultiAMTHitCount
            prmUMCInd.value = mUMCMatchStats(mgInd).UMCIndex
            prmUMCFirstScan.value = .MinScan
            prmUMCLastScan.value = .MaxScan
            prmUMCCount.value = .ClassCount
            prmUMCAbundance.value = .ClassAbundance
            prmUMCBestFit.value = Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), FIT_PRECISION)
            prmUMCAvgMW.value = Round(.ClassMW, MASS_PRECISION)              ' This is usually the median mass of the class, not the average mass
        
            lngPairIndex = -1
            dblExpressionRatio = 0
            If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
                lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, objP1IndFastSearch, objP2IndFastSearch, (LastSearchTypeN14N15 = SEARCH_N15), dblExpressionRatio, 0, 0, 0)
            End If
            
            If lngPairIndex >= 0 Then
                prmExpressionRatio.value = dblExpressionRatio
                prmPairInd.value = lngPairIndex
            Else
                prmExpressionRatio.value = LookupExpressionRatioValue(CallerID, .ClassRepInd, (.ClassRepType = glIsoType))
                prmPairInd.value = -1
            End If
            
        End With
        
        lngMassTagIndexPointer = mMTInd(mUMCMatchStats(mgInd).IDIndex)
        lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
        
        prmMassTagID.value = AMTID(lngMassTagIndexOriginal)

        cmdPutNewPeak.Execute
        ExpCnt = ExpCnt + 1
    Next mgInd
End With

' MonroeMod
AddToAnalysisHistory CallerID, "Export to Peak Results table details: UMC Peaks Match Count = " & ExpCnt

Me.Caption = strCaptionSaved

strExportStatus = ExpCnt & " associations between mass tags and UMC's exported to peak results table."
Set cmdPutNewPeak.ActiveConnection = Nothing
cnNew.Close

If blnUpdateGANETForAnalysisInDB Then
    ' Export the the GANET Slope, Intercept, and Fit to the database
    With GelAnalysis(CallerID)
        strExportStatus = strExportStatus & vbCrLf & ExportGANETtoMTDB(CallerID, .GANET_Slope, .GANET_Intercept, .GANET_Fit)
    End With
End If

Set objP1IndFastSearch = Nothing
Set objP2IndFastSearch = Nothing

ExportMTDBbyUMCToPeakResultsTable = strExportStatus
lngErrorNumber = 0
Exit Function

err_ExportMTDBbyUMC:
ExportMTDBbyUMCToPeakResultsTable = "Error: " & Err.Number & vbCrLf & Err.Description
lngErrorNumber = Err.Number
On Error Resume Next
If Not cnNew Is Nothing Then cnNew.Close
Me.Caption = strCaptionSaved
Set objP1IndFastSearch = Nothing
Set objP2IndFastSearch = Nothing

End Function

Private Function ExportMTDBbyUMCToUMCResultsTable(ByRef lngMDID As Long, Optional blnCreateNewEntryInMMDTable As Boolean = False, Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByRef lngErrorNumber As Long, Optional ByVal blnAddQuantitationDescriptionEntry As Boolean = True, Optional ByVal blnExportUMCsWithNoMatches As Boolean = True, Optional ByVal strIniFileName As String = "") As String
'---------------------------------------------------
'This function will export data to the T_FTICR_UMC_Results table and the T_FTICR_UMC_ResultDetails tables
'
'It will create a new entry in the T_Match_Making_Description if blnCreateNewEntryInMMDTable = True
'If blnAddQuantitationDescriptionEntry = True, then calls ExportMTDBAddQuantitationDescriptionEntry
'  to create a new entry in T_Quantitation_Description and T_Quantitation_MDIDs
'
'Returns a status message
'lngErrorNumber will contain the error number, if an error occurs
'---------------------------------------------------
Dim lngPointer As Long, lngUMCIndex As Long
Dim lngUMCIndexOriginal As Long
Dim lngUMCIndexOriginalLastStored As Long
Dim lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long

Dim lngPairIndex As Long
Dim dblExpressionRatio As Double
Dim dblExpressionRatioStDev As Double
Dim intExpressionRatioChargeStateBasisCount As Integer
Dim lngExpressionRatioMemberBasisCount As Long

Dim objP1IndFastSearch As FastSearchArrayLong
Dim objP2IndFastSearch As FastSearchArrayLong
Dim blnPairsPresent As Boolean

Dim ExpCnt As Long
Dim strCaptionSaved As String
Dim strExportStatus As String
Dim strMassMods As String

'ADO objects for stored procedure adding Match Making row
Dim cnNew As New ADODB.Connection

'ADO objects for stored procedure that adds FTICR UMC rows
Dim cmdPutNewUMC As New ADODB.Command
Dim udtPutUMCParams As udtPutUMCParamsListType
    
'ADO objects for stored procedure adding UMC UMC Details
Dim cmdPutNewUMCMatch As New ADODB.Command
Dim udtPutUMCMatchParams As udtPutUMCMatchParamsListType

Dim blnUMCMatchFound() As Boolean       ' 0-based array

On Error GoTo err_ExportMTDBbyUMC

strCaptionSaved = Me.Caption

' Connect to the database
Me.Caption = "Connecting to the database"
If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
    Debug.Assert False
    lngErrorNumber = -1
    Me.Caption = strCaptionSaved
    ExportMTDBbyUMCToUMCResultsTable = "Error: Unable to establish a connection to the database"
    Exit Function
End If

If blnCreateNewEntryInMMDTable Then
    'first write new analysis in T_Match_Making_Description table
    AddEntryToMatchMakingDescriptionTable cnNew, lngMDID, ExpAnalysisSPName, CallerID, mMatchStatsCount, True
End If

If blnCreateNewEntryInMMDTable Or mMatchStatsCount > 0 Or blnExportUMCsWithNoMatches Then
    ' MonroeMod
    AddToAnalysisHistory CallerID, "Exported UMC Identification results (UMC based ion-by-ion search) to UMC Results table in database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
    If blnCreateNewEntryInMMDTable Then
        AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file
    End If
End If

'nothing to export
If mMatchStatsCount <= 0 And Not blnExportUMCsWithNoMatches Then
    cnNew.Close
    Me.Caption = strCaptionSaved
    Exit Function
End If

' Initialize cmdPutNewUMC and all of the params in udtPutUMCParams
ExportMTDBInitializePutNewUMCParams cnNew, cmdPutNewUMC, udtPutUMCParams, lngMDID, ExpUmcSPName

' Initialize the variables for accessing the AddFTICRUmcMatch SP
ExportMTDBInitializePutUMCMatchParams cnNew, cmdPutNewUMCMatch, udtPutUMCMatchParams, ExpUmcMatchSPName

' Initialize the PairIndex lookup objects
blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)

Select Case LastSearchTypeN14N15
Case SEARCH_N14
     NTypeStr = MOD_TKN_N14
Case SEARCH_N15
     NTypeStr = MOD_TKN_N15
End Select

Me.Caption = "Exporting UMC's to DB: 0 / " & Trim(mMatchStatsCount)

'now export data
ExpCnt = 0

    ' Step through the UMC hits and export information on each hit
    ' mUMCMatchStats() will contain multiple entries for each UMC if the UMC matched multiple mass tags
    ' However, we only want to write one entry for each UMC to T_FTICR_UMC_Results
    ' Thus, we need to keep track of whether or not an entry has been made to T_FTICR_UMC_Results
    ' Luckily, results are stored to mUMCMatchStats() in order of UMC Index
    
    lngUMCIndexOriginalLastStored = -1
    
    For lngPointer = 0 To mMatchStatsCount - 1
        If lngPointer Mod 25 = 0 Then
            Me.Caption = "Exporting UMC's to DB: " & Trim(lngPointer) & " / " & Trim(mMatchStatsCount)
            DoEvents
        End If
        
        lngUMCIndexOriginal = mUMCMatchStats(lngPointer).UMCIndex
        If lngUMCIndexOriginal <> lngUMCIndexOriginalLastStored Then
            ' Add a new row to T_FTICR_UMC_Results
            ' Note: we're recording the Peak FPRType as FPR_Type_Standard, even if we searched only paired UMC's
            ' However, if the UMC is a member of a pair, we'll record the pair index in the database
            lngPairIndex = -1
            dblExpressionRatio = 0
            dblExpressionRatioStDev = 0
            intExpressionRatioChargeStateBasisCount = 0
            lngExpressionRatioMemberBasisCount = 0
            If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
                lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, objP1IndFastSearch, objP2IndFastSearch, (LastSearchTypeN14N15 = SEARCH_N15), dblExpressionRatio, dblExpressionRatioStDev, intExpressionRatioChargeStateBasisCount, lngExpressionRatioMemberBasisCount)
            End If
            
            ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, CallerID, lngUMCIndexOriginal, mUMCMatchStats(lngPointer).MultiAMTHitCount, ClsStat(), FPR_Type_Standard, dblExpressionRatio, dblExpressionRatioStDev, intExpressionRatioChargeStateBasisCount, lngExpressionRatioMemberBasisCount, lngPairIndex
            
            udtPutUMCMatchParams.UMCResultsID.value = FixNullLng(udtPutUMCParams.UMCResultsIDReturn.value)
            
            ' Update lngUMCIndexOriginalLastStored
            lngUMCIndexOriginalLastStored = lngUMCIndexOriginal
        End If
        
        ' Now write an entry to T_FTICR_UMC_ResultDetails
        
        lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngPointer).IDIndex)
        lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
    
        udtPutUMCMatchParams.MassTagID.value = AMTID(lngMassTagIndexOriginal)
        udtPutUMCMatchParams.MatchScore.value = mUMCMatchStats(lngPointer).MemberHitCount
        
        strMassMods = NTypeStr
        If Len(mMTMods(lngMassTagIndexPointer)) > 0 Then
            strMassMods = strMassMods & " " & Trim(mMTMods(lngMassTagIndexPointer))
            If NTypeStr = MOD_TKN_N14 Then
                udtPutUMCMatchParams.MassTagModMass.value = mMTMWN14(mUMCMatchStats(lngPointer).IDIndex) - AMTMW(lngMassTagIndexOriginal)
            Else
                udtPutUMCMatchParams.MassTagModMass.value = mMTMWN15(mUMCMatchStats(lngPointer).IDIndex) - AMTMW(lngMassTagIndexOriginal)
            End If
        Else
            If NTypeStr = MOD_TKN_N14 Then
                udtPutUMCMatchParams.MassTagModMass.value = 0
            Else
                udtPutUMCMatchParams.MassTagModMass.value = glN14N15_DELTA * AMTCNT_N(lngMassTagIndexOriginal)
            End If
        End If
        
        If Len(strMassMods) > PUT_UMC_MATCH_MAX_MODSTRING_LENGTH Then strMassMods = Left(strMassMods, PUT_UMC_MATCH_MAX_MODSTRING_LENGTH)
        udtPutUMCMatchParams.MassTagMods.value = strMassMods
        
        cmdPutNewUMCMatch.Execute
        
        ExpCnt = ExpCnt + 1
    Next lngPointer

    If blnExportUMCsWithNoMatches Then
        ' Also export the UMC's that do not have any hits
        ReDim blnUMCMatchFound(GelUMC(CallerID).UMCCnt)
        
        ' First mark all of the UMC's that have already been exported
        For lngPointer = 0 To mMatchStatsCount - 1
            lngUMCIndexOriginal = mUMCMatchStats(lngPointer).UMCIndex
            blnUMCMatchFound(lngUMCIndexOriginal) = True
        Next lngPointer
        
        With GelUMC(CallerID)
            For lngUMCIndex = 0 To .UMCCnt - 1
                If lngUMCIndex Mod 25 = 0 Then
                    Me.Caption = "Exporting non-matching UMC's: " & Trim(lngUMCIndex) & " / " & Trim(.UMCCnt)
                    DoEvents
                End If
                
                If Not blnUMCMatchFound(lngUMCIndex) Then
                    ' No match was found; export to the database
                    lngPairIndex = -1
                    dblExpressionRatio = 0
                    dblExpressionRatioStDev = 0
                    intExpressionRatioChargeStateBasisCount = 0
                    lngExpressionRatioMemberBasisCount = 0
                    If eClsPaired(lngUMCIndex) <> umcpNone And blnPairsPresent Then
                        lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndex, objP1IndFastSearch, objP2IndFastSearch, (LastSearchTypeN14N15 = SEARCH_N15), dblExpressionRatio, dblExpressionRatioStDev, intExpressionRatioChargeStateBasisCount, lngExpressionRatioMemberBasisCount)
                    End If
                    
                    ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, CallerID, lngUMCIndex, 0, ClsStat(), FPR_Type_Standard, dblExpressionRatio, dblExpressionRatioStDev, intExpressionRatioChargeStateBasisCount, lngExpressionRatioMemberBasisCount, lngPairIndex
                
                End If
            Next lngUMCIndex
        End With
    End If

' MonroeMod
AddToAnalysisHistory CallerID, "Export to UMC Results table details: UMC Peaks Match Count = " & ExpCnt

Me.Caption = strCaptionSaved

strExportStatus = ExpCnt & " associations between mass tags and UMC's exported to UMC results table."
Set cmdPutNewUMC.ActiveConnection = Nothing
Set cmdPutNewUMCMatch.ActiveConnection = Nothing
cnNew.Close

If blnUpdateGANETForAnalysisInDB Then
    ' Export the the GANET Slope, Intercept, and Fit to the database
    With GelAnalysis(CallerID)
        strExportStatus = strExportStatus & vbCrLf & ExportGANETtoMTDB(CallerID, .GANET_Slope, .GANET_Intercept, .GANET_Fit)
    End With
End If

If blnAddQuantitationDescriptionEntry Then
    If lngErrorNumber = 0 And lngMDID >= 0 And ExpCnt > 0 Then
        ExportMTDBAddQuantitationDescriptionEntry Me, CallerID, ExpQuantitationDescription, lngMDID, lngErrorNumber, strIniFileName, 1, 1, 1, Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
    End If
End If

ExportMTDBbyUMCToUMCResultsTable = strExportStatus
lngErrorNumber = 0
Set objP1IndFastSearch = Nothing
Set objP2IndFastSearch = Nothing

Exit Function

err_ExportMTDBbyUMC:
ExportMTDBbyUMCToUMCResultsTable = "Error: " & Err.Number & vbCrLf & Err.Description
lngErrorNumber = Err.Number
On Error Resume Next
If Not cnNew Is Nothing Then cnNew.Close
Me.Caption = strCaptionSaved
Set objP1IndFastSearch = Nothing
Set objP2IndFastSearch = Nothing

End Function

Private Sub Info_NoMTDBLink()
'this message is used twice so ...
MsgBox "Current display is not associated with any Mass Tags database!" & vbCrLf _
     & "Close dialog and establish association (Edit->Select/Modify Database Connection)" & vbCrLf _
     & "or select Mass Tags->Load Legacy MT DB on this dialog to load" & vbCrLf _
     & "data from legacy database!", vbOKOnly, glFGTU
End Sub



Private Sub PickParameters()
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
Call txtAlkylationMWCorrection_LostFocus
Call txtNETFormula_LostFocus
End Sub

Private Sub PopulateComboBoxes()
    Dim intIndex As Integer
    
    With cboResidueToModify
        .Clear
        .AddItem "Full MT"
        For intIndex = 0 To 25
            .AddItem Chr(vbKeyA + intIndex)
        Next intIndex
        .AddItem glPHOSPHORYLATION
        .ListIndex = 0
    End With
    
End Sub

Private Function PrepareSearchN14() As Boolean
'---------------------------------------------------------------
'prepare search of N14 peptide (use loaded peptides masses)
'---------------------------------------------------------------
Dim qsd As New QSDouble
On Error Resume Next
If mMTCnt > 0 Then
   UpdateStatus "Preparing fast N14 search..."
   If qsd.QSAsc(mMTMWN14(), mMTInd()) Then
      Set MWFastSearch = New MWUtil
      If MWFastSearch.Fill(mMTMWN14()) Then PrepareSearchN14 = True
   End If
End If
End Function


Private Function PrepareSearchN15() As Boolean
'---------------------------------------------------------------
'prepare search of N15 peptide (use number of N to correct mass)
'---------------------------------------------------------------
Dim qsd As New QSDouble
On Error Resume Next
If mMTCnt > 0 Then
   UpdateStatus "Preparing fast N15 search..."
   If qsd.QSAsc(mMTMWN15(), mMTInd()) Then
      Set MWFastSearch = New MWUtil
      If MWFastSearch.Fill(mMTMWN15()) Then PrepareSearchN15 = True
   End If
End If
End Function


Private Function PrepareMTArrays() As Boolean
'---------------------------------------------------------------
'prepares masses from loaded mass tags based on specified
'modifications; returns True if succesful, False on any error
'---------------------------------------------------------------
Dim i As Long, j As Long
Dim TmpCnt As Long
Dim CysCnt As Long                 'Cysteine count in peptide
Dim CysLeft As Long                'Cysteine left for modification use
Dim CysUsedPEO As Long             'Cysteine already used in calculation for PEO
Dim CysUsedICAT_D0 As Long         'Cysteine already used in calculation for ICAT_D0
Dim CysUsedICAT_D8 As Long         'Cysteine already used in calculation for ICAT_D8

Dim strResiduesToModify As String   ' One or more residues to modify (single letter amino acid symbols)
Dim dblResidueModMass As Double
Dim ResidueOccurrenceCount As Integer
Dim strResModToken As String

On Error GoTo err_PrepareMTArrays

' Update GelSearchDef(CallerID).AMTSearchMassMods with the current settings
With GelSearchDef(CallerID).AMTSearchMassMods
    .PEO = cChkBox(chkPEO)
    .ICATd0 = cChkBox(chkICATLt)
    .ICATd8 = cChkBox(chkICATHv)
    .Alkylation = cChkBox(chkAlkylation)
    .AlkylationMass = CDblSafe(txtAlkylationMWCorrection)
    If cboResidueToModify.ListIndex > 0 Then
        .ResidueToModify = cboResidueToModify
    Else
        .ResidueToModify = ""
    End If
    
    .ResidueMassModification = CDblSafe(txtResidueToModifyMass)
    txtResidueToModifyMass = Round(.ResidueMassModification, 5)
    
    strResiduesToModify = .ResidueToModify
    dblResidueModMass = .ResidueMassModification
    
    .N15InsteadOfN14 = optN(SEARCH_N15).value
    .DynamicMods = optDBSearchModType(MODS_DYNAMIC).value
End With

If AMTCnt > 0 Then
   UpdateStatus "Preparing arrays for search..."
   'initially reserve space for AMTCnt peptides
   ReDim mMTInd(AMTCnt - 1)
   ReDim mMTOrInd(AMTCnt - 1)
   ReDim mMTMWN14(AMTCnt - 1)
   ReDim mMTMWN15(AMTCnt - 1)
   ReDim mMTNET(AMTCnt - 1)
   ReDim mMTMods(AMTCnt - 1)
   mMTCnt = 0
   For i = 1 To AMTCnt
       mMTCnt = mMTCnt + 1
       mMTInd(mMTCnt - 1) = mMTCnt - 1
       mMTOrInd(mMTCnt - 1) = i             'index; not the ID
       mMTMWN14(mMTCnt - 1) = AMTMW(i)
       mMTMWN15(mMTCnt - 1) = AMTMW(i) + glN14N15_DELTA * AMTCNT_N(i)       ' N15 is always fixed
       Select Case samtDef.NETorRT
       Case glAMT_NET
            mMTNET(mMTCnt - 1) = AMTNET(i)
       Case glAMT_RT
            mMTNET(mMTCnt - 1) = AMTRT(i)
       End Select
       mMTMods(mMTCnt - 1) = ""
   Next i
   If chkPEO.value = vbChecked Then         'correct based on cys number for PEO label
      UpdateStatus "Adding PEO labeled peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTCNT_Cys(mMTOrInd(i))
          If CysCnt > 0 Then
             If GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods Then
                ' Dynamic Mods
                For j = 1 To CysCnt
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * glPEO
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_PEO & "/" & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this mass tag
                mMTMWN14(i) = mMTMWN14(i) + CysCnt * glPEO
                mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                mMTMods(i) = mMTMods(i) & " " & MOD_TKN_PEO & "/" & CysCnt
             End If
          End If
      Next i
   End If
   'yeah, yeah I know that same cysteine can not be labeled with PEO and ICAT at the same
   'time but who cares anyway I can fix this here easily
   If chkICATLt.value = vbChecked Then         'correct based on cys number for ICAT label
      UpdateStatus "Adding D0 ICAT labeled peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTCNT_Cys(mMTOrInd(i))
          CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
          If CysUsedPEO < 0 Then CysUsedPEO = 0
          CysLeft = CysCnt - CysUsedPEO
          If CysLeft > 0 Then
             If GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods Then
                ' Dynamic Mods
                For j = 1 To CysLeft
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * glICAT_Light
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D0 & "/" & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this mass tag
                ' However, if use also has ICAT_d0 enabled, we need to duplicate this
                '  mass tag first
                If chkICATHv.value = vbChecked Then
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + CysLeft * glICAT_Heavy
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & CysLeft
                End If
                
                ' Now update this mass tag to have ICAT_d0 on all the cysteines
                mMTMWN14(i) = mMTMWN14(i) + CysLeft * glICAT_Light
                mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ICAT_D0 & "/" & CysLeft
             End If
          End If
      Next i
   End If
   
   If chkICATHv.value = vbChecked Then         'correct based on cys number for ICAT label
      UpdateStatus "Adding D8 ICAT labeled peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTCNT_Cys(mMTOrInd(i))
          CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
          If CysUsedPEO < 0 Then CysUsedPEO = 0
          CysUsedICAT_D0 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D0)
          If CysUsedICAT_D0 < 0 Then CysUsedICAT_D0 = 0
          CysUsedICAT_D8 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D8)
          If CysUsedICAT_D8 < 0 Then CysUsedICAT_D8 = 0
          CysLeft = CysCnt - CysUsedPEO - CysUsedICAT_D0 - CysUsedICAT_D8
          If CysLeft > 0 Then
             If GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods Then
                ' Dynamic Mods
                For j = 1 To CysLeft
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * glICAT_Heavy
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & j
                Next j
             Else
                If chkICATLt.value = vbChecked Then
                    ' We shouldn't have reached this code since all of the cysteines should
                    '  have been assigned ICAT_d0 or ICAT_d8
                    Debug.Assert False
                Else
                    ' Static Mods
                    ' Simply update the stats for this mass tag
                    mMTMWN14(i) = mMTMWN14(i) + CysLeft * glICAT_Heavy
                    mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                    mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & CysLeft
                End If
             End If
          End If
      Next i
   End If
   
   If chkAlkylation.value = vbChecked Then         'correct based on cys number for alkylation label
      UpdateStatus "Adding alkylated peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTCNT_Cys(mMTOrInd(i))
          CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
          If CysUsedPEO < 0 Then CysUsedPEO = 0
          CysUsedICAT_D0 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D0)
          If CysUsedICAT_D0 < 0 Then CysUsedICAT_D0 = 0
          CysUsedICAT_D8 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D8)
          If CysUsedICAT_D8 < 0 Then CysUsedICAT_D8 = 0
          CysLeft = CysCnt - CysUsedPEO - CysUsedICAT_D0 - CysUsedICAT_D8
          If CysLeft > 0 Then
             If GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods Then
                ' Dynamic Mods
                For j = 1 To CysLeft
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * AlkMWCorrection
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ALK & "/" & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this mass tag
                mMTMWN14(i) = mMTMWN14(i) + CysLeft * AlkMWCorrection
                mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ALK & "/" & CysLeft
             End If
          End If
      Next i
   End If
   
   If dblResidueModMass <> 0 Then
      UpdateStatus "Adding modified residue mass peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
        
        If Len(strResiduesToModify) > 0 Then
          ResidueOccurrenceCount = LookupResidueOccurrence(mMTOrInd(i), strResiduesToModify)
          
          If InStr(strResiduesToModify, "C") > 0 Then
            CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
            If CysUsedPEO < 0 Then CysUsedPEO = 0
            CysUsedICAT_D0 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D0)
            If CysUsedICAT_D0 < 0 Then CysUsedICAT_D0 = 0
            CysUsedICAT_D8 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D8)
            If CysUsedICAT_D8 < 0 Then CysUsedICAT_D8 = 0
            ResidueOccurrenceCount = ResidueOccurrenceCount - CysUsedPEO - CysUsedICAT_D0 - CysUsedICAT_D8
          End If
          strResModToken = MOD_TKN_RES_MOD
        Else
          ' Add dblResidueModMass once to the entire mass tag
          ' Accomplish this by setting ResidueOccurrenceCount to 1
          ResidueOccurrenceCount = 1
          strResModToken = MOD_TKN_MT_MOD
        End If
        
        If ResidueOccurrenceCount > 0 Then
           If GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods Then
              ' Dynamic Mods
              For j = 1 To ResidueOccurrenceCount
                  mMTCnt = mMTCnt + 1
                  mMTInd(mMTCnt - 1) = mMTCnt - 1
                  mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                  mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * dblResidueModMass
                  mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
                  mMTNET(mMTCnt - 1) = mMTNET(i)
                  mMTMods(mMTCnt - 1) = mMTMods(i) & " " & strResModToken & "/" & strResiduesToModify & j
              Next j
           Else
              ' Static Mods
              ' Simply update the stats for this mass tag
              mMTMWN14(i) = mMTMWN14(i) + ResidueOccurrenceCount * dblResidueModMass
              mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTCNT_N(mMTOrInd(i))
              mMTMods(i) = mMTMods(i) & " " & strResModToken & "/" & strResiduesToModify & ResidueOccurrenceCount
           End If
        End If
      Next i
   End If
   
   If mMTCnt > 0 Then
      UpdateStatus "Preparing fast search structures..."
      ReDim Preserve mMTInd(mMTCnt - 1)
      ReDim Preserve mMTOrInd(mMTCnt - 1)
      ReDim Preserve mMTMWN14(mMTCnt - 1)
      ReDim Preserve mMTMWN15(mMTCnt - 1)
      ReDim Preserve mMTNET(mMTCnt - 1)
      ReDim Preserve mMTMods(mMTCnt - 1)
      Select Case N14N15
      Case SEARCH_N14
           If Not PrepareSearchN14() Then
              Debug.Assert False
              Call DestroySearchStructures
              Exit Function
           End If
      Case SEARCH_N15
           If Not PrepareSearchN15() Then
              Debug.Assert False
              Call DestroySearchStructures
              Exit Function
           End If
      End Select
   Else
      Call DestroySearchStructures
   End If

End If
PrepareMTArrays = True
Exit Function

err_PrepareMTArrays:
Select Case Err.Number
Case 9                      'add space in chunks of 10000
   ReDim Preserve mMTInd(mMTCnt + 10000)
   ReDim Preserve mMTOrInd(mMTCnt + 10000)
   ReDim Preserve mMTMWN14(mMTCnt + 10000)
   ReDim Preserve mMTMWN15(mMTCnt + 10000)
   ReDim Preserve mMTNET(mMTCnt + 10000)
   ReDim Preserve mMTMods(mMTCnt + 10000)
   Resume
Case Else
   Debug.Assert False
   Call DestroySearchStructures
End Select
End Function

Private Function GetTokenValue(ByVal s As String, ByVal t As String) As Long
'---------------------------------------------------------------------------
'returns value next to token T in string of type Token1/Value1 Token2/Value2
'-1 if not found or on any error
'---------------------------------------------------------------------------
Dim SSplit() As String
Dim MSplit() As String
Dim i As Long
On Error GoTo exit_GetTokenValue
GetTokenValue = -1

SSplit = Split(s, " ")
For i = 0 To UBound(SSplit)
    If Len(SSplit(i)) > 0 Then
        If InStr(SSplit(i), "/") > 0 Then
            MSplit = Split(SSplit(i), "/")
            If Trim$(MSplit(0)) = t Then
               If IsNumeric(MSplit(1)) Then
                  GetTokenValue = CLng(MSplit(1))
                  Exit Function
               End If
            End If
        End If
    End If
Next i
Exit Function

exit_GetTokenValue:
Debug.Assert False

End Function

Private Sub CheckNETEquationStatus()
    If Not GelAnalysis(CallerID) Is Nothing Then
        If txtNETFormula.Text = ConstructNETFormula(GelAnalysis(CallerID).GANET_Slope, GelAnalysis(CallerID).GANET_Intercept) _
           And InStr(UCase(txtNETFormula), "MINFN") = 0 Then
            mUsingDefaultGANET = True
        Else
            mUsingDefaultGANET = False
        End If
    Else
        mUsingDefaultGANET = False
    End If
    
End Sub

Private Function ConvertScanToNET(lngScanNumber As Long) As Double

    If mUsingDefaultGANET Then
        Debug.Assert InStr(UCase(txtNETFormula), "MINFN") = 0
        ConvertScanToNET = ScanToGANET(CallerID, lngScanNumber)
    Else
        ConvertScanToNET = Elution(lngScanNumber, MinFN, MaxFN)
    End If

End Function

Private Sub DestroySearchStructures()
On Error Resume Next
mMTCnt = 0
Erase mMTInd
Erase mMTOrInd
Erase mMTMWN14
Erase mMTMWN15
Erase mMTNET
Erase mMTMods
Set MWFastSearch = Nothing
End Sub

Private Function DestroyIDStructures()
On Error Resume Next
mMatchStatsCount = 0
Erase mUMCMatchStats
Call ManageCurrID(MNG_ERASE)
End Function

Private Function SearchUMC(ByVal ClassInd As Long) As Long
'-----------------------------------------------------------------------------
'returns number of hits found for UMC with index ClassInd; -1 in case of error
'-----------------------------------------------------------------------------
Dim CurrMW As Double
Dim AbsMWErr As Double
Dim CurrScan As Long
Dim IsMatch As Boolean
Dim i As Long, j As Long
Dim MatchInd1 As Long, MatchInd2 As Long
On Error GoTo err_SearchUMC
If ManageCurrID(MNG_RESET) Then
   If SearchType = SEARCH_PAIRED Then
      Select Case N14N15
      Case SEARCH_N14     'don't search if this class is found only as heavy member
        If eClsPaired(ClassInd) = umcpHeavyUnique Or _
           eClsPaired(ClassInd) = umcpHeavyMultiple Then Exit Function
      Case SEARCH_N15     'don't search if this class is found only as light member
        If eClsPaired(ClassInd) = umcpLightUnique Or _
           eClsPaired(ClassInd) = umcpLightMultiple Then Exit Function
      End Select
   End If
   With GelUMC(CallerID).UMCs(ClassInd)
      For i = 0 To .ClassCount - 1
        Select Case .ClassMType(i)
        Case glCSType
             CurrMW = GelData(CallerID).CSNum(.ClassMInd(i), csfMW)
             CurrScan = GelData(CallerID).CSNum(.ClassMInd(i), csfScan)
        Case glIsoType
             CurrMW = GelData(CallerID).IsoNum(.ClassMInd(i), samtDef.MWField)
             CurrScan = GelData(CallerID).IsoNum(.ClassMInd(i), isfScan)
        End Select
        Select Case samtDef.TolType
        Case gltPPM
           AbsMWErr = CurrMW * samtDef.MWTol * glPPM
        Case gltABS
           AbsMWErr = samtDef.MWTol
        Case Else
           Debug.Assert False
        End Select
        MatchInd1 = 0
        MatchInd2 = -1
        If MWFastSearch.FindIndexRange(CurrMW, AbsMWErr, MatchInd1, MatchInd2) Then
           If MatchInd1 <= MatchInd2 Then
              For j = MatchInd1 To MatchInd2
                  IsMatch = True
                  If samtDef.NETTol >= 0 Then
                     If Abs(ConvertScanToNET(CurrScan) - mMTNET(mMTInd(j))) _
                        > samtDef.NETTol Then IsMatch = False
                  End If
                  If IsMatch Then
                     mgCurrIDCnt = mgCurrIDCnt + 1
                     mgCurrIDInd(mgCurrIDCnt - 1) = j
                  End If
              Next j
           End If
        End If
      Next i
   End With
   If mgCurrIDCnt > 0 Then
      If ManageCurrID(MNG_TRIM) Then
         Call ScoreCurrIDs
         If mgCurrIDCnt > 0 Then
            Call AddCurrIDsToAllIDs(ClassInd)
         End If
      End If
   Else
      Call ManageCurrID(MNG_ERASE)
      SearchUMC = 0
   End If
Else
   UpdateStatus "Error managing memory!"
End If

err_SearchUMC:
Select Case Err.Number
Case 9                  'reserve additional memory
     If ManageCurrID(MNG_ADD_START_SIZE) Then Resume
End Select
SearchUMC = -1
End Function

Private Function ManageCurrID(ByVal ManageType As Long) As Boolean
On Error GoTo exit_ManageCurrID
Select Case ManageType
Case MNG_ERASE
     mgCurrIDCnt = 0
     Erase mgCurrIDInd
     Erase mgCurrIDScore
Case MNG_TRIM
     If mgCurrIDCnt > 0 Then
        ReDim Preserve mgCurrIDInd(mgCurrIDCnt - 1)
        ReDim Preserve mgCurrIDScore(mgCurrIDCnt - 1)
     Else
        ManageCurrID = ManageCurrID(MNG_ERASE)
     End If
Case MNG_RESET
     mgCurrIDCnt = 0
     ReDim mgCurrIDInd(MNG_START_SIZE)
     ReDim mgCurrIDScore(MNG_START_SIZE)
Case MNG_ADD_START_SIZE
     ReDim Preserve mgCurrIDInd(mgCurrIDCnt + MNG_START_SIZE)
     ReDim Preserve mgCurrIDScore(mgCurrIDCnt + MNG_START_SIZE)
Case Else
     If ManageType > 0 Then
        ReDim Preserve mgCurrIDInd(mgCurrIDCnt + ManageType)
        ReDim Preserve mgCurrIDScore(mgCurrIDCnt + ManageType)
     End If
End Select
ManageCurrID = True
exit_ManageCurrID:
End Function

Private Function AddCurrIDsToAllIDs(ClsInd As Long) As Boolean
'---------------------------------------------------------------------------
'returns True if successful; adds current identifications to list of all IDs
'---------------------------------------------------------------------------
Dim i As Long, lngTargetIndex
On Error GoTo err_AddCurrIDsToAllIDs
mMatchStatsCount = mMatchStatsCount + mgCurrIDCnt
ReDim Preserve mUMCMatchStats(mMatchStatsCount - 1)
For i = 0 To mgCurrIDCnt - 1
    lngTargetIndex = mMatchStatsCount - i - 1
    mUMCMatchStats(lngTargetIndex).UMCIndex = ClsInd
    mUMCMatchStats(lngTargetIndex).IDIndex = mgCurrIDInd(i)
    mUMCMatchStats(lngTargetIndex).MemberHitCount = mgCurrIDScore(i)
    mUMCMatchStats(lngTargetIndex).MultiAMTHitCount = mgCurrIDCnt
Next i
AddCurrIDsToAllIDs = True
err_AddCurrIDsToAllIDs:
End Function

Private Sub txtResidueToModifyMass_LostFocus()
    ValidateTextboxValueDbl txtResidueToModifyMass, -10000, 10000, 0
End Sub

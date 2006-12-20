VERSION 5.00
Begin VB.Form frmSearchMT_ConglomerateUMC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Search MT tag DB - Single UMC Mass"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   7605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSetDefaultsForToleranceRefinement 
      Cancel          =   -1  'True
      Caption         =   "Set to Tolerance Refinement Defaults"
      Height          =   495
      Left            =   5640
      TabIndex        =   54
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtDBSearchMinimumPeptideProphetProbability 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Text            =   "0"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdSetDefaults 
      Caption         =   "Set to Defaults"
      Height          =   375
      Left            =   5640
      TabIndex        =   52
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtUniqueMatchStats 
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   6525
      Width           =   7455
   End
   Begin VB.TextBox txtDBSearchMinimumHighDiscriminantScore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Text            =   "0"
      Top             =   2340
      Width           =   615
   End
   Begin VB.ComboBox cboAMTSearchResultsBehavior 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtDBSearchMinimumHighNormalizedScore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.ComboBox cboInternalStdSearchMode 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkUpdateGelDataWithSearchResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update data in current file with results of search"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   600
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearchAllUMCs 
      Caption         =   "Search All UMC's"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemoveAMTMatchesFromUMCs 
      Caption         =   "Remove existing MT matches from UMC's"
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Remove MT reference for current gel"
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame fraMods 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modifications"
      Height          =   1575
      Left            =   0
      TabIndex        =   29
      Top             =   4560
      Width           =   6735
      Begin VB.ComboBox cboResidueToModify 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtResidueToModifyMass 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   39
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame fraOptionFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Index           =   47
         Left            =   5760
         TabIndex        =   44
         Top             =   360
         Width           =   800
         Begin VB.OptionButton optN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N14"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   46
            Top             =   240
            Value           =   -1  'True
            Width           =   700
         End
         Begin VB.OptionButton optN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N15"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   47
            Top             =   525
            Width           =   700
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "N Type:"
            Height          =   255
            Index           =   103
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.Frame fraOptionFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Index           =   49
         Left            =   4460
         TabIndex        =   40
         Top             =   360
         Width           =   1095
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dynamic"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   525
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fixed"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Mod Type:"
            Height          =   255
            Index           =   100
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.TextBox txtAlkylationMWCorrection 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   35
         Text            =   "57.0215"
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkAlkylation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alkylation"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         ToolTipText     =   "Check to add the alkylation mass correction below to all MT Tag masses (added to each cys residue)"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkICATHv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICAT d8"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkICATLt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICAT d0"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkPEO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PEO"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Residue to modify:"
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mass (Da):"
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5640
         X2              =   5640
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
         Index           =   0
         X1              =   1320
         X2              =   1320
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alkylation mass:"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame fraNET 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NET  Calculation"
      Height          =   1455
      Left            =   0
      TabIndex        =   15
      Top             =   3000
      Width           =   5175
      Begin VB.CheckBox chkDisableCustomNETs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable NETs from Warping"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   480
         Width           =   2400
      End
      Begin VB.CheckBox chkUseUMCConglomerateNET 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Class NET for UMCs"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         ToolTipText     =   $"frmSearchMT_ConglomerateUMC.frx":0000
         Top             =   240
         Width           =   2205
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   23
         Text            =   "0.1"
         Top             =   1020
         Width           =   615
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pred. NET for MT Tags"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Use Predicted NET values for the MT tags"
         Top             =   480
         Width           =   2500
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Obs. NET for MT Tags"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Use NET calculated from all peptide observations for each MT tag"
         Top             =   240
         Value           =   -1  'True
         Width           =   2500
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   1020
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "NET T&olerance"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label lblNETFormula 
         BackStyle       =   0  'Transparent
         Caption         =   "Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   810
         Width           =   2415
      End
   End
   Begin VB.Frame fraMWTolerance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1455
      Left            =   5280
      TabIndex        =   24
      Top             =   3000
      Width           =   2175
      Begin VB.ComboBox cboSearchRegionShape 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   160
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Peptide Prophet Probability"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2655
      Width           =   2865
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum PMT Discriminant Score"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2360
      Width           =   2505
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum PMT XCorr"
      Height          =   255
      Index           =   134
      Left            =   120
      TabIndex        =   9
      Top             =   2060
      Width           =   2145
   End
   Begin VB.Label lblInternalStdSearchMode 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Standard Search Mode:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1725
      Width           =   2415
   End
   Begin VB.Label lblUMCMassMode 
      BackStyle       =   0  'Transparent
      Caption         =   "UMC Mass = ??"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblETType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Generic NET"
      Height          =   255
      Left            =   5280
      TabIndex        =   49
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   6240
      Width           =   4935
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuF 
      Caption         =   "&Function"
      Begin VB.Menu mnuFSearchAll 
         Caption         =   "Search &All UMCs"
      End
      Begin VB.Menu mnuFSearchPaired 
         Caption         =   "Search Paired UMCs (skips excluded pairs)"
      End
      Begin VB.Menu mnuFSearchPairedPlusNonPaired 
         Caption         =   "Search Light Members of Pairs &Plus Non-paired UMCs (skips excluded)"
      End
      Begin VB.Menu mnuFSearchNonPaired 
         Caption         =   "Search &Non-paired UMCs"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExcludeAmbiguous 
         Caption         =   "Exclude Ambiguous Pairs (all pairs)"
      End
      Begin VB.Menu mnuFExcludeAmbiguousHitsOnly 
         Caption         =   "Exclude Ambiguous Pairs (only those with hits)"
      End
      Begin VB.Menu mnuFResetExclusionFlags 
         Caption         =   "Reset Exclusion Flags for All Pairs"
      End
      Begin VB.Menu mnuFDeleteExcludedPairs 
         Caption         =   "Delete Excluded Pairs"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFReportByUMC 
         Caption         =   "Report Results by &UMCs..."
      End
      Begin VB.Menu mnuFReportByIon 
         Caption         =   "Report Results by &Ions..."
      End
      Begin VB.Menu mnuFReportIncludeORFs 
         Caption         =   "Include ORFs in Report"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFSepExportToDatabase 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExportResultsToDBbyUMC 
         Caption         =   "Export Results to MT Tag DB (by UMC)"
      End
      Begin VB.Menu mnuFExportDetailedMemberInformation 
         Caption         =   "Export detailed member information for each UMC"
      End
      Begin VB.Menu mnuFSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFMassCalAndToleranceRefinement 
         Caption         =   "&Mass Calibration and Tolerance Refinement"
      End
      Begin VB.Menu mnuFSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMT 
      Caption         =   "&MT Tags"
      Begin VB.Menu mnuMTLoadMT 
         Caption         =   "Load MT Tag DB"
      End
      Begin VB.Menu mnuMTLoadLegacy 
         Caption         =   "Load Legacy MT DB"
      End
      Begin VB.Menu mnuMTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTStatus 
         Caption         =   "MT Tags Status"
      End
   End
   Begin VB.Menu mnuETHeader 
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
Attribute VB_Name = "frmSearchMT_ConglomerateUMC"
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
Const SEARCH_PAIRED_PLUS_NON_PAIRED = 3

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
Private mMTOrInd() As Long      'index of original MT tag (in AMT array)
Private mMTMWN14() As Double    'mass to look for N14
Private mMTMWN15() As Double    'mass to look for N15
Private mMTNET() As Double      'NET value
Private mMTMods() As String     'modification description

Private MWFastSearch As MWUtil

Private mInternalStdIndexPointers() As Long             ' Pointer to entry in UMCInternalStandards.InternalStandards()
Private InternalStdFastSearch As MWUtil

Private AlkMWCorrection As Double
Private N14N15 As Long                  ' SEARCH_N14 or SEARCH_N15
Private SearchType As Long              ' SEARCH_ALL, SEARCH_PAIRED, SEARCH_NON_PAIRED, or SEARCH_PAIRED_PLUS_NON_PAIRED
Private mSearchRegionShape As srsSearchRegionShapeConstants

Private LastSearchTypeN14N15 As Long
Private NTypeStr As String

'following arrays are parallel to the UMCs
Private ClsCnt As Long              'this is not actually neccessary except
Private ClsStat() As Double         'to create nice reports; necessary to use this since we report the Min/Max Charge stats and Average Fit stats
Private eClsPaired() As umcpUMCPairMembershipConstants      ' Keeps track of whether UMC is member of 1 or more pairs

                                
'mUMCMatchStats contains all possible identifications for all UMCs with scores
'as count of each identification hits within the UMC
Private mMatchStatsCount As Long                                'count of UMC-ID matches
Private mUMCMatchStats() As udtUMCMassTagMatchStats             ' 0-based array
Private mSearchUsedCustomNETs As Boolean

' The following hold match stats for each individual UMC
Private mCurrIDCnt As Long
Private mCurrIDMatches() As udtUMCMassTagRawMatches          ' 0-based array

'Expression Evaluator variables for elution time calculation
Private MyExprEva As ExprEvaluator
Private VarVals() As Long
Private MinFN As Long
Private MaxFN As Long

Private ExpAnalysisSPName As String             ' Stored procedure AddMatchMaking
''Private ExpPeakSPName As String               ' Stored procedure AddFTICRPeak; Unused variable
Private ExpUmcSPName As String                  ' Stored procedure AddFTICRUmc
Private ExpUMCMemberSPName As String            ' Stored procedure AddFTICRUmcMember
Private ExpUmcMatchSPName As String             ' Stored procedure AddFTICRUmcMatch
Private ExpUmcInternalStdMatchSPName As String  ' Stored procedure AddFTICRUmcInternalStdMatch
Private ExpQuantitationDescription As String    ' Stored procedure AddQuantitationDescription

Private mUMCCountSkippedSinceRefPresent As Long
Private mUsingDefaultGANET As Boolean
Private eInternalStdSearchMode As issmInternalStandardSearchModeConstants
Private mMTMinimumHighNormalizedScore As Single
Private mMTMinimumHighDiscriminantScore As Single
Private mMTMinimumPeptideProphetProbability As Single

Private mMDTypeSaved As Long

Private mKeyPressAbortProcess As Integer

Private objMTDBNameLookupClass As mtdbMTNames
'

Public Property Get SearchRegionShape() As srsSearchRegionShapeConstants
    SearchRegionShape = mSearchRegionShape
End Property
Public Property Let SearchRegionShape(Value As srsSearchRegionShapeConstants)
    cboSearchRegionShape.ListIndex = Value
    mSearchRegionShape = Value
End Property

Private Function AddCurrIDsToAllIDs(ClsInd As Long) As Boolean
'---------------------------------------------------------------------------
'returns True if successful; adds current identifications to list of all IDs
'---------------------------------------------------------------------------
Dim lngIndex As Long, lngTargetIndex As Long
Dim lngAMTHitCount As Long

On Error GoTo err_AddCurrIDsToAllIDs
mMatchStatsCount = mMatchStatsCount + mCurrIDCnt
ReDim Preserve mUMCMatchStats(mMatchStatsCount - 1)

' Count the number of non Internal Standard matches
lngAMTHitCount = 0
For lngIndex = 0 To mCurrIDCnt - 1
    If Not mCurrIDMatches(lngIndex).IDIsInternalStd Then
        lngAMTHitCount = lngAMTHitCount + 1
    End If
Next lngIndex

For lngIndex = 0 To mCurrIDCnt - 1
    lngTargetIndex = (mMatchStatsCount - mCurrIDCnt) + lngIndex
    With mUMCMatchStats(lngTargetIndex)
        .UMCIndex = ClsInd
        .IDIndex = mCurrIDMatches(lngIndex).IDInd
        .MemberHitCount = mCurrIDMatches(lngIndex).MatchingMemberCount
        .SLiCScore = mCurrIDMatches(lngIndex).SLiCScore
        .DelSLiC = mCurrIDMatches(lngIndex).DelSLiC
        .IDIsInternalStd = mCurrIDMatches(lngIndex).IDIsInternalStd
        .MultiAMTHitCount = lngAMTHitCount
    End With
Next lngIndex
AddCurrIDsToAllIDs = True
err_AddCurrIDsToAllIDs:
End Function

Private Sub CheckNETEquationStatus()
    If RobustNETValuesEnabled() Then
        mUsingDefaultGANET = True
    Else
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
    End If
End Sub

Private Function CountMassTagsInUMCMatchStats() As Long
    ' Returns the number of items in mUMCMatchStats() with .IDIsInternalStd = False
    
    Dim lngMassTagHitCount As Long
    Dim lngIndex As Long
    
    lngMassTagHitCount = 0
    For lngIndex = 0 To mMatchStatsCount - 1
        If Not mUMCMatchStats(lngIndex).IDIsInternalStd Then lngMassTagHitCount = lngMassTagHitCount + 1
    Next lngIndex
    
    CountMassTagsInUMCMatchStats = lngMassTagHitCount

End Function

Private Function ConvertScanToNET(lngScanNumber As Long) As Double

    If mUsingDefaultGANET Then
        ConvertScanToNET = ScanToGANET(CallerID, lngScanNumber)
    Else
        ConvertScanToNET = Elution(lngScanNumber, MinFN, MaxFN)
    End If

End Function

Public Function DeleteExcludedPairsWrapper()
    Dim strMessage As String
    strMessage = DeleteExcludedPairs(CallerID)
    AddToAnalysisHistory CallerID, strMessage
    
    UpdateUMCsPairingStatusNow
End Function

Private Sub DestroyIDStructures()
On Error Resume Next
mMatchStatsCount = 0
Erase mUMCMatchStats
Call ManageCurrID(MNG_ERASE)
End Sub

Private Sub DestroySearchStructures()
    On Error Resume Next
    mMTCnt = 0
    Erase mMTInd
    Erase mMTOrInd
    Erase mMTMWN14
    Erase mMTMWN15
    Erase mMTNET
    Erase mMTMods
    Erase mInternalStdIndexPointers
    Set MWFastSearch = Nothing
    Set InternalStdFastSearch = Nothing
End Sub

Private Sub DisplayCurrentSearchTolerances()
    With samtDef
        txtMWTol.Text = .MWTol
    
        Select Case .TolType
        Case gltPPM
          optTolType(0).Value = True
        Case gltABS
          optTolType(1).Value = True
        Case Else
          Debug.Assert False
        End Select
        
        'NETTol is used both for NET and RT
        If .NETTol >= 0 Then
           txtNETTol.Text = .NETTol
           txtNETTol_Validate False
        Else
           txtNETTol.Text = ""
        End If
    End With
End Sub

Private Sub GenerateUniqueMatchStats(ByRef lngUniqueUMCCount As Long, ByRef lngUniquePMTTagCount As Long, ByRef lngUniqueInternalStdCount As Long)
    ' Determine the number of UMCs with at least one match,
    ' the unique number of MT tags matched, and the unique number of Internal Standards matched
    
    Dim blnUMCHasMatch() As Boolean
    Dim blnPMTTagMatched() As Boolean
    Dim blnInternalStdMatched() As Boolean
    
    Dim lngIndex As Long
    Dim lngUMCIndexOriginal As Long
    Dim lngInternalStdIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long
    Dim lngMassTagIndexOriginal As Long
    
    ReDim blnUMCHasMatch(GelUMC(CallerID).UMCCnt - 1) As Boolean
    ReDim blnPMTTagMatched(AMTCnt) As Boolean
    If UMCInternalStandards.Count > 0 Then
        ReDim blnInternalStdMatched(UMCInternalStandards.Count - 1) As Boolean
    End If
    
    For lngIndex = 0 To mMatchStatsCount - 1
        lngUMCIndexOriginal = mUMCMatchStats(lngIndex).UMCIndex
        If lngUMCIndexOriginal < GelUMC(CallerID).UMCCnt Then
            blnUMCHasMatch(lngUMCIndexOriginal) = True
        Else
            ' Invalid UMC index
            Debug.Assert False
        End If
        
        If mUMCMatchStats(lngIndex).IDIsInternalStd Then
            lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(lngIndex).IDIndex)
            If lngInternalStdIndexOriginal < UMCInternalStandards.Count Then
                blnInternalStdMatched(lngInternalStdIndexOriginal) = True
            Else
                ' Invalid Internal Standard index
                Debug.Assert False
            End If
        Else
            lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngIndex).IDIndex)
            lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            If lngMassTagIndexOriginal <= AMTCnt Then
                blnPMTTagMatched(lngMassTagIndexOriginal) = True
            Else
                ' Invalid MT tag index
                Debug.Assert False
            End If
    End If
    Next lngIndex
    
    lngUniqueUMCCount = 0
    For lngIndex = 0 To UBound(blnUMCHasMatch)
        If blnUMCHasMatch(lngIndex) Then lngUniqueUMCCount = lngUniqueUMCCount + 1
    Next lngIndex
    
    lngUniquePMTTagCount = 0
    For lngIndex = 0 To UBound(blnPMTTagMatched)
        If blnPMTTagMatched(lngIndex) Then lngUniquePMTTagCount = lngUniquePMTTagCount + 1
    Next lngIndex
    
    lngUniqueInternalStdCount = 0
    If UMCInternalStandards.Count > 0 Then
        For lngIndex = 0 To UBound(blnInternalStdMatched)
            If blnInternalStdMatched(lngIndex) Then lngUniqueInternalStdCount = lngUniqueInternalStdCount + 1
        Next lngIndex
    End If
    
End Sub

Private Function DisplayHitSummary(strSearchScope As String) As String

    Dim strMessage As String
    Dim strStats As String
    Dim strSearchItems As String
    Dim strModMassDescription As String
    
    Dim lngUniqueUMCCount As Long
    Dim lngUniquePMTTagCount As Long
    Dim lngUniqueInternalStdCount As Long
    Dim sngUMCMatchPercentage As Single
    
    strMessage = "Hits: " & LongToStringWithCommas(mMatchStatsCount)
    Select Case eInternalStdSearchMode
    Case issmFindWithMassTags
        strSearchItems = "MT tags and/or Internal Stds"
    Case issmFindOnlyInternalStandards
        strSearchItems = "Internal Stds"
    Case Else
        ' Includes issmFindOnlyMassTags
        strSearchItems = "MT tags"
    End Select
    strMessage = strMessage & " " & strSearchItems
    
    ' Determine the unique number of UMCs with matches, the unique MT tag count, and the unique Internal Standard Count
    
    If mUMCCountSkippedSinceRefPresent > 0 Then
        strMessage = strMessage & " (" & Trim(mUMCCountSkippedSinceRefPresent) & " UMC's skipped)"
    End If
    
    UpdateStatus strMessage
    
    GelSearchDef(CallerID).AMTSearchOnUMCs = samtDef
    
    AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched " & strSearchScope & " UMC's for " & strSearchItems & " (searched by UMC conglomerate mass, " & lblUMCMassMode & "; however, all members of a UMC are assigned all matches found for the UMC)", mMatchStatsCount, mMTMinimumHighNormalizedScore, mMTMinimumHighDiscriminantScore, mMTMinimumPeptideProphetProbability, samtDef, True, mSearchUsedCustomNETs)
    
    strModMassDescription = ConstructMassTagModMassDescription(GelSearchDef(CallerID).AMTSearchMassMods)
    If Len(strModMassDescription) > 0 Then
        AddToAnalysisHistory CallerID, strModMassDescription
    End If

    GenerateUniqueMatchStats lngUniqueUMCCount, lngUniquePMTTagCount, lngUniqueInternalStdCount
    If GelUMC(CallerID).UMCCnt > 0 Then
        sngUMCMatchPercentage = lngUniqueUMCCount / CSng(GelUMC(CallerID).UMCCnt) * 100#
    Else
        sngUMCMatchPercentage = 0
    End If
    
    strStats = "UMCs with match = " & LongToStringWithCommas(lngUniqueUMCCount) & " (" & Trim(Round(sngUMCMatchPercentage, 0)) & "%)"
    If eInternalStdSearchMode <> issmFindOnlyInternalStandards Then
        strStats = strStats & "; Unique MT tags matched = " & LongToStringWithCommas(lngUniquePMTTagCount) & " / " & LongToStringWithCommas(mMTCnt)
        If mMTCnt > AMTCnt Then
            strStats = strStats & " (" & LongToStringWithCommas(AMTCnt) & " source MT tags)"
        End If
    End If
    
    If eInternalStdSearchMode <> issmFindOnlyMassTags Then
        strStats = strStats & "; Unique Int Stds = " & LongToStringWithCommas(lngUniqueInternalStdCount) & " / " & LongToStringWithCommas(UMCInternalStandards.Count)
    End If
    
    txtUniqueMatchStats.Text = strStats
    
    AddToAnalysisHistory CallerID, "Match stats: " & strStats

    DisplayHitSummary = strMessage

End Function

Private Function Elution(FN As Long, MinFN As Long, MaxFN As Long) As Double
'---------------------------------------------------
'this function does not care are we using NET or RT
'---------------------------------------------------
VarVals(1) = FN
VarVals(2) = MinFN
VarVals(3) = MaxFN
Elution = MyExprEva.ExprVal(VarVals())
End Function

Private Sub EnableDisableNETFormulaControls()
    Dim i As Integer
    
    txtNETFormula.Enabled = Not RobustNETValuesEnabled()
    lblNETFormula.Enabled = txtNETFormula.Enabled
    mnuETHeader.Enabled = txtNETFormula.Enabled
    
    If RobustNETValuesEnabled() Then
        lblETType.Caption = "Using Custom NETs"
    Else
        For i = mnuET.LBound To mnuET.UBound
            If mnuET(i).Checked Then
               lblETType.Caption = "ET: " & mnuET(i).Caption
               SetETMode val(i)
            End If
        Next i
    End If
End Sub

Public Sub ExcludeAmbiguousPairsWrapper(blnOnlyExaminePairsWithHits As Boolean)
    Dim strMessage As String
    
    If blnOnlyExaminePairsWithHits Then
        strMessage = PairsSearchMarkAmbiguousPairsWithHitsOnly(Me, CallerID)
    Else
        strMessage = PairsSearchMarkAmbiguous(Me, CallerID, True)
    End If
    
    UpdateUMCsPairingStatusNow
    UpdateStatus strMessage
End Sub

Public Function ExportMTDBbyUMC(Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional blnExportUMCMembers As Boolean = False, Optional strIniFileName As String = "", Optional ByRef lngErrorNumber As Long, Optional ByRef lngMDID As Long) As String
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
    cmdRemoveAMTMatchesFromUMCs.Visible = False
        
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        With glbPreferencesExpanded.AutoAnalysisOptions
            blnAddQuantitationEntry = .AddQuantitationDescriptionEntry
            blnExportUMCsWithNoMatches = .ExportUMCsWithNoMatches
        End With
    Else
        eResponse = MsgBox("Export UMC's that do not have any database matches?", vbQuestion + vbYesNo + vbDefaultButton2, "Export Non-Matching UMC's")
        blnExportUMCsWithNoMatches = (eResponse = vbYes)
    End If
    
    ' September 2004: Unsupported code
    ''strStatus = ExportMTDBbyUMCToPeakResultsTable(lngMDID, blnUpdateGANETForAnalysisInDB, lngErrorNumber)
    
    ' Note: The following function call will create a new entry in T_Match_Making_Description
    strStatus = strStatus & vbCrLf & ExportMTDBbyUMCToUMCResultsTable(lngMDID, True, blnUpdateGANETForAnalysisInDB, blnExportUMCMembers, lngErrorNumber, blnAddQuantitationEntry, blnExportUMCsWithNoMatches, strIniFileName)
    
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True
    ExportMTDBbyUMC = strStatus
    
End Function

' September 2004: Unused Function
''Private Function ExportMTDBbyUMCToPeakResultsTable(ByRef lngMDID As Long, Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByRef lngErrorNumber As Long) As String
'''---------------------------------------------------
'''this is simple but long procedure of exporting data
'''results to Organism MT tag database associated with gel
'''
'''We're currently writing the results to the T_Match_Making_Description table and T_FTICR_Peak_Results
'''These tables are designed to hold search results from an ion-by-ion search (either using all ions or using UMC ions only)
'''Since this form uses a UMC by UMC search, and we assign all matches for a UMC to all ions for the UMC, we'll
'''  only export the search results for the class representative ion for each UMC (typically the most abundant ion)
'''
'''Returns a status message
'''lngErrorNumber will contain the error number, if an error occurs
'''lngMDID contains the new MMD_ID value
'''---------------------------------------------------
''Const MASS_PRECISION = 6
''Const FIT_PRECISION = 3
''
''Dim mgInd As Long
''Dim lngUMCIndexOriginal As Long, lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long
''Dim ExpCnt As Long
''
''Dim strCaptionSaved As String
''Dim strExportStatus As String
''
''Dim lngPairIndex As Long
''
''Dim objP1IndFastSearch As FastSearchArrayLong
''Dim objP2IndFastSearch As FastSearchArrayLong
''Dim blnPairsPresent As Boolean
''
''Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
''Dim udtPairMatchStats() As udtPairMatchStatsType
''
'''ADO objects for stored procedure adding Match Making row
''Dim cnNew As New ADODB.Connection
''
'''ADO objects for stored procedure that adds FTICR peak rows
''Dim cmdPutNewPeak As New ADODB.Command
''Dim prmMMDID As New ADODB.Parameter
''Dim prmFTICRID As New ADODB.Parameter
''Dim prmFTICRType As New ADODB.Parameter
''Dim prmScanNumber As New ADODB.Parameter
''Dim prmChargeState As New ADODB.Parameter
''Dim prmMonoisotopicMass As New ADODB.Parameter
''Dim prmAbundance As New ADODB.Parameter
''Dim prmFit As New ADODB.Parameter
''Dim prmExpressionRatio As New ADODB.Parameter
''Dim prmLckID As New ADODB.Parameter
''Dim prmFreqShift As New ADODB.Parameter
''Dim prmMassCorrection As New ADODB.Parameter
''Dim prmMassTagID As New ADODB.Parameter
''Dim prmResType As New ADODB.Parameter
''Dim prmHitsCount As New ADODB.Parameter
''Dim prmUMCInd As New ADODB.Parameter
''Dim prmUMCFirstScan As New ADODB.Parameter
''Dim prmUMCLastScan As New ADODB.Parameter
''Dim prmUMCCount As New ADODB.Parameter
''Dim prmUMCAbundance As New ADODB.Parameter
''Dim prmUMCBestFit As New ADODB.Parameter
''Dim prmUMCAvgMW As New ADODB.Parameter
''Dim prmPairInd As New ADODB.Parameter
''
''On Error GoTo err_ExportMTDBbyUMC
''
''strCaptionSaved = Me.Caption
''
''' Connect to the database
''Me.Caption = "Connecting to the database"
''If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
''    Debug.Assert False
''    lngErrorNumber = -1
''    Me.Caption = strCaptionSaved
''    ExportMTDBbyUMCToPeakResultsTable = "Error: Unable to establish a connection to the database"
''    Exit Function
''End If
''
'''first write new analysis in T_Match_Making_Description table
''' Note that we're using CountMassTagsInUMCMatchStats() to determine the number of items in mUMCMatchStats that are MT tags
''AddEntryToMatchMakingDescriptionTable cnNew, lngMDID, ExpAnalysisSPName, CallerID, CountMassTagsInUMCMatchStats(), GelData(CallerID).CustomNETsDefined, True, strIniFileName
''AddToAnalysisHistory CallerID, "Exported UMC Identification results (single UMC mass) to Peak Results table in database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
''AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file
''
'''nothing to export
''If mMatchStatsCount <= 0 Then
''    cnNew.Close
''    Me.Caption = strCaptionSaved
''    Exit Function
''End If
''
''' Initialize the SP
''InitializeSPCommand cmdPutNewPeak, cnNew, ExpPeakSPName
''
''Set prmMMDID = cmdPutNewPeak.CreateParameter("MMDID", adInteger, adParamInput, , lngMDID)
''cmdPutNewPeak.Parameters.Append prmMMDID
''Set prmFTICRID = cmdPutNewPeak.CreateParameter("FTICRID", adVarChar, adParamInput, 50, Null)
''cmdPutNewPeak.Parameters.Append prmFTICRID
''Set prmFTICRType = cmdPutNewPeak.CreateParameter("FTICRType", adTinyInt, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFTICRType
''Set prmScanNumber = cmdPutNewPeak.CreateParameter("ScanNumber", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmScanNumber
''Set prmChargeState = cmdPutNewPeak.CreateParameter("ChargeState", adSmallInt, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmChargeState
''Set prmMonoisotopicMass = cmdPutNewPeak.CreateParameter("MonoisotopicMass", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMonoisotopicMass
''Set prmAbundance = cmdPutNewPeak.CreateParameter("Abundance", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmAbundance
''Set prmFit = cmdPutNewPeak.CreateParameter("Fit", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFit
''Set prmExpressionRatio = cmdPutNewPeak.CreateParameter("ExpressionRatio", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmExpressionRatio
''Set prmLckID = cmdPutNewPeak.CreateParameter("LckID", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmLckID
''Set prmFreqShift = cmdPutNewPeak.CreateParameter("FreqShift", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFreqShift
''Set prmMassCorrection = cmdPutNewPeak.CreateParameter("MassCorrection", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMassCorrection
''Set prmMassTagID = cmdPutNewPeak.CreateParameter("MassTagID", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMassTagID
''Set prmResType = cmdPutNewPeak.CreateParameter("Type", adInteger, adParamInput, , FPR_Type_Standard)
''cmdPutNewPeak.Parameters.Append prmResType
''Set prmHitsCount = cmdPutNewPeak.CreateParameter("HitCount", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmHitsCount
''Set prmUMCInd = cmdPutNewPeak.CreateParameter("UMCInd", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCInd
''Set prmUMCFirstScan = cmdPutNewPeak.CreateParameter("UMCFirstScan", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCFirstScan
''Set prmUMCLastScan = cmdPutNewPeak.CreateParameter("UMCLastScan", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCLastScan
''Set prmUMCCount = cmdPutNewPeak.CreateParameter("UMCCount", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCCount
''Set prmUMCAbundance = cmdPutNewPeak.CreateParameter("UMCAbundance", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCAbundance
''Set prmUMCBestFit = cmdPutNewPeak.CreateParameter("UMCBestFit", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCBestFit
''Set prmUMCAvgMW = cmdPutNewPeak.CreateParameter("UMCAvgMW", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCAvgMW
''Set prmPairInd = cmdPutNewPeak.CreateParameter("PairInd", adInteger, adParamInput, , -1)
''cmdPutNewPeak.Parameters.Append prmPairInd
''
''' Initialize the PairIndex lookup objects
''blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)
''
''Me.Caption = "Exporting peaks to DB: 0 / " & Trim(mMatchStatsCount)
''
'''now export data
''ExpCnt = 0
''With GelData(CallerID)
''    ' Step through the UMC hits and export information on each hit
''    ' Since the target table is an ion-based table, will use the index and info of the class representative ion
''    For mgInd = 0 To mMatchStatsCount - 1
''        If mgInd Mod 25 = 0 Then
''            Me.Caption = "Exporting peaks to DB: " & Trim(mgInd) & " / " & Trim(mMatchStatsCount)
''            DoEvents
''        End If
''
''        If Not mUMCMatchStats(mgInd).IDIsInternalStd Then
''            ' Only export to T_FTICR_Peak_Results if this is a MT tag hit
''            lngUMCIndexOriginal = mUMCMatchStats(mgInd).UMCIndex
''
''            With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
''                prmFTICRID.value = .ClassRepInd
''                prmFTICRType.value = .ClassRepType
''
''                Select Case .ClassRepType
''                Case glCSType
''                    prmScanNumber.value = GelData(CallerID).CSData(.ClassRepInd).ScanNumber
''                    prmChargeState.value = GelData(CallerID).CSData(.ClassRepInd).Charge
''                    prmMonoisotopicMass.value = Round(.ClassMW, MASS_PRECISION)                 ' Mass of the class rep would be: GelData(CallerID).CSData(.ClassRepInd).AverageMW
''                    prmAbundance.value = .ClassAbundance                                        ' Abundance of the class rep would be: GelData(CallerID).CSData(.ClassRepInd).Abundance
''                    prmFit.value = GelData(CallerID).CSData(.ClassRepInd).MassStDev     'standard deviation
''                    If GelLM(CallerID).CSCnt > 0 Then
''                      prmLckID.value = GelLM(CallerID).CSLckID(.ClassRepInd)
''                      prmFreqShift.value = GelLM(CallerID).CSFreqShift(.ClassRepInd)
''                      prmMassCorrection.value = GelLM(CallerID).CSMassCorrection(.ClassRepInd)
''                    End If
''                Case glIsoType
''                    prmScanNumber.value = GelData(CallerID).IsoData(.ClassRepInd).ScanNumber
''                    prmChargeState.value = GelData(CallerID).IsoData(.ClassRepInd).Charge
''                    prmMonoisotopicMass.value = Round(.ClassMW, MASS_PRECISION)                 ' Mass of the class rep would be: GelData(CallerID).IsoData(.ClassRepInd).MonoisotopicMW
''                    prmAbundance.value = .ClassAbundance                                        ' Abundance of the class rep would be: GelData(CallerID).IsoData(.ClassRepInd).Abundance
''                    prmFit.value = GelData(CallerID).IsoData(.ClassRepInd).Fit
''                    If GelLM(CallerID).IsoCnt > 0 Then
''                      prmLckID.value = GelLM(CallerID).IsoLckID(.ClassRepInd)
''                      prmFreqShift.value = GelLM(CallerID).IsoFreqShift(.ClassRepInd)
''                      prmMassCorrection.value = GelLM(CallerID).IsoMassCorrection(.ClassRepInd)
''                    End If
''                End Select
''
''                ' Note: The multi-hit count value for the UMC is the same as that for the class representative, and can thus be placed in prmHitsCount
''                prmHitsCount.value = mUMCMatchStats(mgInd).MultiAMTHitCount
''                prmUMCInd.value = mUMCMatchStats(mgInd).UMCIndex
''                prmUMCFirstScan.value = .MinScan
''                prmUMCLastScan.value = .MaxScan
''                prmUMCCount.value = .ClassCount
''                prmUMCAbundance.value = .ClassAbundance
''                prmUMCBestFit.value = Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), FIT_PRECISION)
''                prmUMCAvgMW.value = Round(.ClassMW, MASS_PRECISION)              ' This is usually the median mass of the class, not the average mass
''
''            End With
''
''            lngMassTagIndexPointer = mMTInd(mUMCMatchStats(mgInd).IDIndex)
''            lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
''
''            prmMassTagID.value = AMTData(lngMassTagIndexOriginal).ID
''
''            lngPairIndex = -1
''            lngPairMatchCount = 0
''            ReDim udtPairMatchStats(0)
''            InitializePairMatchStats udtPairMatchStats(0)
''            If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
''                lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, objP1IndFastSearch, objP2IndFastSearch, False, (LastSearchTypeN14N15 = SEARCH_N15), lngPairMatchCount, udtPairMatchStats())
''            End If
''
''            ' If pairs exist, then we need to output an entry for each pair that this UMC is a member of
''            If lngPairMatchCount > 0 Then
''                For lngPairMatchIndex = 0 To lngPairMatchCount - 1
''                    With udtPairMatchStats(lngPairMatchIndex)
''                        prmPairInd.value = .PairIndex
''                        prmExpressionRatio.value = .ExpressionRatio
''
''                        cmdPutNewPeak.Execute
''                        ExpCnt = ExpCnt + 1
''                    End With
''                Next lngPairMatchIndex
''            Else
''                With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
''                    prmExpressionRatio.value = LookupExpressionRatioValue(CallerID, .ClassRepInd, (.ClassRepType = glIsoType))
''                    prmPairInd.value = -1
''
''                    cmdPutNewPeak.Execute
''                    ExpCnt = ExpCnt + 1
''                End With
''            End If
''        End If
''    Next mgInd
''End With
''
''' MonroeMod
''AddToAnalysisHistory CallerID, "Export to Peak Results table details: UMC Peaks Match Count = " & ExpCnt
''
''Me.Caption = strCaptionSaved
''
''strExportStatus = ExpCnt & " associations between MT tags and UMC's exported to peak results table."
''Set cmdPutNewPeak.ActiveConnection = Nothing
''cnNew.Close
''
''If blnUpdateGANETForAnalysisInDB Then
''    ' Export the the GANET Slope, Intercept, and Fit to the database
''    With GelAnalysis(CallerID)
''        strExportStatus = strExportStatus & vbCrLf & ExportGANETtoMTDB(CallerID, .GANET_Slope, .GANET_Intercept, .GANET_Fit)
''    End With
''End If
''
''Set objP1IndFastSearch = Nothing
''Set objP2IndFastSearch = Nothing
''
''ExportMTDBbyUMCToPeakResultsTable = strExportStatus
''lngErrorNumber = 0
''Exit Function
''
''err_ExportMTDBbyUMC:
''ExportMTDBbyUMCToPeakResultsTable = "Error: " & Err.Number & vbCrLf & Err.Description
''lngErrorNumber = Err.Number
''On Error Resume Next
''If Not cnNew Is Nothing Then cnNew.Close
''Me.Caption = strCaptionSaved
''Set objP1IndFastSearch = Nothing
''Set objP2IndFastSearch = Nothing
''
''End Function

Private Function ExportMTDBbyUMCToUMCResultsTable(ByRef lngMDID As Long, Optional blnCreateNewEntryInMMDTable As Boolean = False, Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByVal blnExportUMCMembers As Boolean = False, Optional ByRef lngErrorNumber As Long, Optional ByVal blnAddQuantitationDescriptionEntry As Boolean = True, Optional ByVal blnExportUMCsWithNoMatches As Boolean = True, Optional ByVal strIniFileName As String = "") As String
'---------------------------------------------------
'This function will export data to the T_FTICR_UMC_Results table, T_FTICR_UMC_ResultDetails table,
'  and T_FTICR_UMC_InternalStandardDetails table
'
'It will create a new entry in T_Match_Making_Description if blnCreateNewEntryInMMDTable = True
'If blnAddQuantitationDescriptionEntry = True, then calls ExportMTDBAddQuantitationDescriptionEntry
'  to create a new entry in T_Quantitation_Description and T_Quantitation_MDIDs
'
'Returns a status message
'lngErrorNumber will contain the error number, if an error occurs
'---------------------------------------------------
Dim lngPointer As Long, lngUMCIndex As Long
Dim lngUMCIndexCompare As Long
Dim lngUMCIndexOriginal As Long
Dim lngUMCIndexOriginalLastStored As Long

Dim lngUMCIndexOriginalPairOther As Long
Dim lngPeakFPRType As Long
Dim lngPeakFPRTypeLight As Long, lngPeakFPRTypeHeavy As Long

Dim lngPairIndex As Long

Dim objP1IndFastSearch As FastSearchArrayLong
Dim objP2IndFastSearch As FastSearchArrayLong
Dim blnPairsPresent As Boolean

Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
Dim udtPairMatchStats() As udtPairMatchStatsType
Dim lngUMCResultsIDReturn() As Long

Dim blnContinueCompare As Boolean

Dim lngInternalStdMatchCount As Long
Dim MassTagExpCnt As Long
Dim InternalStdExpCnt As Long
Dim strCaptionSaved As String
Dim strExportStatus As String

'ADO objects for stored procedure adding Match Making row
Dim cnNew As New ADODB.Connection

Dim sngDBSchemaVersion As Single

'ADO objects for stored procedure that adds FTICR UMC rows
Dim cmdPutNewUMC As New ADODB.Command
Dim udtPutUMCParams As udtPutUMCParamsListType
    
'ADO objects for stored procedure that adds FTICR UMC member rows
Dim cmdPutNewUMCMember As New ADODB.Command
Dim udtPutUMCMemberParams As udtPutUMCMemberParamsListType
    
'ADO objects for stored procedure adding FTICR UMC Details
Dim cmdPutNewUMCMatch As New ADODB.Command
Dim udtPutUMCMatchParams As udtPutUMCMatchParamsListType

'ADO objects for stored procedure adding FTICR UMC Internal Standard Details
Dim cmPutNewUMCInternalStdMatch As New ADODB.Command
Dim udtPutUMCInternalStdMatchParams As udtPutUMCInternalStdMatchParamsListType

Dim blnUMCMatchFound() As Boolean       ' 0-based array, used to keep track of whether or not the UMC matched any MT tags or Internal Standards

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

' Lookup the DB Schema Version
sngDBSchemaVersion = LookupDBSchemaVersion(cnNew)

If sngDBSchemaVersion < 2 Then
    ' Force UMC Member export to be false
    blnExportUMCMembers = False
End If

If blnCreateNewEntryInMMDTable Then
    'first write new analysis in T_Match_Making_Description table
    lngErrorNumber = AddEntryToMatchMakingDescriptionTable(cnNew, lngMDID, ExpAnalysisSPName, CallerID, CountMassTagsInUMCMatchStats(), mSearchUsedCustomNETs, True, strIniFileName)
Else
    lngErrorNumber = 0
End If

If lngErrorNumber <> 0 Then
    Debug.Assert False
    GoTo err_Cleanup
End If

If blnCreateNewEntryInMMDTable Or mMatchStatsCount > 0 Or blnExportUMCsWithNoMatches Then
    ' MonroeMod
    AddToAnalysisHistory CallerID, "Exported UMC Identification results (single UMC mass) to UMC Results table in database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
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

' Initialize the variables for accessing the AddFTICRUmcMember SP
ExportMTDBInitializePutNewUMCMemberParams cnNew, cmdPutNewUMCMember, udtPutUMCMemberParams, ExpUMCMemberSPName

' Initialize the variables for accessing the AddFTICRUmcMatch SP
ExportMTDBInitializePutUMCMatchParams cnNew, cmdPutNewUMCMatch, udtPutUMCMatchParams, ExpUmcMatchSPName

' Initialize the variables for accessing the AddFTICRUmcInternalStdMatch SP
ExportMTDBInitializePutUMCInternalStdMatchParams cnNew, cmPutNewUMCInternalStdMatch, udtPutUMCInternalStdMatchParams, ExpUmcInternalStdMatchSPName

' Initialize the PairIndex lookup objects
blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)

Select Case LastSearchTypeN14N15
Case SEARCH_N14
     NTypeStr = MOD_TKN_N14
Case SEARCH_N15
     NTypeStr = MOD_TKN_N15
End Select

lngPeakFPRTypeLight = PairsLookupFPRType(CallerID, False)
lngPeakFPRTypeHeavy = PairsLookupFPRType(CallerID, True)

Me.Caption = "Exporting UMC's to DB: 0 / " & Trim(mMatchStatsCount)

'now export data
MassTagExpCnt = 0
InternalStdExpCnt = 0

    ' Step through the UMC hits and export information on each hit
    ' mUMCMatchStats() will contain multiple entries for each UMC if the UMC matched multiple MT tags
    ' Additionally, if the UMC matched an Internal Standard, then that will also be included in mUMCMatchStats()
    ' However, we only want to write one entry for each UMC to T_FTICR_UMC_Results
    ' Thus, we need to keep track of whether or not an entry has been made to T_FTICR_UMC_Results
    ' Luckily, results are stored to mUMCMatchStats() in order of UMC Index
    
    ' We need to keep track of which UMC's are exported to the results table
    ReDim blnUMCMatchFound(GelUMC(CallerID).UMCCnt)
    
    lngUMCIndexOriginalLastStored = -1
    
    For lngPointer = 0 To mMatchStatsCount - 1
        If lngPointer Mod 25 = 0 Then
            Me.Caption = "Exporting UMC's to DB: " & Trim(lngPointer) & " / " & Trim(mMatchStatsCount)
            DoEvents
            If mKeyPressAbortProcess = 2 Then Exit For
        End If
        
        lngUMCIndexOriginal = mUMCMatchStats(lngPointer).UMCIndex
        If lngUMCIndexOriginal <> lngUMCIndexOriginalLastStored Then
            ' Add a new row to T_FTICR_UMC_Results
            ' Note: If we searched only paired UMC's, then record both members of the pairs and set lngPeakFPRType to FPR_Type_N14_N15_L
            '       Additionally, record the pair index in the database and record the opposite pair member
            
            ' Need to perform a look-ahead to determine the number of Internal Standard matches for this UMC Index
            lngInternalStdMatchCount = 0
            lngUMCIndexCompare = lngPointer
            blnContinueCompare = True
            Do
                If mUMCMatchStats(lngUMCIndexCompare).IDIsInternalStd Then
                    lngInternalStdMatchCount = lngInternalStdMatchCount + 1
                End If
                lngUMCIndexCompare = lngUMCIndexCompare + 1
                If lngUMCIndexCompare < mMatchStatsCount Then
                    blnContinueCompare = (mUMCMatchStats(lngUMCIndexCompare).UMCIndex = lngUMCIndexOriginal)
                Else
                    blnContinueCompare = False
                End If
            Loop While blnContinueCompare
            
            lngPairIndex = -1
            lngPairMatchCount = 0
            ReDim udtPairMatchStats(0)
            InitializePairMatchStats udtPairMatchStats(0)
            If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
                lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, objP1IndFastSearch, objP2IndFastSearch, False, (LastSearchTypeN14N15 = SEARCH_N15), lngPairMatchCount, udtPairMatchStats())
            End If
          
            ' If pairs exist, then we need to output an entry for each pair that this UMC is a member of
            If lngPairMatchCount > 0 Then
                ReDim lngUMCResultsIDReturn(lngPairMatchCount - 1)
                
                For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                    ' Lookup whether this UMC is the light or heavy member in the pair
                    With GelP_D_L(CallerID).Pairs(udtPairMatchStats(lngPairMatchIndex).PairIndex)
                        If .P1 = lngUMCIndexOriginal Then
                            lngPeakFPRType = lngPeakFPRTypeLight
                        Else
                            lngPeakFPRType = lngPeakFPRTypeHeavy
                        End If
                    End With
                    
                    ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, blnExportUMCMembers, CallerID, lngUMCIndexOriginal, mUMCMatchStats(lngPointer).MultiAMTHitCount, ClsStat(), udtPairMatchStats(lngPairMatchIndex), lngPeakFPRType, lngInternalStdMatchCount
                    blnUMCMatchFound(lngUMCIndexOriginal) = True
        
                    ' Populate array with return value
                    lngUMCResultsIDReturn(lngPairMatchIndex) = FixNullLng(udtPutUMCParams.UMCResultsIDReturn.Value)
            
                    ' Add the other member of the pair too (typically the heavy member)
                    ' Need to determine the UMC index for the other member of the pair
                    With GelP_D_L(CallerID).Pairs(udtPairMatchStats(lngPairMatchIndex).PairIndex)
                        If .P1 = lngUMCIndexOriginal Then
                            ' Already saved the light member, now save the heavy member
                            lngUMCIndexOriginalPairOther = .P2
                            lngPeakFPRType = lngPeakFPRTypeHeavy
                        Else
                            ' Already saved the heavy member, now save the light member
                            lngUMCIndexOriginalPairOther = .P1
                            lngPeakFPRType = lngPeakFPRTypeLight
                        End If
                        
                        ' Always export the other member of the pair, even if it has already been exported
                        ' Note that we do not record any MT tag hits for the other member of the pair
                        ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, blnExportUMCMembers, CallerID, lngUMCIndexOriginalPairOther, 0, ClsStat(), udtPairMatchStats(lngPairMatchIndex), lngPeakFPRType, 0
                        blnUMCMatchFound(lngUMCIndexOriginalPairOther) = True
                        
                    End With
                    
                Next lngPairMatchIndex
            Else
                lngPeakFPRType = FPR_Type_Standard
            
                ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, blnExportUMCMembers, CallerID, lngUMCIndexOriginal, mUMCMatchStats(lngPointer).MultiAMTHitCount, ClsStat(), udtPairMatchStats(0), lngPeakFPRType, lngInternalStdMatchCount
                blnUMCMatchFound(lngUMCIndexOriginal) = True
        
                udtPutUMCMatchParams.UMCResultsID.Value = FixNullLng(udtPutUMCParams.UMCResultsIDReturn.Value)
                udtPutUMCInternalStdMatchParams.UMCResultsID.Value = udtPutUMCMatchParams.UMCResultsID.Value
                
            End If
        End If
        
        ' Now write the MT tag match
        If lngPairMatchCount > 0 Then
            For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                udtPutUMCMatchParams.UMCResultsID.Value = lngUMCResultsIDReturn(lngPairMatchIndex)
                udtPutUMCInternalStdMatchParams.UMCResultsID.Value = lngUMCResultsIDReturn(lngPairMatchIndex)
                
                ExportMTDBbyUMCToUMCResultDetailsTable lngPointer, udtPutUMCInternalStdMatchParams, cmPutNewUMCInternalStdMatch, udtPutUMCMatchParams, cmdPutNewUMCMatch
            Next lngPairMatchIndex
        Else
            ExportMTDBbyUMCToUMCResultDetailsTable lngPointer, udtPutUMCInternalStdMatchParams, cmPutNewUMCInternalStdMatch, udtPutUMCMatchParams, cmdPutNewUMCMatch
        End If
            
        If mUMCMatchStats(lngPointer).IDIsInternalStd Then
            ' Increment this if we export an Internal Standard
            InternalStdExpCnt = InternalStdExpCnt + 1
        Else
            ' Increment this if we export a MT tag
            MassTagExpCnt = MassTagExpCnt + 1
        End If
        
        ' Update lngUMCIndexOriginalLastStored
        lngUMCIndexOriginalLastStored = lngUMCIndexOriginal
        
    Next lngPointer

    If blnExportUMCsWithNoMatches And mKeyPressAbortProcess < 2 Then
        ' Also export the UMC's that do not have any hits
        ' If SearchType = SEARCH_PAIRED or SEARCH_NON_PAIRED then only export paired or unpaired UMC's without matches
        
        With GelUMC(CallerID)
            For lngUMCIndex = 0 To .UMCCnt - 1
                If lngUMCIndex Mod 25 = 0 Then
                    Me.Caption = "Exporting non-matching UMC's: " & Trim(lngUMCIndex) & " / " & Trim(.UMCCnt)
                    DoEvents
                    If mKeyPressAbortProcess = 2 Then Exit For
                End If
                
                If Not blnUMCMatchFound(lngUMCIndex) Then
                    ' No match was found
                    If SearchType = SEARCH_ALL Or _
                       SearchType = SEARCH_PAIRED_PLUS_NON_PAIRED Or _
                      (SearchType = SEARCH_PAIRED And eClsPaired(lngUMCIndex) <> umcpNone) Or _
                      (SearchType = SEARCH_NON_PAIRED And eClsPaired(lngUMCIndex) = umcpNone) Then
                    
                        ' Export to the database
                        lngPairIndex = -1
                        lngPairMatchCount = 0
                        ReDim udtPairMatchStats(0)
                        InitializePairMatchStats udtPairMatchStats(0)
                        If eClsPaired(lngUMCIndex) <> umcpNone And blnPairsPresent Then
                            lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndex, objP1IndFastSearch, objP2IndFastSearch, False, (LastSearchTypeN14N15 = SEARCH_N15), lngPairMatchCount, udtPairMatchStats())
                        End If
                            
                        ' If pairs exist, then we need to output an entry for each pair that this UMC is a member of
                        If lngPairMatchCount > 0 Then
                            For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                                ' Lookup whether this UMC is the light or heavy member in the pair
                                With GelP_D_L(CallerID).Pairs(udtPairMatchStats(lngPairMatchIndex).PairIndex)
                                    If .P1 = lngUMCIndex Then
                                        lngPeakFPRType = lngPeakFPRTypeLight
                                    Else
                                        lngPeakFPRType = lngPeakFPRTypeHeavy
                                    End If
                                End With
                                        
                                ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, blnExportUMCMembers, CallerID, lngUMCIndex, 0, ClsStat(), udtPairMatchStats(lngPairMatchIndex), lngPeakFPRType, 0
                            Next lngPairMatchIndex
                        Else
                            lngPeakFPRType = FPR_Type_Standard
                        
                            ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, blnExportUMCMembers, CallerID, lngUMCIndex, 0, ClsStat(), udtPairMatchStats(0), lngPeakFPRType, 0
                        End If
                            
                    End If
                End If
            Next lngUMCIndex
        End With
    End If

' MonroeMod
AddToAnalysisHistory CallerID, "Export to UMC Results table details: MT tags Match Count = " & MassTagExpCnt & "; Internal Std Match Count = " & InternalStdExpCnt

Me.Caption = strCaptionSaved

strExportStatus = MassTagExpCnt & " associations between MT tags and UMC's exported (" & Trim(InternalStdExpCnt) & " Internal Standards)."
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
    If lngErrorNumber = 0 And lngMDID >= 0 And (MassTagExpCnt > 0 Or InternalStdExpCnt > 0) Then
        ExportMTDBAddQuantitationDescriptionEntry Me, CallerID, ExpQuantitationDescription, lngMDID, lngErrorNumber, strIniFileName, 1, 1, 1, Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
    End If
End If

ExportMTDBbyUMCToUMCResultsTable = strExportStatus
lngErrorNumber = 0
Set objP1IndFastSearch = Nothing
Set objP2IndFastSearch = Nothing

Exit Function

err_ExportMTDBbyUMC:
Debug.Assert False
LogErrors Err.Number, "ExportMTDBbyUMCToUMCResultsTable"
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "Error exporting matches to the UMC results table: " & Err.Description, vbExclamation + vbOKOnly, glFGTU
End If

err_Cleanup:
On Error Resume Next
If Not cnNew Is Nothing Then cnNew.Close
Me.Caption = strCaptionSaved
Set objP1IndFastSearch = Nothing
Set objP2IndFastSearch = Nothing

If Err.Number <> 0 Then lngErrorNumber = Err.Number
ExportMTDBbyUMCToUMCResultsTable = "Error: " & lngErrorNumber & vbCrLf & Err.Description

End Function

Private Function ExportMTDBbyUMCToUMCResultDetailsTable(lngPointer As Long, ByRef udtPutUMCInternalStdMatchParams As udtPutUMCInternalStdMatchParamsListType, ByRef cmPutNewUMCInternalStdMatch As ADODB.Command, ByRef udtPutUMCMatchParams As udtPutUMCMatchParamsListType, cmdPutNewUMCMatch As ADODB.Command)

    Dim lngInternalStdIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long

    Dim strMassMods As String

    If mUMCMatchStats(lngPointer).IDIsInternalStd Then
    
        ' Write an entry to T_FTICR_UMC_InternalStdDetails
        lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(lngPointer).IDIndex)

        udtPutUMCInternalStdMatchParams.SeqID.Value = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).SeqID
        udtPutUMCInternalStdMatchParams.MatchingMemberCount.Value = mUMCMatchStats(lngPointer).MemberHitCount
        udtPutUMCInternalStdMatchParams.MatchScore.Value = mUMCMatchStats(lngPointer).SLiCScore
        udtPutUMCInternalStdMatchParams.ExpectedNET.Value = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).NET
        udtPutUMCInternalStdMatchParams.DelMatchScore.Value = mUMCMatchStats(lngPointer).DelSLiC
        
        cmPutNewUMCInternalStdMatch.Execute
        
    Else
        ' Write an entry to T_FTICR_UMC_ResultDetails
        
        lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngPointer).IDIndex)
        lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
    
        udtPutUMCMatchParams.MassTagID.Value = AMTData(lngMassTagIndexOriginal).ID
        udtPutUMCMatchParams.MatchingMemberCount.Value = mUMCMatchStats(lngPointer).MemberHitCount
        udtPutUMCMatchParams.MatchScore.Value = mUMCMatchStats(lngPointer).SLiCScore
        
        strMassMods = NTypeStr
        If Len(mMTMods(lngMassTagIndexPointer)) > 0 Then
            strMassMods = strMassMods & " " & Trim(mMTMods(lngMassTagIndexPointer))
            If NTypeStr = MOD_TKN_N14 Then
                udtPutUMCMatchParams.MassTagModMass.Value = mMTMWN14(mUMCMatchStats(lngPointer).IDIndex) - AMTData(lngMassTagIndexOriginal).MW
            Else
                udtPutUMCMatchParams.MassTagModMass.Value = mMTMWN15(mUMCMatchStats(lngPointer).IDIndex) - AMTData(lngMassTagIndexOriginal).MW
            End If
        Else
            If NTypeStr = MOD_TKN_N14 Then
                udtPutUMCMatchParams.MassTagModMass.Value = 0
            Else
                udtPutUMCMatchParams.MassTagModMass.Value = glN14N15_DELTA * AMTData(lngMassTagIndexOriginal).CNT_N
            End If
        End If
        
        If Len(strMassMods) > PUT_UMC_MATCH_MAX_MODSTRING_LENGTH Then strMassMods = Left(strMassMods, PUT_UMC_MATCH_MAX_MODSTRING_LENGTH)
        udtPutUMCMatchParams.MassTagMods.Value = strMassMods
        
        udtPutUMCMatchParams.DelMatchScore.Value = mUMCMatchStats(lngPointer).DelSLiC
        
        cmdPutNewUMCMatch.Execute
    
    End If

End Function

Private Function GetTokenValue(ByVal S As String, ByVal t As String) As Long
'---------------------------------------------------------------------------
'returns value next to token T in string of type Token1/Value1 Token2/Value2
'-1 if not found or on any error
'---------------------------------------------------------------------------
Dim SSplit() As String
Dim MSplit() As String
Dim i As Long
On Error GoTo exit_GetTokenValue
GetTokenValue = -1
SSplit = Split(S, " ")
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
        'loading protein names
        UpdateStatus "Loading ORF info"
        
        If Not GelAnalysis(CallerID) Is Nothing Then
            If Len(GelAnalysis(CallerID).MTDB.cn.ConnectionString) > 0 And Not APP_BUILD_DISABLE_MTS Then
                Me.MousePointer = vbHourglass
                .DBConnectionString = GelAnalysis(CallerID).MTDB.cn.ConnectionString
                .RetrieveSQL = glbPreferencesExpanded.MTSConnectionInfo.sqlGetMTNames
                If .FillData(Me) Then
                   If .DataStatus = dsLoaded Then
                        blnSuccess = True
                    End If
                End If
                Me.MousePointer = vbDefault
            End If
        End If
    End With
    
    InitializeORFInfo = blnSuccess
End Function

Public Sub InitializeSearch()
'------------------------------------------------------------------------------------
'load MT tag database data if neccessary
'if CallerID is associated with MT tag database load that db if not already loaded
'if CallerID is not associated with MT tag database load legacy database
'------------------------------------------------------------------------------------
Dim eResponse As VbMsgBoxResult

On Error Resume Next
Me.MousePointer = vbHourglass
If bLoading Then
    ' Update lblUMCMassMode to reflect the mass mode used to identify the UMC's
    Select Case GelUMC(CallerID).def.ClassMW
    Case UMCClassMassConstants.UMCMassAvg
        lblUMCMassMode = "UMC Mass = Average of the masses of the UMC members"
    Case UMCClassMassConstants.UMCMassRep
        lblUMCMassMode = "UMC Mass = Mass of the UMC Class Representative"
    Case UMCClassMassConstants.UMCMassMed
        lblUMCMassMode = "UMC Mass = Median of the masses of the UMC members"
    Case UMCMassAvgTopX
        lblUMCMassMode = "UMC Mass = Average of top X members of the UMC"
    Case UMCMassMedTopX
        lblUMCMassMode = "UMC Mass = Median of top X members of the UMC"
    Case Else
        lblUMCMassMode = "UMC Mass = ?? Unable to determine; is it a new mass mode?"
    End Select
    
    If GelAnalysis(CallerID) Is Nothing Then
        If AMTCnt > 0 Then    'something is loaded
          If (Len(CurrMTDatabase) > 0 Or Len(CurrLegacyMTDatabase) > 0) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                'MT tag data; we dont know is it appropriate; warn user
                WarnUserUnknownMassTags CallerID
            End If
            lblMTStatus.Caption = ConstructMTStatusText(True)
          
            ' Initialize the MT search object
            If Not CreateNewMTSearchObject() Then
                lblMTStatus.Caption = "Error creating search object."
            Else
               ' Error initializing MT search object
            End If
       
       Else                  'nothing is loaded
            If Len(GelData(CallerID).PathtoDatabase) > 0 And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                If APP_BUILD_DISABLE_MTS Then
                    eResponse = vbYes
                Else
                    eResponse = MsgBox("Current display is not associated with any MT tag database.  Do you want to load the MT tags from the defined legacy MT tag database?" & vbCrLf & GelData(CallerID).PathtoDatabase, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load Legacy MT tags")
                End If
            Else
                eResponse = vbNo
            End If
            
            If eResponse = vbYes Then
                LoadLegacyMassTags
            Else
                WarnUserNotConnectedToDB CallerID, True
                lblMTStatus.Caption = "No MT tags loaded"
            End If
        End If
    Else         'have to have MT tag database loaded
        Call LoadMTDB
    End If
    UpdateStatus "Generating UMC statistics ..."
    ClsCnt = UMCStatistics1(CallerID, ClsStat())
    UpdateStatus "Pairs Count: " & GelP_D_L(CallerID).PCnt
    
    chkDisableCustomNETs.Enabled = GelData(CallerID).CustomNETsDefined
    If APP_BUILD_DISABLE_ADVANCED Then
        chkDisableCustomNETs.Visible = chkDisableCustomNETs.Enabled
    End If
    
    EnableDisableNETFormulaControls
    
    SetETMode etGANET
   
    UpdateStatus "UMCs pairing status ..."
    UpdateUMCsPairingStatusNow
    UpdateStatus "Ready"
    
    'memorize number of scans (to be used with elution)
    MinFN = GelData(CallerID).ScanInfo(1).ScanNumber
    MaxFN = GelData(CallerID).ScanInfo(UBound(GelData(CallerID).ScanInfo)).ScanNumber
    bLoading = False
End If
Me.MousePointer = vbDefault
End Sub

Private Function IsValidMatch(dblCurrMW As Double, AbsMWErr As Double, CurrScan As Long, dblMatchNET As Double, dblMatchMass As Double) As Boolean
    ' Checks if dblCurrMW is within tolerance of the given MT tag
    ' Also checks if the NET equivalent of CurrScan is within tolerance of the NET value for the given MT tag
    ' Returns True if both are within tolerance, false otherwise
    
    Dim InvalidMatch As Boolean
    
    ' If dblCurrMW is not within AbsMWErr of dblMatchMass then this match is inherited
    If Abs(dblCurrMW - dblMatchMass) > AbsMWErr Then
        InvalidMatch = True
    Else
        ' If CurrScan is not within .NETTol of dblMatchNET then this match is inherited
        If samtDef.NETTol >= 0 Then
            If Abs(ConvertScanToNET(CurrScan) - dblMatchNET) > samtDef.NETTol Then
                InvalidMatch = True
            End If
        End If
    End If
    
    IsValidMatch = Not InvalidMatch
End Function

Private Sub LoadLegacyMassTags()

    '------------------------------------------------------------
    'load/reload MT tags
    '------------------------------------------------------------
    Dim eResponse As VbMsgBoxResult
    On Error Resume Next
    'ask user if it wants to replace legitimate MT tag DB with legacy DB
    If Not GelAnalysis(CallerID) Is Nothing And Not APP_BUILD_DISABLE_MTS Then
       eResponse = MsgBox("Current display is associated with MT tag database." & vbCrLf _
                    & "Are you sure you want to use a legacy database for search?", vbYesNoCancel, glFGTU)
       If eResponse <> vbYes Then Exit Sub
    End If
    Me.MousePointer = vbHourglass
    If Len(GelData(CallerID).PathtoDatabase) > 0 Then
       If ConnectToLegacyAMTDB(Me, CallerID, False, True, False) Then
          If CreateNewMTSearchObject() Then
             lblMTStatus.Caption = "Loaded; MT tag count: " & LongToStringWithCommas(AMTCnt)
          Else
             lblMTStatus.Caption = "Error creating search object."
          End If
       Else
          lblMTStatus.Caption = "Error loading MT tags."
       End If
    Else
        WarnUserInvalidLegacyDBPath
    End If
    Me.MousePointer = vbDefault

End Sub

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean
    Dim strMessage As String
    
    Static blnWorking As Boolean
    
    If blnWorking Then Exit Sub
    blnWorking = True
    
    cmdSearchAllUMCs.Enabled = False
    
    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, 0, blnForceReload, True, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTStatus.Caption = ConstructMTStatusText(True)
    
        If Not CreateNewMTSearchObject() Then
           lblMTStatus.Caption = "Error creating search object."
        End If
    Else
        If blnDBConnectionError Then
            strMessage = "Error loading MT tags: database connection error."
        Else
            If Not GelAnalysis(CallerID) Is Nothing Then
                If Len(GelAnalysis(CallerID).MTDB.cn.ConnectionString) > 0 And Not APP_BUILD_DISABLE_MTS Then
                    strMessage = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
                ElseIf Len(GelData(CallerID).PathtoDatabase) > 0 Then
                    strMessage = "Error loading MT tags from Legacy DB"
                Else
                    strMessage = "Error loading MT tags: MT tag database not defined"
                End If
            Else
                strMessage = "Error loading MT tags: MT tag database not defined"
            End If
        End If
    
        lblMTStatus.Caption = strMessage
    End If
    
    cmdSearchAllUMCs.Enabled = True
    blnWorking = False
    
End Sub

Private Function ManageCurrID(ByVal ManageType As Long) As Boolean
On Error GoTo exit_ManageCurrID
Select Case ManageType
Case MNG_ERASE
     mCurrIDCnt = 0
     Erase mCurrIDMatches
Case MNG_TRIM
     If mCurrIDCnt > 0 Then
        ReDim Preserve mCurrIDMatches(mCurrIDCnt - 1)
     Else
        ManageCurrID = ManageCurrID(MNG_ERASE)
     End If
Case MNG_RESET
     mCurrIDCnt = 0
     ReDim mCurrIDMatches(MNG_START_SIZE)
Case MNG_ADD_START_SIZE
     ReDim Preserve mCurrIDMatches(mCurrIDCnt + MNG_START_SIZE)
Case Else
     If ManageType > 0 Then
        ReDim Preserve mCurrIDMatches(mCurrIDCnt + ManageType)
     End If
End Select
ManageCurrID = True
exit_ManageCurrID:
End Function

Private Sub PickParameters()
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
Call txtAlkylationMWCorrection_LostFocus
Call txtNETFormula_LostFocus
End Sub

Private Function PrepareSearchInternalStandards() As Boolean
Dim intIndex As Integer
Dim dblInternalStdMasses() As Double
Dim qsd As New QSDouble
Dim blnSuccess As Boolean

On Error GoTo PrepareSearchInternalStandardsErrorHandler

blnSuccess = False
With UMCInternalStandards
    If .Count > 0 Then
        UpdateStatus "Preparing fast Internal Standard search..."
        ReDim dblInternalStdMasses(.Count - 1)
        ReDim mInternalStdIndexPointers(.Count - 1)
        
        For intIndex = 0 To .Count - 1
            dblInternalStdMasses(intIndex) = .InternalStandards(intIndex).MonoisotopicMass
            mInternalStdIndexPointers(intIndex) = intIndex
        Next intIndex
   
        If qsd.QSAsc(dblInternalStdMasses, mInternalStdIndexPointers) Then
            Set InternalStdFastSearch = New MWUtil
            If InternalStdFastSearch.Fill(dblInternalStdMasses()) Then
                blnSuccess = True
            End If
        End If
    Else
        ReDim mInternalStdIndexPointers(0)
        blnSuccess = True
    End If
End With

PrepareSearchInternalStandards = blnSuccess
Exit Function

PrepareSearchInternalStandardsErrorHandler:
Debug.Assert False
LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.PrepareSearchInternalStandards"
PrepareSearchInternalStandards = False

End Function

Private Function PrepareSearchN14() As Boolean
'---------------------------------------------------------------
'prepare search of N14 peptide (use loaded peptides masses)
'---------------------------------------------------------------
On Error Resume Next
If mMTCnt > 0 Then
   UpdateStatus "Preparing fast N14 search..."
   ' Dim qsd As New QSDouble
   ' Old: If qsd.QSAsc(mMTMWN14(), mMTInd()) Then
   If ShellSortDoubleWithParallelLong(mMTMWN14(), mMTInd(), 0, UBound(mMTMWN14)) Then
      Set MWFastSearch = New MWUtil
      If MWFastSearch.Fill(mMTMWN14()) Then PrepareSearchN14 = True
   End If
End If
End Function

Private Function PrepareSearchN15() As Boolean
'---------------------------------------------------------------
'prepare search of N15 peptide (use number of N to correct mass)
'---------------------------------------------------------------
On Error Resume Next
If mMTCnt > 0 Then
   UpdateStatus "Preparing fast N15 search..."
   ' Dim qsd As New QSDouble
   ' Old: If qsd.QSAsc(mMTMWN15(), mMTInd()) Then
   If ShellSortDoubleWithParallelLong(mMTMWN15(), mMTInd(), 0, UBound(mMTMWN15)) Then
      Set MWFastSearch = New MWUtil
      If MWFastSearch.Fill(mMTMWN15()) Then PrepareSearchN15 = True
   End If
End If
End Function

Private Sub PopulateComboBoxes()
    Dim intIndex As Integer
    
On Error GoTo PopulateComboBoxesErrorHandler

    With cboResidueToModify
        .Clear
        .AddItem "Full MT"
        For intIndex = 0 To 25
            .AddItem Chr(vbKeyA + intIndex)
        Next intIndex
        .AddItem glPHOSPHORYLATION
        .ListIndex = 0
    End With
    
    With cboInternalStdSearchMode
        .Clear
        .AddItem "Search only MT tags", issmFindOnlyMassTags
        .AddItem "Search MT tags & Int Stds", issmFindWithMassTags
        .AddItem "Search only Internal Stds", issmFindOnlyInternalStandards
        
        If APP_BUILD_DISABLE_MTS Then
            .ListIndex = issmFindOnlyMassTags
        Else
            .ListIndex = issmFindWithMassTags
        End If
    End With
    
    With cboAMTSearchResultsBehavior
        .Clear
        .AddItem "Auto remove existing results prior to search", asrbAMTSearchResultsBehaviorConstants.asrbAutoRemoveExisting
        .AddItem "Keep existing results; do not skip UMC's", asrbAMTSearchResultsBehaviorConstants.asrbKeepExisting
        .AddItem "Keep existing results; skip UMC's with results", asrbAMTSearchResultsBehaviorConstants.asrbKeepExistingAndSkip
        .ListIndex = asrbAutoRemoveExisting
    End With
    
    With cboSearchRegionShape
        .Clear
        .AddItem "Elliptical search region"
        .AddItem "Rectangular search region"
        .ListIndex = srsSearchRegionShapeConstants.srsElliptical
    End With
    
    Exit Sub
    
PopulateComboBoxesErrorHandler:
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->PopulateComboBoxes"
End Sub

Private Function PrepareMTArrays() As Boolean
'---------------------------------------------------------------
'prepares masses from loaded MT tags based on specified
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
Dim blnAddMassTag As Boolean

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
    
    .N15InsteadOfN14 = optN(SEARCH_N15).Value
    .DynamicMods = optDBSearchModType(MODS_DYNAMIC).Value
End With

If IsNumeric(txtDBSearchMinimumHighNormalizedScore.Text) Then
    mMTMinimumHighNormalizedScore = CSngSafe(txtDBSearchMinimumHighNormalizedScore.Text)
Else
    mMTMinimumHighNormalizedScore = 0
End If
    
If IsNumeric(txtDBSearchMinimumHighDiscriminantScore.Text) Then
    mMTMinimumHighDiscriminantScore = CSngSafe(txtDBSearchMinimumHighDiscriminantScore.Text)
Else
    mMTMinimumHighDiscriminantScore = 0
End If

If IsNumeric(txtDBSearchMinimumPeptideProphetProbability.Text) Then
    mMTMinimumPeptideProphetProbability = CSngSafe(txtDBSearchMinimumPeptideProphetProbability.Text)
Else
    mMTMinimumPeptideProphetProbability = 0
End If

If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
    If mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
        ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighDiscriminantScore, also taking into account HighNormalizedScore
        ValidateMTMinimumDiscriminantAndPepProphet AMTData(), 1, AMTCnt, mMTMinimumHighDiscriminantScore, mMTMinimumPeptideProphetProbability, mMTMinimumHighNormalizedScore, 2
    Else
        ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighNormalizedScore
        ValidateMTMinimimumHighNormalizedScore AMTData(), 1, AMTCnt, mMTMinimumHighNormalizedScore, 2
    End If
End If

' Record the current state of .CustomNETsDefined
' If chkDisableCustomNETs is checked, then this will have temporarily been set to False
mSearchUsedCustomNETs = GelData(CallerID).CustomNETsDefined

If AMTCnt <= 0 Then
    mMTCnt = 0
Else
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
        If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
            If AMTData(i).HighNormalizedScore >= mMTMinimumHighNormalizedScore And _
               AMTData(i).HighDiscriminantScore >= mMTMinimumHighDiscriminantScore And _
               AMTData(i).PeptideProphetProbability >= mMTMinimumPeptideProphetProbability Then
                blnAddMassTag = True
            Else
                blnAddMassTag = False
            End If
        Else
            blnAddMassTag = True
        End If
        
        If blnAddMassTag Then
            mMTCnt = mMTCnt + 1
            mMTInd(mMTCnt - 1) = mMTCnt - 1
            mMTOrInd(mMTCnt - 1) = i             'index; not the ID
            mMTMWN14(mMTCnt - 1) = AMTData(i).MW
            mMTMWN15(mMTCnt - 1) = AMTData(i).MW + glN14N15_DELTA * AMTData(i).CNT_N       ' N15 is always fixed
            Select Case samtDef.NETorRT
            Case glAMT_NET
                 mMTNET(mMTCnt - 1) = AMTData(i).NET
            Case glAMT_RT_or_PNET
                 mMTNET(mMTCnt - 1) = AMTData(i).PNET
            End Select
            mMTMods(mMTCnt - 1) = ""
        End If
   Next i
   If chkPEO.Value = vbChecked Then         'correct based on cys number for PEO label
      UpdateStatus "Adding PEO labeled peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
          If CysCnt > 0 Then
             If GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods Then
                ' Dynamic Mods
                For j = 1 To CysCnt
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * glPEO
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_PEO & "/" & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this MT tag
                mMTMWN14(i) = mMTMWN14(i) + CysCnt * glPEO
                mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                mMTMods(i) = mMTMods(i) & " " & MOD_TKN_PEO & "/" & CysCnt
             End If
          End If
      Next i
   End If
   'yeah, yeah I know that same cysteine can not be labeled with PEO and ICAT at the same
   'time but who cares anyway I can fix this here easily
   If chkICATLt.Value = vbChecked Then         'correct based on cys number for ICAT label
      UpdateStatus "Adding D0 ICAT labeled peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
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
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D0 & "/" & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this MT tag
                ' However, if use also has ICAT_d0 enabled, we need to duplicate this
                '  MT tag first
                If chkICATHv.Value = vbChecked Then
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + CysLeft * glICAT_Heavy
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & CysLeft
                End If
                
                ' Now update this MT tag to have ICAT_d0 on all the cysteines
                mMTMWN14(i) = mMTMWN14(i) + CysLeft * glICAT_Light
                mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ICAT_D0 & "/" & CysLeft
             End If
          End If
      Next i
   End If
   
   If chkICATHv.Value = vbChecked Then         'correct based on cys number for ICAT label
      UpdateStatus "Adding D8 ICAT labeled peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
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
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & j
                Next j
             Else
                If chkICATLt.Value = vbChecked Then
                    ' We shouldn't have reached this code since all of the cysteines should
                    '  have been assigned ICAT_d0 or ICAT_d8
                    Debug.Assert False
                Else
                    ' Static Mods
                    ' Simply update the stats for this MT tag
                    mMTMWN14(i) = mMTMWN14(i) + CysLeft * glICAT_Heavy
                    mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & CysLeft
                End If
             End If
          End If
      Next i
   End If
   
   If chkAlkylation.Value = vbChecked Then         'correct based on cys number for alkylation label
      UpdateStatus "Adding alkylated peptides..."
      TmpCnt = mMTCnt
      For i = 0 To TmpCnt - 1
          CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
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
                    mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTNET(mMTCnt - 1) = mMTNET(i)
                    mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ALK & "/" & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this MT tag
                mMTMWN14(i) = mMTMWN14(i) + CysLeft * AlkMWCorrection
                mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
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
          ' Add dblResidueModMass once to the entire MT tag
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
                  mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                  mMTNET(mMTCnt - 1) = mMTNET(i)
                  mMTMods(mMTCnt - 1) = mMTMods(i) & " " & strResModToken & "/" & strResiduesToModify & j
              Next j
           Else
              ' Static Mods
              ' Simply update the stats for this MT tag
              mMTMWN14(i) = mMTMWN14(i) + ResidueOccurrenceCount * dblResidueModMass
              mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
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
   End If
End If

If Not PrepareSearchInternalStandards() Then
     Debug.Assert False
     Call DestroySearchStructures
     Exit Function
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

Private Sub RecordSearchResultsInData()
    ' Step through mUMCMatchStats() and add the ID's for each UMC to all of the members of each UMC
    
    Dim lngIndex As Long, lngMemberIndex As Long
    Dim lngUMCIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long                  'absolute index in mMT... arrays
    Dim lngMassTagIndexOriginal As Long                 'absolute index in AMT... arrays
    Dim lngInternalStdIndexOriginal As Long
    Dim lngIonIndexOriginal As Long
    Dim blnAddRef As Boolean
    Dim lngIonCountUpdated As Long
    
    Dim AMTorInternalStdRef As String
    Dim dblMatchMass As Double, dblMatchNET As Double
    Dim dblCurrMW As Double, AbsMWErr As Double
    Dim dblSLiCScore As Double
    Dim dblDelSLiC As Double
    
    Dim CurrScan As Long
     
    'always reinitialize statistics arrays
    InitAMTStat
    
    KeyPressAbortProcess = 0
    
    CheckNETEquationStatus

On Error GoTo RecordSearchResultsInDataErrorHandler

    With GelData(CallerID)
        For lngIndex = 0 To mMatchStatsCount - 1
            If lngIndex Mod 25 = 0 Then
                UpdateStatus "Storing results: " & LongToStringWithCommas(lngIndex) & " / " & LongToStringWithCommas(mMatchStatsCount)
                If KeyPressAbortProcess > 1 Then Exit For
            End If
            
            lngUMCIndexOriginal = mUMCMatchStats(lngIndex).UMCIndex
            
            If mUMCMatchStats(lngIndex).IDIsInternalStd Then
                lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(lngIndex).IDIndex)
                dblMatchMass = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).MonoisotopicMass
                dblMatchNET = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).NET
            Else
                lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngIndex).IDIndex)
                lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
                
                If LastSearchTypeN14N15 = SEARCH_N14 Then
                    ' N14
                    dblMatchMass = mMTMWN14(mUMCMatchStats(lngIndex).IDIndex)
                Else
                    ' N15
                    dblMatchMass = mMTMWN15(mUMCMatchStats(lngIndex).IDIndex)
                End If
                dblMatchNET = AMTData(lngMassTagIndexOriginal).NET
            End If
            
            dblSLiCScore = mUMCMatchStats(lngIndex).SLiCScore
            dblDelSLiC = mUMCMatchStats(lngIndex).DelSLiC
            
            ' Record the search results in each of the members of the UMC
            For lngMemberIndex = 0 To GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassCount - 1
                lngIonIndexOriginal = GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMInd(lngMemberIndex)
                blnAddRef = False
                
                Select Case GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMType(lngMemberIndex)
                Case glCSType
                    dblCurrMW = .CSData(lngIonIndexOriginal).AverageMW
                    CurrScan = .CSData(lngIonIndexOriginal).ScanNumber
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = dblCurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select
                    
                    If mUMCMatchStats(lngIndex).IDIsInternalStd Then
                        AMTorInternalStdRef = ConstructInternalStdReference(.CSData(lngIonIndexOriginal).AverageMW, ConvertScanToNET(CLng(.CSData(lngIonIndexOriginal).ScanNumber)), lngInternalStdIndexOriginal, dblSLiCScore, dblDelSLiC)
                    Else
                        AMTorInternalStdRef = ConstructAMTReference(.CSData(lngIonIndexOriginal).AverageMW, ConvertScanToNET(CLng(.CSData(lngIonIndexOriginal).ScanNumber)), 0, lngMassTagIndexOriginal, dblMatchMass, dblSLiCScore, dblDelSLiC)
                    End If
                    
                    If Len(.CSData(lngIonIndexOriginal).MTID) = 0 Then
                        blnAddRef = True
                    ElseIf InStr(.CSData(lngIonIndexOriginal).MTID, AMTorInternalStdRef) <= 0 Then
                        blnAddRef = True
                    End If
                    
                    If blnAddRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        ' If this specific data point is not within tolerance, then mark it as "Inherited"
                        If Not IsValidMatch(dblCurrMW, AbsMWErr, CurrScan, dblMatchNET, dblMatchMass) Then
                            AMTorInternalStdRef = Trim(AMTorInternalStdRef)
                            If Right(AMTorInternalStdRef, 1) = glARG_SEP Then
                                AMTorInternalStdRef = Left(AMTorInternalStdRef, Len(AMTorInternalStdRef) - 1)
                            End If
                            AMTorInternalStdRef = AMTorInternalStdRef & AMTMatchInheritedMark
                        End If
                        
                        InsertBefore .CSData(lngIonIndexOriginal).MTID, AMTorInternalStdRef
                    End If
                Case glIsoType
                    dblCurrMW = GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField)
                    CurrScan = .IsoData(lngIonIndexOriginal).ScanNumber
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = dblCurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select

                    If mUMCMatchStats(lngIndex).IDIsInternalStd Then
                        AMTorInternalStdRef = ConstructInternalStdReference(GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField), ConvertScanToNET(.IsoData(lngIonIndexOriginal).ScanNumber), lngInternalStdIndexOriginal, dblSLiCScore, dblDelSLiC)
                    Else
                        AMTorInternalStdRef = ConstructAMTReference(GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField), ConvertScanToNET(.IsoData(lngIonIndexOriginal).ScanNumber), 0, lngMassTagIndexOriginal, dblMatchMass, dblSLiCScore, dblDelSLiC)
                    End If
                    
                    ' Only add AMTorInternalStdRef if .MTID does not contain it
                    ' First perform a quick check to see if .MTID is empty
                    ' If it's not empty, then use InStr to see if .MTID contains AMTorInternalStdRef (a relatively slow operation)
                    If Len(.IsoData(lngIonIndexOriginal).MTID) = 0 Then
                        blnAddRef = True
                    ElseIf InStr(.IsoData(lngIonIndexOriginal).MTID, AMTorInternalStdRef) <= 0 Then
                        blnAddRef = True
                    End If
                    
                    If blnAddRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        If Not IsValidMatch(dblCurrMW, AbsMWErr, CurrScan, dblMatchNET, dblMatchMass) Then
                            AMTorInternalStdRef = Trim(AMTorInternalStdRef)
                            If Right(AMTorInternalStdRef, 1) = glARG_SEP Then
                                AMTorInternalStdRef = Left(AMTorInternalStdRef, Len(AMTorInternalStdRef) - 1)
                            End If
                            AMTorInternalStdRef = AMTorInternalStdRef & AMTMatchInheritedMark
                        End If
                        
                        InsertBefore .IsoData(lngIonIndexOriginal).MTID, AMTorInternalStdRef
                    End If
                End Select
            Next lngMemberIndex
        Next lngIndex
    End With
    
    If KeyPressAbortProcess <= 1 Then
        AddToAnalysisHistory CallerID, "Stored search results in ions; recorded all MT tag hits for each UMC in all members of the UMC; total ions updated = " & Trim(lngIonCountUpdated)
    End If
    
    Exit Sub

RecordSearchResultsInDataErrorHandler:
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->RecordSearchResultsInData"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured while storing the search results in the data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub RemoveAMTMatchesFromUMCs(blnQueryUser As Boolean)

    Dim eResponse As VbMsgBoxResult
    
    If blnQueryUser And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Remove MT tag references from the data in the current gel?", vbQuestion + vbYesNoCancel + vbDefaultButton1)
    Else
        eResponse = vbYes
    End If
    
    If eResponse = vbYes Then
        TraceLog 5, "RemoveAMTMatchesFromUMCs", "Calling RemoveAMT"
        RemoveAMT CallerID, glScope.glSc_All
        If eInternalStdSearchMode <> issmFindOnlyMassTags Or cboInternalStdSearchMode.ListIndex <> issmFindOnlyMassTags Then
            TraceLog 5, "RemoveInternalStdFromUMCs", "Calling RemoveInternalStd"
            RemoveInternalStd CallerID, glScope.glSc_All
        End If
        
        TraceLog 5, "RemoveAMTMatchesFromUMCs", "Setting GelStatus(CallerID).Dirty = True"
        
        GelStatus(CallerID).Dirty = True
        If blnQueryUser Then AddToAnalysisHistory CallerID, "Deleted MT tag search results from ions"
        
        TraceLog 5, "RemoveAMTMatchesFromUMCs", "MT tag references removed"
        UpdateStatus "MT tag references removed."
    End If


End Sub

Private Function RobustNETValuesEnabled() As Boolean
    If GelData(CallerID).CustomNETsDefined And Not cChkBox(chkDisableCustomNETs) Then
        RobustNETValuesEnabled = True
    Else
        RobustNETValuesEnabled = False
    End If
End Function

Private Function SearchUMCSingleMass(ByVal ClassInd As Long) As Long
'-----------------------------------------------------------------------------
'returns number of hits found for UMC with index ClassInd; -1 in case of error;
'  -2 if skipped since hit already present
'-----------------------------------------------------------------------------

Dim MWTolAbsBroad As Double     ' MWTol used to compute the MatchScore
Dim NETTolBroad As Double       ' NETTol used to compute the MatchScore

Dim MWTolAbsFinal As Double     ' Final MWErr required
Dim NETTolFinal As Double

Dim dblClassMass As Double

Dim blnProceed As Boolean

Dim lngIndex As Long
Dim lngMassTagIndexPointer As Long

Dim MassTagHitCount As Long

On Error GoTo err_SearchUMCSingleMass

If ManageCurrID(MNG_RESET) Then
    If SearchType = SEARCH_PAIRED Or SearchType = SEARCH_PAIRED_PLUS_NON_PAIRED Then
        Select Case N14N15
        Case SEARCH_N14     'don't search if this class is found only as heavy member
            If eClsPaired(ClassInd) = umcpHeavyUnique Or _
               eClsPaired(ClassInd) = umcpHeavyMultiple Then
                SearchUMCSingleMass = 0
                Exit Function
            End If
        Case SEARCH_N15     'don't search if this class is found only as light member
            If eClsPaired(ClassInd) = umcpLightUnique Or _
               eClsPaired(ClassInd) = umcpLightMultiple Then
                SearchUMCSingleMass = 0
                Exit Function
            End If
        End Select
    End If
    
    ' Define the tolerances
    SearchAMTDefineTolerances CallerID, ClassInd, samtDef, dblClassMass, MWTolAbsBroad, NETTolBroad, MWTolAbsFinal, NETTolFinal
    
    With GelUMC(CallerID)
        blnProceed = True
        If samtDef.SkipReferenced Then
            ' Skip this UMC if one or more of its members have an AMT match defined
            blnProceed = Not IsAMTReferencedByUMC(.UMCs(ClassInd), CallerID)
        End If
    End With
    
    
    If blnProceed Then
        If eInternalStdSearchMode <> issmFindOnlyInternalStandards Then
            ' Search for the MT tags using broad tolerances
            SearchUMCSingleMassAMT GelUMC(CallerID).UMCs(ClassInd), MWTolAbsBroad, NETTolBroad
        End If
        ' MassTagHitCount holds the number of matching MT tags, excluding Internal Standards
        MassTagHitCount = mCurrIDCnt
    Else
        ' Skipped UMC since already has a match
        MassTagHitCount = -2
    End If
    
    If eInternalStdSearchMode <> issmFindOnlyMassTags Then
        ' Search for Internal Standards using broad tolerances
        SearchUMCSingleMassInternalStd GelUMC(CallerID).UMCs(ClassInd), MWTolAbsBroad, NETTolBroad
    End If
     
    If mCurrIDCnt > 0 Then
        ' Populate .IDIndexOriginal
        For lngIndex = 0 To mCurrIDCnt - 1
            If mCurrIDMatches(lngIndex).IDIsInternalStd Then
                lngMassTagIndexPointer = mInternalStdIndexPointers(mCurrIDMatches(lngIndex).IDInd)
                mCurrIDMatches(lngIndex).IDIndexOriginal = lngMassTagIndexPointer
            Else
                lngMassTagIndexPointer = mMTInd(mCurrIDMatches(lngIndex).IDInd)
                mCurrIDMatches(lngIndex).IDIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            End If
        Next lngIndex
        
        
        ' Next compute the Match Scores
        SearchAMTComputeSLiCScores mCurrIDCnt, mCurrIDMatches, dblClassMass, MWTolAbsFinal, NETTolFinal, mSearchRegionShape
        
        If mCurrIDCnt > 0 Then
            If ManageCurrID(MNG_TRIM) Then
                Call AddCurrIDsToAllIDs(ClassInd)
            End If
        End If
    End If
     
    
''    If eInternalStdSearchMode <> issmFindOnlyInternalStandards Then
''        ' Search for MT tags
''        If blnProceed Then
''            ' First search for the MT tags using broad tolerances
''            SearchUMCSingleMassAMT GelUMC(CallerID).UMCs(ClassInd), MWTolAbsBroad, NETTolBroad
''
''            ' Populate .IDIndexOriginal
''            For lngIndex = 0 To mCurrIDCnt - 1
''                lngMassTagIndexPointer = mMTInd(mCurrIDMatches(lngIndex).IDInd)
''                mCurrIDMatches(lngIndex).IDIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
''            Next lngIndex
''
''            ' Next compute the Match Scores
''            SearchAMTComputeSLiCScores mCurrIDCnt, mCurrIDMatches, dblClassMass, MWTolAbsFinal, NETTolFinal
''        End If
''
''        If mCurrIDCnt > 0 Then
''            If ManageCurrID(MNG_TRIM) Then
''                Call AddCurrIDsToAllIDs(ClassInd, False)
''            End If
''            MassTagHitCount = mCurrIDCnt
''        Else
''            Call ManageCurrID(MNG_ERASE)
''        End If
''    End If
''
''    If eInternalStdSearchMode <> issmFindOnlyMassTags Then
''        If ManageCurrID(MNG_RESET) Then
''            ' Search for Internal Standards
''            ' First search for standards using broad tolerances
''            SearchUMCSingleMassInternalStd GelUMC(CallerID).UMCs(ClassInd), MWTolAbsBroad, NETTolBroad
''
''            ' Populate .IDIndexOriginal
''            For lngIndex = 0 To mCurrIDCnt - 1
''                lngMassTagIndexPointer = mMTInd(mCurrIDMatches(lngIndex).IDInd)
''                mCurrIDMatches(lngIndex).IDIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
''            Next lngIndex
''
''            ' Next compute the Match Scores
''            SearchAMTComputeSLiCScores mCurrIDCnt, mCurrIDMatches, dblClassMass, MWTolAbsFinal, samtDef.NETTol
''
''            If mCurrIDCnt > 0 Then
''                If ManageCurrID(MNG_TRIM) Then
''                    Call AddCurrIDsToAllIDs(ClassInd, True)
''                End If
''            Else
''                Call ManageCurrID(MNG_ERASE)
''            End If
''        End If
''    End If
Else
    UpdateStatus "Error managing memory."
    MassTagHitCount = -1
End If

SearchUMCSingleMass = MassTagHitCount

Exit Function

err_SearchUMCSingleMass:
SearchUMCSingleMass = -1
End Function

Private Sub SearchUMCSingleMassAMT(ByRef udtTestUMC As udtUMCType, ByVal dblMWTol As Double, ByVal dblNETTol As Double)
    ' Compare this UMC's mass, NET, and charge with the MT tags

    Dim FastSearchMatchInd As Long
    Dim MatchInd1 As Long, MatchInd2 As Long
    
    Dim dblMassTagMass As Double
    Dim dblMassTagNET As Double

    ' Only need to call MWFastSearch once, sending it udtTestUMC.ClassMW
    MatchInd1 = 0
    MatchInd2 = -1
    If MWFastSearch.FindIndexRange(udtTestUMC.ClassMW, dblMWTol, MatchInd1, MatchInd2) Then
        If MatchInd1 <= MatchInd2 Then
            ' One or more MT tags is within dblMWTol of the median UMC mass
            
            ' Now test each MT tag with dblMWTol and dblNETTol and record the matches
            For FastSearchMatchInd = MatchInd1 To MatchInd2
                
                dblMassTagMass = MWFastSearch.GetMWByIndex(FastSearchMatchInd)
                dblMassTagNET = mMTNET(mMTInd(FastSearchMatchInd))
                
                SearchUMCSingleMassValidate FastSearchMatchInd, dblMWTol, dblNETTol, udtTestUMC, dblMassTagMass, dblMassTagNET, False
            
            Next FastSearchMatchInd
        End If
    End If

End Sub

Private Sub SearchUMCSingleMassInternalStd(ByRef udtTestUMC As udtUMCType, ByVal dblMWTol As Double, ByVal dblNETTol As Double)
    ' Compare this UMC's mass, NET, and charge with the Internal Standard in UMCInternalStandards

    Dim FastSearchMatchInd As Long
    Dim MatchInd1 As Long, MatchInd2 As Long
    Dim udtInternalStd As udtInternalStandardEntryType
    
    If UMCInternalStandards.Count <= 0 Then Exit Sub
    
    ' Only need to call InternalStdFastSearch once, sending it udtTestUMC.ClassMW
    MatchInd1 = 0
    MatchInd2 = -1
    If InternalStdFastSearch.FindIndexRange(udtTestUMC.ClassMW, dblMWTol, MatchInd1, MatchInd2) Then
        If MatchInd1 <= MatchInd2 Then
            ' One or more Internal Standard is within dblMWTol of the median UMC mass
            
            ' Now test each MT tag with dblMWTol and dblNETTol and record the matches
            For FastSearchMatchInd = MatchInd1 To MatchInd2
                
                udtInternalStd = UMCInternalStandards.InternalStandards(mInternalStdIndexPointers(FastSearchMatchInd))
                   
                If SearchUMCTestCharge(udtTestUMC.ClassRepType, udtTestUMC.ClassRepInd, udtInternalStd) Then
                    SearchUMCSingleMassValidate FastSearchMatchInd, dblMWTol, dblNETTol, udtTestUMC, udtInternalStd.MonoisotopicMass, udtInternalStd.NET, True
                End If
                
            Next FastSearchMatchInd
        End If
    End If

End Sub

Private Sub SearchUMCSingleMassValidate(ByVal FastSearchMatchInd As Long, ByVal dblMWTol As Double, ByVal dblNETTol As Double, ByRef udtTestUMC As udtUMCType, ByVal dblMassTagMass As Double, ByVal dblMassTagNET As Double, ByVal blnIsInternalStdMatch As Boolean)
    ' Note: This sub is called by both SearchUMCSingleMassAMT and SearchUMCSingleMassInternalStd
    
    ' Check if the match is within NET and mass tolerance
    ' If it is, increment mCurrIDMatches().MatchingMemberCount
    
    ' Note that since we used udtTestUMC.ClassMW in the call to FindIndexRange(), not all members
    '  of the class will necessarily have a matching mass
    
    ' Additionally, it is possible that the conglomerate class mass will match a MT tag, but none
    ' of the members will match.  An example of this is a UMC with two members, weighing 500.0 and 502.0 Da
    ' The median mass is 501.0 Da.  If the dblMWTol = 0.1, then the median will match a MT tag of 501 Da,
    '  but none of the members will match.  In this case, we'll record the match,
    '  but place a 0 in mCurrIDMatches().MatchingMemberCount

    
    Dim blnFirstMatchFound As Boolean
    Dim lngMemberIndex As Long
    
    Dim dblCurrMW As Double
    Dim dblNETDifference As Double
    
    With udtTestUMC
        ' See if each MassTag is within the NET tolerance of any of the members of the class
        ' Alternatively, if .UseUMCConglomerateNET = True, then use the NET value of the class representative
        
        blnFirstMatchFound = False
        If glbPreferencesExpanded.UseUMCConglomerateNET Then
            If SearchUMCTestNET(.ClassRepType, .ClassRepInd, dblMassTagNET, dblNETTol, dblNETDifference) Then
                
                ' Either: AMT Matches this UMC's median mass and Class Rep NET
                ' Or:     Internal Standard Matches this UMC's median mass, Class Rep NET, and charge
                ' Thus:   Add to mCurrIDMatches()
                
                If mCurrIDCnt > UBound(mCurrIDMatches) Then ManageCurrID (MNG_ADD_START_SIZE)
                
                mCurrIDMatches(mCurrIDCnt).IDInd = FastSearchMatchInd
                mCurrIDMatches(mCurrIDCnt).MatchingMemberCount = 0
                mCurrIDMatches(mCurrIDCnt).SLiCScore = -1             ' Set this to -1 for now
                mCurrIDMatches(mCurrIDCnt).MassErr = .ClassMW - dblMassTagMass
                mCurrIDMatches(mCurrIDCnt).NETErr = dblNETDifference
                mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = blnIsInternalStdMatch
                mCurrIDCnt = mCurrIDCnt + 1
                
                blnFirstMatchFound = True
            End If
        End If
        
        If blnFirstMatchFound Or Not glbPreferencesExpanded.UseUMCConglomerateNET Then
            For lngMemberIndex = 0 To .ClassCount - 1
                If SearchUMCTestNET(CInt(.ClassMType(lngMemberIndex)), .ClassMInd(lngMemberIndex), dblMassTagNET, dblNETTol, dblNETDifference) Then
                    
                    Select Case .ClassMType(lngMemberIndex)
                    Case glCSType
                         dblCurrMW = GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).AverageMW
                    Case glIsoType
                         dblCurrMW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)), samtDef.MWField)
                    End Select
                    
                    If Not blnFirstMatchFound Then
                        ' We haven't had a match for this index yet; add to mCurrIDMatches()
                        
                        If mCurrIDCnt > UBound(mCurrIDMatches) Then ManageCurrID (MNG_ADD_START_SIZE)
                        
                        mCurrIDMatches(mCurrIDCnt).IDInd = FastSearchMatchInd
                        mCurrIDMatches(mCurrIDCnt).MatchingMemberCount = 0
                        mCurrIDMatches(mCurrIDCnt).SLiCScore = -1    ' Set this to -1 for now
                        mCurrIDMatches(mCurrIDCnt).MassErr = .ClassMW - dblMassTagMass
                        mCurrIDMatches(mCurrIDCnt).NETErr = dblNETDifference
                        mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = blnIsInternalStdMatch
                        mCurrIDCnt = mCurrIDCnt + 1
                        blnFirstMatchFound = True
                    End If

                    ' See if the member is within mass tolerance
                    If Abs(dblMassTagMass - dblCurrMW) <= dblMWTol Then
                        ' Yes, within both mass and NET tolerance; increment mCurrIDMatches().MatchingMemberCount
                        mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount = mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount + 1
                    End If
                End If
            Next lngMemberIndex
        End If
    End With

End Sub

Private Function SearchUMCTestCharge(eMemberType As glDistType, lngMemberIndex As Long, udtInternalStd As udtInternalStandardEntryType) As Boolean
    ' Make sure at least one of the charges for this Net Adj Locker is present in the UMC

    Dim blnValidHit As Boolean
    Dim intCharge As Integer
    
    Select Case eMemberType
    Case glCSType
        intCharge = GelData(CallerID).CSData(lngMemberIndex).Charge
    Case glIsoType
        intCharge = GelData(CallerID).IsoData(lngMemberIndex).Charge
    End Select

    ' Make sure at least one of the charges for this Net Adj Locker is present in the UMC
    If intCharge >= udtInternalStd.ChargeMinimum And _
       intCharge <= udtInternalStd.ChargeMaximum Then
       ' Valid Charge
       blnValidHit = True
    Else
        blnValidHit = False
    End If
    
    SearchUMCTestCharge = blnValidHit

End Function

Private Function SearchUMCTestNET(eMemberType As glDistType, lngMemberIndex As Long, dblAMTNet As Double, ByVal dblNETTol As Double, ByRef dblNETDifference As Double) As Boolean
    
    Dim lngScan As Long
    Dim blnNETMatch As Boolean
    
    Select Case eMemberType
    Case glCSType
        lngScan = GelData(CallerID).CSData(lngMemberIndex).ScanNumber
    Case glIsoType
        lngScan = GelData(CallerID).IsoData(lngMemberIndex).ScanNumber
    End Select
    
    blnNETMatch = False
    dblNETDifference = ConvertScanToNET(lngScan) - dblAMTNet
    If dblNETTol > 0 Then
        If Abs(dblNETDifference) <= dblNETTol Then
            blnNETMatch = True
        End If
    Else
        ' NETTol = 0; assume a match
        blnNETMatch = True
    End If

    SearchUMCTestNET = blnNETMatch
    
End Function

Public Sub SetAlkylationMWCorrection(dblMass As Double)
    txtAlkylationMWCorrection = dblMass
    AlkMWCorrection = dblMass
End Sub

Private Sub SetDBSearchModType(blnDynamicMods As Boolean)
    If blnDynamicMods Then
        optDBSearchModType(MODS_DYNAMIC).Value = True
    Else
        optDBSearchModType(MODS_FIXED).Value = True
    End If
    GelSearchDef(CallerID).AMTSearchMassMods.DynamicMods = optDBSearchModType(MODS_DYNAMIC).Value
End Sub

Public Sub SetDBSearchNType(blnUseN15 As Boolean)
    If blnUseN15 Then
        optN(1).Value = True
        N14N15 = SEARCH_N15
    Else
        optN(0).Value = True
        N14N15 = SEARCH_N14
    End If
End Sub

Private Sub SetDefaultOptions(blnUseToleranceRefinementSettings As Boolean)

    Dim udtSearchDef As SearchAMTDefinition
    SetDefaultSearchAMTDef udtSearchDef, UMCNetAdjDef
    
    If blnUseToleranceRefinementSettings Then
        With udtSearchDef
            .MWTol = DEFAULT_TOLERANCE_REFINEMENT_MW_TOL
            .TolType = DEFAULT_TOLERANCE_REFINEMENT_MW_TOL_TYPE
            .NETTol = DEFAULT_TOLERANCE_REFINEMENT_NET_TOL
        End With
    End If
    
    SetCheckBox chkUpdateGelDataWithSearchResults, True
    cboAMTSearchResultsBehavior.ListIndex = asrbAMTSearchResultsBehaviorConstants.asrbAutoRemoveExisting
    
    If blnUseToleranceRefinementSettings Then
        cboSearchRegionShape.ListIndex = srsSearchRegionShapeConstants.srsRectangular
    Else
        cboSearchRegionShape.ListIndex = srsSearchRegionShapeConstants.srsElliptical
    End If
    
    If APP_BUILD_DISABLE_MTS Then
        cboInternalStdSearchMode.ListIndex = issmInternalStandardSearchModeConstants.issmFindOnlyMassTags
    Else
        cboInternalStdSearchMode.ListIndex = issmInternalStandardSearchModeConstants.issmFindWithMassTags
    End If
    
    txtDBSearchMinimumHighNormalizedScore.Text = 0
    txtDBSearchMinimumHighDiscriminantScore.Text = 0
    txtDBSearchMinimumPeptideProphetProbability.Text = 0
    
    optNETorRT(udtSearchDef.NETorRT).Value = True
    SetCheckBox chkUseUMCConglomerateNET, True
    SetCheckBox chkDisableCustomNETs, False
    
    optTolType(udtSearchDef.TolType).Value = True
    txtMWTol.Text = udtSearchDef.MWTol
    
    txtNETTol = udtSearchDef.NETTol
    
    SetCheckBox chkPEO, False
    SetCheckBox chkICATLt, False
    SetCheckBox chkICATHv, False
    SetCheckBox chkAlkylation, False
    txtAlkylationMWCorrection = 57.0215
    
    cboResidueToModify.ListIndex = 0
    txtResidueToModifyMass.Text = 0
    
    optDBSearchModType(MODS_DYNAMIC).Value = True
    optN(0).Value = True
    
    SetETMode etGANET

    PickParameters
    
End Sub

Private Sub SetETMode(eETModeDesired As glETType)
    Dim i As Long
    Dim eETModeToUse As glETType

On Error Resume Next

    If RobustNETValuesEnabled() Then
        lblETType.Caption = "Using Custom NETs"
    Else
        If GelAnalysis(CallerID) Is Nothing Then
            eETModeToUse = etGenericNET
        Else
            eETModeToUse = eETModeDesired
        End If
        
        Select Case eETModeToUse
        Case etGenericNET
            If eETModeDesired <> etGenericNET Then
                txtNETFormula.Text = GelUMCNETAdjDef(CallerID).NETFormula
            Else
                txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
            End If
        Case etTICFitNET
          With GelAnalysis(CallerID)
            If .NET_Slope <> 0 Then
                txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
            Else
                txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
            End If
          End With
          If Err Then
             MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
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
             MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
             Exit Sub
          End If
        End Select
        For i = mnuET.LBound To mnuET.UBound
            If i = eETModeDesired Then
               mnuET(i).Checked = True
               lblETType.Caption = "ET: " & mnuET(i).Caption
            Else
               mnuET(i).Checked = False
            End If
        Next i
        Call txtNETFormula_LostFocus        'make sure expression evaluator is
                                            'initialized for this formula
    End If

End Sub

Public Sub SetInternalStandardSearchMode(eInternalStdSearchMode As issmInternalStandardSearchModeConstants)
    On Error Resume Next
    
    If APP_BUILD_DISABLE_MTS Then
        eInternalStdSearchMode = issmInternalStandardSearchModeConstants.issmFindOnlyMassTags
    End If
    
    cboInternalStdSearchMode.ListIndex = eInternalStdSearchMode
    If cboInternalStdSearchMode.ListIndex < 0 Then cboInternalStdSearchMode.ListIndex = 0
End Sub

Public Sub SetMinimumHighDiscriminantScore(sngMinimumHighDiscriminantScore As Single)
    txtDBSearchMinimumHighDiscriminantScore = sngMinimumHighDiscriminantScore
End Sub

Public Sub SetMinimumHighNormalizedScore(sngMinimumHighNormalizedScore As Single)
    txtDBSearchMinimumHighNormalizedScore = sngMinimumHighNormalizedScore
End Sub

Public Sub SetMinimumPeptideProphetProbability(sngMinimumPeptideProphetProbability As Single)
    txtDBSearchMinimumPeptideProphetProbability = sngMinimumPeptideProphetProbability
End Sub

Private Sub ShowErrorDistribution2DForm()
    frmErrorDistribution2DLoadedData.CallerID = CallerID
    frmErrorDistribution2DLoadedData.Show vbModal
    
    ' Make sure the search tolerances displayed match those now in memory (in case the user performed tolerance refinement)
    DisplayCurrentSearchTolerances
End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuFSepExportToDatabase.Visible = blnVisible
    mnuFExportResultsToDBbyUMC.Visible = blnVisible
    mnuFExportDetailedMemberInformation.Visible = blnVisible

    mnuMTLoadMT.Visible = blnVisible
    
    cboInternalStdSearchMode.Visible = blnVisible
    lblInternalStdSearchMode.Visible = blnVisible
    
End Sub

Private Function ShowOrSaveResultsByIon(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True, Optional ByVal blnIncludeORFInfo As Boolean = True) As Long
'---------------------------------------------------
'report results, listing by data point (by ion)
' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
' If blnIncludeORFInfo = True, then attempts to connect to the database and retrieve the ORF information for each MT tag
'
' Returns 0 if no error, the error number if an error
'---------------------------------------------------
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim strBaseMatchInfo As String
Dim strLineOut As String
Dim fname As String

' Note: AMTRefs() is 1-based
Dim AMTRefs() As String
Dim AMTRefsCnt As Long
Dim i As Long
Dim lngExportCount As Long
Dim strSepChar As String
Dim dblIonMass As Double

Dim lngAMTID() As Long      ' AMT ID's, copied from the globaly array AMTData()
Dim lngIndex As Long

' The following is used to lookup the mass of each MT tag, given the MT tag ID
' It is initialized using AMTData()
Dim objAMTIDFastSearch As New FastSearchArrayLong

' Since AMT masses can be modified (e.g. alkylation), we must use the Pointer determined above
'   to search mMTOrInd() to determine the correct match
' We'll use objMTOrIndFastSearch, initializing using mMTOrInd()
Dim objMTOrIndFastSearch As New FastSearchArrayLong

' In order to add to the confusion, we must actually lookup the mMTOrInd() value in mMTInd()
' This requires a 3rd FastSearch object, initialized using mMTInd()
Dim objMTIndFastSearch As New FastSearchArrayLong

' This last FastSearch object is used to lookup an ORF name
Dim objORFNameFastSearch As New FastSearchArrayLong
Dim blnSuccess As Boolean

On Error GoTo err_ShowOrSaveResultsByIon

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

Select Case LastSearchTypeN14N15
Case SEARCH_N14
     NTypeStr = MOD_TKN_N14
Case SEARCH_N15
     NTypeStr = MOD_TKN_N15
End Select

UpdateStatus "Sorting MT lookup arrays"
mKeyPressAbortProcess = 0

' Construct the MT tag ID lookup arrays
' We need to copy the AMT ID's from AMTData() to lngAMTID() since AMTData().ID is a String array that actually simply holds numbers
If AMTCnt > 0 Then
    ReDim lngAMTID(1 To AMTCnt)
    For lngIndex = 1 To AMTCnt
        lngAMTID(lngIndex) = CLngSafe(AMTData(lngIndex).ID)
    Next lngIndex
Else
    ReDim lngAMTID(1 To 1)
End If

blnSuccess = objAMTIDFastSearch.Fill(lngAMTID())
Debug.Assert blnSuccess

blnSuccess = objMTOrIndFastSearch.Fill(mMTOrInd())
Debug.Assert blnSuccess

blnSuccess = objMTIndFastSearch.Fill(mMTInd())
Debug.Assert blnSuccess

Me.MousePointer = vbHourglass

'temporary file for results output
fname = GetTempFolder() & RawDataTmpFile
If Len(strOutputFilePath) > 0 Then fname = strOutputFilePath
Set ts = fso.OpenTextFile(fname, ForWriting, True)

strSepChar = LookupDefaultSeparationCharacter()

strLineOut = "Index" & strSepChar & "Scan" & strSepChar & "ChargeState" & strSepChar & "MonoMW" & strSepChar & "Abundance" & strSepChar
strLineOut = strLineOut & "Fit" & strSepChar & "ER" & strSepChar & "LockerID" & strSepChar & "FreqShift" & strSepChar & "MassCorrection" & strSepChar & "MultiMassTagHitCount" & strSepChar
strLineOut = strLineOut & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagMods"

If blnIncludeORFInfo Then strLineOut = strLineOut & strSepChar & "MultiORFCount" & strSepChar & "ORFName"
ts.WriteLine strLineOut

With GelData(CallerID)
  If .CSLines > 0 Then ts.WriteLine "Charge State Data Block"
  For i = 1 To .CSLines
      If i Mod 500 = 0 Then
        UpdateStatus "Preparing results: " & Trim(i) & " / " & Trim(.CSLines)
        If mKeyPressAbortProcess > 1 Then Exit For
      End If
      If Not IsNull(.CSData(i).MTID) Then
         If IsAMTReferenced(.CSData(i).MTID) Then
            AMTRefsCnt = GetAMTRefFromString2(.CSData(i).MTID, AMTRefs())
            If AMTRefsCnt > 0 Then
            'for Charge State standard deviation is used on place of Fit
                dblIonMass = .CSData(i).AverageMW
                strBaseMatchInfo = i & strSepChar & .CSData(i).ScanNumber & strSepChar _
                   & .CSData(i).Charge & strSepChar & .CSData(i).AverageMW & strSepChar _
                   & .CSData(i).Abundance & strSepChar & .CSData(i).MassStDev & strSepChar
                strBaseMatchInfo = strBaseMatchInfo & LookupExpressionRatioValue(CallerID, i, False)
                If GelLM(CallerID).CSCnt > 0 Then   'we have mass correction
                   strBaseMatchInfo = strBaseMatchInfo & strSepChar & GelLM(CallerID).CSLckID(i) & strSepChar _
                        & GelLM(CallerID).CSFreqShift(i) & strSepChar _
                        & GelLM(CallerID).CSMassCorrection(i)
                Else
                   strBaseMatchInfo = strBaseMatchInfo & strSepChar & strSepChar & strSepChar
                End If
            
                WriteAMTMatchesForIon ts, strBaseMatchInfo, dblIonMass, AMTRefs(), AMTRefsCnt, objAMTIDFastSearch, objMTOrIndFastSearch, objMTIndFastSearch, lngExportCount, blnIncludeORFInfo, objORFNameFastSearch, strSepChar
            End If
         End If
      End If
  Next i
  If .IsoLines > 0 Then ts.WriteLine "Isotopic Data Block"
  For i = 1 To .IsoLines
      If i Mod 500 = 0 Then
        UpdateStatus "Preparing results: " & Trim(i) & " / " & Trim(.IsoLines)
        If mKeyPressAbortProcess > 1 Then Exit For
      End If
      If Not IsNull(.IsoData(i).MTID) Then
         If IsAMTReferenced(.IsoData(i).MTID) Then
            AMTRefsCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTRefs())
            If AMTRefsCnt > 0 Then
                dblIonMass = .IsoData(i).MonoisotopicMW
                strBaseMatchInfo = i & strSepChar & .IsoData(i).ScanNumber & strSepChar _
                   & .IsoData(i).Charge & strSepChar & .IsoData(i).MonoisotopicMW & strSepChar _
                   & .IsoData(i).Abundance & strSepChar & .IsoData(i).Fit & strSepChar
                strBaseMatchInfo = strBaseMatchInfo & LookupExpressionRatioValue(CallerID, i, True)
                If GelLM(CallerID).IsoCnt > 0 Then
                   strBaseMatchInfo = strBaseMatchInfo & strSepChar & GelLM(CallerID).IsoLckID(i) & strSepChar _
                         & GelLM(CallerID).IsoFreqShift(i) & strSepChar _
                         & GelLM(CallerID).IsoMassCorrection(i)
                Else
                   strBaseMatchInfo = strBaseMatchInfo & strSepChar & strSepChar & strSepChar
                End If
                
                WriteAMTMatchesForIon ts, strBaseMatchInfo, dblIonMass, AMTRefs(), AMTRefsCnt, objAMTIDFastSearch, objMTOrIndFastSearch, objMTIndFastSearch, lngExportCount, blnIncludeORFInfo, objORFNameFastSearch, strSepChar
            End If
         End If
      End If
  Next i
End With
ts.Close

UpdateStatus ""

If blnDisplayResults Then
   frmDataInfo.Tag = "EXP"
   frmDataInfo.Show vbModal
Else
    ' MonroeMod
    AddToAnalysisHistory CallerID, "Exported " & lngExportCount & " search results to text file: " & fname
End If
ShowOrSaveResultsByIon = 0

ShowOrSaveResultsCleanup:

Set ts = Nothing
Set fso = Nothing

Set objAMTIDFastSearch = Nothing
Set objMTOrIndFastSearch = Nothing
Set objMTIndFastSearch = Nothing
Set objORFNameFastSearch = Nothing

Exit Function

err_ShowOrSaveResultsByIon:
Debug.Assert False
ShowOrSaveResultsByIon = Err.Number
LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.ShowOrSaveResultsByIon"
Resume ShowOrSaveResultsCleanup

End Function

Public Function ShowOrSaveResultsByUMC(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True, Optional ByVal blnIncludeORFInfo As Boolean = True) As Long
'-------------------------------------
'report identified unique mass classes
' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
' If blnIncludeORFInfo = True, then attempts to connect to the database and retrieve the ORF information for each MT tag
'
' Returns 0 if no error, the error number if an error
'-------------------------------------
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim strLineOut As String, strLineOutMiddle As String, strLineOutEnd As String
Dim strMinMaxCharges As String
Dim fname As String
Dim mgInd As Long
Dim lngUMCIndexOriginal As Long                     'absolute index of UMC
Dim lngMassTagIndexPointer As Long                  'absolute index in mMT... arrays
Dim lngMassTagIndexOriginal As Long                 'absolute index in AMT... arrays
Dim lngInternalStdIndexOriginal As Long
Dim strSepChar As String

Dim dblMatchMass As Double
Dim dblMatchNET As Double
Dim strMatchID As String
Dim strInternalStdDescription As String
Dim sngPeptideProphetProbability As Single

Dim dblMassErrorPPM As Double
Dim lngScanClassRep As Long
Dim lngScanIndex As Long

Dim dblGANETClassRep As Double
Dim dblGANETError As Double
Dim objORFNameFastSearch As New FastSearchArrayLong
Dim blnSuccess As Boolean

Dim lngPairIndex As Long

Dim objP1IndFastSearch As FastSearchArrayLong
Dim objP2IndFastSearch As FastSearchArrayLong
Dim blnPairsPresent As Boolean

Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
Dim udtPairMatchStats() As udtPairMatchStatsType

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
fname = GetTempFolder() & RawDataTmpFile
If Len(strOutputFilePath) > 0 Then fname = strOutputFilePath
Set ts = fso.OpenTextFile(fname, ForWriting, True)

Select Case LastSearchTypeN14N15
Case SEARCH_N14
     NTypeStr = MOD_TKN_N14
Case SEARCH_N15
     NTypeStr = MOD_TKN_N15
End Select

' Initialize the PairIndex lookup objects
blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)

strSepChar = LookupDefaultSeparationCharacter()

' UMCIndex; ScanStart; ScanEnd; ScanClassRep; GANETClassRep; UMCMonoMW; UMCMWStDev; UMCMWMin; UMCMWMax; UMCAbundance; ClassStatsChargeBasis; ChargeStateMin; ChargeStateMax; UMCMZForChargeBasis; UMCMemberCount; UMCMemberCountUsedForAbu; UMCAverageFit; PairIndex; ExpressionRatio; MultiMassTagHitCount; MassTagID; MassTagMonoMW; MassTagMods; MemberCountMatchingMassTag; MassErrorPPM; GANETError; SLiC_Score; Del_SLiC; IsInternalStdMatch; PeptideProphetProbability; TIC_from_Raw_Data; Deisotoping_Peak_Count
strLineOut = "UMCIndex" & strSepChar & "ScanStart" & strSepChar & "ScanEnd" & strSepChar & "ScanClassRep" & strSepChar & "NETClassRep" & strSepChar & "UMCMonoMW" & strSepChar & "UMCMWStDev" & strSepChar & "UMCMWMin" & strSepChar & "UMCMWMax" & strSepChar & "UMCAbundance" & strSepChar
strLineOut = strLineOut & "ClassStatsChargeBasis" & strSepChar & "ChargeStateMin" & strSepChar & "ChargeStateMax" & strSepChar & "UMCMZForChargeBasis" & strSepChar & "UMCMemberCount" & strSepChar & "UMCMemberCountUsedForAbu" & strSepChar & "UMCAverageFit" & strSepChar & "PairIndex" & strSepChar
strLineOut = strLineOut & "ExpressionRatio" & strSepChar & "ExpressionRatioStDev" & strSepChar & "ExpressionRatioChargeStateBasisCount" & strSepChar & "ExpressionRatioMemberBasisCount" & strSepChar
strLineOut = strLineOut & "MultiMassTagHitCount" & strSepChar
strLineOut = strLineOut & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagMods" & strSepChar & "MemberCountMatchingMassTag" & strSepChar & "MassErrorPPM" & strSepChar & "NETError" & strSepChar & "SLiC_Score" & strSepChar & "Del_SLiC" & strSepChar & "IsInternalStdMatch" & strSepChar & "PeptideProphetProbability" & strSepChar
strLineOut = strLineOut & "TIC_from_Raw_Data" & strSepChar & "Deisotoping_Peak_Count"
If blnIncludeORFInfo Then strLineOut = strLineOut & strSepChar & "MultiORFCount" & strSepChar & "ORFName"
ts.WriteLine strLineOut

    For mgInd = 0 To mMatchStatsCount - 1
        lngUMCIndexOriginal = mUMCMatchStats(mgInd).UMCIndex
        
        If mUMCMatchStats(mgInd).IDIsInternalStd Then
            lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(mgInd).IDIndex)
            With UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal)
                dblMatchMass = .MonoisotopicMass
                dblMatchNET = .NET
                strMatchID = .SeqID
                strInternalStdDescription = .Description
            End With
            sngPeptideProphetProbability = 0
        Else
            lngMassTagIndexPointer = mMTInd(mUMCMatchStats(mgInd).IDIndex)
            lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            
            If LastSearchTypeN14N15 = SEARCH_N14 Then
                ' N14
                dblMatchMass = mMTMWN14(mUMCMatchStats(mgInd).IDIndex)
            Else
                ' N15
                dblMatchMass = mMTMWN15(mUMCMatchStats(mgInd).IDIndex)
            End If
        
            dblMatchNET = AMTData(lngMassTagIndexOriginal).NET
            ' Future: dblMatchNETStDev = AMTData(lngMassTagIndexOriginal).NETStDev
            strMatchID = AMTData(lngMassTagIndexOriginal).ID
            
            sngPeptideProphetProbability = AMTData(lngMassTagIndexOriginal).PeptideProphetProbability
        End If
    
        With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
            Select Case .ClassRepType
            Case gldtCS
                lngScanClassRep = GelData(CallerID).CSData(.ClassRepInd).ScanNumber
            Case gldtIS
                lngScanClassRep = GelData(CallerID).IsoData(.ClassRepInd).ScanNumber
            Case Else
                Debug.Assert False
                lngScanClassRep = (.MinScan + .MaxScan) / 2
            End Select
            
            dblGANETClassRep = ScanToGANET(CallerID, lngScanClassRep)
            
            strLineOut = lngUMCIndexOriginal & strSepChar & .MinScan & strSepChar & .MaxScan & strSepChar & lngScanClassRep & strSepChar & Format(dblGANETClassRep, "0.0000") & strSepChar & Round(.ClassMW, 6) & strSepChar
            strLineOut = strLineOut & Round(.ClassMWStD, 6) & strSepChar & .MinMW & strSepChar & .MaxMW & strSepChar & .ClassAbundance & strSepChar
            
            strMinMaxCharges = ClsStat(lngUMCIndexOriginal, ustChargeMin) & strSepChar & ClsStat(lngUMCIndexOriginal, ustChargeMax) & strSepChar
            
            ' Record ClassStatsChargeBasis, ChargeMin, ChargeMax, UMCMZForChargeBasis, UMCMemberCount, and UMCMemberCountUsedForAbu
            If GelUMC(CallerID).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge & strSepChar
                strLineOut = strLineOut & strMinMaxCharges
                strLineOut = strLineOut & Round(MonoMassToMZ(.ClassMW, .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge), 6) & strSepChar
            Else
                strLineOut = strLineOut & 0 & strSepChar
                strLineOut = strLineOut & strMinMaxCharges
                strLineOut = strLineOut & Round(MonoMassToMZ(.ClassMW, CInt(GelData(CallerID).IsoData(.ClassRepInd).Charge)), 6) & strSepChar
            End If
            
            strLineOut = strLineOut & .ClassCount & strSepChar
            
            ' Record UMCMemberCountUsedForAbu
            If GelUMC(CallerID).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Count & strSepChar
            Else
                strLineOut = strLineOut & .ClassCount & strSepChar
            End If
            
        End With
        
        strLineOut = strLineOut & Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), 3) & strSepChar
        
        ' Now start populating strLineOutEnd
        strLineOutEnd = ""
        
        If mUMCMatchStats(mgInd).IDIsInternalStd Then
            strLineOutEnd = strLineOutEnd & "0" & strSepChar
        Else
            strLineOutEnd = strLineOutEnd & mUMCMatchStats(mgInd).MultiAMTHitCount & strSepChar
        End If
    
        With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
            dblMassErrorPPM = MassToPPM(.ClassMW - dblMatchMass, .ClassMW)
            dblGANETError = dblGANETClassRep - dblMatchNET
        End With
        
        strLineOutEnd = strLineOutEnd & strMatchID & strSepChar & Round(dblMatchMass, 6) & strSepChar
        
        If Not mUMCMatchStats(mgInd).IDIsInternalStd Then
            strLineOutEnd = strLineOutEnd & NTypeStr
            If Len(mMTMods(lngMassTagIndexPointer)) > 0 Then
                strLineOutEnd = strLineOutEnd & " " & mMTMods(lngMassTagIndexPointer)
            End If
        End If
        
        strLineOutEnd = strLineOutEnd & strSepChar & mUMCMatchStats(mgInd).MemberHitCount & strSepChar & Round(dblMassErrorPPM, 4) & strSepChar & Round(dblGANETError, NET_PRECISION)
        strLineOutEnd = strLineOutEnd & strSepChar & Round(mUMCMatchStats(mgInd).SLiCScore, 4)
        strLineOutEnd = strLineOutEnd & strSepChar & Round(mUMCMatchStats(mgInd).DelSLiC, 4)
        strLineOutEnd = strLineOutEnd & strSepChar & mUMCMatchStats(mgInd).IDIsInternalStd
        strLineOutEnd = strLineOutEnd & strSepChar & Round(sngPeptideProphetProbability, 5)
        
        lngScanIndex = LookupScanNumberRelativeIndex(CallerID, lngScanClassRep)
        If lngScanIndex = 0 Then
            lngScanClassRep = LookupScanNumberClosest(CallerID, lngScanClassRep)
            lngScanIndex = LookupScanNumberRelativeIndex(CallerID, lngScanClassRep)
        End If
        
        With GelData(CallerID).ScanInfo(lngScanIndex)
            strLineOutEnd = strLineOutEnd & strSepChar & Round(.TIC, 1)
            strLineOutEnd = strLineOutEnd & strSepChar & .NumIsotopicSignatures
        End With
        
        lngPairIndex = -1
        lngPairMatchCount = 0
        ReDim udtPairMatchStats(0)
        InitializePairMatchStats udtPairMatchStats(0)
        If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
            lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, objP1IndFastSearch, objP2IndFastSearch, False, (LastSearchTypeN14N15 = SEARCH_N15), lngPairMatchCount, udtPairMatchStats())
        End If
        
        If lngPairMatchCount > 0 Then
            For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                With udtPairMatchStats(lngPairMatchIndex)
                    strLineOutMiddle = Trim(.PairIndex) & strSepChar & Trim(.ExpressionRatio) & strSepChar & Trim(.ExpressionRatioStDev) & strSepChar & Trim(.ExpressionRatioChargeStateBasisCount) & strSepChar & Trim(.ExpressionRatioMemberBasisCount) & strSepChar
                    
                    If Not blnIncludeORFInfo Then
                        ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd
                    Else
                        If mUMCMatchStats(mgInd).IDIsInternalStd Then
                            ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd & strSepChar & "1" & strSepChar & strInternalStdDescription
                        Else
                            WriteORFResults ts, strLineOut & strLineOutMiddle & strLineOutEnd, CLngSafe(AMTData(lngMassTagIndexOriginal).ID), objORFNameFastSearch, strSepChar
                        End If
                    End If
                    
                End With
            Next lngPairMatchIndex
        Else
            ' No pair, and thus no expression ratio values
            strLineOutMiddle = Trim(-1) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar
            
            If Not blnIncludeORFInfo Then
                ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd
            Else
                If mUMCMatchStats(mgInd).IDIsInternalStd Then
                    ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd & strSepChar & "1" & strSepChar & strInternalStdDescription
                Else
                    WriteORFResults ts, strLineOut & strLineOutMiddle & strLineOutEnd, CLngSafe(AMTData(lngMassTagIndexOriginal).ID), objORFNameFastSearch, strSepChar
                End If
            End If
            
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
Debug.Assert False
ShowOrSaveResultsByUMC = Err.Number
LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.ShowOrSaveResultsByUMC"
Set fso = Nothing
End Function

Private Sub StartExportResultsToDBbyUMC()
    Dim eResponse As VbMsgBoxResult
    Dim strStatus As String
    Dim strUMCSearchMode As String
    
On Error GoTo ExportResultsToDBErrorHandler
    
    If mMatchStatsCount = 0 And Not glbPreferencesExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches Then
        MsgBox "Search results not found in memory.", vbInformation + vbOKOnly, "Nothing to Export"
    Else
        eResponse = MsgBox("Proceed with exporting of the search results to the database?  This is an advanced feature that should normally only be performed during VIPER Automated PRISM Analysis Mode.  If you continue, you will be prompted for a password.", vbQuestion + vbYesNo + vbDefaultButton1, "Export Results")
        If eResponse = vbYes Then
            If QueryUserForExportToDBPassword(, False) Then
                ' Update the text in MD_Parameters
                strUMCSearchMode = FindSettingInAnalysisHistory(CallerID, UMC_SEARCH_MODE_SETTING_TEXT, , True, ":", ";")
                If Right(strUMCSearchMode, 1) = ")" Then strUMCSearchMode = Left(strUMCSearchMode, Len(strUMCSearchMode) - 1)
                GelAnalysis(CallerID).MD_Parameters = ConstructAnalysisParametersText(CallerID, strUMCSearchMode, AUTO_SEARCH_UMC_CONGLOMERATE)
                
                strStatus = ExportMTDBbyUMC(True, mnuFExportDetailedMemberInformation.Checked)
                MsgBox strStatus, vbInformation + vbOKOnly, glFGTU
            Else
                MsgBox "Invalid password, export aborted.", vbExclamation Or vbOKOnly, "Invalid"
            End If
        End If
    End If
    
    Exit Sub
    
ExportResultsToDBErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.StartExportResultsToDBbyUMC"
    Resume Next

End Sub

Public Function StartSearchAll() As Long
' Returns the number of hits
Dim i As Long
Dim strMessage As String
Dim lngHitCount As Long
Dim blnCustomNETsAreAvailable As Boolean

On Error Resume Next

mKeyPressAbortProcess = 0
cmdSearchAllUMCs.Visible = False
cmdRemoveAMTMatchesFromUMCs.Visible = False
DoEvents

If mMatchStatsCount > 0 Then    'something already identified
   Call DestroyIDStructures
End If
SearchType = SEARCH_ALL
samtDef.SearchScope = glScope.glSc_All
mSearchRegionShape = cboSearchRegionShape.ListIndex

' Note: PrepareMTArrays will update mSearchUsedCustomNETs based on .CustomNETsDefined
blnCustomNETsAreAvailable = GelData(CallerID).CustomNETsDefined
If cChkBox(chkDisableCustomNETs) Then
    GelData(CallerID).CustomNETsDefined = False
End If

CheckNETEquationStatus
eInternalStdSearchMode = cboInternalStdSearchMode.ListIndex

Select Case glbPreferencesExpanded.AMTSearchResultsBehavior
Case asrbAutoRemoveExisting
    RemoveAMTMatchesFromUMCs False
    samtDef.SkipReferenced = False
Case asrbKeepExisting
    samtDef.SkipReferenced = False
Case asrbKeepExistingAndSkip
    samtDef.SkipReferenced = True
Case Else
    Debug.Assert False
    RemoveAMTMatchesFromUMCs False
    samtDef.SkipReferenced = False
End Select

If PrepareMTArrays() Then
    mUMCCountSkippedSinceRefPresent = 0
    txtUniqueMatchStats.Text = ""
    For i = 0 To ClsCnt - 1
        If i Mod 25 = 0 Then
           UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
           If mKeyPressAbortProcess > 1 Then Exit For
        End If
        lngHitCount = SearchUMCSingleMass(i)
        If lngHitCount = -2 Then mUMCCountSkippedSinceRefPresent = mUMCCountSkippedSinceRefPresent + 1
    Next i
    LastSearchTypeN14N15 = N14N15
    
    
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
        strMessage = DisplayHitSummary("all")
        
        If chkUpdateGelDataWithSearchResults Then
            ' Store the search results in the gel data
            If mMatchStatsCount > 0 Then RecordSearchResultsInData
            UpdateStatus strMessage
        End If
    Else
        UpdateStatus "Search aborted."
    End If
Else
   UpdateStatus "Error searching for matches"
End If

GelData(CallerID).CustomNETsDefined = blnCustomNETsAreAvailable
cmdSearchAllUMCs.Visible = True
cmdRemoveAMTMatchesFromUMCs.Visible = True
DoEvents

StartSearchAll = mMatchStatsCount
End Function

Public Function StartSearchPaired() As Long
' Returns the number of hits
Dim i As Long
Dim strMessage As String
Dim lngHitCount As Long
Dim blnCustomNETsAreAvailable As Boolean

On Error Resume Next

mKeyPressAbortProcess = 0
cmdSearchAllUMCs.Visible = False
cmdRemoveAMTMatchesFromUMCs.Visible = False
If mMatchStatsCount > 0 Then    'something already identified
   Call DestroyIDStructures
End If
SearchType = SEARCH_PAIRED
samtDef.SearchScope = glScope.glSc_All
mSearchRegionShape = cboSearchRegionShape.ListIndex

' Note: PrepareMTArrays will update mSearchUsedCustomNETs based on .CustomNETsDefined
blnCustomNETsAreAvailable = GelData(CallerID).CustomNETsDefined
If cChkBox(chkDisableCustomNETs) Then
    GelData(CallerID).CustomNETsDefined = False
End If

CheckNETEquationStatus
eInternalStdSearchMode = cboInternalStdSearchMode.ListIndex

If PrepareMTArrays() Then
    mUMCCountSkippedSinceRefPresent = 0
    For i = 0 To ClsCnt - 1
        If i Mod 25 = 0 Then
           UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
           If mKeyPressAbortProcess > 1 Then Exit For
        End If
        If eClsPaired(i) <> umcpNone Then
           lngHitCount = SearchUMCSingleMass(i)
           If lngHitCount = -2 Then mUMCCountSkippedSinceRefPresent = mUMCCountSkippedSinceRefPresent + 1
        End If
    Next i
    LastSearchTypeN14N15 = N14N15
    
    If GelAnalysis(CallerID).MD_Type = stNotDefined Or GelAnalysis(CallerID).MD_Type = stStandardIndividual Then
        ' Only update MD_Type if it is currently stStandardIndividual
        GelAnalysis(CallerID).MD_Type = stPairsO16O18
    End If

    If mKeyPressAbortProcess <= 1 Then
        strMessage = DisplayHitSummary("paired")
        
        If chkUpdateGelDataWithSearchResults Then
            ' Store the search results in the gel data
            If mMatchStatsCount > 0 Then RecordSearchResultsInData
            UpdateStatus strMessage
        End If
    Else
        UpdateStatus "Search aborted."
    End If
Else
   UpdateStatus "Error searching for matches"
End If

GelData(CallerID).CustomNETsDefined = blnCustomNETsAreAvailable
cmdSearchAllUMCs.Visible = True
cmdRemoveAMTMatchesFromUMCs.Visible = True

StartSearchPaired = mMatchStatsCount
End Function

Public Function StartSearchLightPairsPlusNonPaired() As Long
' Returns the number of hits
Dim i As Long
Dim strMessage As String
Dim lngHitCount As Long
Dim blnCustomNETsAreAvailable As Boolean

On Error Resume Next

' Force N14N15 to be SEARCH_N14
SetDBSearchNType False

mKeyPressAbortProcess = 0
cmdSearchAllUMCs.Visible = False
cmdRemoveAMTMatchesFromUMCs.Visible = False
If mMatchStatsCount > 0 Then    'something already identified
   Call DestroyIDStructures
End If
SearchType = SEARCH_PAIRED_PLUS_NON_PAIRED
samtDef.SearchScope = glScope.glSc_All
mSearchRegionShape = cboSearchRegionShape.ListIndex

' Note: PrepareMTArrays will update mSearchUsedCustomNETs based on .CustomNETsDefined
blnCustomNETsAreAvailable = GelData(CallerID).CustomNETsDefined
If cChkBox(chkDisableCustomNETs) Then
    GelData(CallerID).CustomNETsDefined = False
End If

CheckNETEquationStatus
eInternalStdSearchMode = cboInternalStdSearchMode.ListIndex

If PrepareMTArrays() Then
    mUMCCountSkippedSinceRefPresent = 0
    For i = 0 To ClsCnt - 1
        If i Mod 25 = 0 Then
           UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
           If mKeyPressAbortProcess > 1 Then Exit For
        End If
        lngHitCount = SearchUMCSingleMass(i)
        If lngHitCount = -2 Then mUMCCountSkippedSinceRefPresent = mUMCCountSkippedSinceRefPresent + 1
    Next i
    LastSearchTypeN14N15 = N14N15
    
    If GelAnalysis(CallerID).MD_Type = stNotDefined Or GelAnalysis(CallerID).MD_Type = stStandardIndividual Then
        ' Only update MD_Type if it is currently stStandardIndividual
        GelAnalysis(CallerID).MD_Type = stPairsO16O18
    End If

    If mKeyPressAbortProcess <= 1 Then
        strMessage = DisplayHitSummary("light pairs plus non-paired")
        
        If chkUpdateGelDataWithSearchResults Then
            ' Store the search results in the gel data
            If mMatchStatsCount > 0 Then RecordSearchResultsInData
            UpdateStatus strMessage
        End If
    Else
        UpdateStatus "Search aborted."
    End If
Else
   UpdateStatus "Error searching for matches"
End If

GelData(CallerID).CustomNETsDefined = blnCustomNETsAreAvailable
cmdSearchAllUMCs.Visible = True
cmdRemoveAMTMatchesFromUMCs.Visible = True

StartSearchLightPairsPlusNonPaired = mMatchStatsCount

End Function

Public Function StartSearchNonPaired() As Long
' Returns the number of hits
Dim i As Long
Dim strMessage As String
Dim lngHitCount As Long
Dim blnCustomNETsAreAvailable As Boolean

On Error Resume Next

mKeyPressAbortProcess = 0
cmdSearchAllUMCs.Visible = False
cmdRemoveAMTMatchesFromUMCs.Visible = False
If mMatchStatsCount > 0 Then    'something already identified
   Call DestroyIDStructures
End If
SearchType = SEARCH_NON_PAIRED
samtDef.SearchScope = glScope.glSc_All
mSearchRegionShape = cboSearchRegionShape.ListIndex

' Note: PrepareMTArrays will update mSearchUsedCustomNETs based on .CustomNETsDefined
blnCustomNETsAreAvailable = GelData(CallerID).CustomNETsDefined
If cChkBox(chkDisableCustomNETs) Then
    GelData(CallerID).CustomNETsDefined = False
End If

CheckNETEquationStatus
eInternalStdSearchMode = cboInternalStdSearchMode.ListIndex

If PrepareMTArrays() Then
    mUMCCountSkippedSinceRefPresent = 0
    For i = 0 To ClsCnt - 1
        If i Mod 25 = 0 Then
           UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
           If mKeyPressAbortProcess > 1 Then Exit For
        End If
        If eClsPaired(i) = umcpNone Then
           lngHitCount = SearchUMCSingleMass(i)
           If lngHitCount = -2 Then mUMCCountSkippedSinceRefPresent = mUMCCountSkippedSinceRefPresent + 1
        End If
    Next i
    LastSearchTypeN14N15 = N14N15
    
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
        strMessage = DisplayHitSummary("non-paired")
        
        If chkUpdateGelDataWithSearchResults Then
            ' Store the search results in the gel data
            If mMatchStatsCount > 0 Then RecordSearchResultsInData
            UpdateStatus strMessage
        End If
    Else
        UpdateStatus "Search aborted."
    End If
Else
   UpdateStatus "Error searching for matches"
End If

GelData(CallerID).CustomNETsDefined = blnCustomNETsAreAvailable
cmdSearchAllUMCs.Visible = True
cmdRemoveAMTMatchesFromUMCs.Visible = True

StartSearchNonPaired = mMatchStatsCount
End Function

Private Sub UpdateUMCsPairingStatusNow()
    Dim blnSuccess As Boolean
    blnSuccess = UpdateUMCsPairingStatus(CallerID, eClsPaired())
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub WriteAMTMatchesForIon(ts As TextStream, strLineOutPrefix As String, dblIonMass As Double, AMTRefs() As String, AMTRefsCnt As Long, objAMTIDFastSearch As FastSearchArrayLong, objMTOrIndFastSearch As FastSearchArrayLong, objMTIndFastSearch As FastSearchArrayLong, ByRef lngExportCount, blnIncludeORFInfo As Boolean, objORFNameFastSearch As FastSearchArrayLong, Optional strSepChar As String = glARG_SEP)
    ' Note: AMTRefs() is 1-based
    
    Dim strBaseMatchInfo As String
    Dim strLineOut As String
    Dim lngAMTRefIndex As Long
    Dim lngMassTagID As Long
    
    Dim lngOriginalAMTIndex As Long             ' Index of the AMT in AMTData().MW, etc.
    Dim lngMTOrIndIndexOriginal As Long         ' Index of the AMT in mMTOrInd()
    
    Dim lngMatchingIndices() As Long            ' Used with both objAMTIDFastSearch and objMTOrIndFastSearch
    Dim lngMatchCount As Long
    
    Dim lngMTIndMatchingIndices() As Long       ' Index of the AMT in mMTInd()
    Dim lngMTIndMatchCount As Long
    
    Dim lngPointerIndex As Long
    
    Dim dblAMTMass As Double
    Dim dblBestAMTMass As Double, dblBestAMTMassDiff As Double, strBestAMTMods As String
    
    ' AMTRefsCnt is the number of AMTs that this ion matched (aka MultiMassTagHitCount)
    strBaseMatchInfo = strLineOutPrefix & strSepChar & AMTRefsCnt
    For lngAMTRefIndex = 1 To AMTRefsCnt         'extract MT tag ID
        lngMassTagID = CLng(GetIDFromString(AMTRefs(lngAMTRefIndex), AMTMark, AMTIDEnd))
        
        strLineOut = strBaseMatchInfo & strSepChar & lngMassTagID & strSepChar
        
        If objAMTIDFastSearch.FindMatchingIndices(lngMassTagID, lngMatchingIndices(), lngMatchCount) Then
            ' Match Found
            
            lngOriginalAMTIndex = lngMatchingIndices(0)
            
            ' Now look for lngOriginalAMTIndex in lngMTOrInd()
            ' It could actually be present several times if the mass modifications were
            '  defined as dynamic (rather than static)
            If objMTOrIndFastSearch.FindMatchingIndices(lngOriginalAMTIndex, lngMatchingIndices(), lngMatchCount) Then
                ' Match Found
                
                dblBestAMTMass = 0
                strBestAMTMods = ""
                For lngPointerIndex = 0 To lngMatchCount - 1
                    lngMTOrIndIndexOriginal = lngMatchingIndices(lngPointerIndex)
                    
                    ' Now look for lngMTOrIndIndexOriginal in mMTInd()
                    If objMTIndFastSearch.FindMatchingIndices(lngMTOrIndIndexOriginal, lngMTIndMatchingIndices(), lngMTIndMatchCount) Then
                        ' Match found
                        
                        If LastSearchTypeN14N15 = SEARCH_N14 Then
                            ' N14
                            dblAMTMass = mMTMWN14(lngMTIndMatchingIndices(0))
                        Else
                            ' N15
                            dblAMTMass = mMTMWN15(lngMTIndMatchingIndices(0))
                        End If
                        
                        If dblBestAMTMass = 0 Then
                            dblBestAMTMass = dblAMTMass
                            dblBestAMTMassDiff = Abs(dblAMTMass - dblIonMass)
                            strBestAMTMods = mMTMods(lngMTOrIndIndexOriginal)
                        Else
                            If Abs(dblAMTMass - dblIonMass) < dblBestAMTMassDiff Then
                                dblBestAMTMass = dblAMTMass
                                dblBestAMTMassDiff = Abs(dblAMTMass - dblIonMass)
                                strBestAMTMods = mMTMods(lngMTOrIndIndexOriginal)
                            End If
                        End If
                    End If
                Next lngPointerIndex
                
                dblAMTMass = dblBestAMTMass
                If dblBestAMTMass <> 0 Then
                    Debug.Assert Abs(dblAMTMass - dblIonMass) < 0.5
                    Debug.Assert Abs(dblAMTMass - AMTData(lngOriginalAMTIndex).MW) < 0.0001 Or dblAMTMass > AMTData(lngOriginalAMTIndex).MW
                End If
            Else
                dblAMTMass = 0
            End If
        Else
            dblAMTMass = 0
        End If
        
        strLineOut = strLineOut & Round(dblAMTMass, 6) & strSepChar & NTypeStr
        If Len(strBestAMTMods) > 0 Then
            strLineOut = strLineOut & " " & strBestAMTMods
        End If
        strLineOut = strLineOut & strSepChar
        
        If Not blnIncludeORFInfo Then
            ts.WriteLine strLineOut
        Else
            WriteORFResults ts, strLineOut, lngMassTagID, objORFNameFastSearch, strSepChar
        End If
        
        lngExportCount = lngExportCount + 1
    Next lngAMTRefIndex

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

Private Sub cboAMTSearchResultsBehavior_Click()
    On Error Resume Next
    If Not bLoading Then
        glbPreferencesExpanded.AMTSearchResultsBehavior = cboAMTSearchResultsBehavior.ListIndex
    End If
End Sub

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

Private Sub chkDisableCustomNETs_Click()
    EnableDisableNETFormulaControls
End Sub

Private Sub chkUseUMCConglomerateNET_Click()
    glbPreferencesExpanded.UseUMCConglomerateNET = cChkBox(chkUseUMCConglomerateNET)
End Sub

Private Sub cmdCancel_Click()
    mKeyPressAbortProcess = 2
    KeyPressAbortProcess = 2
End Sub

Private Sub cmdRemoveAMTMatchesFromUMCs_Click()
    RemoveAMTMatchesFromUMCs True
End Sub

Private Sub cmdSearchAllUMCs_Click()
    StartSearchAll
End Sub

Private Sub cmdSetDefaults_Click()
    SetDefaultOptions False
End Sub

Private Sub cmdSetDefaultsForToleranceRefinement_Click()
    SetDefaultOptions True
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

If APP_BUILD_DISABLE_ADVANCED Then
    chkDisableCustomNETs.Visible = False
End If

ShowHidePNNLMenus

'set current Search Definition values
DisplayCurrentSearchTolerances

With samtDef
    If glbPreferencesExpanded.AMTSearchResultsBehavior = asrbKeepExistingAndSkip Then
        .SkipReferenced = True
    Else
        .SkipReferenced = False
    End If
    
    optNETorRT(.NETorRT).Value = True
    
    'save old value and set search on "search all"
    OldSearchFlag = .SearchFlag
    .SearchFlag = 0         'search all
    
    mnuET(etGANET).Checked = True
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

With glbPreferencesExpanded
    cboAMTSearchResultsBehavior.ListIndex = .AMTSearchResultsBehavior
    SetCheckBox chkUseUMCConglomerateNET, .UseUMCConglomerateNET
End With

With glbPreferencesExpanded.MTSConnectionInfo
    ExpAnalysisSPName = .spPutAnalysis
    'ExpPeakSPName = .spPutPeak
    ExpUmcSPName = .spPutUMC
    ExpUMCMemberSPName = .spPutUMCMember
    ExpUmcMatchSPName = .spPutUMCMatch
    ExpUmcInternalStdMatchSPName = .spPutUMCInternalStdMatch
    ExpQuantitationDescription = .spAddQuantitationDescription
End With

If Not GelAnalysis(CallerID) Is Nothing Then
    mMDTypeSaved = GelAnalysis(CallerID).MD_Type
Else
    mMDTypeSaved = stNotDefined
End If

If Len(ExpUmcSPName) = 0 Then
    ExpUmcSPName = "AddFTICRUmc"
End If
Debug.Assert ExpUmcSPName = "AddFTICRUmc"

If Len(ExpUmcMatchSPName) = 0 Then
    ExpUmcMatchSPName = "AddFTICRUmcMatch"
End If
Debug.Assert ExpUmcMatchSPName = "AddFTICRUmcMatch"

If Len(ExpUmcInternalStdMatchSPName) = 0 Then
    ExpUmcInternalStdMatchSPName = "AddFTICRUmcInternalStdMatch"
End If
Debug.Assert ExpUmcInternalStdMatchSPName = "AddFTICRUmcInternalStdMatch"

If Len(ExpQuantitationDescription) = 0 Then
    ExpQuantitationDescription = "AddQuantitationDescription"
End If
Debug.Assert ExpQuantitationDescription = "AddQuantitationDescription"

If Len(ExpAnalysisSPName) = 0 Then
    ExpAnalysisSPName = "AddMatchMaking"
End If
Debug.Assert ExpAnalysisSPName = "AddMatchMaking"

' September 2004: Unused Variable
''If Len(ExpPeakSPName) = 0 Then
''    ExpPeakSPName = "AddFTICRPeak"
''End If
''Debug.Assert ExpPeakSPName = "AddFTICRPeak"

' Possibly add a checkmark to the mnuFReportIncludeORFs menu
mnuFReportIncludeORFs.Checked = glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput
Exit Sub

FormLoadErrorHandler:
LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.Form_Load"
Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
' Restore .SearchFlag using the saved value
samtDef.SearchFlag = OldSearchFlag
If Not objMTDBNameLookupClass Is Nothing Then
    objMTDBNameLookupClass.DeleteData
    Set objMTDBNameLookupClass = Nothing
End If
' Restore .MD_Type from mMDTypeSaved
If Not GelAnalysis(CallerID) Is Nothing Then
    GelAnalysis(CallerID).MD_Type = mMDTypeSaved
End If
End Sub

Private Sub mnuET_Click(Index As Integer)
    SetETMode (Index)
End Sub

Private Sub mnuETHeader_Click()
Call PickParameters
End Sub

Private Sub mnuF_Click()
Call PickParameters
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFDeleteExcludedPairs_Click()
    Me.DeleteExcludedPairsWrapper
End Sub

Private Sub mnuFExcludeAmbiguous_Click()
    Me.ExcludeAmbiguousPairsWrapper False
End Sub

Private Sub mnuFExcludeAmbiguousHitsOnly_Click()
    Me.ExcludeAmbiguousPairsWrapper True
End Sub

Private Sub mnuFExportDetailedMemberInformation_Click()
    mnuFExportDetailedMemberInformation.Checked = Not mnuFExportDetailedMemberInformation.Checked
End Sub

Private Sub mnuFExportResultsToDBbyUMC_Click()
    StartExportResultsToDBbyUMC
End Sub

Private Sub mnuFMassCalAndToleranceRefinement_Click()
    ShowErrorDistribution2DForm
End Sub

Private Sub mnuFReportByIon_Click()
ShowOrSaveResultsByIon "", True, mnuFReportIncludeORFs.Checked
End Sub

Private Sub mnuFReportByUMC_Click()
    ShowOrSaveResultsByUMC "", True, mnuFReportIncludeORFs.Checked
End Sub

Private Sub mnuFReportIncludeORFs_Click()
    mnuFReportIncludeORFs.Checked = Not mnuFReportIncludeORFs.Checked
    glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput = mnuFReportIncludeORFs.Checked
End Sub

Private Sub mnuFResetExclusionFlags_Click()
Dim strMessage As String
strMessage = PairsResetExclusionFlag(CallerID)
UpdateUMCsPairingStatusNow
UpdateStatus strMessage
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

Private Sub mnuFSearchPairedPlusNonPaired_Click()
StartSearchLightPairsPlusNonPaired
End Sub

Private Sub mnuMT_Click()
Call PickParameters
End Sub

Private Sub mnuMTLoadLegacy_Click()
    LoadLegacyMassTags
End Sub

Private Sub mnuMTLoadMT_Click()
'------------------------------------------------------------
'load/reload MT tags
'------------------------------------------------------------
If Not GelAnalysis(CallerID) Is Nothing Then
   Call LoadMTDB(True)
Else
   WarnUserNotConnectedToDB CallerID, True
   lblMTStatus.Caption = "No MT tags loaded"
End If
End Sub

Private Sub mnuMTStatus_Click()
'----------------------------------------------
'displays short MT tags statistics, it might
'help with determining problems with MT tags
'----------------------------------------------
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly
Me.MousePointer = vbDefault
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
   txtAlkylationMWCorrection.Text = glALKYLATION
   AlkMWCorrection = glALKYLATION
End If
End Sub

Private Sub txtDBSearchMinimumHighDiscriminantScore_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumHighDiscriminantScore, 0, 1, 0
End Sub

Private Sub txtDBSearchMinimumHighNormalizedScore_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumHighNormalizedScore, 0, 100000, 0
End Sub

Private Sub txtDBSearchMinimumPeptideProphetProbability_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumPeptideProphetProbability, 0, 1, 0
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   samtDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "Molecular Mass Tolerance should be numeric value.", vbOKOnly
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNETFormula_LostFocus()
'------------------------------------------------
'initialize new expression evaluator
'------------------------------------------------
    If Not InitExprEvaluator(txtNETFormula.Text) Then
       MsgBox "Error in elution calculation formula.", vbOKOnly, glFGTU
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
      MsgBox "NET Tolerance should be number between 0 and 1.", vbOKOnly
      txtNETTol.SetFocus
   Else
      samtDef.NETTol = -1   'do not consider NET when searching
   End If
End If
End Sub

Private Sub txtNETTol_Validate(Cancel As Boolean)
    TextBoxLimitNumberLength txtNETTol, 12
End Sub

Private Sub txtResidueToModifyMass_LostFocus()
    ValidateTextboxValueDbl txtResidueToModifyMass, -10000, 10000, 0
End Sub

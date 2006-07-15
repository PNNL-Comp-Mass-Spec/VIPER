VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSearchForNETAdjustmentUMC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Mass Tags Database For NET Adjustment"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOptionFrame 
      Caption         =   "Miscellaneous"
      Height          =   1890
      Index           =   23
      Left            =   4440
      TabIndex        =   42
      Top             =   3000
      Width           =   3615
      Begin VB.TextBox txtNetAdjMinHighDiscriminantScore 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   48
         Text            =   "0.5"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtNetAdjMinHighNormalizedScore 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   46
         Text            =   "2.5"
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chkEliminateConflictingIDs 
         Caption         =   "Do not use peaks pointing to multiple IDs on NET distance of more than"
         Height          =   615
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtMultiIDMaxNETDist 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   44
         Text            =   "0.1"
         Top             =   300
         Width           =   615
      End
      Begin VB.CheckBox chkUseN15AMTMasses 
         Caption         =   "Use N15 AMT Masses"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblDescription 
         Caption         =   "Minimum MT Discriminant Score"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   1245
         Width           =   2505
      End
      Begin VB.Label lblDescription 
         Caption         =   "Minimum MT XCorr"
         Height          =   255
         Index           =   133
         Left            =   120
         TabIndex        =   45
         Top             =   885
         Width           =   1785
      End
   End
   Begin VB.Frame fraPeakSelectionCriteria 
      Caption         =   "Peaks Selection Criteria"
      Height          =   2775
      Left            =   4440
      TabIndex        =   23
      Top             =   120
      Width           =   6015
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Abundance"
         Height          =   1525
         Index           =   21
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3975
         Begin VB.OptionButton optPeakAbuCriteria 
            Caption         =   "Select LAST peak in class"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   2800
         End
         Begin VB.OptionButton optPeakAbuCriteria 
            Caption         =   "Select FIRST peak in class"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   2800
         End
         Begin VB.TextBox txtPctMaxAbu 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3160
            TabIndex        =   31
            Text            =   "10"
            Top             =   780
            Width           =   495
         End
         Begin VB.OptionButton optPeakAbuCriteria 
            Caption         =   "Select AFTER max. abundance"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   2800
         End
         Begin VB.OptionButton optPeakAbuCriteria 
            Caption         =   "Select AT max. abundance"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   2800
         End
         Begin VB.OptionButton optPeakAbuCriteria 
            Caption         =   "Select BEFORE max. abundance"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   2800
         End
         Begin VB.Label lblDescription 
            Caption         =   "Pct. of max. abundance"
            Height          =   375
            Index           =   49
            Left            =   3000
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Charge State"
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   5775
         Begin VB.CheckBox chkCS 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   480
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   35
            Top             =   480
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   36
            Top             =   480
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "4"
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   37
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "5"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   38
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "6"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   39
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   ">=7"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   40
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "any charge state"
            Height          =   375
            Index           =   7
            Left            =   4440
            TabIndex        =   41
            Top             =   410
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Consider peaks with charge states:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   2535
         End
      End
   End
   Begin VB.Frame fraOptionFrame 
      Caption         =   "Net Adj Lockers"
      Height          =   1695
      Index           =   22
      Left            =   8400
      TabIndex        =   50
      Top             =   3000
      Width           =   2055
      Begin VB.CheckBox chkNetAdjUseLockers 
         Caption         =   "Use Lockers"
         Height          =   255
         Left            =   180
         TabIndex        =   51
         Top             =   300
         Width           =   1700
      End
      Begin VB.CheckBox chkNetAdjUseOldIfFailure 
         Caption         =   "Use old algorithm if failure"
         Height          =   375
         Left            =   180
         TabIndex        =   52
         Top             =   600
         Value           =   1  'Checked
         Width           =   1560
      End
      Begin VB.TextBox txtNetAdjMinLockerMatchCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   54
         Text            =   "3"
         Top             =   1160
         Width           =   615
      End
      Begin VB.Label lblDescription 
         Caption         =   "Minimum match count"
         Height          =   360
         Index           =   111
         Left            =   120
         TabIndex        =   53
         Top             =   1095
         Width           =   1065
      End
   End
   Begin VB.Frame fraIte 
      Caption         =   "Iteration"
      Height          =   3015
      Left            =   120
      TabIndex        =   55
      Top             =   4920
      Width           =   8535
      Begin RichTextLib.RichTextBox rtbIteReport 
         Height          =   2250
         Left            =   3960
         TabIndex        =   71
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3969
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmSearchForNETAdjustmentUMC.frx":0000
      End
      Begin VB.CommandButton cmdUseDefaults 
         Caption         =   "Use Defauts"
         Height          =   315
         Left            =   7320
         TabIndex        =   76
         ToolTipText     =   "Start calculations"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtNetAdjMinIDCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   70
         Text            =   "75"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNetAdjAutoIncrementUMCTopAbuPct 
         Caption         =   "Auto-increment high abu UMC's top percent"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   2380
         Width           =   3495
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   315
         Left            =   5700
         TabIndex        =   74
         ToolTipText     =   "Reset to generic formula and NET tolerance of 0.2"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Iterating"
         Height          =   315
         Left            =   4080
         TabIndex        =   72
         ToolTipText     =   "Start calculations"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   315
         Left            =   7320
         TabIndex        =   75
         ToolTipText     =   "Stops iterations and dumps current results"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   4080
         TabIndex        =   73
         ToolTipText     =   "Stops iterations and dumps current results"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when change less than (or Iterations > max)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   60
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox chkAcceptLastIteration 
         Caption         =   "Accept last iteration as NET adjustment"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   2100
         Width           =   3255
      End
      Begin VB.TextBox txtIteNETDec 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   66
         Text            =   "0.025"
         Top             =   1815
         Width           =   615
      End
      Begin VB.TextBox txtIteMWDec 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   64
         Text            =   "2.5"
         Top             =   1485
         Width           =   615
      End
      Begin VB.TextBox txtIteStopVal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   62
         Text            =   "5"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when number of IDs goes under"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   3015
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when NET tol. goes under"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when MW tol. goes under"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop after iteration number"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.CheckBox chkDecNET 
         Caption         =   "Decrease NET tolerance  by"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   1845
         Width           =   2415
      End
      Begin VB.CheckBox chkDecMW 
         Caption         =   "Decrease MW tolerance by"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   1515
         Width           =   2535
      End
      Begin VB.Label lblDescription 
         Caption         =   "Minimum matching UMC count"
         Height          =   255
         Index           =   47
         Left            =   480
         TabIndex        =   69
         Top             =   2670
         Width           =   2445
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "}"
         BeginProperty Font 
            Name            =   "Script"
            Size            =   36
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   3030
         TabIndex        =   61
         Top             =   480
         Width           =   210
      End
   End
   Begin VB.Frame fraUMCSelectionCriteria 
      Caption         =   "UMC Selection Criteria"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cboPairsUMCsToUseForNETAdjustment 
         Height          =   315
         ItemData        =   "frmSearchForNETAdjustmentUMC.frx":0082
         Left            =   120
         List            =   "frmSearchForNETAdjustmentUMC.frx":0084
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   4000
      End
      Begin VB.TextBox txtMaxUMCScansPct 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Text            =   "10"
         Top             =   880
         Width           =   495
      End
      Begin VB.TextBox txtMinScanRange 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Text            =   "3"
         Top             =   560
         Width           =   495
      End
      Begin VB.TextBox txtUMCAbuTopPct 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Text            =   "20"
         Top             =   1200
         Width           =   420
      End
      Begin VB.CheckBox chkUMCUseTopAbu 
         Caption         =   "Use high-abundance UMCs only - Top"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox txtMinUMCCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Text            =   "3"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblUMCsToUseWhenPairsExist 
         Caption         =   "UMC's to use for NET adjustment when pairs exist"
         Height          =   255
         Index           =   102
         Left            =   120
         TabIndex        =   10
         Top             =   1540
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Maximum percentage of total scans in UMC"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   900
         Width           =   3135
      End
      Begin VB.Label Label6 
         Caption         =   "Minimum scan range for UMC"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   580
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Minimum number of peaks in UMC to use"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   260
         Width           =   2895
      End
   End
   Begin VB.Frame fraID 
      Caption         =   "Search (Identification) Options"
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   2460
      Width           =   4215
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   2000
         Width           =   3735
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   20
         Text            =   "0.2"
         Top             =   1300
         Width           =   1215
      End
      Begin VB.CheckBox chkUseNETForID 
         Caption         =   "Use NET criteria with tolerance"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1300
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Frame fraMWTolerance 
         Caption         =   "MW Tolerance"
         Height          =   960
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1935
         Begin VB.TextBox txtMWTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   160
            TabIndex        =   15
            Text            =   "10"
            Top             =   525
            Width           =   735
         End
         Begin VB.OptionButton optTolType 
            Caption         =   "&ppm"
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   16
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optTolType 
            Caption         =   "&Dalton"
            Height          =   255
            Index           =   1
            Left            =   1020
            TabIndex        =   17
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Tolerance"
            Height          =   255
            Left            =   165
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label lblClassMassNote 
         Caption         =   "Note: The defined Class Mass is used for NET adjustment"
         Height          =   615
         Left            =   2280
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "NET Calculation Formula (FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1740
         Width           =   3495
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "quo vadis domine?"
      Height          =   255
      Left            =   4440
      TabIndex        =   79
      Top             =   8040
      Width           =   4215
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   420
      Left            =   1200
      TabIndex        =   78
      Top             =   7980
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   77
      ToolTipText     =   "Status of the Mass Tag database"
      Top             =   7980
      Width           =   1095
   End
   Begin VB.Menu mnuF 
      Caption         =   "&Function"
      Begin VB.Menu mnuFCalculate 
         Caption         =   "C&alculate Once"
      End
      Begin VB.Menu mnuFIteration 
         Caption         =   "Calculate and Iterate"
      End
      Begin VB.Menu mnuFReport 
         Caption         =   "&Report"
      End
      Begin VB.Menu mnuFExportToMTDB 
         Caption         =   "&Export to MTDB"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuMT 
      Caption         =   "&Mass Tags"
      Begin VB.Menu mnuMTLoad 
         Caption         =   "&Load Mass Tags"
      End
      Begin VB.Menu mnuLoadLegacy 
         Caption         =   "Load L&egacy Mass Tags"
      End
      Begin VB.Menu mnuMTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTStatus 
         Caption         =   "&Status"
      End
   End
   Begin VB.Menu mnuE 
      Caption         =   "&Elution Formula"
      Begin VB.Menu mnuEGeneric 
         Caption         =   "Generic"
      End
      Begin VB.Menu mnuETICFit 
         Caption         =   "TIC Fit"
      End
      Begin VB.Menu mnuEGANET 
         Caption         =   "&GANET"
      End
   End
End
Attribute VB_Name = "frmSearchForNETAdjustmentUMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------
'created: 09/04/2002 nt
'last modified: 02/05/2003 nt
'-----------------------------------------------------------------
'NOTE: in this case NET array contains Lars elution times and RT
'      array contains their standard deviation
'NOTE: Search is always on all data points
'NOTE: no data is changed based on this function
'-----------------------------------------------------------------
'This is how it works; find all matching mass tags then look to
'reduce that number by selecting pairs that will yield best match
'(eliminate all except first instance of each peptide; use higher
'intensity peaks). Look for best Slope and Intercept for IDs by
'least square method!
'-----------------------------------------------------------------
Option Explicit

Const MAX_ID_CNT = 100000            'maximum number of IDs

Const STATE_NET_TOO_DISTANT = 1024
Const STATE_TOO_LONG_ELUTION = 2048
Const STATE_BAD_NET = 4096
Const STATE_OUTSCORED = 8192
Const STATE_ID_NETS_TOO_DISTANT = 16384

Const ITERATION_STOP_NUMBER = 0
Const ITERATION_STOP_MW_TOL_LIMIT = 1
Const ITERATION_STOP_NET_TOL_LIMIT = 2
Const ITERATION_STOP_ID_LIMIT = 3
Const ITERATION_STOP_CHANGE = 4

Const NET_RESOLUTION = 0.01                 'theoretical NET resolution; this defines the absolute minimum value that can be used for the NET Tolerance
Const NET_CHANGE_PCT = 0.25                 'used in iteration that stops when change is insignificant

Const CMD_CAPTION_PAUSE = "Pause"
Const CMD_CAPTION_CONTINUE = "Continue"

Dim GRID() As GR            'GRID array is parallel with AMT array
                            'it is indexed the same way and it's
                            'members are indexes in ID belonging to
                            'AMT (in other words have same ID as
                            'Mass Tags index in AMT arrays)

Dim IDCnt As Long           'identification count
Dim ID() As Long            'index in AMT array or UMCNetLockers.Lockers()
Dim IDInd() As Long         'peak index in data arrays
Dim IDType() As Long        'type of peak
Dim IDState() As Long       'used to clean IDs from duplicates and bads
Dim IDScan() As Long        'this is redundant but will make life easier
Dim IDsAreNetAdjLockers As Boolean

Dim UseUMC() As Boolean     'marks which UMCs should be used in this search
Dim UMCCntAddedSinceLowSegmentCount As Long
Dim UMCSegmentCntWithLowUMCCnt As Long

Dim PeakCnt As Long
Dim PeakInd() As Long
Dim PeakType() As Long
Dim PeakScan() As Long
Dim PeakMW() As Double
Dim PeakUMCInd() As Long        ' Index of the UMC that this peak belongs to

Dim IteSlope() As Double        'array of iteration calculated slopes
Dim IteIntercept() As Double    'array of iteration calculated intercepts
Dim IteAvgDev() As Double       'blah, blah, blah

'in this case CallerID is a public property
Public CallerID As Long

Dim bLoading As Boolean

Dim ScanMin As Long                 'first scan number
Dim ScanMax As Long                 'last scan number
Dim ScanRange As Long               'last-first+1

Dim AdjSlp As Double                'slope of GANET adjustments
Dim AdjInt As Double                'intercept of GANET adjustment
Dim AdjAvD As Double

Dim EditGANETSPName As String
Dim NETExprEva As ExprEvaluator     'expression evaluator for NET
Dim VarVals()                       'variable for expression evaluator

Dim bStop As Boolean
Dim bPause As Boolean

Private mUsingDefaultGANET As Boolean
Private mStopChangeTestResult As Double
Private mUsingNetAdjLockers As Boolean
Private mMTMinimumHighNormalizedScore As Single
Private mMTMinimumHighDiscriminantScore As Single

Private mNetAdjLockerIndexPointers() As Long             ' Pointer to entry in UMCNetLockers.Lockers()
Private objNetAdjLockerSearchUtil As MWUtil

Private mUMCClsStats() As Double

Private Type udtSegmentStatsType
    UMCHitCountUsed As Long
    ArrayCountUnused As Long
    UnusedUMCIndices() As Long          ' 0-based array; holds the indices of the Unused UMC's in this segment
End Type
'

Public Sub CalculateIteration()
    Dim strMessage As String
    Dim blnSuccess As Boolean

On Error GoTo CalculateIterationErrorHandler

    Me.MousePointer = vbHourglass
    
    If UMCNetAdjDef.UseNetAdjLockers Then
        blnSuccess = CalculateIterationUsingNetAdjLockers
        If Not blnSuccess Then
            If Not UMCNetAdjDef.UseOldNetAdjIfFailure Then
                strMessage = "NET Adjustment using lockers failed.  Since UseOldNetAdjIfFailure = False, will use the default Slope and Intercept."
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                    MsgBox strMessage, vbExclamation Or vbOKOnly, "NET Adjustment Failure"
                Else
                    AddToAnalysisHistory CallerID, strMessage
                End If
                
                ResetSlopeAndInterceptToDefault
                
                ' Set this to True so that CalculateIterationAutoIncrementTopAbuPct or CalculateIterationWork isn't called
                blnSuccess = True
            End If
        End If
    Else
        blnSuccess = False
    End If
        
    If Not blnSuccess Then
        If cChkBox(chkUMCUseTopAbu) And glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentAutoIncrementUMCTopAbuPct Then
            CalculateIterationAutoIncrementTopAbuPct
        Else
            ' Call CalculateIterationWork() just once
            CalculateIterationWork 0, False
        End If
    End If
    
    If Not GelAnalysis(CallerID) Is Nothing Then
        If GelAnalysis(CallerID).GANET_Slope <= 0 Then
            ' Negative slope was computed
            ' Inform user if not in auto analysis mode, then reset to the default slope and intercept
            If GelAnalysis(CallerID).GANET_Slope < 0 Then
                strMessage = "The computed slope is negative"
            Else
                strMessage = "The computed slope is zero"
            End If
            
            strMessage = strMessage & "; this is not allowed and consequently the default NET formula will be used."
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox strMessage, vbExclamation Or vbOKOnly, "Invalid GANET Slope"
            Else
                AddToAnalysisHistory CallerID, strMessage
            End If
            
            ' The computed slope was zero or negative; reset the slope and intercept to the default values
            ResetSlopeAndInterceptToDefault
            
            ' Need to assign a non-zero value to GANET_Fit; we'll assign 1.11E-3 with all 1's so it stands out
            GelAnalysis(CallerID).GANET_Fit = 1.11111111111111E-03
        End If
    End If
    
    Me.MousePointer = vbDefault
    Exit Sub

CalculateIterationErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.CalculateIteration"
    Me.MousePointer = vbDefault
    
End Sub

Private Sub CalculateIterationAutoIncrementTopAbuPct()
    
    Dim blnDone As Boolean
    Dim blnUseAbbreviatedFormat As Boolean
    Dim lngOriginalNetAdjMinIDCount As Long
    Dim IterationCount As Long
    Dim strNetAdjFailureReason As String
    Dim blnStartedWithGenericNETEquation As Boolean
    
    ' Call CalculateIterationWork()
    ' If number of ID's is less than the minimum, then
    '   increment txtUMCAbuTopPct and try again
    
    mStopChangeTestResult = 1E+308
    If txtNETFormula = ConstructNETFormulaWithDefaults(UMCNetAdjDef) Then
        blnStartedWithGenericNETEquation = True
    ElseIf txtNETFormula = ConstructNETFormula(0, 0, True) Then
        blnStartedWithGenericNETEquation = True
    Else
        blnStartedWithGenericNETEquation = False
    End If
    
    With glbPreferencesExpanded.AutoAnalysisOptions
        If .NETAdjustmentUMCTopAbuPctIncrement < 5 Then .NETAdjustmentUMCTopAbuPctIncrement = 5
        If .NETAdjustmentUMCTopAbuPctMax < 5 Then .NETAdjustmentUMCTopAbuPctMax = 5
    End With
    
    If UMCNetAdjDef.TopAbuPct < 1 Then
        UMCNetAdjDef.TopAbuPct = 1
        txtUMCAbuTopPct = UMCNetAdjDef.TopAbuPct
    End If
    
    ' Note: On the first call to CalculateIterationWork, we want blnUseAbbreviatedFormat = False
    '       On subsequent calls, we want it True
    blnUseAbbreviatedFormat = False
    
    lngOriginalNetAdjMinIDCount = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount
    blnDone = False
    Do
        IterationCount = 0
        CalculateIterationWork IterationCount, blnUseAbbreviatedFormat
        
        blnUseAbbreviatedFormat = True
        
        With glbPreferencesExpanded.AutoAnalysisOptions
            If (IDCnt < .NETAdjustmentMinIDCount Or IterationCount < .NETAdjustmentMinIterationCount Or AdjSlp <= 0) And Not bStop Then
                    
                blnDone = False
                If IDCnt < .NETAdjustmentMinIDCount Then
                    ' Not enough ID's
                    strNetAdjFailureReason = "Not enough UMC's matched mass tags in the database"
                ElseIf AdjSlp <= 0 Then
                    ' Zero or Negative slope
                    strNetAdjFailureReason = "Negative slope was computed"
                Else
                    ' Not enough iterations
                    ' Check for some situations where we don't need to auto-increment
                    If Not blnStartedWithGenericNETEquation And mStopChangeTestResult <= UMCNetAdjDef.IterationStopValue ^ 2 Then
                        ' We didn't start with the generic formula, and the change in NetMin and NetMax
                        '   was small, so this isn't actually a problem since we do have enough ID's
                        blnDone = True
                    ElseIf UMCNetAdjDef.IterationStopType = ITERATION_STOP_NUMBER And IterationCount >= UMCNetAdjDef.IterationStopValue Then
                        ' The stop type was a fixed number of iterations, and we reached the desired number
                        '  of iterations, so this isn't a problem
                        blnDone = True
                    Else
                        strNetAdjFailureReason = "Number of iterations was lower than the required minimum"
                    End If
                End If
                
                If Not blnDone Then
                    If UMCNetAdjDef.TopAbuPct < .NETAdjustmentUMCTopAbuPctMax Then
                        ' Increment TopAbuPct
                        UMCNetAdjDef.TopAbuPct = UMCNetAdjDef.TopAbuPct + .NETAdjustmentUMCTopAbuPctIncrement
                        If UMCNetAdjDef.TopAbuPct > 100 Then UMCNetAdjDef.TopAbuPct = 100
                        txtUMCAbuTopPct = Trim(UMCNetAdjDef.TopAbuPct)
                        
                        AddToAnalysisHistory CallerID, "NET Adjustment: " & strNetAdjFailureReason & "; incremented high abundance UMC's top percent and will repeat search; increment = " & Trim(.NETAdjustmentUMCTopAbuPctIncrement) & "%; new top abundance percent = " & Trim(UMCNetAdjDef.TopAbuPct) & "%"
                        
                        blnDone = False
                    Else
                        ' We're at the TopAbuPct value specifed by .NETAdjustmentUMCTopAbuPctMax
                        ' Possibly temporarily decrement .NETAdjustmentMinIDCount and repeat
                        If .NETAdjustmentMinIDCount > .NETAdjustmentMinIDCountAbsoluteMinimum Then
                            .NETAdjustmentMinIDCount = .NETAdjustmentMinIDCountAbsoluteMinimum
                            
                            AddToAnalysisHistory CallerID, "NET Adjustment: " & strNetAdjFailureReason & "; decreased minimum ID count (since UMC top percent is already at the maximum) and will repeat search; top abundance percent = " & Trim(UMCNetAdjDef.TopAbuPct) & "%; new minimum ID count = " & Trim(.NETAdjustmentMinIDCount)
                            
                            blnDone = False
                        Else
                            blnDone = True
                        End If
                        
                    End If
                End If
                
                If Not blnDone Then
                    ' Reset .NetTol
                    UMCNetAdjDef.NETTol = .NETAdjustmentInitialNetTol
                    txtNETTol = .NETAdjustmentInitialNetTol
                    
                    ' Reset the NET Formula to the Default
                    txtNETFormula.Text = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
                    ValidateNETFormula
                    
                    DoEvents
                End If
            Else
                blnDone = True
            End If
        End With
    Loop While Not blnDone
    
    ' .NETAdjustmentMinIDCount may have been changed above; restore to the original value
    glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount = lngOriginalNetAdjMinIDCount

End Sub

Private Function CalculateIterationUsingNetAdjLockers() As Boolean
    ' Use the Net Adjustment lockers to determine a slope and intercept
    ' Returns True if success, False if error or GANET slope and intercept could not be determined
    
    Dim blnValidSettings As Boolean
    Dim blnSuccess As Boolean
    
    Dim intNETAdjustmentMinIDCountSaved As Integer
    Dim intNETAdjustmentMinIDCountAbsoluteMinimumSaved As Integer
    Dim intNETAdjustmentMinIterationCountSaved As Integer
    
On Error GoTo CalculateIterationUsingNetAdjLockersErrorHandler

    blnSuccess = False
    
    ' First make sure the Net Adjustment lockers and settings make sense
    blnValidSettings = True
    With UMCNetAdjDef
        If UMCNetLockers.LockerCount < .NetAdjLockerMinimumMatchCount Then
            ' Not enough lockers to satisfy .NetAdjLockerMinimumMatchCount
            ' Cannot use net adjustment lockers
            blnValidSettings = False
        End If
    End With
        
    If blnValidSettings Then
        ' When using Net Lockers, we need to modify some of the parameters
        '   in UMCNetAdjDef and glbPreferencesExpanded.AutoAnalysisOptions
        
        With glbPreferencesExpanded.AutoAnalysisOptions
            intNETAdjustmentMinIDCountSaved = .NETAdjustmentMinIDCount
            intNETAdjustmentMinIDCountAbsoluteMinimumSaved = .NETAdjustmentMinIDCountAbsoluteMinimum
            intNETAdjustmentMinIterationCountSaved = .NETAdjustmentMinIterationCount
            
            .NETAdjustmentMinIDCount = UMCNetLockers.LockerCount                                     ' Desired match count
            .NETAdjustmentMinIDCountAbsoluteMinimum = UMCNetAdjDef.NetAdjLockerMinimumMatchCount        ' Absolute minimum match count
            .NETAdjustmentMinIterationCount = 2
        End With
        
        If Not InitializeNetAdjLockerSearch() Then
            blnSuccess = False
        Else
            mUsingNetAdjLockers = True
            CalculateIterationAutoIncrementTopAbuPct
            mUsingNetAdjLockers = False
            
            ' Need to restore some of the settings changed above
            With glbPreferencesExpanded.AutoAnalysisOptions
                .NETAdjustmentMinIDCount = intNETAdjustmentMinIDCountSaved
                .NETAdjustmentMinIDCountAbsoluteMinimum = intNETAdjustmentMinIDCountAbsoluteMinimumSaved
                .NETAdjustmentMinIterationCount = intNETAdjustmentMinIterationCountSaved
            End With
        
            blnSuccess = True
        
        End If
        
    End If
    
    CalculateIterationUsingNetAdjLockers = blnSuccess
    Exit Function

CalculateIterationUsingNetAdjLockersErrorHandler:
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.CalculateIterationUsingNetAdjLockers"
    mUsingNetAdjLockers = False
    CalculateIterationUsingNetAdjLockers = False
End Function

Private Sub CalculateIterationWork(Optional ByRef IterationStep As Long, Optional blnUseAbbreviatedFormat As Boolean = False)
'--------------------------------------------
'calculate elution adjustment in an iteration
'--------------------------------------------
Dim bDone As Boolean
Dim NETTolVarChange As Double
Dim CurrNET1 As Double, CurrNET2 As Double      'NET for first and last point calculated with last slope,int.
Dim PrevNET1 As Double, PrevNET2 As Double      'this is used as stop criteria for iteration of type 4
Dim blnProceed As Boolean
Dim i As Long

On Error GoTo CalculateIterationWorkErrorHandler

Erase IteSlope
Erase IteIntercept
Erase IteAvgDev
rtbIteReport.Text = ""
PrevNET1 = glHugeOverExp:       PrevNET2 = -glHugeOverExp
cmdPause.Caption = CMD_CAPTION_PAUSE
cmdStart.Visible = False
cmdReset.Visible = False
cmdUseDefaults.Visible = False

If mUsingNetAdjLockers Then
    mMTMinimumHighNormalizedScore = 0
    mMTMinimumHighDiscriminantScore = 0
Else
    mMTMinimumHighNormalizedScore = glbPreferencesExpanded.NetAdjustmentMinHighNormalizedScore
    mMTMinimumHighDiscriminantScore = glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore
    
    If mMTMinimumHighDiscriminantScore > 0 Then
        ' Make sure at least two of the loaded mass tags have score values >= mMTMinimumHighDiscriminantScore, also taking into account HiNormalizedScore
        ValidateMTMinimimumHighDiscriminantScore AMTHiDiscriminantScore(), AMTHiNormalizedScore(), 1, AMTCnt, mMTMinimumHighDiscriminantScore, mMTMinimumHighNormalizedScore, 2
    Else
        ' Make sure at least two of the loaded mass tags have score values >= mMTMinimumHighNormalizedScore
        ValidateMTMinimimumHighNormalizedScore AMTHiNormalizedScore(), 1, AMTCnt, mMTMinimumHighNormalizedScore, 2
    End If
    
End If

bPause = False
bStop = False
Do Until bDone
   IterationStep = IterationStep + 1
   UpdateStatus "Iterating NET adjustment; step " & IterationStep
   ReDim Preserve IteSlope(IterationStep - 1)
   ReDim Preserve IteIntercept(IterationStep - 1)
   ReDim Preserve IteAvgDev(IterationStep - 1)
   
   blnProceed = False
   If IterationStep = 1 Then
      ' On the first iteration, determine which UMC's and which Peaks are to be used
      Call ResetProcedure
      If SelectUMCToUse() > 0 Then
         If SelectPeaksToUse() > 0 Then
            blnProceed = True
         End If
      End If
   Else
      ' On subsequent iterations, simply clear the ID arrays and reset the Hit bit
      ClearIDArrays
      For i = 0 To GelUMC(CallerID).UMCCnt - 1
         With GelUMC(CallerID).UMCs(i)
            .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
         End With
      Next i
      blnProceed = True
   End If
   
   If blnProceed Then
      If mUsingNetAdjLockers Then
           If UMCNetLockers.LockerCount < 1 Then GoTo CalculateIterationWorkExitSub
           If Not SearchNetLockers() Then
              cmdStart.Visible = True
              cmdReset.Visible = True
              cmdUseDefaults.Visible = True
              GoTo CalculateIterationWorkExitSub
           End If
      Else
        If AMTCnt < 1 Then GoTo CalculateIterationWorkExitSub
        If UMCNetAdjDef.UseNET Then
           If Not SearchMassTagsMWNET() Then
              cmdStart.Visible = True
              cmdReset.Visible = True
              cmdUseDefaults.Visible = True
              GoTo CalculateIterationWorkExitSub
           End If
        Else
           If Not SearchMassTagsMW() Then
              cmdStart.Visible = True
              cmdReset.Visible = True
              cmdUseDefaults.Visible = True
              GoTo CalculateIterationWorkExitSub
           End If
        End If
      End If
      
      If FillTheGRID() Then
         If SelectIdentifications() > 1 Then Call CalculateSlopeIntercept
      End If
   End If
   
   IteSlope(IterationStep - 1) = AdjSlp
   IteIntercept(IterationStep - 1) = AdjInt
   IteAvgDev(IterationStep - 1) = AdjAvD
   If AdjAvD < 0 Then
      bDone = True              'can not continue
   Else                         'remember this iteration as the last so far
      If UMCNetAdjDef.IterationAcceptLast Then
         If Not GelAnalysis(CallerID) Is Nothing Then
            GelAnalysis(CallerID).GANET_Fit = AdjAvD
            GelAnalysis(CallerID).GANET_Slope = AdjSlp
            GelAnalysis(CallerID).GANET_Intercept = AdjInt

            ' MonroeMod
            GelStatus(CallerID).Dirty = True
         End If
         
         UMCNetAdjDef.NETFormula = ConstructNETFormula(AdjSlp, AdjInt)
      End If
   End If
   ReportIterationStep IterationStep
   'see if user is sick and tired of iterating
   DoEvents
   If bPause Then ReportPause
   Do While bPause      'loop if paused
      DoEvents
      If bStop Then bPause = False
      ' Sleep for 100 msec to reduce processor usage
      Sleep 100
   Loop
   If bStop Then
      ReportStop
      bDone = True
   End If
   If UMCNetAdjDef.IterationUseMWDec Then
      UMCNetAdjDef.MWTol = UMCNetAdjDef.MWTol - UMCNetAdjDef.IterationMWDec
      txtMWTol.Text = UMCNetAdjDef.MWTol
   End If
   If UMCNetAdjDef.IterationStopType = ITERATION_STOP_CHANGE Then
      NETTolVarChange = UMCNetAdjDef.NETTol * NET_CHANGE_PCT
      UMCNetAdjDef.NETTol = UMCNetAdjDef.NETTol - NETTolVarChange
      If UMCNetAdjDef.NETTol < NET_RESOLUTION Then UMCNetAdjDef.NETTol = NET_RESOLUTION
      txtNETTol.Text = UMCNetAdjDef.NETTol
      txtNETTol_Validate False
   Else                 'constant change;make sure you don't go below NET resolution limits
      If UMCNetAdjDef.IterationUseNETdec Then
         UMCNetAdjDef.NETTol = UMCNetAdjDef.NETTol - UMCNetAdjDef.IterationNETDec
         If UMCNetAdjDef.NETTol < NET_RESOLUTION Then UMCNetAdjDef.NETTol = NET_RESOLUTION
         txtNETTol.Text = UMCNetAdjDef.NETTol
         txtNETTol_Validate False
      End If
   End If
   
   Select Case UMCNetAdjDef.IterationStopType
   Case ITERATION_STOP_NUMBER
        If IterationStep >= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_MW_TOL_LIMIT
        If UMCNetAdjDef.MWTol <= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_NET_TOL_LIMIT
        If UMCNetAdjDef.NETTol <= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_ID_LIMIT
        If IDCnt <= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_CHANGE
        CurrNET1 = AdjSlp * ScanMin + AdjInt
        CurrNET2 = AdjSlp * ScanMax + AdjInt
        '' Debug.Print IDCnt & ", " & CurrNET1 & ", " & CurrNET2 & ", " & (CurrNET1 - PrevNET1) ^ 2 + (CurrNET2 - PrevNET2) ^ 2
        mStopChangeTestResult = (CurrNET1 - PrevNET1) ^ 2 + (CurrNET2 - PrevNET2) ^ 2
        If mStopChangeTestResult <= UMCNetAdjDef.IterationStopValue ^ 2 Then
           ReportIterationStep_Change PrevNET1, PrevNET2, CurrNET1, CurrNET2
           bDone = True
        ElseIf IDCnt < 2 Or IterationStep > glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMaxIterationCount Then
           ReportIterationStep_Change PrevNET1, PrevNET2, CurrNET1, CurrNET2
           bDone = True
        Else
           PrevNET1 = CurrNET1:     PrevNET2 = CurrNET2
        End If
   End Select
   
   If Not bDone And glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        ' Auto Analyzing: need to check if max Iteration Steps reached or IDCnt too low
        ' The IterationStopType should be defined as ITERATION_STOP_CHANGE so that that criterion will be checked above
        Debug.Assert UMCNetAdjDef.IterationStopType = ITERATION_STOP_CHANGE
        With glbPreferencesExpanded.AutoAnalysisOptions
            If IterationStep >= .NETAdjustmentMaxIterationCount Then bDone = True
            If IDCnt < .NETAdjustmentMinIDCount Then bDone = True
        End With
   End If
   
   'don't change NET formula if done
   If Not bDone Then
      UMCNetAdjDef.NETFormula = ConstructNETFormula(AdjSlp, AdjInt)
      txtNETFormula.Text = UMCNetAdjDef.NETFormula
      CheckNETEquationStatus True
      
      If Not InitExprEvaluator(txtNETFormula.Text) Then
         UpdateStatus "Error in elution calculation formula!"
         bDone = True
      End If
   Else
      UpdateStatus "Ready"
   End If
Loop

UpdateAnalysisHistory IterationStep, AdjSlp, AdjInt, AdjAvD, blnUseAbbreviatedFormat

CalculateIterationWorkExitSub:

cmdStart.Visible = True
cmdReset.Visible = True
cmdUseDefaults.Visible = True

Exit Sub

CalculateIterationWorkErrorHandler:
LogErrors Err.Number, "CalculateIterationWork"
Debug.Assert False
' Using Resume Next error handling if the error is a divide by zero error
If Err.Number = 11 Then Resume Next

End Sub

Private Sub DisplayDefaultSettings()

    On Error GoTo DisplayDefaultSettingsErrorHandler
    
    SetDefaultUMCNETAdjDef GelUMCNETAdjDef(CallerID)
    
    ResetExpandedPreferences glbPreferencesExpanded, "NetAdjustmentOptions"
    ResetExpandedPreferences glbPreferencesExpanded, "NetAdjustmentUMCDistributionOptions"
    
    With glbPreferencesExpanded
        .NetAdjustmentUsesN15AMTMasses = False
        .NetAdjustmentMinHighNormalizedScore = 2.5
        .NetAdjustmentMinHighDiscriminantScore = 0.5
        .PairSearchOptions.NETAdjustmentPairedSearchUMCSelection = punaUnpairedPlusPairedLight
    End With
    
    InitializeForm True
    
    Exit Sub

DisplayDefaultSettingsErrorHandler:
    Debug.Assert False
    
End Sub

Private Sub InitializeForm(Optional blnUpdateControlsOnly As Boolean = False)
    Dim i As Long
    
    On Error GoTo InitializeFormErrorHandler
    
    If Not blnUpdateControlsOnly Then
        bLoading = True
        If IsWinLoaded(TrackerCaption) Then Unload frmTracker
    End If
    
    ' MonroeMod
    If CallerID >= 1 And CallerID <= UBound(GelBody) Then UMCNetAdjDef = GelUMCNETAdjDef(CallerID)
    
    With UMCNetAdjDef
        txtMWTol.Text = .MWTol
        Select Case .MWTolType
        Case gltPPM
             optTolType(0).value = True
        Case gltABS
             optTolType(1).value = True
        End Select
        txtNETTol.Text = .NETTol
        txtMinUMCCount.Text = .MinUMCCount
        txtMinScanRange.Text = .MinScanRange
        txtMaxUMCScansPct.Text = .MaxScanPct
        If .TopAbuPct >= 0 Then
           chkUMCUseTopAbu.value = vbChecked
           If glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentAutoIncrementUMCTopAbuPct Then
               txtUMCAbuTopPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
               .TopAbuPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
           Else
               txtUMCAbuTopPct = .TopAbuPct
           End If
        Else
           txtUMCAbuTopPct.Text = -.TopAbuPct
           chkUMCUseTopAbu.value = vbUnchecked
        End If
        optPeakAbuCriteria(.PeakSelection).value = True
        For i = 0 To UBound(.PeakCSSelection)
            If .PeakCSSelection(i) Then
               chkCS(i).value = vbChecked
            Else
               chkCS(i).value = vbUnchecked
            End If
        Next i
        SetCheckBox chkEliminateConflictingIDs, .UseMultiIDMaxNETDist
        txtMultiIDMaxNETDist.Text = .MultiIDMaxNETDist
        txtNETFormula.Text = .NETFormula
        ValidateNETFormula
        
        SetCheckBox chkUseNETForID, .UseNET
        optIteStop(.IterationStopType).value = True
        txtIteStopVal.Text = .IterationStopValue
        
        SetCheckBox chkDecMW, .IterationUseMWDec
        txtIteMWDec.Text = .IterationMWDec
        
        SetCheckBox chkDecNET, .IterationUseNETdec
        txtIteNETDec.Text = .IterationNETDec
        
        SetCheckBox chkAcceptLastIteration, .IterationAcceptLast
        SetCheckBox chkNetAdjUseLockers, .UseNetAdjLockers
        SetCheckBox chkNetAdjUseOldIfFailure, .UseOldNetAdjIfFailure
        txtNetAdjMinLockerMatchCount = .NetAdjLockerMinimumMatchCount
        
    End With
    
    With cboPairsUMCsToUseForNETAdjustment
        .Clear
        .AddItem "All UMC's, regardless of pair or light/heavy status", punaPairedAndUnpaired
        .AddItem "Unpaired UMC's only", punaUnpairedOnly
        .AddItem "Unpaired UMC's and light members of paired UMC's", punaUnpairedPlusPairedLight
        .AddItem "Paired UMC's, both light and heavy members", punaPairedAll
        .AddItem "Paired UMC's, light members only", punaPairedLight
        .AddItem "Paired UMC's, heavy members only", punaPairedHeavy
        .ListIndex = glbPreferencesExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection
        If .ListIndex < 0 Then
            Debug.Assert False
            .ListIndex = punaUnpairedPlusPairedLight
        End If
    End With
    
    With glbPreferencesExpanded.AutoAnalysisOptions
        ' Note: if .NETAdjustmentAutoIncrementUMCTopAbuPct = True, this will automatically
        '       update txtUMCAbuTopPct and set optIteStop(ITERATION_STOP_CHANGE) = True
        SetCheckBox chkNetAdjAutoIncrementUMCTopAbuPct, .NETAdjustmentAutoIncrementUMCTopAbuPct
        txtNetAdjMinIDCount = Trim(.NETAdjustmentMinIDCount)
    End With
    
    With glbPreferencesExpanded
        SetCheckBox chkUseN15AMTMasses, .NetAdjustmentUsesN15AMTMasses
        txtNetAdjMinHighNormalizedScore = .NetAdjustmentMinHighNormalizedScore
        txtNetAdjMinHighDiscriminantScore = .NetAdjustmentMinHighDiscriminantScore
    End With
    
    DoEvents

    Exit Sub

InitializeFormErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.InitiationForm"
    Resume Next
End Sub

Public Sub InitializeNETAdjustment()
'------------------------------------------------------------
'Load mass tags database data if necessary
'If CallerID is associated with mass tags database load that
' database if necessary; if CallerID is not associated with
' mass tags database load legacy database
'Initialize mUMCClsStats()
'------------------------------------------------------------

Dim lngAllUMCCount As Long
Dim eResponse As VbMsgBoxResult

On Error GoTo InitializeNETAdjustmentErrorHandler
UpdateStatus ""
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
   
   GetScanRange CallerID, ScanMin, ScanMax, ScanRange
   If ScanRange < 2 Then
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Scan range for this display could lead to unpredictable results!", vbOKOnly, glFGTU
      Else
        AddToAnalysisHistory CallerID, "Warning, total number of scans is less than 2; NET adjustment cannot be performed"
      End If
   End If
   
   lngAllUMCCount = UMCStatistics1(CallerID, mUMCClsStats())
   Debug.Assert lngAllUMCCount = GelUMC(CallerID).UMCCnt
      
   bLoading = False
   EditGANETSPName = glbPreferencesExpanded.MTSConnectionInfo.spEditGANET
   If Len(EditGANETSPName) > 0 Then         'enable export if neccessary
      mnuFExportToMTDB.Enabled = True
   Else
      mnuFExportToMTDB.Enabled = False
   End If
   If Not InitExprEvaluator(UMCNetAdjDef.NETFormula) Then
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in elution calculation formula!", vbOKOnly, glFGTU
      Else
        AddToAnalysisHistory CallerID, "Error in elution calculation formula: " & UMCNetAdjDef.NETFormula
      End If
      txtNETFormula.SetFocus
   End If
End If
Me.MousePointer = vbDefault
Exit Sub

InitializeNETAdjustmentErrorHandler:
LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.InitializeNETAdjustment"
Resume Next

End Sub
                                                                        
Private Function InitializeNetAdjLockerSearch() As Boolean
    ' Returns True if success, False if error
    
    Dim dblSearchMasses() As Double
    Dim lngIndex As Long
    
    Dim objQSDouble As New QSDouble
    
On Error GoTo InitializeNetAdjLockerSearchErrorHandler

    Set objNetAdjLockerSearchUtil = New MWUtil
    
    With UMCNetLockers
        If .LockerCount > 0 Then
            ReDim dblSearchMasses(0 To .LockerCount - 1)
            ReDim mNetAdjLockerIndexPointers(0 To .LockerCount - 1)
            
            For lngIndex = 0 To .LockerCount - 1
                dblSearchMasses(lngIndex) = .Lockers(lngIndex).MonoisotopicMass
                mNetAdjLockerIndexPointers(lngIndex) = lngIndex
            Next lngIndex
        Else
            ReDim dblSearchMasses(0)
            ReDim mNetAdjLockerIndexPointers(0)
        End If
    End With
            
    ' Need to sort dblSearchMasses before calling objNetAdjLockerSearchUtil.Fill()
    ' Use objQSLong for this
    If UMCNetLockers.LockerCount > 0 Then
        If objQSDouble.QSAsc(dblSearchMasses, mNetAdjLockerIndexPointers) Then
            InitializeNetAdjLockerSearch = objNetAdjLockerSearchUtil.Fill(dblSearchMasses())
        Else
            ' Error sorting
            Debug.Assert False
            InitializeNetAdjLockerSearch = False
        End If
    Else
        InitializeNetAdjLockerSearch = False
    End If
    
    Exit Function
    
InitializeNetAdjLockerSearchErrorHandler:
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.InitializeNetAdjLockerSearch"
    Set objNetAdjLockerSearchUtil = Nothing
    InitializeNetAdjLockerSearch = False
    
End Function

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

Private Sub ValidateNETFormula()
    If Not InitExprEvaluator(txtNETFormula.Text) Then
       If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox "Error in elution calculation formula!", vbOKOnly, glFGTU
       End If
       txtNETFormula = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
       txtNETFormula.SetFocus
    Else
       UMCNetAdjDef.NETFormula = txtNETFormula.Text
       CheckNETEquationStatus
    End If
End Sub

Private Sub ResetSlopeAndInterceptToDefault()
    With GelAnalysis(CallerID)
        .GANET_Slope = UMCNetAdjDef.InitialSlope
        .GANET_Intercept = UMCNetAdjDef.InitialIntercept
    End With
    
    txtNETFormula.Text = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
    ValidateNETFormula
End Sub

Public Sub ResetToGenericNetAdjSettings()
    txtNETFormula.Text = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
    ValidateNETFormula
        
    txtNETTol = Trim(glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol)
    UMCNetAdjDef.NETTol = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol
    
    ' If auto-adjusting UMC Top Abu Pct, then set txtUMCAbuTopPct to the value defined by NETAdjustmentUMCTopAbuPctInitial
    If glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentAutoIncrementUMCTopAbuPct Then
        txtUMCAbuTopPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
        UMCNetAdjDef.TopAbuPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
    Else
        txtUMCAbuTopPct = UMCNetAdjDef.TopAbuPct
    End If
    
End Sub

Private Function SearchMassTagsMW() As Boolean
'----------------------------------------------------------
'searches mass tags for matching masses and returns True if
'OK, False if any error or user canceled the whole thing
'----------------------------------------------------------
Dim i As Long, j As Long
Dim eResponse As VbMsgBoxResult
Dim TmpCnt As Long
Dim Hits() As Long
Dim MWAbsErr As Double
On Error GoTo err_SearchMassTagsMW

UpdateStatus "Searching for mass tags ..."
' These arrays are dimensioned to hold MAX_ID_CNT items
' If too many identifications are found, the Error Number 9 will be raised by ??
ReDim ID(MAX_ID_CNT - 1)            'prepare for the worst case
ReDim IDInd(MAX_ID_CNT - 1)
ReDim IDType(MAX_ID_CNT - 1)
ReDim IDScan(MAX_ID_CNT - 1)
ReDim IDState(MAX_ID_CNT - 1)
IDsAreNetAdjLockers = False

For i = 0 To PeakCnt - 1
    Select Case UMCNetAdjDef.MWTolType
    Case gltPPM
        MWAbsErr = PeakMW(i) * UMCNetAdjDef.MWTol * glPPM
    Case gltABS
        MWAbsErr = UMCNetAdjDef.MWTol
    Case Else
        Debug.Assert False
    End Select
    Select Case UMCNetAdjDef.NETorRT
    Case glAMT_NET
        TmpCnt = GetMTHits1(PeakMW(i), MWAbsErr, -1, -1, Hits())
    Case glAMT_RT
        TmpCnt = GetMTHits2(PeakMW(i), MWAbsErr, -1, -1, Hits())
    End Select
    If TmpCnt > 0 Then
       For j = 0 To TmpCnt - 1
           IDType(IDCnt) = PeakType(i)
           IDInd(IDCnt) = PeakInd(i)
           ID(IDCnt) = Hits(j)
           IDState(IDCnt) = 0
           IDScan(IDCnt) = PeakScan(i)
           IDCnt = IDCnt + 1   'if reaches limit we will have correct
       Next j                  'results by doing increase at the end
       With GelUMC(CallerID).UMCs(PeakUMCInd(i))
          .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
       End With
    End If
Next i

If IDCnt > 0 Then
   ReDim Preserve ID(IDCnt - 1)
   ReDim Preserve IDInd(IDCnt - 1)
   ReDim Preserve IDType(IDCnt - 1)
   ReDim Preserve IDScan(IDCnt - 1)
   ReDim Preserve IDState(IDCnt - 1)
Else
   Call ClearIDArrays
End If
UpdateStatus "Possible identifications: " & IDCnt
SearchMassTagsMW = True
Exit Function

err_SearchMassTagsMW:
Select Case Err.Number
Case 9      'too many identifications
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Too many possible identifications detected!  " _
                   & "To proceed with the first " & MAX_ID_CNT & _
                   " identifications select OK!", vbOKCancel, glFGTU)
    Else
        eResponse = vbOK
    End If
    
    Select Case eResponse
    Case vbOK
        UpdateStatus "Possible identifications: " & IDCnt
        SearchMassTagsMW = True
    Case Else
        UpdateStatus ""
        Call ClearIDArrays
    End Select
Case 7      'short on memory; try to recover by releasing arrays
    Call ClearIDArrays
    UpdateStatus ""
    MsgBox "System low on memory. Process aborted in recovery attempt!", vbOKOnly, glFGTU
Case Else
    UpdateStatus "Error searching for mass tags!"
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.SearchMassTagsMW"
End Select
End Function


Private Function SearchMassTagsMWNET() As Boolean
'----------------------------------------------------------
'searches mass tags for matching masses and returns True if
'OK, False if any error or user canceled the whole thing
'----------------------------------------------------------
Dim i As Long, j As Long
Dim eResponse As VbMsgBoxResult
Dim TmpCnt As Long
Dim Hits() As Long
Dim MWAbsErr As Double
Dim AMTNETMin As Double, AMTNETMax As Double
Dim blnAddMassTag As Boolean

On Error GoTo err_SearchMassTagsMWNET

UpdateStatus "Searching for mass tags ..."
ReDim ID(MAX_ID_CNT - 1)            'prepare for the worst case
ReDim IDInd(MAX_ID_CNT - 1)
ReDim IDType(MAX_ID_CNT - 1)
ReDim IDScan(MAX_ID_CNT - 1)
ReDim IDState(MAX_ID_CNT - 1)
IDsAreNetAdjLockers = False

For i = 0 To PeakCnt - 1
    Select Case UMCNetAdjDef.MWTolType
    Case gltPPM
        MWAbsErr = PeakMW(i) * UMCNetAdjDef.MWTol * glPPM
    Case gltABS
        MWAbsErr = UMCNetAdjDef.MWTol
    End Select
    Select Case UMCNetAdjDef.NETorRT
    Case glAMT_NET
        TmpCnt = GetMTHits1(PeakMW(i), MWAbsErr, ConvertScanToNET(PeakScan(i)), UMCNetAdjDef.NETTol, Hits())
    Case glAMT_RT
        TmpCnt = GetMTHits2(PeakMW(i), MWAbsErr, ConvertScanToNET(PeakScan(i)), UMCNetAdjDef.NETTol, Hits())
    End Select
    If TmpCnt > 0 Then
    
       ' MonroeMod: The following implements the option "Do not use peaks pointing to multiple IDs on NET distance of more than ..."
       If UMCNetAdjDef.UseMultiIDMaxNETDist And TmpCnt > 1 Then
          ' Examine the NET values for the Hits
          ' If the range of NET values is > .MultiIDMaxNETDist then do not use this match
          AMTNETMin = AMTNET(Hits(0))
          AMTNETMax = AMTNET(Hits(0))
          For j = 1 To TmpCnt - 1
             If AMTNET(j) < AMTNETMin Then AMTNETMin = AMTNET(j)
             If AMTNET(j) > AMTNETMax Then AMTNETMax = AMTNET(j)
          Next j
          
          If Abs(AMTNETMax - AMTNETMin) > UMCNetAdjDef.MultiIDMaxNETDist Then
            TmpCnt = 0
          End If
       End If
       
       If TmpCnt > 0 Then
          For j = 0 To TmpCnt - 1
            If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Then
                If AMTHiNormalizedScore(Hits(j)) >= mMTMinimumHighNormalizedScore And AMTHiDiscriminantScore(Hits(j)) >= mMTMinimumHighDiscriminantScore Then
                    blnAddMassTag = True
                Else
                    blnAddMassTag = False
                End If
            Else
                blnAddMassTag = True
            End If
            
            If blnAddMassTag Then
                IDType(IDCnt) = PeakType(i)
                IDInd(IDCnt) = PeakInd(i)
                ID(IDCnt) = Hits(j)
                IDState(IDCnt) = 0
                IDScan(IDCnt) = PeakScan(i)
                IDCnt = IDCnt + 1   'if reaches limit we will have correct results by doing increase at the end
            End If
          Next j
          With GelUMC(CallerID).UMCs(PeakUMCInd(i))
           .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
          End With
       End If
    End If
Next i

If IDCnt > 0 Then
   ReDim Preserve ID(IDCnt - 1)
   ReDim Preserve IDInd(IDCnt - 1)
   ReDim Preserve IDType(IDCnt - 1)
   ReDim Preserve IDScan(IDCnt - 1)
   ReDim Preserve IDState(IDCnt - 1)
Else
   Call ClearIDArrays
End If
UpdateStatus "Possible identifications: " & IDCnt
SearchMassTagsMWNET = True
Exit Function


err_SearchMassTagsMWNET:
Select Case Err.Number
Case 9      'too many identifications
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Too many possible identifications detected!  " _
                   & "To proceed with the first " & MAX_ID_CNT & _
                   " identifications select OK!", vbOKCancel, glFGTU)
    Else
        eResponse = vbOK
    End If
    
    Select Case eResponse
    Case vbOK
        UpdateStatus "Possible identifications: " & IDCnt
        SearchMassTagsMWNET = True
    Case Else
        UpdateStatus ""
        Call ClearIDArrays
    End Select
Case 7      'short on memory; try to recover by releasing arrays
    Call ClearIDArrays
    UpdateStatus ""
    MsgBox "System low on memory. Process aborted in recovery attempt!", vbOKOnly, glFGTU
Case Else
    UpdateStatus "Error searching for mass tags!"
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.SearchMassTagsMWNET"
End Select
End Function

Private Function SearchNetLockers() As Boolean
'----------------------------------------------------------
'Searches GANET Lockers for matching masses
'Returns True if success, or False if an error
'----------------------------------------------------------
    Dim i As Long, j As Long
    Dim Ind1 As Long, Ind2 As Long
    
    Dim lngHitCount As Long
    Dim lngHitCountDimmed As Long
    Dim Hits() As Long
    
    Dim MWAbsErr As Double
    Dim dblNetToMatch As Double
    
    Dim blnValidHit As Boolean
    
    Dim dblNETMin As Double, dblNETMax As Double
    
    On Error GoTo SearchNetAdjLockersErrorHandler
    
    UpdateStatus "Searching for matching Net Lockers ..."
    ReDim ID(PeakCnt)                           ' Reserve space for highest possible number of matches
    ReDim IDInd(PeakCnt)
    ReDim IDType(PeakCnt)
    ReDim IDScan(PeakCnt)
    ReDim IDState(PeakCnt)
    IDsAreNetAdjLockers = True

    For i = 0 To PeakCnt - 1
        Select Case UMCNetAdjDef.MWTolType
        Case gltPPM
            MWAbsErr = PeakMW(i) * UMCNetAdjDef.MWTol * glPPM
        Case gltABS
            MWAbsErr = UMCNetAdjDef.MWTol
        End Select
        
        dblNetToMatch = ConvertScanToNET(PeakScan(i))
        
        lngHitCount = 0
        
        ' Note: objNetAdjLockerSearchUtil should have been filled with the Net Locker masses, sorted ascending
        If objNetAdjLockerSearchUtil.FindIndexRange(PeakMW(i), MWAbsErr, Ind1, Ind2) Then
            If Ind2 >= Ind1 Then
                lngHitCountDimmed = 100
                ReDim Hits(lngHitCountDimmed)
                
                With UMCNetLockers
                    For j = Ind1 To Ind2
                        If ((Abs(dblNetToMatch - .Lockers(mNetAdjLockerIndexPointers(j)).NET) <= UMCNetAdjDef.NETTol)) Or _
                            UMCNetAdjDef.NETTol < 0 Then
                            ' Within NET Tolerance (or tolerance is negative)
                            ' See if the charge is valid
                            
                            blnValidHit = True
                            With .Lockers(mNetAdjLockerIndexPointers(j))
                                ' Make sure at least one of the charges for this Net Locker is present in the UMC
                                If mUMCClsStats(PeakUMCInd(i), ustChargeMin) >= .ChargeMinimum And _
                                   mUMCClsStats(PeakUMCInd(i), ustChargeMin) <= .ChargeMaximum Then
                                   ' Valid Charge
                                ElseIf mUMCClsStats(PeakUMCInd(i), ustChargeMax) >= .ChargeMinimum And _
                                       mUMCClsStats(PeakUMCInd(i), ustChargeMax) <= .ChargeMaximum Then
                                    ' Valid Charge
                                Else
                                    blnValidHit = False
                                End If
                            End With
                            
                            If blnValidHit Then
                                Hits(lngHitCount) = j
                                lngHitCount = lngHitCount + 1
                                If lngHitCount > lngHitCountDimmed Then
                                    lngHitCountDimmed = lngHitCountDimmed + 100
                                    ReDim Preserve Hits(lngHitCountDimmed)
                                End If
                            End If
                        End If
                    Next j
                End With
            End If
        End If
        
        If lngHitCount > 0 Then
        
            ' The following implements the option "Do not use peaks pointing to multiple IDs on NET distance of more than ..."
            If UMCNetAdjDef.UseMultiIDMaxNETDist And lngHitCount > 1 Then
                ' Examine the NET values for the Hits
                ' If the range of NET values is > .MultiIDMaxNETDist then do not use this match
                With UMCNetLockers
                    dblNETMin = .Lockers(Hits(0)).NET
                    dblNETMax = .Lockers(Hits(0)).NET
                    For j = 1 To lngHitCount - 1
                        If .Lockers(Hits(j)).NET < dblNETMin Then dblNETMin = .Lockers(Hits(j)).NET
                        If .Lockers(Hits(j)).NET > dblNETMax Then dblNETMax = .Lockers(Hits(j)).NET
                    Next j
                End With
                
                If Abs(dblNETMax - dblNETMin) > UMCNetAdjDef.MultiIDMaxNETDist Then
                    lngHitCount = 0
                End If
            End If
            
            If lngHitCount > 0 Then
                For j = 0 To lngHitCount - 1
                    IDType(IDCnt) = PeakType(i)
                    IDInd(IDCnt) = PeakInd(i)
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = PeakScan(i)
                    IDCnt = IDCnt + 1
                Next j
                With GelUMC(CallerID).UMCs(PeakUMCInd(i))
                    .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
                    .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_LOCKER_HIT
                End With
            End If
        End If
    Next i
    
    If IDCnt > 0 Then
       ReDim Preserve ID(IDCnt - 1)
       ReDim Preserve IDInd(IDCnt - 1)
       ReDim Preserve IDType(IDCnt - 1)
       ReDim Preserve IDScan(IDCnt - 1)
       ReDim Preserve IDState(IDCnt - 1)
    Else
       Call ClearIDArrays
    End If
    
    UpdateStatus "Possible identifications: " & IDCnt
    SearchNetLockers = True
Exit Function

SearchNetAdjLockersErrorHandler:
    UpdateStatus "Error searching for Net Lockers!"
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.SearchNetLockers"
    SearchNetLockers = False

End Function

Private Sub cboPairsUMCsToUseForNETAdjustment_Click()
    glbPreferencesExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection = cboPairsUMCsToUseForNETAdjustment.ListIndex
End Sub

Private Sub chkAcceptLastIteration_Click()
UMCNetAdjDef.IterationAcceptLast = (chkAcceptLastIteration.value = vbChecked)
End Sub

Private Sub chkCS_Click(Index As Integer)
UMCNetAdjDef.PeakCSSelection(Index) = (chkCS(Index).value = vbChecked)
End Sub

Private Sub chkDecMW_Click()
UMCNetAdjDef.IterationUseMWDec = (chkDecMW.value = vbChecked)
End Sub

Private Sub chkDecNET_Click()
UMCNetAdjDef.IterationUseNETdec = (chkDecNET.value = vbChecked)
End Sub

Private Sub chkEliminateConflictingIDs_Click()
UMCNetAdjDef.UseMultiIDMaxNETDist = (chkEliminateConflictingIDs.value = vbChecked)
End Sub

Private Sub chkNetAdjAutoIncrementUMCTopAbuPct_Click()
    With glbPreferencesExpanded.AutoAnalysisOptions
        .NETAdjustmentAutoIncrementUMCTopAbuPct = cChkBox(chkNetAdjAutoIncrementUMCTopAbuPct)
        If .NETAdjustmentAutoIncrementUMCTopAbuPct Then
            SetCheckBox chkUMCUseTopAbu, True
            txtUMCAbuTopPct = .NETAdjustmentUMCTopAbuPctInitial
            UMCNetAdjDef.TopAbuPct = .NETAdjustmentUMCTopAbuPctInitial
            
            If optIteStop(ITERATION_STOP_CHANGE).value <> True Then
                optIteStop(ITERATION_STOP_CHANGE).value = True
            End If
        End If
    End With
End Sub

Private Sub chkNetAdjUseLockers_Click()
    UMCNetAdjDef.UseNetAdjLockers = cChkBox(chkNetAdjUseLockers)
End Sub

Private Sub chkNetAdjUseOldIfFailure_Click()
    UMCNetAdjDef.UseOldNetAdjIfFailure = cChkBox(chkNetAdjUseOldIfFailure)
End Sub

Private Sub chkUseN15AMTMasses_Click()
    glbPreferencesExpanded.NetAdjustmentUsesN15AMTMasses = cChkBox(chkUseN15AMTMasses)
End Sub

Private Sub chkUseNETForID_Click()
UMCNetAdjDef.UseNET = (chkUseNETForID.value = vbChecked)
End Sub

Private Sub cmdPause_Click()
Select Case cmdPause.Caption
Case CMD_CAPTION_PAUSE
     bPause = True
     cmdPause.Caption = CMD_CAPTION_CONTINUE
Case CMD_CAPTION_CONTINUE
     bPause = False
     cmdPause.Caption = CMD_CAPTION_PAUSE
End Select
End Sub

Private Sub cmdReset_Click()
    ResetToGenericNetAdjSettings
End Sub

Private Sub cmdStart_Click()
    CalculateIteration
End Sub

Private Sub cmdStop_Click()
bStop = True
End Sub

Private Sub cmdUseDefaults_Click()
    DisplayDefaultSettings
End Sub

Private Sub Form_Activate()
    InitializeNETAdjustment
End Sub

Private Sub Form_Load()
    InitializeForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GelUMCNETAdjDef(CallerID) = UMCNetAdjDef
    ' MonroeMod: We need to destroy the search object here in case it was populated with N15 AMT masses
    DestroyMTSearchObject
End Sub

Private Sub mnuE_Click()
Call PickParameters
End Sub

Private Sub mnuEGANET_Click()
On Error Resume Next
With GelAnalysis(CallerID)
    If .GANET_Slope <> 0 Then
        txtNETFormula.Text = ConstructNETFormula(.GANET_Slope, .GANET_Intercept)
    Else
        txtNETFormula.Text = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
    End If
End With
If Err Then
     MsgBox "Make sure display is loaded as analysis! Use New Analysis command from the File menu!", vbOKOnly, glFGTU
     Exit Sub
End If
CheckNETEquationStatus
txtNETFormula.SetFocus
End Sub

Private Sub mnuEGeneric_Click()
    ResetToGenericNetAdjSettings
End Sub

Private Sub mnuETICFit_Click()
On Error Resume Next
With GelAnalysis(CallerID)
     If .NET_Slope <> 0 Then
        txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
     Else
        txtNETFormula.Text = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
     End If
    
End With
If Err Then
   MsgBox "Make sure display is loaded as analysis! Use New Analysis command from the File menu!", vbOKOnly, glFGTU
   Exit Sub
End If
CheckNETEquationStatus
txtNETFormula.SetFocus
End Sub

Private Sub mnuF_Click()
Call PickParameters
End Sub

Private Sub mnuFCalculate_Click()
On Error GoTo exit_mnuFCalculate
Me.MousePointer = vbHourglass
Call ResetProcedure
If SelectUMCToUse() > 0 Then
   If SelectPeaksToUse() > 0 Then
      If UMCNetAdjDef.UseNET Then
         If Not SearchMassTagsMWNET() Then Exit Sub
      Else
         If Not SearchMassTagsMW() Then Exit Sub
      End If
      If Not FillTheGRID() Then Exit Sub
      If SelectIdentifications() > 1 Then
         Call CalculateSlopeIntercept
         If Not GelAnalysis(CallerID) Is Nothing Then
            GelAnalysis(CallerID).GANET_Fit = AdjAvD
            GelAnalysis(CallerID).GANET_Slope = AdjSlp
            GelAnalysis(CallerID).GANET_Intercept = AdjInt
         End If
         MsgBox "Slope: " & AdjSlp & vbCrLf & "Intercept: " & AdjInt & vbCrLf & "Average Deviation: " & AdjAvD, vbOKOnly, glFGTU
      Else
         GoTo err_mnuFCalculate
      End If
   Else
      GoTo err_mnuFCalculate
   End If
Else
   GoTo err_mnuFCalculate
End If

UpdateAnalysisHistory 1, AdjSlp, AdjInt, AdjAvD

exit_mnuFCalculate:
UpdateStatus "Ready"
Me.MousePointer = vbDefault
Exit Sub

err_mnuFCalculate:
MsgBox "Error calculating NET adjustment!(Hint: selected criteria could be the reason for this procedure to fail!)", vbOKOnly, glFGTU
Resume exit_mnuFCalculate
End Sub

Private Sub mnuFClose_Click()
    Unload Me
End Sub

Private Sub mnuFExportToMTDB_Click()
    Dim eResponse As VbMsgBoxResult
    
    eResponse = MsgBox("Proceed with exporting of the NET Slope and Intercept to the database?  This is an advanced feature that should normally only be performed during VIPER Automated PRISM Analysis Mode.  If you continue, you will be prompted for a password.", vbQuestion + vbYesNo + vbDefaultButton1, "Export NET")
    If eResponse = vbYes Then
        If QueryUserForExportToDBPassword(, False) Then
            MsgBox ExportGANETtoMTDB(CallerID, AdjSlp, AdjInt, AdjAvD)
        Else
            MsgBox "Invalid password, export aborted.", vbExclamation Or vbOKOnly, "Invalid"
        End If
    End If
End Sub

Private Sub mnuFIteration_Click()
    CalculateIteration
End Sub

Private Sub mnuFReport_Click()
Call ReportAdjustments
End Sub

Private Sub mnuLoadLegacy_Click()
    LoadLegacyMassTags
End Sub

Private Sub mnuMT_Click()
Call PickParameters
End Sub

Private Sub mnuMTLoad_Click()
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
MsgBox CheckMassTags(), vbOKOnly
End Sub

Private Sub optIteStop_Click(Index As Integer)
UMCNetAdjDef.IterationStopType = Index
' Assign default values to txtIteStopVal
Select Case Index
Case ITERATION_STOP_NUMBER
    txtIteStopVal = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMaxIterationCount
Case ITERATION_STOP_MW_TOL_LIMIT
    txtIteStopVal = "10"
Case ITERATION_STOP_NET_TOL_LIMIT
    txtIteStopVal = "0.02"
Case ITERATION_STOP_ID_LIMIT
    txtIteStopVal = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount
Case ITERATION_STOP_CHANGE
    txtIteStopVal = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentChangeThresholdStopValue
Case Else
    ' Do not set a default
End Select
Call txtIteStopVal_LostFocus
End Sub

Private Sub optPeakAbuCriteria_Click(Index As Integer)
UMCNetAdjDef.PeakSelection = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   UMCNetAdjDef.MWTolType = gltPPM
Else
   UMCNetAdjDef.MWTolType = gltABS
End If
End Sub

Private Sub CleanIdentifications()
'-------------------------------------------------------------
'removes identifications that will not be used from the arrays
'-------------------------------------------------------------
Dim i As Long
Dim NewCnt As Long
On Error Resume Next
UpdateStatus "Restructuring data ..."
For i = 0 To IDCnt - 1
    If IDState(i) = 0 Then
       NewCnt = NewCnt + 1
       ID(NewCnt - 1) = ID(i)
       IDInd(NewCnt - 1) = IDInd(i)
       IDType(NewCnt - 1) = IDType(i)
       IDScan(NewCnt - 1) = IDScan(i)
    End If
Next i
If NewCnt > 0 Then
   ReDim Preserve ID(NewCnt - 1)
   ReDim Preserve IDInd(NewCnt - 1)
   ReDim Preserve IDType(NewCnt - 1)
   ReDim IDState(NewCnt - 1)            'after cleaning always put state to 0
   ReDim Preserve IDScan(NewCnt - 1)
Else
   Call ClearIDArrays
End If
If NewCnt < IDCnt Then
   IDCnt = NewCnt
   ClearTheGRID         'have to recalculate GRID data
   Call FillTheGRID
End If
UpdateStatus ""
End Sub

Public Sub ClearIDArrays()
'-------------------------------------------------------------------
'clears arrays of identifications to be used for UMC NET adjustments
'-------------------------------------------------------------------
Erase ID
Erase IDInd
Erase IDType
Erase IDState
Erase IDScan
IDCnt = 0
End Sub

Public Sub ClearPeakArrays()
'------------------------------------------------------------------
'clears arrays of peaks selected to be used for UMC NET adjustments
'------------------------------------------------------------------
Erase PeakInd
Erase PeakType
Erase PeakMW
Erase PeakScan
Erase PeakUMCInd
PeakCnt = 0
End Sub

Private Sub ReportAdjustments()
'-------------------------------------------------------------------------
'report identifications on which adjustment is based and actual adjustment
'-------------------------------------------------------------------------
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim fname As String
Dim i As Long
On Error Resume Next
UpdateStatus "Generating report ..."
fname = GetTempFolder() & RawDataTmpFile
Set ts = fso.OpenTextFile(fname, ForWriting, True)
ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
'print gel file name and pairs definitions as reference
ts.WriteLine "Gel File: " & GelBody(CallerID).Caption
ts.WriteLine "Reporting NET adjustment(based on Mass Tag NETs) of Unique Mass Classes"
ts.WriteLine
ts.WriteLine "Slope: " & AdjSlp
ts.WriteLine "Intercept: " & AdjInt
ts.WriteLine "Average Deviation: " & AdjAvD
ts.WriteLine
ts.WriteLine "ID" & glARG_SEP & "ID_NET" & glARG_SEP & "Scan"
If IDsAreNetAdjLockers Then
    With UMCNetLockers
        For i = 0 To IDCnt - 1
            With .Lockers(ID(i))
                ts.WriteLine .GANETLockerID & glARG_SEP & .NET & glARG_SEP & IDScan(i)
            End With
        Next i
    End With
Else
    For i = 0 To IDCnt - 1
        ts.WriteLine AMTID(ID(i)) & glARG_SEP & AMTNET(ID(i)) & glARG_SEP & IDScan(i)
    Next i
End If

ts.Close
Set fso = Nothing
UpdateStatus ""
DoEvents
frmDataInfo.Tag = "AdjNET"
frmDataInfo.Show vbModal
End Sub

Private Function FillTheGRID() As Boolean
'----------------------------------------
'fills GRID arrays with ID information
'----------------------------------------
Dim i As Long
Dim DummyInd() As Long      'dummy array(empty) will allow us to
                            'sort only on one array
Dim QSL As QSLong
On Error GoTo err_FillTheGRID
UpdateStatus "Loading data structures ..."
If IDCnt > 0 And AMTCnt > 0 Then
   ReDim GRID(AMTCnt)      'AMT arrays are 1-based
   For i = 0 To IDCnt - 1
       With GRID(ID(i))
           .Count = .Count + 1
           ReDim Preserve .Members(.Count - 1)
           .Members(.Count - 1) = i
       End With
   Next i
   'order members of each group on scan numbers
   For i = 0 To AMTCnt
       If GRID(i).Count > 1 Then
          Set QSL = New QSLong
          If Not QSL.QSAsc(GRID(i).Members, DummyInd) Then GoTo err_FillTheGRID
          Set QSL = Nothing
       End If
   Next i
   FillTheGRID = True
   UpdateStatus ""
Else
   UpdateStatus "Data not found!"
End If
Exit Function

err_FillTheGRID:
Select Case Err.Number
Case 7
   Call ClearIDArrays
   Call ClearTheGRID
   Call ClearPeakArrays
   UpdateStatus ""
   If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
      MsgBox "System low on memory. Process aborted in recovery attempt!", vbOKOnly, glFGTU
   End If
Case Else
   Call ClearIDArrays
   Call ClearTheGRID
   Call ClearPeakArrays
   UpdateStatus "Error loading data structures!"
End Select
End Function

Private Sub UpdateAnalysisHistory(lngIterationCount As Long, dblSlope As Double, dblIntercept As Double, dblAverageDeviation As Double, Optional blnUseAbbreviatedFormat As Boolean = False)
    Dim strDescription As String
    Dim strMessage As String
    
    With UMCNetAdjDef
        If blnUseAbbreviatedFormat Then
            strDescription = "NET Adjustment iteration; " & UMC_NET_ADJ_ITERATION_COUNT & " = " & Trim(lngIterationCount) & "; Mass tolerance = ±" & Trim(UMCNetAdjDef.MWTol) & " " & GetSearchToleranceUnitText(CInt(UMCNetAdjDef.MWTolType)) & "; NET Tolerance = ±" & Trim(UMCNetAdjDef.NETTol)
            If .TopAbuPct >= 0 Then
                strDescription = strDescription & "; Restrict to x% of UMC's = " & Trim(.TopAbuPct) & "%"
            End If
            
            strDescription = strDescription & "; UMC Peaks in tolerance = " & Trim(PeakCnt) & "; " & UMC_NET_ADJ_PEAKS_WITH_DB_HITS & " = " & Trim(IDCnt) & "; NET Formula = " & ConstructNETFormula(dblSlope, dblIntercept)
        Else
            strDescription = "Calculated NET adjustment using UMC's; " & UMC_NET_ADJ_ITERATION_COUNT & " = " & Trim(lngIterationCount) & "; Mass tolerance = ±" & Trim(UMCNetAdjDef.MWTol) & " " & GetSearchToleranceUnitText(CInt(UMCNetAdjDef.MWTolType)) & "; Final NET Tolerance = ±" & Trim(UMCNetAdjDef.NETTol) & "; Peak Abu Criteria = " & GetPeakAbuCriteria(UMCNetAdjDef.PeakSelection)
            strDescription = strDescription & "; Min UMC Peak Count = " & Trim(.MinUMCCount) & "; Min UMC Scan Range = " & Trim(.MinScanRange)
            If .TopAbuPct >= 0 Then
                strDescription = strDescription & "; Restrict to x% of UMC's (sorted by abundance) = " & Trim(.TopAbuPct) & "%"
            End If
            
            strDescription = strDescription & "; UMC Peaks in tolerance = " & Trim(PeakCnt) & "; " & UMC_NET_ADJ_PEAKS_WITH_DB_HITS & " = " & Trim(IDCnt) & "; Slope = " & Trim(AdjSlp) & "; Intercept = " & Trim(AdjInt) & "; Average Deviation = " & Trim(dblAverageDeviation) & "; NET Formula = " & ConstructNETFormula(dblSlope, dblIntercept)
            
        End If
    End With
    
    AddToAnalysisHistory CallerID, strDescription
    
    If (UMCSegmentCntWithLowUMCCnt > 0 Or UMCCntAddedSinceLowSegmentCount > 0) And Not blnUseAbbreviatedFormat Then
        strDescription = "NET Adjustment UMC usage dispersion was low in 1 or more segments; Total segment count = " & Trim(glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions.SegmentCount) & "; Segment count with low UMC counts = " & Trim(UMCSegmentCntWithLowUMCCnt) & "; UMC's added (total) = " & Trim(UMCCntAddedSinceLowSegmentCount)
        AddToAnalysisHistory CallerID, strDescription
    End If
    
    If Not blnUseAbbreviatedFormat Then
        strMessage = "Database for NET adjustment: " & CurrMTDBInfo()
        strMessage = strMessage & "; N15 masses used for AMT's = " & CStr(glbPreferencesExpanded.NetAdjustmentUsesN15AMTMasses)
        strMessage = strMessage & "; Minimum high normalized score for AMT's = " & CStr(mMTMinimumHighNormalizedScore)
        strMessage = strMessage & "; Minimum high discriminant score for AMT's = " & CStr(mMTMinimumHighDiscriminantScore)
        AddToAnalysisHistory CallerID, strMessage
    End If

End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub ClearTheGRID()
'--------------------------------------------
'destroys GRID data structure
'--------------------------------------------
Dim i As Long
For i = 0 To UBound(GRID)
    If GRID(i).Count > 0 Then Erase GRID(i).Members
Next i
Erase GRID
End Sub

'''Private Sub MarkMultiIDsLongDistance()
''''------------------------------------------------------------------------------------
''''sets state of all identifications spanning through the too long range
''''to STATE_ID_NETS_TOO_DISTANT
''''NOTE: this does not work as it should - it should mark IDs coming from the same peak
''''that are too far apart from each other in NET direction
''''------------------------------------------------------------------------------------
'''Dim AllowedScanRange As Double
'''Dim i As Long, j As Long
'''On Error Resume Next
'''UpdateStatus "Eliminating IDs pointed with the same peak ..."
''''pretend that we have NET percentage
'''AllowedScanRange = CLng(UMCNetAdjDef.MultiIDMaxNETDist * ScanRange)
'''For i = 0 To IDCnt - 1
'''  With GRID(ID(i))
'''    If .Count > 1 Then
'''       If (IDScan(.Members(.Count - 1)) - IDScan(.Members(0))) > AllowedScanRange Then
'''          For j = 0 To .Count - 1       'mark them all as too long
'''              IDState(.Members(j)) = IDState(.Members(j)) + STATE_ID_NETS_TOO_DISTANT
'''          Next j
'''       End If
'''    End If
'''  End With
'''Next i
'''UpdateStatus ""
'''End Sub


Private Sub MarkIDsWithBadNET()
'-----------------------------------------------------------------------
'sets state of all identifications with bad NET numbers to STATE_BAD_NET
'-----------------------------------------------------------------------
Dim i As Long
On Error Resume Next
UpdateStatus "Eliminating IDs with bad elution ..."
If Not IDsAreNetAdjLockers Then
    For i = 0 To IDCnt - 1
        If AMTNET(ID(i)) < 0 Or AMTNET(ID(i)) > 1 Then
           IDState(i) = IDState(i) + STATE_BAD_NET
        End If
    Next i
End If
UpdateStatus ""
End Sub

Private Sub Info_NoMTDBLink()
'this message is used twice so ...
MsgBox "Current display is not associated with any Mass Tags database!" & vbCrLf _
     & "Close dialog and establish association (Edit->Select/Modify Database Connection)" & vbCrLf _
     & "or select Mass Tags->Load Legacy MT DB on this dialog to load" & vbCrLf _
     & "data from legacy database!", vbOKOnly, glFGTU
End Sub

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

Private Function ScoreIDs() As Boolean
'-------------------------------------------------------------------
'score ids and set states of IDs that don't make top MaxIDToUse
'to STATE_OUTSCORED
'Score=Log10(Peak Abundance) - Fit for Isotopic peaks and
'Log10(Peak Abundance) for Charge State peaks
'-------------------------------------------------------------------
Dim ScoreOrder() As Long
Dim Score() As Double
Dim qsd As New QSDouble
Dim i As Long
On Error GoTo err_ScoreIDs

ReDim ScoreOrder(IDCnt - 1)
ReDim Score(IDCnt - 1)
With GelData(CallerID)
    For i = 0 To IDCnt - 1
        ScoreOrder(i) = i
        Select Case IDType(i)
        Case glCSType
             Score(i) = (Log(.CSNum(IDInd(i), csfAbu)) / Log(10))
        Case glIsoType
             Score(i) = (Log(.IsoNum(IDInd(i), isfAbu)) / Log(10)) - .IsoNum(IDInd(i), isfFit)
        End Select
    Next i
End With
If Not qsd.QSDesc(Score(), ScoreOrder()) Then GoTo err_ScoreIDs
Set qsd = Nothing

If IDCnt > UMCNetAdjDef.MaxIDToUse Then
   For i = 0 To UMCNetAdjDef.MaxIDToUse - 1
       IDState(ScoreOrder(i)) = 0
   Next i
   For i = UMCNetAdjDef.MaxIDToUse To IDCnt - 1
       IDState(ScoreOrder(i)) = IDState(ScoreOrder(i)) + STATE_OUTSCORED
   Next i
Else
   For i = 0 To IDCnt - 1
       IDState(ScoreOrder(i)) = 0
   Next i
End If
ScoreIDs = True
Exit Function

err_ScoreIDs:
LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.ScoreIDs"
End Function

Private Sub txtIteMWDec_LostFocus()
If IsNumeric(txtIteMWDec.Text) Then
   UMCNetAdjDef.IterationMWDec = CDbl(txtIteMWDec.Text)
Else
   MsgBox "This parameter should be numeric!", vbOKOnly, glFGTU
   txtIteMWDec.SetFocus
End If
End Sub

Private Sub txtIteNETDec_LostFocus()
If IsNumeric(txtIteNETDec.Text) Then
   UMCNetAdjDef.IterationNETDec = CDbl(txtIteNETDec.Text)
Else
   MsgBox "This parameter should be numeric!", vbOKOnly, glFGTU
   txtIteNETDec.SetFocus
End If
End Sub

Private Sub txtIteStopVal_LostFocus()
If IsNumeric(txtIteStopVal.Text) Then
   UMCNetAdjDef.IterationStopValue = CDbl(txtIteStopVal.Text)
   If optIteStop(ITERATION_STOP_CHANGE).value = True Then
      glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentChangeThresholdStopValue = UMCNetAdjDef.IterationStopValue
   ElseIf optIteStop(ITERATION_STOP_ID_LIMIT).value = True Then
      glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount = UMCNetAdjDef.IterationStopValue
      txtNetAdjMinIDCount = txtIteStopVal
   ElseIf optIteStop(ITERATION_STOP_NUMBER).value = True Then
      glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMaxIterationCount = UMCNetAdjDef.IterationStopValue
   End If
Else
   MsgBox "This parameter should be numeric!", vbOKOnly, glFGTU
   txtIteStopVal.SetFocus
End If
End Sub

Private Sub txtMaxUMCScansPct_LostFocus()
If IsNumeric(txtMaxUMCScansPct.Text) Then
   UMCNetAdjDef.MaxScanPct = Abs(CDbl(txtMaxUMCScansPct.Text))
Else
   MsgBox "This parameter should be positive number!", vbOKOnly, glFGTU
   txtMaxUMCScansPct.SetFocus
End If
End Sub

Private Sub txtMinScanRange_LostFocus()
If IsNumeric(txtMinScanRange.Text) Then
   UMCNetAdjDef.MinScanRange = Abs(CLng(txtMinScanRange.Text))
Else
   MsgBox "This parameter should be non-negative integer!", vbOKOnly, glFGTU
   txtMinScanRange.SetFocus
End If
End Sub

Private Sub txtMinUMCCount_LostFocus()
If IsNumeric(txtMinUMCCount.Text) Then
   UMCNetAdjDef.MinUMCCount = Abs(CLng(txtMinUMCCount.Text))
Else
   MsgBox "This parameter should be non-negative integer!", vbOKOnly, glFGTU
   txtMinUMCCount.SetFocus
End If
End Sub

Private Sub txtMultiIDMaxNETDist_LostFocus()
If IsNumeric(txtMultiIDMaxNETDist.Text) Then
   UMCNetAdjDef.MultiIDMaxNETDist = Abs(CDbl(txtMultiIDMaxNETDist.Text))
Else
   MsgBox "This parameter should be non-negative integer!", vbOKOnly, glFGTU
   txtMultiIDMaxNETDist.SetFocus
End If
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   UMCNetAdjDef.MWTol = Abs(CDbl(txtMWTol.Text))
Else
   MsgBox "This parameter should be non-negative number!", vbOKOnly, glFGTU
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNetAdjMinHighDiscriminantScore_LostFocus()
If IsNumeric(txtNetAdjMinHighDiscriminantScore.Text) Then
    glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore = Abs(CSng(txtNetAdjMinHighDiscriminantScore.Text))
    If glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore > 1 Then
        glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore = 0.999
    End If
Else
    MsgBox "This parameter should be non-negative integer!", vbOKOnly, glFGTU
    txtNetAdjMinHighDiscriminantScore.SetFocus
End If
End Sub

Private Sub txtNetAdjMinHighNormalizedScore_LostFocus()
If IsNumeric(txtNetAdjMinHighNormalizedScore.Text) Then
    glbPreferencesExpanded.NetAdjustmentMinHighNormalizedScore = Abs(CSng(txtNetAdjMinHighNormalizedScore.Text))
Else
    MsgBox "This parameter should be non-negative integer!", vbOKOnly, glFGTU
    txtNetAdjMinHighNormalizedScore.SetFocus
End If
End Sub

Private Sub txtNetAdjMinIDCount_LostFocus()
    ValidateTextboxValueLng txtNetAdjMinIDCount, 2, 100000, 75
    If optIteStop(ITERATION_STOP_ID_LIMIT).value = True Then
        txtIteStopVal = txtNetAdjMinIDCount
        Call txtIteStopVal_LostFocus
    End If
    glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount = CLngSafe(txtNetAdjMinIDCount)
End Sub

Private Sub txtNetAdjMinLockerMatchCount_LostFocus()
    ValidateTextboxValueLng txtNetAdjMinLockerMatchCount, 2, 1000, 3
    UMCNetAdjDef.NetAdjLockerMinimumMatchCount = CLngSafe(txtNetAdjMinLockerMatchCount)
End Sub

Private Sub txtNETFormula_LostFocus()
    ValidateNETFormula
End Sub

Private Sub txtNETTol_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtNETTol, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtNETTol_LostFocus()
Dim Tmp As String
Tmp = Trim$(txtNETTol.Text)
If IsNumeric(Tmp) Then
   If Tmp >= 0 And Tmp <= 1 Then
      UMCNetAdjDef.NETTol = CDbl(Tmp)
      Exit Sub
   End If
End If
MsgBox "NET Tolerance should be a number between 0 and 1!", vbOKOnly, glFGTU
txtNETTol.SetFocus
End Sub

' Unused Function (March 2003)
'''Private Function GetUnqIDCntUsed() As Long
''''------------------------------------------------
''''returns number of unique IDs used in calculation
''''------------------------------------------------
'''Dim i As Long, Cnt As Long
'''On Error Resume Next
'''For i = 0 To UBound(GRID)
'''    If GRID(i).Count > 0 Then Cnt = Cnt + 1
'''Next i
'''GetUnqIDCntUsed = Cnt
'''End Function

Private Sub CalculateSlopeIntercept()
'-----------------------------------------------------
'least square method to lay best straight line through
'set of points (xi,yi)
'-----------------------------------------------------
Dim SumY As Double
Dim SumX As Double
Dim SumXY As Double
Dim SumXX As Double
Dim i As Long
UpdateStatus "Calculating Slope & Intercept"
SumY = 0
SumX = 0
SumXY = 0
SumXX = 0
' Loop through all the selected identifications
For i = 0 To IDCnt - 1
    SumX = SumX + IDScan(i)
    If IDsAreNetAdjLockers Then
        With UMCNetLockers.Lockers(ID(i))
            SumY = SumY + .NET
            SumXY = SumXY + IDScan(i) * .NET
        End With
    Else
        SumY = SumY + AMTNET(ID(i))
        SumXY = SumXY + IDScan(i) * AMTNET(ID(i))
    End If
    SumXX = SumXX + IDScan(i) * IDScan(i)
Next i
AdjSlp = (IDCnt * SumXY - SumX * SumY) / (IDCnt * SumXX - SumX * SumX)
AdjInt = (SumY - AdjSlp * SumX) / IDCnt
Call CalculateAvgDev
UpdateStatus ""
Exit Sub


err_CalculateSlopeIntercept:
UpdateStatus "Error calculating slope and intercept!"
AdjSlp = 0
AdjInt = 0
AdjAvD = -1
End Sub


Private Sub CalculateAvgDev()
Dim i As Long
Dim TtlDist As Double
On Error GoTo err_CalculateAvgDev
TtlDist = 0

If IDsAreNetAdjLockers Then
    With UMCNetLockers
        For i = 0 To IDCnt - 1
            TtlDist = TtlDist + (AdjSlp * IDScan(i) + AdjInt - .Lockers(ID(i)).NET) ^ 2
        Next i
    End With
Else
    For i = 0 To IDCnt - 1
        TtlDist = TtlDist + (AdjSlp * IDScan(i) + AdjInt - AMTNET(ID(i))) ^ 2
    Next i
End If

AdjAvD = TtlDist / IDCnt
Exit Sub

err_CalculateAvgDev:
AdjAvD = -1
End Sub


Private Sub PickParameters()
Call txtMinUMCCount_LostFocus
Call txtMinScanRange_LostFocus
Call txtMaxUMCScansPct_LostFocus
Call txtUMCAbuTopPct_LostFocus
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
Call txtMultiIDMaxNETDist_LostFocus
Call txtPctMaxAbu_LostFocus
Call txtIteMWDec_LostFocus
Call txtIteNETDec_LostFocus
Call txtIteStopVal_LostFocus
Call txtIteMWDec_LostFocus
Call txtIteNETDec_LostFocus
Call txtNETFormula_LostFocus
End Sub

Private Function ResetProcedure() As Boolean
'----------------------------------------------------------------------
'resets all arguments and parameters so that we can restart calculation
'----------------------------------------------------------------------
Dim i As Long
ReDim UseUMC(GelUMC(CallerID).UMCCnt - 1)       'initially use all UMCs
For i = 0 To UBound(UseUMC)
    UseUMC(i) = True
Next i
Call ClearIDArrays                              'reset arrays with identifications
Call ClearPeakArrays

If glbPreferencesExpanded.NetAdjustmentUsesN15AMTMasses <> AMTSearchObjectHasN15Masses Then
    ' Need to update the search object
    If Not CreateNewMTSearchObject(glbPreferencesExpanded.NetAdjustmentUsesN15AMTMasses) Then
        lblMTStatus.Caption = "Error creating search object!"
        ResetProcedure = False
    Else
        ResetProcedure = True
    End If
Else
    ResetProcedure = True
End If

End Function


Private Function SelectUMCToUse() As Long
'------------------------------------------------------------------
'selects unique mass classes that will be used to correct NET based
'on specified criteria; returns number of it; -1 on any error
'------------------------------------------------------------------
Dim i As Long
Dim Cnt As Long

Dim ePairedSearchUMCSelection As punaPairsUMCNetAdjustmentConstants

On Error GoTo exit_SelectUMCToUse
UpdateStatus "Selecting UMCs to use ...."
SelectUMCToUse = -1

For i = 0 To GelUMC(CallerID).UMCCnt - 1
    With GelUMC(CallerID).UMCs(i)
        .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
        .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_LOWSEGMENTCOUNT_ADDITION
        .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
    End With
Next i

ePairedSearchUMCSelection = glbPreferencesExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection
    
If ePairedSearchUMCSelection = punaPairedAll Or _
   ePairedSearchUMCSelection = punaPairedLight Or _
   ePairedSearchUMCSelection = punaPairedHeavy Then
    If Not PairsPresent(CallerID) Then
        ' No pairs are present
        ePairedSearchUMCSelection = punaPairedAndUnpaired
    End If
End If
    
If ePairedSearchUMCSelection <> punaPairedAndUnpaired Then
    Select Case ePairedSearchUMCSelection
    Case punaPairedAll, punaPairedLight, punaPairedHeavy
        ' First exclude everything
        ' Then, include UMC's that are paired, depending upon ePairedSearchUMCSelection
        For i = 0 To GelUMC(CallerID).UMCCnt - 1
            UseUMC(i) = False
        Next i
        
        If ePairedSearchUMCSelection = punaPairedAll Then
            ' Add back all UMC's belonging to pairs
            For i = 0 To GelP_D_L(CallerID).PCnt - 1
                UseUMC(GelP_D_L(CallerID).Pairs(i).P1) = True
                UseUMC(GelP_D_L(CallerID).Pairs(i).P2) = True
            Next i
        ElseIf ePairedSearchUMCSelection = punaPairedHeavy Then
            ' Add back UMC's belonging to the heavy member of pairs
            For i = 0 To GelP_D_L(CallerID).PCnt - 1
                UseUMC(GelP_D_L(CallerID).Pairs(i).P2) = True
            Next i
        Else
            ' punaPairedLight
            ' Add back UMC's belonging to the light member of pairs
            For i = 0 To GelP_D_L(CallerID).PCnt - 1
                UseUMC(GelP_D_L(CallerID).Pairs(i).P1) = True
            Next i
        End If
    Case punaUnpairedOnly
        ' Exclude UMC's that are paired
        For i = 0 To GelP_D_L(CallerID).PCnt - 1
            UseUMC(GelP_D_L(CallerID).Pairs(i).P1) = False
            UseUMC(GelP_D_L(CallerID).Pairs(i).P2) = False
        Next i
    Case punaUnpairedPlusPairedLight
        ' Exclude UMC's that belong to heavy members of pairs
        For i = 0 To GelP_D_L(CallerID).PCnt - 1
            UseUMC(GelP_D_L(CallerID).Pairs(i).P2) = False
        Next i
    End Select
    
End If

If UMCNetAdjDef.MinUMCCount > 1 Or UMCNetAdjDef.MinScanRange > 1 Then
    ' filter-out all mass classes with insufficient membership
    '  or
    ' filter-out all mass classes with insufficient scan coverage
    For i = 0 To GelUMC(CallerID).UMCCnt - 1
        If UseUMC(i) = True Then
            UseUMC(i) = UMCSelectionFilterCheck(i)
        End If
    Next i
End If

If UMCNetAdjDef.TopAbuPct >= 0 And UMCNetAdjDef.TopAbuPct < 100 Then
    ' Filter-out low abundant classes
    ' However, if .RequireDispersedUMCSelection = True, then make sure we have some UMC's from all portions of the data
    SelectUMCsToUseWork glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions.RequireDispersedUMCSelection
End If

Debug.Assert UBound(UseUMC) = GelUMC(CallerID).UMCCnt - 1
For i = 0 To UBound(UseUMC)
    With GelUMC(CallerID).UMCs(i)
        If UseUMC(i) Then
            Cnt = Cnt + 1
            ' Turn on the UMC_INDICATOR_BIT_USED_FOR_NET_ADJ bit
            .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
        End If
    End With
Next i
SelectUMCToUse = Cnt
Exit Function

exit_SelectUMCToUse:
Debug.Assert False
LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.SelectUMCToUse"

End Function

Private Sub SelectUMCsToUseWork(blnRequireDispersed As Boolean)
    ' If blnRequireDispersed = True, then assures that the selected UMC's are representative of all parts of the data
    
    Dim Abu() As Double         ' 0-based array; abundances to sort
    Dim TmpInd() As Long        ' 0-based array; original indices in GelUMC()
    Dim qsd As New QSDouble
    
    Dim udtSegmentStats() As udtSegmentStatsType       ' 0-based array
    
    Dim UMCTopAbuPctCnt As Long
    Dim UMCMinimumCntPerSegment As Long
    
    Dim lngIndex As Long, lngSegmentIndex As Long
    Dim lngMaxUnusedUMCs As Long
    Dim lngUMCIndex As Long
    Dim lngMatchingIndices() As Long
    Dim lngMatchCount As Long
    Dim lngUMCsSelectedBeforeSegmentChecking As Long
    Dim lngUMCsSelectedAfterSegmentChecking As Long
    
    Dim lngScanMin As Long, lngScanMax As Long
    Dim lngScansPerSegment As Long
    Dim lngScanCenter As Long, lngWorkingScan As Long
    
    Dim lngSegmentBin As Long
    Dim lngSegmentCount As Long
    
    Dim ePairedSearchUMCSelection As punaPairsUMCNetAdjustmentConstants
    Dim blnAddThisUMC As Boolean

    Dim objP1IndFastSearch As FastSearchArrayLong
    Dim objP2IndFastSearch As FastSearchArrayLong

On Error GoTo SelectUMCsToUseWorkErrorHandler

    UMCTopAbuPctCnt = CLng((UMCNetAdjDef.TopAbuPct / 100) * GelUMC(CallerID).UMCCnt)
    
    ' First select the UMC's to use
    ' What we do here is set the UseUMC() flag for the low abundance UMC's to false
    ' We do not take the pairing preferences into account when we do this; we will
    '  consider that below if blnRequireDispersed = True
    ReDim Abu(GelUMC(CallerID).UMCCnt - 1)
    ReDim TmpInd(GelUMC(CallerID).UMCCnt - 1)
    For lngIndex = 0 To GelUMC(CallerID).UMCCnt - 1
        With GelUMC(CallerID).UMCs(lngIndex)
            Abu(lngIndex) = .ClassAbundance
            TmpInd(lngIndex) = lngIndex
        End With
    Next lngIndex
    
    If qsd.QSDesc(Abu(), TmpInd()) Then
       If UMCTopAbuPctCnt > GelUMC(CallerID).UMCCnt Then UMCTopAbuPctCnt = GelUMC(CallerID).UMCCnt
       If UMCTopAbuPctCnt < 0 Then UMCTopAbuPctCnt = 0
       For lngIndex = UMCTopAbuPctCnt To GelUMC(CallerID).UMCCnt - 1
           UseUMC(TmpInd(lngIndex)) = False
       Next lngIndex
    End If

    UMCCntAddedSinceLowSegmentCount = 0
    UMCSegmentCntWithLowUMCCnt = 0

    If blnRequireDispersed Then
        ' Collect stats on the number of UMC's used per segment
        
        With glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions
            lngSegmentCount = .SegmentCount
            If lngSegmentCount < 1 Then lngSegmentCount = 1
            
            lngScanMin = ScanMin + (ScanMax - ScanMin) * (.ScanPctStart / 100)
            lngScanMax = ScanMin + (ScanMax - ScanMin) * (.ScanPctEnd / 100)
            
            If lngScanMin < ScanMin Then lngScanMin = ScanMin
            If lngScanMax > ScanMax Then lngScanMax = ScanMax
            
        End With
        
        ReDim udtSegmentStats(0 To lngSegmentCount - 1)
        
        ' Determine the total number of unused UMC's
        lngMaxUnusedUMCs = 0
        For lngUMCIndex = 0 To GelUMC(CallerID).UMCCnt - 1
            If Not UseUMC(lngUMCIndex) Then lngMaxUnusedUMCs = lngMaxUnusedUMCs + 1
        Next lngUMCIndex
        
        If lngMaxUnusedUMCs < 10 Then lngMaxUnusedUMCs = 10
        For lngSegmentIndex = 0 To lngSegmentCount - 1
            With udtSegmentStats(lngSegmentIndex)
                .ArrayCountUnused = 0
                ReDim .UnusedUMCIndices(lngMaxUnusedUMCs)
            End With
        Next lngSegmentIndex
        
        lngScansPerSegment = (lngScanMax - lngScanMin) / lngSegmentCount
        
        For lngIndex = 0 To GelUMC(CallerID).UMCCnt - 1
            With GelUMC(CallerID).UMCs(lngIndex)
                ' Compute the center scan of this UMC
                lngScanCenter = (.MaxScan + .MinScan) / 2
            End With
            
            If lngScanCenter >= lngScanMin And lngScanCenter <= lngScanMax Then
                ' Determine which segment this scan corresponds to
                
                ' First subtract lngScanMin from lngScanCenter
                ' For example, if lngScanMin is 100 and lngScanCenter is 250, then lngWorkingScan = 150
                lngWorkingScan = lngScanCenter - lngScanMin
                
                ' Now, dividing lngWorkingScan by lngScansPerSegment and rounding to the nearest integer
                '  actually gives the bin
                ' For example, given lngWorkingScan = 150 and lngScansPerSegment = 1000, Bin = CLng(150/1000) = 0
                lngSegmentBin = CLng(lngWorkingScan / lngScansPerSegment)
                
                If lngSegmentBin < 0 Then lngSegmentBin = 0
                If lngSegmentBin >= lngSegmentCount Then lngSegmentBin = lngSegmentCount - 1
                
                If UseUMC(lngIndex) Then
                    udtSegmentStats(lngSegmentBin).UMCHitCountUsed = udtSegmentStats(lngSegmentBin).UMCHitCountUsed + 1
                Else
                    ' Add to the array of potential UMC's that could be added if needed
                    With udtSegmentStats(lngSegmentBin)
                        .UnusedUMCIndices(.ArrayCountUnused) = lngIndex
                        .ArrayCountUnused = .ArrayCountUnused + 1
                    End With
                End If
            Else
                ' UMC is outside the desired scan range; ignore it
            End If
        Next lngIndex
        
        For lngSegmentIndex = 0 To lngSegmentCount - 1
            With udtSegmentStats(lngSegmentIndex)
                UMCMinimumCntPerSegment = (glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions.MinimumUMCsPerSegmentPctTopAbuPct / 100#) * (UMCNetAdjDef.TopAbuPct / 100#) * (.UMCHitCountUsed + .ArrayCountUnused)
                lngUMCsSelectedBeforeSegmentChecking = lngUMCsSelectedBeforeSegmentChecking + .UMCHitCountUsed
                If .UMCHitCountUsed < UMCMinimumCntPerSegment Then
                    ' Hit count is too low
                    UMCSegmentCntWithLowUMCCnt = UMCSegmentCntWithLowUMCCnt + 1
                    
                    ' If .ArrayCountUnused is more than 0, then fill the Abu() and TmpInd() arrays
                    '  and sort by abundance (ascending)
                    If .ArrayCountUnused > 0 Then
                        
                        ReDim Abu(.ArrayCountUnused - 1)
                        ReDim TmpInd(.ArrayCountUnused - 1)
                        
                        For lngIndex = 0 To .ArrayCountUnused - 1
                            Abu(lngIndex) = GelUMC(CallerID).UMCs(.UnusedUMCIndices(lngIndex)).ClassAbundance
                            TmpInd(lngIndex) = .UnusedUMCIndices(lngIndex)
                        Next lngIndex
                        
                        If qsd.QSAsc(Abu(), TmpInd()) Then
                            ' Add back in the necessary number of UMC's, taking into
                            '   account the value of ePairedSearchUMCSelection and

                            ePairedSearchUMCSelection = glbPreferencesExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection
                            
                            If ePairedSearchUMCSelection <> punaPairedAndUnpaired Then
                                ' Initialize the PairIndex lookup objects
                                If Not PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch) Then
                                    ' No pairs found; pretend we're including all UMC's
                                    ePairedSearchUMCSelection = punaPairedAndUnpaired
                                End If
                            End If
                            
                            lngIndex = 0
                            Do While .UMCHitCountUsed < UMCMinimumCntPerSegment And lngIndex < .ArrayCountUnused
                                
                                lngUMCIndex = TmpInd(lngIndex)
                                blnAddThisUMC = False
                                
                                Select Case ePairedSearchUMCSelection
                                Case punaPairedAndUnpaired, punaPairedAll
                                    ' Add no matter what
                                    blnAddThisUMC = True
                                Case punaPairedLight
                                    ' Only add if UMC is the light member of a pair
                                    If objP1IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        ' Match Found
                                        blnAddThisUMC = True
                                    End If
                                Case punaPairedHeavy
                                    ' Only add if UMC is the light member of a pair
                                    If objP2IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        ' Match Found
                                        blnAddThisUMC = True
                                    End If
                                Case punaUnpairedOnly
                                    ' Only add if UMC does not belong to a pair
                                    If Not objP1IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        If Not objP2IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                            ' No Match Found; use this UMC
                                            blnAddThisUMC = True
                                        End If
                                    End If
                                Case punaUnpairedPlusPairedLight
                                    ' Only add if UMC does not belong to heavy members of pairs
                                    If Not objP2IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        ' No Match Found; use this UMC
                                        blnAddThisUMC = True
                                    End If
                                End Select
                            
                                If blnAddThisUMC Then
                                    ' Make sure the UMC passes the Minimum UMC Member count and
                                    '  Minimum Scan Range requirements, if necessary
                                    
                                    blnAddThisUMC = UMCSelectionFilterCheck(lngUMCIndex)
                                    
                                    If blnAddThisUMC Then
                                        UseUMC(lngUMCIndex) = True
                                        ' Turn on the LowSegmentCountAddedUMC bit
                                        GelUMC(CallerID).UMCs(lngUMCIndex).ClassStatusBits = GelUMC(CallerID).UMCs(lngUMCIndex).ClassStatusBits Or UMC_INDICATOR_BIT_LOWSEGMENTCOUNT_ADDITION
                                        UMCCntAddedSinceLowSegmentCount = UMCCntAddedSinceLowSegmentCount + 1
                                        .UMCHitCountUsed = .UMCHitCountUsed + 1
                                    End If
                                End If
                                
                                lngIndex = lngIndex + 1
                            Loop
                            
                        End If
                        
                    End If
                End If
                lngUMCsSelectedAfterSegmentChecking = lngUMCsSelectedAfterSegmentChecking + .UMCHitCountUsed
            End With
        Next lngSegmentIndex
    End If

Exit Sub

SelectUMCsToUseWorkErrorHandler:
Debug.Assert False
LogErrors Err.Number, "frmSearchForNETAdjustmentUMC.SelectUMCToUseWork"

End Sub

Private Function SelectPeaksToUse() As Long
'------------------------------------------------------------------
'selects peaks that will be used to correct NET based on specified
'criteria; returns number of it; -1 on any error
'NOTE: peaks are selected only from UMCs marked to be used
'------------------------------------------------------------------
Dim i As Long
On Error GoTo exit_SelectPeaksToUse
SelectPeaksToUse = -1
UpdateStatus "Selecting peaks to use ..."
For i = 0 To UBound(UseUMC)
    If UseUMC(i) Then
       Select Case UMCNetAdjDef.PeakSelection
       Case UMCNetConstants.UMCNetBefore
            Call GetUMCPeak_Before(i)
       Case UMCNetConstants.UMCNetAt
            Call GetUMCPeak_At(i)
       Case UMCNetConstants.UMCNetAfter
            Call GetUMCPeak_After(i)
       Case UMCNetConstants.UMCNetFirst
            Call GetUMCPeak_First(i)
       Case UMCNetConstants.UMCNetLast
            Call GetUMCPeak_Last(i)
       End Select
    End If
Next i
If PeakCnt > 0 Then
   ReDim Preserve PeakInd(PeakCnt - 1)
   ReDim Preserve PeakType(PeakCnt - 1)
   ReDim Preserve PeakMW(PeakCnt - 1)
   ReDim Preserve PeakScan(PeakCnt - 1)
   ReDim Preserve PeakUMCInd(PeakCnt - 1)
Else
   ClearPeakArrays
End If
SelectPeaksToUse = PeakCnt
exit_SelectPeaksToUse:
End Function

Private Function SelectIdentifications() As Long
'------------------------------------------------------------------
'selects IDs that will be used to correct NET based on specified
'criteria; returns number of it; -1 on any error
'------------------------------------------------------------------
On Error GoTo exit_SelectIdentifications
UpdateStatus "Selecting IDs ..."
Call ScoreIDs
If UMCNetAdjDef.EliminateBadNET Then Call MarkIDsWithBadNET

'' MonroeMod: The job of Excluding multiple matches over long scan ranges is now part of SearchMassTagsMWNET()
'' If UMCNetAdjDef.UseMultiIDMaxNETDist Then Call MarkMultiIDsLongDistance

Call CleanIdentifications
SelectIdentifications = IDCnt

exit_SelectIdentifications:
End Function

Private Function UMCSelectionFilterCheck(lngUMCIndex) As Boolean
    
    Dim blnValidUMC As Boolean
    
    blnValidUMC = True
    With GelUMC(CallerID).UMCs(lngUMCIndex)
        If UMCNetAdjDef.MinUMCCount > 1 Then                   'filter-out all mass classes with insufficient membership
            If .ClassCount < UMCNetAdjDef.MinUMCCount Then
                blnValidUMC = False
            End If
        End If
        
        If UMCNetAdjDef.MinScanRange > 1 Then                  'filter-out all mass classes with insufficient scan coverage
            If (.MaxScan - .MinScan + 1) < UMCNetAdjDef.MinScanRange Then
                blnValidUMC = False
            End If
        End If
    End With

    UMCSelectionFilterCheck = blnValidUMC

End Function

Private Sub GetUMCPeak_Before(ByVal UMCInd As Long)
'---------------------------------------------------------------
'adds peak from unique mass class UMCInd to peaks to be searched
'looks for first peak before the maximum abundance peak that is
'under UMCNetAdjDef.PeakMaxAbuPct percentage of maximum abundance
'or first peak in the class if such was not found
'---------------------------------------------------------------
Dim HiAbuInd As Long                'index of highest abundance class member
Dim HiAbu As Double                 'highest abundance in class
Dim HiAbuScan As Long               'scan number of highest abundance peak
Dim AbuMax As Double                'abundance threshold
Dim CurrScan As Long
Dim BestInd As Long                 'index in UMC of peak that will be used
Dim BestScan As Long                'scan number of best peak(select first here)
Dim i As Long
On Error GoTo err_GetUMCPeak_Before
HiAbuInd = fUMCHiAbuInd(CallerID, UMCInd)
HiAbu = fUMCHiAbu(CallerID, UMCInd)
If HiAbuInd >= 0 Then
   With GelUMC(CallerID).UMCs(UMCInd)
       Select Case .ClassMType(HiAbuInd)
       Case glCSType
            HiAbuScan = GelData(CallerID).CSNum(.ClassMInd(HiAbuInd), csfScan)
       Case glIsoType
            HiAbuScan = GelData(CallerID).IsoNum(.ClassMInd(HiAbuInd), isfScan)
       End Select
       AbuMax = HiAbu * UMCNetAdjDef.PeakMaxAbuPct / 100
       BestInd = -1
       BestScan = -1
       For i = 0 To .ClassCount - 1
           If i <> HiAbuInd Then
              Select Case .ClassMType(i)
              Case glCSType
                If IsOKChargeState(GelData(CallerID).CSNum(.ClassMInd(i), csfFirstCS)) Then
                   CurrScan = GelData(CallerID).CSNum(.ClassMInd(i), csfScan)
                   If GelData(CallerID).CSNum(.ClassMInd(i), csfAbu) <= AbuMax Then
                      If CurrScan < HiAbuScan Then
                         If CurrScan > BestScan Then
                            BestInd = i
                            BestScan = CurrScan
                         End If
                      End If
                   End If
                End If
              Case glIsoType
                If IsOKChargeState(GelData(CallerID).IsoNum(.ClassMInd(i), isfCS)) Then
                   CurrScan = GelData(CallerID).IsoNum(.ClassMInd(i), isfScan)
                   If GelData(CallerID).IsoNum(.ClassMInd(i), isfAbu) <= AbuMax Then
                      If CurrScan < HiAbuScan Then
                         If CurrScan > BestScan Then
                            BestInd = i
                            BestScan = CurrScan
                         End If
                      End If
                   End If
                End If
              End Select
           End If
       Next i
       If BestInd > 0 Then
          PeakCnt = PeakCnt + 1
          PeakInd(PeakCnt - 1) = GelUMC(CallerID).UMCs(UMCInd).ClassMInd(BestInd)
          PeakType(PeakCnt - 1) = GelUMC(CallerID).UMCs(UMCInd).ClassMType(BestInd)
          PeakScan(PeakCnt - 1) = BestScan
          Select Case PeakType(PeakCnt - 1)
          Case glCSType
               PeakMW(PeakCnt - 1) = GelData(CallerID).CSNum(.ClassMInd(BestInd), csfMW)
          Case glIsoType
            ' MonroeMod: Now always using UMC Class MW rather than mass of most abundant member
            PeakMW(PeakCnt - 1) = .ClassMW
          End Select
          PeakUMCInd(PeakCnt - 1) = UMCInd
       Else                         'select first scan
          GetUMCPeak_First UMCInd
       End If
   End With
End If
Exit Sub

err_GetUMCPeak_Before:
If Err.Number = 9 Then
   ReDim Preserve PeakInd(PeakCnt + 100)
   ReDim Preserve PeakType(PeakCnt + 100)
   ReDim Preserve PeakMW(PeakCnt + 100)
   ReDim Preserve PeakScan(PeakCnt + 100)
   ReDim Preserve PeakUMCInd(PeakCnt + 100)
   Resume
End If
End Sub

Private Sub GetUMCPeak_At(ByVal UMCInd As Long)
'---------------------------------------------------------------
'adds peak from unique mass class UMCInd to peaks to be searched
'---------------------------------------------------------------
Dim HiAbuInd As Long
On Error GoTo err_GetUMCPeak_At
HiAbuInd = fUMCHiAbuInd(CallerID, UMCInd)
If HiAbuInd >= 0 Then
   PeakCnt = PeakCnt + 1
   With GelUMC(CallerID).UMCs(UMCInd)
        PeakInd(PeakCnt - 1) = .ClassMInd(HiAbuInd)
        PeakType(PeakCnt - 1) = .ClassMType(HiAbuInd)
        Select Case PeakType(PeakCnt - 1)
        Case glCSType
            PeakScan(PeakCnt - 1) = GelData(CallerID).CSNum(.ClassMInd(HiAbuInd), csfScan)
            PeakMW(PeakCnt - 1) = GelData(CallerID).CSNum(.ClassMInd(HiAbuInd), csfMW)
        Case glIsoType
            PeakScan(PeakCnt - 1) = GelData(CallerID).IsoNum(.ClassMInd(HiAbuInd), isfScan)
            ' MonroeMod: Now always using UMC Class MW rather than mass of most abundant member
            PeakMW(PeakCnt - 1) = .ClassMW
        End Select
   End With
   PeakUMCInd(PeakCnt - 1) = UMCInd
End If
Exit Sub

err_GetUMCPeak_At:
If Err.Number = 9 Then
   ReDim Preserve PeakInd(PeakCnt + 100)
   ReDim Preserve PeakType(PeakCnt + 100)
   ReDim Preserve PeakMW(PeakCnt + 100)
   ReDim Preserve PeakScan(PeakCnt + 100)
   ReDim Preserve PeakUMCInd(PeakCnt + 100)
   Resume
End If
End Sub

Public Function GetNETAdjustmentIDCount() As Long
    GetNETAdjustmentIDCount = IDCnt
End Function

Private Sub GetUMCPeak_After(ByVal UMCInd As Long)
'---------------------------------------------------------------
'adds peak from unique mass class UMCInd to peaks to be searched
'looks for first peak after the maximum abundance peak that is
'under UMCNetAdjDef.PeakMaxAbuPct percentage of maximum abundance
'or last peak in the class if such was not found
'---------------------------------------------------------------
Dim HiAbuInd As Long                'index of highest abundance class member
Dim HiAbu As Double                 'highest abundance in class
Dim HiAbuScan As Long               'scan number of highest abundance peak
Dim AbuMax As Double                'abundance threshold
Dim CurrScan As Long
Dim BestInd As Long                 'index in UMC of peak that will be used
Dim BestScan As Long                'scan number of best peak(select first here)
Dim i As Long
On Error GoTo err_GetUMCPeak_After
HiAbuInd = fUMCHiAbuInd(CallerID, UMCInd)
HiAbu = fUMCHiAbu(CallerID, UMCInd)
If HiAbuInd >= 0 Then
   With GelUMC(CallerID).UMCs(UMCInd)
       Select Case .ClassMType(HiAbuInd)
       Case glCSType
            HiAbuScan = GelData(CallerID).CSNum(.ClassMInd(HiAbuInd), csfScan)
       Case glIsoType
            HiAbuScan = GelData(CallerID).IsoNum(.ClassMInd(HiAbuInd), isfScan)
       End Select
       AbuMax = HiAbu * UMCNetAdjDef.PeakMaxAbuPct / 100
       BestInd = -1
       BestScan = -1
       For i = 0 To .ClassCount - 1
           If i <> HiAbuInd Then
              Select Case .ClassMType(i)
              Case glCSType
                If IsOKChargeState(GelData(CallerID).CSNum(.ClassMInd(i), csfFirstCS)) Then
                   CurrScan = GelData(CallerID).CSNum(.ClassMInd(i), csfScan)
                   If GelData(CallerID).CSNum(.ClassMInd(i), csfAbu) <= AbuMax Then
                      If CurrScan > HiAbuScan Then
                         If CurrScan < BestScan Then
                            BestInd = i
                            BestScan = CurrScan
                         End If
                      End If
                   End If
                End If
              Case glIsoType
                If IsOKChargeState(GelData(CallerID).IsoNum(.ClassMInd(i), isfCS)) Then
                   CurrScan = GelData(CallerID).IsoNum(.ClassMInd(i), isfScan)
                   If GelData(CallerID).IsoNum(.ClassMInd(i), isfAbu) <= AbuMax Then
                      If CurrScan > HiAbuScan Then
                         If CurrScan < BestScan Then
                            BestInd = i
                            BestScan = CurrScan
                         End If
                      End If
                   End If
                End If
              End Select
           End If
       Next i
       If BestInd > 0 Then
          PeakCnt = PeakCnt + 1
          PeakInd(PeakCnt - 1) = GelUMC(CallerID).UMCs(UMCInd).ClassMInd(BestInd)
          PeakType(PeakCnt - 1) = GelUMC(CallerID).UMCs(UMCInd).ClassMType(BestInd)
          PeakScan(PeakCnt - 1) = BestScan
          Select Case PeakType(PeakCnt - 1)
          Case glCSType
               PeakMW(PeakCnt - 1) = GelData(CallerID).CSNum(.ClassMInd(BestInd), csfMW)
          Case glIsoType
            ' MonroeMod: Now always using UMC Class MW rather than mass of most abundant member
            PeakMW(PeakCnt - 1) = .ClassMW
          End Select
          PeakUMCInd(PeakCnt - 1) = UMCInd
       Else                         'select first scan
          GetUMCPeak_Last UMCInd
       End If
   End With
End If
Exit Sub

err_GetUMCPeak_After:
If Err.Number = 9 Then
   ReDim Preserve PeakInd(PeakCnt + 100)
   ReDim Preserve PeakType(PeakCnt + 100)
   ReDim Preserve PeakMW(PeakCnt + 100)
   ReDim Preserve PeakScan(PeakCnt + 100)
   ReDim Preserve PeakUMCInd(PeakCnt + 100)
   Resume
End If
End Sub


Private Sub GetUMCPeak_First(ByVal UMCInd As Long)
Dim BestScan As Long        'last scan
Dim BestInd As Long         'index in UMCs
Dim CurrCS As Long
Dim CurrScan As Long
Dim i As Long
On Error GoTo err_GetUMCPeak_First

BestInd = -1
BestScan = 1000000
With GelUMC(CallerID).UMCs(UMCInd)
    For i = 0 To .ClassCount - 1
        Select Case .ClassMType(i)
        Case glCSType
             CurrCS = GelData(CallerID).CSNum(.ClassMInd(i), csfFirstCS)
             If IsOKChargeState(CurrCS) Then
                CurrScan = GelData(CallerID).CSNum(.ClassMInd(i), csfScan)
                If CurrScan < BestScan Then
                   BestScan = CurrScan
                   BestInd = i
                End If
             End If
        Case glIsoType
             CurrCS = GelData(CallerID).IsoNum(.ClassMInd(i), isfCS)
             If IsOKChargeState(CurrCS) Then
                CurrScan = GelData(CallerID).IsoNum(.ClassMInd(i), isfScan)
                If CurrScan < BestScan Then
                   BestScan = CurrScan
                   BestInd = i
                End If
             End If
        End Select
    Next i
    If BestInd >= 0 Then
       PeakCnt = PeakCnt + 1
       PeakInd(PeakCnt - 1) = .ClassMInd(BestInd)
       PeakType(PeakCnt - 1) = .ClassMType(BestInd)
       PeakScan(PeakCnt - 1) = BestScan
       Select Case PeakType(PeakCnt - 1)
       Case glCSType
            PeakMW(PeakCnt - 1) = GelData(CallerID).CSNum(.ClassMInd(BestInd), csfMW)
       Case glIsoType
            ' MonroeMod: Now always using UMC Class MW rather than mass of most abundant member
            PeakMW(PeakCnt - 1) = .ClassMW
       End Select
       PeakUMCInd(PeakCnt - 1) = UMCInd
    End If
End With
Exit Sub

err_GetUMCPeak_First:
If Err.Number = 9 Then
   ReDim Preserve PeakInd(PeakCnt + 100)
   ReDim Preserve PeakType(PeakCnt + 100)
   ReDim Preserve PeakMW(PeakCnt + 100)
   ReDim Preserve PeakScan(PeakCnt + 100)
   ReDim Preserve PeakUMCInd(PeakCnt + 100)
   Resume
End If
End Sub


Private Sub GetUMCPeak_Last(ByVal UMCInd As Long)
'---------------------------------------------------------------
'adds peak from unique mass class UMCInd to peaks to be searched
'---------------------------------------------------------------
Dim BestScan As Long        'last scan
Dim BestInd As Long         'index in UMCs
Dim CurrCS As Long
Dim CurrScan As Long
Dim i As Long
On Error GoTo err_GetUMCPeak_Last

BestInd = -1
BestScan = -1
With GelUMC(CallerID).UMCs(UMCInd)
    For i = 0 To .ClassCount - 1
        Select Case .ClassMType(i)
        Case glCSType
             CurrCS = GelData(CallerID).CSNum(.ClassMInd(i), csfFirstCS)
             If IsOKChargeState(CurrCS) Then
                CurrScan = GelData(CallerID).CSNum(.ClassMInd(i), csfScan)
                If CurrScan > BestScan Then
                   BestInd = i
                   BestScan = CurrScan
                End If
             End If
        Case glIsoType
             CurrCS = GelData(CallerID).IsoNum(.ClassMInd(i), isfCS)
             If IsOKChargeState(CurrCS) Then
                CurrScan = GelData(CallerID).IsoNum(.ClassMInd(i), isfScan)
                If CurrScan > BestScan Then
                   BestScan = CurrScan
                   BestInd = i
                End If
             End If
        End Select
    Next i
    If BestInd >= 0 Then
       PeakCnt = PeakCnt + 1
       PeakInd(PeakCnt - 1) = .ClassMInd(BestInd)
       PeakType(PeakCnt - 1) = .ClassMType(BestInd)
       PeakScan(PeakCnt - 1) = BestScan
       Select Case PeakType(PeakCnt - 1)
       Case glCSType
            PeakMW(PeakCnt - 1) = GelData(CallerID).CSNum(.ClassMInd(BestInd), csfMW)
       Case glIsoType
            ' MonroeMod: Now always using UMC Class MW rather than mass of most abundant member
            PeakMW(PeakCnt - 1) = .ClassMW
       End Select
       PeakUMCInd(PeakCnt - 1) = UMCInd
    End If
End With
Exit Sub

err_GetUMCPeak_Last:
If Err.Number = 9 Then
   ReDim Preserve PeakInd(PeakCnt + 100)
   ReDim Preserve PeakType(PeakCnt + 100)
   ReDim Preserve PeakMW(PeakCnt + 100)
   ReDim Preserve PeakScan(PeakCnt + 100)
   ReDim Preserve PeakUMCInd(PeakCnt + 100)
   Resume
End If
End Sub


Private Function IsOKChargeState(ByVal CS As Long) As Boolean
'--------------------------------------------------------------
'returns True if charge state is acceptable to current criteria
'--------------------------------------------------------------
On Error Resume Next
If UMCNetAdjDef.PeakCSSelection(7) Then
   IsOKChargeState = True
Else
   If CS >= 7 Then
      IsOKChargeState = UMCNetAdjDef.PeakCSSelection(6)
   Else
      IsOKChargeState = UMCNetAdjDef.PeakCSSelection(CS - 1)
   End If
End If
End Function


Private Function InitExprEvaluator(ByVal sExpr As String) As Boolean
'-------------------------------------------------------------------
'initializes expression evaluator for elution time
'-------------------------------------------------------------------
On Error Resume Next
Set NETExprEva = New ExprEvaluator
NETExprEva.Vars.add 1, "FN"
NETExprEva.Vars.add 2, "MinFN"
NETExprEva.Vars.add 3, "MaxFN"
NETExprEva.Expr = sExpr
InitExprEvaluator = NETExprEva.IsExprValid
ReDim VarVals(1 To 3)
VarVals(2) = ScanMin        'this will not change in this function
VarVals(3) = ScanMax
End Function

Private Sub CheckNETEquationStatus(Optional blnForceUsingDefaultGANETTrue As Boolean = False)

    If Not GelAnalysis(CallerID) Is Nothing Then
        If (txtNETFormula.Text = ConstructNETFormula(GelAnalysis(CallerID).GANET_Slope, GelAnalysis(CallerID).GANET_Intercept) Or blnForceUsingDefaultGANETTrue) _
           And InStr(UCase(txtNETFormula), "MINFN") = 0 Then
            mUsingDefaultGANET = True
        ElseIf blnForceUsingDefaultGANETTrue Then
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
        ConvertScanToNET = GetElution(lngScanNumber)
    End If

End Function

Private Function GetElution(FN As Long) As Double
'--------------------------------------------------
'this function does not care are we using NET or RT
'--------------------------------------------------
VarVals(1) = FN
GetElution = NETExprEva.ExprVal(VarVals())
End Function

Private Sub txtNETTol_Validate(Cancel As Boolean)
    TextBoxLimitNumberLength txtNETTol, 12
End Sub

Private Sub txtPctMaxAbu_LostFocus()
If IsNumeric(txtPctMaxAbu.Text) Then
   UMCNetAdjDef.PeakMaxAbuPct = Abs(CDbl(txtPctMaxAbu.Text))
Else
   MsgBox "This parameter should be positive number!", vbOKOnly, glFGTU
   txtPctMaxAbu.SetFocus
End If
End Sub

Private Sub txtUMCAbuTopPct_LostFocus()
If IsNumeric(txtUMCAbuTopPct.Text) Then
   UMCNetAdjDef.TopAbuPct = Abs(CDbl(txtUMCAbuTopPct.Text))
Else
   MsgBox "This parameter should be positive number!", vbOKOnly, glFGTU
   txtUMCAbuTopPct.SetFocus
End If
End Sub

Private Sub ReportIterationStep(ByVal IterationStep As Long)
'---------------------------------------------------------------
'add results of an iteration step on bottom of the rich text box
'---------------------------------------------------------------
Dim Rep As String
On Error Resume Next
Rep = "Iteration step: " & IterationStep & vbCrLf
Rep = Rep & "NET formula: " & UMCNetAdjDef.NETFormula & vbCrLf
Select Case UMCNetAdjDef.MWTolType
Case gltABS
    Rep = Rep & "MW Tol: " & UMCNetAdjDef.MWTol & " (Da) - NET Tol: "
Case gltPPM
    Rep = Rep & "MW Tol: " & UMCNetAdjDef.MWTol & " (ppm) - NET Tol: "
End Select
Rep = Rep & UMCNetAdjDef.NETTol & vbCrLf
Rep = Rep & "Peaks count: " & PeakCnt & " - IDs: " & IDCnt & vbCrLf
If AdjAvD >= 0 Then
   Rep = Rep & "Slope: " & AdjSlp & vbCrLf
   Rep = Rep & "Intercept: " & AdjInt & vbCrLf
   Rep = Rep & "Average Deviation: " & AdjAvD & vbCrLf
   Rep = Rep & "NET Min: " & Round(AdjSlp * ScanMin + AdjInt, 4) & ", NET Max: " & Round(AdjSlp * ScanMax + AdjInt, 4) & vbCrLf

   'Debug.Print IterationStep & ", " & IDCnt & "," & Round(AdjSlp * ScanMin + AdjInt, 5) & "," & Round(AdjSlp * ScanMax + AdjInt, 5)
Else
   Rep = "Error or insuficient information for NET adjustment calculation!" & vbCrLf
End If
Rep = Rep & vbCrLf
rtbIteReport.SelStart = Len(rtbIteReport.Text)
rtbIteReport.SelText = Rep

'If IterationStep = 1 Then Debug.Print
'Debug.Print IterationStep & vbTab & PeakCnt & vbTab & IDCnt & vbTab & AdjAvD
End Sub


Private Sub ReportIterationStep_Change(PrevNET1 As Double, PrevNET2 As Double, _
                                       CurrNET1 As Double, CurrNET2 As Double)
'-------------------------------------------------------------------------------
'add results of an iteration step on bottom of the rich text box
'-------------------------------------------------------------------------------
Dim Rep As String
On Error Resume Next
If IDCnt < 2 Then Rep = "Insufficient information for next iteration!" & vbCrLf
Rep = Rep & "Previous iteration NET min/max value: " & Format$(PrevNET1, "0.0000") _
       & "; " & Format$(PrevNET2, "0.0000") & vbCrLf
Rep = Rep & "Last iteration NET min/max value: " & Format$(CurrNET1, "0.0000") _
       & "; " & Format$(CurrNET2, "0.0000") & vbCrLf
rtbIteReport.SelStart = Len(rtbIteReport.Text)
rtbIteReport.SelText = Rep
End Sub


Private Sub ReportStop()
Dim Rep As String
On Error Resume Next
Rep = "Iteration cancelled by user!" & vbCrLf
rtbIteReport.SelStart = Len(rtbIteReport.Text)
rtbIteReport.SelText = Rep
End Sub

Private Sub ReportPause()
Dim Rep As String
On Error Resume Next
Rep = "Iteration paused by user! Press Continue to resume" & vbCrLf
rtbIteReport.SelStart = Len(rtbIteReport.Text)
rtbIteReport.SelText = Rep
UpdateStatus "Paused"
End Sub


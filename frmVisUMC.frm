VERSION 5.00
Begin VB.Form frmVisUMC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UMC Editor"
   ClientHeight    =   9480
   ClientLeft      =   150
   ClientTop       =   825
   ClientWidth     =   11835
   Icon            =   "frmVisUMC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowUMC 
      Caption         =   "Show Unique Mass Classes"
      Height          =   255
      Left            =   4680
      TabIndex        =   62
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Editor Functions/Options"
      Height          =   9375
      Left            =   7320
      TabIndex        =   0
      Top             =   0
      Width           =   4440
      Begin VB.Frame fraAutoRemove 
         Caption         =   "Auto-Remove Options"
         Height          =   3195
         Left            =   120
         TabIndex        =   43
         Top             =   6060
         Width           =   4215
         Begin VB.TextBox txtPercentMaxAbuToUseToGaugeLength 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   59
            Text            =   "25"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtAutoRefineMinimumMemberCount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            TabIndex        =   64
            Text            =   "3"
            Top             =   2160
            Width           =   495
         End
         Begin VB.CheckBox chkRefineUMCLengthByScanRange 
            Caption         =   "Test UMC length using scan range"
            Height          =   375
            Left            =   120
            TabIndex        =   63
            ToolTipText     =   "If True, then considers scan range for the length tests; otherwise, considers member count"
            Top             =   2040
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkRemovePairedHUMC 
            Caption         =   "Remove classes paired as heavy members"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   2880
            Width           =   3615
         End
         Begin VB.CheckBox chkRemovePairedLUMC 
            Caption         =   "Remove classes paired as light members"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   2565
            Width           =   3615
         End
         Begin VB.TextBox txtHiCnt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   56
            Text            =   "500"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CheckBox chkRemoveHiCnt 
            Caption         =   "Remove cls. with more than"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.TextBox txtLoCnt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   53
            Text            =   "3"
            Top             =   960
            Width           =   495
         End
         Begin VB.CheckBox chkRemoveLoCnt 
            Caption         =   "Remove cls. with less than"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtHiAbuPct 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3000
            TabIndex        =   51
            Text            =   "30"
            Top             =   600
            Width           =   615
         End
         Begin VB.CheckBox chkRemoveHiAbu 
            Caption         =   "Remove high intensity classes(%)"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtLoAbuPct 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3000
            TabIndex        =   49
            Text            =   "30"
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox chkRemoveLoAbu 
            Caption         =   "Remove low intensity classes(%)"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.Label lblPercentMaxAbuToUseToGaugeLength 
            Caption         =   "Percent max abundance to use to gauge width"
            Height          =   405
            Left            =   600
            TabIndex        =   58
            Top             =   1560
            Width           =   1845
         End
         Begin VB.Label lblAutoRefineMinimumMemberCount 
            Caption         =   "Minimum member count:"
            Height          =   375
            Left            =   2160
            TabIndex        =   65
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label lblAutoRefineLengthLabel 
            Caption         =   "members"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   57
            Top             =   1360
            Width           =   1000
         End
         Begin VB.Label lblAutoRefineLengthLabel 
            Caption         =   "members"
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   54
            Top             =   1000
            Width           =   1000
         End
      End
      Begin VB.Frame fraAutoMerge 
         Caption         =   "Auto-Merge Options"
         Height          =   1335
         Left            =   120
         TabIndex        =   38
         Top             =   4620
         Width           =   3975
         Begin VB.TextBox txtAutoMergeMaxMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   42
            Text            =   "10000"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtAutoMergeMinMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   41
            Text            =   "2000"
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox cmbAutoMerge 
            Height          =   315
            ItemData        =   "frmVisUMC.frx":030A
            Left            =   240
            List            =   "frmVisUMC.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label5 
            Caption         =   "UMC MW merge range"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   900
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkMultiMemberGroups 
         Caption         =   "List only multi-member groups"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.Frame fraUMCMerger 
         Caption         =   "Selection List"
         Height          =   2055
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   3975
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   325
            Left            =   2940
            TabIndex        =   47
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdDelClass 
            Caption         =   "Del. Class"
            Height          =   325
            Left            =   1980
            TabIndex        =   46
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAddGroup 
            Caption         =   "Add Group"
            Height          =   325
            Left            =   1020
            TabIndex        =   45
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAddClass 
            Caption         =   "Add Class"
            Height          =   325
            Left            =   60
            TabIndex        =   44
            Top             =   1680
            Width           =   975
         End
         Begin VB.ListBox lstSelectedClasses 
            Height          =   1425
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.ComboBox cmbLstGroups 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmVisUMC.frx":030E
         Left            =   120
         List            =   "frmVisUMC.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Frame fraFunction1 
         Caption         =   "Grouping Definition"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtMultiplicity 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   35
            Text            =   "100"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtf1MWTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   26
            Text            =   "0.02"
            Top             =   300
            Width           =   495
         End
         Begin VB.TextBox txtf1MWDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   3
            Text            =   "50"
            Top             =   300
            Width           =   495
         End
         Begin VB.TextBox txtf1ScanDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   2
            Text            =   "5"
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Multiplicity"
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   34
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "MW tol."
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   25
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "MW distance"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Max. scan separation"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   660
            Width           =   1575
         End
      End
      Begin VB.Label lblGroupsCount 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Groups count: "
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   29
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Groups Selection"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   7155
      TabIndex        =   7
      Top             =   0
      Width           =   7215
      Begin VB.PictureBox picDummy1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   7215
         Index           =   1
         Left            =   6840
         ScaleHeight     =   7215
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
      Begin VB.ListBox lstPeaks 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   2550
         Left            =   2400
         TabIndex        =   31
         Top             =   5400
         Width           =   4695
      End
      Begin VB.PictureBox picDummy1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   0
         Left            =   2040
         ScaleHeight     =   7335
         ScaleWidth      =   285
         TabIndex        =   27
         Top             =   960
         Width           =   280
      End
      Begin VB.ListBox lstClasses 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3810
         Left            =   2520
         TabIndex        =   21
         Top             =   1200
         Width           =   4575
      End
      Begin VB.ListBox lstGroups 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   6855
         Left            =   0
         TabIndex        =   19
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblUMCCalculator 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5040
         TabIndex        =   36
         Top             =   120
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Classes belonging to a group"
         ForeColor       =   &H000080FF&
         Height          =   225
         Index           =   1
         Left            =   2880
         TabIndex        =   22
         Top             =   960
         Width           =   2070
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Groups of classes"
         ForeColor       =   &H000080FF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total number of peaks:"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Original UMC count:"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current UMC count:"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblPeaksCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblOrigUMCCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCurrUMCCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Original UMC ratio:"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current UMC ratio:"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblOrigUMCRatio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCurrUMCRatio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Class members"
         ForeColor       =   &H000080FF&
         Height          =   225
         Index           =   2
         Left            =   2880
         TabIndex        =   8
         Top             =   5160
         Width           =   1080
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   8640
      Width           =   4455
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTUMC 
         Caption         =   "UMC Calculator"
         Begin VB.Menu mnuTUMCAvgMW 
            Caption         =   "Avg. MW"
         End
         Begin VB.Menu mnuTUMCRepMW 
            Caption         =   "Rep. MW"
         End
         Begin VB.Menu mnuTUMCStDevMW 
            Caption         =   "StDev. MW"
         End
         Begin VB.Menu mnuTUMCAvgAbu 
            Caption         =   "Avg. Abu."
         End
         Begin VB.Menu mnuTUMCRepAbu 
            Caption         =   "Rep. Abu."
         End
         Begin VB.Menu mnuTUMCSumAbu 
            Caption         =   "Sum. Abu."
         End
      End
      Begin VB.Menu mnuTG 
         Caption         =   "Group"
         Begin VB.Menu mnuTGroup 
            Caption         =   "Mass(ppm)/Scan Proximity"
            Index           =   0
         End
         Begin VB.Menu mnuTGroup 
            Caption         =   "Mass(Da)/Scan Exact Distance"
            Index           =   1
         End
         Begin VB.Menu mnuTGroup 
            Caption         =   "Mass(ppm) Equivalency"
            Index           =   2
         End
         Begin VB.Menu mnuTGroup 
            Caption         =   "Members Sharing Equivalency"
            Index           =   3
         End
         Begin VB.Menu mnuTGSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTGReport 
            Caption         =   "Report"
         End
      End
      Begin VB.Menu mnuTM 
         Caption         =   "Merge"
         Begin VB.Menu mnuTMMergeSelClasses 
            Caption         =   "Merge Selected Classes"
         End
         Begin VB.Menu mnuTMMergeGroup 
            Caption         =   "Merge Group"
         End
         Begin VB.Menu mnuTMMergeAll 
            Caption         =   "Merge All"
         End
         Begin VB.Menu mnuTMSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTMAutoMerge 
            Caption         =   "Auto Merge"
         End
      End
      Begin VB.Menu mnuTR 
         Caption         =   "Remove"
         Begin VB.Menu mnuTRRemoveSelClasses 
            Caption         =   "Remove Selected Classes"
         End
         Begin VB.Menu mnuTRRemoveClass 
            Caption         =   "Remove Class"
         End
         Begin VB.Menu mnuTRRemoveGroup 
            Caption         =   "Remove Group"
         End
         Begin VB.Menu mnuTRSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTRAutoRemove 
            Caption         =   "Auto Remove"
         End
      End
   End
End
Attribute VB_Name = "frmVisUMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Unique Mass Classes Editor - allows inspection, analysis and redefinition of UMC
'created: 04/08/2002 nt
'last modified: 04/04/2003 nt
'--------------------------------------------------------------------------------
Option Explicit

Const MAX_GROUPS_IN_LIST = 250

Const F_MW_SCAN_PROXIMITY = 0
Const F_MW_SCAN_EXACT_DISTANCE = 1
Const F_MW_EQUIVALENCY = 2
Const F_MEMBERS_SHARING_EQUIVALENCY = 3

Const AUTOMERGE_HI_ABU = 0
Const AUTOMERGE_LO_ABU = 1
Const AUTOMERGE_HI_CNT = 2
Const AUTOMERGE_LO_CNT = 3
Const AUTOMERGE_HI_MW = 4
Const AUTOMERGE_LO_MW = 5

Const REMOVE_UMC_MARK = -1

'used when calling UpdateLabels function
Const LBL_ERR = -1
Const LBL_ZLS = 0

Dim bLoading As Long
Dim CallerID As Long

Dim tmp As UMCListType              'all work on this form is done with temporary UMC
Dim TmpInc() As Long        'array parallel with UMCs in Tmp used to determine
                            'what will be included in newly defined UMC
                        
Dim UMCStat() As Double     'precalculated classes statistics used to easily
                            'access class properties and display descriptions
                            
Dim UMCDisplay() As String  'class descriptions used to list classes
Dim UMCInd() As Long        'original sort order
Dim UMCMW() As Double       'classes will be sorted on molecular masses

Dim MWSearch As New MWUtil  'mass range search object

'results of analysis
Dim GrRes As GR2            'result is a group of group of classes

'if number of results is more than MAX_GROUPS_IN_LIST results are
'split in groups which then can be selected/displayed from combo box
Dim LstGroupsCnt As Long
Dim LstGroupsInd1() As Long             'first index of groups belonging to list
Dim LstGroupsInd2() As Long             'last index of groups belonging to list
Dim LstGroupsDisplay() As String        'display name of list portion

Dim CurrLstGroupInd As Long             'index of selected item in the groups combo box
Dim CurrGroupInd As Long                'index of selected group in Res structure
'NOTE: CurrGroupInd=CurrLstGroupInd*MAX_GROUPS_IN_LIST+CurrLstInd
Dim CurrClassInd As Long                'index of selected class
Dim CurrPeakInd As Long                 'index of selected peak

'parameters of grouping
'NOTE: not all parameters are used with all types of grouping, and even same
'parameter could have different interpretation with different functions
Dim f1Type As Long          'type of grouping
Dim f1MWDist As Double      'mw distance(measurement unit depends on type)
Dim f1MWTol As Double       'molecular mass tolerance when needed
Dim f1ScanDist As Long      'scan distance (scan distance 0 means classes must overlap)
Dim f1Multiplicity As Long  'used as an extra parameter (reporting)

'auto merge variables
Dim amType As Long                  'type of auto-range
Dim amMWMin As Long                 'minimum of mass range to merge
Dim amMWMax As Long                 'maximum of mass range to merge

Dim MultiGroupsOnly As Boolean      'if True only groups with multiple membership
                                    'will be displayed
                                    
Dim NeedToSave As Boolean       'if this flag is set ask user does it want to save

Dim PairsUMCInSync As Boolean   'need to be careful with pairs if not in sync with UMC

Dim WithEvents MyViewer As frmUMCView
Attribute MyViewer.VB_VarHelpID = -1

Private mChangeList As String

Public Sub AutoRemoveUMCsWork()
'-------------------------------------------------------
'mark all classes satisfying conditions for removal
'-------------------------------------------------------
Dim i As Long, j As Long
Dim lngScan As Long
Dim arrAbu() As Double
Dim arrInd() As Long
Dim qsAbu As QSDouble
Dim PctCnt As Long
Dim lngOriginalUMCCount As Long
Dim strProcessSummary As String
Dim udtAutoRefine As udtUMCAutoRefineOptionsType

Dim lngMemberIndex As Long
Dim lngScanStart As Long
Dim lngScanEnd As Long
Dim lngCurrentScan As Long

Dim sngThreshold As Single

On Error GoTo AutoRemoveUMCsWorkErrorHandler

' Copy the .UMCAutoRefineOptions to a local variable to make the code a little cleaner (shorter variable names)
udtAutoRefine = glbPreferencesExpanded.UMCAutoRefineOptions

If udtAutoRefine.TestLengthUsingScanRange Then
    ' Need to determine the minimum and maximum scan of each UMC in Tmp
    With tmp
        If .UMCCnt > 0 Then
           For i = 0 To .UMCCnt - 1
               With .UMCs(i)
                   .MinScan = glHugeLong:               .MaxScan = -glHugeLong
                   If .ClassCount > 0 Then
                      For j = 0 To .ClassCount - 1
                          Select Case .ClassMType(j)
                          Case glCSType
                               lngScan = GelData(CallerID).CSData(.ClassMInd(j)).ScanNumber
                          Case glIsoType
                               lngScan = GelData(CallerID).IsoData(.ClassMInd(j)).ScanNumber
                          End Select
                          If lngScan < .MinScan Then .MinScan = lngScan
                          If lngScan > .MaxScan Then .MaxScan = lngScan
                      Next j
                   End If
               End With
               If .UMCs(i).ClassCount > 0 Then
               Else                             'something wrong
                  .UMCs(i).MinScan = -1
                  .UMCs(i).MaxScan = -1
               End If
           Next i
        End If
    End With
End If

UpdateStatus "Selecting classes to remove..."
Me.MousePointer = vbHourglass
With tmp
    lngOriginalUMCCount = .UMCCnt
    If udtAutoRefine.UMCAutoRefineRemoveAbundanceHigh Then
       UpdateStatus "Eliminating high abundance classes ..."
       PctCnt = CLng(.UMCCnt * udtAutoRefine.UMCAutoRefinePctHighAbundance / 100)
       ReDim arrAbu(.UMCCnt - 1)
       ReDim arrInd(.UMCCnt - 1)
       For i = 0 To .UMCCnt - 1
           arrInd(i) = i
           arrAbu(i) = .UMCs(i).ClassAbundance
       Next i
       Set qsAbu = New QSDouble
       If qsAbu.QSDesc(arrAbu(), arrInd()) Then     'eliminate first PctCnt classes
          For i = 0 To PctCnt - 1
              TmpInc(arrInd(i)) = REMOVE_UMC_MARK
          Next i
       End If
       strProcessSummary = "Removed high abundance classes (" & Trim(udtAutoRefine.UMCAutoRefinePctHighAbundance) & "% removed)"
    End If
    If udtAutoRefine.UMCAutoRefineRemoveAbundanceLow Then
       UpdateStatus "Eliminating low abundance classes ..."
       PctCnt = CLng(.UMCCnt * udtAutoRefine.UMCAutoRefinePctLowAbundance / 100)
       ReDim arrAbu(.UMCCnt - 1)
       ReDim arrInd(.UMCCnt - 1)
       For i = 0 To .UMCCnt - 1
           arrInd(i) = i
           arrAbu(i) = .UMCs(i).ClassAbundance
       Next i
       Set qsAbu = New QSDouble
       If qsAbu.QSAsc(arrAbu(), arrInd()) Then     'eliminate first PctCnt classes
          For i = 0 To PctCnt - 1
              TmpInc(arrInd(i)) = REMOVE_UMC_MARK
          Next i
       End If
       If Len(strProcessSummary) > 0 Then strProcessSummary = strProcessSummary & "; "
       strProcessSummary = strProcessSummary & "Removed low abundance classes (" & Trim(udtAutoRefine.UMCAutoRefinePctLowAbundance) & "% removed)"
    End If
    If udtAutoRefine.UMCAutoRefineRemoveCountLow Then
       If Len(strProcessSummary) > 0 Then strProcessSummary = strProcessSummary & "; "
       If Not udtAutoRefine.TestLengthUsingScanRange Then
            UpdateStatus "Removing classes with too few members ..."
            For i = 0 To .UMCCnt - 1
                If .UMCs(i).ClassCount < udtAutoRefine.UMCAutoRefineMinLength Then TmpInc(i) = REMOVE_UMC_MARK
            Next i
            strProcessSummary = strProcessSummary & "Removed classes with low member count (count < " & Trim(udtAutoRefine.UMCAutoRefineMinLength) & ")"
       Else
            UpdateStatus "Removing classes with too short of a scan range ..."
            For i = 0 To .UMCCnt - 1
                If .UMCs(i).MaxScan - .UMCs(i).MinScan + 1 < udtAutoRefine.UMCAutoRefineMinLength Or _
                   .UMCs(i).ClassCount < udtAutoRefine.MinMemberCountWhenUsingScanRange Then
                    TmpInc(i) = REMOVE_UMC_MARK
                End If
            Next i
            strProcessSummary = strProcessSummary & "Removed classes with too short of a scan range (range < " & Trim(udtAutoRefine.UMCAutoRefineMinLength) & " scans) or too few members (count < " & Trim(udtAutoRefine.MinMemberCountWhenUsingScanRange) & ")"
       End If
    End If
    If udtAutoRefine.UMCAutoRefineRemoveCountHigh Then
       If Len(strProcessSummary) > 0 Then strProcessSummary = strProcessSummary & "; "
       If Not udtAutoRefine.TestLengthUsingScanRange Then
            UpdateStatus "Removing classes with too many members ..."
            For i = 0 To .UMCCnt - 1
                If .UMCs(i).ClassCount > udtAutoRefine.UMCAutoRefineMaxLength Then TmpInc(i) = REMOVE_UMC_MARK
            Next i
            strProcessSummary = strProcessSummary & "Removed classes with high member count (count > " & Trim(udtAutoRefine.UMCAutoRefineMaxLength) & ")"
       Else
            UpdateStatus "Removing classes with too long of a scan range ..."
            For i = 0 To .UMCCnt - 1
                If udtAutoRefine.UMCAutoRefinePercentMaxAbuToUseForLength > 0 Then
                    ' Determine the first and last scan numbers in the UMC that have data with an intensity value >= the MaximumIntensity * .UMCAutoRefinePercentMaxAbuToUseForLength / 100
                    With .UMCs(i)
                        ' First, determine the maximum intensity
                        Select Case .ClassRepType
                        Case glCSType
                            sngThreshold = GelData(CallerID).CSData(.ClassRepInd).Abundance * udtAutoRefine.UMCAutoRefinePercentMaxAbuToUseForLength / 100
                            lngScanStart = GelData(CallerID).CSData(.ClassRepInd).ScanNumber
                        Case glIsoType
                            sngThreshold = GelData(CallerID).IsoData(.ClassRepInd).Abundance * udtAutoRefine.UMCAutoRefinePercentMaxAbuToUseForLength / 100
                            lngScanStart = GelData(CallerID).IsoData(.ClassRepInd).ScanNumber
                        End Select
                        lngScanEnd = lngScanStart

                        ' Now examine the data to find the first scan with a data point whose abundance is >= sngThreshold
                        For lngMemberIndex = 0 To .ClassCount - 1
                            
                            Select Case .ClassMType(lngMemberIndex)
                            Case glCSType
                                lngCurrentScan = GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).ScanNumber
                                If GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).Abundance >= sngThreshold And lngCurrentScan < lngScanStart Then
                                    lngScanStart = lngCurrentScan
                                End If
                            Case glIsoType
                                lngCurrentScan = GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)).ScanNumber
                                If GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)).Abundance >= sngThreshold And lngCurrentScan < lngScanStart Then
                                    lngScanStart = lngCurrentScan
                                End If
                            End Select
                        Next lngMemberIndex

                        ' Now examine the data to find the last scan with a data point whose abundance is >= sngThreshold
                        For lngMemberIndex = .ClassCount - 1 To 0 Step -1
                            Select Case .ClassMType(lngMemberIndex)
                            Case glCSType
                                lngCurrentScan = GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).ScanNumber
                                If GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).Abundance >= sngThreshold And lngCurrentScan > lngScanEnd Then
                                    lngScanEnd = lngCurrentScan
                                End If
                            Case glIsoType
                                lngCurrentScan = GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)).ScanNumber
                                If GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)).Abundance >= sngThreshold And lngCurrentScan > lngScanEnd Then
                                    lngScanEnd = lngCurrentScan
                                End If
                            End Select
                        Next lngMemberIndex
                        
                        ' We can now examine the total width of the UMC
                        If lngScanEnd - lngScanStart + 1 > udtAutoRefine.UMCAutoRefineMaxLength Then
                            TmpInc(i) = REMOVE_UMC_MARK
                        End If
                    End With
                    
                Else
                    ' Use the entire length of the UMC to gauge the total scans
                    If .UMCs(i).MaxScan - .UMCs(i).MinScan + 1 > udtAutoRefine.UMCAutoRefineMaxLength Then
                        TmpInc(i) = REMOVE_UMC_MARK
                    End If
                End If
            Next i
            strProcessSummary = strProcessSummary & "Removed classes with too long of a scan range (range > " & Trim(udtAutoRefine.UMCAutoRefineMaxLength) & " scans); "
            strProcessSummary = strProcessSummary & "Max abundance to use to gauge width = " & Trim(udtAutoRefine.UMCAutoRefinePercentMaxAbuToUseForLength) & "%"
            If udtAutoRefine.UMCAutoRefinePercentMaxAbuToUseForLength = 0 Then
                strProcessSummary = strProcessSummary & " (used full width of the UMC to gauge width)"
            End If
            
       End If
    End If
    With GelP_D_L(CallerID)
       If chkRemovePairedLUMC.value = vbChecked Then
          If .PCnt > 0 And PairsUMCInSync Then
             For i = 0 To .PCnt - 1
                 TmpInc(.Pairs(i).P1) = REMOVE_UMC_MARK
             Next i
          End If
          If Len(strProcessSummary) > 0 Then strProcessSummary = strProcessSummary & "; "
          strProcessSummary = strProcessSummary & "Removed classes paired as light members"
       End If
       If chkRemovePairedHUMC.value = vbChecked Then
          If .PCnt > 0 And PairsUMCInSync Then
             For i = 0 To .PCnt - 1
                 TmpInc(.Pairs(i).P2) = REMOVE_UMC_MARK
             Next i
          End If
          If Len(strProcessSummary) > 0 Then strProcessSummary = strProcessSummary & "; "
          strProcessSummary = strProcessSummary & "Removed classes paired as heavy members"
       End If
    End With
End With

UpdateStatus "Recalculating class structure..."
Call ClearGroupsAndLists
Call RemoveClasses
Call ResetUMC
AddToTentativeChangeList "Auto-removed UMC's; Original UMC Count = " & Trim(lngOriginalUMCCount) & "; Count after removal = " & Trim(tmp.UMCCnt) & "; " & strProcessSummary
NeedToSave = True
UpdateStatus ""
Me.MousePointer = vbDefault

Exit Sub

AutoRemoveUMCsWorkErrorHandler:
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "Error in AutoRemoveUMCsWork: " & Err.Description, vbExclamation Or vbOKOnly, "Error"
Else
    LogErrors Err.Number, "AutoRemoveUMCsWork"
End If

End Sub

Public Sub InitializeUMCs()
On Error GoTo err_Activate
   UpdateStatus "Loading..."
   CallerID = Me.Tag
   tmp = GelUMC(CallerID)
   If tmp.UMCCnt <= 0 Then
      MsgBox "No unique mass classes found.", vbOKOnly, glFGTU
      Unload Me
      Exit Sub
   End If
   
   lblPeaksCnt.Caption = GelData(CallerID).CSLines + GelData(CallerID).IsoLines
   lblOrigUMCCnt.Caption = GelUMC(CallerID).UMCCnt
   lblOrigUMCRatio.Caption = Format$(lblOrigUMCCnt.Caption / lblPeaksCnt.Caption, "0.00")
   PairsUMCInSync = GelP_D_L(CallerID).SyncWithUMC
   UpdateStatus "Preparing analysis..."
   If Not ResetUMC() Then GoTo err_Activate
   UpdateStatus ""
   mChangeList = ""
   bLoading = False
   If GelP_D_L(CallerID).PCnt > 0 And PairsUMCInSync Then
      chkRemovePairedLUMC.Enabled = True
      chkRemovePairedHUMC.Enabled = True
   End If
   Set MyViewer = New frmUMCView
   MyViewer.CallerID = CallerID

   Exit Sub
err_Activate:
UpdateStatus "Error preparing analysis..."

End Sub

Private Sub FillComboBoxes()
    With cmbAutoMerge
        .Clear
        .AddItem "Prefer higher UMC abundance"
        .AddItem "Prefer lower UMC abundance"
        .AddItem "Prefer higher UMC count"
        .AddItem "Prefer lower UMC count"
        .AddItem "Prefer higher MW"
        .AddItem "Prefer lower MW"
    End With
    
End Sub

Private Sub chkMultiMemberGroups_Click()
MultiGroupsOnly = (chkMultiMemberGroups.value = vbChecked)
End Sub

Private Sub chkRefineUMCLengthByScanRange_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange = cChkBox(chkRefineUMCLengthByScanRange)
    UpdateDynamicControls
End Sub

Private Sub chkRemoveHiAbu_Click()
glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveAbundanceHigh = cChkBox(chkRemoveHiAbu)
End Sub

Private Sub chkRemoveHiCnt_Click()
glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveCountHigh = cChkBox(chkRemoveHiCnt)
End Sub

Private Sub chkRemoveLoAbu_Click()
glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveAbundanceLow = cChkBox(chkRemoveLoAbu)
End Sub

Private Sub chkRemoveLoCnt_Click()
glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveCountLow = cChkBox(chkRemoveLoCnt)
End Sub

Private Sub chkShowUMC_Click()
If chkShowUMC.value = vbChecked Then
   MyViewer.show
Else
   MyViewer.Hide
End If
End Sub

Private Sub cmbAutoMerge_Click()
amType = cmbAutoMerge.ListIndex
End Sub

Private Sub cmbLstGroups_Click()
Dim i As Long
On Error Resume Next
CurrLstGroupInd = cmbLstGroups.ListIndex
lstGroups.Clear
lstClasses.Clear
lstPeaks.Clear
For i = LstGroupsInd1(CurrLstGroupInd) To LstGroupsInd2(CurrLstGroupInd)
    lstGroups.AddItem GrRes.Members(i).Description
Next i
End Sub

Private Sub cmdAddClass_Click()
'-------------------------------------------------------------
'adds currently selected class to the list of selected classes
'-------------------------------------------------------------
If CurrClassInd >= 0 Then
   With lstSelectedClasses
        .AddItem UMCDisplay(CurrClassInd)
        .ItemData(.NewIndex) = CurrClassInd
   End With
Else
   MsgBox "No class currently selected.", vbOKOnly, glFGTU
End If
End Sub

Private Sub cmdAddGroup_Click()
'-----------------------------------------------------------------------------
'add all classes belonging to the selected group in a list of selected classes
'-----------------------------------------------------------------------------
Dim i As Long
On Error Resume Next
If CurrGroupInd >= 0 Then
   With GrRes.Members(CurrGroupInd)
        For i = 0 To .Count - 1
            lstSelectedClasses.AddItem UMCDisplay(.Members(i))
            lstSelectedClasses.ItemData(lstSelectedClasses.NewIndex) = .Members(i)
        Next i
   End With
Else
   MsgBox "No group selected.", vbOKOnly, glFGTU
End If
End Sub

Private Sub cmdClear_Click()
lstSelectedClasses.Clear
End Sub

Private Sub cmdDelClass_Click()
'--------------------------------------------------------
'removes selected class from the list of selected classes
'--------------------------------------------------------
With lstSelectedClasses
    If .ListIndex >= 0 Then
       .RemoveItem .ListIndex
    Else
       MsgBox "No class selected.", vbOKOnly, glFGTU
    End If
End With
End Sub

Private Sub Form_Activate()
If bLoading Then
   InitializeUMCs
End If
End Sub

Private Sub Form_Load()
FillComboBoxes
Me.Move 100, 100
If IsWinLoaded(TrackerCaption) Then frmTracker.Visible = False
DoEvents
bLoading = True
f1MWDist = txtf1MWDist.Text
f1ScanDist = txtf1ScanDist.Text
f1MWTol = txtf1MWTol.Text
f1Multiplicity = txtMultiplicity.Text
MultiGroupsOnly = chkMultiMemberGroups.value
CurrGroupInd = -1
CurrClassInd = -1
CurrPeakInd = -1
amType = -1
amMWMin = txtAutoMergeMinMW.Text
amMWMax = txtAutoMergeMaxMW.Text

With glbPreferencesExpanded.UMCAutoRefineOptions
    SetCheckBox chkRemoveLoCnt, .UMCAutoRefineRemoveCountLow
    SetCheckBox chkRemoveHiCnt, .UMCAutoRefineRemoveCountHigh
    txtLoCnt = .UMCAutoRefineMinLength
    txtHiCnt = .UMCAutoRefineMaxLength
    txtPercentMaxAbuToUseToGaugeLength = .UMCAutoRefinePercentMaxAbuToUseForLength
    
    SetCheckBox chkRefineUMCLengthByScanRange, .TestLengthUsingScanRange
    txtAutoRefineMinimumMemberCount = .MinMemberCountWhenUsingScanRange
    UpdateDynamicControls
    
    SetCheckBox chkRemoveLoAbu, .UMCAutoRefineRemoveAbundanceLow
    SetCheckBox chkRemoveHiAbu, .UMCAutoRefineRemoveAbundanceHigh
    txtLoAbuPct = .UMCAutoRefinePctLowAbundance
    txtHiAbuPct = .UMCAutoRefinePctHighAbundance
End With

End Sub

Private Function GetIncCount() As Long
'---------------------------------------------
'included distributions are all with TmpInc>=0
'---------------------------------------------
Dim i As Long
Dim Cnt As Long
On Error Resume Next
For i = 0 To tmp.UMCCnt - 1
    If TmpInc(i) >= 0 Then Cnt = Cnt + 1
Next i
GetIncCount = Cnt
End Function

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub UpdateDynamicControls()
    ' Update the UMC auto refine length labels
    If glbPreferencesExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange Then
        chkRemoveLoCnt.Caption = "Remove classes less than"
        chkRemoveHiCnt.Caption = "Remove classes more than"
        lblAutoRefineLengthLabel(0) = "scans wide"
        lblAutoRefineLengthLabel(1) = "scans wide"
        lblAutoRefineMinimumMemberCount.Enabled = True
    Else
        chkRemoveLoCnt.Caption = "Remove cls. with less than"
        chkRemoveHiCnt.Caption = "Remove cls. with more than"
        lblAutoRefineLengthLabel(0) = "members"
        lblAutoRefineLengthLabel(1) = "members"
        lblAutoRefineMinimumMemberCount.Enabled = False
    End If

    txtAutoRefineMinimumMemberCount.Enabled = lblAutoRefineMinimumMemberCount.Enabled
    lblPercentMaxAbuToUseToGaugeLength.Enabled = lblAutoRefineMinimumMemberCount.Enabled
    txtPercentMaxAbuToUseToGaugeLength.Enabled = lblAutoRefineMinimumMemberCount.Enabled
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Res As Long
On Error Resume Next
If NeedToSave Then
   If glbPreferencesExpanded.AutoAnalysisStatus.AutoRefiningUMCs Then
      Res = vbYes
   Else
      Res = MsgBox("Do you want to save the UMC changes?", vbYesNoCancel, glFGTU)
   End If
   Select Case Res
   Case vbYes                               'if Yes save and unload
      UpdateStatus "Saving ..."
      GelUMC(CallerID) = tmp
      
      ' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
      UpdateUMCStatArrays CallerID, False, Me
      
      GelP_D_L(CallerID).SyncWithUMC = PairsUMCInSync
      AddToAnalysisHistory CallerID, mChangeList, False
   Case vbCancel                            'Cancel unload
      Cancel = True
   End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
NeedToSave = False
Erase tmp.UMCs
tmp.UMCCnt = 0
Unload MyViewer
Set MyViewer = Nothing
If IsWinLoaded(TrackerCaption) Then frmTracker.Visible = True
End Sub

Private Sub lstClasses_Click()
Dim FirstScan As Long, LastScan As Long
Dim MinMW As Double, MaxMW As Double
Dim CurrLstInd As Long
On Error GoTo LstClassesErrorHandler
lblUMCCalculator.Caption = ""
CurrLstInd = lstClasses.ListIndex
If CurrLstInd >= 0 Then
   CurrClassInd = GrRes.Members(CurrGroupInd).Members(CurrLstInd)
   If CurrClassInd >= 0 Then
      ListPeaksForClass CurrClassInd
      If MyViewer.Visible Then             'jump to proper region
         FirstScan = tmp.UMCs(CurrClassInd).MinScan - 2
         LastScan = tmp.UMCs(CurrClassInd).MaxScan + 2
         MinMW = tmp.UMCs(CurrClassInd).MinMW - 1
         MaxMW = tmp.UMCs(CurrClassInd).MaxMW + 1
         MyViewer.Zoom_UMC FirstScan, LastScan, MinMW, MaxMW
      End If
   End If
End If
CurrPeakInd = -1
Exit Sub
LstClassesErrorHandler:
Debug.Print "Error occurred in lstClasses_Click: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmVisUMC->lstClasses_Click"
CurrPeakInd = -1
End Sub

Private Sub lstGroups_Click()
Dim CurrLstInd As Long
CurrLstInd = lstGroups.ListIndex
If CurrLstInd >= 0 Then
   CurrGroupInd = CurrLstGroupInd * MAX_GROUPS_IN_LIST + CurrLstInd
   ListClassesForGroup CurrGroupInd
End If
CurrClassInd = -1
CurrPeakInd = -1
End Sub


Private Sub mnuF_Click()
Call PickParameters
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFReset_Click()
'---------------------------------------------------------------
'reset classes to what they were before we started this function
'allow user to change mind if by accident
'---------------------------------------------------------------
Dim Res As Long
If NeedToSave Then
   Res = MsgBox("Changes made to the unique mass classes will be lost. Continue?", vbOKCancel, glFGTU)
   If Res = vbOK Then
      UpdateStatus "Resetting..."
      Erase tmp.UMCs
      tmp = GelUMC(CallerID)
      PairsUMCInSync = GelP_D_L(CallerID).SyncWithUMC
      Call ClearGroupsAndLists
      Call ResetUMC
      NeedToSave = False                  'indicate that nothing changed
      UpdateStatus ""
      mChangeList = ""
   End If
End If
End Sub

Private Sub mnuT_Click()
Call PickParameters
End Sub

Private Sub mnuTGReport_Click()
'--------------------------------------------------------------
'this function does not change anything it just creates reports
'--------------------------------------------------------------
Dim FileName As String
Dim Grouped As Long
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim CurrMWDistance As Double
Dim i As Long
On Error Resume Next
UpdateStatus "Generating grouping reports..."
FileName = GetTempFolder() & RawDataTmpFile
Select Case f1Type
Case F_MW_SCAN_PROXIMITY
    MsgBox "This report is not implemented yet.", vbOKOnly, glFGTU
Case F_MW_SCAN_EXACT_DISTANCE
    If f1Multiplicity > 1 Then
       Me.MousePointer = vbHourglass
       Set ts = fs.OpenTextFile(FileName, ForWriting, True)
       ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
       ts.WriteLine "Gel File: " & GelBody(CallerID).Caption
       ts.WriteLine "Reporting distance analysis of Unique Mass Classes."
       ts.WriteLine "MW Distance: " & f1MWDist
       ts.WriteLine "MW Tolerance: " & f1MWTol
       ts.WriteLine "Scan Distance: " & f1ScanDist
       ts.WriteLine "MW distance range multiplicity: " & f1Multiplicity
       For i = 1 To f1Multiplicity
           CurrMWDistance = i * f1MWDist
           Grouped = GroupByMWDistanceScanCount(CurrMWDistance)
           ts.WriteLine "Groups on distance: " & CurrMWDistance & ": " & Grouped
       Next i
       ts.Close
       Set ts = Nothing
       Set fs = Nothing
       Me.MousePointer = vbDefault
       frmDataInfo.Tag = "Misc"
       frmDataInfo.show vbModal
    Else
       MsgBox "MW distance must be positive number with this option.", vbOKOnly, glFGTU
    End If
Case F_MW_EQUIVALENCY
    MsgBox "This report is not implemented yet.", vbOKOnly, glFGTU
Case F_MEMBERS_SHARING_EQUIVALENCY
    MsgBox "This report is not implemented yet.", vbOKOnly, glFGTU
End Select
UpdateStatus ""
End Sub

Private Sub mnuTGroup_Click(Index As Integer)
Dim Grouped As Boolean
UpdateStatus "Grouping classes..."
'destroy existing groups and lists
ClearGroupsAndLists
f1Type = Index
Select Case f1Type
Case F_MW_SCAN_PROXIMITY
    chkMultiMemberGroups.value = vbChecked
    Grouped = GroupByMWScanProximity()
Case F_MW_SCAN_EXACT_DISTANCE
    chkMultiMemberGroups.value = vbChecked
    If f1MWDist > 0 Then
       Grouped = GroupByMWDistanceScan()
    Else
       MsgBox "MW distance must be positive number with this option.", vbOKOnly, glFGTU
    End If
Case F_MW_EQUIVALENCY
    Grouped = GroupByMWEquivalency()
Case F_MEMBERS_SHARING_EQUIVALENCY
    MsgBox "This option is not implemented yet.", vbOKOnly, glFGTU
    Grouped = True
    'uncomment next statement when ready (and delete previous two)
    'Grouped = GroupByMemberSharingEquivalency()
End Select
If Grouped Then
   UpdateStatus ""
Else
   UpdateStatus "Error grouping classes."
End If
End Sub

Private Sub mnuTMAutoMerge_Click()
'---------------------------------------------------------------
'merge all groups satisfying conditions
'---------------------------------------------------------------
Dim i As Long, j  As Long
Dim BestClassInd As Long
Dim lngOriginalUMCCount As Long
On Error GoTo err_mnuTAutoMerge
UpdateStatus "Merging groups of classes..."
With GrRes
    lngOriginalUMCCount = .Count
    For i = 0 To .Count - 1
        With .Members(i)
            If .Count > 1 Then   'otherwise already merged
                If IsAcceptableForAutoMerge(i) Then
                   BestClassInd = .Members(0)
                   For j = 1 To .Count - 1
                       BestClassInd = MergeClassesAuto(BestClassInd, .Members(j))
                   Next j
                End If
            End If
        End With
    Next i
End With
UpdateStatus "Recalculating class structure..."
Call ClearGroupsAndLists
Call RemoveClasses
Call ResetUMC
If lngOriginalUMCCount > 0 Then AddToTentativeChangeList "Auto-merged UMC's: Original UMC Group Count = " & Trim(lngOriginalUMCCount) & "; Count after merge = " & Trim(GrRes.Count) & "; Auto-merge mode = " & cmbAutoMerge.List(cmbAutoMerge.ListIndex) & "; MinMW = " & txtAutoMergeMinMW & "; MaxMW = " & txtAutoMergeMaxMW
NeedToSave = True
UpdateStatus ""
err_mnuTAutoMerge:
UpdateStatus "Error merging multi member groups."
End Sub

Private Sub mnuTMMergeAll_Click()
'--------------------------------------------------------------
'merges classes from all multi groups to a single member groups
'--------------------------------------------------------------
Dim i As Long, j  As Long
Dim BestClassInd As Long
On Error GoTo err_cmdMergeMergeAll
UpdateStatus "Merging groups of classes..."
With GrRes
    If .Count > 0 Then
       For i = 0 To .Count - 1
           With .Members(i)
               If .Count > 1 Then   'otherwise already merged
                  BestClassInd = .Members(0)
                  For j = 1 To .Count - 1
                      BestClassInd = MergeClasses(BestClassInd, .Members(j))
                  Next j
               End If
           End With
       Next i
    End If
End With
UpdateStatus "Recalculating class structure..."
Call ClearGroupsAndLists
Call RemoveClasses
Call ResetUMC
NeedToSave = True
UpdateStatus ""
Exit Sub
err_cmdMergeMergeAll:
UpdateStatus "Error merging multi member groups."
End Sub

Private Sub mnuTMMergeGroup_Click()
'---------------------------------------------------------------
'merges classes from multi groups to a single member groups from
'currently listed groups
'---------------------------------------------------------------
Dim i As Long, j  As Long
Dim BestClassInd As Long
On Error GoTo err_cmdMergeMergeGrp
UpdateStatus "Merging groups of classes..."
With GrRes
    For i = LstGroupsInd1(CurrLstGroupInd) To LstGroupsInd2(CurrLstGroupInd)
        With .Members(i)
            If .Count > 1 Then   'otherwise already merged
                BestClassInd = .Members(0)
                For j = 1 To .Count - 1
                    BestClassInd = MergeClasses(BestClassInd, .Members(j))
                Next j
            End If
        End With
    Next i
End With
UpdateStatus "Recalculating class structure..."
Call ClearGroupsAndLists
Call RemoveClasses
Call ResetUMC
NeedToSave = True
UpdateStatus ""
Exit Sub
err_cmdMergeMergeGrp:
UpdateStatus "Error merging multi member groups."
End Sub

Private Sub mnuTMMergeSelClasses_Click()
'------------------------------------------------------------------
'merges selected classes; recalculates classes; clears groups and
'list boxes; sets flag to indicate that something was changed from
'where we actually started on this form
'------------------------------------------------------------------
Dim BestClassInd As Long        'index of class that will remain
Dim i As Long
On Error GoTo exit_MergeSelectedClasses
With lstSelectedClasses
    If .ListCount > 1 Then              'nothing to merge otherwise
       BestClassInd = .ItemData(0)
       For i = 1 To .ListCount - 1
           BestClassInd = MergeClasses(BestClassInd, .ItemData(i))
       Next i
    End If
End With
UpdateStatus "Recalculating class structure..."
Call ClearGroupsAndLists
Call RemoveClasses
Call ResetUMC
lstSelectedClasses.Clear
NeedToSave = True       'mark that something changed
exit_MergeSelectedClasses:
End Sub

Private Sub mnuTRAutoRemove_Click()
txtLoAbuPct_LostFocus
txtHiAbuPct_LostFocus
txtLoCnt_LostFocus
txtHiCnt_LostFocus
txtPercentMaxAbuToUseToGaugeLength_LostFocus
txtAutoRefineMinimumMemberCount_LostFocus
AutoRemoveUMCsWork
End Sub

Private Sub mnuTRRemoveClass_Click()
'-------------------------------------------------------
'removes selected unique mass class from structure
'-------------------------------------------------------

End Sub

Private Sub mnuTUMCAvgAbu_Click()
'-------------------------------------------------------
'calculates average abundance of unique mass class
'-------------------------------------------------------
Dim i As Long
Dim DataInd As Long
Dim Res As String
Dim AbuSum As Double
On Error Resume Next
Res = "Error"
If CurrClassInd >= 0 Then
  With GelData(CallerID)
    For i = 0 To tmp.UMCs(CurrClassInd).ClassCount - 1
      DataInd = tmp.UMCs(CurrClassInd).ClassMInd(i)
      Select Case tmp.UMCs(CurrClassInd).ClassMType(i)
      Case glCSType
        AbuSum = AbuSum + .CSData(DataInd).Abundance
      Case glIsoType
        AbuSum = AbuSum + .IsoData(DataInd).Abundance
      End Select
    Next i
    Res = Format$(AbuSum / tmp.UMCs(CurrClassInd).ClassCount, "Scientific")
  End With
End If
lblUMCCalculator.Caption = "UMC: " & CurrClassInd & vbCrLf & "Avg.Abu.= " & Res
End Sub

Private Sub mnuTUMCAvgMW_Click()
'-------------------------------------------------------
'calculates average molecular mass of the class
'-------------------------------------------------------
Dim i As Long
Dim DataInd As Long
Dim Res As String
Dim SumMW As Double
On Error Resume Next
Res = "Error"
If CurrClassInd >= 0 Then
  With GelData(CallerID)
    For i = 0 To tmp.UMCs(CurrClassInd).ClassCount - 1
      DataInd = tmp.UMCs(CurrClassInd).ClassMInd(i)
      Select Case tmp.UMCs(CurrClassInd).ClassMType(i)
      Case glCSType
        SumMW = SumMW + .CSData(DataInd).AverageMW
      Case glIsoType
        SumMW = SumMW + GetIsoMass(.IsoData(DataInd), tmp.def.MWField)
      End Select
      Res = Format$(SumMW / tmp.UMCs(CurrClassInd).ClassCount, "0.0000")
    Next i
  End With
End If
lblUMCCalculator.Caption = "UMC: " & CurrClassInd & vbCrLf & "Avg.MW= " & Res
End Sub

Private Sub mnuTUMCSumAbu_Click()
'-------------------------------------------------------
'calculates sum of abundance of unique mass class
'-------------------------------------------------------
Dim i As Long
Dim DataInd As Long
Dim Res As String
Dim AbuSum As Double
On Error Resume Next
Res = "Error"
If CurrClassInd >= 0 Then
  With GelData(CallerID)
    For i = 0 To tmp.UMCs(CurrClassInd).ClassCount - 1
      DataInd = tmp.UMCs(CurrClassInd).ClassMInd(i)
      Select Case tmp.UMCs(CurrClassInd).ClassMType(i)
      Case glCSType
        AbuSum = AbuSum + .CSData(DataInd).Abundance
      Case glIsoType
        AbuSum = AbuSum + .IsoData(DataInd).Abundance
      End Select
    Next i
    Res = Format$(AbuSum, "Scientific")
  End With
End If
lblUMCCalculator.Caption = "UMC: " & CurrClassInd & vbCrLf & "Sum.Abu.= " & Res
End Sub


Private Sub MyViewer_pvControlDone()
Me.SetFocus
End Sub

Private Sub MyViewer_pvUnload()
chkShowUMC.value = vbUnchecked
End Sub

Private Sub txtAutoMergemaxmw_LostFocus()
Dim tmp As String
tmp = Trim$(txtAutoMergeMaxMW.Text)
If IsNumeric(tmp) Then
   If tmp > 0 Then
      amMWMax = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be positive number.", vbOKOnly, glFGTU
txtAutoMergeMaxMW.SetFocus
End Sub

Private Sub txtAutoMergeMinMW_LostFocus()
Dim tmp As String
tmp = Trim$(txtAutoMergeMinMW.Text)
If IsNumeric(tmp) Then
   If tmp > 0 Then
      amMWMin = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be positive number.", vbOKOnly, glFGTU
txtAutoMergeMinMW.SetFocus
End Sub

Private Sub txtAutoRefineMinimumMemberCount_LostFocus()
If IsNumeric(txtAutoRefineMinimumMemberCount.Text) Then
    glbPreferencesExpanded.UMCAutoRefineOptions.MinMemberCountWhenUsingScanRange = Abs(CLng(txtAutoRefineMinimumMemberCount.Text))
Else
   MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
   txtAutoRefineMinimumMemberCount.SetFocus
End If
End Sub

Private Sub txtf1MWDist_LostFocus()
Dim tmp As String
tmp = Trim$(txtf1MWDist.Text)
If IsNumeric(tmp) Then
   If tmp > 0 Then
      f1MWDist = CDbl(tmp)
      Exit Sub
   End If
End If
'at this point something is wrong
MsgBox "This argument should be positive number.", vbOKOnly, glFGTU
txtf1MWDist.SetFocus
End Sub

Private Sub txtf1MWTol_LostFocus()
Dim tmp As String
tmp = Trim$(txtf1MWTol.Text)
If IsNumeric(tmp) Then
   If tmp > 0 Then
      f1MWTol = CDbl(tmp)
      Exit Sub
   End If
End If
'at this point something is wrong
MsgBox "This argument should be positive number.", vbOKOnly, glFGTU
txtf1MWTol.SetFocus
End Sub

Private Sub txtf1ScanDist_LostFocus()
Dim tmp As String
tmp = Trim$(txtf1ScanDist.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 Then
      f1ScanDist = CLng(tmp)
      Exit Sub
   End If
End If
'at this point something is wrong
MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
txtf1ScanDist.SetFocus
End Sub

Private Function GroupByMWScanProximity() As Boolean
'------------------------------------------------------------
'this option groups close classes based on their proximity
'it uses parameters f1MWTol (in Daltons) and f1ScanDist
'returns True if successful(even if no groups),False on error
'This function also has only sense for multi-member groups!!!
'------------------------------------------------------------
Dim i As Long, j As Long
Dim MyGroupCnt As Long
Dim MyGroup() As Long       'indexes of classes belonging to current class
Dim MinInd As Long
Dim MaxInd As Long
Dim CurrMW As Double
On Error GoTo err_GroupByMWScanProximity

'fill search object with sorted class molecular masses
If Not MWSearch.Fill(UMCMW) Then GoTo err_GroupByMWScanProximity
With tmp
    For i = 0 To .UMCCnt - 1
        MyGroupCnt = 0
        Erase MyGroup
        CurrMW = UMCMW(i)
        MinInd = 0
        MaxInd = -1
        Call MWSearch.FindIndexRange(CurrMW, f1MWTol, MinInd, MaxInd)
        If MinInd < MaxInd Then
           For j = MinInd To MaxInd
               If UMCMW(j) > UMCMW(i) Then        'dont want to make same grouping twice
                  If ScanCloseClasses(UMCInd(i), UMCInd(j)) Then
                     MyGroupCnt = MyGroupCnt + 1
                     ReDim Preserve MyGroup(MyGroupCnt)
                     MyGroup(MyGroupCnt) = UMCInd(j)
                  End If
               End If
           Next j
        End If
        If MyGroupCnt > 0 Then
           MyGroup(0) = UMCInd(i)   'set first element group generator
           With GrRes
                .Count = .Count + 1
                ReDim Preserve .Members(.Count - 1)
                .Members(.Count - 1).Members = MyGroup          'this is array assignment
                .Members(.Count - 1).Count = MyGroupCnt + 1
                .Members(.Count - 1).Description = "Prox.MW: " & Format(CurrMW, "0.0000")
           End With
        End If
    Next i
End With

Call ResolveResults
GroupByMWScanProximity = True
Exit Function

err_GroupByMWScanProximity:
End Function


Private Function GroupByMWDistanceScan() As Boolean
'------------------------------------------------------
'groups classes if they are on certain mass distance
'(within tolerance) and scan separation from each other
'returns True if successful, False on error
'NOTE: this option uses all three parameters; both MW
'      distance and tolerance are in Daltons(absolute)
'NOTE: group members are ordered in ascending order
'------------------------------------------------------
Dim i As Long, j As Long
Dim MyGroupCnt As Long
Dim MyGroup() As Long       'indexes of classes belonging to current class
Dim MinInd As Long
Dim MaxInd As Long
Dim CurrMW As Double
On Error GoTo err_GroupByMWDistanceScan

'fill search object with sorted class molecular masses
If Not MWSearch.Fill(UMCMW) Then GoTo err_GroupByMWDistanceScan
With tmp
    For i = 0 To .UMCCnt - 1
        MyGroupCnt = 0
        Erase MyGroup
        CurrMW = UMCMW(i)
        MinInd = 0
        MaxInd = -1
        Call MWSearch.FindIndexRange(CurrMW + f1MWDist, f1MWTol, MinInd, MaxInd)
        If MinInd < MaxInd Then
           For j = MinInd To MaxInd
               If UMCMW(j) > UMCMW(i) Then        'dont want to make same grouping twice
                  If ScanCloseClasses(UMCInd(i), UMCInd(j)) Then
                     MyGroupCnt = MyGroupCnt + 1
                     ReDim Preserve MyGroup(MyGroupCnt)
                     MyGroup(MyGroupCnt) = UMCInd(j)
                  End If
               End If
           Next j
        End If
        If MyGroupCnt > 0 Then
           MyGroup(0) = UMCInd(i)   'set first element group generator
           With GrRes
                .Count = .Count + 1
                ReDim Preserve .Members(.Count - 1)
                .Members(.Count - 1).Members = MyGroup          'this is array assignment
                .Members(.Count - 1).Count = MyGroupCnt + 1
                .Members(.Count - 1).Description = "Dist. MW: " & Format(CurrMW, "0.0000")
           End With
        End If
    Next i
End With

Call ResolveResults
GroupByMWDistanceScan = True
Exit Function

err_GroupByMWDistanceScan:
End Function


Private Function GroupByMWEquivalency() As Boolean
'----------------------------------------------------------------------
'two classes are equivalent if their mass distance is less than f1MWTol
'returns True if successful, False on error
'----------------------------------------------------------------------
Dim PrevMW As Double
On Error GoTo err_GroupByMWEquivalency
Dim i As Long
PrevMW = -1
With GrRes
    For i = 0 To tmp.UMCCnt - 1
        If UMCMW(i) - PrevMW >= f1MWTol Then    'we have new group
           .Count = .Count + 1
           ReDim Preserve .Members(.Count - 1)
           .Members(.Count - 1).Count = 1
           .Members(.Count - 1).Description = "Equ. MW: " & Format(UMCMW(i), "0.0000")
           ReDim .Members(.Count - 1).Members(0)
           .Members(.Count - 1).Members(0) = UMCInd(i)      'index in GelUMC
        Else                                    'add class to group
           With .Members(.Count - 1)
                .Count = .Count + 1
                ReDim Preserve .Members(.Count - 1)
                .Members(.Count - 1) = UMCInd(i)      'index in GelUMC
           End With
        End If
        PrevMW = UMCMW(i)
    Next i
End With
Call ResolveResults
GroupByMWEquivalency = True
Exit Function

err_GroupByMWEquivalency:
End Function


''Private Function GroupByMemberSharingEquivalency() As Boolean
'''------------------------------------------------------------
'''two classes are equivalent if they share at least one member
'''returns True if successful, False on error
'''------------------------------------------------------------
''On Error GoTo err_GroupByMemberSharingEquivalency
''Dim i As Long
''
''
''
''
''GroupByMemberSharingEquivalency = True
''Exit Function
''
''err_GroupByMemberSharingEquivalency:
''End Function
''


Private Sub ResolveResults()
'------------------------------------------------------------------------
'resolve results obtained from grouping functions in a user friendly form
'eliminates groups with less than 2 class-members if required
'------------------------------------------------------------------------
Dim i As Long
Dim NewResCnt As Long
cmbLstGroups.Clear
cmbLstGroups.Enabled = False
With GrRes
    If MultiGroupsOnly Then
       If GetOneClassGroupsCount() > 0 Then
          NewResCnt = 0
          For i = 0 To .Count - 1
              If .Members(i).Count > 1 Then
                 NewResCnt = NewResCnt + 1
                 .Members(NewResCnt - 1) = .Members(i)
              End If
          Next i
          .Count = NewResCnt
          If NewResCnt > 0 Then
             ReDim Preserve .Members(NewResCnt - 1)
          Else
             Erase .Members
          End If
       End If
    End If
    If .Count > 0 Then
       If .Count > MAX_GROUPS_IN_LIST Then
          If .Count Mod MAX_GROUPS_IN_LIST < 10 Then            'don't allow last list
            LstGroupsCnt = Int(.Count / MAX_GROUPS_IN_LIST)     'to be too small
          Else
            LstGroupsCnt = Int(.Count / MAX_GROUPS_IN_LIST) + 1
          End If
          ReDim LstGroupsInd1(LstGroupsCnt - 1)
          ReDim LstGroupsInd2(LstGroupsCnt - 1)
          ReDim LstGroupsDisplay(LstGroupsCnt - 1)
          For i = 0 To LstGroupsCnt - 1
              LstGroupsInd1(i) = i * MAX_GROUPS_IN_LIST
              LstGroupsInd2(i) = (i + 1) * MAX_GROUPS_IN_LIST - 1
          Next i
          LstGroupsInd2(LstGroupsCnt - 1) = .Count - 1          'make sure last index is really last
          'create names for each part of the list and fill the combo box
          For i = 0 To LstGroupsCnt - 1
              LstGroupsDisplay(i) = "Groups " & LstGroupsInd1(i) & " - " & LstGroupsInd2(i)
              cmbLstGroups.AddItem LstGroupsDisplay(i)
          Next i
          cmbLstGroups.Enabled = True
          'present first group of groups
          cmbLstGroups.ListIndex = 0
       Else
          For i = 0 To .Count - 1
              lstGroups.AddItem .Members(i).Description
          Next i
          CurrLstGroupInd = 0           'important for calculation of absolute index
       End If
       lblGroupsCount.Caption = .Count
    Else
       lblGroupsCount.Caption = "0"
    End If
End With
End Sub



Public Sub ListClassesForGroup(ByVal GroupInd As Long)
'----------------------------------------------------------
'fills list box with class descriptions for specified group
'----------------------------------------------------------
Dim i As Long
On Error Resume Next
lstClasses.Clear
lstPeaks.Clear
If GroupInd >= 0 Then
   With GrRes.Members(GroupInd)
       For i = 0 To .Count
           lstClasses.AddItem UMCDisplay(.Members(i))
       Next i
   End With
End If
End Sub


Private Function PrepareClasses() As Boolean
'------------------------------------------------------------
'calculates UMC statistics and prepares classes display names
'NOTE: this is simplified UMCStatistics2 function
'------------------------------------------------------------
'column 0 - class index in .UMCs
'column 1 - class first scan number
'column 2 - class last scan number
Dim i As Long, j As Long
On Error GoTo err_PrepareClasses

With tmp
   If .UMCCnt > 0 Then
      ReDim UMCStat(.UMCCnt - 1, 2)
      ReDim UMCDisplay(.UMCCnt - 1)
   Else
      ReDim UMCStat(0, 2)
      ReDim UMCDisplay(0)
   End If
   
   For i = 0 To .UMCCnt - 1
      With .UMCs(i)
          UMCStat(i, 0) = i
          If .ClassCount > 0 Then
             'class members are ordered on scan numbers
             Select Case .ClassMType(0)                 'first scan number
             Case gldtCS
               UMCStat(i, 1) = GelData(CallerID).CSData(.ClassMInd(0)).ScanNumber
             Case gldtIS
               UMCStat(i, 1) = GelData(CallerID).IsoData(.ClassMInd(0)).ScanNumber
             End Select
             Select Case .ClassMType(.ClassCount - 1)   'last scan number
             Case gldtCS
               UMCStat(i, 2) = GelData(CallerID).CSData(.ClassMInd(.ClassCount - 1)).ScanNumber
             Case gldtIS
               UMCStat(i, 2) = GelData(CallerID).IsoData(.ClassMInd(.ClassCount - 1)).ScanNumber
             End Select
          Else     'this should not happen
             For j = 0 To 2
                 UMCStat(i, j) = -1
             Next j
          End If
          UMCDisplay(i) = "UMC " & i & "; Scans [" & UMCStat(i, 1) & "," & UMCStat(i, 2) _
                        & "]; MW~" & Format$(.ClassMW, "0.00") & "Da; Count " & .ClassCount
      End With
   Next i
End With
PrepareClasses = True
Exit Function

err_PrepareClasses:
End Function



Private Sub ListPeaksForClass(ByVal ClassInd As Long)
'-------------------------------------------------------
'fills list with description of class members
'description contains type CS or IS; index in Num arrays
'scan number, charge state, fit, abundance
'-------------------------------------------------------
Dim i As Long
Dim DataInd As Long
Dim Desc As String
lstPeaks.Clear
If ClassInd >= 0 Then
  With GelData(CallerID)
    For i = 0 To tmp.UMCs(ClassInd).ClassCount - 1
      DataInd = tmp.UMCs(ClassInd).ClassMInd(i)
      Select Case tmp.UMCs(ClassInd).ClassMType(i)
      Case glCSType
        Desc = "CS " & DataInd & "; Scan " & .CSData(DataInd).ScanNumber & "; " _
               & Format$(.CSData(DataInd).AverageMW, "0.00") & "; " _
               & Format$(.CSData(DataInd).Abundance, "Scientific") _
               & "; CS" & .CSData(DataInd).Charge & "; " _
               & "; Fit NA"

      Case glIsoType
        Desc = "IS " & DataInd & "; Scan " & .IsoData(DataInd).ScanNumber & "; " _
               & Format$(GetIsoMass(.IsoData(DataInd), tmp.def.MWField), "0.00") & "; " _
               & Format$(.IsoData(DataInd).Abundance, "Scientific") _
               & "; CS " & .IsoData(DataInd).Charge & "; Fit " _
               & Format$(.IsoData(DataInd).Fit, "0.00")
      End Select
      lstPeaks.AddItem Desc
    Next i
  End With
End If
End Sub

Private Function MergeClasses(ByVal UMCInd1 As Long, _
                              ByVal UMCInd2 As Long) As Long
'-----------------------------------------------------------------
'merge elements of classes UMCInd1 and UMCInd2 in one class
'resulting class has lower index of UMCInd1 and UMCInd2 and class
'representative from that class - that way preferences in class
'build is preserved; class with higher index is not removed - its
'index in TmpInc is marked with -1; removal of all marked classes
'is done with function RemoveClasses
'Function returns index of MIP class; -1 on any error
'NOTE: If called with UMCInd1=UMCInd2 function will mark class for
'removal without doing anything else
'-----------------------------------------------------------------
Dim MIPClassInd As Long
Dim LIPClassInd As Long
'peaks that has to be added to More Important Class
Dim PeakCnt As Long
Dim PeakType() As Long
Dim PeakInd() As Long
Dim i As Long
On Error GoTo err_MergeClasses
If UMCInd1 = UMCInd2 Then
   TmpInc(UMCInd1) = REMOVE_UMC_MARK
Else
   If UMCInd1 < UMCInd2 Then
      MIPClassInd = UMCInd1
      LIPClassInd = UMCInd2
   Else
      MIPClassInd = UMCInd2
      LIPClassInd = UMCInd1
   End If
   With GelUMC(CallerID).UMCs(LIPClassInd)
       'redimension to highest number of peaks that could be added
       ReDim PeakType(.ClassCount - 1)
       ReDim PeakInd(.ClassCount - 1)
       PeakCnt = 0
       For i = 0 To .ClassCount - 1
           If Not IsClassMember(.ClassMType(i), .ClassMInd(i), MIPClassInd) Then
              PeakCnt = PeakCnt + 1
              PeakType(PeakCnt - 1) = .ClassMType(i)
              PeakInd(PeakCnt - 1) = .ClassMInd(i)
           End If
       Next i
       If PeakCnt > 0 Then
          If PeakCnt < .ClassCount Then 'redimension if necessary
             ReDim Preserve PeakType(PeakCnt - 1)
             ReDim Preserve PeakInd(PeakCnt - 1)
          End If
          If Not AddPeakArrayToTheClass(PeakType(), PeakInd(), MIPClassInd) Then GoTo err_MergeClasses
       End If
   End With
   If Not RecalculateClass(MIPClassInd) Then GoTo err_MergeClasses
   TmpInc(LIPClassInd) = REMOVE_UMC_MARK
End If
MergeClasses = MIPClassInd
Exit Function

err_MergeClasses:
MergeClasses = -1
End Function



Private Function RemoveClasses() As Long
'---------------------------------------------------------------
'update unique mass classes structure by removing classes marked
'as REMOVE_UMC_MARK classes; returns number of removed classes
'or -1 on any error
'---------------------------------------------------------------
Dim i As Long
Dim Cnt As Long, RemoveCnt As Long
On Error GoTo err_RemoveClasses
With tmp
  If GetIncCount() <> .UMCCnt Then          'don't go in this process
    For i = 0 To .UMCCnt - 1                'before first verifying there
        If TmpInc(i) < 0 Then               'is something to remove
           Erase .UMCs(i).ClassMInd
           Erase .UMCs(i).ClassMType
           .UMCs(i).ClassCount = 0
           RemoveCnt = RemoveCnt + 1
        Else
           Cnt = Cnt + 1
           .UMCs(Cnt - 1) = .UMCs(i)
        End If
    Next i
    If Cnt > 0 Then
        ReDim Preserve .UMCs(Cnt - 1)
    Else
        ReDim Preserve .UMCs(0)
    End If
    .UMCCnt = Cnt
    If RemoveCnt > 0 Then PairsUMCInSync = False
  End If
End With
RemoveClasses = RemoveCnt
Exit Function

err_RemoveClasses:
Debug.Assert False
RemoveClasses = -1
End Function


Private Function ResetUMC() As Boolean
'-------------------------------------------------------------
'reset all arrays related with unique mass classes and returns
'True if succesful; this function should be called when form
'loads and after each change in unique mass classes structure
'-------------------------------------------------------------
On Error GoTo err_ResetUMC
Dim dQS As New QSDouble
Dim i As Long

UpdateLabels LBL_ZLS
If Not PrepareClasses() Then GoTo err_ResetUMC
If tmp.UMCCnt > 0 Then
    With tmp
       ReDim TmpInc(.UMCCnt - 1)
       ReDim UMCInd(.UMCCnt - 1)
       ReDim UMCMW(.UMCCnt - 1)
       For i = 0 To .UMCCnt - 1
           UMCMW(i) = .UMCs(i).ClassMW
           UMCInd(i) = i
       Next i
       If Not dQS.QSAsc(UMCMW, UMCInd) Then GoTo err_ResetUMC
    End With
End If

If Not PairsUMCInSync Then                  'disable using pairs if something changed
   chkRemovePairedLUMC.value = vbUnchecked
   chkRemovePairedHUMC.value = vbUnchecked
   chkRemovePairedLUMC.Enabled = False
   chkRemovePairedHUMC.Enabled = False
End If
UpdateLabels 1
ResetUMC = True
Exit Function

err_ResetUMC:
UpdateLabels LBL_ERR
End Function

''Private Function AddPeakToTheClass(ByVal PeakType As Long, _
''                                   ByVal PeakInd As Long, _
''                                   ByVal ClassInd As Long) As Long
'''-------------------------------------------------------------------
'''adds specified peak to the class and returns its index in the class
'''or -1 in case of any error
'''-------------------------------------------------------------------
''On Error GoTo err_AddPeakToTheClass
''With GelUMC(CallerID).UMCs(ClassInd)
''     .ClassCount = .ClassCount + 1
''     ReDim Preserve .ClassMInd(.ClassCount - 1)
''     ReDim Preserve .ClassMType(.ClassCount - 1)
''     .ClassMInd(.ClassCount - 1) = PeakInd
''     .ClassMType(.ClassCount - 1) = PeakType
''     AddPeakToTheClass = .ClassCount - 1
''End With
''Exit Function
''
''err_AddPeakToTheClass:
''AddPeakToTheClass = -1
''End Function


Private Function AddPeakArrayToTheClass(NewPeakType() As Long, _
                                        NewPeakInd() As Long, _
                                        ByVal ClassInd As Long) As Boolean
'-------------------------------------------------------------------------
'adds specified peak array to the class and returns True if successful
'NOTE: has to be careful here since class members have to be ordered on
'      scan numbers
'-------------------------------------------------------------------------
Dim NewPeaksCnt As Long
Dim TTlPeaksCnt As Long
Dim TTlPeakInd() As Long
Dim TTLPeakType() As Long
Dim TTlFN() As Long             'scan number
Dim TTlOrd() As Long            'ordering array
Dim QSL As New QSLong           'ordering object
Dim i As Long, j As Long
On Error GoTo err_AddPeakArrayToTheClass
NewPeaksCnt = UBound(NewPeakInd) + 1
If NewPeaksCnt > 0 Then         'not error if nothing to add
   With GelUMC(CallerID).UMCs(ClassInd)
       'we have to put new and old together and order based on scan numbers
       TTlPeaksCnt = .ClassCount + NewPeaksCnt
       ReDim TTlPeakInd(TTlPeaksCnt - 1)
       ReDim TTLPeakType(TTlPeaksCnt - 1)
       ReDim TTlOrd(TTlPeaksCnt - 1)
       ReDim TTlFN(TTlPeaksCnt - 1)
       For i = 0 To .ClassCount - 1             'first old peaks
           TTlPeakInd(i) = .ClassMInd(i)
           TTLPeakType(i) = .ClassMType(i)
           TTlOrd(i) = i
           Select Case .ClassMType(i)
           Case glCSType
                TTlFN(i) = GelData(CallerID).CSData(.ClassMInd(i)).ScanNumber
           Case glIsoType
                TTlFN(i) = GelData(CallerID).IsoData(.ClassMInd(i)).ScanNumber
           End Select
       Next i
       For i = 0 To NewPeaksCnt - 1             'then new peaks
           j = .ClassCount + i
           TTlOrd(j) = j
           TTlPeakInd(j) = NewPeakInd(i)
           TTLPeakType(j) = NewPeakType(i)
           Select Case TTLPeakType(j)
           Case glCSType
                TTlFN(j) = GelData(CallerID).CSData(TTlPeakInd(j)).ScanNumber
           Case glIsoType
                TTlFN(j) = GelData(CallerID).IsoData(TTlPeakInd(j)).ScanNumber
           End Select
       Next i
       'now order them on scan numbers ascending
       If Not QSL.QSAsc(TTlFN(), TTlOrd()) Then GoTo err_AddPeakArrayToTheClass
       Set QSL = Nothing
       ReDim .ClassMInd(TTlPeaksCnt - 1)
       ReDim .ClassMType(TTlPeaksCnt - 1)
       For i = 0 To TTlPeaksCnt - 1
           .ClassMInd(i) = TTlPeakInd(TTlOrd(i))
           .ClassMType(i) = TTLPeakType(TTlOrd(i))
       Next i
       .ClassCount = TTlPeaksCnt
   End With
End If
AddPeakArrayToTheClass = True
Exit Function

err_AddPeakArrayToTheClass:
End Function



''Private Function RemovePeakFromTheClass(ByVal PeakClassInd As Long, _
''                                        ByVal ClassInd As Long) As Boolean
'''----------------------------------------------------------------------------
'''removes peak with class index from the class and return True if OK.
'''If class becomes empty by removing this peak or peak is class representative
'''class is marked for removal
'''NOTE: Index here is peak index in the class
'''----------------------------------------------------------------------------
''Dim i As Long
''Dim PeakInd As Long
''Dim PeakType As Long
''On Error GoTo err_RemovePeakFromTheClass
''
''With tmp.UMCs(ClassInd)
''    If .ClassCount > 1 Then
''       For i = PeakClassInd To .ClassCount - 2
''           .ClassMType(i) = .ClassMType(i + 1)
''           .ClassMInd(i) = .ClassMInd(i + 1)
''       Next i
''       .ClassCount = .ClassCount - 1
''       ReDim Preserve .ClassMInd(.ClassCount - 1)
''       ReDim Preserve .ClassMType(.ClassCount - 1)
''       'call recalculate class MW/Abu
''    Else
''       .ClassCount = 0
''       Erase .ClassMType
''       Erase .ClassMInd
''       TmpInc(ClassInd) = REMOVE_UMC_MARK
''    End If
''End With
''
''RemovePeakFromTheClass = True
''Exit Function
''
''err_RemovePeakFromTheClass:
''End Function

Private Function RecalculateClass(ByVal ClassInd) As Boolean
'-----------------------------------------------------------
'recalculates class MW and abundance based on current class
'membership and definition
'-----------------------------------------------------------
Dim MWSum As Double
Dim AbuSum As Double
Dim i As Long
On Error GoTo err_RecalculateClass

'no need to waste time on excluded classes
If TmpInc(ClassInd) < 0 Then GoTo err_RecalculateClass
With tmp.UMCs(ClassInd)
    'no need to recalculate for class representative since it can not change
    If (tmp.def.ClassMW <> UMCClassMassConstants.UMCMassRep) Or (tmp.def.ClassAbu <> UMCClassAbundanceConstants.UMCAbuRep) Then
        For i = 0 To .ClassCount - 1
            Select Case .ClassMType(i)
            Case glCSType
                 MWSum = MWSum + GelData(CallerID).CSData(.ClassMInd(i)).AverageMW
                 AbuSum = AbuSum + GelData(CallerID).CSData(.ClassMInd(i)).Abundance
            Case glIsoType
                 MWSum = MWSum + GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(i)), tmp.def.MWField)
                 AbuSum = AbuSum + GelData(CallerID).IsoData(.ClassMInd(i)).Abundance
            End Select
        Next i
        If tmp.def.ClassMW = UMCClassMassConstants.UMCMassAvg Then .ClassMW = MWSum / .ClassCount
        Select Case tmp.def.ClassAbu
        Case UMCClassAbundanceConstants.UMCAbuRep      'do nothing
        Case UMCClassAbundanceConstants.UMCAbuSum
             .ClassAbundance = AbuSum
        Case UMCClassAbundanceConstants.UMCAbuAvg
             .ClassAbundance = AbuSum / .ClassCount
        End Select
    End If
End With
RecalculateClass = True
Exit Function

err_RecalculateClass:
End Function

Private Function GetOneClassGroupsCount() As Long
'------------------------------------------------
'returns number of groups with only one class
'------------------------------------------------
Dim i As Long
Dim Cnt As Long
On Error Resume Next
With GrRes
     For i = 0 To .Count - 1
         If .Members(i).Count <= 1 Then Cnt = Cnt + 1
     Next i
End With
GetOneClassGroupsCount = Cnt
End Function

Private Function IsClassMember(ByVal PeakType As Long, _
                               ByVal PeakInd As Long, _
                               ByVal ClassInd As Long) As Boolean
'---------------------------------------------------------------
'returns True if data point with index PeakInd of PeakType is
'member of class ClassInd
'---------------------------------------------------------------
Dim i As Long
On Error GoTo exit_IsClassMember
With GelUMC(CallerID).UMCs(ClassInd)
    For i = 0 To .ClassCount - 1
        If .ClassMInd(i) = PeakInd Then
           If .ClassMType(i) = PeakType Then
              IsClassMember = True
              Exit For
           End If
        End If
    Next i
End With

exit_IsClassMember:
End Function

Private Sub AddToTentativeChangeList(strProcessDescription As String)
    If Len(mChangeList) > 0 Then mChangeList = mChangeList & vbCrLf
    mChangeList = mChangeList & strProcessDescription
End Sub

Private Sub ClearGroupsAndLists()
'--------------------------------------------
'clears lists with groups, classes and peaks
'--------------------------------------------
Call DestroyGroups
lstGroups.Clear
lstClasses.Clear
lstPeaks.Clear
CurrGroupInd = -1
CurrClassInd = -1
CurrPeakInd = -1
cmbLstGroups.Clear
End Sub


Private Sub UpdateLabels(ByVal lblType As Long)
'----------------------------------------------
'updates labels displaying current counts
'----------------------------------------------
On Error Resume Next
Select Case lblType
Case LBL_ERR
    lblGroupsCount.Caption = "Error"
    lblCurrUMCCnt.Caption = "Error"
    lblCurrUMCRatio.Caption = "Error"
Case LBL_ZLS
    lblGroupsCount.Caption = ""
    lblCurrUMCCnt.Caption = ""
    lblCurrUMCRatio.Caption = ""
Case Else
    lblGroupsCount.Caption = GrRes.Count
    lblCurrUMCCnt.Caption = tmp.UMCCnt
    lblCurrUMCRatio.Caption = Format$(lblCurrUMCCnt.Caption / lblPeaksCnt.Caption, "0.00")
End Select
End Sub

Private Sub DestroyGroups()
'------------------------------------------------
'destroys groups structure
'------------------------------------------------
Dim i As Long
With GrRes
     For i = 0 To .Count - 1
        Erase .Members(i).Members
        .Members(i).Count = 0
        .Members(i).Description = ""
     Next i
     Erase .Members
     .Count = 0
     .Description = ""
End With
End Sub


Private Function ScanCloseClasses(ByVal ClassInd1 As Long, _
                                  ByVal ClassInd2 As Long) As Boolean
'--------------------------------------------------------------------
'returns True if classes ClassInd1 and ClassInd2 are close in regard
'of scan numbers
'NOTE: close here means that distance between closest(in scan regard)
'points of two classes is not more than f1ScanDist
'NOTE: to understand logic of this function draw all possible cases
'      of arangement of two segments in a 1D space
'--------------------------------------------------------------------
Dim ClosestScanDist As Long
If (UMCStat(ClassInd1, 1) < UMCStat(ClassInd2, 1)) And (UMCStat(ClassInd1, 2) < UMCStat(ClassInd2, 1)) Then
    ClosestScanDist = UMCStat(ClassInd2, 1) - UMCStat(ClassInd1, 2)
ElseIf (UMCStat(ClassInd2, 1) < UMCStat(ClassInd1, 1)) And (UMCStat(ClassInd2, 2) < UMCStat(ClassInd1, 1)) Then
    ClosestScanDist = UMCStat(ClassInd1, 1) - UMCStat(ClassInd2, 2)
Else
    ClosestScanDist = 0
End If
ScanCloseClasses = (ClosestScanDist <= f1ScanDist)
End Function


Private Sub txtHiAbuPct_LostFocus()
If IsNumeric(txtHiAbuPct.Text) Then
   glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefinePctHighAbundance = Abs(CDbl(txtHiAbuPct.Text))
Else
   MsgBox "This argument should be non-negative number.", vbOKOnly, glFGTU
   txtHiAbuPct.SetFocus
End If
End Sub

Private Sub txtHiCnt_LostFocus()
If IsNumeric(txtHiCnt.Text) Then
   glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineMaxLength = Abs(CLng(txtHiCnt.Text))
Else
   MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
   txtHiCnt.SetFocus
End If
End Sub

Private Sub txtLoAbuPct_LostFocus()
If IsNumeric(txtLoAbuPct.Text) Then
   glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefinePctLowAbundance = Abs(CDbl(txtLoAbuPct.Text))
Else
   MsgBox "This argument should be non-negative number.", vbOKOnly, glFGTU
   txtLoAbuPct.SetFocus
End If
End Sub

Private Sub txtLoCnt_LostFocus()
If IsNumeric(txtLoCnt.Text) Then
   glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineMinLength = Abs(CLng(txtLoCnt.Text))
Else
   MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
   txtLoCnt.SetFocus
End If
End Sub

Private Sub txtPercentMaxAbuToUseToGaugeLength_LostFocus()
If IsNumeric(txtPercentMaxAbuToUseToGaugeLength.Text) Then
    glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefinePercentMaxAbuToUseForLength = Abs(CLng(txtPercentMaxAbuToUseToGaugeLength.Text))
Else
   MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
   txtPercentMaxAbuToUseToGaugeLength.SetFocus
End If
End Sub

Private Sub txtMultiplicity_LostFocus()
Dim tmp As String
tmp = Trim$(txtMultiplicity.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 Then
      f1Multiplicity = CLng(tmp)
      Exit Sub
   End If
End If
'at this point something is wrong
MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
txtMultiplicity.SetFocus
End Sub


Private Function GroupByMWDistanceScanCount(CurrMWDist As Double) As Long
'------------------------------------------------------------------------
'groups classes if they are on certain mass distance (within tolerance)
'and scan separation from each other
'returns number of groups if successful, -1 on error
'------------------------------------------------------------------------
Dim i As Long, j As Long
Dim GroupsCnt As Long
Dim MyGroupCnt As Long
Dim MinInd As Long
Dim MaxInd As Long
Dim CurrMW As Double
On Error GoTo err_GroupByMWDistanceScanCount

'fill search object with sorted class molecular masses
If Not MWSearch.Fill(UMCMW) Then GoTo err_GroupByMWDistanceScanCount
With tmp
    GroupsCnt = 0
    For i = 0 To .UMCCnt - 1
        MyGroupCnt = 0
        CurrMW = UMCMW(i)
        MinInd = 0
        MaxInd = -1
        Call MWSearch.FindIndexRange(CurrMW + CurrMWDist, f1MWTol, MinInd, MaxInd)
        If MinInd < MaxInd Then
           For j = MinInd To MaxInd
               If UMCMW(j) > UMCMW(i) Then        'dont want to make same grouping twice
                  If ScanCloseClasses(UMCInd(i), UMCInd(j)) Then
                     MyGroupCnt = MyGroupCnt + 1
                  End If
               End If
           Next j
        End If
        If MyGroupCnt > 0 Then      'here we count only multi member groups
           GroupsCnt = GroupsCnt + 1
        End If
    Next i
End With
GroupByMWDistanceScanCount = GroupsCnt
Exit Function

err_GroupByMWDistanceScanCount:
GroupByMWDistanceScanCount = -1
End Function


Private Sub PickParameters()
'--------------------------------------------------------------------
'make sure that option changes are accepted before applying functions
'--------------------------------------------------------------------
Call txtf1MWDist_LostFocus
Call txtf1MWTol_LostFocus
Call txtf1ScanDist_LostFocus
Call txtMultiplicity_LostFocus
Call txtAutoMergeMinMW_LostFocus
Call txtAutoMergemaxmw_LostFocus
Call txtHiAbuPct_LostFocus
Call txtLoAbuPct_LostFocus
Call txtLoCnt_LostFocus
Call txtHiCnt_LostFocus
Call txtPercentMaxAbuToUseToGaugeLength_LostFocus
End Sub


Private Function MergeClassesAuto(ByVal UMCInd1 As Long, ByVal UMCInd2 As Long) As Long
'--------------------------------------------------------------------------------------
'merge elements of classes UMCInd1 and UMCInd2 in one class; resulting class has index
'of UMCInd1 or UMCInd2 depending on amType
'class that gets merged is not removed - its index in TmpInc is marked with -1; removal
'of all marked classes is done with RemoveClasses
'Function returns index of MIP class; -1 on any error
'NOTE: If called with UMCInd1=UMCInd2 function will mark class for removal
'--------------------------------------------------------------------------------------
Dim MIPClassInd As Long         'more important class index(merge result)
Dim LIPClassInd As Long         'less important class index(merge victim)
'peaks that has to be added to More Important Class
Dim PeakCnt As Long
Dim PeakType() As Long
Dim PeakInd() As Long
Dim i As Long
On Error GoTo err_MergeClassesAuto
If UMCInd1 = UMCInd2 Then
   TmpInc(UMCInd1) = REMOVE_UMC_MARK
Else
   With GelUMC(CallerID)
        Select Case amType
        Case AUTOMERGE_HI_ABU
             If .UMCs(UMCInd1).ClassAbundance > .UMCs(UMCInd2).ClassAbundance Then
                MIPClassInd = UMCInd1
                LIPClassInd = UMCInd2
             Else
                MIPClassInd = UMCInd2
                LIPClassInd = UMCInd1
             End If
        Case AUTOMERGE_LO_ABU
             If .UMCs(UMCInd1).ClassAbundance < .UMCs(UMCInd2).ClassAbundance Then
                MIPClassInd = UMCInd1
                LIPClassInd = UMCInd2
             Else
                MIPClassInd = UMCInd2
                LIPClassInd = UMCInd1
             End If
        Case AUTOMERGE_HI_CNT
             If .UMCs(UMCInd1).ClassCount > .UMCs(UMCInd2).ClassCount Then
                MIPClassInd = UMCInd1
                LIPClassInd = UMCInd2
             Else
                MIPClassInd = UMCInd2
                LIPClassInd = UMCInd1
             End If
        Case AUTOMERGE_LO_CNT
             If .UMCs(UMCInd1).ClassCount < .UMCs(UMCInd2).ClassCount Then
                MIPClassInd = UMCInd1
                LIPClassInd = UMCInd2
             Else
                MIPClassInd = UMCInd2
                LIPClassInd = UMCInd1
             End If
        Case AUTOMERGE_HI_MW
             If .UMCs(UMCInd1).ClassMW > .UMCs(UMCInd2).ClassMW Then
                MIPClassInd = UMCInd1
                LIPClassInd = UMCInd2
             Else
                MIPClassInd = UMCInd2
                LIPClassInd = UMCInd1
             End If
        Case AUTOMERGE_LO_MW
             If .UMCs(UMCInd1).ClassMW < .UMCs(UMCInd2).ClassMW Then
                MIPClassInd = UMCInd1
                LIPClassInd = UMCInd2
             Else
                MIPClassInd = UMCInd2
                LIPClassInd = UMCInd1
             End If
        End Select
   End With
   With GelUMC(CallerID).UMCs(LIPClassInd)
       'redimension to highest number of peaks that could be added
       ReDim PeakType(.ClassCount - 1)
       ReDim PeakInd(.ClassCount - 1)
       PeakCnt = 0
       For i = 0 To .ClassCount - 1
           If Not IsClassMember(.ClassMType(i), .ClassMInd(i), MIPClassInd) Then
              PeakCnt = PeakCnt + 1
              PeakType(PeakCnt - 1) = .ClassMType(i)
              PeakInd(PeakCnt - 1) = .ClassMInd(i)
           End If
       Next i
       If PeakCnt > 0 Then
          If PeakCnt < .ClassCount Then 'redimension if necessary
             ReDim Preserve PeakType(PeakCnt - 1)
             ReDim Preserve PeakInd(PeakCnt - 1)
          End If
          If Not AddPeakArrayToTheClass(PeakType(), PeakInd(), MIPClassInd) Then GoTo err_MergeClassesAuto
       End If
   End With
   If Not RecalculateClass(MIPClassInd) Then GoTo err_MergeClassesAuto
   TmpInc(LIPClassInd) = REMOVE_UMC_MARK
End If
MergeClassesAuto = MIPClassInd
Exit Function

err_MergeClassesAuto:
MergeClassesAuto = -1
End Function


Private Function IsAcceptableForAutoMerge(ByVal GroupIndex As Long) As Boolean
'-----------------------------------------------------------------------------
'returns True if group GroupIndex satisfies auto-merge conditions; this func.
'is isolated in case we introduce more options in the future
'-----------------------------------------------------------------------------
Dim i As Long
Dim MW As Double
On Error GoTo exit_IsAcceptableForAutoMerge
With GrRes.Members(GroupIndex)
    For i = 0 To .Count - 1
        MW = GelUMC(CallerID).UMCs(.Members(i)).ClassMW
        If MW < amMWMin Or MW > amMWMax Then GoTo exit_IsAcceptableForAutoMerge
    Next i
End With
IsAcceptableForAutoMerge = True
exit_IsAcceptableForAutoMerge:
End Function



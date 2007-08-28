VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUMCIonNet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LC-MS Feature (UMC) Ion Networks"
   ClientHeight    =   5850
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   141
      Top             =   5280
      Width           =   975
   End
   Begin TabDlg.SSTab tbsTabStrip 
      Height          =   5055
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "1. Find Connections"
      TabPicture(0)   =   "frmUMCIonNet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLCMSFeatureFinderInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraNet(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraUMCScope"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkUseLCMSFeatureFinder"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraDREAMS"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAbortFindConnections"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFindConnectionsThenUMCs"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "2. Edit/Filter Connections"
      TabPicture(1)   =   "frmUMCIonNet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "lblFilterConnections"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "3. Define LC-MS Features using Connections"
      TabPicture(2)   =   "frmUMCIonNet.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraNet(1)"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdFindConnectionsThenUMCs 
         Caption         =   "&Find Connections then LC-MS Features"
         Height          =   615
         Left            =   8280
         TabIndex        =   58
         ToolTipText     =   "Create Net based on current settings, then Find LC-MS Features"
         Top             =   3840
         Width           =   2175
      End
      Begin VB.CommandButton cmdAbortFindConnections 
         Caption         =   "Abort!"
         Height          =   375
         Left            =   8880
         TabIndex        =   151
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame fraDREAMS 
         Caption         =   "DREAMS Options"
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
         Begin VB.OptionButton optEvenOddScanFilter 
            Caption         =   "Process Odd / Even sequentially"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   1605
         End
         Begin VB.OptionButton optEvenOddScanFilter 
            Caption         =   "Even-numbered spectra only"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1485
         End
         Begin VB.OptionButton optEvenOddScanFilter 
            Caption         =   "Use all spectra"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optEvenOddScanFilter 
            Caption         =   "Odd-numbered spectra only"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   500
            Width           =   1485
         End
      End
      Begin VB.CheckBox chkUseLCMSFeatureFinder 
         Caption         =   "Use LCMSFeatureFinder external app"
         Height          =   255
         Left            =   240
         TabIndex        =   147
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2985
      End
      Begin VB.Frame fraUMCScope 
         Caption         =   "Definition Scope"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1815
         Begin VB.OptionButton optDefScope 
            Caption         =   "&All Data Points"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   280
            Width           =   1455
         End
         Begin VB.OptionButton optDefScope 
            Caption         =   "&Current View"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Net Edit"
         Height          =   1695
         Left            =   -74640
         TabIndex        =   60
         Top             =   1260
         Width           =   4575
         Begin VB.TextBox txtNetEditTooDistant 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            TabIndex        =   63
            Text            =   "0.1"
            Top             =   1080
            Width           =   585
         End
         Begin VB.CommandButton cmdRemoveLongConnections 
            Caption         =   "Start"
            Height          =   375
            Left            =   3000
            TabIndex        =   64
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblNetInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Net Info"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   4215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Caption         =   "Eliminate connections longer than"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame fraNet 
         Height          =   4455
         Index           =   1
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   10695
         Begin VB.CommandButton cmdFindUMCsUsingNETConnections 
            Caption         =   "Find &LC-MS Features"
            Height          =   495
            Left            =   9240
            TabIndex        =   139
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdAbortProcessing 
            Caption         =   "Abort!"
            Height          =   375
            Left            =   9240
            TabIndex        =   140
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdReportUMC 
            Caption         =   "&Report"
            Height          =   375
            Left            =   9360
            TabIndex        =   138
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdResetToDefaults 
            Caption         =   "Set to Defaults"
            Height          =   375
            Index           =   1
            Left            =   9120
            TabIndex        =   137
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbUMCDrawType 
            Height          =   315
            ItemData        =   "frmUMCIonNet.frx":0054
            Left            =   1680
            List            =   "frmUMCIonNet.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   4035
            Width           =   2175
         End
         Begin VB.TextBox txtInterpolateMaxGapSize 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7200
            TabIndex        =   134
            Text            =   "0"
            Top             =   4035
            Width           =   495
         End
         Begin VB.CheckBox chkInterpolateMissingIons 
            Caption         =   "Interpolate abundances across gaps"
            Height          =   255
            Left            =   4080
            TabIndex        =   132
            Top             =   3765
            Width           =   3015
         End
         Begin VB.Frame fraLCMSFeatureStats 
            Caption         =   "LC-MS Feature Stats"
            Height          =   3735
            Left            =   120
            TabIndex        =   66
            Top             =   180
            Width           =   3735
            Begin VB.ComboBox cboMolecularMassField 
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   149
               Top             =   2760
               Width           =   1935
            End
            Begin VB.CheckBox chkUseMostAbuChargeStateStatsForClassStats 
               Caption         =   "Use most abundant charge state group stats for class stats"
               Height          =   530
               Left            =   240
               TabIndex        =   145
               ToolTipText     =   "Make single-member classes from unconnected nodes"
               Top             =   3120
               Width           =   2055
            End
            Begin VB.ComboBox cboChargeStateAbuType 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   143
               Top             =   2360
               Width           =   3255
            End
            Begin VB.ComboBox cmbUMCRepresentative 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   540
               Width           =   3255
            End
            Begin VB.ComboBox cmbUMCMW 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   1740
               Width           =   3255
            End
            Begin VB.ComboBox cmbUMCAbu 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   1140
               Width           =   3255
            End
            Begin VB.CheckBox chkUseUntangledAsSingle 
               Caption         =   "Make single member classes"
               Height          =   530
               Left            =   2400
               TabIndex        =   73
               ToolTipText     =   "Make single-member classes from unconnected nodes"
               Top             =   3120
               Width           =   1215
            End
            Begin VB.Label lblMolecularMassField 
               BackStyle       =   0  'Transparent
               Caption         =   "Mass field to use"
               Height          =   255
               Left            =   240
               TabIndex        =   150
               Top             =   2760
               Width           =   1335
            End
            Begin VB.Label lblChargeStateAbuType 
               BackStyle       =   0  'Transparent
               Caption         =   "Most Abu Charge State Group Type"
               Height          =   255
               Left            =   240
               TabIndex        =   144
               Top             =   2120
               Width           =   3135
            End
            Begin VB.Label Label4 
               Caption         =   "Class Representative"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   67
               Top             =   300
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "Class Mass"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   71
               Top             =   1500
               Width           =   1335
            End
            Begin VB.Label Label4 
               Caption         =   "Class Abundance"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   69
               Top             =   900
               Width           =   1335
            End
         End
         Begin TabDlg.SSTab tbsUMCRefinementOptions 
            Height          =   3375
            Left            =   3960
            TabIndex        =   74
            Top             =   180
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   5953
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Auto-Refine Options"
            TabPicture(0)   =   "frmUMCIonNet.frx":0058
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraOptionFrame(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Split Features Options"
            TabPicture(1)   =   "frmUMCIonNet.frx":0074
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraOptionFrame(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Adv Class Stats"
            TabPicture(2)   =   "frmUMCIonNet.frx":0090
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fraClassAbundanceTopX"
            Tab(2).Control(1)=   "fraClassMassTopX"
            Tab(2).ControlCount=   2
            Begin VB.Frame fraClassMassTopX 
               Caption         =   "Class Mass Top X"
               Height          =   1215
               Left            =   -74880
               TabIndex        =   125
               Top             =   1800
               Width           =   4095
               Begin VB.TextBox txtClassMassTopXMinAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   127
                  Text            =   "0"
                  Top             =   240
                  Width           =   900
               End
               Begin VB.TextBox txtClassMassTopXMaxAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   129
                  Text            =   "0"
                  ToolTipText     =   "Maximum abundance to include; use 0 to indicate there infinitely large abundance"
                  Top             =   540
                  Width           =   900
               End
               Begin VB.TextBox txtClassMassTopXMinMembers 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   131
                  Text            =   "3"
                  Top             =   840
                  Width           =   900
               End
               Begin VB.Label lblClassMassTopXMinAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   126
                  Top             =   270
                  Width           =   2535
               End
               Begin VB.Label lblClassMassTopXMaxAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Maximum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   128
                  Top             =   560
                  Width           =   2535
               End
               Begin VB.Label lblClassMassTopXMinMembers 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum members to include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   130
                  Top             =   870
                  Width           =   2535
               End
            End
            Begin VB.Frame fraClassAbundanceTopX 
               Caption         =   "Class Abundance Top X"
               Height          =   1215
               Left            =   -74880
               TabIndex        =   118
               Top             =   480
               Width           =   4095
               Begin VB.TextBox txtClassAbuTopXMinMembers 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   124
                  Text            =   "3"
                  Top             =   840
                  Width           =   900
               End
               Begin VB.TextBox txtClassAbuTopXMaxAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   122
                  Text            =   "0"
                  ToolTipText     =   "Maximum abundance to include; use 0 to indicate there infinitely large abundance"
                  Top             =   540
                  Width           =   900
               End
               Begin VB.TextBox txtClassAbuTopXMinAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   120
                  Text            =   "0"
                  Top             =   240
                  Width           =   900
               End
               Begin VB.Label lblClassAbuTopXMinMembers 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum members to include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   123
                  Top             =   870
                  Width           =   2535
               End
               Begin VB.Label lblClassAbuTopXMaxAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Maximum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   121
                  Top             =   560
                  Width           =   2535
               End
               Begin VB.Label lblClassAbuTopXMinAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   119
                  Top             =   270
                  Width           =   2535
               End
            End
            Begin VB.Frame fraOptionFrame 
               Height          =   2920
               Index           =   1
               Left            =   -74880
               TabIndex        =   97
               Top             =   330
               Width           =   4300
               Begin VB.TextBox txtSplitUMCsStdDevMultiplierForSplitting 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   103
                  Text            =   "1"
                  Top             =   900
                  Width           =   495
               End
               Begin VB.ComboBox cboSplitUMCsScanGapBehavior 
                  Height          =   315
                  Left            =   1800
                  Style           =   2  'Dropdown List
                  TabIndex        =   117
                  Top             =   2580
                  Width           =   2295
               End
               Begin VB.TextBox txtHoleSize 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   114
                  Text            =   "3"
                  Top             =   2220
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsPeakPickingMinimumWidth 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   111
                  Text            =   "4"
                  Top             =   1890
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   108
                  Text            =   "15"
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsMaximumPeakCount 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   105
                  Text            =   "6"
                  Top             =   1230
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsMinimumDifferenceInAvgPpmMass 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   100
                  Text            =   "4"
                  Top             =   570
                  Width           =   495
               End
               Begin VB.CheckBox chkSplitUMCsByExaminingAbundance 
                  Caption         =   "Split LC-MS Features by examining abundance"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   98
                  Top             =   240
                  Width           =   3975
               End
               Begin VB.Label lblSplitUMCsStdDevMultiplierForSplitting 
                  Caption         =   "Mass Std Dev threshold multiplier"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   102
                  Top             =   930
                  Width           =   2700
               End
               Begin VB.Label lblSplitUMCsScanGapBehavior 
                  Caption         =   "Scan gap behavior:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   116
                  Top             =   2610
                  Width           =   1620
               End
               Begin VB.Label lblUnits 
                  Caption         =   "scans"
                  Height          =   255
                  Index           =   3
                  Left            =   3480
                  TabIndex        =   115
                  Top             =   2250
                  Width           =   600
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Max size of scan gap in the feature:"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   113
                  Top             =   2250
                  Width           =   2655
               End
               Begin VB.Label lblSplitUMCsPeakPickingMinimumWidth 
                  Caption         =   "Peak picking minimum width"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   110
                  Top             =   1920
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "scans"
                  Height          =   255
                  Index           =   5
                  Left            =   3480
                  TabIndex        =   112
                  Top             =   1920
                  Width           =   600
               End
               Begin VB.Label lblSplitUMCsPeakDetectIntensityThresholdPercentageOfMax 
                  Caption         =   "Peak picking intensity threshold"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   107
                  Top             =   1590
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "% of max"
                  Height          =   255
                  Index           =   0
                  Left            =   3480
                  TabIndex        =   109
                  Top             =   1590
                  Width           =   705
               End
               Begin VB.Label lblSplitUMCsMaximumPeakCount 
                  Caption         =   "Maximum peak count to split feature"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   104
                  Top             =   1260
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "peaks"
                  Height          =   255
                  Index           =   1
                  Left            =   3480
                  TabIndex        =   106
                  Top             =   1260
                  Width           =   600
               End
               Begin VB.Label lblSplitUMCsMinimumDifferenceInAvgPpmMass 
                  Caption         =   "Minimum difference in average mass"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   99
                  Top             =   600
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "ppm"
                  Height          =   255
                  Index           =   2
                  Left            =   3480
                  TabIndex        =   101
                  Top             =   600
                  Width           =   600
               End
            End
            Begin VB.Frame fraOptionFrame 
               Height          =   2700
               Index           =   0
               Left            =   120
               TabIndex        =   75
               Top             =   300
               Width           =   4545
               Begin VB.TextBox txtMaxLengthPctAllScans 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   89
                  Text            =   "15"
                  Top             =   1520
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveMaxLengthPctAllScans 
                  Caption         =   "Remove cls. with length over"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   88
                  Top             =   1520
                  Width           =   2535
               End
               Begin VB.TextBox txtPercentMaxAbuToUseToGaugeLength 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   92
                  Text            =   "33"
                  Top             =   1840
                  Width           =   495
               End
               Begin VB.TextBox txtAutoRefineMinimumMemberCount 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   96
                  Text            =   "3"
                  Top             =   2300
                  Width           =   495
               End
               Begin VB.CheckBox chkRefineUMCLengthByScanRange 
                  Caption         =   "Test feature length using scan range"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   94
                  ToolTipText     =   "If True, then considers scan range for the length tests; otherwise, considers member count"
                  Top             =   2200
                  Value           =   1  'Checked
                  Width           =   1695
               End
               Begin VB.CheckBox chkRemoveLoAbu 
                  Caption         =   "Remove low intensity classes"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   76
                  Top             =   240
                  Width           =   2550
               End
               Begin VB.TextBox txtLoAbuPct 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   77
                  Text            =   "30"
                  Top             =   240
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveHiAbu 
                  Caption         =   "Remove high intensity classes"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   79
                  Top             =   560
                  Width           =   2550
               End
               Begin VB.TextBox txtHiAbuPct 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   80
                  Text            =   "30"
                  Top             =   560
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveLoCnt 
                  Caption         =   "Remove cls. with less than"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   82
                  Top             =   880
                  Width           =   2295
               End
               Begin VB.TextBox txtLoCnt 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   83
                  Text            =   "3"
                  Top             =   880
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveHiCnt 
                  Caption         =   "Remove cls. with length over"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   85
                  Top             =   1200
                  Width           =   2535
               End
               Begin VB.TextBox txtHiCnt 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   86
                  Text            =   "500"
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "%"
                  Height          =   255
                  Index           =   5
                  Left            =   3600
                  TabIndex        =   81
                  Top             =   590
                  Width           =   270
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "%"
                  Height          =   255
                  Index           =   4
                  Left            =   3600
                  TabIndex        =   78
                  Top             =   270
                  Width           =   270
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "% all scans"
                  Height          =   255
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   90
                  Top             =   1545
                  Width           =   855
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "%"
                  Height          =   255
                  Index           =   2
                  Left            =   3600
                  TabIndex        =   93
                  Top             =   1870
                  Width           =   285
               End
               Begin VB.Label lblPercentMaxAbuToUseToGaugeLength 
                  Caption         =   "Percent max abu for gauging width"
                  Height          =   240
                  Left            =   360
                  TabIndex        =   91
                  Top             =   1845
                  Width           =   2565
               End
               Begin VB.Label lblAutoRefineMinimumMemberCount 
                  Caption         =   "Minimum member count:"
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   95
                  Top             =   2200
                  Width           =   1125
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "members"
                  Height          =   255
                  Index           =   0
                  Left            =   3600
                  TabIndex        =   84
                  Top             =   915
                  Width           =   900
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "members"
                  Height          =   255
                  Index           =   1
                  Left            =   3600
                  TabIndex        =   87
                  Top             =   1230
                  Width           =   900
               End
            End
         End
         Begin VB.Label lblUMCDrawType 
            BackStyle       =   0  'Transparent
            Caption         =   "FeatureDraw Type"
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   4065
            Width           =   1455
         End
         Begin VB.Label lblMaxGapSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum size of gap to interpolate:"
            Height          =   255
            Left            =   4080
            TabIndex        =   133
            Top             =   4065
            Width           =   2535
         End
      End
      Begin VB.Frame fraNet 
         Height          =   3255
         Index           =   0
         Left            =   2040
         TabIndex        =   9
         Top             =   420
         Width           =   8535
         Begin VB.CommandButton cmdResetToOldDefaults 
            Caption         =   "Set to Old Defaults"
            Height          =   250
            Left            =   5280
            TabIndex        =   146
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdResetToDefaults 
            Caption         =   "Set to Defaults"
            Height          =   375
            Index           =   0
            Left            =   6960
            TabIndex        =   14
            Top             =   200
            Width           =   1455
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   4
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   2222
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   3
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1860
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   2
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1500
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   1
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1140
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   0
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   780
            Width           =   855
         End
         Begin VB.CommandButton cmdFindConnections 
            Caption         =   "&Find Connections Only"
            Height          =   375
            Left            =   6000
            TabIndex        =   57
            ToolTipText     =   "Create Net based on current settings"
            Top             =   2760
            Width           =   2415
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   6720
            TabIndex        =   53
            Text            =   "0.1"
            Top             =   2222
            Width           =   735
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   6720
            TabIndex        =   45
            Text            =   "0.1"
            Top             =   1860
            Width           =   735
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   37
            Text            =   "0.1"
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   6720
            TabIndex        =   29
            Text            =   "0.1"
            Top             =   1140
            Width           =   735
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   4
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   2222
            Width           =   975
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   3
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1860
            Width           =   975
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   2
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1500
            Width           =   975
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   1
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6720
            TabIndex        =   21
            Text            =   "0.1"
            Top             =   780
            Width           =   735
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   0
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   780
            Width           =   975
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   4080
            TabIndex        =   50
            Text            =   "1"
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   4080
            TabIndex        =   42
            Text            =   "1"
            Top             =   1860
            Width           =   615
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   4
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2220
            Width           =   2175
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   3
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1860
            Width           =   2175
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   47
            Top             =   2280
            Width           =   700
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   1920
            Width           =   700
         End
         Begin VB.TextBox txtNETType 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   13
            Text            =   "1"
            Top             =   300
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtRejectLongConnections 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   56
            Text            =   "1"
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox cmbMetricType 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   4080
            TabIndex        =   34
            Text            =   "1"
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   4080
            TabIndex        =   26
            Text            =   "1"
            Top             =   1140
            Width           =   615
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   2
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1500
            Width           =   2175
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   1
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1140
            Width           =   2175
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   31
            Top             =   1560
            Width           =   700
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   4080
            TabIndex        =   18
            Text            =   "1"
            Top             =   780
            Width           =   615
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   0
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   780
            Width           =   2175
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   840
            Width           =   700
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   10
            Left            =   4800
            TabIndex        =   51
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   9
            Left            =   4800
            TabIndex        =   43
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   8
            Left            =   4800
            TabIndex        =   35
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   7
            Left            =   4800
            TabIndex        =   27
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   6
            Left            =   4800
            TabIndex        =   19
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Reject connection longer than"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   2790
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   49
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   4
            Left            =   3240
            TabIndex        =   41
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblNETType 
            Caption         =   "Net Type"
            Height          =   255
            Left            =   3360
            TabIndex        =   12
            Top             =   320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Metric Type"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   33
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   25
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   17
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Label lblLCMSFeatureFinderInfo 
         Caption         =   $"frmUMCIonNet.frx":00AC
         Height          =   400
         Left            =   240
         TabIndex        =   148
         Top             =   4080
         Width           =   5415
      End
      Begin VB.Label lblFilterConnections 
         Caption         =   $"frmUMCIonNet.frx":0147
         Height          =   615
         Left            =   -74640
         TabIndex        =   59
         Top             =   540
         Width           =   5055
      End
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   142
      Top             =   5340
      Width           =   9855
   End
End
Attribute VB_Name = "frmUMCIonNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'created: 04/04/2003 nt
'last modified: 05/29/2003 nt
'----------------------------------------------------------------------
'This calculation is very time consuming in original general form
'Optimization is used on first dimension to eliminate uneccessary
'calculation; therefore the efficiency of the algorithm also
'depends on the selection of first dimension (use mass!)
'MetricData members of MyDef structure determines which data dimensions
'to use and how dimensions should be weighted. Constraints can limit
'distance in any single dimension to be "Less Than" or "Greater Than"
'predefined constraint value. Important to notice is that constraints
'work on distances between points; not the data itself
'----------------------------------------------------------------------
Option Explicit

Private Const MAX_NET_SIZE = 1000000

Private Const NET_ADD_RATE = 5000

Private Const NET_EDIT_REJECT_LONG = 0

Private Const HUMCNotUsed As Byte = 0
Private Const HUMCInUse As Byte = 1
Private Const HUMCUsed As Byte = 2

Private Const LCMS_FEATURE_FINDER_APP_NAME As String = "LCMSFeatureFinder.exe"
Private Const LCMS_FEATURE_FINDER_ISOTOPE_FEATURES_FILE As String = "Tmp_Export_LCMSFeaturesToSearch.txt"
Private Const LCMS_FEATURE_FINDER_INI_FILE As String = "Tmp_Export_LCMSFeaturesToSearch.ini"

Private CallerID As Long
Private bLoading As Boolean

Private DataCnt As Long     'count of isotopic data
                            ' If only finding LC-MS Features on data "in current view", then this value may be
                            ' less than the actual data count in the file

' Unused variable
'''Dim DataWeightFactor() As Double    'weighting factor for each dimension
Private DataOInd() As Long              'original index in IsoData array; thus, pointer into GelDraw(CallerID).IsoID()
Private DataVal() As Double             'values to be used in calculations
'this values are dimensioned and weighted to improve calculation speed

Private ResCnt As Long
Private ResInd1() As Long
Private ResInd2() As Long
Private ResDist() As Double
Private ResEliminate() As Boolean

Private MinScan As Long
Private MaxScan As Long

'following arrays are used to optimize calculations by indexing first dimension
Private OptIndO() As Long         'indexes in original data arrays
Private OptValO() As Double       'values in first data dimension used in optimization

Dim MyDef As UMCIonNetDefinition


'settings used to define UMCs from Net
Private UMCMakeSingleMemberClasses As Boolean
Private UMCRepresentative As Long

'helper variables used to fill classes
Private HUMCIsoCnt As Long        'number of Isotopic distributions in a current 2D display
Private HUMCNetCnt As Long        'number of connections in Net
Private HUMCIsoUsed() As Byte     'array parallel with IsoData array indicating should isotopic
                              'distribution be included in the current class; it also helps
                              'if unconnected nodes should be made to classes
Private HUMCEquClsWk() As Long    'array of the same size as IsoData array used to construct each class
                              'this is working array that is never reinitialized to optimize
                              'performance; that means be very careful with it's content
Private HUMCEquClsCnt As Long     'actual size of the current class
Private HUMCEquCls() As Long      'array that will hold actual class of equivalency
Dim HUMCNetUsed() As Byte     'array parallel with NetInd arrays indicating if net connection
                              'was already used in classification
Private DummyInd() As Long        'never to be initialized; used in sort function

Private mLCMSResultsMappingCount As Long
Private mLCMSResultsMappingUMCs() As Long
Private mLCMSResultsMappingDataIndices() As Long

Private mSplitUMCs As clsSplitUMCsByAbundance

Private mAbortProcess As Boolean
Private mCalculating As Boolean
Private mOneSecond As Double

Private Sub AbortProcessing()
    mAbortProcess = True
    
    On Error Resume Next
    If Not mSplitUMCs Is Nothing Then
        mSplitUMCs.AbortProcessingNow
    End If
End Sub

Private Function BuildCurrentClass() As Boolean
'---------------------------------------------------------------------------------------
'builds class for the current settings in the HUMCEquCls array; returns True on success
'class has to be sorted if more than 2 elements (to preserve scan order)
'---------------------------------------------------------------------------------------
Dim i As Long
Dim BestInd As Long
'Dim MySort As New QSLong
On Error GoTo err_BuildCurrentClass

' ToDo: Maybe add parameter blnSortFeatures to this function and don't sort if blnSortFeatures=false
If HUMCEquClsCnt > 2 Then
   ShellSortLong HUMCEquCls, 0, HUMCEquClsCnt - 1
   'If Not MySort.QSAsc(HUMCEquCls(), DummyInd()) Then GoTo err_BuildCurrentClass
   'Set MySort = Nothing
End If

With GelUMC(CallerID)
        
    If .UMCCnt > UBound(.UMCs) Then             'add room if neccessary
        If Not ManageClasses(CallerID, UMCManageConstants.UMCMngAdd) Then GoTo err_BuildCurrentClass
    End If
    
    With .UMCs(.UMCCnt)
          ' ToDo: Consider only expanding the memory used, never contracting
          
          ReDim .ClassMInd(HUMCEquClsCnt - 1)
          ReDim .ClassMType(HUMCEquClsCnt - 1)
          For i = 0 To HUMCEquClsCnt - 1
              .ClassCount = .ClassCount + 1
              .ClassMInd(.ClassCount - 1) = HUMCEquCls(i)
              .ClassMType(.ClassCount - 1) = glIsoType
          Next i
                    
          ' Note: This code has been moved to UMCIonNet.Bas->FindUMCClassRepIndex
          BestInd = FindUMCClassRepIndex(CallerID, GelUMC(CallerID).UMCCnt, CInt(UMCRepresentative))
          
          .ClassRepInd = .ClassMInd(BestInd)
          .ClassRepType = glIsoType
     End With
      .UMCCnt = .UMCCnt + 1
      
     BuildCurrentClass = True
     Exit Function
End With

err_BuildCurrentClass:
ChangeStatus " Error building LC-MS Feature."
End Function

Private Function BuildUMCsUsingmLCMSResultsMapping(ByVal blnShowMessages As Boolean) As Boolean
    Dim lngIndex As Long
    Dim lngCurrentUMC As Long
    Dim intScopeUsedForConnections As Integer
    
    Dim strBaseStatus As String
    Dim strMessage As String
    
    Dim blnSuccess As Boolean
    
On Error GoTo BuildUMCsUsingmLCMSResultsMappingErrorHandler
    blnSuccess = False
    
    If mLCMSResultsMappingCount = 0 Then
        strMessage = "LC-MS results mapping data not in memory; unable to continue"
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            AddToAnalysisHistory CallerID, strMessage
        ElseIf blnShowMessages Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "Invalid Options"
        End If
    ElseIf Not ValidateClassStatOptions() Then
        strMessage = "Invalid class stat options, unable to continue"
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            AddToAnalysisHistory CallerID, strMessage
        ElseIf blnShowMessages Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "Invalid Options"
        End If
    Else
        ' Update GelUMC(CallerID).def now, prior to processing
        With GelUMC(CallerID)
            ' Save the Scope used when finding the connections since
            '  the user may have changed it since then, thus affecting UMCDef
            intScopeUsedForConnections = .def.DefScope
            
            ' Copy the def
            .def = UMCDef
            
            ' Make sure the scope in GelUMC() is correct
            .def.DefScope = intScopeUsedForConnections
        End With
            
        ChangeStatus " Initializing LC-MS Feature structures..."
        
        If Not ManageClasses(CallerID, UMCManageConstants.UMCMngInitialize) Then
            strMessage = "Error initializing LC-MS Feature structures, unable to continue"
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                AddToAnalysisHistory CallerID, strMessage
            ElseIf blnShowMessages Then
                MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
            End If
            blnSuccess = False
        Else
            strBaseStatus = "Defining LC-MS Feature members"
            ChangeStatus strBaseStatus
            
            HUMCEquClsCnt = 0
            ReDim HUMCEquCls(999)
            
            lngCurrentUMC = mLCMSResultsMappingUMCs(0)
            For lngIndex = 0 To mLCMSResultsMappingCount - 1
                  
                If mLCMSResultsMappingUMCs(lngIndex) <> lngCurrentUMC Then
                    BuildCurrentClass
                    
                    lngCurrentUMC = mLCMSResultsMappingUMCs(lngIndex)
                    HUMCEquClsCnt = 0
                End If
                
                If HUMCEquClsCnt = UBound(HUMCEquCls) Then
                    ReDim Preserve HUMCEquCls((UBound(HUMCEquCls) + 1) * 2 - 1)
                End If
                
                HUMCEquCls(HUMCEquClsCnt) = mLCMSResultsMappingDataIndices(lngIndex)
                HUMCEquClsCnt = HUMCEquClsCnt + 1
                
                If lngIndex Mod 1000 = 999 Then
                    ChangeStatus strBaseStatus & ": " & (lngIndex + 1) & " / " & mLCMSResultsMappingCount
                End If
                
                If mAbortProcess Then Exit For
            Next lngIndex
                
            If mAbortProcess Then
                ChangeStatus "Processing aborted."
                blnSuccess = False
            Else
                If HUMCEquClsCnt > 0 Then
                    ' Store the data for the final LC-MS Feature
                    BuildCurrentClass
                End If
            
                ' Refine the LC-MS Features and compute class stats
                blnSuccess = FinalizeNewUMCs()
                
                If blnSuccess Then
                    ChangeStatus "Number of LC-MS Features: " & GelUMC(CallerID).UMCCnt
                Else
                    ChangeStatus "Error creating LC-MS Features."
                End If
            End If
            
            
            If GelUMCDraw(CallerID).Visible Then
                GelBody(CallerID).RequestRefreshPlot
                GelBody(CallerID).csMyCooSys.CoordinateDraw
            End If
        End If
    End If
    
    BuildUMCsUsingmLCMSResultsMapping = blnSuccess
    Exit Function
    
BuildUMCsUsingmLCMSResultsMappingErrorHandler:
    Debug.Print "Error in BuildUMCsUsingmLCMSResultsMapping: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->BuildUMCsUsingmLCMSResultsMapping"
    
    BuildUMCsUsingmLCMSResultsMapping = False

End Function

Private Sub ChangeStatus(ByVal StatusMsg As String)
    lblStatus.Caption = StatusMsg
    DoEvents
End Sub

Private Function CheckOddEvenIterationForDataPoint(ByVal intOddEvenIteration As Integer, ByVal lngOriginalIndex As Long) As Boolean
    If intOddEvenIteration = 1 Then
        ' Return True if the point has an odd scan number
        CheckOddEvenIterationForDataPoint = GelData(CallerID).IsoData(GelDraw(CallerID).IsoID(DataOInd(lngOriginalIndex))).ScanNumber Mod 2 = 1
    ElseIf intOddEvenIteration = 2 Then
        ' Return True if the point has an even scan number
        CheckOddEvenIterationForDataPoint = GelData(CallerID).IsoData(GelDraw(CallerID).IsoID(DataOInd(lngOriginalIndex))).ScanNumber Mod 2 = 0
    Else
        ' intOddEvenIteration is not 1 or 2; return True
        CheckOddEvenIterationForDataPoint = True
    End If
End Function

Private Function CheckOddEvenIterationForScan(ByVal intOddEvenIteration As Integer, ByVal lngScanNumber As Long) As Boolean
    If intOddEvenIteration = 1 Then
        ' Return True if the point has an odd scan number
        CheckOddEvenIterationForScan = lngScanNumber Mod 2 = 1
    ElseIf intOddEvenIteration = 2 Then
        ' Return True if the point has an even scan number
        CheckOddEvenIterationForScan = lngScanNumber Mod 2 = 0
    Else
        ' intOddEvenIteration is not 1 or 2; return True
        CheckOddEvenIterationForScan = True
    End If
End Function

Private Sub CreateNet()
'------------------------------------------------------------------------------
'fills permanent GelUMCIon structures with indexes;
'before filling permanent GelUMCIon structures results are sorted on Ind1/Ind2;
'this will optimize class creation and reduce the total entropy in the Universe
'------------------------------------------------------------------------------
Dim i As Long
Dim TmpCnt As Long
Dim Ind1() As Long, Ind2() As Long, Dist() As Double, SortInd() As Long
Dim blnEraseUMCIonNetworks As Boolean

On Error GoTo CreateNetErrorHandler

blnEraseUMCIonNetworks = False
ChangeStatus " Creating Net structure..."

If ResCnt > 0 Then
   TmpCnt = 0
   ReDim Ind1(ResCnt - 1):   ReDim Ind2(ResCnt - 1):
   ReDim Dist(ResCnt - 1):   ReDim SortInd(ResCnt - 1)
   For i = 0 To ResCnt - 1
       If Not ResEliminate(i) Then
          TmpCnt = TmpCnt + 1
          Ind1(TmpCnt - 1) = DataOInd(ResInd1(i)):   Ind2(TmpCnt - 1) = DataOInd(ResInd2(i))
          Dist(TmpCnt - 1) = ResDist(i):             SortInd(TmpCnt - 1) = TmpCnt - 1
       End If
   Next i
   Call ManageResArrays(amtErase)           'don't need results arrays anymore
   If TmpCnt > 0 Then
      ReDim Preserve Ind1(TmpCnt - 1):   ReDim Preserve Ind2(TmpCnt - 1):
      ReDim Preserve Dist(TmpCnt - 1):   ReDim Preserve SortInd(TmpCnt - 1)
      ChangeStatus "Sorting connections..."
      Call Sort2LongArrays(Ind1(), Ind2(), SortInd())   'sort results on Ind1, Ind2
      With GelUMCIon(CallerID)
         .NetCount = TmpCnt
         ReDim .NetInd1(TmpCnt - 1):   ReDim .NetInd2(TmpCnt - 1):   ReDim .NetDist(TmpCnt - 1)
         .MinDist = glHugeDouble:      .MaxDist = -glHugeDouble
         For i = 0 To TmpCnt - 1
             .NetInd1(i) = Ind1(SortInd(i)):   .NetInd2(i) = Ind2(SortInd(i)):   .NetDist(i) = Dist(SortInd(i))
             If Dist(SortInd(i)) < .MinDist Then .MinDist = Dist(SortInd(i))
             If Dist(SortInd(i)) > .MaxDist Then .MaxDist = Dist(SortInd(i))
         Next i
      End With
   Else
      blnEraseUMCIonNetworks = True
   End If
Else
   Call ManageResArrays(amtErase)
   blnEraseUMCIonNetworks = True
End If

If blnEraseUMCIonNetworks Then
    With GelUMCIon(CallerID)
        .NetCount = 0
         ReDim .NetInd1(0):   ReDim .NetInd2(0):   ReDim .NetDist(0)
    End With
End If

ChangeStatus " Number of connections: " & GelUMCIon(CallerID).NetCount
lblNetInfo.Caption = GetUMCIonNetInfo(CallerID)
Exit Sub

CreateNetErrorHandler:
Debug.Print "Error in CreateNet: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->CreateNet"
Resume Next

End Sub

Private Sub DisplayCurrentOptions()
    Dim blnLoadingSaved As Boolean
    
    blnLoadingSaved = bLoading
    
    bLoading = True
    
    SetDefinition
    SetUMCDefinition
    
    ' MonroeMod: Added Auto-Refine Options
    With glbPreferencesExpanded.UMCAutoRefineOptions
        SetCheckBox chkRemoveLoCnt, .UMCAutoRefineRemoveCountLow
        SetCheckBox chkRemoveHiCnt, .UMCAutoRefineRemoveCountHigh
        SetCheckBox chkRemoveMaxLengthPctAllScans, .UMCAutoRefineRemoveMaxLengthPctAllScans
        
        txtLoCnt = .UMCAutoRefineMinLength
        txtHiCnt = .UMCAutoRefineMaxLength
        txtMaxLengthPctAllScans = .UMCAutoRefineMaxLengthPctAllScans
        txtPercentMaxAbuToUseToGaugeLength = .UMCAutoRefinePercentMaxAbuToUseForLength
        
        SetCheckBox chkRefineUMCLengthByScanRange, .TestLengthUsingScanRange
        txtAutoRefineMinimumMemberCount = .MinMemberCountWhenUsingScanRange
        UpdateDynamicControls
        
        SetCheckBox chkRemoveLoAbu, .UMCAutoRefineRemoveAbundanceLow
        SetCheckBox chkRemoveHiAbu, .UMCAutoRefineRemoveAbundanceHigh
        txtLoAbuPct = .UMCAutoRefinePctLowAbundance
        txtHiAbuPct = .UMCAutoRefinePctHighAbundance
    
        SetCheckBox chkSplitUMCsByExaminingAbundance, .SplitUMCsByAbundance
        With .SplitUMCOptions
            txtSplitUMCsMaximumPeakCount = Trim(.MaximumPeakCountToSplitUMC)
            txtSplitUMCsMinimumDifferenceInAvgPpmMass = Trim(.MinimumDifferenceInAveragePpmMassToSplit)
            txtSplitUMCsStdDevMultiplierForSplitting = Trim(.StdDevMultiplierForSplitting)
            txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax = Trim(.PeakDetectIntensityThresholdPercentageOfMaximum)
            txtSplitUMCsPeakPickingMinimumWidth = Trim(.PeakWidthPointsMinimum)
            cboSplitUMCsScanGapBehavior.ListIndex = .ScanGapBehavior
        End With
    End With
    
    ' MonroeMod: Added Advanced Stats options
    With glbPreferencesExpanded.UMCAdvancedStatsOptions
        txtClassAbuTopXMinAbu = .ClassAbuTopXMinAbu
        txtClassAbuTopXMaxAbu = .ClassAbuTopXMaxAbu
        txtClassAbuTopXMinMembers = .ClassAbuTopXMinMembers
        
        txtClassMassTopXMinAbu = .ClassMassTopXMinAbu
        txtClassMassTopXMaxAbu = .ClassMassTopXMaxAbu
        txtClassMassTopXMinMembers = .ClassMassTopXMinMembers
    End With
    
    cmbUMCDrawType.ListIndex = GelUMCDraw(CallerID).DrawType

    bLoading = blnLoadingSaved

End Sub

Private Sub DisplayDynamicUnits()
    Dim intIndex As Integer
    Dim blnShowConstraints As Boolean
    Dim blnMassBasedDataDim As Boolean
    
On Error GoTo DisplayDynamicUnitsErrorHandler

    For intIndex = 0 To cmbData.Count - 1
        Select Case cmbData(intIndex).ListIndex
        Case uindUMCIonNetDimConstants.uindMonoMW, uindUMCIonNetDimConstants.uindAvgMW, uindUMCIonNetDimConstants.uindTmaMW
            cmbConstraintUnits(intIndex).Visible = True
            cmbConstraintUnits(intIndex).ListIndex = MyDef.MetricData(intIndex).ConstraintUnits
            blnMassBasedDataDim = True
        Case Else
            cmbConstraintUnits(intIndex).Visible = False
            blnMassBasedDataDim = False
        End Select
        
        If cmbConstraint(intIndex).ListIndex > Net_CT_None Then
            blnShowConstraints = True
        Else
            blnShowConstraints = False
        End If
        
        txtConstraint(intIndex).Visible = blnShowConstraints
        cmbConstraintUnits(intIndex).Visible = blnShowConstraints And blnMassBasedDataDim
        
    Next intIndex
    
    Exit Sub
    
DisplayDynamicUnitsErrorHandler:
    Debug.Print "Error in DisplayDynamicUnits: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->DisplayDynamicUnits"
    Resume Next
    
End Sub

Private Function EliminateLongConnections(ByVal TooLong As Double) As Long
'------------------------------------------------------------
'mark long connections for elimination
'returns the number of connections eliminated
'------------------------------------------------------------
Dim i As Long
Dim Count As Long
On Error Resume Next
ChangeStatus " Rejecting long connections..."
For i = 0 To ResCnt - 1
    If ResDist(i) > TooLong Then
       ResEliminate(i) = True
       Count = Count + 1
    End If
Next i
ChangeStatus " Long connections eliminated: " & Count
EliminateLongConnections = Count
End Function

' Unused function (July 2003)
Private Sub EliminateRedundantConnectionsDirect()
'-----------------------------------------------------------------------------------
'marks redundant connections for eliminations; if we have ResInd1(i)=m, ResInd2(i)=n
'and ResInd1(j)=n, ResInd2(j)=m for some i,j then eliminate i if m>n or j if m<=n
'-----------------------------------------------------------------------------------
Dim i As Long, j As Long
Dim Count As Long
On Error Resume Next
ChangeStatus " Eliminating redundancy..."
For i = 0 To DataCnt - 1
    If Not ResEliminate(i) Then
       For j = i + 1 To DataCnt - 1
           If Not ResEliminate(j) Then
              If ResInd1(i) = ResInd2(j) Then
                 If ResInd2(i) = ResInd1(j) Then
                    Count = Count + 1
                    'mark for elimination one where ResInd1>ResInd2
                    If ResInd1(i) > ResInd2(i) Then
                       ResEliminate(i) = True
                    Else
                       ResEliminate(j) = True
                    End If
                 End If
              End If
           End If
       Next j
    End If
Next i
ChangeStatus " Redundant connections eliminated: " & Count
End Sub


Public Function EliminateLongConnections_Net(TooLongConnection As Double) As Long
'--------------------------------------------------------------------
'eliminates long connections dirtectly from the GelUMCIon structures
'NOTE: this function is used for Net editing purposes different from
'EliminateLongConnections which is used when Net is created
'returns the number of connections eliminated
'--------------------------------------------------------------------
Dim i As Long
Dim TmpCnt As Long
Dim lngOriginalConnectionCount As Long

On Error GoTo EliminateLongConnectionsNetErrorHandler

ChangeStatus " Eliminating long connections..."
With GelUMCIon(CallerID)
    lngOriginalConnectionCount = .NetCount
    If .NetCount > 0 Then
       .MinDist = glHugeDouble:     .MaxDist = -glHugeDouble
       For i = 0 To .NetCount - 1
           If .NetDist(i) <= TooLongConnection Then
              TmpCnt = TmpCnt + 1
              .NetInd1(TmpCnt - 1) = .NetInd1(i)
              .NetInd2(TmpCnt - 1) = .NetInd2(i)
              .NetDist(TmpCnt - 1) = .NetDist(i)
              If .NetDist(TmpCnt - 1) < .MinDist Then .MinDist = .NetDist(TmpCnt - 1)
              If .NetDist(TmpCnt - 1) > .MaxDist Then .MaxDist = .NetDist(TmpCnt - 1)
           End If
       Next i
       If TmpCnt > 0 Then
          ReDim Preserve .NetInd1(TmpCnt - 1)
          ReDim Preserve .NetInd2(TmpCnt - 1)
          ReDim Preserve .NetDist(TmpCnt - 1)
       Else
          Erase .NetDist:   Erase .NetInd1:   Erase .NetInd2
       End If
       .NetCount = TmpCnt
       .ThisNetDef.TooDistant = TooLongConnection
    End If
End With
GelSearchDef(CallerID).UMCIonNetDef = GelUMCIon(CallerID).ThisNetDef
ChangeStatus " Number of connections: " & GelUMCIon(CallerID).NetCount
lblNetInfo.Caption = GetUMCIonNetInfo(CallerID)
EliminateLongConnections_Net = lngOriginalConnectionCount - GelUMCIon(CallerID).NetCount

Exit Function

EliminateLongConnectionsNetErrorHandler:
Debug.Print "Error in EliminateLongConnectionsNet: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->EliminateLongConnections_NET"
Resume Next
End Function

Private Function ExportPeaksForUMCFinding(ByVal strOutputFolder As String, ByRef strLCMSFeaturesFilePath As String, ByRef strIniFilePath As String, ByVal intOddEvenIteration As Integer) As Boolean
    Const COL_DELIMITER As String = vbTab
    
    Dim lngIndex As Long
    Dim ISInd() As Long         ' In-scope index
    
    Dim tsOutfile As TextStream
    Dim fso As New FileSystemObject
    
    Dim strMessage As String
    
    Dim strLineOut As String
    Dim strDimensionName As String
    Dim strConstraint As String
    
    Dim strBaseStatus As String
    
    Dim lngGelScanNumberMin As Long
    Dim lngGelScanNumberMax As Long
    Dim intMinLength As Integer
        
    Dim blnMonoMassDefined As Boolean
    Dim blnAvgMassDefined As Boolean
    Dim blnLogAbundanceDefined As Boolean
    Dim blnScanDefined As Boolean
    Dim blnNETDefined As Boolean
    Dim blnFitDefined As Boolean
    
    Dim blnUseGenericNET As Boolean
    Dim blnExportPoint As Boolean
    
On Error GoTo ExportPeaksForUMCFindingErrorHandler

    strBaseStatus = "Exporting loaded peaks to find LC-MS features with external application"
    ChangeStatus strBaseStatus

    If Not GetDataInScope(ISInd(), DataCnt) Then
        mAbortProcess = True
        ExportPeaksForUMCFinding = False
        Exit Function
    End If
    
    ' Update MyDef.NetDim & MyDef.NetActualDim
    If Not UpdateNetDimInfo() Then
        mAbortProcess = True
        ExportPeaksForUMCFinding = False
        Exit Function
    End If

    ' Write out the data in view
    strLCMSFeaturesFilePath = fso.BuildPath(strOutputFolder, LCMS_FEATURE_FINDER_ISOTOPE_FEATURES_FILE)
    Set tsOutfile = fso.CreateTextFile(strLCMSFeaturesFilePath, True)
    
    ' Write the header line
    strLineOut = "scan_num" & COL_DELIMITER & _
                 "charge" & COL_DELIMITER & _
                 "abundance" & COL_DELIMITER & _
                 "mz" & COL_DELIMITER & _
                 "fit" & COL_DELIMITER & _
                 "average_mw" & COL_DELIMITER & _
                 "monoisotopic_mw" & COL_DELIMITER & _
                 "mostabundant_mw" & COL_DELIMITER & _
                 "fwhm" & COL_DELIMITER & _
                 "signal_noise" & COL_DELIMITER & _
                 "mono_abundance" & COL_DELIMITER & _
                 "mono_plus2_abundance" & COL_DELIMITER & _
                 "index"
                 
    tsOutfile.WriteLine strLineOut
    
    With GelData(CallerID)
        For lngIndex = 1 To DataCnt
            If intOddEvenIteration = 0 Then
                blnExportPoint = True
            Else
                blnExportPoint = CheckOddEvenIterationForScan(intOddEvenIteration, .IsoData(ISInd(lngIndex)).ScanNumber)
            End If
            
            If blnExportPoint Then
                With .IsoData(ISInd(lngIndex))
                    ' Note that we're sending the unaltered scan number to the LCMSFeatureFinder, not the relative scan number returned by LookupScanNumberRelativeIndex
                    ' This is required because we're not providing NET values, just scan numbers.
                    ' If we sent the relative scan number, then the NET values computed by the LCMSFeatureFinder would be wrong (unless we also changed the MinScan and MaxScan values written to the Parameter File, but that would confuse the matter even more)
                    ' Therefore, the default settings for calling the LCMSFeatureFinder are to weight by NET but not by scan number
                    ' If the user does weight by scan number and if they're processing LTQ_FT or LTQ_Orbitrap datasets, then they'll likely need to change the scan-number weighting factor

                    strLineOut = Trim(.ScanNumber) & COL_DELIMITER & _
                                 Trim(.Charge) & COL_DELIMITER & _
                                 Trim(.Abundance) & COL_DELIMITER & _
                                 Trim(.MZ) & COL_DELIMITER & _
                                 Trim(.Fit) & COL_DELIMITER & _
                                 Trim(.AverageMW) & COL_DELIMITER & _
                                 Trim(.MonoisotopicMW) & COL_DELIMITER & _
                                 Trim(.MostAbundantMW) & COL_DELIMITER & _
                                 Trim(.FWHM) & COL_DELIMITER & _
                                 Trim(.SignalToNoise) & COL_DELIMITER & _
                                 Trim(.IntensityMono) & COL_DELIMITER & _
                                 Trim(.IntensityMonoPlus2) & COL_DELIMITER & _
                                 Trim(ISInd(lngIndex))
    
                    tsOutfile.WriteLine strLineOut
                End With
            End If
            
            If lngIndex Mod 5000 = 0 Then
                ChangeStatus strBaseStatus & ": " & Trim(lngIndex) & " / " & Trim(DataCnt)
            End If
            
            If mAbortProcess Then Exit For
        Next lngIndex
    End With
   
    tsOutfile.Close
    Set tsOutfile = Nothing
    
    If mAbortProcess Then
        ExportPeaksForUMCFinding = False
        Exit Function
    End If
    
    ChangeStatus "Exporting parameters for finding LC-MS features with external application"
    
     ' Write out the parameters to use to find the LC-MS Features
    strIniFilePath = fso.BuildPath(strOutputFolder, LCMS_FEATURE_FINDER_INI_FILE)
    Set tsOutfile = fso.CreateTextFile(strIniFilePath, True)
    
    tsOutfile.WriteLine "[UMCCreationOptions]"
    
    blnMonoMassDefined = False
    blnAvgMassDefined = False
    blnLogAbundanceDefined = False
    blnScanDefined = False
    blnNETDefined = False
    blnFitDefined = False
    
    blnUseGenericNET = False
    
    For lngIndex = 0 To MyDef.NetDim - 1
        strDimensionName = ""
        strConstraint = ""
        
        With MyDef.MetricData(lngIndex)
            Select Case .DataType
               Case uindUMCIonNetDimConstants.uindMonoMW
                    strDimensionName = "MonoMass"
                    strConstraint = strDimensionName & "Constraint=" & Trim(.ConstraintValue)
                    blnMonoMassDefined = True
                    
               Case uindUMCIonNetDimConstants.uindAvgMW
                    strDimensionName = "AvgMass"
                    strConstraint = strDimensionName & "Constraint=" & Trim(.ConstraintValue)
                    blnAvgMassDefined = True
                    
               Case uindUMCIonNetDimConstants.uindTmaMW
                    ' The most abundant mass; not valid with the LCMSFeatureFinder
                    Debug.Assert False
                    
               Case uindUMCIonNetDimConstants.uindScan
                    strDimensionName = "Scan"
                    blnScanDefined = True
                    
               Case uindUMCIonNetDimConstants.uindFit
                    strDimensionName = "Fit"
                    blnFitDefined = True
                    
               Case uindUMCIonNetDimConstants.uindMZ
                    ' m/z; not valid with the LCMSFeatureFinder
                    Debug.Assert False
                    
               Case uindUMCIonNetDimConstants.uindGenericNET
                    strDimensionName = "NET"
                    blnNETDefined = True
                    blnUseGenericNET = .Use
                    
               Case uindUMCIonNetDimConstants.uindChargeState
                    ' charge; not valid with the LCMSFeatureFinder
                    Debug.Assert False
                    
               Case uindUMCIonNetDimConstants.uindLogAbundance
                    strDimensionName = "LogAbundance"
                    blnLogAbundanceDefined = True
                    
            End Select
        
            If Len(strDimensionName) > 0 Then
                strLineOut = strDimensionName & "Weight="
                If .Use Then
                    strLineOut = strLineOut & Trim(.WeightFactor)
                Else
                    strLineOut = strLineOut & Trim(0)
                End If
                tsOutfile.WriteLine strLineOut
                
                If Len(strConstraint) > 0 Then
                    tsOutfile.WriteLine strConstraint
                    
                    strLineOut = strDimensionName & "ConstraintIsPPM="
                    If .ConstraintUnits = DATA_UNITS_MASS_DA Then
                        strLineOut = strLineOut & "False"
                    Else
                        strLineOut = strLineOut & "True"
                    End If
                    tsOutfile.WriteLine strLineOut
                End If
                
            ElseIf .Use Then
                ' Unknown or inappropriate parameter; this is unexpected
                Debug.Assert False
            End If
        End With
    Next lngIndex

    If Not blnMonoMassDefined Then tsOutfile.WriteLine "MonoMassWeight=0"
    If Not blnAvgMassDefined Then tsOutfile.WriteLine "AvgMassWeight=0"
    If Not blnLogAbundanceDefined Then tsOutfile.WriteLine "LogAbundanceWeight=0"
    If Not blnScanDefined Then tsOutfile.WriteLine "ScanWeight=0"
    If Not blnNETDefined Then tsOutfile.WriteLine "NETWeight=0"
    If Not blnFitDefined Then tsOutfile.WriteLine "FitWeight=0"

    ' Write out some additional settings
    tsOutfile.WriteLine "MaxDistance=" & Trim(MyDef.TooDistant)
    
    strLineOut = "UseGenericNET="
    If blnUseGenericNET Then
        strLineOut = strLineOut & "True"
    Else
        strLineOut = strLineOut & "False"
    End If
    tsOutfile.WriteLine strLineOut
    
    GetScanRange CallerID, lngGelScanNumberMin, lngGelScanNumberMax, 0, 0
    tsOutfile.WriteLine "MinScan=" & Trim(lngGelScanNumberMin)
    tsOutfile.WriteLine "MaxScan=" & Trim(lngGelScanNumberMax)
    
    If UMCMakeSingleMemberClasses Then
        intMinLength = 1
    Else
        intMinLength = 2
    End If
    strLineOut = "MinFeatureLengthPoints=" & Trim(intMinLength)
    tsOutfile.WriteLine strLineOut

    tsOutfile.Close
    Set tsOutfile = Nothing
    
    ExportPeaksForUMCFinding = True
    Exit Function

ExportPeaksForUMCFindingErrorHandler:
Debug.Print "Error in ExportPeaksForUMCFinding: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->ExportPeaksForUMCFinding"

ExportPeaksForUMCFinding = False

End Function

Private Sub SetMolecularMassFieldDropdown(eMWField As mftMassFieldTypeConstants)
    If eMWField <> mftMWAvg And eMWField <> mftMWMono And eMWField <> mftMWTMA Then
        eMWField = GelData(CallerID).Preferences.IsoDataField
    End If
    
    Select Case eMWField
    Case mftMWAvg
        cboMolecularMassField.ListIndex = 0
    Case mftMWMono
        cboMolecularMassField.ListIndex = 1
    Case mftMWTMA
        cboMolecularMassField.ListIndex = 2
    Case Else
        cboMolecularMassField.ListIndex = 1
    End Select
End Sub

Private Function GetMolecularMassFieldFromDropdown() As Integer
    Dim eMWField As mftMassFieldTypeConstants
    
    Select Case cboMolecularMassField.ListIndex
    Case 0
        eMWField = mftMWAvg
    Case 2
        eMWField = mftMWTMA
    Case Else
        ' Includes case 1
        eMWField = mftMWMono
    End Select

    GetMolecularMassFieldFromDropdown = eMWField
End Function

Private Function FinalizeNewUMCs() As Boolean
    
    Dim dblTolPPM As Double
    Dim eTolType As glMassToleranceConstants
    
    Dim blnUMCIndicesUpdated As Boolean
    Dim blnSuccess As Boolean
    
On Error GoTo FinalizeNewUMCsErrorHandler
    blnSuccess = False
    
    ChangeStatus " Managing LC-MS Feature structures..."
    If ManageClasses(CallerID, UMCManageConstants.UMCMngTrim) Then
        
        ' Examine GelUMCIon(CallerID).ThisNetDef to determine the appropriate .Tol and .TolType
        '  to record in GelUMC(Callerid).Def
        LookupUMCIonNetMassTolerances dblTolPPM, eTolType, GelUMCIon(CallerID).ThisNetDef, UMC_IONNET_PPM_CONVERSION_MASS
        
        'set various Unique Mass Classes parameters
        With GelUMC(CallerID).def
            .UMCType = glUMC_TYPE_FROM_NET
            .MWField = GetMolecularMassFieldFromDropdown
            .UMCSharing = False
            .Tol = dblTolPPM            ' IonNet searching doesn't really use ppm, but we'll store ppm here anyway so that it gets exported to the database
            .TolType = eTolType
        End With
        
        ' Make sure UMCDef and GelUMC(CallerID).def are synchronized
        UMCDef = GelUMC(CallerID).def
        
        ' Store UMCDef in GelSearchDef() so that it gets saved to disk
        GelSearchDef(CallerID).UMCDef = UMCDef
        GelSearchDef(CallerID).UMCIonNetDef = GelUMCIon(CallerID).ThisNetDef
        
        glbPreferencesExpanded.UMCIonNetOptions.UMCRepresentative = UMCRepresentative
        GelUMCDraw(CallerID).DrawType = cmbUMCDrawType.ListIndex
        
        AddToAnalysisHistory CallerID, ConstructUMCDefDescription(CallerID, AUTO_ANALYSIS_UMCIonNet, UMCDef, glbPreferencesExpanded.UMCAdvancedStatsOptions, False, True)
        
        ChangeStatus "Calculating LC-MS Feature parameters..."
        
        ' Possibly Auto-Refine the LC-MS Features
        blnUMCIndicesUpdated = AutoRefineUMCs(CallerID, Me)
        
        If Not blnUMCIndicesUpdated Then
            ' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
            blnSuccess = UpdateUMCStatArrays(CallerID, False, Me)
        Else
            blnSuccess = True
        End If
      
        If glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCsByAbundance Then
            ' Be sure to call UpdateUMCStatArrays before using clsSplitUMCsByAbundance
            
            Set mSplitUMCs = New clsSplitUMCsByAbundance
            mSplitUMCs.ExamineUMCs CallerID, Me, GelUMC(CallerID).def.OddEvenProcessingMode, False, True
            
            Set mSplitUMCs = Nothing
        End If
        
    Else
       ChangeStatus " Error managing LC-MS Feature structures."
    End If
    
    FinalizeNewUMCs = blnSuccess
    
    Exit Function
      
FinalizeNewUMCsErrorHandler:
    Debug.Print "Error in FinalizeNewUMCs: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->FinalizeNewUMCs"
    
    FinalizeNewUMCs = False
    
End Function

Private Sub FindBestMatches(ByVal eOddEvenProcessingMode As oepUMCOddEvenProcessingMode)
    Dim i As Long, j As Long                'loop controlers
    Dim iOInd As Long, jOInd As Long        'indexes in original Data arrays
    Dim BestForI As Long                    'index of best match for index i
    Dim ShortestDistance As Double, CurrDistance As Double
    Dim bTooFarAway As Boolean
    Dim lngTickCountLastUpdate As Long, lngNewTickCount As Long
    Dim dtLastUpdateTime As Date

    Dim intOddEvenIteration As Integer
    Dim intOddEvenIterationStart As Integer
    Dim intOddEvenIterationEnd As Integer
    Dim blnComputeDistance As Boolean
    Dim strScanNumMode As String
    
    On Error GoTo err_FindBestMatches
    mAbortProcess = False
    
    Select Case eOddEvenProcessingMode
    Case oepUMCOddEvenProcessingMode.oepOddOnly
        intOddEvenIterationStart = 1
        intOddEvenIterationEnd = 1
    Case oepUMCOddEvenProcessingMode.oepEvenOnly
        intOddEvenIterationStart = 2
        intOddEvenIterationEnd = 2
    Case oepUMCOddEvenProcessingMode.oepOddEvenSequential
        intOddEvenIterationStart = 1
        intOddEvenIterationEnd = 2
    Case Else
        ' Includes case oepUMCOddEvenProcessingMode.oepProcessAll
        intOddEvenIterationStart = 0
        intOddEvenIterationEnd = 0
    End Select
    
    For intOddEvenIteration = intOddEvenIterationStart To intOddEvenIterationEnd
      Select Case intOddEvenIteration
      Case 0: strScanNumMode = ""
      Case 1: strScanNumMode = ", odd numbered spectra"
      Case 2: strScanNumMode = ", even numbered spectra"
      Case Else: strScanNumMode = ", Unknown spectrum numbering mode"
      End Select
      
      Select Case MyDef.MetricType
          Case METRIC_EUCLIDEAN
              Select Case MyDef.NETType
                  Case Net_SPIDER_66                                'remember all connections shorter than threshold
                      For i = 0 To DataCnt - 1
                          iOInd = OptIndO(i)
    
                         ' Only compute the distance if intOddEvenIteration = 0 or if the scan number for
                         ' the data point is the appropriate odd or even value
                          blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, iOInd)
                          If blnComputeDistance Then
                              lngNewTickCount = GetTickCount()     ' Note that GetTickCount returns a negative number after 24 days of computer Uptime and resets to 0 after 48 days
                              If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                                  ' Only update 4 times per second
                                  ChangeStatus ("Calculating line " & i & " / " & Trim(DataCnt) & strScanNumMode)
                                  If mAbortProcess Then Exit For
                                  lngTickCountLastUpdate = lngNewTickCount
                                  dtLastUpdateTime = Now()
                              End If
                              j = i + 1
                              bTooFarAway = (j > DataCnt - 1)
                              Do Until bTooFarAway
                                  If MetricEuclidDim1(i, j) > MyDef.TooDistant Then
                                      bTooFarAway = True
                                  Else
                                      jOInd = OptIndO(j)
    
                                      blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                                      If blnComputeDistance Then
                                          CurrDistance = MetricEuclid(iOInd, jOInd)
                                          If CurrDistance < MyDef.TooDistant Then
                                              If Not SubjectToConstraintEuclid(iOInd, jOInd) Then
                                                  ResCnt = ResCnt + 1
                                                  'put in results original indexes; always smaller index first
                                                  If iOInd < jOInd Then
                                                      ResInd1(ResCnt - 1) = iOInd: ResInd2(ResCnt - 1) = jOInd
                                                  Else
                                                      ResInd1(ResCnt - 1) = jOInd: ResInd2(ResCnt - 1) = iOInd
                                                  End If
                                                  ResDist(ResCnt - 1) = CurrDistance
                                              End If
                                          End If
                                      End If
                                  End If
                                  j = j + 1
                                  If j > DataCnt - 1 Then bTooFarAway = True
                              Loop
                          End If
                      Next i
                  Case Else
                      For i = 0 To DataCnt - 1
                          iOInd = OptIndO(i)
                          BestForI = -1
                          ShortestDistance = glHugeDouble
                          If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                              ' Only update 4 times per second
                              ChangeStatus ("Calculating line " & i & " / " & Trim(DataCnt) & strScanNumMode)
                              If mAbortProcess Then Exit For
                              lngTickCountLastUpdate = lngNewTickCount
                              dtLastUpdateTime = Now()
                          End If
                          If mAbortProcess Then Exit For
                          j = i + 1
                          bTooFarAway = (j > DataCnt - 1)
                          Do Until bTooFarAway
                              If MetricEuclidDim1(i, j) > MyDef.TooDistant Then
                                  bTooFarAway = True
                              Else
                                  jOInd = OptIndO(j)
                                  
                                  blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                                  If blnComputeDistance Then
                                      CurrDistance = MetricEuclid(iOInd, jOInd)
                                      If CurrDistance < ShortestDistance Then
                                          BestForI = j
                                          ShortestDistance = CurrDistance
                                      End If
                                  End If
                              End If
                              j = j + 1
                              If j > DataCnt - 1 Then bTooFarAway = True
                          Loop
                          If ShortestDistance < MyDef.TooDistant Then
                              blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                              If blnComputeDistance Then
                                  If Not SubjectToConstraintEuclid(iOInd, jOInd) Then
                                      ResCnt = ResCnt + 1
                                      'put in results original indexes; always smaller first
                                      If iOInd < OptIndO(BestForI) Then
                                          ResInd1(ResCnt - 1) = iOInd: ResInd2(ResCnt - 1) = OptIndO(BestForI)
                                      Else
                                          ResInd1(ResCnt - 1) = OptIndO(BestForI): ResInd2(ResCnt - 1) = iOInd
                                      End If
                                      ResDist(ResCnt - 1) = ShortestDistance
                                  End If
                              End If
                          End If
                      Next i
              End Select
          Case METRIC_HONDURAS
              Select Case MyDef.NETType
                  Case Net_SPIDER_66                                'remember all connections shorter than threshold
                      For i = 0 To DataCnt - 1
                          iOInd = OptIndO(i)
                          If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                              ' Only update 4 times per second
                              ChangeStatus ("Calculating line " & i & " / " & Trim(DataCnt) & strScanNumMode)
                              If mAbortProcess Then Exit For
                              lngTickCountLastUpdate = lngNewTickCount
                              dtLastUpdateTime = Now()
                          End If
                          If mAbortProcess Then Exit For
                          j = i + 1
                          bTooFarAway = (j > DataCnt - 1)
                          Do Until bTooFarAway
                              If MetricHondurasDim1(i, j) > MyDef.TooDistant Then
                                  bTooFarAway = True
                              Else
                                  jOInd = OptIndO(j)
    
                                ' Only compute the distance if intOddEvenIteration = 0 or if the scan number for
                                ' the data point is the appropriate odd or even value
                                  blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                                  If blnComputeDistance Then
                                      CurrDistance = MetricHonduras(iOInd, jOInd)
                                      If CurrDistance < MyDef.TooDistant Then
                                          If Not SubjectToConstraintHonduras(iOInd, jOInd) Then
                                              ResCnt = ResCnt + 1
                                              'put in results original indexes; always smaller index first
                                              If iOInd < jOInd Then
                                                  ResInd1(ResCnt - 1) = iOInd: ResInd2(ResCnt - 1) = jOInd
                                              Else
                                                  ResInd1(ResCnt - 1) = jOInd: ResInd2(ResCnt - 1) = iOInd
                                              End If
                                              ResDist(ResCnt - 1) = CurrDistance
                                          End If
                                      End If
                                  End If
                              End If
                              j = j + 1
                              If j > DataCnt - 1 Then bTooFarAway = True
                          Loop
                      Next i
                  Case Else
                      For i = 0 To DataCnt - 1
                          iOInd = OptIndO(i)
                          BestForI = -1
                          ShortestDistance = glHugeDouble
                          If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                              ' Only update 4 times per second
                              ChangeStatus ("Calculating line " & i & " / " & Trim(DataCnt) & strScanNumMode)
                              If mAbortProcess Then Exit For
                              lngTickCountLastUpdate = lngNewTickCount
                              dtLastUpdateTime = Now()
                          End If
                          If mAbortProcess Then Exit For
                          j = i + 1
                          bTooFarAway = (j > DataCnt - 1)
                          Do Until bTooFarAway
                              If MetricHondurasDim1(i, j) > MyDef.TooDistant Then
                                  bTooFarAway = True
                              Else
                                  jOInd = OptIndO(j)
    
                                  blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                                  If blnComputeDistance Then
                                      CurrDistance = MetricHonduras(iOInd, jOInd)
                                      If CurrDistance < ShortestDistance Then
                                          BestForI = j
                                          ShortestDistance = CurrDistance
                                      End If
                                  End If
                              End If
                              j = j + 1
                              If j > DataCnt - 1 Then bTooFarAway = True
                          Loop
                          If ShortestDistance < MyDef.TooDistant Then
                              blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                              If blnComputeDistance Then
    
                                  If Not SubjectToConstraintHonduras(iOInd, jOInd) Then
                                      ResCnt = ResCnt + 1
                                      'put in results original indexes; always smaller first
                                      If iOInd < OptIndO(BestForI) Then
                                          ResInd1(ResCnt - 1) = iOInd: ResInd2(ResCnt - 1) = OptIndO(BestForI)
                                      Else
                                          ResInd1(ResCnt - 1) = OptIndO(BestForI): ResInd2(ResCnt - 1) = iOInd
                                      End If
                                      ResDist(ResCnt - 1) = ShortestDistance
                                  End If
                              End If
                          End If
                      Next i
              End Select
          Case METRIC_INFINITY
              Select Case MyDef.NETType
                  Case Net_SPIDER_66                                'remember all connections shorter than threshold
                      For i = 0 To DataCnt - 1
                          iOInd = OptIndO(i)
                          If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                              ' Only update 4 times per second
                              ChangeStatus ("Calculating line " & i & " / " & Trim(DataCnt) & strScanNumMode)
                              If mAbortProcess Then Exit For
                              lngTickCountLastUpdate = lngNewTickCount
                              dtLastUpdateTime = Now()
                          End If
                          If mAbortProcess Then Exit For
                          j = i + 1
                          bTooFarAway = (j > DataCnt - 1)
                          Do Until bTooFarAway
                              If MetricInfinityDim1(i, j) > MyDef.TooDistant Then
                                  bTooFarAway = True
                              Else
                                  jOInd = OptIndO(j)
    
                                ' Only compute the distance if intOddEvenIteration = 0 or if the scan number for
                                ' the data point is the appropriate odd or even value
                                  blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                                  If blnComputeDistance Then
                                      CurrDistance = MetricInfinity(iOInd, jOInd)
                                      If CurrDistance < MyDef.TooDistant Then
                                          If Not SubjectToConstraintInfinity(iOInd, jOInd) Then
                                              ResCnt = ResCnt + 1
                                              'put in results original indexes; always smaller index first
                                              If iOInd < jOInd Then
                                                  ResInd1(ResCnt - 1) = iOInd: ResInd2(ResCnt - 1) = jOInd
                                              Else
                                                  ResInd1(ResCnt - 1) = jOInd: ResInd2(ResCnt - 1) = iOInd
                                              End If
                                              ResDist(ResCnt - 1) = CurrDistance
                                          End If
                                      End If
                                  End If
                              End If
                              j = j + 1
                              If j > DataCnt - 1 Then bTooFarAway = True
                          Loop
                      Next i
                  Case Else
                      For i = 0 To DataCnt - 1
                          iOInd = OptIndO(i)
                          BestForI = -1
                          ShortestDistance = glHugeDouble
                          If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                              ' Only update 4 times per second
                              ChangeStatus ("Calculating line " & i & " / " & Trim(DataCnt) & strScanNumMode)
                              If mAbortProcess Then Exit For
                              lngTickCountLastUpdate = lngNewTickCount
                              dtLastUpdateTime = Now()
                          End If
                          If mAbortProcess Then Exit For
                          j = i + 1
                          bTooFarAway = (j > DataCnt - 1)
                          Do Until bTooFarAway
                              If MetricInfinityDim1(i, j) > MyDef.TooDistant Then
                                  bTooFarAway = True
                              Else
                                  jOInd = OptIndO(j)
    
                                ' Only compute the distance if intOddEvenIteration = 0 or if the scan number for
                                ' the data point is the appropriate odd or even value
                                  blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                                  If blnComputeDistance Then
                                      CurrDistance = MetricInfinity(iOInd, jOInd)
                                      If CurrDistance < ShortestDistance Then
                                          BestForI = j
                                          ShortestDistance = CurrDistance
                                      End If
                                  End If
                              End If
                              j = j + 1
                              If j > DataCnt - 1 Then bTooFarAway = True
                          Loop
                          If ShortestDistance < MyDef.TooDistant Then
                              blnComputeDistance = CheckOddEvenIterationForDataPoint(intOddEvenIteration, jOInd)
                              If blnComputeDistance Then
    
                                  If Not SubjectToConstraintInfinity(iOInd, jOInd) Then
                                      ResCnt = ResCnt + 1
                                      'put in results original indexes; always smaller first
                                      If iOInd < OptIndO(BestForI) Then
                                          ResInd1(ResCnt - 1) = iOInd: ResInd2(ResCnt - 1) = OptIndO(BestForI)
                                      Else
                                          ResInd1(ResCnt - 1) = OptIndO(BestForI): ResInd2(ResCnt - 1) = iOInd
                                      End If
                                      ResDist(ResCnt - 1) = ShortestDistance
                                  End If
                              End If
                          End If
                      Next i
              End Select
      End Select
    
    Next intOddEvenIteration
    

    Call ManageResArrays(amtTrim)
    Exit Sub


err_FindBestMatches:
Select Case Err.Number
Case 9                  'add more space for results and continue
     If ManageResArrays(amtAdd) Then
        Resume
     Else
        Call ManageResArrays(amtTrim)
     End If
Case Else
     LogErrors Err.Number, "frmUMCIonNet.FindBestMatches"
End Select
End Sub


Private Function FindUMCsUsingLCMSFeatureFinder(ByVal blnShowMessages As Boolean) As Boolean
    ' Find LC-MS Features using LCMSFeatureFinder.exe
    ' If the .Exe isn't found or if a problem occurs while finding LC-MS Features, then this function will return False
    
    Const DEFAULT_MAXIMUM_PROCESSING_TIME_MINUTES As Single = 60
    
    Const APP_MONITOR_INTERVAL_MSEC As Integer = 100
    Const STATUS_UPDATE_INTERVAL_MSEC As Integer = 500
    
    Dim strWorkingDirPath As String
    Dim strFeatureFinderAppPath As String
    Dim strArguments As String
    
    Dim strLCMSFeaturesFilePath As String
    Dim strIniFilePath As String
    
    Dim strMessage As String
    Dim strStatusBase As String
    Dim strStatusSpectrumType As String
    
    Dim fso As New FileSystemObject
    Dim objProgRunner As clsProgRunner
    
    Dim sngProcessingTimeSeconds As Single
    Dim sngMaxProcessingTimeMinutes As Single
    Dim dtProcessingStartTime As Date
    
    Dim lngIteration As Long
    Dim lngStatusUpdateIterationCount As Long
        
    Dim eOddEvenProcessingMode As oepUMCOddEvenProcessingMode
    Dim intOddEvenIteration As Integer
    Dim intOddEvenIterationStart As Integer
    Dim intOddEvenIterationEnd As Integer
    Dim intIterationSuccessCount As Integer
        
    Dim blnAbortProcessing As Boolean
    Dim blnSuccess As Boolean
    Dim blnSuccessCurrentIteration As Boolean

On Error GoTo FindUMCsUsingLCMSFeatureFinderErrorHandler
    
    sngMaxProcessingTimeMinutes = DEFAULT_MAXIMUM_PROCESSING_TIME_MINUTES
    If sngMaxProcessingTimeMinutes < 1 Then sngMaxProcessingTimeMinutes = 1
    If sngMaxProcessingTimeMinutes > 300 Then sngMaxProcessingTimeMinutes = 300
    
    blnSuccess = True
    blnSuccessCurrentIteration = False
    blnAbortProcessing = False
    intIterationSuccessCount = 0
    
    ' Check for the existence of LCMSFeatureFinder.exe
    strFeatureFinderAppPath = fso.BuildPath(App.Path, LCMS_FEATURE_FINDER_APP_NAME)
    strWorkingDirPath = App.Path
        
    If Not fso.FileExists(strFeatureFinderAppPath) Then
        strMessage = "LCMS Feature Finder app not found, unable to continue: " & vbCrLf & strFeatureFinderAppPath
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            AddToAnalysisHistory CallerID, strMessage
        ElseIf blnShowMessages Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "File Not Found"
        End If
    Else
        mAbortProcess = False
        mCalculating = True
        ShowHideCommandButtons mCalculating
        
        InitializeLCMSFeatureInfo
    
        eOddEvenProcessingMode = UMCDef.OddEvenProcessingMode
        Select Case UMCDef.OddEvenProcessingMode
        Case oepUMCOddEvenProcessingMode.oepOddOnly
            intOddEvenIterationStart = 1
            intOddEvenIterationEnd = 1
        Case oepUMCOddEvenProcessingMode.oepEvenOnly
            intOddEvenIterationStart = 2
            intOddEvenIterationEnd = 2
        Case oepUMCOddEvenProcessingMode.oepOddEvenSequential
            intOddEvenIterationStart = 1
            intOddEvenIterationEnd = 2
        Case Else
            ' Includes case oepUMCOddEvenProcessingMode.oepProcessAll
            intOddEvenIterationStart = 0
            intOddEvenIterationEnd = 0
        End Select
        
        For intOddEvenIteration = intOddEvenIterationStart To intOddEvenIterationEnd
        
            ' Create two text files for LCMSFeatureFinder.exe to read
            blnSuccessCurrentIteration = ExportPeaksForUMCFinding(strWorkingDirPath, strLCMSFeaturesFilePath, strIniFilePath, intOddEvenIteration)
                    
            If mAbortProcess Then
                GoTo FindUMCsUsingLCMSFeatureFinderCleanup
            End If
            
            If blnSuccessCurrentIteration Then
                strStatusBase = "Calling " & LCMS_FEATURE_FINDER_APP_NAME & " to find the LC-MS Features"
                ChangeStatus strStatusBase
                        
                strArguments = strLCMSFeaturesFilePath
                If InStr(strArguments, " ") > 0 Then
                    strArguments = """" & strArguments & """"
                End If
                strArguments = "/I:" & strArguments
        
                Set objProgRunner = New clsProgRunner
                dtProcessingStartTime = Now()
                
                If objProgRunner.StartProgram(strFeatureFinderAppPath, strArguments, vbNormalNoFocus) Then
                
                    lngStatusUpdateIterationCount = CInt(STATUS_UPDATE_INTERVAL_MSEC / CSng(APP_MONITOR_INTERVAL_MSEC))
                    If lngStatusUpdateIterationCount < 1 Then lngStatusUpdateIterationCount = 1
                    
                    Do While objProgRunner.AppRunning
                        Sleep APP_MONITOR_INTERVAL_MSEC
                        
                        sngProcessingTimeSeconds = (Now - dtProcessingStartTime) * 86400#
                        If sngProcessingTimeSeconds / 60# >= sngMaxProcessingTimeMinutes Then
                            blnAbortProcessing = True
                            strMessage = "LC-MS Feature Finding using the LCMS Feature Finder was aborted because over " & Trim(sngMaxProcessingTimeMinutes) & " minutes has elapsed."
                        ElseIf mAbortProcess Then
                            blnAbortProcessing = True
                            strMessage = "LC-MS Feature Finding using the LCMS Feature Finder was manually aborted by the user after " & Trim(sngProcessingTimeSeconds) & " seconds of processing."
                        End If
                        
                        If blnAbortProcessing Then
                            objProgRunner.AbortProcessing
                            DoEvents
                            
                            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                               MsgBox strMessage, vbOKOnly, glFGTU
                            Else
                               Debug.Assert False
                               LogErrors Err.Number, "frmUMCIonNet->FindUMCsUsingLCMSFeatureFinder", strMessage
                               AddToAnalysisHistory CallerID, strMessage
                            End If
                            
                            ChangeStatus strMessage
                            Exit Do
                        End If
                        
                        If lngIteration Mod lngStatusUpdateIterationCount = 0 Then
                            If intOddEvenIteration = 1 Then
                                strStatusSpectrumType = "; odd-numbered spectra"
                            ElseIf intOddEvenIteration = 2 Then
                                strStatusSpectrumType = "; even-numbered spectra"
                            End If
                            
                            ChangeStatus strStatusBase & strStatusSpectrumType & ": " & Round(sngProcessingTimeSeconds, 1) & " seconds elapsed"
                        End If
                        DoEvents
                        
                        lngIteration = lngIteration + 1
                    Loop
        
                    blnSuccess = Not blnAbortProcessing
                    
                    If blnSuccess Then
                        ' Read the data from the _Features.txt & _PeakToFeatureMap.txt files
                        
                        blnSuccessCurrentIteration = LoadFeatureInfoFromDisk(fso, strWorkingDirPath, strLCMSFeaturesFilePath, blnShowMessages)
                        If blnSuccessCurrentIteration Then
                            intIterationSuccessCount = intIterationSuccessCount + 1
                        End If
                        If mAbortProcess Then blnSuccess = False
                    End If
                End If
            End If
            
            If Not blnSuccess Then Exit For
        
        Next intOddEvenIteration
        
        If intIterationSuccessCount > 0 Then
            blnSuccess = BuildUMCsUsingmLCMSResultsMapping(blnShowMessages)
        Else
            blnSuccess = False
        End If

    End If

FindUMCsUsingLCMSFeatureFinderCleanup:
    mCalculating = False
    ShowHideCommandButtons mCalculating

    FindUMCsUsingLCMSFeatureFinder = blnSuccess
    
    Exit Function

FindUMCsUsingLCMSFeatureFinderErrorHandler:
    Debug.Print "Error in FindUMCsUsingLCMSFeatureFinder: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->FindUMCsUsingLCMSFeatureFinder"

Resume FindUMCsUsingLCMSFeatureFinderCleanup
    blnSuccess = False
    
End Function

Private Sub FindIonNetConnections()
    Dim eResponse As VbMsgBoxResult
    Dim eOddEvenProcessingMode As oepUMCOddEvenProcessingMode
    
    Dim lngConnectionsEliminated As Long
    Dim strUMCIsoDefinition As String
    
    On Error GoTo FindIonNetConnectionsErrorHandler
    
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If MyDef.NETType <= 0 Then
           MsgBox "Net type should be positive integer.", vbOKOnly, glFGTU
           txtNETType.SetFocus
           Exit Sub
        End If
        If GelUMCIon(CallerID).NetCount > 0 Then
           eResponse = MsgBox("Isotopic NET already established. Overwrite?", vbYesNo, glFGTU)
           If eResponse = vbNo Then Exit Sub
        End If
    End If
    
    mAbortProcess = False
    mCalculating = True
    ShowHideCommandButtons mCalculating
    
    GelUMC(CallerID).def.OddEvenProcessingMode = UMCDef.OddEvenProcessingMode
    eOddEvenProcessingMode = UMCDef.OddEvenProcessingMode
    
    If PrepareDataArrays() Then
       If PrepareOptimization() Then
          If ManageResArrays(amtInitialize) Then
             FindBestMatches eOddEvenProcessingMode
             
             lngConnectionsEliminated = EliminateLongConnections(MyDef.TooDistant)
             
             'NOTE: EliminateRedundantConnectionsDirect is time consuming procedure and for now
             'all connection creation procedures create non-redundant connections; it might be
             'neccessary to uncomment it if future code creates redundant connections
             'Call EliminateRedundantConnectionsDirect
             Call CreateNet
             
             'copy current settings to caller structures
             GelUMCIon(CallerID).ThisNetDef = MyDef
             GelSearchDef(CallerID).UMCIonNetDef = MyDef
          
             strUMCIsoDefinition = GetUMCIsoDefinitionText(CallerID, False)
             strUMCIsoDefinition = Replace(strUMCIsoDefinition, ": ", " = ")
             strUMCIsoDefinition = Trim(Replace(strUMCIsoDefinition, vbCrLf, ""))
             If Right(strUMCIsoDefinition, 1) = ";" Then
                strUMCIsoDefinition = Left(strUMCIsoDefinition, Len(strUMCIsoDefinition) - 1)
             End If
             
             AddToAnalysisHistory CallerID, "Found data-point connections (" & AUTO_ANALYSIS_UMCIonNet & "); Connection count = " & Trim(GelUMCIon(CallerID).NetCount) & "; " & strUMCIsoDefinition & "; Connections eliminated by max distance filter = " & Trim(lngConnectionsEliminated)
          Else
             ChangeStatus " Error initializing Net structures."
          End If
       End If
    End If
    
    mCalculating = False
    ShowHideCommandButtons mCalculating
    
    Exit Sub
    
FindIonNetConnectionsErrorHandler:
    Debug.Print "Error in FindIonNetConnections: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->FindIonNetConnections"
    Resume Next
    
End Sub

Private Function FormClassesFromNETsWrapper(Optional ByVal blnShowMessages As Boolean = True) As Boolean
'---------------------------------------------------------------------------
'controls creation of Unique Mass Classes from Net
'calls FormClassesFromNETs if not errors
'Returns True if success; False if failure
'---------------------------------------------------------------------------
    
    Dim eResponse As VbMsgBoxResult
    Dim blnSuccess As Boolean
    Dim strMessage As String
    
    On Error GoTo FormClassesFromNETsWrapperErrorHandler
    
    If mCalculating Then Exit Function
    
    If GelUMCIon(CallerID).NetCount > 0 Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled And blnShowMessages Then
            If GelUMC(CallerID).UMCCnt > 0 Then
                eResponse = MsgBox("Unique Mass Classes structure already exists. Overwrite?", vbYesNo, glFGTU)
                If eResponse <> vbYes Then Exit Function
            End If
            
            If Not ValidateClassStatOptions Then
                Exit Function
            End If
        End If
        
        ' Call FormClassesFromNets
        blnSuccess = FormClassesFromNets()
        
        If blnSuccess Then
           ChangeStatus "Number of LC-MS Features: " & GelUMC(CallerID).UMCCnt
        Else
           ChangeStatus "Error creating LC-MS Features from Nets."
        End If
    Else
        strMessage = "Net elements not found.  Unable to find LC-MS Features."
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
           MsgBox strMessage, vbOKOnly, glFGTU
        Else
           Debug.Assert False
           LogErrors Err.Number, "frmUMCIonNet->FormClassesFromNETsWrapper, .NetCount = 0"
           AddToAnalysisHistory CallerID, "Error in UMCIonNet Searching: " & strMessage
        End If
        blnSuccess = False
    End If
    
    If GelUMCDraw(CallerID).Visible Then
        GelBody(CallerID).RequestRefreshPlot
        GelBody(CallerID).csMyCooSys.CoordinateDraw
    End If
    
    FormClassesFromNETsWrapper = blnSuccess

Exit Function

FormClassesFromNETsWrapperErrorHandler:
Debug.Print "Error occurred in FormClassesFromNETsWrapper: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->FormClassesFromNETsWrapper"
Resume Next

End Function

Private Function FormClassesFromNets() As Boolean
'--------------------------------------------------------------------
'creates Unique Mass Classes from Net structure of current 2D display
'returns True if successful;
'NOTE: if this function is called we have at least one connection in
'GelUMCIon structure
'--------------------------------------------------------------------
Dim bDone As Long
Dim CurrConnInd As Long
Dim CurrInd1 As Long, CurrInd2 As Long
Dim intScopeUsedForConnections As Integer
Dim i As Long
Dim lngTickCountLastUpdate As Long, lngNewTickCount As Long
Dim dtLastUpdateTime As Date

Dim blnSuccess As Boolean
On Error GoTo err_FormClassesFromNets

mAbortProcess = False
mCalculating = True
ShowHideCommandButtons mCalculating

' Update GelUMC(CallerID).def now, prior to processing
With GelUMC(CallerID)
    ' Save the Scope used when finding the connections since
    '  the user may have changed it since then, thus affecting UMCDef
    intScopeUsedForConnections = .def.DefScope
    
    ' Copy the settings from UMCDef to GelUMC(CallerID)
    .def = UMCDef
    
    ' Make sure the scope in GelUMC() is correct
    .def.DefScope = intScopeUsedForConnections
End With

ChangeStatus " Initializing LC-MS Feature structures..."
If ManageClasses(CallerID, UMCManageConstants.UMCMngInitialize) Then
   ChangeStatus " Preparing classification..."
   If PrepareHUMCArrays() Then
      With GelUMCIon(CallerID)
         CurrConnInd = 0
         Do Until bDone
            If HUMCNetUsed(CurrConnInd) = HUMCUsed Then         'already used; go next
               CurrConnInd = CurrConnInd + 1
               If CurrConnInd > HUMCNetCnt - 1 Then bDone = True
            Else                                       'new class; find the whole class
               CurrInd1 = .NetInd1(CurrConnInd):    CurrInd2 = .NetInd2(CurrConnInd)
               HUMCNetUsed(CurrConnInd) = HUMCUsed
               HUMCIsoUsed(CurrInd1) = HUMCInUse:   HUMCIsoUsed(CurrInd2) = HUMCInUse
               'first index  < last index (if this changes this function has to be revised)
               HUMCEquClsWk(0) = CurrInd1:          HUMCEquClsWk(1) = CurrInd2
               HUMCEquClsCnt = 2                    'always start this type of classes with 2 points
               'build class; we have to go in both direction to discover full connection
               For i = CurrConnInd + 1 To HUMCNetCnt - 1
                   If HUMCNetUsed(i) = HUMCNotUsed Then
                      CurrInd1 = .NetInd1(i):    CurrInd2 = .NetInd2(i)
                      'condition in the following two If statements will not be True
                      'simultaneously but this way it will work even if they are
                      If HUMCIsoUsed(CurrInd1) = HUMCInUse Then
                         HUMCNetUsed(i) = HUMCUsed
                         If HUMCIsoUsed(CurrInd2) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(CurrInd2) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = CurrInd2
                         End If
                      End If
                      If HUMCIsoUsed(CurrInd2) = HUMCInUse Then
                         HUMCNetUsed(i) = HUMCUsed
                         If HUMCIsoUsed(CurrInd1) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(CurrInd1) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = CurrInd1
                         End If
                      End If
                   End If
               Next i
               'need to go in another direction to pick up eventual skiping transitions
               For i = HUMCNetCnt - 1 To CurrConnInd + 1 Step -1
                   If HUMCNetUsed(i) = HUMCNotUsed Then
                      CurrInd1 = .NetInd1(i):    CurrInd2 = .NetInd2(i)
                      If HUMCIsoUsed(CurrInd1) = HUMCInUse Then
                         HUMCNetUsed(i) = HUMCUsed
                         If HUMCIsoUsed(CurrInd2) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(CurrInd2) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = CurrInd2
                         End If
                      End If
                      If HUMCIsoUsed(CurrInd2) = HUMCInUse Then
                         HUMCNetUsed(i) = HUMCUsed
                         If HUMCIsoUsed(CurrInd1) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(CurrInd1) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = CurrInd1
                         End If
                      End If
                   End If
               Next i
               'now pack findings to nice small array convenient to create classes
               ReDim HUMCEquCls(HUMCEquClsCnt - 1)
               For i = 0 To HUMCEquClsCnt - 1       'make sure not to use more than belongs to this class
                   HUMCEquCls(i) = HUMCEquClsWk(i)
                   HUMCIsoUsed(HUMCEquCls(i)) = HUMCUsed                'they are used now
               Next i
               'extract and add class to the structure
               Call BuildCurrentClass
               CurrConnInd = CurrConnInd + 1
               
               lngNewTickCount = GetTickCount()     ' Note that GetTickCount returns a negative number after 24 days of computer Uptime and resets to 0 after 48 days
               If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                   ' Only update 4 times per second
                   ChangeStatus "Building LC-MS Feature: " & GelUMC(CallerID).UMCCnt & " (" & Format(CurrConnInd / HUMCNetCnt * 100, "0.00") & "% completed)"
                   lngTickCountLastUpdate = lngNewTickCount
                   dtLastUpdateTime = Now()
                   If mAbortProcess Then bDone = True
               End If
               
               If CurrConnInd > HUMCNetCnt - 1 Then bDone = True
            End If
         Loop
      End With
      
      ' Add single member classes if requested
      If UMCMakeSingleMemberClasses Then Call HUMCAddingSingleMemberUMCs
      
      ' Refine the LC-MS Features and compute class stats
      blnSuccess = FinalizeNewUMCs()
      
   End If
End If

FormClassesFromNetsCleanup:
mCalculating = False
ShowHideCommandButtons mCalculating

FormClassesFromNets = blnSuccess
Exit Function

err_FormClassesFromNets:
Debug.Print "Error in FormClassesFromNets: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->FormClassesFromNets"
blnSuccess = False
Resume FormClassesFromNetsCleanup

End Function

Private Function GetDataInScope(ByRef ISInd() As Long, ByRef DataCnt As Long) As Boolean
    Dim strMessage As String
    
    GelUMC(CallerID).def.DefScope = UMCDef.DefScope
    DataCnt = GetISScope(CallerID, ISInd(), UMCDef.DefScope)

    If DataCnt < 2 Then
       strMessage = "Insufficient number of isotopic data points (must have 2 or more)."
       If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
           MsgBox strMessage, vbOKOnly, glFGTU
       Else
           Debug.Assert False
           LogErrors Err.Number, "frmUMCIonNet->ExportPeaksForUMCFinding, DataCnt < 2"
           AddToAnalysisHistory CallerID, "Error in UMCIonNet Searching: " & strMessage
       End If
              
       GetDataInScope = False
       Exit Function
    End If
    
    GetDataInScope = True
    
End Function

Private Function GetMetricDataMassUnits(lngMetricDataUnits As Long) As String
    Select Case lngMetricDataUnits
    Case DATA_UNITS_MASS_DA
        GetMetricDataMassUnits = "Da"
    Case DATA_UNITS_MASS_PPM
        GetMetricDataMassUnits = "ppm"
    Case Else
        Debug.Assert False
        GetMetricDataMassUnits = "??"
    End Select
End Function

Private Function GetUMCIsoDefinitionText(Ind As Long, Optional ByVal blnMultipleLines As Boolean = True) As String
'-----------------------------------------------------------------------
'returns formatted strDesc of the IonNet for 2D display with index Ind
'-----------------------------------------------------------------------
Dim i As Long
Dim strLineSeparator As String
Dim strAddnlText As String

If blnMultipleLines Then
    strLineSeparator = vbCrLf
Else
    strLineSeparator = "; "
End If

On Error Resume Next
Dim strDesc As String
With GelUMCIon(Ind).ThisNetDef
     Select Case .MetricType
     Case METRIC_EUCLIDEAN
          strDesc = "Metric type: Euclidean" & strLineSeparator
     Case METRIC_HONDURAS
          strDesc = "Metric type: Honduras (a.k.a. Taxicab)" & strLineSeparator
     Case METRIC_INFINITY
          strDesc = "Metric type: Infinity" & strLineSeparator
     End Select
     strDesc = strDesc & "Net type: " & .NETType & strLineSeparator
     strDesc = strDesc & "Max distance: " & .TooDistant & strLineSeparator
     If .NetActualDim > 0 Then
        If blnMultipleLines Then
            strDesc = strDesc & "Metric dimensions description:" & strLineSeparator
        Else
            strDesc = strDesc & "Metric dimensions description; "
        End If
        For i = 0 To .NetDim - 1
            If Not blnMultipleLines Then
                strDesc = strDesc & "Dimension" & Trim(i + 1) & " = "
            End If
            
            If .MetricData(i).Use Then
               Select Case .MetricData(i).DataType
               Case uindUMCIonNetDimConstants.uindMonoMW
                    strAddnlText = "Monoisotopic mass; "
               Case uindUMCIonNetDimConstants.uindAvgMW
                    strAddnlText = "Average mass; "
               Case uindUMCIonNetDimConstants.uindTmaMW
                    strAddnlText = "The most abundant mass; "
               Case uindUMCIonNetDimConstants.uindScan
                    strAddnlText = "Scan; "
               Case uindUMCIonNetDimConstants.uindFit
                    strAddnlText = "Isotopic fit; "
               Case uindUMCIonNetDimConstants.uindMZ
                    strAddnlText = "m/z; "
               Case uindUMCIonNetDimConstants.uindGenericNET
                    strAddnlText = "Generic NET; "
               Case uindUMCIonNetDimConstants.uindChargeState
                    strAddnlText = "Charge state; "
               Case uindUMCIonNetDimConstants.uindLogAbundance
                    strAddnlText = "Log(Abundance); "
               End Select
               strDesc = strDesc & strAddnlText
               strDesc = strDesc & "Weight factor: " & .MetricData(i).WeightFactor & "; "
               strDesc = strDesc & "Constraint: "
               Select Case .MetricData(i).ConstraintType
               Case Net_CT_None
                    strAddnlText = "none"
               Case Net_CT_LT
                    strAddnlText = "Distance < " & .MetricData(i).ConstraintValue
               Case Net_CT_GT
                    strAddnlText = "Distance > " & .MetricData(i).ConstraintValue
               Case Net_CT_EQ
                    strAddnlText = "Distance equal to " & .MetricData(i).ConstraintValue
               End Select
               
               strDesc = strDesc & strAddnlText
               If .MetricData(i).ConstraintType <> Net_CT_None Then
                    Select Case .MetricData(i).DataType
                    Case uindUMCIonNetDimConstants.uindMonoMW, uindUMCIonNetDimConstants.uindAvgMW, uindUMCIonNetDimConstants.uindTmaMW
                        strDesc = strDesc & " " & GetMetricDataMassUnits(.MetricData(i).ConstraintUnits)
                    Case Else
                        ' Do not append the units
                    End Select
               End If
            Else
                strDesc = strDesc & "Unused"
            End If
            strDesc = strDesc & strLineSeparator
        Next i
     Else
        strDesc = strDesc & "Metric strDesc not dimensioned"
     End If
     
    Select Case UMCDef.OddEvenProcessingMode
    Case oepUMCOddEvenProcessingMode.oepOddOnly: strAddnlText = "Process odd-numbered spectra only"
    Case oepUMCOddEvenProcessingMode.oepEvenOnly: strAddnlText = "Process even-numbered  spectra only"
    Case oepUMCOddEvenProcessingMode.oepOddEvenSequential: strAddnlText = "Process odd-numbered spectra then even-numbered spectra sequentially (and independently)"
    Case oepUMCOddEvenProcessingMode.oepProcessAll: strAddnlText = "Process all spectra"
    Case Else: strAddnlText = "Unknown type"
    End Select
    
    strDesc = strDesc & strLineSeparator & strAddnlText
    
    strDesc = strDesc & vbCrLf
End With
GetUMCIsoDefinitionText = strDesc
End Function

Private Function HUMCAddingSingleMemberUMCs() As Long
'------------------------------------------------------------------
'adds unconnected Isotopic distributions as a single-member classes
'returns number of added classes
'NOTE: this function makes sense only in a larger context of an UMC
'from Net procedure
'------------------------------------------------------------------
Dim i As Long
Dim Cnt As Long
Dim ISInd() As Long         ' In-scope index
Dim lngDataInScope As Long
Dim lngOriginalIndex As Long

On Error GoTo err_HUMCAddingSingleMemberUMCs

' Get a list of the data "in-scope"
lngDataInScope = GetISScope(CallerID, ISInd(), GelUMC(CallerID).def.DefScope)

ChangeStatus " Adding single-member LC-MS Features..."
With GelUMC(CallerID)
    For i = 1 To lngDataInScope
        lngOriginalIndex = ISInd(i)
        
        If HUMCIsoUsed(lngOriginalIndex) = HUMCNotUsed Then               'not used in any class
            
            If .UMCCnt > UBound(.UMCs) Then
                 ManageClasses CallerID, UMCManageConstants.UMCMngAdd
            End If
            
            With .UMCs(.UMCCnt)
                 .ClassCount = 1
                 ReDim .ClassMInd(0)
                 ReDim .ClassMType(0)
                 .ClassMInd(0) = lngOriginalIndex
                 .ClassMType(0) = glIsoType
                 .ClassRepInd = lngOriginalIndex
                 .ClassRepType = glIsoType
            End With
            .UMCCnt = .UMCCnt + 1
            
            Cnt = Cnt + 1
        End If
    Next i

End With
HUMCAddingSingleMemberUMCs = Cnt
Exit Function

err_HUMCAddingSingleMemberUMCs:
Select Case Err.Number
Case 9
    ' This code should never be reached
    Debug.Assert False
    
    If ManageClasses(CallerID, UMCManageConstants.UMCMngAdd) Then
       Resume
    Else
       LogErrors Err.Number, "frmUMCIonNet.HUMCAddingSingleMemberUMCs"
    End If
Case Else
    Debug.Print "Error in HUMCAddingSingleMemberUMCs: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->HUMCAddingSingleMemberUMCs"
    ' Do not attempt to resume
End Select
End Function

Private Sub InitializeLCMSFeatureInfo()
    Dim lngSpaceToReserve As Long
    
    lngSpaceToReserve = GelData(CallerID).IsoLines - 1
    If lngSpaceToReserve < 10 Then lngSpaceToReserve = 10
    
    mLCMSResultsMappingCount = 0
    ReDim mLCMSResultsMappingUMCs(lngSpaceToReserve - 1)
    ReDim mLCMSResultsMappingDataIndices(lngSpaceToReserve - 1)
End Sub

Public Sub InitializeUMCSearch()
    
    ' MonroeMod: This code was in Form_Activate
    
On Error GoTo InitializeUMCSearchErrorHandler

    Dim ScanRange As Long
    If bLoading Then
        CallerID = Me.Tag
        lblNetInfo.Caption = GetUMCIonNetInfo(CallerID)
        
        ' Clear the cached LCMSResultsMapping data
        mLCMSResultsMappingCount = 0
        
        ' Copy Def from GelSearchDef(CallerID).UMCDef to UMCDef
        ' Copy Def from GelSearchDef(CallerID).UMCIonNetDef to UMCIonNetDef
        If CallerID >= 1 And CallerID <= UBound(GelBody) Then
            UMCDef = GelSearchDef(CallerID).UMCDef
            UMCIonNetDef = GelSearchDef(CallerID).UMCIonNetDef
        End If
        
        If GelUMCIon(CallerID).NetCount > 0 Then                     'accept settings from caller
           MyDef = GelUMCIon(CallerID).ThisNetDef
           ChangeStatus " Number of lines: " & GelUMCIon(CallerID).NetCount
        Else                                                         'accept setting from UMCIonNetDef (default Def, or last Def used when form was Unloaded)
           MyDef = UMCIonNetDef
           ChangeStatus " No net structure found."
        End If
        
        ' MonroeMod: copy value from .UMCDrawType to .DrawType
        GelUMCDraw(CallerID).DrawType = glbPreferencesExpanded.UMCDrawType
        
        DisplayCurrentOptions
        
        bLoading = False
        GetScanRange CallerID, MinScan, MaxScan, ScanRange
        
        tbsTabStrip.Tab = 0
        tbsUMCRefinementOptions.Tab = 0
    End If

    Exit Sub

InitializeUMCSearchErrorHandler:
    Debug.Print "Error in InitializeUMCSearch: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->InitializeUMCSearch"
    Resume Next
    
End Sub

Private Function LoadFeatureInfoFromDisk(ByRef fso As FileSystemObject, ByVal strWorkingDirPath As String, ByVal strLCMSFeaturesFilePath As String, blnShowMessages As Boolean) As Boolean

    Dim strMessage As String
    Dim strResultsFilePath As String
    Dim strResultingMappingFilePath As String
    
    Dim strLineIn As String
    Dim strSplitLine() As String
    Dim strUMCIsoDefinition As String
    
    Dim tsInFile As TextStream
    Dim objFile As File
    
    Dim lngFileSizeBytes As Long
    Dim lngBytesRead As Long
    
    Dim sngPercentComplete As Single
    
    Dim blnSuccess As Boolean

    blnSuccess = False
    
    
    'strResultsFilePath = fso.BuildPath(strWorkingDirPath, fso.GetBaseName(strLCMSFeaturesFilePath) & "_Features.txt")
    
    strResultingMappingFilePath = fso.BuildPath(strWorkingDirPath, fso.GetBaseName(strLCMSFeaturesFilePath) & "_PeakToFeatureMap.txt")
    If Not fso.FileExists(strResultingMappingFilePath) Then
        strMessage = "LCMS feature to peak map file not found, unable to continue: " & vbCrLf & strResultingMappingFilePath
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            AddToAnalysisHistory CallerID, strMessage
        ElseIf blnShowMessages Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "File Not Found"
        End If
        blnSuccess = False
    Else
        Set objFile = fso.GetFile(strResultingMappingFilePath)
        lngFileSizeBytes = objFile.Size
        If lngFileSizeBytes < 1 Then lngFileSizeBytes = 1
        
        ' Populate mLCMSResultsMappingUMCs() with each item from the _PeakToFeatureMap file
        ' The calling function should have previously called InitializeLCMSFeatureInfo
        
        Set tsInFile = fso.OpenTextFile(strResultingMappingFilePath, ForReading, False)
        
        lngBytesRead = 0
        
        Do While Not tsInFile.AtEndOfStream
            strLineIn = tsInFile.ReadLine
            lngBytesRead = lngBytesRead + Len(strLineIn) + 2
            
            If Len(strLineIn) > 0 Then
                strSplitLine = Split(strLineIn, vbTab, 3)
                If UBound(strSplitLine) >= 1 Then
                    If IsNumeric(strSplitLine(0)) Then
                        ' Parse this line
                        
                        If mLCMSResultsMappingCount = UBound(mLCMSResultsMappingUMCs) Then
                            ReDim Preserve mLCMSResultsMappingUMCs((UBound(mLCMSResultsMappingUMCs) + 1) * 2 - 1)
                            ReDim Preserve mLCMSResultsMappingDataIndices(UBound(mLCMSResultsMappingUMCs))
                        End If
                        
                        mLCMSResultsMappingUMCs(mLCMSResultsMappingCount) = CLng(strSplitLine(0))
                        mLCMSResultsMappingDataIndices(mLCMSResultsMappingCount) = CLng(strSplitLine(1))
                        mLCMSResultsMappingCount = mLCMSResultsMappingCount + 1
                    
                        If mLCMSResultsMappingCount Mod 5000 = 0 Then
                            sngPercentComplete = lngBytesRead / lngFileSizeBytes * 100
                            ChangeStatus "Loading features from disk: " & Round(sngPercentComplete, 1) & "% complete"
                        End If
                    
                    End If
                End If
            End If
        Loop
        tsInFile.Close
        
        ChangeStatus "Loading features from disk: 100% complete"
        
        If mLCMSResultsMappingCount > 0 Then
            blnSuccess = True
            
            ' Copy current settings to caller structures
            GelUMCIon(CallerID).ThisNetDef = MyDef
            GelSearchDef(CallerID).UMCIonNetDef = MyDef
            
            strUMCIsoDefinition = GetUMCIsoDefinitionText(CallerID, False)
            strUMCIsoDefinition = Replace(strUMCIsoDefinition, ": ", " = ")
            strUMCIsoDefinition = Trim(Replace(strUMCIsoDefinition, vbCrLf, ""))
            If Right(strUMCIsoDefinition, 1) = ";" Then
               strUMCIsoDefinition = Left(strUMCIsoDefinition, Len(strUMCIsoDefinition) - 1)
            End If
            
            AddToAnalysisHistory CallerID, "Loaded features found using LCMSFeatureFinder.exe (" & AUTO_ANALYSIS_UMCIonNet & "); Mapping count = " & Trim(mLCMSResultsMappingCount) & "; " & strUMCIsoDefinition

        Else
            strMessage = "Empty data file (" & fso.GetFileName(strResultingMappingFilePath) & "); unable to continue"
            ChangeStatus strMessage
            
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                AddToAnalysisHistory CallerID, strMessage
            End If
        End If
        
    End If
    
    LoadFeatureInfoFromDisk = blnSuccess
    
End Function

Private Function ManageResArrays(ByVal ManageType As ArrayManagementType) As Boolean
Dim MaxResCount As Long
Dim IsoCnt As Long
Dim strMessage As String

On Error GoTo ManageResArraysErrorHandler

'''ChangeStatus " Managing results arrays..."
IsoCnt = GelData(CallerID).IsoLines
If IsoCnt <= 0 Then Exit Function
Select Case ManageType
Case amtErase
    ResCnt = 0
    Erase ResInd1:       Erase ResInd2:       Erase ResDist:      Erase ResEliminate
Case amtInitialize
    If MyDef.NETType < 10 Then
       MaxResCount = IsoCnt * MyDef.NETType
       If MaxResCount > MAX_NET_SIZE Then
          strMessage = "Net would be too big. Select lower Net type."
          If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
             MsgBox strMessage, vbOKOnly, glFGTU
          Else
             Debug.Assert False
             LogErrors Err.Number, "frmUMCIonNET->ManageResArrays, amtInitialize"
             AddToAnalysisHistory CallerID, "Error in UMCIonNet Searching: " & strMessage
          End If
          Exit Function
       Else
          ResCnt = 0
          ReDim ResInd1(MaxResCount - 1):        ReDim ResInd2(MaxResCount - 1)
          ReDim ResDist(MaxResCount - 1):        ReDim ResEliminate(MaxResCount - 1)
       End If
    Else                              'some other type of net; start with
       ResCnt = 0
       ReDim ResInd1(IsoCnt - 1):        ReDim ResInd2(IsoCnt - 1)
       ReDim ResDist(IsoCnt - 1):        ReDim ResEliminate(IsoCnt - 1)
    End If
Case amtAdd
    ReDim Preserve ResInd1(ResCnt + NET_ADD_RATE)
    ReDim Preserve ResInd2(ResCnt + NET_ADD_RATE)
    ReDim Preserve ResDist(ResCnt + NET_ADD_RATE)
    ReDim Preserve ResEliminate(ResCnt + NET_ADD_RATE)
Case amtTrim
    If ResCnt > 0 Then
        ReDim Preserve ResInd1(ResCnt - 1)
        ReDim Preserve ResInd2(ResCnt - 1)
        ReDim Preserve ResDist(ResCnt - 1)
        ReDim Preserve ResEliminate(ResCnt - 1)
    Else
        ReDim Preserve ResInd1(0)
        ReDim Preserve ResInd2(0)
        ReDim Preserve ResDist(0)
        ReDim Preserve ResEliminate(0)
    End If
End Select
ManageResArrays = True
Exit Function

ManageResArraysErrorHandler:
Debug.Print "Error in ManageResArrays: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->ManageResArrays"
Resume Next

End Function

Private Function MetricEuclid(i As Long, j As Long) As Double
'------------------------------------------------------------------
'returns Euclidean distance between two points in MyDef.NetDim-dim space
'i and j are indexes in data arrays; -1 on any error
'------------------------------------------------------------------
Dim k As Long
Dim TmpSum As Double
On Error GoTo err_MetricEuclid
For k = 0 To MyDef.NetDim - 1
    TmpSum = TmpSum + (DataVal(i, k) - DataVal(j, k)) ^ 2
Next k
MetricEuclid = Sqr(TmpSum)
Exit Function

err_MetricEuclid:
MetricEuclid = -1
End Function


Private Function MetricEuclidDim1(i As Long, j As Long) As Double
'------------------------------------------------------------------------
'returns Euclidean distance between two points for the optimization array
'------------------------------------------------------------------------
On Error Resume Next
MetricEuclidDim1 = Abs(OptValO(i) - OptValO(j))
End Function


Private Function MetricEuclidDim1Any(DimInd As Long, i As Long, j As Long) As Double
'----------------------------------------------------------------------------------
'returns Euclidean distance between two points for the data dimension DimInd
'----------------------------------------------------------------------------------
On Error Resume Next
MetricEuclidDim1Any = Abs(DataVal(i, DimInd) - DataVal(j, DimInd))
End Function


Private Function MetricHonduras(i As Long, j As Long) As Double
'----------------------------------------------------------------------
'returns Honduras distance between two points in MyDef.NetDim-dim space
'i and j are indexes in data arrays; -1 on any error
'----------------------------------------------------------------------
Dim k As Long
Dim TmpSum As Double
On Error GoTo err_MetricHonduras
For k = 0 To MyDef.NetDim - 1
    TmpSum = TmpSum + Abs(DataVal(i, k) - DataVal(j, k))
Next k
MetricHonduras = TmpSum
Exit Function

err_MetricHonduras:
MetricHonduras = -1
End Function


'One dimensional Honduras metric is the same as Euclidean
Private Function MetricHondurasDim1(i As Long, j As Long) As Double
'------------------------------------------------------------------------
'returns Honduras distance between two points for the optimization array
'------------------------------------------------------------------------
On Error Resume Next
MetricHondurasDim1 = Abs(OptValO(i) - OptValO(j))
End Function


Private Function MetricHondurasDim1Any(DimInd As Long, i As Long, j As Long) As Double
'------------------------------------------------------------------------------------
'returns Honduras distance between two points for the data dimension DimInd
'------------------------------------------------------------------------------------
On Error Resume Next
MetricHondurasDim1Any = Abs(DataVal(i, DimInd) - DataVal(j, DimInd))
End Function

Private Function MetricInfinity(i As Long, j As Long) As Double
'----------------------------------------------------------------------
'returns Infinity distance between two points in MyDef.NetDim-dim space
'i and j are indexes in data arrays; -1 on any error
'----------------------------------------------------------------------
Dim k As Long
Dim tmpMax As Double
Dim AbsDistance As Double
On Error GoTo err_MetricInfinity
tmpMax = 0
For k = 0 To MyDef.NetDim - 1
    AbsDistance = Abs(DataVal(i, k) - DataVal(j, k))
    If AbsDistance > tmpMax Then tmpMax = AbsDistance
Next k
MetricInfinity = tmpMax
Exit Function

err_MetricInfinity:
MetricInfinity = -1
End Function


'One dimensional Infinity metric is the same as Euclidean
Private Function MetricInfinityDim1(i As Long, j As Long) As Double
'------------------------------------------------------------------------
'returns Infinity distance between two points for the optimization array
'------------------------------------------------------------------------
On Error Resume Next
MetricInfinityDim1 = Abs(OptValO(i) - OptValO(j))
End Function


Private Function MetricInfinityDim1Any(DimInd As Long, i As Long, j As Long) As Double
'------------------------------------------------------------------------------------
'returns Infinity distance between two points for the data dimension DimInd
'------------------------------------------------------------------------------------
On Error Resume Next
MetricInfinityDim1Any = Abs(DataVal(i, DimInd) - DataVal(j, DimInd))
End Function

Private Sub PopulateComboBoxes()
    
    ' MonroeMod: The comboboxes are initialized here, rather than hard-coding their members
    
    Dim intIndex As Integer
    
    With cmbMetricType
        .Clear
        .AddItem "Euclidean"
        .AddItem "Honduras"
        .AddItem "Infinity"
    End With
    
    For intIndex = 0 To chkUse.Count - 1
        With cmbData(intIndex)
            .Clear
            .AddItem "Monoisotopic Mass"
            .AddItem "Average Mass"
            .AddItem "The Most Abundant Mass"
            .AddItem "SCAN"
            .AddItem "Fit"
            .AddItem "m/z"
            .AddItem "Generic NET"
            .AddItem "Charge STATE"
            .AddItem "Log (Abundance)"
        End With
    
        With cmbConstraint(intIndex)
            .Clear
            .AddItem "None"
            .AddItem "L.T."
            .AddItem "G.T."
        End With
    
        With cmbConstraintUnits(intIndex)
            .Clear
            .AddItem "Da"
            .AddItem "ppm"
            .Visible = False
        End With
    
    Next intIndex
    
    With cmbUMCRepresentative
        .Clear
        .AddItem "Highest Abundance"
        .AddItem "Best Isotopic Fit"
        .AddItem "First Scan Distribution"
        .AddItem "Last Scan Distribution"
        .AddItem "Median Scan Distribution"
    End With
    
    With cmbUMCAbu
        .Clear
        .AddItem "Average of Class Abu."
        .AddItem "Sum of Class Abu."
        .AddItem "Abu. of Class Representative"
        .AddItem "Median of Class Abundance"
        .AddItem "Max of Class Abu."
        .AddItem "Sum of Top X Members of Class"
    End With
    
    With cmbUMCMW
        .Clear
        .AddItem "Class Average"
        .AddItem "Mol.Mass Of Class Representative"
        .AddItem "Class Median"
        .AddItem "Average of Top X Members of Class"
        .AddItem "Median of Top X Members of Class"
    End With
    
    With cmbUMCDrawType
        .Clear
        .AddItem "Actual LC-MS Feature"
        .AddItem "LC-MS Feature Full Region"
        .AddItem "LC-MS Feature Intensity"
    End With
    
    With cboChargeStateAbuType
        .Clear
        .AddItem "Highest Abu Sum"
        .AddItem "Most Abu Member"
        .AddItem "Most Members"
    End With
    
    With cboMolecularMassField
        .Clear
        .AddItem "Average"
        .AddItem "Monoisotopic"
        .AddItem "Most abundant"
    End With
    
    With cboSplitUMCsScanGapBehavior
        .Clear
        .AddItem "Ignore scan gaps"
        .AddItem "Split if mass difference"
        .AddItem "Always split"
    End With
    
End Sub

Private Function PPMToDaIfNeeded(dblConstraintValue As Double, DimInd As Long, lngDataIndex As Long) As Double
    
    ' If .DataType is a mass type, and if .ContraintUnits is ppm, then convert
    '  dblConstraintValue from Da to ppm, using DataVal(lngDataIndex, DimInd) as
    '  the basis for the conversion
    
    Select Case MyDef.MetricData(DimInd).DataType
    Case uindUMCIonNetDimConstants.uindMonoMW, uindUMCIonNetDimConstants.uindAvgMW, uindUMCIonNetDimConstants.uindTmaMW
        If MyDef.MetricData(DimInd).ConstraintUnits = DATA_UNITS_MASS_PPM Then
            PPMToDaIfNeeded = dblConstraintValue / 1000000# * DataVal(lngDataIndex, DimInd)
        Else
            PPMToDaIfNeeded = dblConstraintValue
        End If
    Case Else
        PPMToDaIfNeeded = dblConstraintValue
    End Select
    
End Function

Private Function PrepareDataArrays() As Boolean
'------------------------------------------------------------------------
'prepares data arrays and returns True if successful
'------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim lngScanNumberRelativeIndex As Long
    
    Dim strMessage As String
    Dim ISInd() As Long         ' In-scope index

On Error GoTo err_PrepareDataArrays

    ChangeStatus " Preparing arrays..."

    If Not GetDataInScope(ISInd(), DataCnt) Then
        PrepareDataArrays = False
        Exit Function
    End If

    If Not UpdateNetDimInfo() Then
        PrepareDataArrays = False
        Exit Function
    End If

    ReDim DataOInd(DataCnt - 1)
    ReDim DataVal(DataCnt - 1, MyDef.NetDim - 1)
    
    Select Case MyDef.MetricType
    Case METRIC_EUCLIDEAN, METRIC_HONDURAS, METRIC_INFINITY
        With GelData(CallerID)
           For j = 0 To MyDef.NetDim - 1
               If MyDef.MetricData(j).Use Then
                  Select Case MyDef.MetricData(j).DataType
                  Case uindUMCIonNetDimConstants.uindMonoMW
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          DataVal(i - 1, j) = .IsoData(ISInd(i)).MonoisotopicMW * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindAvgMW
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          DataVal(i - 1, j) = .IsoData(ISInd(i)).AverageMW * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindTmaMW
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          DataVal(i - 1, j) = .IsoData(ISInd(i)).MostAbundantMW * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindScan
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          ' When processing odd-only or even-only scans in frmUMCSimple, we divide lngScanNumberRelativeIndex by 2 since we're only keeping every other scan
                          ' However, we will not do that in this function, since a scan gap of 1 is allowed for, and since that can mess up the minimum Scan Width filters applied by the LCMSFeatureFinder
                          lngScanNumberRelativeIndex = LookupScanNumberRelativeIndex(CallerID, .IsoData(ISInd(i)).ScanNumber)
                          DataVal(i - 1, j) = lngScanNumberRelativeIndex * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindFit
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          DataVal(i - 1, j) = .IsoData(ISInd(i)).Fit * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindMZ
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          DataVal(i - 1, j) = .IsoData(ISInd(i)).MZ * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindGenericNET
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          DataVal(i - 1, j) = ((.IsoData(ISInd(i)).ScanNumber - MinScan) / (MaxScan - MinScan)) * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindChargeState
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          DataVal(i - 1, j) = .IsoData(ISInd(i)).Charge * MyDef.MetricData(j).WeightFactor
                      Next i
                  Case uindUMCIonNetDimConstants.uindLogAbundance
                      For i = 1 To DataCnt
                          DataOInd(i - 1) = ISInd(i)
                          If .IsoData(ISInd(i)).Abundance > 0 Then
                            DataVal(i - 1, j) = Log(.IsoData(ISInd(i)).Abundance) / Log(10#) * MyDef.MetricData(j).WeightFactor
                          Else
                            ' Cannot perform Log(0)
                            DataVal(i - 1, j) = 0
                          End If
                      Next i
                  End Select
               End If
           Next j
        End With
    Case Else
        ' This shouldn't get reached
        Debug.Assert False
    End Select
    
    PrepareDataArrays = True
    Exit Function

err_PrepareDataArrays:
    LogErrors Err.Number, "frmUMCIonNet.PrepareArrays"
    ChangeStatus "Error preparing calculation structures."
    PrepareDataArrays = False
    
End Function

Private Function PrepareHUMCArrays() As Boolean
'--------------------------------------------------------------------
'prepares Unique Mass Classes calculation; returns True if successful
'--------------------------------------------------------------------
On Error GoTo err_PrepareHUMCArrays
HUMCIsoCnt = GelData(CallerID).IsoLines
'use arrays with same indexing as in IsoData arrays(element 0 will not be used)
ReDim HUMCIsoUsed(HUMCIsoCnt)
ReDim HUMCEquClsWk(HUMCIsoCnt - 1)               'here we will use 0th element
HUMCNetCnt = GelUMCIon(CallerID).NetCount
ReDim HUMCNetUsed(HUMCNetCnt - 1)
PrepareHUMCArrays = True
err_PrepareHUMCArrays:
End Function

Private Function PrepareOptimization() As Boolean
'-----------------------------------------------------------------
'creates and sorts optimization arrays; returns True if successful
'-----------------------------------------------------------------
Dim i As Long
Dim qsdMySort As New QSDouble
On Error GoTo err_PrepareOptimization
ChangeStatus " Creating optimization structures..."
ReDim OptIndO(DataCnt - 1)
ReDim OptValO(DataCnt - 1)
For i = 0 To DataCnt - 1
    OptIndO(i) = i
    OptValO(i) = DataVal(i, 0)
Next i
PrepareOptimization = qsdMySort.QSAsc(OptValO, OptIndO)
Exit Function


err_PrepareOptimization:
LogErrors Err.Number, "frmUMCIonNet.PrepareOptimization"
ChangeStatus " Error preparing optimization structures."
End Function

Private Sub ResetToDefaults()
    With glbPreferencesExpanded
        With .UMCIonNetOptions
            .ConnectionLengthPostFilterMaxNET = 0.2
            .UMCRepresentative = UMCFROMNet_REP_ABU
            .MakeSingleMemberClasses = False
            
            UMCRepresentative = .UMCRepresentative
            UMCMakeSingleMemberClasses = .MakeSingleMemberClasses
        End With
        
        ResetUMCAdvancedStatsOptions .UMCAdvancedStatsOptions
        ResetUMCAutoRefineOptions .UMCAutoRefineOptions
        
        .UMCDrawType = umcdt_ActualUMC
    End With
    
    SetDefaultUMCDef UMCDef
    SetDefaultUMCIonNetDef MyDef
        
    DisplayCurrentOptions
    
End Sub

Private Sub ResetToOldDefaults()
    Dim eResponse As VbMsgBoxResult
    
    eResponse = MsgBox("Are you sure you want to use old defaults?", vbQuestion Or vbYesNoCancel Or vbDefaultButton3, "Old Defaults")
    If eResponse <> vbYes Then Exit Sub
    
    ResetToDefaults
    
    ' These are old defaults, set in July 2003
    With glbPreferencesExpanded.UMCAutoRefineOptions
        .UMCAutoRefineRemoveCountLow = True
        .UMCAutoRefineMinLength = 2
        
        .UMCAutoRefineRemoveCountHigh = True
        .UMCAutoRefineMaxLength = 400
        
        .UMCAutoRefineRemoveMaxLengthPctAllScans = False
        .UMCAutoRefineMaxLengthPctAllScans = 15
        
        .UMCAutoRefinePercentMaxAbuToUseForLength = 33
        .TestLengthUsingScanRange = True
        .MinMemberCountWhenUsingScanRange = 2
    End With

    ' These are old defaults, set in July 2003
    SetOldDefaultUMCIonNetDef MyDef
    
    DisplayCurrentOptions
End Sub

Private Sub RemoveLongConnections(Optional intEditType As Integer = NET_EDIT_REJECT_LONG)
'--------------------------------------------------------------
'does editing of current display GelUMCIon structure
'--------------------------------------------------------------
On Error GoTo RemoveLongConnectionsErrorHandler
Dim strMessage As String
Dim lngConnectionsEliminated As Long

mCalculating = True
ShowHideCommandButtons mCalculating

Select Case intEditType
Case NET_EDIT_REJECT_LONG
    Dim TooLongConnection As Double
    If IsNumeric(txtNetEditTooDistant.Text) Then
       TooLongConnection = CDbl(txtNetEditTooDistant.Text)
       lngConnectionsEliminated = EliminateLongConnections_Net(TooLongConnection)
       
       AddToAnalysisHistory CallerID, "Removed long connections (" & AUTO_ANALYSIS_UMCIonNet & "); New connection count = " & Trim(GelUMCIon(CallerID).NetCount) & "; Connections removed = " & Trim(lngConnectionsEliminated)
       
       txtRejectLongConnections = CDbl(GelUMCIon(CallerID).ThisNetDef.TooDistant)
    Else
       strMessage = "This argument should be positive number. (txtNetEditTooDistant textbox)"
       If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strMessage, vbOKOnly, glFGTU
          txtNetEditTooDistant.SetFocus
       Else
          Debug.Assert False
          LogErrors Err.Number, "frmUMCIonNet->RemoveLongConnections, txtNetEditTooDistant is not numeric"
          AddToAnalysisHistory CallerID, "Error in UMCIonNet Searching: " & strMessage
       End If
    End If
End Select

mCalculating = False
ShowHideCommandButtons mCalculating

Exit Sub

RemoveLongConnectionsErrorHandler:
Debug.Print "Error in RemoveLongConnections: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->RemoveLongConnections"
Resume Next
End Sub

Private Sub SetDefinition()
'--------------------------------------------------------------
'sets definitions frm structures to control properties
'--------------------------------------------------------------
Dim i As Long
On Error GoTo err_SetDefinition
With MyDef
    txtRejectLongConnections.Text = .TooDistant
    txtNETType.Text = .NETType
    cmbMetricType.ListIndex = .MetricType
    For i = 0 To .NetDim - 1
        With .MetricData(i)
            If .Use Then
               chkUse(i) = vbChecked
            Else
               chkUse(i) = vbUnchecked
            End If
            cmbData(i).ListIndex = .DataType
            ' Note: cmbConstraintUnits() is updated inside DisplayDynamicUnits
            cmbConstraint(i).ListIndex = .ConstraintType
            txtWeightingFactor(i).Text = .WeightFactor
            txtConstraint(i).Text = .ConstraintValue
            cmbConstraintUnits(i).ListIndex = .ConstraintUnits
        End With
    Next i
End With

DisplayDynamicUnits

Exit Sub

err_SetDefinition:
Debug.Print "Error in SetUMCDefinition: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->SetDefinition"
ChangeStatus "Error accepting net definition."
End Sub

Private Sub SetOldDefaultUMCIonNetDef(ByRef udtUMCIonNetDef As UMCIonNetDefinition)
    ' These are old defaults, set in July 2003
    With udtUMCIonNetDef
        .MetricType = METRIC_EUCLIDEAN
        .NETType = Net_SPIDER_66
        .NetDim = 5
        .NetActualDim = 5
        .TooDistant = 0.1
        ReDim .MetricData(.NetDim - 1)
        .MetricData(0).Use = True:  .MetricData(0).DataType = uindUMCIonNetDimConstants.uindMonoMW:   .MetricData(0).WeightFactor = 0.5:   .MetricData(0).ConstraintType = Net_CT_LT:       .MetricData(0).ConstraintValue = 0.025: .MetricData(0).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(1).Use = True:  .MetricData(1).DataType = uindUMCIonNetDimConstants.uindAvgMW:    .MetricData(1).WeightFactor = 0.5:   .MetricData(1).ConstraintType = Net_CT_LT:       .MetricData(1).ConstraintValue = 0.025: .MetricData(1).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(2).Use = True:  .MetricData(2).DataType = uindUMCIonNetDimConstants.uindLogAbundance:   .MetricData(2).WeightFactor = 0.1:   .MetricData(2).ConstraintType = Net_CT_None:     .MetricData(2).ConstraintValue = 0.1:   .MetricData(2).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(3).Use = True:  .MetricData(3).DataType = uindUMCIonNetDimConstants.uindScan:      .MetricData(3).WeightFactor = 0.01:   .MetricData(3).ConstraintType = Net_CT_None:    .MetricData(3).ConstraintValue = 0.01:  .MetricData(3).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(4).Use = True:  .MetricData(4).DataType = uindUMCIonNetDimConstants.uindFit:       .MetricData(4).WeightFactor = 0.1:    .MetricData(4).ConstraintType = Net_CT_None:    .MetricData(4).ConstraintValue = 0.01:  .MetricData(4).ConstraintUnits = DATA_UNITS_MASS_DA
    End With
End Sub

Private Sub SetUMCDefinition()
'----------------------------------------------------------------------------
'sets definitions for UMC from Net procedure based on some settings of UMCDef (mirred at GelUMC().def)
'----------------------------------------------------------------------------
On Error GoTo SetUMCDefinitionErrorHandler

With UMCDef
    cmbUMCMW.ListIndex = .ClassMW
    cmbUMCAbu.ListIndex = .ClassAbu
    cboChargeStateAbuType.ListIndex = .ChargeStateStatsRepType
    
    SetMolecularMassFieldDropdown CInt(.MWField)
    SetCheckBox chkUseMostAbuChargeStateStatsForClassStats, .UMCClassStatsUseStatsFromMostAbuChargeState
    
    optDefScope(.DefScope).Value = True
    optEvenOddScanFilter(.OddEvenProcessingMode).Value = True
    
    SetCheckBox chkInterpolateMissingIons, .InterpolateGaps
    txtInterpolateMaxGapSize = .InterpolateMaxGapSize
    txtHoleSize = .GapMaxSize
    
End With

' Additional UMCIonNet options
With glbPreferencesExpanded.UMCIonNetOptions
    If .ConnectionLengthPostFilterMaxNET = 0 Then
        .ConnectionLengthPostFilterMaxNET = 0.2
    End If
    txtNetEditTooDistant.Text = .ConnectionLengthPostFilterMaxNET
    cmbUMCRepresentative.ListIndex = .UMCRepresentative
    SetCheckBox chkUseUntangledAsSingle, .MakeSingleMemberClasses
End With

Exit Sub

SetUMCDefinitionErrorHandler:
Debug.Print "Error in SetUMCDefinition: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->SetUMCDefinition"
Resume Next
End Sub

Private Sub ShowHideCommandButtons(ByVal blnCalculating As Boolean)
    Dim blnShowConnectionsButtons As Boolean

    fraDREAMS.Enabled = Not blnCalculating
    fraUMCScope.Enabled = Not blnCalculating
    fraNET(0).Enabled = Not blnCalculating
    fraLCMSFeatureStats.Enabled = Not blnCalculating
    fraOptionFrame(0).Enabled = Not blnCalculating
    fraOptionFrame(1).Enabled = Not blnCalculating
    fraClassAbundanceTopX.Enabled = Not blnCalculating
    fraClassMassTopX.Enabled = Not blnCalculating
    
    blnShowConnectionsButtons = Not cChkBox(chkUseLCMSFeatureFinder.Value)
    
    cmdFindConnections.Visible = blnShowConnectionsButtons
        
    cmdFindConnections.Visible = blnShowConnectionsButtons And Not blnCalculating
    cmdFindConnectionsThenUMCs.Visible = Not blnCalculating
    cmdRemoveLongConnections.Visible = Not blnCalculating
    cmdFindUMCsUsingNETConnections.Visible = Not blnCalculating
    cmdClose.Visible = Not blnCalculating
    
    cmdReportUMC.Visible = Not blnCalculating
    cmdResetToOldDefaults.Visible = Not blnCalculating
    cmdResetToDefaults(0).Visible = Not blnCalculating
    cmdResetToDefaults(1).Visible = Not blnCalculating

    cmdAbortFindConnections.Visible = blnCalculating
    cmdAbortProcessing.Visible = blnCalculating
    
    chkUseLCMSFeatureFinder.Enabled = Not blnCalculating

    txtInterpolateMaxGapSize.Enabled = Not blnCalculating
End Sub

Public Function StartUMCSearch() As Boolean
    ' This sub should be called after calling InitializeUMCSearch
    ' It is intended to be called during AutoAnalysis
    '
    ' Returns True if Success, False if Error
    
    Dim blnSuccess As Boolean
    Dim blnAbortedProcessDuringAutoAnalysis As Boolean
    
On Error GoTo StartUMCSearchErrorHandler

    If cChkBox(chkUseLCMSFeatureFinder.Value) Then
        blnSuccess = FindUMCsUsingLCMSFeatureFinder(False)
        
        If Not blnSuccess And Not mAbortProcess Then
            ' Search failed; try to find LC-MS Features using the built-in UMC finding code
            AddToAnalysisHistory CallerID, "Warning: Unable to find LC-MS Features using the LCMS Feature Finder; will instead use the built-in finder"
        End If
    Else
        blnSuccess = False
    End If
    
    If Not blnSuccess Then
        ' 1. Find the connections
        tbsTabStrip.Tab = 0
        FindIonNetConnections
        
        ' If auto-analyzing, ignore the mAbortProcess state
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled And mAbortProcess Then
            blnAbortedProcessDuringAutoAnalysis = True
            mAbortProcess = False
        End If
        
        If Not mAbortProcess Then
            ' 2. Find the LC-MS Features
            tbsTabStrip.Tab = 2
            blnSuccess = FormClassesFromNETsWrapper(False)
        End If
    End If
    
    With GelP_D_L(CallerID)
        If .DltLblType <> ptS_Dlt And .DltLblType <> ptS_Lbl And .DltLblType <> ptS_DltLbl Then
            .SyncWithUMC = False
        End If
    End With
    
    mAbortProcess = mAbortProcess Or blnAbortedProcessDuringAutoAnalysis
    StartUMCSearch = Not mAbortProcess And blnSuccess
    
    Exit Function
    
StartUMCSearchErrorHandler:
    Debug.Print "Error in frmUMCIonNet.StartUMCSearch: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->StartUMCSearch"
    ' Do not attempt to resume
    StartUMCSearch = False
    
End Function

' Note: Status is called by AutoRefineUMCs and by SplitUMCsByAbundance
Public Sub Status(ByVal StatusMsg As String)
    ChangeStatus StatusMsg
End Sub

Private Function SubjectToConstraintEuclid(i As Long, j As Long) As Boolean
'-------------------------------------------------------------------------
'returns True if any used dimension is subject to constraint rule
'-------------------------------------------------------------------------
Dim DimInd As Long

On Error GoTo exit_SubjectToConstrainEuclid
For DimInd = 0 To MyDef.NetDim - 1
  With MyDef.MetricData(DimInd)
    If .Use Then
       Select Case .ConstraintType
       Case Net_CT_None         'no constraint
       Case Net_CT_LT           'distance in this dimension has to be less than constraint value
            If MetricEuclidDim1Any(DimInd, i, j) >= PPMToDaIfNeeded(.ConstraintValue, DimInd, i) Then
               SubjectToConstraintEuclid = True
               Exit Function
            End If
       Case Net_CT_GT           'distance in this dimension has to be more than constraint value
            If MetricEuclidDim1Any(DimInd, i, j) <= PPMToDaIfNeeded(.ConstraintValue, DimInd, i) Then
               SubjectToConstraintEuclid = True
               Exit Function
            End If
       End Select
    End If
  End With
Next DimInd
exit_SubjectToConstrainEuclid:
End Function

Private Function SubjectToConstraintHonduras(i As Long, j As Long) As Boolean
'-------------------------------------------------------------------------
'returns True if any used dimension is subject to constraint rule
'-------------------------------------------------------------------------
Dim DimInd As Long
On Error GoTo exit_SubjectToConstrainHonduras
For DimInd = 0 To MyDef.NetDim - 1
  With MyDef.MetricData(DimInd)
    If .Use Then
       Select Case .ConstraintType
       Case Net_CT_None         'no constraint
       Case Net_CT_LT           'distance in this dimension has to be less than constraint value
            If MetricHondurasDim1Any(DimInd, i, j) >= PPMToDaIfNeeded(.ConstraintValue, DimInd, i) Then
               SubjectToConstraintHonduras = True
               Exit Function
            End If
       Case Net_CT_GT           'distance in this dimension has to be more than constraint value
            If MetricHondurasDim1Any(DimInd, i, j) <= PPMToDaIfNeeded(.ConstraintValue, DimInd, i) Then
               SubjectToConstraintHonduras = True
               Exit Function
            End If
       End Select
    End If
  End With
Next DimInd
exit_SubjectToConstrainHonduras:
End Function


Private Function SubjectToConstraintInfinity(i As Long, j As Long) As Boolean
'---------------------------------------------------------------------------
'returns True if any used dimension is subject to constraint rule
'---------------------------------------------------------------------------
Dim DimInd As Long
On Error GoTo exit_SubjectToConstrainInfinity
For DimInd = 0 To MyDef.NetDim - 1
  With MyDef.MetricData(DimInd)
    If .Use Then
       Select Case .ConstraintType
       Case Net_CT_None         'no constraint
       Case Net_CT_LT           'distance in this dimension has to be less than constraint value
            If MetricInfinityDim1Any(DimInd, i, j) >= PPMToDaIfNeeded(.ConstraintValue, DimInd, i) Then
               SubjectToConstraintInfinity = True
               Exit Function
            End If
       Case Net_CT_GT           'distance in this dimension has to be more than constraint value
            If MetricInfinityDim1Any(DimInd, i, j) <= PPMToDaIfNeeded(.ConstraintValue, DimInd, i) Then
               SubjectToConstraintInfinity = True
               Exit Function
            End If
       End Select
    End If
  End With
Next DimInd
exit_SubjectToConstrainInfinity:
End Function

Private Sub UpdateDynamicControls()
    ' Update the UMC auto refine length labels
    If glbPreferencesExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange Then
        chkRemoveLoCnt.Caption = "Remove cls. with less than"
        chkRemoveHiCnt.Caption = "Remove cls. with length over"
        lblAutoRefineLengthLabel(0) = "scans"
        lblAutoRefineLengthLabel(1) = "scans"
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

    If CDblSafe(txtClassAbuTopXMinAbu) <= 0 And CDblSafe(txtClassAbuTopXMaxAbu) <= 0 Then
        lblClassAbuTopXMinMembers = "Maximum members to include"
    Else
        lblClassAbuTopXMinMembers = "Minimum members to include"
    End If

    If CDblSafe(txtClassMassTopXMinAbu) <= 0 And CDblSafe(txtClassMassTopXMaxAbu) <= 0 Then
        lblClassMassTopXMinMembers = "Maximum members to include"
    Else
        lblClassMassTopXMinMembers = "Minimum members to include"
    End If

End Sub

Private Function UpdateNetDimInfo() As Boolean

    Dim strMessage As String
    Dim i As Integer
    
    MyDef.NetDim = chkUse.Count

    ' Update .NetActualDim
    MyDef.NetActualDim = 0
    For i = 0 To chkUse.Count - 1
        If chkUse(i).Value = vbChecked Then MyDef.NetActualDim = MyDef.NetActualDim + 1
    Next i
    
    If MyDef.NetActualDim < 1 Then
       strMessage = "At least one data dimension has to be selected."
       If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strMessage, vbOKOnly, glFGTU
       Else
          Debug.Assert False
          LogErrors Err.Number, "frmUMCIonNet->UpdateNetDimInfo, MyDef.NetActualDim < 1"
          AddToAnalysisHistory CallerID, "Error in UMCIonNet Searching: " & strMessage
       End If
       UpdateNetDimInfo = False
       Exit Function
    End If
    
    UpdateNetDimInfo = True

End Function

Private Function ValidateClassStatOptions() As Boolean
    If UMCRepresentative < 0 Then
       MsgBox "Class representative type not selected.", vbOKOnly, glFGTU
       cmbUMCRepresentative.SetFocus
        ValidateClassStatOptions = False
       Exit Function
    End If
    If UMCDef.ClassMW < 0 Then
       MsgBox "Class mass type not selected.", vbOKOnly, glFGTU
       cmbUMCMW.SetFocus
        ValidateClassStatOptions = False
       Exit Function
    End If
    If UMCDef.ClassAbu < 0 Then
       MsgBox "Class abundance type not selected.", vbOKOnly, glFGTU
       cmbUMCAbu.SetFocus
        ValidateClassStatOptions = False
       Exit Function
    End If
    
    ValidateClassStatOptions = True
End Function

Private Sub cboMolecularMassField_Click()
    If mCalculating Then
        SetMolecularMassFieldDropdown CInt(UMCDef.MWField)
    Else
        UMCDef.MWField = GetMolecularMassFieldFromDropdown
    End If
End Sub

Private Sub Form_Activate()
    InitializeUMCSearch
End Sub

Private Sub Form_Load()
    mOneSecond = 1 / 24 / 60 / 60
    bLoading = True
    mCalculating = False
    mAbortProcess = False
    PopulateComboBoxes
    ShowHideCommandButtons False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UMCIonNetDef = MyDef
End Sub

Private Sub optDefScope_Click(Index As Integer)
    If mCalculating Then
        optDefScope(UMCDef.DefScope).Value = True
    Else
        UMCDef.DefScope = Index
    End If
End Sub

Private Sub optEvenOddScanFilter_Click(Index As Integer)
    If mCalculating Then
        optEvenOddScanFilter(UMCDef.OddEvenProcessingMode).Value = True
    Else
        UMCDef.OddEvenProcessingMode = Index
    End If
End Sub

Private Sub txtAutoRefineMinimumMemberCount_LostFocus()
If IsNumeric(txtAutoRefineMinimumMemberCount.Text) Then
    glbPreferencesExpanded.UMCAutoRefineOptions.MinMemberCountWhenUsingScanRange = Abs(CLng(txtAutoRefineMinimumMemberCount.Text))
Else
   MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
   txtAutoRefineMinimumMemberCount.SetFocus
End If
End Sub

Private Sub txtClassAbuTopXMaxAbu_Change()
    UpdateDynamicControls
End Sub

Private Sub txtClassAbuTopXMaxAbu_Lostfocus()
    ValidateTextboxValueDbl txtClassAbuTopXMaxAbu, 0, 1E+300, 0
    glbPreferencesExpanded.UMCAdvancedStatsOptions.ClassAbuTopXMaxAbu = CDblSafe(txtClassAbuTopXMaxAbu)
End Sub

Private Sub txtClassAbuTopXMinAbu_Change()
    UpdateDynamicControls
End Sub

Private Sub txtClassAbuTopXMinAbu_Lostfocus()
    ValidateTextboxValueDbl txtClassAbuTopXMinAbu, 0, 1E+300, 0
    glbPreferencesExpanded.UMCAdvancedStatsOptions.ClassAbuTopXMinAbu = CDblSafe(txtClassAbuTopXMinAbu)
End Sub

Private Sub txtClassAbuTopXMinMembers_Lostfocus()
    ValidateTextboxValueLng txtClassAbuTopXMinMembers, 0, 100000, 3
    glbPreferencesExpanded.UMCAdvancedStatsOptions.ClassAbuTopXMinMembers = CLngSafe(txtClassAbuTopXMinMembers)
End Sub

Private Sub txtClassMassTopXMaxAbu_Change()
    UpdateDynamicControls
End Sub

Private Sub txtClassMassTopXMaxAbu_Lostfocus()
    ValidateTextboxValueDbl txtClassMassTopXMaxAbu, 0, 1E+300, 0
    glbPreferencesExpanded.UMCAdvancedStatsOptions.ClassMassTopXMaxAbu = CDblSafe(txtClassMassTopXMaxAbu)
End Sub

Private Sub txtClassMassTopXMinAbu_Change()
    UpdateDynamicControls
End Sub

Private Sub txtClassMassTopXMinAbu_Lostfocus()
    ValidateTextboxValueDbl txtClassMassTopXMinAbu, 0, 1E+300, 0
    glbPreferencesExpanded.UMCAdvancedStatsOptions.ClassMassTopXMinAbu = CDblSafe(txtClassMassTopXMinAbu)
End Sub

Private Sub txtClassMassTopXMinMembers_Lostfocus()
    ValidateTextboxValueLng txtClassMassTopXMinMembers, 0, 100000, 3
    glbPreferencesExpanded.UMCAdvancedStatsOptions.ClassMassTopXMinMembers = CLngSafe(txtClassMassTopXMinMembers)
End Sub

Private Sub txtConstraint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mCalculating Then KeyCode = 0
End Sub

Private Sub txtConstraint_KeyPress(Index As Integer, KeyAscii As Integer)
    If mCalculating Then KeyAscii = 0
End Sub

Private Sub txtConstraint_LostFocus(Index As Integer)
    On Error Resume Next
    If IsNumeric(txtConstraint(Index).Text) Then
       MyDef.MetricData(Index).ConstraintValue = CDbl(txtConstraint(Index).Text)
    Else
       MsgBox "This argument should be positive number.", vbOKOnly, glFGTU
       txtConstraint(Index).SetFocus
    End If
End Sub

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

Private Sub txtHoleSize_LostFocus()
    If IsNumeric(txtHoleSize.Text) Then
       UMCDef.GapMaxSize = CLng(txtHoleSize.Text)
    Else
       MsgBox "This argument should be integer value.", vbOKOnly
       txtHoleSize.SetFocus
    End If
End Sub

Private Sub txtInterpolateMaxGapSize_LostFocus()
    If IsNumeric(txtInterpolateMaxGapSize.Text) Then
       UMCDef.InterpolateMaxGapSize = CLng(txtInterpolateMaxGapSize.Text)
    Else
       MsgBox "This argument should be integer value.", vbOKOnly
       txtInterpolateMaxGapSize.SetFocus
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

Private Sub txtMaxLengthPctAllScans_Lostfocus()
    If IsNumeric(txtMaxLengthPctAllScans.Text) Then
        glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineMaxLengthPctAllScans = Abs(CLng(txtMaxLengthPctAllScans.Text))
    Else
       MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
       txtMaxLengthPctAllScans.SetFocus
    End If
End Sub

Private Sub txtNetEditTooDistant_Lostfocus()
    If IsNumeric(txtNetEditTooDistant.Text) Then
        glbPreferencesExpanded.UMCIonNetOptions.ConnectionLengthPostFilterMaxNET = Abs(txtNetEditTooDistant.Text)
    Else
       MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
       txtNetEditTooDistant.SetFocus
    End If
End Sub

Private Sub txtNETType_LostFocus()
    On Error Resume Next
    If IsNumeric(txtNETType.Text) Then
       MyDef.NETType = CLng(txtNETType.Text)
    Else
       MsgBox "This argument should be positive integer.", vbOKOnly, glFGTU
       txtNETType.SetFocus
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

Private Sub txtRejectLongConnections_LostFocus()
    If IsNumeric(txtRejectLongConnections.Text) Then
       MyDef.TooDistant = CDbl(txtRejectLongConnections.Text)
    Else
       MsgBox "This argument should be positive number.", vbOKOnly, glFGTU
       txtRejectLongConnections.SetFocus
    End If
End Sub
    
Private Sub txtSplitUMCsMaximumPeakCount_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsMaximumPeakCount, 2, 100000, 6
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.MaximumPeakCountToSplitUMC = CLngSafe(txtSplitUMCsMaximumPeakCount)
End Sub

Private Sub txtSplitUMCsMinimumDifferenceInAvgPpmMass_LostFocus()
    ValidateTextboxValueDbl txtSplitUMCsMinimumDifferenceInAvgPpmMass, 0, 10000#, 4
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.MinimumDifferenceInAveragePpmMassToSplit = CDblSafe(txtSplitUMCsMinimumDifferenceInAvgPpmMass)
End Sub

Private Sub txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax, 0, 100, 15
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.PeakDetectIntensityThresholdPercentageOfMaximum = CLngSafe(txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax)
End Sub

Private Sub txtSplitUMCsPeakPickingMinimumWidth_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsPeakPickingMinimumWidth, 0, 1000, 4
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.PeakWidthPointsMinimum = CLngSafe(txtSplitUMCsPeakPickingMinimumWidth)
End Sub

Private Sub txtSplitUMCsStdDevMultiplierForSplitting_LostFocus()
    ValidateTextboxValueDbl txtSplitUMCsStdDevMultiplierForSplitting, 0, 1000, 1
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.StdDevMultiplierForSplitting = CSngSafe(txtSplitUMCsStdDevMultiplierForSplitting)
End Sub

Private Sub txtWeightingFactor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mCalculating Then KeyCode = 0
End Sub

Private Sub txtWeightingFactor_KeyPress(Index As Integer, KeyAscii As Integer)
    If mCalculating Then KeyAscii = 0
End Sub

Private Sub txtWeightingFactor_LostFocus(Index As Integer)
    On Error Resume Next
    If IsNumeric(txtWeightingFactor(Index).Text) Then
       MyDef.MetricData(Index).WeightFactor = CDbl(txtWeightingFactor(Index).Text)
    Else
       MsgBox "This arument should be positive number.", vbOKOnly, glFGTU
       txtWeightingFactor(Index).SetFocus
    End If
End Sub

Private Sub cmdAbortFindConnections_Click()
    AbortProcessing
End Sub

Private Sub cmdAbortProcessing_Click()
    AbortProcessing
End Sub

Private Sub cmdClose_Click()
    Dim eResponse As VbMsgBoxResult
    
    If mCalculating And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Calculations are currently in progress.  Abort them and close the window?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Abort Processing")
        If eResponse <> vbYes Then Exit Sub
        AbortProcessing
    End If
    
    Unload Me
End Sub

Private Sub cmdFindConnections_Click()
    If mCalculating Then Exit Sub
    
    If cChkBox(chkUseLCMSFeatureFinder.Value) Then
        MsgBox "Finding connections is not available when the LCMS Feature Finder external app mode is enabled", vbExclamation + vbOKOnly, "Not Applicable"
    Else
        FindIonNetConnections
    End If
End Sub

Private Sub cmdFindConnectionsThenUMCs_Click()
    If mCalculating Then Exit Sub
    Dim blnUseExternalFinder As Boolean
    
    blnUseExternalFinder = cChkBox(chkUseLCMSFeatureFinder.Value)
    
    If blnUseExternalFinder Then
        FindUMCsUsingLCMSFeatureFinder True
    Else
        FindIonNetConnections
        If Not mAbortProcess Then
            tbsTabStrip.Tab = 2
            FormClassesFromNETsWrapper False
        End If
    End If
End Sub

Private Sub cmdFindUMCsUsingNETConnections_Click()
    If mCalculating Then Exit Sub
    
    If cChkBox(chkUseLCMSFeatureFinder.Value) Then
        If mLCMSResultsMappingCount > 0 Then
            BuildUMCsUsingmLCMSResultsMapping True
        Else
            FindUMCsUsingLCMSFeatureFinder True
        End If
    Else
        FormClassesFromNETsWrapper True
    End If

End Sub

Private Sub cmdRemoveLongConnections_Click()
    If mCalculating Then Exit Sub
    RemoveLongConnections NET_EDIT_REJECT_LONG
End Sub

Private Sub cmdReportUMC_Click()
    If mCalculating Then Exit Sub
    
    Me.MousePointer = vbHourglass
    ChangeStatus "Generating LC-MS Feature report..."
    
    Call ReportUMC(CallerID, "UMCIonNet" & vbCrLf & GetUMCIsoDefinitionText(CallerID))
    
    ChangeStatus ""
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdResetToDefaults_Click(Index As Integer)
    ResetToDefaults
End Sub

Private Sub cmdResetToOldDefaults_Click()
    ResetToOldDefaults
End Sub

Private Sub cboChargeStateAbuType_Click()
    If mCalculating Then
        cboChargeStateAbuType.ListIndex = UMCDef.ChargeStateStatsRepType
    Else
        UMCDef.ChargeStateStatsRepType = cboChargeStateAbuType.ListIndex
    End If
End Sub

Private Sub cboSplitUMCsScanGapBehavior_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.ScanGapBehavior = cboSplitUMCsScanGapBehavior.ListIndex
End Sub

Private Sub chkInterpolateMissingIons_Click()
    If mCalculating Then
        SetCheckBox chkInterpolateMissingIons, UMCDef.InterpolateGaps
    Else
        UMCDef.InterpolateGaps = cChkBox(chkInterpolateMissingIons)
    End If
End Sub

Private Sub chkRemoveMaxLengthPctAllScans_Click()
     glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveMaxLengthPctAllScans = cChkBox(chkRemoveMaxLengthPctAllScans)
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

Private Sub chkSplitUMCsByExaminingAbundance_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCsByAbundance = cChkBox(chkSplitUMCsByExaminingAbundance)
End Sub

Private Sub chkUse_Click(Index As Integer)
    On Error Resume Next
    If mCalculating Then
        SetCheckBox chkUse(Index), MyDef.MetricData(Index).Use
    Else
        MyDef.MetricData(Index).Use = (chkUse(Index).Value = vbChecked)
    End If
End Sub

Private Sub chkUseLCMSFeatureFinder_Click()
    If mCalculating Then Exit Sub
    ShowHideCommandButtons False
End Sub

Private Sub chkUseMostAbuChargeStateStatsForClassStats_Click()
    If mCalculating Then
        SetCheckBox chkUseMostAbuChargeStateStatsForClassStats, UMCDef.UMCClassStatsUseStatsFromMostAbuChargeState
    Else
        UMCDef.UMCClassStatsUseStatsFromMostAbuChargeState = cChkBox(chkUseMostAbuChargeStateStatsForClassStats)
    End If
End Sub

Private Sub chkUseUntangledAsSingle_Click()
    If mCalculating Then
        SetCheckBox chkUseUntangledAsSingle, UMCMakeSingleMemberClasses
    Else
        UMCMakeSingleMemberClasses = cChkBox(chkUseUntangledAsSingle.Value)
        glbPreferencesExpanded.UMCIonNetOptions.MakeSingleMemberClasses = UMCMakeSingleMemberClasses
    End If
End Sub

Private Sub cmbConstraint_Click(Index As Integer)
    On Error Resume Next
    If mCalculating Then
        cmbConstraint(Index).ListIndex = MyDef.MetricData(Index).ConstraintType
    Else
        MyDef.MetricData(Index).ConstraintType = cmbConstraint(Index).ListIndex
        DisplayDynamicUnits
    End If
End Sub

Private Sub cmbData_Click(Index As Integer)
    On Error Resume Next
    If mCalculating Then
        cmbData(Index).ListIndex = MyDef.MetricData(Index).DataType
    Else
        MyDef.MetricData(Index).DataType = cmbData(Index).ListIndex
        DisplayDynamicUnits
    End If
End Sub

Private Sub cmbConstraintUnits_Click(Index As Integer)
    On Error Resume Next
    If mCalculating Then
        cmbConstraintUnits(Index).ListIndex = MyDef.MetricData(Index).ConstraintUnits
    Else
        MyDef.MetricData(Index).ConstraintUnits = cmbConstraintUnits(Index).ListIndex
        
        If Not bLoading Then
            If cmbConstraintUnits(Index).ListIndex = DATA_UNITS_MASS_DA Then
                ' Convert the constraint tolerance from ppm to Da, assuming 1000 m/z
                txtConstraint(Index) = PPMToMass(txtConstraint(Index), 1000)
            Else
                ' Convert the constraint tolerance from Da to ppm, assuming 1000 m/z
                txtConstraint(Index) = MassToPPM(txtConstraint(Index), 1000)
            End If
            If IsNumeric(txtConstraint(Index).Text) Then
                MyDef.MetricData(Index).ConstraintValue = CDbl(txtConstraint(Index).Text)
            End If
        End If
    End If
End Sub

Private Sub cmbMetricType_Click()
    If mCalculating Then
        On Error Resume Next
        cmbMetricType.ListIndex = MyDef.MetricType
    Else
        MyDef.MetricType = cmbMetricType.ListIndex
    End If
End Sub

Private Sub cmbUMCAbu_Click()
    If mCalculating Then
        On Error Resume Next
        cmbUMCAbu.ListIndex = UMCDef.ClassAbu
    Else
        UMCDef.ClassAbu = cmbUMCAbu.ListIndex
    End If
End Sub

Private Sub cmbUMCDrawType_Click()
    If mCalculating Then
        On Error Resume Next
        cmbUMCDrawType.ListIndex = GelUMCDraw(CallerID).DrawType
    Else
        GelUMCDraw(CallerID).DrawType = cmbUMCDrawType.ListIndex
        glbPreferencesExpanded.UMCDrawType = cmbUMCDrawType.ListIndex
    End If
End Sub

Private Sub cmbUMCMW_Click()
    If mCalculating Then
        On Error Resume Next
        cmbUMCMW.ListIndex = UMCDef.ClassMW
    Else
        UMCDef.ClassMW = cmbUMCMW.ListIndex
    End If
End Sub

Private Sub cmbUMCRepresentative_Click()
    If mCalculating Then
        On Error Resume Next
        cmbUMCRepresentative.ListIndex = UMCRepresentative
    Else
        UMCRepresentative = cmbUMCRepresentative.ListIndex
    End If
End Sub



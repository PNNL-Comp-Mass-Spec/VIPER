VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUMCIonNet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UMC Ion Networks"
   ClientHeight    =   5685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   135
      Top             =   4920
      Width           =   975
   End
   Begin TabDlg.SSTab tbsTabStrip 
      Height          =   4695
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "1. Find Connections"
      TabPicture(0)   =   "frmUMCIonNet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraNet(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraUMCScope"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdFindConnectionsThenUMCs"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "2. Edit/Filter Connections"
      TabPicture(1)   =   "frmUMCIonNet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "lblFilterConnections"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "3. Define UMC's using Connections"
      TabPicture(2)   =   "frmUMCIonNet.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraNet(1)"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdFindConnectionsThenUMCs 
         Caption         =   "&Find Connections then UMC's"
         Height          =   375
         Left            =   8040
         TabIndex        =   54
         ToolTipText     =   "Create Net based on current settings, then Find UMC's"
         Top             =   3840
         Width           =   2415
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
            Left            =   240
            TabIndex        =   2
            Top             =   280
            Width           =   1455
         End
         Begin VB.OptionButton optDefScope 
            Caption         =   "&Current View"
            Height          =   255
            Index           =   1
            Left            =   240
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
         TabIndex        =   56
         Top             =   1260
         Width           =   4575
         Begin VB.TextBox txtNetEditTooDistant 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            TabIndex        =   59
            Text            =   "0.1"
            Top             =   1080
            Width           =   585
         End
         Begin VB.CommandButton cmdRemoveLongConnections 
            Caption         =   "Start"
            Height          =   375
            Left            =   3000
            TabIndex        =   60
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
            TabIndex        =   57
            Top             =   240
            Width           =   4215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Caption         =   "Eliminate connections longer than"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame fraNet 
         Height          =   4095
         Index           =   1
         Left            =   -74880
         TabIndex        =   61
         Top             =   360
         Width           =   10695
         Begin VB.CommandButton cmdFindUMCsUsingNETConnections 
            Caption         =   "Find &UMC's"
            Height          =   375
            Left            =   9240
            TabIndex        =   133
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdAbortProcessing 
            Caption         =   "Abort!"
            Height          =   375
            Left            =   9240
            TabIndex        =   134
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdReportUMC 
            Caption         =   "&Report"
            Height          =   375
            Left            =   9360
            TabIndex        =   132
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdResetToDefaults 
            Caption         =   "Set to Defaults"
            Height          =   375
            Index           =   1
            Left            =   9120
            TabIndex        =   131
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbUMCDrawType 
            Height          =   315
            ItemData        =   "frmUMCIonNet.frx":0054
            Left            =   1680
            List            =   "frmUMCIonNet.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   3675
            Width           =   2175
         End
         Begin VB.TextBox txtInterpolateMaxGapSize 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7200
            TabIndex        =   128
            Text            =   "0"
            Top             =   3675
            Width           =   495
         End
         Begin VB.CheckBox chkInterpolateMissingIons 
            Caption         =   "Interpolate abundances across gaps"
            Height          =   255
            Left            =   4080
            TabIndex        =   126
            Top             =   3405
            Width           =   3015
         End
         Begin VB.Frame Frame2 
            Caption         =   "UMC From Net"
            Height          =   3375
            Left            =   120
            TabIndex        =   62
            Top             =   180
            Width           =   3735
            Begin VB.CheckBox chkUseMostAbuChargeStateStatsForClassStats 
               Caption         =   "Use most abundant charge state group stats for class stats"
               Height          =   530
               Left            =   240
               TabIndex        =   139
               ToolTipText     =   "Make single-member classes from unconnected nodes"
               Top             =   2760
               Width           =   2055
            End
            Begin VB.ComboBox cboChargeStateAbuType 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   137
               Top             =   2360
               Width           =   3255
            End
            Begin VB.ComboBox cmbUMCRepresentative 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   540
               Width           =   3255
            End
            Begin VB.ComboBox cmbUMCMW 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   1740
               Width           =   3255
            End
            Begin VB.ComboBox cmbUMCAbu 
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   66
               Top             =   1140
               Width           =   3255
            End
            Begin VB.CheckBox chkUseUntangledAsSingle 
               Caption         =   "Make single member classes"
               Height          =   530
               Left            =   2400
               TabIndex        =   69
               ToolTipText     =   "Make single-member classes from unconnected nodes"
               Top             =   2760
               Width           =   1215
            End
            Begin VB.Label lblChargeStateAbuType 
               BackStyle       =   0  'Transparent
               Caption         =   "Most Abu Charge State Group Type"
               Height          =   255
               Left            =   240
               TabIndex        =   138
               Top             =   2120
               Width           =   3135
            End
            Begin VB.Label Label4 
               Caption         =   "Class Representative"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   63
               Top             =   300
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "Class Mass"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   67
               Top             =   1500
               Width           =   1335
            End
            Begin VB.Label Label4 
               Caption         =   "Class Abundance"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   65
               Top             =   900
               Width           =   1335
            End
         End
         Begin TabDlg.SSTab tbsUMCRefinementOptions 
            Height          =   3135
            Left            =   3960
            TabIndex        =   70
            Top             =   180
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   5530
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Auto-Refine Options"
            TabPicture(0)   =   "frmUMCIonNet.frx":0058
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraOptionFrame(10)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Split UMC's Options"
            TabPicture(1)   =   "frmUMCIonNet.frx":0074
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraOptionFrame(15)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Adv Class Stats"
            TabPicture(2)   =   "frmUMCIonNet.frx":0090
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fraClassMassTopX"
            Tab(2).Control(1)=   "fraClassAbundanceTopX"
            Tab(2).ControlCount=   2
            Begin VB.Frame fraClassMassTopX 
               Caption         =   "Class Mass Top X"
               Height          =   1215
               Left            =   -74880
               TabIndex        =   119
               Top             =   1800
               Width           =   4095
               Begin VB.TextBox txtClassMassTopXMinAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   121
                  Text            =   "0"
                  Top             =   240
                  Width           =   900
               End
               Begin VB.TextBox txtClassMassTopXMaxAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   123
                  Text            =   "0"
                  ToolTipText     =   "Maximum abundance to include; use 0 to indicate there infinitely large abundance"
                  Top             =   540
                  Width           =   900
               End
               Begin VB.TextBox txtClassMassTopXMinMembers 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   125
                  Text            =   "3"
                  Top             =   840
                  Width           =   900
               End
               Begin VB.Label lblClassMassTopXMinAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   120
                  Top             =   270
                  Width           =   2535
               End
               Begin VB.Label lblClassMassTopXMaxAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Maximum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   122
                  Top             =   560
                  Width           =   2535
               End
               Begin VB.Label lblClassMassTopXMinMembers 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum members to include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   124
                  Top             =   870
                  Width           =   2535
               End
            End
            Begin VB.Frame fraClassAbundanceTopX 
               Caption         =   "Class Abundance Top X"
               Height          =   1215
               Left            =   -74880
               TabIndex        =   112
               Top             =   480
               Width           =   4095
               Begin VB.TextBox txtClassAbuTopXMinMembers 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   118
                  Text            =   "3"
                  Top             =   840
                  Width           =   900
               End
               Begin VB.TextBox txtClassAbuTopXMaxAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   116
                  Text            =   "0"
                  ToolTipText     =   "Maximum abundance to include; use 0 to indicate there infinitely large abundance"
                  Top             =   540
                  Width           =   900
               End
               Begin VB.TextBox txtClassAbuTopXMinAbu 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   114
                  Text            =   "0"
                  Top             =   240
                  Width           =   900
               End
               Begin VB.Label lblClassAbuTopXMinMembers 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum members to include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   117
                  Top             =   870
                  Width           =   2535
               End
               Begin VB.Label lblClassAbuTopXMaxAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Maximum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   115
                  Top             =   560
                  Width           =   2535
               End
               Begin VB.Label lblClassAbuTopXMinAbu 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Minimum Abundance to Include"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   113
                  Top             =   270
                  Width           =   2535
               End
            End
            Begin VB.Frame fraOptionFrame 
               Height          =   2640
               Index           =   15
               Left            =   -74880
               TabIndex        =   93
               Top             =   330
               Width           =   4300
               Begin VB.ComboBox cboSplitUMCsScanGapBehavior 
                  Height          =   315
                  Left            =   1800
                  Style           =   2  'Dropdown List
                  TabIndex        =   111
                  Top             =   2220
                  Width           =   2295
               End
               Begin VB.TextBox txtHoleSize 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   108
                  Text            =   "3"
                  Top             =   1890
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsPeakPickingMinimumWidth 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   105
                  Text            =   "4"
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   102
                  Text            =   "15"
                  Top             =   1230
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsMaximumPeakCount 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   99
                  Text            =   "6"
                  Top             =   900
                  Width           =   495
               End
               Begin VB.TextBox txtSplitUMCsMinimumDifferenceInAvgPpmMass 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   96
                  Text            =   "4"
                  Top             =   570
                  Width           =   495
               End
               Begin VB.CheckBox chkSplitUMCsByExaminingAbundance 
                  Caption         =   "Split UMC's by examining abundance"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   94
                  Top             =   240
                  Width           =   3015
               End
               Begin VB.Label lblSplitUMCsScanGapBehavior 
                  Caption         =   "Scan gap behavior:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   110
                  Top             =   2250
                  Width           =   1620
               End
               Begin VB.Label lblUnits 
                  Caption         =   "scans"
                  Height          =   255
                  Index           =   3
                  Left            =   3480
                  TabIndex        =   109
                  Top             =   1920
                  Width           =   600
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Max size of scan gap in the UMC:"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   107
                  Top             =   1920
                  Width           =   2655
               End
               Begin VB.Label lblSplitUMCsPeakPickingMinimumWidth 
                  Caption         =   "Peak picking minimum width"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   104
                  Top             =   1590
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "scans"
                  Height          =   255
                  Index           =   5
                  Left            =   3480
                  TabIndex        =   106
                  Top             =   1590
                  Width           =   600
               End
               Begin VB.Label lblSplitUMCsPeakDetectIntensityThresholdPercentageOfMax 
                  Caption         =   "Peak picking intensity threshold"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   101
                  Top             =   1260
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "% of max"
                  Height          =   255
                  Index           =   0
                  Left            =   3480
                  TabIndex        =   103
                  Top             =   1260
                  Width           =   700
               End
               Begin VB.Label lblSplitUMCsMaximumPeakCount 
                  Caption         =   "Maximum peak count to split UMC"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   98
                  Top             =   930
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "peaks"
                  Height          =   255
                  Index           =   1
                  Left            =   3480
                  TabIndex        =   100
                  Top             =   930
                  Width           =   600
               End
               Begin VB.Label lblSplitUMCsMinimumDifferenceInAvgPpmMass 
                  Caption         =   "Minimum difference in average mass"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   95
                  Top             =   600
                  Width           =   2700
               End
               Begin VB.Label lblUnits 
                  Caption         =   "ppm"
                  Height          =   255
                  Index           =   2
                  Left            =   3480
                  TabIndex        =   97
                  Top             =   600
                  Width           =   600
               End
            End
            Begin VB.Frame fraOptionFrame 
               Height          =   2700
               Index           =   10
               Left            =   120
               TabIndex        =   71
               Top             =   300
               Width           =   4545
               Begin VB.TextBox txtMaxLengthPctAllScans 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   85
                  Text            =   "15"
                  Top             =   1520
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveMaxLengthPctAllScans 
                  Caption         =   "Remove cls. with length over"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   84
                  Top             =   1520
                  Width           =   2535
               End
               Begin VB.TextBox txtPercentMaxAbuToUseToGaugeLength 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   88
                  Text            =   "33"
                  Top             =   1840
                  Width           =   495
               End
               Begin VB.TextBox txtAutoRefineMinimumMemberCount 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   92
                  Text            =   "3"
                  Top             =   2300
                  Width           =   495
               End
               Begin VB.CheckBox chkRefineUMCLengthByScanRange 
                  Caption         =   "Test UMC length using scan range"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   90
                  ToolTipText     =   "If True, then considers scan range for the length tests; otherwise, considers member count"
                  Top             =   2200
                  Value           =   1  'Checked
                  Width           =   1695
               End
               Begin VB.CheckBox chkRemoveLoAbu 
                  Caption         =   "Remove low intensity classes"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   72
                  Top             =   240
                  Width           =   2550
               End
               Begin VB.TextBox txtLoAbuPct 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   73
                  Text            =   "30"
                  Top             =   240
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveHiAbu 
                  Caption         =   "Remove high intensity classes"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   75
                  Top             =   560
                  Width           =   2550
               End
               Begin VB.TextBox txtHiAbuPct 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   76
                  Text            =   "30"
                  Top             =   560
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveLoCnt 
                  Caption         =   "Remove cls. with less than"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   78
                  Top             =   880
                  Width           =   2295
               End
               Begin VB.TextBox txtLoCnt 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   79
                  Text            =   "3"
                  Top             =   880
                  Width           =   495
               End
               Begin VB.CheckBox chkRemoveHiCnt 
                  Caption         =   "Remove cls. with length over"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   81
                  Top             =   1200
                  Width           =   2535
               End
               Begin VB.TextBox txtHiCnt 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   82
                  Text            =   "500"
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "%"
                  Height          =   255
                  Index           =   5
                  Left            =   3600
                  TabIndex        =   77
                  Top             =   590
                  Width           =   270
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "%"
                  Height          =   255
                  Index           =   4
                  Left            =   3600
                  TabIndex        =   74
                  Top             =   270
                  Width           =   270
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "% all scans"
                  Height          =   255
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   86
                  Top             =   1545
                  Width           =   855
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "%"
                  Height          =   255
                  Index           =   2
                  Left            =   3600
                  TabIndex        =   89
                  Top             =   1870
                  Width           =   285
               End
               Begin VB.Label lblPercentMaxAbuToUseToGaugeLength 
                  Caption         =   "Percent max abu for gauging width"
                  Height          =   240
                  Left            =   360
                  TabIndex        =   87
                  Top             =   1845
                  Width           =   2565
               End
               Begin VB.Label lblAutoRefineMinimumMemberCount 
                  Caption         =   "Minimum member count:"
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   91
                  Top             =   2200
                  Width           =   1125
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "members"
                  Height          =   255
                  Index           =   0
                  Left            =   3600
                  TabIndex        =   80
                  Top             =   915
                  Width           =   900
               End
               Begin VB.Label lblAutoRefineLengthLabel 
                  Caption         =   "members"
                  Height          =   255
                  Index           =   1
                  Left            =   3600
                  TabIndex        =   83
                  Top             =   1230
                  Width           =   900
               End
            End
         End
         Begin VB.Label lblUMCDrawType 
            BackStyle       =   0  'Transparent
            Caption         =   "UMC Draw Type"
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   3705
            Width           =   1455
         End
         Begin VB.Label lblMaxGapSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum size of gap to interpolate:"
            Height          =   255
            Left            =   4080
            TabIndex        =   127
            Top             =   3705
            Width           =   2535
         End
      End
      Begin VB.Frame fraNet 
         Height          =   3255
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   420
         Width           =   8535
         Begin VB.CommandButton cmdResetToOldDefaults 
            Caption         =   "Set to Old Defaults"
            Height          =   250
            Index           =   2
            Left            =   5280
            TabIndex        =   140
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdResetToDefaults 
            Caption         =   "Set to Defaults"
            Height          =   375
            Index           =   0
            Left            =   6960
            TabIndex        =   9
            Top             =   200
            Width           =   1455
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   4
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   2222
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   3
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1860
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   2
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1500
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   1
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1140
            Width           =   855
         End
         Begin VB.ComboBox cmbConstraintUnits 
            Height          =   315
            Index           =   0
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   780
            Width           =   855
         End
         Begin VB.CommandButton cmdFindConnections 
            Caption         =   "&Find Connections Only"
            Height          =   375
            Left            =   6000
            TabIndex        =   52
            ToolTipText     =   "Create Net based on current settings"
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CommandButton cmdAbortFindConnections 
            Caption         =   "Abort!"
            Height          =   375
            Left            =   6750
            TabIndex        =   53
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   6720
            TabIndex        =   48
            Text            =   "0.1"
            Top             =   2222
            Width           =   735
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   6720
            TabIndex        =   40
            Text            =   "0.1"
            Top             =   1860
            Width           =   735
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   32
            Text            =   "0.1"
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   6720
            TabIndex        =   24
            Text            =   "0.1"
            Top             =   1140
            Width           =   735
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   4
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   2222
            Width           =   975
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   3
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1860
            Width           =   975
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   2
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1500
            Width           =   975
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   1
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txtConstraint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6720
            TabIndex        =   16
            Text            =   "0.1"
            Top             =   780
            Width           =   735
         End
         Begin VB.ComboBox cmbConstraint 
            Height          =   315
            Index           =   0
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   780
            Width           =   975
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   4080
            TabIndex        =   45
            Text            =   "1"
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   4080
            TabIndex        =   37
            Text            =   "1"
            Top             =   1860
            Width           =   615
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   4
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2220
            Width           =   2175
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   3
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1860
            Width           =   2175
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   42
            Top             =   2280
            Width           =   700
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   34
            Top             =   1920
            Width           =   700
         End
         Begin VB.TextBox txtNETType 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   8
            Text            =   "1"
            Top             =   300
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtRejectLongConnections 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   51
            Text            =   "1"
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox cmbMetricType 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   4080
            TabIndex        =   29
            Text            =   "1"
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   4080
            TabIndex        =   21
            Text            =   "1"
            Top             =   1140
            Width           =   615
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   2
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1500
            Width           =   2175
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   1
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1140
            Width           =   2175
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   1560
            Width           =   700
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtWeightingFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   4080
            TabIndex        =   13
            Text            =   "1"
            Top             =   780
            Width           =   615
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   0
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   780
            Width           =   2175
         End
         Begin VB.CheckBox chkUse 
            Caption         =   "Use"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Width           =   700
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   10
            Left            =   4800
            TabIndex        =   46
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   9
            Left            =   4800
            TabIndex        =   38
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   8
            Left            =   4800
            TabIndex        =   30
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   7
            Left            =   4800
            TabIndex        =   22
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Constraint"
            Height          =   255
            Index           =   6
            Left            =   4800
            TabIndex        =   14
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Reject connection longer than"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   50
            Top             =   2790
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   44
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   4
            Left            =   3240
            TabIndex        =   36
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblNETType 
            Caption         =   "Net Type"
            Height          =   255
            Left            =   3360
            TabIndex        =   7
            Top             =   320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Metric Type"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   28
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   20
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Wt. Factor"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Label lblFilterConnections 
         Caption         =   $"frmUMCIonNet.frx":00AC
         Height          =   615
         Left            =   -74640
         TabIndex        =   55
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
      TabIndex        =   136
      Top             =   4980
      Width           =   7095
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

Const MAX_NET_SIZE = 1000000

Const NET_ADD_RATE = 5000

Const NET_EDIT_REJECT_LONG = 0

Const HUMCNotUsed As Byte = 0
Const HUMCInUse As Byte = 1
Const HUMCUsed As Byte = 2

Dim CallerID As Long
Dim bLoading As Boolean

Dim DataCnt As Long         'count of isotopic data
                            ' If only finding UMC's on data "in current view", then this value may be
                            ' less than the actual data count in the file

' Unused variable
'''Dim DataWeightFactor() As Double    'weighting factor for each dimension
Dim DataOInd() As Long              'original index in IsoData array
Dim DataVal() As Double             'values to be used in calculations
'this values are dimensioned and weighted to improve calculation speed

Dim ResCnt As Long
Dim ResInd1() As Long
Dim ResInd2() As Long
Dim ResDist() As Double
Dim ResEliminate() As Boolean

Dim MinScan As Long
Dim MaxScan As Long

'following arrays are used to optimize calculations by indexing first dimension
Dim OptIndO() As Long         'indexes in original data arrays
Dim OptValO() As Double       'values in first data dimension used in optimization

Dim MyDef As UMCIonNetDefinition


'settings used to define UMCs from Net
Dim UMCMakeSingleMemberClasses As Boolean
Dim UMCRepresentative As Long

'helper variables used to fill classes
Dim HUMCIsoCnt As Long        'number of Isotopic distributions in a current 2D display
Dim HUMCNetCnt As Long        'number of connections in Net
Dim HUMCIsoUsed() As Byte     'array parallel with IsoData array indicating should isotopic
                              'distribution be included in the current class; it also helps
                              'if unconnected nodes should be made to classes
Dim HUMCEquClsWk() As Long    'array of the same size as IsoData array used to construct each class
                              'this is working array that is never reinitialized to optimize
                              'performance; that means be very careful with it's content
Dim HUMCEquClsCnt As Long     'actual size of the current class
Dim HUMCEquCls() As Long      'array that will hold actual class of equivalency
Dim HUMCNetUsed() As Byte     'array parallel with NetInd arrays indicating if net connection
                              'was already used in classification
Dim DummyInd() As Long        'never to be initialized; used in sort function
                              
Private mAbortProcess As Boolean
Private mCalculating As Boolean
Private mOneSecond As Double

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

Public Sub InitializeUMCSearch()
    
    ' MonroeMod: This code was in Form_Activate
    
On Error GoTo InitializeUMCSearchErrorHandler

    Dim ScanRange As Long
    If bLoading Then
        CallerID = Me.Tag
        lblNetInfo.Caption = GetUMCIonNetInfo(CallerID)
        
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

Public Function StartUMCSearch() As Boolean
    ' This sub should be called after calling InitializeUMCSearch
    ' It is intended to be called during AutoAnalysis
    '
    ' Returns True if Success, False if Error
    
    Dim blnSuccess As Boolean
    Dim blnAbortedProcessDuringAutoAnalysis As Boolean
    
On Error GoTo StartUMCSearchErrorHandler

    ' 1. Find the connections
    tbsTabStrip.Tab = 0
    FindIonNetConnections
    
    ' If auto-analyzing, ignore the mAbortProcess state
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled And mAbortProcess Then
        blnAbortedProcessDuringAutoAnalysis = True
        mAbortProcess = False
    End If
    
    If Not mAbortProcess Then
        ' 2. Find the UMC's
        tbsTabStrip.Tab = 2
        blnSuccess = FormClassesFromNETsWrapper(False)
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
        .AddItem "Actual UMC"
        .AddItem "UMC Full Region"
        .AddItem "UMC Intensity"
    End With
    
    With cboChargeStateAbuType
        .Clear
        .AddItem "Highest Abu Sum"
        .AddItem "Most Abu Member"
        .AddItem "Most Members"
    End With
    
    With cboSplitUMCsScanGapBehavior
        .Clear
        .AddItem "Ignore scan gaps"
        .AddItem "Split if mass difference"
        .AddItem "Always split"
    End With
    
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
    UMCMakeSingleMemberClasses = (chkUseUntangledAsSingle.Value = vbChecked)
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
    cmbMetricType.ListIndex = MyDef.MetricType
Else
    MyDef.MetricType = cmbMetricType.ListIndex
End If
End Sub

Private Sub cmbUMCAbu_Click()
If mCalculating Then
    cmbUMCAbu.ListIndex = UMCDef.ClassAbu
Else
    UMCDef.ClassAbu = cmbUMCAbu.ListIndex
End If
End Sub

Private Sub cmbUMCDrawType_Click()
If mCalculating Then
    cmbUMCDrawType.ListIndex = GelUMCDraw(CallerID).DrawType
Else
    GelUMCDraw(CallerID).DrawType = cmbUMCDrawType.ListIndex
    glbPreferencesExpanded.UMCDrawType = cmbUMCDrawType.ListIndex
End If
End Sub

Private Sub cmbUMCMW_Click()
If mCalculating Then
    cmbUMCMW.ListIndex = UMCDef.ClassMW
Else
    UMCDef.ClassMW = cmbUMCMW.ListIndex
End If
End Sub

Private Sub cmbUMCRepresentative_Click()
If mCalculating Then
    cmbUMCRepresentative.ListIndex = UMCRepresentative
Else
    UMCRepresentative = cmbUMCRepresentative.ListIndex
End If
End Sub

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

Private Sub cmdAbortFindConnections_Click()
    mAbortProcess = True
End Sub

Private Sub cmdAbortProcessing_Click()
    mAbortProcess = True
End Sub

Private Sub cmdClose_Click()
    Dim eResponse As VbMsgBoxResult
    
    If mCalculating And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Calculations are currently in progress.  Abort them and close the window?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Abort Processing")
        If eResponse <> vbYes Then Exit Sub
        mAbortProcess = True
    End If
    
    Unload Me
End Sub

Private Sub cmdFindConnections_Click()
    If mCalculating Then Exit Sub
    FindIonNetConnections
End Sub

Private Sub cmdFindConnectionsThenUMCs_Click()
    If mCalculating Then Exit Sub
    FindIonNetConnections
    If Not mAbortProcess Then
        tbsTabStrip.Tab = 2
        FormClassesFromNETsWrapper False
    End If
End Sub

Private Sub cmdFindUMCsUsingNETConnections_Click()
    If mCalculating Then Exit Sub
    FormClassesFromNETsWrapper True
End Sub

Private Sub cmdRemoveLongConnections_Click()
    If mCalculating Then Exit Sub
    RemoveLongConnections NET_EDIT_REJECT_LONG
End Sub

Private Sub cmdReportUMC_Click()
If mCalculating Then Exit Sub
Me.MousePointer = vbHourglass
ChangeStatus "Generating UMC report..."
Call ReportUMC(CallerID, "UMCIonNet" & vbCrLf & GetUMCIsoDefinitionText(CallerID))
ChangeStatus ""
Me.MousePointer = vbDefault
End Sub

Private Sub cmdResetToDefaults_Click(Index As Integer)
    ResetToDefaults
End Sub

Private Sub cmdResetToOldDefaults_Click(Index As Integer)
    ResetToOldDefaults
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
End Sub

' Note: Status is called by AutoRefineUMCs and by SplitUMCsByAbundance
Public Sub Status(ByVal StatusMsg As String)
    ChangeStatus StatusMsg
End Sub

Private Sub ChangeStatus(ByVal StatusMsg As String)
lblStatus.Caption = StatusMsg
DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
UMCIonNetDef = MyDef
End Sub

Private Sub optDefScope_Click(Index As Integer)
UMCDef.DefScope = Index
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

Private Sub DisplayDynamicUnits()
    Dim intIndex As Integer
    Dim blnShowConstraints As Boolean
    Dim blnMassBasedDataDim As Boolean
    
On Error GoTo DisplayDynamicUnitsErrorHandler

    For intIndex = 0 To cmbData.Count - 1
        Select Case cmbData(intIndex).ListIndex
        Case DATA_MONO_MW, DATA_AVG_MW, DATA_TMA_MW
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

Private Sub FindBestMatches()
Dim i As Long, j As Long                'loop controlers
Dim iOInd As Long, jOInd As Long        'indexes in original Data arrays
Dim BestForI As Long                    'index of best match for index i
Dim ShortestDistance As Double, CurrDistance As Double
Dim bTooFarAway As Boolean
Dim lngTickCountLastUpdate As Long, lngNewTickCount As Long
Dim dtLastUpdateTime As Date

On Error GoTo err_FindBestMatches
mAbortProcess = False
Select Case MyDef.MetricType
Case METRIC_EUCLIDEAN
     Select Case MyDef.NETType
     Case Net_SPIDER_66                                'remember all connections shorter than threshold
         For i = 0 To DataCnt - 1
             iOInd = OptIndO(i)
             lngNewTickCount = GetTickCount()     ' Note that GetTickCount returns a negative number after 24 days of computer Uptime and resets to 0 after 48 days
             If lngNewTickCount - lngTickCountLastUpdate > 250 Or Now - dtLastUpdateTime > mOneSecond Then
                ' Only update 4 times per second
                ChangeStatus "Calculating line " & i & " / " & Trim(DataCnt)
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
                   CurrDistance = MetricEuclid(iOInd, jOInd)
                   If CurrDistance < MyDef.TooDistant Then
                      If Not SubjectToConstraintEuclid(iOInd, jOInd) Then
                         ResCnt = ResCnt + 1
                         'put in results original indexes; always smaller index first
                         If iOInd < jOInd Then
                            ResInd1(ResCnt - 1) = iOInd:   ResInd2(ResCnt - 1) = jOInd
                         Else
                            ResInd1(ResCnt - 1) = jOInd:   ResInd2(ResCnt - 1) = iOInd
                         End If
                         ResDist(ResCnt - 1) = CurrDistance
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
                ChangeStatus "Calculating line " & i & " / " & Trim(DataCnt)
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
                   CurrDistance = MetricEuclid(iOInd, jOInd)
                   If CurrDistance < ShortestDistance Then
                      BestForI = j
                      ShortestDistance = CurrDistance
                   End If
                End If
                j = j + 1
                If j > DataCnt - 1 Then bTooFarAway = True
             Loop
             If ShortestDistance < MyDef.TooDistant Then
                If Not SubjectToConstraintEuclid(iOInd, jOInd) Then
                   ResCnt = ResCnt + 1
                   'put in results original indexes; always smaller first
                   If iOInd < OptIndO(BestForI) Then
                      ResInd1(ResCnt - 1) = iOInd:   ResInd2(ResCnt - 1) = OptIndO(BestForI)
                   Else
                      ResInd1(ResCnt - 1) = OptIndO(BestForI):   ResInd2(ResCnt - 1) = iOInd
                   End If
                   ResDist(ResCnt - 1) = ShortestDistance
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
                ChangeStatus "Calculating line " & i & " / " & Trim(DataCnt)
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
                   CurrDistance = MetricHonduras(iOInd, jOInd)
                   If CurrDistance < MyDef.TooDistant Then
                      If Not SubjectToConstraintHonduras(iOInd, jOInd) Then
                         ResCnt = ResCnt + 1
                         'put in results original indexes; always smaller index first
                         If iOInd < jOInd Then
                            ResInd1(ResCnt - 1) = iOInd:   ResInd2(ResCnt - 1) = jOInd
                         Else
                            ResInd1(ResCnt - 1) = jOInd:   ResInd2(ResCnt - 1) = iOInd
                         End If
                         ResDist(ResCnt - 1) = CurrDistance
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
                ChangeStatus "Calculating line " & i & " / " & Trim(DataCnt)
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
                   CurrDistance = MetricHonduras(iOInd, jOInd)
                   If CurrDistance < ShortestDistance Then
                      BestForI = j
                      ShortestDistance = CurrDistance
                   End If
                End If
                j = j + 1
                If j > DataCnt - 1 Then bTooFarAway = True
             Loop
             If ShortestDistance < MyDef.TooDistant Then
                If Not SubjectToConstraintHonduras(iOInd, jOInd) Then
                   ResCnt = ResCnt + 1
                   'put in results original indexes; always smaller first
                   If iOInd < OptIndO(BestForI) Then
                      ResInd1(ResCnt - 1) = iOInd:   ResInd2(ResCnt - 1) = OptIndO(BestForI)
                   Else
                      ResInd1(ResCnt - 1) = OptIndO(BestForI):   ResInd2(ResCnt - 1) = iOInd
                   End If
                   ResDist(ResCnt - 1) = ShortestDistance
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
                ChangeStatus "Calculating line " & i & " / " & Trim(DataCnt)
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
                   CurrDistance = MetricInfinity(iOInd, jOInd)
                   If CurrDistance < MyDef.TooDistant Then
                      If Not SubjectToConstraintInfinity(iOInd, jOInd) Then
                         ResCnt = ResCnt + 1
                         'put in results original indexes; always smaller index first
                         If iOInd < jOInd Then
                            ResInd1(ResCnt - 1) = iOInd:   ResInd2(ResCnt - 1) = jOInd
                         Else
                            ResInd1(ResCnt - 1) = jOInd:   ResInd2(ResCnt - 1) = iOInd
                         End If
                         ResDist(ResCnt - 1) = CurrDistance
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
                ChangeStatus "Calculating line " & i & " / " & Trim(DataCnt)
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
                   CurrDistance = MetricInfinity(iOInd, jOInd)
                   If CurrDistance < ShortestDistance Then
                      BestForI = j
                      ShortestDistance = CurrDistance
                   End If
                End If
                j = j + 1
                If j > DataCnt - 1 Then bTooFarAway = True
             Loop
             If ShortestDistance < MyDef.TooDistant Then
                If Not SubjectToConstraintInfinity(iOInd, jOInd) Then
                   ResCnt = ResCnt + 1
                   'put in results original indexes; always smaller first
                   If iOInd < OptIndO(BestForI) Then
                      ResInd1(ResCnt - 1) = iOInd:   ResInd2(ResCnt - 1) = OptIndO(BestForI)
                   Else
                      ResInd1(ResCnt - 1) = OptIndO(BestForI):   ResInd2(ResCnt - 1) = iOInd
                   End If
                   ResDist(ResCnt - 1) = ShortestDistance
                End If
             End If
         Next i
     End Select
End Select
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


Private Sub CreateNet()
'------------------------------------------------------------------------------
'fills permanent GelUMCIon structures with indexes;
'before filling permanent GelUMCIon structures results are sorted on Ind1/Ind2;
'this will optimize class creation and reduce the total entropy in the Universe
'------------------------------------------------------------------------------
Dim i As Long
Dim TmpCnt As Long
Dim Ind1() As Long, Ind2() As Long, Dist() As Double, SortInd() As Long

On Error GoTo CreateNetErrorHandler

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
   End If
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

Private Function GetUMCIsoDefinitionText(Ind As Long, Optional blnMultipleLines As Boolean = True) As String
'-----------------------------------------------------------------------
'returns formatted definition of the IonNet for 2D display with index Ind
'-----------------------------------------------------------------------
Dim i As Long
Dim strLineSeparator As String

If blnMultipleLines Then
    strLineSeparator = vbCrLf
Else
    strLineSeparator = "; "
End If

On Error Resume Next
Dim Definition As String
With GelUMCIon(Ind).ThisNetDef
     Select Case .MetricType
     Case METRIC_EUCLIDEAN
          Definition = "Metric type: Euclidean" & strLineSeparator
     Case METRIC_HONDURAS
          Definition = "Metric type: Honduras (a.k.a. Taxicab)" & strLineSeparator
     Case METRIC_INFINITY
          Definition = "Metric type: Infinity" & strLineSeparator
     End Select
     Definition = Definition & "Net type: " & .NETType & strLineSeparator
     Definition = Definition & "Max distance: " & .TooDistant & strLineSeparator
     If .NetActualDim > 0 Then
        If blnMultipleLines Then
            Definition = Definition & "Metric dimensions description:" & strLineSeparator
        Else
            Definition = Definition & "Metric dimensions description; "
        End If
        For i = 0 To .NetDim - 1
            If Not blnMultipleLines Then
                Definition = Definition & "Dimension" & Trim(i + 1) & " = "
            End If
            
            If .MetricData(i).Use Then
               Select Case .MetricData(i).DataType
               Case DATA_MONO_MW
                    Definition = Definition & "Monoisotopic mass; "
               Case DATA_AVG_MW
                    Definition = Definition & "Average mass; "
               Case DATA_TMA_MW
                    Definition = Definition & "The most abundant mass; "
               Case DATA_SCAN
                    Definition = Definition & "Scan; "
               Case DATA_FIT
                    Definition = Definition & "Isotopic fit; "
               Case DATA_MOVERZ
                    Definition = Definition & "m/z; "
               Case DATA_GENERIC_NET
                    Definition = Definition & "Generic NET; "
               Case DATA_CHARGE_STATE
                    Definition = Definition & "Charge state; "
               Case DATA_LOG_ABU
                    Definition = Definition & "Log(Abundance); "
               End Select
               Definition = Definition & "Weight factor: " & .MetricData(i).WeightFactor & "; "
               Definition = Definition & "Constraint: "
               Select Case .MetricData(i).ConstraintType
               Case Net_CT_None
                    Definition = Definition & "none"
               Case Net_CT_LT
                    Definition = Definition & "Distance < " & .MetricData(i).ConstraintValue
               Case Net_CT_GT
                    Definition = Definition & "Distance > " & .MetricData(i).ConstraintValue
               Case Net_CT_EQ
                    Definition = Definition & "Distance equal to " & .MetricData(i).ConstraintValue
               End Select
               
               If .MetricData(i).ConstraintType <> Net_CT_None Then
                    Select Case .MetricData(i).DataType
                    Case DATA_MONO_MW, DATA_AVG_MW, DATA_TMA_MW
                        Definition = Definition & " " & GetMetricDataMassUnits(.MetricData(i).ConstraintUnits)
                    Case Else
                        ' Do not append the units
                    End Select
               End If
            Else
                Definition = Definition & "Unused"
            End If
            Definition = Definition & strLineSeparator
        Next i
     Else
        Definition = Definition & "Metric definition not dimensioned"
     End If
     
     Definition = Definition & vbCrLf
End With
GetUMCIsoDefinitionText = Definition
End Function

Private Function PrepareDataArrays() As Boolean
'------------------------------------------------------------------------
'prepares data arrays and returns True if successful
'------------------------------------------------------------------------
On Error GoTo err_PrepareDataArrays
Dim i As Long, j As Long
Dim strMessage As String
Dim ISInd() As Long         ' In-scope index

ChangeStatus " Preparing arrays..."

GelUMC(CallerID).def.DefScope = UMCDef.DefScope
DataCnt = GetISScope(CallerID, ISInd(), UMCDef.DefScope)

If DataCnt < 2 Then
   strMessage = "Insufficient number of isotopic data points."
   If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
       MsgBox strMessage, vbOKOnly, glFGTU
   Else
       Debug.Assert False
       LogErrors Err.Number, "frmUMCIonNet->PrepareDataArrays, DataCnt < 2"
       AddToAnalysisHistory CallerID, "Error in UMCIonNet Searching: " & strMessage
   End If
   Exit Function
End If
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
      LogErrors Err.Number, "frmUMCIonNet->PrepareDataArrays, MyDef.NetActualDim < 1"
      AddToAnalysisHistory CallerID, "Error in UMCIonNet Searching: " & strMessage
   End If
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
              Case DATA_MONO_MW
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = .IsoData(ISInd(i)).MonoisotopicMW * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_AVG_MW
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = .IsoData(ISInd(i)).AverageMW * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_TMA_MW
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = .IsoData(ISInd(i)).MostAbundantMW * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_SCAN
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = LookupScanNumberRelativeIndex(CallerID, .IsoData(ISInd(i)).ScanNumber) * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_FIT
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = .IsoData(ISInd(i)).Fit * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_MOVERZ
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = .IsoData(ISInd(i)).MZ * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_GENERIC_NET
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = ((.IsoData(ISInd(i)).ScanNumber - MinScan) / (MaxScan - MinScan)) * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_CHARGE_STATE
                  For i = 1 To DataCnt
                      DataOInd(i - 1) = ISInd(i)
                      DataVal(i - 1, j) = .IsoData(ISInd(i)).Charge * MyDef.MetricData(j).WeightFactor
                  Next i
              Case DATA_LOG_ABU
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

Private Function PPMToDaIfNeeded(dblConstraintValue As Double, DimInd As Long, lngDataIndex As Long) As Double
    
    ' If .DataType is a mass type, and if .ContraintUnits is ppm, then convert
    '  dblConstraintValue from Da to ppm, using DataVal(lngDataIndex, DimInd) as
    '  the basis for the conversion
    
    Select Case MyDef.MetricData(DimInd).DataType
    Case DATA_MONO_MW, DATA_AVG_MW, DATA_TMA_MW
        If MyDef.MetricData(DimInd).ConstraintUnits = DATA_UNITS_MASS_PPM Then
            PPMToDaIfNeeded = dblConstraintValue / 1000000# * DataVal(lngDataIndex, DimInd)
        Else
            PPMToDaIfNeeded = dblConstraintValue
        End If
    Case Else
        PPMToDaIfNeeded = dblConstraintValue
    End Select
    
End Function

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

Private Sub RemoveLongConnections(Optional intEditType As Integer = NET_EDIT_REJECT_LONG)
'--------------------------------------------------------------
'does editing of current display GelUMCIon structure
'--------------------------------------------------------------
On Error GoTo RemoveLongConnectionsErrorHandler
Dim strMessage As String
Dim lngConnectionsEliminated As Long

cmdRemoveLongConnections.Visible = False
mCalculating = True

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
cmdRemoveLongConnections.Visible = True

Exit Sub

RemoveLongConnectionsErrorHandler:
Debug.Print "Error in RemoveLongConnections: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->RemoveLongConnections"
Resume Next
End Sub

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
        .MetricData(0).Use = True:  .MetricData(0).DataType = DATA_MONO_MW:   .MetricData(0).WeightFactor = 0.5:   .MetricData(0).ConstraintType = Net_CT_LT:       .MetricData(0).ConstraintValue = 0.025: .MetricData(0).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(1).Use = True:  .MetricData(1).DataType = DATA_AVG_MW:    .MetricData(1).WeightFactor = 0.5:   .MetricData(1).ConstraintType = Net_CT_LT:       .MetricData(1).ConstraintValue = 0.025: .MetricData(1).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(2).Use = True:  .MetricData(2).DataType = DATA_LOG_ABU:   .MetricData(2).WeightFactor = 0.1:   .MetricData(2).ConstraintType = Net_CT_None:     .MetricData(2).ConstraintValue = 0.1:   .MetricData(2).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(3).Use = True:  .MetricData(3).DataType = DATA_SCAN:      .MetricData(3).WeightFactor = 0.01:   .MetricData(3).ConstraintType = Net_CT_None:    .MetricData(3).ConstraintValue = 0.01:  .MetricData(3).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(4).Use = True:  .MetricData(4).DataType = DATA_FIT:       .MetricData(4).WeightFactor = 0.1:    .MetricData(4).ConstraintType = Net_CT_None:    .MetricData(4).ConstraintValue = 0.01:  .MetricData(4).ConstraintUnits = DATA_UNITS_MASS_DA
    End With
End Sub

Private Sub SetUMCDefinition()
'----------------------------------------------------------------------------
'sets definitions for UMC from Net procedure based on some settings of GelUMC().def
'----------------------------------------------------------------------------
On Error GoTo SetUMCDefinitionErrorHandler

With UMCDef
    cmbUMCMW.ListIndex = .ClassMW
    cmbUMCAbu.ListIndex = .ClassAbu
    cboChargeStateAbuType.ListIndex = .ChargeStateStatsRepType
    SetCheckBox chkUseMostAbuChargeStateStatsForClassStats, .UMCClassStatsUseStatsFromMostAbuChargeState
    
    optDefScope(.DefScope).Value = True
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

Private Sub txtSplitUMCsMaximumPeakCount_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsMaximumPeakCount, 2, 100, 6
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

Private Sub FindIonNetConnections()
Dim eResponse As VbMsgBoxResult
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
cmdFindConnections.Visible = False
cmdAbortFindConnections.Visible = True
cmdFindConnectionsThenUMCs.Visible = False
mCalculating = True

If PrepareDataArrays() Then
   If PrepareOptimization() Then
      If ManageResArrays(amtInitialize) Then
         Call FindBestMatches
         
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
cmdFindConnections.Visible = True
cmdFindConnectionsThenUMCs.Visible = True
cmdAbortFindConnections.Visible = False

Exit Sub

FindIonNetConnectionsErrorHandler:
Debug.Print "Error in FindIonNetConnections: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->FindIonNetConnections"
Resume Next

End Sub

Private Function FormClassesFromNETsWrapper(Optional blnShowMessages As Boolean = True) As Boolean
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
            
            If UMCRepresentative < 0 Then
               MsgBox "Class representative type not selected.", vbOKOnly, glFGTU
               cmbUMCRepresentative.SetFocus
               Exit Function
            End If
            If UMCDef.ClassMW < 0 Then
               MsgBox "Class mass type not selected.", vbOKOnly, glFGTU
               cmbUMCMW.SetFocus
               Exit Function
            End If
            If UMCDef.ClassAbu < 0 Then
               MsgBox "Class abundance type not selected.", vbOKOnly, glFGTU
               cmbUMCAbu.SetFocus
               Exit Function
            End If
        End If
        
        ' Call FormClassesFromNets
        blnSuccess = FormClassesFromNets()
        
        If blnSuccess Then
           ChangeStatus "Number of UMCs: " & GelUMC(CallerID).UMCCnt
        Else
           ChangeStatus "Error creating UMCs from Nets."
        End If
    Else
        strMessage = "Net elements not found.  Unable to find UMC's."
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
'returns True if successfull;
'NOTE: if this function is called we have at least one connection in
'GelUMCIon structure
'--------------------------------------------------------------------
Dim bDone As Long
Dim CurrConnInd As Long
Dim CurrInd1 As Long, CurrInd2 As Long
Dim blnUMCIndicesUpdated As Boolean
Dim intScopeUsedForConnections As Integer
Dim i As Long
Dim dblTolPPM As Double
Dim eTolType As glMassToleranceConstants
Dim lngTickCountLastUpdate As Long, lngNewTickCount As Long
Dim dtLastUpdateTime As Date

Dim blnSuccess As Boolean
On Error GoTo err_FormClassesFromNets

mAbortProcess = False
cmdFindUMCsUsingNETConnections.Visible = False
cmdFindConnectionsThenUMCs.Visible = False
cmdAbortProcessing.Visible = True
mCalculating = True

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

ChangeStatus " Initializing UMC structures..."
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
                   ChangeStatus "Building Class: " & GelUMC(CallerID).UMCCnt & " (" & Format(CurrConnInd / HUMCNetCnt * 100, "0.00") & "% completed)"
                   lngTickCountLastUpdate = lngNewTickCount
                   dtLastUpdateTime = Now()
                   If mAbortProcess Then bDone = True
               End If
               
               If CurrConnInd > HUMCNetCnt - 1 Then bDone = True
            End If
         Loop
      End With
      
      'add single member classes if requested
      If UMCMakeSingleMemberClasses Then Call HUMCAddingSingleMemberUMCs
      
      ChangeStatus " Managing UMC structures..."
      If ManageClasses(CallerID, UMCManageConstants.UMCMngTrim) Then
        
        ' Examine GelUMCIon(CallerID).ThisNetDef to determine the appropriate .Tol and .TolType
        '  to record in GelUMC(Callerid).Def
        LookupUMCIonNetMassTolerances dblTolPPM, eTolType, GelUMCIon(CallerID).ThisNetDef, UMC_IONNET_PPM_CONVERSION_MASS
        
        'set various Unique Mass Classes parameters
        With GelUMC(CallerID).def
            .UMCType = glUMC_TYPE_FROM_NET
            .MWField = GelData(CallerID).Preferences.IsoDataField
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
      
        ChangeStatus "Calculating UMC parameters..."
      
        ' Possibly Auto-Refine the UMC's
        blnUMCIndicesUpdated = AutoRefineUMCs(CallerID, Me)
        
        If Not blnUMCIndicesUpdated Then
            ' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
            blnSuccess = UpdateUMCStatArrays(CallerID, False, Me)
        Else
            blnSuccess = True
        End If
      
        If glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCsByAbundance Then
           SplitUMCsByAbundance CallerID, Me, False, True
        End If
      
      Else
         ChangeStatus " Error managing UMC structures."
      End If
   End If
End If

FormClassesFromNetsCleanup:
mCalculating = False
cmdFindUMCsUsingNETConnections.Visible = True
cmdFindConnectionsThenUMCs.Visible = True
cmdAbortProcessing.Visible = False
FormClassesFromNets = blnSuccess
Exit Function

err_FormClassesFromNets:
Debug.Print "Error in FormClassesFromNets: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmUMCIonNet->FormClassesFromNets"
blnSuccess = False
Resume FormClassesFromNetsCleanup

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

ChangeStatus " Adding single-member UMCs..."
With GelUMC(CallerID)
    For i = 1 To lngDataInScope
        lngOriginalIndex = ISInd(i)
        
        If HUMCIsoUsed(lngOriginalIndex) = HUMCNotUsed Then               'not used in any class
           .UMCCnt = .UMCCnt + 1
           .UMCs(.UMCCnt - 1).ClassCount = 1
           With .UMCs(.UMCCnt - 1)
               ReDim .ClassMInd(0):                 ReDim .ClassMType(0)
               .ClassMInd(0) = lngOriginalIndex:    .ClassMType(0) = glIsoType
               .ClassRepInd = lngOriginalIndex:     .ClassRepType = glIsoType
           End With
           Cnt = Cnt + 1
        End If
    Next i

End With
HUMCAddingSingleMemberUMCs = Cnt
Exit Function

err_HUMCAddingSingleMemberUMCs:
Select Case Err.Number
Case 9
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


Private Function BuildCurrentClass() As Boolean
'---------------------------------------------------------------------------------------
'builds class for the current settings in the HUMCEquCls array; returns True on success
'class has to be sorted if more than 2 elements(to preserve scan order)
'---------------------------------------------------------------------------------------
Dim i As Long
Dim BestInd As Long
'Dim MySort As New QSLong
On Error GoTo err_BuildCurrentClass

If HUMCEquClsCnt > 2 Then
   ShellSortLong HUMCEquCls, 0, UBound(HUMCEquCls)
   'If Not MySort.QSAsc(HUMCEquCls(), DummyInd()) Then GoTo err_BuildCurrentClass
   'Set MySort = Nothing
End If

With GelUMC(CallerID)
     .UMCCnt = .UMCCnt + 1
     If .UMCCnt > UBound(.UMCs) + 1 Then             'add room if neccessary
          If Not ManageClasses(CallerID, UMCManageConstants.UMCMngAdd) Then GoTo err_BuildCurrentClass
     End If
     With .UMCs(.UMCCnt - 1)
          ReDim .ClassMInd(HUMCEquClsCnt - 1)
          ReDim .ClassMType(HUMCEquClsCnt - 1)
          For i = 0 To HUMCEquClsCnt - 1
              .ClassCount = .ClassCount + 1
              .ClassMInd(.ClassCount - 1) = HUMCEquCls(i)
              .ClassMType(.ClassCount - 1) = glIsoType
          Next i
          ReDim Preserve .ClassMInd(.ClassCount - 1)
          ReDim Preserve .ClassMType(.ClassCount - 1)
                    
          ' Note: This code has been moved to UMCIonNet.Bas->FindUMCClassRepIndex
          BestInd = FindUMCClassRepIndex(CallerID, GelUMC(CallerID).UMCCnt - 1, CInt(UMCRepresentative))
          
          .ClassRepInd = .ClassMInd(BestInd)
          .ClassRepType = glIsoType
     End With
     BuildCurrentClass = True
     Exit Function
End With

err_BuildCurrentClass:
ChangeStatus " Error building UMC."
End Function


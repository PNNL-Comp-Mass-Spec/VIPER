VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmSearchForNETAdjustmentUMC 
   Caption         =   "Search MT Tag Database For NET Adjustment"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   13110
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTimer 
      Interval        =   500
      Left            =   9360
      Top             =   8040
   End
   Begin VB.Frame fraPlots 
      Height          =   7335
      Left            =   12120
      TabIndex        =   123
      Top             =   600
      Width           =   11415
      Begin VIPER.ctlSpectraPlotter ctlPlotSlopeVsScore 
         Height          =   3135
         Left            =   120
         TabIndex        =   124
         Top             =   360
         Width           =   5415
         _ExtentX        =   9340
         _ExtentY        =   4895
      End
      Begin VIPER.ctlSpectraPlotter ctlPlotMassErrorHistogram 
         Height          =   3255
         Left            =   120
         TabIndex        =   125
         Top             =   3720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5741
      End
      Begin VIPER.ctlSpectraPlotter ctlPlotNETRange 
         Height          =   3135
         Left            =   5760
         TabIndex        =   126
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5530
      End
      Begin VIPER.ctlSpectraPlotter ctlPlotNETErrorHistogram 
         Height          =   3255
         Left            =   5760
         TabIndex        =   127
         Top             =   3720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5741
      End
      Begin VB.Label lblPlotNETErrorHistogram 
         Caption         =   "NET Error Histogram: Red = Current, Blue = Best"
         Height          =   255
         Left            =   5760
         TabIndex        =   132
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Label lblPlotMassErrorHistogram 
         Caption         =   "Mass Error Histogram: Red = Current, Blue = Best"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   3500
         Width           =   4455
      End
      Begin VB.Label lblPlotNETRange 
         Caption         =   "Histogrammed NET values: Red = Current, Blue = Best, Green = DB"
         Height          =   255
         Left            =   5760
         TabIndex        =   130
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label lblPlotSlopeVsScore 
         Caption         =   "Slope vs. NET Match Score"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Frame fraRobustNETOptions 
      Height          =   1280
      Left            =   4440
      TabIndex        =   120
      Top             =   3600
      Width           =   6615
      Begin VB.TextBox txtRobustNETProgress 
         Height          =   855
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   122
         Top             =   240
         Width           =   4095
      End
      Begin VB.CheckBox chkRobustNETEnabled 
         Caption         =   "Robust NET Enabled"
         Height          =   190
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   2175
      End
   End
   Begin TabDlg.SSTab tbsNETOptions 
      Height          =   3405
      Left            =   4440
      TabIndex        =   23
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   6006
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Robust NET Options"
      TabPicture(0)   =   "frmSearchForNETAdjustmentUMC.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblRobustNETPredictedIterationCount"
      Tab(0).Control(1)=   "lblRobustNETCurrentSettings"
      Tab(0).Control(2)=   "fraNETSlopeRange"
      Tab(0).Control(3)=   "fraNETInterceptRange"
      Tab(0).Control(4)=   "fraMassShiftPPMRange"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Addnl NET Options"
      TabPicture(1)   =   "frmSearchForNETAdjustmentUMC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOptionFrame(50)"
      Tab(1).Control(1)=   "fraSlopeAndInterceptWarningTolerances"
      Tab(1).Control(2)=   "fraToleranceRefinementPeakCriteria"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Misc. Adv. Options"
      TabPicture(2)   =   "frmSearchForNETAdjustmentUMC.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraNetAdjLockers"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraMiscellaneousAdvanced"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraChargeStateForUMCs"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame fraChargeStateForUMCs 
         Caption         =   "Charge State for UMC Selection"
         Height          =   855
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   5895
         Begin VB.CheckBox chkCS 
            Caption         =   "any charge state"
            Height          =   375
            Index           =   7
            Left            =   4440
            TabIndex        =   80
            Top             =   410
            Width           =   1215
         End
         Begin VB.CheckBox chkCS 
            Caption         =   ">=7"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   79
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "6"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   78
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "5"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   77
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "4"
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   76
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   75
            Top             =   480
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   74
            Top             =   480
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   73
            Top             =   480
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Consider peaks with charge states:"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   72
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraMiscellaneousAdvanced 
         Caption         =   "Miscellaneous"
         Height          =   1890
         Left            =   120
         TabIndex        =   81
         Top             =   1440
         Width           =   3615
         Begin VB.TextBox txtNetAdjMinHighDiscriminantScore 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            TabIndex        =   87
            Text            =   "0.5"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNetAdjMinHighNormalizedScore 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            TabIndex        =   85
            Text            =   "2.5"
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox chkEliminateConflictingIDs 
            Caption         =   "Do not use peaks pointing to multiple IDs on NET distance of more than"
            Height          =   615
            Left            =   240
            TabIndex        =   82
            Top             =   240
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.TextBox txtMultiIDMaxNETDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            TabIndex        =   83
            Text            =   "0.1"
            Top             =   300
            Width           =   615
         End
         Begin VB.CheckBox chkUseN15AMTMasses 
            Caption         =   "Use N15 MT Masses"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum MT Discriminant Score"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   1245
            Width           =   2505
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum MT XCorr"
            Height          =   255
            Index           =   133
            Left            =   120
            TabIndex        =   84
            Top             =   885
            Width           =   1785
         End
      End
      Begin VB.Frame fraNetAdjLockers 
         Caption         =   "Not Used: Net Adj Lockers"
         Height          =   1890
         Left            =   3960
         TabIndex        =   89
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CheckBox chkNetAdjUseLockers 
            Caption         =   "Use Lockers"
            Height          =   255
            Left            =   180
            TabIndex        =   90
            Top             =   300
            Visible         =   0   'False
            Width           =   1700
         End
         Begin VB.CheckBox chkNetAdjUseOldIfFailure 
            Caption         =   "Use old algorithm if failure"
            Height          =   375
            Left            =   180
            TabIndex        =   91
            Top             =   600
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox txtNetAdjMinLockerMatchCount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   93
            Text            =   "3"
            Top             =   1160
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum match count"
            Height          =   360
            Index           =   111
            Left            =   120
            TabIndex        =   92
            Top             =   1095
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Initial NET Mapping"
         Height          =   1815
         Index           =   50
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   2895
         Begin VB.TextBox txtNetAdjInitialIntercept 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   52
            Text            =   "0"
            Top             =   660
            Width           =   855
         End
         Begin VB.TextBox txtNetAdjInitialSlope 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   50
            Text            =   "0.0003"
            Top             =   300
            Width           =   855
         End
         Begin VB.Label lblNetAdjInitialNETStats 
            Caption         =   "Initial NET Stats"
            Height          =   615
            Left            =   120
            TabIndex        =   53
            Top             =   1005
            Width           =   2565
         End
         Begin VB.Label lblNetAdjInitialIntercept 
            Caption         =   "Initial Intercept"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   680
            Width           =   1185
         End
         Begin VB.Label lblNetAdjInitialSlope 
            Caption         =   "Initial Slope"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   320
            Width           =   1140
         End
      End
      Begin VB.Frame fraSlopeAndInterceptWarningTolerances 
         Caption         =   "Warning Tolerances"
         Height          =   1500
         Left            =   -71760
         TabIndex        =   54
         Top             =   480
         Width           =   2895
         Begin VB.TextBox txtNetAdjWarningTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   56
            Text            =   "0.00005"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtNetAdjWarningTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   58
            Text            =   "0.01"
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtNetAdjWarningTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   60
            Text            =   "-1"
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtNetAdjWarningTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   62
            Text            =   "1"
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum Slope"
            Height          =   255
            Index           =   44
            Left            =   120
            TabIndex        =   55
            Top             =   260
            Width           =   1500
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum Slope"
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   57
            Top             =   560
            Width           =   1500
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum Intercept"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   1500
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum Intercept"
            Height          =   255
            Index           =   41
            Left            =   120
            TabIndex        =   61
            Top             =   1160
            Width           =   1500
         End
      End
      Begin VB.Frame fraToleranceRefinementPeakCriteria 
         Caption         =   "Criteria To Use Peak For Refinement"
         Height          =   1280
         Left            =   -71760
         TabIndex        =   63
         Top             =   2040
         Width           =   3255
         Begin VB.TextBox txtToleranceRefinementPercentageOfMaxForWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   68
            Text            =   "60"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtToleranceRefinementMinimumPeakHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   65
            Text            =   "25"
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   70
            Text            =   "2.5"
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Pct of Max for Finding Width"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   67
            Top             =   640
            Width           =   2055
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum Height"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   64
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label lblUnits 
            Caption         =   "counts/bin"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   66
            Top             =   330
            Width           =   840
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum S/N for Low Abu"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   69
            Top             =   940
            Width           =   2055
         End
      End
      Begin VB.Frame fraMassShiftPPMRange 
         Caption         =   "PPM Mass Shift Range"
         Height          =   975
         Left            =   -74880
         TabIndex        =   40
         Top             =   1800
         Width           =   3015
         Begin VB.TextBox txtRobustNETMassShiftPPMStart 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   42
            Text            =   "-6"
            Top             =   240
            Width           =   700
         End
         Begin VB.TextBox txtRobustNETMassShiftPPMEnd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   44
            Text            =   "6"
            Top             =   540
            Width           =   700
         End
         Begin VB.TextBox txtRobustNETMassShiftPPMIncrement 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   46
            Text            =   "6"
            Top             =   540
            Width           =   900
         End
         Begin VB.Label lblDescription 
            Caption         =   "Start"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   41
            Top             =   260
            Width           =   705
         End
         Begin VB.Label lblDescription 
            Caption         =   "End"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Top             =   560
            Width           =   705
         End
         Begin VB.Label lblRobustNETIterate 
            Caption         =   "Increment"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   45
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.Frame fraNETInterceptRange 
         Caption         =   "NET Intercept Range"
         Height          =   1335
         Left            =   -71640
         TabIndex        =   33
         Top             =   420
         Width           =   1935
         Begin VB.TextBox txtRobustNETInterceptStart 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   35
            Text            =   "-0.5"
            Top             =   240
            Width           =   700
         End
         Begin VB.TextBox txtRobustNETInterceptEnd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   37
            Text            =   "0.3"
            Top             =   540
            Width           =   700
         End
         Begin VB.TextBox txtRobustNETInterceptIncrement 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   39
            Text            =   "0.2"
            Top             =   840
            Width           =   700
         End
         Begin VB.Label lblRobustNETIterate 
            Caption         =   "Increment"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   860
            Width           =   825
         End
         Begin VB.Label lblDescription 
            Caption         =   "End"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   36
            Top             =   560
            Width           =   705
         End
         Begin VB.Label lblDescription 
            Caption         =   "Start"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   260
            Width           =   705
         End
      End
      Begin VB.Frame fraNETSlopeRange 
         Caption         =   "NET Slope Range"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   24
         Top             =   420
         Width           =   3015
         Begin VB.OptionButton optRobustNETSlopeIncrement 
            Caption         =   "percent"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   32
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton optRobustNETSlopeIncrement 
            Caption         =   "absolute"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   31
            Top             =   720
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox txtRobustNETSlopeIncrement 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   30
            Text            =   "75"
            Top             =   840
            Width           =   900
         End
         Begin VB.TextBox txtRobustNETSlopeEnd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   28
            Text            =   "0.005"
            Top             =   540
            Width           =   900
         End
         Begin VB.TextBox txtRobustNETSlopeStart 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            TabIndex        =   26
            Text            =   "0.00001"
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblRobustNETIterate 
            Caption         =   "Increment"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   860
            Width           =   825
         End
         Begin VB.Label lblDescription 
            Caption         =   "End"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   560
            Width           =   705
         End
         Begin VB.Label lblDescription 
            Caption         =   "Start"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   260
            Width           =   705
         End
      End
      Begin VB.Label lblRobustNETCurrentSettings 
         Height          =   615
         Left            =   -71640
         TabIndex        =   128
         Top             =   2640
         Width           =   3105
      End
      Begin VB.Label lblRobustNETPredictedIterationCount 
         Caption         =   "Predicted Iteration Count:"
         Height          =   255
         Left            =   -71640
         TabIndex        =   47
         Top             =   1920
         Width           =   3105
      End
   End
   Begin VB.Frame fraIte 
      Caption         =   "Iteration"
      Height          =   3015
      Left            =   120
      TabIndex        =   94
      Top             =   4920
      Width           =   9615
      Begin RichTextLib.RichTextBox rtbIteReport 
         Height          =   2205
         Left            =   3960
         TabIndex        =   110
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3889
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmSearchForNETAdjustmentUMC.frx":0054
      End
      Begin VB.CommandButton cmdUseDefaults 
         Caption         =   "Use Defauts"
         Height          =   315
         Left            =   8040
         TabIndex        =   116
         ToolTipText     =   "Start calculations"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtNetAdjMinIDCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   109
         Text            =   "75"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox chkNetAdjAutoIncrementUMCTopAbuPct 
         Caption         =   "Auto-increment high abu UMC's top percent"
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   2380
         Width           =   3495
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   315
         Left            =   6120
         TabIndex        =   113
         ToolTipText     =   "Reset to generic formula and NET tolerance of 0.2"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Iterating"
         Height          =   315
         Left            =   4080
         TabIndex        =   111
         ToolTipText     =   "Start calculations"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   315
         Left            =   8040
         TabIndex        =   115
         ToolTipText     =   "Stops iterations and dumps current results"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   4080
         TabIndex        =   112
         ToolTipText     =   "Stops iterations and dumps current results"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when change less than (or Iterations > max)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   99
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox chkAcceptLastIteration 
         Caption         =   "Accept last iteration as NET adjustment"
         Height          =   255
         Left            =   240
         TabIndex        =   106
         Top             =   2100
         Width           =   3255
      End
      Begin VB.TextBox txtIteNETDec 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   105
         Text            =   "0.025"
         Top             =   1815
         Width           =   615
      End
      Begin VB.TextBox txtIteMWDec 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   103
         Text            =   "2.5"
         Top             =   1485
         Width           =   615
      End
      Begin VB.TextBox txtIteStopVal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   101
         Text            =   "5"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when number of IDs goes under"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   98
         Top             =   960
         Width           =   3015
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when NET tol. goes under"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   97
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop when MW tol. goes under"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   96
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton optIteStop 
         Caption         =   "Stop after iteration number"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.CheckBox chkDecNET 
         Caption         =   "Decrease NET tolerance  by"
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   1845
         Width           =   2415
      End
      Begin VB.CheckBox chkDecMW 
         Caption         =   "Decrease MW tolerance by"
         Height          =   255
         Left            =   240
         TabIndex        =   102
         Top             =   1515
         Width           =   2535
      End
      Begin MSComctlLib.ProgressBar pbarRobustNET 
         Height          =   315
         Left            =   5640
         TabIndex        =   114
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Max             =   10
         Scrolling       =   1
      End
      Begin VB.Label lblDescription 
         Caption         =   "Minimum matching UMC count"
         Height          =   255
         Index           =   47
         Left            =   480
         TabIndex        =   108
         Top             =   2670
         Width           =   2445
      End
      Begin VB.Label lblDescription 
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
         Index           =   17
         Left            =   3030
         TabIndex        =   100
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
         ItemData        =   "frmSearchForNETAdjustmentUMC.frx":00D6
         Left            =   120
         List            =   "frmSearchForNETAdjustmentUMC.frx":00D8
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   4000
      End
      Begin VB.TextBox txtMaxUMCScansPct 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
      Begin VB.Label lblDescription 
         Caption         =   "Maximum percentage of total scans in UMC"
         Enabled         =   0   'False
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   5
         Top             =   900
         Width           =   3135
      End
      Begin VB.Label lblDescription 
         Caption         =   "Minimum scan range for UMC"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   3
         Top             =   580
         Width           =   3135
      End
      Begin VB.Label lblDescription 
         Caption         =   "%"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   9
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label lblDescription 
         Caption         =   "Minimum number of peaks in UMC to use"
         Height          =   255
         Index           =   11
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
         Begin VB.Label lblDescription 
            Caption         =   "Tolerance"
            Height          =   255
            Index           =   8
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
      Begin VB.Label lblDescription 
         Caption         =   "NET Calculation Formula (FN, MinFN, MaxFN)"
         Height          =   255
         Index           =   14
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
      Left            =   4320
      TabIndex        =   119
      Top             =   8040
      Width           =   4215
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   420
      Left            =   1200
      TabIndex        =   118
      Top             =   7980
      Width           =   3135
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   117
      ToolTipText     =   "Status of the MT Tag database"
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
      Begin VB.Menu mnuFSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLogCalculations 
         Caption         =   "Log Calculations"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMT 
      Caption         =   "&MT Tags"
      Begin VB.Menu mnuMTLoad 
         Caption         =   "&Load MT Tags"
      End
      Begin VB.Menu mnuLoadLegacy 
         Caption         =   "Load L&egacy MT Tags"
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
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewTogglePlotPosition 
         Caption         =   "&Toggle Plot Position"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewAutoZoomPlots 
         Caption         =   "&Auto Zoom Plots"
         Checked         =   -1  'True
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
'This is how it works; find all matching MT tags then look to
'reduce that number by selecting pairs that will yield best match
'(eliminate all except first instance of each peptide; use higher
'intensity peaks). Look for best Slope and Intercept for IDs by
'least square method.
'-----------------------------------------------------------------
Option Explicit

Private Const MAX_ID_CNT As Long = 100000            'maximum number of IDs
Private Const MAX_ROBUST_NET_ITERATION_COUNT As Long = 100000

''Private Const STATE_NET_TOO_DISTANT As Long = 1024
''Private Const STATE_TOO_LONG_ELUTION As Long = 2048
Private Const STATE_BAD_NET As Long = 4096
Private Const STATE_OUTSCORED As Long = 8192
''Private Const STATE_ID_NETS_TOO_DISTANT As Long = 16384

Private Const ITERATION_STOP_NUMBER As Long = 0
Private Const ITERATION_STOP_MW_TOL_LIMIT As Long = 1
Private Const ITERATION_STOP_NET_TOL_LIMIT As Long = 2
Private Const ITERATION_STOP_ID_LIMIT As Long = 3
Private Const ITERATION_STOP_CHANGE As Long = 4

Private Const NET_RESOLUTION As Single = 0.01                'theoretical NET resolution; this defines the absolute minimum value that can be used for the NET Tolerance
''Private Const NET_RESOLUTION_SIMULATED_ANNEALING As Single = 0.001                'theoretical NET resolution during simulated annealing

Private Const NET_CHANGE_PCT As Single = 0.25                'used in iteration that stops when change is insignificant

Private Const NET_HISTOGRAM_BIN_SIZE As Single = 0.05
Private Const NET_HISTOGRAM_DIGITS_AFTER_DECIMAL As Integer = 2

Private Const MASS_ERROR_HISTOGRAM_BIN_SIZE As Single = 0.5
Private Const MASS_ERROR_HISTOGRAM_DIGITS_AFTER_DECIMAL As Integer = 1

Private Const NET_ERROR_HISTOGRAM_BIN_SIZE As Single = 0.005
Private Const NET_ERROR_HISTOGRAM_DIGITS_AFTER_DECIMAL As Integer = 3


Private Const NET_TOL_DIGITS_PRECISION As Integer = 5

Private Const CMD_CAPTION_PAUSE As String = "Pause"
Private Const CMD_CAPTION_CONTINUE As String = "Continue"

Private Const PLOT_FRAME_POSITION_VISIBLE As Long = 120
Private Const PLOT_FRAME_POSITION_OFFSET As Long = 11700

Private Enum naswNetAdjustSlopeWarningConstants
    naswSlopeMinimum = 0
    naswSlopeMaximum = 1
    naswInterceptMinimum = 2
    naswInterceptMaximum = 3
End Enum

Private Type udtIterationStatsType
    MassShiftPPM As Single
    slope As Double
    intercept As Double
    Deviation As Double
    
    InitialSlope As Double
    InitialIntercept As Double
    
    UMCTopAbuPct As Integer
    IterationCount As Integer
    FinalNETTol As Double
    FinalMWTol As Double
    FinalMWTolType As glMassToleranceConstants
    
    UMCSegmentCntWithLowUMCCnt As Long
    UMCCntAddedSinceLowSegmentCount As Long
     
    IDMatchCount As Long
    
    UMCCountSearched As Long
    UMCMatchCount As Long
    
    PointCountSearched As Long
    PointMatchCount As Long
    
    RobustNET As Boolean
    UseInternalStandards As Boolean
    Valid As Boolean

    NETRangeHistogramData As udtBinnedDataType
    MassErrorHistogramData As udtBinnedDataType
    NETErrorHistogramData As udtBinnedDataType
    
    CorrelationCoeff As Single
    MassErrorPPM As Single                  ' This comes from the peak identified in .MassErrorHistogramData
    MassErrorPeakHeight As Long
    
    NETMatchScore As Double

End Type

' GRID stands for Group of Longs, holding IDs
Private GRID() As GR        ' GRID array is parallel with AMT array it is indexed the same way and it's
                            ' members are indexes in ID belonging to AMT (in other words have same ID as MT tags index in AMT arrays)
                            ' Each row in GRID contains the variable .Members(), which holds indices in the ID() array that matched the given AMT

' The ID() arrays are 0-based
Private IDCnt As Long           'identification count
Private ID() As Long                'index in AMT array or UMCInternalStandards.InternalStandards()
Private IDMatchingUMCInd() As Long  ' Index of the UMC that matched the AMT
Private IDMatchAbu() As Double      ' Abundance of the UMC that matched the AMT
Private IDIsoFit() As Single        ' Isotopic fit of the UMC that matched the AMT
Private IDScan() As Long        'this is redundant but will make life easier
Private IDState() As Long       'used to clean IDs from duplicates and bads
Private IDsAreInternalStandards As Boolean

Private mIterationCount As Long  'iteration count

Private UseUMC() As Boolean     'marks which UMCs should be used in this search
Private mUMCCntAddedSinceLowSegmentCount As Long
Private mUMCSegmentCntWithLowUMCCnt As Long

' The following holds the details of the UMCs being searched
Private PeakCnt As Long             ' Number of UMC's searched
Private PeakCntWithMatches As Long  ' Number of UMC's that matched an AMT
Private PeakUMCInd() As Long        ' Index of the UMC
Private PeakMW() As Double          ' Nominally the UMC's class mass, but will be adjusted if mMassShiftPPM is non-zero
Private PeakScan() As Long          ' Scan of the UMC's class rep
Private PeakIsoFit() As Single      ' Isotopic fit of the UMC's class rep

Private UMCTotalPoints As Long      ' The total number of points in the UMCs being searched
Private UMCTotalPointsMatching As Long

'in this case CallerID is a public property
Public CallerID As Long

Private bLoading As Boolean

Private ScanMin As Long                 'first scan number
Private ScanMax As Long                 'last scan number
Private ScanRange As Long               'last-first+1

Private AdjSlp As Double                'New slope resulting from GANET adjustments
Private AdjInt As Double                'New intercept resulting from GANET adjustment
Private AdjAvD As Double

Private EditGANETSPName As String
Private NETExprEva As ExprEvaluator     'expression evaluator for NET
Private VarVals() As Long             'variable for expression evaluator

Private bStop As Boolean
Private bPause As Boolean
Private mRequestUpdateRobustNETIterationCount As Boolean
Private mFormInitialized As Boolean
Private mMassCalibrationInProgress As Boolean

Private mUsingDefaultGANET As Boolean
Private mStopChangeTestResult As Double
Private mUsingInternalStandards As Boolean
Private mMTMinimumHighNormalizedScore As Single
Private mMTMinimumHighDiscriminantScore As Single

Private mNETStatsHistoryCount As Long
Private mNETStatsHistoryBestIndex As Long
Private mNETStatsHistory() As udtIterationStatsType

Private mRobustNETInProgress As Boolean
Private mMassShiftPPM As Single

Private mNegativeSlopeComputed As Boolean
Private mNetAdjFailed As Boolean
Private mEffectiveMinIDCount As Long

Private mCurrentIterationStats As udtIterationStatsType     ' Holds the best results for the current iteration

Private mInternalStandardsNETHistogram As udtBinnedDataType
Private mPMTDataNETHistogram As udtBinnedDataType

Private mInternalStandardIndexPointers() As Long             ' Pointer to entry in UMCInternalStandards.InternalStandards()
Private objInternalStandardSearchUtil As MWUtil

Private mUMCClsStats() As Double

Private objHistogram As New clsHistogram
Private objCorrelate As New clsCorrelation
'

Private Sub AdjustMassErrorHistogramMasses(ByRef udtMassErrorHistogramData As udtBinnedDataType, dblMassCalPPMCorrection As Double)

    With udtMassErrorHistogramData
        .StartBin = .StartBin + Round(dblMassCalPPMCorrection, MASS_ERROR_HISTOGRAM_DIGITS_AFTER_DECIMAL)
    End With
End Sub

Private Sub AppendNETStatsHistoryToAnalysisHistory(ByRef udtNETStatsHistory() As udtIterationStatsType, ByVal lngNETStatsHistoryCount As Long, ByVal lngNETStatsHistoryBestIndex As Long, Optional intTopEntriesToSave As Integer = 5)

    Dim lngCount As Long
    Dim dblScores() As Double
    Dim lngIndices() As Long
    
    Dim objQSDouble As New QSDouble
    Dim lngIndex As Long
    Dim lngIndexStart As Long
    
    Dim blnMatchFound As Boolean
    
    If lngNETStatsHistoryCount <= 0 Then
        ' Nothing to save
        Debug.Assert False
    Else
        lngCount = 0
        ReDim dblScores(lngNETStatsHistoryCount - 1)
        ReDim lngIndices(lngNETStatsHistoryCount - 1)
        
        ' Copy the valid scores from udtNETStatsHistory() into dblScores and lngIndices()
        For lngIndex = 0 To lngNETStatsHistoryCount - 1
            With udtNETStatsHistory(lngIndex)
                If .Valid Then
                    dblScores(lngCount) = .NETMatchScore
                    lngIndices(lngCount) = lngIndex
                    lngCount = lngCount + 1
                End If
            End With
        Next lngIndex
        
        If lngCount > 0 Then
            ReDim Preserve dblScores(lngCount - 1)
            ReDim Preserve lngIndices(lngCount - 1)
            
            ' Sort ascending
            If objQSDouble.QSAsc(dblScores, lngIndices) Then
            
                ' Write out the last intTopEntriesToSave entries
                ' If lngCount < intTopEntriesToSave then write out all the entries
                
                If lngCount <= intTopEntriesToSave Then
                    lngIndexStart = 0
                Else
                    lngIndexStart = lngCount - intTopEntriesToSave
                End If
                
                ' Make sure the final entry is lngNETStatsHistoryBestIndex
                If lngIndices(lngCount - 1) <> lngNETStatsHistoryBestIndex And lngNETStatsHistoryBestIndex >= 0 Then
                    ' Look for lngNETStatsHistoryBestIndex in lngIndices()
                    blnMatchFound = False
                    For lngIndex = lngIndexStart To lngCount - 2
                        If lngIndices(lngIndex) = lngNETStatsHistoryBestIndex Then
                            ' Match found
                            lngIndices(lngIndex) = lngIndices(lngCount - 1)
                            lngIndices(lngCount - 1) = lngNETStatsHistoryBestIndex
                            blnMatchFound = True
                            Exit For
                        End If
                    Next lngIndex
                    
                    If Not blnMatchFound Then
                        ' Append the entry to lngIndices()
                        
                        ReDim Preserve dblScores(lngCount)
                        ReDim Preserve lngIndices(lngCount)
                        
                        dblScores(lngCount) = udtNETStatsHistory(lngNETStatsHistoryBestIndex).NETMatchScore
                        lngIndices(lngCount) = lngNETStatsHistoryBestIndex
                        
                        lngCount = lngCount + 1

                    End If
                End If
                
                For lngIndex = lngIndexStart To lngCount - 1
                    If lngIndex = lngCount - 1 Then
                        UpdateAnalysisHistory udtNETStatsHistory(lngIndices(lngIndex)), False, True
                    Else
                        UpdateAnalysisHistory udtNETStatsHistory(lngIndices(lngIndex)), False, False
                    End If
                Next lngIndex
            
            End If
            
        End If
    End If
    
End Sub

Private Sub AppendToIterationReport(ByVal strMessage As String)

    If Len(strMessage) >= 2 Then
        If Not Right(strMessage, 2) = vbCrLf Then
            strMessage = strMessage & vbCrLf
        End If
    End If
    
    rtbIteReport.SelStart = Len(rtbIteReport.Text)
    rtbIteReport.SelText = strMessage

End Sub

Public Sub CalculateNETAdjustmentStart()

    ' Make sure the settings are sync'd with UMCNetAdjDef
    Call PickParameters
    
    mRobustNETInProgress = False
    tmrTimer.Enabled = False
    
    ' Clear any defined Custom NET values
    CustomNETsClear CallerID
    
    If UMCNetAdjDef.UseRobustNETAdjustment Then
        UMCNetAdjDef.RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETIterative
        CalculateRobustNET
    Else
        CalculateIteration 2
    End If
    
    
    tmrTimer.Enabled = True
End Sub

Private Sub CalculateIteration(intPlotUpdateInterval As Integer)
    ' Note: This sub should be called from CalculateNETAdjustmentStart or CalculateRobustNETIterative
    ' Note: The format of this sub is similar to that of CalculateRobustNETIterative
    
    Dim strMessage As String
    Dim blnSuccess As Boolean

On Error GoTo CalculateIterationErrorHandler

    Me.MousePointer = vbHourglass
    DoEvents
    
    If Not mRobustNETInProgress Then
        ReDim mNETStatsHistory(10)
        
        mNETStatsHistoryCount = 0
        mNETStatsHistoryBestIndex = -1
    
        UpdateNETHistograms
    End If
    
''    ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''    If UMCNetAdjDef.UseNetAdjLockers Then
''        ' Note that the following function will call CalculateIterationAutoIncrementTopAbuPct
''        blnSuccess = CalculateUsingInternalStandards(intPlotUpdateInterval)
''        If Not blnSuccess Then
''            If Not UMCNetAdjDef.UseOldNetAdjIfFailure Then
''                strMessage = "NET Adjustment using lockers (internal standards) failed.  Since UseOldNetAdjIfFailure = False, will use the default Slope and Intercept."
''
''                If mRobustNETInProgress Then
''                    ' Do not log the error
''                Else
''                    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''                        MsgBox strMessage, vbExclamation Or vbOKOnly, "NET Adjustment Failure"
''                    Else
''                        AddToAnalysisHistory CallerID, strMessage
''                    End If
''                End If
''
''                ResetSlopeAndInterceptToDefault
''
''                ' Set this to True so that CalculateIterationAutoIncrementTopAbuPct or CalculateIterationWork isn't called below
''                blnSuccess = True
''            End If
''        End If
''    Else
''        blnSuccess = False
''    End If

    blnSuccess = False
    If Not blnSuccess Then
        If cChkBox(chkUMCUseTopAbu) And glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentAutoIncrementUMCTopAbuPct Then
            CalculateIterationAutoIncrementTopAbuPct intPlotUpdateInterval
        Else
            ' Call CalculateIterationWork() just once
            mIterationCount = 0
            CalculateIterationWork mIterationCount, False, intPlotUpdateInterval
        End If
    End If
    
    ValidatePositiveSlope Not mRobustNETInProgress
    
    Me.MousePointer = vbDefault
    Exit Sub

CalculateIterationErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->CalculateIteration"
    Me.MousePointer = vbDefault
    
End Sub

Private Sub CalculateIterationAutoIncrementTopAbuPct(intPlotUpdateInterval As Integer)
    
    Dim blnDone As Boolean
    Dim blnUseAbbreviatedFormat As Boolean
    Dim lngOriginalNetAdjMinIDCount As Long
    Dim blnStartedWithGenericNETEquation As Boolean
    
    Dim udtPreviousIterationStats As udtIterationStatsType
    
    Dim strMessage As String
    Dim strNetAdjFailureReason As String
    
    ' Call CalculateIterationWork()
    ' If number of ID's is less than the minimum, then
    '   increment txtUMCAbuTopPct and try again
    
    mNetAdjFailed = False
    udtPreviousIterationStats.Valid = False
    mCurrentIterationStats.Valid = False
    
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
        If .NETAdjustmentUMCTopAbuPctMax < 5 Then .NETAdjustmentUMCTopAbuPctMax = 100
    End With
    
    If UMCNetAdjDef.TopAbuPct < 1 Then
        UMCNetAdjDef.TopAbuPct = 1
        txtUMCAbuTopPct = UMCNetAdjDef.TopAbuPct
    End If
    
    ' Note: On the first call to CalculateIterationWork, we want blnUseAbbreviatedFormat = False
    '       On subsequent calls, we want it True
    blnUseAbbreviatedFormat = False
    
    lngOriginalNetAdjMinIDCount = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount
    mEffectiveMinIDCount = lngOriginalNetAdjMinIDCount
    
    blnDone = False
    Do
        mIterationCount = 0
        CalculateIterationWork mIterationCount, blnUseAbbreviatedFormat, intPlotUpdateInterval
        
        blnUseAbbreviatedFormat = True
        With glbPreferencesExpanded.AutoAnalysisOptions
            If (IDCnt < .NETAdjustmentMinIDCount Or mIterationCount < .NETAdjustmentMinIterationCount Or AdjSlp <= 0) And Not bStop Then
                    
                blnDone = False
                SetNetAdjFailed
                If IDCnt < .NETAdjustmentMinIDCount Then
                    ' Not enough ID's
                    strNetAdjFailureReason = "Not enough UMC's matched MT tags in the database"
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
                    ElseIf UMCNetAdjDef.IterationStopType = ITERATION_STOP_NUMBER And mIterationCount >= UMCNetAdjDef.IterationStopValue Then
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
                        
                        If Not mRobustNETInProgress Then
                            AddToAnalysisHistory CallerID, "NET Adjustment: " & strNetAdjFailureReason & "; incremented high abundance UMC's top percent and will repeat search; increment = " & Trim(.NETAdjustmentUMCTopAbuPctIncrement) & "%; new top abundance percent = " & Trim(UMCNetAdjDef.TopAbuPct) & "%"
                        End If
                        
                        blnDone = False
                    Else
                        ' We're at the TopAbuPct value specifed by .NETAdjustmentUMCTopAbuPctMax
                        ' Possibly temporarily decrement .NETAdjustmentMinIDCount and repeat
                        If .NETAdjustmentMinIDCount > .NETAdjustmentMinIDCountAbsoluteMinimum Then
                            .NETAdjustmentMinIDCount = .NETAdjustmentMinIDCountAbsoluteMinimum
                            mEffectiveMinIDCount = .NETAdjustmentMinIDCount
                            
                            If Not mRobustNETInProgress Then
                                AddToAnalysisHistory CallerID, "NET Adjustment: " & strNetAdjFailureReason & "; decreased minimum ID count (since UMC top percent is already at the maximum) and will repeat search; top abundance percent = " & Trim(UMCNetAdjDef.TopAbuPct) & "%; new minimum ID count = " & Trim(.NETAdjustmentMinIDCount)
                            End If
                            
                            blnDone = False
                        Else
                            blnDone = True
                        End If
                        
                    End If
                End If
                
                If Not blnDone Then
                    ' Reset .NETTolIterative
                    UMCNetAdjDef.NETTolIterative = .NETAdjustmentInitialNetTol
                    txtNETTol = Round(.NETAdjustmentInitialNetTol, NET_TOL_DIGITS_PRECISION)
                    
                    ' Reset the NET Formula to the Default
                    ResetSlopeAndInterceptToDefault
                    
                    DoEvents
                End If
            Else
                ' Iteration was successful
                
                ' If the Match Score for the current UMC Top Abu Pct level was greater than the previous iteration,
                ' then repeat at least one more time
                
                If mCurrentIterationStats.UMCTopAbuPct < glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctMax Then
                    If udtPreviousIterationStats.Valid And mCurrentIterationStats.NETMatchScore < udtPreviousIterationStats.NETMatchScore Then
                        ' Previous iteration had a higher score than this one
                        ' Use its values
                        RestoreIterationStats udtPreviousIterationStats
                        
                        If Not mRobustNETInProgress Then
                            strMessage = "Calculated NET adjustment using UMC's; Best score obtained using top " & Round(udtPreviousIterationStats.UMCTopAbuPct, 0) & "% of the UMCs"
                            strMessage = strMessage & "; NET Formula = " & ConstructNETFormula(udtPreviousIterationStats.slope, udtPreviousIterationStats.intercept)
                            strMessage = strMessage & "; NET Match Score = " & Format(udtPreviousIterationStats.NETMatchScore, "0.00")
                        
                            AddToAnalysisHistory CallerID, strMessage
                        End If
                        
                        blnDone = True
                    Else
                    
                        ' Compute one more level of Top Abu Pct
                        
                        UMCNetAdjDef.TopAbuPct = UMCNetAdjDef.TopAbuPct + .NETAdjustmentUMCTopAbuPctIncrement
                        If UMCNetAdjDef.TopAbuPct > 100 Then UMCNetAdjDef.TopAbuPct = 100
                        txtUMCAbuTopPct = Trim(UMCNetAdjDef.TopAbuPct)
                        
                        udtPreviousIterationStats = mCurrentIterationStats
                        blnDone = False
                    End If
                Else
                    ' Stop with the current level of Top Abu Pct
                    blnDone = True
                End If
            End If
        End With
        If bStop Then blnDone = True
    Loop While Not blnDone
    
    ' .NETAdjustmentMinIDCount may have been changed above; restore to the original value
    glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount = lngOriginalNetAdjMinIDCount

End Sub

Private Sub CalculateIterationOneStepOnly()
    ' Note: This sub should only be called interactively since it
    '  displays messages using the MsgBox function
    ' It should not be called during any automated processing or automated iteration looping
    
    On Error GoTo CalculateIterationOneStepOnlyErrorHandler
    Me.MousePointer = vbHourglass
    Call ResetProcedure
    
    mIterationCount = 1
    
    UpdateStatus "Selecting UMCs to use ...."
    If LinearNETAlignmentSelectUMCToUse(CallerID, UseUMC, mUMCCntAddedSinceLowSegmentCount, mUMCSegmentCntWithLowUMCCnt) > 0 Then
        UpdateStatus "Selecting peaks to use ..."
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
             GoTo CalculateIterationOneStepOnlyErrorHandler
          End If
       Else
          GoTo CalculateIterationOneStepOnlyErrorHandler
       End If
    Else
       GoTo CalculateIterationOneStepOnlyErrorHandler
    End If
    
    RecordIterationStats mCurrentIterationStats
    mCurrentIterationStats.IterationCount = 1
    UpdateAnalysisHistory mCurrentIterationStats
    
CalculateIterationOneStepOnlyExit:
    UpdateStatus "Ready"
    Me.MousePointer = vbDefault
    Exit Sub
    
CalculateIterationOneStepOnlyErrorHandler:
    MsgBox "Error calculating NET adjustment. (Hint: selected criteria could be the reason for this procedure to fail.)", vbOKOnly, glFGTU
    Resume CalculateIterationOneStepOnlyExit

End Sub

Private Function CalculateUsingInternalStandards(intPlotUpdateInterval As Integer) As Boolean
    ' Use the Internal Standards to determine a slope and intercept
    ' Returns True if success, False if error or GANET slope and intercept could not be determined
    
    Dim blnValidSettings As Boolean
    Dim blnSuccess As Boolean
    
    Dim intNETAdjustmentMinIDCountSaved As Integer
    Dim intNETAdjustmentMinIDCountAbsoluteMinimumSaved As Integer
    Dim intNETAdjustmentMinIterationCountSaved As Integer
    
On Error GoTo CalculateUsingInternalStandardsErrorHandler

    blnSuccess = False
    
    ' First make sure the Internal Standards and settings make sense
    blnValidSettings = True
    With UMCNetAdjDef
        If UMCInternalStandards.Count < .NetAdjLockerMinimumMatchCount Then
            ' Not enough peptides to satisfy .NetAdjLockerMinimumMatchCount
            ' Cannot use Internal Standards
            blnValidSettings = False
        End If
    End With
        
    If blnValidSettings Then
        ' When using Internal Standards, we need to modify some of the parameters
        '   in UMCNetAdjDef and glbPreferencesExpanded.AutoAnalysisOptions
        
        With glbPreferencesExpanded.AutoAnalysisOptions
            intNETAdjustmentMinIDCountSaved = .NETAdjustmentMinIDCount
            intNETAdjustmentMinIDCountAbsoluteMinimumSaved = .NETAdjustmentMinIDCountAbsoluteMinimum
            intNETAdjustmentMinIterationCountSaved = .NETAdjustmentMinIterationCount
            
            .NETAdjustmentMinIDCount = UMCInternalStandards.Count                                     ' Desired match count
            .NETAdjustmentMinIDCountAbsoluteMinimum = UMCNetAdjDef.NetAdjLockerMinimumMatchCount        ' Absolute minimum match count
            .NETAdjustmentMinIterationCount = 2
        End With
        
        If Not InitializeInternalStandardSearch() Then
            blnSuccess = False
        Else
            mUsingInternalStandards = True
            
            CalculateIterationAutoIncrementTopAbuPct intPlotUpdateInterval
            
            mUsingInternalStandards = False
            
            blnSuccess = True
        End If
        
        ' Need to restore some of the settings changed above
        With glbPreferencesExpanded.AutoAnalysisOptions
            .NETAdjustmentMinIDCount = intNETAdjustmentMinIDCountSaved
            .NETAdjustmentMinIDCountAbsoluteMinimum = intNETAdjustmentMinIDCountAbsoluteMinimumSaved
            .NETAdjustmentMinIterationCount = intNETAdjustmentMinIterationCountSaved
        End With
        
    End If
    
    CalculateUsingInternalStandards = blnSuccess
    Exit Function

CalculateUsingInternalStandardsErrorHandler:
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->CalculateUsingInternalStandards"
    mUsingInternalStandards = False
    CalculateUsingInternalStandards = False
End Function

Private Sub CalculateIterationWork(Optional ByRef IterationStep As Long, Optional blnUseAbbreviatedFormat As Boolean = False, Optional intPlotUpdateInterval As Integer = 5)
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

If Not mRobustNETInProgress Then rtbIteReport.Text = ""

PrevNET1 = glHugeOverExp:       PrevNET2 = -glHugeOverExp

ShowHideControls True
mNetAdjFailed = False
mEffectiveMinIDCount = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount

If intPlotUpdateInterval < 1 Then intPlotUpdateInterval = 1

If IterationStep <= 1 Then
    ValidatePMTScoreFilters mUsingInternalStandards
End If

bPause = False
If Not mRobustNETInProgress Then bStop = False

Do While Not bDone
   IterationStep = IterationStep + 1
   UpdateStatus "Iterating NET adjustment; step " & IterationStep
   
   blnProceed = False
   If IterationStep = 1 Then
      ' On the first iteration, determine which UMC's and which Peaks are to be used
      Call ResetProcedure
      If LinearNETAlignmentSelectUMCToUse(CallerID, UseUMC(), mUMCCntAddedSinceLowSegmentCount, mUMCSegmentCntWithLowUMCCnt) > 0 Then
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
      If mUsingInternalStandards Then
           If UMCInternalStandards.Count < 1 Then
              SetNetAdjFailed
              GoTo CalculateIterationWorkExitSub
           End If
           If Not SearchInternalStandards() Then
              SetNetAdjFailed
              GoTo CalculateIterationWorkExitSub
           End If
      Else
        If AMTCnt < 1 Then
            SetNetAdjFailed
            bStop = True
            GoTo CalculateIterationWorkExitSub
        End If
        If UMCNetAdjDef.UseNET Then
           If Not SearchMassTagsMWNET() Then
              SetNetAdjFailed
              GoTo CalculateIterationWorkExitSub
           End If
        Else
           If Not SearchMassTagsMW() Then
              SetNetAdjFailed
              GoTo CalculateIterationWorkExitSub
           End If
        End If
      End If
      
      If FillTheGRID() Then
         If SelectIdentifications() > 1 Then
            Call CalculateSlopeIntercept
         Else
            SetNetAdjFailed
         End If
      End If
   Else
      SetNetAdjFailed
   End If
   
   If AdjAvD < 0 Then
      bDone = True              'cannot continue
      SetNetAdjFailed
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
         txtNETFormula.Text = UMCNetAdjDef.NETFormula
      End If
   End If
   
   If Not mRobustNETInProgress Then ReportIterationStep IterationStep
   RecordIterationStats mCurrentIterationStats
   
   ' Update the plots every intPlotUpdateInterval iterations
   If IterationStep Mod intPlotUpdateInterval = 0 Then
        If mRobustNETInProgress Then
            ' Copy the current iteration stats into the next available slot in mNETStatsHistory
            ' We're purposely not incrementing mNETStatsHistoryCount by 1 since we don't want this to be a permanently saved value
            mNETStatsHistory(mNETStatsHistoryCount) = mCurrentIterationStats
            UpdatePlots mNETStatsHistory, mNETStatsHistoryCount + 1, mNETStatsHistoryCount, mNETStatsHistoryBestIndex
        Else
            If mNETStatsHistoryCount >= UBound(mNETStatsHistory) Then
                ReDim Preserve mNETStatsHistory(mNETStatsHistoryCount * 2 - 1)
            End If
            
            mNETStatsHistory(mNETStatsHistoryCount) = mCurrentIterationStats
            
            ' The "Best" index is always the most recent one
            mNETStatsHistoryBestIndex = mNETStatsHistoryCount
            mNETStatsHistoryCount = mNETStatsHistoryCount + 1
            
            UpdatePlots mNETStatsHistory, mNETStatsHistoryCount, mNETStatsHistoryCount - 1, mNETStatsHistoryBestIndex
        End If
      
   End If
   
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
      NETTolVarChange = UMCNetAdjDef.NETTolIterative * NET_CHANGE_PCT
      UMCNetAdjDef.NETTolIterative = UMCNetAdjDef.NETTolIterative - NETTolVarChange
      If UMCNetAdjDef.NETTolIterative < NET_RESOLUTION Then UMCNetAdjDef.NETTolIterative = NET_RESOLUTION
      txtNETTol.Text = Round(UMCNetAdjDef.NETTolIterative, NET_TOL_DIGITS_PRECISION)
      txtNETTol_Validate False
   Else                 'constant change;make sure you don't go below NET resolution limits
      If UMCNetAdjDef.IterationUseNETdec Then
         UMCNetAdjDef.NETTolIterative = UMCNetAdjDef.NETTolIterative - UMCNetAdjDef.IterationNETDec
         If UMCNetAdjDef.NETTolIterative < NET_RESOLUTION Then UMCNetAdjDef.NETTolIterative = NET_RESOLUTION
         txtNETTol.Text = Round(UMCNetAdjDef.NETTolIterative, NET_TOL_DIGITS_PRECISION)
         txtNETTol_Validate False
      End If
   End If
   
   Select Case UMCNetAdjDef.IterationStopType
   Case ITERATION_STOP_NUMBER
        If IterationStep >= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_MW_TOL_LIMIT
        If UMCNetAdjDef.MWTol <= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_NET_TOL_LIMIT
        If UMCNetAdjDef.NETTolIterative <= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_ID_LIMIT
        If IDCnt <= UMCNetAdjDef.IterationStopValue Then bDone = True
   Case ITERATION_STOP_CHANGE
        CurrNET1 = AdjSlp * ScanMin + AdjInt
        CurrNET2 = AdjSlp * ScanMax + AdjInt
        '' Debug.Print IDCnt & ", " & CurrNET1 & ", " & CurrNET2 & ", " & (CurrNET1 - PrevNET1) ^ 2 + (CurrNET2 - PrevNET2) ^ 2
        mStopChangeTestResult = (CurrNET1 - PrevNET1) ^ 2 + (CurrNET2 - PrevNET2) ^ 2
        If mStopChangeTestResult <= UMCNetAdjDef.IterationStopValue ^ 2 Then
           If Not mRobustNETInProgress Then ReportIterationStepChange PrevNET1, PrevNET2, CurrNET1, CurrNET2
           bDone = True
        ElseIf IDCnt < 2 Or IterationStep > glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMaxIterationCount Then
           If Not mRobustNETInProgress Then ReportIterationStepChange PrevNET1, PrevNET2, CurrNET1, CurrNET2
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
   
   'Change NET formula if not yet done
   If Not bDone Then
      UMCNetAdjDef.NETFormula = ConstructNETFormula(AdjSlp, AdjInt)
      txtNETFormula.Text = UMCNetAdjDef.NETFormula
      CheckNETEquationStatus True
      
      If Not InitExprEvaluator(txtNETFormula.Text) Then
         UpdateStatus "Error in elution calculation formula."
         bDone = True
      End If
   Else
      UpdateStatus "Ready"
   End If
Loop

If Not mRobustNETInProgress Then
    UpdateAnalysisHistory mCurrentIterationStats, blnUseAbbreviatedFormat
End If

CalculateIterationWorkExitSub:
ShowHideControls False

Exit Sub

CalculateIterationWorkErrorHandler:
' ToDo: Figure out why an "Overflow" error is occuring a lot in this sub
LogErrors Err.Number, "CalculateIterationWork"
'Debug.Assert False
' Using Resume Next error handling if the error is a divide by zero error
If Err.Number = 11 Then Resume Next

End Sub

Private Sub CalculateRobustNET()

    Dim dblInitialSlopeSaved As Double
    Dim dblInitialInterceptSaved As Double
    
    Dim lngIndex As Long
    Dim lngIndexHighestMatchCount As Long
    
    Dim lngBestResultsCount As Long
    Dim lngBestResults() As Long
        
    Dim strMessage As String
    
    Dim udtErrorPeak As udtPeakStatsType
    Dim blnSingleGoodPeakFound As Boolean
    Dim blnMassShiftTooLarge As Boolean, blnPeakTooWide As Boolean
    Dim blnSuccess As Boolean
    
    Const MASS_PPM_ADJUSTMENT_PRECISION As Long = 4
    
    Dim dblMassCalPPMShiftSaved As Double
    Dim dblMassCalPPMCorrection As Double
    
    
    ' Initialize Robust NET Calculations
    mRobustNETInProgress = True

    bPause = False
    bStop = False
    rtbIteReport.Text = ""

    ShowHideControls True
    
    With UMCNetAdjDef
        dblInitialSlopeSaved = .InitialSlope
        dblInitialInterceptSaved = .InitialIntercept
    End With
    
    ' Make sure the NET histograms are up-to-date
    UpdateNETHistograms
    
    ' Call the appopriate Robust NET Function
    CalculateRobustNETIterative

    If mNETStatsHistoryCount > 0 Then
    
        ' Determine the best combination of slope and intercept by evaluating the results
        ' Choose the result with the highest .NETMatchScore value
        ' If more than one value has the same .NETMatchScore value (within 3 decimal places), then choose the one with the highest ID Count value
        lngBestResultsCount = 0
        ReDim lngBestResults(mNETStatsHistoryCount - 1)
        
        mNETStatsHistoryBestIndex = -1
        For lngIndex = 0 To mNETStatsHistoryCount - 1
            If mNETStatsHistory(lngIndex).Valid Then
                If mNETStatsHistoryBestIndex < 0 Then
                    mNETStatsHistoryBestIndex = lngIndex
                Else
                    If Round(mNETStatsHistory(lngIndex).NETMatchScore, 3) >= Round(mNETStatsHistory(mNETStatsHistoryBestIndex).NETMatchScore, 3) Then
                        mNETStatsHistoryBestIndex = lngIndex
                    End If
                End If
            End If
        Next lngIndex
        
        If mNETStatsHistoryBestIndex < 0 Then
            ' No good results were found
            ' Set mNETStatsHistoryCount to 0 so that the default NET slope and intercept will be used
            mNETStatsHistoryCount = 0
        End If
    End If

    If mNETStatsHistoryCount > 0 Then
            ' See if more than one value has the same score as the one at mNETStatsHistoryBestIndex
            lngBestResultsCount = 1
            lngBestResults(0) = mNETStatsHistoryBestIndex
            
            For lngIndex = 0 To mNETStatsHistoryCount - 1
                If lngIndex <> mNETStatsHistoryBestIndex Then
                    If Round(mNETStatsHistory(lngIndex).NETMatchScore, 3) = Round(mNETStatsHistory(mNETStatsHistoryBestIndex).NETMatchScore, 3) Then
                        lngBestResults(lngBestResultsCount) = lngIndex
                        lngBestResultsCount = lngBestResultsCount + 1
                    End If
                End If
            Next lngIndex
            
            If lngBestResultsCount = 1 Then
                ' One best combo of slope, intercept, and mass was found; the index is given by mNETStatsHistoryBestIndex
            Else
                ' Choose the combo with the largest UMCMatchCount value
                ' Copy the index to lngBestResults(0)
                
                lngIndexHighestMatchCount = 0
                For lngIndex = 1 To lngBestResultsCount - 1
                    If mNETStatsHistory(lngBestResults(lngIndex)).UMCMatchCount > mNETStatsHistory(lngBestResults(lngIndexHighestMatchCount)).UMCMatchCount Then
                        lngIndexHighestMatchCount = lngIndex
                    End If
                Next lngIndex
                
                mNETStatsHistoryBestIndex = lngBestResults(lngIndexHighestMatchCount)
            End If
            
            mCurrentIterationStats = mNETStatsHistory(mNETStatsHistoryBestIndex)
        
            AdjSlp = mCurrentIterationStats.slope
            AdjInt = mCurrentIterationStats.intercept
            AdjAvD = mCurrentIterationStats.Deviation
            
            GelStatus(CallerID).Dirty = True
            
            If Not GelAnalysis(CallerID) Is Nothing Then
                With GelAnalysis(CallerID)
                    .GANET_Slope = mCurrentIterationStats.slope
                    .GANET_Intercept = mCurrentIterationStats.intercept
                    .GANET_Fit = mCurrentIterationStats.Deviation
                End With
            End If
            
            UMCNetAdjDef.NETFormula = ConstructNETFormula(AdjSlp, AdjInt)
            txtNETFormula.Text = UMCNetAdjDef.NETFormula
            ValidateNETFormula
            
            ' Find the top 5 highest scoring combos in mNETStatsHistory and write them to the analysis history log
            ' Write out mNETStatsHistory(mNETStatsHistoryBestIndex) last
            AppendNETStatsHistoryToAnalysisHistory mNETStatsHistory, mNETStatsHistoryCount, mNETStatsHistoryBestIndex
            
            ' Validate that the slope is positive
            If ValidatePositiveSlope(True) Then
            
                RestoreIterationStats mNETStatsHistory(mNETStatsHistoryBestIndex)
            
                With mNETStatsHistory(mNETStatsHistoryBestIndex)
                    
                    AppendToIterationReport "Robust NET search complete" & vbCrLf
                    
                    UpdateRobustNETStatus CDbl(.slope), CDbl(.intercept), .MassShiftPPM
                
                End With
            
                ' Update the plots one more time
                UpdatePlots mNETStatsHistory, mNETStatsHistoryCount, mNETStatsHistoryBestIndex, mNETStatsHistoryBestIndex
            
            End If
    End If

    ' Finalize Robust NET Calculations
    UpdateInitialSlopeAndIntercept dblInitialSlopeSaved, dblInitialInterceptSaved
    
    If mNETStatsHistoryCount <= 0 Then
        ' Net Adjustment failed; add entry to log
        strMessage = "Error - Robust NET ended without finding any acceptable combination of slope and intercept"
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strMessage, vbExclamation Or vbOKOnly, "NET Adjustment Failure"
        Else
            AddToAnalysisHistory CallerID, strMessage
        End If
        AppendToIterationReport strMessage & vbCrLf
        
        ResetSlopeAndInterceptToDefault
    
        If Not GelAnalysis(CallerID) Is Nothing Then
            ' Need to assign a non-zero value to GANET_Fit; we'll assign 1.11E-3 with all 1's so it stands out
            GelAnalysis(CallerID).GANET_Fit = 1.11111111111111E-03
        End If
    End If
    
    UpdateStatus ""
    mRobustNETInProgress = False
    
    ShowHideControls False
End Sub

Private Sub CalculateRobustNETIterative()

    Dim sngSlope As Single
    Dim sngIntercept As Single
    
    Dim strMessage As String
    Dim lngLastBestIndex As Long
    Dim blnAppendToIterationReport As Boolean
    
    
On Error GoTo CalculateRobustNETIterativeErrorHandler

    ' Reset the progress bar
    pbarRobustNET.Value = 0
    txtRobustNETProgress.Text = "No valid NET Slope and Intercept has been found yet"

    ' Validate some parameters
    With glbPreferencesExpanded.AutoAnalysisOptions
        ' Always using linear NET intercept incrementing
        ' Make sure the Initial NET Tol value is no greater than 65% of the increment
        If .NETAdjustmentInitialNetTol > Abs(UMCNetAdjDef.RobustNETInterceptIncrement) * 0.65 Then
            .NETAdjustmentInitialNetTol = Round(Abs(UMCNetAdjDef.RobustNETInterceptIncrement) * 0.65, 5)
        End If
    End With
    
    If UMCNetAdjDef.IterationStopValue > 0.0005 Then
        UMCNetAdjDef.IterationStopValue = 0.0005
    End If
    
    ' Simulate the iterations in order to determine the total number of iterations that will be performed
    mNETStatsHistoryCount = PredictRobustNETIterationCount()
    If mNETStatsHistoryCount >= MAX_ROBUST_NET_ITERATION_COUNT Then
        mNETStatsHistoryCount = MAX_ROBUST_NET_ITERATION_COUNT
    End If
    
    ' Make a log Entry
    strMessage = "Starting Robust NET Search; Mode = Iterative; Iteration Count = " & mNETStatsHistoryCount & "; " & GetRobustNETIterationRanges()
    AddToAnalysisHistory CallerID, strMessage
    AppendToIterationReport strMessage & vbCrLf & vbCrLf
    
    If mnuFLogCalculations.Checked Then
        LogNetAdjIteration Now() & ": " & strMessage
        LogNetAdjIteration "MassShiftPPM,Slope,Intercept,MassErrorPeakHeight,PointCountSearched,Deviation,CorrelationCoeff,Computed_MatchPercentage,Computed_InverseLogDeviation,Computed_10^Correlation,MassErrorPPM,NETMatchScore"
    End If
        
    ' Initialize the progress bar
    pbarRobustNET.Min = 0
    pbarRobustNET.Max = mNETStatsHistoryCount
    
    ' Note: We're reserving some extra space in this array for scratch purposes
    ReDim mNETStatsHistory(mNETStatsHistoryCount + 1)
    
    mNETStatsHistoryCount = 0
    mNETStatsHistoryBestIndex = -1
    lngLastBestIndex = -1
    
    ' Iterate through the range of Mass Error, Slope, and Intercept values specified
    
    mMassShiftPPM = UMCNetAdjDef.RobustNETMassShiftPPMStart
    Do While mMassShiftPPM <= UMCNetAdjDef.RobustNETMassShiftPPMEnd
        sngSlope = UMCNetAdjDef.RobustNETSlopeStart
        Do While sngSlope <= UMCNetAdjDef.RobustNETSlopeEnd
            sngIntercept = UMCNetAdjDef.RobustNETInterceptStart
            Do While sngIntercept <= UMCNetAdjDef.RobustNETInterceptEnd
            
                UpdateInitialSlopeAndIntercept CDbl(sngSlope), CDbl(sngIntercept)
                ResetSlopeAndInterceptToDefault
                ResetToGenericNetAdjSettings
            
                UpdateRobustNETStatus CDbl(sngSlope), CDbl(sngIntercept), mMassShiftPPM
                
                mNegativeSlopeComputed = False
                mNetAdjFailed = False
                mEffectiveMinIDCount = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount
            
                ' Make sure the controls are still hidden
                ShowHideControls True
                
                ' Perform the calculations
                CalculateIteration 5
                
                ' Store results
                mNETStatsHistory(mNETStatsHistoryCount) = mCurrentIterationStats
                
                With mNETStatsHistory(mNETStatsHistoryCount)
                    If mNegativeSlopeComputed Or mNetAdjFailed Then
                        .Valid = False
                    Else
                        If mEffectiveMinIDCount < 2 Then mEffectiveMinIDCount = 2
                        If IDCnt < mEffectiveMinIDCount Then
                            .Valid = False
                        Else
                            .Valid = True
                        
                            ' The following assertions are sometimes false if the user cancelled the search
                            ' Otherwise, they should always be true
                            
                            Debug.Assert .slope = AdjSlp
                            Debug.Assert .intercept = AdjInt
                            
                            'Debug.Assert .UMCTopAbuPct = UMCNetAdjDef.TopAbuPct
                            
                            Debug.Assert .IterationCount = mIterationCount
                            'Debug.Assert .FinalNETTol = UMCNetAdjDef.NETTolIterative
                            'Debug.Assert .FinalMWTol = UMCNetAdjDef.MWTol
                            
                            Debug.Assert .UMCMatchCount = PeakCntWithMatches
                            Debug.Assert .PointMatchCount = UMCTotalPointsMatching
                        
                            ' Possibly update the BestIndex value
                            If mNETStatsHistoryBestIndex < 0 Then
                                mNETStatsHistoryBestIndex = mNETStatsHistoryCount
                            Else
                                If mNETStatsHistory(mNETStatsHistoryCount).NETMatchScore > mNETStatsHistory(mNETStatsHistoryBestIndex).NETMatchScore Then
                                    mNETStatsHistoryBestIndex = mNETStatsHistoryCount
                                End If
                            End If
                        
                        End If
                     
                    End If
                End With
                
                ' Update the plots
                UpdatePlots mNETStatsHistory, mNETStatsHistoryCount, mNETStatsHistoryCount, mNETStatsHistoryBestIndex
                
                ' Possibly log the current results
                If mnuFLogCalculations.Checked Then
                    With mNETStatsHistory(mNETStatsHistoryCount)
                        If .Deviation > 0 And .PointCountSearched > 0 Then
                            LogNetAdjIteration .MassShiftPPM & "," & .slope & "," & .intercept & "," & .MassErrorPeakHeight & "," & .PointCountSearched & "," & .Deviation & "," & .CorrelationCoeff & "," & (.MassErrorPeakHeight / .PointCountSearched) * 100 & "," & Log(1 / .Deviation) & "," & 10 ^ .CorrelationCoeff & "," & .MassErrorPPM & "," & .NETMatchScore
                        End If
                    End With
                End If
                
                ' Increment the history count and the progress bar
                mNETStatsHistoryCount = mNETStatsHistoryCount + 1
                pbarRobustNET.Value = mNETStatsHistoryCount
                DoEvents
                
                If mNETStatsHistoryBestIndex >= 0 Then
                    If lngLastBestIndex <> mNETStatsHistoryBestIndex Then
                        blnAppendToIterationReport = True
                        lngLastBestIndex = mNETStatsHistoryBestIndex
                    Else
                        blnAppendToIterationReport = False
                    End If
    
                    UpdateRobustNETStatusBestMatch mNETStatsHistory(mNETStatsHistoryBestIndex), blnAppendToIterationReport
                End If
                            
                ' Possibly pause
                If bPause Then ReportPause
                Do While bPause      'loop if paused
                    DoEvents
                    If bStop Then bPause = False
                    ' Sleep for 100 msec to reduce processor usage
                    Sleep 100
                Loop
                
                If mNETStatsHistoryCount >= MAX_ROBUST_NET_ITERATION_COUNT Then bStop = True
                If bStop Then Exit Do

                ' Increment intercept
                IncrementRobustNETSetting sngIntercept, UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear, UMCNetAdjDef.RobustNETInterceptIncrement
            Loop
            
            If bStop Then Exit Do
            
            ' Increment slope
            IncrementRobustNETSetting sngSlope, UMCNetAdjDef.RobustNETSlopeIncreaseMode, UMCNetAdjDef.RobustNETSlopeIncrement
        Loop
        
        If bStop Then Exit Do
        
        ' Increment the Mass Error
        IncrementRobustNETSetting mMassShiftPPM, UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear, UMCNetAdjDef.RobustNETMassShiftPPMIncrement
    Loop
   
    Exit Sub

CalculateRobustNETIterativeErrorHandler:
    Debug.Assert False
    strMessage = "Error during Robust NET Search: " & Err.Description
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox strMessage, vbOKOnly, glFGTU
    Else
        AddToAnalysisHistory CallerID, strMessage
        AppendToIterationReport strMessage
    End If

End Sub

''Private Sub CalculateRobustNETSimulatedAnnealing()
''    ' Note: The format of this sub is similar to that of CalculateIteration
''
''    Const PLOT_UPDATE_INTERVAL As Integer = 5
''
''    Dim strMessage As String
''    Dim blnSuccess As Boolean
''    Dim lngOriginalNetAdjMinIDCount As Long
''
''On Error GoTo CalculateRobustNETSimulatedAnnealingErrorHandler
''
''    Me.MousePointer = vbHourglass
''    DoEvents
''
''    ' Validate some of the simulated annealing parameters
''    With UMCNetAdjDef
''        If .RobustNETAnnealTemperatureReductionFactor < 0.0001 Or .RobustNETAnnealTemperatureReductionFactor > 0.9999 Then
''            .RobustNETAnnealTemperatureReductionFactor = 0.9
''        End If
''
''        If .RobustNETAnnealSteps < 10 Then .RobustNETAnnealSteps = 10
''        If .RobustNETAnnealTrialsPerStep < 20 Then .RobustNETAnnealTrialsPerStep = 20
''        If .RobustNETAnnealMaxSwapsPerStep < 10 Then .RobustNETAnnealMaxSwapsPerStep = 10
''
''        If .RobustNETSlopeStart = 0 Then .RobustNETSlopeStart = 0.00002
''        If .RobustNETSlopeEnd < .RobustNETSlopeStart Then .RobustNETSlopeEnd = .RobustNETSlopeStart * 10
''    End With
''
''    lngOriginalNetAdjMinIDCount = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount
''    mEffectiveMinIDCount = lngOriginalNetAdjMinIDCount
''    mIterationCount = 1
''
''    If UMCNetAdjDef.UseNetAdjLockers Then
''        ' Note that the following function will call CalculateRobustNETSimulatedAnnealingWork
''        blnSuccess = CalculateUsingInternalStandards(PLOT_UPDATE_INTERVAL, True)
''        If Not blnSuccess Then
''            If Not UMCNetAdjDef.UseOldNetAdjIfFailure Then
''                strMessage = "NET Adjustment using lockers failed.  Since UseOldNetAdjIfFailure = False, will use the default Slope and Intercept."
''
''                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''                    MsgBox strMessage, vbExclamation Or vbOKOnly, "NET Adjustment Failure"
''                Else
''                    AddToAnalysisHistory CallerID, strMessage
''                End If
''
''                ResetSlopeAndInterceptToDefault
''
''                ' Set this to True so that CalculateRobustNETSimulatedAnnealingWork isn't called below
''                blnSuccess = True
''            End If
''        End If
''    Else
''        blnSuccess = False
''    End If
''
''    If Not blnSuccess And Not bStop Then
''        ' Call CalculateRobustNETSimulatedAnnealingWork
''        CalculateRobustNETSimulatedAnnealingWork PLOT_UPDATE_INTERVAL
''    End If
''
''    glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount = lngOriginalNetAdjMinIDCount
''
''    Me.MousePointer = vbDefault
''    Exit Sub
''
''CalculateRobustNETSimulatedAnnealingErrorHandler:
''    Debug.Assert False
''    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->CalculateRobustNETSimulatedAnnealing"
''    Me.MousePointer = vbDefault
''
''End Sub
''
''Private Sub CalculateRobustNETSimulatedAnnealingWork(intPlotUpdateInterval As Integer)
''
''    Dim lngStep As Long
''    Dim lngLastBestIndex As Long
''    Dim lngTrialsPerStep As Long
''
''    Dim dblAnnealTemperature As Double
''
''    Dim blnImprovementFound As Boolean
''    Dim blnRepeatTrial As Boolean
''    Dim blnAppendToIterationReport As Boolean
''
''    Dim intMinimumTopAbuPct As Integer
''    Dim strNetAdjFailureReason As String
''    Dim strMessage As String
''
''On Error GoTo CalculateRobustNETSimulatedAnnealingWorkErrorHandler
''
''    ' Note: We're reserving some extra space in this array for scratch purposes
''    ReDim mNETStatsHistory(UMCNetAdjDef.RobustNETAnnealSteps + 1)
''    mNETStatsHistoryCount = 0
''    mNETStatsHistoryBestIndex = -1
''    lngLastBestIndex = -1
''
''    lngTrialsPerStep = UMCNetAdjDef.RobustNETAnnealTrialsPerStep
''
''    ' Reset the progress bar
''    pbarRobustNET.value = 0
''    txtRobustNETProgress.Text = "No valid NET Slope and Intercept has been found yet"
''
''    ' Initialize the progress bar
''    pbarRobustNET.Min = 0
''    pbarRobustNET.Max = UMCNetAdjDef.RobustNETAnnealSteps
''    DoEvents
''
''    ' Make a log Entry
''    strMessage = "Starting Robust NET Search; Mode = Simulated Annealing; Max Iteration Count = " & UMCNetAdjDef.RobustNETAnnealSteps & "; " & GetRobustNETIterationRanges()
''    AddToAnalysisHistory CallerID, strMessage
''    AppendToIterationReport strMessage & vbCrLf & vbCrLf
''
''    Randomize Timer
''
''    ValidatePMTScoreFilters mUsingInternalStandards
''
''    With UMCNetAdjDef
''        ' The initial temperature parameter is calculated based on the ranges of slope, intercept, and mass shift to search
''        dblAnnealTemperature = 10 * Log(Abs(1 / .RobustNETSlopeEnd - 1 / .RobustNETSlopeStart)) + 100 * Abs(.RobustNETInterceptEnd - .RobustNETInterceptStart) + 5 * Abs(.RobustNETMassShiftPPMEnd - .RobustNETMassShiftPPMStart)
''
''        ' Initialize the values to test
''        AdjAvD = 0
''        AdjSlp = Abs(.RobustNETSlopeEnd + .RobustNETSlopeStart) / 2
''        AdjInt = Abs(.RobustNETInterceptEnd + .RobustNETInterceptStart) / 2
''        mMassShiftPPM = (.RobustNETMassShiftPPMStart + .RobustNETMassShiftPPMEnd) / 2
''
''        intMinimumTopAbuPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
''        UMCNetAdjDef.TopAbuPct = intMinimumTopAbuPct
''        txtUMCAbuTopPct = UMCNetAdjDef.TopAbuPct
''
''        ' Set the NET tolerance
''        txtNETTol = Round(glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentFinalNetTol, NET_TOL_DIGITS_PRECISION)
''        .NETTolIterative = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentFinalNetTol
''
''    End With
''
''    UpdateRobustNETStatus AdjSlp, AdjInt, mMassShiftPPM
''    UpdateInitialSlopeAndIntercept AdjSlp, AdjInt
''    ResetSlopeAndInterceptToDefault
''
''    ' Populate mCurrentIterationStats based on the initial guesses
''    If Not SearchForUMCsMatchingPMTs(True) Then
''        ' Initial search failed
''        Exit Sub
''    Else
''        RecordIterationStats mCurrentIterationStats
''    End If
''
''    lngStep = 1
''    Do While lngStep <= UMCNetAdjDef.RobustNETAnnealSteps And Not bStop
''
''        ' Perform annealing
''        blnImprovementFound = CalculateRobustNETAnnealStep(dblAnnealTemperature, lngTrialsPerStep, UMCNetAdjDef.RobustNETAnnealMaxSwapsPerStep, intMinimumTopAbuPct, intPlotUpdateInterval)
''        If bStop Then Exit Do
''
''        blnRepeatTrial = False
''        If Not blnImprovementFound Then
''            ' Possibly repeat the trial
''
''            If lngStep = 1 Then
''                ' No improvement on first step, will consider raising intMinimumTopAbuPct and lowering .NetAdjustmentMinIDCount
''                blnRepeatTrial = True
''            ElseIf mNETStatsHistoryBestIndex >= 0 Then
''                ' If a valid match has not yet been found, or if the best match found has a score below .NETAdjustmentMinimumNETMatchScore, then repeat the trial
''                If Not mNETStatsHistory(mNETStatsHistoryBestIndex).Valid Then
''                    blnRepeatTrial = True
''                ElseIf mNETStatsHistory(mNETStatsHistoryBestIndex).NETMatchScore < glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinimumNETMatchScore Then
''                    blnRepeatTrial = True
''                End If
''            Else
''                If Not mCurrentIterationStats.Valid Then
''                    blnRepeatTrial = True
''                ElseIf mCurrentIterationStats.NETMatchScore < glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinimumNETMatchScore Then
''                    blnRepeatTrial = True
''                End If
''            End If
''
''            If blnRepeatTrial Then
''                ' Need to repeat the trial
''                ' Increase the minimum TopAbuPct value
''                ' If it's already at the maximum, then decrease .NETAdjustmentMinIDCount
''                ' If that's already been decreased, then increase the trials per step
''                ' If all of this has been tried, then give up
''
''                With glbPreferencesExpanded.AutoAnalysisOptions
''
''                    strNetAdjFailureReason = "No improvement found on Step " & Trim(lngStep) & " of Simulated Annealing"
''
''                    If intMinimumTopAbuPct < .NETAdjustmentUMCTopAbuPctMax Then
''                        ' Increment intMinimumTopAbuPct
''                        intMinimumTopAbuPct = intMinimumTopAbuPct + .NETAdjustmentUMCTopAbuPctIncrement
''                        If intMinimumTopAbuPct > 100 Then intMinimumTopAbuPct = 100
''                        UMCNetAdjDef.TopAbuPct = intMinimumTopAbuPct
''
''                        strMessage = "NET Adjustment: " & strNetAdjFailureReason & "; incremented high abundance UMC's top percent and will repeat search; increment = " & Trim(.NETAdjustmentUMCTopAbuPctIncrement) & "%; new top abundance percent = " & Trim(UMCNetAdjDef.TopAbuPct) & "%"
''                        AddToAnalysisHistory CallerID, strMessage
''                        AppendToIterationReport strMessage & vbCrLf & vbCrLf
''
''                        blnRepeatTrial = True
''                    Else
''                        ' We're at the TopAbuPct value specifed by .NETAdjustmentUMCTopAbuPctMax
''                        ' Possibly temporarily decrement .NETAdjustmentMinIDCount and repeat
''                        If .NETAdjustmentMinIDCount > .NETAdjustmentMinIDCountAbsoluteMinimum Then
''                            .NETAdjustmentMinIDCount = .NETAdjustmentMinIDCountAbsoluteMinimum
''                            mEffectiveMinIDCount = .NETAdjustmentMinIDCount
''
''                            strMessage = "NET Adjustment: " & strNetAdjFailureReason & "; decreased minimum ID count (since UMC top percent is already at the maximum) and will repeat search; top abundance percent = " & Trim(UMCNetAdjDef.TopAbuPct) & "%; new minimum ID count = " & Trim(.NETAdjustmentMinIDCount)
''                            AddToAnalysisHistory CallerID, strMessage
''                            AppendToIterationReport strMessage & vbCrLf & vbCrLf
''
''                            blnRepeatTrial = True
''                        Else
''                            If lngTrialsPerStep = UMCNetAdjDef.RobustNETAnnealTrialsPerStep Then
''                                lngTrialsPerStep = lngTrialsPerStep * 10
''
''                                strMessage = "NET Adjustment: " & strNetAdjFailureReason & "; increasing number of trials to " & Trim(lngTrialsPerStep) & " since no valid matches have yet been found"
''                                AddToAnalysisHistory CallerID, strMessage
''                                AppendToIterationReport strMessage & vbCrLf & vbCrLf
''
''                                blnRepeatTrial = True
''                            Else
''                                blnRepeatTrial = False
''                            End If
''                        End If
''                    End If
''                End With
''
''                If Not blnRepeatTrial Then
''                    Exit Do
''                End If
''
''            Else
''                ' Exit the loop since no improvement was found; not sure if this is best
''                Exit Do
''            End If
''
''        End If
''
''        If Not blnRepeatTrial Then
''            ' Annealing resulted in an improvement, so decrease the temperature and repeat
''            dblAnnealTemperature = dblAnnealTemperature * UMCNetAdjDef.RobustNETAnnealTemperatureReductionFactor
''
''''            If lngTrialsPerStep > UMCNetAdjDef.RobustNETAnnealTrialsPerStep Then
''''                ' The trials per step value was temporarily increased since no matches had been found previously
''''                ' Bump it back down to the defined value
''''                lngTrialsPerStep = UMCNetAdjDef.RobustNETAnnealTrialsPerStep
''''            End If
''''
''            ' Record in mNETStatsHistory()
''            mNETStatsHistory(mNETStatsHistoryCount) = mCurrentIterationStats
''
''            ' Check for new, best score
''            If mNETStatsHistoryBestIndex < 0 Then
''                mNETStatsHistoryBestIndex = mNETStatsHistoryCount
''            Else
''                If mNETStatsHistory(mNETStatsHistoryCount).NETMatchScore > mNETStatsHistory(mNETStatsHistoryBestIndex).NETMatchScore Then
''                    mNETStatsHistoryBestIndex = mNETStatsHistoryCount
''                End If
''            End If
''
''            ' Update the plots
''            UpdatePlots mNETStatsHistory, mNETStatsHistoryCount + 1, mNETStatsHistoryCount, mNETStatsHistoryBestIndex
''
''            mNETStatsHistoryCount = mNETStatsHistoryCount + 1
''
''            If mNETStatsHistoryBestIndex >= 0 Then
''                If lngLastBestIndex <> mNETStatsHistoryBestIndex Then
''                    blnAppendToIterationReport = True
''                    lngLastBestIndex = mNETStatsHistoryBestIndex
''                Else
''                    blnAppendToIterationReport = False
''                End If
''
''                UpdateRobustNETStatusBestMatch mNETStatsHistory(mNETStatsHistoryBestIndex), blnAppendToIterationReport
''            End If
''
''            pbarRobustNET.value = lngStep
''            DoEvents
''
''            lngStep = lngStep + 1
''
''            ' This should always be true
''            Debug.Assert lngStep = mNETStatsHistoryCount + 1
''        End If
''    Loop
''    Exit Sub
''
''CalculateRobustNETSimulatedAnnealingWorkErrorHandler:
''    Debug.Assert False
''    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->CalculateRobustNETSimulatedAnnealingWork"
''    Me.MousePointer = vbDefault
''
''End Sub
''
''Private Function CalculateRobustNETAnnealStep(dblAnnealTemperature As Double, lngTrialsPerStep As Long, lngMaxSwapsPerStep As Long, intTopAbuPctMinimum As Integer, intPlotUpdateInterval) As Boolean
''    ' Note: mCurrentIterationStats holds the best result from the previous step
''
''    Dim intTopAbuPctMax As Integer
''    Const TOP_ABU_PCT_ROUNDINT As Integer = 10
''
''    Dim lngTrial As Long
''    Dim lngTopAbuPct As Long
''    Dim intTopAbuPctTestCount As Integer
''
''    Dim blnImprovementFound As Boolean
''    Dim blnSaveCurrentIteration As Boolean
''
''    Dim blnForceSelectUMCs As Boolean
''
''    Dim dblDelta As Double
''    Dim p As Double, m As Double
''
''    Dim lngIterationHistoryCount As Long
''    Dim udtIterationHistory() As udtIterationStatsType      ' 0-based array
''    Dim lngBestIterationIndex As Long
''
''On Error GoTo CalculateRobustNETAnnealStepErrorHandler
''
''    ReDim udtIterationHistory(lngTrialsPerStep)
''
''    lngIterationHistoryCount = 1
''    udtIterationHistory(0) = mCurrentIterationStats
''    lngBestIterationIndex = 0
''
''    blnImprovementFound = False
''    blnForceSelectUMCs = True
''
''    If intTopAbuPctMinimum < TOP_ABU_PCT_ROUNDINT Then intTopAbuPctMinimum = TOP_ABU_PCT_ROUNDINT
''    intTopAbuPctMax = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctMax
''    If intTopAbuPctMax < intTopAbuPctMinimum Then intTopAbuPctMax = intTopAbuPctMinimum
''
''    If intPlotUpdateInterval < 1 Then intPlotUpdateInterval = 1
''
''    For lngTrial = 1 To lngTrialsPerStep
''
''        ' Pick a new slope, intercept, MassShift, and TopAbuPct
''        ' ToDo: Use our prior search history to guide these guesses
''
''        With UMCNetAdjDef
''            AdjSlp = (.RobustNETSlopeEnd - .RobustNETSlopeStart) * Rnd + .RobustNETSlopeStart
''            AdjInt = (.RobustNETInterceptEnd - .RobustNETInterceptStart) * Rnd + .RobustNETInterceptStart
''            mMassShiftPPM = (.RobustNETMassShiftPPMEnd - .RobustNETMassShiftPPMStart) * Rnd + .RobustNETMassShiftPPMStart
''
''            ' Pick a new Top AbuPct value every 5 trials
''            ' This is done to reduce the number of times we select the UMCs to use
''            If (lngTrial - 1) Mod 5 = 0 Then
''                If intTopAbuPctMinimum = intTopAbuPctMax Then
''                    ' Cannot pick a new Top Abu Pct value
''                    lngTopAbuPct = intTopAbuPctMinimum
''                Else
''                    intTopAbuPctTestCount = 0
''                    Do
''                        lngTopAbuPct = Round((intTopAbuPctMax - intTopAbuPctMinimum) * Rnd / TOP_ABU_PCT_ROUNDINT, 0) * TOP_ABU_PCT_ROUNDINT + intTopAbuPctMinimum
''                        intTopAbuPctTestCount = intTopAbuPctTestCount + 1
''                        If intTopAbuPctTestCount > 100 Or lngTrial = 1 Then Exit Do
''                    Loop While lngTopAbuPct = .TopAbuPct
''
''                    If lngTopAbuPct < intTopAbuPctMinimum Then
''                        ' This shouldn't happen
''                        Debug.Assert False
''                        lngTopAbuPct = intTopAbuPctMinimum
''                    End If
''
''                    If lngTopAbuPct > intTopAbuPctMax Then
''                        ' This shouldn't happen
''                        Debug.Assert False
''                        lngTopAbuPct = intTopAbuPctMax
''                    End If
''
''                End If
''
''                If .TopAbuPct <> lngTopAbuPct Then
''                    .TopAbuPct = lngTopAbuPct
''                    blnForceSelectUMCs = True
''                End If
''                txtUMCAbuTopPct = Trim(UMCNetAdjDef.TopAbuPct)
''            End If
''        End With
''
''        UpdateRobustNETStatus AdjSlp, AdjInt, mMassShiftPPM
''        UpdateInitialSlopeAndIntercept AdjSlp, AdjInt
''        ResetSlopeAndInterceptToDefault
''
''        ' Match the UMCs to the MTs
''        SearchForUMCsMatchingPMTs blnForceSelectUMCs
''        blnForceSelectUMCs = False
''
''        RecordIterationStats mCurrentIterationStats
''
''        If IDCnt < 1 Or mCurrentIterationStats.NETMatchScore = 0 Or Not mCurrentIterationStats.Valid Then
''            blnSaveCurrentIteration = False
''        Else
''            ' See if a better score was found
''            dblDelta = mCurrentIterationStats.NETMatchScore - udtIterationHistory(lngIterationHistoryCount - 1).NETMatchScore
''
''            If dblDelta > 0 Then
''                ' Improved score combo found
''                blnImprovementFound = True
''                blnSaveCurrentIteration = True
''            Else
''                ' Test the Metropolis criterion
''
''                p = Rnd()
''                m = Exp(dblDelta / dblAnnealTemperature)
''
''                If p < m Then
''                    ' Record this as an improved score combo found
''                    ' ?? The java demo sets blnImprovementFound to True, even though the score is lower
''                    blnImprovementFound = True
''                    blnSaveCurrentIteration = True
''                Else
''                    ' Not an improvement
''                    blnSaveCurrentIteration = False
''                End If
''            End If
''        End If
''
''        If blnSaveCurrentIteration Then
''            udtIterationHistory(lngIterationHistoryCount) = mCurrentIterationStats
''
''            If mCurrentIterationStats.NETMatchScore > udtIterationHistory(lngBestIterationIndex).NETMatchScore Then
''                lngBestIterationIndex = lngIterationHistoryCount
''            End If
''
''            lngIterationHistoryCount = lngIterationHistoryCount + 1
''        End If
''
''        ' Update the plots every intPlotUpdateInterval iterations
''        If lngTrial Mod intPlotUpdateInterval = 0 Then
''            ' Copy the current iteration stats into the next available slot in mNETStatsHistory
''            ' We're purposely not incrementing mNETStatsHistoryCount by 1 since we don't want this to be a permanently saved value
''            mNETStatsHistory(mNETStatsHistoryCount) = mCurrentIterationStats
''            UpdatePlots mNETStatsHistory, mNETStatsHistoryCount + 1, mNETStatsHistoryCount, mNETStatsHistoryBestIndex
''        End If
''
''        If lngIterationHistoryCount > lngMaxSwapsPerStep Then
''            ' Found enough good improvements
''            Exit For
''        End If
''
''        ' Possibly pause
''        If bPause Then ReportPause
''        Do While bPause      'loop if paused
''            DoEvents
''            If bStop Then bPause = False
''            ' Sleep for 100 msec to reduce processor usage
''            Sleep 100
''        Loop
''        If bStop Then Exit For
''
''    Next lngTrial
''
''    If blnImprovementFound Then
''        ' Copy the data from the highest scoring iteration index to mCurrentIterationStats
''        ' ?? Or, should we be copying the data from the final trial step back to mCurrentIterationStats, as is done in the java demo ??
''        mCurrentIterationStats = udtIterationHistory(lngBestIterationIndex)
''        RestoreIterationStats mCurrentIterationStats
''    End If
''
''    CalculateRobustNETAnnealStep = blnImprovementFound
''    Exit Function
''
''CalculateRobustNETAnnealStepErrorHandler:
''    Debug.Assert False
''    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->CalculateRobustNETAnnealStep"
''    Me.MousePointer = vbDefault
''
''End Function

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
    If IDsAreInternalStandards Then
        With UMCInternalStandards.InternalStandards(ID(i))
            SumY = SumY + .NET
            SumXY = SumXY + IDScan(i) * .NET
        End With
    Else
        SumY = SumY + AMTData(ID(i)).NET
        SumXY = SumXY + IDScan(i) * AMTData(ID(i)).NET
    End If
    SumXX = SumXX + CDbl(IDScan(i)) ^ 2
Next i
AdjSlp = (IDCnt * SumXY - SumX * SumY) / (IDCnt * SumXX - SumX * SumX)
AdjInt = (SumY - AdjSlp * SumX) / IDCnt
Call CalculateAvgDev
UpdateStatus ""
Exit Sub


err_CalculateSlopeIntercept:
UpdateStatus "Error calculating slope and intercept."
AdjSlp = 0
AdjInt = 0
AdjAvD = -1
End Sub


Private Sub CalculateAvgDev()
Dim i As Long
Dim TtlDist As Double
On Error GoTo err_CalculateAvgDev
TtlDist = 0

If IDsAreInternalStandards Then
    With UMCInternalStandards
        For i = 0 To IDCnt - 1
            TtlDist = TtlDist + (AdjSlp * IDScan(i) + AdjInt - .InternalStandards(ID(i)).NET) ^ 2
        Next i
    End With
Else
    For i = 0 To IDCnt - 1
        TtlDist = TtlDist + (AdjSlp * IDScan(i) + AdjInt - AMTData(ID(i)).NET) ^ 2
    Next i
End If

AdjAvD = TtlDist / IDCnt
Exit Sub

err_CalculateAvgDev:
AdjAvD = -1
End Sub

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

Private Sub CleanIdentifications()
'-------------------------------------------------------------
'removes identifications that will not be used from the arrays
'-------------------------------------------------------------
Dim i As Long
Dim NewCnt As Long
Dim UMCUsed() As Boolean

On Error GoTo CleanIdentificationsErrorHandler

ReDim UMCUsed(GelUMC(CallerID).UMCCnt)

UpdateStatus "Restructuring data ..."
For i = 0 To IDCnt - 1
    If IDState(i) = 0 Then
       NewCnt = NewCnt + 1
       ID(NewCnt - 1) = ID(i)
       IDMatchingUMCInd(NewCnt - 1) = IDMatchingUMCInd(i)
       IDMatchAbu(NewCnt - 1) = IDMatchAbu(i)
       IDIsoFit(NewCnt - 1) = IDIsoFit(i)
       IDScan(NewCnt - 1) = IDScan(i)
       
       UMCUsed(IDMatchingUMCInd(i)) = True
    End If
Next i
If NewCnt > 0 Then
    ReDim Preserve ID(NewCnt - 1)
    ReDim Preserve IDMatchingUMCInd(NewCnt - 1)
    ReDim Preserve IDMatchAbu(NewCnt - 1)
    ReDim Preserve IDIsoFit(NewCnt - 1)
    ReDim Preserve IDScan(NewCnt - 1)
    ReDim IDState(NewCnt - 1)            'after cleaning always put state to 0
Else
    Call ClearIDArrays
End If

If NewCnt < IDCnt Then
    IDCnt = NewCnt
    ClearTheGRID         'have to recalculate GRID data
    Call FillTheGRID

    ' Now that the IDs have been trimmed, update PeakCntWithMatches and UMCTotalPointsMatching
    PeakCntWithMatches = 0
    UMCTotalPointsMatching = 0
    With GelUMC(CallerID)
        For i = 0 To .UMCCnt - 1
            If UMCUsed(i) Then
                PeakCntWithMatches = PeakCntWithMatches + 1
                UMCTotalPointsMatching = UMCTotalPointsMatching + .UMCs(i).ClassCount
            End If
        Next i
    End With

End If
UpdateStatus ""
Exit Sub

CleanIdentificationsErrorHandler:
Debug.Assert False
LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->CleanIdentifications"
End Sub

Public Sub ClearIDArrays()
'-------------------------------------------------------------------
'clears arrays of identifications to be used for UMC NET adjustments
'-------------------------------------------------------------------
Erase ID
Erase IDMatchingUMCInd
Erase IDMatchAbu
Erase IDIsoFit
Erase IDScan
Erase IDState
IDCnt = 0
PeakCntWithMatches = 0
UMCTotalPointsMatching = 0
End Sub

Public Sub ClearPeakArrays()
'------------------------------------------------------------------
'clears arrays of peaks selected to be used for UMC NET adjustments
'------------------------------------------------------------------
Erase PeakUMCInd
Erase PeakMW
Erase PeakScan
Erase PeakIsoFit
PeakCnt = 0
PeakCntWithMatches = 0
UMCTotalPoints = 0
UMCTotalPointsMatching = 0
End Sub

Private Sub ClearPlotDataSingleSeries(ByRef ctlPlot As ctlSpectraPlotter, ByVal intSeriesNumber As Integer)
    Dim dblDataX(1 To 1) As Double      ' 1-based array
    Dim dblDataY(1 To 1) As Double      ' 1-based array
    
On Error GoTo ClearPlotDataSingleSeriesErrorHandler
    
    With ctlPlot
        .EnableDisableDelayUpdating True
        .SetSeriesDataPointCount intSeriesNumber, 1
        .SetDataX intSeriesNumber, dblDataX
        .SetDataY intSeriesNumber, dblDataY
        .EnableDisableDelayUpdating False
    End With
    Exit Sub

ClearPlotDataSingleSeriesErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->ClearPlotDataSingleSeries"
    
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

Private Sub ComputeMassAndNETErrorHistogram(ByRef udtMassErrorHistogramData As udtBinnedDataType, ByRef udtNETErrorHistogramData As udtBinnedDataType, dblSlope As Double, dblIntercept As Double)

    Const ERROR_DATA_DIM_CHUNK = 10000
    
    Dim lngDataCount As Long
    Dim lngDataCountDimmed As Long
    Dim sngMassErrors() As Single       ' 0-based array
    Dim sngNETErrors() As Single        ' 0-based array
    
    Dim lngIDIndex As Long
    Dim lngMemberIndex As Long
    
    Dim blnShiftMass As Boolean
    
    Dim dblCurrMW As Double
    Dim dblCurrNET As Double
    
    On Error GoTo ComputeMassAndNETErrorHistogramErrorHandler
    
    If IDCnt <= 0 Or dblSlope <= 0 Then
        With udtMassErrorHistogramData
            ReDim .Binned(0)
            ReDim .SmoothedBins(0)
            .BinnedCount = 0
        End With
    
        With udtNETErrorHistogramData
            ReDim .Binned(0)
            ReDim .SmoothedBins(0)
            .BinnedCount = 0
        End With
    Else
        ' The Mass and NET error arrays are based on every point in every UMC with a match
        ' We do not know the exact number of points, so we'll reserve data in these arrays in chunks
        
        If Round(mMassShiftPPM, 3) <> 0 Then
            blnShiftMass = True
        Else
            blnShiftMass = False
        End If
        
        lngDataCountDimmed = ERROR_DATA_DIM_CHUNK
        lngDataCount = 0
        ReDim sngMassErrors(0 To lngDataCountDimmed - 1)
        ReDim sngNETErrors(0 To lngDataCountDimmed - 1)
                
        For lngIDIndex = 0 To IDCnt - 1
            With GelUMC(CallerID).UMCs(IDMatchingUMCInd(lngIDIndex))
                
                For lngMemberIndex = 0 To .ClassCount - 1
                    
                    Select Case .ClassMType(lngMemberIndex)
                    Case glCSType
                        dblCurrMW = GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).AverageMW
                        dblCurrNET = ComputeNET(GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).ScanNumber, dblSlope, dblIntercept)
                    Case glIsoType
                        dblCurrMW = GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)).MonoisotopicMW
                        dblCurrNET = ComputeNET(GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)).ScanNumber, dblSlope, dblIntercept)
                    End Select
    
                    If blnShiftMass Then
                        dblCurrMW = dblCurrMW + PPMToMass(CDbl(mMassShiftPPM), dblCurrMW)
                    End If
                    
                    If lngDataCount >= lngDataCountDimmed Then
                        lngDataCountDimmed = lngDataCountDimmed + ERROR_DATA_DIM_CHUNK
                        ReDim Preserve sngMassErrors(lngDataCountDimmed - 1)
                        ReDim Preserve sngNETErrors(lngDataCountDimmed - 1)
                    End If
                    
                    sngMassErrors(lngDataCount) = MassToPPM(dblCurrMW - AMTData(ID(lngIDIndex)).MW, dblCurrMW)
                    sngNETErrors(lngDataCount) = dblCurrNET - AMTData(ID(lngIDIndex)).NET
                    lngDataCount = lngDataCount + 1
                Next lngMemberIndex
                
            End With
        Next lngIDIndex
        
        ' Initialize the histogram object
        With objHistogram
             .RequireNegativeStartBin = True
             .ShowMessages = Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
        End With
                 
         ' Now bin the Mass Error data
         With udtMassErrorHistogramData
             objHistogram.BinSize = MASS_ERROR_HISTOGRAM_BIN_SIZE
             objHistogram.DefaultBinSize = MASS_ERROR_HISTOGRAM_BIN_SIZE
             objHistogram.StartBinDigitsAfterDecimal = MASS_ERROR_HISTOGRAM_DIGITS_AFTER_DECIMAL
             
             If Not objHistogram.BinData(sngMassErrors, lngDataCount, .Binned, .BinnedCount) Then
                 LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputeMassAndNETErrorHistogram->BinMassErrors"
             End If
             
             .StartBin = objHistogram.StartBin
             .BinSize = objHistogram.BinSize
             .BinRangeMaximum = objHistogram.BinRangeMaximum
             .BinCountMaximum = objHistogram.BinCountMaximum
         End With
                 
         ' Now bin the NET Error data
         With udtNETErrorHistogramData
             objHistogram.BinSize = NET_ERROR_HISTOGRAM_BIN_SIZE
             objHistogram.DefaultBinSize = NET_ERROR_HISTOGRAM_BIN_SIZE
             objHistogram.StartBinDigitsAfterDecimal = NET_ERROR_HISTOGRAM_DIGITS_AFTER_DECIMAL
             
             If Not objHistogram.BinData(sngNETErrors, lngDataCount, .Binned, .BinnedCount) Then
                 LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputeMassAndNETErrorHistogram->BinNETErrors"
             End If
             
             .StartBin = objHistogram.StartBin
             .BinSize = objHistogram.BinSize
             .BinRangeMaximum = objHistogram.BinRangeMaximum
             .BinCountMaximum = objHistogram.BinCountMaximum
         End With
    End If
    
    Exit Sub

ComputeMassAndNETErrorHistogramErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->ComputeMassAndNETErrorHistogram"

End Sub

Private Sub ComputeMatchingIDsNETHistogram(ByRef udtNETHistogramData As udtBinnedDataType, ByVal dblSlope As Double, ByVal dblIntercept As Double)
    
    Dim sngUMCNET() As Single           ' 0-based array
    Dim lngIndex As Long
    
    On Error GoTo ComputeMatchingIDsNETHistogramErrorHandler
    
    If IDCnt <= 0 Or dblSlope <= 0 Then
        With udtNETHistogramData
            ReDim .Binned(0)
            ReDim .SmoothedBins(0)
            .BinnedCount = 0
        End With
    Else
        
        ReDim sngUMCNET(0 To IDCnt - 1)
        
        For lngIndex = 0 To IDCnt - 1
            sngUMCNET(lngIndex) = ComputeNET(IDScan(lngIndex), dblSlope, dblIntercept)
        Next lngIndex
        
        ' Initialize the histogram object
        With objHistogram
             .RequireNegativeStartBin = False
             .ShowMessages = Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
        End With
                 
         ' Now bin the data
         With udtNETHistogramData
             objHistogram.BinSize = NET_HISTOGRAM_BIN_SIZE
             objHistogram.DefaultBinSize = NET_HISTOGRAM_BIN_SIZE
             objHistogram.StartBinDigitsAfterDecimal = NET_HISTOGRAM_DIGITS_AFTER_DECIMAL
             
             If Not objHistogram.BinData(sngUMCNET, IDCnt, .Binned, .BinnedCount) Then
                 LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputeMatchingIDsNETHistogram"
             End If
             
             .StartBin = objHistogram.StartBin
             .BinSize = objHistogram.BinSize
             .BinRangeMaximum = objHistogram.BinRangeMaximum
             .BinCountMaximum = objHistogram.BinCountMaximum
         End With
    End If
    
    Exit Sub

ComputeMatchingIDsNETHistogramErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->ComputeMatchingIDsNETHistogram"
        
End Sub

Private Sub ComputeInternalStandardNETHistogram()
    ' Note: This sub is typically called from UpdateNETHistograms
    
    Dim sngNET() As Single           ' 0-based array
    Dim lngIndex As Long
    
    On Error GoTo ComputeInternalStandardNETHistogramErrorHandler
    
    ValidatePMTScoreFilters True
    
    If UMCInternalStandards.Count <= 0 Then
        With mInternalStandardsNETHistogram
            ReDim .Binned(0)
            ReDim .SmoothedBins(0)
            .BinnedCount = 0
        End With
    Else
        
        ReDim sngNET(0 To UMCInternalStandards.Count - 1)
        
        For lngIndex = 0 To UMCInternalStandards.Count - 1
            sngNET(lngIndex) = UMCInternalStandards.InternalStandards(lngIndex).NET
        Next lngIndex
        
        ' Initialize the histogram object
        With objHistogram
             .RequireNegativeStartBin = False
             .ShowMessages = Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
        End With
                 
         ' Now bin the data
         With mInternalStandardsNETHistogram
             objHistogram.BinSize = NET_HISTOGRAM_BIN_SIZE
             objHistogram.DefaultBinSize = NET_HISTOGRAM_BIN_SIZE
             objHistogram.StartBinDigitsAfterDecimal = NET_HISTOGRAM_DIGITS_AFTER_DECIMAL
             
             If Not objHistogram.BinData(sngNET, CLng(UMCInternalStandards.Count), .Binned, .BinnedCount) Then
                 LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputeInternalStandardNETHistogram"
             End If
             
             .StartBin = objHistogram.StartBin
             .BinSize = objHistogram.BinSize
             .BinRangeMaximum = objHistogram.BinRangeMaximum
             .BinCountMaximum = objHistogram.BinCountMaximum
         End With
    End If
    
    Exit Sub

ComputeInternalStandardNETHistogramErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->ComputeInternalStandardNETHistogram"
        
End Sub

Private Sub ComputePMTDataNETHistogram()
    ' Note: This sub is typically called from UpdateNETHistograms

    Dim lngDataCount As Long
    Dim sngAMTNET() As Single           ' 0-based array, though global array AMTData() is 1-based
   
    Dim lngIndex As Long
    Dim blnIncludeAllPMTs As Boolean
    
    On Error GoTo ComputePMTDataNETHistogramErrorHandler
    
    ValidatePMTScoreFilters False
    
    If AMTCnt <= 0 Then
        With mPMTDataNETHistogram
            ReDim .Binned(0)
            ReDim .SmoothedBins(0)
            .BinnedCount = 0
        End With
    Else
        
        ReDim sngAMTNET(0 To AMTCnt - 1)
        
        If Not (mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0) Then
            blnIncludeAllPMTs = True
        Else
            blnIncludeAllPMTs = False
        End If
        
        lngDataCount = 0
        For lngIndex = 1 To AMTCnt
            ' Only include the PMT if it is above the score cutoffs
            If blnIncludeAllPMTs Or (AMTData(lngIndex).HighNormalizedScore >= mMTMinimumHighNormalizedScore And AMTData(lngIndex).HighDiscriminantScore >= mMTMinimumHighDiscriminantScore) Then
                sngAMTNET(lngDataCount) = AMTData(lngIndex).NET
                lngDataCount = lngDataCount + 1
            End If
        Next lngIndex
        
        ' Initialize the histogram object
        With objHistogram
             .RequireNegativeStartBin = False
             .ShowMessages = Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
        End With
                 
         ' Now bin the data
         With mPMTDataNETHistogram
             objHistogram.BinSize = NET_HISTOGRAM_BIN_SIZE
             objHistogram.DefaultBinSize = NET_HISTOGRAM_BIN_SIZE
             objHistogram.StartBinDigitsAfterDecimal = NET_HISTOGRAM_DIGITS_AFTER_DECIMAL
             
             If Not objHistogram.BinData(sngAMTNET, lngDataCount, .Binned, .BinnedCount) Then
                 LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputePMTDataNETHistogram"
             End If
             
             .StartBin = objHistogram.StartBin
             .BinSize = objHistogram.BinSize
             .BinRangeMaximum = objHistogram.BinRangeMaximum
             .BinCountMaximum = objHistogram.BinCountMaximum
         End With
    End If
    
    Exit Sub

ComputePMTDataNETHistogramErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->ComputePMTDataNETHistogram"
        
End Sub

Private Sub ComputeScoreForIterationStats(ByRef udtIterationStats As udtIterationStatsType)
    ' Updates the histograms for the data in udtIterationStats and updates the NETMatchScore
    
    Const PCT_MATCH_WEIGHT As Single = 15
    Const DEVIATION_WEIGHT As Single = 5
    Const CORRELATION_WEIGHT As Single = 50
    Const MASS_ERROR_WEIGHT As Single = 3           ' Note that the absolute value of mass error is subtracted from the score

    Const MIN_ALLOWED_DEVIATION As Double = 0.00000001
    
    Dim udtErrorPeak As udtPeakStatsType
    Dim dblPeakCenter As Double, dblPeakWidth As Double, dblPeakHeight As Double
    Dim sngSignalToNoise As Single
    
    Dim blnMassErrorPeakFound As Boolean
    Dim blnSingleGoodPeakFound As Boolean
    
    With udtIterationStats
        ' Update the NET Histogram
        ComputeMatchingIDsNETHistogram .NETRangeHistogramData, .slope, .intercept
        
        If .NETRangeHistogramData.BinnedCount > 0 Then
            ' Compute the correlation between the NETRangeHistogramData and the data searched
            If .UseInternalStandards Then
                .CorrelationCoeff = CorrelateBinnedData(mInternalStandardsNETHistogram, .NETRangeHistogramData)
            Else
                .CorrelationCoeff = CorrelateBinnedData(mPMTDataNETHistogram, .NETRangeHistogramData)
            End If
            
            If .CorrelationCoeff < 0 Then
                .CorrelationCoeff = 0
            End If
        Else
            .CorrelationCoeff = 0
        End If
        
        ' Update the Mass Error and NET error plot data
        ComputeMassAndNETErrorHistogram .MassErrorHistogramData, .NETErrorHistogramData, .slope, .intercept
        
        ' Determine .MassErrorPPM
        blnMassErrorPeakFound = False
        If .MassErrorHistogramData.BinnedCount > 0 Then
            If FindPeakStatsUsingBinnedErrorData(.MassErrorHistogramData, udtErrorPeak, blnSingleGoodPeakFound) Then
                
                GetPeakStats .MassErrorHistogramData, udtErrorPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, 3
                
                .MassErrorPPM = dblPeakCenter
                .MassErrorPeakHeight = CLng(dblPeakHeight)
                
                blnMassErrorPeakFound = True
            End If
        End If
        
        If Not blnMassErrorPeakFound Then
        
            ' Assign a fake mass error value since a mass error histogram couldn't be generated
            .MassErrorPPM = 10
            If Abs(UMCNetAdjDef.RobustNETMassShiftPPMStart) > .MassErrorPPM Then
                .MassErrorPPM = Abs(UMCNetAdjDef.RobustNETMassShiftPPMStart)
            End If
            
            If Abs(UMCNetAdjDef.RobustNETMassShiftPPMEnd) > .MassErrorPPM Then
                .MassErrorPPM = Abs(UMCNetAdjDef.RobustNETMassShiftPPMEnd)
            End If
            
            If .MassErrorHistogramData.BinnedCount > 0 Then
                ' Use the maximum bin count value for the mass error peak height
                .MassErrorPeakHeight = .MassErrorHistogramData.BinCountMaximum
            Else
                ' No histogram data exists
                ' Divide the PointMatchCount value by an estimate of the number of bins that would normally be present (25 / 0.5 = 50)
                .MassErrorPeakHeight = .PointMatchCount / (25 / MASS_ERROR_HISTOGRAM_BIN_SIZE)
            End If
        End If
        
        
        ' Compute the Match Score value
        On Error Resume Next
        .NETMatchScore = 0
        
        If .Deviation > 0 And .Deviation < MIN_ALLOWED_DEVIATION Then
            .Deviation = MIN_ALLOWED_DEVIATION
        End If
        
        If .PointCountSearched > 0 And .Deviation > 0 Then
            ' Note that Log() is the Natural Log, aka LN()
            
            .NETMatchScore = PCT_MATCH_WEIGHT * (.MassErrorPeakHeight / .PointCountSearched * 100#) + DEVIATION_WEIGHT * Log(1 / .Deviation) + CORRELATION_WEIGHT * .CorrelationCoeff - MASS_ERROR_WEIGHT * Abs(.MassErrorPPM)
        
        End If
    End With
    
End Sub

Private Function ConvertScanToNET(lngScanNumber As Long) As Double

    If mUsingDefaultGANET Then
        Debug.Assert InStr(UCase(txtNETFormula), "MINFN") = 0
        ConvertScanToNET = ScanToGANET(CallerID, lngScanNumber)
    Else
        ConvertScanToNET = GetElution(lngScanNumber)
    End If
    
End Function

Private Function CorrelateBinnedData(ByRef udtList1 As udtBinnedDataType, ByRef udtList2 As udtBinnedDataType) As Single
    
    Dim sngMinimum As Single
    Dim sngMaximum As Single
    
    Dim sngDataList1() As Single
    Dim sngDataList2() As Single
    Dim lngDataCount As Long
    
    Dim lngIndex As Long
    Dim lngCopyStartIndex As Long
    
    On Error GoTo CorrelateBinnedDataErrorHandler
    
    If objCorrelate Is Nothing Then objCorrelate = New clsCorrelation

    ' Assure that the two lists have the same bin size
    If udtList1.BinSize <> udtList2.BinSize Or udtList1.BinSize <= 0 Then
        Debug.Assert False
        CorrelateBinnedData = 0
        Exit Function
    End If

    ' Find the minimum and maximum bin value in udtList1 and udtList2
    sngMinimum = udtList1.StartBin
    If udtList2.StartBin < sngMinimum Then sngMinimum = udtList2.StartBin
    
    sngMaximum = udtList1.BinRangeMaximum
    If udtList2.BinRangeMaximum > sngMaximum Then sngMaximum = udtList2.BinRangeMaximum
    
    ' Reserve space in sngDataList1 and sngDataList2
    lngDataCount = (sngMaximum - sngMinimum) / udtList1.BinSize + 1
        
    If lngDataCount < udtList1.BinnedCount + 1 Then
        Debug.Assert False
        lngDataCount = udtList1.BinnedCount + 1
    End If
    
    If lngDataCount < udtList2.BinnedCount + 1 Then
        Debug.Assert False
        lngDataCount = udtList2.BinnedCount + 1
    End If
    
    
    ReDim sngDataList1(lngDataCount - 1)
    ReDim sngDataList2(lngDataCount - 1)
    
    ' Copy the data from udtList1 into sngDataList1
    lngCopyStartIndex = Round((udtList1.StartBin - sngMinimum) / udtList1.BinSize, 0)
    If lngCopyStartIndex < 0 Then lngCopyStartIndex = 0
    
    For lngIndex = 0 To udtList1.BinnedCount
        sngDataList1(lngCopyStartIndex + lngIndex) = udtList1.Binned(lngIndex)
    Next lngIndex
    
    ' Copy the data from udtList2 into sngDataList2
    lngCopyStartIndex = Round((udtList2.StartBin - sngMinimum) / udtList2.BinSize, 0)
    If lngCopyStartIndex < 0 Then lngCopyStartIndex = 0
    
    For lngIndex = 0 To udtList2.BinnedCount
        sngDataList2(lngCopyStartIndex + lngIndex) = udtList2.Binned(lngIndex)
    Next lngIndex
    
    CorrelateBinnedData = objCorrelate.Correlate(sngDataList1, sngDataList2, cmCorrelationMethodConstants.Pearson)
    Exit Function
    
CorrelateBinnedDataErrorHandler:
    Debug.Assert False
    CorrelateBinnedData = 0

End Function

Private Sub DisplayDefaultSettings()

    On Error GoTo DisplayDefaultSettingsErrorHandler
    
    SetDefaultUMCNETAdjDef GelUMCNETAdjDef(CallerID)
    
    ResetExpandedPreferences glbPreferencesExpanded, "NetAdjustmentOptions"
    ResetExpandedPreferences glbPreferencesExpanded, "NetAdjustmentUMCDistributionOptions"
    ResetExpandedPreferences glbPreferencesExpanded, "RefineMSDataOptions"

    With glbPreferencesExpanded
        .NetAdjustmentUsesN15AMTMasses = False
        .NetAdjustmentMinHighNormalizedScore = 2.5
        .NetAdjustmentMinHighDiscriminantScore = 0.5
        .PairSearchOptions.NETAdjustmentPairedSearchUMCSelection = punaUnpairedPlusPairedLight
    End With
    
    InitializeForm True
    
    EnableDisableControls
    
    Exit Sub

DisplayDefaultSettingsErrorHandler:
    Debug.Assert False
    
End Sub

Private Sub EnableDisableControls(Optional ByVal blnCalculating As Boolean = False)

    Dim intIndex As Integer
    
    Dim blnUseRobustNET As Boolean
    Dim blnEnableRobustNETControls As Boolean
    
    Dim blnUseRobustNETIterative As Boolean

    If mRobustNETInProgress Then blnCalculating = True
    
    blnUseRobustNET = cChkBox(chkRobustNETEnabled)
    
    pbarRobustNET.Visible = blnUseRobustNET And blnCalculating
   
    lblNetAdjInitialSlope.Enabled = Not (blnUseRobustNET Or blnCalculating)
    txtNetAdjInitialSlope.Enabled = lblNetAdjInitialSlope.Enabled
    lblNetAdjInitialIntercept.Enabled = lblNetAdjInitialSlope.Enabled
    txtNetAdjInitialIntercept.Enabled = lblNetAdjInitialSlope.Enabled
    
    blnEnableRobustNETControls = blnUseRobustNET And Not blnCalculating
    fraNETSlopeRange.Enabled = blnEnableRobustNETControls
    fraNETInterceptRange.Enabled = blnEnableRobustNETControls
    fraMassShiftPPMRange.Enabled = blnEnableRobustNETControls
    
    blnUseRobustNETIterative = blnUseRobustNET
    
    txtRobustNETSlopeStart.Enabled = blnUseRobustNET
    txtRobustNETSlopeEnd.Enabled = blnUseRobustNET
    txtRobustNETSlopeIncrement.Enabled = blnUseRobustNETIterative
    
    txtRobustNETInterceptStart.Enabled = blnUseRobustNET
    txtRobustNETInterceptEnd.Enabled = blnUseRobustNET
    txtRobustNETInterceptIncrement.Enabled = blnUseRobustNETIterative
    
    txtRobustNETMassShiftPPMStart.Enabled = blnUseRobustNET
    txtRobustNETMassShiftPPMEnd.Enabled = blnUseRobustNET
    txtRobustNETMassShiftPPMIncrement.Enabled = blnUseRobustNETIterative
    
    For intIndex = 0 To 2
        lblRobustNETIterate(intIndex).Enabled = blnUseRobustNETIterative
    Next intIndex
    
    optRobustNETSlopeIncrement(0).Enabled = blnUseRobustNETIterative
    optRobustNETSlopeIncrement(1).Enabled = blnUseRobustNETIterative
    
    For intIndex = 0 To 3
        txtNetAdjWarningTol(intIndex).Enabled = Not blnUseRobustNET
    Next intIndex
    
    fraChargeStateForUMCs.Enabled = Not blnCalculating
    fraMiscellaneousAdvanced.Enabled = Not blnCalculating
    fraNetAdjLockers.Enabled = Not blnCalculating
        
    ' Charge state for UMC Selection
    chkCS(0).Enabled = Not cChkBox(chkCS(7))
    For intIndex = 1 To 6
        chkCS(intIndex).Enabled = chkCS(0).Enabled
    Next intIndex

    ' Update the NET range stats
    UpdateNetAdjRangeStats UMCNetAdjDef, lblNetAdjInitialNETStats
    
End Sub

Private Function FillTheGRID() As Boolean
    '----------------------------------------
    'fills GRID arrays with ID information
    'For each AMT, records the indices in ID() that matched the AMT
    '----------------------------------------
    Dim i As Long
    Dim GRIDCount As Long
    
    Dim DummyInd() As Long      'dummy array(empty) will allow us to
                                'sort only on one array
    Dim QSL As QSLong
    On Error GoTo err_FillTheGRID
    UpdateStatus "Loading data structures ..."
    
    If mUsingInternalStandards Then
        GRIDCount = UMCInternalStandards.Count
    Else
        GRIDCount = AMTCnt
    End If
    
    If IDCnt > 0 And GRIDCount > 0 Then
       ReDim GRID(GRIDCount)      'AMT arrays are 1-based
       For i = 0 To IDCnt - 1
           With GRID(ID(i))
               ReDim Preserve .Members(.Count)
               .Members(.Count) = i
               .Count = .Count + 1
           End With
       Next i
       
       'order members of each group on ID number
       For i = 0 To GRIDCount
           If GRID(i).Count > 1 Then
              Set QSL = New QSLong
              If Not QSL.QSAsc(GRID(i).Members, DummyInd) Then GoTo err_FillTheGRID
              Set QSL = Nothing
           End If
       Next i
       FillTheGRID = True
       UpdateStatus ""
    Else
       UpdateStatus "Data not found."
    End If
    Exit Function
    
err_FillTheGRID:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->FillTheGRID"
    
    Select Case Err.Number
    Case 7
       Call ClearIDArrays
       Call ClearTheGRID
       Call ClearPeakArrays
       UpdateStatus ""
       If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "System low on memory. Process aborted in recovery attempt.", vbOKOnly, glFGTU
       End If
    Case Else
       Call ClearIDArrays
       Call ClearTheGRID
       Call ClearPeakArrays
       UpdateStatus "Error loading data structures."
    End Select
End Function

Private Function GetElution(FN As Long) As Double
    '--------------------------------------------------
    'this function does not care are we using NET or RT
    '--------------------------------------------------
    VarVals(1) = FN
    GetElution = NETExprEva.ExprVal(VarVals())
End Function

Public Function GetNETAdjustmentIDCount() As Long
    GetNETAdjustmentIDCount = mCurrentIterationStats.IDMatchCount
End Function

Private Function GetRobustNETIterationRanges() As String
    Dim strMessage As String
    
    With UMCNetAdjDef
        strMessage = ""
        strMessage = strMessage & "Slope Start = " & DoubleToStringScientific(.RobustNETSlopeStart) & "; Slope End = " & DoubleToStringScientific(.RobustNETSlopeEnd)
        strMessage = strMessage & "; Intercept Start = " & DoubleToStringScientific(.RobustNETInterceptStart) & "; Intercept End = " & DoubleToStringScientific(.RobustNETInterceptEnd)
        strMessage = strMessage & "; Mass Shift PPM Start = " & Format(.RobustNETMassShiftPPMStart, "0.0") & "; Mass Shift PPM End = " & Format(.RobustNETMassShiftPPMEnd, "0.0")
    End With
    
    GetRobustNETIterationRanges = strMessage

End Function

Private Sub GetUMCPeak_At(ByVal UMCInd As Long)
'---------------------------------------------------------------
'adds peak from unique mass class UMCInd to peaks to be searched
'---------------------------------------------------------------
Dim dblMass As Double

On Error GoTo err_GetUMCPeak_At

' Use the class representative for the UMC's scan number
With GelUMC(CallerID).UMCs(UMCInd)
    If .ClassRepInd >= 0 Then
        PeakCnt = PeakCnt + 1
        UMCTotalPoints = UMCTotalPoints + .ClassCount
        
        Select Case .ClassRepType
        Case glCSType
            PeakScan(PeakCnt - 1) = GelData(CallerID).CSData(.ClassRepInd).ScanNumber
            PeakIsoFit(PeakCnt - 1) = 0
            dblMass = GelData(CallerID).CSData(.ClassMInd(.ClassRepInd)).AverageMW
        Case glIsoType
            PeakScan(PeakCnt - 1) = GelData(CallerID).IsoData(.ClassRepInd).ScanNumber
            PeakIsoFit(PeakCnt - 1) = GelData(CallerID).IsoData(.ClassRepInd).Fit
            ' MonroeMod: Now always using UMC Class MW rather than mass of most abundant member
            dblMass = .ClassMW
        End Select
       
        ' Apply a mass shift if mMassShiftPPM is non-zero
        If Round(mMassShiftPPM, 3) <> 0 Then
            dblMass = dblMass + PPMToMass(CDbl(mMassShiftPPM), dblMass)
        End If
        
        PeakMW(PeakCnt - 1) = dblMass
        PeakUMCInd(PeakCnt - 1) = UMCInd
    End If
End With

Exit Sub

err_GetUMCPeak_At:
If Err.Number = 9 Then
   ReDim Preserve PeakUMCInd(PeakCnt + 100)
   ReDim Preserve PeakMW(PeakCnt + 100)
   ReDim Preserve PeakScan(PeakCnt + 100)
   ReDim Preserve PeakIsoFit(PeakCnt + 100)
   Resume
End If
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

Private Sub IncrementRobustNETSetting(ByRef sngSetting As Single, ByVal eIncrementMode As UMCRobustNETIncrementConstants, ByVal sngIncrementOrFactor As Single)

    If eIncrementMode = UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear Then
        sngSetting = sngSetting + sngIncrementOrFactor
    Else
        ' Increase by a percentage of the original value
        ' Increment must be a number >= 1
        If sngIncrementOrFactor < 1 Then
            sngIncrementOrFactor = 1
        End If
    
        ' The following will increment sngSetting, even if it is negative
        sngSetting = sngSetting + CSng(Abs(sngSetting) * (sngIncrementOrFactor / 100))
        
    End If

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
        
        SetCheckBox chkRobustNETEnabled, .UseRobustNETAdjustment
        
        txtMWTol.Text = .MWTol
        Select Case .MWTolType
        Case gltPPM
             optTolType(0).Value = True
        Case gltABS
             optTolType(1).Value = True
        End Select
        txtNETTol.Text = Round(.NETTolIterative, NET_TOL_DIGITS_PRECISION)
        txtMinUMCCount.Text = .MinUMCCount
        txtMinScanRange.Text = .MinScanRange
        txtMaxUMCScansPct.Text = .MaxScanPct
        If .TopAbuPct >= 0 Then
           chkUMCUseTopAbu.Value = vbChecked
           If glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentAutoIncrementUMCTopAbuPct Then
               txtUMCAbuTopPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
               .TopAbuPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
           Else
               txtUMCAbuTopPct = .TopAbuPct
           End If
        Else
           txtUMCAbuTopPct.Text = -.TopAbuPct
           chkUMCUseTopAbu.Value = vbUnchecked
        End If

        For i = 0 To UBound(.PeakCSSelection)
            If .PeakCSSelection(i) Then
               chkCS(i).Value = vbChecked
            Else
               chkCS(i).Value = vbUnchecked
            End If
        Next i
        SetCheckBox chkEliminateConflictingIDs, .UseMultiIDMaxNETDist
        txtMultiIDMaxNETDist.Text = .MultiIDMaxNETDist
        txtNETFormula.Text = .NETFormula
        ValidateNETFormula
        
        SetCheckBox chkUseNETForID, .UseNET
        optIteStop(.IterationStopType).Value = True
        txtIteStopVal.Text = .IterationStopValue
        
        SetCheckBox chkDecMW, .IterationUseMWDec
        txtIteMWDec.Text = .IterationMWDec
        
        SetCheckBox chkDecNET, .IterationUseNETdec
        txtIteNETDec.Text = .IterationNETDec
        
        SetCheckBox chkAcceptLastIteration, .IterationAcceptLast
        
        txtNetAdjInitialSlope.Text = .InitialSlope
        txtNetAdjInitialIntercept.Text = .InitialIntercept
        
        ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
        '' SetCheckBox chkNetAdjUseLockers, .UseNetAdjLockers
        '' SetCheckBox chkNetAdjUseOldIfFailure, .UseOldNetAdjIfFailure
        '' txtNetAdjMinLockerMatchCount = .NetAdjLockerMinimumMatchCount
  
        txtRobustNETSlopeStart.Text = .RobustNETSlopeStart
        txtRobustNETSlopeEnd.Text = .RobustNETSlopeEnd
        If .RobustNETSlopeIncreaseMode = UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear Then
            optRobustNETSlopeIncrement(0).Value = True
        Else
            optRobustNETSlopeIncrement(1).Value = True
        End If
        txtRobustNETSlopeIncrement.Text = .RobustNETSlopeIncrement
        
        txtRobustNETInterceptStart.Text = .RobustNETInterceptStart
        txtRobustNETInterceptEnd.Text = .RobustNETInterceptEnd
        txtRobustNETInterceptIncrement.Text = .RobustNETInterceptIncrement
        
        txtRobustNETMassShiftPPMStart.Text = .RobustNETMassShiftPPMStart
        txtRobustNETMassShiftPPMEnd.Text = .RobustNETMassShiftPPMEnd
        txtRobustNETMassShiftPPMIncrement.Text = .RobustNETMassShiftPPMIncrement
            
    End With
        
    With cboPairsUMCsToUseForNETAdjustment
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
        
        txtNetAdjWarningTol(naswSlopeMinimum) = Trim(.NETSlopeExpectedMinimum)
        txtNetAdjWarningTol(naswSlopeMaximum) = Trim(.NETSlopeExpectedMaximum)
        txtNetAdjWarningTol(naswInterceptMinimum) = Trim(.NETInterceptExpectedMinimum)
        txtNetAdjWarningTol(naswInterceptMaximum) = Trim(.NETInterceptExpectedMaximum)
        
''        If .NETAdjustmentFinalNetTol < NET_RESOLUTION_SIMULATED_ANNEALING Then
''             .NETAdjustmentFinalNetTol = NET_RESOLUTION_SIMULATED_ANNEALING
''        End If
''        txtNETAdjFinalNETTol = .NETAdjustmentFinalNetTol
''        txtNETAdjMinimumNETMatchScore = Trim(.NETAdjustmentMinimumNETMatchScore)
    End With
    
    With glbPreferencesExpanded
        SetCheckBox chkUseN15AMTMasses, .NetAdjustmentUsesN15AMTMasses
        txtNetAdjMinHighNormalizedScore = .NetAdjustmentMinHighNormalizedScore
        txtNetAdjMinHighDiscriminantScore = .NetAdjustmentMinHighDiscriminantScore
    
''        With .RefineMSDataOptions
''            txtRefineMassCalibrationMaximumShift = .MassCalibrationMaximumShift
''            txtToleranceRefinementMinimumPeakHeight = .MinimumPeakHeight
''            txtToleranceRefinementPercentageOfMaxForWidth = .PercentageOfMaxForFindingWidth
''            txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks = .MinimumSignalToNoiseRatioForLowAbundancePeaks
''        End With
    End With
    
''    UpdateDynamicControls
    
    mFormInitialized = True
    
    tmrTimer.Interval = 500
    tmrTimer.Enabled = True
    
    mRequestUpdateRobustNETIterationCount = True
    
    DoEvents

    Exit Sub

InitializeFormErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->InitiationForm"
    Resume Next
End Sub

Private Sub InitializePlots()

    InitializeOnePlot ctlPlotSlopeVsScore, 1
    InitializeOnePlot ctlPlotNETRange, 3
    InitializeOnePlot ctlPlotMassErrorHistogram, 2
    InitializeOnePlot ctlPlotNETErrorHistogram, 2
    
End Sub

Private Sub InitializeOnePlot(ByRef ctlPlot As ctlSpectraPlotter, intSeriesCount As Integer)
    Dim intSeriesIndex As Integer
        
    Dim dblBlankDataX(1 To 1) As Double
    Dim dblBlankDataY(1 To 1) As Double
    
    With ctlPlot
        .EnableDisableDelayUpdating True
        .SetCurrentGroup 2
        .SetSeriesCount 0

        .SetCurrentGroup 1
        .SetSeriesCount intSeriesCount

        .SetChartType oc2dTypePlot, 1
        .SetStyleDataSymbol vbBlue, oc2dShapeNone, 1

        For intSeriesIndex = 1 To intSeriesCount
            .SetSeriesDataPointCount intSeriesIndex, 1
            
            .SetCurrentSeries intSeriesIndex
            .SetStyleDataSymbol vbBlue, oc2dShapeNone, 5
            
            .SetDataX intSeriesIndex, dblBlankDataX()
            .SetDataY intSeriesIndex, dblBlankDataY()
        Next intSeriesIndex

        ' Set the Tick Spacing the default
        .SetXAxisTickSpacing 1, True

        .SetXAxisAnnotationMethod oc2dAnnotateValues
        .SetXAxisAnnotationPlacement oc2dAnnotateAuto
        
        .SetYAxisAnnotationMethod oc2dAnnotateValues
        .SetYAxisAnnotationPlacement oc2dAnnotateAuto


        .EnableDisableDelayUpdating False
    End With
        
End Sub

Public Sub InitializeNETAdjustment()
'------------------------------------------------------------
'Load MT tag database data if necessary
'If CallerID is associated with MT tag database load that
' database if necessary; if CallerID is not associated with
' MT tag database load legacy database
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
          If (Len(CurrMTDatabase) > 0 Or Len(CurrLegacyMTDatabase) > 0) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            'MT tag data; we don't know is it appropriate; warn user
            WarnUserUnknownMassTags CallerID
         End If
         lblMTStatus.Caption = ConstructMTStatusText(True)
         
         ' Initialize the MT search object
         If Not CreateNewMTSearchObject() Then
            lblMTStatus.Caption = "Error creating search object."
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
   
   GetScanRange CallerID, ScanMin, ScanMax, ScanRange
   If ScanRange < 2 Then
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Scan range for this display could lead to unpredictable results.", vbOKOnly, glFGTU
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
        MsgBox "Error in elution calculation formula.", vbOKOnly, glFGTU
      Else
        AddToAnalysisHistory CallerID, "Error in elution calculation formula: " & UMCNetAdjDef.NETFormula
      End If
      txtNETFormula.SetFocus
   End If
End If
Me.MousePointer = vbDefault
Exit Sub

InitializeNETAdjustmentErrorHandler:
LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->InitializeNETAdjustment"
Resume Next

End Sub
                                                                        
Private Function InitializeInternalStandardSearch() As Boolean
    ' Returns True if success, False if error
    
    Dim dblSearchMasses() As Double
    Dim lngIndex As Long
    
    Dim objQSDouble As New QSDouble
    
On Error GoTo InitializeInternalStandardSearchErrorHandler

    Set objInternalStandardSearchUtil = New MWUtil
    
    With UMCInternalStandards
        If .Count > 0 Then
            ReDim dblSearchMasses(0 To .Count - 1)
            ReDim mInternalStandardIndexPointers(0 To .Count - 1)
            
            For lngIndex = 0 To .Count - 1
                dblSearchMasses(lngIndex) = .InternalStandards(lngIndex).MonoisotopicMass
                mInternalStandardIndexPointers(lngIndex) = lngIndex
            Next lngIndex
        Else
            ReDim dblSearchMasses(0)
            ReDim mInternalStandardIndexPointers(0)
        End If
    End With
            
    ' Need to sort dblSearchMasses before calling objInternalStandardSearchUtil.Fill()
    ' Use objQSLong for this
    If UMCInternalStandards.Count > 0 Then
        If objQSDouble.QSAsc(dblSearchMasses, mInternalStandardIndexPointers) Then
            InitializeInternalStandardSearch = objInternalStandardSearchUtil.Fill(dblSearchMasses())
        Else
            ' Error sorting
            Debug.Assert False
            InitializeInternalStandardSearch = False
        End If
    Else
        InitializeInternalStandardSearch = False
    End If
    
    Exit Function
    
InitializeInternalStandardSearchErrorHandler:
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->InitializeInternalStandardSearch"
    Set objInternalStandardSearchUtil = Nothing
    InitializeInternalStandardSearch = False
    
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
    
    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, 0, blnForceReload, True, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTStatus.Caption = ConstructMTStatusText(True)
    
        If Not CreateNewMTSearchObject() Then
           lblMTStatus.Caption = "Error creating search object."
        End If
    
    Else
        If blnDBConnectionError Then
            lblMTStatus.Caption = "Error loading MT tags: database connection error."
        Else
            lblMTStatus.Caption = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
        End If
    End If

End Sub

Private Sub LogNetAdjIteration(strLogText As String)
    Dim intFileNum As Integer
    Dim strPath As String
    
    On Error Resume Next
    
    strPath = App.Path & "\NetAdjIterationLog.txt"
    intFileNum = FreeFile
    
    Open strPath For Append As intFileNum
    Print #intFileNum, strLogText
    Close intFileNum

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
''
''Public Function ManualRefineMassCalibration(Optional blnOverrideValue As Boolean = False, Optional dblMassAdjustmentOverridePPM As Double = 0) As Boolean
''    ' If blnOverrideValue = True, then dblMassAdjustmentOverride is used
''    '  instead of the one given by txtMassCalibrationNewIncrementalAdjustment
''
''    Dim dblNewMassAdjustmentIncrementPPM As Double
''
''    If Not mFormInitialized Then
''        ManualRefineMassCalibration = False
''        Exit Function
''    End If
''
''    If blnOverrideValue Then
''        If GelSearchDef(CallerID).MassCalibrationInfo.AdjustmentHistoryCount > 0 Then
''            ' Undo any previous mass calibration adjustments when overriding auto adjustment
''            MassCalibrationRevertToOriginal CallerID, False, True, Me
''        End If
''
''        dblNewMassAdjustmentIncrementPPM = dblMassAdjustmentOverridePPM
''    Else
''        dblNewMassAdjustmentIncrementPPM = CDblSafe(txtMassCalibrationNewIncrementalAdjustment.Text)
''    End If
''
''    If dblNewMassAdjustmentIncrementPPM = 0 Then
''        ManualRefineMassCalibration = True
''        Exit Function
''    End If
''
''    mMassCalibrationInProgress = True
''    Me.MousePointer = vbHourglass
''    DoEvents
''
''    ManualRefineMassCalibration = MassCalibrationApplyBulkAdjustment(CallerID, dblNewMassAdjustmentIncrementPPM, gltPPM, True, 0, Me)
''
''    Me.MousePointer = vbDefault
''    mMassCalibrationInProgress = False
''
''End Function


Private Sub MarkIDsWithBadNET()
'-----------------------------------------------------------------------
'sets state of all identifications with bad NET numbers to STATE_BAD_NET
'-----------------------------------------------------------------------
Dim i As Long
On Error Resume Next
UpdateStatus "Eliminating IDs with bad elution ..."
If Not IDsAreInternalStandards Then
    For i = 0 To IDCnt - 1
        If AMTData(ID(i)).NET < 0 Or AMTData(ID(i)).NET > 1 Then
           IDState(i) = IDState(i) + STATE_BAD_NET
        End If
    Next i
End If
UpdateStatus ""
End Sub


Private Sub PickParameters()
    Call txtMinUMCCount_LostFocus
    Call txtMinScanRange_LostFocus
    Call txtMaxUMCScansPct_LostFocus
    Call txtUMCAbuTopPct_LostFocus
    Call txtMWTol_LostFocus
    Call txtNETTol_LostFocus
    Call txtMultiIDMaxNETDist_LostFocus
    Call txtIteMWDec_LostFocus
    Call txtIteNETDec_LostFocus
    Call txtIteStopVal_LostFocus
    Call txtIteMWDec_LostFocus
    Call txtIteNETDec_LostFocus
    Call txtNETFormula_LostFocus
    
    Call txtNetAdjMinHighDiscriminantScore_LostFocus
    Call txtNetAdjMinHighNormalizedScore_LostFocus
    Call txtNetAdjMinIDCount_LostFocus
    
''    Call txtNetAdjMinLockerMatchCount_LostFocus
    
    Call txtRobustNETSlopeEnd_LostFocus
    Call txtRobustNETSlopeIncrement_LostFocus
    Call txtRobustNETSlopeStart_LostFocus
    
    Call txtRobustNETInterceptEnd_LostFocus
    Call txtRobustNETInterceptIncrement_LostFocus
    Call txtRobustNETInterceptStart_LostFocus
    
    Call txtRobustNETMassShiftPPMEnd_LostFocus
    Call txtRobustNETMassShiftPPMIncrement_LostFocus
    Call txtRobustNETMassShiftPPMStart_LostFocus
       
''    Call txtRobustNETAnnealSteps_LostFocus
''    Call txtRobustNETAnnealTrialsPerStep_LostFocus
''    Call txtRobustNETAnnealMaxSwapsPerStep_LostFocus
''    Call txtRobustNETAnnealTemperatureReductionFactor_LostFocus
''    Call txtNETAdjMinimumNETMatchScore_LostFocus
       
    Call txtToleranceRefinementMinimumPeakHeight_LostFocus
    Call txtToleranceRefinementPercentageOfMaxForWidth_Lostfocus
    Call txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks_LostFocus
    
    Call txtNetAdjInitialIntercept_LostFocus
    Call txtNetAdjInitialSlope_LostFocus
    
    Call txtNetAdjWarningTol_LostFocus(0)
    Call txtNetAdjWarningTol_LostFocus(1)
    Call txtNetAdjWarningTol_LostFocus(2)
    Call txtNetAdjWarningTol_LostFocus(3)
    
End Sub

Private Sub PopulateComboBoxes()
    With cboPairsUMCsToUseForNETAdjustment
        .Clear
        .AddItem "All UMC's, regardless of pair or light/heavy status", punaPairedAndUnpaired
        .AddItem "Unpaired UMC's only", punaUnpairedOnly
        .AddItem "Unpaired UMC's and light members of paired UMC's", punaUnpairedPlusPairedLight
        .AddItem "Paired UMC's, both light and heavy members", punaPairedAll
        .AddItem "Paired UMC's, light members only", punaPairedLight
        .AddItem "Paired UMC's, heavy members only", punaPairedHeavy
        .ListIndex = punaPairsUMCNetAdjustmentConstants.punaUnpairedPlusPairedLight
    End With

End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    
    fraPlots.Top = 120

    lngDesiredValue = Me.ScaleWidth - fraPlots.Left - 120
    If lngDesiredValue < 5000 Then lngDesiredValue = 5000
    fraPlots.width = lngDesiredValue
    
    lngDesiredValue = Me.ScaleHeight - fraPlots.Top - 1200
    If lngDesiredValue < 5000 Then lngDesiredValue = 5000
    fraPlots.Height = lngDesiredValue
    
    With ctlPlotSlopeVsScore
        .Left = 120
        .Top = 360
        .width = (fraPlots.width / 2) - 160
        .Height = (fraPlots.Height / 2) - 360
        
        lblPlotSlopeVsScore.Top = .Top - 240
        lblPlotSlopeVsScore.width = .width
        lblPlotSlopeVsScore.Left = .Left
    End With
    
    With ctlPlotNETRange
        .Left = ctlPlotSlopeVsScore.Left + ctlPlotSlopeVsScore.width + 120
        .Top = ctlPlotSlopeVsScore.Top
        .width = ctlPlotSlopeVsScore.width
        .Height = ctlPlotSlopeVsScore.Height
    
        lblPlotNETRange.Top = .Top - 240
        lblPlotNETRange.width = .width
        lblPlotNETRange.Left = .Left
    End With
    
    With ctlPlotMassErrorHistogram
        .Left = ctlPlotSlopeVsScore.Left
        .Top = ctlPlotSlopeVsScore.Top + ctlPlotSlopeVsScore.Height + 360
        .width = ctlPlotSlopeVsScore.width
        .Height = ctlPlotSlopeVsScore.Height - 120
    
        lblPlotMassErrorHistogram.Top = .Top - 240
        lblPlotMassErrorHistogram.width = .width
        lblPlotMassErrorHistogram.Left = .Left
    End With
    
    With ctlPlotNETErrorHistogram
        .Left = ctlPlotNETRange.Left
        .Top = ctlPlotMassErrorHistogram.Top
        .width = ctlPlotNETRange.width
        .Height = ctlPlotMassErrorHistogram.Height
        
        lblPlotNETErrorHistogram.Top = .Top - 240
        lblPlotNETErrorHistogram.width = .width
        lblPlotNETErrorHistogram.Left = .Left
    End With
    
End Sub

Private Function PredictRobustNETIterationCount() As Long

    Dim lngIterationCount As Long
    Dim sngSlope As Single, sngIntercept As Single, sngMassShiftPPM As Single
    
    lngIterationCount = 0
    sngMassShiftPPM = UMCNetAdjDef.RobustNETMassShiftPPMStart
    Do While sngMassShiftPPM <= UMCNetAdjDef.RobustNETMassShiftPPMEnd
        sngSlope = UMCNetAdjDef.RobustNETSlopeStart
        Do While sngSlope <= UMCNetAdjDef.RobustNETSlopeEnd
            sngIntercept = UMCNetAdjDef.RobustNETInterceptStart
            Do While sngIntercept <= UMCNetAdjDef.RobustNETInterceptEnd
                lngIterationCount = lngIterationCount + 1
                If lngIterationCount >= MAX_ROBUST_NET_ITERATION_COUNT Then Exit Do
                IncrementRobustNETSetting sngIntercept, UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear, UMCNetAdjDef.RobustNETInterceptIncrement
            Loop
            If lngIterationCount >= MAX_ROBUST_NET_ITERATION_COUNT Then Exit Do
            IncrementRobustNETSetting sngSlope, UMCNetAdjDef.RobustNETSlopeIncreaseMode, UMCNetAdjDef.RobustNETSlopeIncrement
        Loop
        If lngIterationCount >= MAX_ROBUST_NET_ITERATION_COUNT Then Exit Do
        IncrementRobustNETSetting sngMassShiftPPM, UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear, UMCNetAdjDef.RobustNETMassShiftPPMIncrement
    Loop

    lblRobustNETPredictedIterationCount = "Predicted Iteration Count: " & Trim(lngIterationCount)
    
    PredictRobustNETIterationCount = lngIterationCount
End Function

Private Sub RecordIterationStats(ByRef udtIterationStats As udtIterationStatsType)
    ' Note: The converse of this function is RestoreIterationStats
    
    With udtIterationStats
        .MassShiftPPM = mMassShiftPPM
    
        .slope = AdjSlp
        .intercept = AdjInt
        .Deviation = AdjAvD

        .InitialSlope = UMCNetAdjDef.InitialSlope
        .InitialIntercept = UMCNetAdjDef.InitialIntercept
    
        If UMCNetAdjDef.TopAbuPct < 0 Or UMCNetAdjDef.TopAbuPct > 100 Then
            .UMCTopAbuPct = 100
        Else
            .UMCTopAbuPct = UMCNetAdjDef.TopAbuPct
        End If
        
        .IterationCount = mIterationCount
        
        .FinalNETTol = UMCNetAdjDef.NETTolIterative
        .FinalMWTol = UMCNetAdjDef.MWTol
        .FinalMWTolType = UMCNetAdjDef.MWTolType

        .UMCSegmentCntWithLowUMCCnt = mUMCSegmentCntWithLowUMCCnt
        .UMCCntAddedSinceLowSegmentCount = mUMCCntAddedSinceLowSegmentCount
        
        .IDMatchCount = IDCnt
        
        .UMCCountSearched = PeakCnt
        .UMCMatchCount = PeakCntWithMatches
        
        .PointCountSearched = UMCTotalPoints
        .PointMatchCount = UMCTotalPointsMatching
        
        .RobustNET = mRobustNETInProgress
        .UseInternalStandards = mUsingInternalStandards
        
        If .IDMatchCount < mEffectiveMinIDCount Then
            If .IDMatchCount > 10 Then
                .Valid = False
            End If
            .Valid = False
        Else
            .Valid = True
        End If
        
    End With

    ComputeScoreForIterationStats udtIterationStats
    
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
    ts.WriteLine "Reporting NET adjustment(based on MT tag NETs) of Unique Mass Classes"
    ts.WriteLine
    ts.WriteLine "Slope: " & AdjSlp
    ts.WriteLine "Intercept: " & AdjInt
    ts.WriteLine "Average Deviation: " & AdjAvD
    ts.WriteLine
    ts.WriteLine "ID" & glARG_SEP & "ID_NET" & glARG_SEP & "Scan"
    If IDsAreInternalStandards Then
        With UMCInternalStandards
            For i = 0 To IDCnt - 1
                With .InternalStandards(ID(i))
                    ts.WriteLine .SeqID & glARG_SEP & .NET & glARG_SEP & IDScan(i)
                End With
            Next i
        End With
    Else
        For i = 0 To IDCnt - 1
            ts.WriteLine AMTData(ID(i)).ID & glARG_SEP & AMTData(ID(i)).NET & glARG_SEP & IDScan(i)
        Next i
    End If
    
    ts.Close
    Set fso = Nothing
    UpdateStatus ""
    DoEvents
    frmDataInfo.Tag = "AdjNET"
    frmDataInfo.Show vbModal
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
    Rep = Rep & UMCNetAdjDef.NETTolIterative & vbCrLf
    Rep = Rep & "UMC count used: " & PeakCnt & " - Matches: " & IDCnt & vbCrLf
    If AdjAvD >= 0 Then
       Rep = Rep & "Slope: " & AdjSlp & vbCrLf
       Rep = Rep & "Intercept: " & AdjInt & vbCrLf
       Rep = Rep & "Average Deviation: " & AdjAvD & vbCrLf
       Rep = Rep & "NET Min: " & Round(AdjSlp * ScanMin + AdjInt, 4) & ", NET Max: " & Round(AdjSlp * ScanMax + AdjInt, 4) & vbCrLf
    
       'Debug.Print IterationStep & ", " & IDCnt & "," & Round(AdjSlp * ScanMin + AdjInt, 5) & "," & Round(AdjSlp * ScanMax + AdjInt, 5)
    Else
       Rep = "Error or insuficient information for NET adjustment calculation." & vbCrLf
    End If
    AppendToIterationReport Rep & vbCrLf
    
    'If IterationStep = 1 Then Debug.Print
    'Debug.Print IterationStep & vbTab & PeakCnt & vbTab & IDCnt & vbTab & AdjAvD
End Sub
    
    Private Sub ReportIterationStepChange(PrevNET1 As Double, PrevNET2 As Double, _
                                          CurrNET1 As Double, CurrNET2 As Double)
    '-------------------------------------------------------------------------------
    'add results of an iteration step on bottom of the rich text box
    '-------------------------------------------------------------------------------
    Dim Rep As String
    On Error Resume Next
    If IDCnt < 2 Then Rep = "Insufficient information for next iteration." & vbCrLf
    Rep = Rep & "Previous iteration NET min/max value: " & Format$(PrevNET1, "0.0000") _
           & "; " & Format$(PrevNET2, "0.0000") & vbCrLf
    Rep = Rep & "Last iteration NET min/max value: " & Format$(CurrNET1, "0.0000") _
           & "; " & Format$(CurrNET2, "0.0000") & vbCrLf
    AppendToIterationReport Rep
End Sub

Private Sub ReportStop()
    Dim Rep As String
    On Error Resume Next
    Rep = "Iteration cancelled by user." & vbCrLf
    AppendToIterationReport Rep
End Sub

Private Sub ReportPause()
    Dim Rep As String
    On Error Resume Next
    Rep = "Iteration paused by user. Press Continue to resume" & vbCrLf
    AppendToIterationReport Rep
    UpdateStatus "Paused"
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
            lblMTStatus.Caption = "Error creating search object."
            ResetProcedure = False
        Else
            ResetProcedure = True
        End If
    Else
        ResetProcedure = True
    End If

End Function

Private Sub ResetSlopeAndInterceptToDefault()
    If Not GelAnalysis(CallerID) Is Nothing Then
        With GelAnalysis(CallerID)
            .GANET_Slope = UMCNetAdjDef.InitialSlope
            .GANET_Intercept = UMCNetAdjDef.InitialIntercept
        End With
    End If
    
    UMCNetAdjDef.NETFormula = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
    txtNETFormula.Text = UMCNetAdjDef.NETFormula
    
    ValidateNETFormula
End Sub

Public Sub ResetToGenericNetAdjSettings()
    ResetSlopeAndInterceptToDefault
        
    txtNETTol = Round(glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol, NET_TOL_DIGITS_PRECISION)
    UMCNetAdjDef.NETTolIterative = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol
    
    ' If auto-adjusting UMC Top Abu Pct, then set txtUMCAbuTopPct to the value defined by NETAdjustmentUMCTopAbuPctInitial
    If glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentAutoIncrementUMCTopAbuPct Then
        txtUMCAbuTopPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
        UMCNetAdjDef.TopAbuPct = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentUMCTopAbuPctInitial
    Else
        txtUMCAbuTopPct = UMCNetAdjDef.TopAbuPct
    End If
    
End Sub

Private Sub RestoreIterationStats(ByRef udtIterationStats As udtIterationStatsType)
    ' Note: The converse of this function is RecordIterationStats
    
    With udtIterationStats
        mMassShiftPPM = .MassShiftPPM
        
        AdjSlp = .slope
        AdjInt = .intercept
        AdjAvD = .Deviation
    
        UpdateInitialSlopeAndIntercept .InitialSlope, .InitialIntercept

        UMCNetAdjDef.TopAbuPct = .UMCTopAbuPct
        txtUMCAbuTopPct = Trim(UMCNetAdjDef.TopAbuPct)
        
        mIterationCount = .IterationCount
        
        If UMCNetAdjDef.UseNET Then
            UMCNetAdjDef.NETTolIterative = .FinalNETTol
            txtNETTol = Round(.FinalNETTol, NET_TOL_DIGITS_PRECISION)
        End If
        
        UMCNetAdjDef.MWTol = .FinalMWTol
        txtMWTol = .FinalMWTol
        
        UMCNetAdjDef.MWTolType = .FinalMWTolType
        optTolType(UMCNetAdjDef.MWTolType).Value = True
        
        mUMCSegmentCntWithLowUMCCnt = .UMCSegmentCntWithLowUMCCnt
        mUMCCntAddedSinceLowSegmentCount = .UMCCntAddedSinceLowSegmentCount
        
        ' Do Not Update: IDCnt = .IDMatchCount
        
        ' Do Not Update: PeakCnt = .UMCCountSearched
        PeakCntWithMatches = .UMCMatchCount
        
        ' Do Not Update: UMCTotalPoints = .PointCountSearched
        UMCTotalPointsMatching = .PointMatchCount
        
        
    End With

    ' Store the final NET formula
    If Not GelAnalysis(CallerID) Is Nothing Then
        GelAnalysis(CallerID).GANET_Fit = AdjAvD
        GelAnalysis(CallerID).GANET_Slope = AdjSlp
        GelAnalysis(CallerID).GANET_Intercept = AdjInt
        
        GelStatus(CallerID).Dirty = True
    End If
     
    UMCNetAdjDef.NETFormula = ConstructNETFormula(AdjSlp, AdjInt)
    txtNETFormula.Text = UMCNetAdjDef.NETFormula
    ValidateNETFormula
    
    mCurrentIterationStats = udtIterationStats

End Sub

Private Function ScoreIDs() As Boolean
'-------------------------------------------------------------------
'score ids and set states of IDs that don't make top MaxIDToUse
'to STATE_OUTSCORED
'Score=[Log10(Peak Abundance) - Fit] for Isotopic peaks and
'      [Log10(Peak Abundance)]       for Charge State peaks
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
        Score(i) = (Log(IDMatchAbu(i)) / Log(10)) - IDIsoFit(i)
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
LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->ScoreIDs"
End Function

Private Function SearchForUMCsMatchingPMTs(blnForceSelectUMCs As Boolean) As Boolean
    ' Returns True if the search succeeded (even if there are no matches)
    ' Returns False if an error
    
    Static lngTopAbuPct As Long
    
    Dim i As Integer
    
    Dim blnUpdateUMCs As Boolean
    Dim blnProceed As Boolean
    
    mEffectiveMinIDCount = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount
    
    mNetAdjFailed = False
    
    If blnForceSelectUMCs Or lngTopAbuPct <> UMCNetAdjDef.TopAbuPct Then
        blnUpdateUMCs = True
    Else
        blnUpdateUMCs = False
    End If
    
    blnProceed = False
    If blnUpdateUMCs Then
        ' Determine which UMC's and which Peaks are to be used
        Call ResetProcedure
        If LinearNETAlignmentSelectUMCToUse(CallerID, UseUMC(), mUMCCntAddedSinceLowSegmentCount, mUMCSegmentCntWithLowUMCCnt) > 0 Then
            If SelectPeaksToUse() > 0 Then
                blnProceed = True
                lngTopAbuPct = UMCNetAdjDef.TopAbuPct
            End If
        End If
    Else
        ' Simply clear the ID arrays and reset the Hit bit
        ClearIDArrays
        For i = 0 To GelUMC(CallerID).UMCCnt - 1
           With GelUMC(CallerID).UMCs(i)
              .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
           End With
        Next i
        blnProceed = True
    End If
    
    If blnProceed Then
        If mUsingInternalStandards Then
            If UMCInternalStandards.Count < 1 Then
                SetNetAdjFailed
            End If
            If Not SearchInternalStandards() Then
                SetNetAdjFailed
            End If
        Else
            If AMTCnt < 1 Then
                SetNetAdjFailed
                bStop = True
            End If
            If UMCNetAdjDef.UseNET Then
                If Not SearchMassTagsMWNET() Then
                    SetNetAdjFailed
                End If
            Else
                If Not SearchMassTagsMW() Then
                    SetNetAdjFailed
                End If
            End If
        End If
      
        If Not mNetAdjFailed Then
            If IDCnt < 1 Then
                ' No matches were found
                ReDim GRID(0)
                AdjAvD = 1
            Else
                If FillTheGRID() Then
                    If SelectIdentifications() > 1 Then
                        ' Do not call CalculateSlopeIntercept, but do compute the average deviation
                        Call CalculateAvgDev
                    Else
                        ' Only one match was found
                        AdjAvD = 1
                    End If
                End If
            End If
        End If
    Else
        SetNetAdjFailed
    End If
    
    SearchForUMCsMatchingPMTs = Not mNetAdjFailed
    
End Function

Private Function SearchMassTagsMW() As Boolean
'----------------------------------------------------------
'searches MT tags for matching masses and returns True if
'OK, False if any error or user canceled the whole thing
'----------------------------------------------------------
Dim i As Long, j As Long
Dim eResponse As VbMsgBoxResult
Dim TmpCnt As Long
Dim Hits() As Long
Dim MWAbsErr As Double

On Error GoTo err_SearchMassTagsMW

UpdateStatus "Searching for MT tags ..."
' These arrays are dimensioned to hold MAX_ID_CNT items
' If too many identifications are found, the Error Number 9 will be raised by ??
ReDim ID(MAX_ID_CNT - 1)            'prepare for the worst case
ReDim IDMatchingUMCInd(MAX_ID_CNT - 1)
ReDim IDMatchAbu(MAX_ID_CNT - 1)
ReDim IDIsoFit(MAX_ID_CNT - 1)
ReDim IDScan(MAX_ID_CNT - 1)
ReDim IDState(MAX_ID_CNT - 1)
IDsAreInternalStandards = False

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
    Case glAMT_RT_or_PNET
        TmpCnt = GetMTHits2(PeakMW(i), MWAbsErr, -1, -1, Hits())
    End Select
    If TmpCnt > 0 Then
        With GelUMC(CallerID).UMCs(PeakUMCInd(i))
            PeakCntWithMatches = PeakCntWithMatches + 1
            UMCTotalPointsMatching = UMCTotalPointsMatching + .ClassCount
            
            For j = 0 To TmpCnt - 1
                ID(IDCnt) = Hits(j)
                IDMatchingUMCInd(IDCnt) = PeakUMCInd(i)
                IDMatchAbu(IDCnt) = .ClassAbundance
                IDIsoFit(IDCnt) = PeakIsoFit(i)
                IDScan(IDCnt) = PeakScan(i)
                IDState(IDCnt) = 0
                IDCnt = IDCnt + 1   'if reaches limit we will have correct results by doing increase at the end
            Next j
            .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
        End With
    End If
Next i

If IDCnt > 0 Then
   ReDim Preserve ID(IDCnt - 1)
   ReDim Preserve IDMatchingUMCInd(IDCnt - 1)
   ReDim Preserve IDMatchAbu(IDCnt - 1)
   ReDim Preserve IDIsoFit(IDCnt - 1)
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
        eResponse = MsgBox("Too many possible identifications detected.  " _
                   & "To proceed with the first " & MAX_ID_CNT & _
                   " identifications select OK.", vbOKCancel, glFGTU)
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
    MsgBox "System low on memory. Process aborted in recovery attempt.", vbOKOnly, glFGTU
Case Else
    UpdateStatus "Error searching for MT tags."
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->SearchMassTagsMW"
End Select
End Function


Private Function SearchMassTagsMWNET() As Boolean
'----------------------------------------------------------
'searches MT tags for matching masses and returns True if
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

UpdateStatus "Searching for MT tags ..."
ReDim ID(MAX_ID_CNT - 1)            'prepare for the worst case
ReDim IDMatchingUMCInd(MAX_ID_CNT - 1)
ReDim IDMatchAbu(MAX_ID_CNT - 1)
ReDim IDIsoFit(MAX_ID_CNT - 1)
ReDim IDScan(MAX_ID_CNT - 1)
ReDim IDState(MAX_ID_CNT - 1)
IDsAreInternalStandards = False

For i = 0 To PeakCnt - 1
    Select Case UMCNetAdjDef.MWTolType
    Case gltPPM
        MWAbsErr = PeakMW(i) * UMCNetAdjDef.MWTol * glPPM
    Case gltABS
        MWAbsErr = UMCNetAdjDef.MWTol
    End Select
    Select Case UMCNetAdjDef.NETorRT
    Case glAMT_NET
        TmpCnt = GetMTHits1(PeakMW(i), MWAbsErr, ConvertScanToNET(PeakScan(i)), UMCNetAdjDef.NETTolIterative, Hits())
    Case glAMT_RT_or_PNET
        TmpCnt = GetMTHits2(PeakMW(i), MWAbsErr, ConvertScanToNET(PeakScan(i)), UMCNetAdjDef.NETTolIterative, Hits())
    End Select
    If TmpCnt > 0 Then
    
       ' MonroeMod: The following implements the option "Do not use peaks pointing to multiple IDs on NET distance of more than ..."
       If UMCNetAdjDef.UseMultiIDMaxNETDist And TmpCnt > 1 Then
          ' Examine the NET values for the Hits
          ' If the range of NET values is > .MultiIDMaxNETDist then do not use this match
          AMTNETMin = AMTData(Hits(0)).NET
          AMTNETMax = AMTData(Hits(0)).NET
          For j = 1 To TmpCnt - 1
             If AMTData(j).NET < AMTNETMin Then AMTNETMin = AMTData(j).NET
             If AMTData(j).NET > AMTNETMax Then AMTNETMax = AMTData(j).NET
          Next j
          
          If Abs(AMTNETMax - AMTNETMin) > UMCNetAdjDef.MultiIDMaxNETDist Then
            TmpCnt = 0
          End If
       End If
       
       If TmpCnt > 0 Then
            With GelUMC(CallerID).UMCs(PeakUMCInd(i))
                PeakCntWithMatches = PeakCntWithMatches + 1
                UMCTotalPointsMatching = UMCTotalPointsMatching + .ClassCount
                
                For j = 0 To TmpCnt - 1
                    If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Then
                        If AMTData(Hits(j)).HighNormalizedScore >= mMTMinimumHighNormalizedScore And _
                           AMTData(Hits(j)).HighDiscriminantScore >= mMTMinimumHighDiscriminantScore Then
                            blnAddMassTag = True
                        Else
                            blnAddMassTag = False
                        End If
                    Else
                        blnAddMassTag = True
                    End If
                    
                    If blnAddMassTag Then
                        ID(IDCnt) = Hits(j)
                        IDMatchingUMCInd(IDCnt) = PeakUMCInd(i)
                        IDMatchAbu(IDCnt) = .ClassAbundance
                        IDIsoFit(IDCnt) = PeakIsoFit(i)
                        IDScan(IDCnt) = PeakScan(i)
                        IDState(IDCnt) = 0
                        IDCnt = IDCnt + 1   'if reaches limit we will have correct results by doing increase at the end
                    End If
                Next j
            
               .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
            End With
            
       End If
    End If
Next i

If IDCnt > 0 Then
   ReDim Preserve ID(IDCnt - 1)
   ReDim Preserve IDMatchingUMCInd(IDCnt - 1)
   ReDim Preserve IDMatchAbu(IDCnt - 1)
   ReDim Preserve IDIsoFit(IDCnt - 1)
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
    If Not (glbPreferencesExpanded.AutoAnalysisStatus.Enabled Or mRobustNETInProgress) Then
        eResponse = MsgBox("Too many possible identifications detected.  " _
                   & "To proceed with the first " & MAX_ID_CNT & _
                   " identifications select OK.", vbOKCancel, glFGTU)
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
    MsgBox "System low on memory. Process aborted in recovery attempt.", vbOKOnly, glFGTU
Case Else
    UpdateStatus "Error searching for MT tags."
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->SearchMassTagsMWNET"
End Select
End Function

Private Function SearchInternalStandards() As Boolean
'----------------------------------------------------------
'Searches Internal Standards for matching masses
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
    
On Error GoTo SearchInternalStandardsErrorHandler
    
    UpdateStatus "Searching for matching Internal Standards ..."
    ReDim ID(PeakCnt)                           ' Reserve space for highest possible number of matches
    ReDim IDMatchingUMCInd(PeakCnt)
    ReDim IDMatchAbu(PeakCnt)
    ReDim IDIsoFit(PeakCnt)
    ReDim IDScan(PeakCnt)
    ReDim IDState(PeakCnt)
    IDsAreInternalStandards = True

    For i = 0 To PeakCnt - 1
        Select Case UMCNetAdjDef.MWTolType
        Case gltPPM
            MWAbsErr = PeakMW(i) * UMCNetAdjDef.MWTol * glPPM
        Case gltABS
            MWAbsErr = UMCNetAdjDef.MWTol
        End Select
        
        dblNetToMatch = ConvertScanToNET(PeakScan(i))
        
        lngHitCount = 0
        
        ' Note: objInternalStandardSearchUtil should have been filled with the Internal Standard masses, sorted ascending
        If objInternalStandardSearchUtil.FindIndexRange(PeakMW(i), MWAbsErr, Ind1, Ind2) Then
            If Ind2 >= Ind1 Then
                lngHitCountDimmed = 100
                ReDim Hits(lngHitCountDimmed)
                
                With UMCInternalStandards
                    For j = Ind1 To Ind2
                        If ((Abs(dblNetToMatch - .InternalStandards(mInternalStandardIndexPointers(j)).NET) <= UMCNetAdjDef.NETTolIterative)) Or _
                            UMCNetAdjDef.NETTolIterative < 0 Then
                            ' Within NET Tolerance (or tolerance is negative)
                            ' See if the charge is valid
                            
                            blnValidHit = True
                            With .InternalStandards(mInternalStandardIndexPointers(j))
                                ' Make sure at least one of the charges for this Internal Standard is present in the UMC
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
                With UMCInternalStandards
                    dblNETMin = .InternalStandards(Hits(0)).NET
                    dblNETMax = .InternalStandards(Hits(0)).NET
                    For j = 1 To lngHitCount - 1
                        If .InternalStandards(Hits(j)).NET < dblNETMin Then dblNETMin = .InternalStandards(Hits(j)).NET
                        If .InternalStandards(Hits(j)).NET > dblNETMax Then dblNETMax = .InternalStandards(Hits(j)).NET
                    Next j
                End With
                
                If Abs(dblNETMax - dblNETMin) > UMCNetAdjDef.MultiIDMaxNETDist Then
                    lngHitCount = 0
                End If
            End If
            
            If lngHitCount > 0 Then
                With GelUMC(CallerID).UMCs(PeakUMCInd(i))
                    PeakCntWithMatches = PeakCntWithMatches + 1
                    UMCTotalPointsMatching = UMCTotalPointsMatching + .ClassCount
                
                    For j = 0 To lngHitCount - 1
                        IDMatchingUMCInd(IDCnt) = PeakUMCInd(i)
                        IDMatchAbu(IDCnt) = .ClassAbundance
                        IDIsoFit(IDCnt) = PeakIsoFit(i)
                        ID(IDCnt) = Hits(j)
                        IDState(IDCnt) = 0
                        IDScan(IDCnt) = PeakScan(i)
                        IDCnt = IDCnt + 1
                    Next j
                    
                    .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
                    .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_NET_ADJ_LOCKER_HIT
                End With
            End If
        End If
    Next i
    
    If IDCnt > 0 Then
       ReDim Preserve ID(IDCnt - 1)
       ReDim Preserve IDMatchingUMCInd(IDCnt - 1)
       ReDim Preserve IDMatchAbu(IDCnt - 1)
       ReDim Preserve IDIsoFit(IDCnt - 1)
       ReDim Preserve IDScan(IDCnt - 1)
       ReDim Preserve IDState(IDCnt - 1)
    Else
       Call ClearIDArrays
    End If
    
    UpdateStatus "Possible identifications: " & IDCnt
    SearchInternalStandards = True
Exit Function

SearchInternalStandardsErrorHandler:
    UpdateStatus "Error searching for Internal Standards."
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->SearchInternalStandards"
    SearchInternalStandards = False

End Function

Private Function SelectPeaksToUse() As Long
'------------------------------------------------------------------
'selects peaks that will be used to correct NET based on specified
'criteria; returns number of it; -1 on any error
'NOTE: peaks are selected only from UMCs marked to be used
'Furthermore, we're only selecting one point from each UMC, not every point in every UMC
'------------------------------------------------------------------
Dim i As Long
On Error GoTo exit_SelectPeaksToUse
SelectPeaksToUse = -1
For i = 0 To UBound(UseUMC)
    If UseUMC(i) Then
       ' Always use UMCNetConstants.UMCNetAt
        Call GetUMCPeak_At(i)
    End If
Next i
If PeakCnt > 0 Then
   ReDim Preserve PeakUMCInd(PeakCnt - 1)
   ReDim Preserve PeakMW(PeakCnt - 1)
   ReDim Preserve PeakScan(PeakCnt - 1)
   ReDim Preserve PeakIsoFit(PeakCnt - 1)
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

Private Sub SetNetAdjFailed()
     mNetAdjFailed = True
     mCurrentIterationStats.IDMatchCount = IDCnt
End Sub

Private Sub ShowHideControls(ByVal blnCalculating As Boolean)

If mRobustNETInProgress Then blnCalculating = True

If blnCalculating Then
    cmdPause.Caption = CMD_CAPTION_PAUSE
End If

cmdStart.Visible = Not blnCalculating
cmdReset.Visible = Not blnCalculating
cmdUseDefaults.Visible = Not blnCalculating

chkRobustNETEnabled.Enabled = Not blnCalculating

EnableDisableControls blnCalculating

End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuFExportToMTDB.Visible = blnVisible
    mnuMTLoad.Visible = blnVisible
End Sub

''Private Sub StartManualRefineMassCalibration()
''
''    If mFormInitialized And Not mMassCalibrationInProgress Then
''        mMassCalibrationInProgress = True
''        ManualRefineMassCalibration
''        mMassCalibrationInProgress = False
''    End If
''
''    UpdateDynamicControls
''    txtMassCalibrationNewIncrementalAdjustment.Text = "0"
''
''End Sub
''
''Private Sub StartMassCalibrationRevert()
''
''    If mFormInitialized And Not mMassCalibrationInProgress Then
''        Me.MousePointer = vbHourglass
''        mMassCalibrationInProgress = True
''        DoEvents
''
''        MassCalibrationRevertToOriginal CallerID, True, True, Me
''
''        mMassCalibrationInProgress = False
''        Me.MousePointer = vbDefault
''    End If
''
''    UpdateDynamicControls
''    txtMassCalibrationNewIncrementalAdjustment.Text = "0"
''
''End Sub

Private Sub TogglePlotFramePosition(blnForceOffset As Boolean)
    
    If fraPlots.Left = PLOT_FRAME_POSITION_VISIBLE Or blnForceOffset Then
        fraPlots.Left = PLOT_FRAME_POSITION_OFFSET
    Else
        fraPlots.Left = PLOT_FRAME_POSITION_VISIBLE
    End If
    
    If Me.WindowState = vbNormal Then
        Me.width = PLOT_FRAME_POSITION_OFFSET - 120
        Me.Height = 9300
    End If
    
    PositionControls
    
End Sub

Private Sub UpdateAnalysisHistory(udtCurrentIterationStats As udtIterationStatsType, Optional blnUseAbbreviatedFormat As Boolean = False, Optional ByVal blnIncludeExtendedDetails As Boolean = True)
    Dim strDescription As String
    Dim strMessage As String
    Dim strMWTol As String

    With udtCurrentIterationStats

        If .FinalMWTol >= 1 Then
            strMWTol = Format(.FinalMWTol, "0")
        Else
            strMWTol = Format(.FinalMWTol, "0.0000")
        End If
        strMWTol = strMWTol & " " & GetSearchToleranceUnitText(.FinalMWTolType)
        
        If blnUseAbbreviatedFormat Then
            strDescription = "NET Adjustment iteration; " & UMC_NET_ADJ_ITERATION_COUNT & " = " & Trim(.IterationCount) & "; Mass tolerance = ±" & strMWTol & "; NET Tolerance = ±" & Format(.FinalNETTol, "0.000")
            If .UMCTopAbuPct >= 0 Then
                strDescription = strDescription & "; Restrict to x% of UMC's = " & Trim(.UMCTopAbuPct) & "%"
            End If
            
            strDescription = strDescription & "; " & UMC_NET_ADJ_UMCs_IN_TOLERANCE & " = " & Trim(.PointCountSearched) & "; " & UMC_NET_ADJ_UMCs_WITH_DB_HITS & " = " & Trim(.IDMatchCount) & "; NET Formula = " & ConstructNETFormula(.slope, .intercept)
            strDescription = strDescription & "; NET Match Score = " & Format(.NETMatchScore, "0.00")
        Else
            strDescription = "Calculated NET adjustment using UMC's; " & UMC_NET_ADJ_ITERATION_COUNT & " = " & Trim(.IterationCount) & "; Mass tolerance = ±" & strMWTol & "; Final NET Tolerance = ±" & Format(.FinalNETTol, "0.000")
            If .UMCTopAbuPct >= 0 Then
                strDescription = strDescription & "; Restrict to x% of UMC's (sorted by abundance) = " & Trim(.UMCTopAbuPct) & "%"
            End If
            
            strDescription = strDescription & "; " & UMC_NET_ADJ_UMCs_IN_TOLERANCE & " = " & Trim(.PointCountSearched) & "; " & UMC_NET_ADJ_UMCs_WITH_DB_HITS & " = " & Trim(.IDMatchCount) & "; NET Formula = " & ConstructNETFormula(.slope, .intercept) & "; Average Deviation = " & DoubleToStringScientific(.Deviation)
            
            strDescription = strDescription & "; Initial Slope = " & DoubleToStringScientific(.InitialSlope) & "; Initial Intercept = " & Format(.InitialIntercept, "0.0000") & "; Initial Mass Shift = " & Format(.MassShiftPPM, "0.00") & " ppm"
            
            strDescription = strDescription & "; NET Correlation Coeff = " & Format(.CorrelationCoeff, "0.00")
            strDescription = strDescription & "; NET Match Score = " & Format(.NETMatchScore, "0.00")
        End If
    
        AddToAnalysisHistory CallerID, strDescription
        
        If (.UMCSegmentCntWithLowUMCCnt > 0 Or .UMCCntAddedSinceLowSegmentCount > 0) And blnIncludeExtendedDetails Then
            strDescription = "NET Adjustment UMC usage dispersion was low in 1 or more segments; Total segment count = " & Trim(glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions.SegmentCount) & "; Segment count with low UMC counts = " & Trim(.UMCSegmentCntWithLowUMCCnt) & "; UMC's added (total) = " & Trim(.UMCCntAddedSinceLowSegmentCount)
            AddToAnalysisHistory CallerID, strDescription
        End If
    End With
    
    If Not blnUseAbbreviatedFormat And blnIncludeExtendedDetails Then
        strMessage = "Database for NET adjustment: " & CurrMTDBInfo()
        strMessage = strMessage & "; N15 masses used for MT's = " & CStr(glbPreferencesExpanded.NetAdjustmentUsesN15AMTMasses)
        strMessage = strMessage & "; Minimum high normalized score for NET Adjustment MT's = " & CStr(mMTMinimumHighNormalizedScore)
        strMessage = strMessage & "; Minimum high discriminant score for NET Adjustment MT's = " & CStr(mMTMinimumHighDiscriminantScore)
        AddToAnalysisHistory CallerID, strMessage
    End If

End Sub

''Private Sub UpdateDynamicControls()
''
''    On Error Resume Next
''
''    With GelSearchDef(CallerID).MassCalibrationInfo
''        If .MassUnits = glMassToleranceConstants.gltPPM Then
''            txtMassCalibrationOverallAdjustment.Text = Trim(.OverallMassAdjustment)
''        Else
''            txtMassCalibrationOverallAdjustment.Text = "0"
''        End If
''    End With
''
''End Sub

Private Sub UpdateInitialSlopeAndIntercept(dblSlope As Double, dblIntercept As Double)
    
    UMCNetAdjDef.InitialSlope = dblSlope
    UMCNetAdjDef.InitialIntercept = dblIntercept
    
    txtNetAdjInitialSlope.Text = Format(dblSlope, "0.00000000")
    txtNetAdjInitialIntercept.Text = Format(dblIntercept, "0.00000")

    UpdateNetAdjRangeStats UMCNetAdjDef, lblNetAdjInitialNETStats

End Sub

Private Sub UpdateNETHistograms()
    ComputeInternalStandardNETHistogram
    ComputePMTDataNETHistogram
End Sub

Private Sub UpdatePlots(udtIterationStats() As udtIterationStatsType, lngIterationStatsCount As Long, lngCurrentIndex As Long, lngBestIndex As Long)

    Const MASS_ERROR_PLOT_RANGE_PPM As Double = 20
    Const NET_ERROR_PLOT_RANGE As Double = 0.05
    
    Static blnPlotsHaveData As Boolean
    
    Dim lngIndex As Long
    
    Dim dblDataX() As Double                ' 1-based array
    Dim dblDataY() As Double                ' 1-based array
    
    Dim lngMaxDBNetHistogramValue As Long
    Dim lngMaxDataNetHistogramValue As Long
    Dim sngMultiplier As Single
    
On Error GoTo UpdatePlotsErrorHandler
    
    If lngIterationStatsCount <= 0 Then
        If blnPlotsHaveData Then
            InitializePlots
            blnPlotsHaveData = False
        End If
    Else
        blnPlotsHaveData = True
        
        ' Slope for Score plot
        ReDim dblDataX(1 To lngIterationStatsCount)
        ReDim dblDataY(1 To lngIterationStatsCount)
        
        For lngIndex = 0 To lngIterationStatsCount - 1
            dblDataX(lngIndex + 1) = udtIterationStats(lngIndex).slope
            dblDataY(lngIndex + 1) = udtIterationStats(lngIndex).NETMatchScore
        Next lngIndex
    
        With ctlPlotSlopeVsScore
            .EnableDisableDelayUpdating True
            
            .SetCurrentSeries 1
            .SetStyleDataLine 0, oc2dLineNone, 1

            .SetSeriesDataPointCount 1, lngIterationStatsCount
            .SetDataX 1, dblDataX
            .SetDataY 1, dblDataY
            
            .SetStyleDataSymbol vbBlue, oc2dShapeDot, 5
            .SetStyleDataLine vbBlue, oc2dLineNone, 1
            
            If mnuViewAutoZoomPlots Then
                .ZoomOutFull
                .SetXRange 0, .GetXAxisRangeMax
                If .GetYAxisRangeMax > 0 Then
                    .SetYRange 0, .GetYAxisRangeMax
                End If
            End If
            
            .EnableDisableDelayUpdating False
        End With
    
   
        ' NET Range plot (current iteration is series 1, best iteration is series 2, DB data is series 3
        
        ' Determine the multiplier
        ' If the current data's maximum value is >= the DB data's value (or is <=0), then the multiplier is 1
        ' Otherwise the multiplier is 85% of the DB's NET histogram maximum value divided by the data's NET histogram maximum value
        If lngBestIndex >= 0 Then
            lngMaxDataNetHistogramValue = udtIterationStats(lngBestIndex).NETRangeHistogramData.BinCountMaximum
        Else
            lngMaxDataNetHistogramValue = udtIterationStats(lngCurrentIndex).NETRangeHistogramData.BinCountMaximum
        End If
        
        If lngMaxDataNetHistogramValue > 0 Then
            If udtIterationStats(lngCurrentIndex).UseInternalStandards Then
                lngMaxDBNetHistogramValue = mInternalStandardsNETHistogram.BinCountMaximum
            Else
                lngMaxDBNetHistogramValue = mPMTDataNETHistogram.BinCountMaximum
            End If
            
            If lngMaxDataNetHistogramValue >= lngMaxDBNetHistogramValue Then
                sngMultiplier = 1
            Else
                sngMultiplier = 0.85 * (lngMaxDBNetHistogramValue / CSng(lngMaxDataNetHistogramValue))
                If sngMultiplier < 1 Then sngMultiplier = 1
            End If
        Else
            sngMultiplier = 1
        End If
        
        UpdatePlotAddHistogramData ctlPlotNETRange, 1, vbRed, oc2dShapeBox, udtIterationStats(lngCurrentIndex).NETRangeHistogramData, 0, 1, sngMultiplier
        
        If lngBestIndex >= 0 Then
            UpdatePlotAddHistogramData ctlPlotNETRange, 2, vbBlue, oc2dShapeTriangle, udtIterationStats(lngBestIndex).NETRangeHistogramData, 0, 1, sngMultiplier
        Else
            ClearPlotDataSingleSeries ctlPlotNETRange, 2
        End If
                
        If udtIterationStats(lngCurrentIndex).UseInternalStandards Then
            ' Note: 49512 = RGB(0,192,0) = Green
            UpdatePlotAddHistogramData ctlPlotNETRange, 3, 49152, oc2dShapeDot, mInternalStandardsNETHistogram, 0, 1
        Else
            UpdatePlotAddHistogramData ctlPlotNETRange, 3, 49152, oc2dShapeDot, mPMTDataNETHistogram, 0, 1
        End If
        ctlPlotNETRange.SetXAxisTickSpacing 0.25, False
                
                
        ' Mass Error plot (current iteration and best match if lngBestIndex >= 0)
        UpdatePlotAddHistogramData ctlPlotMassErrorHistogram, 1, vbRed, oc2dShapeNone, udtIterationStats(lngCurrentIndex).MassErrorHistogramData, -MASS_ERROR_PLOT_RANGE_PPM, MASS_ERROR_PLOT_RANGE_PPM
        If lngBestIndex >= 0 Then
            UpdatePlotAddHistogramData ctlPlotMassErrorHistogram, 2, vbBlue, oc2dShapeNone, udtIterationStats(lngBestIndex).MassErrorHistogramData, -MASS_ERROR_PLOT_RANGE_PPM, MASS_ERROR_PLOT_RANGE_PPM
        Else
            ClearPlotDataSingleSeries ctlPlotMassErrorHistogram, 2
        End If
        
        ' NET Error plot (current iteration and best match if lngBestIndex >= 0)
        UpdatePlotAddHistogramData ctlPlotNETErrorHistogram, 1, vbRed, oc2dShapeNone, udtIterationStats(lngCurrentIndex).NETErrorHistogramData, -NET_ERROR_PLOT_RANGE, NET_ERROR_PLOT_RANGE
        If lngBestIndex >= 0 Then
            UpdatePlotAddHistogramData ctlPlotNETErrorHistogram, 2, vbBlue, oc2dShapeNone, udtIterationStats(lngBestIndex).NETErrorHistogramData, -NET_ERROR_PLOT_RANGE, NET_ERROR_PLOT_RANGE
        Else
            ClearPlotDataSingleSeries ctlPlotNETErrorHistogram, 2
        End If
    End If

    Exit Sub
    
UpdatePlotsErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchForNETAdjustmentUMC->UpdatePlots"
     
End Sub

Private Sub UpdatePlotAddHistogramData(ByRef ctlPlot As ctlSpectraPlotter, intSeriesNumber As Integer, lngColor As Long, eShape As OlectraChart2D.ShapeConstants, ByRef udtHistogramData As udtBinnedDataType, dblXRangeMin As Double, dblXRangeMax As Double, Optional ByVal sngYAxisMultiplier As Single = 1)
    
    Dim lngIndex As Long
    Dim lngDataCount As Long
    Dim dblDataX() As Double        ' 1-based array
    Dim dblDataY() As Double        ' 1-based array
    
    With udtHistogramData
        lngDataCount = .BinnedCount + 1
        ReDim dblDataX(1 To lngDataCount)
        ReDim dblDataY(1 To lngDataCount)
      
        For lngIndex = 0 To .BinnedCount
            dblDataX(lngIndex + 1) = .StartBin + lngIndex * .BinSize
            dblDataY(lngIndex + 1) = .Binned(lngIndex) * sngYAxisMultiplier
        Next lngIndex
    End With

    With ctlPlot
        .EnableDisableDelayUpdating True
        
        .SetSeriesDataPointCount intSeriesNumber, lngDataCount
        .SetDataX intSeriesNumber, dblDataX
        .SetDataY intSeriesNumber, dblDataY
        
        .SetCurrentSeries intSeriesNumber
        .SetStyleDataSymbol lngColor, eShape, 5
        .SetStyleDataLine lngColor, oc2dLineSolid, 1
        
        If mnuViewAutoZoomPlots.Checked Then
            '.ZoomOutFull
            .SetXRange dblXRangeMin, dblXRangeMax
        End If
        .EnableDisableDelayUpdating False
    End With
        
End Sub

Private Sub UpdateRobustNETStatus(dblSlope As Double, dblIntercept As Double, sngMassShiftPPM As Single)
    lblRobustNETCurrentSettings.Caption = "Initial Slope = " & Format(dblSlope, "0.000000") & vbCrLf & "Initial Intercept = " & Format(dblIntercept, "0.0000") & vbCrLf & "Mass Shift = " & Round(sngMassShiftPPM, 3) & " ppm"
End Sub

Private Sub UpdateRobustNETStatusBestMatch(udtIterationStats As udtIterationStatsType, blnAppendToIterationReport As Boolean)
    ' Update the Robust Net Best Match
    
    Dim strMessage As String
    
    With udtIterationStats
        strMessage = "Best: " & ConstructNETFormula(.slope, .intercept) & "; " & Round(.MassShiftPPM, 1) & " ppm"
        strMessage = strMessage & vbCrLf & "Match Count = " & .UMCMatchCount & "/" & .UMCCountSearched & "; Iterations = " & .IterationCount
        strMessage = strMessage & vbCrLf & "Match Score = " & Round(.NETMatchScore, 3) & "; Dev. = " & DoubleToStringScientific(.Deviation)
    End With
    
    txtRobustNETProgress.Text = strMessage
    If blnAppendToIterationReport Then
        strMessage = "Iteration: " & mNETStatsHistoryCount & vbCrLf & strMessage & vbCrLf & vbCrLf
        AppendToIterationReport strMessage
    End If
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Function ValidatePositiveSlope(blnResetToDefaultIfInvalid As Boolean) As Boolean
    ' Returns true if GelAnalysis(CallerID).GANET_Slope is > 0; otherwise, returns false
    ' Additionaly, updates mNegativeSlopeComputed
    ' If blnResetToDefaultIfInvalid = True, then logs the error (or informs the user) if the slope is negative, and reset to the default slope and intercept
    
    Dim strMessage As String
    
    mNegativeSlopeComputed = False
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
            
            mNegativeSlopeComputed = True
            
            If blnResetToDefaultIfInvalid Then
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                    MsgBox strMessage, vbExclamation Or vbOKOnly, "Invalid NET Slope"
                Else
                    AddToAnalysisHistory CallerID, strMessage
                End If
                
                ' The computed slope was zero or negative; reset the slope and intercept to the default values
                ResetSlopeAndInterceptToDefault
                
                ' Need to assign a non-zero value to GANET_Fit; we'll assign 1.11E-3 with all 1's so it stands out
                GelAnalysis(CallerID).GANET_Fit = 1.11111111111111E-03
            
            End If
        End If
    End If
    
    ValidatePositiveSlope = Not mNegativeSlopeComputed
    
End Function

Private Sub ValidatePMTScoreFilters(blnUsingInternalStandards As Boolean)
    
    If blnUsingInternalStandards Then
        mMTMinimumHighNormalizedScore = 0
        mMTMinimumHighDiscriminantScore = 0
    Else
        mMTMinimumHighNormalizedScore = glbPreferencesExpanded.NetAdjustmentMinHighNormalizedScore
        mMTMinimumHighDiscriminantScore = glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore
        
        If AMTCnt >= 2 Then
            If mMTMinimumHighDiscriminantScore > 0 Then
                ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighDiscriminantScore, also taking into account HighNormalizedScore
                ' It's possible that mMTMinimumHighDiscriminantScore will be lowered to 0 by the call to ValidateMTMinimumDiscriminantAndPepProphet
                ' If that happens, then call ValidateMTMinimimumHighNormalizedScore
                ValidateMTMinimumDiscriminantAndPepProphet AMTData(), 1, AMTCnt, mMTMinimumHighDiscriminantScore, 0, mMTMinimumHighNormalizedScore, 2
            End If
            
            If mMTMinimumHighDiscriminantScore = 0 Then
                ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighNormalizedScore
                ValidateMTMinimimumHighNormalizedScore AMTData(), 1, AMTCnt, mMTMinimumHighNormalizedScore, 2
            End If
        End If
    End If
End Sub

Private Sub ValidateNETFormula()
    If Not InitExprEvaluator(txtNETFormula.Text) Then
       If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox "Error in elution calculation formula.", vbOKOnly, glFGTU
       End If
       txtNETFormula = ConstructNETFormulaWithDefaults(UMCNetAdjDef)
       txtNETFormula.SetFocus
    Else
       UMCNetAdjDef.NETFormula = txtNETFormula.Text
       CheckNETEquationStatus
    End If
End Sub

Private Sub cboPairsUMCsToUseForNETAdjustment_Click()
    glbPreferencesExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection = cboPairsUMCsToUseForNETAdjustment.ListIndex
End Sub

Private Sub chkAcceptLastIteration_Click()
UMCNetAdjDef.IterationAcceptLast = (chkAcceptLastIteration.Value = vbChecked)
End Sub

Private Sub chkCS_Click(Index As Integer)
UMCNetAdjDef.PeakCSSelection(Index) = (chkCS(Index).Value = vbChecked)
EnableDisableControls
End Sub

Private Sub chkDecMW_Click()
UMCNetAdjDef.IterationUseMWDec = (chkDecMW.Value = vbChecked)
End Sub

Private Sub chkDecNET_Click()
UMCNetAdjDef.IterationUseNETdec = (chkDecNET.Value = vbChecked)
End Sub

Private Sub chkEliminateConflictingIDs_Click()
UMCNetAdjDef.UseMultiIDMaxNETDist = (chkEliminateConflictingIDs.Value = vbChecked)
End Sub

Private Sub chkNetAdjAutoIncrementUMCTopAbuPct_Click()
    With glbPreferencesExpanded.AutoAnalysisOptions
        .NETAdjustmentAutoIncrementUMCTopAbuPct = cChkBox(chkNetAdjAutoIncrementUMCTopAbuPct)
        If .NETAdjustmentAutoIncrementUMCTopAbuPct Then
            SetCheckBox chkUMCUseTopAbu, True
            txtUMCAbuTopPct = .NETAdjustmentUMCTopAbuPctInitial
            UMCNetAdjDef.TopAbuPct = .NETAdjustmentUMCTopAbuPctInitial
            
            If optIteStop(ITERATION_STOP_CHANGE).Value <> True Then
                optIteStop(ITERATION_STOP_CHANGE).Value = True
            End If
        End If
    End With
End Sub

' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''Private Sub chkNetAdjUseLockers_Click()
''    UMCNetAdjDef.UseNetAdjLockers = cChkBox(chkNetAdjUseLockers)
''End Sub
''
''Private Sub chkNetAdjUseOldIfFailure_Click()
''    UMCNetAdjDef.UseOldNetAdjIfFailure = cChkBox(chkNetAdjUseOldIfFailure)
''End Sub

Private Sub chkRobustNETEnabled_Click()
    UMCNetAdjDef.UseRobustNETAdjustment = cChkBox(chkRobustNETEnabled)
    EnableDisableControls
End Sub

Private Sub chkUseN15AMTMasses_Click()
    glbPreferencesExpanded.NetAdjustmentUsesN15AMTMasses = cChkBox(chkUseN15AMTMasses)
End Sub

Private Sub chkUseNETForID_Click()
UMCNetAdjDef.UseNET = (chkUseNETForID.Value = vbChecked)
End Sub

''Private Sub cmdMassCalibrationManual_Click()
''    StartManualRefineMassCalibration
''End Sub
''
''Private Sub cmdMassCalibrationRevert_Click()
''    StartMassCalibrationRevert
''End Sub

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
    If mFormInitialized Then CalculateNETAdjustmentStart
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
    PopulateComboBoxes
    TogglePlotFramePosition True
    ShowHideControls False
    InitializePlots
    tbsNETOptions.Tab = 0
    ShowHidePNNLMenus
    InitializeForm
End Sub

Private Sub Form_Resize()
    PositionControls
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
     MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
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
   MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
   Exit Sub
End If
CheckNETEquationStatus
txtNETFormula.SetFocus
End Sub

Private Sub mnuF_Click()
Call PickParameters
End Sub

Private Sub mnuFCalculate_Click()
    If mFormInitialized Then CalculateIterationOneStepOnly
End Sub

Private Sub mnuFExit_Click()
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
    If mFormInitialized Then CalculateNETAdjustmentStart
End Sub

Private Sub mnuFLogCalculations_Click()
    mnuFLogCalculations.Checked = Not mnuFLogCalculations.Checked
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
MsgBox CheckMassTags(), vbOKOnly
End Sub

Private Sub mnuViewAutoZoomPlots_Click()
    mnuViewAutoZoomPlots.Checked = Not mnuViewAutoZoomPlots.Checked
End Sub

Private Sub mnuViewTogglePlotPosition_Click()
    TogglePlotFramePosition False
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

Private Sub optRobustNETSlopeIncrement_Click(Index As Integer)
    UMCNetAdjDef.RobustNETSlopeIncreaseMode = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   UMCNetAdjDef.MWTolType = gltPPM
Else
   UMCNetAdjDef.MWTolType = gltABS
End If
End Sub

Private Sub tmrTimer_Timer()
    If mRequestUpdateRobustNETIterationCount Then
        PredictRobustNETIterationCount
        mRequestUpdateRobustNETIterationCount = False
    End If
End Sub

Private Sub txtIteMWDec_LostFocus()
If IsNumeric(txtIteMWDec.Text) Then
   UMCNetAdjDef.IterationMWDec = CDbl(txtIteMWDec.Text)
Else
   MsgBox "This parameter should be numeric.", vbOKOnly, glFGTU
   txtIteMWDec.SetFocus
End If
End Sub

Private Sub txtIteNETDec_LostFocus()
If IsNumeric(txtIteNETDec.Text) Then
   UMCNetAdjDef.IterationNETDec = CDbl(txtIteNETDec.Text)
Else
   MsgBox "This parameter should be numeric.", vbOKOnly, glFGTU
   txtIteNETDec.SetFocus
End If
End Sub

Private Sub txtIteStopVal_LostFocus()
If IsNumeric(txtIteStopVal.Text) Then
   UMCNetAdjDef.IterationStopValue = CDbl(txtIteStopVal.Text)
   If optIteStop(ITERATION_STOP_CHANGE).Value = True Then
      glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentChangeThresholdStopValue = UMCNetAdjDef.IterationStopValue
   ElseIf optIteStop(ITERATION_STOP_ID_LIMIT).Value = True Then
      glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount = UMCNetAdjDef.IterationStopValue
      txtNetAdjMinIDCount = txtIteStopVal
   ElseIf optIteStop(ITERATION_STOP_NUMBER).Value = True Then
      glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMaxIterationCount = UMCNetAdjDef.IterationStopValue
   End If
Else
   MsgBox "This parameter should be numeric.", vbOKOnly, glFGTU
   txtIteStopVal.SetFocus
End If
End Sub

Private Sub txtMaxUMCScansPct_LostFocus()
If IsNumeric(txtMaxUMCScansPct.Text) Then
   UMCNetAdjDef.MaxScanPct = Abs(CDbl(txtMaxUMCScansPct.Text))
Else
   MsgBox "This parameter should be positive number.", vbOKOnly, glFGTU
   txtMaxUMCScansPct.SetFocus
End If
End Sub

Private Sub txtMinScanRange_LostFocus()
If IsNumeric(txtMinScanRange.Text) Then
   UMCNetAdjDef.MinScanRange = Abs(CLng(txtMinScanRange.Text))
Else
   MsgBox "This parameter should be non-negative integer.", vbOKOnly, glFGTU
   txtMinScanRange.SetFocus
End If
End Sub

Private Sub txtMinUMCCount_LostFocus()
If IsNumeric(txtMinUMCCount.Text) Then
   UMCNetAdjDef.MinUMCCount = Abs(CLng(txtMinUMCCount.Text))
Else
   MsgBox "This parameter should be non-negative integer.", vbOKOnly, glFGTU
   txtMinUMCCount.SetFocus
End If
End Sub

Private Sub txtMultiIDMaxNETDist_LostFocus()
If IsNumeric(txtMultiIDMaxNETDist.Text) Then
   UMCNetAdjDef.MultiIDMaxNETDist = Abs(CDbl(txtMultiIDMaxNETDist.Text))
Else
   MsgBox "This parameter should be non-negative integer.", vbOKOnly, glFGTU
   txtMultiIDMaxNETDist.SetFocus
End If
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   UMCNetAdjDef.MWTol = Abs(CDbl(txtMWTol.Text))
Else
   MsgBox "This parameter should be a non-negative number.", vbOKOnly, glFGTU
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNetAdjInitialIntercept_LostFocus()
    ValidateTextboxValueDbl txtNetAdjInitialIntercept, -1000, 1000, 0
    UMCNetAdjDef.InitialIntercept = CDblSafe(txtNetAdjInitialIntercept)
    EnableDisableControls
End Sub

Private Sub txtNetAdjInitialSlope_LostFocus()
    ValidateTextboxValueDbl txtNetAdjInitialSlope, 0.0000001, 100, 0.0002
    UMCNetAdjDef.InitialSlope = CDblSafe(txtNetAdjInitialSlope)
    EnableDisableControls
End Sub

Private Sub txtNetAdjMinHighDiscriminantScore_LostFocus()
If IsNumeric(txtNetAdjMinHighDiscriminantScore.Text) Then
    glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore = Abs(CSng(txtNetAdjMinHighDiscriminantScore.Text))
    If glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore > 1 Then
        glbPreferencesExpanded.NetAdjustmentMinHighDiscriminantScore = 0.999
    End If
Else
    MsgBox "This parameter should be non-negative integer.", vbOKOnly, glFGTU
    txtNetAdjMinHighDiscriminantScore.SetFocus
End If
End Sub

Private Sub txtNetAdjMinHighNormalizedScore_LostFocus()
If IsNumeric(txtNetAdjMinHighNormalizedScore.Text) Then
    glbPreferencesExpanded.NetAdjustmentMinHighNormalizedScore = Abs(CSng(txtNetAdjMinHighNormalizedScore.Text))
Else
    MsgBox "This parameter should be non-negative integer.", vbOKOnly, glFGTU
    txtNetAdjMinHighNormalizedScore.SetFocus
End If
End Sub

Private Sub txtNetAdjMinIDCount_LostFocus()
    ValidateTextboxValueLng txtNetAdjMinIDCount, 2, 100000, 75
    If optIteStop(ITERATION_STOP_ID_LIMIT).Value = True Then
        txtIteStopVal = txtNetAdjMinIDCount
        Call txtIteStopVal_LostFocus
    End If
    glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount = CLngSafe(txtNetAdjMinIDCount)
End Sub

' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''Private Sub txtNetAdjMinLockerMatchCount_LostFocus()
''    ValidateTextboxValueLng txtNetAdjMinLockerMatchCount, 2, 1000, 3
''    UMCNetAdjDef.NetAdjLockerMinimumMatchCount = CLngSafe(txtNetAdjMinLockerMatchCount)
''End Sub

''Private Sub txtNETAdjMinimumNETMatchScore_LostFocus()
''    ValidateTextboxValueLng txtNETAdjMinimumNETMatchScore, 1, 10000, 50
''    glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinimumNETMatchScore = CLngSafe(txtNETAdjMinimumNETMatchScore)
''End Sub

Private Sub txtNetAdjWarningTol_LostFocus(Index As Integer)

    With glbPreferencesExpanded.AutoAnalysisOptions
        Select Case Index
        Case naswSlopeMinimum
            ValidateTextboxValueDbl txtNetAdjWarningTol(Index), 0.0000001, 100, 0.000005
            .NETSlopeExpectedMinimum = CDblSafe(txtNetAdjWarningTol(Index))
        Case naswSlopeMaximum
            ValidateTextboxValueDbl txtNetAdjWarningTol(Index), 0.0000001, 100, 0.1
            .NETSlopeExpectedMaximum = CDblSafe(txtNetAdjWarningTol(Index))
        Case naswInterceptMinimum
            ValidateTextboxValueDbl txtNetAdjWarningTol(Index), -1000, 1000, -1
            .NETInterceptExpectedMinimum = CDblSafe(txtNetAdjWarningTol(Index))
        Case naswInterceptMaximum
            ValidateTextboxValueDbl txtNetAdjWarningTol(Index), -1000, 1000, 1
            .NETInterceptExpectedMaximum = CDblSafe(txtNetAdjWarningTol(Index))
        End Select
    End With
    
End Sub

Private Sub txtNETFormula_LostFocus()
    ValidateNETFormula
End Sub

Private Sub txtNETTol_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtNETTol, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtNETTol_LostFocus()
Dim tmp As String
tmp = Trim(txtNETTol.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 And tmp <= 1 Then
      UMCNetAdjDef.NETTolIterative = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "NET Tolerance should be a number between 0 and 1.", vbOKOnly, glFGTU
txtNETTol.SetFocus
End Sub

Private Sub txtNETTol_Validate(Cancel As Boolean)
    TextBoxLimitNumberLength txtNETTol, 12
End Sub

''Private Sub txtRefineMassCalibrationMaximumShift_LostFocus()
''    ValidateTextboxValueDbl txtRefineMassCalibrationMaximumShift, 0, 1E+300, 15
''    glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationMaximumShift = CDblSafe(txtRefineMassCalibrationMaximumShift)
''End Sub
''
''Private Sub txtRobustNETAnnealMaxSwapsPerStep_LostFocus()
''    ValidateTextboxValueLng txtRobustNETAnnealMaxSwapsPerStep, 1, 100000, 50
''    UMCNetAdjDef.RobustNETAnnealMaxSwapsPerStep = CLngSafe(txtRobustNETAnnealMaxSwapsPerStep)
''End Sub
''
''Private Sub txtRobustNETAnnealSteps_LostFocus()
''    ValidateTextboxValueLng txtRobustNETAnnealSteps, 1, 100000, 20
''    UMCNetAdjDef.RobustNETAnnealSteps = CLngSafe(txtRobustNETAnnealSteps)
''End Sub
''
''Private Sub txtRobustNETAnnealTemperatureReductionFactor_LostFocus()
''    ValidateTextboxValueDbl txtRobustNETAnnealTemperatureReductionFactor, 0.0001, 0.9999, 0.9
''    UMCNetAdjDef.RobustNETAnnealTemperatureReductionFactor = CSngSafe(txtRobustNETAnnealTemperatureReductionFactor)
''End Sub
''
''Private Sub txtRobustNETAnnealTrialsPerStep_LostFocus()
''    ValidateTextboxValueLng txtRobustNETAnnealTrialsPerStep, 1, 100000, 250
''    UMCNetAdjDef.RobustNETAnnealTrialsPerStep = CLngSafe(txtRobustNETAnnealTrialsPerStep)
''End Sub

Private Sub txtRobustNETInterceptEnd_LostFocus()
    ValidateTextboxValueDbl txtRobustNETInterceptEnd, -1000, 1000, 0.2
    UMCNetAdjDef.RobustNETInterceptEnd = CSngSafe(txtRobustNETInterceptEnd)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETInterceptIncrement_LostFocus()
    ValidateTextboxValueDbl txtRobustNETInterceptIncrement, -1000, 1000, 0.2
    UMCNetAdjDef.RobustNETInterceptIncrement = CSngSafe(txtRobustNETInterceptIncrement)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETInterceptStart_LostFocus()
    ValidateTextboxValueDbl txtRobustNETInterceptStart, -1000, 1000, -0.4
    UMCNetAdjDef.RobustNETInterceptStart = CSngSafe(txtRobustNETInterceptStart)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETMassShiftPPMEnd_LostFocus()
    ValidateTextboxValueDbl txtRobustNETMassShiftPPMEnd, -10000, 10000, CInt(UMCNetAdjDef.MWTol * 1.5)
    UMCNetAdjDef.RobustNETMassShiftPPMEnd = CSngSafe(txtRobustNETMassShiftPPMEnd)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETMassShiftPPMIncrement_LostFocus()
    ValidateTextboxValueDbl txtRobustNETMassShiftPPMIncrement, -10000, 10000, CInt(UMCNetAdjDef.MWTol * 1.5)
    UMCNetAdjDef.RobustNETMassShiftPPMIncrement = CSngSafe(txtRobustNETMassShiftPPMIncrement)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETMassShiftPPMStart_LostFocus()
    ValidateTextboxValueDbl txtRobustNETMassShiftPPMStart, -10000, 10000, -CInt(UMCNetAdjDef.MWTol * 1.5)
    UMCNetAdjDef.RobustNETMassShiftPPMStart = CSngSafe(txtRobustNETMassShiftPPMStart)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETSlopeEnd_LostFocus()
    ValidateTextboxValueDbl txtRobustNETSlopeEnd, 0.0000001, 100, 0.002
    UMCNetAdjDef.RobustNETSlopeEnd = CSngSafe(txtRobustNETSlopeEnd)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETSlopeIncrement_LostFocus()
    If optRobustNETSlopeIncrement(UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear).Value = True Then
        ValidateTextboxValueDbl txtRobustNETSlopeIncrement, 0.0000001, 100, 0.00005
    Else
        ValidateTextboxValueDbl txtRobustNETSlopeIncrement, 1, 1000, 75
    End If
    
    UMCNetAdjDef.RobustNETSlopeIncrement = CSngSafe(txtRobustNETSlopeIncrement)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtRobustNETSlopeStart_LostFocus()
    ValidateTextboxValueDbl txtRobustNETSlopeStart, 0.0000001, 100, 0.00002
    UMCNetAdjDef.RobustNETSlopeStart = CSngSafe(txtRobustNETSlopeStart)
    mRequestUpdateRobustNETIterationCount = True
End Sub

Private Sub txtToleranceRefinementMinimumPeakHeight_LostFocus()
    ValidateTextboxValueLng txtToleranceRefinementMinimumPeakHeight, 0, 1000000000#, 25
    glbPreferencesExpanded.RefineMSDataOptions.MinimumPeakHeight = CLngSafe(txtToleranceRefinementMinimumPeakHeight)
End Sub

Private Sub txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks_LostFocus()
    ValidateTextboxValueDbl txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks, 0, 100000, 2.5
    glbPreferencesExpanded.RefineMSDataOptions.MinimumSignalToNoiseRatioForLowAbundancePeaks = CSngSafe(txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks)
End Sub

Private Sub txtToleranceRefinementPercentageOfMaxForWidth_Lostfocus()
    ValidateTextboxValueLng txtToleranceRefinementPercentageOfMaxForWidth, 0, 100, 60
    glbPreferencesExpanded.RefineMSDataOptions.PercentageOfMaxForFindingWidth = CLngSafe(txtToleranceRefinementPercentageOfMaxForWidth)
End Sub


Private Sub txtUMCAbuTopPct_LostFocus()
If IsNumeric(txtUMCAbuTopPct.Text) Then
   UMCNetAdjDef.TopAbuPct = Abs(CDbl(txtUMCAbuTopPct.Text))
Else
   MsgBox "This parameter should be positive number.", vbOKOnly, glFGTU
   txtUMCAbuTopPct.SetFocus
End If
End Sub


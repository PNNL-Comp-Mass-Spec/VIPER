VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{C02A7541-5364-11D2-9373-00A02411EBE6}#1.6#0"; "cw3dgrph.ocx"
Begin VB.Form frmMSAlign 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LCMSWarp"
   ClientHeight    =   10410
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14340
   LinkTopic       =   "MS Align"
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboStepsToPerform 
      Height          =   315
      Left            =   11280
      Style           =   2  'Dropdown List
      TabIndex        =   106
      Top             =   9240
      Width           =   2895
   End
   Begin VB.Frame fraMassRefinementStats 
      BackColor       =   &H80000005&
      Caption         =   "Mass Refinment Stats"
      Height          =   1815
      Left            =   9000
      TabIndex        =   93
      Top             =   7800
      Width           =   2175
      Begin VB.TextBox txtMassCalibrationOverallShiftCount 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtMassCalibrationOverallAdjustment 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdMassCalibrationRevert 
         Caption         =   "Revert to Original Masses"
         Height          =   375
         Left            =   50
         TabIndex        =   98
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mass shift count:"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label lblMassCalibrationRefinementUnits 
         Caption         =   "ppm"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   117
         Top             =   1140
         Width           =   600
      End
      Begin VB.Label lblMassCalibrationOverallAdjustment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Avg mass shift (ppm):"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab tbsOptions 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   7800
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3201
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   5
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   -2147483643
      TabCaption(0)   =   "NET Options"
      TabPicture(0)   =   "frmMSAlign.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraNETWarpOptions"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mass Options"
      TabPicture(1)   =   "frmMSAlign.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraMassWarpOptions"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tolerances"
      TabPicture(2)   =   "frmMSAlign.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraNETTolerances"
      Tab(2).Control(1)=   "fraBinningOptions"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Calib Type"
      TabPicture(3)   =   "frmMSAlign.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraMassCalibType"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Adv1"
      TabPicture(4)   =   "frmMSAlign.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label21"
      Tab(4).Control(1)=   "chkWarpMassUseLSQ"
      Tab(4).Control(2)=   "Frame1"
      Tab(4).Control(3)=   "txtWarpMassZScoreTolerance"
      Tab(4).Control(4)=   "fraMTRangeFilters"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Adv2"
      TabPicture(5)   =   "frmMSAlign.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "fraSplitWarpOptions"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Plots"
      TabPicture(6)   =   "frmMSAlign.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraResidualPlotOptions"
      Tab(6).ControlCount=   1
      Begin VB.Frame fraSplitWarpOptions 
         Caption         =   "Split warp options"
         Height          =   1095
         Left            =   240
         TabIndex        =   63
         Top             =   480
         Width           =   4695
         Begin VB.CommandButton cmdSplitWarpResume 
            Caption         =   "Resume"
            Height          =   255
            Left            =   2760
            TabIndex        =   120
            Top             =   760
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkSplitWarpPauseBetweenIterations 
            Caption         =   "Pause between iterations"
            Height          =   200
            Left            =   120
            TabIndex        =   119
            Top             =   760
            Width           =   2295
         End
         Begin VB.TextBox txtSplitWarpMZBoundary 
            Height          =   285
            Left            =   3240
            TabIndex        =   67
            ToolTipText     =   "m/z value at which to split the data"
            Top             =   240
            Width           =   1185
         End
         Begin VB.OptionButton optSplitWarpOnMZ 
            Caption         =   "Split on m/z"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optSplitWarpDisabled 
            Caption         =   "Disabled"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label lblSplitWarpMZBoundary 
            Caption         =   "m/z boundary"
            Height          =   315
            Left            =   1920
            TabIndex        =   66
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.Frame fraResidualPlotOptions 
         Caption         =   "Residual Plot Options"
         Height          =   1350
         Left            =   -74880
         TabIndex        =   68
         Top             =   360
         Width           =   6015
         Begin VB.CheckBox chkAutoZoomOut 
            Caption         =   "Auto-zoom out after alignment"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   960
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.ComboBox cboResidualPlotPointSize 
            Height          =   315
            ItemData        =   "frmMSAlign.frx":00C4
            Left            =   1320
            List            =   "frmMSAlign.frx":00C6
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   240
            Width           =   645
         End
         Begin VB.ComboBox cboResidualPlotTransformationFnLineSize 
            Height          =   315
            ItemData        =   "frmMSAlign.frx":00C8
            Left            =   1320
            List            =   "frmMSAlign.frx":00CA
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   600
            Width           =   645
         End
         Begin VB.TextBox txtResidualPlotMinX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   75
            Text            =   "0"
            Top             =   210
            Width           =   800
         End
         Begin VB.TextBox txtResidualPlotMinY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4920
            TabIndex        =   79
            Text            =   "0"
            Top             =   210
            Width           =   800
         End
         Begin VB.TextBox txtResidualPlotMaxX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   77
            Text            =   "0"
            Top             =   550
            Width           =   800
         End
         Begin VB.TextBox txtResidualPlotMaxY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4920
            TabIndex        =   81
            Text            =   "0"
            Top             =   550
            Width           =   800
         End
         Begin VB.CommandButton cmdResidualPlotSetRange 
            Caption         =   "Set Range for Current Plot"
            Height          =   375
            Left            =   2640
            TabIndex        =   82
            Top             =   900
            Width           =   2055
         End
         Begin VB.CommandButton cmdZoomOutResidualsPlot 
            Cancel          =   -1  'True
            Caption         =   "Zoom Out"
            Height          =   375
            Left            =   4800
            TabIndex        =   83
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Point Size"
            Height          =   285
            Left            =   120
            TabIndex        =   69
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label23 
            Caption         =   "Line Size"
            Height          =   285
            Left            =   120
            TabIndex        =   71
            Top             =   630
            Width           =   975
         End
         Begin VB.Label lblResidualPlotMinX 
            Caption         =   "X Min"
            Height          =   255
            Left            =   2550
            TabIndex        =   74
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblResidualPlotMinY 
            Caption         =   "Y Min"
            Height          =   255
            Left            =   4350
            TabIndex        =   78
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblResidualPlotMaxX 
            Caption         =   "X Max"
            Height          =   255
            Left            =   2550
            TabIndex        =   76
            Top             =   585
            Width           =   600
         End
         Begin VB.Label lblResidualPlotMaxY 
            Caption         =   "Y Max"
            Height          =   255
            Left            =   4350
            TabIndex        =   80
            Top             =   580
            Width           =   600
         End
      End
      Begin VB.Frame fraMTRangeFilters 
         Caption         =   "MT Tag Range Filters (leave blank to ignore)"
         Height          =   975
         Left            =   -72600
         TabIndex        =   54
         Top             =   720
         Width           =   3855
         Begin VB.TextBox txtAMTMassMax 
            Height          =   285
            Left            =   3000
            TabIndex        =   62
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtAMTMassMin 
            Height          =   285
            Left            =   3000
            TabIndex        =   60
            Top             =   220
            Width           =   735
         End
         Begin VB.TextBox txtAMTNetMax 
            Height          =   285
            Left            =   1080
            TabIndex        =   58
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtAMTNetMin 
            Height          =   285
            Left            =   1080
            TabIndex        =   56
            Top             =   220
            Width           =   735
         End
         Begin VB.Label Label25 
            Caption         =   "Mass Max:"
            Height          =   255
            Left            =   2040
            TabIndex        =   61
            Top             =   620
            Width           =   1005
         End
         Begin VB.Label Label24 
            Caption         =   "Mass Min:"
            Height          =   285
            Left            =   2040
            TabIndex        =   59
            Top             =   250
            Width           =   1005
         End
         Begin VB.Label lblAMTNetMax 
            Caption         =   "NET Max:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   620
            Width           =   1000
         End
         Begin VB.Label lblAMTNetMin 
            Caption         =   "NET Min:"
            Height          =   285
            Left            =   120
            TabIndex        =   55
            Top             =   250
            Width           =   1000
         End
      End
      Begin VB.TextBox txtWarpMassZScoreTolerance 
         Height          =   285
         Left            =   -69840
         TabIndex        =   48
         Text            =   "3"
         Top             =   420
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "LSQ Options"
         Height          =   975
         Left            =   -74880
         TabIndex        =   49
         Top             =   720
         Width           =   2175
         Begin VB.TextBox txtWarpMassLSQOutlierZScore 
            Height          =   285
            Left            =   1320
            TabIndex        =   51
            Text            =   "3"
            Top             =   220
            Width           =   735
         End
         Begin VB.TextBox txtWarpMassLSQNumKnots 
            Height          =   285
            Left            =   1320
            TabIndex        =   53
            Text            =   "12"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Outlier z-score:"
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   250
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "# of knots:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   620
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkWarpMassUseLSQ 
         Caption         =   "Use LSQ (least squares fit)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   46
         Top             =   420
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Frame fraMassCalibType 
         Caption         =   "Mass Calibration Type"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   6120
         Begin VB.OptionButton optMassRecalMZRegression 
            Caption         =   "Recalibrate m/z coefficients (Algorithm by Aleksey Tolmachev)"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   260
            Value           =   -1  'True
            Width           =   5800
         End
         Begin VB.OptionButton optMassRecalScanRegression 
            Caption         =   "Recalibrate mass vs. elution time (Algorithm by Vlad Petyuk)"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   560
            Width           =   5800
         End
         Begin VB.OptionButton optMassRecalHybrid 
            Caption         =   "Hybrid Recalibration (Algorithm by Navdeep Jaitly); recalibrate m/z coefficients followed by time drift"
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   860
            Width           =   5800
         End
      End
      Begin VB.Frame fraBinningOptions 
         Caption         =   "Histogram Binning Options"
         Height          =   975
         Left            =   -71880
         TabIndex        =   36
         Top             =   360
         Width           =   3135
         Begin VB.TextBox txtMassBinSizePPM 
            Height          =   285
            Left            =   1560
            TabIndex        =   38
            Text            =   "0.2"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtGANETBinSize 
            Height          =   285
            Left            =   1560
            TabIndex        =   41
            Text            =   "0.001"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblMassBinSizePPM 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Bin Size"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblMassBinSizePPMUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "ppm"
            Height          =   255
            Left            =   2400
            TabIndex        =   39
            Top             =   270
            Width           =   495
         End
         Begin VB.Label lblGANETBinSize 
            BackStyle       =   0  'Transparent
            Caption         =   "NET Bin Size"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   630
            Width           =   1335
         End
      End
      Begin VB.Frame fraNETTolerances 
         Caption         =   "NET Tolerances"
         Height          =   975
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   2895
         Begin VB.TextBox txtMassTolerance 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1920
            TabIndex        =   33
            Text            =   "20"
            Top             =   240
            Width           =   800
         End
         Begin VB.TextBox txtNetTolerance 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1920
            TabIndex        =   35
            Text            =   "0.02"
            Top             =   600
            Width           =   800
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Tolerance (ppm):"
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Tolerance:"
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   630
            Width           =   1695
         End
      End
      Begin VB.Frame fraMassWarpOptions 
         Caption         =   "Mass Warp Options"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtWarpMassMaxJump 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5640
            TabIndex        =   29
            Text            =   "50"
            ToolTipText     =   "Note: This value must be less than or equal to the # of mass delta bins; default is half the number of bins"
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtWarpMassNumMassDeltaBins 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   27
            Text            =   "100"
            Top             =   600
            Width           =   600
         End
         Begin VB.TextBox txtWarpMassNumXSlices 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   25
            Text            =   "20"
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox txtSplineOrder 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   23
            Text            =   "2"
            Top             =   600
            Width           =   615
         End
         Begin VB.CheckBox chkWarpMassAutoRemovePreviousMassCalibrations 
            Caption         =   "Auto-remove previous mass calibration values"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Value           =   1  'Checked
            Width           =   4095
         End
         Begin VB.TextBox txtWarpMassWindowPPM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Text            =   "50"
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Max jump:"
            Height          =   255
            Left            =   4680
            TabIndex        =   28
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "# of mass delta bins:"
            Height          =   405
            Left            =   2400
            TabIndex        =   26
            Top             =   630
            Width           =   1485
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "# of x-axis slices:"
            Height          =   285
            Left            =   2400
            TabIndex        =   24
            Top             =   270
            Width           =   1485
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Spline Order:"
            Height          =   255
            Left            =   30
            TabIndex        =   22
            Top             =   630
            Width           =   1605
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Window (ppm):"
            Height          =   285
            Left            =   0
            TabIndex        =   20
            Top             =   270
            Width           =   1605
         End
      End
      Begin VB.Frame fraNETWarpOptions 
         Caption         =   "NET Warp Options"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtNumSections 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Text            =   "100"
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox txtContractionFactor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   14
            Text            =   "2"
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox txtMaxDistortion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   12
            Text            =   "3"
            Top             =   600
            Width           =   600
         End
         Begin VB.TextBox txtMinMSMSObservations 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   16
            Text            =   "5"
            Top             =   600
            Width           =   600
         End
         Begin VB.TextBox txtMassTagMatchPromiscuity 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   18
            Text            =   "2"
            Top             =   960
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "# of sections:"
            Height          =   240
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   1200
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Contraction Factor:"
            Height          =   240
            Left            =   3120
            TabIndex        =   13
            Top             =   255
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Max Distortion:"
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Min MS/MS observations:"
            Height          =   240
            Left            =   3000
            TabIndex        =   15
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "MassTag Match Promiscuity:"
            Height          =   240
            Left            =   2880
            TabIndex        =   17
            Top             =   960
            Width           =   2280
         End
      End
      Begin VB.Label Label21 
         Caption         =   "z-score tolerance:"
         Height          =   210
         Left            =   -71520
         TabIndex        =   47
         Top             =   450
         Width           =   1695
      End
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H80000005&
      Height          =   735
      Left            =   120
      TabIndex        =   110
      Top             =   9600
      Width           =   11055
      Begin VB.Label lblUMCMassMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LC-MS Feature Mass Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         ToolTipText     =   "Status of the MT tag database"
         Top             =   390
         Width           =   10755
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   135
         Width           =   10815
      End
   End
   Begin VB.CommandButton cmdSetDefaults 
      Caption         =   "Set to Defaults"
      Height          =   375
      Left            =   11280
      TabIndex        =   107
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Frame fraNETAlignmentStats 
      BackColor       =   &H80000005&
      Caption         =   "NET Alignment Stats"
      Height          =   1815
      Left            =   6600
      TabIndex        =   84
      Top             =   7800
      Width           =   2295
      Begin VB.TextBox txtFit 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   1360
         Width           =   1095
      End
      Begin VB.TextBox txtRSquared 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtSlope 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   86
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtIntercept 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   88
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblFit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mean residual:"
         Height          =   375
         Left            =   120
         TabIndex        =   91
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblRSquared 
         BackColor       =   &H00FFFFFF&
         Caption         =   "R-squared:"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Slope:"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Intercept:"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   640
         Width           =   855
      End
   End
   Begin VB.Timer tmrAlignment 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   12840
      TabIndex        =   109
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H80000005&
      Caption         =   "Info"
      Height          =   1335
      Left            =   11280
      TabIndex        =   99
      Top             =   7800
      Width           =   2895
      Begin VB.TextBox txtNumMatched 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   960
         Width           =   1300
      End
      Begin VB.TextBox txtFeatureCountLoaded 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   600
         Width           =   1300
      End
      Begin VB.TextBox txtPMTCountLoaded 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   240
         Width           =   1300
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Match Count:"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   960
         Width           =   1450
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LC-MS Features:"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   600
         Width           =   1450
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MT Tag Count:"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   1450
      End
   End
   Begin VB.CommandButton cmdWarpAlign 
      Caption         =   "Align (Warped)"
      Height          =   375
      Left            =   12840
      TabIndex        =   108
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Frame fraErrors 
      BackColor       =   &H00FFFFFF&
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin CWUIControlsLib.CWGraph graphMassErrors 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4215
         _Version        =   393218
         _ExtentX        =   7435
         _ExtentY        =   5953
         _StockProps     =   71
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   393218
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   62
         C[0]_1          =   16777215
         C[1]_1          =   16777215
         Event_1         =   2
         ClassName_2     =   "CCWGFPlotEvent"
         Owner_2         =   1
         Plots_1         =   3
         ClassName_3     =   "CCWDataPlots"
         Array_3         =   1
         Editor_3        =   4
         ClassName_4     =   "CCWGFPlotArrayEditor"
         Owner_4         =   1
         Array[0]_3      =   5
         ClassName_5     =   "CCWDataPlot"
         opts_5          =   4194367
         Name_5          =   "Plot-1"
         C[0]_5          =   0
         C[1]_5          =   255
         C[2]_5          =   16711680
         C[3]_5          =   16776960
         Event_5         =   2
         X_5             =   6
         ClassName_6     =   "CCWAxis"
         opts_6          =   1599
         Name_6          =   "XAxis"
         Orientation_6   =   2944
         format_6        =   7
         ClassName_7     =   "CCWFormat"
         Scale_6         =   8
         ClassName_8     =   "CCWScale"
         opts_8          =   90112
         rMin_8          =   38
         rMax_8          =   270
         dMax_8          =   10
         discInterval_8  =   1
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   0
         tickopts_6      =   2711
         major_6         =   2
         minor_6         =   1
         Caption_6       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   -2147483640
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         szText_12       =   "Mass Error (ppm)"
         font_12         =   0
         Animator_11     =   0
         Blinker_11      =   0
         Y_5             =   13
         ClassName_13    =   "CCWAxis"
         opts_13         =   1599
         Name_13         =   "YAxis-1"
         Orientation_13  =   2067
         format_13       =   14
         ClassName_14    =   "CCWFormat"
         Scale_13        =   15
         ClassName_15    =   "CCWScale"
         opts_15         =   122880
         rMin_15         =   24
         rMax_15         =   185
         dMax_15         =   10
         discInterval_15 =   1
         Radial_13       =   0
         Enum_13         =   16
         ClassName_16    =   "CCWEnum"
         Editor_16       =   17
         ClassName_17    =   "CCWEnumArrayEditor"
         Owner_17        =   13
         Font_13         =   0
         tickopts_13     =   2711
         major_13        =   2
         minor_13        =   1
         Caption_13      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   -2147483640
         Image_18        =   19
         ClassName_19    =   "CCWTextImage"
         szText_19       =   "Count"
         font_19         =   0
         Animator_18     =   0
         Blinker_18      =   0
         LineStyle_5     =   1
         LineWidth_5     =   1
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Axes_1          =   20
         ClassName_20    =   "CCWAxes"
         Array_20        =   2
         Editor_20       =   21
         ClassName_21    =   "CCWGFAxisArrayEditor"
         Owner_21        =   1
         Array[0]_20     =   6
         Array[1]_20     =   13
         DefaultPlot_1   =   22
         ClassName_22    =   "CCWDataPlot"
         opts_22         =   4194367
         Name_22         =   "[Template]"
         C[0]_22         =   65280
         C[1]_22         =   255
         C[2]_22         =   16711680
         C[3]_22         =   16776960
         Event_22        =   2
         X_22            =   6
         Y_22            =   13
         LineStyle_22    =   1
         LineWidth_22    =   1
         BasePlot_22     =   0
         DefaultXInc_22  =   1
         DefaultPlotPerRow_22=   -1  'True
         Cursors_1       =   23
         ClassName_23    =   "CCWCursors"
         Editor_23       =   24
         ClassName_24    =   "CCWGFCursorArrayEditor"
         Owner_24        =   1
         TrackMode_1     =   11
         GraphBackground_1=   0
         GraphFrame_1    =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         C[0]_25         =   16777215
         C[1]_25         =   16777215
         Image_25        =   26
         ClassName_26    =   "CCWPictImage"
         opts_26         =   1280
         Rows_26         =   1
         Cols_26         =   1
         F_26            =   16777215
         B_26            =   16777215
         ColorReplaceWith_26=   8421504
         ColorReplace_26 =   8421504
         Tolerance_26    =   2
         Animator_25     =   0
         Blinker_25      =   0
         PlotFrame_1     =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   62
         C[0]_27         =   16777215
         C[1]_27         =   16777215
         Image_27        =   28
         ClassName_28    =   "CCWPictImage"
         opts_28         =   1280
         Rows_28         =   1
         Cols_28         =   1
         Pict_28         =   1
         F_28            =   16777215
         B_28            =   16777215
         ColorReplaceWith_28=   8421504
         ColorReplace_28 =   8421504
         Tolerance_28    =   2
         Animator_27     =   0
         Blinker_27      =   0
         Caption_1       =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   62
         C[0]_29         =   -2147483640
         Image_29        =   30
         ClassName_30    =   "CCWTextImage"
         szText_30       =   "Mass Error Histogram"
         font_30         =   0
         Animator_29     =   0
         Blinker_29      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   31
         ClassName_31    =   "CCWBindingHolderArray"
         Editor_31       =   32
         ClassName_32    =   "CCWBindingHolderArrayEditor"
         Owner_32        =   1
         Annotations_1   =   33
         ClassName_33    =   "CCWAnnotations"
         Editor_33       =   34
         ClassName_34    =   "CCWAnnotationArrayEditor"
         Owner_34        =   1
         AnnotationTemplate_1=   35
         ClassName_35    =   "CCWAnnotation"
         opts_35         =   63
         Name_35         =   "[Template]"
         Plot_35         =   22
         Text_35         =   "[Template]"
         TextXPoint_35   =   6.7
         TextYPoint_35   =   6.7
         TextColor_35    =   16777215
         TextFont_35     =   36
         ClassName_36    =   "CCWFont"
         bFont_36        =   -1  'True
         BeginProperty Font_36 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShapeXPoints_35 =   37
         ClassName_37    =   "CDataBuffer"
         Type_37         =   5
         m_cDims;_37     =   1
         m_cElts_37      =   1
         Element[0]_37   =   3.3
         ShapeYPoints_35 =   38
         ClassName_38    =   "CDataBuffer"
         Type_38         =   5
         m_cDims;_38     =   1
         m_cElts_38      =   1
         Element[0]_38   =   3.3
         ShapeFillColor_35=   16777215
         ShapeLineColor_35=   16777215
         ShapeLineWidth_35=   1
         ShapeLineStyle_35=   1
         ShapePointStyle_35=   10
         ShapeImage_35   =   39
         ClassName_39    =   "CCWDrawObj"
         opts_39         =   62
         Image_39        =   40
         ClassName_40    =   "CCWPictImage"
         opts_40         =   1280
         Rows_40         =   1
         Cols_40         =   1
         Pict_40         =   7
         F_40            =   -2147483633
         B_40            =   -2147483633
         ColorReplaceWith_40=   8421504
         ColorReplace_40 =   8421504
         Tolerance_40    =   2
         Animator_39     =   0
         Blinker_39      =   0
         ArrowVisible_35 =   -1  'True
         ArrowColor_35   =   16777215
         ArrowWidth_35   =   1
         ArrowLineStyle_35=   1
         ArrowHeadStyle_35=   1
      End
      Begin CWUIControlsLib.CWGraph graphNetErrors 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Width           =   4215
         _Version        =   393218
         _ExtentX        =   7435
         _ExtentY        =   5953
         _StockProps     =   71
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   393218
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   62
         C[0]_1          =   16777215
         C[1]_1          =   16777215
         Event_1         =   2
         ClassName_2     =   "CCWGFPlotEvent"
         Owner_2         =   1
         Plots_1         =   3
         ClassName_3     =   "CCWDataPlots"
         Array_3         =   1
         Editor_3        =   4
         ClassName_4     =   "CCWGFPlotArrayEditor"
         Owner_4         =   1
         Array[0]_3      =   5
         ClassName_5     =   "CCWDataPlot"
         opts_5          =   4194367
         Name_5          =   "Plot-1"
         C[0]_5          =   0
         C[1]_5          =   255
         C[2]_5          =   16711680
         C[3]_5          =   16776960
         Event_5         =   2
         X_5             =   6
         ClassName_6     =   "CCWAxis"
         opts_6          =   1599
         Name_6          =   "XAxis"
         Orientation_6   =   2944
         format_6        =   7
         ClassName_7     =   "CCWFormat"
         Scale_6         =   8
         ClassName_8     =   "CCWScale"
         opts_8          =   90112
         rMin_8          =   38
         rMax_8          =   270
         dMax_8          =   10
         discInterval_8  =   1
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   0
         tickopts_6      =   2711
         major_6         =   2
         minor_6         =   1
         Caption_6       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   -2147483640
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         szText_12       =   "NET Error"
         font_12         =   0
         Animator_11     =   0
         Blinker_11      =   0
         Y_5             =   13
         ClassName_13    =   "CCWAxis"
         opts_13         =   1599
         Name_13         =   "YAxis-1"
         Orientation_13  =   2067
         format_13       =   14
         ClassName_14    =   "CCWFormat"
         Scale_13        =   15
         ClassName_15    =   "CCWScale"
         opts_15         =   122880
         rMin_15         =   24
         rMax_15         =   185
         dMax_15         =   10
         discInterval_15 =   1
         Radial_13       =   0
         Enum_13         =   16
         ClassName_16    =   "CCWEnum"
         Editor_16       =   17
         ClassName_17    =   "CCWEnumArrayEditor"
         Owner_17        =   13
         Font_13         =   0
         tickopts_13     =   2711
         major_13        =   2
         minor_13        =   1
         Caption_13      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   -2147483640
         Image_18        =   19
         ClassName_19    =   "CCWTextImage"
         szText_19       =   "Count"
         font_19         =   0
         Animator_18     =   0
         Blinker_18      =   0
         LineStyle_5     =   1
         LineWidth_5     =   1
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Axes_1          =   20
         ClassName_20    =   "CCWAxes"
         Array_20        =   2
         Editor_20       =   21
         ClassName_21    =   "CCWGFAxisArrayEditor"
         Owner_21        =   1
         Array[0]_20     =   6
         Array[1]_20     =   13
         DefaultPlot_1   =   22
         ClassName_22    =   "CCWDataPlot"
         opts_22         =   4194367
         Name_22         =   "[Template]"
         C[0]_22         =   65280
         C[1]_22         =   255
         C[2]_22         =   16711680
         C[3]_22         =   16776960
         Event_22        =   2
         X_22            =   6
         Y_22            =   13
         LineStyle_22    =   1
         LineWidth_22    =   1
         BasePlot_22     =   0
         DefaultXInc_22  =   1
         DefaultPlotPerRow_22=   -1  'True
         Cursors_1       =   23
         ClassName_23    =   "CCWCursors"
         Editor_23       =   24
         ClassName_24    =   "CCWGFCursorArrayEditor"
         Owner_24        =   1
         TrackMode_1     =   11
         GraphBackground_1=   0
         GraphFrame_1    =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         C[0]_25         =   16777215
         C[1]_25         =   16777215
         Image_25        =   26
         ClassName_26    =   "CCWPictImage"
         opts_26         =   1280
         Rows_26         =   1
         Cols_26         =   1
         F_26            =   16777215
         B_26            =   16777215
         ColorReplaceWith_26=   8421504
         ColorReplace_26 =   8421504
         Tolerance_26    =   2
         Animator_25     =   0
         Blinker_25      =   0
         PlotFrame_1     =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   62
         C[0]_27         =   16777215
         C[1]_27         =   16777215
         Image_27        =   28
         ClassName_28    =   "CCWPictImage"
         opts_28         =   1280
         Rows_28         =   1
         Cols_28         =   1
         Pict_28         =   1
         F_28            =   16777215
         B_28            =   16777215
         ColorReplaceWith_28=   8421504
         ColorReplace_28 =   8421504
         Tolerance_28    =   2
         Animator_27     =   0
         Blinker_27      =   0
         Caption_1       =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   62
         C[0]_29         =   -2147483640
         Image_29        =   30
         ClassName_30    =   "CCWTextImage"
         szText_30       =   "NET Error Histogram"
         font_30         =   0
         Animator_29     =   0
         Blinker_29      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   31
         ClassName_31    =   "CCWBindingHolderArray"
         Editor_31       =   32
         ClassName_32    =   "CCWBindingHolderArrayEditor"
         Owner_32        =   1
         Annotations_1   =   33
         ClassName_33    =   "CCWAnnotations"
         Editor_33       =   34
         ClassName_34    =   "CCWAnnotationArrayEditor"
         Owner_34        =   1
         AnnotationTemplate_1=   35
         ClassName_35    =   "CCWAnnotation"
         opts_35         =   63
         Name_35         =   "[Template]"
         Plot_35         =   22
         Text_35         =   "[Template]"
         TextXPoint_35   =   6.7
         TextYPoint_35   =   6.7
         TextColor_35    =   16777215
         TextFont_35     =   36
         ClassName_36    =   "CCWFont"
         bFont_36        =   -1  'True
         BeginProperty Font_36 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShapeXPoints_35 =   37
         ClassName_37    =   "CDataBuffer"
         Type_37         =   5
         m_cDims;_37     =   1
         m_cElts_37      =   1
         Element[0]_37   =   3.3
         ShapeYPoints_35 =   38
         ClassName_38    =   "CDataBuffer"
         Type_38         =   5
         m_cDims;_38     =   1
         m_cElts_38      =   1
         Element[0]_38   =   3.3
         ShapeFillColor_35=   16777215
         ShapeLineColor_35=   16777215
         ShapeLineWidth_35=   1
         ShapeLineStyle_35=   1
         ShapePointStyle_35=   10
         ShapeImage_35   =   39
         ClassName_39    =   "CCWDrawObj"
         opts_39         =   62
         Image_39        =   40
         ClassName_40    =   "CCWPictImage"
         opts_40         =   1280
         Rows_40         =   1
         Cols_40         =   1
         Pict_40         =   7
         F_40            =   -2147483633
         B_40            =   -2147483633
         ColorReplaceWith_40=   8421504
         ColorReplace_40 =   8421504
         Tolerance_40    =   2
         Animator_39     =   0
         Blinker_39      =   0
         ArrowVisible_35 =   -1  'True
         ArrowColor_35   =   16777215
         ArrowWidth_35   =   1
         ArrowLineStyle_35=   1
         ArrowHeadStyle_35=   1
      End
   End
   Begin VB.Frame fraScores 
      BackColor       =   &H00FFFFFF&
      Height          =   7575
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin CWUIControlsLib.CWGraph ctlMassVsMZResidual 
         Height          =   4455
         Left            =   1560
         TabIndex        =   118
         Top             =   4200
         Width           =   6135
         _Version        =   393218
         _ExtentX        =   10821
         _ExtentY        =   7858
         _StockProps     =   71
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   393218
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   62
         C[0]_1          =   16777215
         C[1]_1          =   16777215
         Event_1         =   2
         ClassName_2     =   "CCWGFPlotEvent"
         Owner_2         =   1
         Plots_1         =   3
         ClassName_3     =   "CCWDataPlots"
         Array_3         =   3
         Editor_3        =   4
         ClassName_4     =   "CCWGFPlotArrayEditor"
         Owner_4         =   1
         Array[0]_3      =   5
         ClassName_5     =   "CCWDataPlot"
         opts_5          =   4194367
         Name_5          =   "Zero-line"
         C[0]_5          =   0
         C[1]_5          =   255
         C[2]_5          =   16711680
         C[3]_5          =   16776960
         Event_5         =   2
         X_5             =   6
         ClassName_6     =   "CCWAxis"
         opts_6          =   1599
         Name_6          =   "Scan #"
         Orientation_6   =   2944
         format_6        =   7
         ClassName_7     =   "CCWFormat"
         Scale_6         =   8
         ClassName_8     =   "CCWScale"
         opts_8          =   90112
         rMin_8          =   27
         rMax_8          =   397
         dMax_8          =   10
         discInterval_8  =   1
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   0
         tickopts_6      =   2711
         major_6         =   2
         minor_6         =   1
         Caption_6       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   -2147483640
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         font_12         =   0
         Animator_11     =   0
         Blinker_11      =   0
         Y_5             =   13
         ClassName_13    =   "CCWAxis"
         opts_13         =   1599
         Name_13         =   "Mass Residual"
         Orientation_13  =   2067
         format_13       =   14
         ClassName_14    =   "CCWFormat"
         Scale_13        =   15
         ClassName_15    =   "CCWScale"
         opts_15         =   122880
         rMin_15         =   28
         rMax_15         =   267
         dMax_15         =   10
         discInterval_15 =   1
         Radial_13       =   0
         Enum_13         =   16
         ClassName_16    =   "CCWEnum"
         Editor_16       =   17
         ClassName_17    =   "CCWEnumArrayEditor"
         Owner_17        =   13
         Font_13         =   0
         tickopts_13     =   2711
         major_13        =   2
         minor_13        =   1
         Caption_13      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   -2147483640
         Image_18        =   19
         ClassName_19    =   "CCWTextImage"
         font_19         =   0
         Animator_18     =   0
         Blinker_18      =   0
         LineStyle_5     =   5
         LineWidth_5     =   1
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Array[1]_3      =   20
         ClassName_20    =   "CCWDataPlot"
         opts_20         =   4194367
         Name_20         =   "UMCErrors"
         C[0]_20         =   65280
         C[1]_20         =   14090347
         C[2]_20         =   16711680
         C[3]_20         =   16776960
         Event_20        =   2
         X_20            =   6
         Y_20            =   13
         PointStyle_20   =   21
         LineWidth_20    =   4
         BasePlot_20     =   0
         DefaultXInc_20  =   1
         DefaultPlotPerRow_20=   -1  'True
         Array[2]_3      =   21
         ClassName_21    =   "CCWDataPlot"
         opts_21         =   4194367
         Name_21         =   "TransformFunc"
         C[0]_21         =   8388736
         C[1]_21         =   8388736
         C[2]_21         =   16711680
         C[3]_21         =   16776960
         Event_21        =   2
         X_21            =   6
         Y_21            =   13
         LineStyle_21    =   1
         LineWidth_21    =   2
         BasePlot_21     =   0
         DefaultXInc_21  =   1
         DefaultPlotPerRow_21=   -1  'True
         Axes_1          =   22
         ClassName_22    =   "CCWAxes"
         Array_22        =   2
         Editor_22       =   23
         ClassName_23    =   "CCWGFAxisArrayEditor"
         Owner_23        =   1
         Array[0]_22     =   6
         Array[1]_22     =   13
         DefaultPlot_1   =   24
         ClassName_24    =   "CCWDataPlot"
         opts_24         =   4194367
         Name_24         =   "[Template]"
         C[0]_24         =   65280
         C[1]_24         =   255
         C[2]_24         =   16711680
         C[3]_24         =   16776960
         Event_24        =   2
         X_24            =   6
         Y_24            =   13
         PointStyle_24   =   16
         LineWidth_24    =   1
         BasePlot_24     =   0
         DefaultXInc_24  =   1
         DefaultPlotPerRow_24=   -1  'True
         Cursors_1       =   25
         ClassName_25    =   "CCWCursors"
         Editor_25       =   26
         ClassName_26    =   "CCWGFCursorArrayEditor"
         Owner_26        =   1
         TrackMode_1     =   10
         GraphBackground_1=   0
         GraphFrame_1    =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   62
         C[0]_27         =   16777215
         C[1]_27         =   16777215
         Image_27        =   28
         ClassName_28    =   "CCWPictImage"
         opts_28         =   1280
         Rows_28         =   1
         Cols_28         =   1
         F_28            =   16777215
         B_28            =   16777215
         ColorReplaceWith_28=   8421504
         ColorReplace_28 =   8421504
         Tolerance_28    =   2
         Animator_27     =   0
         Blinker_27      =   0
         PlotFrame_1     =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   62
         C[0]_29         =   16777215
         C[1]_29         =   16777215
         Image_29        =   30
         ClassName_30    =   "CCWPictImage"
         opts_30         =   1280
         Rows_30         =   1
         Cols_30         =   1
         Pict_30         =   1
         F_30            =   16777215
         B_30            =   16777215
         ColorReplaceWith_30=   8421504
         ColorReplace_30 =   8421504
         Tolerance_30    =   2
         Animator_29     =   0
         Blinker_29      =   0
         Caption_1       =   31
         ClassName_31    =   "CCWDrawObj"
         opts_31         =   62
         C[0]_31         =   -2147483640
         Image_31        =   32
         ClassName_32    =   "CCWTextImage"
         szText_32       =   "Mass Error vs m/z"
         font_32         =   0
         Animator_31     =   0
         Blinker_31      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   33
         ClassName_33    =   "CCWBindingHolderArray"
         Editor_33       =   34
         ClassName_34    =   "CCWBindingHolderArrayEditor"
         Owner_34        =   1
         Annotations_1   =   35
         ClassName_35    =   "CCWAnnotations"
         Editor_35       =   36
         ClassName_36    =   "CCWAnnotationArrayEditor"
         Owner_36        =   1
         AnnotationTemplate_1=   37
         ClassName_37    =   "CCWAnnotation"
         opts_37         =   63
         Name_37         =   "[Template]"
         Plot_37         =   38
         ClassName_38    =   "CCWDataPlot"
         opts_38         =   4194367
         Name_38         =   "[Template]"
         C[0]_38         =   65280
         C[1]_38         =   255
         C[2]_38         =   16711680
         C[3]_38         =   16776960
         Event_38        =   2
         X_38            =   6
         Y_38            =   13
         LineStyle_38    =   1
         LineWidth_38    =   1
         BasePlot_38     =   0
         DefaultXInc_38  =   1
         DefaultPlotPerRow_38=   -1  'True
         Text_37         =   "[Template]"
         TextXPoint_37   =   6.7
         TextYPoint_37   =   6.7
         TextColor_37    =   16777215
         TextFont_37     =   39
         ClassName_39    =   "CCWFont"
         bFont_39        =   -1  'True
         BeginProperty Font_39 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShapeXPoints_37 =   40
         ClassName_40    =   "CDataBuffer"
         Type_40         =   5
         m_cDims;_40     =   1
         m_cElts_40      =   1
         Element[0]_40   =   3.3
         ShapeYPoints_37 =   41
         ClassName_41    =   "CDataBuffer"
         Type_41         =   5
         m_cDims;_41     =   1
         m_cElts_41      =   1
         Element[0]_41   =   3.3
         ShapeFillColor_37=   16777215
         ShapeLineColor_37=   16777215
         ShapeLineWidth_37=   1
         ShapeLineStyle_37=   1
         ShapePointStyle_37=   10
         ShapeImage_37   =   42
         ClassName_42    =   "CCWDrawObj"
         opts_42         =   62
         Image_42        =   43
         ClassName_43    =   "CCWPictImage"
         opts_43         =   1280
         Rows_43         =   1
         Cols_43         =   1
         Pict_43         =   7
         F_43            =   -2147483633
         B_43            =   -2147483633
         ColorReplaceWith_43=   8421504
         ColorReplace_43 =   8421504
         Tolerance_43    =   2
         Animator_42     =   0
         Blinker_42      =   0
         ArrowVisible_37 =   -1  'True
         ArrowColor_37   =   16777215
         ArrowWidth_37   =   1
         ArrowLineStyle_37=   1
         ArrowHeadStyle_37=   1
      End
      Begin CWUIControlsLib.CWGraph ctlMassVsScanResidual 
         Height          =   4455
         Left            =   2880
         TabIndex        =   116
         Top             =   2880
         Width           =   6135
         _Version        =   393218
         _ExtentX        =   10821
         _ExtentY        =   7858
         _StockProps     =   71
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   393218
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   62
         C[0]_1          =   16777215
         C[1]_1          =   16777215
         Event_1         =   2
         ClassName_2     =   "CCWGFPlotEvent"
         Owner_2         =   1
         Plots_1         =   3
         ClassName_3     =   "CCWDataPlots"
         Array_3         =   3
         Editor_3        =   4
         ClassName_4     =   "CCWGFPlotArrayEditor"
         Owner_4         =   1
         Array[0]_3      =   5
         ClassName_5     =   "CCWDataPlot"
         opts_5          =   4194367
         Name_5          =   "Zero-line"
         C[0]_5          =   0
         C[1]_5          =   255
         C[2]_5          =   16711680
         C[3]_5          =   16776960
         Event_5         =   2
         X_5             =   6
         ClassName_6     =   "CCWAxis"
         opts_6          =   1599
         Name_6          =   "Scan #"
         Orientation_6   =   2944
         format_6        =   7
         ClassName_7     =   "CCWFormat"
         Scale_6         =   8
         ClassName_8     =   "CCWScale"
         opts_8          =   90112
         rMin_8          =   27
         rMax_8          =   397
         dMax_8          =   10
         discInterval_8  =   1
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   0
         tickopts_6      =   2711
         major_6         =   2
         minor_6         =   1
         Caption_6       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   -2147483640
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         font_12         =   0
         Animator_11     =   0
         Blinker_11      =   0
         Y_5             =   13
         ClassName_13    =   "CCWAxis"
         opts_13         =   1599
         Name_13         =   "Mass Residual"
         Orientation_13  =   2067
         format_13       =   14
         ClassName_14    =   "CCWFormat"
         Scale_13        =   15
         ClassName_15    =   "CCWScale"
         opts_15         =   122880
         rMin_15         =   28
         rMax_15         =   267
         dMax_15         =   10
         discInterval_15 =   1
         Radial_13       =   0
         Enum_13         =   16
         ClassName_16    =   "CCWEnum"
         Editor_16       =   17
         ClassName_17    =   "CCWEnumArrayEditor"
         Owner_17        =   13
         Font_13         =   0
         tickopts_13     =   2711
         major_13        =   2
         minor_13        =   1
         Caption_13      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   -2147483640
         Image_18        =   19
         ClassName_19    =   "CCWTextImage"
         font_19         =   0
         Animator_18     =   0
         Blinker_18      =   0
         LineStyle_5     =   5
         LineWidth_5     =   1
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Array[1]_3      =   20
         ClassName_20    =   "CCWDataPlot"
         opts_20         =   4194367
         Name_20         =   "UMCErrors"
         C[0]_20         =   65280
         C[1]_20         =   32768
         C[2]_20         =   16711680
         C[3]_20         =   16776960
         Event_20        =   2
         X_20            =   6
         Y_20            =   13
         PointStyle_20   =   21
         LineWidth_20    =   4
         BasePlot_20     =   0
         DefaultXInc_20  =   1
         DefaultPlotPerRow_20=   -1  'True
         Array[2]_3      =   21
         ClassName_21    =   "CCWDataPlot"
         opts_21         =   4194367
         Name_21         =   "TransformFunc"
         C[0]_21         =   8388736
         C[1]_21         =   8388736
         C[2]_21         =   16711680
         C[3]_21         =   16776960
         Event_21        =   2
         X_21            =   6
         Y_21            =   13
         LineStyle_21    =   1
         LineWidth_21    =   2
         BasePlot_21     =   0
         DefaultXInc_21  =   1
         DefaultPlotPerRow_21=   -1  'True
         Axes_1          =   22
         ClassName_22    =   "CCWAxes"
         Array_22        =   2
         Editor_22       =   23
         ClassName_23    =   "CCWGFAxisArrayEditor"
         Owner_23        =   1
         Array[0]_22     =   6
         Array[1]_22     =   13
         DefaultPlot_1   =   24
         ClassName_24    =   "CCWDataPlot"
         opts_24         =   4194367
         Name_24         =   "[Template]"
         C[0]_24         =   65280
         C[1]_24         =   255
         C[2]_24         =   16711680
         C[3]_24         =   16776960
         Event_24        =   2
         X_24            =   6
         Y_24            =   13
         PointStyle_24   =   16
         LineWidth_24    =   1
         BasePlot_24     =   0
         DefaultXInc_24  =   1
         DefaultPlotPerRow_24=   -1  'True
         Cursors_1       =   25
         ClassName_25    =   "CCWCursors"
         Editor_25       =   26
         ClassName_26    =   "CCWGFCursorArrayEditor"
         Owner_26        =   1
         TrackMode_1     =   10
         GraphBackground_1=   0
         GraphFrame_1    =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   62
         C[0]_27         =   16777215
         C[1]_27         =   16777215
         Image_27        =   28
         ClassName_28    =   "CCWPictImage"
         opts_28         =   1280
         Rows_28         =   1
         Cols_28         =   1
         F_28            =   16777215
         B_28            =   16777215
         ColorReplaceWith_28=   8421504
         ColorReplace_28 =   8421504
         Tolerance_28    =   2
         Animator_27     =   0
         Blinker_27      =   0
         PlotFrame_1     =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   62
         C[0]_29         =   16777215
         C[1]_29         =   16777215
         Image_29        =   30
         ClassName_30    =   "CCWPictImage"
         opts_30         =   1280
         Rows_30         =   1
         Cols_30         =   1
         Pict_30         =   1
         F_30            =   16777215
         B_30            =   16777215
         ColorReplaceWith_30=   8421504
         ColorReplace_30 =   8421504
         Tolerance_30    =   2
         Animator_29     =   0
         Blinker_29      =   0
         Caption_1       =   31
         ClassName_31    =   "CCWDrawObj"
         opts_31         =   62
         C[0]_31         =   -2147483640
         Image_31        =   32
         ClassName_32    =   "CCWTextImage"
         szText_32       =   "Mass Error vs Scan #"
         font_32         =   0
         Animator_31     =   0
         Blinker_31      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   33
         ClassName_33    =   "CCWBindingHolderArray"
         Editor_33       =   34
         ClassName_34    =   "CCWBindingHolderArrayEditor"
         Owner_34        =   1
         Annotations_1   =   35
         ClassName_35    =   "CCWAnnotations"
         Editor_35       =   36
         ClassName_36    =   "CCWAnnotationArrayEditor"
         Owner_36        =   1
         AnnotationTemplate_1=   37
         ClassName_37    =   "CCWAnnotation"
         opts_37         =   63
         Name_37         =   "[Template]"
         Plot_37         =   38
         ClassName_38    =   "CCWDataPlot"
         opts_38         =   4194367
         Name_38         =   "[Template]"
         C[0]_38         =   65280
         C[1]_38         =   255
         C[2]_38         =   16711680
         C[3]_38         =   16776960
         Event_38        =   2
         X_38            =   6
         Y_38            =   13
         LineStyle_38    =   1
         LineWidth_38    =   1
         BasePlot_38     =   0
         DefaultXInc_38  =   1
         DefaultPlotPerRow_38=   -1  'True
         Text_37         =   "[Template]"
         TextXPoint_37   =   6.7
         TextYPoint_37   =   6.7
         TextColor_37    =   16777215
         TextFont_37     =   39
         ClassName_39    =   "CCWFont"
         bFont_39        =   -1  'True
         BeginProperty Font_39 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShapeXPoints_37 =   40
         ClassName_40    =   "CDataBuffer"
         Type_40         =   5
         m_cDims;_40     =   1
         m_cElts_40      =   1
         Element[0]_40   =   3.3
         ShapeYPoints_37 =   41
         ClassName_41    =   "CDataBuffer"
         Type_41         =   5
         m_cDims;_41     =   1
         m_cElts_41      =   1
         Element[0]_41   =   3.3
         ShapeFillColor_37=   16777215
         ShapeLineColor_37=   16777215
         ShapeLineWidth_37=   1
         ShapeLineStyle_37=   1
         ShapePointStyle_37=   10
         ShapeImage_37   =   42
         ClassName_42    =   "CCWDrawObj"
         opts_42         =   62
         Image_42        =   43
         ClassName_43    =   "CCWPictImage"
         opts_43         =   1280
         Rows_43         =   1
         Cols_43         =   1
         Pict_43         =   7
         F_43            =   -2147483633
         B_43            =   -2147483633
         ColorReplaceWith_43=   8421504
         ColorReplace_43 =   8421504
         Tolerance_43    =   2
         Animator_42     =   0
         Blinker_42      =   0
         ArrowVisible_37 =   -1  'True
         ArrowColor_37   =   16777215
         ArrowWidth_37   =   1
         ArrowLineStyle_37=   1
         ArrowHeadStyle_37=   1
      End
      Begin VIPER.ctl2DHeatMap ctlFlatSurface 
         Height          =   1455
         Left            =   240
         TabIndex        =   113
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2566
      End
      Begin VB.CheckBox chkSurfaceShowsZScore 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Plot Z-Score"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkShowNet 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Linear Fit"
         Height          =   255
         Left            =   6480
         TabIndex        =   6
         Top             =   0
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkShowTransform 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Transformation Function"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   0
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin CW3DGraphLib.CWGraph3D ctlCWGraphNI 
         Height          =   4575
         Left            =   4800
         TabIndex        =   114
         Top             =   960
         Visible         =   0   'False
         Width           =   5535
         _Version        =   393217
         _ExtentX        =   9763
         _ExtentY        =   8070
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   393217
         Graph3D_0       =   1
         ClassName_1     =   "CCWGraph3DFrame"
         opts_1          =   62
         C[0]_1          =   16777215
         Plots_1         =   2
         ClassName_2     =   "CCWDataPlots"
         Array_2         =   1
         Editor_2        =   3
         ClassName_3     =   "CCWGFPlotArrayEditor"
         Owner_3         =   1
         Array[0]_2      =   4
         ClassName_4     =   "Plot3DSurface"
         opts_4          =   4194367
         Name_4          =   "Plot-1"
         C[0]_4          =   33554432
         C[1]_4          =   33554432
         C[2]_4          =   33554687
         Event_4         =   0
         X_4             =   5
         ClassName_5     =   "CCWAxis3D"
         opts_5          =   1599
         Name_5          =   "XAxis"
         Orientation_5   =   3
         format_5        =   6
         ClassName_6     =   "CCWFormat"
         Scale_5         =   7
         ClassName_7     =   "CCWScale"
         opts_7          =   98304
         dMax_7          =   10
         discInterval_7  =   1
         Radial_5        =   0
         Enum_5          =   8
         ClassName_8     =   "CCWEnum"
         Editor_8        =   9
         ClassName_9     =   "CCWEnumArrayEditor"
         Owner_9         =   5
         Font_5          =   0
         tickopts_5      =   702
         Caption_5       =   10
         ClassName_10    =   "CCWDrawObj"
         opts_10         =   62
         C[0]_10         =   -2147483640
         Image_10        =   11
         ClassName_11    =   "CCWTextImage"
         font_11         =   0
         Animator_10     =   0
         Blinker_10      =   0
         CaptionFont_5   =   12
         ClassName_12    =   "CCWFont"
         bFont_12        =   -1  'True
         BeginProperty Font_12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelFont_5     =   13
         ClassName_13    =   "CCWFont"
         bFont_13        =   -1  'True
         BeginProperty Font_13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelPrecision_5=   -1
         LabelFormat_5   =   108
         TitleNormal_5   =   -1  'True
         LabelNormal_5   =   -1  'True
         TicksNormal_5   =   -1  'True
         TicksOpposite_5 =   -1  'True
         TicksInside_5   =   -1  'True
         MajorDivisions_5=   5
         MinorDivisions_5=   3
         MajorUnitsInterval_5=   2
         MinorUnitsInterval_5=   0.666666666666667
         DataMin_5       =   4.01897889775899E-287
         DataMax_5       =   4.01897889775899E-287
         Y_4             =   14
         ClassName_14    =   "CCWAxis3D"
         opts_14         =   1599
         Name_14         =   "YAxis"
         Orientation_14  =   3
         format_14       =   15
         ClassName_15    =   "CCWFormat"
         Scale_14        =   16
         ClassName_16    =   "CCWScale"
         opts_16         =   98304
         dMax_16         =   10
         discInterval_16 =   1
         Radial_14       =   0
         Enum_14         =   17
         ClassName_17    =   "CCWEnum"
         Editor_17       =   18
         ClassName_18    =   "CCWEnumArrayEditor"
         Owner_18        =   14
         Font_14         =   0
         tickopts_14     =   702
         Caption_14      =   19
         ClassName_19    =   "CCWDrawObj"
         opts_19         =   62
         C[0]_19         =   -2147483640
         Image_19        =   20
         ClassName_20    =   "CCWTextImage"
         font_20         =   0
         Animator_19     =   0
         Blinker_19      =   0
         CaptionFont_14  =   21
         ClassName_21    =   "CCWFont"
         bFont_21        =   -1  'True
         BeginProperty Font_21 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelFont_14    =   22
         ClassName_22    =   "CCWFont"
         bFont_22        =   -1  'True
         BeginProperty Font_22 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelPrecision_14=   -1
         LabelFormat_14  =   108
         AxisType_14     =   1
         TitleNormal_14  =   -1  'True
         LabelNormal_14  =   -1  'True
         TicksNormal_14  =   -1  'True
         TicksOpposite_14=   -1  'True
         TicksInside_14  =   -1  'True
         MajorDivisions_14=   5
         MinorDivisions_14=   3
         MajorUnitsInterval_14=   2
         MinorUnitsInterval_14=   0.666666666666667
         DataMin_14      =   4.05759922206535E-287
         DataMax_14      =   4.05759922206535E-287
         PointStyle_4    =   31
         LineStyle_4     =   1
         Z_4             =   23
         ClassName_23    =   "CCWAxis3D"
         opts_23         =   1599
         Name_23         =   "ZAxis"
         Orientation_23  =   3
         format_23       =   24
         ClassName_24    =   "CCWFormat"
         Scale_23        =   25
         ClassName_25    =   "CCWScale"
         opts_25         =   65536
         dMax_25         =   10
         discInterval_25 =   1
         Radial_23       =   0
         Enum_23         =   26
         ClassName_26    =   "CCWEnum"
         Editor_26       =   27
         ClassName_27    =   "CCWEnumArrayEditor"
         Owner_27        =   23
         Font_23         =   0
         tickopts_23     =   702
         Caption_23      =   28
         ClassName_28    =   "CCWDrawObj"
         opts_28         =   62
         C[0]_28         =   -2147483640
         Image_28        =   29
         ClassName_29    =   "CCWTextImage"
         font_29         =   0
         Animator_28     =   0
         Blinker_28      =   0
         CaptionFont_23  =   30
         ClassName_30    =   "CCWFont"
         bFont_30        =   -1  'True
         BeginProperty Font_30 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelFont_23    =   31
         ClassName_31    =   "CCWFont"
         bFont_31        =   -1  'True
         BeginProperty Font_31 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelPrecision_23=   -1
         LabelFormat_23  =   108
         AxisType_23     =   2
         TitleNormal_23  =   -1  'True
         LabelNormal_23  =   -1  'True
         TicksNormal_23  =   -1  'True
         TicksOpposite_23=   -1  'True
         TicksInside_23  =   -1  'True
         MajorDivisions_23=   5
         MinorDivisions_23=   3
         MajorUnitsInterval_23=   2
         MinorUnitsInterval_23=   0.666666666666667
         DataMin_23      =   4.07178873405752E-287
         DataMax_23      =   4.07178873405752E-287
         ContourData_4   =   32
         ClassName_32    =   "ContourData"
         opts_32         =   62
         Basis_32        =   3
         Levels_32       =   5
         Interval_32     =   1
         UserDefinedLevelList_32=   33
         ClassName_33    =   "Real64Vector"
         opts_33         =   62
         Contours_32     =   34
         ClassName_34    =   "Contours"
         Editor_34       =   35
         ClassName_35    =   "ContourArrayEditor"
         Owner_35        =   32
         LineWidth_4     =   1
         Style_4         =   5
         PointFrequency_4=   1
         PointSize_4     =   3
         ColorMapStyle_4 =   2
         FillStyle_4     =   1
         ColorMapInterpolate_4=   -1  'True
         ColorMapMode_4  =   2
         ColorMapAutoScale_4=   -1  'True
         ColorMapSize_4  =   9
         ColorMapValue[0]_4=   1.70775349666513E-02
         ColorMapColor[0]_4=   33554432
         ColorMapValue[1]_4=   1.26280815122499
         ColorMapColor[1]_4=   50266367
         ColorMapValue[2]_4=   2.50853876748333
         ColorMapColor[2]_4=   50266112
         ColorMapValue[3]_4=   3.75426938374166
         ColorMapColor[3]_4=   50331392
         ColorMapValue[4]_4=   5
         ColorMapColor[4]_4=   33619712
         ColorMapValue[5]_4=   6.24573061625834
         ColorMapColor[5]_4=   33619967
         ColorMapValue[6]_4=   7.49146123251667
         ColorMapColor[6]_4=   33587455
         ColorMapValue[7]_4=   8.73719184877501
         ColorMapColor[7]_4=   33554687
         ColorMapValue[8]_4=   9.98292246503335
         ColorMapColor[8]_4=   50331647
         PlotType_4      =   3
         ColorMapRange_4 =   9.9658449300667
         ColorMapMinValue_4=   1.70775349666513E-02
         Axes_1          =   36
         ClassName_36    =   "CCWAxes"
         Array_36        =   3
         Editor_36       =   37
         ClassName_37    =   "CCWGFAxisArrayEditor"
         Owner_37        =   1
         Array[0]_36     =   5
         Array[1]_36     =   14
         Array[2]_36     =   23
         DefaultPlot_1   =   38
         ClassName_38    =   "Plot3DSurface"
         opts_38         =   4194367
         Name_38         =   "[Template]"
         C[0]_38         =   33554432
         C[1]_38         =   33554432
         C[2]_38         =   33554687
         Event_38        =   0
         X_38            =   5
         Y_38            =   14
         PointStyle_38   =   31
         LineStyle_38    =   1
         Z_38            =   23
         ContourData_38  =   39
         ClassName_39    =   "ContourData"
         opts_39         =   62
         Basis_39        =   3
         Levels_39       =   5
         Interval_39     =   1
         UserDefinedLevelList_39=   40
         ClassName_40    =   "Real64Vector"
         opts_40         =   62
         Contours_39     =   41
         ClassName_41    =   "Contours"
         Editor_41       =   42
         ClassName_42    =   "ContourArrayEditor"
         Owner_42        =   39
         LineWidth_38    =   1
         Style_38        =   5
         PointFrequency_38=   1
         PointSize_38    =   3
         ColorMapStyle_38=   1
         FillStyle_38    =   1
         ColorMapInterpolate_38=   -1  'True
         ColorMapMode_38 =   2
         ColorMapAutoScale_38=   -1  'True
         ColorMapSize_38 =   3
         ColorMapColor[0]_38=   33554432
         ColorMapValue[1]_38=   0.5
         ColorMapColor[1]_38=   33554687
         ColorMapValue[2]_38=   1
         ColorMapColor[2]_38=   50331647
         PlotType_38     =   3
         ColorMapRange_38=   1
         Cursors_1       =   43
         ClassName_43    =   "CCWCursors"
         Editor_43       =   44
         ClassName_44    =   "CCWGFCursorArrayEditor"
         Owner_44        =   1
         TrackMode_1     =   20
         GraphBackground_1=   0
         GraphFrame_1    =   45
         ClassName_45    =   "CCWDrawObj"
         opts_45         =   62
         Image_45        =   46
         ClassName_46    =   "CCWPictImage"
         opts_46         =   1280
         Rows_46         =   1
         Cols_46         =   1
         F_46            =   -2147483633
         B_46            =   -2147483633
         ColorReplaceWith_46=   8421504
         ColorReplace_46 =   8421504
         Tolerance_46    =   2
         Animator_45     =   0
         Blinker_45      =   0
         PlotFrame_1     =   47
         ClassName_47    =   "CCWDrawObj"
         opts_47         =   62
         C[1]_47         =   16777215
         Image_47        =   48
         ClassName_48    =   "CCWPictImage"
         opts_48         =   1280
         Rows_48         =   1
         Cols_48         =   1
         Pict_48         =   1
         F_48            =   -2147483633
         B_48            =   16777215
         ColorReplaceWith_48=   8421504
         ColorReplace_48 =   8421504
         Tolerance_48    =   2
         Animator_47     =   0
         Blinker_47      =   0
         Caption_1       =   49
         ClassName_49    =   "CCWDrawObj"
         opts_49         =   62
         C[0]_49         =   -2147483640
         Image_49        =   50
         ClassName_50    =   "CCWTextImage"
         font_50         =   0
         Animator_49     =   0
         Blinker_49      =   0
         Lights_1        =   51
         ClassName_51    =   "CCWLights"
         Array_51        =   4
         Editor_51       =   52
         ClassName_52    =   "CCWGFLightArrayEditor"
         Owner_52        =   1
         Array[0]_51     =   53
         ClassName_53    =   "CCWLight"
         opts_53         =   62
         Longitude_53    =   45
         Latitude_53     =   45
         Distance_53     =   1
         Array[1]_51     =   54
         ClassName_54    =   "CCWLight"
         opts_54         =   62
         Longitude_54    =   45
         Latitude_54     =   45
         Distance_54     =   1
         Array[2]_51     =   55
         ClassName_55    =   "CCWLight"
         opts_55         =   62
         Longitude_55    =   45
         Latitude_55     =   45
         Distance_55     =   1
         Array[3]_51     =   56
         ClassName_56    =   "CCWLight"
         opts_56         =   62
         Longitude_56    =   45
         Latitude_56     =   45
         Distance_56     =   1
         FastDraw_1      =   -1  'True
         GridSmoothing_1 =   -1  'True
         GridXY_1        =   -1  'True
         GridXZ_1        =   -1  'True
         viewMode_1      =   3
         aspectRatio_1   =   1.20983606557377
         autoDistance_1  =   -1  'True
         eyeLongitude_1  =   45
         eyeLatitude_1   =   45
         eyeDistance_1   =   1.1
         xCenter_1       =   0.5
         yCenter_1       =   0.5
         zCenter_1       =   0.5
         centerLongitude_1=   110
         ClipData_1      =   -1  'True
      End
      Begin CWUIControlsLib.CWGraph ctlNETResidual 
         Height          =   4455
         Left            =   480
         TabIndex        =   115
         Top             =   2280
         Width           =   6135
         _Version        =   393218
         _ExtentX        =   10821
         _ExtentY        =   7858
         _StockProps     =   71
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   393218
         Graph_0         =   1
         ClassName_1     =   "CCWGraphFrame"
         opts_1          =   62
         C[0]_1          =   16777215
         C[1]_1          =   16777215
         Event_1         =   2
         ClassName_2     =   "CCWGFPlotEvent"
         Owner_2         =   1
         Plots_1         =   3
         ClassName_3     =   "CCWDataPlots"
         Array_3         =   3
         Editor_3        =   4
         ClassName_4     =   "CCWGFPlotArrayEditor"
         Owner_4         =   1
         Array[0]_3      =   5
         ClassName_5     =   "CCWDataPlot"
         opts_5          =   4194367
         Name_5          =   "Zero-line"
         C[0]_5          =   0
         C[1]_5          =   255
         C[2]_5          =   16711680
         C[3]_5          =   16776960
         Event_5         =   2
         X_5             =   6
         ClassName_6     =   "CCWAxis"
         opts_6          =   1599
         Name_6          =   "Scan #"
         Orientation_6   =   2944
         format_6        =   7
         ClassName_7     =   "CCWFormat"
         Scale_6         =   8
         ClassName_8     =   "CCWScale"
         opts_8          =   90112
         rMin_8          =   27
         rMax_8          =   397
         dMax_8          =   10
         discInterval_8  =   1
         Radial_6        =   0
         Enum_6          =   9
         ClassName_9     =   "CCWEnum"
         Editor_9        =   10
         ClassName_10    =   "CCWEnumArrayEditor"
         Owner_10        =   6
         Font_6          =   0
         tickopts_6      =   2711
         major_6         =   2
         minor_6         =   1
         Caption_6       =   11
         ClassName_11    =   "CCWDrawObj"
         opts_11         =   62
         C[0]_11         =   -2147483640
         Image_11        =   12
         ClassName_12    =   "CCWTextImage"
         font_12         =   0
         Animator_11     =   0
         Blinker_11      =   0
         Y_5             =   13
         ClassName_13    =   "CCWAxis"
         opts_13         =   1599
         Name_13         =   "NET Residual"
         Orientation_13  =   2067
         format_13       =   14
         ClassName_14    =   "CCWFormat"
         Scale_13        =   15
         ClassName_15    =   "CCWScale"
         opts_15         =   122880
         rMin_15         =   28
         rMax_15         =   267
         dMax_15         =   10
         discInterval_15 =   1
         Radial_13       =   0
         Enum_13         =   16
         ClassName_16    =   "CCWEnum"
         Editor_16       =   17
         ClassName_17    =   "CCWEnumArrayEditor"
         Owner_17        =   13
         Font_13         =   0
         tickopts_13     =   2711
         major_13        =   2
         minor_13        =   1
         Caption_13      =   18
         ClassName_18    =   "CCWDrawObj"
         opts_18         =   62
         C[0]_18         =   -2147483640
         Image_18        =   19
         ClassName_19    =   "CCWTextImage"
         font_19         =   0
         Animator_18     =   0
         Blinker_18      =   0
         LineStyle_5     =   5
         LineWidth_5     =   1
         BasePlot_5      =   0
         DefaultXInc_5   =   1
         DefaultPlotPerRow_5=   -1  'True
         Array[1]_3      =   20
         ClassName_20    =   "CCWDataPlot"
         opts_20         =   4194367
         Name_20         =   "UMCErrors"
         C[0]_20         =   65280
         C[1]_20         =   255
         C[2]_20         =   16711680
         C[3]_20         =   16776960
         Event_20        =   2
         X_20            =   6
         Y_20            =   13
         PointStyle_20   =   21
         LineWidth_20    =   4
         BasePlot_20     =   0
         DefaultXInc_20  =   1
         DefaultPlotPerRow_20=   -1  'True
         Array[2]_3      =   21
         ClassName_21    =   "CCWDataPlot"
         opts_21         =   4194367
         Name_21         =   "TransformFunc"
         C[0]_21         =   16711680
         C[1]_21         =   16711680
         C[2]_21         =   16711680
         C[3]_21         =   16776960
         Event_21        =   2
         X_21            =   6
         Y_21            =   13
         LineStyle_21    =   1
         LineWidth_21    =   2
         BasePlot_21     =   0
         DefaultXInc_21  =   1
         DefaultPlotPerRow_21=   -1  'True
         Axes_1          =   22
         ClassName_22    =   "CCWAxes"
         Array_22        =   2
         Editor_22       =   23
         ClassName_23    =   "CCWGFAxisArrayEditor"
         Owner_23        =   1
         Array[0]_22     =   6
         Array[1]_22     =   13
         DefaultPlot_1   =   24
         ClassName_24    =   "CCWDataPlot"
         opts_24         =   4194367
         Name_24         =   "[Template]"
         C[0]_24         =   65280
         C[1]_24         =   255
         C[2]_24         =   16711680
         C[3]_24         =   16776960
         Event_24        =   2
         X_24            =   6
         Y_24            =   13
         PointStyle_24   =   16
         LineWidth_24    =   1
         BasePlot_24     =   0
         DefaultXInc_24  =   1
         DefaultPlotPerRow_24=   -1  'True
         Cursors_1       =   25
         ClassName_25    =   "CCWCursors"
         Editor_25       =   26
         ClassName_26    =   "CCWGFCursorArrayEditor"
         Owner_26        =   1
         TrackMode_1     =   10
         GraphBackground_1=   0
         GraphFrame_1    =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   62
         C[0]_27         =   16777215
         C[1]_27         =   16777215
         Image_27        =   28
         ClassName_28    =   "CCWPictImage"
         opts_28         =   1280
         Rows_28         =   1
         Cols_28         =   1
         F_28            =   16777215
         B_28            =   16777215
         ColorReplaceWith_28=   8421504
         ColorReplace_28 =   8421504
         Tolerance_28    =   2
         Animator_27     =   0
         Blinker_27      =   0
         PlotFrame_1     =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   62
         C[0]_29         =   16777215
         C[1]_29         =   16777215
         Image_29        =   30
         ClassName_30    =   "CCWPictImage"
         opts_30         =   1280
         Rows_30         =   1
         Cols_30         =   1
         Pict_30         =   1
         F_30            =   16777215
         B_30            =   16777215
         ColorReplaceWith_30=   8421504
         ColorReplace_30 =   8421504
         Tolerance_30    =   2
         Animator_29     =   0
         Blinker_29      =   0
         Caption_1       =   31
         ClassName_31    =   "CCWDrawObj"
         opts_31         =   62
         C[0]_31         =   -2147483640
         Image_31        =   32
         ClassName_32    =   "CCWTextImage"
         szText_32       =   "Net Error vs Scan #"
         font_32         =   0
         Animator_31     =   0
         Blinker_31      =   0
         DefaultXInc_1   =   1
         DefaultPlotPerRow_1=   -1  'True
         Bindings_1      =   33
         ClassName_33    =   "CCWBindingHolderArray"
         Editor_33       =   34
         ClassName_34    =   "CCWBindingHolderArrayEditor"
         Owner_34        =   1
         Annotations_1   =   35
         ClassName_35    =   "CCWAnnotations"
         Editor_35       =   36
         ClassName_36    =   "CCWAnnotationArrayEditor"
         Owner_36        =   1
         AnnotationTemplate_1=   37
         ClassName_37    =   "CCWAnnotation"
         opts_37         =   63
         Name_37         =   "[Template]"
         Plot_37         =   38
         ClassName_38    =   "CCWDataPlot"
         opts_38         =   4194367
         Name_38         =   "[Template]"
         C[0]_38         =   65280
         C[1]_38         =   255
         C[2]_38         =   16711680
         C[3]_38         =   16776960
         Event_38        =   2
         X_38            =   6
         Y_38            =   13
         LineStyle_38    =   1
         LineWidth_38    =   1
         BasePlot_38     =   0
         DefaultXInc_38  =   1
         DefaultPlotPerRow_38=   -1  'True
         Text_37         =   "[Template]"
         TextXPoint_37   =   6.7
         TextYPoint_37   =   6.7
         TextColor_37    =   16777215
         TextFont_37     =   39
         ClassName_39    =   "CCWFont"
         bFont_39        =   -1  'True
         BeginProperty Font_39 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShapeXPoints_37 =   40
         ClassName_40    =   "CDataBuffer"
         Type_40         =   5
         m_cDims;_40     =   1
         m_cElts_40      =   1
         Element[0]_40   =   3.3
         ShapeYPoints_37 =   41
         ClassName_41    =   "CDataBuffer"
         Type_41         =   5
         m_cDims;_41     =   1
         m_cElts_41      =   1
         Element[0]_41   =   3.3
         ShapeFillColor_37=   16777215
         ShapeLineColor_37=   16777215
         ShapeLineWidth_37=   1
         ShapeLineStyle_37=   1
         ShapePointStyle_37=   10
         ShapeImage_37   =   42
         ClassName_42    =   "CCWDrawObj"
         opts_42         =   62
         Image_42        =   43
         ClassName_43    =   "CCWPictImage"
         opts_43         =   1280
         Rows_43         =   1
         Cols_43         =   1
         Pict_43         =   7
         F_43            =   -2147483633
         B_43            =   -2147483633
         ColorReplaceWith_43=   8421504
         ColorReplace_43 =   8421504
         Tolerance_43    =   2
         Animator_42     =   0
         Blinker_42      =   0
         ArrowVisible_37 =   -1  'True
         ArrowColor_37   =   16777215
         ArrowWidth_37   =   1
         ArrowLineStyle_37=   1
         ArrowHeadStyle_37=   1
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveFlatView 
         Caption         =   "Save &Flat View"
         Begin VB.Menu mnuFileSaveFlatViewAsEMF 
            Caption         =   "as &EMF"
         End
         Begin VB.Menu mnuFileSaveFlatViewAsPNG 
            Caption         =   "as &PNG"
         End
      End
      Begin VB.Menu mnuFileSave3DView 
         Caption         =   "Save &3D View"
         Begin VB.Menu mnuFileSave3DViewAsEMF 
            Caption         =   "as &EMF"
         End
         Begin VB.Menu mnuFileSave3DViewAsPNG 
            Caption         =   "as &PNG"
         End
      End
      Begin VB.Menu mnuFileSaveNETResidualsPlot 
         Caption         =   "Save &NET Residuals Plot"
         Begin VB.Menu mnuFileSaveResidualsPlotAsEMF 
            Caption         =   "as &EMF"
         End
         Begin VB.Menu mnuFileSaveResidualsPlotAsPNG 
            Caption         =   "as &PNG"
         End
      End
      Begin VB.Menu mnuFileSaveMassVsScanResidualsPlot 
         Caption         =   "Save &Mass vs. Scan Residuals Plot"
         Begin VB.Menu mnuFileSaveMassVsScanResidualsPlotAsEMF 
            Caption         =   "as &EMF"
         End
         Begin VB.Menu mnuFileSaveMassVsScanResidualsPlotAsPNG 
            Caption         =   "as &PNG"
         End
      End
      Begin VB.Menu mnuFileSaveMassVsMZResidualsPlot 
         Caption         =   "Save Mass vs. m/&z Residuals Plot"
         Begin VB.Menu mnuFileSaveMassVsMZResidualsPlotAsEMF 
            Caption         =   "as &EMF"
         End
         Begin VB.Menu mnuFileSaveMassVsMZResidualsPlotAsPNG 
            Caption         =   "as &PNG"
         End
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditAlignStart 
         Caption         =   "&Start Align (warping)"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyFlatViewToClipboard 
         Caption         =   "&Copy Flat View to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditCopy3DViewToClipboard 
         Caption         =   "Copy &3D View to Clipboard"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditCopyResidualsPlotToClipboard 
         Caption         =   "Copy &NET Residual Plot to Clipboard"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditCopyMassVsScanResidualsPlotToClipboard 
         Caption         =   "Copy &Mass vs. Scan Residual Plot to Clipboard"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditCopyMassVsMZResidualsPlotToClipboard 
         Caption         =   "Copy &Mass vs. m/z Residual Plot to Clipboard"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyNETResidualValues 
         Caption         =   "Copy NET Residual Values to Clipboard"
      End
      Begin VB.Menu mnuEditCopyMassMassVsScanResidualValues 
         Caption         =   "Copy Mass vs. Scan Residual Values to Clipboard"
      End
      Begin VB.Menu mnuEditCopylMassVsMZResidualValues 
         Caption         =   "Copy Mass vs. m/z Residual Values to Clipboard"
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyAlignmentScoresToClipboard 
         Caption         =   "Copy Alignment Scores to Clipboard"
      End
      Begin VB.Menu mnuEditCopyAlignmentFunctionToClipboard 
         Caption         =   "Copy Alignment Function to Clipboard"
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu cmdEditShowDefaults 
         Caption         =   "Show Default &Options"
      End
      Begin VB.Menu mnuEditLoadPMTsFromFile 
         Caption         =   "Load &MT Tags From File"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLoadUMCsFromFile 
         Caption         =   "Load &UMCs From File"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFlatFiew 
         Caption         =   "&Flat View"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuView2DView 
         Caption         =   "&2D View"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuView3DView 
         Caption         =   "&3D View"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewNETResidualsPlot 
         Caption         =   "&NET Residuals Plot (toggle mode)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewMassVsScanResidualsPlot 
         Caption         =   "&Mass Residuals vs Scan Plot (toggle mode)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewMassVsMZResidualsPlot 
         Caption         =   "&Mass Residuals vs. m/z Plot (toggle mode)"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuPMTTags 
      Caption         =   "&MT Tags"
      Begin VB.Menu mnuPMTTagsLoadFromDB 
         Caption         =   "Load MT Tags from &DB"
      End
      Begin VB.Menu mnuPMTTagsLoadFromLegacyDB 
         Caption         =   "Load MT Tags from &Legacy DB"
      End
      Begin VB.Menu mnuPMTTagsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPMTTagsLoadMTStats 
         Caption         =   "Load MT &FScore Stats"
      End
      Begin VB.Menu mnuPMTTagsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPMTTagsStatus 
         Caption         =   "MT Tags &Status"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowSizeToA 
         Caption         =   "Set to Size &A"
      End
      Begin VB.Menu mnuWindowSizeToB 
         Caption         =   "Set to Size &B"
      End
      Begin VB.Menu mnuWindowSizeToC 
         Caption         =   "Set to Size &C"
      End
      Begin VB.Menu mnuWindowSizeToD 
         Caption         =   "Set to Size &D"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMSAlign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const MAX_NET_TOL As Single = 0.15

Private Const MAXIMUM_MASS_SHIFT_PPM As Long = 10000
Private Const MAXIMUM_PERCENTAGE_SHIFTED_OVER_MAX As Single = 5     ' If more than 5% of the data is shifted more than MAXIMUM_MASS_SHIFT_PPM, then all shifting will be reverted

Private Const PMT_COLUMN_COUNT As Integer = 4
Private Enum PMTColumnConstants
    pccNET = 0
    pccMonoisotopicMass = 1
    pccObservationCount = 2             ' Must be an integer
    pccPMTTagID = 3                     ' Must be an integer
End Enum

Private Const FEATURE_COLUMN_COUNT As Integer = 7
Private Enum FeatureColumnConstants
    fccClassMZ = 0
    fccClassMass = 1
    fccScanClassRep = 2                 ' Must be an integer
    fccClassAbundance = 3
    fccPairIndex = 4                    ' Must be an integer; -1 if not a member of a pair; if part of one or more pairs, then only lists the PairIndex of the first pair the UMC belongs to
    fccPMTTagID = 5                     ' Must be an integer
    fccUMCID = 6                        ' Must be an integer
End Enum

Private Enum pvmPlotViewModeConstants
    pvmFlatView = 0
    pvm2DView = 1
    pvm3DView = 2
    pvmLinearFitNETResidualsPlot = 3
    pvmWarpedFitNETResidualsPlot = 4
    pvmMassResidualsScanPlot = 5
    pvmMassResidualsCorrectedScanPlot = 6
    pvmMassResidualsMZPlot = 7
    pvmMassResidualsCorrectedMZPlot = 8
End Enum

Private Enum UMCFileColumnConstants
    ufcUMCID = 0
    ufcStartScan = 1
    ufcEndScan = 2
    ufcCenterScan = 3
    ufcClassMonoisotopicMass = 5
    ufcClassAbundance = 9
    ufcClassMZ = 13
    ufcMemberCount = 14
    ufcPairIndex = 17
    ufcPMTTagID = 23
End Enum

Public Enum MassMatchProcessingStateConstants
    pscUninitialized = 0
    pscRunning = 1
    pscComplete = 2
    pscError = 3
    pscAborted = 4
    pscInsufficientMatches = 5
End Enum

Private Enum MassMatchStepsToPerformConstants
    mmsWarpTime = 0                         ' Corresponds to UMCRobustNETModeConstants.UMCRobustNETWarpTime
    mmsWarpTimeAndMass = 1                  ' Corresponds to UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass
End Enum

Private Enum fwsFormWindowSizeConstants
    fwsSizeA = 0
    fwsSizeB = 1
    fwsSizeC = 2
    fwsSizeD = 3
End Enum

' 2D Array, ranging from 0 to mLocalPMTCount-1 in the 1st dimension and 0 to PMT_COLUMN_COUNT-1 in the second dimension
' Columns for the 2nd dimension are given by PMTColumnConstants
Private mLocalPMTCount As Long
Private mLocalPMTs() As Double

' 2D Array; containing a subset of mLocalPMTs, filtered on UMCNetAdjDef.MSWarpOptions.MinimumPMTTagObsCount
Private mLocalPMTsFiltered() As Double

' 2D Array, ranging from 0 to mLocalFeatureCount-1 in the 1st dimension and 0 to FEATURE_COLUMN_COUNT-1 in the second dimension
' Columns for the 2nd dimension are given by FeatureColumnConstants
Private mLocalFeatureCount As Long
Private mLocalFeatures() As Double

Private mLocalFeaturesAreFiltered As Boolean

' This variable is used when calling LCMSWarp multiple times, first with data below a boundary, and then with data above a boundary
Private mSplitWarpIteration As Integer
Private mPauseSplitWarpProcessing As Boolean

' 2D Array of variants, ranging from 0 to UBound(mMatches, 1) in the first dimension and containing 0 to 3 in the 2nd dimension
Private mMatches As Variant
Private Enum mcMatchColumns
    mcUMCIndex = 0
    mcPMTTagIndex = 1
    mcNETError = 2
    mcMassError = 3
End Enum

' 2D Array of variants, ranging from 0 to UBound(mMatchScores, 1) in the first dimension and 0 to UBound(mMatchScores, 2) in the second dimension
' The number of rows corresponds to the number of time slices; the number of columns corresponds to the number of regions the NET scale was divided into
Private mMatchScores As Variant

' 2D Array with the alignment function values, ranging from UBound(mAlignmentFunc, 1) in the first dimension and from 0 to 8 in the 2nd dimension
' Columns for the 2nd dimension are given by afAlignmentFuncColumnConstants
Private mAlignmentFunc As Variant
Private Enum afAlignmentFuncColumnConstants
    TimeSlice = 0
    AlignmentFuncY1 = 1
    AlignmentFuncY2 = 2
    ScanStart = 3
    ScanEnd = 4
    NETStart = 5
    NETEnd = 6
    ScoreA = 7
    ScoreB = 8
End Enum

Private mCalibrationType As rmcUMCRobustNETWarpMassCalibrationType
Private mSplineOrder As Integer

' 2D arrays holding Residual values
' 1a. scan vs. net error using linear function
Private mScanMTNETvsLinearNETResidual() As Double
' 1b. scan vs. net error using warping function
Private mScanMTNETvsCustomNETResidual() As Double
' 1c. scan vs. (net from ms warp - net from linear function)
Private mScanCustomNetVsLinearNETResidual() As Double

' 2a. scan vs. mass error before warping
Private mScanVsMassError() As Double
' 2b. scan vs. mass error after warping
Private mScanVsMassErrorCorrected() As Double

' 3a. m/z vs. mass error before warping
Private mMZVsMassError() As Double
' 3b. m/z vs. mass error after warping
Private mMZVsMassErrorCorrected() As Double

' Scan Number to NET mapping
' 2D array ranging from 0 to UBound(mAlignmentFunc, 1) in the first dimension and 0 to 1 in the 2nd dimension
Private mTransformRT As Variant
Private Enum trcTransformRTColumns
    trcScanNum = 0
    trcNET = 1                          ' Transformed NET for the given scan number
End Enum

' UMC to Transformed NET mapping
' 2D array ranging from 0 to mLocalFeatureCount-1 in the first dimension and 0 to 1 in the 2nd dimension
Private mPepTransformRT As Variant
Private Enum ptcPepTransformRTColumns
    ptcUMCIndex = 0
    ptcUMCNET = 1                       ' Transformed NET for the given UMC index
End Enum

' Mass Error Histogram
' 2D array of doubles with X values in the first dimension and Y values in the second dimension
Private mMassErrorHistogram As Variant
Private Enum mecMassErrorHistogramColumns
    mecMassErrorPPMBin = 0
    mecBinCount = 1
End Enum

' NET Error Histogram
' 2D array of doubles with X values in the first dimension and Y values in the second dimension
Private mNetErrorHistogram As Variant
Private Enum necNETErrorHistogramColumns
    necNETErrorBin = 0
    necBinCount = 1
End Enum


' Column 1 is scan# and Column 2 is Net value
Private mScan2Nets As Variant

' MassMatch object for aligning and matching data
Private mMassMatchObject As New MassMatchCOM.CMassMatchWrapper
Private mMassMatchState As MassMatchProcessingStateConstants
Private mProcessingStartTime As Date
Private mAlignmentFinalizedOrAborted As Boolean

Private mLocalGelUpdated As Boolean

Private mLoading As Boolean
Private mControlsEnabled As Boolean
Private mRecalibratingMassDuringAutoAnalysis As Boolean
Private mAbortRequested As Boolean

Private mMinNet As Double
Private mMaxNet As Double
Private mMinScan As Long
Private mMaxScan As Long
Private mMinMZ As Single
Private mMaxMZ As Single

Private mLinearNETResidualMin As Single
Private mLinearNETResidualMax As Single
Private mMassResidualMin As Single
Private mMassResidualMax As Single

Private mNETLineSlopeSaved As Double
Private mNETLineInterceptSaved As Double

Private mMostRecentPlotViewMode As pvmPlotViewModeConstants
Private mMostRecentNETResidualsViewMode As pvmPlotViewModeConstants
Private mMostRecentMassVsScanResidualsViewMode As pvmPlotViewModeConstants
Private mMostRecentMassVsMZResidualsViewMode As pvmPlotViewModeConstants

Private mResidualPlotPointSize As Integer
Private mResidualPlotTransformationFnLineSize As Integer
Private mUpdatingPlotRanges As Boolean

Public CallerID As Long

Public Property Get LocalGelUpdated() As Boolean
    LocalGelUpdated = mLocalGelUpdated
End Property
Public Property Let LocalGelUpdated(blnValue As Boolean)
    mLocalGelUpdated = blnValue
End Property

Public Property Get MassMatchState() As MassMatchProcessingStateConstants
    MassMatchState = mMassMatchState
End Property

Public Property Get AlignmentFinalizedOrAborted() As Boolean
    AlignmentFinalizedOrAborted = mAlignmentFinalizedOrAborted
End Property

Public Property Get RecalibratingMassDuringAutoAnalysis() As Boolean
    RecalibratingMassDuringAutoAnalysis = mRecalibratingMassDuringAutoAnalysis
End Property
Public Property Let RecalibratingMassDuringAutoAnalysis(blnValue As Boolean)
    mRecalibratingMassDuringAutoAnalysis = blnValue
End Property


Private Function CheckVsMinimum(ByVal lngValue As Long, Optional ByVal lngMinimum As Long = 0) As Long
    If lngValue < lngMinimum Then lngValue = lngMinimum
    CheckVsMinimum = lngValue
End Function

Private Sub ClearLocalFeaturesArray()
    ReDim mLocalFeatures(0, FEATURE_COLUMN_COUNT - 1)
    mLocalFeaturesAreFiltered = True
    txtFeatureCountLoaded.Text = "0"
End Sub

Private Sub ClearLocalPMTsArray()
    ReDim mLocalPMTs(0, PMT_COLUMN_COUNT - 1)
    txtPMTCountLoaded.Text = "0"
End Sub

Private Function ConstructHistoryTextForMSWarpSettings(ByVal strSecondsElapsed As String) As String
    Dim strMessage As String
    
    With UMCNetAdjDef.MSWarpOptions
        strMessage = "MS Warp settings" & _
                     "; Number of sections = " & Trim(.NumberOfSections) & _
                     "; Contraction factor = " & Trim(.ContractionFactor) & _
                     "; Max distortion = " & Trim(.MaxDistortion) & _
                     "; Minimum MT tag Obs Count = " & Trim(.MinimumPMTTagObsCount) & _
                     "; Match promiscuity = " & Trim(.MatchPromiscuity) & _
                     "; Processing time = " & strSecondsElapsed & " seconds"
    End With
    
    ConstructHistoryTextForMSWarpSettings = strMessage
End Function

Private Function ConstructHistoryTextForMassRecalibration() As String
    Dim strMessage As String
    
    With UMCNetAdjDef.MSWarpOptions
        strMessage = "Mass window = " & Trim(Round(.MassWindowPPM, 3)) & " ppm" & _
                     "; Mass calibration type = " & LookupMassCalibrationTypeName(CInt(.MassCalibrationType)) & _
                     "; Spline order = " & Trim(.MassSplineOrder) & _
                     "; # x-axis slices = " & Trim(.MassNumXSlices) & _
                     "; # mass delta bins = " & Trim(.MassNumMassDeltaBins) & _
                     "; Max jump = " & Trim(Round(.MassMaxJump, 1)) & " ppm" & _
                     "; z-score tolerance = " & Trim(Round(.MassZScoreTolerance, 2)) & _
                     "; Use LSQ = " & CStr(.MassUseLSQ)
                     
        If .MassUseLSQ Then
            strMessage = strMessage & "; Outlier z-score = " & Trim(Round(.MassLSQOutlierZScore, 2)) & _
                                      "; Number of knots for LSQ = " & Trim(.MassLSQNumKnots)
        End If
    
    End With
    
    ConstructHistoryTextForMassRecalibration = strMessage
End Function

Private Function ConstructHistoryTextForNETAlignment(ByVal strSecondsElapsed As String) As String
  Dim strMessage As String
  
    strMessage = UMC_NET_ADJ_ITERATION_COUNT & " = 1" & _
                 "; Mass tolerance = ±" & Trim(Round(UMCNetAdjDef.MWTol, 3)) & " ppm" & _
                 "; NET Tolerance = ±" & Format(UMCNetAdjDef.MSWarpOptions.NETTol, "0.000")
                    
    '' If, in the future, we restrict the alignment to the top x% of LC-MS Features, then use this text to record that
    ' If .UMCTopAbuPct >= 0 Then
    '     strMessage = strMessage & "; Restrict to x% of LC-MS Features = " & Trim(.UMCTopAbuPct) & "%"
    ' End If
    
    If MatchesAreValid Then
        strMessage = strMessage & _
                     "; " & UMC_NET_ADJ_UMCs_IN_TOLERANCE & " = " & Trim(mLocalFeatureCount) & _
                     "; " & UMC_NET_ADJ_UMCs_WITH_DB_HITS & " = " & Trim(UBound(mMatches, 1) + 1)
    Else
        strMessage = strMessage & _
                     "; " & UMC_NET_ADJ_UMCs_IN_TOLERANCE & " = " & Trim(mLocalFeatureCount) & _
                     "; " & UMC_NET_ADJ_UMCs_WITH_DB_HITS & " = 0"
    End If
    
                    
    ConstructHistoryTextForNETAlignment = strMessage
End Function

Public Function CopyTwoDimensionalVariantToClipboardOrFile(ByRef varTwoDimensionalArray As Variant, Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long

    ' Returns 0 if success; error number if error
    
    Dim lngRowIndexMax As Long
    Dim lngColumnIndexMax As Long
    Dim strData() As String     ' Ranges from 0 to lngRowIndexMax since row 0 is the header row
    
    Dim lngRowIndex As Long
    Dim lngColumnIndex As Long
    
    Dim fso As FileSystemObject
    Dim tsOutfile As TextStream
    Dim strTextToCopy As String
    
On Error Resume Next
    lngRowIndexMax = UBound(varTwoDimensionalArray, 2)
    lngColumnIndexMax = UBound(varTwoDimensionalArray, 1)
    If lngRowIndexMax = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        CopyTwoDimensionalVariantToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo CopyTwoDimensionalVariantToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    ' Header row is strData(0)
    ' Data is from strData(1) to strData(lngRowIndexMax)
    ReDim strData(0 To lngRowIndexMax + 1)
    
    ' Define the header row
    strData(0) = "Row"
    For lngColumnIndex = 0 To lngColumnIndexMax
        strData(0) = strData(0) & vbTab & "Column" & LTrim(CStr(lngColumnIndex + 1))
    Next lngColumnIndex
    
    ' Fill strData()
    For lngRowIndex = 0 To lngRowIndexMax
        ' Store the first column of data
        strData(lngRowIndex + 1) = "Row" & LTrim(CStr(lngRowIndex + 1))
        
        ' Store the remaining columns
        For lngColumnIndex = 0 To lngColumnIndexMax - 1
            strData(lngRowIndex + 1) = strData(lngRowIndex + 1) & vbTab & Round(varTwoDimensionalArray(lngColumnIndex, lngRowIndex), 4)
        
        Next lngColumnIndex
    Next lngRowIndex
      
    If Len(strFilePath) > 0 Then
        Set fso = New FileSystemObject
        Set tsOutfile = fso.CreateTextFile(strFilePath, True)
        
        For lngRowIndex = 0 To UBound(strData)
           tsOutfile.WriteLine strData(lngRowIndex)
        Next lngRowIndex
        
        tsOutfile.Close
    Else
        strTextToCopy = FlattenStringArray(strData(), UBound(strData) + 1, vbCrLf, False)
        Clipboard.Clear
        Clipboard.SetText strTextToCopy, vbCFText
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    CopyTwoDimensionalVariantToClipboardOrFile = 0
    Exit Function

CopyTwoDimensionalVariantToClipboardOrFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    CopyTwoDimensionalVariantToClipboardOrFile = Err.Number

End Function

Public Function CopyAlignmentFunctionToClipboardOrFile(Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long

    ' Returns 0 if success; error number if error
    
    Dim lngColumnIndexMax As Long
    Dim lngRowIndexMax As Long
    Dim strData() As String     ' Ranges from 0 to lngRowIndexMax since row 0 is the header row
    
    Dim lngColumnIndex As Long
    Dim lngRowIndex As Long
    
    Dim fso As FileSystemObject
    Dim tsOutfile As TextStream
    Dim strTextToCopy As String
    
On Error Resume Next
    lngRowIndexMax = UBound(mAlignmentFunc, 1)
    lngColumnIndexMax = UBound(mAlignmentFunc, 2)
    If lngRowIndexMax = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        CopyAlignmentFunctionToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo CopyAlignmentFunctionToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    ' Header row is strData(0)
    ' Data is from strData(1) to strData(lngColumnIndexMax)
    ReDim strData(0 To lngRowIndexMax + 1)
    
    ' Define the header row
    strData(0) = "TimeSlice" & vbTab & _
                 "AlignmentFuncY1" & vbTab & _
                 "AlignmentFuncY2" & vbTab & _
                 "ScanStart" & vbTab & _
                 "ScanEnd" & vbTab & _
                 "NETStart" & vbTab & _
                 "NETEnd" & vbTab & _
                 "ScoreA" & vbTab & _
                 "ScoreB"

    ' Fill strData()
    For lngRowIndex = 0 To lngRowIndexMax
        ' Store the first column of data
        strData(lngRowIndex + 1) = Round(mAlignmentFunc(lngRowIndex, 0), 4)
        
        ' Store the remaining columns
        For lngColumnIndex = 1 To lngColumnIndexMax
            strData(lngRowIndex + 1) = strData(lngRowIndex + 1) & vbTab & Round(mAlignmentFunc(lngRowIndex, lngColumnIndex), 4)
        Next lngColumnIndex
    Next lngRowIndex
      
    If Len(strFilePath) > 0 Then
        Set fso = New FileSystemObject
        Set tsOutfile = fso.CreateTextFile(strFilePath, True)
        
        For lngRowIndex = 0 To UBound(strData)
           tsOutfile.WriteLine strData(lngRowIndex)
        Next lngRowIndex
        
        tsOutfile.Close
    Else
        strTextToCopy = FlattenStringArray(strData(), UBound(strData) + 1, vbCrLf, False)
        Clipboard.Clear
        Clipboard.SetText strTextToCopy, vbCFText
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    CopyAlignmentFunctionToClipboardOrFile = 0
    Exit Function

CopyAlignmentFunctionToClipboardOrFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    CopyAlignmentFunctionToClipboardOrFile = Err.Number

End Function

Public Function CopyAlignmentScoresToClipboardOrFile(Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    CopyAlignmentScoresToClipboardOrFile = CopyTwoDimensionalVariantToClipboardOrFile(mMatchScores, strFilePath, blnShowMessages)
End Function

Private Sub CopyLocalDataToGel(strSecondsElapsed As String, blnUpdateMassValues As Boolean)

    Const USE_ARIBTRARY_SCAN_NUMBERS As Boolean = False
    
    Dim lngIndex As Long
    Dim lngStartIndex As Long
    
    Dim lngDataCountNew As Long
    Dim lngDataCountOld As Long
    
    Dim lngDataCountTotal As Long
    Dim lngResidualDataTargetIndex As Long
    Dim lngNETDataCountExpected As Long
    
    Dim lngResidualNETTargetIndexStart As Long
    Dim lngResidualMassTargetIndexStart As Long
    
    Dim lngScanCount As Long
    Dim lngScans() As Long
    Dim dblScanNETs() As Double
    
    Dim dblPredictedNETLinear As Double
    Dim dblMeanNETResidual As Double
    Dim dblMassTagNET As Double

    Dim lngScanNumber As Long
    Dim dblCustomNET As Double
    Dim dblMZ As Double
    Dim dblMassErrorCalibrated As Double
      
    Dim intTopAbuPct As Integer         ' Value between 0 and 100; percentage of features to use, ordering features by decreasing abundance
    Dim dblSlope As Double
    Dim dblIntercept As Double
    Dim dblRSquared As Double
    
    Dim lngFeatureCountUsed As Long     ' Number of features used to determine Slope and Intercept
    Dim intRoundingDivisor As Integer
    Dim sngMinMZInitial As Single
    
    Dim strLastGoodPosition As String
    Dim strMessage As String
    
    Dim blnSuccess As Boolean
    Dim dblMassError As Double
   
On Error GoTo CopyLocalDataToGelErrorHandler

    '--------------------------------------------------
    ' First, update the NET values for each scan number
    '--------------------------------------------------
    strLastGoodPosition = "Update the NET values for each scan number"
    If USE_ARIBTRARY_SCAN_NUMBERS Then
        lngScanCount = UBound(mTransformRT, 1) + 100
        ReDim lngScans(lngScanCount - 1)
        For lngIndex = 0 To lngScanCount - 1
            lngScans(lngIndex) = lngIndex
        Next lngIndex
    Else
          ' Create an array listing the known scan numbers
        With GelData(CallerID)
            ' Note: .ScanInfo is 1-based
            
            lngScanCount = UBound(.ScanInfo)
            ReDim lngScans(lngScanCount - 1)
            For lngIndex = 1 To lngScanCount
                lngScans(lngIndex - 1) = .ScanInfo(lngIndex).ScanNumber
            Next lngIndex
        
        End With
        
    End If
     
    strLastGoodPosition = "Call GetNETsFromScans"
    blnSuccess = GetNETsFromScans(lngScans, dblScanNETs)
    
    If blnSuccess Then
        If USE_ARIBTRARY_SCAN_NUMBERS Then
            For lngIndex = 0 To lngScanCount
                If lngIndex Mod 100 = 0 Then
                    Debug.Print lngScans(lngIndex), dblScanNETs(lngIndex)
                End If
            Next lngIndex
        Else
            With GelData(CallerID)
                For lngIndex = 1 To UBound(.ScanInfo)
                    .ScanInfo(lngIndex).CustomNET = dblScanNETs(lngIndex - 1)
                Next lngIndex
            
                ' The last few scans will have invalid NET values; check for this
                ' ToDo: This is a bug that Deep needs to fix
                
                lngStartIndex = UBound(.ScanInfo) * 0.95
                If lngStartIndex < 1 Then lngStartIndex = 1
                
                For lngIndex = lngStartIndex To UBound(.ScanInfo) - 1
                    If .ScanInfo(lngIndex + 1).CustomNET < .ScanInfo(lngIndex).CustomNET Then
                        .ScanInfo(lngIndex + 1).CustomNET = .ScanInfo(lngIndex).CustomNET
                        Debug.Print Now() & ": Updated invalid NET value for scan " & Trim(.ScanInfo(lngIndex + 1).ScanNumber) & " since less than the current scan's NET value"
                    End If
                Next lngIndex
            
            End With
            
            GelData(CallerID).CustomNETsDefined = True
            GelStatus(CallerID).Dirty = True
        End If
    Else
        UpdateStatus "Error updating Scan to NET mapping"
    End If
    
    If blnSuccess And GelData(CallerID).CustomNETsDefined Then
        
        '--------------------------------------------------
        ' Confirm that the feature's class NET matches the new Custom NET values
        '--------------------------------------------------
        
        With GelUMC(CallerID)
            Debug.Assert mLocalFeatureCount - 1 = UBound(mPepTransformRT, 1)
            strLastGoodPosition = "Confirm that the feature's class NET matches the custom NET value"
            For lngIndex = 0 To UBound(mPepTransformRT, 1)
                If GetUMCScanMZAndNET(mPepTransformRT(lngIndex, ptcPepTransformRTColumns.ptcUMCIndex), lngScanNumber, dblMZ, dblCustomNET) Then
                    If Round(dblCustomNET, 2) <> Round(mPepTransformRT(lngIndex, ptcPepTransformRTColumns.ptcUMCNET), 2) Then
                        ' NET values do not agree
                        ''Debug.Assert False
                        Debug.Print Now() & ": NET Values do not agree for scan " & lngScanNumber & ": " & Format(dblCustomNET, "0.0000") & " vs. " & Format(mPepTransformRT(lngIndex, ptcPepTransformRTColumns.ptcUMCNET), "0.0000")
                    End If
                End If
            Next lngIndex
        End With
        
        If MatchesAreValid Then
            lngDataCountNew = UBound(mMatches, 1) + 1
        Else
            lngDataCountNew = 0
        End If
        
        If lngDataCountNew = 0 Then
            If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Or _
               mSplitWarpIteration < 2 Then
                ' Erase the residual value arrays
                ReDim mScanMTNETvsLinearNETResidual(1, 0)
                ReDim mScanMTNETvsCustomNETResidual(1, 0)
                ReDim mScanCustomNetVsLinearNETResidual(1, 0)
                ReDim mScanVsMassError(1, 0)
                ReDim mScanVsMassErrorCorrected(1, 0)
                ReDim mMZVsMassError(1, 0)
                ReDim mMZVsMassErrorCorrected(1, 0)
                    
                mLinearNETResidualMin = 0
                mLinearNETResidualMax = 0
            
            End If
        Else
            
            '--------------------------------------------------
            ' Estimate the Linear NET slope and intercept
            ' Always start with 20% of the LC-MS Features, then step up in increments of 20%
            '--------------------------------------------------
            intTopAbuPct = 20
            strLastGoodPosition = "Estimate the Linear NET"
            Do
                blnSuccess = EstimateLinearNET(intTopAbuPct, dblSlope, dblIntercept, dblRSquared, lngFeatureCountUsed)
                
                intTopAbuPct = intTopAbuPct + 20
            Loop While Not blnSuccess And intTopAbuPct <= 100
            
            ' Now that EstimateLinearNET is done, re-update the status bit for UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
            With GelUMC(CallerID)
                strLastGoodPosition = "Update the status bit for UMC_INDICATOR_BIT_USED_FOR_NET_ADJ"
                
                ' First clear the UMC_INDICATOR_BIT_USED_FOR_NET_ADJ status bit
                For lngIndex = 0 To .UMCCnt - 1
                    .UMCs(lngIndex).ClassStatusBits = .UMCs(lngIndex).ClassStatusBits And Not UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
                Next lngIndex

                ' Now update the status bit for the LC-MS Features in mPepTransformRT()
                For lngIndex = 0 To UBound(mPepTransformRT, 1)
                    With .UMCs(mPepTransformRT(lngIndex, ptcPepTransformRTColumns.ptcUMCIndex))
                        .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
                    End With
                Next lngIndex
            End With
            
        
            If UMCNetAdjDef.MSWarpOptions.SplitWarpMode <> swmDisabled And mSplitWarpIteration >= 2 Then
                ' We will append new residual values data to the arrays
                ' Need to reserve more room in the arrays
                
                lngDataCountOld = UBound(mScanMTNETvsLinearNETResidual, 2) + 1
                lngDataCountTotal = lngDataCountNew + lngDataCountOld
                
                If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmSplitOnMZ Then
                    ' Clear out the NET residual arrays since we only keep the final NET alignment
                    ReDim mScanMTNETvsLinearNETResidual(1, lngDataCountNew)
                    ReDim mScanMTNETvsCustomNETResidual(1, lngDataCountNew)
                    ReDim mScanCustomNetVsLinearNETResidual(1, lngDataCountNew)
                    lngResidualNETTargetIndexStart = 0
                    lngNETDataCountExpected = lngDataCountNew
                Else
                    ' Do not clear the NET residual arrays
                    ReDim Preserve mScanMTNETvsLinearNETResidual(1, lngDataCountTotal - 1)
                    ReDim Preserve mScanMTNETvsCustomNETResidual(1, lngDataCountTotal - 1)
                    ReDim Preserve mScanCustomNetVsLinearNETResidual(1, lngDataCountTotal - 1)
                    lngResidualNETTargetIndexStart = lngDataCountOld
                    lngNETDataCountExpected = lngDataCountTotal
                End If
                
                ReDim Preserve mScanVsMassError(1, lngDataCountTotal - 1)
                ReDim Preserve mScanVsMassErrorCorrected(1, lngDataCountTotal - 1)
                ReDim Preserve mMZVsMassError(1, lngDataCountTotal - 1)
                ReDim Preserve mMZVsMassErrorCorrected(1, lngDataCountTotal - 1)
                lngResidualMassTargetIndexStart = lngDataCountOld
                
            Else

                ' Initialize the residual value arrays
                ReDim mScanMTNETvsLinearNETResidual(1, lngDataCountNew - 1)
                ReDim mScanMTNETvsCustomNETResidual(1, lngDataCountNew - 1)
                ReDim mScanCustomNetVsLinearNETResidual(1, lngDataCountNew - 1)
                ReDim mScanVsMassError(1, lngDataCountNew - 1)
                ReDim mScanVsMassErrorCorrected(1, lngDataCountNew - 1)
                ReDim mMZVsMassError(1, lngDataCountNew - 1)
                ReDim mMZVsMassErrorCorrected(1, lngDataCountNew - 1)
            
                mLinearNETResidualMin = 0
                mLinearNETResidualMax = 0
                
                mMassResidualMin = 0
                mMassResidualMax = 0
                
                mMinMZ = 100000
                mMaxMZ = -100000
            
                lngDataCountTotal = lngDataCountNew
            
                lngResidualNETTargetIndexStart = 0
                lngNETDataCountExpected = lngDataCountNew
                
                lngResidualMassTargetIndexStart = 0
                
            End If

            '--------------------------------------------------
            ' Compute the residual NET values for the data
            '--------------------------------------------------
            strLastGoodPosition = "Compute the residual NET values for the data"
            lngResidualDataTargetIndex = lngResidualNETTargetIndexStart
            For lngIndex = 0 To lngDataCountNew - 1
            
                ' Lookup the scan number and custom NET value for lngIndex
                strLastGoodPosition = "Lookup the scan number and custom NET value for lngIndex " & Trim(lngIndex)
                If GetUMCScanMZAndNET(mMatches(lngIndex, mcMatchColumns.mcUMCIndex), lngScanNumber, dblMZ, dblCustomNET) Then
                    
                    '--------------------------------------------------
                    ' Compute residual NETs
                    '--------------------------------------------------
                    strLastGoodPosition = "Compute residual NETs (lngIndex = " & Trim(lngIndex) & ")"
                    
                    ' Use dblCustomNET to compute the NET of the matched MT tag.
                    dblMassTagNET = dblCustomNET - mMatches(lngIndex, mcMatchColumns.mcNETError)
                    
                    dblPredictedNETLinear = dblSlope * lngScanNumber + dblIntercept
                    
                    ' Subtract the NET of the MT tag from the linear model to get residual from the linear model
                    mScanMTNETvsLinearNETResidual(0, lngResidualDataTargetIndex) = lngScanNumber
                    mScanMTNETvsLinearNETResidual(1, lngResidualDataTargetIndex) = dblMassTagNET - dblPredictedNETLinear
                    
                    ' Keep track of the minimum and maximum residual NET values
                    If mScanMTNETvsLinearNETResidual(1, lngResidualDataTargetIndex) < mLinearNETResidualMin Then
                        mLinearNETResidualMin = mScanMTNETvsLinearNETResidual(1, lngResidualDataTargetIndex)
                    End If
                    If mScanMTNETvsLinearNETResidual(1, lngResidualDataTargetIndex) > mLinearNETResidualMax Then
                        mLinearNETResidualMax = mScanMTNETvsLinearNETResidual(1, lngResidualDataTargetIndex)
                    End If
                    
                    ' Subtract the NET of the MT tag from the custom model to get residual from the custom model
                    mScanMTNETvsCustomNETResidual(0, lngResidualDataTargetIndex) = lngScanNumber
                    mScanMTNETvsCustomNETResidual(1, lngResidualDataTargetIndex) = dblMassTagNET - dblCustomNET
                    
                    mScanCustomNetVsLinearNETResidual(0, lngResidualDataTargetIndex) = lngScanNumber
                    mScanCustomNetVsLinearNETResidual(1, lngResidualDataTargetIndex) = dblCustomNET - dblPredictedNETLinear
                    
                    lngResidualDataTargetIndex = lngResidualDataTargetIndex + 1
                End If
            Next lngIndex
            
            ' Round mLinearNETResidualMin and mLinearNETResidualMax to the nearest 0.001 or nearest 0.01
            strLastGoodPosition = "Round mLinearNETResidualMin and mLinearNETResidualMax"
            If mLinearNETResidualMax - mLinearNETResidualMin < 0.005 Then
                intRoundingDivisor = 10000
            ElseIf mLinearNETResidualMax - mLinearNETResidualMin < 0.05 Then
                intRoundingDivisor = 1000
            Else
                intRoundingDivisor = 100
            End If
            
            mLinearNETResidualMin = Int(mLinearNETResidualMin * intRoundingDivisor) / intRoundingDivisor
            mLinearNETResidualMax = Int(mLinearNETResidualMax * intRoundingDivisor) / intRoundingDivisor + 1 / intRoundingDivisor
            
            If lngResidualDataTargetIndex < lngNETDataCountExpected Then
                ' This is unexpected; some of the data was not valid
                Debug.Assert False
                If lngResidualDataTargetIndex > 0 Then
                    ReDim Preserve mScanMTNETvsLinearNETResidual(1, lngResidualDataTargetIndex - 1)
                    ReDim Preserve mScanMTNETvsCustomNETResidual(1, lngResidualDataTargetIndex - 1)
                    ReDim Preserve mScanCustomNetVsLinearNETResidual(1, lngResidualDataTargetIndex - 1)
                Else
                    ReDim mScanMTNETvsLinearNETResidual(1, 0)
                    ReDim mScanMTNETvsCustomNETResidual(1, 0)
                    ReDim mScanCustomNetVsLinearNETResidual(1, 0)
                End If
            End If
            
            If lngResidualDataTargetIndex > 0 Then
            
                ' Compute the mean residual value
                dblMeanNETResidual = 0
                For lngIndex = 0 To lngResidualDataTargetIndex - 1
                    ' Compute the mean NET residual (difference between the MT tag NET value and the custom NET value)
                    dblMeanNETResidual = dblMeanNETResidual + Abs(mScanMTNETvsCustomNETResidual(1, lngIndex))
                Next lngIndex
                dblMeanNETResidual = dblMeanNETResidual / lngResidualDataTargetIndex
                
                SortTransformScanResidual
            Else
                ' Assign a very poor fit mean NET residual value
                dblMeanNETResidual = 1
            End If
            
            strLastGoodPosition = "Call DisplaySlopeAndIntercept"
            DisplaySlopeAndIntercept dblSlope, dblIntercept, dblRSquared, dblMeanNETResidual
            
            ' Store the Slope and Intercept in  GelAnalysis(CallerID)
            If Not GelAnalysis(CallerID) Is Nothing Then
                With GelAnalysis(CallerID)
                    .GANET_Slope = dblSlope
                    .GANET_Intercept = dblIntercept
                    .GANET_Fit = dblMeanNETResidual
                End With
            End If
        
            ' Update the analysis history with the settings used;
            
            ' Note: The text used in this update statement should match the similar text
            '       in frmSearchForNETAdjustmentUMC->UpdateAnalysisHistory
            ' This is important since other portions of the program scrape the analysis history log
            '  to obtain some of these values
            
            strMessage = "MS Warp NET Alignment performed; " & ConstructHistoryTextForNETAlignment(strSecondsElapsed)
        
            strMessage = strMessage & "; Closest Linear NET Formula = " & ConstructNETFormula(dblSlope, dblIntercept) & _
                                      "; R-squared = " & Format(dblRSquared, "0.000") & _
                                      "; Fit (mean residual) = " & Format(dblMeanNETResidual, "0.00E-00")
            
            AddToAnalysisHistory CallerID, strMessage
        
            If Not mRecalibratingMassDuringAutoAnalysis Then
                AddToAnalysisHistory CallerID, ConstructHistoryTextForMSWarpSettings(strSecondsElapsed)
            End If
                 
            
            If blnUpdateMassValues Then
                '--------------------------------------------------
                ' Update the mass values for each data point
                '--------------------------------------------------
                strLastGoodPosition = "Call RecalibrateMassesUsingWarpedData"
                blnSuccess = RecalibrateMassesUsingWarpedData
            End If
            
            '--------------------------------------------------
            ' Compute the residual Mass value for the data
            '--------------------------------------------------
            strLastGoodPosition = "Compute the residual Mass value for the data"
            lngResidualDataTargetIndex = lngResidualMassTargetIndexStart
            For lngIndex = 0 To lngDataCountNew - 1
            
                ' Lookup the scan number, m/z, and custom NET value for lngIndex
                strLastGoodPosition = "Lookup the scan number, m/z, and custom NET value for lngIndex " & Trim(lngIndex)
                If GetUMCScanMZAndNET(mMatches(lngIndex, mcMatchColumns.mcUMCIndex), lngScanNumber, dblMZ, dblCustomNET) Then
                    '--------------------------------------------------
                    ' Compute residual mass values (vs. scan)
                    '--------------------------------------------------
                    strLastGoodPosition = "Compute residual mass values vs. scan (lngIndex = " & Trim(lngIndex) & ")"
                    
                    ' Get the Mass error between match and the UMC
                    dblMassError = mMatches(lngIndex, mcMatchColumns.mcMassError)
                    dblMassError = ValidateNotInfinity(dblMassError)
                    
                    mScanVsMassError(0, lngResidualDataTargetIndex) = lngScanNumber
                    mScanVsMassError(1, lngResidualDataTargetIndex) = dblMassError
                    
                    If blnUpdateMassValues Then
                        dblMassErrorCalibrated = dblMassError - GetPPMShift(dblMZ, CDbl(lngScanNumber))
                                            
                        mScanVsMassErrorCorrected(0, lngResidualDataTargetIndex) = lngScanNumber
                        mScanVsMassErrorCorrected(1, lngResidualDataTargetIndex) = dblMassErrorCalibrated
                    Else
                        mScanVsMassErrorCorrected(0, lngResidualDataTargetIndex) = lngScanNumber
                        mScanVsMassErrorCorrected(1, lngResidualDataTargetIndex) = dblMassError
                    End If
                    
                    
                    ' Keep track of the minimum and maximum residual mass values
                    If mScanVsMassErrorCorrected(1, lngResidualDataTargetIndex) < mMassResidualMin Then
                        mMassResidualMin = mScanVsMassErrorCorrected(1, lngResidualDataTargetIndex)
                    End If
                    If mScanVsMassErrorCorrected(1, lngResidualDataTargetIndex) > mMassResidualMax Then
                        mMassResidualMax = mScanVsMassErrorCorrected(1, lngResidualDataTargetIndex)
                    End If
                
                    '--------------------------------------------------
                    ' Compute residual mass values (vs. m/z)
                    '--------------------------------------------------
                    strLastGoodPosition = "Compute residual mass values vs. m/z (lngIndex = " & Trim(lngIndex) & ")"
                    
                    mMZVsMassError(0, lngResidualDataTargetIndex) = dblMZ
                    mMZVsMassError(1, lngResidualDataTargetIndex) = dblMassError
                    
                    If blnUpdateMassValues Then
                        mMZVsMassErrorCorrected(0, lngResidualDataTargetIndex) = dblMZ
                        mMZVsMassErrorCorrected(1, lngResidualDataTargetIndex) = dblMassErrorCalibrated
                    Else
                        mMZVsMassErrorCorrected(0, lngResidualDataTargetIndex) = dblMZ
                        mMZVsMassErrorCorrected(1, lngResidualDataTargetIndex) = dblMassError
                    End If
                End If
            
                ' Possibly update mMinMZ and mMaxMZ
                If dblMZ < mMinMZ Then
                    mMinMZ = dblMZ
                End If
                If dblMZ > mMaxMZ Then
                    mMaxMZ = dblMZ
                End If
                
                lngResidualDataTargetIndex = lngResidualDataTargetIndex + 1
            Next lngIndex
            
            ' Round mMassResidualMin and mMassResidualMax to the nearest 1, 5, or 10
            If mMassResidualMax - mMassResidualMin < 10 Then
                intRoundingDivisor = 1
            ElseIf mMassResidualMax - mMassResidualMin < 100 Then
                intRoundingDivisor = 5
            Else
                intRoundingDivisor = 10
            End If
            
            RoundValuesToNearestMultiple mMassResidualMin, mMassResidualMax, intRoundingDivisor
            
            ' Round mMinMZ and mMaxMZ to the nearest 50
            intRoundingDivisor = 50
            sngMinMZInitial = mMinMZ
            RoundValuesToNearestMultiple mMinMZ, mMaxMZ, intRoundingDivisor
        
            If Abs(sngMinMZInitial - mMinMZ) < 25 Then
                mMinMZ = mMinMZ - 50
            End If
        End If
    End If
    
    mLocalGelUpdated = True
    Exit Sub
    
CopyLocalDataToGelErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error copying results into gel data structures: " & Err.Description & vbCrLf & "Last good position: " & strLastGoodPosition & "(lngIndex = " & Trim(lngIndex) & ")", vbExclamation + vbOKOnly, "Error"
    Else
        LogErrors Err.Number, "frmMSAlign.CopyLocalDataToGel"
    End If
    mLocalGelUpdated = True
    
End Sub

Public Function CopyResidualNETValuesToClipboardOrFile(Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    ' Returns 0 if success; error number if error
    
    Dim lngDataCount As Long
    Dim lngCustomNetVsLinearNETResidualDataCount As Long
    Dim strData() As String     ' Ranges from 0 to lngDataCount since row 0 is the header row
    Dim sngSortKeys() As Single ' Parallel with strData(); used to sort strData after populating it
    
    Dim lngIndex As Long
    
On Error Resume Next
    lngDataCount = UBound(mScanMTNETvsLinearNETResidual, 2) + 1
    lngCustomNetVsLinearNETResidualDataCount = UBound(mScanCustomNetVsLinearNETResidual, 2) + 1
    
    If lngDataCount <= 1 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        CopyResidualNETValuesToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo CopyResidualNETValuesToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    ' Header row is strData(0)
    ' Data is from strData(1) to strData(lngDataCount)
    ReDim strData(0 To lngDataCount)
    ReDim sngSortKeys(0 To lngDataCount)
    
    ' Fill strData()
    ' Define the header row
    strData(0) = "Scan Number" & vbTab & "Linear NET Residual" & vbTab & "Warped NET Residual" & vbTab & "NET Transform Function"
    sngSortKeys(0) = -1E+30
    
    For lngIndex = 1 To lngDataCount
        sngSortKeys(lngIndex) = mScanMTNETvsLinearNETResidual(0, lngIndex - 1)
        strData(lngIndex) = Round(mScanMTNETvsLinearNETResidual(0, lngIndex - 1), 1) & vbTab & _
                            Round(mScanMTNETvsLinearNETResidual(1, lngIndex - 1), 6) & vbTab & _
                            Round(mScanMTNETvsCustomNETResidual(1, lngIndex - 1), 6) & vbTab
        
        If lngIndex < lngCustomNetVsLinearNETResidualDataCount Then
            strData(lngIndex) = strData(lngIndex) & Round(mScanCustomNetVsLinearNETResidual(1, lngIndex - 1), 6)
        End If
    Next lngIndex
      
    CopyResidualsFinalize strData, sngSortKeys, strFilePath
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    CopyResidualNETValuesToClipboardOrFile = 0
    Exit Function

CopyResidualNETValuesToClipboardOrFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting NET residuals: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    CopyResidualNETValuesToClipboardOrFile = Err.Number
    
End Function

Public Function CopyResidualMassVsScanValuesToClipboardOrFile(Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    ' Returns 0 if success; error number if error
    
    Dim lngDataCount As Long
    Dim strData() As String     ' Ranges from 0 to lngDataCount since row 0 is the header row
    Dim sngSortKeys() As Single ' Parallel with strData(); used to sort strData after populating it
    
    Dim lngIndex As Long
    
On Error Resume Next
    lngDataCount = UBound(mScanVsMassError, 2)
    If lngDataCount = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        CopyResidualMassVsScanValuesToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo CopyResidualMassVsScanValuesToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    ' Header row is strData(0)
    ' Data is from strData(1) to strData(lngDataCount)
    ReDim strData(0 To lngDataCount)
    ReDim sngSortKeys(0 To lngDataCount)
    
    ' Fill strData()
    ' Define the header row
    strData(0) = "Scan Number" & vbTab & "Initial Mass Residual" & vbTab & "Warped Mass Residual"
    sngSortKeys(0) = -1E+30
    
    For lngIndex = 1 To lngDataCount
        sngSortKeys(lngIndex) = mScanVsMassError(0, lngIndex - 1)
        strData(lngIndex) = Round(mScanVsMassError(0, lngIndex - 1), 1) & vbTab & _
                            Round(mScanVsMassError(1, lngIndex - 1), 4) & vbTab & _
                            Round(mScanVsMassErrorCorrected(1, lngIndex - 1), 4)
    Next lngIndex
      
    CopyResidualsFinalize strData, sngSortKeys, strFilePath
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    CopyResidualMassVsScanValuesToClipboardOrFile = 0
    Exit Function

CopyResidualMassVsScanValuesToClipboardOrFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting NET residuals: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    CopyResidualMassVsScanValuesToClipboardOrFile = Err.Number
    
End Function

Public Function CopyResidualMassVsMZValuesToClipboardOrFile(Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    ' Returns 0 if success; error number if error
    
    Dim lngDataCount As Long
    Dim strData() As String     ' Ranges from 0 to lngDataCount since row 0 is the header row
    Dim sngSortKeys() As Single ' Parallel with strData(); used to sort strData after populating it
    
    Dim lngIndex As Long
    
On Error Resume Next
    lngDataCount = UBound(mMZVsMassError, 2)
    If lngDataCount = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        CopyResidualMassVsMZValuesToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo CopyResidualMassVsScanValuesToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    ' Header row is strData(0)
    ' Data is from strData(1) to strData(lngDataCount)
    ReDim strData(0 To lngDataCount)
    ReDim sngSortKeys(0 To lngDataCount)
    
    ' Fill strData()
    ' Define the header row
    strData(0) = "m/z" & vbTab & "Initial Mass Residual" & vbTab & "Warped Mass Residual"
    sngSortKeys(0) = -1E+30
    
    For lngIndex = 1 To lngDataCount
        sngSortKeys(lngIndex) = mMZVsMassError(0, lngIndex - 1)
        strData(lngIndex) = Round(mMZVsMassError(0, lngIndex - 1), 4) & vbTab & _
                            Round(mMZVsMassError(1, lngIndex - 1), 4) & vbTab & _
                            Round(mMZVsMassErrorCorrected(1, lngIndex - 1), 4)
    Next lngIndex
      
    CopyResidualsFinalize strData, sngSortKeys, strFilePath
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    CopyResidualMassVsMZValuesToClipboardOrFile = 0
    Exit Function

CopyResidualMassVsScanValuesToClipboardOrFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting NET residuals: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    CopyResidualMassVsMZValuesToClipboardOrFile = Err.Number
    
End Function

Private Sub CopyResidualsFinalize(ByRef strData() As String, ByRef sngSortKeys() As Single, ByVal strFilePath As String)
  Dim fso As FileSystemObject
  Dim tsOutfile As TextStream
  
  Dim lngIndex As Long
  Dim strTextToCopy As String
  
  ' Sort sngSortKeys() and sort strData() parallel to it
  ShellSortSingleWithParallelString sngSortKeys, strData, 0, UBound(sngSortKeys)
  
  If Len(strFilePath) > 0 Then
        Set fso = New FileSystemObject
        Set tsOutfile = fso.CreateTextFile(strFilePath, True)
        
        For lngIndex = 0 To UBound(strData)
           tsOutfile.WriteLine strData(lngIndex)
        Next lngIndex
        
        tsOutfile.Close
    Else
        strTextToCopy = FlattenStringArray(strData(), UBound(strData) + 1, vbCrLf, False)
        Clipboard.Clear
        Clipboard.SetText strTextToCopy, vbCFText
    End If
End Sub

Private Sub DisplayResidualPlotAxisRangesCurrent()
    Select Case mMostRecentPlotViewMode
    Case pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot
        DisplayResidualPlotAxisRanges ctlMassVsScanResidual, 0, 1
    Case pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot
        DisplayResidualPlotAxisRanges ctlMassVsMZResidual, 0, 1
    Case Else
        ' Includes pvmLinearFitNETResidualsPlot & pvmWarpedFitNETResidualsPlot
        DisplayResidualPlotAxisRanges ctlNETResidual, 0, 3
    End Select
End Sub

Private Sub DisplayResidualPlotAxisRanges(ByRef objPlot As CWGraph, ByVal intDigitsAfterDecimalForX As Integer, ByVal intDigitsAfterDecimalForY As Integer)
    If mUpdatingPlotRanges Then Exit Sub
    
On Error Resume Next

    With objPlot
        txtResidualPlotMinX.Text = Round(.Plots(1).XAxis.Minimum, intDigitsAfterDecimalForX)
        txtResidualPlotMaxX.Text = Round(.Plots(1).XAxis.Maximum, intDigitsAfterDecimalForX)
    
        txtResidualPlotMinY.Text = Round(.Plots(1).YAxis.Minimum, intDigitsAfterDecimalForY)
        txtResidualPlotMaxY.Text = Round(.Plots(1).YAxis.Maximum, intDigitsAfterDecimalForY)
    End With

End Sub

Private Sub DisplaySlopeAndIntercept(ByVal dblSlope As Double, ByVal dblIntercept As Double, ByVal dblRSquared As Double, ByVal dblMeanNETResidual As Double)
    txtSlope.Text = Format(dblSlope, "0.000E-00")
    txtIntercept.Text = Format(dblIntercept, "0.000")

    txtRSquared.Text = Format(dblRSquared, "0.000")
    txtFit = Format(dblMeanNETResidual, "0.00E-00")
End Sub

Private Sub EnableDisableControls(blnEnabled As Boolean)
    mControlsEnabled = blnEnabled
    
    If blnEnabled Then
        Me.MousePointer = vbDefault
    Else
        mAbortRequested = False
        Me.MousePointer = vbHourglass
    End If
    
    fraNETWarpOptions.Enabled = blnEnabled
    fraMassWarpOptions.Enabled = blnEnabled
    
    fraNETTolerances.Enabled = blnEnabled
    fraBinningOptions.Enabled = blnEnabled
    fraMassCalibType.Enabled = blnEnabled

    cmdMassCalibrationRevert.Enabled = blnEnabled
    cboStepsToPerform.Enabled = blnEnabled
    cmdSetDefaults.Enabled = blnEnabled
    
    cmdWarpAlign.Visible = blnEnabled
    cmdAbort.Visible = Not blnEnabled
    cmdAbort.Caption = "Abort"
    
    mnuEdit.Enabled = blnEnabled
    mnuPMTTags.Enabled = blnEnabled
    
    Select Case cboStepsToPerform.ListIndex
    Case MassMatchStepsToPerformConstants.mmsWarpTimeAndMass
        fraMassWarpOptions.Enabled = blnEnabled
    Case Else
        fraMassWarpOptions.Enabled = False
    End Select
    
    DoEvents
End Sub

Private Function EstimateLinearNET(ByVal intTopAbuPct As Integer, ByRef dblSlope As Double, ByRef dblIntercept As Double, ByRef dblRSquared As Double, byreflngFeatureCountUsed As Long) As Boolean
    ' Order the features by descending abundance
    ' Select the top intTopAbuPct percent of the features and use them to estimate the
    '  linear fit between the Custom NETs and the scan numbers
    
    Dim udtUMCNetAdjDefSaved As NetAdjDefinition
    Dim lngNETAdjustmentMinIDCountSaved As Long
    Dim blnRequireDispersedUMCSelectionSaved  As Boolean
    Dim lngMinIDCountThreshold As Long
    Dim lngMinIDCountThresholdAlt As Long
    
    Dim lngIndex As Long
    Dim lngDataCount As Long
    Dim lngValidDataCount As Long
    Dim lngScanNumber As Long
    
    Dim dblCustomNET As Double
    Dim dblMZ As Double
    
    Dim UseUMC() As Boolean
    Dim lngUMCCntAddedSinceLowSegmentCount As Long
    Dim lngUMCSegmentCntWithLowUMCCnt As Long
        
    Dim SumY As Double
    Dim SumX As Double
    Dim SumXX As Double
    Dim SumXY As Double
    Dim SumYY As Double
      
    Dim dblStatTemp As Double
    
    Dim blnSuccess As Boolean
    blnSuccess = False
    
    ' Record the number of matches
    If MatchesAreValid Then
        lngDataCount = UBound(mMatches, 1) + 1
    Else
        lngDataCount = 0
    End If
    
    '--------------------------------------------------
    ' Validate/override some of the NET Adj settings
    '--------------------------------------------------
    
On Error GoTo EstimateLinearNETErrorHandler

    udtUMCNetAdjDefSaved = UMCNetAdjDef
    With glbPreferencesExpanded
        lngNETAdjustmentMinIDCountSaved = .AutoAnalysisOptions.NETAdjustmentMinIDCount
        blnRequireDispersedUMCSelectionSaved = .NetAdjustmentUMCDistributionOptions.RequireDispersedUMCSelection
    End With
    
    With UMCNetAdjDef
        .TopAbuPct = intTopAbuPct
        If .MinUMCCount < 3 Then .MinUMCCount = 3
        If .MinScanRange < 3 Then .MinScanRange = 3
        .PeakCSSelection(7) = True                      ' Always use all charge states
    End With
    
    With glbPreferencesExpanded
        lngMinIDCountThreshold = GelUMC(CallerID).UMCCnt * 0.05
        lngMinIDCountThresholdAlt = lngDataCount * 0.1
        If lngMinIDCountThresholdAlt >= 10 And lngMinIDCountThresholdAlt < lngMinIDCountThreshold Then
            lngMinIDCountThreshold = lngMinIDCountThresholdAlt
        End If
           
        If lngMinIDCountThreshold < 4 Then lngMinIDCountThreshold = 4
        If lngMinIDCountThreshold > 100 Then lngMinIDCountThreshold = 100
        
        .AutoAnalysisOptions.NETAdjustmentMinIDCount = lngMinIDCountThreshold
        
        .NetAdjustmentUMCDistributionOptions.RequireDispersedUMCSelection = False
    End With
           

    '--------------------------------------------------
    ' Select the most abundant LC-MS Features that pass the UMC selection criteria defined in UMCNetAdjDef
    '--------------------------------------------------
    
    ' Initially set all LC-MS Features for use
    ReDim UseUMC(GelUMC(CallerID).UMCCnt - 1)
    For lngIndex = 0 To UBound(UseUMC)
        UseUMC(lngIndex) = True
    Next lngIndex
    
    LinearNETAlignmentSelectUMCToUse CallerID, UseUMC(), lngUMCCntAddedSinceLowSegmentCount, lngUMCSegmentCntWithLowUMCCnt
    
    
    '--------------------------------------------------
    ' Step through the features that have valid matches
    ' Only use those LC-MS Features with UseUMC() = True
    ' Use the scan and CustomNET for each feature to compute
    ' a least squares line of scan vs. CustomNET, along with
    ' the R-Squared value
    '--------------------------------------------------
    
    SumY = 0
    SumX = 0
    SumXY = 0
    SumXX = 0

    lngValidDataCount = 0
    For lngIndex = 0 To lngDataCount - 1
        If UseUMC(mMatches(lngIndex, mcMatchColumns.mcUMCIndex)) Then
            If GetUMCScanMZAndNET(mMatches(lngIndex, mcMatchColumns.mcUMCIndex), lngScanNumber, dblMZ, dblCustomNET) Then
                SumX = SumX + lngScanNumber
                SumY = SumY + dblCustomNET
                SumXX = SumXX + CDbl(lngScanNumber) ^ 2
                SumXY = SumXY + lngScanNumber * dblCustomNET
                SumYY = SumYY + dblCustomNET * dblCustomNET
                lngValidDataCount = lngValidDataCount + 1
            End If
        End If
    Next lngIndex
    
    If lngValidDataCount > 1 Then
        dblSlope = (lngValidDataCount * SumXY - SumX * SumY) / (lngValidDataCount * SumXX - SumX * SumX)
        dblIntercept = (SumY - dblSlope * SumX) / lngValidDataCount
    
        dblStatTemp = (lngValidDataCount * SumXY - SumX * SumY) _
            / Sqr((lngValidDataCount * SumXX - (SumX * SumX)) _
            * (lngValidDataCount * SumYY - (SumY * SumY)))
        dblRSquared = dblStatTemp * dblStatTemp
        
        If lngValidDataCount >= glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount Then
            blnSuccess = True
        Else
            blnSuccess = False
        End If
    Else
        blnSuccess = False
    End If
    
EstimateLinearNETContinue:

    ' -------------------------------------------------
    ' Restore the saved settings
    ' -------------------------------------------------
    UMCNetAdjDef = udtUMCNetAdjDefSaved
    With glbPreferencesExpanded
        .AutoAnalysisOptions.NETAdjustmentMinIDCount = lngNETAdjustmentMinIDCountSaved
        .NetAdjustmentUMCDistributionOptions.RequireDispersedUMCSelection = blnRequireDispersedUMCSelectionSaved
    End With

    EstimateLinearNET = blnSuccess
    Exit Function

EstimateLinearNETErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error estimating the Linear NET value: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.EstimateLinearNET"
    End If
    
    Resume EstimateLinearNETContinue
    
End Function

Private Sub FilterAndAlignFeatures(ByVal intIteration As Integer)
    Const NET_MIN As Double = 0
    Const NET_MAX As Double = 1
    
    ' On the first call to this function, intIteration should be 1
    ' On the second call, it should be 2
        
    Dim blnStartAlignment As Boolean
    
    blnStartAlignment = False
    
    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Then
        If intIteration <= 1 Then
            If mLocalFeaturesAreFiltered Then
                ' Cached local features are filtered; need to update the cached data
                PopulateLocalFeaturesArray True
            End If
            
            blnStartAlignment = True
        Else
           ' No more iterations are required
        End If
        
    ElseIf UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmSplitOnMZ Then
        If intIteration = 0 Then intIteration = 1
        If intIteration <= 2 Then
            mSplitWarpIteration = intIteration
            
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled And _
               mSplitWarpIteration > 1 And _
               chkSplitWarpPauseBetweenIterations.Value = vbChecked Then
               
                ' User has requested that we pause between each iteration (normally there are just 2 iterations)
                mPauseSplitWarpProcessing = True
                cmdSplitWarpResume.Visible = True
                cmdAbort.Caption = "Paused (Abort)"
                
                Do
                    Sleep 50
                    DoEvents
                    If mAlignmentFinalizedOrAborted Then
                        cmdSplitWarpResume.Visible = False
                        Exit Sub
                    End If
                Loop While mPauseSplitWarpProcessing
                
                cmdSplitWarpResume.Visible = False
                cmdAbort.Caption = "Abort"
                DoEvents
                
            End If
            
            ' Filter and cache the features
            PopulateLocalFeaturesArray True

            blnStartAlignment = True
        Else
            ' No more iterations are required
        End If
    End If
    
    If blnStartAlignment Then
        tmrAlignment.Enabled = True
        
        With UMCNetAdjDef
            ' Make sure this is true
            .UseRobustNETAdjustment = True
            
            ' Reset the status to pscUninitialized (necessary in case the status is pscError or pscAborted)
            Call mMassMatchObject.ResetStatus
            
            ' Set the MassMatch Options
            Call mMassMatchObject.SetNetOptions(CLng(.MSWarpOptions.NumberOfSections), _
                                             CInt(.MSWarpOptions.ContractionFactor), _
                                             CInt(.MSWarpOptions.MaxDistortion), _
                                             CDbl(.MSWarpOptions.NETTol), _
                                             CDbl(NET_MIN), _
                                             CDbl(NET_MAX), _
                                             CLng(.MSWarpOptions.MatchPromiscuity))
                                             
            Call mMassMatchObject.SetMassOptions(CDbl(.MWTol), _
                                             CLng(.MSWarpOptions.MassNumMassDeltaBins), _
                                             CDbl(.MSWarpOptions.MassWindowPPM), _
                                             CLng(.MSWarpOptions.MassMaxJump), _
                                             CLng(.MSWarpOptions.MassNumXSlices), _
                                             CDbl(UMCNetAdjDef.MSWarpOptions.MassZScoreTolerance), _
                                             CLng(.MSWarpOptions.MassUseLSQ))
                                             
            Call mMassMatchObject.SetMassLSQOptions(.MSWarpOptions.MassLSQNumKnots, UMCNetAdjDef.MSWarpOptions.MassLSQOutlierZScore)
            
            mCalibrationType = .MSWarpOptions.MassCalibrationType
            mSplineOrder = .MSWarpOptions.MassSplineOrder
    
            Call mMassMatchObject.SetRegressionOrder(mSplineOrder)
            Call mMassMatchObject.SetRecalibrationType(mCalibrationType)
            
            If .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass Then
                Call mMassMatchObject.SetAlignmentType(1)
            Else
                .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTime
                Call mMassMatchObject.SetAlignmentType(0)
            End If
        End With
        
        ' Perform the alignment
        ' Note: This process occurs asynchronously; when complete then
        '       QueryMassMatchProgress will call FinalizeAlignment
        Call mMassMatchObject.MS2MSMSDBAlignPeptidesThreaded(mLocalFeatures, mLocalPMTsFiltered)
    End If
    
End Sub

Private Sub FinalizeAlignment(strSecondsElapsed As String)
    Dim strMessage As String
    Dim blnUpdateMassValues As Boolean
    Dim blnAutoZoomOut As Boolean

    Dim intNewIteration As Integer
    
    Dim varMassErrorHistogramOld As Variant
    Dim varNetErrorHistogramOld As Variant
    Dim blnMergeCachedHistogramData As Boolean
    
On Error GoTo FinalizeAlignmentErrorHandler

    ' Stop the timer
    tmrAlignment.Enabled = False

    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode <> swmDisabled Then
        If mSplitWarpIteration >= 2 Then
            ' Cache some of the error histogram values so that we can add them to the new values returned by .GetResults
            
            varMassErrorHistogramOld = mMassErrorHistogram
            varNetErrorHistogramOld = mNetErrorHistogram
            
            blnMergeCachedHistogramData = True
        End If
    End If
    
    ' Obtain the alignment results
    With glbPreferencesExpanded.ErrorPlottingOptions
        mMassMatchObject.GetResults mMatchScores, mAlignmentFunc, mMatches, mPepTransformRT, mTransformRT, mMassErrorHistogram, mNetErrorHistogram, CDbl(.MassBinSizePPM), CDbl(.GANETBinSize)
    End With
    
    If blnMergeCachedHistogramData Then
        ' Merge the cached histogram data with the new histograms
        
        MergeHistogramData mMassErrorHistogram, varMassErrorHistogramOld, glbPreferencesExpanded.ErrorPlottingOptions.MassBinSizePPM
        
        MergeHistogramData mNetErrorHistogram, varNetErrorHistogramOld, glbPreferencesExpanded.ErrorPlottingOptions.GANETBinSize
    End If
    
    If UMCNetAdjDef.RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass Then
        blnUpdateMassValues = True
    Else
        blnUpdateMassValues = False
    End If
    
    ' Find range of features and MT tags
    mMassMatchObject.GetBounds mMinNet, mMaxNet, mMinScan, mMaxScan
    
    If MatchesAreValid Then
        txtNumMatched.Text = LongToStringWithCommas(UBound(mMatches, 1) + 1)
    Else
        txtNumMatched.Text = "0"
    End If
    
    UpdateStatus "Processing: Updating plots (" & strSecondsElapsed & " seconds elapsed)"
    
    ' Update the UMC masses and NETs with the aligned values
    ' This needs to be done before calling PlotSurfaceData in order to populate txtSlope and txtIntercept
    CopyLocalDataToGel strSecondsElapsed, blnUpdateMassValues
    
    ' Update the Flat Surface
    PlotSurfaceData
    
    ' Update the NI Surface
    ctlCWGraphNI.Plot3DSimpleSurface mMatchScores
    
    ' Update the mass and NET error histograms
    If UBound(mMassErrorHistogram) >= 0 Then
        graphMassErrors.PlotXY mMassErrorHistogram, False
    Else
        graphMassErrors.ClearData
    End If
    
    If UBound(mNetErrorHistogram) >= 0 Then
        graphNetErrors.PlotXY mNetErrorHistogram, False
    Else
        graphNetErrors.ClearData
    End If
    
    blnAutoZoomOut = cChkBox(chkAutoZoomOut)
    Select Case mMostRecentPlotViewMode
    Case pvmLinearFitNETResidualsPlot, pvmWarpedFitNETResidualsPlot
        UpdateNETResidualsPlot blnAutoZoomOut
    Case pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot
        UpdateMassResidualsScanPlot blnAutoZoomOut
    Case pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot
        UpdateMassResidualsMZPlot blnAutoZoomOut
    Case Else
        If blnAutoZoomOut Then
            ZoomOutResidualsPlot True
        End If
    End Select
    
    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode <> swmDisabled Then
        If mSplitWarpIteration < 2 Then
            intNewIteration = mSplitWarpIteration + 1
            FilterAndAlignFeatures intNewIteration
            Exit Sub
        End If
    End If
    
    strMessage = "Processing: 100% complete (" & strSecondsElapsed & " seconds elapsed)"
    If blnUpdateMassValues Then
        With GelSearchDef(CallerID).MassCalibrationInfo
            If .AdjustmentHistoryCount >= 1 Then
                strMessage = strMessage & "; applied " & Round(.AdjustmentHistory(.AdjustmentHistoryCount - 1), 4) & " ppm average mass shift"
            End If
        End With
    End If
    
    UpdateStatus strMessage
    mAlignmentFinalizedOrAborted = True

    Exit Sub
    
FinalizeAlignmentErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error finishing up the alignment: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.FinalizeAlignment"
    End If
    
    mAlignmentFinalizedOrAborted = True

End Sub

Public Function GetNETsFromScans(ByRef lngScanNumbers() As Long, ByRef dblScanNETs() As Double) As Boolean

    Dim dblFinalSectionNETStart As Double
    Dim dblFinalSectionScanStart As Double
    Dim dblFinalSectionNETEnd As Double
    Dim dblFinalSectionScanEnd As Double
    
    Dim lngScanNumber As Long
    Dim lngScanIndex As Long
    
    Dim lngIndexMax As Long
    Dim lngIndexTransformedMax As Long
    
    Dim lngTimeSliceCount As Long
    
    Dim blnSuccess As Boolean
    
On Error GoTo GetNETsFromScansErrorHandler

    lngIndexMax = UBound(lngScanNumbers)
    lngIndexTransformedMax = UBound(mTransformRT, 1)
    
    ReDim dblScanNETs(lngIndexMax)
    
    lngTimeSliceCount = UBound(mAlignmentFunc, 1)
    
    If lngTimeSliceCount > 0 Then
        dblFinalSectionNETStart = mTransformRT(lngTimeSliceCount - 1, trcTransformRTColumns.trcNET)
        dblFinalSectionNETEnd = mTransformRT(lngTimeSliceCount, trcTransformRTColumns.trcNET)
        dblFinalSectionScanStart = mTransformRT(lngTimeSliceCount - 1, trcTransformRTColumns.trcScanNum)
        dblFinalSectionScanEnd = mTransformRT(lngTimeSliceCount, trcTransformRTColumns.trcScanNum)
            
        For lngScanIndex = 0 To lngIndexMax
            lngScanNumber = lngScanNumbers(lngScanIndex)
            If lngScanNumber > lngIndexTransformedMax Then
                dblScanNETs(lngScanIndex) = dblFinalSectionNETEnd + (dblFinalSectionNETEnd - dblFinalSectionNETStart) * (lngScanNumber - dblFinalSectionScanEnd) / (dblFinalSectionScanEnd - dblFinalSectionScanStart)
            Else
                dblScanNETs(lngScanIndex) = mTransformRT(lngScanNumber, trcTransformRTColumns.trcNET)
            End If
        Next lngScanIndex
        
        blnSuccess = True
    Else
        blnSuccess = False
    End If

    GetNETsFromScans = blnSuccess
    Exit Function

GetNETsFromScansErrorHandler:
    Debug.Assert False
    GetNETsFromScans = False
    
End Function

Private Function GetPPMShift(ByVal MZ As Double, ByVal Scan As Long) As Double
    
    Dim dblPPMShift As Double
    Dim intIndex As Integer
    
On Error GoTo GetPPMShiftErrorHandler
    Call mMassMatchObject.GetPPMShift(MZ, Scan, dblPPMShift)
    
    ' Assure that dblPPMShift is not equal to -1.#IND
    GetPPMShift = ValidateNotInfinity(dblPPMShift)
    
    Exit Function
    
GetPPMShiftErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.GetPPMShift"
    GetPPMShift = 0
    
End Function

Private Function GetUMCScanMZAndNET(ByVal lngUMCIndex As Long, ByRef lngScanNumber As Long, ByRef dblMZ As Double, ByRef dblClassNET As Double) As Boolean

On Error GoTo GetUMCScanMZAndNETErrorHandler

    With GelUMC(CallerID).UMCs(lngUMCIndex)
        Select Case .ClassRepType
        Case gldtCS
            lngScanNumber = GelData(CallerID).CSData(.ClassRepInd).ScanNumber
            dblMZ = GelData(CallerID).CSData(.ClassRepInd).MZ
            dblClassNET = ScanToGANET(CallerID, lngScanNumber)
            GetUMCScanMZAndNET = True
        Case gldtIS
            lngScanNumber = GelData(CallerID).IsoData(.ClassRepInd).ScanNumber
            dblMZ = GelData(CallerID).IsoData(.ClassRepInd).MZ
            dblClassNET = ScanToGANET(CallerID, lngScanNumber)
            GetUMCScanMZAndNET = True
        Case Else
            ' This shouldn't happen
            GetUMCScanMZAndNET = False
        End Select
    End With

    Exit Function
    
GetUMCScanMZAndNETErrorHandler:
    Debug.Assert False
    GetUMCScanMZAndNET = False
    
End Function

Private Sub InitializeCWGraph()

On Error GoTo InitializeCWGraphErrorHandler

    With ctlCWGraphNI
        .PlotTemplate.XAxis.Labels.Font.Size = 8
        .PlotTemplate.YAxis.Labels.Font.Size = 8
        .PlotTemplate.ZAxis.Labels.Font.Size = 8
        
        .PlotTemplate.XAxis.Labels.Normal = True
        .PlotTemplate.XAxis.Labels.Opposite = False
        .PlotTemplate.YAxis.Labels.Normal = True
        .PlotTemplate.YAxis.Labels.Opposite = False
        .PlotTemplate.ZAxis.Labels.Normal = True
        .PlotTemplate.ZAxis.Labels.Opposite = False
        
        .GridXY = True
        .GridXZ = True
        .GridYZ = False
    
'        .Lighting = False
'        .ProjectionStyle = cwOrthographic
    End With
    
    Exit Sub

InitializeCWGraphErrorHandler:

    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error initializing the CWGraph: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.InitializeCWGraph"
    End If
    
End Sub

Private Sub InitializeForm()

    Dim strLastGoodLocation As String
    
On Error GoTo FormLoadErrorHandler

    strLastGoodLocation = "Start"
    
    mControlsEnabled = False
    mLoading = True
    mNETLineSlopeSaved = 0
    mNETLineInterceptSaved = 0
    
    strLastGoodLocation = "Redim Arrays"
    
    ReDim mScanMTNETvsLinearNETResidual(1, 0)
    ReDim mScanMTNETvsCustomNETResidual(1, 0)
    ReDim mScanCustomNetVsLinearNETResidual(1, 0)
            
    ReDim mScanVsMassError(1, 0)
    ReDim mScanVsMassErrorCorrected(1, 0)
    
    ReDim mMZVsMassError(1, 0)
    ReDim mMZVsMassErrorCorrected(1, 0)
    
    mMinMZ = 0
    mMaxMZ = 3000
    mResidualPlotPointSize = 2
    mResidualPlotTransformationFnLineSize = 2
    
    strLastGoodLocation = "InitializeCWGraph"
    InitializeCWGraph
    ctlCWGraphNI.Visible = False
    
    strLastGoodLocation = "PopulateComboBoxes"
    PopulateComboBoxes
    
    strLastGoodLocation = "PositionControls"
    PositionControls

    cmdAbort.Visible = False
    
    strLastGoodLocation = "ClearLocalFeaturesArray"
    ClearLocalFeaturesArray
    ClearLocalPMTsArray
    
    strLastGoodLocation = "Unload frmTracker"
    If IsWinLoaded(TrackerCaption) Then Unload frmTracker
    DoEvents
    
    strLastGoodLocation = "Call ctlFlatSurface.ShowLine"
    ctlFlatSurface.ShowLine = cChkBox(chkShowTransform)
    UpdatePlotViewModeToFlatView
    
    strLastGoodLocation = "SetDefaultOptions"
    SetDefaultOptions
    tbsOptions.TabIndex = 3
    
    mLoading = False
    mControlsEnabled = True
    
    Exit Sub

FormLoadErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.InitializeForm"
    Resume Next
    
End Sub

Public Sub InitializeSearch()
    '------------------------------------------------------------------------------------
    ' Load MT tag database data if neccessary
    ' If CallerID is associated with MT tag database load that DB if not already loaded
    ' If CallerID is not associated with MT tag database load legacy database
    '------------------------------------------------------------------------------------
    
    Dim eResponse As VbMsgBoxResult
    
On Error GoTo InitializeSearchErrorHandler

    Me.MousePointer = vbHourglass

    If CallerID >= 1 And CallerID <= UBound(GelUMC) Then
        ' Update lblUMCMassMode to reflect the mass mode used to identify the LC-MS Features
        Select Case GelUMC(CallerID).def.ClassMW
        Case UMCClassMassConstants.UMCMassAvg
            lblUMCMassMode = "LC-MS Feature Mass = Average of the masses of the LC-MS Feature members"
        Case UMCClassMassConstants.UMCMassRep
            lblUMCMassMode = "LC-MS Feature Mass = Mass of the LC-MS Feature Class Representative"
        Case UMCClassMassConstants.UMCMassMed
            lblUMCMassMode = "LC-MS Feature Mass = Median of the masses of the LC-MS Feature members"
        Case UMCMassAvgTopX
            lblUMCMassMode = "LC-MS Feature Mass = Average of top X members of the LC-MS Feature"
        Case UMCMassMedTopX
            lblUMCMassMode = "LC-MS Feature Mass = Median of top X members of the LC-MS Feature"
        Case Else
            lblUMCMassMode = "LC-MS Feature Mass = ?? Unable to determine; is it a new mass mode?"
        End Select
        
        mSplitWarpIteration = 0
        PopulateLocalFeaturesArray False
    
        UMCNetAdjDef = GelUMCNETAdjDef(CallerID)
    
        Me.Caption = "LCMSWarp (" & CompactPathString(GelStatus(CallerID).GelFilePathFull, 80) & ")"
        
    Else
        lblUMCMassMode = "CallerID is not defined; LC-MS Feature Mass mode is unknown"
        ClearLocalFeaturesArray
    End If
    
    If UMCNetAdjDef.MSWarpOptions.NumberOfSections = 0 Or UMCNetAdjDef.MSWarpOptions.MaxDistortion = 0 Or UMCNetAdjDef.MSWarpOptions.ContractionFactor = 0 Then
        SetDefaultUMCNETAdjDef UMCNetAdjDef
    End If
        
    UpdateControlValues False
    UpdateMassCalibrationStats
    
    ClearLocalPMTsArray
    If GelAnalysis(CallerID) Is Nothing Then
        If AMTCnt > 0 Then    'something is loaded
            If (Len(CurrMTDatabase) > 0 Or Len(CurrLegacyMTDatabase) > 0) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                ' PMT data; we dont know is it appropriate; warn user
                WarnUserUnknownMassTags CallerID
            End If
            PopulateLocalPMTsArray
            UpdateStatus "Found MT tags in memory though current display is not associated with a specific MT tag database"
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
                LoadPMTsFromLegacyDB
            Else
                WarnUserNotConnectedToDB CallerID, True
                UpdateStatus "No MT tags loaded"
            End If
        End If
    Else
        ' Force loading of MT tags
        Call LoadPMTsFromDB(False, False)
        
        ' Display the slope and intercept
        With GelAnalysis(CallerID)
            DisplaySlopeAndIntercept .GANET_Slope, .GANET_Intercept, 0, .GANET_Fit
        End With
    End If
    
    EnableDisableControls True
    
    Me.MousePointer = vbDefault
    UpdatePlotViewModeToFlatView
    
    Exit Sub

InitializeSearchErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.InitializeSearch"
    Resume Next

End Sub

Private Sub LoadPMTsFromDB(blnForceReload As Boolean, blnLoadMTStats As Boolean)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean
    Dim strMessage As String
    
    EnableDisableControls False
    
    UpdateStatus "Confirming that MT tags are present in memory"
    
    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, blnLoadMTStats, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
        PopulateLocalPMTsArray
        
        If blnAMTsWereLoaded Then
            strMessage = "MT tags loaded from DB"
        Else
            strMessage = "Ready"
        End If
    Else
        If blnDBConnectionError Then
            strMessage = "Error loading MT tags: database connection error."
        Else
            strMessage = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
        End If
    End If
    
    UpdateStatus strMessage
    EnableDisableControls True
    
End Sub

Private Sub LoadPMTsFromFile()
    Dim OldPath As String
    Dim i As Long
    Dim j As Integer
    Dim FileNum As Integer
    
    Dim strLineIn As String
    Dim strSplitLine() As String
    
    
    Dim FileName As String
    Dim val_read As Double
    Dim Index As Long
    Dim num_members As Long
    
    Dim tempDimCount As Long
    Dim dblTempPMTs() As Double

On Error GoTo LoadPMTsFromFileErrorHandler

    If mLoading Then Exit Sub
    
    FileName = SelectFile(Me.hwnd, "Load MT tags File ...", , False, "*.csv", "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*", 1)
    If Len(FileName) = 0 Then
        Exit Sub
    End If

    tempDimCount = 1024 * 2
    ReDim dblTempPMTs(PMT_COLUMN_COUNT - 1, tempDimCount)
    
    FileNum = FreeFile
    Open FileName For Input Access Read As FileNum
    
    If FileNum = -1 Then
         Exit Sub
    End If

    EnableDisableControls False
    UpdateStatus "Loading MT tags"

    Line Input #FileNum, strLineIn
    mLocalPMTCount = 0
    Do While Not EOF(FileNum)
        Line Input #FileNum, strLineIn
                
        strSplitLine = Split(strLineIn, ",")
        For i = 0 To UBound(strSplitLine)
            If i >= PMT_COLUMN_COUNT Then Exit For
            dblTempPMTs(i, mLocalPMTCount) = val(strSplitLine(i))
        Next i
        
        If dblTempPMTs(PMTColumnConstants.pccObservationCount, mLocalPMTCount) > UMCNetAdjDef.MSWarpOptions.MinimumPMTTagObsCount Then
            mLocalPMTCount = mLocalPMTCount + 1
            If mLocalPMTCount >= tempDimCount Then
              ReDim Preserve dblTempPMTs(PMT_COLUMN_COUNT - 1, 2 * tempDimCount)
              tempDimCount = 2 * tempDimCount
            End If
        End If
        
        If mLocalPMTCount Mod 100 = 0 Then
            If mAbortRequested Then Exit Do
            UpdateStatus "Loading MT tags: " & mLocalPMTCount
            DoEvents
        End If
    Loop
    Close FileNum
    
    txtPMTCountLoaded.Text = LongToStringWithCommas(mLocalPMTCount)
    ReDim mLocalPMTs(mLocalPMTCount - 1, PMT_COLUMN_COUNT - 1)
    For i = 0 To mLocalPMTCount - 1
        For j = 0 To PMT_COLUMN_COUNT - 1
            mLocalPMTs(i, j) = dblTempPMTs(j, i)
        Next j
    Next i

    If mAbortRequested Then
        UpdateStatus "Aborted load"
    Else
        UpdateStatus "Ready"
    End If

ExitSub:
    EnableDisableControls True
    Exit Sub
    
LoadPMTsFromFileErrorHandler:
    MsgBox "Error in LoadPMTsFromFile: " & Err.Description
    UpdateStatus "Aborted loading MT tags"
    Resume ExitSub
    
End Sub

Private Sub LoadPMTsFromLegacyDB()

    '------------------------------------------------------------
    ' Load/reload MT tags from Legacy DB
    '------------------------------------------------------------
    Dim eResponse As VbMsgBoxResult
    Dim strMessage As String
    
On Error GoTo LoadPMTsFromLegacyDBErrorHandler

    'ask user if it wants to replace legitimate MT tags DB with legacy DB
    If Not GelAnalysis(CallerID) Is Nothing And Not APP_BUILD_DISABLE_MTS Then
       eResponse = MsgBox("Current display is associated with MT tag database." & vbCrLf _
                    & "Are you sure you want to use a legacy database for search?", vbYesNoCancel, glFGTU)
       If eResponse <> vbYes Then Exit Sub
    End If
    Me.MousePointer = vbHourglass
    EnableDisableControls False
    
    If Len(GelData(CallerID).PathtoDatabase) > 0 Then
        If ConnectToLegacyAMTDB(Me, CallerID, False, True, False) Then
            PopulateLocalPMTsArray
            strMessage = "Loaded MT tags from Legacy DB"
        Else
            strMessage = "Error loading MT tags from Legacy DB"
        End If
    Else
        WarnUserInvalidLegacyDBPath
        strMessage = "Undefined Legacy DB Path"
    End If
    
    EnableDisableControls True
    Me.MousePointer = vbDefault

    UpdateStatus strMessage
    Exit Sub

LoadPMTsFromLegacyDBErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.LoadPMTsFromLegacyDB"
    Resume Next
    
End Sub

Private Sub LoadLCMSFeaturesFromFile()
    Dim OldPath As String
    Dim i As Long
    Dim j As Integer
    
    Dim strLineIn As String
    Dim strSplitLine() As String
    
    Dim FileNum As Integer
    Dim FileName As String
    Dim min_scan_members As Integer
    Dim val_read As Double
    Dim lngClassMemberCount As Long
    
    Dim lngDimCount As Long
    Dim dblTempUMCs() As Double
    Dim Index As Long
    
On Error GoTo LoadLCMSFeaturesFromFileErrorHandler

    If mLoading Then Exit Sub
    
    FileName = SelectFile(Me.hwnd, "Load LC-MS Features File ...", , False, "*.txt", "UMC Files (*umcs.txt)|*umcs.txt|All Files (*.*)|*.*", 1)
    If Len(FileName) = 0 Then
        Exit Sub
    End If
    
    lngDimCount = 1024 * 2
    ReDim dblTempUMCs(FEATURE_COLUMN_COUNT - 1, lngDimCount)
    
    min_scan_members = 3
    

    FileNum = FreeFile
    Open FileName For Input Access Read As FileNum
    
    If FileNum = -1 Then
         Exit Sub
    End If
    
    EnableDisableControls False
    UpdateStatus "Loading LC-MS Features"
    
    Line Input #FileNum, strLineIn
    mLocalFeatureCount = 0
    Do While Not EOF(FileNum)
        Line Input #FileNum, strLineIn
     
        lngClassMemberCount = 0
        strSplitLine = Split(strLineIn, vbTab)
        For i = 0 To UBound(strSplitLine)
            Select Case i
            Case UMCFileColumnConstants.ufcCenterScan
                dblTempUMCs(FeatureColumnConstants.fccScanClassRep, mLocalFeatureCount) = val(strSplitLine(i))
            Case UMCFileColumnConstants.ufcClassMonoisotopicMass
                dblTempUMCs(FeatureColumnConstants.fccClassMass, mLocalFeatureCount) = val(strSplitLine(i))
            Case UMCFileColumnConstants.ufcClassAbundance
                dblTempUMCs(FeatureColumnConstants.fccClassAbundance, mLocalFeatureCount) = val(strSplitLine(i))
                
            Case UMCFileColumnConstants.ufcPairIndex
                dblTempUMCs(FeatureColumnConstants.fccPairIndex, mLocalFeatureCount) = val(strSplitLine(i))
            Case UMCFileColumnConstants.ufcPMTTagID
                dblTempUMCs(FeatureColumnConstants.fccPMTTagID, mLocalFeatureCount) = val(strSplitLine(i))
            Case UMCFileColumnConstants.ufcUMCID
                dblTempUMCs(FeatureColumnConstants.fccUMCID, mLocalFeatureCount) = val(strSplitLine(i))
            Case UMCFileColumnConstants.ufcMemberCount
                lngClassMemberCount = val(strSplitLine(i))
            End Select
        Next i
        
        If lngClassMemberCount + 1 > min_scan_members Then
            mLocalFeatureCount = mLocalFeatureCount + 1
        End If
        If mLocalFeatureCount >= lngDimCount Then
            ReDim Preserve dblTempUMCs(FEATURE_COLUMN_COUNT - 1, 2 * lngDimCount)
            lngDimCount = 2 * lngDimCount
        End If
       
        If mLocalFeatureCount Mod 100 = 0 Then
            If mAbortRequested Then Exit Do
            UpdateStatus "Loading LC-MS Features: " & mLocalFeatureCount
            DoEvents
        End If
    Loop
    Close FileNum
    
    txtFeatureCountLoaded.Text = LongToStringWithCommas(mLocalFeatureCount)
    ReDim mLocalFeatures(mLocalFeatureCount - 1, FEATURE_COLUMN_COUNT - 1)
    For i = 0 To mLocalFeatureCount - 1
        For j = 0 To FEATURE_COLUMN_COUNT - 1
            mLocalFeatures(i, j) = dblTempUMCs(j, i)
        Next j
    Next i
    
    mLocalFeaturesAreFiltered = False

    If mAbortRequested Then
        UpdateStatus "Aborted load"
    Else
        UpdateStatus "Ready"
    End If

ExitSub:
   EnableDisableControls True
    Exit Sub
    
LoadLCMSFeaturesFromFileErrorHandler:
    MsgBox "Error in LoadLCMSFeaturesFromFile: " & Err.Description
    UpdateStatus "Aborted loading LC-MS Features"
    Resume ExitSub
    
End Sub

Private Function LookupMassCalibrationTypeName(eRobustNETWarpMassCalibrationType As rmcUMCRobustNETWarpMassCalibrationType) As String
    Select Case eRobustNETWarpMassCalibrationType
    Case rmcUMCRobustNETWarpMassCalibrationType.rmcMZRegressionRecal
        LookupMassCalibrationTypeName = "m/z coefficient recalibration"
    Case rmcUMCRobustNETWarpMassCalibrationType.rmcScanRegressionRecal
        LookupMassCalibrationTypeName = "mass vs. elution time recalibration"
    Case rmcUMCRobustNETWarpMassCalibrationType.rmcHybridRecal
        LookupMassCalibrationTypeName = "hybrid recalibration"
    Case Else
        LookupMassCalibrationTypeName = "unknown recalibration type"
    End Select
    
End Function

Private Sub MergeHistogramData(varHistogram As Variant, varHistogramOld As Variant, sngBinSize As Single)
    Dim intIndex As Integer
    Dim intIndexTarget As Integer
    Dim intBestIndex As Integer
    
    Dim dblBestDiff As Double
    Dim dblNewDiff As Double
    
On Error GoTo MergeHistogramDataErrorHander


    For intIndex = 0 To UBound(varHistogramOld, 1)
        ' Find the bin in varHistogram that is closest to varHistogramOld(intIndex, 0)
        intBestIndex = 0
        dblBestDiff = Abs(varHistogram(0, 0) - varHistogramOld(intIndex, 0))
        
        For intIndexTarget = 1 To UBound(varHistogram, 1)
            dblNewDiff = Abs(varHistogram(intIndexTarget, 0) - varHistogramOld(intIndex, 0))
            If dblNewDiff < dblBestDiff Then
                dblBestDiff = dblNewDiff
                intBestIndex = intIndexTarget
            End If
        Next intIndexTarget
        
        varHistogram(intBestIndex, 1) = varHistogram(intBestIndex, 1) + varHistogramOld(intIndex, 1)
    Next intIndex
    
    Exit Sub

MergeHistogramDataErrorHander:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.MergeHistogramData"
    
End Sub

Private Sub SetMostRecentPlotViewMode(eNewViewMode As pvmPlotViewModeConstants)
    TraceLog 3, "SetMostRecentPlotViewMode", "Update view mode to: " & eNewViewMode
    mMostRecentPlotViewMode = eNewViewMode

    Select Case mMostRecentPlotViewMode
    Case pvmLinearFitNETResidualsPlot, pvmWarpedFitNETResidualsPlot
        mMostRecentNETResidualsViewMode = mMostRecentPlotViewMode
    Case pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot
        mMostRecentMassVsScanResidualsViewMode = mMostRecentPlotViewMode
    Case pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot
        mMostRecentMassVsMZResidualsViewMode = mMostRecentPlotViewMode
    End Select
 
End Sub

Private Sub UpdateNETResidualsPlot(blnZoomOut As Boolean)

    Dim dblStraightLine() As Double
    Dim blnUseWarpedResiduals As Boolean
    
On Error GoTo UpdateNETResidualsPlotErrorHandler
    
    TraceLog 3, "UpdateNETResidualsPlot", "Current view mode is: " & mMostRecentPlotViewMode

    ReDim dblStraightLine(1, 1)
    
    dblStraightLine(0, 0) = mMinScan
    dblStraightLine(1, 0) = 0
    dblStraightLine(0, 1) = mMaxScan
    dblStraightLine(1, 1) = 0
    
    ctlNETResidual.ClearData
    
    If mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
        blnUseWarpedResiduals = False
        If UBound(mScanMTNETvsLinearNETResidual, 2) <= 0 Then Exit Sub
    Else
        ' Assume ePlotMode = pvmPlotViewModeConstants.pvmWarpedFitNETResidualsPlot
        blnUseWarpedResiduals = True
        If UBound(mScanMTNETvsCustomNETResidual, 2) <= 0 Then Exit Sub
    End If
    
    ' Plot the residuals
    ctlNETResidual.Plots(1).PlotXY dblStraightLine
    
    If blnUseWarpedResiduals Then
        TraceLog 3, "UpdateNETResidualsPlot", "Plotting Warped residuals"
        ctlNETResidual.Plots(2).PlotXY mScanMTNETvsCustomNETResidual
        ctlNETResidual.Caption = "Residuals Using MS Warp NET Alignment"
    Else
        TraceLog 3, "UpdateNETResidualsPlot", "Plotting Linear NET residuals"
        ctlNETResidual.Plots(2).PlotXY mScanMTNETvsLinearNETResidual
        ctlNETResidual.Caption = "Residuals Using Linear NET Fit (prior to warping)"
    End If
    
    If Not blnUseWarpedResiduals And chkShowTransform Then
        TraceLog 3, "UpdateNETResidualsPlot", "Plotting Transform Function on residuals plot"
        ctlNETResidual.Plots(3).PlotXY mScanCustomNetVsLinearNETResidual
        ctlNETResidual.Plots(3).LineWidth = mResidualPlotTransformationFnLineSize
    End If
    
    ctlNETResidual.Axes(1).Caption = "Scan #"
    ctlNETResidual.Axes(2).Caption = "NET Residual  "
    
    ctlNETResidual.Plots(2).LineWidth = mResidualPlotPointSize
    
    If blnZoomOut Then ZoomOutNETResidualsPlot
    DisplayResidualPlotAxisRanges ctlNETResidual, 0, 3
    
    Exit Sub
    
UpdateNETResidualsPlotErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.UpdateNETResidualsPlot"
    
End Sub

Private Sub UpdateMassResidualsMZPlot(blnZoomOut As Boolean)

    Dim dblStraightLine() As Double
    
On Error GoTo UpdateMassResidualsMZPlotErrorHandler
    
    TraceLog 3, "UpdateMassResidualsMZPlot", "Current view mode is: " & mMostRecentPlotViewMode

    ReDim dblStraightLine(1, 1)
    
    dblStraightLine(0, 0) = mMinMZ
    dblStraightLine(1, 0) = 0
    dblStraightLine(0, 1) = mMaxMZ
    dblStraightLine(1, 1) = 0
    
    ctlMassVsMZResidual.ClearData
    
    ' Plot the residuals
    ctlMassVsMZResidual.Plots(1).PlotXY dblStraightLine
    
    If mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmMassResidualsMZPlot Then
        If UBound(mMZVsMassError, 2) <= 0 Then Exit Sub
        TraceLog 3, "UpdateMassResidualsMZPlot", "Plotting Mass residuals vs m/z"
        ctlMassVsMZResidual.Plots(2).PlotXY mMZVsMassError
        ctlMassVsMZResidual.Caption = "Mass Error (ppm) vs m/z"
    
    ElseIf mMostRecentPlotViewMode = pvmMassResidualsCorrectedMZPlot Then
        If UBound(mMZVsMassErrorCorrected, 2) <= 0 Then Exit Sub
        TraceLog 3, "UpdateMassResidualsMZPlot", "Plotting Mass error recalibrated vs m/z"
        ctlMassVsMZResidual.Plots(2).PlotXY mMZVsMassErrorCorrected
        ctlMassVsMZResidual.Caption = "Mass Error (ppm) vs m/z After Recalibration"
    End If
    
    ctlMassVsMZResidual.Axes(1).Caption = "m/z"
    ctlMassVsMZResidual.Axes(2).Caption = "Mass Residual  "
    
    ctlMassVsMZResidual.Plots(2).LineWidth = mResidualPlotPointSize
    
    If blnZoomOut Then ZoomOutMassVsMZResidualsPlot
    DisplayResidualPlotAxisRanges ctlMassVsMZResidual, 0, 1
    
    Exit Sub
    
UpdateMassResidualsMZPlotErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.UpdateMassResidualsMZPlot"
    
End Sub

Private Sub UpdateMassResidualsScanPlot(blnZoomOut As Boolean)

    Dim dblStraightLine() As Double
    
On Error GoTo UpdateMassResidualsScanPlotErrorHandler
    
    TraceLog 3, "UpdateMassResidualsScanPlot", "Current view mode is: " & mMostRecentPlotViewMode

    ReDim dblStraightLine(1, 1)
    
    dblStraightLine(0, 0) = mMinScan
    dblStraightLine(1, 0) = 0
    dblStraightLine(0, 1) = mMaxScan
    dblStraightLine(1, 1) = 0
    
    ctlMassVsScanResidual.ClearData
    
    ' Plot the residuals
    ctlMassVsScanResidual.Plots(1).PlotXY dblStraightLine
    
    If mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmMassResidualsScanPlot Then
        If UBound(mScanVsMassError, 2) <= 0 Then Exit Sub
        TraceLog 3, "UpdateMassResidualsScanPlot", "Plotting Mass residuals vs Scan"
        ctlMassVsScanResidual.Plots(2).PlotXY mScanVsMassError
        ctlMassVsScanResidual.Caption = "Mass Error (ppm) vs Scan #"
    
    ElseIf mMostRecentPlotViewMode = pvmMassResidualsCorrectedScanPlot Then
        If UBound(mScanVsMassErrorCorrected, 2) <= 0 Then Exit Sub
        TraceLog 3, "UpdateMassResidualsScanPlot", "Plotting Mass error recalibrated vs scan"
        ctlMassVsScanResidual.Plots(2).PlotXY mScanVsMassErrorCorrected
        ctlMassVsScanResidual.Caption = "Mass Error (ppm) vs Scan # After Recalibration"
    
    End If
    
    ctlMassVsScanResidual.Axes(1).Caption = "Scan #"
    ctlMassVsScanResidual.Axes(2).Caption = "Mass Residual  "
    
    ctlMassVsScanResidual.Plots(2).LineWidth = mResidualPlotPointSize
    
    If blnZoomOut Then ZoomOutMassVsScanResidualsPlot
    DisplayResidualPlotAxisRanges ctlMassVsScanResidual, 0, 1
    
    Exit Sub
    
UpdateMassResidualsScanPlotErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.UpdateMassResidualsScanPlot"
    
End Sub

Private Function MatchesAreValid() As Boolean
    On Error GoTo MatchesAreValidErrorHandler
    
    If VarType(mMatches) = vbEmpty Then
        MatchesAreValid = False
    Else
        MatchesAreValid = True
    End If
    
    Exit Function

MatchesAreValidErrorHandler:
    MatchesAreValid = False
End Function

Private Function MTPassesFilters(ByVal lngMTIndex As Long, ByVal lngObsCountMinimum As Long, _
                                 ByVal blnUseNETFilter As Boolean, ByVal dblNETMinimum As Double, ByVal dblNETMaximum As Double, _
                                 ByVal blnUseMassFilter As Boolean, ByVal dblMassMinimum As Double, ByVal dblMassMaximum As Double) As Boolean

    Dim blnPassesFilters As Boolean
    
    If mLocalPMTs(lngMTIndex, PMTColumnConstants.pccObservationCount) >= lngObsCountMinimum Then
        blnPassesFilters = True
        If blnUseNETFilter Then
            If mLocalPMTs(lngMTIndex, PMTColumnConstants.pccNET) < dblNETMinimum Or _
               mLocalPMTs(lngMTIndex, PMTColumnConstants.pccNET) > dblNETMaximum Then
                blnPassesFilters = False
            End If
        End If
        
        If blnUseMassFilter Then
            If mLocalPMTs(lngMTIndex, PMTColumnConstants.pccMonoisotopicMass) < dblMassMinimum Or _
               mLocalPMTs(lngMTIndex, PMTColumnConstants.pccMonoisotopicMass) > dblMassMaximum Then
                blnPassesFilters = False
            End If
        End If
    Else
        blnPassesFilters = False
    End If
    
    MTPassesFilters = blnPassesFilters
    
End Function
            
Private Function PlotNETLine() As Boolean
    ' Returns True if the plot is refreshed

On Error GoTo PlotNETLineErrorHandler
    
    If IsNumeric(txtSlope) And IsNumeric(txtIntercept) Then
        mNETLineSlopeSaved = CDblSafe(txtSlope)
        mNETLineInterceptSaved = CDblSafe(txtIntercept)
        
        ctlFlatSurface.SetNetSlopeAndIntercept mNETLineSlopeSaved, mNETLineInterceptSaved
        ctlFlatSurface.ShowNetLine = (chkShowNet.Value)
        ctlFlatSurface.RefreshPlotNow
        
        PlotNETLine = True
    Else
        PlotNETLine = False
    End If

    Exit Function

PlotNETLineErrorHandler:
    Debug.Assert False
    PlotNETLine = False
End Function

Private Sub PlotSurfaceData()
    If Not MatchesAreValid Then Exit Sub
    
    With ctlFlatSurface
        .ZScoreMode = cChkBox(chkSurfaceShowsZScore)
        .Plot3DSimpleSurface (mMatchScores)
        .PlotLine (mAlignmentFunc)
        .SetBounds mMinScan, mMaxScan, mMinNet, mMaxNet
    End With
        
    If Not PlotNETLine Then
        ctlFlatSurface.RefreshPlotNow
    End If
    
End Sub

Private Sub PopulateComboBoxes()
    With cboStepsToPerform
        .Clear
        .AddItem "NET Alignment Only", MassMatchStepsToPerformConstants.mmsWarpTime
        .AddItem "NET and Mass Alignment", MassMatchStepsToPerformConstants.mmsWarpTimeAndMass
        .ListIndex = MassMatchStepsToPerformConstants.mmsWarpTime
    End With
    
    With cboResidualPlotPointSize
        .Clear
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .ListIndex = 2
    End With

    With cboResidualPlotTransformationFnLineSize
        .Clear
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .ListIndex = 2
    End With
    
End Sub

Private Sub PopulateLocalPMTsArray()
    ' Copy the MT tags into mLocalPMTs

    Dim lngIndex As Long
    
On Error GoTo PopulateLocalPMTsArrayErrorHandler
    
    mLocalPMTCount = AMTCnt
    ReDim mLocalPMTs(mLocalPMTCount - 1, PMT_COLUMN_COUNT - 1)
    
    ' Copy from AMTData to mLocalPMTs
    ' Note that AMTData() is a 1-based array
    For lngIndex = 0 To AMTCnt - 1
        With AMTData(lngIndex + 1)
            mLocalPMTs(lngIndex, PMTColumnConstants.pccNET) = .NET
            mLocalPMTs(lngIndex, PMTColumnConstants.pccMonoisotopicMass) = .MW
            mLocalPMTs(lngIndex, PMTColumnConstants.pccObservationCount) = .MSMSObsCount
            mLocalPMTs(lngIndex, PMTColumnConstants.pccPMTTagID) = .ID
        End With
    Next lngIndex
    
    txtPMTCountLoaded.Text = LongToStringWithCommas(mLocalPMTCount)
    Exit Sub

PopulateLocalPMTsArrayErrorHandler:
    Debug.Assert False
    
End Sub

Private Sub PopulateLocalFeaturesArray(ByVal blnLogToAnalysisHistory As Boolean)

    ' Copy the features into mLocalFeatures

    Dim lngUMCIndex As Long
    
    Dim objP1IndFastSearch As FastSearchArrayLong
    Dim objP2IndFastSearch As FastSearchArrayLong
    Dim blnPairsPresent As Boolean
    Dim lngPairIndex As Long
    Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
    Dim udtPairMatchStats() As udtPairMatchStatsType

    Dim dblMZ As Double
    
    Dim blnUseFeature As Boolean
    Dim blnKeepFeaturesBelowBoundary As Boolean
    Dim lngFeatureCountPassingFilter As Long
    
    Dim dblFeatureMZs() As Double
    Dim lngFeatureScans() As Long
    
    Dim strMessage As String
    
On Error GoTo PopulateLocalFeaturesArrayErrorHandler

    If CallerID < 1 Then Exit Sub
    
    ' Initialize the PairIndex lookup objects
    blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)
    
    
    If GelUMC(CallerID).UMCCnt <= 0 Then
        ReDim mLocalFeatures(0, FEATURE_COLUMN_COUNT - 1)
        mLocalFeaturesAreFiltered = True
        mMinMZ = 0
        mMaxMZ = 3000
    Else
        With GelUMC(CallerID)
            ReDim dblFeatureMZs(.UMCCnt)
            ReDim lngFeatureScans(.UMCCnt)
            
            mMinMZ = 100000
            mMaxMZ = -100000
            
            ' First extract out the m/z and scan values for the features; we will use these below
            ' We're also populating mMinMZ and mMaxMZ at this time
            For lngUMCIndex = 0 To .UMCCnt - 1
                With .UMCs(lngUMCIndex)
                    Select Case .ClassRepType
                    Case gldtCS
                        ' It is important to convert the monoisotopic mass of the UMC to m/z using MonoMassToMZ, rather than using the m/z value of the class rep, since the m/z value is not changed when recalibrating the data
                        ' Also, use .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge rather than GelData(CallerID).CSData(.ClassRepInd).Charge
                        dblFeatureMZs(lngUMCIndex) = MonoMassToMZ(.ClassMW, .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge)
                        lngFeatureScans(lngUMCIndex) = GelData(CallerID).CSData(.ClassRepInd).ScanNumber
                        dblMZ = GelData(CallerID).CSData(.ClassRepInd).MZ
                    Case gldtIS
                        ' It is important to convert the monoisotopic mass of the UMC to m/z using MonoMassToMZ, rather than using the m/z value of the class rep, since the m/z value is not changed when recalibrating the data
                        ' Also, use .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge rather than GelData(CallerID).IsoData(.ClassRepInd).Charge
                        dblFeatureMZs(lngUMCIndex) = MonoMassToMZ(.ClassMW, .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge)
                        lngFeatureScans(lngUMCIndex) = GelData(CallerID).IsoData(.ClassRepInd).ScanNumber
                        dblMZ = GelData(CallerID).IsoData(.ClassRepInd).MZ
                    Case Else
                        ' This shouldn't happen; ignore it
                        Debug.Assert False
                    End Select
                End With
                
                If dblMZ < mMinMZ Then
                    mMinMZ = dblMZ
                End If
                If dblMZ > mMaxMZ Then
                    mMaxMZ = dblMZ
                End If
            
            Next lngUMCIndex
            
            If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Then
                lngFeatureCountPassingFilter = .UMCCnt
                mLocalFeaturesAreFiltered = False
            Else
                ' Filtering the UMCs on m/z and/or scan
                ' Need to first determine how many UMCs pass the filters
                
                If mSplitWarpIteration = 1 Then
                    blnKeepFeaturesBelowBoundary = True
                Else
                    blnKeepFeaturesBelowBoundary = False
                End If

                lngFeatureCountPassingFilter = 0
                For lngUMCIndex = 0 To .UMCCnt - 1
                    
                    If FeaturePassesSplitWarpFilter(dblFeatureMZs(lngUMCIndex), lngFeatureScans(lngUMCIndex), blnKeepFeaturesBelowBoundary) Then
                        ' The feature passes the filters; increment the total count
                        lngFeatureCountPassingFilter = lngFeatureCountPassingFilter + 1
                    End If
                Next lngUMCIndex
                
                If UMCNetAdjDef.MSWarpOptions.SplitWarpMode <> swmDisabled Then
                    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmSplitOnMZ Then
                        strMessage = "Features for MS Warp NET Alignment are split at " & CStr(UMCNetAdjDef.MSWarpOptions.SplitWarpMZBoundary) & " m/z"
                    Else
                        strMessage = "Unknown split warp mode for MS Warp NET Alignment"
                    End If
                End If

                strMessage = strMessage & "; using " & LongToStringWithCommas(lngFeatureCountPassingFilter) & " / " & LongToStringWithCommas(.UMCCnt) & " LC-MS Features"
                If blnKeepFeaturesBelowBoundary Then
                    strMessage = strMessage & " below boundary"
                Else
                    strMessage = strMessage & " above boundary"
                End If
                
                If blnLogToAnalysisHistory Then
                    AddToAnalysisHistory CallerID, strMessage
                End If
                
            End If
            
            If lngFeatureCountPassingFilter > 0 Then
                If UBound(mLocalFeatures, 1) <> lngFeatureCountPassingFilter Then
                    ReDim mLocalFeatures(lngFeatureCountPassingFilter - 1, FEATURE_COLUMN_COUNT - 1)
                End If
            Else
                ReDim mLocalFeatures(0, FEATURE_COLUMN_COUNT - 1)
            End If
    
            mLocalFeatureCount = 0
            If lngFeatureCountPassingFilter > 0 Then
                ' Copy data from GelUMC() to mLocalFeatures
                
                ' Step through the UMCs
                ' Copy those that pass the filters into mLocalFeatures (or all if no filters)
                mLocalFeatureCount = 0
                For lngUMCIndex = 0 To .UMCCnt - 1
                
                    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Then
                        blnUseFeature = True
                    Else
                        blnUseFeature = FeaturePassesSplitWarpFilter(dblFeatureMZs(lngUMCIndex), lngFeatureScans(lngUMCIndex), blnKeepFeaturesBelowBoundary)
                    End If
                    
                    If blnUseFeature Then

                        With .UMCs(lngUMCIndex)
                            mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccUMCID) = lngUMCIndex
                            mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccClassMass) = .ClassMW
                            mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccClassAbundance) = .ClassAbundance
                            
                            mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccClassMZ) = dblFeatureMZs(lngUMCIndex)
                            mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccScanClassRep) = lngFeatureScans(lngUMCIndex)
            
                            
                            ' Could extract the MT tag IDs if needed using:
            ''                Dim lngUMCsInViewCount As Long, lngUMCInViewCountDimmed As Long
            ''                Dim udtUMCsInView() As udtUMCMassTagMatchStats          ' 0-based array
            ''
            ''                ExtractMTHitsFromUMCMembers CallerID, lngUMCIndex, False, udtUMCsInView, lngUMCsInViewCount, lngUMCInViewCountDimmed, False, False
            ''                ExtractMTHitsFromUMCMembers CallerID, lngUMCIndex, True, udtUMCsInView, lngUMCsInViewCount, lngUMCInViewCountDimmed, False, False
            
                            mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccPMTTagID) = -1
                            mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccPairIndex) = -1
                            If blnPairsPresent Then
                                lngPairIndex = -1
                                lngPairMatchCount = 0
                                ReDim udtPairMatchStats(0)
                                InitializePairMatchStats udtPairMatchStats(0)
                                
                                lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndex, objP1IndFastSearch, objP2IndFastSearch, True, False, lngPairMatchCount, udtPairMatchStats())
                                
                                If lngPairMatchCount > 0 Then
                                    For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                                        ' Note: Only save the first pair that this UMC belongs to
                                        mLocalFeatures(mLocalFeatureCount, FeatureColumnConstants.fccPairIndex) = udtPairMatchStats(lngPairMatchIndex).PairIndex
                                        Exit For
                                    Next lngPairMatchIndex
                                End If
                            End If
                            
                        End With
                    
                        mLocalFeatureCount = mLocalFeatureCount + 1
                    End If

                Next lngUMCIndex
            End If
            
            If mLocalFeatureCount < .UMCCnt Then
                mLocalFeaturesAreFiltered = True
            Else
                mLocalFeaturesAreFiltered = False
            End If

        End With
    End If
    
    ' Round mMinMZ and mMaxMZ to the nearest 50
    RoundValuesToNearestMultiple mMinMZ, mMaxMZ, 50
            
    txtFeatureCountLoaded.Text = LongToStringWithCommas(mLocalFeatureCount)
    Exit Sub
    
PopulateLocalFeaturesArrayErrorHandler:
    Debug.Assert False
    
End Sub

Private Function FeaturePassesSplitWarpFilter(ByVal dblMZ As Double, ByVal lngScan As Long, ByVal blnKeepFeaturesBelowBoundary As Boolean) As Boolean
    Dim blnUseFeature As Boolean
    
    blnUseFeature = True
    If UMCNetAdjDef.MSWarpOptions.SplitWarpMZBoundary > 0 Then
        If dblMZ < UMCNetAdjDef.MSWarpOptions.SplitWarpMZBoundary Then
            blnUseFeature = blnKeepFeaturesBelowBoundary
        Else
            blnUseFeature = Not blnKeepFeaturesBelowBoundary
        End If
    End If
    
    FeaturePassesSplitWarpFilter = blnUseFeature

End Function

Private Sub PositionControls()
    
    Const button_border As Long = 150
    Const fraborder As Long = 100
    Const graph_border As Long = 200
    
On Error GoTo FormResizeErrorHandler
    
    
    fraStatus.Top = CheckVsMinimum(Me.ScaleHeight - fraStatus.Height - fraborder, 120)
    tbsOptions.Top = CheckVsMinimum(fraStatus.Top - tbsOptions.Height - fraborder / 2, 120)
    
    fraInfo.Top = tbsOptions.Top
    fraNETAlignmentStats.Top = tbsOptions.Top
    fraMassRefinementStats.Top = tbsOptions.Top
    
    cmdSetDefaults.Top = CheckVsMinimum(Me.ScaleHeight - cmdSetDefaults.Height - button_border, 120)
    cmdWarpAlign.Top = cmdSetDefaults.Top
    cmdAbort.Top = cmdWarpAlign.Top
    cmdAbort.Left = cmdWarpAlign.Left
    cboStepsToPerform.Top = fraMassRefinementStats.Top + fraMassRefinementStats.Height - cboStepsToPerform.Height
    
    fraScores.width = CheckVsMinimum(Me.ScaleWidth - fraScores.Left - 25)
    fraScores.Height = CheckVsMinimum(tbsOptions.Top - fraScores.Top - fraborder)
    fraErrors.Height = fraScores.Height
    
    ctlFlatSurface.width = CheckVsMinimum(fraScores.width - 2 * ctlFlatSurface.Left)
    
    graphMassErrors.Height = CheckVsMinimum(fraErrors.Height / 2 - 3 * graph_border / 2)
    graphMassErrors.Top = graph_border
    
    graphNetErrors.Height = graphMassErrors.Height
    graphNetErrors.Top = 2 * graph_border + graphMassErrors.Height
    
    ctlFlatSurface.Height = CheckVsMinimum(fraScores.Height - 2 * graph_border) + graph_border / 3
    ctlFlatSurface.Top = graph_border + graph_border / 3
    
    With ctlCWGraphNI
        .Left = ctlFlatSurface.Left
        .width = ctlFlatSurface.width
        .Height = ctlFlatSurface.Height
        .Top = ctlFlatSurface.Top
    End With
    With ctlNETResidual
        .Left = ctlFlatSurface.Left
        .width = ctlFlatSurface.width
        .Height = ctlFlatSurface.Height
        .Top = ctlFlatSurface.Top
    End With
    With ctlMassVsScanResidual
        .Left = ctlFlatSurface.Left
        .width = ctlFlatSurface.width
        .Height = ctlFlatSurface.Height
        .Top = ctlFlatSurface.Top
    End With
    With ctlMassVsMZResidual
        .Left = ctlFlatSurface.Left
        .width = ctlFlatSurface.width
        .Height = ctlFlatSurface.Height
        .Top = ctlFlatSurface.Top
    End With
    
    Exit Sub
    
FormResizeErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub PossiblyUpdateNETLinePlot()
    If chkShowNet.Value = vbChecked Then
        ' See if txtSlope or txtIntercept contains a different value than was last plotted
        
        If CDblSafe(txtSlope) <> mNETLineSlopeSaved Or CDblSafe(txtIntercept) <> mNETLineInterceptSaved Then
            PlotNETLine
        End If
    End If

End Sub

Private Sub QueryMassMatchProgress()
    Dim intPercentComplete As Integer
    Dim strSecondsElapsed As String
    Dim strMessage As String
    
On Error GoTo QueryMassMatchProgressErrorHandler

    mMassMatchState = mMassMatchObject.GetState
    
    strMessage = mMassMatchObject.ProgressMessage
    strSecondsElapsed = Round((Now - mProcessingStartTime) * 24 * 60 * 60, 1)
    Select Case mMassMatchState
    Case MassMatchProcessingStateConstants.pscRunning
        intPercentComplete = mMassMatchObject.ProgressPercentComplete
        UpdateStatus strMessage & ": " & intPercentComplete & "% complete (" & strSecondsElapsed & " seconds elapsed)"
    Case MassMatchProcessingStateConstants.pscComplete
        FinalizeAlignment strSecondsElapsed
    Case Is >= MassMatchProcessingStateConstants.pscError
        tmrAlignment.Enabled = False
        
        strMessage = "Error during MS Warp NET Alignment; state = " & mMassMatchState & "; " & ConstructHistoryTextForNETAlignment(strSecondsElapsed)
        AddToAnalysisHistory CallerID, strMessage
        AddToAnalysisHistory CallerID, ConstructHistoryTextForMSWarpSettings(strSecondsElapsed)
    
        If mMassMatchState = pscInsufficientMatches Then
            UpdateStatus "Error during MS Warp NET Alignment; Insufficient matches"
        Else
            UpdateStatus strMessage
        End If
        
        If UMCNetAdjDef.RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass Then
            strMessage = "MS Warp Mass Recalibration settings; " & ConstructHistoryTextForMassRecalibration()
            AddToAnalysisHistory CallerID, strMessage
        End If
        
        mAlignmentFinalizedOrAborted = True
    End Select

    If Not tmrAlignment.Enabled Then
        EnableDisableControls True
    End If
    
    Exit Sub

QueryMassMatchProgressErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.QueryMassMatchProgress"

End Sub

Private Function RecalibrateMassesUsingWarpedData() As Boolean
    Dim lngIndex As Long

    Dim dblMassShiftPPM As Double

    Dim dblWeightingSum As Double
    Dim dblIntensitySum As Double
    Dim lngAdjustmentCount As Long
    Dim lngTotalDataPoints As Long
    
    ' Note that the mass shift avg is a weighted average value with units of ppm
    Dim dblMassShiftPPMAvg As Double

    ' Note: If more than MAXIMUM_PERCENTAGE_SHIFTED_OVER_MAX percent of the data
    '  is shifted more than MAXIMUM_MASS_SHIFT_PPM ppm, then the mass calibration will not be applied
    Dim lngDataCountShiftedTooFar As Long
    Dim sngPercentage As Single                  ' Number between 0 and 100
    
    Dim eMassType As glMassToleranceConstants
    
    Dim strMessage As String
    Dim blnSuccess As Boolean
    
    Dim blnKeepFeaturesBelowBoundary As Boolean
    Dim blnUpdateMass As Boolean
    
On Error GoTo RecalibrateMassesUsingWarpedDataErrorHandler

    dblWeightingSum = 0
    dblIntensitySum = 0
    lngAdjustmentCount = 0

        
    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode <> swmDisabled Then
        ' Filtering the UMCs on m/z and/or scan
        
        If mSplitWarpIteration = 1 Then
            blnKeepFeaturesBelowBoundary = True
        Else
            blnKeepFeaturesBelowBoundary = False
        End If
    End If
    
    With GelData(CallerID)
        lngTotalDataPoints = .CSLines + .IsoLines
        
        If lngTotalDataPoints > 0 Then
            ' First count the number of data points that will be shifted more than MAXIMUM_PERCENTAGE_SHIFTED_OVER_MAX ppm
            ' Process CS data
            lngDataCountShiftedTooFar = 0
            For lngIndex = 1 To .CSLines
                If .CSData(lngIndex).MZ > 0 Then
                    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Then
                        blnUpdateMass = True
                    Else
                        blnUpdateMass = FeaturePassesSplitWarpFilter(.CSData(lngIndex).MZ, .CSData(lngIndex).ScanNumber, blnKeepFeaturesBelowBoundary)
                    End If

                    If blnUpdateMass Then
                        If Abs(GetPPMShift(.CSData(lngIndex).MZ, .CSData(lngIndex).ScanNumber)) > MAXIMUM_MASS_SHIFT_PPM Then
                            lngDataCountShiftedTooFar = lngDataCountShiftedTooFar + 1
                        End If
                    End If
                End If
            Next lngIndex
    
            ' Process isotopic data
            For lngIndex = 1 To .IsoLines
                If .IsoData(lngIndex).MonoisotopicMW > 0 Then
                    If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Then
                        blnUpdateMass = True
                    Else
                        blnUpdateMass = FeaturePassesSplitWarpFilter(.IsoData(lngIndex).MZ, .IsoData(lngIndex).ScanNumber, blnKeepFeaturesBelowBoundary)
                    End If

                    If blnUpdateMass Then
                        If Abs(GetPPMShift(.IsoData(lngIndex).MZ, .IsoData(lngIndex).ScanNumber)) > MAXIMUM_MASS_SHIFT_PPM Then
                            lngDataCountShiftedTooFar = lngDataCountShiftedTooFar + 1
                        End If
                    End If
                End If
            Next lngIndex
            
            sngPercentage = lngDataCountShiftedTooFar / CDbl(lngTotalDataPoints) * 100#
            If sngPercentage > MAXIMUM_PERCENTAGE_SHIFTED_OVER_MAX Then
                ' Too much of the data is shifted more than MAXIMUM_PERCENTAGE_SHIFTED_OVER_MAX
                
                strMessage = "Error: MS Warp Mass Recalibration could not be performed since " & Format(sngPercentage, "0.0") & "% of the data would be shifted more than " & Trim(MAXIMUM_MASS_SHIFT_PPM) & " ppm; " & ConstructHistoryTextForMassRecalibration()
                AddToAnalysisHistory CallerID, strMessage
        
                lngAdjustmentCount = 0
            Else
                ' Process CS data
                For lngIndex = 1 To .CSLines
                    If .CSData(lngIndex).MZ > 0 Then
                    
                        If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Then
                            blnUpdateMass = True
                        Else
                            blnUpdateMass = FeaturePassesSplitWarpFilter(.CSData(lngIndex).MZ, .CSData(lngIndex).ScanNumber, blnKeepFeaturesBelowBoundary)
                        End If
                        
                        If blnUpdateMass Then
                            ' Note: We're taking the value returned from GetPPMShift times -1
                            dblMassShiftPPM = -1 * GetPPMShift(.CSData(lngIndex).MZ, .CSData(lngIndex).ScanNumber)
                            
                            If dblMassShiftPPM <> 0 Then
                                dblWeightingSum = dblWeightingSum + (dblMassShiftPPM * .CSData(lngIndex).Abundance)
                                dblIntensitySum = dblIntensitySum + .CSData(lngIndex).Abundance
                                
                                MassCalibrationApplyAdjustmentOnePoint .CSData(lngIndex), dblMassShiftPPM, False
                                lngAdjustmentCount = lngAdjustmentCount + 1
                            End If
                        End If
                        
                        If lngIndex Mod 5000 = 0 Then
                            UpdateStatus "Updating mass calibration: " & Round(lngIndex / lngTotalDataPoints * 100, 1) & "% done"
                        End If
                    End If
                Next lngIndex
        
                ' Process isotopic data
                For lngIndex = 1 To .IsoLines
                    If .IsoData(lngIndex).MonoisotopicMW > 0 Then
                        
                        If UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled Then
                            blnUpdateMass = True
                        Else
                            blnUpdateMass = FeaturePassesSplitWarpFilter(.IsoData(lngIndex).MZ, .IsoData(lngIndex).ScanNumber, blnKeepFeaturesBelowBoundary)
                        End If
                        
                        If blnUpdateMass Then
                            ' Note: We're taking the value returned from GetPPMShift times -1
                            dblMassShiftPPM = -1 * GetPPMShift(.IsoData(lngIndex).MZ, .IsoData(lngIndex).ScanNumber)
            
                            If dblMassShiftPPM <> 0 Then
                                dblWeightingSum = dblWeightingSum + (dblMassShiftPPM * .IsoData(lngIndex).Abundance)
                                dblIntensitySum = dblIntensitySum + .IsoData(lngIndex).Abundance
                                
                                MassCalibrationApplyAdjustmentOnePoint .IsoData(lngIndex), dblMassShiftPPM, True
                                lngAdjustmentCount = lngAdjustmentCount + 1
                            End If
                        End If
        
                        If lngIndex Mod 5000 = 0 Then
                            UpdateStatus "Updating mass calibration: " & Round((lngIndex + .CSLines) / lngTotalDataPoints * 100, 1) & "% done"
                        End If
                    End If
                Next lngIndex
            End If
        End If
    End With

    If lngAdjustmentCount > 0 Then
        ' Compute the intensity-weighted average mass shift
        If dblIntensitySum > 0 Then
            dblMassShiftPPMAvg = Round(dblWeightingSum / dblIntensitySum, 4)
        Else
            dblMassShiftPPMAvg = 0
        End If

        strMessage = "MS Warp Mass Recalibration performed; " & ConstructHistoryTextForMassRecalibration()
        AddToAnalysisHistory CallerID, strMessage
    
        ' Update .MassCalibrationInfo with the bulk stats
        eMassType = glMassToleranceConstants.gltPPM
        blnSuccess = MassCalibrationUpdateHistory(CallerID, dblMassShiftPPMAvg, eMassType, True, True, glbPreferencesExpanded.ErrorPlottingOptions.MassBinSizePPM, True)

        ' Recompute the UMC class mass stats
        UpdateStatus "Updating mass calibration: 100% done; now recomputing LC-MS Feature stats"
        blnSuccess = CalculateClasses(CallerID, True, False, Me)

        ' Update mLocalFeatures
        PopulateLocalFeaturesArray False

        ' Update the stats displayed
        UpdateMassCalibrationStats
    End If
    
    RecalibrateMassesUsingWarpedData = True
    Exit Function

RecalibrateMassesUsingWarpedDataErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.RecalibrateMassesUsingWarpedData"
    RecalibrateMassesUsingWarpedData = False
        
End Function

Private Sub RequestAbort()

On Error GoTo RequestAbortErrorHandler
    If Not mMassMatchObject Is Nothing Then
        mMassMatchObject.Abort
    End If
    
    tmrAlignment.Enabled = False
    EnableDisableControls True
    
    mAlignmentFinalizedOrAborted = True
    
    Exit Sub
    
RequestAbortErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmMSAlign.RequestAbort"
    
End Sub

Private Sub ResidualPlotSetRange()
    Select Case mMostRecentPlotViewMode
    Case pvmLinearFitNETResidualsPlot, pvmWarpedFitNETResidualsPlot
        ResidualPlotSetRangeWork ctlNETResidual, 0, mMaxScan * 10, mMinScan, mMaxScan, mLinearNETResidualMin - 2, mLinearNETResidualMax + 2, mLinearNETResidualMin, mLinearNETResidualMax
        DisplayResidualPlotAxisRangesCurrent
    Case pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot
        ResidualPlotSetRangeWork ctlMassVsScanResidual, 0, mMaxScan * 10, mMinScan, mMaxScan, mMassResidualMin - 500, mMassResidualMax + 500, mMassResidualMin, mMassResidualMax
        DisplayResidualPlotAxisRangesCurrent
    Case pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot
        ResidualPlotSetRangeWork ctlMassVsMZResidual, 0, mMaxMZ * 10, mMinMZ, mMaxMZ, mMassResidualMin - 500, mMassResidualMax + 500, mMassResidualMin, mMassResidualMax
        DisplayResidualPlotAxisRangesCurrent
    Case Else
        ' Residuals plot is not visible; ignore this
    End Select
End Sub

Private Sub ResidualPlotSetRangeWork(ByRef objPlot As CWGraph, ByVal dblMinimumXVal As Double, ByVal dblMaximumXVal As Double, ByVal dblDefaultMinimumXVal As Double, ByVal dblDefaultMaximumXVal As Double, ByVal dblMinimumYVal As Double, ByVal dblMaximumYVal As Double, ByVal dblDefaultMinimumYVal As Double, ByVal dblDefaultMaximumYVal As Double)

On Error GoTo ResidualPlotSetRangeWorkErrorHandler

    mUpdatingPlotRanges = True
    
    With objPlot
        With .Plots(1).XAxis
            .AutoScale = False
            .Minimum = ValidateTextboxValueDbl(txtResidualPlotMinX, dblMinimumXVal, dblMaximumXVal, dblDefaultMinimumXVal)
            .Maximum = ValidateTextboxValueDbl(txtResidualPlotMaxX, dblMinimumXVal, dblMaximumXVal, dblDefaultMaximumXVal)
        End With
        
        With .Plots(1).YAxis
            .AutoScale = False
            .Minimum = ValidateTextboxValueDbl(txtResidualPlotMinY, dblMinimumYVal, dblMaximumYVal, dblDefaultMinimumYVal)
            .Maximum = ValidateTextboxValueDbl(txtResidualPlotMaxY, dblMinimumYVal, dblMaximumYVal, dblDefaultMaximumYVal)
        End With
    End With

    mUpdatingPlotRanges = False
    Exit Sub

ResidualPlotSetRangeWorkErrorHandler:
    Debug.Print "Error in ResidualPlotSetRangeWork: " & Err.Description
    Debug.Assert False
    Resume Next
End Sub

Private Sub RoundValuesToNearestMultiple(ByRef sngValueA As Single, ByRef sngValueB As Single, ByVal intRoundingDivisor As Integer)
    Dim sngNewValue As Single
   
    sngValueA = Int(sngValueA / intRoundingDivisor) * intRoundingDivisor
    
    sngNewValue = Int(sngValueB / intRoundingDivisor) * intRoundingDivisor
    If sngValueB > sngNewValue Then
        sngValueB = sngNewValue + intRoundingDivisor
    Else
        sngValueB = sngNewValue
    End If
End Sub

Public Function Save3DViewToClipboardOrEMF(strFilePath As String, Optional blnSaveHQ As Boolean = False) As Boolean
    ' If strFilePath is empty then copies to the clipboard
    ' Otherwise, saves to strFilePath
        
    Dim eCurrentViewMode As pvmPlotViewModeConstants
    
On Error GoTo Save3DViewErrorHandler

    eCurrentViewMode = mMostRecentPlotViewMode
    
    If Not (eCurrentViewMode = pvm2DView Or eCurrentViewMode = pvm3DView) Then
        UpdatePlotViewModeTo3DView
    End If

    If Len(strFilePath) > 0 Then
        ' Note: The .ControlImageEx function is only available in Measurement Studio v6.0 if you
        '  download and install the patch from http://digital.ni.com/softlib.nsf/websearch/2AAC97491D073A6C86256EEF005374CE?opendocument&node=132060_US
        ' After updating, the c:\windows\system32\cwui.ocx file should be 2,335,240 bytes with date 7/24/2004 2:20 am
        ' Also, make sure the installer does not install an out-of-date cwui.ocx file in the c:\program files\viper folder
        If blnSaveHQ Then
            SavePicture ctlCWGraphNI.ControlImageEx(400, 400), strFilePath
        Else
            SavePicture ctlCWGraphNI.ControlImageEx(ctlCWGraphNI.width / 15, ctlCWGraphNI.Height / 15), strFilePath
        End If
    Else
        Clipboard.Clear
        Clipboard.SetData ctlCWGraphNI.ControlImage, vbCFMetafile
    End If
    
    If Not (eCurrentViewMode = pvm2DView Or eCurrentViewMode = pvm3DView) Then
        UpdatePlotViewMode eCurrentViewMode
    End If

    Save3DViewToClipboardOrEMF = True
    Exit Function
    
Save3DViewErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If Len(strFilePath) > 0 Then
            MsgBox "Error saving 3D view to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
        Else
            MsgBox "Error copying 3D view to clipboard: " & Err.Description, vbExclamation + vbOKOnly, "Error"
        End If
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.Save3DViewToClipboardOrEMF"
    End If
 
    Save3DViewToClipboardOrEMF = False

End Function

Private Sub Save3DViewToFile(ePicfileType As pftPictureFileTypeConstants)
    Dim strFilePath As String
    Dim objRemoteSaveFileHandler As clsRemoteSaveFileHandler
    Dim strWorkingFilePath As String
    Dim blnSuccess As Boolean
    
On Error GoTo Save3DViewToFileErrorHandler

    Select Case ePicfileType
    Case pftPictureFileTypeConstants.pftEMF, pftPictureFileTypeConstants.pftWMF
        ' Saving EMF file
        strFilePath = SelectFile(Me.hwnd, "Save picture as EMF ...", , True, "*.emf", "EMF Files (*.emf)|*.emf|All Files (*.*)|*.*", 1)

        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "emf")
            Set objRemoteSaveFileHandler = New clsRemoteSaveFileHandler
            strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
            
            Save3DViewToClipboardOrEMF strWorkingFilePath
            blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
        End If
    Case Else
        ' Includes pftPictureFileTypeConstants.pftPNG
        ' Saving PNG file
        strFilePath = SelectFile(Me.hwnd, "Save picture as PNG ...", , True, "*.png", "PNG Files (*.png)|*.png|All Files (*.*)|*.*", 1)
        
        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "png")
            Save3DViewToPNG strFilePath
        End If
    End Select

    Exit Sub
    
Save3DViewToFileErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving 3D view to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.Save3DViewToFile"
    End If
    
End Sub

Public Function Save3DViewToPNG(strFilePath As String) As Boolean
    Dim strEmfFilePath As String, strWorkingFilePath As String
    Dim blnSuccess As Boolean
    Dim lngReturn As Long
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
    
On Error GoTo Save3DViewToPNGErrorHandler

    strFilePath = FileExtensionForce(strFilePath, "png")
    strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
    
    strEmfFilePath = FileExtensionForce(strWorkingFilePath, "emf")
    blnSuccess = Save3DViewToClipboardOrEMF(strEmfFilePath, True)
    
    If blnSuccess Then
        lngReturn = ConvertEmfToPng(strEmfFilePath, strWorkingFilePath, ctlCWGraphNI.width / Screen.TwipsPerPixelX, ctlCWGraphNI.Height / Screen.TwipsPerPixelY)
        If lngReturn = 0 Then
            blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
        Else
            objRemoteSaveFileHandler.DeleteTempFile
            blnSuccess = False
        End If
    End If
    
    Save3DViewToPNG = blnSuccess
    Exit Function

Save3DViewToPNGErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving 3D view to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.Save3DViewToPNG"
    End If
       
End Function

Public Function SaveFlatViewToClipboardOrEMF(strFilePath As String) As Boolean
    ' If strFilePath is empty then copies to the clipboard
    ' Otherwise, saves to strFilePath
    
    Dim eCurrentViewMode As pvmPlotViewModeConstants
    
On Error GoTo SaveFlatViewErrorHandler
    
    eCurrentViewMode = mMostRecentPlotViewMode
    If Not eCurrentViewMode = pvmFlatView Then
        UpdatePlotViewModeToFlatView
    End If
         
    Me.ctlFlatSurface.RefreshPlotNow
    If Len(strFilePath) > 0 Then
        Me.ctlFlatSurface.Draw2EMF strFilePath
    Else
        Me.ctlFlatSurface.Draw2EMF2Clipboard
    End If
    
    If eCurrentViewMode <> pvmFlatView Then
        UpdatePlotViewMode eCurrentViewMode
    End If
    
    SaveFlatViewToClipboardOrEMF = True
    Exit Function

SaveFlatViewErrorHandler:
   If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If Len(strFilePath) > 0 Then
            MsgBox "Error saving flat view to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
        Else
            MsgBox "Error copying flat view to clipboard: " & Err.Description, vbExclamation + vbOKOnly, "Error"
        End If
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.SaveFlatViewToClipboardOrEMF"
    End If
 
    SaveFlatViewToClipboardOrEMF = False
    
End Function

Private Sub SaveFlatViewToFile(ePicfileType As pftPictureFileTypeConstants)
    Dim strFilePath As String
    Dim objRemoteSaveFileHandler As clsRemoteSaveFileHandler
    Dim strWorkingFilePath As String
    Dim blnSuccess As Boolean
    
On Error GoTo SaveFlatViewToFileErrorHandler

    Select Case ePicfileType
    Case pftPictureFileTypeConstants.pftEMF, pftPictureFileTypeConstants.pftWMF
        ' Saving EMF file
        strFilePath = SelectFile(Me.hwnd, "Save picture as EMF ...", , True, "*.emf", "EMF Files (*.emf)|*.emf|All Files (*.*)|*.*", 1)

        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "emf")
            Set objRemoteSaveFileHandler = New clsRemoteSaveFileHandler
            strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
            
            SaveFlatViewToClipboardOrEMF strWorkingFilePath
            blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
        End If
    Case Else
        ' Includes pftPictureFileTypeConstants.pftPNG
        ' Saving PNG file
        strFilePath = SelectFile(Me.hwnd, "Save picture as PNG ...", , True, "*.png", "PNG Files (*.png)|*.png|All Files (*.*)|*.*", 1)
        
        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "png")
            SaveFlatViewToPNG strFilePath
        End If
    End Select

    Exit Sub
    
SaveFlatViewToFileErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving flat view to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.SaveFlatViewToFile"
    End If
    
End Sub

Public Function SaveFlatViewToPNG(strFilePath As String) As Boolean
    Dim strEmfFilePath As String, strWorkingFilePath As String
    Dim blnSuccess As Boolean
    Dim lngReturn As Long
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
    
On Error GoTo SaveFlatViewToPNGErrorHandler

    strFilePath = FileExtensionForce(strFilePath, "png")
    strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
    
    strEmfFilePath = FileExtensionForce(strWorkingFilePath, "emf")
    blnSuccess = SaveFlatViewToClipboardOrEMF(strEmfFilePath)
    
    If blnSuccess Then
        lngReturn = ConvertEmfToPng(strEmfFilePath, strWorkingFilePath, ctlFlatSurface.width / Screen.TwipsPerPixelX, ctlFlatSurface.Height / Screen.TwipsPerPixelY)
        If lngReturn = 0 Then
            blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
        Else
            objRemoteSaveFileHandler.DeleteTempFile
            blnSuccess = False
        End If
    End If
    
    SaveFlatViewToPNG = blnSuccess
    Exit Function

SaveFlatViewToPNGErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving flat view to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.SaveFlatViewToPNG"
    End If
        
End Function

Private Function SaveMassVsMZResidualsPlotToClipboardOrEMF(ByVal strFilePath As String, Optional ByVal blnSaveHQ As Boolean = False) As Boolean
    SaveMassVsMZResidualsPlotToClipboardOrEMF = SaveResidualsPlotToClipboardOrEMF( _
                ctlMassVsMZResidual, _
                pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot, _
                strFilePath, blnSaveHQ)
End Function

Public Function SaveMassVsMZResidualsPlotToFile(ByVal ePicfileType As pftPictureFileTypeConstants) As Boolean
    SaveMassVsMZResidualsPlotToFile = SaveResidualsPlotToFile( _
                ctlMassVsMZResidual, _
                pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot, _
                ePicfileType)
End Function

Public Function SaveMassVsMZResidualsPlotToPNG(strFilePath As String) As Boolean
    SaveMassVsMZResidualsPlotToPNG = SaveResidualsPlotToPNG( _
                ctlMassVsMZResidual, _
                pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot, _
                strFilePath)
End Function


Private Function SaveMassVsScanResidualsPlotToClipboardOrEMF(ByVal strFilePath As String, Optional ByVal blnSaveHQ As Boolean = False) As Boolean
    SaveMassVsScanResidualsPlotToClipboardOrEMF = SaveResidualsPlotToClipboardOrEMF( _
                ctlMassVsScanResidual, _
                pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot, _
                strFilePath, blnSaveHQ)
End Function

Public Function SaveMassVsScanResidualsPlotToFile(ByVal ePicfileType As pftPictureFileTypeConstants) As Boolean
    SaveMassVsScanResidualsPlotToFile = SaveResidualsPlotToFile( _
                ctlMassVsScanResidual, _
                pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot, _
                ePicfileType)
End Function

Public Function SaveMassVsScanResidualsPlotToPNG(strFilePath As String) As Boolean
    SaveMassVsScanResidualsPlotToPNG = SaveResidualsPlotToPNG( _
                ctlMassVsScanResidual, _
                pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot, _
                strFilePath)
End Function


Private Function SaveNETResidualsPlotToClipboardOrEMF(ByVal strFilePath As String, Optional ByVal blnSaveHQ As Boolean = False) As Boolean
    SaveNETResidualsPlotToClipboardOrEMF = SaveResidualsPlotToClipboardOrEMF( _
                ctlNETResidual, _
                pvmLinearFitNETResidualsPlot, pvmWarpedFitNETResidualsPlot, _
                strFilePath, blnSaveHQ)
End Function

Public Function SaveNETResidualsPlotToFile(ByVal ePicfileType As pftPictureFileTypeConstants) As Boolean
   SaveNETResidualsPlotToFile = SaveResidualsPlotToFile( _
                ctlNETResidual, _
                pvmLinearFitNETResidualsPlot, pvmWarpedFitNETResidualsPlot, _
                ePicfileType)
End Function

Public Function SaveNETResidualsPlotToPNG(strFilePath As String) As Boolean
    SaveNETResidualsPlotToPNG = SaveResidualsPlotToPNG( _
                ctlNETResidual, _
                pvmLinearFitNETResidualsPlot, pvmWarpedFitNETResidualsPlot, _
                strFilePath)
End Function


Private Function SaveResidualsPlotToClipboardOrEMF(ByRef ctlPlot As CWGraph, ByVal eLinearPlotViewMode As pvmPlotViewModeConstants, ByVal eWarpedPlotViewMode As pvmPlotViewModeConstants, ByVal strFilePath As String, ByVal blnSaveHQ As Boolean) As Boolean
    ' If strFilePath is empty then copies to the clipboard
    ' Otherwise, saves to strFilePath
        
    Dim eCurrentViewMode As pvmPlotViewModeConstants
    
On Error GoTo SaveResidualsPlotErrorHandler

    TraceLog 5, "SaveResidualsPlotToClipboardOrEMF", "Save " & ctlPlot.Name & " to: " & strFilePath

    eCurrentViewMode = mMostRecentPlotViewMode
    If Not (eCurrentViewMode = eLinearPlotViewMode Or eCurrentViewMode = eWarpedPlotViewMode) Then
        TraceLog 3, "SaveResidualsPlotToClipboardOrEMF", ctlPlot.Name & " is visible, but the View Mode is " & mMostRecentPlotViewMode & "; updating View Mode to " & eWarpedPlotViewMode
        UpdatePlotViewMode eWarpedPlotViewMode
    End If

    If Len(strFilePath) > 0 Then
        ' Note: The .ControlImageEx function is only available in Measurement Studio v6.0 if you
        '  download and install the patch from http://digital.ni.com/softlib.nsf/websearch/2AAC97491D073A6C86256EEF005374CE?opendocument&node=132060_US
        ' After updating, the c:\windows\system32\cwui.ocx file should be 2,335,240 bytes with date 7/24/2004 2:20 am
        ' Also, make sure the installer does not install an out-of-date cwui.ocx file in the c:\program files\viper folder
        If blnSaveHQ Then
            SavePicture ctlPlot.ControlImageEx(400, 400), strFilePath
        Else
            SavePicture ctlPlot.ControlImageEx(ctlPlot.width / 15, ctlPlot.Height / 15), strFilePath
        End If
    Else
        Clipboard.Clear
        Clipboard.SetData ctlPlot.ControlImage, vbCFMetafile
    End If
    
    If Not (eCurrentViewMode = eLinearPlotViewMode Or eCurrentViewMode = eWarpedPlotViewMode) Then
        TraceLog 3, "SaveResidualsPlotToClipboardOrEMF", "Updating View Mode back to " & eCurrentViewMode
        UpdatePlotViewMode eCurrentViewMode
    End If

    SaveResidualsPlotToClipboardOrEMF = True
    Exit Function
    
SaveResidualsPlotErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If Len(strFilePath) > 0 Then
            MsgBox "Error saving residuals plot from " & ctlPlot.Name & " to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
        Else
            MsgBox "Error copying residuals plot from " & ctlPlot.Name & " to clipboard: " & Err.Description, vbExclamation + vbOKOnly, "Error"
        End If
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.SaveResidualsPlotToClipboardOrEMF"
    End If
 
    SaveResidualsPlotToClipboardOrEMF = False

End Function

Private Function SaveResidualsPlotToFile(ByRef ctlPlot As CWGraph, ByVal eLinearPlotViewMode As pvmPlotViewModeConstants, ByVal eWarpedPlotViewMode As pvmPlotViewModeConstants, ByVal ePicfileType As pftPictureFileTypeConstants) As Boolean
    Const SAVE_HQ As Boolean = False
    Dim strFilePath As String
    Dim blnSuccess As Boolean
    
On Error GoTo SaveResidualsPlotToFileErrorHandler

    Select Case ePicfileType
    Case pftPictureFileTypeConstants.pftEMF, pftPictureFileTypeConstants.pftWMF
        ' Saving EMF file
        strFilePath = SelectFile(Me.hwnd, "Save picture as EMF ...", , True, "*.emf", "EMF Files (*.emf)|*.emf|All Files (*.*)|*.*", 1)

        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "emf")
            blnSuccess = SaveResidualsPlotToClipboardOrEMF(ctlPlot, eLinearPlotViewMode, eWarpedPlotViewMode, strFilePath, SAVE_HQ)
        End If
    Case Else
        ' Includes pftPictureFileTypeConstants.pftPNG
        ' Saving PNG file
        strFilePath = SelectFile(Me.hwnd, "Save picture as PNG ...", , True, "*.png", "PNG Files (*.png)|*.png|All Files (*.*)|*.*", 1)
        
        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "png")
            blnSuccess = SaveResidualsPlotToPNG(ctlPlot, eLinearPlotViewMode, eWarpedPlotViewMode, strFilePath)
        End If
    End Select

    SaveResidualsPlotToFile = blnSuccess
  Exit Function
    
SaveResidualsPlotToFileErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving residuals plot to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.SaveResidualsPlotToFile"
    End If
    SaveResidualsPlotToFile = False
End Function

Private Function SaveResidualsPlotToPNG(ByRef ctlPlot As CWGraph, ByVal eLinearPlotViewMode As pvmPlotViewModeConstants, ByVal eWarpedPlotViewMode As pvmPlotViewModeConstants, ByVal strFilePath As String) As Boolean
    Dim strEmfFilePath As String, strWorkingFilePath As String
    Dim blnSuccess As Boolean
    Dim lngReturn As Long
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
    
On Error GoTo SaveResidualsPlotToPNGErrorHandler

    strFilePath = FileExtensionForce(strFilePath, "png")
    strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
    strEmfFilePath = FileExtensionForce(strWorkingFilePath, "emf")
    
    blnSuccess = SaveResidualsPlotToClipboardOrEMF(ctlPlot, eLinearPlotViewMode, eWarpedPlotViewMode, strEmfFilePath, True)
      
    If blnSuccess Then
        lngReturn = ConvertEmfToPng(strEmfFilePath, strWorkingFilePath, ctlPlot.width / Screen.TwipsPerPixelX, ctlPlot.Height / Screen.TwipsPerPixelY)
        If lngReturn = 0 Then
            blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
        Else
            objRemoteSaveFileHandler.DeleteTempFile
            blnSuccess = False
        End If
    End If
    
    SaveResidualsPlotToPNG = blnSuccess
    Exit Function

SaveResidualsPlotToPNGErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving residuals plot from " & ctlPlot.Name & " to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.SaveResidualsPlotToPNG"
    End If
    SaveResidualsPlotToPNG = False
    
End Function

Private Sub SetDefaultOptions()
    UpdateControlValues True
End Sub

Public Sub SetPlotPointSize(intPointSize As Integer)
    ' intPointSize should be between 0 and 5
    On Error Resume Next
    cboResidualPlotPointSize.ListIndex = intPointSize
End Sub

Private Sub SetWindowSize(eNewSize As fwsFormWindowSizeConstants)
    Select Case eNewSize
    Case fwsSizeA
        Me.width = 759 * Screen.TwipsPerPixelX
        Me.Height = 551 * Screen.TwipsPerPixelY
    Case fwsSizeB
        Me.width = 800 * Screen.TwipsPerPixelX
        Me.Height = 600 * Screen.TwipsPerPixelY
    Case fwsSizeC
        Me.width = 1024 * Screen.TwipsPerPixelX
        Me.Height = 768 * Screen.TwipsPerPixelY
    Case fwsSizeD
        Me.width = 1280 * Screen.TwipsPerPixelX
        Me.Height = 1024 * Screen.TwipsPerPixelY
    Case Else
        ' Leave unchanged
        Debug.Assert False
    End Select
End Sub
Private Sub ShellSortSingleWithParallelString(ByRef sngValues() As Single, ByRef strText() As String, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)

    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim sngValSaved As Single
    Dim strValSaved As String
    
    Dim lngSwapCount As Long
    
On Error GoTo ShellSortSingleWithParallelStringErrorHandler

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            sngValSaved = sngValues(lngIndex)
            strValSaved = strText(lngIndex)
                
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If sngValues(lngIndexCompare) <= sngValSaved Then Exit For
                
                sngValues(lngIndexCompare + lngIncrement) = sngValues(lngIndexCompare)
                strText(lngIndexCompare + lngIncrement) = strText(lngIndexCompare)
            
                lngSwapCount = lngSwapCount + 1
            Next lngIndexCompare
            
            sngValues(lngIndexCompare + lngIncrement) = sngValSaved
            strText(lngIndexCompare + lngIncrement) = strValSaved
        
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop
    
    'Debug.Print "Swapped " & lngSwapCount & " times using Shell Sort"
    
    Exit Sub

ShellSortSingleWithParallelStringErrorHandler:
    Debug.Assert False
End Sub

Private Sub ShellSortTransformScanResidual()
    
    Dim lngLowIndex As Long
    Dim lngHighIndex As Long

    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim dblValSaved1 As Double
    Dim dblValSaved2 As Double
    
    Dim lngSwapCount As Long
    
On Error GoTo ShellSort2DArrayErrorHandler

    lngLowIndex = 0
    lngHighIndex = UBound(mScanCustomNetVsLinearNETResidual, 2)

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            dblValSaved1 = mScanCustomNetVsLinearNETResidual(0, lngIndex)
            dblValSaved2 = mScanCustomNetVsLinearNETResidual(1, lngIndex)
                
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If mScanCustomNetVsLinearNETResidual(0, lngIndexCompare) <= dblValSaved1 Then Exit For
                
                mScanCustomNetVsLinearNETResidual(0, lngIndexCompare + lngIncrement) = mScanCustomNetVsLinearNETResidual(0, lngIndexCompare)
                mScanCustomNetVsLinearNETResidual(1, lngIndexCompare + lngIncrement) = mScanCustomNetVsLinearNETResidual(1, lngIndexCompare)
            
                lngSwapCount = lngSwapCount + 1
            Next lngIndexCompare
            
            mScanCustomNetVsLinearNETResidual(0, lngIndexCompare + lngIncrement) = dblValSaved1
            mScanCustomNetVsLinearNETResidual(1, lngIndexCompare + lngIncrement) = dblValSaved2
        
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop
    
    'Debug.Print "Swapped " & lngSwapCount & " times using Shell Sort"
    
    Exit Sub

ShellSort2DArrayErrorHandler:
    Debug.Assert False
  
End Sub

Private Sub ShowHideTransformLine(blnShowLine As Boolean)

On Error GoTo ShowHideTransformLineErrorHandler
    
    ctlFlatSurface.ShowLine = blnShowLine
    ctlFlatSurface.RefreshPlotNow
    TraceLog 3, "chkShowTransform_Click", "Call UpdateResidualPlots"
    UpdateResidualPlots
    
    Exit Sub
    
ShowHideTransformLineErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error showing or hiding the transform function line: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.ShowHideTransformLine"
    End If
    End Sub
    
Private Sub SortTransformScanResidual()
    Dim last_index As Long
    Dim next_index As Long
    
    ' Sort the scan residuals
    ShellSortTransformScanResidual
  
    ' Remove redundant points
    last_index = 0
    next_index = 0
    Do While next_index < UBound(mScanCustomNetVsLinearNETResidual, 2)
        last_index = next_index
        next_index = last_index + 1
        Do While next_index <= UBound(mScanCustomNetVsLinearNETResidual, 2)
            If mScanCustomNetVsLinearNETResidual(0, last_index) = mScanCustomNetVsLinearNETResidual(0, next_index) Then
                next_index = next_index + 1
            Else
                Exit Do
            End If
        Loop
        If next_index > UBound(mScanCustomNetVsLinearNETResidual, 2) Then Exit Do
        mScanCustomNetVsLinearNETResidual(0, last_index + 1) = mScanCustomNetVsLinearNETResidual(0, next_index)
        mScanCustomNetVsLinearNETResidual(1, last_index + 1) = mScanCustomNetVsLinearNETResidual(1, next_index)
    Loop
    
    If last_index < UBound(mScanCustomNetVsLinearNETResidual, 2) Then
        ReDim Preserve mScanCustomNetVsLinearNETResidual(1, last_index)
    End If
End Sub

Private Sub ShowAboutBox()
    Dim strMessage As String
    
    strMessage = "Algorithm developed by Navdeep Jaitly.  " & vbCrLf & vbCrLf & _
                 "Journal Article: " & _
                 "'Robust Algorithm for Alignment of Liquid Chromatography-Mass Spectrometry " & _
                 "Analyses in an Accurate Mass and Time Tag Data Analysis Pipeline', " & _
                 "N. Jaitly, M.E. Monroe, V.A. Petyuk, T.R.W. Clauss, J.N. Adkins, and R.D. Smith.  " & _
                 "Analytical Chemistry, 78 (21), 7397-7409 (2006).  " & vbCrLf & vbCrLf & _
                 "Implementation by Navdeep Jaitly and Matthew Monroe"
    
    MsgBox strMessage, vbInformation Or vbOKOnly, "About LCMSWarp"
End Sub

Public Sub StartAlignment()
    Dim lngPMTCount As Long
    Dim i As Long, j As Long
    
    Dim lngObsCountFilterActual
    
    Dim blnUseNETFilter As Boolean
    Dim blnUseMassFilter As Boolean
    
    Dim dblNETMinimum As Double
    Dim dblNETMaximum As Double
    Dim dblMassMinimum As Double
    Dim dblMassMaximum As Double
    
    Dim strFilterMessage As String
    
    mAlignmentFinalizedOrAborted = True
    If mLoading Then Exit Sub
    
On Error GoTo StartAlignmentErrorHandler
    
    If mLocalPMTCount = 0 Or mLocalFeatureCount = 0 Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            If mLocalPMTCount = 0 Then
                MsgBox "MT tags are not present in memory.  Please close this dialog and use menu option 'Steps->3. Select MT tags' in the main window to connect to a database or define a legacy database (Access DB file).", vbExclamation + vbOKOnly, "Error"
            Else
                MsgBox "LC-MS Features are not present in memory.  Please close this dialog and use menu item 'Steps->2. Find LC-MS Features (UMCs)' in the main window to cluster the data into unique mass classes.", vbExclamation + vbOKOnly, "Error"
            End If
        End If
        Exit Sub
    End If
    
    EnableDisableControls False
    mLocalGelUpdated = False
    mAlignmentFinalizedOrAborted = False
    mPauseSplitWarpProcessing = False
    cmdSplitWarpResume.Visible = False
    
    If GelSearchDef(CallerID).MassCalibrationInfo.AdjustmentHistoryCount > 0 And UMCNetAdjDef.RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass Then
        If cChkBox(chkWarpMassAutoRemovePreviousMassCalibrations) Then
            StartMassCalibrationRevert False
        End If
    End If
    
    mMassMatchState = MassMatchProcessingStateConstants.pscRunning
    
    mProcessingStartTime = Now
    tmrAlignment.Enabled = True
    
    ' Clear the surface plot and the 3D plot
    ctlFlatSurface.ClearData
    ctlCWGraphNI.ClearData
    
    ' Clear the Mass and NET error histograms
    graphMassErrors.ClearData
    graphNetErrors.ClearData


    blnUseNETFilter = False
    dblNETMinimum = 0
    dblNETMaximum = 0
    If IsNumeric(txtAMTNetMin) Then dblNETMinimum = CDblSafe(txtAMTNetMin)
    If IsNumeric(txtAMTNetMax) Then dblNETMaximum = CDblSafe(txtAMTNetMax)
    If dblNETMinimum > 0 Or dblNETMaximum > 0 Then
        If dblNETMinimum > 0 And dblNETMaximum = 0 Then
            dblNETMaximum = 10
        End If
        blnUseNETFilter = True
    End If
    
    blnUseMassFilter = False
    dblMassMinimum = 0
    dblMassMaximum = 0
    If IsNumeric(txtAMTMassMin) Then dblMassMinimum = CDblSafe(txtAMTMassMin)
    If IsNumeric(txtAMTMassMax) Then dblMassMaximum = CDblSafe(txtAMTMassMax)
    If dblMassMinimum > 0 Or dblMassMaximum > 0 Then
        If dblMassMinimum > 0 And dblMassMaximum = 0 Then
            dblMassMaximum = 1000000
        End If
        blnUseMassFilter = True
    End If
    
    ' Populate mLocalPMTsFiltered()
    ' First find the number of MT tags passing the observation count filter
    lngPMTCount = 0
    For i = 0 To mLocalPMTCount - 1
        If mLocalPMTs(i, PMTColumnConstants.pccObservationCount) >= UMCNetAdjDef.MSWarpOptions.MinimumPMTTagObsCount Then
            lngPMTCount = lngPMTCount + 1
        End If
    Next i
    
    If lngPMTCount = 0 Then
        ' No MT tags match the filter; use all of them
        lngObsCountFilterActual = 0
    Else
        ' One ore more MT Tags has an observation count >= .MinimumPMTTagObsCount
        lngObsCountFilterActual = UMCNetAdjDef.MSWarpOptions.MinimumPMTTagObsCount
    End If

    ' Now determine the number of MT tags that will pass both the observation count filter and the mass and NET filters
    lngPMTCount = 0
    For i = 0 To mLocalPMTCount - 1
        If MTPassesFilters(i, lngObsCountFilterActual, blnUseNETFilter, dblNETMinimum, dblNETMaximum, blnUseMassFilter, dblMassMinimum, dblMassMaximum) Then
            lngPMTCount = lngPMTCount + 1
        End If
    Next i
    
    strFilterMessage = ""
    If blnUseMassFilter Or blnUseNETFilter Then
        If blnUseNETFilter Then
            strFilterMessage = strFilterMessage & "; NET between " & Round(dblNETMinimum, 3) & " and " & Round(dblNETMaximum, 3)
        End If
        
        If blnUseMassFilter Then
            strFilterMessage = strFilterMessage & "; Mass between " & Round(dblMassMinimum, 3) & " and " & Round(dblMassMaximum, 3)
        End If
    End If
    
    If lngPMTCount = 0 Then
        strFilterMessage = "The specified NET and/or Mass filters excluded all of the MT Tags" & strFilterMessage
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strFilterMessage, vbExclamation + vbOKOnly, "Error"
        Else
            AddToAnalysisHistory CallerID, strFilterMessage
        End If
        EnableDisableControls True
        mAlignmentFinalizedOrAborted = True
        Exit Sub
    Else
        txtPMTCountLoaded.Text = LongToStringWithCommas(lngPMTCount)
    End If
    
    ' Reserve space in mLocalPMTsFiltered for a maximum of lngPMTCount PMTs
    ReDim mLocalPMTsFiltered(lngPMTCount - 1, PMT_COLUMN_COUNT - 1)
    
    ' Now copy the valid data from mLocalPMTs into mLocalPMTsFiltered
    ' We will also apply the mass and NET filters at this time (if applicable)
    lngPMTCount = 0
    For i = 0 To mLocalPMTCount - 1
        If MTPassesFilters(i, lngObsCountFilterActual, blnUseNETFilter, dblNETMinimum, dblNETMaximum, blnUseMassFilter, dblMassMinimum, dblMassMaximum) Then
            For j = 0 To PMT_COLUMN_COUNT - 1
                mLocalPMTsFiltered(lngPMTCount, j) = mLocalPMTs(i, j)
            Next j
            lngPMTCount = lngPMTCount + 1
        End If
    Next i

    ' Verify that mLocalPMTsFiltered was populated the way we expected it to be
    Debug.Assert UBound(mLocalPMTsFiltered, 1) = lngPMTCount - 1
    
    If Len(strFilterMessage) > 0 Then
        strFilterMessage = "AMT Tags used for alignment are filtered" & strFilterMessage
        AddToAnalysisHistory CallerID, strFilterMessage
    End If
    
    ' The following function will filter the features to align (if applicable) then call mMassMatchObject.MS2MSMSDBAlignPeptidesThreaded
    FilterAndAlignFeatures 1

    Exit Sub

StartAlignmentErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error starting alignment: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmMSAlign.StartAlignment"
    End If
       
    mAlignmentFinalizedOrAborted = True
    
End Sub

Private Sub StartMassCalibrationRevert(blnQueryUserToConfirm As Boolean)
    Dim blnDataUpdated As Boolean
    
    UpdateStatus "Reverting mass calibration to original, uncorrected values."
    blnDataUpdated = MassCalibrationRevertToOriginal(CallerID, blnQueryUserToConfirm, True, Me)
  
    If blnDataUpdated Then
        ' Update mLocalFeatures
        PopulateLocalFeaturesArray False
        
        UpdateMassCalibrationStats
        
        UpdateStatus "Finished reverting mass calibration to original, uncorrected values."
    Else
        UpdateStatus "Mass calibration revert not performed."
    End If
    
End Sub

Private Sub ToggleMassVsMZResidualsPlotMode()
    ' If the current plot show is a Mass vs. m/z residuals plot, then toggles to the other Mass vs. m/z residuals plot
    ' Otherwise, examines mMostRecentMassVsmzResidualsViewMode and switches to the mode stored in that variable
    '  provided it is a mass vs. m/z residuals mode
    ' If neither of the above is true, then switches to mode pvmMassResidualsMZPlot
    
    Select Case mMostRecentPlotViewMode
    Case pvmMassResidualsMZPlot
        UpdatePlotViewMode pvmMassResidualsCorrectedMZPlot
    Case pvmMassResidualsCorrectedMZPlot
        UpdatePlotViewMode pvmMassResidualsMZPlot
    Case Else
        If mMostRecentMassVsMZResidualsViewMode = pvmMassResidualsMZPlot Or _
            mMostRecentMassVsMZResidualsViewMode = pvmMassResidualsCorrectedMZPlot Then
            UpdatePlotViewMode mMostRecentMassVsMZResidualsViewMode
        Else
            UpdatePlotViewMode pvmMassResidualsMZPlot
        End If
    End Select
End Sub

Private Sub ToggleMassVsScanResidualsPlotMode()
    ' If the current plot show is a Mass vs. Scan residuals plot, then toggles to the other Mass vs. Scan residuals plot
    ' Otherwise, examines mMostRecentMassVsScanResidualsViewMode and switches to the mode stored in that variable
    '  provided it is a mass vs. scan residuals mode
    ' If neither of the above is true, then switches to mode pvmMassResidualsScanPlot
    
    Select Case mMostRecentPlotViewMode
    Case pvmMassResidualsScanPlot
        UpdatePlotViewMode pvmMassResidualsCorrectedScanPlot
    Case pvmMassResidualsCorrectedScanPlot
        UpdatePlotViewMode pvmMassResidualsScanPlot
    Case Else
        If mMostRecentMassVsScanResidualsViewMode = pvmMassResidualsScanPlot Or _
            mMostRecentMassVsScanResidualsViewMode = pvmMassResidualsCorrectedScanPlot Then
            UpdatePlotViewMode mMostRecentMassVsScanResidualsViewMode
        Else
            UpdatePlotViewMode pvmMassResidualsScanPlot
        End If
    End Select
End Sub

Private Sub ToggleNETResidualsPlotMode()
    ' If the current plot show is a NET residuals plot, then toggles to the other NET residuals plot
    ' Otherwise, examines mMostRecentNETResidualsViewMode and switches to the mode stored in that variable
    '  provided it is a NET residuals mode
    ' If neither of the above is true, then switches to mode pvmLinearFitNETResidualsPlot
    
    Select Case mMostRecentPlotViewMode
    Case pvmLinearFitNETResidualsPlot
        UpdatePlotViewMode pvmWarpedFitNETResidualsPlot
    Case pvmWarpedFitNETResidualsPlot
        UpdatePlotViewMode pvmLinearFitNETResidualsPlot
    Case Else
        If mMostRecentNETResidualsViewMode = pvmLinearFitNETResidualsPlot Or _
            mMostRecentNETResidualsViewMode = pvmWarpedFitNETResidualsPlot Then
            UpdatePlotViewMode mMostRecentNETResidualsViewMode
        Else
            UpdatePlotViewMode pvmLinearFitNETResidualsPlot
        End If
    End Select
End Sub

Private Sub UpdateControlValues(ByVal blnUseDefaults As Boolean)
    
    If blnUseDefaults Then
        SetDefaultUMCNETAdjDef UMCNetAdjDef
    End If
    
    With UMCNetAdjDef
        txtNumSections.Text = .MSWarpOptions.NumberOfSections
        txtMaxDistortion.Text = .MSWarpOptions.MaxDistortion
        txtContractionFactor.Text = .MSWarpOptions.ContractionFactor
        txtMinMSMSObservations.Text = .MSWarpOptions.MinimumPMTTagObsCount
        
        Select Case .MWTolType
        Case gltPPM
            txtMassTolerance.Text = .MWTol
        Case gltABS
            txtMassTolerance.Text = MassToPPM(.MWTol, 1000)
        Case Else
            Debug.Assert False
        End Select

        If .MSWarpOptions.NETTol <= 0 Then .MSWarpOptions.NETTol = 0.02
        txtNetTolerance.Text = Round(.MSWarpOptions.NETTol, 4)
        
        txtMassTagMatchPromiscuity.Text = .MSWarpOptions.MatchPromiscuity
    
        txtSplineOrder.Text = .MSWarpOptions.MassSplineOrder
        txtWarpMassWindowPPM.Text = .MSWarpOptions.MassWindowPPM
    
        txtWarpMassNumXSlices.Text = .MSWarpOptions.MassNumXSlices
        txtWarpMassNumMassDeltaBins.Text = .MSWarpOptions.MassNumMassDeltaBins
        txtWarpMassMaxJump.Text = .MSWarpOptions.MassMaxJump
    
        txtWarpMassZScoreTolerance.Text = .MSWarpOptions.MassZScoreTolerance
        SetCheckBox chkWarpMassUseLSQ, .MSWarpOptions.MassUseLSQ
        txtWarpMassLSQNumKnots.Text = .MSWarpOptions.MassLSQNumKnots
        txtWarpMassLSQOutlierZScore.Text = .MSWarpOptions.MassLSQOutlierZScore
    
        txtAMTMassMin.Text = ""
        txtAMTMassMax.Text = ""
        txtAMTNetMin.Text = ""
        txtAMTNetMax.Text = ""
    
        Select Case .MSWarpOptions.MassCalibrationType
        Case rmcUMCRobustNETWarpMassCalibrationType.rmcMZRegressionRecal
            optMassRecalMZRegression.Value = True
        Case rmcUMCRobustNETWarpMassCalibrationType.rmcScanRegressionRecal
            optMassRecalScanRegression.Value = True
        Case rmcUMCRobustNETWarpMassCalibrationType.rmcHybridRecal
            optMassRecalHybrid.Value = True
        Case Else
            ' Default to Hybrid
            optMassRecalHybrid.Value = True
        End Select
        
        If .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass Then
            cboStepsToPerform.ListIndex = MassMatchStepsToPerformConstants.mmsWarpTimeAndMass
        Else
            cboStepsToPerform.ListIndex = MassMatchStepsToPerformConstants.mmsWarpTime
        End If
        
        Select Case .MSWarpOptions.SplitWarpMode
        Case swmSplitOnMZ
            optSplitWarpOnMZ.Value = True
        Case Else
            ' Includes swmDisabled
            optSplitWarpDisabled.Value = True
        End Select
        
        txtSplitWarpMZBoundary.Text = .MSWarpOptions.SplitWarpMZBoundary
        
    End With
    
    With glbPreferencesExpanded.ErrorPlottingOptions
        If blnUseDefaults Then
            .MassBinSizePPM = DEFAULT_MASS_BIN_SIZE_PPM
            .GANETBinSize = DEFAULT_GANET_BIN_SIZE
        End If
        
        txtMassBinSizePPM = .MassBinSizePPM
        txtGANETBinSize = .GANETBinSize
    End With
    
    SetCheckBox chkWarpMassAutoRemovePreviousMassCalibrations, True
    
End Sub

Private Sub UpdateMassCalibrationStats()
    Dim dblMassShiftPPM As Double
    
    With GelSearchDef(CallerID).MassCalibrationInfo
        Select Case .MassUnits
        Case gltABS
            ' Compute PPM value at 1000 Da
            dblMassShiftPPM = MassToPPM(.OverallMassAdjustment, 1000)
        Case Else
            ' Includes gltPPM
            dblMassShiftPPM = .OverallMassAdjustment
        End Select
        txtMassCalibrationOverallAdjustment.Text = Format(dblMassShiftPPM, "0.0000")
        txtMassCalibrationOverallShiftCount.Text = .AdjustmentHistoryCount
    End With
End Sub

Private Sub UpdatePlotViewMode(ePlotViewMode As pvmPlotViewModeConstants)
    
    Dim blnShowNETandZScoreCheckbox As Boolean
    Dim blnShowTransformCheckbox As Boolean
    Dim ePreviousPlotViewMode As pvmPlotViewModeConstants
    
On Error GoTo UpdatePlotViewModeErrorHandler

    TraceLog 3, "UpdatePlotViewMode", "Old ViewMode: " & mMostRecentPlotViewMode & ", New ViewMode: " & ePlotViewMode
    ePreviousPlotViewMode = mMostRecentPlotViewMode
    SetMostRecentPlotViewMode ePlotViewMode

    ctlFlatSurface.Visible = False
    ctlCWGraphNI.Visible = False
    ctlNETResidual.Visible = False
    ctlMassVsScanResidual.Visible = False
    ctlMassVsMZResidual.Visible = False

    Select Case ePlotViewMode
    Case pvmPlotViewModeConstants.pvmFlatView
        blnShowNETandZScoreCheckbox = True
        blnShowTransformCheckbox = True
        cmdZoomOutResidualsPlot.Enabled = False
        
        ctlFlatSurface.Visible = True
        ctlFlatSurface.RefreshPlotNow
        
    Case pvmPlotViewModeConstants.pvm2DView, pvmPlotViewModeConstants.pvm3DView
        blnShowNETandZScoreCheckbox = False
        blnShowTransformCheckbox = False
        cmdZoomOutResidualsPlot.Enabled = False
        
        ctlCWGraphNI.Visible = True
        
        If ePlotViewMode = pvmPlotViewModeConstants.pvm2DView Then
            With ctlCWGraphNI
                .ViewDistance = 0.75
                .ViewLatitude = 0
                .ViewLongitude = -270
            End With
        Else
            With ctlCWGraphNI
                ctlCWGraphNI.ViewDistance = 0.85
                ctlCWGraphNI.ViewLatitude = 45
                ctlCWGraphNI.ViewLongitude = -290
            End With
        End If
    Case pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot, pvmPlotViewModeConstants.pvmWarpedFitNETResidualsPlot
        blnShowNETandZScoreCheckbox = False
        If ePlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
            blnShowTransformCheckbox = True
        Else
            blnShowTransformCheckbox = False
        End If
        cmdZoomOutResidualsPlot.Enabled = True
        
        ctlNETResidual.Visible = True
        
        TraceLog 3, "UpdatePlotViewMode", "Call UpdateNETResidualsPlot"
        UpdateNETResidualsPlot False
            
    Case pvmPlotViewModeConstants.pvmMassResidualsScanPlot, pvmPlotViewModeConstants.pvmMassResidualsCorrectedScanPlot
        blnShowNETandZScoreCheckbox = False
        If ePlotViewMode = pvmPlotViewModeConstants.pvmMassResidualsScanPlot Then
            blnShowTransformCheckbox = True
        Else
            blnShowTransformCheckbox = False
        End If
        cmdZoomOutResidualsPlot.Enabled = True
        
        ctlMassVsScanResidual.Visible = True
        
        TraceLog 3, "UpdatePlotViewMode", "Call UpdateMassResidualsScanPlot"
        UpdateMassResidualsScanPlot False
    
    Case pvmPlotViewModeConstants.pvmMassResidualsMZPlot, pvmPlotViewModeConstants.pvmMassResidualsCorrectedMZPlot
        blnShowNETandZScoreCheckbox = False
        If ePlotViewMode = pvmPlotViewModeConstants.pvmMassResidualsMZPlot Then
            blnShowTransformCheckbox = True
        Else
            blnShowTransformCheckbox = False
        End If
        cmdZoomOutResidualsPlot.Enabled = True
        
        ctlMassVsMZResidual.Visible = True
        
        TraceLog 3, "UpdatePlotViewMode", "Call UpdateMassResidualsMZPlot"
        UpdateMassResidualsMZPlot False
        
    Case Else
        ' Unknown mode
        Debug.Assert False
    End Select
        
    chkShowNet.Visible = blnShowNETandZScoreCheckbox
    chkShowTransform.Visible = blnShowTransformCheckbox
    chkSurfaceShowsZScore.Visible = blnShowNETandZScoreCheckbox

    Exit Sub

UpdatePlotViewModeErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in UpdatePlotViewMode: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        LogErrors Err.Number, "frmMSAlign.UpdatePlotViewMode"
    End If

End Sub

Public Sub UpdatePlotViewModeToMassVsMZResiduals()
    UpdatePlotViewMode pvmMassResidualsMZPlot
End Sub
Public Sub UpdatePlotViewModeToMassVsMZCorrectedResiduals()
    UpdatePlotViewMode pvmMassResidualsCorrectedMZPlot
End Sub

Public Sub UpdatePlotViewModeToMassVsScanResiduals()
    UpdatePlotViewMode pvmMassResidualsScanPlot
End Sub
Public Sub UpdatePlotViewModeToMassVsScanCorrectedResiduals()
    UpdatePlotViewMode pvmMassResidualsCorrectedScanPlot
End Sub

Public Sub UpdatePlotViewModeToFlatView()
    UpdatePlotViewMode pvmFlatView
End Sub

Public Sub UpdatePlotViewModeTo2DView()
    UpdatePlotViewMode pvm2DView
End Sub

Public Sub UpdatePlotViewModeTo3DView()
    UpdatePlotViewMode pvm3DView
End Sub

Public Sub UpdatePlotViewModeToLinearFitNETResidualsPlot()
    UpdatePlotViewMode pvmLinearFitNETResidualsPlot
End Sub

Public Sub UpdatePlotViewModeToWarpedFitNETResidualsPlot()
    UpdatePlotViewMode pvmWarpedFitNETResidualsPlot
End Sub

Private Sub UpdateResidualPlots()
    UpdateNETResidualsPlot False
    UpdateMassResidualsScanPlot False
    UpdateMassResidualsMZPlot False
End Sub

Private Sub UpdateStatus(strNewStatus As String)
    lblStatus = strNewStatus
    DoEvents
End Sub

Private Sub ValidateNETWarpMassMaxJump()
    Dim intMaxValue As Integer
    intMaxValue = CIntSafe(txtWarpMassNumMassDeltaBins.Text)
    If intMaxValue < 2 Then intMaxValue = 2
    
    UMCNetAdjDef.MSWarpOptions.MassMaxJump = ValidateTextboxValueLng(txtWarpMassMaxJump, 0, CLng(intMaxValue), intMaxValue / 2)
End Sub

Private Function ValidateNotInfinity(ByVal dblValue As Double) As Double
    On Error Resume Next
    If dblValue = 0 Then
        If Err.Number <> 0 Then
            Err.Clear
            dblValue = 0
        End If
    End If
    ValidateNotInfinity = dblValue
End Function

Private Sub ZoomOutCWGraph(ByRef objPlot As CWGraph)
    On Error Resume Next
    objPlot.Plots(1).XAxis.AutoScaleNow
    objPlot.Plots(1).YAxis.AutoScaleNow
End Sub

Private Sub ZoomOutResidualsPlot(blnZoomOutAll As Boolean)
    If blnZoomOutAll Then
        ZoomOutNETResidualsPlot
        ZoomOutMassVsScanResidualsPlot
        ZoomOutMassVsMZResidualsPlot
    Else
        Select Case mMostRecentPlotViewMode
        Case pvmLinearFitNETResidualsPlot, pvmWarpedFitNETResidualsPlot
            ZoomOutNETResidualsPlot
        Case pvmMassResidualsScanPlot, pvmMassResidualsCorrectedScanPlot
            ZoomOutMassVsScanResidualsPlot
        Case pvmMassResidualsMZPlot, pvmMassResidualsCorrectedMZPlot
            ZoomOutMassVsMZResidualsPlot
        Case Else
            ' Residuals plot is not visible; ignore this
        End Select
    End If

    DisplayResidualPlotAxisRangesCurrent
End Sub

Private Sub ZoomOutNETResidualsPlot()
    Dim sngNETWarpNETTolMinimum As Single
    
    '------------------------------------
    ' NET Residuals Plot
    '------------------------------------
    
    ' Autoscale the X Axis of the NET Residuals plot
    With ctlNETResidual.Plots(1).XAxis
        .AutoScale = False
        .Minimum = mMinScan
        .Maximum = mMaxScan
    End With
    
    ' Assure that the Y axis range on ctlNETResidual stays constant between modes pvmLinearFitNETResidualsPlot and pvmWarpedFitNETResidualsPlot
    sngNETWarpNETTolMinimum = Abs(UMCNetAdjDef.MSWarpOptions.NETTol)
    If sngNETWarpNETTolMinimum < 0.005 Then sngNETWarpNETTolMinimum = 0.005
    
    With ctlNETResidual.Plots(1).YAxis
        .AutoScale = False
        If mLinearNETResidualMin <> 0 And mLinearNETResidualMax <> 0 Then
            If mLinearNETResidualMin < -sngNETWarpNETTolMinimum Then
                If mLinearNETResidualMin > -(sngNETWarpNETTolMinimum * 10) Or mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
                    .Minimum = mLinearNETResidualMin
                Else
                    .Minimum = -(sngNETWarpNETTolMinimum * 10)
                End If
            Else
                .Minimum = -sngNETWarpNETTolMinimum
            End If
            If mLinearNETResidualMax > sngNETWarpNETTolMinimum Then
                If mLinearNETResidualMin < sngNETWarpNETTolMinimum * 10 Or mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
                    .Maximum = mLinearNETResidualMax
                Else
                    .Maximum = sngNETWarpNETTolMinimum * 10
                End If
            Else
                .Maximum = sngNETWarpNETTolMinimum
            End If
        Else
            .Minimum = -sngNETWarpNETTolMinimum
            .Maximum = sngNETWarpNETTolMinimum
        End If
        '.AutoScaleNow
    End With

End Sub

Private Sub ZoomOutMassVsScanResidualsPlot()
    Dim sngMassWarpMassTolMinimum As Single

    '------------------------------------
    ' Mass vs. Scan Residuals Plot
    '------------------------------------
    
    ' Autoscale the X Axis of the Mass vs. Scan Residuals plot
    With ctlMassVsScanResidual.Plots(1).XAxis
        .AutoScale = False
        .Minimum = mMinScan
        .Maximum = mMaxScan
    End With

    ' Assure that the Y axis range on ctlMassVsScanResidual stays constant between modes pvmMassResidualsScanPlot and pvmMassResidualsCorrectedScanPlot
    If UMCNetAdjDef.MWTolType = glMassToleranceConstants.gltPPM Then
        sngMassWarpMassTolMinimum = UMCNetAdjDef.MWTol
    Else
        sngMassWarpMassTolMinimum = 20
    End If
    
    With ctlMassVsScanResidual.Plots(1).YAxis
       .AutoScale = False
        If mMassResidualMin <> 0 And mMassResidualMax <> 0 Then
            If mMassResidualMin < -sngMassWarpMassTolMinimum Then
                If mMassResidualMin > -(sngMassWarpMassTolMinimum * 3) Or mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
                    .Minimum = mMassResidualMin
                Else
                    .Minimum = -(sngMassWarpMassTolMinimum * 3)
                End If
            Else
                .Minimum = -sngMassWarpMassTolMinimum
            End If
            If mMassResidualMax > sngMassWarpMassTolMinimum Then
                If mMassResidualMin < sngMassWarpMassTolMinimum * 3 Or mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
                    .Maximum = mMassResidualMax
                Else
                    .Maximum = sngMassWarpMassTolMinimum * 3
                End If
            Else
                .Maximum = sngMassWarpMassTolMinimum
            End If
        Else
            .Minimum = -sngMassWarpMassTolMinimum
            .Maximum = sngMassWarpMassTolMinimum
        End If
        '.AutoScaleNow
    End With

End Sub

Private Sub ZoomOutMassVsMZResidualsPlot()
    Dim sngMassWarpMassTolMinimum As Single
    
    '------------------------------------
    ' Mass vs. m/z Residuals Plot
    '------------------------------------
    
    ' Autoscale the X Axis of the Mass vs. m/z Residuals plot
    With ctlMassVsMZResidual.Plots(1).XAxis
        .AutoScale = False
        .Minimum = mMinMZ
        .Maximum = mMaxMZ
    End With

    ' Assure that the Y axis range on ctlMassVsMZResidual stays constant between modes pvmMassResidualsMZPlot and pvmMassResidualsCorrectedMZPlot
    If UMCNetAdjDef.MWTolType = glMassToleranceConstants.gltPPM Then
        sngMassWarpMassTolMinimum = UMCNetAdjDef.MWTol
    Else
        sngMassWarpMassTolMinimum = 20
    End If
    
    With ctlMassVsMZResidual.Plots(1).YAxis
       .AutoScale = False
        If mMassResidualMin <> 0 And mMassResidualMax <> 0 Then
            If mMassResidualMin < -sngMassWarpMassTolMinimum Then
                If mMassResidualMin > -(sngMassWarpMassTolMinimum * 3) Or mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
                    .Minimum = mMassResidualMin
                Else
                    .Minimum = -(sngMassWarpMassTolMinimum * 3)
                End If
            Else
                .Minimum = -sngMassWarpMassTolMinimum
            End If
            If mMassResidualMax > sngMassWarpMassTolMinimum Then
                If mMassResidualMin < sngMassWarpMassTolMinimum * 3 Or mMostRecentPlotViewMode = pvmPlotViewModeConstants.pvmLinearFitNETResidualsPlot Then
                    .Maximum = mMassResidualMax
                Else
                    .Maximum = sngMassWarpMassTolMinimum * 3
                End If
            Else
                .Maximum = sngMassWarpMassTolMinimum
            End If
        Else
            .Minimum = -sngMassWarpMassTolMinimum
            .Maximum = sngMassWarpMassTolMinimum
        End If
        '.AutoScaleNow
    End With
    
End Sub

Private Sub cboResidualPlotPointSize_Click()
    mResidualPlotPointSize = cboResidualPlotPointSize.ListIndex + 1
    If mResidualPlotPointSize < 0 Then mResidualPlotPointSize = 0
    UpdateResidualPlots
End Sub

Private Sub cboResidualPlotTransformationFnLineSize_Click()
    mResidualPlotTransformationFnLineSize = cboResidualPlotTransformationFnLineSize.ListIndex
    If mResidualPlotTransformationFnLineSize < 0 Then mResidualPlotTransformationFnLineSize = 0
    UpdateResidualPlots
End Sub

Private Sub cboStepsToPerform_Click()
    EnableDisableControls mControlsEnabled
    If Not mLoading Then
        Select Case cboStepsToPerform.ListIndex
        Case mmsWarpTimeAndMass
            UMCNetAdjDef.RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass
        Case Else
            UMCNetAdjDef.RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTime
        End Select
    End If
End Sub

Private Sub chkShowNet_Click()
    PlotNETLine
End Sub

Private Sub chkShowTransform_Click()
    ShowHideTransformLine cChkBox(chkShowTransform)
End Sub

Private Sub chkSurfaceShowsZScore_Click()
    PlotSurfaceData
End Sub

Private Sub chkWarpMassUseLSQ_Click()
    UMCNetAdjDef.MSWarpOptions.MassUseLSQ = cChkBox(chkWarpMassUseLSQ)
End Sub

Private Sub cmdAbort_Click()
    RequestAbort
End Sub

Private Sub cmdEditShowDefaults_Click()
    SetDefaultOptions
End Sub

Private Sub cmdMassCalibrationRevert_Click()
    StartMassCalibrationRevert True
End Sub

Private Sub cmdResidualPlotSetRange_Click()
    ResidualPlotSetRange
End Sub

Private Sub cmdSetDefaults_Click()
    SetDefaultOptions
End Sub

Private Sub cmdSplitWarpResume_Click()
    mPauseSplitWarpProcessing = False
End Sub

'Private Sub cmdMassCalib_Click()
'    Dim num_sections As Long, contraction_factor As Integer, max_discontinuity_jump As Integer
'    Dim mass_tolerance As Double, net_tolerance As Double
'    Dim dblNETMin As Double, dblNETMax As Double
'    Dim scores() As Double
'    Dim i As Long
'
'    num_sections = Val(txtNumSections)
'    contraction_factor = Val(txtContractionFactor)
'    max_discontinuity_jump = Val(txtMaxDistortion)
'    mass_tolerance = Val(txtMassTolerance)
'    net_tolerance = Val(txtNetTolerance)
'    dblNETMin = 0
'    dblNETMax = 1
'
'    mMassMatchState = MassMatchProcessingStateConstants.pscMassCalibRunning
'    mProcessingStartTime = Now
'    tmrAlignment.Enabled = True
'
'    graphMassErrors.ClearData
'
'    Call mMassMatchObject.SetOptions(num_sections, contraction_factor, max_discontinuity_jump, mass_tolerance, net_tolerance, dblNETMin, dblNETMax)
'    Call mMassMatchObject.CalculateMS2MSMSDBMassCalibrationThreaded(mLocalFeatures, mLocalPMTs)
'
'End Sub

Private Sub cmdWarpAlign_Click()
    StartAlignment
End Sub

Private Sub cmdZoomOutResidualsPlot_Click()
    ZoomOutResidualsPlot False
End Sub

Private Sub ctlMassVsMZResidual_Zoom()
    DisplayResidualPlotAxisRanges ctlMassVsMZResidual, 0, 1
End Sub

Private Sub ctlMassVsScanResidual_Zoom()
    DisplayResidualPlotAxisRanges ctlMassVsScanResidual, 0, 1
End Sub

Private Sub ctlNetResidual_Zoom()
    DisplayResidualPlotAxisRanges ctlNETResidual, 0, 3
End Sub

Private Sub Form_Activate()
    InitializeSearch
End Sub

Public Sub Form_Load()
    InitializeForm
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GelUMCNETAdjDef(CallerID) = UMCNetAdjDef
End Sub

Private Sub graphMassErrors_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        ZoomOutCWGraph graphMassErrors
    End If
End Sub

Private Sub graphMassErrors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        ZoomOutCWGraph graphMassErrors
    End If
End Sub

Private Sub graphNetErrors_KeyPress(KeyAscii As Integer)
  If KeyAscii = 8 Then
       ZoomOutCWGraph graphNetErrors
    End If
End Sub

Private Sub graphNetErrors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
         ZoomOutCWGraph graphNetErrors
    End If
End Sub

Private Sub mnuEditAlignStart_Click()
    StartAlignment
End Sub

Private Sub mnuEditCopy3DViewToClipboard_Click()
    Save3DViewToClipboardOrEMF ""
End Sub

Private Sub mnuEditCopyAlignmentFunctionToClipboard_Click()
    CopyAlignmentFunctionToClipboardOrFile
End Sub

Private Sub mnuEditCopyAlignmentScoresToClipboard_Click()
    CopyAlignmentScoresToClipboardOrFile
End Sub

Private Sub mnuEditCopyFlatViewToClipboard_Click()
    SaveFlatViewToClipboardOrEMF ""
End Sub

Private Sub mnuEditCopylMassVsMZResidualValues_Click()
    CopyResidualMassVsMZValuesToClipboardOrFile
End Sub

Private Sub mnuEditCopyMassMassVsScanResidualValues_Click()
    CopyResidualMassVsScanValuesToClipboardOrFile
End Sub

Private Sub mnuEditCopyMassVsmzResidualsPlotToClipboard_Click()
    SaveMassVsMZResidualsPlotToClipboardOrEMF ""
End Sub

Private Sub mnuEditCopyMassVsScanResidualsPlotToClipboard_Click()
    SaveMassVsScanResidualsPlotToClipboardOrEMF ""
End Sub

Private Sub mnuEditCopyNETResidualValues_Click()
    CopyResidualNETValuesToClipboardOrFile
End Sub

Private Sub mnuEditCopyResidualsPlotToClipboard_Click()
    SaveNETResidualsPlotToClipboardOrEMF ""
End Sub

Private Sub mnuEditLoadPMTsFromFile_Click()
    LoadPMTsFromFile
End Sub

Private Sub mnuEditLoadLCMSFeaturesFromFile_Click()
    LoadLCMSFeaturesFromFile
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave3DViewAsEMF_Click()
    Save3DViewToFile pftEMF
End Sub

Private Sub mnuFileSave3DViewAsPNG_Click()
    Save3DViewToFile pftPNG
End Sub

Private Sub mnuFileSaveFlatViewAsEMF_Click()
    SaveFlatViewToFile pftPictureFileTypeConstants.pftEMF
End Sub

Private Sub mnuFileSaveFlatViewAsPNG_Click()
    SaveFlatViewToFile pftPictureFileTypeConstants.pftPNG
End Sub

Private Sub mnuFileSaveMassVsMZResidualsPlotAsEMF_Click()
    SaveMassVsMZResidualsPlotToFile pftEMF
End Sub

Private Sub mnuFileSaveMassVsMZResidualsPlotAsPNG_Click()
    SaveMassVsMZResidualsPlotToFile pftPNG
End Sub

Private Sub mnuFileSaveMassVsScanResidualsPlotAsEMF_Click()
    SaveMassVsScanResidualsPlotToFile pftEMF
End Sub

Private Sub mnuFileSaveMassVsScanResidualsPlotAsPNG_Click()
    SaveMassVsScanResidualsPlotToFile pftPNG
End Sub

Private Sub mnuFileSaveResidualsPlotAsEMF_Click()
    SaveNETResidualsPlotToFile pftEMF
End Sub

Private Sub mnuFileSaveResidualsPlotAsPNG_Click()
    SaveNETResidualsPlotToFile pftPNG
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAboutBox
End Sub

Private Sub mnuPMTTagsLoadFromDB_Click()
    '------------------------------------------------------------
    'load/reload MT tags
    '------------------------------------------------------------
    If Not GelAnalysis(CallerID) Is Nothing Then
        Call LoadPMTsFromDB(True, False)
    Else
        WarnUserNotConnectedToDB CallerID, True
    End If
End Sub

Private Sub mnuPMTTagsLoadFromLegacyDB_Click()
    LoadPMTsFromLegacyDB
End Sub

Private Sub mnuPMTTagsLoadMTStats_Click()
    '------------------------------------------------------------
    'load/reload MT tags
    '------------------------------------------------------------
    If Not GelAnalysis(CallerID) Is Nothing Then
        Call LoadPMTsFromDB(True, True)
    Else
        WarnUserNotConnectedToDB CallerID, True
    End If
End Sub

Private Sub mnuPMTTagsStatus_Click()
    '----------------------------------------------
    'displays short MT tags statistics, it might
    'help with determining problems with MT tags
    '----------------------------------------------
    MsgBox CheckMassTags(), vbOKOnly
End Sub

Private Sub mnuView2DView_Click()
    UpdatePlotViewModeTo2DView
End Sub

Private Sub mnuView3DView_Click()
    UpdatePlotViewModeTo3DView
End Sub

Private Sub mnuViewFlatFiew_Click()
    UpdatePlotViewModeToFlatView
End Sub

Private Sub mnuViewMassVsMZResidualsPlot_Click()
    ToggleMassVsMZResidualsPlotMode
End Sub

Private Sub mnuViewMassVsScanResidualsPlot_Click()
    ToggleMassVsScanResidualsPlotMode
End Sub

Private Sub mnuViewNETResidualsPlot_Click()
    ToggleNETResidualsPlotMode
End Sub

Private Sub mnuWindowSizeToA_Click()
    SetWindowSize fwsFormWindowSizeConstants.fwsSizeA
End Sub

Private Sub mnuWindowSizeToB_Click()
    SetWindowSize fwsFormWindowSizeConstants.fwsSizeB
End Sub

Private Sub mnuWindowSizeToC_Click()
    SetWindowSize fwsFormWindowSizeConstants.fwsSizeC
End Sub

Private Sub mnuWindowSizeToD_Click()
    SetWindowSize fwsFormWindowSizeConstants.fwsSizeD
End Sub

Private Sub optMassRecalHybrid_Click()
    If Not mLoading Then
        UMCNetAdjDef.MSWarpOptions.MassCalibrationType = rmcUMCRobustNETWarpMassCalibrationType.rmcHybridRecal
    End If
End Sub

Private Sub optMassRecalMZRegression_Click()
    If Not mLoading Then
        UMCNetAdjDef.MSWarpOptions.MassCalibrationType = rmcUMCRobustNETWarpMassCalibrationType.rmcMZRegressionRecal
    End If
End Sub

Private Sub optMassRecalScanRegression_Click()
    If Not mLoading Then
        UMCNetAdjDef.MSWarpOptions.MassCalibrationType = rmcUMCRobustNETWarpMassCalibrationType.rmcScanRegressionRecal
    End If
End Sub

Private Sub optSplitWarpDisabled_Click()
    If Not mLoading Then
        UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmDisabled
    End If
End Sub

Private Sub optSplitWarpOnMZ_Click()
    If Not mLoading Then
        UMCNetAdjDef.MSWarpOptions.SplitWarpMode = swmSplitOnMZ
        If UMCNetAdjDef.MSWarpOptions.SplitWarpMZBoundary <= 0 Then
            ' Use the m/z boundary that is appropriate for the Thermo Exactive instrument
            UMCNetAdjDef.MSWarpOptions.SplitWarpMZBoundary = 505.7
            txtSplitWarpMZBoundary.Text = UMCNetAdjDef.MSWarpOptions.SplitWarpMZBoundary
        End If
    End If
End Sub

Private Sub tmrAlignment_Timer()
    QueryMassMatchProgress
End Sub

Private Sub txtAMTMassMax_LostFocus()
    If Len(txtAMTMassMax) > 0 Then
        ValidateTextboxValueDbl txtAMTMassMax, 0, 1000000, -1000
        If txtAMTMassMax = "-1000" Then
            txtAMTMassMax = ""
        End If
    End If
End Sub

Private Sub txtAMTMassMin_LostFocus()
    If Len(txtAMTMassMin) > 0 Then
        ValidateTextboxValueDbl txtAMTMassMin, 0, 1000000, -1000
        If txtAMTMassMin = "-1000" Then
            txtAMTMassMin = ""
        End If
    End If
End Sub

Private Sub txtAMTNetMax_LostFocus()
    If Len(txtAMTNetMax) > 0 Then
        ValidateTextboxValueDbl txtAMTNetMax, -100, 100, -1000
        If txtAMTNetMax = "-1000" Then
            txtAMTNetMax = ""
        End If
    End If
End Sub

Private Sub txtAMTNetMin_LostFocus()
    If Len(txtAMTNetMin) > 0 Then
        ValidateTextboxValueDbl txtAMTNetMin, -100, 100, -1000
        If txtAMTNetMin = "-1000" Then
            txtAMTNetMin = ""
        End If
    End If
End Sub

Private Sub txtContractionFactor_LostFocus()
    UMCNetAdjDef.MSWarpOptions.ContractionFactor = ValidateTextboxValueLng(txtContractionFactor, 1, 10, 3)
End Sub

Private Sub txtGANETBinSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGANETBinSize, KeyAscii, True, True
End Sub

Private Sub txtGANETBinSize_LostFocus()
    glbPreferencesExpanded.ErrorPlottingOptions.GANETBinSize = ValidateTextboxValueDbl(txtGANETBinSize, 0.00001, 5, DEFAULT_GANET_BIN_SIZE)
End Sub

Private Sub txtIntercept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And chkShowNet.Value = vbChecked Then PlotNETLine
End Sub

Private Sub txtIntercept_LostFocus()
    ValidateTextboxValueDbl txtIntercept, -10, 10, 0
    PossiblyUpdateNETLinePlot
End Sub

Private Sub txtMassBinSizePPM_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMassBinSizePPM, KeyAscii, True, True
End Sub

Private Sub txtMassBinSizePPM_LostFocus()
    glbPreferencesExpanded.ErrorPlottingOptions.MassBinSizePPM = ValidateTextboxValueDbl(txtMassTolerance, 0.01, 10000, DEFAULT_MASS_BIN_SIZE_PPM)
End Sub

Private Sub txtMassTagMatchPromiscuity_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MatchPromiscuity = ValidateTextboxValueLng(txtMassTagMatchPromiscuity, 1, 1000, 2)
End Sub

Private Sub txtMassTolerance_LostFocus()
    UMCNetAdjDef.MWTol = ValidateTextboxValueDbl(txtMassTolerance, 0.0001, 10000, 20)
    UMCNetAdjDef.MWTolType = glMassToleranceConstants.gltPPM
End Sub

Private Sub txtMaxDistortion_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MaxDistortion = ValidateTextboxValueLng(txtMaxDistortion, 1, 200, 10)
End Sub

Private Sub txtMinMSMSObservations_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MinimumPMTTagObsCount = ValidateTextboxValueLng(txtMinMSMSObservations, 0, 1000000, 5)
End Sub

Private Sub txtSplitWarpMZBoundary_LostFocus()
    UMCNetAdjDef.MSWarpOptions.SplitWarpMZBoundary = CSng(ValidateTextboxValueDbl(txtSplitWarpMZBoundary, 0, 100000, 0))
End Sub

Private Sub txtSplitWarpMZBoundary_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtSplitWarpMZBoundary, KeyAscii, True, True, False
End Sub

Private Sub txtNetTolerance_LostFocus()
    UMCNetAdjDef.MSWarpOptions.NETTol = ValidateTextboxValueDbl(txtNetTolerance, 0.000001, MAX_NET_TOL, 0.02)
End Sub

Private Sub txtNumSections_LostFocus()
    UMCNetAdjDef.MSWarpOptions.NumberOfSections = ValidateTextboxValueLng(txtNumSections, 5, 500, 100)
End Sub

Private Sub txtResidualPlotMaxX_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ResidualPlotSetRange
End Sub

Private Sub txtResidualPlotMaxX_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtResidualPlotMaxX, KeyAscii, True, True, True
End Sub

Private Sub txtResidualPlotMaxY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ResidualPlotSetRange
End Sub

Private Sub txtResidualPlotMaxY_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtResidualPlotMaxY, KeyAscii, True, True, True
End Sub

Private Sub txtResidualPlotMinX_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ResidualPlotSetRange
End Sub

Private Sub txtResidualPlotMinX_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtResidualPlotMinX, KeyAscii, True, True, True
End Sub

Private Sub txtResidualPlotMinY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ResidualPlotSetRange
End Sub

Private Sub txtResidualPlotMinY_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtResidualPlotMinY, KeyAscii, True, True, True
End Sub

Private Sub txtSlope_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And chkShowNet.Value = vbChecked Then PlotNETLine
End Sub

Private Sub txtSlope_LostFocus()
    ValidateTextboxValueDbl txtSlope, 0.000001, 0.2, 0.001
    PossiblyUpdateNETLinePlot
End Sub

Private Sub txtSplineOrder_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MassSplineOrder = ValidateTextboxValueLng(txtSplineOrder, 0, 3, 2)
End Sub

Private Sub txtWarpMassLSQOutlierZScore_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MassLSQOutlierZScore = ValidateTextboxValueDbl(txtWarpMassLSQOutlierZScore, 0.1, 10, 3)
End Sub

Private Sub txtWarpMassNumMassDeltaBins_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MassNumMassDeltaBins = ValidateTextboxValueLng(txtWarpMassNumMassDeltaBins, 1, 200, 100)
    ValidateNETWarpMassMaxJump
End Sub

Private Sub txtWarpMassMaxJump_LostFocus()
    ValidateNETWarpMassMaxJump
End Sub

Private Sub txtWarpMassLSQNumKnots_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MassLSQNumKnots = ValidateTextboxValueLng(txtWarpMassLSQNumKnots, 1, 50, 12)
End Sub

Private Sub txtWarpMassNumXSlices_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MassNumXSlices = ValidateTextboxValueLng(txtWarpMassNumXSlices, 0, 50, 20)
End Sub

Private Sub txtWarpMassWindowPPM_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MassWindowPPM = ValidateTextboxValueDbl(txtWarpMassWindowPPM, 1, 10000, 50)
End Sub

Private Sub txtWarpMassZScoreTolerance_LostFocus()
    UMCNetAdjDef.MSWarpOptions.MassZScoreTolerance = ValidateTextboxValueDbl(txtWarpMassZScoreTolerance, 0.1, 10, 3)
End Sub

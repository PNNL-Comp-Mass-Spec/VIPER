VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmErrorDistribution2DLoadedData 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tolerance Refinement (Mass and GANET Error Plots)"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraControlsAndPlotContainer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6200
      Left            =   5880
      TabIndex        =   51
      Top             =   0
      Width           =   7000
      Begin VB.TextBox txtMaximumAbundance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         TabIndex        =   59
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtMinimumSLiC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         TabIndex        =   57
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkUseUMCClassStats 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use UMC class stats"
         Height          =   255
         Left            =   3720
         TabIndex        =   55
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chkIncludeInternalStandards 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include Internal Standard matches"
         Height          =   255
         Left            =   3720
         TabIndex        =   54
         Top             =   0
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Frame fraOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Options"
         Height          =   2295
         Left            =   -120
         TabIndex        =   64
         Top             =   1200
         Width           =   6615
         Begin VB.CheckBox chkShowPeakEdges 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Peak Edges"
            Height          =   255
            Left            =   2520
            TabIndex        =   71
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtMassRangePPM 
            Height          =   285
            Left            =   2040
            TabIndex        =   77
            Text            =   "100"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtButterworthFrequency 
            Height          =   285
            Left            =   2760
            TabIndex        =   87
            Text            =   "0.15"
            Top             =   1920
            Width           =   735
         End
         Begin VB.CheckBox chkShowSmoothedData 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Smoothed Data"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkCenterYAxis 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Center Y Axis"
            Height          =   255
            Left            =   2520
            TabIndex        =   70
            Top             =   480
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkDrawLinesBetweenPoints 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Connect Points with Line"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   720
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkShowGridlines 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Gridlines"
            Height          =   255
            Left            =   2520
            TabIndex        =   69
            Top             =   240
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtGANETRange 
            Height          =   285
            Left            =   2040
            TabIndex        =   83
            Text            =   "0.3"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtGANETBinSize 
            Height          =   285
            Left            =   5280
            TabIndex        =   85
            Text            =   "0.005"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtMassBinSizePPM 
            Height          =   285
            Left            =   5280
            TabIndex        =   80
            Text            =   "0.5"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtGraphPointSize 
            Height          =   285
            Left            =   5280
            TabIndex        =   73
            Text            =   "2"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtGraphLineWidth 
            Height          =   285
            Left            =   5280
            TabIndex        =   75
            Text            =   "3"
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chkAutoScaleXRange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Auto Scale X Range"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkShowPointSymbols 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Point Symbols"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   480
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.Label lblButterworthFrequency 
            BackStyle       =   0  'Transparent
            Caption         =   "Butterworth Sampling Frequency"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   1950
            Width           =   2415
         End
         Begin VB.Label lblMassRangePPM 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Range (� 0)"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   1230
            Width           =   1575
         End
         Begin VB.Label lblMassRangeUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "ppm"
            Height          =   255
            Left            =   2880
            TabIndex        =   78
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label lblGANETRange 
            BackStyle       =   0  'Transparent
            Caption         =   "GANET Range (� 0)"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   1590
            Width           =   1575
         End
         Begin VB.Label lblGANETBinSize 
            BackStyle       =   0  'Transparent
            Caption         =   "GANET Bin Size"
            Height          =   255
            Left            =   3840
            TabIndex        =   84
            Top             =   1590
            Width           =   1335
         End
         Begin VB.Label lblMassBinSizePPMUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "ppm"
            Height          =   255
            Left            =   6120
            TabIndex        =   81
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label lblMassBinSizePPM 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Bin Size"
            Height          =   255
            Left            =   3840
            TabIndex        =   79
            Top             =   1230
            Width           =   1335
         End
         Begin VB.Label lblGraphPointSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Point Size"
            Height          =   255
            Left            =   4080
            TabIndex        =   72
            Top             =   270
            Width           =   855
         End
         Begin VB.Label lblGraphLineWidth 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Width"
            Height          =   255
            Left            =   4080
            TabIndex        =   74
            Top             =   630
            Width           =   975
         End
      End
      Begin VB.CheckBox chkShowToleranceRefinementControls 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Tolerance Refinement Controls"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   400
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.ComboBox cboErrorDisplayMode 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   0
         Width           =   3015
      End
      Begin VIPER.ctlSpectraPlotter ctlPlotter 
         Height          =   4815
         Left            =   0
         TabIndex        =   63
         Top             =   1320
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8493
      End
      Begin VB.Label lblMaximumAbundance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Abundance"
         Height          =   195
         Left            =   3840
         TabIndex        =   58
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblMinimumSLiC 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum SLiC"
         Height          =   195
         Left            =   4320
         TabIndex        =   56
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblMTStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         Height          =   495
         Left            =   1080
         TabIndex        =   61
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblMTDBSTatus 
         BackStyle       =   0  'Transparent
         Caption         =   "MTDB Status:"
         Height          =   255
         Left            =   0
         TabIndex        =   60
         ToolTipText     =   "Status of the MT tag database"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Ready"
         Height          =   255
         Left            =   0
         TabIndex        =   62
         Top             =   960
         Width           =   3735
      End
   End
   Begin VB.Frame fraToleranceRefinementContainer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Frame fraRelativeRisk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Relative Risk Statistics"
         Height          =   735
         Left            =   120
         TabIndex        =   95
         Top             =   6120
         Width           =   3015
         Begin VB.TextBox txtRelativeRiskStatistics 
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   96
            Top             =   220
            Width           =   2835
         End
      End
      Begin VB.Frame fraUMCMassStatistics 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UMC Mass Statistics"
         Height          =   1395
         Left            =   3240
         TabIndex        =   88
         Top             =   4680
         Width           =   2550
         Begin VB.TextBox txtUMCMassStatistics 
            BorderStyle     =   0  'None
            Height          =   1020
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   89
            Text            =   "frmErrorDistribution2D.frx":0000
            Top             =   280
            Width           =   2355
         End
      End
      Begin VB.Frame fraCurrentDBSearchMassTolerances 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current DB Search Tolerances"
         Height          =   745
         Left            =   3240
         TabIndex        =   45
         Top             =   6120
         Width           =   2535
         Begin VB.TextBox txtDBSearchMassTolerances 
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   46
            Text            =   "frmErrorDistribution2D.frx":0051
            Top             =   220
            Width           =   2355
         End
      End
      Begin VB.Frame fraNETCalibrationPlotStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NET Calibration Plot Status"
         Height          =   1300
         Left            =   3240
         TabIndex        =   49
         Top             =   6960
         Width           =   2550
         Begin VB.TextBox txtNETCalibrationPeakCenter 
            BorderStyle     =   0  'None
            Height          =   1000
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   50
            Text            =   "frmErrorDistribution2D.frx":0088
            Top             =   220
            Width           =   2355
         End
      End
      Begin VB.Frame fraMassCalibrationStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mass Calibration Plot Status"
         Height          =   1300
         Left            =   120
         TabIndex        =   47
         Top             =   6960
         Width           =   3015
         Begin VB.TextBox txtMassCalibrationPeakCenter 
            BorderStyle     =   0  'None
            Height          =   1000
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   48
            Text            =   "frmErrorDistribution2D.frx":00E0
            Top             =   220
            Width           =   2835
         End
      End
      Begin VB.Frame fraToleranceRefinementPeakCriteria 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Criteria To Use Peak For Refinement"
         Height          =   1335
         Left            =   120
         TabIndex        =   37
         Top             =   4680
         Width           =   3015
         Begin VB.TextBox txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   44
            Text            =   "2.5"
            Top             =   920
            Width           =   615
         End
         Begin VB.TextBox txtToleranceRefinementMinimumPeakHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   39
            Text            =   "25"
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtToleranceRefinementPercentageOfMaxForWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   42
            Text            =   "60"
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblDescription 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Minimum S/N for Low Abu"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   940
            Width           =   2055
         End
         Begin VB.Label lblUnits 
            BackColor       =   &H00FFFFFF&
            Caption         =   "counts/bin"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   40
            Top             =   330
            Width           =   840
         End
         Begin VB.Label lblDescription 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Minimum Height"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   38
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label lblDescription 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pct of Max for Finding Width"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   41
            Top             =   640
            Width           =   2055
         End
      End
      Begin TabDlg.SSTab tbsRefinement 
         Height          =   4605
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8123
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "Mass Calibration Refinement"
         TabPicture(0)   =   "frmErrorDistribution2D.frx":0138
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblMassCalibrationRefinementDescription"
         Tab(0).Control(1)=   "lblMassCalibrationRefinementUnits(2)"
         Tab(0).Control(2)=   "lblMassCalibrationAdjustment"
         Tab(0).Control(3)=   "lblMassCalibrationRefinementUnits(1)"
         Tab(0).Control(4)=   "lblMassCalibrationOverallAdjustment"
         Tab(0).Control(5)=   "cmdMassCalibrationRefinementStart"
         Tab(0).Control(6)=   "fraMassCalibrationRefinement"
         Tab(0).Control(7)=   "cmdMassCalibrationManual"
         Tab(0).Control(8)=   "txtMassCalibrationNewIncrementalAdjustment"
         Tab(0).Control(9)=   "txtMassCalibrationOverallAdjustment"
         Tab(0).Control(10)=   "cmdMassCalibrationRevert"
         Tab(0).Control(11)=   "cmdRecomputeHistograms"
         Tab(0).Control(12)=   "cmdAbortProcessing"
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "Tolerance Refinement"
         TabPicture(1)   =   "frmErrorDistribution2D.frx":0154
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "cmdMassToleranceRefinementStart"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "fraToleranceRefinementMass"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "fraToleranceRefinementGANET"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdNETToleranceRefinementStart"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.CommandButton cmdAbortProcessing 
            Caption         =   "Abort Processing"
            Height          =   375
            Left            =   -74160
            TabIndex        =   93
            Top             =   4200
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton cmdRecomputeHistograms 
            Caption         =   "&Recompute Histograms"
            Height          =   375
            Left            =   -74160
            TabIndex        =   94
            Top             =   4200
            Width           =   2055
         End
         Begin VB.CommandButton cmdMassCalibrationRevert 
            Caption         =   "Revert to Original"
            Height          =   375
            Left            =   -73920
            TabIndex        =   90
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtMassCalibrationOverallAdjustment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   2430
            Width           =   1100
         End
         Begin VB.TextBox txtMassCalibrationNewIncrementalAdjustment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -72960
            TabIndex        =   15
            Text            =   "0"
            Top             =   2790
            Width           =   1100
         End
         Begin VB.CommandButton cmdMassCalibrationManual 
            Caption         =   "Manually recalibrate"
            Height          =   375
            Left            =   -73080
            TabIndex        =   10
            Top             =   1620
            Width           =   1815
         End
         Begin VB.CommandButton cmdNETToleranceRefinementStart 
            Caption         =   "Start NET Tol Refinement"
            Height          =   615
            Left            =   2160
            TabIndex        =   36
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Frame fraMassCalibrationRefinement 
            Caption         =   "Mass Calibration Refinement"
            Height          =   1215
            Left            =   -74760
            TabIndex        =   2
            Top             =   360
            Width           =   3500
            Begin VB.TextBox txtRefineMassCalibrationMaximumShift 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2040
               TabIndex        =   7
               Text            =   "0"
               Top             =   720
               Width           =   615
            End
            Begin VB.Frame fraMassCalibrationRefinementMassType 
               Caption         =   "Tolerance Type"
               Height          =   855
               Left            =   240
               TabIndex        =   3
               Top             =   240
               Width           =   1455
               Begin VB.OptionButton optRefineMassCalibrationMassType 
                  Caption         =   "Dalton"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   5
                  Top             =   520
                  Width           =   855
               End
               Begin VB.OptionButton optRefineMassCalibrationMassType 
                  Caption         =   "ppm"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   4
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   855
               End
            End
            Begin VB.Label lblMaximumShift 
               Caption         =   "Maximum shift"
               Height          =   255
               Left            =   2040
               TabIndex        =   6
               Top             =   480
               Width           =   1350
            End
            Begin VB.Label lblMassCalibrationRefinementUnits 
               Caption         =   "ppm"
               Height          =   255
               Index           =   0
               Left            =   2760
               TabIndex        =   8
               Top             =   750
               Width           =   600
            End
         End
         Begin VB.Frame fraToleranceRefinementGANET 
            Caption         =   "NET Tolerance Refinement"
            Height          =   1335
            Left            =   240
            TabIndex        =   28
            Top             =   2400
            Width           =   3735
            Begin VB.TextBox txtRefineDBSearchNETToleranceAdjustmentMultiplier 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               TabIndex        =   91
               Text            =   "1"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtRefineDBSearchNETToleranceMinimum 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               TabIndex        =   30
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtRefineDBSearchNETToleranceMaximum 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               TabIndex        =   33
               Text            =   "0"
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lblDescription 
               Caption         =   "Adjustment multiplier"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   92
               Top             =   990
               Width           =   1575
            End
            Begin VB.Label lblDescription 
               Caption         =   "Minimum Tolerance"
               Height          =   255
               Index           =   84
               Left            =   120
               TabIndex        =   29
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label lblUnits 
               Caption         =   "NET"
               Height          =   255
               Index           =   8
               Left            =   2520
               TabIndex        =   31
               Top             =   270
               Width           =   600
            End
            Begin VB.Label lblDescription 
               Caption         =   "Maximum Tolerance"
               Height          =   255
               Index           =   83
               Left            =   120
               TabIndex        =   32
               Top             =   630
               Width           =   1455
            End
            Begin VB.Label lblUnits 
               Caption         =   "NET"
               Height          =   255
               Index           =   7
               Left            =   2520
               TabIndex        =   34
               Top             =   630
               Width           =   600
            End
         End
         Begin VB.Frame fraToleranceRefinementMass 
            Caption         =   "Mass Tolerance Refinement"
            Height          =   1815
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   3735
            Begin VB.TextBox txtRefineDBSearchMassToleranceAdjustmentMultiplier 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               TabIndex        =   26
               Text            =   "1"
               Top             =   960
               Width           =   615
            End
            Begin VB.ComboBox cboMassToleranceRefinementMethod 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   1320
               Width           =   3495
            End
            Begin VB.TextBox txtRefineDBSearchMassToleranceMaximum 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               TabIndex        =   23
               Text            =   "0"
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtRefineDBSearchMassToleranceMinimum 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1800
               TabIndex        =   20
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblDescription 
               Caption         =   "Adjustment multiplier"
               Height          =   255
               Index           =   94
               Left            =   120
               TabIndex        =   25
               Top             =   990
               Width           =   1695
            End
            Begin VB.Label lblUnits 
               Caption         =   "ppm"
               Height          =   255
               Index           =   6
               Left            =   2520
               TabIndex        =   24
               Top             =   630
               Width           =   600
            End
            Begin VB.Label lblDescription 
               Caption         =   "Maximum Tolerance"
               Height          =   255
               Index           =   76
               Left            =   120
               TabIndex        =   22
               Top             =   630
               Width           =   1455
            End
            Begin VB.Label lblUnits 
               Caption         =   "ppm"
               Height          =   255
               Index           =   5
               Left            =   2520
               TabIndex        =   21
               Top             =   270
               Width           =   600
            End
            Begin VB.Label lblDescription 
               Caption         =   "Minimum Tolerance"
               Height          =   255
               Index           =   72
               Left            =   120
               TabIndex        =   19
               Top             =   270
               Width           =   1455
            End
         End
         Begin VB.CommandButton cmdMassToleranceRefinementStart 
            Caption         =   "Start Mass Tol Refinement"
            Height          =   615
            Left            =   360
            TabIndex        =   35
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton cmdMassCalibrationRefinementStart 
            Caption         =   "Auto recalibrate"
            Height          =   375
            Left            =   -74760
            TabIndex        =   9
            Top             =   1620
            Width           =   1575
         End
         Begin VB.Label lblMassCalibrationOverallAdjustment 
            Caption         =   "Overall Adjustment"
            Height          =   255
            Left            =   -74760
            TabIndex        =   11
            Top             =   2460
            Width           =   1700
         End
         Begin VB.Label lblMassCalibrationRefinementUnits 
            Caption         =   "ppm"
            Height          =   255
            Index           =   1
            Left            =   -71760
            TabIndex        =   13
            Top             =   2460
            Width           =   600
         End
         Begin VB.Label lblMassCalibrationAdjustment 
            Caption         =   "New Adjustment (relative to overall)"
            Height          =   375
            Left            =   -74760
            TabIndex        =   14
            Top             =   2710
            Width           =   1695
         End
         Begin VB.Label lblMassCalibrationRefinementUnits 
            Caption         =   "ppm"
            Height          =   255
            Index           =   2
            Left            =   -71760
            TabIndex        =   16
            Top             =   2820
            Width           =   600
         End
         Begin VB.Label lblMassCalibrationRefinementDescription 
            Caption         =   "Mass Calibration Refinement Description"
            Height          =   855
            Left            =   -74880
            TabIndex        =   17
            Top             =   3200
            Width           =   4005
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveDataToTextFile 
         Caption         =   "Save Data to Text File"
      End
      Begin VB.Menu mnuSaveBinnedDataToTextFile 
         Caption         =   "Save Binned Data to &Text File"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveChartPicture 
         Caption         =   "Save Chart as &PNG"
         Index           =   1
      End
      Begin VB.Menu mnuSaveChartPicture 
         Caption         =   "Save Chart as &JPEG"
         Index           =   2
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopyErrors 
         Caption         =   "Copy Differences (ppm, Da, and NET)"
      End
      Begin VB.Menu mnuCopyErrorsBinned 
         Caption         =   "Copy Binned Differences (ppm, Da, and NET)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyVisibleBinnedDifferences 
         Caption         =   "&Copy Binned Differences (visible only)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy as &BMP"
         Index           =   0
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy as &WMF"
         Index           =   1
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy as &EMF"
         Index           =   2
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetToleranceRefinementOptionsToDefault 
         Caption         =   "Set Tolerance Refinement Options To Defaults"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView3DErrorDistributions 
         Caption         =   "&3D Error Distributions (Mass vs. GANET)"
      End
      Begin VB.Menu mnuView3DErrorDistributionsInverted 
         Caption         =   "&3D Error Distributions Inverted"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmErrorDistribution2DLoadedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ccmCopyChartMode
    ccmBMP = 0
    ccmWMF = 1
    ccmEMF = 2
End Enum

Public Enum mdmMassErrorDisplayModeConstants
    mdmMassErrorPPM = 0
    mdmMassErrorDa = 1
    mdmGanetError = 2
End Enum

Private Type udtUMCStatsDetailsType
    MassWidthMedian As Double
    MassWidthMaximum As Double
    MassStDevMedian As Double
    MassStDevMaximum As Double
End Type

Private Type udtUMCStatsType
    Count As Long
    PPMStats As udtUMCStatsDetailsType
    DaStats As udtUMCStatsDetailsType
End Type

Private Const MASS_PPM_ADJUSTMENT_PRECISION = 4
Private Const MASS_DA_ADJUSTMENT_PRECISION = 8
Private Const GANET_ADJUSTMENT_PRECISION = 4

Public CallerID As Long

Private mErrorsCount As Long
Private mErrorsCountDimmed As Long
Private mRawMassErrorsPPM() As Single        ' 0-based array, in PPM
Private mRawMassErrorsDa() As Single         ' 0-based array, in Da
Private mRAWGanetErrors() As Single          ' 0-based array

Private mMassPPMErrors As udtBinnedDataType
Private mMassDaErrors As udtBinnedDataType
Private mGanetErrors As udtBinnedDataType

Private mUMCStats As udtUMCStatsType

Private mAMTIDSorted() As Long                  ' 0-based array
Private mAMTIDSortedInd() As Long               ' 0-based array

Private mInternalStdIDSorted() As Long                 ' 0-based array
Private mInternalStdIDSortedInd() As Long              ' 0-based array

Private mAMTIndicesInitialized As Boolean
Private mUpdatingControls As Boolean
Private mCalculating As Boolean
Private mFormInitialized As Boolean

Private mMassCalErrorPeakCached As udtErrorPlottingPeakCacheType
Private mNETTolErrorPeakCached As udtErrorPlottingPeakCacheType

Private mGraphTitle As String

'Expression Evaluator variables for elution time calculation
Private MyExprEva As ExprEvaluator
Private VarVals() As Long
Private MinFN As Long
Private MaxFN As Long

Private objHistogram As New clsHistogram

Private mGelAnalysisIsValid As Boolean
Private mAbortProcessing As Boolean
'

Private Function AddNewErrValues(ByRef Refs() As String, ByVal RefsCnt As Long, ByVal dblIonMass As Double, ByVal sngIonGANET As Single, ByVal dblIonAbundance As Double, ByVal dblMassErrPPMCorrection As Double, ByVal blnInternalStdMatch As Boolean) As Boolean
    
    Dim strRefID As String
    Dim strPPMMassError As String
    
    Dim lngMatchIndex As Long
    Dim lngIndexPointer As Long
    Dim lngMTIndex As Long, lngInternalStdIndex As Long
    
    Dim dblRefMW As Double, dblRefNET As Double
    
    Dim dblMassErrorDa As Double                ' Mass error newly computed here (in Da)
    Dim dblMassErrorPPM As Double               ' Mass error newly computed here (in ppm); if dblMassErrorPPMFromID is smaller, then it gets updated to that
    Dim dblMassErrorPPMFromID As Double         ' Mass error recorded in Refs()
    
    Dim sngGANETError As Single
    Dim blnValidSLiC As Boolean
    Dim blnHitFound As Boolean
    
    Static LastErrorDisplayTime As Date
    
On Error GoTo AddNewErrValuesErrorHandler

    If glbPreferencesExpanded.RefineMSDataOptions.MaximumAbundance > 0 Then
        If dblIonAbundance > glbPreferencesExpanded.RefineMSDataOptions.MaximumAbundance Then
            ' Do not use this data point
            AddNewErrValues = False
            Exit Function
        End If
    End If
    

    blnHitFound = False
    If RefsCnt > 0 Then
        For lngMatchIndex = 1 To RefsCnt
            
            lngMTIndex = -1
            lngInternalStdIndex = -1
            
            If blnInternalStdMatch Then
            
                ' Extract MT tag ID
                strRefID = GetIDFromString(Refs(lngMatchIndex), INT_STD_MARK, AMTIDEnd)
                
                ' Extract the recorded mass error (in ppm)
                strPPMMassError = GetMWErrFromString(Refs(lngMatchIndex))
            
                blnValidSLiC = True
                If glbPreferencesExpanded.RefineMSDataOptions.MinimumSLiC > 0 Then
                    If GetSLiCFromString(Refs(lngMatchIndex)) < glbPreferencesExpanded.RefineMSDataOptions.MinimumSLiC Then
                        blnValidSLiC = False
                    End If
                End If
                
                If blnValidSLiC Then
                    lngIndexPointer = BinarySearchLng(mInternalStdIDSorted(), CLngSafe(strRefID))
                    If lngIndexPointer >= 0 Then
                        lngInternalStdIndex = mInternalStdIDSortedInd(lngIndexPointer)
                        
                        Debug.Assert InStr(Refs(lngMatchIndex), UMCInternalStandards.InternalStandards(lngInternalStdIndex).SeqID) > 0
                        dblRefMW = UMCInternalStandards.InternalStandards(lngInternalStdIndex).MonoisotopicMass
                        dblRefNET = UMCInternalStandards.InternalStandards(lngInternalStdIndex).NET
    
                    Else
                        lngInternalStdIndex = -1
                    End If
                End If
                
            Else
                ' Extract MT tag ID
                strRefID = GetIDFromString(Refs(lngMatchIndex), AMTMark, AMTIDEnd)
                
                ' Extract the recorded mass error (in ppm)
                strPPMMassError = GetMWErrFromString(Refs(lngMatchIndex))
                
                blnValidSLiC = True
                If glbPreferencesExpanded.RefineMSDataOptions.MinimumSLiC > 0 Then
                    If GetSLiCFromString(Refs(lngMatchIndex)) < glbPreferencesExpanded.RefineMSDataOptions.MinimumSLiC Then
                        blnValidSLiC = False
                    End If
                End If

                If blnValidSLiC Then
                    lngIndexPointer = BinarySearchLng(mAMTIDSorted(), CLngSafe(strRefID))
                    If lngIndexPointer >= 0 Then
                        lngMTIndex = mAMTIDSortedInd(lngIndexPointer)
                    
                        Debug.Assert InStr(Refs(lngMatchIndex), AMTData(lngMTIndex).ID) > 0
                        dblRefMW = AMTData(lngMTIndex).MW
                        dblRefNET = AMTData(lngMTIndex).NET
                    
                    Else
                        lngMTIndex = -1
                    End If
                End If
                
            End If
            
            If blnValidSLiC Then
                
                If IsNumeric(strPPMMassError) Then
                    dblMassErrorPPMFromID = CDbl(strPPMMassError) + dblMassErrPPMCorrection
                Else
                    ' This probably shouldn't happen
                    Debug.Assert False
                    dblMassErrorPPMFromID = 0
                End If
                
                If AMTCnt <= 0 Or lngMTIndex >= 0 Or lngInternalStdIndex >= 0 Then
                    
                    If AMTCnt = 0 Then
                        ' AMTs not in memory; must use the mass error that was stored in Refs
                        ' In addition, will not have a NET error since that isn't stored in the Ref string
                        dblMassErrorPPM = dblMassErrorPPMFromID
                        dblMassErrorDa = PPMToMass(dblMassErrorPPM, dblIonMass)
                        sngGANETError = 0
                    Else
                        dblMassErrorDa = dblIonMass - dblRefMW
                        dblMassErrorPPM = MassToPPM(dblMassErrorDa, dblIonMass)
                        
                        ' dblMassErrorPPM will be drastically off if the user searched the database using modified MT tag masses (e.g. alkylation, ICAT, or N15)
                        ' For this reason, we can compare dblMassErrorPPM with dblMassErrorPPMFromID
                        ' If dblMassErrorPPMFromID is a little more than half what dblMassErrorPPM is, then
                        '  we'll use dblMassErrorPPMFromID instead
                        ' We don't always want to use the dblMassErrorPPMFromID value since it is normally always positive,
                        '  though I've added modifications to store negative values (where appropriate) when the MT tag mass is, in fact, modified (for example, see RecordSearchResultsInData)
                        '
                        ' If the mass error from the ID is less than the computed mass error * 45%, then update the mass error
                        ' The reason for the decrease of the computed mass error by 55% is to try to prevent inadvertent updating due to rounding errors
                        If Abs(dblMassErrorPPMFromID) < Abs(dblMassErrorPPM) * 0.45 Then
                            If dblMassErrorPPMFromID = 0 Then
                                Debug.Assert Abs(dblMassErrorPPM) < 0.01
                            End If
                            
                            dblMassErrorPPM = dblMassErrorPPMFromID
                            dblMassErrorDa = PPMToMass(dblMassErrorPPM, dblIonMass)
                        End If
                    
                        sngGANETError = sngIonGANET - dblRefNET
                    End If
                    
                    
                    ' Add Errors to mRawMassErrorsPPM(), mRawMassErrorsDa(), and mRAWGanetErrors()
                    mRawMassErrorsPPM(mErrorsCount) = CSng(MassToPPM(dblMassErrorDa, dblIonMass))
                    mRawMassErrorsDa(mErrorsCount) = CSng(dblMassErrorDa)
                    mRAWGanetErrors(mErrorsCount) = sngGANETError
                    
                    ''mDataSourceIonIndex(mErrorsCount) = lngIonIndex
                    
                    mErrorsCount = mErrorsCount + 1
                    If mErrorsCount >= mErrorsCountDimmed Then
                        mErrorsCountDimmed = mErrorsCountDimmed + 100
                        ReDim Preserve mRawMassErrorsPPM(mErrorsCountDimmed)
                        ReDim Preserve mRawMassErrorsDa(mErrorsCountDimmed)
                        ReDim Preserve mRAWGanetErrors(mErrorsCountDimmed)
                        ''ReDim Preserve mDataSourceIonIndex(mErrorsCountDimmed)
                    End If
                    
                    blnHitFound = True
                Else
                    ' AMTID or InternalStdID not found; user could be connected to a different database than that originally used for searching
                    ' Another possibility is that the GANET value for this MT tag was not null when the search was performed,
                    '  but now it is null
                    ' Dispaly a message in the debug window every 0.5 second
                    If (Now() - LastErrorDisplayTime) * 24# * 60# * 60# >= 0.5 Then
                        Debug.Print strRefID & " not found in currently loaded MT tags"
                        LastErrorDisplayTime = Now
                    End If
                End If
            End If
        Next lngMatchIndex
    End If
    
    AddNewErrValues = blnHitFound
    Exit Function

AddNewErrValuesErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Unexpected error in AddNewErrValues" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
    End If
    LogErrors Err.Number, "AddNewErrValues"
    AddNewErrValues = False
    
End Function

Private Sub ComputeErrors(blnForceUpdate As Boolean)
    
    Dim lngIndex As Long
    Dim lngCSDataIndex As Long
    Dim lngIsoDataIndex As Long
    
    Dim IsoField As Integer
    Dim lngDataWithHits As Long

    Dim dblMassErrPPMCorrection As Double

    If mUpdatingControls Then
        Exit Sub
    End If
    
    If mCalculating And Not blnForceUpdate Then
        Exit Sub
    End If
    
    If CallerID < 1 Or CallerID > UBound(GelData()) Then Exit Sub
    
    If AMTCnt <= 0 Then
        UpdateStatus "Warning: MT tags not loaded; only mass errors will be shown"
    End If
    
    If glbPreferencesExpanded.RefineMSDataOptions.UseUMCClassStats Then
        ' Must have UMCs in memory
        If GelUMC(CallerID).UMCCnt <= 0 Then
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "UMCs not found in memory.  Unable to compute mass errors using UMC class stats. Please use menu item 'Steps->2. Find UMCs' in the main window to cluster the data into unique mass classes.", vbInformation + vbOKOnly, "No UMCs"
            End If
            UpdateStatus "Not ready"
            Exit Sub
        End If
    End If
    
    UpdateStatus "Populating mass error bins"
    EnableDisableCalculating True, True
    mAbortProcessing = False
    
On Error GoTo ComputeMassErrorsErrorHandler
    
    mErrorsCountDimmed = 100
    ReDim mRawMassErrorsPPM(mErrorsCountDimmed)
    ReDim mRawMassErrorsDa(mErrorsCountDimmed)
    ReDim mRAWGanetErrors(mErrorsCountDimmed)
    ''ReDim mDataSourceIonIndex(mErrorsCountDimmed)
    mErrorsCount = 0
    
    IsoField = GelData(CallerID).Preferences.IsoDataField
    If IsoField < isfMWavg Or IsoField > isfMWavg Then IsoField = isfMWMono
    
    
    ' First fill mRawMassErrorsPPM(), mRawMassErrorsDa(), and mRAWGanetErrors() with all of the observed errors
    If glbPreferencesExpanded.RefineMSDataOptions.UseUMCClassStats Then
        ' Working with UMCs
        With GelUMC(CallerID)
            For lngIndex = 0 To .UMCCnt - 1
                With .UMCs(lngIndex)
                    ' Just examine the class rep for each UMC
                    Select Case .ClassRepType
                    Case gldtCS
                        ' Note: Since the match to the MT tag is stored with a mass error value relative to a given data point's mass,
                        '       we need to compute the difference in mass between the class rep and the class mass (converted to ppm)
                        dblMassErrPPMCorrection = MassToPPM(.ClassMW - GelData(CallerID).CSData(.ClassRepInd).AverageMW, .ClassMW)
                        If ComputeErrorsExtractValues(GelData(CallerID).CSData(.ClassRepInd).MTID, .ClassMW, .ClassAbundance, GelData(CallerID).CSData(.ClassRepInd).ScanNumber, dblMassErrPPMCorrection) Then
                           ' Hit found
                           lngDataWithHits = lngDataWithHits + 1
                       End If
                    Case gldtIS
                        ' See note above
                        dblMassErrPPMCorrection = MassToPPM(.ClassMW - GetIsoMass(GelData(CallerID).IsoData(.ClassRepInd), IsoField), .ClassMW)
                        If ComputeErrorsExtractValues(GelData(CallerID).IsoData(.ClassRepInd).MTID, .ClassMW, .ClassAbundance, GelData(CallerID).IsoData(.ClassRepInd).ScanNumber, dblMassErrPPMCorrection) Then
                            ' Hit found
                            lngDataWithHits = lngDataWithHits + 1
                        End If
                    Case Else
                        Debug.Assert False
                    End Select
                End With
                If lngIndex Mod 500 = 0 Then
                    UpdateStatus "Populating mass error bins: " & Trim(lngIndex) & " / " & Trim(.UMCCnt)
                    If mAbortProcessing Then Exit For
                End If
            Next lngIndex
        End With
    Else
        ' Working with individual data points
        With GelData(CallerID)
            ' Charge state data
            For lngCSDataIndex = 1 To .CSLines
                With .CSData(lngCSDataIndex)
                    If ComputeErrorsExtractValues(.MTID, .AverageMW, .Abundance, .ScanNumber, 0) Then
                        ' Hit found
                        lngDataWithHits = lngDataWithHits + 1
                    End If
                End With
                If lngCSDataIndex Mod 1000 = 0 Then
                    UpdateStatus "Populating mass error bins: " & Trim(lngCSDataIndex) & " / " & Trim(.CSLines)
                    If mAbortProcessing Then Exit For
                End If
            Next lngCSDataIndex
            
            ' Isotopic data
            For lngIsoDataIndex = 1 To .IsoLines
                With .IsoData(lngIsoDataIndex)
                    If ComputeErrorsExtractValues(.MTID, GetIsoMass(GelData(CallerID).IsoData(lngIsoDataIndex), IsoField), .Abundance, .ScanNumber, 0) Then
                        ' Hit found
                        lngDataWithHits = lngDataWithHits + 1
                    End If
                End With
                If lngIsoDataIndex Mod 1000 = 0 Then
                    UpdateStatus "Populating mass error bins: " & Trim(lngIsoDataIndex) & " / " & Trim(.IsoLines)
                    If mAbortProcessing Then Exit For
                End If
            Next lngIsoDataIndex
        End With
    End If
    
    If mAbortProcessing Then
        mErrorsCount = 0
    End If
    
    If mErrorsCount > 0 Then
        UpdateStatus "Binning data"
        
        ' Initialize the histogram object
        With objHistogram
            .RequireNegativeStartBin = True
            .ShowMessages = Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
        End With
                
        ' Now bin the data
        ' First the mass errors (in PPM)
        With mMassPPMErrors
            objHistogram.BinSize = glbPreferencesExpanded.ErrorPlottingOptions.MassBinSizePPM
            objHistogram.DefaultBinSize = 0.5
            objHistogram.StartBinDigitsAfterDecimal = 0
            
            If Not objHistogram.BinData(mRawMassErrorsPPM, mErrorsCount, .Binned, .BinnedCount) Then
                LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputeErrors"
            End If
            
            .StartBin = objHistogram.StartBin
            .BinSize = objHistogram.BinSize
            .BinRangeMaximum = objHistogram.BinRangeMaximum
            
            ReDim .SmoothedBins(UBound(.Binned))
            For lngIndex = 0 To UBound(.Binned)
                .SmoothedBins(lngIndex) = .Binned(lngIndex)
            Next lngIndex
        End With
        
        ' Now the mass errors (in Da)
        With mMassDaErrors
            objHistogram.BinSize = Round(PPMToMass(CDbl(glbPreferencesExpanded.ErrorPlottingOptions.MassBinSizePPM), 2000), 6)
            objHistogram.DefaultBinSize = 0.005
            objHistogram.StartBinDigitsAfterDecimal = 3
            
            If Not objHistogram.BinData(mRawMassErrorsDa, mErrorsCount, .Binned, .BinnedCount) Then
                LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputeErrors"
            End If
            
            .StartBin = objHistogram.StartBin
            .BinSize = objHistogram.BinSize
            .BinRangeMaximum = objHistogram.BinRangeMaximum
        
            ReDim .SmoothedBins(UBound(.Binned))
            For lngIndex = 0 To UBound(.Binned)
                .SmoothedBins(lngIndex) = .Binned(lngIndex)
            Next lngIndex
        End With

        ' Finally, the GANET errors
        With mGanetErrors
            objHistogram.BinSize = glbPreferencesExpanded.ErrorPlottingOptions.GANETBinSize
            objHistogram.DefaultBinSize = 0.005
            objHistogram.StartBinDigitsAfterDecimal = 1
        
            If Not objHistogram.BinData(mRAWGanetErrors, mErrorsCount, .Binned, .BinnedCount) Then
                LogErrors objHistogram.ErrorNumber, "frmErrorDistribution2DLoaddedData->ComputeErrors"
            End If
            
            .StartBin = objHistogram.StartBin
            .BinSize = objHistogram.BinSize
            .BinRangeMaximum = objHistogram.BinRangeMaximum
        
            ReDim .SmoothedBins(UBound(.Binned))
            For lngIndex = 0 To UBound(.Binned)
                .SmoothedBins(lngIndex) = .Binned(lngIndex)
            Next lngIndex
        End With

    Else
        ' This will happen if there were no matches found when searching the database
        '  or if the user aborted processing
        mMassPPMErrors.BinnedCount = 0
        mMassDaErrors.BinnedCount = 0
        mGanetErrors.BinnedCount = 0
        
        ReDim mMassPPMErrors.SmoothedBins(1)
        ReDim mMassDaErrors.SmoothedBins(1)
        ReDim mGanetErrors.SmoothedBins(1)
        
    End If
    
    EnableDisableCalculating False
    
    UpdateStatus "Updating plot"
    UpdatePlot
    
''    If ctlPlotter.GetSeriesCount > 1 Then
''        ' Update the plot again due to a refresh bug with the plot that prevents both series
''        ' from being visible on the first update (very odd behavior that I can't track down)
''        UpdatePlot
''    End If
    
    UpdateStatus "Data points with 1 or more hits = " & Trim(lngDataWithHits)
    
    Exit Sub
    
ComputeMassErrorsErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Unexpected error in ComputeErrors" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "ComputeErrors"
    EnableDisableCalculating False

End Sub

'''Private Function ComputeErrorsExtractValuesCS(ByVal lngCSDataIndex As Long, ByVal dblMass As Double) As Boolean
'''    Dim blnSuccess As Boolean
'''    With GelData(CallerID)
'''        blnSuccess = ComputeErrorsExtractValues(.CSData(lngCSDataIndex).MTID, dblMass, .CSData(lngCSDataIndex).ScanNumber)
'''    End With
'''    ComputeErrorsExtractValuesCS = blnSuccess
'''End Function
'''
'''Private Function ComputeErrorsExtractValuesIso(ByVal lngIsoDataIndex As Long, ByVal dblMass As Double) As Boolean
'''    Dim blnSuccess As Boolean
'''    With GelData(CallerID)
'''        blnSuccess = ComputeErrorsExtractValues(.IsoData(lngIsoDataIndex).MTID, dblMass, .IsoData(lngIsoDataIndex).ScanNumber)
'''    End With
'''    ComputeErrorsExtractValuesIso = blnSuccess
'''End Function

Private Function ComputeErrorsExtractValues(ByVal strRefString As String, ByVal dblIonMass As Double, ByVal dblIonAbundance As Double, ByVal lngScanNumber As Long, dblMassErrPPMCorrection As Double) As Boolean
    Dim sngIonGANET As Single
    
    Dim Refs() As String         ' 1-based array
    Dim RefsCnt As Long
    
    Dim blnHitFound As Boolean
    
    If IsAMTReferenced(strRefString) Then
        RefsCnt = GetAMTRefFromString2(strRefString, Refs())
        sngIonGANET = ConvertScanToNET(lngScanNumber)
        
        ' Compute mass error in ppm between AMT and actual data
        ' Record in mRawMassErrorsPPM(), mRawMassErrorsDa(), and mRAWGanetErrors()
        blnHitFound = AddNewErrValues(Refs(), RefsCnt, dblIonMass, sngIonGANET, dblIonAbundance, dblMassErrPPMCorrection, False)
    End If
    
    If glbPreferencesExpanded.RefineMSDataOptions.IncludeInternalStdMatches Then
        If IsInternalStdReferenced(strRefString) Then
            RefsCnt = GetInternalStdRefFromString2(strRefString, Refs())
            
            If Not blnHitFound Then
                sngIonGANET = ConvertScanToNET(lngScanNumber)
            End If
        
            blnHitFound = AddNewErrValues(Refs(), RefsCnt, dblIonMass, sngIonGANET, dblIonAbundance, dblMassErrPPMCorrection, True)
        End If
    End If
    
    ComputeErrorsExtractValues = blnHitFound

End Function

Private Sub ConstructLookupArrays()
    Dim objQSLong As New QSLong
    Dim lngIndex As Long

    UpdateStatus "Initializing index arrays"
    
On Error GoTo ConstructAMTIndexLookupArrayErrorHandler
    
    If AMTCnt > 0 Then
        ReDim mAMTIDSorted(0 To AMTCnt - 1)
        ReDim mAMTIDSortedInd(0 To AMTCnt - 1)
    
        For lngIndex = 1 To AMTCnt
            Debug.Assert IsNumeric(AMTData(lngIndex).ID)
            mAMTIDSorted(lngIndex - 1) = AMTData(lngIndex).ID        ' Note that AMTData() is a 1-based array
            mAMTIDSortedInd(lngIndex - 1) = lngIndex
        Next lngIndex
    
        If objQSLong.QSAsc(mAMTIDSorted(), mAMTIDSortedInd()) Then
            ' All is fine
        Else
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "Error while constructing the MT tag ID index arrays", vbInformation + vbOKOnly, "Error"
            End If
            ReDim mAMTIDSorted(0)
            ReDim mAMTIDSortedInd(0)
        End If
    Else
        ReDim mAMTIDSorted(0)
        ReDim mAMTIDSortedInd(0)
    End If
    
    If UMCInternalStandards.Count > 0 Then
        With UMCInternalStandards
            ReDim mInternalStdIDSorted(0 To .Count - 1)
            ReDim mInternalStdIDSortedInd(0 To .Count - 1)
        
            For lngIndex = 0 To .Count - 1
                mInternalStdIDSorted(lngIndex) = .InternalStandards(lngIndex).SeqID
                mInternalStdIDSortedInd(lngIndex) = lngIndex
            Next lngIndex
        
            If objQSLong.QSAsc(mInternalStdIDSorted(), mInternalStdIDSortedInd()) Then
                ' All is fine
            Else
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                    MsgBox "Error while constructing the InternalStdID index arrays", vbInformation + vbOKOnly, "Error"
                End If
                ReDim mInternalStdIDSorted(0)
                ReDim mInternalStdIDSortedInd(0)
            End If
        End With
    Else
        ReDim mInternalStdIDSorted(0)
        ReDim mInternalStdIDSortedInd(0)
    End If
    
    Set objQSLong = Nothing
    Exit Sub
    
ConstructAMTIndexLookupArrayErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Unexpected error in ConstructLookupArrays()" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "ConstructLookupArrays"
    
End Sub

Private Function ConvertScanToNET(lngScanNumber As Long) As Double

    If mGelAnalysisIsValid Then
        ConvertScanToNET = ScanToGANET(CallerID, lngScanNumber)
    Else
        ConvertScanToNET = Elution(lngScanNumber, MinFN, MaxFN)
    End If

End Function

Private Sub EnableDisableCalculating(ByVal blnCalculating As Boolean, Optional ByVal blnShowAbortButton As Boolean = False)
    Dim blnEnableOptions As Boolean
    
    blnEnableOptions = Not blnCalculating
    
    mCalculating = blnCalculating
    cmdAbortProcessing.Visible = blnCalculating And blnShowAbortButton
    
    chkIncludeInternalStandards.Enabled = blnEnableOptions
    txtMassRangePPM.Enabled = blnEnableOptions
    txtGANETRange.Enabled = blnEnableOptions
    txtMassBinSizePPM.Enabled = blnEnableOptions
    txtGANETBinSize.Enabled = blnEnableOptions
    txtButterworthFrequency.Enabled = blnEnableOptions
    
    cboErrorDisplayMode.Enabled = blnEnableOptions
    chkUseUMCClassStats.Enabled = blnEnableOptions
    txtMinimumSLiC.Enabled = blnEnableOptions
    
    If blnCalculating Then
        Me.MousePointer = vbHourglass
    Else
        Me.MousePointer = vbDefault
    End If
    
    DoEvents

End Sub

Private Function ExportErrorsToClipboardOrFile(Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Boolean
    ' Returns True if success, False if failure

    Dim strErrorData() As String
    Dim strTextToCopy As String
    
    Dim OutFileNum As Integer
    Dim lngIndex As Long, lngOutputArrayCount As Long
    
    If mErrorsCount = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        ExportErrorsToClipboardOrFile = False
        Exit Function
    End If
    
On Error GoTo ExportMassErrorsErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    ' Header row is strErrorData(0)
    ' Data is from strErrorData(1) to strErrorData(mMassPPMErrors.BinnedCount + mGANETErrors.BinnedCount)
    ReDim strErrorData(0 To mErrorsCount)
    
    ' Fill strErrorData() with the Mass Errors and GANET errors
    strErrorData(0) = "MassErrorPPM" & vbTab & "MassErrorDa" & vbTab & "NETError"
    
    For lngIndex = 0 To mErrorsCount - 1
        strErrorData(lngIndex + 1) = Trim(mRawMassErrorsPPM(lngIndex)) & vbTab & Trim(mRawMassErrorsDa(lngIndex)) & vbTab & Trim(mRAWGanetErrors(lngIndex))
        
        If lngIndex Mod 1000 = 0 Then UpdateStatus "Exporting: " & Trim(lngIndex) & " / " & mErrorsCount
    Next lngIndex
    lngOutputArrayCount = mErrorsCount + 1
    
    If Len(strFilePath) > 0 Then
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        For lngIndex = 0 To lngOutputArrayCount - 1
            Print #OutFileNum, strErrorData(lngIndex)
        Next lngIndex
        
        Close #OutFileNum
    Else
        UpdateStatus "Exporting: Preparing for clipboard"
        strTextToCopy = FlattenStringArray(strErrorData(), lngOutputArrayCount, vbCrLf, False)
        Clipboard.Clear
        Clipboard.SetText strTextToCopy, vbCFText
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    ExportErrorsToClipboardOrFile = True
    Exit Function
    
ExportMassErrorsErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    ExportErrorsToClipboardOrFile = False

End Function

Private Sub DisplayErrorPlotPeakStats()
        
    Dim blnValidPeakFound As Boolean
    Dim blnSingleGoodPeakFound As Boolean
    
    Dim udtMassErrorPeak As udtPeakStatsType
    Dim udtNETErrorPeak As udtPeakStatsType
    
    Dim lngDigitsOfPrecisionToRoundMassTo As Long
    Dim lngMultiplier As Long
    
    Dim strUnits As String
    
    Dim dblPeakCenter As Double, dblPeakWidth As Double, dblPeakHeight As Double
    Dim sngSignalToNoise As Single
    
    Dim strRelativeRisk As String
    Dim strSignalToNoise As String
    
On Error GoTo DisplayErrorPlotPeakStatsErrorHandler

    If mUpdatingControls Then Exit Sub
    
    ' Reset these to 0 for now
    With mMassCalErrorPeakCached
        .Center = 0
        .width = 0
        .Height = 0
        .SingleValidPeak = False
        With .PeakStats
            .IndexOfMaximum = 0
            .IndexBaseLeft = 0
            .IndexBaseRight = 0
            .TruePositiveArea = 0
            .FalsePositiveArea = 0
        End With
    End With
    
    With mNETTolErrorPeakCached
        .Center = 0
        .width = 0
        .Height = 0
        .SingleValidPeak = False
        With .PeakStats
            .IndexOfMaximum = 0
            .IndexBaseLeft = 0
            .IndexBaseRight = 0
            .TruePositiveArea = 0
            .FalsePositiveArea = 0
        End With
    End With

    
    If cboErrorDisplayMode.ListIndex = mdmMassErrorDa Then
        lngDigitsOfPrecisionToRoundMassTo = MASS_DA_ADJUSTMENT_PRECISION - 1
        lngMultiplier = 1000
        strUnits = "mDa"            ' Millidaltons
        
        blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mMassDaErrors, udtMassErrorPeak, blnSingleGoodPeakFound)
    
        If blnValidPeakFound Then
            GetPeakStats mMassDaErrors, udtMassErrorPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, lngDigitsOfPrecisionToRoundMassTo
        
            With mMassCalErrorPeakCached
                .Center = MassToPPM(dblPeakCenter, 1000)
                .width = MassToPPM(dblPeakWidth, 1000)
                .Height = dblPeakHeight
                ' .SignalToNoise = sngSignalToNoise
                .SingleValidPeak = blnSingleGoodPeakFound
                .PeakStats = udtMassErrorPeak
            End With
        End If
        
    Else
        lngDigitsOfPrecisionToRoundMassTo = MASS_PPM_ADJUSTMENT_PRECISION - 1
        lngMultiplier = 1
        strUnits = GetSearchToleranceUnitText(gltPPM)
    
        blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mMassPPMErrors, udtMassErrorPeak, blnSingleGoodPeakFound)
        
        If blnValidPeakFound Then
            GetPeakStats mMassPPMErrors, udtMassErrorPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, lngDigitsOfPrecisionToRoundMassTo
        
            With mMassCalErrorPeakCached
                .Center = dblPeakCenter
                .width = dblPeakWidth
                .Height = dblPeakHeight
                .SignalToNoise = sngSignalToNoise
                .SingleValidPeak = blnSingleGoodPeakFound
                .PeakStats = udtMassErrorPeak
            End With
        End If
    End If
    
    strRelativeRisk = ""
    If blnValidPeakFound Then
        If sngSignalToNoise >= 3 Then
            strSignalToNoise = Round(sngSignalToNoise, 0)
        Else
            strSignalToNoise = Round(sngSignalToNoise, 1)
        End If
        txtMassCalibrationPeakCenter = "Peak Center: " & Round(dblPeakCenter * lngMultiplier, 3) & " " & strUnits & vbCrLf & "Peak Width: " & Round(dblPeakWidth * lngMultiplier, 3) & " " & strUnits & vbCrLf & "Peak Height: " & Round(dblPeakHeight, 0) & " counts/bin" & vbCrLf & "S/N: " & strSignalToNoise & vbCrLf & "Single good peak: " & blnSingleGoodPeakFound
        
        With mMassCalErrorPeakCached
            strRelativeRisk = DisplayErrorPlotRelativeRisk("Mass: ", .PeakStats.TruePositiveArea, .PeakStats.FalsePositiveArea) & vbCrLf
        End With
    Else
        txtMassCalibrationPeakCenter = "A valid peak could not be found"
        strRelativeRisk = "Valid mass error peak not found" & vbCrLf
    End If
    
    blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mGanetErrors, udtNETErrorPeak, blnSingleGoodPeakFound)
    If blnValidPeakFound Then
        GetPeakStats mGanetErrors, udtNETErrorPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, GANET_ADJUSTMENT_PRECISION
    
        With mNETTolErrorPeakCached
            .Center = dblPeakCenter
            .width = dblPeakWidth
            .Height = dblPeakHeight
            .SignalToNoise = sngSignalToNoise
            .SingleValidPeak = blnSingleGoodPeakFound
            .PeakStats = udtNETErrorPeak
        
            strRelativeRisk = strRelativeRisk & DisplayErrorPlotRelativeRisk("NET: ", .PeakStats.TruePositiveArea, .PeakStats.FalsePositiveArea)
        End With
    
        If sngSignalToNoise >= 3 Then
            strSignalToNoise = Round(sngSignalToNoise, 0)
        Else
            strSignalToNoise = Round(sngSignalToNoise, 1)
        End If
        txtNETCalibrationPeakCenter = "Peak Center: " & Round(dblPeakCenter, 4) & " NET" & vbCrLf & "Peak Width: " & Round(dblPeakWidth, 4) & " NET" & vbCrLf & "Peak Height: " & Round(dblPeakHeight, 0) & " counts/bin" & vbCrLf & "S/N: " & strSignalToNoise & vbCrLf & "Single good peak: " & blnSingleGoodPeakFound
    Else
        txtNETCalibrationPeakCenter = "A valid peak could not be found"
        strRelativeRisk = strRelativeRisk & "Valid NET error peak not found"
    End If

    txtRelativeRiskStatistics = strRelativeRisk
    
    Exit Sub

DisplayErrorPlotPeakStatsErrorHandler:
    Debug.Print "Error in DisplayErrorPlotPeakStats: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "DisplayErrorPlotPeakStats"
    
End Sub

Private Function DisplayErrorPlotRelativeRisk(ByVal strLineLabel As String, ByVal lngTruePositiveArea As Long, ByVal lngFalsePositiveArea As Long) As String

    Dim strRelativeRisk As String
    Dim sngRelativeRisk As Single
    
    If lngTruePositiveArea + lngFalsePositiveArea > 0 Then
        sngRelativeRisk = lngFalsePositiveArea / (lngTruePositiveArea + lngFalsePositiveArea)
    Else
        sngRelativeRisk = 0
    End If

    strRelativeRisk = strLineLabel & Round(sngRelativeRisk * 100, 1) & "% (FP: " & lngFalsePositiveArea & ", TP: " & lngTruePositiveArea & ")"

    DisplayErrorPlotRelativeRisk = strRelativeRisk
    
End Function

Private Sub ComputeCurrentUMCStats()
    ' Update UMC Statistics
    
    Dim lngIndex As Long
    
    Dim dblMassWidthsPPM() As Double
    Dim dblMassWidthsDa() As Double
    Dim dblMassStDevPPM() As Double
    Dim dblMassStDevDa() As Double
    
    Dim objStats As New StatDoubles
    
On Error GoTo ComputeCurrentUMCStatsErrorHandler

    With GelUMC(CallerID)
        mUMCStats.Count = .UMCCnt
        If .UMCCnt > 0 Then
            
            ReDim dblMassWidthsPPM(0 To .UMCCnt - 1)
            ReDim dblMassWidthsDa(0 To .UMCCnt - 1)
            
            ReDim dblMassStDevPPM(0 To .UMCCnt - 1)
            ReDim dblMassStDevDa(0 To .UMCCnt - 1)
                          
            
            For lngIndex = 0 To .UMCCnt - 1
                With .UMCs(lngIndex)
                    dblMassWidthsDa(lngIndex) = .MaxMW - .MinMW
                    dblMassWidthsPPM(lngIndex) = MassToPPM(dblMassWidthsDa(lngIndex), .ClassMW)
                    
                    dblMassStDevDa(lngIndex) = .ClassMWStD
                    dblMassStDevPPM(lngIndex) = MassToPPM(dblMassStDevDa(lngIndex), .ClassMW)
                End With
            Next lngIndex
        End If
    End With
    
    With mUMCStats
        If .Count > 0 Then
            With .PPMStats
                objStats.Fill dblMassWidthsPPM()
                .MassWidthMedian = objStats.Median
                .MassWidthMaximum = objStats.Maximum
                
                objStats.Fill dblMassStDevPPM()
                .MassStDevMedian = objStats.Median
                .MassStDevMaximum = objStats.Maximum
            End With
            
            With .DaStats
                objStats.Fill dblMassWidthsDa()
                .MassWidthMedian = objStats.Median
                .MassWidthMaximum = objStats.Maximum
                
                objStats.Fill dblMassStDevDa()
                .MassStDevMedian = objStats.Median
                .MassStDevMaximum = objStats.Maximum
            End With
            
        Else
            With .PPMStats
                .MassWidthMedian = 0
                .MassWidthMaximum = 0
                .MassStDevMedian = 0
                .MassStDevMaximum = 0
            End With
            .DaStats = .PPMStats
        End If
    End With
    
    DisplayCurrentUMCStats

    Exit Sub

ComputeCurrentUMCStatsErrorHandler:
    Debug.Print "Error in DisplayUMCStats: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "ComputeCurrentUMCStats"
    
End Sub

Private Sub DisplayCurrentDBSearchTolerances()
    
    Dim strTolerances As String
    
On Error GoTo DisplayCurrentDBSearchTolerancesErrorHandler

    With GelSearchDef(CallerID).AMTSearchOnUMCs
        strTolerances = "DB Mass Tolerance: " & .MWTol & " " & GetSearchToleranceUnitText(CInt(.TolType)) & vbCrLf
        strTolerances = strTolerances & "DB NET Tolerance: " & .NETTol & " NET"
    End With

    txtDBSearchMassTolerances = strTolerances
    
    Exit Sub
    
DisplayCurrentDBSearchTolerancesErrorHandler:
    Debug.Print "DisplayCurrentDBSearchTolerances: " & Err.Description
    Debug.Assert False
    
End Sub

Private Sub DisplayCurrentUMCStats()
    ' Note: Call ComputeCurrentUMCStats before calling this sub
    
    Dim strUMCStats As String, strUnits As String
    Dim udtUMCStatDetails As udtUMCStatsDetailsType
    Dim lngDigitsToroundTo As Long
    Dim lngMultiplier As Long
    
On Error GoTo DisplayCurrentUMCStatsErrorHandler

    strUMCStats = "Count: " & mUMCStats.Count & " UMC's"
    If cboErrorDisplayMode.ListIndex = mdmMassErrorDa Then
        udtUMCStatDetails = mUMCStats.DaStats
        strUnits = " mDa"       ' Millidaltons
        lngDigitsToroundTo = MASS_DA_ADJUSTMENT_PRECISION - 5
        If lngDigitsToroundTo < 0 Then lngDigitsToroundTo = 0
        lngMultiplier = 1000
    Else
        ' Show stats in ppm
        udtUMCStatDetails = mUMCStats.PPMStats
        strUnits = " " & GetSearchToleranceUnitText(gltPPM)
        lngDigitsToroundTo = MASS_PPM_ADJUSTMENT_PRECISION - 1
        lngMultiplier = 1
    End If
        
    If mUMCStats.Count > 0 Then
        With udtUMCStatDetails
            strUMCStats = strUMCStats & vbCrLf & "Median Width: " & Round(.MassWidthMedian * lngMultiplier, lngDigitsToroundTo) & strUnits
            strUMCStats = strUMCStats & vbCrLf & "Maximum Width: " & Round(.MassWidthMaximum * lngMultiplier, lngDigitsToroundTo) & strUnits
            strUMCStats = strUMCStats & vbCrLf & "Median StDev: " & Round(.MassStDevMedian * lngMultiplier, lngDigitsToroundTo) & strUnits
            strUMCStats = strUMCStats & vbCrLf & "Maximum StDev: " & Round(.MassStDevMaximum * lngMultiplier, lngDigitsToroundTo) & strUnits
        End With
    End If
    
    txtUMCMassStatistics = strUMCStats

    Exit Sub
    
DisplayCurrentUMCStatsErrorHandler:
    Debug.Print "DisplayCurrentUMCStats: " & Err.Description
    Debug.Assert False
    
End Sub

Private Function Elution(FN As Long, MinFN As Long, MaxFN As Long) As Double
'---------------------------------------------------
'this function does not care are we using NET or RT
'---------------------------------------------------
VarVals(1) = FN
VarVals(2) = MinFN
VarVals(3) = MaxFN
Elution = MyExprEva.ExprVal(VarVals())
End Function

Public Function ExportErrorsBinnedToClipboardOrFile(strFilePath As String, blnShowMessages As Boolean, blnVisibleDifferencesOnly As Boolean) As Long
    ' Returns 0 if success, the error code if an error

    Dim strErrors() As String
    Dim strTextToCopy As String
    
    Dim lngDigitsToRound As Long
    Dim strBinAsScientific As String
    Dim lngCharLoc As Long
    
    Dim OutFileNum As Integer
    Dim lngIndex As Long, lngOutputArrayCount As Long
    Dim lngCountAvailable As Long
    Dim sngErrorValue As Single
    
    Dim strHeaderSuffix As String
    Dim strYValues As String
    
    Dim intDisplayModeSaved As Integer
    
    ' Save the current value of cboErrorDisplayMode
    intDisplayModeSaved = cboErrorDisplayMode.ListIndex
    
    If blnVisibleDifferencesOnly Then
        Select Case intDisplayModeSaved
        Case mdmMassErrorDisplayModeConstants.mdmMassErrorPPM
            lngCountAvailable = mMassPPMErrors.BinnedCount
        Case mdmMassErrorDisplayModeConstants.mdmMassErrorDa
            lngCountAvailable = mMassDaErrors.BinnedCount
        Case mdmMassErrorDisplayModeConstants.mdmGanetError
            lngCountAvailable = mGanetErrors.BinnedCount
        Case Else
            lngCountAvailable = 0
        End Select
    Else
        lngCountAvailable = mMassPPMErrors.BinnedCount + mMassDaErrors.BinnedCount + mGanetErrors.BinnedCount
    End If
    
    If lngCountAvailable = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        ExportErrorsBinnedToClipboardOrFile = -1
        Exit Function
    End If

    
On Error GoTo ExportMassErrorsBinnedErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    strHeaderSuffix = vbTab & "Count" & vbTab
    If cChkBox(chkShowSmoothedData) Then
        strHeaderSuffix = strHeaderSuffix & "Smoothed_Count" & vbTab
    End If
    strHeaderSuffix = strHeaderSuffix & "Comment"
    
    ' Header row is strErrors(0)
    ' Data is from strErrors(1) to strErrors(mMassPPMErrors.BinnedCount + mGANETErrors.BinnedCount)
    ReDim strErrors(0 To lngCountAvailable + 6)
    lngOutputArrayCount = 0
    
    If Not blnVisibleDifferencesOnly Or intDisplayModeSaved = mdmMassErrorDisplayModeConstants.mdmMassErrorPPM Then
        ' 1. Fill strErrors() with the Mass Errors (PPM)
        If cboErrorDisplayMode.ListIndex <> mdmMassErrorDisplayModeConstants.mdmMassErrorPPM Then
            cboErrorDisplayMode.ListIndex = mdmMassErrorDisplayModeConstants.mdmMassErrorPPM
        End If
        
        strErrors(lngOutputArrayCount) = "MassErrorPPM" & strHeaderSuffix
        lngOutputArrayCount = lngOutputArrayCount + 1
        
        strBinAsScientific = Format(mMassPPMErrors.BinSize, "0E+00")
        lngCharLoc = InStr(strBinAsScientific, "E")
        If lngCharLoc > 0 Then
            strBinAsScientific = Mid(strBinAsScientific, lngCharLoc + 1)
            If Left(strBinAsScientific, 1) = "+" Then
                lngDigitsToRound = 1
            Else
                lngDigitsToRound = Abs(CLngSafe(strBinAsScientific)) + 1
            End If
        Else
            lngDigitsToRound = 1
        End If
        
        For lngIndex = 0 To mMassPPMErrors.BinnedCount - 1
            sngErrorValue = mMassPPMErrors.StartBin + lngIndex * mMassPPMErrors.BinSize
            strYValues = Trim(mMassPPMErrors.Binned(lngIndex))
            If cChkBox(chkShowSmoothedData) Then
                strYValues = strYValues & vbTab & Trim(mMassPPMErrors.SmoothedBins(lngIndex))
            End If
            
            strErrors(lngOutputArrayCount) = Round(sngErrorValue, lngDigitsToRound) & vbTab & strYValues & vbTab & LookupErrorBinComment(lngIndex, mMassCalErrorPeakCached)
            lngOutputArrayCount = lngOutputArrayCount + 1
        Next lngIndex
    End If
    
    If Not blnVisibleDifferencesOnly Or intDisplayModeSaved = mdmMassErrorDisplayModeConstants.mdmMassErrorDa Then
        ' 2. Fill strErrors() with the Mass errors (Da)
        If cboErrorDisplayMode.ListIndex <> mdmMassErrorDisplayModeConstants.mdmMassErrorDa Then
            cboErrorDisplayMode.ListIndex = mdmMassErrorDisplayModeConstants.mdmMassErrorDa
        End If
        
        strErrors(lngOutputArrayCount) = vbCrLf & "MassErrorDa" & strHeaderSuffix
        lngOutputArrayCount = lngOutputArrayCount + 1
        
        strBinAsScientific = Format(mMassDaErrors.BinSize, "0E+00")
        lngCharLoc = InStr(strBinAsScientific, "E")
        If lngCharLoc > 0 Then
            strBinAsScientific = Mid(strBinAsScientific, lngCharLoc + 1)
            If Left(strBinAsScientific, 1) = "+" Then
                lngDigitsToRound = 1
            Else
                lngDigitsToRound = Abs(CLngSafe(strBinAsScientific)) + 1
            End If
        Else
            lngDigitsToRound = 1
        End If
        
        For lngIndex = 0 To mMassDaErrors.BinnedCount - 1
            sngErrorValue = mMassDaErrors.StartBin + lngIndex * mMassDaErrors.BinSize
            strYValues = Trim(mMassDaErrors.Binned(lngIndex))
            If cChkBox(chkShowSmoothedData) Then
                strYValues = strYValues & vbTab & Trim(mMassDaErrors.SmoothedBins(lngIndex))
            End If
            
            strErrors(lngOutputArrayCount) = Round(sngErrorValue, lngDigitsToRound) & vbTab & strYValues & vbTab & LookupErrorBinComment(lngIndex, mMassCalErrorPeakCached)
            lngOutputArrayCount = lngOutputArrayCount + 1
        Next lngIndex
    End If
    
    
    If Not blnVisibleDifferencesOnly Or intDisplayModeSaved = mdmMassErrorDisplayModeConstants.mdmGanetError Then
        ' 3. Fill strErrors() with the GANET errors
        If cboErrorDisplayMode.ListIndex <> mdmMassErrorDisplayModeConstants.mdmGanetError Then
            cboErrorDisplayMode.ListIndex = mdmMassErrorDisplayModeConstants.mdmGanetError
        End If
        
        strErrors(lngOutputArrayCount) = vbCrLf & "NETError" & strHeaderSuffix
        lngOutputArrayCount = lngOutputArrayCount + 1
        
        strBinAsScientific = Format(mGanetErrors.BinSize, "0E+00")
        lngCharLoc = InStr(strBinAsScientific, "E")
        If lngCharLoc > 0 Then
            strBinAsScientific = Mid(strBinAsScientific, lngCharLoc + 1)
            If Left(strBinAsScientific, 1) = "+" Then
                lngDigitsToRound = 1
            Else
                lngDigitsToRound = Abs(CLngSafe(strBinAsScientific)) + 1
            End If
        Else
            lngDigitsToRound = 1
        End If
        
        For lngIndex = 0 To mGanetErrors.BinnedCount - 1
            sngErrorValue = mGanetErrors.StartBin + lngIndex * mGanetErrors.BinSize
            strYValues = Trim(mGanetErrors.Binned(lngIndex))
            If cChkBox(chkShowSmoothedData) Then
                strYValues = strYValues & vbTab & Trim(mGanetErrors.SmoothedBins(lngIndex))
            End If
            
            strErrors(lngOutputArrayCount) = Round(sngErrorValue, lngDigitsToRound) & vbTab & strYValues & vbTab & LookupErrorBinComment(lngIndex, mNETTolErrorPeakCached)
            lngOutputArrayCount = lngOutputArrayCount + 1
        Next lngIndex
        
    End If
    
    ' 4. Save to file or copy to clipboard
    If Len(strFilePath) > 0 Then
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        For lngIndex = 0 To lngOutputArrayCount - 1
            Print #OutFileNum, strErrors(lngIndex)
        Next lngIndex
        
        Close #OutFileNum
    Else
        strTextToCopy = FlattenStringArray(strErrors(), lngOutputArrayCount, vbCrLf, False)
        Clipboard.Clear
        Clipboard.SetText strTextToCopy, vbCFText
    End If
    
     ' Restore the value of cboErrorDisplayMode
    If cboErrorDisplayMode.ListIndex <> intDisplayModeSaved Then
        cboErrorDisplayMode.ListIndex = intDisplayModeSaved
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    ExportErrorsBinnedToClipboardOrFile = 0
    Exit Function


ExportMassErrorsBinnedErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting binned data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    ExportErrorsBinnedToClipboardOrFile = Err.Number
    
End Function

Public Sub RecordMassCalPeakStatsNow()
    glbPreferencesExpanded.AutoAnalysisCachedData.MassCalErrorPeakCached = mMassCalErrorPeakCached
End Sub

Public Sub RecordNETTolPeakStatsNow()
    glbPreferencesExpanded.AutoAnalysisCachedData.NETTolErrorPeakCached = mNETTolErrorPeakCached
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

Public Sub InitializeForm()
    UpdateStatus "Initializing"
    cmdRecomputeHistograms.Visible = False
    
    ' We're calling EnableDisableCalculating to disable the controls while intitializing the form
    '  but we must set mCalculating to false, lest some of the initialization events are skipped
    EnableDisableCalculating True, False
    mCalculating = False
    
    If GelAnalysis(CallerID) Is Nothing Then
       If AMTCnt > 0 Then    'something is loaded
          If (Len(CurrMTDatabase) > 0 Or Len(CurrLegacyMTDatabase) > 0) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
             'MT tag data; we dont know is it appropriate; warn user
             WarnUserUnknownMassTags CallerID
          End If
          lblMTStatus.Caption = ConstructMTStatusText(False)
       
          If Not mAMTIndicesInitialized Then
              ConstructLookupArrays
              mAMTIndicesInitialized = True
          End If
       
       Else                  'nothing is loaded
          lblMTStatus.Caption = "No MT tags loaded"
       
        ReDim mAMTIDSorted(0)
        ReDim mAMTIDSortedInd(0)

        ReDim mInternalStdIDSorted(0)
        ReDim mInternalStdIDSortedInd(0)
       
       End If
       
       ' Can't use the GelAnalysis() object to determine slope and intercept
       ' Need to use the expression evaluator instead
       
       'memorize number of scans (to be used with elution)
       MinFN = GelData(CallerID).ScanInfo(1).ScanNumber
       MaxFN = GelData(CallerID).ScanInfo(UBound(GelData(CallerID).ScanInfo)).ScanNumber
       
        If GelData(CallerID).CustomNETsDefined Then
            mGelAnalysisIsValid = True
        Else
            mGelAnalysisIsValid = False
            If Not InitExprEvaluator(GelUMCNETAdjDef(CallerID).NETFormula) Then
               InitExprEvaluator ConstructNETFormulaWithDefaults(UMCNetAdjDef)
            End If
        End If
       
    Else
       mGelAnalysisIsValid = True
       LoadMTDB False
    End If
    
    mUpdatingControls = True
    
    ' Update the controls with the values in .ErrorPlottingOptions
    With glbPreferencesExpanded.ErrorPlottingOptions
        txtMassRangePPM = .MassRangePPM
        txtMassBinSizePPM = .MassBinSizePPM
        txtGANETRange = .GANETRange
        txtGANETBinSize = .GANETBinSize
        
        txtButterworthFrequency = .ButterWorthFrequency
        
        With .Graph2DOptions
            SetCheckBox chkAutoScaleXRange, .AutoScaleXAxis
            SetCheckBox chkShowPointSymbols, .ShowPointSymbols
            SetCheckBox chkDrawLinesBetweenPoints, .DrawLinesBetweenPoints
            SetCheckBox chkShowGridlines, .ShowGridLines
            SetCheckBox chkCenterYAxis, .CenterYAxis
            SetCheckBox chkShowSmoothedData, .ShowSmoothedData
            SetCheckBox chkShowPeakEdges, .ShowPeakEdges

            txtGraphPointSize = .PointSizePixels
            txtGraphLineWidth = .LineWidthPixels
        End With
    End With
    
    With glbPreferencesExpanded.RefineMSDataOptions
        txtToleranceRefinementMinimumPeakHeight = Trim(.MinimumPeakHeight)
        txtToleranceRefinementPercentageOfMaxForWidth = Trim(.PercentageOfMaxForFindingWidth)
        txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks = Trim(.MinimumSignalToNoiseRatioForLowAbundancePeaks)
        
        txtRefineMassCalibrationMaximumShift = Trim(.MassCalibrationMaximumShift)
        If .MassCalibrationTolType = gltPPM Then
            optRefineMassCalibrationMassType(0).Value = True
        Else
            optRefineMassCalibrationMassType(1).Value = True
        End If
        cboMassToleranceRefinementMethod.ListIndex = .MassToleranceRefinementMethod
        txtRefineDBSearchMassToleranceMinimum = Trim(.MassToleranceMinimum)
        txtRefineDBSearchMassToleranceMaximum = Trim(.MassToleranceMaximum)
        txtRefineDBSearchMassToleranceAdjustmentMultiplier = Trim(.MassToleranceAdjustmentMultiplier)
        txtRefineDBSearchNETToleranceMinimum = Trim(.NETToleranceMinimum)
        txtRefineDBSearchNETToleranceMaximum = Trim(.NETToleranceMaximum)
        txtRefineDBSearchNETToleranceAdjustmentMultiplier = Trim(.NETToleranceAdjustmentMultiplier)
        SetCheckBox chkIncludeInternalStandards, .IncludeInternalStdMatches
        SetCheckBox chkUseUMCClassStats, .UseUMCClassStats
        txtMinimumSLiC = Trim(.MinimumSLiC)
        txtMaximumAbundance = Trim(.MaximumAbundance)
    End With
    
    With GelSearchDef(CallerID).MassCalibrationInfo
        If .AdjustmentHistoryCount > 0 Then
            If .MassUnits = gltPPM Then
                If optRefineMassCalibrationMassType(0).Value <> True Then
                    optRefineMassCalibrationMassType(0).Value = True
                End If
            Else
                If optRefineMassCalibrationMassType(1).Value <> True Then
                    optRefineMassCalibrationMassType(1).Value = True
                End If
            End If
        End If
    End With
    
    mUpdatingControls = False
    
    UpdateDynamicControls
    
    ComputeErrors True
    
    cmdRecomputeHistograms.Visible = True
    EnableDisableCalculating False
    
    mFormInitialized = True
    
End Sub

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean
    Dim blnDBConnectionError As Boolean
    Dim eResponse As VbMsgBoxResult
    
    If AMTCnt = 0 And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("MT tags not in memory.  Load from the database?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load MT tags")
    Else
        eResponse = vbYes
    End If
    
    If eResponse = vbYes Then
        If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, 0, blnForceReload, True, blnAMTsWereLoaded, blnDBConnectionError) Then
            lblMTStatus.Caption = ConstructMTStatusText(False)
        Else
            If blnDBConnectionError Then
                lblMTStatus.Caption = "Error loading MT tags: database connection error!"
            Else
                lblMTStatus.Caption = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
            End If
        End If
    Else
        If AMTCnt > 0 Then
            lblMTStatus.Caption = ConstructMTStatusText(False)
        Else
            lblMTStatus.Caption = "Using cached MT data"
        End If
    End If
    
    If Not mAMTIndicesInitialized Then
        ConstructLookupArrays
        mAMTIndicesInitialized = True
    End If

End Sub

Private Function LookupErrorBinComment(lngIndex As Long, ByRef udtErrorPlottingPeakCache As udtErrorPlottingPeakCacheType) As String

    Select Case lngIndex
        Case udtErrorPlottingPeakCache.PeakStats.IndexBaseLeft
            LookupErrorBinComment = "Left base"
        Case udtErrorPlottingPeakCache.PeakStats.IndexLeft
            LookupErrorBinComment = "Left at " & glbPreferencesExpanded.RefineMSDataOptions.PercentageOfMaxForFindingWidth & "% of Max"
        Case udtErrorPlottingPeakCache.PeakStats.IndexOfMaximum
            LookupErrorBinComment = "Maximum"
        Case udtErrorPlottingPeakCache.PeakStats.IndexOfCenterOfMass
            LookupErrorBinComment = "Center of mass"
        Case udtErrorPlottingPeakCache.PeakStats.IndexRight
            LookupErrorBinComment = "Right at " & glbPreferencesExpanded.RefineMSDataOptions.PercentageOfMaxForFindingWidth & "% of Max"
        Case udtErrorPlottingPeakCache.PeakStats.IndexBaseRight
            LookupErrorBinComment = "Right base"
        Case Else
            LookupErrorBinComment = ""
    End Select

End Function

Public Function ManualRefineMassCalibration(Optional blnOverrideValue As Boolean = False, Optional dblMassAdjustmentOverridePPM As Double = 0) As Boolean
    ' If blnOverrideValue = True, then dblMassAdjustmentOverride is used
    '  instead of the one given by txtMassCalibrationNewIncrementalAdjustment
    
    Dim dblNewMassAdjustmentIncrement As Double
    Dim eMassType As glMassToleranceConstants
    Dim blnSuccess As Boolean
    Dim blnProceed As Boolean
    
    If Not mFormInitialized Then
        ManualRefineMassCalibration = False
        Exit Function
    End If
    
    If blnOverrideValue Then
        If GelSearchDef(CallerID).MassCalibrationInfo.AdjustmentHistoryCount > 0 Then
            ' Undo any previous mass calibration adjustments when overriding auto adjustment
            EnableDisableCalculating True, False
            blnProceed = MassCalibrationRevertToOriginal(False, False, True, Me)
            If blnProceed Then
                UpdateUMCStatsAndRecomputeErrors
            End If
            EnableDisableCalculating False
        End If
        
        ' Make sure ppm mode is enabled
        optRefineMassCalibrationMassType(0).Value = True
        
        ' Verify that the mode changed; if previous adjustments were made with Da mode, then it will not change
        If optRefineMassCalibrationMassType(0).Value = True Then
            dblNewMassAdjustmentIncrement = dblMassAdjustmentOverridePPM
        Else
            ManualRefineMassCalibration = False
            Exit Function
        End If
    Else
        dblNewMassAdjustmentIncrement = CDblSafe(txtMassCalibrationNewIncrementalAdjustment)
    End If
    
    If dblNewMassAdjustmentIncrement = 0 Then
        ManualRefineMassCalibration = True
        Exit Function
    End If
    
    If optRefineMassCalibrationMassType(0).Value = True Then
        eMassType = gltPPM
    Else
        eMassType = gltABS
    End If
    
    EnableDisableCalculating True, False
    
    blnSuccess = MassCalibrationApplyBulkAdjustment(CallerID, dblNewMassAdjustmentIncrement, eMassType, True, 0, Me)
    
    EnableDisableCalculating False
    
    If blnSuccess Then
        UpdateUMCStatsAndRecomputeErrors
    End If

    
    ManualRefineMassCalibration = blnSuccess

End Function

Private Sub PopulateComboBoxes()
    mUpdatingControls = True
    
    With cboErrorDisplayMode
        .Clear
        .AddItem "Mass Error (PPM)"
        .AddItem "Mass Error (Da)"
        .AddItem "NET Error"
        .ListIndex = mdmMassErrorDisplayModeConstants.mdmMassErrorPPM
    End With
    
    With cboMassToleranceRefinementMethod
        .Clear
        .AddItem "Mass Error Plot Width at % of Max"
        .AddItem "Median UMC Mass StDev"
        .AddItem "Maximum UMC Mass StDev"
        .AddItem "Median UMC Mass Width"
        .AddItem "Maximum UMC Mass Width"
        .ListIndex = mtrMassErrorPlotWidthAtPctOfMax
    End With
    
    mUpdatingControls = False
End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    Dim dblToleranceControlsVisible As Boolean
    
    dblToleranceControlsVisible = cChkBox(chkShowToleranceRefinementControls)
    
    fraToleranceRefinementContainer.Visible = dblToleranceControlsVisible
    
    With fraControlsAndPlotContainer
        If dblToleranceControlsVisible Then
            .Left = fraToleranceRefinementContainer.Left + fraToleranceRefinementContainer.width
        Else
            .Left = 120
        End If
        
        lngDesiredValue = Me.ScaleWidth - .Left
        If lngDesiredValue < 4000 Then lngDesiredValue = 4000
        .width = lngDesiredValue
        
        lngDesiredValue = Me.ScaleHeight
        If lngDesiredValue < 4000 Then lngDesiredValue = 4000
        .Height = lngDesiredValue
        
    End With
    
    With ctlPlotter
        .Top = lblStatus.Top + lblStatus.Height
        .Left = 0
        
        .width = fraControlsAndPlotContainer.width
        lngDesiredValue = fraControlsAndPlotContainer.Height - .Top
        If lngDesiredValue < 1000 Then
            ' This shouldn't happen
            Debug.Assert False
            lngDesiredValue = 1000
        End If
        .Height = lngDesiredValue
    End With
    
    fraOptions.Top = ctlPlotter.Top
    fraOptions.Left = ctlPlotter.Left
End Sub

Public Function RefineMassCalibrationStart(Optional ByRef blnValidPeakFound As Boolean, Optional ByRef blnMassShiftTooLarge As Boolean = False, Optional ByRef blnPeakTooWide As Boolean = False, Optional ByVal blnFindPeakOnly As Boolean) As Boolean
    ' Returns True if success, False if failure
    ' If a valid peak is not found, sets blnValidPeakFound = False, but returns True
    ' If blnFindPeakOnly is True, then looks for a peak and updates the analysis history with the peak stats, but does not shift the peak
    
    Dim eMassType As glMassToleranceConstants
    Dim udtPeak As udtPeakStatsType
    
    Dim blnSuccess As Boolean
    Dim blnSingleGoodPeakFound As Boolean
    
    Dim strMessage As String
    
On Error GoTo RefineMassCalibrationStartErrorHandler

    If Not mFormInitialized Then Exit Function
    
    EnableDisableCalculating True, False
    eMassType = glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationTolType
    If eMassType = gltPPM Then
        blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mMassPPMErrors, udtPeak, blnSingleGoodPeakFound)
    Else
        blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mMassDaErrors, udtPeak, blnSingleGoodPeakFound)
    End If
    
    If blnValidPeakFound Then
        
        ' Update txtMassCalibrationNewIncrementalAdjustment to be 0 so that the user doesn't accidentally shift the data further
        txtMassCalibrationNewIncrementalAdjustment = Trim(0)
        
        If eMassType = gltPPM Then
            blnSuccess = RefineMassCalibrationWork(CallerID, mMassPPMErrors, udtPeak, eMassType, MASS_PPM_ADJUSTMENT_PRECISION, 0, blnSingleGoodPeakFound, blnMassShiftTooLarge, blnPeakTooWide, True, blnFindPeakOnly, True, Me)
        Else
            blnSuccess = RefineMassCalibrationWork(CallerID, mMassDaErrors, udtPeak, eMassType, MASS_DA_ADJUSTMENT_PRECISION, 0, blnSingleGoodPeakFound, blnMassShiftTooLarge, blnPeakTooWide, True, blnFindPeakOnly, True, Me)
        End If
        
        If blnSuccess And Not blnFindPeakOnly Then
            UpdateUMCStatsAndRecomputeErrors
        End If
        
    Else
        strMessage = "Unable to determine a mass calibration adjustment value since no valid peaks could be found in the mass error plot"
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strMessage & ".", vbExclamation + vbOKOnly, "Valid Peak Not Found"
        Else
            With glbPreferencesExpanded.RefineMSDataOptions
                strMessage = "Warning - " & strMessage & "; Minimum peak height = " & .MinimumPeakHeight & " counts/bin; Percentage of peak max for finding peak width = " & .PercentageOfMaxForFindingWidth & "%; Minimum SLiC = " & .MinimumSLiC & "; Minimum Signal/Noise = " & Trim(.MinimumSignalToNoiseRatioForLowAbundancePeaks)
                If .MaximumAbundance > 0 Then
                    strMessage = strMessage & "; Maximum Abundance = " & Trim(.MaximumAbundance)
                End If
            End With

            AddToAnalysisHistory CallerID, strMessage
        End If
        
        ' Set blnSuccess to True since this is a warning, not an error
        blnSuccess = True
    End If
    
    EnableDisableCalculating False
    RefineMassCalibrationStart = blnSuccess
    Exit Function

RefineMassCalibrationStartErrorHandler:
    Debug.Print "Error in RefineMassCalibrationStart: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "RefineMassCalibrationStart"
    EnableDisableCalculating False

    
End Function

Private Function RefineMassCalibrationWork(lngGelIndex As Long, ByRef udtBinnedErrorData As udtBinnedDataType, ByRef udtPeak As udtPeakStatsType, ByVal eMassType As glMassToleranceConstants, ByVal lngDigitsOfPrecisionToRoundTo As Long, ByVal dblPreAppliedMassShiftPPM As Double, ByVal blnSingleGoodPeakFound As Boolean, ByRef blnMassShiftTooLarge As Boolean, ByRef blnPeakTooWide As Boolean, blnMakeLogEntry As Boolean, blnFindPeakOnly As Boolean, blnInformIfChangeTooSmall As Boolean, frmCallingForm As VB.Form)
    ' Looks for a peak in udtBinnedErrorData, populating udtPeak with the peak stats
    ' If a valid peak is found, then considers shifting the data to move the peak to be centered at 0
    
    Dim dblPeakCenter As Double, dblPeakWidth As Double, dblPeakHeight As Double
    Dim sngSignalToNoise As Single
    Dim sngBinSize As Single
    
    Dim blnSuccess As Boolean
    
    Dim blnChangeTooSmall As Boolean
    Dim dblMassAdjustmentIncrement As Double
    
    Dim strMessage As String
    Dim strMassCalibrationPeakStats As String

    GetPeakStats udtBinnedErrorData, udtPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, lngDigitsOfPrecisionToRoundTo
    strMassCalibrationPeakStats = MASS_CALIBRATION_PEAK_STATS_START & " = " & dblPeakHeight & ", " & dblPeakWidth & ", " & dblPeakCenter & ", " & sngSignalToNoise & " " & MASS_CALIBRATION_PEAK_STATS_END

    blnChangeTooSmall = False
    dblMassAdjustmentIncrement = 0
    
    ' Make sure dblPeakWidth is less than .MassToleranceMaximum
    ' If it isn't, then don't perform mass calibration since we run the risk of shifting the data outside of the search window
    If dblPeakWidth > glbPreferencesExpanded.RefineMSDataOptions.MassToleranceMaximum Then
        strMessage = "The mass tolerance determined from the mass error plot was "
        If dblPeakWidth < 1 Then
            strMessage = strMessage & Format(dblPeakWidth, "0.0000")
        Else
            strMessage = strMessage & Format(dblPeakWidth, "0.0")
        End If
        strMessage = strMessage & " " & GetSearchToleranceUnitText(eMassType) & "; This value is larger than the defined maximum peak width for mass tolerance refinement.  Thus, no mass adjustment was performed."
        
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "Peak width too large"
        Else
            If Right(strMessage, 1) = "." Then
                strMessage = Left(strMessage, Len(strMessage) - 1)
            Else
                ' strMessage should always end in a period
                Debug.Assert False
            End If
            strMessage = strMessage & "; " & strMassCalibrationPeakStats
            AddToAnalysisHistory lngGelIndex, strMessage
        End If
        
        ' This is a warning, not an error
        ' Set blnSuccess to True and blnPeakTooWide to True
        blnSuccess = True
        blnPeakTooWide = True
    Else
        blnSuccess = True
        blnPeakTooWide = False
    End If

    If Not blnPeakTooWide Then
        With udtBinnedErrorData
            sngBinSize = .BinSize
            
            dblMassAdjustmentIncrement = -(.StartBin + udtPeak.IndexOfCenterOfMass * .BinSize)
            
            If eMassType = gltPPM And dblPreAppliedMassShiftPPM <> 0 Then
                dblMassAdjustmentIncrement = dblMassAdjustmentIncrement + dblPreAppliedMassShiftPPM
            End If
            
            dblMassAdjustmentIncrement = Round(dblMassAdjustmentIncrement, lngDigitsOfPrecisionToRoundTo)
            
            If (eMassType = gltPPM And Abs(dblMassAdjustmentIncrement) < 0.1) Or _
               (eMassType <> gltPPM And Abs(dblMassAdjustmentIncrement) < 0.0001) Then
                ' No point in shifting; the change is too small
                blnChangeTooSmall = True
            End If
        End With
        
        
        If blnChangeTooSmall Then
            strMessage = "The mass calibration adjustment value computed using the mass error plot was nearly 0.  Thus, no mass adjustment was performed."
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled And blnInformIfChangeTooSmall Then
                MsgBox strMessage, vbExclamation + vbOKOnly, "Adjustment Too Small"
            Else
                If Right(strMessage, 1) = "." Then
                    strMessage = Left(strMessage, Len(strMessage) - 1)
                Else
                    ' strMessage should always end in a period
                    Debug.Assert False
                End If
                strMessage = strMessage & "; " & strMassCalibrationPeakStats
                AddToAnalysisHistory lngGelIndex, strMessage
            End If
                
            ' This is neither an error nor a warning; set blnSuccess to True
            blnSuccess = True
        Else
            ' See if larger than the maximum allowed shift
            If Abs(dblMassAdjustmentIncrement) > glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationMaximumShift Then
                strMessage = "The mass calibration adjustment value computed using the mass error plot was too large (" & dblMassAdjustmentIncrement & " " & GetSearchToleranceUnitText(eMassType) & " vs. limit of " & Trim(glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationMaximumShift) & ").  Thus, no mass adjustment was performed."
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                    MsgBox strMessage, vbExclamation + vbOKOnly, "Adjustment Exceeds Maximum"
                Else
                    If Right(strMessage, 1) = "." Then
                        strMessage = Left(strMessage, Len(strMessage) - 1)
                    Else
                        ' strMessage should always end in a period
                        Debug.Assert False
                    End If
                    strMessage = strMessage & "; " & strMassCalibrationPeakStats
                    strMessage = "Error - " & strMessage
                    AddToAnalysisHistory lngGelIndex, strMessage
                    blnMassShiftTooLarge = True
                End If
                
                ' Flag this as an error, since we probably want to re-analyze this file; set blnSuccess to false
                blnSuccess = False
            Else
                With glbPreferencesExpanded.RefineMSDataOptions
                    If blnFindPeakOnly Then
                        strMessage = "Examined mass calibration"
                    Else
                        strMessage = "Adjusted mass calibration"
                    End If
                    strMessage = strMessage & " using the peak identified on the mass error plot; " & strMassCalibrationPeakStats & "; Percentage of peak max for finding peak width = " & .PercentageOfMaxForFindingWidth & "%; Minimum SLiC = " & .MinimumSLiC
                    
                    If .MaximumAbundance > 0 Then
                        strMessage = strMessage & "; Maximum Abundance = " & Trim(.MaximumAbundance)
                    End If
                    strMessage = strMessage & "; Single good peak found = " & Trim(blnSingleGoodPeakFound)
                End With
                AddToAnalysisHistory lngGelIndex, strMessage
                
                If blnFindPeakOnly Then
                    blnSuccess = True
                Else
                    blnSuccess = MassCalibrationApplyBulkAdjustment(lngGelIndex, dblMassAdjustmentIncrement, eMassType, blnMakeLogEntry, sngBinSize, frmCallingForm)
                End If
            
            End If
        End If
    End If
    
    RefineMassCalibrationWork = blnSuccess
End Function

Public Function RefineDBSearchMassToleranceStart(Optional ByRef blnValidPeakFound As Boolean, Optional ByRef blnPeakTooWide As Boolean) As Boolean
    
    ' Note that blnValidPeakFound is will be set to True even when the .MassToleranceRefinementMethod
    '  is not mtrMassErrorPlotWidthAtPctOfMax; this is done since the other methods rely on UMC statistics
    ' It will only be set to False if the UMC stats are empty
    
    Dim udtPeak As udtPeakStatsType
    
    Dim eMassType As glMassToleranceConstants
    Dim dblPeakCenter As Double, dblPeakWidth As Double, dblPeakHeight As Double
    Dim sngSignalToNoise As Single
    Dim strMethodName As String, strMessage As String
    Dim strPeakOptionsMsg As String
    Dim strFilterOptionsMsg As String
    
    Dim blnSuccess As Boolean
    Dim blnUsingMassErrorPlot As Boolean
    Dim blnSingleGoodPeakFound As Boolean
    
On Error GoTo RefineDBSearchMassToleranceStartErrorHandler

    If Not mFormInitialized Then Exit Function
    
    EnableDisableCalculating True, False
    blnPeakTooWide = False
    
    ' Set eMassType; this is not changed below
    eMassType = gltPPM
    
    ' Set this to true in case the .MassToleranceRefinementMethod is not mtrMassErrorPlotWidthAtPctOfMax
    ' This only affects the setting of the error bit during auto analysis
    blnValidPeakFound = True
    
    Select Case glbPreferencesExpanded.RefineMSDataOptions.MassToleranceRefinementMethod
    Case mtrMassErrorPlotWidthAtPctOfMax
        ' Adjust using the mass error plot
        If eMassType = gltPPM Then
            blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mMassPPMErrors, udtPeak, blnSingleGoodPeakFound)
        Else
            blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mMassDaErrors, udtPeak, blnSingleGoodPeakFound)
        End If
        
        strMethodName = "mass error plot"
        blnUsingMassErrorPlot = True
        
        If blnValidPeakFound Then
            If eMassType = gltPPM Then
                GetPeakStats mMassPPMErrors, udtPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, MASS_PPM_ADJUSTMENT_PRECISION
            Else
                GetPeakStats mMassDaErrors, udtPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, MASS_DA_ADJUSTMENT_PRECISION
            End If
        End If
    
    Case mtrMedianUMCMassStDev
        dblPeakWidth = Round(mUMCStats.PPMStats.MassStDevMedian, MASS_PPM_ADJUSTMENT_PRECISION)
        If mUMCStats.Count > 0 Then blnValidPeakFound = True
        strMethodName = "median UMC mass StDev"
        
    Case mtrMaximumUMCMassStDev
        dblPeakWidth = Round(mUMCStats.PPMStats.MassStDevMaximum, MASS_PPM_ADJUSTMENT_PRECISION)
        If mUMCStats.Count > 0 Then blnValidPeakFound = True
        strMethodName = "maximum UMC mass StDev"
    
    Case mtrMedianUMCMassWidth
        dblPeakWidth = Round(mUMCStats.PPMStats.MassWidthMedian, MASS_PPM_ADJUSTMENT_PRECISION)
        If mUMCStats.Count > 0 Then blnValidPeakFound = True
        strMethodName = "median UMC mass width"
    
    Case mtrMaximumUMCMassWidth
        dblPeakWidth = Round(mUMCStats.PPMStats.MassWidthMaximum, MASS_PPM_ADJUSTMENT_PRECISION)
        If mUMCStats.Count > 0 Then blnValidPeakFound = True
        strMethodName = "maximum UMC mass width"
        
    Case Else
        ' Invalid (or un-coded) method
        Debug.Assert False
    End Select
    
    With glbPreferencesExpanded.RefineMSDataOptions
        strPeakOptionsMsg = "Percentage of peak max for finding peak width = " & .PercentageOfMaxForFindingWidth & "%"
        
        strFilterOptionsMsg = "Use UMC class stats = " & Trim(.UseUMCClassStats)
        strFilterOptionsMsg = strFilterOptionsMsg & "; Minimum SLiC Score = " & Trim(.MinimumSLiC)
        If .MaximumAbundance > 0 Then
            strFilterOptionsMsg = strFilterOptionsMsg & "; Maximum Abundance = " & Trim(.MaximumAbundance)
        End If
    End With
    
    If blnValidPeakFound Then
        
        With glbPreferencesExpanded.RefineMSDataOptions
        
            ' Multiply dblPeakWidth by .MassToleranceAdjustmentMultiplier
            If .MassToleranceAdjustmentMultiplier <= 0 Then
                Debug.Assert False
                .MassToleranceAdjustmentMultiplier = 1
            End If
            
            dblPeakWidth = dblPeakWidth * .MassToleranceAdjustmentMultiplier
            
            ' Make sure dblPeakWidth is within the minimum and maximum limits defined
            If dblPeakWidth < .MassToleranceMinimum Or dblPeakWidth > .MassToleranceMaximum Then
                strMessage = "The mass tolerance determined from the " & strMethodName & " was " & dblPeakWidth & " " & GetSearchToleranceUnitText(eMassType) & "; This value is outside of the defined limits for mass tolerance refinement.  Thus, the mass tolerance was not changed."
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                    MsgBox strMessage, vbInformation + vbOKOnly, "Outside Limits"
                Else
                    AddToAnalysisHistory CallerID, strMessage
                End If
                
                ' This is a warning, not an error
                ' Set blnSuccess to True and blnPeakTooWide to True
                blnSuccess = True
                If dblPeakWidth > .MassToleranceMaximum Then blnPeakTooWide = True
            Else
                blnSuccess = True
            End If
        End With
        
        If blnSuccess And Not blnPeakTooWide Then
            With GelSearchDef(CallerID)
                .AMTSearchOnIons.MWTol = dblPeakWidth
                .AMTSearchOnIons.TolType = eMassType
                
                .AMTSearchOnUMCs.MWTol = dblPeakWidth
                .AMTSearchOnUMCs.TolType = eMassType
                
                .AMTSearchOnPairs.MWTol = dblPeakWidth
                .AMTSearchOnPairs.TolType = eMassType
                
                samtDef = .AMTSearchOnUMCs
            End With
            
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "The mass tolerance determined using the " & strMethodName & " was " & dblPeakWidth & " " & GetSearchToleranceUnitText(eMassType) & "; the search tolerance has been updated.", vbInformation + vbOKOnly, "Tolerance Updated"
            End If
        
            strMessage = SEARCH_MASS_TOL_DETERMINED & " using " & strMethodName & "; tolerance = " & dblPeakWidth & " " & GetSearchToleranceUnitText(eMassType)
            If blnUsingMassErrorPlot Then
                strMessage = strMessage & "; " & strPeakOptionsMsg
            End If
            strMessage = strMessage & "; " & strFilterOptionsMsg
                            
            strMessage = strMessage & "; Single good peak found = " & CStr(blnSingleGoodPeakFound)
            
            AddToAnalysisHistory CallerID, strMessage
            
        End If
        
    Else
        If blnUsingMassErrorPlot Then
            strMessage = "Unable to determine an optimal mass tolerance since no valid peaks could be found in the mass error plot"
        Else
            strMessage = "Unable to determine an optimal mass tolerance since there are no UMC's in memory"
        End If
        
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strMessage & ".", vbExclamation + vbOKOnly, "Adjustment Not Performed"
        Else
            If blnUsingMassErrorPlot Then
                strMessage = strMessage & "; Minimum peak height = " & glbPreferencesExpanded.RefineMSDataOptions.MinimumPeakHeight & " counts/bin; " & strPeakOptionsMsg
            End If
            strMessage = strMessage & "; " & strFilterOptionsMsg
            
            strMessage = "Warning - " & strMessage
            AddToAnalysisHistory CallerID, strMessage
        End If
        
        ' Set blnSuccess to True since this is a warning, not an error
        blnSuccess = True
    End If
    
    DisplayCurrentDBSearchTolerances
    
    EnableDisableCalculating False
    
    RefineDBSearchMassToleranceStart = blnSuccess
    Exit Function

RefineDBSearchMassToleranceStartErrorHandler:
    Debug.Print "Error in RefineDBSearchMassToleranceStart: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "RefineDBSearchMassToleranceStart"
    EnableDisableCalculating False

End Function

Public Function RefineDBSearchNETToleranceStart(Optional ByRef blnValidPeakFound As Boolean, Optional ByRef blnPeakTooWide As Boolean) As Boolean
    
    Dim udtPeak As udtPeakStatsType
    
    Dim dblPeakCenter As Double, dblPeakWidth As Double, dblPeakHeight As Double
    Dim sngSignalToNoise As Single
    Dim strMessage As String
    Dim strPeakOptionsMsg As String
    Dim strFilterOptionsMsg As String
    
    Dim blnSuccess As Boolean
    Dim blnSingleGoodPeakFound As Boolean
    
On Error GoTo RefineDBSearchNETToleranceStartErrorHandler

    If Not mFormInitialized Then Exit Function
    
    EnableDisableCalculating True, False
    blnPeakTooWide = False
    
    ' Adjust using the NET error plot
    blnValidPeakFound = FindPeakStatsUsingBinnedErrorData(mGanetErrors, udtPeak, blnSingleGoodPeakFound)
    If blnValidPeakFound Then
        GetPeakStats mGanetErrors, udtPeak, dblPeakCenter, dblPeakWidth, dblPeakHeight, sngSignalToNoise, GANET_ADJUSTMENT_PRECISION
    End If
    
    With glbPreferencesExpanded.RefineMSDataOptions
        strPeakOptionsMsg = "Percentage of peak max for finding peak width = " & .PercentageOfMaxForFindingWidth & "%"
    
        strFilterOptionsMsg = "Use UMC class stats = " & Trim(.UseUMCClassStats)
        strFilterOptionsMsg = strFilterOptionsMsg & "; Minimum SLiC Score = " & Trim(.MinimumSLiC)
        If .MaximumAbundance > 0 Then
            strFilterOptionsMsg = strFilterOptionsMsg & "; Maximum Abundance = " & Trim(.MaximumAbundance)
        End If
    End With
    
    If blnValidPeakFound Then
        
        With glbPreferencesExpanded.RefineMSDataOptions
        
            ' Multiply dblPeakWidth by .NETToleranceAdjustmentMultiplier
            If .NETToleranceAdjustmentMultiplier <= 0 Then
                Debug.Assert False
                .NETToleranceAdjustmentMultiplier = 1
            End If

            dblPeakWidth = dblPeakWidth * .NETToleranceAdjustmentMultiplier
            
            ' Make sure dblPeakWidth is within the minimum and maximum limits defined
            If dblPeakWidth < .NETToleranceMinimum Or dblPeakWidth > .NETToleranceMaximum Then
                blnSuccess = False
                
                strMessage = "The NET tolerance determined from the NET error plot was " & dblPeakWidth & " NET; This value is outside of the defined limits for NET tolerance refinement.  Thus, the NET tolerance was not changed."
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                    MsgBox strMessage, vbInformation + vbOKOnly, "Outside Limits"
                Else
                    AddToAnalysisHistory CallerID, strMessage
                End If
                
                ' This is a warning, not an error
                ' Set blnSuccess to True and blnPeakTooWide to True
                blnSuccess = True
                blnPeakTooWide = True
            Else
                blnSuccess = True
            End If
        End With
        
        If blnSuccess And Not blnPeakTooWide Then
            With GelSearchDef(CallerID)
                .AMTSearchOnIons.NETTol = dblPeakWidth
                .AMTSearchOnUMCs.NETTol = dblPeakWidth
                .AMTSearchOnPairs.NETTol = dblPeakWidth
                samtDef = .AMTSearchOnUMCs
            End With
            
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "The NET tolerance determined using the NET error plot was " & dblPeakWidth & " NET; The search tolerance has been updated.", vbInformation + vbOKOnly, "Tolerance Updated"
            End If
        
            strMessage = SEARCH_NET_TOL_DETERMINED & " using the NET error plot; tolerance = " & dblPeakWidth & " NET"
            strMessage = strMessage & "; " & NET_TOL_PEAK_STATS_START & " = " & dblPeakHeight & ", " & dblPeakWidth & ", " & dblPeakCenter & ", " & sngSignalToNoise & " " & NET_TOL_PEAK_STATS_END

            strMessage = strMessage & "; " & strPeakOptionsMsg & "; " & strFilterOptionsMsg
            strMessage = strMessage & "; Single good peak found = " & CStr(blnSingleGoodPeakFound)
            AddToAnalysisHistory CallerID, strMessage
            
        End If
        
    Else
        strMessage = "Unable to determine an optimal NET tolerance since no valid peaks could be found in the NET error plot."
        
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "Adjustment Not Performed"
        Else
            strMessage = strMessage & "; Minimum peak height = " & glbPreferencesExpanded.RefineMSDataOptions.MinimumPeakHeight & " counts/bin; " & strPeakOptionsMsg & "; " & strFilterOptionsMsg
            strMessage = "Warning - " & strMessage
            AddToAnalysisHistory CallerID, strMessage
        End If
        
        ' Set blnSuccess to True since this is a warning, not an error
        blnSuccess = True
    End If
    
    DisplayCurrentDBSearchTolerances
    
    EnableDisableCalculating False
    
    RefineDBSearchNETToleranceStart = blnSuccess
    Exit Function

RefineDBSearchNETToleranceStartErrorHandler:
    Debug.Print "Error in RefineDBSearchNETToleranceStart: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "RefineDBSearchNETToleranceStart"
    EnableDisableCalculating False
End Function

Public Function SaveChartPictureToFile(blnSaveAsPNG As Boolean, Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    ' If blnSaveAsPNG = True, then saves a PNG file
    ' If blnSaveAsPNG = False, then saves a JPG file
    
    ' Returns 0 if success, the error code if an error
    
    Dim strPictureFormat As String
    Dim strPictureExtension As String
    Dim objRemoteSaveFileHandler As clsRemoteSaveFileHandler
    Dim strWorkingFilePath As String
    Dim blnSuccess As Boolean

On Error GoTo SaveChartPictureToFileErrorHandler

    If blnSaveAsPNG Then
        strPictureFormat = "PNG"
        strPictureExtension = ".png"
    Else
        strPictureFormat = "JPG"
        strPictureExtension = ".jpg"
    End If
    
    If Len(strFilePath) = 0 Then
        strFilePath = SelectFile(Me.hwnd, "Enter filename", "", True, "MassErrors" & strPictureExtension, strPictureFormat & " Files (*." & strPictureExtension & ")|*." & strPictureExtension & "|All Files (*.*)|*.*")
    End If
    
    If Len(strFilePath) > 0 Then
        strFilePath = FileExtensionForce(strFilePath, strPictureExtension)
        
        Set objRemoteSaveFileHandler = New clsRemoteSaveFileHandler
        strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
        
        If ctlPlotter.SaveChartPictureToFile(blnSaveAsPNG, strWorkingFilePath) Then
            blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
            SaveChartPictureToFile = 0
        Else
            SaveChartPictureToFile = -1
        End If
    Else
        SaveChartPictureToFile = -1
    End If
    
    Exit Function

SaveChartPictureToFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error while saving a picture of the graph to disk:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    SaveChartPictureToFile = Err.Number
    
End Function

Public Sub SetPlotMode(ePlotMode As mdmMassErrorDisplayModeConstants)
    On Error GoTo SetPlotModeErrorHandler
    
    cboErrorDisplayMode.ListIndex = ePlotMode
    Exit Sub

SetPlotModeErrorHandler:
    Debug.Print "Invalid plot mode in SetPlotMode: " & ePlotMode
    Debug.Assert False
    
End Sub

Private Sub SetToleranceRefinementOptionsToDefault()
    ResetExpandedPreferences glbPreferencesExpanded, "RefineMSDataOptions"
    ResetExpandedPreferences glbPreferencesExpanded, "ErrorPlottingOptions"
    InitializeForm
End Sub

Public Function ShowHideToleranceRefinementControls(blnShowRefinementControls As Boolean)
    If Not cChkBox(chkShowToleranceRefinementControls) = blnShowRefinementControls Then
        SetCheckBox chkShowToleranceRefinementControls, blnShowRefinementControls
    End If
    PositionControls
End Function

Public Function ShowOrCompute3DErrorDistributions(blnShowCumulativeData As Boolean, Optional blnShowForm As Boolean = True) As Long

    ' Returns 0 if success, the error code if an error

    Dim sngMassBinSizePPM As Single
    Dim sngGANETBinSize As Single
    Dim dblPredictedBinCount As Double
    
On Error GoTo ShowOrCompute3DErrorDistribErrorHandler

    If mErrorsCount > 0 Then
        With glbPreferencesExpanded.ErrorPlottingOptions
            If .MassBinSizePPM <= 0 Then .MassBinSizePPM = DEFAULT_MASS_BIN_SIZE_PPM
            If .GANETBinSize <= 0 Then .GANETBinSize = DEFAULT_GANET_BIN_SIZE
            
            ' Adjust the bin size if more than 100 bins will be needed
            sngMassBinSizePPM = .MassBinSizePPM
            Do
                dblPredictedBinCount = ((mMassPPMErrors.BinRangeMaximum - mMassPPMErrors.StartBin) / sngMassBinSizePPM) + 2
                If dblPredictedBinCount > 100 Then
                    sngMassBinSizePPM = sngMassBinSizePPM * 2
                Else
                    Exit Do
                End If
            Loop
            
            sngGANETBinSize = .GANETBinSize
            Do
                dblPredictedBinCount = ((mGanetErrors.BinRangeMaximum - mGanetErrors.StartBin) / sngGANETBinSize) + 2
                If dblPredictedBinCount > 100 Then
                    sngGANETBinSize = sngGANETBinSize * 2
                Else
                    Exit Do
                End If
            Loop
            
            ShowOrCompute3DErrorDistributions = frmErrorDistribution3DFromFile.InitializeDataUsingArrays(mRawMassErrorsPPM(), mRAWGanetErrors(), mErrorsCount, sngMassBinSizePPM, sngGANETBinSize, 200, mGraphTitle, blnShowCumulativeData, Me, blnShowForm)
            
        End With
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data in memory.  Nothing to graph.", vbInformation + vbOKOnly, "No data"
        End If
        ShowOrCompute3DErrorDistributions = -1
    End If
    
    Exit Function
    
ShowOrCompute3DErrorDistribErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in ShowOrCompute3DErrorDistributions:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    ShowOrCompute3DErrorDistributions = Err.Number
    
End Function

Private Sub ShowHideOptions(Optional blnForceHide As Boolean)
    If blnForceHide Then
        fraOptions.Visible = False
    Else
        fraOptions.Visible = Not fraOptions.Visible
    End If
    
    mnuViewOptions.Checked = fraOptions.Visible
End Sub

Private Sub StartMassCalibrationRevert()
    Dim blnSuccess As Boolean
    
    If Not mCalculating Then
        EnableDisableCalculating True, False
        blnSuccess = MassCalibrationRevertToOriginal(CallerID, True, True, Me)
        EnableDisableCalculating False
        
        If blnSuccess Then
            UpdateUMCStatsAndRecomputeErrors
        End If
    End If
End Sub

Private Sub UpdateDynamicControls()
    
    ' Mass calibration tolerance type description
    If glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationTolType = gltPPM Then
        lblMassCalibrationRefinementDescription = "Note: All data is shifted by a constant ppm value, and thus a varying absolute Da value.  The ppm shift amount is determined by the location of the peak apex in a ppm-based mass-error plot."
        lblMassCalibrationRefinementUnits(0) = "ppm"
    Else
        lblMassCalibrationRefinementDescription = "Note: All data is shifted linearly by a fixed amount, determined by the location of the peak apex in a Dalton-based mass-error plot."
        lblMassCalibrationRefinementUnits(0) = "Da"
    End If
    
    lblMassCalibrationRefinementUnits(1) = lblMassCalibrationRefinementUnits(0)
    lblMassCalibrationRefinementUnits(2) = lblMassCalibrationRefinementUnits(0)

    txtMassCalibrationOverallAdjustment = GelSearchDef(CallerID).MassCalibrationInfo.OverallMassAdjustment
End Sub

Private Sub UpdatePlot()
    Dim strPlotTitle As String
    Dim dblRange As Double
    Dim intSeriesCount As Integer
    Dim intCurrentSeries As Integer
    
    If mUpdatingControls Then Exit Sub

On Error GoTo GraphMassErrorsErrorHandler
    
    ' The Error Plot peak stats need to be displayed first in order to update mMassCalErrorPeakCached and mMassCalErrorPeakCached
    DisplayErrorPlotPeakStats
   
    ' Olectra Chart requires that the data arrays be 1-based

    strPlotTitle = ""
    If Not GelAnalysis(CallerID) Is Nothing Then
        If GelAnalysis(CallerID).Job > 0 Then
            strPlotTitle = "Job " & Trim(GelAnalysis(CallerID).Job) & ": "
        End If
    End If
    strPlotTitle = strPlotTitle & StripFullPath(ExtractInputFilePath(CallerID))
    
    Select Case cboErrorDisplayMode.ListIndex
    Case mdmGanetError
        With mNETTolErrorPeakCached
            strPlotTitle = strPlotTitle & vbCrLf & DisplayErrorPlotRelativeRisk("Relative Risk: ", .PeakStats.TruePositiveArea, .PeakStats.FalsePositiveArea) & vbCrLf
        End With
    Case Else
        ' Includes mdmMassErrorPPM & mdmMassErrorDa
        With mMassCalErrorPeakCached
            strPlotTitle = strPlotTitle & vbCrLf & DisplayErrorPlotRelativeRisk("Relative Risk: ", .PeakStats.TruePositiveArea, .PeakStats.FalsePositiveArea) & vbCrLf
        End With
    End Select
            
    
    With ctlPlotter
        ' Delay updating the chart
        .EnableDisableDelayUpdating True
        
        .SetLabelGraphTitle strPlotTitle
        
        .SetChartType oc2dTypePlot, 1
        .SetCurrentGroup 1

        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowSmoothedData Then
            intSeriesCount = 2
        Else
            intSeriesCount = 1
        End If
        
        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowPeakEdges Then
            intSeriesCount = intSeriesCount + 1
        End If
        
        .SetSeriesCount intSeriesCount
    End With
    
    If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowSmoothedData Then
        UpdatePlotAddSeries 1, vbBlue, False
        UpdatePlotAddSeries 2, vbRed, True
        intCurrentSeries = 3
    Else
        UpdatePlotAddSeries 1, vbBlue, False
        intCurrentSeries = 2
    End If
    
    If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowPeakEdges Then
        UpdatePlotAddPeakEdges intCurrentSeries, RGB(128, 0, 0)
    End If
    
    ctlPlotter.SetCurrentSeries 1
        
    With ctlPlotter
        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.AutoScaleXAxis Then
            .AutoScaleXNow
        Else
            Select Case cboErrorDisplayMode.ListIndex
            Case mdmMassErrorPPM
                dblRange = glbPreferencesExpanded.ErrorPlottingOptions.MassRangePPM
            Case mdmMassErrorDa
                dblRange = PPMToMass(CDbl(glbPreferencesExpanded.ErrorPlottingOptions.MassRangePPM), 2000)
            Case mdmGanetError
                dblRange = glbPreferencesExpanded.ErrorPlottingOptions.GANETRange
            Case Else
                Debug.Assert False
                dblRange = 50
            End Select

            .SetXRange -dblRange, dblRange
        End If

        ' Set the Tick Spacing the default
        .SetXAxisTickSpacing 1, True

        .SetXAxisAnnotationMethod oc2dAnnotateValues
        .SetXAxisAnnotationPlacement oc2dAnnotateAuto
        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.CenterYAxis Then
            .SetYAxisOriginVsXAxis 0
        Else
            .SetYAxisOriginVsXAxis .GetXAxisRangeMin()
        End If
        
        .SetYAxisAnnotationMethod oc2dAnnotateValues
        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.CenterYAxis Then
            .SetYAxisAnnotationPlacement oc2dAnnotateMinimum
        Else
            .SetYAxisAnnotationPlacement oc2dAnnotateAuto
        End If

        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowGridLines Then
            .SetYAxisGridlines oc2dLineDotted
        Else
            .SetYAxisGridlines oc2dLineNone
        End If

        ' Restore the chart to update
        .EnableDisableDelayUpdating False
    End With

    Exit Sub
    
GraphMassErrorsErrorHandler:
    Debug.Print "Error in UpdatePlot: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmTICAndBPIPlots->UpdatePlot"
    
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error while populating graph: " & vbCrLf & Err.Description, vbInformation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub UpdatePlotAddSeries(intSeriesNumber As Integer, lngSeriesColor As Long, blnUseSmoothedData As Boolean)

    Dim strXAxisTitle As String

    With ctlPlotter
        .SetCurrentSeries intSeriesNumber
        
        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowPointSymbols And Not blnUseSmoothedData Then
            .SetStyleDataSymbol lngSeriesColor, oc2dShapeCross, glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.PointSizePixels
        Else
            .SetStyleDataSymbol lngSeriesColor, oc2dShapeNone, 5
        End If

        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.DrawLinesBetweenPoints Or blnUseSmoothedData Then
            .SetStyleDataLine lngSeriesColor, oc2dLineSolid, glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.LineWidthPixels
        Else
            .SetStyleDataLine lngSeriesColor, oc2dLineNone, 1
        End If

        .SetStyleDataFill lngSeriesColor, oc2dFillSolid
        
        Select Case cboErrorDisplayMode.ListIndex
        Case mdmMassErrorPPM
            ' The following will not be true if MassBinSizePPM contains more than 1 significant figure;
            '  in that case, it will have been rounded to 1 sig fig
            ' Debug.Assert glbPreferencesExpanded.ErrorPlottingOptions.MassBinSizePPM = mMassPPMErrors.BinSize
            UpdatePlotAddData intSeriesNumber, blnUseSmoothedData, mMassPPMErrors
            strXAxisTitle = "Mass Error (ppm)"
        Case mdmMassErrorDa
            UpdatePlotAddData intSeriesNumber, blnUseSmoothedData, mMassDaErrors
            strXAxisTitle = "Mass Error (Da)"
        Case mdmGanetError
            Debug.Assert glbPreferencesExpanded.ErrorPlottingOptions.GANETBinSize = mGanetErrors.BinSize
            UpdatePlotAddData intSeriesNumber, blnUseSmoothedData, mGanetErrors
            strXAxisTitle = "NET Error"
        Case Else
            ' This shouldn't happen
            Debug.Assert False
        End Select
    
        .SetLabelXAxis strXAxisTitle
        
        If glbPreferencesExpanded.RefineMSDataOptions.UseUMCClassStats Then
            .SetLabelYAxis "Count (UMCs)"
        Else
            .SetLabelYAxis "Count (Individual Peaks)"
        End If
    
    End With
    
End Sub

Private Sub UpdatePlotAddData(ByVal intSeriesNumber As Integer, ByVal blnUseSmoothedData As Boolean, ByRef udtBinnedData As udtBinnedDataType)
    
    Dim lngIndex As Long
    Dim dblXData() As Double    ' 1-based array
    Dim dblYData() As Double    ' 1-based array
    
    Dim lngMaxIndex As Long
    
    With udtBinnedData
        ctlPlotter.SetCurrentGroup 1
        
        ctlPlotter.SetSeriesDataPointCount intSeriesNumber, .BinnedCount + 1
        
        If .BinnedCount > 0 Then
    
            ReDim dblXData(1 To .BinnedCount + 1)
            ReDim dblYData(1 To .BinnedCount + 1)
    
            If blnUseSmoothedData Then
                lngMaxIndex = .BinnedCount
                If UBound(.SmoothedBins) < .BinnedCount Then
                    lngMaxIndex = .BinnedCount
                End If
                
                For lngIndex = 0 To .BinnedCount
                    dblXData(lngIndex + 1) = .StartBin + lngIndex * .BinSize
                    
                    ' Note: The smoothed data can sometimes be a negative number; we'll clip the data at 0
                    '       to avoid plotting irregularities
                    If .SmoothedBins(lngIndex) > 0 Then
                        dblYData(lngIndex + 1) = .SmoothedBins(lngIndex)
                    End If
                Next lngIndex
                
            Else
                For lngIndex = 0 To .BinnedCount
                    dblXData(lngIndex + 1) = .StartBin + lngIndex * .BinSize
                    dblYData(lngIndex + 1) = .Binned(lngIndex)
                Next lngIndex
            End If
            
            ctlPlotter.SetDataX intSeriesNumber, dblXData()
            ctlPlotter.SetDataY intSeriesNumber, dblYData()
        
        End If
        
    End With
    
End Sub

Private Sub UpdatePlotAddPeakEdges(ByVal intSeriesNumber As Integer, lngSeriesColor As Long)

On Error GoTo UpdatePlotAddPeakEdgesErrorHandler
    
    Select Case cboErrorDisplayMode.ListIndex
    Case mdmMassErrorPPM
        UpdatePlotAddPeakEdgeData intSeriesNumber, lngSeriesColor, mMassPPMErrors, mMassCalErrorPeakCached
    Case mdmMassErrorDa
        UpdatePlotAddPeakEdgeData intSeriesNumber, lngSeriesColor, mMassDaErrors, mMassCalErrorPeakCached
    Case mdmGanetError
        UpdatePlotAddPeakEdgeData intSeriesNumber, lngSeriesColor, mGanetErrors, mNETTolErrorPeakCached
    Case Else
        ' This shouldn't happen
        Debug.Assert False
        Exit Sub
    End Select

    Exit Sub
    
UpdatePlotAddPeakEdgesErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error while adding peak edges to the graph: " & vbCrLf & Err.Description, vbInformation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub UpdatePlotAddPeakEdgeData(ByVal intSeriesNumber As Integer, lngSeriesColor As Long, ByRef udtBinnedData As udtBinnedDataType, ByRef udtErrorPeak As udtErrorPlottingPeakCacheType)
  
    Const DATA_POINT_COUNT As Integer = 2
    Dim dblXData() As Double    ' 1-based array
    Dim dblYData() As Double    ' 1-based array
  
    With ctlPlotter
        .SetCurrentGroup 1
        
        .SetCurrentSeries intSeriesNumber
        
        .SetStyleDataSymbol lngSeriesColor, oc2dShapeNone, 5
        .SetStyleDataLine lngSeriesColor, oc2dLineSolid, glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.LineWidthPixels
        .SetStyleDataFill lngSeriesColor, oc2dFillSolid
        
        .SetSeriesDataPointCount intSeriesNumber, DATA_POINT_COUNT
    End With
    
    ReDim dblXData(1 To DATA_POINT_COUNT)
    ReDim dblYData(1 To DATA_POINT_COUNT)
    
    With udtBinnedData
        If .BinnedCount > 0 Then
            ' Add a line to show the peak boundaries
            dblXData(1) = .StartBin + udtErrorPeak.PeakStats.IndexBaseLeft * .BinSize
            dblXData(2) = .StartBin + udtErrorPeak.PeakStats.IndexBaseRight * .BinSize
            
            If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowSmoothedData Then
                dblYData(1) = .SmoothedBins(udtErrorPeak.PeakStats.IndexBaseLeft)
                dblYData(2) = .SmoothedBins(udtErrorPeak.PeakStats.IndexBaseRight)
                If dblYData(1) <= 0 Then dblYData(1) = .Binned(udtErrorPeak.PeakStats.IndexBaseLeft)
                If dblYData(2) <= 0 Then dblYData(2) = .Binned(udtErrorPeak.PeakStats.IndexBaseRight)
            Else
                dblYData(1) = .Binned(udtErrorPeak.PeakStats.IndexBaseLeft)
                dblYData(2) = .Binned(udtErrorPeak.PeakStats.IndexBaseRight)
            End If
            
            If dblYData(1) < 0 Then dblYData(1) = 0
            If dblYData(2) < 0 Then dblYData(2) = 0
                    
            ctlPlotter.SetDataX intSeriesNumber, dblXData()
            ctlPlotter.SetDataY intSeriesNumber, dblYData()
        End If
    End With
        
End Sub

Public Sub UpdateUMCStatsAndRecomputeErrors()
    ComputeCurrentUMCStats
    ComputeErrors True
    UpdateDynamicControls
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
    lblStatus.Caption = Msg
    DoEvents
End Sub

Private Sub cboErrorDisplayMode_Click()
    UpdatePlot
    DisplayCurrentUMCStats
    mnuEditCopyVisibleBinnedDifferences.Caption = "&Copy Binned Differences; " & cboErrorDisplayMode.Text
End Sub

Private Sub cboMassToleranceRefinementMethod_Click()
    glbPreferencesExpanded.RefineMSDataOptions.MassToleranceRefinementMethod = cboMassToleranceRefinementMethod.ListIndex
End Sub

Private Sub chkAutoScaleXRange_Click()
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.AutoScaleXAxis = cChkBox(chkAutoScaleXRange)
    UpdatePlot
End Sub

Private Sub chkCenterYAxis_Click()
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.CenterYAxis = cChkBox(chkCenterYAxis)
    UpdatePlot
End Sub

Private Sub chkDrawLinesBetweenPoints_Click()
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.DrawLinesBetweenPoints = cChkBox(chkDrawLinesBetweenPoints)
    UpdatePlot
End Sub

Private Sub chkIncludeInternalStandards_Click()
    glbPreferencesExpanded.RefineMSDataOptions.IncludeInternalStdMatches = cChkBox(chkIncludeInternalStandards)
    ComputeErrors False
End Sub

Private Sub chkShowGridlines_Click()
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowGridLines = cChkBox(chkShowGridlines)
    UpdatePlot
End Sub

Private Sub chkShowPeakEdges_Click()
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowPeakEdges = cChkBox(chkShowPeakEdges)
    UpdatePlot
End Sub

Private Sub chkShowPointSymbols_Click()
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowPointSymbols = cChkBox(chkShowPointSymbols)
    UpdatePlot
End Sub

Private Sub chkShowSmoothedData_Click()
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowSmoothedData = cChkBox(chkShowSmoothedData)
    UpdatePlot
End Sub

Private Sub chkShowToleranceRefinementControls_Click()
    ShowHideToleranceRefinementControls cChkBox(chkShowToleranceRefinementControls)
End Sub

Private Sub chkUseUMCClassStats_Click()
    glbPreferencesExpanded.RefineMSDataOptions.UseUMCClassStats = cChkBox(chkUseUMCClassStats)
    ComputeErrors False
End Sub

Private Sub cmdAbortProcessing_Click()
    mAbortProcessing = True
End Sub

Private Sub cmdMassCalibrationManual_Click()
    If Not mCalculating Then ManualRefineMassCalibration False
End Sub

Private Sub cmdMassCalibrationRefinementStart_Click()
    If Not mCalculating Then RefineMassCalibrationStart
End Sub

Private Sub cmdMassCalibrationRevert_Click()
    StartMassCalibrationRevert
End Sub

Private Sub cmdMassToleranceRefinementStart_Click()
    If Not mCalculating Then RefineDBSearchMassToleranceStart
End Sub

Private Sub cmdNETToleranceRefinementStart_Click()
    If Not mCalculating Then RefineDBSearchNETToleranceStart
End Sub

Private Sub cmdRecomputeHistograms_Click()
    ComputeErrors False
End Sub

Private Sub Form_Activate()
    InitializeForm
End Sub

Private Sub Form_Load()
    Dim dblBlankDataX(1 To 1) As Double
    Dim dblBlankDataY(1 To 1) As Double
    
    Me.ScaleMode = vbTwips
    
    Me.width = 12870
    Me.Height = 10000

    
    mAMTIndicesInitialized = False
    mFormInitialized = False
    
    With ctlPlotter
        .EnableDisableDelayUpdating True
        .SetCurrentGroup 2
        .SetSeriesCount 0

        .SetCurrentGroup 1
        
        If glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.ShowSmoothedData Then
            .SetSeriesCount 2
        Else
            .SetSeriesCount 1
        End If

        .SetSeriesDataPointCount 1, 1
        .SetDataX 1, dblBlankDataX()
        .SetDataY 1, dblBlankDataY()

        .EnableDisableDelayUpdating False
    End With
    
    tbsRefinement.Tab = 0
    PositionControls
    
    ComputeCurrentUMCStats
    DisplayCurrentDBSearchTolerances
    
    ShowHideOptions True
    
    PopulateComboBoxes
    
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub mnuCopyChart_Click(Index As Integer)
    Select Case Index
    Case ccmWMF
        ctlPlotter.CopyToClipboard oc2dFormatMetafile
    Case ccmEMF
        ctlPlotter.CopyToClipboard oc2dFormatEnhMetafile
    Case Else
        ' Includes ccmBMP
        ctlPlotter.CopyToClipboard oc2dFormatBitmap
    End Select
End Sub

Private Sub mnuCopyErrors_Click()
    ExportErrorsToClipboardOrFile
End Sub

Private Sub mnuCopyErrorsBinned_Click()
    ExportErrorsBinnedToClipboardOrFile "", True, False
End Sub

Private Sub mnuEditCopyVisibleBinnedDifferences_Click()
    ExportErrorsBinnedToClipboardOrFile "", True, True
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSaveBinnedDataToTextFile_Click()
    Dim strFilePath As String
    
    strFilePath = SelectFile(Me.hwnd, "Enter filename", "", True, "MassErrorsBinned.txt", "Text Files (*.txt)|*.txt|All Files (*.*)|*.*")

    If Len(strFilePath) > 0 Then
        ExportErrorsBinnedToClipboardOrFile strFilePath, True, False
    End If
End Sub

Private Sub mnuSaveChartPicture_Click(Index As Integer)
    If Index = pftPictureFileTypeConstants.pftJPG Then
        SaveChartPictureToFile False
    Else
        ' Inclues pftPictureFileTypeConstants.pftPNG
        SaveChartPictureToFile True
    End If
End Sub

Private Sub mnuSaveDataToTextFile_Click()
    Dim strFilePath As String
    
    strFilePath = SelectFile(Me.hwnd, "Enter filename", "", True, "MassErrors.txt", "Text Files (*.txt)|*.txt|All Files (*.*)|*.*")

    If Len(strFilePath) > 0 Then
        ExportErrorsToClipboardOrFile strFilePath
    End If
End Sub

Private Sub mnuSetToleranceRefinementOptionsToDefault_Click()
    SetToleranceRefinementOptionsToDefault
End Sub

Private Sub mnuView3DErrorDistributions_Click()
    ShowOrCompute3DErrorDistributions False
End Sub

Private Sub mnuView3DErrorDistributionsInverted_Click()
    ShowOrCompute3DErrorDistributions True
End Sub

Private Sub mnuViewOptions_Click()
    ShowHideOptions
End Sub

Private Sub optRefineMassCalibrationMassType_Click(Index As Integer)
    Dim blnShowMessage As Boolean
    Dim strMessage As String
    
    If GelSearchDef(CallerID).MassCalibrationInfo.AdjustmentHistoryCount = 0 Then
        If Index = 0 Then
            glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationTolType = gltPPM
        Else
            glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationTolType = gltABS
        End If
    Else
        ' Do not allow user to set units if at least one adjustment has been applied
        If GelSearchDef(CallerID).MassCalibrationInfo.MassUnits = gltPPM Then
            If optRefineMassCalibrationMassType(0).Value <> True Then
                optRefineMassCalibrationMassType(0).Value = True
                blnShowMessage = True
            End If
        Else
            If optRefineMassCalibrationMassType(1).Value <> True Then
                optRefineMassCalibrationMassType(1).Value = True
                blnShowMessage = True
            End If
        End If
        
        If blnShowMessage Then
            strMessage = "Previous calibration adjustments were performed using " & GetSearchToleranceUnitText(CInt(GelSearchDef(CallerID).MassCalibrationInfo.MassUnits)) & " units.  You are not allowed to perform additional adjustments using differing units.  You must use Revert to Original masses before adjusting using a different unit base."
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled And blnShowMessage Then
                MsgBox strMessage, vbInformation + vbOKOnly, "Incompatible Units"
            Else
                AddToAnalysisHistory CallerID, strMessage
            End If
        End If
    End If
    
    UpdateDynamicControls
End Sub

Private Sub txtButterworthFrequency_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtButterworthFrequency_Validate (False)
        ComputeErrors False
    End If
End Sub

Private Sub txtButterworthFrequency_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtButterworthFrequency, KeyAscii, True, True
End Sub

Private Sub txtButterworthFrequency_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtButterworthFrequency, 0.01, 100, 0.2
    With glbPreferencesExpanded.ErrorPlottingOptions
        If .ButterWorthFrequency <> CSngSafe(txtButterworthFrequency) Then
            .ButterWorthFrequency = CSngSafe(txtButterworthFrequency)
        End If
    End With
End Sub

Private Sub txtGANETBinSize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtGANETBinSize_Validate (False)
        ComputeErrors False
    End If
End Sub
Private Sub txtGANETBinSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGANETBinSize, KeyAscii, True, True
End Sub

Private Sub txtGANETBinSize_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtGANETBinSize, 0.00001, 5, DEFAULT_GANET_BIN_SIZE
    With glbPreferencesExpanded.ErrorPlottingOptions
        If .GANETBinSize <> CSngSafe(txtGANETBinSize) Then
            .GANETBinSize = CSngSafe(txtGANETBinSize)
        End If
    End With
End Sub

Private Sub txtganetRange_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGANETRange, KeyAscii, True, True
End Sub

Private Sub txtganetRange_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtGANETRange, 0.01, 5, 0.3
    glbPreferencesExpanded.ErrorPlottingOptions.GANETRange = CSngSafe(txtGANETRange)
    UpdatePlot
End Sub

Private Sub txtGraphLineWidth_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphLineWidth, KeyAscii, True, False
End Sub

Private Sub txtGraphLineWidth_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtGraphLineWidth, 1, 20, 3
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.LineWidthPixels = CLngSafe(txtGraphLineWidth)
    UpdatePlot
End Sub

Private Sub txtGraphPointSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphPointSize, KeyAscii, True, False
End Sub

Private Sub txtGraphPointSize_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtGraphPointSize, 1, 20, 2
    glbPreferencesExpanded.ErrorPlottingOptions.Graph2DOptions.PointSizePixels = CLngSafe(txtGraphPointSize)
    UpdatePlot
End Sub

Private Sub txtMassBinSizePPM_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        txtMassBinSizePPM_Validate (False)
        ComputeErrors False
    End If
End Sub

Private Sub txtMassBinSizePPM_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMassBinSizePPM, KeyAscii, True, True
End Sub

Private Sub txtMassBinSizePPM_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtMassBinSizePPM, 0.01, 10000, DEFAULT_MASS_BIN_SIZE_PPM
    With glbPreferencesExpanded.ErrorPlottingOptions
        If .MassBinSizePPM <> CSngSafe(txtMassBinSizePPM) Then
            .MassBinSizePPM = CSngSafe(txtMassBinSizePPM)
        End If
    End With
End Sub

Private Sub txtMassRangePPM_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMassRangePPM, KeyAscii, True, False
End Sub

Private Sub txtMassRangePPM_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtMassRangePPM, 1, 10000, 100
    glbPreferencesExpanded.ErrorPlottingOptions.MassRangePPM = CSngSafe(txtMassRangePPM)
    UpdatePlot
End Sub

Private Sub txtMaximumAbundance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtMaximumAbundance_Validate (False)
        ComputeErrors False
    End If
End Sub

Private Sub txtMaximumAbundance_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtMaximumAbundance, 0, 1E+300, 0
    
    With glbPreferencesExpanded.RefineMSDataOptions
        If .MaximumAbundance <> CDblSafe(txtMaximumAbundance) Then
            .MaximumAbundance = CDblSafe(txtMaximumAbundance)
        End If
    End With
End Sub

Private Sub txtMinimumSLiC_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        txtMinimumSLiC_Validate (False)
        ComputeErrors False
    End If
End Sub

Private Sub txtMinimumSLiC_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtMinimumSLiC, 0, 1, 0
    
    With glbPreferencesExpanded.RefineMSDataOptions
        If .MinimumSLiC <> CSngSafe(txtMinimumSLiC) Then
            .MinimumSLiC = CSngSafe(txtMinimumSLiC)
        End If
    End With
End Sub

Private Sub txtRefineDBSearchMassToleranceAdjustmentMultiplier_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchMassToleranceAdjustmentMultiplier, 0.0001, 10000, 1
    glbPreferencesExpanded.RefineMSDataOptions.MassToleranceAdjustmentMultiplier = CDblSafe(txtRefineDBSearchMassToleranceAdjustmentMultiplier)
End Sub

Private Sub txtRefineDBSearchMassToleranceMaximum_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchMassToleranceMaximum, 0, 1E+300, 10
    If Not ValidateDualTextBoxes(txtRefineDBSearchMassToleranceMinimum, txtRefineDBSearchMassToleranceMaximum, False, 0, 1E+300, 0) Then
        txtRefineDBSearchMassToleranceMinimum.SetFocus
    End If
    glbPreferencesExpanded.RefineMSDataOptions.MassToleranceMaximum = CDblSafe(txtRefineDBSearchMassToleranceMaximum)
End Sub

Private Sub txtRefineDBSearchMassToleranceMinimum_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchMassToleranceMinimum, 0, 1E+300, 1
    If Not ValidateDualTextBoxes(txtRefineDBSearchMassToleranceMinimum, txtRefineDBSearchMassToleranceMaximum, True, 0, 1E+300, 0) Then
        txtRefineDBSearchMassToleranceMaximum.SetFocus
    End If
    glbPreferencesExpanded.RefineMSDataOptions.MassToleranceMinimum = CDblSafe(txtRefineDBSearchMassToleranceMinimum)
End Sub

Private Sub txtRefineDBSearchNETToleranceAdjustmentMultiplier_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchNETToleranceAdjustmentMultiplier, 0.0001, 10000, 1
    glbPreferencesExpanded.RefineMSDataOptions.NETToleranceAdjustmentMultiplier = CDblSafe(txtRefineDBSearchNETToleranceAdjustmentMultiplier)
End Sub

Private Sub txtRefineDBSearchNETToleranceMaximum_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchNETToleranceMaximum, 0.0001, 100, 0.2
    If Not ValidateDualTextBoxes(txtRefineDBSearchNETToleranceMinimum, txtRefineDBSearchNETToleranceMaximum, False, 0.0001, 100, 0) Then
        txtRefineDBSearchNETToleranceMinimum.SetFocus
    End If
   glbPreferencesExpanded.RefineMSDataOptions.NETToleranceMaximum = CDblSafe(txtRefineDBSearchNETToleranceMaximum)
End Sub

Private Sub txtRefineDBSearchNETToleranceMinimum_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchNETToleranceMinimum, 0.0001, 100, 0.01
    If Not ValidateDualTextBoxes(txtRefineDBSearchNETToleranceMinimum, txtRefineDBSearchNETToleranceMaximum, True, 0.0001, 100, 0) Then
        txtRefineDBSearchNETToleranceMaximum.SetFocus
    End If
    glbPreferencesExpanded.RefineMSDataOptions.NETToleranceMinimum = CDblSafe(txtRefineDBSearchNETToleranceMinimum)
End Sub

Private Sub txtRefineMassCalibrationMaximumShift_LostFocus()
    ValidateTextboxValueDbl txtRefineMassCalibrationMaximumShift, 0, 1E+300, 15
    glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationMaximumShift = CDblSafe(txtRefineMassCalibrationMaximumShift)
End Sub

Private Sub txtToleranceRefinementMinimumPeakHeight_LostFocus()
    ValidateTextboxValueLng txtToleranceRefinementMinimumPeakHeight, 0, 1000000000#, 25
    glbPreferencesExpanded.RefineMSDataOptions.MinimumPeakHeight = CLngSafe(txtToleranceRefinementMinimumPeakHeight)
End Sub

Private Sub txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks_Change()
    ' Yes, I want this code to be in a _Change procedure
    If IsNumeric(txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks) Then
        If val(txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks) >= 0 And val(txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks) <= 100000 Then
            glbPreferencesExpanded.RefineMSDataOptions.MinimumSignalToNoiseRatioForLowAbundancePeaks = CSngSafe(txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks)
            DisplayErrorPlotPeakStats
        End If
    End If
End Sub

Private Sub txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks_LostFocus()
    ValidateTextboxValueDbl txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks, 0, 100000, 2.5
    glbPreferencesExpanded.RefineMSDataOptions.MinimumSignalToNoiseRatioForLowAbundancePeaks = CSngSafe(txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks)
End Sub

Private Sub txtToleranceRefinementPercentageOfMaxForWidth_Change()
    ' Yes, I want this code to be in a _Change procedure
    If IsNumeric(txtToleranceRefinementPercentageOfMaxForWidth) Then
        If val(txtToleranceRefinementPercentageOfMaxForWidth) >= 0 And val(txtToleranceRefinementPercentageOfMaxForWidth) <= 100 Then
            glbPreferencesExpanded.RefineMSDataOptions.PercentageOfMaxForFindingWidth = CLngSafe(txtToleranceRefinementPercentageOfMaxForWidth)
            DisplayErrorPlotPeakStats
        End If
    End If
End Sub

Private Sub txtToleranceRefinementPercentageOfMaxForWidth_Lostfocus()
    ValidateTextboxValueLng txtToleranceRefinementPercentageOfMaxForWidth, 0, 100, 60
    glbPreferencesExpanded.RefineMSDataOptions.PercentageOfMaxForFindingWidth = CLngSafe(txtToleranceRefinementPercentageOfMaxForWidth)
End Sub


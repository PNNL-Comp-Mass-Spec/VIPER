VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditAnalysisSettings 
   Caption         =   "Edit Analysis Settings"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9870
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Interval        =   500
      Left            =   11400
      Top             =   8640
   End
   Begin TabDlg.SSTab tbsTabStrip 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   13785
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   4
      TabsPerRow      =   8
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "1. Load and Filter"
      TabPicture(0)   =   "frmEditAnalysisSettings.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDescription(50)"
      Tab(0).Control(1)=   "fraOptionFrame(4)"
      Tab(0).Control(2)=   "fraOptionFrame(5)"
      Tab(0).Control(3)=   "fraOptionFrame(6)"
      Tab(0).Control(4)=   "fraOptionFrame(7)"
      Tab(0).Control(5)=   "txtPEKFileExtensionPreferenceOrder"
      Tab(0).Control(6)=   "fraOptionFrame(9)"
      Tab(0).Control(7)=   "fraOptionFrame(8)"
      Tab(0).Control(8)=   "fraOptionFrame(22)"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "2. LC-MS Features"
      TabPicture(1)   =   "frmEditAnalysisSettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkSkipFindUMCs"
      Tab(1).Control(1)=   "cboChargeStateAbuType"
      Tab(1).Control(2)=   "chkUseMostAbuChargeStateStatsForClassStats"
      Tab(1).Control(3)=   "tbsUMCRefinementOptions"
      Tab(1).Control(4)=   "fraUMCIonNetOptions"
      Tab(1).Control(5)=   "chkInterpolateMissingIons"
      Tab(1).Control(6)=   "txtInterpolateMaxGapSize"
      Tab(1).Control(7)=   "cmbUMCAbu"
      Tab(1).Control(8)=   "cmbUMCMW"
      Tab(1).Control(9)=   "fraUMCSearch200x"
      Tab(1).Control(10)=   "cboUMCSearchMode"
      Tab(1).Control(11)=   "lblDescription(61)"
      Tab(1).Control(12)=   "lblDescription(58)"
      Tab(1).Control(13)=   "lblDescription(64)"
      Tab(1).Control(14)=   "lblDescription(63)"
      Tab(1).Control(15)=   "lblDescription(115)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "3. MT Tags"
      TabPicture(2)   =   "frmEditAnalysisSettings.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOptionFrame(19)"
      Tab(2).Control(1)=   "fraOptionFrame(44)"
      Tab(2).Control(2)=   "cmdSelectOtherDB"
      Tab(2).Control(3)=   "fraSelectingMassTags"
      Tab(2).Control(4)=   "fraOptionFrame(16)"
      Tab(2).Control(5)=   "lblDescription(121)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "4. Pairs"
      TabPicture(3)   =   "frmEditAnalysisSettings.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraOptionFrame(56)"
      Tab(3).Control(1)=   "fraOptionFrame(53)"
      Tab(3).Control(2)=   "fraOptionFrame(52)"
      Tab(3).Control(3)=   "fraOptionFrame(48)"
      Tab(3).Control(4)=   "fraOptionFrame(41)"
      Tab(3).Control(5)=   "fraOptionFrame(40)"
      Tab(3).Control(6)=   "fraOptionFrame(38)"
      Tab(3).Control(7)=   "fraOptionFrame(37)"
      Tab(3).Control(8)=   "cboPairsIdentificationMode"
      Tab(3).Control(9)=   "lblDescription(16)"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "5. NET Adjustment"
      TabPicture(4)   =   "frmEditAnalysisSettings.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "lblDescription(156)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "chkSkipGANETSlopeAndInterceptComputation"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "tbsNETOptions"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "fraOptionFrame(18)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "chkRobustNETEnabled"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cboRobustNETAdjustmentMode"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "pctOptionsInIniFile(3)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "pctOptionsInIniFile(4)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "6. Refinement"
      TabPicture(5)   =   "frmEditAnalysisSettings.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblToleranceRefinementExplanation"
      Tab(5).Control(1)=   "lblMassCalibrationRefinementDescription"
      Tab(5).Control(2)=   "fraOptionFrame(1)"
      Tab(5).Control(3)=   "fraOptionFrame(0)"
      Tab(5).Control(4)=   "fraOptionFrame(3)"
      Tab(5).Control(5)=   "fraOptionFrame(2)"
      Tab(5).Control(6)=   "fraToleranceRefinementPeakCriteria"
      Tab(5).Control(7)=   "chkRefineDBSearchMassTolerance"
      Tab(5).Control(8)=   "chkRefineDBSearchNETTolerance"
      Tab(5).Control(9)=   "cboToleranceRefinementMethod"
      Tab(5).Control(10)=   "fraExpecationMaximizationOptions"
      Tab(5).Control(11)=   "chkRefineDBSearchTolUseMinMaxIfOutOfRange"
      Tab(5).ControlCount=   12
      TabCaption(6)   =   "7. DB Search"
      TabPicture(6)   =   "frmEditAnalysisSettings.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblToleranceRefinementWarning"
      Tab(6).Control(1)=   "lblDescription(70)"
      Tab(6).Control(2)=   "fraOptionFrame(30)"
      Tab(6).Control(3)=   "fraOptionFrame(26)"
      Tab(6).Control(4)=   "cmdAddRemove(2)"
      Tab(6).Control(5)=   "fraOptionFrame(29)"
      Tab(6).Control(6)=   "chkDoNotSaveOrExport"
      Tab(6).Control(7)=   "lstDBSearchModes"
      Tab(6).Control(8)=   "cboDBSearchMode"
      Tab(6).Control(9)=   "fraOptionFrame(45)"
      Tab(6).Control(10)=   "fraOptionFrame(31)"
      Tab(6).ControlCount=   11
      TabCaption(7)   =   "8. Saving/Plotting"
      TabPicture(7)   =   "frmEditAnalysisSettings.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblSavingAndExportingStatus"
      Tab(7).Control(1)=   "lblDescription(136)"
      Tab(7).Control(2)=   "fraOptionFrame(32)"
      Tab(7).Control(3)=   "fraOptionFrame(33)"
      Tab(7).Control(4)=   "pctOptionsInIniFile(0)"
      Tab(7).Control(5)=   "pctOptionsInIniFile(1)"
      Tab(7).ControlCount=   6
      Begin VB.PictureBox pctOptionsInIniFile 
         Height          =   2800
         Index           =   4
         Left            =   4150
         Picture         =   "frmEditAnalysisSettings.frx":00E0
         ScaleHeight     =   2745
         ScaleWidth      =   6885
         TabIndex        =   543
         Top             =   4680
         Width           =   6940
      End
      Begin VB.PictureBox pctOptionsInIniFile 
         Height          =   5055
         Index           =   3
         Left            =   120
         Picture         =   "frmEditAnalysisSettings.frx":15F52
         ScaleHeight     =   4995
         ScaleWidth      =   3855
         TabIndex        =   542
         Top             =   2640
         Width           =   3915
      End
      Begin VB.CheckBox chkRefineDBSearchTolUseMinMaxIfOutOfRange 
         Caption         =   "Use min or max tol if out of range"
         Height          =   375
         Left            =   -65760
         TabIndex        =   540
         Top             =   3720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Frame fraExpecationMaximizationOptions 
         Caption         =   "Expectation Maximization Options"
         Height          =   2055
         Left            =   -69360
         TabIndex        =   392
         Top             =   4100
         Width           =   3495
         Begin VB.TextBox txtEMRefineMassErrorPeakToleranceEstimatePPM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   394
            Text            =   "6"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtEMRefineNETErrorPeakToleranceEstimate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   397
            Text            =   "0.01"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtEMRefinePercentOfDataToExclude 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   399
            Text            =   "10"
            Top             =   960
            Width           =   495
         End
         Begin VB.CheckBox chkEMRefineMassTolForceUseAllDataPointErrors 
            Caption         =   "Use single data point errors for Mass"
            Height          =   255
            Left            =   120
            TabIndex        =   401
            Top             =   1360
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox chkEMRefineNETTolForceUseAllDataPointErrors 
            Caption         =   "Use single data point errors for NET"
            Height          =   255
            Left            =   120
            TabIndex        =   402
            Top             =   1680
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.Label lblDescription 
            Caption         =   "ppm"
            Height          =   255
            Index           =   166
            Left            =   2880
            TabIndex        =   395
            Top             =   270
            Width           =   540
         End
         Begin VB.Label lblDescription 
            Caption         =   "Mass error tol. estimate"
            Height          =   255
            Index           =   165
            Left            =   120
            TabIndex        =   393
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label lblDescription 
            Caption         =   "NETerror tol. estimate"
            Height          =   255
            Index           =   164
            Left            =   120
            TabIndex        =   396
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label lblDescription 
            Caption         =   "Data to exclude from extremes"
            Height          =   255
            Index           =   163
            Left            =   120
            TabIndex        =   398
            Top             =   990
            Width           =   2175
         End
         Begin VB.Label lblDescription 
            Caption         =   "%"
            Height          =   255
            Index           =   167
            Left            =   3000
            TabIndex        =   400
            Top             =   990
            Width           =   195
         End
      End
      Begin VB.ComboBox cboToleranceRefinementMethod 
         Height          =   315
         Left            =   -69360
         Style           =   2  'Dropdown List
         TabIndex        =   391
         Top             =   3720
         Width           =   3495
      End
      Begin VB.CheckBox chkRefineDBSearchNETTolerance 
         Caption         =   "Refine DB NET search tol. using NET error plot"
         Height          =   495
         Left            =   -66480
         TabIndex        =   381
         Top             =   1800
         Width           =   2400
      End
      Begin VB.CheckBox chkRefineDBSearchMassTolerance 
         Caption         =   "Refine DB mass search tol. using mass error plot"
         Height          =   495
         Left            =   -69240
         TabIndex        =   371
         Top             =   1800
         Width           =   2400
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Data Count Filter"
         Height          =   1065
         Index           =   22
         Left            =   -69960
         TabIndex        =   56
         Top             =   4080
         Width           =   5895
         Begin VB.CheckBox chkMaximumDataCountEnabled 
            Caption         =   "Maximum data count filter enabled (requires pre-scan of data file)"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   270
            Width           =   5415
         End
         Begin VB.TextBox txtMaximumDataCountToLoad 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   59
            ToolTipText     =   "Higher abundance data is favored when determine the data to load"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum data count to load"
            Height          =   255
            Index           =   137
            Left            =   120
            TabIndex        =   58
            Top             =   645
            Width           =   2055
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Legacy MT DB Path"
         Height          =   1215
         Index           =   19
         Left            =   -74880
         TabIndex        =   226
         Top             =   6240
         Width           =   10215
         Begin VB.CheckBox chkUseLegacyDBForMTs 
            Caption         =   "Use Legacy DB for MT tags"
            Height          =   255
            Left            =   120
            TabIndex        =   229
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtLegacyAMTDatabasePath 
            Height          =   315
            Left            =   120
            TabIndex        =   227
            Top             =   285
            Width           =   8895
         End
         Begin VB.CommandButton cmdBrowseAMT 
            Caption         =   "B&rowse"
            Height          =   375
            Left            =   9240
            TabIndex        =   228
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.ComboBox cboRobustNETAdjustmentMode 
         Height          =   315
         ItemData        =   "frmEditAnalysisSettings.frx":2B1B8
         Left            =   6600
         List            =   "frmEditAnalysisSettings.frx":2B1BA
         Style           =   2  'Dropdown List
         TabIndex        =   329
         Top             =   4200
         Width           =   4365
      End
      Begin VB.PictureBox pctOptionsInIniFile 
         Height          =   2145
         Index           =   1
         Left            =   -68400
         Picture         =   "frmEditAnalysisSettings.frx":2B1BC
         ScaleHeight     =   2085
         ScaleWidth      =   2865
         TabIndex        =   528
         Top             =   2880
         Width           =   2925
      End
      Begin VB.CheckBox chkSkipFindUMCs 
         Caption         =   "Skip Finding LC-MS Features (only appropriate if auto-processing .Gel files)"
         Height          =   420
         Left            =   -74760
         TabIndex        =   68
         Top             =   420
         Width           =   4455
      End
      Begin VB.CheckBox chkRobustNETEnabled 
         Caption         =   "Robust NET Enabled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   4200
         TabIndex        =   327
         Top             =   4230
         Width           =   2175
      End
      Begin VB.Frame fraToleranceRefinementPeakCriteria 
         Caption         =   "Criteria To Use Peak For Refinement"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   352
         Top             =   4320
         Width           =   3015
         Begin VB.TextBox txtToleranceRefinementPercentageOfMaxForWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   357
            Text            =   "60"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtToleranceRefinementMinimumPeakHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   354
            Text            =   "25"
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   359
            Text            =   "2.5"
            Top             =   920
            Width           =   615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Pct of Max for Finding Width"
            Height          =   255
            Index           =   147
            Left            =   120
            TabIndex        =   356
            Top             =   640
            Width           =   2055
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum Height"
            Height          =   255
            Index           =   146
            Left            =   120
            TabIndex        =   353
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label lblDescription 
            Caption         =   "counts/bin"
            Height          =   255
            Index           =   8
            Left            =   2040
            TabIndex        =   355
            Top             =   330
            Width           =   840
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum S/N for Low Abu"
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   358
            Top             =   940
            Width           =   2055
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "ER Calculation Mode"
         Height          =   615
         Index           =   56
         Left            =   -71040
         TabIndex        =   522
         Top             =   720
         Width           =   5055
         Begin VB.OptionButton optERCalc 
            Caption         =   "Ratio (L/H)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   525
            ToolTipText     =   "Ratio (Abundance of Light member/ Abundance of Heavy Member)"
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optERCalc 
            Caption         =   "Log. Ratio"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   524
            ToolTipText     =   "Logarithmic Ratio (Abundance of Light member/ Abundance of Heavy Member)"
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optERCalc 
            Caption         =   "Symmetric Ratio"
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   523
            ToolTipText     =   "Zero Symetric Ratio (Abundance of Light member/ Abundance of Heavy Member)"
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.PictureBox pctOptionsInIniFile 
         Height          =   2220
         Index           =   0
         Left            =   -74880
         Picture         =   "frmEditAnalysisSettings.frx":2BAEE
         ScaleHeight     =   2160
         ScaleWidth      =   9345
         TabIndex        =   521
         Top             =   5450
         Width           =   9400
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Pair member removal after database search"
         Height          =   1215
         Index           =   53
         Left            =   -64800
         TabIndex        =   287
         Top             =   6120
         Visible         =   0   'False
         Width           =   5295
         Begin VB.OptionButton optRemovePairMemberHitsRemoveHeavy 
            Caption         =   "Remove heavy member hits"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   291
            Top             =   920
            Width           =   2295
         End
         Begin VB.OptionButton optRemovePairMemberHitsRemoveHeavy 
            Caption         =   "Remove light member hits"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   290
            Top             =   680
            Width           =   2295
         End
         Begin VB.CheckBox chkRemovePairMemberHitsAfterDBSearch 
            Caption         =   "Enable removal"
            Height          =   300
            Left            =   240
            TabIndex        =   289
            Top             =   720
            Width           =   1560
         End
         Begin VB.Label lblDescription 
            Caption         =   "This option will find pairs, then search the database, then remove the light or heavy member of the pairs that match MT tags."
            Height          =   420
            Index           =   113
            Left            =   240
            TabIndex        =   288
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Pair Search and ER Calculation Options"
         Height          =   2415
         Index           =   52
         Left            =   -69360
         TabIndex        =   273
         Top             =   3600
         Width           =   5295
         Begin VB.CheckBox chkPairEROptions 
            Caption         =   "Enable I-Report ER computation"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   279
            Top             =   1320
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.CheckBox chkPairEROptions 
            Caption         =   "Remove outlier ER values using Grubb's test (95% conf.)"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   280
            Top             =   1635
            Value           =   1  'Checked
            Width           =   4695
         End
         Begin VB.CheckBox chkPairEROptions 
            Caption         =   "Repeatedly remove outliers"
            Height          =   300
            Index           =   6
            Left            =   360
            TabIndex        =   281
            Top             =   1875
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.TextBox txtRemoveOutlierERsMinimumDataPointCount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   283
            Text            =   "3"
            Top             =   1965
            Width           =   615
         End
         Begin VB.ComboBox cboAverageERsWeightingMode 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   277
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkPairEROptions 
            Caption         =   "Use identical charge states for expression ratio"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   275
            Top             =   495
            Value           =   1  'Checked
            Width           =   4095
         End
         Begin VB.CheckBox chkPairEROptions 
            Caption         =   "Require matching charge states for pair"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   274
            Top             =   240
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox chkPairEROptions 
            Caption         =   "Compute ER Scan by Scan"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   278
            Top             =   1020
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkPairEROptions 
            Caption         =   "Average ER's for all charge states"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   276
            Top             =   765
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum final data point count"
            Height          =   405
            Index           =   140
            Left            =   2880
            TabIndex        =   282
            Top             =   1875
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboChargeStateAbuType 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CheckBox chkUseMostAbuChargeStateStatsForClassStats 
         Caption         =   "Use most abundant charge state group stats for class stats"
         Height          =   520
         Left            =   -71760
         TabIndex        =   80
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Text file saving options"
         Height          =   855
         Index           =   48
         Left            =   -69360
         TabIndex        =   284
         Top             =   6480
         Width           =   3375
         Begin VB.CheckBox chkPairsSaveStatisticsTextFile 
            Caption         =   "Save pairs statistics (binned)  to text file"
            Height          =   300
            Left            =   120
            TabIndex        =   286
            Top             =   480
            Value           =   1  'Checked
            Width           =   3195
         End
         Begin VB.CheckBox chkPairsSaveTextFile 
            Caption         =   "Save pairs to text file"
            Height          =   300
            Left            =   120
            TabIndex        =   285
            Top             =   240
            Value           =   1  'Checked
            Width           =   3000
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Text Export Options"
         Height          =   1995
         Index           =   31
         Left            =   -74760
         TabIndex        =   456
         Top             =   5520
         Width           =   3255
         Begin VB.CheckBox chkSaveUMCStatisticsToTextFile 
            Caption         =   "Save all LC-MS Feature to text file (not just those with DB matches)"
            Height          =   375
            Left            =   240
            TabIndex        =   459
            ToolTipText     =   "Saves statistics, including scan, mass, NET, and member information, for all UMC's"
            Top             =   1010
            Width           =   2900
         End
         Begin VB.TextBox txtOutputFileSeparationCharacter 
            Height          =   285
            Left            =   2280
            TabIndex        =   461
            Text            =   "<TAB>"
            ToolTipText     =   "Type a single separation character, like a comma or a semicolon, or enter the text <TAB> to indicate a tab."
            Top             =   1545
            Width           =   615
         End
         Begin VB.CheckBox chkDBSearchWriteIDResultsByIonAfterAutoSearches 
            Caption         =   "Write search results by ion to text file after auto search completion"
            Height          =   495
            Left            =   240
            TabIndex        =   457
            Top             =   200
            Width           =   2805
         End
         Begin VB.CheckBox chkDBSearchIncludeORFNameInTextFileOutput 
            Caption         =   "Include ORF name in text file output"
            Height          =   255
            Left            =   240
            TabIndex        =   458
            Top             =   710
            Value           =   1  'Checked
            Width           =   2900
         End
         Begin VB.Label lblDescription 
            Caption         =   "Text output file separation character"
            Height          =   405
            Index           =   68
            Left            =   240
            TabIndex        =   460
            Top             =   1485
            Width           =   1695
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Pairs Search Options"
         Height          =   2175
         Index           =   41
         Left            =   -74760
         TabIndex        =   264
         Top             =   5160
         Width           =   5295
         Begin VB.CheckBox chkPairsExcludeAmbiguousKeepMostConfident 
            Caption         =   "Ambiguous pairs exclusion keeps most confident pair"
            Height          =   255
            Left            =   360
            TabIndex        =   526
            Top             =   1130
            Value           =   1  'Checked
            Width           =   4485
         End
         Begin VB.ComboBox cboPairsUMCsToUseForNETAdjustment 
            Height          =   315
            ItemData        =   "frmEditAnalysisSettings.frx":6EE20
            Left            =   120
            List            =   "frmEditAnalysisSettings.frx":6EE22
            Style           =   2  'Dropdown List
            TabIndex        =   272
            Top             =   1680
            Width           =   4935
         End
         Begin VB.CheckBox chkPairsExcludeOutOfERRange 
            Caption         =   "Exclude pairs outside of the ER inclusion range"
            Height          =   255
            Left            =   120
            TabIndex        =   269
            Top             =   620
            Width           =   4455
         End
         Begin VB.CheckBox chkPairsExcludeAmbiguous 
            Caption         =   "Exclude ambiguous pairs after database search (prior to exporting)"
            Height          =   255
            Left            =   120
            TabIndex        =   270
            Top             =   875
            Width           =   5055
         End
         Begin VB.TextBox txtPairsERMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2280
            TabIndex        =   266
            Text            =   "-5"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtPairsERMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3720
            TabIndex        =   268
            Text            =   "5"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblDescription 
            Caption         =   "LC-MS Feature to use for NET adjustment"
            Height          =   255
            Index           =   102
            Left            =   120
            TabIndex        =   271
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lblDescription 
            Caption         =   "to"
            Height          =   255
            Index           =   99
            Left            =   3360
            TabIndex        =   267
            ToolTipText     =   "285"
            Top             =   280
            Width           =   255
         End
         Begin VB.Label lblDescription 
            Caption         =   "ER Inclusion Range:"
            Height          =   255
            Index           =   98
            Left            =   240
            TabIndex        =   265
            Top             =   280
            Width           =   1935
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Tolerances"
         Height          =   1250
         Index           =   40
         Left            =   -74760
         TabIndex        =   256
         Top             =   3760
         Width           =   5295
         Begin VB.OptionButton optPairTolType 
            Caption         =   "&Dalton"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   545
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optPairTolType 
            Caption         =   "&ppm"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   544
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtPairsScanTolApex 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4080
            TabIndex        =   263
            Text            =   "5"
            Top             =   900
            Width           =   495
         End
         Begin VB.TextBox txtPairsScanTolEdge 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4080
            TabIndex        =   261
            Text            =   "5"
            Top             =   575
            Width           =   495
         End
         Begin VB.CheckBox chkPairsRequireOverlapAtApex 
            Caption         =   "Require pair-classes &overlap at feature apexes"
            Height          =   255
            Left            =   240
            TabIndex        =   262
            ToolTipText     =   "If checked pair classes have to show at least once in the same scan"
            Top             =   900
            Width           =   3600
         End
         Begin VB.TextBox txtPairsDeltaTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   258
            Text            =   "0.02"
            Top             =   180
            Width           =   855
         End
         Begin VB.CheckBox chkPairsRequireOverlapAtEdge 
            Caption         =   "Require pair-classes &overlap at feature edges"
            Height          =   255
            Left            =   240
            TabIndex        =   260
            ToolTipText     =   "If checked pair classes have to show at least once in the same scan"
            Top             =   600
            Value           =   1  'Checked
            Width           =   3600
         End
         Begin VB.Label lblDescription 
            Caption         =   "Scan Tolerance:"
            Height          =   255
            Index           =   97
            Left            =   3840
            TabIndex        =   259
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label lblDescription 
            Caption         =   "Pair Tolerance:"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   257
            Top             =   215
            Width           =   1335
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Isotopic Labeling Values (e.g. ICAT)"
         Height          =   2055
         Index           =   38
         Left            =   -71040
         TabIndex        =   245
         Top             =   1440
         Width           =   5055
         Begin VB.TextBox txtPairsMaxLblDiff 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            TabIndex        =   255
            Text            =   "1"
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtPairsHeavyLabelDelta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   249
            Text            =   "8.05"
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtPairsMinMaxLbl 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   251
            Text            =   "1"
            Top             =   660
            Width           =   855
         End
         Begin VB.TextBox txtPairsMinMaxLbl 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   253
            Text            =   "5"
            Top             =   1020
            Width           =   855
         End
         Begin VB.TextBox txtPairsLabel 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   247
            Text            =   "442.2249697"
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max. difference between number of labels in light/heavy:"
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   24
            Left            =   120
            TabIndex        =   254
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label lblDescription 
            Caption         =   "Heavy/Light Delta:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   21
            Left            =   2520
            TabIndex        =   248
            Top             =   340
            Width           =   1455
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Labels:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   250
            Top             =   700
            Width           =   900
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max Labels:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   252
            Top             =   1060
            Width           =   900
         End
         Begin VB.Label lblDescription 
            Caption         =   "Label (Light)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   246
            Top             =   340
            Width           =   975
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Isotopic Ratio Delta Values"
         Height          =   2295
         Index           =   37
         Left            =   -74760
         TabIndex        =   232
         Top             =   1440
         Width           =   3495
         Begin VB.TextBox txtPairsMinMaxDelta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1920
            TabIndex        =   244
            Text            =   "1"
            Top             =   1920
            Width           =   615
         End
         Begin VB.CommandButton cmdSetPairsToC13 
            Caption         =   "C12/C13"
            Height          =   300
            Left            =   1200
            TabIndex        =   236
            Top             =   645
            Width           =   975
         End
         Begin VB.TextBox txtPairsDelta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   234
            Text            =   "0.9970356"
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txtPairsMinMaxDelta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2760
            TabIndex        =   242
            Text            =   "100"
            Top             =   1620
            Width           =   615
         End
         Begin VB.TextBox txtPairsMinMaxDelta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   240
            Text            =   "1"
            Top             =   1620
            Width           =   615
         End
         Begin VB.CommandButton cmdSetPairsToN15 
            Caption         =   "N14/N15"
            Height          =   300
            Left            =   120
            TabIndex        =   235
            Top             =   645
            Width           =   975
         End
         Begin VB.CommandButton cmdSetPairsToO18 
            Caption         =   "O16/O18"
            Height          =   300
            Left            =   2280
            TabIndex        =   237
            Top             =   645
            Width           =   975
         End
         Begin VB.CheckBox chkPairsAutoMinMaxDelta 
            Caption         =   "C&alculate N14/N15 Min/Max Deltas from class molecular mass"
            Height          =   420
            Left            =   240
            TabIndex        =   238
            Top             =   1155
            Width           =   2775
         End
         Begin VB.Label lblDescription 
            Caption         =   "Delta count step size:"
            Height          =   255
            Index           =   141
            Left            =   120
            TabIndex        =   243
            Top             =   1935
            Width           =   1860
         End
         Begin VB.Label lblDescription 
            Caption         =   "Delta:"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   233
            Top             =   340
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max Deltas:"
            Height          =   255
            Index           =   19
            Left            =   1880
            TabIndex        =   241
            Top             =   1640
            Width           =   900
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Deltas:"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   239
            Top             =   1640
            Width           =   900
         End
      End
      Begin VB.ComboBox cboPairsIdentificationMode 
         Height          =   315
         ItemData        =   "frmEditAnalysisSettings.frx":6EE24
         Left            =   -74760
         List            =   "frmEditAnalysisSettings.frx":6EE26
         Style           =   2  'Dropdown List
         TabIndex        =   231
         Top             =   960
         Width           =   3495
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Search (Identification) Tolerances"
         Height          =   1720
         Index           =   18
         Left            =   240
         TabIndex        =   293
         Top             =   840
         Width           =   3735
         Begin VB.Frame fraOptionFrame 
            Caption         =   "MW Tolerance"
            Height          =   975
            Index           =   36
            Left            =   240
            TabIndex        =   294
            Top             =   260
            Width           =   1935
            Begin VB.OptionButton optNETAdjTolType 
               Caption         =   "&Dalton"
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   298
               Top             =   600
               Width           =   855
            End
            Begin VB.OptionButton optNETAdjTolType 
               Caption         =   "&ppm"
               Height          =   255
               Index           =   0
               Left            =   1020
               TabIndex        =   297
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.TextBox txtNETAdjMWTol 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   160
               TabIndex        =   296
               Text            =   "10"
               Top             =   525
               Width           =   735
            End
            Begin VB.Label lblDescription 
               Caption         =   "Tolerance"
               Height          =   255
               Index           =   51
               Left            =   165
               TabIndex        =   295
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox txtNetAdjInitialNETTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   301
            Text            =   "0.2"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblDescription 
            Caption         =   "Note: The defined Class Mass is used for NET adjustment"
            Height          =   810
            Index           =   7
            Left            =   2265
            TabIndex        =   299
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label lblDescription 
            Caption         =   "Initial NET tolerance"
            Height          =   255
            Index           =   52
            Left            =   240
            TabIndex        =   300
            Top             =   1350
            Width           =   1695
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Mass Tolerance Refinement"
         Height          =   1350
         Index           =   2
         Left            =   -69360
         TabIndex        =   372
         Top             =   2280
         Width           =   2655
         Begin VB.TextBox txtRefineDBSearchMassToleranceMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   374
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtRefineDBSearchMassToleranceMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   377
            Text            =   "0"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtRefineDBSearchMassToleranceAdjustmentMultiplier 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   380
            Text            =   "1"
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum Tol."
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   373
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label lblDescription 
            Caption         =   "ppm"
            Height          =   255
            Index           =   86
            Left            =   1920
            TabIndex        =   375
            Top             =   270
            Width           =   540
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum Tol."
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   376
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label lblDescription 
            Caption         =   "ppm"
            Height          =   255
            Index           =   87
            Left            =   1920
            TabIndex        =   378
            Top             =   630
            Width           =   540
         End
         Begin VB.Label lblDescription 
            Caption         =   "Adjustment multiplier"
            Height          =   255
            Index           =   94
            Left            =   120
            TabIndex        =   379
            Top             =   990
            Width           =   1605
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "NET Tolerance Refinement"
         Height          =   1350
         Index           =   3
         Left            =   -66600
         TabIndex        =   382
         Top             =   2280
         Width           =   2655
         Begin VB.TextBox txtRefineDBSearchNETToleranceMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   387
            Text            =   "0"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtRefineDBSearchNETToleranceMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   384
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtRefineDBSearchNETToleranceAdjustmentMultiplier 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   390
            Text            =   "1"
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Adjustment multiplier"
            Height          =   255
            Index           =   93
            Left            =   120
            TabIndex        =   389
            Top             =   990
            Width           =   1605
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum Tol."
            Height          =   255
            Index           =   84
            Left            =   120
            TabIndex        =   386
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum Tol."
            Height          =   255
            Index           =   83
            Left            =   120
            TabIndex        =   383
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label lblDescription 
            Caption         =   "NET"
            Height          =   255
            Index           =   90
            Left            =   1920
            TabIndex        =   388
            Top             =   630
            Width           =   540
         End
         Begin VB.Label lblDescription 
            Caption         =   "NET"
            Height          =   255
            Index           =   89
            Left            =   1920
            TabIndex        =   385
            Top             =   270
            Width           =   540
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Criteria To Use Peak For Refinement"
         Height          =   2280
         Index           =   0
         Left            =   -74880
         TabIndex        =   331
         Top             =   1920
         Width           =   5415
         Begin VB.TextBox txtToleranceRefinementFilter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   347
            Text            =   "0"
            Top             =   1880
            Width           =   615
         End
         Begin VB.ComboBox cboSearchRegionShape 
            Height          =   315
            Index           =   0
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   339
            Top             =   640
            Width           =   2775
         End
         Begin VB.TextBox txtToleranceRefinementFilter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   351
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtToleranceRefinementFilter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   349
            Text            =   "0"
            Top             =   1635
            Width           =   615
         End
         Begin VB.CheckBox chkUseUMCClassStats 
            Caption         =   "Use LC-MS Feature class stats"
            Height          =   255
            Left            =   2520
            TabIndex        =   341
            Top             =   1275
            Width           =   2745
         End
         Begin VB.TextBox txtToleranceRefinementFilter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   345
            Text            =   "0"
            Top             =   1580
            Width           =   615
         End
         Begin VB.CheckBox chkRefineDBSearchIncludeInternalStds 
            Caption         =   "Include Internal Standard matches"
            Height          =   375
            Left            =   2520
            TabIndex        =   340
            Top             =   960
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.TextBox txtToleranceRefinementFilter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   343
            Text            =   "0"
            Top             =   1280
            Width           =   615
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Molecular Mass Tolerance"
            Height          =   975
            Index           =   39
            Left            =   120
            TabIndex        =   332
            Top             =   240
            Width           =   2175
            Begin VB.OptionButton optAutoToleranceRefinementDBSearchTolType 
               Caption         =   "&Dalton"
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   336
               Top             =   600
               Width           =   855
            End
            Begin VB.OptionButton optAutoToleranceRefinementDBSearchTolType 
               Caption         =   "&ppm"
               Height          =   255
               Index           =   0
               Left            =   1020
               TabIndex        =   335
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.TextBox txtAutoToleranceRefinementDBSearchMWTol 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   160
               TabIndex        =   334
               Text            =   "25"
               ToolTipText     =   "Database search tolerance"
               Top             =   525
               Width           =   735
            End
            Begin VB.Label lblDescription 
               Caption         =   "Tolerance"
               Height          =   255
               Index           =   91
               Left            =   165
               TabIndex        =   333
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox txtAutoToleranceRefinementDBSearchNETTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3720
            TabIndex        =   338
            Text            =   "0.1"
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Pep Prophet"
            Height          =   255
            Index           =   157
            Left            =   120
            TabIndex        =   346
            Top             =   1900
            Width           =   1545
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Abundance"
            Height          =   225
            Index           =   14
            Left            =   2520
            TabIndex        =   350
            Top             =   1940
            Width           =   1575
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum SLiC"
            Height          =   225
            Index           =   11
            Left            =   2520
            TabIndex        =   348
            Top             =   1665
            Width           =   1695
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Discrim. Score"
            Height          =   255
            Index           =   144
            Left            =   120
            TabIndex        =   344
            Top             =   1600
            Width           =   1545
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum XCorr"
            Height          =   255
            Index           =   135
            Left            =   120
            TabIndex        =   342
            Top             =   1300
            Width           =   1185
         End
         Begin VB.Label lblDescription 
            Caption         =   "NET tolerance"
            Height          =   255
            Index           =   92
            Left            =   2520
            TabIndex        =   337
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Mass Calibration Refinement"
         Height          =   1695
         Index           =   1
         Left            =   -74880
         TabIndex        =   360
         Top             =   5760
         Width           =   5415
         Begin VB.TextBox txtRefineMassCalibrationOverridePPM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            TabIndex        =   369
            Text            =   "0"
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox chkRefineMassCalibration 
            Caption         =   "Re-calibrate (shift mass) using mass error plot"
            Height          =   375
            Left            =   240
            TabIndex        =   361
            Top             =   240
            Width           =   3765
         End
         Begin VB.TextBox txtRefineMassCalibrationMaximumShift 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   366
            Text            =   "0"
            Top             =   1200
            Width           =   615
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Tolerance Type"
            Height          =   855
            Index           =   25
            Left            =   240
            TabIndex        =   362
            Top             =   720
            Width           =   1455
            Begin VB.OptionButton optRefineMassCalibrationMassType 
               Caption         =   "Dalton"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   364
               Top             =   520
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optRefineMassCalibrationMassType 
               Caption         =   "ppm"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   363
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Label lblDescription 
            Caption         =   "ppm"
            Height          =   255
            Index           =   139
            Left            =   4320
            TabIndex        =   370
            Top             =   1230
            Width           =   600
         End
         Begin VB.Label lblDescription 
            Caption         =   "Override Shift Amount"
            Height          =   255
            Index           =   138
            Left            =   3600
            TabIndex        =   368
            Top             =   960
            Width           =   1700
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum shift"
            Height          =   255
            Index           =   77
            Left            =   2040
            TabIndex        =   365
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblMassCalibrationRefinementUnits 
            Caption         =   "ppm"
            Height          =   255
            Left            =   2760
            TabIndex        =   367
            Top             =   1230
            Width           =   600
         End
      End
      Begin VB.Frame fraOptionFrame 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   45
         Left            =   -70440
         TabIndex        =   419
         Top             =   840
         Width           =   375
         Begin VB.CommandButton cmdAddRemove 
            Caption         =   " >"
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   421
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmdAddRemove 
            Caption         =   "<"
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   420
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.ComboBox cboDBSearchMode 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   405
         Top             =   1260
         Width           =   4095
      End
      Begin VB.ListBox lstDBSearchModes 
         Height          =   840
         Left            =   -69960
         TabIndex        =   422
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox chkDoNotSaveOrExport 
         Caption         =   "Disable all saving and exporting"
         Height          =   255
         Left            =   -74760
         TabIndex        =   403
         Top             =   600
         Width           =   3015
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Options for selected Database Search Mode"
         Height          =   3735
         Index           =   29
         Left            =   -70680
         TabIndex        =   424
         Top             =   2160
         Width           =   6735
         Begin VB.TextBox txtDBSearchMinimumPeptideProphetProbability 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            TabIndex        =   534
            Text            =   "0"
            Top             =   2400
            Width           =   615
         End
         Begin VB.TextBox txtDBSearchMinimumHighDiscriminantScore 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            TabIndex        =   447
            Text            =   "0"
            Top             =   2100
            Width           =   615
         End
         Begin VB.TextBox txtDBSearchMinimumHighNormalizedScore 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            TabIndex        =   446
            Text            =   "0"
            Top             =   1800
            Width           =   615
         End
         Begin VB.ComboBox cboInternalStdSearchMode 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   445
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Modifications"
            Height          =   1530
            Index           =   42
            Left            =   120
            TabIndex        =   425
            Top             =   240
            Width           =   6375
            Begin VB.CheckBox chkMTAdditionalMass 
               Caption         =   "PEO"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   426
               Top             =   360
               Width           =   975
            End
            Begin VB.CheckBox chkMTAdditionalMass 
               Caption         =   "ICAT d0"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   427
               Top             =   720
               Width           =   975
            End
            Begin VB.CheckBox chkMTAdditionalMass 
               Caption         =   "ICAT d8"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   428
               Top             =   1080
               Width           =   975
            End
            Begin VB.CheckBox chkMTAdditionalMass 
               Caption         =   "Alkylation"
               Height          =   255
               Index           =   3
               Left            =   1320
               TabIndex        =   429
               ToolTipText     =   "Check to add the alkylation mass correction below to all MT tag masses (added to each cys residue)"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtAlkylationMWCorrection 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1440
               TabIndex        =   431
               Text            =   "57.0215"
               Top             =   960
               Width           =   855
            End
            Begin VB.Frame fraOptionFrame 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   1095
               Index           =   49
               Left            =   4160
               TabIndex        =   436
               Top             =   360
               Width           =   1095
               Begin VB.OptionButton optDBSearchModType 
                  Caption         =   "Decoy"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   547
                  Top             =   720
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton optDBSearchModType 
                  Caption         =   "Fixed"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   438
                  Top             =   240
                  Width           =   750
               End
               Begin VB.OptionButton optDBSearchModType 
                  Caption         =   "Dynamic"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   439
                  Top             =   500
                  Width           =   975
               End
               Begin VB.Label lblDescription 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mod Type:"
                  Height          =   255
                  Index           =   100
                  Left            =   120
                  TabIndex        =   437
                  Top             =   0
                  Width           =   900
               End
            End
            Begin VB.Frame fraOptionFrame 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   855
               Index           =   47
               Left            =   5400
               TabIndex        =   440
               Top             =   360
               Width           =   800
               Begin VB.OptionButton optDBSearchNType 
                  Caption         =   "N15"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   443
                  Top             =   525
                  Width           =   700
               End
               Begin VB.OptionButton optDBSearchNType 
                  Caption         =   "N14"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   442
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
                  TabIndex        =   441
                  Top             =   0
                  Width           =   600
               End
            End
            Begin VB.TextBox txtResidueToModifyMass 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   435
               Text            =   "0"
               Top             =   1080
               Width           =   855
            End
            Begin VB.ComboBox cboResidueToModify 
               Height          =   315
               Left            =   2760
               Style           =   2  'Dropdown List
               TabIndex        =   433
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Alkylation mass:"
               Height          =   255
               Index           =   108
               Left            =   1320
               TabIndex        =   430
               Top             =   720
               Width           =   1215
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   1200
               X2              =   1200
               Y1              =   240
               Y2              =   1440
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   4080
               X2              =   4080
               Y1              =   240
               Y2              =   1440
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   5280
               X2              =   5280
               Y1              =   240
               Y2              =   1440
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   2520
               X2              =   2520
               Y1              =   240
               Y2              =   1440
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Mass (Da):"
               Height          =   255
               Index           =   107
               Left            =   2640
               TabIndex        =   434
               Top             =   840
               Width           =   975
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Residue to modify:"
               Height          =   255
               Index           =   101
               Left            =   2640
               TabIndex        =   432
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdBrowseForFolder 
            Caption         =   "Browse"
            Height          =   255
            Left            =   5700
            TabIndex        =   455
            Top             =   3360
            Width           =   855
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "MT tags"
            Height          =   825
            Index           =   46
            Left            =   5040
            TabIndex        =   448
            Top             =   1800
            Visible         =   0   'False
            Width           =   1500
            Begin VB.OptionButton optPairsDBSearchLabelAssumption 
               Caption         =   "Not Labeled"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   449
               ToolTipText     =   "MT tags are not labeled (mass in DB is the mass of the unmodified sequence)"
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optPairsDBSearchLabelAssumption 
               Caption         =   "Labeled"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   450
               ToolTipText     =   "MT tags are labeled (mass in DB is already modified)"
               Top             =   500
               Width           =   1215
            End
         End
         Begin VB.CheckBox chkDBSearchExportToDB 
            Caption         =   "Export search results to database"
            Height          =   255
            Left            =   180
            TabIndex        =   451
            Top             =   2820
            Width           =   3015
         End
         Begin VB.CheckBox chkDBSearchWriteResultsToTextFile 
            Caption         =   "Write search results to text file"
            Height          =   255
            Left            =   3240
            TabIndex        =   452
            Top             =   2820
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.TextBox txtDBSearchAlternateOutputFolderPath 
            Height          =   285
            Left            =   180
            TabIndex        =   454
            Top             =   3360
            Width           =   5295
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum XCorr"
            Height          =   255
            Index           =   161
            Left            =   2760
            TabIndex        =   533
            Top             =   1820
            Width           =   1185
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Discrim. Score"
            Height          =   255
            Index           =   143
            Left            =   2760
            TabIndex        =   532
            Top             =   2120
            Width           =   1545
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Pep Prophet"
            Height          =   255
            Index           =   134
            Left            =   2760
            TabIndex        =   531
            Top             =   2420
            Width           =   1545
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Internal Std Search Mode:"
            Height          =   255
            Index           =   112
            Left            =   180
            TabIndex        =   444
            Top             =   1880
            Width           =   2415
         End
         Begin VB.Label lblDescription 
            Caption         =   "Alternate output folder path (ignored during PRISM-initiated analysis)"
            Height          =   255
            Index           =   40
            Left            =   180
            TabIndex        =   453
            Top             =   3105
            Width           =   4905
         End
      End
      Begin VB.CommandButton cmdAddRemove 
         Caption         =   "Remove All"
         Height          =   255
         Index           =   2
         Left            =   -69960
         TabIndex        =   423
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Search Options"
         Height          =   2715
         Index           =   26
         Left            =   -74760
         TabIndex        =   406
         Top             =   1740
         Width           =   3855
         Begin VB.ComboBox cboSearchRegionShape 
            Height          =   315
            Index           =   1
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   530
            Top             =   2320
            Width           =   3000
         End
         Begin VB.CheckBox chkUseUMCConglomerateNET 
            Caption         =   "Use Class NET for LC-MS Features"
            Height          =   615
            Left            =   2280
            TabIndex        =   411
            ToolTipText     =   $"frmEditAnalysisSettings.frx":6EE28
            Top             =   720
            Width           =   1485
         End
         Begin VB.TextBox txtDBSearchNETTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   418
            Text            =   "0.1"
            Top             =   1830
            Width           =   735
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Molecular Mass Tolerance"
            Height          =   925
            Index           =   28
            Left            =   240
            TabIndex        =   412
            Top             =   1340
            Width           =   2175
            Begin VB.TextBox txtDBSearchMWTol 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   160
               TabIndex        =   414
               Text            =   "10"
               Top             =   525
               Width           =   735
            End
            Begin VB.OptionButton optDBSearchTolType 
               Caption         =   "&ppm"
               Height          =   255
               Index           =   0
               Left            =   1020
               TabIndex        =   415
               Top             =   300
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optDBSearchTolType 
               Caption         =   "&Dalton"
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   416
               Top             =   560
               Width           =   855
            End
            Begin VB.Label lblDescription 
               Caption         =   "Tolerance"
               Height          =   255
               Index           =   57
               Left            =   165
               TabIndex        =   413
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Molecular Mass Field"
            Height          =   1095
            Index           =   27
            Left            =   240
            TabIndex        =   407
            Top             =   200
            Width           =   1935
            Begin VB.OptionButton optDBSearchMWField 
               Caption         =   "A&verage"
               Height          =   255
               Index           =   0
               Left            =   80
               TabIndex        =   408
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton optDBSearchMWField 
               Caption         =   "&Monoisotopic"
               Height          =   255
               Index           =   1
               Left            =   80
               TabIndex        =   409
               Top             =   500
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton optDBSearchMWField 
               Caption         =   "&The Most Abundant"
               Height          =   255
               Index           =   2
               Left            =   80
               TabIndex        =   410
               Top             =   760
               Width           =   1815
            End
         End
         Begin VB.Label lblDescription 
            Caption         =   "NET tolerance"
            Height          =   255
            Index           =   69
            Left            =   2520
            TabIndex        =   417
            Top             =   1580
            Width           =   1095
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Search Result Export Options"
         Height          =   1485
         Index           =   30
         Left            =   -70680
         TabIndex        =   462
         Top             =   6000
         Width           =   5295
         Begin VB.CheckBox chkExportResultsFileUsesJobNumber 
            Caption         =   "Export results file uses job number instead of dataset name"
            Height          =   255
            Left            =   360
            TabIndex        =   466
            Top             =   1080
            Value           =   1  'Checked
            Width           =   4515
         End
         Begin VB.CheckBox chkSetIsConfirmedForDBSearchMatches 
            Caption         =   "Set Is_Confirmed = 1 for database search matches"
            Height          =   255
            Left            =   360
            TabIndex        =   463
            Top             =   300
            Value           =   1  'Checked
            Width           =   4000
         End
         Begin VB.CheckBox chkAddQuantitationDescriptionEntry 
            Caption         =   "Add Quantitation Description Entry"
            Height          =   255
            Left            =   360
            TabIndex        =   464
            Top             =   560
            Value           =   1  'Checked
            Width           =   4000
         End
         Begin VB.CheckBox chkExportUMCsWithNoMatches 
            Caption         =   "Export LC-MS Features with no database matches"
            Height          =   255
            Left            =   360
            TabIndex        =   465
            Top             =   820
            Width           =   4000
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Error Plot Save Options"
         Height          =   1935
         Index           =   33
         Left            =   -74760
         TabIndex        =   479
         Top             =   3120
         Width           =   6255
         Begin VB.CheckBox chkSaveErrorGraphicMass 
            Caption         =   "Save mass error plot"
            Height          =   255
            Left            =   240
            TabIndex        =   480
            Top             =   300
            Width           =   2400
         End
         Begin VB.CheckBox chkSaveErrorGraphicGANET 
            Caption         =   "Save NET error plot"
            Height          =   255
            Left            =   3360
            TabIndex        =   481
            Top             =   240
            Width           =   2400
         End
         Begin VB.CheckBox chkSaveErrorGraphic3D 
            Caption         =   "Save 3D error plot (Mass vs. NET)"
            Height          =   255
            Left            =   240
            TabIndex        =   482
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtErrorGraphicMassRangePPM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   484
            Text            =   "100"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtErrorGraphicGANETRange 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   490
            Text            =   "0.3"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtErrorGraphicGANETBinSize 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4920
            TabIndex        =   492
            Text            =   "0.005"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtErrorGraphicMassBinSizePPM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4920
            TabIndex        =   487
            Text            =   "1"
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Range (± 0)"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   483
            Top             =   1110
            Width           =   1575
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "ppm"
            Height          =   255
            Index           =   10
            Left            =   2880
            TabIndex        =   485
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "NET Range (± 0)"
            Height          =   255
            Index           =   73
            Left            =   240
            TabIndex        =   489
            Top             =   1470
            Width           =   1575
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "NET Bin Size"
            Height          =   255
            Index           =   74
            Left            =   3480
            TabIndex        =   491
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "ppm"
            Height          =   255
            Index           =   82
            Left            =   5760
            TabIndex        =   488
            Top             =   1110
            Width           =   405
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Bin Size"
            Height          =   255
            Index           =   71
            Left            =   3480
            TabIndex        =   486
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Gel Data"
         Height          =   2055
         Index           =   32
         Left            =   -74760
         TabIndex        =   468
         Top             =   960
         Width           =   5535
         Begin VB.CheckBox chkSaveGelFile 
            Caption         =   "Save Gel File (raw data and DB Search Results)"
            Height          =   255
            Left            =   240
            TabIndex        =   469
            Top             =   240
            Width           =   3855
         End
         Begin VB.CheckBox chkSavePictureGraphic 
            Caption         =   "Save graphic of 2D display"
            Height          =   375
            Left            =   240
            TabIndex        =   472
            Top             =   1200
            Width           =   3015
         End
         Begin VB.ComboBox cboSavePictureGraphicFileType 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   473
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox txtPictureGraphicSizeWidthPixels 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4680
            TabIndex        =   476
            Text            =   "1024"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtPictureGraphicSizeHeightPixels 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4680
            TabIndex        =   478
            Text            =   "768"
            Top             =   1680
            Width           =   615
         End
         Begin VB.CheckBox chkExtendedFileSaveModePreferred 
            Caption         =   "Include LC-MS Features with the raw data when saving"
            Height          =   255
            Left            =   600
            TabIndex        =   470
            Top             =   510
            Width           =   4815
         End
         Begin VB.CheckBox chkSaveGelFileOnError 
            Caption         =   "Save gel file if an error occurs"
            Height          =   255
            Left            =   240
            TabIndex        =   471
            Top             =   800
            Width           =   3855
         End
         Begin VB.Label lblDescription 
            Caption         =   "File type"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   504
            Top             =   1590
            Width           =   975
         End
         Begin VB.Label lblDescription 
            Caption         =   "Width"
            Height          =   255
            Index           =   5
            Left            =   3960
            TabIndex        =   475
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Height"
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   477
            Top             =   1710
            Width           =   615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Pixels"
            Height          =   255
            Index           =   3
            Left            =   4680
            TabIndex        =   474
            Top             =   1080
            Width           =   615
         End
      End
      Begin TabDlg.SSTab tbsUMCRefinementOptions 
         Height          =   3120
         Left            =   -69360
         TabIndex        =   155
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5503
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Auto-Refine Options"
         TabPicture(0)   =   "frmEditAnalysisSettings.frx":6EEBF
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraOptionFrame(10)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Split Features Options"
         TabPicture(1)   =   "frmEditAnalysisSettings.frx":6EEDB
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraOptionFrame(15)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Adv Class Stats"
         TabPicture(2)   =   "frmEditAnalysisSettings.frx":6EEF7
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraOptionFrame(55)"
         Tab(2).ControlCount=   1
         Begin VB.Frame fraOptionFrame 
            Height          =   2700
            Index           =   10
            Left            =   120
            TabIndex        =   539
            Top             =   320
            Width           =   4545
            Begin VB.TextBox txtAutoRefineTextbox 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   3000
               TabIndex        =   188
               Text            =   "15"
               Top             =   1520
               Width           =   495
            End
            Begin VB.CheckBox chkAutoRefineCheckbox 
               Caption         =   "Remove cls. with length over"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   187
               Top             =   1520
               Width           =   2535
            End
            Begin VB.TextBox txtAutoRefineTextbox 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   3000
               TabIndex        =   191
               Text            =   "33"
               Top             =   1840
               Width           =   495
            End
            Begin VB.TextBox txtAutoRefineTextbox 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   3480
               TabIndex        =   195
               Text            =   "3"
               Top             =   2300
               Width           =   495
            End
            Begin VB.CheckBox chkAutoRefineCheckbox 
               Caption         =   "Test feature length using scan range"
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   193
               ToolTipText     =   "If True, then considers scan range for the length tests; otherwise, considers member count"
               Top             =   2200
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox chkAutoRefineCheckbox 
               Caption         =   "Remove low intensity classes"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   175
               Top             =   240
               Width           =   2550
            End
            Begin VB.TextBox txtAutoRefineTextbox 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   3000
               TabIndex        =   176
               Text            =   "30"
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox chkAutoRefineCheckbox 
               Caption         =   "Remove high intensity classes"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   178
               Top             =   560
               Width           =   2550
            End
            Begin VB.TextBox txtAutoRefineTextbox 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   3000
               TabIndex        =   179
               Text            =   "30"
               Top             =   560
               Width           =   495
            End
            Begin VB.CheckBox chkAutoRefineCheckbox 
               Caption         =   "Remove cls. with less than"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   181
               Top             =   880
               Width           =   2295
            End
            Begin VB.TextBox txtAutoRefineTextbox 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   3000
               TabIndex        =   182
               Text            =   "3"
               Top             =   880
               Width           =   495
            End
            Begin VB.CheckBox chkAutoRefineCheckbox 
               Caption         =   "Remove cls. with length over"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   184
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox txtAutoRefineTextbox 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   3000
               TabIndex        =   185
               Text            =   "500"
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label lblAutoRefineLengthLabel 
               Caption         =   "%"
               Height          =   255
               Index           =   5
               Left            =   3600
               TabIndex        =   180
               Top             =   590
               Width           =   270
            End
            Begin VB.Label lblAutoRefineLengthLabel 
               Caption         =   "%"
               Height          =   255
               Index           =   4
               Left            =   3600
               TabIndex        =   177
               Top             =   270
               Width           =   270
            End
            Begin VB.Label lblAutoRefineLengthLabel 
               Caption         =   "% all scans"
               Height          =   255
               Index           =   3
               Left            =   3600
               TabIndex        =   189
               Top             =   1545
               Width           =   855
            End
            Begin VB.Label lblAutoRefineLengthLabel 
               Caption         =   "%"
               Height          =   255
               Index           =   2
               Left            =   3600
               TabIndex        =   192
               Top             =   1870
               Width           =   285
            End
            Begin VB.Label lblAutoRefineSpecialLabel 
               Caption         =   "Percent max abu for gauging width"
               Height          =   240
               Index           =   0
               Left            =   360
               TabIndex        =   190
               Top             =   1845
               Width           =   2565
            End
            Begin VB.Label lblAutoRefineSpecialLabel 
               Caption         =   "Minimum member count:"
               Height          =   375
               Index           =   1
               Left            =   2280
               TabIndex        =   194
               Top             =   2200
               Width           =   1125
            End
            Begin VB.Label lblAutoRefineLengthLabel 
               Caption         =   "members"
               Height          =   255
               Index           =   0
               Left            =   3600
               TabIndex        =   183
               Top             =   915
               Width           =   900
            End
            Begin VB.Label lblAutoRefineLengthLabel 
               Caption         =   "members"
               Height          =   255
               Index           =   1
               Left            =   3600
               TabIndex        =   186
               Top             =   1230
               Width           =   900
            End
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Class Abundance Top X"
            Height          =   1215
            Index           =   17
            Left            =   -74760
            TabIndex        =   513
            Top             =   360
            Width           =   4095
            Begin VB.TextBox txtClassAbuTopXMinMaxAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   2880
               TabIndex        =   516
               Text            =   "0"
               Top             =   240
               Width           =   900
            End
            Begin VB.TextBox txtClassAbuTopXMinMaxAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   2880
               TabIndex        =   515
               Text            =   "0"
               ToolTipText     =   "Maximum abundance to include; use 0 to indicate there infinitely large abundance"
               Top             =   540
               Width           =   900
            End
            Begin VB.TextBox txtClassAbuTopXMinMembers 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2880
               TabIndex        =   514
               Text            =   "3"
               Top             =   840
               Width           =   900
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Minimum Abundance to Include"
               Height          =   255
               Index           =   88
               Left            =   120
               TabIndex        =   519
               Top             =   270
               Width           =   2535
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Maximum Abundance to Include"
               Height          =   255
               Index           =   104
               Left            =   120
               TabIndex        =   518
               Top             =   560
               Width           =   2535
            End
            Begin VB.Label lblClassAbuTopXMinMembers 
               BackStyle       =   0  'Transparent
               Caption         =   "Minimum members to include"
               Height          =   255
               Left            =   120
               TabIndex        =   517
               Top             =   870
               Width           =   2535
            End
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Class Mass Top X"
            Height          =   1215
            Index           =   55
            Left            =   -74760
            TabIndex        =   506
            Top             =   1600
            Width           =   4095
            Begin VB.TextBox txtClassMassTopXMinMembers 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2880
               TabIndex        =   509
               Text            =   "3"
               Top             =   840
               Width           =   900
            End
            Begin VB.TextBox txtClassMassTopXMinMaxAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   2880
               TabIndex        =   508
               Text            =   "0"
               ToolTipText     =   "Maximum abundance to include; use 0 to indicate there infinitely large abundance"
               Top             =   540
               Width           =   900
            End
            Begin VB.TextBox txtClassMassTopXMinMaxAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   2880
               TabIndex        =   507
               Text            =   "0"
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lblClassMassTopXMinMembers 
               BackStyle       =   0  'Transparent
               Caption         =   "Minimum members to include"
               Height          =   255
               Left            =   120
               TabIndex        =   512
               Top             =   870
               Width           =   2535
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Maximum Abundance to Include"
               Height          =   255
               Index           =   106
               Left            =   120
               TabIndex        =   511
               Top             =   560
               Width           =   2535
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Minimum Abundance to Include"
               Height          =   255
               Index           =   105
               Left            =   120
               TabIndex        =   510
               Top             =   270
               Width           =   2535
            End
         End
         Begin VB.Frame fraOptionFrame 
            BorderStyle     =   0  'None
            Height          =   2340
            Index           =   15
            Left            =   -74760
            TabIndex        =   156
            Top             =   440
            Width           =   4260
            Begin VB.TextBox txtSplitUMCsStdDevMultiplierForSplitting 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   162
               Text            =   "1"
               Top             =   660
               Width           =   495
            End
            Begin VB.CheckBox chkSplitUMCsByExaminingAbundance 
               Caption         =   "Split LC-MS Features by examining abundance"
               Height          =   255
               Left            =   0
               TabIndex        =   157
               Top             =   0
               Width           =   4215
            End
            Begin VB.TextBox txtSplitUMCsMinimumDifferenceInAvgPpmMass 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   159
               Text            =   "4"
               Top             =   330
               Width           =   495
            End
            Begin VB.TextBox txtSplitUMCsMaximumPeakCount 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   164
               Text            =   "6"
               Top             =   990
               Width           =   495
            End
            Begin VB.TextBox txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   167
               Text            =   "15"
               Top             =   1320
               Width           =   495
            End
            Begin VB.TextBox txtSplitUMCsPeakPickingMinimumWidth 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               TabIndex        =   170
               Text            =   "4"
               Top             =   1650
               Width           =   495
            End
            Begin VB.TextBox txtHoleSize 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   2760
               TabIndex        =   173
               Text            =   "3"
               Top             =   1980
               Width           =   495
            End
            Begin VB.ComboBox cboSplitUMCsScanGapBehavior 
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   197
               Top             =   2340
               Width           =   2295
            End
            Begin VB.Label lblSplitUMCsStdDevMultiplierForSplitting 
               Caption         =   "Mass Std Dev threshold multiplier"
               Height          =   255
               Left            =   0
               TabIndex        =   161
               Top             =   690
               Width           =   2700
            End
            Begin VB.Label lblDescription 
               Caption         =   "ppm"
               Height          =   255
               Index           =   131
               Left            =   3360
               TabIndex        =   160
               Top             =   360
               Width           =   600
            End
            Begin VB.Label lblDescription 
               Caption         =   "Minimum difference in average mass"
               Height          =   255
               Index           =   114
               Left            =   0
               TabIndex        =   158
               Top             =   360
               Width           =   2700
            End
            Begin VB.Label lblDescription 
               Caption         =   "peaks"
               Height          =   255
               Index           =   127
               Left            =   3360
               TabIndex        =   165
               Top             =   1020
               Width           =   600
            End
            Begin VB.Label lblDescription 
               Caption         =   "Maximum peak count to split feature"
               Height          =   255
               Index           =   122
               Left            =   0
               TabIndex        =   163
               Top             =   1020
               Width           =   2700
            End
            Begin VB.Label lblDescription 
               Caption         =   "% of max"
               Height          =   255
               Index           =   128
               Left            =   3360
               TabIndex        =   168
               Top             =   1350
               Width           =   705
            End
            Begin VB.Label lblDescription 
               Caption         =   "Peak picking intensity threshold"
               Height          =   255
               Index           =   123
               Left            =   0
               TabIndex        =   166
               Top             =   1350
               Width           =   2700
            End
            Begin VB.Label lblDescription 
               Caption         =   "scans"
               Height          =   255
               Index           =   129
               Left            =   3360
               TabIndex        =   171
               Top             =   1680
               Width           =   600
            End
            Begin VB.Label lblDescription 
               Caption         =   "Peak picking minimum width"
               Height          =   255
               Index           =   124
               Left            =   0
               TabIndex        =   169
               Top             =   1680
               Width           =   2700
            End
            Begin VB.Label lblDescription 
               BackStyle       =   0  'Transparent
               Caption         =   "Max size of scan gap in the feature"
               Height          =   255
               Index           =   125
               Left            =   0
               TabIndex        =   172
               Top             =   2010
               Width           =   2655
            End
            Begin VB.Label lblDescription 
               Caption         =   "scans"
               Height          =   255
               Index           =   130
               Left            =   3360
               TabIndex        =   174
               Top             =   2010
               Width           =   600
            End
            Begin VB.Label lblDescription 
               Caption         =   "Scan gap behavior:"
               Height          =   255
               Index           =   126
               Left            =   0
               TabIndex        =   196
               Top             =   2370
               Width           =   1620
            End
         End
      End
      Begin VB.Frame fraUMCIonNetOptions 
         Caption         =   "LC-MS Feature Ion Network Options"
         Height          =   4095
         Left            =   -74760
         TabIndex        =   104
         Top             =   3600
         Width           =   9135
         Begin VB.Frame fraOptionFrame 
            Height          =   3050
            Index           =   11
            Left            =   120
            TabIndex        =   108
            Top             =   930
            Width           =   8535
            Begin VB.CheckBox chkUse 
               Caption         =   "Use"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   113
               Top             =   840
               Width           =   700
            End
            Begin VB.ComboBox cmbData 
               Height          =   315
               Index           =   0
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   114
               Top             =   780
               Width           =   2175
            End
            Begin VB.TextBox txtWeightingFactor 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   4080
               TabIndex        =   116
               Text            =   "1"
               Top             =   780
               Width           =   615
            End
            Begin VB.CheckBox chkUse 
               Caption         =   "Use"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   121
               Top             =   1200
               Width           =   700
            End
            Begin VB.CheckBox chkUse 
               Caption         =   "Use"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   129
               Top             =   1560
               Width           =   700
            End
            Begin VB.ComboBox cmbData 
               Height          =   315
               Index           =   1
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   1140
               Width           =   2175
            End
            Begin VB.ComboBox cmbData 
               Height          =   315
               Index           =   2
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   130
               Top             =   1500
               Width           =   2175
            End
            Begin VB.TextBox txtWeightingFactor 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   4080
               TabIndex        =   124
               Text            =   "1"
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtWeightingFactor 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   4080
               TabIndex        =   132
               Text            =   "1"
               Top             =   1500
               Width           =   615
            End
            Begin VB.ComboBox cmbMetricType 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   110
               Top             =   300
               Width           =   1695
            End
            Begin VB.TextBox txtRejectLongConnections 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2520
               TabIndex        =   154
               Text            =   "1"
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox txtNETType 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4320
               TabIndex        =   112
               Text            =   "1"
               Top             =   300
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CheckBox chkUse 
               Caption         =   "Use"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   137
               Top             =   1920
               Width           =   700
            End
            Begin VB.CheckBox chkUse 
               Caption         =   "Use"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   145
               Top             =   2280
               Width           =   700
            End
            Begin VB.ComboBox cmbData 
               Height          =   315
               Index           =   3
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   1860
               Width           =   2175
            End
            Begin VB.ComboBox cmbData 
               Height          =   315
               Index           =   4
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   146
               Top             =   2220
               Width           =   2175
            End
            Begin VB.TextBox txtWeightingFactor 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   4080
               TabIndex        =   140
               Text            =   "1"
               Top             =   1860
               Width           =   615
            End
            Begin VB.TextBox txtWeightingFactor 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   4080
               TabIndex        =   148
               Text            =   "1"
               Top             =   2220
               Width           =   615
            End
            Begin VB.ComboBox cmbConstraint 
               Height          =   315
               Index           =   0
               Left            =   5640
               Style           =   2  'Dropdown List
               TabIndex        =   118
               Top             =   780
               Width           =   975
            End
            Begin VB.TextBox txtConstraint 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   6720
               TabIndex        =   119
               Text            =   "0.1"
               Top             =   780
               Width           =   735
            End
            Begin VB.ComboBox cmbConstraint 
               Height          =   315
               Index           =   1
               Left            =   5640
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   1140
               Width           =   975
            End
            Begin VB.ComboBox cmbConstraint 
               Height          =   315
               Index           =   2
               Left            =   5640
               Style           =   2  'Dropdown List
               TabIndex        =   134
               Top             =   1500
               Width           =   975
            End
            Begin VB.ComboBox cmbConstraint 
               Height          =   315
               Index           =   3
               Left            =   5640
               Style           =   2  'Dropdown List
               TabIndex        =   142
               Top             =   1860
               Width           =   975
            End
            Begin VB.ComboBox cmbConstraint 
               Height          =   315
               Index           =   4
               Left            =   5640
               Style           =   2  'Dropdown List
               TabIndex        =   150
               Top             =   2222
               Width           =   975
            End
            Begin VB.TextBox txtConstraint 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   6720
               TabIndex        =   127
               Text            =   "0.1"
               Top             =   1140
               Width           =   735
            End
            Begin VB.TextBox txtConstraint 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   6720
               TabIndex        =   135
               Text            =   "0.1"
               Top             =   1500
               Width           =   735
            End
            Begin VB.TextBox txtConstraint 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   6720
               TabIndex        =   143
               Text            =   "0.1"
               Top             =   1860
               Width           =   735
            End
            Begin VB.TextBox txtConstraint 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   6720
               TabIndex        =   151
               Text            =   "0.1"
               Top             =   2222
               Width           =   735
            End
            Begin VB.ComboBox cmbConstraintUnits 
               Height          =   315
               Index           =   0
               Left            =   7560
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   780
               Width           =   855
            End
            Begin VB.ComboBox cmbConstraintUnits 
               Height          =   315
               Index           =   1
               Left            =   7560
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   1140
               Width           =   855
            End
            Begin VB.ComboBox cmbConstraintUnits 
               Height          =   315
               Index           =   2
               Left            =   7560
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   1500
               Width           =   855
            End
            Begin VB.ComboBox cmbConstraintUnits 
               Height          =   315
               Index           =   3
               Left            =   7560
               Style           =   2  'Dropdown List
               TabIndex        =   144
               Top             =   1860
               Width           =   855
            End
            Begin VB.ComboBox cmbConstraintUnits 
               Height          =   315
               Index           =   4
               Left            =   7560
               Style           =   2  'Dropdown List
               TabIndex        =   152
               Top             =   2222
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Wt. Factor"
               Height          =   255
               Index           =   0
               Left            =   3240
               TabIndex        =   115
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Wt. Factor"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   123
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Wt. Factor"
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   131
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label lblDescription 
               Caption         =   "Metric Type"
               Height          =   255
               Index           =   79
               Left            =   240
               TabIndex        =   109
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblDescription 
               Caption         =   "Net Type"
               Height          =   255
               Index           =   80
               Left            =   3240
               TabIndex        =   111
               Top             =   360
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Wt. Factor"
               Height          =   255
               Index           =   4
               Left            =   3240
               TabIndex        =   139
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Wt. Factor"
               Height          =   255
               Index           =   5
               Left            =   3240
               TabIndex        =   147
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label lblDescription 
               Caption         =   "Reject connection longer than"
               Height          =   255
               Index           =   81
               Left            =   240
               TabIndex        =   153
               Top             =   2670
               Width           =   2175
            End
            Begin VB.Label Label2 
               Caption         =   "Constraint"
               Height          =   255
               Index           =   6
               Left            =   4800
               TabIndex        =   117
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Constraint"
               Height          =   255
               Index           =   7
               Left            =   4800
               TabIndex        =   125
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Constraint"
               Height          =   255
               Index           =   8
               Left            =   4800
               TabIndex        =   133
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Constraint"
               Height          =   255
               Index           =   9
               Left            =   4800
               TabIndex        =   141
               Top             =   1920
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Constraint"
               Height          =   255
               Index           =   10
               Left            =   4800
               TabIndex        =   149
               Top             =   2280
               Width           =   735
            End
         End
         Begin VB.CheckBox chkUMCIonNetMakeSingleMemberClasses 
            Caption         =   "Make single-member classes"
            Height          =   255
            Left            =   3720
            TabIndex        =   107
            ToolTipText     =   "Make single-member classes from unconnected nodes"
            Top             =   560
            Width           =   2535
         End
         Begin VB.ComboBox cmbUMCRepresentative 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   540
            Width           =   3255
         End
         Begin VB.Label lblDescription 
            Caption         =   "Class Representative"
            Height          =   255
            Index           =   78
            Left            =   120
            TabIndex        =   105
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkInterpolateMissingIons 
         Caption         =   "Interpolate abundances across gaps"
         Height          =   375
         Left            =   -71760
         TabIndex        =   77
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtInterpolateMaxGapSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70200
         TabIndex        =   79
         Text            =   "0"
         Top             =   2280
         Width           =   495
      End
      Begin VB.ComboBox cmbUMCAbu 
         Height          =   315
         ItemData        =   "frmEditAnalysisSettings.frx":6EF13
         Left            =   -74760
         List            =   "frmEditAnalysisSettings.frx":6EF15
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   2205
         Width           =   2775
      End
      Begin VB.ComboBox cmbUMCMW 
         Height          =   315
         ItemData        =   "frmEditAnalysisSettings.frx":6EF17
         Left            =   -74760
         List            =   "frmEditAnalysisSettings.frx":6EF19
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Frame fraUMCSearch200x 
         Caption         =   "UMC 2002/2003 Options"
         Height          =   3855
         Left            =   -74760
         TabIndex        =   81
         Top             =   3600
         Width           =   6375
         Begin VB.ComboBox cmbCountType 
            Height          =   315
            ItemData        =   "frmEditAnalysisSettings.frx":6EF1B
            Left            =   2880
            List            =   "frmEditAnalysisSettings.frx":6EF1D
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtHoleNum 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5520
            TabIndex        =   95
            Text            =   "0"
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox txtHoleSize 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   5520
            TabIndex        =   97
            Text            =   "0"
            Top             =   1500
            Width           =   495
         End
         Begin VB.Frame fraUMC2003UniqueOptions 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   2880
            TabIndex        =   101
            Top             =   2400
            Width           =   3255
            Begin VB.CheckBox chkAllowSharing 
               Caption         =   "Allow members sharing among classes"
               Height          =   255
               Left            =   0
               TabIndex        =   102
               Top             =   0
               Width           =   3015
            End
            Begin VB.CheckBox chkUMCDefRequireIdenticalCharge 
               Caption         =   "Require Identical Charge (not yet implemented)"
               Enabled         =   0   'False
               Height          =   495
               Left            =   0
               TabIndex        =   103
               Top             =   360
               Width           =   3015
            End
         End
         Begin VB.Frame fraUMC2002UniqueOptions 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   2880
            TabIndex        =   98
            Top             =   1920
            Width           =   3255
            Begin VB.TextBox txtHolePct 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               TabIndex        =   100
               Text            =   "0"
               Top             =   60
               Width           =   495
            End
            Begin VB.Label lblDescription 
               Caption         =   "Percentage of allowed scan holes in the Unique Mass Class:"
               Height          =   495
               Index           =   59
               Left            =   0
               TabIndex        =   99
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.CheckBox chkUMCShrinkingBoxWeightAverageMassByIntensity 
            Caption         =   "Weight average mass by intensity during shrinking box search"
            Height          =   735
            Left            =   120
            TabIndex        =   91
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Molecular Mass Tolerance"
            Height          =   1215
            Index           =   13
            Left            =   120
            TabIndex        =   86
            Top             =   1680
            Width           =   2175
            Begin VB.TextBox txtUMCSearchMassTol 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   160
               TabIndex        =   88
               Text            =   "10"
               Top             =   640
               Width           =   735
            End
            Begin VB.OptionButton optUMCSearchTolType 
               Caption         =   "&ppm"
               Height          =   255
               Index           =   0
               Left            =   1020
               TabIndex        =   89
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optUMCSearchTolType 
               Caption         =   "&Dalton"
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   90
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblDescription 
               Caption         =   "Tolerance"
               Height          =   255
               Index           =   15
               Left            =   160
               TabIndex        =   87
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Molecular Mass Field"
            Height          =   1215
            Index           =   12
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   1935
            Begin VB.OptionButton optUMCSearchMWField 
               Caption         =   "A&verage"
               Height          =   255
               Index           =   0
               Left            =   80
               TabIndex        =   83
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton optUMCSearchMWField 
               Caption         =   "&Monoisotopic"
               Height          =   255
               Index           =   1
               Left            =   80
               TabIndex        =   84
               Top             =   540
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton optUMCSearchMWField 
               Caption         =   "&The Most Abundant"
               Height          =   255
               Index           =   2
               Left            =   80
               TabIndex        =   85
               Top             =   840
               Width           =   1815
            End
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum size of scan gap in the Unique Mass Class:"
            Height          =   495
            Index           =   66
            Left            =   2880
            TabIndex        =   96
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label lblDescription 
            Caption         =   "Maximum number of scan gaps in the Unique Mass Class:"
            Height          =   495
            Index           =   65
            Left            =   2880
            TabIndex        =   94
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblDescription 
            Caption         =   "Count Type"
            Height          =   255
            Index           =   62
            Left            =   2880
            TabIndex        =   92
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "MT Tag Database Settings"
         Height          =   3015
         Index           =   44
         Left            =   -74880
         TabIndex        =   198
         Top             =   480
         Width           =   6735
         Begin VB.TextBox txtCurrentDBScoreThreshold 
            Height          =   285
            Index           =   2
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   538
            Text            =   "0"
            Top             =   2520
            Width           =   495
         End
         Begin VB.CheckBox chkCurrentDBLimitToPMTsFromDataset 
            Caption         =   "Limit to MT tags from Dataset for Job"
            Height          =   375
            Left            =   120
            TabIndex        =   208
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox txtCurrentDBScoreThreshold 
            Height          =   285
            Index           =   1
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   211
            Text            =   "0"
            Top             =   2220
            Width           =   495
         End
         Begin VB.TextBox txtCurrentDBName 
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   200
            Top             =   240
            Width           =   5055
         End
         Begin VB.CheckBox chkCurrentDBAMTsOnly 
            Caption         =   "AMT's Only"
            Height          =   255
            Left            =   120
            TabIndex        =   206
            Top             =   1400
            Width           =   1455
         End
         Begin VB.TextBox txtCurrentDBAllowedModifications 
            Height          =   615
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   209
            Top             =   1200
            Width           =   3255
         End
         Begin VB.CheckBox chkCurrentDBConfirmedOnly 
            Caption         =   "Confirmed Only"
            Height          =   255
            Left            =   120
            TabIndex        =   205
            Top             =   1160
            Width           =   1455
         End
         Begin VB.CheckBox chkCurrentDBLockersOnly 
            Caption         =   "Lockers Only"
            Height          =   255
            Left            =   120
            TabIndex        =   207
            Top             =   1640
            Width           =   1455
         End
         Begin VB.CommandButton cmdSelectMassTags 
            Caption         =   "&Select MT tags"
            Height          =   375
            Left            =   4920
            TabIndex        =   203
            ToolTipText     =   "Select the MT tags to use"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtCurrentDBScoreThreshold 
            Height          =   285
            Index           =   0
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   210
            Text            =   "0"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Minimum XCorr"
            Height          =   255
            Index           =   162
            Left            =   3360
            TabIndex        =   537
            Top             =   1935
            Width           =   2025
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Discrim. Score"
            Height          =   255
            Index           =   145
            Left            =   3360
            TabIndex        =   536
            Top             =   2240
            Width           =   1545
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min Pep Prophet"
            Height          =   255
            Index           =   119
            Left            =   3360
            TabIndex        =   535
            Top             =   2540
            Width           =   1545
         End
         Begin VB.Label lblCurrentDBInternalStdExplicitOrMTSubset 
            Caption         =   "Explicit Internal Standard:"
            Height          =   255
            Left            =   120
            TabIndex        =   204
            Top             =   900
            Width           =   5175
         End
         Begin VB.Label lblCurrentDBNETValueType 
            Caption         =   "Avg Obs NET - from DB"
            Height          =   255
            Left            =   120
            TabIndex        =   213
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label lblDescription 
            Caption         =   "NET Value Type:"
            Height          =   255
            Index           =   120
            Left            =   120
            TabIndex        =   212
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblDescription 
            Caption         =   "Database Name:"
            Height          =   255
            Index           =   116
            Left            =   120
            TabIndex        =   199
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label lblDescription 
            Caption         =   "Count of selected MT tags in current DB:"
            Height          =   255
            Index           =   117
            Left            =   120
            TabIndex        =   201
            Top             =   660
            Width           =   3255
         End
         Begin VB.Label lblCurrentDBMassTagCount 
            Caption         =   "0"
            Height          =   255
            Left            =   3480
            TabIndex        =   202
            Top             =   660
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdSelectOtherDB 
         Caption         =   "&Select other DB"
         Height          =   375
         Left            =   -74880
         TabIndex        =   215
         ToolTipText     =   "Select the MT tags to use"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Frame fraSelectingMassTags 
         Height          =   1800
         Left            =   -69600
         TabIndex        =   222
         Top             =   4320
         Width           =   4695
         Begin VB.CommandButton cmdSelectingMassTagsOK 
            Caption         =   "&Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   225
            ToolTipText     =   "Select MT tags to load for search"
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdSelectingMassTagsCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   720
            TabIndex        =   224
            ToolTipText     =   "Select MT tags to load for search"
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblSelectingMassTags 
            Caption         =   $"frmEditAnalysisSettings.frx":6EF1F
            Height          =   1095
            Left            =   120
            TabIndex        =   223
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "General Database options"
         Height          =   1515
         Index           =   16
         Left            =   -74880
         TabIndex        =   216
         Top             =   4560
         Width           =   4815
         Begin VB.TextBox txtDBConnectionTimeoutSeconds 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3720
            TabIndex        =   221
            Text            =   "30"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtDBConnectionRetryAttemptMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3720
            TabIndex        =   219
            Text            =   "5"
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkUseMassTagsWithNullNET 
            Caption         =   "Use MT tags with Null NET"
            Height          =   255
            Left            =   240
            TabIndex        =   217
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblDescription 
            Caption         =   "Database Connection Timeout (seconds)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   220
            Top             =   1110
            Width           =   3375
         End
         Begin VB.Label lblDescription 
            Caption         =   "Database Connection Maximum Retry Count"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   218
            Top             =   750
            Width           =   3375
         End
      End
      Begin VB.ComboBox cboUMCSearchMode 
         Height          =   315
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   960
         Width           =   2055
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "M/Z Range (isotopic data only)"
         Height          =   1180
         Index           =   8
         Left            =   -69960
         TabIndex        =   50
         Top             =   2800
         Width           =   4335
         Begin VB.TextBox txtIsoMinMaxMZ 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   55
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtIsoMinMaxMZ 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3000
            TabIndex        =   53
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkIsoMZRange 
            Caption         =   "Use only Isotopic data with m/z within range"
            Height          =   495
            Left            =   240
            TabIndex        =   51
            Top             =   405
            Width           =   1935
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   54
            Top             =   765
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min"
            Height          =   255
            Index           =   54
            Left            =   2400
            TabIndex        =   52
            Top             =   405
            Width           =   495
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Specialized filters"
         Height          =   1920
         Index           =   9
         Left            =   -69960
         TabIndex        =   60
         Top             =   5300
         Width           =   5895
         Begin VB.ComboBox cboEvenOddScanNumber 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtDupTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            TabIndex        =   65
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkDupElimination 
            Caption         =   "&Exclude duplicates from the Isotopic data"
            Height          =   375
            Left            =   3240
            TabIndex        =   63
            ToolTipText     =   "Isotopic Data Only"
            Top             =   200
            Width           =   1935
         End
         Begin VB.CheckBox chkSecGuessElimination 
            Caption         =   "Exclude less likely guess"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   62
            ToolTipText     =   "Check this option to keep data more likely (in comparisson with other data)"
            Top             =   600
            Width           =   2400
         End
         Begin VB.CheckBox chkSecGuessElimination 
            Caption         =   "Exclude &second guess"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   61
            ToolTipText     =   "Check this option to keep data with better fit from calculation"
            Top             =   260
            Width           =   2400
         End
         Begin VB.Label lblDescription 
            Caption         =   "Even/Odd Scan Number Filtering (use this filter for DREAMS-based data files)"
            Height          =   450
            Index           =   60
            Left            =   120
            TabIndex        =   66
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label lblDescription 
            Caption         =   "Duplicate Tolerance"
            Height          =   375
            Index           =   67
            Left            =   3240
            TabIndex        =   64
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.TextBox txtPEKFileExtensionPreferenceOrder 
         Height          =   285
         Left            =   -71880
         TabIndex        =   2
         Text            =   "_ic.csv, .csv, _ic.pek,_s.pek,.pek,DeCal.pek-3,.pek-3"
         Top             =   480
         Width           =   7815
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Isotopic data options"
         Height          =   1755
         Index           =   7
         Left            =   -69960
         TabIndex        =   38
         Top             =   960
         Width           =   5865
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Use Isotopic MW From"
            Height          =   1215
            Index           =   14
            Left            =   3840
            TabIndex        =   46
            Top             =   240
            Width           =   1935
            Begin VB.OptionButton optIsoDataFrom 
               Caption         =   "Most &abundant"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   49
               Top             =   840
               Width           =   1455
            End
            Begin VB.OptionButton optIsoDataFrom 
               Caption         =   "&Average field"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton optIsoDataFrom 
               Caption         =   "&Monoisotopic field"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   48
               Top             =   540
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.TextBox txtExcludeIsoByFit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   40
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkExcludeIsoByFit 
            Caption         =   "Ex&clude data with calculated isotopic fit worse (higher) than:"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   310
            Width           =   2895
         End
         Begin VB.TextBox txtIsoMinMaxCS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   45
            Text            =   "0"
            Top             =   1210
            Width           =   735
         End
         Begin VB.TextBox txtIsoMinMaxCS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2880
            TabIndex        =   43
            Text            =   "0"
            Top             =   910
            Width           =   735
         End
         Begin VB.CheckBox chkIsoUseCSRange 
            Caption         =   "Use only charge states within range"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   1030
            Width           =   1815
         End
         Begin VB.Label lblDescription 
            Caption         =   "Last C.S."
            Height          =   255
            Index           =   26
            Left            =   2040
            TabIndex        =   44
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label lblDescription 
            Caption         =   "First C.S."
            Height          =   255
            Index           =   27
            Left            =   2040
            TabIndex        =   42
            Top             =   975
            Width           =   735
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Scan Range and NET Range"
         Height          =   1935
         Index           =   6
         Left            =   -74760
         TabIndex        =   27
         Top             =   5400
         Width           =   4575
         Begin VB.TextBox txtRestrictGANETRangeMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   37
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtRestrictGANETRangeMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3000
            TabIndex        =   35
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkRestrictGANETRange 
            Caption         =   "Use only scans whose NET value is within range"
            Height          =   615
            Left            =   120
            TabIndex        =   33
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtRestrictScanRangeMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   32
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtRestrictScanRangeMinMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3000
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkRestrictScanRange 
            Caption         =   "Use only scans within range"
            Height          =   615
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max"
            Height          =   255
            Index           =   29
            Left            =   2280
            TabIndex        =   36
            Top             =   1485
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min"
            Height          =   255
            Index           =   28
            Left            =   2280
            TabIndex        =   34
            Top             =   1125
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max"
            Height          =   255
            Index           =   30
            Left            =   2280
            TabIndex        =   31
            Top             =   645
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min"
            Height          =   255
            Index           =   31
            Left            =   2280
            TabIndex        =   29
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Abundance Range"
         Height          =   2175
         Index           =   5
         Left            =   -74760
         TabIndex        =   15
         Top             =   3120
         Width           =   4575
         Begin VB.TextBox txtCSMinMaxAbu 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   25
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtCSMinMaxAbu 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   23
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox chkCSAbuRange 
            Caption         =   "Use only Charge State data with abundance within range  "
            Height          =   615
            Left            =   240
            TabIndex        =   21
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkIsoAbuRange 
            Caption         =   "Use only Isotopic data with abundance within range  "
            Height          =   615
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtIsoMinMaxAbu 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtIsoMinMaxAbu 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkCSIsoSameRangeAbu 
            Caption         =   "Same range for Charge State and Isotopic data"
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   1680
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max"
            Height          =   255
            Index           =   33
            Left            =   2520
            TabIndex        =   24
            Top             =   1365
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min"
            Height          =   255
            Index           =   32
            Left            =   2520
            TabIndex        =   22
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min"
            Height          =   255
            Index           =   34
            Left            =   2520
            TabIndex        =   17
            Top             =   285
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max"
            Height          =   255
            Index           =   35
            Left            =   2520
            TabIndex        =   19
            Top             =   645
            Width           =   495
         End
      End
      Begin VB.Frame fraOptionFrame 
         Caption         =   "Molecular Mass Range"
         Height          =   2055
         Index           =   4
         Left            =   -74760
         TabIndex        =   3
         Top             =   960
         Width           =   4575
         Begin VB.CheckBox chkCSMWRange 
            Caption         =   "Use only Charge State data with molecular mass within range  "
            Height          =   615
            Left            =   240
            TabIndex        =   9
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtCSMinMaxMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   11
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtCSMinMaxMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   13
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox chkCSIsoSameRangeMW 
            Caption         =   "Same range for Charge State and Isotopic data"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   1680
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.TextBox txtIsoMinMaxMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtIsoMinMaxMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkIsoMWRange 
            Caption         =   "Use only Isotopic data with molecular mass within range  "
            Height          =   615
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min"
            Height          =   255
            Index           =   39
            Left            =   2520
            TabIndex        =   10
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max"
            Height          =   255
            Index           =   38
            Left            =   2520
            TabIndex        =   12
            Top             =   1365
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Max"
            Height          =   255
            Index           =   36
            Left            =   2520
            TabIndex        =   7
            Top             =   645
            Width           =   495
         End
         Begin VB.Label lblDescription 
            Caption         =   "Min"
            Height          =   255
            Index           =   37
            Left            =   2520
            TabIndex        =   5
            Top             =   285
            Width           =   495
         End
      End
      Begin TabDlg.SSTab tbsNETOptions 
         Height          =   3405
         Left            =   4200
         TabIndex        =   302
         Top             =   480
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   6006
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Robust NET Options"
         TabPicture(0)   =   "frmEditAnalysisSettings.frx":6EFF2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "pctOptionsInIniFile(2)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Charge Selection"
         TabPicture(1)   =   "frmEditAnalysisSettings.frx":6F00E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "pctOptionsInIniFile(5)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "MS Warp Options"
         TabPicture(2)   =   "frmEditAnalysisSettings.frx":6F02A
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "lblDescription(160)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "lblDescription(155)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "fraOptionFrame(34)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "fraOptionFrame(35)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "cboMSWarpMassCalibrationType"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).ControlCount=   5
         Begin VB.PictureBox pctOptionsInIniFile 
            Height          =   1125
            Index           =   5
            Left            =   -74880
            Picture         =   "frmEditAnalysisSettings.frx":6F046
            ScaleHeight     =   1065
            ScaleWidth      =   6255
            TabIndex        =   546
            Top             =   480
            Width           =   6315
         End
         Begin VB.PictureBox pctOptionsInIniFile 
            Height          =   2655
            Index           =   2
            Left            =   -74760
            Picture         =   "frmEditAnalysisSettings.frx":767E8
            ScaleHeight     =   2595
            ScaleWidth      =   5595
            TabIndex        =   541
            Top             =   480
            Width           =   5655
         End
         Begin VB.ComboBox cboMSWarpMassCalibrationType 
            Height          =   315
            ItemData        =   "frmEditAnalysisSettings.frx":869FE
            Left            =   2040
            List            =   "frmEditAnalysisSettings.frx":86A00
            Style           =   2  'Dropdown List
            TabIndex        =   529
            Top             =   2880
            Width           =   3165
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "Mass Warp Options"
            Height          =   2175
            Index           =   35
            Left            =   3120
            TabIndex        =   315
            Top             =   600
            Width           =   2775
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   17
               Left            =   1920
               TabIndex        =   323
               Text            =   "100"
               Top             =   1400
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   16
               Left            =   1920
               TabIndex        =   321
               Text            =   "20"
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   15
               Left            =   1920
               TabIndex        =   319
               Text            =   "2"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   14
               Left            =   1920
               TabIndex        =   317
               Text            =   "50"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   18
               Left            =   1920
               TabIndex        =   325
               Text            =   "50"
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label lblDescription 
               Caption         =   "Mass Window (ppm)"
               Height          =   255
               Index           =   154
               Left            =   120
               TabIndex        =   316
               Top             =   270
               Width           =   2145
            End
            Begin VB.Label lblDescription 
               Caption         =   "# of mass delta bins"
               Height          =   255
               Index           =   153
               Left            =   120
               TabIndex        =   322
               Top             =   1430
               Width           =   1785
            End
            Begin VB.Label lblDescription 
               Caption         =   "# of x-axis time slices"
               Height          =   255
               Index           =   152
               Left            =   120
               TabIndex        =   320
               Top             =   990
               Width           =   1785
            End
            Begin VB.Label lblDescription 
               Caption         =   "Spline Order"
               Height          =   255
               Index           =   151
               Left            =   120
               TabIndex        =   318
               Top             =   630
               Width           =   2145
            End
            Begin VB.Label lblDescription 
               Caption         =   "Max jump"
               Height          =   255
               Index           =   150
               Left            =   120
               TabIndex        =   324
               Top             =   1830
               Width           =   1785
            End
         End
         Begin VB.Frame fraOptionFrame 
            Caption         =   "NET Warp Options"
            Height          =   2175
            Index           =   34
            Left            =   120
            TabIndex        =   304
            Top             =   600
            Width           =   2775
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   13
               Left            =   1920
               TabIndex        =   314
               Text            =   "2"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   9
               Left            =   1920
               TabIndex        =   306
               Text            =   "100"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   10
               Left            =   1920
               TabIndex        =   308
               Text            =   "3"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   11
               Left            =   1920
               TabIndex        =   310
               Text            =   "2"
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtRobustNETOption 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   12
               Left            =   1920
               TabIndex        =   312
               Text            =   "5"
               Top             =   1400
               Width           =   735
            End
            Begin VB.Label lblDescription 
               Caption         =   "Match Promiscuity"
               Height          =   255
               Index           =   118
               Left            =   120
               TabIndex        =   313
               Top             =   1830
               Width           =   1785
            End
            Begin VB.Label lblDescription 
               Caption         =   "Max Distortion"
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   307
               Top             =   630
               Width           =   2145
            End
            Begin VB.Label lblDescription 
               Caption         =   "Contraction Factor"
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   309
               Top             =   990
               Width           =   1785
            End
            Begin VB.Label lblDescription 
               Caption         =   "Minimum MS/MS Obs Count for MT tags"
               Height          =   495
               Index           =   158
               Left            =   120
               TabIndex        =   311
               Top             =   1300
               Width           =   1785
            End
            Begin VB.Label lblDescription 
               Caption         =   "Number of Sections"
               Height          =   255
               Index           =   159
               Left            =   120
               TabIndex        =   305
               Top             =   270
               Width           =   2145
            End
         End
         Begin VB.Label lblDescription 
            Caption         =   "Mass Calibration Type"
            Height          =   255
            Index           =   155
            Left            =   240
            TabIndex        =   326
            Top             =   2880
            Width           =   1665
         End
         Begin VB.Label lblDescription 
            Caption         =   "Note: Mass and NET tolerance values are set under Search Tolerances (at the left)"
            Height          =   255
            Index           =   160
            Left            =   120
            TabIndex        =   303
            Top             =   360
            Width           =   6345
         End
      End
      Begin VB.CheckBox chkSkipGANETSlopeAndInterceptComputation 
         Caption         =   "Skip NET Slope and Intercept Computation"
         Height          =   255
         Left            =   240
         TabIndex        =   292
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblDescription 
         Caption         =   "Adjustment Mode"
         Height          =   240
         Index           =   156
         Left            =   6600
         TabIndex        =   328
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblMassCalibrationRefinementDescription 
         Caption         =   "Mass Calibration Refinement Description"
         Height          =   975
         Left            =   -69120
         TabIndex        =   527
         Top             =   6480
         Width           =   4695
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditAnalysisSettings.frx":86A02
         Height          =   375
         Index           =   136
         Left            =   -74760
         TabIndex        =   520
         Top             =   5160
         Width           =   9855
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Most Abu Charge State Group Type"
         Height          =   255
         Index           =   61
         Left            =   -74760
         TabIndex        =   75
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label lblDescription 
         Caption         =   "Pairs Identification Mode"
         Height          =   255
         Index           =   16
         Left            =   -74760
         TabIndex        =   230
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblToleranceRefinementExplanation 
         Caption         =   "Tolerance Refinement explanation"
         Height          =   1215
         Left            =   -74685
         TabIndex        =   330
         Top             =   600
         Width           =   9495
      End
      Begin VB.Label lblDescription 
         Caption         =   "Database Search Mode"
         Height          =   255
         Index           =   70
         Left            =   -74760
         TabIndex        =   404
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label lblToleranceRefinementWarning 
         Caption         =   $"frmEditAnalysisSettings.frx":86A8B
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   -74760
         TabIndex        =   505
         Top             =   4560
         Width           =   3855
      End
      Begin VB.Label lblSavingAndExportingStatus 
         Caption         =   "Saving and exporting is enabled."
         Height          =   255
         Left            =   -74760
         TabIndex        =   467
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label lblDescription 
         Caption         =   "Maximum size of gap to interpolate:"
         Height          =   420
         Index           =   58
         Left            =   -71760
         TabIndex        =   78
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblDescription 
         Caption         =   "Class Molecular Mass"
         Height          =   255
         Index           =   64
         Left            =   -74760
         TabIndex        =   71
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblDescription 
         Caption         =   "Class Abundance"
         Height          =   255
         Index           =   63
         Left            =   -74760
         TabIndex        =   73
         Top             =   1965
         Width           =   1335
      End
      Begin VB.Label lblDescription 
         Caption         =   "The above settings are overridden during PRISM-initiated VIPER Analysis by the values in T_Peak_Matching_Task"
         Height          =   255
         Index           =   121
         Left            =   -74880
         TabIndex        =   214
         Top             =   3720
         Width           =   8415
      End
      Begin VB.Label lblDescription 
         Caption         =   "LC-MS Feature Search Mode"
         Height          =   255
         Index           =   115
         Left            =   -74760
         TabIndex        =   69
         Top             =   975
         Width           =   2655
      End
      Begin VB.Label lblDescription 
         Caption         =   "Input File extension preference order"
         Height          =   255
         Index           =   50
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame fraOptionFrame 
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   43
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   493
      Top             =   8400
      Width           =   11055
      Begin VB.CommandButton cmdRevert 
         Caption         =   "Revert"
         Height          =   375
         Left            =   9480
         TabIndex        =   501
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdIniFileOpen 
         Caption         =   "Read from Settings File"
         Height          =   375
         Left            =   7320
         TabIndex        =   499
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton cmdIniFileSave 
         Caption         =   "Save to Settings File"
         Height          =   375
         Left            =   7320
         TabIndex        =   500
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdResetToDefaults 
         Caption         =   "Set to Defaults"
         Height          =   375
         Left            =   240
         TabIndex        =   496
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   9480
         TabIndex        =   502
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cboGelFileInMemoryToUse 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   498
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton cmdApplyToSelectedGel 
         Caption         =   "Apply to Selected Gel"
         Height          =   375
         Left            =   0
         TabIndex        =   495
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton cmdReadFromSelectedGel 
         Caption         =   "Read from Selected Gel"
         Height          =   375
         Left            =   0
         TabIndex        =   494
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lblGelFileInMemoryToUse 
         Caption         =   "Gel file (in memory) to read or update"
         Height          =   255
         Left            =   2280
         TabIndex        =   497
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Label lblWorking 
      Caption         =   "Loading settings from .Ini file"
      Height          =   675
      Left            =   240
      TabIndex        =   503
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmEditAnalysisSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' This form can be used to view all of the analysis options in a concise, tabbed format
' Settings files (.ini files) can be saved or loaded for the current options
' The options can be applied to any of the .Gel files in memory
' Alternatively, the options can be read from any of the .Gel files in memory
'

Private Const MAX_MW = 1000000
Private Const MAX_MZ = 100000
Private Const max_scan = 1000000

Private Const OPTION_FRAME_DELTA_PAIRS = 37
Private Const OPTION_FRAME_LABEL_PAIRS = 38
Private Const OPTION_FRAME_CONTROLS = 43
Private Const OPTION_FRAME_CURRENT_CONNECTION = 44
Private Const OPTION_FRAME_DB_SEARCH_MT_LABELED = 46
Private Const OPTION_FRAME_GENERAL_DB_OPTIONS = 16
Private Const OPTION_FRAME_LEGACY_DB = 19
Private Const OPTION_FRAME_NET_WARP_OPTIONS = 34
Private Const OPTION_FRAME_MASS_WARP_OPTIONS = 35

Private Const OPTION_FRAME_NETSlopeRange = 20
Private Const OPTION_FRAME_NETInterceptRange = 21
Private Const OPTION_FRAME_MassShiftPPMRange = 24

Private Const LABEL_DESCRIPTION_PRISM_MESSAGE = 121
Private Const LABEL_DESCRIPTION_ALTERNATE_OUTPUT_FOLDER_PATH = 40
Private Const LABEL_DESCRIPTION_INTERNAL_STD_SEARCH_MODE = 112

Private Enum usmUMCSearchModeConstants
    usmUMC2002 = 0                      ' Obsolete search mode (July 2004)
    usmUMC2003 = 1
    usmUMCIonNet = 2
End Enum

Private Enum pimPairsIdentificationModeConstants
    pimNone = 0
    pimDelta = 1
    pimLabel = 2
End Enum

Private Enum cbamCheckBoxAdditionaMassConstants
    cbamPEO = 0
    cbamICATd0 = 1
    cbamICATd8 = 2
    cbamAlkylation = 3
End Enum

Private Enum cbnaCheckBoxNetAdjustConstants
    cbnaDecreaseMW = 0
    cbnaDecreaseNET = 1
End Enum

Private Enum uarUMCAutoRefineOptionsConstants
    uarRemoveLowIntensity = 0
    uarRemoveHighIntensity = 1
    uarRemoveShort = 2
    uarRemoveLong = 3
    uarTestLengthUsingScanRange = 4
    uarPercentMaxAbuToUseToGaugeLength = 5
    uarMinimumMemberCount = 6
    uarRemoveMaxLengthPctAllScans = 7
End Enum

Private Enum eosEvenOddScanConstants
    eosAllScans = 0
    eosOddScans = 1
    eosEvenScans = 2
End Enum

Private Enum rpmhRemovePairMemberHitsConstants
    rpmhRemoveLight = 0
    rpmhRemoveHeavy = 1
End Enum

Private Enum psoPairSearchOptionsConstants
    psoRequireMatchingChargeStatesForPairMembers = 0
    psoUseIdenticalChargesForER = 1
    psoAverageERsAllChargeStates = 2
    psoComputeERScanByScan = 3
    psoIReportEREnable = 4
    psoRemoveOutlierERs = 5
    psoRemoveOutlierERsIterate = 6
End Enum

Private Enum rnoRobustNetOptionsConstants
''    rnoNETSlopeStart = 0
''    rnoNETSlopeEnd = 1
''    rnoNETSlopeIncrement = 2
''    rnoNETInterceptStart = 3
''    rnoNETInterceptEnd = 4
''    rnoNETInterceptIncrement = 5
''    rnoMassShiftPPMStart = 6
''    rnoMassShiftPPMEnd = 7
''    rnoMassShiftPPMIncrement = 8
    rnoWarpNumberOfSections = 9
    rnoWarpMaxDistortion = 10
    rnoWarpContractionFactor = 11
    rnoWarpMinimumPMTTagObsCount = 12
    rnoWarpMatchPromiscuity = 13
    rnoWarpMassWindow = 14
    rnoWarpMassSplineOrder = 15
    rnoWarpNumXAxisSlices = 16
    rnoWarpNumMassDeltaBins = 17
    rnoWarpMassMaxJump = 18
End Enum

Private Enum trfToleranceRefinementFilterOptionsConstants
    trfMinimumHighNormalizedScore = 0
    trfMinimumDiscriminant = 1
    trfMinimumPeptideProphet = 2
    trfMinimumSLiC = 3
    trfMaximumAbundance = 4
End Enum

Private Enum trfCurrentDBScoreThresholdConstants
    cdstHighNormalizedScore = 0
    cdstDiscriminant = 1
    cdstPeptideProphet = 2
End Enum

Private Type udtSettingsContainerType
    Prefs As GelPrefs
    PrefsExpanded As udtPreferencesExpandedType
    UMCDef As UMCDefinition
    UMCIonNetDef As UMCIonNetDefinition
    UMCNetAdjDef As NetAdjDefinition
    UMCInternalStandards As udtInternalStandardsType
    AMTDef As SearchAMTDefinition
    DBSettings As udtDBSettingsType
End Type

Private Const SEARCH_N14 = 0
Private Const SEARCH_N15 = 1

Private Const MODS_FIXED = 0            ' Aka Static
Private Const MODS_DYNAMIC = 1
Private Const MODS_DECOY = 2

Private Const ASSUME_MT_NOT_LABELED = 0
Private Const ASSUME_MT_LABELED = 1

Private mCurrentSettings As udtSettingsContainerType
Private mSettingsSaved As udtSettingsContainerType

Private mAutoIniFileNameOverridden As Boolean
Private mUpdatingControls As Boolean
Private mUpdatingModMassControls As Boolean
''Private mRequestUpdateRobustNETIterationCount As Boolean

Private mGelIndexInstantiatingForm As Long
Private mActiveGelIndex As Long

Private mIniFilePathSaved As String

Private objSelectMassTags As FTICRAnalysis

Private WithEvents objMTConnectionSelector As DummyAnalysisInitiator
Attribute objMTConnectionSelector.VB_VarHelpID = -1

Private Sub AddNewDBSearchMode()
    Dim eDBSearchMode As dbsmDatabaseSearchModeConstants
    Dim strNewSearchMode As String
    Dim intIndex As Integer
    Dim intNewIndex As Integer
    
On Error GoTo AddNewDBSearchModeErrorHandler
    
    If lstDBSearchModes.ListCount >= MAX_AUTO_SEARCH_MODE_COUNT Then
        Exit Sub
    End If
    
    If mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount > MAX_AUTO_SEARCH_MODE_COUNT Then
        Debug.Assert False
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount = MAX_AUTO_SEARCH_MODE_COUNT
    End If
    
    If lstDBSearchModes.ListCount <> mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount Then
        Debug.Assert False
        lstDBSearchModes.Clear
        For intIndex = 0 To mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount
            lstDBSearchModes.AddItem mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intIndex).SearchMode
        Next intIndex
    End If
    
    If cboDBSearchMode.ListIndex < 0 Or cboDBSearchMode.ListIndex > DB_SEARCH_MODE_MAX_INDEX Then
        Exit Sub
    End If
    
    eDBSearchMode = cboDBSearchMode.ListIndex
    Select Case eDBSearchMode
    Case dbsmExportUMCsOnly: strNewSearchMode = AUTO_SEARCH_EXPORT_UMCS_ONLY
    Case dbsmIndividualPeaks: strNewSearchMode = AUTO_SEARCH_ORGANISM_MTDB
    Case dbsmIndividualPeaksInUMCsWithoutNET: strNewSearchMode = AUTO_SEARCH_UMC_MTDB
    Case dbsmIndividualPeaksInUMCsWithNET: strNewSearchMode = AUTO_SEARCH_UMC_HERETIC           ' No longer supported (June 2004)
    Case dbsmConglomerateUMCsWithNET: strNewSearchMode = AUTO_SEARCH_UMC_CONGLOMERATE
    Case dbsmIndividualPeaksInUMCsPaired: strNewSearchMode = AUTO_SEARCH_UMC_HERETIC_PAIRED     ' No longer supported (June 2004)
    Case dbsmIndividualPeaksInUMCsUnpaired: strNewSearchMode = AUTO_SEARCH_UMC_HERETIC_UNPAIRED ' No longer supported (June 2004)
    Case dbsmConglomerateUMCsPaired: strNewSearchMode = AUTO_SEARCH_UMC_CONGLOMERATE_PAIRED
    Case dbsmConglomerateUMCsUnpaired: strNewSearchMode = AUTO_SEARCH_UMC_CONGLOMERATE_UNPAIRED
    Case dbsmConglomerateUMCsLightPairsPlusUnpaired: strNewSearchMode = AUTO_SEARCH_UMC_CONGLOMERATE_LIGHT_PAIRS_PLUS_UNPAIRED
    Case dbsmPairsN14N15: strNewSearchMode = AUTO_SEARCH_PAIRS_N14N15                           ' No longer supported (july 2004)
    Case dbsmPairsN14N15ConglomerateMass: strNewSearchMode = AUTO_SEARCH_PAIRS_N14N15_CONGLOMERATEMASS
    Case dbsmPairsICAT: strNewSearchMode = AUTO_SEARCH_PAIRS_ICAT
    Case dbsmPairsPEO: strNewSearchMode = AUTO_SEARCH_PAIRS_PEO
    Case Else
        ' This shouldn't happen
        Debug.Assert False
        Exit Sub
    End Select
    
    If eDBSearchMode >= DB_SEARCH_MODE_PAIR_MODE_START_INDEX Then
        If cboPairsIdentificationMode.ListIndex = 0 Then
           MsgBox "Please select a pairs identification mode on the Pairs tab before selecting a pairs-based database search mode.", vbInformation + vbOKOnly, "Pairs Mode Not Defined"
           Exit Sub
        Else
            If eDBSearchMode = dbsmPairsICAT Or eDBSearchMode = dbsmPairsPEO Then
                If cboPairsIdentificationMode.ListIndex <> pimLabel Then
                End If
            End If
        End If
    End If


    lstDBSearchModes.AddItem strNewSearchMode

    With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions
        intNewIndex = .AutoAnalysisSearchModeCount
        .AutoAnalysisSearchModeCount = .AutoAnalysisSearchModeCount + 1
        
        With .AutoAnalysisSearchMode(intNewIndex)
            .SearchMode = strNewSearchMode
        
            If mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount > 1 Then
                .AlternateOutputFolderPath = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intNewIndex - 1).AlternateOutputFolderPath
            End If
            
            If APP_BUILD_DISABLE_MTS Then
                .InternalStdSearchMode = issmFindOnlyMassTags
            Else
                .InternalStdSearchMode = issmFindWithMassTags
            End If
            
            .DBSearchMinimumHighNormalizedScore = 0
            .DBSearchMinimumHighDiscriminantScore = 0
            .DBSearchMinimumPeptideProphetProbability = 0
            
            .ExportResultsToDatabase = False
            .ExportUMCMembers = False
            .WriteResultsToTextFile = True
        End With
    End With
    
    lstDBSearchModes.ListIndex = lstDBSearchModes.ListCount - 1
    
    UpdateDynamicControls

    Exit Sub

AddNewDBSearchModeErrorHandler:
    Debug.Print "Error in AddNewDBSearchMode: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmEditAnalysisSettings->AddNewDBSearchMode"
    Resume Next
    
End Sub

Private Sub ApplyDBSettingsToGelInMemory(udtDBSettings As udtDBSettingsType)
    
    Dim udtExistingAnalysisInfo As udtGelAnalysisInfoType
    
    Me.MousePointer = vbHourglass
    DoEvents
    
    If udtDBSettings.IsDeleted Then
        ' Leave GelAnalysis(mActiveGelIndex) unchanged
    Else
        If GelAnalysis(mActiveGelIndex) Is Nothing Then
            Set GelAnalysis(mActiveGelIndex) = New FTICRAnalysis
            udtExistingAnalysisInfo.ValidAnalysisDataPresent = False
        Else
            FillGelAnalysisInfo udtExistingAnalysisInfo, GelAnalysis(mActiveGelIndex)
        End If
        
        FillGelAnalysisObject GelAnalysis(mActiveGelIndex), udtDBSettings.AnalysisInfo
        
        If udtExistingAnalysisInfo.ValidAnalysisDataPresent Then
        
            ' Update GelAnalysis() with the settings in udtAnalysisInfo
            ' However, do not update .MTDB or the DBStuff() collection since we want the settings in
            '  udtDBSettings to take precedence
            FillGelAnalysisObject GelAnalysis(mActiveGelIndex), udtExistingAnalysisInfo, False, False
            
        End If
        
        IniFileUpdateRecentDatabaseConnectionInfo udtDBSettings
        
        AddToAnalysisHistory mActiveGelIndex, "Database connection defined: " & udtDBSettings.DatabaseName & "; " & udtDBSettings.SelectedMassTagCount & " MT tags"
    
        GelStatus(mActiveGelIndex).Dirty = True
    End If

    Me.MousePointer = vbDefault

End Sub

Private Sub ApplyToSelectedGel()

    Dim eResponse As VbMsgBoxResult
    
    If cboGelFileInMemoryToUse.ListCount <= 0 Then
        MsgBox "No gels are present in memory; unable to continue.", vbInformation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If mActiveGelIndex < 1 Or mActiveGelIndex > UBound(GelBody()) Then
        MsgBox "A valid gel is not selected in the 'Gel file (in memory) to read or update' list."
        Exit Sub
    End If

    eResponse = MsgBox("Are you sure you want to apply the settings to the selected gel?", vbQuestion + vbYesNoCancel, "Apply Settings")
    If eResponse <> vbYes Then Exit Sub
    
    With mCurrentSettings
        GelData(mActiveGelIndex).Preferences = .Prefs
        GelData(mActiveGelIndex).PathtoDatabase = .PrefsExpanded.LegacyAMTDBPath
        
        glbPreferencesExpanded = .PrefsExpanded
        GelSearchDef(mActiveGelIndex).UMCDef = .UMCDef
        GelSearchDef(mActiveGelIndex).UMCIonNetDef = .UMCIonNetDef
        GelUMCNETAdjDef(mActiveGelIndex) = .UMCNetAdjDef
        GelSearchDef(mActiveGelIndex).AMTSearchOnUMCs = .AMTDef
        GelSearchDef(mActiveGelIndex).AMTSearchOnIons = .AMTDef
        GelSearchDef(mActiveGelIndex).AMTSearchOnPairs = .AMTDef
        
        GelP_D_L(mActiveGelIndex).SearchDef = .PrefsExpanded.PairSearchOptions.SearchDef
        
        ApplyDBSettingsToGelInMemory .DBSettings
        
        ApplyAutoAnalysisFilter .PrefsExpanded.AutoAnalysisFilterPrefs, mActiveGelIndex, False
    
        With .PrefsExpanded.AutoAnalysisOptions
            If .AutoAnalysisSearchModeCount > 0 Then
                GelSearchDef(mActiveGelIndex).AMTSearchMassMods = .AutoAnalysisSearchMode(0).MassMods
            End If
        End With
    
    End With
    
    ' Filter the gel in memory with the defined filters
    
    ' Assign udtWorkingParams.GelIndex to frmFilter.Tag, then call .InitializeControls
    frmFilter.Tag = mActiveGelIndex
    frmFilter.InitializeControls True
    
End Sub

Private Sub AutoShowHideWarpingTab()
    If RobustNETWarpEnabled() Then
        tbsNETOptions.Tab = 2
    Else
        tbsNETOptions.Tab = 0
    End If
End Sub

Private Sub DisplayCurrentAnalysisSettings()
    Dim intIndex As Integer
    Dim intDBSearchModeIndexSaved As Integer
    
On Error GoTo DisplayCurrentAnalysisSettingsErrorHandler

    mUpdatingControls = True
    
    With mCurrentSettings
        With .PrefsExpanded
            chkExtendedFileSaveModePreferred.Enabled = True
            SetCheckBox chkExtendedFileSaveModePreferred, .ExtendedFileSaveModePreferred
            
            ' Note: Skipping .AutoAdjSize
            ' Note: Skipping .AutoSizeMultiplier
            ' Note: Skipping .UMCDrawType
            
            ' Note: Skipping .UsePEKBasedERValues
            ' Note: Skipping .UseMassTagsWithNullMass
            
            SetCheckBox chkUseMassTagsWithNullNET, .UseMassTagsWithNullNET
            
            ' Note: Skipping .UseSTAC
            ' Note: Skipping .STACUsesPriorProbability
            ' Note: Skipping .STACAlignsDriftTime
            ' Note: Skipping .KeepTempSTACFiles
            ' Note: Skipping .IReportAutoAddMonoPlus4AndMinus4Data
            
            SetCheckBox chkUseUMCConglomerateNET, .UseUMCConglomerateNET
            ' Note: Skipping .NetAdjustmentUsesN15AMTMasses
            
            ' Note: Skipping .NetAdjustmentMinHighNormalizedScore
            ' Note: Skipping .NetAdjustmentMinHighDiscriminantScore
            
            txtLegacyAMTDatabasePath.Text = .LegacyAMTDBPath
        End With
        
        With .PrefsExpanded.UMCAutoRefineOptions
            SetCheckBox chkAutoRefineCheckbox(uarRemoveShort), .UMCAutoRefineRemoveCountLow
            SetCheckBox chkAutoRefineCheckbox(uarRemoveLong), .UMCAutoRefineRemoveCountHigh
            SetCheckBox chkAutoRefineCheckbox(uarRemoveMaxLengthPctAllScans), .UMCAutoRefineRemoveMaxLengthPctAllScans
            
            txtAutoRefineTextbox(uarRemoveShort) = Trim(.UMCAutoRefineMinLength)
            txtAutoRefineTextbox(uarRemoveLong) = Trim(.UMCAutoRefineMaxLength)
            txtAutoRefineTextbox(uarRemoveMaxLengthPctAllScans) = Trim(.UMCAutoRefineMaxLengthPctAllScans)
            
            txtAutoRefineTextbox(uarPercentMaxAbuToUseToGaugeLength) = Trim(.UMCAutoRefinePercentMaxAbuToUseForLength)
            
            SetCheckBox chkAutoRefineCheckbox(uarTestLengthUsingScanRange), .TestLengthUsingScanRange
            txtAutoRefineTextbox(uarMinimumMemberCount) = .MinMemberCountWhenUsingScanRange
            
            SetCheckBox chkAutoRefineCheckbox(uarRemoveLowIntensity), .UMCAutoRefineRemoveAbundanceLow
            SetCheckBox chkAutoRefineCheckbox(uarRemoveHighIntensity), .UMCAutoRefineRemoveAbundanceHigh
            txtAutoRefineTextbox(uarRemoveLowIntensity) = Trim(.UMCAutoRefinePctLowAbundance)
            txtAutoRefineTextbox(uarRemoveHighIntensity) = Trim(.UMCAutoRefinePctHighAbundance)
        
            SetCheckBox chkSplitUMCsByExaminingAbundance, .SplitUMCsByAbundance
            With .SplitUMCOptions
                txtSplitUMCsMaximumPeakCount = Trim(.MaximumPeakCountToSplitUMC)
                txtSplitUMCsMinimumDifferenceInAvgPpmMass = Trim(.MinimumDifferenceInAveragePpmMassToSplit)
                txtSplitUMCsStdDevMultiplierForSplitting = Trim(.StdDevMultiplierForSplitting)
                txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax = Trim(.PeakDetectIntensityThresholdPercentageOfMaximum)
                txtSplitUMCsPeakPickingMinimumWidth = Trim(.PeakWidthPointsMinimum)
                SetComboBox cboSplitUMCsScanGapBehavior, .ScanGapBehavior, "Split UMC Scan Gap Behavior"
            End With
        
        End With
        
        With .PrefsExpanded.UMCAdvancedStatsOptions
            txtClassAbuTopXMinMaxAbu(0) = .ClassAbuTopXMinAbu
            txtClassAbuTopXMinMaxAbu(1) = .ClassAbuTopXMaxAbu
            txtClassAbuTopXMinMembers = .ClassAbuTopXMinMembers
            
            txtClassMassTopXMinMaxAbu(0) = .ClassMassTopXMinAbu
            txtClassMassTopXMinMaxAbu(1) = .ClassMassTopXMaxAbu
            txtClassMassTopXMinMembers = .ClassMassTopXMinMembers
        End With
        
        With .PrefsExpanded.ErrorPlottingOptions
                txtErrorGraphicMassRangePPM = Trim(.MassRangePPM)
                txtErrorGraphicMassBinSizePPM = Trim(.MassBinSizePPM)
                txtErrorGraphicGANETRange = Trim(.GANETRange)
                txtErrorGraphicGANETBinSize = Trim(.GANETBinSize)
                
                ' Note: Not setting .DriftTimeBinSize here
                
                ' Note: Not setting .Graph2DOptions here
                ' Note: Not setting .Graph3DOptions here
        End With
        
        With .PrefsExpanded.PairSearchOptions
            Select Case LCase(.PairSearchMode)
            Case LCase(AUTO_FIND_PAIRS_NONE): SetComboBox cboPairsIdentificationMode, pimNone, "Pairs identification mode"
            Case LCase(AUTO_FIND_PAIRS_DELTA): SetComboBox cboPairsIdentificationMode, pimDelta, "Pairs identification mode"
            Case LCase(AUTO_FIND_PAIRS_LABEL): SetComboBox cboPairsIdentificationMode, pimLabel, "Pairs identification mode"
            End Select
            
            With .SearchDef
                txtPairsDelta = .DeltaMass
                txtPairsDeltaTol = .DeltaMassTolerance
                If .DeltaMassTolType = gltPPM Then
                    optPairTolType(0).Value = True
                Else
                    optPairTolType(1).Value = True
                End If
                
                SetCheckBox chkPairsAutoMinMaxDelta, .AutoCalculateDeltaMinMaxCount
                txtPairsMinMaxDelta(0) = .DeltaCountMin
                txtPairsMinMaxDelta(1) = .DeltaCountMax
                txtPairsMinMaxDelta(2) = .DeltaStepSize
                
                txtPairsLabel = .LightLabelMass
                txtPairsHeavyLabelDelta = .HeavyLightMassDifference
                txtPairsMinMaxLbl(0) = .LabelCountMin
                txtPairsMinMaxLbl(1) = .LabelCountMax
                txtPairsMaxLblDiff = .MaxDifferenceInNumberOfLightHeavyLabels
            
                SetCheckBox chkPairsRequireOverlapAtEdge, .RequireUMCOverlap
                SetCheckBox chkPairsRequireOverlapAtApex, .RequireUMCOverlapAtApex
                
                txtPairsScanTolEdge = .ScanTolerance
                txtPairsScanTolApex = .ScanToleranceAtApex
                
                txtPairsERMinMax(0) = .ERInclusionMin
                txtPairsERMinMax(1) = .ERInclusionMax
                
                SetCheckBox chkPairEROptions(psoRequireMatchingChargeStatesForPairMembers), .RequireMatchingChargeStatesForPairMembers
                SetCheckBox chkPairEROptions(psoUseIdenticalChargesForER), .UseIdenticalChargesForER
                SetCheckBox chkPairEROptions(psoAverageERsAllChargeStates), .AverageERsAllChargeStates
                SetCheckBox chkPairEROptions(psoComputeERScanByScan), .ComputeERScanByScan
                SetCheckBox chkPairEROptions(psoIReportEREnable), .IReportEROptions.Enabled
                SetCheckBox chkPairEROptions(psoRemoveOutlierERs), .RemoveOutlierERs
                SetCheckBox chkPairEROptions(psoRemoveOutlierERsIterate), .RemoveOutlierERsIterate
                
                If .RemoveOutlierERsMinimumDataPointCount < 2 Then .RemoveOutlierERsMinimumDataPointCount = 2
                txtRemoveOutlierERsMinimumDataPointCount = .RemoveOutlierERsMinimumDataPointCount
                
                SetComboBox cboAverageERsWeightingMode, .AverageERsWeightingMode, "Average ERs Weighting Mode"
                
                optERCalc(.ERCalcType).Value = True
            End With
            
            SetCheckBox chkPairsExcludeOutOfERRange, .AutoExcludeOutOfERRange
            SetCheckBox chkPairsExcludeAmbiguous, .AutoExcludeAmbiguous
            SetCheckBox chkPairsExcludeAmbiguousKeepMostConfident, .KeepMostConfidentAmbiguous
            
            SetCheckBox chkRemovePairMemberHitsAfterDBSearch, .AutoAnalysisRemovePairMemberHitsAfterDBSearch
            If .AutoAnalysisRemovePairMemberHitsRemoveHeavy Then
                optRemovePairMemberHitsRemoveHeavy(rpmhRemoveHeavy).Value = True
            Else
                optRemovePairMemberHitsRemoveHeavy(rpmhRemoveLight).Value = True
            End If
            
            SetCheckBox chkPairsSaveTextFile, .AutoAnalysisSavePairsToTextFile
            SetCheckBox chkPairsSaveStatisticsTextFile, .AutoAnalysisSavePairsStatisticsToTextFile
            
            SetComboBox cboPairsUMCsToUseForNETAdjustment, .NETAdjustmentPairedSearchUMCSelection, "Paired UMCs to use for NET Adjustment"
            
        End With
        
        With .PrefsExpanded.RefineMSDataOptions
            txtToleranceRefinementMinimumPeakHeight = Trim(.MinimumPeakHeight)
            txtToleranceRefinementPercentageOfMaxForWidth = Trim(.PercentageOfMaxForFindingWidth)
            txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks = Trim(.MinimumSignalToNoiseRatioForLowAbundancePeaks)
            
            txtRefineMassCalibrationMaximumShift = Trim(.MassCalibrationMaximumShift)
            If .MassCalibrationTolType = gltPPM Then
                optRefineMassCalibrationMassType(0).Value = True
            Else
                optRefineMassCalibrationMassType(1).Value = True
            End If
            txtRefineDBSearchMassToleranceMinMax(0) = Trim(.MassToleranceMinimum)
            txtRefineDBSearchMassToleranceMinMax(1) = Trim(.MassToleranceMaximum)
            txtRefineDBSearchMassToleranceAdjustmentMultiplier = Trim(.MassToleranceAdjustmentMultiplier)
            
            txtRefineDBSearchNETToleranceMinMax(0) = Trim(.NETToleranceMinimum)
            txtRefineDBSearchNETToleranceMinMax(1) = Trim(.NETToleranceMaximum)
            txtRefineDBSearchNETToleranceAdjustmentMultiplier = Trim(.NETToleranceAdjustmentMultiplier)
            
            SetComboBox cboToleranceRefinementMethod, .ToleranceRefinementMethod, "Tolerance Refinement Method"
            SetCheckBox chkRefineDBSearchTolUseMinMaxIfOutOfRange, .UseMinMaxIfOutOfRange
            
            txtEMRefineMassErrorPeakToleranceEstimatePPM.Text = Trim(.EMMassErrorPeakToleranceEstimatePPM)
            txtEMRefineNETErrorPeakToleranceEstimate.Text = Trim(.EMNETErrorPeakToleranceEstimate)
            txtEMRefinePercentOfDataToExclude.Text = Trim(.EMPercentOfDataToExclude)
            
            SetCheckBox chkEMRefineMassTolForceUseAllDataPointErrors, .EMMassTolRefineForceUseSingleDataPointErrors
            SetCheckBox chkEMRefineNETTolForceUseAllDataPointErrors, .EMNETTolRefineForceUseSingleDataPointErrors
            
            SetCheckBox chkRefineDBSearchIncludeInternalStds, .IncludeInternalStdMatches
            SetCheckBox chkUseUMCClassStats, .UseUMCClassStats
            txtToleranceRefinementFilter(trfMinimumSLiC) = Trim(.MinimumSLiC)
            txtToleranceRefinementFilter(trfMaximumAbundance) = Trim(.MaximumAbundance)
        End With
        
        With .PrefsExpanded.AutoAnalysisOptions
            SetCheckBox chkDoNotSaveOrExport, .DoNotSaveOrExport
            
            SetCheckBox chkSkipFindUMCs, .SkipFindUMCs
            SetCheckBox chkSkipGANETSlopeAndInterceptComputation, .SkipGANETSlopeAndInterceptComputation
            
            txtDBConnectionRetryAttemptMax = Trim(.DBConnectionRetryAttemptMax)
            txtDBConnectionTimeoutSeconds = Trim(.DBConnectionTimeoutSeconds)
            SetCheckBox chkExportResultsFileUsesJobNumber, .ExportResultsFileUsesJobNumberInsteadOfDataSetName
            
            SetCheckBox chkSaveGelFile, .SaveGelFile
            SetCheckBox chkSaveGelFileOnError, .SaveGelFileOnError
            
            ' Need to enable and set cboSavePictureGraphicFileType before setting chkSavePictureGraphic
            cboSavePictureGraphicFileType.Enabled = True
            SetComboBox cboSavePictureGraphicFileType, .SavePictureGraphicFileType - 1, "Save Picture graphic file type"
            SetCheckBox chkSavePictureGraphic, .SavePictureGraphic
            
            txtPictureGraphicSizeWidthPixels = Trim(.SavePictureWidthPixels)
            txtPictureGraphicSizeHeightPixels = Trim(.SavePictureHeightPixels)
            
            ' Note: Skipping .SaveInternalStdHitsAndData
            
            SetCheckBox chkSaveErrorGraphicMass, .SaveErrorGraphicMass
            SetCheckBox chkSaveErrorGraphicGANET, .SaveErrorGraphicGANET
            SetCheckBox chkSaveErrorGraphic3D, .SaveErrorGraphic3D
            
            ' Note: Skipping .SaveErrorGraphicDriftTime
            
            ' Note: Skipping .SaveErrorGraphicFileType
            ' Note: Skipping .SaveErrorGraphSizeWidthPixels
            ' Note: Skipping .SaveErrorGraphSizeHeightPixels
            
            ' Note: txtNetAdjInitialNETTol is updated below
            
            ' Note: Skipping .NETAdjustmentMaxIterationCount
            ' Note: Skipping .NETAdjustmentMinIDCount
            ' Note: Skipping .NETAdjustmentMinIDCountAbsoluteMinimum
            ' Note: Skipping .NETAdjustmentChangeThresholdStopValue
            
            ' Note: Skipping .NETAdjustmentAutoIncrementUMCTopAbuPct
            ' Note: Skipping .NETAdjustmentUMCTopAbuPctIncrement
            
            
            ' Note: Skipping .NETSlopeExpectedMinimum, .NETSlopeExpectedMaximum, .NETInterceptExpectedMinimum, and .NETInterceptExpectedMaximum
            
            If cboUMCSearchMode.ListCount >= 3 Then
                If .UMCSearchMode = AUTO_ANALYSIS_UMCListType2002 Then
                    ' Obsolete search mode (July 2004)
                    SetComboBox cboUMCSearchMode, usmUMC2002, "UMC Search Mode"
                ElseIf .UMCSearchMode = AUTO_ANALYSIS_UMC2003 Then
                    SetComboBox cboUMCSearchMode, usmUMC2003, "UMC Search Mode"
                Else
                    SetComboBox cboUMCSearchMode, usmUMCIonNet, "UMC Search Mode"
                End If
            End If
            
            SetCheckBox chkUMCShrinkingBoxWeightAverageMassByIntensity, .UMCShrinkingBoxWeightAverageMassByIntensity
            
            txtOutputFileSeparationCharacter = .OutputFileSeparationCharacter
            txtPEKFileExtensionPreferenceOrder = .PEKFileExtensionPreferenceOrder
            SetCheckBox chkDBSearchWriteIDResultsByIonAfterAutoSearches, .WriteIDResultsByIonToTextFileAfterAutoSearches
            SetCheckBox chkSaveUMCStatisticsToTextFile, .SaveUMCStatisticsToTextFile
            SetCheckBox chkDBSearchIncludeORFNameInTextFileOutput, .IncludeORFNameInTextFileOutput
            SetCheckBox chkSetIsConfirmedForDBSearchMatches, .SetIsConfirmedForDBSearchMatches
            
            SetCheckBox chkAddQuantitationDescriptionEntry, .AddQuantitationDescriptionEntry
            SetCheckBox chkExportUMCsWithNoMatches, .ExportUMCsWithNoMatches
            
            SetComboBox cboSearchRegionShape(1), .DBSearchRegionShape, "DB Search region Shape"
            
            SetCheckBox chkUseLegacyDBForMTs, .UseLegacyDBForMTs
            
            With .AutoToleranceRefinement
                If .DBSearchTolType = gltPPM Then
                    optAutoToleranceRefinementDBSearchTolType(0).Value = True
                Else
                    optAutoToleranceRefinementDBSearchTolType(1).Value = True
                End If
                
                txtAutoToleranceRefinementDBSearchMWTol = .DBSearchMWTol
                txtAutoToleranceRefinementDBSearchNETTol = .DBSearchNETTol
                
                SetComboBox cboSearchRegionShape(0), .DBSearchRegionShape, "DB Search region Shape for Tolerance Refinement"
                
                txtToleranceRefinementFilter(trfMinimumHighNormalizedScore) = .DBSearchMinimumHighNormalizedScore
                txtToleranceRefinementFilter(trfMinimumDiscriminant) = .DBSearchMinimumHighDiscriminantScore
                txtToleranceRefinementFilter(trfMinimumPeptideProphet) = .DBSearchMinimumPeptideProphetProbability
                
                txtRefineMassCalibrationOverridePPM = .RefineMassCalibrationOverridePPM
                
                SetCheckBox chkRefineMassCalibration, .RefineMassCalibration
                SetCheckBox chkRefineDBSearchMassTolerance, .RefineDBSearchMassTolerance
                SetCheckBox chkRefineDBSearchNETTolerance, .RefineDBSearchNETTolerance
                
                ' Ignore: .UseRefinementWhenUsingSTAC
                ' Ignore: .RefinedTolMassMultiplierWhenUsingSTAC
                ' Ignore: .RefinedTolNETMultiplierWhenUsingSTAC
            End With
            
            intDBSearchModeIndexSaved = lstDBSearchModes.ListIndex
            lstDBSearchModes.Clear
            For intIndex = 0 To .AutoAnalysisSearchModeCount - 1
                lstDBSearchModes.AddItem .AutoAnalysisSearchMode(intIndex).SearchMode
            Next intIndex
            
            If lstDBSearchModes.ListCount > 0 Then
                If intDBSearchModeIndexSaved >= 0 And intDBSearchModeIndexSaved < lstDBSearchModes.ListCount Then
                    lstDBSearchModes.ListIndex = intDBSearchModeIndexSaved
                Else
                    lstDBSearchModes.ListIndex = 0
                End If
                
                SetComboBox cboDBSearchMode, LookupDBSearchModeIndex(.AutoAnalysisSearchMode(0).SearchMode), "DB Search Mode"
            End If
            
        End With
        
        With .PrefsExpanded.AutoAnalysisFilterPrefs
            SetCheckBox chkDupElimination, .ExcludeDuplicates
            txtDupTolerance = Trim(.ExcludeDuplicatesTolerance)
            
            SetCheckBox chkExcludeIsoByFit, .ExcludeIsoByFit
            txtExcludeIsoByFit = .ExcludeIsoByFitMaxVal
            
            SetCheckBox chkSecGuessElimination(0), .ExcludeIsoSecondGuess
            SetCheckBox chkSecGuessElimination(1), .ExcludeIsoLessLikelyGuess
            
            ' Note: Skipping .ExcludeCSByStdDevMaxVal
            ' Note: Skipping .ExcludeCSByStdDev
            
            SetCheckBox chkIsoAbuRange, .RestrictIsoByAbundance
            txtIsoMinMaxAbu(0) = Trim(.RestrictIsoAbundanceMin)
            txtIsoMinMaxAbu(1) = Trim(.RestrictIsoAbundanceMax)
            
            SetCheckBox chkIsoMWRange, .RestrictIsoByMass
            txtIsoMinMaxMW(0) = Trim(.RestrictIsoMassMin)
            txtIsoMinMaxMW(1) = Trim(.RestrictIsoMassMax)
            
            SetCheckBox chkIsoMZRange, .RestrictIsoByMZ
            txtIsoMinMaxMZ(0) = Trim(.RestrictIsoMZMin)
            txtIsoMinMaxMZ(1) = Trim(.RestrictIsoMZMax)
            
            SetCheckBox chkIsoUseCSRange, .RestrictIsoByChargeState
            txtIsoMinMaxCS(0) = Trim(.RestrictIsoChargeStateMin)
            txtIsoMinMaxCS(1) = Trim(.RestrictIsoChargeStateMax)
            
            SetCheckBox chkCSAbuRange, .RestrictCSByAbundance
            txtCSMinMaxAbu(0) = Trim(.RestrictCSAbundanceMin)
            txtCSMinMaxAbu(1) = Trim(.RestrictCSAbundanceMax)
            
            SetCheckBox chkCSMWRange, .RestrictCSByMass
            txtCSMinMaxMW(0) = Trim(.RestrictCSMassMin)
            txtCSMinMaxMW(1) = Trim(.RestrictCSMassMax)
            
            SetCheckBox chkRestrictScanRange, .RestrictScanRange
            txtRestrictScanRangeMinMax(0) = Trim(.RestrictScanRangeMin)
            txtRestrictScanRangeMinMax(1) = Trim(.RestrictScanRangeMax)
            
            SetCheckBox chkRestrictGANETRange, .RestrictGANETRange
            txtRestrictGANETRangeMinMax(0) = Trim(.RestrictGANETRangeMin)
            txtRestrictGANETRangeMinMax(1) = Trim(.RestrictGANETRangeMax)
            
            If .RestrictToEvenScanNumbersOnly Then
                SetComboBox cboEvenOddScanNumber, eosEvenScans, "Even Odd Scan Number Load Mode"
            ElseIf .RestrictToOddScanNumbersOnly Then
                SetComboBox cboEvenOddScanNumber, eosOddScans, "Even Odd Scan Number Load Mode"
            Else
                SetComboBox cboEvenOddScanNumber, eosAllScans, "Even Odd Scan Number Load Mode"
            End If
            
            txtMaximumDataCountToLoad = Trim(.MaximumDataCountToLoad)
            SetCheckBox chkMaximumDataCountEnabled, .MaximumDataCountEnabled
            
            ' Note: Skipping .FilterLCMSFeatures
            ' Note: Skipping .LCMSFeatureAbuMin
            ' Note: Skipping .LCMSFeatureScanCountMin
            ' Note: Skipping .IMSConformerScoreMin

        End With
        
        ' .PrefsExpanded.AutoAnalysisDBInfo
        DisplayCurrentDBSettings
        
        With .Prefs
            .DupTolerance = mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeDuplicatesTolerance
            .IsoDataFit = mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeIsoByFitMaxVal
            optIsoDataFrom(.IsoDataField - MW_FIELD_OFFSET).Value = True
        End With
        
        With .UMCDef
            
            If .UMCType < cmbCountType.ListCount Then
                SetComboBox cmbCountType, .UMCType, "UMC creation mode"
            End If
            
            'Note: .DefScope is ignored here
            optUMCSearchMWField(.MWField - MW_FIELD_OFFSET).Value = True
            Select Case .TolType
            Case gltPPM
                optUMCSearchTolType(0).Value = True
            Case gltABS
                optUMCSearchTolType(1).Value = True
            Case Else
                Debug.Assert False
            End Select
            
            txtUMCSearchMassTol = Trim(.Tol)
            SetCheckBox chkAllowSharing, .UMCSharing
            SetCheckBox chkUMCDefRequireIdenticalCharge, .UMCUniCS
            SetComboBox cmbUMCAbu, .ClassAbu, "UMC Class Abundance"
            SetComboBox cmbUMCMW, .ClassMW, "UMC Class Mass"
            SetComboBox cboChargeStateAbuType, .ChargeStateStatsRepType, "UMC Charge State Stats Rep Type"
            SetCheckBox chkUseMostAbuChargeStateStatsForClassStats, .UMCClassStatsUseStatsFromMostAbuChargeState
            
            txtHoleNum = Trim(.GapMaxCnt)
            txtHoleSize(0) = Trim(.GapMaxSize)
            txtHoleSize(1) = Trim(.GapMaxSize)
            txtHolePct = Trim(Round(CLng(.GapMaxPct * 100), 2))    ' UMCListType2002 only
            
            ' Note: .UMCNETType is not used
            ' Note: .UMCMaxAbuPctBf is not used
            ' Note: .UMCMaxAbuPctAf is not used
            ' Note: .UMCMaxAbuEtPctBf is not used
            ' Note: .UMCMaxAbuEtPctAf is not used
            
            ' Note: .UMCMinCnt is the same as .PrefsExpanded.AutoRefineMinLength; do not define here
            ' Note: .UMCMaxCnt is the same as .PrefsExpanded.AutoRefineMaxLength; do not define here
            
            SetCheckBox chkInterpolateMissingIons, .InterpolateGaps
            txtInterpolateMaxGapSize = Trim(.InterpolateMaxGapSize)
            ' Note: .InterpolationType is not used
        
        End With
        
        With .UMCIonNetDef
            txtRejectLongConnections.Text = .TooDistant
            txtNETType.Text = .NETType
            SetComboBox cmbMetricType, .MetricType, "UMC Metric Type"
            For intIndex = 0 To .NetDim - 1
                With .MetricData(intIndex)
                    SetCheckBox chkUse(intIndex), .Use
                    SetComboBox cmbData(intIndex), .DataType, "UMC Metric Data Type"

                    ' Note: cmbConstraintUnits() is updated inside UpdateDynamicControls
                    SetComboBox cmbConstraint(intIndex), .ConstraintType, "UMC Metric Constraint"
                    txtWeightingFactor(intIndex) = .WeightFactor
                    txtConstraint(intIndex) = .ConstraintValue
                    SetComboBox cmbConstraintUnits(intIndex), .ConstraintUnits, "UMC Constraint Units"
                End With
            Next intIndex
        End With
        
        ' Additional UMCIonNet options
        With .PrefsExpanded.UMCIonNetOptions
            SetComboBox cmbUMCRepresentative, .UMCRepresentative, "UMC Representative"
            SetCheckBox chkUMCIonNetMakeSingleMemberClasses, .MakeSingleMemberClasses
            ' Note: Skipping: .LimitToSingleChargeState
        End With
        
        With .UMCNetAdjDef
        
            ' Note: Skipping: .MinUMCCount
            ' Note: Skipping: .MinScanRange
            ' Note: Skipping: .MaxScanPct
            ' Note: Skipping: .TopAbuPct
        
            ' Note: Skipping: .TopAbuPct
            
            ' Note: Skipping: .PeakCSSelection(intIndex)        (ranges from 0 to 7)
            
            optNETAdjTolType(.MWTolType).Value = True
            txtNETAdjMWTol = Trim(.MWTol)
        
            
            ' Note: .NETTolIterative is ignored here
            ' Note: .NETorRT is ignored here
            ' Note: .UseNET is ignored here
            
            ' Note: Skipping: .UseMultiIDMaxNETDist
            ' Note: Skipping: .MultiIDMaxNETDist
            
            ' Note: Skipping: .EliminateBadNET
            ' Note: Skipping: .MaxIDToUse
            
            ' Note: .IterationStopType is ignored here; During Auto Analysis,
            '       .IterationStopType is forced to be 4 (ITERATION_STOP_CHANGE)
            ' Note: .IterationStopValue is ignored here
            
            ' Note: Skipping: .IterationUseMWDec
            ' Note: Skipping: .IterationMWDec
            ' Note: Skipping: .IterationUseNETdec
            ' Note: Skipping: .IterationNETDec
            
            ' Note: .IterationAcceptLast is ignored here
            
            ' Note: Skipping: .InitialSlope)
            ' Note: Skipping: .InitialIntercept)
            ''UpdateNetAdjRangeStats mCurrentSettings.UMCNetAdjDef, lblNetAdjInitialNETStats

            ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
            '' SetCheckBox chkNetAdjUseLockers, .UseNetAdjLockers
            '' SetCheckBox chkNetAdjUseOldIfFailure, .UseOldNetAdjIfFailure
            '' txtNETAlignmentOption(naoNetAdjMinLockerMatchCount) = Trim(.NetAdjLockerMinimumMatchCount)
            
            SetCheckBox chkRobustNETEnabled, .UseRobustNETAdjustment
            
            If APP_BUILD_DISABLE_LCMSWARP Then
                If .RobustNETAdjustmentMode <> UMCRobustNETModeConstants.UMCRobustNETIterative Then
                    .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETIterative
                End If
            End If
            
            SetComboBox cboRobustNETAdjustmentMode, .RobustNETAdjustmentMode, "Robust NET Adjstument Mode"
            
            If RobustNETWarpEnabled() Then
                txtNetAdjInitialNETTol = Trim(.MSWarpOptions.NETTol)
            Else
                txtNetAdjInitialNETTol = Trim(mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol)
            End If
            
            ' Note: Skipping: .RobustNETSlopeIncreaseMode
            
            ' Note: Skipping: .RobustNETSlopeStart
            ' Note: Skipping: .RobustNETSlopeEnd
            ' Note: Skipping: .RobustNETSlopeIncrement
            
            ' Note: Skipping: .RobustNETInterceptStart
            ' Note: Skipping: .RobustNETInterceptEnd
            ' Note: Skipping: .RobustNETInterceptIncrement
            
            ' Note: Skipping: .RobustNETMassShiftPPMStart
            ' Note: Skipping: .RobustNETMassShiftPPMEnd
            ' Note: Skipping: RobustNETMassShiftPPMIncrement
            
            With .MSWarpOptions
            
                SetComboBox cboMSWarpMassCalibrationType, .MassCalibrationType, "MS Warp Mass Calibration Type"
                
                txtRobustNETOption(rnoWarpNumberOfSections).Text = Trim(.NumberOfSections)
                txtRobustNETOption(rnoWarpMaxDistortion).Text = Trim(.MaxDistortion)
                txtRobustNETOption(rnoWarpContractionFactor).Text = Trim(.ContractionFactor)
                txtRobustNETOption(rnoWarpMinimumPMTTagObsCount).Text = Trim(.MinimumPMTTagObsCount)
                txtRobustNETOption(rnoWarpMatchPromiscuity).Text = Trim(.MatchPromiscuity)
                
                txtRobustNETOption(rnoWarpMassWindow).Text = Trim(.MassWindowPPM)
                txtRobustNETOption(rnoWarpMassSplineOrder).Text = Trim(.MassSplineOrder)
                txtRobustNETOption(rnoWarpNumXAxisSlices).Text = Trim(.MassNumXSlices)
                txtRobustNETOption(rnoWarpNumMassDeltaBins).Text = Trim(.MassNumMassDeltaBins)
                txtRobustNETOption(rnoWarpMassMaxJump).Text = Trim(.MassMaxJump)
                
                ValidateNETWarpMassMaxJump
            End With

            AutoShowHideWarpingTab
        End With
        
        With .AMTDef
            ' Note: .SearchScope is ignored here
            ' Note: SearchFlag is ignored here
            optDBSearchMWField(.MWField - MW_FIELD_OFFSET).Value = True
            If .TolType = gltPPM Then
                optDBSearchTolType(0).Value = True
            Else
                optDBSearchTolType(1).Value = True
            End If
            ' Note: NETorRT is ignored here
            ' Note: Formula is ignored here
            txtDBSearchMWTol = .MWTol
            txtDBSearchNETTolerance = .NETTol
            
            ' Note: .UseDriftTime is ignored here
            ' Note: .DriftTimeTol is ignored here
            
            ' Note: MassTag is ignored here
            ' Note: MaxMassTags is ignored here
            ' Note: SkipReferenced is ignored here
            ' Note: SaveNCnt is ignored here
        End With
        
    End With
    
    mUpdatingControls = False
    
    UpdateDynamicControls
    DisplayDataValidationWarnings True

    Exit Sub

DisplayCurrentAnalysisSettingsErrorHandler:
    Debug.Print "Error occured in DisplayCurrentAnalysisSettings: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmEditAnalysisSettings->DisplayCurrentAnalysisSettings"
    Resume Next
    
End Sub

Private Sub DisplayCurrentDBSettings()
    Dim blnShowDBSchema1Controls As Boolean

    With mCurrentSettings.DBSettings
        If Not .IsDeleted Then
            txtCurrentDBName = .DatabaseName
            lblCurrentDBMassTagCount = .SelectedMassTagCount
            
            If .DBSchemaVersion >= 2 Then
                lblCurrentDBInternalStdExplicitOrMTSubset = "Internal Standard Explicit: " & .InternalStandardExplicit
            Else
                If .MassTagSubsetID = -1 Then
                    lblCurrentDBInternalStdExplicitOrMTSubset = "Mass tag subset ID: "
                Else
                    lblCurrentDBInternalStdExplicitOrMTSubset = "Mass tag subset ID: " & Trim(.MassTagSubsetID)
                End If
            End If
            
            SetCheckBox chkCurrentDBAMTsOnly, .AMTsOnly
            SetCheckBox chkCurrentDBConfirmedOnly, .ConfirmedOnly
            SetCheckBox chkCurrentDBLockersOnly, .LockersOnly
            SetCheckBox chkCurrentDBLimitToPMTsFromDataset, .LimitToPMTsFromDataset
            
            If .DBSchemaVersion >= 2 Then
                blnShowDBSchema1Controls = False
            Else
                blnShowDBSchema1Controls = True
            End If
            
            chkCurrentDBAMTsOnly.Visible = blnShowDBSchema1Controls
            chkCurrentDBLockersOnly.Visible = blnShowDBSchema1Controls
            chkCurrentDBLimitToPMTsFromDataset.Visible = Not blnShowDBSchema1Controls
                        
            txtCurrentDBScoreThreshold(trfCurrentDBScoreThresholdConstants.cdstHighNormalizedScore) = .MinimumHighNormalizedScore
            txtCurrentDBScoreThreshold(trfCurrentDBScoreThresholdConstants.cdstDiscriminant) = .MinimumHighDiscriminantScore
            txtCurrentDBScoreThreshold(trfCurrentDBScoreThresholdConstants.cdstPeptideProphet) = .MinimumPeptideProphetProbability
            
            lblCurrentDBNETValueType.Caption = LookupNETValueTypeDescription(CInt(.NETValueType))
            If .ModificationList = "-1" Then
                txtCurrentDBAllowedModifications = ""
            Else
                txtCurrentDBAllowedModifications = .ModificationList
            End If
        Else
            txtCurrentDBName = ""
            lblCurrentDBMassTagCount = ""
            lblCurrentDBInternalStdExplicitOrMTSubset = ""
            chkCurrentDBAMTsOnly = vbUnchecked
            chkCurrentDBConfirmedOnly = vbUnchecked
            chkCurrentDBLockersOnly = vbUnchecked
            chkCurrentDBLimitToPMTsFromDataset = vbUnchecked
            txtCurrentDBAllowedModifications = ""
        End If
        
        ' Synchronize .AutoAnalysisDBInfo and .AnalysisInfo
        mCurrentSettings.PrefsExpanded.AutoAnalysisDBInfo = .AnalysisInfo
    End With

End Sub

Private Sub DisplayDataValidationWarnings(Optional blnResetWarningCounter As Boolean = False)
    Const WARNING_COUNT = 1
    Static dtLastDisplayTime(WARNING_COUNT) As Date
    
    If mUpdatingControls Then Exit Sub
    If blnResetWarningCounter Then Erase dtLastDisplayTime()
    
    With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement
        If .DBSearchMWTol < mCurrentSettings.AMTDef.MWTol And _
            (.RefineDBSearchMassTolerance Or .RefineMassCalibration) Then
            
            If dtLastDisplayTime(0) + 3 / 60 / 24 < Now() Then
                ' More than 3 minutes has elapsed since the last time we warned the user about this
                ' Warn them again
                MsgBox "Tolerance Refinement is enabled, but the Tolerance Refinement search tolerance (" & Trim(.DBSearchMWTol) & " " & GetSearchToleranceUnitText(.DBSearchTolType) & ", see Tab 6) is smaller than the database search tolerance (" & Trim(mCurrentSettings.AMTDef.MWTol) & " " & GetSearchToleranceUnitText(CInt(mCurrentSettings.AMTDef.TolType)) & ", see Tab 7).  Normally the Tolerance Refinement search tolerance is set larger than the database search tolerance.", vbInformation Or vbOKOnly, "Warning"
                
                dtLastDisplayTime(0) = Now()
            End If
        End If
    End With
            
End Sub

Private Sub EnableDisableControlButtons(blnEnable As Boolean)
    
    cmdClose.Enabled = blnEnable
    cmdReadFromSelectedGel.Enabled = blnEnable
    cmdApplyToSelectedGel.Enabled = blnEnable
    cmdIniFileOpen.Enabled = blnEnable
    cmdIniFileSave.Enabled = blnEnable
    cmdResetToDefaults.Enabled = blnEnable
    cboGelFileInMemoryToUse.Enabled = blnEnable
    
End Sub

Private Sub EnableDisableScanByScanAndIReport(blnEnable As Boolean)
    If cChkBox(chkPairEROptions(psoComputeERScanByScan)) <> blnEnable Then
        SetCheckBox chkPairEROptions(psoComputeERScanByScan), blnEnable
    End If
    If cChkBox(chkPairEROptions(psoIReportEREnable)) <> blnEnable Then
        SetCheckBox chkPairEROptions(psoIReportEREnable), blnEnable
    End If
End Sub

Private Sub GetOLEData(vData As Variant)
    ' Handle a drag/drop to this form
    ' From http://blackbeltvb.com/hax-ole.htm
    
    Dim CF As Long, k As Long
    Dim fso As FileSystemObject
    
    Do
        CF = CF + 1
        If CF = &H18 Then Exit Do
        If vData.GetFormat(CF) Then
            Select Case CF
                Case vbCFText
                    If TypeOf vData Is Clipboard Then
                        'Debug.Print "Text - " & CF & vbCrLf & vData.GetText
                    Else
                        'Debug.Print "Text - " & CF & vbCrLf & vData.GetData(CF)
                    End If
                Case vbCFBitmap, vbCFMetafile, vbCFEMetafile
                    ' picPicture = vData.GetData(CF)
                    'Debug.Print "Graphic - " & CF
                Case vbCFDIB
                    'Debug.Print "DIB - " & CF
                Case vbCFPalette
                    'Debug.Print "Palette - " & CF
                Case vbCFFiles
                    ' Note: can iterate through the files using:
                    Set fso = New FileSystemObject
                    For k = 1 To vData.Files.Count
                        If UCase(fso.GetExtensionName(vData.Files(k))) = "INI" Then
                            IniFileLoadSettingsLocal vData.Files(k)
                            Exit Do
                        End If
                    Next k
                    Set fso = Nothing
                    'Debug.Print "Files - " & CF
                Case vbCFRTF
                    'Debug.Print "RTF - " & CF
                Case Else
                    'Debug.Print "Unknown - " & CF
            End Select
        End If
    Loop
End Sub

Private Sub HandleMTConnectionSelectorDialogClose()
    '--------------------------------------------------
    'accept settings if new analysis is specified
    '--------------------------------------------------
        
    Dim objGelAnalysis As New FTICRAnalysis
    
    Me.MousePointer = vbHourglass

    On Error GoTo HandleMTConnectionSelectorDialogCloseErrorHandler

    If Not objMTConnectionSelector.NewAnalysis Is Nothing Then
       
        Set objGelAnalysis = objMTConnectionSelector.NewAnalysis
    
        FillDBSettingsUsingAnalysisObject mCurrentSettings.DBSettings, objGelAnalysis
        mCurrentSettings.PrefsExpanded.AutoAnalysisDBInfo = mCurrentSettings.DBSettings.AnalysisInfo
    End If
    
HandleMTConnectionSelectorDialogCloseContinue:
    Set objMTConnectionSelector = Nothing
    
    EnableDisableControlButtons True
    
    ' Determine number of matching MT tags for the given settings
    mCurrentSettings.DBSettings.SelectedMassTagCount = GetMassTagMatchCount(mCurrentSettings.DBSettings, LookupCurrentJob(), Me)
    mCurrentSettings.DBSettings.IsDeleted = False
    
    DisplayCurrentDBSettings
    
    Me.MousePointer = vbDefault

    Exit Sub
    
HandleMTConnectionSelectorDialogCloseErrorHandler:
    LogErrors Err.Number, "frmEditAnalysisSettings.objMTConnectionSelector_DialogClosed"
    MsgBox "Error initiating new dummy analysis.", vbOKOnly
    Resume HandleMTConnectionSelectorDialogCloseContinue

End Sub

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

Private Sub IniFileLoadSettingsLocal(Optional strNewIniFilePath As String = "")
    
    Dim fso As New FileSystemObject
    
    If Len(strNewIniFilePath) = 0 Then
        strNewIniFilePath = SelectFile(Me.hwnd, "Select .Ini file to load", fso.GetParentFolderName(mIniFilePathSaved), False, fso.GetFileName(mIniFilePathSaved), "Ini files (*.ini)|*.ini|All Files (*.*)|*.*", 0)
    End If
    
    Set fso = Nothing
    If Len(strNewIniFilePath) = 0 Then Exit Sub
    
    mIniFilePathSaved = strNewIniFilePath
    mAutoIniFileNameOverridden = False

    fraOptionFrame(OPTION_FRAME_CONTROLS).Visible = False
    lblWorking.Visible = True
    lblWorking = "Loading settings from .Ini file:" & vbCrLf & CompactPathString(strNewIniFilePath, 125)
    Me.MousePointer = vbHourglass
    DoEvents
    
    With mCurrentSettings
        IniFileLoadSettings .PrefsExpanded, .UMCDef, .UMCIonNetDef, .UMCNetAdjDef, .UMCInternalStandards, .AMTDef, .Prefs, mIniFilePathSaved, True
    
        .DBSettings.AnalysisInfo = .PrefsExpanded.AutoAnalysisDBInfo
        
        FillDBSettingsUsingAnalysisInfoUDT .DBSettings
    End With
    
    DisplayCurrentAnalysisSettings

    Me.MousePointer = vbDefault
    lblWorking.Visible = False
    fraOptionFrame(OPTION_FRAME_CONTROLS).Visible = True
    DoEvents
    
End Sub

Private Sub IniFileSaveSettingsLocal()
    Dim strNewIniFilePath As String
    Dim fso As FileSystemObject
    Dim strSuggestedName As String, strSuggestedNameToDisplay As String
    Dim strUMCSearchType As String
    Dim intIndex As Integer
    Dim blnExportDB As Boolean
    
    Dim blnUMCIonNetEnabled As Boolean
    Dim dblTolPPM As Double
    Dim eTolType As glMassToleranceConstants
    Dim eTolTypeActual As glMassToleranceConstants
    
On Error GoTo IniFileSaveSettingsLocalErrorHandler

    With mCurrentSettings
        
        With .PrefsExpanded.AutoAnalysisFilterPrefs
            If .RestrictToEvenScanNumbersOnly Then
                strSuggestedName = "EvenScans_"
            ElseIf .RestrictToOddScanNumbersOnly Then
                strSuggestedName = "OddScans_"
            Else
                strSuggestedName = ""
            End If
        End With
        
        blnUMCIonNetEnabled = .PrefsExpanded.AutoAnalysisOptions.UMCSearchMode = AUTO_ANALYSIS_UMCIonNet
        If blnUMCIonNetEnabled Then
            LookupUMCIonNetMassTolerances dblTolPPM, eTolType, .UMCIonNetDef, eTolTypeActual, UMC_IONNET_PPM_CONVERSION_MASS
            
            If dblTolPPM < 0 Then
                ' This is unexpected
                Debug.Assert False
                dblTolPPM = 0
            End If
            
            strSuggestedName = strSuggestedName & "UMCIonNet"
            If eTolTypeActual = gltPPM Then
                strSuggestedName = strSuggestedName & "_" & Trim(dblTolPPM) & GetSearchToleranceUnitText(eTolType) & "Max"
            End If
        Else
            strSuggestedName = strSuggestedName & "UMC" & Trim(.UMCDef.Tol) & GetSearchToleranceUnitText(CInt(.UMCDef.TolType))
                    
            Select Case .PrefsExpanded.AutoAnalysisOptions.UMCSearchMode
            Case AUTO_ANALYSIS_UMCListType2002, AUTO_ANALYSIS_UMC2003
                Select Case .UMCDef.UMCType
                Case 0: strUMCSearchType = "FHI"
                Case 1: strUMCSearchType = "FBF"
                Case 2: strUMCSearchType = "MinCnt"
                Case 3: strUMCSearchType = "MaxCnt"
                Case 4: strUMCSearchType = "UnqAmt"
                Case 5: strUMCSearchType = "SBFI"
                Case 6: strUMCSearchType = "SBFF"
                Case Else: strUMCSearchType = ""
                End Select
            Case Else
                strUMCSearchType = .PrefsExpanded.AutoAnalysisOptions.UMCSearchMode
            End Select
            
            strSuggestedName = strSuggestedName & "_" & strUMCSearchType
        End If
        
        If .PrefsExpanded.PairSearchOptions.PairSearchMode <> AUTO_FIND_PAIRS_NONE Then
            strSuggestedName = strSuggestedName & "_FindPairs"
        End If
        
        With .PrefsExpanded.PairSearchOptions
            If .AutoAnalysisRemovePairMemberHitsAfterDBSearch Then
                If .AutoAnalysisRemovePairMemberHitsRemoveHeavy Then
                    strSuggestedName = strSuggestedName & "_RemoveHeavyHits"
                Else
                    strSuggestedName = strSuggestedName & "_RemoveLightHits"
                End If
            End If
        End With
        
        If .PrefsExpanded.NetAdjustmentUsesN15AMTMasses Then
            strSuggestedName = strSuggestedName & "_N15AMT"
        End If
        
        If .PrefsExpanded.AutoAnalysisOptions.SkipFindUMCs Then
            strSuggestedName = strSuggestedName & "_SkipFindUMCs"
        End If
        
        If .PrefsExpanded.AutoAnalysisOptions.SkipGANETSlopeAndInterceptComputation Then
            strSuggestedName = strSuggestedName & "_SkipNetAdj"
        Else
            strSuggestedName = strSuggestedName & "_NetAdj"
            If RobustNETWarpEnabled() Then
                strSuggestedName = strSuggestedName & "Warp"
            End If
            strSuggestedName = strSuggestedName & "Tol" & Trim(.UMCNetAdjDef.MWTol) & GetSearchToleranceUnitText(CInt(.UMCNetAdjDef.MWTolType))
        End If
        
        strSuggestedName = strSuggestedName & "_AMT" & Trim(.AMTDef.MWTol) & GetSearchToleranceUnitText(CInt(.AMTDef.TolType))
        strSuggestedName = strSuggestedName & "_" & Trim(.AMTDef.NETTol) & "NET"
        
        If .AMTDef.UseDriftTime Then
            strSuggestedName = strSuggestedName & "_" & Trim(.AMTDef.DriftTimeTol) & "msecDT"
        End If
        
        With .PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement
            If .RefineMassCalibration Or .RefineDBSearchMassTolerance Or .RefineDBSearchNETTolerance Then
                strSuggestedName = strSuggestedName & "_RefineTol"
            End If
        End With
        
        With .PrefsExpanded.AutoAnalysisOptions
            For intIndex = 0 To .AutoAnalysisSearchModeCount - 1
                If .AutoAnalysisSearchMode(intIndex).ExportResultsToDatabase Then
                    blnExportDB = True
                End If
            Next intIndex
        End With
        
        If blnExportDB Then
            strSuggestedName = strSuggestedName & "_Export"
            If .PrefsExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches Then
                strSuggestedName = strSuggestedName & "AllUMCs"
            End If
        End If
        
        With .PrefsExpanded.AutoAnalysisOptions
            If .SavePictureGraphic Or .SaveErrorGraphicMass Or .SaveErrorGraphicGANET Or .SaveErrorGraphic3D Or .SaveInternalStdHitsAndData Then
                strSuggestedName = strSuggestedName & "_SaveGraphics"
                If .SaveGelFile Then
                    strSuggestedName = strSuggestedName & "AndGel"
                End If
            ElseIf .SaveGelFile Then
                strSuggestedName = strSuggestedName & "_SaveGel"
            End If
            If .WriteIDResultsByIonToTextFileAfterAutoSearches Then
                strSuggestedName = strSuggestedName & "_SavePeaks"
            End If
            If .SaveUMCStatisticsToTextFile Then
                strSuggestedName = strSuggestedName & "_SaveUMCs"
            End If
        End With
        
        strSuggestedName = strSuggestedName & "_" & Format(Now(), "yyyy-mm-dd") & ".ini"
        
    End With
    
    Set fso = New FileSystemObject
    
    If mAutoIniFileNameOverridden And Len(mIniFilePathSaved) > 0 Then
        ' User previously over-rode the suggested name; thus, use the previously used name
        strSuggestedNameToDisplay = fso.GetFileName(mIniFilePathSaved)
    Else
        strSuggestedNameToDisplay = strSuggestedName
    End If
        
    strNewIniFilePath = SelectFile(Me.hwnd, "Select .Ini file to save to", fso.GetParentFolderName(mIniFilePathSaved), True, strSuggestedNameToDisplay, "Ini files (*.ini)|*.ini|All Files (*.*)|*.*", 0)
    If Len(strNewIniFilePath) = 0 Then
        Set fso = Nothing
        Exit Sub
    End If
    
    If LCase(fso.GetFileName(strNewIniFilePath)) = LCase(strSuggestedName) Then
        mAutoIniFileNameOverridden = False
    Else
        mAutoIniFileNameOverridden = True
    End If
    
    If LCase(fso.GetExtensionName(strNewIniFilePath)) <> "ini" Then
        strNewIniFilePath = strNewIniFilePath & ".ini"
    End If
    
    Set fso = Nothing
    
    mIniFilePathSaved = strNewIniFilePath
    
    fraOptionFrame(OPTION_FRAME_CONTROLS).Visible = False
    lblWorking.Visible = True
    lblWorking = "Saving settings to .Ini file:" & vbCrLf & CompactPathString(strNewIniFilePath, 125)
    Me.MousePointer = vbHourglass
    DoEvents
    
    With mCurrentSettings
        .PrefsExpanded.LastAutoAnalysisIniFilePath = mIniFilePathSaved
        IniFileSaveSettings .PrefsExpanded, .UMCDef, .UMCIonNetDef, .UMCNetAdjDef, .UMCInternalStandards, .AMTDef, .Prefs, mIniFilePathSaved, True
    End With
    glbPreferencesExpanded.LastAutoAnalysisIniFilePath = mIniFilePathSaved
    
    Me.MousePointer = vbDefault
    lblWorking.Visible = False
    fraOptionFrame(OPTION_FRAME_CONTROLS).Visible = True
    DoEvents
    
    Exit Sub
    
IniFileSaveSettingsLocalErrorHandler:
    Debug.Print "Error in IniFileSaveSettingsLocal: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmEditAnalysisSettings->IniFileSaveSettingsLocal"
    Resume Next
    
End Sub

Private Sub InitializeForm()
    Me.OLEDropMode = 1
    
    SizeAndCenterWindow Me, cWindowUpperThird, 11550, 10400, False
    
    PositionControls
    tbsTabStrip.Tab = 0
    tbsUMCRefinementOptions.Tab = 0
    
    PopulateComboBoxes
    
    lblToleranceRefinementExplanation = ""
    lblToleranceRefinementExplanation = lblToleranceRefinementExplanation & "Tolerance Refinement involves performing an initial database search using the tolerances defined below and the search method defined on the next tab. "
    lblToleranceRefinementExplanation = lblToleranceRefinementExplanation & "Following this, Mass Error and NET Error plots are produced, and the peak apex and peak width are determined.  "
    lblToleranceRefinementExplanation = lblToleranceRefinementExplanation & "The location of the peak apex can be used to dynamically re-calibrate the data by applying a mass shift to all data points in memory.  "
    lblToleranceRefinementExplanation = lblToleranceRefinementExplanation & "The width of the peak can be used to determine the optimum mass tolerance to use for database searching.  "
    lblToleranceRefinementExplanation = lblToleranceRefinementExplanation & "Lastly, the width of the NET Error plot peak can be used to determine the optimum NET tolerance to use for database searching."
    
    lblToleranceRefinementWarning = "Note: Since Tolerance Refinement is enabled, the above mass and NET tolerances are used as the search tolerances only if tolerance refinement fails."
    
    chkMTAdditionalMass(cbamPEO).ToolTipText = glPEO & " Da"
    chkMTAdditionalMass(cbamICATd0).ToolTipText = glICAT_Light & " Da"
    chkMTAdditionalMass(cbamICATd8).ToolTipText = glICAT_Heavy & " Da"
    
    chkMTAdditionalMass(cbamAlkylation).ToolTipText = "Check to add the alkylation mass correction below to all MT tag masses (added to each cys residue)"
    txtAlkylationMWCorrection = glALKYLATION
    
    txtPEKFileExtensionPreferenceOrder.Text = DEFAULT_PEK_FILE_EXTENSION_ORDER
    
    With fraOptionFrame(OPTION_FRAME_LABEL_PAIRS)
        .Top = fraOptionFrame(OPTION_FRAME_DELTA_PAIRS).Top
        .Left = fraOptionFrame(OPTION_FRAME_DELTA_PAIRS).Left
    End With
    
    fraOptionFrame(OPTION_FRAME_DELTA_PAIRS).Visible = True
    fraOptionFrame(OPTION_FRAME_LABEL_PAIRS).Visible = False

    UpdateDynamicControls
    
    chkCurrentDBLimitToPMTsFromDataset.Left = chkCurrentDBConfirmedOnly.Left
    chkCurrentDBLimitToPMTsFromDataset.Visible = False

    ''tmrTimer.Interval = 500
    ''tmrTimer.Enabled = True
    ''mRequestUpdateRobustNETIterationCount = True

    mAutoIniFileNameOverridden = False
    mIniFilePathSaved = ""

    ResetToDefaults
    mCurrentSettings.DBSettings.IsDeleted = True
    mSettingsSaved = mCurrentSettings
    
    On Error Resume Next
    With mCurrentSettings
        .PrefsExpanded = glbPreferencesExpanded
        If UBound(GelBody) >= 1 Then
            .Prefs = GelData(1).Preferences
            .UMCDef = GelSearchDef(1).UMCDef
            .UMCIonNetDef = GelSearchDef(1).UMCIonNetDef
            .UMCNetAdjDef = GelUMCNETAdjDef(1)
            .AMTDef = GelSearchDef(1).AMTSearchOnUMCs
            FillDBSettingsUsingAnalysisObject .DBSettings, GelAnalysis(1)
        End If
    End With
    DisplayCurrentAnalysisSettings

End Sub

Private Sub LookupDefaultUMCIonNetValues(ByVal intControlIndex As Integer, ByRef dblMinimum As Double, ByRef dblMaximum As Double, ByRef dblDefault As Double, ByRef blnReturnDefaultWeightFactor As Boolean)
    ' If blnReturnDefaultWeightFactor = True, then returns the default Weight Factor
    ' If blnReturnDefaultWeightFactor = False, then returns the default constraint
    
    If intControlIndex < 0 Or intControlIndex > cmbData.Count - 1 Then
        Debug.Assert False
        Exit Sub
    End If
    
    dblMinimum = 0
    dblMaximum = 100

    Select Case cmbData(intControlIndex).ListIndex
    Case uindUMCIonNetDimConstants.uindMonoMW, uindUMCIonNetDimConstants.uindAvgMW, uindUMCIonNetDimConstants.uindTmaMW
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.5
        Else
            If cmbConstraintUnits(intControlIndex).ListIndex = DATA_UNITS_MASS_DA Then
                dblDefault = 0.015
                dblMaximum = 100
            Else
                dblDefault = 15
                dblMaximum = 100000
            End If
        End If
    Case uindUMCIonNetDimConstants.uindScan
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.01
        Else
            ' Scan doesn't normally have a constraint
            dblDefault = 1
            dblMinimum = 1
            dblMaximum = max_scan
        End If
    Case uindUMCIonNetDimConstants.uindFit
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.1
        Else
            ' Fit doesn't normally have a constraint
            dblDefault = 0.2
            dblMaximum = 1
        End If
    Case uindUMCIonNetDimConstants.uindMZ
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.1
        Else
            ' m/z doesn't normally have a constraint
            dblDefault = 1
            dblMaximum = 100
        End If
    Case uindUMCIonNetDimConstants.uindGenericNET
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.1
        Else
            ' Generic NET doesn't normally have a constraint
            dblDefault = 0.1
            dblMaximum = 100
        End If
    Case uindUMCIonNetDimConstants.uindChargeState
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.1
        Else
            ' charge state doesn't normally have a constraint
            dblDefault = 1
            dblMaximum = 100
        End If
    Case uindUMCIonNetDimConstants.uindLogAbundance
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.1
        Else
            ' Abundance doesn't normally have a constraint
            dblDefault = 1
            dblMaximum = 100000
        End If
    Case uindUMCIonNetDimConstants.uindIMSDriftTime
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.1
            ' Drift time doesn't normally have a constraint
            dblMinimum = 0
            dblMaximum = 100000
        End If
    Case Else
        Debug.Assert False
        If blnReturnDefaultWeightFactor Then
            dblDefault = 0.1
        Else
            dblDefault = 1
            dblMaximum = 100
        End If
    End Select
    
End Sub

Private Function LookupCurrentJob() As Long
    Dim lngCurrentJob As Long

On Error Resume Next

    If Not GelAnalysis(mActiveGelIndex) Is Nothing Then
        LookupCurrentJob = GelAnalysis(mActiveGelIndex).Job
    Else
        LookupCurrentJob = 0
    End If
    
End Function

Private Sub PopulateComboBoxes()
    Dim intLoadedGelCount As Integer
    Dim intGelIndex As Integer
    Dim intIndex As Integer
    
    mUpdatingControls = True

    With cboEvenOddScanNumber
        .Clear
        .AddItem "Use All Scans"
        .AddItem "Use Odd Scans"
        .AddItem "Use Even Scans"
        .ListIndex = eosAllScans
    End With

    With cboUMCSearchMode
        .Clear
        .AddItem AUTO_ANALYSIS_UMCListType2002 & " (obsolete)", usmUMC2002
        .AddItem AUTO_ANALYSIS_UMC2003, usmUMC2003
        .AddItem AUTO_ANALYSIS_UMCIonNet, usmUMCIonNet
        .ListIndex = usmUMCIonNet
    End With
    
    With cboDBSearchMode
        .Clear
        .AddItem "Export LC-MS Features Only (No DB Search)", dbsmExportUMCsOnly
        .AddItem "Individual peaks", dbsmIndividualPeaks
        .AddItem "Individual peaks in LC-MS Features without NET", dbsmIndividualPeaksInUMCsWithoutNET
        .AddItem "Peaks in LC-MS Features with NET (not supported)", dbsmIndividualPeaksInUMCsWithNET
        .AddItem "Conglomerate LC-MS Feature masses with NET (preferred)", dbsmConglomerateUMCsWithNET
        .AddItem "Peaks in LC-MS Features with NET, paired (not supported)", dbsmIndividualPeaksInUMCsPaired
        .AddItem "Peaks in LC-MS Features with NET, unpaired (not supported)", dbsmIndividualPeaksInUMCsUnpaired
        .AddItem "Conglomerate LC-MS Feature masses - Paired Only", dbsmConglomerateUMCsPaired
        .AddItem "Conglomerate LC-MS Feature masses - Unpaired Only", dbsmConglomerateUMCsUnpaired
        .AddItem "N14/N15 pairs (individual LC-MS Feature peaks, not supported)", dbsmPairsN14N15
        .AddItem "N14/N15 pairs (conglomerate LC-MS Feature masses)", dbsmPairsN14N15ConglomerateMass
        .AddItem "ICAT pairs (individual peaks in LC-MS Features)", dbsmPairsICAT
        .AddItem "PEO pairs (individual peaks in LC-MS Features)", dbsmPairsPEO
        .AddItem "Congolmerate LC-MS Feature masses - Light Pairs Plus Unpaired", dbsmConglomerateUMCsLightPairsPlusUnpaired
        .ListIndex = dbsmConglomerateUMCsWithNET
    End With
    UpdateDBSearchModeTooltip
    
    ' Note: We must populate the Pairs comboboxes after populating cboDBSearchMode due to ValidatePairsIdentificationMode()
    With cboPairsIdentificationMode
        .Clear
        .AddItem AUTO_FIND_PAIRS_NONE, pimNone
        .AddItem "Delta Pairs (e.g N14/N15)", pimDelta
        .AddItem "Label Pairs (e.g. ICAT or PEO)", pimLabel
        .ListIndex = pimNone
    End With
    
    With cboPairsUMCsToUseForNETAdjustment
        .Clear
        .AddItem "All LC-MS Features, regardless of pair or light/heavy status", punaPairedAndUnpaired
        .AddItem "Unpaired LC-MS Features only", punaUnpairedOnly
        .AddItem "Unpaired LC-MS Features plus light members of paired LC-MS Features", punaUnpairedPlusPairedLight
        .AddItem "Paired LC-MS Features, both light and heavy members", punaPairedAll
        .AddItem "Paired LC-MS Features, light members only", punaPairedLight
        .AddItem "Paired LC-MS Features, heavy members only", punaPairedHeavy
        .ListIndex = punaUnpairedPlusPairedLight
    End With
    
    With cboRobustNETAdjustmentMode
        .Clear
        .AddItem "Iterative mode"
        If APP_BUILD_DISABLE_LCMSWARP Then
            .ListIndex = UMCRobustNETModeConstants.UMCRobustNETIterative
        Else
            .AddItem "MS Warp NET Alignment"
            .AddItem "MS Warp NET and Mass Alignment"
            .ListIndex = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass
        End If
    End With

    With cboMSWarpMassCalibrationType
        .Clear
        .AddItem "Recalibrate m/z coefficients"
        .AddItem "Recalibrate mass vs. elution time"
        .AddItem "Hybrid Recalibration (preferred)"
        .ListIndex = rmcUMCRobustNETWarpMassCalibrationType.rmcHybridRecal
    End With
    
    With cboAverageERsWeightingMode
        .Clear
        .AddItem "No weighting"
        .AddItem "Weight by Abu"
        .AddItem "Weight by Members"
        .ListIndex = aewAbundance
    End With
    
    With cboSavePictureGraphicFileType
        .Clear
        .AddItem "PNG file (.PNG)"
        .AddItem "Jpeg file (.JPG)"
        .AddItem "Windows Meta File (.WMF)"
        .AddItem "Extended Meta File (.EMF)"
        .AddItem "Windows Bitmap (.BMP)"
        .ListIndex = 0
    End With
    
''    With cboSaveErrorGraphicFileType
''        .Clear
''        .AddItem "PNG"
''        .AddItem "Jpeg file (.JPG)"
''        .ListIndex = 1
''    End With
    
    On Error Resume Next
    intLoadedGelCount = UBound(GelBody())
    
    With cboGelFileInMemoryToUse
        .Clear
        For intGelIndex = 1 To intLoadedGelCount
            .AddItem Trim(intGelIndex) & ": " & CompactPathString(GelBody(intGelIndex).Caption, 46)
        Next intGelIndex
        
        If .ListCount > 0 Then
            If mGelIndexInstantiatingForm > 0 And mGelIndexInstantiatingForm <= .ListCount Then
                .ListIndex = mGelIndexInstantiatingForm - 1
            Else
                .ListIndex = 0
            End If
        End If
    End With
    
    ' UMC Search Options
    With cmbCountType
        .Clear
        .AddItem "Favor Higher Intensity"
        .AddItem "Favor Better Fit"
        .AddItem "Minimize Count"
        .AddItem "Maximize Count"
        .AddItem "Unique MT"
        .AddItem "Shrinking Box Favor Intensity"
        .AddItem "Shrinking Box Favor Fit"
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
    
    With cboChargeStateAbuType
        .Clear
        .AddItem "Highest Abu Sum"
        .AddItem "Most Abu Member"
        .AddItem "Most Members"
    End With
    
    ' UMCIonNet Options
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
            .AddItem "Drift Time"
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
    End With
    
    With cmbUMCMW
        .Clear
        .AddItem "Class Average"
        .AddItem "Mol.Mass Of Class Representative"
        .AddItem "Class Median"
    End With
    
    With cboSplitUMCsScanGapBehavior
        .Clear
        .AddItem "Ignore scan gaps"
        .AddItem "Split if mass difference"
        .AddItem "Always split"
    End With
    
    With cboInternalStdSearchMode
        .Clear
        .AddItem "Search only MT tags", issmFindOnlyMassTags
        .AddItem "Search MT tags & lockers", issmFindWithMassTags
        .AddItem "Search only Internal Stds", issmFindOnlyInternalStandards
        
        If APP_BUILD_DISABLE_MTS Then
            .ListIndex = issmFindOnlyMassTags
        Else
            .ListIndex = issmFindWithMassTags
        End If
    End With
        
    ' Tolerance refinement
    With cboToleranceRefinementMethod
        .Clear
        .AddItem "Expectation Maximization"
        .AddItem "Mass Error Plot Width at % of Max"
        .AddItem "Median LC-MS Feature Mass StDev"
        .AddItem "Maximum LC-MS Feature Mass StDev"
        .AddItem "Median LC-MS Feature Mass Width"
        .AddItem "Maximum LC-MS Feature Mass Width"
        .ListIndex = mtrExpectationMaximization
    End With

    For intIndex = 0 To 1
        With cboSearchRegionShape(intIndex)
            .Clear
            .AddItem "Elliptical search region"
            .AddItem "Rectangular search region"
            .ListIndex = srsSearchRegionShapeConstants.srsElliptical
        End With
    Next intIndex
    
    mUpdatingModMassControls = True
    With cboResidueToModify
        .Clear
        .AddItem "Full MT"
        For intIndex = 0 To 25
            .AddItem Chr(vbKeyA + intIndex)
        Next intIndex
        .AddItem glPHOSPHORYLATION
        .ListIndex = 0
    End With
    mUpdatingModMassControls = False
    
    mUpdatingControls = False
        
End Sub

Private Sub PositionControls()
    
    fraUMC2002UniqueOptions.BackColor = &H8000000F
    fraUMC2003UniqueOptions.BackColor = &H8000000F
    
    fraUMC2003UniqueOptions.Top = fraUMC2002UniqueOptions.Top
    
    fraUMCSearch200x.Visible = True
    fraUMCIonNetOptions.Visible = False
    
    With fraUMCSearch200x
        .Top = fraUMCIonNetOptions.Top
        .Left = fraUMCIonNetOptions.Left
    End With
    
    fraSelectingMassTags.Visible = False
    
    With lblWorking
        .Visible = False
        .Top = fraOptionFrame(OPTION_FRAME_CONTROLS).Top + 120
        .Left = fraOptionFrame(OPTION_FRAME_CONTROLS).Left + 120
        .Height = 675
    End With
    
    fraOptionFrame(OPTION_FRAME_CONTROLS).Top = tbsTabStrip.Top + tbsTabStrip.Height + 120
    
End Sub

''Private Function PredictRobustNETIterationCount() As Long
''
''    Const MAX_ROBUST_NET_ITERATION_COUNT As Long = 100000
''
''    Dim lngIterationCount As Long
''    Dim sngSlope As Single, sngIntercept As Single, sngMassShiftPPM As Single
''
''    lngIterationCount = 0
''
''    With mCurrentSettings.UMCNetAdjDef
''        sngMassShiftPPM = .RobustNETMassShiftPPMStart
''        Do While sngMassShiftPPM <= .RobustNETMassShiftPPMEnd
''            sngSlope = .RobustNETSlopeStart
''            Do While sngSlope <= .RobustNETSlopeEnd
''                sngIntercept = .RobustNETInterceptStart
''                Do While sngIntercept <= .RobustNETInterceptEnd
''                    lngIterationCount = lngIterationCount + 1
''                    If lngIterationCount >= MAX_ROBUST_NET_ITERATION_COUNT Then Exit Do
''                    IncrementRobustNETSetting sngIntercept, UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear, .RobustNETInterceptIncrement
''                Loop
''                If lngIterationCount >= MAX_ROBUST_NET_ITERATION_COUNT Then Exit Do
''                IncrementRobustNETSetting sngSlope, .RobustNETSlopeIncreaseMode, .RobustNETSlopeIncrement
''            Loop
''            If lngIterationCount >= MAX_ROBUST_NET_ITERATION_COUNT Then Exit Do
''            IncrementRobustNETSetting sngMassShiftPPM, UMCRobustNETIncrementConstants.UMCRobustNETIncrementLinear, .RobustNETMassShiftPPMIncrement
''        Loop
''    End With
''
''    lblRobustNETPredictedIterationCount = "Predicted Iteration Count: " & Trim(lngIterationCount)
''
''    PredictRobustNETIterationCount = lngIterationCount
''
''End Function

Private Sub ReadFromSelectedGel()
    
    If cboGelFileInMemoryToUse.ListCount <= 0 Then
        MsgBox "No gels are present in memory; unable to continue.", vbInformation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If mActiveGelIndex < 1 Or mActiveGelIndex > UBound(GelBody()) Then
        MsgBox "A valid gel is not selected in the 'Gel file (in memory) to read or update' list."
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    DoEvents
    
    With mCurrentSettings
        .Prefs = GelData(mActiveGelIndex).Preferences
        .PrefsExpanded = glbPreferencesExpanded
        .PrefsExpanded.LegacyAMTDBPath = GelData(mActiveGelIndex).PathtoDatabase
        
        .UMCDef = GelSearchDef(mActiveGelIndex).UMCDef
        .UMCIonNetDef = GelSearchDef(mActiveGelIndex).UMCIonNetDef
        .UMCNetAdjDef = GelUMCNETAdjDef(mActiveGelIndex)
        .AMTDef = GelSearchDef(mActiveGelIndex).AMTSearchOnUMCs
        FillDBSettingsUsingAnalysisObject .DBSettings, GelAnalysis(mActiveGelIndex)
            
        ' Determine number of matching MT tags for the given settings
        If .DBSettings.SelectedMassTagCount = 0 Then
            .DBSettings.SelectedMassTagCount = GetMassTagMatchCount(.DBSettings, LookupCurrentJob(), Me)
        End If
        
        With .PrefsExpanded.AutoAnalysisOptions
            If .AutoAnalysisSearchModeCount > 0 Then
                .AutoAnalysisSearchMode(0).MassMods = GelSearchDef(mActiveGelIndex).AMTSearchMassMods
            End If
        End With
    End With
    
    DisplayCurrentAnalysisSettings
    Me.MousePointer = vbDefault
    
End Sub

Private Sub RemoveSelectedSearchMode()
    Dim intIndex As Integer
    Dim intIndexToRemove As Integer
    
On Error GoTo RemoveSelectedSearchModeErrorHandler

    If lstDBSearchModes.ListCount < 1 Then Exit Sub
    
    intIndexToRemove = -1
    For intIndex = 0 To lstDBSearchModes.ListCount - 1
        If lstDBSearchModes.Selected(intIndex) Then
            intIndexToRemove = intIndex
            Exit For
        End If
    Next intIndex
    
    If intIndexToRemove >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions
            If .AutoAnalysisSearchModeCount = 1 Then
                .AutoAnalysisSearchModeCount = 0
                lstDBSearchModes.Clear
            Else
                For intIndex = intIndexToRemove To .AutoAnalysisSearchModeCount - 1
                    .AutoAnalysisSearchMode(intIndex) = .AutoAnalysisSearchMode(intIndex + 1)
                Next intIndex
                
                .AutoAnalysisSearchModeCount = .AutoAnalysisSearchModeCount - 1
                
                lstDBSearchModes.Clear
                For intIndex = 0 To .AutoAnalysisSearchModeCount - 1
                    lstDBSearchModes.AddItem .AutoAnalysisSearchMode(intIndex).SearchMode
                Next intIndex
                If intIndexToRemove < lstDBSearchModes.ListCount Then
                    lstDBSearchModes.ListIndex = intIndexToRemove
                Else
                    lstDBSearchModes.ListIndex = lstDBSearchModes.ListCount - 1
                End If
            End If
        End With
    
        UpdateDynamicControls
    End If
    
    Exit Sub

RemoveSelectedSearchModeErrorHandler:
    Debug.Print "Error in RemoveSelectedSearchMode: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmEditAnalysisSettings->RemoveSelectedSearchMode"
    Resume Next
    
End Sub

Private Sub ResetToDefaults()
    Dim strLegacyDBPathSaved As String
    
    With mCurrentSettings
        strLegacyDBPathSaved = .PrefsExpanded.LegacyAMTDBPath
                
        ResetGelPrefs .Prefs
        ResetExpandedPreferences .PrefsExpanded, "", True
        
        SetDefaultUMCDef .UMCDef
        SetDefaultUMCIonNetDef .UMCIonNetDef
        SetDefaultUMCNETAdjDef .UMCNetAdjDef
        
        SetDefaultSearchAMTDef .AMTDef, .UMCNetAdjDef
        
        If APP_BUILD_DISABLE_MTS Then
            .PrefsExpanded.LegacyAMTDBPath = strLegacyDBPathSaved
        End If
    End With
    
    SetCheckBox chkCSIsoSameRangeAbu, True
    SetCheckBox chkCSIsoSameRangeMW, True
    
    DisplayCurrentAnalysisSettings
    UpdateUMCIonNetFindSingleMemberClasses
End Sub

Private Sub RevertToSavedSettings()
    Dim eResponse As VbMsgBoxResult
    
    eResponse = MsgBox("Restore saved analysis settings?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Revert")
    
    If eResponse = vbYes Then
        mCurrentSettings = mSettingsSaved
        DisplayCurrentAnalysisSettings
    End If
End Sub

Private Function RobustNETWarpEnabled() As Boolean
    If cboRobustNETAdjustmentMode.ListIndex >= UMCRobustNETModeConstants.UMCRobustNETWarpTime And cChkBox(chkRobustNETEnabled) Then
        RobustNETWarpEnabled = True
    Else
        RobustNETWarpEnabled = False
    End If
End Function

Private Sub SelectAlternateOutputFolder()
    Dim strNewFolder As String
    
    strNewFolder = BrowseForFileOrFolder(Me.hwnd, txtDBSearchAlternateOutputFolderPath, "Select Output Folder", True)
    
    If Len(strNewFolder) > 0 Then
        txtDBSearchAlternateOutputFolderPath = strNewFolder
    End If
End Sub

Private Sub SelectLegacyDatabase()
    Dim strNewFilePath As String
    
    With mCurrentSettings.PrefsExpanded
        strNewFilePath = SelectLegacyMTDB(Me, .LegacyAMTDBPath)
        If Len(strNewFilePath) > 0 Then
            txtLegacyAMTDatabasePath.Text = strNewFilePath
            .LegacyAMTDBPath = strNewFilePath
        End If
    End With
End Sub

Private Sub SelectMassTagsForCurrentDB()
    
    If mCurrentSettings.DBSettings.IsDeleted Then
        MsgBox "Cannot select MT tags since no database is currently defined.  Please choose Select Other DB", vbExclamation + vbOKOnly, "No connection"
        Exit Sub
    End If
    
    lblSelectingMassTags.Caption = "The MT tag selection window should now be visible.  When done selecting the MT tags, press OK on that Window.  Next, press the OK button below.  To cancel (or if you can't see the window), select Cancel."
    
    With fraSelectingMassTags
        .Left = fraOptionFrame(OPTION_FRAME_CURRENT_CONNECTION).Left
        .Top = fraOptionFrame(OPTION_FRAME_CURRENT_CONNECTION).Top
        .width = tbsTabStrip.width - 300
        .Height = tbsTabStrip.Height - 750
        .Visible = True
        .ZOrder
    End With
    
    If objSelectMassTags Is Nothing Then
        Set objSelectMassTags = New FTICRAnalysis
    End If
        
    FillGelAnalysisObject objSelectMassTags, mCurrentSettings.DBSettings.AnalysisInfo
    
    ' Use the following to display the MT tags selection window
    ' Unfortunately, there is no way to wait for this to finish
    ' Thus the reason for fraSelectingMassTags above, which fills the window to cover the other controls,
    '  and requires the user to click OK or Cancel when done selecting MT tags
    
    objSelectMassTags.MTDB.SelectMassTags glInitFile
    
End Sub

Public Sub SetCallerID(lngCallerID As Long)
    ' Use lngCallerID = 0 to indicate no gel (i.e. form called when no gels were loaded)
    
    mGelIndexInstantiatingForm = lngCallerID
    
    If mGelIndexInstantiatingForm > 0 And cboGelFileInMemoryToUse.ListCount >= mGelIndexInstantiatingForm Then
        SetComboBox cboGelFileInMemoryToUse, mGelIndexInstantiatingForm - 1, "Gel File in Memory"
        ReadFromSelectedGel
        mSettingsSaved = mCurrentSettings
    End If
    
End Sub

Private Sub SetComboBox(ByRef cboControl As ComboBox, ByVal intListIndex As Integer, ByVal strComboBoxDescription As String)

    On Error GoTo SetComboBoxErrorHandler

    cboControl.ListIndex = intListIndex

    Exit Sub

SetComboBoxErrorHandler:
   
    Dim strMessage As String
    strMessage = "Error updating combo box '" & strComboBoxDescription & "' (" & cboControl.Name & ") to select item at index " & intListIndex & ": " & Err.Description
    
    Debug.Assert False
    MsgBox strMessage, vbExclamation + vbOKOnly, "Error"

End Sub

Private Sub SetPairSearchDeltas(dblDeltaMass As Double, DeltaCountMin As Integer, DeltaCountMax As Integer, Optional DeltaStepSize As Long = 1)
    txtPairsDelta = dblDeltaMass
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaMass = dblDeltaMass
    
    
    txtPairsMinMaxDelta(0) = DeltaCountMin
    txtPairsMinMaxDelta(1) = DeltaCountMax
    txtPairsMinMaxDelta(2) = DeltaStepSize
    
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaCountMin = DeltaCountMin
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaCountMax = DeltaCountMax
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaStepSize = DeltaStepSize
End Sub

Private Sub ShowHidePNNLMenus()
    Dim strMessage As String
    
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    lblDescription(LABEL_DESCRIPTION_PRISM_MESSAGE).Visible = blnVisible
    fraOptionFrame(OPTION_FRAME_CURRENT_CONNECTION).Visible = blnVisible
    fraOptionFrame(OPTION_FRAME_GENERAL_DB_OPTIONS).Visible = blnVisible
    cmdSelectOtherDB.Visible = blnVisible

    chkSetIsConfirmedForDBSearchMatches.Enabled = blnVisible
    chkAddQuantitationDescriptionEntry.Enabled = blnVisible
    chkExportUMCsWithNoMatches.Enabled = blnVisible

    cboInternalStdSearchMode.Visible = blnVisible
    lblDescription(LABEL_DESCRIPTION_INTERNAL_STD_SEARCH_MODE).Visible = blnVisible

    If blnVisible Then
        strMessage = "Alternate output folder path (ignored during PRISM-initiated analysis)"
    Else
        fraOptionFrame(OPTION_FRAME_LEGACY_DB).Top = 600
        strMessage = "Alternate output folder path"
    End If
    
    lblDescription(LABEL_DESCRIPTION_ALTERNATE_OUTPUT_FOLDER_PATH).Caption = strMessage
    
    If APP_BUILD_DISABLE_LCMSWARP Then
        tbsNETOptions.TabEnabled(2) = False
        tbsNETOptions.TabCaption(2) = "Unused"
    End If

End Sub

Private Sub UpdateDBSearchModeTooltip()
    
    Dim strToolTip As String
    
    ' Update the ToolTip for cboDBSearchMode
    Select Case cboDBSearchMode.ListIndex
    Case dbsmExportUMCsOnly: strToolTip = "Exports all LC-MS Features to the database without actually searching for MT tag hits"
    Case dbsmIndividualPeaks: strToolTip = "Searches DB using individual peaks (ions); does not use LC-MS Features"
    Case dbsmIndividualPeaksInUMCsWithoutNET: strToolTip = "Searches DB using each peak in each LC-MS Feature; uses mass, but not NET"
    Case dbsmIndividualPeaksInUMCsWithNET: strToolTip = "This search mode is no longer supported"
    Case dbsmConglomerateUMCsWithNET: strToolTip = "Searches DB using the class mass for each LC-MS Feature (1 search per LC-MS Feature); uses class mass and NET of class representative"
    Case dbsmIndividualPeaksInUMCsPaired: strToolTip = "This search mode is no longer supported"
    Case dbsmIndividualPeaksInUMCsUnpaired: strToolTip = "This search mode is no longer supported"
    Case dbsmConglomerateUMCsPaired: strToolTip = "Searches DB using using the class mass for each paired LC-MS Feature, using either the light or heavy member (specify below)"
    Case dbsmConglomerateUMCsUnpaired: strToolTip = "Searches DB using using the class mass for each unpaired LC-MS Feature"
    Case dbsmConglomerateUMCsLightPairsPlusUnpaired: strToolTip = "Searches DB using the class mass of the light member of each paired LC-MS Feature, plus the masses of the unpaired LC-MS Features"
    Case dbsmPairsN14N15: strToolTip = "This search mode is no longer supported"
    Case dbsmPairsN14N15ConglomerateMass: strToolTip = "Searches DB using the class mass of each paired LC-MS Feature, light members only; uses mass and NET, plus considers number of N atoms"
    Case dbsmPairsICAT: strToolTip = "Searches DB using each peak in each paired LC-MS Feature, light members only; uses mass and NET, plus considers number of Cys residues"
    Case dbsmPairsPEO: strToolTip = "Searches DB using each peak in each paired LC-MS Feature, light members only; uses mass and NET, plus considers number of N atoms and number of Cys residues"
    Case Else
        ' This shouldn't happen
        Debug.Assert False
        strToolTip = ""
    End Select
    
    cboDBSearchMode.ToolTipText = strToolTip

End Sub

Private Sub UpdateDynamicControls()
    Dim eUMCSearchMode As usmUMCSearchModeConstants
    Dim intSearchModeIndex As Integer
    Dim intIndex As Integer
    Dim blnShowConstraints As Boolean
    Dim blnMassBasedDataDim As Boolean
    Dim eDBSearchMode As dbsmDatabaseSearchModeConstants
    
    Dim blnEMEnabled As Boolean
    
On Error GoTo UpdateDynamicControlsErrorHandler

    ' UMC Options
    eUMCSearchMode = cboUMCSearchMode.ListIndex
    
    Select Case eUMCSearchMode
    Case usmUMC2002
        fraUMCSearch200x.Visible = True
        fraUMCIonNetOptions.Visible = False
        fraUMC2002UniqueOptions.Visible = True
        fraUMC2003UniqueOptions.Visible = False
        chkUMCShrinkingBoxWeightAverageMassByIntensity.Visible = True
    Case usmUMC2003
        fraUMCSearch200x.Visible = True
        fraUMCIonNetOptions.Visible = False
        fraUMC2002UniqueOptions.Visible = False
        fraUMC2003UniqueOptions.Visible = True
        chkUMCShrinkingBoxWeightAverageMassByIntensity.Visible = False
    Case Else
        fraUMCSearch200x.Visible = False
        fraUMCIonNetOptions.Visible = True
    End Select
    
    ' Update the UMC auto refine length labels
    If mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange Then
        chkAutoRefineCheckbox(uarRemoveShort).Caption = "Remove cls. with less than"
        chkAutoRefineCheckbox(uarRemoveLong).Caption = "Remove cls. with length over"
        lblAutoRefineLengthLabel(0) = "scans"
        lblAutoRefineLengthLabel(1) = "scans"
        lblAutoRefineSpecialLabel(0).Enabled = True
    Else
        chkAutoRefineCheckbox(uarRemoveShort).Caption = "Remove cls. with less than"
        chkAutoRefineCheckbox(uarRemoveLong).Caption = "Remove cls. with more than"
        lblAutoRefineLengthLabel(0) = "members"
        lblAutoRefineLengthLabel(1) = "members"
        lblAutoRefineSpecialLabel(0).Enabled = False
    End If
    
    lblAutoRefineSpecialLabel(1).Enabled = lblAutoRefineSpecialLabel(0).Enabled
    txtAutoRefineTextbox(uarMinimumMemberCount).Enabled = lblAutoRefineSpecialLabel(0).Enabled
    txtAutoRefineTextbox(uarPercentMaxAbuToUseToGaugeLength).Enabled = lblAutoRefineSpecialLabel(0).Enabled
    
    ' UMC advanced class stats
    If CDblSafe(txtClassAbuTopXMinMaxAbu(0)) <= 0 And CDblSafe(txtClassAbuTopXMinMaxAbu(1)) <= 0 Then
        lblClassAbuTopXMinMembers = "Maximum members to include"
    Else
        lblClassAbuTopXMinMembers = "Minimum members to include"
    End If

    If CDblSafe(txtClassMassTopXMinMaxAbu(0)) <= 0 And CDblSafe(txtClassMassTopXMinMaxAbu(1)) <= 0 Then
        lblClassMassTopXMinMembers = "Maximum members to include"
    Else
        lblClassMassTopXMinMembers = "Minimum members to include"
    End If
    
    ' UMCIonNetDef Controls
    For intIndex = 0 To cmbData.Count - 1
        Select Case cmbData(intIndex).ListIndex
        Case uindUMCIonNetDimConstants.uindMonoMW, uindUMCIonNetDimConstants.uindAvgMW, uindUMCIonNetDimConstants.uindTmaMW
            cmbConstraintUnits(intIndex).Visible = True
            SetComboBox cmbConstraintUnits(intIndex), mCurrentSettings.UMCIonNetDef.MetricData(intIndex).ConstraintUnits, "UMC Metric Data Constraint Units"
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
    
''    ' Net Adjustment Options
''    ' chkUMCUseTopAbu should be checked if chkUMCUseTopAbu is >=0; otherwise, it should be unchecked
''    If cChkBox(chkUMCUseTopAbu) Then
''        If val(txtNETAlignmentOption(naoUMCAbuTopPct)) < 0 Then
''            txtNETAlignmentOption(naoUMCAbuTopPct) = Abs(val(txtNETAlignmentOption(naoUMCAbuTopPct)))
''            mCurrentSettings.UMCNetAdjDef.TopAbuPct = CLngSafe(txtNETAlignmentOption(naoUMCAbuTopPct))
''        End If
''        txtNETAlignmentOption(naoUMCAbuTopPct).Visible = True
''        lblUMCAbuTopPct.Visible = True
''    Else
''        If val(txtNETAlignmentOption(naoUMCAbuTopPct)) >= 0 Then
''            txtNETAlignmentOption(naoUMCAbuTopPct) = -Abs(val(txtNETAlignmentOption(naoUMCAbuTopPct)))
''            If val(txtNETAlignmentOption(naoUMCAbuTopPct)) = 0 Then txtNETAlignmentOption(naoUMCAbuTopPct) = "-90"
''            mCurrentSettings.UMCNetAdjDef.TopAbuPct = CLngSafe(txtNETAlignmentOption(naoUMCAbuTopPct))
''        End If
''        txtNETAlignmentOption(naoUMCAbuTopPct).Visible = False
''        lblUMCAbuTopPct.Visible = False
''    End If
    
''    ' Charge state for UMC Selection
''    chkCS(0).Enabled = Not cChkBox(chkCS(7))
''    For intIndex = 1 To 6
''        chkCS(intIndex).Enabled = chkCS(0).Enabled
''    Next intIndex
    
    ' Robust NET Options
    
    Dim blnUseRobustNET As Boolean
    
    Dim blnUseRobustNETIterative As Boolean
    Dim blnUseMSWarp As Boolean

    blnUseRobustNET = cChkBox(chkRobustNETEnabled)
    
    cboRobustNETAdjustmentMode.Enabled = blnUseRobustNET
    ''fraOptionFrame(OPTION_FRAME_NETSlopeRange).Enabled = blnUseRobustNET
    ''fraOptionFrame(OPTION_FRAME_NETInterceptRange).Enabled = blnUseRobustNET
    ''fraOptionFrame(OPTION_FRAME_MassShiftPPMRange).Enabled = blnUseRobustNET
    
    blnUseMSWarp = RobustNETWarpEnabled()
    
    blnUseRobustNETIterative = blnUseRobustNET And Not (blnUseMSWarp)
    
    cboMSWarpMassCalibrationType.Enabled = blnUseMSWarp
    
    fraOptionFrame(OPTION_FRAME_NET_WARP_OPTIONS).Enabled = blnUseMSWarp
    If mCurrentSettings.UMCNetAdjDef.RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass And _
       Not APP_BUILD_DISABLE_LCMSWARP Then
        fraOptionFrame(OPTION_FRAME_MASS_WARP_OPTIONS).Enabled = blnUseMSWarp
    Else
        fraOptionFrame(OPTION_FRAME_MASS_WARP_OPTIONS).Enabled = False
    End If
        
    ' Saving and Exporting
    If cChkBox(chkDoNotSaveOrExport) Then
        lblSavingAndExportingStatus.Caption = "Saving and exporting is currently disabled (see tab 7)."
    Else
        lblSavingAndExportingStatus.Caption = "Saving and exporting is enabled."
    End If
    
    ' Search Mode
    intSearchModeIndex = lstDBSearchModes.ListIndex
    If intSearchModeIndex >= 0 And intSearchModeIndex < mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intSearchModeIndex)
            chkDBSearchExportToDB.Enabled = Not APP_BUILD_DISABLE_MTS
            chkDBSearchWriteResultsToTextFile.Enabled = True
            txtDBSearchAlternateOutputFolderPath.Enabled = True
            
            eDBSearchMode = LookupDBSearchModeIndex(.SearchMode)
            Select Case eDBSearchMode
            Case dbsmPairsICAT, dbsmPairsPEO
                fraOptionFrame(OPTION_FRAME_DB_SEARCH_MT_LABELED).Visible = True
            Case Else
                fraOptionFrame(OPTION_FRAME_DB_SEARCH_MT_LABELED).Visible = False
            End Select
                        
            Select Case eDBSearchMode
            Case dbsmConglomerateUMCsWithNET, dbsmConglomerateUMCsPaired, dbsmConglomerateUMCsUnpaired, dbsmConglomerateUMCsLightPairsPlusUnpaired, dbsmExportUMCsOnly
                cboInternalStdSearchMode.Enabled = True
            Case Else
                cboInternalStdSearchMode.Enabled = False
            End Select
            
            SetCheckBox chkDBSearchExportToDB, .ExportResultsToDatabase
            SetCheckBox chkDBSearchWriteResultsToTextFile, .WriteResultsToTextFile
            txtDBSearchAlternateOutputFolderPath = .AlternateOutputFolderPath
            SetComboBox cboInternalStdSearchMode, .InternalStdSearchMode, "Internal standard search mode"
            
            txtDBSearchMinimumHighNormalizedScore = Trim(.DBSearchMinimumHighNormalizedScore)
            txtDBSearchMinimumHighDiscriminantScore = Trim(.DBSearchMinimumHighDiscriminantScore)
            txtDBSearchMinimumPeptideProphetProbability = Trim(.DBSearchMinimumPeptideProphetProbability)
            
            
            If .MassMods.ModMode = 2 Then
                optDBSearchModType(MODS_DECOY).Value = True
            ElseIf .MassMods.ModMode = 1 Then
                optDBSearchModType(MODS_DYNAMIC).Value = True
            Else
                ' Assumed fixed
                optDBSearchModType(MODS_FIXED).Value = True
            End If
                      
            If .MassMods.N15InsteadOfN14 Then
                optDBSearchNType(SEARCH_N15).Value = True
            Else
                optDBSearchNType(SEARCH_N14).Value = True
            End If
            
            If .PairSearchAssumeMassTagsAreLabeled Then
                optPairsDBSearchLabelAssumption(ASSUME_MT_LABELED).Value = True
            Else
                optPairsDBSearchLabelAssumption(ASSUME_MT_NOT_LABELED).Value = True
            End If
            
            With .MassMods
                SetCheckBox chkMTAdditionalMass(cbamPEO), .PEO
                SetCheckBox chkMTAdditionalMass(cbamICATd0), .ICATd0
                SetCheckBox chkMTAdditionalMass(cbamICATd8), .ICATd8
                SetCheckBox chkMTAdditionalMass(cbamAlkylation), .Alkylation
                txtAlkylationMWCorrection = Trim(.AlkylationMass)
                
                mUpdatingModMassControls = True
                If Len(.ResidueToModify) >= 1 Then
                    For intIndex = 0 To cboResidueToModify.ListCount - 1
                        If UCase(cboResidueToModify.List(intIndex)) = UCase(.ResidueToModify) Then
                            SetComboBox cboResidueToModify, intIndex, "Mass Mod Residue to modify"
                            Exit For
                        End If
                    Next intIndex
                Else
                    SetComboBox cboResidueToModify, 0, "Mass Mod Residue to modify"
                End If
                mUpdatingModMassControls = False
                txtResidueToModifyMass = Round(.ResidueMassModification, 5)
            
            End With
                        
        End With
    Else
        cboInternalStdSearchMode.Enabled = False
        chkDBSearchExportToDB.Enabled = False
        chkDBSearchWriteResultsToTextFile.Enabled = False
        txtDBSearchAlternateOutputFolderPath.Enabled = False
    End If
    
    ' Linked Filter Values
    VerifyLinkedFilterValues True, True
    
    ' Mass Calibration Warning Display
    If cChkBox(chkRefineDBSearchMassTolerance) Or cChkBox(chkRefineDBSearchNETTolerance) Then
        lblToleranceRefinementWarning.Visible = True
        
        If cChkBox(chkRefineDBSearchMassTolerance) And cChkBox(chkRefineDBSearchNETTolerance) Then
            ' Refine both Mass and NET tolerance are on
            lblToleranceRefinementWarning = "Note: Since Tolerance Refinement is enabled, the above mass and NET tolerances are used as the search tolerances only if tolerance refinement fails."
        ElseIf cChkBox(chkRefineDBSearchMassTolerance) Then
            ' Refine Mass Tolerance is on
            lblToleranceRefinementWarning = "Note: Since Tolerance Refinement is enabled, the above mass tolerance is used as the search tolerance only if tolerance refinement fails."
        Else
            ' Refine NET Tolerance is on
            lblToleranceRefinementWarning = "Note: Since Tolerance Refinement is enabled, the above NET tolerance is used as the search tolerance only if tolerance refinement fails."
        End If
    Else
        lblToleranceRefinementWarning.Visible = False
    End If
   
    ' Mass calibration tolerance type description
    If mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MassCalibrationTolType = gltPPM Then
        lblMassCalibrationRefinementDescription = "Note: All data is shifted by a constant ppm value, and thus a varying absolute Da value.  The ppm shift amount is determined by the location of the peak apex in a ppm-based mass-error plot."
        lblMassCalibrationRefinementUnits = "ppm"
    Else
        lblMassCalibrationRefinementDescription = "Note: All data is shifted linearly by a fixed amount, determined by the location of the peak apex in a Dalton-based mass-error plot."
        lblMassCalibrationRefinementUnits = "Da"
    End If
    
    ' Refinement options
    txtRefineDBSearchMassToleranceMinMax(0).Enabled = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance
    txtRefineDBSearchMassToleranceMinMax(1).Enabled = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance
    txtRefineDBSearchMassToleranceAdjustmentMultiplier.Enabled = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance
    
    txtRefineDBSearchNETToleranceMinMax(0).Enabled = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance
    txtRefineDBSearchNETToleranceMinMax(1).Enabled = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance
    txtRefineDBSearchNETToleranceAdjustmentMultiplier.Enabled = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance
    
    
    If mCurrentSettings.PrefsExpanded.RefineMSDataOptions.ToleranceRefinementMethod = mtrExpectationMaximization Then
        blnEMEnabled = True
    Else
        blnEMEnabled = False
    End If
    
    txtEMRefineMassErrorPeakToleranceEstimatePPM.Enabled = blnEMEnabled
    txtEMRefineNETErrorPeakToleranceEstimate.Enabled = blnEMEnabled
    txtEMRefinePercentOfDataToExclude.Enabled = blnEMEnabled
    chkEMRefineMassTolForceUseAllDataPointErrors.Enabled = blnEMEnabled
    chkEMRefineNETTolForceUseAllDataPointErrors.Enabled = blnEMEnabled
    
    ' Dynamic save options
    chkExtendedFileSaveModePreferred.Enabled = cChkBox(chkSaveGelFile)
    cboSavePictureGraphicFileType.Enabled = cChkBox(chkSavePictureGraphic)
    
    ' Pairs delta min/max options
    txtPairsMinMaxDelta(0).Enabled = Not cChkBox(chkPairsAutoMinMaxDelta)
    txtPairsMinMaxDelta(1).Enabled = txtPairsMinMaxDelta(0).Enabled
    
    ' Pairs ER calculation options
    With mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef
        chkPairEROptions(psoAverageERsAllChargeStates).Enabled = (.UseIdenticalChargesForER And .RequireMatchingChargeStatesForPairMembers)
        cboAverageERsWeightingMode.Enabled = (.AverageERsAllChargeStates And chkPairEROptions(psoAverageERsAllChargeStates).Enabled)
    
        chkPairEROptions(psoRemoveOutlierERs).Enabled = .ComputeERScanByScan
        chkPairEROptions(psoRemoveOutlierERsIterate).Enabled = .ComputeERScanByScan And .RemoveOutlierERs
        txtRemoveOutlierERsMinimumDataPointCount.Enabled = .ComputeERScanByScan And .RemoveOutlierERs
        
        chkPairEROptions(psoIReportEREnable).Enabled = .ComputeERScanByScan
    End With
    
    chkPairsExcludeAmbiguousKeepMostConfident.Enabled = cChkBox(chkPairsExcludeAmbiguous)
    
    Exit Sub
    
UpdateDynamicControlsErrorHandler:
    Debug.Print "Error in UpdateDynamicControls: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmEditAnalysisSettings->UpdateDynamicControls"
    Resume Next
        
End Sub

Private Sub UpdateUMCIonNetFindSingleMemberClasses()
    Dim blnFindSingleMemberClasses As Boolean
    
    If mUpdatingControls Then Exit Sub
        
    blnFindSingleMemberClasses = True
    If cChkBox(chkAutoRefineCheckbox(uarRemoveShort)) Then
        If IsNumeric(txtAutoRefineTextbox(uarRemoveShort)) Then
            If val(txtAutoRefineTextbox(uarRemoveShort)) >= 2 Then
                blnFindSingleMemberClasses = False
            End If
        End If
    End If

    SetCheckBox chkUMCIonNetMakeSingleMemberClasses, blnFindSingleMemberClasses
End Sub

Private Sub ValidateChangedInitialNETTol()
    If RobustNETWarpEnabled() Then
        ValidateTextboxValueDbl txtNetAdjInitialNETTol, 0.0001, 10, 0.02
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.NETTol = CDblSafe(txtNetAdjInitialNETTol)
    Else
        ValidateTextboxValueDbl txtNetAdjInitialNETTol, 0.0001, 100, 0.2
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol = CDblSafe(txtNetAdjInitialNETTol)
    End If
    
    ' Reset mAutoIniFileNameOverridden
    mAutoIniFileNameOverridden = False
End Sub

Private Sub ValidateNETWarpMassMaxJump()
    Dim intMaxValue As Integer
    
    intMaxValue = CIntSafe(txtRobustNETOption(rnoWarpNumMassDeltaBins).Text)
    If intMaxValue < 2 Then intMaxValue = 2
    
    mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MassMaxJump = ValidateTextboxValueLng(txtRobustNETOption(rnoWarpMassMaxJump), 0, CLng(intMaxValue), intMaxValue / 2)
End Sub

Private Sub ValidateRobustNETAdjustmentModeSettings()
    mCurrentSettings.UMCNetAdjDef.RobustNETAdjustmentMode = cboRobustNETAdjustmentMode.ListIndex
    mCurrentSettings.UMCNetAdjDef.UseRobustNETAdjustment = cChkBox(chkRobustNETEnabled)
    UpdateDynamicControls
    
    If Not mUpdatingControls Then
        If RobustNETWarpEnabled() Then
            txtNetAdjInitialNETTol.Text = mCurrentSettings.UMCNetAdjDef.MSWarpOptions.NETTol
        Else
            txtNetAdjInitialNETTol.Text = mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol
        End If
        
        ' Reset mAutoIniFileNameOverridden
        mAutoIniFileNameOverridden = False
    End If
End Sub

Private Function ValidatePairsIdentificationMode() As Boolean
    ' Returns True if the pairs identification mode is valid; otherwise, returns false
    
    Dim intIndex As Integer
    Dim intMatchIndex As Integer
    
    Dim blnPairFindingEnabled As Boolean
    Dim blnPairDBSearchingEnabled As Boolean
    Dim blnInvalidPairSearchCombo As Boolean
    
    Dim strMessage As String
    Dim eResponse As VbMsgBoxResult
    
    Static blnWorking As Boolean
    
On Error GoTo ValidatePairsIdentificationModeErrorHandler

    If blnWorking Then Exit Function
    blnWorking = True
    
    If cboPairsIdentificationMode.ListIndex = 0 Then
        blnPairFindingEnabled = False
    Else
        blnPairFindingEnabled = True
    End If
    
    blnPairDBSearchingEnabled = False
    With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions
        If .AutoAnalysisSearchModeCount > 0 Then
            ' See if any of the AutoAnalysisSearchModes is a pairs-based search
            For intIndex = 0 To .AutoAnalysisSearchModeCount - 1
                If LookupDBSearchModeIndex(.AutoAnalysisSearchMode(intIndex).SearchMode) >= DB_SEARCH_MODE_PAIR_MODE_START_INDEX Then
                    blnPairDBSearchingEnabled = True
                    Exit For
                End If
            Next intIndex
        End If
    End With
    
    If Not blnPairFindingEnabled Then
        If blnPairDBSearchingEnabled Then
            If mUpdatingControls Then
                eResponse = vbYes
            Else
                strMessage = "Warning: Turning off pairs identification will automatically remove any pairs-based database searching.  Do you really want to do this?"
                eResponse = MsgBox(strMessage, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Remove Pairs-based DB Searching")
            End If
            
            If eResponse <> vbYes Then
                If mCurrentSettings.PrefsExpanded.PairSearchOptions.PairSearchMode = AUTO_FIND_PAIRS_LABEL Then
                    SetComboBox cboPairsIdentificationMode, pimLabel, "Pairs identification mode"
                Else
                    SetComboBox cboPairsIdentificationMode, pimDelta, "Pairs identification mode"
                End If
                blnPairFindingEnabled = True
            Else
                blnPairFindingEnabled = False
            End If
            
            If Not blnPairFindingEnabled Then
                With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions
                    ' Remove any pairs-based searches
                    intIndex = 0
                    Do While intIndex < .AutoAnalysisSearchModeCount
                        If LookupDBSearchModeIndex(.AutoAnalysisSearchMode(intIndex).SearchMode) >= DB_SEARCH_MODE_PAIR_MODE_START_INDEX Then
                            If lstDBSearchModes.ListCount > 0 Then
                                lstDBSearchModes.ListIndex = intIndex
                                RemoveSelectedSearchMode
                            Else
                                intIndex = intIndex + 1
                            End If
                        Else
                            intIndex = intIndex + 1
                        End If
                    Loop
                    blnPairDBSearchingEnabled = False
                End With
            End If
        End If
    End If
    
    blnInvalidPairSearchCombo = False
    If blnPairFindingEnabled Then
        If mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsAfterDBSearch Then
            ' Make sure blnPairDBSearchingEnabled = True and that another search mode is present after the pair-based DB search
                        
            strMessage = "You have enabled the 'Pair member removal after database search' option, but have not defined a pairs-based database search mode.  Please do so, and make sure a non pairs-based search mode is present following the pairs-based search mode.  Otherwise, please disable the 'Pair member removal after database search' option."
            If blnPairDBSearchingEnabled Then
                With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions
                    If .AutoAnalysisSearchModeCount > 0 Then
                        ' Find the last pairs-based search mode
                        intMatchIndex = -1
                        For intIndex = .AutoAnalysisSearchModeCount - 1 To 0 Step -1
                            If LookupDBSearchModeIndex(.AutoAnalysisSearchMode(intIndex).SearchMode) >= DB_SEARCH_MODE_PAIR_MODE_START_INDEX Then
                                ' Match found
                                intMatchIndex = intIndex
                                Exit For
                            End If
                        Next intIndex
                        
                        If intMatchIndex < 0 Or intIndex = .AutoAnalysisSearchModeCount - 1 Then
                            ' No non pairs-based search modes after the pairs-based search mode
                            strMessage = "You have enabled the 'Pair member removal after database search' option, but do not have any non pairs-based database search modes defined after the pairs-based database search.  Please define a non pairs-based search mode and make sure it occurs after the pairs-based search.  Otherwise, please disable the 'Pair member removal after database search' option."
                            MsgBox strMessage, vbExclamation + vbOKOnly, "Invalid DB Search Mode"
                            blnInvalidPairSearchCombo = True
                        End If
                    Else
                        MsgBox strMessage, vbExclamation + vbOKOnly, "Invalid DB Search Mode"
                        blnInvalidPairSearchCombo = True
                    End If
                End With
                
            Else
                MsgBox strMessage, vbExclamation + vbOKOnly, "Invalid DB Search Mode"
                blnInvalidPairSearchCombo = True
            End If
        End If
    End If
    
    With mCurrentSettings.PrefsExpanded.PairSearchOptions
        Select Case cboPairsIdentificationMode.ListIndex
        Case pimNone:
            .PairSearchMode = AUTO_FIND_PAIRS_NONE
        Case pimDelta:
            .PairSearchMode = AUTO_FIND_PAIRS_DELTA
            fraOptionFrame(OPTION_FRAME_DELTA_PAIRS).Visible = True
            fraOptionFrame(OPTION_FRAME_LABEL_PAIRS).Visible = False
        Case pimLabel:
            .PairSearchMode = AUTO_FIND_PAIRS_LABEL
            fraOptionFrame(OPTION_FRAME_DELTA_PAIRS).Visible = False
            fraOptionFrame(OPTION_FRAME_LABEL_PAIRS).Visible = True
        Case Else
            Debug.Assert False
        End Select
    End With
    
    blnWorking = False
    
    ValidatePairsIdentificationMode = Not blnInvalidPairSearchCombo
    Exit Function

ValidatePairsIdentificationModeErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmEditAnalysisSettings.ValidatePairsIdentificationMode"
    blnWorking = False
    ValidatePairsIdentificationMode = Not blnInvalidPairSearchCombo
End Function

Private Sub VerifyLinkedFilterValues(Optional blnFavorIsoValues As Boolean = True, Optional blnFavorMinValue As Boolean = True)
    
    ' The static variable blnUpdating is used to prevent circular recursive calling
    Static blnUpdating As Boolean
    
    If blnUpdating Then Exit Sub
    
    blnUpdating = True
    
    If cChkBox(chkCSIsoSameRangeAbu) Then
        If blnFavorIsoValues Then
            txtCSMinMaxAbu(0) = txtIsoMinMaxAbu(0)
            txtCSMinMaxAbu(1) = txtIsoMinMaxAbu(1)
            SetCheckBox chkCSAbuRange, cChkBox(chkIsoAbuRange.Value)
        Else
            txtIsoMinMaxAbu(0) = txtCSMinMaxAbu(0)
            txtIsoMinMaxAbu(1) = txtCSMinMaxAbu(1)
            SetCheckBox chkIsoAbuRange, cChkBox(chkCSAbuRange.Value)
        End If
    End If
    
    If cChkBox(chkCSIsoSameRangeMW) Then
        If blnFavorIsoValues Then
            txtCSMinMaxMW(0) = txtIsoMinMaxMW(0)
            txtCSMinMaxMW(1) = txtIsoMinMaxMW(1)
            SetCheckBox chkCSMWRange, cChkBox(chkIsoMWRange.Value)
        Else
            txtIsoMinMaxMW(0) = txtCSMinMaxMW(0)
            txtIsoMinMaxMW(1) = txtCSMinMaxMW(1)
            SetCheckBox chkIsoMWRange, cChkBox(chkCSMWRange.Value)
        End If
    End If
    
    If Not mUpdatingControls Then
        ValidateDualTextBoxes txtIsoMinMaxAbu(0), txtIsoMinMaxAbu(1), blnFavorMinValue, 0, 1E+300, 0
        txtIsoMinMaxAbu_LostFocus 0
        txtIsoMinMaxAbu_LostFocus 1
        
        ValidateDualTextBoxes txtCSMinMaxAbu(0), txtCSMinMaxAbu(1), blnFavorMinValue, 0, 1E+300, 0
        txtCSMinMaxAbu_LostFocus 0
        txtCSMinMaxAbu_LostFocus 1
        
        ValidateDualTextBoxes txtIsoMinMaxMW(0), txtIsoMinMaxMW(1), blnFavorMinValue, 0, MAX_MW, 0
        txtIsoMinMaxMW_LostFocus 0
        txtIsoMinMaxMW_LostFocus 1
        
        ValidateDualTextBoxes txtIsoMinMaxMZ(0), txtIsoMinMaxMZ(1), blnFavorMinValue, 0, MAX_MZ, 0
        txtIsoMinMaxMZ_LostFocus 0
        txtIsoMinMaxMZ_LostFocus 1
        
        ValidateDualTextBoxes txtCSMinMaxMW(0), txtCSMinMaxMW(1), blnFavorMinValue, 0, MAX_MW, 0
        txtCSMinMaxMW_LostFocus 0
        txtCSMinMaxMW_LostFocus 1
    End If
    
    blnUpdating = False
    
End Sub

Private Sub cboAverageERsWeightingMode_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.AverageERsWeightingMode = cboAverageERsWeightingMode.ListIndex
End Sub

Private Sub cboChargeStateAbuType_Click()
mCurrentSettings.UMCDef.ChargeStateStatsRepType = cboChargeStateAbuType.ListIndex
End Sub

Private Sub cboDBSearchMode_Click()
    UpdateDBSearchModeTooltip
End Sub

Private Sub cboEvenOddScanNumber_Click()
    With mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs
        Select Case cboEvenOddScanNumber.ListIndex
        Case eosEvenScans
            .RestrictToEvenScanNumbersOnly = True
            .RestrictToOddScanNumbersOnly = False
        Case eosOddScans
            .RestrictToEvenScanNumbersOnly = False
            .RestrictToOddScanNumbersOnly = True
        Case Else
            .RestrictToEvenScanNumbersOnly = False
            .RestrictToOddScanNumbersOnly = False
        End Select
    End With
End Sub

Private Sub cboInternalStdSearchMode_Click()
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If cboInternalStdSearchMode.ListIndex < 0 Then
                If APP_BUILD_DISABLE_MTS Then
                    .InternalStdSearchMode = issmFindOnlyMassTags
                Else
                    .InternalStdSearchMode = issmFindWithMassTags
                End If
            Else
                .InternalStdSearchMode = cboInternalStdSearchMode.ListIndex
            End If
        End With
    End If
End Sub

Private Sub cboGelFileInMemoryToUse_Click()
    mActiveGelIndex = cboGelFileInMemoryToUse.ListIndex + 1
End Sub

Private Sub cboToleranceRefinementMethod_Click()
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.ToleranceRefinementMethod = cboToleranceRefinementMethod.ListIndex
    UpdateDynamicControls
End Sub

Private Sub cboMSWarpMassCalibrationType_Click()
    mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MassCalibrationType = cboMSWarpMassCalibrationType.ListIndex
End Sub

Private Sub cboPairsIdentificationMode_Click()
    ValidatePairsIdentificationMode
End Sub

Private Sub cboPairsUMCsToUseForNETAdjustment_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection = cboPairsUMCsToUseForNETAdjustment.ListIndex
End Sub

Private Sub cboResidueToModify_Click()
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If cboResidueToModify.ListIndex > 0 Then
                .MassMods.ResidueToModify = cboResidueToModify
            Else
                .MassMods.ResidueToModify = ""
            End If
            
            If Not mUpdatingModMassControls Then
                If .MassMods.ResidueToModify = glPHOSPHORYLATION Then
                    txtResidueToModifyMass = Trim(glPHOSPHORYLATION_Mass)
                    .MassMods.ResidueMassModification = glPHOSPHORYLATION_Mass
                Else
                    ' For safety reasons, reset txtResidueToModifyMass to "0"
                    txtResidueToModifyMass = "0"
                    .MassMods.ResidueMassModification = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub cboRobustNETAdjustmentMode_Click()
    ValidateRobustNETAdjustmentModeSettings
End Sub

Private Sub cboSavePictureGraphicFileType_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SavePictureGraphicFileType = cboSavePictureGraphicFileType.ListIndex + 1
End Sub

Private Sub cboSearchRegionShape_Click(Index As Integer)
    Dim eSearchRegionShape As srsSearchRegionShapeConstants
    
    If cboSearchRegionShape(Index).ListIndex = srsSearchRegionShapeConstants.srsRectangular Then
        eSearchRegionShape = srsSearchRegionShapeConstants.srsRectangular
    Else
        eSearchRegionShape = srsSearchRegionShapeConstants.srsElliptical
    End If
    
    Select Case Index
    Case 0
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchRegionShape = eSearchRegionShape
    Case 1
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.DBSearchRegionShape = eSearchRegionShape
    Case Else
        ' Unknown index
        Debug.Assert False
    End Select
End Sub

Private Sub cboSplitUMCsScanGapBehavior_Click()
    mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.SplitUMCOptions.ScanGapBehavior = cboSplitUMCsScanGapBehavior.ListIndex
End Sub

Private Sub cboUMCSearchMode_Click()
    
On Error GoTo UMCSearchModeClickErrorHandler

    Select Case cboUMCSearchMode.ListIndex
    Case usmUMC2002
        ' Now Obsolete
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.UMCSearchMode = AUTO_ANALYSIS_UMCListType2002
        SetComboBox cboUMCSearchMode, usmUMC2002, "UMC Search Mode"
    Case usmUMC2003
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.UMCSearchMode = AUTO_ANALYSIS_UMC2003
        SetComboBox cboUMCSearchMode, usmUMC2003, "UMC Search Mode"
    Case Else
        ' Includes usmUMCIonNet
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.UMCSearchMode = AUTO_ANALYSIS_UMCIonNet
        SetComboBox cboUMCSearchMode, usmUMCIonNet, "UMC Search Mode"
    End Select
    
    UpdateDynamicControls
    
    ' Reset mAutoIniFileNameOverridden
    mAutoIniFileNameOverridden = False
    
    Exit Sub

UMCSearchModeClickErrorHandler:
    Debug.Print "Error in cboUMCSearchMode_Click: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmEditAnalysisSettings->cboUMCSearchMode_Click"
    Resume Next
End Sub

Private Sub chkAddQuantitationDescriptionEntry_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AddQuantitationDescriptionEntry = cChkBox(chkAddQuantitationDescriptionEntry)
End Sub

Private Sub chkAllowSharing_Click()
    mCurrentSettings.UMCDef.UMCSharing = cChkBox(chkAllowSharing)
End Sub

Private Sub chkAutoRefineCheckbox_Click(Index As Integer)
    
    Select Case Index
    Case uarRemoveLowIntensity
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveAbundanceLow = cChkBox(chkAutoRefineCheckbox(Index))
    Case uarRemoveHighIntensity
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveAbundanceHigh = cChkBox(chkAutoRefineCheckbox(Index))
    Case uarRemoveShort
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveCountLow = cChkBox(chkAutoRefineCheckbox(Index))
        UpdateUMCIonNetFindSingleMemberClasses
    Case uarRemoveLong
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveCountHigh = cChkBox(chkAutoRefineCheckbox(Index))
    Case uarTestLengthUsingScanRange
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange = cChkBox(chkAutoRefineCheckbox(Index))
        UpdateDynamicControls
    Case uarRemoveMaxLengthPctAllScans
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveMaxLengthPctAllScans = cChkBox(chkAutoRefineCheckbox(Index))
    Case Else
        Debug.Assert False
    End Select
    
    UpdateDynamicControls
    
End Sub

Private Sub chkCSAbuRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictCSByAbundance = cChkBox(chkCSAbuRange)
    If cChkBox(chkCSIsoSameRangeAbu) Then SetCheckBox chkIsoAbuRange, cChkBox(chkCSAbuRange)
End Sub

Private Sub chkCSIsoSameRangeAbu_Click()
    VerifyLinkedFilterValues True, True
End Sub

Private Sub chkCSIsoSameRangeMW_Click()
    VerifyLinkedFilterValues True, True
End Sub

Private Sub chkCSMWRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictCSByMass = cChkBox(chkCSMWRange)
    If cChkBox(chkCSIsoSameRangeMW) Then SetCheckBox chkIsoMWRange, cChkBox(chkCSMWRange)
End Sub

Private Sub chkCurrentDBAMTsOnly_Click()
    SetCheckBox chkCurrentDBAMTsOnly, mCurrentSettings.DBSettings.AMTsOnly
End Sub

Private Sub chkCurrentDBConfirmedOnly_Click()
    SetCheckBox chkCurrentDBConfirmedOnly, mCurrentSettings.DBSettings.ConfirmedOnly
End Sub

Private Sub chkCurrentDBLimitToPMTsFromDataset_Click()
    SetCheckBox chkCurrentDBLimitToPMTsFromDataset, mCurrentSettings.DBSettings.LimitToPMTsFromDataset
End Sub

Private Sub chkCurrentDBLockersOnly_Click()
    SetCheckBox chkCurrentDBLockersOnly, mCurrentSettings.DBSettings.LockersOnly
End Sub

Private Sub chkDBSearchExportToDB_Click()
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            .ExportResultsToDatabase = cChkBox(chkDBSearchExportToDB)
        End With
    End If
End Sub

Private Sub chkDBSearchIncludeORFNameInTextFileOutput_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput = cChkBox(chkDBSearchIncludeORFNameInTextFileOutput)
End Sub

Private Sub chkDBSearchWriteIDResultsByIonAfterAutoSearches_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.WriteIDResultsByIonToTextFileAfterAutoSearches = cChkBox(chkDBSearchWriteIDResultsByIonAfterAutoSearches)
End Sub

Private Sub chkDBSearchWriteResultsToTextFile_Click()
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            .WriteResultsToTextFile = cChkBox(chkDBSearchWriteResultsToTextFile)
        End With
    End If
End Sub

Private Sub chkDoNotSaveOrExport_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.DoNotSaveOrExport = cChkBox(chkDoNotSaveOrExport)
    UpdateDynamicControls
End Sub

Private Sub chkDupElimination_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeDuplicates = cChkBox(chkDupElimination)
End Sub

Private Sub chkEMRefineMassTolForceUseAllDataPointErrors_Click()
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.EMMassTolRefineForceUseSingleDataPointErrors = cChkBox(chkEMRefineMassTolForceUseAllDataPointErrors)
End Sub

Private Sub chkEMRefineNETTolForceUseAllDataPointErrors_Click()
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.EMNETTolRefineForceUseSingleDataPointErrors = cChkBox(chkEMRefineNETTolForceUseAllDataPointErrors)
End Sub

Private Sub chkExcludeIsoByFit_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeIsoByFit = cChkBox(chkExcludeIsoByFit)
End Sub

Private Sub chkExportResultsFileUsesJobNumber_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.ExportResultsFileUsesJobNumberInsteadOfDataSetName = cChkBox(chkExportResultsFileUsesJobNumber)
End Sub

Private Sub chkExportUMCsWithNoMatches_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches = cChkBox(chkExportUMCsWithNoMatches)
End Sub

Private Sub chkExtendedFileSaveModePreferred_Click()
    mCurrentSettings.PrefsExpanded.ExtendedFileSaveModePreferred = cChkBox(chkExtendedFileSaveModePreferred)
End Sub

Private Sub chkInterpolateMissingIons_Click()
    mCurrentSettings.UMCDef.InterpolateGaps = cChkBox(chkInterpolateMissingIons)
End Sub

Private Sub chkIsoAbuRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoByAbundance = cChkBox(chkIsoAbuRange)
    If cChkBox(chkCSIsoSameRangeAbu) Then SetCheckBox chkCSAbuRange, cChkBox(chkIsoAbuRange)
End Sub

Private Sub chkIsoMWRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoByMass = cChkBox(chkIsoMWRange)
    If cChkBox(chkCSIsoSameRangeMW) Then SetCheckBox chkCSMWRange, cChkBox(chkIsoMWRange)
End Sub

Private Sub chkIsoMZRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoByMZ = cChkBox(chkIsoMZRange)
End Sub

Private Sub chkIsoUseCSRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoByChargeState = cChkBox(chkIsoUseCSRange)
End Sub

Private Sub chkMaximumDataCountEnabled_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.MaximumDataCountEnabled = cChkBox(chkMaximumDataCountEnabled)
End Sub

Private Sub chkMTAdditionalMass_Click(Index As Integer)
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            Select Case Index
            Case cbamPEO
                ' PEO
                .MassMods.PEO = cChkBox(chkMTAdditionalMass(Index))
            Case cbamICATd0
                ' ICAT d0
                .MassMods.ICATd0 = cChkBox(chkMTAdditionalMass(Index))
            Case cbamICATd8
                ' ICAT d8
                .MassMods.ICATd8 = cChkBox(chkMTAdditionalMass(Index))
            Case cbamAlkylation
                ' Alkylation
                .MassMods.Alkylation = cChkBox(chkMTAdditionalMass(Index))
                If .MassMods.Alkylation Then
                    txtAlkylationMWCorrection = Trim(glALKYLATION)
                Else
                    txtAlkylationMWCorrection = Trim(0)
                End If
            End Select
        End With
    End If
    
End Sub

' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''Private Sub chkNetAdjUseLockers_Click()
''    mCurrentSettings.UMCNetAdjDef.UseNetAdjLockers = cChkBox(chkNetAdjUseLockers)
''End Sub
''
''Private Sub chkNetAdjUseOldIfFailure_Click()
''    mCurrentSettings.UMCNetAdjDef.UseOldNetAdjIfFailure = cChkBox(chkNetAdjUseOldIfFailure)
''End Sub

Private Sub chkPairEROptions_Click(Index As Integer)

    Dim blnValue As Boolean
    blnValue = cChkBox(chkPairEROptions(Index))

    With mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef
        Select Case Index
        Case psoRequireMatchingChargeStatesForPairMembers
            .RequireMatchingChargeStatesForPairMembers = blnValue
        Case psoUseIdenticalChargesForER
            .UseIdenticalChargesForER = blnValue
        Case psoAverageERsAllChargeStates
            .AverageERsAllChargeStates = blnValue
        Case psoComputeERScanByScan
            .ComputeERScanByScan = blnValue
        Case psoIReportEREnable
            .IReportEROptions.Enabled = blnValue
        Case psoRemoveOutlierERs
            .RemoveOutlierERs = blnValue
        Case psoRemoveOutlierERsIterate
            .RemoveOutlierERsIterate = blnValue
        Case Else
            Debug.Assert False
        End Select
    End With
    UpdateDynamicControls
    
End Sub

Private Sub chkPairsAutoMinMaxDelta_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.AutoCalculateDeltaMinMaxCount = cChkBox(chkPairsAutoMinMaxDelta)
    UpdateDynamicControls
End Sub

Private Sub chkPairsExcludeAmbiguous_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoExcludeAmbiguous = cChkBox(chkPairsExcludeAmbiguous)
    UpdateDynamicControls
End Sub

Private Sub chkPairsExcludeAmbiguousKeepMostConfident_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.KeepMostConfidentAmbiguous = cChkBox(chkPairsExcludeAmbiguousKeepMostConfident)
End Sub

Private Sub chkPairsExcludeOutOfERRange_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoExcludeOutOfERRange = cChkBox(chkPairsExcludeOutOfERRange)
End Sub

Private Sub chkPairsRequireOverlapAtApex_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.RequireUMCOverlapAtApex = cChkBox(chkPairsRequireOverlapAtApex)
End Sub

Private Sub chkPairsRequireOverlapAtEdge_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.RequireUMCOverlap = cChkBox(chkPairsRequireOverlapAtEdge)
End Sub

Private Sub chkPairsSaveStatisticsTextFile_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoAnalysisSavePairsStatisticsToTextFile = cChkBox(chkPairsSaveStatisticsTextFile)
End Sub

Private Sub chkPairsSaveTextFile_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoAnalysisSavePairsToTextFile = cChkBox(chkPairsSaveTextFile)
End Sub

Private Sub chkRefineDBSearchIncludeInternalStds_Click()
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.IncludeInternalStdMatches = cChkBox(chkRefineDBSearchIncludeInternalStds)
End Sub

Private Sub chkRefineDBSearchMassTolerance_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance = cChkBox(chkRefineDBSearchMassTolerance)
    DisplayDataValidationWarnings

    ' Reset mAutoIniFileNameOverridden
    mAutoIniFileNameOverridden = False
    UpdateDynamicControls
End Sub

Private Sub chkRefineDBSearchNETTolerance_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance = cChkBox(chkRefineDBSearchNETTolerance)

    ' Reset mAutoIniFileNameOverridden
    mAutoIniFileNameOverridden = False
    UpdateDynamicControls
End Sub

Private Sub chkRefineDBSearchTolUseMinMaxIfOutOfRange_Click()
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.UseMinMaxIfOutOfRange = cChkBox(chkRefineDBSearchTolUseMinMaxIfOutOfRange)
End Sub

Private Sub chkRefineMassCalibration_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibration = cChkBox(chkRefineMassCalibration)
    DisplayDataValidationWarnings

    ' Reset mAutoIniFileNameOverridden
    mAutoIniFileNameOverridden = False
End Sub

Private Sub chkRemovePairMemberHitsAfterDBSearch_Click()
    mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsAfterDBSearch = cChkBox(chkRemovePairMemberHitsAfterDBSearch)
End Sub

Private Sub chkRestrictGANETRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictGANETRange = cChkBox(chkRestrictGANETRange)
End Sub

Private Sub chkRestrictScanRange_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictScanRange = cChkBox(chkRestrictScanRange)
End Sub

Private Sub chkRobustNETEnabled_Click()
    ValidateRobustNETAdjustmentModeSettings
    AutoShowHideWarpingTab
End Sub

Private Sub chkSaveErrorGraphic3D_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SaveErrorGraphic3D = cChkBox(chkSaveErrorGraphic3D)
End Sub

Private Sub chkSaveErrorGraphicGANET_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SaveErrorGraphicGANET = cChkBox(chkSaveErrorGraphicGANET)
End Sub

Private Sub chkSaveErrorGraphicMass_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SaveErrorGraphicMass = cChkBox(chkSaveErrorGraphicMass)
End Sub

Private Sub chkSaveGelFile_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SaveGelFile = cChkBox(chkSaveGelFile)
    UpdateDynamicControls
End Sub

Private Sub chkSaveGelFileOnError_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SaveGelFileOnError = cChkBox(chkSaveGelFileOnError)
    UpdateDynamicControls
End Sub

Private Sub chkSavePictureGraphic_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SavePictureGraphic = cChkBox(chkSavePictureGraphic)
    UpdateDynamicControls
End Sub

Private Sub chkSaveUMCStatisticsToTextFile_Click()
    With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions
        .SaveUMCStatisticsToTextFile = cChkBox(chkSaveUMCStatisticsToTextFile)
    End With
End Sub

Private Sub chkSecGuessElimination_Click(Index As Integer)
    With mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs
        .ExcludeIsoSecondGuess = cChkBox(chkSecGuessElimination(0))
        .ExcludeIsoLessLikelyGuess = cChkBox(chkSecGuessElimination(1))
    End With
End Sub

Private Sub chkSetIsConfirmedForDBSearchMatches_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SetIsConfirmedForDBSearchMatches = cChkBox(chkSetIsConfirmedForDBSearchMatches)
End Sub

Private Sub chkSkipGANETSlopeAndInterceptComputation_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SkipGANETSlopeAndInterceptComputation = cChkBox(chkSkipGANETSlopeAndInterceptComputation)
End Sub

Private Sub chkSkipFindUMCs_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SkipFindUMCs = cChkBox(chkSkipFindUMCs)
End Sub

Private Sub chkSplitUMCsByExaminingAbundance_Click()
    mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.SplitUMCsByAbundance = cChkBox(chkSplitUMCsByExaminingAbundance)
End Sub

Private Sub chkUMCDefRequireIdenticalCharge_Click()
    mCurrentSettings.UMCDef.UMCUniCS = cChkBox(chkUMCDefRequireIdenticalCharge)
End Sub

Private Sub chkUMCShrinkingBoxWeightAverageMassByIntensity_Click()
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.UMCShrinkingBoxWeightAverageMassByIntensity = cChkBox(chkUMCShrinkingBoxWeightAverageMassByIntensity)
End Sub

Private Sub chkUse_Click(Index As Integer)
    mCurrentSettings.UMCIonNetDef.MetricData(Index).Use = cChkBox(chkUse(Index))
End Sub

Private Sub chkUseLegacyDBForMTs_Click()
    If APP_BUILD_DISABLE_MTS Then
        If chkUseLegacyDBForMTs.Value = vbUnchecked Then
            chkUseLegacyDBForMTs.Value = vbChecked
        End If
    End If
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.UseLegacyDBForMTs = cChkBox(chkUseLegacyDBForMTs)
End Sub

Private Sub chkUseMassTagsWithNullNET_Click()
    mCurrentSettings.PrefsExpanded.UseMassTagsWithNullNET = cChkBox(chkUseMassTagsWithNullNET)
End Sub

Private Sub chkUMCIonNetMakeSingleMemberClasses_Click()
    mCurrentSettings.PrefsExpanded.UMCIonNetOptions.MakeSingleMemberClasses = cChkBox(chkUMCIonNetMakeSingleMemberClasses)
End Sub

Private Sub chkUseMostAbuChargeStateStatsForClassStats_Click()
    mCurrentSettings.UMCDef.UMCClassStatsUseStatsFromMostAbuChargeState = cChkBox(chkUseMostAbuChargeStateStatsForClassStats)
End Sub

Private Sub chkUseUMCClassStats_Click()
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.UseUMCClassStats = cChkBox(chkUseUMCClassStats)
End Sub

Private Sub chkUseUMCConglomerateNET_Click()
    mCurrentSettings.PrefsExpanded.UseUMCConglomerateNET = cChkBox(chkUseUMCConglomerateNET)
End Sub

Private Sub cmbConstraint_Click(Index As Integer)
    mCurrentSettings.UMCIonNetDef.MetricData(Index).ConstraintType = cmbConstraint(Index).ListIndex
    UpdateDynamicControls
End Sub

Private Sub cmbConstraintUnits_Click(Index As Integer)
    mCurrentSettings.UMCIonNetDef.MetricData(Index).ConstraintUnits = cmbConstraintUnits(Index).ListIndex
    
    If Not mUpdatingControls Then
        If cmbConstraintUnits(Index).ListIndex = DATA_UNITS_MASS_DA Then
            ' Convert the constraint tolerance from ppm to Da, assuming 1000 m/z
            txtConstraint(Index) = PPMToMass(txtConstraint(Index), 1000)
        Else
            ' Convert the constraint tolerance from Da to ppm, assuming 1000 m/z
            txtConstraint(Index) = MassToPPM(txtConstraint(Index), 1000)
        End If
        
        If IsNumeric(txtConstraint(Index).Text) Then
            mCurrentSettings.UMCIonNetDef.MetricData(Index).ConstraintValue = CDbl(txtConstraint(Index).Text)
        End If
    End If

End Sub

Private Sub cmbCountType_Click()
    mCurrentSettings.UMCDef.UMCType = cmbCountType.ListIndex
End Sub

Private Sub cmbData_Click(Index As Integer)
    mCurrentSettings.UMCIonNetDef.MetricData(Index).DataType = cmbData(Index).ListIndex
    UpdateDynamicControls
End Sub

Private Sub cmbMetricType_Click()
    mCurrentSettings.UMCIonNetDef.MetricType = cmbMetricType.ListIndex
End Sub

Private Sub cmbUMCAbu_Click()
    mCurrentSettings.UMCDef.ClassAbu = cmbUMCAbu.ListIndex
End Sub

Private Sub cmbUMCMW_Click()
    mCurrentSettings.UMCDef.ClassMW = cmbUMCMW.ListIndex
End Sub

Private Sub cmbUMCRepresentative_Click()
    mCurrentSettings.PrefsExpanded.UMCIonNetOptions.UMCRepresentative = cmbUMCRepresentative.ListIndex
End Sub

Private Sub cmdAddRemove_Click(Index As Integer)
    Select Case Index
    Case 0
        ' Remove
        RemoveSelectedSearchMode
    Case 1
        ' Add
        AddNewDBSearchMode
    Case 2
        ' Remove all
        lstDBSearchModes.Clear
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount = 0
        UpdateDynamicControls
    End Select
    
End Sub

Private Sub cmdApplyToSelectedGel_Click()
    ApplyToSelectedGel
End Sub

Private Sub cmdBrowseAMT_Click()
    SelectLegacyDatabase
End Sub

Private Sub cmdBrowseForFolder_Click()
    SelectAlternateOutputFolder
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdIniFileOpen_Click()
    IniFileLoadSettingsLocal
End Sub

Private Sub cmdIniFileSave_Click()
    If ValidatePairsIdentificationMode() Then
        IniFileSaveSettingsLocal
    End If
End Sub

Private Sub cmdReadFromSelectedGel_Click()
    ReadFromSelectedGel
End Sub

Private Sub cmdResetToDefaults_Click()
    ResetToDefaults
End Sub

Private Sub cmdRevert_Click()
    RevertToSavedSettings
End Sub

Private Sub cmdSelectingMassTagsCancel_Click()
    fraSelectingMassTags.Visible = False
    HideMTConnectionClassForm objMTConnectionSelector

End Sub

Private Sub cmdSelectingMassTagsOK_Click()
    fraSelectingMassTags.Visible = False

    FillDBSettingsUsingAnalysisObject mCurrentSettings.DBSettings, objSelectMassTags
    
    ' Determine number of matching MT tags for the given settings
    mCurrentSettings.DBSettings.SelectedMassTagCount = GetMassTagMatchCount(mCurrentSettings.DBSettings, LookupCurrentJob(), Me)
    
    DisplayCurrentDBSettings

End Sub

Private Sub cmdSelectMassTags_Click()
    SelectMassTagsForCurrentDB
End Sub

Private Sub cmdSelectOtherDB_Click()
    On Error Resume Next
    
    EnableDisableControlButtons False
    
    Set objMTConnectionSelector = New DummyAnalysisInitiator
    objMTConnectionSelector.GetNewAnalysisDialog glInitFile

End Sub

Private Sub cmdSetPairsToC13_Click()
    SetPairSearchDeltas glC12C13_DELTA, 1, 100
    EnableDisableScanByScanAndIReport False
End Sub

Private Sub cmdSetPairsToN15_Click()
    SetPairSearchDeltas glN14N15_DELTA, 1, 100
    EnableDisableScanByScanAndIReport False
End Sub

Private Sub cmdSetPairsToO18_Click()
    SetPairSearchDeltas glO16O18_DELTA, 1, 1
    EnableDisableScanByScanAndIReport True
End Sub

Private Sub Form_Load()
    InitializeForm
    ShowHidePNNLMenus
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
GetOLEData Data
End Sub

Private Sub fraOptionFrame_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetOLEData Data
End Sub

Private Sub lstDBSearchModes_Click()
    UpdateDynamicControls
End Sub

Private Sub objMTConnectionSelector_DialogClosed()
    HandleMTConnectionSelectorDialogClose
End Sub

Private Sub optAutoToleranceRefinementDBSearchTolType_Click(Index As Integer)
    With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement
        If Index = 0 Then
            .DBSearchTolType = gltPPM
        Else
            .DBSearchTolType = gltABS
        End If
    End With
End Sub

Private Sub optDBSearchModType_Click(Index As Integer)
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex).MassMods
            If Index = MODS_DECOY Then
                .ModMode = 2
            ElseIf Index = MODS_DYNAMIC Then
                .ModMode = 1
            Else
                .ModMode = 0
            End If
        End With
    End If
End Sub

Private Sub optDBSearchMWField_Click(Index As Integer)
    mCurrentSettings.AMTDef.MWField = Index + MW_FIELD_OFFSET
End Sub

Private Sub optDBSearchNType_Click(Index As Integer)
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex).MassMods
            If Index = SEARCH_N15 Then
                .N15InsteadOfN14 = True
            Else
                .N15InsteadOfN14 = False
            End If
        End With
    End If
End Sub

Private Sub optDBSearchTolType_Click(Index As Integer)
    If Index = 0 Then
        mCurrentSettings.AMTDef.TolType = gltPPM
    Else
        mCurrentSettings.AMTDef.TolType = gltABS
    End If
End Sub

Private Sub optERCalc_Click(Index As Integer)
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.ERCalcType = Index
End Sub

Private Sub optIsoDataFrom_Click(Index As Integer)
     mCurrentSettings.Prefs.IsoDataField = Index + MW_FIELD_OFFSET
End Sub

Private Sub optPairsDBSearchLabelAssumption_Click(Index As Integer)
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If Index = ASSUME_MT_LABELED Then
                .PairSearchAssumeMassTagsAreLabeled = True
            Else
                .PairSearchAssumeMassTagsAreLabeled = False
            End If
            
        End With
    End If
End Sub

Private Sub optPairTolType_Click(Index As Integer)
    If Index = 0 Then
       mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaMassTolType = gltPPM
    Else
       mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaMassTolType = gltABS
    End If
End Sub

Private Sub optRefineMassCalibrationMassType_Click(Index As Integer)
    If Index = 0 Then
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MassCalibrationTolType = gltPPM
    Else
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MassCalibrationTolType = gltABS
    End If
    
    UpdateDynamicControls
End Sub

Private Sub optNETAdjTolType_Click(Index As Integer)
    If Index = 0 Then
        mCurrentSettings.UMCNetAdjDef.MWTolType = gltPPM
    Else
        mCurrentSettings.UMCNetAdjDef.MWTolType = gltABS
    End If
End Sub

Private Sub optRemovePairMemberHitsRemoveHeavy_Click(Index As Integer)
    If Index = rpmhRemoveLight Then
        mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsRemoveHeavy = False
    Else
        mCurrentSettings.PrefsExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsRemoveHeavy = True
    End If
End Sub

Private Sub optUMCSearchMWField_Click(Index As Integer)
    mCurrentSettings.UMCDef.MWField = Index + MW_FIELD_OFFSET
End Sub

Private Sub optUMCSearchTolType_Click(Index As Integer)
    If Index = 0 Then
        mCurrentSettings.UMCDef.TolType = gltPPM
    Else
        mCurrentSettings.UMCDef.TolType = gltABS
    End If
End Sub

Private Sub tbsTabStrip_OLEDragDrop(Data As TabDlg.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetOLEData Data
End Sub

''Private Sub tmrTimer_Timer()
''    If mRequestUpdateRobustNETIterationCount Then
''        PredictRobustNETIterationCount
''        mRequestUpdateRobustNETIterationCount = False
''    End If
''End Sub

Private Sub txtAlkylationMWCorrection_Change()
    ' Note: This needs to fire on Change and not on LostFocus.
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If IsNumeric(txtAlkylationMWCorrection) Then
                .MassMods.AlkylationMass = CDblSafe(txtAlkylationMWCorrection)
            Else
                txtAlkylationMWCorrection = glALKYLATION
                .MassMods.AlkylationMass = glALKYLATION
            End If
        End With
    End If
End Sub

Private Sub txtAutoRefineTextbox_LostFocus(Index As Integer)

    Select Case Index
    Case uarRemoveLowIntensity
        ValidateTextboxValueDbl txtAutoRefineTextbox(Index), 0, 100, 30
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefinePctLowAbundance = CDblSafe(txtAutoRefineTextbox(Index))
    Case uarRemoveHighIntensity
        ValidateTextboxValueDbl txtAutoRefineTextbox(Index), 0, 100, 30
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefinePctHighAbundance = CDblSafe(txtAutoRefineTextbox(Index))
    Case uarRemoveShort
        ValidateTextboxValueLng txtAutoRefineTextbox(Index), 0, 100000, 3
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineMinLength = CLngSafe(txtAutoRefineTextbox(Index))
        UpdateUMCIonNetFindSingleMemberClasses
    Case uarRemoveLong
        ValidateTextboxValueLng txtAutoRefineTextbox(Index), 1, 100000, 300
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineMaxLength = CLngSafe(txtAutoRefineTextbox(Index))
    Case uarRemoveMaxLengthPctAllScans
        ValidateTextboxValueLng txtAutoRefineTextbox(Index), 1, 100, 20
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefineMaxLengthPctAllScans = CLngSafe(txtAutoRefineTextbox(Index))
    Case uarPercentMaxAbuToUseToGaugeLength
        ValidateTextboxValueLng txtAutoRefineTextbox(Index), 1, 100, 25
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.UMCAutoRefinePercentMaxAbuToUseForLength = CLngSafe(txtAutoRefineTextbox(Index))
    Case uarMinimumMemberCount
        ValidateTextboxValueLng txtAutoRefineTextbox(Index), 0, 100000, 3
        mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.MinMemberCountWhenUsingScanRange = CLngSafe(txtAutoRefineTextbox(Index))
    Case Else
        Debug.Assert False
    End Select

End Sub

Private Sub txtAutoToleranceRefinementDBSearchMWTol_LostFocus()
    ValidateTextboxValueDbl txtAutoToleranceRefinementDBSearchMWTol, 0, 1E+300, 25
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchMWTol = CDblSafe(txtAutoToleranceRefinementDBSearchMWTol)
    DisplayDataValidationWarnings
End Sub

Private Sub txtAutoToleranceRefinementDBSearchNETTol_LostFocus()
    ValidateTextboxValueDbl txtAutoToleranceRefinementDBSearchNETTol, 0.0001, 100, 0.1
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchNETTol = CDblSafe(txtAutoToleranceRefinementDBSearchNETTol)
End Sub

Private Sub txtClassAbuTopXMinMaxAbu_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtClassAbuTopXMinMaxAbu(0), 0, 1E+300, 0
        mCurrentSettings.PrefsExpanded.UMCAdvancedStatsOptions.ClassAbuTopXMinAbu = CDblSafe(txtClassAbuTopXMinMaxAbu(0))
    Else
        ValidateTextboxValueDbl txtClassAbuTopXMinMaxAbu(1), 0, 1E+300, 0
        mCurrentSettings.PrefsExpanded.UMCAdvancedStatsOptions.ClassAbuTopXMaxAbu = CDblSafe(txtClassAbuTopXMinMaxAbu(1))
    End If
End Sub

Private Sub txtClassAbuTopXMinMembers_Lostfocus()
    ValidateTextboxValueLng txtClassAbuTopXMinMembers, 0, 100000, 3
    mCurrentSettings.PrefsExpanded.UMCAdvancedStatsOptions.ClassAbuTopXMinMembers = CLngSafe(txtClassAbuTopXMinMembers)
End Sub

Private Sub txtClassMassTopXMinMaxAbu_Change(Index As Integer)
    UpdateDynamicControls
End Sub

Private Sub txtClassMassTopXMinMaxAbu_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtClassMassTopXMinMaxAbu(0), 0, 1E+300, 0
        mCurrentSettings.PrefsExpanded.UMCAdvancedStatsOptions.ClassMassTopXMinAbu = CDblSafe(txtClassMassTopXMinMaxAbu(0))
    Else
        ValidateTextboxValueDbl txtClassMassTopXMinMaxAbu(1), 0, 1E+300, 0
        mCurrentSettings.PrefsExpanded.UMCAdvancedStatsOptions.ClassMassTopXMaxAbu = CDblSafe(txtClassMassTopXMinMaxAbu(1))
    End If
End Sub

Private Sub txtClassMassTopXMinMembers_Lostfocus()
    ValidateTextboxValueLng txtClassMassTopXMinMembers, 0, 100000, 3
    mCurrentSettings.PrefsExpanded.UMCAdvancedStatsOptions.ClassMassTopXMinMembers = CLngSafe(txtClassMassTopXMinMembers)
End Sub

Private Sub txtConstraint_LostFocus(Index As Integer)
    Dim dblMinimum As Double, dblMaximum As Double, dblDefault As Double
    
    LookupDefaultUMCIonNetValues Index, dblMinimum, dblMaximum, dblDefault, False
    
    ValidateTextboxValueDbl txtConstraint(Index), dblMinimum, dblMaximum, dblDefault
    mCurrentSettings.UMCIonNetDef.MetricData(Index).ConstraintValue = CDbl(txtConstraint(Index).Text)
End Sub

Private Sub txtCSMinMaxAbu_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtCSMinMaxAbu(0), 0, 1E+300, 0
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictCSAbundanceMin = CDblSafe(txtCSMinMaxAbu(0))
        VerifyLinkedFilterValues False, True
    Else
        ValidateTextboxValueDbl txtCSMinMaxAbu(1), 0, 1E+300, 1E+15
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictCSAbundanceMax = CDblSafe(txtCSMinMaxAbu(1))
        VerifyLinkedFilterValues False, False
    End If
End Sub

Private Sub txtCSMinMaxMW_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtCSMinMaxMW(0), 0, MAX_MW, 400
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictCSMassMin = CDblSafe(txtCSMinMaxMW(0))
        VerifyLinkedFilterValues False, True
    Else
        ValidateTextboxValueDbl txtCSMinMaxMW(1), 0, MAX_MW, 6000
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictCSMassMax = CDblSafe(txtCSMinMaxMW(1))
        VerifyLinkedFilterValues False, False
    End If
End Sub

Private Sub txtDBConnectionRetryAttemptMax_LostFocus()
    ValidateTextboxValueLng txtDBConnectionRetryAttemptMax, 1, 100, 3
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.DBConnectionRetryAttemptMax = CLngSafe(txtDBConnectionRetryAttemptMax)
End Sub

Private Sub txtDBConnectionTimeoutSeconds_LostFocus()
    ValidateTextboxValueLng txtDBConnectionTimeoutSeconds, 10, 100000, 300
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.DBConnectionTimeoutSeconds = CLngSafe(txtDBConnectionTimeoutSeconds)
End Sub

Private Sub txtDBSearchAlternateOutputFolderPath_Change()
    ' Note: This needs to fire on Change and not on LostFocus.
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            .AlternateOutputFolderPath = txtDBSearchAlternateOutputFolderPath
        End With
    End If
End Sub

Private Sub txtDBSearchMinimumHighDiscriminantScore_Change()
    ' Note: This needs to fire on Change and not on LostFocus.
    Dim blnNotNumeric As Boolean
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If IsNumeric(txtDBSearchMinimumHighDiscriminantScore) Then
                .DBSearchMinimumHighDiscriminantScore = CSngSafe(txtDBSearchMinimumHighDiscriminantScore)
                blnNotNumeric = False
            Else
                blnNotNumeric = True
            End If
            
            If blnNotNumeric Or .DBSearchMinimumHighDiscriminantScore < 0 Or .DBSearchMinimumHighDiscriminantScore > 1 Then
                txtDBSearchMinimumHighDiscriminantScore = 0
                .DBSearchMinimumHighDiscriminantScore = 0
            End If
        End With
    End If
End Sub

Private Sub txtDBSearchMinimumHighNormalizedScore_Change()
    ' Note: This needs to fire on Change and not on LostFocus.
    Dim blnNotNumeric As Boolean
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If IsNumeric(txtDBSearchMinimumHighNormalizedScore) Then
                .DBSearchMinimumHighNormalizedScore = CSngSafe(txtDBSearchMinimumHighNormalizedScore)
                blnNotNumeric = False
            Else
                blnNotNumeric = True
            End If
            
            If blnNotNumeric Or .DBSearchMinimumHighNormalizedScore < 0 Or .DBSearchMinimumHighNormalizedScore > 100000 Then
                txtDBSearchMinimumHighNormalizedScore = 0
                .DBSearchMinimumHighNormalizedScore = 0
            End If
        End With
    End If

End Sub

Private Sub txtDBSearchMinimumPeptideProphetProbability_Change()
    ' Note: This needs to fire on Change and not on LostFocus.
    Dim blnNotNumeric As Boolean
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If IsNumeric(txtDBSearchMinimumPeptideProphetProbability) Then
                .DBSearchMinimumPeptideProphetProbability = CSngSafe(txtDBSearchMinimumPeptideProphetProbability)
                blnNotNumeric = False
            Else
                blnNotNumeric = True
            End If
            
            If blnNotNumeric Or .DBSearchMinimumPeptideProphetProbability < 0 Or .DBSearchMinimumPeptideProphetProbability > 1 Then
                txtDBSearchMinimumPeptideProphetProbability = 0
                .DBSearchMinimumPeptideProphetProbability = 0
            End If
        End With
    End If

End Sub
Private Sub txtDBSearchMWTol_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMWTol, 0, 1E+300, 10
    mCurrentSettings.AMTDef.MWTol = CDblSafe(txtDBSearchMWTol)

    DisplayDataValidationWarnings
    
    ' Reset mAutoIniFileNameOverridden
    mAutoIniFileNameOverridden = False
End Sub

Private Sub txtDBSearchNETTolerance_LostFocus()
    ValidateTextboxValueDbl txtDBSearchNETTolerance, 0.0001, 100, 0.1
    mCurrentSettings.AMTDef.NETTol = CDblSafe(txtDBSearchNETTolerance)

    ' Reset mAutoIniFileNameOverridden
    mAutoIniFileNameOverridden = False
End Sub

Private Sub txtDupTolerance_LostFocus()
    ValidateTextboxValueDbl txtDupTolerance, 0, MAX_MW, 2
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeDuplicatesTolerance = CDblSafe(txtDupTolerance)
    With mCurrentSettings
        .Prefs.DupTolerance = .PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeDuplicatesTolerance
    End With
End Sub

Private Sub txtEMRefineMassErrorPeakToleranceEstimatePPM_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtEMRefineMassErrorPeakToleranceEstimatePPM, KeyAscii, True, True
End Sub

Private Sub txtEMRefineMassErrorPeakToleranceEstimatePPM_LostFocus()
    ValidateTextboxValueDbl txtEMRefineMassErrorPeakToleranceEstimatePPM, 0.0001, 10000, 6
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.EMMassErrorPeakToleranceEstimatePPM = CSngSafe(txtEMRefineMassErrorPeakToleranceEstimatePPM)
End Sub

Private Sub txtEMRefineNETErrorPeakToleranceEstimate_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtEMRefineNETErrorPeakToleranceEstimate, KeyAscii, True, True
End Sub

Private Sub txtEMRefineNETErrorPeakToleranceEstimate_LostFocus()
    ValidateTextboxValueDbl txtEMRefineNETErrorPeakToleranceEstimate, 0.00001, 1, 0.05
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.EMNETErrorPeakToleranceEstimate = CSngSafe(txtEMRefineNETErrorPeakToleranceEstimate)
End Sub

Private Sub txtEMRefinePercentOfDataToExclude_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtEMRefinePercentOfDataToExclude, KeyAscii, True, False
End Sub

Private Sub txtEMRefinePercentOfDataToExclude_LostFocus()
    ValidateTextboxValueDbl txtEMRefinePercentOfDataToExclude, 0, 100, 10
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.EMPercentOfDataToExclude = CIntSafe(txtEMRefinePercentOfDataToExclude)
End Sub

Private Sub txtErrorGraphicGANETBinSize_LostFocus()
    ValidateTextboxValueDbl txtErrorGraphicGANETBinSize, 0.0001, 10, DEFAULT_GANET_BIN_SIZE
    mCurrentSettings.PrefsExpanded.ErrorPlottingOptions.GANETBinSize = CDblSafe(txtErrorGraphicGANETBinSize)
End Sub

Private Sub txtErrorGraphicGANETRange_LostFocus()
    ValidateTextboxValueDbl txtErrorGraphicGANETRange, 0.001, 10, 0.3
    mCurrentSettings.PrefsExpanded.ErrorPlottingOptions.GANETRange = CDblSafe(txtErrorGraphicGANETRange)
End Sub

Private Sub txtErrorGraphicMassBinSizePPM_LostFocus()
    ValidateTextboxValueDbl txtErrorGraphicMassBinSizePPM, 0.01, 10000, DEFAULT_MASS_BIN_SIZE_PPM
    mCurrentSettings.PrefsExpanded.ErrorPlottingOptions.MassBinSizePPM = CDblSafe(txtErrorGraphicMassBinSizePPM)
End Sub

Private Sub txtErrorGraphicMassRangePPM_LostFocus()
    ValidateTextboxValueDbl txtErrorGraphicMassRangePPM, 1, 10000, 100
    mCurrentSettings.PrefsExpanded.ErrorPlottingOptions.MassRangePPM = CDblSafe(txtErrorGraphicMassRangePPM)
End Sub

Private Sub txtExcludeIsoByFit_LostFocus()
    ValidateTextboxValueDbl txtExcludeIsoByFit, 0, 100, 0.15
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeIsoByFitMaxVal = CDblSafe(txtExcludeIsoByFit)

    With mCurrentSettings
        .Prefs.IsoDataFit = .PrefsExpanded.AutoAnalysisFilterPrefs.ExcludeIsoByFitMaxVal
    End With

End Sub

Private Sub txtHoleNum_LostFocus()
    ValidateTextboxValueLng txtHoleNum, 0, max_scan, 10
    mCurrentSettings.UMCDef.GapMaxCnt = CLngSafe(txtHoleNum)
End Sub

Private Sub txtHolePct_LostFocus()
    ' UMCListType2002 Only
    ValidateTextboxValueDbl txtHolePct, 0, 100, 80
    mCurrentSettings.UMCDef.GapMaxPct = Round(CDblSafe(txtHolePct) / 100#, 4)
End Sub

Private Sub txtHoleSize_Change(Index As Integer)
    If Index = 0 Then
        txtHoleSize(1) = txtHoleSize(0)
    Else
        txtHoleSize(0) = txtHoleSize(1)
    End If
End Sub

Private Sub txtHoleSize_LostFocus(Index As Integer)
    ValidateTextboxValueLng txtHoleSize(Index), 0, max_scan, 4
    mCurrentSettings.UMCDef.GapMaxSize = CLngSafe(txtHoleSize(Index))
End Sub

Private Sub txtInterpolateMaxGapSize_LostFocus()
    ValidateTextboxValueLng txtInterpolateMaxGapSize, 0, max_scan, 3
    mCurrentSettings.UMCDef.InterpolateMaxGapSize = CLngSafe(txtInterpolateMaxGapSize)
End Sub

Private Sub txtIsoMinMaxAbu_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtIsoMinMaxAbu(0), 0, 1E+300, 0
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoAbundanceMin = CDblSafe(txtIsoMinMaxAbu(0))
        VerifyLinkedFilterValues True, True
    Else
        ValidateTextboxValueDbl txtIsoMinMaxAbu(1), 0, 1E+300, 1E+15
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoAbundanceMax = CDblSafe(txtIsoMinMaxAbu(1))
        VerifyLinkedFilterValues True, False
    End If
End Sub

Private Sub txtIsoMinMaxCS_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueLng txtIsoMinMaxCS(0), 0, 100, 0
        If Not ValidateDualTextBoxes(txtIsoMinMaxCS(0), txtIsoMinMaxCS(1), True, 0, 100, 0) Then
            txtIsoMinMaxCS(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoChargeStateMin = CLngSafe(txtIsoMinMaxCS(0))
    Else
        ValidateTextboxValueLng txtIsoMinMaxCS(1), 0, 100, 6
        If Not ValidateDualTextBoxes(txtIsoMinMaxCS(0), txtIsoMinMaxCS(1), False, 0, 100, 0) Then
            txtIsoMinMaxCS(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoChargeStateMax = CLngSafe(txtIsoMinMaxCS(1))
    End If
End Sub

Private Sub txtIsoMinMaxMW_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtIsoMinMaxMW(0), 0, MAX_MW, 400
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoMassMin = CDblSafe(txtIsoMinMaxMW(0))
        VerifyLinkedFilterValues True, True
    Else
        ValidateTextboxValueDbl txtIsoMinMaxMW(1), 0, MAX_MW, 6000
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoMassMax = CDblSafe(txtIsoMinMaxMW(1))
        VerifyLinkedFilterValues True, False
    End If
End Sub

Private Sub txtIsoMinMaxMZ_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtIsoMinMaxMZ(0), 1, MAX_MZ, 0
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoMZMin = CDblSafe(txtIsoMinMaxMZ(0))
        VerifyLinkedFilterValues True, True
    Else
        ValidateTextboxValueDbl txtIsoMinMaxMZ(1), 1, MAX_MZ, 6000
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictIsoMZMax = CDblSafe(txtIsoMinMaxMZ(1))
        VerifyLinkedFilterValues True, False
    End If
End Sub

Private Sub txtLegacyAMTDatabasePath_LostFocus()
    mCurrentSettings.PrefsExpanded.LegacyAMTDBPath = txtLegacyAMTDatabasePath.Text
End Sub

Private Sub txtMaximumDataCountToLoad_LostFocus()
    ValidateTextboxValueLng txtMaximumDataCountToLoad, 0, 10000000#, DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD
    mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.MaximumDataCountToLoad = CLngSafe(txtMaximumDataCountToLoad)
End Sub

Private Sub txtNetAdjInitialNETTol_LostFocus()
    ValidateChangedInitialNETTol
End Sub

Private Sub txtNETAdjMWTol_LostFocus()
    ValidateTextboxValueDbl txtNETAdjMWTol, 0, 1E+300, 10
    mCurrentSettings.UMCNetAdjDef.MWTol = CDblSafe(txtNETAdjMWTol)
End Sub

Private Sub txtNETType_LostFocus()
    ValidateTextboxValueLng txtNETType, 1, 1000, Net_SPIDER_66
    mCurrentSettings.UMCIonNetDef.NETType = CLngSafe(txtNETType)
End Sub

Private Sub txtOutputFileSeparationCharacter_LostFocus()
    If Len(txtOutputFileSeparationCharacter) <> 1 Then
        If txtOutputFileSeparationCharacter <> SEPARATION_CHARACTER_TAB_STRING Then
            txtOutputFileSeparationCharacter = SEPARATION_CHARACTER_TAB_STRING
        End If
    End If
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.OutputFileSeparationCharacter = txtOutputFileSeparationCharacter
End Sub

Private Sub txtPairsDelta_LostFocus()
    ValidateTextboxValueDbl txtPairsDelta, 0.01, MAX_MW, glN14_N15CorrMW
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaMass = CDblSafe(txtPairsDelta)
End Sub

Private Sub txtPairsDeltaTol_LostFocus()
    ValidateTextboxValueDbl txtPairsDeltaTol, 0.00001, MAX_MW, 0.02
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaMassTolerance = CDblSafe(txtPairsDeltaTol)
End Sub

Private Sub txtPairsERMinMax_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtPairsERMinMax(0), -1E+300, 1E+300, -5
        If Not ValidateDualTextBoxes(txtPairsERMinMax(0), txtPairsERMinMax(1), True, -1E+300, 1E+300, 0) Then
            txtPairsERMinMax(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.ERInclusionMin = CDblSafe(txtPairsERMinMax(0))
    Else
        ValidateTextboxValueDbl txtPairsERMinMax(1), -1E+300, 1E+300, 5
        If Not ValidateDualTextBoxes(txtPairsERMinMax(0), txtPairsERMinMax(1), False, -1E+300, 1E+300, 0) Then
            txtPairsERMinMax(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.ERInclusionMax = CDblSafe(txtPairsERMinMax(1))
    End If
End Sub

Private Sub txtPairsHeavyLabelDelta_LostFocus()
    ValidateTextboxValueDbl txtPairsHeavyLabelDelta, 0.01, MAX_MW, Round(glICAT_Heavy - glICAT_Light, 3)
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.HeavyLightMassDifference = CDblSafe(txtPairsHeavyLabelDelta)
End Sub

Private Sub txtPairsLabel_LostFocus()
    ValidateTextboxValueDbl txtPairsLabel, 0.01, MAX_MW, glICAT_Light
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.LightLabelMass = CDblSafe(txtPairsLabel)
End Sub

Private Sub txtPairsMaxLblDiff_LostFocus()
    ValidateTextboxValueLng txtPairsMaxLblDiff, 1, MAX_MW, 1
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels = CLngSafe(txtPairsMaxLblDiff)
End Sub

Private Sub txtPairsMinMaxDelta_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        ValidateTextboxValueLng txtPairsMinMaxDelta(0), 1, MAX_MW, 1
        If Not ValidateDualTextBoxes(txtPairsMinMaxDelta(0), txtPairsMinMaxDelta(1), True, 1, MAX_MW, 0) Then
            txtPairsMinMaxDelta(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaCountMin = CLngSafe(txtPairsMinMaxDelta(0))
    Case 1
        ValidateTextboxValueLng txtPairsMinMaxDelta(1), 1, MAX_MW, 100
        If Not ValidateDualTextBoxes(txtPairsMinMaxDelta(0), txtPairsMinMaxDelta(1), False, 1, MAX_MW, 0) Then
            txtPairsMinMaxDelta(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaCountMax = CLngSafe(txtPairsMinMaxDelta(1))
    Case 2
        ValidateTextboxValueLng txtPairsMinMaxDelta(2), 1, MAX_MW, 1
        mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.DeltaStepSize = CLngSafe(txtPairsMinMaxDelta(2))
    Case Else
        MsgBox "Error: Unknown Index sent to txtPairsMinMaxDelta_LostFocus", vbExclamation + vbOKOnly, "Error"
    End Select
End Sub

Private Sub txtPairsMinMaxLbl_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueLng txtPairsMinMaxLbl(0), 1, MAX_MW, 1
        If Not ValidateDualTextBoxes(txtPairsMinMaxLbl(0), txtPairsMinMaxLbl(1), True, 1, MAX_MW, 0) Then
            txtPairsMinMaxLbl(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.LabelCountMin = CLngSafe(txtPairsMinMaxLbl(0))
    Else
        ValidateTextboxValueLng txtPairsMinMaxLbl(1), 1, MAX_MW, 5
        If Not ValidateDualTextBoxes(txtPairsMinMaxLbl(0), txtPairsMinMaxLbl(1), False, 1, MAX_MW, 0) Then
            txtPairsMinMaxLbl(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.LabelCountMax = CLngSafe(txtPairsMinMaxLbl(1))
    End If
End Sub

Private Sub txtPairsScanTolApex_LostFocus()
    ValidateTextboxValueLng txtPairsScanTolApex, 1, max_scan, 10
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.ScanToleranceAtApex = CLngSafe(txtPairsScanTolApex)
End Sub

Private Sub txtPairsScanTolEdge_LostFocus()
    ValidateTextboxValueLng txtPairsScanTolEdge, 1, max_scan, 5
    mCurrentSettings.PrefsExpanded.PairSearchOptions.SearchDef.ScanTolerance = CLngSafe(txtPairsScanTolEdge)
End Sub

Private Sub txtPEKFileExtensionPreferenceOrder_LostFocus()
    If Len(txtPEKFileExtensionPreferenceOrder) = 0 Then
        txtPEKFileExtensionPreferenceOrder = DEFAULT_PEK_FILE_EXTENSION_ORDER
    End If
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.PEKFileExtensionPreferenceOrder = txtPEKFileExtensionPreferenceOrder
End Sub

Private Sub txtPictureGraphicSizeHeightPixels_LostFocus()
    ValidateTextboxValueLng txtPictureGraphicSizeHeightPixels, 50, 100000, 768
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SavePictureHeightPixels = CLngSafe(txtPictureGraphicSizeHeightPixels)
End Sub

Private Sub txtPictureGraphicSizeWidthPixels_LostFocus()
    ValidateTextboxValueLng txtPictureGraphicSizeWidthPixels, 50, 100000, 1024
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.SavePictureWidthPixels = CLngSafe(txtPictureGraphicSizeWidthPixels)
End Sub

Private Sub txtRefineDBSearchMassToleranceAdjustmentMultiplier_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchMassToleranceAdjustmentMultiplier, 0.0001, 10000, 1
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MassToleranceAdjustmentMultiplier = CDblSafe(txtRefineDBSearchMassToleranceAdjustmentMultiplier)
End Sub

Private Sub txtRefineDBSearchMassToleranceMinMax_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtRefineDBSearchMassToleranceMinMax(0), 0, 1E+300, 1
        If Not ValidateDualTextBoxes(txtRefineDBSearchMassToleranceMinMax(0), txtRefineDBSearchMassToleranceMinMax(1), True, 0, 1E+300, 0) Then
            txtRefineDBSearchMassToleranceMinMax(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MassToleranceMinimum = CDblSafe(txtRefineDBSearchMassToleranceMinMax(0))
    Else
        ValidateTextboxValueDbl txtRefineDBSearchMassToleranceMinMax(1), 0, 1E+300, 10
        If Not ValidateDualTextBoxes(txtRefineDBSearchMassToleranceMinMax(0), txtRefineDBSearchMassToleranceMinMax(1), False, 0, 1E+300, 0) Then
            txtRefineDBSearchMassToleranceMinMax(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MassToleranceMaximum = CDblSafe(txtRefineDBSearchMassToleranceMinMax(1))
    End If
End Sub

Private Sub txtRefineDBSearchNETToleranceAdjustmentMultiplier_LostFocus()
    ValidateTextboxValueDbl txtRefineDBSearchNETToleranceAdjustmentMultiplier, 0.0001, 10000, 1
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.NETToleranceAdjustmentMultiplier = CDblSafe(txtRefineDBSearchNETToleranceAdjustmentMultiplier)
End Sub

Private Sub txtRefineDBSearchNETToleranceMinMax_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtRefineDBSearchNETToleranceMinMax(0), 0.0001, 100, 0.01
        If Not ValidateDualTextBoxes(txtRefineDBSearchNETToleranceMinMax(0), txtRefineDBSearchNETToleranceMinMax(1), True, 0.0001, 100, 0) Then
            txtRefineDBSearchNETToleranceMinMax(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.NETToleranceMinimum = CDblSafe(txtRefineDBSearchNETToleranceMinMax(0))
    Else
        ValidateTextboxValueDbl txtRefineDBSearchNETToleranceMinMax(1), 0.0001, 100, 0.2
        If Not ValidateDualTextBoxes(txtRefineDBSearchNETToleranceMinMax(0), txtRefineDBSearchNETToleranceMinMax(1), False, 0.0001, 100, 0) Then
            txtRefineDBSearchNETToleranceMinMax(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.NETToleranceMaximum = CDblSafe(txtRefineDBSearchNETToleranceMinMax(1))
    End If
    
End Sub

Private Sub txtRefineMassCalibrationMaximumShift_LostFocus()
    ValidateTextboxValueDbl txtRefineMassCalibrationMaximumShift, 0, 1E+300, 15
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MassCalibrationMaximumShift = CDblSafe(txtRefineMassCalibrationMaximumShift)
End Sub

Private Sub txtRefineMassCalibrationOverridePPM_LostFocus()
    ValidateTextboxValueDbl txtRefineMassCalibrationOverridePPM, -1E+300, 1E+300, 0
    mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibrationOverridePPM = CDblSafe(txtRefineMassCalibrationOverridePPM)
End Sub

Private Sub txtRejectLongConnections_LostFocus()
    ValidateTextboxValueDbl txtRejectLongConnections, 0, 100, 0.1
    mCurrentSettings.UMCIonNetDef.TooDistant = CDblSafe(txtRejectLongConnections)
End Sub

Private Sub txtResidueToModifyMass_Change()
    ' Note: This needs to fire on Change and not on LostFocus.
    If lstDBSearchModes.ListIndex >= 0 Then
        With mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(lstDBSearchModes.ListIndex)
            If IsNumeric(txtResidueToModifyMass) Then
                .MassMods.ResidueMassModification = CDblSafe(txtResidueToModifyMass)
            Else
                txtResidueToModifyMass = "0"
                .MassMods.ResidueMassModification = 0
            End If
        End With
    End If
End Sub

Private Sub txtRestrictGANETRangeMinMax_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueDbl txtRestrictGANETRangeMinMax(0), -100, 100, -1
        If Not ValidateDualTextBoxes(txtRestrictGANETRangeMinMax(0), txtRestrictGANETRangeMinMax(1), False, -100, 100, 0) Then
            txtRestrictGANETRangeMinMax(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictGANETRangeMin = CDblSafe(txtRestrictGANETRangeMinMax(0))
    Else
        ValidateTextboxValueDbl txtRestrictGANETRangeMinMax(1), -100, 100, 2
        If Not ValidateDualTextBoxes(txtRestrictGANETRangeMinMax(0), txtRestrictGANETRangeMinMax(1), False, -100, 100, 0) Then
            txtRestrictGANETRangeMinMax(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictGANETRangeMax = CDblSafe(txtRestrictGANETRangeMinMax(1))
    End If
End Sub

Private Sub txtRestrictScanRangeMinMax_LostFocus(Index As Integer)
    If Index = 0 Then
        ValidateTextboxValueLng txtRestrictScanRangeMinMax(0), 0, max_scan, 0
        If Not ValidateDualTextBoxes(txtRestrictScanRangeMinMax(0), txtRestrictScanRangeMinMax(1), True, 0, max_scan, 0) Then
            txtRestrictScanRangeMinMax(1).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictScanRangeMin = CLngSafe(txtRestrictScanRangeMinMax(0))
    Else
        ValidateTextboxValueLng txtRestrictScanRangeMinMax(1), 0, max_scan, 1500
        If Not ValidateDualTextBoxes(txtRestrictScanRangeMinMax(0), txtRestrictScanRangeMinMax(1), False, 0, max_scan, 0) Then
            txtRestrictScanRangeMinMax(0).SetFocus
        End If
        mCurrentSettings.PrefsExpanded.AutoAnalysisFilterPrefs.RestrictScanRangeMax = CLngSafe(txtRestrictScanRangeMinMax(1))
    End If
End Sub

Private Sub txtRobustNETOption_LostFocus(Index As Integer)
    Select Case Index
    Case rnoWarpNumberOfSections
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.NumberOfSections = ValidateTextboxValueLng(txtRobustNETOption(Index), 5, 500, 100)
    
    Case rnoWarpMaxDistortion
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MaxDistortion = ValidateTextboxValueLng(txtRobustNETOption(Index), 1, 200, 3)
    
    Case rnoWarpContractionFactor
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.ContractionFactor = ValidateTextboxValueLng(txtRobustNETOption(Index), 1, 10, 2)
    
    Case rnoWarpMinimumPMTTagObsCount
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MinimumPMTTagObsCount = ValidateTextboxValueLng(txtRobustNETOption(Index), 1, 100000, 5)
    
    Case rnoWarpMatchPromiscuity
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MatchPromiscuity = ValidateTextboxValueLng(txtRobustNETOption(Index), 1, 1000, 2)
        
    Case rnoWarpMassWindow
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MassWindowPPM = ValidateTextboxValueDbl(txtRobustNETOption(Index), 1, 10000, 50)
        
    Case rnoWarpMassSplineOrder
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MassSplineOrder = ValidateTextboxValueLng(txtRobustNETOption(Index), 0, 3, 2)
        
    Case rnoWarpNumXAxisSlices
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MassNumXSlices = ValidateTextboxValueLng(txtRobustNETOption(Index), 0, 50, 20)
        
    Case rnoWarpNumMassDeltaBins
        mCurrentSettings.UMCNetAdjDef.MSWarpOptions.MassNumMassDeltaBins = ValidateTextboxValueLng(txtRobustNETOption(Index), 1, 200, 100)
        ValidateNETWarpMassMaxJump
        
    Case rnoWarpMassMaxJump
        ValidateNETWarpMassMaxJump
    Case Else
        ' Unknown index
        Debug.Assert False
    End Select

    ''mRequestUpdateRobustNETIterationCount = True

End Sub

Private Sub txtSplitUMCsMaximumPeakCount_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsMaximumPeakCount, 2, 100000, 6
    mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.SplitUMCOptions.MaximumPeakCountToSplitUMC = CLngSafe(txtSplitUMCsMaximumPeakCount)
End Sub

Private Sub txtSplitUMCsMinimumDifferenceInAvgPpmMass_LostFocus()
    ValidateTextboxValueDbl txtSplitUMCsMinimumDifferenceInAvgPpmMass, 0, 10000#, 4
    mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.SplitUMCOptions.MinimumDifferenceInAveragePpmMassToSplit = CDblSafe(txtSplitUMCsMinimumDifferenceInAvgPpmMass)
End Sub

Private Sub txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax, 0, 100, 15
    mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.SplitUMCOptions.PeakDetectIntensityThresholdPercentageOfMaximum = CLngSafe(txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax)
End Sub

Private Sub txtSplitUMCsPeakPickingMinimumWidth_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsPeakPickingMinimumWidth, 0, 1000, 4
    mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.SplitUMCOptions.PeakWidthPointsMinimum = CLngSafe(txtSplitUMCsPeakPickingMinimumWidth)
End Sub

Private Sub txtSplitUMCsStdDevMultiplierForSplitting_LostFocus()
    ValidateTextboxValueDbl txtSplitUMCsStdDevMultiplierForSplitting, 0, 1000, 1
    mCurrentSettings.PrefsExpanded.UMCAutoRefineOptions.SplitUMCOptions.StdDevMultiplierForSplitting = CSngSafe(txtSplitUMCsStdDevMultiplierForSplitting)
End Sub

Private Sub txtToleranceRefinementFilter_LostFocus(Index As Integer)
    Select Case Index
    Case trfToleranceRefinementFilterOptionsConstants.trfMinimumHighNormalizedScore
        ValidateTextboxValueDbl txtToleranceRefinementFilter(Index), 0, 100000, 2.5
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchMinimumHighNormalizedScore = CSngSafe(txtToleranceRefinementFilter(Index))
        
    Case trfToleranceRefinementFilterOptionsConstants.trfMinimumDiscriminant
        ValidateTextboxValueDbl txtToleranceRefinementFilter(Index), 0, 1, 0.5
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchMinimumHighDiscriminantScore = CSngSafe(txtToleranceRefinementFilter(Index))
        
    Case trfToleranceRefinementFilterOptionsConstants.trfMinimumPeptideProphet
        ValidateTextboxValueDbl txtToleranceRefinementFilter(Index), 0, 1, 0.5
        mCurrentSettings.PrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchMinimumPeptideProphetProbability = CSngSafe(txtToleranceRefinementFilter(Index))
        
    Case trfToleranceRefinementFilterOptionsConstants.trfMinimumSLiC
        ValidateTextboxValueDbl txtToleranceRefinementFilter(Index), 0, 1, 0
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MinimumSLiC = CSngSafe(txtToleranceRefinementFilter(Index))
        
    Case trfToleranceRefinementFilterOptionsConstants.trfMaximumAbundance
        ValidateTextboxValueDbl txtToleranceRefinementFilter(Index), 0, 1E+300, 0
        mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MaximumAbundance = CDblSafe(txtToleranceRefinementFilter(Index))
        
    Case Else
        ' Unknown Index
        Debug.Assert False
    End Select
End Sub

Private Sub txtToleranceRefinementMinimumPeakHeight_LostFocus()
    ValidateTextboxValueLng txtToleranceRefinementMinimumPeakHeight, 0, 1000000000#, 25
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MinimumPeakHeight = CLngSafe(txtToleranceRefinementMinimumPeakHeight)
End Sub

Private Sub txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks_LostFocus()
    ValidateTextboxValueDbl txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks, 0, 100000, 2.5
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.MinimumSignalToNoiseRatioForLowAbundancePeaks = CSngSafe(txtToleranceRefinementMinimumSignalToNoiseForLowAbuPeaks)
End Sub

Private Sub txtToleranceRefinementPercentageOfMaxForWidth_Lostfocus()
    ValidateTextboxValueDbl txtToleranceRefinementPercentageOfMaxForWidth, 0, 100, 30
    mCurrentSettings.PrefsExpanded.RefineMSDataOptions.PercentageOfMaxForFindingWidth = CLngSafe(txtToleranceRefinementPercentageOfMaxForWidth)
End Sub

Private Sub txtUMCSearchMassTol_LostFocus()
    ValidateTextboxValueDbl txtUMCSearchMassTol, 0, 1E+300, 12.5
    mCurrentSettings.UMCDef.Tol = CDblSafe(txtUMCSearchMassTol)
End Sub

Private Sub txtWeightingFactor_LostFocus(Index As Integer)
    Dim dblMinimum As Double, dblMaximum As Double, dblDefault As Double
    
    LookupDefaultUMCIonNetValues Index, dblMinimum, dblMaximum, dblDefault, True
    
    ValidateTextboxValueDbl txtWeightingFactor(Index), dblMinimum, dblMaximum, dblDefault
    mCurrentSettings.UMCIonNetDef.MetricData(Index).WeightFactor = CDblSafe(txtWeightingFactor(Index))
End Sub


VERSION 5.00
Begin VB.Form frmGraph 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Virtual 2D Display"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8670
   ClipControls    =   0   'False
   DrawStyle       =   2  'Dot
   Icon            =   "frmGraph.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   578
   Begin VB.Timer tmrRefreshTimer 
      Interval        =   250
      Left            =   7920
      Top             =   120
   End
   Begin VB.PictureBox picGraph 
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Mask Pen
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5100
      Left            =   0
      ScaleHeight     =   336
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   0
      Top             =   0
      Width           =   7890
      Begin VIPER.SelToolBox SelToolBox1 
         Height          =   2190
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3863
         BackColor       =   -2147483633
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   105
      Begin VB.Menu mnuNew 
         Caption         =   "&New (Load peak list file)"
         HelpContextID   =   105
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuNewAnalysis 
         Caption         =   "New Anal&ysis (Choose from DMS)"
      End
      Begin VB.Menu mnuNewAutoAnalysis 
         Caption         =   "New Automatic Analysis (Choose manually)"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open .Gel file"
         HelpContextID   =   105
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         HelpContextID   =   105
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Gel file"
         HelpContextID   =   105
      End
      Begin VB.Menu mnuSaveAsCompressed 
         Caption         =   "Save &As Without Extended Info"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As With Extended Info"
         HelpContextID   =   105
      End
      Begin VB.Menu mnuSaveAsLegacyData 
         Caption         =   "Save As Legacy Data Format"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveWYS 
         Caption         =   "Save What You See (&WYS) As New Gel file"
      End
      Begin VB.Menu mnuSavePic 
         Caption         =   "Sa&ve Picture"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavePEKFileUsingAllDataPoints 
         Caption         =   "Save loaded data to new PEK file based on original"
      End
      Begin VB.Menu mnuSavePEKFileUsingDataPointsInView 
         Caption         =   "Save new PEK file using Data Points in View"
      End
      Begin VB.Menu mnuSavePEKFileUsingUMCs 
         Caption         =   "Save new PEK file using UMCs in View"
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveLoadEditAnalysisSettings 
         Caption         =   "Save/Load/Edit analysis settings"
      End
      Begin VB.Menu mnuSaveSettingsToIniFile 
         Caption         =   "Save current settings to &Ini file"
      End
      Begin VB.Menu mnuResetAllOptionsToDefaults 
         Caption         =   "Reset all options to defaults"
      End
      Begin VB.Menu mnuFileSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         HelpContextID   =   105
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Print Set&up"
      End
      Begin VB.Menu mnuFileSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         HelpContextID   =   105
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RecentFile8"
         Index           =   8
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSteps 
      Caption         =   "&Steps"
      Begin VB.Menu mnuStepsFile 
         Caption         =   "1a. &Load PEK File"
         Begin VB.Menu mnuStepsFileNew 
            Caption         =   "&New (Choose any PEK file)"
         End
         Begin VB.Menu mnuStepsFileNewChooseFromDMS 
            Caption         =   "New Anal&ysis (Choose from DMS)"
         End
      End
      Begin VB.Menu mnuStepsFilter 
         Caption         =   "1b. &Filter"
      End
      Begin VB.Menu mnuStepsUMCs 
         Caption         =   "2. Find LC-MS Features (&UMCs)"
         Begin VB.Menu mnuStepsUMCMode 
            Caption         =   "UMC 2003 (faster)"
            Index           =   0
         End
         Begin VB.Menu mnuStepsUMCMode 
            Caption         =   "UMC Ion Networks (better)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuStepsSelectMassTags 
         Caption         =   "3. Select &MT Tags (Connect to DB)"
      End
      Begin VB.Menu mnuStepsFindPairs 
         Caption         =   "4. Find &Pairs"
         Begin VB.Menu mnuStepsFindPairsMode 
            Caption         =   "Delta (UMC) - N14/N15 or O16/O18"
            Index           =   0
         End
         Begin VB.Menu mnuStepsFindPairsMode 
            Caption         =   "Label (UMC) - ICAT"
            Index           =   1
         End
         Begin VB.Menu mnuStepsFindPairsMode 
            Caption         =   "Delta-Label (UMC)"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu mnuStepsNETAdjustment 
         Caption         =   "5. NET &Adjustment"
         Begin VB.Menu mnuStepsNETAdjustmentMode 
            Caption         =   "Align and Warp data Using MS Warp (preferred)"
            Index           =   0
         End
         Begin VB.Menu mnuStepsNETAdjustmentMode 
            Caption         =   "Traditional NET Adjustment"
            Index           =   1
         End
      End
      Begin VB.Menu mnuStepsDBSearch 
         Caption         =   "6. &Database Search"
         Begin VB.Menu mnuStepsDBSearchConglomerateMass 
            Caption         =   "UMC Single Mass"
         End
      End
      Begin VB.Menu mnuStepsToleranceRefinement 
         Caption         =   "7. Mass Calibration and &Tolerance Refinement"
      End
      Begin VB.Menu mnuStepsPairsDBSearch 
         Caption         =   "8. Database Search using Pairs"
         Begin VB.Menu mnuStepsPairsDBSearchUsingPairs 
            Caption         =   "UMC N14/N15 Pairs"
            Index           =   0
         End
         Begin VB.Menu mnuStepsPairsDBSearchUsingPairs 
            Caption         =   "UMC ICAT Pairs"
            Index           =   1
         End
         Begin VB.Menu mnuStepsPairsDBSearchUsingPairs 
            Caption         =   "UMC Labeled N14/N15 Pairs "
            Index           =   2
         End
      End
      Begin VB.Menu mnuStepsSaveQCPlots 
         Caption         =   "9. Save &QC Plots..."
      End
      Begin VB.Menu mnuStepsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStepsAnalysisHistoryLog 
         Caption         =   "&View Analysis History Log"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   106
      Begin VB.Menu mnuEditCopyBMP 
         Caption         =   "Copy As &BMP"
      End
      Begin VB.Menu mnuEditCopyWMF 
         Caption         =   "Copy As &WMF"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditCopyEMF 
         Caption         =   "Copy As EMF"
      End
      Begin VB.Menu mnuEditCopyEMFIncludeFileNameAndTime 
         Caption         =   "Include Filename and Current Time"
      End
      Begin VB.Menu mnuEditCopyEMFIncludeTextLabels 
         Caption         =   "Include Text Labels"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSCopyPointsInView 
         Caption         =   "Copy Points In View"
         Begin VB.Menu mnuSCopyPointsInViewToClipboard 
            Caption         =   "Copy to Clipboard"
         End
         Begin VB.Menu mnuSCopyPointsInViewToFile 
            Caption         =   "Copy to File"
         End
         Begin VB.Menu mnuSCopyPointsInViewIncludeSearchResults 
            Caption         =   "Include DB Search Matches"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSCopyUMCsInView 
         Caption         =   "Copy LC-MS Features (UMCs) In View"
         Begin VB.Menu mnuSCopyPointsInViewByUMCtoClipboard 
            Caption         =   "Copy to Clipboard"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuSCopyPointsInViewByUMCtoFile 
            Caption         =   "Copy to File"
         End
         Begin VB.Menu mnuSCopyPointsInViewByUMCIncludeSearchResults 
            Caption         =   "Include DB Search Matches"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSCopyPoints 
         Caption         =   "Copy Points Options"
         Begin VB.Menu mnuSCopyPointsInViewToClipboardAuto 
            Caption         =   "Auto Copy Points on Zoom Change"
         End
         Begin VB.Menu mnuSCopyScansIncludeEmptyScans 
            Caption         =   "Include Empty Scans When Copying"
         End
         Begin VB.Menu mnuSCopyPointsInViewOneHitPerLine 
            Caption         =   "Copy Points Includes All Matches on One Line"
         End
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFS 
         Caption         =   "Find &Special"
         Begin VB.Menu mnuEditFSDenseAreas 
            Caption         =   "Areas Of Highest &Density"
         End
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "Clear"
         Begin VB.Menu mnuEditClearIDAll 
            Caption         =   "ID - &All"
         End
         Begin VB.Menu mnuEditClearIDNoId 
            Caption         =   "ID - With no DB reference"
         End
         Begin VB.Menu mnuEditClearIDBadDelta 
            Caption         =   "ID - &Bad Delta Pairs"
         End
         Begin VB.Menu mnuEditClearIDBadDeltaMT 
            Caption         =   "ID - &Bad Delta MT"
            Begin VB.Menu mnuEditClearIDBadDeltaMTN14N15 
               Caption         =   "&N14/N15"
            End
            Begin VB.Menu mnuEditClearIDBadDeltaMTICAT 
               Caption         =   "I&CAT"
            End
         End
         Begin VB.Menu mnuEditClearER 
            Caption         =   "&ER"
         End
      End
      Begin VB.Menu mnuEditSep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUMC 
         Caption         =   "LC-MS Features (&Unique Mass Classes)"
      End
      Begin VB.Menu mnuEditAvgUMC 
         Caption         =   "&Avg. UMC Member Masses"
      End
      Begin VB.Menu mnuEditDiscreteMWs 
         Caption         =   "Discrete &Molecular Masses"
      End
      Begin VB.Menu mnuEditCalibration 
         Caption         =   "&Calibration"
      End
      Begin VB.Menu mnuEditAbundance 
         Caption         =   "Abundances"
      End
      Begin VB.Menu mnuEditAdjScans 
         Caption         =   "Adjacent Scans"
      End
      Begin VB.Menu mnuEditNETFormula 
         Caption         =   "NET Formula"
      End
      Begin VB.Menu mnuEditNETAdj 
         Caption         =   "NET Adjustment"
         Begin VB.Menu mnuEditNETAdjustment 
            Caption         =   "Individual ID Peaks"
            Index           =   0
         End
         Begin VB.Menu mnuEditNETAdjustment 
            Caption         =   "UMC ID Peaks"
            Index           =   1
         End
         Begin VB.Menu mnuEditNETAdjustment 
            Caption         =   "Align and Warp Data using MSAlign"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEditResidualDisplay 
         Caption         =   "Init. Residual Display"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditParameters 
         Caption         =   "&Display Parameters and Paths"
      End
      Begin VB.Menu mnuEditComment 
         Caption         =   "Comm&ent"
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDatabaseConnection 
         Caption         =   "Select/Modify Database Connection"
      End
      Begin VB.Menu mnuEditMTStatus 
         Caption         =   "MT tag DB Stat&us"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Info"
      Begin VB.Menu mnuViewAnalysisHistory 
         Caption         =   "Analysis History &Log"
      End
      Begin VB.Menu mnuViewFileInfo 
         Caption         =   "&File Info"
         HelpContextID   =   107
      End
      Begin VB.Menu mnuIMTSettings 
         Caption         =   "MT tag S&ettings"
      End
      Begin VB.Menu mnuIAnalysisInfo 
         Caption         =   "MT tag DB &Analysis Info"
         Begin VB.Menu mnuIMTDBAnalysisInfo 
            Caption         =   "&Analysis Info"
         End
         Begin VB.Menu mnuIMTDBProcessingInfo 
            Caption         =   "&Processing Parameters"
         End
      End
      Begin VB.Menu mnuInfoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRawData 
         Caption         =   "&Raw Data"
         Begin VB.Menu mnuViewRDGel 
            Caption         =   "&Structured as in .GEL file"
         End
         Begin VB.Menu mnuViewRDPek 
            Caption         =   "Original .&PEK file"
         End
      End
      Begin VB.Menu mnuIDeviceCaps 
         Caption         =   "De&vice Caps"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      HelpContextID   =   107
      Begin VB.Menu mnuViewChargeState 
         Caption         =   "&Charge State"
         HelpContextID   =   107
      End
      Begin VB.Menu mnuViewIsotopic 
         Caption         =   "I&sotopic"
         HelpContextID   =   107
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewNormalDisplay 
         Caption         =   "&Normal View"
         Checked         =   -1  'True
         HelpContextID   =   107
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuViewDiffDisplay 
         Caption         =   "Co&mparative View"
         HelpContextID   =   107
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuViewCSMap 
         Caption         =   "Charge State Ma&p"
      End
      Begin VB.Menu mnuSep1234 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVUMC 
         Caption         =   "Unique Mass Classes"
      End
      Begin VB.Menu mnuViewSepOverlay 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVOverlay 
         Caption         =   "O&verlay"
      End
      Begin VB.Menu mnu2lsOverlaysManager 
         Caption         =   "Overlays Manager"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPI 
         Caption         =   "&pI View"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewFN 
         Caption         =   "&File (Scan) Number View"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewNET 
         Caption         =   "NET View"
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewZOrder 
         Caption         =   "Charge State on &Top"
         Checked         =   -1  'True
         HelpContextID   =   107
         Index           =   0
      End
      Begin VB.Menu mnuViewZOrder 
         Caption         =   "Isotopic on T&op"
         HelpContextID   =   107
         Index           =   1
      End
      Begin VB.Menu mnuViewSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTICandBPIPlots 
         Caption         =   "TIC and BPI Plots"
      End
      Begin VB.Menu mnuViewUMCBrowser 
         Caption         =   "LC-MS Feature (UMC) Browser"
      End
      Begin VB.Menu mnuViewPairsBrowser 
         Caption         =   "Pairs Browser"
      End
      Begin VB.Menu mnuViewDistributions 
         Caption         =   "&Distributions"
      End
      Begin VB.Menu mnuViewSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewZoomRegionListEditor 
         Caption         =   "Zoom Region List Editor"
      End
      Begin VB.Menu mnuVMTDisplay 
         Caption         =   "MT tags Display"
      End
      Begin VB.Menu mnuViewSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelToolBox 
         Caption         =   "Selection Calc&ulator"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Tool&bar"
         Checked         =   -1  'True
         HelpContextID   =   107
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuViewTracker 
         Caption         =   "Coor&dinates"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVMenuMode 
         Caption         =   "Menu Mode"
         Begin VB.Menu mnuVMenuModeSelect 
            Caption         =   "&Basic Menus"
            Index           =   0
         End
         Begin VB.Menu mnuVMenuModeSelect 
            Caption         =   "&DB Search Menus (no pairs)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuVMenuModeSelect 
            Caption         =   "DB Search Menus (with &pairs)"
            Index           =   2
         End
         Begin VB.Menu mnuVMenuModeSelect 
            Caption         =   "&Full Menus"
            Index           =   3
         End
         Begin VB.Menu mnuVMenuModeSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMenuModeIncludeObsolete 
            Caption         =   "Include Obsolete Menus"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      HelpContextID   =   108
      Begin VB.Menu mnu2lsFilter 
         Caption         =   "&Filter Points"
         HelpContextID   =   108
      End
      Begin VB.Menu mnu2lsFilterGraph 
         Caption         =   "Filter Points - Graph"
      End
      Begin VB.Menu mnu2lsFilterPointsByMass 
         Caption         =   "Filter Points by Mass (Auto-remove noise)"
      End
      Begin VB.Menu mnu2lsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2lsShowAll 
         Caption         =   "Show &All Points"
         HelpContextID   =   108
      End
      Begin VB.Menu mnu2lsShowUMCOnly 
         Caption         =   "Show &UMC Points"
      End
      Begin VB.Menu mnu2lsSel 
         Caption         =   "S&election"
         Begin VB.Menu mnu2lsAddVisiblePointsToSelected 
            Caption         =   "Add &Visible Points to Selected"
         End
         Begin VB.Menu mnu2lsLockSelection 
            Caption         =   "&Lock"
         End
         Begin VB.Menu mnu2lsClearSelection 
            Caption         =   "&Clear"
         End
         Begin VB.Menu mnu2lsSepSel1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu2lsExcludeSelection 
            Caption         =   "&Exclude Selected Points"
            Enabled         =   0   'False
            HelpContextID   =   108
         End
         Begin VB.Menu mnu2lsExcludeAllButSelection 
            Caption         =   "Exclude All &But Selection"
            Enabled         =   0   'False
            HelpContextID   =   108
         End
         Begin VB.Menu mnu2lsExcludeVisiblePoints 
            Caption         =   "Select and Exclude Visible Points"
            Shortcut        =   ^W
         End
         Begin VB.Menu mnu2lsSepSel2 
            Caption         =   "-"
         End
         Begin VB.Menu mnu2lsSelAvgMW 
            Caption         =   "Avg. &MW"
         End
         Begin VB.Menu mnu2lsSelAvgInt 
            Caption         =   "Avg. I&ntensity"
         End
         Begin VB.Menu mnu2lsSelAvgFit 
            Caption         =   "Avg.& Fit"
         End
         Begin VB.Menu mnu2lsSelAvgER 
            Caption         =   "Avg. E&R"
         End
      End
      Begin VB.Menu mnu2lsShowData 
         Caption         =   "Show &Data"
         Enabled         =   0   'False
         HelpContextID   =   108
         Begin VB.Menu mnu2lsShowAllData 
            Caption         =   "All &Data"
         End
         Begin VB.Menu mnu2lsShowCurrentData 
            Caption         =   "Current &Scope"
         End
         Begin VB.Menu mnu2lsShowSelected 
            Caption         =   "Selected &Point"
         End
      End
      Begin VB.Menu mnu2lsSepShowSpectra 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2lsShowSpectrum 
         Caption         =   "Show Mass &Spectrum"
         Begin VB.Menu mnu2lsShowSpectrumForLastSelectedPt 
            Caption         =   "for &Last Selected Point"
            HelpContextID   =   108
         End
         Begin VB.Menu mnu2lsShowSpectrumAllSelectedPts 
            Caption         =   "for All Selected Points (&Individual)"
         End
         Begin VB.Menu mnu2lsShowSpectrumSumSelectedPts 
            Caption         =   "for All Selected Points (&Sum)"
         End
      End
      Begin VB.Menu mnu2lsShowSpectrumNearestCursorPoint 
         Caption         =   "Show Mass Spectrum for Point Nearest Cursor"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu2lsCloseAllICR2LSMassSpectra 
         Caption         =   "Close all ICR-2LS Mass Spectra"
      End
      Begin VB.Menu mnu2lsSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2lsUMC2003 
         Caption         =   "Unique Mass Classes 2003 (faster)"
      End
      Begin VB.Menu mnu2lsUMCIonNet 
         Caption         =   "Unique Mass Classes Ion Net (better)"
      End
      Begin VB.Menu mnu2lsSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2lsSearchAMT 
         Caption         =   "Search &MT Tag Database"
         Begin VB.Menu mnu2lsSearchAMTOld 
            Caption         =   "&Legacy MT Database"
         End
         Begin VB.Menu mnu2lsSearchAMTNew 
            Caption         =   "&Org. MT tags Database"
         End
         Begin VB.Menu mnu2lsSearchUMCMassTags 
            Caption         =   "UMC Search - All points in UMC"
         End
         Begin VB.Menu mnu2lsSearchUMCSingleMass 
            Caption         =   "UMC Search - Single UMC Mass (preferred)"
         End
         Begin VB.Menu mnu2lsSepSearchPairs 
            Caption         =   "-"
         End
         Begin VB.Menu mnu2lsSearchUMCPairs 
            Caption         =   "UMC N14/N15 or O16/O18 &Pairs"
         End
         Begin VB.Menu mnu2lsSearchUMCPairs_ICAT 
            Caption         =   "UMC ICAT Pairs"
         End
         Begin VB.Menu mnu2lsSearchUMCPairs_PEON14N15 
            Caption         =   "UMC Labeled N14/N15 Pairs "
         End
      End
      Begin VB.Menu mnu2lsToleranceRefinement 
         Caption         =   "Mass Calibration and Tolerance Refinement"
      End
      Begin VB.Menu mnu2lsShowAMTRecord 
         Caption         =   "Sho&w MT Record"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu2lsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2lsORFCenteredSearchPRISM 
         Caption         =   "ORF-&Centered Search - PRISM"
      End
      Begin VB.Menu mnu2lsORFCenteredSearchFASTA 
         Caption         =   "ORF-Centered Search - &FASTA"
      End
      Begin VB.Menu mnu2LSSearchForORFs 
         Caption         =   "Search For ORFs"
      End
      Begin VB.Menu mnu2lsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2LsZoomIn 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu mnu2lsZoomOut 
         Caption         =   "Zoom &Out"
         HelpContextID   =   108
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu2lsZoomOutOneLevel 
         Caption         =   "Zoom Out One &Level"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu2lsSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2lsResetGraph 
         Caption         =   "&Reset Graph"
         HelpContextID   =   108
      End
      Begin VB.Menu mnu2lsSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2lsPEKFunctions 
         Caption         =   "PEK Functions"
         Begin VB.Menu mnu2LSMergePEKFiles 
            Caption         =   "Merge PEK Files"
         End
         Begin VB.Menu mnu2LSSplitPEKFiles 
            Caption         =   "Split PEK Files"
         End
      End
      Begin VB.Menu mnu2lsOptions 
         Caption         =   "O&ptions"
      End
   End
   Begin VB.Menu mnuS 
      Caption         =   "&Special"
      Begin VB.Menu mnuSLoadScope 
         Caption         =   "&Load Scope"
      End
      Begin VB.Menu mnuSMS_MSSearch 
         Caption         =   "&MS/MS Search"
      End
      Begin VB.Menu mnuSSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSUMCLockMass 
         Caption         =   "UMC Lock &Mass"
      End
      Begin VB.Menu mnuSLockMass 
         Caption         =   "PRISM Loc&k Mass"
      End
      Begin VB.Menu mnuSIntCalLockMass 
         Caption         =   "Internal &Calibration Lock Mass"
      End
      Begin VB.Menu mnuSSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSExportData 
         Caption         =   "Export Raw Data"
      End
      Begin VB.Menu mnuSExportResults 
         Caption         =   "Export &Results by Ion"
      End
      Begin VB.Menu mnuSSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSExcludeExc 
         Caption         =   "Exclude Except"
      End
      Begin VB.Menu mnuSDltLbl 
         Caption         =   "Delta &Label Pairs"
         Begin VB.Menu mnuSDlt_S 
            Caption         =   "Delta (Individual)"
         End
         Begin VB.Menu mnuSDltUMC 
            Caption         =   "Delta (UMC)"
         End
         Begin VB.Menu mnuSLbl_S 
            Caption         =   "Label (Individual)"
         End
         Begin VB.Menu mnuSLblUMC 
            Caption         =   "Label (UMC)"
         End
         Begin VB.Menu mnuSDltLbl_S 
            Caption         =   "Delta-Label (Individual)"
         End
         Begin VB.Menu mnuSDltLblUMC 
            Caption         =   "Delta-Label (UMC)"
         End
         Begin VB.Menu mnuSERAnalysis 
            Caption         =   "ER &Analysis"
         End
         Begin VB.Menu mnuSDltLblReport 
            Caption         =   "&Report"
         End
      End
      Begin VB.Menu mnuSAttentionList 
         Caption         =   "Attention List"
      End
      Begin VB.Menu mnuSSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplitUMCs 
         Caption         =   "Split UMC's by Examining Abundances"
      End
      Begin VB.Menu mnuShowSplitUMCs 
         Caption         =   "Show Split UMC's"
      End
      Begin VB.Menu mnuSSepShow 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowNETAdjUMCs 
         Caption         =   "Show UMC's Used for NET Adjustment"
      End
      Begin VB.Menu mnuShowLowSegmentCountUMCs 
         Caption         =   "Show Low Segment Count UMC's Added"
      End
      Begin VB.Menu mnuShowNETAdjUMCsWithDBHit 
         Caption         =   "Show NET Adjustment UMC's with a DB Hit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      HelpContextID   =   109
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileH 
         Caption         =   "Tile &Horizontally"
         HelpContextID   =   109
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "Tile &Vertically"
         HelpContextID   =   109
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
         HelpContextID   =   109
      End
      Begin VB.Menu mnuArangeIcons 
         Caption         =   "&Arrange Icons"
         HelpContextID   =   109
      End
      Begin VB.Menu mnuWindowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowSizeToDim 
         Caption         =   "Size to fill a Powerpoint Slide"
         Index           =   0
      End
      Begin VB.Menu mnuWindowSizeToDim 
         Caption         =   "Size to 640x480"
         Index           =   1
      End
      Begin VB.Menu mnuWindowSizeToDim 
         Caption         =   "Size to 800x600"
         Index           =   2
      End
      Begin VB.Menu mnuWindowSizeToDim 
         Caption         =   "Size to 1024x768"
         Index           =   3
      End
      Begin VB.Menu mnuWindowSizeToDim 
         Caption         =   "Size to 1280x1024"
         Index           =   4
      End
      Begin VB.Menu mnuWindowSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowSizeUpdateAll 
         Caption         =   "Resize all windows"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   110
      Begin VB.Menu mnuVisual2DGelHelp 
         Caption         =   "&Help Topics"
         HelpContextID   =   110
      End
      Begin VB.Menu mnuHelpSetTraceLogLevel 
         Caption         =   "Set Trace Log Level"
      End
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         HelpContextID   =   110
      End
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Last modified: 01/29/2003 nt
'----------------------------------------------------------------------------
Option Explicit

Const ConvLPDP = 1
Const ConvDPLP = 2

' MonroeMod
Const REFRESH_TIMER_INTERVAL = 250       ' Minimum time between actual screen refreshes, in milliseconds

Private Enum enaEditNETAdjustmentMenuConstants
    enaIndividualPoints = 0
    enaUMCIterative = 1
    enaMSAlign = 2
End Enum

Private Enum namStepsNETAdjustmentModeConstants
    namMSAlign = 0
    namUMCIterative = 1
End Enum

Private Type udtRawSpectraDisplayStatsType
    ScanNumber As Long
    ChargeState As Integer
    TargetMZ As Double
    IsMoverZ As Boolean
    VisibleMZMinimum As Double
    VisibleMZMaximum As Double
End Type

Dim bLoad As Boolean    'true if loading of form
Dim bResize As Boolean  'if True allow resize event
Dim bPaint As Boolean   'if True allow paint event
Dim nMyIndex As Long
Dim OldMeScaleW As Long
Dim OldMeScaleH As Long

' MonroeMod: New Variable
Private bNeedToUpdate As Boolean
Private mAutoCopyPointsMaxCount As Long
Public mFileSaveMode As fsFileSaveModeConstants

Private lAction As Long
Private MouseButton As Integer
Private HotID As Long           'hot spot ID
Private HotType As Integer      'hot spot type
Private LastID As Long          'last selected point
Private LastType As Integer     'last selected point type

Private mLastCursorPosFN As Long
Private mLastCursorPosMass As Double

Public fgDisplay As Integer         'public properties
Public fgZOrder As Integer
Public fgSelBoxVisible As Boolean
Public fgSelProtected As Boolean    'when locked it can't be cleared
'''Public fgLMOnFreqShift As Boolean   'when set Lock Mass function locks
'''                                   'directly on values found in DFFS array


Public fgRebuild As Boolean  'when activated if this property
                             'is True rebuild graph

Dim gbZoomX1 As Double      'private variables to help with
Dim gbZoomY1 As Double      'zooming, clipping, and mouse
Dim gbZoomX2 As Double      'tracking features
Dim gbZoomY2 As Double

'''Dim guShiftX As Single
'''Dim guShiftY As Single

Dim paPoints() As POINTAPI  'used to calculate Dev to Log coordinates

'coordinate system variable for the gel
Public WithEvents csMyCooSys As CooSysR
Attribute csMyCooSys.VB_VarHelpID = -1
'selection object for the gel
Public WithEvents GelSel As Sel
Attribute GelSel.VB_VarHelpID = -1

Private Sub AddVisiblePointsToSelection()
    
    Dim lngIonPointerArray() As Long           ' 1-based array
    Dim lngIonCount As Long
    Dim lngIndex As Long
    
    ' Retrieve an array of the ion indices of the ions currently "In Scope"
    ' Note that GetISScope will ReDim lngIonPointerArray() automatically
    lngIonCount = GetISScope(nMyIndex, lngIonPointerArray(), glScope.glSc_Current)
    
    For lngIndex = 1 To lngIonCount
        GelSel.AddToIsoSelection lngIonPointerArray(lngIndex)
    Next lngIndex
    
    bNeedToUpdate = True
End Sub

Private Sub ChooseNewAnalysisUsingDMS()
    On Error GoTo NewAnalysisErrorHandler
    Set MDIForm1.MyAnalysisInit = New AnalysisInitiator
    MDIForm1.MyAnalysisInit.GetNewAnalysisDialog glInitFile

    Exit Sub
    
NewAnalysisErrorHandler:
    MsgBox "Error initiating a new analysis of a file on DMS (" & Err.Number & "):" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub ClearSelectedPoints(Optional ByVal blnRefreshPlot As Boolean = True)
    GelSel.Clear
    If fgSelBoxVisible Then SelToolBox1.ClearResults
    If blnRefreshPlot Then picGraph.Refresh
End Sub

Private Sub csMyCooSys_CooSysChanged()
picGraph.Refresh
UpdateTICPlotAndFeatureBrowsersIfNeeded
End Sub

Private Sub ExcludeVisiblePointsOneStep()
    ClearSelectedPoints False
    AddVisiblePointsToSelection
    ExcludeSelection nMyIndex
End Sub

Public Sub GetCurrentZoomArea(lngXMin As Long, lngXMax As Long, dblYMin As Double, dblYMax As Double)
    
    On Error GoTo GetCurrentZoomAreaErrorHandler
    
    lngXMin = csMyCooSys.CurrRXMin
    lngXMax = csMyCooSys.CurrRXMax
    dblYMin = csMyCooSys.CurrRYMin
    dblYMax = csMyCooSys.CurrRYMax
    Exit Sub

GetCurrentZoomAreaErrorHandler:
    ' Error: csMyCooSys probably doesn't exist
    Debug.Assert False
    
End Sub

Private Sub HandleSelToolboxClick(ByVal eFormat As ssrfSelectionStatsResultFormatConstants)
    Dim CSField As Integer
    Dim IsoField As Integer
    Dim udtResults As GelRes
    Dim NumFormat As String
    
    NumFormat = GetNumFormat(4)
    Select Case SelToolBox1.DataField
    Case glSEL_DF_MW
        SelStatsCompute nMyIndex, glFIELD_MW, NumFormat, eFormat, udtResults
    Case glSEL_DF_INTENSITY
        SelStatsCompute nMyIndex, glFIELD_ABU, NumFormat, eFormat, udtResults
    Case glSEL_DF_FIT
        SelStatsCompute nMyIndex, glFIELD_FIT, NumFormat, eFormat, udtResults
    Case glSEL_DF_ER
        SelStatsCompute nMyIndex, glFIELD_ER, NumFormat, eFormat, udtResults
    Case Else
        Debug.Assert False
    End Select
    SelToolBox1.SetResults udtResults.CSRes, udtResults.IsoRes, udtResults.AllRes
    SelToolBox1.Refresh

End Sub

Private Sub LoadNewDataFile()
    MDIStatus True, "Loading ... please be patient"
    FileNew (Me.hwnd)
End Sub

Private Sub csMyCooSys_MWScaleChange()
'logarithmic scale might need to be initialized
If csMyCooSys.csYScale = glVAxisLog Then
   With GelDraw(nMyIndex)
        If .CSCount > 0 Then
           If IsArrayEmpty(.CSLogMW()) Then
              ReDim .CSLogMW(1 To .CSCount)
              InitDrawCSLogMW (nMyIndex)
           End If
        End If
        If .IsoCount > 0 Then
           If IsArrayEmpty(.IsoLogMW()) Then
              ReDim .IsoLogMW(1 To .IsoCount)
              InitDrawIsoLogMW (nMyIndex)
           End If
        End If
   End With
End If
csMyCooSys.InitFNType
csMyCooSys.CoordinateDraw
picGraph.Refresh
End Sub

Public Sub SetXAxisLabelType(blnNETLabels As Boolean)

    mnuViewPI.Checked = False
    If blnNETLabels Then
        mnuViewFN.Checked = False
        mnuViewNET.Checked = True
        If Not csMyCooSys Is Nothing Then csMyCooSys.csType = glNETCooSys
        GelData(nMyIndex).Preferences.CooType = glNETCooSys
    Else
        mnuViewFN.Checked = True
        mnuViewNET.Checked = False
        If Not csMyCooSys Is Nothing Then csMyCooSys.csType = glFNCooSys
        GelData(nMyIndex).Preferences.CooType = glFNCooSys
    End If
    
    picGraph.Refresh

End Sub

Public Sub SetWindowSize(eWindowSize As wscWindowSizeConstants, Optional blnUpdatingAllGels As Boolean = False)
    Dim lngWidthPixels As Long, lngHeightPixels As Long
    Dim lngGelIndex As Long
    
    Select Case eWindowSize
    Case wscsize640by480
        lngWidthPixels = 640
        lngHeightPixels = 480
    Case wscSizeForPowerpoint
        lngWidthPixels = 880
        lngHeightPixels = 660
    Case wscSize1024by768
        lngWidthPixels = 1024
        lngHeightPixels = 768
    Case wscSize1280by1024
        lngWidthPixels = 1280
        lngHeightPixels = 1024
    Case Else
        ' Size to 800 by 600
        ' Includes wscSize800by600
        lngWidthPixels = 800
        lngHeightPixels = 600
    End Select
    
    On Error GoTo SetWindowSizeErrorHandler
    With Me
        If .WindowState = vbNormal Then
            .ScaleMode = vbTwips
            .width = lngWidthPixels * Screen.TwipsPerPixelX
            .Height = lngHeightPixels * Screen.TwipsPerPixelY
        End If
    End With
    
    If mnuWindowSizeUpdateAll.Checked And Not blnUpdatingAllGels Then
        For lngGelIndex = 1 To UBound(GelBody())
            If lngGelIndex <> nMyIndex And Not GelStatus(lngGelIndex).Deleted Then
                GelBody(lngGelIndex).SetWindowSize eWindowSize, True
            End If
        Next lngGelIndex
    End If
    
    With MDIForm1
        If .width < Me.width * 1.05 Then
            .width = Me.width * 1.05
        End If
        
        If .Height < Me.Height * 1.05 + 1000 Then
            .Height = Me.Height * 1.05 + 1000
        End If
    End With
    
    Exit Sub
    
SetWindowSizeErrorHandler:
    Debug.Print "Error in SetWindowSize: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmGraph->SetWindowSize"
End Sub

Private Sub ShowErrorDistribution2DForm()
    On Error Resume Next
    
    If GelUMC(nMyIndex).UMCCnt <= 0 Then
        MsgBox "You must cluster the data into Unique Mass Classes before running Mass Calibration and Tolerance Refinement.  Please use menu item 'Steps->2. Find UMCs' to accomplish this.", vbInformation + vbOKOnly, "No UMCs"
    Else
        frmErrorDistribution2DLoadedData.CallerID = nMyIndex
        frmErrorDistribution2DLoadedData.Show vbModal
    End If
End Sub

Private Sub ShowFilterForm()
    lAction = glNoAction
    frmFilter.Tag = nMyIndex
    frmFilter.Show
    picGraph.Refresh
End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuNewAnalysis.Visible = blnVisible
    mnuStepsFileNewChooseFromDMS.Visible = blnVisible
End Sub

Private Sub ShowMSAlignForm()
    Dim objMSAlign As frmMSAlign

    If APP_BUILD_DISABLE_ADVANCED Then
        MsgBox "Function not enabled in this version; action cancelled.", vbExclamation + vbOKOnly, "Not enabled"
        Exit Sub
    End If

    Set objMSAlign = New frmMSAlign
    objMSAlign.CallerID = nMyIndex

On Error GoTo ShowMSAlignFormErrorHandler
    objMSAlign.Show vbModal
    Set objMSAlign = Nothing
    
    Exit Sub
    
ShowMSAlignFormErrorHandler:
    Debug.Assert False
    Set objMSAlign = Nothing
End Sub

Private Function ShowMSSpectrum(ByVal lngScanNumber As Long, ByVal blnIsMoverZ As Boolean, ByVal dblTargetMZ As Double, ByRef hScope As Integer, ByVal dblVisibleMZMinimum As Double, ByVal dblVisibleMZMaximum As Double, Optional blnInformUserOnError As Boolean = True) As Boolean
    Dim blnSuccess As Boolean
    
    Select Case GelStatus(nMyIndex).SourceDataRawFileType
    Case rfcZippedSFolders
        blnSuccess = ICR2LSLoadSpectrumViaCache(nMyIndex, lngScanNumber, blnIsMoverZ, dblTargetMZ, hScope, dblVisibleMZMinimum, dblVisibleMZMaximum)
    Case rfcFinniganRaw
        blnSuccess = ICR2LSLoadFinniganSpectrum(nMyIndex, lngScanNumber, dblTargetMZ, hScope, dblVisibleMZMinimum, dblVisibleMZMaximum)
    Case Else
        If blnInformUserOnError Then
            MsgBox "Unable to determine the raw data file type. " & vbCrLf & "Use 'Edit->Display Parameters and Paths' to enter a valid path to the folder containing the Finnigan .Raw file or zipped S-Folders.", vbExclamation + vbOKOnly, "Error"
        End If
        hScope = -1
        blnSuccess = False
    End Select

    ShowMSSpectrum = blnSuccess
End Function

Private Sub ShowMSSpectrumForAllSelected(ByVal blnSumSpectra As Boolean)

    ' Construct a unique FN list for all selected points, then show each
    ' spectrum individually, or sum if blnSumSpectra = True (can only sum spectra in ICR-2LS)
    
    Const MAX_SPECTRA_TO_SHOW As Long = 15
    
    Dim lngSpecCount As Long
    Dim udtSpectra() As udtRawSpectraDisplayStatsType
    
    Dim objQS As QSDouble
    Dim dblSortKey() As Double
    Dim lngPointerArray() As Long
    
    Dim lngSpecCountUnique As Long
    Dim udtSpectraUnique() As udtRawSpectraDisplayStatsType
    
    Dim lngIndex As Long
    Dim lngPointerIndex As Long             ' Pointer into GelData().IsoData or .CSData
   
    Dim eResponse As VbMsgBoxResult
    Dim blnSuccess As Boolean
    
    Dim hScope As Integer
    Dim hScopeBase As Integer
    
On Error GoTo ShowSpectrumErrorHandler

    DetermineSourceDataRawFileType nMyIndex, False
    
    lngSpecCount = 0
    ReDim udtSpectra(0)
    
    If GelBody(nMyIndex).GelSel.CSSelCnt > 0 Then
        ReDim udtSpectra(GelBody(nMyIndex).GelSel.CSSelCnt - 1)
        
        For lngIndex = 1 To GelBody(nMyIndex).GelSel.CSSelCnt
            lngPointerIndex = GelBody(nMyIndex).GelSel.Value(lngIndex, glCSType)
            
            With udtSpectra(lngSpecCount)
                .ScanNumber = GelData(nMyIndex).CSData(lngPointerIndex).ScanNumber
                .ChargeState = 1
                .IsMoverZ = False
                .TargetMZ = GelData(nMyIndex).CSData(lngPointerIndex).AverageMW
                .VisibleMZMinimum = .TargetMZ - glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
                .VisibleMZMaximum = .TargetMZ + glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
            End With
            
            lngSpecCount = lngSpecCount + 1
            
        Next lngIndex
    End If
    
    If GelBody(nMyIndex).GelSel.IsoSelCnt > 0 Then
        ReDim Preserve udtSpectra(lngSpecCount + GelBody(nMyIndex).GelSel.IsoSelCnt - 1)
        
        For lngIndex = 1 To GelBody(nMyIndex).GelSel.IsoSelCnt
           lngPointerIndex = GelBody(nMyIndex).GelSel.Value(lngIndex, glIsoType)
        
            With udtSpectra(lngSpecCount)
                .ScanNumber = GelData(nMyIndex).IsoData(lngPointerIndex).ScanNumber
                .ChargeState = GelData(nMyIndex).IsoData(lngPointerIndex).Charge
                .IsMoverZ = GelData(nMyIndex).Preferences.IsoICR2LSMOverZ
                If .IsMoverZ Then
                   .TargetMZ = GelData(nMyIndex).IsoData(lngPointerIndex).MZ    'use m/z
                Else
                   .TargetMZ = GelData(nMyIndex).IsoData(lngPointerIndex).MostAbundantMW     'use the most abundant mass
                End If
                
                If Not ShowMSSpectrumGetViewRange(.ScanNumber, GelData(nMyIndex).IsoData(lngPointerIndex).MonoisotopicMW, 0, .VisibleMZMinimum, .VisibleMZMaximum) Then
                    ' No points found, use default zoom window width
                    .VisibleMZMinimum = .TargetMZ - glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
                    .VisibleMZMaximum = .TargetMZ + glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
                End If
            End With
            
            lngSpecCount = lngSpecCount + 1
            
        Next lngIndex
    End If

    
    If lngSpecCount > 1 Then
        ' Remove scan number redundancy from udtSpectra
        
        ReDim dblSortKey(lngSpecCount - 1)
        ReDim lngPointerArray(lngSpecCount - 1)
        
        For lngIndex = 0 To lngSpecCount - 1
            dblSortKey(lngIndex) = udtSpectra(lngIndex).ScanNumber + udtSpectra(lngIndex).ChargeState / 10
            lngPointerArray(lngIndex) = lngIndex
        Next lngIndex
        
        Set objQS = New QSDouble
        If objQS.QSAsc(dblSortKey, lngPointerArray) Then
            ReDim udtSpectraUnique(lngSpecCount - 1)
            
            udtSpectraUnique(0) = udtSpectra(lngPointerArray(0))
            lngSpecCountUnique = 1

            For lngIndex = 1 To lngSpecCount - 1
                With udtSpectra(lngPointerArray(lngIndex - 1))
                    If udtSpectra(lngPointerArray(lngIndex)).ScanNumber <> .ScanNumber Then
                        udtSpectraUnique(lngSpecCountUnique) = udtSpectra(lngPointerArray(lngIndex))
                        lngSpecCountUnique = lngSpecCountUnique + 1
                    ElseIf Not blnSumSpectra Then
                        ' Keep the second spectrum if it has a different charge state,
                        '  but not if we're summing spectra
                        If udtSpectra(lngPointerArray(lngIndex)).ChargeState <> .ChargeState Then
                            udtSpectraUnique(lngSpecCountUnique) = udtSpectra(lngPointerArray(lngIndex))
                            lngSpecCountUnique = lngSpecCountUnique + 1
                        End If
                    End If
                End With
            Next lngIndex
            
            If lngSpecCount <> lngSpecCountUnique Then
                ' Copy the data back into udtSpectra
                
                ReDim Preserve udtSpectraUnique(lngSpecCountUnique - 1)
                lngSpecCount = lngSpecCountUnique
                udtSpectra = udtSpectraUnique
            End If
        Else
            Debug.Assert False
        End If
    End If
    
    If blnSumSpectra And lngSpecCount > 25 Then
        eResponse = MsgBox("You have selected data points from " & Trim(lngSpecCount) & " distinct spectra (scans).  Summing this many spectra will take awhile.  Are you sure you want to continue?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Large Number of Spectra")
        If eResponse <> vbYes Then Exit Sub
    End If

    If lngSpecCount > 1 Then
        frmProgress.InitializeForm "Loading spectra", 0, lngSpecCount, True
        frmProgress.MoveToBottomCenter
    End If
    
    If blnSumSpectra Then
        With udtSpectra(0)
            hScope = -1
            ShowMSSpectrum .ScanNumber, .IsMoverZ, .TargetMZ, hScope, .VisibleMZMinimum, .VisibleMZMaximum
        End With
        
        ' Note: hScope will be -1 if things didn't work right
        If hScope >= 0 Then
            hScopeBase = hScope
            For lngIndex = 1 To lngSpecCount - 1
                With udtSpectra(lngIndex)
                    hScope = -1
                    blnSuccess = ShowMSSpectrum(.ScanNumber, .IsMoverZ, .TargetMZ, hScope, .VisibleMZMinimum, .VisibleMZMaximum, False)
                    DoEvents
                End With
                
                objICR2LS.ScopeMath hScopeBase, hScope, "+"
                objICR2LS.KillScope = hScope
                
                frmProgress.UpdateProgressBar lngIndex + 1
                DoEvents
                If KeyPressAbortProcess > 1 Or Not blnSuccess Then Exit For
            Next lngIndex
        
        
'            Dim intStatus As Integer
'            Dim strComment As String
'
'            ' ToDo: Implement this
'            Debug.Assert False
'
'            'objICR2LS.ClearScopeComments hScopeBase
'            intStatus = objICR2LS.PeakTransform(hScopeBase, Round(udtSpectra(0).TargetMZ, 6))
'            If intStatus = 0 Then
'                strComment = objICR2LS.GetScopeComments(hScopeBase)
'            End If
        
        End If
    Else
        If lngSpecCount > MAX_SPECTRA_TO_SHOW Then
            MsgBox "You have selected data points from " & Trim(lngSpecCount) & " distinct spectra (scans), but ICR-2LS can only display at most " & MAX_SPECTRA_TO_SHOW & " spectra.  The number of spectra will therefore be limited to " & MAX_SPECTRA_TO_SHOW, vbInformation + vbOKOnly, "Too many spectra"
            lngSpecCount = MAX_SPECTRA_TO_SHOW
        End If
        
        For lngIndex = 0 To lngSpecCount - 1
            With udtSpectra(lngIndex)
                blnSuccess = ShowMSSpectrum(.ScanNumber, .IsMoverZ, .TargetMZ, hScope, .VisibleMZMinimum, .VisibleMZMaximum, True)
            End With
        
            If lngSpecCount > 1 Then
                frmProgress.UpdateProgressBar lngIndex + 1
                DoEvents
                If KeyPressAbortProcess > 1 Or Not blnSuccess Then Exit For
            End If
        
        Next lngIndex
    End If
    
    If lngSpecCount > 1 Then frmProgress.HideForm
    
    Exit Sub
    
ShowSpectrumErrorHandler:
    MsgBox "Error displaying spectra for all selected points " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    
End Sub

Private Sub ShowMSSpectrumForPointNearestCursor()

    Dim lngScanNumber As Long
    Dim lngMaxScanNumber As Long
    
    Dim dblTargetMonoIsoMass As Double
    Dim dblTargetMZ As Double
    Dim intCharge As Integer
    
    Dim dblVisibleMZMinimum As Double
    Dim dblVisibleMZMaximum As Double
    
    Dim blnIsMoverZ As Boolean
    
    Dim hScope As Integer
    
On Error GoTo ShowSpectrumErrorHandler
            
    blnIsMoverZ = False
    lngScanNumber = mLastCursorPosFN
    dblTargetMonoIsoMass = mLastCursorPosMass
    
    With GelData(nMyIndex)
        If lngScanNumber < .ScanInfo(1).ScanNumber Then lngScanNumber = .ScanInfo(1).ScanNumber
        
        lngMaxScanNumber = .ScanInfo(UBound(.ScanInfo)).ScanNumber
        If lngScanNumber > lngMaxScanNumber Then lngScanNumber = lngMaxScanNumber
    End With
    
    ' Find the nearest, valid scan number
    If LookupScanNumberRelativeIndex(nMyIndex, lngScanNumber) = 0 Then
        ' lngScanNumber does not represent a scan with data
        ' Look for the nearest scan that actually contains data
        
        lngScanNumber = LookupScanNumberClosest(nMyIndex, lngScanNumber)
    End If

    If lngScanNumber = 0 Then
        MsgBox "Unable to determine the scan number that the cursor is nearest to (" & CStr(lngScanNumber) & " isn't a valid scan number)", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If dblTargetMonoIsoMass < 1 Then dblTargetMonoIsoMass = 1000
    
    ' Need to determine most appropriate charge, then compute the target M/Z value based on the currently visible zoom region
    
    If ShowMSSpectrumGetViewRange(lngScanNumber, dblTargetMonoIsoMass, intCharge, dblVisibleMZMinimum, dblVisibleMZMaximum) Then
        blnIsMoverZ = True
        dblTargetMZ = ConvoluteMass(dblTargetMonoIsoMass, 1, intCharge)
    End If
    
    DetermineSourceDataRawFileType nMyIndex, False
    
    ShowMSSpectrum lngScanNumber, blnIsMoverZ, dblTargetMZ, hScope, dblVisibleMZMinimum, dblVisibleMZMaximum
        
    Exit Sub

ShowSpectrumErrorHandler:
    MsgBox "Error displaying spectrum for scan nearest cursor (scan " & Trim(lngScanNumber) & " and m/z " & Round(dblTargetMZ, 3) & ")" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"

End Sub

Private Sub ShowMSSpectrumForLastID()

    Dim lngScanNumber As Long
    Dim dblTargetMZ As Double
    Dim blnIsMoverZ As Boolean

    Dim dblVisibleMZMinimum As Double
    Dim dblVisibleMZMaximum As Double

    Dim hScope As Integer
    
On Error GoTo ShowSpectrumErrorHandler

    With GelData(nMyIndex)
        Select Case LastType
        Case glNoType
             Exit Sub
        Case glCSType
            lngScanNumber = .CSData(LastID).ScanNumber
            
            blnIsMoverZ = False
            'dblTargetMZ = .CSData(LastID).AverageMW / .CSData(LastID).Charge + glMASS_CC
            dblTargetMZ = .CSData(LastID).AverageMW
            
            dblVisibleMZMinimum = dblTargetMZ - glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
            dblVisibleMZMaximum = dblTargetMZ + glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
            
        Case glIsoType
             lngScanNumber = .IsoData(LastID).ScanNumber
             
             blnIsMoverZ = .Preferences.IsoICR2LSMOverZ
             If blnIsMoverZ Then
                dblTargetMZ = .IsoData(LastID).MZ    'use m/z
             Else
                dblTargetMZ = .IsoData(LastID).MostAbundantMW     'use the most abundant mass
             End If
             
            If Not ShowMSSpectrumGetViewRange(lngScanNumber, .IsoData(LastID).MonoisotopicMW, 0, dblVisibleMZMinimum, dblVisibleMZMaximum) Then
                ' No points found, use default zoom window width
                dblVisibleMZMinimum = dblTargetMZ - glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
                dblVisibleMZMaximum = dblTargetMZ + glbPreferencesExpanded.ICR2LSSpectrumViewZoomWindowWidthMZ
            End If
             
        End Select
    End With
    
    DetermineSourceDataRawFileType nMyIndex, False
    
    ShowMSSpectrum lngScanNumber, blnIsMoverZ, dblTargetMZ, hScope, dblVisibleMZMinimum, dblVisibleMZMaximum
    
    Exit Sub
    
ShowSpectrumErrorHandler:
    MsgBox "Error displaying spectrum for scan " & Trim(lngScanNumber) & " and m/z " & Round(dblTargetMZ, 3) & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"

End Sub

Private Function ShowMSSpectrumGetViewRange(ByVal lngScanNumber As Long, ByVal dblTargetMonoIsoMass As Double, ByRef intChargeClosestPoint As Integer, ByRef dblVisibleMZMinimum As Double, ByRef dblVisibleMZMaximum As Double) As Boolean
    ' Determine the charge of the data point nearest lngScanNumber and dblTargetMonoIsoMass
    ' Additionally, determine the mass range of the data in view
    ' Convolute the mass range using the charge, and return the values in dblVisibleMZMinimum and dblVisibleMZMaximum
    '
    ' Returns true if data was found in range
    
    Dim lngIonCount As Long
    Dim lngIonPointerArray() As Long
    
    Dim lngBestIndex As Long        ' Index in the .IsoData() array
    Dim dblBestDistance As Double
    
    Dim dblDistance As Double
    Dim lngIndex As Long

    Dim dblMWMinimum As Double      ' Monoisotopic mass minimum
    Dim dblMWMaximum As Double      ' Monoisotopic mass maximum
    Dim dblCorrection As Double
    
    Dim dblCurrentViewMassMin As Double
    Dim dblCurrentViewMassMax As Double
    
    ' Find the datapoint closest to lngScanNumber and dblTargetMonoIsoMass
    lngIonCount = GetISScope(nMyIndex, lngIonPointerArray(), glScope.glSc_Current)

    If lngIonCount > 0 Then
        
        dblBestDistance = 1E+307
        lngBestIndex = lngIonPointerArray(1)
        
        With GelData(nMyIndex)
            dblMWMinimum = .IsoData(lngBestIndex).MonoisotopicMW
            dblMWMaximum = .IsoData(lngBestIndex).MonoisotopicMW
            
            For lngIndex = 1 To lngIonCount
                dblDistance = Sqr((.IsoData(lngIonPointerArray(lngIndex)).ScanNumber - lngScanNumber) ^ 2 + (.IsoData(lngIonPointerArray(lngIndex)).MonoisotopicMW - dblTargetMonoIsoMass) ^ 2)
                If dblDistance < dblBestDistance Then
                    dblBestDistance = dblDistance
                    lngBestIndex = lngIonPointerArray(lngIndex)
                    If .IsoData(lngBestIndex).MonoisotopicMW < dblMWMinimum Then
                        dblMWMinimum = .IsoData(lngBestIndex).MonoisotopicMW
                    End If
                
                    If .IsoData(lngBestIndex).MonoisotopicMW > dblMWMaximum Then
                        dblMWMaximum = .IsoData(lngBestIndex).MonoisotopicMW
                    End If
                End If
            Next lngIndex
            
            ' Determine the currently zoomed view range
            csMyCooSys.GetVisibleDimensions 0, 0, dblCurrentViewMassMin, dblCurrentViewMassMax

            If dblMWMinimum < dblCurrentViewMassMin Then
                ' This shouldn't normally occur; if it does, then the
                '  coordinate system is out of sync with the visible data
                Debug.Assert False
                dblCurrentViewMassMin = dblMWMinimum
            End If
            
            If dblMWMaximum > dblCurrentViewMassMax Then
                ' This shouldn't normally occur; if it does, then the
                '  coordinate system is out of sync with the visible data
                Debug.Assert False
                dblCurrentViewMassMax = dblMWMaximum
            End If
                    
            ' Determine the charge for the data point at lngBestIndex
            intChargeClosestPoint = .IsoData(lngBestIndex).Charge
            
            dblVisibleMZMinimum = ConvoluteMass(dblCurrentViewMassMin, 1, intChargeClosestPoint)
            dblVisibleMZMaximum = ConvoluteMass(dblCurrentViewMassMax, 1, intChargeClosestPoint)
            
            ' Require a minimum view size of 5 m/z
            dblCorrection = 5 - (dblVisibleMZMaximum - dblVisibleMZMinimum)
            If dblCorrection > 0 Then
                dblVisibleMZMinimum = dblVisibleMZMinimum - 1 * dblCorrection / 5
                dblVisibleMZMaximum = dblVisibleMZMaximum + 4 * dblCorrection / 5
            End If
        End With
        
        ShowMSSpectrumGetViewRange = True
    Else
        ShowMSSpectrumGetViewRange = False
    End If

End Function

Private Sub ShowOrganizeDatabaseConnectionsForm()
    frmOrganizeDBConnections.Tag = nMyIndex
    frmOrganizeDBConnections.InitializeForm
    frmOrganizeDBConnections.Show vbModeless, MDIForm1
End Sub

Private Sub ShowPairsDeltaUMCForm()
    On Error Resume Next
    frmUMCDltPairs.Tag = nMyIndex
    frmUMCDltPairs.FormMode = pfmDelta
    frmUMCDltPairs.Show vbModal
    UpdateTICPlotAndFeatureBrowsersIfNeeded True
End Sub

Private Sub ShowPairsDeltaLabelUMCForm()
    On Error Resume Next
    frmUMCDltPairs.Tag = nMyIndex
    frmUMCDltPairs.FormMode = pfmDeltaLabel
    frmUMCDltPairs.Show vbModal
    UpdateTICPlotAndFeatureBrowsersIfNeeded True
End Sub

Private Sub ShowPairsLabelUMCForm()
    On Error Resume Next
    frmUMCDltPairs.Tag = nMyIndex
    frmUMCDltPairs.FormMode = pfmLabel
    frmUMCDltPairs.Show vbModal
    UpdateTICPlotAndFeatureBrowsersIfNeeded True
End Sub

Private Sub ShowSearchMTDBUMCPairsN14N15()
    '----------------------------------------------
    'displays search of MT tag database form for
    'UMC N14/N15 Pairs
    '----------------------------------------------
    Dim UMCPairsSearch As New frmSearchMTPairs
    On Error Resume Next
    If GelP_D_L(nMyIndex).PCnt > 0 Then
       UMCPairsSearch.CallerID = nMyIndex
       UMCPairsSearch.Show vbModal
       Set UMCPairsSearch = Nothing
    Else
       MsgBox "Paired UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes then use menu item 'Steps->4. Find Pairs' to find paired UMCs.", vbOKOnly, glFGTU
    End If
End Sub

Private Sub ShowSearchMTDBUMCPairsPEON14N15()
    '--------------------------------------------------------------
    'displays search of MT tag DB form for UMC PEO N14/N15 Pairs
    '--------------------------------------------------------------
    Dim UMCPairsSearch As New frmSearchMTPairs_PEO
    On Error Resume Next
    If GelP_D_L(nMyIndex).PCnt > 0 Then
       UMCPairsSearch.CallerID = nMyIndex
       UMCPairsSearch.Show vbModal
       Set UMCPairsSearch = Nothing
    Else
       MsgBox "Paired UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes then use menu item 'Steps->4. Find Pairs' to find paired UMCs.", vbOKOnly, glFGTU
    End If
End Sub

Private Sub ShowSearchMTDBUMCPairsICAT()
    '-------------------------------------------------------
    'displays search of MT tag DB form for UMC ICAT Pairs
    '-------------------------------------------------------
    Dim UMCPairsSearch As New frmSearchMTPairs_ICAT
    On Error Resume Next
    If GelP_D_L(nMyIndex).PCnt > 0 Then
       UMCPairsSearch.CallerID = nMyIndex
       UMCPairsSearch.Show vbModal
       Set UMCPairsSearch = Nothing
    Else
       MsgBox "Paired UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes then use menu item 'Steps->4. Find Pairs' to find paired UMCs.", vbOKOnly, glFGTU
    End If
End Sub

Private Sub ShowSearchMTDBUMCSingleMassForm()
    '---------------------------------------------------
    'displays search of MT tag database form for UMCs
    '---------------------------------------------------
    Dim UMC_IDSearch As New frmSearchMT_ConglomerateUMC
    On Error Resume Next
    If GelUMC(nMyIndex).UMCCnt > 0 Then
       UMC_IDSearch.CallerID = nMyIndex
       UMC_IDSearch.Show vbModal
       Set UMC_IDSearch = Nothing
    Else
       MsgBox "You must cluster the data into Unique Mass Classes before searching for matching MT tags.  Please use menu item 'Steps->2. Find UMCs' to accomplish this.", vbOKOnly, glFGTU
    End If
End Sub

Private Sub ShowUMC2003Form()
    frmUMCSimple.Tag = nMyIndex
    frmUMCSimple.Show vbModal
    If GelUMCDraw(nMyIndex).Visible Then
        bNeedToUpdate = True
        csMyCooSys.CoordinateDraw
        picGraph.Refresh
    End If
    UpdateTICPlotAndFeatureBrowsersIfNeeded True
End Sub

Private Sub ShowUMCIonNetworksForm()
    frmUMCIonNet.Tag = nMyIndex
    frmUMCIonNet.Show vbModal
    If GelUMCDraw(nMyIndex).Visible Then
        bNeedToUpdate = True
        csMyCooSys.CoordinateDraw
        picGraph.Refresh
    End If
    UpdateTICPlotAndFeatureBrowsersIfNeeded True
End Sub

Private Sub ShowUMCNetAdjustmentForm()
    Dim NETAdj As New frmSearchForNETAdjustmentUMC
    On Error Resume Next
    If GelUMC(nMyIndex).UMCCnt > 0 Then
       NETAdj.CallerID = nMyIndex
       NETAdj.Show vbModal
       Set NETAdj = Nothing
    Else
       MsgBox "You must cluster the data into Unique Mass Classes before NET Adjustment.  Please use menu item 'Steps->2. Find UMCs' to accomplish this.", vbOKOnly, glFGTU
    End If
End Sub

Private Sub Form_Activate()
    ActivateGraph
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then ZoomActionCancel
End Sub

Private Sub Form_Load()
bLoad = True
bResize = False
fgSelBoxVisible = False

Me.width = 8375
Me.Height = 6100 - GetSystemMetrics(SM_CYMENU) * Screen.TwipsPerPixelY

ShowHidePNNLMenus

If GetChildCount = 1 Then 'Mesa only open gel (Jar Jar)
   frmTracker.Show        'toolbar and tracker visible
Else                      'already some open gels - adjust or die!
   mnuViewTracker.Checked = IsWinLoaded(TrackerCaption)
   mnuViewToolbar.Checked = MDIForm1.picToolbar.Visible
End If
Set GelSel = New Sel
' MonroeMod
tmrRefreshTimer.Interval = REFRESH_TIMER_INTERVAL
mnuVMenuModeIncludeObsolete.Checked = glbPreferencesExpanded.MenuModeIncludeObsolete
UpdateMenuMode glbPreferencesExpanded.MenuModeDefault

mnuSCopyPointsInViewIncludeSearchResults.Checked = glbPreferencesExpanded.CopyPointsInViewIncludeSearchResultsChecked
mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked = glbPreferencesExpanded.CopyPointsInViewByUMCIncludeSearchResultsChecked

' Make sure Scan Number mode is enabled
mnuViewFN_Click
 
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim sMsg As String
Dim sFileName As String
Dim nResponse As Integer

' MonroeMod
Dim blnSuccess As Boolean

On Error GoTo err_QueryUnload

If Me.Tag <= 0 Then Exit Sub       'for forms created from the unique mw count don't ask user to save
If GelStatus(nMyIndex).Dirty Then  'something changed
   sFileName = Me.Caption
   sMsg = "File [" & sFileName & "] has changed."
   sMsg = sMsg & vbCrLf & "Do you want to save the changes?"
   nResponse = MsgBox(sMsg, vbYesNoCancel + vbExclamation, MDIForm1.Caption)
   Select Case nResponse
   Case 2   'User chose Cancel; cancel unload
        Cancel = True
   Case 6   'User chose Yes; save and unload
        If Left(Me.Caption, 8) = "Untitled" Then 'never saved before
           sFileName = ""
        Else  'form caption has file name
           sFileName = Me.Caption
        End If
        ' MonroeMod
        blnSuccess = SaveFileAsInit(mFileSaveMode)
        If Not blnSuccess Then Cancel = True
   Case 7   'User choose No; unload, don't save
        Cancel = False
   End Select
End If
Exit Sub

err_QueryUnload:
If Len(Me.Tag) <= 0 Then Exit Sub   'must be something that was never activated
Cancel = True
End Sub

Private Sub Form_Resize()
If (bResize And (Me.ScaleHeight > 0) And (Me.ScaleWidth > 0)) Then
   If Me.ScaleHeight <> OldMeScaleH Then bPaint = False
   picGraph.width = Me.ScaleWidth
   bPaint = True
   picGraph.Height = Me.ScaleHeight
End If
OldMeScaleH = Me.ScaleHeight
OldMeScaleW = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetGelStateToDeleted nMyIndex
Set csMyCooSys = Nothing
Set GelSel = Nothing
If GetChildCount < 1 Then
   GetRecentFiles
   If IsWinLoaded(TrackerCaption) Then Unload frmTracker
End If
DestroyStructures nMyIndex
MDIForm1.ProperToolbar False
End Sub

Private Sub GelSel_ChangeCSSel()
SelToolBox1.SelCSCount = GelSel.CSSelCnt
If fgSelBoxVisible Then SelToolBox1.ClearResults
End Sub

Private Sub GelSel_ChangeIsoSel()
SelToolBox1.SelIsoCount = GelSel.IsoSelCnt
If fgSelBoxVisible Then SelToolBox1.ClearResults
End Sub

Private Sub mnu2lsAddVisiblePointsToSelected_Click()
    AddVisiblePointsToSelection
End Sub

Private Sub mnu2lsClearSelection_Click()
    ClearSelectedPoints True
End Sub

Private Sub mnu2lsCloseAllICR2LSMassSpectra_Click()
    ICR2LSCloseAllScopes True
End Sub

Private Sub mnu2lsExcludeAllButSelection_Click()
'exclude all points except those currently selected and clear selection
ExcludeAllButSelection nMyIndex
picGraph.Refresh
End Sub

Private Sub mnu2lsExcludeSelection_Click()
'exclude all points currently selected and clear selection
ExcludeSelection nMyIndex
picGraph.Refresh
End Sub

Private Sub mnu2lsExcludeVisiblePoints_Click()
    ExcludeVisiblePointsOneStep
End Sub

Private Sub mnu2lsFilter_Click()
ShowFilterForm
End Sub

Private Sub mnu2lsFilterGraph_Click()
frmFilterGraph.Tag = nMyIndex
frmFilterGraph.Show vbModal
End Sub

' Unused function (February 2005)
''Private Sub mnu2lsFilterIsoCom_Click()
''' Legacy: Filter on Isotopic Composition
''On Error Resume Next
''frmFilterIsoCom.Tag = nMyIndex
''frmFilterIsoCom.Show vbModal
''End Sub

Private Sub mnu2lsFilterPointsByMass_Click()
    frmExcludeMassRange.SetCallerID nMyIndex
    frmExcludeMassRange.Show vbModal
End Sub

Private Sub mnu2lsLockSelection_Click()
mnu2lsLockSelection.Checked = Not mnu2lsLockSelection.Checked
fgSelProtected = mnu2lsLockSelection.Checked
End Sub

Private Sub mnu2LSMergePEKFiles_Click()
frmPEKMerge.Show vbModal
End Sub

Private Sub mnu2lsShowSpectrumAllSelectedPts_Click()
    ShowMSSpectrumForAllSelected False
End Sub

Private Sub mnu2lsShowSpectrumNearestCursorPoint_Click()
    ' Menu is bound to key Ctrl+D
    ShowMSSpectrumForPointNearestCursor
End Sub

Private Sub mnu2lsShowSpectrumSumSelectedPts_Click()
    ShowMSSpectrumForAllSelected True
End Sub

Private Sub mnu2LSSplitPEKFiles_Click()
frmPEKSplit.Show vbModal
End Sub

Private Sub mnu2lsToleranceRefinement_Click()
    ShowErrorDistribution2DForm
End Sub

Private Sub mnu2lsSearchUMCSingleMass_Click()
    ShowSearchMTDBUMCSingleMassForm
End Sub

Private Sub mnu2lsUMC2003_Click()
    ShowUMC2003Form
End Sub

Private Sub mnu2lsUMCIonNet_Click()
    ShowUMCIonNetworksForm
End Sub

'' Obsolete form
''Private Sub mnu2lsUMCWithAutoRefine_Click()
''On Error Resume Next
''frmUMCWithAutoRefine.Tag = nMyIndex
''frmUMCWithAutoRefine.Show vbModal
''If GelUMCDraw(nMyIndex).Visible Then
''    bNeedToUpdate = True
''    csMyCooSys.CoordinateDraw
''    picGraph.Refresh
''End If
''End Sub
''
Private Sub mnuEditCopyEMF_Click()
GelDrawMetafile nMyIndex, False, "", True, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeFilenameAndDate, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeTextLabels
End Sub

Private Sub mnuEditCopyEMFIncludeTextLabels_Click()
    SetEditCopyEMFOptions glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeFilenameAndDate, Not glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeTextLabels
End Sub

Private Sub mnuEditCopyEMFIncludeFileNameAndTime_Click()
    SetEditCopyEMFOptions Not glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeFilenameAndDate, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeTextLabels
End Sub

Private Sub mnuEditNETAdjustment_Click(Index As Integer)
    Dim NETAdj As frmSearchForNETAdjustment
    
    Select Case Index
    Case enaEditNETAdjustmentMenuConstants.enaIndividualPoints
        On Error Resume Next
        Set NETAdj = New frmSearchForNETAdjustment
        NETAdj.CallerID = nMyIndex
        NETAdj.Show vbModal
        Set NETAdj = Nothing
    Case enaEditNETAdjustmentMenuConstants.enaUMCIterative
        ShowUMCNetAdjustmentForm
    Case enaEditNETAdjustmentMenuConstants.enaMSAlign
        ShowMSAlignForm
    Case Else
        ' Unknown index
        Debug.Assert False
    End Select
End Sub

Private Sub mnuFile_Click()
    Select Case mFileSaveMode
    Case fsIncludeExtended
        mnuSave.Caption = "&Save (With Extended Info)"
    Case fsNoExtended
        mnuSave.Caption = "&Save (No Extended Info)"
    Case fsUnknown
        mnuSave.Caption = "&Save"
    End Select
    
    If mFileSaveMode = fsNoExtended Then
                  mnuSaveAs.Caption = "Save As With Extended Info"
        mnuSaveAsCompressed.Caption = "Save &As Without Extended Info"
    Else
                  mnuSaveAs.Caption = "Save &As With Extended Info"
        mnuSaveAsCompressed.Caption = "Save As Without Extended Info"
    End If
    
End Sub

Private Sub mnuHelpSetTraceLogLevel_Click()
    SetTraceLogLevel 0, True
End Sub

Private Sub mnuNewAnalysis_Click()
    ChooseNewAnalysisUsingDMS
End Sub

Private Sub mnuNewAutoAnalysis_Click()
Dim udtAutoParams As udtAutoAnalysisParametersType

InitializeAutoAnalysisParameters udtAutoParams
udtAutoParams.ShowMessages = True

AutoAnalysisStart udtAutoParams
End Sub

Private Sub mnuResetAllOptionsToDefaults_Click()
    Dim eResponse As VbMsgBoxResult
    
    eResponse = MsgBox("Are you sure you want to reset all options to their default values?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Reset Options")
    If eResponse = vbYes Then
        ResetGelPrefs GelData(nMyIndex).Preferences
        
        ResetExpandedPreferences glbPreferencesExpanded, "", True
        
        SetDefaultUMCDef GelSearchDef(nMyIndex).UMCDef
        SetDefaultUMCIonNetDef GelSearchDef(nMyIndex).UMCIonNetDef
        SetDefaultUMCNETAdjDef GelUMCNETAdjDef(nMyIndex)
        
        SetDefaultSearchAMTDef samtDef, GelUMCNETAdjDef(nMyIndex)
        
        GelSearchDef(nMyIndex).AMTSearchOnUMCs = samtDef
        GelSearchDef(nMyIndex).AMTSearchOnIons = samtDef
        GelSearchDef(nMyIndex).AMTSearchOnPairs = samtDef
        
        GelP_D_L(nMyIndex).SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
        
    End If
End Sub

Private Sub mnuSaveAsLegacyData_Click()
    SaveFileAsInit fsLegacy
End Sub

Private Sub mnuSaveLoadEditAnalysisSettings_Click()
    frmEditAnalysisSettings.SetCallerID nMyIndex
    frmEditAnalysisSettings.Show vbModeless, MDIForm1
End Sub

Private Sub mnuSavePEKFileUsingAllDataPoints_Click()
    SavePEKFileUsingOriginalPEKFileAsTemplate
End Sub

Private Sub mnuSavePEKFileUsingDataPointsInView_Click()
    SavePEKFileUsingDataPoints True
End Sub

Private Sub mnuSavePEKFileUsingUMCs_Click()
    SavePEKFileUsingUMCs True
End Sub

Private Sub mnuSaveSettingsToIniFile_Click()
    SaveCurrentSettingsToIniFile nMyIndex
End Sub

Private Sub mnuSCopyPointsInViewByUMCIncludeSearchResults_Click()
    mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked = Not mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked
    glbPreferencesExpanded.CopyPointsInViewByUMCIncludeSearchResultsChecked = mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked
End Sub

Private Sub mnuSCopyPointsInViewByUMCtoFile_Click()
    CopyAllUMCsInView -1, True
End Sub

Private Sub mnuSCopyPointsInViewIncludeSearchResults_Click()
    mnuSCopyPointsInViewIncludeSearchResults.Checked = Not mnuSCopyPointsInViewIncludeSearchResults.Checked
    glbPreferencesExpanded.CopyPointsInViewIncludeSearchResultsChecked = mnuSCopyPointsInViewIncludeSearchResults.Checked
End Sub

Private Sub mnuSCopyPointsInViewOneHitPerLine_Click()
    mnuSCopyPointsInViewOneHitPerLine.Checked = Not mnuSCopyPointsInViewOneHitPerLine.Checked
    SynchronizeCopyDataOptions
End Sub

Private Sub mnuSCopyPointsInViewToFile_Click()
    CopyAllPointsInView -1, True
End Sub

Private Sub mnuShowLowSegmentCountUMCs_Click()
lAction = glNoAction
If GelUMC(nMyIndex).UMCCnt > 0 Then
   ShowNetAdjUMCPoints nMyIndex, UMC_INDICATOR_BIT_LOWSEGMENTCOUNT_ADDITION
   picGraph.Refresh
Else
   MsgBox "UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuShowNetAdjUMCs_Click()
lAction = glNoAction
If GelUMC(nMyIndex).UMCCnt > 0 Then
   ShowNetAdjUMCPoints nMyIndex, UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
   picGraph.Refresh
Else
   MsgBox "UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuShowNETAdjUMCsWithDBHit_Click()
lAction = glNoAction
If GelUMC(nMyIndex).UMCCnt > 0 Then
   ShowNetAdjUMCPoints nMyIndex, UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
   picGraph.Refresh
Else
   MsgBox "UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes.", vbOKOnly, glFGTU
End If

End Sub

Private Sub mnuShowSplitUMCs_Click()
lAction = glNoAction
If GelUMC(nMyIndex).UMCCnt > 0 Then
   ShowSplitUMCPoints nMyIndex
   picGraph.Refresh
Else
   MsgBox "UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes.", vbOKOnly, glFGTU
End If
End Sub

''Private Sub mnuSpecialTestAutoNetWarp_Click()
''    TestAutoNETWarp nMyIndex
''End Sub

Private Sub mnuSplitUMCs_Click()
    SplitUMCsByAbundance nMyIndex, Me, True, False
End Sub

Private Sub mnuStepsAnalysisHistoryLog_Click()
    ViewAnalysisHistory nMyIndex
End Sub

Private Sub mnuStepsDBSearchConglomerateMass_Click()
    ShowSearchMTDBUMCSingleMassForm
End Sub

Private Sub mnuStepsFileNew_Click()
    LoadNewDataFile
End Sub

Private Sub mnuStepsFileNewChooseFromDMS_Click()
    ChooseNewAnalysisUsingDMS
End Sub

Private Sub mnuStepsFilter_Click()
    ShowFilterForm
End Sub

Private Sub mnuStepsFindPairsMode_Click(Index As Integer)
    Select Case Index
    Case 0
        ShowPairsDeltaUMCForm
    Case 1
        ShowPairsLabelUMCForm
    Case 2
        ShowPairsDeltaLabelUMCForm
    Case Else
        ' Unknkown mode
        Debug.Assert False
        ShowPairsDeltaUMCForm
    End Select
End Sub

Private Sub mnuStepsNETAdjustmentMode_Click(Index As Integer)
    If Index = namStepsNETAdjustmentModeConstants.namUMCIterative Then
        ShowUMCNetAdjustmentForm
    Else
        ShowMSAlignForm
    End If
End Sub

Private Sub mnuStepsPairsDBSearchUsingPairs_Click(Index As Integer)
    Select Case Index
    Case 0
        ShowSearchMTDBUMCPairsN14N15
    Case 1
        ShowSearchMTDBUMCPairsICAT
    Case 2
        ShowSearchMTDBUMCPairsPEON14N15
    Case Else
        ' Unknkown mode
        Debug.Assert False
        ShowSearchMTDBUMCPairsN14N15
    End Select
End Sub

Private Sub mnuStepsSaveQCPlots_Click()
    AutoGenerateQCPlots Me, nMyIndex
End Sub

Private Sub mnuStepsSelectMassTags_Click()
    ShowOrganizeDatabaseConnectionsForm
End Sub

Private Sub mnuStepsToleranceRefinement_Click()
    ShowErrorDistribution2DForm
End Sub

Private Sub mnuStepsUMCMode_Click(Index As Integer)
    If Index = 0 Then
        ShowUMC2003Form
    Else
        ShowUMCIonNetworksForm
    End If
End Sub

Private Sub mnuViewNET_Click()
SetXAxisLabelType True
End Sub

Private Sub mnuViewPairsBrowser_Click()
    frmPairBrowser.CallerIDNew = nMyIndex
    frmPairBrowser.Show vbModeless
End Sub

Private Sub mnuViewTICandBPIPlots_Click()
    frmTICAndBPIPlots.CallerID = nMyIndex
    frmTICAndBPIPlots.Show vbModeless
End Sub

Private Sub mnuViewUMCBrowser_Click()
    frmUMCBrowser.CallerIDNew = nMyIndex
    frmUMCBrowser.Show vbModeless
End Sub

Private Sub mnuViewZoomRegionListEditor_Click()
    frmZoomRegionList.CallerID = nMyIndex
    frmZoomRegionList.Show
End Sub

Private Sub mnuVMenuModeIncludeObsolete_Click()
    mnuVMenuModeIncludeObsolete.Checked = Not mnuVMenuModeIncludeObsolete.Checked
    UpdateVisibleMenus
End Sub

Private Sub mnuVMenuModeSelect_Click(Index As Integer)
    UpdateMenuMode Index
End Sub

Private Sub mnuEditMTStatus_Click()
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly
Me.MousePointer = vbDefault
End Sub

Private Sub mnu2lsOptions_Click()
Dim blnAutoAdjSizeSaved As Boolean
Dim i As Integer

frmOptions.Tag = nMyIndex
If IsWinLoaded(TrackerCaption) Then frmTracker.Hide
bPaint = False

' Save the current zoom settings
Dim lngXMin As Long, lngXMax As Long
Dim dblYMin As Double, dblYMax As Double
lngXMin = csMyCooSys.CurrRXMin
lngXMax = csMyCooSys.CurrRXMax
dblYMin = csMyCooSys.CurrRYMin
dblYMax = csMyCooSys.CurrRYMax

' Save the current state of glbPreferencesExpanded.AutoAdjSize
blnAutoAdjSizeSaved = glbPreferencesExpanded.AutoAdjSize

' Show the options form
frmOptions.Show vbModal

' Need to bump up vWhatever if glbPreferencesExpanded.AutoAdjSize changed
If glbPreferencesExpanded.AutoAdjSize <> blnAutoAdjSizeSaved And Abs(vWhatever) = 1 Then vWhatever = 2

Select Case Abs(vWhatever)
Case 1  'redraw only me
    If vWhatever < 0 Then InitDrawERColors nMyIndex
    BuildCooSys
Case 2  'redraw all because colors changed
    SetBackForeColorObjects
    SetCSIsoColorObjects
    SetDDRColorObjects
    SetSelColorObjects
    GelColorsChange False, vWhatever
Case 3  'redraw all and set options for others same as mine
    SetBackForeColorObjects
    SetCSIsoColorObjects
    SetDDRColorObjects
    SetSelColorObjects
    GelColorsChange True, vWhatever
End Select

' Make sure the view is zoomed correctly
csMyCooSys.ZoomInR lngXMin, dblYMin, lngXMax, dblYMax

If Abs(vWhatever) >= 2 Then
    ' Need to call Coordinate draw for the other Gels (it gets called for this gel in the above call to .ZoomInR
    For i = 1 To UBound(GelBody)
        If i <> nMyIndex And Not GelStatus(i).Deleted Then
            GelBody(i).csMyCooSys.CoordinateDraw
        End If
    Next i
End If

bNeedToUpdate = True

If IsWinLoaded(TrackerCaption) Then frmTracker.Show
picGraph.SetFocus
End Sub


Private Sub mnu2lsORFCenteredSearchFASTA_Click()
frmProtoViewFASTA.Tag = nMyIndex
frmProtoViewFASTA.Show vbModal
End Sub

Private Sub mnu2lsORFCenteredSearchPRISM_Click()
Dim Resp As Long
If GelAnalysis(nMyIndex) Is Nothing Then
   MsgBox "Current display is not associated with any MT tag database.  Unable to display the protein viewer.", vbOKOnly, glFGTU
   Exit Sub
Else
   'need to check only connection string to the database; if not the same
   'need to load MT tags from the database before proceeding
   If AMTCnt <= 0 Then
      Resp = MsgBox("MT tag data not loaded. Load now?", vbYesNoCancel, glFGTU)
      Select Case Resp
      Case vbYes
         Me.MousePointer = vbHourglass
' MonroeMod
         LoadMassTags nMyIndex, Me          'it does not have to be anything loaded for us to use frmProtoView
         Me.MousePointer = vbDefault
      Case vbNo
      Case vbCancel
         Exit Sub
      End Select
   Else
      If CurrMTDatabase <> GelAnalysis(nMyIndex).MTDB.cn.ConnectionString Then
         Resp = MsgBox("Loaded database does not match required data. Load correct data?", vbYesNoCancel, glFGTU)
         Select Case Resp
            Case vbYes
            Me.MousePointer = vbHourglass
' MonroeMod
            LoadMassTags nMyIndex, Me          'it does not have to be anything loaded for us to use frmProtoView
            Me.MousePointer = vbDefault
         Case Else
            Exit Sub
         End Select
      End If
   End If
End If
frmProtoViewPRISM.Tag = nMyIndex
frmProtoViewPRISM.Show vbModal
End Sub

Private Sub mnu2lsOverlaysManager_Click()
frmOverlayManager.Show
End Sub

Private Sub mnu2lsResetGraph_Click()
ResetGraph True, True, fgDisplay
End Sub

Private Sub mnu2lsSearchAMTNew_Click()
'---------------------------------------------
'displays search of MT tag database form
'---------------------------------------------
Dim OrgDBSearch As New frmSearchMT
On Error Resume Next
OrgDBSearch.CallerID = nMyIndex
OrgDBSearch.Show vbModal
Set OrgDBSearch = Nothing
End Sub

Private Sub mnu2lsSearchAMTOld_Click()
'-------------------------------------
'search legacy databases
'-------------------------------------
frmSearchAMT.Tag = nMyIndex
frmSearchAMT.Show vbModal
End Sub

Private Sub mnu2LSSearchForORFs_Click()
On Error Resume Next
frmORFSearch.Tag = nMyIndex
frmORFSearch.Show vbModal
End Sub

Private Sub mnu2lsSearchUMCMassTags_Click()
'--------------------------------------------------
'search MT tag database on the blueprint of the
'unique mass classes; contains also export function
'--------------------------------------------------
Dim MyUMCSearch As New frmUMCIdentification
On Error Resume Next
If GelUMC(nMyIndex).UMCCnt > 0 Then
   MyUMCSearch.CallerID = nMyIndex
   MyUMCSearch.Show vbModal
   Set MyUMCSearch = Nothing
Else
   MsgBox "UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnu2lsSearchUMCPairs_Click()
    ShowSearchMTDBUMCPairsN14N15
End Sub

Private Sub mnu2lsSearchUMCPairs_ICAT_Click()
    ShowSearchMTDBUMCPairsICAT
End Sub

Private Sub mnu2lsSearchUMCPairs_PEON14N15_Click()
    ShowSearchMTDBUMCPairsPEON14N15
End Sub

Private Sub mnu2lsSelAvgER_Click()
Dim aInt() As Double
Dim SelCnt As Long
Dim aAvg As Double
SelCnt = GetSelectionFieldNumeric(nMyIndex, glFIELD_ER, aInt())
If SelCnt > 0 Then
   Avg_D aInt(), SelCnt, aAvg
   If SelCnt > 0 Then
      MsgBox "Avg.ER (" & SelCnt & " points): " & aAvg, vbOKOnly
   ElseIf SelCnt < 0 Then
      MsgBox "Error calculating average.", vbOKOnly
   Else
      MsgBox "No valid selection ER data found.", vbOKOnly
   End If
Else
   MsgBox "No selection ER data found.", vbOKOnly
End If
End Sub

Private Sub mnu2lsSelAvgFit_Click()
Dim aInt() As Double
Dim SelCnt As Long
Dim aAvg As Double
SelCnt = GetSelectionFieldNumeric(nMyIndex, glFIELD_FIT, aInt())
If SelCnt > 0 Then
   Avg_D aInt(), SelCnt, aAvg
   If SelCnt > 0 Then
      MsgBox "Avg.Fit (" & SelCnt & " points): " & aAvg, vbOKOnly
   ElseIf SelCnt < 0 Then
      MsgBox "Error calculating average.", vbOKOnly
   Else
      MsgBox "No valid selection Fit data found.", vbOKOnly
   End If
Else
   MsgBox "No selection Fit data found.", vbOKOnly
End If
End Sub

Private Sub mnu2lsSelAvgInt_Click()
Dim aInt() As Double
Dim SelCnt As Long
Dim aAvg As Double
SelCnt = GetSelectionFieldNumeric(nMyIndex, glFIELD_ABU, aInt())
If SelCnt > 0 Then
   Avg_D aInt(), SelCnt, aAvg
   If SelCnt > 0 Then
      MsgBox "Avg.Intensity (" & SelCnt & " points): " & Format$(aAvg, "Scientific"), vbOKOnly
   ElseIf SelCnt < 0 Then
      MsgBox "Error calculating average.", vbOKOnly
   Else
      MsgBox "No valid selection Intensity data found.", vbOKOnly
   End If
Else
   MsgBox "No selection Intensity data found.", vbOKOnly
End If
End Sub

Private Sub mnu2lsSelAvgMW_Click()
Dim aInt() As Double
Dim SelCnt As Long
Dim aAvg As Double
SelCnt = GetSelectionFieldNumeric(nMyIndex, glFIELD_MW, aInt())
If SelCnt > 0 Then
   Avg_D aInt(), SelCnt, aAvg
   If SelCnt > 0 Then
      MsgBox "Avg.MW (" & SelCnt & " points): " & aAvg, vbOKOnly
   ElseIf SelCnt < 0 Then
      MsgBox "Error calculating average.", vbOKOnly
   Else
      MsgBox "No valid selection MW data found.", vbOKOnly
   End If
Else
   MsgBox "No selection MW data found.", vbOKOnly
End If
End Sub

Private Sub mnu2lsShowAll_Click()
Dim i As Integer
lAction = glNoAction
'set all points to visible (positive ids) and clear selection
With GelData(nMyIndex)
     For i = 1 To MAX_FILTER_COUNT
        .DataFilter(i, 0) = False
     Next i
     .DataFilter(fltID, 1) = 0      'identity
     GelCSIncludeAll (nMyIndex)
     GelIsoIncludeAll (nMyIndex)
End With
picGraph.Refresh
End Sub

Private Sub mnu2lsShowAllData_Click()
'remains to be defined
MsgBox "This command is disabled (under revision).", vbOKOnly
End Sub

' No longer supported (March 2006)
''Private Sub mnu2lsShowAMTRecord_Click()
''Dim sID As String
''If dbAMT Is Nothing Then
''   If Not ConnectToLegacyAMTDB(nMyIndex, True) Then Exit Sub 'error message will be displayed
''End If
''Select Case LastType
''Case glCSType
''     sID = GelData(nMyIndex).CSData(LastID).MTID
''Case glIsoType
''     sID = GelData(nMyIndex).IsoData(LastID).MTID
''End Select
''sID = GetIDFromString(sID, AMTMark, AMTIDEnd)
''If Len(sID) > 0 Then
''   MsgBox GetAMTRecordByID(sID), vbOKOnly, "MT Database Record"
''Else
''   MsgBox "No MT reference found for current selection.", vbOKOnly, "MT Database Record"
''End If
''End Sub

Private Sub mnu2lsShowCurrentData_Click()
'remains to be defined
MsgBox "This command is disabled (under revision).", vbOKOnly
End Sub

Private Sub mnu2lsShowSelected_Click()
'remains to be defined
MsgBox "This command is disabled (under revision).", vbOKOnly
End Sub

Private Sub mnu2lsShowSpectrumForLastSelectedPt_Click()
    ShowMSSpectrumForLastID
End Sub

Private Sub mnu2lsShowUMCOnly_Click()
lAction = glNoAction
If GelUMC(nMyIndex).UMCCnt > 0 Then
   Call fUMCSpotsOnly(nMyIndex)
   picGraph.Refresh
Else
   MsgBox "UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnu2LsZoomIn_Click()
frmZoomIn.Tag = nMyIndex
frmZoomIn.Show vbModal
End Sub

Private Sub mnu2lsZoomOut_Click()
csMyCooSys.ZoomOut
End Sub

Private Sub mnu2lsZoomOutOneLevel_Click()
csMyCooSys.ZoomPrevious
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, MDIForm1
End Sub

Private Sub mnuArangeIcons_Click()
MDIForm1.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
MDIForm1.Arrange vbCascade
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuEditAbundance_Click()
frmAbuEditor.Tag = nMyIndex
frmAbuEditor.Show vbModal
End Sub

Private Sub mnuEditAdjScans_Click()
frmAdjScan.Tag = nMyIndex
frmAdjScan.Show vbModal
End Sub

Private Sub mnuEditAvgUMC_Click()
Dim Resp As Long
On Error Resume Next
If GelUMC(nMyIndex).UMCCnt > 0 Then
   Resp = MsgBox("Application of this function will change masses from original file. Continue with averaging masses of UMCs?", vbYesNo, glFGTU)
   If Resp <> vbYes Then Exit Sub
   Me.MousePointer = vbHourglass
   If UMCAverageMass(nMyIndex) Then
      GelStatus(nMyIndex).Dirty = True
      ResetGraph True, True, fgDisplay
   Else
      MsgBox "Error calculating average UMC masses.", vbOKOnly, glFGTU
   End If
   Me.MousePointer = vbDefault
Else
   MsgBox "Unique mass classes not found. Use function Unique Mass Classes from Tools menu first.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuEditCalibration_Click()
frmCalibration.Tag = nMyIndex
frmCalibration.Show vbModal
End Sub

Private Sub mnuEditClearER_Click()
'--------------------------------------------------
'deletes all expression ratios and resets the graph
'--------------------------------------------------
ResetERValues nMyIndex
With GelData(nMyIndex)
    .DataFilter(fltAR, 0) = 0
    .DataFilter(fltAR, 1) = -1
    .DataFilter(fltAR, 2) = -1
End With
ResetGraph
End Sub

Private Sub mnuEditClearIDAll_Click()
'---------------------------------------------
'remove all data from ID fields
'---------------------------------------------
Me.MousePointer = vbHourglass
CleanIDData nMyIndex
Me.MousePointer = vbDefault
End Sub

Private Sub mnuEditClearIDBadDelta_Click()
Me.MousePointer = vbHourglass
RemoveBadMTs_Delta nMyIndex
Me.MousePointer = vbDefault
End Sub

Private Sub mnuEditClearIDBadDeltaMTICAT_Click()
Me.MousePointer = vbHourglass
If AMTCnt = 0 Then
    ConfirmMassTagsAndInternalStdsLoaded Me, nMyIndex, True
End If
RemoveBadMTs_DeltaMT nMyIndex, PAIR_ICAT
Me.MousePointer = vbDefault
End Sub

Private Sub mnuEditClearIDBadDeltaMTN14N15_Click()
Me.MousePointer = vbHourglass
If AMTCnt = 0 Then
    ConfirmMassTagsAndInternalStdsLoaded Me, nMyIndex, True
End If
RemoveBadMTs_DeltaMT nMyIndex, PAIR_N14N15
Me.MousePointer = vbDefault
End Sub

Private Sub mnuEditClearIDNoId_Click()
Me.MousePointer = vbHourglass
RemoveIDWithoutID nMyIndex
Me.MousePointer = vbDefault
End Sub

Private Sub mnuEditComment_Click()
frmComment.Caption = "Comment - " & Me.Caption
frmComment.Tag = nMyIndex
frmComment.Show vbModal
End Sub

Private Sub mnuEditCopyBMP_Click()
On Error GoTo err_CopyGraph
picGraph.AutoRedraw = True
GelDrawScreen nMyIndex, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeFilenameAndDate, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeTextLabels
Clipboard.Clear
Clipboard.SetData picGraph.Image, vbCFBitmap
picGraph.AutoRedraw = False

With glbPreferencesExpanded.GraphicExportOptions
    If .CopyEMFIncludeFilenameAndDate Or Not .CopyEMFIncludeTextLabels Then
        picGraph.Refresh
    End If
End With

Exit Sub

err_CopyGraph:
MsgBox "Unexpected error." & vbCrLf & sErrLogReference, vbOKOnly
LogErrors Err.Number, "Copy Graph"
End Sub

Private Sub mnuEditCopyWMF_Click()
GelDrawMetafile nMyIndex, False, "", False, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeFilenameAndDate, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeTextLabels
End Sub

Private Sub mnuEditDatabaseConnection_Click()
    ShowOrganizeDatabaseConnectionsForm
End Sub

Private Sub mnuEditDiscreteMWs_Click()
frmDiscreteDisplay.Tag = nMyIndex
frmDiscreteDisplay.Show vbModal
End Sub

Private Sub mnuEditFind_Click()
frmFind.Tag = nMyIndex
frmFind.Show vbModal
End Sub

Private Sub mnuEditFSDenseAreas_Click()
frmFindDenseAreas.Tag = nMyIndex
frmFindDenseAreas.Show vbModal
End Sub

Private Sub mnuEditNETFormula_Click()
    frmNETFormulaEditor.Tag = nMyIndex
    frmNETFormulaEditor.Show vbModal
End Sub

Private Sub mnuEditParameters_Click()
frmParameters.Tag = nMyIndex
frmParameters.Show vbModeless, MDIForm1
End Sub

Private Sub mnuEditResidualDisplay_Click()
Call FillDisplay0ResidualCounts(nMyIndex)
End Sub

Private Sub mnuEditUMC_Click()
If GelUMC(nMyIndex).UMCCnt > 0 Then
   frmVisUMC.Tag = nMyIndex
   frmVisUMC.Show vbModal
    UpdateTICPlotAndFeatureBrowsersIfNeeded True
Else
   MsgBox "UMCs not found. Please use menu item 'Steps->2. Find UMCs' to cluster the data into unique mass classes.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuExit_Click()
Unload MDIForm1
End Sub

Private Sub mnuNew_Click()
LoadNewDataFile
End Sub

Private Sub mnuOpen_Click()
MDIStatus True, "Loading ... please be patient"
FileOpenProc (Me.hwnd)
End Sub

''' MonroeMod: New subroutine
''Private Sub mnuORFViewer_Click()
''    ORFViewerLoader.ShowORFViewerForm Me.Caption, CInt(nMyIndex), False
''    If mFileSaveMode = fsNoExtended Then mFileSaveMode = fsUnknown
''End Sub

Private Sub mnuPrint_Click()
If GoPrint(Me.Tag) < 0 Then MsgBox "Error printing.", vbOKOnly
End Sub

Private Sub mnuPrintSetup_Click()
PrinterSetupAPIDlg (Me.hwnd)
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
    ' MonroeMod
    Dim strFullFilePath As String
    
    strFullFilePath = RecentFileLookUpFullPath(mnuRecentFiles(Index).Caption)
    If Len(strFullFilePath) > 0 Then
        MDIStatus True, "Loading ... please be patient"
        ReadGelFile strFullFilePath
    End If
End Sub


Private Sub mnuSAttentionList_Click()
If GelP_D_L(nMyIndex).PCnt > 0 Then
   frmAttentionList.Tag = nMyIndex
   frmAttentionList.Show vbModal
Else
   MsgBox "No pairs detected. To find pairs use '" & _
    "Delta Label'" & "' functions from the Special menu.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuSave_Click()
Dim sFileName As String
If (Left(Me.Caption, 8) = "Untitled") Or (Len(Me.Caption) <= 0) Then
   mFileSaveMode = fsUnknown
    SaveFileAsWrapper "", True, True, False, True
Else
   sFileName = Me.Caption
    ' MonroeMod
    Select Case mFileSaveMode
    Case fsNoExtended
        SaveFileAsWrapper sFileName, False, False, True
    Case fsIncludeExtended
        SaveFileAsWrapper sFileName, False, False, False
    Case Else
        SaveFileAsWrapper sFileName, False, True, False, glbPreferencesExpanded.ExtendedFileSaveModePreferred
    End Select
End If
End Sub

Private Sub mnuSaveAs_Click()
    SaveFileAsInit fsIncludeExtended
End Sub

' MonroeMod: New subroutine (Also need a New menu item named mnuSaveAsCompressed)
Private Sub mnuSaveAsCompressed_Click()
    SaveFileAsInit fsNoExtended
End Sub

Private Sub mnuSavePic_Click()
'-----------------------------
Dim sSaveFileName As String
Dim PicSaveType As pftPictureFileTypeConstants

On Error Resume Next

sSaveFileName = FileSaveProc(Me.hwnd, StripFullPath(Me.Caption), fstFileSaveTypeConstants.fstPIC, PicSaveType)

SaveFileAsPicture nMyIndex, sSaveFileName, PicSaveType

End Sub

Private Sub mnuSaveWYS_Click()
MDIStatus True, "Saving file ..."
SaveWYSAs nMyIndex, mFileSaveMode
MDIStatus False, ""
End Sub

Private Sub mnuSCopyPointsInViewToClipboard_Click()
    CopyAllPointsInView
End Sub

Private Sub mnuSCopyPointsInViewByUMCtoClipboard_Click()
    CopyAllUMCsInView
End Sub

Private Sub mnuSCopyPointsInViewToClipboardAuto_Click()
    Dim strResponse As String
    mnuSCopyPointsInViewToClipboardAuto.Checked = Not mnuSCopyPointsInViewToClipboardAuto.Checked
    If mnuSCopyPointsInViewToClipboardAuto.Checked Then
        strResponse = InputBox("Enabling automatic copy of data points in current view to clipboard on zoom change.  Please enter the maximum number of data points to copy (0 to cancel):", "Max Auto Copy Points", "5000")
        mAutoCopyPointsMaxCount = val(strResponse)
        If Len(strResponse) = 0 Or mAutoCopyPointsMaxCount <= 0 Then
            mnuSCopyPointsInViewToClipboardAuto.Checked = False
        Else
        End If
    End If
End Sub

Private Sub mnuSCopyScansIncludeEmptyScans_Click()
    mnuSCopyScansIncludeEmptyScans.Checked = Not mnuSCopyScansIncludeEmptyScans.Checked
    SynchronizeCopyDataOptions
End Sub

Private Sub mnuSDlt_S_Click()
On Error Resume Next
frmS_DltPairs.Tag = nMyIndex
frmS_DltPairs.Show vbModal
End Sub

Private Sub mnuSDltLbl_S_Click()
On Error Resume Next
frmS_DltLblPairs.Tag = nMyIndex
frmS_DltLblPairs.Show vbModal
End Sub

Private Sub mnuSDltLblReport_Click()
'-----------------------------------
'generic report
'-----------------------------------
MsgBox "Not implemented. Coming soon.", vbOKOnly, glFGTU
End Sub

Private Sub mnuSDltLblUMC_Click()
    ShowPairsDeltaLabelUMCForm
End Sub

Private Sub mnuSDltUMC_Click()
    ShowPairsDeltaUMCForm
End Sub

Private Sub mnuSERAnalysis_Click()
On Error Resume Next
frmERAnalysis.Tag = nMyIndex
frmERAnalysis.Show vbModal
End Sub

Private Sub mnuSExcludeExc_Click()
frmExcludeWhatsNot.Tag = nMyIndex
frmExcludeWhatsNot.Show vbModal
End Sub

' No longer supported (March 2006)
''Begin VB.Menu mnuSExport
''   Caption = "&Export to FTICR_AMT DB"
''   Begin VB.Menu mnuSExportAll
''      Caption = "&All"
''   End
''   Begin VB.Menu mnuSExportCurrent
''      Caption = "&Current View"
''   End
''   Begin VB.Menu mnuSExportAllID
''      Caption = "All I&dentified"
''   End
''   Begin VB.Menu mnuSExportCurrentID
''      Caption = "Current View Id&entified"
''   End
''End

' No longer supported (March 2006)
''Private Sub mnuSExportAll_Click()
''Me.MousePointer = vbHourglass
''F_AExport nMyIndex, glScope.glSc_All
''Me.MousePointer = vbDefault
''End Sub
''
''Private Sub mnuSExportAllID_Click()
''Me.MousePointer = vbHourglass
''F_AExportIsoAMTHits nMyIndex, glScope.glSc_All
''Me.MousePointer = vbDefault
''End Sub
''
''Private Sub mnuSExportCurrent_Click()
''Me.MousePointer = vbHourglass
''F_AExport nMyIndex, glScope.glSc_Current
''Me.MousePointer = vbDefault
''End Sub
''
''Private Sub mnuSExportCurrentID_Click()
''Me.MousePointer = vbHourglass
''F_AExportIsoAMTHits nMyIndex, glScope.glSc_Current
''Me.MousePointer = vbDefault
''End Sub

Private Sub mnuSExportData_Click()
frmExportData.Show vbModal
End Sub

Private Sub mnuSExportResults_Click()
Dim MyExport As New frmExportResults
MyExport.CallerID = nMyIndex
MyExport.Show vbModal
Set MyExport = Nothing
End Sub

Private Sub mnuSIntCalLockMass_Click()
frmIntCalLM.Tag = nMyIndex
frmIntCalLM.Show vbModal
End Sub


Private Sub mnuSLbl_S_Click()
On Error Resume Next
frmS_LblPairs.Tag = nMyIndex
frmS_LblPairs.Show vbModal
End Sub

Private Sub mnuSLblUMC_Click()
    ShowPairsLabelUMCForm
End Sub

Private Sub mnuSLoadScope_Click()
Dim strScope As String
Dim strFilter As String
On Error Resume Next
strFilter = "All files(*.*)" & Chr(0) & "*.*" & Chr(0)
strScope = OpenFileAPIDlg(Me.hwnd, strFilter, 1, "Load Scope")
If Len(strScope) > 0 Then ICR2LSLoadScope strScope, 0
End Sub

''Private Sub mnuSLockMass_Click()
''frmMTLockMass.Tag = nMyIndex
''frmMTLockMass.Show vbModal
''End Sub

' No longer supported (March 2006)
''Private Sub mnuSMS_MSSearch_Click()
''frmMSMSSearch.Tag = nMyIndex
''frmMSMSSearch.Show vbModal
''End Sub


Private Sub mnuSUMCLockMass_Click()
frmUMCLockMass.Tag = nMyIndex
frmUMCLockMass.Show vbModal
End Sub

Private Sub mnuTileH_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileV_Click()
MDIForm1.Arrange vbTileVertical
End Sub

Public Function BuildGraph() As Boolean
On Error GoTo err_BuildGraph

'when this form starts loading we know there are some data in file
With GelData(nMyIndex)
     If .Preferences.IsoDataField <> glPreferences.IsoDataField Then
        .Preferences.IsoDataField = glPreferences.IsoDataField
        GelStatus(nMyIndex).Dirty = True
     End If
End With
InitDrawER nMyIndex

mnuViewChargeState.Enabled = GelDraw(nMyIndex).CSVisible
mnuViewChargeState.Checked = GelDraw(nMyIndex).CSVisible
mnuViewIsotopic.Enabled = GelDraw(nMyIndex).IsoVisible
mnuViewIsotopic.Checked = GelDraw(nMyIndex).IsoVisible
mnuViewPI.Enabled = GelData(nMyIndex).pICooSysEnabled
fgDisplay = glNormalDisplay
fgZOrder = glCSOnTop
mnuViewPI.Enabled = GelData(nMyIndex).pICooSysEnabled
'initialize coordinate system and redraw
If csMyCooSys Is Nothing Then Set csMyCooSys = New CooSysR
BuildCooSys
BuildGraph = True
MDIStatus False, "Done"
Exit Function

err_BuildGraph:
End Function

Private Sub mnuTools_Click()
    AdjustToolsMenu
End Sub

Private Sub mnuIDeviceCaps_Click()
Dim MyCaps As New frmDeviceCaps
MyCaps.ParentDC = picGraph.hDC
MyCaps.Show vbModal
Set MyCaps = Nothing
End Sub

Private Sub mnuViewAnalysisHistory_Click()
    ViewAnalysisHistory nMyIndex
End Sub

Private Sub mnuViewChargeState_Click()
mnuViewChargeState.Checked = Not mnuViewChargeState.Checked
If mnuViewChargeState.Checked Then
   GelDraw(nMyIndex).CSVisible = True
Else
   GelDraw(nMyIndex).CSVisible = False
End If
picGraph.Refresh
End Sub

'Private Sub mnuViewCorrelations_Click()
'frmCorrelations.Tag = nMyIndex
'frmCorrelations.Show vbModal
'End Sub

Private Sub mnuViewCSMap_Click()
ShowChargeStateMap
End Sub

Private Sub mnuViewDiffDisplay_Click()
ShowDiffDisplay
End Sub

Private Sub mnuViewDistributions_Click()
frmDistributions.Tag = nMyIndex
frmDistributions.Show vbModal
End Sub

Private Sub mnuViewFileInfo_Click()
frmDataInfo.Tag = -nMyIndex
frmDataInfo.Show
End Sub

Private Sub mnuViewFN_Click()
SetXAxisLabelType False
End Sub

Private Sub mnuViewIsotopic_Click()
mnuViewIsotopic.Checked = Not mnuViewIsotopic.Checked
If mnuViewIsotopic.Checked Then
   GelDraw(nMyIndex).IsoVisible = True
Else
   GelDraw(nMyIndex).IsoVisible = False
End If
picGraph.Refresh
End Sub

Private Sub mnuViewNormalDisplay_Click()
ShowNormalView
End Sub

Private Sub mnuViewPI_Click()
mnuViewPI.Checked = True
mnuViewFN.Checked = False
mnuViewNET.Checked = False
If Not csMyCooSys Is Nothing Then csMyCooSys.csType = glPICooSys
picGraph.Refresh
End Sub

Private Sub mnuViewRDGel_Click()
Me.MousePointer = vbHourglass
lAction = glNoAction
WriteRawDataFile nMyIndex
frmDataInfo.Tag = ""
frmDataInfo.Show
Me.MousePointer = vbDefault
End Sub

Private Sub mnuViewRDPek_Click()
Dim sMsg As String
Dim Res
lAction = glNoAction
If FileExists(GelData(nMyIndex).FileName) Then
   Me.MousePointer = vbHourglass
   frmDataInfo.Tag = nMyIndex
   frmDataInfo.Show
   Me.MousePointer = vbDefault
Else
   sMsg = "Original PEK/CSV/mzXML/mzData file not found. Do you want to see data structured as in GEL file instead?"
   Res = MsgBox(sMsg, vbOKCancel)
   If Res = vbOK Then mnuViewRDGel_Click
End If
End Sub

Private Sub mnuViewSelToolBox_Click()
fgSelBoxVisible = Not fgSelBoxVisible
mnuViewSelToolBox.Checked = fgSelBoxVisible
SelToolBox1.Visible = fgSelBoxVisible
End Sub

Private Sub mnuViewToolbar_Click()
If mnuViewToolbar.Checked Then
   SyncMenuCmdToolbar False
Else
   SyncMenuCmdToolbar True
End If
MDIForm1.picToolbar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuViewTracker_Click()
On Error Resume Next
If mnuViewTracker.Checked = True Then
   glTracking = False
   Unload frmTracker
Else
   SyncMenuCmdTracker True
   frmTracker.Show
End If
End Sub

Private Sub mnuViewZOrder_Click(Index As Integer)
mnuViewZOrder(0).Checked = Not mnuViewZOrder(0).Checked
mnuViewZOrder(1).Checked = Not mnuViewZOrder(1).Checked
If mnuViewZOrder(0).Checked Then
   fgZOrder = glCSOnTop
Else
   fgZOrder = glIsoOnTop
End If
picGraph.Refresh
End Sub

Private Sub mnuVisual2DGelHelp_Click()
    MsgBox "Please see the Powerpoint Help file (e.g. VIPER_HelpFile_v3.20.ppt).  If inside PNNL, you can find this file at \\floyd\software\VIPER\ .  If outside PNNL, please visit http://ncrr.pnl.gov/software/ for help on using this software.", vbInformation + vbOKOnly, glFGTU
End Sub

Private Sub mnuIMTDBAnalysisInfo_Click()
Dim Msg As String
If Not GelAnalysis(nMyIndex) Is Nothing Then
   Msg = GelAnalysis(nMyIndex).GetJobInfo
   If Len(Msg) <= 0 Then
      Msg = "This analysis has link with MT tag database but was not loaded as analysis."
   End If
Else
   Msg = "Current gel not loaded as analysis."
End If
MsgBox Msg, vbOKOnly, glFGTU
End Sub

Private Sub mnuIMTDBProcessingInfo_Click()
Dim Msg As String
If Not GelAnalysis(nMyIndex) Is Nothing Then
   Msg = GelAnalysis(nMyIndex).GetParameters
   If Len(Msg) <= 0 Then
      Msg = "This analysis has link with MT tag database but was not loaded as analysis."
   End If
Else
   Msg = "Current gel not loaded as analysis."
End If
MsgBox Msg, vbOKOnly, glFGTU
End Sub

Private Sub mnuVMTDisplay_Click()
If AMTCnt > 0 Then
   Call Display0        'always recreate database display
Else
   MsgBox "No MT tags(database peptides) found.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuIMTSettings_Click()
If Len(CurrLegacyMTDatabase) > 0 Then
    MsgBox "Legacy database: " & CurrLegacyMTDatabase, vbOKOnly, glFGTU
ElseIf Not GelAnalysis(nMyIndex) Is Nothing Then
   frmDataInfo.Tag = nMyIndex + glMIL
   frmDataInfo.Show
Else
    MsgBox "No MT tag database loaded.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuVOverlay_Click()
On Error Resume Next
Load frmGraphOverlay
End Sub


Private Sub mnuVUMC_Click()
mnuVUMC.Checked = Not mnuVUMC.Checked
GelUMCDraw(nMyIndex).Visible = mnuVUMC.Checked
bNeedToUpdate = True
csMyCooSys.CoordinateDraw
picGraph.Refresh
End Sub

Private Sub mnuWindowSizeToDim_Click(Index As Integer)
    SetWindowSize CInt(Index)
End Sub

Private Sub mnuWindowSizeUpdateAll_Click()
    mnuWindowSizeUpdateAll.Checked = Not mnuWindowSizeUpdateAll.Checked
End Sub

Private Sub picGraph_Click()
LastType = HotType
Select Case HotType
Case glNoType
    'if selection not protected
    If ((Not fgSelProtected) And (MouseButton = vbLeftButton)) Then
        'and there is something selected
        If GelSel.CSSelCnt > 0 Or GelSel.IsoSelCnt > 0 Then
            ClearSelectedPoints True
        End If
    End If
    LastID = -1
Case glCSType
    LastID = HotID
    If (MouseButton = vbLeftButton) Then
        GelSel.AddToCSSelection HotID
        GelDrawSelectionAdd nMyIndex, glCSType, HotID
    End If
Case glIsoType
    LastID = HotID
    If (MouseButton = vbLeftButton) Then
      GelSel.AddToIsoSelection HotID
      GelDrawSelectionAdd nMyIndex, glIsoType, HotID
    End If
End Select
End Sub

Private Sub picGraph_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
Source.Move x, y
End Sub

Private Sub picGraph_GotFocus()
lAction = glNoAction
If bPaint Then picGraph.Refresh
bPaint = False
End Sub

Private Sub picGraph_KeyUp(KeyCode As Integer, Shift As Integer)
Dim TextToCopy As String
If (Shift And vbCtrlMask) > 0 Then
   Select Case KeyCode
   Case vbKeyI  'copy Identity field if Hot
       Select Case HotType
       Case glCSType
          TextToCopy = GelData(nMyIndex).CSData(HotID).MTID
       Case glIsoType
          TextToCopy = GelData(nMyIndex).IsoData(HotID).MTID
       End Select
       If Len(TextToCopy) > 0 Then
          Clipboard.SetText TextToCopy
       Else
          Clipboard.Clear
       End If
   End Select
End If
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ReDim paPoints(0)
paPoints(0).x = x
paPoints(0).y = y
DevLogConversion ConvDPLP, 1
If HotType = glNoType Then
   If Button = vbLeftButton Then lAction = glActionZoom
Else
   lAction = glActionHit
   frmTracker.Tag = nMyIndex
End If
MouseButton = Button
Select Case Button
Case vbLeftButton
    ZoomActionStart x, y, paPoints()
Case vbRightButton
    If lAction = glActionZoom Then
        ZoomActionCancel
    Else
        AdjustToolsMenu
        lAction = glNoAction
        Me.PopupMenu mnuTools
    End If
Case Else
  lAction = glNoAction
End Select
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Update the zoom box dimensions in the status window
 Dim ZoomX1 As Double, ZoomY1 As Double, ZoomX2 As Double, ZoomY2 As Double
 Dim dblNET As Double
' MonroeMod: changed from (0) to (1)
ReDim paPoints(1)

If csMyCooSys Is Nothing Then Exit Sub

paPoints(0).x = x
paPoints(0).y = y
DevLogConversion ConvDPLP, 1
GetHotSpot nMyIndex, paPoints(0).x, paPoints(0).y, HotID, HotType

If glTracking Then SetTrackingLabels nMyIndex, HotType, HotID
If lAction = glActionZoom Then
   picGraph.Line (gbZoomX1, gbZoomY1)-(gbZoomX2, gbZoomY2), , B
   gbZoomX2 = x
   gbZoomY2 = y
   picGraph.Line (gbZoomX1, gbZoomY1)-(gbZoomX2, gbZoomY2), , B
   
   If glTracking Then
        paPoints(0).x = gbZoomX1
        paPoints(0).y = gbZoomY1
        paPoints(1).x = gbZoomX2
        paPoints(1).y = gbZoomY2
        DevLogConversion ConvDPLP, 2
        
        ZoomX1 = paPoints(0).x
        ZoomY1 = paPoints(0).y
        ZoomX2 = paPoints(1).x
        ZoomY2 = paPoints(1).y
        
        csMyCooSys.LPToRP ZoomX1, ZoomY1, ZoomX2, ZoomY2
        frmTracker.lblIdentity = GetZoomBoxDimensions(ZoomX1, ZoomY1, ZoomX2, ZoomY2, True, True, True)
   
   End If
Else
    ' Update the current position in the tracker box
    ZoomX1 = paPoints(0).x
    ZoomY1 = paPoints(0).y
    csMyCooSys.LPToRP ZoomX1, ZoomY1, 0, 0
        
    mLastCursorPosFN = ZoomX1
    mLastCursorPosMass = ZoomY1
        
    If glTracking And HotType = glNoType Then
        dblNET = ScanToGANET(nMyIndex, CLng(ZoomX1))
        frmTracker.lblIdentity = FormatCurrentPosition(ZoomX1, ZoomY1, True, True) & " (NET " & Format(dblNET, "0.000") & ")"
    End If
End If
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then ZoomActionEnd
End Sub

Public Sub ActivateGraph(Optional blnShowFilterDialogIfLoading As Boolean = True)
    
    MDIForm1.ProperToolbar True
    lAction = glNoAction
    If bLoad Then
       glTracking = False
       If Len(Me.Tag) = 0 Then Me.Tag = 0
       nMyIndex = Abs(Me.Tag)
       If nMyIndex = 0 Then mnuEditResidualDisplay.Enabled = True
       If Not InitDrawData(nMyIndex) Then
          MsgBox "Error initializing drawing structures.", vbOKOnly, glFGTU
          Unload Me
          Exit Sub
       End If
       If BuildGraph Then
          If blnShowFilterDialogIfLoading Then
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                frmFilter.Tag = nMyIndex
                frmFilter.InitializeControls True
            Else
               ShowFilterForm
            End If
          End If
       Else
          If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Error building a gel. Probable cause - invalid file format.", vbOKOnly, glFGTU
            MDIStatus False, "Done"
          Else
            Debug.Print "Error in frmGraphsActivateGraph: " & Err.Description
            Debug.Assert False
            LogErrors Err.Number, "frmGraph->ActiveGraph"
          End If
          Exit Sub
       End If
       If GelData(nMyIndex).Fileinfo = UMRFileInfo Then
          mnuViewRDPek.Enabled = False
       Else
          mnuViewRDPek.Enabled = True
       End If
       bLoad = False
       bResize = True
       If IsWinLoaded(TrackerCaption) Then glTracking = True
       Form_Resize
    Else
       If fgRebuild Then
          fgRebuild = False
          BuildGraph
       End If
    End If

    UpdateTICPlotAndFeatureBrowsersIfNeeded
    
End Sub

Private Sub AdjustToolsMenu()
    Dim blnEnabled As Boolean
    
    If GelSel.CSSelCnt > 0 Or GelSel.IsoSelCnt > 0 Then blnEnabled = True
    mnu2lsExcludeSelection.Enabled = blnEnabled
    mnu2lsExcludeAllButSelection.Enabled = blnEnabled
    mnu2lsShowSpectrum.Enabled = blnEnabled
    
    ' No longer supported (March 2006)
    '' mnu2lsShowAMTRecord.Enabled = blnEnabled
End Sub

Public Sub CopyAllPointsInView(Optional ByVal lngMaxPointsCountToCopy As Long = -1, Optional blnPromptForFileToExportTo As Boolean = False, Optional strFilePathForce As String = "")
    
    Dim i As Long, j As Long
    Dim dblMW As Double, dblMtoZ As Double, dblAbu As Double
    Dim dblAbuIReportMWMono As Double, dblAbuIReport2Da As Double
    Dim dblNET As Double
    Dim lngIonIndex As Long
    Dim intCharge As Integer
    Dim dblFit As Double
    Dim dblExpressionRatio As Double
    Dim lngFN As Long, lngFNPrevious As Long
    Dim lngIonPointerArray() As Long           ' 1-based array
    Dim lngIonCount As Long
    
    Dim blnCSPoints As Boolean
    
    ' The following two arrays are used to look up the mass of each MT tag, given the MT tag ID
    ' If the user specified a mass modification (like alkylation, ICAT, or N15), then the standard mass
    '  for the MT tag will not be correct; we won't try to correct for this since it would require a bit of guessing, and
    '  the user can get this information using the official Results by UMC or Results by Ion report on the search form
    Dim blnIncludeAMTMassDetails As Boolean
    Dim lngAMTID() As Long, lngAMTIDPointer() As Long
    Dim lngIndex As Long
    Dim QSL As New QSLong
    
    Dim lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long
    Dim dblAMTMW As Double, dblAMTNet As Double, dblAMTNetStDev As Double
    
    Dim strDBMatchList As String, strDBMatchOrMatches As String
    Dim strUMCIndices As String
    Dim lngCharLoc As Long
    
    Dim lngCurrIDCnt As Long
    Dim udtCurrIDMatchStats() As udtUMCMassTagMatchStats        ' 0-based array
    
    Dim strFlattenedList As String
    Dim strExport() As String
    Dim lngExportCount As Long
    Dim lngExportCountDimmed As Long
    
    Dim eResponse As VbMsgBoxResult
    
    Dim strFilePath As String
    Dim strSepChar As String
    Dim OutFileNum As Integer
    
    Dim blnIncludeEmptyScans As Boolean
    Dim blnOutputAllMatchesOnOneLine As Boolean
    Dim blnIReportData As Boolean
    
On Error GoTo CopyAllPointsInViewErrorHandler

    ' Retrieve an array of the ion indices of the ions currently "In Scope"
    ' Note that GetISScope will ReDim lngIonPointerArray() automatically
    lngIonCount = GetISScope(nMyIndex, lngIonPointerArray(), glScope.glSc_Current)
    
    If lngIonCount = 0 Then
        ' No isotopic points in view; what about CS points?
        lngIonCount = GetCSScope(nMyIndex, lngIonPointerArray(), glScope.glSc_Current)
        If lngIonCount > 0 Then
            blnCSPoints = True
        End If
    End If
    
    If lngIonCount > 0 Then
        If blnPromptForFileToExportTo Then
            ' Save to a file
            strFilePath = SelectFile(Me.hwnd, "Enter file name to copy points to", "", True, "DataInView.txt", , 2)
            If Len(strFilePath) = 0 Then Exit Sub
        ElseIf Len(strFilePathForce) > 0 Then
            strFilePath = strFilePathForce
        End If
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No points found in the current view", vbInformation + vbOKOnly, glFGTU
        End If
        Exit Sub
    End If
    
    blnIncludeEmptyScans = mnuSCopyScansIncludeEmptyScans.Checked
    blnOutputAllMatchesOnOneLine = mnuSCopyPointsInViewOneHitPerLine.Checked
    
    If lngMaxPointsCountToCopy > 0 Then
        If lngIonCount > lngMaxPointsCountToCopy Then lngIonCount = lngMaxPointsCountToCopy
    End If
    
    frmProgress.InitializeForm "Prepraring data for copying", 0, lngIonCount, False, False, False
    
    lngExportCount = 0
    lngExportCountDimmed = lngIonCount
    
    ReDim strExport(lngExportCountDimmed)
    
    strSepChar = LookupDefaultSeparationCharacter()
    
    If AMTCnt = 0 Then
        ' See if any of the ions in memory (regardless of scope) have matches defined
        ' If any do, then attempt to connect to the database, provided mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked = True
        
        If mnuSCopyPointsInViewIncludeSearchResults.Checked Then
            For lngIonIndex = 1 To GelData(nMyIndex).IsoLines
                strDBMatchList = FixNull(GelData(nMyIndex).IsoData(lngIonIndex).MTID)
                If Len(strDBMatchList) > 0 Then
                    eResponse = MsgBox("One or more points contains database matches.  However, the appropriate ORF information is not currently loaded.  Load it now (if No, then re-enable this option using Edit->Copy Points in View->Include DB Search Matches)?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load MT tags")
                    If eResponse = vbYes Then
                        ConfirmMassTagsAndInternalStdsLoaded Me, nMyIndex, True
                    Else
                        mnuSCopyPointsInViewIncludeSearchResults.Checked = False
                    End If
                    Exit For
                End If
            Next lngIonIndex
        End If
    End If
    
    If AMTCnt > 0 Then
        blnIncludeAMTMassDetails = True
            
        ' Construct the MT tag mass lookup arrays
        ReDim lngAMTID(0 To AMTCnt - 1)
        ReDim lngAMTIDPointer(0 To AMTCnt - 1)
        
        For lngIndex = 0 To AMTCnt - 1
            lngAMTID(lngIndex) = AMTData(lngIndex + 1).ID        ' Note: AMTData() is 1-based and is sorted based on AMT MW
            lngAMTIDPointer(lngIndex) = lngIndex + 1
        Next lngIndex
        
        If Not QSL.QSAsc(lngAMTID(), lngAMTIDPointer()) Then
            ' This is unexpected
            Debug.Assert False
        End If
        
        Set QSL = Nothing
    End If
    
    If Len(strFilePath) > 0 Then
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
    End If
    
    With GelData(nMyIndex)
        strExport(0) = "Scan" & strSepChar & "NET" & strSepChar & "Index" & strSepChar & "Abundance" & strSepChar
        If Not blnCSPoints And ((.DataStatusBits And GEL_DATA_STATUS_BIT_IREPORT) = GEL_DATA_STATUS_BIT_IREPORT) Then
            blnIReportData = True
            strExport(0) = strExport(0) & "Monoiso Mass Abu" & strSepChar & "Monoiso Mass +2 Da Abu" & strSepChar
        Else
            blnIReportData = False
        End If
        
        If blnCSPoints Then
            strExport(0) = strExport(0) & "Charge State" & strSepChar & "Fit" & strSepChar & "ExpressionRatio" & strSepChar & "M/Z of Most Abu" & strSepChar & "MW" & strSepChar & "UMC Indices"
        Else
            strExport(0) = strExport(0) & "Charge State" & strSepChar & "Fit" & strSepChar & "ExpressionRatio" & strSepChar & "M/Z of Most Abu" & strSepChar & GetIsoDescription(.Preferences.IsoDataField) & strSepChar & "UMC Indices"
        End If
        
        If blnIncludeAMTMassDetails Then
            strExport(0) = strExport(0) & strSepChar & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagNET" & strSepChar & "MassTagNETStDev" & strSepChar & "SLiC Score" & strSepChar & "MassDiff (ppm)"
        Else
            strExport(0) = strExport(0) & strSepChar & "DB Matches"
        End If
        
        If Len(strFilePath) > 0 Then Print #OutFileNum, strExport(0)
        
        lngExportCount = 1
        For i = 1 To lngIonCount
            If blnCSPoints Then
                lngFN = .CSData(lngIonPointerArray(i)).ScanNumber
            Else
                lngFN = .IsoData(lngIonPointerArray(i)).ScanNumber
            End If
            
            If lngFN > lngFNPrevious Then
                If lngFNPrevious <> 0 And blnIncludeEmptyScans Then
                    ' Include blank lines for scans without any points; useful when lining up data in Excel
                    If lngExportCount + (lngFN - lngFNPrevious - 2) >= lngExportCountDimmed Then
                        lngExportCountDimmed = lngExportCountDimmed + (lngFN - lngFNPrevious - 2) + 1
                        ReDim Preserve strExport(lngExportCountDimmed)
                    End If
                    
                    For j = lngFNPrevious + 1 To lngFN - 1
                        strExport(lngExportCount) = j           ' Record the scan number
                        lngExportCount = lngExportCount + 1
                    Next j
                End If
                lngFNPrevious = lngFN
            End If
            
            lngIonIndex = lngIonPointerArray(i)
            If blnCSPoints Then
                dblAbu = .CSData(lngIonIndex).Abundance
                intCharge = .CSData(lngIonIndex).Charge
                dblFit = 0
                dblExpressionRatio = 0
                dblMtoZ = 0
                dblMW = .CSData(lngIonIndex).AverageMW
                
                strUMCIndices = ""
                strDBMatchList = .CSData(lngIonIndex).MTID
            
            Else
                dblAbu = .IsoData(lngIonIndex).Abundance
            
                If blnIReportData Then
                    dblAbuIReportMWMono = .IsoData(lngIonIndex).IntensityMono
                    dblAbuIReport2Da = .IsoData(lngIonIndex).IntensityMonoPlus2
                End If
                
                intCharge = .IsoData(lngIonIndex).Charge
                dblFit = .IsoData(lngIonIndex).Fit
                dblExpressionRatio = .IsoData(lngIonIndex).ExpressionRatio
                
                ' Note: the m/z value stored in .IsoData() corresponds to the m/z of the most abundant peak in the isotope envelope, and therefore may not agree perfectly with dblMW
                dblMtoZ = .IsoData(lngIonIndex).MZ
                dblMW = GetIsoMass(.IsoData(lngIonIndex), .Preferences.IsoDataField)
                
                strUMCIndices = ConstructUMCIndexList(nMyIndex, lngIonIndex, glIsoType)
                strDBMatchList = .IsoData(lngIonIndex).MTID
            
            End If
            
            dblNET = ScanToGANET(nMyIndex, lngFN)
            
            Do
                If blnOutputAllMatchesOnOneLine Or Len(strDBMatchList) = 0 Then
                    strDBMatchOrMatches = strDBMatchList
                    strDBMatchList = ""
                Else
                    lngCharLoc = InStr(strDBMatchList, glARG_SEP)
                    If lngCharLoc > 0 Then
                        strDBMatchOrMatches = Left(strDBMatchList, lngCharLoc - 1)
                        strDBMatchList = Trim(Mid(strDBMatchList, lngCharLoc + 1))
                    Else
                        strDBMatchOrMatches = strDBMatchList
                        strDBMatchList = ""
                    End If
                End If
                
                strExport(lngExportCount) = lngFN & strSepChar & Format$(dblNET, "0.0000") & strSepChar & lngIonIndex & strSepChar & Round(dblAbu, 0) & strSepChar
                
                If blnIReportData Then
                    strExport(lngExportCount) = strExport(lngExportCount) & Round(dblAbuIReportMWMono, 0) & strSepChar & Round(dblAbuIReport2Da, 0) & strSepChar
                End If
                
                strExport(lngExportCount) = strExport(lngExportCount) & Trim(intCharge) & strSepChar & Format(dblFit, "0.000") & strSepChar & Format(dblExpressionRatio, "0.000000") & strSepChar & Format(dblMtoZ, "0.0000") & strSepChar & Format(dblMW, "0.0000") & strSepChar & strUMCIndices & strSepChar
                
                If blnIncludeAMTMassDetails And Not blnOutputAllMatchesOnOneLine And Len(strDBMatchOrMatches) > 0 Then
                    ' Parse out the match details from strDBMatchOrMatches
                    
                    lngCurrIDCnt = 0
                    ExtractMTHitsFromMatchList strDBMatchOrMatches, False, lngCurrIDCnt, udtCurrIDMatchStats(), True
                    
                    If lngCurrIDCnt > 0 Then
                        Debug.Assert lngCurrIDCnt = 1
                        
                        If mnuSCopyPointsInViewIncludeSearchResults.Checked And _
                           blnIncludeAMTMassDetails And udtCurrIDMatchStats(0).IDIndex >= 0 And _
                           Not udtCurrIDMatchStats(0).IDIsInternalStd Then
                            lngMassTagIndexPointer = BinarySearchLng(lngAMTID(), udtCurrIDMatchStats(0).IDIndex, 0, AMTCnt - 1)
                            
                            If lngMassTagIndexPointer >= 0 Then
                                lngMassTagIndexOriginal = lngAMTIDPointer(lngMassTagIndexPointer)
                                dblAMTMW = AMTData(lngMassTagIndexOriginal).MW
                                dblAMTNet = AMTData(lngMassTagIndexOriginal).NET
                                dblAMTNetStDev = AMTData(lngMassTagIndexOriginal).NETStDev
                                Debug.Assert AMTData(lngMassTagIndexOriginal).ID = udtCurrIDMatchStats(0).IDIndex
                            Else
                                dblAMTMW = 0
                                dblAMTNet = 0
                                dblAMTNetStDev = 0
                            End If
                        Else
                            dblAMTMW = 0
                            dblAMTNet = 0
                            dblAMTNetStDev = 0
                        End If
                        
                        
                        With udtCurrIDMatchStats(0)
                            ' "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagNET" & strSepChar & "MassTagNETStDev" & strSepChar & "SLiC Score" & strSepChar & "MassDiff (ppm)"
                            strExport(lngExportCount) = strExport(lngExportCount) & .IDIndex & strSepChar & dblAMTMW & strSepChar & dblAMTNet & strSepChar & dblAMTNetStDev & strSepChar & .SLiCScore & strSepChar & .MassDiffPPM
                        End With
                        
                    Else
                        ' This will happen if the match was an Internal Standard match
                        
                    End If
                Else
                    ' Do not include AMT Mass details, either because blnIncludeAMTMassDetails is false or blnOutputAllMatchesOnOneLine = True
                    strExport(lngExportCount) = strExport(lngExportCount) & strDBMatchOrMatches
                End If
                
                If Len(strFilePath) > 0 Then
                    Print #OutFileNum, strExport(lngExportCount)
                End If
                
                lngExportCount = lngExportCount + 1
                If lngExportCount >= lngExportCountDimmed Then
                    lngExportCountDimmed = lngExportCountDimmed + 500
                    ReDim Preserve strExport(lngExportCountDimmed)
                End If
            Loop While Len(strDBMatchList) > 0
            
            If i Mod 100 = 0 Then
                frmProgress.UpdateProgressBar i
                If KeyPressAbortProcess > 1 Then Exit For
            End If
        Next i
    End With
    
    If Len(strFilePath) > 0 Then
        ' Close the file
        Close OutFileNum
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            If KeyPressAbortProcess > 1 Then
                MsgBox "Copy data cancelled.", vbInformation + vbOKOnly, glFGTU
            Else
                MsgBox "Data saved to file: " & strFilePath, vbInformation + vbOKOnly, glFGTU
            End If
        End If
    Else
        If KeyPressAbortProcess > 1 Then
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "Copy data cancelled.", vbInformation + vbOKOnly, glFGTU
            End If
        Else
            ' Copy to the clipboard
            strFlattenedList = FlattenStringArray(strExport(), lngExportCount, vbCrLf)
            
            On Error Resume Next
            Clipboard.Clear
            Clipboard.SetText strFlattenedList, vbCFText
        End If
    End If
    
    frmProgress.HideForm
    Exit Sub

CopyAllPointsInViewErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error occurred in CopyAllPointsInView:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        LogErrors Err.Number, "CopyAllPointsInView"
    End If
    frmProgress.HideForm
    
End Sub

Public Sub CopyAllUMCsInView(Optional ByVal lngMaxPointsCountToCopy As Long = -1, Optional blnPromptForFileToExportTo As Boolean = False, Optional strFilePathForce As String = "")
    
    Const EXPORT_STEP_COUNT = 5
    
    Dim lngAllUMCCount As Long
    Dim blnUMCPresent() As Boolean          ' Records whether or not each UMC is present
                                            ' Corresponds to data in GelUMC().UMCs() and is thus 0-based
    Dim ClsStat() As Double                 ' Holds Stats on each UMC, including min and max scan number
    
    ' The Match Stats variables
    ' Note that a given UMC will be present multiple times in udtUMCsInView if the UMC had multiple MT tag matches
    Dim lngUMCsInViewCount As Long, lngUMCInViewCountDimmed As Long
    Dim udtUMCsInView() As udtUMCMassTagMatchStats          ' 0-based array
    
    ' The following variable will hold the indices of the ions currently In View (in the current scope)
    
    Dim lngCSPointerArray() As Long            ' 1-based array (dictated by GetCSScope)
    Dim lngIsoPointerArray() As Long            ' 1-based array (dictated by GetISScope)
    Dim lngCSCount As Long
    Dim lngIsoCount As Long
    Dim lngIonIndex As Long
    Dim lngUMCIndex As Long, lngUMCIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long
    
    Dim lngScanClassRep As Long
    Dim dblGANETClassRep As Double, dblAMTMW As Double, dblAMTNet As Double, dblAMTNetStDev As Double
    Dim sngPeptideProphetProbability As Single
    
    ' The following two arrays are used to look up the mass of each MT tag, given the MT tag ID
    ' If the user specified a mass modification (like alkylation, ICAT, or N15), then the standard mass
    '  for the MT tag will not be correct; we won't try to correct for this since it would require a bit of guessing, and
    '  the user can get this information using the official Results by UMC or Results by Ion report on the search form
    Dim blnIncludeAMTMass As Boolean
    Dim lngAMTID() As Long, lngAMTIDPointer() As Long
    Dim lngIndex As Long
    Dim QSL As New QSLong
    
    Dim strDBMatchList As String

    Dim strLineOut As String, strLineOutMiddle As String, strLineOutEnd As String
    Dim strMinMaxCharges As String
    Dim strFlattenedList As String
    Dim strExport() As String
    Dim lngExportCount As Long
    Dim lngExportCountDimmed As Long
    Dim lngProgessStepCount As Long
    
    Dim eResponse As VbMsgBoxResult
    Dim blnMatchFound As Boolean
        
    Dim strFilePath As String
    Dim strSepChar As String
    Dim OutFileNum As Integer
    
    Dim lngPairIndex As Long
    
    Dim objP1IndFastSearch As FastSearchArrayLong
    Dim objP2IndFastSearch As FastSearchArrayLong
    Dim blnPairsPresent As Boolean
    Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
    Dim udtPairMatchStats() As udtPairMatchStatsType
    
On Error GoTo CopyAllUMCsInViewErrorHandler

    lngAllUMCCount = GelUMC(nMyIndex).UMCCnt
    
    If lngAllUMCCount > 0 Then
        If blnPromptForFileToExportTo Then
            ' Save to a file
            strFilePath = SelectFile(Me.hwnd, "Enter file name to copy UMC's to", "", True, "UMCsInView.txt", , 2)
            If Len(strFilePath) = 0 Then Exit Sub
        ElseIf Len(strFilePathForce) > 0 Then
            strFilePath = strFilePathForce
        End If
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No UMC's are present in memory", vbInformation + vbOKOnly, glFGTU
        End If
        Exit Sub
    End If
    
    lngProgessStepCount = 0
    frmProgress.InitializeForm "Preparing data", 0, EXPORT_STEP_COUNT, False, True, False
    frmProgress.InitializeSubtask "Generating UMC statistics", 0, 1
    
    lngAllUMCCount = UMCStatistics1(nMyIndex, ClsStat())
    Debug.Assert lngAllUMCCount = GelUMC(nMyIndex).UMCCnt
    
    lngProgessStepCount = lngProgessStepCount + 1
    frmProgress.UpdateProgressBar lngProgessStepCount
    
    ' Reserve space in the UMCPresent array
    ReDim blnUMCPresent(lngAllUMCCount)
    
    ' Initialize the PairIndex lookup objects
    blnPairsPresent = PairIndexLookupInitialize(nMyIndex, objP1IndFastSearch, objP2IndFastSearch)
        
    strSepChar = LookupDefaultSeparationCharacter()
    
    If AMTCnt = 0 Then
        ' See if any of the ions in memory (regardless of scope) have matches defined
        ' If any do, then attempt to connect to the database, provided mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked = True
        
        If mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked Then
            ' Check the CS points
            For lngIonIndex = 1 To GelData(nMyIndex).CSLines
                strDBMatchList = FixNull(GelData(nMyIndex).CSData(lngIonIndex).MTID)
                If Len(strDBMatchList) > 0 Then
                    blnMatchFound = True
                    Exit For
                End If
            Next lngIonIndex
            
            If Not blnMatchFound Then
                ' Check the Iso points
                For lngIonIndex = 1 To GelData(nMyIndex).IsoLines
                    strDBMatchList = FixNull(GelData(nMyIndex).IsoData(lngIonIndex).MTID)
                    If Len(strDBMatchList) > 0 Then
                        blnMatchFound = True
                        Exit For
                    End If
                Next lngIonIndex
            End If
            
            If blnMatchFound Then
                eResponse = MsgBox("One or more UMC's contains database matches.  However, the appropriate ORF (protein) information is not currently loaded.  Load it now (if No, then re-enable this option using Edit->Copy UMC's in View->Include DB Search Matches)?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load MT tags")
                If eResponse = vbYes Then
                    ConfirmMassTagsAndInternalStdsLoaded Me, nMyIndex, True
                Else
                    mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked = False
                End If
            End If
        End If
    End If
    
    If AMTCnt > 0 Then
        blnIncludeAMTMass = True
            
        ' Construct the MT tag mass lookup arrays
        ReDim lngAMTID(0 To AMTCnt - 1)
        ReDim lngAMTIDPointer(0 To AMTCnt - 1)
        
        For lngIndex = 0 To AMTCnt - 1
            lngAMTID(lngIndex) = AMTData(lngIndex + 1).ID        ' Note: AMTData() is 1-based and is sorted based on AMT MW
            lngAMTIDPointer(lngIndex) = lngIndex + 1
        Next lngIndex
        
        If Not QSL.QSAsc(lngAMTID(), lngAMTIDPointer()) Then
            ' This is unexpected
            Debug.Assert False
        End If
        
        Set QSL = Nothing
    End If
    
    ' Step 1: Retrieve an array of the ion indices of the ions currently "In Scope"
    ' Note that GetCSScope and GetISScope will ReDim lngCSPointerArray() and lngIsoPointerArray() automatically
    lngCSCount = GetCSScope(nMyIndex, lngCSPointerArray(), glSc_Current)
    lngIsoCount = GetISScope(nMyIndex, lngIsoPointerArray(), glScope.glSc_Current)
    
    lngProgessStepCount = lngProgessStepCount + 1
    frmProgress.UpdateProgressBar lngProgessStepCount
    frmProgress.InitializeSubtask "Finding UMC's in view", 0, lngAllUMCCount
    
    ' Step 2: Set blnUMCPresent() to True for the UMC's that the ions currently "In Scope" belong to
    lngUMCInViewCountDimmed = 0
    For lngIonIndex = 1 To lngCSCount
        With GelDataLookupArrays(nMyIndex).CSUMCs(lngCSPointerArray(lngIonIndex))
            For lngUMCIndex = 0 To .UMCCount - 1
                If Not blnUMCPresent(.UMCs(lngUMCIndex)) Then
                    blnUMCPresent(.UMCs(lngUMCIndex)) = True
                    lngUMCInViewCountDimmed = lngUMCInViewCountDimmed + 1
                End If
            Next lngUMCIndex
        End With
    Next lngIonIndex
    
    For lngIonIndex = 1 To lngIsoCount
        With GelDataLookupArrays(nMyIndex).IsoUMCs(lngIsoPointerArray(lngIonIndex))
            For lngUMCIndex = 0 To .UMCCount - 1
                If Not blnUMCPresent(.UMCs(lngUMCIndex)) Then
                    blnUMCPresent(.UMCs(lngUMCIndex)) = True
                    lngUMCInViewCountDimmed = lngUMCInViewCountDimmed + 1
                End If
            Next lngUMCIndex
        End With
    Next lngIonIndex

    ' Reserve space in the Match Stats array
    lngUMCsInViewCount = 0
    If lngUMCInViewCountDimmed < 100 Then lngUMCInViewCountDimmed = 100
    ReDim udtUMCsInView(lngUMCInViewCountDimmed)

    ' Step 3: Populate the Match Stats arrays based on the members of each UMC, regardless of whether or not they're in scope
    For lngUMCIndex = 0 To lngAllUMCCount - 1
        If blnUMCPresent(lngUMCIndex) Then
            ExtractMTHitsFromUMCMembers nMyIndex, lngUMCIndex, False, udtUMCsInView, lngUMCsInViewCount, lngUMCInViewCountDimmed, Not mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked, False
            ExtractMTHitsFromUMCMembers nMyIndex, lngUMCIndex, True, udtUMCsInView, lngUMCsInViewCount, lngUMCInViewCountDimmed, Not mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked, False
        End If
        
        If lngUMCIndex Mod 100 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngUMCIndex
            If KeyPressAbortProcess > 1 Then
                frmProgress.HideForm
                Exit Sub
            End If
        End If
    Next lngUMCIndex
    
    If lngUMCsInViewCount = 0 Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            If blnPromptForFileToExportTo Then
                MsgBox "No UMC's found in the current view; nothing was saved to disk.", vbInformation + vbOKOnly, glFGTU
            Else
                MsgBox "No UMC's found in the current view", vbInformation + vbOKOnly, glFGTU
            End If
        End If
        frmProgress.HideForm
        Exit Sub
    End If
    
    lngProgessStepCount = lngProgessStepCount + 1
    frmProgress.UpdateProgressBar lngProgessStepCount
    frmProgress.InitializeSubtask "Preparing UMC Match stats", 0, lngUMCsInViewCount
    
    lngExportCount = 0
    lngExportCountDimmed = lngCSCount + lngIsoCount
    ReDim strExport(lngExportCountDimmed)
    
    ' UMCIndex; ScanStart; ScanEnd; ScanClassRep; GANETClassRep; UMCMonoMW; UMCMWStDev; UMCMWMin; UMCMWMax; UMCAbundance; ClassStatsChargeBasis; ChargeStateMin; ChargeStateMax; UMCMZForChargeBasis; UMCMemberCount; UMCMemberCountUsedForAbu; UMCAverageFit; PairIndex; ExpressionRatio; ExpressionRatioStDev; ExpressionRatioBasisCount; MultiMassTagHitCount; MassTagID; MassTagMonoMW; MassTagNET; MassTagNETStDev; SLiC Score; DelSLiC; MemberCountMatchingMassTag; IsInternalStdMatch; PeptideProphetProbability
    strLineOut = "UMCIndex" & strSepChar & "ScanStart" & strSepChar & "ScanEnd" & strSepChar & "ScanClassRep" & strSepChar & "NETClassRep" & strSepChar & "UMCMonoMW" & strSepChar & "UMCMWStDev" & strSepChar & "UMCMWMin" & strSepChar & "UMCMWMax" & strSepChar & "UMCAbundance" & strSepChar
    strLineOut = strLineOut & "ClassStatsChargeBasis" & strSepChar & "ChargeStateMin" & strSepChar & "ChargeStateMax" & strSepChar & "UMCMZForChargeBasis" & strSepChar & "UMCMemberCount" & strSepChar & "UMCMemberCountUsedForAbu" & strSepChar & "UMCAverageFit" & strSepChar & "PairIndex" & strSepChar
    strLineOut = strLineOut & "ExpressionRatio" & strSepChar & "ExpressionRatioStDev" & strSepChar & "ExpressionRatioChargeStateBasisCount" & strSepChar & "ExpressionRatioMemberBasisCount" & strSepChar
    strLineOut = strLineOut & "MultiMassTagHitCount" & strSepChar & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagNET" & strSepChar & "MassTagNETStDev" & strSepChar & "SLiC Score" & strSepChar & "DelSLiC" & strSepChar & "MemberCountMatchingMassTag" & strSepChar & "IsInternalStdMatch" & strSepChar & "PeptideProphetProbability"
    
    strExport(0) = strLineOut
    If Len(strFilePath) > 0 Then
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        Print #OutFileNum, strExport(0)
    End If
    
    lngExportCount = 1
    
    ' Step 4: Output the UMC's; exit the For loop if lngExportCount becomes greater than lngMaxPointsCountToCopy
    For lngUMCIndex = 0 To lngUMCsInViewCount - 1
    
        lngUMCIndexOriginal = udtUMCsInView(lngUMCIndex).UMCIndex
        
        With GelUMC(nMyIndex).UMCs(lngUMCIndexOriginal)
            Select Case .ClassRepType
            Case gldtCS
                lngScanClassRep = GelData(nMyIndex).CSData(.ClassRepInd).ScanNumber
            Case gldtIS
                lngScanClassRep = GelData(nMyIndex).IsoData(.ClassRepInd).ScanNumber
            Case Else
                Debug.Assert False
                lngScanClassRep = (.MinScan + .MaxScan) / 2
            End Select
            
            Debug.Assert ClsStat(lngUMCIndexOriginal, ustScanStart) = .MinScan
            Debug.Assert ClsStat(lngUMCIndexOriginal, ustScanEnd) = .MaxScan
        
            dblGANETClassRep = ScanToGANET(nMyIndex, lngScanClassRep)
            
            strLineOut = lngUMCIndexOriginal & strSepChar & .MinScan & strSepChar & .MaxScan & strSepChar & lngScanClassRep & strSepChar & Format(dblGANETClassRep, "0.0000") & strSepChar & Round(.ClassMW, 6) & strSepChar
            strLineOut = strLineOut & Round(.ClassMWStD, 6) & strSepChar & .MinMW & strSepChar & .MaxMW & strSepChar & .ClassAbundance & strSepChar
            
            strMinMaxCharges = ClsStat(lngUMCIndexOriginal, ustChargeMin) & strSepChar & ClsStat(lngUMCIndexOriginal, ustChargeMax) & strSepChar
            
            ' Record ClassStatsChargeBasis, ChargeMin, ChargeMax, UMCMZForChargeBasis, UMCMemberCount, and UMCMemberCountUsedForAbu
            If GelUMC(nMyIndex).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge & strSepChar
                strLineOut = strLineOut & strMinMaxCharges
                strLineOut = strLineOut & Round(MonoMassToMZ(.ClassMW, .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge), 6) & strSepChar
            Else
                strLineOut = strLineOut & 0 & strSepChar
                strLineOut = strLineOut & strMinMaxCharges
                strLineOut = strLineOut & Round(MonoMassToMZ(.ClassMW, CInt(GelData(nMyIndex).IsoData(.ClassRepInd).Charge)), 6) & strSepChar
            End If
        
            strLineOut = strLineOut & .ClassCount & strSepChar
            
            ' Record UMCMemberCountUsedForAbu
            If GelUMC(nMyIndex).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Count & strSepChar
            Else
                strLineOut = strLineOut & .ClassCount & strSepChar
            End If
        
        End With
        
        strLineOut = strLineOut & Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), 3) & strSepChar
        
        ' Now start populating strLineOutEnd
        strLineOutEnd = ""
        strLineOutEnd = strLineOutEnd & udtUMCsInView(lngUMCIndex).MultiAMTHitCount & strSepChar
        
        If mnuSCopyPointsInViewByUMCIncludeSearchResults.Checked And _
           blnIncludeAMTMass And udtUMCsInView(lngUMCIndex).IDIndex >= 0 And _
           Not udtUMCsInView(lngUMCIndex).IDIsInternalStd Then
            lngMassTagIndexPointer = BinarySearchLng(lngAMTID(), udtUMCsInView(lngUMCIndex).IDIndex, 0, AMTCnt - 1)
            
            If lngMassTagIndexPointer >= 0 Then
                lngMassTagIndexOriginal = lngAMTIDPointer(lngMassTagIndexPointer)
                dblAMTMW = AMTData(lngMassTagIndexOriginal).MW
                dblAMTNet = AMTData(lngMassTagIndexOriginal).NET
                dblAMTNetStDev = AMTData(lngMassTagIndexOriginal).NETStDev
                sngPeptideProphetProbability = AMTData(lngMassTagIndexOriginal).PeptideProphetProbability
                Debug.Assert AMTData(lngMassTagIndexOriginal).ID = udtUMCsInView(lngUMCIndex).IDIndex
            Else
                dblAMTMW = 0
                dblAMTNet = 0
                dblAMTNetStDev = 0
                sngPeptideProphetProbability = 0
            End If
        Else
            dblAMTMW = 0
            dblAMTNet = 0
            dblAMTNetStDev = 0
            sngPeptideProphetProbability = 0
        End If
        
        strLineOutEnd = strLineOutEnd & udtUMCsInView(lngUMCIndex).IDIndex & strSepChar & Round(dblAMTMW, 6) & strSepChar & Round(dblAMTNet, 4) & strSepChar & Round(dblAMTNetStDev, 4)
        strLineOutEnd = strLineOutEnd & strSepChar & udtUMCsInView(lngUMCIndex).SLiCScore
        strLineOutEnd = strLineOutEnd & strSepChar & udtUMCsInView(lngUMCIndex).DelSLiC
        strLineOutEnd = strLineOutEnd & strSepChar & udtUMCsInView(lngUMCIndex).MemberHitCount
        strLineOutEnd = strLineOutEnd & strSepChar & udtUMCsInView(lngUMCIndex).IDIsInternalStd
        strLineOutEnd = strLineOutEnd & strSepChar & Round(sngPeptideProphetProbability, 5)
        
        lngPairIndex = -1
        lngPairMatchCount = 0
        ReDim udtPairMatchStats(0)
        InitializePairMatchStats udtPairMatchStats(0)
        If blnPairsPresent Then
            lngPairIndex = PairIndexLookupSearch(nMyIndex, lngUMCIndexOriginal, objP1IndFastSearch, objP2IndFastSearch, True, False, lngPairMatchCount, udtPairMatchStats())
        End If
        
        If lngPairMatchCount > 0 Then
            For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                With udtPairMatchStats(lngPairMatchIndex)
                    strLineOutMiddle = Trim(.PairIndex) & strSepChar & Trim(.ExpressionRatio) & strSepChar & Trim(.ExpressionRatioStDev) & strSepChar & Trim(.ExpressionRatioChargeStateBasisCount) & strSepChar & Trim(.ExpressionRatioMemberBasisCount) & strSepChar
                    CopyAllUMCsInViewAppendLine strLineOut & strLineOutMiddle & strLineOutEnd, strExport(), strFilePath, OutFileNum, lngExportCount, lngExportCountDimmed
                End With
            Next lngPairMatchIndex
        Else
            ' No pair, and thus no expression ratio values
            strLineOutMiddle = Trim(-1) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar
            
            strLineOut = strLineOut & strLineOutMiddle & strLineOutEnd
            CopyAllUMCsInViewAppendLine strLineOut, strExport(), strFilePath, OutFileNum, lngExportCount, lngExportCountDimmed
        End If

        If lngMaxPointsCountToCopy > 0 Then
            If lngExportCount > lngMaxPointsCountToCopy Then Exit For
        End If

        If lngUMCIndex Mod 100 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngUMCIndex
            If KeyPressAbortProcess > 1 Then
                frmProgress.HideForm
                Exit Sub
            End If
        End If

    Next lngUMCIndex
    
    lngProgessStepCount = lngProgessStepCount + 1
    frmProgress.UpdateProgressBar lngProgessStepCount
    frmProgress.InitializeSubtask "Copying to clipboard", 0, 1
    
    If Len(strFilePath) > 0 Then
        ' Close the file
        Close OutFileNum
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            If KeyPressAbortProcess > 1 Then
                MsgBox "Copy data cancelled.", vbInformation + vbOKOnly, glFGTU
            Else
                MsgBox "Data saved to file: " & strFilePath, vbInformation + vbOKOnly, glFGTU
            End If
        End If
    Else
        If KeyPressAbortProcess > 1 Then
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "Copy data cancelled.", vbInformation + vbOKOnly, glFGTU
            End If
        Else
            ' Copy to the clipboard
            strFlattenedList = FlattenStringArray(strExport(), lngExportCount, vbCrLf)
            
            On Error Resume Next
            Clipboard.Clear
            Clipboard.SetText strFlattenedList, vbCFText
        End If
    End If
    
    frmProgress.HideForm
    Exit Sub

CopyAllUMCsInViewErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error occurred in CopyAllUMCsInView:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
    LogErrors Err.Number, "CopyAllUMCsInView"
    End If
    frmProgress.HideForm
End Sub

Private Sub CopyAllUMCsInViewAppendLine(ByVal strNewLine As String, ByRef strExport() As String, ByVal strFilePath As String, ByVal OutFileNum As Integer, ByRef lngExportCount As Long, ByRef lngExportCountDimmed As Long)
    
    strExport(lngExportCount) = strNewLine
    If Len(strFilePath) > 0 Then Print #OutFileNum, strExport(lngExportCount)
    lngExportCount = lngExportCount + 1
    
    If lngExportCount >= lngExportCountDimmed Then
        lngExportCountDimmed = lngExportCountDimmed + 500
        ReDim Preserve strExport(lngExportCountDimmed)
    End If

End Sub

Public Sub NormalViewMenu()
mnu2lsFilter.Enabled = True
mnu2lsShowAll.Enabled = True
mnuViewChargeState.Enabled = True
mnuViewIsotopic.Enabled = True
mnuViewZOrder(0).Enabled = True
mnuViewZOrder(1).Enabled = True
End Sub

' MonroeMod: New Function
Public Sub RequestRefreshPlot()
    bNeedToUpdate = True
End Sub

Public Sub ResetGraph(Optional blnReapplyFilters As Boolean = True, Optional blnShowFilterForm As Boolean = True, Optional intViewToShow As Integer = glNormalDisplay)
    lAction = glNoAction
    If BuildGraph Then
        If blnReapplyFilters Then
            If blnShowFilterForm Then
                ShowFilterForm
            Else
                frmFilter.Tag = nMyIndex
                frmFilter.InitializeControls True
            End If
        End If
    End If

    Select Case intViewToShow
    Case glDifferentialDisplay
        ShowDiffDisplay
    Case glChargeStateMapDisplay
        ShowChargeStateMap
    Case Else
        ' Includes glNormalDisplay
        ShowNormalView
    End Select
    
    
End Sub

Private Function SaveFileAsInit(eFileSaveMode As fsFileSaveModeConstants, Optional blnPromptForPath As Boolean = True)
    Dim sSaveFileName As String
    Dim eFileSaveModeSaved As fsFileSaveModeConstants
    Dim blnSuccess As Boolean
    
    If (Left(Me.Caption, 8) = "Untitled") Or Len(Me.Caption) = 0 Then
        sSaveFileName = ""
    Else
        sSaveFileName = Me.Caption
    End If
    
    eFileSaveModeSaved = mFileSaveMode
    mFileSaveMode = eFileSaveMode
    
    Select Case mFileSaveMode
    Case fsNoExtended
        blnSuccess = SaveFileAsWrapper(sSaveFileName, blnPromptForPath, False, True)
    Case fsIncludeExtended
        blnSuccess = SaveFileAsWrapper(sSaveFileName, blnPromptForPath, False, False, True)
    Case fsLegacy
        blnSuccess = SaveFileAsWrapper(sSaveFileName, blnPromptForPath, False, False, False, True)
    Case Else
        ' Includes fsUnknown
        blnSuccess = SaveFileAsWrapper(sSaveFileName, blnPromptForPath, True, False, glbPreferencesExpanded.ExtendedFileSaveModePreferred)
    End Select
    
    If Not blnSuccess Then mFileSaveMode = eFileSaveModeSaved
    SaveFileAsInit = blnSuccess
    
End Function

' MonroeMod: New function; used to call SaveFileAs()
Private Function SaveFileAsWrapper(strSuggestedSaveFilePath As String, blnPromptForPath As Boolean, Optional blnPromptToIncludeExtendedInfo As Boolean = False, Optional blnForceNoExtendedInfo As Boolean = False, Optional blnDefaultToIncludeExtendedInfo As Boolean = False, Optional blnUseLegacyFileFormat As Boolean = False) As Boolean
    ' Returns True if successful save, or False if cancelled
    
    Dim sSaveFileName As String
    Dim SaveType As fstFileSaveTypeConstants
    Dim eResponse As VbMsgBoxResult
    Dim blnIncludeUMCData As Boolean, blnIncludePairsData As Boolean
    Dim intDefaultButtonIndex As Integer
    
    If Me.Tag < 0 Then
       SaveType = fstFileSaveTypeConstants.fstUMR
    Else
       SaveType = fstFileSaveTypeConstants.fstGel
    End If
    
    If blnPromptToIncludeExtendedInfo Then
        If blnDefaultToIncludeExtendedInfo Then
            intDefaultButtonIndex = vbDefaultButton1
        Else
            intDefaultButtonIndex = vbDefaultButton2
        End If
        
          ' No longer supported (March 2006)
''        If GelORFData(nMyIndex).ORFCount > 0 Then
''            eResponse = MsgBox("Include ORF (protein) Data and ORF MT tags when saving?", vbQuestion + vbYesNoCancel + intDefaultButtonIndex, "Include ORF Data")
''            If eResponse = vbCancel Then
''                SaveFileAsWrapper = False
''                Exit Function
''            Else
''                blnIncludeORFData = (eResponse = vbYes)
''            End If
''        Else
''            blnIncludeORFData = True
''        End If
        
        If GelUMC(nMyIndex).UMCCnt > 0 Then
            eResponse = MsgBox("Include UMC data when saving?", vbQuestion + vbYesNoCancel + intDefaultButtonIndex, "Include UMC Data")
            If eResponse = vbCancel Then
                SaveFileAsWrapper = False
                Exit Function
            Else
                blnIncludeUMCData = (eResponse = vbYes)
            End If
        Else
            blnIncludeUMCData = True
        End If
        
        ''If blnIncludeUMCData And (GelP(nMyIndex).PCnt > 0 Or GelP_D_L(nMyIndex).PCnt > 0) Then
        If blnIncludeUMCData And (GelP_D_L(nMyIndex).PCnt > 0) Then
            eResponse = MsgBox("Include Pairs data when saving?", vbQuestion + vbYesNoCancel + intDefaultButtonIndex, "Include Pairs Data")
            If eResponse = vbCancel Then
                SaveFileAsWrapper = False
                Exit Function
            Else
                blnIncludePairsData = (eResponse = vbYes)
            End If
        Else
            blnIncludePairsData = blnIncludeUMCData
        End If
    Else
        blnIncludeUMCData = Not blnForceNoExtendedInfo
        blnIncludePairsData = Not blnForceNoExtendedInfo
    End If
    
    If blnIncludeUMCData And blnIncludePairsData Then
        mFileSaveMode = fsIncludeExtended
    ElseIf Not (blnIncludeUMCData Or blnIncludePairsData) Then
        mFileSaveMode = fsNoExtended
    Else
        mFileSaveMode = fsUnknown
    End If
    
    ' No longer supported (March 2006)
    ''If GelORFData(nMyIndex).ORFCount = 0 And GelUMC(nMyIndex).UMCCnt = 0 Then mFileSaveMode = fsUnknown
    
    If blnUseLegacyFileFormat Then mFileSaveMode = fsLegacy
    
    If strSuggestedSaveFilePath = "" Then
       sSaveFileName = FileSaveProc(Me.hwnd, SuggestionByIndex(nMyIndex, "gel"), SaveType)
    Else
        If blnPromptForPath Then
            sSaveFileName = FileSaveProc(Me.hwnd, StripFullPath(Me.Caption), SaveType)
        Else
            sSaveFileName = strSuggestedSaveFilePath
        End If
    End If
    
    If sSaveFileName <> "" Then
        SaveFileAs sSaveFileName, Not blnIncludeUMCData, Not blnIncludePairsData, nMyIndex, mFileSaveMode
        SaveFileAsWrapper = True
    Else
        SaveFileAsWrapper = False
    End If

End Function

Private Sub SavePEKFileUsingDataPoints(ByVal blnLimitToDataInView As Boolean)
    Dim blnSuccess As Boolean
    Me.MousePointer = vbHourglass
    MDIForm1.lblStatus.Caption = "Constructing PEK file."
    DoEvents
    blnSuccess = GeneratePEKFileUsingDataPoints(nMyIndex, blnLimitToDataInView, "", Me.hwnd)
    MDIForm1.lblStatus.Caption = ""
    Me.MousePointer = vbDefault
End Sub

Private Sub SavePEKFileUsingUMCs(ByVal blnLimitToUMCsInView As Boolean)
    Dim blnSuccess As Boolean
    Me.MousePointer = vbHourglass
    MDIForm1.lblStatus.Caption = "Constructing PEK file."
    DoEvents
    blnSuccess = GeneratePEKFileUsingUMCs(nMyIndex, blnLimitToUMCsInView, "", Me.hwnd)
    MDIForm1.lblStatus.Caption = ""
    Me.MousePointer = vbDefault
End Sub

Private Sub SavePEKFileUsingOriginalPEKFileAsTemplate()
    ' Note that this option requires that the original PEK file be available
    
    Dim blnSuccess As Boolean
    
    Me.MousePointer = vbHourglass
    MDIForm1.lblStatus.Caption = "Constructing PEK file."
    DoEvents
    blnSuccess = WriteGELAsPEK(nMyIndex, Me.hwnd)
    MDIForm1.lblStatus.Caption = ""
    Me.MousePointer = vbDefault
End Sub

Private Sub picGraph_Paint()
On Error Resume Next
If IsWinLoaded(TrackerCaption) Then frmTracker.Visible = True
If Not Err Then
   ' MonroeMod: Sending request to paint graph to RequestRefreshPlot
   If (Not bLoad) Then RequestRefreshPlot
End If
End Sub

Private Sub picGraph_Resize()
If bPaint Then picGraph.Refresh
End Sub

Public Sub BuildCooSys()
With csMyCooSys
     .BuildingCS = True
     .CSIndex = nMyIndex
     .csOriginXY = GelData(nMyIndex).Preferences.CooOrigin
     .csXOrient = GelData(nMyIndex).Preferences.CooHOrientation
     .csYOrient = GelData(nMyIndex).Preferences.CooVOrientation
     .csType = GelData(nMyIndex).Preferences.CooType
     .csYScale = GelData(nMyIndex).Preferences.CooVAxisScale
     .BuildingCS = False
End With
End Sub

Private Sub DevLogConversion(ByVal ConversionType As Integer, ByVal NumOfPoints As Integer)
Dim OldDC As Long, Res As Long

OldDC = SaveDC(picGraph.hDC)
GelCooSys nMyIndex, picGraph.hDC
Select Case ConversionType
Case ConvDPLP
    Res = DPtoLP(picGraph.hDC, paPoints(0), NumOfPoints)
Case ConvLPDP
    Res = LPtoDP(picGraph.hDC, paPoints(0), NumOfPoints)
End Select
Res = RestoreDC(picGraph.hDC, OldDC)
End Sub

Private Sub SelToolBox1_ClickAvg()
    HandleSelToolboxClick ssrfAverage
End Sub

Private Sub SelToolBox1_ClickMax()
    HandleSelToolboxClick ssrfMaximum
End Sub

Private Sub SelToolBox1_ClickMin()
    HandleSelToolboxClick ssrfMinimum
End Sub

Private Sub SelToolBox1_ClickRange()
    HandleSelToolboxClick ssrfRange
End Sub

Private Sub SelToolBox1_ClickStD()
    HandleSelToolboxClick ssrfStDev
End Sub

Private Sub SelToolBox1_DblClick(ByVal DblClickType As DblClickPosition)
Dim BckClr As Long
Select Case DblClickType
Case DblClickTL
     Call mnuViewSelToolBox_Click
Case DblClickBR
     BckClr = SelToolBox1.BackColor
     Call GetColorAPIDlg(Me.hwnd, BckClr)
     If BckClr >= 0 Then SelToolBox1.BackColor = BckClr
End Select
End Sub

Private Sub SelToolBox1_MouseDown()
SelToolBox1.Drag vbBeginDrag
End Sub

Private Sub SelToolBox1_MouseUp()
SelToolBox1.Drag vbEndDrag
End Sub

Public Sub ShowChargeStateMap()
NormalViewMenu
mnuViewNormalDisplay.Checked = False
mnuViewDiffDisplay.Checked = False
mnuViewCSMap.Checked = True
fgDisplay = glChargeStateMapDisplay
InitDrawChargeStateMap nMyIndex
picGraph.Refresh
End Sub

Public Sub ShowDiffDisplay()
NormalViewMenu
mnuViewNormalDisplay.Checked = False
mnuViewDiffDisplay.Checked = True
mnuViewCSMap.Checked = False
fgDisplay = glDifferentialDisplay
InitDrawER nMyIndex
picGraph.Refresh
End Sub

Public Sub ShowNormalView()
NormalViewMenu
mnuViewNormalDisplay.Checked = True
mnuViewDiffDisplay.Checked = False
mnuViewCSMap.Checked = False
fgDisplay = glNormalDisplay
picGraph.Refresh
End Sub

Private Sub tmrRefreshTimer_Timer()
    If bNeedToUpdate Then
        If APIDrawingAborted Then
            If DateDiff("s", APIDrawStartTime, Now()) < 2 Then
                bNeedToUpdate = False
                Exit Sub
            Else
                APIDrawingAborted = False
            End If
        End If
        GelDrawScreen nMyIndex
        bNeedToUpdate = False
    End If
End Sub

Private Sub UpdateMenuMode(ByVal eMenuMode As mmMenuModeConstants)
    Dim intIndex As Integer
    
On Error GoTo UpdateMenuModeErrorHandler

    If eMenuMode = mmBasic Then mnuVMenuModeIncludeObsolete.Checked = False
    
    For intIndex = 0 To MENU_MODE_COUNT - 1
        mnuVMenuModeSelect(intIndex).Checked = (intIndex = eMenuMode)
    Next intIndex
    
    UpdateVisibleMenus
    Exit Sub

UpdateMenuModeErrorHandler:
    MsgBox "Error in UpdateMenuMode (" & Err.Number & "):" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    
End Sub

Public Sub UpdateTICPlotAndFeatureBrowsersIfNeeded(Optional blnForceFeatureBrowserUpdate As Boolean = False)
    If IsLoaded("frmTICAndBPIPlots") Then
        frmTICAndBPIPlots.CallerID = nMyIndex
        frmTICAndBPIPlots.AutoUpdatePlot
    End If
    
    If IsLoaded("frmPairBrowser") Then
        frmPairBrowser.CallerIDNew = nMyIndex
        frmPairBrowser.AutoUpdatePlot blnForceFeatureBrowserUpdate
    End If
    
    If IsLoaded("frmUMCBrowser") Then
        frmUMCBrowser.CallerIDNew = nMyIndex
        frmUMCBrowser.AutoUpdatePlot blnForceFeatureBrowserUpdate
    End If
    
End Sub

Private Sub UpdateVisibleMenus()
    Dim intIndex As Integer
    Dim eMenuMode As mmMenuModeConstants
    Dim blnBasic As Boolean, blnDBSimple As Boolean, blnDBPairs As Boolean, blnFull As Boolean
    
On Error GoTo UpdateVisibleMenusErrorHandler

    glbPreferencesExpanded.MenuModeIncludeObsolete = mnuVMenuModeIncludeObsolete.Checked
    
    If APP_BUILD_DISABLE_ADVANCED Then
        glbPreferencesExpanded.MenuModeIncludeObsolete = False
        mnuVMenuModeIncludeObsolete.Visible = False
        mnuVMenuModeSep.Visible = False
    End If
    
    For intIndex = 0 To MENU_MODE_COUNT - 1
        If mnuVMenuModeSelect(intIndex).Checked Then
            eMenuMode = intIndex
            Exit For
        End If
    Next intIndex
    
    glbPreferencesExpanded.MenuModeDefault = eMenuMode
    
    Select Case eMenuMode
    Case mmBasic
        blnBasic = True
    Case mmDBNoPairs
        blnDBSimple = True
    Case mmDBWithPairs
        blnDBSimple = True
        blnDBPairs = True
    Case mmFull
        blnDBSimple = True
        blnDBPairs = True
        blnFull = True
    Case Else
        Debug.Assert False
        eMenuMode = mmFull
    End Select
    
    mnuNewAutoAnalysis.Visible = Not blnBasic
    mnuSaveAsCompressed.Visible = Not blnBasic
    mnuEditUMC.Visible = Not blnBasic
    mnuEditAvgUMC.Visible = Not blnBasic
    mnuEditAbundance.Visible = Not blnBasic
    mnuEditNETFormula.Visible = Not blnBasic
    
    If blnBasic Then
        mnuEditNETAdjustment(enaEditNETAdjustmentMenuConstants.enaIndividualPoints).Visible = False
        mnuEditNETAdjustment(enaEditNETAdjustmentMenuConstants.enaUMCIterative).Visible = APP_BUILD_DISABLE_ADVANCED
        mnuEditNETAdjustment(enaEditNETAdjustmentMenuConstants.enaMSAlign).Visible = Not APP_BUILD_DISABLE_ADVANCED
    
        mnuStepsNETAdjustmentMode(namStepsNETAdjustmentModeConstants.namUMCIterative).Visible = APP_BUILD_DISABLE_ADVANCED
        mnuStepsNETAdjustmentMode(namStepsNETAdjustmentModeConstants.namMSAlign).Visible = Not APP_BUILD_DISABLE_ADVANCED
    Else
        mnuEditNETAdjustment(enaEditNETAdjustmentMenuConstants.enaIndividualPoints).Visible = glbPreferencesExpanded.MenuModeIncludeObsolete
        mnuEditNETAdjustment(enaEditNETAdjustmentMenuConstants.enaUMCIterative).Visible = True
        mnuEditNETAdjustment(enaEditNETAdjustmentMenuConstants.enaMSAlign).Visible = Not APP_BUILD_DISABLE_ADVANCED
    
        mnuStepsNETAdjustmentMode(namStepsNETAdjustmentModeConstants.namUMCIterative).Visible = True
        mnuStepsNETAdjustmentMode(namStepsNETAdjustmentModeConstants.namMSAlign).Visible = Not APP_BUILD_DISABLE_ADVANCED
    End If
   
    mnuEditClear.Visible = Not blnBasic
    mnuEditParameters.Visible = Not blnBasic
    mnu2lsShowSpectrum.Visible = Not blnBasic
    mnu2lsShowSpectrumNearestCursorPoint.Visible = Not blnBasic
    mnu2lsCloseAllICR2LSMassSpectra.Visible = Not blnBasic
    mnu2lsSepShowSpectra.Visible = Not blnBasic
    
    mnu2lsToleranceRefinement.Visible = Not blnBasic
    
    mnuViewSepOverlay.Visible = Not APP_BUILD_DISABLE_ADVANCED
    mnuVOverlay.Visible = mnuViewSepOverlay.Visible
    mnu2lsOverlaysManager.Visible = mnuViewSepOverlay.Visible
 
    mnuEditSep2.Visible = Not blnBasic
    mnuEditSep3.Visible = Not blnBasic
    mnu2lsSep8.Visible = Not blnBasic
    mnu2lsSep9.Visible = Not blnBasic
    
    mnuSplitUMCs.Visible = Not blnBasic
    mnuShowSplitUMCs.Visible = Not blnBasic
    mnuSSepShow.Visible = Not blnBasic

    mnuShowNETAdjUMCs.Visible = Not blnBasic
    mnuShowLowSegmentCountUMCs.Visible = Not blnBasic
    mnuShowNETAdjUMCsWithDBHit.Visible = Not blnBasic
    
    mnuViewZoomRegionListEditor.Visible = Not blnBasic
    
    mnuSSep1.Visible = glbPreferencesExpanded.MenuModeIncludeObsolete
    mnuSSep3.Visible = blnFull Or glbPreferencesExpanded.MenuModeIncludeObsolete
    
    mnuEditDatabaseConnection.Visible = blnDBSimple
    mnuEditMTStatus.Visible = blnDBSimple
    mnuVMTDisplay.Visible = blnDBSimple
    mnu2lsSearchAMT.Visible = blnDBSimple
    mnu2lsSearchUMCSingleMass.Visible = blnDBSimple
    
    mnuSSep7.Visible = blnDBPairs
    
    mnuEditClearER.Visible = blnDBPairs
    mnuEditClearIDBadDelta.Visible = blnDBPairs
    mnuEditClearIDBadDeltaMT.Visible = blnDBPairs
    mnuViewPairsBrowser.Visible = blnDBPairs
    mnu2lsSepSearchPairs.Visible = blnDBPairs
    mnu2lsSearchUMCPairs.Visible = blnDBPairs
    mnu2lsSearchUMCPairs_ICAT.Visible = blnDBPairs
    mnu2lsSearchUMCPairs_PEON14N15.Visible = blnDBPairs
    mnuSDltLbl.Visible = blnDBPairs
    mnuSAttentionList.Visible = blnDBPairs
    
    mnuEditDiscreteMWs.Visible = blnFull
    mnuEditCalibration.Visible = blnFull
    mnuEditCopyEMF.Visible = blnFull
    mnuEditAdjScans.Visible = blnFull
    mnuEditResidualDisplay.Visible = blnFull
    mnuEditMTStatus.Visible = blnFull

    mnuIMTSettings.Visible = Not blnBasic
    mnuIAnalysisInfo.Visible = Not blnBasic
    mnuInfoSep1.Visible = blnFull
    mnuViewDiffDisplay.Visible = Not blnBasic
    mnuViewRawData.Visible = blnFull
    mnuIDeviceCaps.Visible = blnFull

    'mnuViewCorrelations.Visible = blnFull
    mnuSLoadScope.Visible = blnFull
    
    ' Obsolete menus
    ' Note: If APP_BUILD_DISABLE_ADVANCED = True, then .MenuModeIncludeObsolete will be False
    With glbPreferencesExpanded
        mnu2lsORFCenteredSearchPRISM.Visible = .MenuModeIncludeObsolete
        mnu2lsORFCenteredSearchFASTA.Visible = .MenuModeIncludeObsolete
        mnu2lsSearchAMTOld.Visible = .MenuModeIncludeObsolete
        mnu2lsSearchAMTNew.Visible = .MenuModeIncludeObsolete
        mnu2lsSearchUMCMassTags.Visible = .MenuModeIncludeObsolete
        mnu2LSSearchForORFs.Visible = .MenuModeIncludeObsolete
        mnu2lsShowData.Visible = .MenuModeIncludeObsolete
        mnu2lsSep3.Visible = .MenuModeIncludeObsolete
        mnuSMS_MSSearch.Visible = .MenuModeIncludeObsolete
        mnuSUMCLockMass.Visible = .MenuModeIncludeObsolete
        mnuSLockMass.Visible = .MenuModeIncludeObsolete
        mnuSIntCalLockMass.Visible = .MenuModeIncludeObsolete
        mnuSExcludeExc.Visible = .MenuModeIncludeObsolete
        mnuEditFS.Visible = .MenuModeIncludeObsolete
    
' No longer supported (March 2006)
''        mnuSExport.Visible = .MenuModeIncludeObsolete
''        mnu2lsShowAMTRecord.Visible = .MenuModeIncludeObsolete
    
    End With
    
    Exit Sub

UpdateVisibleMenusErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub SynchronizeCopyDataOptions()
    Dim lngGelIndex As Long
    
    ' Make sure the menu option is the same for the other gels
    For lngGelIndex = 1 To UBound(GelBody())
        GelBody(lngGelIndex).mnuSCopyScansIncludeEmptyScans.Checked = mnuSCopyScansIncludeEmptyScans.Checked
        GelBody(lngGelIndex).mnuSCopyPointsInViewOneHitPerLine.Checked = mnuSCopyPointsInViewOneHitPerLine.Checked
    Next lngGelIndex

End Sub

Private Sub ZoomActionCancel()
    ' Cancel the zoom
    lAction = glNoAction
    ZoomActionEnd
    bNeedToUpdate = True
End Sub

Private Sub ZoomActionEnd()
    Dim Res As Long
    ReDim paPoints(1)
    Select Case lAction
    Case glNoAction   'remove restrictions on cursor
        Res = ClipCursorByNum(&O0)
    Case glActionZoom
        Res = ClipCursorByNum(&O0)
        picGraph.Line (gbZoomX1, gbZoomY1)-(gbZoomX2, gbZoomY2), , B
        picGraph.DrawStyle = vbSolid
        If (Abs(gbZoomX1 - gbZoomX2) > 10) And (Abs(gbZoomY1 - gbZoomY2) > 10) Then
           paPoints(0).x = gbZoomX1
           paPoints(0).y = gbZoomY1
           paPoints(1).x = gbZoomX2
           paPoints(1).y = gbZoomY2
           DevLogConversion ConvDPLP, 2
           csMyCooSys.ZoomIn paPoints(0).x, paPoints(0).y, paPoints(1).x, paPoints(1).y
        End If
    End Select
    lAction = glNoAction
    
    If mnuSCopyPointsInViewToClipboardAuto.Checked Then CopyAllPointsInView mAutoCopyPointsMaxCount
End Sub

Private Sub ZoomActionStart(x As Single, y As Single, paPoints() As POINTAPI)
  Dim Res As Long
  Dim rcClip As Rect
  Dim paClip As POINTAPI

  If lAction = glActionZoom Then
     gbZoomX1 = x
     gbZoomY1 = y
     gbZoomX2 = gbZoomX1
     gbZoomY2 = gbZoomY1
'clip cursor on viewport
     paPoints(0).x = gbZoomX1
     paPoints(0).y = gbZoomY1
     DevLogConversion ConvDPLP, 1
     If (paPoints(0).x < LDfX0) Or (paPoints(0).x > LDfXE) Or _
        (paPoints(0).y < LDfY0) Or (paPoints(0).y > LDfYE) Then
        lAction = glNoAction
        Exit Sub
     End If
     paClip.x = 0
     paClip.y = 0
     Res = ClientToScreen(picGraph.hwnd, paClip)
     csMyCooSys.GetViewPortRectangle paClip.x, paClip.y, rcClip.Top, rcClip.Left, rcClip.Bottom, rcClip.Right
     Res = ClipCursor(rcClip)
     picGraph.DrawStyle = vbDot
  End If

End Sub

VERSION 5.00
Begin VB.Form frmFileLoadOptions 
   Caption         =   "File Load Options"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   7470
   Begin VB.Frame fraFilterLCMSFeatures 
      Caption         =   "LC-MS Feature Filters"
      Height          =   1455
      Left            =   3600
      TabIndex        =   39
      Top             =   3960
      Width           =   3615
      Begin VB.TextBox txtLCMSFeatureScanCountMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   45
         Text            =   "2"
         Top             =   640
         Width           =   975
      End
      Begin VB.TextBox txtLCMSFeatureAbuMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   43
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtIMSConformerScoreMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   47
         Text            =   "0"
         Top             =   1035
         Width           =   975
      End
      Begin VB.OptionButton optFilterLCMSFeatures 
         Caption         =   "On"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optFilterLCMSFeatures 
         Caption         =   "Off"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblLCMSFeatureScanCountMin 
         Caption         =   "Min LC Scans"
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Min Abundance"
         Height          =   255
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblIMSConformerScoreMin 
         Caption         =   "Min Conformer Score"
         Height          =   375
         Left            =   960
         TabIndex        =   46
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame fraPredefinedLCMSFeatureOptions 
      Caption         =   "Predefined LC-MS Feature Options"
      Height          =   2655
      Left            =   120
      TabIndex        =   32
      Top             =   4560
      Width           =   3375
      Begin VB.ComboBox cboLCMSFeatureClassAbundanceMode 
         Height          =   315
         ItemData        =   "frmFileLoadOptions.frx":0000
         Left            =   120
         List            =   "frmFileLoadOptions.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   2160
         Width           =   3000
      End
      Begin VB.CheckBox chkSplitUMCsByExaminingAbundance 
         Caption         =   "Split LC-MS features after loading"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1600
         Width           =   3015
      End
      Begin VB.ComboBox cboLCMSFeaturePointsLoadMode 
         Height          =   315
         ItemData        =   "frmFileLoadOptions.frx":0004
         Left            =   120
         List            =   "frmFileLoadOptions.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1200
         Width           =   3000
      End
      Begin VB.TextBox txtAutoMapDataPointsMassTolerancePPM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   34
         Text            =   "5"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblDescription 
         Caption         =   "Class Abundance Mode"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   54
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblDescription 
         Caption         =   "Choose which points to load:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblDescription 
         Caption         =   "ppm"
         Height          =   240
         Index           =   2
         Left            =   2760
         TabIndex        =   35
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblAutoMapDataPointsMassTolerancePPM 
         Caption         =   "Mass Tolerance for auto-mapping data points to predefined features"
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraDREAMS 
      Caption         =   "DREAMS Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   3375
      Begin VB.OptionButton optEvenOddScanFilter 
         Caption         =   "Only load odd-numbered scans"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   540
         Width           =   2800
      End
      Begin VB.OptionButton optEvenOddScanFilter 
         Caption         =   "Load all scans"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   2800
      End
      Begin VB.OptionButton optEvenOddScanFilter 
         Caption         =   "Only load even-numbered scans"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   2800
      End
   End
   Begin VB.CommandButton cmdSetToDefaults 
      Caption         =   "Set to &Defaults"
      Height          =   375
      Left            =   5640
      TabIndex        =   51
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame fraOptionFrame 
      Caption         =   "Data Count / Data Intensity Percentage Filter"
      Height          =   1695
      Index           =   22
      Left            =   3600
      TabIndex        =   19
      Top             =   2160
      Width           =   3615
      Begin VB.CheckBox chkTotalIntensityPercentageFilterEnabled 
         Caption         =   "Enable total intensity percentage filter"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtTotalIntensityPercentageFilter 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         Text            =   "90"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtMaximumDataCountToLoad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   "125000"
         ToolTipText     =   "Higher abundance data is favored when determine the data to load"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkMaximumDataCountEnabled 
         Caption         =   "Enable maximum data count filter"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblDescription 
         Caption         =   "%"
         Height          =   240
         Index           =   1
         Left            =   2760
         TabIndex        =   27
         Top             =   1360
         Width           =   735
      End
      Begin VB.Label lblDescription 
         Caption         =   "Cumulative % to retain"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   1360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Note: the data file must be pre-scanned"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblDescription 
         Caption         =   "Maximum points to load:"
         Height          =   255
         Index           =   137
         Left            =   240
         TabIndex        =   22
         Top             =   755
         Width           =   1935
      End
   End
   Begin VB.Frame fraMSLevelFilter 
      Caption         =   "MS Level Filter"
      Height          =   1215
      Left            =   3600
      TabIndex        =   48
      Top             =   5565
      Width           =   1455
      Begin VB.ListBox lstMSLevelFilter 
         Height          =   840
         ItemData        =   "frmFileLoadOptions.frx":0008
         Left            =   120
         List            =   "frmFileLoadOptions.frx":000A
         MultiSelect     =   2  'Extended
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraCSandIsoDataFilter 
      Caption         =   "Data Type Filter"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   3375
      Begin VB.OptionButton optCSandIsoDataFilterMode 
         Caption         =   "Only load Isotopic data"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   3000
      End
      Begin VB.OptionButton optCSandIsoDataFilterMode 
         Caption         =   "Only load Charge State data"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   3000
      End
      Begin VB.OptionButton optCSandIsoDataFilterMode 
         Caption         =   "Load all data from the input file"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   52
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   5640
      TabIndex        =   50
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Frame fraIsoFitFilter 
      Caption         =   "Isotopic Fit Filter"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
      Begin VB.OptionButton optIsoFitFilter 
         Caption         =   "Off"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optIsoFitFilter 
         Caption         =   "On"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtIsoFitMaxValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "0.15"
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblMaxFit 
         Caption         =   "Max"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraAbuFilter 
      Caption         =   "Abundance Filter"
      Height          =   975
      Left            =   3600
      TabIndex        =   8
      Top             =   1080
      Width           =   3615
      Begin VB.TextBox txtAbuFilterMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Text            =   "0"
         Top             =   560
         Width           =   1095
      End
      Begin VB.TextBox txtAbuFilterMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "1E+15"
         Top             =   200
         Width           =   1095
      End
      Begin VB.OptionButton optAbuFilter 
         Caption         =   "On"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAbuFilter 
         Caption         =   "Off"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblMinAbu 
         Caption         =   "Min"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   590
         Width           =   495
      End
      Begin VB.Label lblMaxAbu 
         Caption         =   "Max"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   230
         Width           =   495
      End
   End
   Begin VB.Label lblFileSize 
      Caption         =   "0 MB"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFilePathCaption 
      Caption         =   "Input File info:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFilePath 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   430
      Width           =   7095
   End
End
Attribute VB_Name = "frmFileLoadOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FILE_SIZE_THRESHOLD_BYTES As Long = 5242880

Private mFileSize As Long
Private mFileType As ifmInputFileModeConstants

Private mLoadCancelled As Boolean

Public Property Let AutoMapDataPointsMassTolerancePPM(ByVal sngValue As Single)
    txtAutoMapDataPointsMassTolerancePPM = sngValue
End Property
Public Property Get AutoMapDataPointsMassTolerancePPM() As Single
    If IsNumeric(txtAutoMapDataPointsMassTolerancePPM) Then
        AutoMapDataPointsMassTolerancePPM = txtAutoMapDataPointsMassTolerancePPM
    Else
        AutoMapDataPointsMassTolerancePPM = 5
    End If
End Property

Public Property Let FilterLCMSFeatures(ByVal blnEnable As Boolean)
    If blnEnable Then
        optFilterLCMSFeatures(1).Value = True
    Else
        optFilterLCMSFeatures(0).Value = True
    End If
End Property
Public Property Get FilterLCMSFeatures() As Boolean
    FilterLCMSFeatures = optFilterLCMSFeatures(1).Value
End Property

Public Property Let FilterOnIsoFit(ByVal blnEnable As Boolean)
    If blnEnable Then
        optIsoFitFilter(1).Value = True
    Else
        optIsoFitFilter(0).Value = True
    End If
End Property
Public Property Get FilterOnIsoFit() As Boolean
    FilterOnIsoFit = optIsoFitFilter(1).Value
End Property

Public Property Let FilterOnAbundance(ByVal blnEnable As Boolean)
    If blnEnable Then
        optAbuFilter(1).Value = True
    Else
        optAbuFilter(0).Value = True
    End If
End Property
Public Property Get FilterOnAbundance() As Boolean
    FilterOnAbundance = optAbuFilter(1).Value
End Property

Public Property Let DataFilterMode(ByVal eDataFilterMode As dfmCSandIsoDataFilterModeConstants)
    Select Case eDataFilterMode
    Case dfmCSandIsoDataFilterModeConstants.dfmLoadCSDataOnly
        optCSandIsoDataFilterMode(dfmCSandIsoDataFilterModeConstants.dfmLoadCSDataOnly).Value = True
    Case dfmCSandIsoDataFilterModeConstants.dfmLoadIsoDataOnly
        optCSandIsoDataFilterMode(dfmCSandIsoDataFilterModeConstants.dfmLoadIsoDataOnly).Value = True
    Case Else
        ' Includes dfmCSandIsoDataFilterModeConstants.dfmLoadAllData
        optCSandIsoDataFilterMode(dfmCSandIsoDataFilterModeConstants.dfmLoadAllData).Value = True
    End Select
End Property
Public Property Get DataFilterMode() As dfmCSandIsoDataFilterModeConstants
    If optCSandIsoDataFilterMode(dfmCSandIsoDataFilterModeConstants.dfmLoadCSDataOnly).Value = True Then
        DataFilterMode = dfmCSandIsoDataFilterModeConstants.dfmLoadCSDataOnly
    ElseIf optCSandIsoDataFilterMode(dfmCSandIsoDataFilterModeConstants.dfmLoadIsoDataOnly).Value = True Then
        DataFilterMode = dfmCSandIsoDataFilterModeConstants.dfmLoadIsoDataOnly
    Else
        DataFilterMode = dfmCSandIsoDataFilterModeConstants.dfmLoadAllData
    End If
End Property

Public Property Let EvenOddScanFilterMode(ByVal eEvenOddScanFilterMode As eosEvenOddScanFilterModeConstants)
    Select Case eEvenOddScanFilterMode
    Case eosEvenOddScanFilterModeConstants.eosLoadOddScansOnly
        optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadOddScansOnly).Value = True
    Case eosEvenOddScanFilterModeConstants.eosLoadEvenScansOnly
        optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadEvenScansOnly).Value = True
    Case Else
        ' Includes eosEvenOddScanFilterModeConstants.eosLoadAllScans
        optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadAllScans).Value = True
    End Select
End Property
Public Property Get EvenOddScanFilterMode() As eosEvenOddScanFilterModeConstants
    If optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadOddScansOnly).Value = True Then
        EvenOddScanFilterMode = eosEvenOddScanFilterModeConstants.eosLoadOddScansOnly
    ElseIf optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadEvenScansOnly).Value = True Then
        EvenOddScanFilterMode = eosEvenOddScanFilterModeConstants.eosLoadEvenScansOnly
    Else
        EvenOddScanFilterMode = eosEvenOddScanFilterModeConstants.eosLoadAllScans
    End If
End Property

Public Property Let AbuFilterMax(ByVal dblValue As Double)
    txtAbuFilterMax = dblValue
End Property
Public Property Get AbuFilterMax() As Double
    If IsNumeric(txtAbuFilterMax) Then
        AbuFilterMax = txtAbuFilterMax
    Else
        AbuFilterMax = 1E+15
    End If
End Property

Public Property Let AbuFilterMin(ByVal dblValue As Double)
    txtAbuFilterMin = dblValue
End Property
Public Property Get AbuFilterMin() As Double
    If IsNumeric(txtAbuFilterMin) Then
        AbuFilterMin = txtAbuFilterMin
    Else
        AbuFilterMin = 0
    End If
End Property

Public Property Let IMSConformerScoreMin(ByVal dblValue As Double)
    txtIMSConformerScoreMin.Text = dblValue
End Property
Public Property Get IMSConformerScoreMin() As Double
    If IsNumeric(txtIMSConformerScoreMin) Then
        IMSConformerScoreMin = CDbl(txtIMSConformerScoreMin.Text)
    Else
        IMSConformerScoreMin = DEFAULT_IMS_CONFORMER_COMBINED_SCORE_MINIMUM
    End If
End Property

Public Property Let LCMSFeatureAbuMin(ByVal dblValue As Double)
    txtLCMSFeatureAbuMin = dblValue
End Property
Public Property Get LCMSFeatureAbuMin() As Double
    If IsNumeric(txtLCMSFeatureAbuMin) Then
        LCMSFeatureAbuMin = CDbl(txtLCMSFeatureAbuMin.Text)
    Else
        LCMSFeatureAbuMin = 0
    End If
End Property

Public Property Let LCMSFeatureScanCountMin(ByVal lngValue As Long)
    txtLCMSFeatureScanCountMin = lngValue
End Property
Public Property Get LCMSFeatureScanCountMin() As Long
    If IsNumeric(txtLCMSFeatureScanCountMin) Then
        LCMSFeatureScanCountMin = CLng(txtLCMSFeatureScanCountMin.Text)
    Else
        LCMSFeatureScanCountMin = 0
    End If
End Property

Public Property Let IsoFitMax(ByVal dblValue As Double)
    txtIsoFitMaxValue.Text = dblValue
End Property
Public Property Get IsoFitMax() As Double
    If IsNumeric(txtIsoFitMaxValue) Then
        IsoFitMax = CDbl(txtIsoFitMaxValue.Text)
    Else
        IsoFitMax = 0.15
    End If
End Property

' 0 means Abundance Sum
' 1 means Abundance Max
Public Property Let LCMSFeatureClassAbundanceMode(ByVal intValue As Integer)
    cboLCMSFeatureClassAbundanceMode.ListIndex = intValue
End Property
Public Property Get LCMSFeatureClassAbundanceMode() As Integer
    LCMSFeatureClassAbundanceMode = cboLCMSFeatureClassAbundanceMode.ListIndex
End Property

Public Property Let LCMSFeaturePointsLoadMode(ByVal intValue As Integer)
    cboLCMSFeaturePointsLoadMode.ListIndex = intValue
End Property
Public Property Get LCMSFeaturePointsLoadMode() As Integer
    LCMSFeaturePointsLoadMode = cboLCMSFeaturePointsLoadMode.ListIndex
End Property

Public Property Let LCMSFeatureSplitUMCsByExaminingAbundance(ByVal Value As Boolean)
    SetCheckBox chkSplitUMCsByExaminingAbundance, Value
End Property
Public Property Get LCMSFeatureSplitUMCsByExaminingAbundance() As Boolean
    LCMSFeatureSplitUMCsByExaminingAbundance = cChkBox(chkSplitUMCsByExaminingAbundance)
End Property

Public Property Get LoadCancelled() As Boolean
    LoadCancelled = mLoadCancelled
End Property

Public Property Let MaximumDataCountEnabled(ByVal Value As Boolean)
    SetCheckBox chkMaximumDataCountEnabled, Value
End Property
Public Property Get MaximumDataCountEnabled() As Boolean
    MaximumDataCountEnabled = cChkBox(chkMaximumDataCountEnabled)
End Property

Public Property Let MaximumDataCountToLoad(ByVal Value As Long)
    txtMaximumDataCountToLoad.Text = Value
End Property
Public Property Get MaximumDataCountToLoad() As Long
    If IsNumeric(txtMaximumDataCountToLoad) Then
        MaximumDataCountToLoad = CLng(txtMaximumDataCountToLoad)
    Else
        MaximumDataCountToLoad = DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD
    End If
End Property

Public Property Let TotalIntensityPercentageFilterEnabled(ByVal Value As Boolean)
    SetCheckBox chkTotalIntensityPercentageFilterEnabled, Value
End Property
Public Property Get TotalIntensityPercentageFilterEnabled() As Boolean
    TotalIntensityPercentageFilterEnabled = cChkBox(chkTotalIntensityPercentageFilterEnabled)
End Property

Public Property Let TotalIntensityPercentageFilter(ByVal Value As Single)
    txtTotalIntensityPercentageFilter.Text = Value
End Property
Public Property Get TotalIntensityPercentageFilter() As Single
    If IsNumeric(txtTotalIntensityPercentageFilter) Then
        TotalIntensityPercentageFilter = CSng(txtTotalIntensityPercentageFilter)
    Else
        TotalIntensityPercentageFilter = DEFAULT_TOTAL_INTENSITY_PERCENTAGE_TO_LOAD
    End If
End Property

Public Sub GetMSLevelFilter(ByRef blnMSLevelFilter() As Boolean)
    Dim intIndex As Integer
    Dim intSelectedCount As Integer
    
    ReDim blnMSLevelFilter(lstMSLevelFilter.ListCount - 1)
    
    intSelectedCount = 0
    For intIndex = 0 To lstMSLevelFilter.ListCount - 1
        blnMSLevelFilter(intIndex) = lstMSLevelFilter.Selected(intIndex)
        If blnMSLevelFilter(intIndex) Then intSelectedCount = intSelectedCount + 1
    Next intIndex
    
    If intSelectedCount = 0 Then
        blnMSLevelFilter(0) = True
    End If
End Sub

Private Sub EnableDisableControls()
    Dim blnEnableDataTypeFilter As Boolean
    Dim blnEnableIsoFitFilter As Boolean
    Dim blnEnableMSLevelFilter As Boolean
    Dim blnEnableDreamsFilters As Boolean
    Dim blnEnableAbundanceFilters As Boolean
    Dim blnEnableDatacountFilters As Boolean
    Dim blnEnableLCMSFeatureFilters As Boolean
    
    Dim intIndex As Integer
    
    Select Case mFileType
    Case ifmInputFileModeConstants.ifmCSVFile, ifmInputFileModeConstants.ifmPEKFile
        blnEnableDataTypeFilter = True
        blnEnableIsoFitFilter = True
        blnEnableMSLevelFilter = True
        blnEnableDreamsFilters = True
        blnEnableAbundanceFilters = True
        blnEnableDatacountFilters = True
        blnEnableLCMSFeatureFilters = False
    
    Case ifmInputFileModeConstants.ifmmzXMLFile, ifmInputFileModeConstants.ifmmzXMLFileWithXMLExtension
        blnEnableDataTypeFilter = False
        blnEnableIsoFitFilter = False
        blnEnableMSLevelFilter = True
        blnEnableDreamsFilters = True
        blnEnableAbundanceFilters = True
        blnEnableDatacountFilters = True
        blnEnableLCMSFeatureFilters = False
    
    Case ifmInputFileModeConstants.ifmmzDataFile, ifmInputFileModeConstants.ifmmzDataFileWithXMLExtension
        blnEnableDataTypeFilter = False
        blnEnableIsoFitFilter = False
        blnEnableMSLevelFilter = True
        blnEnableDreamsFilters = True
        blnEnableAbundanceFilters = True
        blnEnableDatacountFilters = True
        blnEnableLCMSFeatureFilters = False
    
    Case ifmDelimitedTextFile
        blnEnableDataTypeFilter = False
         
        If cboLCMSFeaturePointsLoadMode.ListIndex = plmLoadAllPoints Then
            blnEnableIsoFitFilter = True
            blnEnableAbundanceFilters = True
        Else
            blnEnableIsoFitFilter = False
            blnEnableAbundanceFilters = False
        End If
        
        blnEnableMSLevelFilter = False
        blnEnableDreamsFilters = False
        blnEnableDatacountFilters = False
        blnEnableLCMSFeatureFilters = True
    
    Case Else
        ' Unknown file type
        Debug.Assert False
        blnEnableDataTypeFilter = False
        blnEnableIsoFitFilter = False
        blnEnableMSLevelFilter = False
        blnEnableDreamsFilters = False
        blnEnableAbundanceFilters = False
        blnEnableDatacountFilters = False
        blnEnableLCMSFeatureFilters = False

    End Select

    optCSandIsoDataFilterMode(0).Enabled = blnEnableDataTypeFilter
    optCSandIsoDataFilterMode(1).Enabled = blnEnableDataTypeFilter
    optCSandIsoDataFilterMode(2).Enabled = blnEnableDataTypeFilter
    
    optIsoFitFilter(0).Enabled = blnEnableIsoFitFilter
    optIsoFitFilter(1).Enabled = blnEnableIsoFitFilter
    txtIsoFitMaxValue.Enabled = blnEnableIsoFitFilter
    
    optEvenOddScanFilter(0).Enabled = blnEnableDreamsFilters
    optEvenOddScanFilter(1).Enabled = blnEnableDreamsFilters
    optEvenOddScanFilter(2).Enabled = blnEnableDreamsFilters
    If Not blnEnableDreamsFilters Then
        optEvenOddScanFilter(0).Value = True
    End If

    optAbuFilter(0).Enabled = blnEnableAbundanceFilters
    optAbuFilter(1).Enabled = blnEnableAbundanceFilters
    If Not blnEnableAbundanceFilters Then
        optAbuFilter(0).Value = True
    End If
    
    chkMaximumDataCountEnabled.Enabled = blnEnableDatacountFilters
    chkTotalIntensityPercentageFilterEnabled.Enabled = blnEnableDatacountFilters
    If Not blnEnableDatacountFilters Then
        chkMaximumDataCountEnabled.Value = vbUnchecked
        chkTotalIntensityPercentageFilterEnabled.Value = vbUnchecked
    End If
    
    optFilterLCMSFeatures(0).Enabled = blnEnableLCMSFeatureFilters
    optFilterLCMSFeatures(1).Enabled = blnEnableLCMSFeatureFilters
    
    txtLCMSFeatureAbuMin.Enabled = blnEnableLCMSFeatureFilters
    txtLCMSFeatureScanCountMin.Enabled = blnEnableLCMSFeatureFilters
    txtIMSConformerScoreMin.Enabled = blnEnableLCMSFeatureFilters
        
    lstMSLevelFilter.Enabled = blnEnableMSLevelFilter
    If Not blnEnableMSLevelFilter Then
        lstMSLevelFilter.Selected(0) = True
        For intIndex = 1 To lstMSLevelFilter.ListCount - 1
            lstMSLevelFilter.Selected(intIndex) = False
        Next intIndex
    Else
        ' Could use this to default to MS1 data
        ' However, for now, we're defaulting to an MSLevel filter of "All Scans"
''        lstMSLevelFilter.Selected(0) = False
''        lstMSLevelFilter.Selected(1) = True
''        For intIndex = 1 To lstMSLevelFilter.ListCount - 1
''            lstMSLevelFilter.Selected(intIndex) = False
''        Next intIndex
    End If

    txtAutoMapDataPointsMassTolerancePPM.Enabled = blnEnableLCMSFeatureFilters
    cboLCMSFeaturePointsLoadMode.Enabled = blnEnableLCMSFeatureFilters
    chkSplitUMCsByExaminingAbundance.Enabled = blnEnableLCMSFeatureFilters
    
End Sub

Private Sub ResetToDefaults()
    Dim intIndex As Integer
    
    Me.MousePointer = vbDefault
    
    Me.FilterOnIsoFit = True
    Me.FilterOnAbundance = False
    Me.FilterLCMSFeatures = True
    
    Me.AbuFilterMin = 0
    Me.AbuFilterMax = 1E+15
    
    Me.IsoFitMax = 0.15
    
    optFilterLCMSFeatures(1).Value = True
    
    Me.LCMSFeatureAbuMin = 0
    Me.LCMSFeatureScanCountMin = DEFAULT_LCMS_FEATURE_SCAN_COUNT_MINIMUM
    Me.IMSConformerScoreMin = DEFAULT_IMS_CONFORMER_COMBINED_SCORE_MINIMUM
    
    chkMaximumDataCountEnabled.Value = vbChecked
    Me.txtMaximumDataCountToLoad = DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD
    
    chkTotalIntensityPercentageFilterEnabled.Value = vbUnchecked
    Me.txtTotalIntensityPercentageFilter.Text = DEFAULT_TOTAL_INTENSITY_PERCENTAGE_TO_LOAD
    
    mLoadCancelled = False
    
    optCSandIsoDataFilterMode(dfmCSandIsoDataFilterModeConstants.dfmLoadAllData).Value = True
    
    If Len(lblFilePath.Caption) = 0 Then
        optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadAllScans).Value = True
    ElseIf InStr(LCase(lblFilePath.Caption), "dreams") > 0 Then
        optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadEvenScansOnly).Value = True
    Else
        optEvenOddScanFilter(eosEvenOddScanFilterModeConstants.eosLoadAllScans).Value = True
    End If
    
    With lstMSLevelFilter
        .Clear
        .AddItem "All Scans"
        For intIndex = 1 To 4
            .AddItem "MS" + Trim(intIndex) & " scans"
        Next intIndex
        .Selected(1) = True
    End With
    
    With cboLCMSFeaturePointsLoadMode
        .Clear
        .AddItem ("Load All Points")                    ' 0
        .AddItem ("Load Mapped Points Only")            ' 1
        .AddItem ("Load 1 Point Per LC-MS Feature")     ' 2
        .ListIndex = 1
    End With
    
    With cboLCMSFeatureClassAbundanceMode
        .Clear
        .AddItem ("Abundance Sum")
        .AddItem ("Abundance Max")
        .ListIndex = 0
    End With
    
    txtAutoMapDataPointsMassTolerancePPM.Text = "5"
    chkSplitUMCsByExaminingAbundance.Value = vbUnchecked
    
    ' Call SetFileType to set file-type specific filters
    SetFileType mFileType
    
End Sub

Public Sub SetFileType(eFileType As ifmInputFileModeConstants)
    
    mFileType = eFileType
    
    EnableDisableControls
    
End Sub

Public Sub SetFilePath(strFilePath As String)
    ' Note: This sub will call SetFileType
    
    Dim eFileType As ifmInputFileModeConstants
    
    On Error GoTo InitializeErrorHandler
    
    If FileExists(strFilePath) Then
        mFileSize = FileLen(strFilePath)
    
        If mFileSize > FILE_SIZE_THRESHOLD_BYTES Then
            optIsoFitFilter(1).Value = True
        Else
            optIsoFitFilter(1).Value = False
        End If
        
        lblFileSize = Round(mFileSize / 1024 / 1024, 1) & " MB"
    Else
        lblFileSize = "?? MB"
    End If
    
    lblFilePath = CompactPathString(strFilePath, 65)
    
    If DetermineFileType(strFilePath, eFileType) Then
        SetFileType eFileType
    End If
    
    Exit Sub

InitializeErrorHandler:
    Debug.Print "Error in SetFilePath: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmFileLoadOptions->SetFilePath"
    Resume Next
    
End Sub

Private Sub cboLCMSFeaturePointsLoadMode_Click()
    EnableDisableControls
End Sub

Private Sub cmdCancel_Click()
    mLoadCancelled = True
    Me.Hide
End Sub

Private Sub cmdLoad_Click()
    mLoadCancelled = False
    Me.Hide
End Sub

Private Sub cmdSetToDefaults_Click()
    ResetToDefaults
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowUpperThird, -1, -1, False
    mFileType = ifmInputFileModeConstants.ifmPEKFile
    ResetToDefaults
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub txtAbuFilterMax_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAbuFilterMax, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtAbuFilterMax_LostFocus()
    ValidateDualTextBoxes txtAbuFilterMin, txtAbuFilterMax, False, 0, 1E+200, 1
End Sub

Private Sub txtAbuFilterMin_Change()
    If IsNumeric(txtAbuFilterMin) Then
        If val(txtAbuFilterMin) > 0 Then
            optAbuFilter(1).Value = True
        End If
    End If
End Sub

Private Sub txtAbuFilterMin_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAbuFilterMin, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtAbuFilterMin_LostFocus()
    ValidateDualTextBoxes txtAbuFilterMin, txtAbuFilterMax, True, 0, 1E+200, 1
End Sub

Private Sub txtIMSConformerScoreMin_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIMSConformerScoreMin, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtIMSConformerScoreMin_LostFocus()
    ValidateTextboxValueDbl txtIMSConformerScoreMin, 0, 1, 0.75
End Sub

Private Sub txtIsoFitMaxValue_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoFitMaxValue, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtIsoFitMaxValue_LostFocus()
    ValidateTextboxValueDbl txtIsoFitMaxValue, 0, 1, 0.15
End Sub

Private Sub txtLCMSFeatureAbuMin_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtLCMSFeatureAbuMin, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtLCMSFeatureAbuMin_LostFocus()
    ValidateTextboxValueDbl txtLCMSFeatureAbuMin, 0, 1E+200, 0
End Sub

Private Sub txtLCMSFeatureScanCountMin_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtLCMSFeatureScanCountMin, KeyAscii, True, False, False, False, True, False, False, False, False, True
End Sub

Private Sub txtLCMSFeatureScanCountMin_LostFocus()
    ValidateTextboxValueLng txtLCMSFeatureScanCountMin, 1, 1000000#, 2
End Sub

Private Sub txtMaximumDataCountToLoad_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMaximumDataCountToLoad, KeyAscii, True, False
End Sub

Private Sub txtTotalIntensityPercentageFilter_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtTotalIntensityPercentageFilter, KeyAscii, True, True, False
End Sub

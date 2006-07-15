VERSION 5.00
Begin VB.Form frmFileLoadOptions 
   Caption         =   "File Load Options"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDREAMS 
      Caption         =   "DREAMS Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   3520
      Width           =   3375
      Begin VB.OptionButton optEvenOddScanFilter 
         Caption         =   "Only load odd-numbered scans"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   540
         Width           =   2800
      End
      Begin VB.OptionButton optEvenOddScanFilter 
         Caption         =   "Load all scans"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   2800
      End
      Begin VB.OptionButton optEvenOddScanFilter 
         Caption         =   "Only load even-numbered scans"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   2800
      End
   End
   Begin VB.CommandButton cmdSetToDefaults 
      Caption         =   "Set to &Defaults"
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame fraOptionFrame 
      Caption         =   "Data Count Filter"
      Height          =   1215
      Index           =   22
      Left            =   3600
      TabIndex        =   19
      Top             =   2160
      Width           =   3615
      Begin VB.TextBox txtMaximumDataCountToLoad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Text            =   "125000"
         ToolTipText     =   "Higher abundance data is favored when determine the data to load"
         Top             =   800
         Width           =   1095
      End
      Begin VB.CheckBox chkMaximumDataCountEnabled 
         Caption         =   "Maximum data count filter enabled (requires pre-scan of data file)"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   230
         Width           =   2895
      End
      Begin VB.Label lblDescription 
         Caption         =   "Maximum data count to load"
         Height          =   255
         Index           =   137
         Left            =   120
         TabIndex        =   21
         Top             =   830
         Width           =   2055
      End
   End
   Begin VB.Frame fraMSLevelFilter 
      Caption         =   "MS Level Filter"
      Height          =   1215
      Left            =   3600
      TabIndex        =   27
      Top             =   3520
      Width           =   1455
      Begin VB.ListBox lstMSLevelFilter 
         Height          =   840
         ItemData        =   "frmFileLoadOptions.frx":0000
         Left            =   120
         List            =   "frmFileLoadOptions.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraCSandIsoDataFilter 
      Caption         =   "Data Type Filter"
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   3375
      Begin VB.OptionButton optCSandIsoDataFilterMode 
         Caption         =   "Only load Isotopic data"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   3000
      End
      Begin VB.OptionButton optCSandIsoDataFilterMode 
         Caption         =   "Only load Charge State data"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   540
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
      TabIndex        =   31
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   5640
      TabIndex        =   29
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Frame fraIsoFitFilter 
      Caption         =   "Isotopic Fit Filter"
      Height          =   975
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

Public Property Get LoadCancelled() As Boolean
    LoadCancelled = mLoadCancelled
End Property

Public Property Let FilterOnIsoFit(blnEnable As Boolean)
    If blnEnable Then
        optIsoFitFilter(1).Value = blnEnable
    Else
        optIsoFitFilter(0).Value = blnEnable
    End If
End Property
Public Property Get FilterOnIsoFit() As Boolean
    FilterOnIsoFit = optIsoFitFilter(1).Value
End Property

Public Property Let FilterOnAbundance(blnEnable As Boolean)
    If blnEnable Then
        optAbuFilter(1).Value = True
    Else
        optAbuFilter(0).Value = True
    End If
End Property
Public Property Get FilterOnAbundance() As Boolean
    FilterOnAbundance = optAbuFilter(1).Value
End Property

Public Property Let DataFilterMode(eDataFilterMode As dfmCSandIsoDataFilterModeConstants)
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

Public Property Let EvenOddScanFilterMode(eEvenOddScanFilterMode As eosEvenOddScanFilterModeConstants)
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

Public Property Let AbuFilterMax(dblValue As Double)
    txtAbuFilterMax = dblValue
End Property
Public Property Get AbuFilterMax() As Double
    If IsNumeric(txtAbuFilterMax) Then
        AbuFilterMax = txtAbuFilterMax
    Else
        AbuFilterMax = 1E+15
    End If
End Property

Public Property Let AbuFilterMin(dblValue As Double)
    txtAbuFilterMin = dblValue
End Property
Public Property Get AbuFilterMin() As Double
    If IsNumeric(txtAbuFilterMin) Then
        AbuFilterMin = txtAbuFilterMin
    Else
        AbuFilterMin = 0
    End If
End Property

Public Property Let IsoFitMax(dblValue As Double)
    txtIsoFitMaxValue = dblValue
End Property
Public Property Get IsoFitMax() As Double
    If IsNumeric(txtIsoFitMaxValue) Then
        IsoFitMax = txtIsoFitMaxValue
    Else
        IsoFitMax = 0.15
    End If
End Property

Public Property Let MaximumDataCountEnabled(Value As Boolean)
    SetCheckBox chkMaximumDataCountEnabled, Value
End Property
Public Property Get MaximumDataCountEnabled() As Boolean
    MaximumDataCountEnabled = cChkBox(chkMaximumDataCountEnabled)
End Property

Public Property Let MaximumDataCountToLoad(Value As Long)
    txtMaximumDataCountToLoad.Text = Value
End Property
Public Property Get MaximumDataCountToLoad() As Long
    If IsNumeric(MaximumDataCountToLoad) Then
        MaximumDataCountToLoad = CLng(txtMaximumDataCountToLoad)
    Else
        MaximumDataCountToLoad = DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD
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

Private Sub ResetToDefaults()
    Dim intIndex As Integer
    
    Me.MousePointer = vbDefault
    
    Me.FilterOnIsoFit = True
    Me.FilterOnAbundance = False
    
    Me.AbuFilterMin = 0
    Me.AbuFilterMax = 1E+15
    
    Me.IsoFitMax = 0.15
    
    chkMaximumDataCountEnabled.Value = vbChecked
    Me.txtMaximumDataCountToLoad = DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD
    
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
    
    SetFileType mFileType
End Sub

Public Sub SetFileType(eFileType As ifmInputFileModeConstants)
    Dim blnEnableDataTypeFilter As Boolean
    Dim blnEnableIsoFitFilter As Boolean
    Dim blnEnableMSLevelFilter As Boolean
    Dim intIndex As Integer
    
    mFileType = eFileType
    Select Case mFileType
    Case ifmInputFileModeConstants.ifmCSVFile, ifmInputFileModeConstants.ifmPEKFile
        blnEnableDataTypeFilter = True
        blnEnableIsoFitFilter = True
        blnEnableMSLevelFilter = False
    Case ifmInputFileModeConstants.ifmmzXMLFile, ifmInputFileModeConstants.ifmmzXMLFileWithXMLExtension
        blnEnableDataTypeFilter = False
        blnEnableIsoFitFilter = False
        blnEnableMSLevelFilter = True
    Case ifmInputFileModeConstants.ifmmzDataFile, ifmInputFileModeConstants.ifmmzDataFileWithXMLExtension
        blnEnableDataTypeFilter = False
        blnEnableIsoFitFilter = False
        blnEnableMSLevelFilter = True
    Case Else
        ' Unknown file type
        Debug.Assert False
        blnEnableDataTypeFilter = False
        blnEnableIsoFitFilter = False
        blnEnableMSLevelFilter = False
    End Select

    optCSandIsoDataFilterMode(0).Enabled = blnEnableDataTypeFilter
    optCSandIsoDataFilterMode(1).Enabled = blnEnableDataTypeFilter
    optCSandIsoDataFilterMode(2).Enabled = blnEnableDataTypeFilter
    
    optIsoFitFilter(0).Enabled = blnEnableIsoFitFilter
    optIsoFitFilter(1).Enabled = blnEnableIsoFitFilter
    txtIsoFitMaxValue.Enabled = blnEnableIsoFitFilter
    
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

End Sub

Public Sub SetFilePath(strFilePath As String)
    ' Note: This sub will call SetFileType
    
    Dim eFileType As ifmInputFileModeConstants
    
    On Error GoTo InitializeErrorHandler
    
    mFileSize = FileLen(strFilePath)

    If mFileSize > FILE_SIZE_THRESHOLD_BYTES Then
        optIsoFitFilter(1).Value = True
    Else
        optIsoFitFilter(1).Value = False
    End If
    
    lblFileSize = Round(mFileSize / 1024 / 1024, 1) & " MB"
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

Private Sub txtIsoFitMaxValue_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoFitMaxValue, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtIsoFitMaxValue_LostFocus()
    ValidateTextboxValueDbl txtIsoFitMaxValue, 0, 1, 0.15
End Sub

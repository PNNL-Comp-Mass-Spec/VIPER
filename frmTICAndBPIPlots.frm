VERSION 5.00
Begin VB.Form frmTICAndBPIPlots 
   BackColor       =   &H00FFFFFF&
   Caption         =   "TIC and BPI Plots"
   ClientHeight    =   5805
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAutoUpdatePlot 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Auto update plot using current zoom range"
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   240
      Width           =   3615
   End
   Begin VB.Timer tmrUpdatePlot 
      Interval        =   200
      Left            =   7800
      Top             =   480
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      Height          =   2745
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      Begin VB.CheckBox chkClipOutlierValues 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clip Outlier Values"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtClipOutliersFactor 
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "100"
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cboPointShape 
         Height          =   315
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2085
         Width           =   1695
      End
      Begin VB.ComboBox cboPointShape 
         Height          =   315
         Index           =   0
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkSmoothUsingMovingAverage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Smooth Using Moving Average (Ctrl+M)"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1500
         Width           =   3255
      End
      Begin VB.TextBox txtSmoothUsingMovingAverage 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "3"
         Top             =   1860
         Width           =   615
      End
      Begin VB.CheckBox chkNormalizeYAxis 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Normalize Y Axis to 100  (Ctrl+N)"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkPlotNETOnXAxis 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Plot NET on X Axis"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkDrawLinesBetweenPoints 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Connect Points with Line"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkShowGridlines 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Gridlines"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtGraphPointSize 
         Height          =   285
         Left            =   5280
         TabIndex        =   16
         Text            =   "3"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtGraphLineWidth 
         Height          =   285
         Left            =   5280
         TabIndex        =   18
         Text            =   "1"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkAutoScaleXRange 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Scale X Range"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkShowPointSymbols 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Point Symbols"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Label lblClipOutliersFactor 
         BackStyle       =   0  'Transparent
         Caption         =   "fold from median"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   2300
         Width           =   1335
      End
      Begin VB.Label bllSeries2Style 
         BackStyle       =   0  'Transparent
         Caption         =   "Second Series"
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblPointColorSelection 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   23
         ToolTipText     =   "Double click to change"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label bllSeries1Style 
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Data"
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   1040
         Width           =   1095
      End
      Begin VB.Label lblPointColorSelection 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   20
         ToolTipText     =   "Double click to change"
         Top             =   1040
         Width           =   375
      End
      Begin VB.Label lblSmoothUsingMovingAverageUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "points"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1890
         Width           =   450
      End
      Begin VB.Label lblSmoothUsingMovingAverage 
         BackStyle       =   0  'Transparent
         Caption         =   "Window width"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label lblGraphPointSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Point Size"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblGraphLineWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Line Width"
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   630
         Width           =   975
      End
   End
   Begin VB.ComboBox cboTICToDisplay 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   4815
   End
   Begin VIPER.ctlSpectraPlotter ctlChromatogramPlot 
      Height          =   4935
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8705
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      Height          =   255
      Left            =   5040
      TabIndex        =   25
      Top             =   0
      Width           =   4935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveDataToTextFile 
         Caption         =   "Save Data to Text File"
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
      Begin VB.Menu mnuCopyChromatogram 
         Caption         =   "&Copy Data"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy Data for All Chromatograms"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy as &BMP"
         Index           =   0
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy as &WMF"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy as &EMF"
         Index           =   2
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Plot &Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewKeepWindowOnTop 
         Caption         =   "&Keep Window On Top"
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "frmTICAndBPIPlots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ccmCopyChartMode
    ccmWMF = 1
    ccmEMF = 2
    ccmBMP = 0
End Enum

Private Type udtBPIInfoType
    Mass As Double              ' Monoisotopic Mass for Isotopic data; Average mass for CS data
    OriginalIndex As Long
    CSData As Boolean           ' True if OriginalIndex points to data in .CSData()
End Type

Private Type udtChromatogramType
    RawIntensity() As Double            ' 1-based array  (Olectra Chart requires 1-based)
    NormalizedIntensity() As Double     ' 1-based array  (Olectra Chart requires 1-based)
    MaximumValue As Double
    SmoothWindowWidth As Long           ' Set to 0 if smoothing wasn't used; otherwise, set to the window size used for smoothing
    DualSeriesData As Boolean
End Type

Private Type udtTICandBPIChromType
    Chromatograms(TIC_AND_BPI_TYPE_COUNT - 1) As udtChromatogramType
    
    BPIInfo() As udtBPIInfoType             ' 1-based array
    GANETVals() As Double                   ' 1-based array
    
    ScanCount As Long               ' The above arrays range from 1 to ScanCount
    ScanNumberStart As Long
    ScanNumberEnd  As Long
End Type

Public CallerID As Long

Private mChromData As udtTICandBPIChromType
Private mCurrentPlotType As tbcTICAndBPIConstants

Private mUpdatingControls As Boolean
Private mFormInitialized As Boolean

Private mNeedToRecomputeChromatogram As Boolean
Private mNeedToUpdatePlot As Boolean
Private mAutoUpdatePlot As Boolean

Private mWindowStayOnTopEnabled As Boolean

Public Sub AutoUpdatePlot()
    If mAutoUpdatePlot Then
        mNeedToRecomputeChromatogram = True
    End If
End Sub

Public Function ClipDataVsMedian(ByRef dblValues() As Double, ByVal dblClipOutliersFactor As Double) As Boolean
    
    ' Looks for values in dblValues that are more than dblClipOutliersFactor away from the median
    ' Replaces their value with the maximum or minimum allowable value
    
    Dim objStatDoubles As New StatDoubles
    
    Dim dblMedian As Double
    Dim dblClipMin As Double
    Dim dblClipMax As Double
    
    Dim lngIndex As Long
    Dim blnSuccess As Boolean
    
    On Error GoTo ClipDataVsMedianErrorHandler

    ' Find the median of the data in dblValuesSorted
    If objStatDoubles.Fill(dblValues) Then
        
        dblMedian = objStatDoubles.Median
        
        dblClipMax = dblMedian * dblClipOutliersFactor
        
        If dblClipOutliersFactor > 0 Then
            dblClipMin = dblMedian / dblClipOutliersFactor
        Else
            dblClipMin = 0
        End If
        
        For lngIndex = LBound(dblValues) To UBound(dblValues)
            If dblValues(lngIndex) > dblClipMax Then
                dblValues(lngIndex) = dblClipMax
                'dblValues(lngIndex) = dblMedian
            Else
                If dblValues(lngIndex) < dblClipMin Then
                    dblValues(lngIndex) = dblClipMin
                    'dblValues(lngIndex) = dblMedian
                End If
            End If
        Next lngIndex
    Else
        blnSuccess = False
    End If
    
    ClipDataVsMedian = blnSuccess
    Exit Function

ClipDataVsMedianErrorHandler:
    Debug.Assert False
    ClipDataVsMedian = False
    
End Function

Private Sub ComputeAndDisplayChromatogram()
    
    Dim lngCSCount As Long, lngIsoCount As Long
    Dim lngCSPointerArray() As Long         ' 1-based array
    Dim lngIsoPointerArray() As Long        ' 1-based array
    
    Dim lngOriginalIndex As Long, lngPointerIndex As Long
    Dim lngChromIndex As Long
    
    Dim lngScanNumber As Long
    Dim dblAbundance As Double
    Dim dblTimeDomainDataMaxValue As Double
    
    Dim lngCountDimmed As Long
    Dim eChromType As tbcTICAndBPIConstants
    
    Dim dblMaximum As Double
    
    If mUpdatingControls Then Exit Sub
    mNeedToRecomputeChromatogram = False
    
On Error GoTo ComputeAndDisplayChromatogramErrorHandler
    
    If CallerID < 1 Or CallerID > UBound(GelData()) Then Exit Sub
    If GelStatus(CallerID).Deleted Then Exit Sub
    
    UpdateStatus "Preparing chromatogram"
    
    With mChromData
        ' Determine the number of scans for the current gel and initialize the arrays
        GetScanRange CallerID, .ScanNumberStart, .ScanNumberEnd, .ScanCount
        
        lngCountDimmed = .ScanCount
        If lngCountDimmed = 0 Then lngCountDimmed = 1
        
        ' Initialize the chromatogram arrays
        For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
            With .Chromatograms(eChromType)
                ReDim .RawIntensity(1 To lngCountDimmed)
                ReDim .NormalizedIntensity(1 To lngCountDimmed)
                .MaximumValue = 0
                .DualSeriesData = False
            End With
        Next eChromType
        ReDim .BPIInfo(1 To lngCountDimmed)
        ReDim .GANETVals(1 To lngCountDimmed)
        
        ' Two of the chromatograms are actually dual series chromatograms, rather than raw and normalized data
        .Chromatograms(tbcDeisotopingIntensityThresholds).DualSeriesData = True
        .Chromatograms(tbcDeisotopingPeakCounts).DualSeriesData = True

    End With
    
    ' Retrieve an array of the ion indices of the ions currently "In Scope"
    ' Note that GetISScope will ReDim each PointerArray() automatically
    lngCSCount = GetCSScope(CallerID, lngCSPointerArray(), glScope.glSc_Current)
    lngIsoCount = GetISScope(CallerID, lngIsoPointerArray(), glScope.glSc_Current)
    
    With glbPreferencesExpanded.TICAndBPIPlottingOptions
        If .TimeDomainDataMaxValue > 0 Then
            dblTimeDomainDataMaxValue = .TimeDomainDataMaxValue
        Else
            dblTimeDomainDataMaxValue = glHugeDouble
        End If
    End With

    
    ' Construct the chromatograms
    With mChromData
        ' Charge state data
        For lngPointerIndex = 1 To lngCSCount
            lngOriginalIndex = lngCSPointerArray(lngPointerIndex)
            
            lngScanNumber = GelData(CallerID).CSData(lngOriginalIndex).ScanNumber
            dblAbundance = GelData(CallerID).CSData(lngOriginalIndex).Abundance
            
            lngChromIndex = lngScanNumber - .ScanNumberStart + 1
            
            With .Chromatograms(tbcTICFromCurrentDataIntensities)
                .RawIntensity(lngChromIndex) = .RawIntensity(lngChromIndex) + dblAbundance
            End With
            
            With .Chromatograms(tbcTICFromCurrentDataPointCounts)
                .RawIntensity(lngChromIndex) = .RawIntensity(lngChromIndex) + 1
            End With
            
            If dblAbundance > .Chromatograms(tbcBPIFromCurrentDataIntensities).RawIntensity(lngChromIndex) Then
                .Chromatograms(tbcBPIFromCurrentDataIntensities).RawIntensity(lngChromIndex) = dblAbundance
                With .BPIInfo(lngChromIndex)
                    .Mass = GelData(CallerID).CSData(lngOriginalIndex).AverageMW
                    .OriginalIndex = lngOriginalIndex
                    .CSData = True
                End With
            End If
            
            If lngPointerIndex Mod 250 = 0 Then UpdateStatus "Preparing chromatograms: " & Trim(lngPointerIndex) & " / " & Trim(lngCSCount)
        Next lngPointerIndex
        
        ' Isotopic data
        For lngPointerIndex = 1 To lngIsoCount
            lngOriginalIndex = lngIsoPointerArray(lngPointerIndex)
            
            lngScanNumber = GelData(CallerID).IsoData(lngOriginalIndex).ScanNumber
            dblAbundance = GelData(CallerID).IsoData(lngOriginalIndex).Abundance
            
            lngChromIndex = lngScanNumber - .ScanNumberStart + 1
            
            With .Chromatograms(tbcTICFromCurrentDataIntensities)
                .RawIntensity(lngChromIndex) = .RawIntensity(lngChromIndex) + dblAbundance
            End With
            
            With .Chromatograms(tbcTICFromCurrentDataPointCounts)
                .RawIntensity(lngChromIndex) = .RawIntensity(lngChromIndex) + 1
            End With
            
            If dblAbundance > .Chromatograms(tbcBPIFromCurrentDataIntensities).RawIntensity(lngChromIndex) Then
                .Chromatograms(tbcBPIFromCurrentDataIntensities).RawIntensity(lngChromIndex) = dblAbundance
                With .BPIInfo(lngChromIndex)
                    .Mass = GelData(CallerID).IsoData(lngOriginalIndex).MonoisotopicMW
                    .OriginalIndex = lngOriginalIndex
                    .CSData = False
                End With
            End If
            
            If lngPointerIndex Mod 250 = 0 Then UpdateStatus "Preparing chromatograms: " & Trim(lngPointerIndex) & " / " & Trim(lngIsoCount)
        Next lngPointerIndex
    
        ' Data from .ScanInfo()
        ' Must lookup the scan number in ScanInfo()
        For lngPointerIndex = 1 To UBound(GelData(CallerID).ScanInfo())
            lngScanNumber = GelData(CallerID).ScanInfo(lngPointerIndex).ScanNumber
            lngChromIndex = lngScanNumber - .ScanNumberStart + 1
            
            With .Chromatograms(tbcTICFromTimeDomain)
                .RawIntensity(lngChromIndex) = GelData(CallerID).ScanInfo(lngPointerIndex).TimeDomainSignal
                
                If .RawIntensity(lngChromIndex) > dblTimeDomainDataMaxValue Then
                    .RawIntensity(lngChromIndex) = dblTimeDomainDataMaxValue
                End If
            End With
    
            With .Chromatograms(tbcTICFromRawData)
                .RawIntensity(lngChromIndex) = GelData(CallerID).ScanInfo(lngPointerIndex).TIC
            End With
            With .Chromatograms(tbcBPIFromRawData)
                .RawIntensity(lngChromIndex) = GelData(CallerID).ScanInfo(lngPointerIndex).BPI
            End With
            
            With .Chromatograms(tbcDeisotopingIntensityThresholds)
                .RawIntensity(lngChromIndex) = GelData(CallerID).ScanInfo(lngPointerIndex).PeakIntensityThreshold
                .NormalizedIntensity(lngChromIndex) = GelData(CallerID).ScanInfo(lngPointerIndex).PeptideIntensityThreshold
            End With
            With .Chromatograms(tbcDeisotopingPeakCounts)
                .RawIntensity(lngChromIndex) = GelData(CallerID).ScanInfo(lngPointerIndex).NumDeisotoped
                .NormalizedIntensity(lngChromIndex) = GelData(CallerID).ScanInfo(lngPointerIndex).NumPeaks
            End With

        Next lngPointerIndex
  
        For lngChromIndex = 1 To .ScanCount
            .GANETVals(lngChromIndex) = ScanToGANET(CallerID, lngChromIndex + .ScanNumberStart - 1)
        Next lngChromIndex
    
        ' Step through each of the chromatograms and interpolate the abundances between non-zero data points
        '  that have scan gaps > 1 but, according to .ScanInfo(), are actually adjacent scans
        For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
            InterpolateChromatogramGaps CallerID, .Chromatograms(eChromType).RawIntensity(), 1, .ScanCount, .ScanNumberStart
        Next eChromType
        
        If glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage Then
            ' Smooth each of the chromatograms using a moving average filter
            For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
                .Chromatograms(eChromType).SmoothWindowWidth = glbPreferencesExpanded.TICAndBPIPlottingOptions.MovingAverageWindowWidth
                SmoothViaMovingAverage .Chromatograms(eChromType).RawIntensity(), 1, .ScanCount, .Chromatograms(eChromType).SmoothWindowWidth, 1
            Next eChromType
        
            ' Additionally need to smooth the second series data in the dual series (Deisotoping) chromatograms
            For eChromType = tbcDeisotopingIntensityThresholds To tbcDeisotopingPeakCounts
                .Chromatograms(eChromType).SmoothWindowWidth = glbPreferencesExpanded.TICAndBPIPlottingOptions.MovingAverageWindowWidth
                SmoothViaMovingAverage .Chromatograms(eChromType).NormalizedIntensity(), 1, .ScanCount, .Chromatograms(eChromType).SmoothWindowWidth, 1
            Next eChromType
        
        Else
            For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
                .Chromatograms(eChromType).SmoothWindowWidth = 0
            Next eChromType
        End If
        
        ' Optionally clip the data relative to the median
        ' Normalize each of the chromatograms and store in .NormalizedIntensity
        For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
            With .Chromatograms(eChromType)
                
                If glbPreferencesExpanded.TICAndBPIPlottingOptions.ClipOutliers Then
                    ' Clip the data relative to the median
                    ClipDataVsMedian .RawIntensity, glbPreferencesExpanded.TICAndBPIPlottingOptions.ClipOutliersFactor
                    If .DualSeriesData Then
                        ClipDataVsMedian .NormalizedIntensity, glbPreferencesExpanded.TICAndBPIPlottingOptions.ClipOutliersFactor
                    End If
                End If
                          
                dblMaximum = 0
                For lngChromIndex = 1 To mChromData.ScanCount
                    If .RawIntensity(lngChromIndex) > dblMaximum Then
                        dblMaximum = .RawIntensity(lngChromIndex)
                    End If
                Next lngChromIndex
                
                If Not .DualSeriesData Then
                    If dblMaximum > 0 Then
                        For lngChromIndex = 1 To mChromData.ScanCount
                            .NormalizedIntensity(lngChromIndex) = .RawIntensity(lngChromIndex) / dblMaximum * 100
                        Next lngChromIndex
                    End If
                End If
                .MaximumValue = dblMaximum
            End With
        Next eChromType
        
    End With
    
    UpdatePlot
    
    Exit Sub
    
ComputeAndDisplayChromatogramErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Unexpected error in ComputeAndDisplayChromatogram" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    LogErrors Err.Number, "ComputeAndDisplayChromatogram"
    
End Sub

Private Function ExportChromatogramCheckInclusion(eChromType, blnIncludeAllChromatograms, blnExportCurrentViewChromatogramsOnly) As Boolean

    Dim blnInclude As Boolean

    If blnIncludeAllChromatograms Then
        blnInclude = True
    ElseIf blnExportCurrentViewChromatogramsOnly Then
        Select Case eChromType
        Case tbcTICFromCurrentDataIntensities, tbcBPIFromCurrentDataIntensities, tbcTICFromCurrentDataPointCounts
            blnInclude = True
        Case Else
            blnInclude = False
        End Select
    ElseIf eChromType = mCurrentPlotType Then
        blnInclude = True
    Else
        blnInclude = False
    End If

    ExportChromatogramCheckInclusion = blnInclude
    
End Function

Public Function ExportChromatogramDataToClipboardOrFile(Optional strFilePath As String = "", Optional blnIncludeAllChromatograms As Boolean = False, Optional blnShowMessages As Boolean = True, Optional blnExportCurrentViewChromatogramsOnly As Boolean = False) As Long
    ' Returns 0 if success, the error code if an error

    Dim strData() As String
    Dim strTextToCopy As String
    Dim lngScanNumber As Long
    
    Dim OutFileNum As Integer
    Dim lngIndex As Long, lngOutputArrayCount As Long
    Dim eChromType As tbcTICAndBPIConstants
    Dim blnUseNormalized As Boolean
    
    If mChromData.ScanCount = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        ExportChromatogramDataToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo ExportChromatogramDataToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass
    UpdateStatus "Exporting"
    
    blnUseNormalized = glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis
    
    ' Header row is strData(0)
    ' Data is from strData(1) to strData(mChromData.ScanCount)
    ReDim strData(0 To mChromData.ScanCount + 2)
    
    ' Fill strData()
    ' Define the header row
    strData(0) = "ScanNumber" & vbTab & "NET"
    For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
        If ExportChromatogramCheckInclusion(eChromType, blnIncludeAllChromatograms, blnExportCurrentViewChromatogramsOnly) Then
            strData(0) = strData(0) & vbTab & GetChromDescription(eChromType, False)
        End If
    Next eChromType
    
    With mChromData
        For lngIndex = 1 To mChromData.ScanCount
            lngScanNumber = lngIndex + mChromData.ScanNumberStart - 1
            strData(lngIndex) = Trim(lngScanNumber) & vbTab & Round(.GANETVals(lngIndex), 5)
            
            For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
                If ExportChromatogramCheckInclusion(eChromType, blnIncludeAllChromatograms, blnExportCurrentViewChromatogramsOnly) Then
                    If blnUseNormalized Then
                        strData(lngIndex) = strData(lngIndex) & vbTab & Trim(.Chromatograms(eChromType).NormalizedIntensity(lngIndex))
                    Else
                        strData(lngIndex) = strData(lngIndex) & vbTab & Trim(.Chromatograms(eChromType).RawIntensity(lngIndex))
                    End If
                End If
            Next eChromType
        Next lngIndex
        lngOutputArrayCount = .ScanCount + 1
    End With
    
    If Len(strFilePath) > 0 Then
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        For lngIndex = 0 To lngOutputArrayCount - 1
            Print #OutFileNum, strData(lngIndex)
        Next lngIndex
        
        Close #OutFileNum
    Else
        strTextToCopy = FlattenStringArray(strData(), lngOutputArrayCount, vbCrLf, False)
        Clipboard.Clear
        Clipboard.SetText strTextToCopy, vbCFText
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus "Ready"
    
    ExportChromatogramDataToClipboardOrFile = 0
    Exit Function

ExportChromatogramDataToClipboardOrFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting chromatogram data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    ExportChromatogramDataToClipboardOrFile = Err.Number
    
End Function

Public Sub ForceRecomputeChromatogram()
    ' Forces a call to ComputeAndDisplayChromatogram
    ComputeAndDisplayChromatogram
End Sub

Public Sub ForceUpdatePlot()
    ' Forces a call to UpdatePlot
    UpdatePlot
End Sub

Public Function GetChromDescription(ePlotType As tbcTICAndBPIConstants, blnIncludeSpaces As Boolean) As String

    Dim strDescription As String
    
    If blnIncludeSpaces Then
        Select Case ePlotType
        Case tbcTICFromCurrentDataIntensities
            strDescription = "TIC From Data Intensities"
        Case tbcBPIFromCurrentDataIntensities
            strDescription = "Base Peak Intensity (BPI)"
        Case tbcTICFromTimeDomain
            strDescription = "TIC From Time Domain Signal"
        Case tbcTICFromCurrentDataPointCounts
            strDescription = "TIC From Data Point Counts"
        Case tbcTICFromRawData
            strDescription = "TIC From Raw Data"
        Case tbcBPIFromRawData
            strDescription = "BPI From Raw Data"
        Case tbcDeisotopingIntensityThresholds
            strDescription = "Deisotoping Intensity Thresholds"
        Case tbcDeisotopingPeakCounts
            strDescription = "Deisotoping Peak Counts"
        Case Else
            Debug.Assert False
            strDescription = "Unknown Chromatogram"
        End Select
    Else
        Select Case ePlotType
        Case tbcTICFromCurrentDataIntensities
            strDescription = "TICFromDataIntensities"
        Case tbcBPIFromCurrentDataIntensities
            strDescription = "BPI"
        Case tbcTICFromTimeDomain
            strDescription = "TICFromTimeDomainSignal"
        Case tbcTICFromCurrentDataPointCounts
            strDescription = "TICFromDataPointCounts"
        Case tbcTICFromRawData
            strDescription = "TICFromRawData"
        Case tbcBPIFromRawData
            strDescription = "BPIFromRawData"
        Case tbcDeisotopingIntensityThresholds
            strDescription = "DeisotopingIntensityThresholds"
        Case tbcDeisotopingPeakCounts
            strDescription = "DeisotopingPeakCounts"
        Case Else
            strDescription = "UnknownChromatogram"
        End Select
    End If
    
    GetChromDescription = strDescription
    
End Function

Public Sub InitializeForm()
    mUpdatingControls = True
    
On Error GoTo InitializeFormErrorHandler

    ' Update the controls with the values in .ErrorPlottingOptions
    With glbPreferencesExpanded.TICAndBPIPlottingOptions
        SetCheckBox chkPlotNETOnXAxis, .PlotNETOnXAxis
        SetCheckBox chkNormalizeYAxis, .NormalizeYAxis
        SetCheckBox chkSmoothUsingMovingAverage, .SmoothUsingMovingAverage
        txtSmoothUsingMovingAverage = Trim(.MovingAverageWindowWidth)
        
        With .Graph2DOptions
            SetCheckBox chkAutoScaleXRange, .AutoScaleXAxis
            SetCheckBox chkShowPointSymbols, .ShowPointSymbols
            SetCheckBox chkDrawLinesBetweenPoints, .DrawLinesBetweenPoints
            SetCheckBox chkShowGridlines, .ShowGridLines
            
            txtGraphPointSize = Trim(.PointSizePixels)
            
            If .PointShape < 1 Or .PointShape > OlectraChart2D.ShapeConstants.oc2dShapeSquare Then
                .PointShape = OlectraChart2D.ShapeConstants.oc2dShapeDot
            End If
            cboPointShape(0).ListIndex = .PointShape - 1
           
            lblPointColorSelection(0).BackColor = .PointAndLineColor
            
            txtGraphLineWidth = Trim(.LineWidthPixels)
        End With
    
        If .PointShapeSeries2 < 1 Or .PointShapeSeries2 > OlectraChart2D.ShapeConstants.oc2dShapeSquare Then
            .PointShapeSeries2 = OlectraChart2D.ShapeConstants.oc2dShapeDot
        End If
        cboPointShape(1).ListIndex = .PointShapeSeries2
        lblPointColorSelection(1).BackColor = .PointAndLineColorSeries2
        
        SetCheckBox chkClipOutlierValues, .ClipOutliers
        txtClipOutliersFactor = .ClipOutliersFactor
        
        ToggleWindowStayOnTop .KeepWindowOnTop
    End With
    mUpdatingControls = False
    
    ComputeAndDisplayChromatogram
    
    mFormInitialized = True
    Exit Sub

InitializeFormErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmTICandBPIPlots->InitializeForm", Err.Description, CallerID
    Resume Next

End Sub

Private Sub InitializePlot()
    
    Dim dblBlankDataX(1 To 1) As Double
    Dim dblBlankDataY(1 To 1) As Double
    
    With ctlChromatogramPlot
        .PopulateSymbolStyleComboBox cboPointShape(0)
        .PopulateSymbolStyleComboBox cboPointShape(1)
        
        .EnableDisableDelayUpdating True
        .SetCurrentGroup 2
        .SetSeriesCount 0

        .SetCurrentGroup 1
        .SetSeriesCount 1

        .SetSeriesDataPointCount 1, 1
        .SetDataX 1, dblBlankDataX()
        .SetDataY 1, dblBlankDataY()

        .EnableDisableDelayUpdating False
    End With

End Sub

Private Sub PopulateComboBoxes()
    
    mUpdatingControls = True
    With cboTICToDisplay
        .Clear
        .AddItem "TIC derived from data intensities (current data in view)"
        .AddItem "BPI (base peak intensity chromatogram, current data in view)"
        .AddItem "TIC of the Time Domain Signal"
        .AddItem "TIC derived from count of data points (current data in view)"
        .AddItem "TIC from raw data"
        .AddItem "BPI from raw data"
        .AddItem "Deisotoping Intensity Thresholds"
        .AddItem "Deisotoping Peak Counts"
        
        .ListIndex = tbcTICFromCurrentDataIntensities
    End With
    mUpdatingControls = False
    
End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    Static blnResizing As Boolean
    
    If blnResizing Then Exit Sub
    blnResizing = True
    
    With cboTICToDisplay
        .Top = 60
        .Left = 120
    End With
    
    With lblStatus
        .Top = 30
        .Left = cboTICToDisplay.Left + cboTICToDisplay.width + 120
    End With
    
    With chkAutoUpdatePlot
        .Top = 240
        .Left = lblStatus.Left
    End With
    
    With ctlChromatogramPlot
        .Left = 0
        .Top = cboTICToDisplay.Top + cboTICToDisplay.Height + 120
        lngDesiredValue = Me.ScaleWidth
        If lngDesiredValue < 3000 Then lngDesiredValue = 3000
        .width = lngDesiredValue

        lngDesiredValue = Me.ScaleHeight - .Top
        If lngDesiredValue < 2000 Then lngDesiredValue = 2000
        .Height = lngDesiredValue
    End With

    fraOptions.Left = ctlChromatogramPlot.Left
    fraOptions.Top = ctlChromatogramPlot.Top
    
    blnResizing = False
    
End Sub

Public Function SaveChartPictureToFile(blnSaveAsPNG As Boolean, Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    ' If blnSaveAsPNG = True, then saves a PNG file
    ' If blnSaveAsPNG = False, then saves a JPG file
    
    ' Returns 0 if success, the error code if an error
    
    Dim strPictureFormat As String
    Dim strPictureExtension As String
    Dim strChromatogramFileNameBase As String
    
    Dim objRemoteSaveFileHandler As clsRemoteSaveFileHandler
    Dim strWorkingFilePath As String
    Dim blnSuccess As Boolean
    
    strChromatogramFileNameBase = GetChromDescription(mCurrentPlotType, False)
    
On Error GoTo SaveChartPictureToFileErrorHandler

    If blnSaveAsPNG Then
        strPictureFormat = "PNG"
        strPictureExtension = ".png"
    Else
        strPictureFormat = "JPG"
        strPictureExtension = ".jpg"
    End If
    
    If Len(strFilePath) = 0 Then
        strFilePath = SelectFile(Me.hwnd, "Enter filename", "", True, strChromatogramFileNameBase & strPictureExtension, strPictureFormat & " Files (*." & strPictureExtension & ")|*." & strPictureExtension & "|All Files (*.*)|*.*")
    End If
    
    If Len(strFilePath) > 0 Then
        strFilePath = FileExtensionForce(strFilePath, strPictureExtension)
        Set objRemoteSaveFileHandler = New clsRemoteSaveFileHandler
        strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
            
        If ctlChromatogramPlot.SaveChartPictureToFile(blnSaveAsPNG, strWorkingFilePath) Then
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

Private Sub SelectCustomColor(lblThisLabel As Label)
    Dim lngTemporaryColor As Long
    
    lngTemporaryColor = lblThisLabel.BackColor
    Call GetColorAPIDlg(Me.hwnd, lngTemporaryColor)
    If lngTemporaryColor >= 0 Then
        lblThisLabel.BackColor = lngTemporaryColor
        mNeedToUpdatePlot = True
    End If
End Sub

Public Sub SetPlotMode(eChromType As tbcTICAndBPIConstants)
    On Error Resume Next
    If eChromType >= 0 And eChromType < TIC_AND_BPI_TYPE_COUNT And eChromType < cboTICToDisplay.ListCount Then
        ' Note: This will automatically call UpdatePlot, unless the timer
        '  has been disabled using TogglePlotUpdateTimerEnabled
        cboTICToDisplay.ListIndex = eChromType
        
        If Not tmrUpdatePlot.Enabled Then
            UpdatePlot
        End If
    Else
        Debug.Assert False
    End If
End Sub

Private Sub ShowHideOptions(Optional blnForceHide As Boolean)
    If blnForceHide Then
        fraOptions.Visible = False
    Else
        fraOptions.Visible = Not fraOptions.Visible
    End If
    
    mnuViewOptions.Checked = fraOptions.Visible
End Sub

Public Sub SetNormalizeYAxisOption(blnEnableNormalizing As Boolean)
    If Not cChkBox(chkNormalizeYAxis) = blnEnableNormalizing Then
        SetCheckBox chkNormalizeYAxis, blnEnableNormalizing
    End If
    Debug.Assert glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis = blnEnableNormalizing
End Sub

Public Sub SetPlotNETOnXAxis(blnEnableNETOnXAxis As Boolean)
    SetCheckBox chkPlotNETOnXAxis, blnEnableNETOnXAxis
End Sub

Public Sub SetSmoothUsingMovingAverage(blnEnableSmoothing As Boolean)
    If Not cChkBox(chkSmoothUsingMovingAverage) = blnEnableSmoothing Then
        SetCheckBox chkSmoothUsingMovingAverage, blnEnableSmoothing
    End If
    Debug.Assert glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage = blnEnableSmoothing
End Sub

Public Sub TogglePlotUpdateTimerEnabled(blnEnableTimer As Boolean)
    ' This sub should only be called during auto-analysis
    ' If blnEnableTimer = false, then the timer is disabled, effectively preventing the
    '  plots from updating when cboTICToDisplay is clicked (or changed programmatically)
    
    tmrUpdatePlot.Enabled = blnEnableTimer
    
End Sub

Private Sub ToggleWindowStayOnTop(blnEnableStayOnTop As Boolean)
    
    mnuViewKeepWindowOnTop.Checked = blnEnableStayOnTop
    glbPreferencesExpanded.TICAndBPIPlottingOptions.KeepWindowOnTop = blnEnableStayOnTop
    If mWindowStayOnTopEnabled = blnEnableStayOnTop Then Exit Sub
    
    Me.ScaleMode = vbTwips
    
    WindowStayOnTop Me.hwnd, blnEnableStayOnTop, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)
    
    mWindowStayOnTopEnabled = blnEnableStayOnTop

End Sub

Private Sub UpdatePlot()
    Dim strPlotTitle As String
    Dim strXAxisTitle As String
    Dim strYAxisTitle As String
    Dim strFormatString As String
    Dim blnNoteAdded As Double
    
    Dim lngXMin As Long, lngXMax As Long
    
    Dim intSeries As Integer
    Dim intSeriesCount As Integer
    
    Dim eSymbolShape As OlectraChart2D.ShapeConstants
    Dim lngLineColor As Long
    
    Dim udtGraphOptions As udtGraph2DOptionsType
    
    If mUpdatingControls Then Exit Sub
    mNeedToUpdatePlot = False
    
    UpdateStatus "Updating plot"

On Error GoTo GraphChromatogramErrorHandler
    
    If CallerID < 1 Or CallerID > UBound(GelData()) Then Exit Sub
    If GelStatus(CallerID).Deleted Then Exit Sub
    
    ' Olectra Chart requires that the data arrays be 1-based

    strPlotTitle = ""
    If Not GelAnalysis(CallerID) Is Nothing Then
        If GelAnalysis(CallerID).Job > 0 Then
            strPlotTitle = "Job " & Trim(GelAnalysis(CallerID).Job) & ": "
        End If
    End If
    strPlotTitle = strPlotTitle & StripFullPath(ExtractInputFilePath(CallerID))
    
    ' Determine the plot type and thus the series count
    If cboTICToDisplay.ListIndex >= 0 And cboTICToDisplay.ListIndex < TIC_AND_BPI_TYPE_COUNT Then
        mCurrentPlotType = cboTICToDisplay.ListIndex
    Else
        ' This shouldn't happen
        Debug.Assert False
        mCurrentPlotType = tbcBPIFromCurrentDataIntensities
    End If
    
    If mChromData.Chromatograms(mCurrentPlotType).DualSeriesData Then
        intSeriesCount = 2
    Else
        intSeriesCount = 1
    End If
    
    With ctlChromatogramPlot
        ' Delay updating the chart
        .EnableDisableDelayUpdating True

        .SetLabelGraphTitle strPlotTitle

        .SetChartType oc2dTypePlot, 1
        .SetCurrentGroup 1
        .SetSeriesCount intSeriesCount

        ' Copying to local variable to make code cleaner
        udtGraphOptions = glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions
    
        For intSeries = 1 To intSeriesCount
            .SetCurrentSeries intSeries
            
            If intSeries = 1 Then
                eSymbolShape = val(udtGraphOptions.PointShape)
                lngLineColor = udtGraphOptions.PointAndLineColor
            Else
                eSymbolShape = val(glbPreferencesExpanded.TICAndBPIPlottingOptions.PointShapeSeries2)
                lngLineColor = glbPreferencesExpanded.TICAndBPIPlottingOptions.PointAndLineColorSeries2
            End If
            
            If udtGraphOptions.ShowPointSymbols Then
                .SetStyleDataSymbol lngLineColor, eSymbolShape, udtGraphOptions.PointSizePixels
            Else
                .SetStyleDataSymbol lngLineColor, OlectraChart2D.ShapeConstants.oc2dShapeNone, 5
            End If
    
            If udtGraphOptions.DrawLinesBetweenPoints Then
                .SetStyleDataLine lngLineColor, oc2dLineSolid, udtGraphOptions.LineWidthPixels
            Else
                .SetStyleDataLine lngLineColor, oc2dLineNone, 1
            End If
    
            .SetStyleDataFill lngLineColor, oc2dFillSolid
        Next intSeries
        
        .SetCurrentSeries 1

        ' Plot the data
        If intSeriesCount = 2 Then
            UpdatePlotAddData mChromData.Chromatograms(mCurrentPlotType).NormalizedIntensity(), mChromData.ScanCount, 1, 2
            UpdatePlotAddData mChromData.Chromatograms(mCurrentPlotType).RawIntensity(), mChromData.ScanCount, 2, 2
        Else
            If glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis Then
                UpdatePlotAddData mChromData.Chromatograms(mCurrentPlotType).NormalizedIntensity(), mChromData.ScanCount, 1, 1
            Else
                UpdatePlotAddData mChromData.Chromatograms(mCurrentPlotType).RawIntensity(), mChromData.ScanCount, 1, 1
            End If
        End If
        
        strYAxisTitle = GetChromDescription(mCurrentPlotType, True)
        If glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis Then
            
            If mChromData.Chromatograms(mCurrentPlotType).MaximumValue >= 100000 Then
                strFormatString = "0.000E+00"
            ElseIf mChromData.Chromatograms(mCurrentPlotType).MaximumValue >= 100 Then
                strFormatString = "#0.0"
            Else
                strFormatString = "#0.000"
            End If
            strYAxisTitle = strYAxisTitle & " (Maximum = " & Format(mChromData.Chromatograms(mCurrentPlotType).MaximumValue, strFormatString)
            blnNoteAdded = True
        End If
        
        If mChromData.Chromatograms(mCurrentPlotType).SmoothWindowWidth > 0 Then
            If blnNoteAdded Then
                strYAxisTitle = strYAxisTitle & ", "
            Else
                strYAxisTitle = strYAxisTitle & " ("
            End If
            strYAxisTitle = strYAxisTitle & "Smooth window width = " & Trim(mChromData.Chromatograms(mCurrentPlotType).SmoothWindowWidth)
            blnNoteAdded = True
        End If
        
        If blnNoteAdded Then strYAxisTitle = strYAxisTitle & ")"

        If glbPreferencesExpanded.TICAndBPIPlottingOptions.PlotNETOnXAxis Then
            strXAxisTitle = "Normalized Elution Time (NET)"
        Else
            strXAxisTitle = "Scan Number"
        End If

        .SetLabelXAxis strXAxisTitle
        .SetLabelYAxis strYAxisTitle

        If udtGraphOptions.AutoScaleXAxis Then
            If mAutoUpdatePlot Then
                ' Determine the actual scan range the user has zoomed in on
                GelBody(CallerID).GetCurrentZoomArea lngXMin, lngXMax, 0, 0
                
                If glbPreferencesExpanded.TICAndBPIPlottingOptions.PlotNETOnXAxis Then
                    ' Need to convert Scan numbers to NET values
                    .SetXRange ScanToGANET(CallerID, lngXMin), ScanToGANET(CallerID, lngXMax)
                Else
                    .SetXRange CDbl(lngXMin), CDbl(lngXMax)
                End If
            Else
                .AutoScaleXNow
            End If
        End If

        ' Set the Tick Spacing the default
        .SetXAxisTickSpacing 1, True

        .SetXAxisAnnotationMethod oc2dAnnotateValues
        .SetXAxisAnnotationPlacement oc2dAnnotateAuto
        
        .SetYAxisAnnotationMethod oc2dAnnotateValues
        .SetYAxisAnnotationPlacement oc2dAnnotateAuto

        If udtGraphOptions.ShowGridLines Then
            .SetYAxisGridlines oc2dLineDotted
        Else
            .SetYAxisGridlines oc2dLineNone
        End If

        ' Restore the chart to update
        .EnableDisableDelayUpdating False

    End With

    UpdateStatus "Ready"
    Exit Sub
    
GraphChromatogramErrorHandler:
    Debug.Print "Error in UpdatePlot: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmTICAndBPIPlots->UpdatePlot"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error while populating graph: " & vbCrLf & Err.Description, vbInformation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub UpdatePlotAddData(dblChromatogramDataOneBased() As Double, lngDataCount As Long, intSeries As Integer, intSeriesCount As Integer)
    Dim lngIndex As Long
    Dim dblXData() As Double
    
    With ctlChromatogramPlot
        .SetSeriesDataPointCount intSeries, lngDataCount
        
        If lngDataCount > 0 Then
    
            If glbPreferencesExpanded.TICAndBPIPlottingOptions.PlotNETOnXAxis Then
                .SetDataX intSeries, mChromData.GANETVals()
            Else
                ReDim dblXData(1 To lngDataCount)
                
                For lngIndex = 1 To lngDataCount
                    dblXData(lngIndex) = mChromData.ScanNumberStart + lngIndex - 1
                Next lngIndex
                
                .SetDataX intSeries, dblXData()
            End If
            
            .SetDataY intSeries, dblChromatogramDataOneBased()
        End If
        
    End With
    
End Sub

Private Sub UpdateStatus(strMessage As String)
    lblStatus = strMessage
    DoEvents
End Sub

Private Sub cboPointShape_Click(Index As Integer)
    If mFormInitialized Then
        If Index = 0 Then
            glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.PointShape = cboPointShape(Index).ListIndex + 1
        Else
            glbPreferencesExpanded.TICAndBPIPlottingOptions.PointShapeSeries2 = cboPointShape(Index).ListIndex + 1
        End If
    End If
    mNeedToUpdatePlot = True
End Sub

Private Sub cboTICToDisplay_Click()
    mNeedToUpdatePlot = True
End Sub

Private Sub chkAutoScaleXRange_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.AutoScaleXAxis = cChkBox(chkAutoScaleXRange)
    mNeedToUpdatePlot = True
End Sub

Private Sub chkAutoUpdatePlot_Click()
    mAutoUpdatePlot = cChkBox(chkAutoUpdatePlot)
    mNeedToRecomputeChromatogram = True
End Sub

Private Sub chkClipOutlierValues_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.ClipOutliers = cChkBox(chkClipOutlierValues)
    mNeedToRecomputeChromatogram = True
End Sub

Private Sub chkDrawLinesBetweenPoints_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.DrawLinesBetweenPoints = cChkBox(chkDrawLinesBetweenPoints)
    mNeedToUpdatePlot = True
End Sub

Private Sub chkNormalizeYAxis_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis = cChkBox(chkNormalizeYAxis)
    mNeedToUpdatePlot = True
End Sub

Private Sub chkPlotNETOnXAxis_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.PlotNETOnXAxis = cChkBox(chkPlotNETOnXAxis)
    mNeedToUpdatePlot = True
End Sub

Private Sub chkShowGridlines_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.ShowGridLines = cChkBox(chkShowGridlines)
    mNeedToUpdatePlot = True
End Sub

Private Sub chkShowPointSymbols_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.ShowPointSymbols = cChkBox(chkShowPointSymbols)
    mNeedToUpdatePlot = True
End Sub

Private Sub chkSmoothUsingMovingAverage_Click()
    glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage = cChkBox(chkSmoothUsingMovingAverage)
    mNeedToRecomputeChromatogram = True
End Sub

Private Sub Form_Activate()
    If Not mFormInitialized Then InitializeForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift Or vbCtrlMask Then
        If KeyCode = vbKeyN Then
            SetCheckBox chkNormalizeYAxis, Not cChkBox(chkNormalizeYAxis)
        ElseIf KeyCode = vbKeyM Then
            SetCheckBox chkSmoothUsingMovingAverage, Not cChkBox(chkSmoothUsingMovingAverage)
        End If
    End If
End Sub

Private Sub Form_Load()
    
    SizeAndCenterWindow Me, cWindowLowerThird, 10500, 7500, False
    
    Me.ScaleMode = vbTwips
    mFormInitialized = False
    
    InitializePlot
    
    PositionControls
    
    ShowHideOptions True
    
    PopulateComboBoxes
    
    tmrUpdatePlot.Enabled = True
    tmrUpdatePlot.Interval = 200
    
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub lblPointColorSelection_Click(Index As Integer)
    SelectCustomColor lblPointColorSelection(Index)
    If Index = 0 Then
        glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.PointAndLineColor = lblPointColorSelection(Index).BackColor
    Else
        glbPreferencesExpanded.TICAndBPIPlottingOptions.PointAndLineColorSeries2 = lblPointColorSelection(Index).BackColor
    End If
End Sub

Private Sub mnuCopyAll_Click()
    ExportChromatogramDataToClipboardOrFile "", True, True
End Sub

Private Sub mnuCopyChart_Click(Index As Integer)
    Select Case Index
    Case ccmWMF
        ctlChromatogramPlot.CopyToClipboard oc2dFormatMetafile
    Case ccmEMF
        ctlChromatogramPlot.CopyToClipboard oc2dFormatEnhMetafile
    Case Else
        ' Includes ccmBMP
        ctlChromatogramPlot.CopyToClipboard oc2dFormatBitmap
    End Select
End Sub

Private Sub mnuCopyChromatogram_Click()
    ExportChromatogramDataToClipboardOrFile "", False, True
End Sub

Private Sub mnuExit_Click()
    Unload Me
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
    
    strFilePath = SelectFile(Me.hwnd, "Enter filename", "", True, "TICAndBPIPlots.txt", "Text Files (*.txt)|*.txt|All Files (*.*)|*.*")

    If Len(strFilePath) > 0 Then
        ExportChromatogramDataToClipboardOrFile strFilePath, True
    End If
End Sub

Private Sub mnuViewKeepWindowOnTop_Click()
    ToggleWindowStayOnTop Not mWindowStayOnTopEnabled
End Sub

Private Sub mnuViewOptions_Click()
    ShowHideOptions
End Sub

Private Sub tmrUpdatePlot_Timer()
    If mNeedToRecomputeChromatogram Then
        mNeedToRecomputeChromatogram = False
        ComputeAndDisplayChromatogram
    ElseIf mNeedToUpdatePlot Then
        mNeedToUpdatePlot = False
        UpdatePlot
    End If
End Sub

Private Sub txtClipOutliersFactor_LostFocus()
    If IsNumeric(txtClipOutliersFactor) Then
        With glbPreferencesExpanded.TICAndBPIPlottingOptions
            If .ClipOutliersFactor <> val(txtClipOutliersFactor) Then
                .ClipOutliersFactor = val(txtClipOutliersFactor)
                If cChkBox(chkClipOutlierValues) Then mNeedToRecomputeChromatogram = True
            End If
        End With
    End If
End Sub

Private Sub txtGraphLineWidth_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphLineWidth, KeyAscii, True, False
End Sub

Private Sub txtGraphLineWidth_LostFocus()
    ValidateTextboxValueLng txtGraphLineWidth, 1, 20, 3
    glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.LineWidthPixels = CLngSafe(txtGraphLineWidth)
    mNeedToUpdatePlot = True
End Sub

Private Sub txtGraphPointSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphPointSize, KeyAscii, True, False
End Sub

Private Sub txtGraphPointSize_LostFocus()
    ValidateTextboxValueLng txtGraphPointSize, 1, 20, 2
    glbPreferencesExpanded.TICAndBPIPlottingOptions.Graph2DOptions.PointSizePixels = CLngSafe(txtGraphPointSize)
    mNeedToUpdatePlot = True
End Sub

Private Sub txtSmoothUsingMovingAverage_Change()
    ValidateTextboxValueLng txtSmoothUsingMovingAverage, 3, 1001, 3
    With glbPreferencesExpanded.TICAndBPIPlottingOptions
        .MovingAverageWindowWidth = CLngSafe(txtSmoothUsingMovingAverage)
        
        ' Make sure .MovingAverageWindowWidth is odd
        If .MovingAverageWindowWidth Mod 2 = 0 Then
            ' Number is even; add one to the number
            .MovingAverageWindowWidth = .MovingAverageWindowWidth + 1
        End If
    End With
    mNeedToRecomputeChromatogram = True
End Sub

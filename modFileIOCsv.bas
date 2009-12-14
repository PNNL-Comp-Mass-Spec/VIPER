Attribute VB_Name = "modFileIOCsv"
Option Explicit

Public Const CSV_ISOS_IC_FILE_SUFFIX As String = "isos_ic.csv"
Public Const CSV_ISOS_FILE_SUFFIX As String = "isos.csv"
Public Const CSV_ISOS_PAIRS_SUFFIX As String = "pairs_isos.csv"
Public Const CSV_SCANS_FILE_SUFFIX As String = "scans.csv"

Public Const LCMS_FEATURES_FILE_SUFFIX As String = "LCMSFeatures.txt"
Public Const LCMS_FEATURE_TO_PEAK_MAP_FILE_SUFFIX As String = "LCMSFeatureToPeakMap.txt"

Public Const CSV_COLUMN_HEADER_UNKNOWN_WARNING As String = "Warning: unknown column headers"
Public Const CSV_COLUMN_HEADER_MISSING_WARNING As String = "Warning: expected important column headers"

' Note: These should all be lowercase string values
Private Const SCANS_COLUMN_SCAN_NUM As String = "scan_num"
Private Const SCANS_COLUMN_FRAME_NUM As String = "frame_num"        ' Represents IMS Frame Number; VIPER treats this as scan_num
Private Const SCANS_COLUMN_TIME_A As String = "time"
Private Const SCANS_COLUMN_TIME_B As String = "scan_time"           ' Represents the elution time of the given scan (or IMS frame); used in non-IMS data and also used in the 2008 version of the IMS _scans.csv file format
Private Const SCANS_COLUMN_FRAME_TIME As String = "frame_time"      ' Represents the elution time of the IMS Frame Number; VIPER treats this as scan_time
Private Const SCANS_COLUMN_DRIFT_TIME As String = "drift_time"      ' Old column that was only used in the 2008 version of the IMS file format
Private Const SCANS_COLUMN_TYPE As String = "type"
Private Const SCANS_COLUMN_NUM_DEISOTOPED As String = "num_deisotoped"
Private Const SCANS_COLUMN_NUM_PEAKS As String = "num_peaks"
Private Const SCANS_COLUMN_TIC As String = "tic"
Private Const SCANS_COLUMN_BPI_MZ As String = "bpi_mz"
Private Const SCANS_COLUMN_BPI As String = "bpi"
Private Const SCANS_COLUMN_TIME_DOMAIN_SIGNAL As String = "time_domain_signal"
Private Const SCANS_COLUMN_PEAK_INTENSITY_THRESHOLD As String = "peak_intensity_threshold"
Private Const SCANS_COLUMN_PEPTIDE_INTENSITY_THRESHOLD As String = "peptide_intensity_threshold"
Private Const SCANS_COLUMN_IMS_FRAME_PRESSURE As String = "frame_pressure"
Private Const SCANS_COLUMN_IMS_FRAME_PRESSURE_FRONT As String = "frame_pressure_front"
Private Const SCANS_COLUMN_IMS_FRAME_PRESSURE_BACK As String = "frame_pressure_back"

' Note: These should all be lowercase string values
Private Const ISOS_COLUMN_SCAN_NUM_A As String = "scan_num"
Private Const ISOS_COLUMN_SCAN_NUM_B As String = "lc_scan_num"  ' Represented Frame Number in the 2008 version of the IMS File format
Private Const ISOS_COLUMN_FRAME_NUM As String = "frame_num"     ' Represents MS Frame Number; VIPER treats this as scan_num
Private Const ISOS_COLUMN_IMS_SCAN_NUM As String = "ims_scan_num"
Private Const ISOS_COLUMN_CHARGE As String = "charge"
Private Const ISOS_COLUMN_ABUNDANCE As String = "abundance"
Private Const ISOS_COLUMN_MZ As String = "mz"
Private Const ISOS_COLUMN_FIT As String = "fit"
Private Const ISOS_COLUMN_AVERAGE_MW As String = "average_mw"
Private Const ISOS_COLUMN_MONOISOTOPIC_MW As String = "monoisotopic_mw"
Private Const ISOS_COLUMN_MOSTABUNDANT_MW As String = "mostabundant_mw"
Private Const ISOS_COLUMN_FWHM As String = "fwhm"
Private Const ISOS_COLUMN_SIGNAL_NOISE As String = "signal_noise"
Private Const ISOS_COLUMN_MONO_ABUNDANCE As String = "mono_abundance"
Private Const ISOS_COLUMN_MONO_PLUS2_ABUNDANCE As String = "mono_plus2_abundance"
Private Const ISOS_COLUMN_MONO_PLUS4_ABUNDANCE As String = "mono_plus4_abundance"
Private Const ISOS_COLUMN_MONO_MINUS4_ABUNDANCE As String = "mono_minus4_abundance"
Private Const ISOS_COLUMN_IMS_DRIFT_TIME As String = "drift_time"
Private Const ISOS_COLUMN_IMS_CUMULATIVE_DRIFT_TIME As String = "cumulative_drift_time"

Private Const SCAN_INFO_DIM_CHUNK As Long = 10000
Private Const ISO_DATA_DIM_CHUNK As Long = 25000

Private Const SCAN_FILE_COLUMN_COUNT As Integer = 13
Private Enum ScanFileColumnConstants
    ScanNumber = 0
    ScanTime = 1
    ScanType = 2
    NumDeisotoped = 3
    NumPeaks = 4
    TIC = 5
    BPImz = 6
    BPI = 7
    TimeDomainSignal = 8
    PeakIntensityThreshold = 9
    PeptideIntensityThreshold = 10
    IMSFramePressureFront = 11               ' Only present in IMS datafiles
    IMSFramePressureBack = 12               ' Only present in IMS datafiles
End Enum

Private Const ISOS_FILE_COLUMN_COUNT As Integer = 15
Private Enum IsosFileColumnConstants
    ScanNumber = 0
    Charge = 1
    Abundance = 2
    MZ = 3
    Fit = 4
    AverageMW = 5
    MonoisotopicMW = 6
    MostAbundantMW = 7
    FWHM = 8
    SignalToNoise = 9
    MonoAbundance = 10
    MonoPlus2Abundance = 11
    MonoPlus4Abundance = 12
    MonoMinus4Abundance = 13
    IMSDriftTime = 14               ' Only present in IMS datafiles
End Enum

Private Enum rmReadModeConstants
    rmPrescanData = 0
    rmStoreData = 1
    rmReadComplete = 2
End Enum

Private mGelIndex As Long
Private mScanInfoCount As Long

Private mEvenOddScanFilter As Boolean
Private mEvenOddModCompareVal As Integer

Private mMaxFit As Double
Private mFilterByAbundance As Boolean
Private mAbundanceMin As Double
Private mAbundanceMax As Double

Private mMaximumDataCountEnabled As Boolean
Private mMaximumDataCountToLoad As Long

Private mTotalIntensityPercentageFilterEnabled As Boolean
Private mTotalIntensityPercentageFilter As Single

Private mPrescannedData As clsFileIOPrescannedData

Private mPointsToKeep As Dictionary
Private mHashMapOfPointsKept As Dictionary
Private mFeatureToScanMap As Dictionary

Private mValidDataPointCount As Long
Private mSubtaskMessage As String

Private mReadMode As rmReadModeConstants
Private mCurrentProgressStep As Integer

Private Sub DuplicateIsoLineDataPoint(ByRef udtSrcIsoData As udtIsotopicDataType, ByRef udtTargetIsoData() As udtIsotopicDataType, ByVal lngTargetIndex As Long, ByVal dblTargetMassDelta As Double, ByVal sngTargetIntensity As Single, ByVal eIReportTagType As irtIReportTagTypeConstants)

    Static intErrorLogCount As Integer
    
    On Error GoTo DuplicateIsoLineDataPointErrorHandler
    
    If lngTargetIndex > UBound(udtTargetIsoData) Then
        ' Increase the amount of space reserved for udtTargetIsoData by 50%
        ' Note that udtTargetIsoData() is a 1-based array
        ReDim Preserve udtTargetIsoData((UBound(udtTargetIsoData)) * 1.5)
    End If
    
    udtTargetIsoData(lngTargetIndex) = udtSrcIsoData
    
    With udtTargetIsoData(lngTargetIndex)
        .MonoisotopicMW = udtSrcIsoData.MonoisotopicMW + dblTargetMassDelta
        .AverageMW = .MonoisotopicMW
        .MostAbundantMW = .MonoisotopicMW
        
        .IntensityMono = sngTargetIntensity
        .Abundance = .IntensityMono
        
        .MZ = ConvoluteMass(.MonoisotopicMW, 0, .Charge)
        .IntensityMonoPlus2 = 0
        .IsotopeLabel = iltIsotopeLabelTagConstants.iltNone
        .IReportTagType = eIReportTagType
        .IMSDriftTime = udtSrcIsoData.IMSDriftTime
    End With
    
    Exit Sub

DuplicateIsoLineDataPointErrorHandler:
    Debug.Assert False
    If intErrorLogCount < 10 Then
        LogErrors Err.Number, "DuplicateIsoLineDataPoint"
        intErrorLogCount = intErrorLogCount + 1
    End If

End Sub

Private Function GetColumnValueDbl(ByRef strData() As String, ByVal intColumnIndex As Integer, Optional ByVal dblDefaultValue As Double = 0) As Double
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        GetColumnValueDbl = CDbl(strData(intColumnIndex))
    Else
        GetColumnValueDbl = dblDefaultValue
    End If
    
    Exit Function
    
GetColumnValueErrorHandler:
    Debug.Assert False
    GetColumnValueDbl = 0
    
End Function

Private Function GetColumnValueLng(ByRef strData() As String, ByVal intColumnIndex As Integer, Optional ByVal lngDefaultValue As Long = 0) As Long
    
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        GetColumnValueLng = CLng(strData(intColumnIndex))
    Else
        GetColumnValueLng = lngDefaultValue
    End If
    
    Exit Function
    
GetColumnValueErrorHandler:
    Debug.Assert False
    GetColumnValueLng = 0
    
End Function

Private Function GetColumnValueSng(ByRef strData() As String, ByVal intColumnIndex As Integer, Optional ByVal sngDefaultValue As Single = 0) As Single
    Dim dblValue As Double
    
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        dblValue = CDbl(strData(intColumnIndex))
        If dblValue > 1E+38 Then
            dblValue = 1E+38
        ElseIf dblValue < -1E+38 Then
            dblValue = -1E+38
        End If
        GetColumnValueSng = CSng(dblValue)
    Else
        GetColumnValueSng = sngDefaultValue
    End If
    
    Exit Function
    
GetColumnValueErrorHandler:
    Debug.Assert False
    GetColumnValueSng = 0
    
End Function

Public Function GetDatasetNameFromDecon2LSFilename(ByVal strFilePath As String) As String
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim strBase As String
    
    strBase = ""

    strFileName = fso.GetFileName(strFilePath)
    
    If StringEndsWith(strFileName, "_" & CSV_ISOS_IC_FILE_SUFFIX) Then
        strBase = StringTrimEnd(strFileName, "_" & CSV_ISOS_IC_FILE_SUFFIX)
    ElseIf StringEndsWith(strFileName, "_" & CSV_ISOS_FILE_SUFFIX) Then
        strBase = StringTrimEnd(strFileName, "_" & CSV_ISOS_FILE_SUFFIX)
    ElseIf StringEndsWith(strFileName, "_" & CSV_ISOS_PAIRS_SUFFIX) Then
        strBase = StringTrimEnd(strFileName, "_" & CSV_ISOS_PAIRS_SUFFIX)
    ElseIf StringEndsWith(strFileName, "_" & CSV_SCANS_FILE_SUFFIX) Then
        strBase = StringTrimEnd(strFileName, "_" & CSV_SCANS_FILE_SUFFIX)
    Else
        strBase = fso.GetBaseName(strFileName)
    End If
    
    Set fso = Nothing
    
    GetDatasetNameFromDecon2LSFilename = strBase
    
End Function

Private Function StringEndsWith(ByVal strText As String, ByVal strComparisonText As String) As Boolean
  
    Dim blnMatchFound As Boolean
  
    blnMatchFound = False
    If Len(strText) >= Len(strComparisonText) Then
        If LCase(Right(strText, Len(strComparisonText))) = LCase(strComparisonText) Then
            blnMatchFound = True
        End If
    End If
    
    StringEndsWith = blnMatchFound
    
End Function

Private Function StringTrimEnd(ByVal strText As String, ByVal strTextToTrim As String) As String
    Dim intTrimLength As Integer
    Dim strTrimmedText As String
    
    intTrimLength = Len(strTextToTrim)
    
    strTrimmedText = strText
    If Len(strTrimmedText) >= intTrimLength Then
        If LCase(Right(strTrimmedText, intTrimLength)) = LCase(strTextToTrim) Then
            strTrimmedText = Left(strTrimmedText, Len(strTrimmedText) - intTrimLength)
        End If
    End If
    
    StringTrimEnd = strTrimmedText
End Function

Private Function GetDefaultIsosColumnHeaders(blnRequiredColumnsOnly As Boolean, blnIncludeIMSFileHeaders As Boolean) As String
    Dim strHeaders As String
    
    If blnIncludeIMSFileHeaders Then
        strHeaders = ISOS_COLUMN_FRAME_NUM & ", " & ISOS_COLUMN_IMS_SCAN_NUM
    Else
        strHeaders = ISOS_COLUMN_SCAN_NUM_A
    End If
    
    If Not blnRequiredColumnsOnly Then
        strHeaders = strHeaders & ", " & ISOS_COLUMN_CHARGE
    End If
    
    strHeaders = strHeaders & ", " & ISOS_COLUMN_ABUNDANCE
    If Not blnRequiredColumnsOnly Then
        strHeaders = strHeaders & ", " & ISOS_COLUMN_MZ
        strHeaders = strHeaders & ", " & ISOS_COLUMN_FIT
        strHeaders = strHeaders & ", " & ISOS_COLUMN_AVERAGE_MW
    End If
    
    strHeaders = strHeaders & ", " & ISOS_COLUMN_MONOISOTOPIC_MW
    If Not blnRequiredColumnsOnly Then
        strHeaders = strHeaders & ", " & ISOS_COLUMN_MOSTABUNDANT_MW
        strHeaders = strHeaders & ", " & ISOS_COLUMN_FWHM
        strHeaders = strHeaders & ", " & ISOS_COLUMN_SIGNAL_NOISE
        strHeaders = strHeaders & ", " & ISOS_COLUMN_MONO_ABUNDANCE
        strHeaders = strHeaders & ", " & ISOS_COLUMN_MONO_PLUS2_ABUNDANCE
        strHeaders = strHeaders & ", " & ISOS_COLUMN_MONO_PLUS4_ABUNDANCE
        strHeaders = strHeaders & ", " & ISOS_COLUMN_MONO_MINUS4_ABUNDANCE
        
        If blnIncludeIMSFileHeaders Then
            strHeaders = strHeaders & ", " & ISOS_COLUMN_IMS_DRIFT_TIME
            strHeaders = strHeaders & ", " & ISOS_COLUMN_IMS_CUMULATIVE_DRIFT_TIME
        End If
    End If

    GetDefaultIsosColumnHeaders = strHeaders
End Function

Private Function GetDefaultScansColumnHeaders(blnRequiredColumnsOnly As Boolean, blnIncludeIMSFileHeaders As Boolean) As String
    Dim strHeaders As String

    If blnIncludeIMSFileHeaders Then
        strHeaders = SCANS_COLUMN_FRAME_NUM & ", " & SCANS_COLUMN_FRAME_TIME
    Else
        strHeaders = SCANS_COLUMN_SCAN_NUM & ", " & SCANS_COLUMN_TIME_B
    End If
    
    strHeaders = strHeaders & ", " & SCANS_COLUMN_TYPE
    
    If Not blnRequiredColumnsOnly Then
        strHeaders = strHeaders & ", " & SCANS_COLUMN_NUM_DEISOTOPED
        strHeaders = strHeaders & ", " & SCANS_COLUMN_NUM_PEAKS
        strHeaders = strHeaders & ", " & SCANS_COLUMN_TIC
        strHeaders = strHeaders & ", " & SCANS_COLUMN_BPI_MZ
        strHeaders = strHeaders & ", " & SCANS_COLUMN_BPI
        strHeaders = strHeaders & ", " & SCANS_COLUMN_TIME_DOMAIN_SIGNAL
        strHeaders = strHeaders & ", " & SCANS_COLUMN_PEAK_INTENSITY_THRESHOLD
        strHeaders = strHeaders & ", " & SCANS_COLUMN_PEPTIDE_INTENSITY_THRESHOLD
        
        If blnIncludeIMSFileHeaders Then
            strHeaders = strHeaders & ", " & SCANS_COLUMN_IMS_FRAME_PRESSURE_FRONT & ", " & SCANS_COLUMN_IMS_FRAME_PRESSURE_BACK
        End If
    End If
    
    GetDefaultScansColumnHeaders = strHeaders
End Function

Public Function LoadNewCSV(ByVal CSVFilePath As String, ByVal lngGelIndex As Long, _
                           ByVal MaxFit As Double, _
                           ByVal blnFilterByAbundance As Boolean, _
                           ByVal dblMinAbu As Double, ByVal dblMaxAbu As Double, _
                           ByVal blnMaximumDataCountEnabled As Boolean, ByVal lngMaximumDataCountToLoad As Long, _
                           ByVal blnTotalIntensityPercentageFilterEnabled, ByVal sngTotalIntensityPercentageFilter, _
                           ByVal eScanFilterMode As eosEvenOddScanFilterModeConstants, _
                           ByVal eDataFilterMode As dfmCSandIsoDataFilterModeConstants, _
                           ByVal blnLoadPredefinedLCMSFeatures As Boolean, _
                           ByVal sngAutoMapDataPointsMassTolerancePPM As Single, _
                           ByRef strErrorMessage, _
                           ByVal plmPointsLoadMode As Integer) As Long
    '-------------------------------------------------------------------------
    'Returns 0 if data successfuly loaded, -2 if data set is too large,
    '-3 if problems with scan numbers, -4 if no data found, -5 if user cancels load,
    '-6 for file not found or invalid file
    '-7 for file problem that user was already notified about
    '-10 for any other error
    'call this function with huge MaxFit or MaxFit <= 0 to load all values
    'Set blnFilterByAbundance to True to use dblMinAbu and dblMaxAbu to filter the abundance values
    'eDataFilterMode is ignored by this function
    '-------------------------------------------------------------------------
    
    Dim intProgressCount As Integer
    Dim blnFilePrescanEnabled As Boolean
    
    Dim strScansFilePath As String
    Dim strIsosFilePath As String
    Dim strBaseFilePath As String
    
    Dim strLCMSFeaturesFilePath As String
    Dim strLCMSFeatureToPeakMapFilePath As String
    
    Dim eResponse As VbMsgBoxResult
    
    Dim fso As New FileSystemObject
    Dim tsInFile As TextStream
    Dim strLineIn As String
    
    Dim blnValidScansFile As Boolean
    Dim blnValidDataPoint As Boolean
    Dim blnSuccess As Boolean
    
    Dim lngReturnValue As Long
    
    Dim lngScansFileByteCount As Long
    Dim lngByteCountTotal As Long
    Dim lngTotalBytesRead As Long
    
    Dim objSplitUMCs As clsSplitUMCsByAbundance
    
On Error GoTo LoadNewCSVErrorHandler

    ' Update the filter variables
    mGelIndex = lngGelIndex
    mMaxFit = MaxFit
    mFilterByAbundance = blnFilterByAbundance
    mAbundanceMin = dblMinAbu
    mAbundanceMax = dblMaxAbu

    mMaximumDataCountEnabled = blnMaximumDataCountEnabled
    mMaximumDataCountToLoad = lngMaximumDataCountToLoad
    
    mTotalIntensityPercentageFilterEnabled = blnTotalIntensityPercentageFilterEnabled
    mTotalIntensityPercentageFilter = sngTotalIntensityPercentageFilter

    If mMaximumDataCountEnabled Or mTotalIntensityPercentageFilterEnabled Then
        intProgressCount = 5
        blnFilePrescanEnabled = True
    
        If mMaximumDataCountToLoad < 10 Then mMaximumDataCountToLoad = 10
        If sngTotalIntensityPercentageFilter < 1 Then sngTotalIntensityPercentageFilter = 1
        If sngTotalIntensityPercentageFilter > 100 Then sngTotalIntensityPercentageFilter = 100
    Else
        intProgressCount = 3
        blnFilePrescanEnabled = False
    End If

    mCurrentProgressStep = 0
    frmProgress.InitializeForm "Loading data file", mCurrentProgressStep, intProgressCount, False, True, True, MDIForm1
    lngReturnValue = -10
    
    ' Resolve the CSV FilePath given to the ScansFilePath and the IsosFilePath variables
    
    If Not ResolveCSVFilePaths(CSVFilePath, strScansFilePath, strIsosFilePath, strBaseFilePath) Then
        strErrorMessage = "Unable to resolve the given FilePath to the _isos.csv and _scans.csv files: " & vbCrLf & CSVFilePath
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
        End If
        LoadNewCSV = -7
        Exit Function
    End If
    
    ' Validate that the input file(s) exist
    If Not fso.FileExists(strIsosFilePath) Then
        strErrorMessage = "Decon2LS _isos.csv file not found: " & vbCrLf & strIsosFilePath
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
        End If
        LoadNewCSV = -7
        Exit Function
    End If
    
    If blnLoadPredefinedLCMSFeatures Then
    
        ' Define the path to the LCMSFeature file and the FeatureToPeakMap file
        strLCMSFeaturesFilePath = strBaseFilePath & LCMS_FEATURES_FILE_SUFFIX
        strLCMSFeatureToPeakMapFilePath = strBaseFilePath & LCMS_FEATURE_TO_PEAK_MAP_FILE_SUFFIX
        
        ' Make sure the LCMSFeatures file exists
        If Not fso.FileExists(strLCMSFeaturesFilePath) Then
            strErrorMessage = "The LCMSFeatures file does not exist: " & vbCrLf & strLCMSFeaturesFilePath
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
            End If
            LoadNewCSV = -7
            Exit Function
        End If
    
        ' See if the FeatureToPeakMap file exists; it's OK if it doesn't
        If Not fso.FileExists(strLCMSFeatureToPeakMapFilePath) Then
            strErrorMessage = "Warning: The LCMS Feature to Peak Map file does not exist: " & vbCrLf & strLCMSFeatureToPeakMapFilePath & vbCrLf & "VIPER will infer the feature to peak mapping by looking for peaks within the scan range of each LC-MS feature and within " & sngAutoMapDataPointsMassTolerancePPM & " ppm of the feature's mass"
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
            End If
        End If
    
    End If
    
    blnValidScansFile = True
    If Not fso.FileExists(strScansFilePath) Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            eResponse = MsgBox("CSV Scans file not found: " & vbCrLf & strScansFilePath & vbCrLf & "Load the Isos.csv file anyway?  If yes, then scan type will be assumed to be MS and scan time will be unknown.  Choose No or Cancel to abort.", vbExclamation + vbYesNoCancel + vbDefaultButton3, glFGTU)
        Else
            AddToAnalysisHistory mGelIndex, "Error: CSV Scans file not found: " & strScansFilePath
            eResponse = vbCancel
        End If
        
        If eResponse = vbCancel Or eResponse = vbNo Then
            LoadNewCSV = -7
            Exit Function
        Else
            blnValidScansFile = False
        End If
    End If
    
    ' Initialize the even/odd scan filter variables
    mEvenOddScanFilter = False
    
    If Not blnLoadPredefinedLCMSFeatures Then
        If eScanFilterMode = eosLoadOddScansOnly Then
            mEvenOddScanFilter = True
            mEvenOddModCompareVal = 1                     ' Use scans where Scan Mod 2 = 1
        ElseIf eScanFilterMode = eosLoadEvenScansOnly Then
            mEvenOddScanFilter = True
            mEvenOddModCompareVal = 0                     ' Use scans where Scan Mod 2 = 0
        End If
    End If
    
    ' Initialize the progress bar
    lngTotalBytesRead = 0
    lngByteCountTotal = FileLen(strIsosFilePath)
    
    If blnValidScansFile Then
        lngScansFileByteCount = FileLen(strScansFilePath)
        lngByteCountTotal = lngByteCountTotal + lngScansFileByteCount
    End If
    
    frmProgress.InitializeSubtask "Reading data", 0, lngByteCountTotal
    
    mScanInfoCount = 0
    ReDim GelData(mGelIndex).ScanInfo(0)
    
    GelData(mGelIndex).DataStatusBits = 0
    
    If blnValidScansFile Then
        ' Read the scans file and populate .ScanInfo
        lngReturnValue = ReadCSVScanFile(fso, strScansFilePath, strBaseFilePath, lngTotalBytesRead)
    Else
        lngReturnValue = 0
    End If
    
    If plmPointsLoadMode >= plmLoadMappedPointsOnly Then
        Dim objReadLCMSFeatures As clsFileIOPredefinedLCMSFeatures
        Set objReadLCMSFeatures = New clsFileIOPredefinedLCMSFeatures
        lngReturnValue = objReadLCMSFeatures.CreatePeakListFromFeatureToPeakMapFile(strLCMSFeatureToPeakMapFilePath, mPointsToKeep)
        If plmPointsLoadMode = plmLoadOnePointPerLCMSFeature Then
            lngReturnValue = objReadLCMSFeatures.RefinePeakListFromFeatureFile(strLCMSFeaturesFilePath, mPointsToKeep, mFeatureToScanMap)
        End If
    End If
    
    If lngReturnValue = 0 Then
        ' Read the Isos file
        ' Note that the CSV Isos file only contains isotopic data, not charge state data
        lngReturnValue = ReadCSVIsosFile(fso, strIsosFilePath, strBaseFilePath, _
                                         lngScansFileByteCount, lngByteCountTotal, lngTotalBytesRead, _
                                         blnValidScansFile, blnFilePrescanEnabled, _
                                         blnLoadPredefinedLCMSFeatures, plmPointsLoadMode)
    
        If lngReturnValue = 0 Then
            If blnLoadPredefinedLCMSFeatures Then
                lngReturnValue = ReadLCMSFeatureFiles(fso, strLCMSFeaturesFilePath, strLCMSFeatureToPeakMapFilePath, sngAutoMapDataPointsMassTolerancePPM, plmPointsLoadMode)
            End If
        End If
    End If
    
    If glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCsByAbundance And plmPointsLoadMode < plmLoadOnePointPerLCMSFeature Then
        Set objSplitUMCs = New clsSplitUMCsByAbundance
        blnSuccess = UpdateUMCStatArrays(mGelIndex, True, False, frmProgress)
        objSplitUMCs.ExamineUMCs mGelIndex, frmProgress, GelUMC(mGelIndex).def.OddEvenProcessingMode, True, False
        Set objSplitUMCs = Nothing
    End If
        
    LoadNewCSV = lngReturnValue
    Exit Function

LoadNewCSVErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "LoadNewCSV"
    
    strErrorMessage = "Error in LoadNewCSVErrorHandler: " & Err.Description
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    LoadNewCSV = lngReturnValue
    
End Function

Private Function ReadCSVIsosFile(ByRef fso As FileSystemObject, ByVal strIsosFilePath As String, ByVal strBaseFilePath As String, _
                                 ByVal lngScansFileByteCount As Long, ByVal lngByteCountTotal As Long, _
                                 ByRef lngTotalBytesRead As Long, ByVal blnValidScansFile As Boolean, _
                                 ByVal blnFilePrescanEnabled As Boolean, _
                                 ByVal blnLoadPredefinedLCMSFeatures As Boolean, _
                                 ByVal plmPointsLoadMode As Integer) As Long

    ' Returns 0 if no error, the error number if an error

    Dim lngIndex As Long
    Dim lngNewIsoDataCount As Long
    Dim lngDataCountUpdated As Long
    Dim lngDataPointsAdded As Long
    Dim lngReturnValue As Long
    
    Dim objFile As File
    Dim objFolder As Folder
    
    Dim blnMonoPlus2DataPresent As Boolean
    Dim blnMonoPlus4DataPresent As Boolean
    Dim MaxMZ As Double
    Dim intColumnMapping() As Integer
    
    Dim sngMonoPlus4Intensities() As Single
    Dim sngMonoMinus4Intensities() As Single
    
    Dim blnIgnoreAllFiltersAndLoadAllData As Boolean
    
On Error GoTo ReadCSVIsosFileErrorHandler

    ' If we're loading predefined LC/MS features, then we need to ignore all filters and load all of the data
    blnIgnoreAllFiltersAndLoadAllData = blnLoadPredefinedLCMSFeatures
    
    If blnIgnoreAllFiltersAndLoadAllData Then
        ' Make sure blnFilePrescanEnabled is false, since we're ignoring all filters and loading all of the data
        blnFilePrescanEnabled = False
    End If

    ReDim intColumnMapping(ISOS_FILE_COLUMN_COUNT - 1) As Integer
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To ISOS_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    If Len(strBaseFilePath) = 0 Then
        strBaseFilePath = fso.GetBaseName(strIsosFilePath)
    End If
        
    With GelData(mGelIndex)
        If Not blnValidScansFile Then
            mScanInfoCount = 0
            ReDim .ScanInfo(SCAN_INFO_DIM_CHUNK)
        End If
        
        .LinesRead = 0
        .DataLines = 0
        .CSLines = 0
        .IsoLines = 0
        
        ReDim .IsoData(ISO_DATA_DIM_CHUNK)
    End With

    If blnFilePrescanEnabled Then
        mReadMode = rmReadModeConstants.rmPrescanData
    Else
        mReadMode = rmReadModeConstants.rmStoreData
    End If
    
    Do While mReadMode < rmReadModeConstants.rmReadComplete
        If blnFilePrescanEnabled Then
            If mReadMode = rmPrescanData Then
                mCurrentProgressStep = 0
            Else
                mCurrentProgressStep = 2
            End If
        Else
            mCurrentProgressStep = 0
        End If
        
        frmProgress.UpdateProgressBar mCurrentProgressStep
        
        If mReadMode = rmReadModeConstants.rmPrescanData Then
            ' Initialize the prescanned data class
            Set mPrescannedData = New clsFileIOPrescannedData
            
            With mPrescannedData
                .MaximumDataCountEnabled = mMaximumDataCountEnabled
                .MaximumDataCountToLoad = mMaximumDataCountToLoad
                .TotalIntensityPercentageFilterEnabled = mTotalIntensityPercentageFilterEnabled
                .TotalIntensityPercentageFilter = mTotalIntensityPercentageFilter
            End With
            
            mSubtaskMessage = "Pre-scanning Isotopic CSV file to determine data to load"
        Else
            mSubtaskMessage = "Reading Isotopic CSV file"
        End If
        frmProgress.InitializeSubtask mSubtaskMessage, 0, lngByteCountTotal
        frmProgress.UpdateSubtaskProgressBar lngScansFileByteCount
        
        ' Reset the tracking variables
        mValidDataPointCount = 0
        lngTotalBytesRead = 0
    
        lngReturnValue = ReadCSVIsosFileWork(fso, strIsosFilePath, lngTotalBytesRead, _
                                             intColumnMapping, blnValidScansFile, _
                                             sngMonoPlus4Intensities, sngMonoMinus4Intensities, _
                                             blnFilePrescanEnabled, blnIgnoreAllFiltersAndLoadAllData, _
                                             plmPointsLoadMode)
        If lngReturnValue <> 0 Then
            ' Error occurred
            Debug.Assert False
            ReadCSVIsosFile = lngReturnValue
            Exit Function
        End If
        
        If mReadMode = rmReadModeConstants.rmPrescanData Then
            If KeyPressAbortProcess > 1 Then
                ' User Cancelled Load
                ReadCSVIsosFile = -5
                frmProgress.HideForm
                Exit Function
            End If
            
            mCurrentProgressStep = mCurrentProgressStep + 1
            frmProgress.UpdateProgressBar mCurrentProgressStep
        
            mPrescannedData.ParsePrescannedData
        End If
        
        mReadMode = mReadMode + 1
    Loop
    
    ' Check for no data loaded
    If GelData(mGelIndex).CSLines + GelData(mGelIndex).IsoLines = 0 Then
       ReadCSVIsosFile = -4
       frmProgress.HideForm
       Exit Function
    End If
    
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load; keep the data in memory, but write an entry to the analysis history
        AddToAnalysisHistory mGelIndex, "Warning: File only partially loaded since user cancelled the loading process"
        KeyPressAbortProcess = 1
    End If
    
    mCurrentProgressStep = mCurrentProgressStep + 1
    frmProgress.UpdateProgressBar mCurrentProgressStep
    frmProgress.InitializeSubtask "Initializing data structures", 0, 1

    ' If IMS data is present, then update .DataStatusBits
    If intColumnMapping(IsosFileColumnConstants.IMSDriftTime) >= 0 Then
        GelData(mGelIndex).DataStatusBits = GelData(mGelIndex).DataStatusBits Or GEL_DATA_STATUS_BIT_IMS_DATA
    End If

    ' Look for the presence of MonoPlus2 data
    blnMonoPlus2DataPresent = False
    If intColumnMapping(IsosFileColumnConstants.MonoPlus2Abundance) >= 0 Then
        For lngIndex = 1 To GelData(mGelIndex).IsoLines
            If GelData(mGelIndex).IsoData(lngIndex).IntensityMonoPlus2 > 0 Then
                blnMonoPlus2DataPresent = True
                Exit For
            End If
        Next lngIndex
    End If

    ' Look for the presence of MonoPlus4 and MonoMinus4 data
    blnMonoPlus4DataPresent = False
    If intColumnMapping(IsosFileColumnConstants.MonoPlus4Abundance) >= 0 Then
        For lngIndex = 1 To GelData(mGelIndex).IsoLines
            If sngMonoPlus4Intensities(lngIndex) > 0 Then
                blnMonoPlus4DataPresent = True
                Exit For
            End If
        Next lngIndex
    End If
  
    If blnMonoPlus4DataPresent Then
        If glbPreferencesExpanded.IReportAutoAddMonoPlus4AndMinus4Data Then
            Dim udtNewIsoData() As udtIsotopicDataType
            ReDim udtNewIsoData(GelData(mGelIndex).IsoLines * 1.5)
            
            lngNewIsoDataCount = 0
            For lngIndex = 1 To GelData(mGelIndex).IsoLines
                lngNewIsoDataCount = lngNewIsoDataCount + 1
                If lngNewIsoDataCount > UBound(udtNewIsoData) Then
                    ReDim Preserve udtNewIsoData((UBound(udtNewIsoData)) * 1.5)
                End If
                
                udtNewIsoData(lngNewIsoDataCount) = GelData(mGelIndex).IsoData(lngIndex)
                                
                If sngMonoPlus4Intensities(lngIndex) > 0 Then
                    lngNewIsoDataCount = lngNewIsoDataCount + 1
                    DuplicateIsoLineDataPoint GelData(mGelIndex).IsoData(lngIndex), udtNewIsoData, lngNewIsoDataCount, glO16O18_DELTA, sngMonoPlus4Intensities(lngIndex), irtIReportTagTypeConstants.irtMonoPlus4
                End If
            
                If sngMonoMinus4Intensities(lngIndex) > 0 Then
                    lngNewIsoDataCount = lngNewIsoDataCount + 1
                    DuplicateIsoLineDataPoint GelData(mGelIndex).IsoData(lngIndex), udtNewIsoData, lngNewIsoDataCount, -glO16O18_DELTA, sngMonoMinus4Intensities(lngIndex), irtIReportTagTypeConstants.irtMonoMinus4
                End If
            Next lngIndex
            
            If lngNewIsoDataCount > GelData(mGelIndex).IsoLines Then
                ' Copy the data from udtNewIsoData() back into GelData(mGelIndex).IsoData
                ReDim GelData(mGelIndex).IsoData(lngNewIsoDataCount)
                For lngIndex = 1 To lngNewIsoDataCount
                    GelData(mGelIndex).IsoData(lngIndex) = udtNewIsoData(lngIndex)
                Next lngIndex
                
                lngDataPointsAdded = lngNewIsoDataCount - GelData(mGelIndex).IsoLines
                GelData(mGelIndex).IsoLines = lngNewIsoDataCount
                
                AddToAnalysisHistory mGelIndex, "Added " & CStr(lngDataPointsAdded) & " new data points using the '" & ISOS_COLUMN_MONO_PLUS4_ABUNDANCE & "' and '" & ISOS_COLUMN_MONO_MINUS4_ABUNDANCE & "' columns in the " & CSV_ISOS_PAIRS_SUFFIX & " file; total data point count = " & CStr(lngNewIsoDataCount)
            
                GelData(mGelIndex).DataStatusBits = GelData(mGelIndex).DataStatusBits Or GEL_DATA_STATUS_BIT_ADDED_MONOPLUSMINUS4_DATA
            
            End If
        End If
    End If
    
    
    With GelData(mGelIndex)
         ' Old: .PathtoDataFiles = GetPathWOFileName(CurrDataFName)
         ' New: data file folder path is the folder one folder up from .Filename's folder if .Filename's folder contains _Auto00000
         '      if .Filename's folder does not contain _Auto0000, then simply use .Filename's folder
        .PathtoDataFiles = DetermineParentFolderPath(.FileName)
        
        ' Note: CS Data is not loaded by this function
        ReDim .CSData(0)
        
        ' Find the minimum and maximum MW, Abundance, and MZ values, and set the filters
        MaxMZ = 0
        If .IsoLines > 0 Then
            ReDim Preserve .IsoData(.IsoLines)
            
            .MinMW = glHugeOverExp
            .MaxMW = 0
            .MinAbu = glHugeOverExp
            .MaxAbu = 0
            
            For lngIndex = 1 To .IsoLines
            If .IsoData(lngIndex).Abundance < .MinAbu Then .MinAbu = .IsoData(lngIndex).Abundance
                If .IsoData(lngIndex).Abundance > .MaxAbu Then .MaxAbu = .IsoData(lngIndex).Abundance
                
                FindMWExtremes .IsoData(lngIndex), .MinMW, .MaxMW, MaxMZ
            Next lngIndex
        
        Else
            ReDim .IsoData(0)
            .MinAbu = 0
            .MaxAbu = 0
            .MinMW = 0
            .MaxMW = 0
        End If
        
        ' If the IReport column was present and at least one entry had a non-zero MonoPlus2Abundance value,
        '  then set the GEL_DATA_STATUS_BIT_IREPORT data status bit
        If blnMonoPlus2DataPresent Then
            .DataStatusBits = .DataStatusBits Or GEL_DATA_STATUS_BIT_IREPORT
        Else
            .DataStatusBits = .DataStatusBits And Not GEL_DATA_STATUS_BIT_IREPORT
        End If
   
        .DataFilter(fltCSAbu, 2) = .MaxAbu             'put initial filters on max
        .DataFilter(fltIsoAbu, 2) = .MaxAbu
        .DataFilter(fltCSMW, 2) = .MaxMW
        .DataFilter(fltIsoMW, 2) = .MaxMW
        .DataFilter(fltIsoMZ, 2) = MaxMZ
        
        .DataFilter(fltEvenOddScanNumber, 0) = False
        .DataFilter(fltEvenOddScanNumber, 1) = 0       ' Show all scan numbers
        
        .DataFilter(fltIsoCS, 2) = 1000                'maximum charge state
         
      
        If Not blnValidScansFile Then
            If mScanInfoCount > 0 Then
                ReDim Preserve .ScanInfo(mScanInfoCount)
            Else
                ReDim .ScanInfo(0)
            End If
        End If
        
    End With
    
    If Not blnValidScansFile Then
        ' Elution time wasn't defined
        ' Define the default elution time to range from 0 to 1
        DefineDefaultElutionTimes GelData(mGelIndex).ScanInfo, 0, 1
        
        UpdateGelAdjacentScanPointerArrays mGelIndex
    End If
    
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load
        ReadCSVIsosFile = -5
        frmProgress.HideForm
        Exit Function
    End If
    
    ' Update the progress bar
    mCurrentProgressStep = mCurrentProgressStep + 1
    frmProgress.UpdateProgressBar mCurrentProgressStep
    frmProgress.InitializeSubtask "Sorting isotopic data", 0, GelData(mGelIndex).IsoLines
    
    If Not blnLoadPredefinedLCMSFeatures Then
        ' Sort the data, though we skip this step if we have loaded predefined LCMSFeatures
        SortIsotopicData mGelIndex
    End If
    
    If (GelData(mGelIndex).DataStatusBits And GEL_DATA_STATUS_BIT_IREPORT) = GEL_DATA_STATUS_BIT_IREPORT Then
        ' Fix the mono plus 2 abundance values
        FixIsosMonoPlus2Abu mGelIndex
    End If
        
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load
        ReadCSVIsosFile = -5
        frmProgress.HideForm
        Exit Function
    End If
    
    ReadCSVIsosFile = 0
    Exit Function

ReadCSVIsosFileErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadCSVIsosFile"
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    ReadCSVIsosFile = lngReturnValue

End Function

Private Function ReadCSVIsosFileWork(ByRef fso As FileSystemObject, _
                                     ByVal strIsosFilePath As String, _
                                     ByRef lngTotalBytesRead As Long, _
                                     ByRef intColumnMapping() As Integer, _
                                     ByVal blnValidScansFile As Boolean, _
                                     ByRef sngMonoPlus4Intensities() As Single, _
                                     ByRef sngMonoMinus4Intensities() As Single, _
                                     ByVal blnFilePrescanEnabled As Boolean, _
                                     ByVal blnIgnoreAllFiltersAndLoadAllData As Boolean, _
                                     ByVal plmPointsLoadMode As Integer) As Long
    
    Dim tsInFile As TextStream
    Dim strLineIn As String
    
    Dim lngLinesRead As Long
    Dim lngIndex As Long
    Dim lngScanNumber As Long
    Dim lngIsoIndex As Long
    Dim lngFeatureIndex As Long
    
    Dim blnColumnsDefined As Boolean
    Dim blnDataLine As Boolean
    Dim blnValidDataPoint As Boolean
    Dim blnStoreDataPoint As Boolean
    
    Dim strData() As String
    Dim strColumnHeader As String
    
    Dim sngFit As Single
    Dim sngAbundance As Single
    Dim intCharge As Integer
    
    Dim strUnknownColumnList As String
    Dim strMessage As String
    Dim lngReturnValue As Long
    
    Set mHashMapOfPointsKept = New Dictionary
    
On Error GoTo ReadCSVIsosFileWorkErrorHandler
    
    ' Reserve space in these arrays, but only if we're not pre-scanning the data
    If mReadMode <> rmReadModeConstants.rmPrescanData Then
        ReDim sngMonoPlus4Intensities(UBound(GelData(mGelIndex).IsoData))
        ReDim sngMonoMinus4Intensities(UBound(GelData(mGelIndex).IsoData))
    End If
    
    lngLinesRead = 0
    lngIsoIndex = 1 'Start index at 1 to handle LCMS_Features.txt type of index
    
    Set tsInFile = fso.OpenTextFile(strIsosFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream

        strLineIn = tsInFile.ReadLine
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If lngLinesRead Mod 100 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngTotalBytesRead, True
            If KeyPressAbortProcess > 1 Then Exit Do
        End If
        
        If blnColumnsDefined Then
            lngLinesRead = lngLinesRead + 1
            
            If mReadMode = rmReadModeConstants.rmStoreData Then
                GelData(mGelIndex).LinesRead = lngLinesRead
            End If
        End If
        
        ' Check for valid line (must contain at least one comma and must be
        ' the header line or start with a number)
        strLineIn = Trim(strLineIn)
        If InStr(strLineIn, ",") > 0 Then
            If blnColumnsDefined Then
                blnDataLine = IsNumeric(Left(strLineIn, 1))
            Else
                ' Haven't found the column header line yet
                ' If the line starts with a number, then assume the column header line is missing and use the default column order
                ' If the line starts with text, then assume this is the column header line
                
                If IsNumeric(Left(strLineIn, 1)) Then
                    ' Use the default column mappings
                   
                    intColumnMapping(IsosFileColumnConstants.ScanNumber) = IsosFileColumnConstants.ScanNumber
                    intColumnMapping(IsosFileColumnConstants.Charge) = IsosFileColumnConstants.Charge
                    intColumnMapping(IsosFileColumnConstants.Abundance) = IsosFileColumnConstants.Abundance
                    intColumnMapping(IsosFileColumnConstants.MZ) = IsosFileColumnConstants.MZ
                    intColumnMapping(IsosFileColumnConstants.Fit) = IsosFileColumnConstants.Fit
                    intColumnMapping(IsosFileColumnConstants.AverageMW) = IsosFileColumnConstants.AverageMW
                    intColumnMapping(IsosFileColumnConstants.MonoisotopicMW) = IsosFileColumnConstants.MonoisotopicMW
                    intColumnMapping(IsosFileColumnConstants.MostAbundantMW) = IsosFileColumnConstants.MostAbundantMW
                    intColumnMapping(IsosFileColumnConstants.FWHM) = IsosFileColumnConstants.FWHM
                    intColumnMapping(IsosFileColumnConstants.SignalToNoise) = IsosFileColumnConstants.SignalToNoise
                    intColumnMapping(IsosFileColumnConstants.MonoAbundance) = IsosFileColumnConstants.MonoAbundance
                    intColumnMapping(IsosFileColumnConstants.MonoPlus2Abundance) = IsosFileColumnConstants.MonoPlus2Abundance
                    intColumnMapping(IsosFileColumnConstants.MonoPlus4Abundance) = IsosFileColumnConstants.MonoPlus4Abundance
                    intColumnMapping(IsosFileColumnConstants.MonoMinus4Abundance) = IsosFileColumnConstants.MonoMinus4Abundance
                    intColumnMapping(IsosFileColumnConstants.IMSDriftTime) = IsosFileColumnConstants.IMSDriftTime

                    ' Column headers were not present
                     AddToAnalysisHistory mGelIndex, "Isos file " & fso.GetFileName(strIsosFilePath) & " did not contain column headers; using the default headers (" & GetDefaultIsosColumnHeaders(False, False) & ")"

                    blnDataLine = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, ",")
                    strUnknownColumnList = ""
                    
                    For lngIndex = 0 To UBound(strData)
                        
                        strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                        
                        Select Case strColumnHeader
                        Case ISOS_COLUMN_SCAN_NUM_A, ISOS_COLUMN_SCAN_NUM_B: intColumnMapping(IsosFileColumnConstants.ScanNumber) = lngIndex
                        Case ISOS_COLUMN_FRAME_NUM
                            ' We treat IMS frame number as if its the primary scan number
                            intColumnMapping(IsosFileColumnConstants.ScanNumber) = lngIndex
                        Case ISOS_COLUMN_IMS_SCAN_NUM
                            ' Ignore this column; VIPER does not track the IMS scan number
                        Case ISOS_COLUMN_CHARGE: intColumnMapping(IsosFileColumnConstants.Charge) = lngIndex
                        Case ISOS_COLUMN_ABUNDANCE: intColumnMapping(IsosFileColumnConstants.Abundance) = lngIndex
                        Case ISOS_COLUMN_MZ: intColumnMapping(IsosFileColumnConstants.MZ) = lngIndex
                        Case ISOS_COLUMN_FIT: intColumnMapping(IsosFileColumnConstants.Fit) = lngIndex
                        Case ISOS_COLUMN_AVERAGE_MW: intColumnMapping(IsosFileColumnConstants.AverageMW) = lngIndex
                        Case ISOS_COLUMN_MONOISOTOPIC_MW: intColumnMapping(IsosFileColumnConstants.MonoisotopicMW) = lngIndex
                        Case ISOS_COLUMN_MOSTABUNDANT_MW: intColumnMapping(IsosFileColumnConstants.MostAbundantMW) = lngIndex
                        Case ISOS_COLUMN_FWHM: intColumnMapping(IsosFileColumnConstants.FWHM) = lngIndex
                        Case ISOS_COLUMN_SIGNAL_NOISE: intColumnMapping(IsosFileColumnConstants.SignalToNoise) = lngIndex
                        Case ISOS_COLUMN_MONO_ABUNDANCE: intColumnMapping(IsosFileColumnConstants.MonoAbundance) = lngIndex
                        Case ISOS_COLUMN_MONO_PLUS2_ABUNDANCE: intColumnMapping(IsosFileColumnConstants.MonoPlus2Abundance) = lngIndex
                        Case ISOS_COLUMN_MONO_PLUS4_ABUNDANCE: intColumnMapping(IsosFileColumnConstants.MonoPlus4Abundance) = lngIndex
                        Case ISOS_COLUMN_MONO_MINUS4_ABUNDANCE: intColumnMapping(IsosFileColumnConstants.MonoMinus4Abundance) = lngIndex
                        Case ISOS_COLUMN_IMS_DRIFT_TIME: intColumnMapping(IsosFileColumnConstants.IMSDriftTime) = lngIndex
                        Case ISOS_COLUMN_IMS_CUMULATIVE_DRIFT_TIME
                            ' Ignore this column; VIPER does not track the IMS cumulative drift time
                        Case Else
                            ' Unknown column header; ignore it, but post an entry to the analysis history
                            If Len(strUnknownColumnList) > 0 Then
                                strUnknownColumnList = strUnknownColumnList & ", "
                            End If
                            strUnknownColumnList = strUnknownColumnList & Trim(strData(lngIndex))
                            
                            Debug.Assert False
                        End Select
                        
                    Next lngIndex
                    
                    If Len(strUnknownColumnList) > 0 Then
                        ' Unknown column header; ignore it, but post an entry to the analysis history
                        AddToAnalysisHistory mGelIndex, CSV_COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & fso.GetFileName(strIsosFilePath) & ": " & strUnknownColumnList & "; Known columns are: " & vbCrLf & GetDefaultIsosColumnHeaders(False, False)
                    End If
                    
                    blnDataLine = False
                
                End If
                
                ' Warn the user if any of the important columns are missing
                If intColumnMapping(IsosFileColumnConstants.ScanNumber) < 0 Or _
                   intColumnMapping(IsosFileColumnConstants.Abundance) < 0 Or _
                   intColumnMapping(IsosFileColumnConstants.MonoisotopicMW) < 0 Then
                   
                    If mReadMode = rmStoreData Then
                        strMessage = CSV_COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strIsosFilePath) & "; the expected columns are: " & vbCrLf & GetDefaultIsosColumnHeaders(True, False)
                        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                            MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
                        End If
                        AddToAnalysisHistory mGelIndex, strMessage
                    End If
                End If
                
                blnColumnsDefined = True
            End If
        End If
                
        If blnDataLine Then
            
            strData = Split(strLineIn, ",")
            
            If UBound(strData) >= 0 Then
                lngScanNumber = GetColumnValueLng(strData, intColumnMapping(IsosFileColumnConstants.ScanNumber), -1)
            Else
                lngScanNumber = -1
            End If
    
            If lngScanNumber >= 0 And (Not mEvenOddScanFilter Or (lngScanNumber Mod 2 = mEvenOddModCompareVal)) Then
        
                sngFit = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.Fit), 0)
                sngAbundance = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.Abundance), 0)
                intCharge = CInt(GetColumnValueLng(strData, intColumnMapping(IsosFileColumnConstants.Charge), 1))
                
                blnValidDataPoint = True
                
                If Not blnIgnoreAllFiltersAndLoadAllData Then
                    If sngFit <= mMaxFit Or mMaxFit <= 0 Then
                        If mFilterByAbundance Then
                            If sngAbundance < mAbundanceMin Or sngAbundance > mAbundanceMax Then
                                blnValidDataPoint = False
                            End If
                        End If
                    Else
                        blnValidDataPoint = False
                    End If
                End If
                
                If plmPointsLoadMode >= plmLoadMappedPointsOnly And blnValidDataPoint Then
                    If Not mPointsToKeep.Exists(lngIsoIndex) Then
                        blnValidDataPoint = False
                    Else
                        If plmPointsLoadMode = plmLoadOnePointPerLCMSFeature Then
                            lngFeatureIndex = mPointsToKeep.Item(lngIsoIndex)
                            
                            If Not mFeatureToScanMap.Item(lngFeatureIndex) = lngScanNumber Then
                                blnValidDataPoint = False
                            End If
                        End If
                    End If
                End If
            
                If blnValidDataPoint Then
                
                    mHashMapOfPointsKept.add lngIsoIndex, GelData(mGelIndex).IsoLines + 1
                
                    If mReadMode = rmReadModeConstants.rmPrescanData Then
                        mPrescannedData.AddDataPoint sngAbundance, intCharge, mValidDataPointCount
                    Else
                        If blnFilePrescanEnabled And Not blnIgnoreAllFiltersAndLoadAllData Then
                            If mPrescannedData.GetAbundanceByIndex(mValidDataPointCount) >= 0 Then
                                blnStoreDataPoint = True
                            Else
                                blnStoreDataPoint = False
                            End If
                        Else
                            blnStoreDataPoint = True
                        End If
                        
                        If blnStoreDataPoint Then
                            With GelData(mGelIndex)
                                .DataLines = .DataLines + 1
                                .IsoLines = .IsoLines + 1
                        
                                If Not blnValidScansFile Then
                                    ' Possibly add a new entry to .ScanInfo()
                                    ' Assumes data in the _isos.csv file is sorted by ascending scan number
                                    
                                    If mScanInfoCount = 0 Then
                                        mScanInfoCount = 1
                                        With .ScanInfo(1)
                                            .ScanNumber = lngScanNumber
                                            .ElutionTime = 0
                                            .ScanType = 1
                                        End With
                                    Else
                                        If .ScanInfo(mScanInfoCount).ScanNumber < lngScanNumber Then
                                            mScanInfoCount = mScanInfoCount + 1
                                            
                                            If mScanInfoCount > UBound(.ScanInfo) Then
                                                ReDim Preserve .ScanInfo(UBound(.ScanInfo) + SCAN_INFO_DIM_CHUNK)
                                            End If
                                            
                                            With .ScanInfo(mScanInfoCount)
                                                .ScanNumber = lngScanNumber
                                                .ElutionTime = 0
                                                .ScanType = 1
                                            End With
                                        End If
                                    End If
                                End If
                
                                If .IsoLines > UBound(.IsoData) Then
                                    ReDim Preserve .IsoData(UBound(.IsoData) + ISO_DATA_DIM_CHUNK)
                                    
                                    If mReadMode <> rmReadModeConstants.rmPrescanData Then
                                        ReDim Preserve sngMonoPlus4Intensities(UBound(.IsoData))
                                        ReDim Preserve sngMonoMinus4Intensities(UBound(.IsoData))
                                    End If
                                End If
                               
                                With .IsoData(.IsoLines)
                                    .ScanNumber = lngScanNumber
                                    .Charge = intCharge
                                    .Abundance = sngAbundance
                                    .MZ = GetColumnValueDbl(strData, intColumnMapping(IsosFileColumnConstants.MZ), 0)
                                    .Fit = sngFit
                                    .AverageMW = GetColumnValueDbl(strData, intColumnMapping(IsosFileColumnConstants.AverageMW), 0)
                                    .MonoisotopicMW = GetColumnValueDbl(strData, intColumnMapping(IsosFileColumnConstants.MonoisotopicMW), 0)
                                    .MostAbundantMW = GetColumnValueDbl(strData, intColumnMapping(IsosFileColumnConstants.MostAbundantMW), 0)
                                    .FWHM = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.FWHM), 0)
                                    .SignalToNoise = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.SignalToNoise), 0)
                                    .IntensityMono = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.MonoAbundance), 0)
                                    .IntensityMonoPlus2 = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.MonoPlus2Abundance), 0)
                                    .IMSDriftTime = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.IMSDriftTime), 0)
                                End With
                            
                                If mReadMode <> rmReadModeConstants.rmPrescanData Then
                                    sngMonoPlus4Intensities(.IsoLines) = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.MonoPlus4Abundance), 0)
                                    sngMonoMinus4Intensities(.IsoLines) = GetColumnValueSng(strData, intColumnMapping(IsosFileColumnConstants.MonoMinus4Abundance), 0)
                                End If
                            End With
                        End If
                    End If
                    
                    mValidDataPointCount = mValidDataPointCount + 1
                End If
            End If
            
            lngIsoIndex = lngIsoIndex + 1
        End If
    Loop
        
    tsInFile.Close
    ReadCSVIsosFileWork = 0
    Exit Function

ReadCSVIsosFileWorkErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadCSVIsosFileWork"
    
    On Error Resume Next
    tsInFile.Close
    
    ReadCSVIsosFileWork = lngReturnValue
End Function

Private Function ReadCSVScanFile(ByRef fso As FileSystemObject, ByVal strScansFilePath As String, ByVal strBaseFilePath As String, ByRef lngTotalBytesRead As Long) As Long
    ' Returns 0 if no error, the error number if an error

    Dim tsInFile As TextStream
    Dim strLineIn As String
    Dim strUnknownColumnList As String
    Dim strMessage As String
    
    Dim lngIndex As Long
    Dim lngReturnValue As Long
    
    Dim blnColumnsDefined As Boolean
    Dim blnDataLine As Boolean
    
    Dim strData() As String
    Dim strColumnHeader As String
    
    Dim intColumnMapping() As Integer
    Dim intScanType As Integer
    Dim sngMaxElutionTime As Single
    
On Error GoTo ReadCSVScanFileErrorHandler

    ReDim intColumnMapping(SCAN_FILE_COLUMN_COUNT - 1)
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To SCAN_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    If Len(strBaseFilePath) = 0 Then
        strBaseFilePath = fso.GetBaseName(strScansFilePath)
    End If
    
    frmProgress.UpdateCurrentSubTask "Reading Scan Info file"
    
    With GelData(mGelIndex)
        ReDim .ScanInfo(SCAN_INFO_DIM_CHUNK)        ' 1-based array
    
        mScanInfoCount = 0
        Set tsInFile = fso.OpenTextFile(strScansFilePath, ForReading, False)
        Do While Not tsInFile.AtEndOfStream
    
            strLineIn = tsInFile.ReadLine
            lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
            
            If mScanInfoCount Mod 50 = 0 Then
                frmProgress.UpdateSubtaskProgressBar lngTotalBytesRead
            End If
            
            ' Check for valid line (must contain at least one comma and must be
            ' the header line or start with a number)
            strLineIn = Trim(strLineIn)
            If InStr(strLineIn, ",") > 0 Then
                If blnColumnsDefined Then
                    blnDataLine = IsNumeric(Left(strLineIn, 1))
                Else
                    ' Haven't found the column header line yet
                    ' If the line starts with a number, then assume the column header line is missing and use the default column order
                    ' If the line starts with text, then assume this is the column header line
                    
                    If IsNumeric(Left(strLineIn, 1)) Then
                        ' Use the default column mappings

                        intColumnMapping(ScanFileColumnConstants.ScanNumber) = ScanFileColumnConstants.ScanNumber
                        intColumnMapping(ScanFileColumnConstants.ScanTime) = ScanFileColumnConstants.ScanTime
                        intColumnMapping(ScanFileColumnConstants.ScanType) = ScanFileColumnConstants.ScanType
                        intColumnMapping(ScanFileColumnConstants.NumDeisotoped) = ScanFileColumnConstants.NumDeisotoped
                        intColumnMapping(ScanFileColumnConstants.NumPeaks) = ScanFileColumnConstants.NumPeaks
                        intColumnMapping(ScanFileColumnConstants.TIC) = ScanFileColumnConstants.TIC
                        intColumnMapping(ScanFileColumnConstants.BPImz) = ScanFileColumnConstants.BPImz
                        intColumnMapping(ScanFileColumnConstants.BPI) = ScanFileColumnConstants.BPI
                        intColumnMapping(ScanFileColumnConstants.TimeDomainSignal) = ScanFileColumnConstants.TimeDomainSignal
                        intColumnMapping(ScanFileColumnConstants.PeakIntensityThreshold) = ScanFileColumnConstants.PeakIntensityThreshold
                        intColumnMapping(ScanFileColumnConstants.PeptideIntensityThreshold) = ScanFileColumnConstants.PeptideIntensityThreshold
                        intColumnMapping(ScanFileColumnConstants.IMSFramePressureFront) = ScanFileColumnConstants.IMSFramePressureFront
                        intColumnMapping(ScanFileColumnConstants.IMSFramePressureBack) = ScanFileColumnConstants.IMSFramePressureBack

                        ' Column headers were not present
                         AddToAnalysisHistory mGelIndex, "Scans file " & fso.GetFileName(strScansFilePath) & " did not contain column headers; using the default headers (" & GetDefaultScansColumnHeaders(False, False) & ")"

                        blnDataLine = True
                    Else
                        ' Define the column mappings
                        strData = Split(strLineIn, ",")
                        strUnknownColumnList = ""
                        
                        For lngIndex = 0 To UBound(strData)

                            strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                            
                            Select Case strColumnHeader
                            Case SCANS_COLUMN_SCAN_NUM: intColumnMapping(ScanFileColumnConstants.ScanNumber) = lngIndex
                            Case SCANS_COLUMN_FRAME_NUM
                                ' VIPER treats frame_num as scan_num
                                intColumnMapping(ScanFileColumnConstants.ScanNumber) = lngIndex
                            Case SCANS_COLUMN_TIME_A, SCANS_COLUMN_TIME_B: intColumnMapping(ScanFileColumnConstants.ScanTime) = lngIndex
                            Case SCANS_COLUMN_FRAME_TIME
                                ' VIPER treats frame_time as scan_time
                                intColumnMapping(ScanFileColumnConstants.ScanTime) = lngIndex
                            Case SCANS_COLUMN_DRIFT_TIME
                                ' Old column that was only used in the 2008 version of the IMS file format
                                ' Ignore this column
                            Case SCANS_COLUMN_TYPE: intColumnMapping(ScanFileColumnConstants.ScanType) = lngIndex
                            Case SCANS_COLUMN_NUM_DEISOTOPED: intColumnMapping(ScanFileColumnConstants.NumDeisotoped) = lngIndex
                            Case SCANS_COLUMN_NUM_PEAKS: intColumnMapping(ScanFileColumnConstants.NumPeaks) = lngIndex
                            Case SCANS_COLUMN_TIC: intColumnMapping(ScanFileColumnConstants.TIC) = lngIndex
                            Case SCANS_COLUMN_BPI_MZ: intColumnMapping(ScanFileColumnConstants.BPImz) = lngIndex
                            Case SCANS_COLUMN_BPI: intColumnMapping(ScanFileColumnConstants.BPI) = lngIndex
                            Case SCANS_COLUMN_TIME_DOMAIN_SIGNAL: intColumnMapping(ScanFileColumnConstants.TimeDomainSignal) = lngIndex
                            Case SCANS_COLUMN_PEAK_INTENSITY_THRESHOLD: intColumnMapping(ScanFileColumnConstants.PeakIntensityThreshold) = lngIndex
                            Case SCANS_COLUMN_PEPTIDE_INTENSITY_THRESHOLD: intColumnMapping(ScanFileColumnConstants.PeptideIntensityThreshold) = lngIndex
                            Case SCANS_COLUMN_IMS_FRAME_PRESSURE, SCANS_COLUMN_IMS_FRAME_PRESSURE_FRONT
                                intColumnMapping(ScanFileColumnConstants.IMSFramePressureFront) = lngIndex
                            Case SCANS_COLUMN_IMS_FRAME_PRESSURE_BACK
                                intColumnMapping(ScanFileColumnConstants.IMSFramePressureBack) = lngIndex
                            Case Else
                                ' Unknown column header; ignore it, but post an entry to the analysis history
                                If Len(strUnknownColumnList) > 0 Then
                                    strUnknownColumnList = strUnknownColumnList & ", "
                                End If
                                strUnknownColumnList = strUnknownColumnList & Trim(strData(lngIndex))
                                
                                Debug.Assert False
                            End Select
                            
                        Next lngIndex
                        
                        If Len(strUnknownColumnList) > 0 Then
                            ' Unknown column header; ignore it, but post an entry to the
                            AddToAnalysisHistory mGelIndex, CSV_COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & fso.GetFileName(strScansFilePath) & ": " & strUnknownColumnList & "; Known columns are: " & vbCrLf & GetDefaultScansColumnHeaders(False, False)
                        End If
                        
                        blnDataLine = False
                    
                    End If
                    
                    ' Warn the user if any of the important columns are missing
                    If intColumnMapping(ScanFileColumnConstants.ScanNumber) < 0 Or _
                       intColumnMapping(ScanFileColumnConstants.ScanTime) < 0 Or _
                       intColumnMapping(ScanFileColumnConstants.ScanType) < 0 Then
                       
                        strMessage = CSV_COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strScansFilePath) & "; the expected columns are: " & vbCrLf & GetDefaultScansColumnHeaders(True, False)
                        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                            MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
                        End If
                        AddToAnalysisHistory mGelIndex, strMessage
                    End If
                    
                    blnColumnsDefined = True
                End If
            End If
                    
            If blnDataLine Then
                
                strData = Split(strLineIn, ",")
                intScanType = CInt(GetColumnValueLng(strData, intColumnMapping(ScanFileColumnConstants.ScanType), 1))
    
                If intScanType <= 1 Then
                    mScanInfoCount = mScanInfoCount + 1
                    If mScanInfoCount > UBound(.ScanInfo) Then
                        ReDim Preserve .ScanInfo(UBound(.ScanInfo) + SCAN_INFO_DIM_CHUNK)
                    End If
                    
                    ' Update the scan Info data
                    With .ScanInfo(mScanInfoCount)
                        .ScanNumber = GetColumnValueLng(strData, intColumnMapping(ScanFileColumnConstants.ScanNumber), 0)
                        .ElutionTime = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.ScanTime), 0)
                        If .ElutionTime > sngMaxElutionTime Then sngMaxElutionTime = .ElutionTime
                        
                        .ScanType = intScanType
                        If .ScanType < 1 Then .ScanType = 1
    
                        .ScanFileName = strBaseFilePath & "." & Format(.ScanNumber, "00000")
                        .ScanPI = 0
    
                        .NumDeisotoped = GetColumnValueLng(strData, intColumnMapping(ScanFileColumnConstants.NumDeisotoped), 0)
                        .NumPeaks = GetColumnValueLng(strData, intColumnMapping(ScanFileColumnConstants.NumPeaks), 0)
                        
                        .TIC = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.TIC), 0)
                        .BPI = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.BPI), 0)
                        .BPImz = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.BPImz), 0)
    
                        .TimeDomainSignal = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.TimeDomainSignal), 0)
    
                        .PeakIntensityThreshold = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.PeakIntensityThreshold), 0)
                        .PeptideIntensityThreshold = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.PeptideIntensityThreshold), 0)
    
                        .FrequencyShift = 0
                    
                        .IMSFramePressure = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.IMSFramePressureFront), 0)
                        
                        ' Note: Not reading IMSFramePressureBack
                        ' .IMSFramePressureBack = GetColumnValueSng(strData, intColumnMapping(ScanFileColumnConstants.IMSFramePressureBack), 0)
                    
                    End With
                End If
            End If
        
        Loop
        
        tsInFile.Close
    
        If mScanInfoCount > 0 Then
            ReDim Preserve .ScanInfo(mScanInfoCount)
        Else
            ReDim .ScanInfo(0)
        End If
    End With
    
    If sngMaxElutionTime = 0 Then
        ' Elution time wasn't defined
        ' Define the default elution time to range from 0 to 1
        DefineDefaultElutionTimes GelData(mGelIndex).ScanInfo, 0, 1
    End If
    
    UpdateGelAdjacentScanPointerArrays mGelIndex
    
    ReadCSVScanFile = 0
    Exit Function

ReadCSVScanFileErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadCSVScanFile"
    
    On Error Resume Next
    tsInFile.Close
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    ReadCSVScanFile = lngReturnValue

End Function

Private Function ReadLCMSFeatureFiles(ByRef fso As FileSystemObject, _
                                      ByVal strLCMSFeaturesFilePath As String, _
                                      ByVal strLCMSFeatureToPeakMapFilePath As String, _
                                      ByVal sngAutoMapDataPointsMassTolerancePPM As Single, _
                                      ByVal plmPointsLoadMode As Integer) As Long
                                      
    Dim objReadLCMSFeatures As clsFileIOPredefinedLCMSFeatures
    Dim lngReturnCode As Long
    
    Set objReadLCMSFeatures = New clsFileIOPredefinedLCMSFeatures
    
    
    objReadLCMSFeatures.ShowDialogBoxes = Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
    objReadLCMSFeatures.AutoMapDataPointsMassTolerancePPM = sngAutoMapDataPointsMassTolerancePPM
    
    objReadLCMSFeatures.PointsLoadMode = plmPointsLoadMode
    
    If plmPointsLoadMode >= plmLoadMappedPointsOnly Then
        objReadLCMSFeatures.HashMapOfPointsKept = mHashMapOfPointsKept
    End If
    
    lngReturnCode = objReadLCMSFeatures.LoadLCMSFeatureFiles(strLCMSFeaturesFilePath, strLCMSFeatureToPeakMapFilePath, mGelIndex)
    
    ReadLCMSFeatureFiles = lngReturnCode
    
End Function

Private Function ResolveCSVFilePaths(ByVal strFilePath As String, ByRef strScansFilePath As String, ByRef strIsosFilePath As String, ByRef strBaseFilePath As String) As Boolean
    ' Define the _scans.csv and _isos.csv FilePaths, given strFilePath
    ' strFilePath could contain either the _scans.csv name or the _isos.csv name
    ' Does not confirm that the files actually exist
    
    
    Dim lngCharLoc As Long
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    strScansFilePath = ""
    strIsosFilePath = ""
    strBaseFilePath = ""
    
    lngCharLoc = InStr(LCase(strFilePath), LCase(CSV_SCANS_FILE_SUFFIX))
    If lngCharLoc >= 1 Then
        ' strFilePath contains the _scans.csv file to this function
        ' Look for the corresponding _isos.csv file
        strScansFilePath = strFilePath
        strBaseFilePath = Left(strFilePath, lngCharLoc - 1)
        strIsosFilePath = strBaseFilePath & CSV_ISOS_IC_FILE_SUFFIX
        
        If Not FileExists(strIsosFilePath) Then
            strIsosFilePath = strBaseFilePath & CSV_ISOS_PAIRS_SUFFIX
        End If
        
        If Not FileExists(strIsosFilePath) Then
            strIsosFilePath = strBaseFilePath & CSV_ISOS_FILE_SUFFIX
        End If
        
        blnSuccess = True
        
    Else
        ' Assume strFilePath contains the _isos.csv file (or similar)
        ' Look for the Scans.csv file
        
        ' Define the base name
        ' First look for pairs_isos.csv
        lngCharLoc = InStr(LCase(strFilePath), LCase(CSV_ISOS_PAIRS_SUFFIX))
        If lngCharLoc < 1 Then
            ' No match, look for isos.csv
            lngCharLoc = InStr(LCase(strFilePath), LCase(CSV_ISOS_IC_FILE_SUFFIX))
            If lngCharLoc < 1 Then
                ' No match, look for isos.csv
                lngCharLoc = InStr(LCase(strFilePath), LCase(CSV_ISOS_FILE_SUFFIX))
                If lngCharLoc < 1 Then
                    ' No match
                End If
            End If
        End If
                
        If lngCharLoc >= 1 Then
            strIsosFilePath = strFilePath
            
            strBaseFilePath = Left(strFilePath, lngCharLoc - 1)
            strScansFilePath = strBaseFilePath & CSV_SCANS_FILE_SUFFIX
            
            blnSuccess = True
        Else
            blnSuccess = False
        End If
        
    End If
    
    ResolveCSVFilePaths = blnSuccess
    
End Function

Private Function StripQuotes(ByVal strText As String) As String

    If Len(strText) > 2 Then
        If Left(strText, 1) = """" And Right(strText, 1) = """" Then
            strText = Mid(strText, 2, Len(strText) - 2)
        End If
    End If
             
    StripQuotes = strText
End Function

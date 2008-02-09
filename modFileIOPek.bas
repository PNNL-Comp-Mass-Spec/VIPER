Attribute VB_Name = "modFileIOPEK"
Option Explicit

'following set of constants is used with PEK functions
Public Const PEK_D_FILENAME = "Filename:"
Private Const PEK_D_TIME_DOMAIN_SIGNAL_LEVEL = "Time domain signal level:"
''Private Const PEK_D_MEDIA = "Media type:"
''Private Const PEK_D_mOVERz_RANGE = "m/z Range:"
''Private Const PEK_D_MAX_CS = "Maximum CS:"
''Private Const PEK_D_THRESHOLD = "Threshold:"
''Private Const PEK_D_MIN_SNR = "Minimum peak S/N:"
''Private Const PEK_D_MAX_FIT = "Maximum fit:"
''Private Const PEK_N_CS_BLOCK = "First CS,"
''Private Const PEK_N_IS_BLOCK = "CS,  Abundance,   m/z,"
Private Const PEK_D_PEAKS_CNT = "Number of peaks in spectrum ="
Private Const PEK_D_IS_CNT = "Number of isotopic distributions detected ="
''Private Const PEK_N_CS_BLOCK_1 = "Charge state mass transform results:"
''Private Const PEK_N_IS_BLOCK_1 = "Isotopic mass transform results:"
''Private Const PEK_N_CALIBRATION = "Calibration:"

'this constants are used only during reading of PEK file
'they are used to identify line by it's left 8 characters
Private Const t8CALIBRATION1 = "Calibrat"
Private Const t8CALIBRATION2 = " Calibra"
Private Const t8FILENAME = "Filename"
Private Const t8DATABASE = "Database"
Public Const t8DATA_CS = "First CS"
Public Const t8DATA_ISO = "CS,  Abu"
Public Const t8DATA_ISO_Tabbed = "CS" & vbTab & "Abund"

Private Const t8MEDIATYPE = "Media ty"
Private Const t8FREQSHIFT = " Freq sh"
''Private Const t8DELTADATABLOCK = "Monoisot"
''Private Const t8DELTABLOCKEND = "End of d"      'not used
''Private Const t8DELTA = " Delta ="
''Private Const t8DELTA_TOL = " Toleren"
''Private Const t8DELTA_TOL_1 = " Toleran"
''Private Const t8TAG_MASS = " Tag Mas"
Private Const t8MAX_DELTAS = " Maximum"
Private Const t8TTL_SEQ_TIME = ""               'time for one scan
Private Const t8TIME_DOMAIN = "Time dom"

Private Const t3ARG_A = "A ="
Private Const t3ARG_B = "B ="
Private Const t3ARG_C = "C ="
Private Const t3ARG_D = "D ="
Private Const t3ARG_E = "E ="
Private Const t3ARG_F = "F ="
Private Const t3ARG_G = "G ="
Private Const t3ARG_H = "H ="
Private Const t3ARG_I = "I ="
Private Const t3ARG_J = "J ="
Private Const t3EQUATION = "m/z"

Private Const t4RT = "RT ="

Private Const t5WIFF = "wiff-"    'used with scan number in QTof files

Private Const LINE_NOTHING = -1
Private Const LINE_CALIBRATION = 0
Private Const LINE_DATA_CS = 1
Private Const LINE_DATA_ISO = 2
Private Const LINE_FILENAME_AKA_SCAN_NUMBER = 3
Private Const LINE_DATABASE = 4
Private Const LINE_EQUATION = 5
Private Const LINE_CAL_ARGUMENT = 6
Private Const LINE_FREQUENCY = 7
Private Const LINE_INTENSITY = 8
Private Const LINE_MEDIA = 9
''Private Const LINE_DELTA = 10
''Private Const LINE_DELTA_TOLERANCE = 11
''Private Const LINE_DELTA_TAGMASS = 12
''Private Const LINE_DELTA_MAX = 13
''Private Const LINE_DATA_DD = 14
''Private Const LINE_DATA_DD_END = 15
Private Const LINE_WHATEVER = 16
Private Const LINE_TIME_DOMAIN_SIGNAL = 17
Private Const LINE_RETENTION_TIME = 18
Private Const LINE_NUMBER_OF_PEAKS = 19
Private Const LINE_NUMBER_OF_ISOTOPIC_DISTRIBUTIONS = 20

Private Const SCAN_INFO_DIM_CHUNK As Long = 10000
Private Const ISO_DATA_DIM_CHUNK As Long = 25000

Private Const ISONUM_FIELD_COUNT_MAX = 12       ' Reserve space for up to 12 numeric fields; we may not use all of these

' Note: These should all be lowercase string values
Private Const ISOS_COLUMN_CS As String = "cs"
Private Const ISOS_COLUMN_ABUNDANCE As String = "abundance"
Private Const ISOS_COLUMN_MZ As String = "m/z"
Private Const ISOS_COLUMN_ISO_FIT As String = "fit"
Private Const ISOS_COLUMN_AVG_MASS As String = "average mw"
Private Const ISOS_COLUMN_MONO_MASS As String = "monoisotopic mw"
Private Const ISOS_COLUMN_MOST_ABU_MASS As String = "most abundant mw"          ' Aka TMA (The Most Abundant)
Private Const ISOS_COLUMN_IREP_MW_MONO_ABU As String = "imono"
Private Const ISOS_COLUMN_IREP_2DA As String = "i+2"
Private Const ISOS_COLUMN_PCT_MOST_ABUNDANT As String = "pct most abundant"
Private Const ISOS_COLUMN_ELEMENT_LABEL_N14N15 As String = "n14/n15"
Private Const ISOS_COLUMN_ELEMENT_LABEL_N14N15_RATIO As String = "n14/n15 ratio"


' Note: These should all be lowercase string values
Private Const CS_COLUMN_FIRST_CS As String = "first cs"
Private Const CS_COLUMN_NUMBER_OF_CS As String = "number of cs"
Private Const CS_COLUMN_ABUNDANCE As String = "abundance"
Private Const CS_COLUMN_MASS As String = "mass"
Private Const CS_COLUMN_STD_DEV As String = "standard deviation"

' The following corresponds to varData() in LoadNewPEK and LineNow for Isotopic data
' The values represent the column number in the file that holds the given type of data, though the data after column irdMWMono can vary
' If a header row is present, then use ???? to map column number to column content type
Private Const ISOS_DATA_COLUMN_COUNT As Integer = 14
Private Enum irdIsoRawDataIndex
    irdCS = 0
    irdAbu = 1
    irdMZ = 2
    irdIsoFit = 3
    irdMWavg = 4
    irdMWMono = 5
    irdMWTMA = 6                    ' Aka TMA (The Most Abundant)
    irdIRepMWMonoAbu = 7            ' Previously was irdERorN14N15OrIRepMWMonoAbu = 7
    irdIRep2Da = 8                  ' Previously was irdN14N15RatioOrIRep2Da = 8
    irdIsotopeTag = 9
    irdIsotopeTagRatio = 10
    irdPctMostAbundant = 11
    irdDBMatchMassError = 12
    irdPeptideIdentity = 13
End Enum

' The following corresponds to varData() in LoadNewPEK and LineNow for Charge state data
Private Const CS_DATA_COLUMN_COUNT As Integer = 5
Private Enum crdCSRawDataIndex
    crdCS = 0
    crdNumberOfCS = 1
    crdAbu = 2
    crdAverageMW = 3
    crdMWStDev = 4
    crdMTID = 5
End Enum

''' The following corresponds to the TmpNum() array in LoadNewPEK
''Private Enum itmIsoTempNumIndex
''    itmLineType = 1
''    itmScanNumber = 2
''    itmCS = 3
''    itmAbu = 4
''    itmMOverZ = 5
''    itmFit = 6
''    itmMWavg = 7
''    itmMWMono = 8
''    itmMWMA = 9
''    itmIsotopicFitRatio = 10
''    itmIsotopicAtomCount = 11
''End Enum

Private Enum rmReadModeConstants
    rmPrescanData = 0
    rmStoreData = 1
    rmReadComplete = 2
End Enum

Private mGelIndex As Long
Private mScanInfoCount As Long
Private mMaxElutionTime As Single

Private mEvenOddScanFilter As Boolean
Private mEvenOddModCompareVal As Integer

Private mMaxFit As Double
Private mFilterByAbundance As Boolean
Private mAbundanceMin As Double
Private mAbundanceMax As Double

Private mDataFilterMode As dfmCSandIsoDataFilterModeConstants

Private mMaximumDataCountEnabled As Boolean
Private mMaximumDataCountToLoad As Long

Private mPrescannedData As clsFileIOPrescannedData

Private mIReportFile As Boolean
Private mIsotopeLabelTagsPresent As Boolean

Private mValidDataPointCount As Long
Private mSubtaskMessage As String

Private mReadMode As rmReadModeConstants
Private mCurrentProgressStep As Integer

Private mWarnedMissingRequiredIsosColumnHeaders As Boolean
Private mWarnedUnknownIsosColumnHeader As Boolean

Private mWarnedMissingRequiredCSColumnHeaders As Boolean
Private mWarnedUnknownCSColumnHeader As Boolean

'control of loading
' Note: ThisLine keeps track of whether we are reading Charge State (not deisotoped) or Isotopic (deisotoped) data
Private ThisLine As Integer
Private CalibrationIn As Boolean   'read only once
Private DatabaseIn As Boolean      'read only once
Private MediaTypeIn As Boolean     'read only once
'

Private Function ExtractTimeDomainSignalFromPEK(ByVal strInputFilePath As String, ByVal Ind As Long) As Boolean
'-------------------------------------------------------------------------
' Looks for the "Time domain signal level" line in the PEK file
' Note that _ic.pek files have "Time domain signal level" values of 1 throughout the file (a bug)
' Also, note that DeCal PEK files do not have the "Time domain signal level" entry
' Thus, if blnUseOriginalPEKFile = True, then this function looks for and tries to use
'  the original .PEK file in the directory contained in strInputFilePath
'
' This function should be called from LoadNewPEK, and it thus will update the
'  SubTask Progress Bar rather than the main progress bar
'
' Returns True if data loaded, False otherwise
'-------------------------------------------------------------------------

Const SCAN_DIM_CHUNK = 1000

Dim strProcessedPEKExtensionList() As String
Dim lngExtensionCount As Long

Dim strMoreExtensions() As String
Dim lngMoreExtensionAdd As Long

Dim strTestExtension As String
Dim strTestFilePath As String, strWorkingPEKFilePath As String

Dim InFileNum As Integer
Dim strCurrentLine As String
Dim intLineType As Integer

Dim lngIndex As Long, lngCompareIndex As Long
Dim blnMatched As Boolean

Dim varValue As Variant
Dim lngCurrentScanNumber As Long
Dim lngCurrentScanNumberSaved As Long

Dim dblMostRecentSignal As Double

Dim lngByteCountTotal As Long
Dim lngTotalBytesRead As Long

Dim lngTimeDomainValueCount As Long     ' Number of time domain values found
Dim lngMaxScanNumberDimmed As Long
Dim dblTimeDomainSignal() As Double     ' 1-based array: Time domain signal level
                                        ' Array index corresponds to scan number

Dim blnAutoNumberScans As Boolean
Dim objScanNumberTracker As clsScanNumberTracker

Dim fso As New FileSystemObject

On Error GoTo ExtractTimeDomainSignalFromPEKErrorHandler

' Make sure .ScanInfo() is not empty
lngCurrentScanNumber = UBound(GelData(Ind).ScanInfo())
If lngCurrentScanNumber = 0 Then
    ' This shouldn't happen
    Debug.Assert False
    Set fso = Nothing
    ExtractTimeDomainSignalFromPEK = False
    Exit Function
End If

' Initialize the TimeDomainSignal array
' Reserve space for SCAN_DIM_CHUNK = 1000 scans at a time
lngMaxScanNumberDimmed = SCAN_DIM_CHUNK
If lngCurrentScanNumber > lngMaxScanNumberDimmed Then
    lngMaxScanNumberDimmed = lngCurrentScanNumber
End If
ReDim dblTimeDomainSignal(lngMaxScanNumberDimmed)

' Initialize strWorkingPEKFilePath
strWorkingPEKFilePath = strInputFilePath

' If strInputFilePath ends in one of the extensions specified in DEFAULT_PEK_FILE_EXTENSION_ORDER,
'  then look for an un-processed PEK in the folder given by strInputFilePath
lngExtensionCount = ParseString(DEFAULT_PEK_FILE_EXTENSION_ORDER, strProcessedPEKExtensionList(), 100, ",", "", True, True, False)

' In addition, combine the extensions defined in .PEKFileExtensionPreferenceOrder with those found in DEFAULT_PEK_FILE_EXTENSION_ORDER
lngMoreExtensionAdd = ParseString(glbPreferencesExpanded.AutoAnalysisOptions.PEKFileExtensionPreferenceOrder, strMoreExtensions(), 100, ",", "", True, True, False)

If lngMoreExtensionAdd > 0 Then
    For lngIndex = 0 To lngMoreExtensionAdd - 1
        blnMatched = False
        For lngCompareIndex = 0 To lngExtensionCount - 1
            If LCase(Trim(strProcessedPEKExtensionList(lngCompareIndex))) = LCase(Trim(strMoreExtensions(lngIndex))) Then
                blnMatched = True
                Exit For
            End If
        Next lngCompareIndex
        
        If Not blnMatched Then
            ReDim Preserve strProcessedPEKExtensionList(lngExtensionCount)
            strProcessedPEKExtensionList(lngExtensionCount) = Trim(strMoreExtensions(lngIndex))
            lngExtensionCount = lngExtensionCount + 1
        End If
    Next lngIndex
End If

' See if strInputFilePath ends in an extension listed in strProcessedPEKExtensionList()
' However, skip any extension that is ".pek" or "pek"
' Additionally, skip any extension that ends in ."csv" or ".mzXML"
For lngIndex = 0 To lngExtensionCount - 1
    strTestExtension = LCase(Trim(strProcessedPEKExtensionList(lngIndex)))
    If strTestExtension <> ".pek" And strTestExtension <> ".pek" And Right(strTestExtension, 3) <> "csv" And Right(strTestExtension, 3) <> "xml" Then
        If Right(LCase(strInputFilePath), Len(strTestExtension)) = strTestExtension Then
            strTestFilePath = Left(strInputFilePath, Len(strInputFilePath) - Len(strTestExtension)) & ".pek"
            If fso.FileExists(strTestFilePath) Then
                strWorkingPEKFilePath = strTestFilePath
                Exit For
            End If
        End If
    End If
Next lngIndex

lngByteCountTotal = FileLen(strWorkingPEKFilePath)
frmProgress.InitializeSubtask "Scanning file for Time Domain Signal values", 0, lngByteCountTotal


InFileNum = FreeFile()
Open strWorkingPEKFilePath For Input As InFileNum

' Read each line in strWorkingPEKFilePath and parse
' Note that the Time Domain Signal line comes before the Filename line,
' and we thus do not know the scan number (file number) until after we've read each time domain signal value

blnAutoNumberScans = False
Set objScanNumberTracker = New clsScanNumberTracker

lngCurrentScanNumber = 0
objScanNumberTracker.Reset

Do While Not EOF(InFileNum)
    Line Input #InFileNum, strCurrentLine
       
    lngTotalBytesRead = lngTotalBytesRead + Len(strCurrentLine) + 2      ' Add 2 bytes to account for CrLf at end of line
    If lngTotalBytesRead Mod 100 = 0 Then
        frmProgress.UpdateSubtaskProgressBar lngTotalBytesRead
        If KeyPressAbortProcess > 1 Then Exit Do
    End If
     
    LineNowScanNumberOrTimeDomainSignal strCurrentLine, intLineType, varValue
    
    Select Case intLineType
    Case LINE_FILENAME_AKA_SCAN_NUMBER
        
        If blnAutoNumberScans Then
            lngCurrentScanNumber = objScanNumberTracker.GetNextAutoNumberedScan()
        Else
            lngCurrentScanNumber = CLng(varValue)
            objScanNumberTracker.AddScanNumberAndUpdateAverageIncrement lngCurrentScanNumber
        End If
       
        If lngCurrentScanNumberSaved > lngCurrentScanNumber Then     'can not accept non-ascending scan numbers
            ' Auto-number from now on
            If Not blnAutoNumberScans Then
                objScanNumberTracker.SetAutoNumberIncrementToCurrentAverage
                blnAutoNumberScans = True
            End If
            lngCurrentScanNumber = objScanNumberTracker.GetNextAutoNumberedScan()
        End If
        
        Do While lngCurrentScanNumber > lngMaxScanNumberDimmed
            lngMaxScanNumberDimmed = lngMaxScanNumberDimmed + SCAN_DIM_CHUNK
            ReDim Preserve dblTimeDomainSignal(lngMaxScanNumberDimmed)
        Loop
        
        dblTimeDomainSignal(lngCurrentScanNumber) = dblMostRecentSignal
        lngTimeDomainValueCount = lngTimeDomainValueCount + 1
        
        lngCurrentScanNumberSaved = lngCurrentScanNumber
        
    Case LINE_TIME_DOMAIN_SIGNAL
        dblMostRecentSignal = varValue
    End Select
Loop

Close InFileNum

If lngTimeDomainValueCount > 0 Then
    With GelData(Ind)
        For lngIndex = 1 To UBound(.ScanInfo)
            lngCurrentScanNumber = .ScanInfo(lngIndex).ScanNumber
            If lngCurrentScanNumber <= lngMaxScanNumberDimmed Then
                .ScanInfo(lngIndex).TimeDomainSignal = dblTimeDomainSignal(lngCurrentScanNumber)
            Else
                .ScanInfo(lngIndex).TimeDomainSignal = 0
            End If
        Next lngIndex
    End With
    ExtractTimeDomainSignalFromPEK = True
Else
    ExtractTimeDomainSignalFromPEK = False
End If

Set fso = Nothing

Exit Function

ExtractTimeDomainSignalFromPEKErrorHandler:
    Debug.Print "Error in ExtractTimeDomainSignalFromPEK: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "ExtractTimeDomainSignalFromPEK"
    Set fso = Nothing

End Function

Public Function LoadNewPEK(ByVal strPEKFilePath As String, ByVal lngGelIndex As Long, _
                           ByVal MaxFit As Double, _
                           ByVal blnFilterByAbundance As Boolean, _
                           ByVal dblMinAbu As Double, ByVal dblMaxAbu As Double, _
                           ByVal blnMaximumDataCountEnabled As Boolean, ByVal lngMaximumDataCountToLoad As Long, _
                           ByVal eScanFilterMode As eosEvenOddScanFilterModeConstants, _
                           ByVal eDataFilterMode As dfmCSandIsoDataFilterModeConstants) As Long
    '-------------------------------------------------------------------------
    'Returns 0 if data successfuly loaded, -2 if data set is too large,
    '-3 if problems with scan numbers, -4 if no data found, -5 if user cancels load,
    '-6 for file not found or invalid file
    '-7 for file problem that user was already notified about
    '-10 for any other error
    'call this function with huge MaxFit or MaxFit <= 0 to load all values
    'Set blnFilterByAbundance to True to use dblMinAbu and dblMaxAbu to filter the abundance values
    '-------------------------------------------------------------------------
    
    Dim intProgressCount As Integer
    Dim lngReturnValue As Long
    Dim lngIndex As Long
    
    Dim fso As New FileSystemObject
    
    Dim strParFileSetting As String
    
    Dim blnValid As Boolean
    Dim blnSuccess As Boolean
    Dim blnSkipTimeDomainLoad As Boolean
    
    Dim lngByteCountTotal As Long
    Dim lngTotalBytesRead As Long
    
    Dim MaxMZ As Double
    
On Error GoTo LoadNewPEKErrorHandler
    
    ' Update the filter variables
    mGelIndex = lngGelIndex
    mMaxFit = MaxFit
    mFilterByAbundance = blnFilterByAbundance
    mAbundanceMin = dblMinAbu
    mAbundanceMax = dblMaxAbu

    mDataFilterMode = eDataFilterMode
    
    mMaximumDataCountEnabled = blnMaximumDataCountEnabled
    mMaximumDataCountToLoad = lngMaximumDataCountToLoad
    
    If mMaximumDataCountEnabled Then
        If mMaximumDataCountToLoad < 10 Then mMaximumDataCountToLoad = 10
        intProgressCount = 6
    Else
        intProgressCount = 4
    End If

    mCurrentProgressStep = 0
    frmProgress.InitializeForm "Loading data file", mCurrentProgressStep, intProgressCount, False, True, True, MDIForm1
    lngReturnValue = -10
    
    ' Validate that the input file exists
    If Not fso.FileExists(strPEKFilePath) Then
        LoadNewPEK = -6
        Exit Function
    End If
    
    ' Initialize the even/odd scan filter variables
    mEvenOddScanFilter = False
    If eScanFilterMode = eosLoadOddScansOnly Then
        mEvenOddScanFilter = True
        mEvenOddModCompareVal = 1                     ' Use scans where Scan Mod 2 = 1
    ElseIf eScanFilterMode = eosLoadEvenScansOnly Then
        mEvenOddScanFilter = True
        mEvenOddModCompareVal = 0                     ' Use scans where Scan Mod 2 = 0
    End If
    
    ' Initialize the progress bar
    lngTotalBytesRead = 0
    lngByteCountTotal = FileLen(strPEKFilePath)
    frmProgress.InitializeSubtask "Reading data", 0, lngByteCountTotal
    
    mIReportFile = False
    If LookupICR2LSParFileSetting(strPEKFilePath, "chkIreport", strParFileSetting) Then
        If Trim(strParFileSetting) = "1" Then
            mIReportFile = True
        End If
    End If
    
    mIsotopeLabelTagsPresent = False
    
    With GelData(mGelIndex)
        ' Reserve space in the arrays
        ReDim .ScanInfo(SCAN_INFO_DIM_CHUNK)
        
        .LinesRead = 0
        .DataLines = 0
        .CSLines = 0
        .IsoLines = 0
        
        .DataStatusBits = 0
        
        ReDim .CSData(ISO_DATA_DIM_CHUNK)
        ReDim .IsoData(ISO_DATA_DIM_CHUNK)
    End With

    If mMaximumDataCountEnabled Then
        mReadMode = rmReadModeConstants.rmPrescanData
    Else
        mReadMode = rmReadModeConstants.rmStoreData
    End If
    
    mWarnedMissingRequiredIsosColumnHeaders = False
    mWarnedUnknownIsosColumnHeader = False
    mWarnedMissingRequiredCSColumnHeaders = False
    mWarnedUnknownCSColumnHeader = False

    Do While mReadMode < rmReadModeConstants.rmReadComplete
        If mMaximumDataCountEnabled Then
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
            mPrescannedData.MaximumDataCountToLoad = mMaximumDataCountToLoad
            
            mSubtaskMessage = "Pre-scanning PEK file to determine data to load"
        Else
            mSubtaskMessage = "Reading PEK file"
        End If
        frmProgress.InitializeSubtask mSubtaskMessage, 0, lngByteCountTotal
    
        ' Reset the tracking variables
        mValidDataPointCount = 0
        lngTotalBytesRead = 0
        
        lngReturnValue = ReadPEKFile(fso, strPEKFilePath, lngTotalBytesRead)
        If lngReturnValue <> 0 Then
            ' Error occurred
            Debug.Assert False
            LoadNewPEK = lngReturnValue
            Exit Function
        End If
        
        If mReadMode = rmReadModeConstants.rmPrescanData Then
            If KeyPressAbortProcess > 1 Then
                ' User Cancelled Load
                LoadNewPEK = -5
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
       LoadNewPEK = -4
       frmProgress.HideForm
       Exit Function
    End If

    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load; keep the data in memory, but write an entry to the analysis history
        AddToAnalysisHistory mGelIndex, "Warning: File only partially loaded since user cancelled the loading process"
        blnSkipTimeDomainLoad = True
        KeyPressAbortProcess = 1
    End If
    
    mCurrentProgressStep = mCurrentProgressStep + 1
    frmProgress.UpdateProgressBar mCurrentProgressStep
    frmProgress.InitializeSubtask "Initializing data structures", 0, 1

    With GelData(mGelIndex)
         ' Old: .PathtoDataFiles = GetPathWOFileName(CurrDataFName)
         ' New: data file folder path is the folder one folder up from .Filename's folder if .Filename's folder contains _Auto00000
         '      if .Filename's folder does not contain _Auto0000, then simply use .Filename's folder
        .PathtoDataFiles = DetermineParentFolderPath(.FileName)
        
        .DataLines = .IsoLines + .CSLines
        
        If mScanInfoCount > 0 Then
            ReDim Preserve .ScanInfo(mScanInfoCount)
        Else
            ReDim .ScanInfo(0)
        End If
        
        If mIReportFile Then
            .DataStatusBits = .DataStatusBits Or GEL_DATA_STATUS_BIT_IREPORT
        Else
            .DataStatusBits = .DataStatusBits And Not GEL_DATA_STATUS_BIT_IREPORT
        End If
        
        If mIsotopeLabelTagsPresent Then
            .DataStatusBits = .DataStatusBits Or GEL_DATA_STATUS_BIT_ISOTOPE_LABEL_TAG
        Else
            .DataStatusBits = .DataStatusBits And Not GEL_DATA_STATUS_BIT_ISOTOPE_LABEL_TAG
        End If
        
        ' .Pek files cannot have MonoPlusMinus4 data, so clear the bit
        .DataStatusBits = .DataStatusBits And Not GEL_DATA_STATUS_BIT_ADDED_MONOPLUSMINUS4_DATA

        If mMaxElutionTime = 0 Then
            ' Elution time wasn't defined
            ' Define the default elution time to range from 0 to 1
            DefineDefaultElutionTimes GelData(mGelIndex).ScanInfo, 0, 1
        End If
        
        UpdateGelAdjacentScanPointerArrays mGelIndex
        
        ' MonroeMod
        AddToAnalysisHistory mGelIndex, "File Loaded; Charge State Data Points = " & Trim(.CSLines) & "; Isotopic (deconvoluted) Data Points = " & Trim(.IsoLines)
    
    End With

    With GelData(mGelIndex)
        ' Find the minimum and maximum MW, Abundance, and MZ values, and set the filters
        MaxMZ = 0
        If .CSLines + .IsoLines > 0 Then
            .MinMW = glHugeOverExp
            .MaxMW = 0
            .MinAbu = glHugeOverExp
            .MaxAbu = 0
            
            If .CSLines > 0 Then
                ReDim Preserve .CSData(.CSLines)

                For lngIndex = 1 To .CSLines
                    If .CSData(lngIndex).Abundance < .MinAbu Then .MinAbu = .CSData(lngIndex).Abundance
                    If .CSData(lngIndex).Abundance > .MaxAbu Then .MaxAbu = .CSData(lngIndex).Abundance
                        
                    If .CSData(lngIndex).AverageMW > .MaxMW Then
                        .MaxMW = .CSData(lngIndex).AverageMW
                        MaxMZ = .MaxMW
                    End If
                    If .CSData(lngIndex).AverageMW < .MinMW Then .MinMW = .CSData(lngIndex).AverageMW
                Next lngIndex
            Else
                ReDim .CSData(0)
            End If
            
            If .IsoLines > 0 Then
                ReDim Preserve .IsoData(.IsoLines)
                
                For lngIndex = 1 To .IsoLines
                    If .IsoData(lngIndex).Abundance < .MinAbu Then .MinAbu = .IsoData(lngIndex).Abundance
                    If .IsoData(lngIndex).Abundance > .MaxAbu Then .MaxAbu = .IsoData(lngIndex).Abundance
                        
                    FindMWExtremes .IsoData(lngIndex), .MinMW, .MaxMW, MaxMZ
                Next lngIndex
            Else
                ReDim .IsoData(0)
            End If
    
        Else
            ReDim .IsoData(0)
            ReDim .CSData(0)
            .MinAbu = 0
            .MaxAbu = 0
            .MinMW = 0
            .MaxMW = 0
        End If
        
        .DataFilter(fltCSAbu, 2) = .MaxAbu             'put initial filters on max
        .DataFilter(fltIsoAbu, 2) = .MaxAbu
        .DataFilter(fltCSMW, 2) = .MaxMW
        .DataFilter(fltIsoMW, 2) = .MaxMW
        .DataFilter(fltIsoMZ, 2) = MaxMZ
        
        .DataFilter(fltEvenOddScanNumber, 0) = False
        .DataFilter(fltEvenOddScanNumber, 1) = 0       ' Show all scan numbers
        
        .DataFilter(fltIsoCS, 2) = 1000                'maximum charge state
    End With
    
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load
        LoadNewPEK = -5
        frmProgress.HideForm
        Exit Function
    Else
        mCurrentProgressStep = mCurrentProgressStep + 1
        frmProgress.UpdateProgressBar mCurrentProgressStep
        frmProgress.InitializeSubtask "Sorting isotopic data", 0, GelData(mGelIndex).IsoLines
    End If
    
    ' Sort the data
    SortIsotopicData mGelIndex
    
    If (GelData(mGelIndex).DataStatusBits And GEL_DATA_STATUS_BIT_IREPORT) = GEL_DATA_STATUS_BIT_IREPORT Then
        ' Fix the mono plus 2 abundance values
        FixIsosMonoPlus2Abu mGelIndex
    End If
    
    If Not blnSkipTimeDomainLoad Then
        ' Extract the time domain signals from the .PEK file
        mCurrentProgressStep = mCurrentProgressStep + 1
        frmProgress.UpdateProgressBar mCurrentProgressStep
        blnSuccess = ExtractTimeDomainSignalFromPEK(strPEKFilePath, mGelIndex)
        Debug.Assert blnSuccess
    End If
    
    LoadNewPEK = 0
    frmProgress.HideForm
    Exit Function

LoadNewPEKErrorHandler:
    Select Case Err.Number
    Case 9
         ' Does this error occur?  If yes, I should figure out why and prevent it from happening
         Debug.Assert False
         Resume Next
    Case Else
         LogErrors Err.Number, "LoadNewPEK"
    End Select
    LoadNewPEK = -10
End Function

Private Function ReadPEKFile(ByRef fso As FileSystemObject, ByVal strPEKFilePath As String, ByRef lngTotalBytesRead As Long) As Long
    Dim tsInFile As TextStream
    Dim strLineIn As String

    Dim lngLinesRead As Long

    Dim blnAutoNumberScans As Boolean
    Dim objScanNumberTracker As clsScanNumberTracker

    Dim LineType As Integer
    Dim Special As String
    Dim CalArgCnt As Long

    ' The following holds the data read from the .Pek file (varData(0) through varData(intDataColumnCount-1))
    ' It is a variant array to allow for both text and numbers
    Dim intDataColumnCount As Integer
    Dim varData(0 To ISONUM_FIELD_COUNT_MAX) As Variant

    Dim CurrDataFName As String
    Dim lngScanNumber As Long
    Dim CurrDataElutionTime As Single

    Dim CurrDataTIC As Double
    Dim CurrDataBPI As Double
    Dim CurrDataBPImz As Double

    Dim sngAbundance As Single
    Dim sngFit As Single
    
    Dim strResponse As String
    Dim eResponse As VbMsgBoxResult
    
    Dim blnValidDataPoint As Boolean
    Dim blnStoreDataPoint As Boolean
    
    Dim lngReturnValue As Long
    
    Dim intIsosColumnMapping() As Integer
    Dim intCSColumnMapping() As Integer
   
On Error GoTo ReadPEKFileErrorHandler
    
    ReDim intIsosColumnMapping(ISOS_DATA_COLUMN_COUNT) As Integer
    ReDim intCSColumnMapping(CS_DATA_COLUMN_COUNT) As Integer
    
    ' Initialize intIsosColumnMapping to the default column mappings
    SetDefaultIsosColumnMapping intIsosColumnMapping
    SetDefaultCSColumnMapping intCSColumnMapping
    
    ' Initialize control variables
    lngLinesRead = 0
    mScanInfoCount = 0
    CalArgCnt = 0
    
    'initialize control variables for LineNow procedure
    ThisLine = LINE_NOTHING
    CalibrationIn = False
    
    DatabaseIn = False
    MediaTypeIn = False
    
    blnAutoNumberScans = False
    Set objScanNumberTracker = New clsScanNumberTracker
    
    CurrDataElutionTime = 0
    mMaxElutionTime = 0
        
    Set tsInFile = fso.OpenTextFile(strPEKFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream
        
        strLineIn = tsInFile.ReadLine
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
          
        If lngLinesRead Mod 100 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngTotalBytesRead
            If KeyPressAbortProcess > 1 Then Exit Do
        End If
        
        lngLinesRead = lngLinesRead + 1
        
        If mReadMode = rmReadModeConstants.rmStoreData Then
            GelData(mGelIndex).LinesRead = lngLinesRead
        End If
        
        LineNow strLineIn, LineType, Special, varData, intDataColumnCount, intIsosColumnMapping, intCSColumnMapping, strPEKFilePath
        Select Case LineType
        Case LINE_FILENAME_AKA_SCAN_NUMBER
            CurrDataFName = varData(0)
            
            If blnAutoNumberScans Then
                varData(1) = objScanNumberTracker.GetNextAutoNumberedScan()
            Else
                ' Update the average increment value
                objScanNumberTracker.AddScanNumberAndUpdateAverageIncrement CLng(varData(1))
            End If
           
            If lngScanNumber > varData(1) Then     'cannot accept non-ascending scan numbers
                If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Or mReadMode = rmReadModeConstants.rmPrescanData Then
                    ' Assume auto-numbering is OK
                    eResponse = vbYes
                Else
                    eResponse = MsgBox("Error in scan order found after scan " & lngScanNumber & ".  Next scan number is " & varData(1) & ". Choose Yes to auto-number remaining spectra sequentially.  Choose No to keep the data loaded up to this point.  Choose Cancel to abort loading.", vbYesNoCancel + vbDefaultButton1, glFGTU)
                End If
                
                If eResponse = vbCancel Then
                    ReadPEKFile = -3              'cancel read operation
                    frmProgress.HideForm
                    Exit Function
                ElseIf eResponse = vbNo Then
                    Exit Do                   'ignore the remaining scans
                Else
                    ' Auto-number from now on
                    If Not blnAutoNumberScans Then
                        objScanNumberTracker.SetAutoNumberIncrementToCurrentAverage
                        
                        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                            ' Query the user to confirm the auto-number increment value
                            strResponse = InputBox("Please confirm the value to increment each scan number by when auto-numbering.  This value has been auto-computed based on the scans read so far.  Note that decimal numbers are typical for LTQ-FT or LTQ-Orbitrap data.", "Increment value", Round(objScanNumberTracker.AutoNumberIncrement, 4))
                            If Len(strResponse) > 0 Then
                                If IsNumeric(strResponse) Then
                                    objScanNumberTracker.AutoNumberIncrement = CSngSafe(strResponse)
                                End If
                            End If
                        End If
                        
                        AddToAnalysisHistory mGelIndex, "Non-ascending scan number found.  Auto-numbering sequentially starting with scan " & Trim(Round(lngScanNumber + objScanNumberTracker.AutoNumberIncrement, 0)) & "; Increment value = " & Trim(objScanNumberTracker.AutoNumberIncrement)
                        blnAutoNumberScans = True
                    End If
                    varData(1) = objScanNumberTracker.GetNextAutoNumberedScan()
                End If
            End If
           
            lngScanNumber = varData(1)
            If mReadMode = rmReadModeConstants.rmStoreData Then
                With GelData(mGelIndex)
                    If mScanInfoCount > 0 Then
                        ' Store the TIC and BPI information in the previous scan
                        With .ScanInfo(mScanInfoCount)
                            .TIC = CurrDataTIC
                            .BPI = CurrDataBPI
                            .BPImz = CurrDataBPImz
                        End With
                    End If
                    CurrDataTIC = 0
                    CurrDataBPI = 0
                    CurrDataBPImz = 0
                    
                    mScanInfoCount = mScanInfoCount + 1
                    If mScanInfoCount > UBound(.ScanInfo) Then
                       ReDim Preserve .ScanInfo(UBound(.ScanInfo) + SCAN_INFO_DIM_CHUNK)
                    End If
                    
                    With .ScanInfo(mScanInfoCount)
                       .ScanNumber = lngScanNumber
                       .ElutionTime = CurrDataElutionTime
                       .ScanType = 1
                       .ScanFileName = fso.GetFileName(CurrDataFName)
                    End With
                End With
            End If
            
       Case LINE_FREQUENCY  'frequency shifts
            If mReadMode = rmReadModeConstants.rmStoreData Then
                GelData(mGelIndex).ScanInfo(mScanInfoCount).FrequencyShift = varData(0)
                ''GelData(mGelIndex).DFFS(mScanInfoCount) = varData(0)
            End If
       Case LINE_INTENSITY  ' MonrodMod: Storing Time Domain Signal Level here (found in function ExtractTimeDomainSignalFromPEK)
            If mReadMode = rmReadModeConstants.rmStoreData Then
                GelData(mGelIndex).ScanInfo(mScanInfoCount).TimeDomainSignal = varData(0)
                ''GelData(mGelIndex).DFIN(mScanInfoCount) = varData(0)
            End If
       Case LINE_DATABASE                         'this can happen only once
            ' No longer supported (March 2006)
            ''GelData(mGelIndex).PathtoDatabase = varData(0)
       Case LINE_MEDIA
            If mReadMode = rmReadModeConstants.rmStoreData Then
                GelData(mGelIndex).MediaType = varData(0)
            End If
       Case LINE_EQUATION
            If mReadMode = rmReadModeConstants.rmStoreData Then
                If Not CalibrationIn Then    'read calibration only once
                   GelData(mGelIndex).CalEquation = varData(0)
                   CalibrationIn = True
                Else
                   CalibrationIn = False
                End If
            End If
       Case LINE_CAL_ARGUMENT
            If mReadMode = rmReadModeConstants.rmStoreData Then
                If CalibrationIn Then        'read only once(looks strange but it works)
                   CalArgCnt = CalArgCnt + 1
                   If CalArgCnt <= 10 Then GelData(mGelIndex).CalArg(CalArgCnt) = varData(0)
                End If
            End If
       Case LINE_DATA_CS
            If Not mEvenOddScanFilter Or (lngScanNumber Mod 2 = mEvenOddModCompareVal) Then
              
                sngAbundance = GetColumnValueSng(varData, intCSColumnMapping(crdCSRawDataIndex.crdAbu), -1)
              
                blnValidDataPoint = True
                If mFilterByAbundance Then
                    If sngAbundance > -1 Then
                        If sngAbundance < mAbundanceMin Or sngAbundance > mAbundanceMax Then blnValidDataPoint = False
                    Else
                        blnValidDataPoint = False
                    End If
                End If
                
                If mDataFilterMode = dfmLoadIsoDataOnly Then
                    ' Skip CS data
                    blnValidDataPoint = False
                End If
                
                If blnValidDataPoint Then
                    If mReadMode = rmReadModeConstants.rmPrescanData Then
                        sngAbundance = GetColumnValueSng(varData, intCSColumnMapping(crdCSRawDataIndex.crdAbu), -1)
                        mPrescannedData.AddDataPoint sngAbundance, mValidDataPointCount
                    Else
                        If mMaximumDataCountEnabled Then
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
                                .CSLines = .CSLines + 1
                                
                                If .CSLines > UBound(.CSData) Then
                                    ReDim Preserve .CSData(UBound(.CSData) + ISO_DATA_DIM_CHUNK)
                                End If

                                With .CSData(.CSLines)
                                    .ScanNumber = lngScanNumber
                                    
                                    .Charge = CInt(GetColumnValueLng(varData, intCSColumnMapping(crdCSRawDataIndex.crdCS), 1))
                                    .ChargeCount = CInt(GetColumnValueLng(varData, intCSColumnMapping(crdCSRawDataIndex.crdNumberOfCS), 1))
                                    .Abundance = sngAbundance
                                    .Fit = 0
                                    .AverageMW = GetColumnValueDbl(varData, intCSColumnMapping(crdCSRawDataIndex.crdAverageMW), 0)
                                    .MonoisotopicMW = 0
                                    .MostAbundantMW = 0
                                    .MassStDev = GetColumnValueDbl(varData, intCSColumnMapping(crdCSRawDataIndex.crdMWStDev), 0)
                                    
                                    .MZ = 0
                                    
                                    .FWHM = 0
                                    .SignalToNoise = 0
                                    .IntensityMono = 0
                                    .IntensityMonoPlus2 = 0

                                    ''If IsNumeric(varData(6)) Then .IsotopicFitRatio = varData(6)  'Now holds ratio of N14/N15; Legacy: expected mw
                                    ''If IsNumeric(varData(7)) Then .IsotopicAtomCount = varData(7)  'Legacy: DB match mass error
''                                    If Not IsNull(varData(crdMTID)) Then
''                                        If Len(CStr(varData(crdMTID))) > 0 Then .MTID = varData(crdMTID)
''                                    End If
''
                                    ' Update the TIC and BPI data
                                    CurrDataTIC = CurrDataTIC + .Abundance
                                    If .Abundance > CurrDataBPI Then
                                        CurrDataBPI = .Abundance
                                        CurrDataBPImz = .AverageMW
                                    End If
                                End With
                            End With
                        End If
                    End If
                    
                    mValidDataPointCount = mValidDataPointCount + 1
                End If
            End If
       Case LINE_DATA_ISO   'data line - isotopic
            ' Possibly filter on Fit
            sngFit = GetColumnValueSng(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdIsoFit), 0)
            
            If sngFit <= mMaxFit Or mMaxFit <= 0 Then
                If Not mEvenOddScanFilter Or (lngScanNumber Mod 2 = mEvenOddModCompareVal) Then
                    
                    sngAbundance = GetColumnValueDbl(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdAbu), -1)
                    
                    blnValidDataPoint = True
                    If mFilterByAbundance Then
                        If sngAbundance > -1 Then
                            If sngAbundance < mAbundanceMin Or sngAbundance > mAbundanceMax Then blnValidDataPoint = False
                        Else
                            blnValidDataPoint = False
                        End If
                    End If
                    
                    If mDataFilterMode = dfmLoadCSDataOnly Then
                        ' Skip Iso data
                        blnValidDataPoint = False
                    End If
                    
                    If blnValidDataPoint Then
                        If mReadMode = rmReadModeConstants.rmPrescanData Then
                            sngAbundance = GetColumnValueSng(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdAbu), -1)
                            mPrescannedData.AddDataPoint sngAbundance, mValidDataPointCount
                        Else
                            If mMaximumDataCountEnabled Then
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
                                    .IsoLines = .IsoLines + 1
                                        
                                    If .IsoLines > UBound(.IsoData) Then
                                        ReDim Preserve .IsoData(UBound(.IsoData) + ISO_DATA_DIM_CHUNK)
                                    End If
                                    
                                    With .IsoData(.IsoLines)
                                        .ScanNumber = lngScanNumber
                                        
                                        .Charge = CInt(GetColumnValueLng(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdCS), 1))
                                        .Abundance = sngAbundance
                                        .MZ = GetColumnValueDbl(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdMZ), 0)
                                        .Fit = sngFit
                                        .AverageMW = GetColumnValueDbl(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdMWavg), 0)
                                        .MonoisotopicMW = GetColumnValueDbl(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdMWMono), 0)
                                        .MostAbundantMW = GetColumnValueDbl(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdMWTMA), 0)
                                        .FWHM = 0
                                        .SignalToNoise = 0
                                        
                                        .IsotopeLabel = iltIsotopeLabelTagConstants.iltNone
                                        
                                        If intIsosColumnMapping(irdIsoRawDataIndex.irdIsotopeTag) >= 0 Or _
                                           intIsosColumnMapping(irdIsoRawDataIndex.irdIRepMWMonoAbu) >= 0 Then
                                            ' IReport values are defined
                                            .IntensityMono = GetColumnValueSng(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdIRepMWMonoAbu), 0)
                                            .IntensityMonoPlus2 = GetColumnValueSng(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdIRep2Da), 0)
                                        Else
                                            .IntensityMono = 0
                                            .IntensityMonoPlus2 = 0
                                            
                                            If intIsosColumnMapping(irdIsoRawDataIndex.irdIsotopeTag) > 0 Then
                                                .IsotopeLabel = GetIsotopeLabelTagCode(GetColumnValueString(varData, intIsosColumnMapping(irdIsoRawDataIndex.irdIsotopeTag), ""))
                                            Else
                                                ' See if vardata(7) contains some text
                                            
                                                If Not IsNumeric(varData(7)) And Not IsNull(varData(7)) Then
                                                    ' Assume varData(7) contains an isotope label
                                                    .IsotopeLabel = GetIsotopeLabelTagCode(CStr(varData(7)))

                                                    If .IsotopeLabel <> iltNone Then
                                                        If IsNumeric(varData(8)) And Not IsNull(varData(8)) Then
                                                            ' Assume varData(8) contains the ratio of N14 to N15
                                                            '.IsotopicFitRatio = CSng(varData(8))
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            
                                            If .IsotopeLabel <> iltIsotopeLabelTagConstants.iltNone Then
                                                ' Isotope tag found; set mIsotopeLabelTagsPresent to true
                                                mIsotopeLabelTagsPresent = True
                                            End If

                                        End If
                                        
                                        
                                 

                                        '' Old method used before switch to intIsosColumnMapping()
                                        ''    If mIReportFile Or IsNumeric(varData(irdERorN14N15OrIRepMWMonoAbu)) And IsNumeric(varData(irdN14N15RatioOrIRep2Da)) Then
                                        ''        ' Assume IReport data
                                        ''        If IsNumeric(varData(irdERorN14N15OrIRepMWMonoAbu)) Then
                                        ''            .IntensityMono = varData(irdERorN14N15OrIRepMWMonoAbu)
                                        ''            mIReportFile = True
                                        ''        End If
                                        ''
                                        ''        If IsNumeric(varData(irdN14N15RatioOrIRep2Da)) Then
                                        ''            .IntensityMonoPlus2 = varData(irdN14N15RatioOrIRep2Da)
                                        ''            mIReportFile = True
                                        ''        End If
                                        ''    Else
                                        ''        If Not IsNumeric(varData(irdERorN14N15OrIRepMWMonoAbu)) Then
                                        ''            ' Assume irdERorN14N15OrIRepMWMonoAbu contains an isotope label
                                        ''            '.IsotopeLabel = varData(irdERorN14N15OrIRepMWMonoAbu)        'Now holds ratio of N14/N15; Legacy: expected mw
                                        ''
                                        ''            ' Assume varData(irdN14N15RatioOrIRep2Da) contains .IsotopicFitRatio data
                                        ''            If IsNumeric(varData(irdN14N15RatioOrIRep2Da)) Then
                                        ''                '.IsotopicFitRatio = varData(irdN14N15RatioOrIRep2Da)        'Now holds ratio of N14/N15; Legacy: expected mw
                                        ''            End If
                                        ''        End If
                                        ''    End If
                                        ''
                                        ''    If Not IsNull(varData(irdPeptideIdentity)) Then
                                        ''        If Len(CStr(varData(irdPeptideIdentity))) > 0 Then .MTID = varData(irdPeptideIdentity)
                                        ''    End If
                                                                            
                                        ' Update the TIC and BPI data
                                        CurrDataTIC = CurrDataTIC + .Abundance
                                        If .Abundance > CurrDataBPI Then
                                            CurrDataBPI = .Abundance
                                            CurrDataBPImz = .MZ
                                        End If
                                        
                                    End With
                                End With
                            End If
                        End If
                        
                        mValidDataPointCount = mValidDataPointCount + 1
                    End If
                 
                End If
            End If
    ''    Case LINE_DELTA
    ''       Deltas(mScanInfoCount).Delta = varData(0)
    ''    Case LINE_DELTA_TOLERANCE
    ''       Deltas(mScanInfoCount).Tolerance = varData(0)
    ''    Case LINE_DELTA_TAGMASS
    ''       Deltas(mScanInfoCount).TagMass = varData(0)
    ''    Case LINE_DELTA_MAX
    ''       Deltas(mScanInfoCount).MaxDeltas = varData(0)
    ''    Case LINE_DATA_DD
    ''       DDCnt = DDCnt + 1
    ''       If Deltas(mScanInfoCount).MinInd < 0 Then Deltas(mScanInfoCount).MinInd = DDCnt
    ''       Deltas(mScanInfoCount).MaxInd = DDCnt
    ''       'read data here
    ''       DDFNs(DDCnt) = GelData(mGelIndex).ScanInfo(mScanInfoCount).ScanNumber
    ''       DDMWs(DDCnt) = varData(0)
    ''       DDD(DDCnt) = varData(1)
    ''       DDRatio(DDCnt) = varData(2)
        Case LINE_RETENTION_TIME
            If IsNumeric(varData(0)) Then
                CurrDataElutionTime = varData(0)
                If CurrDataElutionTime > mMaxElutionTime Then mMaxElutionTime = CurrDataElutionTime
            End If
        Case LINE_NUMBER_OF_PEAKS
            If mReadMode = rmReadModeConstants.rmStoreData Then
                If IsNumeric(varData(0)) Then
                    GelData(mGelIndex).ScanInfo(mScanInfoCount).NumPeaks = CLng(varData(0))
                End If
            End If
        Case LINE_NUMBER_OF_ISOTOPIC_DISTRIBUTIONS
            If mReadMode = rmReadModeConstants.rmStoreData Then
                If IsNumeric(varData(0)) Then
                    GelData(mGelIndex).ScanInfo(mScanInfoCount).NumDeisotoped = CLng(varData(0))
                End If
            End If
       End Select
    Loop
    
    tsInFile.Close
    ReadPEKFile = 0
    Exit Function

ReadPEKFileErrorHandler:
    If Err.Number = 94 Then
        ' Invalid use of Null; this shouldn't normally happen
        Debug.Assert False
        Resume Next
    End If

    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadPEKFile"
    
    On Error Resume Next
    tsInFile.Close
    
    ReadPEKFile = lngReturnValue
    
End Function

Private Function LookupICR2LSParFileSetting(strInputFilePath As String, strSettingNameToLookup As String, ByRef strSettingValue As String) As Boolean
    ' Looks for a .Par file with the same name as strInputFilePath, but ending in .Par
    ' If found, opens it and looks for strSettingNameToLookup
    ' If found, returns True and returns the setting value in strSettingValue (even if it's a number)
    '
    ' If the .Par file isn't found, or the setting isn't found, or an error occurs then returns False
    
    Dim fso As FileSystemObject
    Dim tsParFile As TextStream
    
    Dim strParFilePath As String
    Dim strLineIn As String
    Dim strKeyName As String, strKeyValue As String
    
    Dim lngCharLoc As Long
    
    Dim blnSettingFound As Boolean
    
On Error GoTo LookupICR2LSParFileSettingErrorHandler

    blnSettingFound = False
    strSettingValue = ""
    
    Set fso = New FileSystemObject
    strParFilePath = fso.GetParentFolderName(strInputFilePath)
    
    strParFilePath = fso.BuildPath(strParFilePath, fso.GetBaseName(strInputFilePath) & ".par")
    
    If fso.FileExists(strParFilePath) Then
        Set tsParFile = fso.OpenTextFile(strParFilePath, ForReading)
        
        Do While Not tsParFile.AtEndOfStream
            strLineIn = tsParFile.ReadLine()
            
            lngCharLoc = InStr(strLineIn, "=")
            If lngCharLoc > 0 Then
                strKeyName = Trim(Left(strLineIn, lngCharLoc - 1))
                strKeyValue = Trim(Mid(strLineIn, lngCharLoc + 1))
                
                If LCase(strKeyName) = LCase(strSettingNameToLookup) Then
                    strSettingValue = strKeyValue
                    blnSettingFound = True
                    Exit Do
                End If
            End If
        Loop
        tsParFile.Close
        Set tsParFile = Nothing
        
    End If
    
    Set fso = Nothing

    LookupICR2LSParFileSetting = blnSettingFound
    Exit Function

LookupICR2LSParFileSettingErrorHandler:
    Debug.Print "Error in LookupICR2LSParFileSetting: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "Module1->LookupICR2LSParFileSetting"
    LookupICR2LSParFileSetting = False

End Function

Private Sub LineNow(ByVal strLineIn As String, ByRef LineType As Integer, ByRef Special As String, _
                    ByRef varData As Variant, ByRef intDataColumnCount As Integer, _
                    ByRef intIsosColumnMapping() As Integer, ByRef intCSColumnMapping() As Integer, _
                    ByVal strFilePath As String)
                    
    'strLineIn is line, LineType returns type of line, Special string to be carried back
    'varData is array of variants that returns actual values
    Dim k As Integer
    On Error GoTo err_LineNow
    
    intDataColumnCount = 0
    For k = 0 To UBound(varData)
        varData(k) = Null
    Next k
    
    Special = ""
    
    Select Case Left(strLineIn, 8)
    Case t8FILENAME
       varData(0) = Right(Trim(strLineIn), Len(strLineIn) - 9)     'File name
       varData(1) = ExtractScanNumberFromFilenameLine(CStr(varData(0)))
       LineType = LINE_FILENAME_AKA_SCAN_NUMBER
    Case t8FREQSHIFT
       varData(0) = GetNumberEqual(strLineIn)
       If Len(varData(0)) > 0 Then
          varData(0) = val(varData(0))
       Else
          varData(0) = 0
       End If
       LineType = LINE_FREQUENCY
    Case t8DATABASE
       If Not DatabaseIn Then
          varData(0) = Right(Trim(strLineIn), Len(strLineIn) - 9)
          LineType = LINE_DATABASE
          If Len(varData(0)) > 0 Then DatabaseIn = True
       Else
          LineType = LINE_NOTHING
       End If
    Case t8MEDIATYPE
       If Not MediaTypeIn Then
          varData(0) = Trim$(Right(Trim(strLineIn), Len(strLineIn) - 11))
          LineType = LINE_MEDIA
          MediaTypeIn = True
       Else
          LineType = LINE_NOTHING
       End If
    Case t8DATA_CS
        ThisLine = LINE_DATA_CS
        LineType = LINE_NOTHING
        ParseCSColumnHeaders strLineIn, intCSColumnMapping, strFilePath
        
    Case t8DATA_ISO, t8DATA_ISO_Tabbed
        ThisLine = LINE_DATA_ISO
        LineType = LINE_NOTHING
        ParseIsosColumnHeaders strLineIn, intIsosColumnMapping, strFilePath
    
    Case t8CALIBRATION1, t8CALIBRATION2
       If Not CalibrationIn Then
          ThisLine = LINE_CALIBRATION
          LineType = LINE_NOTHING
       End If
    ''Case t8DELTA
    ''   varData(0) = GetNumberEqual(strLineIn)      'delta
    ''   If Len(varData(0)) > 0 Then
    ''      varData(0) = Val(varData(0))
    ''   Else
    ''      varData(0) = -1
    ''   End If
    ''   LineType = LINE_DELTA
    ''Case t8DELTA_TOL, t8DELTA_TOL_1
    ''   varData(0) = GetNumberEqual(strLineIn)      'delta
    ''   If Len(varData(0)) > 0 Then
    ''      varData(0) = Val(varData(0))
    ''   Else
    ''      varData(0) = -1
    ''   End If
    ''   LineType = LINE_DELTA_TOLERANCE
    ''Case t8TAG_MASS
    ''   varData(0) = GetNumberEqual(strLineIn)      'delta
    ''   If Len(varData(0)) > 0 Then
    ''      varData(0) = Val(varData(0))
    ''   Else
    ''      varData(0) = -1
    ''   End If
    ''   LineType = LINE_DELTA_TAGMASS
    ''Case t8MAX_DELTAS
    ''   varData(0) = GetNumberEqual(strLineIn)      'delta
    ''   If Len(varData(0)) > 0 Then
    ''      varData(0) = Val(varData(0))
    ''   Else
    ''      varData(0) = -1
    ''   End If
    ''   LineType = LINE_DELTA_MAX
    ''Case t8DELTADATABLOCK
    ''   ThisLine = LINE_DATA_DD
    ''   LineType = LINE_NOTHING
    Case Else
        If Left(strLineIn, 4) = t4RT Then
            varData(0) = GetNumberEqual(strLineIn)      ' Retention time
            If Len(varData(0)) > 0 Then
                varData(0) = val(varData(0))
                LineType = LINE_RETENTION_TIME
            Else
                varData(0) = 0
                LineType = LINE_NOTHING
            End If
        
        ElseIf Left(strLineIn, Len(PEK_D_PEAKS_CNT)) = PEK_D_PEAKS_CNT Then
            varData(0) = GetNumberEqual(strLineIn)      ' Number of peaks
            If Len(varData(0)) > 0 Then
                varData(0) = val(varData(0))
                LineType = LINE_NUMBER_OF_PEAKS
            Else
                varData(0) = 0
                LineType = LINE_NOTHING
            End If
        
        ElseIf Left(strLineIn, Len(PEK_D_IS_CNT)) = PEK_D_IS_CNT Then
            varData(0) = GetNumberEqual(strLineIn)      ' Number of isotopic distributions
            If Len(varData(0)) > 0 Then
                varData(0) = val(varData(0))
                LineType = LINE_NUMBER_OF_ISOTOPIC_DISTRIBUTIONS
            Else
                varData(0) = 0
                LineType = LINE_NOTHING
            End If
    
        Else
            Select Case ThisLine
            Case LINE_DATA_CS
                 If IsDataLine(Trim$(strLineIn), varData, intDataColumnCount, Special, strFilePath) Then
                    If intDataColumnCount >= 5 Then
                       LineType = LINE_DATA_CS
                    Else
                       LineType = LINE_NOTHING
                    End If
                    ' The 6th entry in a charge state based line of data could be various things
                    ' In a legacy .PEK file, it could be a number, in which case we used to append "M" to Special; now, we ignore it
                    ' If the 6th entry isn't a number, then we used to append just the first letter of the entry to Special
                    ' We now ignore the 8th entry in this function and deal with it in the LoadNewPEK function instead
                 Else
                    LineType = LINE_NOTHING
                    ThisLine = LINE_NOTHING
                 End If
            Case LINE_DATA_ISO
                 If IsDataLine(Trim$(strLineIn), varData, intDataColumnCount, Special, strFilePath) Then
                    If intDataColumnCount >= 7 Then
                       LineType = LINE_DATA_ISO
                    Else
                       LineType = LINE_NOTHING
                    End If
                    ' The 8th entry in a isotopic based line of data could be various things
                    ' In a legacy .PEK file, it could be a number, in which case we append "M" to Special
                    ' If the 8th entry wasn't a number, then we used to append just the first letter of the entry to Special
                    ' We now ignore the 8th entry in this function and deal with it in the LoadNewPEK function instead
                 Else
                    LineType = LINE_NOTHING
                    ThisLine = LINE_NOTHING
                 End If
            Case LINE_CALIBRATION
                 Select Case Left$(Trim$(strLineIn), 3)
                 Case t3EQUATION
                      varData(0) = Trim$(strLineIn)
                      LineType = LINE_EQUATION
                 Case t3ARG_A, t3ARG_B, t3ARG_C, t3ARG_D, t3ARG_E, _
                      t3ARG_F, t3ARG_G, t3ARG_H, t3ARG_H, t3ARG_I
                      varData(0) = Trim$(Right$(Trim$(strLineIn), Len(Trim$(strLineIn)) - 3))
                      If Not IsNumeric(varData(0)) Then varData(0) = 0
                      LineType = LINE_CAL_ARGUMENT
                 Case Else
                      ThisLine = LINE_NOTHING
                      LineType = LINE_NOTHING
                 End Select
    ''        Case LINE_DATA_DD
    ''             If IsDataLine(Trim$(strLineIn), varData, intDataColumnCount, Special, strFilePath) Then
    ''                If intDataColumnCount = 3 Then
    ''                   LineType = LINE_DATA_DD
    ''                Else
    ''                   LineType = LINE_NOTHING
    ''                End If
    ''             Else
    ''                LineType = LINE_NOTHING
    ''                ThisLine = LINE_NOTHING
    ''             End If
            Case Else
                 ThisLine = LINE_NOTHING
                 LineType = LINE_NOTHING
            End Select
        
        End If
       
    End Select
    Exit Sub

err_LineNow: 'ignore the line
    LineType = LINE_NOTHING
End Sub

Public Function ExtractScanNumberFromFilenameLine(ByVal strLine As String, Optional ByRef strTextBeforeScanNumber As String = "", Optional ByRef strTextAfterScanNumber As String = "") As Long

    ' strLine contains the scan number (aka the file number)
    ' Typical forms for strLine are the following (for scan 3275)
    '  1) ICR Data: "Filename: C:\DMS_ICR_WorkDir2\APQ_3Conc_A-2_20Nov03_Andro_0929-13\s001\APQ_3Conc_.03275"
    '  2) Analysis Wiff data: "Filename: C:\DMS_ICR_WorkDir1\ShewSO407_8December03_Saturn-C2.wiff-3275:1"
    '  3) CDF Data: "Filename: D:\ethan\CDF\BSA19JAN04.CDF - 3275"
    '  4) MassLynx Data: "Filename: D:\ethan\MassLynx\BSA19JAN04.Raw - 3275"
    '
    ' The calling function should have already removed "Filename: " from strLine before calling this function, though that isn't a requirement
    
    Dim StartPos As Long
    Dim lngCharLoc As Long
    Dim intAscVal As String
    
    Dim strScanNumber As String
    Dim lngScanNumber As Long
    
On Error GoTo ExtractScanNumberFromFilenameLineErrorHandler

    strTextBeforeScanNumber = ""
    strTextAfterScanNumber = ""
    
    strLine = Trim(strLine)
    If InStr(LCase(strLine), t5WIFF) > 0 Then
        ' Line contains .wiff- and is thus of the form
        '  ?????wiff-123:1
        ' The scan number is the number before the colon
        ' Thus, we'll remove the :1 from strLine
        StartPos = InStr(LCase(strLine), t5WIFF) + Len(t5WIFF)
        lngCharLoc = InStr(StartPos, strLine, ":")
        If lngCharLoc > 0 Then
            strTextAfterScanNumber = Mid(strLine, lngCharLoc)
            strLine = Left(strLine, lngCharLoc - 1)
        Else
            ' Colon not found; don't do anything
        End If
    End If
        
    lngScanNumber = 0
    strScanNumber = ""
    
    ' Now find the longest contiguous set of numbers at the end of strLine
    lngCharLoc = Len(strLine)
    Do While lngCharLoc > 0
        If IsNumeric(Mid(strLine, lngCharLoc, 1)) Then
            strScanNumber = Mid(strLine, lngCharLoc, 1) & strScanNumber
            lngCharLoc = lngCharLoc - 1
        Else
            Exit Do
        End If
    Loop
    
    If lngCharLoc > 0 Then
        strTextBeforeScanNumber = Left(strLine, lngCharLoc)
    End If
    
    If IsNumeric(strScanNumber) Then
        lngScanNumber = CLng(strScanNumber)
    End If
    
    ExtractScanNumberFromFilenameLine = lngScanNumber
    Exit Function

ExtractScanNumberFromFilenameLineErrorHandler:
    Debug.Print "Error in ExtractScanNumberFromFilenameLine: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "Module1->ExtractScanNumberFromFilenameLine"
    Resume Next
    
End Function

Private Sub LineNowScanNumberOrTimeDomainSignal(ByVal strThisLine As String, ByRef intLineType As Integer, ByRef varValue As Variant)
' This function only looks for lines beginning with "Time domain signal level:" or "Filename:"
' It is used by Sub ExtractTimeDomainSignalFromPEK
' strThisLine is line, intLineType returns type of line,
' varValue is variant that returns actual values

On Error GoTo LineNowScanNumberOrTimeDomainSignalErrorHandler

Select Case Left(strThisLine, 8)
Case t8FILENAME
    varValue = ExtractScanNumberFromFilenameLine(strThisLine)
    intLineType = LINE_FILENAME_AKA_SCAN_NUMBER
Case t8TIME_DOMAIN
    strThisLine = Trim(Mid(strThisLine, Len(PEK_D_TIME_DOMAIN_SIGNAL_LEVEL) + 1))
    If Left(strThisLine, 1) = vbTab Then strThisLine = Mid(strThisLine, 2)
    If IsNumeric(strThisLine) Then
        varValue = val(strThisLine)     ' Time domain signal level
    Else
        varValue = 0
    End If
    intLineType = LINE_TIME_DOMAIN_SIGNAL
Case Else
    intLineType = LINE_NOTHING
End Select

Exit Sub

LineNowScanNumberOrTimeDomainSignalErrorHandler:
    Debug.Print "Error in LineNowScanNumberOrTimeDomainSignal: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "LineNowScanNumberOrTimeDomainSignal"
    intLineType = LINE_WHATEVER
    
End Sub


Private Function IsDataLine(ByVal strLineIn As String, ByRef varData As Variant, ByRef intDataColumnCount As Integer, ByRef Special As String, ByVal strFilePath As String) As Boolean
    Dim k As Integer, i As Integer
    Dim LineElement As Variant
    Dim TmpLine As String
    Dim Done As Boolean
    
    Static blnMaxColCountReachedErrorLogged As Boolean
    
    intDataColumnCount = 0
    
    TmpLine = Trim$(strLineIn)
    If Len(TmpLine) > 0 Then
       If Not IsNumeric(Left(TmpLine, 1)) Then
          ' Line doesn't start with a number
          Special = Left(TmpLine, 1)
        
          TmpLine = Mid(TmpLine, 2)
          
          If Special <> "*" Then Special = ""
       Else
          Special = ""
       End If
    End If
    If Len(TmpLine) > 0 Then
       If IsNumeric(Left(TmpLine, 1)) Then
          intDataColumnCount = 0
          Do While Not Done
             k = InStr(TmpLine, vbTab)
             If k > 0 Then
                LineElement = Left(TmpLine, k - 1)
             Else
                LineElement = Trim(TmpLine)
                Done = True
             End If
             If Len(LineElement) > 0 Then
                If IsNumeric(LineElement) Then
                   varData(intDataColumnCount) = val(LineElement)
                Else
                   varData(intDataColumnCount) = LineElement
                End If
             End If
             intDataColumnCount = intDataColumnCount + 1
             
             If intDataColumnCount > ISONUM_FIELD_COUNT_MAX Or intDataColumnCount > UBound(varData) Then
                ' Data file has more than ISONUM_FIELD_COUNT_MAX fields; stop parsing the line
                If Not blnMaxColCountReachedErrorLogged Then
                    blnMaxColCountReachedErrorLogged = True
                    LogErrors 0, "Module1.bas->IsDataLine", "PEK data file encountered with more than " & ISONUM_FIELD_COUNT_MAX & " columns in the data block: " & strFilePath
                End If
                Done = True
             Else
                TmpLine = Mid(TmpLine, k + 1)
             End If
          Loop
          
          'make sure that there are no dot-zeros among data
          If intDataColumnCount > 0 Then
             For i = 0 To intDataColumnCount - 1
                 If varData(i) = "." Then varData(i) = 0
             Next i
             IsDataLine = True
          End If
       End If
    End If
End Function

Private Function GetNumberEqual(ByVal sLine As String) As String
'returns number as string after "=";
'empty string if none or not numeric
Dim EqualSignPos As Long
Dim sNumber As String
EqualSignPos = InStr(1, sLine, "=")
If EqualSignPos > 0 Then
   sNumber = Trim$(Right$(sLine, Len(sLine) - EqualSignPos))
   If IsNumeric(sNumber) Then
      GetNumberEqual = sNumber
   Else     'try to recover for cases like .03, -.234, 345.
      Select Case Left$(sNumber, 1)
      Case "."
           If IsNumeric("0" & sNumber) Then
              GetNumberEqual = "0" & sNumber
              Exit Function
           End If
      Case "-"
           If Mid$(sNumber, 2, 1) = "." Then
              If IsNumeric("-0." & Right$(sNumber, Len(sNumber) - 2)) Then
                 GetNumberEqual = "-0." & Right$(sNumber, Len(sNumber) - 2)
                 Exit Function
              End If
           End If
      End Select
      If Right$(sNumber, 1) = "." Then
         If IsNumeric(sNumber & "0") Then
            GetNumberEqual = sNumber & "0"
         End If
      End If
   End If
End If
End Function

'NEXT BLOCK OF FUNCTIONS ARE USED TO REWRITE PEK FILE
'FROM ORIGINAL PEK FILE AND VALUES FOUND IN GEL FILE

Public Function GeneratePEKFileUsingDataPoints(ByVal lngGelIndex As Long, ByVal blnLimitToDataPointsInView As Boolean, ByVal strFilePathForce As String, ByVal hwndOwner As Long) As Boolean
    ' Creates new barebones PEK file using the data points in memory

    ' If strFilePathForce contains text then that file path will be used (and the user will not be prompted)

    Const CHARGE_STATE_DATA_DECIMAL As Single = 0.1
    Const ISOTOPIC_DATA_DECIMAL As Single = 0.2

    Const EXPORT_STEP_COUNT As Integer = 2

    Dim strFilePath As String
    Dim strSuggestedName As String
    Dim strSepChar As String
    Dim strScanHeaderFileName As String

    Dim strLineOut As String

    Dim tsOutfile As TextStream
    Dim fso As FileSystemObject
    Dim blnSuccess As Boolean
    Dim blnResponse As Boolean

    Dim lngCurrentScan As Long
    Dim lngDataPointCountInScan As Long               ' Counts the number of CS-based LC-MS Features or the number of Isotopic-based LC-MS Features in a given scan -- not both
    Dim lngScanIndex As Long

    Dim lngCSPointerArray() As Long             ' 1-based array (dictated by GetCSScope)
    Dim lngIsoPointerArray() As Long            ' 1-based array (dictated by GetISScope)

    Dim lngPointsInViewCount As Long
    Dim sngPointsInViewScanNumbers() As Single    ' 0-based array; each scan number is appened with CHARGE_STATE_DATA_DECIMAL or ISOTOPIC_DATA_DECIMAL -- .1 is used for ChargeState data and .2 is used for Isotopic data
    Dim lngPointsInViewIndices() As Long        ' 0-based array
            
    Dim lngCSCount As Long
    Dim lngIsoCount As Long
    Dim lngIonIndex As Long
    Dim lngIndex As Long

    Dim lngScanInfoMaxIndex As Long

    Dim blnIsoHeaderWritten As Boolean
    Dim blnIsotopicDataPresent As Boolean           ' This is set to True if any CS data is present
    
    Dim blnValidDataPoint As Boolean
    Dim blnIsotopicDataPoint As Boolean
    Dim blnAborted As Boolean
    
On Error GoTo GeneratePEKFileUsingDataPointsErrorHandler

    ' Retrieve an array of the ion indices of the ions currently "In Scope"
    ' Note that GetCSScope and GetISScope will ReDim lngCSPointerArray() and lngIsoPointerArray() automatically
    lngCSCount = GetCSScope(lngGelIndex, lngCSPointerArray(), glSc_Current)
    lngIsoCount = GetISScope(lngGelIndex, lngIsoPointerArray(), glScope.glSc_Current)

    lngPointsInViewCount = lngCSCount + lngIsoCount
    If lngPointsInViewCount > 0 Then
        Set fso = New FileSystemObject
        
        If Len(strFilePathForce) = 0 Then
            ' Save to a file
            
            If Len(GelData(lngGelIndex).FileName) > 0 Then
                strSuggestedName = fso.GetBaseName(GelData(lngGelIndex).FileName) & "_Subset.pek"
            Else
                strSuggestedName = "DataPointsInView.pek"
            End If
            
            strFilePath = SelectFile(hwndOwner, "Enter file name to create using data points in view", "", True, strSuggestedName, "All Files (*.*)|*.*|PEK Files (*.pek)|*.pek", 2)
            If Len(strFilePath) = 0 Then Exit Function
        Else
            strFilePath = strFilePathForce
        End If
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data points were found visible in the current range (i.e. in the current zoom range).", vbInformation + vbOKOnly, glFGTU
        End If
        Exit Function
    End If
        
    ' Populate lngPointsInViewIndices and sngPointsInViewScanNumbers
    ReDim sngPointsInViewScanNumbers(lngPointsInViewCount - 1)
    ReDim lngPointsInViewIndices(lngPointsInViewCount - 1)

    frmProgress.InitializeForm "Saving data in view", 0, lngPointsInViewCount
    blnAborted = False
    
    lngPointsInViewCount = 0
    For lngIndex = 1 To lngCSCount
        sngPointsInViewScanNumbers(lngPointsInViewCount) = GelData(lngGelIndex).CSData(lngCSPointerArray(lngIndex)).ScanNumber + CHARGE_STATE_DATA_DECIMAL
        lngPointsInViewIndices(lngPointsInViewCount) = lngCSPointerArray(lngIndex)
        lngPointsInViewCount = lngPointsInViewCount + 1
    Next lngIndex
    
    If lngIsoCount > 0 Then
        blnIsotopicDataPresent = True
        For lngIndex = 1 To lngIsoCount
            sngPointsInViewScanNumbers(lngPointsInViewCount) = GelData(lngGelIndex).IsoData(lngIsoPointerArray(lngIndex)).ScanNumber + ISOTOPIC_DATA_DECIMAL
            lngPointsInViewIndices(lngPointsInViewCount) = lngIsoPointerArray(lngIndex)
            lngPointsInViewCount = lngPointsInViewCount + 1
        Next lngIndex
    End If
    
    ' Sort sngPointsInViewScanNumbers() and sort lngPointsInViewIndices() parallel with it
    ' We're using QSSingle since the scan numbers all end in 0.1 or 0.2
    Dim objSort As New QSSingle
    blnSuccess = objSort.QSAsc(sngPointsInViewScanNumbers, lngPointsInViewIndices)
    If blnSuccess = False Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Error sorting the PointsInViewScanNumbers array", vbInformation + vbOKOnly, glFGTU
        End If
        frmProgress.HideForm
        Exit Function
    End If

    Set tsOutfile = fso.CreateTextFile(strFilePath, True)

    strScanHeaderFileName = fso.GetFileName(strFilePath)

    ' Define the initial scan number
    lngCurrentScan = CInt(sngPointsInViewScanNumbers(0))
    lngDataPointCountInScan = 0

    ' Find lngCurrentScan in .ScanInfo() which is guaranteed to be sorted ascending
    lngScanInfoMaxIndex = UBound(GelData(lngGelIndex).ScanInfo)
    lngScanIndex = 1
    Do While lngScanIndex < lngScanInfoMaxIndex
        If GelData(lngGelIndex).ScanInfo(lngScanIndex).ScanNumber < lngCurrentScan Then
            lngScanIndex = lngScanIndex + 1
        Else
            Exit Do
        End If
    Loop

    If GelData(lngGelIndex).ScanInfo(lngScanIndex).ScanNumber <> lngCurrentScan Then
          ' Scan numbers don't match; this is unexpected and it means we cannot continue
          Debug.Assert False

          tsOutfile.Close
          frmProgress.HideForm

          If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
              MsgBox "Error saving PEK file; scan " & lngCurrentScan & " was not found in GelData().ScanInfo() array.  Unable to continue.  ", vbExclamation + vbOKOnly, glFGTU
          End If

          GeneratePEKFileUsingDataPoints = False
          Exit Function
    End If

    ' Write the file header
    tsOutfile.WriteLine App.Title & " - Version " & GetProgramVersion() & ", " & APP_BUILD_DATE
    tsOutfile.WriteLine "PEK file generated from processed data points"
    tsOutfile.WriteLine "Original PEK file: " & GelData(lngGelIndex).FileName
    tsOutfile.WriteLine

    ' Write the scan header
    GeneratePEKFileWriteCSandScanHeader tsOutfile, lngGelIndex, lngScanIndex, strScanHeaderFileName
    blnIsoHeaderWritten = False

    ' Generate the PEK file
    strSepChar = vbTab
    For lngIndex = 0 To lngPointsInViewCount - 1

        ' Write empty scan blocks if necessary
        Do While lngCurrentScan < CInt(sngPointsInViewScanNumbers(lngIndex))
            ' Close out the last scan
            If Not blnIsoHeaderWritten Then
                ' Scan only contained CS Data
                If blnIsotopicDataPresent Then
                    ' However, there is Isotopic data in the data file, so write the Isotopic Header anyway
                    GeneratePEKFileWriteIsoHeader tsOutfile, lngGelIndex
                    tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngDataPointCountInScan)
                    tsOutfile.WriteLine "Number of isotopic distributions detected = " & Trim(0)
                Else
                    ' Scan only contained CS data (or contained no data) and the data file only contains CS data
                    ' Write out the following only
                    tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngDataPointCountInScan)
                End If
            Else
                ' Scan contained Isotopic data
                tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngDataPointCountInScan)
                tsOutfile.WriteLine "Number of isotopic distributions detected = " & Trim(lngDataPointCountInScan)
            End If
            tsOutfile.WriteLine

            lngDataPointCountInScan = 0

            If lngScanIndex < lngScanInfoMaxIndex Then
                lngCurrentScan = GelData(lngGelIndex).ScanInfo(lngScanIndex + 1).ScanNumber
                lngScanIndex = lngScanIndex + 1
            Else
                ' We've passed the last scan in .ScanInfo; exit the for loop (and thus do not write out any more data)
                If lngIndex < lngPointsInViewCount - 1 Then
                    ' One or more data points has scan numbers greater than lngCurrentScan
                    Debug.Assert False
                End If
                Exit For
            End If

            ' Write the CS header
            ' First CS,    Number of CS,   Abundance,   Mass,   Standard deviation
            GeneratePEKFileWriteCSandScanHeader tsOutfile, lngGelIndex, lngScanIndex, strScanHeaderFileName

            blnIsoHeaderWritten = False
        Loop

        blnValidDataPoint = True
        If Abs(sngPointsInViewScanNumbers(lngIndex) - CInt(sngPointsInViewScanNumbers(lngIndex)) - CHARGE_STATE_DATA_DECIMAL) < 0.04 Then
            blnIsotopicDataPoint = False
        Else
            blnIsotopicDataPoint = True
        End If
        
        If Not blnIsotopicDataPoint Then
            ' CS data point
            If blnIsoHeaderWritten Then
                ' The Iso header has been written already for this scan
                ' Due to the way sngPointsInViewScanNumbers() was populated and sorted, this shouldn't happen
                ' Skip this UMC
                blnValidDataPoint = False
                Debug.Assert False
            End If
        Else
            If Not blnIsoHeaderWritten Then
                ' This is the first isotopic-based UMC encountered for this scan; Write the scan header
                GeneratePEKFileWriteIsoHeader tsOutfile, lngGelIndex
                blnIsoHeaderWritten = True

                ' Reset lngDataPointCountInScan since it ignores CS data when Isotopic data is present
                lngDataPointCountInScan = 0
            End If
        End If

        If blnValidDataPoint Then
            ' Record ClassStatsChargeBasis, UMCMZForChargeBasis
            If Not blnIsotopicDataPoint Then
                With GelData(lngGelIndex).CSData(lngPointsInViewIndices(lngIndex))
                    ' First CS,    Number of CS,   Abundance,   Mass,   Standard deviation
                    tsOutfile.WriteLine " " & Trim(.Charge) & strSepChar & _
                                Trim(.ChargeCount) & strSepChar & _
                                Trim(.Abundance) & strSepChar & _
                                Round(.AverageMW, 3) & strSepChar & _
                                Round(.MassStDev, 4)
                
                End With
            Else
                With GelData(lngGelIndex).IsoData(lngPointsInViewIndices(lngIndex))
                    ' CS,  Abundance,   m/z,   Fit,    Average MW, Monoisotopic MW,    Most abundant MW
                    strLineOut = " " & Trim(.Charge) & strSepChar & _
                                    Trim(.Abundance) & strSepChar & _
                                    Round(.MZ, 6) & strSepChar & _
                                    Trim(.Fit) & strSepChar & _
                                    Round(.AverageMW, 6) & strSepChar & _
                                    Round(.MonoisotopicMW, 6) & strSepChar & _
                                    Round(.MostAbundantMW, 6)

                    If .IsotopeLabel <> iltNone Then
                        strLineOut = strLineOut & strSepChar & GetIsotopeLabelTagName(.IsotopeLabel)
                    End If

                    tsOutfile.WriteLine strLineOut

                End With
            End If
            
            lngDataPointCountInScan = lngDataPointCountInScan + 1
        End If

        If lngIndex Mod 100 = 0 Then
            frmProgress.UpdateProgressBar lngIndex
            If KeyPressAbortProcess > 1 Then
                blnAborted = True
                Exit For
            End If
        End If

    Next lngIndex

    ' Close out the final scan
    tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngDataPointCountInScan)
    tsOutfile.WriteLine "Number of isotopic distributions detected = " & Trim(lngDataPointCountInScan)
    tsOutfile.WriteLine

    tsOutfile.Close
    frmProgress.HideForm

    If blnAborted Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Process aborted; saved the first " & Trim(lngPointsInViewCount) & " data points to file:" & vbCrLf & strFilePath, vbExclamation + vbOKOnly, "Aborted"
        End If
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Save complete; saved " & Trim(lngPointsInViewCount) & " data points to file:" & vbCrLf & strFilePath, vbInformation + vbOKOnly, "Done"
        End If
    End If

    GeneratePEKFileUsingDataPoints = blnSuccess
    Exit Function

GeneratePEKFileUsingDataPointsErrorHandler:
    Debug.Print "Error in sub GeneratePEKFileUsingDataPoints(): " & Err.Description
    Debug.Assert False

    LogErrors Err.Number, "GeneratePEKFileUsingDataPoints"

    On Error Resume Next
    If Not tsOutfile Is Nothing Then tsOutfile.Close

    GeneratePEKFileUsingDataPoints = False

End Function

Public Function GeneratePEKFileUsingUMCs(ByVal lngGelIndex As Long, ByVal blnLimitToUMCsInView As Boolean, ByVal strFilePathForce As String, ByVal hwndOwner As Long) As Boolean
    ' Creates new barebones PEK file using the LC-MS Features in memory
    ' Only outputs one entry for each UMC (class mass and class rep)
    
    ' If strFilePathForce contains text then that file path will be used (and the user will not be prompted)

    Const CHARGE_STATE_DATA_DECIMAL As Single = 0.1
    Const ISOTOPIC_DATA_DECIMAL As Single = 0.2
    
    Const EXPORT_STEP_COUNT As Integer = 2
    
    Dim strFilePath As String
    Dim strSuggestedName As String
    Dim strSepChar As String
    Dim strScanHeaderFileName As String
    
    Dim tsOutfile As TextStream
    Dim fso As FileSystemObject
    Dim blnSuccess As Boolean
    Dim blnResponse As Boolean
    
    Dim lngIndex As Long
    Dim lngAllUMCCount As Long
    Dim lngUMCsInViewCount As Long
    
    Dim lngCurrentScan As Long
    Dim lngUMCCountInScan As Long               ' Counts the number of CS-based LC-MS Features or the number of Isotopic-based LC-MS Features in a given scan -- not both
    Dim lngScanIndex As Long
    
    Dim lngCSPointerArray() As Long
    Dim lngIsoPointerArray() As Long            ' 1-based array (dictated by GetISScope)
    
    Dim lngCSCount As Long
    Dim lngIsoCount As Long
    Dim lngIonIndex As Long
    Dim lngUMCIndex As Long

    Dim sngScanClassRep As Single

    Dim blnUMCPresent() As Boolean          ' Records whether or not each UMC is present
    Dim lngUMCsInView() As Long             ' 0-based array; holds the indices of the LC-MS Features in view
    Dim sngUMCsClassRepScan() As Single     ' 0-based array; holds the scan numbers of the class rep for the LC-MS Features in view; parallel to lngUMCsInView; each scan number is appened with CHARGE_STATE_DATA_DECIMAL or ISOTOPIC_DATA_DECIMAL -- .1 is used for LC-MS Features from ChargeState data and .2 is used for LC-MS Features from Isotopic data
    
    Dim lngScanInfoMaxIndex As Long
    
    Dim lngProgessStepCount As Long
    
    Dim intCS As Integer
    Dim dblMZ As Double
    
    Dim blnIsoHeaderWritten As Boolean
    Dim blnValidUMC As Boolean
    Dim blnIsotopicUMCsPresent As Boolean           ' This is set to True if any of the LC-MS Features have .ClassRepType = gldtIS
    Dim blnAborted As Boolean
    
On Error GoTo GeneratePEKFileUsingUMCsErrorHandler

        
    ' Need to generate list of LC-MS Features that will be exported and sort them on ascending scan number of the class rep
    ' Then, step through the list and write an entry for each UMC

    lngAllUMCCount = GelUMC(lngGelIndex).UMCCnt
    strSepChar = vbTab
       If lngAllUMCCount > 0 Then
        Set fso = New FileSystemObject
        
        If Len(strFilePathForce) = 0 Then
            ' Save to a file
            
            If Len(GelData(lngGelIndex).FileName) > 0 Then
                strSuggestedName = fso.GetBaseName(GelData(lngGelIndex).FileName) & "_UMCs.pek"
            Else
                strSuggestedName = "UMCsInView.pek"
            End If
            
            strFilePath = SelectFile(hwndOwner, "Enter file name to create using LC-MS Features in view", "", True, strSuggestedName, "All Files (*.*)|*.*|PEK Files (*.pek)|*.pek", 2)
            If Len(strFilePath) = 0 Then Exit Function
        Else
            strFilePath = strFilePathForce
        End If
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No LC-MS Features are present in memory", vbInformation + vbOKOnly, glFGTU
        End If
        Exit Function
    End If
 
    
    lngProgessStepCount = 0
    frmProgress.InitializeForm "Preparing data", 0, EXPORT_STEP_COUNT, False, True, False
    frmProgress.InitializeSubtask "Finding LC-MS Features in View", 0, lngAllUMCCount
    blnAborted = False
    
    ' Reserve space for the UMC Presence array
    ReDim blnUMCPresent(lngAllUMCCount)
    
    ' Step 1: Retrieve an array of the ion indices of the ions currently "In Scope"
    ' Note that GetCSScope and GetISScope will ReDim lngCSPointerArray() and lngIsoPointerArray() automatically
    lngCSCount = GetCSScope(lngGelIndex, lngCSPointerArray(), glSc_Current)
    lngIsoCount = GetISScope(lngGelIndex, lngIsoPointerArray(), glScope.glSc_Current)
    
    ' Step 2: Set blnUMCPresent() to True for the LC-MS Features that the ions currently "In Scope" belong to
    lngUMCsInViewCount = 0
    For lngIonIndex = 1 To lngCSCount
        With GelDataLookupArrays(lngGelIndex).CSUMCs(lngCSPointerArray(lngIonIndex))
            For lngUMCIndex = 0 To .UMCCount - 1
                If Not blnUMCPresent(.UMCs(lngUMCIndex)) Then
                    blnUMCPresent(.UMCs(lngUMCIndex)) = True
                    lngUMCsInViewCount = lngUMCsInViewCount + 1
                End If
            Next lngUMCIndex
        End With
    Next lngIonIndex
    
    For lngIonIndex = 1 To lngIsoCount
        With GelDataLookupArrays(lngGelIndex).IsoUMCs(lngIsoPointerArray(lngIonIndex))
            For lngUMCIndex = 0 To .UMCCount - 1
                If Not blnUMCPresent(.UMCs(lngUMCIndex)) Then
                    blnUMCPresent(.UMCs(lngUMCIndex)) = True
                    lngUMCsInViewCount = lngUMCsInViewCount + 1
                End If
            Next lngUMCIndex
        End With
    Next lngIonIndex

    If lngUMCsInViewCount = 0 Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No LC-MS Features found in the current view; nothing was saved to disk.", vbInformation + vbOKOnly, glFGTU
        End If
        frmProgress.HideForm
        Exit Function
    End If


    ' Populate lngUMCsInView and sngUMCsClassRepScan
    ReDim lngUMCsInView(lngUMCsInViewCount - 1)
    ReDim sngUMCsClassRepScan(lngUMCsInViewCount - 1)
    
    lngUMCsInViewCount = 0
    For lngUMCIndex = 0 To lngAllUMCCount - 1
        If blnUMCPresent(lngUMCIndex) Then
            lngUMCsInView(lngUMCsInViewCount) = lngUMCIndex
            
            With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
                Select Case .ClassRepType
                Case gldtCS
                    sngScanClassRep = GelData(lngGelIndex).CSData(.ClassRepInd).ScanNumber + CHARGE_STATE_DATA_DECIMAL
                Case gldtIS
                    sngScanClassRep = GelData(lngGelIndex).IsoData(.ClassRepInd).ScanNumber + ISOTOPIC_DATA_DECIMAL
                    blnIsotopicUMCsPresent = True
                Case Else
                    Debug.Assert False
                    sngScanClassRep = (.MinScan + .MaxScan) / 2 + ISOTOPIC_DATA_DECIMAL
                End Select
            End With
            sngUMCsClassRepScan(lngUMCsInViewCount) = sngScanClassRep
            
            lngUMCsInViewCount = lngUMCsInViewCount + 1
        End If

        If lngUMCIndex Mod 100 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngUMCIndex
            If KeyPressAbortProcess > 1 Then
                frmProgress.HideForm
                Exit Function
            End If
        End If
    Next lngUMCIndex
    
    ' Sort sngUMCsClassRepScan() and sort lngUMCsInView() parallel with it
    ' We're using QSSingle since the scan numbers all end in 0.1 or 0.2
    Dim objSort As New QSSingle
    blnSuccess = objSort.QSAsc(sngUMCsClassRepScan, lngUMCsInView)
    If blnSuccess = False Then
        frmProgress.HideForm
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Error sorting the LC-MS Features Class Rep scan array", vbInformation + vbOKOnly, glFGTU
        End If
        Exit Function
    End If
    
    lngProgessStepCount = lngProgessStepCount + 1
    frmProgress.UpdateProgressBar lngProgessStepCount
    frmProgress.InitializeSubtask "Saving data", 0, lngUMCsInViewCount
    
    Set tsOutfile = fso.CreateTextFile(strFilePath, True)
    
    strScanHeaderFileName = fso.GetFileName(strFilePath)
    
    ' Define the initial scan number
    lngCurrentScan = CInt(sngUMCsClassRepScan(0))
    lngUMCCountInScan = 0
    
    ' Find lngCurrentScan in .ScanInfo() which is guaranteed to be sorted ascending
    lngScanInfoMaxIndex = UBound(GelData(lngGelIndex).ScanInfo)
    lngScanIndex = 1
    Do While lngScanIndex < lngScanInfoMaxIndex
        If GelData(lngGelIndex).ScanInfo(lngScanIndex).ScanNumber < lngCurrentScan Then
            lngScanIndex = lngScanIndex + 1
        Else
            Exit Do
        End If
    Loop
    
    If GelData(lngGelIndex).ScanInfo(lngScanIndex).ScanNumber <> lngCurrentScan Then
          ' Scan numbers don't match; this is unexpected and it means we cannot continue
          Debug.Assert False
      
          tsOutfile.Close
          frmProgress.HideForm
          
          If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
              MsgBox "Error saving PEK file; scan " & lngCurrentScan & " was not found in GelData().ScanInfo() array.  Unable to continue.  ", vbExclamation + vbOKOnly, glFGTU
          End If
          
          GeneratePEKFileUsingUMCs = False
          Exit Function
    End If
    
    ' Write the file header
    tsOutfile.WriteLine App.Title & " - Version " & GetProgramVersion() & ", " & APP_BUILD_DATE
    tsOutfile.WriteLine "PEK file generated from LC-MS Feature data"
    tsOutfile.WriteLine "Original PEK file: " & GelData(lngGelIndex).FileName
    tsOutfile.WriteLine

    ' Write the scan header
    GeneratePEKFileWriteCSandScanHeader tsOutfile, lngGelIndex, lngScanIndex, strScanHeaderFileName
    blnIsoHeaderWritten = False
    
    ' Step 4: Generate the PEK file
    For lngIndex = 0 To lngUMCsInViewCount - 1
        
        ' Write empty scan blocks if necessary
        Do While lngCurrentScan < CInt(sngUMCsClassRepScan(lngIndex))
            ' Close out the last scan
            If Not blnIsoHeaderWritten Then
                ' Scan only contained CS Data
                If blnIsotopicUMCsPresent Then
                    ' However, there are Isotopic LC-MS Features in the data file, so write the Isotopic Header anyway
                    GeneratePEKFileWriteIsoHeader tsOutfile, lngGelIndex
                    tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngUMCCountInScan)
                    tsOutfile.WriteLine "Number of isotopic distributions detected = " & Trim(0)
                Else
                    ' Scan only contained CS data (or contained no data) and the data file only contains CS data
                    ' Write out the following only
                    tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngUMCCountInScan)
                End If
            Else
                ' Scan contained Isotopic data
                tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngUMCCountInScan)
                tsOutfile.WriteLine "Number of isotopic distributions detected = " & Trim(lngUMCCountInScan)
            End If
            tsOutfile.WriteLine
            
            lngUMCCountInScan = 0
                        
            If lngScanIndex < lngScanInfoMaxIndex Then
                lngCurrentScan = GelData(lngGelIndex).ScanInfo(lngScanIndex + 1).ScanNumber
                lngScanIndex = lngScanIndex + 1
            Else
                ' We've passed the last scan in .ScanInfo; exit the for loop (and thus do not write out any more LC-MS Features)
                If lngIndex < lngUMCsInViewCount - 1 Then
                    ' One or more LC-MS Features has scan numbers greater than lngCurrentScan
                    Debug.Assert False
                End If
                Exit For
            End If
            
            ' Write the CS header
            ' First CS,    Number of CS,   Abundance,   Mass,   Standard deviation
            GeneratePEKFileWriteCSandScanHeader tsOutfile, lngGelIndex, lngScanIndex, strScanHeaderFileName
            
            blnIsoHeaderWritten = False
        Loop
        
        With GelUMC(lngGelIndex).UMCs(lngUMCsInView(lngIndex))
            blnValidUMC = True
        
            Select Case .ClassRepType
            Case gldtCS
                If blnIsoHeaderWritten Then
                    ' The Iso header has been written already for this scan
                    ' Due to the way sngUMCsClassRepScan() was populated and sorted, this shouldn't happen
                    ' Skip this UMC
                    blnValidUMC = False
                    Debug.Assert False
                End If
            Case gldtIS
                If Not blnIsoHeaderWritten Then
                    ' This is the first isotopic-based UMC encountered for this scan; Write the scan header
                    GeneratePEKFileWriteIsoHeader tsOutfile, lngGelIndex
                    blnIsoHeaderWritten = True
                    
                    ' Reset lngUMCCountInScan since it ignores CS data when Isotopic data is present
                    lngUMCCountInScan = 0
                End If
            Case Else
                blnValidUMC = False
            End Select
            
            If blnValidUMC Then
                ' Record ClassStatsChargeBasis, UMCMZForChargeBasis
                If GelUMC(lngGelIndex).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                    intCS = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge & strSepChar
                    dblMZ = MonoMassToMZ(.ClassMW, .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge)
                Else
                    ' Use charge of class rep
                    If .ClassRepType = gldtCS Then
                        intCS = GelData(lngGelIndex).CSData(.ClassRepInd).Charge & strSepChar
                        dblMZ = MonoMassToMZ(.ClassMW, GelData(lngGelIndex).CSData(.ClassRepInd).Charge)
                    Else
                        ' .ClassRepType = gldtIS
                        intCS = GelData(lngGelIndex).IsoData(.ClassRepInd).Charge & strSepChar
                        dblMZ = MonoMassToMZ(.ClassMW, GelData(lngGelIndex).IsoData(.ClassRepInd).Charge)
                    End If
                End If
                    
                If .ClassRepType = gldtCS Then
                
                    ' First CS,    Number of CS,   Abundance,   Mass,   Standard deviation
                    tsOutfile.WriteLine " " & Trim(intCS) & strSepChar & _
                                    Trim(GelData(lngGelIndex).CSData(.ClassRepInd).ChargeCount) & strSepChar & _
                                    Trim(.ClassAbundance) & strSepChar & _
                                    Round(GelData(lngGelIndex).CSData(.ClassRepInd).AverageMW, 3) & strSepChar & _
                                    Round(GelData(lngGelIndex).CSData(.ClassRepInd).MassStDev, 4)
                
                Else
                    ' .ClassRepType = gldtIS
                    
                    ' CS,  Abundance,   m/z,   Fit,    Average MW, Monoisotopic MW,    Most abundant MW
                    tsOutfile.WriteLine " " & Trim(intCS) & strSepChar & _
                                        Trim(.ClassAbundance) & strSepChar & _
                                        Round(dblMZ, 6) & strSepChar & _
                                        Trim(GelData(lngGelIndex).IsoData(.ClassRepInd).Fit) & strSepChar & _
                                        Round(GelData(lngGelIndex).IsoData(.ClassRepInd).AverageMW, 6) & strSepChar & _
                                        Round(.ClassMW, 6) & strSepChar & _
                                        Round(GelData(lngGelIndex).IsoData(.ClassRepInd).MostAbundantMW, 6)
                
                End If
                
            
                lngUMCCountInScan = lngUMCCountInScan + 1
            End If
        End With
        

        If lngIndex Mod 100 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngIndex
            If KeyPressAbortProcess > 1 Then
                blnAborted = True
                Exit For
            End If
        End If
        
    Next lngIndex
        
    ' Close out the final scan
    tsOutfile.WriteLine "Number of peaks in spectrum = " & Trim(lngUMCCountInScan)
    tsOutfile.WriteLine "Number of isotopic distributions detected = " & Trim(lngUMCCountInScan)
    tsOutfile.WriteLine
        
    tsOutfile.Close
    frmProgress.HideForm
    
    If blnAborted Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Process aborted; saved the first " & Trim(lngUMCsInViewCount) & " LC-MS Features to file:" & vbCrLf & strFilePath, vbExclamation + vbOKOnly, "Aborted"
        End If
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Save complete; saved " & Trim(lngUMCsInViewCount) & " LC-MS Features to file:" & vbCrLf & strFilePath, vbInformation + vbOKOnly, "Done"
        End If
    End If
    
    GeneratePEKFileUsingUMCs = blnSuccess
    Exit Function

GeneratePEKFileUsingUMCsErrorHandler:
    Debug.Print "Error in sub GeneratePEKFileUsingUMCs(): " & Err.Description
    Debug.Assert False
    
    LogErrors Err.Number, "GeneratePEKFileUsingUMCs"
    
    On Error Resume Next
    If Not tsOutfile Is Nothing Then tsOutfile.Close
    
    GeneratePEKFileUsingUMCs = False

End Function

Private Sub GeneratePEKFileWriteCSandScanHeader(ByRef tsOutfile As TextStream, lngGelIndex As Long, lngScanIndex As Long, strFileName As String)
    Dim strHeader As String
    
    tsOutfile.WriteLine "Time domain signal level:  " & GelData(lngGelIndex).ScanInfo(lngScanIndex).TimeDomainSignal
    tsOutfile.WriteLine "Filename: " & strFileName & "." & Format(GelData(lngGelIndex).ScanInfo(lngScanIndex).ScanNumber, "00000")
    tsOutfile.WriteLine "ScanType: Survey Scan"
    tsOutfile.WriteLine "Charge state mass transform results:"
    
    strHeader = GetDefaultPEKFileCSHeader()
    tsOutfile.WriteLine strHeader

End Sub

Private Sub GeneratePEKFileWriteIsoHeader(ByRef tsOutfile As TextStream, ByVal lngGelIndex As Long)
    Dim strHeader As String

    tsOutfile.WriteLine "Isotopic mass transform results:"
    
    strHeader = GetDefaultPEKFileIsoHeader(lngGelIndex, True)
    tsOutfile.WriteLine strHeader

End Sub

Private Function GetDefaultPEKFileCSHeader() As String
    GetDefaultPEKFileCSHeader = "First CS,    Number of CS,   Abundance,   Mass,   Standard deviation"
End Function

Private Function GetColumnValueDbl(ByRef varData() As Variant, ByVal intColumnIndex As Integer, Optional ByVal dblDefaultValue As Double = 0) As Double
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        GetColumnValueDbl = CDbl(varData(intColumnIndex))
    Else
        GetColumnValueDbl = dblDefaultValue
    End If
    
    Exit Function
    
GetColumnValueErrorHandler:
    Debug.Assert False
    GetColumnValueDbl = 0
    
End Function

Private Function GetColumnValueLng(ByRef varData() As Variant, ByVal intColumnIndex As Integer, Optional ByVal lngDefaultValue As Long = 0) As Long
    
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        GetColumnValueLng = CLng(varData(intColumnIndex))
    Else
        GetColumnValueLng = lngDefaultValue
    End If
    
    Exit Function
    
GetColumnValueErrorHandler:
    Debug.Assert False
    GetColumnValueLng = 0
    
End Function

Private Function GetColumnValueSng(ByRef varData() As Variant, ByVal intColumnIndex As Integer, Optional ByVal sngDefaultValue As Single = 0) As Single
    Dim dblValue As Double
    
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        dblValue = CDbl(varData(intColumnIndex))
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

Private Function GetColumnValueString(ByRef varData() As Variant, ByVal intColumnIndex As Integer, Optional ByVal strDefaultValue As String = "") As String
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        If Not IsNull(varData(intColumnIndex)) Then
            GetColumnValueString = CStr(varData(intColumnIndex))
        Else
            GetColumnValueString = strDefaultValue
        End If
    Else
        GetColumnValueString = strDefaultValue
    End If
    
    Exit Function
    
GetColumnValueErrorHandler:
    Debug.Assert False
    GetColumnValueString = ""
    
End Function

Private Function GetDefaultPEKFileIsoHeader(ByVal lngGelIndex As Long, ByVal blnUseGelStatusBits As Boolean) As String
    ' If lngGelIndex is > 0, then will optionally add the "IsotopeLabel" header if .IsoData().IsotopeLabel is not blank
    
    Dim strHeader As String
    Dim blnAddIsotopeHeaders As Boolean
    
    strHeader = "CS,  Abundance,   m/z,   Fit,    Average MW, Monoisotopic MW,    Most abundant MW"
    
On Error GoTo GetDefaultPEKFileIsoHeaderErrorHandler

    If lngGelIndex > 0 Then
        With GelData(lngGelIndex)
            blnAddIsotopeHeaders = False
            If blnUseGelStatusBits Then
                If ((.DataStatusBits And GEL_DATA_STATUS_BIT_ISOTOPE_LABEL_TAG) = GEL_DATA_STATUS_BIT_ISOTOPE_LABEL_TAG) Then
                    blnAddIsotopeHeaders = True
                End If
            Else
                If .IsoLines > 0 Then
                    If .IsoData(1).IsotopeLabel <> iltIsotopeLabelTagConstants.iltNone Then
                        blnAddIsotopeHeaders = True
                    End If
                End If
            End If
            
            If blnAddIsotopeHeaders Then
                strHeader = strHeader & ",  IsotopeLabel,  IsotopeRatio"
            End If
        End With
    End If
    
    GetDefaultPEKFileIsoHeader = strHeader
    Exit Function

GetDefaultPEKFileIsoHeaderErrorHandler:
    Debug.Print "Error in GetDefaultPEKFileIsoHeader: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "GetDefaultPEKFileIsoHeader"
    
    GetDefaultPEKFileIsoHeader = strHeader
    
End Function

Private Sub ParseCSColumnHeaders(ByVal strLineIn As String, ByRef intColumnMapping() As Integer, ByVal strInputFilePath As String)
    ' Examine strLineIn to define the column mappings
    
    Dim strUnknownColumnList As String
    Dim strData() As String
    
    Dim strMessage As String
    
    Dim intNewColumnMapping(CS_DATA_COLUMN_COUNT) As Integer
    Dim lngIndex As Long
    
    ' First replace any commas in strLineIn with tabs
    strLineIn = Replace(strLineIn, ",", vbTab)
    
    ' Now split on tabs
    strData = Split(strLineIn, vbTab)
    strUnknownColumnList = ""
    
    For lngIndex = 0 To UBound(intNewColumnMapping)
        intNewColumnMapping(lngIndex) = -1
    Next lngIndex
    
    For lngIndex = 0 To UBound(strData)
        If lngIndex >= CS_DATA_COLUMN_COUNT Then
            ' Too many column headers
            Debug.Assert False
            Exit For
        End If
        
        Select Case LCase(Trim(strData(lngIndex)))
        Case CS_COLUMN_FIRST_CS: intNewColumnMapping(crdCSRawDataIndex.crdCS) = lngIndex
        Case CS_COLUMN_NUMBER_OF_CS: intNewColumnMapping(crdCSRawDataIndex.crdNumberOfCS) = lngIndex
        Case CS_COLUMN_ABUNDANCE: intNewColumnMapping(crdCSRawDataIndex.crdAbu) = lngIndex
        Case CS_COLUMN_MASS: intNewColumnMapping(crdCSRawDataIndex.crdAverageMW) = lngIndex
        Case CS_COLUMN_STD_DEV: intNewColumnMapping(crdCSRawDataIndex.crdMWStDev) = lngIndex
        Case Else
            ' Unknown column header; ignore it, but post an entry to the analysis history
            If Len(strUnknownColumnList) > 0 Then
                strUnknownColumnList = strUnknownColumnList & ", "
            End If
            strUnknownColumnList = strUnknownColumnList & Trim(strData(lngIndex))
            
            Debug.Assert False
        End Select
        
    Next lngIndex
                    
    If intNewColumnMapping(crdCSRawDataIndex.crdCS) >= 0 And _
       intNewColumnMapping(crdCSRawDataIndex.crdAbu) >= 0 And _
       intNewColumnMapping(crdCSRawDataIndex.crdAverageMW) >= 0 Then
       
        ' Required columns are present; we can continue
        For lngIndex = 0 To UBound(intColumnMapping)
            intColumnMapping(lngIndex) = intNewColumnMapping(lngIndex)
        Next lngIndex
    
    Else
        ' Warn the user that the column headers weren't recognized
        If mReadMode = rmStoreData And Not mWarnedMissingRequiredCSColumnHeaders Then
            strMessage = CSV_COLUMN_HEADER_MISSING_WARNING & " not found in file " & strInputFilePath & "; default columns are: " & vbCrLf & GetDefaultPEKFileCSHeader
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
            End If
            AddToAnalysisHistory mGelIndex, strMessage
            mWarnedMissingRequiredCSColumnHeaders = True
        End If
    End If
                    
    If Len(strUnknownColumnList) > 0 Then
        If mReadMode = rmStoreData And Not mWarnedUnknownCSColumnHeader Then
            ' Unknown column header; ignore it, but post an entry to the
            AddToAnalysisHistory mGelIndex, CSV_COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & strInputFilePath & ": " & strUnknownColumnList & "; default columns are: " & vbCrLf & GetDefaultPEKFileCSHeader
            mWarnedUnknownCSColumnHeader = True
        End If
    End If

End Sub

Private Sub ParseIsosColumnHeaders(ByVal strLineIn As String, ByRef intColumnMapping() As Integer, ByVal strInputFilePath As String)
    ' Examine strLineIn to define the column mappings
    
    Dim strUnknownColumnList As String
    Dim strData() As String
    
    Dim strMessage As String
    
    Dim intNewColumnMapping(ISOS_DATA_COLUMN_COUNT) As Integer
    Dim lngIndex As Long

    ' First replace any commas in strLineIn with tabs
    strLineIn = Replace(strLineIn, ",", vbTab)
    
    ' Now split on tabs
    strData = Split(strLineIn, vbTab)
    strUnknownColumnList = ""
    
    For lngIndex = 0 To UBound(intNewColumnMapping)
        intNewColumnMapping(lngIndex) = -1
    Next lngIndex
    
    For lngIndex = 0 To UBound(strData)
        If lngIndex >= ISOS_DATA_COLUMN_COUNT Then
            ' Too many column headers
            Debug.Assert False
            Exit For
        End If
        
        Select Case LCase(Trim(strData(lngIndex)))
        Case ISOS_COLUMN_CS: intNewColumnMapping(irdIsoRawDataIndex.irdCS) = lngIndex
        Case ISOS_COLUMN_ABUNDANCE: intNewColumnMapping(irdIsoRawDataIndex.irdAbu) = lngIndex
        Case ISOS_COLUMN_MZ: intNewColumnMapping(irdIsoRawDataIndex.irdMZ) = lngIndex
        Case ISOS_COLUMN_ISO_FIT: intNewColumnMapping(irdIsoRawDataIndex.irdIsoFit) = lngIndex
        Case ISOS_COLUMN_AVG_MASS: intNewColumnMapping(irdIsoRawDataIndex.irdMWavg) = lngIndex
        Case ISOS_COLUMN_MONO_MASS: intNewColumnMapping(irdIsoRawDataIndex.irdMWMono) = lngIndex
        Case ISOS_COLUMN_MOST_ABU_MASS: intNewColumnMapping(irdIsoRawDataIndex.irdMWTMA) = lngIndex
        Case ISOS_COLUMN_IREP_MW_MONO_ABU: intNewColumnMapping(irdIsoRawDataIndex.irdIRepMWMonoAbu) = lngIndex
        Case ISOS_COLUMN_IREP_2DA: intNewColumnMapping(irdIsoRawDataIndex.irdIRep2Da) = lngIndex
        Case ISOS_COLUMN_PCT_MOST_ABUNDANT: intNewColumnMapping(irdIsoRawDataIndex.irdPctMostAbundant) = lngIndex
        Case ISOS_COLUMN_ELEMENT_LABEL_N14N15: intNewColumnMapping(irdIsoRawDataIndex.irdIsotopeTag) = lngIndex
        Case ISOS_COLUMN_ELEMENT_LABEL_N14N15_RATIO: intNewColumnMapping(irdIsoRawDataIndex.irdIsotopeTagRatio) = lngIndex
        Case Else
            ' Unknown column header; ignore it, but post an entry to the analysis history
            If Len(strUnknownColumnList) > 0 Then
                strUnknownColumnList = strUnknownColumnList & ", "
            End If
            strUnknownColumnList = strUnknownColumnList & Trim(strData(lngIndex))
            
            Debug.Assert False
        End Select
        
    Next lngIndex
                    
    If intNewColumnMapping(irdIsoRawDataIndex.irdCS) >= 0 And _
       intNewColumnMapping(irdIsoRawDataIndex.irdAbu) >= 0 And _
       intNewColumnMapping(irdIsoRawDataIndex.irdMWMono) >= 0 Then
       
        ' Required columns are present; we can continue
        For lngIndex = 0 To UBound(intColumnMapping)
            intColumnMapping(lngIndex) = intNewColumnMapping(lngIndex)
        Next lngIndex
    
    Else
        ' Warn the user that the column headers weren't recognized
        If mReadMode = rmStoreData And Not mWarnedMissingRequiredIsosColumnHeaders Then
            strMessage = CSV_COLUMN_HEADER_MISSING_WARNING & " not found in file " & strInputFilePath & "; default columns are: " & vbCrLf & GetDefaultPEKFileIsoHeader(mGelIndex, False)
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
            End If
            AddToAnalysisHistory mGelIndex, strMessage
            mWarnedMissingRequiredIsosColumnHeaders = True
        End If
    End If
                    
    If Len(strUnknownColumnList) > 0 Then
        If mReadMode = rmStoreData And Not mWarnedUnknownIsosColumnHeader Then
            ' Unknown column header; ignore it, but post an entry to the
            AddToAnalysisHistory mGelIndex, CSV_COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & strInputFilePath & ": " & strUnknownColumnList & "; default columns are: " & vbCrLf & GetDefaultPEKFileIsoHeader(mGelIndex, False)
            mWarnedUnknownIsosColumnHeader = True
        End If
    End If

End Sub

Private Sub SetDefaultCSColumnMapping(ByRef intCSColumnMapping() As Integer)
    Dim lngIndex As Long
    
    ' First set the column mappings to -1 (not present)
    For lngIndex = 0 To UBound(intCSColumnMapping)
        intCSColumnMapping(lngIndex) = -1
    Next lngIndex
     
    ' Now set the default columns
    intCSColumnMapping(crdCSRawDataIndex.crdCS) = 0
    intCSColumnMapping(crdCSRawDataIndex.crdNumberOfCS) = 1
    intCSColumnMapping(crdCSRawDataIndex.crdAbu) = 2
    intCSColumnMapping(crdCSRawDataIndex.crdAverageMW) = 3
    intCSColumnMapping(crdCSRawDataIndex.crdMWStDev) = 4

End Sub

Private Sub SetDefaultIsosColumnMapping(ByRef intIsosColumnMapping() As Integer)
    Dim lngIndex As Long
    
    ' First set the column mappings to -1 (not present)
    For lngIndex = 0 To UBound(intIsosColumnMapping)
        intIsosColumnMapping(lngIndex) = -1
    Next lngIndex
    
    ' Now set the default columns
    intIsosColumnMapping(irdIsoRawDataIndex.irdCS) = 0
    intIsosColumnMapping(irdIsoRawDataIndex.irdAbu) = 1
    intIsosColumnMapping(irdIsoRawDataIndex.irdMZ) = 2
    intIsosColumnMapping(irdIsoRawDataIndex.irdIsoFit) = 3
    intIsosColumnMapping(irdIsoRawDataIndex.irdMWavg) = 4
    intIsosColumnMapping(irdIsoRawDataIndex.irdMWMono) = 5
    intIsosColumnMapping(irdIsoRawDataIndex.irdMWTMA) = 6
    intIsosColumnMapping(irdIsoRawDataIndex.irdIRepMWMonoAbu) = 7
    intIsosColumnMapping(irdIsoRawDataIndex.irdIRep2Da) = 8

End Sub

Public Function WriteGELAsPEK(ByVal lngGelIndex As Long, ByVal hwndOwner As Long) As Boolean
    '------------------------------------------------------------------------------------
    'writes gel file as PEK file; recreates PEK file by combining lines from PEK file and
    'data from the display; returns True on success; False otherwise
    'NOTE : if PEK file parameters changed (calibration equation for example) new PEK
    'file rewrites old parameters (from original file)
    '------------------------------------------------------------------------------------
    Dim OldPEK As String
    Dim OldPEKNum As Integer
    Dim NewPEK As String
    Dim TmpPEK As String
    Dim TmpPEKNum As Integer
    Dim CurrLine As String
    Dim LineType As Integer
    Dim FN As Integer
    Dim CurrFN As Integer
    
    Dim eResponse As VbMsgBoxResult
    Dim lngTotalBytesRead As Long, lngLineCount As Long
    Dim lngIsoDataStartIndex As Long
    Dim blnAborted As Boolean
    
    Dim fso As FileSystemObject
    
    On Error GoTo WriteGELAsPEKErrorHandler
    
    Set fso = New FileSystemObject
    
    OldPEK = GelData(lngGelIndex).FileName
    If Not fso.FileExists(OldPEK) Then
        eResponse = MsgBox("Original PEK file not found. Without the original file, this function cannot proceed. Do you want to specify path to the original PEK file?", vbYesNo, glFGTU)
        If eResponse = vbYes Then
            If Len(GelData(lngGelIndex).FileName) > 0 Then
                OldPEK = fso.GetFileName(GelData(lngGelIndex).FileName)
            Else
                OldPEK = ""
            End If
            
            OldPEK = SelectFile(hwndOwner, "Choose original PEK file:", "", False, OldPEK, "All Files (*.*)|*.*|PEK Files (*.pek)|*.pek", 2)
            If Len(OldPEK) > 0 Then
                DoEvents
                GelData(lngGelIndex).FileName = OldPEK     'note new path to original PEK
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    NewPEK = fso.GetBaseName(OldPEK) & "_Processed.pek"
    
    NewPEK = SelectFile(hwndOwner, "Enter new PEK file name", "", True, NewPEK, "All Files (*.*)|*.*|PEK Files (*.pek)|*.pek", 2)
    If Len(NewPEK) <= 0 Then
        Exit Function
    End If
    
    frmProgress.InitializeForm "Create new .PEK file based on loaded data and old .PEK file", 0, FileLen(OldPEK), True, False, True, MDIForm1
    blnAborted = False
    
    ' Open the old PEK file
    OldPEKNum = FreeFile()
    Open OldPEK For Input As OldPEKNum
    
    TmpPEK = GetTempFolder() & "NewPEKFile" & CLng(Rnd(1) * 100000!) & ".pek"
    TmpPEKNum = FreeFile()
    lngIsoDataStartIndex = 1
    
    Open TmpPEK For Output As TmpPEKNum
    Do Until EOF(OldPEKNum)
        Line Input #OldPEKNum, CurrLine
        
        lngTotalBytesRead = lngTotalBytesRead + Len(CurrLine) + 2      ' Add 2 bytes to account for CrLf at end of line
        lngLineCount = lngLineCount + 1
        If lngLineCount Mod 250 = 0 Then
            frmProgress.UpdateProgressBar lngTotalBytesRead
            If KeyPressAbortProcess > 1 Then
                blnAborted = True
                Exit Do
            End If
        End If
       
        LineNow1 CurrLine, LineType, FN, OldPEK
        Select Case LineType
        Case LINE_FILENAME_AKA_SCAN_NUMBER
             CurrFN = FN
             Print #TmpPEKNum, CurrLine
             'print also frequency shift if asked
             If glWriteFreqShift Then Print #TmpPEKNum, glLaV2DG_FREQUENCY_SHIFT _
                                 & Format$(GelData(lngGelIndex).ScanInfo(GetDFIndex(lngGelIndex, CurrFN)).FrequencyShift, "0.0000")
        Case LINE_DATA_CS
             ' Get the default CS header line
             CurrLine = GetDefaultPEKFileCSHeader
             Print #TmpPEKNum, CurrLine
             WriteCSDataToPEK TmpPEKNum, lngGelIndex, CurrFN
        Case LINE_DATA_ISO
             ' Get the default ISO header line
             CurrLine = GetDefaultPEKFileIsoHeader(lngGelIndex, True)
             Print #TmpPEKNum, CurrLine
             WriteIsoDataToPEK TmpPEKNum, lngGelIndex, CurrFN, lngIsoDataStartIndex
        Case LINE_WHATEVER
             Print #TmpPEKNum, CurrLine
        Case LINE_NOTHING
             ' Ignore CS & Iso data from original file
        Case Else
             ' Shouldn't normally reach this code, but if we do, just write out CurrLine
             Print #TmpPEKNum, CurrLine
        End Select
    Loop
    Close TmpPEKNum
    Close OldPEKNum
    DoEvents
    
    FileCopy TmpPEK, NewPEK
    DoEvents
    Kill TmpPEK
    
    frmProgress.HideForm
    If blnAborted Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Process aborted; saved the first " & Trim(lngIsoDataStartIndex) & " data points to file:" & vbCrLf & NewPEK, vbExclamation + vbOKOnly, "Aborted"
        End If
    Else
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Save complete; saved " & Trim(GelData(lngGelIndex).CSLines + GelData(lngGelIndex).IsoLines) & " data points to file:" & vbCrLf & NewPEK, vbInformation + vbOKOnly, "Done"
        End If
    End If
    
    WriteGELAsPEK = True
    Exit Function
    
WriteGELAsPEKErrorHandler:
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        LogErrors Err.Number, "WriteGELasPEK"
    Else
        MsgBox "Error in WriteGELasPEK: " & Err.Description
    End If
    
    frmProgress.HideForm
    
    On Error Resume Next
    Close OldPEKNum
    Close TmpPEKNum
    
End Function

Private Sub LineNow1(ByVal strLineIn As String, ByRef LineType As Integer, ByRef FN As Integer, ByVal strFilePath As String)
    '------------------------------------------------------------------------------
    'this is simpler version of LineNow used when PEK file is created from display
    'strLineIn is line, LineType returns type of line, FN returns scan number
    '------------------------------------------------------------------------------
    Dim intDataColumnCount As Integer
    Dim varData(ISONUM_FIELD_COUNT_MAX) As Variant
    Dim Special As String

On Error GoTo err_LineNow1
    
    Select Case Left(strLineIn, 8)
    Case t8FILENAME
        FN = ExtractScanNumberFromFilenameLine(strLineIn)
        LineType = LINE_FILENAME_AKA_SCAN_NUMBER
    Case t8DATA_CS
       ThisLine = LINE_DATA_CS
       LineType = LINE_DATA_CS
    Case t8DATA_ISO, t8DATA_ISO_Tabbed
       ThisLine = LINE_DATA_ISO
       LineType = LINE_DATA_ISO
    Case Else
       Select Case ThisLine
       Case LINE_DATA_CS
            If IsDataLine(Trim$(strLineIn), varData, intDataColumnCount, Special, strFilePath) Then
               LineType = LINE_NOTHING
            Else
               LineType = LINE_WHATEVER
               ThisLine = LINE_NOTHING
            End If
       Case LINE_DATA_ISO
            If IsDataLine(Trim$(strLineIn), varData, intDataColumnCount, Special, strFilePath) Then
               LineType = LINE_NOTHING
            Else
               LineType = LINE_WHATEVER
               ThisLine = LINE_NOTHING
            End If
       Case Else
            ThisLine = LINE_NOTHING
            LineType = LINE_WHATEVER
       End Select
    End Select
    Exit Sub

err_LineNow1:        'write unrecognized line
    LineType = LINE_WHATEVER
End Sub

Private Sub WriteCSDataToPEK(ByVal hfile As Integer, ByVal Ind As Long, ByVal FN As Integer)
'-------------------------------------------------------------------------------------------
'hFile - handle to open file, Ind - GelData index, FN - scan number
'NOTE: PEK data are ordered(descending) on Intensity
'-------------------------------------------------------------------------------------------
Dim FNCnt As Long
Dim Indx() As Long
Dim Abu() As Double
Dim sLine As String
Dim qsdSort As New QSDouble
Dim i As Long
On Error Resume Next
With GelData(Ind)
    If .CSLines > 0 Then
       ReDim Indx(1 To .CSLines)
       ReDim Abu(1 To .CSLines)
       FNCnt = 0
       For i = 1 To .CSLines
           If .CSData(i).ScanNumber = FN Then
              FNCnt = FNCnt + 1
              Indx(FNCnt) = i
              Abu(FNCnt) = .CSData(i).Abundance
           ElseIf .CSData(i).ScanNumber > FN Then
              Exit For
           End If
       Next i
       If FNCnt > 0 Then
          ReDim Preserve Indx(1 To FNCnt)
          ReDim Preserve Abu(1 To FNCnt)
          If qsdSort.QSDesc(Abu(), Indx()) Then
             For i = 1 To FNCnt
                'this part is always included
                sLine = .CSData(Indx(i)).Charge & vbTab & .CSData(Indx(i)).ChargeCount _
                    & vbTab & Format$(.CSData(Indx(i)).Abundance, "Scientific") _
                    & vbTab & Format$(.CSData(Indx(i)).AverageMW, "0.0000") _
                    & vbTab & Format$(.CSData(Indx(i)).MassStDev, "0.0000")
                'add ER if included
''                If Not IsNull(.CSVar(Indx(i), csvfMTDDRatio)) Then
''                   If IsNumeric(.CSVar(Indx(i), csvfMTDDRatio)) Then
''                      sLine = sLine & vbTab & .CSVar(Indx(i), csvfMTDDRatio)
''                   Else
''                      If InStr(1, .CSVar(Indx(i), csvfIsotopeLabel), "C") Then
''                         sLine = sLine & vbTab & "C12"
''                      ElseIf InStr(1, .CSVar(Indx(i), csvfIsotopeLabel), "N") Then
''                         sLine = sLine & vbTab & "Normal"
''                      End If
''                   End If
''                End If
''                sLine = sLine & vbTab & .CSData(Indx(i)).ExpressionRatio
''                'add database fit and error if included
''                If .CSData(Indx(i)).IsotopicFitRatio > 0 Then
''                   'add Tab if we did not have ER in this line
''                   If IsNull(.CSVar(Indx(i), 2)) Then sLine = sLine & vbTab
''                   sLine = sLine & vbTab & Format$(.CSData(Indx(i)).IsotopicFitRatio, "0.0000") _
''                       & vbTab & Format$(.CSData(Indx(i)).IsotopicAtomCount, "0.0000")
''                End If

                Print #hfile, sLine
             Next i
          End If
       End If
    End If
End With
End Sub

Private Sub WriteIsoDataToPEK(ByVal hfile As Integer, ByVal Ind As Long, ByVal FN As Integer, ByRef lngIsoDataStartIndex As Long)
'--------------------------------------------------------------------------------------------
'hFile - handle to open file, Ind - GelData index, FN - scan number
'NOTE: PEK data are ordered(descending) on Intensity
'Parameter lngIsoDataStartIndex is used to speed the search of the data in .IsoData
'If a match is found for FN, then updates lngIsoDataStartIndex to the index of the last match found
'--------------------------------------------------------------------------------------------
Dim DataMatchCount As Long
Dim Indx() As Long
Dim Abu() As Double
Dim sLine As String
Dim strAppendText As String
Dim qsdSort As New QSDouble
Dim i As Long
Dim blnSuccess As Boolean

Dim strIsotopeLabel As String
''Dim blnLegacyIsotopeLabel As Boolean

With GelData(Ind)
  If .IsoLines > 0 Then
     ReDim Indx(1 To .IsoLines)
     ReDim Abu(1 To .IsoLines)
     
     DataMatchCount = 0
     If lngIsoDataStartIndex < 1 Then lngIsoDataStartIndex = 1
     If lngIsoDataStartIndex > .IsoLines Then lngIsoDataStartIndex = .IsoLines
     For i = lngIsoDataStartIndex To .IsoLines
         If .IsoData(i).ScanNumber = FN Then
            DataMatchCount = DataMatchCount + 1
            Indx(DataMatchCount) = i
            Abu(DataMatchCount) = .IsoData(i).Abundance        'Intensity
         ElseIf .IsoData(i).ScanNumber > FN Then
            Exit For
         End If
     Next i
     If DataMatchCount > 0 Then
        lngIsoDataStartIndex = lngIsoDataStartIndex + DataMatchCount
        
        ReDim Preserve Indx(1 To DataMatchCount)
        ReDim Preserve Abu(1 To DataMatchCount)
        
        If DataMatchCount > 1 Then
            blnSuccess = qsdSort.QSDesc(Abu(), Indx())
        Else
            blnSuccess = True
        End If
        
        If blnSuccess Then
           For i = 1 To DataMatchCount
''              If Not IsNull(.IsoVar(Indx(i), isvfIsotopeLabel)) Then
''                 strIsotopeLabel = CStr(.IsoVar(Indx(i), isvfIsotopeLabel))
''              Else
''                 strIsotopeLabel = ""
''              End If
              
''              If InStr(strIsotopeLabel, "*") Then
''                 ' Legacy data may have an asterisk in strIsotopeLabel
''                 sLine = "*"
''                 blnLegacyIsotopeLabel = True
''              Else
''                 sLine = " "
''                 blnLegacyIsotopeLabel = False
''              End If

              sLine = " " & Trim(.IsoData(Indx(i)).Charge) & vbTab & _
                        Format$(.IsoData(Indx(i)).Abundance, "Scientific") & vbTab & _
                        Format$(.IsoData(Indx(i)).MZ, "0.0000") & vbTab & _
                        Format$(.IsoData(Indx(i)).Fit, "0.0000") & vbTab & _
                        Format$(.IsoData(Indx(i)).AverageMW, "0.0000") & vbTab & _
                        Format$(.IsoData(Indx(i)).MonoisotopicMW, "0.0000") & vbTab & _
                        Format$(.IsoData(Indx(i)).MostAbundantMW, "0.0000")

              If .IsoData(Indx(i)).IsotopeLabel <> iltNone Then
                sLine = sLine & vbTab & GetIsotopeLabelTagName(.IsoData(Indx(i)).IsotopeLabel)
              End If

''              strAppendText = ""
''              If Len(strIsotopeLabel) > 0 Then
''                If IsNumeric(strIsotopeLabel) Or blnLegacyIsotopeLabel Then
''                    ' strIsotopeLabel probably contains and AMT NET value; do not output the value
''                    strAppendText = ""
''                    blnLegacyIsotopeLabel = True
''                Else
''                    ' strIsotopeLabel probably contains N14 or N15 or O16 or O18
''                    strAppendText = strIsotopeLabel
''                End If
''              Else
''                If Not IsNull(.IsoVar(Indx(i), isvfMTDDRatio)) Then
''                  If IsNumeric(.IsoVar(Indx(i), isvfMTDDRatio)) Then
''                      strAppendText = .IsoData(Indx(i)).ExpressionRatio
''                      blnLegacyIsotopeLabel = True
''                   End If
''                End If
''              End If
''              sLine = sLine & vbTab & strAppendText
''                sLine = sLine & vbTab & .IsoData(Indx(i)).ExpressionRatio
                
''              If Not blnLegacyIsotopeLabel Then
''                If .IsoData(Indx(i)).IsotopicFitRatio > 0 Then
''                   sLine = sLine & vbTab & Format$(.IsoData(Indx(i)).IsotopicFitRatio, "0.0000")
''                   sLine = sLine & vbTab & Trim(.IsoData(Indx(i)).IsotopicAtomCount)
''                End If
''              End If
              
              Print #hfile, sLine
           Next i
        End If
     End If
  End If
End With
End Sub


Attribute VB_Name = "modFileIOCsv"
Option Explicit

Public Const CSV_ISOS_IC_FILE_SUFFIX As String = "isos_ic.csv"
Public Const CSV_ISOS_FILE_SUFFIX As String = "isos.csv"
Public Const CSV_SCANS_FILE_SUFFIX As String = "scans.csv"
Public Const CSV_COLUMN_HEADER_UNKNOWN_WARNING As String = "Warning: unknown column headers"
Public Const CSV_COLUMN_HEADER_MISSING_WARNING As String = "Warning: expected important column headers"

' Note: These should all be lowercase string values
Private Const ISOS_COLUMN_SCAN_NUM As String = "scan_num"
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

' Note: These should all be lowercase string values
Private Const SCANS_COLUMN_SCAN_NUM As String = "scan_num"
Private Const SCANS_COLUMN_TIME_A As String = "time"
Private Const SCANS_COLUMN_TIME_B As String = "scan_time"
Private Const SCANS_COLUMN_TYPE As String = "type"
Private Const SCANS_COLUMN_NUM_DEISOTOPED As String = "num_deisotoped"
Private Const SCANS_COLUMN_NUM_PEAKS As String = "num_peaks"
Private Const SCANS_COLUMN_TIC As String = "tic"
Private Const SCANS_COLUMN_BPI_MZ As String = "bpi_mz"
Private Const SCANS_COLUMN_BPI As String = "bpi"
Private Const SCANS_COLUMN_TIME_DOMAIN_SIGNAL As String = "time_domain_signal"
Private Const SCANS_COLUMN_PEAK_INTENSITY_THRESHOLD As String = "peak_intensity_threshold"
Private Const SCANS_COLUMN_PEPTIDE_INTENSITY_THRESHOLD As String = "peptide_intensity_threshold"


Private Const SCAN_INFO_DIM_CHUNK As Long = 10000
Private Const ISO_DATA_DIM_CHUNK As Long = 25000

Private Const SCAN_FILE_COLUMN_COUNT As Integer = 11
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
End Enum

Private Const ISOS_FILE_COLUMN_COUNT As Integer = 12
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

Private mPrescannedData As clsFileIOPrescannedData

Private mValidDataPointCount As Long
Private mSubtaskMessage As String

Private mReadMode As rmReadModeConstants
Private mCurrentProgressStep As Integer

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

Private Function GetDefaultIsosColumnHeaders(blnRequiredColumnsOnly As Boolean) As String
    Dim strHeaders As String
    
    strHeaders = ISOS_COLUMN_SCAN_NUM
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
    End If

    GetDefaultIsosColumnHeaders = strHeaders
End Function

Private Function GetDefaultScansColumnHeaders(blnRequiredColumnsOnly As Boolean) As String
    Dim strHeaders As String

    strHeaders = SCANS_COLUMN_SCAN_NUM
    strHeaders = strHeaders & ", " & SCANS_COLUMN_TIME_B
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
    End If
    
    GetDefaultScansColumnHeaders = strHeaders
End Function

Public Function LoadNewCSV(ByVal CSVFilePath As String, ByVal lngGelIndex As Long, _
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
    'eDataFilterMode is ignored by this function
    '-------------------------------------------------------------------------
    
    Dim intProgressCount As Integer
    
    Dim strScansFilePath As String
    Dim strIsosFilePath As String
    Dim strBaseName As String
    
    Dim eResponse As VbMsgBoxResult
    
    Dim fso As New FileSystemObject
    Dim tsInFile As TextStream
    Dim strLineIn As String
    
    Dim blnValidScanFile As Boolean
    Dim blnValidDataPoint As Boolean
    
    Dim lngReturnValue As Long
    
    Dim lngScansFileByteCount As Long
    Dim lngByteCountTotal As Long
    Dim lngTotalBytesRead As Long
    
On Error GoTo LoadNewCSVErrorHandler

    ' Update the filter variables
    mGelIndex = lngGelIndex
    mMaxFit = MaxFit
    mFilterByAbundance = blnFilterByAbundance
    mAbundanceMin = dblMinAbu
    mAbundanceMax = dblMaxAbu

    mMaximumDataCountEnabled = blnMaximumDataCountEnabled
    mMaximumDataCountToLoad = lngMaximumDataCountToLoad
    
    If mMaximumDataCountEnabled Then
        If mMaximumDataCountToLoad < 10 Then mMaximumDataCountToLoad = 10
        intProgressCount = 5
    Else
        intProgressCount = 3
    End If

    mCurrentProgressStep = 0
    frmProgress.InitializeForm "Loading data file", mCurrentProgressStep, intProgressCount, False, True, True, MDIForm1
    lngReturnValue = -10
    
    ' Resolve the CSV FilePath given to the ScansFilePath and the IsosFilePath variables
    
    If Not ResolveCSVFilePaths(CSVFilePath, strScansFilePath, strIsosFilePath, strBaseName) Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Unable to resolve the given FilePath to the _isos.csv and _scans.csv files: " & vbCrLf & CSVFilePath, vbExclamation + vbOKOnly, glFGTU
        End If
        LoadNewCSV = -7
        Exit Function
    End If
    
    ' Validate that the input file(s) exist
    If Not fso.FileExists(strIsosFilePath) Then
        LoadNewCSV = -6
        Exit Function
    End If
    
    blnValidScanFile = True
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
            blnValidScanFile = False
        End If
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
    lngByteCountTotal = FileLen(strIsosFilePath)
''    If mMaximumDataCountEnabled Then
''        lngByteCountTotal = lngByteCountTotal * 2
''    End If
    
    If blnValidScanFile Then
        lngScansFileByteCount = FileLen(strScansFilePath)
        lngByteCountTotal = lngByteCountTotal + lngScansFileByteCount
    End If
    
    frmProgress.InitializeSubtask "Reading data", 0, lngByteCountTotal
    
    mScanInfoCount = 0
    ReDim GelData(mGelIndex).ScanInfo(0)
    If blnValidScanFile Then
        ' Read the scans file and populate .ScanInfo
        lngReturnValue = ReadCSVScanFile(fso, strScansFilePath, strBaseName, lngTotalBytesRead)
    Else
        lngReturnValue = 0
    End If
    
    If lngReturnValue = 0 Then
        ' Read the Isos file
        ' Note that the CSV Isos file only contains isotopic data, not charge state data
        lngReturnValue = ReadCSVIsosFile(fso, strIsosFilePath, strBaseName, lngScansFileByteCount, lngByteCountTotal, lngTotalBytesRead, blnValidScanFile)
    End If
        
    LoadNewCSV = lngReturnValue
    Exit Function

LoadNewCSVErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "LoadNewCSV"
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    LoadNewCSV = lngReturnValue
    
End Function

Private Function ReadCSVIsosFile(ByRef fso As FileSystemObject, ByVal strIsosFilePath As String, ByVal strBaseName As String, _
                                 ByVal lngScansFileByteCount As Long, ByVal lngByteCountTotal As Long, _
                                 ByRef lngTotalBytesRead As Long, ByVal blnValidScanFile As Boolean) As Long

    ' Returns 0 if no error, the error number if an error

    Dim lngIndex As Long
    Dim lngReturnValue As Long
    
    Dim objFile As File
    Dim objFolder As Folder
    
    Dim blnMonoPlus2DataPresent As Boolean
    Dim MaxMZ As Double
    Dim intColumnMapping() As Integer
    
On Error GoTo ReadCSVIsosFileErrorHandler

    ReDim intColumnMapping(ISOS_FILE_COLUMN_COUNT - 1) As Integer
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To ISOS_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    If Len(strBaseName) = 0 Then
        strBaseName = fso.GetBaseName(strIsosFilePath)
    End If
        
    With GelData(mGelIndex)
        If Not blnValidScanFile Then
            mScanInfoCount = 0
            ReDim .ScanInfo(SCAN_INFO_DIM_CHUNK)
        End If
        
        .LinesRead = 0
        .DataLines = 0
        .CSLines = 0
        .IsoLines = 0
        
        ReDim .IsoData(ISO_DATA_DIM_CHUNK)
    End With

    If mMaximumDataCountEnabled Then
        mReadMode = rmReadModeConstants.rmPrescanData
    Else
        mReadMode = rmReadModeConstants.rmStoreData
    End If
    
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
            
            mSubtaskMessage = "Pre-scanning Isotopic CSV file to determine data to load"
        Else
            mSubtaskMessage = "Reading Isotopic CSV file"
        End If
        frmProgress.InitializeSubtask mSubtaskMessage, 0, lngByteCountTotal
        frmProgress.UpdateSubtaskProgressBar lngScansFileByteCount
        
            ' Reset the tracking variables
        mValidDataPointCount = 0
        lngTotalBytesRead = 0
    
        lngReturnValue = ReadCSVIsosFileWork(fso, strIsosFilePath, lngTotalBytesRead, intColumnMapping, blnValidScanFile)
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
            
    With GelData(mGelIndex)
         ' Old: .PathtoDataFiles = GetPathWOFileName(CurrDataFName)
         ' New: data file folder path is the folder one folder up from .Filename's folder if .Filename's folder contains _Auto00000
         '      if .Filename's folder does not contain _Auto0000, then simply use .Filename's folder
        .PathtoDataFiles = DetermineParentFolderPath(.FileName)
        
        ' Note: CS Data is not loaded by this function
        ReDim .CSData(0)
        
        ' Find the minimum and maximum MW, Abundance, and MZ values, and set the filters
        MaxMZ = 0
        blnMonoPlus2DataPresent = False
        If .IsoLines > 0 Then
            ReDim Preserve .IsoData(.IsoLines)
            
            .MinMW = glHugeOverExp
            .MaxMW = 0
            .MinAbu = glHugeOverExp
            .MaxAbu = 0
            
            For lngIndex = 1 To .IsoLines
            If .IsoData(lngIndex).Abundance < .MinAbu Then .MinAbu = .IsoData(lngIndex).Abundance
                If .IsoData(lngIndex).Abundance > .MaxAbu Then .MaxAbu = .IsoData(lngIndex).Abundance
                
                If intColumnMapping(IsosFileColumnConstants.MonoPlus2Abundance) >= 0 And Not blnMonoPlus2DataPresent Then
                    If .IsoData(lngIndex).IntensityMonoPlus2 > 0 Then blnMonoPlus2DataPresent = True
                End If
                
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
         
      
        If Not blnValidScanFile Then
            If mScanInfoCount > 0 Then
                ReDim Preserve .ScanInfo(mScanInfoCount)
            Else
                ReDim .ScanInfo(0)
            End If
        End If
        
    End With
    
    If Not blnValidScanFile Then
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

Private Function ReadCSVIsosFileWork(ByRef fso As FileSystemObject, ByVal strIsosFilePath As String, ByRef lngTotalBytesRead As Long, ByRef intColumnMapping() As Integer, ByVal blnValidScanFile As Boolean) As Long
    
    Dim tsInFile As TextStream
    Dim strLineIn As String
    
    Dim lngLinesRead As Long
    Dim lngIndex As Long
    Dim lngScanNumber As Long
    
    Dim blnColumnsDefined As Boolean
    Dim blnDataLine As Boolean
    Dim blnValidDataPoint As Boolean
    Dim blnStoreDataPoint As Boolean
    
    Dim strData() As String
    Dim strColumnHeader As String
    
    Dim sngFit As Single
    Dim sngAbundance As Single
    
    Dim strUnknownColumnList As String
    Dim strMessage As String
    Dim lngReturnValue As Long
    
On Error GoTo ReadCSVIsosFileWorkErrorHandler
    
    lngLinesRead = 0
    
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

                    ' Column headers were not present
                     AddToAnalysisHistory mGelIndex, "Isos file " & fso.GetFileName(strIsosFilePath) & " did not contain column headers; using the default headers (" & GetDefaultIsosColumnHeaders(False) & ")"

                    blnDataLine = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, ",")
                    strUnknownColumnList = ""
                    
                    For lngIndex = 0 To UBound(strData)
                        If lngIndex >= ISOS_FILE_COLUMN_COUNT Then
                            ' Too many column headers
                            Debug.Assert False
                            Exit For
                        End If
                        
                        strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                        
                        Select Case strColumnHeader
                        Case ISOS_COLUMN_SCAN_NUM: intColumnMapping(IsosFileColumnConstants.ScanNumber) = lngIndex
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
                        AddToAnalysisHistory mGelIndex, CSV_COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & fso.GetFileName(strIsosFilePath) & ": " & strUnknownColumnList & "; Known columns are: " & vbCrLf & GetDefaultIsosColumnHeaders(False)
                    End If
                    
                    blnDataLine = False
                
                End If
                
                ' Warn the user if any of the important columns are missing
                If intColumnMapping(IsosFileColumnConstants.ScanNumber) < 0 Or _
                   intColumnMapping(IsosFileColumnConstants.Abundance) < 0 Or _
                   intColumnMapping(IsosFileColumnConstants.MonoisotopicMW) < 0 Then
                   
                    If mReadMode = rmStoreData Then
                        strMessage = CSV_COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strIsosFilePath) & "; the expected columns are: " & vbCrLf & GetDefaultIsosColumnHeaders(True)
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
                
                blnValidDataPoint = True
                If sngFit <= mMaxFit Or mMaxFit <= 0 Then
                    If mFilterByAbundance Then
                        If sngAbundance < mAbundanceMin Or sngAbundance > mAbundanceMax Then
                            blnValidDataPoint = False
                        End If
                    End If
                Else
                    blnValidDataPoint = False
                End If

                If blnValidDataPoint Then
                
                    If mReadMode = rmReadModeConstants.rmPrescanData Then
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
                                .DataLines = .DataLines + 1
                                .IsoLines = .IsoLines + 1
                        
                                If Not blnValidScanFile Then
                                    ' Possibly add a new entry to .ScanInfo()
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
                                End If
                               
                                With .IsoData(.IsoLines)
                                    .ScanNumber = lngScanNumber
                                    .Charge = CInt(GetColumnValueLng(strData, intColumnMapping(IsosFileColumnConstants.Charge), 1))
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
                                End With
                            End With
                        End If
                    End If
                    
                    mValidDataPointCount = mValidDataPointCount + 1
                End If
            End If
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

Private Function ReadCSVScanFile(ByRef fso As FileSystemObject, ByVal strScansFilePath As String, ByVal strBaseName As String, ByRef lngTotalBytesRead As Long) As Long
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
    
    If Len(strBaseName) = 0 Then
        strBaseName = fso.GetBaseName(strScansFilePath)
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

                        ' Column headers were not present
                         AddToAnalysisHistory mGelIndex, "Scans file " & fso.GetFileName(strScansFilePath) & " did not contain column headers; using the default headers (" & GetDefaultScansColumnHeaders(False) & ")"

                        blnDataLine = True
                    Else
                        ' Define the column mappings
                        strData = Split(strLineIn, ",")
                        strUnknownColumnList = ""
                        
                        For lngIndex = 0 To UBound(strData)
                            If lngIndex >= SCAN_FILE_COLUMN_COUNT Then
                                ' Too many column headers
                                Debug.Assert False
                                Exit For
                            End If

                            strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                            
                            Select Case strColumnHeader
                            Case SCANS_COLUMN_SCAN_NUM: intColumnMapping(ScanFileColumnConstants.ScanNumber) = lngIndex
                            Case SCANS_COLUMN_TIME_A, SCANS_COLUMN_TIME_B: intColumnMapping(ScanFileColumnConstants.ScanTime) = lngIndex
                            Case SCANS_COLUMN_TYPE: intColumnMapping(ScanFileColumnConstants.ScanType) = lngIndex
                            Case SCANS_COLUMN_NUM_DEISOTOPED: intColumnMapping(ScanFileColumnConstants.NumDeisotoped) = lngIndex
                            Case SCANS_COLUMN_NUM_PEAKS: intColumnMapping(ScanFileColumnConstants.NumPeaks) = lngIndex
                            Case SCANS_COLUMN_TIC: intColumnMapping(ScanFileColumnConstants.TIC) = lngIndex
                            Case SCANS_COLUMN_BPI_MZ: intColumnMapping(ScanFileColumnConstants.BPImz) = lngIndex
                            Case SCANS_COLUMN_BPI: intColumnMapping(ScanFileColumnConstants.BPI) = lngIndex
                            Case SCANS_COLUMN_TIME_DOMAIN_SIGNAL: intColumnMapping(ScanFileColumnConstants.TimeDomainSignal) = lngIndex
                            Case SCANS_COLUMN_PEAK_INTENSITY_THRESHOLD: intColumnMapping(ScanFileColumnConstants.PeakIntensityThreshold) = lngIndex
                            Case SCANS_COLUMN_PEPTIDE_INTENSITY_THRESHOLD: intColumnMapping(ScanFileColumnConstants.PeptideIntensityThreshold) = lngIndex
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
                            AddToAnalysisHistory mGelIndex, CSV_COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & fso.GetFileName(strScansFilePath) & ": " & strUnknownColumnList & "; Known columns are: " & vbCrLf & GetDefaultScansColumnHeaders(False)
                        End If
                        
                        blnDataLine = False
                    
                    End If
                    
                    ' Warn the user if any of the important columns are missing
                    If intColumnMapping(ScanFileColumnConstants.ScanNumber) < 0 Or _
                       intColumnMapping(ScanFileColumnConstants.ScanTime) < 0 Or _
                       intColumnMapping(ScanFileColumnConstants.ScanType) < 0 Then
                       
                        strMessage = CSV_COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strScansFilePath) & "; the expected columns are: " & vbCrLf & GetDefaultScansColumnHeaders(True)
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
    
                        .ScanFileName = strBaseName & "." & Format(.ScanNumber, "00000")
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

Private Function ResolveCSVFilePaths(ByVal strFilePath As String, ByRef strScansFilePath As String, ByRef strIsosFilePath As String, ByRef strBaseName As String) As Boolean
    ' Define the _scans.csv and _isos.csv FilePaths, given strFilePath
    ' strFilePath could contain either the _scans.csv name or the _isos.csv name
    ' Does not confirm that the files actually exist
    
    
    Dim lngCharLoc As Long
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    strScansFilePath = ""
    strIsosFilePath = ""
    strBaseName = ""
    
    lngCharLoc = InStr(LCase(strFilePath), LCase(CSV_SCANS_FILE_SUFFIX))
    If lngCharLoc >= 1 Then
        ' strFilePath contains the _scans.csv file to this function
        ' Look for the corresponding _isos.csv file
        strScansFilePath = strFilePath
        strBaseName = Left(strFilePath, lngCharLoc - 1)
        strIsosFilePath = strBaseName & CSV_ISOS_IC_FILE_SUFFIX
        
        If Not FileExists(strIsosFilePath) Then
            strIsosFilePath = strBaseName & CSV_ISOS_FILE_SUFFIX
        End If
        blnSuccess = True
        
    Else
        ' Assume strFilePath contains the _isos.csv file (or similar)
        ' Look for the Scans.csv file
        
        ' Define the base name
        ' First look for isos_ic.csv
        lngCharLoc = InStr(LCase(strFilePath), LCase(CSV_ISOS_IC_FILE_SUFFIX))
        If lngCharLoc < 1 Then
            ' No match, look for isos.csv
            lngCharLoc = InStr(LCase(strFilePath), LCase(CSV_ISOS_FILE_SUFFIX))
            If lngCharLoc < 1 Then
                ' No match
            End If
        End If
        
        If lngCharLoc >= 1 Then
            strIsosFilePath = strFilePath
            
            strBaseName = Left(strFilePath, lngCharLoc - 1)
            strScansFilePath = strBaseName & CSV_SCANS_FILE_SUFFIX
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

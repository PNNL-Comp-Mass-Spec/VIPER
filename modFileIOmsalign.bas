Attribute VB_Name = "modFileIOmsalign"
Option Explicit

Public Const MSALIGN_FILE_SUFFIX As String = ".msalign"

Private Const MSDECONV_FILE_SUFFIX As String = "_msdeconv"

' Note: These should all be upper case string values
Private Const MSALIGN_TAG_BEGIN_IONS As String = "BEGIN IONS"
Private Const MSALIGN_TAG_ID As String = "ID"
Private Const MSALIGN_TAG_SCANS As String = "SCANS"
Private Const MSALIGN_TAG_END_IONS As String = "END IONS"

Private Const SCAN_INFO_DIM_CHUNK As Long = 10000
Private Const ISO_DATA_DIM_CHUNK As Long = 25000

Private Enum rmReadModeConstants
    rmPrescanData = 0
    rmStoreData = 1
    rmReadComplete = 2
End Enum

Private Type udtIsoDataCurrentScanType
    CurrentDataLine As Long
    IsoDataIndex As Long
    FeatureIndex As Long
End Type

Private mGelIndex As Long
Private mScanInfoCount As Long

Private mEvenOddScanFilter As Boolean
Private mEvenOddModCompareVal As Integer

Private mFilterByAbundance As Boolean
Private mAbundanceMin As Double
Private mAbundanceMax As Double

Private mMaximumDataCountEnabled As Boolean
Private mMaximumDataCountToLoad As Long

Private mTotalIntensityPercentageFilterEnabled As Boolean
Private mTotalIntensityPercentageFilter As Single

Private mPrescannedData As clsFileIOPrescannedData

Private mValidDataPointCount As Long
Private mSubtaskMessage As String

Private mReadMode As rmReadModeConstants
Private mCurrentProgressStep As Integer

Public Function GetDatasetNameFromMSAlignFilename(ByVal strFilePath As String) As String
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim strBase As String
    
    strBase = ""

    strFileName = fso.GetFileName(strFilePath)
    
    If StringEndsWith(strFileName, MSALIGN_FILE_SUFFIX) Then
        strBase = StringTrimEnd(strFileName, MSALIGN_FILE_SUFFIX)
    Else
        strBase = fso.GetBaseName(strFileName)
    End If
    
    Set fso = Nothing
    
    If StringEndsWith(strBase, MSDECONV_FILE_SUFFIX) Then
        strBase = StringTrimEnd(strBase, MSDECONV_FILE_SUFFIX)
    End If
    
    GetDatasetNameFromMSAlignFilename = strBase
    
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

Public Function LoadNewMSAlign(ByVal MSAlignFilePath As String, ByVal lngGelIndex As Long, _
                                ByVal blnFilterByAbundance As Boolean, _
                                ByVal dblMinAbu As Double, ByVal dblMaxAbu As Double, _
                                ByVal blnMaximumDataCountEnabled As Boolean, ByVal lngMaximumDataCountToLoad As Long, _
                                ByVal blnTotalIntensityPercentageFilterEnabled, ByVal sngTotalIntensityPercentageFilter, _
                                ByVal eScanFilterMode As eosEvenOddScanFilterModeConstants, _
                                ByVal eDataFilterMode As dfmCSandIsoDataFilterModeConstants, _
                                ByRef strErrorMessage) As Long
                           
    '-------------------------------------------------------------------------
    'Returns 0 if data successfuly loaded, -2 if data set is too large,
    '-3 if problems with scan numbers, -4 if no data found, -5 if user cancels load,
    '-6 for file not found or invalid file
    '-7 for file problem that user was already notified about
    '-10 for any other error
    'Set blnFilterByAbundance to True to use dblMinAbu and dblMaxAbu to filter the abundance values
    'eDataFilterMode is ignored by this function
    '-------------------------------------------------------------------------
    
    Dim intProgressCount As Integer
    Dim blnFilePrescanEnabled As Boolean
    
    Dim strMessage As String
    
    Dim eResponse As VbMsgBoxResult
    
    Dim fso As New FileSystemObject
    Dim objFile As Object
    Dim tsInFile As TextStream
    Dim strLineIn As String
    
    Dim blnSuccess As Boolean
    
    Dim lngCharLoc As Long
    Dim lngReturnValue As Long
    Dim lngIndex As Long
    
    Dim dblByteCountTotal As Double
    Dim dblTotalBytesRead As Double
    
    ' This HashTable maps the the line number of the data point in the input file with the index in GelData(mGelIndex).IsoData() that the data point is stored
    Dim objHashMapOfPointsKept As clsParallelLngArrays

On Error GoTo LoadNewMSAlignErrorHandler

    ' Update the filter variables
    mGelIndex = lngGelIndex
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
    
    ' Validate that the input file(s) exist
    If Not fso.FileExists(MSAlignFilePath) Then
        strErrorMessage = "Error: MSAlign file not found: " & vbCrLf & MSAlignFilePath
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
        End If
        AddToAnalysisHistory mGelIndex, strErrorMessage
        LoadNewMSAlign = -7
        Exit Function
    End If
    
    ' Initialize the hash tables
    Set objHashMapOfPointsKept = New clsParallelLngArrays
    
    ' Initialize the even/odd scan filter variables
    mEvenOddScanFilter = False
    
    If eScanFilterMode = eosLoadOddScansOnly Then
        mEvenOddScanFilter = True
        mEvenOddModCompareVal = 1                     ' Use scans where Scan Mod 2 = 1
    ElseIf eScanFilterMode = eosLoadEvenScansOnly Then
        mEvenOddScanFilter = True
        mEvenOddModCompareVal = 0                     ' Use scans where Scan Mod 2 = 0
    End If
    
    On Error Resume Next
    
    ' Initialize the progress bar
    dblTotalBytesRead = 0
    dblByteCountTotal = -1
    
    Set objFile = fso.GetFile(MSAlignFilePath)
    dblByteCountTotal = objFile.Size
    
    frmProgress.InitializeSubtask "Reading data", 0, 100
    
    On Error GoTo LoadNewMSAlignErrorHandler
    
    mScanInfoCount = 0
    ReDim GelData(mGelIndex).ScanInfo(0)
    
    GelData(mGelIndex).DataStatusBits = 0
    
    ' Read the .msalign file
    ' Note that the msalign file only contains isotopic data, not charge state data
    lngReturnValue = ReadMSAlignFile(fso, MSAlignFilePath, _
                                     dblByteCountTotal, _
                                     dblTotalBytesRead, _
                                     blnFilePrescanEnabled, _
                                     objHashMapOfPointsKept)
     
    LoadNewMSAlign = lngReturnValue
    Exit Function

LoadNewMSAlignErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "LoadNewMSAlign"
    
    strErrorMessage = "Error in LoadNewMSAlignErrorHandler: " & Err.Description
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    LoadNewMSAlign = lngReturnValue
    
End Function

Private Function ReadMSAlignFile(ByRef fso As FileSystemObject, ByVal strMSAlignFilePath As String, _
                                 ByVal dblByteCountTotal As Double, _
                                 ByRef dblTotalBytesRead As Double, _
                                 ByVal blnFilePrescanEnabled As Boolean, _
                                 ByRef objHashMapOfPointsKept As clsParallelLngArrays) As Long

    ' Returns 0 if no error, the error number if an error

    Dim lngIndex As Long
    Dim lngNewIsoDataCount As Long
    Dim lngDataCountUpdated As Long
    Dim lngDataPointsAdded As Long
    Dim lngReturnValue As Long
    
    Dim objFile As File
    Dim objFolder As Folder
    
    Dim MaxMZ As Double
    
On Error GoTo ReadMSAlignFileErrorHandler
    
    With GelData(mGelIndex)
        mScanInfoCount = 0
        ReDim .ScanInfo(SCAN_INFO_DIM_CHUNK)
        
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
            
            mSubtaskMessage = "Pre-scanning MSAlign file to determine data to load"
        Else
            mSubtaskMessage = "Reading MSAlign file"
            
        End If
        
        frmProgress.InitializeSubtask mSubtaskMessage, 0, 100
        
        ' Reset the tracking variables
        mValidDataPointCount = 0
        dblTotalBytesRead = 0
    
        lngReturnValue = ReadMSAlignFileWork(fso, strMSAlignFilePath, dblTotalBytesRead, dblByteCountTotal, _
                                             blnFilePrescanEnabled, _
                                             objHashMapOfPointsKept)
        
        If lngReturnValue <> 0 Then
            ' Error occurred
            Debug.Assert False
            ReadMSAlignFile = lngReturnValue
            Exit Function
        End If
        
        If mReadMode = rmReadModeConstants.rmPrescanData Then
            If KeyPressAbortProcess > 1 Then
                ' User Cancelled Load
                ReadMSAlignFile = -5
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
       ReadMSAlignFile = -4
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
            
        .DataStatusBits = .DataStatusBits And Not GEL_DATA_STATUS_BIT_IREPORT
   
        .DataFilter(fltCSAbu, 2) = .MaxAbu             'put initial filters on max
        .DataFilter(fltIsoAbu, 2) = .MaxAbu
        .DataFilter(fltCSMW, 2) = .MaxMW
        .DataFilter(fltIsoMW, 2) = .MaxMW
        .DataFilter(fltIsoMZ, 2) = MaxMZ
        
        .DataFilter(fltEvenOddScanNumber, 0) = False
        .DataFilter(fltEvenOddScanNumber, 1) = 0       ' Show all scan numbers
        
        .DataFilter(fltIsoCS, 2) = 1000                'maximum charge state
         
      
        If mScanInfoCount > 0 Then
            ReDim Preserve .ScanInfo(mScanInfoCount)
        Else
            ReDim .ScanInfo(0)
        End If
        
    End With
    
    ' Elution time wasn't defined
    ' Define the default elution time to range from 0 to 1
    DefineDefaultElutionTimes GelData(mGelIndex).ScanInfo, 0, 1
    
    UpdateGelAdjacentScanPointerArrays mGelIndex
    
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load
        ReadMSAlignFile = -5
        frmProgress.HideForm
        Exit Function
    End If
    
    ' Update the progress bar
    mCurrentProgressStep = mCurrentProgressStep + 1
    frmProgress.UpdateProgressBar mCurrentProgressStep
    
    ' Sort the data
    frmProgress.InitializeSubtask "Sorting isotopic data", 0, GelData(mGelIndex).IsoLines
    SortIsotopicData mGelIndex
        
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load
        ReadMSAlignFile = -5
        frmProgress.HideForm
        Exit Function
    End If
    
    ReadMSAlignFile = 0
    Exit Function

ReadMSAlignFileErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadMSAlignFile"
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    ReadMSAlignFile = lngReturnValue

End Function

Private Function ReadMSAlignFileWork(ByRef fso As FileSystemObject, _
                                     ByVal strMSAlignFilePath As String, _
                                     ByRef dblTotalBytesRead As Double, _
                                     ByVal dblByteCountTotal As Double, _
                                     ByVal blnFilePrescanEnabled As Boolean, _
                                     ByRef objHashMapOfPointsKept As clsParallelLngArrays) As Long
    
    Dim tsInFile As TextStream
    Dim strLineIn As String
    
    Dim lngCurrentDataLine As Long
    
    Dim blnInsideIonBlock As Boolean
    Dim intCharLoc As Integer
    
    Dim lngCurrentID As Long
    Dim lngCurrentScan As Long
    Dim lngErrorCount As Long

    Dim strKeyName As String
    Dim strValue As String
    
    Dim blnValidDataPoint As Boolean
    
    Dim lngScanNumber As Long
    
    Dim blnStoreDataPoint As Boolean

    Dim strData() As String
    Dim dblMonoisotopicMass As Double
    Dim sngAbundance As Single
    Dim intCharge As Integer

    Dim lngReturnValue As Long
                                                                      
On Error GoTo ReadMSAlignFileWorkErrorHandler
    
    ' Make sure objHashMapOfPointsKept is empty
    objHashMapOfPointsKept.Clear
    
    lngErrorCount = 0
    
    ' Need to start at 1 to remain consistent with .IsoLines starting at 1
    ' and to remain consistent with objPointsToKeep
    lngCurrentDataLine = 1
    
    Set tsInFile = fso.OpenTextFile(strMSAlignFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream

        strLineIn = tsInFile.ReadLine
        dblTotalBytesRead = dblTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If lngCurrentDataLine Mod 500 = 0 Then
            If dblByteCountTotal > 0 Then
                frmProgress.UpdateSubtaskProgressBar CSng(dblTotalBytesRead / dblByteCountTotal * 100#), True
            Else
                DoEvents
            End If
            If KeyPressAbortProcess > 1 Then Exit Do
        End If
        
        strLineIn = Trim(strLineIn)
        
        ' See if line starts with a known tag
        If strLineIn = MSALIGN_TAG_BEGIN_IONS Then
            blnInsideIonBlock = True
            lngCurrentID = -1
            lngCurrentScan = -1
        ElseIf strLineIn = MSALIGN_TAG_END_IONS Then
            blnInsideIonBlock = False
        ElseIf blnInsideIonBlock And Len(strLineIn) > 0 Then
            
            ' See if line contains an equals sign
            intCharLoc = InStr(strLineIn, "=")
            If intCharLoc > 0 Then
                strKeyName = Left(strLineIn, intCharLoc - 1)
                strValue = Mid(strLineIn, intCharLoc + 1)
                
                Select Case strKeyName
                Case MSALIGN_TAG_ID
                    lngCurrentID = CLngSafe(strValue)
                Case MSALIGN_TAG_SCANS
                    If IsNumeric(strValue) Then
                        lngCurrentScan = CLngSafe(strValue)
                    Else
                        lngErrorCount = lngErrorCount + 1
                        If lngErrorCount < 25 Then
                            AddToAnalysisHistory mGelIndex, "Error: Non-numeric value after SCANS tag: " & strLineIn
                        End If
                    End If
                Case Else
                    ' Unknown key; skip it
                End Select
            Else
                ' No equals sign, does line start with a number?
                If IsNumeric(Left(strLineIn, 1)) Then
                    ' Yes, this is a data line
                    ' We can only continue if we know the scan number
                    
                    blnValidDataPoint = False
                    
                    If lngCurrentScan < 0 Then
                        lngErrorCount = lngErrorCount + 1
                        If lngErrorCount < 25 Then
                            AddToAnalysisHistory mGelIndex, "Error: Numeric line found, but we don't yet know the scan number: " & strLineIn
                        End If
                    Else
                        lngScanNumber = lngCurrentScan
                        
                        strData = Split(strLineIn, vbTab)
                        
                        If UBound(strData) > 0 Then
                            ' Valid data line
                            blnValidDataPoint = True
                        End If
                    End If
                    
                    If blnValidDataPoint And (Not mEvenOddScanFilter Or (lngScanNumber Mod 2 = mEvenOddModCompareVal)) Then
        
                        
                        dblMonoisotopicMass = CDbl(strData(0))
                        sngAbundance = CSng(strData(1))
                        intCharge = CInt(strData(2))
                        
                        Debug.Assert intCharge > 0
                                                
                        If mFilterByAbundance Then
                            If sngAbundance < mAbundanceMin Or sngAbundance > mAbundanceMax Then
                                blnValidDataPoint = False
                            End If
                        End If
                   
                        If blnValidDataPoint Then
        
                            If mReadMode = rmReadModeConstants.rmPrescanData Then
                                mPrescannedData.AddDataPoint sngAbundance, intCharge, mValidDataPointCount
                            Else
                                
                                If blnFilePrescanEnabled Then
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
                                
                                        ' Possibly add a new entry to .ScanInfo()
                                        ' Assumes data in the .msalign file is sorted by ascending scan number
                                        
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
                        
                                        If .IsoLines > UBound(.IsoData) Then
                                            If .IsoLines >= 2200000 Then
                                                ' Maximum array length reached; if we try to ReDim more memory, we'll get error 438
                                                ' Must abort loading any new data
                                                
                                                AddToAnalysisHistory mGelIndex, "Error: Maximum number of supported Isotopic Data points has been loaded (2200000); aborting load"
                                                Debug.Assert False
                                                
                                                Exit Do
                                            End If
                                            
                                            ReDim Preserve .IsoData(UBound(.IsoData) + ISO_DATA_DIM_CHUNK)
                                            
                                        End If
                                       
                                        With .IsoData(.IsoLines)
                                            .ScanNumber = lngScanNumber
                                            .Charge = intCharge
                                            .Abundance = sngAbundance
                                            .MZ = ConvoluteMass(dblMonoisotopicMass, 0, intCharge)
                                            .Fit = 0
                                            .AverageMW = dblMonoisotopicMass
                                            .MonoisotopicMW = dblMonoisotopicMass
                                            .MostAbundantMW = dblMonoisotopicMass
                                            .FWHM = 0
                                            .SignalToNoise = 0
                                            .IntensityMono = 0
                                            .IntensityMonoPlus2 = 0
                                            .IMSDriftTime = 0
                                            
                                            .SaturationFlag = 0
                                            
                                        End With
                                    
                                    End With
    
                                    ' Keep track of the mapping between the line number of the data point in the input file
                                    ' and the index value in GelData(mGelIndex).IsoData() where this data point has been stored
                                    objHashMapOfPointsKept.add lngCurrentDataLine, GelData(mGelIndex).IsoLines
                                                                    
                                End If
                            End If
                            
                            mValidDataPointCount = mValidDataPointCount + 1
                        End If


                    End If
                End If
            End If
        End If
        
            
        lngCurrentDataLine = lngCurrentDataLine + 1
    
        If mReadMode = rmReadModeConstants.rmStoreData Then
            ' Update .LinesRead
            GelData(mGelIndex).LinesRead = lngCurrentDataLine
        End If
        
    Loop
    
    If mReadMode <> rmReadModeConstants.rmPrescanData Then
        AddToAnalysisHistory mGelIndex, "Processed " & Format(GelData(mGelIndex).LinesRead, "0,000") & " data lines; retained " & Format(mValidDataPointCount, "0,000") & " data points"
    
        ' Sort the data in objHashMapOfPointsKept
        objHashMapOfPointsKept.SortNow
    End If
    
    tsInFile.Close
    ReadMSAlignFileWork = 0
    Exit Function

ReadMSAlignFileWorkErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadMSAlignFileWork"
    
    If Err.Number = 438 Then
    End If
    
    If Not tsInFile Is Nothing Then
        On Error Resume Next
        tsInFile.Close
    End If
    
    ReadMSAlignFileWork = lngReturnValue
End Function


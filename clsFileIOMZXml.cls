VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileIOMZXml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Copyright 2006, Battelle Memorial Institute.  All Rights Reserved.
' Program started March 24, 2006
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------

Private Const SCAN_INFO_DIM_CHUNK As Long = 10000
Private Const ISO_DATA_DIM_CHUNK As Long = 25000
Private Const MAX_MS_LEVEL As Integer = 9

Private Const SKIP_DATA_POINT_FLAG As Single = -1

Private Enum rmReadModeConstants
    rmPrescanData = 0
    rmStoreData = 1
    rmReadComplete = 2
End Enum

Private WithEvents mMZXmlReader As MZXmlFileReaderDLL.mzXMLReader
Attribute mMZXmlReader.VB_VarHelpID = -1

Private mFirstScanRead As Boolean
Private mReadMode As rmReadModeConstants
Private mCurrentProgressStep As Integer

Private mScanInfoCount As Long
Private mValidDataPointCount As Long
Private mSubtaskMessage As String

Private mEvenOddScanFilter As Boolean
Private mEvenOddModCompareVal As Integer

Private mMaxElutionTime As Single

Private mGelIndex As Long

Private mMSLevelFilter() As Boolean             ' Used to filter which MSLevels to use; mMSLevelFilter(0) means to use all levels
Private mFilterByAbundance As Boolean
Private mAbundanceMin As Double
Private mAbundanceMax As Double

Private mMaximumDataCountEnabled As Boolean
Private mMaximumDataCountToLoad As Long

Private mPrescannedData As clsFileIOPrescannedData

Private Sub InitializeLocalVariables()
    mFirstScanRead = False
    
    mScanInfoCount = 0
    mValidDataPointCount = 0
    
    mEvenOddScanFilter = False
    mEvenOddModCompareVal = 0
    
    mMaxElutionTime = 0
    
    mFilterByAbundance = False
    mAbundanceMin = 0
    mAbundanceMax = 1E+300

    ReDim mMSLevelFilter(MAX_MS_LEVEL)
    mMSLevelFilter(0) = True
    
    mMaximumDataCountEnabled = False
    mMaximumDataCountToLoad = 400000
    
End Sub

Public Function LoadNewMZXML(ByVal mzXMLFilePath As String, ByVal lngGelIndex As Long, _
                             ByVal MaxFit As Double, _
                             ByVal blnFilterByAbundance As Boolean, _
                             ByVal dblMinAbu As Double, ByVal dblMaxAbu As Double, _
                             ByVal blnMaximumDataCountEnabled As Boolean, ByVal lngMaximumDataCountToLoad As Long, _
                             ByVal eScanFilterMode As eosEvenOddScanFilterModeConstants, _
                             ByVal eDataFilterMode As dfmCSandIsoDataFilterModeConstants, _
                             ByRef blnMSLevelFilter() As Boolean) As Long
    '-------------------------------------------------------------------------
    ' Returns 0 if data successfuly loaded, -2 if data set is too large,
    ' -3 if problems with scan numbers, -4 if no data found, -5 if user cancels load,
    ' -6 for file not found or invalid file
    ' -7 for file problem that user was already notified about
    ' -10 for any other error
    ' Set mFilterByAbundance to True to use dblMinAbu and dblMaxAbu to filter the abundance values
    ' Set blnMaximumDataCountEnabled to only keep the top lngMaximumDataCountToLoad points in the input file (favoring higher abundance points)
    ' eDataFilterMode is ignored by this function
    '-------------------------------------------------------------------------
    
    Dim fso As New FileSystemObject
    Dim intIndex As Integer
    Dim intProgressCount As Integer
    
    Dim lngReturnValue As Long
    
On Error GoTo LoadNewMZXMLErrorHandler
    
    InitializeLocalVariables
    
    ' Update the filter variables
    mGelIndex = lngGelIndex
    mFilterByAbundance = blnFilterByAbundance
    mAbundanceMin = dblMinAbu
    mAbundanceMax = dblMaxAbu
    
    On Error Resume Next
    For intIndex = 0 To UBound(blnMSLevelFilter)
        If intIndex > UBound(mMSLevelFilter) Then Exit For
        mMSLevelFilter(intIndex) = blnMSLevelFilter(intIndex)
    Next intIndex
        
On Error GoTo LoadNewMZXMLErrorHandler
        
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

    ' Validate that the input file(s) exist
    If Not fso.FileExists(mzXMLFilePath) Then
        LoadNewMZXML = -6
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
    frmProgress.InitializeSubtask "Reading data", 0, 100
    
    ' Read the mzXML file and populate GelData(mGelIndex)
    lngReturnValue = ReadMZXMLFile(mzXMLFilePath)
        
    LoadNewMZXML = lngReturnValue
    Exit Function

LoadNewMZXMLErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "LoadNewMZXML"
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    LoadNewMZXML = lngReturnValue
    
End Function

Private Function ReadMZXMLFile(ByVal mzXMLFilePath As String) As Long
    ' Returns 0 if no error, the error number if an error

    Dim lngIndex As Long
    Dim lngReturnValue As Long
    
    Dim fso As New FileSystemObject
    Dim objFile As File
    Dim objFolder As Folder
    
    Dim MaxMZ As Double
    Dim blnSuccess As Boolean
    
On Error GoTo ReadMZXMLFileErrorHandler
    
    With GelData(mGelIndex)
        
        ReDim .ScanInfo(SCAN_INFO_DIM_CHUNK)
        ReDim .IsoData(ISO_DATA_DIM_CHUNK)
        
        .LinesRead = 0
        .DataLines = 0
        .CSLines = 0
        .IsoLines = 0
    End With
    
    Set mMZXmlReader = New MZXmlFileReaderDLL.mzXMLReader
    
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
            
            mSubtaskMessage = "Pre-scanning mzXML file to determine data to load"
        Else
            mSubtaskMessage = "Reading mzXML file"
        End If
        frmProgress.InitializeSubtask mSubtaskMessage, 0, 100

        ' Reset the tracking variables
        mFirstScanRead = False
        mValidDataPointCount = 0

        ' Note that mMZXmlReader raises event ScanRead for each scan read
        ' When the entire file has been read, control will return to this function
        mMZXmlReader.CacheDataInMemory = False
        blnSuccess = mMZXmlReader.ReadMZXMLFile(mzXMLFilePath)
        
        If Len(mMZXmlReader.ErrorMessage) And mReadMode < rmReadModeConstants.rmReadComplete - 1 Then
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                AddToAnalysisHistory mGelIndex, mMZXmlReader.ErrorMessage
            Else
                MsgBox "Error reading file: " & mMZXmlReader.ErrorMessage
            End If
            If mValidDataPointCount = 0 Then
                Exit Do
            End If
        End If
        
        If mReadMode = rmReadModeConstants.rmPrescanData Then
            If KeyPressAbortProcess > 1 Then
                ' User Cancelled Load
                ReadMZXMLFile = -5
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
       ReadMZXMLFile = -4
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
         ' New: data file folder path is the folder one folder up from .Filename's folder
        Set objFile = fso.GetFile(.FileName)
        Set objFolder = objFile.ParentFolder
        .PathtoDataFiles = objFolder.Path
        
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
        
        ' IReport column was not present
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
    
    If mMaxElutionTime = 0 Then
        ' Elution time wasn't defined
        ' Define the default elution time to range from 0 to 1
        DefineDefaultElutionTimes GelData(mGelIndex).ScanInfo, 0, 1
    End If
    
    UpdateGelAdjacentScanPointerArrays mGelIndex
    
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load
        ReadMZXMLFile = -5
        frmProgress.HideForm
        Exit Function
    Else
        mCurrentProgressStep = mCurrentProgressStep + 1
        frmProgress.UpdateProgressBar mCurrentProgressStep
        frmProgress.InitializeSubtask "Sorting isotopic data", 0, GelData(mGelIndex).IsoLines
    End If
    
    SortIsotopicData mGelIndex
    
    If KeyPressAbortProcess > 1 Then
        ' User Cancelled Load
        ReadMZXMLFile = -5
        frmProgress.HideForm
        Exit Function
    End If
    
    Set mMZXmlReader = Nothing
    
    ReadMZXMLFile = 0
    Exit Function

ReadMZXMLFileErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadMZXMLFile"
    
    Set mMZXmlReader = Nothing
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    ReadMZXMLFile = lngReturnValue

End Function

Private Sub StoreNextScan(ByRef MZList() As Double, ByRef IntensityList() As Single, ByRef objScanInfo As MZXmlFileReaderDLL.mzXMLSpectrumInfo)
    Dim blnUseScan As Boolean
    Dim blnValidDataPoint As Boolean
    Dim blnStoreDataPoint As Boolean
    
    Dim lngDataIndex As Long
    
On Error GoTo StoreNextScanErrorHandler

    If objScanInfo.ScanNumber >= 0 And (Not mEvenOddScanFilter Or (objScanInfo.ScanNumber Mod 2 = mEvenOddModCompareVal)) Then
        ' Make sure this scan's MSLevel is valid
        
        If mMSLevelFilter(0) Then
            blnUseScan = True
        Else
            If objScanInfo.MSLevel >= 0 And objScanInfo.MSLevel <= UBound(mMSLevelFilter) Then
                blnUseScan = mMSLevelFilter(objScanInfo.MSLevel)
            Else
                blnUseScan = False
            End If
        End If
        
        If blnUseScan Then
            For lngDataIndex = 0 To UBound(MZList)
                
                blnValidDataPoint = True
                If mFilterByAbundance Then
                    If IntensityList(lngDataIndex) < mAbundanceMin Or IntensityList(lngDataIndex) > mAbundanceMax Then
                        blnValidDataPoint = False
                    End If
                End If
            
                If blnValidDataPoint Then
                    If mReadMode = rmReadModeConstants.rmPrescanData Then
                        mPrescannedData.AddDataPoint IntensityList(lngDataIndex), mValidDataPointCount
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
                        
                                ' Possibly add a new entry to .ScanInfo()
                                If mScanInfoCount = 0 Then
                                    mScanInfoCount = 1
                                    With .ScanInfo(1)
                                        .ScanNumber = objScanInfo.ScanNumber
                                        .ElutionTime = objScanInfo.RetentionTimeMin
                                        .ScanType = objScanInfo.MSLevel
                                        
                                        mMaxElutionTime = .ElutionTime
                                    End With
                                Else
                                    If .ScanInfo(mScanInfoCount).ScanNumber < objScanInfo.ScanNumber Then
                                        mScanInfoCount = mScanInfoCount + 1
                        
                                        If mScanInfoCount > UBound(.ScanInfo) Then
                                            ReDim Preserve .ScanInfo(UBound(.ScanInfo) + SCAN_INFO_DIM_CHUNK)
                                        End If
                        
                                        With .ScanInfo(mScanInfoCount)
                                            .ScanNumber = objScanInfo.ScanNumber
                                            .ElutionTime = objScanInfo.RetentionTimeMin
                                            .ScanType = objScanInfo.MSLevel
                                        
                                            If .ElutionTime > mMaxElutionTime Then
                                                mMaxElutionTime = .ElutionTime
                                            End If
                                        End With
                                    End If
                                End If
                        
                                If .IsoLines > UBound(.IsoData) Then
                                    ReDim Preserve .IsoData(UBound(.IsoData) + ISO_DATA_DIM_CHUNK)
                                End If
                        
                                With .IsoData(.IsoLines)
                                    .ScanNumber = objScanInfo.ScanNumber
                                    .Charge = 1
                                    .Abundance = IntensityList(lngDataIndex)
                                    .MZ = MZList(lngDataIndex)
                                    .Fit = 0
                                    .AverageMW = MZList(lngDataIndex)
                                    .MonoisotopicMW = MZList(lngDataIndex)
                                    .MostAbundantMW = MZList(lngDataIndex)
                                    .FWHM = 0
                                    .SignalToNoise = 1
                                    .IntensityMono = 0
                                    .IntensityMonoPlus2 = 0
                                End With
                            End With
                        End If
                    End If
                    
                    ' Increment mValidDataPointCount
                    mValidDataPointCount = mValidDataPointCount + 1
                    
                End If
            Next lngDataIndex
        End If
    End If

    Exit Sub

StoreNextScanErrorHandler:
    Debug.Assert False
End Sub

Private Sub Class_Initialize()
    InitializeLocalVariables
End Sub

Private Sub Class_Terminate()
    Set mMZXmlReader = Nothing
End Sub

Private Sub mMZXmlReader_ScanRead()
    Dim MZList() As Double
    Dim IntensityList() As Single
    Dim objScanInfo As MZXmlFileReaderDLL.mzXMLSpectrumInfo
    
    If Not mFirstScanRead Then
        frmProgress.InitializeSubtask mSubtaskMessage, 0, mMZXmlReader.FileInfoScanCount
        mFirstScanRead = True
    End If

    frmProgress.UpdateSubtaskProgressBar mMZXmlReader.ScanNumber
    DoEvents
    
    If KeyPressAbortProcess > 1 Then
        mMZXmlReader.AbortProcessingNow
        Exit Sub
    Else
        mMZXmlReader.GetCurrentScanData MZList, IntensityList, objScanInfo
        StoreNextScan MZList, IntensityList, objScanInfo
    End If
    
End Sub



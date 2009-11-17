VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileIOPredefinedLCMSFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Copyright 2009, Battelle Memorial Institute.  All Rights Reserved.
' Class created August 1, 2009
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------

' This class reads pre-defined LC/MS features from a text file, populating the data structures for the given Gel
' It requires that the Decon2LS _isos.csv and _scans.csv files have already been loaded into memory

' It will read two files:
'  _Features.txt file (see below for valid column headers)
'  _FeatureToPeakMap.txt file (see below for valid column headers)

Private Const COLUMN_HEADER_UNKNOWN_WARNING As String = "Warning: unknown column headers"
Private Const COLUMN_HEADER_MISSING_WARNING As String = "Warning: expected important column headers"


Private Const LCMSFEATURES_COLUMN_FEATURE_INDEX As String = "feature_index"
Private Const LCMSFEATURES_COLUMN_SCAN As String = "scan"
Private Const LCMSFEATURES_COLUMN_SCAN_START As String = "scan_start"
Private Const LCMSFEATURES_COLUMN_SCAN_END As String = "scan_end"
Private Const LCMSFEATURES_COLUMN_SCAN_ALIGNED As String = "scan_aligned"
Private Const LCMSFEATURES_COLUMN_NET As String = "net"
Private Const LCMSFEATURES_COLUMN_MONO_MASS As String = "monoisotopic_mass"
Private Const LCMSFEATURES_COLUMN_MONO_MASS_CALIBRATED As String = "monoisotopic_mass_calibrated"
Private Const LCMSFEATURES_COLUMN_ABUNDANCE As String = "abundance"
Private Const LCMSFEATURES_COLUMN_CLASS_REP_MZ As String = "class_rep_mz"
Private Const LCMSFEATURES_COLUMN_CLASS_REP_CHARGE As String = "class_rep_charge"

Private Const PEAK_MAP_COLUMN_FEATURE_INDEX As String = "feature_index"
Private Const PEAK_MAP_COLUMN_PEAK_INDEX As String = "peak_index"

Private Const LCMSFEATURES_FILE_COLUMN_COUNT As Integer = 11
Private Enum LCMSFeaturesFileColumnConstants
    FeatureIndex = 0
    Scan = 1
    ScanStart = 2
    ScanEnd = 3
    ScanAligned = 4
    NET = 5
    MonoisotopicMass = 6
    MonoisotopicMassCalibrated = 7
    Abundance = 8
    ClassRepMZ = 9
    ClassRepCharge = 10
End Enum

Private Const PEAK_MAP_FILE_COLUMN_COUNT As Integer = 2
Private Enum LCMSFeatureToPeakMapFileColumnConstants
    FeatureIndex = 0
    PeakIndex = 1
End Enum

' Class-wide variables
Private mGelIndex As Integer
Private mAutoMapDataPointsMassTolerancePPM As Double

Private mFeatureToPeakMapCount As Long
Private mFeatureToPeakMapFeatureIdx() As Long
Private mFeatureToPeakMapPeakIdx() As Long          ' These indices are the index values in GelData().IsoData() -- note that .IsoData() is a 1-based array

Private mErrorMessage As String
Private mShowDialogBoxes As Boolean

Private mProgressDescription As String
Private mProgressPctComplete As Single          ' Number between 0 and 100


' Indexing schema modelled after that used in frmUMCSimple, written by Nikola Tolic in Spring 2003
' All of the arrays are 0-based

Private CSCnt As Long               'count of CS data points included in count
Private ISCnt As Long               'count of IS data points included in count

' The O_ arrays contain the data for the ions to be searched
' We could get by with just the O_Index() and O_Type() arrays, but then we'd have to
'  continually be looking up values in GelData().CSLines and GelData().IsoLines
' Copying the values to these arrays speeds up the search
'
Private O_Cnt As Long               'total number of ions to search
Private O_Index() As Long           'index in CS/Iso arrays
Private O_Type() As glDistType      'type of data(CS/Iso)
Private O_MW() As Double            'mass array
''Private O_Charge() As Double        'charge
''Private O_Order() As Double         'Abundance
Private O_Scan() As Long            'scan number

Private S_MW() As Double            'sorted mass array used for fast search

Private IndMW() As Long             'index on mass
''Private IndOrder() As Long          'index on order (ordered by decreasing abundance)

Private IsMatched() As Boolean      'is already matched

Private MWRangeFinder As MWUtil     'fast search of mass range

' Properties
Public Property Get ShowDialogBoxes() As Boolean
    ShowDialogBoxes = mShowDialogBoxes
End Property
Public Property Let ShowDialogBoxes(ByVal Value As Boolean)
    mShowDialogBoxes = Value
End Property

Public Property Get AutoMapDataPointsMassTolerancePPM() As Double
    AutoMapDataPointsMassTolerancePPM = mAutoMapDataPointsMassTolerancePPM
End Property
Public Property Let AutoMapDataPointsMassTolerancePPM(ByVal Value As Double)
    mAutoMapDataPointsMassTolerancePPM = Value
End Property

Private Sub Class_Initialize()
    mGelIndex = 0
    mAutoMapDataPointsMassTolerancePPM = 5
End Sub

' Functions

Private Function AddFeatureMembersToUMCs(ByRef lngFeatureToPeakMapFeatureIdx() As Long, _
                                         ByRef lngFeatureToPeakMapPeakIdx() As Long, _
                                         ByVal lngFeatureToPeakMapCount As Long) As Boolean
    
    ' Be sure to call SortFeatureMapMembers before calling this function because this function
    '  assumes that the data in lngFeatureToPeakMapFeatureIdx and lngFeatureToPeakMapPeakIdx
    '  is properly sorted and is unique
    
    Dim lngIndex As Long
    Dim lngIndexStart As Long
    Dim lngIndexEnd As Long
    Dim lngPeakIndexCount As Long
    
    Dim lngFeatureIndexCurrent
    
    Dim lngClassRepIndexPointer As Long
    
    Dim dblAbundance As Double
    Dim dblBestAbundance As Double
    Dim dblClassAbundance As Double
            
    Dim lngScan As Long
    Dim lngMinScan As Long
    Dim lngMaxScan As Long
    
    Dim dblMass As Double
    Dim dblMinMass As Double
    Dim dblMaxMass As Double
    
    Dim blnSuccess As Boolean
    
On Error GoTo AddFeatureMembersToUMCsErrorHandler

    ' Set this to true for now
    blnSuccess = True
    
    lngIndexStart = 0
    Do While lngIndexStart < mFeatureToPeakMapCount
    
        lngFeatureIndexCurrent = mFeatureToPeakMapFeatureIdx(lngIndexStart)
        lngIndexEnd = lngIndexStart
                
        Do While lngIndexEnd + 1 < mFeatureToPeakMapCount
            If mFeatureToPeakMapFeatureIdx(lngIndexEnd + 1) = lngFeatureIndexCurrent Then
                lngIndexEnd = lngIndexEnd + 1
            Else
                Exit Do
            End If
        Loop
    
        lngPeakIndexCount = lngIndexEnd - lngIndexStart + 1
        
        If lngPeakIndexCount = 0 Then
            ' Feature has no members; this is unexpected
            Debug.Assert False
        Else
            
            ' Determine the class rep index (i.e. the feature with the highest intensity)
            
            lngClassRepIndexPointer = lngIndexStart
            If Not GetIsoDataInfo(mFeatureToPeakMapPeakIdx(lngClassRepIndexPointer), lngScan, dblMass, dblAbundance) Then
                ' mFeatureToPeakMapPeakIdx() did not point to a valid index in GelData().IsoData()
                Debug.Assert False
            End If
            
            dblBestAbundance = dblAbundance
            dblClassAbundance = dblAbundance
            
            lngMinScan = lngScan
            lngMaxScan = lngScan
            
            dblMinMass = dblMass
            dblMaxMass = dblMass
            
            For lngIndex = lngIndexStart + 1 To lngIndexEnd
                
                If Not GetIsoDataInfo(mFeatureToPeakMapPeakIdx(lngIndex), lngScan, dblMass, dblAbundance) Then
                    ' lngClassRepIndexPointer did not point to a valid index in GelData().IsoData()
                    Debug.Assert False
                Else
                    If dblAbundance > dblBestAbundance Then
                        dblBestAbundance = dblAbundance
                        lngClassRepIndexPointer = lngIndex
                    End If
                    
                    dblClassAbundance = dblClassAbundance + dblBestAbundance
                    
                    If lngScan < lngMinScan Then lngMinScan = lngScan
                    If lngScan > lngMaxScan Then lngMaxScan = lngScan
                    
                    If dblMass < dblMinMass Then dblMinMass = dblMass
                    If dblMass > dblMaxMass Then dblMaxMass = dblMass
                
                End If
                
            Next lngIndex
            
            
            With GelUMC(mGelIndex).UMCs(lngFeatureIndexCurrent)
                .ClassRepInd = mFeatureToPeakMapPeakIdx(lngClassRepIndexPointer)
                .ClassRepType = gldtIS
                .ClassCount = lngPeakIndexCount
                ReDim .ClassMInd(lngPeakIndexCount - 1)       ' members of the UMC(including representative)         (0-based array); index of the member in GelData
                ReDim .ClassMType(lngPeakIndexCount - 1)      ' type of member: gldtCS or gldtIS
                
                For lngIndex = lngIndexStart To lngIndexEnd
                    .ClassMInd(lngIndex - lngIndexStart) = mFeatureToPeakMapPeakIdx(lngIndex)
                    .ClassMType(lngIndex - lngIndexStart) = gldtIS
                Next lngIndex
                
                ' Only update .ClassAbundance if it is 0
                If .ClassAbundance = 0 Then
                    .ClassAbundance = dblClassAbundance
                End If
                
                ' Do not update Class mass data
                '.ClassMW
                '.ClassMWStD
                '.ClassMassCorrectionDa
                
                ' Do not update these items
                '.ClassScore
                '.ClassNET
                '.ClassStatusBits
                
                ' Only update .MinScan, .MaxScan, .MinMass, or .MaxMass if the actual values are less than or greater than the recorded values
                
                If lngMinScan < .MinScan Then .MinScan = lngMinScan
                If lngMaxScan > .MaxScan Then .MaxScan = lngMaxScan
                
                If dblMinMass < .MinMW Then .MinMW = dblMinMass
                If dblMaxMass > .MaxMW Then .MaxMW = dblMaxMass
                
                ' Leave the .ChargeStateStats values unchanged
                '.ChargeStateStatsRepInd
                '.ChargeStateCount
                
                ' Leave the .PercentMembersIReport values unchanged
                '.PercentMembersIReportMonoPlus4
                '.PercentMembersIReportMonoMinus4
                
            End With
        End If
        
        lngIndexStart = lngIndexEnd + 1
    Loop
    
    AddFeatureMembersToUMCs = blnSuccess
    Exit Function

AddFeatureMembersToUMCsErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "AddFeatureMembersToUMCsErrorHandler"
    
    AddFeatureMembersToUMCs = False


End Function

Private Function AutoMapDataPointsToFeatures(ByVal dblMassTolerancePPM As Double) As Boolean
    ' dblMassTolerancePPM defines the tolerance to search when finding points that are likely mapped to the given UMC
    
    Dim lngUMCCount As Long
    Dim dblUMCAbundance() As Double
    Dim lngUMCAbuPointerArray() As Long
            
    Dim lngPointerIndex As Long
    Dim lngUMCIndex As Long
    Dim lngIndex As Long
    
    Dim lngMinScan As Long
    Dim lngMaxScan As Long
    Dim dblMassToleranceDa As Double
            
    Dim dblSearchMass As Double
            
    Dim MWRangeMinInd As Long
    Dim MWRangeMaxInd As Long
            
    Dim strMessage As String
    Dim blnSuccess As Boolean
    
On Error GoTo AutoMapDataPointsToFeaturesErrorHandler

    ' Set this to false for now
    blnSuccess = False

    If dblMassTolerancePPM <= 0 Then
        dblMassTolerancePPM = 5
    End If
    
    If GelUMC(mGelIndex).UMCCnt <= 0 Then
        mErrorMessage = "No UMCs are in memory; unable to auto-map the data points"
        AutoMapDataPointsToFeatures = False
        Exit Function
    End If
    
    ' Setup the indices that will be used for searching the data
    blnSuccess = InitializeSearchIndices()
    If blnSuccess Then
        blnSuccess = CreateIndMW()
        
        ''If blnSuccess Then
        ''    blnSuccess = CreateIndOrder()
        ''End If

    End If
    
    If blnSuccess Then
    
        ' Reset the map count
        mFeatureToPeakMapCount = 0
        
        ' Make sure all items in IsMatched() are false
        For lngIndex = 0 To O_Cnt - 1
            IsMatched(lngIndex) = False
        Next lngIndex
    
    
        ' We want to step through the UMCs by decreasing abundance
        ' Populate some arrays to help us do this
        lngUMCCount = GelUMC(mGelIndex).UMCCnt
        ReDim dblUMCAbundance(lngUMCCount - 1)
        ReDim lngUMCAbuPointerArray(lngUMCCount - 1)
        
        For lngIndex = 0 To lngUMCCount - 1
            dblUMCAbundance(lngIndex) = GelUMC(mGelIndex).UMCs(lngIndex).ClassAbundance
            lngUMCAbuPointerArray(lngIndex) = lngIndex
        Next lngIndex
        
        ShellSortDoubleWithParallelLong dblUMCAbundance, lngUMCAbuPointerArray, 0, lngUMCCount - 1
        
        ' Step through the data by descending abundance
        For lngPointerIndex = lngUMCCount - 1 To 0 Step -1
            
            lngUMCIndex = lngUMCAbuPointerArray(lngPointerIndex)
            With GelUMC(mGelIndex).UMCs(lngUMCIndex)
                lngMinScan = .MinScan
                lngMaxScan = .MaxScan
                
                dblMassToleranceDa = PPMToMass(dblMassTolerancePPM, .ClassMW)
                
                dblSearchMass = .ClassMW
                
            End With
            
            ' Look for the data points within the mass range of this UMC
            
            MWRangeMinInd = 0
            MWRangeMaxInd = O_Cnt - 1
            If MWRangeFinder.FindIndexRange(dblSearchMass, dblMassToleranceDa, MWRangeMinInd, MWRangeMaxInd) Then
                If MWRangeMaxInd >= MWRangeMinInd Then
                    ' Now determine which points are within the scan range of this UMC (and are not yet used)
                    
                    For lngIndex = MWRangeMinInd To MWRangeMaxInd
                    
                        If O_Type(IndMW(lngIndex)) = gldtIS And Not IsMatched(IndMW(lngIndex)) Then
                            
                            ' Validate that this data point's mass is within dblMassToleranceDa of the UMC's class mass
                            Debug.Assert Abs(O_MW(IndMW(lngIndex)) - dblSearchMass) <= dblMassToleranceDa
                            
                            If O_Scan(IndMW(lngIndex)) >= lngMinScan And O_Scan(IndMW(lngIndex)) <= lngMaxScan Then
                                
                                ' Data point is within the mass range and the scan range
                                ' Add to the mFeatureToPeakMap arrays
                                
                                If mFeatureToPeakMapCount >= UBound(mFeatureToPeakMapFeatureIdx) Then
                                    ' Reserve more room in the mFeatureToPeakMap arrays
                                    ReDim Preserve mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount * 2 - 1)
                                    ReDim Preserve mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount * 2 - 1)
                                End If
                                
                                mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount) = lngUMCIndex
                                mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount) = O_Index(IndMW(lngIndex))
                                mFeatureToPeakMapCount = mFeatureToPeakMapCount + 1
                                
                                IsMatched(IndMW(lngIndex)) = True
                            End If
                        Else
                            ' Data point is not an isotopic data point; or it was already matched
                            ' Do not re-use it
                        End If
                    Next lngIndex
                    
                End If
            End If
        Next lngPointerIndex
        
        ' Done searching
        
        If mFeatureToPeakMapCount <= 0 Then
            strMessage = "Error: AutoMapDataPointsToFeatures did not find any data points within the mass and scan ranges of the loaded LCMS Features"
            If mShowDialogBoxes Then
                MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
            End If
            AddToAnalysisHistory mGelIndex, strMessage
            
            blnSuccess = False
        Else
            ' Shrink the mFeatureToPeakMap arrays to length mFeatureToPeakMapCount
            ReDim Preserve mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount - 1)
            ReDim Preserve mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount - 1)
            blnSuccess = True
        End If
        
    End If
    
    AutoMapDataPointsToFeatures = blnSuccess
    Exit Function

AutoMapDataPointsToFeaturesErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "AutoMapDataPointsToFeatures"
    
    AutoMapDataPointsToFeatures = False

End Function

Private Function CreateIndMW() As Boolean
    '--------------------------------------------------------------
    'creates index on molecular mass; sorts its members and creates
    'fast search object; returns True if successful
    '--------------------------------------------------------------
    Dim qsDbl As New QSDouble
    On Error GoTo err_CreateIndMW
    mProgressDescription = "Creating MW index"
    S_MW() = O_MW()                                             'array assignment
    CreateIndMW = qsDbl.QSAsc(S_MW, IndMW)
    Set MWRangeFinder = New MWUtil
    If Not MWRangeFinder.Fill(S_MW) Then GoTo err_CreateIndMW
    
exit_CreateIndMW:
    Set qsDbl = Nothing
    Exit Function
    
err_CreateIndMW:
    Erase IndMW
    Erase S_MW
    Resume exit_CreateIndMW
    mProgressDescription = "Error creating MW index"
End Function

''Private Function CreateIndOrder() As Boolean
''    '--------------------------------------------------------------
''    'creates index on order by decreasing abundance and returns True if successful;
''    '--------------------------------------------------------------
''    Dim TmpOrder() As Double
''    Dim qsDbl As New QSDouble
''    On Error GoTo err_CreateIndOrder
''    mProgressDescription = "Creating order index"
''    TmpOrder() = O_Order()                                      'array assignment
''    CreateIndOrder = qsDbl.QSDesc(TmpOrder, IndOrder)           'higher intensity is better
''
''exit_CreateIndOrder:
''    Set qsDbl = Nothing
''    Exit Function
''
''err_CreateIndOrder:
''    Erase IndOrder
''    Resume exit_CreateIndOrder
''    mProgressDescription = "Error creating order index"
''End Function

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

Private Function GetDefaultLCMSFeatureColumnHeaders(ByVal blnRequiredColumnsOnly As Boolean) As String
    Dim strHeaders As String

    strHeaders = LCMSFEATURES_COLUMN_FEATURE_INDEX & vbTab & LCMSFEATURES_COLUMN_SCAN
   
    If Not blnRequiredColumnsOnly Then
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_SCAN_START
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_SCAN_END
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_SCAN_ALIGNED
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_NET
    End If
    
    strHeaders = strHeaders & LCMSFEATURES_COLUMN_MONO_MASS
    If Not blnRequiredColumnsOnly Then
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_MONO_MASS_CALIBRATED
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_ABUNDANCE
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_CLASS_REP_MZ
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_CLASS_REP_CHARGE
    End If
    
    GetDefaultLCMSFeatureColumnHeaders = strHeaders
End Function

Private Function GetDefaultLCMSFeatureToPeakMapColumnHeaders() As String
    Dim strHeaders As String

    strHeaders = PEAK_MAP_COLUMN_FEATURE_INDEX & vbTab & PEAK_MAP_COLUMN_PEAK_INDEX
    
    GetDefaultLCMSFeatureToPeakMapColumnHeaders = strHeaders
End Function

Private Function GetIsoDataInfo(ByVal lngDataPointIndex As Long, ByRef lngScan As Long, ByRef dblMass As Double, ByRef dblAbundance As Double) As Boolean
    On Error GoTo GetIsoDataAbundanceErrorHandler

    lngScan = 0
    dblMass = 0
    dblAbundance = 0
    
    lngScan = GelData(mGelIndex).IsoData(lngDataPointIndex).ScanNumber
    dblMass = GelData(mGelIndex).IsoData(lngDataPointIndex).MonoisotopicMW
    dblAbundance = GelData(mGelIndex).IsoData(lngDataPointIndex).Abundance
    
    GetIsoDataInfo = True
    Exit Function
    
GetIsoDataAbundanceErrorHandler:
    GetIsoDataInfo = False
    
End Function

Private Function InitializeSearchIndices() As Boolean
    ' Initializes the search indices
    ' Returns True if success, or False if failure or no data in memory
    
    ' This sub modelled after that in frmUMCSimple, written by Nikola Tolic in Spring 2003
    
    Dim MaxCnt As Long
    Dim CSInd() As Long
    Dim ISInd() As Long
    Dim i As Long
    
    On Error GoTo InitializeSearchIndicesErrorHandler
    
    mProgressDescription = "Loading arrays"
    
    MaxCnt = GelData(mGelIndex).CSLines + GelData(mGelIndex).IsoLines
    If MaxCnt > 0 Then
       ReDim O_Index(MaxCnt - 1)
       ReDim O_Type(MaxCnt - 1)
       ReDim O_MW(MaxCnt - 1)
       ''ReDim O_Charge(MaxCnt - 1)
       ''ReDim O_Order(MaxCnt - 1)
       ReDim O_Scan(MaxCnt - 1)
       
       O_Cnt = 0
       With GelData(mGelIndex)
         For i = 1 To .CSLines
            O_Index(O_Cnt) = i
            O_Type(O_Cnt) = gldtCS
            O_MW(O_Cnt) = .CSData(i).AverageMW
            ''O_Charge(O_Cnt) = .CSData(i).Charge
            O_Scan(O_Cnt) = .CSData(i).ScanNumber
            ''O_Order(O_Cnt) = .CSData(i).Abundance
            O_Cnt = O_Cnt + 1
         Next i
         
         For i = 1 To .IsoLines
            O_Index(O_Cnt) = i
            O_Type(O_Cnt) = gldtIS
            O_MW(O_Cnt) = GetIsoMass(.IsoData(i), UMCDef.MWField)
            ''O_Charge(O_Cnt) = .IsoData(i).Charge
            O_Scan(O_Cnt) = .IsoData(i).ScanNumber
            ''O_Order(O_Cnt) = .IsoData(i).Abundance
            O_Cnt = O_Cnt + 1
         Next i
         
       End With
    End If
    If O_Cnt <= 0 Then
        mProgressDescription = "No data in memory"
    End If
    
InitializeSearchIndicesErrorHandlerExit:

    If O_Cnt > 0 Then
       If O_Cnt < MaxCnt Then
            ReDim Preserve O_Index(O_Cnt - 1)
            ReDim Preserve O_Type(O_Cnt - 1)
            ReDim Preserve O_MW(O_Cnt - 1)
            ''ReDim Preserve O_Charge(O_Cnt - 1)
            ''ReDim Preserve O_Order(O_Cnt - 1)
            ReDim Preserve O_Scan(O_Cnt - 1)
        End If
        
       'initialize index arrays
       ReDim IndMW(O_Cnt - 1)
       ''ReDim IndOrder(O_Cnt - 1)
       
       For i = 0 To O_Cnt - 1
           IndMW(i) = i
           ''IndOrder(i) = i
       Next i
       ReDim IsMatched(O_Cnt - 1)
       InitializeSearchIndices = True
    Else
       Erase O_Index
       Erase O_Type
       Erase O_MW
       ''Erase O_Charge
       ''Erase O_Order
       Erase O_Scan
       InitializeSearchIndices = False
    End If
    
    Exit Function
    
InitializeSearchIndicesErrorHandler:
    Debug.Assert False
    O_Cnt = 0               'this will cause everything to be cleared
    Resume InitializeSearchIndicesErrorHandlerExit
    mErrorMessage = "Error loading arrays"
    
End Function

Public Function LoadLCMSFeatureFiles(ByVal strLCMSFeaturesFilePath As String, _
                                     ByVal strLCMSFeatureToPeakMapFilePath As String, _
                                     ByVal lngGelIndex As Long) As Long
                                     
    '-------------------------------------------------------------------------
    'Returns 0 if data successfuly loaded, -2 if data set is too large,
    '-3 if problems with scan numbers, -4 if no data found, -5 if user cancels load,
    '-6 for file not found or invalid file
    '-7 for file problem that user was already notified about
    '-10 for any other error
    '-------------------------------------------------------------------------
    
    Dim fso As New FileSystemObject
        
    Dim lngByteCountTotal As Long
    Dim lngTotalBytesRead As Long

    Dim blnLoadFeatureToMapFile As Boolean
    Dim blnSuccess As Boolean
    
    Dim lngReturnValue As Long
        
On Error GoTo LoadLCMSFeatureFilesErrorHandler
   
    ' Using a dictionary object as a hashtable
    Dim htLCMSFeatureIndexMap As Dictionary
    Set htLCMSFeatureIndexMap = New Dictionary
    htLCMSFeatureIndexMap.RemoveAll
       
    mGelIndex = lngGelIndex
    mErrorMessage = ""
    
    ' Initially reserve space for 50000 entries in the mFeatureToPeakMap arrays
    mFeatureToPeakMapCount = 0
    ReDim mFeatureToPeakMapFeatureIdx(49999)
    ReDim mFeatureToPeakMapPeakIdx(49999)
    
    If Not fso.FileExists(strLCMSFeaturesFilePath) Then
        mErrorMessage = "File not found: " & strLCMSFeaturesFilePath
        LoadLCMSFeatureFiles = -6
        Exit Function
    End If
    
    blnLoadFeatureToMapFile = fso.FileExists(strLCMSFeatureToPeakMapFilePath)
    
    If Not blnLoadFeatureToMapFile Then
        ' The LCMSFeature to Peak Map file was not found; will try to auto-map the loaded data to the LCMS Features
        mErrorMessage = "Note: File not found: " & strLCMSFeatureToPeakMapFilePath
        AddToAnalysisHistory mGelIndex, mErrorMessage
    End If
       
    ' Initialize GelUMC(mGelIndex)
    If Not ManageClasses(mGelIndex, UMCManageConstants.UMCMngInitialize) Then
       mErrorMessage = "Error initializing LC-MS Feature memory structures."
       LoadLCMSFeatureFiles = -10
       Exit Function
    End If
       
    ' Initialize GelUMC().Def using UMCDef
    
    GelUMC(mGelIndex).def = UMCDef
    
    ' Make sure some of the settings are up-to-date
    With GelUMC(mGelIndex).def
        .UMCType = glUMC_TYPE_FROM_NET
        .DefScope = glScope.glSc_All
        .MWField = mftMWMono
        .TolType = glMassToleranceConstants.gltPPM
        .Tol = mAutoMapDataPointsMassTolerancePPM
        .UMCSharing = False
        .ClassAbu = UMCClassAbundanceConstants.UMCAbuSum
        .OddEvenProcessingMode = oepUMCOddEvenProcessingMode.oepProcessAll
    End With
    
    ' Initialize the progress
    lngTotalBytesRead = 0
    lngByteCountTotal = FileLen(strLCMSFeaturesFilePath)
    If blnLoadFeatureToMapFile Then
        lngByteCountTotal = lngByteCountTotal + FileLen(strLCMSFeatureToPeakMapFilePath)
    End If
    
    mProgressDescription = "Reading data"
    mProgressPctComplete = 0 / CSng(lngByteCountTotal) * 100
              
    ' Read the LCMSFeature file
    lngReturnValue = ReadLCMSFeaturesFile(fso, strLCMSFeaturesFilePath, htLCMSFeatureIndexMap, lngByteCountTotal, lngTotalBytesRead)
       
    If lngReturnValue = 0 Then
        If blnLoadFeatureToMapFile Then
            ' Read the LCMSFeature to Peak Map file
            lngReturnValue = ReadLCMSFeatureToPeakMapFile(fso, strLCMSFeatureToPeakMapFilePath, htLCMSFeatureIndexMap, lngByteCountTotal, lngTotalBytesRead)
        Else
            blnSuccess = AutoMapDataPointsToFeatures(mAutoMapDataPointsMassTolerancePPM)
            
            lngReturnValue = 0
            If Not blnSuccess Then
                If mErrorMessage = "" Then
                    mErrorMessage = "Error auto-mapping featuers to loaded data points"
                End If
                lngReturnValue = -10
            End If
            
        End If
        
        If lngReturnValue = 0 Then
           
            ' Make sure the mFeatureToPeakMap arrays are sorted by Feature Index, then peak index
            ' In addition, make sure there are no duplicates
            SortFeatureMapMembers mFeatureToPeakMapFeatureIdx, mFeatureToPeakMapPeakIdx, mFeatureToPeakMapCount
            
            ' Now update GelUMC() with the data in mFeatureToPeakMap
            blnSuccess = AddFeatureMembersToUMCs(mFeatureToPeakMapFeatureIdx, mFeatureToPeakMapPeakIdx, mFeatureToPeakMapCount)
        
            ' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
            UpdateUMCStatArrays mGelIndex, False, False
        
        End If
        
    End If
        
    LoadLCMSFeatureFiles = lngReturnValue
    Exit Function

LoadLCMSFeatureFilesErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "LoadLCMSFeatureFiles"
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    LoadLCMSFeatureFiles = lngReturnValue
    
End Function
    
Private Function ReadLCMSFeaturesFile(ByRef fso As FileSystemObject, _
                                      ByVal strLCMSFeaturesFilePath As String, _
                                      ByRef htLCMSFeatureIndexMap As Dictionary, _
                                      ByVal lngByteCountTotal As Long, _
                                      ByRef lngTotalBytesRead As Long) As Long
                                      
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
    
    Dim lngFeatureIndex As Long
    Dim lngScanCenter As Long
    Dim dblMonoisotopicMass As Double
    Dim dblClassAbundance As Double
    
    Dim blnSuccess As Boolean
    
On Error GoTo ReadLCMSFeaturesFileErrorHandler

    ReDim intColumnMapping(LCMSFEATURES_FILE_COLUMN_COUNT - 1)
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To LCMSFEATURES_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    mProgressDescription = "Reading LCMS Features file"
    
    GelUMC(mGelIndex).UMCCnt = 0
    
    Set tsInFile = fso.OpenTextFile(strLCMSFeaturesFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream

        strLineIn = tsInFile.ReadLine
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If GelUMC(mGelIndex).UMCCnt Mod 50 = 0 Then
            mProgressPctComplete = lngTotalBytesRead / CSng(lngByteCountTotal) * 100
        End If
        
        ' Check for valid line (must contain at least one tab and must be
        ' the header line or start with a number)
        strLineIn = Trim(strLineIn)
        If InStr(strLineIn, vbTab) > 0 Then
            If blnColumnsDefined Then
                blnDataLine = IsNumeric(Left(strLineIn, 1))
            Else
                ' Haven't found the column header line yet
                ' If the line starts with a number, then assume the column header line is missing and use the default column order
                ' If the line starts with text, then assume this is the column header line
                
                If IsNumeric(Left(strLineIn, 1)) Then
                    ' Use the default column mappings

                    intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex) = LCMSFeaturesFileColumnConstants.FeatureIndex
                    intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) = LCMSFeaturesFileColumnConstants.Scan
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart) = LCMSFeaturesFileColumnConstants.ScanStart
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd) = LCMSFeaturesFileColumnConstants.ScanEnd
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ScanAligned) = LCMSFeaturesFileColumnConstants.ScanAligned
                    intColumnMapping(LCMSFeaturesFileColumnConstants.NET) = LCMSFeaturesFileColumnConstants.NET
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass) = LCMSFeaturesFileColumnConstants.MonoisotopicMass
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMassCalibrated) = LCMSFeaturesFileColumnConstants.MonoisotopicMassCalibrated
                    intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance) = LCMSFeaturesFileColumnConstants.Abundance
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepMZ) = LCMSFeaturesFileColumnConstants.ClassRepMZ
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepCharge) = LCMSFeaturesFileColumnConstants.ClassRepCharge
                    

                    ' Column headers were not present
                     AddToAnalysisHistory mGelIndex, "LCMSFeatures file " & fso.GetFileName(strLCMSFeaturesFilePath) & " did not contain column headers; using the default headers (" & GetDefaultLCMSFeatureColumnHeaders(False) & ")"

                    blnDataLine = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, vbTab)
                    strUnknownColumnList = ""
                    
                    For lngIndex = 0 To UBound(strData)
                        If lngIndex >= LCMSFEATURES_FILE_COLUMN_COUNT Then
                            ' Too many column headers; ignore the extra columns
                            Debug.Assert False
                            Exit For
                        End If

                        strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                        
                        Select Case strColumnHeader
                        Case LCMSFEATURES_COLUMN_FEATURE_INDEX:         intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN:                  intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_START:            intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_END:              intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_ALIGNED:          intColumnMapping(LCMSFeaturesFileColumnConstants.ScanAligned) = lngIndex
                        Case LCMSFEATURES_COLUMN_NET:                   intColumnMapping(LCMSFeaturesFileColumnConstants.NET) = lngIndex
                        Case LCMSFEATURES_COLUMN_MONO_MASS:             intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass) = lngIndex
                        Case LCMSFEATURES_COLUMN_MONO_MASS_CALIBRATED:  intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMassCalibrated) = lngIndex
                        Case LCMSFEATURES_COLUMN_ABUNDANCE:             intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance) = lngIndex
                        Case LCMSFEATURES_COLUMN_CLASS_REP_MZ:          intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepMZ) = lngIndex
                        Case LCMSFEATURES_COLUMN_CLASS_REP_CHARGE:      intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepCharge) = lngIndex

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
                        AddToAnalysisHistory mGelIndex, COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & fso.GetFileName(strLCMSFeaturesFilePath) & ": " & strUnknownColumnList & "; Known columns are: " & vbCrLf & GetDefaultLCMSFeatureColumnHeaders(False)
                    End If
                    
                    blnDataLine = False
                
                End If
                
                ' Warn the user if any of the important columns are missing
                If intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex) < 0 Or _
                   intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) < 0 Or _
                   intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass) < 0 Then
                   
                    strMessage = COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strLCMSFeaturesFilePath) & "; the expected columns are: " & vbCrLf & GetDefaultLCMSFeatureColumnHeaders(True)
                    If mShowDialogBoxes Then
                        MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
                    End If
                    AddToAnalysisHistory mGelIndex, strMessage
                End If
                
                blnColumnsDefined = True
            End If
        End If
                
        If blnDataLine Then
            
            strData = Split(strLineIn, vbTab)
            
            ' Extract out the Feature index; ignore any features defined more than once
            lngFeatureIndex = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex), -1)
            
            If lngFeatureIndex >= 0 And Not htLCMSFeatureIndexMap.Exists(lngFeatureIndex) Then
                
                lngScanCenter = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.Scan), 0)
                dblMonoisotopicMass = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass), 0)
                dblClassAbundance = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance), 0)
                
                If GelUMC(mGelIndex).UMCCnt >= UBound(GelUMC(mGelIndex).UMCs) Then
                    ' Reserve more room in .UMCs()
                    blnSuccess = ManageClasses(mGelIndex, UMCMngAdd)
                End If
                
                ' Define the new .UMC
                With GelUMC(mGelIndex).UMCs(GelUMC(mGelIndex).UMCCnt)
                    
                    .ClassRepInd = 0
                    .ClassRepType = gldtIS
                    .ClassCount = 0
                    ReDim .ClassMInd(0)         ' members of the UMC(including representative)         (0-based array); index of the member in GelData
                    ReDim .ClassMType(0)        ' type of member: gldtCS or gldtIS
                    .ClassAbundance = dblClassAbundance
                    
                    .ClassMW = dblMonoisotopicMass
                    .ClassMWStD = 0
                    .ClassMassCorrectionDa = 0
                    
                    .ClassScore = 0
                    .ClassNET = 0
                    .ClassStatusBits = 0
                    
                    .MinScan = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart), lngScanCenter)
                    .MaxScan = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd), lngScanCenter)
                    .MinMW = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass), 0)
                    .MaxMW = .MinMW
                                        
                    .ChargeStateStatsRepInd = 0
                    .ChargeStateCount = 0
                    ReDim .ChargeStateBasedStats(0)
                    With .ChargeStateBasedStats(0)
                        .Charge = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepCharge), 0)
                        .Count = 1
                        .Mass = dblMonoisotopicMass
                        .MassStD = 0
                        .Abundance = dblClassAbundance
                        .GroupRepIndex = 0          ' Index of the group representative; pointer to .ClassMInd() and .ClassMType() arrays, NOT a direct pointer to GelData
                    End With
                    
                    .PercentMembersIReportMonoPlus4 = 0
                    .PercentMembersIReportMonoMinus4 = 0
                End With
                
                ' Store the mapping between Index in the LCMSFeatures file and Index in VIPER
                htLCMSFeatureIndexMap.add lngFeatureIndex, GelUMC(mGelIndex).UMCCnt
                
                GelUMC(mGelIndex).UMCCnt = GelUMC(mGelIndex).UMCCnt + 1
            End If
        End If
    
    Loop
    
    tsInFile.Close

    ' Shrink GelUMC().UMCs to length .UMCCnt
    blnSuccess = ManageClasses(mGelIndex, UMCMngTrim)
    
    ReadLCMSFeaturesFile = 0
    Exit Function

ReadLCMSFeaturesFileErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadLCMSFeaturesFile"
    
    On Error Resume Next
    tsInFile.Close
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    ReadLCMSFeaturesFile = lngReturnValue

End Function

Private Function ReadLCMSFeatureToPeakMapFile(ByRef fso As FileSystemObject, _
                                              ByVal strLCMSFeatureToPeakMapFilePath As String, _
                                              ByRef htLCMSFeatureIndexMap As Dictionary, _
                                              ByVal lngByteCountTotal As Long, _
                                              ByRef lngTotalBytesRead As Long) As Long

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
    
    Dim lngFeatureIndex As Long
    Dim lngPeakIndex As Long
    Dim lngUMCIndex As Long
    
    Dim blnSuccess As Boolean
    
On Error GoTo ReadLCMSFeatureToPeakMapFileErrorHandler

    lngReturnValue = 0
    
    ReDim intColumnMapping(PEAK_MAP_FILE_COLUMN_COUNT - 1)
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To PEAK_MAP_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    mProgressDescription = "Reading LCMS Feature to Peak Map file"
    
    ' We will load the mappings into the mFeatureToPeakMap arrays, then update GelUMC
    ' The calling function should have already reserved space in the mFeatureToPeakMap arrays
    mFeatureToPeakMapCount = 0
    
    Set tsInFile = fso.OpenTextFile(strLCMSFeatureToPeakMapFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream

        strLineIn = tsInFile.ReadLine
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If mFeatureToPeakMapCount Mod 500 = 0 Then
            mProgressPctComplete = lngTotalBytesRead / CSng(lngByteCountTotal) * 100
        End If
        
        ' Check for valid line (must contain at least one tab and must be
        ' the header line or start with a number)
        strLineIn = Trim(strLineIn)
        If InStr(strLineIn, vbTab) > 0 Then
            If blnColumnsDefined Then
                blnDataLine = IsNumeric(Left(strLineIn, 1))
            Else
                ' Haven't found the column header line yet
                ' If the line starts with a number, then assume the column header line is missing and use the default column order
                ' If the line starts with text, then assume this is the column header line
                
                If IsNumeric(Left(strLineIn, 1)) Then
                    ' Use the default column mappings

                    intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex) = LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex
                    intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.PeakIndex) = LCMSFeatureToPeakMapFileColumnConstants.PeakIndex
                    
                    ' Column headers were not present
                     AddToAnalysisHistory mGelIndex, "LCMSFeatureToPeakMap file " & fso.GetFileName(strLCMSFeatureToPeakMapFilePath) & " did not contain column headers; using the default headers (" & GetDefaultLCMSFeatureToPeakMapColumnHeaders() & ")"

                    blnDataLine = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, vbTab)
                    strUnknownColumnList = ""
                    
                    For lngIndex = 0 To UBound(strData)
                        If lngIndex >= PEAK_MAP_FILE_COLUMN_COUNT Then
                            ' Too many column headers; ignore the extra columns
                            Debug.Assert False
                            Exit For
                        End If

                        strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                        
                        Select Case strColumnHeader
                        Case PEAK_MAP_COLUMN_FEATURE_INDEX:    intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex) = lngIndex
                        Case PEAK_MAP_COLUMN_PEAK_INDEX:       intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.PeakIndex) = lngIndex
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
                        AddToAnalysisHistory mGelIndex, COLUMN_HEADER_UNKNOWN_WARNING & " found in file " & fso.GetFileName(strLCMSFeatureToPeakMapFilePath) & ": " & strUnknownColumnList & "; Known columns are: " & vbCrLf & GetDefaultLCMSFeatureToPeakMapColumnHeaders()
                    End If
                    
                    blnDataLine = False
                
                End If
                
                ' Warn the user if any of the important columns are missing
                If intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex) < 0 Or _
                   intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.PeakIndex) < 0 Then
                   
                    strMessage = COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strLCMSFeatureToPeakMapFilePath) & "; the expected columns are: " & vbCrLf & GetDefaultLCMSFeatureToPeakMapColumnHeaders()
                    If mShowDialogBoxes Then
                        MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
                    End If
                    AddToAnalysisHistory mGelIndex, strMessage
                End If
                
                blnColumnsDefined = True
            End If
        End If
                
        If blnDataLine Then
            
            strData = Split(strLineIn, vbTab)
            
            ' Extract out the Feature index and Peak Index
            lngFeatureIndex = GetColumnValueLng(strData, intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex), -1)
            lngPeakIndex = GetColumnValueLng(strData, intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.PeakIndex), -1)
            
            If lngFeatureIndex >= 0 And lngPeakIndex >= 0 Then
                
                ' Add this mapping to the mFeatureToPeakMap arrays
                
                If mFeatureToPeakMapCount >= UBound(mFeatureToPeakMapFeatureIdx) Then
                    ' Reserve more room in the mFeatureToPeakMap arrays
                    ReDim Preserve mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount * 2 - 1)
                    ReDim Preserve mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount * 2 - 1)
                End If
                
                mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount) = lngFeatureIndex
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' IMPORTANT NOTE: We might need to add 1 to lngPeakIndex since the data in mFeatureToPeakMapPeakIdx() will be  ''
                '                 used to refer to data in GelData().IsoData(), and .IsoData() is a 1-based array;             ''
                '                 in contrast, the peak indices in the LCMSFeature to Peak Map file might be 0-based           ''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount) = lngPeakIndex
                
                mFeatureToPeakMapCount = mFeatureToPeakMapCount + 1
            End If
        End If
    
    Loop
    
    tsInFile.Close
    
    If mFeatureToPeakMapCount <= 0 Then
        strMessage = "Error: feature to peak map file was empty; will try to auto-determine the mapping between LC-MS Features and loaded data points"
        If mShowDialogBoxes Then
            MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
        End If
        AddToAnalysisHistory mGelIndex, strMessage
        
        ' The following will populate the mFeatureToPeakMap arrays
        AutoMapDataPointsToFeatures (mAutoMapDataPointsMassTolerancePPM)
    Else
        ' Shrink the mFeatureToPeakMap arrays to length mFeatureToPeakMapCount
        ReDim Preserve mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount - 1)
        ReDim Preserve mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount - 1)
    End If
    
    ReadLCMSFeatureToPeakMapFile = 0
    Exit Function

ReadLCMSFeatureToPeakMapFileErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "ReadLCMSFeatureToPeakMapFile"
    
    On Error Resume Next
    tsInFile.Close
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    ReadLCMSFeatureToPeakMapFile = lngReturnValue

End Function


Private Sub SortFeatureMapMembers(ByRef lngFeatureToPeakMapFeatureIdx() As Long, _
                                  ByRef lngFeatureToPeakMapPeakIdx() As Long, _
                                  ByVal lngFeatureToPeakMapCount As Long)

    Dim lngIndex As Long
    Dim lngIndexStart As Long
    Dim lngIndexEnd As Long
    Dim lngFeatureIndexCurrent As Long
    
    Dim lngPeakIndexCount As Long
    Dim lngPeakIndicesForFeature() As Long
    
    Dim lngNewFeatureToPeakMapCount As Long
    Dim lngNewPeakIndexCount As Long
    
    ' Makes sure the peak indices associated with each Feature Index are sorted (and have no duplicates)
    
On Error GoTo SortFeatureMapMembersErrorHandler

    If mFeatureToPeakMapCount <= 0 Then
        ' Nothing to do; this is unexpected
        Debug.Assert False
        Exit Sub
    End If
    
    ' First sort by feature index
    ShellSortLongWithParallelLong mFeatureToPeakMapFeatureIdx, mFeatureToPeakMapPeakIdx, 0, mFeatureToPeakMapCount - 1
    
    ' Reserve space in lngPeakIndicesForFeature
    lngPeakIndexCount = 0
    ReDim lngPeakIndicesForFeature(999)
    
    ' Sort the peak indices mapped to each feature
    ' Do this by stepping through mFeatureToPeakMapFeatureIdx and copying the
    '  entries for a given feature from mFeatureToPeakMapPeakIdx() to lngPeakIndicesForFeature()
    ' Next, sort lngPeakIndicesForFeature() and remove duplicates
    ' Finally, copy back over to mFeatureToPeakMapFeatureIdx and mFeatureToPeakMapPeakIdx
    
    lngIndexStart = 0
    lngNewFeatureToPeakMapCount = 0
    Do While lngIndexStart < mFeatureToPeakMapCount
    
        lngFeatureIndexCurrent = mFeatureToPeakMapFeatureIdx(lngIndexStart)
        lngIndexEnd = lngIndexStart
                
        Do While lngIndexEnd + 1 < mFeatureToPeakMapCount
            If mFeatureToPeakMapFeatureIdx(lngIndexEnd + 1) = lngFeatureIndexCurrent Then
                lngIndexEnd = lngIndexEnd + 1
            Else
                Exit Do
            End If
        Loop
    
        ' Make sure lngPeakIndicesForFeature has enough room for this feature's peaks
        If lngIndexEnd - lngIndexStart + 1 > UBound(lngPeakIndicesForFeature) Then
            ReDim lngPeakIndicesForFeature(lngIndexEnd - lngIndexStart)
        End If
        
        lngPeakIndexCount = 0
        For lngIndex = lngIndexStart To lngIndexEnd
            lngPeakIndicesForFeature(lngPeakIndexCount) = mFeatureToPeakMapPeakIdx(lngIndex)
            lngPeakIndexCount = lngPeakIndexCount + 1
        Next lngIndex
        
        ' Now sort mFeatureToPeakMapPeakIdx
        ShellSortLong lngPeakIndicesForFeature, 0, lngPeakIndexCount - 1
        
        ' Remove any duplicates in mFeatureToPeakMapPeakIdx by copying in place
        lngNewPeakIndexCount = 1
        For lngIndex = 1 To lngPeakIndexCount - 1
            If lngPeakIndicesForFeature(lngIndex) <> lngPeakIndicesForFeature(lngIndex - 1) Then
                ' Peak index is new; keep it
                ' Only need to copy if lngNewPeakIndexCount is not equal to lngIndex, and this will only happen if we found (and removed) a duplicate
                If lngNewPeakIndexCount <> lngIndex Then
                    lngPeakIndicesForFeature(lngNewPeakIndexCount) = lngPeakIndicesForFeature(lngIndex)
                End If
                lngNewPeakIndexCount = lngNewPeakIndexCount + 1
            End If
        Next lngIndex
        lngPeakIndexCount = lngNewPeakIndexCount
        
        ' Now copy the data back into mFeatureToPeakMapFeatureIdx and mFeatureToPeakMapPeakIdx
        For lngIndex = 0 To lngPeakIndexCount - 1
            mFeatureToPeakMapFeatureIdx(lngNewFeatureToPeakMapCount) = lngFeatureIndexCurrent
            mFeatureToPeakMapPeakIdx(lngNewFeatureToPeakMapCount) = lngPeakIndicesForFeature(lngIndex)
            lngNewFeatureToPeakMapCount = lngNewFeatureToPeakMapCount + 1
        Next lngIndex
        
        lngIndexStart = lngIndexEnd + 1
    Loop
    
    Exit Sub

SortFeatureMapMembersErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "SortFeatureMapMembers"
        
End Sub

Private Function StripQuotes(ByVal strText As String) As String

    If Len(strText) > 2 Then
        If Left(strText, 1) = """" And Right(strText, 1) = """" Then
            strText = Mid(strText, 2, Len(strText) - 2)
        End If
    End If
             
    StripQuotes = strText
End Function


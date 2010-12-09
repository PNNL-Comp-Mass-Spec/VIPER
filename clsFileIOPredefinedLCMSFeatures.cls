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
' Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------

' This class reads pre-defined LC/MS features from a text file, populating the data structures for the given Gel
' It requires that the Decon2LS _isos.csv and _scans.csv files have already been loaded into memory

' It will read two files:
'  _Features.txt file (see below for valid column headers)
'  _FeatureToPeakMap.txt file (see below for valid column headers)

Private Const COLUMN_HEADER_UNKNOWN_WARNING As String = "Warning: unknown column headers"
Private Const COLUMN_HEADER_MISSING_WARNING As String = "Warning: expected important column headers"


Private Const LCMSFEATURES_COLUMN_FEATURE_INDEX As String = "feature_index"
Private Const LCMSFEATURES_COLUMN_ORIGINAL_INDEX As String = "original_index"
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
Private Const LCMSFEATURES_COLUMN_CHARGE_MAX As String = "charge_max"
Private Const LCMSFEATURES_COLUMN_DRIFT_TIME As String = "drift_time"
Private Const LCMSFEATURES_COLUMN_CONFORMATION_FIT_SCORE As String = "conformation_fit_score"
Private Const LCMSFEATURES_COLUMN_LC_FIT_SCORE As String = "lc_fit_score"
Private Const LCMSFEATURES_COLUMN_AVERAGE_ISOTOPIC_FIT As String = "average_isotopic_fit"
Private Const LCMSFEATURES_COLUMN_MEMBERS_PERCENTAGE As String = "members_percentage"
Private Const LCMSFEATURES_COLUMN_COMBINED_SCORE As String = "combined_score"

Private Const LCMSFEATURES_COLUMN_AVERAGE_MONO_MASS As String = "average_mono_mass"
Private Const LCMSFEATURES_COLUMN_UMC_MW_MIN As String = "umc_mw_min"
Private Const LCMSFEATURES_COLUMN_UMC_MW_MAX As String = "umc_mw_max"
Private Const LCMSFEATURES_COLUMN_UMC_MEMBER_COUNT As String = "umc_member_count"
Private Const LCMSFEATURES_COLUMN_ABUNDANCE_MAX As String = "max_abundance"

Private Const PEAK_MAP_COLUMN_FEATURE_INDEX As String = "feature_index"
Private Const PEAK_MAP_COLUMN_PEAK_INDEX As String = "peak_index"
Private Const PEAK_MAP_COLUMN_FILTERED_PEAK_INDEX As String = "filtered_peak_index"

Private Const LCMSFEATURES_FILE_COLUMN_COUNT As Integer = 20
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
    MWMin = 11
    MWMax = 12
    MemberCount = 13
    DriftTime = 14
    ConformationFitScore = 15
    LCFitScore = 16
    AverageIsotopicFit = 17
    MembersPercentage = 18
    CombinedScore = 19
End Enum

Private Const PEAK_MAP_FILE_COLUMN_COUNT As Integer = 3
Private Enum LCMSFeatureToPeakMapFileColumnConstants
    FeatureIndex = 0
    PeakIndex = 1
    FilteredPeakIndex = 2
End Enum

' Class-wide variables
Private mGelIndex As Integer
Private mAutoMapDataPointsMassTolerancePPM As Double

' The following are parallel arrays
' They are used to determine the indices of data in GelData().IsoData() that map to each feature
Private mFeatureToPeakMapCount As Long
Private mFeatureToPeakMapFeatureIdx() As Long
Private mFeatureToPeakMapPeakIdx() As Long          ' These indices are the index values in GelData().IsoData() -- note that .IsoData() is a 1-based array

Private mCachedFeatureToPeakMapDefined As Boolean

Private mErrorMessage As String
Private mShowDialogBoxes As Boolean

Private mReadingFilteredIsosFile As Boolean         ' When True, then will use column Filtered_Peak_Index instead of column Peak_Index in the _LCMSFeatureToPeakMap.txt file

Private mPointsLoadMode As plmPointsLoadModeConstants
Private mHashMapOfPointsKept As clsParallelLngArrays

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

Private mProgressForm As frmProgress

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

Public Property Get PointsLoadMode() As plmPointsLoadModeConstants
    PointsLoadMode = mPointsLoadMode
End Property
Public Property Let PointsLoadMode(ByVal plmValue As plmPointsLoadModeConstants)
    mPointsLoadMode = plmValue
End Property

Public Property Get HashMapOfPointsKept() As clsParallelLngArrays
    Set HashMapOfPointsKept = mHashMapOfPointsKept
End Property
Public Property Let HashMapOfPointsKept(ByRef objHashMap As clsParallelLngArrays)
    Set mHashMapOfPointsKept = objHashMap
End Property
        
Public Property Let ProgressForm(ByRef objForm As frmProgress)
    Set mProgressForm = objForm
End Property

Public Property Let ReadingFilteredIsosFile(ByVal Value As Boolean)
    mReadingFilteredIsosFile = Value
End Property
Public Property Get ReadingFilteredIsosFile() As Boolean
    ReadingFilteredIsosFile = mReadingFilteredIsosFile
End Property
Private Sub Class_Initialize()
    mGelIndex = 0
    mAutoMapDataPointsMassTolerancePPM = 5
End Sub

' Functions

Private Function AddFeatureMembersToUMCs(ByRef objLCMSFeatureIndexMap As clsParallelLngArrays) As Boolean
    
    ' Be sure to call SortFeatureMapMembers before calling this function because this function
    '  assumes that the data in mFeatureToPeakMapFeatureIdx and mFeatureToPeakMapPeakIdx
    '  is properly sorted and is unique

    Dim lngIndex As Long
    Dim lngIndexStart As Long
    Dim lngIndexEnd As Long
    Dim lngPeakIndexCount As Long
    
    Dim lngFeatureIndexOriginal As Long
    Dim lngFeatureIndexInVIPER As Long
    
    Dim lngDataPointsMapped As Long
    Dim lngFeaturesProcessed As Long
    
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
    Dim dblMassDiff As Double
    
    Dim blnMatchFound As Boolean
    Dim blnSuccess As Boolean
    
    Dim lngFeatureCountNoMappedPoints As Long
    Dim lngMassDiffWarningCount As Long
    
On Error GoTo AddFeatureMembersToUMCsErrorHandler

    ' Set this to true for now
    blnSuccess = True
    
    lngDataPointsMapped = 0
    lngFeaturesProcessed = 0
    
    lngIndexStart = 0
    Do While lngIndexStart < mFeatureToPeakMapCount
    
        lngFeatureIndexOriginal = mFeatureToPeakMapFeatureIdx(lngIndexStart)
        lngIndexEnd = lngIndexStart
                
        Do While lngIndexEnd + 1 < mFeatureToPeakMapCount
            If mFeatureToPeakMapFeatureIdx(lngIndexEnd + 1) = lngFeatureIndexOriginal Then
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
            
            ' Dereference lngFeatureIndexOriginal to lngFeatureIndexInVIPER using objLCMSFeatureIndexMap
            lngFeatureIndexInVIPER = objLCMSFeatureIndexMap.GetItemForKey(lngFeatureIndexOriginal, blnMatchFound)
            
            ''' Temp Debug
            ''Debug.Assert lngFeatureIndexInVIPER < 143 Or lngFeatureIndexInVIPER > 145
            
            If Not blnMatchFound Then
                ' Could not dereference feature
                ' This will happen if we filtered the features by abundance or conformer fit score
                lngClassRepIndexPointer = 0
            Else
            
                ' Determine the class rep index (i.e. the feature with the highest intensity)
                
                lngClassRepIndexPointer = lngIndexStart
                If Not GetIsoDataInfo(mFeatureToPeakMapPeakIdx(lngClassRepIndexPointer), lngScan, dblMass, dblAbundance) Then
                    ' mFeatureToPeakMapPeakIdx() did not point to a valid index in GelData().IsoData()
                    ' This is possible if data was filtered when loading the _isos.csv file
                    Debug.Assert False
                End If
                              
                dblBestAbundance = dblAbundance
                dblClassAbundance = dblAbundance
                
                lngMinScan = lngScan
                lngMaxScan = lngScan
                
                dblMinMass = dblMass
                dblMaxMass = dblMass
                    
                ' If more than one mapped point for this feature, then step through the mapped points
                '  to update dblClassAbundance, lngMinScan, lngMaxScan, dblMinMass, and dblMaxMass
                For lngIndex = lngIndexStart + 1 To lngIndexEnd
                    
                    If Not GetIsoDataInfo(mFeatureToPeakMapPeakIdx(lngIndex), lngScan, dblMass, dblAbundance) Then
                        ' mFeatureToPeakMapPeakIdx() did not point to a valid index in GelData().IsoData()
                        ' This is possible if data was filtered when loading the _isos.csv file
                        Debug.Assert False
                    Else
                        If dblAbundance > dblBestAbundance Then
                            dblBestAbundance = dblAbundance
                            lngClassRepIndexPointer = lngIndex
                        End If
                        
                        dblClassAbundance = dblClassAbundance + dblBestAbundance
                        
                        If lngIndex = lngIndexStart + 1 Then
                            If lngMinScan = 0 And dblMinMass = 0 Then
                                ' The data point at index lngIndexStart didn't resolve to any valid, loaded data
                                ' Update the min/max scan variables
                                lngMinScan = lngScan
                                lngMaxScan = lngScan
                                
                                dblMinMass = dblMass
                                dblMaxMass = dblMass
                            End If
                        End If
                        
                        If lngScan < lngMinScan And lngScan > 0 Then lngMinScan = lngScan
                        If lngScan > lngMaxScan Then lngMaxScan = lngScan
                        
                        If dblMass < dblMinMass And dblMass > 0 Then dblMinMass = dblMass
                        If dblMass > dblMaxMass Then dblMaxMass = dblMass
                    
                    End If
                    
                Next lngIndex
                    
                ' Now associate the data points with the LC-MS Feature
                With GelUMC(mGelIndex).UMCs(lngFeatureIndexInVIPER)
                    
                    .ClassRepInd = mFeatureToPeakMapPeakIdx(lngClassRepIndexPointer)
                    .ClassRepType = gldtIS
                    .ClassCount = lngPeakIndexCount               ' This is the expected count; the actual count could be lower due to filtering when loading the _isos.csv file; we'll update lngPeakIndexCount in the following for loop with the actual count
                    ReDim .ClassMInd(lngPeakIndexCount - 1)       ' members of the UMC(including representative)         (0-based array); index of the member in GelData
                    ReDim .ClassMType(lngPeakIndexCount - 1)      ' type of member: gldtCS or gldtIS
                    
                    lngPeakIndexCount = 0
                    For lngIndex = lngIndexStart To lngIndexEnd
                        If mFeatureToPeakMapPeakIdx(lngIndex) > 0 Then
                            
                            ' Unless something really wacky is going on, the mass difference between the feature and
                            '  each member data point should be less than 5 Da
                            dblMassDiff = .ClassMW - GelData(mGelIndex).IsoData(mFeatureToPeakMapPeakIdx(lngIndex)).MonoisotopicMW
                            
                            If Abs(dblMassDiff) >= 5 Then
                                lngMassDiffWarningCount = lngMassDiffWarningCount + 1
                                If lngMassDiffWarningCount < 5 Then
                                    Debug.Assert False
                                End If
                            End If
                            
                            .ClassMInd(lngPeakIndexCount) = mFeatureToPeakMapPeakIdx(lngIndex)
                            .ClassMType(lngPeakIndexCount) = gldtIS
                            lngPeakIndexCount = lngPeakIndexCount + 1
                        End If
                    Next lngIndex
                    lngDataPointsMapped = lngDataPointsMapped + lngPeakIndexCount
                    
                    ' Update .ClassCount if lngPeakIndexCount is less than .ClassCount
                    ' Note that .ClassCountPredefinedLCMSFeatures tracks the original value in .ClassCount
                    If lngPeakIndexCount < .ClassCount Then
                        .ClassCount = lngPeakIndexCount
                        
                        If .ClassCount = 0 Then
                            ReDim Preserve .ClassMInd(0)
                            ReDim Preserve .ClassMType(0)
                            .ClassMInd(0) = 0
                            .ClassMType(0) = 0
                        Else
                            ReDim Preserve .ClassMInd(.ClassCount - 1)
                            ReDim Preserve .ClassMType(.ClassCount - 1)
                        End If
                        
                    End If
                    
                    ' Only update .ClassAbundance if it is 0
                    ' The reason is that the .ClassAbundance value loaded from disk will be much more accurate (and much higher) if we only loaded one point per LC-MS Feature
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
                    
                    If lngMinScan < .MinScan And lngMinScan > 0 Then .MinScan = lngMinScan
                    If lngMaxScan > .MaxScan Then .MaxScan = lngMaxScan
                    
                    If dblMinMass < .MinMW And dblMinMass > 0 Then .MinMW = dblMinMass
                    If dblMaxMass > .MaxMW Then .MaxMW = dblMaxMass
                    
                    ' Leave the .ChargeStateStats values unchanged
                    '.ChargeStateStatsRepInd
                    '.ChargeStateCount
                    
                    ' Leave the .PercentMembersIReport values unchanged
                    '.PercentMembersIReportMonoPlus4
                    '.PercentMembersIReportMonoMinus4
                    
                End With
                
            End If
            
            lngFeaturesProcessed = lngFeaturesProcessed + 1

        End If
        
        lngIndexStart = lngIndexEnd + 1
    Loop
    
    ' Now step through the UMCs and look for any with UBound(.ClassMType) = 0 and .ClassMType(0) = 0
    ' These are features that were loaded into memory, but have no mapped points
    ' Need to change .ClassCount to 0 for these features
    lngFeatureCountNoMappedPoints = 0
    For lngIndex = 0 To GelUMC(mGelIndex).UMCCnt - 1
        If UBound(GelUMC(mGelIndex).UMCs(lngIndex).ClassMType) = 0 Then
            If GelUMC(mGelIndex).UMCs(lngIndex).ClassMType(0) = 0 Then
                ' This feature doesn't have any mapped points
                ' This code won't normally be reached
                
                lngFeatureCountNoMappedPoints = lngFeatureCountNoMappedPoints + 1
                If lngFeatureCountNoMappedPoints < 5 Then
                    ' Only stop the IDE the first 5 times this happens
                    Debug.Assert False
                End If
                
                'Debug.Print lngIndex
                'If lngFeatureCountNoMappedPoints Mod 195 = 0 Then
                '    Debug.Assert False
                'End If
                
                GelUMC(mGelIndex).UMCs(lngIndex).ClassCount = 0
            End If
            
            Debug.Assert GelUMC(mGelIndex).UMCs(lngIndex).ClassCount <= 1
        End If
    Next lngIndex
    
    AddToAnalysisHistory mGelIndex, "Mapped " & Format(lngDataPointsMapped, "0,000") & " data points to " & Format(lngFeaturesProcessed, "0,000") & " LC-MS Features"

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
                                    ReDim Preserve mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount * 2)
                                    ReDim Preserve mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount * 2)
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

Public Function CacheFeatureToPeakMapIndices(ByVal lngGelIndex As Long, _
                                             ByVal strLCMSFeatureToPeakMapFilePath As String, _
                                             ByRef objPointsToKeep As clsParallelLngArrays, _
                                             ByRef objFeatureToScanMap As clsParallelLngArrays) As Long

    '-------------------------------------------------------------------------
    ' Reads the LCMSFeatureToPeakMap file and populates a hash table
    ' with a mapping from peak index to feature index
    '
    ' Only keeps points with an entry in objFeatureToScanMap
    '
    'Returns 0 if data successfuly loaded,
    '-6 for file not found or invalid file,
    '-10 for any other error
    '-------------------------------------------------------------------------
                                        
    Dim fso As New FileSystemObject
    
    Dim lngByteCountTotal As Long
    Dim lngTotalBytesRead As Long
    
    Dim tsInFile As TextStream
    Dim strLineIn As String
    Dim strMessage As String
    
    Dim lngReturnValue As Long
    Dim lngIndex As Long
    Dim lngFeatureIndex As Long
    Dim lngPeakIndex As Long
    
    Dim lngDataLineCount As Long
    
    Dim lngFeatureIndexCurrent As Long
    Dim blnFeaturePassesFilters As Boolean
    
    Dim blnColumnsDefined As Boolean
    Dim blnDataLine As Boolean
    Dim blnMissingColumns As Boolean
    
    Dim strData() As String
    Dim strColumnHeader As String
    
    Dim intColumnMapping() As Integer
    
On Error GoTo CacheFeatureToPeakMapIndicesErrorHandler

    If objPointsToKeep Is Nothing Then
        Set objPointsToKeep = New clsParallelLngArrays
    Else
        objPointsToKeep.Clear
    End If
 
    If Not fso.FileExists(strLCMSFeatureToPeakMapFilePath) Then
        mErrorMessage = "File not found: " & strLCMSFeatureToPeakMapFilePath
        CacheFeatureToPeakMapIndices = -6
        Exit Function
    End If
        
    ' Make sure objFeatureToScanMap is sorted
    objFeatureToScanMap.SortNow
       
    ReDim intColumnMapping(PEAK_MAP_FILE_COLUMN_COUNT - 1)

    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To PEAK_MAP_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
            
    ' Initialize the progress
    lngTotalBytesRead = 0
    lngByteCountTotal = FileLen(strLCMSFeatureToPeakMapFilePath)
    
    mProgressDescription = "Caching LCMS feature to peak mapping"
    If Not mProgressForm Is Nothing Then
        mProgressForm.InitializeSubtask mProgressDescription, 0, 105
    End If
    
    lngDataLineCount = 0
    lngFeatureIndexCurrent = -1
    blnFeaturePassesFilters = True
    
    Set tsInFile = fso.OpenTextFile(strLCMSFeatureToPeakMapFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream
        strLineIn = tsInFile.ReadLine
        
        ' Check for valid line (must contain at least one tab and must be
        ' the header line or start with a number)
        strLineIn = Trim(strLineIn)
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If lngDataLineCount Mod 500 = 0 Then
            mProgressPctComplete = lngTotalBytesRead / CSng(lngByteCountTotal) * 100
            If Not mProgressForm Is Nothing Then
                mProgressForm.UpdateSubtaskProgressBar mProgressPctComplete, False
                If KeyPressAbortProcess > 1 Then Exit Do
            End If
        End If
                
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
                
                    blnDataLine = True
                    blnColumnsDefined = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, vbTab)
                
                    For lngIndex = 0 To UBound(strData)
                        
                        strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                    
                        Select Case strColumnHeader
                        Case PEAK_MAP_COLUMN_FEATURE_INDEX:    intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex) = lngIndex
                        Case PEAK_MAP_COLUMN_PEAK_INDEX:       intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.PeakIndex) = lngIndex
                        Case PEAK_MAP_COLUMN_FILTERED_PEAK_INDEX: intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FilteredPeakIndex) = lngIndex
                        Case Else
                            ' Ignore this column; either it's unknown or we don't need it in this function
                        End Select
                    
                    Next lngIndex
                
                    blnDataLine = False
            
                End If
            
                ' Warn the user if any of the important columns are missing
                
                blnMissingColumns = False
                If mReadingFilteredIsosFile Then
                    If intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex) < 0 Or _
                       intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FilteredPeakIndex) < 0 Then
                       blnMissingColumns = True
                    End If
                Else
                    If intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex) < 0 Or _
                       intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.PeakIndex) < 0 Then
                       blnMissingColumns = True
                    End If
                End If
                
                If blnMissingColumns Then
                    strMessage = COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strLCMSFeatureToPeakMapFilePath) & "; the expected columns are: " & PEAK_MAP_COLUMN_FEATURE_INDEX & " and " & PEAK_MAP_COLUMN_PEAK_INDEX & " or " & PEAK_MAP_COLUMN_FILTERED_PEAK_INDEX
                    If mShowDialogBoxes Then
                        MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
                    End If
                    AddToAnalysisHistory lngGelIndex, strMessage
                End If
                
                blnColumnsDefined = True
            End If
        End If
            
        If blnDataLine Then
        
            strData = Split(strLineIn, vbTab)
        
            ' Extract out the peak index and feature index that it is mapped to
            
            If mReadingFilteredIsosFile Then
                ' The filtered peak index column will often be blank
                ' Check for this using Len(strData(2))
                lngPeakIndex = GetColumnValueLng(strData, intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FilteredPeakIndex), -1)
            Else
                lngPeakIndex = GetColumnValueLng(strData, intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.PeakIndex), -1)
            End If
            
            If lngPeakIndex >= 0 Then
                lngFeatureIndex = GetColumnValueLng(strData, intColumnMapping(LCMSFeatureToPeakMapFileColumnConstants.FeatureIndex), -1)
            End If
            
            If lngFeatureIndex >= 0 And lngPeakIndex >= 0 Then
                ' Add 1 to keep indexing in sync with VIPER data structures
                lngPeakIndex = lngPeakIndex + 1
    
                If lngFeatureIndex <> lngFeatureIndexCurrent Then
                    ' Update lngFeatureIndexCurrent and see if it is present in objFeatureToScanMap
                    lngFeatureIndexCurrent = lngFeatureIndex
                    If objFeatureToScanMap.Exists(lngFeatureIndex) Then
                        blnFeaturePassesFilters = True
                    Else
                        blnFeaturePassesFilters = False
                    End If
                End If
                
                If blnFeaturePassesFilters Then
                    objPointsToKeep.add lngPeakIndex, lngFeatureIndex
                End If
            End If
        
            lngDataLineCount = lngDataLineCount + 1
        End If

    Loop
    
    tsInFile.Close
           
    If objPointsToKeep.Count <= 0 Then
        strMessage = "Warning: feature to peak map file was empty"
        AddToAnalysisHistory lngGelIndex, strMessage
    Else
        AddToAnalysisHistory lngGelIndex, "Loaded feature to peak map information for " & Format(objPointsToKeep.Count, "0,000") & " data points"
    End If
    
    ' Note: Do not sort the data in objPointsToKeep
    ' We want to leave it sorted by the order that was in the _LCMSFeatureToPeakMap file for now, since
    ' we plan to copy the data and place in objPointsToKeepSortedByPeak later
    
    CacheFeatureToPeakMapIndices = 0
    Exit Function

CacheFeatureToPeakMapIndicesErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "CacheFeatureToPeakMapIndices"
    
    On Error Resume Next
    tsInFile.Close
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    CacheFeatureToPeakMapIndices = lngReturnValue
End Function

Public Function CachePredefinedLCMSFeatures(ByVal lngGelIndex As Long, _
                                            ByVal strLCMSFeaturesFilePath As String, _
                                            ByRef objFeatureToScanMap As clsParallelLngArrays, _
                                            ByVal dblLCMSFeatureAbuMin As Double, _
                                            ByVal lngLCMSFeatureScanCountMin As Long, _
                                            ByVal dblIMSConformerScoreMin As Double) As Long
                                            
    '-------------------------------------------------------------------------
    ' Reads the LCMSFeature file and populates a hash table
    ' with a mapping from feature index to scan number of the center of the feature
    '
    ' If blnFilterFeatures = True, then will filter out features with:
    '   abundance < dblLCMSAbundanceMin or conformation score < dblIMSConformerScoreMin
    '
    'Returns 0 if data successfuly loaded,
    '-6 for file not found or invalid file,
    '-10 for any other error
    '-------------------------------------------------------------------------
                                        
    Dim fso As New FileSystemObject
    
    Dim lngByteCountTotal As Long
    Dim lngTotalBytesRead As Long
    
    Dim tsInFile As TextStream
    Dim strLineIn As String
    Dim strMessage As String
    
    Dim lngReturnValue As Long
    Dim lngIndex As Long
    Dim lngFeatureIndex As Long
    
    Dim lngScanStart As Long
    Dim lngScanEnd As Long
    Dim lngScanCenter As Long
    Dim lngLCMSFeatureScanCount As Long
    
    Dim blnColumnsDefined As Boolean
    Dim blnDataLine As Boolean
    
    Dim blnFilterFeatures As Boolean
    Dim blnValidFeature As Boolean
    
    Dim strData() As String
    Dim strColumnHeader As String

    Dim intColumnMapping() As Integer
    
    Dim dblAbundance As Double
    
    Dim dblIMSConformerScore As Double
    Dim dblLCFitScore As Double
    Dim dblAvgIsotopicFit As Double
    Dim dblMembersPercentage As Double
    Dim dblCombinedScore As Double
    
    Dim dblConformerScoreToFilterOn As Double

            
On Error GoTo CacheScanCenterForLCMSFeaturesErrorHandler
    
    Set objFeatureToScanMap = New clsParallelLngArrays
    
    If Not fso.FileExists(strLCMSFeaturesFilePath) Then
        mErrorMessage = "File not found: " & strLCMSFeaturesFilePath
        CachePredefinedLCMSFeatures = -6
        Exit Function
    End If
    
    If dblLCMSFeatureAbuMin > 0 Or lngLCMSFeatureScanCountMin > 1 Or dblIMSConformerScoreMin > 0 Then
        blnFilterFeatures = True
    Else
        blnFilterFeatures = False
    End If
    
    ReDim intColumnMapping(LCMSFEATURES_FILE_COLUMN_COUNT - 1)
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To LCMSFEATURES_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    ' Initialize the progress
    lngTotalBytesRead = 0
    lngByteCountTotal = FileLen(strLCMSFeaturesFilePath)
    
    mProgressDescription = "Caching and filtering LC-MS features"
    If Not mProgressForm Is Nothing Then
        mProgressForm.InitializeSubtask mProgressDescription, 0, 100
    End If
    
    Set tsInFile = fso.OpenTextFile(strLCMSFeaturesFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream
        strLineIn = tsInFile.ReadLine
        
        ' Check for valid line (must contain at least one tab and must be
        ' the header line or start with a number)
        strLineIn = Trim(strLineIn)
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If objFeatureToScanMap.Count Mod 50 = 0 Then
            mProgressPctComplete = lngTotalBytesRead / CSng(lngByteCountTotal) * 100
            If Not mProgressForm Is Nothing Then
                mProgressForm.UpdateSubtaskProgressBar mProgressPctComplete, False
            End If
            If KeyPressAbortProcess > 1 Then Exit Do
        End If
        
        
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
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart) = LCMSFeaturesFileColumnConstants.ScanStart
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd) = LCMSFeaturesFileColumnConstants.ScanEnd
                    intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) = LCMSFeaturesFileColumnConstants.Scan
                    intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance) = LCMSFeaturesFileColumnConstants.Abundance
                    
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ConformationFitScore) = LCMSFeaturesFileColumnConstants.ConformationFitScore
                    intColumnMapping(LCMSFeaturesFileColumnConstants.LCFitScore) = LCMSFeaturesFileColumnConstants.LCFitScore
                    intColumnMapping(LCMSFeaturesFileColumnConstants.AverageIsotopicFit) = LCMSFeaturesFileColumnConstants.AverageIsotopicFit
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MembersPercentage) = LCMSFeaturesFileColumnConstants.MembersPercentage
                    intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore) = LCMSFeaturesFileColumnConstants.CombinedScore
    
                    blnDataLine = True
                    blnColumnsDefined = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, vbTab)
                    
                    For lngIndex = 0 To UBound(strData)
                       
                        strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                        
                        Select Case strColumnHeader
                        Case LCMSFEATURES_COLUMN_FEATURE_INDEX:         intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_START:            intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_END:              intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN:                  intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) = lngIndex
                        Case LCMSFEATURES_COLUMN_ABUNDANCE:             intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance) = lngIndex
                        Case LCMSFEATURES_COLUMN_CONFORMATION_FIT_SCORE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.ConformationFitScore) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_LC_FIT_SCORE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.LCFitScore) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_AVERAGE_ISOTOPIC_FIT
                            intColumnMapping(LCMSFeaturesFileColumnConstants.AverageIsotopicFit) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_MEMBERS_PERCENTAGE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.MembersPercentage) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_COMBINED_SCORE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore) = lngIndex
                            
                        Case Else
                            ' Ignore this column; either it's unknown or we don't need it in this function
                        End Select
                                                
                    Next lngIndex
                    
                    blnDataLine = False
                
                End If

                ' Warn the user if any of the important columns are missing
                If intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex) < 0 Or _
                   intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) < 0 Then
                   
                    strMessage = COLUMN_HEADER_MISSING_WARNING & " not found in file " & fso.GetFileName(strLCMSFeaturesFilePath) & "; the expected columns are: " & LCMSFEATURES_COLUMN_FEATURE_INDEX & " and " & LCMSFEATURES_COLUMN_SCAN
                    If mShowDialogBoxes Then
                        MsgBox strMessage, vbExclamation + vbOKOnly, glFGTU
                    End If
                    AddToAnalysisHistory lngGelIndex, strMessage
                End If
                
                blnColumnsDefined = True
            End If
        End If
        
        If blnDataLine Then
        
            strData = Split(strLineIn, vbTab)
            
            ' Extract out the Feature index and central scan
            lngFeatureIndex = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex), -1)
            
            lngScanStart = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart), 0)
            lngScanEnd = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd), 0)
            lngScanCenter = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.Scan), 0)
            
            dblAbundance = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance), 0)
            
            dblIMSConformerScore = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ConformationFitScore), 0)
            dblLCFitScore = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.LCFitScore), 0)
            dblAvgIsotopicFit = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.AverageIsotopicFit), 0)
            dblMembersPercentage = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MembersPercentage), 0)
            dblCombinedScore = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore), 0)
            
            If intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore) >= 0 Then
                ' Filter using Combined_Score
                dblConformerScoreToFilterOn = dblCombinedScore
            Else
                ' Filter using Conformer Score
                dblConformerScoreToFilterOn = dblIMSConformerScore
            End If
            
            If lngFeatureIndex >= 0 Then
                
                blnValidFeature = True
                If blnFilterFeatures Then
                
                    lngLCMSFeatureScanCount = lngScanEnd - lngScanStart + 1
                    
                    If dblLCMSFeatureAbuMin > 0 And dblAbundance < dblLCMSFeatureAbuMin Then
                        blnValidFeature = False
                    ElseIf lngLCMSFeatureScanCountMin > 1 And lngLCMSFeatureScanCount < lngLCMSFeatureScanCountMin Then
                        blnValidFeature = False
                    ElseIf dblIMSConformerScoreMin > 0 And dblConformerScoreToFilterOn < dblIMSConformerScoreMin Then
                        blnValidFeature = False
                    End If
                End If
                
                If blnValidFeature Then
                    objFeatureToScanMap.add lngFeatureIndex, lngScanCenter
                End If
                
            End If
        
        End If

    Loop
    
    tsInFile.Close
    
    ' Sort the data in objFeatureToScanMap
    objFeatureToScanMap.SortNow
    
    CachePredefinedLCMSFeatures = 0
    Exit Function

CacheScanCenterForLCMSFeaturesErrorHandler:
    Debug.Assert False
    lngReturnValue = Err.Number
    LogErrors Err.Number, "CachePredefinedLCMSFeatures"
    
    On Error Resume Next
    tsInFile.Close
    
    If lngReturnValue = 0 Then lngReturnValue = -10
    CachePredefinedLCMSFeatures = lngReturnValue
End Function

Public Sub ClearFeatureToPeakMapping()
    mCachedFeatureToPeakMapDefined = False
End Sub

' Copies data from objPeakToFeatureMapping to the local arrays
' During the copy, will update the PeakIndex value using mHashMapOfPointsKept, which has a mapping from the peak index in the data file to the peak index in VIPER's memory space
' In objPeakToFeatureMapping:
'  Key is PeakIndex
'  Value is FeatureIndex
' Ideally, the data will already be sorted by FeatureIndex (since that's the order the data is written to the _LCMSFeatureToPeakMap.txt file)
Public Sub DefineFeatureToPeakMapping(ByRef objPeakToFeatureMapping As clsParallelLngArrays)
    Dim lngIndex As Long
    Dim lngCachedCount As Long
    
    Dim lngFeatureIndex As Long
    
    Dim lngPeakIndex As Long
    Dim lngPeakIndexNew As Long
    Dim blnMatchFound As Boolean
    
    If objPeakToFeatureMapping Is Nothing Then
        Exit Sub
    End If
    
    lngCachedCount = objPeakToFeatureMapping.Count
    If lngCachedCount = 0 Then
        Exit Sub
    End If
    
    ' Initialize the progress bar since this might take awhile
    If Not mProgressForm Is Nothing Then
        mProgressForm.InitializeSubtask "Copying cached peak to feature mapping", 0, 100
    End If
            
    ' Make sure mHashMapOfPointsKept is sorted
    If Not mHashMapOfPointsKept.IsSorted Then
        mHashMapOfPointsKept.SortNow
    End If
    
    mFeatureToPeakMapCount = 0
        
    ReDim mFeatureToPeakMapFeatureIdx(lngCachedCount - 1)
    ReDim mFeatureToPeakMapPeakIdx(lngCachedCount - 1)
    
    ' Copy the data from objPeakToFeatureMapping to the FeatureToPeakMap arrays
    For lngIndex = 0 To lngCachedCount - 1
    
        lngPeakIndex = objPeakToFeatureMapping.KeyByIndex(lngIndex)
        
        ' Update the peak index using mHashMapOfPointsKept
        lngPeakIndexNew = mHashMapOfPointsKept.GetItemForKey(lngPeakIndex, blnMatchFound)
        
        If blnMatchFound Then
            
            ''' Temp Debug
            ''Debug.Assert lngPeakIndexNew <> 50599
            
            lngFeatureIndex = objPeakToFeatureMapping.ValueByIndex(lngIndex)
            
            ''' Temp Debug
            ''Debug.Assert lngFeatureIndex < 143 Or lngFeatureIndex > 145
            
            mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount) = lngFeatureIndex
            mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount) = lngPeakIndexNew
            mFeatureToPeakMapCount = mFeatureToPeakMapCount + 1
        Else
            ' Index not in mHashMapOfPointsKept; skip this mapping
        End If
        
        If lngIndex Mod 1000 = 0 Then
            mProgressPctComplete = lngIndex / CSng(lngCachedCount) * 100
            
            If Not mProgressForm Is Nothing Then
                mProgressForm.UpdateSubtaskProgressBar mProgressPctComplete, False
            End If
        End If
    Next lngIndex

    ' Shrink the arrays
    If mFeatureToPeakMapCount > 0 Then
        ReDim Preserve mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount - 1)
        ReDim Preserve mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount - 1)
    Else
        ReDim Preserve mFeatureToPeakMapFeatureIdx(0)
        ReDim Preserve mFeatureToPeakMapPeakIdx(0)
    End If
    
    mCachedFeatureToPeakMapDefined = True
    
End Sub

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
    GetColumnValueDbl = dblDefaultValue
    
End Function

Private Function GetColumnValueLng(ByRef strData() As String, ByVal intColumnIndex As Integer, Optional ByVal lngDefaultValue As Long = 0) As Long
    
    On Error GoTo GetColumnValueErrorHandler
    
    If intColumnIndex >= 0 Then
        If Len(strData(intColumnIndex)) > 0 Then
            GetColumnValueLng = CLng(strData(intColumnIndex))
        Else
            GetColumnValueLng = lngDefaultValue
        End If
    Else
        GetColumnValueLng = lngDefaultValue
    End If
    
    Exit Function
    
GetColumnValueErrorHandler:
    Debug.Assert False
    GetColumnValueLng = lngDefaultValue
    
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
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_CHARGE_MAX
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_DRIFT_TIME
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_CONFORMATION_FIT_SCORE
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_LC_FIT_SCORE
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_AVERAGE_ISOTOPIC_FIT
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_MEMBERS_PERCENTAGE
        strHeaders = strHeaders & vbTab & LCMSFEATURES_COLUMN_COMBINED_SCORE
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
                                     ByVal lngGelIndex As Long, _
                                     ByVal dblLCMSFeatureAbuMin As Double, _
                                     ByVal lngLCMSFeatureScanCountMin As Long, _
                                     ByVal dblIMSConformerScoreMin As Double) As Long
                                     
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
   
    Dim objLCMSFeatureIndexMap As clsParallelLngArrays
    Set objLCMSFeatureIndexMap = New clsParallelLngArrays
    objLCMSFeatureIndexMap.RemoveAll
       
    mGelIndex = lngGelIndex
    mErrorMessage = ""
    
    If mCachedFeatureToPeakMapDefined And mFeatureToPeakMapCount > 0 Then
        ' The mFeatureToPeakMap arrays have already been populated
    Else
        ' Initially reserve space for 50000 entries in the mFeatureToPeakMap arrays
        mFeatureToPeakMapCount = 0
        ReDim mFeatureToPeakMapFeatureIdx(50000)
        ReDim mFeatureToPeakMapPeakIdx(50000)
    End If
    
    If Not fso.FileExists(strLCMSFeaturesFilePath) Then
        mErrorMessage = "File not found: " & strLCMSFeaturesFilePath
        AddToAnalysisHistory mGelIndex, mErrorMessage
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
        AddToAnalysisHistory mGelIndex, mErrorMessage
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
        If Not (mCachedFeatureToPeakMapDefined And mFeatureToPeakMapCount > 0) Then
            lngByteCountTotal = lngByteCountTotal + FileLen(strLCMSFeatureToPeakMapFilePath)
        End If
    End If
    
    mProgressDescription = "Reading data"
    mProgressPctComplete = 0 / CSng(lngByteCountTotal) * 100
    
    ' Read the LCMSFeature file
    lngReturnValue = ReadLCMSFeaturesFile(fso, strLCMSFeaturesFilePath, _
                                          objLCMSFeatureIndexMap, _
                                          lngByteCountTotal, lngTotalBytesRead, _
                                          dblLCMSFeatureAbuMin, _
                                          lngLCMSFeatureScanCountMin, _
                                          dblIMSConformerScoreMin)
       
    If lngReturnValue = 0 Then
        If blnLoadFeatureToMapFile Then
            If mCachedFeatureToPeakMapDefined And mFeatureToPeakMapCount > 0 Then
                ' We do not need to re-read the feature to peak map file since
                ' mFeatureToPeakMapFeatureIdx() and mFeatureToPeakMapPeakIdx() have already been populated
                lngReturnValue = 0
            Else
                ' Read the LCMSFeature to Peak Map file
                lngReturnValue = ReadLCMSFeatureToPeakMapFile(fso, strLCMSFeatureToPeakMapFilePath, lngByteCountTotal, lngTotalBytesRead)
            End If
        Else
            blnSuccess = AutoMapDataPointsToFeatures(mAutoMapDataPointsMassTolerancePPM)
            
            lngReturnValue = 0
            If Not blnSuccess Then
                If mErrorMessage = "" Then
                    mErrorMessage = "Error auto-mapping featuers to loaded data points"
                    AddToAnalysisHistory mGelIndex, mErrorMessage
                End If
                lngReturnValue = -10
            End If
            
        End If
        
        If lngReturnValue = 0 Then
           
            mProgressDescription = "Sort peak to feature mapping"
            If Not mProgressForm Is Nothing Then
                mProgressForm.UpdateCurrentSubTask mProgressDescription
            End If

            ' Make sure the mFeatureToPeakMap arrays are sorted by Feature Index, then peak index
            ' In addition, make sure there are no duplicates
            SortFeatureMapMembers
            
            mProgressDescription = "Associating peaks with features"
            If Not mProgressForm Is Nothing Then
                mProgressForm.UpdateCurrentSubTask mProgressDescription
            End If
            
            ' Now update GelUMC() with the data in mFeatureToPeakMap
            blnSuccess = AddFeatureMembersToUMCs(objLCMSFeatureIndexMap)
        
            ' Override some settings in the UMC Def
            GelUMC(mGelIndex).def.UMCClassStatsUseStatsFromMostAbuChargeState = False
       
            GelUMC(mGelIndex).def.LoadedPredefinedLCMSFeatures = True
            
            If mPointsLoadMode = plmLoadOnePointPerLCMSFeature Then
                GelUMC(mGelIndex).def.OnePointPerLCMSFeature = True
            Else
                GelUMC(mGelIndex).def.OnePointPerLCMSFeature = False
            End If
        
            ' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
            Dim blnComputeClassMassAndAbundance As Boolean
            blnComputeClassMassAndAbundance = False
    
            UpdateUMCStatArrays mGelIndex, blnComputeClassMassAndAbundance, False

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
                                      ByRef objLCMSFeatureIndexMap As clsParallelLngArrays, _
                                      ByVal lngByteCountTotal As Long, _
                                      ByRef lngTotalBytesRead As Long, _
                                      ByVal dblLCMSFeatureAbuMin As Double, _
                                      ByVal lngLCMSFeatureScanCountMin As Long, _
                                      ByVal dblIMSConformerScoreMin As Double) As Long
                                      
    ' Returns 0 if no error, the error number if an error

    Dim tsInFile As TextStream
    Dim strLineIn As String
    Dim strUnknownColumnList As String
    Dim strMessage As String
    
    Dim lngIndex As Long
    Dim lngReturnValue As Long
    
    Dim blnColumnsDefined As Boolean
    Dim blnDataLine As Boolean
    Dim blnKeepUMC As Boolean
    
    Dim blnWarnedMissingCombinedScoreColumn As Boolean
    
    Dim strData() As String
    Dim strColumnHeader As String
    
    Dim intColumnMapping() As Integer
    
    Dim lngFeatureIndex As Long
    
    Dim lngScanStart As Long
    Dim lngScanEnd As Long
    Dim lngScanCenter As Long
    Dim lngLCMSFeatureScanCount As Long
    
    Dim dblMonoisotopicMass As Double
    Dim dblClassAbundance As Double
    
    Dim dblIMSConformerScore As Double
    Dim dblLCFitScore As Double
    Dim dblAvgIsotopicFit As Double
    Dim dblMembersPercentage As Double
    Dim dblCombinedScore As Double
    
    Dim dblConformerScoreToFilterOn As Double
    
    Dim blnSuccess As Boolean
    
On Error GoTo ReadLCMSFeaturesFileErrorHandler

    ReDim intColumnMapping(LCMSFEATURES_FILE_COLUMN_COUNT - 1)
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To LCMSFEATURES_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    mProgressDescription = "Reading LCMS Features file"
    If Not mProgressForm Is Nothing Then
        mProgressForm.InitializeSubtask mProgressDescription, 0, 100
    End If
    
    GelUMC(mGelIndex).UMCCnt = 0
    
    Set tsInFile = fso.OpenTextFile(strLCMSFeaturesFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream

        strLineIn = tsInFile.ReadLine
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If GelUMC(mGelIndex).UMCCnt Mod 50 = 0 Then
            mProgressPctComplete = lngTotalBytesRead / CSng(lngByteCountTotal) * 100
            If Not mProgressForm Is Nothing Then
                mProgressForm.UpdateSubtaskProgressBar mProgressPctComplete, True
            End If
            
            If KeyPressAbortProcess >= 2 Then
                Exit Do
            End If
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
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart) = LCMSFeaturesFileColumnConstants.ScanStart
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd) = LCMSFeaturesFileColumnConstants.ScanEnd
                    intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) = LCMSFeaturesFileColumnConstants.Scan
                    intColumnMapping(LCMSFeaturesFileColumnConstants.NET) = LCMSFeaturesFileColumnConstants.NET
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass) = LCMSFeaturesFileColumnConstants.MonoisotopicMass
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMassCalibrated) = LCMSFeaturesFileColumnConstants.MonoisotopicMassCalibrated
                    intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance) = LCMSFeaturesFileColumnConstants.Abundance
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepMZ) = LCMSFeaturesFileColumnConstants.ClassRepMZ
                    intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepCharge) = LCMSFeaturesFileColumnConstants.ClassRepCharge
                    
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MWMin) = LCMSFeaturesFileColumnConstants.MWMin
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MWMax) = LCMSFeaturesFileColumnConstants.MWMax
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MemberCount) = LCMSFeaturesFileColumnConstants.MemberCount
                    
                    intColumnMapping(LCMSFeaturesFileColumnConstants.DriftTime) = LCMSFeaturesFileColumnConstants.DriftTime

                    intColumnMapping(LCMSFeaturesFileColumnConstants.ConformationFitScore) = LCMSFeaturesFileColumnConstants.ConformationFitScore
                    intColumnMapping(LCMSFeaturesFileColumnConstants.LCFitScore) = LCMSFeaturesFileColumnConstants.LCFitScore
                    intColumnMapping(LCMSFeaturesFileColumnConstants.AverageIsotopicFit) = LCMSFeaturesFileColumnConstants.AverageIsotopicFit
                    intColumnMapping(LCMSFeaturesFileColumnConstants.MembersPercentage) = LCMSFeaturesFileColumnConstants.MembersPercentage
                    intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore) = LCMSFeaturesFileColumnConstants.CombinedScore

                    ' Column headers were not present
                     AddToAnalysisHistory mGelIndex, "LCMSFeatures file " & fso.GetFileName(strLCMSFeaturesFilePath) & " did not contain column headers; using the default headers (" & GetDefaultLCMSFeatureColumnHeaders(False) & ")"

                    blnDataLine = True
                    blnColumnsDefined = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, vbTab)
                    strUnknownColumnList = ""
                    
                    For lngIndex = 0 To UBound(strData)
                        strColumnHeader = StripQuotes(LCase(Trim(strData(lngIndex))))
                        
                        Select Case strColumnHeader
                        Case LCMSFEATURES_COLUMN_FEATURE_INDEX:         intColumnMapping(LCMSFeaturesFileColumnConstants.FeatureIndex) = lngIndex
                        Case LCMSFEATURES_COLUMN_ORIGINAL_INDEX
                            ' Ignore this column
                        Case LCMSFEATURES_COLUMN_SCAN:                  intColumnMapping(LCMSFeaturesFileColumnConstants.Scan) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_START:            intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_END:              intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd) = lngIndex
                        Case LCMSFEATURES_COLUMN_SCAN_ALIGNED
                            ' Legacy column, ignore it; shouldn't even be present anymore
                        Case LCMSFEATURES_COLUMN_NET:                   intColumnMapping(LCMSFeaturesFileColumnConstants.NET) = lngIndex
                        Case LCMSFEATURES_COLUMN_MONO_MASS:             intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass) = lngIndex
                        Case LCMSFEATURES_COLUMN_MONO_MASS_CALIBRATED:  intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMassCalibrated) = lngIndex
                        Case LCMSFEATURES_COLUMN_ABUNDANCE:             intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance) = lngIndex
                        Case LCMSFEATURES_COLUMN_CLASS_REP_MZ:          intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepMZ) = lngIndex
                        Case LCMSFEATURES_COLUMN_CLASS_REP_CHARGE:      intColumnMapping(LCMSFeaturesFileColumnConstants.ClassRepCharge) = lngIndex
                        Case LCMSFEATURES_COLUMN_UMC_MW_MIN:            intColumnMapping(LCMSFeaturesFileColumnConstants.MWMin) = lngIndex
                        Case LCMSFEATURES_COLUMN_UMC_MW_MAX:            intColumnMapping(LCMSFeaturesFileColumnConstants.MWMax) = lngIndex
                        Case LCMSFEATURES_COLUMN_UMC_MEMBER_COUNT:      intColumnMapping(LCMSFeaturesFileColumnConstants.MemberCount) = lngIndex

                        Case LCMSFEATURES_COLUMN_AVERAGE_MONO_MASS:
                            ' Ignore this column
                        Case LCMSFEATURES_COLUMN_ABUNDANCE_MAX:
                            ' Ignore this column
                        Case LCMSFEATURES_COLUMN_CHARGE_MAX
                            ' Ignore this column
                            
                        Case LCMSFEATURES_COLUMN_DRIFT_TIME
                            intColumnMapping(LCMSFeaturesFileColumnConstants.DriftTime) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_CONFORMATION_FIT_SCORE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.ConformationFitScore) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_LC_FIT_SCORE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.LCFitScore) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_AVERAGE_ISOTOPIC_FIT
                            intColumnMapping(LCMSFeaturesFileColumnConstants.AverageIsotopicFit) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_MEMBERS_PERCENTAGE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.MembersPercentage) = lngIndex
                            
                        Case LCMSFEATURES_COLUMN_COMBINED_SCORE
                            intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore) = lngIndex
                      
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
                        ' Unknown column headers were found; post an entry to the analysis history
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
            
            If lngFeatureIndex >= 0 And Not objLCMSFeatureIndexMap.Exists(lngFeatureIndex) Then
                
                lngScanStart = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart), 0)
                lngScanEnd = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd), 0)
                lngScanCenter = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.Scan), 0)
                dblMonoisotopicMass = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MonoisotopicMass), 0)
                dblClassAbundance = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.Abundance), 0)
                
                dblIMSConformerScore = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ConformationFitScore), 0)
                dblLCFitScore = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.LCFitScore), 0)
                dblAvgIsotopicFit = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.AverageIsotopicFit), 0)
                dblMembersPercentage = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MembersPercentage), 0)
                dblCombinedScore = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore), 0)
                
                If intColumnMapping(LCMSFeaturesFileColumnConstants.CombinedScore) >= 0 Then
                    ' Filter using Combined_Score
                    dblConformerScoreToFilterOn = dblCombinedScore
                Else
                    ' Filter using Conformer Score
                    dblConformerScoreToFilterOn = dblIMSConformerScore
                    
                    ' Log a warning the first time we do this
                    If Not blnWarnedMissingCombinedScoreColumn Then
                        AddToAnalysisHistory mGelIndex, "Warning: IMS Conformer Combined_Score value not found; will filter on Conformation_Fit_Score"
                        blnWarnedMissingCombinedScoreColumn = True
                    End If
                End If
                
                blnKeepUMC = True
                If dblLCMSFeatureAbuMin > 0 And dblClassAbundance < dblLCMSFeatureAbuMin Then
                    blnKeepUMC = False
                End If
    
                lngLCMSFeatureScanCount = lngScanEnd - lngScanStart + 1
                If lngLCMSFeatureScanCountMin > 1 And lngLCMSFeatureScanCount < lngLCMSFeatureScanCountMin Then
                    blnKeepUMC = False
                End If

                If dblIMSConformerScoreMin > 0 And dblConformerScoreToFilterOn < dblIMSConformerScoreMin Then
                    blnKeepUMC = False
                End If
    
                If blnKeepUMC Then
    
                    If GelUMC(mGelIndex).UMCCnt >= UBound(GelUMC(mGelIndex).UMCs) Then
                        ' Reserve more room in .UMCs()
                        blnSuccess = ManageClasses(mGelIndex, UMCMngAdd)
                    End If
                    
                    ' Define the new .UMC
                    With GelUMC(mGelIndex).UMCs(GelUMC(mGelIndex).UMCCnt)
                        
                        .ClassRepInd = 0
                        .ClassRepType = gldtIS
                        .ClassCount = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MemberCount), 0)
                        
                        If .ClassCount > 32767 Then
                            ' This is a 32-bit integer; cap the class count at 32767
                            .ClassCountPredefinedLCMSFeatures = 32767
                        Else
                            .ClassCountPredefinedLCMSFeatures = .ClassCount
                        End If
                        
                        
                        ReDim .ClassMInd(0)         ' members of the UMC(including representative)         (0-based array); index of the member in GelData
                        ReDim .ClassMType(0)        ' type of member: gldtCS or gldtIS
                        .ClassAbundance = dblClassAbundance
                        
                        .ClassMW = dblMonoisotopicMass
                        .ClassMWStD = 0
                        .ClassMassCorrectionDa = 0
                        
                        .ClassScore = dblConformerScoreToFilterOn
                        .ClassNET = 0
                        .ClassStatusBits = 0
                        
                        .MinScan = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanStart), lngScanCenter)
                        .MaxScan = GetColumnValueLng(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.ScanEnd), lngScanCenter)
                        
                        .MinMW = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MWMin), 0)
                        .MaxMW = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.MWMax), 0)
    
                        .DriftTime = GetColumnValueDbl(strData, intColumnMapping(LCMSFeaturesFileColumnConstants.DriftTime), 0)
                        
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
                    objLCMSFeatureIndexMap.add lngFeatureIndex, GelUMC(mGelIndex).UMCCnt
                    
                    GelUMC(mGelIndex).UMCCnt = GelUMC(mGelIndex).UMCCnt + 1
                    
                End If
            End If
        End If
    
    Loop
    
    tsInFile.Close

    ' Sort the data in objLCMSFeatureIndexMap
    objLCMSFeatureIndexMap.SortNow
    
    ' Shrink GelUMC().UMCs to length .UMCCnt
    blnSuccess = ManageClasses(mGelIndex, UMCMngTrim)
    
    AddToAnalysisHistory mGelIndex, "Loaded " & Format(GelUMC(mGelIndex).UMCCnt, "0,000") & " predefined LC-MS features"

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
                                              ByVal lngByteCountTotal As Long, _
                                              ByRef lngTotalBytesRead As Long) As Long

    ' Returns 0 if no error, the error number if an error

    Dim tsInFile As TextStream
    Dim strLineIn As String
    Dim strUnknownColumnList As String
    Dim strMessage As String
    
    Dim lngIndex As Long
    Dim lngReturnValue As Long
    Dim lngDataLineCount As Long
    
    Dim blnColumnsDefined As Boolean
    Dim blnDataLine As Boolean
    Dim blnValidPeak As Boolean
    
    Dim strData() As String
    Dim strColumnHeader As String
    
    Dim intColumnMapping() As Integer
    
    Dim lngFeatureIndex As Long
    Dim lngPeakIndex As Long
    Dim lngUMCIndex As Long
    
    Dim blnSuccess As Boolean
    
On Error GoTo ReadLCMSFeatureToPeakMapFileErrorHandler

    lngReturnValue = 0
    lngDataLineCount = 0
    
    ReDim intColumnMapping(PEAK_MAP_FILE_COLUMN_COUNT - 1)
    
    ' Set the column mappings to -1 (not present) for now
    For lngIndex = 0 To PEAK_MAP_FILE_COLUMN_COUNT - 1
        intColumnMapping(lngIndex) = -1
    Next lngIndex
    
    mProgressDescription = "Reading LCMS Feature to Peak Map file"
    If Not mProgressForm Is Nothing Then
        mProgressForm.UpdateCurrentSubTask mProgressDescription
    End If

    ' Make sure mHashMapOfPointsKept is sorted
    If Not mHashMapOfPointsKept.IsSorted Then
        mHashMapOfPointsKept.SortNow
    End If

    ' We will load the mappings into the mFeatureToPeakMap arrays, then update GelUMC
    ' The calling function should have already reserved space in the mFeatureToPeakMap arrays
    mFeatureToPeakMapCount = 0
    
    Set tsInFile = fso.OpenTextFile(strLCMSFeatureToPeakMapFilePath, ForReading, False)
    Do While Not tsInFile.AtEndOfStream

        strLineIn = tsInFile.ReadLine
        lngTotalBytesRead = lngTotalBytesRead + Len(strLineIn) + 2          ' Add 2 bytes to account for CrLf at end of line
        
        If mFeatureToPeakMapCount Mod 500 = 0 Then
            mProgressPctComplete = lngTotalBytesRead / CSng(lngByteCountTotal) * 100
            If Not mProgressForm Is Nothing Then
                mProgressForm.UpdateSubtaskProgressBar mProgressPctComplete, True
            End If
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
                    blnColumnsDefined = True
                Else
                    ' Define the column mappings
                    strData = Split(strLineIn, vbTab)
                    strUnknownColumnList = ""
                    
                    For lngIndex = 0 To UBound(strData)

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
                        ' Unknown column headers were found; post an entry to the analysis history
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
            
            ' Add 1 to lngPeakIndex since the data in the _LCMS_Features.txt file is 0-based while the
            '  data in GelData().IsoData() and .IsoData() is 1-based
            lngPeakIndex = lngPeakIndex + 1
            
            blnValidPeak = True
            
            If mPointsLoadMode >= plmLoadMappedPointsOnly Then
                If mPointsLoadMode = plmLoadOnePointPerLCMSFeature Then
                    If Not mHashMapOfPointsKept.Exists(lngPeakIndex) Then
                        blnValidPeak = False
                    End If
                End If
            End If
            
            If blnValidPeak Then
                ' Update the peak index using mHashMapOfPointsKept
                lngPeakIndex = mHashMapOfPointsKept.Item(lngPeakIndex)
            End If
                
            If blnValidPeak And lngFeatureIndex >= 0 And lngPeakIndex >= 0 Then
                
                ' Add this mapping to the mFeatureToPeakMap arrays
                
                If mFeatureToPeakMapCount >= UBound(mFeatureToPeakMapFeatureIdx) Then
                    ' Reserve more room in the mFeatureToPeakMap arrays
                    ReDim Preserve mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount * 2)
                    ReDim Preserve mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount * 2)
                End If
                
                mFeatureToPeakMapFeatureIdx(mFeatureToPeakMapCount) = lngFeatureIndex
                mFeatureToPeakMapPeakIdx(mFeatureToPeakMapCount) = lngPeakIndex
                
                mFeatureToPeakMapCount = mFeatureToPeakMapCount + 1
            End If
        
            lngDataLineCount = lngDataLineCount + 1
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
    
        AddToAnalysisHistory mGelIndex, "Parsed mapping data for " & Format(lngDataLineCount, "0,000") & " data points; retained feature to peak map information for " & Format(mFeatureToPeakMapCount, "0,000") & " data points"
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

' Sorts the data in mFeatureToPeakMapFeatureIdx and mFeatureToPeakMapPeakIdx
Private Sub SortFeatureMapMembers()

    Dim lngIndex As Long
    Dim lngIndexStart As Long
    Dim lngIndexEnd As Long
    Dim lngFeatureIndexCurrent As Long
    
    Dim lngPeakIndexCount As Long
    Dim lngPeakIndicesForFeature() As Long
    
    Dim lngNewFeatureToPeakMapCount As Long
    Dim lngNewPeakIndexCount As Long
    
    Dim blnSortRequired As Boolean
    
    ' Makes sure the peak indices associated with each Feature Index are sorted (and have no duplicates)
    
On Error GoTo SortFeatureMapMembersErrorHandler

    If mFeatureToPeakMapCount <= 0 Then
        ' Nothing to do; this is unexpected
        Debug.Assert False
        Exit Sub
    End If
    
    ' First check whether a sort is even required
    blnSortRequired = False
    For lngIndex = 0 To mFeatureToPeakMapCount - 2
        If mFeatureToPeakMapFeatureIdx(lngIndex) > mFeatureToPeakMapFeatureIdx(lngIndex + 1) Then
            blnSortRequired = True
            Exit For
        End If
    Next lngIndex
    
    If blnSortRequired Then
        ' Sort by feature index
        ShellSortLongWithParallelLong mFeatureToPeakMapFeatureIdx, mFeatureToPeakMapPeakIdx, 0, mFeatureToPeakMapCount - 1
    End If
    
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
                
        ''' Temp Debug
        '' Debug.Assert lngFeatureIndexCurrent < 143 Or lngFeatureIndexCurrent > 145
        
        Do While lngIndexEnd + 1 < mFeatureToPeakMapCount
            If mFeatureToPeakMapFeatureIdx(lngIndexEnd + 1) = lngFeatureIndexCurrent Then
                lngIndexEnd = lngIndexEnd + 1
            Else
                Exit Do
            End If
        Loop
    
        ' Now see if a sort is required or if duplicate values exist or if any of the entries have mFeatureToPeakMapPeakIdx(lngIndex) = 0
        blnSortRequired = False
        For lngIndex = lngIndexStart To lngIndexEnd - 1
            If mFeatureToPeakMapPeakIdx(lngIndex) = 0 Then
                blnSortRequired = True
                Exit For
            ElseIf mFeatureToPeakMapPeakIdx(lngIndex) >= mFeatureToPeakMapPeakIdx(lngIndex + 1) Then
                blnSortRequired = True
                Exit For
            End If
        Next lngIndex
        
        ' Make sure lngPeakIndicesForFeature has enough room for this feature's peaks
        If lngIndexEnd - lngIndexStart + 1 > UBound(lngPeakIndicesForFeature) Then
            ReDim lngPeakIndicesForFeature(lngIndexEnd - lngIndexStart)
        End If
        
        lngPeakIndexCount = 0
        For lngIndex = lngIndexStart To lngIndexEnd
            If mFeatureToPeakMapPeakIdx(lngIndex) > 0 Then
                lngPeakIndicesForFeature(lngPeakIndexCount) = mFeatureToPeakMapPeakIdx(lngIndex)
                lngPeakIndexCount = lngPeakIndexCount + 1
            End If
        Next lngIndex
        
        If blnSortRequired Then
    
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
            
        End If
        
        ' Now copy the data back into mFeatureToPeakMapFeatureIdx and mFeatureToPeakMapPeakIdx
        For lngIndex = 0 To lngPeakIndexCount - 1
            mFeatureToPeakMapFeatureIdx(lngNewFeatureToPeakMapCount) = lngFeatureIndexCurrent
            mFeatureToPeakMapPeakIdx(lngNewFeatureToPeakMapCount) = lngPeakIndicesForFeature(lngIndex)
            lngNewFeatureToPeakMapCount = lngNewFeatureToPeakMapCount + 1
        Next lngIndex
        
        lngIndexStart = lngIndexEnd + 1
    Loop
    
    ' Update mFeatureToPeakMapCount to match the new count (they will differ only if duplicates were removed)
    mFeatureToPeakMapCount = lngNewFeatureToPeakMapCount
    
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


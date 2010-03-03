VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilterDataArrayMaxCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' This class can be used to select the top N data points in a list, sorting descending
' It does not require a full sort of the data, which allows for faster filtering of the data
'
' To use, first call AddDataPoint() for each source data point, specifying the value to sort on and a data point index
' When done, call FilterData()
'  This routine will determine which data points to retain
'  For the remaining points, their data values will be changed to mSkipDataPointFlag (defaults to -1)


Private Const DEFAULT_SKIP_DATA_POINT_FLAG As Single = -1
    
' 4 steps in Sub FilterDataByMaxDataCountToLoad
Private Const SUBTASK_STEP_COUNT As Integer = 4
    

Private mDataCount As Long
Private mDataValues() As Single
Private mDataIndices() As Long

Private mMaximumDataCountToKeep As Long

Private mSkipDataPointFlag As Single

Private mTotalIntensityPercentageFilterEnabled As Boolean
Private mTotalIntensityPercentageFilter As Single

Private mProgress As Single

Public Property Let MaximumDataCountToLoad(Value As Long)
    mMaximumDataCountToKeep = Value
End Property
Public Property Get MaximumDataCountToLoad() As Long
    MaximumDataCountToLoad = mMaximumDataCountToKeep
End Property

Public Property Get Progress() As Single
    Progress = mProgress
End Property

Public Property Let SkipDataPointFlag(Value As Single)
    mSkipDataPointFlag = Value
End Property
Public Property Get SkipDataPointFlag() As Single
    SkipDataPointFlag = mSkipDataPointFlag
End Property

Public Property Let TotalIntensityPercentageFilterEnabled(Value As Boolean)
    mTotalIntensityPercentageFilterEnabled = Value
End Property
Public Property Get TotalIntensityPercentageFilterEnabled() As Boolean
    TotalIntensityPercentageFilterEnabled = mTotalIntensityPercentageFilterEnabled
End Property

Public Property Let TotalIntensityPercentageFilter(Value As Single)
    mTotalIntensityPercentageFilter = Value
End Property
Public Property Get TotalIntensityPercentageFilter() As Single
    TotalIntensityPercentageFilter = mTotalIntensityPercentageFilter
End Property


Public Sub AddDataPoint(ByVal sngAbundance As Single, ByVal lngDataPointIndex As Long)
    
    If mDataCount >= UBound(mDataValues) Then
        ReDim Preserve mDataValues(UBound(mDataValues) * 1.5)
        ReDim Preserve mDataIndices(UBound(mDataValues))
    End If
    
    mDataValues(mDataCount) = sngAbundance
    mDataIndices(mDataCount) = lngDataPointIndex
    
    mDataCount = mDataCount + 1
End Sub

Public Sub Clear()
    Const INITIAL_MEMORY_RESERVE As Long = 50000

    mMaximumDataCountToKeep = 400000
    
    mTotalIntensityPercentageFilterEnabled = False
    mTotalIntensityPercentageFilter = 90
                
    mDataCount = 0
    ReDim mDataValues(INITIAL_MEMORY_RESERVE)
    ReDim mDataIndices(INITIAL_MEMORY_RESERVE)
End Sub

Public Function GetAbundanceByIndex(ByVal lngDataPointIndex As Long) As Single
    If lngDataPointIndex >= 0 And lngDataPointIndex < mDataCount Then
        GetAbundanceByIndex = mDataValues(lngDataPointIndex)
    Else
        ' Invalid data point index value
        GetAbundanceByIndex = -1
        Debug.Assert False
    End If
End Function

Public Sub FilterData()

    If mDataCount <= 0 Then
        ' Nothing to do
    Else
        ' Shrink the arrays to mDataCount
        ReDim Preserve mDataValues(mDataCount - 1)
        ReDim Preserve mDataIndices(mDataCount - 1)

        FilterDataByMaxDataCountToKeep

    End If
    
End Sub

Private Sub FilterDataByMaxDataCountToKeep()

    Const HISTOGRAM_BIN_COUNT As Long = 5000
    
    Dim lngIndex As Long
    Dim lngPointTotal As Long
    Dim lngBinCount As Long
    Dim lngTargetBin As Long
    Dim lngBinToSort As Long
    Dim lngOriginalDataArrayIndex As Long
    
    Dim blnUseFullDataSort As Boolean
    
    Dim sngMaxAbundance As Single
    Dim dblBinSize As Double
    
    Dim lngHistogramBinCounts() As Long
    Dim dblHistogramBinStartIntensity() As Double
    
    Dim dblBinToSortAbundanceMinimum As Double
    Dim dblBinToSortAbundanceMaximum As Double
                
    Dim sngBinToSortAbundances() As Single
    Dim lngBinToSortDataIndices() As Long
    Dim lngBinToSortDataCount As Long
    Dim lngDataCountImplicitlyIncluded As Long
    
On Error GoTo ParsePrescannedDataErrorHandler

    mProgress = 0

    blnUseFullDataSort = False
    If mDataCount = 0 Then
        ' No data loaded
    ElseIf mDataCount <= mMaximumDataCountToKeep Then
        ' Loaded less than mMaximumDataCountToKeep data points
        ' Nothing to filter
    Else
    
        ' In order to speed up the sorting, we're first going to make a histogram
        '  (aka frequency distribution) of the abundances in mDataValues
    
        ' First, determine the maximum abundance value in mDataValues
        sngMaxAbundance = -1E+38
        For lngIndex = 0 To mDataCount - 1
            If mDataValues(lngIndex) > sngMaxAbundance Then
                sngMaxAbundance = mDataValues(lngIndex)
            End If
        Next lngIndex
        
        ' Round sngMaxAbundance up to the next highest integer
        sngMaxAbundance = Int(sngMaxAbundance) + 1
        
        ' Now determine the histogram bin size
        dblBinSize = sngMaxAbundance / HISTOGRAM_BIN_COUNT
        If dblBinSize < 1 Then dblBinSize = 1
        
        ' Initialize lngHistogramData
        lngBinCount = sngMaxAbundance / dblBinSize + 1
        ReDim lngHistogramBinCounts(lngBinCount - 1)
        ReDim dblHistogramBinStartIntensity(lngBinCount - 1)
        
        For lngIndex = 0 To lngBinCount - 1
            dblHistogramBinStartIntensity(lngIndex) = lngIndex * dblBinSize
        Next lngIndex
        
        ' Parse mDataValues to populate lngHistogramBinCounts
        For lngIndex = 0 To mDataCount - 1
            If mDataValues(lngIndex) <= 0 Then
                lngTargetBin = 0
            Else
                lngTargetBin = Int(mDataValues(lngIndex) / dblBinSize)
            End If
            
            If lngTargetBin < lngBinCount - 1 Then
                If mDataValues(lngIndex) >= dblHistogramBinStartIntensity(lngTargetBin + 1) Then
                    lngTargetBin = lngTargetBin + 1
                End If
            End If

            lngHistogramBinCounts(lngTargetBin) = lngHistogramBinCounts(lngTargetBin) + 1
            
            Debug.Assert mDataValues(lngIndex) >= dblHistogramBinStartIntensity(lngTargetBin)
            
            If lngIndex Mod 10000 = 0 Then
                mProgress = (0 + (lngIndex + 1) / CDbl(mDataCount)) / SUBTASK_STEP_COUNT * 100#
            End If
        Next lngIndex
        
        ' Now examine the frequencies in lngHistogramBinCounts() to determine the minimum value to consider when sorting
        lngPointTotal = 0
        lngBinToSort = -1
        For lngIndex = lngBinCount - 1 To 0 Step -1
            lngPointTotal = lngPointTotal + lngHistogramBinCounts(lngIndex)
            If lngPointTotal >= mMaximumDataCountToKeep Then
                lngBinToSort = lngIndex
                Exit For
            End If
        Next lngIndex
        
        mProgress = 1 / SUBTASK_STEP_COUNT * 100#
        
        If lngBinToSort >= 0 Then
            ' Find the data with intensity >= dblHistogramBinStartIntensity(lngBinToSort)
            ' We actually only need to sort the data in bin lngBinToSort
            
            dblBinToSortAbundanceMinimum = dblHistogramBinStartIntensity(lngBinToSort)
            dblBinToSortAbundanceMaximum = sngMaxAbundance + 1
            If lngBinToSort < lngBinCount - 1 Then
                dblBinToSortAbundanceMaximum = dblHistogramBinStartIntensity(lngBinToSort + 1)
            End If
            
            If dblBinToSortAbundanceMaximum = dblBinToSortAbundanceMinimum Then
                ' Is this code ever reached?
                ' If yes, then the code below won't populate sngBinToSortAbundances() and lngBinToSortDataIndices() with any data
                Debug.Assert False
                blnUseFullDataSort = True
            End If
            
            If Not blnUseFullDataSort Then
                lngBinToSortDataCount = 0
                If lngHistogramBinCounts(lngBinToSort) > 0 Then
                    ReDim sngBinToSortAbundances(lngHistogramBinCounts(lngBinToSort) - 1)
                    ReDim lngBinToSortDataIndices(lngHistogramBinCounts(lngBinToSort) - 1)
                Else
                    ' Is this code ever reached?
                    Debug.Assert False
                    blnUseFullDataSort = True
                End If
            End If
            
            If Not blnUseFullDataSort Then
                lngDataCountImplicitlyIncluded = 0
                For lngIndex = 0 To mDataCount - 1
                    If mDataValues(lngIndex) < dblBinToSortAbundanceMinimum Then
                        ' Skip this data point when re-reading the input data file
                        mDataValues(lngIndex) = mSkipDataPointFlag
                    ElseIf mDataValues(lngIndex) < dblBinToSortAbundanceMaximum Then
                        ' Value is in the bin to sort; add to the BinToSort arrays
                        sngBinToSortAbundances(lngBinToSortDataCount) = mDataValues(lngIndex)
                        lngBinToSortDataIndices(lngBinToSortDataCount) = mDataIndices(lngIndex)
                        lngBinToSortDataCount = lngBinToSortDataCount + 1
                    Else
                        lngDataCountImplicitlyIncluded = lngDataCountImplicitlyIncluded + 1
                    End If
                    
                    If lngIndex Mod 10000 = 0 Then
                        mProgress = (1 + (lngIndex + 1) / CDbl(mDataCount)) / SUBTASK_STEP_COUNT * 100#
                    End If
                Next lngIndex
                
                If lngBinToSortDataCount > 0 Then
                    If lngBinToSortDataCount < UBound(sngBinToSortAbundances) + 1 Then
                        ReDim Preserve sngBinToSortAbundances(lngBinToSortDataCount - 1)
                        ReDim Preserve lngBinToSortDataIndices(lngBinToSortDataCount - 1)
                    End If
                Else
                    ' This code shouldn't be reached
                    Debug.Assert False
                End If
                
                If mMaximumDataCountToKeep - lngDataCountImplicitlyIncluded - lngBinToSortDataCount = 0 Then
                    ' No need to sort and examine the data for BinToSort since we'll ultimately include all of it
                Else
                    SortAndMarkPointsToSkip sngBinToSortAbundances, lngBinToSortDataIndices, lngBinToSortDataCount, mMaximumDataCountToKeep - lngDataCountImplicitlyIncluded, SUBTASK_STEP_COUNT
                End If
            
                ' Synchronize the data in sngBinToSortAbundances and lngBinToSortDataIndices with mDataValues and mDataValues
                ' mDataValues and mDataIndices have not been sorted and therefore mDataIndices should currently be sorted ascending on "valid data point index"
                ' lngBinToSortDataIndices should also currently be sorted ascending on "valid data point index" so the following Do Loop within a For Loop should sync things up
                
                lngOriginalDataArrayIndex = 0
                For lngIndex = 0 To lngBinToSortDataCount - 1
                    Do While lngBinToSortDataIndices(lngIndex) > mDataIndices(lngOriginalDataArrayIndex)
                        lngOriginalDataArrayIndex = lngOriginalDataArrayIndex + 1
                    Loop
                    
                    If sngBinToSortAbundances(lngIndex) = mSkipDataPointFlag Then
                        If mDataIndices(lngOriginalDataArrayIndex) = lngBinToSortDataIndices(lngIndex) Then
                            mDataValues(lngOriginalDataArrayIndex) = mSkipDataPointFlag
                        Else
                        ' This code shouldn't be reached
                            Debug.Assert False
                        End If
                    End If
                    lngOriginalDataArrayIndex = lngOriginalDataArrayIndex + 1
                
                    If lngBinToSortDataCount < 1000 Or lngBinToSortDataCount Mod 100 = 0 Then
                        mProgress = (3 + (lngIndex + 1) / CDbl(lngBinToSortDataCount)) / SUBTASK_STEP_COUNT * 100#
                    End If
                Next lngIndex
            End If
        Else
            blnUseFullDataSort = True
        End If
        
        If blnUseFullDataSort Then
            ' This shouldn't normally be necessary
            Debug.Assert False
            
            ' We have to sort all of the data; this can be quite slow
            SortAndMarkPointsToSkip mDataValues, mDataIndices, mDataCount, mMaximumDataCountToKeep, SUBTASK_STEP_COUNT
        End If
        
    End If

    mProgress = 4 / SUBTASK_STEP_COUNT * 100#
    
    Exit Sub

ParsePrescannedDataErrorHandler:
    LogErrors Err.Number, "ParsePrescannedData"
    Debug.Assert False

End Sub

' This is sub uses a full sort to filter the data
' This will be slow for large arrays and you should therefore use FilterDataByMaxDataCountToKeep if possible
Private Sub SortAndMarkPointsToSkip(ByRef sngAbundances() As Single, ByRef lngDataIndices() As Long, ByVal lngDataCount As Long, ByVal lngMaximumDataCountInArraysToLoad As Long, ByVal intSubtaskStepCount As Integer)

    Dim lngIndex As Long

    If lngDataCount > 0 Then
        ' Sort sngAbundances ascending, sorting lngDataIndices in parallel
        ShellSortSingleWithParallelLong sngAbundances, lngDataIndices, 0, lngDataCount - 1

        mProgress = (2.333 / intSubtaskStepCount) * 100#

        ' Change the abundance values to mSkipDataPointFlag for data up to index lngDataCount-lngMaximumDataCountInArraysToLoad-1
        For lngIndex = 0 To lngDataCount - lngMaximumDataCountInArraysToLoad - 1
            sngAbundances(lngIndex) = mSkipDataPointFlag
        Next lngIndex

        mProgress = (2.666 / intSubtaskStepCount) * 100#

        ' Re-sort, this time on lngDataIndices with sngAbundances in parallel
        ShellSortLongWithParallelSingle lngDataIndices, sngAbundances, 0, lngDataCount - 1
    End If

    mProgress = 3 / intSubtaskStepCount * 100#

End Sub

Private Sub Class_Initialize()
    mSkipDataPointFlag = DEFAULT_SKIP_DATA_POINT_FLAG
    Me.Clear
End Sub




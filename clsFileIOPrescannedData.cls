VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileIOPrescannedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The following arrays are populated when mMaximumDataCountEnabled = True
' They contain the abundance values and data indices of the data that passes
'  the various import filters (EvenOddScanFilter, MSLevel, Abundance, Fit)
' After pre-scanning the file, mPrescanDataAbundances is sorted and mPrescanDataIndices is sorted in parallel
' The top mMaximumDataCountToLoad points by abundance are determined, and the abundance for the remaining points
'  is set to SKIP_DATA_POINT_FLAG = -1 in mPrescanDataAbundances
' The arrays are then re-sorted, this time on mPrescanDataIndices with in parallel
' The file is re-read, and only the data with mPrescanDataAbundances(lngIndex) >= 0 is kept

Private Const SKIP_DATA_POINT_FLAG As Single = -1

Private mPrescanDataCount As Long
Private mPrescanDataAbundances() As Single
Private mPrescanDataIndices() As Long

Private mMaximumDataCountToLoad As Long

Public Property Get MaximumDataCountToLoad() As Long
    MaximumDataCountToLoad = mMaximumDataCountToLoad
End Property
Public Property Let MaximumDataCountToLoad(Value As Long)
    mMaximumDataCountToLoad = Value
End Property

Public Sub AddDataPoint(sngAbundance As Single, lngDataPointIndex As Long)
    
    If mPrescanDataCount >= UBound(mPrescanDataAbundances) Then
        ReDim Preserve mPrescanDataAbundances(UBound(mPrescanDataAbundances) * 1.5)
        ReDim Preserve mPrescanDataIndices(UBound(mPrescanDataAbundances))
    End If
    
    mPrescanDataAbundances(mPrescanDataCount) = sngAbundance
    mPrescanDataIndices(mPrescanDataCount) = lngDataPointIndex
    
    mPrescanDataCount = mPrescanDataCount + 1
End Sub

Public Sub Clear()
    Const INITIAL_MEMORY_RESERVE As Long = 25000

    mMaximumDataCountToLoad = 400000
                
    mPrescanDataCount = 0
    ReDim mPrescanDataAbundances(INITIAL_MEMORY_RESERVE)
    ReDim mPrescanDataIndices(INITIAL_MEMORY_RESERVE)
End Sub

Public Function GetAbundanceByIndex(lngDataPointIndex As Long) As Single
    If lngDataPointIndex >= 0 And lngDataPointIndex < mPrescanDataCount Then
        GetAbundanceByIndex = mPrescanDataAbundances(lngDataPointIndex)
    Else
        ' Invalid data point index value
        GetAbundanceByIndex = -1
        Debug.Assert False
    End If
End Function

Public Sub ParsePrescannedData()
    Const HISTOGRAM_BIN_COUNT As Long = 5000
    Const SUBTASK_STEP_COUNT As Integer = 4
    
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
    
    Dim dblSubtaskProgress As Double
    
On Error GoTo ParsePrescannedDataErrorHandler

    If mPrescanDataCount <= 0 Then
        ' Nothing to do
    Else
        dblSubtaskProgress = 0
        frmProgress.InitializeSubtask "Sorting pre-scanned data", 0, 100
    
        ' Shrink the arrays to mPrescanDataCount
        ReDim Preserve mPrescanDataAbundances(mPrescanDataCount - 1)
        ReDim Preserve mPrescanDataIndices(mPrescanDataCount - 1)
        
        blnUseFullDataSort = False
        If mPrescanDataCount > mMaximumDataCountToLoad Then
        
            ' In order to speed up the sorting, we're first going to make a histogram
            '  (aka frequency distribution) of the abundances in mPrescanDataAbundances
        
            ' First, determine the maximum abundance value in mPrescanDataAbundances
            sngMaxAbundance = -1E+38
            For lngIndex = 0 To mPrescanDataCount - 1
                If mPrescanDataAbundances(lngIndex) > sngMaxAbundance Then
                    sngMaxAbundance = mPrescanDataAbundances(lngIndex)
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
            
            ' Parse mPrescanDataAbundances to populate lngHistogramBinCounts
            For lngIndex = 0 To mPrescanDataCount - 1
                If mPrescanDataAbundances(lngIndex) <= 0 Then
                    lngTargetBin = 0
                Else
                    lngTargetBin = Int(mPrescanDataAbundances(lngIndex) / dblBinSize)
                End If
                
                If lngTargetBin < lngBinCount - 1 Then
                    If mPrescanDataAbundances(lngIndex) >= dblHistogramBinStartIntensity(lngTargetBin + 1) Then
                        lngTargetBin = lngTargetBin + 1
                    End If
                End If

                lngHistogramBinCounts(lngTargetBin) = lngHistogramBinCounts(lngTargetBin) + 1
                
                Debug.Assert mPrescanDataAbundances(lngIndex) >= dblHistogramBinStartIntensity(lngTargetBin)
                
                If lngIndex Mod 10000 = 0 Then
                    dblSubtaskProgress = (0 + (lngIndex + 1) / CDbl(mPrescanDataCount)) / SUBTASK_STEP_COUNT * 100#
                    frmProgress.UpdateSubtaskProgressBar dblSubtaskProgress
                End If
            Next lngIndex
            
            ' Now examine the frequencies in lngHistogramBinCounts() to determine the minimum abundance to consider when sorting
            lngPointTotal = 0
            lngBinToSort = -1
            For lngIndex = lngBinCount - 1 To 0 Step -1
                lngPointTotal = lngPointTotal + lngHistogramBinCounts(lngIndex)
                If lngPointTotal >= mMaximumDataCountToLoad Then
                    lngBinToSort = lngIndex
                    Exit For
                End If
            Next lngIndex
            
            frmProgress.UpdateSubtaskProgressBar 1 / SUBTASK_STEP_COUNT * 100#
            
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
                    For lngIndex = 0 To mPrescanDataCount - 1
                        If mPrescanDataAbundances(lngIndex) < dblBinToSortAbundanceMinimum Then
                            ' Skip this data point when re-reading the input data file
                            mPrescanDataAbundances(lngIndex) = SKIP_DATA_POINT_FLAG
                        ElseIf mPrescanDataAbundances(lngIndex) < dblBinToSortAbundanceMaximum Then
                            ' Value is in the bin to sort; add to the BinToSort arrays
                            sngBinToSortAbundances(lngBinToSortDataCount) = mPrescanDataAbundances(lngIndex)
                            lngBinToSortDataIndices(lngBinToSortDataCount) = mPrescanDataIndices(lngIndex)
                            lngBinToSortDataCount = lngBinToSortDataCount + 1
                        Else
                            lngDataCountImplicitlyIncluded = lngDataCountImplicitlyIncluded + 1
                        End If
                        
                        If lngIndex Mod 10000 = 0 Then
                            dblSubtaskProgress = (1 + (lngIndex + 1) / CDbl(mPrescanDataCount)) / SUBTASK_STEP_COUNT * 100#
                            frmProgress.UpdateSubtaskProgressBar dblSubtaskProgress
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
                    
                    If mMaximumDataCountToLoad - lngDataCountImplicitlyIncluded - lngBinToSortDataCount = 0 Then
                        ' No need to sort and examine the data for BinToSort since we'll ultimately include all of it
                    Else
                        PrescannedDataSortAndMarkPointsToSkip sngBinToSortAbundances, lngBinToSortDataIndices, lngBinToSortDataCount, mMaximumDataCountToLoad - lngDataCountImplicitlyIncluded, SUBTASK_STEP_COUNT
                    End If
                
                    ' Synchronize the data in sngBinToSortAbundances and lngBinToSortDataIndices with mPrescanDataAbundances and mPrescanDataAbundances
                    ' mPrescanDataAbundances and mPrescanDataIndices have not been sorted and therefore mPrescanDataIndices should be sorted ascending on on "valid data point index"
                    ' lngBinToSortDataIndices should also be sorted ascending on "valid data point index" so the following Do Loop should sync things up
                    
                    lngOriginalDataArrayIndex = 0
                    For lngIndex = 0 To lngBinToSortDataCount - 1
                        Do While lngBinToSortDataIndices(lngIndex) > mPrescanDataIndices(lngOriginalDataArrayIndex)
                            lngOriginalDataArrayIndex = lngOriginalDataArrayIndex + 1
                        Loop
                        
                        If sngBinToSortAbundances(lngIndex) = SKIP_DATA_POINT_FLAG Then
                            If mPrescanDataIndices(lngOriginalDataArrayIndex) = lngBinToSortDataIndices(lngIndex) Then
                                mPrescanDataAbundances(lngOriginalDataArrayIndex) = SKIP_DATA_POINT_FLAG
                            Else
                            ' This code shouldn't be reached
                                Debug.Assert False
                            End If
                        End If
                        lngOriginalDataArrayIndex = lngOriginalDataArrayIndex + 1
                    
                        If lngBinToSortDataCount < 1000 Or lngBinToSortDataCount Mod 100 = 0 Then
                            dblSubtaskProgress = (3 + (lngIndex + 1) / CDbl(lngBinToSortDataCount)) / SUBTASK_STEP_COUNT * 100#
                            frmProgress.UpdateSubtaskProgressBar dblSubtaskProgress
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
                PrescannedDataSortAndMarkPointsToSkip mPrescanDataAbundances, mPrescanDataIndices, mPrescanDataCount, mMaximumDataCountToLoad, SUBTASK_STEP_COUNT
            End If
            
        End If
    
        frmProgress.UpdateSubtaskProgressBar SUBTASK_STEP_COUNT * 100#
    End If
    
    Exit Sub

ParsePrescannedDataErrorHandler:
    LogErrors Err.Number, "ReadMZXMLFile"
    Debug.Assert False

End Sub

Public Sub PrescannedDataSortAndMarkPointsToSkip(ByRef sngAbundances() As Single, ByRef lngDataIndices() As Long, ByVal lngDataCount As Long, ByVal lngMaximumDataCountInArraysToLoad As Long, ByVal intSubtaskStepCount As Integer)
    Dim lngIndex As Long
    
    If lngDataCount > 0 Then
        ' Sort sngAbundances ascending, sorting lngDataIndices in parallel
        ShellSortSingleWithParallelLong sngAbundances, lngDataIndices, 0, lngDataCount - 1
    
        frmProgress.UpdateSubtaskProgressBar (2.333 / intSubtaskStepCount) * 100#
    
        ' Change the abundance values to SKIP_DATA_POINT_FLAG for data up to index lngDataCount-lngMaximumDataCountInArraysToLoad-1
        For lngIndex = 0 To lngDataCount - lngMaximumDataCountInArraysToLoad - 1
            sngAbundances(lngIndex) = SKIP_DATA_POINT_FLAG
        Next lngIndex
        
        frmProgress.UpdateSubtaskProgressBar (2.666 / intSubtaskStepCount) * 100#
        
        ' Re-sort, this time on lngDataIndices with sngAbundances in parallel
        ShellSortLongWithParallelSingle lngDataIndices, sngAbundances, 0, lngDataCount - 1
    End If
    
    frmProgress.UpdateSubtaskProgressBar 3 / intSubtaskStepCount * 100#
    
End Sub

Private Sub Class_Initialize()
    Me.Clear
End Sub


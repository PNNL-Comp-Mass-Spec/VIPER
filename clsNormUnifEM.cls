VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNormUnifEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Implments normal, uniform, expectation maximization

Private Const VARIANCE_MINIMUM As Double = 0.0001
Private Const DEFAULT_VARIANCE As Double = 0.1

Private Const DEFAULT_ITERATION_COUNT As Integer = 32
Private Const DEFAULT_DATA_RANGE_EXCLUSION_PERCENT As Integer = 5

Private Const MIN_ALLOWED_ITERATION_COUNT As Integer = 5
Private Const MAX_ALLOWED_DATA_RANGE_EXCLUSION_PERCENT As Integer = 75

Private mIterationCount As Integer
Private mPercentOfDataToExclude As Integer       ' Value between 0 and MAX_ALLOWED_DATA_RANGE_EXCLUSION_PERCENT; used when determining minimum and maximum values in the range of data to use for Expectation Maximization; excludes mPercentOfDataToExclude of the data (by removing DataCount*mPercentOfDataToExclude/100/2 values from the extremes of the array)

Private mMean As Double
Private mVariance As Double
Private mNormFraction As Double

Private mAbortProcessing As Boolean
Private mProgressPercentComplete As Single          ' Value between 0 and 100
Private mProgressMessage As String

Public Event ProgessReset()
Public Event ProgessUpdate()
Public Event ProgessComplete()

Public Property Get IterationCount() As Integer
    IterationCount = mIterationCount
End Property
Public Property Let IterationCount(Value As Integer)
    If Value < MIN_ALLOWED_ITERATION_COUNT Then Value = MIN_ALLOWED_ITERATION_COUNT
    mIterationCount = Value
End Property

Public Property Get PercentOfDataToExclude() As Integer
    PercentOfDataToExclude = mPercentOfDataToExclude
End Property
Public Property Let PercentOfDataToExclude(Value As Integer)
    If Value < 0 Then Value = 0
    If Value > MAX_ALLOWED_DATA_RANGE_EXCLUSION_PERCENT Then Value = MAX_ALLOWED_DATA_RANGE_EXCLUSION_PERCENT
    mPercentOfDataToExclude = Value
End Property

Public Property Get Mean() As Double
    Mean = mMean
End Property
Public Property Get Variance() As Double
    Variance = mVariance
End Property
Public Property Get NormFraction() As Double
    NormFraction = mNormFraction
End Property

Public Property Get ProgressPercentComplete() As Single
    ProgressPercentComplete = mProgressPercentComplete
End Property
Public Property Get ProgressMessage() As String
    ProgressMessage = mProgressMessage
End Property

Public Sub AbortProcessingNow()
    mAbortProcessing = True
End Sub

Public Function ComputeDistribution(ByRef dblValues() As Double, ByVal dblVarianceEstimate As Double) As Boolean
    ' When calling, set dblVarianceEstimate to a rough estimate of error. For example, if we think mass error
    ' is 5 ppm at the base of a mass error histogram (where the base is 2.5*sigma),
    ' then compute dblVarianceEstimate = (5 / 2.5)^2
    '
    ' Also, note that if mPercentOfDataToExclude > 0, then dblValues() will be sorted ascending
    '  and we'll only use the mPercentOfDataToExclude of the data (removing DataCount*mPercentOfDataToExclude/100/2 values from the extremes of the array)
        
    Dim blnSuccess As Boolean
    
    Dim lngPointNum As Long
    Dim intIteration As Integer
    Dim lngDataCount As Long
    Dim lngTargetDataPoint As Long
    
    Dim dblMinimumValue As Double
    Dim dblMaximumValue As Double
    
    Dim dblAlternateMinimumValue As Double
    Dim dblAlternateMaximumValue As Double
    
    Dim dblInverseRange As Double       ' Aka u
    Dim dblMeanNext As Double
    Dim dblVarianceNext As Double
    Dim dblNormFractionNext As Double
    
    Dim dblCurrentValue As Double
    Dim dblDiff As Double
    Dim dblNormalProbability As Double
    Dim dblPosteriorNormalProbability As Double
            
    Dim dblUniformProbabilityVals() As Double
    
    Dim dblVariancePrevious As Double
    Dim intIterationCountUnchanged As Integer
    
On Error GoTo ComputeDistributionErrorHandler

    ResetProgress "Starting calculations"
    ResetStats

    lngDataCount = UBound(dblValues) + 1
    If lngDataCount <= 1 Then
        mMean = 0
        mVariance = 0.1
        mNormFraction = 0
        ComputeDistribution = False
        Exit Function
    End If
    
    If dblVarianceEstimate <> 0 Then
        mVariance = dblVarianceEstimate
    Else
        mVariance = DEFAULT_VARIANCE
    End If
    
    ' Determine the minimum and maximum values
    dblMinimumValue = dblValues(0)
    dblMaximumValue = dblValues(0)
    For lngPointNum = 1 To lngDataCount - 1
        If dblValues(lngPointNum) < dblMinimumValue Then dblMinimumValue = dblValues(lngPointNum)
        If dblValues(lngPointNum) > dblMaximumValue Then dblMaximumValue = dblValues(lngPointNum)
    Next lngPointNum
    
    If dblMinimumValue = dblMaximumValue Then
        ' Minimum and maximum are the same
        mMean = dblMaximumValue
        mVariance = 0.1
        mNormFraction = 0
        ComputeDistribution = False
        Exit Function
    End If
    
    If mPercentOfDataToExclude > 0 And mPercentOfDataToExclude < 100 Then
        UpdateProgress "Sorting the Data", 0
        
        ' Sort the data
        ShellSortDouble dblValues, 0, UBound(dblValues)
        
        ' Set dblMinimumValue to the value of the data point at the mPercentOfDataToExclude'th percentile
        lngTargetDataPoint = lngDataCount * (mPercentOfDataToExclude / 100# / 2#)
        If lngTargetDataPoint < 0 Then lngTargetDataPoint = 0
        dblAlternateMinimumValue = dblValues(lngTargetDataPoint)
        
        lngTargetDataPoint = lngDataCount * (1 - (mPercentOfDataToExclude / 100# / 2#))
        If lngTargetDataPoint >= lngDataCount Then lngTargetDataPoint = lngDataCount - 1
        dblAlternateMaximumValue = dblValues(lngTargetDataPoint)
        
        If dblAlternateMaximumValue > dblAlternateMinimumValue Then
            dblMinimumValue = dblAlternateMinimumValue
            dblMaximumValue = dblAlternateMaximumValue
        End If
    End If
    
    dblInverseRange = 1 / (dblMaximumValue - dblMinimumValue)
    
    ReDim dblUniformProbabilityVals(lngDataCount - 1)

    dblVariancePrevious = 0
    intIterationCountUnchanged = 0
    
    UpdateProgress "Processing", 0
    If mAbortProcessing Then
        ComputeDistribution = False
        Exit Function
    End If
    
    ' Initialize the mean and norm fraction; variance should have been initialized above
    mMean = 0
    mNormFraction = 0.5
    For intIteration = 0 To mIterationCount - 1
        dblMeanNext = 0
        dblVarianceNext = 0
        dblNormFractionNext = 0
        For lngPointNum = 0 To lngDataCount - 1
            dblCurrentValue = dblValues(lngPointNum)
            dblDiff = (dblCurrentValue - mMean)
            dblNormalProbability = Exp(-(0.5 * dblDiff * dblDiff) / mVariance) / (Sqr(2 * 3.14159) * Sqr(mVariance))
            dblPosteriorNormalProbability = (dblNormalProbability * mNormFraction) / (dblNormalProbability * mNormFraction + (1 - mNormFraction) * dblInverseRange)
            dblUniformProbabilityVals(lngPointNum) = dblPosteriorNormalProbability

            dblNormFractionNext = dblNormFractionNext + dblPosteriorNormalProbability
            dblMeanNext = dblMeanNext + dblPosteriorNormalProbability * dblCurrentValue
            dblVarianceNext = dblVarianceNext + dblPosteriorNormalProbability * (dblCurrentValue - mMean) * (dblCurrentValue - mMean)
        Next lngPointNum
        
        mNormFraction = dblNormFractionNext / lngDataCount
        mMean = dblMeanNext / dblNormFractionNext
        mVariance = dblVarianceNext / dblNormFractionNext
        
        If mVariance < VARIANCE_MINIMUM Then
            Exit For
        End If
        
        If intIteration > 0 Then
            If mVariance = dblVariancePrevious Then
                intIterationCountUnchanged = intIterationCountUnchanged + 1
            ElseIf mVariance > 0 Then
                If Abs((mVariance - dblVariancePrevious) / mVariance) < 0.0001 Then
                    intIterationCountUnchanged = intIterationCountUnchanged + 1
                Else
                    intIterationCountUnchanged = 0
                End If
            Else
                intIterationCountUnchanged = 0
            End If
            
             If intIterationCountUnchanged = 3 Then
                ' Same variance 3 times in a row
                Exit For
            End If
        End If
        
        dblVariancePrevious = mVariance
        
        UpdateProgress "Processing", (intIteration + 1) / CSng(mIterationCount) * 100
        If mAbortProcessing Then
            ComputeDistribution = False
        Exit Function
    End If

    Next intIteration
    
    UpdateProgressComplete
    
    blnSuccess = True
    ComputeDistribution = blnSuccess
    
    Exit Function
    
ComputeDistributionErrorHandler:
    Debug.Assert False
    ComputeDistribution = False

End Function

Private Sub ResetProgress(ByVal strMessage As String)
    mAbortProcessing = False
    mProgressPercentComplete = 0
    mProgressMessage = strMessage
    RaiseEvent ProgessReset
End Sub

Private Sub ResetStats()
    mMean = 0
    mVariance = 0
    mNormFraction = 0.5
End Sub

Public Sub SetDefaultOptions()
    mIterationCount = DEFAULT_ITERATION_COUNT
    mPercentOfDataToExclude = DEFAULT_DATA_RANGE_EXCLUSION_PERCENT
End Sub

Private Sub ShellSortDouble(ByRef dblArray() As Double, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim dblCompareVal As Double

On Error GoTo ShellSortDoubleErrorHandler

' sort array[lngLowIndex..lngHighIndex]

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            dblCompareVal = dblArray(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If dblArray(lngIndexCompare) <= dblCompareVal Then Exit For
                dblArray(lngIndexCompare + lngIncrement) = dblArray(lngIndexCompare)
            Next lngIndexCompare
            dblArray(lngIndexCompare + lngIncrement) = dblCompareVal
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop

''    Debug.Assert VerifySort(dblArray(), lngLowIndex, lngHighIndex)
    Exit Sub

ShellSortDoubleErrorHandler:
    Debug.Assert False
End Sub

Private Sub UpdateProgress(ByVal strMessage As String, ByVal sngPercentComplete As Single)
    mProgressMessage = strMessage
    mProgressPercentComplete = sngPercentComplete
    RaiseEvent ProgessUpdate
    DoEvents
End Sub

Private Sub UpdateProgressComplete()
     UpdateProgress "Complete", 100
     RaiseEvent ProgessComplete
     DoEvents
End Sub
Private Sub Class_Initialize()
    SetDefaultOptions
    ResetProgress ""
    ResetStats
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatisticalMoments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Computes the Statistical Moments for a peak in a given array
' The array should ideally contain just one peak
' The user may optionally provide the starting and ending indices in the array to
'  use initially for iteration

Private Type udtPeakStatsType
    MaximumIntensity As Double
    MaximumIntensityLoc As Long
    Plates As Long
    PeakArea As Double
    PeakCenter As Double
    Variance As Double
    Sigma As Double
    WidthAtFourSigma As Double
    Skew As Double                  ' < 0 means fronted, > 0 means tailed
    lngIterationCount As Long
End Type

Private mData() As Double       ' 0-based array
Private mDataCount As Long
Private mLowIndex As Long
Private mHighIndex As Long
Private mDeltaX As Double

Private mPeakStats As udtPeakStatsType

Private mIndexLeftInitial As Long
Private mIndexRightInitial As Long

Private mCurIndexLeft As Long
Private mCurIndexRight As Long

Public Sub AutoFindIterationStartPoints(ByRef lngIndexLeft As Long, ByRef lngIndexRight As Long, Optional ByVal sngPercentHighestIntensityTarget As Single = 0.2, Optional ByVal lngMinimumDistanceFromCenter As Long = 1, Optional blnUpdateIterationStartPoints As Boolean = True)
    ' Finds the highest point in mData()
    ' Searches left and right until a point is found whose intensity is sngPercentHighestIntensityTarget of the highest point
    ' However, if lngMinimumDistanceFromCenter is > 1, then makes sure the point is at least
    '  that distance from the highest point
    
    Dim lngIndex As Long
    Dim dblMaximumValue As Double, dblThreshold As Double
    Dim lngIndexOfMaximum As Long
    
    If mDataCount = 0 Then Exit Sub
    
    lngIndexOfMaximum = mLowIndex
    dblMaximumValue = mData(mLowIndex)
    For lngIndex = mLowIndex + 1 To mHighIndex
        If mData(lngIndex) > dblMaximumValue Then
            dblMaximumValue = mData(lngIndex)
            lngIndexOfMaximum = lngIndex
        End If
    Next lngIndex
    
    ' Compute the threshold to look for
    dblThreshold = sngPercentHighestIntensityTarget * dblMaximumValue
    
    ' Initialize lngIndexLeft and lngIndexRight to the endpoints
    mIndexLeftInitial = mLowIndex
    mIndexRightInitial = mHighIndex
    
    lngIndexLeft = mIndexLeftInitial
    lngIndexRight = mIndexRightInitial
    
    ' Find the right point
    For lngIndex = lngIndexOfMaximum + 1 To mHighIndex
        If mData(lngIndex) <= dblThreshold Then
            If lngIndex - lngIndexOfMaximum >= lngMinimumDistanceFromCenter Then
                lngIndexRight = lngIndex
                Exit For
            End If
        End If
    Next lngIndex
    
    ' Find the left point
    For lngIndex = lngIndexOfMaximum - 1 To mLowIndex Step -1
        If mData(lngIndex) <= dblThreshold Then
            If lngIndexOfMaximum - lngIndex >= lngMinimumDistanceFromCenter Then
                lngIndexLeft = lngIndex
                Exit For
            End If
        End If
    Next lngIndex
    
    If blnUpdateIterationStartPoints Then
        ' Update mIndexRightInitial and mIndexLeftInitial
        mIndexLeftInitial = lngIndexLeft
        mIndexRightInitial = lngIndexRight
    End If
    
End Sub

Public Sub ClearData()
    ReDim mData(0)
    mDataCount = 0
End Sub

Public Function GetDeltaX() As Double
    GetDeltaX = mDeltaX
End Function

Public Sub GetIterationStartPoints(ByRef lngIndexLeft As Long, ByRef lngIndexRight As Long)
    
    lngIndexLeft = mIndexLeftInitial
    lngIndexRight = mIndexRightInitial

End Sub

Public Function SetData(dblDataArray() As Double, lngLowIndex As Long, lngHighIndex As Long) As Boolean
    ' Returns True if success, False if failure

    Dim lngIndex As Long
    
On Error GoTo SetDataErrorHandler

    If lngLowIndex > lngHighIndex Then
        SetData = False
        Exit Function
    End If
    
    ReDim mData(lngLowIndex To lngHighIndex)
    
    For lngIndex = lngLowIndex To lngHighIndex
        mData(lngIndex) = dblDataArray(lngIndex)
    Next lngIndex

    mLowIndex = lngLowIndex
    mHighIndex = lngHighIndex
    mDataCount = lngHighIndex - lngLowIndex + 1
    
    SetData = True
    Exit Function
    
SetDataErrorHandler:
    Debug.Print "Error in SetData: " & Err.Description
    Debug.Assert False
    SetData = False
    
End Function

Public Sub SetDeltaX(dblDeltaX As Double)
    If dblDeltaX > 0 Then
        mDeltaX = dblDeltaX
    End If
End Sub

Public Sub SetIterationStartPoints(lngIndexLeft As Long, lngIndexRight As Long)
    ' Update mIndexRightInitial and mIndexLeftInitial
    mIndexLeftInitial = lngIndexLeft
    mIndexRightInitial = lngIndexRight

    If mIndexLeftInitial < mLowIndex Then mIndexLeftInitial = mLowIndex
    If mIndexRightInitial > mHighIndex Then mIndexRightInitial = mHighIndex
    
End Sub

Public Function FindStatisticalMoments(Optional ByRef lngIndexLeftReturn As Long, Optional ByRef lngIndexRightReturn As Long, Optional sngTargetSigmaValue As Single = 2, Optional lngMaximumIterations As Long = 1000) As Boolean
    ' Uses Iteration to find the statistical moments for the data in mData()
    ' Will iterate until mIndexLeftInitial and mIndexRightInitial are at +-sngTargetSigmaValue
    ' Thus, if sngTargetSigmaValue = 2, then will find 4 sigma width
    ' Returns True if success, False if failure
    
    Dim LastLP As Long, LastRP As Long, CalcLP As Long, CalcRP As Long
    Dim lngIterationCount As Long
    Dim lngIndex As Long
    Dim dblThreshold As Double
    Dim blnSuccess As Boolean
    
    Dim Moments(0 To 3) As Double       ' 0-based array
    Dim TempArray() As Double           ' 0-based array
    Dim TempArrayCount As Long
    
On Error GoTo FindStatisticalMomentsErrorHandler
    
    lngIndexLeftReturn = mIndexLeftInitial
    lngIndexRightReturn = mIndexRightInitial
    
    If mIndexRightInitial = 0 And mIndexRightInitial = 0 Then
        AutoFindIterationStartPoints mIndexLeftInitial, mIndexRightInitial
    End If
    
    mCurIndexLeft = mIndexLeftInitial
    mCurIndexRight = mIndexRightInitial
    
    Do
        LastLP = mCurIndexLeft
        LastRP = mCurIndexRight
        
        TempArrayCount = mCurIndexRight - mCurIndexLeft + 1
        ReDim TempArray(0 To mCurIndexRight - mCurIndexLeft)
        For lngIndex = mCurIndexLeft To mCurIndexRight
            TempArray(lngIndex - mCurIndexLeft) = mData(lngIndex)
        Next lngIndex
        
        blnSuccess = CalculateMoments(TempArray(), TempArrayCount, Moments(), mCurIndexLeft, mDeltaX)
    
        If Not blnSuccess Then Exit Do
        
        ' Find the index nearest X to Moments(1) - sngTargetSigmaValue*Sqr(Moments(3))
        CalcLP = Round(Moments(1) - sngTargetSigmaValue * Sqr(Moments(3)), 0)
        CalcRP = Round(Moments(1) + sngTargetSigmaValue * Sqr(Moments(3)), 0)
    
        If mCurIndexLeft < CalcLP Then
            mCurIndexLeft = mCurIndexLeft + 1
        Else
            If mCurIndexLeft > CalcLP Then mCurIndexLeft = mCurIndexLeft - 1
        End If
        
        If mCurIndexRight < CalcRP Then
            mCurIndexRight = mCurIndexRight + 1
        Else
            If mCurIndexRight > CalcRP Then mCurIndexRight = mCurIndexRight - 1
        End If
        
        If mCurIndexLeft = LastLP And mCurIndexRight = LastRP Then Exit Do
    
        lngIterationCount = lngIterationCount + 1
    Loop While lngIterationCount <= lngMaximumIterations

    If blnSuccess Then
        ' Find some statistics
        With mPeakStats
            .MaximumIntensity = mData(mCurIndexLeft)
            .MaximumIntensityLoc = mCurIndexLeft
            For lngIndex = mCurIndexLeft To mCurIndexRight
                If mData(lngIndex) > .MaximumIntensity Then
                    .MaximumIntensity = mData(lngIndex)
                    .MaximumIntensityLoc = lngIndex
                End If
            Next lngIndex
            
            If Moments(2) <> 0 Then
                .Plates = Moments(1) ^ 2 / Moments(2)
            Else
                .Plates = 0
            End If
            .PeakArea = Moments(0)
            .PeakCenter = Moments(1)
            .Variance = Moments(2)
            .Sigma = Sqr(.Variance)
            .WidthAtFourSigma = 4 * .Sigma
            .Skew = Moments(3) / .Sigma ^ 3
            
            .lngIterationCount = lngIterationCount
        End With
    End If
    
    FindStatisticalMoments = blnSuccess
    Exit Function

FindStatisticalMomentsErrorHandler:
    Debug.Print "Error in FindStatisticalMoments: " & Err.Description
    Debug.Assert False
    FindStatisticalMoments = False
End Function


Private Function CalculateMoments(TempArray() As Double, TempArrayCount As Long, Moments() As Double, lngStartX As Long, Optional dblDeltaX As Double = 1) As Boolean
    ' TempArray() is 0-based
    ' The values in TempArray() are altered by this function, so do not send the original array
    ' Moments() ranges from Moments(0) to Moments(3)
    ' Returns True if success, False if failure

    Dim lngIndex As Long
    Dim ArrayModified() As Double
    
On Error GoTo CalculateMomentsErrorHandler
    
    ' Calculate the zeroth moment (the peak area)
    Moments(0) = FindArea(TempArray(), TempArrayCount, dblDeltaX)
    
    If Moments(0) = 0 Then
        Moments(1) = 0
        Moments(2) = 0
        Moments(3) = 0
        CalculateMoments = True
        Exit Function
    End If
    
    ReDim ArrayModified(0 To TempArrayCount - 1)
    
    ' Take every data point in TempArray() times the point's x value (adjust using lngStartX)
    For lngIndex = 0 To TempArrayCount - 1
        ArrayModified(lngIndex) = TempArray(lngIndex) * (lngIndex + lngStartX)
    Next lngIndex
    
    ' Calculate the first moment (center of mass = X location of peak center)
    Moments(1) = FindArea(ArrayModified(), TempArrayCount, dblDeltaX) / Moments(0)
    
    ' Take every data point in TempArray() times (X - Moments(1))^2
    For lngIndex = 0 To TempArrayCount - 1
        ArrayModified(lngIndex) = TempArray(lngIndex) * ((lngIndex + lngStartX) - Moments(1)) ^ 2
    Next lngIndex
    
    ' Calculate the second moment (variance = sigma^2)
    Moments(2) = FindArea(ArrayModified(), TempArrayCount, dblDeltaX) / Moments(0)
    
    ' Modify TempArray() further
    For lngIndex = 0 To TempArrayCount - 1
        ArrayModified(lngIndex) = ArrayModified(lngIndex) * ((lngIndex + lngStartX) - Moments(1))
    Next lngIndex
    
    ' Finally, calculate the 3rd moment (indicative of skewing)
    Moments(3) = FindArea(ArrayModified(), TempArrayCount, dblDeltaX) / Moments(0)
    
    CalculateMoments = True
    Exit Function

CalculateMomentsErrorHandler:
    Debug.Print "Error in CalculateMoments: " & Err.Description
    Debug.Assert False
    CalculateMoments = False
    
End Function

Private Function FindArea(dblArray() As Double, lngArrayCount As Long, Optional dblDeltaX As Double = 1) As Double
    ' dblArray() should be 0-based
    
    ' Finds the area under the curve, using the trapezoidal integration
    
    Dim lngIndex As Long
    Dim dblArea As Double
    
    dblArea = 0
    For lngIndex = 0 To lngArrayCount - 2
        ' Area of a trapezoid (turned on its side) is:
        '   0.5 * d * (h1 + h2)
        ' where d is the distance between two points, and h1 and h2 are the intensities
        '   at the 2 points
        ' The distance between each of the points in dblArray() is given by dblDeltaX
        
        dblArea = dblArea + 0.5 * dblDeltaX * (dblArray(lngIndex) + dblArray(lngIndex + 1))
    Next lngIndex
    
    FindArea = dblArea
    
End Function

Private Sub Class_Initialize()
    ReDim mData(0)
    mDataCount = 0
    mDeltaX = 1
End Sub

Private Sub Class_Terminate()
    ReDim mData(0)
    mDataCount = 0
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeakDetection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Peak detection routines
' Written by Matthew Monroe in roughly 2001 at UNC (Chapel Hill, NC)
' Kevin Lan provided the concept of Magnitude Concavity fitting
' Ported from LabView code to VB 6 in June 2003 at PNNL (Richland, WA)
' Ported from VB 6 to VB.NET in October 2003
' Synched the VB 6 version with the VB.NET version in February 2005

' Last modified February 20, 2005

Private Enum eTermFunctionConstants
    One = 0
    x = 1
    LogX = 2
    Log10X = 3
    ExpX = 4
    SinX = 5
    CosX = 6
    TanX = 7
    ATanX = 8
End Enum

Private Type udtLeastSquaresFitEquationTermType
    Func As eTermFunctionConstants
    Power As Double
    Coefficient As Double
    Inverse As Boolean
    
    ParamResult As Double        ' Stores the coefficient determined for the fit
End Type

Public Function ComputeSlope(ByRef dblXValsZeroBased() As Double, ByRef dblYValsZeroBased() As Double, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long) As Double
    ' Returns the slope for the given segment of data in dblXVals and dblYVals
    ' Returns 0 if an error
    
    Const POLYNOMIAL_ORDER As Integer = 1

    Dim lngSegmentCount As Long
    Dim lngIndex As Long

    Dim dblSegmentX() As Double
    Dim dblSegmentY() As Double

    Dim dblCoefficients() As Double
    Dim dblSlope As Double
    
On Error GoTo ComputeSlopeErrorHandler
    
    lngSegmentCount = lngEndIndex - lngStartIndex + 1

    If lngSegmentCount > 0 Then
        ReDim dblSegmentX(lngSegmentCount - 1)
        ReDim dblSegmentY(lngSegmentCount - 1)
    
        ' Copy the desired segment of data from dblXVals to dblSegmentX and dblYVals to dblSegmentY
        For lngIndex = lngStartIndex To lngEndIndex
            dblSegmentX(lngIndex - lngStartIndex) = dblXValsZeroBased(lngIndex)
            dblSegmentY(lngIndex - lngStartIndex) = dblYValsZeroBased(lngIndex)
        Next lngIndex
    
        ' Compute the coefficients for the curve fit
        LeastSquaresFit dblSegmentX, dblSegmentY, dblCoefficients, POLYNOMIAL_ORDER
        dblSlope = dblCoefficients(1)
    Else
        dblSlope = 0
    End If
    
    ComputeSlope = dblSlope
    Exit Function

ComputeSlopeErrorHandler:
    Debug.Print "Error in ComputeSlope: " & Err.Description
    Debug.Assert False
    ComputeSlope = 0
    
End Function

Public Function DetectPeaks(dblXValsZeroBased() As Double, dblYValsZeroBased() As Double, lngSourceDataCount As Long, dblIntensityThresholdAbsoluteMinimum As Double, lngPeakWidthPointsMinimum As Long, ByRef lngPeakLocations() As Long, ByRef lngPeakEdgesLeft() As Long, ByRef lngPeakEdgesRight() As Long, ByRef dblPeakAreas() As Double, Optional lngPeakDetectIntensityThresholdPercentageOfMaximum As Long = 0, Optional lngPeakWidthInSigma As Long = 4, Optional blnUseValleysForPeakWidth As Boolean = True, Optional blnMovePeakLocationToMaxIntensity As Boolean = True) As Long
    ' Finds peaks in the parallel arrays dblXValsZeroBased() and dblYValsZeroBased()
    ' dblIntensityThreshold is the minimum absolute intensity allowable for a peak
    ' lngPeakDetectIntensityThresholdPercentageOfMaximum allows one to specify a minimum intensity as a percentage of the maximum peak intensity
    ' Note that the maximum value of dblIntensityThreshold vs. MaxValue*lngPeakDetectIntensityThresholdPercentageOfMaximum is used as the minimum
    ' For example, if dblIntensityThreshold = 10 and lngPeakDetectIntensityThresholdPercentageOfMaximum =  5 (indicating 5%),
    '   then if the maximum of dblYValsZeroBased() is 50, then the minimum intensity of identified peaks is 10, and not 2.5
    '   However, if the maximum of dblYValsZeroBased() is 500, then the minimum intensity of identified peaks is 50, and not 10
    
    ' Returns the locations of the peaks in lngPeakLocations() -- indices of the peak apexes in the source arrays
    ' Returns the left edges of the peaks (in points, not actual units) in intPeakEdgesLeft()       -- These values could be negative if blnUseValleysForPeakWidth = False
    ' Returns the right edges of the peaks in intPeakEdgesRight()                                   -- These values could be larger than intSourceDataCount-1 if blnUseValleysForPeakWidth = False
    ' Returns the areas of the peaks in dblPeakAreas()
    
    ' Note: Compute peak width using: lngPeakWidthPoints = lngPeakEdgesRight(lngPeakLocationsCount) - lngPeakEdgesLeft(lngPeakLocationsCount) + 1

    ' The function returns the number of peaks found; if none are found, returns 0
    
    ' Uses the Magnitude-Concavity method, wherein a second order
    '   polynomial is fit to the points within the window, giving a_2*x^2 + a_1*x + a_0
    '   Given this, a_1 is the first derivative and a_2 is the second derivative
    ' From this, the first derivative gives the index of the peak apex
    ' The standard deviation (s) can be found using:
    '   s = sqrt(-h(t_r) / h''(t_r))
    '  where h(t_r) is the height of the peak at the peak center
    '  and h''(t_r) is the height of the second derivative of the peak
    ' In chromatography, the baseline peak WidthInPoints = 4*dblSigma
    
    Const PEAK_LOCS_DIM_CHUNK As Integer = 10
    
    Dim lngIndex As Long, lngIndexFirst As Long, lngIndexLast As Long
    Dim lngCompareIndex As Long
    Dim lngPeakHalfWidth As Long
    Dim lngPeakLocationsCount As Long, lngPeakLocationsCountDimmed As Long
    
    Dim lngLowIntensityPointCount As Long
    
    Dim lngDataIndexCheck As Long, lngDataIndexCheckStart As Long, lngDataIndexCheckEnd As Long
    
    Dim dblMaximumIntensity As Double, dblIntensityThreshold As Double
    Dim dblSigma As Double
    Dim lngWidthInPoints As Long
    Dim dblFirstDerivative() As Double, dblSecondDerivative() As Double
    
    Dim dblXValsForArea() As Double, dblYValsForArea() As Double
    Dim lngThisPeakWidthInPoints As Long
    Dim lngThisPeakStartIndex As Long, lngThisPeakEndIndex As Long
    Dim lngAreaValsCopyIndex As Long
    
On Error GoTo DetectPeaksErrorHandler
    
    ' Initialize lngPeakLocations()
    lngPeakLocationsCount = 0
    lngPeakLocationsCountDimmed = PEAK_LOCS_DIM_CHUNK
    ReDim lngPeakLocations(lngPeakLocationsCountDimmed)
    ReDim lngPeakEdgesLeft(lngPeakLocationsCountDimmed)
    ReDim lngPeakEdgesRight(lngPeakLocationsCountDimmed)
    ReDim dblPeakAreas(lngPeakLocationsCountDimmed)

    If lngSourceDataCount <= 0 Then
        DetectPeaks = 0
        Exit Function
    End If
    
    ' Reserve space for the first and second derivatives
    ReDim dblFirstDerivative(0 To lngSourceDataCount - 1)
    ReDim dblSecondDerivative(0 To lngSourceDataCount - 1)
    
    ' The mid point width is the minimum width divided by 2, rounded down
    lngPeakHalfWidth = Int(lngPeakWidthPointsMinimum / 2#)
    
    ' Find the maximum intensity in the source data
    dblMaximumIntensity = 0
    For lngIndex = 0 To lngSourceDataCount - 1
        If dblYValsZeroBased(lngIndex) > dblMaximumIntensity Then
            dblMaximumIntensity = dblYValsZeroBased(lngIndex)
        End If
    Next lngIndex
    
    dblIntensityThreshold = dblMaximumIntensity * (lngPeakDetectIntensityThresholdPercentageOfMaximum / 100#)
    If dblIntensityThreshold < dblIntensityThresholdAbsoluteMinimum Then
        dblIntensityThreshold = dblIntensityThresholdAbsoluteMinimum
    End If
    
    ' Exit the function if none of the data is above the minimum threshold
    If dblMaximumIntensity < dblIntensityThreshold Then
        DetectPeaks = 0
        Exit Function
    End If
    
    ' Do the actual work
    FitSegments dblXValsZeroBased(), dblYValsZeroBased(), lngSourceDataCount, lngPeakWidthPointsMinimum, lngPeakHalfWidth, dblFirstDerivative(), dblSecondDerivative()
    
    If lngPeakWidthInSigma < 1 Then lngPeakWidthInSigma = 1
    
    ' Examine the First Derivative function and look for zero crossings (in the downward direction)
    ' If looking for valleys, would look for zero crossings in the upward direction
    ' Only significant if intensity of point is above threshold
    lngPeakLocationsCount = 0
    If lngPeakWidthPointsMinimum <= 0 Then lngPeakWidthPointsMinimum = 1
    
    ' We'll start looking for peaks halfway into lngPeakWidthPointsMinimum
    lngIndexFirst = lngPeakHalfWidth
    lngIndexLast = lngSourceDataCount - 1 - lngPeakHalfWidth

    For lngIndex = lngIndexFirst To lngIndexLast
        If dblFirstDerivative(lngIndex) > 0 And dblFirstDerivative(lngIndex + 1) < 0 Then
            ' Possible peak
            If dblYValsZeroBased(lngIndex) >= dblIntensityThreshold Or dblYValsZeroBased(lngIndex + 1) >= dblIntensityThreshold Then
                ' Actual peak

                lngPeakLocations(lngPeakLocationsCount) = lngIndex

                If blnUseValleysForPeakWidth Then
                    ' Determine the peak width by looking for the adjacent valleys
                    ' If, while looking, we find lngPeakWidthPointsMinimum / 2 points in a row with intensity values below dblIntensityThreshold, then
                    ' set the edge lngPeakHalfWidth - 1 points closer to the peak maximum
            
                    If lngIndex > 0 Then
                        lngPeakEdgesLeft(lngPeakLocationsCount) = 0
                        lngLowIntensityPointCount = 0
                        For lngCompareIndex = lngIndex - 1 To 0 Step -1
                            If dblFirstDerivative(lngCompareIndex) <= 0 And dblFirstDerivative(lngCompareIndex + 1) >= 0 Then
                                ' Found a valley; this is the left edge
                                lngPeakEdgesLeft(lngPeakLocationsCount) = lngCompareIndex + 1
                                Exit For
                            ElseIf dblYValsZeroBased(lngCompareIndex) < dblIntensityThreshold Then
                                lngLowIntensityPointCount = lngLowIntensityPointCount + 1
                                If lngLowIntensityPointCount > lngPeakHalfWidth Then
                                    lngPeakEdgesLeft(lngPeakLocationsCount) = lngCompareIndex + (lngPeakHalfWidth - 1)
                                    Exit For
                                End If
                            Else
                                lngLowIntensityPointCount = 0
                            End If
                        Next lngCompareIndex
                    Else
                        lngPeakEdgesLeft(lngPeakLocationsCount) = 0
                    End If

                    If lngIndex < lngSourceDataCount - 2 Then
                        lngPeakEdgesRight(lngPeakLocationsCount) = lngSourceDataCount - 1
                        lngLowIntensityPointCount = 0
                        For lngCompareIndex = lngIndex + 1 To lngSourceDataCount - 2
                            If dblFirstDerivative(lngCompareIndex) <= 0 And dblFirstDerivative(lngCompareIndex + 1) >= 0 Then
                                ' Found a valley; this is the right edge
                                lngPeakEdgesRight(lngPeakLocationsCount) = lngCompareIndex
                                Exit For
                             ElseIf dblYValsZeroBased(lngCompareIndex) < dblIntensityThreshold Then
                                lngLowIntensityPointCount = lngLowIntensityPointCount + 1
                                If lngLowIntensityPointCount > lngPeakHalfWidth Then
                                    lngPeakEdgesRight(lngPeakLocationsCount) = lngCompareIndex - (lngPeakHalfWidth - 1)
                                    Exit For
                                End If
                            Else
                                lngLowIntensityPointCount = 0
                                End If
                        Next lngCompareIndex
                    Else
                        lngPeakEdgesRight(lngPeakLocationsCount) = lngSourceDataCount - 1
                    End If
                
                    If lngPeakEdgesLeft(lngPeakLocationsCount) > lngPeakLocations(lngPeakLocationsCount) Then
                        Debug.Print "Left edge is > peak center; this is unexpected (clsPeakDetection->DetectPeaks)"
                        Debug.Assert False
                        lngPeakEdgesLeft(lngPeakLocationsCount) = lngPeakLocations(lngPeakLocationsCount)
                    End If

                    If lngPeakEdgesRight(lngPeakLocationsCount) < lngPeakLocations(lngPeakLocationsCount) Then
                        Debug.Print "Right edge is < peak center; this is unexpected (clsPeakDetection->DetectPeaks)"
                        Debug.Assert False
                        lngPeakEdgesRight(lngPeakLocationsCount) = lngPeakLocations(lngPeakLocationsCount)
                    End If
                
                Else
                    ' Examine the Second Derivative to determine peak Width (in points)
                
                    ' If lngPeakWidthPointsMinimum is too small, then the computed width could be too large
                    
                    dblSigma = 0
                    
                    ' If dblSecondDerivative(lngIndex)) is tiny, the following division will fail
                    On Error Resume Next
                    dblSigma = Sqr(Abs(-dblYValsZeroBased(lngIndex) / dblSecondDerivative(lngIndex)))
                    lngWidthInPoints = CLngRoundUp(lngPeakWidthInSigma * dblSigma)
                    On Error GoTo DetectPeaksErrorHandler
                    
                    If lngWidthInPoints > 4 * lngSourceDataCount Then
                        ' Predicted width is over 4 times the data count
                        ' Set it to be 4 times the data count
                        lngWidthInPoints = lngSourceDataCount * 4
                    End If
                    
                    If lngWidthInPoints < 2 Then lngWidthInPoints = 2

                    ' If the peak width is odd, then center around lngIndex
                    ' Otherwise, offset to the right of lngIndex
                    If lngWidthInPoints Mod 2 = 0 Then
                        ' Even number
                        lngPeakEdgesLeft(lngPeakLocationsCount) = lngIndex - CLng(lngWidthInPoints / 2)
                        lngPeakEdgesRight(lngPeakLocationsCount) = lngIndex + CLng(lngWidthInPoints / 2) - 1
                    Else
                        ' Odd number
                        lngPeakEdgesLeft(lngPeakLocationsCount) = lngIndex - CLng((lngWidthInPoints - 1) / 2)
                        lngPeakEdgesRight(lngPeakLocationsCount) = lngIndex + CLng((lngWidthInPoints - 1) / 2)
                    End If
                End If

                lngPeakLocationsCount = lngPeakLocationsCount + 1
                If lngPeakLocationsCount >= lngPeakLocationsCountDimmed Then
                    lngPeakLocationsCountDimmed = lngPeakLocationsCountDimmed + PEAK_LOCS_DIM_CHUNK
                    ReDim Preserve lngPeakLocations(lngPeakLocationsCountDimmed)
                    ReDim Preserve lngPeakEdgesLeft(lngPeakLocationsCountDimmed)
                    ReDim Preserve lngPeakEdgesRight(lngPeakLocationsCountDimmed)
                    ReDim Preserve dblPeakAreas(lngPeakLocationsCountDimmed)
                End If
            End If
        End If
    Next lngIndex

    ' Shrink the arrays to the proper length
    If lngPeakLocationsCount > 0 Then
        ReDim Preserve lngPeakLocations(lngPeakLocationsCount - 1)
        ReDim Preserve lngPeakEdgesLeft(lngPeakLocationsCount - 1)
        ReDim Preserve lngPeakEdgesRight(lngPeakLocationsCount - 1)
        ReDim Preserve dblPeakAreas(lngPeakLocationsCount - 1)
    Else
        ReDim lngPeakLocations(0)
        ReDim lngPeakEdgesLeft(0)
        ReDim lngPeakEdgesRight(0)
        ReDim dblPeakAreas(0)
    End If

    ' Compute the peak areas
    For lngIndex = 0 To lngPeakLocationsCount - 1
        lngThisPeakWidthInPoints = lngPeakEdgesRight(lngIndex) - lngPeakEdgesLeft(lngIndex) + 1
        
        If lngThisPeakWidthInPoints > 0 Then
            If lngThisPeakWidthInPoints = 1 Then
                ' I don't think this can happen
                ' Just in case, we'll set the area equal to the peak intensity
                dblPeakAreas(lngIndex) = dblYValsZeroBased(lngPeakLocations(lngIndex))
            Else
                ReDim dblXValsForArea(lngThisPeakWidthInPoints - 1)
                ReDim dblYValsForArea(lngThisPeakWidthInPoints - 1)
                
                lngThisPeakStartIndex = lngPeakEdgesLeft(lngIndex)
                lngThisPeakEndIndex = lngPeakEdgesRight(lngIndex)
                
                If lngThisPeakStartIndex < 0 Then
                    ' This will happen if the width is too large, or if not all of the peak's data was included in the data arrays
                    lngThisPeakStartIndex = 0
                End If
                
                If lngThisPeakEndIndex >= lngSourceDataCount Then
                    ' This will happen if the width is too large, or if not all of the peak's data was included in the data arrays
                    lngThisPeakEndIndex = lngSourceDataCount - 1
                End If
                
                For lngAreaValsCopyIndex = lngThisPeakStartIndex To lngThisPeakEndIndex
                    dblXValsForArea(lngAreaValsCopyIndex - lngThisPeakStartIndex) = dblXValsZeroBased(lngAreaValsCopyIndex)
                    dblYValsForArea(lngAreaValsCopyIndex - lngThisPeakStartIndex) = dblYValsZeroBased(lngAreaValsCopyIndex)
                Next lngAreaValsCopyIndex
                
                dblPeakAreas(lngIndex) = FindArea(dblXValsForArea(), dblYValsForArea(), lngThisPeakWidthInPoints)

            End If
        Else
            Debug.Print "0-width peak; this shouldn't happen (clsPeakDetection->DetectPeaks)"
            Debug.Assert False
            dblPeakAreas(lngIndex) = 0
        End If
    Next lngIndex
    
    If blnMovePeakLocationToMaxIntensity Then
        For lngIndex = 0 To lngPeakLocationsCount - 1
            ' The peak finder often determines the peak center to be a few points away from the peak apex -- check for this
            ' Define the maximum allowed peak apex shift to be 33% of lngPeakWidthPointsMinimum
            ' Note: Int() is equivalent to VB.NET's Math.Floor()
            lngDataIndexCheckStart = lngPeakLocations(lngIndex) - Int(lngPeakWidthPointsMinimum / 3)
            If lngDataIndexCheckStart < 0 Then lngDataIndexCheckStart = 0
    
            lngDataIndexCheckEnd = lngPeakLocations(lngIndex) + Int(lngPeakWidthPointsMinimum / 3)
            If lngDataIndexCheckEnd > lngSourceDataCount - 1 Then lngDataIndexCheckEnd = lngSourceDataCount - 1
    
            dblMaximumIntensity = dblYValsZeroBased(lngPeakLocations(lngIndex))
            For lngDataIndexCheck = lngDataIndexCheckStart To lngDataIndexCheckEnd
                If dblYValsZeroBased(lngDataIndexCheck) > dblMaximumIntensity Then
                    lngPeakLocations(lngIndex) = lngDataIndexCheck
                    dblMaximumIntensity = dblYValsZeroBased(lngDataIndexCheck)
                End If
            Next lngDataIndexCheck
    
            If lngPeakLocations(lngIndex) < lngPeakEdgesLeft(lngIndex) Then lngPeakEdgesLeft(lngIndex) = lngPeakLocations(lngIndex)
            If lngPeakLocations(lngIndex) > lngPeakEdgesRight(lngIndex) Then lngPeakEdgesRight(lngIndex) = lngPeakLocations(lngIndex)
        Next lngIndex
    End If
    
    DetectPeaks = lngPeakLocationsCount
    Exit Function

DetectPeaksErrorHandler:
    Debug.Print "Error in DetectPeaks: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "clsPeakDetection->DetectPeaks"
    
End Function

Private Function FindArea(dblXVals() As Double, dblYVals() As Double, lngArrayCount As Long) As Double
    ' dblYVals() should be 0-based
    
    ' Finds the area under the curve, using trapezoidal integration
    
    Dim lngIndex As Long
    Dim dblArea As Double
    
    dblArea = 0
    For lngIndex = 0 To lngArrayCount - 2
        ' Area of a trapezoid (turned on its side) is:
        '   0.5 * d * (h1 + h2)
        ' where d is the distance between two points, and h1 and h2 are the intensities
        '   at the 2 points
        
        dblArea = dblArea + 0.5 * Abs(dblXVals(lngIndex + 1) - dblXVals(lngIndex)) * (dblYVals(lngIndex) + dblYVals(lngIndex + 1))
    Next lngIndex
    
    FindArea = dblArea
    
End Function

Private Sub FitSegments(dblXVals() As Double, dblYVals() As Double, lngSourceDataCount As Long, lngPeakWidthPointsMinimum As Long, lngPeakWidthMidPoint As Long, ByRef dblFirstDerivative() As Double, ByRef dblSecondDerivative() As Double)
    ' dblXVals() and dblYVals() are zero-based arrays

    Const POLYNOMIAL_ORDER As Integer = 2

    Dim dblSegmentX() As Double
    Dim dblSegmentY() As Double
    
    Dim dblCoefficients() As Double
    
    Dim lngSubIndex As Long, lngStartIndex As Long
    Dim lngMidPointIndex As Long

    ' If POLYNOMIAL_ORDER < 2 Then POLYNOMIAL_ORDER = 2
    ' If POLYNOMIAL_ORDER > 9 Then POLYNOMIAL_ORDER = 9
    
    ReDim dblSegmentX(0 To lngPeakWidthPointsMinimum - 1)
    ReDim dblSegmentY(0 To lngPeakWidthPointsMinimum - 1)

    For lngStartIndex = 0 To lngSourceDataCount - lngPeakWidthPointsMinimum - 1
        
        ' Copy the desired segment of data from dblXVals to dblSegmentX and dblYVals to dblSegmentY
        For lngSubIndex = lngStartIndex To lngStartIndex + lngPeakWidthPointsMinimum - 1
            dblSegmentX(lngSubIndex - lngStartIndex) = dblXVals(lngSubIndex)
            dblSegmentY(lngSubIndex - lngStartIndex) = dblYVals(lngSubIndex)
        Next lngSubIndex
        
        ' Compute the coefficients for the curve fit
        LeastSquaresFit dblSegmentX(), dblSegmentY(), dblCoefficients(), POLYNOMIAL_ORDER
        
        ' Compute the dblFirstDerivative at the midpoint
        lngMidPointIndex = lngStartIndex + lngPeakWidthMidPoint
        dblFirstDerivative(lngMidPointIndex) = 2 * dblCoefficients(2) * dblXVals(lngMidPointIndex) + dblCoefficients(1)
        dblSecondDerivative(lngMidPointIndex) = 2 * dblCoefficients(2)

    Next lngStartIndex

End Sub

Private Function LeastSquaresFit(ByRef dblXVals() As Double, ByRef dblYVals() As Double, ByRef dblCoefficients() As Double, ByVal intPolynomialOrder As Integer) As Boolean

    ' Code from article "Fit for Purpose" written by Steven Abbot
    ' and published in the February 2003 issue of Hardcore Visual Basic.
    ' Code excerpted from the VB6 program FitIt
    ' URL: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnhcvb03/html/hcvb03b1.asp

    Dim udtEquationTerms() As udtLeastSquaresFitEquationTermType
    Dim intTerm As Integer
    Dim blnSuccess As Boolean

    ReDim udtEquationTerms(intPolynomialOrder)
    ReDim dblCoefficients(intPolynomialOrder)

    If UBound(dblXVals) + 1 < intPolynomialOrder + 1 Then
        ' Not enough data to fit a curve
        blnSuccess = False
    Else

        ' Define equation for "ax^0 + bx^1 + cx^2", which is the same as "a + bx + cx^2"
        For intTerm = 0 To intPolynomialOrder
            With udtEquationTerms(intTerm)
                .Coefficient = 1                        ' a, b, c in the above equation
                .Func = eTermFunctionConstants.x        ' X
                .Power = intTerm                        ' Power that X is raised to
                .Inverse = False                        ' Whether or not to inverse the entire term

                .ParamResult = 0
            End With
        Next intTerm

        blnSuccess = LLSqFit(dblXVals, dblYVals, udtEquationTerms)
        For intTerm = 0 To intPolynomialOrder
            dblCoefficients(intTerm) = udtEquationTerms(intTerm).ParamResult
        Next
    End If

    LeastSquaresFit = blnSuccess

End Function

Private Function LLSqFit(ByRef DataX() As Double, ByRef DataY() As Double, ByRef udtEquationTerms() As udtLeastSquaresFitEquationTermType) As Boolean

    'Linear Least Squares Fit

    Dim intEquationTermCount As Integer
    Dim i As Long, j As Long, k As Long, L As Long, m As Long
    Dim ym As Double

    Dim beta() As Double
    Dim CoVar() As Double           ' 2D array, CoVar(,)
    Dim PFuncVal() As Double
    
    Dim blnSuccess As Boolean

    intEquationTermCount = UBound(udtEquationTerms) + 1
    
    ReDim beta(UBound(DataX))
    ReDim CoVar(intEquationTermCount - 1, intEquationTermCount - 1)
    ReDim PFuncVal(intEquationTermCount - 1)

    For i = 0 To UBound(DataX)
        GetLVals DataX(i), udtEquationTerms, PFuncVal
        ym = DataY(i)
        For L = 0 To intEquationTermCount - 1
            For m = 0 To L
                CoVar(L, m) = CoVar(L, m) + PFuncVal(L) * PFuncVal(m)
            Next m
            beta(L) = beta(L) + ym * PFuncVal(L)
        Next L
    Next i
    For j = 1 To intEquationTermCount - 1
        For k = 0 To j - 1
            CoVar(k, j) = CoVar(j, k)
        Next k
    Next j

    If GaussJordan(CoVar, udtEquationTerms, beta) Then
        For L = 0 To intEquationTermCount - 1
            udtEquationTerms(L).ParamResult = beta(L)
        Next L

        blnSuccess = True
    Else
        ' Error fitting; clear dblCoefficients
        For L = 0 To intEquationTermCount - 1
            udtEquationTerms(L).ParamResult = 0
        Next L

        blnSuccess = False
    End If
    
    LLSqFit = blnSuccess

End Function

Private Sub GetLVals(ByVal x As Double, ByRef udtEquationTerms() As udtLeastSquaresFitEquationTermType, ByRef PFuncVal() As Double)
    ' Get values for Linear Least Squares
    ' udtEquationTerms() is a 0-based array defining the form of each term

    Dim i As Long, v As Double

    ' Use the following for a 2nd order polynomial fit
    ''Define the formula via PFuncVal
    ''In this case NTerms=3 and y=a+bx+cx^2
    'PFuncVal(1) = 1
    'PFuncVal(2) = X
    'PFuncVal(3) = X ^ 2

    'f = "1,X,Log(X),Log10(X),Exp(X),Sin(X),Cos(X),Tan(X),ATAN(X)"
    For i = 0 To UBound(udtEquationTerms)
        With udtEquationTerms(i)
            Select Case .Func
                Case eTermFunctionConstants.One
                    v = 1
                Case eTermFunctionConstants.x
                    v = x ^ .Power
                Case eTermFunctionConstants.LogX
                    If .Coefficient * x <= 0 Then
                        v = 0
                    Else
                        v = Log(.Coefficient * x) ^ .Power
                    End If
                Case eTermFunctionConstants.Log10X
                    If .Coefficient * x <= 0 Then
                        v = 0
                    Else
                        v = Log10(.Coefficient * x) ^ .Power
                    End If
                Case eTermFunctionConstants.ExpX
                    v = Exp(.Coefficient * x) ^ .Power
                Case eTermFunctionConstants.SinX
                    v = Sin(.Coefficient * x) ^ .Power
                Case eTermFunctionConstants.CosX
                    v = Cos(.Coefficient * x) ^ .Power
                Case eTermFunctionConstants.TanX
                    v = Tan(.Coefficient * x) ^ .Power
                Case eTermFunctionConstants.ATanX
                    v = Atn(.Coefficient * x) ^ .Power
            End Select

            If .Inverse Then
                If v = 0 Then
                    PFuncVal(i) = 0
                Else 'NOT V...
                    PFuncVal(i) = 1 / v
                End If
            Else 'INV(I) = FALSE
                PFuncVal(i) = v
            End If

        End With
    Next i

End Sub

Private Function GaussJordan(ByRef A() As Double, ByRef udtEquationTerms() As udtLeastSquaresFitEquationTermType, ByRef B() As Double) As Boolean

    ' GaussJordan elimination for LLSq and LM solving
    ' Returns True if success, False if an error
    ' Note that A() is a 2D array, A(,)
    
    Dim indxc() As Long, indxr() As Long, ipiv() As Long
    Dim n As Long

    n = UBound(udtEquationTerms) + 1

    ReDim indxc(n - 1)
    ReDim indxr(n - 1)
    ReDim ipiv(n - 1)

    Dim i As Long, icol As Long, irow As Long, j As Long, k As Long, L As Long, ll As Long
    Dim Big As Double, Dum As Double, PivInv As Double

On Error GoTo GaussJordanErrorHandler

    For i = 0 To n - 1
        Big = 0
        For j = 0 To n - 1
            If ipiv(j) <> 1 Then
                For k = 0 To n - 1
                    If ipiv(k) = 0 Then
                        If Abs(A(j, k)) >= Big Then
                            Big = Abs(A(j, k))
                            irow = j
                            icol = k
                        End If
                    End If
                Next k
            End If
        Next j

        ipiv(icol) = ipiv(icol) + 1
        If irow <> icol Then
            For L = 0 To n - 1
                Dum = A(irow, L)
                A(irow, L) = A(icol, L)
                A(icol, L) = Dum
            Next L
            Dum = B(irow)
            B(irow) = B(icol)
            B(icol) = Dum
        End If

        indxr(i) = irow
        indxc(i) = icol
        If A(icol, icol) = 0 Then
            ' Error, the matrix was singular
            GaussJordan = False
            Exit Function
        End If

        PivInv = 1 / A(icol, icol)
        A(icol, icol) = 1
        For L = 0 To n - 1
            A(icol, L) = A(icol, L) * PivInv
        Next L

        B(icol) = B(icol) * PivInv
        For ll = 0 To n - 1
            If ll <> icol Then
                Dum = A(ll, icol)
                A(ll, icol) = 0
                For L = 0 To n - 1
                    A(ll, L) = A(ll, L) - A(icol, L) * Dum
                Next L
                B(ll) = B(ll) - B(icol) * Dum
            End If
        Next ll
    Next i

    For L = n - 1 To 0 Step -1
        If indxr(L) <> indxc(L) Then
            For k = 0 To n - 1
                Dum = A(k, indxr(L))
                A(k, indxr(L)) = A(k, indxc(L))
                A(k, indxc(L)) = Dum
            Next k
        End If
    Next L

    GaussJordan = True
    Exit Function

GaussJordanErrorHandler:
    Debug.Print "Error in GaussJordan: " & Err.Description
    Debug.Assert False
    
    GaussJordan = False
    
End Function


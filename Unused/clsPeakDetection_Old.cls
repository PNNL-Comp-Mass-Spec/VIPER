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

' Last modified December 29, 2003

Private Const EODOUBLE As Integer = 1
Private Const EOFLOAT As Integer = 2
Private Const EOLONG As Integer = 3
Private Const EOSHORT As Integer = 4

Private Declare Function EoLeastSquaresFit Lib "eols.dll" _
    (XData As Any, _
    YData As Any, _
    ByVal iNDataPoints As Long, _
    ByVal iNCoefficients As Long, _
    ByVal fnUserEquation As Long, _
    Coef As Any, _
    ByVal iDataType As Long, _
    ByVal iSaveStateFlag As Long, _
    ByRef handle As Long) As Long

''The following is only needed when using eolsrt.dll; used for real-time least squares fitting, utilizing data buffering
''Private Declare Sub EoLeastSquaresFitClose Lib "eolsrt.dll" (ByRef handle As Long)

' This structure was used by code that is now commented out
'    Private Structure udtFitCoefficientsType
'        Public B0 As Double
'        Public B1 As Double
'        Public B2 As Double
'    End Structure

Private mEolsDllNotFound As Boolean

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
    ' Note: Compute peak width using: lngPeakWidthsPoints = lngPeakEdgesRight(lngPeakLocationsCount) - lngPeakEdgesLeft(lngPeakLocationsCount) + 1

    ' Returns the areas of the peaks in dblPeakAreas()

    ' The function returns the number of peaks found; if none are found, returns 0
    ' If eols.dll cannot be found, returns -1
    
    ' Use the Magnitude-Concavity method, wherein a second order
    '   polynomial is fit to the points within the window, giving a_2*x^2 + a_1*x + a_0
    '   Given this, a_1 is the first derivative and a_2 is the second derivative
    ' From this, the first derivative gives the index of the peak apex
    ' The standard deviation (s) can be found using:
    '   s = sqrt(-h(t_r) / h''(t_r))
    '  where h(t_r) is the height of the peak at the peak center
    '  and h''(t_r) is the height of the second derivative of the peak
    ' In chromatography, the baseline peak dblWidthInPoints = 4*dblSigma
    
    Const PEAK_LOCS_DIM_CHUNK As Integer = 10
    
    Dim lngIndex As Long, lngIndexFirst As Long, lngIndexLast As Long
    Dim lngCompareIndex As Long
    Dim lngPeakWidthMidPoint As Long
    Dim lngPeakLocationsCount As Long, lngPeakLocationsCountDimmed As Long
    Dim lngDataIndexCheck As Long, lngDataIndexCheckStart As Long, lngDataIndexCheckEnd As Long
    
    Dim dblMaximumIntensity As Double, dblIntensityThreshold As Double
    Dim dblSigma As Double, dblWidthInPoints As Double
    Dim lngWidthInPoints As Long
    Dim dblFirstDerivative() As Double, dblSecondDerivative() As Double
    
    Dim dblXValsForArea() As Double, dblYValsForArea() As Double
    Dim lngThisPeakWidthInPoints As Long
    Dim lngThisPeakStartIndex As Long, lngThisPeakEndIndex As Long
    Dim lngAreaValsCopyIndex As Long
    
    ' Initialize lngPeakLocations()
    lngPeakLocationsCount = 0
    lngPeakLocationsCountDimmed = PEAK_LOCS_DIM_CHUNK
    ReDim lngPeakLocations(lngPeakLocationsCountDimmed)
    ReDim lngPeakEdgesLeft(lngPeakLocationsCountDimmed)
    ReDim lngPeakEdgesRight(lngPeakLocationsCountDimmed)
    ReDim dblPeakAreas(lngPeakLocationsCountDimmed)
    
    On Error Resume Next
    
    ' The following two statements are necessary to ensure that the current directory
    '  is the application path, which should contain eols.dll
    ChDir App.Path
    ChDrive App.Path
    Debug.Assert CurDir() = App.Path
    
On Error GoTo DetectPeaksErrorHandler

    If lngSourceDataCount <= 0 Then
        DetectPeaks = 0
        Exit Function
    End If
    
    ' Reserve space for the first and second derivatives
    ReDim dblFirstDerivative(0 To lngSourceDataCount - 1)
    ReDim dblSecondDerivative(0 To lngSourceDataCount - 1)
    
    ' The mid point width is the minimum width divided by 2, rounded down
    lngPeakWidthMidPoint = Int(lngPeakWidthPointsMinimum / 2#)
    
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
    FitSegments dblXValsZeroBased(), dblYValsZeroBased(), lngSourceDataCount, lngPeakWidthPointsMinimum, lngPeakWidthMidPoint, dblFirstDerivative(), dblSecondDerivative()
    
    If mEolsDllNotFound Then
        DetectPeaks = -1
        Exit Function
    End If
    
    If lngPeakWidthInSigma < 1 Then lngPeakWidthInSigma = 1
    
    ' Examine the First Derivative function and look for zero crossings (in the downward direction)
    ' If looking for valleys, would look for zero crossings in the upward direction
    ' Only significant if intensity of point is above threshold
    lngPeakLocationsCount = 0
    If lngPeakWidthPointsMinimum <= 0 Then lngPeakWidthPointsMinimum = 1
    
    ' We'll start looking for peaks halfway into lngPeakWidthPointsMinimum
    lngIndexFirst = lngPeakWidthMidPoint
    lngIndexLast = lngSourceDataCount - 1 - lngPeakWidthMidPoint

    For lngIndex = lngIndexFirst To lngIndexLast
        If dblFirstDerivative(lngIndex) > 0 And dblFirstDerivative(lngIndex + 1) < 0 Then
            ' Possible peak
            If dblYValsZeroBased(lngIndex) >= dblIntensityThreshold Or dblYValsZeroBased(lngIndex + 1) >= dblIntensityThreshold Then
                ' Actual peak

                lngPeakLocations(lngPeakLocationsCount) = lngIndex

                If blnUseValleysForPeakWidth Then
                    ' Determine the peak width by looking for the adjacent valleys
                    If lngIndex > 0 Then
                        lngPeakEdgesLeft(lngPeakLocationsCount) = 0
                        For lngCompareIndex = lngIndex - 1 To 0 Step -1
                            If dblFirstDerivative(lngCompareIndex) <= 0 And dblFirstDerivative(lngCompareIndex + 1) >= 0 Then
                                ' Found a valley; this is the left edge
                                lngPeakEdgesLeft(lngPeakLocationsCount) = lngCompareIndex + 1
                                Exit For
                            End If
                        Next lngCompareIndex
                    Else
                        lngPeakEdgesLeft(lngPeakLocationsCount) = 0
                    End If

                    If lngIndex < lngSourceDataCount - 2 Then
                        lngPeakEdgesRight(lngPeakLocationsCount) = lngSourceDataCount - 1
                        For lngCompareIndex = lngIndex + 1 To lngSourceDataCount - 2
                            If dblFirstDerivative(lngCompareIndex) <= 0 And dblFirstDerivative(lngCompareIndex + 1) >= 0 Then
                                ' Found a valley; this is the right edge
                                lngPeakEdgesRight(lngPeakLocationsCount) = lngCompareIndex
                                Exit For
                            End If
                        Next lngCompareIndex
                    Else
                        lngPeakEdgesRight(lngPeakLocationsCount) = lngSourceDataCount - 1
                    End If
                Else
                    ' Examine the Second Derivative to determine peak Width (in points)
                
                    ' If lngPeakWidthPointsMinimum is too small, then the computed width could be too large
                    
                    dblSigma = 0
                    
                    ' If dblSecondDerivative(lngIndex)) is tiny, the following division will fail
                    On Error Resume Next
                    dblSigma = Sqr(Abs(-dblYValsZeroBased(lngIndex) / dblSecondDerivative(lngIndex)))
                    dblWidthInPoints = lngPeakWidthInSigma * dblSigma
                    On Error GoTo DetectPeaksErrorHandler
                    
                    If dblWidthInPoints = 0 Then dblWidthInPoints = 1
                    
                    If dblWidthInPoints > 4 * lngSourceDataCount Then
                        ' Predicted width is over 4 times the data count
                        ' Set it to be 4 times the data count
                        dblWidthInPoints = lngSourceDataCount * 4
                    End If
                    
                    If dblWidthInPoints < 2 Then dblWidthInPoints = 2
                    lngWidthInPoints = CLngRoundUp(dblWidthInPoints)

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
            ' 0-width peak; this shouldn't happen
            Debug.Assert False
            dblPeakAreas(lngIndex) = 0
        End If
    Next lngIndex
    
    If blnMovePeakLocationToMaxIntensity Then
        For lngIndex = 0 To lngPeakLocationsCount - 1
            ' The peak finder often determines the peak center to be a few points away from the peak apex -- check for this
            ' Define the maximum allowed peak apex shift to be 33% of lngPeakWidthPointsMinimum
            lngDataIndexCheckStart = lngPeakLocations(lngIndex) - lngPeakWidthPointsMinimum / 3
            If lngDataIndexCheckStart < 0 Then lngDataIndexCheckStart = 0
    
            lngDataIndexCheckEnd = lngPeakLocations(lngIndex) + lngPeakWidthPointsMinimum / 3
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


Private Sub FitSegments(dblXVals() As Double, dblYVals() As Double, lngSourceDataCount As Long, lngPeakWidthPointsMinimum As Long, lngPeakWidthMidPoint As Long, ByRef dblFirstDerivative() As Double, ByRef dblSecondDerivative() As Double, Optional ByVal lngPolynomialOrder As Long = 2)
    ' dblXVals() and dblYVals() are zero-based arrays

    Dim dblSegmentX() As Double
    Dim dblSegmentY() As Double
    
    Dim dblCoefficients() As Double
    
    Dim lngSubIndex As Long, lngStartIndex As Long
    Dim lngMidPointIndex As Long

    If lngPolynomialOrder < 2 Then lngPolynomialOrder = 2
    If lngPolynomialOrder > 9 Then lngPolynomialOrder = 9
    
    ' Note: For a 2nd order equation, dblCoefficients(0), (1), and (2) correspond to C0, C1, and C2 in the equation:
    '       y = C0 +  C1 x  +  C2 x^2
    ReDim dblCoefficients(lngPolynomialOrder + 1)

    ReDim dblSegmentX(0 To lngPeakWidthPointsMinimum - 1)
    ReDim dblSegmentY(0 To lngPeakWidthPointsMinimum - 1)

    For lngStartIndex = 0 To lngSourceDataCount - lngPeakWidthPointsMinimum - 1
        
        ' Copy the desired segment of data from dblXVals to dblSegmentX and dblYVals to dblSegmentY
        For lngSubIndex = lngStartIndex To lngStartIndex + lngPeakWidthPointsMinimum - 1
            dblSegmentX(lngSubIndex - lngStartIndex) = dblXVals(lngSubIndex)
            dblSegmentY(lngSubIndex - lngStartIndex) = dblYVals(lngSubIndex)
        Next lngSubIndex
        
        ' Compute the coefficients for the curve fit
        LeastSquaresFit dblSegmentX(), dblSegmentY(), lngPeakWidthPointsMinimum, dblCoefficients(), lngPolynomialOrder
        
        ' Compute the dblFirstDerivative at the midpoint
        lngMidPointIndex = lngStartIndex + lngPeakWidthMidPoint
        dblFirstDerivative(lngMidPointIndex) = 2 * dblCoefficients(2) * dblXVals(lngMidPointIndex) + dblCoefficients(1)
        dblSecondDerivative(lngMidPointIndex) = 2 * dblCoefficients(2)

        If mEolsDllNotFound Then Exit For
    Next lngStartIndex

End Sub

Private Sub LeastSquaresFit(dblXVals() As Double, dblYVals() As Double, lngDataCount As Long, dblCoefficients() As Double, lngPolynomialOrder As Long)
    ' Uses the EoLeastSquaresFit function in the eols.dll file to compute a least squares fit on the portion of the data between lngIndexStart and lngIndexEnd
    ' lngPolynomialOrder should be between 2 and 9
    ' dblXVals() should range from 0 to lngDataCount-1
    
    Dim lngReturnCode As Long
    
On Error GoTo LeastSquaresFitErrorHandler

    ' Note: For a 2nd order equation, dblCoefficients(0), (1), and (2) correspond to C0, C1, and C2 in the equation:
    '       y = C0 +  C1 x  +  C2 x^2
    lngReturnCode = EoLeastSquaresFit(dblXVals(0), dblYVals(0), lngDataCount, lngPolynomialOrder + 1, 0, dblCoefficients(0), EODOUBLE, 0, 0)
    Debug.Assert lngReturnCode = 1
    
    Exit Sub
    
LeastSquaresFitErrorHandler:
    If Err.Number = 53 Then
        If Not mEolsDllNotFound Then
            Debug.Print "Count not find the eols.dll file"
            mEolsDllNotFound = True
            Debug.Assert False
        
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                ' Inform the user that we couldn't find eols.dll
                MsgBox "Could not find the eols.dll file; please assure it is located in the application folder and in \Windows\System32\"
            Else
                LogErrors Err.Number, "clsPeakDetection.LeastSquaresFit, Could not find eols.dll"
            End If
        End If
    Else
        Debug.Print "Error in LeastSquaresFit: " & Err.Description
        Debug.Assert False
        LogErrors Err.Number, "frmTICAndBPIPlots->LeastSquaresFit"
    End If
    
End Sub


' The following function is supposed to fit a 2nd order equation to dblXVals() and dblYVals() using a least squares fit
' However, it doesn't work
' Thus, we're using a call to EoLeastSquaresFit (eols.dll) in the previous sub, LeastSquaresFit
'Private Sub CurveFitCoeff(dblXVals() As Double, dblYVals() As Double, lngDataCount As Long, dblCoefficients() As Double)
'    ' Returns coefficients for fitting a 2nd order equation to dblXVals() and dblYVals()
'
'    ' Note: dblCoefficients(0), (1), and (2) correspond to C0, C1, and C2 in the equation:
'    '       y = C0 +  C1 x  +  C2 x^2
'
'    Dim SX As Double, sx2 As Double, sx3 As Double, sx4 As Double
'    Dim SY As Double
'    Dim sxy As Double, sx2y As Double
'    Dim Denominator As Double
'
'    SX = SumXpower(dblXVals(), 0, lngDataCount - 1, 1)
'    sx2 = SumXpower(dblXVals(), 0, lngDataCount - 1, 2)
'    sx3 = SumXpower(dblXVals(), 0, lngDataCount - 1, 3)
'    sx4 = SumXpower(dblXVals(), 0, lngDataCount - 1, 4)
'    SY = SumY(dblYVals(), 0, lngDataCount - 1)
'    sxy = SumXpowerTimesY(dblXVals(), dblYVals(), 0, lngDataCount - 1, 1)
'    sx2y = SumXpowerTimesY(dblXVals(), dblYVals(), 0, lngDataCount - 1, 2)
'
'    ReDim dblCoefficients(0 To 2)
'
'    Denominator = (SX ^ 2 * sx4 - 2 * SX * sx2 * sx3 + sx2 ^ 3 - sx2 * sx4 * lngDataCount + sx3 ^ 2 * lngDataCount)
'    If Denominator > 0 Then
'        dblCoefficients(0) = -(SX * (sx3 * sx2y - sxy * sx4) - sx2 ^ 2 * sx2y + sx2 * (SY * sx4 + sx3 * sxy) - SY * sx3 ^ 2) / Denominator
'        dblCoefficients(1) = -(SX * (sx2 * sx2y - SY * sx4) - sx2 ^ 2 * sxy + sx2 * SY * sx3 - lngDataCount * (sx3 * sx2y - sxy * sx4)) / Denominator
'        dblCoefficients(2) = (SX ^ 2 * sx2y - SX * (sx2 * sxy + SY * sx3) + sx2 ^ 2 * SY - sx2 * sx2y * lngDataCount + sx3 * sxy * lngDataCount) / Denominator
'    Else
'        dblCoefficients(0) = 0
'        dblCoefficients(1) = 0
'        dblCoefficients(2) = 0
'    End If
'End Sub
'
'Private Function SumXpower(dblXVals() As Double, lngIndexStart As Long, lngIndexEnd As Long, Optional Power As Integer = 1) As Double
'    ' Don't allow power = 0 since 0^0 = error
'
'    Dim lngIndex As Long, dblSum As Double
'
'    If Power < 1 Then Power = 1
'
'    dblSum = 0
'    For lngIndex = lngIndexStart To lngIndexEnd
'        dblSum = dblSum + dblXVals(lngIndex) ^ Power
'    Next lngIndex
'
'    SumXpower = dblSum
'End Function
'
'Private Function SumXpowerTimesY(dblXVals() As Double, dblYVals() As Double, lngIndexStart As Long, lngIndexEnd As Long, Optional Power As Integer = 1) As Double
'    ' Don't allow power = 0 since 0^0 = error
'
'    Dim lngIndex As Long, dblSum As Double
'
'    If Power < 1 Then Power = 1
'
'    dblSum = 0
'    For lngIndex = lngIndexStart To lngIndexEnd
'        dblSum = dblSum + (dblXVals(lngIndex) ^ Power) * dblYVals(lngIndex)
'    Next lngIndex
'
'    SumXpowerTimesY = dblSum
'End Function
'
'Private Function SumY(dblYVals() As Double, lngIndexStart As Long, lngIndexEnd As Long) As Double
'    Dim lngIndex As Long, dblSum As Double
'
'    dblSum = 0
'    For lngIndex = lngIndexStart To lngIndexEnd
'        dblSum = dblSum + dblYVals(lngIndex)
'    Next lngIndex
'
'    SumY = dblSum
'End Function




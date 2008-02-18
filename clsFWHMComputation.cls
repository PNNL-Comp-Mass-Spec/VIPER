VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFWHMComputation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''Public Function ComputeFWHM(ByRef lngScanNumbers() As Long, ByRef sngIntensityList() As Single, ByRef udtSICPeak As udtSICStatsPeakType, ByVal blnSubtractBaselineNoise As Boolean) As Long
''    ' Note: The calling function should have already populated udtSICPeak.MaxIntensityValue, plus .IndexMax, .IndexBaseLeft, and .IndexBaseRight
''    ' If blnSubtractBaselineNoise is True, then this function also uses udtSICPeak.BaselineNoiseStats....
''    ' Note: This function returns the FWHM value in units of scan number; it does not update the value stored in udtSICPeak
''    ' This function does, however, update udtSICPeak.IndexMax if it is not between udtSICPeak.IndexBaseLeft and udtSICPeak.IndexBaseRight
''
''    Const ALLOW_NEGATIVE_VALUES As Boolean = False
''    Dim sngFWHMScanStart, sngFWHMScanEnd As Single
''    Dim lngFWHMScans As Long
''
''    Dim lngDataIndex As Long
''    Dim sngTargetIntensity As Single
''    Dim sngMaximumIntensity As Single
''
''    Dim sngY1, sngY2 As Single
''
''    On Error GoTo ComputeFWHMErrorHandler
''
''    ' Determine the full width at half max (FWHM), in units of absolute scan number
''
''    With udtSICPeak
''
''        If .IndexMax <= .IndexBaseLeft Or .IndexMax >= .IndexBaseRight Then
''            ' Find the index of the maximum (between .IndexBaseLeft and .IndexBaseRight)
''            sngMaximumIntensity = 0
''            If .IndexMax < .IndexBaseLeft Or .IndexMax > .IndexBaseRight Then
''                .IndexMax = .IndexBaseLeft
''            End If
''
''            For lngDataIndex = .IndexBaseLeft To .IndexBaseRight
''                If sngIntensityList(lngDataIndex) > sngMaximumIntensity Then
''                    .IndexMax = lngDataIndex
''                    sngMaximumIntensity = sngIntensityList(lngDataIndex)
''                End If
''            Next lngDataIndex
''        End If
''
''        ' Look for the intensity halfway down the peak (correcting for baseline noise level if blnSubtractBaselineNoise = True)
''        If blnSubtractBaselineNoise Then
''            sngTargetIntensity = BaselineAdjustIntensity(.MaxIntensityValue, .BaselineNoiseStats.NoiseLevel, ALLOW_NEGATIVE_VALUES) / 2
''
''            If sngTargetIntensity <= 0 Then
''                ' The maximum intensity of the peak is below the baseline; do not correct for baseline noise level
''                sngTargetIntensity = .MaxIntensityValue / 2
''                blnSubtractBaselineNoise = False
''            End If
''        Else
''            sngTargetIntensity = .MaxIntensityValue / 2
''        End If
''
''        If sngTargetIntensity > 0 Then
''
''            ' Start the search at each peak edge to thus determine the largest FWHM value
''            sngFWHMScanStart = -1
''            For lngDataIndex = .IndexBaseLeft To .IndexMax - 1
''                If blnSubtractBaselineNoise Then
''                    sngY1 = BaselineAdjustIntensity(sngIntensityList(lngDataIndex), .BaselineNoiseStats.NoiseLevel, ALLOW_NEGATIVE_VALUES)
''                    sngY2 = BaselineAdjustIntensity(sngIntensityList(lngDataIndex + 1), .BaselineNoiseStats.NoiseLevel, ALLOW_NEGATIVE_VALUES)
''                Else
''                    sngY1 = sngIntensityList(lngDataIndex)
''                    sngY2 = sngIntensityList(lngDataIndex + 1)
''                End If
''
''                If sngY1 > sngTargetIntensity Or sngY2 > sngTargetIntensity Then
''                    If sngY1 <= sngTargetIntensity And sngY2 >= sngTargetIntensity Then
''                        blnSuccess = InterpolateX(sngFWHMScanStart, lngScanNumbers(lngDataIndex), lngScanNumbers(lngDataIndex + 1), sngY1, sngY2, sngTargetIntensity)
''                    Else
''                        ' sngTargetIntensity is not between sngY1 and sngY2; simply use lngDataIndex
''                        If lngDataIndex = .IndexBaseLeft Then
''                            ' At the start of the peak; use the scan number halfway between .IndexBaseLeft and .IndexMax
''                            sngFWHMScanStart = lngScanNumbers(lngDataIndex + CInt(Math.Round((.IndexMax - .IndexBaseLeft) / 2, 0)))
''                        Else
''                            ' This code will probably never be reached
''                            sngFWHMScanStart = lngScanNumbers(lngDataIndex)
''                        End If
''                    End If
''                    Exit For
''                End If
''            Next lngDataIndex
''            If sngFWHMScanStart < 0 Then
''                If .IndexMax > .IndexBaseLeft Then
''                    sngFWHMScanStart = lngScanNumbers(.IndexMax - 1)
''                Else
''                    sngFWHMScanStart = lngScanNumbers(.IndexBaseLeft)
''                End If
''            End If
''
''            sngFWHMScanEnd = -1
''            For lngDataIndex = .IndexBaseRight - 1 To .IndexMax Step -1
''                If blnSubtractBaselineNoise Then
''                    sngY1 = BaselineAdjustIntensity(sngIntensityList(lngDataIndex), .BaselineNoiseStats.NoiseLevel, ALLOW_NEGATIVE_VALUES)
''                    sngY2 = BaselineAdjustIntensity(sngIntensityList(lngDataIndex + 1), .BaselineNoiseStats.NoiseLevel, ALLOW_NEGATIVE_VALUES)
''                Else
''                    sngY1 = sngIntensityList(lngDataIndex)
''                    sngY2 = sngIntensityList(lngDataIndex + 1)
''                End If
''
''                If sngY1 > sngTargetIntensity Or sngY2 > sngTargetIntensity Then
''                    If sngY1 >= sngTargetIntensity And sngY2 <= sngTargetIntensity Then
''                        blnSuccess = InterpolateX(sngFWHMScanEnd, lngScanNumbers(lngDataIndex), lngScanNumbers(lngDataIndex + 1), sngY1, sngY2, sngTargetIntensity)
''                    Else
''                        ' sngTargetIntensity is not between sngY1 and sngY2; simply use lngDataIndex+1
''                        If lngDataIndex = .IndexBaseRight - 1 Then
''                            ' At the end of the peak; use the scan number halfway between .IndexBaseRight and .IndexMax
''                            sngFWHMScanEnd = lngScanNumbers(lngDataIndex + 1 - CInt(Math.Round((.IndexBaseRight - .IndexMax) / 2, 0)))
''                        Else
''                            ' This code will probably never be reached
''                            sngFWHMScanEnd = lngScanNumbers(lngDataIndex + 1)
''                        End If
''                    End If
''                    Exit For
''                End If
''            Next lngDataIndex
''            If sngFWHMScanEnd < 0 Then
''                If .IndexMax < .IndexBaseRight Then
''                    sngFWHMScanEnd = lngScanNumbers(.IndexMax + 1)
''                Else
''                    sngFWHMScanEnd = lngScanNumbers(.IndexBaseRight)
''                End If
''            End If
''
''            lngFWHMScans = CInt(Math.Round(sngFWHMScanEnd - sngFWHMScanStart, 0))
''            If lngFWHMScans <= 0 Then lngFWHMScans = 0
''        Else
''            ' Maximum intensity value is <= 0
''            ' Set FWHM to 1
''            lngFWHMScans = 1
''        End If
''
''    End With
''
''    ComputeFWHM = lngFWHMScans
''    Exit Function
''
''ComputeFWHMErrorHandler:
''    ' Error finding FWHM
''    Debug.Assert False
''    Debug.Print "Error in clsFWHMComputation.ComputeFWHM: " & Err.Description
''
''    ComputeFWHM = 0
''End Function
''
Private Function InterpolateX(ByRef sngInterpolatedXValue As Single, ByVal X1 As Long, ByVal x2 As Long, ByVal Y1 As Single, ByVal Y2 As Single, ByVal sngTargetY As Single) As Boolean
    ' Determines the X value that corresponds to sngTargetY by interpolating the line between (X1, Y1) and (X2, Y2)
    ' Returns True on success, false on error

    Dim sngDeltaY As Single
    Dim sngFraction As Single
    Dim lngDeltaX As Long
    Dim sngTargetX As Single

    sngDeltaY = Y2 - Y1                                 ' This is y-two minus y-one
    sngFraction = (sngTargetY - Y1) / sngDeltaY
    lngDeltaX = x2 - X1                                 ' This is x-two minus x-one

    sngTargetX = sngFraction * lngDeltaX + X1

    If Abs(sngTargetX - X1) >= 0 And Abs(sngTargetX - x2) >= 0 Then
        sngInterpolatedXValue = sngTargetX
        InterpolateX = True
    Else
        ' TargetX is not between X1 and X2; this shouldn't happen
        Debug.Assert False
        InterpolateX = False
    End If

End Function

Private Function InterpolateY(ByRef sngInterpolatedIntensity As Single, ByVal X1 As Long, ByVal x2 As Long, ByVal Y1 As Single, ByVal Y2 As Single, ByVal sngXValToInterpolate As Single) As Boolean
    ' Given two X,Y coordinates interpolate or extrapolate to determine the Y value that would be seen for a given X value

    Dim lngScanDifference As Long

    lngScanDifference = x2 - X1
    If lngScanDifference <> 0 Then
        sngInterpolatedIntensity = Y1 + (Y2 - Y1) * ((sngXValToInterpolate - X1) / lngScanDifference)
        InterpolateY = True
    Else
        ' sngXValToInterpolate is not between X1 and X2; cannot interpolate
        InterpolateY = False
    End If
End Function


Attribute VB_Name = "modBinnedDataPeakStats"
Option Explicit

' These functions are used to find the best, single peak in histogramed mass-error or NET-error data
' Written by Matthew Monroe
' Started February 14, 2005

Public Enum eNoiseThresholdModes
    AbsoluteThreshold = 0
    TrimmedMeanByAbundance = 1
    TrimmedMeanByCount = 2
    TrimmedMedianByAbundance = 3
End Enum

Public Type udtBinnedDataType
    BinnedCount As Long
    Binned() As Long            ' 0-based array, but ranges from index 0 to index .BinnedCount; Counts of number of hits for each bin
    SmoothedBins() As Double
    
    StartBin As Single
    BinSize As Single
    BinRangeMaximum As Single
    
    BinCountMaximum As Long
End Type

Public Type udtPeakStatsType
    MaximumIntensity As Double
    SignalToNoiseRatio As Single
    IndexOfMaximum As Long
    IndexOfCenterOfMass As Double       ' Decimal index, to allow for the center of mass falling between two bins
    IndexBaseLeft As Long               ' Index of the left side of the peak at the base of the peak
    IndexBaseRight As Long              ' Index of the right side of the peak at the base of the peak
    IndexLeft As Long                   ' Note: This index is ultimately the index of the left side of the peak at udtFindPeaksOptions.PercentageOfMaximumForPeakWidth of the peak maximum
    IndexRight As Long                  ' Note: This index is ultimately the index of the right side of the peak at udtFindPeaksOptions.PercentageOfMaximumForPeakWidth of the peak maximum

    TruePositiveArea As Long                ' Area in the peak that is above the background level
    FalsePositiveArea As Long               ' Area in the peak that is below the background level
End Type

Public Type udtNoiseThresholdOptionsType
    NoiseThresholdMode As eNoiseThresholdModes           ' Method to use to determine the noise level; default is eNoiseThresholdModes.TrimmedMedianByAbundance
    NoiseThresholdIntensity As Single                    ' Absolute noise threshold to apply; typically 0
    NoiseFractionLowIntensityDataToAverage As Single     ' Typically 0.33 for binned error histograms if ExcludePeakDataFromNoiseComputation = False, ignored if ExcludePeakDataFromNoiseComputation = True
    ExcludePeakDataFromNoiseComputation As Boolean
    MinimumNoiseThresholdLevel As Single                 ' If the noise threshold computed is less than this value, then will use this value to compute S/N; additionally, this is used as the minimum intensity threshold when computing a trimmed noise level
End Type

Public Type udtFindPeaksOptionsType
    PeakWidthPointsMinimum As Long
    MinimumPeakIntensity As Double
    PercentageOfMaximumForPeakWidth As Long
    ButterWorthFrequency As Single
    NoiseThresholdOptions As udtNoiseThresholdOptionsType
End Type

Private Function ComputeAverageNoiseLevelCheckCounts(ByVal lngValidDataCountA As Long, ByVal lngValidDataCountB As Long, ByVal dblSumA As Double, ByVal dblSumB As Double, ByVal lngMinimumCount As Long, ByRef dblNoiseLevel As Double) As Boolean
    
    If lngMinimumCount < 1 Then lngMinimumCount = 1
    If lngValidDataCountA >= lngMinimumCount Or lngValidDataCountB >= lngMinimumCount Then
        If lngValidDataCountA >= lngMinimumCount And lngValidDataCountB >= lngMinimumCount Then
            ' Both meet the minimum; return the lowest noise level
            dblNoiseLevel = dblSumA / lngValidDataCountA
            If dblSumB / lngValidDataCountB < dblNoiseLevel Then
                dblNoiseLevel = dblSumB / lngValidDataCountB
            End If
        ElseIf lngValidDataCountA >= lngMinimumCount Then
            ' Side A meets the minimum
            dblNoiseLevel = dblSumA / lngValidDataCountA
        Else
            ' Side B meets the minimum
            dblNoiseLevel = dblSumB / lngValidDataCountB
        End If
        ComputeAverageNoiseLevelCheckCounts = True
    Else
        ComputeAverageNoiseLevelCheckCounts = False
    End If

End Function

Private Function ComputeAverageNoiseLevelExcludingRegion(ByRef dblData() As Double, ByVal lngPeakWidthPoints As Long, ByVal lngIndexStart As Long, ByVal lngIndexEnd As Long, ByVal lngExclusionIndexStart As Long, ByVal lngExclusionIndexEnd As Long, ByRef udtNoiseThresholdOptions As udtNoiseThresholdOptionsType, ByVal blnIgnoreNonPositiveData As Boolean) As Double

    ' Compute the average intensity level between lngIndexStart and lngExclusionIndexStart
    ' Also compute the average between lngExclusionIndexEnd and lngIndexEnd
    ' Return the smaller of the two averages provided sufficient data points were used to compute the average
    ' If neither segment contains sufficient data points, then call ComputeTrimmedNoiseLevel
    '
    ' Preferably, require lngPeakWidthPoints/2 points
    ' If neither side has that many points, then require at least MINIMUM_PEAK_WIDTH=3 points
    
    Const MINIMUM_PEAK_WIDTH = 3
    
    Dim lngIndex As Long
    Dim lngValidDataCountA As Long
    Dim lngValidDataCountB As Long
    
    Dim dblSumA As Double
    Dim dblSumB As Double
    
    Dim dblNoiseLevel As Double
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    
    If lngExclusionIndexStart >= lngIndexStart And lngExclusionIndexStart <= lngIndexEnd And _
       lngExclusionIndexEnd >= lngIndexStart And lngExclusionIndexEnd <= lngIndexEnd And _
       lngExclusionIndexStart <= lngExclusionIndexEnd Then
       
        lngValidDataCountA = 0
        dblSumA = 0
        For lngIndex = lngIndexStart To lngExclusionIndexStart - 1
            If dblData(lngIndex) > udtNoiseThresholdOptions.MinimumNoiseThresholdLevel Then
                dblSumA = dblSumA + dblData(lngIndex)
                lngValidDataCountA = lngValidDataCountA + 1
            End If
        Next lngIndex
        
        lngValidDataCountB = 0
        dblSumB = 0
        For lngIndex = lngExclusionIndexEnd + 1 To lngIndexEnd
            If dblData(lngIndex) > udtNoiseThresholdOptions.MinimumNoiseThresholdLevel Then
                dblSumB = dblSumB + dblData(lngIndex)
                lngValidDataCountB = lngValidDataCountB + 1
            End If
        Next lngIndex
''
''        If Int(lngPeakWidthPoints / 2) >= MINIMUM_PEAK_WIDTH * 2 Then
''            blnSuccess = ComputeAverageNoiseLevelCheckCounts(lngValidDataCountA, lngValidDataCountB, dblSumA, dblSumB, Int(lngPeakWidthPoints / 2), dblNoiseLevel)
''        End If
''
''        If Not blnSuccess Then
''            blnSuccess = ComputeAverageNoiseLevelCheckCounts(lngValidDataCountA, lngValidDataCountB, dblSumA, dblSumB, MINIMUM_PEAK_WIDTH * 2, dblNoiseLevel)
''        End If
''
''        If Not blnSuccess Then
            blnSuccess = ComputeAverageNoiseLevelCheckCounts(lngValidDataCountA, lngValidDataCountB, dblSumA, dblSumB, MINIMUM_PEAK_WIDTH, dblNoiseLevel)
''        End If
        
    End If
    
    If Not blnSuccess Then
        Dim udtNoiseThresholdOptionsOverride As udtNoiseThresholdOptionsType
        
        udtNoiseThresholdOptionsOverride = udtNoiseThresholdOptions
        With udtNoiseThresholdOptionsOverride
            .NoiseThresholdMode = eNoiseThresholdModes.TrimmedMedianByAbundance
            .NoiseFractionLowIntensityDataToAverage = 0.33
        End With
        
        dblNoiseLevel = ComputeTrimmedNoiseLevel(dblData, lngIndexStart, lngIndexEnd, udtNoiseThresholdOptionsOverride, True)
    End If
    
    ComputeAverageNoiseLevelExcludingRegion = dblNoiseLevel
    
End Function

Private Sub ComputeNoiseLevelInPeakVicinity(ByRef udtPeak As udtPeakStatsType, ByRef dblData() As Double, ByVal lngDataCount As Long, ByRef udtFindPeaksOptions As udtFindPeaksOptionsType)

    Const IgnoreNonPositiveData As Boolean = True
    
    Dim dblNoiseThresholdIntensity As Double
    Dim lngIndexStart As Long
    Dim lngIndexEnd As Long
    Dim lngPeakWidthPoints As Long
    
    ' Only use a portion of the data to compute the noise level
    ' The number of points to extend from the left and right is based on the peak width divided by 2
    lngPeakWidthPoints = udtPeak.IndexBaseRight - udtPeak.IndexBaseLeft
    If lngPeakWidthPoints < udtFindPeaksOptions.PeakWidthPointsMinimum Then
        lngPeakWidthPoints = udtFindPeaksOptions.PeakWidthPointsMinimum
    End If
    
    lngIndexStart = udtPeak.IndexBaseLeft - lngPeakWidthPoints
    lngIndexEnd = udtPeak.IndexBaseRight + lngPeakWidthPoints
    
    If lngIndexStart < 0 Then lngIndexStart = 0
    If lngIndexEnd >= lngDataCount Then lngIndexEnd = lngDataCount - 1
    
    If udtFindPeaksOptions.NoiseThresholdOptions.ExcludePeakDataFromNoiseComputation Then
        dblNoiseThresholdIntensity = ComputeAverageNoiseLevelExcludingRegion(dblData, lngPeakWidthPoints, lngIndexStart, lngIndexEnd, udtPeak.IndexBaseLeft, udtPeak.IndexBaseRight, udtFindPeaksOptions.NoiseThresholdOptions, IgnoreNonPositiveData)
    Else
        dblNoiseThresholdIntensity = ComputeTrimmedNoiseLevel(dblData, lngIndexStart, lngIndexEnd, udtFindPeaksOptions.NoiseThresholdOptions, IgnoreNonPositiveData)
    End If
    
    With udtFindPeaksOptions.NoiseThresholdOptions
        If dblNoiseThresholdIntensity < .MinimumNoiseThresholdLevel And .MinimumNoiseThresholdLevel > 0 Then
            dblNoiseThresholdIntensity = .MinimumNoiseThresholdLevel
        End If
    End With

    If dblNoiseThresholdIntensity > 0 Then
        If udtPeak.IndexOfMaximum >= 0 And udtPeak.IndexOfMaximum < lngDataCount Then
            udtPeak.SignalToNoiseRatio = dblData(udtPeak.IndexOfMaximum) / dblNoiseThresholdIntensity
        Else
            ' This shouldn't happen
            Debug.Assert False
            udtPeak.SignalToNoiseRatio = 0
        End If
    Else
        udtPeak.SignalToNoiseRatio = 0
    End If
    
End Sub

Private Function ComputeTrimmedNoiseLevel(ByRef dblData() As Double, ByVal lngIndexStart As Long, ByVal lngIndexEnd As Long, ByRef udtNoiseThresholdOptions As udtNoiseThresholdOptionsType, ByVal blnIgnoreNonPositiveData As Boolean) As Double
    Dim lngDataSortedCount As Long
    Dim dblDataSorted() As Double
    Dim EmptyArray() As Long            ' Empty array; required for call to .QSAsc()

    Dim dblIntensityThreshold As Double
    Dim dblSum As Double

    Dim dblNoiseThreshold As Double
    
    Dim lngIndex As Long
    Dim lngThresholdPointIndex As Long
    Dim lngValidDataCount As Long

    Dim lngCountForAverage As Long
    
    Dim objQSDouble As New QSDouble
    
    If lngIndexEnd - lngIndexStart < 0 Then
        ComputeTrimmedNoiseLevel = 0
        Exit Function
    End If

    ' Copy the data into dblDataSorted
    lngDataSortedCount = lngIndexEnd - lngIndexStart + 1
    ReDim dblDataSorted(lngDataSortedCount - 1)
    
    For lngIndex = lngIndexStart To lngIndexEnd
        dblDataSorted(lngIndex - lngIndexStart) = dblData(lngIndex)
    Next lngIndex
    

    ' Sort the array
    If Not objQSDouble.QSAsc(dblDataSorted, EmptyArray) Then
        Debug.Assert False
        ComputeTrimmedNoiseLevel = 0
        Exit Function
    End If

    If blnIgnoreNonPositiveData Then
        ' Remove data with a value <= udtNoiseThresholdOptions.MinimumNoiseThresholdLevel

        If dblDataSorted(0) <= udtNoiseThresholdOptions.MinimumNoiseThresholdLevel Then
            lngValidDataCount = 0
            For lngIndex = 0 To lngDataSortedCount - 1
                If dblDataSorted(lngIndex) > udtNoiseThresholdOptions.MinimumNoiseThresholdLevel Then
                    dblDataSorted(lngValidDataCount) = dblDataSorted(lngIndex)
                    lngValidDataCount = lngValidDataCount + 1
                End If
            Next lngIndex

            If lngValidDataCount < lngDataSortedCount Then
                lngDataSortedCount = lngValidDataCount
            End If

            If lngDataSortedCount = 0 Then
                ComputeTrimmedNoiseLevel = 0
                Exit Function
            End If
        End If

    End If

    Select Case udtNoiseThresholdOptions.NoiseThresholdMode
    Case eNoiseThresholdModes.TrimmedMeanByAbundance, eNoiseThresholdModes.TrimmedMeanByCount

        If udtNoiseThresholdOptions.NoiseThresholdMode = eNoiseThresholdModes.TrimmedMeanByAbundance Then
            ' Average the data that has intensity values less than
            '  Minimum + udtNoiseThresholdOptions.NoiseFractionLowIntensityDataToAverage * (Maximum - Minimum)
            With udtNoiseThresholdOptions
                dblIntensityThreshold = dblDataSorted(0) + .NoiseFractionLowIntensityDataToAverage * (dblDataSorted(lngDataSortedCount - 1) - dblDataSorted(0))
            End With

            lngCountForAverage = 0
            dblSum = 0
            For lngIndex = 0 To lngDataSortedCount - 1
                If dblDataSorted(lngIndex) <= dblIntensityThreshold Then
                    dblSum = dblSum + dblDataSorted(lngIndex)
                Else
                    lngCountForAverage = lngIndex
                    Exit For
                End If
            Next lngIndex
        Else
            ' eNoiseThresholdModes.TrimmedMeanByCount
            ' Find the index of the data point at lngDataSortedCount * udtNoiseThresholdOptions.NoiseFractionLowIntensityDataToAverage and
            ' average the data from the start to that index
            lngThresholdPointIndex = CLng(Round((lngDataSortedCount - 1) * udtNoiseThresholdOptions.NoiseFractionLowIntensityDataToAverage, 0))

            lngCountForAverage = lngThresholdPointIndex + 1
            dblSum = 0
            For lngIndex = 0 To lngThresholdPointIndex
                dblSum = dblSum + dblDataSorted(lngIndex)
            Next lngIndex

        End If

        If lngCountForAverage > 0 Then
            ' Return the average
            dblNoiseThreshold = dblSum / CDbl(lngCountForAverage)
        Else
            ' No data to average; define the noise level to be the minimum intensity
            dblNoiseThreshold = dblDataSorted(0)
        End If

    Case eNoiseThresholdModes.TrimmedMedianByAbundance
        ' Find the median of the data that has intensity values less than
        '  Minimum + udtNoiseThresholdOptions.NoiseFractionLowIntensityDataToAverage * (Maximum - Minimum)
        With udtNoiseThresholdOptions
            dblIntensityThreshold = dblDataSorted(0) + .NoiseFractionLowIntensityDataToAverage * (dblDataSorted(lngDataSortedCount - 1) - dblDataSorted(0))
        End With

        ' Find the first point with an intensity value <= dblIntensityThreshold
        lngThresholdPointIndex = BinarySearchDblFindNearest(dblDataSorted, dblIntensityThreshold, 0, lngDataSortedCount - 1, False)

        If lngThresholdPointIndex Mod 2 = 0 Then
            ' Even value
            dblNoiseThreshold = dblDataSorted(CLng(lngThresholdPointIndex / 2))
        Else
            ' Odd value; average the values on either side of lngThresholdPointIndex/2
            lngIndex = CLng((lngThresholdPointIndex - 1) / 2)
            If lngIndex < 0 Then lngIndex = 0
            dblSum = dblDataSorted(lngIndex)

            lngIndex = lngIndex + 1
            If lngIndex = lngDataSortedCount Then lngIndex = lngDataSortedCount - 1
            dblSum = dblSum + dblDataSorted(lngIndex)

            dblNoiseThreshold = dblSum / 2#
        End If

    Case Else
        ' Unknown mode
        Debug.Print "Unknown Noise Threshold Mode encountered: " & udtNoiseThresholdOptions.NoiseThresholdMode
        Debug.Assert False
        dblNoiseThreshold = 0
    End Select

    ComputeTrimmedNoiseLevel = dblNoiseThreshold
    
End Function

Public Function FindPeakStatsUsingBinnedErrorData(ByRef udtBinnedData As udtBinnedDataType, ByRef udtPeak As udtPeakStatsType, ByRef blnSingleGoodPeakFound As Boolean) As Boolean
    ' Returns True if success, False if error
    
    Const MINIMUM_PEAK_WIDTH = 3    ' Essentially minimum width, in bins; may want to make this adjustable
    
    Dim dblXData() As Double        ' 0-based array
    Dim dblYData() As Double        ' 0-based array
    
    Dim lngDataCount As Long
    
    Dim dblDeltaX As Double
    Dim lngIndex As Long
    Dim blnSuccess As Boolean
            
    Dim udtFindPeaksOptions As udtFindPeaksOptionsType
    
    ' Set this to true for now; we'll change it to false if there is a problem
    blnSuccess = True
    blnSingleGoodPeakFound = False
    
    If blnSuccess Then
        If udtBinnedData.BinnedCount <= 0 Then
            ' No data; no point in trying to find a peak
            blnSuccess = False
        Else
            With udtBinnedData
                ReDim dblXData(.BinnedCount)
                ReDim dblYData(.BinnedCount)
                ReDim .SmoothedBins(.BinnedCount)
                
                lngDataCount = .BinnedCount + 1
                dblDeltaX = Round(.BinSize, 6)
                
                For lngIndex = 0 To .BinnedCount
                    dblXData(lngIndex) = .StartBin + lngIndex * .BinSize
                    dblYData(lngIndex) = .Binned(lngIndex)
                    .SmoothedBins(lngIndex) = .Binned(lngIndex)
                Next lngIndex
            End With
    
            ' The majority of these options are hard coded
            With udtFindPeaksOptions
                .PeakWidthPointsMinimum = MINIMUM_PEAK_WIDTH
                .MinimumPeakIntensity = glbPreferencesExpanded.RefineMSDataOptions.MinimumPeakHeight
                .PercentageOfMaximumForPeakWidth = glbPreferencesExpanded.RefineMSDataOptions.PercentageOfMaxForFindingWidth
                .ButterWorthFrequency = glbPreferencesExpanded.ErrorPlottingOptions.ButterWorthFrequency
                
                With .NoiseThresholdOptions
                    .NoiseThresholdMode = eNoiseThresholdModes.TrimmedMedianByAbundance
                    .NoiseThresholdIntensity = 0
                    .NoiseFractionLowIntensityDataToAverage = 0.33
                    .ExcludePeakDataFromNoiseComputation = True
                    .MinimumNoiseThresholdLevel = 1
                End With
            End With

            blnSuccess = FindPeaksWork(udtPeak, dblXData(), dblYData(), udtBinnedData.SmoothedBins(), lngDataCount, dblDeltaX, udtFindPeaksOptions, blnSingleGoodPeakFound)
            
            If blnSuccess Then
                '' Uncomment for debugging purposes
                '' Debug.Print Now() & "; Intensity=" & Round(udtPeak.MaximumIntensity, 3) & ", S/N=" & Round(udtPeak.SignalToNoiseRatio, 3)
                
                ' If the peak intensity is less than 5 times the minimum peak height, then check its signal to noise ratio against the minimum signal to noise ratio
                ' Otherwise, ignore the S/N ratio since it's method of calculation isn't as robust as desired
                If udtPeak.MaximumIntensity < glbPreferencesExpanded.RefineMSDataOptions.MinimumPeakHeight * 5 Then
                    If udtPeak.SignalToNoiseRatio < glbPreferencesExpanded.RefineMSDataOptions.MinimumSignalToNoiseRatioForLowAbundancePeaks Then
                        ' Peak was found, but its signal to noise ratio is too low
                        ' Do not return the peak
                        
                        blnSuccess = False
                        blnSingleGoodPeakFound = False
                    End If
                End If
            End If
        End If
    End If
    
    FindPeakStatsUsingBinnedErrorData = blnSuccess
End Function

Private Function FindPeaksWork(ByRef udtPeak As udtPeakStatsType, ByRef dblXData() As Double, ByRef dblYData() As Double, ByRef dblYDataSmoothed() As Double, lngDataCount As Long, dblDeltaX As Double, ByRef udtFindPeaksOptions As udtFindPeaksOptionsType, ByRef blnSingleGoodPeakFound As Boolean) As Boolean
    ' dblXData() and dblYData() should be 0-based
    ' Returns True if a valid peak is found; false otherwise

    Const PeakDetectIntensityThresholdPercentageOfMaximum = 15
    Const PeakWidthInSigma = 3
    
    Dim blnValidPeakFound As Boolean
    Dim blnAmbiguousPeaksPresent As Boolean
    
    Dim lngIndex As Long
    Dim lngIndexOfMostIntensePeak As Long
    
    Dim lngLeftIndexA As Long, lngLeftIndexB As Long
    Dim lngRightIndexA As Long, lngRightIndexB As Long
    
    Dim lngPeakCount As Long
    Dim lngPeakLocations() As Long                  ' 0-based array
    Dim lngPeakEdgesLeft() As Long                  ' 0-based array, parallel with lngPeakLocations()
    Dim lngPeakEdgesRight() As Long                 ' 0-based array, parallel with lngPeakLocations()
    Dim dblPeakAreas() As Double                    ' 0-based array, parallel with lngPeakLocations()
    
    Dim dblIntensityThresholdCheck As Double
    Dim dblPeakMinimum As Double, dblPeakMaximum As Double
    Dim dblThreshold As Double
    
    Dim objPeakFinder As New clsPeakDetection
    Dim blnDataIsSmoothed As Boolean
    Dim strErrorMessage As String
    
    
    ' Validate .PeakWidthPointsMinimum
    If dblDeltaX >= 1 Then
        If udtFindPeaksOptions.PeakWidthPointsMinimum < 5 Then
            udtFindPeaksOptions.PeakWidthPointsMinimum = 5
        End If
    Else
        If udtFindPeaksOptions.PeakWidthPointsMinimum < 7 Then
            udtFindPeaksOptions.PeakWidthPointsMinimum = 7
        End If
    End If
    
    If lngDataCount <= 0 Then
        ' Don't call this function unless lngDataCount is > 0
        Debug.Assert False
        
        udtPeak.IndexOfMaximum = 0
        udtPeak.MaximumIntensity = 0
        FindPeaksWork = False
        Exit Function
    End If
    
    If udtFindPeaksOptions.PercentageOfMaximumForPeakWidth < 0 Then udtFindPeaksOptions.PercentageOfMaximumForPeakWidth = 0
    If udtFindPeaksOptions.PercentageOfMaximumForPeakWidth > 100 Then udtFindPeaksOptions.PercentageOfMaximumForPeakWidth = 100
    
    blnSingleGoodPeakFound = False
    
    ' 1. Find the Peak Maximum (using the unsmoothed data)
    With udtPeak
        .IndexOfMaximum = 0
        .IndexLeft = .IndexOfMaximum
        .IndexRight = .IndexOfMaximum
        .MaximumIntensity = dblYData(0)
    End With
    
    For lngIndex = 1 To lngDataCount - 1
        If dblYData(lngIndex) > udtPeak.MaximumIntensity Then
            udtPeak.IndexOfMaximum = lngIndex
            udtPeak.MaximumIntensity = dblYData(lngIndex)
        End If
    Next lngIndex
    
    ' 2. Make sure the maximum value is at least as large as .MinimumPeakIntensity
    If udtPeak.MaximumIntensity >= udtFindPeaksOptions.MinimumPeakIntensity Then
    
        ' 3. Smooth the Y data, and store in dblSmoothedYData()
        ' Since we're using a Butterworth filter, we increase lngPeakWidthPointsMinimum if too small, compared to 1/SamplingFrequency
        blnDataIsSmoothed = FindPeakStatsWorkSmoothData(dblXData(), dblYData(), dblYDataSmoothed(), lngDataCount, udtFindPeaksOptions, strErrorMessage)
        
        If blnDataIsSmoothed Then
            ' Smoothing succeeded
            
            ' Find the Peak Maximum using the smoothed data
            With udtPeak
                .IndexOfMaximum = 0
                .IndexLeft = .IndexOfMaximum
                .IndexRight = .IndexOfMaximum
                .MaximumIntensity = dblYDataSmoothed(0)
            End With
            
            For lngIndex = 1 To lngDataCount - 1
                If dblYDataSmoothed(lngIndex) > udtPeak.MaximumIntensity Then
                    udtPeak.IndexOfMaximum = lngIndex
                    udtPeak.MaximumIntensity = dblYDataSmoothed(lngIndex)
                End If
            Next lngIndex
                        
        Else
            ' Data smoothing failed
            ' Copy the data from dblYData() into dblYDataSmoothed() since we use dblYDataSmoothed() from now on
            ReDim dblYDataSmoothed(lngDataCount - 1)
            For lngIndex = 0 To lngDataCount - 1
                dblYDataSmoothed(lngIndex) = dblYData(lngIndex)
            Next lngIndex
        End If
    
        
        ' 4. Find the peaks
        lngPeakCount = objPeakFinder.DetectPeaks(dblXData(), dblYDataSmoothed(), lngDataCount, _
                                                 udtFindPeaksOptions.MinimumPeakIntensity, _
                                                 udtFindPeaksOptions.PeakWidthPointsMinimum, _
                                                 lngPeakLocations(), lngPeakEdgesLeft(), lngPeakEdgesRight(), _
                                                 dblPeakAreas(), PeakDetectIntensityThresholdPercentageOfMaximum, _
                                                 PeakWidthInSigma, True, True)
               
               
        ' 5. See if a single, good peak is present
        If lngPeakCount > 1 Then
            ' See if the highest intensity peak is at least 2 times higher than all of the other peaks
            ' If it is, then call that the single, valid peak
            
            lngIndexOfMostIntensePeak = 0
            For lngIndex = 0 To lngPeakCount - 1
                If dblYDataSmoothed(lngPeakLocations(lngIndex)) > dblYDataSmoothed(lngPeakLocations(lngIndexOfMostIntensePeak)) Then
                    lngIndexOfMostIntensePeak = lngIndex
                End If
            Next lngIndex
            
            dblIntensityThresholdCheck = dblYDataSmoothed(lngPeakLocations(lngIndexOfMostIntensePeak)) / 2
            
            ' See if any of the peaks is larger than dblIntensityThresholdCheck
            blnAmbiguousPeaksPresent = False
            For lngIndex = 0 To lngPeakCount - 1
                If lngIndex <> lngIndexOfMostIntensePeak Then
                    If dblYDataSmoothed(lngPeakLocations(lngIndex)) >= dblIntensityThresholdCheck Then
                        blnAmbiguousPeaksPresent = True
                        Exit For
                    End If
                End If
            Next lngIndex
        
            If Not blnAmbiguousPeaksPresent Then
                 If lngIndexOfMostIntensePeak > 0 Then
                    lngPeakLocations(0) = lngPeakLocations(lngIndexOfMostIntensePeak)
                    lngPeakEdgesLeft(0) = lngPeakEdgesLeft(lngIndexOfMostIntensePeak)
                    lngPeakEdgesRight(0) = lngPeakEdgesRight(lngIndexOfMostIntensePeak)
                    dblPeakAreas(0) = dblPeakAreas(lngIndexOfMostIntensePeak)
                End If
                lngPeakCount = 1
            End If
        End If
    
        If lngPeakCount = 1 Then
            ' Yes, a single good peak was found
            With udtPeak
                .IndexLeft = lngPeakEdgesLeft(0)
                .IndexRight = lngPeakEdgesRight(0)
                .IndexOfMaximum = lngPeakLocations(0)
                
                If .IndexLeft < 0 Then .IndexLeft = 0
                If .IndexRight >= lngDataCount Then .IndexRight = lngDataCount - 1
                If .IndexOfMaximum < .IndexLeft Then .IndexOfMaximum = .IndexLeft
                If .IndexOfMaximum > .IndexRight Then .IndexOfMaximum = .IndexRight
                
                ' Store the current values of .IndexLeft and .IndexRight
                .IndexBaseLeft = .IndexLeft
                .IndexBaseRight = .IndexRight
                
                '' UpdatePeakEdges udtPeak, dblXData(), dblYDataSmoothed(), lngDataCount, lngPeakWidthPointsMinimum, lngPercentageOfMaximumForPeakWidth
                
                ' udtPeak.IndexLeft and udtPeak.IndexRight are the peak edges at the base of the peak
                ' We want to know the peak width at udtFindPeaksOptions.PercentageOfMaximumForPeakWidth of the peak maximum
                ' Since the peak is often present with a y-axis offset, find the minimum and maximum values
                '  of the data between .IndexLeft and .IndexRight, then compute maximum - minimum and take
                '  the result times lngPercentageOfMaximumForPeakWidth/100; now find the data points that
                '  cross this threshold, looking both from the edges to the middle and the middle to the edges.
                ' If a discrepancy exists between the threshold crossing points, then compute the distance
                '  halfway between the threshold crossing points.  Define this threshold position as the new peak edges
            
                dblPeakMinimum = dblYDataSmoothed(.IndexLeft)
                dblPeakMaximum = dblYDataSmoothed(.IndexLeft)
                For lngIndex = .IndexLeft + 1 To .IndexRight
                    If dblYDataSmoothed(lngIndex) < dblPeakMinimum Then dblPeakMinimum = dblYDataSmoothed(lngIndex)
                    If dblYDataSmoothed(lngIndex) > dblPeakMaximum Then dblPeakMaximum = dblYDataSmoothed(lngIndex)
                Next lngIndex
                
                ' Only proceed if the observed peak height is >= 50% of the minimum peak intensity
                ' This should always be the case, but checking just to be sure
                If dblPeakMaximum - dblPeakMinimum > udtFindPeaksOptions.MinimumPeakIntensity / 2 Then
                    ' Compute the intensity threshold
                    dblThreshold = (dblPeakMaximum - dblPeakMinimum) * CDbl(udtFindPeaksOptions.PercentageOfMaximumForPeakWidth) / 100# + dblPeakMinimum
                    
                    ' Find the indices where the intensity threshold is crossed
                    
                    lngLeftIndexA = .IndexLeft              ' Note that LeftIndexA < LeftIndexB
                    lngLeftIndexB = .IndexOfMaximum
                    For lngIndex = .IndexLeft To .IndexOfMaximum
                        If dblYDataSmoothed(lngIndex) >= dblThreshold Then
                            lngLeftIndexA = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                    
                    For lngIndex = .IndexOfMaximum To .IndexLeft Step -1
                        If dblYDataSmoothed(lngIndex) <= dblThreshold Then
                            lngLeftIndexB = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                    
                    If lngLeftIndexB < lngLeftIndexA Then
                        If lngLeftIndexA > 0 Then lngLeftIndexA = lngLeftIndexA - 1
                        lngLeftIndexB = lngLeftIndexB
                    End If
                    
                    If lngLeftIndexA <> lngLeftIndexB Then
                        ' Find the midpoint between A and B
                        lngLeftIndexA = lngLeftIndexA + Abs(lngLeftIndexB - lngLeftIndexA) / 2
                    End If
                    .IndexLeft = lngLeftIndexA
                    
                    
                    lngRightIndexB = .IndexOfMaximum        ' Note that RightIndexB < RightIndexA
                    lngRightIndexA = .IndexRight
                    For lngIndex = .IndexRight To .IndexOfMaximum Step -1
                        If dblYDataSmoothed(lngIndex) >= dblThreshold Then
                            lngRightIndexA = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                    
                    For lngIndex = .IndexOfMaximum To .IndexRight
                        If dblYDataSmoothed(lngIndex) <= dblThreshold Then
                            lngRightIndexB = lngIndex
                            Exit For
                        End If
                    Next lngIndex
                    
                    If lngRightIndexB > lngRightIndexA Then
                        If lngRightIndexA < lngDataCount Then lngRightIndexA = lngRightIndexA + 1
                        lngRightIndexB = lngRightIndexA
                    End If
                    
                    If lngRightIndexA <> lngRightIndexB Then
                        ' Find the midpoint between A and B
                        lngRightIndexA = lngRightIndexA - Abs(lngRightIndexA - lngRightIndexB) / 2
                    End If
                    .IndexRight = lngRightIndexA
                    
                End If
            End With
            
            udtPeak.IndexOfCenterOfMass = FindCenterOfMassIndex(udtPeak, dblXData, dblYDataSmoothed, lngDataCount, dblDeltaX)
            
            blnValidPeakFound = True
            blnSingleGoodPeakFound = True
        Else
            ' No, couldn't find a single good peak; use FindPeakStatsSimple instead
            blnValidPeakFound = FindPeakStatsSimple(udtPeak, dblXData, dblYDataSmoothed, lngDataCount, dblDeltaX, udtFindPeaksOptions)
        End If
        
        If blnValidPeakFound Then
            
            FindPeakStatsWorkExtendPeakBaseEdges udtPeak, dblYDataSmoothed(), lngDataCount
            ComputeNoiseLevelInPeakVicinity udtPeak, dblYDataSmoothed(), lngDataCount, udtFindPeaksOptions
            ComputeRelativeRiskAreas udtPeak, dblYData(), dblYDataSmoothed(), lngDataCount, udtFindPeaksOptions
        Else
            With udtPeak
                .SignalToNoiseRatio = 0
                .TruePositiveArea = 0
                .FalsePositiveArea = 0
            End With
        End If
    Else
        blnValidPeakFound = False
    End If

'' Uncomment for debugging purposes
''    If udtPeak.SignalToNoiseRatio > 200 Then
''        For lngIndex = 0 To lngDataCount - 1
''            Debug.Print dblYData(lngIndex) & " " & dblYDataSmoothed(lngIndex)
''        Next lngIndex
''    End If
    
    FindPeaksWork = blnValidPeakFound
    
End Function

Private Sub FindPeakStatsWorkExtendPeakBaseEdges(ByRef udtPeak As udtPeakStatsType, ByRef dblYDataSmoothed() As Double, ByVal lngDataCount As Long)
    ' Extend the base index positions left and right until:
    '  a) the intensity stops decreasing
    '  b) we hit a negative value, or
    '  c) we fall below the median of the data not within the peak
    
    Dim dblMedianLeft As Double
    Dim dblMedianRight As Double
    
    ' Exit sub if dblYDataSmoothed() contains fewer than 2 data points
    If lngDataCount <= 1 Then
        Exit Sub
    End If
    
    ' Find the median of the data left of .IndexBaseLeft and right of .IndexBaseRight
    dblMedianLeft = FindPeakStatsWorkComputeMedian(0, udtPeak.IndexBaseLeft, dblYDataSmoothed(), lngDataCount)
    dblMedianRight = FindPeakStatsWorkComputeMedian(udtPeak.IndexBaseRight, lngDataCount - 1, dblYDataSmoothed(), lngDataCount)
        
    With udtPeak
        Do While .IndexBaseLeft > 1
            If dblYDataSmoothed(.IndexBaseLeft - 1) < dblYDataSmoothed(.IndexBaseLeft) And dblYDataSmoothed(.IndexBaseLeft - 1) > dblMedianLeft Then
                .IndexBaseLeft = .IndexBaseLeft - 1
            Else
                Exit Do
            End If
        Loop
    
        Do While .IndexBaseRight < lngDataCount - 2
            If dblYDataSmoothed(.IndexBaseRight + 1) < dblYDataSmoothed(.IndexBaseRight) And dblYDataSmoothed(.IndexBaseRight + 1) > dblMedianRight Then
                .IndexBaseRight = .IndexBaseRight + 1
            Else
                Exit Do
            End If
        Loop
    End With
    
End Sub

Private Function FindPeakStatsWorkComputeMedian(lngIndexStart As Long, lngIndexEnd As Long, ByRef dblData() As Double, ByVal lngDataCount As Long) As Double
    
    Dim lngIndex As Long
    
    Dim lngDataSortedMaxIndex As Long
    Dim dblMedian As Double
    
    Dim dblDataSorted() As Double
    Dim EmptyArray() As Long            ' Empty array; required for call to .QSAsc()

    Dim objQSDouble As New QSDouble
    
    ' Populate dblDataSorted()
    ReDim dblDataSorted(lngIndexEnd - lngIndexStart)
    For lngIndex = lngIndexStart To lngIndexEnd
        dblDataSorted(lngIndex - lngIndexStart) = dblData(lngIndex)
    Next lngIndex
    lngDataSortedMaxIndex = UBound(dblDataSorted)
    
    ' Sort the array
    If Not objQSDouble.QSAsc(dblDataSorted, EmptyArray) Then
        Debug.Assert False
        Exit Function
    End If
    
    If lngDataSortedMaxIndex Mod 2 = 0 Then
        ' Odd number of points
        dblMedian = dblDataSorted(CLng((lngDataSortedMaxIndex) / 2))
    Else
        ' Even number of points in dblDataSorted
        ' Average the values on either side of lngDataSortedMaxIndex/2
        
        lngIndex = CLng((lngDataSortedMaxIndex - 1) / 2)
        If lngIndex < 0 Then lngIndex = 0
        dblMedian = dblDataSorted(lngIndex)

        lngIndex = lngIndex + 1
        If lngIndex > lngDataSortedMaxIndex Then lngIndex = lngDataSortedMaxIndex
        dblMedian = (dblMedian + dblDataSorted(lngIndex)) / 2#
    End If
    
    If dblMedian < 0 Then dblMedian = 0
    
    FindPeakStatsWorkComputeMedian = dblMedian
        
End Function

Private Function FindPeakStatsWorkSmoothData(ByRef dblXData() As Double, ByRef dblYData() As Double, ByRef dblYDataSmoothed() As Double, lngDataCount As Long, ByRef udtFindPeaksOptions As udtFindPeaksOptionsType, ByRef strErrorMessage As String) As Boolean
    ' Returns True if the data was smoothed; false if not or an error
    ' The smoothed data is returned in dblYDataSmoothed(), which is Dimmed inside this function

    Dim intFilterThirdWidth As Integer
    Dim blnSuccess As Boolean

    Dim lngPeakWidthPointsCompare As Long
    Dim lngIndex As Long
    
    Dim objFilter As New clsButterworthFilter

    ReDim dblYDataSmoothed(lngDataCount - 1)

    For lngIndex = 0 To lngDataCount - 1
        dblYDataSmoothed(lngIndex) = dblYData(lngIndex)
    Next lngIndex
    
    ' Filter the data with a Butterworth filter
    blnSuccess = objFilter.ButterworthFilter(dblYDataSmoothed(), 0, lngDataCount - 1, udtFindPeaksOptions.ButterWorthFrequency)
    If Not blnSuccess Then
        Debug.Print "Error with the Butterworth filter (modBinnedDataPeakStats->FindPeakStatsWorkSmothData): " & strErrorMessage
        Debug.Assert False
        blnSuccess = False
    Else
        ' Data was smoothed
        ' Validate that lngPeakWidthPointsMinimum is large enough
        If udtFindPeaksOptions.ButterWorthFrequency > 0 Then
            lngPeakWidthPointsCompare = CLng(Round(1 / udtFindPeaksOptions.ButterWorthFrequency, 0))
            If udtFindPeaksOptions.PeakWidthPointsMinimum < lngPeakWidthPointsCompare Then
                udtFindPeaksOptions.PeakWidthPointsMinimum = lngPeakWidthPointsCompare
            End If
        End If

        blnSuccess = True
    End If

    FindPeakStatsWorkSmoothData = blnSuccess
End Function

Private Function FindCenterOfMassIndex(ByRef udtPeak As udtPeakStatsType, ByRef dblXData() As Double, ByRef dblYData() As Double, lngDataCount As Long, dblDeltaX As Double) As Integer

    Dim lngIndex As Long
    Dim dblXOffset As Double
    Dim dblSumXY As Double, dblSumY As Double
    
    Dim lngCOMIndex As Long
    
    ' Find the center of mass (C.O.M.) of the peak
    ' From: http://www.galactic.com/Algorithms/pc_ctrmass.htm
    ' Formula is: X Coordinate of the C.O.M. = Sum (x_i*y_i) / (Sum(y_i) where i is the index from .IndexLeft to .IndexRight
    ' This formula cannot handle negative x data; thus, if dblXdata(lngindex) is < 0, then apply an offset
    
    dblXOffset = 0
    For lngIndex = 0 To lngDataCount - 1
        If dblXData(lngIndex) < dblXOffset Then
            dblXOffset = dblXData(lngIndex)
        End If
    Next lngIndex
    dblXOffset = Abs(dblXOffset)
    
    With udtPeak
        dblSumXY = 0
        dblSumY = 0
        For lngIndex = .IndexLeft To .IndexRight
            dblSumXY = dblSumXY + (dblXData(lngIndex) + dblXOffset) * dblYData(lngIndex)
            dblSumY = dblSumY + dblYData(lngIndex)
        Next lngIndex
        
        If dblSumY > 0 Then
            ' Note: Must divide by dblDeltaX to get the correct Index using the above center of mass formula
            lngCOMIndex = Round((dblSumXY / dblSumY) / dblDeltaX, 4)
        Else
            ' This shouldn't happen
            Debug.Assert False
            lngCOMIndex = .IndexOfMaximum
        End If
    End With

    FindCenterOfMassIndex = lngCOMIndex

End Function

Public Sub GetPeakStats(ByRef udtBinnedErrorData As udtBinnedDataType, ByRef udtThisPeak As udtPeakStatsType, ByRef dblPeakCenter As Double, ByRef dblPeakWidth As Double, ByRef dblPeakHeight As Double, ByRef sngSignalToNoise As Single, lngDigitsOfPrecisionToRoundTo As Long)

On Error GoTo GetPeakStatsErrorHandler

    With udtBinnedErrorData
        dblPeakCenter = Round(.StartBin + udtThisPeak.IndexOfCenterOfMass * .BinSize, lngDigitsOfPrecisionToRoundTo)
        
        ' Peak width at the base
        dblPeakWidth = Round((udtThisPeak.IndexRight * .BinSize) - (udtThisPeak.IndexLeft * .BinSize), lngDigitsOfPrecisionToRoundTo)
        
        dblPeakHeight = Round(udtThisPeak.MaximumIntensity, lngDigitsOfPrecisionToRoundTo)
        sngSignalToNoise = Round(udtThisPeak.SignalToNoiseRatio, 1)
    End With

    Exit Sub
    
GetPeakStatsErrorHandler:
    Debug.Print "Error in GetPeakStats: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "GetPeakStats"
    Resume Next
    
End Sub

' FindPeakStatsSimple functions
Private Function FindPeakStatsSimple(ByRef udtPeak As udtPeakStatsType, ByRef dblXData() As Double, ByRef dblYData() As Double, lngDataCount As Long, dblDeltaX As Double, udtFindPeaksOptions As udtFindPeaksOptionsType) As Boolean
    ' dblXData() and dblYData() should be 0-based
    ' Returns True if a valid peak is found; false otherwise
    ' Assumes udtPeak.MaximumIntensity has already been populated
    
    Const PERCENTAGE_OF_MAX_FOR_PEAK_WIDTH_AT_BASE As Integer = 10
    
    Dim dblThreshold As Double
    
On Error GoTo FindPeakStatsSimpleErrorHandler
    
    ' Step left and right until we find a point whose intensity is less than lngPercentageOfMaximumForPeakWidth of the maximum intensity
    dblThreshold = udtPeak.MaximumIntensity * CDbl(udtFindPeaksOptions.PercentageOfMaximumForPeakWidth) / 100#
    With udtPeak
        .IndexLeft = .IndexOfMaximum
        .IndexRight = .IndexOfMaximum
        FindPeakStatsSimpleExtendEdges .IndexLeft, .IndexRight, dblThreshold, dblYData(), lngDataCount, udtFindPeaksOptions.PeakWidthPointsMinimum
    End With
    
    ' Do the same, but attempt to find the peak edges at the base of the peak
    dblThreshold = udtPeak.MaximumIntensity * CDbl(PERCENTAGE_OF_MAX_FOR_PEAK_WIDTH_AT_BASE) / 100#
    With udtPeak
        .IndexBaseLeft = .IndexOfMaximum
        .IndexBaseRight = .IndexOfMaximum
        FindPeakStatsSimpleExtendEdges .IndexBaseLeft, .IndexBaseRight, dblThreshold, dblYData(), lngDataCount, udtFindPeaksOptions.PeakWidthPointsMinimum
    End With
    
    udtPeak.IndexOfCenterOfMass = FindCenterOfMassIndex(udtPeak, dblXData, dblYData, lngDataCount, dblDeltaX)

    FindPeakStatsSimple = True
    Exit Function

FindPeakStatsSimpleErrorHandler:
    Debug.Print "Error in FindPeakStatsSimple: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "FindPeakStatsSimple"
    FindPeakStatsSimple = False

End Function

Private Sub FindPeakStatsSimpleExtendEdges(ByRef lngIndexLeft As Long, ByRef lngIndexRight As Long, ByVal dblThreshold As Double, ByRef dblYData() As Double, ByVal lngDataCount As Long, ByVal lngPeakWidthPointsMinimum As Long)

    Do While lngIndexLeft > 0
        lngIndexLeft = lngIndexLeft - 1
        If dblYData(lngIndexLeft) <= dblThreshold Then
            Exit Do
        End If
    Loop

    Do While lngIndexRight < lngDataCount - 1
        lngIndexRight = lngIndexRight + 1
        If dblYData(lngIndexRight) <= dblThreshold Then
            Exit Do
        End If
    Loop

    ' Make sure (lngIndexRight - lngIndexLeft + 1) is >= udtFindPeaksOptions.PeakWidthPointsMinimum
    Do While (lngIndexRight - lngIndexLeft + 1) < lngPeakWidthPointsMinimum
        ' Change lngIndexRight and lngIndexLeft until their width is at least PeakWidthPointsMinimum
        lngIndexLeft = lngIndexLeft - 1
        lngIndexRight = lngIndexRight + 1
        If lngIndexLeft < 0 Then lngIndexLeft = 0
        If lngIndexRight >= lngDataCount Then lngIndexRight = lngDataCount - 1
        If lngIndexLeft = 0 Or lngIndexRight = lngDataCount - 1 Then Exit Do
    Loop
    
End Sub

' Relative Risk Functions
Private Function ComputeMeanInRange(ByRef dblData() As Double, ByVal lngIndexStart As Long, ByVal lngIndexEnd As Long, ByVal blnIgnoreNonPositiveData As Boolean) As Double
   
    Dim lngIndex As Long
   
    Dim dblSum As Double
    Dim lngValidDataCount As Long
    
    dblSum = 0
    lngValidDataCount = 0
    For lngIndex = lngIndexStart To lngIndexEnd
        If Not blnIgnoreNonPositiveData Or dblData(lngIndex) > 0 Then
            dblSum = dblSum + dblData(lngIndex)
            lngValidDataCount = lngValidDataCount + 1
        End If
    Next lngIndex
    
    If lngValidDataCount > 0 Then
        ComputeMeanInRange = dblSum / lngValidDataCount
    Else
        ComputeMeanInRange = 0
    End If
    
End Function

Private Sub ComputeRelativeRiskAreas(ByRef udtPeak As udtPeakStatsType, ByRef dblYData() As Double, ByRef dblYDataSmoothed() As Double, ByVal lngDataCount As Long, ByRef udtFindPeaksOptions As udtFindPeaksOptionsType)

    Const IgnoreNonPositiveData As Boolean = True

    Dim lngIndex As Long
    Dim lngPeakWidthPoints As Long
    
    Dim lngIndexLeftStart As Long
    Dim lngIndexLeftEnd As Long
    
    Dim lngIndexRightStart As Long
    Dim lngIndexRightEnd As Long
    
    Dim lngBinCountInPeak As Long
    
    Dim dblMeanLeft As Double
    Dim dblMeanRight As Double
    Dim dblMeanDiff As Double           ' Absolute value of difference between dblMeanLeft and dblMeanRight
    Dim dblMinMean As Double            ' Minimum of dblMeanLeft and dblMeanRight
    
    Dim dblTruePositiveArea As Double
    Dim dblFalsePositiveArea As Double

    ' Only use a portion of the data to compute the background levels
    ' The number of points to extend from the left and right is based on the peak width divided by 2
    lngPeakWidthPoints = udtPeak.IndexBaseRight - udtPeak.IndexBaseLeft
    If lngPeakWidthPoints < udtFindPeaksOptions.PeakWidthPointsMinimum Then
        lngPeakWidthPoints = udtFindPeaksOptions.PeakWidthPointsMinimum
    End If
    
    lngIndexLeftStart = udtPeak.IndexBaseLeft - lngPeakWidthPoints
    lngIndexLeftEnd = udtPeak.IndexBaseLeft - 1
    If lngIndexLeftStart < 0 Then lngIndexLeftStart = 0
    If lngIndexLeftEnd < 0 Then lngIndexLeftEnd = 0
    
    lngIndexRightStart = udtPeak.IndexBaseRight + 1
    lngIndexRightEnd = udtPeak.IndexBaseRight + lngPeakWidthPoints
    If lngIndexRightStart >= lngDataCount Then lngIndexRightStart = lngDataCount - 1
    If lngIndexRightEnd >= lngDataCount Then lngIndexRightEnd = lngDataCount - 1
    
'    For lngIndex = 0 To lngDataCount - 1
'        Debug.Print dblYData(lngIndex) & "," & dblYDataSmoothed(lngIndex)
'    Next lngIndex
    
    ' Compute the mean value of the data left of the peak
    dblMeanLeft = ComputeMeanInRange(dblYDataSmoothed, lngIndexLeftStart, lngIndexLeftEnd, IgnoreNonPositiveData)
    
    ' Compute the mean value of the data right of the peak
    dblMeanRight = ComputeMeanInRange(dblYDataSmoothed, lngIndexRightStart, lngIndexRightEnd, IgnoreNonPositiveData)
    
    ' For extremely narrow peaks, the smoothed peak can be wider than the true Y data
    ' Therefore, we may need to adjust udtPeak.IndexBaseLeft and udtPeak.IndexBaseRight to the point
    '  where the first non-negative value exists
    Do While udtPeak.IndexBaseLeft < udtPeak.IndexOfMaximum
        If dblYData(udtPeak.IndexBaseLeft) <= 0 Then
            udtPeak.IndexBaseLeft = udtPeak.IndexBaseLeft + 1
        Else
            Exit Do
        End If
    Loop
    
    Do While udtPeak.IndexBaseRight > udtPeak.IndexOfMaximum
        If dblYData(udtPeak.IndexBaseRight) <= 0 Then
            udtPeak.IndexBaseRight = udtPeak.IndexBaseRight - 1
        Else
            Exit Do
        End If
    Loop
    
    ' Compute the false positive area
    ' Determine the number of bins in the peak
    lngBinCountInPeak = udtPeak.IndexBaseRight - udtPeak.IndexBaseLeft + 1
    
    dblMeanDiff = Abs(dblMeanLeft - dblMeanRight)
    dblMinMean = dblMeanLeft
    If dblMeanRight < dblMinMean Then dblMinMean = dblMeanRight
    
    ' The area under the peak that accounts for the false positive data is a trapezoid
    dblFalsePositiveArea = lngBinCountInPeak * dblMinMean + 0.5 * lngBinCountInPeak * dblMeanDiff
        
    ' Compute the true positive area by summing the data intensities within the peak (using the true Y data)
    '  then subtracting dblFalsePositiveArea
    dblTruePositiveArea = 0
    For lngIndex = udtPeak.IndexBaseLeft To udtPeak.IndexBaseRight
        dblTruePositiveArea = dblTruePositiveArea + dblYData(lngIndex)
    Next lngIndex
    dblTruePositiveArea = dblTruePositiveArea - dblFalsePositiveArea
    
    If dblTruePositiveArea < 0 Then
        Debug.Assert False
        ' This shouldn't happen
        dblTruePositiveArea = 0
    End If
        
    With udtPeak
        .TruePositiveArea = CLng(dblTruePositiveArea)
        .FalsePositiveArea = CLng(dblFalsePositiveArea)
    End With
End Sub


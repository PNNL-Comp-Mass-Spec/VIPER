VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSplitUMCsByAbundance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtUMCPeakSplitType
    PeakStartPoint As Long
    PeakEndPoint As Long
    MemberIndexStartPoint As Long
    MemberIndexEndPoint As Long
    MedianMass As Double
    StDevMass As Double
    Valid As Boolean
End Type

Private mAbortRequested As Boolean

Private Sub Class_Initialize()
    mAbortRequested = False
End Sub

Public Sub AbortProcessingNow()
    mAbortRequested = True
End Sub

Public Sub ExamineUMCs(lngGelIndex As Long, frmCallingForm As VB.Form, Optional blnShowMessages As Boolean = True, Optional blnUseStatusFunctionOnCallingForm As Boolean = False)
    
    ' Examines the intensities of the ions in each UMC
    ' Looks for UMC's that have two or more distinct peaks in a localized TIC plot (total ion chromatogram)
    ' If multiple members in a UMC are in the same scan, then can either combine their intensities, or use the intensity of the highest ion
    ' If a UMC is found to have two or more peaks, then splits the UMC into 2 (or more parts), proportioning
    '  the members along a scan-based division point
    ' Note: Be sure the UMC Statistics are up to date (by calling UpdateUMCStatArrays) before calling this function
    ' If blnUseStatusFunctionOnCallingForm = True, then calls frmCallingForm.Status for the status messages
    '
    ' In addition, if .GapMaxSize is > 0, then will also split UMC's with gaps larger than .GapMaxSize, regardless of intensity or mass
    
    Dim lngUMCCountAtStart As Long
    Dim lngUMCIndex As Long
    Dim lngScanIndex As Long
    Dim lngCompareIndex As Long
    
    Dim lngMemberScanNumbersRelative() As Long      ' Scan numbers of the members of each UMC (relative, since we call LookupScanNumberRelativeIndex() to lookup the index in .ScanInfo())
    Dim lngClassMemberCount As Long
    
    Dim dblXData() As Double        ' Scan numbers: note, the peak finder algorithm expects evenly spaced data; thus, must use zero's for scans without data
    Dim dblYData() As Double        ' Sum of the intensity of all of the ions falling in the given scan
    Dim lngDataCount As Long
    
    Dim lngMinimumClassCountToCheck As Long
    Dim lngPeakHalfWidthPoints As Long
    
    Dim lngPeakCount As Long
    Dim lngPeakLocations() As Long                  ' 0-based array
    Dim lngPeakEdgesLeft() As Long                  ' 0-based array, parallel with lngPeakLocations(); Not used in this function, though we pass it to objPeakFinder.DetectPeaks
    Dim lngPeakEdgesRight() As Long                 ' 0-based array, parallel with lngPeakLocations(); Not used in this function, though we pass it to objPeakFinder.DetectPeaks
    Dim dblPeakAreas() As Double                    ' 0-based array, parallel with lngPeakLocations(); Not used in this function, though we pass it to objPeakFinder.DetectPeaks
    Dim udtPeakLocs() As udtUMCPeakSplitType        ' 0-based array, parallel with lngPeakLocations()
    
    Dim lngPeakIndex As Long, lngSubIndex As Long
    Dim lngSwapVal As Long
    
    Dim blnSuccess As Boolean
    Dim lngUMCsSplit As Long, lngUMCsAdded As Long
    
    Dim strMessage As String
    Dim strCaptionSaved As String
    
    Dim udtOptions As udtSplitUMCsByAbundanceOptionsType
    Dim objPeakFinder As New clsPeakDetection
    
On Error GoTo SplitUMCsByAbundanceErrorHandler

    mAbortRequested = False

    ' Validate some of the peak-picking parameters
    With glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions
        If .PeakWidthPointsMinimum < 3 Then .PeakWidthPointsMinimum = 4
        If .PeakWidthInSigma < 1 Then .PeakWidthInSigma = 3
        
        If .PeakDetectIntensityThresholdPercentageOfMaximum < 1 Then .PeakDetectIntensityThresholdPercentageOfMaximum = 1
        
        If .MaximumPeakCountToSplitUMC < 2 Then .MaximumPeakCountToSplitUMC = 2
    
        lngPeakHalfWidthPoints = CInt(.PeakWidthPointsMinimum / 2#)
    End With
    
    ' Copy the settings to udtOptions (to make the code a bit cleaner)
    udtOptions = glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions
    
    ' Minimum UMC size must be at least udtOptions.PeakWidthPointsMinimum * 2
    lngMinimumClassCountToCheck = udtOptions.PeakWidthPointsMinimum * 2
    
    lngUMCCountAtStart = GelUMC(lngGelIndex).UMCCnt
    lngUMCsAdded = 0
    lngUMCsSplit = 0
    
    strCaptionSaved = frmCallingForm.Caption
    strMessage = "Finding UMC's needing to be split: 0 / " & Trim(lngUMCCountAtStart)
    If blnUseStatusFunctionOnCallingForm Then
        frmCallingForm.Status strMessage
    Else
        frmCallingForm.Caption = strMessage
    End If
    
    
    If GelSearchDef(lngGelIndex).UMCDef.GapMaxSize > 0 And _
       glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.ScanGapBehavior <> susgIgnoreScanGaps Then
       
        ' First split any UMC's with gaps that are too large
        ' This should only affect UMC's found using UMCIonNet, and not UMCListType2002 or UMC2003
        For lngUMCIndex = 0 To lngUMCCountAtStart - 1
            If GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassCount > GelSearchDef(lngGelIndex).UMCDef.GapMaxSize Then
                SplitUMCsPopulateArrays lngGelIndex, lngUMCIndex, dblXData(), dblYData(), lngDataCount, lngMemberScanNumbersRelative(), lngClassMemberCount
                    
                ' Look for gaps
                ' Record their presence by recording each group as a peak in udtPeakLocs
                
                ' At most, there will be lngClassMemberCount peaks
                ReDim udtPeakLocs(0 To lngClassMemberCount - 1)
                
                ' The scan gap based peak finder doesn't use lngPeakLocations
                ' However, we'll populate it anyway for completeness since SplitUMCsComputeMedianMass needs to be passed the array
                ReDim lngPeakLocations(0 To lngClassMemberCount - 1)
                
                lngPeakCount = 0
                    
                For lngScanIndex = 0 To lngDataCount - 2
                    lngCompareIndex = lngScanIndex + 1
                    Do While dblYData(lngCompareIndex) = 0
                        lngCompareIndex = lngCompareIndex + 1
                    Loop
                    
                    If lngCompareIndex - lngScanIndex - 1 > GelSearchDef(lngGelIndex).UMCDef.GapMaxSize Then
                        If lngPeakCount = 0 Then
                            ' Need to add the first peak
                            With udtPeakLocs(0)
                                .PeakStartPoint = 0
                            End With
                            lngPeakCount = 1
                        End If
                        
                        With udtPeakLocs(lngPeakCount - 1)
                            .PeakEndPoint = lngCompareIndex - 1
                            ' Define the peak location to be half way between .PeakStartPoint and lngScanIndex,
                            '  since the data between lngScanIndex + 1 and lngCompareIndex all has an intensity of 0
                            lngPeakLocations(lngPeakCount - 1) = (.PeakStartPoint + lngScanIndex) / 2
                        End With
                        
                        udtPeakLocs(lngPeakCount).PeakStartPoint = lngCompareIndex
                        lngPeakCount = lngPeakCount + 1
                        lngScanIndex = lngCompareIndex - 1
                    End If
                Next lngScanIndex
                
                If lngPeakCount > 1 Then
                    With udtPeakLocs(lngPeakCount - 1)
                        .PeakEndPoint = lngDataCount - 1
                        lngPeakLocations(lngPeakCount - 1) = (.PeakEndPoint + .PeakStartPoint) / 2
                    End With
                    
                    ' I believe only UMC Ion Net will produce UMC's that need to be split using scan gaps
                    ' Still, if .def.UMCType is not glUMC_TYPE_FROM_NET, then that is OK too
                    Debug.Assert GelUMC(lngGelIndex).def.UMCType = glUMC_TYPE_FROM_NET
                    
                    ' Now expand the peaks until their StartPoint and EndPoint values touch
                    SplitUMCsWidenAdjacentPeaks udtPeakLocs(), lngPeakCount, dblYData()
                    
                    SplitUMCsLookupMemberIndices udtPeakLocs(), lngPeakCount, dblXData(), lngMemberScanNumbersRelative(), lngClassMemberCount
                    
                    If glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.ScanGapBehavior = susgSplitIfMassDifference Then
                        ' Combine the peaks that have similar masses

                        ' Now compute the average mass of the points in each sub-peak
                        SplitUMCsComputeMedianMass lngGelIndex, lngUMCIndex, udtPeakLocs(), lngPeakCount, lngPeakLocations(), udtOptions.MinimumDifferenceInAveragePpmMassToSplit
                        If lngPeakCount > 0 Then
                            Debug.Assert udtPeakLocs(lngPeakCount - 1).MemberIndexEndPoint = lngClassMemberCount - 1
                        End If
                    End If
                    
                    ' Actually split the UMC
                    If SplitUMCsDoSplitting(lngGelIndex, lngUMCIndex, udtPeakLocs(), lngPeakCount, lngMemberScanNumbersRelative, lngUMCsAdded) Then
                        lngUMCsSplit = lngUMCsSplit + 1
                    End If
                End If
                
            End If
            
            If lngUMCIndex Mod 50 = 0 Then
                strMessage = "Splitting UMC's with gaps that are too large: " & Trim(lngUMCIndex) & " / " & Trim(lngUMCCountAtStart)
                If blnUseStatusFunctionOnCallingForm Then
                    frmCallingForm.Status strMessage
                Else
                    frmCallingForm.Caption = strMessage
                End If
                DoEvents
            End If
            
            If mAbortRequested Then Exit For
        Next lngUMCIndex
    End If
    
    If Not mAbortRequested Then
        ' Now see if any of the UMC's need to be split due to subregions with differing abundances and mass
        For lngUMCIndex = 0 To lngUMCCountAtStart - 1
            If GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassCount >= lngMinimumClassCountToCheck Then
                
                SplitUMCsPopulateArrays lngGelIndex, lngUMCIndex, dblXData(), dblYData(), lngDataCount, lngMemberScanNumbersRelative(), lngClassMemberCount
                
                With udtOptions
                    lngPeakCount = objPeakFinder.DetectPeaks(dblXData(), dblYData(), lngDataCount, _
                                                            .PeakDetectIntensityThresholdAbsoluteMinimum, _
                                                            .PeakWidthPointsMinimum, _
                                                            lngPeakLocations(), lngPeakEdgesLeft(), lngPeakEdgesRight(), _
                                                            dblPeakAreas(), _
                                                            .PeakDetectIntensityThresholdPercentageOfMaximum, _
                                                            .PeakWidthInSigma, False, True)
                End With
                
                If lngPeakCount > 1 And lngPeakCount <= udtOptions.MaximumPeakCountToSplitUMC Then
                
                    ' Possibly Split the UMC into multiple UMC's
                    ' First, find the start and end of each peak
                    ' Next, find the average mass of the points in each peak
                    ' Then, find those peaks that have average masses significantly different than the adjacent peak
                    
                    ReDim udtPeakLocs(0 To lngPeakCount - 1)
                    
                    For lngPeakIndex = lngPeakCount - 1 To 1 Step -1
                        ' Make sure this peak and the previous peak are at least udtOptions.PeakWidthPointsMinimum scans apart
                        ' In addition, make sure this peak's apex still points to a valid index in the
                        '  current UMC; this could happen to one of the internal peaks if one of the other peaks included it as part of a split
                        
                        udtPeakLocs(lngPeakIndex).Valid = False
                        If (lngPeakLocations(lngPeakIndex) - lngPeakLocations(lngPeakIndex - 1) >= udtOptions.PeakWidthPointsMinimum) Then
                        
                            If lngPeakIndex = lngPeakCount - 1 Then
                                udtPeakLocs(lngPeakIndex).Valid = True
                                udtPeakLocs(lngPeakIndex).PeakEndPoint = lngClassMemberCount - 1
                            Else
                                If lngPeakLocations(lngPeakIndex) < udtPeakLocs(lngPeakIndex + 1).PeakStartPoint Then
                                    udtPeakLocs(lngPeakIndex).Valid = True
                                    udtPeakLocs(lngPeakIndex).PeakEndPoint = udtPeakLocs(lngPeakIndex + 1).PeakStartPoint - 1
                                Else
                                    udtPeakLocs(lngPeakIndex).Valid = False
                                End If
                            End If
                        
                            If udtPeakLocs(lngPeakIndex).Valid Then
                                ' Define the start point for this peak
                                udtPeakLocs(lngPeakIndex).PeakStartPoint = lngPeakLocations(lngPeakIndex) - lngPeakHalfWidthPoints
                                
                                ' Define the end point for the next peak to the left of this peak
                                udtPeakLocs(lngPeakIndex - 1).PeakEndPoint = lngPeakLocations(lngPeakIndex - 1) + lngPeakHalfWidthPoints
                                
                                If udtPeakLocs(lngPeakIndex - 1).PeakEndPoint > udtPeakLocs(lngPeakIndex).PeakStartPoint Then
                                    ' The peaks overlap
                                    ' Swap the start and end indices
                                    lngSwapVal = udtPeakLocs(lngPeakIndex).PeakStartPoint
                                    udtPeakLocs(lngPeakIndex).PeakStartPoint = udtPeakLocs(lngPeakIndex - 1).PeakEndPoint
                                    udtPeakLocs(lngPeakIndex - 1).PeakEndPoint = lngSwapVal
                                ElseIf udtPeakLocs(lngPeakIndex - 1).PeakEndPoint = udtPeakLocs(lngPeakIndex).PeakStartPoint Then
                                    udtPeakLocs(lngPeakIndex).PeakStartPoint = udtPeakLocs(lngPeakIndex).PeakStartPoint + 1
                                End If
                            End If
                        End If
                    Next lngPeakIndex
                    
                    With udtPeakLocs(0)
                        .PeakStartPoint = 0
                        If .PeakEndPoint = 0 Then
                            .PeakEndPoint = lngPeakLocations(0) + lngPeakHalfWidthPoints
                            For lngSubIndex = 1 To lngPeakCount - 1
                                If udtPeakLocs(lngSubIndex).PeakStartPoint > 0 Then
                                    If .PeakEndPoint >= udtPeakLocs(lngSubIndex).PeakStartPoint Then
                                        udtPeakLocs(lngSubIndex).PeakStartPoint = .PeakEndPoint + 1
                                    End If
                                    Exit For
                                End If
                            Next lngSubIndex
                        End If
                        .Valid = True
                    End With
                    
                    lngPeakIndex = 1
                    Do While lngPeakIndex <= lngPeakCount - 1
                        If Not udtPeakLocs(lngPeakIndex).Valid Then
                            ' Remove this peak from lngPeakLocations() and udtPeakLocs()
                            For lngSubIndex = lngPeakIndex To lngPeakCount - 2
                                lngPeakLocations(lngSubIndex) = lngPeakLocations(lngSubIndex + 1)
                                udtPeakLocs(lngSubIndex) = udtPeakLocs(lngSubIndex + 1)
                            Next lngSubIndex
                            lngPeakCount = lngPeakCount - 1
                            
                            If lngPeakIndex <= lngPeakCount - 1 Then
                                With udtPeakLocs(lngPeakIndex)
                                    ' Expand the peak to start one point to the right of the next peak to the left
                                    .PeakStartPoint = udtPeakLocs(lngPeakIndex - 1).PeakEndPoint + 1
                                End With
                            End If
                        Else
                            lngPeakIndex = lngPeakIndex + 1
                        End If
                    Loop
                    
                    If lngPeakCount > 1 Then
                        
                        ' Now expand the peaks until their StartPoint and EndPoint values touch
                        SplitUMCsWidenAdjacentPeaks udtPeakLocs(), lngPeakCount, dblYData()
                        
                        SplitUMCsLookupMemberIndices udtPeakLocs(), lngPeakCount, dblXData(), lngMemberScanNumbersRelative(), lngClassMemberCount
                                                
                        ' Now compute the average mass of the points in each sub-peak
                        SplitUMCsComputeMedianMass lngGelIndex, lngUMCIndex, udtPeakLocs(), lngPeakCount, lngPeakLocations(), udtOptions.MinimumDifferenceInAveragePpmMassToSplit
                        If lngPeakCount > 0 Then
                            Debug.Assert udtPeakLocs(lngPeakCount - 1).MemberIndexEndPoint = lngClassMemberCount - 1
                        End If
                        
                        If lngPeakCount > 1 Then
                            ' Actually split the UMC
                            If SplitUMCsDoSplitting(lngGelIndex, lngUMCIndex, udtPeakLocs(), lngPeakCount, lngMemberScanNumbersRelative, lngUMCsAdded) Then
                                lngUMCsSplit = lngUMCsSplit + 1
                            End If
                        End If
                    End If
                    
                End If
            End If
            
            If lngUMCIndex Mod 50 = 0 Then
                strMessage = "Splitting UMC's with multiple abundance profiles: " & Trim(lngUMCIndex) & " / " & Trim(lngUMCCountAtStart)
                If blnUseStatusFunctionOnCallingForm Then
                    frmCallingForm.Status strMessage
                Else
                    frmCallingForm.Caption = strMessage
                End If
                DoEvents
            End If
            
            If mAbortRequested Then Exit For
        Next lngUMCIndex
    End If
    
    If lngUMCsAdded > 0 Or lngUMCsSplit > 0 Then
        ' Need to recompute the UMC Statistic arrays and store the updated Class Representative Mass
        blnSuccess = UpdateUMCStatArrays(lngGelIndex, False, frmCallingForm)
        Debug.Assert blnSuccess
    
        frmCallingForm.Caption = strCaptionSaved

        strMessage = "Split UMC's with separate regions; UMC Count before splitting = " & lngUMCCountAtStart & "; UMC's split = " & lngUMCsSplit & "; UMC's Added = " & lngUMCsAdded
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled And blnShowMessages Then
            MsgBox strMessage, vbInformation + vbOKOnly, "Done"
        End If
        AddToAnalysisHistory lngGelIndex, strMessage
    
        ' Possibly Auto-Refine the UMC's
        AutoRefineUMCs lngGelIndex, frmCallingForm
    
    Else
        frmCallingForm.Caption = strCaptionSaved
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled And blnShowMessages Then
            MsgBox "No UMC's were found that needed splitting.", vbInformation + vbOKOnly, "Done"
        End If
    End If
    
    Set objPeakFinder = Nothing
    
    Exit Sub
    
SplitUMCsByAbundanceErrorHandler:
    Debug.Print "Error in SplitUMCsByAbundance: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "SplitUMCsBybundance"

End Sub

Private Sub SplitUMCsComputeMedianMass(lngGelIndex As Long, lngUMCIndex As Long, ByRef udtPeakLocs() As udtUMCPeakSplitType, ByRef lngPeakCount As Long, lngPeakLocations() As Long, dblMinimumDifferenceInMedianPpmMassToSplit As Double)

    Dim lngPeakIndex As Long
    Dim lngSubIndex As Long
    
    Dim lngMemberIndex As Long
    Dim lngPeakMemberCount As Long
    
    Dim dblMassSum As Double, dblMassDiff As Double

    Dim blnProceed As Boolean
    Dim blnCombinePeaks As Boolean
    
    Dim objStats As New StatDoubles
    Dim dblMassData() As Double
    
On Error GoTo SplitUMCsComputeAverageMassErrorHandler

    ' Only split if the median mass is at least dblMinimumDifferenceInMedianPpmMassToSplit away from the previous peak
    ' Also make sure the mass difference is at least as large as the StDev of the masses in each peak
    Do
        blnProceed = True
        ' Compute the median mass and StDev in mass for each peak
        With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
            For lngPeakIndex = 0 To lngPeakCount - 1
                dblMassSum = 0
                
                lngPeakMemberCount = udtPeakLocs(lngPeakIndex).MemberIndexEndPoint - udtPeakLocs(lngPeakIndex).MemberIndexStartPoint + 1
                
                If lngPeakIndex > 0 Then
                    Debug.Assert udtPeakLocs(lngPeakIndex).MemberIndexStartPoint = udtPeakLocs(lngPeakIndex - 1).MemberIndexEndPoint + 1
                End If
                
                ReDim dblMassData(lngPeakMemberCount - 1)
                For lngMemberIndex = udtPeakLocs(lngPeakIndex).MemberIndexStartPoint To udtPeakLocs(lngPeakIndex).MemberIndexEndPoint
                    Select Case .ClassMType(lngMemberIndex)
                    Case gldtCS
                        dblMassData(lngMemberIndex - udtPeakLocs(lngPeakIndex).MemberIndexStartPoint) = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).AverageMW
                    Case gldtIS
                        dblMassData(lngMemberIndex - udtPeakLocs(lngPeakIndex).MemberIndexStartPoint) = GetIsoMass(GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)), GelData(lngGelIndex).Preferences.IsoDataField)
                    Case Else
                        ' This shouldn't happen; ignore it
                        Debug.Assert False
                    End Select
                Next lngMemberIndex
                
                ' Compute the median and StDev
                If objStats.Fill(dblMassData) Then
                    udtPeakLocs(lngPeakIndex).MedianMass = objStats.Median
                    udtPeakLocs(lngPeakIndex).StDevMass = objStats.StDev
                End If
            
            Next lngPeakIndex
        End With
        
        ' Compare the computed median mass of each peak to the adjacent peak
        ' If the mass differences are too small, combine the peaks, then set blnProceed = False
        '  so that the median masses are re-computed
        lngPeakIndex = 0
        Do While lngPeakIndex <= lngPeakCount - 2
            dblMassDiff = Abs(udtPeakLocs(lngPeakIndex).MedianMass - udtPeakLocs(lngPeakIndex + 1).MedianMass)
            
            If MassToPPM(dblMassDiff, udtPeakLocs(lngPeakIndex).MedianMass) < dblMinimumDifferenceInMedianPpmMassToSplit Then
                ' Mass difference is too small; combine the peaks
                blnCombinePeaks = True
            Else
                ' Make sure mass difference is at least as large as the StDev in mass of each peak
                If dblMassDiff < udtPeakLocs(lngPeakIndex).StDevMass Or dblMassDiff < udtPeakLocs(lngPeakIndex + 1).StDevMass Then
                    ' Mass difference is too small to be significant, compared to the StDev values
                    ' Combine the peaks
                    blnCombinePeaks = True
                Else
                    ' Split the peaks
                    blnCombinePeaks = False
                End If
            End If
            
            If blnCombinePeaks Then
                With udtPeakLocs(lngPeakIndex)
                    .PeakEndPoint = udtPeakLocs(lngPeakIndex + 1).PeakEndPoint
                    .MemberIndexEndPoint = udtPeakLocs(lngPeakIndex + 1).MemberIndexEndPoint
                End With
                
                For lngSubIndex = lngPeakIndex + 1 To lngPeakCount - 2
                    lngPeakLocations(lngSubIndex) = lngPeakLocations(lngSubIndex + 1)
                    udtPeakLocs(lngSubIndex) = udtPeakLocs(lngSubIndex + 1)
                Next lngSubIndex
                lngPeakCount = lngPeakCount - 1
                
                ' Set blnProceed to false, then Exit the Do Loop and re-compute the median masses
                blnProceed = False
                Exit Do
            Else
                lngPeakIndex = lngPeakIndex + 1
            End If
        Loop
        
    Loop While Not blnProceed And lngPeakCount > 1
    
    Exit Sub
    
SplitUMCsComputeAverageMassErrorHandler:
    Debug.Print "Error in SplitUMCsComputeMedianMass: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "SplitUMCsComputeMedianMass"

End Sub

Private Function SplitUMCsDoSplitting(ByVal lngGelIndex As Long, ByVal lngUMCIndex As Long, ByRef udtPeakLocs() As udtUMCPeakSplitType, lngPeakCount As Long, ByRef lngMemberScanNumbersRelative() As Long, ByRef lngUMCsAdded As Long) As Boolean

    Dim lngPeakIndex As Long
    Dim lngSplitIndexMemberIndex As Long
    Dim udtNewUMC As udtUMCType

    Dim lngMemberIndex As Long
    Dim lngTargetIndex As Long
    Dim lngClassRepIndexInUMC As Long
    Dim intUMCRepresentativeType As Integer
    
    Dim eClassRepIonType As glDistType
    
    Dim blnProceed As Boolean
    Dim blnClassSplit As Boolean
    
On Error GoTo SplitUMCsDoSplittingErrorHandler

    blnClassSplit = False
    intUMCRepresentativeType = glbPreferencesExpanded.UMCIonNetOptions.UMCRepresentative
    
    ' Work from right to left
    ' If lngPeakCount is just 1, then nothing will happen
    For lngPeakIndex = lngPeakCount - 1 To 1 Step -1
    
        blnProceed = False
        
        With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
            lngSplitIndexMemberIndex = udtPeakLocs(lngPeakIndex).MemberIndexStartPoint
            If lngSplitIndexMemberIndex > 0 Then
                ' Move the members from the current UMC to udtNewUMC
                
                udtNewUMC.ClassCount = .ClassCount - lngSplitIndexMemberIndex
                ReDim udtNewUMC.ClassMInd(0 To udtNewUMC.ClassCount - 1)
                ReDim udtNewUMC.ClassMType(0 To udtNewUMC.ClassCount - 1)
                
                ' Set these to 0 for now
                udtNewUMC.ClassAbundance = 0
                udtNewUMC.ClassMW = 0
                udtNewUMC.ClassMWStD = 0
                udtNewUMC.ClassRepInd = 0
                udtNewUMC.ClassRepType = gldtIS
                
                udtNewUMC.MinMW = 0
                udtNewUMC.MaxMW = 0
                udtNewUMC.MinScan = 0       ' The scan numbers recorded here will be relative scan numbers, not actual ones
                udtNewUMC.MaxScan = 0       ' The scan numbers recorded here will be relative scan numbers, not actual ones
                
                ' Set .ClassStatusBits = UMC_INDICATOR_BIT_SPLIT_UMC = 16 to record that this was a split UMC
                udtNewUMC.ClassStatusBits = UMC_INDICATOR_BIT_SPLIT_UMC
                
                For lngMemberIndex = lngSplitIndexMemberIndex To .ClassCount - 1
                    lngTargetIndex = lngMemberIndex - lngSplitIndexMemberIndex
                    
                    udtNewUMC.ClassMInd(lngTargetIndex) = .ClassMInd(lngMemberIndex)
                    udtNewUMC.ClassMType(lngTargetIndex) = .ClassMType(lngMemberIndex)
                    
                    If lngMemberIndex = lngSplitIndexMemberIndex Then
                        udtNewUMC.MinScan = lngMemberScanNumbersRelative(lngMemberIndex)
                        udtNewUMC.MaxScan = udtNewUMC.MinScan
                    Else
                        If lngMemberScanNumbersRelative(lngMemberIndex) < udtNewUMC.MinScan Then udtNewUMC.MinScan = lngMemberScanNumbersRelative(lngMemberIndex)
                        If lngMemberScanNumbersRelative(lngMemberIndex) > udtNewUMC.MaxScan Then udtNewUMC.MaxScan = lngMemberScanNumbersRelative(lngMemberIndex)
                    End If
                Next lngMemberIndex
                    
                blnProceed = True
            Else
                ' This shouldn't happen
                Debug.Assert False
            End If
            
            If blnProceed Then
                ' Remove the points from the current UMC
                .ClassCount = lngSplitIndexMemberIndex
                
                ' Update the MaxScan value
                ' Need to dereference from the relative value to the true value using .ScanInfo()
                .MaxScan = GelData(lngGelIndex).ScanInfo(lngMemberScanNumbersRelative(lngSplitIndexMemberIndex - 1)).ScanNumber
                
                ReDim Preserve .ClassMInd(0 To .ClassCount - 1)
                ReDim Preserve .ClassMType(0 To .ClassCount - 1)
                
                ' Reset the various stats for this UMC
                .ClassAbundance = 0
                .ClassMW = 0
                .ClassMWStD = 0
                .ClassRepInd = 0
                .ClassRepType = gldtIS
                
                ' Set .ClassStatusBits = UMC_INDICATOR_BIT_SPLIT_UMC = 16 to record that this was a split UMC
                .ClassStatusBits = UMC_INDICATOR_BIT_SPLIT_UMC
                
                .MinMW = 0
                .MaxMW = 0
                
                ' The class was split, so set this to true
                blnClassSplit = True
                
            End If
        End With
        
        If blnProceed Then
            ' If using UMCIonNet searching, and if .MakeSingleMemberClasses = False, then make sure .MemberCount is not 1
            If GelUMC(lngGelIndex).def.UMCType = glUMC_TYPE_FROM_NET And Not glbPreferencesExpanded.UMCIonNetOptions.MakeSingleMemberClasses Then
                If udtNewUMC.ClassCount = 1 Then
                    blnProceed = False
                End If
            End If
            
            ' If auto-refining is turned on, then see if the new UMC is too short
            If glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveCountLow Then
                If glbPreferencesExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange Then
                    If udtNewUMC.ClassCount < glbPreferencesExpanded.UMCAutoRefineOptions.MinMemberCountWhenUsingScanRange Then
                        blnProceed = False
                    ElseIf udtNewUMC.MaxScan - udtNewUMC.MinScan + 1 < glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineMinLength Then
                        blnProceed = False
                    End If
                Else
                    If udtNewUMC.ClassCount < glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineMinLength Then
                        blnProceed = False
                    End If
                End If
            End If
            
        End If
        
        If blnProceed Then
            With GelUMC(lngGelIndex)
                ' Add udtNewUMC to .UMCs()
                ' Need to update .MinScan and .MaxScan to be the correct scan numbers
                udtNewUMC.MinScan = GelData(lngGelIndex).ScanInfo(udtNewUMC.MinScan).ScanNumber
                udtNewUMC.MaxScan = GelData(lngGelIndex).ScanInfo(udtNewUMC.MaxScan).ScanNumber
                
                ReDim Preserve .UMCs(.UMCCnt)
                .UMCs(.UMCCnt) = udtNewUMC
                .UMCCnt = .UMCCnt + 1
            End With
            
            ' Find the index of the class representative
            lngClassRepIndexInUMC = FindUMCClassRepIndex(lngGelIndex, GelUMC(lngGelIndex).UMCCnt - 1, intUMCRepresentativeType, eClassRepIonType)
            
            With GelUMC(lngGelIndex).UMCs(GelUMC(lngGelIndex).UMCCnt - 1)
                .ClassRepInd = .ClassMInd(lngClassRepIndexInUMC)
                .ClassRepType = eClassRepIonType
            End With
            
            lngUMCsAdded = lngUMCsAdded + 1
            blnClassSplit = True
        End If
                            
    Next lngPeakIndex

    If blnClassSplit Then
        ' Need to re-determine the class representative for this UMC
        lngClassRepIndexInUMC = FindUMCClassRepIndex(lngGelIndex, lngUMCIndex, intUMCRepresentativeType, eClassRepIonType)
        With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
            .ClassRepInd = .ClassMInd(lngClassRepIndexInUMC)
            .ClassRepType = eClassRepIonType
        End With
    End If
    
    SplitUMCsDoSplitting = blnClassSplit
    Exit Function

SplitUMCsDoSplittingErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "SplitUMCsDoSplitting"
    Debug.Print "Error in SplitUMCsDoSplitting: " & Err.Description
    SplitUMCsDoSplitting = False

End Function

Private Sub SplitUMCsLookupMemberIndices(ByRef udtPeakLocs() As udtUMCPeakSplitType, ByVal lngPeakCount As Long, ByRef dblXData() As Double, ByRef lngMemberScanNumbersRelative() As Long, ByVal lngClassMemberCount As Long)
    
    Dim lngPreviousPeakMemberIndexStart As Long
    Dim lngPeakIndex As Long
    Dim lngMemberIndex As Long
    
    Dim blnSuccess As Boolean
    
    ' Determine the member indices corresponding to .PeakStartPoint and .PeakEndPoint in each peak
    ' Initialize lngPreviousPeakMemberIndexStart to be 1 more than the last valid member index (i.e. equal to lngClassMemberCount)
    lngPreviousPeakMemberIndexStart = lngClassMemberCount
    For lngPeakIndex = lngPeakCount - 1 To 1 Step -1
    
        With udtPeakLocs(lngPeakIndex)
            .MemberIndexEndPoint = lngPreviousPeakMemberIndexStart - 1
            
            blnSuccess = False
            For lngMemberIndex = 0 To .MemberIndexEndPoint
                If lngMemberScanNumbersRelative(lngMemberIndex) >= dblXData(.PeakStartPoint) Then
                    .MemberIndexStartPoint = lngMemberIndex
                    blnSuccess = True
                    Exit For
                End If
            Next lngMemberIndex
            
            If Not blnSuccess Then
                ' This shouldn't happen
                Debug.Assert False
                .MemberIndexStartPoint = 0
            End If
            
            lngPreviousPeakMemberIndexStart = .MemberIndexStartPoint
        End With
        
        If lngPreviousPeakMemberIndexStart <= 0 Then
            Debug.Assert False
            Exit For
        End If
    Next lngPeakIndex
    
    With udtPeakLocs(0)
        .MemberIndexStartPoint = 0
        .MemberIndexEndPoint = udtPeakLocs(1).MemberIndexStartPoint - 1
    End With
    
End Sub

Private Sub SplitUMCsPopulateArrays(ByVal lngGelIndex As Long, ByVal lngUMCIndex As Long, ByRef dblXData() As Double, ByRef dblYData() As Double, ByRef lngDataCount As Long, ByRef lngMemberScanNumbersRelative() As Long, ByRef lngClassMemberCount As Long)

    Dim lngScanIndex As Long
    Dim lngMemberIndex As Long
    
    Dim lngScanNumberStart As Long
    Dim lngCurrentScanNumber As Long
    Dim dblCurrentIntensity As Double


    With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
        
        lngClassMemberCount = .ClassCount
        
        ' For LTQ-FT data, there can be gaps between scan numbers due to MS/MS scans interleaved between the MS scans
        ' To account for this, we'll lookup the relative index for each scan number and use it, rather than the actual scan number
        lngScanNumberStart = LookupScanNumberRelativeIndex(lngGelIndex, .MinScan)
        lngDataCount = LookupScanNumberRelativeIndex(lngGelIndex, .MaxScan) - lngScanNumberStart + 1
        
        ' At most, we'll need space for lngDataCount scans
        ReDim dblXData(0 To lngDataCount - 1)
        ReDim dblYData(0 To lngDataCount - 1)
        
        ' Initialize dblXData()
        For lngScanIndex = 0 To lngDataCount - 1
            dblXData(lngScanIndex) = lngScanNumberStart + lngScanIndex
        Next lngScanIndex
        
        ' Initialize lngMemberScanNumbersRelative()
        ReDim lngMemberScanNumbersRelative(0 To .ClassCount - 1)
        
        For lngMemberIndex = 0 To .ClassCount - 1
            Select Case .ClassMType(lngMemberIndex)
            Case gldtCS
                lngCurrentScanNumber = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).ScanNumber
                dblCurrentIntensity = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).Abundance
            Case gldtIS
                lngCurrentScanNumber = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).ScanNumber
                dblCurrentIntensity = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).Abundance
            Case Else
                ' This shouldn't happen; ignore it
                Debug.Assert False
                lngCurrentScanNumber = -1
            End Select
            
            ' Add this ion's intensity to the appropriate index of dblXData() and dblYData()
            If lngCurrentScanNumber >= 0 Then
                lngCurrentScanNumber = LookupScanNumberRelativeIndex(lngGelIndex, lngCurrentScanNumber)
                
                lngScanIndex = lngCurrentScanNumber - lngScanNumberStart
                Debug.Assert dblXData(lngScanIndex) = lngCurrentScanNumber
                If dblCurrentIntensity < 0 Then
                    ' Does this ever happen?  I doubt it (January 2004)
                    ' We don't want to allow negative intensities since we determine whether any data
                    '  is present in the given scan by checking if dblYData() > 0
                    Debug.Assert False
                    dblCurrentIntensity = Abs(dblCurrentIntensity)
                ElseIf dblCurrentIntensity = 0 Then
                    ' Does this ever happen?  I doubt it (January 2004)
                    ' We don't want to allow intensities of 0 since we determine whether any data
                    '  is present in the given scan by checking if dblYData() > 0
                    dblCurrentIntensity = 1
                End If
                    
                dblYData(lngScanIndex) = dblYData(lngScanIndex) + dblCurrentIntensity
                
                lngMemberScanNumbersRelative(lngMemberIndex) = lngCurrentScanNumber
            End If
        Next lngMemberIndex
    End With

End Sub

Private Sub SplitUMCsWidenAdjacentPeaks(ByRef udtPeakLocs() As udtUMCPeakSplitType, lngPeakCount As Long, ByRef dblYData() As Double)
        
    Dim lngPeakIndex As Long
    Dim lngScanIndex As Long
    Dim lngSplitIndexScan As Long
    
    For lngPeakIndex = 0 To lngPeakCount - 2
        
        If udtPeakLocs(lngPeakIndex).PeakEndPoint >= udtPeakLocs(lngPeakIndex + 1).PeakStartPoint - 1 Then
            If udtPeakLocs(lngPeakIndex).PeakEndPoint = udtPeakLocs(lngPeakIndex + 1).PeakStartPoint - 1 Then
                ' The peaks already touch
            Else
                ' This peak overlaps the next one; this shouldn't happen
                ' There must be a problem with the code in the above For Loop (before the Do loop)
                Debug.Assert False
                udtPeakLocs(lngPeakIndex).PeakEndPoint = udtPeakLocs(lngPeakIndex + 1).PeakStartPoint - 1
            End If
        Else
            ' Look for the minimum between udtPeakLocs(lngPeakIndex).PeakEndPoint and udtPeakLocs(lngPeakIndex + 1).PeakStartPoint
            
            lngSplitIndexScan = udtPeakLocs(lngPeakIndex).PeakEndPoint
            For lngScanIndex = udtPeakLocs(lngPeakIndex).PeakEndPoint + 1 To udtPeakLocs(lngPeakIndex + 1).PeakStartPoint
                If dblYData(lngScanIndex) <= dblYData(lngSplitIndexScan) Then
                    lngSplitIndexScan = lngScanIndex
                End If
            Next lngScanIndex
            
            ' Update this peak to end at lngSplitIndexScan and the next to start at lngSplitIndexScan + 1
            udtPeakLocs(lngPeakIndex).PeakEndPoint = lngSplitIndexScan
            udtPeakLocs(lngPeakIndex + 1).PeakStartPoint = lngSplitIndexScan + 1
        End If
        
    Next lngPeakIndex

End Sub



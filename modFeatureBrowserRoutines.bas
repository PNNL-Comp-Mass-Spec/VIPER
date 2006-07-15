Attribute VB_Name = "modFeatureBrowserRoutines"
Option Explicit
    
Public Enum mruMassRangeUnitsConstants
    mruDa = 0
    mruPpm = 1
End Enum

Public Enum sruScanRangeUnitsConstants
    sruScan = 0
    sruNet = 1
End Enum
    
Public Function BrowseFeaturesDeleteSelected(lstFeatures As ListBox, ByRef InfoSortedPointerArray() As Long, ByRef blnFeatureValid() As Boolean, ByRef DeletedStackCount As Long, ByRef DeletedFeatureStack() As Long) As Boolean
    ' Note: This sub is very similar to BrowseFeaturesUnDeleteSelected
    
    Dim lngItemIndex As Long
    Dim lngDereferencedIndex As Long
    Dim blnUpdateListBox As Boolean
    
On Error GoTo DeleteSelectedFeatureErrorHandler
    
    ' Find each item selected in lstFeatures
    For lngItemIndex = 0 To lstFeatures.ListCount - 1
        If lstFeatures.Selected(lngItemIndex) Then
        
            lngDereferencedIndex = InfoSortedPointerArray(lngItemIndex)
            
            If Not blnFeatureValid(lngDereferencedIndex) Then
                ' UMC already deleted
            Else
                blnFeatureValid(lngDereferencedIndex) = False
            
                If DeletedStackCount > 0 Then
                    If DeletedFeatureStack(DeletedStackCount - 1) <> lngDereferencedIndex Then
                        ReDim Preserve DeletedFeatureStack(DeletedStackCount)
                        DeletedFeatureStack(DeletedStackCount) = lngDereferencedIndex
                        DeletedStackCount = DeletedStackCount + 1
                        blnUpdateListBox = True
                    End If
                Else
                    DeletedStackCount = 1
                    ReDim DeletedFeatureStack(0)
                    DeletedFeatureStack(0) = lngDereferencedIndex
                    blnUpdateListBox = True
                End If
            End If
        End If
    Next lngItemIndex
    
    BrowseFeaturesDeleteSelected = blnUpdateListBox
    
    Exit Function

DeleteSelectedFeatureErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "BrowseFeaturesDeleteSelected", Err.Description
    BrowseFeaturesDeleteSelected = False

End Function

Public Sub BrowseFeaturesPopulateUMCPlotData(lngGelIndex As Long, blnPlotAllChargeStates As Boolean, intChargesUsed() As Integer, ByVal lngUMCIndex As Long, ByRef lngDataPointCount As Long, ByRef dblScanData() As Double, ByRef dblAbuData() As Double, ByVal blnUseMaxValueEachScan As Boolean)
    ' Note that the arrays are 1-based, for compatibility reasons with the plot control
    
    Dim lngScanMin As Long
    Dim lngScanIndex As Long
    Dim lngScanCountNew As Long
    
    Dim lngGapSize As Long
    Dim lngMaxGapSize As Long
    Dim lngScanIndexCompare As Long
    
    Dim lngScanNumberRelative As Long
    
    Dim blnCopyDataPoint As Boolean
    Dim blnCopyingGapData As Boolean
    
    Dim intChargeIndex As Integer
    
    With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
        lngScanMin = .MinScan
        lngDataPointCount = .MaxScan - lngScanMin + 1
    End With
    
    ReDim dblScanData(1 To lngDataPointCount)
    ReDim dblAbuData(1 To lngDataPointCount)
      
    If blnPlotAllChargeStates Or intChargesUsed(0) = 0 Then
        ' Sum all charge states
        PopulateUMCAbuDataWork dblAbuData(), lngGelIndex, lngUMCIndex, 0, lngScanMin, False
    Else
        ' Only use the charge states listed in intChargesUsed
        For intChargeIndex = 0 To UBound(intChargesUsed())
            PopulateUMCAbuDataWork dblAbuData(), lngGelIndex, lngUMCIndex, intChargesUsed(intChargeIndex), lngScanMin, blnUseMaxValueEachScan
        Next intChargeIndex
    End If
    
    ' Interpolate the abundances between non-zero data points that have
    ' scan gaps > 1 but, according to .ScanInfo(), are actually adjacent scans
    InterpolateChromatogramGaps lngGelIndex, dblAbuData(), 1, lngDataPointCount, lngScanMin
    
    
    ' Need to populate dblScanData() with scan numbers since we'll be removing unwanted scans next
    For lngScanIndex = 1 To lngDataPointCount
        dblScanData(lngScanIndex) = lngScanIndex + lngScanMin - 1
    Next lngScanIndex
    
    ' Remove the points with an abundance of 0, provided the gap size is less than .InterpolateMaxGapSize
    lngMaxGapSize = GelUMC(lngGelIndex).def.InterpolateMaxGapSize
    
    lngScanCountNew = 0
    blnCopyingGapData = False
    For lngScanIndex = 1 To lngDataPointCount
        If dblAbuData(lngScanIndex) = 0 Then
            If blnCopyingGapData Then
                blnCopyDataPoint = True
            Else
                If lngScanIndex = 1 Or lngScanIndex = lngDataPointCount Then
                    blnCopyDataPoint = True
                Else
                    ' Set lngGapSize to a large value to start with
                    lngGapSize = lngDataPointCount - lngScanIndex + 1
                    
                    lngScanNumberRelative = LookupScanNumberRelativeIndex(lngGelIndex, CLng(dblScanData(lngScanIndex)))
                    
                    ' Now step through the points, looking for the next non-zero point
                    blnCopyDataPoint = False
                    For lngScanIndexCompare = lngScanIndex + 1 To lngDataPointCount
                        If dblAbuData(lngScanIndexCompare) > 0 Then
                            ' Find the gap distance to the next scan with data
                            lngGapSize = LookupScanNumberRelativeIndex(lngGelIndex, CLng(dblScanData(lngScanIndexCompare))) - lngScanNumberRelative
                            Exit For
                        End If
                    Next lngScanIndexCompare
                
                    If lngGapSize > lngMaxGapSize Then
                        blnCopyingGapData = True
                        blnCopyDataPoint = True
                    End If
                End If
            End If
        Else
            blnCopyDataPoint = True
            blnCopyingGapData = False
        End If
       
        If blnCopyDataPoint Then
            lngScanCountNew = lngScanCountNew + 1
            dblScanData(lngScanCountNew) = dblScanData(lngScanIndex)
            dblAbuData(lngScanCountNew) = dblAbuData(lngScanIndex)
        End If
    Next lngScanIndex
    
    If lngScanCountNew <= 0 Then
        ReDim dblScanData(1 To 1)
        ReDim dblAbuData(1 To 1)
        lngScanCountNew = 0
    Else
        ' Make sure there is a zero at the beginning of the array
        If dblAbuData(1) <> 0 Then
            lngScanCountNew = lngScanCountNew + 1
            If lngScanCountNew > lngDataPointCount Then
                ReDim Preserve dblScanData(1 To lngScanCountNew)
                ReDim Preserve dblAbuData(1 To lngScanCountNew)
            End If
            
            For lngScanIndex = lngScanCountNew To 2 Step -1
                dblAbuData(lngScanIndex) = dblAbuData(lngScanIndex - 1)
                dblScanData(lngScanIndex) = dblScanData(lngScanIndex - 1)
            Next lngScanIndex
    
            dblAbuData(1) = 0
            dblScanData(1) = dblScanData(2) - 1
        End If
        
        ' Make sure there is a zero at the end of the array
        If dblAbuData(lngScanCountNew) <> 0 Then
            lngScanCountNew = lngScanCountNew + 1
            If lngScanCountNew > lngDataPointCount Then
                ReDim Preserve dblScanData(1 To lngScanCountNew)
                ReDim Preserve dblAbuData(1 To lngScanCountNew)
            End If
            
            dblAbuData(lngScanCountNew) = 0
            dblScanData(lngScanCountNew) = dblScanData(lngScanCountNew - 1) + 1
        End If
        
        If lngScanCountNew < lngDataPointCount Then
            ReDim Preserve dblScanData(1 To lngScanCountNew)
            ReDim Preserve dblAbuData(1 To lngScanCountNew)
        End If
        lngDataPointCount = lngScanCountNew
    End If
    
End Sub

Private Sub PopulateUMCAbuDataWork(ByRef dblAbundance() As Double, ByVal lngGelIndex As Long, ByVal UMCIndex As Long, ByVal intTargetCharge As Integer, ByVal lngScanNumberStart As Long, blnUseMaxValueEachScan As Boolean)
    ' Note: The algorithms in this function are the same as those in
    '       Pairs.bas->CalcDltLblPairsScanByScanPopulate
    '
    ' Note that the dblAbundance() array is 1-based, for compatibility reasons with the plot control
    
    Dim lngMemberIndex As Long
    Dim lngScan As Long, lngScanIndex As Long
    Dim intCharge As Integer
    Dim dblAbu As Double
    
    With GelUMC(lngGelIndex).UMCs(UMCIndex)
        For lngMemberIndex = 0 To .ClassCount - 1
            Select Case .ClassMType(lngMemberIndex)
            Case gldtCS
                 lngScan = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).ScanNumber
                 intCharge = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).Charge
                 dblAbu = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).Abundance
            Case gldtIS
                 lngScan = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).ScanNumber
                 intCharge = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).Charge
                 dblAbu = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).Abundance
            End Select
        
            If intTargetCharge <= 0 Or intCharge = intTargetCharge Then
                ' Note: Must add 1 due to 1-based array
                lngScanIndex = lngScan - lngScanNumberStart + 1
                If lngScanIndex < 1 Then
                    ' This shouldn't happen
                    Debug.Assert False
                Else
                    If blnUseMaxValueEachScan Then
                        If dblAbu > dblAbundance(lngScanIndex) Then
                            dblAbundance(lngScanIndex) = dblAbu
                        End If
                    Else
                        dblAbundance(lngScanIndex) = dblAbundance(lngScanIndex) + dblAbu
                    End If
                End If
            End If
        Next lngMemberIndex
    End With

End Sub

Public Function BrowseFeaturesUndeleteSelected(lstFeatures As ListBox, ByRef InfoSortedPointerArray() As Long, ByRef blnFeatureValid() As Boolean, ByRef DeletedStackCount As Long, ByRef DeletedFeatureStack() As Long) As Boolean
    ' Note: This sub is very similar to BrowseFeaturesDeleteSelected
    
    Dim lngItemIndex As Long
    Dim lngDereferencedIndex As Long
    Dim lngIndex As Long, lngTargetIndex As Long
    Dim blnUpdateListBox As Boolean
    
On Error GoTo UnDeleteSelectedFeatureErrorHandler

    ' Find each item selected in lstFeatures
    For lngItemIndex = 0 To lstFeatures.ListCount - 1
        If lstFeatures.Selected(lngItemIndex) Then
            lngDereferencedIndex = InfoSortedPointerArray(lngItemIndex)
            
            If blnFeatureValid(lngDereferencedIndex) Then
                ' Pair already not deleted
            Else
                blnFeatureValid(lngDereferencedIndex) = True
            
                If DeletedStackCount > 0 Then
                    ' Make sure lngDereferencedIndex is not in DeletedFeatureStack
                    lngTargetIndex = 0
                    For lngIndex = 0 To DeletedStackCount - 1
                        If DeletedFeatureStack(lngIndex) <> lngDereferencedIndex Then
                            DeletedFeatureStack(lngTargetIndex) = DeletedFeatureStack(lngIndex)
                            lngTargetIndex = lngTargetIndex + 1
                        End If
                    Next lngIndex
                    If lngTargetIndex < DeletedStackCount Then
                        DeletedStackCount = lngTargetIndex
                        blnUpdateListBox = True
                    End If
                Else
                    DeletedStackCount = 1
                    ReDim DeletedFeatureStack(0)
                    DeletedFeatureStack(0) = lngDereferencedIndex
                    blnUpdateListBox = True
                End If
            End If
        End If
    Next lngItemIndex
    
    BrowseFeaturesUndeleteSelected = blnUpdateListBox
    
    Exit Function

UnDeleteSelectedFeatureErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmPairBrowser->UnDeleteSelectedPair", Err.Description
    BrowseFeaturesUndeleteSelected = False
    
End Function

Public Sub BrowseFeaturesZoom2DPlot(ByRef udtOptions As udtFeatureBrowserOptionsType, lngGelIndex As Long, lngUMCIndex As Long, Optional lngUMCIndex2 As Long = -1)
    ' If lngUMCIndex2 is >=0 then examines it, in addition to lngUMCIndex
    
    Dim lngScanMin As Long, lngScanMax As Long
    Dim dblMassMin As Double, dblMassMax As Double

    Dim blnUsePpm As Boolean
    Dim blnUseNET As Boolean
    
    Dim dblMassHalfWidthUser As Double
    Dim dblMassHalfWidthDa As Double
    Dim dblCentralMass As Double

    Dim lngCentralScan As Long
    
    Dim sngScanWidth As Single
    Dim dblScanHalfWidth As Double
    Dim lngScanHalfWidth As Long

    Dim lngIndex As Long
    Dim lngIonCount As Long
    Dim lngIonPointerArray() As Long

On Error GoTo BrowseFeaturesZoom2DPlotErrorHandler

    dblMassHalfWidthUser = udtOptions.MassRangeZoom

    sngScanWidth = udtOptions.ScanRangeZoom

    If udtOptions.MassRangeUnits = mruPpm Then
        blnUsePpm = True
    Else
        blnUsePpm = False
    End If

    If udtOptions.ScanRangeUnits = sruNet Then
        blnUseNET = True
    Else
        blnUseNET = False
    End If

    With GelUMC(lngGelIndex)
        If lngUMCIndex2 >= 0 Then
            BrowseFeaturesLookupScanAndMassLimits .UMCs(lngUMCIndex), .UMCs(lngUMCIndex2), lngScanMin, lngScanMax, dblMassMin, dblMassMax
            dblCentralMass = (.UMCs(lngUMCIndex).ClassMW + .UMCs(lngUMCIndex2).ClassMW) / 2
        Else
            ' Only consider the first UMC; sending it twice to BrowseFeaturesLookupScanAndMassLimits won't hurt anything
            BrowseFeaturesLookupScanAndMassLimits .UMCs(lngUMCIndex), .UMCs(lngUMCIndex), lngScanMin, lngScanMax, dblMassMin, dblMassMax
            dblCentralMass = .UMCs(lngUMCIndex).ClassMW
        End If
    End With
    
    ' Determine the central scan
    lngCentralScan = Round((lngScanMin + lngScanMax) / 2, 0)

    If blnUsePpm Then
        dblMassHalfWidthUser = PPMToMass(dblMassHalfWidthUser, dblCentralMass)
    End If
    dblMassHalfWidthDa = Abs(dblMassHalfWidthUser / 2)

    If dblMassHalfWidthDa < 0.00001 Then dblMassHalfWidthDa = 0.00001
    
    dblScanHalfWidth = Abs(sngScanWidth / 2)
    If blnUseNET Then
        ' Convert from NET to scan
        lngScanHalfWidth = GANETToScan(lngGelIndex, dblScanHalfWidth)
    Else
        lngScanHalfWidth = Round(dblScanHalfWidth, 0)
    End If
    
    If lngScanHalfWidth < 2 Then lngScanHalfWidth = 2
    
    If udtOptions.FixedDimensionsForAutoZoom Then
        
        dblMassMin = dblCentralMass - dblMassHalfWidthDa
        dblMassMax = dblCentralMass + dblMassHalfWidthDa
        
        lngScanMin = lngCentralScan - lngScanHalfWidth
        lngScanMax = lngCentralScan + lngScanHalfWidth
    Else
        ' The dimensions define the edge width around the feature
        dblMassMin = dblMassMin - dblMassHalfWidthDa
        dblMassMax = dblMassMax + dblMassHalfWidthDa
        
        lngScanMin = lngScanMin - lngScanHalfWidth
        lngScanMax = lngScanMax + lngScanHalfWidth
    End If
    
    ' Zoom the 2D plot
    ZoomGelToDimensions lngGelIndex, CSng(lngScanMin), dblMassMin, CSng(lngScanMax), dblMassMax
    
    If udtOptions.HighlightMembers Then
        ' Highlight the points that are members of this Feature
        
        GelBody(lngGelIndex).GelSel.Clear
        
        ' Retrieve an array of the ion indices of the ions currently "In Scope" and part of the light UMC
        ' Note that GetISScope will ReDim lngIonPointerArray() automatically
        lngIonCount = GetISScopeFilterByUMC(lngGelIndex, lngIonPointerArray(), glScope.glSc_Current, lngUMCIndex)
        For lngIndex = 1 To lngIonCount
            GelBody(lngGelIndex).GelSel.AddToIsoSelection lngIonPointerArray(lngIndex)
        Next lngIndex
        
        If lngUMCIndex2 >= 0 Then
            lngIonCount = GetISScopeFilterByUMC(lngGelIndex, lngIonPointerArray(), glScope.glSc_Current, lngUMCIndex2)
            For lngIndex = 1 To lngIonCount
                GelBody(lngGelIndex).GelSel.AddToIsoSelection lngIonPointerArray(lngIndex)
            Next lngIndex
        End If
        
        GelBody(lngGelIndex).RequestRefreshPlot
    End If
    
    Exit Sub
    
BrowseFeaturesZoom2DPlotErrorHandler:
    Debug.Assert False
    MsgBox "Error auto zooming: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    LogErrors Err.Number, "BrowseFeaturesZoom2DPlot", Err.Description, lngGelIndex

End Sub

Public Sub BrowseFeaturesLookupScanAndMassLimits(ByRef udtUMC As udtUMCType, ByRef udtUMC2 As udtUMCType, ByRef lngScanMin As Long, ByRef lngScanMax As Long, ByRef dblMassMin As Double, ByRef dblMassMax As Double)
    lngScanMin = udtUMC.MinScan
    lngScanMax = udtUMC.MaxScan
    
    If udtUMC2.MinScan < lngScanMin Then
        lngScanMin = udtUMC2.MinScan
    End If
    
    If udtUMC2.MaxScan > lngScanMax Then
        lngScanMax = udtUMC2.MaxScan
    End If

    dblMassMin = udtUMC.MinMW
    dblMassMax = udtUMC.MaxMW
    
    If udtUMC2.MinMW < dblMassMin Then
        dblMassMin = udtUMC2.MinMW
    End If
    
    If udtUMC2.MaxMW > dblMassMax Then
        dblMassMax = udtUMC2.MaxMW
    End If

End Sub



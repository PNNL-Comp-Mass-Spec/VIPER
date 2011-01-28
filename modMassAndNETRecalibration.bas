Attribute VB_Name = "modMassAndNETRecalibration"
Option Explicit

Private Sub AdjustAMTRefMwErrValues(ByRef strAMTRef As String, dblPPMAdjust As Double)
    
    Dim strAMTRefNew As String, strMWErrPPM As String
    Dim lngCharIndex As Long, lngPos1 As Long, lngPos2 As Long
    
    If Len(strAMTRef) > 0 Then
        If IsAMTReferenced(strAMTRef) Then
            ' Correct the ppm error values stored in the Refs, since they are now wrong
            
            strAMTRefNew = strAMTRef
            
            lngCharIndex = 1
            Do While lngCharIndex < Len(strAMTRefNew)
                lngPos1 = InStr(lngCharIndex, strAMTRefNew, MWErrMark)
                If lngPos1 > 0 Then
                    lngPos1 = lngPos1 + Len(MWErrMark)
                    lngPos2 = InStr(lngPos1, strAMTRefNew, MWErrEnd)
                    If lngPos2 > 0 Then
                        strMWErrPPM = Mid(strAMTRefNew, lngPos1, lngPos2 - lngPos1)
                        
                        If IsNumeric(strMWErrPPM) Then
                            strMWErrPPM = Round(CDbl(strMWErrPPM) + dblPPMAdjust, 2)
                            strAMTRefNew = Left(strAMTRefNew, lngPos1 - 1) & strMWErrPPM & Mid(strAMTRefNew, lngPos2)
                        Else
                            Debug.Assert False
                        End If
                    Else
                        Debug.Assert False
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
                lngCharIndex = lngPos2 + 1
            Loop
            
            strAMTRef = strAMTRefNew
         End If
    End If

End Sub

Public Function MassCalibrationApplyBulkAdjustment(ByVal lngGelIndex As Long, _
                                                   ByVal dblIncrementalShift As Double, _
                                                   ByVal eMassType As glMassToleranceConstants, _
                                                   Optional ByVal blnMakeLogEntry As Boolean = True, _
                                                   Optional ByVal sngBinSizeUsedDuringAutoCalibration As Single = 0, _
                                                   Optional ByRef frmCallingForm As VB.Form) As Boolean
    ' Returns True if successful, False if not
    ' If called during Auto calibration, sngBinSizeUsedDuringAutoCalibration is sent so that it may be recorded in the analysis history
    
    Dim lngIndex As Long
    Dim i As Long
    Dim dblMassShiftPPM As Double
    
    Dim blnProceed As Boolean, blnSuccess As Boolean
    
On Error GoTo ApplyMassCalibrationAdjustmentErrorHandler

    blnProceed = MassCalibrationUpdateHistory(lngGelIndex, dblIncrementalShift, eMassType, False, blnMakeLogEntry, sngBinSizeUsedDuringAutoCalibration, False)
    
    If blnProceed Then
        With GelData(lngGelIndex)
            
            Select Case eMassType
            Case gltABS
                For lngIndex = 1 To .CSLines
                    ' Convert the absolute shift to ppm then call MassCalibrationApplyAdjustmentOnePoint
                    dblMassShiftPPM = MassToPPM(dblIncrementalShift, .CSData(lngIndex).AverageMW)
                    MassCalibrationApplyAdjustmentOnePoint .CSData(lngIndex), dblMassShiftPPM, False
                Next lngIndex
                
                For lngIndex = 1 To .IsoLines
                    ' Convert the absolute shift to ppm then call MassCalibrationApplyAdjustmentOnePoint
                    ' Note that we're using the Monoisotopic mass to convert to ppm and then applying the same ppm value to the various isotopic-related masses
                    dblMassShiftPPM = MassToPPM(dblIncrementalShift, .IsoData(lngIndex).MonoisotopicMW)
                    MassCalibrationApplyAdjustmentOnePoint .IsoData(lngIndex), dblMassShiftPPM, True
                Next lngIndex
            Case gltPPM
                For lngIndex = 1 To .CSLines
                    MassCalibrationApplyAdjustmentOnePoint .CSData(lngIndex), dblIncrementalShift, False
                Next lngIndex
                
                For lngIndex = 1 To .IsoLines
                    MassCalibrationApplyAdjustmentOnePoint .IsoData(lngIndex), dblIncrementalShift, True
                Next lngIndex
            Case Else
                ' This shouldn't happen
                Debug.Assert False
            End Select
        End With
    End If
    
    If blnProceed Then
    
        ' Now update the .ClassMW, .MinMW, and .MaxMW values associated with each UMC
        ' This step will actually only be performed if GelUMC().def.LoadedPredefinedLCMSFeatures = True
        MassCalibrationUpdateUMCClassStats lngGelIndex
    
        ' Need to recompute the UMC Statistic arrays and store the updated Class Representative Mass
        ' However, if we loaded predefined LCMSFeatures, then the call to MassCalibrationUpdateUMCClassStats above
        '  will have already updated the class rep mass
        
        Dim blnComputeClassMass As Boolean
        Dim blnComputeClassAbundance As Boolean
        
        If GelUMC(lngGelIndex).def.LoadedPredefinedLCMSFeatures Then
            blnComputeClassMass = False
            blnComputeClassAbundance = False
        Else
            blnComputeClassMass = True
            blnComputeClassAbundance = True
        End If
                
        blnSuccess = UpdateUMCStatArrays(lngGelIndex, blnComputeClassMass, blnComputeClassAbundance, False, frmCallingForm)
        Debug.Assert blnSuccess
        
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "A new mass calibration has been applied (incremental shift of " & dblIncrementalShift & " " & GetSearchToleranceUnitText(eMassType) & ").  You should repeat any database searches done previously."
        End If
    End If
    
    MassCalibrationApplyBulkAdjustment = blnProceed
    Exit Function
    
ApplyMassCalibrationAdjustmentErrorHandler:
    Debug.Print "Error in MassCalibrationApplyBulkAdjustment: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "MassCalibrationApplyBulkAdjustment"
    MassCalibrationApplyBulkAdjustment = False
    
End Function

Public Sub MassCalibrationApplyAdjustmentOnePoint(ByRef udtDataPoint As udtIsotopicDataType, ByVal dblMassShiftPPM As Double, ByVal blnIsoData As Boolean)
    
    If blnIsoData Then
        With udtDataPoint
            .AverageMW = MassCalibrationApplyAdjustmentOnePointWork(.AverageMW, dblMassShiftPPM)
            .MonoisotopicMW = MassCalibrationApplyAdjustmentOnePointWork(.MonoisotopicMW, dblMassShiftPPM)
            .MostAbundantMW = MassCalibrationApplyAdjustmentOnePointWork(.MostAbundantMW, dblMassShiftPPM)
            .MZ = MassCalibrationApplyAdjustmentOnePointWork(.MZ, dblMassShiftPPM)
        
            AdjustAMTRefMwErrValues .MTID, dblMassShiftPPM
                
            If .MassShiftCount < 255 Then .MassShiftCount = .MassShiftCount + 1
            .MassShiftOverallPPM = .MassShiftOverallPPM + dblMassShiftPPM
        End With
    Else
        With udtDataPoint
            .AverageMW = MassCalibrationApplyAdjustmentOnePointWork(.AverageMW, dblMassShiftPPM)
            
            If .MassShiftCount < 255 Then .MassShiftCount = .MassShiftCount + 1
            .MassShiftOverallPPM = .MassShiftOverallPPM + dblMassShiftPPM
        End With
    End If

End Sub

Private Function MassCalibrationApplyAdjustmentOnePointWork(ByVal dblMass As Double, ByVal dblMassShiftPPM As Double) As Double
    Dim dblMassShiftDa As Double
    
    dblMassShiftDa = dblMassShiftPPM / 1000000# * dblMass
    MassCalibrationApplyAdjustmentOnePointWork = dblMass + dblMassShiftDa
End Function

Private Function MassCalibrationRevertAdjustmentOnePoint(ByVal dblMass As Double, ByVal dblMassShiftPPM As Double) As Double
    Dim dblDivisor As Double
 
    dblDivisor = 1 + (dblMassShiftPPM / 1000000#)
    If dblDivisor <> 0 Then
        MassCalibrationRevertAdjustmentOnePoint = Round(dblMass / dblDivisor, 8)
    Else
        MassCalibrationRevertAdjustmentOnePoint = dblMass
    End If

End Function

Public Function MassCalibrationRevertToOriginal(ByVal lngGelIndex As Long, Optional ByVal blnQueryUserToConfirm As Boolean = True, Optional ByVal blnMakeLogEntry As Boolean = True, Optional ByRef frmCallingForm As VB.Form) As Boolean
    ' Returns True if the mass calibration was averted
    ' Returns False if the user cancelled the operation or an error occurred
    
    Dim eResponse As VbMsgBoxResult
    Dim lngIndex As Long
    Dim i As Long
    
    Dim blnDataUpdated As Boolean
    Dim blnSuccess As Boolean
    
On Error GoTo MassCalibrationRevertToOriginalErrorHandler

    If GelSearchDef(lngGelIndex).MassCalibrationInfo.AdjustmentHistoryCount <= 0 Then
        If blnQueryUserToConfirm And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No prior mass calibration adjustments were found in memory.", vbExclamation + vbOKOnly, "Nothing to Do"
        End If
        MassCalibrationRevertToOriginal = False
        Exit Function
    End If
    
    If blnQueryUserToConfirm And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Revert to original masses by removing all mass calibration adjustments?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Revert to Original Masses")
        If eResponse <> vbYes Then
            MassCalibrationRevertToOriginal = False
            Exit Function
        End If
    End If
    
    With GelData(lngGelIndex)
        For lngIndex = 1 To .CSLines
            With .CSData(lngIndex)
                If .MassShiftCount > 0 Then
                    .AverageMW = MassCalibrationRevertAdjustmentOnePoint(.AverageMW, .MassShiftOverallPPM)

                    .MassShiftOverallPPM = 0
                    .MassShiftCount = 0
                    blnDataUpdated = True
                End If
            End With
        Next lngIndex
        
        For lngIndex = 1 To .IsoLines
            With .IsoData(lngIndex)
                If .MassShiftCount > 0 Then
                    .AverageMW = MassCalibrationRevertAdjustmentOnePoint(.AverageMW, .MassShiftOverallPPM)
                    .MonoisotopicMW = MassCalibrationRevertAdjustmentOnePoint(.MonoisotopicMW, .MassShiftOverallPPM)
                    .MostAbundantMW = MassCalibrationRevertAdjustmentOnePoint(.MostAbundantMW, .MassShiftOverallPPM)
                    .MZ = MassCalibrationRevertAdjustmentOnePoint(.MZ, .MassShiftOverallPPM)

                    AdjustAMTRefMwErrValues .MTID, -CDbl(.MassShiftOverallPPM)
                    
                    .MassShiftOverallPPM = 0
                    .MassShiftCount = 0
                    blnDataUpdated = True
                End If
            End With
        Next lngIndex
    End With
    
    If GelUMC(lngGelIndex).def.LoadedPredefinedLCMSFeatures Then
        ' We loaded predefined LCMSFeatures; need to undo the changes made by MassCalibrationUpdateUMCClassStats
        ' Update the .ClassMW, .MinMW, and .MaxMW values for each UMC
        
        For i = 0 To GelUMC(lngGelIndex).UMCCnt - 1
            With GelUMC(lngGelIndex).UMCs(i)
                If .MassShiftCount > 0 Then
                    .ClassMW = MassCalibrationRevertAdjustmentOnePoint(.ClassMW, .MassShiftOverallPPM)
                    .MinMW = MassCalibrationRevertAdjustmentOnePoint(.MinMW, .MassShiftOverallPPM)
                    .MaxMW = MassCalibrationRevertAdjustmentOnePoint(.MaxMW, .MassShiftOverallPPM)
                    .MassShiftOverallPPM = 0
                    .MassShiftCount = 0
                End If
            End With
        Next i
        
    End If
    
    With GelSearchDef(lngGelIndex).MassCalibrationInfo
        If blnMakeLogEntry And blnDataUpdated Then
            AddToAnalysisHistory lngGelIndex, "Previous mass calibration adjustments have been reversed; Prior Adjustment Count = " & Trim(.AdjustmentHistoryCount)
        End If
        
        .AdjustmentHistoryCount = 0
        ReDim .AdjustmentHistory(0)
        .OverallMassAdjustment = 0
    End With
    
    If blnDataUpdated Then
        ' Need to recompute the UMC Statistic arrays and store the updated Class Representative Mass
        ' However, if we loaded predefined LCMSFeatures, then code earlier in this function
        '  has already updated the class rep mass
        
        Dim blnComputeClassMass As Boolean
        Dim blnComputeClassAbundance As Boolean
                
        If GelUMC(lngGelIndex).def.LoadedPredefinedLCMSFeatures Then
            blnComputeClassMass = False
            blnComputeClassAbundance = False
        Else
            blnComputeClassMass = True
            blnComputeClassAbundance = True
        End If
        
        blnSuccess = UpdateUMCStatArrays(lngGelIndex, blnComputeClassMass, blnComputeClassAbundance, False, frmCallingForm)
        Debug.Assert blnSuccess
    End If
    
    MassCalibrationRevertToOriginal = blnDataUpdated
    
    Exit Function
    
MassCalibrationRevertToOriginalErrorHandler:
    Debug.Print "Error in MassCalibrationRevertToOriginal: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "MassCalibrationRevertToOriginal"
    MassCalibrationRevertToOriginal = False
    
End Function

Public Function MassCalibrationUpdateHistory(ByVal lngGelIndex As Long, ByVal dblIncrementalShift As Double, ByVal eMassType As glMassToleranceConstants, ByVal blnResetHistoryIfConflictingMassType As Boolean, ByVal blnMakeLogEntry As Boolean, ByVal sngBinSizeUsedDuringAutoCalibration As Single, ByVal blnUsingMSAlign As Boolean) As Boolean
    
    Dim blnProceed As Boolean
    Dim strMessage As String
    
    With GelSearchDef(lngGelIndex).MassCalibrationInfo
        If .AdjustmentHistoryCount = 0 Then
            ReDim .AdjustmentHistory(0)
            .MassUnits = eMassType
            blnProceed = True
        Else
            If eMassType <> .MassUnits Then
                If blnResetHistoryIfConflictingMassType Then
                    ReDim .AdjustmentHistory(0)
                    .MassUnits = eMassType
                    blnProceed = True
                Else
                    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                        MsgBox "Unable to apply the new mass calibration adjustment since previous adjustments were in " & GetSearchToleranceUnitText(CInt(.MassUnits)) & " while the new adjustment is defined in " & GetSearchToleranceUnitText(eMassType), vbExclamation + vbOKOnly, "Mismatched Units"
                        blnProceed = False
                    End If
                End If
            Else
                ReDim Preserve .AdjustmentHistory(.AdjustmentHistoryCount + 1)
                blnProceed = True
            End If
        End If
        
        If blnProceed Then
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                If .AdjustmentHistoryCount <> 0 Or .OverallMassAdjustment <> 0 Then
                    If GelUMCNETAdjDef(lngGelIndex).UseRobustNETAdjustment And GelUMCNETAdjDef(lngGelIndex).RobustNETAdjustmentMode >= UMCRobustNETModeConstants.UMCRobustNETWarpTime Then
                        ' This is acceptable behavior during AutoAnalysis
                    Else
                        ' This is potentially a bug that shouldn't happen: This function should never be called twice during AutoAnalysis
                        Debug.Assert False
                    End If
                End If
            End If
            
            .AdjustmentHistoryCount = .AdjustmentHistoryCount + 1
            .AdjustmentHistory(.AdjustmentHistoryCount - 1) = dblIncrementalShift
        
            .OverallMassAdjustment = .OverallMassAdjustment + dblIncrementalShift
            
            Debug.Assert .MassUnits = eMassType
            
            
            If blnMakeLogEntry Then
                With GelSearchDef(lngGelIndex).MassCalibrationInfo
                    If blnUsingMSAlign Then
                        strMessage = "Mass calibration adjustment applied; Used warped masses from MS Warp; Avg "
                    Else
                        strMessage = "Mass calibration adjustment applied; "
                    End If
                    
                    strMessage = strMessage & "Mass Shift = " & Format(dblIncrementalShift, "0.0000") & " " & GetSearchToleranceUnitText(eMassType)
                    If sngBinSizeUsedDuringAutoCalibration > 0 Then
                        strMessage = strMessage & "; Mass error plot bin size = " & Trim(sngBinSizeUsedDuringAutoCalibration) & " " & GetSearchToleranceUnitText(CInt(.MassUnits))
                    End If
                    strMessage = strMessage & "; New overall adjustment = " & Format(.OverallMassAdjustment, "0.0000") & " " & GetSearchToleranceUnitText(CInt(.MassUnits))
                    strMessage = strMessage & "; Total adjustments applied = " & .AdjustmentHistoryCount
                End With
                
                AddToAnalysisHistory lngGelIndex, strMessage
            End If
        
        
        End If
    End With

    MassCalibrationUpdateHistory = blnProceed
    
End Function

Public Function MassCalibrationUpdateUMCClassStats(ByVal lngGelIndex As Long) As Boolean

    ' Returns True if successful, False if not
    
    Dim i As Long
    Dim lngClassMIndexPointer As Long
    Dim lngClassRepInd As Long
    Dim lngClassRepType As Long
    
    Dim dblMassShiftPPM As Double
    Dim bytMassShiftCount As Byte
    
On Error GoTo MassCalibrationUpdateUMCClassStatsErrorHandler

    If GelUMC(lngGelIndex).def.LoadedPredefinedLCMSFeatures Then
        ' Need to update the .MinMW and .MaxMW values for each UMC
        ' We will not update .ClassMW since UpdateUMCStatArrays will call CalculateClasses, and CalculateClasses will re-compute .ClassMW using the ClassRep
        
        With GelUMC(lngGelIndex)
            
            For i = 0 To .UMCCnt - 1
                If .def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                    With .UMCs(i)
                        ' Obtain the class rep info from the best charge state group
                        lngClassMIndexPointer = .ChargeStateBasedStats(.ChargeStateStatsRepInd).GroupRepIndex
                        If lngClassMIndexPointer < 0 Then
                            lngClassRepInd = .ClassRepInd
                            lngClassRepType = .ClassRepType
                        Else
                            lngClassRepInd = .ClassMInd(lngClassMIndexPointer)
                            lngClassRepType = .ClassMType(lngClassMIndexPointer)
                        End If
                    End With
                Else
                    lngClassRepInd = .UMCs(i).ClassRepInd
                    lngClassRepType = .UMCs(i).ClassRepType
                End If
                
                If lngClassRepInd < 0 Then
                    dblMassShiftPPM = 0
                Else
                    ' Lookup the overall mass shift currently applied to this point
                    Select Case lngClassRepType
                        Case glCSType
                            dblMassShiftPPM = GelData(lngGelIndex).CSData(lngClassRepInd).MassShiftOverallPPM
                            bytMassShiftCount = GelData(lngGelIndex).CSData(lngClassRepInd).MassShiftCount
                        Case glIsoType
                            dblMassShiftPPM = GelData(lngGelIndex).IsoData(lngClassRepInd).MassShiftOverallPPM
                            bytMassShiftCount = GelData(lngGelIndex).IsoData(lngClassRepInd).MassShiftCount
                        Case Else
                            dblMassShiftPPM = 0
                    End Select
                End If
                
                With .UMCs(i)
                
                    ' Revert any adjustments already applied to .ClassMW, .MinMW, .MaxMW
                    If .MassShiftCount > 0 Then
                        .ClassMW = MassCalibrationRevertAdjustmentOnePoint(.ClassMW, .MassShiftOverallPPM)
                        .MinMW = MassCalibrationRevertAdjustmentOnePoint(.MinMW, .MassShiftOverallPPM)
                        .MaxMW = MassCalibrationRevertAdjustmentOnePoint(.MaxMW, .MassShiftOverallPPM)
                        .MassShiftOverallPPM = 0
                        .MassShiftCount = 0
                    End If
                    
                    ' Apply the new mass shift
                    If dblMassShiftPPM <> 0 Then
                        .ClassMW = MassCalibrationApplyAdjustmentOnePointWork(.ClassMW, dblMassShiftPPM)
                        .MinMW = MassCalibrationApplyAdjustmentOnePointWork(.MinMW, dblMassShiftPPM)
                        .MaxMW = MassCalibrationApplyAdjustmentOnePointWork(.MaxMW, dblMassShiftPPM)
                    
                        .MassShiftOverallPPM = dblMassShiftPPM
                        .MassShiftCount = bytMassShiftCount
                    End If
                End With
               
            Next i
        End With
    End If
    
    MassCalibrationUpdateUMCClassStats = True
    Exit Function
    
MassCalibrationUpdateUMCClassStatsErrorHandler:
    Debug.Print "Error in MassCalibrationUpdateUMCClassStats: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "MassCalibrationUpdateUMCClassStats"
    MassCalibrationUpdateUMCClassStats = False
    
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDltLblPairsUMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtERStatsType
    ERAvg As Double             ' Single ER or average ER
    ERStDev As Double           ' ER Standard Deviation if averaging
    ERCount As Long             ' Number of values used to average, or 1 if no averaging
    EROutlierCountRemoved As Long       ' Number of ER values removed, since they were found to be outliers
    
    TotalAbundance As Double    ' The sum of the abundances in both members of the pair, possibly for a given charge state
    TotalMemberCount As Long    ' The sum of the member counts in both members of a pair, possibly for a given charge state
End Type

Private Type udtPairsERChargeStateStatsType
    Charge As Integer
    ERStats As udtERStatsType
End Type
    
Private Type udtScanByScanDetailsType
    Abundance As Double
    MemberCount As Long
    MemberTypeMostAbu As Long               ' Type of member: gldtCS or gldtIS
    MemberIndexMostAbu As Long              ' Pointer to entry in GelData().IsoData() or GelData().CSData()
End Type


Private Type udtScanByScanStatsType
    ScanNumStart As Long
    ScanNumEnd As Long
    
    ScanNumCount As Long
    ScanDetailsLt() As udtScanByScanDetailsType
    ScanDetailsHv() As udtScanByScanDetailsType
    ScanER() As Double
    ScanAbu2PlusCorrected() As Single
    ScanAbu4PlusCorrected() As Single
    
    ERStats As udtERStatsType
End Type

Private Type udtCorrectedAbundanceStatsByScanType
    ScanNumber As Long
    LightAbundance As Double
    HeavyAbundance As Double
    ER As Double
    ScanAbu2PlusCorrected As Single
    ScanAbu4PlusCorrected As Single
End Type

Private Type udtCorrectedAbundanceStatsType
    AlignedScanCount As Long
    AlignedScanStats() As udtCorrectedAbundanceStatsByScanType
End Type

' mCorrectedAbundanceStats should be parallel to GelP_D_L(mGelIndex).Pairs
Private mCorrectedAbundanceStatsCount As Long
Private mCorrectedAbundanceStats() As udtCorrectedAbundanceStatsType

Private mGelIndex As Long

Private Sub Class_Initialize()
    mGelIndex = 0
    
    InitializeLocalVariables 10
End Sub

Public Sub CalcDltLblPairsER_UMC(ByVal lngGelIndex As Long, ByRef strMessage As String)
    '----------------------------------------------------------
    'creates space for and calculates expression ratio for UMC-based pairs
    '----------------------------------------------------------
    Dim lngPairIndex As Long
    Dim lngPairCountWithOutliers As Long
    Dim dblOutlierFractionSum As Double
    
    mGelIndex = lngGelIndex
    
    Call InitDltLblPairsER(mGelIndex, GelP_D_L(mGelIndex).SearchDef.ERCalcType)
    
    InitializeLocalVariables GelP_D_L(mGelIndex).PCnt
    
    lngPairCountWithOutliers = 0
    dblOutlierFractionSum = 0
    
    For lngPairIndex = 0 To GelP_D_L(mGelIndex).PCnt - 1
        CalcDltLblPairsER_UMCSinglePair lngPairIndex, lngPairCountWithOutliers, dblOutlierFractionSum
    Next lngPairIndex
  
    strMessage = "Calculated pair expression ratios; Pair count = " & GelP_D_L(mGelIndex).PCnt
    If lngPairCountWithOutliers > 0 Then
        strMessage = strMessage & "; Removed outlier ER values from " & lngPairCountWithOutliers & " pairs"
        strMessage = strMessage & "; Average % of ER values of pair that were removed as outliers = " & Round(dblOutlierFractionSum / CDbl(lngPairCountWithOutliers) * 100, 2) & "%"
    End If
    
    AddToAnalysisHistory mGelIndex, strMessage
    
End Sub

Private Sub CalcDltLblPairsER_UMCSinglePair(ByVal lngPairIndex As Long, ByRef lngPairCountWithOutliers As Long, ByRef dblOutlierFractionSum As Double)
    
    Dim LtAbu As Double, HvAbu As Double
    
    Dim intChargeStateIndexA As Integer, intChargeStateIndexB As Integer
    Dim intIndex As Integer
    
    Dim intChargeStateStatsCount As Integer
    Dim udtPairsERChargeStateStats() As udtPairsERChargeStateStatsType
    
    Dim udtUMCA As udtUMCType
    Dim udtUMCB As udtUMCType
    
    Dim udtERStats As udtERStatsType
    
    ' Note: When computing a weighted average, use ER = dblERProductSum / dblWeightingSum
    Dim dblERProductSum As Double       ' Sum of ER times value used for weighting
    Dim dblStDevProductSum As Double    ' Sum of ER StDev times value used for weighting
    Dim dblMemberCountProductSum As Double  ' Sum of ER member count value times value used for weighting
    Dim dblOutlierCountRemovedProductSum As Double        ' Sum of ER outlier count removed times value used for weighting
    Dim dblWeightingSum As Double       ' Sum of values used for weighting
    
    Dim eWeightingMode As aewAverageERsWeightingModeConstants

    Dim blnValidER As Boolean
    Dim blnAveragedAcrossChargeStates As Boolean
    Dim blnIReportData As Boolean
    
    eWeightingMode = GelP_D_L(mGelIndex).SearchDef.AverageERsWeightingMode
    If (GelData(mGelIndex).DataStatusBits And GEL_DATA_STATUS_BIT_IREPORT) = GEL_DATA_STATUS_BIT_IREPORT Then
        ' Enable I-Report based ER computation
        ' This will only be used if .ComputeERScanByScan = True
        blnIReportData = glbPreferencesExpanded.PairSearchOptions.SearchDef.IReportEROptions.Enabled
    Else
        blnIReportData = False
    End If
    
    If GelP_D_L(mGelIndex).SearchDef.RequireMatchingChargeStatesForPairMembers And _
       GelP_D_L(mGelIndex).SearchDef.UseIdenticalChargesForER And _
       GelP_D_L(mGelIndex).SearchDef.AverageERsAllChargeStates Then
        ' Need to compute ER value for each charge state matching between the LC-MS Features, then average them together with a weighted average
        
        udtUMCA = GelUMC(mGelIndex).UMCs(GelP_D_L(mGelIndex).Pairs(lngPairIndex).P1)
        udtUMCB = GelUMC(mGelIndex).UMCs(GelP_D_L(mGelIndex).Pairs(lngPairIndex).P2)
        
        ' Determine the maximum possible number of charge states that could match up
        intChargeStateStatsCount = 1
        If udtUMCA.ChargeStateCount > intChargeStateStatsCount Then
            intChargeStateStatsCount = udtUMCA.ChargeStateCount
        End If
        If udtUMCB.ChargeStateCount > intChargeStateStatsCount Then
            intChargeStateStatsCount = udtUMCB.ChargeStateCount
        End If
                    
        ReDim udtPairsERChargeStateStats(intChargeStateStatsCount - 1)
        intChargeStateStatsCount = 0
        
        For intChargeStateIndexA = 0 To udtUMCA.ChargeStateCount - 1
            For intChargeStateIndexB = 0 To udtUMCB.ChargeStateCount - 1
                If udtUMCA.ChargeStateBasedStats(intChargeStateIndexA).Charge = _
                    udtUMCB.ChargeStateBasedStats(intChargeStateIndexB).Charge Then
                    ' Matching charge states found
                   
                    With udtPairsERChargeStateStats(intChargeStateStatsCount)
                        .Charge = udtUMCA.ChargeStateBasedStats(intChargeStateIndexA).Charge
                        
                        ' Call CalcDltLblPairsERWork to compute the ER for this charge state pair
                        ' Note that .ERStats is returned ByRef
                        ' Also, note that we pass ectER_RAT for the ERCalcType since we need normal ER ratios for the weighted averaging
                        blnValidER = CalcDltLblPairsERWork(mGelIndex, lngPairIndex, .Charge, ectER_RAT, .ERStats, blnIReportData)
                   End With
                   
                   If blnValidER Then intChargeStateStatsCount = intChargeStateStatsCount + 1
                   
                End If
            Next intChargeStateIndexB
        Next intChargeStateIndexA

        If intChargeStateStatsCount = 0 Then
            ' No charge states were found in common between the two LC-MS Features in this pair
            ' This normally shouldn't be the case
            udtERStats.ERAvg = ER_CALC_ERR
            blnAveragedAcrossChargeStates = False
        Else
            blnAveragedAcrossChargeStates = True
            If intChargeStateStatsCount = 1 Then
                ' Only one charge state; nothing to average
                udtERStats = udtPairsERChargeStateStats(0).ERStats
            
            Else
                ' Average the ER values for all of the charge states, optionally weighting by Abundance or member count
                
                dblERProductSum = 0
                dblStDevProductSum = 0
                dblWeightingSum = 0
                dblMemberCountProductSum = 0
                dblOutlierCountRemovedProductSum = 0
                
                For intIndex = 0 To intChargeStateStatsCount - 1
                    If udtPairsERChargeStateStats(intIndex).ERStats.ERAvg = ER_CALC_ERR Then
                        ' This code shouldn't be reached
                        Debug.Assert False
                    Else
                        If eWeightingMode = aewAbundance Then
                            dblERProductSum = dblERProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERAvg * udtPairsERChargeStateStats(intIndex).ERStats.TotalAbundance
                            dblStDevProductSum = dblStDevProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERStDev * udtPairsERChargeStateStats(intIndex).ERStats.TotalAbundance
                            dblMemberCountProductSum = dblMemberCountProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERCount * udtPairsERChargeStateStats(intIndex).ERStats.TotalAbundance
                            dblOutlierCountRemovedProductSum = dblOutlierCountRemovedProductSum + udtPairsERChargeStateStats(intIndex).ERStats.EROutlierCountRemoved * udtPairsERChargeStateStats(intIndex).ERStats.TotalAbundance
                            
                            dblWeightingSum = dblWeightingSum + udtPairsERChargeStateStats(intIndex).ERStats.TotalAbundance
                        ElseIf eWeightingMode = aewMemberCounts Then
                            dblERProductSum = dblERProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERAvg * udtPairsERChargeStateStats(intIndex).ERStats.TotalMemberCount
                            dblStDevProductSum = dblStDevProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERStDev * udtPairsERChargeStateStats(intIndex).ERStats.TotalMemberCount
                            dblMemberCountProductSum = dblMemberCountProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERCount * udtPairsERChargeStateStats(intIndex).ERStats.TotalMemberCount
                            dblOutlierCountRemovedProductSum = dblOutlierCountRemovedProductSum + udtPairsERChargeStateStats(intIndex).ERStats.EROutlierCountRemoved * udtPairsERChargeStateStats(intIndex).ERStats.TotalMemberCount
                            
                            dblWeightingSum = dblWeightingSum + udtPairsERChargeStateStats(intIndex).ERStats.TotalMemberCount
                        Else
                            ' aewNoWeighting
                            dblERProductSum = dblERProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERAvg
                            dblStDevProductSum = dblStDevProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERStDev
                            dblMemberCountProductSum = dblMemberCountProductSum + udtPairsERChargeStateStats(intIndex).ERStats.ERCount
                            dblOutlierCountRemovedProductSum = dblOutlierCountRemovedProductSum + udtPairsERChargeStateStats(intIndex).ERStats.EROutlierCountRemoved
                            
                            dblWeightingSum = dblWeightingSum + 1#
                        End If
                    End If
                Next intIndex
                
                If dblWeightingSum > 0 Then
                    udtERStats.ERAvg = dblERProductSum / dblWeightingSum
                    udtERStats.ERStDev = dblStDevProductSum / dblWeightingSum
                    udtERStats.ERCount = Round(dblMemberCountProductSum / dblWeightingSum, 0)
                    If dblOutlierCountRemovedProductSum > 0 Then
                        udtERStats.EROutlierCountRemoved = Round(dblOutlierCountRemovedProductSum / dblWeightingSum, 0)
                        If udtERStats.EROutlierCountRemoved < 1 Then
                            udtERStats.EROutlierCountRemoved = 1
                        End If
                    Else
                        udtERStats.EROutlierCountRemoved = 1
                    End If
                Else
                    ' The Weighting Sum is 0
                    ' This shouldn't happen
                    Debug.Assert False
                    udtERStats.ERAvg = ER_CALC_ERR
                    udtERStats.ERStDev = 0
                    udtERStats.ERCount = 0
                    udtERStats.EROutlierCountRemoved = 0
                End If
                
                If udtERStats.ERAvg <> ER_CALC_ERR Then
                    ' We specified an ERCalcType of Ratio (ectER_RAT=0) when calling CalcDltLblPairsERWork
                    ' Convert to the correct ratio type, if necessary
                    Select Case GelP_D_L(mGelIndex).SearchDef.ERCalcType
                    Case ectER_RAT
                        ' Nothing to convert
                    Case ectER_LOG
                        udtERStats.ERAvg = LogERViaRatER(udtERStats.ERAvg)
                        udtERStats.ERStDev = LogERViaRatER(udtERStats.ERStDev)
                    Case ectER_ALT
                        udtERStats.ERAvg = AltERViaRatER(udtERStats.ERAvg)
                        udtERStats.ERStDev = AltERViaRatER(udtERStats.ERStDev)
                    End Select
                End If
            End If
        End If
        
    Else
        ' Not averaging ER's across charge states
        ' Just compute one ER value for this pair
        blnValidER = CalcDltLblPairsERWork(mGelIndex, lngPairIndex, 0, GelP_D_L(mGelIndex).SearchDef.ERCalcType, udtERStats, blnIReportData)
        blnAveragedAcrossChargeStates = False
        intChargeStateStatsCount = 1
        
        If Not blnValidER Then
            udtERStats.ERAvg = ER_CALC_ERR
        End If
    End If

    ' Store udtERStats.ERAvg
    With GelP_D_L(mGelIndex).Pairs(lngPairIndex)
        .ER = udtERStats.ERAvg
        .ERStDev = udtERStats.ERStDev
        .ERChargeStateBasisCount = intChargeStateStatsCount
        
        If intChargeStateStatsCount > 0 And blnAveragedAcrossChargeStates Then
            ReDim .ERChargesUsed(intChargeStateStatsCount - 1)
            For intIndex = 0 To intChargeStateStatsCount - 1
                .ERChargesUsed(intIndex) = udtPairsERChargeStateStats(intIndex).Charge
            Next intIndex
        Else
            ReDim .ERChargesUsed(0)
            .ERChargesUsed(0) = 0
        End If
        
        If udtERStats.EROutlierCountRemoved > 0 Then
            lngPairCountWithOutliers = lngPairCountWithOutliers + 1
            If udtERStats.ERCount + udtERStats.EROutlierCountRemoved > 0 Then
                dblOutlierFractionSum = dblOutlierFractionSum + udtERStats.EROutlierCountRemoved / (udtERStats.ERCount + udtERStats.EROutlierCountRemoved)
            End If
        End If
        
        .ERMemberBasisCount = udtERStats.ERCount
        
    End With

End Sub

Private Function CalcDltLblPairsERWork(ByVal lngGelIndex As Long, ByVal lngPairIndex As Long, ByVal intChargeStateToMatch As Integer, eERCalcType As Integer, ByRef udtERStats As udtERStatsType, blnIReportData As Boolean) As Boolean
    ' Computes and returns an ER value and associated stats in udtERStats
    ' Function returns True if success, false if an error
    ' Sets udtERStats.ERAvg = ER_CALC_ERR if an error occurs
    
    ' If intChargeStateToMatch is > 0 then only uses the specified charge state values in the two LC-MS Features
    ' If intChargeStateToMatch is > 0 and the given charge state is only present in one of the LC-MS Features, then returns an ER value of ER_CALC_ERR and sets udtERStats.TotalAbundance to 0 and udtERStats.TotalMemberCount to 0

    Dim LtAbu As Double, HvAbu As Double
    
    Dim intLtUMCCharge As Integer, intHvUMCCharge As Integer
    Dim intIndex As Integer, intIndexCompare As Integer
    Dim intLtUMCChargeStateIndex As Integer
    Dim intHVUMCChargeStateIndex As Integer
    
    Dim LtUMCIndex As Long
    Dim HvUMCIndex As Long
    Dim lngScanIndex As Long
    
    Dim blnMatchFound As Boolean
    Dim blnValidER As Boolean
    Dim blnComputedERScanByScan As Boolean
    
    Dim lngChargeIndexToTest() As Long              ' 0-based array of charge indices in the light UMC to test; sorted parallel with dblChargeIndexIntensities
    Dim dblChargeIndexIntensities() As Double       ' 0-based array of the intensities for the charge indices in the light UMC
    Dim intChargeToFind As Integer
    
    ' This udt is used when computing stats scan-by-scan
    Dim udtScanByScanStats As udtScanByScanStatsType
    Dim intTargetChargeLt As Integer, intTargetChargeHv As Integer
    
    Dim objSort As QSDouble
    
    blnComputedERScanByScan = False
    LtUMCIndex = GelP_D_L(lngGelIndex).Pairs(lngPairIndex).P1
    HvUMCIndex = GelP_D_L(lngGelIndex).Pairs(lngPairIndex).P2
    
    blnValidER = False
    With udtERStats
        .ERAvg = 0
        .ERCount = 0
        .EROutlierCountRemoved = 0
        .ERStDev = 0
        .TotalAbundance = 0
        .TotalMemberCount = 0
    End With
    
    If GelP_D_L(lngGelIndex).SearchDef.UseIdenticalChargesForER Then
        ' Use identical charges for ER
        If intChargeStateToMatch > 0 Then
            ' We're required to use a given charge state
            ' See if both the light and heavy UMC have this charge state
            
            blnMatchFound = False
            intLtUMCChargeStateIndex = -1
            With GelUMC(lngGelIndex).UMCs(LtUMCIndex)
                For intIndex = 0 To .ChargeStateCount - 1
                    If .ChargeStateBasedStats(intIndex).Charge = intChargeStateToMatch Then
                        LtAbu = .ChargeStateBasedStats(intIndex).Abundance
                        intLtUMCChargeStateIndex = intIndex
                        blnMatchFound = True
                        Exit For
                    End If
                Next intIndex
            End With
            
            If intLtUMCChargeStateIndex >= 0 Then
                ' Light member has the charge; what about the heavy member?
                blnMatchFound = False
                With GelUMC(lngGelIndex).UMCs(HvUMCIndex)
                    For intIndex = 0 To .ChargeStateCount - 1
                        If .ChargeStateBasedStats(intIndex).Charge = intChargeStateToMatch Then
                            HvAbu = .ChargeStateBasedStats(intIndex).Abundance
                            intHVUMCChargeStateIndex = intIndex
                            blnMatchFound = True
                            Exit For
                        End If
                    Next intIndex
                End With
            End If
            
        Else
        
            ' Determine appropriate charge state index in each UMC to use for ClassAbundance
            ' Do this by first looking for matching charge states, then, if no matching
            '  charge states can be found using the most abundant charge state in each UMC
            '  (unless intChargeStateToMatch > 0)
            
            blnMatchFound = False
            With GelUMC(lngGelIndex).UMCs(LtUMCIndex)
                LtAbu = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Abundance
                intLtUMCCharge = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge
                intLtUMCChargeStateIndex = .ChargeStateStatsRepInd
            End With
            
            With GelUMC(lngGelIndex).UMCs(HvUMCIndex)
                HvAbu = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Abundance
                intHvUMCCharge = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge
                intHVUMCChargeStateIndex = .ChargeStateStatsRepInd
                If intLtUMCCharge = intHvUMCCharge Then
                    ' Both LC-MS Features have the same representative charge state
                    blnMatchFound = True
                Else
                    ' the LC-MS Features do not have the same representative charge state
                    ' Compare the representative abundances to see which is larger
                    If LtAbu >= HvAbu Then
                        ' Look for intLtUMCCharge in the heavy UMC
                        For intIndex = 0 To .ChargeStateCount - 1
                            If .ChargeStateBasedStats(intIndex).Charge = intLtUMCCharge Then
                                HvAbu = .ChargeStateBasedStats(intIndex).Abundance
                                intHVUMCChargeStateIndex = intIndex
                                blnMatchFound = True
                                Exit For
                            End If
                        Next intIndex
                    End If
                End If
            End With
    
            If Not blnMatchFound Then
                With GelUMC(lngGelIndex).UMCs(LtUMCIndex)
                    ' Look for intHvUMCCharge in the light UMC
                    For intIndex = 0 To .ChargeStateCount - 1
                        If .ChargeStateBasedStats(intIndex).Charge = intHvUMCCharge Then
                            LtAbu = .ChargeStateBasedStats(intIndex).Abundance
                            intLtUMCChargeStateIndex = intIndex
                            blnMatchFound = True
                            Exit For
                        End If
                    Next intIndex
                End With
            End If
            
            If Not blnMatchFound And LtAbu < HvAbu Then
                ' We haven't yet looked for intLtUMCCharge in the heavy UMC; do this now
                With GelUMC(lngGelIndex).UMCs(HvUMCIndex)
                    For intIndex = 0 To .ChargeStateCount - 1
                        If .ChargeStateBasedStats(intIndex).Charge = intLtUMCCharge Then
                            HvAbu = .ChargeStateBasedStats(intIndex).Abundance
                            intHVUMCChargeStateIndex = intIndex
                            blnMatchFound = True
                            Exit For
                        End If
                    Next intIndex
                End With
            End If
            
            If Not blnMatchFound Then
                ' The most abundant charge state of either UMC is not present in the other UMC
                ' Determine the 2nd most abundant charge in the light UMC and look for that charge in the heavy UMC
                ' If not found, then try the 3rd most abundant charge, etc.
                ' However, we're still requiring that identical charge states be used
                With GelUMC(lngGelIndex).UMCs(LtUMCIndex)
                    If .ChargeStateCount > 1 Then
                        ReDim lngChargeIndexToTest(.ChargeStateCount - 1)
                        ReDim dblChargeIndexIntensities(.ChargeStateCount - 1)
                        
                        For intIndex = 0 To .ChargeStateCount - 1
                            lngChargeIndexToTest(intIndex) = intIndex
                            dblChargeIndexIntensities(intIndex) = .ChargeStateBasedStats(intIndex).Abundance
                        Next intIndex
                        
                        ' Sort the arrays parallel, sorting descending on abundance
                        Set objSort = New QSDouble
                        If objSort.QSDesc(dblChargeIndexIntensities, lngChargeIndexToTest) Then
                            For intIndex = 0 To .ChargeStateCount - 1
                                intChargeToFind = .ChargeStateBasedStats(lngChargeIndexToTest(intIndex)).Charge
                                For intIndexCompare = 0 To GelUMC(lngGelIndex).UMCs(HvUMCIndex).ChargeStateCount - 1
                                    If GelUMC(lngGelIndex).UMCs(HvUMCIndex).ChargeStateBasedStats(intIndexCompare).Charge = intChargeToFind Then
                                        LtAbu = .ChargeStateBasedStats(lngChargeIndexToTest(intIndex)).Abundance
                                        intLtUMCChargeStateIndex = intIndex
                                        
                                        HvAbu = GelUMC(lngGelIndex).UMCs(HvUMCIndex).ChargeStateBasedStats(intIndexCompare).Abundance
                                        intHVUMCChargeStateIndex = intIndexCompare
                                        blnMatchFound = True
                                        Exit For
                                    End If
                                Next intIndexCompare
                                If blnMatchFound Then Exit For
                            Next intIndex
                        End If
                        Set objSort = Nothing
                    End If
    
                End With
                
            End If
        End If
        
        ' If, at this point, blnMatchFound = False, then no matching charge states could
        '   be found and we're forced to use the most abundant charge state for each UMC
        ' However, return ER_CALC_ERR if required to match a charge state and no match is found
        
        If intChargeStateToMatch > 0 And Not blnMatchFound Then
            ' We required a specific charge state and no match was found
            ' Return ER_CALC_ERR
            udtERStats.ERAvg = ER_CALC_ERR
            blnValidER = False
        Else
            ' Either a match was found, or no match was found but no specific charge state was required
            ' In either case, we now know LtAbu, HvAbu, intLtUMCChargeStateIndex, and intHVUMCChargeStateIndex for LtUMCIndex and HvUMCIndex
            
            If GelP_D_L(lngGelIndex).SearchDef.ComputeERScanByScan Then
                ' Compute the ER scan by scan
                
                With GelUMC(lngGelIndex).UMCs(LtUMCIndex)
                    intTargetChargeLt = .ChargeStateBasedStats(intLtUMCChargeStateIndex).Charge
                End With
                
                With GelUMC(lngGelIndex).UMCs(HvUMCIndex)
                    intTargetChargeHv = .ChargeStateBasedStats(intHVUMCChargeStateIndex).Charge
                End With
            
                If intChargeStateToMatch > 0 Then Debug.Assert intChargeStateToMatch = intTargetChargeLt And intChargeStateToMatch = intTargetChargeHv
                
                blnValidER = CalcDltLblPairsERScanByScan(udtScanByScanStats, lngGelIndex, LtUMCIndex, HvUMCIndex, intTargetChargeLt, intTargetChargeHv, blnIReportData)
                blnComputedERScanByScan = True
                
                If blnValidER Then
                    udtERStats = udtScanByScanStats.ERStats
                End If
                
            Else
                ' Do not compute the ER scan by scan
                Select Case eERCalcType
                Case ectER_RAT
                    udtERStats.ERAvg = RatER(LtAbu, HvAbu)
                Case ectER_LOG
                    udtERStats.ERAvg = LogER(LtAbu, HvAbu)
                Case ectER_ALT
                    udtERStats.ERAvg = AltER(LtAbu, HvAbu)
                End Select
                udtERStats.ERCount = 1
                
                With GelUMC(lngGelIndex).UMCs(LtUMCIndex).ChargeStateBasedStats(intLtUMCChargeStateIndex)
                    udtERStats.TotalAbundance = .Abundance
                    udtERStats.TotalMemberCount = .Count
                End With
                
                With GelUMC(lngGelIndex).UMCs(HvUMCIndex).ChargeStateBasedStats(intHVUMCChargeStateIndex)
                    udtERStats.TotalAbundance = udtERStats.TotalAbundance + .Abundance
                    udtERStats.TotalMemberCount = udtERStats.TotalMemberCount + .Count
                End With
            
                blnValidER = True
            End If
        End If
    
    Else
        ' Do not require identical charges for ER
        If GelP_D_L(lngGelIndex).SearchDef.ComputeERScanByScan Then
            ' Compute the ER scan by scan
            blnValidER = CalcDltLblPairsERScanByScan(udtScanByScanStats, lngGelIndex, LtUMCIndex, HvUMCIndex, 0, 0, blnIReportData)
            blnComputedERScanByScan = True
            
            If blnValidER Then
                udtERStats = udtScanByScanStats.ERStats
            End If
        Else
            ' Simply use the class abundance of each UMC
            With GelUMC(lngGelIndex).UMCs(LtUMCIndex)
                LtAbu = .ClassAbundance
                udtERStats.TotalAbundance = .ClassAbundance
                udtERStats.TotalMemberCount = .ClassCount
            End With
            
            With GelUMC(lngGelIndex).UMCs(HvUMCIndex)
                HvAbu = .ClassAbundance
                udtERStats.TotalAbundance = udtERStats.TotalAbundance + .ClassAbundance
                udtERStats.TotalMemberCount = udtERStats.TotalMemberCount + .ClassCount
            End With
        
            Select Case eERCalcType
            Case ectER_RAT
                  udtERStats.ERAvg = RatER(LtAbu, HvAbu)
            Case ectER_LOG
                  udtERStats.ERAvg = LogER(LtAbu, HvAbu)
            Case ectER_ALT
                  udtERStats.ERAvg = AltER(LtAbu, HvAbu)
            End Select
            udtERStats.ERCount = 1
            
            blnValidER = True
        End If
    End If

    If mCorrectedAbundanceStatsCount > UBound(mCorrectedAbundanceStats) Then
        ' This code shouldn't normally be reached
        Debug.Assert False
        ReDim Preserve mCorrectedAbundanceStats((UBound(mCorrectedAbundanceStats) + 1) * 2 - 1)
    End If
    
    With mCorrectedAbundanceStats(lngPairIndex)
        If blnComputedERScanByScan Then
            ' Store the ScanAbu2PlusCorrected and ScanAbu4PlusCorrected values so they can be accessed later
            
            ' First step through udtScanByScanStats and count the number of scans with a light or heavy abundance value
            .AlignedScanCount = 0
            For lngScanIndex = 0 To udtScanByScanStats.ScanNumCount - 1
                If udtScanByScanStats.ScanDetailsLt(lngScanIndex).Abundance <> 0 Or _
                   udtScanByScanStats.ScanDetailsHv(lngScanIndex).Abundance <> 0 Then
                   .AlignedScanCount = .AlignedScanCount + 1
                End If
            Next lngScanIndex
            
            If .AlignedScanCount = 0 Then
                ReDim .AlignedScanStats(0)
            Else
                ReDim .AlignedScanStats(.AlignedScanCount - 1)
                
                .AlignedScanCount = 0
                For lngScanIndex = 0 To udtScanByScanStats.ScanNumCount - 1
                    If udtScanByScanStats.ScanDetailsLt(lngScanIndex).Abundance <> 0 Or _
                       udtScanByScanStats.ScanDetailsHv(lngScanIndex).Abundance <> 0 Then
                       With .AlignedScanStats(.AlignedScanCount)
                            .ScanNumber = udtScanByScanStats.ScanNumStart + lngScanIndex
                            .LightAbundance = udtScanByScanStats.ScanDetailsLt(lngScanIndex).Abundance
                            .HeavyAbundance = udtScanByScanStats.ScanDetailsHv(lngScanIndex).Abundance
                            .ER = udtScanByScanStats.ScanER(lngScanIndex)
                            .ScanAbu2PlusCorrected = udtScanByScanStats.ScanAbu2PlusCorrected(lngScanIndex)
                            .ScanAbu4PlusCorrected = udtScanByScanStats.ScanAbu4PlusCorrected(lngScanIndex)
                       End With
                       .AlignedScanCount = .AlignedScanCount + 1
                    End If
                Next lngScanIndex
            End If
        Else
            .AlignedScanCount = 0
        End If
    End With
    mCorrectedAbundanceStatsCount = mCorrectedAbundanceStatsCount + 1
    
    If Not blnValidER Then
        With udtERStats
            If .ERAvg <> 0 Then .ERAvg = ER_CALC_ERR
            .ERCount = 0
            .EROutlierCountRemoved = 0
            .ERStDev = 0
            .TotalAbundance = 0
            .TotalMemberCount = 0
        End With
    End If

    CalcDltLblPairsERWork = blnValidER

End Function

Private Function CalcDltLblPairsERScanByScan(ByRef udtScanByScanStats As udtScanByScanStatsType, ByVal lngGelIndex As Long, ByVal LtUMCIndex As Long, ByVal HvUMCIndex As Long, ByVal intTargetChargeLt As Integer, ByVal intTargetChargeHv As Integer, blnIReportData As Boolean) As Boolean
    ' Compares the data in the LC-MS Features given by LtUMCIndex and HvUMCIndex and computes ER values scan-by-scan
    ' If intTargetChargeLt or intTargetChargeHv is non-zero, then only uses the data with the given charge state;
    '  otherwise, uses all data in the UMC
    
    ' The steps:
    ' 1. Determine the minimum and maximum scan number in the light and heavy LC-MS Features
    '    For speed purposes, using the precomputed .MinScan and .MaxScan values,
    '     even though we'll probably reserve more memory than needed
    ' 2. Reserve space in udtScanByScanStats as needed
    ' 3. Populate the arrays
    ' 4. Look for scans with data in each LC-MS Features and computer an ER value
    '    If blnIReportData = True, then use the CalcDltLblPairsERViaIReport function to compute the value
    ' 5. Optionally, remove outlier ER values
    ' 6. Compute the ER values, weighting by abundance
    
    
    Dim lngScanIndex As Long
    Dim lngIndex As Long
    
    Dim blnUseMaxValueEachScan As Boolean
    Dim blnSuccess As Boolean
    Dim blnConvertedToSymmetricER As Boolean
    
    Dim lngERListCount As Long
    Dim dblERList() As Double                 ' Note that the ER values are converted to base-10 Log values prior to filtering with the Grubb's test filter
    Dim lngERListIndexPointers() As Long      ' Pointer of index that ER value comes from
    
    Dim dblAbu2DaCorrected As Double      ' Corrected monoisotopic + 2 Da abundance
    Dim dblAbu4DaCorrected As Double      ' Corrected monoisotopic + 4 Da abundance
    
    Dim dblERListLtOverHv() As Double
    
    Dim lngScanCountInvalidER As Long
    Dim dblFractionValid As Double
    
    Dim dblLightHeavyAbuSum As Double
    
    ' Note: When computing a weighted average, use ER = dblERProductSum / dblWeightingSum
    Dim dblERProductSum As Double
    Dim dblWeightingSum As Double
    
    Dim lngMinFinalValueCount As Long
    Dim lngEROutlierCountRemoved As Long
    Dim eConfidenceLevel As eclConfidenceLevelConstants

    Dim objStatDoubles As New StatDoubles
    Dim objOutlierFilter As New clsGrubbsTestOutlierFilter
    
On Error GoTo CalcDltLblPairsERScanByScanErrorHandler

    ' Make sure these variables are 0
    With udtScanByScanStats.ERStats
        .ERAvg = 0
        .ERStDev = 0
        .ERCount = 0
        .EROutlierCountRemoved = 0
        .TotalAbundance = 0
        .TotalMemberCount = 0
    End With

    ' Determine the minimum and maximum scan numbers
    With GelUMC(lngGelIndex).UMCs(LtUMCIndex)
        udtScanByScanStats.ScanNumStart = .MinScan
        udtScanByScanStats.ScanNumEnd = .MaxScan
    End With
    
    With GelUMC(lngGelIndex).UMCs(HvUMCIndex)
        If .MinScan < udtScanByScanStats.ScanNumStart Then udtScanByScanStats.ScanNumStart = .MinScan
        If .MaxScan > udtScanByScanStats.ScanNumEnd Then udtScanByScanStats.ScanNumEnd = .MaxScan
    End With
    
    If udtScanByScanStats.ScanNumStart < 0 Or udtScanByScanStats.ScanNumEnd < 0 Then
        blnSuccess = False
    Else
        ' Reserve space in udtScanByScanStats.ScanDetails() to hold the intensities by scan
        
        With udtScanByScanStats
            .ScanNumCount = .ScanNumEnd - .ScanNumStart + 1
            If .ScanNumCount <= 0 Then
                Debug.Assert False
                ' this shouldn't happen
                blnSuccess = False
            Else
                ReDim .ScanDetailsLt(.ScanNumCount - 1)
                ReDim .ScanDetailsHv(.ScanNumCount - 1)
                ReDim .ScanER(.ScanNumCount - 1)
                ReDim .ScanAbu2PlusCorrected(.ScanNumCount - 1)
                ReDim .ScanAbu4PlusCorrected(.ScanNumCount - 1)

                blnSuccess = True
            End If
        End With
    End If

    If blnSuccess Then
        ' Populate the .ScanDetails arrays
        blnUseMaxValueEachScan = GelP_D_L(lngGelIndex).SearchDef.UseIdenticalChargesForER
        
        CalcDltLblPairsScanByScanPopulate udtScanByScanStats.ScanDetailsLt, lngGelIndex, LtUMCIndex, intTargetChargeLt, udtScanByScanStats.ScanNumStart, blnUseMaxValueEachScan
        CalcDltLblPairsScanByScanPopulate udtScanByScanStats.ScanDetailsHv, lngGelIndex, HvUMCIndex, intTargetChargeHv, udtScanByScanStats.ScanNumStart, blnUseMaxValueEachScan
        
        ' Step through the scans and compute the ER value for each pair of scans with data in each scan
        ' If a valid ER value, then add to dblERList and lngERListIndexPointers
        
        With udtScanByScanStats
            lngERListCount = 0
            lngScanCountInvalidER = 0
            If .ScanNumCount > 0 Then
                ReDim dblERList(.ScanNumCount - 1)
                ReDim lngERListIndexPointers(.ScanNumCount - 1)
                
                For lngScanIndex = 0 To .ScanNumCount - 1
                    If .ScanDetailsLt(lngScanIndex).Abundance > 0 And .ScanDetailsHv(lngScanIndex).Abundance > 0 Then
                        If blnIReportData Then
                            .ScanER(lngScanIndex) = CalcDltLblPairsERViaIReport(lngGelIndex, .ScanDetailsLt(lngScanIndex), .ScanDetailsHv(lngScanIndex), dblAbu2DaCorrected, dblAbu4DaCorrected)
                        Else
                            dblAbu2DaCorrected = 0
                            dblAbu4DaCorrected = 0
                            .ScanER(lngScanIndex) = RatER(.ScanDetailsLt(lngScanIndex).Abundance, .ScanDetailsHv(lngScanIndex).Abundance)
                        End If
                        
                        .ScanAbu2PlusCorrected(lngScanIndex) = dblAbu2DaCorrected
                        .ScanAbu4PlusCorrected(lngScanIndex) = dblAbu4DaCorrected
                        
                        If .ScanER(lngScanIndex) <> ER_CALC_ERR Then
                            dblERList(lngERListCount) = .ScanER(lngScanIndex)
                            lngERListIndexPointers(lngERListCount) = lngScanIndex
                            
                            lngERListCount = lngERListCount + 1
                        Else
                            lngScanCountInvalidER = lngScanCountInvalidER + 1
                        End If
                    End If
                Next lngScanIndex
                
                If blnIReportData Then
                    If lngScanCountInvalidER + lngERListCount Then
                        ' Check if at least x% of the scans had a valid ER value
                        dblFractionValid = lngERListCount / (lngScanCountInvalidER + lngERListCount)
                        
                        If dblFractionValid < glbPreferencesExpanded.PairSearchOptions.SearchDef.IReportEROptions.MinimumFractionScansWithValidER Then
                            ' Not enough of the scan-by-scan pairs for this pair had valid ER values
                            ' Set lngERListCount to 1 and set its ER to ER_CALC_ERR
                            
                            lngERListCount = 1
                            dblERList(0) = ER_CALC_ERR
                            lngERListIndexPointers(0) = 0
                            
                            dblERProductSum = ER_CALC_ERR
                            dblWeightingSum = 1
                        End If

                    End If
                    
                End If
            End If
            
        End With
        
        ' Call objStatDoubles to compute the StDev and Count stats
        ' Use dblERProductSum and dblWeightingSum to compute a Weighted ER value
        blnConvertedToSymmetricER = False
        lngEROutlierCountRemoved = 0
        If lngERListCount > 0 Then
            
            If glbPreferencesExpanded.PairSearchOptions.OutlierRemovalUsesSymmetricERs Then
                ' Convert the ER values to symmetric ER values (0 means unchanged)
                If lngERListCount = 1 And dblERList(lngIndex) = ER_CALC_ERR Then
                    ' Do not convert to symmetric ER
                Else
                    For lngIndex = 0 To lngERListCount - 1
                        If dblERList(lngIndex) > 0 Then
                            dblERList(lngIndex) = AltERViaRatER(dblERList(lngIndex))
                        Else
                            Debug.Assert False
                            dblERList(lngIndex) = 0
                        End If
                    Next lngIndex
                    blnConvertedToSymmetricER = True
                End If
            End If
            
            ' Need to shrink dblERList and lngERListIndexPointers to only include the non-zero ER values
            ReDim Preserve dblERList(lngERListCount - 1)
            ReDim Preserve lngERListIndexPointers(lngERListCount - 1)
            
            If GelP_D_L(lngGelIndex).SearchDef.RemoveOutlierERs Then
                
                ' Remove outlier ER values
                lngMinFinalValueCount = GelP_D_L(lngGelIndex).SearchDef.RemoveOutlierERsMinimumDataPointCount
                If lngMinFinalValueCount < 2 Then lngMinFinalValueCount = 2
                
                If lngERListCount > lngMinFinalValueCount Then
                    Select Case GelP_D_L(lngGelIndex).SearchDef.RemoveOutlierERsConfidenceLevel
                    Case ecl97Pct
                        eConfidenceLevel = ecl97Pct
                    Case ecl99Pct
                        eConfidenceLevel = ecl99Pct
                    Case Else
                         ' Includes ecl95Pct
                         eConfidenceLevel = ecl95Pct
                    End Select
                    
                    objOutlierFilter.ConfidenceLevel = eConfidenceLevel
                    objOutlierFilter.MinFinalValueCount = lngMinFinalValueCount
                    objOutlierFilter.RemoveMultipleValues = GelP_D_L(lngGelIndex).SearchDef.RemoveOutlierERsIterate

                    If objOutlierFilter.RemoveOutliers(dblERList, lngERListIndexPointers, lngEROutlierCountRemoved) Then
                        ' Successfully removed outliers
                        
                        If lngEROutlierCountRemoved > 0 Then
                            'Debug.Print "Removed " & lngEROutlierCountRemoved & " outliers for pair with LC-MS Features " & LtUMCIndex & " and " & HvUMCIndex
                        End If
                        
                        lngERListCount = lngERListCount - lngEROutlierCountRemoved
                        If lngERListCount <> UBound(dblERList) + 1 Then
                            ' Array size doesn't agree with expected value
                            Debug.Assert False
                            lngERListCount = UBound(dblERList) + 1
                        End If
                    Else
                        Debug.Assert False
                    End If
                
                End If
            End If
            
            If lngERListCount > 1 Then
                ' Now that outliers have been removed, populate the weighting variables
                
                With udtScanByScanStats
                    dblLightHeavyAbuSum = 0
                    dblERProductSum = 0
                    dblWeightingSum = 0
                    
                    For lngIndex = 0 To lngERListCount - 1
                    
                        Debug.Assert dblERList(lngIndex) <> ER_CALC_ERR
                    
                        dblLightHeavyAbuSum = .ScanDetailsLt(lngERListIndexPointers(lngIndex)).Abundance + .ScanDetailsHv(lngERListIndexPointers(lngIndex)).Abundance
                        dblERProductSum = dblERProductSum + dblERList(lngIndex) * dblLightHeavyAbuSum
                      
                        dblWeightingSum = dblWeightingSum + dblLightHeavyAbuSum
                        
                        .ERStats.TotalAbundance = .ERStats.TotalAbundance + dblLightHeavyAbuSum
                        .ERStats.TotalMemberCount = .ERStats.TotalMemberCount + .ScanDetailsLt(lngERListIndexPointers(lngIndex)).MemberCount + .ScanDetailsHv(lngERListIndexPointers(lngIndex)).MemberCount
                
                    Next lngIndex
                End With
            Else
                If dblERList(0) <> ER_CALC_ERR Then
                    dblERProductSum = dblERList(0)
                
                    udtScanByScanStats.ERStats.TotalAbundance = udtScanByScanStats.ScanDetailsLt(lngERListIndexPointers(0)).Abundance + udtScanByScanStats.ScanDetailsHv(lngERListIndexPointers(0)).Abundance
                    udtScanByScanStats.ERStats.TotalMemberCount = udtScanByScanStats.ScanDetailsLt(lngERListIndexPointers(0)).MemberCount + udtScanByScanStats.ScanDetailsHv(lngERListIndexPointers(0)).MemberCount
                
                End If
                dblWeightingSum = 1
            End If
            
            ' Note: If .OutlierRemovalUsesSymmetricERs, then objStatDoubles
            '       is filled with the symmetric ER values
            
            If lngERListCount = 1 And dblERList(0) = ER_CALC_ERR Then
                blnSuccess = False
            Else
                blnSuccess = objStatDoubles.Fill(dblERList)
            End If
            
            If blnSuccess Then
                With udtScanByScanStats.ERStats
                    If dblWeightingSum > 0 Then
                        ' Store a weighed mean ER
                        .ERAvg = dblERProductSum / dblWeightingSum
                    Else
                        ' This shouldn't happen
                        Debug.Assert False
                        
                        ' Could use the following to store an unweighted mean ER
                        .ERAvg = objStatDoubles.Mean
                    End If
                    
                    .ERCount = objStatDoubles.Count
                    .EROutlierCountRemoved = lngEROutlierCountRemoved
                    
                    .ERStDev = objStatDoubles.StDev
                    If blnConvertedToSymmetricER Then
                        ' If .erAvg >=0 then we can use objStatDoubles.StDev for the StDev
                        ' Otherwise, need to recompute the StDev using the regular light/heavy values
                        If .ERAvg < 0 Then
                            ReDim dblERListLtOverHv(lngERListCount - 1)
                            For lngIndex = 0 To lngERListCount - 1
                                dblERListLtOverHv(lngIndex) = RatERViaAltER(dblERList(lngIndex))
                            Next lngIndex
                            
                            If objStatDoubles.Fill(dblERListLtOverHv) Then
                                .ERStDev = objStatDoubles.StDev
                            End If
                        End If
                    End If
                
                    If blnConvertedToSymmetricER And .ERAvg <> ER_CALC_ERR Then
                        ' Convert the ER value back to regular light/heavy values
                        ' Note: Do not convert the .ERStDev value (this was handled above)
                        .ERAvg = RatERViaAltER(.ERAvg)
                    End If
                
                End With
            Else
                udtScanByScanStats.ERStats.ERAvg = ER_CALC_ERR
            End If
        Else
            ' No overlapping scans; store 0 but set blnSuccess = False
            With udtScanByScanStats.ERStats
                .ERAvg = 0
                .ERStDev = 0
                .ERCount = 0
                .EROutlierCountRemoved = 0
            End With
            blnSuccess = False
        End If
    Else
        udtScanByScanStats.ERStats.ERAvg = ER_CALC_ERR
    End If
    
    CalcDltLblPairsERScanByScan = blnSuccess
    Exit Function

CalcDltLblPairsERScanByScanErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in CalcDltLblPairsERScanByScan: " & Err.Description
    Else
        LogErrors Err.Number, "CalcDltLblPairsERScanByScan"
    End If
    
    udtScanByScanStats.ERStats.ERAvg = ER_CALC_ERR
End Function

Private Function CalcDltLblPairsERViaIReport(ByVal lngGelIndex As Long, ByRef udtScanDetailsLt As udtScanByScanDetailsType, ByRef udtScanDetailsHv As udtScanByScanDetailsType, ByRef dblAbu2DaCorrected As Double, ByRef dblAbu4DaCorrected As Double) As Double

    ' Use udtScanDetailsLt.MemberIndexMostAbu and udtScanDetailsHv.MemberIndexMostAbu
    '  to compute corrected abundances for the MonoIsotopic peak, the M + 2Da peak, and the M + 4 Da peak
    
    Dim dblMonoIsoMass As Double
    
    Dim dblAbuMonoIso As Double     ' Raw monoisotopic abundance
    Dim dblAbu2DaRaw As Double      ' Raw (uncorrected) monoisotopic + 2 Da abundance
    Dim dblAbu4DaRaw As Double      ' Raw (uncorrected) monoisotopic + 4 Da abundance
    
    Dim dblM2OverM0 As Double
    Dim dblM4OverM0 As Double
    
    Dim dblER As Double
    
    Dim blnError As Boolean
    
On Error GoTo CalcDltLblPairsERViaIReportErrorHandler

    dblER = ER_CALC_ERR
    blnError = False
    
    With GelData(lngGelIndex)
        Select Case udtScanDetailsLt.MemberTypeMostAbu
        Case gldtCS
            ' This code shouldn't be reached
            ' We cannot compute I Report values with CS data
            Debug.Assert False
            blnError = True
        Case gldtIS
            dblMonoIsoMass = .IsoData(udtScanDetailsLt.MemberIndexMostAbu).MonoisotopicMW
            dblAbuMonoIso = .IsoData(udtScanDetailsLt.MemberIndexMostAbu).IntensityMono
            dblAbu2DaRaw = .IsoData(udtScanDetailsLt.MemberIndexMostAbu).IntensityMonoPlus2
        Case Else
            blnError = True
        End Select
    End With
    
    If Not blnError Then
        With GelData(lngGelIndex)
            Select Case udtScanDetailsHv.MemberTypeMostAbu
            Case gldtCS
                ' This code shouldn't be reached
                ' We cannot compute I Report values with CS data
                Debug.Assert False
                blnError = True
            Case gldtIS
                dblAbu4DaRaw = .IsoData(udtScanDetailsHv.MemberIndexMostAbu).IntensityMono
            Case Else
                blnError = True
            End Select
        End With
        
        If Not blnError Then
            ' Uncomment for debugging
            'Debug.Print dblMonoIsoMass & vbTab & dblAbuMonoIso & vbTab & dblAbu2DaRaw & vbTab & dblAbu4DaRaw
            
            With glbPreferencesExpanded.PairSearchOptions.SearchDef.IReportEROptions
                With .NaturalAbundanceRatio2Coeff
                    dblM2OverM0 = .Multiplier * dblMonoIsoMass ^ .Exponent
                End With
                
                With .NaturalAbundanceRatio4Coeff
                    dblM4OverM0 = .Multiplier * dblMonoIsoMass ^ .Exponent
                End With
                
                dblAbu2DaCorrected = dblAbu2DaRaw - dblM2OverM0 * dblAbuMonoIso
                
                dblAbu4DaCorrected = dblAbu4DaRaw - dblM4OverM0 * dblAbuMonoIso
                
                If dblAbu4DaCorrected <= 0 Then
                    ' Condition 1
                    ' False pair
                    dblER = ER_CALC_ERR
                ElseIf dblAbu2DaCorrected < 0.15 * dblAbu4DaCorrected Then
                    ' Condition 2
                    udtScanDetailsHv.Abundance = Round(1.15 * dblAbu4DaCorrected, 0)
                    dblER = dblAbuMonoIso / udtScanDetailsHv.Abundance
                Else
                    ' Condition 3
                    If dblAbu2DaCorrected + dblAbu4DaCorrected > 0 Then
                        udtScanDetailsHv.Abundance = Round(dblAbu2DaCorrected + dblAbu4DaCorrected, 0)
                        dblER = dblAbuMonoIso / udtScanDetailsHv.Abundance
                    Else
                        dblER = ER_CALC_ERR
                    End If
                End If
                
            End With
        End If
    End If
    
    
    CalcDltLblPairsERViaIReport = dblER
    
    Exit Function

CalcDltLblPairsERViaIReportErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in CalcDltLblPairsERViaIReport: " & Err.Description
    Else
        LogErrors Err.Number, "CalcDltLblPairsERViaIReport"
    End If
    
    CalcDltLblPairsERViaIReport = ER_CALC_ERR
End Function

Private Sub CalcDltLblPairsScanByScanPopulate(ByRef udtScanDetails() As udtScanByScanDetailsType, ByVal lngGelIndex As Long, ByVal UMCIndex As Long, ByVal intTargetCharge As Integer, ByVal lngScanByScanStatsScanNumStart As Long, blnUseMaxValueEachScan As Boolean)
    ' Note: The algorithms in this function are the same as those in
    '       frmPairBrowser.PopulateUMCAbuData
    
    Dim lngMemberIndex As Long
    Dim lngScan As Long, lngScanIndex As Long
    Dim intCharge As Integer
    Dim dblAbu As Double
    Dim dblMaxAbu() As Double
    
On Error GoTo CalcDltLblPairsScanByScanPopulateErrorHandler

    With GelUMC(lngGelIndex).UMCs(UMCIndex)
        ReDim dblMaxAbu(UBound(udtScanDetails))
        
        For lngScanIndex = 0 To UBound(udtScanDetails)
            dblMaxAbu(lngScanIndex) = ER_CALC_ERR
        Next lngScanIndex
        
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
                lngScanIndex = lngScan - lngScanByScanStatsScanNumStart
                If lngScanIndex < 0 Then
                    ' This shouldn't happen
                    Debug.Assert False
                Else
                    If blnUseMaxValueEachScan Then
                        If dblAbu > udtScanDetails(lngScanIndex).Abundance Then
                            udtScanDetails(lngScanIndex).Abundance = dblAbu
                            udtScanDetails(lngScanIndex).MemberCount = 1
                        End If
                    Else
                        udtScanDetails(lngScanIndex).Abundance = udtScanDetails(lngScanIndex).Abundance + dblAbu
                        udtScanDetails(lngScanIndex).MemberCount = udtScanDetails(lngScanIndex).MemberCount + 1
                    End If
                
                    
                    If dblAbu > dblMaxAbu(lngScanIndex) Or lngMemberIndex = 0 Then
                        dblMaxAbu(lngScanIndex) = dblAbu
                        udtScanDetails(lngScanIndex).MemberTypeMostAbu = .ClassMType(lngMemberIndex)
                        udtScanDetails(lngScanIndex).MemberIndexMostAbu = .ClassMInd(lngMemberIndex)
                    End If
                
                End If
            End If
        Next lngMemberIndex
    End With

    Exit Sub

CalcDltLblPairsScanByScanPopulateErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in CalcDltLblPairsScanByScanPopulate: " & Err.Description
    Else
        LogErrors Err.Number, "CalcDltLblPairsScanByScanPopulate"
    End If

End Sub

Public Sub GetScanByScanStatsForPair(ByVal lngPairIndex As Long, ByRef lngAlignedScanCount As Long, ByRef lngScanNumbers() As Long, ByRef dblLtAbundance() As Double, ByRef dblHvAbundance() As Double, ByRef dblER() As Double, ByRef sngAbu2DaCorrected() As Single, ByRef sngAbu4DaCorrected() As Single)
    Dim lngScanIndex As Long
    
    If lngPairIndex < mCorrectedAbundanceStatsCount Then
        With mCorrectedAbundanceStats(lngPairIndex)
            lngAlignedScanCount = .AlignedScanCount
            
            If lngAlignedScanCount >= UBound(lngScanNumbers) Then
                ReDim lngScanNumbers(lngAlignedScanCount - 1)
                ReDim dblLtAbundance(lngAlignedScanCount - 1)
                ReDim dblHvAbundance(lngAlignedScanCount - 1)
                ReDim dblER(lngAlignedScanCount - 1)
                ReDim sngAbu2DaCorrected(lngAlignedScanCount - 1)
                ReDim sngAbu4DaCorrected(lngAlignedScanCount - 1)
            End If
            
            For lngScanIndex = 0 To .AlignedScanCount - 1
                With .AlignedScanStats(lngScanIndex)
                    lngScanNumbers(lngScanIndex) = .ScanNumber
                    dblLtAbundance(lngScanIndex) = .LightAbundance
                    dblHvAbundance(lngScanIndex) = .HeavyAbundance
                    dblER(lngScanIndex) = .ER
                    sngAbu2DaCorrected(lngScanIndex) = .ScanAbu2PlusCorrected
                    sngAbu4DaCorrected(lngScanIndex) = .ScanAbu4PlusCorrected
                End With
            Next lngScanIndex
        End With
    Else
        lngAlignedScanCount = 0
    End If
End Sub

Private Sub InitializeLocalVariables(ByVal lngPairCount As Long)
    mCorrectedAbundanceStatsCount = 0
    
    If lngPairCount < 1 Then
        ReDim mCorrectedAbundanceStats(9)
    Else
        ReDim mCorrectedAbundanceStats(lngPairCount - 1)
    End If
End Sub

Private Function AltERViaRatER(ByVal dblRatER As Double) As Double
    '------------------------------------------------------
    'Converts ratio-based ER to shifted symmetric abundance ratio
    '------------------------------------------------------
    On Error GoTo err_AltER
    If dblRatER >= 1 Then
       AltERViaRatER = dblRatER - 1
    Else
       AltERViaRatER = 1 - (1 / dblRatER)
    End If
    Exit Function
    
err_AltER:
    AltERViaRatER = ER_CALC_ERR
End Function

Private Function LogERViaRatER(ByVal dblRatER As Double) As Double
    '------------------------------------------------------
    'Converts ratio-based ER to a logarithmic scale
    '------------------------------------------------------
    On Error Resume Next
    If dblRatER > 0 Then
        LogERViaRatER = Log(dblRatER)
        If Err Then LogERViaRatER = ER_CALC_ERR
    Else
        LogERViaRatER = ER_CALC_ERR
    End If
End Function



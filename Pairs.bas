Attribute VB_Name = "Module18"
'isotopic labeling analysis definitions, procedures
'last modified 05/31/2002 nt
'--------------------------------------------------
Option Explicit

Public Const ER_CALC_ERR As Double = -1E+307
Public Const ER_NO_RATIO As Single = 0

Public Const PAIR_ICAT = 0
Public Const PAIR_N14N15 = 1

'''Public Const PAIR_L_MARK = "(L)"
'''Public Const PAIR_H_MARK = "(H)"
Public Const PAIR_DLT_MARK = "DLT:"

Public Const glLMTYPE_PAIRS_N14 = 0
'''Public Const glLMTYPE_PAIRS_N15 = 1
'''Public Const glLMTYPE_SINGLES_N14 = 2
'''Public Const glLMTYPE_SINGLES_N15 = 3
'''
'''Public Const glER_Light As Double = 99.99999
'''Public Const glER_Heavy As Double = 0.00001

'LblDlt pairing labels - keep UMC and individual separated
'''Public Const glPDL_NONE = 0

'''Public Const glPDL_N14_N15_PEO_SOLO = 1001
'''Public Const glPDL_ICAT_SOLO = 1002
'''Public Const glPDL_N14_N15_SOLO = 1003
'''
'''Public Const glPDL_N14_N15_PEO_UMC = 2001
'''Public Const glPDL_ICAT_UMC = 2002
'''Public Const glPDL_N14_N15_UMC = 2003

'next set of constants might be interesting for database search
'it shows what information should be matched during search

Public Const glLBL_NONE = 0     'indicates label not used
'''Public Const glLBL_CYS = 1      'cysteine label
'''Public Const glLBL_LYS = 2      'lysine label

Public Const glDLT_NONE = 0     'indicates delta not used
'''Public Const glDLT_N14_N15 = 1  'N14/N15 delta
'''Public Const glDLT_C12_C13 = 2  'C12/C13 delta
'LblDlt expression ratio constants
Public Const glER_NONE = 0              'no ER applied

'expression ratio constants
Public Enum ectERCalcTypeConstants
    ectER_RAT = 0           'lt/hv
    ectER_LOG = 1           'log(lt/hv)
    ectER_ALT = 2           'lt/hv - 1 if lt/hv>=1; otherwise 1-hv/lt
End Enum
                                        
Public Const glER_SOLO_RAT = 1          'lt/hv
Public Const glER_SOLO_LOG = 2          'log(lt/hv)
Public Const glER_SOLO_ALT = 3          'lt/hv - 1 if lt/hv>=1
                                        '1-hv/lt   if lt/hv<1
Public Const glER_UMC_REP_RAT = 4       'use class representative
Public Const glER_UMC_REP_LOG = 5       'intensity
Public Const glER_UMC_REP_ALT = 6

Public Const glER_UMC_AVG_RAT = 7       'use class members average
Public Const glER_UMC_AVG_LOG = 8       'intensity
Public Const glER_UMC_AVG_ALT = 9

Public Const glER_UMC_SUM_RAT = 10      'use sum of class members
Public Const glER_UMC_SUM_LOG = 11      'intensities
Public Const glER_UMC_SUM_ALT = 12

Public Const glPair_All = 0             ' Included, excluded, and neutral pairs
Public Const glPAIR_Inc = 1             'included
Public Const glPAIR_Neu = 0             'neutral (upon initialization)
Public Const glPAIR_Exc = -1            'excluded

Public Enum umcpUMCPairMembershipConstants
    umcpNone = 0                    'UMC is not member of any pair
    umcpLightUnique = 1             'UMC is light member of unique pair
    umcpLightMultiple = 2           'UMC is light member of multiple pairs
    umcpHeavyUnique = 3             'UMC is heavy member of unique pair
    umcpHeavyMultiple = 4           'UMC is heavy member of multiple pairs
    umcpLightHeavyMix = 5           'UMC is light member of at least one pair and heavy member of at least one pair
End Enum

Public Type PairDefinition
   Delta As Double                  'delta of isotopic label(neutron mass)
   DeltaTol As Double               'delta tolerance(Da) when detecting pairs
   Case2500 As Boolean              'for masses>2500 Da error of -1 nitrogen atom
                                    'could be introduced in deconvolution
                                    'if True that case is also considered when looking for pairs
   StopAfterEachScan As Boolean
   SearchForSingles As Boolean      'if True singles(distributions detected
                                    'as N14 or N15 will also be counted and
                                    'marked
   SaveN15Singles As Boolean        'if True singles detected as N15
                                    'will during save operation get
                                    'tagged with AMT mark(marked as identified)
   SinglesMMA As Double             'MMA(ppm) for singles determination
                                    'it should match MMA used during AMT
                                    'search because for N14 that search results
                                    'is used here; for N15 search is performed
                                    'with this argument
   MultiAMTHits As Integer          '0 ignore pairs/singles when hit
                                    'with multiple AMTs
                                    '1 pick best fit if hit with multiple AMTs
                                    '2 pick whatever matches number of nitrogen atoms
   MassLockType As Integer
End Type

Public Type ERStatHelper    'just helper type for ER statistics
    ERCnt As Long           'total count
    ERBadL As Long          'out of left ER range
    ERBadR As Long          'out of right ER range
    ERAvg As Double         'average of "in range" ERs
    ERStD As Double         'standard deviation of "in range" ERs
End Type

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
    
    ERStats As udtERStatsType
End Type

Public PairDef As PairDefinition

'-----------------------------------------------------------------------
'The idea behind DltLbl pairs is that Delta stands for something
'that is different 'the same way' in heavy and light member of pair
'while Label stands for something that is different in 'different way'
'For example ICAT would be label because number of Cys labeled in
'light and heavy memeber does not have to match (although we can explore
'that case by treating ICAT D8-D0 as Delta) while N14/N15 pairs would
'be Delta since all N14 atoms in light member are replaced by N15 in
'heavy member
'-----------------------------------------------------------------------


Public Function InitDltLblPairs(ByVal Ind As Long) As Boolean

'--------------------------------------------------------
'initializes pairs information for N14/N15 Lbl analysis
'initially reserve space for 40000 pairs
'space for ER should be reserved as needed
'--------------------------------------------------------
On Error GoTo err_InitDltLblPairs
With GelP_D_L(Ind)
    .DltLblType = ptNone
    .SearchDef.ERCalcType = glER_NONE
    .PCnt = 0
    ReDim .Pairs(40000)
End With
InitDltLblPairs = True
Exit Function

err_InitDltLblPairs:
If Err.Number = 7 Then  'out of memory; try to recover
    DestroyDltLblPairs Ind, False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Not enough memory for this operation.", vbOKOnly
    End If
End If
LogErrors Err.Number, "InitDltLblPairs"
End Function


Public Function AddDltLblPairs(ByVal Ind As Long, _
                               ByVal AddRate As Long) As Boolean
'------------------------------------------------------------------
'increase size for Delta Labeled arrays
'------------------------------------------------------------------
Dim NewSize As Long
On Error GoTo err_AddDltLblPairs
With GelP_D_L(Ind)
    NewSize = .PCnt + AddRate
    ReDim Preserve .Pairs(NewSize)
End With
AddDltLblPairs = True
Exit Function

err_AddDltLblPairs:
If Err.Number = 7 Then  'out of memory; try to recover
    DestroyDltLblPairs Ind, False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Not enough memory for this operation.", vbOKOnly
    End If
End If
LogErrors Err.Number, "AddDltLblPairs"
End Function

Public Sub ResetERValues(ByVal Ind As Long)
    Dim i As Long
    With GelData(Ind)
        For i = 1 To .CSLines
            .CSData(i).ExpressionRatio = ER_NO_RATIO
        Next i
        For i = 1 To .IsoLines
            .IsoData(i).ExpressionRatio = ER_NO_RATIO
        Next i
    End With
End Sub

Public Function TrimDltLblPairs(ByVal Ind As Long) As Boolean
'---------------------------------------------------------------
'this procedure should be called after processing is done;
'arrays are trimmed/erased to free some memeory
'---------------------------------------------------------------
On Error GoTo err_TrimDltLblPairs
With GelP_D_L(Ind)
    If .PCnt > 0 Then
       ReDim Preserve .Pairs(.PCnt - 1)
    Else
       DestroyDltLblPairs Ind, False
    End If
End With
TrimDltLblPairs = True
Exit Function

err_TrimDltLblPairs:
LogErrors Err.Number, "TrimDltLblPairs"
End Function

Public Sub DestroyDltLblPairs(ByVal Ind As Long, Optional blnUpdateAnalysisHistory As Boolean = True)

On Error GoTo DestroyPairsResume
    If GelP_D_L(Ind).PCnt > 0 Then
        If blnUpdateAnalysisHistory Then
            AddToAnalysisHistory Ind, "Cleared all pairs"
        End If
    End If
    
DestroyPairsResume:
    DestroyDltLblPairsLocal Ind
End Sub

Private Sub DestroyDltLblPairsLocal(ByVal Ind As Long)
'--------------------------------------------------
'erases N14/N15 Lbl arrays
'--------------------------------------------------
On Error Resume Next
With GelP_D_L(Ind)
    .DltLblType = ptNone
    .SearchDef.ERCalcType = glER_NONE
    .PCnt = 0
    Erase .Pairs
End With
End Sub

Public Sub InitDltLblPairsER(ByVal Ind As Long, _
                             Optional ByVal ERCalcType As ectERCalcTypeConstants = glER_NONE, _
                             Optional blnAddEntryToAnalysisHistory As Boolean = True)
'----------------------------------------------------------
'creates space for expression ratio; this operation should
'be called before any ER operation
'----------------------------------------------------------
On Error Resume Next
Dim i As Long
With GelP_D_L(Ind)
    .SearchDef.ERCalcType = ERCalcType
    If .PCnt > 0 Then
        For i = 0 To .PCnt - 1
            With .Pairs(i)
                .ER = 0
                .ERStDev = 0
                .ERChargeStateBasisCount = 0
                ReDim .ERChargesUsed(0)
                .ERMemberBasisCount = 0
            End With
        Next i
    End If
End With

' Initialize all of the ER entries to 0
ResetERValues Ind

If blnAddEntryToAnalysisHistory Then AddToAnalysisHistory Ind, "Cleared pair expression ratios"
End Sub


Public Sub CalcDltLblPairsER_UMC(ByVal Ind As Long, ByRef strMessage As String)
    '----------------------------------------------------------
    'creates space for and calculates expression ratio for UMC-based pairs
    '----------------------------------------------------------
    Dim lngPairIndex As Long
    Dim lngPairCountWithOutliers As Long
    Dim dblOutlierFractionSum As Double
    
    Call InitDltLblPairsER(Ind, GelP_D_L(Ind).SearchDef.ERCalcType)
    
    lngPairCountWithOutliers = 0
    dblOutlierFractionSum = 0
    
    For lngPairIndex = 0 To GelP_D_L(Ind).PCnt - 1
        CalcDltLblPairsER_UMCSinglePair Ind, lngPairIndex, lngPairCountWithOutliers, dblOutlierFractionSum
    Next lngPairIndex
  
    strMessage = "Calculated pair expression ratios; Pair count = " & GelP_D_L(Ind).PCnt
    If lngPairCountWithOutliers > 0 Then
        strMessage = strMessage & "; Removed outlier ER values from " & lngPairCountWithOutliers & " pairs"
        strMessage = strMessage & "; Average % of ER values of pair that were removed as outliers = " & Round(dblOutlierFractionSum / CDbl(lngPairCountWithOutliers) * 100, 2) & "%"
    End If
    
    AddToAnalysisHistory Ind, strMessage
    
End Sub

Private Sub CalcDltLblPairsER_UMCSinglePair(ByVal Ind As Long, ByVal lngPairIndex As Long, ByRef lngPairCountWithOutliers As Long, ByRef dblOutlierFractionSum As Double)
    
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
    
    eWeightingMode = GelP_D_L(Ind).SearchDef.AverageERsWeightingMode
    If (GelData(Ind).DataStatusBits And GEL_DATA_STATUS_BIT_IREPORT) = GEL_DATA_STATUS_BIT_IREPORT Then
        ' Enable I-Report based ER computation
        ' This will only be used if .ComputeERScanByScan = True
        blnIReportData = glbPreferencesExpanded.PairSearchOptions.SearchDef.IReportEROptions.Enabled
    Else
        blnIReportData = False
    End If
    
    If GelP_D_L(Ind).SearchDef.RequireMatchingChargeStatesForPairMembers And _
       GelP_D_L(Ind).SearchDef.UseIdenticalChargesForER And _
       GelP_D_L(Ind).SearchDef.AverageERsAllChargeStates Then
        ' Need to compute ER value for each charge state matching between the UMC's, then average them together with a weighted average
        
        udtUMCA = GelUMC(Ind).UMCs(GelP_D_L(Ind).Pairs(lngPairIndex).P1)
        udtUMCB = GelUMC(Ind).UMCs(GelP_D_L(Ind).Pairs(lngPairIndex).P2)
        
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
                        blnValidER = CalcDltLblPairsERWork(Ind, GelP_D_L(Ind).Pairs(lngPairIndex).P1, GelP_D_L(Ind).Pairs(lngPairIndex).P2, .Charge, ectER_RAT, .ERStats, blnIReportData)
                   End With
                   
                   If blnValidER Then intChargeStateStatsCount = intChargeStateStatsCount + 1
                   
                End If
            Next intChargeStateIndexB
        Next intChargeStateIndexA

        If intChargeStateStatsCount = 0 Then
            ' No charge states were found in common between the two UMC's in this pair
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
                    Select Case GelP_D_L(Ind).SearchDef.ERCalcType
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
        blnValidER = CalcDltLblPairsERWork(Ind, GelP_D_L(Ind).Pairs(lngPairIndex).P1, GelP_D_L(Ind).Pairs(lngPairIndex).P2, 0, GelP_D_L(Ind).SearchDef.ERCalcType, udtERStats, blnIReportData)
        blnAveragedAcrossChargeStates = False
        intChargeStateStatsCount = 1
        
        If Not blnValidER Then
            udtERStats.ERAvg = ER_CALC_ERR
        End If
    End If

    ' Store udtERStats.ERAvg
    With GelP_D_L(Ind).Pairs(lngPairIndex)
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

Private Function CalcDltLblPairsERWork(ByVal lngGelIndex As Long, ByVal LtUMCIndex As Long, ByVal HvUMCIndex As Long, ByVal intChargeStateToMatch As Integer, eERCalcType As Integer, ByRef udtERStats As udtERStatsType, blnIReportData As Boolean) As Boolean
    ' Computes and returns an ER value and associated stats in udtERStats
    ' Function returns True if success, false if an error
    ' Sets udtERStats.ERAvg = ER_CALC_ERR if an error occurs
    
    ' If intChargeStateToMatch is > 0 then only uses the specified charge state values in the two UMC's
    ' If intChargeStateToMatch is > 0 and the given charge state is only present in one of the UMC's, then returns an ER value of ER_CALC_ERR and sets udtERStats.TotalAbundance to 0 and udtERStats.TotalMemberCount to 0

    Dim LtAbu As Double, HvAbu As Double
    
    Dim intLtUMCCharge As Integer, intHvUMCCharge As Integer
    Dim intIndex As Integer, intIndexCompare As Integer
    Dim intLtUMCChargeStateIndex As Integer
    Dim intHVUMCChargeStateIndex As Integer
    
    Dim blnMatchFound As Boolean, blnValidER As Boolean
    
    Dim lngChargeIndexToTest() As Long              ' 0-based array of charge indices in the light UMC to test; sorted parallel with dblChargeIndexIntensities
    Dim dblChargeIndexIntensities() As Double       ' 0-based array of the intensities for the charge indices in the light UMC
    Dim intChargeToFind As Integer
    
    ' This udt is used when computing stats scan-by-scan
    Dim udtScanByScanStats As udtScanByScanStatsType
    Dim intTargetChargeLt As Integer, intTargetChargeHv As Integer
    
    Dim objSort As QSDouble
    
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
                    ' Both UMC's have the same representative charge state
                    blnMatchFound = True
                Else
                    ' The UMC's do not have the same representative charge state
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
    ' Compares the data in the UMC's given by LtUMCIndex and HvUMCIndex and computes ER values scan-by-scan
    ' If intTargetChargeLt or intTargetChargeHv is non-zero, then only uses the data with the given charge state;
    '  otherwise, uses all data in the UMC
    
    ' The steps:
    ' 1. Determine the minimum and maximum scan number in the light and heavy UMC's
    '    For speed purposes, using the precomputed .MinScan and .MaxScan values,
    '     even though we'll probably reserve more memory than needed
    ' 2. Reserve space in udtScanByScanStats as needed
    ' 3. Populate the arrays
    ' 4. Look for scans with data in each UMC's and computer an ER value
    '    If blnIReportData = True, then use the CalcDltLblPairsERViaIReport function to compute the value
    ' 5. Optionally, remove outlier ER values
    ' 6. Compute the ER values, weighting by abundance
    
    
    Dim lngScanIndex As Long
    Dim lngIndex As Long
    
    Dim blnUseMaxValueEachScan As Boolean
    Dim blnSuccess As Boolean
    
    Dim lngERListCount As Long
    Dim dblERList() As Double                 ' Note that the ER values are converted to base-10 Log values prior to filtering with the Grubb's test filter
    Dim lngERListIndexPointers() As Long      ' Pointer of index that ER value comes from
    
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
                            .ScanER(lngScanIndex) = CalcDltLblPairsERViaIReport(lngGelIndex, .ScanDetailsLt(lngScanIndex), .ScanDetailsHv(lngScanIndex))
                        Else
                            .ScanER(lngScanIndex) = RatER(.ScanDetailsLt(lngScanIndex).Abundance, .ScanDetailsHv(lngScanIndex).Abundance)
                        End If
                        
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
                            'Debug.Print "Removed " & lngEROutlierCountRemoved & " outliers for pair with UMCs " & LtUMCIndex & " and " & HvUMCIndex
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
                    
                    .ERStDev = objStatDoubles.StDev
                    .ERCount = objStatDoubles.Count
                    .EROutlierCountRemoved = lngEROutlierCountRemoved
                
                    If glbPreferencesExpanded.PairSearchOptions.OutlierRemovalUsesSymmetricERs And .ERAvg <> ER_CALC_ERR Then
                        ' Convert the ER value back to regular light/heavy values
                        ' Note: Do not need to convert the .ERStDev value, since it is already correct
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

Private Function CalcDltLblPairsERViaIReport(lngGelIndex As Long, ByRef udtScanDetailsLt As udtScanByScanDetailsType, ByRef udtScanDetailsHv As udtScanByScanDetailsType) As Double

    ' Use udtScanDetailsLt.MemberIndexMostAbu and udtScanDetailsHv.MemberIndexMostAbu
    '  to compute corrected abundances for the MonoIsotopic peak, the M + 2Da peak, and the M + 4 Da peak
    
    Dim dblMonoIsoMass As Double
    
    Dim dblAbuMonoIso As Double     ' Raw monoisotopic abundance
    Dim dblAbu2DaRaw As Double      ' Raw (uncorrected) monoisotopic + 2 Da abundance
    Dim dblAbu4DaRaw As Double      ' Raw (uncorrected) monoisotopic + 4 Da abundance
    
    Dim dblAbu2DaCorrected As Double      ' Corrected monoisotopic + 2 Da abundance
    Dim dblAbu4DaCorrected As Double      ' Corrected monoisotopic + 4 Da abundance
    
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

Public Sub CalcDltLblPairsER_Solo(ByVal Ind As Long)
'-----------------------------------------------------
'creates space for and calculates expression ratio for
'individual distribution pairs - solo pairs are always
'of isotopic peak type
'-----------------------------------------------------
Dim i As Long
GelP_D_L(Ind).SearchDef.ERCalcType = glbPreferencesExpanded.PairSearchOptions.SearchDef.ERCalcType
Call InitDltLblPairsER(Ind, GelP_D_L(Ind).SearchDef.ERCalcType, False)
With GelP_D_L(Ind)
  Select Case .SearchDef.ERCalcType
  Case ectER_RAT
    For i = 0 To .PCnt - 1
        .Pairs(i).ER = RatER(GelData(Ind).IsoData(.Pairs(i).P1).Abundance, _
                           GelData(Ind).IsoData(.Pairs(i).P2).Abundance)
        .Pairs(i).ERMemberBasisCount = 1
    Next i
  Case ectER_LOG
    For i = 0 To .PCnt - 1
        .Pairs(i).ER = LogER(GelData(Ind).IsoData(.Pairs(i).P1).Abundance, _
                           GelData(Ind).IsoData(.Pairs(i).P2).Abundance)
        .Pairs(i).ERMemberBasisCount = 1
    Next i
  Case ectER_ALT
    For i = 0 To .PCnt - 1
        .Pairs(i).ER = AltER(GelData(Ind).IsoData(.Pairs(i).P1).Abundance, _
                           GelData(Ind).IsoData(.Pairs(i).P2).Abundance)
        .Pairs(i).ERMemberBasisCount = 1
    Next i
  End Select
End With
AddToAnalysisHistory Ind, "Recalculated expression ratios"
End Sub

' Unused Function (May 2003)
'''Public Sub ClearDltLblPairsER(ByVal Ind As Long)
''''-----------------------------------------------
''''deletes expression ratio for delta label pairs
''''-----------------------------------------------
'''On Error Resume Next
'''GelP_D_L(Ind).SearchDef.ERCalcType = glER_NONE
'''Erase GelP_D_L(Ind).P1P2ER
'''Erase GelP_D_L(Ind).P1P2ERStDev
'''Erase GelP_D_L(Ind).P1P2ERBasisCount
'''End Sub


Public Function GetERDesc(ByVal ERCalcType As Long) As String
'--------------------------------------------------------
'returns description of ER type code
'--------------------------------------------------------
Select Case ERCalcType
Case glER_NONE
    GetERDesc = "None"
Case glER_SOLO_RAT
    GetERDesc = "Individual distributions; Ratio (Light/Heavy)"
Case glER_SOLO_LOG
    GetERDesc = "Individual distributions; Logarithmic"
Case glER_SOLO_ALT
    GetERDesc = "Individual distributions; Symmetric Ratio"
Case glER_UMC_REP_RAT
    GetERDesc = "UMC; Class Representative; Ratio (Light/Heavy)"
Case glER_UMC_REP_LOG
    GetERDesc = "UMC; Class Representative; Logarithmic"
Case glER_UMC_REP_ALT
    GetERDesc = "UMC; Class Representative; Symmetric Ratio"
Case glER_UMC_AVG_RAT
    GetERDesc = "UMC; Class Average; Ratio (Light/Heavy)"
Case glER_UMC_AVG_LOG
    GetERDesc = "UMC; Class Average; Logarithmic"
Case glER_UMC_AVG_ALT
    GetERDesc = "UMC; Class Average; Symmetric Ratio"
Case glER_UMC_SUM_RAT
    GetERDesc = "UMC; Class Sum; Ratio (Light/Heavy)"
Case glER_UMC_SUM_LOG
    GetERDesc = "UMC; Class Sum; Logarithmic"
Case glER_UMC_SUM_ALT
    GetERDesc = "UMC; Class Sum; Symmetric Ratio"
Case Else
    GetERDesc = "Unknown"
End Select
End Function

' Unused Function (March 2003)
'''Public Sub DltLblPairsReport(ByVal Ind As Long)
''''--------------------------------------------------
''''generates and displays generic DltLbl pairs report
''''--------------------------------------------------
'''Dim FileNum As Integer
'''Dim FileNam As String
'''Dim sLine As String
'''Dim i As Long
'''Dim strSepChar as string
'''On Error Resume Next
'''strSepChar = LookupDefaultSeparationCharacter()
'''If GelP_D_L(Ind).PCnt > 0 Then
'''   With GelP_D_L(Ind)
'''     FileNum = FreeFile
'''     FileNam = GetTempFolder() & RawDataTmpFile
'''     Open FileNam For Output As FileNum
'''     'print gel file name and Search definition as reference
'''     Print #FileNum, "Gel File: " & GelBody(Ind).Caption
'''
'''     Print #FileNum, "Label type: "
'''     Print #FileNum, "Delta type: "
'''     Print #FileNum, "Delta Mass= " & .SearchDef.DeltaMass
'''     Print #FileNum, "Label Mass= " & .SearchDef.LightLabelMass
'''     Print #FileNum, "ER Type: " & GetERDesc(.SearchDef.ERCalcType)
'''
'''     sLine = "Lt.ID" & strSepChar & "Lt.MW" & strSepChar & "Lt.Scan" & strSepChar _
'''           & "Lt.Int" & strSepChar & "Lt.Lbl.Cnt" & strSepChar & "Hv.ID" & strSepChar _
'''           & "Hv.MW" & strSepChar & "Hv.Scan" & strSepChar & "Hv.Int" & strSepChar _
'''           & "Hv.Lbl.Cnt" & strSepChar & "Hv.Dlt.Cnt"
'''
'''     Print #FileNum, sLine
'''     For i = 0 To .PCnt - 1
'''       sLine = .Pairs(i).P1 & strSepChar & GelData(Ind).IsoData(.Pairs(i).P1).MonoisotopicMW & strSepChar _
'''               & GelData(Ind).IsoData(.Pairs(i).P1, 1) & strSepChar _
'''               & GelData(Ind).IsoData(.Pairs(i).P1, 3) & strSepChar & .P1LblCnt(i) & strSepChar _
'''               & .Pairs(i).P2 & strSepChar & GelData(Ind).IsoData(.Pairs(i).P2).MonoisotopicMW & strSepChar _
'''               & GelData(Ind).IsoData(.Pairs(i).P2, 1) & strSepChar _
'''               & GelData(Ind).IsoData(.Pairs(i).P2, 3) & strSepChar _
'''               & .P2LblCnt(i) & strSepChar & .P2DltCnt(i)
'''       Print #FileNum, sLine
'''     Next i
'''     Close FileNum
'''   End With
'''   frmDataInfo.Tag = "Dlt_Lbl"
'''   frmDataInfo.Show vbModal
'''Else
'''   MsgBox "No pairs found.", vbOKOnly
'''End If
'''End Sub

Public Function PairsLookupFPRType(lngGelIndex As Long, blnHeavyPairMember As Boolean) As Long
    
    Dim lngPeakFPRType As Long
    
    Select Case GelAnalysis(lngGelIndex).MD_Type
    Case stPairsICAT
        If blnHeavyPairMember Then
            lngPeakFPRType = FPR_Type_ICAT_H
        Else
            lngPeakFPRType = FPR_Type_ICAT_L
        End If
    Case stPairsLysC12C13
        If blnHeavyPairMember Then
            lngPeakFPRType = FPR_Type_C12_C13_H
        Else
            lngPeakFPRType = FPR_Type_C12_C13_L
        End If
    Case stPairsPEO
        If blnHeavyPairMember Then
            lngPeakFPRType = FPR_Type_PEO_H
        Else
            lngPeakFPRType = FPR_Type_PEO_L
        End If
    Case stPairsPhIAT
        If blnHeavyPairMember Then
            lngPeakFPRType = FPR_Type_PhIAT_H
        Else
            lngPeakFPRType = FPR_Type_PhIAT_L
        End If
    Case stPairsPEON14N15
        If blnHeavyPairMember Then
            lngPeakFPRType = FPR_Type_PEO_N14_N15_H
        Else
            lngPeakFPRType = FPR_Type_PEO_N14_N15_L
        End If
    Case stPairsO16O18
        If blnHeavyPairMember Then
            lngPeakFPRType = FPR_Type_O16_O18_H
        Else
            lngPeakFPRType = FPR_Type_O16_O18_L
        End If

    Case Else
        ' Includes stPairsN14N15
        If blnHeavyPairMember Then
            lngPeakFPRType = FPR_Type_N14_N15_H
        Else
            lngPeakFPRType = FPR_Type_N14_N15_L
        End If
    End Select
    
    PairsLookupFPRType = lngPeakFPRType
    
End Function

Public Function PairsResetExclusionFlag(ByVal Ind As Long) As String
'---------------------------------------------------
'Resets the exclusion flag for all pairs to glPAIR_Neutral
'---------------------------------------------------
Dim i As Long
Dim OKCnt As Long
Dim DeletedCnt As Long
Dim strMessage As String

On Error Resume Next
With GelP_D_L(Ind)
  If .PCnt > 0 Then
    For i = 0 To .PCnt - 1
        .Pairs(i).STATE = glPAIR_Neu
    Next i
    
    strMessage = "Reset all pairs to neutral inclusion state; total pair count = " & .PCnt
  Else
    strMessage = "No pairs were found in memory"
  End If
End With
PairsResetExclusionFlag = strMessage
End Function

Public Function PairsSearchMarkAmbiguous(frmCallingForm As VB.Form, lngGelIndex As Long, blnUMCBasedPairs As Boolean, Optional lngTemporarilyExcludedPairCount As Long = 0) As String
' Returns a status message
' blnUMCBasedPairs determines whether ResolveDltLblPairsUMC or ResolveDltLblPairs_S is called

Dim OKCnt As Long
Dim ExcludedCount As Long
Dim lngPairIndex As Long

Dim strPairType As String
Dim strMessage As String

On Error GoTo PairsSearchMarkAmbiguousErrorHandler

If blnUMCBasedPairs Then
    strPairType = "UMC-based"
Else
    strPairType = "peak-based"
End If

If GelP_D_L(lngGelIndex).PCnt > 0 Then
    frmCallingForm.MousePointer = vbHourglass
   
    If blnUMCBasedPairs Then
        OKCnt = ResolveDltLblPairsUMC(lngGelIndex)
    Else
        OKCnt = ResolveDltLblPairs_S(lngGelIndex)
    End If
   
    With GelP_D_L(lngGelIndex)
        OKCnt = 0
        ExcludedCount = 0
        For lngPairIndex = 0 To .PCnt - 1
            If .Pairs(lngPairIndex).STATE = glPAIR_Exc Then
                ExcludedCount = ExcludedCount + 1
            Else
                OKCnt = OKCnt + 1
            End If
        Next lngPairIndex
        ExcludedCount = ExcludedCount - lngTemporarilyExcludedPairCount
        OKCnt = OKCnt + lngTemporarilyExcludedPairCount
        If ExcludedCount < 0 Then ExcludedCount = 0
    End With
   
    frmCallingForm.MousePointer = vbDefault
    If OKCnt >= 0 Then
        With GelP_D_L(lngGelIndex)
            strMessage = "Total number of pairs: " & .PCnt & "; " & "Number of unambiguous pairs: " & OKCnt
            
            AddToAnalysisHistory lngGelIndex, "Identified ambiguous " & strPairType & " pairs; Unambiguous Count = " & Trim(OKCnt) & "; Ambiguous Count = " & Trim(.PCnt - OKCnt) & "; Total Pairs Count = " & Trim(.PCnt)
        End With
    Else
        strMessage = "Error resolving ambiguous pairs."
        If blnUMCBasedPairs Then
            strMessage = strMessage & " Make sure pairs are synchronized with latest UMC."
        End If
    End If
Else
   strMessage = "No pairs found."
End If

PairsSearchMarkAmbiguous = strMessage
Exit Function

PairsSearchMarkAmbiguousErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in PairsSearchMarkAmbiguous: " & Err.Description, vbOKOnly, glFGTU
    Else
        LogErrors Err.Number, "PairsSearchMarkAmbiguous", Err.Description, lngGelIndex
    End If

End Function

Public Function PairsSearchMarkAmbiguousPairsWithHitsOnly(frmCallingForm As VB.Form, lngGelIndex As Long) As String
    
    Dim strMessage As String
    
    Dim lngIndex As Long
    
    Dim blnPairHasUMCWithHit() As Boolean       ' 0-based; Parallel to GelP_D_L().Pairs(); true if the pair has a UMC with one or more hits
    
    Dim lngPairsWithoutHitsCount As Long
    Dim lngPairsWithoutHits() As Long           ' 0-based; List of indices in GelP_D_L().Pairs that do not have any hits
    Dim intPairsWithoutHitsStateSaved() As Integer
    
On Error GoTo PairsSearchMarkAmbiguousPairsWithHitsOnlyErrorHandler

    With GelP_D_L(lngGelIndex)
        If .PCnt > 0 Then
        
            ReDim blnPairHasUMCWithHit(.PCnt - 1)
            ReDim lngPairsWithoutHits(.PCnt - 1)
            ReDim intPairsWithoutHitsStateSaved(.PCnt - 1)
            
            ' Determine which pairs have hits
            For lngIndex = 0 To .PCnt - 1
                With .Pairs(lngIndex)
                    If IsAMTReferencedByUMC(GelUMC(lngGelIndex).UMCs(.P1), lngGelIndex) Then
                        blnPairHasUMCWithHit(lngIndex) = True
                    Else
                        If IsAMTReferencedByUMC(GelUMC(lngGelIndex).UMCs(.P2), lngGelIndex) Then
                            blnPairHasUMCWithHit(lngIndex) = True
                        End If
                    End If
                End With
            Next lngIndex
            
            ' Step through the pairs again, and exclude those that are currently included but don't have any hits
            ' Keep track of the excluded pairs using lngPairsWithoutHits()
            lngPairsWithoutHitsCount = 0
            For lngIndex = 0 To .PCnt - 1
                With .Pairs(lngIndex)
                    If Not blnPairHasUMCWithHit(lngIndex) And .STATE <> glPAIR_Exc Then
                        lngPairsWithoutHits(lngPairsWithoutHitsCount) = lngIndex
                        intPairsWithoutHitsStateSaved(lngPairsWithoutHitsCount) = .STATE
                        lngPairsWithoutHitsCount = lngPairsWithoutHitsCount + 1
                        
                        .STATE = glPAIR_Exc
                    End If
                End With
            Next lngIndex
            
            ' Now exclude the ambiguous pairs
            strMessage = PairsSearchMarkAmbiguous(frmCallingForm, lngGelIndex, True, lngPairsWithoutHitsCount)
            
            ' Now restore the pairs that were excluded because they had no hits
            For lngIndex = 0 To lngPairsWithoutHitsCount - 1
                With .Pairs(lngPairsWithoutHits(lngIndex))
                    .STATE = intPairsWithoutHitsStateSaved(lngIndex)
                End With
            Next lngIndex
        
        Else
            strMessage = "No pairs found."
        End If
    End With

    PairsSearchMarkAmbiguousPairsWithHitsOnly = strMessage
    Exit Function

PairsSearchMarkAmbiguousPairsWithHitsOnlyErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.PairsSearchMarkAmbiguousPairsWithHitsOnly"
    PairsSearchMarkAmbiguousPairsWithHitsOnly = "Error in code (PairsSearchMarkAmbiguousPairsWithHitsOnly)"
    
End Function

Public Sub PairSearchIncludeNeutralPairs(lngGelIndex As Long)
    ' Changes the PState() value for those pairs with a state of glPAIR_Neu to glPAIR_Inc
    
    Dim i As Long

On Error GoTo PairSearchChangeIncludeNeutralPairsErrorHandler

    If GelP_D_L(lngGelIndex).PCnt > 0 Then
        With GelP_D_L(lngGelIndex)
            For i = 0 To .PCnt - 1
                If .Pairs(i).STATE = glPAIR_Neu Then
                   .Pairs(i).STATE = glPAIR_Inc
                End If
            Next i
        End With
    End If

    Exit Sub

PairSearchChangeIncludeNeutralPairsErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in PairSearchChangeIncludeNeutralPairs: " & Err.Description, vbOKOnly, glFGTU
    Else
        LogErrors Err.Number, "PairSearchChangeIncludeNeutralPairs"
    End If
    
End Sub

Public Function PairsSearchMarkBadER(ERMin As Double, ERMax As Double, lngGelIndex As Long, blnUMCBasedPairs As Boolean) As String
' Returns a status message

Dim i As Long
Dim lngExclusionCount As Long
Dim strPairType As String
Dim strMessage As String

On Error GoTo PairsSearchMarkBadERErrorHandler

If blnUMCBasedPairs Then
    strPairType = "UMC-based"
Else
    strPairType = "peak-based"
End If

If GelP_D_L(lngGelIndex).PCnt > 0 Then
    With GelP_D_L(lngGelIndex)
        .SearchDef.ERInclusionMin = ERMin
        .SearchDef.ERInclusionMax = ERMax
        
        lngExclusionCount = 0
        For i = 0 To .PCnt - 1
            If (.Pairs(i).ER < ERMin) Or (.Pairs(i).ER > ERMax) Then
               .Pairs(i).STATE = glPAIR_Exc
               lngExclusionCount = lngExclusionCount + 1
            Else
               .Pairs(i).STATE = glPAIR_Inc
            End If
        Next i
    
        strMessage = "Excluded " & Trim(lngExclusionCount) & " pairs (" & Trim(.PCnt) & " pairs total)"
        
        AddToAnalysisHistory lngGelIndex, "Excluded " & strPairType & " pairs out of the expression ratio range; Pairs Excluded = " & Trim(lngExclusionCount) & "; Total Pairs Count = " & Trim(.PCnt) & "; Minimum ER = " & Trim(ERMin) & "; Maximum ER = " & Trim(ERMax)
    End With
Else
    strMessage = "No pairs found. Make sure that Find Pairs function was applied first."
End If

PairsSearchMarkBadER = strMessage
Exit Function

PairsSearchMarkBadERErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in PairsSearchMarkBadER: " & Err.Description, vbOKOnly, glFGTU
    Else
        LogErrors Err.Number, "PairsSearchMarkBadER", Err.Description, lngGelIndex
    End If

End Function

Private Function RatER(ByVal LtAbu As Double, ByVal HvAbu As Double) As Double
'------------------------------------------------------
'returns ER as standard ratio of abundance of light and
'heavy pair members
'------------------------------------------------------
On Error Resume Next
RatER = LtAbu / HvAbu
If Err Then RatER = ER_CALC_ERR
End Function

Private Function RatERViaAltER(ByVal dblAltER As Double) As Double
'------------------------------------------------------
'Converts shifted symmetric abundance ratio values to normal ratio-based ER values
'------------------------------------------------------
On Error GoTo err_AltER

If dblAltER >= 0 Then
    RatERViaAltER = dblAltER + 1
Else
    RatERViaAltER = 1 / (1 - dblAltER)
End If
Exit Function

err_AltER:
RatERViaAltER = ER_CALC_ERR
End Function

Private Function LogER(ByVal LtAbu As Double, ByVal HvAbu As Double) As Double
'------------------------------------------------------
'returns ER as logarithm of abundance ratio of light
'and heavy pair members
'------------------------------------------------------
On Error Resume Next
LogER = Log(LtAbu / HvAbu)
If Err Then LogER = ER_CALC_ERR
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

Private Function AltER(ByVal LtAbu As Double, ByVal HvAbu As Double) As Double
'------------------------------------------------------
'returns ER as shifted symmetric abundance ratio of light
'and heavy pair members
'------------------------------------------------------
Dim tmpER As Double
On Error GoTo err_AltER
tmpER = LtAbu / HvAbu
If tmpER >= 1 Then
   AltER = tmpER - 1
Else
   AltER = 1 - (1 / tmpER)
End If
Exit Function

err_AltER:
AltER = ER_CALC_ERR
End Function

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

Public Sub ReportDltLblPairs_S(ByVal Ind As Long, PState As Integer, Optional strFilePath As String = "")
'-------------------------------------------------------------------
'report all pairs results (for individual peaks) in temporary file
'PState determines which pairs will be reported; exc., inc. or all
'-------------------------------------------------------------------
Dim fname As String
Dim sLine As String
Dim LPart As String, HPart As String
Dim i As Long
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim blnSaveToDiskOnly As Boolean
Dim blnUserNotifiedOfError As Boolean
Dim strSepChar As String

On Error GoTo ReportDltLblPairsErrorHandler

If Len(strFilePath) > 0 Then
   fname = strFilePath
   blnSaveToDiskOnly = True
Else
   fname = GetTempFolder() & RawDataTmpFile
End If

strSepChar = LookupDefaultSeparationCharacter()

With GelP_D_L(Ind)
   If .PCnt > 0 Then
      Set ts = fso.OpenTextFile(fname, ForWriting, True)
      ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
      'print gel file name and pairs definitions as reference
      ts.WriteLine "Gel File: " & GelBody(Ind).Caption
      ts.WriteLine "Reporting delta-label pairs for individual peaks(isotopic only)."
      Select Case PState
      Case glPAIR_Inc
           ts.WriteLine "Included pairs."
      Case glPAIR_Exc
           ts.WriteLine "Excluded pairs."
      Case Else
           ts.WriteLine "All pairs."
      End Select
      ts.WriteLine "Label mass: " & .SearchDef.LightLabelMass
      ts.WriteLine "Delta mass: " & .SearchDef.DeltaMass
      Select Case .SearchDef.ERCalcType
      Case ectER_RAT
        ts.WriteLine "ER calculation: Ratio; AbuLight/AbuHeavy"
      Case ectER_LOG
        ts.WriteLine "ER calculation: Logarithmic Ratio; Log(AbuLight/AbuHeavy)"
      Case ectER_ALT
        ts.WriteLine "ER calculation: 0-Shifted Symmetric Ratio; (AbuL/AbuH)-1 for AbuL>=AbuH; 1-(AbuH/AbuL) for AbuL<AbuH"
      Case Else
        ts.WriteLine "ER calculation: Unknown"
      End Select
      ts.WriteLine
      'header line
      sLine = "Light Index" & strSepChar & "Light MW" & strSepChar & "Light Abu" & strSepChar _
            & "Light Scan" & strSepChar & "Light Lbl Count" & strSepChar _
            & "Hight Index" & strSepChar & "Hight MW" & strSepChar & "High Abu" & strSepChar _
            & "Hight Scan" & strSepChar & "High Lbl Count" & strSepChar & "Delta Count" & strSepChar & "ER"
            
      ts.WriteLine sLine
      For i = 0 To .PCnt - 1
        With .Pairs(i)
            If .STATE = PState Or Abs(PState) <> 1 Then
              LPart = .P1 & strSepChar & Round(GetIsoMass(GelData(Ind).IsoData(.P1), GelData(Ind).Preferences.IsoDataField), 6) & strSepChar _
                       & GelData(Ind).IsoData(.P1).Abundance & strSepChar _
                       & GelData(Ind).IsoData(.P1).ScanNumber & strSepChar & .P1LblCnt
              HPart = .P2 & strSepChar & Round(GetIsoMass(GelData(Ind).IsoData(.P2), GelData(Ind).Preferences.IsoDataField), 6) & strSepChar _
                       & GelData(Ind).IsoData(.P2).Abundance & strSepChar _
                       & GelData(Ind).IsoData(.P2).ScanNumber & strSepChar & .P2LblCnt & strSepChar & .P2DltCnt
              sLine = LPart & strSepChar & HPart & strSepChar & .ER
              ts.WriteLine sLine
            End If
        End With
      Next i
      ts.Close
      DoEvents
      If Not blnSaveToDiskOnly Then
         frmDataInfo.Tag = "DLT_LBL"
         frmDataInfo.Show vbModal
      End If
   Else
      If Not blnSaveToDiskOnly And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "No pairs found.", vbOKOnly, glFGTU
      End If
   End If
End With

Exit Sub

ReportDltLblPairsErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If Not blnUserNotifiedOfError Then
            MsgBox "Error in ReportDltLblPairs_S: " & Err.Description, vbOKOnly, glFGTU
            blnUserNotifiedOfError = True
        End If
    Else
        LogErrors Err.Number, "Pairs.bas->ReportDltLblPairs_S"
        AddToAnalysisHistory Ind, "Error in ReportDltLblPairs_S: " & Err.Description
    End If
    Resume Next

End Sub

Public Sub ReportDltLblPairsUMCWrapper(ByVal Ind As Long, PState As Integer, Optional strFilePath As String = "")
    ' This sub will create the ClsStat() array for the loaded UMC's,
    '  then call ReportDltLblPairs_UMC
    ' Call this sub to report or save the pairs without having to load frmUMCDltPairs
    
    Dim ClsStat() As Double                 ' Holds Stats on each UMC, including min and max scan number
    Dim lngAllUMCCount As Long
    
    lngAllUMCCount = UMCStatistics1(Ind, ClsStat())
    Debug.Assert lngAllUMCCount = GelUMC(Ind).UMCCnt

    ReportDltLblPairs_UMC Ind, ClsStat(), PState, strFilePath
    
End Sub

Private Sub ReportDltLblPairs_UMC(ByVal Ind As Long, _
                                 ClsStat() As Double, _
                                 PState As Integer, _
                                 Optional strFilePath As String = "")
'------------------------------------------------------
'report all pairs results (for unique mass classes) in
'temporary file; PState determines which pairs will be
'reported; excluded, included or all
'If Len(strFilePath) = 0, then displays report using frmDataInfo;
'  otherwise, saves the report to strFilePath
'PState can be 0 (aka glPAIR_Inc) for all pairs, 1 for Included only (aka glPAIR_Inc),
'  or -1 for Excluded only (aka glPAIR_Exc)
'------------------------------------------------------
Dim fname As String
Dim sLine As String
Dim i As Long
Dim intChargeStateBasis As Integer
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim blnSaveToDiskOnly As Boolean
Dim blnUserNotifiedOfError As Boolean
Dim strSepChar As String

On Error GoTo ReportDltLblPairsUMCErrorHandler

If Len(strFilePath) > 0 Then
   fname = strFilePath
   blnSaveToDiskOnly = True
Else
   fname = GetTempFolder() & RawDataTmpFile
End If

strSepChar = LookupDefaultSeparationCharacter()

With GelP_D_L(Ind)
   If .PCnt >= 0 Then
      Set ts = fso.OpenTextFile(fname, ForWriting, True)
      ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
      'print gel file name and pairs definitions as reference
      ts.WriteLine "Gel File: " & GelBody(Ind).Caption
      ts.WriteLine "Reporting delta-label pairs for unique mass classes."
      Select Case PState
      Case glPAIR_Inc
           ts.WriteLine "Included pairs."
      Case glPAIR_Exc
           ts.WriteLine "Excluded pairs."
      Case Else
           ts.WriteLine "All pairs."
      End Select
      ts.WriteLine "Label mass: " & .SearchDef.LightLabelMass
      ts.WriteLine "Delta mass: " & .SearchDef.DeltaMass
      Select Case .SearchDef.ERCalcType
      Case ectER_RAT
        ts.WriteLine "ER calculation: Ratio; AbuLight/AbuHeavy"
      Case ectER_LOG
        ts.WriteLine "ER calculation: Logarithmic Ratio; Log(AbuLight/AbuHeavy)"
      Case ectER_ALT
        ts.WriteLine "ER calculation: 0-Shifted Symmetric Ratio; (AbuL/AbuH)-1 for AbuL>=AbuH; 1-(AbuH/AbuL) for AbuL<AbuH"
      Case Else
        ts.WriteLine "ER calculation: Unknown"
      End Select
      ts.WriteLine
      ts.WriteLine "Unique Mass Class definition"
      ts.Write GetUMCDefDesc(GelUMC(Ind).def)
      ts.WriteLine
      ts.WriteLine
      
      'header line
      sLine = ""
      sLine = sLine & "Pair Index" & strSepChar & "UMC Light Index" & strSepChar & "Light MW" & strSepChar & "Light Abu" & strSepChar
      sLine = sLine & "Light Lbl Count" & strSepChar & "Light ScanStart" & strSepChar & "Light ScanEnd" & strSepChar & "Light Charge State Basis" & strSepChar & "Light Charge Basis MZ" & strSepChar
      sLine = sLine & "UMC Heavy Index" & strSepChar & "Heavy MW" & strSepChar & "Heavy Abu" & strSepChar
      sLine = sLine & "Heavy Lbl Count" & strSepChar & "Delta Count" & strSepChar
      sLine = sLine & "Heavy ScanStart" & strSepChar & "Heavy ScanEnd" & strSepChar & "Heavy Charge State Basis" & strSepChar & "Heavy Charge Basis MZ" & strSepChar
      sLine = sLine & "ER" & strSepChar & "ER StDev" & strSepChar & "ER Charge State Basis Count" & strSepChar & "ER Member Basis Count"
      If PState = glPair_All Then sLine = sLine & strSepChar & "State"
      
      ts.WriteLine sLine
      For i = 0 To .PCnt - 1
         With .Pairs(i)
            If .STATE = PState Or Abs(PState) <> 1 Then
                sLine = ""
                
                ' Old method, referencing ClsStat()
                'sLine = sLine & i & strSepChar & .P1 & strSepChar & Round(ClsStat(.P1, ustClassMW), 6) & strSepChar & ClsStat(.P1, ustClassIntensity) & strSepChar
                'sLine = sLine & .P1LblCnt & strSepChar & ClsStat(.P1, ustScanStart) & strSepChar & ClsStat(.P1, ustScanEnd) & strSepChar
                
                ' First the light member
                ' New method, grabbing values directly from GelUMC().UMCs()
                sLine = sLine & i & strSepChar & .P1 & strSepChar & Round(GelUMC(Ind).UMCs(.P1).ClassMW, 6) & strSepChar & GelUMC(Ind).UMCs(.P1).ClassAbundance & strSepChar
                sLine = sLine & .P1LblCnt & strSepChar & GelUMC(Ind).UMCs(.P1).MinScan & strSepChar & GelUMC(Ind).UMCs(.P1).MaxScan & strSepChar

                ' Record ChargeBasis and UMCMZForChargeBasis
                If GelUMC(Ind).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                    intChargeStateBasis = GelUMC(Ind).UMCs(.P1).ChargeStateBasedStats(GelUMC(Ind).UMCs(.P1).ChargeStateStatsRepInd).Charge
                    sLine = sLine & Trim(intChargeStateBasis) & strSepChar
                Else
                    intChargeStateBasis = CInt(GelData(Ind).IsoData(GelUMC(Ind).UMCs(.P1).ClassRepInd).Charge)
                    sLine = sLine & 0 & strSepChar
                End If
                sLine = sLine & Round(MonoMassToMZ(GelUMC(Ind).UMCs(.P1).ClassMW, intChargeStateBasis), 6) & strSepChar
                    
                ' Now the heavy menber
                sLine = sLine & .P2 & strSepChar & Round(GelUMC(Ind).UMCs(.P2).ClassMW, 6) & strSepChar & GelUMC(Ind).UMCs(.P2).ClassAbundance & strSepChar
                sLine = sLine & .P2LblCnt & strSepChar & .P2DltCnt & strSepChar
                sLine = sLine & GelUMC(Ind).UMCs(.P2).MinScan & strSepChar & GelUMC(Ind).UMCs(.P2).MaxScan & strSepChar
                
                ' Record ChargeBasis and UMCMZForChargeBasis
                If GelUMC(Ind).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                    intChargeStateBasis = GelUMC(Ind).UMCs(.P2).ChargeStateBasedStats(GelUMC(Ind).UMCs(.P2).ChargeStateStatsRepInd).Charge
                    sLine = sLine & Trim(intChargeStateBasis) & strSepChar
                Else
                    intChargeStateBasis = CInt(GelData(Ind).IsoData(GelUMC(Ind).UMCs(.P2).ClassRepInd).Charge)
                    sLine = sLine & 0 & strSepChar
                End If
                sLine = sLine & Round(MonoMassToMZ(GelUMC(Ind).UMCs(.P2).ClassMW, intChargeStateBasis), 6) & strSepChar
                
                sLine = sLine & Round(.ER, 6) & strSepChar & Round(.ERStDev, 6) & strSepChar & .ERChargeStateBasisCount & strSepChar & .ERMemberBasisCount
                    
                If PState = glPair_All Then
                    Select Case .STATE
                    Case glPAIR_Inc
                        sLine = sLine & strSepChar & "Included"
                    Case glPAIR_Exc
                        sLine = sLine & strSepChar & "Excluded"
                    Case Else
                        ' Neutral; we'll call it Included
                        ' If this code gets encountered, it means we may want to call
                        '  PairSearchIncludeNeutralPairs() prior to calling this function
                        sLine = sLine & strSepChar & "Included"
                    End Select
                End If
                ts.WriteLine sLine
            End If
         End With
      Next i
      ts.Close
      DoEvents
      If Not blnSaveToDiskOnly Then
         frmDataInfo.Tag = "DLT_LBL"
         frmDataInfo.Show vbModal
      End If
   Else
      If Not blnSaveToDiskOnly And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "No pairs found.", vbOKOnly, glFGTU
      End If
   End If
End With

Exit Sub

ReportDltLblPairsUMCErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If Not blnUserNotifiedOfError Then
            MsgBox "Error in ReportDltLblPairs_UMC: " & Err.Description, vbOKOnly, glFGTU
            blnUserNotifiedOfError = True
        End If
    Else
        LogErrors Err.Number, "Pairs.bas->ReportDltLblPairs_UMC"
        AddToAnalysisHistory Ind, "Error in ReportDltLblPairs_UMC: " & Err.Description
    End If
    Resume Next

End Sub

Public Sub ReportERStat(ByVal Ind As Long, _
                        ERBin() As Double, _
                        ERAll() As Long, _
                        ERInc() As Long, _
                        ERExc() As Long, _
                        AllS As ERStatHelper, _
                        IncS As ERStatHelper, _
                        ExcS As ERStatHelper, _
                        Optional strFilePath As String = "")
'------------------------------------------------------
'reports ER statistics
'If Len(strFilePath) = 0, then displays report using frmDataInfo;
'  otherwise, saves the report to strFilePath
'------------------------------------------------------
Dim fname As String
Dim sLine As String
Dim i As Long
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim blnSaveToDiskOnly As Boolean
Dim blnUserNotifiedOfError As Boolean
Dim strSepChar As String

On Error GoTo ReportERStatErrorHandler

If Len(strFilePath) > 0 Then
    fname = strFilePath
    blnSaveToDiskOnly = True
Else
    fname = GetTempFolder() & RawDataTmpFile
End If

strSepChar = LookupDefaultSeparationCharacter()

Set ts = fso.OpenTextFile(fname, ForWriting, True)
ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
ts.WriteLine "Gel File: " & GelBody(Ind).Caption
ts.WriteLine "Reporting ER statistics."
Select Case GelP_D_L(Ind).SearchDef.ERCalcType
Case ectER_RAT
  ts.WriteLine "ER calculation: Ratio; AbuL/AbuH"
Case ectER_LOG
  ts.WriteLine "ER calculation: Logarithmic Ratio; Log(AbuL/AbuH)"
Case ectER_ALT
  ts.WriteLine "ER calculation: 0-Shifted Symmetric Ratio; (AbuL/AbuH)-1 for AbuL>=AbuH; 1-(AbuH/AbuL) for AbuL<AbuH"
Case Else
  ts.WriteLine "ER calculation: Unknown"
End Select
ts.WriteLine
ts.WriteLine "Overall Statistics"
ts.WriteLine strSepChar & "All Pairs" & strSepChar & "Included Pairs" _
            & strSepChar & "Excluded Pairs"
ts.WriteLine "Total count" & strSepChar & "" & AllS.ERCnt & strSepChar & IncS.ERCnt & strSepChar & ExcS.ERCnt
ts.WriteLine "Out of left ER range" & strSepChar & "" & AllS.ERBadL & strSepChar & IncS.ERBadL & strSepChar & ExcS.ERBadL
ts.WriteLine "Out of right ER range" & strSepChar & "" & AllS.ERBadR & strSepChar & IncS.ERBadR & strSepChar & ExcS.ERBadR
ts.WriteLine
sLine = "ER Bin" & strSepChar & "All Pairs" & strSepChar _
      & "Included Pairs" & strSepChar & "Excluded Pairs"
ts.WriteLine sLine
For i = 0 To 1000
    sLine = ERBin(i) & strSepChar & ERAll(i) & strSepChar _
            & ERInc(i) & strSepChar & ERExc(i)
    ts.WriteLine sLine
Next i
ts.Close
DoEvents
If Not blnSaveToDiskOnly Then
    frmDataInfo.Tag = "ER_STAT"
    frmDataInfo.Show vbModal
End If

Exit Sub

ReportERStatErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If Not blnUserNotifiedOfError Then
            MsgBox "Error in ReportERStat: " & Err.Description, vbOKOnly, glFGTU
            blnUserNotifiedOfError = True
        End If
    Else
        LogErrors Err.Number, "Pairs.bas->ReportERStat"
        AddToAnalysisHistory Ind, "Error in ReportERStat: " & Err.Description
    End If
    Resume Next
End Sub

Private Function ResolveDltLblPairsUMC(ByVal Ind As Long) As Long
'----------------------------------------------------------------
'resolves ambiguous pairs and returns number of OK pairs; mark
'ambiguous pairs glPAIR_Exc, OK pairs glPAIR_Inc
'pairs marked glPAIR_Exc are not considered in this procedure
'NOTE: last statement makes order of steps in resolving pairs
'important and this procedure should always be applied last

'If blnKeepMostConfidentAmbiguous = False, then
    'only pairs made from classes that participate in only one
    'pair are included after this procedure. This is very restrictive
    'rule - the idea is that at this point we already have excluded
    'majority of random pairs(with ER range and database search)
'Otherwise, when blnKeepMostConfidentAmbiguous = True, then
    'removes the pairs that contains the least believable ER values UMC is shared between several pairs

'NOTE: number of OK pairs returned is the number of pairs that
'this procedure declared to be unambiguous
'----------------------------------------------------------------
Dim ClsLCnt() As Long      'parallel with UMC classes; shows how
Dim ClsHCnt() As Long      'many times each class was used as L/H
Dim ClsCnt As Long
Dim IncPairsCnt As Long

Dim MatchingPairsCount As Long
Dim MatchingPairs() As Long         ' Indices of members in GelP_D_L().Pairs
Dim dblMassDiff As Double
Dim dblMassDiffMin As Double
Dim dblMassDiffMax As Double

Dim i As Long, j As Long
Dim blnKeepMostConfidentAmbiguous As Boolean

blnKeepMostConfidentAmbiguous = glbPreferencesExpanded.PairSearchOptions.KeepMostConfidentAmbiguous

ClsCnt = GelUMC(Ind).UMCCnt

    If GelP_D_L(Ind).SyncWithUMC And GelP_D_L(Ind).PCnt > 0 And ClsCnt > 0 Then
        'count how many times each class appears as light and
        'heavy member of a pair(ignore already excluded pairs)
        ReDim ClsLCnt(ClsCnt - 1)
        ReDim ClsHCnt(ClsCnt - 1)
       
        With GelP_D_L(Ind)
            For i = 0 To .PCnt - 1
                If Not .Pairs(i).STATE = glPAIR_Exc Then
                   ClsLCnt(.Pairs(i).P1) = ClsLCnt(.Pairs(i).P1) + 1
                   ClsHCnt(.Pairs(i).P2) = ClsHCnt(.Pairs(i).P2) + 1
                End If
            Next i
        End With
        
        For i = 0 To GelP_D_L(Ind).PCnt - 1
           If Not GelP_D_L(Ind).Pairs(i).STATE = glPAIR_Exc Then
              If ClsLCnt(GelP_D_L(Ind).Pairs(i).P1) = 1 And ClsHCnt(GelP_D_L(Ind).Pairs(i).P1) = 0 _
                 And ClsLCnt(GelP_D_L(Ind).Pairs(i).P2) = 0 And ClsHCnt(GelP_D_L(Ind).Pairs(i).P2) = 1 Then
                 GelP_D_L(Ind).Pairs(i).STATE = glPAIR_Inc
                 IncPairsCnt = IncPairsCnt + 1
              Else
                ' Ambiguous pair; either the light or heavy UMC is shared among multiple pairs
                If blnKeepMostConfidentAmbiguous Then
                
                    ' ToDo: Check this; decide whether or not to check for ClsHCnt() = 0
                    ' Debug.Assert False
                    
                    If ClsLCnt(GelP_D_L(Ind).Pairs(i).P1) > 1 And ClsHCnt(GelP_D_L(Ind).Pairs(i).P1) = 0 Then
                    End If
                    
                    If ClsLCnt(GelP_D_L(Ind).Pairs(i).P1) > 1 Then
                        ' The light member of the pair is a member of several pairs, and it is the light member in all of them
                        ' Only keep one of the pairs that the light member is a member of
                        ' Choose the pair to keep based on the one with the largest sum of Light + Heavy
                        
                        ' Construct a list of the other pairs that include .Pairs(i).P1
                        With GelP_D_L(Ind)
                            MatchingPairsCount = 0
                            ReDim MatchingPairs(0)
                            For j = 0 To .PCnt - 1
                                If .Pairs(j).P1 = .Pairs(i).P1 Then
                                    ReDim Preserve MatchingPairs(MatchingPairsCount)
                                    MatchingPairs(MatchingPairsCount) = j
                                    MatchingPairsCount = MatchingPairsCount + 1
                                End If
                            Next j
                        End With
                        
                        Debug.Assert MatchingPairsCount > 1
                        If MatchingPairsCount > 0 Then
                            ResolveDltLblPairsUMCFindBest Ind, MatchingPairs(), MatchingPairsCount
                        Else
                            ' This shouldn't happen
                            Debug.Assert False
                             GelP_D_L(Ind).Pairs(i).STATE = glPAIR_Inc
                            IncPairsCnt = IncPairsCnt + 1
                        End If
                        
                    'ElseIf ClsLCnt(GelP_D_L(Ind).Pairs(i).P2) = 0 And ClsHCnt(GelP_D_L(Ind).Pairs(i).P1) > 1 Then
                    ElseIf ClsHCnt(GelP_D_L(Ind).Pairs(i).P1) > 1 Then
                        ' The heavy member of the pair is a member of several pairs, and it is the heavy member in all of them
                        ' Choose the pair to keep based on the one with the largest sum of Light + Heavy
                        ' However, since we perform database searching based on the light member's mass, we'll only examine those pairs that have the same
                        '  spacing between heavy and light masses as this pair (+- tolerance)
                        
                        ' Construct a list of the other pairs that include .Pairs(i).P2, but only include them
                        '  if they have the same spacing between the light and heavy members
                        With GelP_D_L(Ind)
                        
                            dblMassDiff = Abs(GelUMC(Ind).UMCs(.Pairs(i).P2).ClassMW - GelUMC(Ind).UMCs(.Pairs(i).P1).ClassMW)
                            dblMassDiffMin = dblMassDiff - .SearchDef.DeltaMassTolerance * 2
                            dblMassDiffMax = dblMassDiff + .SearchDef.DeltaMassTolerance * 2
                            
                            MatchingPairsCount = 0
                            ReDim MatchingPairs(0)
                            For j = 0 To .PCnt - 1
                                If .Pairs(j).P2 = .Pairs(i).P2 Then
                                    ' Make sure the mass difference is allowable
                                    dblMassDiff = Abs(GelUMC(Ind).UMCs(.Pairs(j).P2).ClassMW - GelUMC(Ind).UMCs(.Pairs(j).P1).ClassMW)
                                    
                                    If dblMassDiff >= dblMassDiffMin And dblMassDiff <= dblMassDiffMax Then
                                        ReDim Preserve MatchingPairs(MatchingPairsCount)
                                        MatchingPairs(MatchingPairsCount) = j
                                        MatchingPairsCount = MatchingPairsCount + 1
                                    End If
                                End If
                            Next j
                        End With
                        
                        Debug.Assert MatchingPairsCount >= 1
                        If MatchingPairsCount > 0 Then
                            ResolveDltLblPairsUMCFindBest Ind, MatchingPairs(), MatchingPairsCount
                        Else
                            ' This shouldn't happen
                            Debug.Assert False
                             GelP_D_L(Ind).Pairs(i).STATE = glPAIR_Inc
                            IncPairsCnt = IncPairsCnt + 1
                        End If
                        
                    Else
                        ' Since we don't know what else to do, exclude the pair
                        GelP_D_L(Ind).Pairs(i).STATE = glPAIR_Exc
                    End If
                    
                Else
                    GelP_D_L(Ind).Pairs(i).STATE = glPAIR_Exc
                End If
              End If
           End If
       Next i
       ResolveDltLblPairsUMC = IncPairsCnt
    Else
       ResolveDltLblPairsUMC = -1
    End If

End Function

Private Sub ResolveDltLblPairsUMCFindBest(ByVal Ind As Long, MatchingPairs() As Long, MatchingPairsCount As Long)
    
    Dim dblHighestSum As Double
    Dim dblCompareSum As Double
    Dim IndexBestSum As Long
    Dim i As Long
    
    With GelP_D_L(Ind)
        dblHighestSum = GelUMC(Ind).UMCs(.Pairs(0).P1).ClassAbundance + GelUMC(Ind).UMCs(.Pairs(0).P2).ClassAbundance
        IndexBestSum = 0
            
        ' Determine the pair with the highest sum of Light + Heavy
        For i = 1 To MatchingPairsCount - 1
            dblCompareSum = GelUMC(Ind).UMCs(.Pairs(MatchingPairs(i)).P1).ClassAbundance + GelUMC(Ind).UMCs(.Pairs(MatchingPairs(i)).P2).ClassAbundance
            If dblCompareSum > dblHighestSum Then
                dblHighestSum = dblCompareSum
                IndexBestSum = i
            End If
        Next i
        
        ' Exclude all pairs in MatchingPairs() except the one at IndexBestSum
        For i = 0 To MatchingPairsCount - 1
            If i <> IndexBestSum Then
                .Pairs(MatchingPairs(i)).STATE = glPAIR_Exc
            End If
        Next i
    End With
    
End Sub

Private Function ResolveDltLblPairs_S(ByVal Ind As Long) As Long
'----------------------------------------------------------------
'resolves ambiguous pairs and returns number of OK pairs; mark
'ambiguous pairs glPAIR_Exc, OK pairs glPAIR_Inc
'pairs marked glPAIR_Exc are not considered in this procedure
'NOTE: last statement makes order of steps in resolving pairs
'important and this procedure should always be applied last
'NOTE: only pairs made from peaks that participate in only one
'pair are included after this procedure. This is very exclusive
'rule - the idea is that at this point we already have excluded
'majority of random pairs(with ER range and database search)
'NOTE: number of OK pairs returned is the number of pairs that
'this procedure declared to be unambiguous
'----------------------------------------------------------------
Dim IsoLCnt() As Long      'parallel with isotopic data-shows how
Dim IsoHCnt() As Long      'many times each peak was used as L/H
Dim IsoCnt As Long
Dim IncPairsCnt As Long
Dim i As Long
IsoCnt = GelData(Ind).IsoLines
With GelP_D_L(Ind)
    If .PCnt > 0 And IsoCnt > 0 Then
       'count how many times each peak appears as light and
       'heavy member of a pair(ignore already excluded pairs)
       ReDim IsoLCnt(IsoCnt)
       ReDim IsoHCnt(IsoCnt)
       For i = 0 To .PCnt - 1
           If Not .Pairs(i).STATE = glPAIR_Exc Then
              IsoLCnt(.Pairs(i).P1) = IsoLCnt(.Pairs(i).P1) + 1
              IsoHCnt(.Pairs(i).P2) = IsoHCnt(.Pairs(i).P2) + 1
           End If
       Next i
       'there could be more sophisticated ways to resolve
       'pairs to get higher number of resolved OK pairs
       For i = 0 To .PCnt - 1
           If Not .Pairs(i).STATE = glPAIR_Exc Then
              If IsoLCnt(.Pairs(i).P1) = 1 And IsoHCnt(.Pairs(i).P1) = 0 _
                 And IsoLCnt(.Pairs(i).P2) = 0 And IsoHCnt(.Pairs(i).P2) = 1 Then
                 .Pairs(i).STATE = glPAIR_Inc
                 IncPairsCnt = IncPairsCnt + 1
              Else
                 .Pairs(i).STATE = glPAIR_Exc
              End If
           End If
       Next i
       ResolveDltLblPairs_S = IncPairsCnt
    Else
       ResolveDltLblPairs_S = -1
    End If
End With
End Function

Public Function DeleteExcludedPairs(ByVal Ind As Long) As String
'---------------------------------------------------
'deletes pairs marked as glPAIR_Exc to save some mem
'returns a status message
'---------------------------------------------------
Dim i As Long
Dim OKCnt As Long
Dim DeletedCnt As Long
Dim strMessage As String

On Error Resume Next
With GelP_D_L(Ind)
  If .PCnt > 0 Then
    For i = 0 To .PCnt - 1
        If .Pairs(i).STATE <> glPAIR_Exc Then
           OKCnt = OKCnt + 1
           .Pairs(OKCnt - 1) = .Pairs(i)
        Else
            DeletedCnt = DeletedCnt + 1
        End If
    Next i
    .PCnt = OKCnt
    If .PCnt > 0 Then
       ReDim Preserve .Pairs(.PCnt - 1)
    Else
       Erase .Pairs
    End If
    
    strMessage = "Deleted excluded Pairs; number deleted = " & Trim(DeletedCnt) & "; new total pair count = " & .PCnt
    If DeletedCnt > 0 Then AddToAnalysisHistory Ind, strMessage
  Else
    strMessage = "No pairs were found in memory"
  End If
End With
DeleteExcludedPairs = strMessage
End Function


Public Function GetPairsTypeDesc(ByVal Ind As Long) As String
'------------------------------------------------------------
'returns description of type of pairs contained in GelP_D_L
'------------------------------------------------------------
On Error Resume Next
Select Case GelP_D_L(Ind).DltLblType
Case ptUMCDlt
  GetPairsTypeDesc = "UMC - Delta Pairs"
Case ptUMCLbl
  GetPairsTypeDesc = "UMC - Label Pairs"
Case ptUMCDltLbl
  GetPairsTypeDesc = "UMC - Delta-Label Pairs"
Case ptS_Dlt
  GetPairsTypeDesc = "Individual - Delta Pairs"
Case ptS_Lbl
  GetPairsTypeDesc = "Individual - Label Pairs"
Case ptS_DltLbl
  GetPairsTypeDesc = "Individual - Delta-Label Pairs"
End Select
End Function

Public Sub FillUMC_ERs(ByVal Ind As Long)
'-----------------------------------------------------
'fills expression numbers from GelP_D_L to data arrays
'UMC pairs
'-----------------------------------------------------
Dim i As Long, j As Long
Dim CurrType As Long
Dim CurrInd As Long
Dim CurrER As Double

On Error GoTo FillUMCERsErrorHandler

' Initialize all of the ER entries to 0
ResetERValues Ind

' Now copy the ER values from .Pairs().ER to .ExpressionRatio
With GelP_D_L(Ind)
    For i = 0 To .PCnt - 1
        CurrER = .Pairs(i).ER
        If CurrER < glHugeUnderExp Then
            CurrER = glHugeUnderExp
        ElseIf CurrER > glHugeOverExp Then
            CurrER = glHugeOverExp
        End If
        With GelUMC(Ind).UMCs(.Pairs(i).P1)
            For j = 0 To .ClassCount - 1
                CurrType = .ClassMType(j)
                CurrInd = .ClassMInd(j)
                Select Case CurrType
                Case glCSType
                    GelData(Ind).CSData(CurrInd).ExpressionRatio = CurrER
                Case glIsoType
                    GelData(Ind).IsoData(CurrInd).ExpressionRatio = CurrER
                End Select
            Next j
        End With
        With GelUMC(Ind).UMCs(.Pairs(i).P2)
            For j = 0 To .ClassCount - 1
                CurrType = .ClassMType(j)
                CurrInd = .ClassMInd(j)
                Select Case CurrType
                Case glCSType
                    GelData(Ind).CSData(CurrInd).ExpressionRatio = CurrER
                Case glIsoType
                    GelData(Ind).IsoData(CurrInd).ExpressionRatio = CurrER
                End Select
            Next j
        End With
    Next i
End With
    
GelStatus(Ind).Dirty = True
Exit Sub

FillUMCERsErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "Pairs.Bas->FillUMC_ERs"
    Resume Next
    
End Sub

Public Sub FillSolo_ERs(ByVal Ind As Long)
'-----------------------------------------------------
'fills expression numbers from GelP_D_L to data arrays
'Solo pairs
'-----------------------------------------------------
Dim i As Long
On Error Resume Next

' Initialize all of the ER entries to 0
ResetERValues Ind

' Now copy the ER values from .Pairs(i).ER to .IsoData().ExpressionRatio
With GelP_D_L(Ind)
    For i = 0 To .PCnt - 1
        GelData(Ind).IsoData(.Pairs(i).P1).ExpressionRatio = .Pairs(i).ER
        GelData(Ind).IsoData(.Pairs(i).P2).ExpressionRatio = .Pairs(i).ER
    Next i
End With

GelStatus(Ind).Dirty = True

End Sub

Public Function ChargeStatesMatch(lngGelIndex As Long, i As Long, j As Long) As Boolean
    ' Checks if the two UMC's have matching charge states
    ' i and j specify the UMC indices in GelUMC(lngGelIndex).UMCs()
    '
    ' Returns True if they do, False if they do not
    
    Dim blnChargesMatch As Boolean
    Dim intIndex As Integer
    Dim intIndexCompare As Integer
    
    blnChargesMatch = False
    With GelUMC(lngGelIndex)
        For intIndex = 0 To .UMCs(i).ChargeStateCount - 1
            For intIndexCompare = 0 To .UMCs(j).ChargeStateCount - 1
                If .UMCs(i).ChargeStateBasedStats(intIndex).Charge = .UMCs(j).ChargeStateBasedStats(intIndexCompare).Charge Then
                    blnChargesMatch = True
                    Exit For
                End If
            Next intIndexCompare
            If blnChargesMatch Then Exit For
        Next intIndex
    End With
    
    ChargeStatesMatch = blnChargesMatch
    
End Function

Public Function PairsOverlapAtEdgesWithinTol(lngGelIndex As Long, i As Long, j As Long, lngScanTolerance As Long) As Boolean
    ' Checks if the two UMC's overlap at the edges
    ' i and j specify the UMC indices in GelUMC(lngGelIndex).UMCs()
    '
    ' Returns True if they do, False if they do not
    
    Dim blnOverlap As Boolean
    
    blnOverlap = False
    With GelUMC(lngGelIndex)
        If ((Abs(.UMCs(j).MinScan - .UMCs(i).MinScan) <= lngScanTolerance) And _
            (Abs(.UMCs(j).MaxScan - .UMCs(i).MaxScan) <= lngScanTolerance)) Then
            blnOverlap = True
        Else
           ' If beginnings and/or ends differ by more than ScanTol
           ' but one class is completely overlapped with another,
           ' then this still can be considered a good pair
           If (.UMCs(j).MinScan <= .UMCs(i).MinScan And .UMCs(j).MaxScan >= .UMCs(i).MaxScan) Or _
              (.UMCs(i).MinScan <= .UMCs(j).MinScan And .UMCs(i).MaxScan >= .UMCs(j).MaxScan) Then
                blnOverlap = True
            End If
        End If
    End With
    
    PairsOverlapAtEdgesWithinTol = blnOverlap
End Function

Public Function UpdateUMCsPairingStatus(lngGelIndex As Long, eClsPaired() As umcpUMCPairMembershipConstants) As Boolean
    '----------------------------------------------------------------------
    'examines pairing status of each unique mass classes; this information
    'can be used as filter during identification; returns True if OK
    '
    'Note that this function skips excluded pairs
    '----------------------------------------------------------------------
    
    Dim i As Long
    
    On Error GoTo UpdateUMCsPairingStatusErrorHandler
    
    If GelUMC(lngGelIndex).UMCCnt <= 0 Then
        ReDim eClsPaired(0)
        UpdateUMCsPairingStatus = True
        Exit Function
    End If
    
    ReDim eClsPaired(GelUMC(lngGelIndex).UMCCnt - 1)
    
    With GelP_D_L(lngGelIndex)
        If .PCnt > 0 Then
            For i = 0 To .PCnt - 1
                With .Pairs(i)
                    If .STATE <> glPAIR_Exc Then
                        'light class
                        Select Case eClsPaired(.P1)
                        Case umcpNone
                             eClsPaired(.P1) = umcpLightUnique
                        Case umcpLightUnique
                             eClsPaired(.P1) = umcpLightMultiple
                        Case umcpHeavyUnique, umcpHeavyMultiple
                             eClsPaired(.P1) = umcpLightHeavyMix
                        Case umcpLightMultiple, umcpLightHeavyMix
                             'no changes
                        Case Else
                            ' Unknown eClsPaired value; this shouldn't happen
                            Debug.Assert False
                        End Select
                        
                        'heavy class
                        Select Case eClsPaired(.P2)
                        Case umcpNone
                             eClsPaired(.P2) = umcpHeavyUnique
                        Case umcpHeavyUnique
                             eClsPaired(.P2) = umcpHeavyMultiple
                        Case umcpLightUnique, umcpLightMultiple
                             eClsPaired(.P2) = umcpLightHeavyMix
                        Case umcpHeavyMultiple, umcpLightHeavyMix
                             'no changes
                        Case Else
                            ' Unknown eClsPaired value; this shouldn't happen
                            Debug.Assert False
                        End Select
                    End If
                End With
            Next i
        End If
    End With
    UpdateUMCsPairingStatus = True
    Exit Function
    
UpdateUMCsPairingStatusErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "Pairs.bas->UpdateUMCsPairingStatus"
    UpdateUMCsPairingStatus = False
End Function



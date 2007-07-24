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
    GetERDesc = "LC-MS Feature; Class Representative; Ratio (Light/Heavy)"
Case glER_UMC_REP_LOG
    GetERDesc = "LC-MS Feature; Class Representative; Logarithmic"
Case glER_UMC_REP_ALT
    GetERDesc = "LC-MS Feature; Class Representative; Symmetric Ratio"
Case glER_UMC_AVG_RAT
    GetERDesc = "LC-MS Feature; Class Average; Ratio (Light/Heavy)"
Case glER_UMC_AVG_LOG
    GetERDesc = "LC-MS Feature; Class Average; Logarithmic"
Case glER_UMC_AVG_ALT
    GetERDesc = "LC-MS Feature; Class Average; Symmetric Ratio"
Case glER_UMC_SUM_RAT
    GetERDesc = "LC-MS Feature; Class Sum; Ratio (Light/Heavy)"
Case glER_UMC_SUM_LOG
    GetERDesc = "LC-MS Feature; Class Sum; Logarithmic"
Case glER_UMC_SUM_ALT
    GetERDesc = "LC-MS Feature; Class Sum; Symmetric Ratio"
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

Public Function PairsLookupFPRType(ByVal lngGelIndex As Long, ByVal blnHeavyPairMember As Boolean) As Long
    
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
    strPairType = "LC-MS Feature-based"
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
            strMessage = strMessage & " Make sure pairs are synchronized with latest LC-MS Feature."
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
    
    Dim lngindex As Long
    
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
            For lngindex = 0 To .PCnt - 1
                With .Pairs(lngindex)
                    If IsAMTReferencedByUMC(GelUMC(lngGelIndex).UMCs(.P1), lngGelIndex) Then
                        blnPairHasUMCWithHit(lngindex) = True
                    Else
                        If IsAMTReferencedByUMC(GelUMC(lngGelIndex).UMCs(.P2), lngGelIndex) Then
                            blnPairHasUMCWithHit(lngindex) = True
                        End If
                    End If
                End With
            Next lngindex
            
            ' Step through the pairs again, and exclude those that are currently included but don't have any hits
            ' Keep track of the excluded pairs using lngPairsWithoutHits()
            lngPairsWithoutHitsCount = 0
            For lngindex = 0 To .PCnt - 1
                With .Pairs(lngindex)
                    If Not blnPairHasUMCWithHit(lngindex) And .STATE <> glPAIR_Exc Then
                        lngPairsWithoutHits(lngPairsWithoutHitsCount) = lngindex
                        intPairsWithoutHitsStateSaved(lngPairsWithoutHitsCount) = .STATE
                        lngPairsWithoutHitsCount = lngPairsWithoutHitsCount + 1
                        
                        .STATE = glPAIR_Exc
                    End If
                End With
            Next lngindex
            
            ' Now exclude the ambiguous pairs
            strMessage = PairsSearchMarkAmbiguous(frmCallingForm, lngGelIndex, True, lngPairsWithoutHitsCount)
            
            ' Now restore the pairs that were excluded because they had no hits
            For lngindex = 0 To lngPairsWithoutHitsCount - 1
                With .Pairs(lngPairsWithoutHits(lngindex))
                    .STATE = intPairsWithoutHitsStateSaved(lngindex)
                End With
            Next lngindex
        
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

Public Sub PairSearchIncludeNeutralPairs(ByVal lngGelIndex As Long)
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
    strPairType = "LC-MS Feature-based"
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

Public Function RatER(ByVal LtAbu As Double, ByVal HvAbu As Double) As Double
'------------------------------------------------------
'returns ER as standard ratio of abundance of light and
'heavy pair members
'------------------------------------------------------
On Error Resume Next
RatER = LtAbu / HvAbu
If Err Then RatER = ER_CALC_ERR
End Function

Public Function RatERViaAltER(ByVal dblAltER As Double) As Double
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

Public Function LogER(ByVal LtAbu As Double, ByVal HvAbu As Double) As Double
'------------------------------------------------------
'returns ER as logarithm of abundance ratio of light
'and heavy pair members
'------------------------------------------------------
On Error Resume Next
LogER = Log(LtAbu / HvAbu)
If Err Then LogER = ER_CALC_ERR
End Function

Public Function AltER(ByVal LtAbu As Double, ByVal HvAbu As Double) As Double
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
    ' This sub will create the ClsStat() array for the loaded LC-MS Features,
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
Dim blnReportN15IncompleteIncorporation As Boolean

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
        
        If .SearchDef.N15IncompleteIncorporationMode Then
            sLine = sLine & strSepChar & "N15 Incorporation %"
            blnReportN15IncompleteIncorporation = True
        Else
            blnReportN15IncompleteIncorporation = False
        End If
        
        If PState = glPair_All Then
            ' Only report the state if the reporting mode is All Pairs
            sLine = sLine & strSepChar & "State"
        End If
          
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
                    
                    If blnReportN15IncompleteIncorporation Then
                        sLine = sLine & strSepChar & Round(.DeltaAtomPercentIncorporation, 1)
                    End If
                    
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

Public Function ChargeStatesMatch(ByVal lngGelIndex As Long, ByVal i As Long, ByVal j As Long) As Boolean
    ' Checks if the two LC-MS Features have matching charge states
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

Public Function PairsOverlapAtEdgesWithinTol(ByVal lngGelIndex As Long, ByVal i As Long, ByVal j As Long, ByVal lngScanTolerance As Long) As Boolean
    ' Checks if the two LC-MS Features overlap at the edges
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

Public Function UpdateUMCsPairingStatus(ByVal lngGelIndex As Long, ByRef eClsPaired() As umcpUMCPairMembershipConstants) As Boolean
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



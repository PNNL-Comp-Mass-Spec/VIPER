VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInSilicoDigest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' This class can be used to perform an in-silico digest of an amino acid sequence
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Utilizes the MolecularWeightCalculator class (MwtWinDll.dll)
'
' Last Modified November 17, 2003
'

Private Const ONE_END_TOKEN = "one end"

Private Const crCleavageRuleCount = 10
Public Enum crCleavageRuleConstants
    crNoRule = 0
    crConventionalTrypsin = 1
    crEricPartialTrypsin = 2
    crTrypsinPlusFVLEY = 3
    crKROneEnd = 4
    crTerminiiOnly = 5
    crChymotrypsin = 6
    crChymotrypsinAndTrypsin = 7
    crGluC = 8
    crCyanBr = 9
End Enum

Private Type udtCleavageRulesType
    Description As String
    Residues As String
    RuleIDInParallax As Integer             ' Unused in this software
End Type

Private Type udtCleavageRuleList
    RuleCount As Integer
    Rules() As udtCleavageRulesType         ' 0-based array
End Type

Private mCleavageRules As udtCleavageRuleList

Private mMwtWin As MolecularWeightCalculator
Private mMwtWinLoaded As Boolean

'''Public Sub TestDigester()
'''    Dim strFragments() As String
'''    Dim strFragmentNames() As String
'''    Dim dblFragmentMasses() As Double
'''
'''    objInSilicoDigest.DigestSequence "MNASELQPVVQFLTSTAPFDTFSEEFIRRCAKAVIIGYYSKASGFVQFDAEAPKLYLVRSGAFEVRDPEGVLLDRVAEGEFFGFSTLLSGEKVVNRVAILEDSLVYHFPQALFDQLRSESRHFDKFFTRAFAKRLRHEARFKAKDLATTSRISTLMSSSPITIDAHATVTQAALLMRNSRVSSLLVTDNHKLVGILTDKDLRNRVLASGLDGQIAVHQAMTTSPISISSNALIFEAMLLMSEHNIHHLPIIDEQNTDEVKAIGMVTSTDILRGQGSQPLLLIGEIERQRDLASLISVSKQIPVLLQNLISADARAEEIGRVLTSVTDALTRRLIVLNQQILGEAPMAFCWLTFGSQGRQDQAACSDQDNGLLVAEEMDDFAKGYFDALTHAVCAGLDQCGYMFCPGNIMAQNPKWRMSLNQWQKVFEKWVITPEPKALMHASIFFDMRSVYGPQSLFDALQDKVLAQTKDNDIFLAGMAGNSLIESPPLGFFRKFVLERDGSEVKGIDLKHKGNALINDIARVYALSAGIREVNTAKRIRALMDANILNRKDALNLADAHEFIAHMRLSNQGYQHTQGLKISNYLLPGHLSSLVRHQLRDAFKVVHDAQSGMKMKFMRSFMNASELQPVVQFLTSTAPFDTFSEEFIRRCAKAVIIGYYSKASGFVQFDAEAPKLYLVR", strFragments(), strFragmentNames(), dblFragmentMasses(), 1, crConventionalTrypsin, 4, 5000, True
'''End Sub

Public Function CheckSequenceAgainstCleavageRule(strSequence As String, ByVal eRuleID As crCleavageRuleConstants, Optional ByRef intRuleMatchCount As Integer = 0) As Boolean
    ' Checks strSequence against the rule given by eRuleID
    ' See sub InitializeCleavageRules for a list of the rules
    ' Returns True if valid, False if invalid
    ' intRuleMatchCount returns 0, 1, or 2:  0 if neither end matches, 1 if one end matches, 2 if both ends match
    '
    ' In order to check for Exception residues, strSequence must be in the form "R.ABCDEFGK.L" so that the residue following the final residue of the fragment can be examined
    
    Dim strRuleResidues As String, strExceptionSuffixResidues As String
    Dim blnAllowPartialCleavage As Boolean
    
    intRuleMatchCount = 0
    If eRuleID >= 0 And eRuleID < mCleavageRules.RuleCount Then
        If eRuleID = 0 Then
            ' No cleavage rule; no point in checking
            CheckSequenceAgainstCleavageRule = True
            intRuleMatchCount = 2
        Else
                        
            strRuleResidues = GetCleaveageRuleResiduesSymbols(eRuleID)
            strExceptionSuffixResidues = GetCleavageExceptionSuffixResidues(eRuleID)
            blnAllowPartialCleavage = GetCleavageAllowPartialCleavage(eRuleID)
            
            If mMwtWinLoaded Then
                CheckSequenceAgainstCleavageRule = mMwtWin.Peptide.CheckSequenceAgainstCleavageRule(strSequence, strRuleResidues, strExceptionSuffixResidues, blnAllowPartialCleavage, , , , intRuleMatchCount)
            Else
                ' objMwtWin not loaded; assume True
                CheckSequenceAgainstCleavageRule = True
            End If
            
        End If
    Else
        ' No rule selected; assume true
        CheckSequenceAgainstCleavageRule = True
        intRuleMatchCount = 2
    End If
    
End Function

Public Function CountTrypticsInSequence(strSequence As String) As Long
    Dim lngTrypticCount As Long
    Dim lngStartSearchLoc As Long, lngReturnResidueStart As Long, lngReturnResidueEnd As Long
    Dim strFragment As String
       
    On Error GoTo CountTrypticsInSequenceErrorHandler
    
    lngTrypticCount = 0
    lngStartSearchLoc = 1
    
    If Len(strSequence) > 0 And mMwtWinLoaded Then
        Do
            strFragment = mMwtWin.Peptide.GetTrypticPeptideNext(strSequence, lngStartSearchLoc, lngReturnResidueStart, lngReturnResidueEnd)
            If Len(strFragment) > 0 Then
                lngTrypticCount = lngTrypticCount + 1
                lngStartSearchLoc = lngReturnResidueEnd + 1
            Else
                Exit Do
            End If
        Loop
    End If
    
    CountTrypticsInSequence = lngTrypticCount
    
    Exit Function

CountTrypticsInSequenceErrorHandler:
    ' Note: Error -2147024770 or 429 will occur if unable to properly access objMwtWin
    CountTrypticsInSequence = 0
    
End Function

Public Function DigestSequence(ByVal strProteinSequence As String, ByRef strFragments() As String, ByRef strFragmentNames() As String, ByRef dblFragmentMasses() As Double, Optional intMaxMissedCleavages As Integer = 0, Optional eCleavageRuleID As crCleavageRuleConstants = crConventionalTrypsin, Optional intMinFragmentResidueCount As Integer = 4, Optional dblMaxFragmentMass As Double = 5000, Optional blnRemoveDuplicateSequences As Boolean = False, Optional strProteinName As String = "") As Long
    ' Digests strProteinSequence using the sequence rule given by eCleavageRuleID
    ' Returns the fragments in strFragments() and their names in strFragmentNames()
    ' If blnRemoveDuplicateSequences = True, then only returns the first occurrence of each unique sequence
    
    Dim htFragmentsUniqueList As New Scripting.Dictionary
        
    Dim lngFragCountDimmed As Long
    Dim lngTrypticFragmentCount As Long, lngFragmentCountTotal As Long
    Dim intTrypticIndex As Integer
    Dim lngSearchStartLoc As Long
    Dim lngResidueStartLoc As Long, lngResidueEndLoc As Long
    Dim lngResidueLength As Long
    Dim lngResidueLengthStart As Long
    
    Dim intIndex As Integer

    Dim intTrypticFragCacheCountDimmed As Integer
    Dim intTrypticFragCacheCount As Integer
    Dim strTrypticFragCache() As String             ' 0-based array
    Dim lngTrypticFragStartLocs() As Long           ' 0-based array, parallel to strTypticFragmentCache()
    Dim lngTrypticFragEndLocs() As Long             ' 0-based array, parallel to strTypticFragmentCache()
    
    Dim strPeptideSequence As String, strPeptideSequenceBase As String
    Dim strRuleResidues As String, strExceptionSuffixResidues As String
    
    Dim blnPeptideAdded As Boolean
    Dim objProgress As frmProgress
    
    If Not mMwtWinLoaded Or Len(strProteinSequence) = 0 Then
        ReDim strFragments(0)
        ReDim strFragmentNames(0)
        ReDim dblFragmentMasses(0)
        DigestSequence = 0
        Exit Function
    End If
    
    ' Make sure the correct mass mode is enabled
    mMwtWin.SetElementMode emIsotopicMass
    
    If intMaxMissedCleavages < 0 Then intMaxMissedCleavages = 0
    If intMaxMissedCleavages > 20 Then intMaxMissedCleavages = 20
    
    strRuleResidues = GetCleaveageRuleResiduesSymbols(eCleavageRuleID)
    strExceptionSuffixResidues = GetCleavageExceptionSuffixResidues(eCleavageRuleID)
    
    lngTrypticFragmentCount = CountTrypticsInSequence(strProteinSequence)
    
    ' Increment lngTrypticFragmentCount to account for missed cleavages
    ' This will be drastically low if using partial cleavage, but it is a starting point
    lngFragmentCountTotal = 0
    For intIndex = intMaxMissedCleavages + 1 To 1 Step -1
        lngFragmentCountTotal = lngFragmentCountTotal + intIndex * lngTrypticFragmentCount
    Next intIndex
    
    lngFragCountDimmed = lngFragmentCountTotal
    ReDim strFragments(lngFragmentCountTotal)
    ReDim strFragmentNames(lngFragmentCountTotal)
    ReDim dblFragmentMasses(lngFragmentCountTotal)
    
    intTrypticFragCacheCountDimmed = 10
    ReDim strTrypticFragCache(0 To intTrypticFragCacheCountDimmed - 1)
    ReDim lngTrypticFragEndLocs(0 To intTrypticFragCacheCountDimmed - 1)
    ReDim lngTrypticFragStartLocs(0 To intTrypticFragCacheCountDimmed - 1)
    
    lngFragmentCountTotal = 0
    intTrypticFragCacheCount = 0
    lngSearchStartLoc = 1
    
    ' Using the GetTrypticPeptideNext function to retrieve the sequence for each tryptic peptide
    '   is faster than using the GetTrypticPeptideByFragmentNumber function
    ' Populate strTrypticFragCache()
    Do
        strPeptideSequence = mMwtWin.Peptide.GetTrypticPeptideNext(strProteinSequence, lngSearchStartLoc, lngResidueStartLoc, lngResidueEndLoc, strRuleResidues, strExceptionSuffixResidues)
        If Len(strPeptideSequence) > 0 Then
            strTrypticFragCache(intTrypticFragCacheCount) = strPeptideSequence
            lngTrypticFragStartLocs(intTrypticFragCacheCount) = lngResidueStartLoc
            lngTrypticFragEndLocs(intTrypticFragCacheCount) = lngResidueEndLoc
            intTrypticFragCacheCount = intTrypticFragCacheCount + 1
            
            If intTrypticFragCacheCount >= intTrypticFragCacheCountDimmed Then
                intTrypticFragCacheCountDimmed = intTrypticFragCacheCountDimmed + 10
                ReDim Preserve strTrypticFragCache(0 To intTrypticFragCacheCountDimmed - 1)
                ReDim Preserve lngTrypticFragStartLocs(0 To intTrypticFragCacheCountDimmed - 1)
                ReDim Preserve lngTrypticFragEndLocs(0 To intTrypticFragCacheCountDimmed - 1)
            End If
            
            lngSearchStartLoc = lngResidueEndLoc + 1
        Else
            Exit Do
        End If
    Loop
    
    If eCleavageRuleID = crKROneEnd Then
        Set objProgress = New frmProgress
        objProgress.MoveToBottomCenter
        objProgress.Show
        objProgress.InitializeForm "Digesting protein " & strProteinName, 0, intTrypticFragCacheCount * 2, False, False
    End If
        
    For intTrypticIndex = 0 To intTrypticFragCacheCount - 1
        strPeptideSequenceBase = ""
        strPeptideSequence = ""
        lngResidueStartLoc = lngTrypticFragStartLocs(intTrypticIndex)
        
        For intIndex = 0 To intMaxMissedCleavages
            If intTrypticIndex + intIndex >= intTrypticFragCacheCount Then
                Exit For
            End If
            
            If eCleavageRuleID = crKROneEnd Then
                ' Partially tryptic cleavage rule: Add all partially tryptic fragments
                If intIndex = 0 Then
                    lngResidueLengthStart = intMinFragmentResidueCount
                Else
                    lngResidueLengthStart = 1
                End If
                
                For lngResidueLength = lngResidueLengthStart To Len(strTrypticFragCache(intTrypticIndex + intIndex))
                    If intIndex > 0 Then
                        lngResidueEndLoc = lngTrypticFragEndLocs(intTrypticIndex + intIndex - 1) + lngResidueLength
                    Else
                        lngResidueEndLoc = lngResidueStartLoc + lngResidueLength - 1
                    End If
                    
                    strPeptideSequence = strPeptideSequenceBase & Left(strTrypticFragCache(intTrypticIndex + intIndex), lngResidueLength)
                                        
                    If Len(strPeptideSequence) >= intMinFragmentResidueCount Then
                        PossiblyAddPeptide strPeptideSequence, intTrypticIndex, intIndex, lngResidueStartLoc, lngResidueEndLoc, htFragmentsUniqueList, strFragments, strFragmentNames, dblFragmentMasses, blnRemoveDuplicateSequences, dblMaxFragmentMass, lngFragmentCountTotal, lngFragCountDimmed, eCleavageRuleID
                    End If
                Next lngResidueLength
            Else
                ' Normal cleavage rule
                lngResidueEndLoc = lngTrypticFragEndLocs(intTrypticIndex + intIndex)
                
                strPeptideSequence = strPeptideSequence & strTrypticFragCache(intTrypticIndex + intIndex)
                If Len(strPeptideSequence) >= intMinFragmentResidueCount Then
                    PossiblyAddPeptide strPeptideSequence, intTrypticIndex, intIndex, lngResidueStartLoc, lngResidueEndLoc, htFragmentsUniqueList, strFragments, strFragmentNames, dblFragmentMasses, blnRemoveDuplicateSequences, dblMaxFragmentMass, lngFragmentCountTotal, lngFragCountDimmed, eCleavageRuleID
                End If
            End If
            
            strPeptideSequenceBase = strPeptideSequenceBase & strTrypticFragCache(intTrypticIndex + intIndex)
        Next intIndex
        
        If eCleavageRuleID = crKROneEnd Then
            objProgress.UpdateProgressBar intTrypticIndex
        End If
    Next intTrypticIndex
    
    If eCleavageRuleID = crKROneEnd Then
        ' Partially tryptic cleavage rule: Add all partially tryptic fragments, working from the end toward the front
        For intTrypticIndex = intTrypticFragCacheCount - 1 To 0 Step -1
            strPeptideSequenceBase = ""
            strPeptideSequence = ""
            lngResidueEndLoc = lngTrypticFragEndLocs(intTrypticIndex)
            
            For intIndex = 0 To intMaxMissedCleavages
                If intTrypticIndex - intIndex < 0 Then
                    Exit For
                End If
                
                If intIndex = 0 Then
                    lngResidueLengthStart = intMinFragmentResidueCount
                Else
                    lngResidueLengthStart = 1
                End If
                    
                ' We can limit the following for loop to the peptide length - 1 since those peptides using the full peptide will have already been added above
                For lngResidueLength = lngResidueLengthStart To Len(strTrypticFragCache(intTrypticIndex - intIndex)) - 1
                    ' The following is wrong:
                    Debug.Assert False
                    If intIndex > 0 Then
                        lngResidueStartLoc = lngTrypticFragStartLocs(intTrypticIndex - intIndex + 1) - lngResidueLength
                    Else
                        lngResidueStartLoc = lngResidueEndLoc - (lngResidueLength - 1)
                    End If
                    
                    ' Use of the Right() function grabs characters from the right edge of the string, but returns them in the forward reading direction (i.e., their order is not reversed)
                    strPeptideSequence = Right(strTrypticFragCache(intTrypticIndex - intIndex), lngResidueLength) & strPeptideSequenceBase
                                        
                    If Len(strPeptideSequence) >= intMinFragmentResidueCount Then
                        blnPeptideAdded = PossiblyAddPeptide(strPeptideSequence, intTrypticIndex, intIndex, lngResidueStartLoc, lngResidueEndLoc, htFragmentsUniqueList, strFragments, strFragmentNames, dblFragmentMasses, blnRemoveDuplicateSequences, dblMaxFragmentMass, lngFragmentCountTotal, lngFragCountDimmed, eCleavageRuleID)
                    End If
                    
                Next lngResidueLength
                
                strPeptideSequenceBase = strTrypticFragCache(intTrypticIndex - intIndex) & strPeptideSequenceBase
            Next intIndex
            
            'objProgress.UpdateProgressBar intTrypticFragCacheCount + (intTrypticFragCacheCount - intTrypticIndex)
            objProgress.UpdateProgressBar intTrypticFragCacheCount * 2 - intTrypticIndex

        Next intTrypticIndex
        
        Unload objProgress
        Set objProgress = Nothing
    End If
    
    
    If lngFragmentCountTotal <> lngFragCountDimmed And lngFragmentCountTotal > 0 Then
        ReDim Preserve strFragments(lngFragmentCountTotal - 1)
        ReDim Preserve strFragmentNames(lngFragmentCountTotal - 1)
        ReDim Preserve dblFragmentMasses(lngFragmentCountTotal - 1)
    End If
    
    DigestSequence = lngFragmentCountTotal
    
End Function

Private Function PossiblyAddPeptide(ByVal strPeptideSequence As String, ByRef intTrypticIndex As Integer, ByRef intMissedCleavageCount As Integer, ByRef lngResidueStartLoc As Long, ByRef lngResidueEndLoc As Long, ByRef htFragmentsUniqueList As Scripting.Dictionary, ByRef strFragments() As String, ByRef strFragmentNames() As String, ByRef dblFragmentMasses() As Double, ByRef blnRemoveDuplicateSequences As Boolean, ByRef dblMaxFragmentMass As Double, ByRef lngFragmentCountTotal As Long, ByRef lngFragCountDimmed As Long, ByRef eCleavageRuleID As crCleavageRuleConstants) As Boolean

    Dim blnAddFragment As Boolean
    Dim dblPeptideMass As Double
    
    blnAddFragment = True
    If blnRemoveDuplicateSequences Then
        If htFragmentsUniqueList.Exists(strPeptideSequence) Then
            blnAddFragment = False
        Else
            htFragmentsUniqueList.add strPeptideSequence, 1
        End If
    End If
    
    If blnAddFragment Then
    
        mMwtWin.Peptide.SetSequence strPeptideSequence, ntgHydrogen, ctgHydroxyl, False, False
        dblPeptideMass = mMwtWin.Peptide.GetPeptideMass
        
        If dblPeptideMass <= dblMaxFragmentMass Then
            ' Add to strFragments()
            
            strFragments(lngFragmentCountTotal) = strPeptideSequence
            
            If eCleavageRuleID = crConventionalTrypsin Then
                strFragmentNames(lngFragmentCountTotal) = "t" & Trim(intTrypticIndex + 1) & "." & Trim(intMissedCleavageCount + 1)
            Else
                strFragmentNames(lngFragmentCountTotal) = Trim(lngResidueStartLoc) & "." & Trim(lngResidueEndLoc)
            End If
            
            dblFragmentMasses(lngFragmentCountTotal) = dblPeptideMass
            
            lngFragmentCountTotal = lngFragmentCountTotal + 1
            If lngFragmentCountTotal >= lngFragCountDimmed Then
                If eCleavageRuleID = crKROneEnd Then
                    lngFragCountDimmed = lngFragCountDimmed + 1000
                Else
                    lngFragCountDimmed = lngFragCountDimmed + 100
                End If
                ReDim Preserve strFragments(lngFragCountDimmed)
                ReDim Preserve strFragmentNames(lngFragCountDimmed)
                ReDim Preserve dblFragmentMasses(lngFragCountDimmed)
            End If
        End If
    End If

    PossiblyAddPeptide = blnAddFragment
End Function

Public Function GetCleavageAllowPartialCleavage(eRuleID As crCleavageRuleConstants) As Boolean
    Dim blnAllowPartialCleavage As Boolean
    
    Select Case eRuleID
    Case crEricPartialTrypsin
        blnAllowPartialCleavage = True
    Case crKROneEnd
        blnAllowPartialCleavage = True
    Case Else
        blnAllowPartialCleavage = False
    End Select
    
    If InStr(UCase(mCleavageRules.Rules(eRuleID).Residues), UCase(ONE_END_TOKEN)) > 0 Then
        blnAllowPartialCleavage = True
    End If
    
    GetCleavageAllowPartialCleavage = blnAllowPartialCleavage
End Function

Public Function GetCleavageExceptionSuffixResidues(eRuleID As crCleavageRuleConstants) As String
            
    If eRuleID = crConventionalTrypsin Or eRuleID = crKROneEnd Then
        GetCleavageExceptionSuffixResidues = "P"
    Else
        GetCleavageExceptionSuffixResidues = ""
    End If

End Function

Public Function GetCleaveageRuleName(eRuleID As crCleavageRuleConstants) As String
    If eRuleID >= 0 And eRuleID < mCleavageRules.RuleCount Then
        GetCleaveageRuleName = mCleavageRules.Rules(eRuleID).Description
    End If
End Function

Public Function GetCleaveageRuleResiduesDescription(eRuleID As crCleavageRuleConstants) As String
    If eRuleID >= 0 And eRuleID < mCleavageRules.RuleCount Then
        GetCleaveageRuleResiduesDescription = mCleavageRules.Rules(eRuleID).Residues
    Else
        GetCleaveageRuleResiduesDescription = ""
    End If
End Function

Public Function GetCleaveageRuleResiduesSymbols(eRuleID As crCleavageRuleConstants) As String
    Dim strResiduesSymbols As String
    
    Select Case eRuleID
    Case crKROneEnd
        strResiduesSymbols = "KR"
    Case crTerminiiOnly
        strResiduesSymbols = "-"
    Case Else
        strResiduesSymbols = mCleavageRules.Rules(eRuleID).Residues
    End Select

    If InStr(UCase(strResiduesSymbols), UCase(ONE_END_TOKEN)) > 0 Then
        strResiduesSymbols = Trim(Replace(strResiduesSymbols, ONE_END_TOKEN, ""))
    End If
    
    GetCleaveageRuleResiduesSymbols = strResiduesSymbols
        
End Function

Public Function GetCleaveageRuleIDInParallax(eRuleID As crCleavageRuleConstants) As Integer
    If eRuleID >= 0 And eRuleID < mCleavageRules.RuleCount Then
        GetCleaveageRuleIDInParallax = mCleavageRules.Rules(eRuleID).RuleIDInParallax
    End If
    
End Function

Private Sub InitializeCleavageRules()
    
    With mCleavageRules
        .RuleCount = crCleavageRuleCount
        ReDim .Rules(.RuleCount)            ' 0-based array
        
        With .Rules(crNoRule)
            .Description = "No cleavage rule"
            .Residues = ""
            .RuleIDInParallax = 0
        End With
        
        With .Rules(crConventionalTrypsin)
            .Description = "Fully Tryptic"
            .Residues = "KR"
            .RuleIDInParallax = 10
        End With
        
        With .Rules(crEricPartialTrypsin)
            .Description = "Eric's Partial Trypsin"         ' Allows partial cleavage
            .Residues = "KRFYVEL"
            .RuleIDInParallax = 11
        End With
        
        With .Rules(crTrypsinPlusFVLEY)
            .Description = "Trypsin Plus FVLEY"             ' Does not allow partial cleavage
            .Residues = "KRFYVEL"
            .RuleIDInParallax = 12
        End With
        
        With .Rules(crKROneEnd)
            .Description = "Half (Partial) Trypsin "
            .Residues = "KR" & " " & ONE_END_TOKEN
            .RuleIDInParallax = 13
        End With
        
        With .Rules(crTerminiiOnly)
            .Description = "Peptide Database"
            .Residues = "terminii only"
            .RuleIDInParallax = 20
        End With
        
        With .Rules(crChymotrypsin)
            .Description = "Chymotrypsin"
            .Residues = "FWYL"
            .RuleIDInParallax = 30
        End With
        
        With .Rules(crChymotrypsinAndTrypsin)
            .Description = "Chymotrypsin + Trypsin"
            .Residues = "FWYLKR"
            .RuleIDInParallax = 31
        End With
        
        With .Rules(crGluC)
            .Description = "GluC"
            .Residues = "ED"
            .RuleIDInParallax = 40
        End With
        
        With .Rules(crCyanBr)
            .Description = "CyanBr"
            .Residues = "M"
            .RuleIDInParallax = 50
        End With
    
    End With

End Sub

Private Sub InitializeMwtWinInternal()

On Error GoTo InitializeMwtWinError
    
    Set mMwtWin = New MolecularWeightCalculator
    
    ' If the above statement didn't generate an error, then try to set the element mode
    ' If this doesn't generate an error, then mMwtWin loaded successfully
    mMwtWin.SetElementMode emIsotopicMass
    mMwtWinLoaded = True
    
    Exit Sub
    
InitializeMwtWinError:
    
    mMwtWinLoaded = False
    
End Sub

Public Property Get CleaveageRuleCount() As Long
    CleaveageRuleCount = mCleavageRules.RuleCount
End Property

Private Sub Class_Initialize()

    InitializeMwtWinInternal
    
    InitializeCleavageRules
        
End Sub

Private Sub Class_Terminate()
    Set mMwtWin = Nothing
End Sub


Attribute VB_Name = "OrfViewerRoutines"
Option Explicit

' Written by Matthew Monroe, PNNL
' Started January 2, 2003

' Public constants
Public Const DEFAULT_ORF_PICTURE_HEIGHT = 250
Public Const DEFAULT_ORF_PICTURE_WIDTH = 300
Public Const DEFAULT_ORF_PICTURE_SPACING_PIXELS = 1

Public Const DEFAULT_ORF_MAX_SPOT_SIZE_PIXELS = 500
Public Const DEFAULT_ORF_MIN_SPOT_SIZE_PIXELS = 75

Public Const DEFAULT_ORF_MASS_RANGE_PPM = 25
Public Const DEFAULT_ORF_NET_RANGE = 0.12
Public Const DEFAULT_ORF_MASS_TAG_MASS_ERROR_PPM = 20
Public Const DEFAULT_ORF_MASS_TAG_NET_ERROR = 0.1

Public Const DEFAULT_ORF_LISTVIEW_INTENSITY_SCALAR = 1000000
Public Const DEFAULT_ORF_PICTURE_ION_TO_UMC_INTENSITY_SCALING_RATIO = 2

Public Const ORF_VIEWER_MASS_TAG_STRING = "Mass Tag "
Public Const ORF_VIEWER_ION_STRING = "Ion "
Public Const ORF_VIEWER_UMC_STRING = "UMC "
Public Const ORF_VIEWER_ID_DELIMETER = ":"
Public Const ORF_VIEWER_UMC_ION_LIST_START_STRING = " = "
Public Const ORF_VIEWER_UMC_REPRESENTATIVE_MEMBER_INDICATOR = "**"
Public Const ORF_VIEWER_UMC_ION_LIST_DELIMETER = ", "

' Values stored in .OtherInfo
Public Const UMC_COUNT_LAST_RECORD_ION_MATCH_CALL = "UMCCountLastRecordIonMatchCall"

' Remote DB Table Constants
Public Const TBL_MASS_TAGS = "T_Mass_Tags"
Public Const TBL_MTSUBSET_MASS_TAGS = "T_MTSubset_Mass_Tags"
Private Const TBL_ORF_REFERENCE = "T_ORF_Reference"
Private Const TBL_EXTERNAL_DATABASES = "T_External_Databases"
Private Const TBL_ORF_DETAILS = "T_ORF"
Private Const TBL_MASS_TAG_TO_ORF_MAP = "T_Mass_Tag_to_ORF_Map"
Private Const TBL_MASS_TAGS_NET = "T_Mass_Tags_NET"

' Field Constants
Private Const FIELD_ORF_DB_NAME = "ORF_DB_Name"
Private Const FIELD_EFFECTIVE_DATE = "Effective_Date"

' Other constants
Private Const ORF_DIM_CHUNK = 100                   ' Reserve memory for ORF's in chunks of 100
    

' The following index array can be used to quickly find any mass tag by mass tag Ref ID (using a binary search)
' MassTagRefID() is filled with the RefID's of all of the mass tags, and is sorted ascending
' The MassTagRefIDPointer() gets sorted along with MassTagRefID(), allowing for dereferencing
' The ORFIndex() and MassTagIndex() arrays do not get sorted
Private Type udtMassTagRefIDIndexType
    MassTagCount As Long

    MassTagRefID() As Long              ' 0-based array.  Sorted, ascending
    MassTagRefIDPointer() As Long
    
    ORFIndex() As Long            ' Pointer to be used to dereference x in GelORFData(GelIndex).Orfs(x)
    MassTagIndex() As Long        ' Pointer to be used to dereference x in GelORFMassTags(GelIndex).Orfs(ORFIndex).MassTags(x)

End Type

Public gOrfViewerOptionsSavedGelList As udtORFViewerGelListType
Public gOrfViewerOptionsCurrentGelList As udtORFViewerGelListType

Private objGANET As GANETClass
Private mGANETLoaded As Boolean

Private objInSilicoDigest As New clsInSilicoDigest
'

' Unused Function (March 2003)
'''Public Function ComputeMassByLocationInProtein(strProteinSequence As String, lngResidueStart As Long, lngResidueEnd As Long, blnProtonated As Boolean) As Double
'''    ' Determines the mass of the residues in strProteinSequenceranging from
'''    '  lngResidueStart to lngResidueEnd
'''    Dim eNTerminus As ntgNTerminusGroupConstants
'''
'''    If Len(strProteinSequence) = 0 Then
'''        ComputeMassByLocationInProtein = 0
'''    ElseIf lngResidueStart < 1 Or lngResidueEnd < 1 Then
'''        ComputeMassByLocationInProtein = 0
'''    Else
'''        If blnProtonated Then
'''            eNTerminus = ntgHydrogenPlusProton
'''        Else
'''            eNTerminus = ntgHydrogen
'''        End If
'''
'''        If gMwtWinLoaded Then
'''            objMwtWin.Peptide.SetSequence Mid(strProteinSequence, lngResidueStart, lngResidueEnd - lngResidueStart + 1), eNTerminus, ctgHydroxyl, False, False
'''            ComputeMassByLocationInProtein = objMwtWin.Peptide.GetPeptideMass
'''        Else
'''            ComputeMassByLocationInProtein = 0
'''        End If
'''    End If
'''
'''End Function

Public Function CheckSequenceAgainstCleavageRuleWrapper(strSequence As String, intRuleID As Integer, Optional ByRef intRuleMatchCount As Integer = 0) As Boolean
    CheckSequenceAgainstCleavageRuleWrapper = objInSilicoDigest.CheckSequenceAgainstCleavageRule(strSequence, CInt(intRuleID), intRuleMatchCount)
End Function

Public Sub ComputeTheoreticalTrypticMassTags(ByRef udtGelORF As udtORFListType, ByRef udtGelORFMassTags As udtORFMassTagsListType, ByVal lngGelIndex As Long, Optional ByVal lngMinimumResidueCountToUse As Long = 5, Optional ByVal dblMaximumPeptideMassToUse As Double = 6000, Optional ByRef blnCopiedValuesFromOtherGel As Boolean = False)
    ' Examine the mass tags for each ORF to determine all of the missing tryptic mass tags
    ' Add the missing tryptic mass tags to GelORFMassTags(), with the correct mass value
    '  and theoretical NET values
    '
    ' Peptides shorter than lngMinimumResidueCountToUse residues or heavier than dblMaximumPeptideMassToUse Da are not included
    
    Dim blnTrypticMassTagsPresent() As Boolean           ' 1-based array; records whether a given tryptic mass tag is present
    Dim lngORFIndex As Long
    Dim lngMassTagIndex As Long, lngMassTagIndexCompare As Long
    Dim intTrypticFragmentCount As Integer
    Dim intTrypticMassTagIndex As Integer
    Dim intTrypticMassTagStartNum As Integer, intTrypticMassTagsInARow As Integer
    Dim strFormatString As String
    Dim strPeptideSequence As String
    Dim dblPeptideMass As Double
    Dim lngResidueStartLoc As Long, lngResidueEndLoc As Long
    Dim lngSearchStartLoc As Long
    Dim lngTheoreticalMassTagsAdded As Long
    Dim intORFMassTagCount As Integer
    
    Dim OutSeqFileNum As Integer, OutRefFileNum As Integer
    Dim strOutFilePath As String
    Dim intOpenFileFailCount As Integer
    Dim blnOutputToFile As Boolean
    Dim eResponse As VbMsgBoxResult
    Dim lngSourceGelIndex As Long
    Dim blnMatchFound As Boolean
    Dim lngNewMassTagCount As Long
    
    blnCopiedValuesFromOtherGel = False
    If UBound(GelBody()) > 1 Then
        ' Examine the other Gels to see if any have the needed theoretical tryptic mass tag information
        ' If one is found, copy the information from that Gel to this one
        For lngSourceGelIndex = 1 To UBound(GelBody())
            If lngSourceGelIndex <> lngGelIndex Then
                With GelORFMassTags(lngSourceGelIndex)
                    If GelORFData(lngSourceGelIndex).Definition.MTDBConnectionString = GelORFData(lngGelIndex).Definition.MTDBConnectionString And _
                       GelORFData(lngSourceGelIndex).Definition.DataParsedCompletely Then
                        If .ORFCount = GelORFMassTags(lngGelIndex).ORFCount Then
                            If .Definition.IncludesTheoreticalTrypticMassTags And .Definition.TheoreticalTrypticMassTagsSuccessfullyAdded Then
                                blnMatchFound = True
                            End If
                        End If
                    End If
                End With
                If blnMatchFound Then Exit For
            End If
        Next lngSourceGelIndex
    
        If blnMatchFound Then
            frmProgress.InitializeForm "Copying from other Gel file in memory", 0, GelORFMassTags(lngSourceGelIndex).ORFCount, False, False, True

            blnCopiedValuesFromOtherGel = True
            With GelORFMassTags(lngSourceGelIndex)
                For lngORFIndex = 0 To .ORFCount - 1
                    With .Orfs(lngORFIndex)
                        If .RefID = GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).RefID Then
                            ' Reserve extra memory in the target array, if needed
                            If GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount < .MassTagCount Then
                                ReDim Preserve GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(.MassTagCount)
                            End If
                            
                            For lngMassTagIndex = 0 To .MassTagCount - 1
                                If .MassTags(lngMassTagIndex).IsTheoretical Then
                                    ' Found a theoretical mass tag; make sure it isn't already present in the target array
                                    blnMatchFound = False
                                    For lngMassTagIndexCompare = 0 To GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount - 1
                                        If GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(lngMassTagIndexCompare).MassTagRefID = .MassTags(lngMassTagIndex).MassTagRefID Then
                                            blnMatchFound = True
                                            Exit For
                                        End If
                                    Next lngMassTagIndexCompare
                                    
                                    If Not blnMatchFound Then
                                        lngNewMassTagCount = GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount + 1
                                        GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount = lngNewMassTagCount
                                        GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(lngNewMassTagCount - 1) = .MassTags(lngMassTagIndex)
                                    End If
                                End If
                            Next lngMassTagIndex
                        Else
                            ' This shouldn't happen
                            Debug.Assert False
                            blnCopiedValuesFromOtherGel = False
                        End If
                    End With
                    If lngORFIndex Mod 100 = 0 Then
                        frmProgress.UpdateProgressBar lngORFIndex
                    End If
                Next lngORFIndex
            End With
            
            frmProgress.HideForm
            If blnCopiedValuesFromOtherGel Then
                AddToAnalysisHistory lngGelIndex, "Loaded theoretical mass tags for ORFs from other Gel file in memory; FilePath = " & GelData(lngSourceGelIndex).FileName & "); Total theoretical Mass Tags Added = " & lngTheoreticalMassTagsAdded & "; ORFs processed = " & GelORFMassTags(lngGelIndex).ORFCount
                Exit Sub
            End If
        End If
    End If
    
    If gMwtWinLoaded Then
        eResponse = MsgBox("Compute theoretical tryptic mass tags?", vbQuestion + vbYesNo + vbDefaultButton2, "Include Theoretical MT")
    Else
        MsgBox "Unable to compute theoretical tryptic mass tags since the Molecular Weight Calculator Dll file (MwtWinDll.Dll) could not be loaded.", vbInformation + vbOKOnly, "Unable to compute"
        eResponse = vbNo
    End If
    
    If eResponse = vbNo Or udtGelORF.ORFCount = 0 Then
        udtGelORFMassTags.Definition.IncludesTheoreticalTrypticMassTags = False
        Exit Sub
    Else
        udtGelORFMassTags.Definition.IncludesTheoreticalTrypticMassTags = True
    End If
    
    Debug.Assert udtGelORF.ORFCount = udtGelORFMassTags.ORFCount
    If udtGelORF.ORFCount <> udtGelORFMassTags.ORFCount Then
        udtGelORFMassTags.ORFCount = udtGelORF.ORFCount
        
    End If
    
    frmProgress.InitializeForm "Computing theoretical tryptic mass tags", 0, udtGelORF.ORFCount, True, False, True, MDIForm1
    
On Error GoTo ComputeTheoreticalTrypticMassTagsErrorHandler

    ' Make sure we're using isotopic masses
    objMwtWin.SetElementMode emIsotopicMass
    
    blnOutputToFile = Not mGANETLoaded
    
    If blnOutputToFile Then
        ' Need to write the sequences of the theoretical ions to an output file so that Lars can compute a NET value for each
        ' I will write the sequences to OutSeqFileNum
        ' and the ORFIndex, MassTagIndex, and Sequence to OutRefFileNum
    
        ' I first try to write to a file based on the FileName stored in GelData().FileName
        
        If Len(GelBody(lngGelIndex).Caption) > 0 Then
            strOutFilePath = GelBody(lngGelIndex).Caption & ".NET"
        Else
            strOutFilePath = AppendToPath(App.Path, "Temp_ComputeNETValues.NET")
        End If
    End If
    
    lngTheoreticalMassTagsAdded = 0
    intOpenFileFailCount = 0
    
ComputeTheoreticalTrypticMassTagsOpenFiles:
    
On Error GoTo ComputeTheoreticalTrypticMassTagsOpenFileError
    
    If blnOutputToFile Then
        OutSeqFileNum = FreeFile()
        Open strOutFilePath For Output As #OutSeqFileNum
        
        OutRefFileNum = FreeFile()
        strOutFilePath = strOutFilePath & "REF"
        Open strOutFilePath For Output As #OutRefFileNum
    End If
    
ComputeTheoreticalTrypticMassTagsWork:

On Error GoTo ComputeTheoreticalTrypticMassTagsErrorHandler

    strFormatString = ListViewConstructSortKeyFormatString(udtGelORF.ORFCount - 1)
    
    For lngORFIndex = 0 To udtGelORF.ORFCount - 1
        intTrypticFragmentCount = udtGelORF.Orfs(lngORFIndex).TrypticFragmentCount
        ReDim blnTrypticMassTagsPresent(intTrypticFragmentCount)
        
        With udtGelORFMassTags.Orfs(lngORFIndex)
                        
            For lngMassTagIndex = 0 To .MassTagCount - 1
                With .MassTags(lngMassTagIndex).Location
                    If ParseTrypticName(.TrypticFragmentName, intTrypticMassTagStartNum, intTrypticMassTagsInARow) Then
                        ' We're only recording tryptic hits that were not several tryptic mass tags in a row
                        If intTrypticMassTagsInARow = 1 Then
                            If intTrypticMassTagStartNum <= intTrypticFragmentCount Then
                                blnTrypticMassTagsPresent(intTrypticMassTagStartNum) = True
                            Else
                                ' This shouldn't happen
                                Debug.Assert False
                            End If
                        End If
                    End If
                End With
            Next lngMassTagIndex
            
            ' Determine how many mass tags will be added for this ORF
            intORFMassTagCount = .MassTagCount
            For intTrypticMassTagIndex = 1 To intTrypticFragmentCount
                If Not blnTrypticMassTagsPresent(intTrypticMassTagIndex) Then
                    intORFMassTagCount = intORFMassTagCount + 1
                End If
            Next intTrypticMassTagIndex
            
            ReDim Preserve .MassTags(intORFMassTagCount)
            
            lngSearchStartLoc = 1
            For intTrypticMassTagIndex = 1 To intTrypticFragmentCount
                ' Using the GetTrypticPeptideNext function to retrieve the sequence for each tryptic peptide
                '   is faster than using the GetTrypticPeptideByFragmentNumber function for just those Tryptic Mass Tags that are missing
                strPeptideSequence = objMwtWin.Peptide.GetTrypticPeptideNext(udtGelORF.Orfs(lngORFIndex).Sequence, lngSearchStartLoc, lngResidueStartLoc, lngResidueEndLoc)
                lngSearchStartLoc = lngResidueEndLoc + 1
                
                If Not blnTrypticMassTagsPresent(intTrypticMassTagIndex) Then
                    
                    ' Confirm every 25th peptide to make sure the GetTrypticPeptideNext and the GetTrypticPeptideByFragmentNumber functions agree
                    If intTrypticMassTagIndex Mod 25 = 0 Then
                        Debug.Assert strPeptideSequence = objMwtWin.Peptide.GetTrypticPeptideByFragmentNumber(udtGelORF.Orfs(lngORFIndex).Sequence, intTrypticMassTagIndex, lngResidueStartLoc, lngResidueEndLoc)
                    End If
                    
                    ' Only add a new theoretical mass tag if at least lngMinimumResidueCountToUse residues long and
                    ' the mass is < dblMaximumPeptideMassToUse Da
                    If Len(strPeptideSequence) >= lngMinimumResidueCountToUse Then
                        objMwtWin.Peptide.SetSequence strPeptideSequence, ntgHydrogen, ctgHydroxyl, False, False
                        dblPeptideMass = objMwtWin.Peptide.GetPeptideMass
                        
                        If dblPeptideMass <= dblMaximumPeptideMassToUse Then
                            ' Add a new, theoretical mass tag
                            .MassTagCount = .MassTagCount + 1
                            lngTheoreticalMassTagsAdded = lngTheoreticalMassTagsAdded + 1
                           
                            With .MassTags(.MassTagCount - 1)
                                .MassTagRefID = CLng("-1" & Format(lngORFIndex + 1, strFormatString) & Format(intTrypticMassTagIndex, "000"))
                                .Location.TrypticFragmentName = "t" & Trim(intTrypticMassTagIndex) & ".1"
                                .Location.ResidueStart = lngResidueStartLoc
                                .Location.ResidueEnd = lngResidueEndLoc
                                
                                .IsTheoretical = True
                                .IsLocker = False
                                .IsAMT = False
                                .GANET = LookupPredictedGANET(strPeptideSequence)
                                .Mass = dblPeptideMass
                            End With
                           
                            If blnOutputToFile Then
                                Print #OutSeqFileNum, strPeptideSequence
                                Print #OutRefFileNum, Trim(lngORFIndex) & vbTab & Trim(.MassTagCount - 1) & vbTab & strPeptideSequence
                            End If
                        End If
                    End If
                End If
            Next intTrypticMassTagIndex
        End With
        
        If lngORFIndex Mod 5 = 0 Then
            frmProgress.UpdateProgressBar lngORFIndex
            frmProgress.UpdateCurrentSubTask LongToStringWithCommas(Round(lngTheoreticalMassTagsAdded / 100#, 0) * 100) & " theoretical tryptic mass tags added"
            If KeyPressAbortProcess > 1 Then
                AddToAnalysisHistory lngGelIndex, "User prematurely aborted addition of theoretical tryptic mass tags"
                Exit For
            End If
        End If
    Next lngORFIndex
    
    With GelORFMassTags(lngGelIndex).Definition
        .TheoreticalTrypticMassTagsSuccessfullyAdded = False
        
        If lngTheoreticalMassTagsAdded > 0 Then
            AddToAnalysisHistory lngGelIndex, "Theoretical tryptic mass tags added to the ORF Mass tag list; Total Added = " & lngTheoreticalMassTagsAdded
            .IncludesTheoreticalTrypticMassTags = True
            If KeyPressAbortProcess <= 1 Then
                .TheoreticalTrypticMassTagsSuccessfullyAdded = True
            End If
        Else
            .IncludesTheoreticalTrypticMassTags = False
        End If
    End With
    
    frmProgress.HideForm
    If blnOutputToFile Then
        Close #OutSeqFileNum
        Close #OutRefFileNum
    End If
    
    Exit Sub

ComputeTheoreticalTrypticMassTagsOpenFileError:
    Close
    intOpenFileFailCount = intOpenFileFailCount + 1
    If intOpenFileFailCount = 1 Then
        strOutFilePath = AppendToPath(App.Path, "Temp_ComputeNETValues.NET")
        Resume ComputeTheoreticalTrypticMassTagsOpenFiles
    Else
        blnOutputToFile = False
        Resume ComputeTheoreticalTrypticMassTagsWork
    End If
       
ComputeTheoreticalTrypticMassTagsErrorHandler:
    If Err.Number = -2147024770 Or Err.Number = 429 Then
        MsgBox "Error connecting to MwtWinDll.Dll; you probably need to re-install this application or the Molecular Weight Calculator to properly register the DLL", vbExclamation + vbOKOnly, "Error"
    Else
        MsgBox "An error occurred in sub ComputeTheoreticalTrypticMassTags: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
End Sub

Private Function ParseTrypticName(ByVal strTrypticFragmentName As String, ByRef intTrypticStartNum As Integer, ByRef intTrypticCount As Integer) As Boolean
    ' Examines strTrypticFragmentName to determine the tryptic fragment number and
    '  number of sequential tryptic fragments in a row, returning the value in the ByRef variables
    ' For example, t5.1 means tryptic fragment 5 and only fragment 5
    '              t5.3 means tryptic fragments 5, 6, and 7, so intTrypticStartNum = 5 and intTryptiCount = 3
    '
    ' Returns true if found tx or tx.y (where x and y are numbers)
    ' Returns false if didn't find tx or tx.y
    
    Dim intPeriodLoc As Integer
    
    If Len(strTrypticFragmentName) > 0 Then
        If LCase(Left(strTrypticFragmentName, 1)) = "t" Then
            ' Tryptic peptide
            intPeriodLoc = InStr(strTrypticFragmentName, ".")
            If intPeriodLoc > 0 Then
                ' Multiple tryptic peptides in a row (tryptic peptide with missed cleavages)
                intTrypticStartNum = val(Mid(strTrypticFragmentName, 2, intPeriodLoc - 2))
                intTrypticCount = val(Mid(strTrypticFragmentName, intPeriodLoc + 1))
            Else
                ' Single tryptic peptide
                intTrypticStartNum = val(Mid(strTrypticFragmentName, 2))
                intTrypticCount = 1
            End If
            
            ParseTrypticName = True
        Else
            ' This is unexpected
            ' Only tryptic peptides should have data in strTrypticFragmentName
            Debug.Assert False
            
            ParseTrypticName = False
        End If
    Else
        ParseTrypticName = False
    End If

End Function

Private Function GetAMTRefIDNext(ByVal strRefString As String, ByRef lngStartSearchLoc As Long) As Long
    ' Looks for the next AMTRef in strRefString, returning the AMT RefID value if found
    ' Also increments lngStartSearchLoc to be placed after AMTIDEnd if a match is found
    ' Returns 0 if not found
    
    Dim lngCharLoc As Long, lngAMTIDEndLoc As Long
    Dim lngAMTRefID As Long
    
    lngCharLoc = InStr(lngStartSearchLoc, strRefString, AMTMark)
    
    If lngCharLoc > 0 Then
        lngCharLoc = lngCharLoc + Len(AMTMark)
        If lngCharLoc <= Len(strRefString) Then
            lngAMTIDEndLoc = InStr(lngCharLoc, strRefString, AMTIDEnd)
            If lngAMTIDEndLoc > 0 Then
                lngAMTRefID = val(Mid$(strRefString, lngCharLoc, lngAMTIDEndLoc - lngCharLoc))
            End If
        End If
        
        ' Update lngStartSearchLoc so that the next time this function is called,
        '  the search will continue after this match
        lngStartSearchLoc = lngCharLoc + lngAMTIDEndLoc
    End If
        
    GetAMTRefIDNext = lngAMTRefID

End Function

Public Function GetNumberOfIncludedGels(udtGelDisplayListAndOptions As udtORFViewerGelListType) As Long
    Dim lngGelIndex As Long
    Dim lngGelIncludeCount As Long
    
    For lngGelIndex = 1 To udtGelDisplayListAndOptions.GelCount
        If udtGelDisplayListAndOptions.Gels(lngGelIndex).IncludeGel Then
            lngGelIncludeCount = lngGelIncludeCount + 1
        End If
    Next lngGelIndex

    GetNumberOfIncludedGels = lngGelIncludeCount
End Function

Public Function GetSequencePortion(ByVal strProteinSequence As String, ByVal lngResidueStart As Long, ByVal lngResidueEnd As Long, Optional ByVal blnIncludePrefixAndSuffixResidues As Boolean = True) As String

    Const TERMINII_SYMBOL = "-"
    Const SEPARATION_CHAR = "."
    
    Dim lngSequenceLength As Long
    Dim strSequence As String, strPrefix As String, strSuffix As String
    
    lngSequenceLength = Len(strProteinSequence)
    If lngSequenceLength = 0 Then
        GetSequencePortion = ""
        Exit Function
    End If
    
    If lngResidueStart < 1 Then
        Debug.Assert False
        lngResidueStart = 1
    End If
    
    If lngResidueEnd > Len(strProteinSequence) Then
        Debug.Assert False
        lngResidueEnd = Len(strProteinSequence)
    ElseIf lngResidueEnd < lngResidueStart Then
        Debug.Assert False
        lngResidueEnd = lngResidueStart
    End If
    
    strSequence = Mid(strProteinSequence, lngResidueStart, lngResidueEnd - lngResidueStart + 1)
    
    If blnIncludePrefixAndSuffixResidues Then
        If lngResidueStart <= 1 Then
            strPrefix = TERMINII_SYMBOL
        Else
            strPrefix = Mid(strProteinSequence, lngResidueStart - 1, 1)
        End If
    
        If lngResidueStart > lngSequenceLength Then
            strSuffix = TERMINII_SYMBOL
        Else
            strSuffix = Mid(strProteinSequence, lngResidueEnd + 1, 1)
        End If
    
        strSequence = strPrefix & SEPARATION_CHAR & strSequence & SEPARATION_CHAR & strSuffix
    End If
    
    GetSequencePortion = strSequence
    
End Function

Public Function GetUnusedORFViewerSpotColor(ByRef lngSuggestedUMCColor As Long, Optional lngCurrentGelIndex As Long = -1) As Long
    ' if lngCurrentGelIndex > 0 then ignores specified gel when determining which colors are currently in use
    
    Const COLOR_COUNT = 4
    Dim lngSuggestedColors(COLOR_COUNT, 2) As Long      ' 2-D array to hold colors for both spots and UMC's
    Dim blnSuggestedColorInUse(COLOR_COUNT) As Boolean
    
    lngSuggestedColors(0, 0) = RGB(128, 128, 255): lngSuggestedColors(0, 1) = RGB(0, 0, 255)
    lngSuggestedColors(1, 0) = RGB(255, 128, 128): lngSuggestedColors(1, 1) = RGB(255, 0, 0)
    lngSuggestedColors(2, 0) = RGB(128, 255, 128): lngSuggestedColors(2, 1) = RGB(0, 255, 0)
    lngSuggestedColors(3, 0) = RGB(255, 128, 255): lngSuggestedColors(3, 1) = RGB(255, 0, 255)
    
    Dim lngGelIndex As Long, intColorIndex As Integer
    Dim lngSuggestedColor As Long
    
    With gOrfViewerOptionsCurrentGelList
        If lngCurrentGelIndex > 0 Then
            If .Gels(lngCurrentGelIndex).IonSpotColor = 0 Then
                lngCurrentGelIndex = -1
            End If
        End If
        For lngGelIndex = 1 To .GelCount
            If .Gels(lngGelIndex).IncludeGel Then
                If lngCurrentGelIndex < 1 Or lngGelIndex <> lngCurrentGelIndex Then
                    For intColorIndex = 0 To COLOR_COUNT - 1
                        If lngSuggestedColors(intColorIndex, 0) = .Gels(lngGelIndex).IonSpotColor Then
                            blnSuggestedColorInUse(intColorIndex) = True
                        End If
                    Next intColorIndex
                End If
            End If
        Next lngGelIndex
        
        For intColorIndex = 0 To COLOR_COUNT - 1
            If Not blnSuggestedColorInUse(intColorIndex) Then
                lngSuggestedColor = lngSuggestedColors(intColorIndex, 0)
                lngSuggestedUMCColor = lngSuggestedColors(intColorIndex, 1)
                Exit For
            End If
        Next intColorIndex
    End With
    
    If lngSuggestedColor <= 0 Then
        lngSuggestedColor = lngSuggestedColors(0, 0)
        lngSuggestedUMCColor = lngSuggestedColors(0, 1)
    End If
    
    GetUnusedORFViewerSpotColor = lngSuggestedColor
End Function

Public Sub IncludeTrypticMassTagsInAMTs(lngGelIndex As Long, frmCallingForm As VB.Form)

    Const MEMORY_RESERVE_CHUNK_SIZE As Long = 50000
    Dim CurrSeq As String

    ' The following can be used to include theoretical mass tags in the loaded mass tags
    Dim lngORFIndex As Long, lngMTIndex As Long
    
On Error GoTo IncludeTrypticMassTagsInAMTsErrorHandler

    With GelORFMassTags(lngGelIndex)
        For lngORFIndex = 0 To .ORFCount - 1
            With .Orfs(lngORFIndex)
                For lngMTIndex = 0 To .MassTagCount - 1
                    With .MassTags(lngMTIndex)
                        If .IsTheoretical Then
                            AMTCnt = AMTCnt + 1
                            AMTData(AMTCnt).ID = .MassTagRefID
                            CurrSeq = GetSequencePortion(GelORFData(lngGelIndex).Orfs(lngORFIndex).Sequence, .Location.ResidueStart, .Location.ResidueEnd, False)
                            AMTData(AMTCnt).Sequence = CurrSeq
                            AMTData(AMTCnt).HighNormalizedScore = 0
                            AMTData(AMTCnt).HighDiscriminantScore = 0
                            AMTData(AMTCnt).MW = .Mass
                            AMTData(AMTCnt).NET = .GANET
                            AMTData(AMTCnt).MSMSObsCount = 1
                            AMTData(AMTCnt).NETStDev = 0
                            AMTData(AMTCnt).PNET = .GANET
                            AMTData(AMTCnt).CNT_N = NitrogenCount(CurrSeq)
                            AMTData(AMTCnt).CNT_Cys = AACount(CurrSeq, "C")       'look for cysteine
                        End If
                    End With
                Next lngMTIndex
            End With
            If lngORFIndex Mod 100 = 0 Then frmCallingForm.Caption = "Loading theoretical mass tags : " & lngORFIndex & " / " & .ORFCount
        Next lngORFIndex
    End With

Exit Sub

IncludeTrypticMassTagsInAMTsErrorHandler:

Select Case Err.Number
Case 9                       'need more room for mass tags
    ReDim Preserve AMTData(1 To AMTCnt + MEMORY_RESERVE_CHUNK_SIZE)
    Resume
End Select

Debug.Print "Error in IncludeTrypticMassTagsInAMTs: " & Err.Description
Debug.Assert False


End Sub

Public Sub InitializeGANET(Optional blnInformUserOnLoadFailure As Boolean = True)

On Error GoTo InitializeGANETError
    
    Const TEST_PEPTIDE = "ACDEFGHIKLMNPQRSTVWY"
    Dim sngGANET As Single
    
    Set objGANET = New GANETClass
    
    ' Initialize the GANET prediction code by predicting the NET value of a typical peptide
    sngGANET = objGANET.GetNET(TEST_PEPTIDE)
    
    ' On 2/18/2003, sngGANET = 0.5945874
    ' On 3/23/2004, sngGANET = 0.51
    ' On 6/28/2004, sngGANET = 0.42
    Debug.Assert Round(sngGANET, 2) = 0.43
    
    mGANETLoaded = True
    Exit Sub
    
InitializeGANETError:
    If blnInformUserOnLoadFailure Then
        MsgBox "Unable to initialize the GANET computation class: " & vbCrLf & Err.Description
    End If
    
    mGANETLoaded = False
    
End Sub


' Call LoadORFsFromMTDB() before calling this function
Public Function LoadMassTagsForORFSFromMTDB(ByVal strMTDBConnectionString As String, lngGelIndex As Long, blnLoadPMTs As Boolean, ByRef blnCopiedValuesFromOtherGel As Boolean, Optional blnLoadORFInfoOnly As Boolean = False) As Long
    ' Returns 0 if success; the error number if an error
    ' Note: when blnLoadORFInfoOnly = True, then only loads the ORF info; does not load any AMT's or PMT's
    
    Dim cnnConnection As ADODB.Connection
    Dim rstRecordset As ADODB.Recordset
    Dim strSQL As String
    Dim strSqlSelect As String, strSqlFrom As String, strSqlWhere As String, strSqlOrderBy As String
    Dim strSqlJoin1 As String, strSqlJoin2 As String
    
    Dim lngProgressStep As Long, lngIndex As Long
    Dim lngORFDimCount As Long
    Dim lngORFRefID As Long
    Dim strExpectedORFName As String
    Dim strMassTagLocationInORF As String
    Dim strTrypticName As String
    Dim lngArrayIndex As Long, lngArrayIndexPrevious As Long
    Dim lngError As Long
    Dim lngSourceGelIndex As Long
    Dim lngORFIndex As Long
    Dim blnMatchFound As Boolean
    Dim lngRowCount As Long, lngTotalMassTagCount As Long
    Dim strMessage As String
    
On Error GoTo LoadMassTagsForORFsFromMTDBErrorHandler
    
    If blnLoadPMTs Then
        strMessage = "Obtaining AMT's and PMT's for the loaded ORFs"
    Else
        strMessage = "Obtaining AMT's for the loaded ORFs"
    End If
    
    frmProgress.InitializeForm strMessage, 0, 1, True, False, True, MDIForm1
    
    blnCopiedValuesFromOtherGel = False
    If UBound(GelBody()) > 1 Then
        ' Examine the other Gels to see if any have the needed mass tag information
        ' If one is found, copy the information from that Gel to this one
        For lngSourceGelIndex = 1 To UBound(GelBody())
            If lngSourceGelIndex <> lngGelIndex Then
                With GelORFData(lngSourceGelIndex)
                    If .Definition.MTDBConnectionString = GelORFData(lngGelIndex).Definition.MTDBConnectionString And _
                       .Definition.DataParsedCompletely Then
                        If .ORFCount = GelORFData(lngGelIndex).ORFCount Then
                            If GelORFMassTags(lngSourceGelIndex).Definition.IncludesPMTs = blnLoadPMTs Then
                                frmProgress.UpdateCurrentSubTask "Copying from other Gel file in memory"

                                GelORFMassTags(lngGelIndex) = GelORFMassTags(lngSourceGelIndex)
                                blnMatchFound = True
                            End If
                        End If
                    End If
                End With
                If blnMatchFound Then Exit For
            End If
        Next lngSourceGelIndex
    
        If blnMatchFound Then
            frmProgress.UpdateCurrentSubTask "Parsing data"
            With GelORFMassTags(lngGelIndex)
                For lngORFIndex = 0 To .ORFCount - 1
                    lngTotalMassTagCount = lngTotalMassTagCount + .Orfs(lngORFIndex).MassTagCount
                Next lngORFIndex
            End With
            AddToAnalysisHistory lngGelIndex, "Loaded mass tags for ORFs from other Gel file in memory; FilePath = " & GelData(lngSourceGelIndex).FileName & "); Total Mass Tags Loaded = " & lngTotalMassTagCount & "; ORFs processed = " & GelORFMassTags(lngGelIndex).ORFCount & "; PMTs Loaded = " & GelORFMassTags(lngGelIndex).Definition.IncludesPMTs

            LoadMassTagsForORFSFromMTDB = 0
            blnCopiedValuesFromOtherGel = True
            frmProgress.HideForm
            Exit Function
        End If
    End If
    
    frmProgress.UpdateCurrentSubTask "Connecting to database"
    lngProgressStep = 0

    ' Make sure we're using isotopic masses
    If gMwtWinLoaded Then objMwtWin.SetElementMode emIsotopicMass
    
    ' Before we connect to the database, synchronize GelORFData(lngGelIndex) and GelORFMassTags(lngGelIndex)
    With GelORFMassTags(lngGelIndex)
        ' Synchronize GelORFMassTags(lngGelIndex) with GelORFData(lngGelIndex)
        .ORFCount = GelORFData(lngGelIndex).ORFCount
        
        If GelORFData(lngGelIndex).ORFCount = 0 Then
            ' This is unexpected
            ' I'll reserve space for both
            Debug.Assert False
            
            lngORFDimCount = ORF_DIM_CHUNK
            ReDim GelORFData(lngGelIndex).Orfs(lngORFDimCount)
            ReDim .Orfs(lngORFDimCount)
        Else
            lngORFDimCount = UBound(GelORFData(lngGelIndex).Orfs()) + 1
            ReDim .Orfs(lngORFDimCount)
        End If
        
        For lngIndex = 0 To GelORFData(lngGelIndex).ORFCount - 1
            .Orfs(lngIndex).RefID = GelORFData(lngGelIndex).Orfs(lngIndex).RefID
            .Orfs(lngIndex).MassTagCount = 0
            ReDim .Orfs(lngIndex).MassTags(0)
        Next lngIndex
    End With
    
    If blnLoadORFInfoOnly Then
        GelORFMassTags(lngGelIndex).Definition.DataParsedCompletely = True
        LoadMassTagsForORFSFromMTDB = 0
        Exit Function
    End If
    
    ' Now connect to the database
    If Not EstablishConnection(cnnConnection, strMTDBConnectionString) Then
        LoadMassTagsForORFSFromMTDB = 1
        Exit Function
    End If

    ' Construct the SQL string to grab all of the AMT's, including the ORF they belong to,
    '  their sequence, the AMT's NET value, and modification information
    ' Note also that it's important to sort (ORDER BY) MTORF.Mass_Tag_Name so that the entries are
    ' returned in the same order as was used in sub LoadORFsFromMTDB (they were ordered by Reference)
    
    ' strSQL should be equivalent to the following:
    '
    'SELECT MTORF.Ref_ID, MTORF.Mass_Tag_Name, MTORF.MT_ID, MT.Is_Locker, MT.Is_AMT, MT.Monoisotopic_Mass, MTNET.Avg_GANET, MT.Dynamic_Mod_List, MT.Static_Mod_List, MT.Dyn_Mod_ID, MT.Static_Mod_Id
    'FROM ( dbo_T_Mass_Tags AS MT
    '       LEFT JOIN dbo_T_Mass_Tag_to_ORF_Map AS MTORF ON MT.Mass_Tag_ID = MTORF.MT_ID)
    '        LEFT JOIN dbo_T_Mass_Tags_NET AS MTNET ON MT.Mass_Tag_ID = MTNET.Mass_Tag_ID
    
    'WHERE ((MT.Is_AMT) = 1)
    'ORDER BY MTORF.Mass_Tag_Name, MTORF.Ref_ID;
    
    strSqlSelect = "SELECT MTORF.Ref_ID, MTORF.Mass_Tag_Name, MTORF.MT_ID, MT.Is_Locker, MT.Is_AMT, MT.Monoisotopic_Mass, MT.Peptide, MTNET.Avg_GANET, MT.Dynamic_Mod_List, MT.Static_Mod_List, MT.Dyn_Mod_ID, MT.Static_Mod_Id"
        ' Note the parenthesis is required in the following due to the join
    strSqlFrom = "  FROM (" & TBL_MASS_TAGS & " AS MT"
    strSqlJoin1 = " LEFT JOIN " & TBL_MASS_TAG_TO_ORF_MAP & " AS MTORF ON MT.Mass_Tag_ID = MTORF.MT_ID)"
    strSqlJoin2 = " LEFT JOIN " & TBL_MASS_TAGS_NET & " AS MTNET ON MT.Mass_Tag_ID = MTNET.Mass_Tag_ID"
    
    If Not blnLoadPMTs Then
        strSqlWhere = " WHERE ((MT.Is_AMT) = 1)"
        GelORFMassTags(lngGelIndex).Definition.IncludesPMTs = False
    Else
        ' Load all mass tags, including PMT's
        strSqlWhere = " "
        GelORFMassTags(lngGelIndex).Definition.IncludesPMTs = True
    End If
    
    strSqlOrderBy = " ORDER BY MTORF.Mass_Tag_Name, MTORF.Ref_ID"
    
    strSQL = strSqlSelect & " " & strSqlFrom & " " & strSqlJoin1 & " " & strSqlJoin2 & " " & strSqlWhere & " " & strSqlOrderBy
    
    Set rstRecordset = New ADODB.Recordset

    frmProgress.UpdateCurrentSubTask "Connecting to ORF Reference table in MTDB"
    
    rstRecordset.CursorLocation = adUseClient
    rstRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    rstRecordset.ActiveConnection = Nothing
    
    ' If user requested AMT's only, see if any records were found
    ' If not, inform user, then try to load PMT's instead
    If Not blnLoadPMTs Then
        If rstRecordset.RecordCount = 0 Then
            MsgBox "No AMT's were found in the given database (" & GelORFData(lngGelIndex).Definition.MTDBName & ").  Will load PMT's instead.", vbExclamation + vbOKOnly, "Error"
            rstRecordset.Close
            
            strSqlWhere = " "
            GelORFMassTags(lngGelIndex).Definition.IncludesPMTs = True
            
            strSQL = strSqlSelect & " " & strSqlFrom & " " & strSqlJoin1 & " " & strSqlJoin2 & " " & strSqlWhere & " " & strSqlOrderBy
            
            Set rstRecordset = New ADODB.Recordset
            rstRecordset.CursorLocation = adUseClient
            rstRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
            rstRecordset.ActiveConnection = Nothing
        End If
    End If
    
    With GelORFMassTags(lngGelIndex)
        
        If rstRecordset.RecordCount = 0 Then
            MsgBox "No mass tags were found in the given database (" & GelORFData(lngGelIndex).Definition.MTDBName & ")", vbExclamation + vbOKOnly, "Error"
            .Definition.DataParsedCompletely = True
        Else
            frmProgress.InitializeForm strMessage, 0, rstRecordset.RecordCount, True, False, True, MDIForm1
            frmProgress.UpdateCurrentSubTask "Loading data"
            lngRowCount = 0
            
            .Definition.DataParsedCompletely = False
            
            lngArrayIndex = 0
            lngArrayIndexPrevious = 0
            Do While Not rstRecordset.EOF
                lngORFRefID = FixNullLng(rstRecordset!Ref_ID, -1)
                
                If lngORFRefID >= 0 Then
                    ' Find the record in .Orfs() with lngORFRefID
                    blnMatchFound = False
                    Do
                        If .Orfs(lngArrayIndex).RefID = lngORFRefID Then
                            blnMatchFound = True
                            lngArrayIndexPrevious = lngArrayIndex
                            Exit Do
                        Else
                            lngArrayIndex = lngArrayIndex + 1
                            If lngArrayIndex >= .ORFCount Then
                                lngArrayIndex = 0
                            End If
                            If lngArrayIndex = lngArrayIndexPrevious Then
                                ' Search looped without finding a match
                                Exit Do
                            End If
                        End If
                    Loop
                    
                    If Not blnMatchFound Then
                        ' Need to add a new Orf entry
                        .ORFCount = .ORFCount + 1
                        GelORFData(lngGelIndex).ORFCount = .ORFCount
                        If .ORFCount > lngORFDimCount Then
                            lngORFDimCount = lngORFDimCount + ORF_DIM_CHUNK
                            ReDim Preserve .Orfs(lngORFDimCount)
                            ReDim Preserve GelORFData(lngGelIndex).Orfs(lngORFDimCount)
                        End If
                        lngArrayIndex = .ORFCount - 1
                        .Orfs(lngArrayIndex).RefID = lngORFRefID
                        GelORFData(lngGelIndex).Orfs(lngArrayIndex).RefID = lngORFRefID
                    End If
                    
                    strExpectedORFName = GelORFData(lngGelIndex).Orfs(lngArrayIndex).Reference
                    
                    With .Orfs(lngArrayIndex)
                        .MassTagCount = .MassTagCount + 1
                        ReDim Preserve .MassTags(.MassTagCount)
                        
                        lngTotalMassTagCount = lngTotalMassTagCount + 1
                        
                        With .MassTags(.MassTagCount - 1)
                            .MassTagRefID = FixNullLng(rstRecordset!MT_ID)
                            strMassTagLocationInORF = FixNull(rstRecordset!Mass_Tag_Name)
                            If Len(strMassTagLocationInORF) = 0 Then
                                ' Null value; need to determine location manually
                                lngError = 1
                            Else
                                lngError = ParseMassTagsLocationDescription(strExpectedORFName, FixNull(rstRecordset!Mass_Tag_Name), .Location)
                            End If
                            
                            If lngError <> 0 Then
                                ' Manually determine location of the mass tag in the protein
                                ' Calling .GetTrypticName will fill .ResidueStart and .ResidueEnd properly
                                ' Only record TrypticName in .TrypticFragmentName if it starts with a t (and is thus fully tryptic)
                                If gMwtWinLoaded Then
                                    strTrypticName = objMwtWin.Peptide.GetTrypticName(GelORFData(lngGelIndex).Orfs(lngArrayIndex).Sequence, rstRecordset!Peptide, .Location.ResidueStart, .Location.ResidueEnd)
                                End If
                                If LCase(Left(strTrypticName, 1)) = "t" Then
                                    .Location.TrypticFragmentName = strTrypticName
                                Else
                                    .Location.TrypticFragmentName = ""
                                End If
                                If .Location.ResidueStart = 0 And .Location.ResidueEnd = 0 Then
                                    ' Unable to determine the Tryptic Name
                                    Debug.Print "Unable to determine the tryptic name for Mass Tag " & .MassTagRefID & "; ORF = " & strExpectedORFName & "; Sequence = " & FixNull(rstRecordset!Peptide)
                                End If
                            End If
                            
                            .IsTheoretical = False
                            .IsLocker = FixNullLng(rstRecordset!Is_Locker)
                            .IsAMT = FixNullLng(rstRecordset!Is_AMT)
                            .GANET = FixNullDbl(rstRecordset!Avg_GANET)
                            .Mass = FixNullDbl(rstRecordset!Monoisotopic_Mass)
                            
                            If rstRecordset!Dyn_Mod_ID > 1 Or rstRecordset!Static_Mod_ID > 1 Then
                                .IsModified = True
                                .StaticModList = rstRecordset!Static_Mod_List
                                .DynamicModList = rstRecordset!Dynamic_Mod_List
                            Else
                                .IsModified = False
                            End If
                            
                            ' Check the mass of every 50th mass tag loaded
                            ' Note that the masses will not agree if the mass tag is modified (e.g. N15, PEO or ICAT)
                            If GelORFMassTags(lngGelIndex).Orfs(lngArrayIndex).MassTagCount Mod 50 = 0 Then
                                If gMwtWinLoaded Then
                                    objMwtWin.Peptide.SetSequence rstRecordset!Peptide, ntgHydrogen, ctgHydroxyl, False
                                    If Round(.Mass, 3) <> Round(objMwtWin.Peptide.GetPeptideMass, 3) Then
                                        If Round(.Mass, 2) <> Round(objMwtWin.Peptide.GetPeptideMass, 2) Then
                                            ' Assert that the mass tag is modified; if it isn't, and the masses don't agree, we probably have a problem
                                            Debug.Assert .IsModified
                                        End If
                                    End If
                                End If
                            End If
                        End With
                    End With
                    
                    lngRowCount = lngRowCount + 1
                    If lngRowCount Mod 10 = 0 Then
                        frmProgress.UpdateProgressBar lngRowCount
                        If KeyPressAbortProcess > 1 Then
                            AddToAnalysisHistory lngGelIndex, "User prematurely aborted load of mass tags for ORFs from database"
                            Exit Do
                        End If
                    End If
                Else
                    ' Found a null value for Ref_ID
                    ' This is unexpected
                    Debug.Print "Null value for Ref_ID found; Peptide sequence is " & FixNull(rstRecordset!Peptide)
                End If
                rstRecordset.MoveNext
            Loop
            
            If KeyPressAbortProcess <= 1 Then
                .Definition.DataParsedCompletely = True
            Else
                .Definition.DataParsedCompletely = False
            End If
        End If
    
    End With
    
    rstRecordset.Close
    
    AddToAnalysisHistory lngGelIndex, "Loaded mass tags for ORFs from database (" & GelORFData(lngGelIndex).Definition.MTDBName & "); Total Mass Tags Loaded = " & lngTotalMassTagCount & "; ORFs processed = " & GelORFMassTags(lngGelIndex).ORFCount & "; PMTs Loaded = " & GelORFMassTags(lngGelIndex).Definition.IncludesPMTs
    
    LoadMassTagsForORFSFromMTDB = 0
    
LoadMassTagsForORFsFromMTDBCleanUp:
    On Error Resume Next
    If rstRecordset.STATE <> adStateClosed Then rstRecordset.Close
    If cnnConnection.STATE <> adStateClosed Then cnnConnection.Close
    Set rstRecordset = Nothing
    Set cnnConnection = Nothing
    
    frmProgress.HideForm
    
    Exit Function

LoadMassTagsForORFsFromMTDBErrorHandler:
    If Err.Number = -2147024770 Or Err.Number = 429 Then
        MsgBox "Error connecting to MwtWinDll.Dll; you probably need to re-install this application or the Molecular Weight Calculator to properly register the DLL", vbExclamation + vbOKOnly, "Error"
    Else
        MsgBox "Error while obtaining the AMT's for the loaded ORFs: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    LoadMassTagsForORFSFromMTDB = AssureNonZero(Err.Number)
    
    lngORFDimCount = ORF_DIM_CHUNK
    If GelORFData(lngGelIndex).ORFCount > lngORFDimCount Then
        lngORFDimCount = GelORFData(lngGelIndex).ORFCount
    End If
    
    ReDim GelORFData(lngGelIndex).Orfs(lngORFDimCount)
    ReDim GelORFMassTags(lngGelIndex).Orfs(lngORFDimCount)
    
    Resume LoadMassTagsForORFsFromMTDBCleanUp
    
End Function

' This function should be called before calling LoadMassTagsForORFSFromMTDB()
Public Function LoadORFsFromMTDB(ByVal strMTDBConnectionString As String, lngGelIndex As Long, ByRef blnCopiedValuesFromOtherGel As Boolean) As Long
    ' Returns 0 if success; the error number if an error
    
    Const PROGRESS_STEP_COUNT = 2
    
    Dim cnnConnection As ADODB.Connection
    Dim rstRecordset As ADODB.Recordset
    Dim strSQL As String
    Dim strSqlSelect As String, strSqlFrom As String, strSqlWhere As String, strSqlOrderBy As String
    Dim strMessage As String
    Dim lngExtendedORFDataCount As Long
    
    Dim lngProgressStep As Long
    Dim lngORFDimCount As Long
    Dim lngORFTableRowIndex As Long, lngORFArrayPointer As Long, lngORFArrayPointerPrevious As Long
    Dim lngIndex As Long
    Dim lngORFIndex As Long
    Dim lngSourceGelIndex As Long
    Dim blnMatchFound As Boolean
    Dim lngORFID As Long
    Dim intCharLoc As Integer
    
On Error GoTo LoadORFsFromMTDBErrorHandler
    
    blnCopiedValuesFromOtherGel = False
    If UBound(GelBody()) > 1 Then
        ' Examine the other Gels to see if any have the needed ORF information
        ' If one is found, copy the information from that Gel to this one
        For lngSourceGelIndex = 1 To UBound(GelBody())
            If lngSourceGelIndex <> lngGelIndex Then
                With GelORFData(lngSourceGelIndex)
                    If .ORFCount > 0 And _
                       .Definition.MTDBConnectionString = strMTDBConnectionString And _
                       .Definition.DataParsedCompletely Then
                        frmProgress.UpdateCurrentSubTask "Copying from other Gel file in memory"

                        GelORFData(lngGelIndex) = GelORFData(lngSourceGelIndex)
                        blnMatchFound = True
                    End If
                End With
                If blnMatchFound Then Exit For
            End If
        Next lngSourceGelIndex
    
        If blnMatchFound Then
            frmProgress.UpdateCurrentSubTask "Parsing data"

            ' Initialize the Ion Match values
            With GelORFData(lngGelIndex)
                For lngORFIndex = 0 To .ORFCount - 1
                    With .Orfs(lngORFIndex)
                        .IonMatchCount = 0
                        ReDim .IonMatches(1)
                        .UMCMatchCount = 0
                        ReDim .UMCMatches(1)
                    End With
                Next lngORFIndex
            End With
            
            AddToAnalysisHistory lngGelIndex, "Loaded ORFs from other Gel file in memory; FilePath = " & GelData(lngSourceGelIndex).FileName & "Database = (" & GelORFData(lngGelIndex).Definition.MTDBName & "); Total ORFs Loaded = " & GelORFData(lngGelIndex).ORFCount

            LoadORFsFromMTDB = 0
            blnCopiedValuesFromOtherGel = True
            frmProgress.HideForm
            Exit Function
        End If
    End If
    
    frmProgress.InitializeForm "Obtaining ORF data", 0, PROGRESS_STEP_COUNT, True, True, True, MDIForm1
    frmProgress.InitializeSubtask "Connecting to database", 0, 1
    lngProgressStep = 0
    
    If Not EstablishConnection(cnnConnection, strMTDBConnectionString) Then
        LoadORFsFromMTDB = 1
        Exit Function
    End If
    
    strSqlSelect = "SELECT *"
    strSqlFrom = "  FROM " & TBL_ORF_REFERENCE
    strSqlWhere = ""
    strSqlOrderBy = " ORDER BY Reference"
    
    strSQL = strSqlSelect & " " & strSqlFrom & " " & strSqlWhere & " " & strSqlOrderBy
    
    Set rstRecordset = New ADODB.Recordset

    frmProgress.InitializeSubtask "Connecting to ORF Reference table in MTDB", 0, 1
    
    rstRecordset.CursorLocation = adUseClient
    rstRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    rstRecordset.ActiveConnection = Nothing
    
    With GelORFData(lngGelIndex)
        .ORFCount = 0
        
        lngORFDimCount = ORF_DIM_CHUNK
        ReDim Preserve .Orfs(lngORFDimCount)
        
        With .Definition
            .DateDataObtained = CStr(Format(Now(), "mm/dd/yyyy Hh:Nn:Ss AM/PM"))
            .MTDBConnectionString = strMTDBConnectionString
            .MTDBName = ExtractDBNameFromConnectionString(.MTDBConnectionString)
        End With
    
        If rstRecordset.RecordCount = 0 Then
            MsgBox "No ORF's were found in the given database (" & .Definition.MTDBName & ")", vbExclamation + vbOKOnly, "Error"
            .Definition.DataParsedCompletely = True
        Else
            
            frmProgress.InitializeSubtask "Loading data", 0, rstRecordset.RecordCount
            
            .Definition.DataParsedCompletely = False
            
            Do While Not rstRecordset.EOF
                With .Orfs(.ORFCount)
                    .Reference = FixNull(rstRecordset!Reference)
                    .RefID = FixNullLng(rstRecordset!Ref_ID)
                    .ORFID = FixNullLng(rstRecordset!ORF_ID)
                    .Sequence = FixNull(rstRecordset!Protein_Sequence)
                    .MassMonoisotopic = FixNullDbl(rstRecordset!Monoisotopic_Mass)
                End With
                .ORFCount = .ORFCount + 1
                If .ORFCount > lngORFDimCount Then
                    lngORFDimCount = lngORFDimCount + ORF_DIM_CHUNK
                    ReDim Preserve .Orfs(lngORFDimCount)
                End If
            
                If .ORFCount Mod 10 = 0 Then
                    frmProgress.UpdateSubtaskProgressBar .ORFCount
                    If KeyPressAbortProcess > 1 Then
                        AddToAnalysisHistory lngGelIndex, "User prematurely aborted load of ORFs from database"
                        Exit Do
                    End If
                End If
                rstRecordset.MoveNext
            
            Loop
            
            If KeyPressAbortProcess <= 1 Then
                .Definition.DataParsedCompletely = True
            Else
                .Definition.DataParsedCompletely = False
            End If
        End If
    End With
    
    rstRecordset.Close
    
    AddToAnalysisHistory lngGelIndex, "Loaded ORFs from database (" & GelORFData(lngGelIndex).Definition.MTDBName & "); Total ORFs Loaded = " & GelORFData(lngGelIndex).ORFCount
    
    ' Next determine the name of the ORF database
    frmProgress.InitializeSubtask "Determining ORF database name", 0, 1
    
    strSqlSelect = "SELECT " & FIELD_ORF_DB_NAME & ", " & FIELD_EFFECTIVE_DATE
    strSqlFrom = "  FROM " & TBL_EXTERNAL_DATABASES
    strSqlWhere = ""
    strSqlOrderBy = "ORDER BY Effective_Date DESC"           ' Use Descending to sort Descending
    
    strSQL = strSqlSelect & " " & strSqlFrom & " " & strSqlWhere & " " & strSqlOrderBy

    rstRecordset.CursorLocation = adUseClient
    rstRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    rstRecordset.ActiveConnection = Nothing
    
    If rstRecordset.RecordCount = 0 Then
        strMessage = "The " & TBL_EXTERNAL_DATABASES & " Table is missing from the database (" & GelORFData(lngGelIndex).Definition.MTDBName & "). Unable to retrieve extended ORF info."
        MsgBox strMessage, vbInformation + vbOKOnly, "Missing table"
        AddToAnalysisHistory lngGelIndex, strMessage
    Else
    
        GelORFData(lngGelIndex).Definition.ORFDBName = FixNull(rstRecordset!ORF_DB_Name)
        
        If Len(GelORFData(lngGelIndex).Definition.ORFDBName) = 0 Then
            MsgBox "The ORF database name could not be found in table " & TBL_EXTERNAL_DATABASES & " in database " & GelORFData(lngGelIndex).Definition.MTDBName & ". Unable to retrieve extended ORF info.", vbInformation + vbOKOnly, "Missing data"
        Else
            rstRecordset.Close
            cnnConnection.Close
            Set cnnConnection = Nothing
            
            ' Copy the connection string from the MTDB connection string, but replace the MTDB table with the ORFDB table
            With GelORFData(lngGelIndex).Definition
                intCharLoc = InStr(.MTDBConnectionString, .MTDBName)
                
                If intCharLoc = 0 Then
                    ' Problem
                    .ORFDBConnectionString = ""
                Else
                    .ORFDBConnectionString = Left(.MTDBConnectionString, intCharLoc - 1) & .ORFDBName & Mid(.MTDBConnectionString, intCharLoc + Len(.MTDBName))
                End If
            End With
            
            If Len(GelORFData(lngGelIndex).Definition.ORFDBConnectionString) = 0 Then
                MsgBox "Unable to construct the ORF DB connection string, since unable to parse the MTDB connection string:" & vbCrLf & GelORFData(lngGelIndex).Definition.MTDBConnectionString & vbCrLf & "Thus, unable to retrieve extended ORF info.", vbInformation + vbOKOnly, "Error"
                GelORFData(lngGelIndex).Definition.DataParsedCompletely = False
            Else
                If GelORFData(lngGelIndex).Definition.ORFDBName = "(pending)" Then
                    strMessage = "The ORF database name is listed as '(pending)'. Therefore, unable to retrieve extended ORF info."
                    MsgBox strMessage, vbInformation + vbOKOnly, "Missing Data"
                    AddToAnalysisHistory lngGelIndex, strMessage
                    GelORFData(lngGelIndex).Definition.DataParsedCompletely = True
                Else
    
                    ' Finally, connect to the ORF database and obtain additional information about the ORF's
                    lngProgressStep = lngProgressStep + 1
                    frmProgress.UpdateProgressBar lngProgressStep
                    frmProgress.InitializeSubtask "Connecting to ORF database", 0, 1
                    
                    If Not EstablishConnection(cnnConnection, GelORFData(lngGelIndex).Definition.ORFDBConnectionString) Then
                        GoTo LoadORFsFromMTDBCleanup
                    End If
                    
                    frmProgress.InitializeSubtask "Connecting to ORF table in ORF database", 0, 1
                    
                    strSqlSelect = "SELECT ORF_ID, Reference, Description_From_FASTA, Location_Start, Location_Stop, Strand, Reading_Frame, Isoelectric_Point, CAI, Monoisotopic_Mass, Average_Mass, Molecular_Formula"
                    strSqlFrom = "  FROM " & TBL_ORF_DETAILS
                    strSqlWhere = ""
                    strSqlOrderBy = " ORDER BY Reference"
                    
                    strSQL = strSqlSelect & " " & strSqlFrom & " " & strSqlWhere & " " & strSqlOrderBy
                    
                    rstRecordset.CursorLocation = adUseClient
                    rstRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
                    rstRecordset.ActiveConnection = Nothing
                    
                    With GelORFData(lngGelIndex)
                    
                        If rstRecordset.RecordCount = 0 Then
                            MsgBox "No data was found in the ORF table (" & TBL_ORF_DETAILS & " ) in the ORF database (" & .Definition.ORFDBName & ").  This is unexpected.", vbInformation + vbOKOnly, "Error"
                            .Definition.DataParsedCompletely = True
                        Else
                            ' The data in this table should be a one-to-one match to the data in .Orfs()
                            ' In case it's not, I'll use a Do-Loop with an index pointer
                            
                            frmProgress.InitializeSubtask "Obtaining additional ORF data", 0, rstRecordset.RecordCount
                            lngORFTableRowIndex = 0
                            lngORFArrayPointer = 0
                            lngORFArrayPointerPrevious = 0
                            .Definition.DataParsedCompletely = False
                            Do While Not rstRecordset.EOF
                            
                                lngORFID = FixNullLng(rstRecordset!ORF_ID)
                                
                                ' Find the item in .Orfs() with this ORFID
                                blnMatchFound = False
                                Do
                                    If lngORFID = .Orfs(lngORFArrayPointer).ORFID Then
                                        blnMatchFound = True
                                        lngORFArrayPointerPrevious = lngORFArrayPointer
                                        Exit Do
                                    Else
                                        lngORFArrayPointer = lngORFArrayPointer + 1
                                        If lngORFArrayPointer >= .ORFCount Then
                                            lngORFArrayPointer = 0
                                        End If
                                        If lngORFArrayPointer = lngORFArrayPointerPrevious Then
                                            ' Search looped without finding a match
                                            Exit Do
                                        End If
                                    End If
                                Loop
                                
                                If blnMatchFound Then
                                    lngExtendedORFDataCount = lngExtendedORFDataCount + 1
                                    With .Orfs(lngORFArrayPointer)
                                        ' The following was previously obtained, and should match that in this database
                                        Debug.Assert .Reference = FixNull(rstRecordset!Reference)
                                        Debug.Assert Round(.MassMonoisotopic, 4) = Round(FixNullDbl(rstRecordset!Monoisotopic_Mass), 4)
                                        
                                        .Description = FixNull(rstRecordset!Description_From_FASTA)
                                        .LocationStart = FixNullLng(rstRecordset!Location_Start)
                                        .LocationStop = FixNullLng(rstRecordset!Location_Stop)
                                        .Strand = FixNull(rstRecordset!Strand)
                                        .ReadingFrame = FixNullLng(rstRecordset!Reading_Frame)
                                        .pi = FixNullDbl(rstRecordset!Isoelectric_Point)
                                        .CAI = FixNullDbl(rstRecordset!CAI)
                                        .MassAverage = FixNullDbl(rstRecordset!Average_Mass)
                                        .MolecularFormula = FixNullDbl(rstRecordset!Molecular_Formula)
                                    End With
                                Else
                                    ' No match found; this is unexpected
                                    Debug.Print "Caution: No match found in .ORFS() for ORFID = " & lngORFID
                                End If
                                
                                lngORFTableRowIndex = lngORFTableRowIndex + 1
                                If lngORFTableRowIndex Mod 10 = 0 Then
                                    frmProgress.UpdateSubtaskProgressBar lngORFTableRowIndex
                                    If KeyPressAbortProcess > 1 Then
                                        AddToAnalysisHistory lngGelIndex, "User prematurely aborted load of extended ORF info from database"
                                        Exit Do
                                    End If
                                End If
                                rstRecordset.MoveNext
                            Loop
                            If KeyPressAbortProcess <= 1 Then
                                .Definition.DataParsedCompletely = True
                            Else
                                .Definition.DataParsedCompletely = False
                            End If
                        End If
                    End With
                
                    AddToAnalysisHistory lngGelIndex, "Loaded extended ORF data from database (" & ExtractDBNameFromConnectionString(GelORFData(lngGelIndex).Definition.ORFDBConnectionString) & "); Total ORFs Processed = " & lngExtendedORFDataCount
                
                    If lngExtendedORFDataCount <> GelORFData(lngGelIndex).ORFCount Then
                        ' Didn't find extended info on all of the ORFs
                        ' Find the ones with missing extended data
                        With GelORFData(lngGelIndex)
                            For lngIndex = 0 To .ORFCount - 1
                                If Len(.Orfs(lngIndex).Description) = 0 Then
                                    Debug.Print "No extended info found for ORF " & .Orfs(lngIndex).Reference
                                End If
                            Next lngIndex
                        End With
                    End If
                End If
            End If
        End If
    End If
    
    LoadORFsFromMTDB = 0
    
LoadORFsFromMTDBCleanup:
    On Error Resume Next
    If rstRecordset.STATE <> adStateClosed Then rstRecordset.Close
    If cnnConnection.STATE <> adStateClosed Then cnnConnection.Close
    Set rstRecordset = Nothing
    Set cnnConnection = Nothing
    
    frmProgress.HideForm
    
    LoadORFsFromMTDB = 0
    Exit Function

LoadORFsFromMTDBErrorHandler:
    MsgBox "Error while obtaining ORF data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    LoadORFsFromMTDB = AssureNonZero(Err.Number)
    Resume LoadORFsFromMTDBCleanup
    
End Function

Public Function LookupPredictedGANET(strSequence As String) As Single
    Dim strTrimmedSequence As String
    
    ' Provided mGANETLoaded = True, calls the .GetNET function in the objGANET class
    '  to get the predicted GANET for the sequence
    
    If mGANETLoaded Then
        strTrimmedSequence = TrimSequence(strSequence)
        LookupPredictedGANET = objGANET.GetNET(strTrimmedSequence)
    Else
        LookupPredictedGANET = 0.5
    End If
End Function

Public Sub InitializeGelDisplayOptions(ByRef udtGelDisplayListAndOptions As udtORFViewerGelListType, ByVal lngGelIndex As Long, blnDeleted As Boolean)
    
    If lngGelIndex > UBound(udtGelDisplayListAndOptions.Gels()) Then
        Debug.Assert False
        Exit Sub
    End If
    
    ' Note: gOrfViewerOptionsCurrentGelList.Gels() is 1-based, as is udtGelDisplayListAndOptions.Gels()
    gOrfViewerOptionsCurrentGelList = udtGelDisplayListAndOptions
    With udtGelDisplayListAndOptions.Gels(lngGelIndex)
        .Deleted = blnDeleted
        .FileName = ""
        .IsoLines = 0
        .NETAdjustmentType = natGANET
        
        .IonSpotColor = GetUnusedORFViewerSpotColor(.UMCSpotColor, lngGelIndex)
        .IonSpotColorSelected = RGB(0, 128, 0)
        .IonSpotShape = sCircle
        
        .UMCSpotColorSelected = .IonSpotColorSelected
        .UMCSpotShape = sTriangleWithExtents
    End With
End Sub

Public Function InitializeGelDisplayListAndOptions(ByRef udtGelDisplayListAndOptionsToUpdate As udtORFViewerGelListType, ByVal lngGelIndexAssureIncluded As Long) As Boolean
    ' Updates udtGelDisplayListAndOptionsToUpdate, returning true if necessary to populate the ORFGroupList
    ' if lngGelIndexAssureIncluded >=0 then sets .IncludeGel to True for gel with index lngGelIndexAssureIncluded
        
    Dim intLoadedGelCount As Integer
    Dim lngGelIndex As Long
    Dim blnUpdateORFGroupList As Boolean
    
    intLoadedGelCount = UBound(GelBody())
    
    With udtGelDisplayListAndOptionsToUpdate
        If .GelCount < intLoadedGelCount Then
            ReDim Preserve .Gels(intLoadedGelCount)
            
            ' For the newly reserved memory, set .Deleted to True (necessary for the logic below)
            For lngGelIndex = .GelCount + 1 To intLoadedGelCount
                InitializeGelDisplayOptions udtGelDisplayListAndOptionsToUpdate, lngGelIndex, True
            Next lngGelIndex
        End If
        
        If intLoadedGelCount = 0 Then
            If .GelCount > 0 Then
                ' Set .Deleted to True
                For lngGelIndex = 1 To .GelCount
                    .Gels(lngGelIndex).Deleted = True
                Next lngGelIndex
                
                blnUpdateORFGroupList = True
            End If
        Else
            .GelCount = intLoadedGelCount
        
            ' Make sure the data in udtGelDisplayListAndOptions is synchronized with the currently loaded gels
            For lngGelIndex = 1 To .GelCount
                With .Gels(lngGelIndex)
                    If GelStatus(lngGelIndex).Deleted Then
                        If .Deleted = False Then
                            .Deleted = True
                            .IncludeGel = False
                            blnUpdateORFGroupList = True
                        End If
                    Else
                        If .Deleted = True Then
                            InitializeGelDisplayOptions udtGelDisplayListAndOptionsToUpdate, lngGelIndex, False
                            blnUpdateORFGroupList = True
                        End If
                    End If
                    
                    If Not .Deleted Then
                        If .FileName <> GelData(lngGelIndex).FileName Or _
                           .IsoLines <> GelData(lngGelIndex).IsoLines Then
                           
                            .FileName = GelData(lngGelIndex).FileName
                            .IsoLines = GelData(lngGelIndex).IsoLines
                            
                            If lngGelIndexAssureIncluded >= 0 Then
                                .IncludeGel = (lngGelIndexAssureIncluded = lngGelIndex)
                            Else
                                .IncludeGel = False
                            End If
                            
                            If Not GelAnalysis(lngGelIndex) Is Nothing Then
                                .MTDBName = ExtractDBNameFromConnectionString(GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString)
                            End If
                            blnUpdateORFGroupList = True
                        
                        End If
                        
                        ' Make sure the currently saved file name is correct
                        .GelFileName = GelBody(lngGelIndex).Caption
                    End If
                    
                End With
            Next lngGelIndex
        End If
    End With
    
    InitializeGelDisplayListAndOptions = blnUpdateORFGroupList
    
End Function

Public Sub ORFViewerOptionsInitializeForm(udtCurrentGelList As udtORFViewerGelListType, lngFormID As Long)
    Dim intIndex As Integer
        
    gOrfViewerOptionsSavedGelList = udtCurrentGelList
    gOrfViewerOptionsCurrentGelList = udtCurrentGelList
    
    frmORFViewerOptions.InitializeGeneralOptions
    frmORFViewerOptions.SetGelInListChanged False
    
    With gOrfViewerOptionsCurrentGelList
        frmORFViewerOptions.lstAvailableGels.Clear
        
        For intIndex = 1 To .GelCount
            With .Gels(intIndex)
                If .Deleted = False Then
                    frmORFViewerOptions.lstAvailableGels.AddItem Trim(intIndex) & ": " & StripFullPath(.GelFileName)
                End If
            End With
        Next intIndex
    End With
    
    frmORFViewerOptions.UpdateGelInUseList
    frmORFViewerOptions.SetCallingFormID lngFormID
    
    frmORFViewerOptions.show
End Sub

Private Function ParseMassTagsLocationDescription(ByVal strExpectedORFName As String, ByVal strLocationDescription As String, ByRef udtLocation As udtORFPeptideLocationType) As Long
    ' Returns 0 if success; 1 if failure
    
    Dim lngCharLoc As Integer
    Dim strDescription As String
    Dim strResidueStart As String, strResidueStop As String
    
    ' Look for strExpectedORFName in strLocationDescription
    ' If not found, will need to do some extra work
    
    lngCharLoc = InStr(strLocationDescription, strExpectedORFName)
    
    If lngCharLoc > 0 Then
        strDescription = Mid(strLocationDescription, lngCharLoc + Len(strExpectedORFName))
        If Left(strDescription, 1) = "." Then strDescription = Mid(strDescription, 2)
    Else
        ' strExpectedORFName was not found in strLocationDescription
        ' This is unlikely, but is possible
        If Len(strLocationDescription) > 0 Then
            ' Look for a period instead
            lngCharLoc = InStr(strLocationDescription, ".")
                    
            If lngCharLoc > 0 Then
                strDescription = Mid(strLocationDescription, lngCharLoc + 1)
            End If
        End If
    End If
    
    With udtLocation
        If Len(strDescription) > 0 Then
            If LCase(Left(strDescription, 1)) = "t" Then
                ' Tryptic peptide
                .TrypticFragmentName = LCase(strDescription)
                
                ' Need to compute the start and end residues
                ' I'll do this in a separate function, named UpdateORFStatistics

            Else
                .TrypticFragmentName = ""
                lngCharLoc = InStr(strDescription, ".")
                If lngCharLoc > 0 Then
                    strResidueStart = Left(strDescription, lngCharLoc - 1)
                    strResidueStop = Mid(strDescription, lngCharLoc + 1)
                Else
                    ' This is unexpected
                    Debug.Assert False
                    strResidueStart = strDescription
                End If
                    
                If IsNumeric(strResidueStart) Then
                    .ResidueStart = val(strResidueStart)
                Else
                    .ResidueStart = 0
                End If
            
                If IsNumeric(strResidueStop) Then
                    ' Note: Need to subtract one here since the value recorded in
                    '       strLocationDescription is actually the residue number after the last residue in the peptide
                    '       For example, DR0001.1.5 means the first four residues, thus I record a 1 and a 4
                    ' Also note that the function objMwtWin.Peptide.GetTrypticName()
                    '  actually returns the start and stop residue, and is thus of the form 1.4 for the above example
                    .ResidueEnd = val(strResidueStop) - 1
                Else
                    .ResidueEnd = 0
                End If
            
            End If
            
            ParseMassTagsLocationDescription = 0
        Else
            .ResidueStart = 0
            .ResidueEnd = 0
            .TrypticFragmentName = ""
            
            ParseMassTagsLocationDescription = 1
        End If
    End With
    
End Function

Public Sub RecordIonMatchesInORFMassTags(ByVal lngGelIndex As Long)
    ' Examine the Ions and UMC's for this gel and record the AMT matches in GelORFMassTags()
    
    ' To speed up the searching, uses a sorted list of mass tag RefID's (stored in udtMassTagRefIDIndex),
    '  searching them using a non-recursive Binary Search
    Dim udtMassTagRefIDIndex As udtMassTagRefIDIndexType
    
    Dim lngIonIndex As Long, lngUMCIndex As Long, lngClassMemberIndex As Long
    Dim lngCompareIndex As Long
    Dim lngStartSearchLoc As Long
    Dim lngMassTagCount As Long
    Dim lngAMTRefID As Long
    Dim lngMatchingIndex As Long
    Dim lngORFIndex As Long, lngMassTagIndexPointer As Long
    Dim lngMissingMassTagCount As Long
    Dim lngTotalIonMatchCount As Long
    Dim blnProceed As Boolean
    Dim strMessage As String
    
    ' First populate udtMassTagRefIDIndex with the AMT's for this gel
    UpdateAmtRefIDMassIndexingArray udtMassTagRefIDIndex, lngGelIndex
    lngMassTagCount = udtMassTagRefIDIndex.MassTagCount
    
    ' Now parse the Ion Matches for this gel
    ' Note that the GetAMTRefFromString1() function could be used to extract AMT matches from ions
    ' Instead, for speed purposes, I've copied the necessary code, removed unneeded items,
    '  and created the new function GetAMTRefID

    frmProgress.InitializeForm "Examining AMT matches", 0, 2, True, True, True, MDIForm1
    frmProgress.InitializeSubtask "Gel Ions", 0, GelData(lngGelIndex).IsoLines
    
    ' First clear all of the stored .IonMatch and .UMCMatch values
    With GelORFData(lngGelIndex)
        For lngORFIndex = 0 To .ORFCount - 1
            .Orfs(lngORFIndex).IonMatchCount = 0
            .Orfs(lngORFIndex).UMCMatchCount = 0
            ReDim .Orfs(lngORFIndex).IonMatches(0)
            ReDim .Orfs(lngORFIndex).UMCMatches(0)
        Next lngORFIndex
    End With
    
    ' Note: .IsoData() is 1-based
    For lngIonIndex = 1 To GelData(lngGelIndex).IsoLines
        If Len(GelData(lngGelIndex).IsoData(lngIonIndex).MTID) > 0 Then
            lngStartSearchLoc = 1
            Do
                ' Look for the next Amt RefID (for example, AMT:12345)
                ' GetAMTRefIDNext returns 0 if no match
                lngAMTRefID = GetAMTRefIDNext(GelData(lngGelIndex).IsoData(lngIonIndex).MTID, lngStartSearchLoc)
                If lngAMTRefID > 0 Then
                    ' Look for lngAMTRefID in udtMassTagRefIDIndex.MassTagRefID()
                    lngMatchingIndex = BinarySearchLng(udtMassTagRefIDIndex.MassTagRefID(), lngAMTRefID, 0, lngMassTagCount - 1)
                    
                    If lngMatchingIndex >= 0 Then
                        With udtMassTagRefIDIndex
                            lngMassTagIndexPointer = .MassTagRefIDPointer(lngMatchingIndex)
                            lngORFIndex = .ORFIndex(lngMassTagIndexPointer)
                        End With
                        
                        With GelORFData(lngGelIndex).Orfs(lngORFIndex)
                            ReDim Preserve .IonMatches(.IonMatchCount)
                            With .IonMatches(.IonMatchCount)
                                .MassTagIndex = udtMassTagRefIDIndex.MassTagIndex(lngMassTagIndexPointer)
                                .IonDataIndex = lngIonIndex
                                
                                ' Make sure we really did find the match
                                Debug.Assert GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(.MassTagIndex).MassTagRefID = lngAMTRefID
                            End With
                            .IonMatchCount = .IonMatchCount + 1
                        End With
                        
                        lngTotalIonMatchCount = lngTotalIonMatchCount + 1
                    Else
                        ' Couldn't find Mass Tag
                        lngMissingMassTagCount = lngMissingMassTagCount + 1
                    End If
                Else
                    Exit Do
                End If
            Loop
        End If
        
        If lngIonIndex Mod 50 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngIonIndex
            If KeyPressAbortProcess > 1 Then
                AddToAnalysisHistory lngGelIndex, "User prematurely aborted process of recording ion matches in ORF mass tags"
                Exit For
            End If
        End If
    Next lngIonIndex
    

    ' Now parse the UMC's for this gel
    ' As of Jan 2003, this software does not record ion matches for a UMC, per se
    ' Rather, when searching the database using UMC's, it simply searches the database using
    '  each ion in each UMC.  Thus, I'll iterate through the UMC's, and iterate through the
    '  ions in each UMC
    '
    ' Note: March 2003: The new SearchMT_ConglomerateUMC function does search for ion matches on a UMC by UMC basis
    '                   However, it still stores the results on an ion by ion basis, storing the same hit results for all ions of a UMC
    
    frmProgress.UpdateProgressBar 1
    frmProgress.InitializeSubtask "Gel UMCs", 0, GelUMC(lngGelIndex).UMCCnt
    
    ' Note: .UMCs() is 0-based
    For lngUMCIndex = 0 To GelUMC(lngGelIndex).UMCCnt - 1
        For lngClassMemberIndex = 0 To GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassCount - 1
            If GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassMType(lngClassMemberIndex) = gldtIS Then
                lngIonIndex = GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassMInd(lngClassMemberIndex)
                
                ' Note .IsoData() is 1-based
                If lngIonIndex <= GelData(lngGelIndex).IsoLines Then
                  If Len(GelData(lngGelIndex).IsoData(lngIonIndex).MTID) > 0 Then
                    lngStartSearchLoc = 1
                    Do
                        ' Look for the next Amt RefID (for example, AMT:12345)
                        ' GetAMTRefIDNext returns 0 if no match
                        lngAMTRefID = GetAMTRefIDNext(GelData(lngGelIndex).IsoData(lngIonIndex).MTID, lngStartSearchLoc)
                        If lngAMTRefID > 0 Then
                            ' Look for lngAMTRefID in udtMassTagRefIDIndex udtMassTagRefIDIndex.MassTagRefID()
                            lngMatchingIndex = BinarySearchLng(udtMassTagRefIDIndex.MassTagRefID(), lngAMTRefID, 0, lngMassTagCount - 1)
                            If lngMatchingIndex >= 0 Then
                                With udtMassTagRefIDIndex
                                    lngMassTagIndexPointer = .MassTagRefIDPointer(lngMatchingIndex)
                                    lngORFIndex = .ORFIndex(lngMassTagIndexPointer)
                                End With
                                
                                With GelORFData(lngGelIndex).Orfs(lngORFIndex)
                                    ' Only add this UMC if it isn't already present
                                    blnProceed = True
                                    For lngCompareIndex = 0 To .UMCMatchCount - 1
                                        If .UMCMatches(lngCompareIndex).UMCDataIndex = lngUMCIndex Then
                                            blnProceed = False
                                            Exit For
                                        End If
                                    Next lngCompareIndex
                                    
                                    If blnProceed Then
                                        ReDim Preserve .UMCMatches(.UMCMatchCount)
                                        With .UMCMatches(.UMCMatchCount)
                                            .MassTagIndex = udtMassTagRefIDIndex.MassTagIndex(lngMassTagIndexPointer)
                                            .UMCDataIndex = lngUMCIndex
                                            .ClassMemberIndex = lngClassMemberIndex
                                            
                                            ' Make sure we really did find the match
                                            Debug.Assert GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(.MassTagIndex).MassTagRefID = lngAMTRefID
                                        End With
                                        .UMCMatchCount = .UMCMatchCount + 1
                                    End If
                                End With
                            End If
                            
                        Else
                            Exit Do
                        End If
                    Loop
                 End If
                End If
            End If
        Next lngClassMemberIndex
        
        If lngUMCIndex Mod 50 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngUMCIndex
            If KeyPressAbortProcess > 1 Then
                AddToAnalysisHistory lngGelIndex, "User prematurely aborted process of recording UMC ion matches in ORF mass tags"
                Exit For
            End If
        End If
    Next lngUMCIndex
    
    UpdateValueInStringByKey GelORFData(lngGelIndex).Definition.OtherInfo, UMC_COUNT_LAST_RECORD_ION_MATCH_CALL, Trim(GelUMC(lngGelIndex).UMCCnt)
    
    If lngTotalIonMatchCount = 0 Then
        strMessage = "None of the ions or UMC's has recorded mass tag matches.  Although the ORF viewer will display the Mass Tags and ions or UMC's, none of the ion counts will be correct.  After you have performed a database search against the ions or UMC's, choose 'Refresh Source Data' from the File menu of the ORF Viewer to populate the ion matches."
        MsgBox strMessage, vbExclamation + vbOKOnly, "No ion matches"
    End If
    
    If lngMissingMassTagCount > 0 Then
        strMessage = Trim(Str(lngMissingMassTagCount)) & " ions contain mass tag matches to mass tags not located in the currently linked mass tag database (" & Trim(lngTotalIonMatchCount + lngMissingMassTagCount) & " total matches in memory).  These could easily be matches to PMT's if you did not include PMT's when loading the ORF's.  Otherwise, this may indicate that the gel is not properly linked with the database.  If necessary, unlink, then re-link to the MTDB, re-perform the ion searching, then choose 'Reload ORFs from MTDB' from the File menu of the ORF viewer.  To include PMT's, choose 'Set Included Analyses and Options' from the View menu, then check 'Load PMT's in addition to AMT's', then choose 'Reload ORFs from MTDB' from the File menu of the ORF viewer."
        MsgBox strMessage, vbInformation + vbOKOnly, "Missing mass tags"
    End If
    
    frmProgress.HideForm
End Sub

' Unused Function (March 2003)
'''Public Sub SynchronizeORFData(lngSourceGelIndex As Long)
'''    ' Synchronize ORF information among all loaded gells
'''
'''    Dim lngGelIndex As Long
'''    Dim lngORFIndex As Long
'''
'''    If lngSourceGelIndex < 1 Or lngSourceGelIndex > UBound(GelBody()) Then
'''        MsgBox "Invalid source gel index"
'''        Exit Sub
'''    End If
'''
'''    frmProgress.InitializeForm "Synchronizing ORF data", 0, UBound(GelBody()), True, True, True
'''
'''    For lngGelIndex = 1 To UBound(GelBody())
'''        frmProgress.InitializeSubtask TrimFileName(GelBody(lngGelIndex).Caption), 0, GelORFData(lngSourceGelIndex).ORFCount
'''
'''        If lngGelIndex <> lngSourceGelIndex Then
'''            GelORFData(lngGelIndex) = GelORFData(lngSourceGelIndex)
'''            GelORFMassTags(lngGelIndex) = GelORFMassTags(lngSourceGelIndex)
'''
'''            With GelORFData(lngGelIndex)
'''                For lngORFIndex = 0 To .ORFCount - 1
'''                    With .Orfs(lngORFIndex)
'''                        .IonMatchCount = 0
'''                        ReDim .IonMatches(1)
'''                        .UMCMatchCount = 0
'''                        ReDim .UMCMatches(1)
'''                    End With
'''                    If lngORFIndex Mod 100 = 0 Then frmProgress.UpdateSubtaskProgressBar lngORFIndex
'''                Next lngORFIndex
'''            End With
'''
'''            GelStatus(lngGelIndex).Dirty = True
'''        End If
'''    Next lngGelIndex
'''
'''    MsgBox "Synchronized ORF data from gel " & TrimFileName(GelBody(lngSourceGelIndex).Caption) & " to the other gels."
'''    frmProgress.HideForm
'''
'''End Sub

Public Function TrimSequence(ByVal strSequence As String) As String
    ' Looks for, and removes . or - symbols from the first or second, and
    ' last or second-to-the-last positions in strSequence
    ' For example, if strSequence = "A.BCDEFG.H"
    '  then this function will return BCDEFG
    ' Also, if strSequence = "-BCDEFG."
    '  then this function will return BCDEFG
    
    Dim strTrimmedSequence As String
    Dim lngSequenceLength As Long
    
    strTrimmedSequence = Trim(strSequence)
    If Not IsCharacter(Mid(strTrimmedSequence, 2, 1)) Then
        strTrimmedSequence = Mid(strTrimmedSequence, 3)
    End If
    
    lngSequenceLength = Len(strTrimmedSequence)
    If Not IsCharacter(Mid(strTrimmedSequence, lngSequenceLength - 1, 1)) Then
        strTrimmedSequence = Left(strTrimmedSequence, lngSequenceLength - 2)
    End If
    
    If Not IsCharacter(Left(strTrimmedSequence, 1)) Then
        strTrimmedSequence = Mid(strTrimmedSequence, 2)
    End If
    
    If Not IsCharacter(Right(strTrimmedSequence, 1)) Then
        strTrimmedSequence = Left(strTrimmedSequence, Len(strTrimmedSequence) - 1)
    End If

    TrimSequence = strTrimmedSequence

End Function

Private Sub UpdateAmtRefIDMassIndexingArray(udtMassTagRefIDIndex As udtMassTagRefIDIndexType, lngGelIndex As Long)

    Const MASS_TAG_DIM_CHUNK = 500
    
    Dim objQSLong As New QSLong
    
    Dim lngORFIndex As Long
    Dim lngMassTagIndex As Long
    
    Dim lngMassTagDimCount As Long

    frmProgress.InitializeForm "Building AMT reference arrays", 0, GelORFMassTags(lngGelIndex).ORFCount, True, False, True, MDIForm1
    
    With udtMassTagRefIDIndex
        .MassTagCount = 0
        lngMassTagDimCount = MASS_TAG_DIM_CHUNK
        
        ReDim .MassTagRefID(MASS_TAG_DIM_CHUNK)
        ReDim .MassTagRefIDPointer(MASS_TAG_DIM_CHUNK)
        ReDim .ORFIndex(MASS_TAG_DIM_CHUNK)
        ReDim .MassTagIndex(MASS_TAG_DIM_CHUNK)
    End With
    
    If GelORFMassTags(lngGelIndex).ORFCount < 1 Then
        Exit Sub
    End If

    For lngORFIndex = 0 To GelORFMassTags(lngGelIndex).ORFCount - 1
        If GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount > 0 Then
            
            With GelORFMassTags(lngGelIndex).Orfs(lngORFIndex)
                
                For lngMassTagIndex = 0 To .MassTagCount - 1
                    
                    ' Add an entry to udtMassTagRefIDIndex
                    udtMassTagRefIDIndex.MassTagRefID(udtMassTagRefIDIndex.MassTagCount) = .MassTags(lngMassTagIndex).MassTagRefID
                    
                    udtMassTagRefIDIndex.MassTagRefIDPointer(udtMassTagRefIDIndex.MassTagCount) = udtMassTagRefIDIndex.MassTagCount
                    
                    udtMassTagRefIDIndex.ORFIndex(udtMassTagRefIDIndex.MassTagCount) = lngORFIndex
                    udtMassTagRefIDIndex.MassTagIndex(udtMassTagRefIDIndex.MassTagCount) = lngMassTagIndex
                    
                    udtMassTagRefIDIndex.MassTagCount = udtMassTagRefIDIndex.MassTagCount + 1
                    If udtMassTagRefIDIndex.MassTagCount >= lngMassTagDimCount Then
                        lngMassTagDimCount = lngMassTagDimCount + MASS_TAG_DIM_CHUNK
                        ReDim Preserve udtMassTagRefIDIndex.MassTagRefID(lngMassTagDimCount)
                        ReDim Preserve udtMassTagRefIDIndex.MassTagRefIDPointer(lngMassTagDimCount)
                        ReDim Preserve udtMassTagRefIDIndex.ORFIndex(lngMassTagDimCount)
                        ReDim Preserve udtMassTagRefIDIndex.MassTagIndex(lngMassTagDimCount)
                    End If
                Next lngMassTagIndex
            End With
        End If
        
        If lngORFIndex Mod 50 = 0 Then
            frmProgress.UpdateProgressBar lngORFIndex
            If KeyPressAbortProcess > 1 Then
                AddToAnalysisHistory lngGelIndex, "User prematurely aborted process of updating the AMT Reference ID mass indexing array"
                Exit For
            End If
        End If
    Next lngORFIndex
    
    ' Sort the .MassTagRefID() array
    ' First need to ReDim the arrays to remove unused memory space
    lngMassTagDimCount = udtMassTagRefIDIndex.MassTagCount
    If lngMassTagDimCount = 0 Then lngMassTagDimCount = 1
    ReDim Preserve udtMassTagRefIDIndex.MassTagRefID(0 To lngMassTagDimCount - 1)
    ReDim Preserve udtMassTagRefIDIndex.MassTagRefIDPointer(0 To lngMassTagDimCount - 1)
    ReDim Preserve udtMassTagRefIDIndex.ORFIndex(0 To lngMassTagDimCount - 1)
    ReDim Preserve udtMassTagRefIDIndex.MassTagIndex(0 To lngMassTagDimCount - 1)
    
    If Not objQSLong.QSAsc(udtMassTagRefIDIndex.MassTagRefID(), udtMassTagRefIDIndex.MassTagRefIDPointer()) Then
        ' Failure with QSort
        Debug.Assert False
    End If
    
    Set objQSLong = Nothing
    
    frmProgress.HideForm
End Sub


Public Sub UpdateORFStatistics(lngGelIndex As Long)
    Dim lngORFIndex As Long, lngMassTagIndex As Long
    Dim intTrypticStartNum  As Integer, intTrypticCount As Integer
    Dim lngResidueStartForCurrent As Long, lngResidueEndForCurrent As Long
    Dim strFragment As String
    
    If GelORFData(lngGelIndex).ORFCount = 0 Then Exit Sub
    
    Debug.Assert GelORFData(lngGelIndex).ORFCount = GelORFMassTags(lngGelIndex).ORFCount
    
    frmProgress.InitializeForm "Parsing AMT's and ORFs", 0, GelORFData(lngGelIndex).ORFCount, True, False, True, MDIForm1
    
    For lngORFIndex = 0 To GelORFData(lngGelIndex).ORFCount - 1
        GelORFData(lngGelIndex).Orfs(lngORFIndex).TrypticFragmentCount = objInSilicoDigest.CountTrypticsInSequence(GelORFData(lngGelIndex).Orfs(lngORFIndex).Sequence)
        
        With GelORFMassTags(lngGelIndex).Orfs(lngORFIndex)
            For lngMassTagIndex = 0 To .MassTagCount - 1
                With .MassTags(lngMassTagIndex).Location
                    If Len(.TrypticFragmentName) > 0 Then
                        If ParseTrypticName(.TrypticFragmentName, intTrypticStartNum, intTrypticCount) And gMwtWinLoaded Then
                            ' Determine the start and end residue number of this mass tag
                            strFragment = objMwtWin.Peptide.GetTrypticPeptideByFragmentNumber(GelORFData(lngGelIndex).Orfs(lngORFIndex).Sequence, intTrypticStartNum, lngResidueStartForCurrent, lngResidueEndForCurrent)
                            .ResidueStart = lngResidueStartForCurrent
                            
                            If intTrypticCount = 1 Then
                                .ResidueEnd = lngResidueEndForCurrent
                            Else
                                strFragment = objMwtWin.Peptide.GetTrypticPeptideByFragmentNumber(GelORFData(lngGelIndex).Orfs(lngORFIndex).Sequence, intTrypticStartNum + intTrypticCount - 1, lngResidueStartForCurrent, lngResidueEndForCurrent)
                                .ResidueEnd = lngResidueEndForCurrent
                            End If
                        End If
                    End If
                End With
            Next lngMassTagIndex
        End With
        
        If lngORFIndex Mod 10 = 0 Then
            frmProgress.UpdateProgressBar lngORFIndex
            If KeyPressAbortProcess > 1 Then
                AddToAnalysisHistory lngGelIndex, "User prematurely aborted process of updating ORF statistics"
                Exit For
            End If
        End If
    Next lngORFIndex
    
    frmProgress.HideForm
End Sub

Public Sub UpdateSavedGelListAndOptions(udtCurrentGelDisplayListAndOptions As udtORFViewerGelListType)
    ' Copies data from udtCurrentGelDisplayListAndOptions to
    '  GelORFViewerSavedGelListAndOptions().SavedGelListAndOptions
    '  for all gels currently "Included"
    
    Dim lngGelIndex As Long
    
    If ORFViewerLoader.InitializingORFViewerForm Then Exit Sub
    
    
    With udtCurrentGelDisplayListAndOptions
        For lngGelIndex = 1 To .GelCount
            If .Gels(lngGelIndex).IncludeGel Then
                GelORFViewerSavedGelListAndOptions(lngGelIndex).IsDefined = True
                GelORFViewerSavedGelListAndOptions(lngGelIndex).SavedGelListAndOptions = udtCurrentGelDisplayListAndOptions
            End If
        Next lngGelIndex
    End With
    
End Sub

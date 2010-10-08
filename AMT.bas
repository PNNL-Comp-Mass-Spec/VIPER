Attribute VB_Name = "Module13"
'last modified 01/10/2002 nt
'------------------------------------------------------------------------
'this module is centered around SearchAMT function. It contains procedure
'to connect with the AMT database and load essential data from it to be
'used for faster search. Currently it supports search on MW only or MW
'and NET fields and is optimized for speed (although it is possible to
'optimize it even more)
'Search parameters are contained in gl.var. of type SearchAMTDefinition
'------------------------------------------------------------------------
Option Explicit

Public Enum dbgDatabaseGenerationConstants
    dbgUnknown = 0
    dbgGeneration1 = 1          ' Legacy Access file DB with table [AMT]; no Retention Time; no N atoms count
    dbgGeneration0800 = 2       ' Legacy Access file DB with table [AMT]; no N atoms
    dbgGeneration0900 = 3       ' Legacy Access file DB with table [AMT]
    dbgGeneration1000 = 4       ' Legacy Access file DB with table [AMT]; field names changes; added field "AA_Cystine_Count"; added field "MSMS_Obs_Count"
    dbgMTSOffline = 5           ' Access DB with tables [T_Mass_Tags] and [T_Mass_Tags_GANET]
    dbgMTSOnline = 6            ' Sql Server MTS Database
End Enum

' These fieldes apply to dbgGeneration1 through dbgGeneration1000
Private Const DB_FIELD_AMT_OLD_ID = "ID"
Private Const DB_FIELD_AMT_NEW_ID = "AMT_ID"
Private Const DB_FIELD_AMT_PEPTIDE = "Peptide"
Private Const DB_FIELD_AMT_MW = "AMTMonoisotopicMass"
Private Const DB_FIELD_AMT_NET = "NET"
Private Const DB_FIELD_AMT_RETENTION = "RetentionTime"            ' Stored in the AMTData().PNET; ignored if field PNET is present
Private Const DB_FIELD_AMT_PNET = "PNET"
Private Const DB_FIELD_AMT_NitrogenAtom = "NitrogenAtom"
Private Const DB_FIELD_AMT_CysCount = "AA_Cystine_Count"
Private Const DB_FIELD_AMT_Status = "Status"
Private Const DB_FIELD_AMT_MSMSObsCount = "MSMS_Obs_Count"
Private Const DB_FIELD_AMT_HighNormalizedScore = "High_Normalized_Score"
Private Const DB_FIELD_AMT_HighDiscriminantScore = "High_Discriminant_Score"
Private Const DB_FIELD_AMT_PeptideProphetProbability = "Peptide_Prophet_Probability"

Private Const DB_FIELD_PROTEIN_AMT_ID As String = "AMT_ID"
Private Const DB_FIELD_PROTEIN_Protein_ID As String = "Protein_ID"
Private Const DB_FIELD_PROTEIN_Protein_Name As String = "Protein_Name"

' These fields apply to dbgMTSOffline
Private Const DB_FIELD_TMASSTAGS_MASS_TAG_ID = "Mass_Tag_ID"
Private Const DB_FIELD_TMASSTAGS_PEPTIDE = "Peptide"
Private Const DB_FIELD_TMASSTAGS_MW = "Monoisotopic_Mass"
Private Const DB_FIELD_TMASSTAGS_MSMSObsCount = "Peptide_Obs_Count_Passing_Filter"
Private Const DB_FIELD_TMASSTAGS_HighNormalizedScore = "High_Normalized_Score"
Private Const DB_FIELD_TMASSTAGS_HighDiscriminantScore = "High_Discriminant_Score"
Private Const DB_FIELD_TMASSTAGS_PeptideProphetProbability = "High_Peptide_Prophet_Probability"

Private Const DB_FIELD_TMTNET_NET = "Avg_GANET"
Private Const DB_FIELD_TMTNET_PNET = "PNET"
Private Const DB_FIELD_TMTNET_STDEV = "StD_GANET"
Private Const DB_FIELD_TMTNET_CNT_NET = "Cnt_GANET"

Private Const DB_FIELD_TPROTEINS_Protein_ID As String = "Ref_ID"
Private Const DB_FIELD_TPROTEINS_Protein_Name As String = "Reference"
Private Const DB_FIELD_TPROTEINS_Protein_Description As String = "Description"
Private Const DB_FIELD_TPROTEINS_Protein_Sequence As String = "Protein_Sequence"
Private Const DB_FIELD_TPROTEINS_Protein_ResidueCount As String = "Protein_Residue_Count"
Private Const DB_FIELD_TPROTEINS_Protein_Mass As String = "Monoisotopic_Mass"

Private Const DB_FIELD_TMASSTAGSPROTEINMAP_MASS_TAG_Name = "Mass_Tag_Name"


Public Const glAMT_NET = 0
Public Const glAMT_RT_or_PNET = 1

Public Const glAMT_CONFIRM_NO = 0
Public Const glAMT_CONFIRM_PPM = 1
Public Const glAMT_CONFIRM_PPM_NET = 2
Public Const glAMT_CONFIRM_N14_N15 = 4
Public Const glAMT_CONFIRM_CYS = 8
Public Const glAMT_CONFIRM_MSMS1 = 256
Public Const glAMT_CONFIRM_MSMS2 = 512
Public Const glAMT_CONFIRM_MSMS3PLUS = 1024

Public Const AMTMark = "AMT:"
Public Const AMTIDEnd = "("
Public Const MWErrMark = "(MW Err: "
Public Const MWErrEnd = "ppm)"

Public Const NETErrMark = "(NET Err: "
Public Const NETErrEnd = ")"
Public Const RTErrMark = "(RT Err: "
Public Const RTErrEnd = ")"

Public Const NCntMark = "(N: "

'generic elution time (in use since 07/2001)
Public Const ETErrMark = "(ET Err: "
Public Const ETErrEnd = ")"

Public Const LCK_MARK = "LCK:"
Public Const LckIDEnd = "("
'continue to use AMT mark until the final transition
'''Public Const MTG_MARK = "MTG:"

'Public Const NET_LCK_MARK = "NETLK:"            ' This cannot end in LCK: since that would cause problems with LCK_MARK = "LCK:"
Public Const INT_STD_MARK = "ISTD:"
Public Const INT_STD_IDEnd = "("

Public Const MTSLiCMark = "(SLiC: "
Public Const MTSLiCEnd = ")"
Public Const MTSLiCMarkLength = 7

Public Const MTDelSLiCMark = "(DelSLiC: "
Public Const MTDelSLiCEnd = ")"

Public Const MTUPMark = "(UP: "
Public Const MTUpEND = ")"

Public Const MTDltMark = "Dlt:"         'used both for labeling
Public Const MTNCntMark = "(N:"         'used if N count is stored
'''Public Const MTCysCntMark = "(Cys:"     'used if Cys count is stored

Public Const MTEndMark = ")"

'''Public Const MTLightMark = "LT"         'light
'''Public Const MTHeavyMark = "HV"         'heavy
'''Public Const MTLightLblMark = "LL"      'light labeled
'''Public Const MTHeavyLblMark = "HL"      'heavy labeled
'''Public Const MTLightModMark = "LM"      'light modified
'''Public Const MTHeavyModMark = "HM"      'heavy modified
'''
'''Public Const AMTN15Mark = "(N15)"

Public Enum asrbAMTSearchResultsBehaviorConstants
    asrbAutoRemoveExisting = 0
    asrbKeepExisting = 1
    asrbKeepExistingAndSkip = 2
End Enum

Public Enum srsSearchRegionShapeConstants
    srsElliptical = 0
    srsRectangular = 1
End Enum

' MonroeMod: Used to mark ions that inherit a match from another member of a UMC
Public Const AMTMatchInheritedMark = "(inherited)"

Private Const TABLE_NAME_AMT = "AMT"
Private Const TABLE_NAME_AMT_PROTEINS = "AMT_Proteins"
Private Const TABLE_NAME_AMT_TO_PROTEIN_MAP = "AMT_to_Protein_Map"

Private Const TABLE_NAME_T_MASS_TAGS = "T_Mass_Tags"
Private Const TABLE_NAME_T_MASS_TAGS_NET = "T_Mass_Tags_NET"
Private Const TABLE_NAME_T_PROTEINS = "T_Proteins"
Private Const TABLE_NAME_T_MASS_TAG_TO_PROTEIN_MAP = "T_Mass_Tag_to_Protein_Map"
Private Const TABLE_NAME_T_MASS_TAG_TO_PROTEIN_MAP_ALT = "T_Mass_Tags_to_Protein_Map"

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 2 through version 8
Public Type SearchAMTDefinition
    SearchScope As Integer       'current scope or all points (actually type glScope) ; Not used on frmSearchMT_ConglomerateUMC
    SearchFlag As Integer        'defines which AMTs are included in search; corresponds to constants glAMT_CONFIRM_ ; if 0, then all data is searched ; Not used on frmSearchMT_ConglomerateUMC
    MWField As Integer           'MWField for Isotopic data
    TolType As Integer           'ppm or Dalton; represented by 0 = ppm = gltPPM or 2 = Da = gltABS (actually type glMassToleranceConstants)
    NETorRT As Integer           'Use NET or use RT (on some forms, use NET or use PNET)
    Formula As String            'Formula to calculate NET
    MWTol As Double              'actual MW Tolerance (search is +/- this tolerance)
    NETTol As Double             'NET Tolerance (search is +/- this tolerance)
    MassTag As Double            'if used special search is performed
    MaxMassTags As Long          'maximum number of MT tags
    SkipReferenced As Boolean    'if True skip data points already referenced by AMT ID
    SaveNCnt As Boolean          'if true keep also number of N atoms in ID field; however, data referenced with Not Found will be searched
End Type

Public Type SearchORFDefinition
    SearchScope As Long
    MWField As Long
    MWTol As Double
    MWTolType As Long
    Mods As Collection
End Type

Public Type udtUMCMassTagRawMatches
    IDInd As Long                           ' Match ID; note that this is not de-referenced, so one must use mMTInd() or mInternalStdIndexPointers()
    IDIndexOriginal As Long                 ' Dereferenced pointer, directly into AMTData() array
    MatchingMemberCount As Long
    StandardizedSquaredDistance As Double
    SLiCScoreNumerator As Double
    StacOrSLiC As Double                    ' STAC Score or SLiC Score (Spatially Localized Confidence score)
    DelScore As Double                      ' Similar to DelCN, difference in STAC or SLiC score between top match and match with score value one less than this score
    UniquenessProbability As Double         ' Uniqueness Probability from STAC
    MassErr As Double                       ' Observed difference (error) between MT tag mass and UMC class mass (in Da)
    NETErr As Double                        ' Observed difference (error) between MT tag NET and UMC class NET
    IDIsInternalStd As Boolean
End Type

Public Type udtAMTDataType
    ID As Long                          'AMT ID
    flag As Integer                     'Status field
    MW As Double                        'Theoretical molecular weight
    NET As Double                       'elution time (average NET)
    MSMSObsCount As Long                'number of observations by MS/MS
    NETStDev As Double                  'elution time standard deviation
    NETCount As Long                    'number of values averaged together to compute the average NET value
    PNET As Single                      'Theoretical NET (from DB)  (previously, held retention time, in seconds)
    CNT_N As Integer                    'count of N atoms
    CNT_Cys As Integer                  'count of Cysteines
    HighNormalizedScore As Single       'High normalized score (typically XCorr)
    HighDiscriminantScore As Single     'High discriminant score
    PeptideProphetProbability As Single 'High Peptide Prophet Probability
    Sequence As String                  'peptide sequences;
End Type

Public Type udtAMTScoreStatsType
    MTID As Long                          'AMT ID
    MW As Double                        'Theoretical molecular weight
    NET As Double                       'elution time
    NETStDev As Double                  'elution time standard deviation
    NETCount As Integer
    MSMSObsCount As Long                'number of observations by MS/MS
    HighNormalizedScore As Single       'High normalized score (typically XCorr)
    HighDiscriminantScore As Single     'High discriminant score
    PeptideProphetProbability As Single 'High Peptide Prophet Probability
    ModCount As Integer
    TrypticState As Byte
    PepProphetObsCountCS1 As Integer
    PepProphetObsCountCS2 As Integer
    PepProphetObsCountCS3 As Integer
    PepProphetFScoreCS1 As Single
    PepProphetFScoreCS2 As Single
    PepProphetFScoreCS3 As Single
    PassesFilters As Byte
End Type


Private Type udtAMTFieldPresentProteinTblType
    ID As Boolean
    Name As Boolean
    Description As Boolean
    Sequence As Boolean
    ResidueCount As Boolean
    Mass As Boolean
End Type

Private Type udtAMTFieldPresentProteinPeptideMapTblType
    MTID As Boolean
    MTName As Boolean
    ProteinID As Boolean
End Type

Private Type udtAMTFieldPresentType
    MTID As Boolean
    PeptideSequence As Boolean
    MW As Boolean
    NET As Boolean
    NETStDev As Boolean
    NETCount As Boolean
    Status As Boolean
    RetentionTime As Boolean
    PNET As Boolean
    NitrogenAtom As Boolean
    CysCount As Boolean
    MSMSObsCount As Boolean
    HighNormalizedScore As Boolean
    HighDiscriminantScore As Boolean
    PeptideProphetProbability As Boolean
    
    ProteinInfo As udtAMTFieldPresentProteinTblType
    ProteinPeptideMap As udtAMTFieldPresentProteinPeptideMapTblType
End Type

'once open AMT database stays open for the duration of the application
Public dbAMT As Database

' For dbgGeneration1000 and earlier, this is True if tables AMT_Proteins & AMT_to_Protein_Map exist
' For dbgMTSOffline, this is True if tables T_Proteins & T_Mass_Tag_to_Protein_Map exist
Private AMTProteinTablesExist As Boolean

Public AMTCnt As Long         'global count of AMTs; can be used anywhere in code
Public AMTGeneration As dbgDatabaseGenerationConstants

'used to keep track of changes in Search AMT function during application run
Public samtDef As SearchAMTDefinition

Public sorfDef As SearchORFDefinition

'search flag as an array (for more efficient searching)
Private aSearchFlag() As Boolean

'Global array of data loaded from AMT database
'Array is sorted on MW (since the Stored Procedure returns the data that way)
'This is a 1-based array, ranging from 1 to AMTCnt
Public AMTData() As udtAMTDataType

'Global array of data loaded from AMT database
'Array is sorted on MTID (since the Stored Procedure returns the data that way)
'This is a 0-based array, ranging from 0 to AMTScoreStatsCnt-1
Public AMTScoreStatsCnt As Long
Public AMTScoreStats() As udtAMTScoreStatsType

'arrays down are used to compare AMT database with
'current gel and eventually recalculate NET for gels
Public AMTHits() As Long
Public AMTMWErr() As Double     'sum of absolute values of absolute errors

'following arrays are used for both NET and RT calculation
Public AMTNETErr() As Double    'sum of absolute NET/RT errors (direction could help)
Public AMTNETMin() As Double    'min of NET/RT range
Public AMTNETMax() As Double    'max of NET/RT range


'ORF arrays; for now everything I need is coming from MassTags database
Public ORFCnt As Long               'ORF count
Public ORFID() As Long              'ORF ids

Public MTtoORFMapCount As Long      'number of MT tags - ORF mappings
Public MTIDMap() As Long            'parallel arrays that establish MT tags to Proteins mapping; 1-based array
Public ORFIDMap() As Long           'parallel arrays that establish MT tags to Proteins mapping; 1-based array
Public ORFRefNames() As String      'Protein (ORF) names; 1-based array

' Unused variables (July 2003)
''''names from the T_Mass_Tags_to_ORF_Map are not loaded by default but could be
''''requested from the Overlay drawing function; resources should be freed after use
'''Public nameCnt As Long
'''Public nameMTID() As Long
'''Public nameMTName() As String

'object used to fast locate index ranges in AMTMW
Public mwutSearch As MWUtil

Private mSearchObjectHasN15Masses As Boolean        ' Set to True when mwutSearch was filled with N15 masses

'counts number of hits to the AMT data (non-unique)
Dim HitsCount As Long
'Expression Evaluator variables
Dim MyExprEva As ExprEvaluator
Dim VarVals() As Long

Public Function ConnectToLegacyAMTDB(ByRef frmCallingForm As Form, ByVal lngGelIndex As Long, ByVal AskUser As Boolean, ByVal blnLoadProteinInfo As Boolean, ByVal blnIncludeProteinsForMassTagsNotImMemory As Boolean) As Boolean
    ' Load AMT data from an Access database
    
    Static strLegacyDBNotifiedProteinTableFieldError As String
    
    Dim lngTableRowCount As Long
    Dim eDBGeneration As dbgDatabaseGenerationConstants
    Dim strMTToProteinMapTableName As String
    
    Dim eResponse As VbMsgBoxResult
    
    Dim udtFieldPresent As udtAMTFieldPresentType
    Dim strErrorMessage As String
    Dim blnSuccess As Boolean
    Dim strDBPathFull As String
    Dim fso As New FileSystemObject
    
On Error GoTo err_ConnectToLegacyAMTDB
    
    ConnectToLegacyAMTDB = False
    If AskUser And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("2DGel will access AMT database: " & vbCrLf _
                     & GelData(lngGelIndex).PathtoDatabase & vbCrLf & "Multiple versions of the AMT " _
                     & "database might exist. Make sure that listed file " _
                     & "is 'The AMT' database. Choose OK to continue. " _
                     & "(To specify different database use Options dialog.)" _
                     , vbOKCancel)
        If eResponse <> vbOK Then
            ConnectToLegacyAMTDB = False
            Exit Function
        End If
    End If
    
    strDBPathFull = GelData(lngGelIndex).PathtoDatabase
    If Not fso.FileExists(strDBPathFull) Then
        ' Look for the DB in the same folder as the .Gel file
        If Len(GelStatus(lngGelIndex).GelFilePathFull) > 0 Then
            strDBPathFull = fso.BuildPath(fso.GetParentFolderName(GelStatus(lngGelIndex).GelFilePathFull), fso.GetFileName(strDBPathFull))
        Else
            strDBPathFull = fso.BuildPath(App.Path, fso.GetFileName(strDBPathFull))
        End If
        
        If fso.FileExists(strDBPathFull) Then
            GelData(lngGelIndex).PathtoDatabase = strDBPathFull
        End If
    End If
    
    If Not fso.FileExists(GelData(lngGelIndex).PathtoDatabase) Then
        strErrorMessage = "Database file not found: " & GelData(lngGelIndex).PathtoDatabase
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            AddToAnalysisHistory lngGelIndex, strErrorMessage
        Else
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
        End If
        ConnectToLegacyAMTDB = False
        Exit Function
    End If
    
    Set dbAMT = DBEngine.Workspaces(0).OpenDatabase(GelData(lngGelIndex).PathtoDatabase, False, True)
    
    blnSuccess = DetermineLegacyAMTDBFormat(dbAMT, eDBGeneration)
    If Not blnSuccess Then
        strErrorMessage = "Could not find the expected tables in the database. Connection with AMT database will be closed. " & _
                          "Make sure the Access database contains a table named [" & TABLE_NAME_AMT & "] or the tables " & _
                          "[" & TABLE_NAME_T_MASS_TAGS & "] and [" & TABLE_NAME_T_MASS_TAGS_NET & "]." & vbCrLf & "File: " & GelData(lngGelIndex).PathtoDatabase
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            AddToAnalysisHistory lngGelIndex, strErrorMessage
        Else
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
        End If
        
        Set dbAMT = Nothing
        Exit Function
    Else
        
    End If
    
    If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
        ' Count the number of rows in tables TABLE_NAME_T_MASS_TAGS & TABLE_NAME_T_MASS_TAGS_NET
        lngTableRowCount = GetTableRowCount(dbAMT, TABLE_NAME_T_MASS_TAGS, True, GetLegacyDBRequiredMTTableFields(eDBGeneration), lngGelIndex)
        If lngTableRowCount <= 0 Then
            Set dbAMT = Nothing
            Exit Function
        End If
    
        lngTableRowCount = GetTableRowCount(dbAMT, TABLE_NAME_T_MASS_TAGS_NET, True, GetLegacyDBRequiredMTTableFields(eDBGeneration), lngGelIndex)
        If lngTableRowCount <= 0 Then
            Set dbAMT = Nothing
            Exit Function
        End If
    Else
        ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
        
        ' Count the number of rows in TABLE_NAME_AMT
        lngTableRowCount = GetTableRowCount(dbAMT, TABLE_NAME_AMT, True, GetLegacyDBRequiredMTTableFields(eDBGeneration), lngGelIndex)
        If lngTableRowCount <= 0 Then
            Set dbAMT = Nothing
            Exit Function
        End If
    End If
    
    AMTProteinTablesExist = LegacyDBCheckForProteinTables(eDBGeneration, GelData(lngGelIndex).PathtoDatabase, strMTToProteinMapTableName)
    
    ' Determine which fields are present
    ' Note that eDBGeneration will get updated here if it is dbgGeneration1 through dbgGeneration1000
    If Not EnumerateLegacyAMTFields(GelData(lngGelIndex).PathtoDatabase, lngGelIndex, udtFieldPresent, eDBGeneration) Then
        strErrorMessage = GetLegacyDBRequiredMTTableFields(eDBGeneration) & vbCrLf & "File: " & GelData(lngGelIndex).PathtoDatabase
        
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            AddToAnalysisHistory lngGelIndex, strErrorMessage
        Else
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
        End If
    
        Set dbAMT = Nothing
        Exit Function
    End If
    
    If AMTProteinTablesExist Then
        If Not EnumerateProteinTableFields(GelData(lngGelIndex).PathtoDatabase, lngGelIndex, eDBGeneration, strMTToProteinMapTableName, udtFieldPresent) Then
            If strLegacyDBNotifiedProteinTableFieldError <> GelData(lngGelIndex).PathtoDatabase Then
                strErrorMessage = GetLegacyDBRequiredProteinTableFields(eDBGeneration) & vbCrLf & "File: " & GelData(lngGelIndex).PathtoDatabase
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                    MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
                End If
            End If
            AMTProteinTablesExist = False
        End If
    End If
    
    ' The data should be ready to load
    ' First load the MT tags
    blnSuccess = LegacyDBLoadAMTData(frmCallingForm, GelData(lngGelIndex).PathtoDatabase, lngGelIndex, udtFieldPresent, eDBGeneration)
    
    If blnSuccess And blnLoadProteinInfo And AMTProteinTablesExist Then
        ' Now load the proteins
        blnSuccess = LegacyDBLoadProteinData(frmCallingForm, GelData(lngGelIndex).PathtoDatabase, lngGelIndex, blnIncludeProteinsForMassTagsNotImMemory, eDBGeneration, strMTToProteinMapTableName, udtFieldPresent)
    End If
    
    If blnSuccess Then
        glbPreferencesExpanded.LegacyAMTDBPath = GelData(lngGelIndex).PathtoDatabase
    End If
    
    ConnectToLegacyAMTDB = blnSuccess

Exit Function

err_ConnectToLegacyAMTDB:
    If Err.Number = 3024 Then
        strErrorMessage = "Error connecting to database; " & Err.Description
    Else
        strErrorMessage = "Error connecting to database: " & GelData(lngGelIndex).PathtoDatabase & "; " & Err.Description & " (Error Number " & Trim(Err.Number) & ")"
    End If
    
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        AddToAnalysisHistory lngGelIndex, strErrorMessage
    Else
        MsgBox strErrorMessage, vbExclamation + vbOKOnly, "Error"
    End If

End Function

Public Sub CloseConnections()
On Error Resume Next
If Not (dbAMT Is Nothing) Then Set dbAMT = Nothing
End Sub

Public Function ConstructAMTReference(ByVal MW As Double, _
                                      ByVal NETRT As Double, _
                                      ByVal Delta As Long, _
                                      ByVal AMTMatchIndex As Long, _
                                      ByVal dblAMTMass As Double, _
                                      ByVal dblStacOrSLiC As Double, _
                                      ByVal dblDelSLiCScore As Double, _
                                      ByVal dblUPScore As Double) As String
                                      
    'returns AMT reference string based on MW and samtDef
    'this function is called from SearchAMT and similar functions
    
    Dim AMTRef As String
    Dim MWTolRef As Double
    
    Dim blnStoreAbsoluteValueOfError As Boolean
    
    ' Note, we are no longer storing the absolute value of errors in the AMT Ref for the data
    blnStoreAbsoluteValueOfError = False
    
    On Error GoTo exit_ConstructAMTReference
    
    If blnStoreAbsoluteValueOfError Then
        MWTolRef = Abs(MW - dblAMTMass)
    Else
        MWTolRef = MW - dblAMTMass
    End If
    
    ' The following assertion will fail if we used a huge search tolerance
    'Debug.Assert Abs(MWTolRef) < 1
    
    'put AMT ID and actual errors
    If samtDef.SaveNCnt Then
        AMTRef = AMTMark & Trim(AMTData(AMTMatchIndex).ID) & _
                 MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd & _
                 MTSLiCMark & Round(dblStacOrSLiC, 4) & MTSLiCEnd & _
                 MTDelSLiCMark & Round(dblDelSLiCScore, 4) & MTDelSLiCEnd & _
                 MTUPMark & Round(dblUPScore, 4) & MTUpEND & _
                 MTNCntMark & AMTData(AMTMatchIndex).CNT_N & MTEndMark
    Else
          
        AMTRef = AMTMark & Trim(AMTData(AMTMatchIndex).ID) & _
                 MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd & _
                 MTSLiCMark & Round(dblStacOrSLiC, 4) & MTSLiCEnd & _
                 MTDelSLiCMark & Round(dblDelSLiCScore, 4) & MTDelSLiCEnd & _
                 MTUPMark & Round(dblUPScore, 4) & MTUpEND
    End If
  
    If Delta > 0 Then AMTRef = AMTRef & MTDltMark & Delta
    
    'do statistics
    AMTHits(AMTMatchIndex) = AMTHits(AMTMatchIndex) + 1
    AMTMWErr(AMTMatchIndex) = AMTMWErr(AMTMatchIndex) + Abs(MWTolRef)
    Select Case samtDef.NETorRT
    Case glAMT_NET
        ' 7/26/2004 MEM - Reversed the order of subtraction to be Observed - Database, consistent with the MWTolRef calculation above
         AMTNETErr(AMTMatchIndex) = AMTNETErr(AMTMatchIndex) + (NETRT - AMTData(AMTMatchIndex).NET)
    Case glAMT_RT_or_PNET
         If (AMTData(AMTMatchIndex).PNET >= 0) Then    'there are some negative RTs
            ' 7/26/2004 MEM - Reversed the order of subtraction to be Observed - Database, consistent with the MWTolRef calculation above
            AMTNETErr(AMTMatchIndex) = AMTNETErr(AMTMatchIndex) + (NETRT - AMTData(AMTMatchIndex).PNET)
         End If
    End Select
    If NETRT < AMTNETMin(AMTMatchIndex) Then AMTNETMin(AMTMatchIndex) = NETRT
    If NETRT > AMTNETMax(AMTMatchIndex) Then AMTNETMax(AMTMatchIndex) = NETRT

exit_ConstructAMTReference:
    ConstructAMTReference = AMTRef & glARG_SEP & " "

End Function

Public Function ConstructInternalStdReference(ByVal MW As Double, _
                                              ByVal NETRT As Double, _
                                              ByVal InternalStdIndex As Long, _
                                              ByVal dblSLiCScore As Double, _
                                              ByVal dblDelSLiCScore As Double, _
                                              ByVal dblUPScore As Double) As String
    
    'returns InternalStd reference string based on MW and samtDef
    'this function is called from SearchAMT and similar functions
    
    Dim IntStdRef As String
    Dim MWTolRef As Double
    Dim NETTolRef As Double
    
    Dim sMWTolRef As String
    Dim sNETTolRef As String

    Dim blnStoreAbsoluteValueOfError As Boolean
    
    ' Note, we are no longer storing the absolute value of errors in the AMT Ref for the data
    blnStoreAbsoluteValueOfError = False

On Error GoTo exit_ConstructInternalStdReference
    
    With UMCInternalStandards.InternalStandards(InternalStdIndex)
        If blnStoreAbsoluteValueOfError Then
            MWTolRef = Abs(MW - .MonoisotopicMass)
        Else
            MWTolRef = MW - .MonoisotopicMass
        End If
        
        NETTolRef = .NET - NETRT

        ' The following assertion will fail if we used a huge search tolerance
        Debug.Assert Abs(MWTolRef) < 1
        
        sMWTolRef = MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd
        sNETTolRef = NETErrMark & Format$(NETTolRef, "0.000") & NETErrEnd
        
        'put Internal Standard ID and actual errors
        IntStdRef = INT_STD_MARK & .SeqID & sMWTolRef & sNETTolRef & _
                    MTSLiCMark & Round(dblSLiCScore, 4) & MTSLiCEnd & _
                    MTDelSLiCMark & Round(dblDelSLiCScore, 4) & MTSLiCEnd & _
                    MTUPMark & Round(dblUPScore, 4) & MTUpEND & _
                    glARG_SEP & " "
                    
    End With

exit_ConstructInternalStdReference:
    ConstructInternalStdReference = IntStdRef

End Function

Private Function DetermineLegacyAMTDBFormat(ByRef dbAMT As Database, ByRef eDBGeneration As dbgDatabaseGenerationConstants) As Boolean
    Dim objTableDef As TableDef
    Dim blnMatchFound As Boolean
    
    ' First look for table TABLE_NAME_T_MASS_TAGS
    blnMatchFound = False
    For Each objTableDef In dbAMT.TableDefs
        If UCase(objTableDef.Name) = UCase(TABLE_NAME_T_MASS_TAGS) Then
            blnMatchFound = True
            Exit For
        End If
    Next objTableDef
    
    If blnMatchFound Then
        eDBGeneration = dbgMTSOffline
    Else
        For Each objTableDef In dbAMT.TableDefs
            If UCase(objTableDef.Name) = UCase(TABLE_NAME_AMT) Then
                blnMatchFound = True
                Exit For
            End If
        Next objTableDef
        
        If blnMatchFound Then
            ' We'll set eDBGeneration to dbgGeneration1000 for now
            ' It could also be dbgGeneration1, dbgGeneration0800, or dbgGeneration0900
            ' The correct format will be determined in EnumerateLegacyAMTFields
            eDBGeneration = dbgGeneration1000
        End If
    End If
    
    DetermineLegacyAMTDBFormat = blnMatchFound
    
End Function

Public Function FillMWSearchObject(ByRef objMWUtil As MWUtil) As Boolean
    Dim dblMW() As Double
    Dim lngIndex As Long
    Dim blnSuccess As Boolean
    
    ReDim dblMW(LBound(AMTData) To UBound(AMTData))
    For lngIndex = LBound(AMTData) To UBound(AMTData)
        dblMW(lngIndex) = AMTData(lngIndex).MW
    Next lngIndex
    
    blnSuccess = mwutSearch.Fill(dblMW())
    FillMWSearchObject = blnSuccess
    
End Function

Private Function GetTableRowCount(ByRef objThisDB As Database, ByVal strTableName As String, ByVal blnInformUserIfMissingOrNoData As Boolean, Optional ByVal strRequiredTableFields As String = "", Optional lngGelIndex As Long = 0) As Long
    ' Determines the number of rows in table strTableName in database objThisDB
    ' Returns 0 if the table is not found or no rows
    
    Dim strErrorMessage As String
    Dim lngTableRowCount As Long
    
    On Error Resume Next
    
    strErrorMessage = ""
    lngTableRowCount = objThisDB.TableDefs(strTableName).RecordCount
    
    If blnInformUserIfMissingOrNoData Then
        If Err Then
            strErrorMessage = "Error accessing the [" & strTableName & "] table. Make sure the Access database contains a table named [" & strTableName & "]."
        ElseIf lngTableRowCount <= 0 Then
            strErrorMessage = "No records found in the [" & strTableName & "] table."
        End If
        
        If Len(strErrorMessage) > 0 Then
            strErrorMessage = strErrorMessage & "  Connection with AMT database will be closed. " & strRequiredTableFields & vbCrLf & "File: " & GelData(lngGelIndex).PathtoDatabase
            lngTableRowCount = 0
            
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                AddToAnalysisHistory lngGelIndex, strErrorMessage
            Else
                MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
            End If
        End If
    End If
    
    If Err Then
        Err.Clear
    End If
    
    GetTableRowCount = lngTableRowCount
    
End Function

' Unused function (March 2006)
''Public Function GetAMTRecordByID(ByVal ID As String) As String
'''retrieves record from AMT by it's ID (primary key)
''MsgBox "Not implemented at the moment.", vbOKOnly, glFGTU
''End Function

Private Function LegacyDBCheckForProteinTables(ByVal eDBGeneration As dbgDatabaseGenerationConstants, ByVal strLegacyDBPath As String, ByRef strMTToProteinMapTableName As String) As Boolean

    Static strLegacyDBPathSaved As String
    
    Dim blnAMTProteinTablesExist As Boolean
    
    Dim strProteinTableName As String
    Dim strMTToProteinMapTableNameDefault As String
    Dim strMTtoProteinMapTableNameAlt As String
    
    Dim strErrorMessage As String
    
    Dim lngTableRowCount As Long
    
    blnAMTProteinTablesExist = False
    If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
        strProteinTableName = "[" & TABLE_NAME_T_PROTEINS & "]"
        strMTToProteinMapTableNameDefault = "[" & TABLE_NAME_T_MASS_TAG_TO_PROTEIN_MAP & "]"
        strMTtoProteinMapTableNameAlt = "[" & TABLE_NAME_T_MASS_TAG_TO_PROTEIN_MAP_ALT & "]"
    Else
        ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
        strProteinTableName = "[" & TABLE_NAME_AMT_PROTEINS & "]"
        strMTToProteinMapTableNameDefault = "[" & TABLE_NAME_AMT_TO_PROTEIN_MAP & "]"
        strMTtoProteinMapTableNameAlt = ""
    End If

    strMTToProteinMapTableName = strMTToProteinMapTableNameDefault
    
    ' See if the Protein mapping tables exist
    lngTableRowCount = GetTableRowCount(dbAMT, strProteinTableName, False)
    If lngTableRowCount > 0 Then
        lngTableRowCount = GetTableRowCount(dbAMT, strMTToProteinMapTableNameDefault, False)
        If lngTableRowCount > 0 Then
            blnAMTProteinTablesExist = True
        ElseIf Len(strMTtoProteinMapTableNameAlt) > 0 Then
            lngTableRowCount = GetTableRowCount(dbAMT, strMTtoProteinMapTableNameAlt, False)
            If lngTableRowCount > 0 Then
                strMTToProteinMapTableName = strMTtoProteinMapTableNameAlt
                blnAMTProteinTablesExist = True
            End If
        End If
    End If

    If Not blnAMTProteinTablesExist Then
        If strLegacyDBPathSaved <> strLegacyDBPath Then
            strLegacyDBPathSaved = strLegacyDBPath
        
            strErrorMessage = "Could not find the " & strProteinTableName & " and/or " & strMTToProteinMapTableName & " tables in the Access database.  " & _
                              "Analysis will continue but protein information will not be loaded.  To load protein information, make sure the database contains these tables.  " & _
                              GetLegacyDBRequiredProteinTableFields(eDBGeneration) & vbCrLf & "File: " & strLegacyDBPath
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
            End If
        End If
    End If
    
    LegacyDBCheckForProteinTables = blnAMTProteinTablesExist
    
End Function
    
Private Function LegacyDBLoadAMTData(ByRef frmCallingForm As VB.Form, ByVal strLegacyDBFilePath As String, ByVal lngGelIndex As Long, ByRef udtFieldPresent As udtAMTFieldPresentType, ByVal eDBGeneration As dbgDatabaseGenerationConstants) As Boolean
'------------------------------------------------------------
'loads data from AMT table in a Microsoft Access file into arrays
'returns True on success
'this function can be called at any time to refresh arrays
'------------------------------------------------------------
    
    Dim rsAMTSQL As String
    Dim rsAMT As Recordset
    Dim IDFieldName As String
    Dim strMTTable As String
    Dim strNETTable As String
    Dim strErrorMessage As String
    Dim strCaptionSaved As String
    
    Dim lngMassTagCountWithNullValues As Long
    
    ' Save the form's caption
    strCaptionSaved = frmCallingForm.Caption
    

On Error GoTo err_LegacyDBLoadAMTData
    
    Select Case eDBGeneration
    Case dbgGeneration1
        strMTTable = "[" & TABLE_NAME_AMT & "]."
        IDFieldName = DB_FIELD_AMT_OLD_ID
        rsAMTSQL = " SELECT " & strMTTable & DB_FIELD_AMT_OLD_ID & ", " & _
                                strMTTable & DB_FIELD_AMT_MW & ", " & _
                                strMTTable & DB_FIELD_AMT_NET & ", " & _
                                strMTTable & DB_FIELD_AMT_Status & _
                   " FROM [" & TABLE_NAME_AMT & "]" & _
                   " ORDER BY " & strMTTable & DB_FIELD_AMT_MW & ";"
               
    Case dbgGeneration0800
        strMTTable = "[" & TABLE_NAME_AMT & "]."
        IDFieldName = DB_FIELD_AMT_OLD_ID
        rsAMTSQL = " SELECT " & strMTTable & DB_FIELD_AMT_OLD_ID & ", " & _
                                strMTTable & DB_FIELD_AMT_MW & ", " & _
                                strMTTable & DB_FIELD_AMT_NET & ", " & _
                                strMTTable & DB_FIELD_AMT_Status & ", " & _
                                strMTTable & DB_FIELD_AMT_RETENTION & _
                   " FROM [" & TABLE_NAME_AMT & "]" & _
                   " ORDER BY " & strMTTable & DB_FIELD_AMT_MW & ";"
               
    Case dbgGeneration0900
        strMTTable = "[" & TABLE_NAME_AMT & "]."
        IDFieldName = DB_FIELD_AMT_OLD_ID
        rsAMTSQL = " SELECT " & strMTTable & DB_FIELD_AMT_OLD_ID & ", " & _
                                strMTTable & DB_FIELD_AMT_MW & ", " & _
                                strMTTable & DB_FIELD_AMT_NET & ", " & _
                                strMTTable & DB_FIELD_AMT_Status & ", " & _
                                strMTTable & DB_FIELD_AMT_RETENTION & ", " & _
                                strMTTable & DB_FIELD_AMT_NitrogenAtom & _
                   " FROM [" & TABLE_NAME_AMT & "]" & _
                   " ORDER BY " & strMTTable & DB_FIELD_AMT_MW & ";"
               
    Case dbgGeneration1000
        strMTTable = "[" & TABLE_NAME_AMT & "]."
        IDFieldName = DB_FIELD_AMT_NEW_ID
        rsAMTSQL = "SELECT " & strMTTable & DB_FIELD_AMT_NEW_ID & ", " & _
                               strMTTable & DB_FIELD_AMT_MW & ", " & _
                               strMTTable & DB_FIELD_AMT_NET
    
        If udtFieldPresent.PeptideSequence Then
            rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_PEPTIDE
        End If
    
        If udtFieldPresent.Status Then
            rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_Status
        End If
        
        If udtFieldPresent.PNET Then
            rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_PNET
            
            ' Do not read the Retention column if the PNET column is present
            udtFieldPresent.RetentionTime = False
        End If
                          
        If udtFieldPresent.RetentionTime Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_RETENTION
        If udtFieldPresent.NitrogenAtom Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_NitrogenAtom
        If udtFieldPresent.CysCount Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_CysCount
        If udtFieldPresent.MSMSObsCount Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_MSMSObsCount
        If udtFieldPresent.HighNormalizedScore Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_HighNormalizedScore
        If udtFieldPresent.HighDiscriminantScore Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_HighDiscriminantScore
        If udtFieldPresent.PeptideProphetProbability Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_AMT_PeptideProphetProbability
                          
        rsAMTSQL = rsAMTSQL & " FROM [" & TABLE_NAME_AMT & "]" & _
                              " ORDER BY " & strMTTable & DB_FIELD_AMT_MW & ";"
    Case dbgMTSOffline
        
        ' Query both T_Mass_Tags and T_Mass_Tags_NET in one query
        ' We can use an Inner Join since EnumerateLegacyAMTFields has already validated that both tables exist
        
        strMTTable = "[" & TABLE_NAME_T_MASS_TAGS & "]."
        strNETTable = "[" & TABLE_NAME_T_MASS_TAGS_NET & "]."
        
        rsAMTSQL = "SELECT " & strMTTable & DB_FIELD_TMASSTAGS_MASS_TAG_ID & ", " & _
                               strMTTable & DB_FIELD_TMASSTAGS_MW & ", " & _
                               strNETTable & DB_FIELD_TMTNET_NET
    
        If udtFieldPresent.PeptideSequence Then
            rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_TMASSTAGS_PEPTIDE
        End If
        
        If udtFieldPresent.PNET Then
            rsAMTSQL = rsAMTSQL & ", " & strNETTable & DB_FIELD_TMTNET_PNET
        End If
        
        If udtFieldPresent.NETStDev Then
            rsAMTSQL = rsAMTSQL & ", " & strNETTable & DB_FIELD_TMTNET_STDEV
        End If
        
        If udtFieldPresent.NETCount Then
            rsAMTSQL = rsAMTSQL & ", " & strNETTable & DB_FIELD_TMTNET_CNT_NET
        End If
                          
        If udtFieldPresent.MSMSObsCount Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_TMASSTAGS_MSMSObsCount
        If udtFieldPresent.HighNormalizedScore Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_TMASSTAGS_HighNormalizedScore
        If udtFieldPresent.HighDiscriminantScore Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_TMASSTAGS_HighDiscriminantScore
        If udtFieldPresent.PeptideProphetProbability Then rsAMTSQL = rsAMTSQL & ", " & strMTTable & DB_FIELD_TMASSTAGS_PeptideProphetProbability
                          
        rsAMTSQL = rsAMTSQL & " FROM [" & TABLE_NAME_T_MASS_TAGS & "] INNER JOIN" & _
                                   " [" & TABLE_NAME_T_MASS_TAGS_NET & "] ON " & _
                                   " " & strMTTable & DB_FIELD_TMASSTAGS_MASS_TAG_ID & " = " & _
                                   " " & strNETTable & DB_FIELD_TMASSTAGS_MASS_TAG_ID & _
                              " ORDER BY " & strMTTable & DB_FIELD_TMASSTAGS_MW & ";"

    Case Else
        ' Unknown version
        Debug.Assert False
        LegacyDBLoadAMTData = False
    End Select
    
    Set rsAMT = dbAMT.OpenRecordset(rsAMTSQL, dbOpenForwardOnly)
    
    ' Initially reserve space for 1000 entries
    AMTCnt = 0
    ReDim AMTData(1 To 1000)

    ' Clear MTtoORFMapCount
    MTtoORFMapCount = 0
        
    ' Clear the Score Stats array; this data cannot be loaded from Legacy DBs
    AMTScoreStatsCnt = 0
    ReDim AMTScoreStats(0)
    
    Do Until rsAMT.EOF
        If AMTCnt >= UBound(AMTData) Then
            If AMTCnt < 250000 Then
                ReDim Preserve AMTData(1 To UBound(AMTData) * 2)
            Else
                ReDim Preserve AMTData(1 To UBound(AMTData) * 1.5)
            End If
        End If
        
        If LegacyDBLoadAMTDataWork(rsAMT, AMTCnt + 1, lngMassTagCountWithNullValues, udtFieldPresent, eDBGeneration) Then
            AMTCnt = AMTCnt + 1
        End If
        
        rsAMT.MoveNext
    
        If AMTCnt Mod 100 = 0 Then frmCallingForm.Caption = "Loading MT Tags: " & LongToStringWithCommas(AMTCnt)
    Loop
    rsAMT.Close
    
exit_LegacyDBLoadAMTData:
    
    ' Restore the caption on the calling form
    frmCallingForm.Caption = strCaptionSaved
    
    If (AMTCnt > 0) Then
        ReDim Preserve AMTData(1 To AMTCnt)
         
         ' Update the AMT staleness stats
        With glbPreferencesExpanded.MassTagStalenessOptions
            .AMTLoadTime = Now()
            .AMTCountInDB = AMTCnt
            .AMTCountWithNulls = lngMassTagCountWithNullValues
        End With
        
        LegacyDBLoadAMTData = True
    Else
        AMTCnt = 0
        LegacyDBLoadAMTData = False
    End If
    
    ' mark that currently loaded data is coming from legacy database
    AMTGeneration = eDBGeneration
    CurrMTDatabase = ""
    CurrLegacyMTDatabase = strLegacyDBFilePath
    CurrMTSchemaVersion = 0
    With CurrMTFilteringOptions
        .MTSubsetID = -1
        .MTIncList = ""
    End With
    
    If GelAnalysis(lngGelIndex) Is Nothing Then
        Set GelAnalysis(lngGelIndex) = New FTICRAnalysis
    
        Dim udtEmptyAnalysisInfo As udtGelAnalysisInfoType
        FillGelAnalysisObject GelAnalysis(lngGelIndex), udtEmptyAnalysisInfo
    End If
    
    Set rsAMT = Nothing
    Exit Function

err_LegacyDBLoadAMTData:

    strErrorMessage = "Error loading AMT data from file: " & strLegacyDBFilePath & "; " & Err.Description & " (Error Number " & Trim(Err.Number) & ")"
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        AddToAnalysisHistory lngGelIndex, strErrorMessage
    Else
        MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
    End If

    Resume exit_LegacyDBLoadAMTData

End Function

Private Function LegacyDBLoadAMTDataWork(ByRef rsAMT As Recordset, ByVal lngIndex As Long, ByRef lngMassTagCountWithNullValues As Long, ByRef udtFieldPresent As udtAMTFieldPresentType, ByVal eDBGeneration As dbgDatabaseGenerationConstants) As Boolean
    Const NET_VALUE_IF_NULL As Single = -100000

    Dim blnSuccess As Boolean
    Dim blnComputeNitrogenCount As Boolean
    
On Error GoTo LegacyDBLoadAMTDataWorkErrorHandler
    
    ' Note: This function assumes (and requires) that udtFieldPresent.MTID, udtFieldPresent.MW,
    '       and udtFieldPresent.NET are all True
    
    blnSuccess = False
    With rsAMT
        Select Case eDBGeneration
        Case dbgMTSOffline
            
            ' Initialize this entry
            InitializeAMTDataEntry AMTData(lngIndex), NET_VALUE_IF_NULL
                        
            AMTData(lngIndex).ID = CLng(.Fields(DB_FIELD_TMASSTAGS_MASS_TAG_ID).Value)
            AMTData(lngIndex).MW = CDbl(.Fields(DB_FIELD_TMASSTAGS_MW).Value)
            
            If IsNull(.Fields(DB_FIELD_TMTNET_NET).Value) Then
                AMTData(lngIndex).NET = NET_VALUE_IF_NULL
            Else
                AMTData(lngIndex).NET = CDbl(.Fields(DB_FIELD_TMTNET_NET).Value)
            End If
            
            
            If udtFieldPresent.PeptideSequence Then
                If Not IsNull(.Fields(DB_FIELD_TMASSTAGS_PEPTIDE).Value) Then
                    AMTData(lngIndex).Sequence = CStr(.Fields(DB_FIELD_TMASSTAGS_PEPTIDE).Value)
                End If
            End If
            
            If udtFieldPresent.PNET Then
                If Not IsNull(.Fields(DB_FIELD_TMTNET_PNET).Value) Then
                    AMTData(lngIndex).PNET = CSng(.Fields(DB_FIELD_TMTNET_PNET).Value)
                End If
            End If
            
            If udtFieldPresent.NETStDev Then
                If Not IsNull(.Fields(DB_FIELD_TMTNET_STDEV).Value) Then
                    AMTData(lngIndex).NETStDev = CSng(.Fields(DB_FIELD_TMTNET_STDEV).Value)
                End If
            End If
            
            If udtFieldPresent.NETCount Then
               If Not IsNull(.Fields(DB_FIELD_TMTNET_CNT_NET).Value) Then
                  AMTData(lngIndex).NETCount = CLng(.Fields(DB_FIELD_TMTNET_CNT_NET).Value)
               End If
            End If
            
            
            ' Correct for files that have PNET defined but not NET, or vice versa
            ' This program uses AMTData().NET by default; AMTData().PNET historically held the retention time, in seconds, but now holds Predicted NET
            ' If one is missing from the Access DB file, then we'll copy the value from the other column to the missing column
            
            If AMTData(lngIndex).NET = NET_VALUE_IF_NULL And AMTData(lngIndex).PNET > NET_VALUE_IF_NULL Then
                 AMTData(lngIndex).NET = AMTData(lngIndex).PNET
            ElseIf AMTData(lngIndex).NET > NET_VALUE_IF_NULL And AMTData(lngIndex).PNET = NET_VALUE_IF_NULL Then
                 AMTData(lngIndex).PNET = AMTData(lngIndex).NET
            End If
            
            If AMTData(lngIndex).NET = NET_VALUE_IF_NULL Then
                lngMassTagCountWithNullValues = lngMassTagCountWithNullValues + 1
            End If
            
            If udtFieldPresent.MSMSObsCount Then
               If Not IsNull(.Fields(DB_FIELD_TMASSTAGS_MSMSObsCount).Value) Then
                  AMTData(lngIndex).MSMSObsCount = CLng(.Fields(DB_FIELD_TMASSTAGS_MSMSObsCount).Value)
               End If
            End If
            
            If udtFieldPresent.HighNormalizedScore Then
               If Not IsNull(.Fields(DB_FIELD_TMASSTAGS_HighNormalizedScore).Value) Then
                  AMTData(lngIndex).HighNormalizedScore = CSng(.Fields(DB_FIELD_TMASSTAGS_HighNormalizedScore).Value)
               End If
            End If
            
            If udtFieldPresent.HighDiscriminantScore Then
               If Not IsNull(.Fields(DB_FIELD_TMASSTAGS_HighDiscriminantScore).Value) Then
                  AMTData(lngIndex).HighDiscriminantScore = CSng(.Fields(DB_FIELD_TMASSTAGS_HighDiscriminantScore).Value)
               End If
            End If
            
            If udtFieldPresent.PeptideProphetProbability Then
               If Not IsNull(.Fields(DB_FIELD_TMASSTAGS_PeptideProphetProbability).Value) Then
                  AMTData(lngIndex).PeptideProphetProbability = CSng(.Fields(DB_FIELD_TMASSTAGS_PeptideProphetProbability).Value)
               End If
            End If
                     
            blnSuccess = True
        Case dbgGeneration1000, dbgGeneration1, dbgGeneration0800, dbgGeneration0900
            
             ' Initialize this entry
            InitializeAMTDataEntry AMTData(lngIndex), NET_VALUE_IF_NULL
            
            If eDBGeneration < dbgGeneration1000 Then
                AMTData(lngIndex).ID = CLng(.Fields(DB_FIELD_AMT_OLD_ID).Value)
            Else
                AMTData(lngIndex).ID = CLng(.Fields(DB_FIELD_AMT_NEW_ID).Value)
            End If
            
            AMTData(lngIndex).MW = CDbl(.Fields(DB_FIELD_AMT_MW).Value)
            
            If IsNull(.Fields(DB_FIELD_AMT_NET).Value) Then
                AMTData(lngIndex).NET = NET_VALUE_IF_NULL
            Else
                AMTData(lngIndex).NET = CDbl(.Fields(DB_FIELD_AMT_NET).Value)
            End If
            
            If udtFieldPresent.PeptideSequence Then
                If Not IsNull(.Fields(DB_FIELD_AMT_PEPTIDE).Value) Then
                    AMTData(lngIndex).Sequence = .Fields(DB_FIELD_AMT_PEPTIDE).Value
                End If
            End If
            
            If udtFieldPresent.Status Then
                AMTData(lngIndex).flag = CInt(.Fields(DB_FIELD_AMT_Status).Value)
            End If
            
            If udtFieldPresent.PNET Then
                If Not IsNull(.Fields(DB_FIELD_AMT_PNET).Value) Then
                    AMTData(lngIndex).PNET = CSng(.Fields(DB_FIELD_AMT_PNET).Value)
                End If
            End If
            
            If udtFieldPresent.RetentionTime Then
                If Not IsNull(.Fields(DB_FIELD_AMT_RETENTION).Value) Then
                    AMTData(lngIndex).PNET = CSng(.Fields(DB_FIELD_AMT_RETENTION).Value)
                End If
            End If
            
            ' Correct for files that have RT defined but not NET, or vice versa
            ' This program uses AMTData().NET by default; AMTData().PNET historically held the retention time, in seconds, but now holds Predicted NET
            ' If one is missing from the Access DB file, then we'll copy the value from the other column to the missing column
            
            If AMTData(lngIndex).NET = NET_VALUE_IF_NULL And AMTData(lngIndex).PNET > NET_VALUE_IF_NULL Then
                 AMTData(lngIndex).NET = AMTData(lngIndex).PNET
            ElseIf AMTData(lngIndex).NET > NET_VALUE_IF_NULL And AMTData(lngIndex).PNET = NET_VALUE_IF_NULL Then
                 AMTData(lngIndex).PNET = AMTData(lngIndex).NET
            End If
            
            If AMTData(lngIndex).NET = NET_VALUE_IF_NULL Then
                lngMassTagCountWithNullValues = lngMassTagCountWithNullValues + 1
            End If
            
            blnComputeNitrogenCount = True
            If udtFieldPresent.NitrogenAtom Then
                If Not IsNull(.Fields(DB_FIELD_AMT_NitrogenAtom).Value) Then
                    AMTData(lngIndex).CNT_N = CLng(.Fields(DB_FIELD_AMT_NitrogenAtom).Value)
                    blnComputeNitrogenCount = False
                End If
            End If
            
            If blnComputeNitrogenCount And udtFieldPresent.PeptideSequence Then
                If Len(AMTData(lngIndex).Sequence) > 0 Then
                    AMTData(lngIndex).CNT_N = NitrogenCount(AMTData(lngIndex).Sequence)
                End If
            End If
            
            If udtFieldPresent.CysCount Then
               If Not IsNull(.Fields(DB_FIELD_AMT_CysCount).Value) Then
                  AMTData(lngIndex).CNT_Cys = CLng(.Fields(DB_FIELD_AMT_CysCount).Value)
               End If
            End If
            
            If udtFieldPresent.MSMSObsCount Then
               If Not IsNull(.Fields(DB_FIELD_AMT_MSMSObsCount).Value) Then
                  AMTData(lngIndex).MSMSObsCount = CLng(.Fields(DB_FIELD_AMT_MSMSObsCount).Value)
                  AMTData(lngIndex).NETCount = AMTData(lngIndex).MSMSObsCount
               End If
            End If
            
            If udtFieldPresent.HighNormalizedScore Then
               If Not IsNull(.Fields(DB_FIELD_AMT_HighNormalizedScore).Value) Then
                  AMTData(lngIndex).HighNormalizedScore = CSng(.Fields(DB_FIELD_AMT_HighNormalizedScore).Value)
               End If
            End If
            
            If udtFieldPresent.HighDiscriminantScore Then
               If Not IsNull(.Fields(DB_FIELD_AMT_HighDiscriminantScore).Value) Then
                  AMTData(lngIndex).HighDiscriminantScore = CSng(.Fields(DB_FIELD_AMT_HighDiscriminantScore).Value)
               End If
            End If
            
            If udtFieldPresent.PeptideProphetProbability Then
               If Not IsNull(.Fields(DB_FIELD_AMT_PeptideProphetProbability).Value) Then
                  AMTData(lngIndex).PeptideProphetProbability = CSng(.Fields(DB_FIELD_AMT_PeptideProphetProbability).Value)
               End If
            End If
            
            blnSuccess = True
        Case Else
            ' Unknown version
            Debug.Assert False
        End Select

    End With
    
    LegacyDBLoadAMTDataWork = blnSuccess
    Exit Function

LegacyDBLoadAMTDataWorkErrorHandler:
    ' Data conversion error; skip this row
    
    Debug.Assert False
    LegacyDBLoadAMTDataWork = False
    
End Function

Private Function EnumerateLegacyAMTFields(ByVal strLegacyDBFilePath As String, ByVal lngGelIndex As Long, ByRef udtFieldPresent As udtAMTFieldPresentType, ByRef eDBGeneration As dbgDatabaseGenerationConstants) As Long
    'enumerates and returns AMT table fields and returns
    'its count; returns -1 on any error
    'also sets generation attribute to the current database
    Dim tdAMT As TableDef           ' Used with dbgGeneration1000
    Dim tdAMTNET As TableDef        ' Used with dbgMTSOffline
    
    Dim strMTTableName As String
    Dim strMTNETTableName As String
        
    Dim fldAny As Field
    Dim blnIsNewGeneration As Boolean
    Dim strErrorMessage As String
    Dim AMTFldCnt As Long
    
    With udtFieldPresent
        .MTID = False
        .PeptideSequence = False
        .MW = False
        .NET = False
        .NETStDev = False
        .NETCount = False
        .Status = False
        .RetentionTime = False
        .PNET = False
        .NitrogenAtom = False
        .CysCount = False
        .MSMSObsCount = False
        .HighNormalizedScore = False
        .HighDiscriminantScore = False
        .PeptideProphetProbability = False
        
        With .ProteinInfo
            .ID = False
            .Name = False
            .Description = False
            .Sequence = False
            .ResidueCount = False
            .Mass = False
        End With
        
        With .ProteinPeptideMap
            .MTID = False
            .MTName = False
            .ProteinID = False
        End With
    End With
    
On Error GoTo err_EnumerateLegacyAMTFields:
    
    If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
        strMTTableName = TABLE_NAME_T_MASS_TAGS
        strMTNETTableName = TABLE_NAME_T_MASS_TAGS_NET
    Else
        ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
        strMTTableName = TABLE_NAME_AMT
        strMTNETTableName = ""
    End If
    
    Set tdAMT = dbAMT.TableDefs(strMTTableName)
    AMTFldCnt = tdAMT.Fields.Count
    If AMTFldCnt > 0 Then
        For Each fldAny In tdAMT.Fields
            If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
                Select Case LCase(fldAny.Name)
                Case LCase(DB_FIELD_TMASSTAGS_MASS_TAG_ID):             udtFieldPresent.MTID = True
                Case LCase(DB_FIELD_TMASSTAGS_PEPTIDE):                 udtFieldPresent.PeptideSequence = True
                Case LCase(DB_FIELD_TMASSTAGS_MW):                      udtFieldPresent.MW = True
                Case LCase(DB_FIELD_TMASSTAGS_MSMSObsCount):            udtFieldPresent.MSMSObsCount = True
                Case LCase(DB_FIELD_TMASSTAGS_HighNormalizedScore):     udtFieldPresent.HighNormalizedScore = True
                Case LCase(DB_FIELD_TMASSTAGS_HighDiscriminantScore):   udtFieldPresent.HighDiscriminantScore = True
                Case LCase(DB_FIELD_TMASSTAGS_PeptideProphetProbability):     udtFieldPresent.PeptideProphetProbability = True
                Case Else
                    ' Unknown field; skip it
                End Select
            Else
                ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
                Select Case LCase(fldAny.Name)
                Case LCase(DB_FIELD_AMT_OLD_ID)
                    udtFieldPresent.MTID = True
                    blnIsNewGeneration = False
                Case LCase(DB_FIELD_AMT_NEW_ID)
                    udtFieldPresent.MTID = True
                    blnIsNewGeneration = True
                Case LCase(DB_FIELD_AMT_PEPTIDE):           udtFieldPresent.PeptideSequence = True
                Case LCase(DB_FIELD_AMT_MW):                udtFieldPresent.MW = True
                Case LCase(DB_FIELD_AMT_NET):               udtFieldPresent.NET = True
                Case LCase(DB_FIELD_AMT_Status):            udtFieldPresent.Status = True
                Case LCase(DB_FIELD_AMT_RETENTION):         udtFieldPresent.RetentionTime = True
                Case LCase(DB_FIELD_AMT_PNET):              udtFieldPresent.PNET = True
                Case LCase(DB_FIELD_AMT_NitrogenAtom):              udtFieldPresent.NitrogenAtom = True
                Case LCase(DB_FIELD_AMT_CysCount):                  udtFieldPresent.CysCount = True
                Case LCase(DB_FIELD_AMT_MSMSObsCount):              udtFieldPresent.MSMSObsCount = True
                Case LCase(DB_FIELD_AMT_HighNormalizedScore):       udtFieldPresent.HighNormalizedScore = True
                Case LCase(DB_FIELD_AMT_HighDiscriminantScore):     udtFieldPresent.HighDiscriminantScore = True
                Case LCase(DB_FIELD_AMT_PeptideProphetProbability):     udtFieldPresent.PeptideProphetProbability = True
                Case Else
                    ' Unknown field; skip it
                End Select
            End If
        Next fldAny
    End If
    
    If Len(strMTNETTableName) > 0 Then
    
        Set tdAMT = dbAMT.TableDefs(strMTNETTableName)
        AMTFldCnt = tdAMT.Fields.Count
        If AMTFldCnt > 0 Then
            For Each fldAny In tdAMT.Fields
                If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
                    Select Case LCase(fldAny.Name)
                    Case LCase(DB_FIELD_TMTNET_NET):     udtFieldPresent.NET = True
                    Case LCase(DB_FIELD_TMTNET_PNET):    udtFieldPresent.PNET = True
                    Case LCase(DB_FIELD_TMTNET_STDEV):   udtFieldPresent.NETStDev = True
                    Case LCase(DB_FIELD_TMTNET_CNT_NET): udtFieldPresent.NETCount = True
                    Case Else
                        ' Unknown field; skip it
                    End Select
                 Else
                     ' Nothing to do
                 End If
           Next fldAny
        End If
    
    End If
    
    ' Record type (generation) of database currently in use
    If eDBGeneration <> dbgMTSOffline Then
        If blnIsNewGeneration Then
            eDBGeneration = dbgGeneration1000
        Else
            If udtFieldPresent.RetentionTime Then
                If udtFieldPresent.NitrogenAtom Then
                    eDBGeneration = dbgGeneration0900
                Else
                    eDBGeneration = dbgGeneration0800
                End If
            Else
                eDBGeneration = dbgGeneration1
            End If
        End If
    End If
    
exit_EnumerateLegacyAMTFields:
    If udtFieldPresent.MTID And udtFieldPresent.MW And udtFieldPresent.NET Then
        EnumerateLegacyAMTFields = True
    Else
        EnumerateLegacyAMTFields = False
    End If
    
    Set tdAMT = Nothing
    Exit Function

err_EnumerateLegacyAMTFields:

    strErrorMessage = "Error enumerating AMT fields in table " & TABLE_NAME_AMT & " in file: " & strLegacyDBFilePath & "; " & Err.Description & " (Error Number " & Trim(Err.Number) & ")"
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        AddToAnalysisHistory lngGelIndex, strErrorMessage
    Else
        MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
    End If

    Resume exit_EnumerateLegacyAMTFields
End Function

Private Function EnumerateProteinTableFields(ByVal strLegacyDBFilePath As String, ByVal lngGelIndex As Long, ByVal eDBGeneration As dbgDatabaseGenerationConstants, ByVal strMTToProteinMapTableName As String, ByRef udtFieldPresent As udtAMTFieldPresentType)
    ' Makes sure the required protein fields are present in the protein tables
    ' Returns True if present, false if not
    Dim tdAMT As TableDef
    Dim fldAny As Field
    
    Dim strErrorMessage As String
    Dim strProteinTableName As String
    Dim strMTToProteinMapTableNameLcl As String
        
    Dim blnSuccess As Boolean
    
On Error GoTo err_EnumerateProteinTableFields:
   
    blnSuccess = False
    
    If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
        strProteinTableName = "[" & TABLE_NAME_T_PROTEINS & "]"
        strMTToProteinMapTableNameLcl = strMTToProteinMapTableName
    Else
        ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
        strProteinTableName = "[" & TABLE_NAME_AMT_PROTEINS & "]"
        strMTToProteinMapTableNameLcl = "[" & TABLE_NAME_AMT_TO_PROTEIN_MAP & "]"
    End If
    
    Set tdAMT = dbAMT.TableDefs(strProteinTableName)
    If tdAMT.Fields.Count > 0 Then
        For Each fldAny In tdAMT.Fields
            If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
                Select Case LCase(fldAny.Name)
                Case LCase(DB_FIELD_TPROTEINS_Protein_ID): udtFieldPresent.ProteinInfo.ID = True
                Case LCase(DB_FIELD_TPROTEINS_Protein_Name): udtFieldPresent.ProteinInfo.Name = True
                Case LCase(DB_FIELD_TPROTEINS_Protein_Description): udtFieldPresent.ProteinInfo.Description = True
                Case LCase(DB_FIELD_TPROTEINS_Protein_Sequence): udtFieldPresent.ProteinInfo.Sequence = True
                Case LCase(DB_FIELD_TPROTEINS_Protein_ResidueCount): udtFieldPresent.ProteinInfo.ResidueCount = True
                Case LCase(DB_FIELD_TPROTEINS_Protein_Mass): udtFieldPresent.ProteinInfo.Mass = True
                Case Else
                    ' Unknown field; skip it
                End Select
            Else
                ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
                Select Case LCase(fldAny.Name)
                Case LCase(DB_FIELD_PROTEIN_Protein_ID): udtFieldPresent.ProteinInfo.ID = True
                Case LCase(DB_FIELD_PROTEIN_Protein_Name): udtFieldPresent.ProteinInfo.Name = True
                End Select
            End If
        Next fldAny
    End If
    
    If udtFieldPresent.ProteinInfo.ID And udtFieldPresent.ProteinInfo.Name Then
        
        Set tdAMT = dbAMT.TableDefs(strMTToProteinMapTableNameLcl)
        If tdAMT.Fields.Count > 0 Then
            For Each fldAny In tdAMT.Fields
                If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
                    Select Case LCase(fldAny.Name)
                    Case LCase(DB_FIELD_TMASSTAGS_MASS_TAG_ID): udtFieldPresent.ProteinPeptideMap.MTID = True
                    Case LCase(DB_FIELD_TMASSTAGSPROTEINMAP_MASS_TAG_Name): udtFieldPresent.ProteinPeptideMap.MTName = True
                    Case LCase(DB_FIELD_TPROTEINS_Protein_ID): udtFieldPresent.ProteinPeptideMap.ProteinID = True
                    End Select
                Else
                    ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
                    Select Case LCase(fldAny.Name)
                    Case LCase(DB_FIELD_PROTEIN_AMT_ID): udtFieldPresent.ProteinPeptideMap.MTID = True
                    Case LCase(DB_FIELD_PROTEIN_Protein_ID): udtFieldPresent.ProteinPeptideMap.ProteinID = True
                    End Select
                End If
            Next fldAny
        End If
        
        If udtFieldPresent.ProteinPeptideMap.MTID And udtFieldPresent.ProteinPeptideMap.ProteinID Then
            blnSuccess = True
        End If
    End If
    
exit_EnumerateProteinTableFields:
    EnumerateProteinTableFields = blnSuccess
    Set tdAMT = Nothing
    Exit Function

err_EnumerateProteinTableFields:

    strErrorMessage = "Error enumerating the fields in tables " & strProteinTableName & " and " & strMTToProteinMapTableNameLcl & " in file: " & strLegacyDBFilePath & "; " & Err.Description & " (Error Number " & Trim(Err.Number) & ")"
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        AddToAnalysisHistory lngGelIndex, strErrorMessage
    Else
        MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
    End If

    Resume exit_EnumerateProteinTableFields
End Function

Private Function LegacyDBLoadProteinData(ByRef frmCallingForm As VB.Form, ByVal strLegacyDBFilePath As String, ByVal lngGelIndex As Long, ByVal blnIncludeORFsForMassTagsNotInMemory As Boolean, ByVal eDBGeneration As dbgDatabaseGenerationConstants, ByVal strMTToProteinMapTableName As String, ByRef udtFieldPresent As udtAMTFieldPresentType) As Boolean
    '---------------------------------------------------------------------------
    ' Obtains the mappings between MT tags and Protein IDs from the given legacy DB
    ' Also retrieves the Protein Names
    ' When blnIncludeORFsForMassTagsNotInMemory = True, then retrieves all MT to Protein mappings and Protein names
    ' When blnIncludeORFsForMassTagsNotInMemory = False, then only records those Protein mappings that correspond to a MT tag in memory
    ' The second method is generally faster, and definitely uses less memory if only a subset of all of the MT tags are in memory
    '---------------------------------------------------------------------------
    
    Dim rsSQL As String
    Dim rsMT_ORF_Map As Recordset
    Dim strProteinTableName As String
    Dim strMTToProteinMapTableNameLcl As String
    
    Dim lngMassTagIDToAdd As Long
    Dim blnProceed As Boolean
    Dim strCaptionSaved As String
    
    Dim strMTName As String
    
    Dim AMTIDsSorted() As Long          ' 1-based array to stay consistent with AMTData()
    Dim EmptyArray() As Long            ' Never allocate any memory for this; simply pass to objQSLong.QSAsc
    
    Dim lngAMTIndex As Long
    Dim lngORFMapItemsExamined As Long
    
    Dim i As Long
    
    Dim strErrorMessage As String
    Dim blnSuccess As Boolean
    
    Dim objQSLong As QSLong
    
    ' Save the form's caption
    strCaptionSaved = frmCallingForm.Caption

    If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
        strProteinTableName = "[" & TABLE_NAME_T_PROTEINS & "]"
        strMTToProteinMapTableNameLcl = strMTToProteinMapTableName
    Else
        ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
        strProteinTableName = "[" & TABLE_NAME_AMT_PROTEINS & "]"
        strMTToProteinMapTableNameLcl = "[" & TABLE_NAME_AMT_TO_PROTEIN_MAP & "]"
    End If
  
On Error GoTo err_LegacyDBLoadProteinData
    
    If Not blnIncludeORFsForMassTagsNotInMemory Then
        ' Need to fill a search object to speed up the search
        ' Since we only need to know if an AMT ID is in memory, we can duplicate the AMTData().ID data
        '  and sort it ascending, then supply it directly to BinarySearchLng
        ' Can't simply copy the entire array since AMTData().ID is a string array
        ' Although we could search a string array, I'd rather take the time to copy to a Long array since the searches should then be faster
        If AMTCnt > 0 Then
            ReDim AMTIDsSorted(1 To AMTCnt)
            For lngAMTIndex = 1 To AMTCnt
                AMTIDsSorted(lngAMTIndex) = AMTData(lngAMTIndex).ID
            Next lngAMTIndex
            Set objQSLong = New QSLong
            objQSLong.QSAsc AMTIDsSorted(), EmptyArray()
            Set objQSLong = Nothing
        End If
    End If
    
    ' Clear MTtoORFMapCount
    MTtoORFMapCount = 0
    
    If eDBGeneration = dbgDatabaseGenerationConstants.dbgMTSOffline Then
        rsSQL = " SELECT " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_TMASSTAGS_MASS_TAG_ID & ", " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_TPROTEINS_Protein_ID & ", " & strProteinTableName & "." & DB_FIELD_TPROTEINS_Protein_Name
        If udtFieldPresent.ProteinPeptideMap.MTName Then
            rsSQL = rsSQL & ", " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_TMASSTAGSPROTEINMAP_MASS_TAG_Name
        End If
        
        rsSQL = rsSQL & " FROM " & strProteinTableName & " INNER JOIN " & strMTToProteinMapTableNameLcl & " ON " & strProteinTableName & "." & DB_FIELD_TPROTEINS_Protein_ID & " = " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_TPROTEINS_Protein_ID & _
                         " ORDER BY " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_TMASSTAGS_MASS_TAG_ID & ", " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_TPROTEINS_Protein_ID & ";"
    Else
        ' Assume eDBGeneration is dbgGeneration1 through dbgGeneration1000
        rsSQL = " SELECT " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_PROTEIN_AMT_ID & ", " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_PROTEIN_Protein_ID & ", " & strProteinTableName & "." & DB_FIELD_PROTEIN_Protein_Name & _
                " FROM " & strProteinTableName & " INNER JOIN " & strMTToProteinMapTableNameLcl & " ON " & strProteinTableName & "." & DB_FIELD_PROTEIN_Protein_ID & " = " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_PROTEIN_Protein_ID & _
                " ORDER BY " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_PROTEIN_AMT_ID & ", " & strMTToProteinMapTableNameLcl & "." & DB_FIELD_PROTEIN_Protein_ID & ";"
    End If


    Set rsMT_ORF_Map = dbAMT.OpenRecordset(rsSQL, dbOpenSnapshot)
    rsMT_ORF_Map.MoveLast
    MTtoORFMapCount = rsMT_ORF_Map.RecordCount
    
    ' Reserve space for the mappings
    ReDim MTIDMap(1 To MTtoORFMapCount)
    ReDim ORFIDMap(1 To MTtoORFMapCount)
    ReDim ORFRefNames(1 To MTtoORFMapCount)

    ' Reset MTtoORFMapCount back to 0 since we may not load all of the mappings
    MTtoORFMapCount = 0
    
    i = 0
    With rsMT_ORF_Map
        .MoveFirst
        Do Until .EOF
            lngMassTagIDToAdd = FixNullLng(.Fields(0).Value)
            If blnIncludeORFsForMassTagsNotInMemory Then
                blnProceed = True
            Else
                If AMTCnt > 0 Then
                    If BinarySearchLng(AMTIDsSorted(), lngMassTagIDToAdd, 1, AMTCnt) >= 0 Then
                        blnProceed = True
                    Else
                        blnProceed = False
                    End If
                Else
                    blnProceed = False
                End If
            End If
            
            If blnProceed Then
                MTtoORFMapCount = MTtoORFMapCount + 1
                
                MTIDMap(MTtoORFMapCount) = lngMassTagIDToAdd
                ORFIDMap(MTtoORFMapCount) = FixNullLng(.Fields(1).Value)
                ORFRefNames(MTtoORFMapCount) = FixNull(.Fields(2).Value)

                If udtFieldPresent.ProteinPeptideMap.MTName Then
                    If Not IsNull(.Fields(DB_FIELD_TMASSTAGSPROTEINMAP_MASS_TAG_Name).Value) Then
                        strMTName = .Fields(DB_FIELD_TMASSTAGSPROTEINMAP_MASS_TAG_Name).Value
                    End If
                End If
            
            End If
           
            .MoveNext
        
            i = i + 1
            If i Mod 100 = 0 Then frmCallingForm.Caption = "Loading MT to Protein Mapping: " & LongToStringWithCommas(i)
        Loop
    End With
    rsMT_ORF_Map.Close

exit_LegacyDBLoadProteinData:
    
    ' Restore the caption on the calling form
    frmCallingForm.Caption = strCaptionSaved
    
    If MTtoORFMapCount > 0 Then
        ' Possibly shrink the arrays
        If MTtoORFMapCount < UBound(MTIDMap) Then
            ReDim Preserve MTIDMap(1 To MTtoORFMapCount)
            ReDim Preserve ORFIDMap(1 To MTtoORFMapCount)
            ReDim Preserve ORFRefNames(1 To MTtoORFMapCount)
        End If
        blnSuccess = True
    Else
        Erase MTIDMap
        Erase ORFIDMap
        Erase ORFRefNames
        blnSuccess = False
    End If
    
    Set rsMT_ORF_Map = Nothing
    
    LegacyDBLoadProteinData = blnSuccess
    Exit Function

err_LegacyDBLoadProteinData:

    strErrorMessage = "Error loading Protein data from tables " & TABLE_NAME_AMT_PROTEINS & " and " & TABLE_NAME_AMT_TO_PROTEIN_MAP & " in file: " & strLegacyDBFilePath & "; " & Err.Description & " (Error Number " & Trim(Err.Number) & ")"
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        AddToAnalysisHistory lngGelIndex, strErrorMessage
    Else
        MsgBox strErrorMessage, vbExclamation + vbOKOnly, glFGTU
    End If

    Resume exit_LegacyDBLoadProteinData

End Function

Public Sub RemoveAMT(ByVal Ind As Long, ByVal Scope As Integer)
'removes AMT reference from the gel; the only samtDef parameter
'that this function takes into account is SearchScope; that way
'we can test search with various parameters on small portion of file
'and later clean it if we dont want to keep the search results
Dim i As Long

Select Case Scope
Case glScope.glSc_All
  With GelData(Ind)
      If .CSLines > 0 Then
         For i = 1 To .CSLines
            If i Mod 1000 = 1 Then TraceLog 4, "RemoveAMT", "Calling CleanAMTRef .CSData(" & i & ")"
               ' Old Method:
               '' CleanAMTRef .CSData(i).MTID
               .CSData(i).MTID = ""
         Next i
      End If
      If .IsoLines > 0 Then
         For i = 1 To .IsoLines
            If i Mod 1000 = 1 Then TraceLog 4, "RemoveAMT", "Calling CleanAMTRef .IsoData(" & i & ")"
            
            ' Old Method:
            '' CleanAMTRef .IsoData(i).MTID
            
            ' New Method:
            .IsoData(i).MTID = ""
         Next i
      End If
  End With
Case glScope.glSc_Current
  With GelData(Ind)
    If .CSLines > 0 Then
       For i = 1 To .CSLines
        If i Mod 1000 = 1 Then TraceLog 4, "RemoveAMT", "Possibly calling CleanAMTRef .CSData(" & i & ")"
         If GelDraw(Ind).CSID(i) > 0 And GelDraw(Ind).CSR(i) > 0 Then
            CleanAMTRef .CSData(i).MTID
         End If
       Next i
    End If
    If .IsoLines > 0 Then
       For i = 1 To .IsoLines
        If i Mod 1000 = 1 Then TraceLog 4, "RemoveAMT", "Possibly calling CleanAMTRef .IsoData(" & i & ")"
         If GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0 Then
            CleanAMTRef .IsoData(i).MTID
         End If
       Next i
    End If
  End With
End Select
End Sub

Public Sub RemoveInternalStd(ByVal Ind As Long, ByVal Scope As Integer)
' Removes ISTD reference from the gel
' Scope can be glScope.glSc_All or glScope.glSc_Current

Dim i As Long

Select Case Scope
Case glScope.glSc_All
  With GelData(Ind)
      If .CSLines > 0 Then
         For i = 1 To .CSLines
             If i Mod 1000 = 1 Then TraceLog 4, "RemoveInternalStd", "Calling CleanInternalStdRef .CSData(" & i & ")"
             CleanInternalStdRef .CSData(i).MTID
         Next i
      End If
      If .IsoLines > 0 Then
         For i = 1 To .IsoLines
             If i Mod 1000 = 1 Then TraceLog 4, "RemoveInternalStd", "Calling CleanInternalStdRef .IsoData(" & i & ")"
             CleanInternalStdRef .IsoData(i).MTID
         Next i
      End If
  End With
Case glScope.glSc_Current
  With GelData(Ind)
    If .CSLines > 0 Then
       For i = 1 To .CSLines
         If i Mod 1000 = 1 Then TraceLog 4, "RemoveInternalStd", "Possibly calling CleanInternalStdRef .CSData(" & i & ")"
         If GelDraw(Ind).CSID(i) > 0 And GelDraw(Ind).CSR(i) > 0 Then
            CleanInternalStdRef .CSData(i).MTID
         End If
       Next i
    End If
    If .IsoLines > 0 Then
       For i = 1 To .IsoLines
         If i Mod 1000 = 1 Then TraceLog 4, "RemoveInternalStd", "Possibly calling CleanInternalStdRef .IsoData(" & i & ")"
         If GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0 Then
            CleanInternalStdRef .IsoData(i).MTID
         End If
       Next i
    End If
  End With
End Select
End Sub

' MonroeMod
' Additional Mod made 3/7/2003: Removed the redundant code in this function by
'  moving the samtDef.SkipReferenced check and the samtDef.SearchScope check
Public Function SearchAMT(ByVal Ind As Long, _
                          ByVal sExpr As String, frmCallingForm As VB.Form) As Long
'searches AMT database for MWs from GelData(ind) based
'on values in global variable samtDef arguments.
'To optimize for speed arrays loaded from the AMT table
'are actually searched rather than database recordsets
'SearchFlag determines which AMTs are included in search
Dim MinFN As Long
Dim MaxFN As Long
Dim ScanWidth As Long
Dim AMTRef As String
Dim IsoF As Integer     'Isotopic MW field; just shortcut
Dim i As Long
Dim blnProceed As Boolean

If Not GelData(Ind).CustomNETsDefined Then
    If Not InitExprEvaluator(sExpr) Then
       SearchAMT = -2
       Exit Function
    End If
End If

' MonroeMod
Dim strCaptionSaved As String
strCaptionSaved = frmCallingForm.Caption

With GelData(Ind)
   HitsCount = 0
   Set mwutSearch = New MWUtil
   If Not FillMWSearchObject(mwutSearch) Then GoTo err_SearchAMT
   IsoF = samtDef.MWField
   GetScanRange Ind, MinFN, MaxFN, ScanWidth
   If ScanWidth <= 0 And samtDef.NETTol >= 0 Then GoTo err_SearchAMT  'can not do it
   SetAMTSearchFlags samtDef.SearchFlag, aSearchFlag()
        If .CSLines > 0 Then
           For i = 1 To .CSLines
            ' MonroeMod Begin
            If i Mod 100 = 0 Then
                frmCallingForm.Caption = "Working: " & i & " / " & .CSLines
                DoEvents
            End If
            ' MonroeMod Finish
             
             If samtDef.SearchScope = glScope.glSc_All Or (GelDraw(Ind).CSID(i) > 0 And GelDraw(Ind).CSR(i) > 0) Then
                ' Proceed if using all the data, or if the ion is in the current scope
                If samtDef.SkipReferenced Then
                   blnProceed = Not IsAMTReferenced(.CSData(i).MTID)
                Else
                   blnProceed = True
                End If
                If blnProceed Then
                   If samtDef.NETTol >= 0 Then
                      Select Case samtDef.NETorRT
                      Case glAMT_NET
                        AMTRef = GetAMTReferenceMWNET(.CSData(i).AverageMW, NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), 0)
                      Case glAMT_RT_or_PNET
                        AMTRef = GetAMTReferenceMWRT(.CSData(i).AverageMW, NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), 0)
                      End Select
                   Else
                      AMTRef = GetAMTReferenceMW(.CSData(i).AverageMW, NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), 0)
                   End If
                   InsertBefore .CSData(i).MTID, AMTRef
                End If
             End If
           Next i
        End If
        If .IsoLines > 0 Then
           For i = 1 To .IsoLines
            ' MonroeMod Begin
            If i Mod 100 = 0 Then
                frmCallingForm.Caption = "Working: " & i & " / " & .IsoLines
                DoEvents
            End If
            ' MonroeMod Finish
            
             If samtDef.SearchScope = glScope.glSc_All Or (GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0) Then
                ' Proceed if using all the data, or if the ion is in the current scope
                If samtDef.SkipReferenced Then
                   blnProceed = Not IsAMTReferenced(.IsoData(i).MTID)
                Else
                   blnProceed = True
                End If
                If blnProceed Then
                   If samtDef.NETTol >= 0 Then
                      Select Case samtDef.NETorRT
                      Case glAMT_NET
                        AMTRef = GetAMTReferenceMWNET(GetIsoMass(.IsoData(i), IsoF), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), 0)
                      Case glAMT_RT_or_PNET
                        AMTRef = GetAMTReferenceMWRT(GetIsoMass(.IsoData(i), IsoF), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), 0)
                      End Select
                   Else
                      AMTRef = GetAMTReferenceMW(GetIsoMass(.IsoData(i), IsoF), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), 0)
                   End If
                   InsertBefore .IsoData(i).MTID, AMTRef
                End If
             End If
           Next i
        End If
   SearchAMT = HitsCount
   
exit_SearchAMT:
End With

' MonroeMod
frmCallingForm.Caption = strCaptionSaved

Set mwutSearch = Nothing
Exit Function

err_SearchAMT:
SearchAMT = -1
GoTo exit_SearchAMT
End Function

Public Sub SearchAMTComputeSLiCScores(ByRef lngCurrIDCnt As Long, _
                                      ByRef udtCurrIDMatches() As udtUMCMassTagRawMatches, _
                                      ByVal dblClassMass As Double, _
                                      ByVal dblMWTolFinal As Double, _
                                      ByVal dblNETTolFinal As Double, _
                                      ByVal eSearchRegionShape As srsSearchRegionShapeConstants, _
                                      ByVal blnUsingPrecomputedSTACScores As Boolean, _
                                      ByVal blnFilterUsingFinalTolerances As Boolean)
    Dim lngIndex As Long
    
    Dim dblMassStDevPPM As Double
    Dim dblMassStDevAbs As Double
    
    Dim dblNETStDevCombined As Double
    Dim dblNumeratorSum As Double
            
    Dim lngMassTagIndexOriginal As Long
    
    Dim lngNewIDCount As Long
    
    If lngCurrIDCnt = 0 Then Exit Sub
    
    ' Compute the match scores (aka SLiC scores)
    
On Error GoTo ComputeSLiCScoresErrorHandler
    
    If Not blnUsingPrecomputedSTACScores Then
    
        dblMassStDevPPM = glbPreferencesExpanded.SLiCScoreOptions.MassPPMStDev
        If dblMassStDevPPM <= 0 Then dblMassStDevPPM = 3
        
        dblMassStDevAbs = PPMToMass(dblMassStDevPPM, dblClassMass)
        If dblMassStDevAbs <= 0 Then
            Debug.Assert False
            LogErrors 0, "AMT.Bas->SearchAMTComputeSLiCScores", "dblMassStDevAbs was <= 0, which isn't allowed"
            dblMassStDevAbs = 0.003
        End If
        
        ' Compute the standarized squared distance and the numerator sum
        dblNumeratorSum = 0
        For lngIndex = 0 To lngCurrIDCnt - 1
            
            ' December 2005: .UseAMTNETStDev is now always forced to be false
    ''        If glbPreferencesExpanded.SLiCScoreOptions.UseAMTNETStDev Then
    ''            ' The NET StDev is computed by combining the default NETStDev value with the AMT's specific NETStDev
    ''            ' The combining is done by "adding in quadrature", which means to square each number, add together, and take the square root
    ''
    ''            If udtCurrIDMatches(lngIndex).IDIsInternalStd Then
    ''                ' Internal Standard match; Internal Standards do not have NETStDev values
    ''                dblNETStDevCombined = glbPreferencesExpanded.SLiCScoreOptions.NETStDev
    ''            Else
    ''                ' MT tag match
    ''                lngMassTagIndexOriginal = udtCurrIDMatches(lngIndex).IDIndexOriginal
    ''                dblNETStDevCombined = Sqr(glbPreferencesExpanded.SLiCScoreOptions.NETStDev ^ 2 + AMTData(lngMassTagIndexOriginal).NETStDev ^ 2)
    ''            End If
    ''
    ''        Else
                ' Simply use the default NETStDev value
                dblNETStDevCombined = glbPreferencesExpanded.SLiCScoreOptions.NETStDev
    ''        End If
            
            If dblNETStDevCombined <= 0 Then
                Debug.Assert False
                LogErrors 0, "AMT.Bas->SearchAMTComputeSLiCScores", "dblNETStDevCombined was <= 0, which isn't allowed"
                dblNETStDevCombined = 0.025
            End If
            
            With udtCurrIDMatches(lngIndex)
                .StandardizedSquaredDistance = .MassErr ^ 2 / dblMassStDevAbs ^ 2 + .NETErr ^ 2 / dblNETStDevCombined ^ 2
                
                .SLiCScoreNumerator = (1 / (dblMassStDevAbs * dblNETStDevCombined)) * Exp(-.StandardizedSquaredDistance / 2)
                
                dblNumeratorSum = dblNumeratorSum + .SLiCScoreNumerator
            End With
        Next lngIndex
        
        ' Compute the match score for each match
        For lngIndex = 0 To lngCurrIDCnt - 1
            With udtCurrIDMatches(lngIndex)
                If dblNumeratorSum > 0 Then
                    .StacOrSLiC = Round(.SLiCScoreNumerator / dblNumeratorSum, 5)
                Else
                    .StacOrSLiC = 0
                End If
            End With
        Next lngIndex
        
    End If
    
    If lngCurrIDCnt > 1 Then
        ' Sort by SLiCScore descending (need a custom sort routine since udtCurrIDMatches is a UDT)
        ShellSortCurrIDMatches lngCurrIDCnt, udtCurrIDMatches
    End If
    
    If lngCurrIDCnt > 0 Then
        ' Compute the DelSLiC value
        ' If there is only one match, then the DelSLiC value is 1
        ' If there is more than one match, then the highest scoring match gets a DelSLiC value,
        '  computed by subtracting the next lower scoring value from the highest scoring value;
        '  all other matches get a DelSLiC score of 0
        ' This allows one to quickly identify the LC-MS Features with a single match (DelSLiC = 1) or with a match
        '  distinct from other matches (DelSLiC > threshold)
        
        If lngCurrIDCnt > 1 Then
            udtCurrIDMatches(0).DelScore = (udtCurrIDMatches(0).StacOrSLiC - udtCurrIDMatches(1).StacOrSLiC)
            
            For lngIndex = 1 To lngCurrIDCnt - 1
                udtCurrIDMatches(lngIndex).DelScore = 0
            Next lngIndex
        Else
            udtCurrIDMatches(0).DelScore = 1
        End If
        
        If blnFilterUsingFinalTolerances Then
            ' Now filter the list using the tighter tolerances:
            '   MWTol is dblMWTolFinal and NET Tol is dblNETTolFinal
            ' Since we're shrinking the array, we can copy in place
            '
            ' When testing whether to keep the match or not, we're testing whether the match is
            '  in the ellipse or in the rectangle bounded by dblMWTolFinal and dblNETTolFinal
            ' Note that these are half-widths of the ellipse or rectangle
            lngNewIDCount = 0
            For lngIndex = 0 To lngCurrIDCnt - 1
                If TestPointInRegion(udtCurrIDMatches(lngIndex).NETErr, udtCurrIDMatches(lngIndex).MassErr, dblNETTolFinal, dblMWTolFinal, eSearchRegionShape) Then
                    udtCurrIDMatches(lngNewIDCount) = udtCurrIDMatches(lngIndex)
                    lngNewIDCount = lngNewIDCount + 1
                End If
            Next lngIndex
        Else
            lngNewIDCount = lngCurrIDCnt
        End If
    End If
 
    If lngNewIDCount = 0 Then
        lngCurrIDCnt = 0
        ReDim udtCurrIDMatches(0)
    ElseIf lngNewIDCount < lngCurrIDCnt Then
        lngCurrIDCnt = lngNewIDCount
        ReDim Preserve udtCurrIDMatches(lngNewIDCount - 1)
    End If
    
    Exit Sub
    
ComputeSLiCScoresErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "AMT.Bas->SearchAMTComputeSLiCScores"
    
End Sub

Public Sub SearchAMTDefineTolerances(ByVal lngGelIndex As Long, ByVal UMCIndex As Long, _
                                     ByRef udtAMTDef As SearchAMTDefinition, ByRef dblClassMass As Double, _
                                     ByRef MWTolAbsBroad As Double, ByRef NETTolBroad As Double, _
                                     ByRef MWTolAbsFinal As Double, ByRef NETTolFinal As Double)

    Const STDEV_SCALING_FACTOR As Integer = 2
    
    Dim MWTolPPMBroad As Double
    
    With GelUMC(lngGelIndex).UMCs(UMCIndex)
        dblClassMass = .ClassMW
        
        ' The weight to search on is the class mass, not each member's mass
        Select Case udtAMTDef.TolType
        Case gltPPM
            MWTolAbsFinal = dblClassMass * udtAMTDef.MWTol * glPPM
            MWTolPPMBroad = udtAMTDef.MWTol
        Case gltABS
            MWTolAbsFinal = udtAMTDef.MWTol
            If dblClassMass > 0 Then
                MWTolPPMBroad = udtAMTDef.MWTol / dblClassMass / glPPM
            Else
                MWTolPPMBroad = glbPreferencesExpanded.SLiCScoreOptions.MassPPMStDev
            End If
        Case Else
            Debug.Assert False
        End Select
    End With
    
    With glbPreferencesExpanded.SLiCScoreOptions
        If .AutoDefineSLiCScoreThresholds Then
            ' Define the Mass StDev (in ppm) using the narrow mass tolerance divided by 2 = STDEV_SCALING_FACTOR
            Select Case udtAMTDef.TolType
            Case gltPPM
                .MassPPMStDev = udtAMTDef.MWTol / STDEV_SCALING_FACTOR
            Case gltABS
                If dblClassMass > 0 Then
                    .MassPPMStDev = udtAMTDef.MWTol / dblClassMass / glPPM / STDEV_SCALING_FACTOR
                Else
                    .MassPPMStDev = 3
                End If
            Case Else
                Debug.Assert False
            End Select
            
            ' Define the Net StDev using the narrow NET tolerance divided by 2 = STDEV_SCALING_FACTOR
            .NETStDev = udtAMTDef.NETTol / STDEV_SCALING_FACTOR
        End If
        
        If MWTolPPMBroad < .MassPPMStDev * .MaxSearchDistanceMultiplier * STDEV_SCALING_FACTOR Then
            MWTolPPMBroad = .MassPPMStDev * .MaxSearchDistanceMultiplier * STDEV_SCALING_FACTOR
        End If
        NETTolBroad = .NETStDev * .MaxSearchDistanceMultiplier * STDEV_SCALING_FACTOR
        If NETTolBroad < udtAMTDef.NETTol Then NETTolBroad = udtAMTDef.NETTol
    End With
    
    NETTolFinal = udtAMTDef.NETTol
    
    ' Convert from PPM to Absolute mass
    MWTolAbsBroad = dblClassMass * MWTolPPMBroad * glPPM

End Sub
 
 Private Sub ShellSortCurrIDMatches(ByRef lngCurrIDCnt As Long, ByRef udtCurrIDMatches() As udtUMCMassTagRawMatches)
    Dim lngLowIndex As Long
    Dim lngHighIndex As Long
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim udtCompareVal As udtUMCMassTagRawMatches

On Error GoTo ShellSortCurrIDMatchesErrorHandler

' sort array[lngLowIndex..lngHighIndex]

    lngLowIndex = 0
    lngHighIndex = lngCurrIDCnt - 1
    
    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            udtCompareVal = udtCurrIDMatches(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If udtCurrIDMatches(lngIndexCompare).StacOrSLiC > udtCompareVal.StacOrSLiC Then Exit For
                udtCurrIDMatches(lngIndexCompare + lngIncrement) = udtCurrIDMatches(lngIndexCompare)
            Next lngIndexCompare
            udtCurrIDMatches(lngIndexCompare + lngIncrement) = udtCompareVal
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop

    Exit Sub

ShellSortCurrIDMatchesErrorHandler:
    Debug.Assert False

End Sub

Private Function TestPointInRegion(ByVal dblPointX As Double, ByVal dblPointY As Double, ByVal dblXTol As Double, ByVal dblYTol As Double, ByVal eSearchRegionShape As srsSearchRegionShapeConstants) As Boolean
    ' Tests whether a point is within the rectangular or the elliptical region defined by dblXTol and dblYTol
    '
    ' The equation for the points along the edge of an ellipse is x^2/a^2 + y^2/b^2 = 1 where a and b are
    ' the half-widths of the ellipse and x and y are the coordinates of each point on the ellipse's perimeter
    '
    ' If blnUseEllipticalBoundary = True, then this function takes x, y, a, and b as inputs
    '  and computes the result of this equation.  If the result is <= 1, then the point
    '  at x,y is inside the ellipse
    
On Error GoTo TestPointInEllipseErrorHandler

    If eSearchRegionShape = srsSearchRegionShapeConstants.srsRectangular Then
        ' Test if point is within the rectangle bounded by the tolerances
        If Abs(dblPointX) <= dblXTol And Abs(dblPointY) <= dblYTol Then
            TestPointInRegion = True
        Else
            TestPointInRegion = False
        End If
    Else
        ' Assume srsSearchRegionShapeConstants.srsElliptical
        ' Test if point is within the ellipse bounded by the tolerances
        If dblPointX ^ 2 / dblXTol ^ 2 + dblPointY ^ 2 / dblYTol ^ 2 <= 1 Then
            TestPointInRegion = True
        Else
            TestPointInRegion = False
        End If
    End If

Exit Function

TestPointInEllipseErrorHandler:
' Error; return false
TestPointInRegion = False

End Function

Private Function GetAMTReferenceMW(ByVal MW As Double, _
                                   ByVal NETRT As Double, _
                                   ByVal Delta As Long, _
                                   Optional ByVal blnStoreAbsoluteValueOfError As Boolean = False) As String
'returns AMT reference string based on MW and samtDef
'this function is called only from SearchAMT function
'NETRT is here used only to generate statistic(NET & RT)
'and not as a search criteria
Dim AMTRef As String
Dim MWTolRef As Double
Dim sMWTolRef As String
Dim FirstInd As Long
Dim LastInd As Long
Dim AbsTol As Double
Dim i As Long
On Error GoTo exit_GetAMTReferenceMW

Select Case samtDef.TolType
Case gltPPM
    AbsTol = MW * samtDef.MWTol * glPPM
Case gltABS
    AbsTol = samtDef.MWTol
Case Else
    Debug.Assert False
End Select
If mwutSearch.FindIndexRange(MW, AbsTol, FirstInd, LastInd) Then
   For i = FirstInd To LastInd
     If IsGoodAMTFlag(AMTData(i).flag) Then
        HitsCount = HitsCount + 1
        If blnStoreAbsoluteValueOfError Then
            MWTolRef = Abs(MW - AMTData(i).MW)
        Else
            MWTolRef = MW - AMTData(i).MW
        End If
        sMWTolRef = MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd
        'put AMT ID and actual errors
        AMTRef = AMTRef & AMTMark & Trim(AMTData(i).ID) & sMWTolRef
        If samtDef.SaveNCnt Then AMTRef = AMTRef & MTNCntMark & AMTData(i).CNT_N & MTEndMark
        If Delta > 0 Then AMTRef = AMTRef & MTDltMark & Delta
        AMTRef = AMTRef & glARG_SEP & Chr$(32)
        'do statistics
        AMTHits(i) = AMTHits(i) + 1
        AMTMWErr(i) = AMTMWErr(i) + MWTolRef
        Select Case samtDef.NETorRT
        Case glAMT_NET
             ' 7/26/2004 MEM - Reversed the order of subtraction to be Observed - Database, consistent with the MWTolRef calculation above
             AMTNETErr(i) = AMTNETErr(i) + (NETRT - AMTData(i).NET)
        Case glAMT_RT_or_PNET
             If (AMTData(i).PNET >= 0) Then    'there are some negative RTs
                AMTNETErr(i) = AMTNETErr(i) + (NETRT - AMTData(i).PNET)
             End If
        End Select
        If NETRT < AMTNETMin(i) Then AMTNETMin(i) = NETRT
        If NETRT > AMTNETMax(i) Then AMTNETMax(i) = NETRT
     End If
   Next i
End If
'If Len(AMTRef) <= 0 Then AMTRef = AMTMark & NoHarvest & glARG_SEP & Chr$(32)
exit_GetAMTReferenceMW:
GetAMTReferenceMW = AMTRef
End Function

Private Function GetAMTReferenceMWNET(ByVal MW As Double, _
                                      ByVal NET As Double, _
                                      ByVal Delta As Long, _
                                      Optional ByVal blnStoreAbsoluteValueOfError As Boolean = False) As String
'returns AMT reference string based on MW, NET and samtDef
'this function is called only from SearchAMT function
Dim AMTRef As String
Dim MWTolRef As Double
Dim sMWTolRef As String
Dim NETTolRef As Double
Dim sNETTolRef As String
Dim FirstInd As Long
Dim LastInd As Long
Dim AbsTol As Double
Dim i As Long
On Error GoTo exit_GetAMTReferenceMWNET

Select Case samtDef.TolType
Case gltPPM
    AbsTol = MW * samtDef.MWTol * glPPM
Case gltABS
    AbsTol = samtDef.MWTol
Case Else
    Debug.Assert False
End Select
If mwutSearch.FindIndexRange(MW, AbsTol, FirstInd, LastInd) Then
   For i = FirstInd To LastInd
     If ((Abs(NET - AMTData(i).NET) <= samtDef.NETTol) And (IsGoodAMTFlag(AMTData(i).flag))) Then
        HitsCount = HitsCount + 1
        If blnStoreAbsoluteValueOfError Then
            MWTolRef = Abs(MW - AMTData(i).MW)
        Else
            MWTolRef = MW - AMTData(i).MW
        End If
        sMWTolRef = MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd
        ' 7/26/2004 MEM - Reversed the order of subtraction to be Observed - Database, consistent with the MWTolRef calculation above
        NETTolRef = (NET - AMTData(i).NET)
        sNETTolRef = NETErrMark & Format$(NETTolRef, "0.000") & NETErrEnd
        'put AMT ID and actual errors
        AMTRef = AMTRef & AMTMark & Trim(AMTData(i).ID) & sMWTolRef & sNETTolRef
        If samtDef.SaveNCnt Then AMTRef = AMTRef & MTNCntMark & AMTData(i).CNT_N & MTEndMark
        If Delta > 0 Then AMTRef = AMTRef & MTDltMark & Delta
        AMTRef = AMTRef & glARG_SEP & Chr$(32)
        'do statistics
        AMTHits(i) = AMTHits(i) + 1
        AMTMWErr(i) = AMTMWErr(i) + MWTolRef
        AMTNETErr(i) = AMTNETErr(i) + NETTolRef
        If NET < AMTNETMin(i) Then AMTNETMin(i) = NET
        If NET > AMTNETMax(i) Then AMTNETMax(i) = NET
     End If
   Next i
End If
'If Len(AMTRef) = 0 Then AMTRef = AMTMark & NoHarvest & glARG_SEP & Chr$(32)
exit_GetAMTReferenceMWNET:
GetAMTReferenceMWNET = AMTRef
End Function


Private Function GetAMTReferenceMWRT(ByVal MW As Double, _
                                     ByVal RT As Double, _
                                     ByVal Delta As Double, _
                                     Optional ByVal blnStoreAbsoluteValueOfError As Boolean = False) As String
'returns AMT reference string based on MW, RT and samtDef
'this function is called only from SearchAMT function
Dim AMTRef As String
Dim MWTolRef As Double
Dim sMWTolRef As String
Dim RTolRef As Double
Dim sRTolRef As String
Dim FirstInd As Long
Dim LastInd As Long
Dim AbsTol As Double
Dim i As Long
On Error GoTo exit_GetAMTReferenceMWRT

Select Case samtDef.TolType
Case gltPPM
    AbsTol = MW * samtDef.MWTol * glPPM
Case gltABS
    AbsTol = samtDef.MWTol
Case Else
    Debug.Assert False
End Select
If mwutSearch.FindIndexRange(MW, AbsTol, FirstInd, LastInd) Then
   For i = FirstInd To LastInd
     If ((Abs(RT - AMTData(i).PNET) <= samtDef.NETTol) And (IsGoodAMTFlag(AMTData(i).flag))) Then
        HitsCount = HitsCount + 1
        If blnStoreAbsoluteValueOfError Then
            MWTolRef = Abs(MW - AMTData(i).MW)
        Else
            MWTolRef = MW - AMTData(i).MW
        End If
        sMWTolRef = MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd
        RTolRef = (AMTData(i).PNET - RT)
        sRTolRef = RTErrMark & Format$(RTolRef, "0.000") & RTErrEnd
        'put AMT ID and actual errors
        AMTRef = AMTRef & AMTMark & Trim(AMTData(i).ID) & sMWTolRef & sRTolRef
        If samtDef.SaveNCnt Then AMTRef = AMTRef & MTNCntMark & AMTData(i).CNT_N & MTEndMark
        If Delta > 0 Then AMTRef = AMTRef & MTDltMark & Delta
        AMTRef = AMTRef & glARG_SEP & Chr$(32)
        'do statistics; AMTs with negative RT swill not be identified anyways
        AMTHits(i) = AMTHits(i) + 1
        AMTMWErr(i) = AMTMWErr(i) + MWTolRef
        AMTNETErr(i) = AMTNETErr(i) + RTolRef
        If RT < AMTNETMin(i) Then AMTNETMin(i) = RT
        If RT > AMTNETMax(i) Then AMTNETMax(i) = RT
     End If
   Next i
End If
'If Len(AMTRef) = 0 Then AMTRef = AMTMark & NoHarvest & glARG_SEP & Chr$(32)
exit_GetAMTReferenceMWRT:
GetAMTReferenceMWRT = AMTRef
End Function

Private Function GetAMTRefFromString(ByVal S As String, _
                                     ByVal StartPos As Integer, _
                                     Optional ByRef AMTMatchStart As Integer) As String
'extracts and returns first AMT reference from string s
'starting from StartPos, returns empty string if not found
'updates AMTMatchStart to the position of the first character of the match in s
Dim AMTPos As Integer
Dim AMTEnd As Integer
On Error Resume Next
AMTPos = InStr(StartPos, S, AMTMark)
If AMTPos > 0 Then
   AMTEnd = InStr(AMTPos, S, glARG_SEP)
   If AMTEnd > 0 Then
      GetAMTRefFromString = Mid$(S, AMTPos, AMTEnd - AMTPos)
   Else
      GetAMTRefFromString = Mid$(S, AMTPos)
   End If
   AMTMatchStart = AMTPos
Else
   GetAMTRefFromString = ""
   AMTMatchStart = 0
End If
End Function

Private Function GetInternalStdRefFromString(ByVal S As String, _
                                        ByVal StartPos As Integer, _
                                        Optional ByRef MatchStart As Integer) As String
'extracts and returns first ISTD reference from string s
'starting from StartPos, returns empty string if not found
'updates MatchStart to the position of the first character of the match in s
Dim ISTDPos As Integer
Dim ISTDEnd As Integer
On Error Resume Next
ISTDPos = InStr(StartPos, S, INT_STD_MARK)
If ISTDPos > 0 Then
   ISTDEnd = InStr(ISTDPos, S, glARG_SEP)
   If ISTDEnd > 0 Then
      GetInternalStdRefFromString = Mid$(S, ISTDPos, ISTDEnd - ISTDPos)
   Else
      GetInternalStdRefFromString = Mid$(S, ISTDPos)
   End If
   MatchStart = ISTDPos
Else
   GetInternalStdRefFromString = ""
   MatchStart = 0
End If
End Function

Public Function GetInternalStdRefFromString2(ByVal S As String, _
                                        ByRef Refs() As String) As Long
'Fills array Refs with Internal Standard references and returns length of array
'Note that array is 1-based
'Similar to GetAMTRefFromString2

Dim Cnt  As Long
Dim IntStdRef As String
Dim StartPos As Integer
Dim IntStdMatchStart As Integer
Dim Done As Boolean

Cnt = 0
If Len(S) > 0 Then
   ReDim Refs(1 To Len(S))
   StartPos = 1
   Do Until Done
      IntStdRef = GetInternalStdRefFromString(S, StartPos, IntStdMatchStart)
      If Len(IntStdRef) > 0 Then
         Cnt = Cnt + 1
         Refs(Cnt) = IntStdRef
         StartPos = IntStdMatchStart + Len(IntStdRef)
         If StartPos > Len(S) Then
            Done = True
         End If
      Else
         Done = True
      End If
   Loop
End If
If Cnt > 0 Then
   ReDim Preserve Refs(1 To Cnt)
Else
   Erase Refs
End If
GetInternalStdRefFromString2 = Cnt
End Function

Public Function IsAMTReferenced(S As String) As Boolean
''returns True if string s contains AMT reference other than Not Found
''If IsNull(S) Then
''   IsAMTReferenced = False
''Else
   If InStr(1, S, AMTMark) > 0 Then
      'if reference is "Not Found" consider it not referenced
      If InStr(1, GetAMTRefFromString(S, 1), NoHarvest) > 0 Then
         IsAMTReferenced = False
      Else
         IsAMTReferenced = True
      End If
   Else
      IsAMTReferenced = False
   End If
''End If
End Function

Public Function IsInternalStdReferenced(S As String) As Boolean
''returns True if string s contains an INT_STD_MARK reference
''If IsNull(S) Then
''   IsInternalStdReferenced = False
''Else
   If InStr(1, S, INT_STD_MARK) > 0 Then
      IsInternalStdReferenced = True
   Else
      IsInternalStdReferenced = False
   End If
''End If
End Function

Public Function IsAMTReferencedByUMC(udtUMC As udtUMCType, lngGelIndex As Long) As Boolean
    Dim i As Long
    Dim blnAMTMatchPresent As Boolean
    
    With udtUMC
        blnAMTMatchPresent = False
        For i = 0 To .ClassCount - 1
            Select Case .ClassMType(i)
            Case glCSType
                If IsAMTReferenced(GelData(lngGelIndex).CSData(.ClassMInd(i)).MTID) Then
                    blnAMTMatchPresent = True
                    Exit For
                End If
            Case glIsoType
                If IsAMTReferenced(GelData(lngGelIndex).IsoData(.ClassMInd(i)).MTID) Then
                    blnAMTMatchPresent = True
                    Exit For
                End If
            End Select
        Next i
    End With
    
    IsAMTReferencedByUMC = blnAMTMatchPresent
End Function

Public Sub InitializeAMTDataEntry(ByRef udtAMTDataPoint As udtAMTDataType, Optional NET_VALUE_IF_NULL As Single = -100000)
    udtAMTDataPoint.ID = 0
    udtAMTDataPoint.flag = 0
    udtAMTDataPoint.MW = 0
    udtAMTDataPoint.NET = 0
    udtAMTDataPoint.MSMSObsCount = 1
    udtAMTDataPoint.NETStDev = 0
    udtAMTDataPoint.NETCount = 0
    udtAMTDataPoint.PNET = NET_VALUE_IF_NULL
    udtAMTDataPoint.CNT_N = -1
    udtAMTDataPoint.CNT_Cys = -1
    udtAMTDataPoint.HighNormalizedScore = 0
    udtAMTDataPoint.HighDiscriminantScore = 0
    udtAMTDataPoint.PeptideProphetProbability = 0
    udtAMTDataPoint.Sequence = ""
End Sub


Public Function IsAMTMatchInherited(S As String) As Boolean
'returns True if string s contains AMTMatchInheritedMark
    
If IsNull(S) Then
    IsAMTMatchInherited = False
Else
    If InStr(1, S, AMTMatchInheritedMark) > 0 Then
        IsAMTMatchInherited = True
    Else
        IsAMTMatchInherited = False
    End If
End If
    
End Function

Private Sub CleanAMTRef(S As Variant)
Dim sTmp As String
Dim AMTRef As String
Dim Done As Boolean
On Error GoTo err_CleanAMTRef

If Not IsNull(S) Then
    sTmp = CStr(S)  'this will trigger error for Null
    If Len(sTmp) > 0 Then
       Do Until Done
          AMTRef = GetAMTRefFromString(sTmp, 1)
          If Len(AMTRef) > 0 Then
             Remove1stSubstring sTmp, AMTRef
             If Len(sTmp) = 0 Then Done = True
          Else
             Done = True
          End If
       Loop
       S = sTmp
    End If
End If

err_CleanAMTRef:
End Sub

Private Sub CleanInternalStdRef(S As Variant)
Dim sTmp As String
Dim IntStdRef As String
Dim Done As Boolean
On Error GoTo err_CleanInternalStdRef

If Not IsNull(S) Then
    sTmp = CStr(S)  'this will trigger error for Null
    If Len(sTmp) > 0 Then
       Do Until Done
          IntStdRef = GetInternalStdRefFromString(sTmp, 1)
          If Len(IntStdRef) > 0 Then
             Remove1stSubstring sTmp, IntStdRef
          Else
             Done = True
          End If
       Loop
       S = sTmp
    End If
End If

err_CleanInternalStdRef:
End Sub

Public Function GetAMTRefFromString1(ByVal S As String, _
                                     ByRef aAMT() As String) As Long
'fills array aAMT with AMT IDs and returns number of it
'Note that array aAMT is 1-based
'Note that S contains both AMT IDs and mass error information
'In contrast, aAMT() will just contain the AMT ID values
Dim Cnt As Long
Dim AMTRef As String
Dim StartPos As Integer
Dim AMTMatchStart As Integer
Dim Done As Boolean
Dim RefStringLength As Integer
Dim PosAdjust As Integer

RefStringLength = Len(S)
PosAdjust = Len(AMTMark) - 1

If RefStringLength > 0 Then
   ReDim aAMT(1 To 2)
   Cnt = 0
   StartPos = 1
   Do Until Done
      AMTRef = GetAMTRefFromString(S, StartPos, AMTMatchStart)
      If Len(AMTRef) > 0 Then
         Cnt = Cnt + 1
         If Cnt > UBound(aAMT) Then
            ReDim Preserve aAMT(1 To UBound(aAMT) * 2)
         End If
        
         aAMT(Cnt) = GetIDFromString(AMTRef, AMTMark, AMTIDEnd)
         StartPos = AMTMatchStart + Len(AMTRef)
         If StartPos >= RefStringLength - PosAdjust Then
            Done = True
         End If
      Else
         Done = True
      End If
   Loop
End If
If Cnt > 0 Then
   ReDim Preserve aAMT(1 To Cnt)
Else
   Erase aAMT
End If
GetAMTRefFromString1 = Cnt
End Function

Public Function GetAMTSearchDefDesc() As String
Dim sTmp As String
On Error GoTo exit_GetAMTSearchDefDesc
sTmp = "AMT Search Definition:" & vbCrLf
With samtDef
    Select Case .SearchScope
    Case glScope.glSc_All
         sTmp = sTmp & "AMT search on all data points." & vbCrLf
    Case glScope.glSc_Current
         sTmp = sTmp & "AMT search on data points currently in scope." & vbCrLf
    End Select
    If .SkipReferenced Then
         sTmp = sTmp & "Previously AMT referenced data points not included in search(previous references not included in statistics)." & vbCrLf
    Else
         sTmp = sTmp & "Previously AMT referenced data points included in search(previous references not included in statistics)." & vbCrLf
    End If
    Select Case .MWField
    Case 6
         sTmp = sTmp & "MW field(Isotopic): Average" & vbCrLf
    Case 7
         sTmp = sTmp & "MW field(Isotopic): Monoisotopic" & vbCrLf
    Case 8
         sTmp = sTmp & "MW field(Isotopic): The Most Abundant" & vbCrLf
    End Select
    Select Case .TolType
    Case gltPPM
         sTmp = sTmp & "MW Tolerance: " & .MWTol & " ppm" & vbCrLf
    Case gltABS
         sTmp = sTmp & "MW Tolerance: " & .MWTol & " Da" & vbCrLf
    Case Else
        Debug.Assert False
    End Select
    Select Case .NETorRT
    Case glAMT_NET
        sTmp = sTmp & "NET Calculation Formula: " & .Formula & vbCrLf
        If .NETTol < 0 Then
           sTmp = sTmp & "NET Tolerance: Not used as criteria in search."
        Else
           sTmp = sTmp & "NET Tolerance: " & .NETTol & vbCrLf
        End If
    Case glAMT_RT_or_PNET
        sTmp = sTmp & "RT Calculation Formula: " & .Formula & vbCrLf
        If .NETTol < 0 Then
           sTmp = sTmp & "RT Tolerance: Not used as criteria in search."
        Else
           sTmp = sTmp & "RT Tolerance: " & .NETTol & vbCrLf
        End If
    End Select
    If .SearchFlag > 0 Then
         sTmp = sTmp & "Search limited by condition: " & .SearchFlag & vbCrLf
    Else
         sTmp = sTmp & "Search over all AMTs found in database." & vbCrLf
    End If
End With
exit_GetAMTSearchDefDesc:
GetAMTSearchDefDesc = sTmp
End Function

Public Function LookupResidueOccurrence(ByVal lngAMTID As Long, ByVal strResidues As String) As Integer
    ' This function now handles multiple residues in strResidues
    
    Dim strSequence As String
    Dim intIndex As Integer
    Dim intResidueCount As Integer
    Static blnUserWarnedOfError As Boolean
    
    On Error GoTo LookupResidueOccurrenceErrorHandler

    ' Counts number of times that strResidues is present in AMTData(lngAMTID).Sequence
    strSequence = AMTData(lngAMTID).Sequence
    
    intResidueCount = 0
    For intIndex = 1 To Len(strResidues)
        intResidueCount = intResidueCount + AACount(strSequence, Mid(strResidues, intIndex, 1))
    Next intIndex
    
    LookupResidueOccurrence = intResidueCount
    
    Exit Function

LookupResidueOccurrenceErrorHandler:
    If Not blnUserWarnedOfError Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Error occurred when looking up the residue count in the given MT tag sequence (lngAMTID = " & lngAMTID & "):" & Err.Description, vbExclamation Or vbOKOnly, "Error"
        End If
        blnUserWarnedOfError = True
    End If
    LookupResidueOccurrence = -1
    
End Function

Private Function NET_RT(ByVal lngGelIndex As Long, ByVal FN As Long, ByVal MinFN As Long, ByVal MaxFN As Long) As Double
    If GelData(lngGelIndex).CustomNETsDefined Then
        NET_RT = ScanToGANET(lngGelIndex, FN)
    Else
        'this function does not care are we using NET or RT
        VarVals(1) = FN
        VarVals(2) = MinFN
        VarVals(3) = MaxFN
        NET_RT = MyExprEva.ExprVal(VarVals())
    End If
End Function

Private Function InitExprEvaluator(ByVal sExpr As String) As Boolean
On Error Resume Next
Set MyExprEva = New ExprEvaluator
With MyExprEva
    .Vars.add 1, "FN"
    .Vars.add 2, "MinFN"
    .Vars.add 3, "MaxFN"
    .Expr = sExpr
    InitExprEvaluator = .IsExprValid
    ReDim VarVals(1 To 3)
End With
End Function

Public Function GetAMTRefFromString2(ByVal S As String, _
                                     ByRef aAMT() As String) As Long
'Fills array aAMT with AMT references and returns number of it
'Note that array aAMT is 1-based
'Difference with GetAMTRefFromString1 is that here is returned
'the whole reference with errors; and Not founds are not counted
Dim Cnt  As Long
Dim AMTRef As String
Dim StartPos As Integer
Dim AMTMatchStart As Integer
Dim Done As Boolean

Dim RefStringLength As Integer
Dim PosAdjust As Integer

RefStringLength = Len(S)
PosAdjust = Len(AMTMark) - 1

Cnt = 0
If RefStringLength > 0 Then
   ReDim aAMT(1 To 4)
   StartPos = 1
   Do Until Done
      AMTRef = GetAMTRefFromString(S, StartPos, AMTMatchStart)
      If Len(AMTRef) > 0 Then
         'do not take empty(Not found) reference
         If InStr(1, AMTRef, NoHarvest) <= 0 Then
            Cnt = Cnt + 1
            If Cnt > UBound(aAMT) Then
               ReDim Preserve aAMT(1 To UBound(aAMT) * 2)
            End If
            aAMT(Cnt) = AMTRef
         End If
         StartPos = AMTMatchStart + Len(AMTRef)
         If StartPos > RefStringLength - PosAdjust Then
            Done = True
         End If
      Else
         Done = True
      End If
   Loop
End If
If Cnt > 0 Then
   ReDim Preserve aAMT(1 To Cnt)
Else
   Erase aAMT
End If
GetAMTRefFromString2 = Cnt
End Function

Public Function GetMWErrFromString(ByVal S As String) As String
'returns MW error from AMT string (always in ppm)
Dim Pos1 As Integer, Pos2 As Integer
Pos1 = InStr(1, S, MWErrMark)
If Pos1 > 0 Then
   Pos1 = Pos1 + Len(MWErrMark)
   Pos2 = InStr(Pos1, S, MWErrEnd)
   If Pos2 > 0 Then GetMWErrFromString = Mid$(S, Pos1, Pos2 - Pos1)
Else
    GetMWErrFromString = ""
End If
End Function

Public Function GetETErrFromString(ByVal S As String) As String
'-------------------------------------------------------------
'returns ET error from AMT string; there are different elution
'measurements but suggestion is to always use generic ET mark
'-------------------------------------------------------------
Dim Pos1 As Integer, Pos2 As Integer
Pos1 = InStr(1, S, ETErrMark)
If Pos1 > 0 Then
   Pos1 = Pos1 + Len(ETErrMark)
   Pos2 = InStr(Pos1, S, ETErrEnd)
   If Pos2 > 0 Then GetETErrFromString = Mid$(S, Pos1, Pos2 - Pos1)
Else
    GetETErrFromString = ""
End If
End Function

''Public Function GetNETErrFromString(ByVal S As String) As String
'''-------------------------------------------------------------
'''returns NET error from AMT string
'''-------------------------------------------------------------
''Dim Pos1 As Integer, Pos2 As Integer
''Pos1 = InStr(1, S, NETErrMark)
''If Pos1 > 0 Then
''   Pos1 = Pos1 + Len(NETErrMark)
''   Pos2 = InStr(Pos1, S, NETErrEnd)
''   If Pos2 > 0 Then GetNETErrFromString = Mid$(S, Pos1, Pos2 - Pos1)
''Else
''    GetNETErrFromString = ""
''End If
''End Function

Public Function GetSLiCFromString(ByVal S As String) As String
'-------------------------------------------------------------
'returns SLiC Score from AMT string
'-------------------------------------------------------------
Dim Pos1 As Integer, Pos2 As Integer
Pos1 = InStr(1, S, MTSLiCMark)
If Pos1 > 0 Then
   Pos1 = Pos1 + MTSLiCMarkLength
   Pos2 = InStr(Pos1, S, MTSLiCEnd)
   If Pos2 > 0 Then GetSLiCFromString = Mid$(S, Pos1, Pos2 - Pos1)
Else
    GetSLiCFromString = ""
End If

End Function

'''Public Function GetAMTBestErrIndex(AMTs() As String) As Long
''''returns index with best(smallest) error; -1 on any error
'''Dim BestAMTErr As Double
'''Dim BestAMTErrInd As Long
'''Dim sAMTMWErr As String
'''Dim AMTMWErr As Double
'''Dim i As Long
'''On Error GoTo exit_GetAMTBestErrIndex
'''
'''BestAMTErrInd = -1
'''BestAMTErr = glHugeOverExp
'''For i = 1 To UBound(AMTs)
'''    sAMTMWErr = GetMWErrFromString(AMTs(i))
'''    If IsNumeric(sAMTMWErr) Then
'''       AMTMWErr = CDbl(sAMTMWErr)
'''       If AMTMWErr < BestAMTErr Then
'''          BestAMTErr = AMTMWErr
'''          BestAMTErrInd = i
'''       End If
'''    End If
'''Next i
'''
'''exit_GetAMTBestErrIndex:
'''GetAMTBestErrIndex = BestAMTErrInd
'''End Function

Private Sub SetAMTSearchFlags(ByVal flag As Integer, _
                             ByRef aFlags() As Boolean)
Dim pos As Integer

ReDim aFlags(-1 To 10)
If flag > 0 Then
   For pos = 0 To 10
       If flag Mod 2 > 0 Then aFlags(pos) = True
       flag = flag \ 2      'integer division
   Next pos
Else
   aFlags(-1) = True
End If
End Sub

Private Function GetLegacyDBRequiredMTTableFields(ByVal eDBGeneration As dbgDatabaseGenerationConstants) As String
    Dim strMessage As String
    
    If eDBGeneration = dbgMTSOffline Then
        strMessage = "The [" & TABLE_NAME_T_MASS_TAGS & "] table should contain the fields: " & DB_FIELD_TMASSTAGS_MASS_TAG_ID & " and " & DB_FIELD_TMASSTAGS_MW & ".  " & _
                     "It can optionally contain the fields: " & DB_FIELD_TMASSTAGS_PEPTIDE & ", " & DB_FIELD_TMASSTAGS_MSMSObsCount & ", " & DB_FIELD_TMASSTAGS_HighNormalizedScore & ", " & DB_FIELD_TMASSTAGS_HighDiscriminantScore & ", and " & DB_FIELD_TMASSTAGS_PeptideProphetProbability & "." & _
                     "In addition, the [" & TABLE_NAME_T_MASS_TAGS_NET & "] table should contain the fields: " & DB_FIELD_TMASSTAGS_MASS_TAG_ID & " and " & DB_FIELD_TMTNET_NET & ", plus optionally " & DB_FIELD_TMTNET_PNET & " and " & DB_FIELD_TMTNET_STDEV & ".  "
    Else
        strMessage = "The [" & TABLE_NAME_AMT & "] table should contain the fields: " & DB_FIELD_AMT_NEW_ID & ", " & DB_FIELD_AMT_MW & ", " & DB_FIELD_AMT_NET & ", " & DB_FIELD_AMT_Status & ", and " & DB_FIELD_AMT_RETENTION & " or " & DB_FIELD_AMT_PNET & ".  " & _
                     "It can optionally contain the fields: " & DB_FIELD_AMT_PEPTIDE & ", " & DB_FIELD_AMT_MSMSObsCount & ", " & DB_FIELD_AMT_HighNormalizedScore & ", " & DB_FIELD_AMT_HighDiscriminantScore & ", " & DB_FIELD_AMT_PeptideProphetProbability & ", " & DB_FIELD_AMT_NitrogenAtom & ", and " & DB_FIELD_AMT_CysCount & "."
    End If
    
    GetLegacyDBRequiredMTTableFields = strMessage
End Function

Private Function GetLegacyDBRequiredProteinTableFields(ByVal eDBGeneration As dbgDatabaseGenerationConstants) As String
    Dim strMessage As String
    
    If eDBGeneration = dbgMTSOffline Then
        strMessage = "The [" & TABLE_NAME_T_PROTEINS & "] table should contain the fields " & DB_FIELD_TPROTEINS_Protein_ID & " and " & DB_FIELD_TPROTEINS_Protein_Name & ".  " & _
                     "The [" & TABLE_NAME_T_MASS_TAG_TO_PROTEIN_MAP & "] table should contain the fields " & DB_FIELD_TMASSTAGS_MASS_TAG_ID & " and " & DB_FIELD_TPROTEINS_Protein_ID & "."
    Else
        strMessage = "The [" & TABLE_NAME_AMT_PROTEINS & "] table should contain the fields " & DB_FIELD_PROTEIN_Protein_ID & " and " & DB_FIELD_PROTEIN_Protein_Name & ".  " & _
                     "The [" & TABLE_NAME_AMT_TO_PROTEIN_MAP & "] table should contain the fields " & DB_FIELD_PROTEIN_AMT_ID & " and " & DB_FIELD_PROTEIN_Protein_ID & "."
    End If
    
    GetLegacyDBRequiredProteinTableFields = strMessage
End Function

Private Function IsGoodAMTFlag(ByVal AMTFlag As Integer) As Boolean
'compares AMTFlag with SearchFlag and returnes True if AMTFlag
'satisfies SearchFlag conditions; SearchFlag array needs to be
'set prior to call to this function
Dim AMTF() As Boolean
Dim i As Integer
SetAMTSearchFlags AMTFlag, AMTF()
If aSearchFlag(-1) Then     'no search conditions
   IsGoodAMTFlag = True
Else
   For i = 1 To 7                       'if for any of Flags 1 to 7
       If aSearchFlag(i) Then           'SearchFlag is set and AMT is
          If Not AMTF(i) Then           'not AMT is not good enough
             IsGoodAMTFlag = False
             Exit Function
          End If
       End If
   Next i
   If aSearchFlag(0) Then                   'high accuracy AMT could be
      If (Not (AMTF(0) Or AMTF(1))) Then    'also marked as high accuracy
         IsGoodAMTFlag = False              'and NET condition
         Exit Function
      End If
   End If
   If aSearchFlag(8) Then
      If (Not (AMTF(8) Or AMTF(9) Or AMTF(10))) Then
         IsGoodAMTFlag = False
         Exit Function
      End If
   End If
   If aSearchFlag(9) Then
      If (Not (AMTF(9) Or AMTF(10))) Then
         IsGoodAMTFlag = False
         Exit Function
      End If
   End If
   If aSearchFlag(10) Then
      If (Not AMTF(10)) Then
         IsGoodAMTFlag = False
         Exit Function
      End If
   End If
   IsGoodAMTFlag = True
End If
End Function

' Unused Function (March 2003)
'''Private Function CntGoodAMTs() As Long
''''returns number of OKFlagged AMTs; search flag needs to be set
''''before use of this function; this was used to test behavior
''''of Flags
'''Dim i As Long
'''Dim Cnt As Long
'''Cnt = 0
'''For i = 1 To UBound(AMTData)
'''    If IsGoodAMTFlag(AMTData(i).Flag) Then Cnt = Cnt + 1
'''Next i
'''CntGoodAMTs = Cnt
'''End Function

Public Sub InitAMTStat()
    'redimensions and initialize statistic arrays
    Dim i As Long
    If AMTCnt > 0 Then
       ReDim AMTHits(1 To AMTCnt)
       ReDim AMTMWErr(1 To AMTCnt)
       ReDim AMTNETErr(1 To AMTCnt)
       ReDim AMTNETMin(1 To AMTCnt)
       ReDim AMTNETMax(1 To AMTCnt)
       'only last 2 arrays need special initialization
       For i = 1 To AMTCnt
           AMTNETMin(i) = glHugeOverExp
           AMTNETMax(i) = -1
       Next i
    End If
End Sub

Public Sub DestroyAMTStat()
'free some memory; it might come handy
Erase AMTHits
Erase AMTMWErr
Erase AMTNETErr
Erase AMTNETMin
Erase AMTNETMax
End Sub

' Unused Function (March 2003)
'''Public Function GetLockerAMTInd(ByVal ID As String) As Long
''''string should contain only one Locker mark so we take first
''''returns index in AMT array of AMT before Locker Mark
''''if any or -1 if none or on any error
'''
'''Dim LMMarkPos As Integer
'''Dim AMTIDs() As String
'''Dim AMTsCnt As Long
'''On Error GoTo err_GetLockerAMTInd
'''LMMarkPos = InStr(1, ID, glMASS_LOCKER_MARK)
'''If LMMarkPos > 0 Then
'''   'retrieve all AMT reference in front of Locker mark
'''   AMTsCnt = GetAMTRefFromString1(Left$(ID, LMMarkPos - 1), AMTIDs())
'''   If AMTsCnt > 0 Then
'''      GetLockerAMTInd = GetAMTRefInd(AMTIDs(AMTsCnt))
'''      Exit Function
'''   End If
'''End If
'''
'''err_GetLockerAMTInd:
'''GetLockerAMTInd = -1
'''End Function

' Unused Function (April 2008)
''Public Function GetAMTRefInd(ByVal AMTRef As Long) As Long
'''returns index in AMT array of AMTRef; -1 if not found or error
''Dim i As Long
''On Error GoTo err_GetAMTRefInd
''For i = 1 To AMTCnt
''    If AMTRef = AMTData(i).ID Then
''       GetAMTRefInd = i
''       Exit Function
''    End If
''Next i
''err_GetAMTRefInd:
''GetAMTRefInd = -1
''End Function

' Unused Function (March 2003)
'''Public Function GetBestAMTMatchInd(ByVal s As String) As Long
''''---------------------------------------------------------------
''''if s contains more than one AMT reference function returns
''''negative Index of reference with smallest MW error(first if
''''no error reference is found
''''if s contains only one AMT function returns its index
''''function returns 0 if no AMT reference is found or on any error
''''---------------------------------------------------------------
'''Dim AMTsCnt As Long
'''Dim AMTs() As String
'''Dim sMWErr As String
'''Dim MWErr As Double
'''Dim MinMWErr As Double
'''Dim MinMWErrInd As Long
'''Dim i As Long
'''On Error GoTo err_GetBestAMTMatchInd
'''
''''pick all ATM reference (together with errors)
'''AMTsCnt = GetAMTRefFromString2(s, AMTs())
'''If AMTsCnt > 1 Then
'''   MinMWErr = glHugeOverExp
'''   MinMWErrInd = 0
'''   For i = 1 To AMTsCnt
'''       sMWErr = GetMWErrFromString(AMTs(i))
'''       If IsNumeric(sMWErr) Then
'''          MWErr = CDbl(sMWErr)
'''       Else
'''          MWErr = glHugeOverExp
'''       End If
'''       If MWErr < MinMWErr Then
'''          MinMWErr = MWErr
'''          MinMWErrInd = i
'''       End If
'''   Next i
'''   If MinMWErrInd > 0 Then  'negative index so that we know there were multiple hits
'''      GetBestAMTMatchInd = -GetAMTRefInd(GetIDFromString(AMTs(MinMWErrInd), AMTMark, AMTIDEnd))
'''   Else
'''      GetBestAMTMatchInd = -GetAMTRefInd(GetIDFromString(AMTs(1), AMTMark, AMTIDEnd))
'''   End If
'''ElseIf AMTsCnt = 1 Then
'''   GetBestAMTMatchInd = GetAMTRefInd(GetIDFromString(AMTs(1), AMTMark, AMTIDEnd))
'''Else
'''   GetBestAMTMatchInd = 0
'''End If
'''Exit Function
'''
'''err_GetBestAMTMatchInd:
'''GetBestAMTMatchInd = 0
'''End Function


Public Function SearchAMTWithTag1(ByVal Ind As Long, _
                                 ByVal sExpr As String) As Long
'searches AMT database for MWs from GelData(ind) based
'on values in global variable samtDef arguments.
'To optimize for speed arrays loaded from the AMT table
'are actually searched rather than database recordsets
'SearchFlag determines which AMTs are included in search
'lm:03/31/2001;nt; no idea why this UMC_LO condition was present
Dim MinFN As Long
Dim MaxFN As Long
Dim ScanWidth As Long
Dim AMTRef As String
Dim IsoF As Integer     'Isotopic MW field; just shortcut
Dim i As Long, j As Long

If Not GelData(Ind).CustomNETsDefined Then
    If Not InitExprEvaluator(sExpr) Then
       SearchAMTWithTag1 = -2
       Exit Function
    End If
End If

With GelData(Ind)
   HitsCount = 0
   Set mwutSearch = New MWUtil
   If Not FillMWSearchObject(mwutSearch) Then GoTo err_SearchAMTWithTag1
   IsoF = samtDef.MWField
   GetScanRange Ind, MinFN, MaxFN, ScanWidth
   If ScanWidth <= 0 And samtDef.NETTol >= 0 Then GoTo err_SearchAMTWithTag1  'can not do it
   SetAMTSearchFlags samtDef.SearchFlag, aSearchFlag()
   Select Case samtDef.SearchScope
   Case glScope.glSc_All                 'search all data
     If samtDef.SkipReferenced Then
        If .CSLines > 0 Then
           For i = 1 To .CSLines
             If Not IsAMTReferenced(.CSData(i).MTID) Then
                For j = 1 To samtDef.MaxMassTags
                  If samtDef.NETTol >= 0 Then
                     Select Case samtDef.NETorRT
                     Case glAMT_NET
                       AMTRef = GetAMTReferenceMWNET(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                     Case glAMT_RT_or_PNET
                       AMTRef = GetAMTReferenceMWRT(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                     End Select
                  Else
                     AMTRef = GetAMTReferenceMW(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                  End If
                  InsertBefore .CSData(i).MTID, AMTRef
                Next j
             End If
           Next i
        End If
        If .IsoLines > 0 Then
          For i = 1 To .IsoLines
            If Not IsAMTReferenced(.IsoData(i).MTID) Then
              For j = 1 To samtDef.MaxMassTags
                 If samtDef.NETTol >= 0 Then
                    Select Case samtDef.NETorRT
                    Case glAMT_NET
                      AMTRef = GetAMTReferenceMWNET(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                    Case glAMT_RT_or_PNET
                      AMTRef = GetAMTReferenceMWRT(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                    End Select
                 Else
                    AMTRef = GetAMTReferenceMW(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                 End If
                 InsertBefore .IsoData(i).MTID, AMTRef
               Next j
             End If
          Next i
        End If
     Else
        If .CSLines > 0 Then
          For i = 1 To .CSLines
            For j = 1 To samtDef.MaxMassTags
              If samtDef.NETTol >= 0 Then
                 Select Case samtDef.NETorRT
                 Case glAMT_NET
                   AMTRef = GetAMTReferenceMWNET(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                 Case glAMT_RT_or_PNET
                   AMTRef = GetAMTReferenceMWRT(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                 End Select
              Else
                 AMTRef = GetAMTReferenceMW(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
              End If
              InsertBefore .CSData(i).MTID, AMTRef
            Next j
          Next i
        End If
        If .IsoLines > 0 Then
          For i = 1 To .IsoLines
            For j = 1 To samtDef.MaxMassTags
              If samtDef.NETTol >= 0 Then
                 Select Case samtDef.NETorRT
                 Case glAMT_NET
                   AMTRef = GetAMTReferenceMWNET(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                 Case glAMT_RT_or_PNET
                   AMTRef = GetAMTReferenceMWRT(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                 End Select
              Else
                 AMTRef = GetAMTReferenceMW(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
              End If
              InsertBefore .IsoData(i).MTID, AMTRef
            Next j
          Next i
        End If
     End If
   Case glScope.glSc_Current             'search current view data
     If samtDef.SkipReferenced Then
        If .CSLines > 0 Then
           For i = 1 To .CSLines
              If GelDraw(Ind).CSID(i) > 0 And GelDraw(Ind).CSR(i) > 0 Then
                 If Not IsAMTReferenced(.CSData(i).MTID) Then
                    For j = 1 To samtDef.MaxMassTags
                      If samtDef.NETTol >= 0 Then
                        Select Case samtDef.NETorRT
                        Case glAMT_NET
                           AMTRef = GetAMTReferenceMWNET(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                        Case glAMT_RT_or_PNET
                           AMTRef = GetAMTReferenceMWRT(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                        End Select
                      Else
                        AMTRef = GetAMTReferenceMW(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                      End If
                      InsertBefore .CSData(i).MTID, AMTRef
                    Next j
                  End If
              End If
           Next i
        End If
        If .IsoLines > 0 Then
           For i = 1 To .IsoLines
              If GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0 Then
                 If Not IsAMTReferenced(.IsoData(i).MTID) Then
                    For j = 1 To samtDef.MaxMassTags
                      If samtDef.NETTol >= 0 Then
                         Select Case samtDef.NETorRT
                         Case glAMT_NET
                           AMTRef = GetAMTReferenceMWNET(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                         Case glAMT_RT_or_PNET
                           AMTRef = GetAMTReferenceMWRT(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                         End Select
                      Else
                         AMTRef = GetAMTReferenceMW(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                      End If
                      InsertBefore .IsoData(i).MTID, AMTRef
                    Next j
                 End If
              End If
           Next i
        End If
     Else
        If .CSLines > 0 Then
           For i = 1 To .CSLines
              If GelDraw(Ind).CSID(i) > 0 And GelDraw(Ind).CSR(i) > 0 Then
                 For j = 1 To samtDef.MaxMassTags
                   If samtDef.NETTol >= 0 Then
                      Select Case samtDef.NETorRT
                      Case glAMT_NET
                        AMTRef = GetAMTReferenceMWNET(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                      Case glAMT_RT_or_PNET
                        AMTRef = GetAMTReferenceMWRT(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                      End Select
                   Else
                      AMTRef = GetAMTReferenceMW(.CSData(i).AverageMW - (j * samtDef.MassTag), NET_RT(Ind, .CSData(i).ScanNumber, MinFN, MaxFN), j)
                   End If
                   InsertBefore .CSData(i).MTID, AMTRef
                 Next j
              End If
           Next i
        End If
        If .IsoLines > 0 Then
           For i = 1 To .IsoLines
              If GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0 Then
                 For j = 1 To samtDef.MaxMassTags
                   If samtDef.NETTol >= 0 Then
                      Select Case samtDef.NETorRT
                      Case glAMT_NET
                        AMTRef = GetAMTReferenceMWNET(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                      Case glAMT_RT_or_PNET
                        AMTRef = GetAMTReferenceMWRT(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                      End Select
                   Else
                      AMTRef = GetAMTReferenceMW(GetIsoMass(.IsoData(i), IsoF) - (j * samtDef.MassTag), NET_RT(Ind, .IsoData(i).ScanNumber, MinFN, MaxFN), j)
                   End If
                   InsertBefore .IsoData(i).MTID, AMTRef
                 Next j
              End If
           Next i
        End If
      End If
   End Select
   SearchAMTWithTag1 = HitsCount
   
exit_SearchAMTWithTag1:
End With
Set mwutSearch = Nothing
Exit Function

err_SearchAMTWithTag1:
SearchAMTWithTag1 = -1
GoTo exit_SearchAMTWithTag1
End Function

' Unused Function (March 2003)
'''Public Function IsLCKReferenced(s) As Boolean
''''-----------------------------------------------
''''returns True if string s contains LCK reference
''''-----------------------------------------------
'''If IsNull(s) Then
'''   IsLCKReferenced = False
'''Else
'''   If InStr(1, s, LCK_MARK) > 0 Then
'''      IsLCKReferenced = True
'''   Else
'''      IsLCKReferenced = False
'''   End If
'''End If
'''End Function
'''
'''Public Function IsMTGReferenced(s) As Boolean
''''-----------------------------------------------
''''returns True if string s contains MTG reference
''''-----------------------------------------------
'''If IsNull(s) Then
'''   IsMTGReferenced = False
'''Else
'''   If InStr(1, s, MTG_MARK) > 0 Then
'''      IsMTGReferenced = True
'''   Else
'''      IsMTGReferenced = False
'''   End If
'''End If
'''End Function


Public Function CheckMassTags() As String
'--------------------------------------------------
'returns string with some important data parameters
'--------------------------------------------------
Dim sTmp As String
Dim i As Long

Dim IDNulls As Long
Dim IDOKs As Long

Dim MWNulls As Long
Dim MWOKs As Long
Dim MWOthers As Long
Dim MWMin As Double, MWMax As Double

Dim ETAllNulls As Long
Dim ETAllOKs As Long
Dim ETAllOthers As Long
Dim ETAllMin As Double, ETAllMax As Double

Dim ET1stNulls As Long
Dim ET1stOKs As Long
Dim ET1stOthers As Long
Dim ET1stMin As Double, ET1stMax As Double

If AMTCnt <= 0 Then
   CheckMassTags = "No MT tags loaded."
   Exit Function
End If

MWMin = glHugeOverExp
MWMax = -glHugeOverExp
ETAllMin = glHugeOverExp
ETAllMax = -glHugeOverExp
ET1stMin = glHugeOverExp
ET1stMax = -glHugeOverExp

For i = 1 To AMTCnt
    
    If IsNull(AMTData(i).ID) Then        'check ID strings
       IDNulls = IDNulls + 1
    Else
       IDOKs = IDOKs + 1
    End If
    
    
    If IsNull(AMTData(i).MW) Then
       MWNulls = MWNulls + 1
    Else
       If IsNumeric(AMTData(i).MW) Then
          If AMTData(i).MW > 0 Then
             MWOKs = MWOKs + 1
             If AMTData(i).MW > MWMax Then MWMax = AMTData(i).MW
             If AMTData(i).MW < MWMin Then MWMin = AMTData(i).MW
          Else
             MWOthers = MWOthers + 1
          End If
       Else
            MWOthers = MWOthers + 1
       End If
    End If
    
    
    If IsNull(AMTData(i).NET) Then
       ETAllNulls = ETAllNulls + 1
    Else
       If IsNumeric(AMTData(i).NET) Then
          If AMTData(i).NET > 0 Then                 'should also check for <=1
             ETAllOKs = ETAllOKs + 1
             If AMTData(i).NET > ETAllMax Then ETAllMax = AMTData(i).NET
             If AMTData(i).NET < ETAllMin Then ETAllMin = AMTData(i).NET
          Else
             ETAllOthers = ETAllOthers + 1
          End If
       Else
            ETAllOthers = ETAllOthers + 1
       End If
    End If
    
    If IsNull(AMTData(i).PNET) Then
       ET1stNulls = ET1stNulls + 1
    Else
       If IsNumeric(AMTData(i).PNET) Then
          If AMTData(i).PNET > 0 Then
             ET1stOKs = ET1stOKs + 1
             If AMTData(i).PNET > ET1stMax Then ET1stMax = AMTData(i).PNET
             If AMTData(i).PNET < ET1stMin Then ET1stMin = AMTData(i).PNET
          Else
             ET1stOthers = ET1stOthers + 1
          End If
       Else
            ET1stOthers = ET1stOthers + 1
       End If
    End If
    
    
Next i

If Len(CurrMTDatabase) > 0 Then         'MT tag database
   sTmp = "Database type: MTS" & vbCrLf
   sTmp = sTmp & "Connection string: " & CurrMTDatabase & vbCrLf
   sTmp = sTmp & "MTSubset: " & CurrMTFilteringOptions.MTSubsetID & vbCrLf
   sTmp = sTmp & "Inclusion list: " & CurrMTFilteringOptions.MTIncList & vbCrLf
Else                                    'must be legacy database
   sTmp = "Database type: Legacy Access DB" & vbCrLf
   If Len(CurrLegacyMTDatabase) > 0 Then
      sTmp = sTmp & "Path: " & CurrLegacyMTDatabase & vbCrLf
   Else
      sTmp = sTmp & "Path: " & glbPreferencesExpanded.LegacyAMTDBPath & vbCrLf
   End If
End If

sTmp = sTmp & vbCrLf & "MT tags: " & Format$(AMTCnt, "###,###,##0") & vbCrLf
sTmp = sTmp & "OK MT tag ID: " & Format$(IDOKs, "###,###,##0") & vbCrLf
sTmp = sTmp & "Null MT tag ID: " & Format$(IDNulls, "###,###,##0") & vbCrLf

sTmp = sTmp & "OK Masses: " & Format$(MWOKs, "###,###,##0") & vbCrLf
sTmp = sTmp & "Null Masses: " & Format$(MWNulls, "###,###,##0") & vbCrLf
sTmp = sTmp & "Bad Masses: " & Format$(MWOthers, "###,###,##0") & vbCrLf
If MWOKs > 0 Then
    sTmp = sTmp & "Mass Range: " & Format$(MWMin, "0.000000") & " - " & Format$(MWMax, "0.000000") & vbCrLf
End If

sTmp = sTmp & vbCrLf & "OK ET (NET/all results): " & Format$(ETAllOKs, "###,###,##0") & vbCrLf
sTmp = sTmp & "Null ET (NET/all results): " & Format$(ETAllNulls, "###,###,##0") & vbCrLf
sTmp = sTmp & "Bad ET (NET/all results): " & Format$(ETAllOthers, "###,###,##0") & vbCrLf
If ETAllOKs > 0 Then
    sTmp = sTmp & "ET Range: " & Format$(ETAllMin, "0.000000") & " - " & Format$(ETAllMax, "0.000000") & vbCrLf
End If

sTmp = sTmp & vbCrLf & "OK ET (RT/best results): " & Format$(ET1stOKs, "###,###,##0") & vbCrLf
sTmp = sTmp & "Null ET (RT/best results): " & Format$(ET1stNulls, "###,###,##0") & vbCrLf
sTmp = sTmp & "Bad ET (RT/best results): " & Format$(ET1stOthers, "###,###,##0") & vbCrLf
If ET1stOKs > 0 Then
    sTmp = sTmp & "ET Range: " & Format$(ET1stMin, "0.000000") & " - " & Format$(ET1stMax, "0.000000") & vbCrLf
End If

If AMTScoreStatsCnt > 0 Then
    sTmp = sTmp & vbCrLf
    sTmp = sTmp & "MT Stats: " & Format$(AMTScoreStatsCnt, "###,###,##0")
End If

CheckMassTags = sTmp
End Function

'------------------------------------------------------------
'assumption for next two functions is that MT tag molecular
'masses are loaded in mwutSearch object in ascending order
'------------------------------------------------------------
Public Function GetMTHits1(ByVal MW As Double, _
                           ByVal MWTol As Double, _
                           ByVal ET As Double, _
                           ByVal ETTol As Double, _
                           ByRef HitsInd() As Long) As Long
'------------------------------------------------------------
'returns number of MT tags(from currently loaded set that
'satisfy MW/ET condition); indices of hits are returned in
'HitsInd array; -1 is returned on any error
'if ETTol < 0 then ET criteria is not used (only MW is used)
'NOTE: ET for this function means NET (AMTData array)
'------------------------------------------------------------
Dim Ind1 As Long
Dim Ind2 As Long
Dim i As Long
Dim TmpCnt As Long
On Error GoTo err_GetMTHits1:

If mwutSearch.FindIndexRange(MW, MWTol, Ind1, Ind2) Then
   ReDim HitsCnt(100)       'should be plenty
   If ETTol >= 0 Then       'use ET tolerance
      For i = Ind1 To Ind2
        If ((Abs(ET - AMTData(i).NET) <= ETTol)) Then
           TmpCnt = TmpCnt + 1
           HitsInd(TmpCnt - 1) = i
        End If
      Next i
   Else                     'ET does not matter
      For i = Ind1 To Ind2
        TmpCnt = TmpCnt + 1
        HitsInd(TmpCnt - 1) = i
      Next i
   End If
End If
exit_GetMTHits1:
If TmpCnt > 0 Then
   ReDim Preserve HitsInd(TmpCnt - 1)
Else
   Erase HitsInd
End If
GetMTHits1 = TmpCnt
Exit Function

err_GetMTHits1:
Select Case Err.Number
Case 9                              'need more room for hits
    If TmpCnt > 1000 Then           'too many hits
       Resume exit_GetMTHits1
    Else
       ReDim Preserve HitsInd(TmpCnt + 100)
       Resume
    End If
Case Else
    Debug.Assert False
    TmpCnt = -1
    Resume exit_GetMTHits1
End Select
End Function


Public Function GetMTHits2(ByVal MW As Double, _
                           ByVal MWTol As Double, _
                           ByVal ET As Double, _
                           ByVal ETTol As Double, _
                           ByRef HitsInd() As Long) As Long
'------------------------------------------------------------
'returns number of MT tags(from currently loaded set that
'satisfy MW/ET condition); indices of hits are returned in
'HitsInd array; -1 is returned on any error
'if ETTol < 0 then ET criteria is not used (only MW is used)
'NOTE: ET for this function means RT (AMTData().PNET array)
'------------------------------------------------------------
Dim Ind1 As Long
Dim Ind2 As Long
Dim i As Long
Dim TmpCnt As Long
On Error GoTo err_GetMTHits2:

If mwutSearch.FindIndexRange(MW, MWTol, Ind1, Ind2) Then
   ReDim HitsCnt(100)       'should be plenty
   If ETTol >= 0 Then       'use ET tolerance
      For i = Ind1 To Ind2
        If ((Abs(ET - AMTData(i).PNET) <= ETTol)) Then
           TmpCnt = TmpCnt + 1
           HitsInd(TmpCnt - 1) = i
        End If
      Next i
   Else                     'ET does not matter
      For i = Ind1 To Ind2
        TmpCnt = TmpCnt + 1
        HitsInd(TmpCnt - 1) = i
      Next i
   End If
End If
exit_GetMTHits2:
If TmpCnt > 0 Then
   ReDim Preserve HitsInd(TmpCnt - 1)
Else
   Erase HitsInd
End If
GetMTHits2 = TmpCnt
Exit Function

err_GetMTHits2:
Select Case Err.Number
Case 9                              'need more room for hits
    If TmpCnt > 1000 Then           'too many hits
       Resume exit_GetMTHits2
    Else
       ReDim Preserve HitsInd(TmpCnt + 100)
       Resume
    End If
Case Else
    TmpCnt = -1
    Resume exit_GetMTHits2
End Select
End Function

Public Function CreateNewMTSearchObject(Optional ByVal blnUseN15AMTMasses As Boolean = False) As Boolean
'-------------------------------------------------
'create new object for fast search of MT tags
'If blnUseN15AMTMasses = True, then the search object is populated with the N15 forms of the MT tags
'Currently, only frmSearchForNETAdjustmentUMC uses this feature
'-------------------------------------------------

Dim dblN15AMTMasses() As Double
Dim lngIndex As Long

On Error GoTo err_CreateNewMTSearchObject
Set mwutSearch = New MWUtil

If blnUseN15AMTMasses Then
    If AMTCnt > 0 Then
        ReDim dblN15AMTMasses(1 To AMTCnt)
        For lngIndex = 1 To AMTCnt
            dblN15AMTMasses(lngIndex) = AMTData(lngIndex).MW + glN14N15_DELTA * AMTData(lngIndex).CNT_N
        Next lngIndex
    Else
        ReDim dblN15AMTMasses(1)
    End If
    
    mSearchObjectHasN15Masses = True
    CreateNewMTSearchObject = mwutSearch.Fill(dblN15AMTMasses())
Else
    mSearchObjectHasN15Masses = False
    CreateNewMTSearchObject = FillMWSearchObject(mwutSearch)
End If

Exit Function

err_CreateNewMTSearchObject:
LogErrors Err.Number, "CreateNewMTSEarchObject"
Set mwutSearch = Nothing
End Function

Public Function DestroyMTSearchObject() As Boolean
'-------------------------------------------------
'destroy object for fast search of MT tags
'-------------------------------------------------
On Error Resume Next
Set mwutSearch = Nothing
mSearchObjectHasN15Masses = False
End Function

Public Function AMTSearchObjectHasN15Masses() As Boolean
    ' Returns the value of mSearchObjectHasN15Masses
    AMTSearchObjectHasN15Masses = mSearchObjectHasN15Masses
End Function

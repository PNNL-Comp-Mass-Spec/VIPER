Attribute VB_Name = "Module20"
'last modified: 08/05/2002 nt
'------------------------------------------------------------
Option Explicit

'secondary delimiter (used only when passing delimited line
'of arguments to stored procedures)
Public Const DELI = "/"

'constants synchronized with MT database table T_FPR_Type_Name (field FT_ID and FT_Name)
'NOTE: Ideally, this should be loaded from the database and not hard-coded
Public Const FPR_Type_NA As Integer = -1
Public Const FPR_Type_Standard As Integer = 0
Public Const FPR_Type_N14_N15_L As Integer = 1
Public Const FPR_Type_N14_N15_H As Integer = 2
Public Const FPR_Type_ICAT_L As Integer = 3
Public Const FPR_Type_ICAT_H As Integer = 4
Public Const FPR_Type_C12_C13_L As Integer = 5
Public Const FPR_Type_C12_C13_H As Integer = 6
Public Const FPR_Type_PEO_L As Integer = 7
Public Const FPR_Type_PEO_H As Integer = 8
Public Const FPR_Type_PhIAT_L As Integer = 9
Public Const FPR_Type_PhIAT_H As Integer = 10
Public Const FPR_Type_PEO_N14_N15_L As Integer = 11
Public Const FPR_Type_PEO_N14_N15_H As Integer = 12
Public Const FPR_Type_O16_O18_L As Integer = 13
Public Const FPR_Type_O16_O18_H As Integer = 14
Public Const FPR_Type_MSMS As Integer = 10000

Public Const NAME_SUBSET As String = "MTSubset ID"
Public Const NAME_INC_LIST As String = "Search Inclusion List"
Public Const NAME_CONFIRMED_ONLY As String = "Confirmed Only"
Public Const NAME_ACCURATE_ONLY As String = "Accurate Only"
Public Const NAME_LOCKERS_ONLY As String = "Lockers Only"
Public Const NAME_LIMIT_TO_PMTS_FROM_DATASET As String = "Limit to PMTs from Dataset"

Public Const NAME_MINIMUM_HIGH_NORMALIZED_SCORE As String = "Minimum High Normalized Score"
Public Const NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE As String = "Minimum High Discriminant Score"
Public Const NAME_MINIMUM_PEPTIDE_PROPHET_PROBABILITY As String = "Minimum Peptide Prophet Probability"
Public Const NAME_MINIMUM_PMT_QUALITY_SCORE As String = "MinimumPMTQualityScore"

Public Const NAME_NET_VALUE_TYPE As String = "NETValueType"
Public Const NAME_EXPERIMENT_INCLUSION_FILTER As String = "Experiment Inclusion Filter"
Public Const NAME_EXPERIMENT_EXCLUSION_FILTER As String = "Experiment Exclusion Filter"
Public Const NAME_INTERNAL_STANDARD_EXPLICIT As String = "Internal Standard Explicit"

Public Const NAME_GET_DB_SCHEMA_VERSION As String = "sp_GetDBSchemaVersion"

Public Const NAME_LOCKERS_TYPE As String = "Locker Type ID"
Public Const NAME_LOCKERS_MIN_SCORE As String = "Locker Min Score"
Public Const NAME_LOCKERS_CALLER_ID As String = "Locker Caller ID"

Public Type udtMTFilteringOptionsType
    CurrentJob As Long
    
    MTSubsetID As Long                  ' Only used in Schema Version 1; ID of current MT Subset("-1" if none)
    MTIncList As String                 ' -1 for all in Schema Version 1; Blank for all in Schema Version 2
    ConfirmedOnly As Boolean
    AccurateOnly As Boolean             ' Only used in Schema Version 1
    LockersOnly As Boolean              ' Only used in Schema Version 1
    LimitToPMTsFromDataset As Boolean           ' Only used in Schema Version 2
    
    MinimumHighNormalizedScore As Single
    MinimumHighDiscriminantScore As Single      ' Only used in Schema Version 2
    MinimumPeptideProphetProbability As Single  ' Only used in Schema Version 2
    MinimumPMTQualityScore As Single
    
    ExperimentInclusionFilter As String         ' Only used in Schema Version 2
    ExperimentExclusionFilter As String         ' Only used in Schema Version 2
    InternalStandardExplicit As String          ' Only used in Schema Version 2
    NETValueType As Integer                     ' Actually type nvtNetValueTypeConstants
    
    LoadConformers As Boolean
End Type

Private Const MT_FIELD_INDEX_MAX As Integer = 20
Private Enum mtfMTFieldConstants
    Mass_Tag_ID = 0
    Peptide = 1             ' Peptide sequence
    Monoisotopic_Mass = 2
    NET_Value_to_Use = 3
    NET_Obs_Count = 4
    PNET = 5
    High_Normalized_Score = 6
    StD_GANET = 7
    High_Discriminant_Score = 8
    Peptide_Obs_Count_Passing_Filter = 9
    Mod_Count = 10
    Mod_Description = 11
    High_Peptide_Prophet_Probability = 12
    Min_MSGF_SpecProb = 13
    Cleavage_State = 14
    Conformer_ID = 15
    Conformer_Charge = 16
    Conformer = 17
    Drift_Time_Avg = 18
    Drift_Time_StDev = 19
    Conformer_Obs_Count = 20
End Enum

'for legacy database CurrMTDatabase is zero-length string
'for MT tag database CurrLegacyMTDatabase is zero length string
Public CurrMTDatabase As String     'connection string of current organism MT tag database
Public CurrLegacyMTDatabase As String

' For the following, 0 means legacy, 1 means MTS Schema 1 (defunct as of ~2006), and 2 means MTS Schema 2
Public CurrMTSchemaVersion As Single

Public CurrMTFilteringOptions As udtMTFilteringOptionsType

' Unused in June 2011
'' DB from which MT Stats are currently loaded
'Public CurrMTStatsDatabase As String
'Public CurrMTStatsFilteringOptions As udtMTFilteringOptionsType
        
'modification list as obtained from MTMain database
'''Public ModCnt As Long
'''Public ModID() As Long
'''Public ModSymbol() As String
'''Public ModDescription() As String
'''Public ModSDFlag() As String
'''Public ModMassCorrection() As Double


' MonroeMod: Added frmCallingForm, along with several statements referring to frmCallingForm (see below)
'            In addition, now using SP GetMassTagsGANETParam or GetMassTagsPlusConformers, which take individual parameters rather than
'            a delimited list of parameters
Public Function LoadMassTags(ByVal lngGelIndex As Long, _
                             frmCallingForm As VB.Form, _
                             Optional intDBConnectionTimeOutSeconds As Integer = 300, _
                             Optional ByRef blnDBConnectionError As Boolean = False, _
                             Optional blnLoadtheoreticalMTFromGelORFMT As Boolean = False) As Boolean
                             
    '------------------------------------------------------------
    ' Executes command that retrieves MT tags from Organism MT tag database
    ' Returns True if at least one MT tag loaded.
    ' Additionally, sets blnDBConnectionError to True if an error
    '  occurs when connecting to the database, or when running the SP
    ' This way, even if LoadMassTags returns false, if blnDBConnectionError = True
    '  then we'll know we don't have a database connection problem; instead
    '  there are simply no MT tags, or no MT tags with NET values
    '------------------------------------------------------------
    Dim cnNew As New ADODB.Connection
    Dim sCommand As String
    Dim rsMassTags As New ADODB.Recordset
    
    Dim cmdGetMassTags As New ADODB.Command
    
    ' Stored procedure parameters
    Dim prmMTsubsetID As ADODB.Parameter
    Dim prmMTInclusionList As ADODB.Parameter
    Dim prmAMTsOnly As ADODB.Parameter
    Dim prmConfirmedOnly As ADODB.Parameter
    Dim prmLockersOnly As ADODB.Parameter
    Dim prmMinimumPMTQualityScore As ADODB.Parameter
    Dim prmMinimumHighNormalizedScore As ADODB.Parameter
    Dim prmNETValueType As ADODB.Parameter
    
    Dim prmMinimumHighDiscriminantScore As ADODB.Parameter
    Dim prmExperimentInclusionFilter As ADODB.Parameter
    Dim prmExperimentExclusionFilter As ADODB.Parameter
    Dim prmJobToFilterOnByDataset As ADODB.Parameter
    
    Dim prmMinimumPeptideProphetProbability As ADODB.Parameter
    
    ' MonroeMod
    Dim strProgressDots As String
    Dim blnSkipMassTag As Boolean
    Dim lngMassTagsParseCount As Long, lngMassTagCountWithNullValues As Long
    ''Dim blnUseTheoreticalNETs As Boolean
    
    Dim intIndex As Integer
    Dim lngErrorCode As Long
    Dim strMessage As String
    
    Const MASS_VALUE_IF_NULL As Double = 0
    Const NET_VALUE_IF_NULL As Single = -100000
    Const MEMORY_RESERVE_CHUNK_SIZE As Long = 50000
    
    Dim udtFilteringOptions As udtMTFilteringOptionsType
    
    Dim sngDBSchemaVersion As Single
    
    Dim ErrCnt As Long                              'list only first 10 errors
    
    On Error GoTo err_LoadMassTags
    
    If GelAnalysis(lngGelIndex) Is Nothing Then
        blnDBConnectionError = True
        LoadMassTags = False
        Exit Function
    End If
    
    'reserve space for 75000 MT tags; increase in chunks of 50000 after that
    ReDim AMTData(1 To 75000)
    
    ' Lookup the current MT tags filter options
    LookupMTFilteringOptions lngGelIndex, udtFilteringOptions
    
    ' Initialize the field mapping
    ReDim intFieldMapping(MT_FIELD_INDEX_MAX) As Integer
    For intIndex = 0 To MT_FIELD_INDEX_MAX
        intFieldMapping(intIndex) = -1
    Next intIndex
    
    If (GelData(lngGelIndex).DataStatusBits And GEL_DATA_STATUS_BIT_IMS_DATA) = GEL_DATA_STATUS_BIT_IMS_DATA Then
        udtFilteringOptions.LoadConformers = True
    Else
        udtFilteringOptions.LoadConformers = False
    End If
    
    If udtFilteringOptions.LoadConformers Then
        sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetMassTagsPlusConformers
        If Len(sCommand) <= 0 Then
            sCommand = "GetMassTagsPlusConformers"
            glbPreferencesExpanded.MTSConnectionInfo.spGetMassTags = sCommand
        End If
        
        ' Define the field mapping for obtaining AMTs plus conformer information
        intFieldMapping(mtfMTFieldConstants.Mass_Tag_ID) = 0
        intFieldMapping(mtfMTFieldConstants.Peptide) = 1
        intFieldMapping(mtfMTFieldConstants.Monoisotopic_Mass) = 2
        intFieldMapping(mtfMTFieldConstants.NET_Value_to_Use) = 3
        intFieldMapping(mtfMTFieldConstants.NET_Obs_Count) = 4
        intFieldMapping(mtfMTFieldConstants.PNET) = 5
        intFieldMapping(mtfMTFieldConstants.High_Normalized_Score) = 6
        intFieldMapping(mtfMTFieldConstants.StD_GANET) = 7
        intFieldMapping(mtfMTFieldConstants.High_Discriminant_Score) = 8
        intFieldMapping(mtfMTFieldConstants.Peptide_Obs_Count_Passing_Filter) = 9
        intFieldMapping(mtfMTFieldConstants.Mod_Count) = 10
        intFieldMapping(mtfMTFieldConstants.Mod_Description) = 11
        intFieldMapping(mtfMTFieldConstants.High_Peptide_Prophet_Probability) = 12
        intFieldMapping(mtfMTFieldConstants.Min_MSGF_SpecProb) = 13
        intFieldMapping(mtfMTFieldConstants.Cleavage_State) = 14
        intFieldMapping(mtfMTFieldConstants.Conformer_ID) = 15
        intFieldMapping(mtfMTFieldConstants.Conformer_Charge) = 16
        intFieldMapping(mtfMTFieldConstants.Conformer) = 17
        intFieldMapping(mtfMTFieldConstants.Drift_Time_Avg) = 18
        intFieldMapping(mtfMTFieldConstants.Drift_Time_StDev) = 19
        intFieldMapping(mtfMTFieldConstants.Conformer_Obs_Count) = 20

    Else
        sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetMassTags
        If Len(sCommand) <= 0 Or LCase(sCommand) = "getmasstagsganet" Then
            ' Note that GetMassTagsGANET was replaced by "GetMassTagsGANETParam" in 2004
            sCommand = "GetMassTagsGANETParam"
            glbPreferencesExpanded.MTSConnectionInfo.spGetMassTags = sCommand
        End If
        
        ' Define the field mapping for obtaining AMTs
        intFieldMapping(mtfMTFieldConstants.Mass_Tag_ID) = 0
        intFieldMapping(mtfMTFieldConstants.Peptide) = 1
        intFieldMapping(mtfMTFieldConstants.Monoisotopic_Mass) = 2
        intFieldMapping(mtfMTFieldConstants.NET_Value_to_Use) = 3
        intFieldMapping(mtfMTFieldConstants.NET_Obs_Count) = 12         ' Field "Cnt_GANET"
        intFieldMapping(mtfMTFieldConstants.PNET) = 4
        intFieldMapping(mtfMTFieldConstants.High_Normalized_Score) = 5
        intFieldMapping(mtfMTFieldConstants.StD_GANET) = 6
        intFieldMapping(mtfMTFieldConstants.High_Discriminant_Score) = 7
        intFieldMapping(mtfMTFieldConstants.Peptide_Obs_Count_Passing_Filter) = 8
        intFieldMapping(mtfMTFieldConstants.Mod_Count) = 9
        intFieldMapping(mtfMTFieldConstants.Mod_Description) = 10
        intFieldMapping(mtfMTFieldConstants.High_Peptide_Prophet_Probability) = 11

    End If
   
    AMTGeneration = dbgMTSOnline
    Screen.MousePointer = vbHourglass
    AMTCnt = 0
    lngMassTagsParseCount = 0
    lngMassTagCountWithNullValues = 0
    
    ' Clear MTtoORFMapCount
    MTtoORFMapCount = 0
    
    ''If udtFilteringOptions.NETValueType = nvtTheoreticalNET Then
    ''    blnUseTheoreticalNETs = True
    ''    InitializeGANET False
    ''Else
    ''    blnUseTheoreticalNETs = False
    ''End If
    
    TraceLog 5, "LoadMassTags", "EstablishConnection"
    TraceLog 5, "LoadMassTags", "Connection String = " & GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString
    
    If Not EstablishConnection(cnNew, GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, False) Then
        TraceLog 5, "LoadMassTags", "EstablishConnection returned false"
        Debug.Assert False
        
        If InStr(LCase(GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString), "pogo") > 0 Then
            GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString = Replace(GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, "pogo", "albert", 1, 1, vbTextCompare)
        End If
            
        blnDBConnectionError = True
        LoadMassTags = False
        Exit Function
    End If
    
    ' MonroeMod
    Dim strCaptionSaved As String
    strCaptionSaved = frmCallingForm.Caption
    frmCallingForm.Caption = "Initializing DB connection"
    
    TraceLog 3, "LoadMassTags", "LookupDBSchemaVersion"
    
    sngDBSchemaVersion = LookupDBSchemaVersion(cnNew)
    
    ' Now that we know the DB Schema version, check udtFilteringOptions.MTIncList
    If Len(udtFilteringOptions.MTIncList) = 0 And sngDBSchemaVersion < 2 Then
        ' This will be true if performing an auto analysis and the user had ModificationList=
        '  in the .Ini file
        ' udtFilteringOptions.MTIncList cannot be blank if Schema Version 1; set to -1
        Debug.Print ".DBStuff(NAME_INC_LIST).value was blank; assuming value is -1"
        udtFilteringOptions.MTIncList = "-1"
    End If
    
    TraceLog 3, "LoadMassTags", "Initialize LoadMassTags SPCommand"
    
    'create and tune command object to retrieve MT tags
    ' Initialize the SP
    InitializeSPCommand cmdGetMassTags, cnNew, sCommand
    
    TraceLog 3, "LoadMassTags", "Append parameters to cmdGetMassTags"
    
    If sngDBSchemaVersion < 2 Then
        Set prmMTsubsetID = cmdGetMassTags.CreateParameter("MTSubSetID", adInteger, adParamInput, , udtFilteringOptions.MTSubsetID)
        cmdGetMassTags.Parameters.Append prmMTsubsetID
        
        Set prmMTInclusionList = cmdGetMassTags.CreateParameter("GlobModsIncList", adVarChar, adParamInput, 255, udtFilteringOptions.MTIncList)
        cmdGetMassTags.Parameters.Append prmMTInclusionList
    
        Set prmAMTsOnly = cmdGetMassTags.CreateParameter("AmtsOnly", adTinyInt, adParamInput, , BoolToTinyInt(udtFilteringOptions.AccurateOnly))
        cmdGetMassTags.Parameters.Append prmAMTsOnly
    Else
        Set prmMTInclusionList = cmdGetMassTags.CreateParameter("MassCorrectionIDFilterList", adVarChar, adParamInput, 255, udtFilteringOptions.MTIncList)
        cmdGetMassTags.Parameters.Append prmMTInclusionList
    End If
    
    Set prmConfirmedOnly = cmdGetMassTags.CreateParameter("ConfirmedOnly", adTinyInt, adParamInput, , BoolToTinyInt(udtFilteringOptions.ConfirmedOnly))
    cmdGetMassTags.Parameters.Append prmConfirmedOnly
    
    If sngDBSchemaVersion < 2 Then
        Set prmLockersOnly = cmdGetMassTags.CreateParameter("LockersOnly", adTinyInt, adParamInput, , BoolToTinyInt(udtFilteringOptions.LockersOnly))
        cmdGetMassTags.Parameters.Append prmLockersOnly
    End If
    
    Set prmMinimumHighNormalizedScore = cmdGetMassTags.CreateParameter("MinimumHighNormalizedScore", adSingle, adParamInput, , udtFilteringOptions.MinimumHighNormalizedScore)
    cmdGetMassTags.Parameters.Append prmMinimumHighNormalizedScore
    
    Set prmMinimumPMTQualityScore = cmdGetMassTags.CreateParameter("MinimumPMTQualityScore", adDecimal, adParamInput)
    With prmMinimumPMTQualityScore
        .precision = 9
        .NumericScale = 5
        .Value = ValueToSqlDecimal(udtFilteringOptions.MinimumPMTQualityScore, sdcSqlDecimal9x5)
    End With
    cmdGetMassTags.Parameters.Append prmMinimumPMTQualityScore
    
    Set prmNETValueType = cmdGetMassTags.CreateParameter("NETValueType", adTinyInt, adParamInput, , udtFilteringOptions.NETValueType)
    cmdGetMassTags.Parameters.Append prmNETValueType
    
    If sngDBSchemaVersion >= 2 Then
        Set prmMinimumHighDiscriminantScore = cmdGetMassTags.CreateParameter("MinimumHighDiscriminantScore", adSingle, adParamInput, , udtFilteringOptions.MinimumHighDiscriminantScore)
        cmdGetMassTags.Parameters.Append prmMinimumHighDiscriminantScore
        
        Set prmExperimentInclusionFilter = cmdGetMassTags.CreateParameter("ExperimentFilter", adVarChar, adParamInput, 64, udtFilteringOptions.ExperimentInclusionFilter)
        cmdGetMassTags.Parameters.Append prmExperimentInclusionFilter
        
        Set prmExperimentExclusionFilter = cmdGetMassTags.CreateParameter("ExperimentExclusionFilter", adVarChar, adParamInput, 64, udtFilteringOptions.ExperimentExclusionFilter)
        cmdGetMassTags.Parameters.Append prmExperimentExclusionFilter
    
        Set prmJobToFilterOnByDataset = cmdGetMassTags.CreateParameter("JobToFilterOnByDataset", adInteger, adParamInput, , 0)
        If udtFilteringOptions.LimitToPMTsFromDataset Then
            prmJobToFilterOnByDataset.Value = udtFilteringOptions.CurrentJob
        End If
        cmdGetMassTags.Parameters.Append prmJobToFilterOnByDataset
    
        Set prmMinimumPeptideProphetProbability = cmdGetMassTags.CreateParameter("MinimumPeptideProphetProbability", adSingle, adParamInput, , udtFilteringOptions.MinimumPeptideProphetProbability)
        cmdGetMassTags.Parameters.Append prmMinimumPeptideProphetProbability
    End If
    
    'procedure returns error number or 0 if OK
    If intDBConnectionTimeOutSeconds = 0 Then intDBConnectionTimeOutSeconds = 300
    TraceLog 3, "LoadMassTags", "cmdGetMassTags.CommandTimeout = " & intDBConnectionTimeOutSeconds
    cmdGetMassTags.CommandTimeout = intDBConnectionTimeOutSeconds
    
    TraceLog 5, "LoadMassTags", "cmdGetMassTags.Execute"
    Set rsMassTags = cmdGetMassTags.Execute(, , adAsyncExecute)
    
    Do While (cmdGetMassTags.STATE And adStateExecuting)
        Sleep 500
        strProgressDots = strProgressDots & "."
        If Len(strProgressDots) > 30 Then strProgressDots = "."
        frmCallingForm.Caption = "Waiting to load MT Tags" & strProgressDots
        DoEvents
    Loop
    
    TraceLog 5, "LoadMassTags", "Done executing cmdGetMassTags"
    frmCallingForm.Caption = "Loading MT tags: "
    DoEvents
    ' MonroeMod Finish
    
    ''' Uncomment the following to limit the NET range of the loaded AMT tags
    ''Dim blnSkipNETSOutOfRange As Boolean
    ''blnSkipNETSOutOfRange = True
    
    
    
    If rsMassTags.STATE = 0 And udtFilteringOptions.LimitToPMTsFromDataset Then
        strMessage = "'Limit to Dataset For Job' was enabled but Job " & Trim(udtFilteringOptions.CurrentJob) & " was not found in database " & ExtractDBNameFromConnectionString(GelAnalysis(lngGelIndex).MTDB.cn)
        AddToAnalysisHistory lngGelIndex, strMessage

        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "Error Loading MT tags"
        End If
    End If
    
    If rsMassTags.STATE <> 0 Then
        'load MT tag data
        TraceLog 5, "LoadMassTags", "Start loading MT tags"
        Do Until rsMassTags.EOF
           AMTCnt = AMTCnt + 1
           
           ' Initialize this entry
           InitializeAMTDataEntry AMTData(AMTCnt), NET_VALUE_IF_NULL
                                   
           AMTData(AMTCnt).ID = LoadMassTagsGetValueLng(rsMassTags, intFieldMapping, mtfMTFieldConstants.Mass_Tag_ID, -1)
           AMTData(AMTCnt).Sequence = LoadMassTagsGetValueStr(rsMassTags, intFieldMapping, mtfMTFieldConstants.Peptide, "")
           
           AMTData(AMTCnt).HighNormalizedScore = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.High_Normalized_Score, 0)
           AMTData(AMTCnt).HighDiscriminantScore = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.High_Discriminant_Score, 0)
           AMTData(AMTCnt).PeptideProphetProbability = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.High_Peptide_Prophet_Probability, 0)
           
            ' MonroeMod: Store -1 as the Mass value when the MT tag Mass value is Null
            AMTData(AMTCnt).MW = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.Monoisotopic_Mass, MASS_VALUE_IF_NULL)
            
            ' MonroeMod: Store -100000 as the NET value when the MT tag NET value is Null
            AMTData(AMTCnt).NET = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.NET_Value_to_Use, NET_VALUE_IF_NULL)
            AMTData(AMTCnt).NETStDev = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.StD_GANET, 0)
            
            AMTData(AMTCnt).NETCount = LoadMassTagsGetValueLng(rsMassTags, intFieldMapping, mtfMTFieldConstants.NET_Obs_Count, 0)
            
            AMTData(AMTCnt).MSMSObsCount = LoadMassTagsGetValueLng(rsMassTags, intFieldMapping, mtfMTFieldConstants.Peptide_Obs_Count_Passing_Filter, 1)
           
            AMTData(AMTCnt).PNET = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.PNET, 0)
           
            ' MonroeMod: the NitrogenCount() Function replaces the ELCount Function
            AMTData(AMTCnt).CNT_N = NitrogenCount(AMTData(AMTCnt).Sequence)
            
            AMTData(AMTCnt).CNT_Cys = AACount(AMTData(AMTCnt).Sequence, "C")       'look for cysteine
           
            If udtFilteringOptions.LoadConformers Then
                ' Load conformer information
                AMTData(AMTCnt).Conformer_ID = LoadMassTagsGetValueLng(rsMassTags, intFieldMapping, mtfMTFieldConstants.Conformer_ID, 0)
                AMTData(AMTCnt).ConformerCharge = LoadMassTagsGetValueLng(rsMassTags, intFieldMapping, mtfMTFieldConstants.Conformer_Charge, 0)
                AMTData(AMTCnt).ConformerNum = LoadMassTagsGetValueLng(rsMassTags, intFieldMapping, mtfMTFieldConstants.Conformer, 0)
                 
                AMTData(AMTCnt).Drift_Time_Avg = LoadMassTagsGetValueDbl(rsMassTags, intFieldMapping, mtfMTFieldConstants.Drift_Time_Avg, 0)
                AMTData(AMTCnt).Conformer_Obs_Count = LoadMassTagsGetValueLng(rsMassTags, intFieldMapping, mtfMTFieldConstants.Conformer_Obs_Count, 0)
            End If
            
            ' Advance to the next row
            rsMassTags.MoveNext
    
            ' MonroeMod: Possibly skip MT tags with null mass values or null NET values
            blnSkipMassTag = False
            If AMTData(AMTCnt).MW = MASS_VALUE_IF_NULL Then
                If Not glbPreferencesExpanded.UseMassTagsWithNullMass Then
                   blnSkipMassTag = True
                End If
                lngMassTagCountWithNullValues = lngMassTagCountWithNullValues + 1
            End If
            
            If AMTData(AMTCnt).NET = NET_VALUE_IF_NULL Then
                If Not glbPreferencesExpanded.UseMassTagsWithNullNET Then
                   blnSkipMassTag = True
                End If
                lngMassTagCountWithNullValues = lngMassTagCountWithNullValues + 1
            End If
           
           
''           If blnSkipNETSOutOfRange Then
''                If AMTData(AMTCnt).NET < 0 Or AMTData(AMTCnt).NET > 1 Then
''                    blnSkipMassTag = True
''                End If
''           End If
        
           If blnSkipMassTag Then
                AMTCnt = AMTCnt - 1
           End If
           lngMassTagsParseCount = lngMassTagsParseCount + 1
           
            ' MonroeMod
            If AMTCnt Mod 100 = 0 Then
                If AMTCnt Mod 1000 = 0 Then
                    TraceLog 3, "LoadMassTags", "Reading MT tags from DB, AMTCnt = " & LongToStringWithCommas(AMTCnt)
                Else
                    TraceLog 1, "LoadMassTags", "Reading MT tags from DB, AMTCnt = " & LongToStringWithCommas(AMTCnt)
                End If
                
                If lngMassTagsParseCount = AMTCnt Then
                    frmCallingForm.Caption = "Loading MT tags: " & LongToStringWithCommas(lngMassTagsParseCount)
                Else
                    frmCallingForm.Caption = "Loading MT tags: " & LongToStringWithCommas(AMTCnt) & " valid of " & LongToStringWithCommas(lngMassTagsParseCount) & " PMT's"
                End If
            End If
            DoEvents
        Loop
        rsMassTags.Close

    End If

    
    ' Update the AMT staleness stats
    With glbPreferencesExpanded.MassTagStalenessOptions
        .AMTLoadTime = Now()
        .AMTCountInDB = lngMassTagsParseCount
        .AMTCountWithNulls = lngMassTagCountWithNullValues
    End With
    
    '' Code that was used by the ORFViewer; No longer supported (March 2006)
    ''Dim blnCopiedValuesFromOtherGel As Boolean
    ''blnLoadtheoreticalMTFromGelORFMT = False
    ''If blnLoadtheoreticalMTFromGelORFMT Then
    ''    InitializeGANET
    ''    LoadORFsFromMTDB GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, lngGelIndex, blnCopiedValuesFromOtherGel
    ''    LoadMassTagsForORFSFromMTDB GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, lngGelIndex, True, blnCopiedValuesFromOtherGel
    ''    UpdateORFStatistics lngGelIndex
    ''    ComputeTheoreticalTrypticMassTags GelORFData(lngGelIndex), GelORFMassTags(lngGelIndex), lngGelIndex
    ''    IncludeTrypticMassTagsInAMTs lngGelIndex, frmCallingForm
    ''End If
    
    'clean things and exit
exit_LoadMassTags:
    Screen.MousePointer = vbDefault
    On Error Resume Next
    Set cmdGetMassTags.ActiveConnection = Nothing
    cnNew.Close
    
    TraceLog 5, "LoadMassTags", "Done reading MT tags, AMTCnt = " & AMTCnt
    
    If AMTCnt > 0 Then
       If AMTCnt < UBound(AMTData) Then
          ReDim Preserve AMTData(1 To AMTCnt)
       End If
    Else
       Erase AMTData
    End If
    
    
    ' MonroeMod
    ' Restore the caption on the calling form
    frmCallingForm.Caption = strCaptionSaved
    LoadMassTags = (AMTCnt > 0)
    
    'remember which database is currently loaded
    CurrMTDatabase = GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString
    CurrLegacyMTDatabase = ""
    CurrMTSchemaVersion = sngDBSchemaVersion
    
    'remember the current filtering options
    CurrMTFilteringOptions = udtFilteringOptions
    
    Exit Function

err_LoadMassTags:
    Select Case Err.Number
    Case 9                       'need more room for MT tags
        ReDim Preserve AMTData(1 To AMTCnt + MEMORY_RESERVE_CHUNK_SIZE)
        Resume
    Case 13, 94                  'Type Mismatch or Invalid Use of Null
        Resume Next              'just ignore it
    Case 3265, 3704              'two common database connection errors
        '2nd attempt will probably work so let user know they should try again
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Error loading MT tags from the database. Error could " _
                 & "have been caused by network/server issues(timeout) so you " _
                 & "might try loading again with Refresh function: " & Err.Description, vbOKOnly, glFGTU
        End If
        blnDBConnectionError = True
    Case Else
        TraceLog 3, "LoadMassTags", "Error occurred (ErrCnt=" & Trim(ErrCnt) & "): " & Err.Description
        
        ErrCnt = ErrCnt + 1
        If ErrCnt < 10 Then
           LogErrors Err.Number, "LoadMassTags", Err.Description
           Resume Next
        End If
        blnDBConnectionError = True
    End Select
    AMTCnt = -1
    GoTo exit_LoadMassTags
End Function

Private Function LoadMassTagsGetValueDbl(ByRef rsMassTags As ADODB.Recordset, ByRef intFieldMapping() As Integer, eField As mtfMTFieldConstants, ByVal dblValueIfMissing As Double) As Double
    If intFieldMapping(eField) < 0 Then
        LoadMassTagsGetValueDbl = dblValueIfMissing
    Else
        LoadMassTagsGetValueDbl = FixNullDbl(rsMassTags.Fields(intFieldMapping(eField)).Value, dblValueIfMissing)
    End If
End Function

Private Function LoadMassTagsGetValueLng(ByRef rsMassTags As ADODB.Recordset, ByRef intFieldMapping() As Integer, eField As mtfMTFieldConstants, ByVal lngValueIfMissing As Long) As Long
    If intFieldMapping(eField) < 0 Then
        LoadMassTagsGetValueLng = lngValueIfMissing
    Else
        LoadMassTagsGetValueLng = FixNullLng(rsMassTags.Fields(intFieldMapping(eField)).Value, lngValueIfMissing)
    End If
End Function
            
Private Function LoadMassTagsGetValueStr(ByRef rsMassTags As ADODB.Recordset, ByRef intFieldMapping() As Integer, eField As mtfMTFieldConstants, ByVal strValueIfMissing As String) As String
    If intFieldMapping(eField) < 0 Then
        LoadMassTagsGetValueStr = strValueIfMissing
    Else
        LoadMassTagsGetValueStr = FixNull(rsMassTags.Fields(intFieldMapping(eField)).Value)
    End If
End Function
            
' Unused Function (June 2011)
''Public Function LoadMTStats(ByVal lngGelIndex As Long, frmCallingForm As VB.Form, Optional intDBConnectionTimeOutSeconds As Integer = 300, Optional ByRef blnDBConnectionError As Boolean = False) As Boolean
''
''    '------------------------------------------------------------
''    ' Executes command that retrieves MT Stats from an AMT Tag database
''    ' Returns True if at least one MT Stat entry was loaded.
''    ' Additionally, sets blnDBConnectionError to True if an error
''    '  occurs when connecting to the database, or when running the SP
''    ' This way, even if LoadMTStats returns false, if blnDBConnectionError = True
''    '  then we'll know we don't have a database connection problem; instead
''    '  there are simply no MT Stat entries
''    '------------------------------------------------------------
''
''    Dim cnNew As New ADODB.Connection
''    Dim sCommand As String
''    Dim rsMTStats As New ADODB.Recordset
''
''    Dim cmdGetMTStats As New ADODB.Command
''
''    ' Stored procedure parameters
''    Dim prmNonFilterPassingMTsSamplingFraction As ADODB.Parameter
''    Dim prmNonFilterPassingMTsMaxCount As ADODB.Parameter
''    Dim prmMinimumHighNormalizedScore As ADODB.Parameter
''    Dim prmMinimumHighDiscriminantScore As ADODB.Parameter
''    Dim prmMinimumPeptideProphetProbability As ADODB.Parameter
''    Dim prmMinimumPMTQualityScore As ADODB.Parameter
''
''    Dim strProgressDots As String
''
''    Dim lngErrorCode As Long
''    Dim strMessage As String
''
''    Const MASS_VALUE_IF_NULL As Double = 0
''    Const NET_VALUE_IF_NULL As Single = -100000
''    Const FSCORE_VALUE_IF_NULL As Integer = -100
''    Const MEMORY_RESERVE_CHUNK_SIZE As Long = 50000
''
''    Const Default_NonFilterPassingMTsSamplingFraction As Single = 0.5
''    Const Default_NonFilterPassingMTsMaxCount As Long = 500000
''
''    Dim udtFilteringOptions As udtMTFilteringOptionsType
''
''    Dim ErrCnt As Long                              'list only first 10 errors
''
''    On Error GoTo err_LoadMTStats
''
''    If GelAnalysis(lngGelIndex) Is Nothing Then
''        blnDBConnectionError = True
''        LoadMTStats = False
''        Exit Function
''    End If
''
''    ' Reserve space for 50000 MT Stat entries
''    ' Memory reserved will be increased as needed
''    ReDim AMTScoreStats(49999)
''
''    ' Lookup the current MT tags filter options
''    LookupMTFilteringOptions lngGelIndex, udtFilteringOptions
''
''    On Error Resume Next
''
''    sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetMTStats
''    If Len(sCommand) <= 0 Then
''        blnDBConnectionError = True
''        LoadMTStats = False
''        Exit Function
''    End If
''
''    On Error GoTo err_LoadMTStats
''
''    AMTScoreStatsCnt = 0
''
''    TraceLog 5, "LoadMTStats", "EstablishConnection"
''    TraceLog 5, "LoadMTStats", "Connection String = " & GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString
''
''    If Not EstablishConnection(cnNew, GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, False) Then
''        TraceLog 5, "LoadMTStats", "EstablishConnection returned false"
''        Debug.Assert False
''
''        If InStr(LCase(GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString), "pogo") > 0 Then
''            GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString = Replace(GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, "pogo", "albert", 1, 1, vbTextCompare)
''        End If
''
''        blnDBConnectionError = True
''        LoadMTStats = False
''        Exit Function
''    End If
''
''    Dim strCaptionSaved As String
''    strCaptionSaved = frmCallingForm.Caption
''    frmCallingForm.Caption = "Initializing DB connection"
''
''    TraceLog 3, "LoadMTStats", "Initialize LoadMTStats SPCommand"
''
''    'create and tune command object to retrieve MT tags
''    ' Initialize the SP
''    InitializeSPCommand cmdGetMTStats, cnNew, sCommand
''
''    TraceLog 3, "LoadMTStats", "Append parameters to cmdGetMTStats"
''
''    Set prmNonFilterPassingMTsSamplingFraction = cmdGetMTStats.CreateParameter("NonFilterPassingMTsSamplingFraction", adSingle, adParamInput, , Default_NonFilterPassingMTsSamplingFraction)
''    cmdGetMTStats.Parameters.Append prmNonFilterPassingMTsSamplingFraction
''
''    Set prmNonFilterPassingMTsMaxCount = cmdGetMTStats.CreateParameter("NonFilterPassingMTsMaxCount", adInteger, adParamInput, , Default_NonFilterPassingMTsMaxCount)
''    cmdGetMTStats.Parameters.Append prmNonFilterPassingMTsMaxCount
''
''    Set prmMinimumHighNormalizedScore = cmdGetMTStats.CreateParameter("MinimumHighNormalizedScore", adSingle, adParamInput, , udtFilteringOptions.MinimumHighNormalizedScore)
''    cmdGetMTStats.Parameters.Append prmMinimumHighNormalizedScore
''
''    Set prmMinimumHighDiscriminantScore = cmdGetMTStats.CreateParameter("MinimumHighDiscriminantScore", adSingle, adParamInput, , udtFilteringOptions.MinimumHighDiscriminantScore)
''    cmdGetMTStats.Parameters.Append prmMinimumHighDiscriminantScore
''
''    Set prmMinimumPeptideProphetProbability = cmdGetMTStats.CreateParameter("MinimumPeptideProphetProbability", adSingle, adParamInput, , udtFilteringOptions.MinimumPeptideProphetProbability)
''    cmdGetMTStats.Parameters.Append prmMinimumPeptideProphetProbability
''
''    Set prmMinimumPMTQualityScore = cmdGetMTStats.CreateParameter("MinimumPMTQualityScore", adSingle, adParamInput, , udtFilteringOptions.MinimumPMTQualityScore)
''    cmdGetMTStats.Parameters.Append prmMinimumPMTQualityScore
''
''
''    'procedure returns error number or 0 if OK
''    If intDBConnectionTimeOutSeconds = 0 Then intDBConnectionTimeOutSeconds = 300
''    TraceLog 3, "LoadMTStats", "cmdGetMTStats.CommandTimeout = " & intDBConnectionTimeOutSeconds
''    cmdGetMTStats.CommandTimeout = intDBConnectionTimeOutSeconds
''
''    TraceLog 5, "LoadMTStats", "cmdGetMTStats.Execute"
''    Set rsMTStats = cmdGetMTStats.Execute(, , adAsyncExecute)
''
''    Do While (cmdGetMTStats.STATE And adStateExecuting)
''        Sleep 500
''        strProgressDots = strProgressDots & "."
''        If Len(strProgressDots) > 30 Then strProgressDots = "."
''        frmCallingForm.Caption = "Waiting to load MT Stats" & strProgressDots
''        DoEvents
''    Loop
''
''    TraceLog 5, "LoadMTStats", "Done executing cmdGetMTStats"
''    frmCallingForm.Caption = "Loading MT stats: "
''    DoEvents
''
''
''    If rsMTStats.STATE <> 0 Then
''        'load MT Stats
''        TraceLog 5, "LoadMTStats", "Start loading MT Stats"
''        Do Until rsMTStats.EOF
''
''            If AMTScoreStatsCnt >= UBound(AMTScoreStats) Then
''                If UBound(AMTScoreStats) < 1000000 Then
''                    ReDim Preserve AMTScoreStats((UBound(AMTScoreStats) + 1) * 2 - 1)
''                Else
''                    ReDim Preserve AMTScoreStats((UBound(AMTScoreStats) + 1) * 1.5 - 1)
''                End If
''            End If
''
''            With rsMTStats
''                AMTScoreStats(AMTScoreStatsCnt).MTID = CLng(.Fields(0).Value)
''                AMTScoreStats(AMTScoreStatsCnt).MW = FixNullDbl(.Fields(1).Value, MASS_VALUE_IF_NULL)
''                AMTScoreStats(AMTScoreStatsCnt).NET = FixNullDbl(.Fields(2).Value, NET_VALUE_IF_NULL)
''                AMTScoreStats(AMTScoreStatsCnt).NETStDev = FixNullDbl(.Fields(3).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).NETCount = FixNullLng(.Fields(4).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).MSMSObsCount = FixNullLng(.Fields(5).Value, 1)
''                AMTScoreStats(AMTScoreStatsCnt).HighNormalizedScore = FixNullDbl(.Fields(6).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).HighDiscriminantScore = FixNullDbl(.Fields(7).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).PeptideProphetProbability = FixNullDbl(.Fields(8).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).ModCount = FixNullInt(.Fields(9).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).TrypticState = FixNullInt(.Fields(10).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).PepProphetObsCountCS1 = FixNullLng(.Fields(11).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).PepProphetObsCountCS2 = FixNullLng(.Fields(12).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).PepProphetObsCountCS3 = FixNullLng(.Fields(13).Value, 0)
''                AMTScoreStats(AMTScoreStatsCnt).PepProphetFScoreCS1 = FixNullDbl(.Fields(14).Value, FSCORE_VALUE_IF_NULL)
''                AMTScoreStats(AMTScoreStatsCnt).PepProphetFScoreCS2 = FixNullDbl(.Fields(15).Value, FSCORE_VALUE_IF_NULL)
''                AMTScoreStats(AMTScoreStatsCnt).PepProphetFScoreCS3 = FixNullDbl(.Fields(16).Value, FSCORE_VALUE_IF_NULL)
''                AMTScoreStats(AMTScoreStatsCnt).PassesFilters = FixNullInt(.Fields(17).Value, 0)
''            End With
''
''            AMTScoreStatsCnt = AMTScoreStatsCnt + 1
''            rsMTStats.MoveNext
''
''            If AMTScoreStatsCnt Mod 500 = 0 Then
''                If AMTScoreStatsCnt Mod 5000 = 0 Then
''                     TraceLog 3, "LoadMTStats", "Reading MT stats from DB, AMTScoreStatsCnt = " & LongToStringWithCommas(AMTScoreStatsCnt)
''                Else
''                     TraceLog 1, "LoadMTStats", "Reading MT stats from DB, AMTScoreStatsCnt = " & LongToStringWithCommas(AMTScoreStatsCnt)
''                End If
''
''                frmCallingForm.Caption = "Loading MT stats: " & LongToStringWithCommas(AMTScoreStatsCnt)
''            End If
''
''           DoEvents
''        Loop
''        rsMTStats.Close
''
''    End If
''
''    ' Update the AMT Stats staleness stats
''    With glbPreferencesExpanded.MassTagStalenessOptions
''        .AMTStatsLoadTime = Now()
''    End With
''
''    'clean things and exit
''exit_LoadMTStats:
''    Screen.MousePointer = vbDefault
''    On Error Resume Next
''    Set cmdGetMTStats.ActiveConnection = Nothing
''    cnNew.Close
''
''    TraceLog 5, "LoadMTStats", "Done reading MT Stats, MT Stats Count = " & AMTScoreStatsCnt
''
''    If AMTScoreStatsCnt > 0 Then
''       If AMTScoreStatsCnt < UBound(AMTScoreStats) Then
''          ReDim Preserve AMTScoreStats(AMTScoreStatsCnt - 1)
''       End If
''    Else
''       ReDim AMTScoreStats(0)
''    End If
''
''
''    ' MonroeMod
''    ' Restore the caption on the calling form
''    frmCallingForm.Caption = strCaptionSaved
''    If (AMTScoreStatsCnt > 0) Then
''        LoadMTStats = True
''    Else
''        LoadMTStats = False
''    End If
''
''    'remember which database is currently loaded
''    CurrMTStatsDatabase = GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString
''
''    'remember the current filtering options
''    CurrMTStatsFilteringOptions = udtFilteringOptions
''
''    Exit Function
''
''err_LoadMTStats:
''    Select Case Err.Number
''    Case 13, 94                  'Type Mismatch or Invalid Use of Null
''        Resume Next              'just ignore it
''    Case 3265, 3704              'two common database connection errors
''        '2nd attempt will probably work so let user know they should try again
''        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''            MsgBox "Error loading MT stats from the database. Error could " _
''                 & "have been caused by network/server issues(timeout) so you " _
''                 & "might try loading again with Refresh function.", vbOKOnly, glFGTU
''        End If
''        blnDBConnectionError = True
''    Case Else
''        TraceLog 3, "LoadMTStats", "Error occurred (ErrCnt=" & Trim(ErrCnt) & "): " & Err.Description
''
''        ErrCnt = ErrCnt + 1
''        If ErrCnt < 10 Then
''           LogErrors Err.Number, "LoadMTStats"
''           Resume Next
''        End If
''        blnDBConnectionError = True
''    End Select
''    AMTScoreStatsCnt = -1
''    GoTo exit_LoadMTStats
''
''End Function

Public Function LoadMassTagToProteinMapping(frmCallingForm As VB.Form, Ind As Long, blnIncludeORFsForMassTagsNotInMemory As Boolean) As Boolean
    '---------------------------------------------------------------------------
    'executes command that retrieves list of mappings between MT tags and ORFs
    'also retrieves the ORF reference name
    'When blnIncludeORFsForMassTagsNotInMemory = True, then retrieves all MT to ORF mappings and ORF names
    'When blnIncludeORFsForMassTagsNotInMemory = False, then only records those ORF mappings that correspond to a MT tag in memory
    'The second method is generally faster, and definitely uses less memory if only a subset of all of the MT tags are in memory
    '---------------------------------------------------------------------------
    Dim cnNew As New ADODB.Connection
    Dim sCommand As String
    Dim udtFilteringOptions As udtMTFilteringOptionsType
    Dim rsMT_ORF_Map As New ADODB.Recordset
    Dim cmdGetMap As New ADODB.Command
    Dim prmConfirmedOnly As ADODB.Parameter
    Dim prmMinimumPMTQualityScore As ADODB.Parameter
    Dim prmMinimumHighNormalizedScore As ADODB.Parameter
    Dim prmMinimumHighDiscriminantScore As ADODB.Parameter
    Dim prmMinimumPeptideProphetProbability As ADODB.Parameter
    
    Dim intMTIDField As Integer, intProteinIDField As Integer
    Dim intReferenceField As Integer, intInternalRefID As Integer
    Dim intFieldIndex As Integer
    Dim lngMassTagIDToAdd As Long
    Dim blnProceed As Boolean
    Dim strCaptionSaved As String
    Dim AMTIDsSorted() As Long          ' 1-based array to stay consistent with AMTData()
    Dim EmptyArray() As Long            ' Never allocate any memory for this; simply pass to objQSLong.QSAsc
    Dim lngAMTIndex As Long
    Dim lngORFMapItemsExamined As Long
    Dim blnSuccess As Boolean
    
    Dim objQSLong As QSLong
    
    'reserve space for 75000 mappings; increase in chunks of 10000 after that
    ReDim MTIDMap(1 To 75000)
    ReDim ORFIDMap(1 To 75000)
    ReDim ORFRefNames(1 To 75000)
    
    ' Save the form's caption
    strCaptionSaved = frmCallingForm.Caption
    
    ' Lookup the current MT tags filter options
    LookupMTFilteringOptions Ind, udtFilteringOptions
    
    On Error Resume Next
    sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetMassTagToProteinNameMap
    
    If Len(sCommand) <= 0 Then
        sCommand = "GetMassTagToProteinNameMap"
    End If
    
    On Error GoTo err_LoadMassTagToProteinMapping
    
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
    
    If Len(sCommand) <= 0 Then Exit Function
    Screen.MousePointer = vbHourglass
    MTtoORFMapCount = 0
    
    If Not EstablishConnection(cnNew, GelAnalysis(Ind).MTDB.cn.ConnectionString, False) Then
        Debug.Assert False
        LoadMassTagToProteinMapping = False
        Exit Function
    End If
    
    'create and tune command object to retrieve MT tag to protein mappings
    ' Initialize the SP
    InitializeSPCommand cmdGetMap, cnNew, sCommand
    
    If blnIncludeORFsForMassTagsNotInMemory Then
        ' Update the filtering options to effectively not filter; this way all MT to protein mappings will be loaded
        With udtFilteringOptions
            .ConfirmedOnly = False
            .MinimumHighNormalizedScore = 0
            .MinimumHighDiscriminantScore = 0
            .MinimumPeptideProphetProbability = 0
            .MinimumPMTQualityScore = 0
        End With
    End If
    
    Set prmConfirmedOnly = cmdGetMap.CreateParameter("ConfirmedOnly", adTinyInt, adParamInput, , BoolToTinyInt(udtFilteringOptions.ConfirmedOnly))
    cmdGetMap.Parameters.Append prmConfirmedOnly
    
    Set prmMinimumHighNormalizedScore = cmdGetMap.CreateParameter("MinimumHighNormalizedScore", adSingle, adParamInput, , udtFilteringOptions.MinimumHighNormalizedScore)
    cmdGetMap.Parameters.Append prmMinimumHighNormalizedScore
    
    Set prmMinimumPMTQualityScore = cmdGetMap.CreateParameter("MinimumPMTQualityScore", adSingle, adParamInput, , udtFilteringOptions.MinimumPMTQualityScore)
    cmdGetMap.Parameters.Append prmMinimumPMTQualityScore
    
    Set prmMinimumHighDiscriminantScore = cmdGetMap.CreateParameter("MinimumHighDiscriminantScore", adSingle, adParamInput, , udtFilteringOptions.MinimumHighDiscriminantScore)
    cmdGetMap.Parameters.Append prmMinimumHighDiscriminantScore
    
    Set prmMinimumPeptideProphetProbability = cmdGetMap.CreateParameter("MinimumPeptideProphetProbability", adSingle, adParamInput, , udtFilteringOptions.MinimumPeptideProphetProbability)
    cmdGetMap.Parameters.Append prmMinimumPeptideProphetProbability
    
    'procedure returns error number or 0 if OK
    Set rsMT_ORF_Map = cmdGetMap.Execute
    With rsMT_ORF_Map
        ' Determine the field indices
        intMTIDField = 0
        intProteinIDField = 1
        intReferenceField = -1
        intInternalRefID = -1
        For intFieldIndex = 0 To .Fields.Count - 1
            Select Case .Fields(intFieldIndex).Name
            Case "Mass_Tag_ID", "MT_ID": intMTIDField = intFieldIndex
            Case "Protein_ID", "ORF_ID": intProteinIDField = intFieldIndex      ' Protein_ID, globally unique across DBs
            Case "Reference": intReferenceField = intFieldIndex
            Case "Ref_ID": intInternalRefID = intFieldIndex                 ' Ref_ID for this protein in this AMT database
            Case Else
                Debug.Assert False
            End Select
        Next intFieldIndex
        
        lngORFMapItemsExamined = 0
        Do Until .EOF
           lngMassTagIDToAdd = FixNullLng(.Fields(intMTIDField).Value)
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
              ORFIDMap(MTtoORFMapCount) = FixNullLng(.Fields(intProteinIDField).Value)      ' This will be Null if the ORF_ID column in T_ORF_Reference is null
              If intReferenceField >= 0 Then ORFRefNames(MTtoORFMapCount) = FixNull(.Fields(intReferenceField).Value)
           End If
           .MoveNext
           lngORFMapItemsExamined = lngORFMapItemsExamined + 1
           If lngORFMapItemsExamined Mod 100 = 0 Then frmCallingForm.Caption = "Loading Protein data: " & LongToStringWithCommas(lngORFMapItemsExamined)
        Loop
    End With
    rsMT_ORF_Map.Close
    blnSuccess = True

'clean things and exit
exit_LoadMassTagToProteinMapping:
    On Error Resume Next
    Set cmdGetMap.ActiveConnection = Nothing
    cnNew.Close
    If MTtoORFMapCount > 0 Then
       If MTtoORFMapCount < UBound(MTIDMap) Then
          ReDim Preserve MTIDMap(1 To MTtoORFMapCount)
          ReDim Preserve ORFIDMap(1 To MTtoORFMapCount)
          ReDim Preserve ORFRefNames(1 To MTtoORFMapCount)
       End If
    Else
       Erase MTIDMap
       Erase ORFIDMap
       Erase ORFRefNames
    End If
    ' Restore the caption on the calling form
    frmCallingForm.Caption = strCaptionSaved
    Screen.MousePointer = vbDefault
    LoadMassTagToProteinMapping = blnSuccess
    Exit Function

err_LoadMassTagToProteinMapping:
    Select Case Err.Number
    Case 9                       'need more room for MT tags
        Err.Clear
        ReDim Preserve MTIDMap(1 To MTtoORFMapCount + 10000)
        ReDim Preserve ORFIDMap(1 To MTtoORFMapCount + 10000)
        ReDim Preserve ORFRefNames(1 To MTtoORFMapCount + 10000)
        Resume
    Case 13, 94                  'Type Mismatch or Invalid Use of Null
        Resume Next              'just ignore it
    Case 3265, 3704              'two errors I have encountered
        '2nd attempt will probably work so let user know it should try again
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Error loading mapping between MT tags and Proteins from the database. Error could " _
                 & "have been caused by network/server issues(timeout) so you " _
                 & "might try loading again with Refresh function.", vbOKOnly, glFGTU
        End If
    Case Else
        Debug.Assert False
        LogErrors Err.Number, "LoadMassTagToProteinMapping"
    End Select
    MTtoORFMapCount = -1
    GoTo exit_LoadMassTagToProteinMapping
End Function

Public Function BoolToTinyInt(blnOption As Boolean) As Integer
    
    If blnOption Then
        BoolToTinyInt = 1
    Else
        BoolToTinyInt = 0
    End If
    
End Function

Public Sub LookupMTFilteringOptions(ByVal lngGelIndex As Long, ByRef udtFilteringOptions As udtMTFilteringOptionsType)

    On Error Resume Next
    
    'retrieve parameters that should be used with this MT tags search
    With GelAnalysis(lngGelIndex).MTDB
        udtFilteringOptions.CurrentJob = GelAnalysis(lngGelIndex).MD_Reference_Job
        
        udtFilteringOptions.MTSubsetID = CLng(.DBStuff(NAME_SUBSET).Value)   ' Long Integer
        If Err Then
           udtFilteringOptions.MTSubsetID = -1
           Err.Clear
        End If
        udtFilteringOptions.MTIncList = .DBStuff(NAME_INC_LIST).Value        ' String: DB_Schema_Version 1: e.g. -1 or Dynamic 1 and Static 1; DB_Schema_Version 2: e.g. "" or "1014" or "Not 1014"
        If Err Then
           udtFilteringOptions.MTIncList = ""
           Err.Clear
        End If
        udtFilteringOptions.ConfirmedOnly = CBool(.DBStuff(NAME_CONFIRMED_ONLY).Value)  ' Stored as String: True or False
        If Err Then
           udtFilteringOptions.ConfirmedOnly = False
           Err.Clear
        End If
        udtFilteringOptions.AccurateOnly = CBool(.DBStuff(NAME_ACCURATE_ONLY).Value)    ' Stored as String: True or False
        If Err Then
           udtFilteringOptions.AccurateOnly = False
           Err.Clear
        End If
        udtFilteringOptions.LockersOnly = CBool(.DBStuff(NAME_LOCKERS_ONLY).Value)      ' Stored as String: True or False
        If Err Then
           udtFilteringOptions.LockersOnly = False
           Err.Clear
        End If
        udtFilteringOptions.LimitToPMTsFromDataset = CBool(.DBStuff(NAME_LIMIT_TO_PMTS_FROM_DATASET).Value)      ' Stored as String: True or False
        If Err Then
           udtFilteringOptions.LimitToPMTsFromDataset = False
           Err.Clear
        End If

        udtFilteringOptions.MinimumHighNormalizedScore = .DBStuff(NAME_MINIMUM_HIGH_NORMALIZED_SCORE).Value      ' Single
        If Err Then
            udtFilteringOptions.MinimumHighNormalizedScore = 0
            Err.Clear
        End If
        udtFilteringOptions.MinimumHighDiscriminantScore = .DBStuff(NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE).Value      ' Single
        If Err Then
            udtFilteringOptions.MinimumHighDiscriminantScore = 0
            Err.Clear
        End If
        udtFilteringOptions.MinimumPeptideProphetProbability = .DBStuff(NAME_MINIMUM_PEPTIDE_PROPHET_PROBABILITY).Value      ' Single
        If Err Then
            udtFilteringOptions.MinimumPeptideProphetProbability = 0
            Err.Clear
        End If
        udtFilteringOptions.MinimumPMTQualityScore = .DBStuff(NAME_MINIMUM_PMT_QUALITY_SCORE).Value      ' Single
        If Err Then
           udtFilteringOptions.MinimumPMTQualityScore = 0
           Err.Clear
        End If
        
        udtFilteringOptions.ExperimentInclusionFilter = .DBStuff(NAME_EXPERIMENT_INCLUSION_FILTER).Value      ' String
        If Err Then
            udtFilteringOptions.ExperimentInclusionFilter = ""
            Err.Clear
        End If
        udtFilteringOptions.ExperimentExclusionFilter = .DBStuff(NAME_EXPERIMENT_EXCLUSION_FILTER).Value      ' String
        If Err Then
            udtFilteringOptions.ExperimentExclusionFilter = ""
            Err.Clear
        End If
        
        udtFilteringOptions.InternalStandardExplicit = .DBStuff(NAME_INTERNAL_STANDARD_EXPLICIT).Value      ' String
        If Err Then
            udtFilteringOptions.InternalStandardExplicit = ""
            Err.Clear
        End If
        
        udtFilteringOptions.NETValueType = .DBStuff(NAME_NET_VALUE_TYPE).Value        ' Integer
        If Err Then
            udtFilteringOptions.NETValueType = nvtNetValueTypeConstants.nvtGANET
            Err.Clear
        End If
    End With

End Sub

Public Function GetORFRecord(ByVal Ind As Long, _
                             ByVal ORFID As Long) As String
'--------------------------------------------------------------
'returns ORF record in a form of lines of text Column: Value
'--------------------------------------------------------------
Dim cnNew As New ADODB.Connection
Dim sCommand As String
Dim rsORFRow As New ADODB.Recordset
Dim cmdGetORFRow As New ADODB.Command
Dim prmORFID As ADODB.Parameter
Dim col As ADODB.Field
Dim tmp As String

On Error Resume Next
sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetORFRecord
If Len(sCommand) <= 0 Then
   GetORFRecord = "Database access key not found in initialization file."
   Exit Function
End If

On Error GoTo err_GetORFRecord
Screen.MousePointer = vbHourglass

If Not EstablishConnection(cnNew, GelAnalysis(Ind).MTDB.cn.ConnectionString, False) Then
    Debug.Assert False
    GetORFRecord = ""
    Exit Function
End If

'create and tune command object to retrieve MT tags
' Initialize the SP
InitializeSPCommand cmdGetORFRow, cnNew, sCommand

Set prmORFID = cmdGetORFRow.CreateParameter("ORFID", adInteger, adParamInput, , ORFID)
cmdGetORFRow.Parameters.Append prmORFID
Set rsORFRow = cmdGetORFRow.Execute
rsORFRow.MoveFirst
For Each col In rsORFRow.Fields
    tmp = tmp & col.Name & ": " & col.Value & vbCrLf
Next
rsORFRow.Close

'clean things and exit
exit_GetORFRecord:
Set cmdGetORFRow.ActiveConnection = Nothing
cnNew.Close
Screen.MousePointer = vbDefault
GetORFRecord = tmp
Exit Function

err_GetORFRecord:
tmp = Err.Number & " - " & Err.Description
Resume exit_GetORFRecord
End Function


Public Function GetORFSequence(ByVal Ind As Long, _
                               ByVal ORFID As Long) As String
'--------------------------------------------------------------
'returns ORF sequence in a form of lines of text Column: Value
'--------------------------------------------------------------
Dim cnNew As New ADODB.Connection
Dim sCommand As String
Dim rsORFSeq As New ADODB.Recordset
Dim cmdGetORFSeq As New ADODB.Command
Dim prmORFID As ADODB.Parameter

On Error Resume Next
sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetORFSeq
If Len(sCommand) <= 0 Then
   GetORFSequence = "Database access key not found in initialization file."
   Exit Function
End If

On Error GoTo err_GetORFSequence
Screen.MousePointer = vbHourglass

If Not EstablishConnection(cnNew, GelAnalysis(Ind).MTDB.cn.ConnectionString, False) Then
    Debug.Assert False
    GetORFSequence = ""
    Exit Function
End If

'create and tune command object to retrieve MT tags
' Initialize the SP
InitializeSPCommand cmdGetORFSeq, cnNew, sCommand

Set prmORFID = cmdGetORFSeq.CreateParameter("ORFID", adInteger, adParamInput, , ORFID)
cmdGetORFSeq.Parameters.Append prmORFID
Set rsORFSeq = cmdGetORFSeq.Execute
GetORFSequence = rsORFSeq.Fields(0).Value
rsORFSeq.Close

'clean things and exit
exit_GetORFSequence:
Set cmdGetORFSeq.ActiveConnection = Nothing
cnNew.Close
Screen.MousePointer = vbDefault
Exit Function

err_GetORFSequence:
GetORFSequence = Err.Number & " - " & Err.Description
Resume exit_GetORFSequence
End Function

Public Sub ClearGelAnalysisObject(ByVal lngGelIndex As Long, blnSetToNothing As Boolean)
    Dim udtEmptyAnalysisInfo As udtGelAnalysisInfoType
    
    If Not GelAnalysis(lngGelIndex) Is Nothing Then
        If blnSetToNothing Then
            With GelAnalysis(lngGelIndex)
                 .DestroyParameters
                 .MTDB.DestroyDBStuff
                 .MD_Reference_Job = -1
                 Set .MTDB = Nothing
            End With
            Set GelAnalysis(lngGelIndex) = Nothing
        Else
            ' Don't actually set GelAnalysis() to nothing; just clear the contents
            ' However, keep the slope, intercept, and fit unchanged
            
            With udtEmptyAnalysisInfo
                .GANET_Slope = GelAnalysis(lngGelIndex).GANET_Slope
                .GANET_Intercept = GelAnalysis(lngGelIndex).GANET_Intercept
                .GANET_Fit = GelAnalysis(lngGelIndex).GANET_Fit
            End With
        
            Set GelAnalysis(lngGelIndex) = New FTICRAnalysis
            FillGelAnalysisObject GelAnalysis(lngGelIndex), udtEmptyAnalysisInfo
        End If
    End If
    
End Sub

Public Function CurrMTDBInfo() As String
    Dim strMessage As String

On Error Resume Next

    If AMTCnt = 0 And Len(CurrMTDatabase) = 0 Then
        strMessage = "No MT tags are loaded"
    Else
        strMessage = ""
        strMessage = strMessage & "MT tag database = "
        
        If Len(CurrMTDatabase) = 0 Then
            If Len(CurrLegacyMTDatabase) > 0 Then
                strMessage = strMessage & CurrLegacyMTDatabase
            Else
                strMessage = strMessage & "Unknown"
            End If
        Else
            strMessage = strMessage & ExtractDBNameFromConnectionString(CurrMTDatabase)
        End If
        
        strMessage = strMessage & "; MT tag Count = " & Trim(AMTCnt)
        
        With CurrMTFilteringOptions
            strMessage = strMessage & "; MT Subset ID = " & .MTSubsetID
            If .ConfirmedOnly Then strMessage = strMessage & "; Confirmed MT tags Only"
            If .AccurateOnly Then strMessage = strMessage & "; Accurate MT tags Only"
            If .LockersOnly Then strMessage = strMessage & "; Locker MT tags Only"
            If .LimitToPMTsFromDataset Then strMessage = strMessage & "; Limiting to MT tags from Dataset for job " & .CurrentJob
            
            strMessage = strMessage & "; Minimum High Normalized Score = " & Trim(.MinimumHighNormalizedScore)
            strMessage = strMessage & "; Minimum High Discriminant Score = " & Trim(.MinimumHighDiscriminantScore)
            strMessage = strMessage & "; Minimum Peptide Prophet Probability = " & Trim(.MinimumPeptideProphetProbability)
            strMessage = strMessage & "; Minimum PMT Quality Score = " & Trim(.MinimumPMTQualityScore)
            
            If Len(.ExperimentInclusionFilter) > 0 Then
                strMessage = strMessage & "; Experiment Inclusion Filter = " & .ExperimentInclusionFilter
            End If
            If Len(.ExperimentExclusionFilter) > 0 Then
                strMessage = strMessage & "; Experiment Exclusion Filter = " & .ExperimentExclusionFilter
            End If
            strMessage = strMessage & "; Net Value Type = " & LookupNETValueTypeDescription(val(.NETValueType))
            
            If .MTIncList = "-1" Or Len(.MTIncList) = 0 Then
                strMessage = strMessage & "; Inclusion List = All"
            Else
                strMessage = strMessage & "; Inclusion List = " & .MTIncList
            End If
            
            If Len(.InternalStandardExplicit) > 0 Then
                strMessage = strMessage & "; Explicit Internal Standard = " & .InternalStandardExplicit
            End If
        End With
        
        ' Unused variables (June 2011)
        ' strMessage = strMessage & "; MT Stats DB = " & CurrMTStatsDatabase
        ' strMessage = strMessage & "; MT Stats Count = " & Trim(AMTScoreStatsCnt)
    End If
    
    CurrMTDBInfo = strMessage

End Function

Public Function ExportGANETtoMTDB(CallerID As Long, dblGANETSlope As Double, dblGANETIntercept As Double, dblGANETAvgDev As Double) As String
'----------------------------------------------------
'this is simple but long procedure of exporting GANET
'parameters for the FTICR analysis to the database
'
'The mupETCalc_GANET Stored Procedure can be executed
'to assign an ET value to the matches in the T_FTICR_Peak_Results table
'----------------------------------------------------
Dim EditGANETSPName As String
'ADO objects for editing stored procedure
Dim cnNew As New ADODB.Connection
Dim cmdEditJob As New ADODB.Command
Dim prmJob As New ADODB.Parameter           'job number that will be edited
Dim prmFit As New ADODB.Parameter           'GANET fit - actually avg.deviation
Dim prmSlope As New ADODB.Parameter         'GANET slope
Dim prmIntercept As New ADODB.Parameter     'GANET intercept

Dim prmTotalScanCount As New ADODB.Parameter
Dim prmScanStart As New ADODB.Parameter
Dim prmScanEnd As New ADODB.Parameter
Dim prmDurationMinutes As New ADODB.Parameter

Dim lngScanStart As Long, lngScanEnd As Long, lngScanCount As Long

On Error GoTo ExportGANETtoMTDBErrorHandler

   EditGANETSPName = glbPreferencesExpanded.MTSConnectionInfo.spEditGANET
   
   If Len(EditGANETSPName) = 0 Then
       ' This shouldn't happen
       Debug.Assert False
       EditGANETSPName = "EditFAD_GANET"
   End If

   ' Look up the scan number statistics
   GetScanRange CallerID, lngScanStart, lngScanEnd, 0, lngScanCount
    
   If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
       Debug.Assert False
       ExportGANETtoMTDB = "Error: Unable to establish a connection to the database"
       Exit Function
   End If
    
   ' Initialize the SP
   InitializeSPCommand cmdEditJob, cnNew, EditGANETSPName
    
   Set prmJob = cmdEditJob.CreateParameter("FAD_Job", adInteger, adParamInput, , GelAnalysis(CallerID).MD_Reference_Job)
   cmdEditJob.Parameters.Append prmJob
   Set prmFit = cmdEditJob.CreateParameter("GANETFit", adDouble, adParamInput, , dblGANETAvgDev)
   cmdEditJob.Parameters.Append prmFit
   Set prmSlope = cmdEditJob.CreateParameter("GANETSlope", adDouble, adParamInput, , dblGANETSlope)
   cmdEditJob.Parameters.Append prmSlope
   Set prmIntercept = cmdEditJob.CreateParameter("GANETIntercept", adDouble, adParamInput, , dblGANETIntercept)
   cmdEditJob.Parameters.Append prmIntercept
   
   Set prmTotalScanCount = cmdEditJob.CreateParameter("TotalScanCount", adInteger, adParamInput, , lngScanCount)
   cmdEditJob.Parameters.Append prmTotalScanCount
   Set prmScanStart = cmdEditJob.CreateParameter("ScanStart", adInteger, adParamInput, , lngScanStart)
   cmdEditJob.Parameters.Append prmScanStart
   Set prmScanEnd = cmdEditJob.CreateParameter("ScanEnd", adInteger, adParamInput, , lngScanEnd)
   cmdEditJob.Parameters.Append prmScanEnd
   Set prmDurationMinutes = cmdEditJob.CreateParameter("DurationMinutes", adDouble, adParamInput)         ' Leave as the default value
   cmdEditJob.Parameters.Append prmDurationMinutes
      
   cmdEditJob.Execute
   Set cmdEditJob.ActiveConnection = Nothing

   ' MonroeMod
   AddToAnalysisHistory CallerID, "Exported NET adjustment values and scan range statistics; FAD_Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; Slope = " & DoubleToStringScientific(dblGANETSlope, 4) & "; Intercept = " & DoubleToStringScientific(dblGANETIntercept, 3) & "; Average Deviation = " & DoubleToStringScientific(dblGANETAvgDev) & "; Total Scan Count= " & Trim(lngScanCount) & "; Scan Start = " & Trim(lngScanStart) & "; Scan End = " & Trim(lngScanEnd)

ExportGANETtoMTDB = "Export of NET and scan range statistics for analysis was successful"

Exit Function

ExportGANETtoMTDBErrorHandler:
ExportGANETtoMTDB = "Error occurred during export of NET and scan range statistics for analysis: " & Err.Description

End Function

Public Function FillDisplay0ResidualCounts(ByVal Ind As Long) As Boolean
'---------------------------------------------------------------------------
'goes through identification of an display Ind and fills the most abundant
'mass field with number of an amino acids in an identification of that field
'NOTE: Unfortunately this can be done only for isotopic data; this is done
'for some strange visualization (Kostas)
'---------------------------------------------------------------------------
Dim I As Long
Dim CurrID As Long
Dim cnNew As New ADODB.Connection

If Not EstablishConnection(cnNew, GelAnalysis(Ind).MTDB.cn.ConnectionString, False) Then
    Debug.Assert False
    FillDisplay0ResidualCounts = False
    Exit Function
End If

On Error GoTo FillDisplay0ResidualCountsErrorHandler

With GelData(Ind)
    If .CSLines > 0 Then
       For I = 1 To .CSLines
           CurrID = CLng(GetTagValueFromText(CStr(.CSData(I).MTID), "MT:"))
           .CSData(I).AverageMW = Len(GetPeptSeqForID_MT(CurrID, Ind, cnNew))
       Next I
    End If
    If .IsoLines > 0 Then
       For I = 1 To .IsoLines
           CurrID = CLng(GetTagValueFromText(CStr(.IsoData(I).MTID), "MT:"))
           .IsoData(I).MostAbundantMW = Len(GetPeptSeqForID_MT(CurrID, Ind, cnNew))
       Next I
    End If
    GelData(Ind).MaxMW = 1000
    GelData(Ind).MinMW = 1
End With

Set cnNew = Nothing

Exit Function

FillDisplay0ResidualCountsErrorHandler:
    Debug.Assert False
    Set cnNew = Nothing
    
End Function

Private Function GetPeptSeqForID_MT(ByVal ID As Long, ByVal Ind As Long, ByVal cnNew As ADODB.Connection) As String
'--------------------------------------------------------------
'returns ORF sequence in a form of lines of text Column: Value
'--------------------------------------------------------------
Dim sCommand As String
Dim rsMTSeq As New ADODB.Recordset
Dim cmdGetMTSeq As New ADODB.Command
Dim prmMTID As ADODB.Parameter
On Error GoTo exit_GetMTSequence
sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetMassTagSeq
    
'create and tune command object to retrieve MT tags
' Initialize the SP
InitializeSPCommand cmdGetMTSeq, cnNew, sCommand

Set prmMTID = cmdGetMTSeq.CreateParameter("MassTagID", adInteger, adParamInput, , ID)
cmdGetMTSeq.Parameters.Append prmMTID
Set rsMTSeq = cmdGetMTSeq.Execute
GetPeptSeqForID_MT = rsMTSeq.Fields(0).Value
rsMTSeq.Close
exit_GetMTSequence:
Set cmdGetMTSeq.ActiveConnection = Nothing
End Function

Public Function LookupDBSchemaVersion(cnConnection As ADODB.Connection) As Single
    
    Dim cmdGetDBSchemaVersion As New ADODB.Command
    Dim prmDBSchemaVersion As ADODB.Parameter
    Dim sngDBSchemaVersion As Single
    
    Dim sCommand As String
    
On Error GoTo LookupDBSchemaVersionErrorHandler
    
    sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetDBSchemaVersion
    If Len(sCommand) <= 0 Then
        sCommand = "GetDBSchemaVersion"
    End If
    
    InitializeSPCommand cmdGetDBSchemaVersion, cnConnection, sCommand
    
    Set prmDBSchemaVersion = cmdGetDBSchemaVersion.CreateParameter("DBSchemaVersion", adSingle, adParamOutput, , 0)
    cmdGetDBSchemaVersion.Parameters.Append prmDBSchemaVersion
    
    cmdGetDBSchemaVersion.Execute
    
    sngDBSchemaVersion = prmDBSchemaVersion.Value
    If sngDBSchemaVersion = 0 Then
        ' Assume 2 if unknown
        sngDBSchemaVersion = 2
    End If
    
    LookupDBSchemaVersion = prmDBSchemaVersion.Value
    Exit Function

LookupDBSchemaVersionErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "LookupDBSchemaVersion"
    ' Assume 2 if unknown
    LookupDBSchemaVersion = 2
    
End Function

Public Function LookupDBSchemaVersionViaCNString(strConnectionString As String) As Single
    Static htSchemaVersionsSaved As Scripting.Dictionary
    
    Dim cnNew As ADODB.Connection
    Dim sngSchemaVersion As Single
    
    If htSchemaVersionsSaved Is Nothing Then
        Set htSchemaVersionsSaved = New Scripting.Dictionary
    End If
    
    sngSchemaVersion = 0
    If Len(strConnectionString) > 0 Then
        If htSchemaVersionsSaved.Exists(strConnectionString) Then
            sngSchemaVersion = CSng(htSchemaVersionsSaved.Item(strConnectionString))
        End If
    
        If sngSchemaVersion = 0 Then
            If EstablishConnection(cnNew, strConnectionString, False) Then
                sngSchemaVersion = LookupDBSchemaVersion(cnNew)
                htSchemaVersionsSaved.add strConnectionString, sngSchemaVersion
                cnNew.Close
            End If
        End If
    End If
    
    LookupDBSchemaVersionViaCNString = sngSchemaVersion
End Function

Public Function GetMassTagNameDisplay(MTNames() As String) As String
'-------------------------------------------------------------------
'returns first name from the list and if more than one total number
'in parentheses(empty string on any error)
'-------------------------------------------------------------------
Dim FirstNameInList As String
Dim ListCnt As Long
Dim I As Long
On Error GoTo err_GetMassTagNameDisplay
ListCnt = UBound(MTNames) + 1
If ListCnt > 0 Then
   For I = 0 To ListCnt - 1
       If Len(MTNames(I)) > 0 Then
          FirstNameInList = MTNames(I)
          Exit For
       End If
   Next I
End If
If Len(FirstNameInList) <= 0 Then FirstNameInList = IdUnknown
If ListCnt > 1 Then
   GetMassTagNameDisplay = FirstNameInList & "(" & ListCnt & ")"
Else
   GetMassTagNameDisplay = FirstNameInList
End If
Exit Function

err_GetMassTagNameDisplay:
GetMassTagNameDisplay = ""
End Function


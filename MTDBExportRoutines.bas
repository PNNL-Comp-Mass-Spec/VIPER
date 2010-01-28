Attribute VB_Name = "MTDBExportRoutines"
Option Explicit

Public Type udtPutUMCParamsListType
    MDID As ADODB.Parameter
    UMCInd As ADODB.Parameter
    MemberCount As ADODB.Parameter
    MemberCountUsedForAbu As ADODB.Parameter
    UMCScore As ADODB.Parameter
    ScanFirst As ADODB.Parameter
    ScanLast As ADODB.Parameter
    ScanMaxAbundance As ADODB.Parameter
    ClassMass As ADODB.Parameter
    MonoisotopicMassMin As ADODB.Parameter
    MonoisotopicMassMax As ADODB.Parameter
    MonoisotopicMassStDev As ADODB.Parameter
    MonoisotopicMassMaxAbu As ADODB.Parameter
    ClassAbundance As ADODB.Parameter
    AbundanceMin As ADODB.Parameter
    AbundanceMax As ADODB.Parameter
    ChargeStateMin As ADODB.Parameter
    ChargeStateMax As ADODB.Parameter
    ChargeStateMaxAbu As ADODB.Parameter
    FitAverage As ADODB.Parameter
    FitMin As ADODB.Parameter
    FitMax As ADODB.Parameter
    FitStDev As ADODB.Parameter
    ElutionTime As ADODB.Parameter
    ExpressionRatio As ADODB.Parameter
    ExpressionRatioStDev As ADODB.Parameter
    ExpressionRatioChargeStateBasisCount As ADODB.Parameter
    ExpressionRatioMemberBasisCount As ADODB.Parameter
    PeakFPRType As ADODB.Parameter                  ' 0 = Standard, 1 = Pair - N14/N15 - Light, 2 = Pair - N14/N15 - Heavy, etc.
    MassTagHitCount As ADODB.Parameter
    PairUMCInd As ADODB.Parameter                   ' Index of the pair that this UMC belongs to; -1 if not a menber of a pair
    UMCResultsIDReturn As ADODB.Parameter           ' Return value of the index of the row just added
    ClassStatsChargeBasis As ADODB.Parameter        ' Charge state of the charge group used for determing Class Mass and Class Abundance when GelUMC().def.UMCClassStatsUseStatsFromMostAbuChargeState = True; Otherwise use 0
    InternalStdCount As ADODB.Parameter             ' The number of Internal Standards that this UMC matched

' Future parameters
''    LabellingEfficiencyF As ADODB.Parameter
''    LogERCorrectedForF As ADODB.Parameter           ' Base-2 log
''    LogERStandardError As ADODB.Parameter
End Type

Public Type udtPutUMCMemberParamsListType
    UMCResultsID As ADODB.Parameter
    MemberTypeID As ADODB.Parameter
    IndexInUMC As ADODB.Parameter
    ScanNumber As ADODB.Parameter
    MZ As ADODB.Parameter
    ChargeState As ADODB.Parameter
    MonoisotopicMass As ADODB.Parameter
    Abundance As ADODB.Parameter
    IsotopicFit As ADODB.Parameter
    ElutionTime As ADODB.Parameter
    IsChargeStateRep As ADODB.Parameter
End Type

Public Const PUT_UMC_MATCH_MAX_MODSTRING_LENGTH = 50            ' Maximum length of .MassTagMods
Public Type udtPutUMCMatchParamsListType
    UMCResultsID As ADODB.Parameter
    MassTagID As ADODB.Parameter
    MatchingMemberCount As ADODB.Parameter
    MatchScore As ADODB.Parameter
    MatchState As ADODB.Parameter
    SetIsConfirmedForMT As ADODB.Parameter
    MassTagMods As ADODB.Parameter
    MassTagModMass As ADODB.Parameter
    DelMatchScore As ADODB.Parameter
End Type

Public Type udtPutUMCInternalStdMatchParamsListType
    UMCResultsID As ADODB.Parameter
    SeqID As ADODB.Parameter
    MatchingMemberCount As ADODB.Parameter
    MatchScore As ADODB.Parameter
    MatchState As ADODB.Parameter
    ExpectedNET As ADODB.Parameter
    DelMatchScore As ADODB.Parameter
End Type

Public Type udtPutUMCCSStatsParamsListType
    UMCResultsID As ADODB.Parameter
    ChargeState As ADODB.Parameter
    MemberCount As ADODB.Parameter
    MonoisotopicMass As ADODB.Parameter
    Abundance As ADODB.Parameter
    ElutionTime As ADODB.Parameter
    DriftTime As ADODB.Parameter
End Type

Private Const MASS_PRECISION = 6
Private Const FIT_PRECISION = 3
Private Const NET_PRECISION = 5

Public Function AddEntryToMatchMakingDescriptionTable(ByRef cnNew As ADODB.Connection, ByRef lngMDID As Long, ByVal ExpAnalysisSPName As String, ByVal lngGelIndex As Long, ByVal lngMatchHitCount As Long, ByVal blnUsedCustomNETs As Boolean, Optional ByVal blnSetStateToOK As Boolean = True, Optional ByVal strIniFileName As String = "") As Long
    ' Returns 0 if success, the error number if an error
    
    Dim cmdPutNewMM As New ADODB.Command
    Dim prmRefJob As New ADODB.Parameter        'reference job
    Dim prmFile As New ADODB.Parameter          'file name
    Dim prmType As New ADODB.Parameter          'type of analysis
    Dim prmParameters As New ADODB.Parameter    'analysis parameters
    Dim prmPeaksCount As New ADODB.Parameter    'count of peaks to be exported
    Dim prmIDVal As New ADODB.Parameter         'ID returned from stored procedure
    Dim prmToolVersion As New ADODB.Parameter   'Viper version string
    
    Dim prmComparisonMassTagCount As New ADODB.Parameter        ' Number of MT tags loaded from database
    Dim prmUMCTolerancePPM As New ADODB.Parameter               ' Tolerance for finding LC-MS Features
    Dim prmUMCCount As New ADODB.Parameter                      ' Number of LC-MS Features (after filtering and refinement, if applicable)
    Dim prmNetAdjTolerancePPM As New ADODB.Parameter            ' NET Adjustment mass tolerance
    Dim prmNetAdjNETMin As New ADODB.Parameter                  ' NET Adjustment result: NET value of first scan
    Dim prmNetAdjNETMax As New ADODB.Parameter                  ' NET Adjustment result: NET value of last scan
    Dim prmNetAdjUMCsHitCount As New ADODB.Parameter            ' NET Adjustment hit count after final iteration
    Dim prmNetAdjTopAbuPct As New ADODB.Parameter               ' NET Adjustment Top Abu Percent value after final iteration
    Dim prmNetAdjIterationCount As New ADODB.Parameter          ' NET Adjustment Iteration Count
    Dim prmMMATolerancePPM As New ADODB.Parameter               ' DB Search mass tolerance
    Dim prmNETTolerance As New ADODB.Parameter                  ' DB Search net tolerance
    Dim prmState As New ADODB.Parameter                         ' MD_State value; 1 = New, 2 = OK, 5 = Updated (i.e. old)
    Dim prmGANETFit As New ADODB.Parameter                      ' GANET_Fit for this analysis
    Dim prmGANETSlope As New ADODB.Parameter                    ' GANET_Slope for this analysis
    Dim prmGANETIntercept As New ADODB.Parameter                ' GANET_Intercept for this analysis
    Dim prmRefineMassCalPPMShift As New ADODB.Parameter         ' Amount of shift for mass calibration
    Dim prmRefineMassCalPeakHeightCounts As New ADODB.Parameter ' Peak height of the mass error plot for mass calibration
    Dim prmRefineMassTolUsed As New ADODB.Parameter             ' 1 if mass tolerance refinement was used
    Dim prmRefineNETTolUsed As New ADODB.Parameter              ' 1 if net tolerance refinement was used
    Dim prmMinimumHighNormalizedScore As New ADODB.Parameter    ' Minimum High Normalized Score for MT tags loaded from database
    Dim prmMinimumPMTQualityScore As New ADODB.Parameter        ' Minimum PMT Quality Score for MT tags loaded from database
    Dim prmIniFileName As New ADODB.Parameter                   ' Ini File Name (if applicable); blank otherwise
    
    Dim prmMinimumHighDiscriminantScore As New ADODB.Parameter  ' Minimum High Discriminant Score for MT tags loaded from database
    Dim prmExperimentInclusionFilter As New ADODB.Parameter     ' Experiment Inclusion Filter for MT tags loaded from database
    Dim prmExperimentExclusionFilter As New ADODB.Parameter     ' Experiment Exclusion Filter for MT tags loaded from database

    Dim prmRefineMassCalPeakWidthPPM As New ADODB.Parameter     ' Peak width of the mass error plot for mass calibration
    Dim prmRefineMassCalPeakCenterPPM As New ADODB.Parameter    ' Peak center of the mass error plot for mass calibration
    
    Dim prmRefineNETTolPeakHeightCounts As New ADODB.Parameter  ' Peak height of the NET error plot for NET tolerance adjustment
    Dim prmRefineNETTolPeakWidthNET As New ADODB.Parameter      ' Peak height of the NET error plot for NET tolerance adjustment
    Dim prmRefineNETTolPeakCenterNET As New ADODB.Parameter     ' Peak height of the NET error plot for NET tolerance adjustment
    
    Dim prmLimitToPMTsFromDataset As New ADODB.Parameter        ' 1 if the MT tags were limited to only come from the dataset associated with the loaded job
    
    Dim prmMinimumPeptideProphetProbability As New ADODB.Parameter  ' Minimum Peptide Prophet Probability for MT tags loaded from database
    
    Dim strEntryInAnalysisHistory As String, lngValueFromAnalysisHistory As Long
    Dim strNetAdjUMCsWithDBHits As String
    Dim lngHistoryIndexOfMatch As Long
    
    Dim udtMassCalErrorPeakCached As udtErrorPlottingPeakCacheType
    Dim udtNETTolErrorPeakCached As udtErrorPlottingPeakCacheType

    Dim lngErrorNumber As Long
    
    Dim lngGelScanNumberMin As Long, lngGelScanNumberMax As Long
    Dim dblMassCalPPMShift As Double

On Error GoTo AddEntryToMatchMakingDescriptionTableErrorHandler

    ' First, write new analysis in T_Match_Making_Description table
    ' Initialize the SP
    InitializeSPCommand cmdPutNewMM, cnNew, ExpAnalysisSPName
        
    'this procedure takes 17 input parameters and has 1 output
    Set prmRefJob = cmdPutNewMM.CreateParameter("Reference_Job", adInteger, adParamInput, , GelAnalysis(lngGelIndex).MD_Reference_Job)
    cmdPutNewMM.Parameters.Append prmRefJob
    Set prmFile = cmdPutNewMM.CreateParameter("File", adVarChar, adParamInput, 255, GelAnalysis(lngGelIndex).MD_file)
    cmdPutNewMM.Parameters.Append prmFile
    Set prmType = cmdPutNewMM.CreateParameter("Type", adInteger, adParamInput, , GelAnalysis(lngGelIndex).MD_Type)
    cmdPutNewMM.Parameters.Append prmType
    Set prmParameters = cmdPutNewMM.CreateParameter("Parameters", adVarChar, adParamInput, 2048, GelAnalysis(lngGelIndex).MD_Parameters)
    cmdPutNewMM.Parameters.Append prmParameters
    Set prmPeaksCount = cmdPutNewMM.CreateParameter("PeaksCount", adInteger, adParamInput, , lngMatchHitCount)
    cmdPutNewMM.Parameters.Append prmPeaksCount
    Set prmIDVal = cmdPutNewMM.CreateParameter("MatchMakingID", adInteger, adParamOutput)
    cmdPutNewMM.Parameters.Append prmIDVal
    
    Set prmToolVersion = cmdPutNewMM.CreateParameter("ToolVersion", adVarChar, adParamInput, 128, GetMyNameVersion())
    cmdPutNewMM.Parameters.Append prmToolVersion
    
    Set prmComparisonMassTagCount = cmdPutNewMM.CreateParameter("ComparisonMassTagCount", adInteger, adParamInput, , AMTCnt)
    cmdPutNewMM.Parameters.Append prmComparisonMassTagCount
    
    Set prmUMCTolerancePPM = cmdPutNewMM.CreateParameter("UMCTolerancePPM", adDecimal, adParamInput)
    With prmUMCTolerancePPM
        .precision = 9
        .NumericScale = 4
        .Value = ValueToSqlDecimal(GelUMC(lngGelIndex).def.Tol, sdcSqlDecimal9x4)
    End With
    cmdPutNewMM.Parameters.Append prmUMCTolerancePPM
    
    Set prmUMCCount = cmdPutNewMM.CreateParameter("UMCCount", adInteger, adParamInput, , GelUMC(lngGelIndex).UMCCnt)
    cmdPutNewMM.Parameters.Append prmUMCCount
    
    Set prmNetAdjTolerancePPM = cmdPutNewMM.CreateParameter("NetAdjTolerancePPM", adDecimal, adParamInput)
    With prmNetAdjTolerancePPM
        .precision = 9
        .NumericScale = 4
        .Value = ValueToSqlDecimal(GelUMCNETAdjDef(lngGelIndex).MWTol, sdcSqlDecimal9x4)
    End With
    cmdPutNewMM.Parameters.Append prmNetAdjTolerancePPM
    
    ' UMC Hit count during NET adjustment
    ' This value is stored identically in the analysis history whether or not Custom NETs were used
    strNetAdjUMCsWithDBHits = FindSettingInAnalysisHistory(lngGelIndex, UMC_NET_ADJ_UMCs_WITH_DB_HITS, lngHistoryIndexOfMatch, True, "=", ";")
    If IsNumeric(strNetAdjUMCsWithDBHits) Then
        lngValueFromAnalysisHistory = CLng(strNetAdjUMCsWithDBHits)
    Else
        ' This will happen if the user did not perform NET adjustment
        ' This will be the case if they already know the NET slope and intercept, which is entirely possible if the job was previously analyzed using VIPER
        lngValueFromAnalysisHistory = 0
    End If
    Set prmNetAdjUMCsHitCount = cmdPutNewMM.CreateParameter("NetAdjUMCsHitCount", adInteger, adParamInput, , lngValueFromAnalysisHistory)
    cmdPutNewMM.Parameters.Append prmNetAdjUMCsHitCount
    
    
    ' Top Abu Pct
    If blnUsedCustomNETs Then
        ' Always record a value of 100 when using NET warping
        lngValueFromAnalysisHistory = 100
    Else
        lngValueFromAnalysisHistory = GelUMCNETAdjDef(lngGelIndex).TopAbuPct
    End If
    
    Set prmNetAdjTopAbuPct = cmdPutNewMM.CreateParameter("NetAdjTopAbuPct", adTinyInt, adParamInput, , lngValueFromAnalysisHistory)
    cmdPutNewMM.Parameters.Append prmNetAdjTopAbuPct
    
    
    ' Iteration Count  (Max value to store in DB is 255 since datatype is TinyInt)
    If blnUsedCustomNETs Then
        ' Always record an iteration count of 1 when using NET Warping
        lngValueFromAnalysisHistory = 1
    Else
        strNetAdjUMCsWithDBHits = FindSettingInAnalysisHistory(lngGelIndex, UMC_NET_ADJ_ITERATION_COUNT, lngHistoryIndexOfMatch, True, "=", ";")
        If IsNumeric(strNetAdjUMCsWithDBHits) Then
            lngValueFromAnalysisHistory = CLng(strNetAdjUMCsWithDBHits)
            If lngValueFromAnalysisHistory > 255 Then lngValueFromAnalysisHistory = 255
        Else
            ' This will happen if the user did not perform NET adjustment
            ' This will be the case if they already know the NET slope and intercept, which is entirely possible if the job was previously analyzed using VIPER
            lngValueFromAnalysisHistory = 0
        End If
    End If
    Set prmNetAdjIterationCount = cmdPutNewMM.CreateParameter("NetAdjIterationCount", adTinyInt, adParamInput, , CInt(lngValueFromAnalysisHistory))
    cmdPutNewMM.Parameters.Append prmNetAdjIterationCount
    
    
    ' MMA Tolerance
    Set prmMMATolerancePPM = cmdPutNewMM.CreateParameter("MMATolerancePPM", adDecimal, adParamInput)
    With prmMMATolerancePPM
        .precision = 9
        .NumericScale = 4
        .Value = ValueToSqlDecimal(samtDef.MWTol, sdcSqlDecimal9x4)
    End With
    cmdPutNewMM.Parameters.Append prmMMATolerancePPM
    
    Set prmNETTolerance = cmdPutNewMM.CreateParameter("NETTolerance", adDecimal, adParamInput)
    With prmNETTolerance
        .precision = 9
        .NumericScale = 5
        .Value = ValueToSqlDecimal(samtDef.NETTol, sdcSqlDecimal9x5)
    End With
    cmdPutNewMM.Parameters.Append prmNETTolerance
    
    Set prmState = cmdPutNewMM.CreateParameter("State", adTinyInt, adParamInput)
    If blnSetStateToOK Then
        prmState.Value = MMD_STATE_OK
    Else
        prmState.Value = MMD_STATE_NEW
    End If
    cmdPutNewMM.Parameters.Append prmState
    
    
    Set prmGANETFit = cmdPutNewMM.CreateParameter("GANETFit", adDouble, adParamInput)
    Set prmGANETSlope = cmdPutNewMM.CreateParameter("GANETSlope", adDouble, adParamInput)
    Set prmGANETIntercept = cmdPutNewMM.CreateParameter("GANETIntercept", adDouble, adParamInput)
        
    If Not GelAnalysis(lngGelIndex) Is Nothing Then
        With GelAnalysis(lngGelIndex)
            prmGANETFit.Value = DoubleToStringScientific(.GANET_Fit, 6)
            prmGANETSlope.Value = DoubleToStringScientific(.GANET_Slope, 6)
            prmGANETIntercept.Value = DoubleToStringScientific(.GANET_Intercept, 6)
        End With
    End If
    
    cmdPutNewMM.Parameters.Append prmGANETFit
    cmdPutNewMM.Parameters.Append prmGANETSlope
    cmdPutNewMM.Parameters.Append prmGANETIntercept
    
    
    ' Determine the scan range for the current gel
    GetScanRange lngGelIndex, lngGelScanNumberMin, lngGelScanNumberMax, 0
    
    Set prmNetAdjNETMin = cmdPutNewMM.CreateParameter("NetAdjNETMin", adDecimal, adParamInput)
    With prmNetAdjNETMin
        .precision = 9
        .NumericScale = 5
        .Value = ValueToSqlDecimal(Round(ScanToGANET(lngGelIndex, lngGelScanNumberMin), 5), sdcSqlDecimal9x5)
    End With
    cmdPutNewMM.Parameters.Append prmNetAdjNETMin
    
    Set prmNetAdjNETMax = cmdPutNewMM.CreateParameter("NetAdjNETMax", adDecimal, adParamInput)
    With prmNetAdjNETMax
        .precision = 9
        .NumericScale = 5
        .Value = ValueToSqlDecimal(Round(ScanToGANET(lngGelIndex, lngGelScanNumberMax), 5), sdcSqlDecimal9x5)
    End With
    cmdPutNewMM.Parameters.Append prmNetAdjNETMax
    
    ' Look up the Mass Calibration PPM shift (if any)
    With GelSearchDef(lngGelIndex).MassCalibrationInfo
        dblMassCalPPMShift = .OverallMassAdjustment
        If .MassUnits = gltABS Then
            ' Need to convert to ppm; we'll use 1000 as the conversion m/z
            dblMassCalPPMShift = MassToPPM(dblMassCalPPMShift, 1000)
        End If
    End With
    
    Set prmRefineMassCalPPMShift = cmdPutNewMM.CreateParameter("RefineMassCalPPMShift", adDecimal, adParamInput)
    With prmRefineMassCalPPMShift
        .precision = 9
        .NumericScale = 4
        .Value = ValueToSqlDecimal(dblMassCalPPMShift, sdcSqlDecimal9x4)
    End With
    cmdPutNewMM.Parameters.Append prmRefineMassCalPPMShift
    
    ' Lookup the Mass and NET Error peak stats, either from .AutoAnalysisCachedData, or from the analysis history
    LookupMassAndNETErrorPeakStats lngGelIndex, udtMassCalErrorPeakCached, udtNETTolErrorPeakCached
    
    Set prmRefineMassCalPeakHeightCounts = cmdPutNewMM.CreateParameter("RefineMassCalPeakHeightCounts", adInteger, adParamInput, , udtMassCalErrorPeakCached.Height)
    cmdPutNewMM.Parameters.Append prmRefineMassCalPeakHeightCounts
    
    ' Determine if DB Search Mass Tolerance refinement was used
    strEntryInAnalysisHistory = FindSettingInAnalysisHistory(lngGelIndex, SEARCH_MASS_TOL_DETERMINED, lngHistoryIndexOfMatch, True, "=", ";")
    If lngHistoryIndexOfMatch >= 0 Then
        lngValueFromAnalysisHistory = 1
    Else
        lngValueFromAnalysisHistory = 0
    End If
    
    Set prmRefineMassTolUsed = cmdPutNewMM.CreateParameter("RefineMassTolUsed", adTinyInt, adParamInput, , CInt(lngValueFromAnalysisHistory))
    cmdPutNewMM.Parameters.Append prmRefineMassTolUsed
    
    ' Determine if DB Search NET Tolerance refinement was used
    strEntryInAnalysisHistory = FindSettingInAnalysisHistory(lngGelIndex, SEARCH_NET_TOL_DETERMINED, lngHistoryIndexOfMatch, True, "=", ";")
    If lngHistoryIndexOfMatch >= 0 Then
        lngValueFromAnalysisHistory = 1
    Else
        lngValueFromAnalysisHistory = 0
    End If
    
    Set prmRefineNETTolUsed = cmdPutNewMM.CreateParameter("RefineNETTolUsed", adTinyInt, adParamInput, , CInt(lngValueFromAnalysisHistory))
    cmdPutNewMM.Parameters.Append prmRefineNETTolUsed
    
    Set prmMinimumHighNormalizedScore = cmdPutNewMM.CreateParameter("MinimumHighNormalizedScore", adDecimal, adParamInput)
    With prmMinimumHighNormalizedScore
        .precision = 9
        .NumericScale = 5
        .Value = ValueToSqlDecimal(CurrMTFilteringOptions.MinimumHighNormalizedScore, sdcSqlDecimal9x4)
    End With
    cmdPutNewMM.Parameters.Append prmMinimumHighNormalizedScore
    
    Set prmMinimumPMTQualityScore = cmdPutNewMM.CreateParameter("MinimumPMTQualityScore", adDecimal, adParamInput)
    With prmMinimumPMTQualityScore
        .precision = 9
        .NumericScale = 5
        .Value = ValueToSqlDecimal(CurrMTFilteringOptions.MinimumPMTQualityScore, sdcSqlDecimal9x4)
    End With
    cmdPutNewMM.Parameters.Append prmMinimumPMTQualityScore
        
    Set prmIniFileName = cmdPutNewMM.CreateParameter("IniFileName", adVarChar, adParamInput, 255, strIniFileName)
    cmdPutNewMM.Parameters.Append prmIniFileName
    
    Set prmMinimumHighDiscriminantScore = cmdPutNewMM.CreateParameter("MinimumHighDiscriminantScore", adSingle, adParamInput, , CurrMTFilteringOptions.MinimumHighDiscriminantScore)
    cmdPutNewMM.Parameters.Append prmMinimumHighDiscriminantScore
        
    Set prmExperimentInclusionFilter = cmdPutNewMM.CreateParameter("ExperimentFilter", adVarChar, adParamInput, 64, CurrMTFilteringOptions.ExperimentInclusionFilter)
    cmdPutNewMM.Parameters.Append prmExperimentInclusionFilter
        
    Set prmExperimentExclusionFilter = cmdPutNewMM.CreateParameter("ExperimentExclusionFilter", adVarChar, adParamInput, 64, CurrMTFilteringOptions.ExperimentExclusionFilter)
    cmdPutNewMM.Parameters.Append prmExperimentExclusionFilter
        
    Set prmRefineMassCalPeakWidthPPM = cmdPutNewMM.CreateParameter("RefineMassCalPeakWidthPPM", adSingle, adParamInput, , udtMassCalErrorPeakCached.width)
    cmdPutNewMM.Parameters.Append prmRefineMassCalPeakWidthPPM
    Set prmRefineMassCalPeakCenterPPM = cmdPutNewMM.CreateParameter("RefineMassCalPeakCenterPPM", adSingle, adParamInput, , udtMassCalErrorPeakCached.Center)
    cmdPutNewMM.Parameters.Append prmRefineMassCalPeakCenterPPM
    
    Set prmRefineNETTolPeakHeightCounts = cmdPutNewMM.CreateParameter("RefineNETTolPeakHeightCounts", adInteger, adParamInput, , udtNETTolErrorPeakCached.Height)
    cmdPutNewMM.Parameters.Append prmRefineNETTolPeakHeightCounts
    Set prmRefineNETTolPeakWidthNET = cmdPutNewMM.CreateParameter("RefineNETTolPeakWidthNET", adSingle, adParamInput, , udtNETTolErrorPeakCached.width)
    cmdPutNewMM.Parameters.Append prmRefineNETTolPeakWidthNET
    Set prmRefineNETTolPeakCenterNET = cmdPutNewMM.CreateParameter("RefineNETTolPeakCenterNET", adSingle, adParamInput, , udtNETTolErrorPeakCached.Center)
    cmdPutNewMM.Parameters.Append prmRefineNETTolPeakCenterNET
    
    Set prmLimitToPMTsFromDataset = cmdPutNewMM.CreateParameter("LimitToPMTsFromDataset", adTinyInt, adParamInput, , BoolToTinyInt(CurrMTFilteringOptions.LimitToPMTsFromDataset))
    cmdPutNewMM.Parameters.Append prmLimitToPMTsFromDataset
    
    Set prmMinimumPeptideProphetProbability = cmdPutNewMM.CreateParameter("MinimumPeptideProphetProbability", adSingle, adParamInput, , CurrMTFilteringOptions.MinimumPeptideProphetProbability)
    cmdPutNewMM.Parameters.Append prmMinimumPeptideProphetProbability
    

    ' Call the SP
    cmdPutNewMM.Execute
    lngMDID = prmIDVal.Value
    Set cmdPutNewMM.ActiveConnection = Nothing
    
    AddEntryToMatchMakingDescriptionTable = 0
    Exit Function

AddEntryToMatchMakingDescriptionTableErrorHandler:
    Debug.Assert False
    lngErrorNumber = Err.Number
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error making entry in Match Making Description (most likely the job number is not defined in T_FTICR_Analysis_Description): " & Err.Description, vbExclamation + vbOKOnly, glFGTU
    End If
    
    LogErrors Err.Number, "AddEntryToMatchMakingDescription", Err.Description, lngGelIndex
    
    If lngErrorNumber = 0 Then
        AddEntryToMatchMakingDescriptionTable = 50000
    Else
        AddEntryToMatchMakingDescriptionTable = lngErrorNumber
    End If
    
End Function

Public Sub ExportMTDBInitializePutNewUMCParams(cnNew As ADODB.Connection, cmdPutNewUMC As ADODB.Command, udtPutUMCParams As udtPutUMCParamsListType, lngMDID As Long, strStoredProcName As String)

' Initialize the SP
InitializeSPCommand cmdPutNewUMC, cnNew, strStoredProcName

With udtPutUMCParams
    
    Set .MDID = New ADODB.Parameter
    Set .UMCInd = New ADODB.Parameter
    Set .MemberCount = New ADODB.Parameter
    Set .MemberCountUsedForAbu = New ADODB.Parameter
    Set .UMCScore = New ADODB.Parameter
    Set .ScanFirst = New ADODB.Parameter
    Set .ScanLast = New ADODB.Parameter
    Set .ScanMaxAbundance = New ADODB.Parameter
    Set .ClassMass = New ADODB.Parameter
    Set .MonoisotopicMassMin = New ADODB.Parameter
    Set .MonoisotopicMassMax = New ADODB.Parameter
    Set .MonoisotopicMassStDev = New ADODB.Parameter
    Set .MonoisotopicMassMaxAbu = New ADODB.Parameter
    Set .ClassAbundance = New ADODB.Parameter
    Set .AbundanceMin = New ADODB.Parameter
    Set .AbundanceMax = New ADODB.Parameter
    Set .ChargeStateMin = New ADODB.Parameter
    Set .ChargeStateMax = New ADODB.Parameter
    Set .ChargeStateMaxAbu = New ADODB.Parameter
    Set .FitAverage = New ADODB.Parameter
    Set .FitMin = New ADODB.Parameter
    Set .FitMax = New ADODB.Parameter
    Set .FitStDev = New ADODB.Parameter
    Set .ElutionTime = New ADODB.Parameter
    Set .ExpressionRatio = New ADODB.Parameter
    Set .ExpressionRatioStDev = New ADODB.Parameter
    Set .ExpressionRatioChargeStateBasisCount = New ADODB.Parameter
    Set .ExpressionRatioMemberBasisCount = New ADODB.Parameter
    
    Set .PeakFPRType = New ADODB.Parameter
    Set .MassTagHitCount = New ADODB.Parameter
    Set .PairUMCInd = New ADODB.Parameter
    Set .UMCResultsIDReturn = New ADODB.Parameter
    Set .ClassStatsChargeBasis = New ADODB.Parameter
    Set .InternalStdCount = New ADODB.Parameter

' Future parameters
''    Set .LabellingEfficiencyF = New ADODB.Parameter
''    Set .LogERCorrectedForF = New ADODB.Parameter
''    Set .LogERStandardError = New ADODB.Parameter

    Set .MDID = cmdPutNewUMC.CreateParameter("MDID", adInteger, adParamInput, , lngMDID)
    cmdPutNewUMC.Parameters.Append .MDID
    Set .UMCInd = cmdPutNewUMC.CreateParameter("UMCInd", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .UMCInd
    Set .MemberCount = cmdPutNewUMC.CreateParameter("MemberCount", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .MemberCount
    Set .UMCScore = cmdPutNewUMC.CreateParameter("UMCScore", adDouble, adParamInput, , 0)     ' Not currently used
    cmdPutNewUMC.Parameters.Append .UMCScore
    
    Set .ScanFirst = cmdPutNewUMC.CreateParameter("ScanFirst", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ScanFirst
    Set .ScanLast = cmdPutNewUMC.CreateParameter("ScanLast", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ScanLast
    Set .ScanMaxAbundance = cmdPutNewUMC.CreateParameter("ScanMaxAbundance", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ScanMaxAbundance
    
    Set .ClassMass = cmdPutNewUMC.CreateParameter("ClassMass", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ClassMass
    Set .MonoisotopicMassMin = cmdPutNewUMC.CreateParameter("MonoisotopicMassMin", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .MonoisotopicMassMin
    Set .MonoisotopicMassMax = cmdPutNewUMC.CreateParameter("MonoisotopicMassMax", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .MonoisotopicMassMax
    Set .MonoisotopicMassStDev = cmdPutNewUMC.CreateParameter("MonoisotopicMassStDev", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .MonoisotopicMassStDev
    Set .MonoisotopicMassMaxAbu = cmdPutNewUMC.CreateParameter("MonoisotopicMassMaxAbu", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .MonoisotopicMassMaxAbu
    
    Set .ClassAbundance = cmdPutNewUMC.CreateParameter("ClassAbundance", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ClassAbundance
    Set .AbundanceMin = cmdPutNewUMC.CreateParameter("AbundanceMin", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .AbundanceMin
    Set .AbundanceMax = cmdPutNewUMC.CreateParameter("AbundanceMax", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .AbundanceMax
    
    Set .ChargeStateMin = cmdPutNewUMC.CreateParameter("ChargeStateMin", adSmallInt, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ChargeStateMin
    Set .ChargeStateMax = cmdPutNewUMC.CreateParameter("ChargeStateMax", adSmallInt, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ChargeStateMax
    Set .ChargeStateMaxAbu = cmdPutNewUMC.CreateParameter("ChargeStateMaxAbu", adSmallInt, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ChargeStateMaxAbu
    
    Set .FitAverage = cmdPutNewUMC.CreateParameter("FitAverage", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .FitAverage
    Set .FitMin = cmdPutNewUMC.CreateParameter("FitMin", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .FitMin
    Set .FitMax = cmdPutNewUMC.CreateParameter("FitMax", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .FitMax
    Set .FitStDev = cmdPutNewUMC.CreateParameter("FitStDev", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .FitStDev
    
    Set .ElutionTime = cmdPutNewUMC.CreateParameter("ElutionTime", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ElutionTime
    Set .ExpressionRatio = cmdPutNewUMC.CreateParameter("Expression_Ratio", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ExpressionRatio
    
    Set .PeakFPRType = cmdPutNewUMC.CreateParameter("PeakFPRType", adInteger, adParamInput, , FPR_Type_Standard)
    cmdPutNewUMC.Parameters.Append .PeakFPRType
    Set .MassTagHitCount = cmdPutNewUMC.CreateParameter("MassTagHitCount", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .MassTagHitCount
    Set .PairUMCInd = cmdPutNewUMC.CreateParameter("PairUMCInd", adInteger, adParamInput, , -1)
    cmdPutNewUMC.Parameters.Append .PairUMCInd
    
    Set .UMCResultsIDReturn = cmdPutNewUMC.CreateParameter("UMCResultsID", adInteger, adParamOutput)
    cmdPutNewUMC.Parameters.Append .UMCResultsIDReturn
    
    Set .ClassStatsChargeBasis = cmdPutNewUMC.CreateParameter("ClassStatsChargeBasis", adTinyInt, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ClassStatsChargeBasis
    
    ' This parameter is named GANETLockerCount for legacy reasons
    Set .InternalStdCount = cmdPutNewUMC.CreateParameter("GANETLockerCount", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .InternalStdCount

    Set .ExpressionRatioStDev = cmdPutNewUMC.CreateParameter("ExpressionRatioStDev", adDouble, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ExpressionRatioStDev
    Set .ExpressionRatioChargeStateBasisCount = cmdPutNewUMC.CreateParameter("ExpressionRatioChargeStateBasisCount", adSmallInt, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ExpressionRatioChargeStateBasisCount
    Set .ExpressionRatioMemberBasisCount = cmdPutNewUMC.CreateParameter("ExpressionRatioMemberBasisCount", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .ExpressionRatioMemberBasisCount

    Set .MemberCountUsedForAbu = cmdPutNewUMC.CreateParameter("MemberCountUsedForAbu", adInteger, adParamInput, , 0)
    cmdPutNewUMC.Parameters.Append .MemberCountUsedForAbu

' Future parameters
''    Set .LabellingEfficiencyF = cmdPutNewUMC.CreateParameter("LabellingEfficiencyF", adSingle, adParamInput, , 0)
''    cmdPutNewUMC.Parameters.Append .LabellingEfficiencyF
''    Set .LogERCorrectedForF = cmdPutNewUMC.CreateParameter("LogERCorrectedForF", adSingle, adParamInput, , 0)
''    cmdPutNewUMC.Parameters.Append .LogERCorrectedForF
''    Set .LogERStandardError = cmdPutNewUMC.CreateParameter("LogERStandardError", adSingle, adParamInput, , 0)
''    cmdPutNewUMC.Parameters.Append .LogERStandardError

End With

End Sub

Public Sub ExportMTDBInitializePutNewUMCMemberParams(cnNew As ADODB.Connection, cmdPutNewUMCMember As ADODB.Command, udtPutUMCMemberParams As udtPutUMCMemberParamsListType, strStoredProcName As String)
    
    ' Initialize the SP
    InitializeSPCommand cmdPutNewUMCMember, cnNew, strStoredProcName
        
    With udtPutUMCMemberParams
        Set .UMCResultsID = New ADODB.Parameter
        Set .MemberTypeID = New ADODB.Parameter
        Set .IndexInUMC = New ADODB.Parameter
        Set .ScanNumber = New ADODB.Parameter
        Set .MZ = New ADODB.Parameter
        Set .ChargeState = New ADODB.Parameter
        Set .MonoisotopicMass = New ADODB.Parameter
        Set .Abundance = New ADODB.Parameter
        Set .IsotopicFit = New ADODB.Parameter
        Set .ElutionTime = New ADODB.Parameter
        Set .IsChargeStateRep = New ADODB.Parameter
        
        Set .UMCResultsID = cmdPutNewUMCMember.CreateParameter("UMCResultsID", adInteger, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .UMCResultsID
        
        Set .MemberTypeID = cmdPutNewUMCMember.CreateParameter("MemberTypeID", adTinyInt, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .MemberTypeID
        
        Set .IndexInUMC = cmdPutNewUMCMember.CreateParameter("IndexInUMC", adSmallInt, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .IndexInUMC
        
        Set .ScanNumber = cmdPutNewUMCMember.CreateParameter("ScanNumber", adInteger, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .ScanNumber
        
        Set .MZ = cmdPutNewUMCMember.CreateParameter("MZ", adDouble, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .MZ
        
        Set .ChargeState = cmdPutNewUMCMember.CreateParameter("ChargeState", adSmallInt, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .ChargeState
        
        Set .MonoisotopicMass = cmdPutNewUMCMember.CreateParameter("MonoisotopicMass", adDouble, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .MonoisotopicMass
        
        Set .Abundance = cmdPutNewUMCMember.CreateParameter("Abundance", adDouble, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .Abundance
        Set .IsotopicFit = cmdPutNewUMCMember.CreateParameter("IsotopicFit", adSingle, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .IsotopicFit
        
        Set .ElutionTime = cmdPutNewUMCMember.CreateParameter("ElutionTime", adSingle, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .ElutionTime
        Set .IsChargeStateRep = cmdPutNewUMCMember.CreateParameter("IsChargeStateRep", adTinyInt, adParamInput, , 0)
        cmdPutNewUMCMember.Parameters.Append .IsChargeStateRep
        
    End With

End Sub

Public Sub ExportMTDBInitializePutUMCMatchParams(cnNew As ADODB.Connection, cmdPutNewUMCMatch As ADODB.Command, udtPutUMCMatchParams As udtPutUMCMatchParamsListType, strStoredProcName As String)
    
' Initialize the SP
InitializeSPCommand cmdPutNewUMCMatch, cnNew, strStoredProcName
    
With udtPutUMCMatchParams
    Set .UMCResultsID = New ADODB.Parameter
    Set .MassTagID = New ADODB.Parameter
    Set .MatchingMemberCount = New ADODB.Parameter
    Set .MatchScore = New ADODB.Parameter
    Set .MatchState = New ADODB.Parameter
    Set .SetIsConfirmedForMT = New ADODB.Parameter
    Set .MassTagMods = New ADODB.Parameter
    Set .MassTagModMass = New ADODB.Parameter
    Set .DelMatchScore = New ADODB.Parameter
    
    Set .UMCResultsID = cmdPutNewUMCMatch.CreateParameter("UMCResultsID", adInteger, adParamInput, , 0)
    cmdPutNewUMCMatch.Parameters.Append .UMCResultsID
    Set .MassTagID = cmdPutNewUMCMatch.CreateParameter("MassTagID", adInteger, adParamInput, , 0)
    cmdPutNewUMCMatch.Parameters.Append .MassTagID
    
    Set .MatchingMemberCount = cmdPutNewUMCMatch.CreateParameter("MatchingMemberCount", adInteger, adParamInput, , 0)
    cmdPutNewUMCMatch.Parameters.Append .MatchingMemberCount
    
    Set .MatchScore = cmdPutNewUMCMatch.CreateParameter("MatchScore", adDecimal, adParamInput)
    With .MatchScore
        .precision = 9
        .NumericScale = 5
        '' .value = -1     ' Default: -1
    End With
    cmdPutNewUMCMatch.Parameters.Append .MatchScore
    
    ' Note: For this search mode, all matches are deemed Hits
    Set .MatchState = cmdPutNewUMCMatch.CreateParameter("MatchState", adTinyInt, adParamInput, , MATCH_STATE_HIT)
    cmdPutNewUMCMatch.Parameters.Append .MatchState
    
    Set .SetIsConfirmedForMT = cmdPutNewUMCMatch.CreateParameter("SetIsConfirmedForMT", adTinyInt, adParamInput)
    If glbPreferencesExpanded.AutoAnalysisOptions.SetIsConfirmedForDBSearchMatches Then
        .SetIsConfirmedForMT.Value = 1
    Else
        .SetIsConfirmedForMT.Value = 0
    End If
    cmdPutNewUMCMatch.Parameters.Append .SetIsConfirmedForMT
    
    Set .MassTagMods = cmdPutNewUMCMatch.CreateParameter("MassTagMods", adVarChar, adParamInput, PUT_UMC_MATCH_MAX_MODSTRING_LENGTH, "")
    cmdPutNewUMCMatch.Parameters.Append .MassTagMods
    
    Set .MassTagModMass = cmdPutNewUMCMatch.CreateParameter("MassTagModMass", adDouble, adParamInput, , 0)
    cmdPutNewUMCMatch.Parameters.Append .MassTagModMass
   
    Set .DelMatchScore = cmdPutNewUMCMatch.CreateParameter("DelMatchScore", adDecimal, adParamInput)
    With .DelMatchScore
        .precision = 9
        .NumericScale = 5
        '' .value = 0     ' Default: 0
    End With
    cmdPutNewUMCMatch.Parameters.Append .DelMatchScore
   
End With

End Sub

Public Sub ExportMTDBInitializePutUMCInternalStdMatchParams(cnNew As ADODB.Connection, cmPutNewUMCInternalStdMatch As ADODB.Command, udtPutUMCInternalStdMatchParams As udtPutUMCInternalStdMatchParamsListType, strStoredProcName As String)
    
' Initialize the SP
InitializeSPCommand cmPutNewUMCInternalStdMatch, cnNew, strStoredProcName
    
With udtPutUMCInternalStdMatchParams
    Set .UMCResultsID = New ADODB.Parameter
    Set .SeqID = New ADODB.Parameter
    Set .MatchingMemberCount = New ADODB.Parameter
    Set .MatchScore = New ADODB.Parameter
    Set .MatchState = New ADODB.Parameter
    Set .ExpectedNET = New ADODB.Parameter
    Set .DelMatchScore = New ADODB.Parameter
    
    Set .UMCResultsID = cmPutNewUMCInternalStdMatch.CreateParameter("UMCResultsID", adInteger, adParamInput, , 0)
    cmPutNewUMCInternalStdMatch.Parameters.Append .UMCResultsID
    Set .SeqID = cmPutNewUMCInternalStdMatch.CreateParameter("SeqID", adInteger, adParamInput, , 0)
    cmPutNewUMCInternalStdMatch.Parameters.Append .SeqID
    
    Set .MatchingMemberCount = cmPutNewUMCInternalStdMatch.CreateParameter("MatchingMemberCount", adInteger, adParamInput, , 0)
    cmPutNewUMCInternalStdMatch.Parameters.Append .MatchingMemberCount
    
    Set .MatchScore = cmPutNewUMCInternalStdMatch.CreateParameter("MatchScore", adDecimal, adParamInput)
    With .MatchScore
        .precision = 9
        .NumericScale = 5
        '' .value = -1     ' Default: -1
    End With
    cmPutNewUMCInternalStdMatch.Parameters.Append .MatchScore
    
    ' Note: For this search mode, all matches are deemed Hits
    Set .MatchState = cmPutNewUMCInternalStdMatch.CreateParameter("MatchState", adTinyInt, adParamInput, , MATCH_STATE_HIT)
    cmPutNewUMCInternalStdMatch.Parameters.Append .MatchState
    
    Set .ExpectedNET = cmPutNewUMCInternalStdMatch.CreateParameter("ExpectedNET", adDouble, adParamInput, , 0)
    cmPutNewUMCInternalStdMatch.Parameters.Append .ExpectedNET
   
    Set .DelMatchScore = cmPutNewUMCInternalStdMatch.CreateParameter("DelMatchScore", adDecimal, adParamInput)
    With .DelMatchScore
        .precision = 9
        .NumericScale = 5
        '' .value = 0     ' Default: 0
    End With
    cmPutNewUMCInternalStdMatch.Parameters.Append .DelMatchScore
   
End With

End Sub


Public Sub ExportMTDBInitializePutUMCCSStatsParams(cnNew As ADODB.Connection, cmdPutNewUMCCSStat As ADODB.Command, udtPutUMCCSStatsParams As udtPutUMCCSStatsParamsListType, strStoredProcName As String)
  
    ' Initialize the SP
    InitializeSPCommand cmdPutNewUMCCSStat, cnNew, strStoredProcName
        
    With udtPutUMCCSStatsParams
        Set .UMCResultsID = New ADODB.Parameter
        Set .ChargeState = New ADODB.Parameter
        Set .MemberCount = New ADODB.Parameter
        Set .MonoisotopicMass = New ADODB.Parameter
        Set .Abundance = New ADODB.Parameter
        Set .ElutionTime = New ADODB.Parameter
        Set .DriftTime = New ADODB.Parameter
       
        Set .UMCResultsID = cmdPutNewUMCCSStat.CreateParameter("UMCResultsID", adInteger, adParamInput, , 0)
        cmdPutNewUMCCSStat.Parameters.Append .UMCResultsID
        
        Set .ChargeState = cmdPutNewUMCCSStat.CreateParameter("ChargeState", adSmallInt, adParamInput, , 0)
        cmdPutNewUMCCSStat.Parameters.Append .ChargeState
       
        Set .MemberCount = cmdPutNewUMCCSStat.CreateParameter("MemberCount", adSmallInt, adParamInput, , 0)
        cmdPutNewUMCCSStat.Parameters.Append .MemberCount
        
        Set .MonoisotopicMass = cmdPutNewUMCCSStat.CreateParameter("MonoisotopicMass", adDouble, adParamInput, , 0)
        cmdPutNewUMCCSStat.Parameters.Append .MonoisotopicMass
        
        Set .Abundance = cmdPutNewUMCCSStat.CreateParameter("Abundance", adDouble, adParamInput, , 0)
        cmdPutNewUMCCSStat.Parameters.Append .Abundance
        
        Set .ElutionTime = cmdPutNewUMCCSStat.CreateParameter("ElutionTime", adSingle, adParamInput, , 0)
        cmdPutNewUMCCSStat.Parameters.Append .ElutionTime

        Set .DriftTime = cmdPutNewUMCCSStat.CreateParameter("DriftTime", adSingle, adParamInput, , 0)
        cmdPutNewUMCCSStat.Parameters.Append .DriftTime
        
    End With

End Sub

Public Sub ExportMTDBAddUMCResultRow( _
            ByRef cmdPutNewUMC As ADODB.Command, _
            ByRef udtPutUMCParams As udtPutUMCParamsListType, _
            ByRef cmdPutNewUMCMember As ADODB.Command, _
            ByRef udtPutUMCMemberParams As udtPutUMCMemberParamsListType, _
            ByRef cmdPutNewUMCCSStats As ADODB.Command, _
            ByRef udtPutUMCCSStatsParams As udtPutUMCCSStatsParamsListType, _
            ByVal blnExportUMCMembers As Boolean, _
            ByVal lngGelIndex As Long, _
            ByVal lngUMCIndexOriginal As Long, _
            ByVal lngMassTagHitCount As Long, _
            ByRef ClsStat() As Double, _
            ByRef udtPairMatchStats As udtPairMatchStatsType, _
            Optional ByVal lngPeakFPRType As Long = FPR_Type_Standard, _
            Optional lngInternalStdMatchCount As Long = 0)
            
    ' Adds row to T_FTICR_UMC_Results table
    ' Also adds row to T_FTICR_UMC_CS_Stats table
    
    ' If blnExportUMCMembers, then adds rows to T_FTICR_UMC_Members table
    ' Note that DBs must have DB Schema Version >= 2 in order to save UMC members
    
    ' Note: when dblExpressionRatioOverride is 0, we use the expression ratio value
    '       returned by LookupExpressionRatioValue()
    ' When non-zero, then dblExpressionRatioOverride is used
    
    Dim lngScanNumberMin As Long, lngScanNumberMax As Long
    Dim lngMemberIndex As Long, lngDataIndex As Long
    Dim intChargeIndex As Integer
    Dim lngUMCResultsIDInDB As Long
    
On Error GoTo AddUMCErrorHandler
    
    udtPutUMCParams.UMCInd.Value = lngUMCIndexOriginal
    
    With GelUMC(lngGelIndex).UMCs(lngUMCIndexOriginal)
        udtPutUMCParams.MemberCount.Value = .ClassCount
        
        udtPutUMCParams.ScanFirst.Value = .MinScan
        udtPutUMCParams.ScanLast.Value = .MaxScan
        
        udtPutUMCParams.ClassMass.Value = Round(.ClassMW, MASS_PRECISION)
        
        udtPutUMCParams.MonoisotopicMassMin.Value = Round(.MinMW, MASS_PRECISION)
        udtPutUMCParams.MonoisotopicMassMax.Value = Round(.MaxMW, MASS_PRECISION)
        
        udtPutUMCParams.MonoisotopicMassStDev.Value = Round(.ClassMWStD, MASS_PRECISION)
        
        ' Note that the MWStDev value stored in ClsStat(, ustMassStDev) might be slightly different than .ClassMWStD
        ' Thus, the following may possibly be False
        Debug.Assert Round(udtPutUMCParams.MonoisotopicMassStDev.Value, 1) = Round(.ClassMWStD, 1)
        
        Select Case .ClassRepType
        Case glCSType
            udtPutUMCParams.ScanMaxAbundance.Value = GelData(lngGelIndex).CSData(.ClassRepInd).ScanNumber
            udtPutUMCParams.MonoisotopicMassMaxAbu.Value = Round(GelData(lngGelIndex).CSData(.ClassRepInd).AverageMW, MASS_PRECISION)
            udtPutUMCParams.ChargeStateMaxAbu.Value = GelData(lngGelIndex).CSData(.ClassRepInd).Charge
            udtPutUMCParams.AbundanceMax.Value = GelData(lngGelIndex).CSData(.ClassRepInd).Abundance
            
        Case glIsoType
            udtPutUMCParams.ScanMaxAbundance.Value = GelData(lngGelIndex).IsoData(.ClassRepInd).ScanNumber
            udtPutUMCParams.MonoisotopicMassMaxAbu.Value = Round(GelData(lngGelIndex).IsoData(.ClassRepInd).MonoisotopicMW, MASS_PRECISION)
            udtPutUMCParams.ChargeStateMaxAbu.Value = GelData(lngGelIndex).IsoData(.ClassRepInd).Charge
            udtPutUMCParams.AbundanceMax.Value = GelData(lngGelIndex).IsoData(.ClassRepInd).Abundance
        End Select
        
        udtPutUMCParams.ExpressionRatio.Value = Round(udtPairMatchStats.ExpressionRatio, 6)
        udtPutUMCParams.ExpressionRatioStDev.Value = Round(udtPairMatchStats.ExpressionRatioStDev, 6)
        udtPutUMCParams.ExpressionRatioChargeStateBasisCount.Value = udtPairMatchStats.ExpressionRatioChargeStateBasisCount
        udtPutUMCParams.ExpressionRatioMemberBasisCount.Value = udtPairMatchStats.ExpressionRatioMemberBasisCount
        
        udtPutUMCParams.ExpressionRatio.Value = Round(udtPairMatchStats.ExpressionRatio, 6)
        
        ' The following should always be true:
        Debug.Assert Round(ClsStat(lngUMCIndexOriginal, ustClassRepMW), MASS_PRECISION) = Round(udtPutUMCParams.MonoisotopicMassMaxAbu.Value, MASS_PRECISION)
    
        udtPutUMCParams.ClassAbundance.Value = .ClassAbundance
        
        udtPutUMCParams.AbundanceMin.Value = ClsStat(lngUMCIndexOriginal, ustAbundanceMin)
        
        ' The following should always be equal,
        '  unless the class stats charge basis doesn't contain the most intense data point (which does occur occasionally)
        If udtPutUMCParams.AbundanceMax.Value <> ClsStat(lngUMCIndexOriginal, ustAbundanceMax) Then
            If udtPutUMCParams.AbundanceMax.Value > 0 And ClsStat(lngUMCIndexOriginal, ustAbundanceMax) > 0 Then
                If udtPutUMCParams.AbundanceMax.Value > ClsStat(lngUMCIndexOriginal, ustAbundanceMax) Then
                    Debug.Assert udtPutUMCParams.AbundanceMax.Value / ClsStat(lngUMCIndexOriginal, ustAbundanceMax) < 2
                Else
                    'Debug.Assert ClsStat(lngUMCIndexOriginal, ustAbundanceMax) / udtPutUMCParams.AbundanceMax.Value < 2
                End If
            Else
                Debug.Assert False
            End If
        End If
        
        udtPutUMCParams.ChargeStateMin.Value = ClsStat(lngUMCIndexOriginal, ustChargeMin)
        udtPutUMCParams.ChargeStateMax.Value = ClsStat(lngUMCIndexOriginal, ustChargeMax)
        
        udtPutUMCParams.FitAverage.Value = Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), FIT_PRECISION)
        udtPutUMCParams.FitMin.Value = Round(ClsStat(lngUMCIndexOriginal, ustFitMin), FIT_PRECISION)
        udtPutUMCParams.FitMax.Value = Round(ClsStat(lngUMCIndexOriginal, ustFitMax), FIT_PRECISION)
        udtPutUMCParams.FitStDev.Value = Round(ClsStat(lngUMCIndexOriginal, ustFitStDev), FIT_PRECISION)
        
        ' Convert from scan number to NET
        udtPutUMCParams.ElutionTime.Value = Round(ScanToGANET(lngGelIndex, udtPutUMCParams.ScanMaxAbundance.Value), NET_PRECISION)
        
        udtPutUMCParams.PeakFPRType.Value = lngPeakFPRType
        udtPutUMCParams.MassTagHitCount.Value = lngMassTagHitCount
        udtPutUMCParams.PairUMCInd = udtPairMatchStats.PairIndex
        
        If GelUMC(lngGelIndex).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
            udtPutUMCParams.ClassStatsChargeBasis.Value = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge
            udtPutUMCParams.MemberCountUsedForAbu.Value = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Count
        Else
            udtPutUMCParams.ClassStatsChargeBasis.Value = 0
            udtPutUMCParams.MemberCountUsedForAbu.Value = .ClassCount
        End If

        udtPutUMCParams.InternalStdCount.Value = lngInternalStdMatchCount
    
' Future parameters
''        udtPutUMCParams.LabellingEfficiencyF = udtPairMatchStats.LabellingEfficiencyF
''        udtPutUMCParams.LogERCorrectedForF = udtPairMatchStats.LogERCorrectedForF
''        udtPutUMCParams.LogERStandardError = udtPairMatchStats.LogERStandardError
    
    End With
    
    cmdPutNewUMC.Execute
    
    
On Error GoTo AddAddnlInfoErrorHandler

    lngUMCResultsIDInDB = FixNullLng(udtPutUMCParams.UMCResultsIDReturn.Value)
    
    ExportMTDBAddUMCCSStatsRow cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, lngGelIndex, lngUMCIndexOriginal, lngUMCResultsIDInDB
    
    If blnExportUMCMembers Then
        ExportMTDBAddUMCResultMemberRow cmdPutNewUMCMember, udtPutUMCMemberParams, lngGelIndex, lngUMCIndexOriginal, lngUMCResultsIDInDB
    End If
    
    Exit Sub

AddUMCErrorHandler:
    ' Error populating or executing cmdPutNewUMC
    
    Debug.Print "Error occurred at AddUMCErrorHandler: Err Code = " & Err.Number & vbCrLf & Err.Description
    Debug.Assert False
    
    Err.Raise Err.Number
    Exit Sub

AddAddnlInfoErrorHandler:

    Debug.Print "Error occurred at AddAddnlInfoErrorHandler: Err Code = " & Err.Number & vbCrLf & Err.Description
    Debug.Assert False
    
    Err.Raise Err.Number
    Exit Sub
    
End Sub

Private Sub ExportMTDBAddUMCCSStatsRow( _
            ByRef cmdPutNewUMCCSStats As ADODB.Command, _
            ByRef udtPutUMCCSStatsParams As udtPutUMCCSStatsParamsListType, _
            ByVal lngGelIndex As Long, _
            ByVal lngUMCIndexOriginal As Long, _
            ByVal lngUMCResultsIDInDB As Long)

    ' Adds row to T_FTICR_UMC_CS_Stats table

    Dim intChargeIndex As Integer
    Dim lngClassMIndexPointer As Long
    
On Error GoTo AddUMCCSStatsErrorHandler


    ' Now add the charge stats for this UMC to T_FTICR_UMC_CS_Stats (if the table exists in the database)
    
    ' Store the UMCResultsID value
    udtPutUMCCSStatsParams.UMCResultsID = lngUMCResultsIDInDB
    
    With GelUMC(lngGelIndex).UMCs(lngUMCIndexOriginal)

        For intChargeIndex = 0 To .ChargeStateCount - 1
                
            udtPutUMCCSStatsParams.ChargeState = .ChargeStateBasedStats(intChargeIndex).Charge
            udtPutUMCCSStatsParams.MemberCount = .ChargeStateBasedStats(intChargeIndex).Count
            udtPutUMCCSStatsParams.MonoisotopicMass = Round(.ChargeStateBasedStats(intChargeIndex).Mass, MASS_PRECISION)
            udtPutUMCCSStatsParams.Abundance = .ChargeStateBasedStats(intChargeIndex).Abundance
            
            lngClassMIndexPointer = .ChargeStateBasedStats(intChargeIndex).GroupRepIndex
             
            If lngClassMIndexPointer < 0 Then
                udtPutUMCCSStatsParams.ElutionTime = 0
                udtPutUMCCSStatsParams.DriftTime = 0
            Else
                Select Case .ClassMType(lngClassMIndexPointer)
                Case glCSType
                    udtPutUMCCSStatsParams.ElutionTime = Round(ScanToGANET(lngGelIndex, GelData(lngGelIndex).CSData(.ClassMInd(lngClassMIndexPointer)).ScanNumber), NET_PRECISION)
                    udtPutUMCCSStatsParams.DriftTime = GelData(lngGelIndex).CSData(lngClassMIndexPointer).IMSDriftTime
                
                Case glIsoType
                    udtPutUMCCSStatsParams.ElutionTime = Round(ScanToGANET(lngGelIndex, GelData(lngGelIndex).IsoData(.ClassMInd(lngClassMIndexPointer)).ScanNumber), NET_PRECISION)
                    udtPutUMCCSStatsParams.DriftTime = GelData(lngGelIndex).IsoData(lngClassMIndexPointer).IMSDriftTime
                
                End Select
            End If
            
            cmdPutNewUMCCSStats.Execute
            
        Next intChargeIndex

    End With
    

    Exit Sub

AddUMCCSStatsErrorHandler:
    ' Error populating or executing cmdPutNewUMCCSStats
    
    Debug.Print "Error occurred in ExportMTDBAddUMCCSStatsRow: Err Code = " & Err.Number & vbCrLf & Err.Description
    Debug.Assert False
    
    Err.Raise Err.Number
    
End Sub

Private Sub ExportMTDBAddUMCResultMemberRow( _
            ByRef cmdPutNewUMCMember As ADODB.Command, _
            ByRef udtPutUMCMemberParams As udtPutUMCMemberParamsListType, _
            ByVal lngGelIndex As Long, _
            ByVal lngUMCIndexOriginal As Long, _
            ByVal lngUMCResultsIDInDB As Long)
            
    ' Adds row to T_FTICR_UMC_Members table
    
    Dim lngMemberIndex As Long, lngDataIndex As Long
    Dim intChargeIndex As Integer
    
On Error GoTo AddUMCMembersErrorHandler

    ' Now add the members of the UMC to T_FTICR_UMC_Members (if the table exists in the database)
    ' Store the UMCResultsID value
    udtPutUMCMemberParams.UMCResultsID.Value = lngUMCResultsIDInDB
    
    With GelUMC(lngGelIndex).UMCs(lngUMCIndexOriginal)
        For lngMemberIndex = 0 To .ClassCount - 1
            udtPutUMCMemberParams.IndexInUMC = lngMemberIndex
            lngDataIndex = .ClassMInd(lngMemberIndex)
            
            Select Case .ClassMType(lngMemberIndex)
            Case gldtCS
                udtPutUMCMemberParams.MemberTypeID = gldtCS
            
                udtPutUMCMemberParams.ScanNumber = GelData(lngGelIndex).CSData(lngDataIndex).ScanNumber
            
                udtPutUMCMemberParams.MZ = GelData(lngGelIndex).CSData(lngDataIndex).AverageMW
                udtPutUMCMemberParams.ChargeState = GelData(lngGelIndex).CSData(lngDataIndex).Charge
                udtPutUMCMemberParams.MonoisotopicMass = GelData(lngGelIndex).CSData(lngDataIndex).AverageMW
                udtPutUMCMemberParams.Abundance = GelData(lngGelIndex).CSData(lngDataIndex).Abundance
                udtPutUMCMemberParams.IsotopicFit = GelData(lngGelIndex).CSData(lngDataIndex).MassStDev
                udtPutUMCMemberParams.ElutionTime = ScanToGANET(lngGelIndex, GelData(lngGelIndex).CSData(lngDataIndex).ScanNumber)
            
            Case gldtIS
                udtPutUMCMemberParams.MemberTypeID = gldtIS
            
                udtPutUMCMemberParams.ScanNumber = GelData(lngGelIndex).IsoData(lngDataIndex).ScanNumber
            
                udtPutUMCMemberParams.MZ = GelData(lngGelIndex).IsoData(lngDataIndex).MZ
                udtPutUMCMemberParams.ChargeState = GelData(lngGelIndex).IsoData(lngDataIndex).Charge
                udtPutUMCMemberParams.MonoisotopicMass = GelData(lngGelIndex).IsoData(lngDataIndex).MonoisotopicMW
                udtPutUMCMemberParams.Abundance = GelData(lngGelIndex).IsoData(lngDataIndex).Abundance
                udtPutUMCMemberParams.IsotopicFit = GelData(lngGelIndex).IsoData(lngDataIndex).Fit
                udtPutUMCMemberParams.ElutionTime = ScanToGANET(lngGelIndex, GelData(lngGelIndex).IsoData(lngDataIndex).ScanNumber)
            
            Case Else
                ' This shouldn't happen; don't export data point if .ClassMType(lngMemberIndex) = 0
                Debug.Assert False
                udtPutUMCMemberParams.MemberTypeID = 0
            End Select
            
            If udtPutUMCMemberParams.MemberTypeID > 0 Then
            
                ' Check whether or not data point is the Charge State Based Stats group rep (most abundant point within the charge state)
                udtPutUMCMemberParams.IsChargeStateRep = 0
                For intChargeIndex = 0 To .ChargeStateCount - 1
                    If lngMemberIndex = .ChargeStateBasedStats(intChargeIndex).GroupRepIndex Then
                        udtPutUMCMemberParams.IsChargeStateRep = 1
                        Exit For
                    End If
                Next intChargeIndex
                    
                cmdPutNewUMCMember.Execute
            End If
        
        Next lngMemberIndex
        
    End With

    Exit Sub

AddUMCMembersErrorHandler:
    ' Error populating or executing cmdPutNewUMCMember
    
    Debug.Print "Error occurred in ExportMTDBAddUMCResultMemberRow: Err Code = " & Err.Number & vbCrLf & Err.Description
    Debug.Assert False
    
    Err.Raise Err.Number
    
End Sub

Public Function ExportMTDBAddQuantitationDescriptionEntry(ByRef frmCallingForm As Form, ByVal lngGelIndex As Long, ByVal strQuantitationDescriptionSP As String, ByVal lngMDID As Long, ByRef lngErrorNumber As Long, Optional ByVal strIniFileName As String = "", Optional ByVal intReplicate As Integer = 1, Optional ByVal intFraction As Integer = 1, Optional ByVal intTopLevelFraction As Integer = 1, Optional ByVal blnProcessImmediately As Boolean = False) As String
'---------------------------------------------------
'This function will add a new entry to T_Quantitation_Description and T_Quantitation_MDIDs in the Database
'
'Returns a status message
'lngErrorNumber will contain the error number, if an error occurs
'---------------------------------------------------

    ' Use a longer timeout than usual since Quantitation processing can take a while
    Const STORED_PROCEDURE_TIMEOUT_SEC = 600        ' 10 minutes
    Dim lngTimeoutOverride As Long
    
    Dim strSampleName As String
    Dim strComment As String
    Dim strExportStatus As String
    Dim lngQuantitationID As Long
    Dim lngQMDIDID As Long
    Dim lngEntriesProcessed As Long
    Dim lngSecElapsed As Long
    Dim strCaptionSaved As String
    Dim strCaptionAddOn As String
    
    'ADO objects for stored procedure
    Dim cnNew As New ADODB.Connection
    Dim cmdPutQuantitationDesc As New ADODB.Command
    Dim prmSampleName As New ADODB.Parameter
    Dim prmMDID As New ADODB.Parameter
    Dim prmReplicate As New ADODB.Parameter
    Dim prmFraction As New ADODB.Parameter
    Dim prmTopLevelFraction As New ADODB.Parameter
    Dim prmComment As New ADODB.Parameter
    Dim prmProcessImmediately As New ADODB.Parameter
    
    Dim prmQuantitationID As New ADODB.Parameter            ' Output
    Dim prmQ_MDID_ID As New ADODB.Parameter                 ' Output
    Dim prmEntriesProcessedReturn As New ADODB.Parameter    ' Output

    Dim prmLookupDefaultOptions As New ADODB.Parameter

    ' The remaining parameter are not defined or supplied to the SP since defaults are used instead
''    Dim prmFractionHighestAbuToUse As New ADODB.Parameter
''    Dim prmNormalizeToStandardAbundances As New ADODB.Parameter
''    Dim prmStandardAbundanceMin As New ADODB.Parameter
''    Dim prmStandardAbundanceMax As New ADODB.Parameter
''    Dim prmUMCAbundanceMode As New ADODB.Parameter
''    Dim prmExpressionRatioMode As New ADODB.Parameter
''    Dim prmMinimumHighNormalizedScore As New ADODB.Parameter
''    Dim prmMinimumHighDiscriminantScore As New ADODB.Parameter
''    Dim prmMinimumPMTQualityScore As New ADODB.Parameter
''    Dim prmMinimumPeptideLength As New ADODB.Parameter
''    Dim prmMinimumMatchScore As New ADODB.Parameter
''    Dim prmMinimumDelMatchScore As New ADODB.Parameter
''    Dim prmMinimumPeptideReplicateCount As New ADODB.Parameter
''    Dim prmORFCoverageComputationLevel As New ADODB.Parameter
''    Dim prmInternalStdInclusionMode As New ADODB.Parameter
''    Dim prmMinimumPeptideProphetProbability as New ADODB.Parameter
    
    On Error GoTo ExportMTDBAddQuantitationDescriptionEntryErrorHandler
    
    strCaptionSaved = frmCallingForm.Caption
    
    ' Define the sample name and comment
    ' The sample name should not contain any spaces, while the comment may
    ' Both have a maximum length of 255 characters
    
    strSampleName = GelAnalysis(lngGelIndex).Dataset
    If Len(strSampleName) = 0 Then
        strSampleName = GelAnalysis(lngGelIndex).MD_file
    End If
    If Len(strSampleName) = 0 Then
        strSampleName = "SampleEntered_" & Format(Now(), "yyyy-mm-dd_HH:nn_AmPm")
    End If
    If Len(strSampleName) > 255 Then strSampleName = Left(strSampleName, 255)
    
    strComment = "Job: " & GelAnalysis(lngGelIndex).MD_Reference_Job & ", File: " & GelAnalysis(lngGelIndex).MD_file
    If Len(strComment) > 255 Then strComment = Left(strComment, 255)
    
    ' Connect to the database
    frmCallingForm.Caption = "Connecting to the database"
    
    lngTimeoutOverride = STORED_PROCEDURE_TIMEOUT_SEC
    If lngTimeoutOverride < glbPreferencesExpanded.AutoAnalysisOptions.DBConnectionTimeoutSeconds Then
        lngTimeoutOverride = glbPreferencesExpanded.AutoAnalysisOptions.DBConnectionTimeoutSeconds
    End If
    
    If Not EstablishConnection(cnNew, GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, False, lngTimeoutOverride) Then
        Debug.Assert False
        lngErrorNumber = -1
        frmCallingForm.Caption = strCaptionSaved
        ExportMTDBAddQuantitationDescriptionEntry = "Error: Unable to establish a connection to the database"
        Exit Function
    End If
    
    ' Initialize the SP
    InitializeSPCommand cmdPutQuantitationDesc, cnNew, strQuantitationDescriptionSP
    
    ' Override the default timeout to use 620 seconds instead
    cmdPutQuantitationDesc.CommandTimeout = STORED_PROCEDURE_TIMEOUT_SEC + 20   ' Tack on an extra 20 seconds so an error doesn't occur in the Do-Loop
    
    Set prmSampleName = cmdPutQuantitationDesc.CreateParameter("SampleName", adVarChar, adParamInput, 255, strSampleName)
    cmdPutQuantitationDesc.Parameters.Append prmSampleName
    Set prmMDID = cmdPutQuantitationDesc.CreateParameter("MDID", adInteger, adParamInput, , lngMDID)
    cmdPutQuantitationDesc.Parameters.Append prmMDID
    
    Set prmReplicate = cmdPutQuantitationDesc.CreateParameter("Replicate", adSmallInt, adParamInput, , intReplicate)
    cmdPutQuantitationDesc.Parameters.Append prmReplicate
    Set prmFraction = cmdPutQuantitationDesc.CreateParameter("Fraction", adSmallInt, adParamInput, , intFraction)
    cmdPutQuantitationDesc.Parameters.Append prmFraction
    Set prmTopLevelFraction = cmdPutQuantitationDesc.CreateParameter("TopLevelFraction", adSmallInt, adParamInput, , intTopLevelFraction)
    cmdPutQuantitationDesc.Parameters.Append prmTopLevelFraction
    
    Set prmComment = cmdPutQuantitationDesc.CreateParameter("Comment", adVarChar, adParamInput, 255, strComment)
    cmdPutQuantitationDesc.Parameters.Append prmComment
    
    Set prmProcessImmediately = cmdPutQuantitationDesc.CreateParameter("ProcessImmediately", adTinyInt, adParamInput)
    If blnProcessImmediately Then
        prmProcessImmediately.Value = 1
    Else
        prmProcessImmediately.Value = 0
    End If
    cmdPutQuantitationDesc.Parameters.Append prmProcessImmediately
    
''    ' Legacy parameters; no longer supported
''    Set prmIniFileName = cmdPutQuantitationDesc.CreateParameter("IniFileName", adVarChar, adParamInput, 255, strIniFileName)
''    cmdPutQuantitationDesc.Parameters.Append prmIniFileName
    
''    ' In each of the Set calls, if .Value is left undefined, then the default value defined by the SP will be used
''    '
''    Set prmFractionHighestAbuToUse = cmdPutQuantitationDesc.CreateParameter("Fraction_Highest_Abu_To_Use", adDecimal, adParamInput)
''    With prmFractionHighestAbuToUse
''        .precision = 9
''        .NumericScale = 8
''        '' .value = 0.33     ' Default: 0.33
''    End With
''    cmdPutQuantitationDesc.Parameters.Append prmFractionHighestAbuToUse
''
''    Set prmNormalizeToStandardAbundances = cmdPutQuantitationDesc.CreateParameter("Normalize_To_Standard_Abundances", adTinyInt, adParamInput)      ' Default: 1
''    cmdPutQuantitationDesc.Parameters.Append prmNormalizeToStandardAbundances
''    Set prmStandardAbundanceMin = cmdPutQuantitationDesc.CreateParameter("Standard_Abundance_Min", adDouble, adParamInput)                          ' Default: 0
''    cmdPutQuantitationDesc.Parameters.Append prmStandardAbundanceMin
''    Set prmStandardAbundanceMax = cmdPutQuantitationDesc.CreateParameter("Standard_Abundance_Max", adDouble, adParamInput)                          ' Default: 5000000000
''    cmdPutQuantitationDesc.Parameters.Append prmStandardAbundanceMax
''    Set prmMinimumCriteriaORFMassDaDivisor = cmdPutQuantitationDesc.CreateParameter("Minimum_Criteria_ORFMassDaDivisor", adInteger, adParamInput)   ' Default: 15000
''    cmdPutQuantitationDesc.Parameters.Append prmMinimumCriteriaORFMassDaDivisor
''    Set prmMinimumCriteriaUniqueMTCountMinimum = cmdPutQuantitationDesc.CreateParameter("Minimum_Criteria_UniqueMTCountMinimum", adInteger, adParamInput)   ' Default: 2
''    cmdPutQuantitationDesc.Parameters.Append prmMinimumCriteriaUniqueMTCountMinimum
''    Set prmMinimumCriteriaMTIonMatchCountMinimum = cmdPutQuantitationDesc.CreateParameter("Minimum_Criteria_MTIonMatchCountMinimum", adInteger, adParamInput)   ' Default: 6
''    cmdPutQuantitationDesc.Parameters.Append prmMinimumCriteriaMTIonMatchCountMinimum
''    Set prmMinimumCriteriaFractionScansMatchingSingleMassTagMinimum = cmdPutQuantitationDesc.CreateParameter("Minimum_Criteria_FractionScansMatchingSingleMassTagMinimum", adDecimal, adParamInput)
''    With prmMinimumCriteriaFractionScansMatchingSingleMassTagMinimum
''        .precision = 9
''        .NumericScale = 8
''    ''    .value = 0.5     ' Default: 0.5
''    End With
''    cmdPutQuantitationDesc.Parameters.Append prmMinimumCriteriaFractionScansMatchingSingleMassTagMinimum
''
''    Set prmRemoveOutlierAbundancesForReplicates = cmdPutQuantitationDesc.CreateParameter("RemoveOutlierAbundancesForReplicates", adTinyInt, adParamInput)   ' Default: 1
''    cmdPutQuantitationDesc.Parameters.Append prmRemoveOutlierAbundancesForReplicates
''
''    Set prmFractionCrossReplicateAvgInRange = cmdPutQuantitationDesc.CreateParameter("FractionCrossReplicateAvgInRange", adDecimal, adParamInput)
''    With prmFractionCrossReplicateAvgInRange
''        .precision = 9
''        .NumericScale = 5
''    ''    .value = 0.8     ' Default: 0.8
''    End With
''    cmdPutQuantitationDesc.Parameters.Append prmFractionCrossReplicateAvgInRange
    
    
    Set prmQuantitationID = cmdPutQuantitationDesc.CreateParameter("Quantitation_ID", adInteger, adParamOutput)
    cmdPutQuantitationDesc.Parameters.Append prmQuantitationID
    Set prmQ_MDID_ID = cmdPutQuantitationDesc.CreateParameter("Q_MDID_ID", adInteger, adParamOutput)
    cmdPutQuantitationDesc.Parameters.Append prmQ_MDID_ID
    
    Set prmEntriesProcessedReturn = cmdPutQuantitationDesc.CreateParameter("EntriesProcessedReturn", adInteger, adParamOutput)     ' This value will be populated if @ProcessImmediately = 1
    cmdPutQuantitationDesc.Parameters.Append prmEntriesProcessedReturn
    
    ' Instruct the SP to lookup the options in T_Quantitation_Defaults
    Set prmLookupDefaultOptions = cmdPutQuantitationDesc.CreateParameter("LookupDefaultOptions", adTinyInt, adParamInput, , 1)
    cmdPutQuantitationDesc.Parameters.Append prmLookupDefaultOptions
    
    
    frmCallingForm.Caption = "Adding Quantitation Entry"
    
    ' Actually call the SP, using an asynchronous call so that we can provide
    '  feedback to the user as we wait for it to finish
    cmdPutQuantitationDesc.Execute , , adAsyncExecute
    
    lngSecElapsed = 0
    Do
        If lngSecElapsed >= STORED_PROCEDURE_TIMEOUT_SEC Then Exit Do
            
        ' Sleep for 1 second
        Sleep 1000
        lngSecElapsed = lngSecElapsed + 1
        
        If lngSecElapsed > 1 Then
            ' Append a . to the caption every second
            strCaptionAddOn = strCaptionAddOn & "."
            If Len(strCaptionAddOn) > 10 Then strCaptionAddOn = ""
            frmCallingForm.Caption = "Summarizing Protein Abundances " & strCaptionAddOn
            DoEvents
        End If
    Loop While cmdPutQuantitationDesc.STATE = adStateExecuting
    
    ' The following 3 variables are outputs from the SP
    lngQuantitationID = FixNullLng(prmQuantitationID.Value)             ' New Quantitation_ID
    lngQMDIDID = FixNullLng(prmQ_MDID_ID.Value)                         ' Index of the new row in T_Quantitation_MDIDs
    lngEntriesProcessed = FixNullLng(prmEntriesProcessedReturn.Value)   ' Quantitation_ID entries processed (typically 1)
    
    Set cmdPutQuantitationDesc.ActiveConnection = Nothing
    cnNew.Close
    
    ' Add an entry to the analysis history
    strExportStatus = "Added entry to T_Quantation_Description and T_Quantitation_MDIDs in the database (" & ExtractDBNameFromConnectionString(GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString) & "); Quantitation_ID = " & lngQuantitationID & "; Q_MDID_ID = " & lngQMDIDID
    If blnProcessImmediately Then
        strExportStatus = strExportStatus & "; QuantitationProcessStart processed " & Trim(lngEntriesProcessed) & " entries"
    End If
    AddToAnalysisHistory lngGelIndex, strExportStatus
    
    ExportMTDBAddQuantitationDescriptionEntry = strExportStatus
    lngErrorNumber = 0
    frmCallingForm.Caption = strCaptionSaved
    Exit Function
    
ExportMTDBAddQuantitationDescriptionEntryErrorHandler:
    Debug.Print "Error, probably timeout: Err Code = " & Err.Number & vbCrLf & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "ExportMTDBAddQuantitationDescriptionEntry (Possibly timeout error)"
    ExportMTDBAddQuantitationDescriptionEntry = "Error: " & Err.Number & vbCrLf & Err.Description
    lngErrorNumber = Err.Number
    On Error Resume Next
    If Not cnNew Is Nothing Then cnNew.Close
    frmCallingForm.Caption = strCaptionSaved

End Function


Attribute VB_Name = "AutoAnalysisProcedures"
Option Explicit

Private Const DATAFILE_LOAD_ERROR_BIT = 1               ' 2^0
Private Const INIFILE_LOAD_ERROR_BIT = 2                ' 2^1
Private Const DATABASE_ERROR_BIT = 4                    ' 2^2
Private Const UMC_COUNT_ERROR_BIT = 8                   ' 2^3
Private Const GANET_ERROR_BIT = 16                      ' 2^4
Private Const SEARCH_ERROR_BIT = 32                     ' 2^5
Private Const EXPORTRESULTS_ERROR_BIT = 64              ' 2^6
Private Const TOLERANCE_REFINEMENT_ERROR_BIT = 128      ' 2^7
Private Const SAVE_GRAPHIC_ERROR_BIT = 256              ' 2^8
Private Const SAVE_ERROR_DISTRIBUTION_ERROR_BIT = 512   ' 2^9
Private Const SAVE_CHROMATOGRAM_ERROR_BIT = 1024        ' 2^10
Private Const MASS_TAGS_NULL_COUNTS_HIGH_ERROR_BIT = 2048   '2^11
Private Const PAIRS_BASED_DB_SEARCH_ERROR_BIT = 4096        '2^12

Private Const UMC_SEARCH_ABORTED_WARNING_BIT = 1        ' 2^0
Private Const NET_ADJUSTMENT_SKIPPED_WARNING_BIT = 2    ' 2^1
Private Const GANET_SLOPE_WARNING_BIT = 4               ' 2^2
Private Const GANET_INTERCEPT_WARNING_BIT = 8           ' 2^3
Private Const NET_ADJUSTMENT_LOW_ID_COUNT_WARNING_BIT = 16          ' 2^4
Private Const TOLERANCE_REFINEMENT_WARNING_PEAK_NOT_FOUND_BIT = 32  ' 2^5
Private Const TOLERANCE_REFINEMENT_WARNING_BIT_PEAK_TOO_WIDE = 64   ' 2^6
Private Const PICTURE_FORMAT_WARNING_BIT = 128                      ' 2^7
Private Const INVALID_EXPORT_OPTION_WARNING_BIT = 256               ' 2^8
Private Const NO_DATABASE_HITS_WARNING_BIT = 512                    ' 2^9
Private Const PATH_FILE_ERROR_WARNING_BIT = 1024                    ' 2^10
Private Const NO_PAIRS_WARNING_BIT = 2048                           ' 2^11
Private Const MISCELLANEOUS_MESSAGE_WARNING_BIT = 4096              ' 2^12

Private Const HTML_INDEX_FILE_DATA_FILE_LINE_START As String = "Data file:"
Private Const HTML_INDEX_FILE_NAME As String = "Index.html"
Private Const HTML_INDEX_FOLDER_LIST_TABLE_START As String = "<TR><TH>Folder</TH>"
Private Const HTML_INDEX_FOLDER_LIST_TABLE_END As String = "</TABLE>"

Private Const HTML_SUMMARY_FILE_HEADING_NET_ALIGNMENT As String = "NET Alignment Surface"
Private Const HTML_SUMMARY_FILE_HEADING_NET_ALIGNMENT_MASS_CAL As String = "NET Alignment Surface with Mass Calibration"

Private Const HTML_SUMMARY_FILE_HEADING_NET_RESIDUALS As String = "NET Alignment Residuals"
Private Const HTML_SUMMARY_FILE_HEADING_NET_RESIDUALS_MASS_CAL As String = "NET Alignment Residuals with Mass Calibration"

Private Const EXPORT_TO_DB_PASSWORD As String = "mt4real"
Private Const RESIDUALS_PLOT_POINT_SIZE As Integer = 2

Private Const CTL2DHEATMAP_ERROR_MESSAGE As String = "Failed to load control 'ctl2DHeatMap'"

Private Type udtAutoAnalysisOutputFileInfoType
    FileName As String
    Description As String
    width As Integer                ' Typically 420
    TableRow As Integer             ' 1, 2, 3, etc.
    TableColumn As Integer          ' 1, 2, etc.
End Type

Private Type udtAutoAnalysisWorkingParamsType
    GelIndex As Long
    ErrorBits As Long
    WarningBits As Long
    ResultsFileNameBase As String
    GelOutputFolder As String
    GelFilePath As String
    GraphicOutputFileInfoCount As Integer
    GraphicOutputFileInfo() As udtAutoAnalysisOutputFileInfoType            ' 0-based array
    TICPlotsStartRow As Integer                         ' Defaults to 4, but gets bumped up to 5 if NET Alignment plots are included
    ts As TextStream                                    ' LogFile TextStream
    NextHistoryIndexToCopy As Long                      ' Next entry in GelSearchDef().AnalysisHistory that needs to be copied to the log
    LoadedGelFile As Boolean
    NETDefined As Boolean
End Type


Public Const DB_SEARCH_MODE_MAX_INDEX = 13
Public Const DB_SEARCH_MODE_PAIR_MODE_START_INDEX = 5
Public Enum dbsmDatabaseSearchModeConstants
    dbsmNone = -1
    dbsmExportUMCsOnly = 0
    dbsmIndividualPeaks = 1
    dbsmIndividualPeaksInUMCsWithoutNET = 2
    dbsmIndividualPeaksInUMCsWithNET = 3        ' No longer supported (June 2004)
    dbsmConglomerateUMCsWithNET = 4
    dbsmIndividualPeaksInUMCsPaired = 5         ' No longer supported (June 2004)
    dbsmIndividualPeaksInUMCsUnpaired = 6       ' No longer supported (June 2004)
    dbsmConglomerateUMCsPaired = 7
    dbsmConglomerateUMCsUnpaired = 8
    dbsmPairsN14N15 = 9                         ' No longer supported (July 2004)
    dbsmPairsN14N15ConglomerateMass = 10
    dbsmPairsICAT = 11
    dbsmPairsPEO = 12
    dbsmConglomerateUMCsLightPairsPlusUnpaired = 13
End Enum

Private Enum ehmErrorHistogramModeConstants
    ehmBeforeRefinement = 0
    ehmAfterLCMSWARP = 1
    ehmFinalTolerances = 2
End Enum

Private mRedirectedOutputFolderMessage As String

' mMemoryLog holds all of the data sent to AutoAnalysisLog
Private mMemoryLog As String

Private Sub AddNewOutputFileForHtml(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByVal strFileName As String, ByVal strDescription As String, ByVal intTableRow As Integer, ByVal intTableColumn As Integer, Optional ByVal intWidth As Integer = 450)

    With udtWorkingParams
        .GraphicOutputFileInfoCount = .GraphicOutputFileInfoCount + 1
        
        If .GraphicOutputFileInfoCount <= 1 Then
            .GraphicOutputFileInfoCount = 1
            ReDim .GraphicOutputFileInfo(0)
        Else
            ReDim Preserve .GraphicOutputFileInfo(.GraphicOutputFileInfoCount - 1)
        End If
        
        With .GraphicOutputFileInfo(.GraphicOutputFileInfoCount - 1)
            .FileName = strFileName
            .Description = strDescription
            .TableRow = intTableRow
            .TableColumn = intTableColumn
            .width = intWidth
        End With
    End With

End Sub

''Public Sub PrintAsciiCodes()
''    Dim intIndex As Integer
''    For intIndex = 33 To 255
''        If intIndex <> 127 Then
''            Debug.Print intIndex & ": " & Chr(intIndex)
''        End If
''    Next intIndex
''End Sub

Public Sub ApplyAutoAnalysisFilter(FilterPrefs As udtAutoAnalysisFilterPrefsType, lngGelIndex As Long, Optional blnAddCommentsToAnalysisHistory As Boolean = True)
    
    With GelData(lngGelIndex)
        .DataFilter(fltDupTolerance, 0) = FilterPrefs.ExcludeDuplicates
        .DataFilter(fltDupTolerance, 1) = FilterPrefs.ExcludeDuplicatesTolerance
        
        .DataFilter(fltIsoFit, 0) = FilterPrefs.ExcludeIsoByFit
        .DataFilter(fltIsoFit, 1) = FilterPrefs.ExcludeIsoByFitMaxVal
        
        .DataFilter(fltCSStDev, 0) = FilterPrefs.ExcludeCSByStdDev
        .DataFilter(fltCSStDev, 1) = FilterPrefs.ExcludeCSByStdDevMaxVal
        
        If FilterPrefs.ExcludeIsoSecondGuess Then
            .DataFilter(fltCase2CloseResults, 0) = True
            .DataFilter(fltCase2CloseResults, 1) = 1
        ElseIf FilterPrefs.ExcludeIsoLessLikelyGuess Then
            .DataFilter(fltCase2CloseResults, 0) = True
            .DataFilter(fltCase2CloseResults, 1) = 2
        Else
            .DataFilter(fltCase2CloseResults, 0) = False
            .DataFilter(fltCase2CloseResults, 1) = 0
        End If

        .DataFilter(fltCSAbu, 0) = FilterPrefs.RestrictCSByAbundance
        .DataFilter(fltCSAbu, 1) = FilterPrefs.RestrictCSAbundanceMin
        .DataFilter(fltCSAbu, 2) = FilterPrefs.RestrictCSAbundanceMax
        
        .DataFilter(fltIsoAbu, 0) = FilterPrefs.RestrictIsoByAbundance
        .DataFilter(fltIsoAbu, 1) = FilterPrefs.RestrictIsoAbundanceMin
        .DataFilter(fltIsoAbu, 2) = FilterPrefs.RestrictIsoAbundanceMax
        
        ' If filtering either Iso by Mass or CS by Mass, then apply to both
        If FilterPrefs.RestrictIsoByMass And Not FilterPrefs.RestrictCSByMass Then
            FilterPrefs.RestrictCSByMass = True
            FilterPrefs.RestrictCSMassMin = FilterPrefs.RestrictIsoMassMin
            FilterPrefs.RestrictCSMassMax = FilterPrefs.RestrictIsoMassMax
        ElseIf Not FilterPrefs.RestrictIsoByMass And FilterPrefs.RestrictCSByMass Then
            FilterPrefs.RestrictIsoByMass = True
            FilterPrefs.RestrictIsoMassMin = FilterPrefs.RestrictCSMassMin
            FilterPrefs.RestrictIsoMassMax = FilterPrefs.RestrictCSMassMax
        End If
        
        .DataFilter(fltCSMW, 0) = FilterPrefs.RestrictCSByMass
        .DataFilter(fltCSMW, 1) = FilterPrefs.RestrictCSMassMin
        .DataFilter(fltCSMW, 2) = FilterPrefs.RestrictCSMassMax
        
        .DataFilter(fltIsoMW, 0) = FilterPrefs.RestrictIsoByMass
        .DataFilter(fltIsoMW, 1) = FilterPrefs.RestrictIsoMassMin
        .DataFilter(fltIsoMW, 2) = FilterPrefs.RestrictIsoMassMax
        
        .DataFilter(fltIsoMZ, 0) = FilterPrefs.RestrictIsoByMZ
        .DataFilter(fltIsoMZ, 1) = FilterPrefs.RestrictIsoMZMin
        .DataFilter(fltIsoMZ, 2) = FilterPrefs.RestrictIsoMZMax
        
        .DataFilter(fltIsoCS, 0) = FilterPrefs.RestrictIsoByChargeState
        .DataFilter(fltIsoCS, 1) = FilterPrefs.RestrictIsoChargeStateMin
        .DataFilter(fltIsoCS, 2) = FilterPrefs.RestrictIsoChargeStateMax
        
        If FilterPrefs.RestrictToEvenScanNumbersOnly Then
            .DataFilter(fltEvenOddScanNumber, 0) = True
            .DataFilter(fltEvenOddScanNumber, 1) = 2
        ElseIf FilterPrefs.RestrictToOddScanNumbersOnly Then
            .DataFilter(fltEvenOddScanNumber, 0) = True
            .DataFilter(fltEvenOddScanNumber, 1) = 1
        Else
            .DataFilter(fltEvenOddScanNumber, 0) = False
            .DataFilter(fltEvenOddScanNumber, 1) = 0
        End If

    End With
    
    If blnAddCommentsToAnalysisHistory Then
        With FilterPrefs
            If .ExcludeDuplicates Then AddToAnalysisHistory lngGelIndex, "Filter applied: Exclude Duplicates; tolerance = " & .ExcludeDuplicatesTolerance
            If .ExcludeIsoByFit Then AddToAnalysisHistory lngGelIndex, "Filter applied: Exclude by Isotopic Fit; max fit = " & .ExcludeIsoByFitMaxVal
            If .ExcludeIsoSecondGuess Then AddToAnalysisHistory lngGelIndex, "Filter applied: Exclude Second Guess"
            If .ExcludeIsoLessLikelyGuess Then AddToAnalysisHistory lngGelIndex, "Filter applied: Exclude Less Likely Guess"
            If .ExcludeCSByStdDev Then AddToAnalysisHistory lngGelIndex, "Filter applied: Exclude CS by Std Dev; max std dev  = " & .ExcludeCSByStdDevMaxVal
            If .RestrictIsoByAbundance Then AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict Iso by Abundance; min abundance = " & .RestrictIsoAbundanceMin & "; max abundance = " & .RestrictIsoAbundanceMax
            If .RestrictIsoByMass Then AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict Iso by Mass; min mass = " & .RestrictIsoMassMin & "; max mass = " & .RestrictIsoMassMax
            If .RestrictIsoByMZ Then AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict Iso by m/z; min m/z = " & .RestrictIsoMZMin & "; max m/z = " & .RestrictIsoMZMax
            If .RestrictIsoByChargeState Then AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict Iso by Charge State; min charge state = " & .RestrictIsoChargeStateMin & "; max charge state = " & .RestrictIsoChargeStateMax
            If .RestrictCSByAbundance Then AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict CS by Abundance; min abundance = " & .RestrictCSAbundanceMin & "; max abundance = " & .RestrictCSAbundanceMax
            If .RestrictCSByMass Then AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict CS by Mass; min mass = " & .RestrictCSMassMin & "; max mass = " & .RestrictCSMassMax
            If .RestrictToEvenScanNumbersOnly Then
                AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict data to even-numbered scans only"
            ElseIf .RestrictToOddScanNumbersOnly Then
                AddToAnalysisHistory lngGelIndex, "Filter applied: Restrict data to odd-numbered scans only"
            End If
        End With
    End If
    
End Sub

Public Function AutoAnalysisStart(ByRef udtAutoParams As udtAutoAnalysisParametersType, Optional ByVal blnClearMemoryLog As Boolean = True, Optional ByRef blnPrintDebugInfo As Boolean = False) As Boolean
    ' Automatically load and analyze a .PEK/.CSV/.mzXML/.mzData file
    ' If udtAutoParams.InputFilePath is blank and udtAutoParams.ShowMessages = True, then the user is shown an open file box
    ' If strOptionsIniFile is blank, then the currently enabled options are used (i.e. UMC searching, NET Adjustment, DB searching)
    ' If udtAutoParams.GelIndexToForce is > 0 then the data will be loaded into the gel with the given index
    '
    ' For true automation, the file given by strOptionsIniFile should be customized with the desired options and sent to this function
    ' If udtAutoParams.AutoDMSAnalysisManuallyInitiated = True then the user manually chose the file using New Analysis, but requested auto-analysis by choosing a .Ini file
    '
    ' This function returns True on success and False on failure
    
    Dim fso As New FileSystemObject
    
    Dim udtWorkingParams As udtAutoAnalysisWorkingParamsType
    
    Dim blnLogFileIsInTempFolder As Boolean
    Dim blnSuccess As Boolean
    Dim blnDBReadyToLoad As Boolean, blnDBLoaded As Boolean
    Dim blnToleranceRefinementComplete As Boolean
    Dim strErrorMessage As String
    
    Dim blnRefinementWasPerformed As Boolean
    Dim blnAbortSinceNETNotDefined As Boolean
    Dim intHtmlFileColumnOverride As Integer
    
On Error GoTo AutoAnalysis_ErrorHandler
    
    ' 0a. Initialize .AutoAnalysisCachedData
    AutoAnalysisInitializeCachedData glbPreferencesExpanded.AutoAnalysisCachedData
    
    ' 0b. Initialize udtWorkingParams
    AutoAnalysisInitializeWorkingParams udtWorkingParams
    
    ' 1. Initialize logging, both in memory and possibly to disk
    If blnPrintDebugInfo Then Debug.Print vbCrLf & Now() & " = " & "Initialize logging"
    blnSuccess = AutoAnalysisInitializeLogging(udtWorkingParams, udtAutoParams, fso, blnClearMemoryLog, blnLogFileIsInTempFolder)
    
    ' 2. Possibly check for the existence of the .Ini file
    If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Check for existence of .Ini file"
    If blnSuccess And udtAutoParams.FullyAutomatedPRISMMode Then
        If Not fso.FileExists(udtAutoParams.FilePaths.IniFilePath) Then
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - The specified options .Ini file does not exist: " & udtAutoParams.FilePaths.IniFilePath
            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or INIFILE_LOAD_ERROR_BIT
            blnSuccess = False
        Else
            blnSuccess = True
        End If
    End If

    ' 3. Load the PEK/CSV/mzXML/mzData/GEL file (if no error)
    If blnSuccess Then
        If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Load PEK/CSV/mzXML/mzData/GEL file"
        blnSuccess = AutoAnalysisLoadInputFile(udtWorkingParams, udtAutoParams, fso)
    End If
    
    ' Proceed if the file loaded successfully
    If blnSuccess And udtWorkingParams.GelIndex > 0 Then
    
    ' 4. Need to make sure the gel gets fully drawn in case any of the analysis functions depend on this
        ' I'm not sure if this is truly necessary, but it doesn't hurt
        If blnPrintDebugInfo Then Debug.Print Now() & " = " & "GelBody.ActivateGraph()"
        GelBody(udtWorkingParams.GelIndex).ActivateGraph False
        
    ' 5. Load the Auto Analysis .Ini file (if any)
        If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Load the .Ini file"
        blnDBReadyToLoad = AutoAnalysisLoadOptions(udtWorkingParams, udtAutoParams, fso)
        
    ' 6. Define the file output paths; Copy the .Ini file to the output folder
        If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Determine Output Paths"
        AutoAnalysisDefineFilePaths udtWorkingParams, udtAutoParams, fso
        
        ' Flush the log file
        AutoAnalysisFlushLogfile udtWorkingParams, udtAutoParams, fso
        
    ' 7. Possibly filter the data
        If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Filter Data"
        AutoAnalysisFilterData udtWorkingParams, udtAutoParams
    
    ' 8. Possibly filter out noise streaks
        AutoAnalysisFilterNoiseStreaks udtWorkingParams, udtAutoParams
        
    ' 9. Possibly (typically) Find the LC-MS Features
        If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Find LC-MS Features"
        AutoAnalysisFindUMCs udtWorkingParams, udtAutoParams
                
        ' Flush the log file
        AutoAnalysisFlushLogfile udtWorkingParams, udtAutoParams, fso
        If udtAutoParams.ExitViperASAP Then GoTo AutoAnalysis_CleanUp
        
    ' 10. Possibly Find the pairs
        AutoAnalysisFindPairs udtWorkingParams, udtAutoParams, fso
        
    ' 11. Attempt to actually load the MT tags
        If blnDBReadyToLoad Then
            If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Load MT tags"
            blnDBLoaded = AutoAnalysisLoadDB(udtWorkingParams, udtAutoParams, fso)
        End If
        
        ' Flush the log file
        AutoAnalysisFlushLogfile udtWorkingParams, udtAutoParams, fso
        If udtAutoParams.ExitViperASAP Then GoTo AutoAnalysis_CleanUp
        
    ' 12. If all is well, then perform NET adjustment and search the database
        If Not blnDBLoaded Then
            ' No DB: Log the error
            strErrorMessage = "Error - Database not in memory; unable to proceed with NET adjustment and DB search"
            If APP_BUILD_DISABLE_MTS Then
                strErrorMessage = strErrorMessage & ".  Be sure a Legacy MT database is defined in the .Ini file"
            End If
            AutoAnalysisLog udtAutoParams, udtWorkingParams, strErrorMessage
            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or DATABASE_ERROR_BIT
        ElseIf GelUMC(udtWorkingParams.GelIndex).UMCCnt <= 0 Then
            ' No LC-MS Features: Log the error
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - No LC-MS Features were found; unable to proceed with NET adjustment and DB search"
            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or UMC_COUNT_ERROR_BIT
        Else
        
            ' 13a. If a mass calibration override is defined, then apply it now, prior to NET adjustment
            blnToleranceRefinementComplete = False
            If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibration And _
               glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibrationOverridePPM <> 0 Then
                If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Tolerance Refinement"
                AutoAnalysisToleranceRefinement udtWorkingParams, udtAutoParams, fso
                blnToleranceRefinementComplete = True
            
                ' Flush the log file
                AutoAnalysisFlushLogfile udtWorkingParams, udtAutoParams, fso
                If udtAutoParams.ExitViperASAP Then GoTo AutoAnalysis_CleanUp
            End If

            ' 13b. Perform NET Adjustment
            If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Net Adjustment"
            AutoAnalysisPerformNETAdjustment udtWorkingParams, udtAutoParams, fso
            
            If udtAutoParams.ExitViperASAP Then GoTo AutoAnalysis_CleanUp
            
            If udtWorkingParams.NETDefined Then
                blnAbortSinceNETNotDefined = False
            Else
                blnAbortSinceNETNotDefined = True
                If glbPreferencesExpanded.AutoAnalysisOptions.IgnoreNETAdjustmentFailure Then
                    If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibration And _
                       GelUMCNETAdjDef(udtWorkingParams.GelIndex).UseRobustNETAdjustment And _
                       GelUMCNETAdjDef(udtWorkingParams.GelIndex).RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass And _
                       Not APP_BUILD_DISABLE_LCMSWARP Then
                        blnAbortSinceNETNotDefined = False
                    End If
                End If
            End If

            If blnAbortSinceNETNotDefined Then
                ' NET could not be defined; typically due to an error
                ' Cannot search the data
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - NET mapping not defined; unable to proceed with DB search"
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
            Else
                If udtWorkingParams.NETDefined Then
                    ' 14. Search for Internal Standards and save Picture and UMC text file to disk
                    If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Search for Internal Standards"
                    AutoAnalysisSaveInternalStdHits udtWorkingParams, udtAutoParams, fso
                End If
                
                ' 15. Tolerance Refinement
                If Not blnToleranceRefinementComplete Then
                    If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Tolerance Refinement"
                    AutoAnalysisToleranceRefinement udtWorkingParams, udtAutoParams, fso
                    blnToleranceRefinementComplete = True
                End If
            
                ' Flush the log file
                AutoAnalysisFlushLogfile udtWorkingParams, udtAutoParams, fso
                If udtAutoParams.ExitViperASAP Then GoTo AutoAnalysis_CleanUp
                
                ' 16. Search the Database for matches with the loaded data
                If blnPrintDebugInfo Then Debug.Print Now() & " = " & "Search DB"
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Database info: " & CurrMTDBInfo()
                AutoAnalysisSearchDatabase udtWorkingParams, udtAutoParams, fso
            End If
            
            ' Flush the log file
            AutoAnalysisFlushLogfile udtWorkingParams, udtAutoParams, fso
            If udtAutoParams.ExitViperASAP Then GoTo AutoAnalysis_CleanUp
            
            ' 17. Possibly save a graphical picture of the data
            AutoAnalysisSavePictureGraphic udtWorkingParams, udtAutoParams, fso
            
            ' 18. Possibly save a picture or text file of the mass and NET error distribution
            blnRefinementWasPerformed = CheckAutoToleranceRefinementEnabled(glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement, udtAutoParams, udtWorkingParams, False)
            If GelUMCNETAdjDef(udtWorkingParams.GelIndex).RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETIterative Then
                intHtmlFileColumnOverride = 2
            Else
                intHtmlFileColumnOverride = 0
            End If
            AutoAnalysisSaveErrorDistributions udtWorkingParams, udtAutoParams, fso, False, Not blnRefinementWasPerformed, ehmErrorHistogramModeConstants.ehmFinalTolerances, intHtmlFileColumnOverride
        End If
        
    ' 19. Possibly save the Points and/or LC-MS Features to disk
        AutoAnalysisSaveUMCsToDisk udtWorkingParams, udtAutoParams, fso

    ' 20. Possibly save a picture and text file of the various chromatograms
        AutoAnalysisSaveChromatograms udtWorkingParams, udtAutoParams, fso
    
    ' 21. Create a browsable Index.html file
        AutoAnalysisGenerateHTMLBrowsingFile udtWorkingParams, udtAutoParams, fso
    
    ' Obsolete: Update the browsable Index.html file up one folder level to include this output folder
    ' No longer need to do this since the Index Creator tool is updating these files every 4 hours on Albert
    '' AutoAnalysisUpdateParentHTMLBrowsingFile udtWorkingParams, udtAutoParams, fso
    
        ' Flush the log file
        AutoAnalysisFlushLogfile udtWorkingParams, udtAutoParams, fso
        If udtAutoParams.ExitViperASAP Then GoTo AutoAnalysis_CleanUp
    
    ' 22. Possibly save the Gel file to disk
        AutoAnalysisSaveGelFile udtWorkingParams
        
        ' Write the latest AnalysisHistory information to the log
        AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
        
    ' 23. Possibly close the loaded file; Change .Dirty to False so that user is not prompted about close
        If udtAutoParams.AutoCloseFileWhenDone Then
            GelStatus(udtWorkingParams.GelIndex).Dirty = False
            Unload GelBody(udtWorkingParams.GelIndex)
        End If
    End If
    
    ' 24. Write the ErrorCode and WarningCode lines to the log file
    AutoAnalysisLog udtAutoParams, udtWorkingParams, "ErrorCode=" & Trim(udtWorkingParams.ErrorBits)
    AutoAnalysisLog udtAutoParams, udtWorkingParams, "WarningCode=" & Trim(udtWorkingParams.WarningBits)
    
    If udtWorkingParams.ErrorBits <> 0 Then blnSuccess = False
    udtAutoParams.ErrorBits = udtWorkingParams.ErrorBits
    udtAutoParams.WarningBits = udtWorkingParams.WarningBits
    
    AutoAnalysisStart = blnSuccess
    
AutoAnalysis_CleanUp:
    ' 24. Cleanup steps
    AutoAnalysisCleanup udtWorkingParams, udtAutoParams, fso, blnLogFileIsInTempFolder
    Exit Function


AutoAnalysis_ErrorHandler:
    Debug.Assert False
    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - An error has occurred during auto analysis: " & Err.Description
    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - Aborting"
    If udtAutoParams.ShowMessages Then
        MsgBox "An error has occurred during auto analysis: " & Err.Description, vbExclamation + vbOKOnly, "Aborting"
    End If
    Resume AutoAnalysis_CleanUp
    
End Function

Private Sub AutoAnalysisAppendLatestHistoryToLog(ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType)
    
    Dim lngHistoryIndex As Long

On Error GoTo AutoAnalysisAppendLatestHistoryToLogErrorHandler

    With GelSearchDef(udtWorkingParams.GelIndex)
        For lngHistoryIndex = udtWorkingParams.NextHistoryIndexToCopy To .AnalysisHistoryCount - 1
            AutoAnalysisLog udtAutoParams, udtWorkingParams, .AnalysisHistory(lngHistoryIndex)
            udtWorkingParams.NextHistoryIndexToCopy = lngHistoryIndex + 1
        Next lngHistoryIndex
    End With

    Exit Sub
    
AutoAnalysisAppendLatestHistoryToLogErrorHandler:
    If Err.Number = 9 Then
        ' Subscript out of range; loading of .Pek, .CSV, .mzXML, or .mzData file probably failed
        Err.Clear
    Else
        Debug.Print "Error in AutoAnalysisAppendLatestHistory: " & Err.Description
        Debug.Assert False
        LogErrors Err.Number, "AutoAnalysisAppendLatestHistoryLog"
        Err.Clear
    End If
    
End Sub

Private Sub AutoAnalysisCleanup(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject, blnLogFileIsInTempFolder As Boolean)
    
    Dim strLogFilePathRenamed As String, strLogFilePathNew As String
    Dim strMessage As String
    
    On Error Resume Next
    If Not udtWorkingParams.ts Is Nothing Then
        udtWorkingParams.ts.Close
        Set udtWorkingParams.ts = Nothing
    End If
    
    ' If necessary, move the log file from the temp folder to strResultsFilePath's folder
    If blnLogFileIsInTempFolder And Len(udtWorkingParams.ResultsFileNameBase) > 0 Then
        ' First rename the log file to match udtWorkingParams.ResultsFileNameBase
        strLogFilePathRenamed = fso.BuildPath(fso.GetParentFolderName(udtAutoParams.FilePaths.LogFilePath), udtWorkingParams.ResultsFileNameBase & ".log")
        If fso.FileExists(strLogFilePathRenamed) Then
            fso.DeleteFile strLogFilePathRenamed, True
        End If
        fso.MoveFile udtAutoParams.FilePaths.LogFilePath, strLogFilePathRenamed
        strLogFilePathNew = fso.BuildPath(fso.GetParentFolderName(udtWorkingParams.GelFilePath), fso.GetFileName(strLogFilePathRenamed))
        fso.CopyFile strLogFilePathRenamed, strLogFilePathNew, True
        
        ' Update .FilePaths.LogFilePath
        udtAutoParams.FilePaths.LogFilePath = strLogFilePathNew
    End If
    
    If udtWorkingParams.ErrorBits <> 0 Or udtWorkingParams.WarningBits <> 0 Then
        If udtAutoParams.ShowMessages Then
            strMessage = ""
            With udtWorkingParams
                If .ErrorBits <> 0 Then
                    strMessage = "One or more errors occurred during auto analysis: " & vbCrLf
                    strMessage = strMessage & LookupErrorBitDescription(.ErrorBits)
                End If
                
                If .WarningBits <> 0 Then
                    If Len(strMessage) > 0 Then strMessage = strMessage & vbCrLf & vbCrLf
                    strMessage = strMessage & "One or more warnings were recorded during auto analysis: " & vbCrLf
                    strMessage = strMessage & LookupWarningBitDescription(.WarningBits)
                End If
            End With
            
            If (udtWorkingParams.WarningBits And PATH_FILE_ERROR_WARNING_BIT) = PATH_FILE_ERROR_WARNING_BIT And Len(mRedirectedOutputFolderMessage) > 0 Then
                If Len(strMessage) > 0 Then strMessage = strMessage & vbCrLf & vbCrLf
                strMessage = strMessage & mRedirectedOutputFolderMessage
            End If
            
            strMessage = strMessage & vbCrLf & vbCrLf & "Please see the log file for more information: " & vbCrLf & udtAutoParams.FilePaths.LogFilePath
            MsgBox strMessage, vbInformation Or vbOKOnly, "Auto Analysis Status"
        End If
    End If
    
    Set fso = Nothing
    With glbPreferencesExpanded.AutoAnalysisStatus
        .AutoAnalysisTimeStamp = ""
        .Enabled = False
    End With

End Sub

Private Sub AutoAnalysisDefineFilePaths(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
    
    Dim strTestFilePath As String
    Dim strIniFilePathCopyDestination As String
    Dim strErrorMessage As String
    
    Dim blnIgnoreErrors As Boolean
    Dim blnErrorOccurred As Boolean
    Dim ts As TextStream
    
On Error GoTo AutoAnalysisDefineFilePathsErrorHandler
    
    blnIgnoreErrors = False
    blnErrorOccurred = False
    mRedirectedOutputFolderMessage = ""
    
    ' Define the default gel output folder
    ' It will be .AlternateOutputFolderPath if that exists, or the parent folder
    '   of the input .Pek, .CSV, .mzXML, or .mzData file if .AlternateOutputFolderPath doesn't exist
    With glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(0)
        ' Note that any value defined in udtAutoParams.FilePaths.OutputFolderPath will
        '  have been copied to .AlternateOutputFolderPath in sub AutoAnalysisLoadOptions
        If Len(.AlternateOutputFolderPath) > 0 Then
            If Not fso.FolderExists(.AlternateOutputFolderPath) Then
                ' Folder not found; try to create it
                CreateFolderByPath .AlternateOutputFolderPath
            End If
        End If
        
        If fso.FolderExists(.AlternateOutputFolderPath) Then
            udtWorkingParams.GelOutputFolder = .AlternateOutputFolderPath
        Else
            udtWorkingParams.GelOutputFolder = fso.GetParentFolderName(udtAutoParams.FilePaths.InputFilePath)
            
            If Len(.AlternateOutputFolderPath) > 0 Then
                strErrorMessage = "Error - Output folder does not exist: " & .AlternateOutputFolderPath & " "
            Else
                strErrorMessage = "Warning - AlternateOutputFolderPath is blank (or missing) in the .Ini file"
            End If
            
            strErrorMessage = strErrorMessage + "; Will attempt to use folder: " & udtWorkingParams.GelOutputFolder
            AutoAnalysisLog udtAutoParams, udtWorkingParams, strErrorMessage
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or PATH_FILE_ERROR_WARNING_BIT
        End If
        
        ' Make sure we have write-permission to udtWorkingParams.GelOutputFolder
        ' Do this by creating an empty text file there, then deleting it
        ' If either operation fails, then change udtWorkingParams.GelOutputFolder to a subfolder in the local computer's temp folder
        strTestFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, "TempTest" & Trim(Abs(GetTickCount())) & ".tmp")
        
        blnIgnoreErrors = True
        
        Set ts = fso.CreateTextFile(strTestFilePath, True)
        If Not blnErrorOccurred Then
            ts.Write "Test"
            ts.Close
        End If
        Set ts = Nothing
        
        If Not blnErrorOccurred Then
            fso.DeleteFile strTestFilePath, True
        End If
        
        blnIgnoreErrors = False
        
        If blnErrorOccurred Then
            ' Unable to create a test text file in the output folder; use a different folder instead
            strTestFilePath = fso.GetParentFolderName(udtAutoParams.FilePaths.InputFilePath)
            If Len(strTestFilePath) = 0 Then
                ' This is unexpected; perhaps the .Pek, .CSV, .mzXML, or .mzData file is in the root of a drive (like C:\)
                Debug.Assert False
                strTestFilePath = "ViperAutoAnalysis_" & Format(Now(), "yyyymmdd_hhnnss")
            Else
                strTestFilePath = fso.GetBaseName(strTestFilePath)
            End If
            
            udtWorkingParams.GelOutputFolder = fso.BuildPath(GetTempFolder(), strTestFilePath)
            
            
            If Not fso.FolderExists(udtWorkingParams.GelOutputFolder) Then
                ' Folder not found; try to create it
                CreateFolderByPath udtWorkingParams.GelOutputFolder
            End If
            
            mRedirectedOutputFolderMessage = "Error - Write permission was denied in the output folder: " & fso.GetParentFolderName(udtAutoParams.FilePaths.InputFilePath) & vbCrLf & "  Will instead use the folder: " & udtWorkingParams.GelOutputFolder
            AutoAnalysisLog udtAutoParams, udtWorkingParams, mRedirectedOutputFolderMessage
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or PATH_FILE_ERROR_WARNING_BIT
        End If
        
    End With
    
    ' Define the base results file name
    udtWorkingParams.ResultsFileNameBase = fso.GetBaseName(udtAutoParams.FilePaths.InputFilePath)
    If glbPreferencesExpanded.AutoAnalysisOptions.ExportResultsFileUsesJobNumberInsteadOfDataSetName _
       And Not GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
        With GelAnalysis(udtWorkingParams.GelIndex)
            If .MD_Reference_Job > 0 Then
                udtWorkingParams.ResultsFileNameBase = "Job" & Trim(.MD_Reference_Job)
            End If
        End With
    End If
    
    With udtWorkingParams
        ' Define the gel file name
        .GelFilePath = fso.BuildPath(.GelOutputFolder, .ResultsFileNameBase & ".gel")
    
        ' Update the caption to reflect the target path (even if the gel will actually never get saved)
        GelBody(.GelIndex).Caption = .GelFilePath
        GelStatus(.GelIndex).GelFilePathFull = GetFilePathFull(.GelFilePath)
    End With
    
    ' Copy the .Ini file to the output folder
    If Len(udtAutoParams.FilePaths.IniFilePath) > 0 Then
        If fso.FileExists(udtAutoParams.FilePaths.IniFilePath) Then
            
            strIniFilePathCopyDestination = fso.BuildPath(udtWorkingParams.GelOutputFolder, fso.GetFileName(udtAutoParams.FilePaths.IniFilePath))
            
            blnErrorOccurred = False
            
            blnIgnoreErrors = True
            fso.CopyFile udtAutoParams.FilePaths.IniFilePath, strIniFilePathCopyDestination
            blnIgnoreErrors = False
            
            If blnErrorOccurred Then
                blnErrorOccurred = False
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Error occurred while trying to copy the .Ini file to the output folder: " & strIniFilePathCopyDestination
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or PATH_FILE_ERROR_WARNING_BIT
            End If
        End If
    End If
        
    Exit Sub

AutoAnalysisDefineFilePathsErrorHandler:
    If blnIgnoreErrors Then
        ' The error was probably permission denied (70)
        Debug.Assert Err.Number = 70
        
        ' Note that error 76 will result from "path too long", which results if the path length, including the .Ini file name, is over 255 characters
        
        blnErrorOccurred = True
        Err.Clear
        Resume Next
    Else
        
    End If
    
End Sub

Private Sub AutoAnalysisFilterData(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, Optional blnUpdateLog As Boolean = True)
    
    Dim lngGelScanNumberMin As Long, lngGelScanNumberMax As Long
    Dim strMessage As String
        
On Error GoTo FilterDataErrorHandler
        
    With glbPreferencesExpanded
        ' Rather than check whether or not filtering is enabled, simply set all of the
        '  options in GelData(udtWorkingParams.GelIndex)
        ApplyAutoAnalysisFilter .AutoAnalysisFilterPrefs, udtWorkingParams.GelIndex, blnUpdateLog
        
        ' Assign udtWorkingParams.GelIndex to frmFilter.Tag, then call .InitializeControls
        ' Since glbPreferencesExpanded.AutoAnalysisStatus.Enabled = True, will
        '  automatically unload the form after the filter is applied
        frmFilter.Tag = udtWorkingParams.GelIndex
        frmFilter.InitializeControls True
        
        ' If filtering by scan range, then apply that filter now
        With .AutoAnalysisFilterPrefs
            If .RestrictScanRange Then
                ' Determine the minimum and maximum scan numbers possible
                GetScanRange udtWorkingParams.GelIndex, lngGelScanNumberMin, lngGelScanNumberMax, 0, 0

                ValidateValueLng .RestrictScanRangeMin, lngGelScanNumberMin, lngGelScanNumberMax, lngGelScanNumberMin
                ValidateValueLng .RestrictScanRangeMax, lngGelScanNumberMin, lngGelScanNumberMax, lngGelScanNumberMax

                ZoomGelToDimensions udtWorkingParams.GelIndex, CSng(.RestrictScanRangeMin), 0, CSng(.RestrictScanRangeMax), 0
        
                If blnUpdateLog Then
                    AddToAnalysisHistory udtWorkingParams.GelIndex, "Scan range restricted; min scan = " & .RestrictScanRangeMin & "; max scan = " & .RestrictScanRangeMax
                End If
            End If
        End With
    End With
    
    If blnUpdateLog Then
        ' Write the latest AnalysisHistory information to the log
        AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
    End If
    
    Exit Sub
    
FilterDataErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while filtering data during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

Private Sub AutoAnalysisFilterNoiseStreaks(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType)
    
    Dim strMessage As String

On Error GoTo AutoAnalysisFilterNoiseStreaksErrorHandler
        
    With glbPreferencesExpanded
        If .AutoAnalysisOptions.AutoRemoveNoiseStreaks Then
            With frmExcludeMassRange
                .SetCallerID udtWorkingParams.GelIndex
                .Show vbModeless
                .InitializeForm
                .AutoPopulateStart True
                .IncludeExcludeIons True
            End With
            Unload frmExcludeMassRange
        End If
    End With
    
    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
    
    Exit Sub
    
AutoAnalysisFilterNoiseStreaksErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while filtering noise streaks during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    
End Sub

Private Sub AutoAnalysisFindPairs(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
    
    Dim objPairsSearchForm As frmUMCDltPairs
    Dim ePairFormMode As pfmPairFormMode
    Dim blnSuccess As Boolean
    Dim strPairsFilePath As String
    Dim strMessage As String
    
    Dim intIndex As Integer
    
On Error GoTo FindPairsErrorHandler

    ' Possibly find the Pairs
    Select Case LCase(glbPreferencesExpanded.PairSearchOptions.PairSearchMode)
    Case LCase(AUTO_FIND_PAIRS_NONE)
        ' Nothing to do
        ' Leave objPairsSearchForm undefined (aka Is Nothing)
    Case LCase(AUTO_FIND_PAIRS_DELTA)
        Set objPairsSearchForm = frmUMCDltPairs
        ePairFormMode = pfmDelta
    Case LCase(AUTO_FIND_PAIRS_LABEL)
        Set objPairsSearchForm = frmUMCDltPairs
        ePairFormMode = pfmLabel
    Case Else
        Debug.Assert False
    End Select

    If Not objPairsSearchForm Is Nothing Then
    
        With objPairsSearchForm
            .Tag = udtWorkingParams.GelIndex
            .FormMode = ePairFormMode
            .Show vbModeless
            .InitializeForm
            blnSuccess = .FindPairsWrapper(False)
        End With
        
        With glbPreferencesExpanded.PairSearchOptions
            If .AutoAnalysisDeltaMassAddnlCount > 0 Then
                
                ' Search for the additional delta masses defined in .AutoAnalysisDeltaMassAddnl()
                For intIndex = 0 To .AutoAnalysisDeltaMassAddnlCount - 1
                    ' Disable clearing existing pairs when searching for new pairs
                    objPairsSearchForm.AutoClearPairsWhenFindingPairs = False
                    ' Update the delta mass
                    objPairsSearchForm.SetDeltaMass .AutoAnalysisDeltaMassAddnl(intIndex)
                    ' Search for new pairs
                    objPairsSearchForm.FindPairsWrapper (False)
                Next intIndex
            End If
        End With
            
        If glbPreferencesExpanded.PairSearchOptions.AutoExcludeOutOfERRange Then
            objPairsSearchForm.MarkBadERPairs
        Else
            ' Change .Pairs().State to glPAIR_Inc for all pairs with .State = glPAIR_Neu
            PairSearchIncludeNeutralPairs udtWorkingParams.GelIndex
        End If
        
        If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisSavePairsToTextFile Then
            strPairsFilePath = udtWorkingParams.ResultsFileNameBase & "_Pairs.txt"
            strPairsFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strPairsFilePath)
            objPairsSearchForm.ReportPairs 0, strPairsFilePath
        End If
        
        If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisSavePairsStatisticsToTextFile Then
            strPairsFilePath = udtWorkingParams.ResultsFileNameBase & "_PairStatistics.txt"
            strPairsFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strPairsFilePath)
            objPairsSearchForm.ReportERStatistics strPairsFilePath
        End If
        
        If glbPreferencesExpanded.PairSearchOptions.AutoExcludeOutOfERRange Then
            strMessage = DeleteExcludedPairs(udtWorkingParams.GelIndex)
            AddToAnalysisHistory udtWorkingParams.GelIndex, strMessage
        End If
        
        Unload objPairsSearchForm
        Set objPairsSearchForm = Nothing
    
        ' Write the latest AnalysisHistory information to the log
        AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
        
        If GelP_D_L(udtWorkingParams.GelIndex).PCnt = 0 Then
            ' No Pairs were found
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - No Pairs were found "
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or NO_PAIRS_WARNING_BIT
        Else
            If Not blnSuccess Then
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - pairs searching was aborted prematurely"
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or UMC_SEARCH_ABORTED_WARNING_BIT
            End If
        End If
    End If
    
    Exit Sub

FindPairsErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while finding Pairs during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    
End Sub

Private Sub AutoAnalysisFindUMCs(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType)
    
    Dim objUMCSearchForm As VB.Form
    Dim strMessage As String
    Dim blnSuccess As Boolean
    
On Error GoTo FindUMCsErrorHandler
    
    If glbPreferencesExpanded.AutoAnalysisOptions.SkipFindUMCs Then
        If GelUMC(udtWorkingParams.GelIndex).UMCCnt = 0 Then
            ' No LC-MS Features are present in memory
            strMessage = "The option SkipFindUMCs was set to True, but no LC-MS Features are present in memory -- unable to continue auto analysis."
            AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
            If udtAutoParams.ShowMessages Then MsgBox strMessage, vbExclamation + vbOKOnly, "No LC-MS Features"
        Else
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Skipped searching for LC-MS Features (" & Trim(GelUMC(udtWorkingParams.GelIndex).UMCCnt) & " LC-MS Features already in memory)"
            blnSuccess = True
        End If
    Else
        ' Note: If you want to alter the UMC finding mode while debugging, here is your chance to do so
        '       (change from False to True in the following If statement, or just move the cursor to the .UMCSearchMode = AUTO_ANALYSIS_UMC2003 line)
        Debug.Assert False
        If False Then
            ' For example, change to UMC2003 using:
            glbPreferencesExpanded.AutoAnalysisOptions.UMCSearchMode = AUTO_ANALYSIS_UMC2003
            
            ' Also, use this to assure the data is not exported to the database
            glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(0).ExportResultsToDatabase = False
        End If
        
        ' Find the LC-MS Features (Note: Using LCase to avoid case conversion problems)
        Select Case LCase(glbPreferencesExpanded.AutoAnalysisOptions.UMCSearchMode)
        Case LCase(AUTO_ANALYSIS_UMCListType2002)
            ' Old mass-window-limited search method; no longer supported by auto-analysis; use frmUMCSimple instead
            Set objUMCSearchForm = frmUMCSimple
        Case LCase(AUTO_ANALYSIS_UMC2003)
            ' Improved mass-window-limited search method
            Set objUMCSearchForm = frmUMCSimple
        Case Else
            ' Includes AUTO_ANALYSIS_UMCIonNet
            Set objUMCSearchForm = frmUMCIonNet
        End Select
        
        With objUMCSearchForm
            If GelSearchDef(udtWorkingParams.GelIndex).UMCDef.MWField = 0 Then
                Debug.Assert False
                GelSearchDef(udtWorkingParams.GelIndex).UMCDef.MWField = mftMWMono
            End If
            
            .Tag = udtWorkingParams.GelIndex
            ' Note: Must use vbModeLess to prevent App from waiting for form to close
            .Show vbModeless
            .InitializeUMCSearch
            blnSuccess = .StartUMCSearch()
        End With
        Unload objUMCSearchForm
        Set objUMCSearchForm = Nothing
        
        ' Write the latest AnalysisHistory information to the log
        AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
        
        If GelUMC(udtWorkingParams.GelIndex).UMCCnt = 0 Then
            ' No LC-MS Features were found
            strMessage = "No LC-MS Features were found -- unable to continue auto analysis.  Perhaps the filter parameters were inappropriate for this dataset."
            AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
            If udtAutoParams.ShowMessages Then MsgBox strMessage, vbExclamation + vbOKOnly, "No LC-MS Features"
        Else
            If Not blnSuccess Then
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - LC-MS Feature search was aborted prematurely"
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or UMC_SEARCH_ABORTED_WARNING_BIT
            End If
        End If
    End If

    Exit Sub
    
FindUMCsErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while finding LC-MS Features during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

Private Sub AutoAnalysisFlushLogfile(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
    ' Flush the log file by closing, then re-opening it
    Dim strNewFilePath As String
    
On Error GoTo FlushLogFileErrorHandler

    If Not udtWorkingParams.ts Is Nothing Then
        udtWorkingParams.ts.Close
        Set udtWorkingParams.ts = Nothing
    End If
    
    ' Reopen the log file
ReopenLogFile:
On Error Resume Next
    With udtAutoParams.FilePaths
        If Len(.LogFilePath) > 0 Then
            If Not .LogFilePathError Then
                Set udtWorkingParams.ts = fso.OpenTextFile(.LogFilePath, ForAppending, True)
                
                If (udtWorkingParams.ts Is Nothing Or Err.Number <> 0) Then
                    ' Error re-opening the log file
                    ' Append .txt to the filename and try again
                    Debug.Assert False
                    Err.Clear
                    strNewFilePath = .LogFilePath & ".txt"
                    
                    Set udtWorkingParams.ts = fso.OpenTextFile(strNewFilePath, ForAppending, True)
                    
                    If udtWorkingParams.ts Is Nothing Then
                        If udtAutoParams.ShowMessages Then
                            MsgBox "Unable to initialize the log file (" & .LogFilePath & ").  Also tried " & strNewFilePath & ": " & Err.Description
                        End If
                        .LogFilePathError = True
                    Else
                        .LogFilePath = strNewFilePath
                    End If
                End If
                Err.Clear
            End If
        End If
    End With
    
    Exit Sub
    
FlushLogFileErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "AutoAnalysisFlushLogfile", "Error closing the log file: " & Err.Description
    Resume ReopenLogFile
    
End Sub

Private Function AutoAnalysisGenerateHTMLFolderLinkText(strSubFolderName As String, strDataFileName As String) As String
    
    Const strQ As String = """"
    Dim strOutLine As String

    strOutLine = "<TR><TD><A href=" & strQ & strSubFolderName & "/" & HTML_INDEX_FILE_NAME & strQ & ">" & strSubFolderName & "</a></TD>"
    
    If Len(strDataFileName) > 0 Then
        strOutLine = strOutLine & "<TD>" & strDataFileName & "</TD>"
    End If
    strOutLine = strOutLine & "</TR>"

    AutoAnalysisGenerateHTMLFolderLinkText = strOutLine
    
End Function

Private Function AutoAnalysisGenerateHTMLLookupDataFileName(strFolderPath As String, ByRef fso As FileSystemObject) As String

    Dim tsInFile As TextStream
    
    Dim strFilePath As String
    Dim strDataFileName As String
    
    Dim strLineIn As String
    Dim strCompareText As String
    Dim lngCharLoc As Long
    
    strCompareText = UCase(HTML_INDEX_FILE_DATA_FILE_LINE_START)
    
    strDataFileName = ""
    strFilePath = fso.BuildPath(strFolderPath, HTML_INDEX_FILE_NAME)
    
    If fso.FileExists(strFilePath) Then
        Set tsInFile = fso.OpenTextFile(strFilePath, ForReading)
        Do While Not tsInFile.AtEndOfStream
            strLineIn = tsInFile.ReadLine
            
            lngCharLoc = InStr(UCase(strLineIn), strCompareText)
            
            If lngCharLoc > 0 Then
                strDataFileName = Trim(Mid(strLineIn, lngCharLoc + 11))
                If Right(UCase(strDataFileName), 4) = "<BR>" Then
                    strDataFileName = Left(strDataFileName, Len(strDataFileName) - 4)
                End If
                
                Exit Do
            End If
        Loop
        tsInFile.Close
    End If
    
    AutoAnalysisGenerateHTMLLookupDataFileName = strDataFileName
    
End Function

Private Sub AutoAnalysisGenerateHTMLSubfolderListFile(objFolder As Folder, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)

    Const strQ As String = """"
    
    Dim objSubFolder As Folder
    Dim tsOutput As TextStream
    
    Dim strHtmlFilepath As String
    Dim strOutLine As String
    Dim strMessage As String
    
    Dim intSubFolderCount As Integer
    Dim strSubFolders() As String
    Dim strDataFileNames() As String
    
    Dim blnDataFileNamesPresent As Boolean
    
    Dim intIndex As Integer
    
On Error GoTo AutoAnalysisGenerateHTMLSubfolderListFileErrorHandler

    If Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
        
        blnDataFileNamesPresent = False
        
        ' Cache the subfolder names
        If objFolder.SubFolders.Count > 0 Then
            ReDim strSubFolders(objFolder.SubFolders.Count - 1)
            
            intSubFolderCount = 0
            For Each objSubFolder In objFolder.SubFolders
                strSubFolders(intSubFolderCount) = objSubFolder.Path
                intSubFolderCount = intSubFolderCount + 1
                DoEvents
            Next objSubFolder
            
            If intSubFolderCount > 0 Then
                ' Sort strSubFolders()
                ShellSortString strSubFolders, 0, intSubFolderCount - 1
                
                ReDim strDataFileNames(intSubFolderCount - 1)
                
                ' Lookup the datafile name for each Folder
                For intIndex = 0 To intSubFolderCount - 1
                    ' Look for an Index.html file in the subfolder
                    ' If it exists, open it and look for the line "Date file: "
                    ' If found, return the filename
                    strDataFileNames(intIndex) = AutoAnalysisGenerateHTMLLookupDataFileName(strSubFolders(intIndex), fso)
                    If Len(strDataFileNames(intIndex)) > 0 Then
                        blnDataFileNamesPresent = True
                    End If
                    DoEvents
                Next intIndex
            End If
        Else
            intSubFolderCount = 0
            ReDim strSubFolders(0)
            ReDim strDataFileNames(0)
        End If
        
        ' Create an Index.html file that lists the subfolders of this folder
        strHtmlFilepath = fso.BuildPath(objFolder.Path, HTML_INDEX_FILE_NAME)
        
        Set tsOutput = fso.CreateTextFile(strHtmlFilepath, True)
        
        tsOutput.WriteLine "<!DOCTYPE html PUBLIC " & strQ & "-//W3C//DTD HTML 4.0 Transitional//EN" & strQ & ">"
        tsOutput.WriteLine "<HTML>"
        tsOutput.WriteLine "  <HEAD>"
        tsOutput.WriteLine "    <META HTTP-EQUIV=" & strQ & "Content-Type" & strQ & " CONTENT=" & strQ & "text/html; charset=iso-8859-1" & strQ & ">"
        tsOutput.WriteLine "    <TITLE>"
        
        tsOutput.WriteLine "      " & objFolder.Name
        tsOutput.WriteLine "    </TITLE>"
        tsOutput.WriteLine "  </HEAD>"
        tsOutput.WriteLine "  <BODY>"
        tsOutput.WriteLine "    <P><FONT SIZE=+1>" & objFolder.Path & "</FONT></P>"
        tsOutput.WriteLine "    <P>"
        tsOutput.WriteLine "        <a href=" & strQ & "../Index.html" & strQ & ">Up one folder</a>"
        tsOutput.WriteLine "    </P>"
        
        If intSubFolderCount > 0 Then
            
            If blnDataFileNamesPresent Then
                tsOutput.WriteLine "<TABLE Border=1 Width=75%>"
                strOutLine = HTML_INDEX_FOLDER_LIST_TABLE_START & "<TH>Datafile</TH></TR>"
            Else
                tsOutput.WriteLine "<TABLE Border=1>"
                strOutLine = "<TR><TH>Folder</TH></TR>"
            End If
            tsOutput.WriteLine strOutLine
            
            ' Cache the subfolders
            For intIndex = 0 To intSubFolderCount - 1
                If KeyPressAbortProcess > 1 Then Exit For
                
                strOutLine = AutoAnalysisGenerateHTMLFolderLinkText(fso.GetFileName(strSubFolders(intIndex)), strDataFileNames(intIndex))
                
                tsOutput.WriteLine strOutLine
            Next intIndex
            tsOutput.WriteLine HTML_INDEX_FOLDER_LIST_TABLE_END
        Else
            tsOutput.WriteLine "<P>No subfolders are present.</P>"
        End If

        tsOutput.WriteLine "  </BODY>"
        tsOutput.WriteLine "</HTML>"
        tsOutput.Close
        
    End If

    Exit Sub
    
AutoAnalysisGenerateHTMLSubfolderListFileErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while creating the browsable Index.html of subfolders file during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

Private Sub AutoAnalysisGenerateHTMLBrowsingFile(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject, Optional ByVal strVersionOverride As String = "", Optional ByVal strDateOverride As String = "")

    Const strQ As String = """"
    
    Dim strHtmlFilepath As String
    Dim strMessage As String
    
    Dim strOutputLine As String
    Dim strRows() As String
    Dim strRightAlign As String
    
    Dim tsOutput As TextStream
    Dim objOutputFolder As Folder
    Dim objFile As File
    
    Dim strTempPath As String
    Dim strExtension As String
    
    Dim blnLookupStatsInDB As Boolean
    Dim lngLCMSFeatureCount As Long
    Dim lngLCMSFeatureCountWithHits As Long
    Dim lngUniqueMassTagCount As Long
    
    Dim intPicFilesCount As Integer
    Dim strPicFiles() As String
    
    Dim intTextFilesCount As Integer
    Dim strTextFiles() As String
    
    Dim intOtherFilesCount As Integer
    Dim strOtherFiles() As String
    
    Dim intIndex As Integer
    
    Dim intMaxRowIndex As Integer
    Dim intMaxColIndex As Integer
    
    Dim intRowNumber As Integer
    Dim intColNumber As Integer
    
On Error GoTo GenerateHTMLBrowsingFileErrorHandler

    If Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
        ' Create an Index.html file that can be used to browse the results
        
        strHtmlFilepath = fso.BuildPath(udtWorkingParams.GelOutputFolder, HTML_INDEX_FILE_NAME)
        
        Set tsOutput = fso.CreateTextFile(strHtmlFilepath, True)
        
        tsOutput.WriteLine "<!DOCTYPE html PUBLIC " & strQ & "-//W3C//DTD HTML 4.0 Transitional//EN" & strQ & ">"
        tsOutput.WriteLine "<HTML>"
        tsOutput.WriteLine "  <HEAD>"
        tsOutput.WriteLine "    <META HTTP-EQUIV=" & strQ & "Content-Type" & strQ & " CONTENT=" & strQ & "text/html; charset=iso-8859-1" & strQ & ">"
        tsOutput.WriteLine "    <TITLE>"
        
        If Len(udtAutoParams.FilePaths.OutputFolderPath) > 0 Then
            Debug.Assert udtWorkingParams.GelOutputFolder = udtAutoParams.FilePaths.OutputFolderPath
        End If
        
        tsOutput.WriteLine "      " & fso.GetFileName(udtWorkingParams.GelOutputFolder)
        tsOutput.WriteLine "    </TITLE>"
        tsOutput.WriteLine "  </HEAD>"
        tsOutput.WriteLine "  <BODY>"
        tsOutput.WriteLine "    <P>"
        If Len(udtWorkingParams.GelOutputFolder) > 0 Then
            tsOutput.WriteLine "        Report name: " & fso.GetFileName(udtWorkingParams.GelOutputFolder) & "<BR>"
        End If
        tsOutput.WriteLine "        " & HTML_INDEX_FILE_DATA_FILE_LINE_START & " " & fso.GetFileName(udtAutoParams.FilePaths.InputFilePath) & "<BR>"
        
        strTempPath = udtAutoParams.FilePaths.InputFilePath
        intIndex = InStr(strTempPath, fso.GetFileName(udtAutoParams.FilePaths.InputFilePath))
        If intIndex > 1 Then
            strTempPath = Left(strTempPath, intIndex - 1)
        End If
        
        tsOutput.WriteLine "        Source folder: <A HREF=" & strQ & strTempPath & strQ & ">" & strTempPath & "</a><BR>"
        tsOutput.WriteLine "    </P>"
        
        lngLCMSFeatureCount = -1
        lngLCMSFeatureCountWithHits = -1
        lngUniqueMassTagCount = -1
        
        blnLookupStatsInDB = False
        If udtAutoParams.MTDBOverride.Enabled And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            ' Note that we cannot use LookupMatchStatsForPeakMatchingTask during auto-analysis since the
            '  Peak matching table will not yet have the appropriate MD_ID value
            If udtAutoParams.MTDBOverride.PeakMatchingTaskID > 0 And Len(udtAutoParams.MTDBOverride.ServerName) > 0 And Len(udtAutoParams.MTDBOverride.MTDBName) > 0 Then
                blnLookupStatsInDB = True
            End If
        End If
        
        If APP_BUILD_DISABLE_MTS Then
            blnLookupStatsInDB = False
        End If
        
        If blnLookupStatsInDB Then
            If LookupMatchStatsForPeakMatchingTask(udtAutoParams.MTDBOverride.ServerName, _
                                                   udtAutoParams.MTDBOverride.MTDBName, _
                                                   udtAutoParams.MTDBOverride.PeakMatchingTaskID, _
                                                   0, _
                                                   lngLCMSFeatureCount, _
                                                   lngLCMSFeatureCountWithHits, _
                                                   lngUniqueMassTagCount) Then
                ' Success looking up stats from DB
            Else
                ' DB Lookup failed
                ' DB Search possibly was set to not export results to database
                blnLookupStatsInDB = False
            End If
        End If
        
        If Not blnLookupStatsInDB Then
            If udtWorkingParams.GelIndex > 0 Then
                ' Examine the data in memory to determine the match stats
                LookupMatchingUMCStats udtWorkingParams.GelIndex, lngLCMSFeatureCount, lngLCMSFeatureCountWithHits, lngUniqueMassTagCount
            Else
                ' Data is not in memory and task ID is undefined
                ' Cannot obtain these values from anywhere
            End If
        End If
        
        tsOutput.WriteLine "    <P>"
        tsOutput.WriteLine "        <TABLE>"
        
        ReDim strRows(2)
        
        strRightAlign = "<TD align=" & strQ & "right" & strQ & ">"
        strRows(0) = "<TD>Total LC-MS Feature count:" & strRightAlign & Trim(lngLCMSFeatureCount) & "<TD width=100>"
        strRows(1) = "<TD>Feature count with hits:" & strRightAlign & Trim(lngLCMSFeatureCountWithHits) & "<TD width=100>"
        strRows(2) = "<TD>Unique mass tag count:" & strRightAlign & Trim(lngUniqueMassTagCount) & "<TD width=100>"
      
        ' Write out the Mass Error and NET Error peak stats
        With glbPreferencesExpanded.AutoAnalysisCachedData.MassCalErrorPeakCached
            strRows(0) = strRows(0) & "<TD>Mass error peak center:" & strRightAlign & Trim(.Center) & "<TD width=100>ppm"
            strRows(1) = strRows(1) & "<TD>Mass error peak width:" & strRightAlign & Trim(.width) & "<TD width=100>ppm"
            strRows(2) = strRows(2) & "<TD>Mass error peak height:" & strRightAlign & Trim(.Height) & "<TD width=100>counts"
        End With
        
        With glbPreferencesExpanded.AutoAnalysisCachedData.NETTolErrorPeakCached
            strRows(0) = strRows(0) & "<TD>NET error peak center:" & strRightAlign & Trim(.Center) & "<TD>NET"
            strRows(1) = strRows(1) & "<TD>NET error peak width:" & strRightAlign & Trim(.width) & "<TD>NET"
            strRows(2) = strRows(2) & "<TD>NET error peak height:" & strRightAlign & Trim(.Height) & "<TD>counts"
        End With
        
        For intIndex = 0 To UBound(strRows)
            tsOutput.WriteLine "        <TR>" & strRows(intIndex) & "</TR>"
        Next intIndex
        
        tsOutput.WriteLine "        </TABLE>"
        

        tsOutput.WriteLine "    </P>"
        tsOutput.WriteLine ""
        
        ' Write the graphic files
        ' First determine the maximum number of rows and maximum number of columns
        With udtWorkingParams
            intMaxRowIndex = 0
            intMaxColIndex = 0
             
            For intIndex = 0 To .GraphicOutputFileInfoCount - 1
                With .GraphicOutputFileInfo(intIndex)
                    If .TableRow > intMaxRowIndex Then
                        intMaxRowIndex = .TableRow
                    End If
                
                    If .TableColumn > intMaxColIndex Then
                        intMaxColIndex = .TableColumn
                    End If
                End With
            Next intIndex
        End With
        
        If intMaxRowIndex > 0 Then
            tsOutput.WriteLine "  <TABLE>"
            With udtWorkingParams
                For intRowNumber = 1 To intMaxRowIndex
                    tsOutput.WriteLine "    <TR>"
                    
                    For intColNumber = 1 To intMaxColIndex
                        strOutputLine = "     <TD>"
                        
                        For intIndex = 0 To .GraphicOutputFileInfoCount - 1
                            With .GraphicOutputFileInfo(intIndex)
                                If .TableRow = intRowNumber And .TableColumn = intColNumber Then
                                    If Right(strOutputLine, 1) <> ">" Then strOutputLine = strOutputLine & "<BR>"
                                    strOutputLine = strOutputLine & "<A HREF=" & strQ & .FileName & strQ & ">"
                                    strOutputLine = strOutputLine & "<IMG Src=" & strQ & .FileName & strQ & " WIDTH=" & Trim(.width) & "></a><BR>"
                                    strOutputLine = strOutputLine & .Description
                                End If
                            End With
                        Next intIndex
                        
                        strOutputLine = strOutputLine & "</TD>"
                        tsOutput.WriteLine strOutputLine
                    Next intColNumber
                
                    tsOutput.WriteLine "    </TR>"
                Next intRowNumber
            End With
            
            tsOutput.WriteLine "  </TABLE>"
        End If
        
        tsOutput.WriteLine ""
        tsOutput.WriteLine "    <P>"
        tsOutput.WriteLine "      Windows folder link: <A href=" & strQ & udtWorkingParams.GelOutputFolder & strQ & ">" & udtWorkingParams.GelOutputFolder & "</a>"
        tsOutput.WriteLine "    </P>"
        
        ' Examine the output folder to generate a complete list of the graphics files and text files present there
        Set objOutputFolder = fso.GetFolder(udtWorkingParams.GelOutputFolder)
        
        intPicFilesCount = 0
        intTextFilesCount = 0
        intOtherFilesCount = 0
        ReDim strPicFiles(0)
        ReDim strTextFiles(0)
        ReDim strOtherFiles(0)
        
        For Each objFile In objOutputFolder.Files
            strExtension = LCase(Trim(fso.GetExtensionName(objFile.Name)))
            If strExtension = "jpg" Or strExtension = "png" Or strExtension = "wmf" Or strExtension = "emf" Or strExtension = "gif" Then
                ReDim Preserve strPicFiles(intPicFilesCount)
                strPicFiles(intPicFilesCount) = objFile.Name
                intPicFilesCount = intPicFilesCount + 1
            ElseIf strExtension = "txt" Then
                ReDim Preserve strTextFiles(intTextFilesCount)
                strTextFiles(intTextFilesCount) = objFile.Name
                intTextFilesCount = intTextFilesCount + 1
            Else
                If LCase(objFile.Name) <> LCase(HTML_INDEX_FILE_NAME) Then
                    ReDim Preserve strOtherFiles(intOtherFilesCount)
                    strOtherFiles(intOtherFilesCount) = objFile.Name
                    intOtherFilesCount = intOtherFilesCount + 1
                End If
            End If
        Next objFile

        intMaxRowIndex = intTextFilesCount
        If intPicFilesCount > intMaxRowIndex Then intMaxRowIndex = intPicFilesCount
        
        tsOutput.WriteLine "  <TABLE>"
        tsOutput.WriteLine "    <TR>"
        tsOutput.WriteLine "     <TH>Picture files</TH><TH>Text files</TH>"
        tsOutput.WriteLine "    </TR>"
        
        For intRowNumber = 1 To intMaxRowIndex
            tsOutput.WriteLine "    <TR>"
            
            strOutputLine = "     <TD>"
            If intRowNumber <= intPicFilesCount Then
                strOutputLine = strOutputLine & "<A HREF=" & strQ & strPicFiles(intRowNumber - 1) & strQ & ">" & strPicFiles(intRowNumber - 1) & "</a>"
            End If
            strOutputLine = strOutputLine & "</TD>"
            
            strOutputLine = strOutputLine & " <TD>"
            If intRowNumber <= intTextFilesCount Then
                strOutputLine = strOutputLine & "<A HREF=" & strQ & strTextFiles(intRowNumber - 1) & strQ & ">" & strTextFiles(intRowNumber - 1) & "</a>"
            End If
            strOutputLine = strOutputLine & "</TD>"
            
            tsOutput.WriteLine strOutputLine
        
            tsOutput.WriteLine "    </TR>"
        Next intRowNumber
        tsOutput.WriteLine "  </TABLE>"
        
        tsOutput.WriteLine "  <TABLE>"
        tsOutput.WriteLine "    <TR>"
        tsOutput.WriteLine "     <TH>Other files</TH>"
        tsOutput.WriteLine "    </TR>"
        
        For intRowNumber = 1 To intOtherFilesCount
            tsOutput.WriteLine "    <TR>"
            
            strOutputLine = "     <TD>"
            strOutputLine = strOutputLine & "<A HREF=" & strQ & strOtherFiles(intRowNumber - 1) & strQ & ">" & strOtherFiles(intRowNumber - 1) & "</a>"
            strOutputLine = strOutputLine & "</TD>"
            
            tsOutput.WriteLine strOutputLine
        
            tsOutput.WriteLine "    </TR>"
        Next intRowNumber
        tsOutput.WriteLine "  </TABLE>"

        tsOutput.WriteLine ""
        tsOutput.WriteLine "<P>"
        
        If Len(strVersionOverride) = 0 Then
            strVersionOverride = GetMyNameVersion(False, True)
        End If
        
        If Len(strDateOverride) = 0 Then
            strDateOverride = Format(Now(), "MMMM dd, yyyy at hh:nn AMPM")
        End If
        
        tsOutput.WriteLine "  Generated " & strDateOverride & " by Viper (v" & strVersionOverride & ")"
        tsOutput.WriteLine "</P>"

        tsOutput.WriteLine "  </BODY>"
        tsOutput.WriteLine "</HTML>"
        tsOutput.Close
        
    End If

    Exit Sub
    
GenerateHTMLBrowsingFileErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while creating the browsable Index.html file: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    
End Sub

' September 2004: Unused function
''Private Sub AutoAnalysisUpdateParentHTMLBrowsingFile(byref udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, byref fso As FileSystemObject)
''
''    Dim strIndexHtmlFilePath As String
''
''    Dim objFolder As Folder
''    Dim objParentFolder As Folder
''    Dim tsIndexHtmlFile As TextStream
''
''    Dim lngIndex As Long
''
''    Dim lngLineCount As Long
''    Dim lngLineCountDimmed As Long
''    Dim strLines() As String         ' Cache of entire file; loaded into memory for simplicity
''
''    Dim blnFolderListFound As Boolean
''    Dim blnLinkAdded As Boolean
''
''    Dim strMessage As String
''
''On Error GoTo AutoAnalysisUpdateParentHTMLBrowsingFileErrorHandler:
''
''    ' See if an Index.html file exists in the parent folder of the output folder
''
''    Set objFolder = fso.GetFolder(udtWorkingParams.GelOutputFolder)
''
''    strIndexHtmlFilePath = fso.BuildPath(objFolder.ParentFolder.Path, HTML_INDEX_FILE_NAME)
''
''    If fso.FileExists(strIndexHtmlFilePath) Then
''        ' File exists; open it and append an entry for the Index.html file in the output folder, along with the dataset name
''
''        Set tsIndexHtmlFile = fso.OpenTextFile(strIndexHtmlFilePath, ForReading, False)
''
''        lngLineCountDimmed = 1000
''        lngLineCount = 0
''        ReDim strLines(lngLineCountDimmed)
''
''        Do While Not tsIndexHtmlFile.AtEndOfStream
''            strLines(lngLineCount) = tsIndexHtmlFile.ReadLine
''            lngLineCount = lngLineCount + 1
''            If lngLineCount >= lngLineCountDimmed Then
''                lngLineCountDimmed = lngLineCountDimmed + 1000
''                ReDim Preserve strLines(lngLineCountDimmed)
''            End If
''
''            If Not blnFolderListFound Then
''                ' See if the line starts with "<TR><TH>Folder"
''                If UCase(Trim(Left(strLines(lngLineCount - 1), Len(HTML_INDEX_FOLDER_LIST_TABLE_START)))) = UCase(HTML_INDEX_FOLDER_LIST_TABLE_START) Then
''                    blnFolderListFound = True
''                End If
''            ElseIf Not blnLinkAdded Then
''                ' See if the line starts with "</TABLE>"
''                If UCase(Trim(Left(strLines(lngLineCount - 1), Len(HTML_INDEX_FOLDER_LIST_TABLE_END)))) = UCase(HTML_INDEX_FOLDER_LIST_TABLE_END) Then
''                    ' </TABLE> line found
''                    ' Insert the link for the new folder here
''
''                    AutoAnalysisUpdateParentHTMLBrowsingFileInsertLink strLines(), lngLineCount, lngLineCountDimmed, objFolder.Name, fso.GetFileName(udtAutoParams.FilePaths.InputFilePath)
''
''                    blnLinkAdded = True
''                End If
''
''            End If
''        Loop
''
''        tsIndexHtmlFile.Close
''
''        ' Wait 200 msec, just to be safe
''        Sleep 200
''
''        If Not blnLinkAdded Then
''            ' Need to add the link to the folder
''            ' Step backward through strLines() until the first non-blank line that isn't   </BODY> or </HTML>
''
''            lngIndex = lngLineCount - 1
''            Do While lngIndex > 0
''                If Len(strLines(lngIndex)) > 0 Then
''                    Select Case UCase(Trim(strLines(lngIndex)))
''                    Case "</BODY>"
''                    Case "</HTML>"
''                    Case Else
''                        ' Place the output file here
''                        AutoAnalysisUpdateParentHTMLBrowsingFileInsertLink strLines(), lngLineCount, lngLineCountDimmed, objFolder.Name, fso.GetFileName(udtAutoParams.FilePaths.InputFilePath)
''                        blnLinkAdded = True
''                        Exit Do
''                    End Select
''                End If
''                lngIndex = lngIndex - 1
''            Loop
''        End If
''
''        ' Now write the data out
''        If blnLinkAdded Then
''            Set tsIndexHtmlFile = fso.OpenTextFile(strIndexHtmlFilePath, ForWriting, True)
''            For lngIndex = 0 To lngLineCount - 1
''                tsIndexHtmlFile.WriteLine strLines(lngIndex)
''            Next lngIndex
''            tsIndexHtmlFile.Close
''        End If
''    Else
''        ' File doesn't exist, use GenerateAutoAnalysisHtmlFiles to create it
''        ' We need to determine the appropriate start folder (possibly up more than one folder from objFolder.ParentFolder)
''
''        Set objParentFolder = fso.GetFolder(objFolder.Path)
''
''        On Error Resume Next
''        Do
''            If Len(FixNull(objParentFolder.ParentFolder.Path)) = 0 Then
''                Exit Do
''            Else
''                If fso.FolderExists(objParentFolder.ParentFolder.Path) Then
''                    Set objParentFolder = fso.GetFolder(objParentFolder.ParentFolder.Path)
''                Else
''                    Exit Do
''                End If
''            End If
''        Loop
''
''On Error GoTo AutoAnalysisUpdateParentHTMLBrowsingFileErrorHandler:
''
''        ' Call GenerateAutoAnalysisHtmlFiles, sending the start folder path and a folder path target mask
''        GenerateAutoAnalysisHtmlFiles objParentFolder.Path, objFolder.Path, False, 0, 0, False
''
''    End If
''
''
''    Exit Sub
''
''AutoAnalysisUpdateParentHTMLBrowsingFileErrorHandler:
''    Debug.Assert False
''    strMessage = "Error - An error has occurred while creating/updating the browsable Index.html file for the parent folder of the output folder during auto analysis: " & Err.Description
''    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
''        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
''    Else
''        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
''    End If
''
''End Sub
''
''Private Sub AutoAnalysisUpdateParentHTMLBrowsingFileInsertLink(ByRef strLines() As String, ByRef lngLineCount As Long, ByRef lngLineCountDimmed As Long, ByVal strFolderName As String, ByVal strDataFileName As String)
''
''    ' First, shift the </TABLE> line by one
''    strLines(lngLineCount) = strLines(lngLineCount - 1)
''    lngLineCount = lngLineCount + 1
''    If lngLineCount >= lngLineCountDimmed Then
''        lngLineCountDimmed = lngLineCountDimmed + 1000
''        ReDim Preserve strLines(lngLineCountDimmed)
''    End If
''
''    ' Now store the link to the new folder in strlines(lngLineCount-2)
''    strLines(lngLineCount - 2) = AutoAnalysisGenerateHTMLFolderLinkText(strFolderName, strDataFileName)
''
''End Sub

Private Sub AutoAnalysisInitializeCachedData(ByRef udtAutoAnalysisCachedData As udtAutoAnalysisCachedDataType)
    
    With udtAutoAnalysisCachedData
        With .MassCalErrorPeakCached
            .Center = 0
            .width = 0
            .Height = 0
            .SingleValidPeak = False
        End With
    
        With .MassCalErrorPeakCached
            .Center = 0
            .width = 0
            .Height = 0
            .SingleValidPeak = False
        End With
    
        .Initialized = True
    End With
    
End Sub

Private Function AutoAnalysisInitializeLogging(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject, blnClearMemoryLog As Boolean, ByRef blnLogFileIsInTempFolder As Boolean) As Boolean
    ' Returns True if success, False otherwise

    Dim strMessage As String

On Error GoTo InitializeLoggingErrorHandler
    
    With glbPreferencesExpanded.AutoAnalysisStatus
        .AutoAnalysisTimeStamp = Format(Now(), "yyyy.mm.dd Hh:Nn:Ss")
        .Enabled = True
    End With
    
    mRedirectedOutputFolderMessage = ""
    If blnClearMemoryLog Then AutoAnalysisMemoryLogClear
    
    ' If udtAutoParams.FilePaths.LogFilePath = "", then start a new log in the temporary folder
    ' If .LogFilePath isn't empty, then see if it contains a folder path rather than a file path
    On Error Resume Next
    With udtAutoParams.FilePaths
        If Len(.LogFilePath) = 0 Then
            .LogFilePath = GetTempFolder & "VIPER_AutoAnalysisLogFile.Log"
            If fso.FileExists(.LogFilePath) Then
                fso.DeleteFile .LogFilePath, True
            End If
            blnLogFileIsInTempFolder = True
        ElseIf FolderExists(.LogFilePath) Then
            ' .LogFilePath holds a folder path, rather than a path to a specific file
            .LogFilePath = fso.BuildPath(.LogFilePath, "VIPER_AutoAnalysisLogFile.Log")
        End If
        
        CreateFolderByPath fso.GetParentFolderName(.LogFilePath)
        
        ' Initialize the Log file (open for appending)
        On Error Resume Next
        .LogFilePathError = False
        Set udtWorkingParams.ts = fso.OpenTextFile(.LogFilePath, ForAppending, True)
        
        If (udtWorkingParams.ts Is Nothing Or Err.Number <> 0) Then
            If udtAutoParams.ShowMessages Then
                MsgBox "Unable to initialize the log file (" & .LogFilePath & "): " & Err.Description
            End If
            .LogFilePathError = True
        End If
        Err.Clear
        
        strMessage = "VIPER Analysis Started"
        If Len(udtAutoParams.ComputerName) > 0 Then
            strMessage = strMessage & " on " & udtAutoParams.ComputerName
        End If
        
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage & " (version " & GetMyNameVersion(False, True) & ")"
        
        udtWorkingParams.NextHistoryIndexToCopy = 0
    End With
 
    AutoAnalysisInitializeLogging = True
    Exit Function

InitializeLoggingErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while initializing the memory and/or disk log during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    AutoAnalysisInitializeLogging = False

End Function

Private Sub AutoAnalysisInitializeWorkingParams(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType)
    With udtWorkingParams
        .GelIndex = 0
        .ErrorBits = 0
        .WarningBits = 0
        .ResultsFileNameBase = ""
        .GelOutputFolder = ""
        .GelFilePath = ""
        .GraphicOutputFileInfoCount = 0
        ReDim .GraphicOutputFileInfo(0)
        .TICPlotsStartRow = 4
        .NextHistoryIndexToCopy = 0
        .LoadedGelFile = False
        .NETDefined = False
    End With
End Sub

Private Function AutoAnalysisLoadDB(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject) As Boolean
    ' Returns True if the database is loaded, False otherwise
    
    Dim blnDBLoaded As Boolean
    Dim intConnectAttemptCount As Integer
    Dim strDBName As String
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean
    Dim strMessage As String
    
On Error GoTo LoadDBErrorHandler

    If GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
        blnDBLoaded = False
        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - Could not connect to the database -- the connection string is not defined"
    Else
        ' Note that intConnectAttemptCount is sent ByRef so that we can record the number of connection attempts required
        blnDBLoaded = ConfirmMassTagsAndInternalStdsLoaded(MDIForm1, udtWorkingParams.GelIndex, udtAutoParams.ShowMessages, intConnectAttemptCount, False, True, blnAMTsWereLoaded, blnDBConnectionError)
        If AMTCnt <= 0 Then blnDBLoaded = False
        strDBName = ExtractDBNameFromConnectionString(GelAnalysis(udtWorkingParams.GelIndex).MTDB.cn.ConnectionString)
        If blnDBLoaded Then
            If intConnectAttemptCount > 1 Then
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Note - Required " & intConnectAttemptCount & " attempts to connect to the database (" & strDBName & ")"
            End If
        
            With glbPreferencesExpanded.MassTagStalenessOptions
                If .AMTCountWithNulls / CDbl(.AMTCountInDB) >= .MaximumFractionAMTsWithNulls Then
                    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - The percentage of MT tags in the database (" & strDBName & ") with null mass or null NET values is abnormally high: " & Trim(Round(.AMTCountWithNulls / CDbl(.AMTCountInDB) * 100#, 2)) & "% of " & Trim(.AMTCountInDB) & " MT tags"
                    udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or MASS_TAGS_NULL_COUNTS_HIGH_ERROR_BIT
                ElseIf .AMTCountWithNulls >= .MaximumCountAMTsWithNulls Then
                    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - The number of MT tags in the database (" & strDBName & ") with null mass or null NET values is abnormally high: " & Trim(.AMTCountWithNulls) & " of " & Trim(.AMTCountInDB) & " MT tags"
                    udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or MASS_TAGS_NULL_COUNTS_HIGH_ERROR_BIT
                End If
            End With
        Else
            If Len(GelData(udtWorkingParams.GelIndex).PathtoDatabase) = 0 And Len(glbPreferencesExpanded.LegacyAMTDBPath) > 0 Then
                GelData(udtWorkingParams.GelIndex).PathtoDatabase = glbPreferencesExpanded.LegacyAMTDBPath
            End If
            
            If Len(GelData(udtWorkingParams.GelIndex).PathtoDatabase) > 0 Then
                strDBName = fso.GetFileName(GelData(udtWorkingParams.GelIndex).PathtoDatabase)
                blnDBLoaded = ConnectToLegacyAMTDB(MDIForm1, udtWorkingParams.GelIndex, False, True, False)
            End If
            If Not blnDBLoaded Then
                If blnDBConnectionError Then
                    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - Could not connect to the database (" & strDBName & "); tried " & intConnectAttemptCount & " times."
                Else
                    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - No valid MT tags were found (" & strDBName & "); possibly due to missing NET values"
                End If
            End If
        End If
    End If

    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
    
    AutoAnalysisLoadDB = blnDBLoaded
    Exit Function

LoadDBErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while loading the database during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    AutoAnalysisLoadDB = False

End Function

Private Function AutoAnalysisLoadOptions(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject) As Boolean
    ' Returns True if the database is defined and MT tags are ready to be loaded
    ' Returns False otherwise

    Dim udtAnalysisInfo As udtGelAnalysisInfoType
    
    Dim strDBInIniFile As String, strDBSelectedByUser As String
    Dim strParentFolderName As String
    Dim strMessage As String
    
    Dim intAutoSearchIndex As Integer
    Dim lngCharLoc As Long
    
    Dim dblSlope As Double
    Dim dblIntercept As Double
    
    Dim blnDBReadyToLoad As Boolean
    Dim eResponse As VbMsgBoxResult
    
On Error GoTo LoadOptionsErrorHandler

    ' Before loading the .Ini file, need to save a few parameters from GelAnalysis(udtWorkingParams.GelIndex)
    '  if this Sub was called from MDIForm1.MyAnalysisInit_DialogClosed()
    If udtAutoParams.AutoDMSAnalysisManuallyInitiated Then
        If GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
            ' This is unexpected
            Debug.Assert False
            udtAutoParams.AutoDMSAnalysisManuallyInitiated = False
        Else
            FillGelAnalysisInfo udtAnalysisInfo, GelAnalysis(udtWorkingParams.GelIndex)
        End If
    End If
    
    If Len(udtAutoParams.FilePaths.IniFilePath) > 0 Then
        ' IniFileLoadSettings requires a full path
        ' Thus, if udtAutoParams.FilePaths.IniFilePath does not contain "\" anywhere, then prepend with App.Path
        If InStr(udtAutoParams.FilePaths.IniFilePath, "\") = 0 Then
            udtAutoParams.FilePaths.IniFilePath = fso.BuildPath(App.Path, udtAutoParams.FilePaths.IniFilePath)
        End If
    
        If fso.FileExists(udtAutoParams.FilePaths.IniFilePath) Then
            
            SetDefaultDefinitions
            IniFileLoadSettings glbPreferencesExpanded, UMCDef, UMCIonNetDef, UMCNetAdjDef, UMCInternalStandards, samtDef, glPreferences, udtAutoParams.FilePaths.IniFilePath, True
            
            ' UMCDef and samtDef got updated during the load
            ' Need to set the Scope for each to "Current View" in case filters are enabled
            ' This will only cause a problem if the user zooms in while auto-analysis is occurring
            UMCDef.DefScope = glScope.glSc_Current
            samtDef.SearchScope = glScope.glSc_Current
            
            If UMCDef.MWField = 0 Then
                Debug.Assert False
                UMCDef.MWField = mftMWMono
            End If
            
            If samtDef.MWField = 0 Then
                Debug.Assert False
                UMCDef.MWField = mftMWMono
            End If
            
            ' Update GelSearchDef accordingly
            With GelSearchDef(udtWorkingParams.GelIndex)
                .UMCDef = UMCDef
                .UMCIonNetDef = UMCIonNetDef
                .AMTSearchOnIons = samtDef
                .AMTSearchOnUMCs = samtDef
                .AMTSearchOnPairs = samtDef
            End With
            
            GelP_D_L(udtWorkingParams.GelIndex).SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
            
            ' In addition, update GelUMCNETAdjDef
            GelUMCNETAdjDef(udtWorkingParams.GelIndex) = UMCNetAdjDef
            
            ' Update .LastAutoAnalysisIniFilePath
            glbPreferencesExpanded.LastAutoAnalysisIniFilePath = udtAutoParams.FilePaths.IniFilePath
            
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "AutoAnalysis options loaded from .Ini file: " & udtAutoParams.FilePaths.IniFilePath
        Else
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - The specified options .Ini file does not exist: " & udtAutoParams.FilePaths.IniFilePath
            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or INIFILE_LOAD_ERROR_BIT
        End If
    Else
        UMCDef.DefScope = glScope.glSc_Current
        samtDef.SearchScope = glScope.glSc_Current
    End If
    
    ' Now that the options have been loaded, fill the file path variables
    ' If a folder is present in udtAutoParams.FilePaths.OutputFolderPath, then this
    '  overrides all .AlternateOutputFolderPath values
    If Len(udtAutoParams.FilePaths.OutputFolderPath) > 0 Then
        For intAutoSearchIndex = 0 To MAX_AUTO_SEARCH_MODE_COUNT - 1
            glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).AlternateOutputFolderPath = udtAutoParams.FilePaths.OutputFolderPath
        Next intAutoSearchIndex
    End If
    
    ' Next attempt to initialize GelAnalysis(udtWorkingParams.GelIndex)
    If GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
        Set GelAnalysis(udtWorkingParams.GelIndex) = New FTICRAnalysis
        dblSlope = 0
        dblIntercept = 0
        
        ClearGelAnalysisObject udtWorkingParams.GelIndex, False
    Else
        If udtWorkingParams.LoadedGelFile Then
            dblSlope = GelAnalysis(udtWorkingParams.GelIndex).GANET_Slope
            dblIntercept = GelAnalysis(udtWorkingParams.GelIndex).GANET_Intercept
        End If
    End If
    
    If glbPreferencesExpanded.AutoAnalysisDBInfoIsValid Then
        FillGelAnalysisObject GelAnalysis(udtWorkingParams.GelIndex), glbPreferencesExpanded.AutoAnalysisDBInfo
    End If
    
    If udtWorkingParams.LoadedGelFile And dblSlope <> 0 Then
        GelAnalysis(udtWorkingParams.GelIndex).GANET_Slope = dblSlope
        GelAnalysis(udtWorkingParams.GelIndex).GANET_Intercept = dblIntercept
    End If
    
    If udtAutoParams.MTDBOverride.Enabled Then
        GelAnalysis(udtWorkingParams.GelIndex).MTDB.cn.ConnectionString = udtAutoParams.MTDBOverride.ConnectionString
    End If
    
    If Len(GelAnalysis(udtWorkingParams.GelIndex).MTDB.cn.ConnectionString) = 0 Or APP_BUILD_DISABLE_MTS Then
        If Len(GelData(udtWorkingParams.GelIndex).PathtoDatabase) = 0 And Len(glbPreferencesExpanded.LegacyAMTDBPath) > 0 Then
            GelData(udtWorkingParams.GelIndex).PathtoDatabase = glbPreferencesExpanded.LegacyAMTDBPath
        End If
        
        If Len(glbPreferencesExpanded.LegacyAMTDBPath) > 0 Then
            blnDBReadyToLoad = True
        Else
            blnDBReadyToLoad = False
        End If
    Else
        blnDBReadyToLoad = True
    End If

    If APP_BUILD_DISABLE_MTS Then
        udtAutoParams.AutoDMSAnalysisManuallyInitiated = False
        udtAutoParams.MTDBOverride.Enabled = False
    End If
    
    If udtAutoParams.AutoDMSAnalysisManuallyInitiated Then
        ' Update GelAnalysis() with the settings in udtAnalysisInfo
        ' However, do not update .MTDB or the DBStuff() collection since we want the settings in the
        '  .Ini file to take precedence
        FillGelAnalysisObject GelAnalysis(udtWorkingParams.GelIndex), udtAnalysisInfo, False, False
        
        ' If .cn.ConnectionString values do not agree between that in the .Ini file and that in udtAnalysisInfo, then
        '  query the user about which to use
        With GelAnalysis(udtWorkingParams.GelIndex).MTDB.cn
            strDBInIniFile = ExtractDBNameFromConnectionString(.ConnectionString)
            strDBSelectedByUser = ExtractDBNameFromConnectionString(udtAnalysisInfo.MTDB.ConnectionString)
            
            If strDBInIniFile <> strDBSelectedByUser Then
                If strDBInIniFile = "Unknown" Or Not udtAutoParams.ShowMessages Then
                    eResponse = vbNo
                Else
                    eResponse = MsgBox("The database in the .Ini file (" & strDBInIniFile & ") is not the same as the database you chose in the New Analysis dialog (" & strDBSelectedByUser & ").  Use the database in the .Ini file?", vbQuestion + vbYesNo + vbDefaultButton2, "Database conflict")
                End If
                
                If eResponse = vbNo Then
                    .ConnectionString = udtAnalysisInfo.MTDB.ConnectionString
                End If
            End If
        
            strDBSelectedByUser = ExtractDBNameFromConnectionString(.ConnectionString)
            If Len(strDBSelectedByUser) > 0 And strDBSelectedByUser <> "Unknown" Then
                blnDBReadyToLoad = True
            End If
        
        End With
    ElseIf udtAutoParams.MTDBOverride.Enabled Then
        ' PRISM Automation Mode is Enabled
        If Not GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
            With GelAnalysis(udtWorkingParams.GelIndex)
                With .MTDB
                    .DBStuff(NAME_SUBSET).Value = udtAutoParams.MTDBOverride.MTSubsetID
                    .DBStuff(NAME_INC_LIST).Value = udtAutoParams.MTDBOverride.ModList
                    .DBStuff(NAME_CONFIRMED_ONLY).Value = udtAutoParams.MTDBOverride.ConfirmedOnly
                    .DBStuff(NAME_ACCURATE_ONLY).Value = udtAutoParams.MTDBOverride.AMTsOnly
                    .DBStuff(NAME_LOCKERS_ONLY).Value = udtAutoParams.MTDBOverride.LockersOnly
                    .DBStuff(NAME_LIMIT_TO_PMTS_FROM_DATASET).Value = udtAutoParams.MTDBOverride.LimitToPMTsFromDataset
                    
                    .DBStuff(NAME_MINIMUM_HIGH_NORMALIZED_SCORE).Value = udtAutoParams.MTDBOverride.MinimumHighNormalizedScore
                    .DBStuff(NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE).Value = udtAutoParams.MTDBOverride.MinimumHighDiscriminantScore
                    .DBStuff(NAME_MINIMUM_PEPTIDE_PROPHET_PROBABILITY).Value = udtAutoParams.MTDBOverride.MinimumPeptideProphetProbability
                    .DBStuff(NAME_MINIMUM_PMT_QUALITY_SCORE).Value = udtAutoParams.MTDBOverride.MinimumPMTQualityScore
                    
                    .DBStuff(NAME_EXPERIMENT_INCLUSION_FILTER).Value = udtAutoParams.MTDBOverride.ExperimentInclusionFilter
                    .DBStuff(NAME_EXPERIMENT_EXCLUSION_FILTER).Value = udtAutoParams.MTDBOverride.ExperimentExclusionFilter
                    .DBStuff(NAME_INTERNAL_STANDARD_EXPLICIT).Value = udtAutoParams.MTDBOverride.InternalStandardExplicit
                    
                    .DBStuff(NAME_NET_VALUE_TYPE).Value = udtAutoParams.MTDBOverride.NETValueType
                    .cn.ConnectionString = udtAutoParams.MTDBOverride.ConnectionString
                End With
                
                If Len(.Dataset_Folder) = 0 Then .Dataset_Folder = udtAutoParams.FilePaths.DatasetFolder
                If Len(.Dataset) = 0 Then .Dataset = udtAutoParams.FilePaths.DatasetFolder
            End With
            
            blnDBReadyToLoad = True
        Else
            ' This probably shouldn't happen
            Debug.Assert False
        End If
    End If
    
    If Not udtAutoParams.MTDBOverride.Enabled Then
        If Len(GelAnalysis(udtWorkingParams.GelIndex).MTDB.DBStuff(NAME_INC_LIST).Value) = 0 Then
            ' It is important that Name_Inc_List not be blank for DB Schema Version < 2
            ' Change to -1
            ' This defines the modifications to consider; for no mods, use "Dynamic 1 and Static 1" in DB Schema version 1 and "Not Any" in DB Schema version 2
            GelAnalysis(udtWorkingParams.GelIndex).MTDB.DBStuff(NAME_INC_LIST).Value = "-1"
        End If
    End If
    
    If blnDBReadyToLoad Then
        If Not GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
            With GelAnalysis(udtWorkingParams.GelIndex)
                If udtAutoParams.JobNumber >= 0 Then
                    .Dataset_ID = udtAutoParams.DatasetID
                    .Job = udtAutoParams.JobNumber
                    .MD_Reference_Job = udtAutoParams.JobNumber
                Else
                    If Not udtAutoParams.AutoDMSAnalysisManuallyInitiated Then
                        .Dataset_ID = -1
                        .Job = -1
                        .MD_Reference_Job = -1
                        .MD_Type = glbPreferencesExpanded.AutoAnalysisOptions.MDType
                    End If
                End If
            End With
        End If
    Else
        If udtAutoParams.ShowMessages Then
            ' Show the DB Connections window to let the user define the database connection
            frmOrganizeDBConnections.Tag = udtWorkingParams.GelIndex
            frmOrganizeDBConnections.InitializeForm
            frmOrganizeDBConnections.Show vbModeless, MDIForm1
            
            frmOrganizeDBConnections.WaitUntilFormClose
            
            ' Assume user actually chose a database (we'll check for this in a second)
            blnDBReadyToLoad = True
        End If
    End If
    
    If Not udtAutoParams.AutoDMSAnalysisManuallyInitiated And Not GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
        ' Need to populate GelAnalysis(udtWorkingParams.GelIndex) with the needed parameters
        With GelAnalysis(udtWorkingParams.GelIndex)
            If .MD_Reference_Job < 0 Then
                ' Try to determine the job number from udtAutoParams.InputFilePath
                ' This will only work if udtAutoParams.InputFilePath is on one of the main servers (like Gigasax or Proto-1)
                strParentFolderName = fso.GetParentFolderName(udtAutoParams.FilePaths.InputFilePath)
                
                ' The above returned the full path to the parent folder
                ' We just want the name
                strParentFolderName = fso.GetFileName(strParentFolderName)
                
                lngCharLoc = InStr(UCase(strParentFolderName), "_AUTO")
                If lngCharLoc > 0 Then
                    strParentFolderName = Mid(strParentFolderName, lngCharLoc + Len("_AUTO"))
                    If IsNumeric(strParentFolderName) Then
                        .MD_Reference_Job = CLng(strParentFolderName)
                        .Job = .MD_Reference_Job
                    End If
                End If
            End If

            .MD_file = StripFullPath(udtAutoParams.FilePaths.InputFilePath)
            .MD_Type = glbPreferencesExpanded.AutoAnalysisOptions.MDType
        End With
    End If

    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    AutoAnalysisLoadOptions = blnDBReadyToLoad
    
    Exit Function
    
LoadOptionsErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while loading/setting options during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    AutoAnalysisLoadOptions = False
    
End Function

Private Function AutoAnalysisLoadInputFile(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject) As Boolean
    ' Returns True if success, False if an error
    
    Const MAX_WILDCARD_MATCHES = 1000
    Const MAX_FILE_EXTENSIONS_PREF_LIST_COUNT = 50
    
    Dim strSuggestedIniFileName As String, strStrippedPathForIniFile As String

    Dim strWildcardFileMatch As String
    Dim strWildcardFileMatches(MAX_WILDCARD_MATCHES) As String      ' 0-based array
    Dim intWildcardFileMatchesCount As Integer
    Dim intFileExtensionsPrefListCount As Integer, intExtensionLength As Integer
    Dim strFileExtensionsPrefList(MAX_FILE_EXTENSIONS_PREF_LIST_COUNT) As String        ' 0-based array
    Dim strSubString As String
    
    Dim strKeyValue As String
    Dim strDefault As String
    Dim strMessage As String
    Dim strParsedPath As String
    Dim strParentFolderPath As String
    
    Dim intCharLoc As Integer
    Dim strCaption As String
    
    Dim eFileType As ifmInputFileModeConstants
    
    Dim blnSuccess As Boolean, blnInvalidFolder As Boolean
    
    Dim strHistoryMatch As String
    Dim lngHistoryIndexLastMatch As Long
    
    Dim intExtensionIndex As Integer, intWildCardFileIndex As Integer

    Dim intAutoAnalysisSearchModeCount As Integer, intIndex As Integer
    
    Dim eResponse As VbMsgBoxResult
    Dim strErrorMessage As String
    Dim fsInputFile As File
    
On Error GoTo LoadInputFileErrorHandler

    If udtAutoParams.ShowMessages And Len(udtAutoParams.FilePaths.IniFilePath) = 0 Then
        ' Prompt for a custom .Ini file
        If Len(glbPreferencesExpanded.LastAutoAnalysisIniFilePath) > 0 Then
            strSuggestedIniFileName = StripFullPath(glbPreferencesExpanded.LastAutoAnalysisIniFilePath, strStrippedPathForIniFile)
        Else
            strSuggestedIniFileName = ""
            strStrippedPathForIniFile = ""
        End If
        
        
        strParsedPath = udtAutoParams.FilePaths.InputFilePath
        If Left(strParsedPath, 2) = "\\" Then
            ' The file is probably on the MTS Storage Servers; look up the folder name
            strParsedPath = Mid(strParsedPath, 3)
            intCharLoc = InStr(strParsedPath, "\")
            If intCharLoc > 0 Then
                strParsedPath = Mid(strParsedPath, intCharLoc + 1)
                
                intCharLoc = InStr(strParsedPath, "\")
                If intCharLoc > 0 Then
                    strParsedPath = Left(strParsedPath, intCharLoc - 1)
                Else
                strParsedPath = ""
                End If
            Else
                strParsedPath = ""
            End If
        Else
            strParsedPath = ""
        End If
        
        If Len(strParsedPath) > 0 Then
            strCaption = "Select .Ini file; root folder is " & strParsedPath
        Else
            strCaption = "Please select .Ini file to use for automated analysis"
        End If
        strCaption = strCaption & "; Input file is " & fso.GetBaseName(udtAutoParams.FilePaths.InputFilePath)
        
        udtAutoParams.FilePaths.IniFilePath = SelectFile(MDIForm1.hwnd, strCaption, strStrippedPathForIniFile, False, strSuggestedIniFileName, "Ini Files (*.ini)|*.ini", 1, True)
    
        If Len(udtAutoParams.FilePaths.IniFilePath) = 0 Then
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - No .Ini file was chosen"
            AutoAnalysisLoadInputFile = False
            Exit Function
        End If
    End If

    ' See if udtAutoParams.FilePaths.InputFilePath exists
    If Len(udtAutoParams.FilePaths.InputFilePath) > 0 Then
        blnSuccess = fso.FileExists(udtAutoParams.FilePaths.InputFilePath)
        If Not blnSuccess And InStr(udtAutoParams.FilePaths.InputFilePath, "*") Then
            ' File not found, however a wildcard is present in udtAutoParams.FilePaths.InputFilePath
            ' Find all files that match the wildcard
            ' If more than one match, choose the best one based on .PEKFileExtensionPreferenceOrder
            
            ' Need to pre-read the .Ini file to look up the PEKFileExtensionPreferenceOrder setting
            ' See below for additional comments about pre-reading the .Ini file
            If Len(udtAutoParams.FilePaths.IniFilePath) > 0 Then
                ' IniFileReadSingleSetting requires a full path
                ' Thus, if udtAutoParams.FilePaths.IniFilePath does not contain "\" anywhere, then prepend with App.Path
                If InStr(udtAutoParams.FilePaths.IniFilePath, "\") = 0 Then
                    udtAutoParams.FilePaths.IniFilePath = fso.BuildPath(App.Path, udtAutoParams.FilePaths.IniFilePath)
                End If
            
                If FileExists(udtAutoParams.FilePaths.IniFilePath) Then
                    With glbPreferencesExpanded.AutoAnalysisOptions
                        strKeyValue = IniFileReadSingleSetting("AutoAnalysisOptions", "PEKFileExtensionPreferenceOrder", .PEKFileExtensionPreferenceOrder, udtAutoParams.FilePaths.IniFilePath)
                        .PEKFileExtensionPreferenceOrder = strKeyValue
                    End With
                End If
            End If
            
            
            ' First make sure the folder exists
            strParentFolderPath = fso.GetParentFolderName(udtAutoParams.FilePaths.InputFilePath)
            If Len(strParentFolderPath) > 0 And Not fso.FolderExists(strParentFolderPath) Then
                blnInvalidFolder = True
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - File not found (invalid folder path): " & udtAutoParams.FilePaths.InputFilePath
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or DATAFILE_LOAD_ERROR_BIT
            Else
                intWildcardFileMatchesCount = 0
                strWildcardFileMatch = Dir(udtAutoParams.FilePaths.InputFilePath)
                Do While Len(strWildcardFileMatch) > 0 And intWildcardFileMatchesCount < MAX_WILDCARD_MATCHES
                    strWildcardFileMatches(intWildcardFileMatchesCount) = strWildcardFileMatch
                    intWildcardFileMatchesCount = intWildcardFileMatchesCount + 1
                    strWildcardFileMatch = Dir()
                Loop
                
                If intWildcardFileMatchesCount = 1 Then
                    strWildcardFileMatch = strWildcardFileMatches(0)
                ElseIf intWildcardFileMatchesCount > 1 Then
                    ' Find the best matching file
                    ' First parse PEKFileExtensionPreferenceOrder
                    intFileExtensionsPrefListCount = ParseString(glbPreferencesExpanded.AutoAnalysisOptions.PEKFileExtensionPreferenceOrder, strFileExtensionsPrefList(), MAX_FILE_EXTENSIONS_PREF_LIST_COUNT, ",", "", True, False, False)
                    
                    If intFileExtensionsPrefListCount = 0 Then
                        Debug.Assert False
                        intFileExtensionsPrefListCount = 7
                        strFileExtensionsPrefList(0) = CSV_ISOS_IC_FILE_SUFFIX
                        strFileExtensionsPrefList(1) = CSV_ISOS_FILE_SUFFIX
                        strFileExtensionsPrefList(2) = ".mzxml"
                        strFileExtensionsPrefList(3) = ".mzdata"
                        strFileExtensionsPrefList(4) = "mzxml.xml"
                        strFileExtensionsPrefList(5) = "mzdata.xml"
                        strFileExtensionsPrefList(6) = ".pek"
                    End If
                    
                    ' Now step through strFileExtensionsPrefList() and see if any of the files in strWildcardFileMatches() match
                    strWildcardFileMatch = ""
                    For intExtensionIndex = 0 To intFileExtensionsPrefListCount - 1
                        strFileExtensionsPrefList(intExtensionIndex) = LCase(Trim(strFileExtensionsPrefList(intExtensionIndex)))
                        intExtensionLength = Len(strFileExtensionsPrefList(intExtensionIndex))
                        If intExtensionLength > 0 Then
                        
                            If strFileExtensionsPrefList(intExtensionIndex) = ".pek" Or strFileExtensionsPrefList(intExtensionIndex) = ".csv" Then
                                ' Preferentially choose the .pek. or .csv file over the _*.??? file (like _ic.pek or _s.pek or _ic.csv or _s.csv)
                                
                                ' This takes some extra logic
                                For intWildCardFileIndex = 0 To intWildcardFileMatchesCount - 1
                                    If LCase(Right(strWildcardFileMatches(intWildCardFileIndex), intExtensionLength)) = strFileExtensionsPrefList(intExtensionIndex) Then
                                        ' Match found; make sure it doesn't have the form _x.??? or _xx.??? where x is any character
                                        
                                        strSubString = LCase(Right(strWildcardFileMatches(intWildCardFileIndex), 7))
                                        If Left(strSubString, 1) = "_" Then
                                            ' Skip this file
                                        Else
                                            strSubString = LCase(Right(strWildcardFileMatches(intWildCardFileIndex), 6))
                                            If Left(strSubString, 1) = "_" Then
                                                ' Skip this file
                                            Else
                                                strWildcardFileMatch = strWildcardFileMatches(intWildCardFileIndex)
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next intWildCardFileIndex
                            End If
                            
                            If Len(strWildcardFileMatch) = 0 Then
                                For intWildCardFileIndex = 0 To intWildcardFileMatchesCount - 1
                                    If LCase(Right(strWildcardFileMatches(intWildCardFileIndex), intExtensionLength)) = strFileExtensionsPrefList(intExtensionIndex) Then
                                        ' Match found; record in strWildcardFileMatch and stop searching
                                        strWildcardFileMatch = strWildcardFileMatches(intWildCardFileIndex)
                                        Exit For
                                    End If
                                Next intWildCardFileIndex
                            End If
                        Else
                            ' This is unexpected
                            Debug.Assert False
                        End If
                        If Len(strWildcardFileMatch) > 0 Then Exit For
                    Next intExtensionIndex
                
                    If Len(strWildcardFileMatch) = 0 Then
                        ' None of the files in strWildcardFileMatches() matched any of the extensions in strFileExtensionsPrefList()
                        ' Simply use the first file in strWildcardFileMatches()
                        ' This really shouldn't happen
                        Debug.Assert False
                        strWildcardFileMatch = strWildcardFileMatches(0)
                    End If
                End If
                
                If Len(strWildcardFileMatch) > 0 Then
                    strWildcardFileMatch = fso.BuildPath(fso.GetParentFolderName(udtAutoParams.FilePaths.InputFilePath), strWildcardFileMatch)
                    blnSuccess = fso.FileExists(strWildcardFileMatch)
                    If blnSuccess Then
                        Set fsInputFile = fso.GetFile(strWildcardFileMatch)
                        udtAutoParams.FilePaths.InputFilePath = fsInputFile.Path
                    Else
                        ' The wildcard specified in udtAutoParams.FilePaths.InputFilePath didn't match a file
                        ' An error will be logged below
                    End If
                End If
            End If
        End If
    End If
        
    ' Need to pre-read the .Ini file to look up the ExcludeIsoByFit, ExcludeIsoByFitMaxVal,
    '  RestrictIsoByAbundance, RestrictIsoAbundanceMin, RestrictIsoAbundanceMax,
    '  UsePEKBasedERValues, RestrictToEvenScanNumbersOnly, RestrictToOddScanNumbersOnly,
    '  MaximumDataCountEnabled, and database export settings
    
    ' We cannot read the entire .Ini file yet, since we haven't yet loaded the .Pek, .Csv, .mzXML, .mzData, or .Gel file and populated
    '  a GelBody() object and a GelData() object
    ' Thus, we'll use the IniFileReadSingleSetting function to read the necessary settings
    
    If Len(udtAutoParams.FilePaths.IniFilePath) > 0 Then
        ' IniFileReadSingleSetting requires a full path
        ' Thus, if udtAutoParams.FilePaths.IniFilePath does not contain "\" anywhere, then prepend with App.Path
        If InStr(udtAutoParams.FilePaths.IniFilePath, "\") = 0 Then
            udtAutoParams.FilePaths.IniFilePath = fso.BuildPath(App.Path, udtAutoParams.FilePaths.IniFilePath)
        End If
    
        With glbPreferencesExpanded
            If FileExists(udtAutoParams.FilePaths.IniFilePath) Then
                With .AutoAnalysisFilterPrefs
                    
                    strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "ExcludeIsoByFit", "False", udtAutoParams.FilePaths.IniFilePath)
                    .ExcludeIsoByFit = CBoolSafe(strKeyValue)
                    .ExcludeIsoByFitMaxVal = 0.15
                    
                    strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "RestrictIsoByAbundance", "False", udtAutoParams.FilePaths.IniFilePath)
                    .RestrictIsoByAbundance = CBoolSafe(strKeyValue)
                    .RestrictIsoAbundanceMin = 0
                    .RestrictIsoAbundanceMax = 1E+15
                    
                    If .RestrictIsoByAbundance Then
                        strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "RestrictIsoAbundanceMin", Trim(.RestrictIsoAbundanceMin), udtAutoParams.FilePaths.IniFilePath)
                        If IsNumeric(strKeyValue) Then
                            .RestrictIsoAbundanceMin = val(strKeyValue)
                        End If
                    
                        strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "RestrictIsoAbundanceMax", Trim(.RestrictIsoAbundanceMax), udtAutoParams.FilePaths.IniFilePath)
                        If IsNumeric(strKeyValue) Then
                            .RestrictIsoAbundanceMax = val(strKeyValue)
                        End If
                    End If
                    
                    If .ExcludeIsoByFit Then
                        strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "ExcludeIsoByFitMaxVal", Trim(.ExcludeIsoByFitMaxVal), udtAutoParams.FilePaths.IniFilePath)
                        If IsNumeric(strKeyValue) Then
                            .ExcludeIsoByFitMaxVal = val(strKeyValue)
                        End If
                    End If
                    
                    strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "RestrictToEvenScanNumbersOnly", "False", udtAutoParams.FilePaths.IniFilePath)
                    .RestrictToEvenScanNumbersOnly = CBoolSafe(strKeyValue)
                    
                    strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "RestrictToOddScanNumbersOnly", "False", udtAutoParams.FilePaths.IniFilePath)
                    .RestrictToOddScanNumbersOnly = CBoolSafe(strKeyValue)
                    
                    If udtAutoParams.FullyAutomatedPRISMMode Then
                        strDefault = "False"
                    Else
                        strDefault = "True"
                    End If
                    strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "MaximumDataCountEnabled", strDefault, udtAutoParams.FilePaths.IniFilePath)
                    .MaximumDataCountEnabled = CBoolSafe(strKeyValue)
                    
                    If .MaximumDataCountEnabled Then
                        strKeyValue = IniFileReadSingleSetting("AutoAnalysisFilterPrefs", "MaximumDataCountToLoad", Trim(.MaximumDataCountToLoad), udtAutoParams.FilePaths.IniFilePath)
                        If IsNumeric(strKeyValue) Then
                            .MaximumDataCountToLoad = val(strKeyValue)
                        End If
                    End If
                End With
                
                With .AutoAnalysisOptions
                    strKeyValue = IniFileReadSingleSetting("AutoAnalysisOptions", "PEKFileExtensionPreferenceOrder", .PEKFileExtensionPreferenceOrder, udtAutoParams.FilePaths.IniFilePath)
                    .PEKFileExtensionPreferenceOrder = strKeyValue
                End With
                
                strKeyValue = IniFileReadSingleSetting("ExpandedPreferences", "UsePEKBasedERValues", "False", udtAutoParams.FilePaths.IniFilePath)
                .UsePEKBasedERValues = CBoolSafe(strKeyValue)
                
                If APP_BUILD_DISABLE_MTS Then
                    udtAutoParams.InvalidExportPassword = True
                End If
                
                If Not udtAutoParams.FullyAutomatedPRISMMode And Not udtAutoParams.InvalidExportPassword Then
                    
                    ' See if the user has Export to DB enabled
                    ' If they do, prompt for the export password
                    strKeyValue = IniFileReadSingleSetting("AutoAnalysisOptions", "DoNotSaveOrExport", "False", udtAutoParams.FilePaths.IniFilePath)
                    
                    If CBoolSafe(strKeyValue) Then
                        ' User has DoNotSaveOrExport enabled
                        ' That will bypass all exporting; thus no need for password prompting
                    Else
                    
                        ' See if key DBExportPW is present
                        strKeyValue = IniFileReadSingleSetting("AutoAnalysisOptions", "DBExportPW", "", udtAutoParams.FilePaths.IniFilePath)
                        If strKeyValue <> EXPORT_TO_DB_PASSWORD Then
                            ' Query user
                            strKeyValue = IniFileReadSingleSetting("AutoAnalysisOptions", "AutoAnalysisSearchModeCount", "0", udtAutoParams.FilePaths.IniFilePath)
                            intAutoAnalysisSearchModeCount = CIntSafe(strKeyValue)
                            
                            For intIndex = 1 To intAutoAnalysisSearchModeCount
                                strKeyValue = IniFileReadSingleSetting("AutoAnalysisSearchMode" & Trim(intIndex), "ExportResultsToDatabase", "False", udtAutoParams.FilePaths.IniFilePath)
                                
                                If CBoolSafe(strKeyValue) Then
                                    strMessage = "The Export Results to Database option is enabled in the .Ini file.  Exporting results to the database is an advanced feature that should normally only be performed during VIPER Automated PRISM Analysis Mode.  Please enter the password for exporting results to the database, or click Cancel to continue auto-analysis without database exporting."
                                    If Not QueryUserForExportToDBPassword(strMessage) Then
                                        udtAutoParams.InvalidExportPassword = True
                                    End If
                                    Exit For
                                End If
                            Next intIndex
                        End If
                    End If
                End If
                
            Else
                With .AutoAnalysisFilterPrefs
                    .ExcludeIsoByFit = False
                    .ExcludeIsoByFitMaxVal = 0.15
                    .RestrictCSByAbundance = False
                    .RestrictIsoByAbundance = False
                    
                    .RestrictToEvenScanNumbersOnly = False
                    .RestrictToOddScanNumbersOnly = False
                    
                    If udtAutoParams.FullyAutomatedPRISMMode Then
                        .MaximumDataCountEnabled = False
                    Else
                        .MaximumDataCountEnabled = True
                        .MaximumDataCountToLoad = DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD
                    End If
                End With
                
                .UsePEKBasedERValues = False
                
                udtAutoParams.InvalidExportPassword = True
            End If
        End With
        
    End If
    
    
    If Not blnSuccess Then
        ' udtAutoParams.FilePaths.InputFilePath is empty or doesn't exist
        If udtAutoParams.ShowMessages Then
            ' Prompt user for file and try to load
            If Len(udtAutoParams.FilePaths.InputFilePath) > 0 Then
                MsgBox "Input file path does not point to a valid " & KNOWN_FILE_EXTENSIONS & " file: " & vbCrLf & udtAutoParams.FilePaths.InputFilePath & vbCrLf & "Please choose a valid file (extension " & KNOWN_FILE_EXTENSIONS_WITH_GEL & ")", vbExclamation + vbOKOnly, "File not found"
            End If
            
            udtWorkingParams.GelIndex = FileNew(MDIForm1.hwnd, "", 0, strErrorMessage)
            If udtWorkingParams.GelIndex > 0 Then
                ' File loaded successfully
                udtAutoParams.FilePaths.InputFilePath = GelData(udtWorkingParams.GelIndex).FileName
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Loading File; " & GetFileInfo(GelData(udtWorkingParams.GelIndex).FileName)
                blnSuccess = True
            Else
                ' Load failed
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - " & strErrorMessage
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or DATAFILE_LOAD_ERROR_BIT
                blnSuccess = False
            End If
        Else
            ' Load failed
            If Not blnInvalidFolder Then
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - File not found: " & udtAutoParams.FilePaths.InputFilePath
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or DATAFILE_LOAD_ERROR_BIT
            End If
        End If
    Else
        ' udtAutoParams.FilePaths.InputFilePath was found; try to load
        
        If DetermineFileType(udtAutoParams.FilePaths.InputFilePath, eFileType) Then
            If eFileType = ifmGelFile Then
                ' Loading a .Gel file; use ReadGelFile
                udtWorkingParams.GelIndex = ReadGelFile(udtAutoParams.FilePaths.InputFilePath, udtAutoParams.GelIndexToForce)
                udtWorkingParams.LoadedGelFile = True
            Else
                ' Loading a .Pek, .CSV, .mzXML, or .mzData file; use FileNew
                If udtAutoParams.GelIndexToForce > 0 And udtAutoParams.GelIndexToForce <= UBound(GelBody()) Then
                    udtWorkingParams.GelIndex = FileNew(MDIForm1.hwnd, udtAutoParams.FilePaths.InputFilePath, udtAutoParams.GelIndexToForce, strErrorMessage)
                Else
                    udtWorkingParams.GelIndex = FileNew(MDIForm1.hwnd, udtAutoParams.FilePaths.InputFilePath, 0, strErrorMessage)
                End If
                udtWorkingParams.LoadedGelFile = False
            End If
        Else
            ' Invalid file type
            strErrorMessage = "Unknown file type for " & fso.GetFileName(udtAutoParams.FilePaths.InputFilePath) & "; should be " & KNOWN_FILE_EXTENSIONS_WITH_GEL
        End If
        
        If udtWorkingParams.GelIndex > 0 Then
            ' File loaded successfully
            blnSuccess = True
        Else
            ' Load failed
            
            If fso.FileExists(udtAutoParams.FilePaths.InputFilePath) Then
                ' Try to determine the size of the file
                On Error Resume Next
                Set fsInputFile = fso.GetFile(udtAutoParams.FilePaths.InputFilePath)
                
                On Error GoTo LoadInputFileErrorHandler
                If fsInputFile Is Nothing Then
                    strErrorMessage = strErrorMessage & " (file not found)"
                Else
                    strErrorMessage = strErrorMessage & " (file size = " & Trim(Round(fsInputFile.Size / 1024, 0)) & " Kb)"
                    Set fsInputFile = Nothing
                End If
            Else
                strErrorMessage = strErrorMessage & " (file not found)"
            End If
            
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - " & strErrorMessage & ": " & udtAutoParams.FilePaths.InputFilePath
            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or DATAFILE_LOAD_ERROR_BIT
            blnSuccess = False
        End If
        
    End If
    
    If blnSuccess Then
        ' Check for file load warning messages
        strHistoryMatch = FindSettingInAnalysisHistory(udtWorkingParams.GelIndex, CSV_COLUMN_HEADER_UNKNOWN_WARNING, lngHistoryIndexLastMatch, False, ":", ";")
        If lngHistoryIndexLastMatch >= 0 Then
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or MISCELLANEOUS_MESSAGE_WARNING_BIT
        End If
        
        strHistoryMatch = FindSettingInAnalysisHistory(udtWorkingParams.GelIndex, CSV_COLUMN_HEADER_MISSING_WARNING, lngHistoryIndexLastMatch, False, ":", ";")
        If lngHistoryIndexLastMatch >= 0 Then
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or MISCELLANEOUS_MESSAGE_WARNING_BIT
        End If
    End If
    
    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    AutoAnalysisLoadInputFile = blnSuccess
    Exit Function
    
LoadInputFileErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while loading the " & KNOWN_FILE_EXTENSIONS_WITH_GEL & " file during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or DATAFILE_LOAD_ERROR_BIT
    AutoAnalysisLoadInputFile = False
End Function

Private Sub AutoAnalysisLog(ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, strLogText As String)
    Dim strLogTextWithExtras As String
    
On Error Resume Next
    
    If Len(glbPreferencesExpanded.AutoAnalysisStatus.AutoAnalysisTimeStamp) > 0 Then
        strLogTextWithExtras = glbPreferencesExpanded.AutoAnalysisStatus.AutoAnalysisTimeStamp & " > " & strLogText
    Else
        strLogTextWithExtras = strLogText
    End If
    
    If Not udtWorkingParams.ts Is Nothing Then
        udtWorkingParams.ts.WriteLine strLogTextWithExtras
    End If
    
    ' Append to mMemoryLog
    mMemoryLog = mMemoryLog & strLogTextWithExtras & vbCrLf
    
    ' Check for CTL2DHEATMAP_ERROR_MESSAGE
    If InStr(strLogText, CTL2DHEATMAP_ERROR_MESSAGE) > 0 Then
        ' Message found; exit Viper as soon as possible
        udtAutoParams.ExitViperASAP = True
        udtAutoParams.ExitViperReason = CTL2DHEATMAP_ERROR_MESSAGE
        udtAutoParams.RestartAfterExit = True
    End If
End Sub

Private Sub AutoAnalysisMemoryLogClear()
    mMemoryLog = ""
End Sub

Public Function AutoAnalysisMemoryLogGet() As String
    AutoAnalysisMemoryLogGet = mMemoryLog
End Function

Private Sub AutoAnalysisPerformNETAdjustment(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
    
    Const ITERATION_STOP_CHANGE = 4
    
    Dim blnSkipGANETComputation As Boolean
    Dim dblGANETSlope As Double, dblGANETIntercept As Double
    Dim lngTimerStart As Long
    
    Dim strWarningMessage As String
    Dim strNetAdjUMCsWithDBHits As String
    Dim strMessage As String
    Dim strLastGoodLocation As String
    
    Dim blnUseRobustNETWarping As Boolean

    Dim eRobustNETAdjustmentModeSaved As UMCRobustNETModeConstants
    
    Dim objMSAlign As frmMSAlign
    
On Error GoTo PerformNETAdjustmentErrorHandler
    
    With glbPreferencesExpanded
        strLastGoodLocation = "Check .SkipGANETSlopeAndInterceptComputation"
        blnSkipGANETComputation = .AutoAnalysisOptions.SkipGANETSlopeAndInterceptComputation
        If blnSkipGANETComputation And Not udtWorkingParams.LoadedGelFile Then
            If .AutoAnalysisDBInfoIsValid Then
                With .AutoAnalysisDBInfo
                    If .GANET_Slope <> 0 Then
                        GelAnalysis(udtWorkingParams.GelIndex).GANET_Slope = .GANET_Slope
                        GelAnalysis(udtWorkingParams.GelIndex).GANET_Intercept = .GANET_Intercept
                        ' Need to assign a non-zero value to GANET_Fit; we'll assign 1.11E-3 with all 1's so it stands out
                        GelAnalysis(udtWorkingParams.GelIndex).GANET_Fit = 1.11111111111111E-03
                    Else
                        blnSkipGANETComputation = False
                    End If
                End With
            Else
                blnSkipGANETComputation = False
            End If
        
            If Not blnSkipGANETComputation Then
                If GelUMCNETAdjDef(udtWorkingParams.GelIndex).InitialSlope <> 0 Then
                    GelAnalysis(udtWorkingParams.GelIndex).GANET_Slope = GelUMCNETAdjDef(udtWorkingParams.GelIndex).InitialSlope
                    GelAnalysis(udtWorkingParams.GelIndex).GANET_Intercept = GelUMCNETAdjDef(udtWorkingParams.GelIndex).InitialIntercept
                    ' Need to assign a non-zero value to GANET_Fit; we'll assign 1.11E-3 with all 1's so it stands out
                    GelAnalysis(udtWorkingParams.GelIndex).GANET_Fit = 1.11111111111111E-03
                    blnSkipGANETComputation = True
                End If
            End If
        End If
    End With
    
    If blnSkipGANETComputation Then
        strWarningMessage = "Warning - NET Adjustment has been skipped"
        udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or NET_ADJUSTMENT_SKIPPED_WARNING_BIT
        udtWorkingParams.NETDefined = True
    Else
        ' Compute the NET Adjustment
        ' This could be Simple Iterative NET adjustment, Robust NET Iterative adjustment, or UMCRobustNETWarp
        ' For iterative adjustment during auto analysis, the three stop criteria are:
        '   Max Iterations, Minimum ID count, and Change Threshold
        ' Must set Iteration Stop Type to ITERATION_STOP_CHANGE
        
        strLastGoodLocation = "Customize GelUMCNETAdjDef(udtWorkingParams.GelIndex)"
        With GelUMCNETAdjDef(udtWorkingParams.GelIndex)
            .IterationStopType = ITERATION_STOP_CHANGE
            .IterationStopValue = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentChangeThresholdStopValue
            .NETTolIterative = glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentInitialNetTol
            .NETFormula = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(udtWorkingParams.GelIndex))        ' Make sure .NETFormula is the default
            .UseNET = True
            .IterationAcceptLast = True
            
            If APP_BUILD_DISABLE_LCMSWARP Then
                If .RobustNETAdjustmentMode <> UMCRobustNETModeConstants.UMCRobustNETIterative Then
                    .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETIterative
                End If
            End If
            
            If .UseRobustNETAdjustment And .RobustNETAdjustmentMode >= UMCRobustNETModeConstants.UMCRobustNETWarpTime Then
                blnUseRobustNETWarping = True
            Else
                blnUseRobustNETWarping = False
            End If
        End With
        
        ' In case NET adjustment fails, make sure .Slope and .Intercept are 0
        If Not GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
            ' Examine .GANET_Slope and .GANET_Intercept for unusual values
            With GelAnalysis(udtWorkingParams.GelIndex)
                .GANET_Slope = 0
                .GANET_Intercept = 0
            End With
        Else
            ' This shouldn't happen
            Debug.Assert False
        End If
        
        If blnUseRobustNETWarping And Not APP_BUILD_DISABLE_LCMSWARP Then
            With GelUMCNETAdjDef(udtWorkingParams.GelIndex)
                eRobustNETAdjustmentModeSaved = .RobustNETAdjustmentMode
                .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTime
            End With
            
            strLastGoodLocation = "Instantiate objMSAlign"
            Set objMSAlign = New frmMSAlign
            DoEvents
            Sleep 250
            DoEvents
            
            ' Perform NET Alignment using MS Warp
            ' For now, just perform NET alignment and not mass recalibration (which is done in AutoAnalysisToleranceRefinement)
            
            strLastGoodLocation = "Set objMSAlign.CallerID"
            objMSAlign.CallerID = udtWorkingParams.GelIndex
            objMSAlign.RecalibratingMassDuringAutoAnalysis = False
            
            DoEvents
            Sleep 250
            DoEvents
            strLastGoodLocation = "Show form objMSAlign"
            objMSAlign.Show vbModeless
            
            strLastGoodLocation = "InitializeSearch"
            objMSAlign.InitializeSearch
            objMSAlign.SetPlotPointSize RESIDUALS_PLOT_POINT_SIZE

            strLastGoodLocation = "StartAlignment"
            objMSAlign.StartAlignment
        
            lngTimerStart = Timer()
            Do
                Sleep 200
                DoEvents
                strLastGoodLocation = "Waiting while objMSAlign.MassMatchState = pscRunning (" & CStr(Timer - lngTimerStart) & " seconds elapsed)"
            Loop While objMSAlign.MassMatchState = pscRunning
          
            If objMSAlign.MassMatchState = pscComplete Then
                strLastGoodLocation = "objMSAlign.MassMatchState now equals pscComplete"
                
                udtWorkingParams.NETDefined = True

                ' Alignment succeeded
                ' Need to wait until objMSAlign.LocalGelUpdated = True
                strLastGoodLocation = "Wait objMSAlign.LocalGelUpdated = True"
                Do
                    Sleep 50
                Loop While Not objMSAlign.LocalGelUpdated
                
                ' Alignment is completed, the custom NET values have been updated,
                ' and the slope and intercept have been defined
                
                ' Save the NET Alignment Surface and NET Residual plots to disk
                strLastGoodLocation = "Call AutoAnalysisSaveNETSurfaceAndResidualPlots"
                AutoAnalysisSaveNETSurfaceAndResidualPlots objMSAlign, udtAutoParams, udtWorkingParams, fso, False
                
            Else
                ' Alignment failed
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
                udtWorkingParams.NETDefined = False
            End If
            
            strLastGoodLocation = "Unload objMSAlign"
            Unload objMSAlign
            
            ' Restore RobustNETAdjustmentMode
            With GelUMCNETAdjDef(udtWorkingParams.GelIndex)
                .RobustNETAdjustmentMode = eRobustNETAdjustmentModeSaved
            End With
        Else
            
            strLastGoodLocation = "With frmSearchForNETAdjustmentUMC"
            With frmSearchForNETAdjustmentUMC
                strLastGoodLocation = "Set frmSearchForNETAdjustmentUMC.CallerID"
                .CallerID = udtWorkingParams.GelIndex
                
                ' Note: Must use vbModeLess to prevent App from waiting for form to close
                strLastGoodLocation = "Show frmSearchForNETAdjustmentUMC"
                .Show vbModeless
                
                strLastGoodLocation = "InitializeNETAdjustment"
                .InitializeNETAdjustment
                
                ' Perform NET adjustment
                strLastGoodLocation = "CalculateNETAdjustmentStart"
                .CalculateNETAdjustmentStart
                
                ' Check that the number of IDs was at least .NETAdjustmentMinIDCount
                If .GetNETAdjustmentIDCount() < glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount Then
                    strWarningMessage = "Warning - Number of LC-MS Features used for NET adjustment is low: " & Trim(.GetNETAdjustmentIDCount) & " vs. expected minimum of " & Trim(glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIDCount)
                    udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or NET_ADJUSTMENT_LOW_ID_COUNT_WARNING_BIT
                Else
                    ' Check if the number of iterations was too low
                    strNetAdjUMCsWithDBHits = FindSettingInAnalysisHistory(udtWorkingParams.GelIndex, UMC_NET_ADJ_ITERATION_COUNT, 0, True, "=", ";")
                    If IsNumeric(strNetAdjUMCsWithDBHits) Then
                        If CLng(strNetAdjUMCsWithDBHits) < glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIterationCount Then
                            strWarningMessage = "Warning - Number of NET adjustment iterations is low: " & Trim(strNetAdjUMCsWithDBHits) & " vs. expected minimum of " & Trim(glbPreferencesExpanded.AutoAnalysisOptions.NETAdjustmentMinIterationCount)
                            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or NET_ADJUSTMENT_LOW_ID_COUNT_WARNING_BIT
                        End If
                    End If
                End If
            End With
            udtWorkingParams.NETDefined = True
            strLastGoodLocation = "Unload frmSearchForNETAdjustmentUMC"
            Unload frmSearchForNETAdjustmentUMC
        End If

    End If

    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
    
    If Len(strWarningMessage) > 0 Then
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strWarningMessage
    End If

    If Not GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
        ' Examine .GANET_Slope and .GANET_Intercept for unusual values
        strLastGoodLocation = "Examine .GANET_Slope and .GANET_Intercept"
        With GelAnalysis(udtWorkingParams.GelIndex)
            dblGANETSlope = .GANET_Slope
            dblGANETIntercept = .GANET_Intercept
        End With
        
        If dblGANETSlope = 0 Then
            ' Warn if the Slope = 0
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - The NET Slope is 0; search results will be incorrect"
            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
        Else
            ' Raise an error if the Slope is outside of the expected range
            ' Raise a warning if the Intercept is outside of the expected range
            If GelUMCNETAdjDef(udtWorkingParams.GelIndex).UseRobustNETAdjustment And Not blnUseRobustNETWarping Then
                ' Used Robust NET (but not NET warping); use the Robust NET ranges as the expected minima and maxima
                With GelUMCNETAdjDef(udtWorkingParams.GelIndex)
                    If dblGANETSlope < .RobustNETSlopeStart / 2 Or dblGANETSlope > .RobustNETSlopeEnd * 2 Then
                        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - NET_Slope is outside of the expected range: " & Trim(dblGANETSlope) & " vs. expected range of " & Trim(.RobustNETSlopeStart / 2) & " to " & Trim(.RobustNETSlopeEnd * 2)
                        udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
                    End If
                    
                    If dblGANETIntercept < .RobustNETInterceptStart Or dblGANETIntercept > .RobustNETInterceptEnd Then
                        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - NET_Intercept is outside of the expected range: " & Trim(dblGANETIntercept) & " vs. expected range of " & Trim(.RobustNETInterceptStart) & " to " & Trim(.RobustNETInterceptEnd)
                        udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or GANET_INTERCEPT_WARNING_BIT
                    End If
                End With
            Else
                ' Used Iterative NET or NET Warping; use .NETSlopeExpectedMinimum and .NETSlopeExpectedMaximum
                With glbPreferencesExpanded.AutoAnalysisOptions
                    If dblGANETSlope < .NETSlopeExpectedMinimum Or dblGANETSlope > .NETSlopeExpectedMaximum Then
                        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - NET_Slope is outside of the expected range: " & Trim(dblGANETSlope) & " vs. expected range of " & Trim(.NETSlopeExpectedMinimum) & " to " & Trim(.NETSlopeExpectedMaximum)
                        udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
                    End If
                    
                    If dblGANETIntercept < .NETInterceptExpectedMinimum Or dblGANETIntercept > .NETInterceptExpectedMaximum Then
                        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - NET_Intercept is outside of the expected range: " & Trim(dblGANETIntercept) & " vs. expected range of " & Trim(.NETInterceptExpectedMinimum) & " to " & Trim(.NETInterceptExpectedMaximum)
                        udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or GANET_INTERCEPT_WARNING_BIT
                    End If
                End With
            End If
        End If
    Else
        ' This shouldn't happen
        Debug.Assert False
    End If
    
    Exit Sub
    
PerformNETAdjustmentErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while adjusting the NET values during auto analysis (LastGoodLocation=" & strLastGoodLocation & "): " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

    If Err.Number = 7 Then
        ' Out of Memory Error; stop Viper as soon as possible
        udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
        udtAutoParams.ExitViperASAP = True
        udtAutoParams.ExitViperReason = "Out of memory"
        udtAutoParams.RestartAfterExit = True
    End If
    
End Sub

Private Sub AutoAnalysisSaveNETSurfaceAndResidualPlots(ByRef objMSAlign As frmMSAlign, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef fso As FileSystemObject, ByVal blnMassAndNETAlignment As Boolean)
    
    Dim lngWindowWidth As Long
    Dim lngWindowHeight As Long
    Dim lngWindowWidthTwips As Long
    Dim lngWindowHeightTwips As Long
   
    Dim lngDefaultMaxMontageWidth As Long
    Dim lngDefaultMaxMontageHeight As Long
   
    Dim strFilePath As String
    Dim strFilePathB As String
    Dim strCombinedFilePath As String
    Dim strWorkingCombinedFilePath As String
    Dim strUniqueSuffix As String
    
    Dim strLastGoodLocation As String
    Dim strMessage As String
    
    Dim lngResult As Long
    Dim blnSaveSuccessful As Boolean
    
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
   
On Error GoTo AutoAnalysisSaveNETSurfaceAndResidualPlotsErrorhandler

    ' Enlarge the window to the desired size (pixels)
    
    strLastGoodLocation = "Validate the dimensions"
    lngWindowWidth = glbPreferencesExpanded.AutoAnalysisOptions.SavePictureWidthPixels
    lngWindowHeight = glbPreferencesExpanded.AutoAnalysisOptions.SavePictureHeightPixels
    ValidateValueLng lngWindowWidth, 64, 10000, 1024
    ValidateValueLng lngWindowHeight, 64, 10000, 768
    
    With objMSAlign
        ' The user may have minimized the window during auto analysis
        strLastGoodLocation = "Assure the WindowState = vbNormal"
        If .WindowState <> vbNormal Then
            .WindowState = vbNormal
        End If
        
        strLastGoodLocation = "Resize the window"
        .ScaleMode = vbTwips
        lngWindowWidthTwips = lngWindowWidth * Screen.TwipsPerPixelX * 1.3
        lngWindowHeightTwips = lngWindowHeight * Screen.TwipsPerPixelY * 1.3
        .width = lngWindowWidthTwips
        .Height = lngWindowHeightTwips
        .Top = 0
        .Left = 0
        DoEvents
    
        ' -----------------------------------------
        ' Save the NET Alignment Surface (Deep's Control)
        ' -----------------------------------------
        
        strLastGoodLocation = "Define the filename"
        If Not blnMassAndNETAlignment Then
            strFilePath = udtWorkingParams.ResultsFileNameBase & "_NETAlignmentSurface.png"
        Else
            strFilePath = udtWorkingParams.ResultsFileNameBase & "_NETandMassAlignmentSurface.png"
        End If

        strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
        
        strLastGoodLocation = "Save Flat View to: " & strFilePath
        TraceLog 5, "AutoAnalysisSaveNETSurfaceAndResidualPlots", strLastGoodLocation
        blnSaveSuccessful = .SaveFlatViewToPNG(strFilePath)
        
        If blnSaveSuccessful Then
            If Not blnMassAndNETAlignment Then
                AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), HTML_SUMMARY_FILE_HEADING_NET_ALIGNMENT, 4, 1
                udtWorkingParams.TICPlotsStartRow = 5
            Else
                ' Replace the "NET Alignment Surface" entry in .GraphicOutputFileInfo with this one
                ReplaceOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), HTML_SUMMARY_FILE_HEADING_NET_ALIGNMENT, HTML_SUMMARY_FILE_HEADING_NET_ALIGNMENT_MASS_CAL, 4, 1
                udtWorkingParams.TICPlotsStartRow = 5
            End If
        End If
    
    
        ' -----------------------------------------
        ' Shrink the window to 55% of the current height so that the two NET residuals plots will look nice stacked
        ' In addition, shrink the width to 75% of the current width
        ' -----------------------------------------
        strLastGoodLocation = "Shrink the window"
        .width = lngWindowWidthTwips * 0.75
        .Height = lngWindowHeightTwips * 0.55
        DoEvents
        
        lngDefaultMaxMontageWidth = .width / 23.25466
        lngDefaultMaxMontageHeight = .Height / 29.5233
        
        ' -----------------------------------------
        ' Save the Linear Fit NET Residuals plot
        ' -----------------------------------------
        strUniqueSuffix = "_" & Mid(Format(Rnd(), "0.0000000"), 3)
        strFilePath = udtWorkingParams.ResultsFileNameBase & "_NETAlignmentResidualsVsLinearFit" & strUniqueSuffix & ".png"
        strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
        strFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, True)
        
        strLastGoodLocation = "Call UpdatePlotViewModeToLinearFitNETResidualsPlot"
        TraceLog 5, "AutoAnalysisSaveNETSurfaceAndResidualPlots", strLastGoodLocation
        .UpdatePlotViewModeToLinearFitNETResidualsPlot
        Sleep 100
        DoEvents
        
        strLastGoodLocation = "Save Linear Fit NET Residuals Plot to: " & strFilePath
        TraceLog 5, "AutoAnalysisSaveNETSurfaceAndResidualPlots", strLastGoodLocation
        blnSaveSuccessful = .SaveNETResidualsPlotToPNG(strFilePath)
        
        If blnSaveSuccessful Then
            ' Save the Warped NET Residuals plot
            strFilePathB = udtWorkingParams.ResultsFileNameBase & "_NETAlignmentResidualsWarped" & strUniqueSuffix & ".png"
            strFilePathB = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePathB)
            strFilePathB = objRemoteSaveFileHandler.GetTempFilePath(strFilePathB, True)
            
            strLastGoodLocation = "Call UpdatePlotViewModeToWarpedFitNETResidualsPlot"
            TraceLog 5, "AutoAnalysisSaveNETSurfaceAndResidualPlots", strLastGoodLocation
            .UpdatePlotViewModeToWarpedFitNETResidualsPlot
            Sleep 100
            DoEvents
            
            strLastGoodLocation = "Save Warped NET Residuals Plot to: " & strFilePathB
            TraceLog 5, "AutoAnalysisSaveNETSurfaceAndResidualPlots", strLastGoodLocation
            blnSaveSuccessful = .SaveNETResidualsPlotToPNG(strFilePathB)
            
            If blnSaveSuccessful Then
                ' Create a montage using the two files
                If Not blnMassAndNETAlignment Then
                    strCombinedFilePath = udtWorkingParams.ResultsFileNameBase & "_NETAlignmentResiduals.png"
                Else
                    strCombinedFilePath = udtWorkingParams.ResultsFileNameBase & "_NETAlignmentResidualsWithMassCal.png"
                End If
                
                strCombinedFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strCombinedFilePath)
                strWorkingCombinedFilePath = objRemoteSaveFileHandler.GetTempFilePath(strCombinedFilePath, False)
                
                strLastGoodLocation = "CreateMontageImageFile"
                lngResult = CreateMontageImageFile(udtWorkingParams, strFilePath, strFilePathB, strWorkingCombinedFilePath, lngDefaultMaxMontageWidth, lngDefaultMaxMontageHeight)
                
                If lngResult = 0 Then
                    objRemoteSaveFileHandler.MoveTempFileToFinalDestination
                    blnSaveSuccessful = True
                Else
                    blnSaveSuccessful = False
                End If
                
                If blnSaveSuccessful Then
                    strLastGoodLocation = "AddNewOutputFileForHtml or ReplaceOutputFileForHtml"
                    
                    If Not blnMassAndNETAlignment Then
                        AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strCombinedFilePath), HTML_SUMMARY_FILE_HEADING_NET_RESIDUALS, 4, 2
                        udtWorkingParams.TICPlotsStartRow = 5
                    Else
                        ' Replace the "NET Alignment Residuals" entry in .GraphicOutputFileInfo with this one
                        ReplaceOutputFileForHtml udtWorkingParams, fso.GetFileName(strCombinedFilePath), HTML_SUMMARY_FILE_HEADING_NET_RESIDUALS, HTML_SUMMARY_FILE_HEADING_NET_RESIDUALS_MASS_CAL, 4, 2
                        udtWorkingParams.TICPlotsStartRow = 5
                    End If
                End If
            End If
        End If
    
        ' Restore the window size
        .width = lngWindowWidthTwips
        .Height = lngWindowHeightTwips
        DoEvents
    End With
    Exit Sub

AutoAnalysisSaveNETSurfaceAndResidualPlotsErrorhandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while saving NET surface and residual plots during auto analysis (LastGoodLocation=" & strLastGoodLocation & "): " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

''Public Sub TestAutoNETWarp(byval lngGelIndex As Long, Optional byval strResultsFileNameBase As String = "MyTestFile")
''
''    Dim dblGANETSlope As Double, dblGANETIntercept As Double
''
''    Dim lngWindowWidth As Long
''    Dim lngWindowHeight As Long
''    Dim lngWindowWidthTwips As Long
''    Dim lngWindowHeightTwips As Long
''
''    Dim strOutputFolder As String
''    Dim strFilePath As String
''    Dim strFilePathB As String
''    Dim strCombinedFilePath As String
''    Dim strWorkingCombinedFilePath As String
''    Dim strUniqueSuffix As String
''
''    Dim lngResult As Long
''    Dim blnSaveSuccessful As Boolean
''
''    Dim objMSAlign As frmMSAlign
''    Dim fso As FileSystemObject
''
''    strOutputFolder = App.Path
''
''    Set objMSAlign = New frmMSAlign
''    With objMSAlign
''        ' Perform Alignment using MS Warp
''
''        .CallerID = lngGelIndex
''        .RecalibratingMassDuringAutoAnalysis = False
''        .show vbModeless
''        .InitializeSearch
''        .StartAlignment
''
''        Do
''            Sleep 200
''            DoEvents
''        Loop While objMSAlign.MassMatchState = pscRunning
''
''        If objMSAlign.MassMatchState = pscComplete Then
''            ' Alignment succeeded
''            ' Need to wait until objMSAlign.LocalGelUpdated = True
''            Do
''                Sleep 50
''            Loop While Not objMSAlign.LocalGelUpdated
''
''            ' Alignment is completed, the custom NET values have been updated,
''            ' and the slope and intercept have been defined
''
''            ' Save several plots to disk
''
''            ' Enlarge the window to the desired size (pixels)
''            lngWindowWidth = glbPreferencesExpanded.AutoAnalysisOptions.SavePictureWidthPixels
''            lngWindowHeight = glbPreferencesExpanded.AutoAnalysisOptions.SavePictureHeightPixels
''            ValidateValueLng lngWindowWidth, 64, 10000, 1024
''            ValidateValueLng lngWindowHeight, 64, 10000, 768
''
''            ' The user may have minimized the window during auto analysis
''            If .WindowState <> vbNormal Then
''                .WindowState = vbNormal
''            End If
''
''            .ScaleMode = vbTwips
''            lngWindowWidthTwips = lngWindowWidth * Screen.TwipsPerPixelX * 1.3
''            lngWindowHeightTwips = lngWindowHeight * Screen.TwipsPerPixelY * 1.3
''            .width = lngWindowWidthTwips
''            .Height = lngWindowHeightTwips
''            .Top = 0
''            .Left = 0
''            DoEvents
''
''            Set fso = New FileSystemObject
''
''            ' Save the NET Alignment Surface
''            strFilePath = strResultsFileNameBase & "_NETAlignmentSurface.png"
''            strFilePath = fso.BuildPath(strOutputFolder, strFilePath)
''
''            TraceLog 5, "AutoAnalysisPerformNETAdjustment", "Save Flat View to: " & strFilePath
''            blnSaveSuccessful = .SaveFlatViewToPNG(strFilePath)
''
''            If blnSaveSuccessful Then
''                Debug.Print "Save successful: " & strFilePath
''            Else
''                Debug.Print "Save failed: " & strFilePath
''            End If
''
''
''            ' Shrink the window to 50% of the current height so that the two residuals plots will look nice stacked
''            ' In addition, shrink the width to 70% of the current width
''            .width = lngWindowWidthTwips * 0.75
''            .Height = lngWindowHeightTwips * 0.55
''            DoEvents
''
''            ' Save the Linear Fit Residuals plot
''            strFilePath = strResultsFileNameBase & "_NETAlignmentResidualsVsLinearFit.png"
''            strFilePath = fso.BuildPath(strOutputFolder, strFilePath)
''
''            TraceLog 5, "AutoAnalysisPerformNETAdjustment", "Call UpdatePlotViewModeToLinearFitNETResidualsPlot"
''            .UpdatePlotViewModeToLinearFitNETResidualsPlot
''            Sleep 100
''            DoEvents
''
''            TraceLog 5, "AutoAnalysisPerformNETAdjustment", "Save Linear Fit Residuals Plot to: " & strFilePath
''            blnSaveSuccessful = .SaveNETResidualsPlotToPNG(strFilePath)
''
''            If blnSaveSuccessful Then
''                ' Save the Warped Residuals plot
''                strFilePathB = strResultsFileNameBase & "_NETAlignmentResidualsWarped.png"
''                strFilePathB = fso.BuildPath(strOutputFolder, strFilePathB)
''
''                TraceLog 5, "AutoAnalysisPerformNETAdjustment", "Call UpdatePlotViewModeToWarpedFitNETResidualsPlot"
''                .UpdatePlotViewModeToWarpedFitNETResidualsPlot
''                Sleep 100
''                DoEvents
''
''                TraceLog 5, "AutoAnalysisPerformNETAdjustment", "Save Warped Residuals Plot to: " & strFilePathB
''                blnSaveSuccessful = .SaveNETResidualsPlotToPNG(strFilePathB)
''
''                If blnSaveSuccessful Then
''                    ' Create a montage using the two files
''                    strCombinedFilePath = strResultsFileNameBase & "_NETAlignmentResiduals.png"
''                    strCombinedFilePath = fso.BuildPath(strOutputFolder, strCombinedFilePath)
''
''                    lngResult = CreateMontageImageFile(udtWorkingParams, strFilePath, strFilePathB, strWorkingCombinedFilePath)
''                    If lngResult = 0 Then
''                        blnSaveSuccessful = True
''                    Else
''                        blnSaveSuccessful = False
''                    End If
''
''                      If blnSaveSuccessful Then
''                        Debug.Print "Save successful: " & strFilePath
''                    Else
''                        Debug.Print "Save failed: " & strFilePath
''                    End If
''                End If
''            End If
''
''            ' Restore the window size
''            .width = lngWindowWidthTwips
''            .Height = lngWindowHeightTwips
''            DoEvents
''
''        Else
''            ' Alignment failed
''            Debug.Print "Alignment Failed"
''        End If
''    End With
''    Unload objMSAlign
''
''End Sub

Private Sub AutoAnalysisRemovePairMemberHits(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType)
    ' Looks for unambiguous pairs
    ' If they have any database hits, then removes either the light or heavy UMC for the pair
    
    
    Dim lngUMCIndex As Long
    Dim lngUMCIndexLight As Long
    
    Dim lngPairIndex As Long
    
    Dim lngUMCCountToRemove As Long
    Dim lngNewUMCCount As Long
    Dim lngNewPairCount As Long
    
    Dim blnPossiblyRemoveUMC As Boolean
    
    Dim blnRemoveUMC() As Boolean           ' 0-based array; corresponds to GelUMC().UMCs
    Dim lngNewUMCIndex() As Long
    
    Dim eClsPaired() As umcpUMCPairMembershipConstants
    
    Dim strMessage As String
    Dim strMessageSuffix As String
    
On Error GoTo RemovePairMemberHitsErrorHandler

    If GelUMC(udtWorkingParams.GelIndex).UMCCnt <= 0 Then
        ' No LC-MS Features; nothing to do
        Exit Sub
    End If
    
    UpdateUMCsPairingStatus udtWorkingParams.GelIndex, eClsPaired()
    
    With GelUMC(udtWorkingParams.GelIndex)
        If .UMCCnt > 0 Then
        ReDim blnRemoveUMC(0 To .UMCCnt - 1)
        ReDim lngNewUMCIndex(0 To .UMCCnt - 1)
        End If
    End With

    If GelP_D_L(udtWorkingParams.GelIndex).PCnt > 0 Then
        For lngPairIndex = 0 To GelP_D_L(udtWorkingParams.GelIndex).PCnt - 1
        
            blnPossiblyRemoveUMC = False
            If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsRemoveHeavy Then
                ' Only consider removing the heavy member UMC if its state is umcpHeavyUnique or umcpHeavyMultiple
                lngUMCIndex = GelP_D_L(udtWorkingParams.GelIndex).Pairs(lngPairIndex).P2
                If eClsPaired(lngUMCIndex) = umcpHeavyUnique Or eClsPaired(lngUMCIndex) = umcpHeavyMultiple Then
                    blnPossiblyRemoveUMC = True
                End If
            Else
                ' Only consider removing the light member UMC if its state is umcpLightUnique or umcpLightMultiple
                lngUMCIndex = GelP_D_L(udtWorkingParams.GelIndex).Pairs(lngPairIndex).P1
                If eClsPaired(lngUMCIndex) = umcpLightUnique Or eClsPaired(lngUMCIndex) = umcpLightMultiple Then
                    blnPossiblyRemoveUMC = True
                End If
            End If
            
            lngUMCIndexLight = GelP_D_L(udtWorkingParams.GelIndex).Pairs(lngPairIndex).P1
            
            If blnPossiblyRemoveUMC Then
                ' See if any of the members of the light UMC have any database hits
                ' We need to examine the light UMC since database matches are always recorded in the light member, and not in the heavy member
                With GelUMC(udtWorkingParams.GelIndex)
                    If IsAMTReferencedByUMC(.UMCs(lngUMCIndexLight), udtWorkingParams.GelIndex) Then
                        If Not blnRemoveUMC(lngUMCIndex) Then
                            blnRemoveUMC(lngUMCIndex) = True
                            lngUMCCountToRemove = lngUMCCountToRemove + 1
                            Exit For
                        End If
                    End If
                End With
                
            End If
        Next lngPairIndex
    End If
    
            
    strMessage = "Removed members of pairs with DB hits; LC-MS Features removed = " & Trim(lngUMCCountToRemove)
    If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsRemoveHeavy Then
        strMessageSuffix = "; Removed LC-MS Features that were heavy members of pairs with DB hits"
    Else
        strMessageSuffix = "; Removed LC-MS Features that were light members of pairs with DB hits"
    End If
            
    If lngUMCCountToRemove <= 0 Then
        AddToAnalysisHistory udtWorkingParams.GelIndex, strMessage & strMessageSuffix
    Else
        ' Remove the LC-MS Features
        With GelUMC(udtWorkingParams.GelIndex)
        
            If lngUMCCountToRemove < .UMCCnt Then
                lngNewUMCCount = 0
                For lngUMCIndex = 0 To .UMCCnt - 1
                    If Not blnRemoveUMC(lngUMCIndex) Then
                        lngNewUMCIndex(lngUMCIndex) = lngNewUMCCount
                        .UMCs(lngNewUMCCount) = .UMCs(lngUMCIndex)
                        lngNewUMCCount = lngNewUMCCount + 1
                    End If
                Next lngUMCIndex
            
                If .UMCCnt <> lngNewUMCCount Then
                    .UMCCnt = lngNewUMCCount
                    ReDim Preserve .UMCs(.UMCCnt - 1)
                Else
                    ' This code shouldn't be reached
                    Debug.Assert False
                End If
            
            Else
                .UMCCnt = 0
                ReDim .UMCs(0)
            End If
            
        End With

        strMessage = strMessage & "; New LC-MS Feature count = " & Trim(lngNewUMCCount)
        AddToAnalysisHistory udtWorkingParams.GelIndex, strMessage & strMessageSuffix
                    
        ' Need to recompute the UMC Statistic arrays and store the updated Class Representative Mass
        UpdateUMCStatArrays udtWorkingParams.GelIndex, False
        
        ' Now that we have removed some LC-MS Features, we need to remove any pairs that used those LC-MS Features
        With GelP_D_L(udtWorkingParams.GelIndex)
            If .PCnt > 0 Then
                lngNewPairCount = 0
                
                For lngPairIndex = 0 To .PCnt - 1
                    
                    If blnRemoveUMC(.Pairs(lngPairIndex).P1) Or blnRemoveUMC(.Pairs(lngPairIndex).P2) Then
                        ' Do not copy this pair
                    Else
                        .Pairs(lngNewPairCount) = .Pairs(lngPairIndex)
                        lngNewPairCount = lngNewPairCount + 1
                    End If
                Next lngPairIndex
                
                strMessage = "Removed pairs that contained deleted LC-MS Features; Pairs removed = " & Trim(.PCnt - lngNewPairCount) & "; New pair count = " & Trim(lngNewPairCount)
                AddToAnalysisHistory udtWorkingParams.GelIndex, strMessage
                
                If .PCnt <> lngNewPairCount Then
                    .PCnt = lngNewPairCount
                    ReDim Preserve .Pairs(.PCnt - 1)
                End If
            End If
        End With
        
    End If

RemovePairMemberHitsCleanup:
    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    Exit Sub

RemovePairMemberHitsErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while removing pair member database hits: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    Resume RemovePairMemberHitsCleanup:

End Sub

Private Sub AutoAnalysisSaveChromatograms(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
    
    Dim strFilePath As String
    Dim strGraphicExtension As String
    Dim strMessage As String
    
    Dim lngError As Long
    Dim lngWidthTwips As Long, lngHeightTwips As Long
    Dim eChromType As tbcTICAndBPIConstants
    Dim blnProceed As Boolean
    Dim blnNormalizeYAxis As Boolean, blnSmoothUsingMovingAverage As Boolean
    Dim blnSaveAsPNG As Boolean
    Dim blnPlotNETOnXAxisSaved As Boolean
    
On Error GoTo SaveChromatogramsErrorHandler
    
    lngError = 0
    With glbPreferencesExpanded.AutoAnalysisOptions
        If (.SavePlotTIC Or .SavePlotBPI Or .SavePlotTICTimeDomain Or _
            .SavePlotTICDataPointCounts Or .SavePlotTICDataPointCountsHitsOnly Or _
            .SavePlotTICFromRawData Or .SavePlotBPIFromRawData Or _
            .SavePlotDeisotopingIntensityThresholds Or .SavePlotDeisotopingPeakCounts) And Not .DoNotSaveOrExport Then
            
            ' Need to show all ions (passing filters), not just the hits
            ' Probably only have hits visible, due to code in AutoAnalysisSavePictureGraphic
            ' Change the view to only show the ions with hits
            GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 0) = True
            GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 1) = 0               ' Set to 0 to show all data
            
            AutoAnalysisFilterData udtWorkingParams, udtAutoParams, False
            
            ' Make sure Y Axis normalization is disabled when saving the data to the text file
            blnNormalizeYAxis = glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis
            If blnNormalizeYAxis Then
                ' Temporarily disable Y Axis normalizing
                glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis = False
            End If
            
            ' Make sure smoothing is disabled when saving the data to the text file
            blnSmoothUsingMovingAverage = glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage
            If blnSmoothUsingMovingAverage Then
                ' Temporarily disable smoothing
                glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage = False
            End If
            
            frmTICAndBPIPlots.CallerID = udtWorkingParams.GelIndex
            frmTICAndBPIPlots.Show vbModeless
            frmTICAndBPIPlots.InitializeForm
            frmTICAndBPIPlots.TogglePlotUpdateTimerEnabled False
            
            Select Case .SaveErrorGraphicFileType
            Case pftPictureFileTypeConstants.pftPNG
                strGraphicExtension = ".png"
                blnSaveAsPNG = True
            Case pftPictureFileTypeConstants.pftJPG
                strGraphicExtension = ".jpg"
                blnSaveAsPNG = False
            Case Else
                strGraphicExtension = ".png"
                blnSaveAsPNG = True
                If Not (udtWorkingParams.WarningBits And PICTURE_FORMAT_WARNING_BIT) Then
                    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Unknown error picture type specified; valid options are " & GetErrorGraphicsTypeList()
                    udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or PICTURE_FORMAT_WARNING_BIT
                End If
            End Select
            
            ' Save a text file of the chromatogram data
            strFilePath = udtWorkingParams.ResultsFileNameBase & "_Chromatograms.txt"
            strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
            lngError = frmTICAndBPIPlots.ExportChromatogramDataToClipboardOrFile(strFilePath, True, False)
            
            ' Possibly re-enable y-axis normalization
            If blnNormalizeYAxis <> glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis Then
                frmTICAndBPIPlots.SetNormalizeYAxisOption blnNormalizeYAxis
            End If
            
            ' Possibly re-enable smoothing
            If blnSmoothUsingMovingAverage <> glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage Then
                frmTICAndBPIPlots.SetSmoothUsingMovingAverage blnSmoothUsingMovingAverage
            End If
            
            frmTICAndBPIPlots.ForceRecomputeChromatogram
            
            ' Restore the window size
            ValidateValueLng .SaveErrorGraphSizeWidthPixels, 64, 10000, 1024
            ValidateValueLng .SaveErrorGraphSizeHeightPixels, 64, 10000, 768
            
            ' The user may have minimized the window during auto analysis
            If frmTICAndBPIPlots.WindowState <> vbNormal Then
                frmTICAndBPIPlots.WindowState = vbNormal
            End If
            
            lngWidthTwips = .SaveErrorGraphSizeWidthPixels * Screen.TwipsPerPixelX
            lngHeightTwips = .SaveErrorGraphSizeHeightPixels * Screen.TwipsPerPixelY * 1.2
            
            frmTICAndBPIPlots.ScaleMode = vbTwips
            frmTICAndBPIPlots.width = lngWidthTwips
            frmTICAndBPIPlots.Height = lngHeightTwips
            DoEvents
            
            For eChromType = 0 To TIC_AND_BPI_TYPE_COUNT - 1
                If lngError <> 0 Then Exit For

                blnProceed = False
                Select Case eChromType
                Case tbcTICFromCurrentDataIntensities
                    If .SavePlotTIC Then blnProceed = True
                Case tbcBPIFromCurrentDataIntensities
                    If .SavePlotBPI Then blnProceed = True
                Case tbcTICFromTimeDomain
                    If .SavePlotTICTimeDomain Then blnProceed = True
                Case tbcTICFromCurrentDataPointCounts
                    If .SavePlotTICDataPointCounts Then blnProceed = True
                Case tbcTICFromRawData
                    If .SavePlotTICFromRawData Then blnProceed = True
                Case tbcBPIFromRawData
                    If .SavePlotBPIFromRawData Then blnProceed = True
                Case tbcDeisotopingIntensityThresholds
                    If .SavePlotDeisotopingIntensityThresholds Then blnProceed = True
                Case tbcDeisotopingPeakCounts
                    If .SavePlotDeisotopingPeakCounts Then blnProceed = True
                Case Else
                    ' This shouldn't happen
                    Debug.Assert False
                End Select
                
                If blnProceed Then
                    ' Note: Since tmrUpdatePlot was disabled above, control will not return to this function until after the plot gets updated
                    frmTICAndBPIPlots.SetPlotMode eChromType
                    
                    strFilePath = udtWorkingParams.ResultsFileNameBase & "_" & frmTICAndBPIPlots.GetChromDescription(eChromType, False)
                    If eChromType = tbcTICFromCurrentDataIntensities Or eChromType = tbcBPIFromCurrentDataIntensities Then
                        If glbPreferencesExpanded.TICAndBPIPlottingOptions.PlotNETOnXAxis Then
                            strFilePath = strFilePath & "_NET"
                        Else
                            strFilePath = strFilePath & "_Scan"
                        End If
                    End If
                    strFilePath = strFilePath & strGraphicExtension
                    strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                    
                    lngError = frmTICAndBPIPlots.SaveChartPictureToFile(blnSaveAsPNG, strFilePath, False)
                
                    Select Case eChromType
                    Case tbcTICFromCurrentDataIntensities
                        AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), "Total Ion Chromatogram (TIC)", udtWorkingParams.TICPlotsStartRow, 1
                    Case tbcDeisotopingIntensityThresholds
                        AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), "Deisotoping Intensity Thresholds", udtWorkingParams.TICPlotsStartRow + 1, 1
                    Case tbcDeisotopingPeakCounts
                        AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), "Deisotoping Peak Counts", udtWorkingParams.TICPlotsStartRow + 1, 2
                    Case Else
                        ' Do not add to the Html file
                    End Select
                    
                    
                    If eChromType = tbcTICFromCurrentDataIntensities Or eChromType = tbcBPIFromCurrentDataIntensities Then
                        ' Also save the TIC and BPI vs. GANET
                        ' (or if we just saved vs. GANET, save vs. scan)
                        
                        blnPlotNETOnXAxisSaved = glbPreferencesExpanded.TICAndBPIPlottingOptions.PlotNETOnXAxis
                        frmTICAndBPIPlots.SetPlotNETOnXAxis Not blnPlotNETOnXAxisSaved
                        frmTICAndBPIPlots.ForceUpdatePlot
                        
                        strFilePath = udtWorkingParams.ResultsFileNameBase & "_" & frmTICAndBPIPlots.GetChromDescription(eChromType, False)
                        If Not blnPlotNETOnXAxisSaved Then
                            strFilePath = strFilePath & "_NET"
                        Else
                            strFilePath = strFilePath & "_Scan"
                        End If
                        strFilePath = strFilePath & strGraphicExtension
                        strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                        
                        lngError = frmTICAndBPIPlots.SaveChartPictureToFile(blnSaveAsPNG, strFilePath, False)
                        
                        If eChromType = tbcBPIFromCurrentDataIntensities Then
                            AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), "Base Peak Intensity (BPI) Chromatogram", udtWorkingParams.TICPlotsStartRow, 2
                        End If
                        
                        frmTICAndBPIPlots.SetPlotNETOnXAxis blnPlotNETOnXAxisSaved
                        frmTICAndBPIPlots.ForceUpdatePlot
                        glbPreferencesExpanded.TICAndBPIPlottingOptions.PlotNETOnXAxis = blnPlotNETOnXAxisSaved
                    
                    End If
                
                End If
            Next eChromType
            
            
            If .SavePlotTICDataPointCountsHitsOnly Then
                ' Need to save a TIC constructed from the data point intensities, using only the DB hits
                ' Also save a TIC constructed from the data point counts, using only the DB hits
                
                ' Change the view to only show the ions with hits
                GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 0) = True
                GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 1) = 2               ' Set to 2 to exclude unidentified
                
                ' Assign udtWorkingParams.GelIndex to frmFilter.Tag, then call .InitializeControls
                ' Since glbPreferencesExpanded.AutoAnalysisStatus.Enabled = True, will
                '  automatically unload the form after the filter is applied
                frmFilter.Tag = udtWorkingParams.GelIndex
                frmFilter.InitializeControls True

                ' Recompute the chromatogram, using only the data in view
                frmTICAndBPIPlots.ForceRecomputeChromatogram
                
                ' Make sure the plot shows the desired TIC (use current data point intensities)
                eChromType = tbcTICFromCurrentDataIntensities
                frmTICAndBPIPlots.SetPlotMode eChromType
                
                ' Save to disk
                strFilePath = udtWorkingParams.ResultsFileNameBase & "_" & frmTICAndBPIPlots.GetChromDescription(eChromType, False) & "DBHitsOnly"
                strFilePath = strFilePath & strGraphicExtension
                strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                
                lngError = frmTICAndBPIPlots.SaveChartPictureToFile(blnSaveAsPNG, strFilePath, False)
            
            
                ' Make sure the plot shows the desired TIC (use current data point counts)
                eChromType = tbcTICFromCurrentDataPointCounts
                frmTICAndBPIPlots.SetPlotMode eChromType
                
                ' Save to disk
                strFilePath = udtWorkingParams.ResultsFileNameBase & "_" & frmTICAndBPIPlots.GetChromDescription(eChromType, False) & "DBHitsOnly"
                strFilePath = strFilePath & strGraphicExtension
                strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                
                lngError = frmTICAndBPIPlots.SaveChartPictureToFile(blnSaveAsPNG, strFilePath, False)
            
            
                ' Save the chromatograms constructed using only the data in view to a text file
                ' Make sure Y Axis normalization is disabled
                blnNormalizeYAxis = glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis
                If blnNormalizeYAxis Then
                    ' Temporarily disable Y Axis normalizing
                    glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis = False
                End If
                
                ' Make sure smoothing is disabled
                blnSmoothUsingMovingAverage = glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage
                If blnSmoothUsingMovingAverage Then
                    ' Temporarily disable smoothing
                    glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage = False
                End If
            
                ' Recompute the chromatogram
                frmTICAndBPIPlots.ForceRecomputeChromatogram
            
                ' Save a text file of the chromatogram data
                strFilePath = udtWorkingParams.ResultsFileNameBase & "_ChromatogramsDBHitsOnly.txt"
                strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                lngError = frmTICAndBPIPlots.ExportChromatogramDataToClipboardOrFile(strFilePath, False, False, True)
                
                ' Possibly re-enable y-axis normalization
                If blnNormalizeYAxis <> glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis Then
                    glbPreferencesExpanded.TICAndBPIPlottingOptions.NormalizeYAxis = blnNormalizeYAxis
                    frmTICAndBPIPlots.SetNormalizeYAxisOption blnNormalizeYAxis
                End If
                
                ' Possibly re-enable smoothing
                If blnSmoothUsingMovingAverage <> glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage Then
                    glbPreferencesExpanded.TICAndBPIPlottingOptions.SmoothUsingMovingAverage = blnSmoothUsingMovingAverage
                    frmTICAndBPIPlots.SetSmoothUsingMovingAverage blnSmoothUsingMovingAverage
                End If
            
            End If
            
            
            If lngError <> 0 Then
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - Unable to save a graphic of the chromatogram plots to disk (" & strFilePath & "): " & Error(lngError)
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SAVE_CHROMATOGRAM_ERROR_BIT
            End If
        
            Unload frmTICAndBPIPlots
        End If
    End With

    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    Exit Sub

SaveChromatogramsErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while saving chromatogram plots during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

Private Sub AutoAnalysisSaveErrorDistributions(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject, ByVal blnSavingDataDuringToleranceRefinement As Boolean, ByVal blnUpdateCachedErrorPeakStats As Boolean, ByVal eErrorHistogramMode As ehmErrorHistogramModeConstants, ByVal intHtmlFileColumnOverride As Integer)
    ' Note: blnUpdateCachedErrorPeakStats is only valid if blnSavingDataDuringToleranceRefinement = False
    ' Note: strFileNameSuffix is only valid if blnSavingDataDuringToleranceRefinement = True
    ' If intHtmlFileColumnOverride is <> 0, then will save the graphic in the given column
    
    Const SAVE_FINAL_TOLERANCE_PLOTS As Boolean = True
    
    Dim strFilePath As String
    Dim strGraphicExtension As String
    Dim strMessage As String
    Dim strFileNameSuffix As String
    
    Dim lngError As Long
    Dim lngWidthTwips As Long, lngHeightTwips As Long
    
    Dim intRowToUse As Integer, intColumnToUse As Integer
    Dim strFigureCaption As String
    
On Error GoTo SaveErrorGraphicErrorHandler
    
    If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount < 1 Then
        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Saving of the error distributions was requested, but no database searching was performed; skipping this step"
        Exit Sub
    End If
    
    Select Case eErrorHistogramMode
    Case ehmErrorHistogramModeConstants.ehmBeforeRefinement
        strFileNameSuffix = "_BeforeRefinement"
    Case ehmErrorHistogramModeConstants.ehmAfterLCMSWARP
        strFileNameSuffix = "_AfterLCMSWARP"
    Case ehmErrorHistogramModeConstants.ehmFinalTolerances
        strFileNameSuffix = ""
    Case Else
        ' Unknown mode
        Debug.Assert False
        strFileNameSuffix = ""
    End Select
    
    lngError = 0
    With glbPreferencesExpanded.AutoAnalysisOptions
        If (.SaveErrorGraphicMass Or .SaveErrorGraphicGANET Or .SaveErrorGraphic3D Or blnSavingDataDuringToleranceRefinement) And _
            Not .DoNotSaveOrExport Then
            
            If Not blnSavingDataDuringToleranceRefinement Then
                frmErrorDistribution2DLoadedData.CallerID = udtWorkingParams.GelIndex
                frmErrorDistribution2DLoadedData.Show vbModeless
                frmErrorDistribution2DLoadedData.InitializeForm
            
                If blnUpdateCachedErrorPeakStats Then
                    frmErrorDistribution2DLoadedData.RecordMassCalPeakStatsNow
                    frmErrorDistribution2DLoadedData.RecordNETTolPeakStatsNow
                End If
            Else
                ' The form was already loaded and initialized by the calling function
            End If
            
            Select Case .SaveErrorGraphicFileType
            Case pftPictureFileTypeConstants.pftPNG
                strGraphicExtension = ".png"
            Case pftPictureFileTypeConstants.pftJPG
                strGraphicExtension = ".jpg"
            Case Else
                strGraphicExtension = ".png"
                If Not (udtWorkingParams.WarningBits And PICTURE_FORMAT_WARNING_BIT) Then
                    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Unknown error picture type specified; valid options are " & GetErrorGraphicsTypeList()
                    udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or PICTURE_FORMAT_WARNING_BIT
                End If
            End Select
            
            strFilePath = udtWorkingParams.ResultsFileNameBase & "_MassAndGANETErrors" & strFileNameSuffix & ".txt"
            strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
            lngError = frmErrorDistribution2DLoadedData.ExportErrorsBinnedToClipboardOrFile(strFilePath, False, False)
            If lngError = -1 Then
                ' No data in memory; nothing to graph
                lngError = 0
            End If
            
            ValidateValueLng .SaveErrorGraphSizeWidthPixels, 64, 10000, 800
            ValidateValueLng .SaveErrorGraphSizeHeightPixels, 64, 10000, 600
            
            ' The user may have minimized the window during auto analysis
            If frmErrorDistribution2DLoadedData.WindowState <> vbNormal Then
                frmErrorDistribution2DLoadedData.WindowState = vbNormal
            End If
            
            lngWidthTwips = .SaveErrorGraphSizeWidthPixels * Screen.TwipsPerPixelX
            lngHeightTwips = .SaveErrorGraphSizeHeightPixels * Screen.TwipsPerPixelY
            
            ' Make sure the tolerance refinement controls are hidden
            frmErrorDistribution2DLoadedData.ShowHideToleranceRefinementControls False
            
            ' Set the form to the desired size
            frmErrorDistribution2DLoadedData.ScaleMode = vbTwips
            frmErrorDistribution2DLoadedData.width = lngWidthTwips
            frmErrorDistribution2DLoadedData.Height = lngHeightTwips
            DoEvents
            
            If (.SaveErrorGraphicMass Or blnSavingDataDuringToleranceRefinement) And lngError = 0 Then
                frmErrorDistribution2DLoadedData.SetPlotMode mdmMassErrorPPM
                strFilePath = udtWorkingParams.ResultsFileNameBase & "_MassErrors" & strFileNameSuffix & strGraphicExtension
                strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                If .SaveErrorGraphicFileType = pftPictureFileTypeConstants.pftJPG Then
                    lngError = frmErrorDistribution2DLoadedData.SaveChartPictureToFile(False, strFilePath, False)
                Else
                    lngError = frmErrorDistribution2DLoadedData.SaveChartPictureToFile(True, strFilePath, False)
                End If
                
                intRowToUse = 2
                Select Case eErrorHistogramMode
                Case ehmErrorHistogramModeConstants.ehmBeforeRefinement
                    strFigureCaption = "Mass Errors Before Refinement"
                    intColumnToUse = 1
                Case ehmErrorHistogramModeConstants.ehmAfterLCMSWARP
                    strFigureCaption = "Mass Errors After LCMSWARP"
                    intColumnToUse = 2
                Case ehmErrorHistogramModeConstants.ehmFinalTolerances
                    strFigureCaption = "Mass Errors with Final Tolerances"
                    If SAVE_FINAL_TOLERANCE_PLOTS Then
                        intColumnToUse = 3
                    Else
                        intColumnToUse = 0
                    End If
                Case Else
                    ' Unknown mode
                    Debug.Assert False
                End Select
                        
                If intHtmlFileColumnOverride <> 0 Then
                    intColumnToUse = intHtmlFileColumnOverride
                End If
            
                If intColumnToUse <> 0 Then
                    AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), strFigureCaption, intRowToUse, intColumnToUse
                End If
            End If
            
            If (.SaveErrorGraphicGANET Or blnSavingDataDuringToleranceRefinement) And lngError = 0 Then
                frmErrorDistribution2DLoadedData.SetPlotMode mdmGanetError
                strFilePath = udtWorkingParams.ResultsFileNameBase & "_GANETErrors" & strFileNameSuffix & strGraphicExtension
                strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                If .SaveErrorGraphicFileType = pftPictureFileTypeConstants.pftJPG Then
                    lngError = frmErrorDistribution2DLoadedData.SaveChartPictureToFile(False, strFilePath, False)
                Else
                    lngError = frmErrorDistribution2DLoadedData.SaveChartPictureToFile(True, strFilePath, False)
                End If
            
                intRowToUse = 3
                Select Case eErrorHistogramMode
                Case ehmErrorHistogramModeConstants.ehmBeforeRefinement
                    strFigureCaption = "NET Errors Before Refinement"
                    intColumnToUse = 1
                Case ehmErrorHistogramModeConstants.ehmAfterLCMSWARP
                    strFigureCaption = "NET Errors After LCMSWARP"
                    intColumnToUse = 2
                Case ehmErrorHistogramModeConstants.ehmFinalTolerances
                    strFigureCaption = "NET Errors with Final Tolerances"
                    If SAVE_FINAL_TOLERANCE_PLOTS Then
                        intColumnToUse = 3
                    Else
                        intColumnToUse = 0
                    End If
                Case Else
                    ' Unknown mode
                    Debug.Assert False
                End Select
                        
                If intHtmlFileColumnOverride <> 0 Then
                    intColumnToUse = intHtmlFileColumnOverride
                End If
            
                If intColumnToUse <> 0 Then
                    AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strFilePath), strFigureCaption, intRowToUse, intColumnToUse
                End If
            End If
            
            If Not blnSavingDataDuringToleranceRefinement And .SaveErrorGraphic3D And lngError = 0 Then
                strFilePath = udtWorkingParams.ResultsFileNameBase & "_MassAndGANETErrors3D" & strGraphicExtension
                strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                
                lngError = frmErrorDistribution2DLoadedData.ShowOrCompute3DErrorDistributions(False, False)
                
                If lngError = 0 Then
                    ' Note that the above call should have loaded frmErrorDistribution3DFromFile
                    ' frmChart3D should also now be open, but it can't be directly accessed since it
                    '   was loaded as a local variable of frmErrorDistribution3DFromFile
                    ' Save a picture of the graph by calling .SaveGraphPicture
                    If .SaveErrorGraphicFileType = pftPictureFileTypeConstants.pftJPG Then
                        frmErrorDistribution3DFromFile.SaveGraphPicture False, strFilePath, lngWidthTwips, lngHeightTwips
                    Else
                        frmErrorDistribution3DFromFile.SaveGraphPicture True, strFilePath, lngWidthTwips, lngHeightTwips
                    End If
                ElseIf lngError = -1 Then
                    ' No data in memory; nothing to graph
                    lngError = 0
                End If
                frmErrorDistribution3DFromFile.UnloadMyself
            End If
        
            If lngError <> 0 Then
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - Unable to save a graphic of the error distributions to disk (" & strFilePath & "): " & Error(lngError)
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SAVE_ERROR_DISTRIBUTION_ERROR_BIT
            End If
        
            If Not blnSavingDataDuringToleranceRefinement Then
                Unload frmErrorDistribution2DLoadedData
            End If
        End If
    End With

    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    Exit Sub

SaveErrorGraphicErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while saving error distribution info during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    
End Sub

Private Sub AutoAnalysisSaveInternalStdHits(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
    ' Search for Internal Standard using AutoAnalysisSearchDatabase
    ' Export the matching LC-MS Features to a text file, but not to the database
    ' After searching, use AutoAnalysisSavePictureGraphic to save a graphic of the hits
    
    Dim udtAutoAnalysisOptionsSaved As udtAutoAnalysisOptionsType
    Dim udtAMTDefSaved As SearchAMTDefinition
    Dim strMessage As String

On Error GoTo AutoAnalysisSaveInternalStdHitsErrorHandler

    If Not glbPreferencesExpanded.AutoAnalysisOptions.SaveInternalStdHitsAndData Or _
        glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
        ' Skip searching for Internal Standard
        Exit Sub
    End If
        
    ' 1. Copy settings from .AutoAnalysisOptions to udtAutoAnalysisOptionsSaved
    '    Disable exporting for database searching
    '    Enable exporting to disk
    '    In addition, make sure only one database search mode exists
    udtAutoAnalysisOptionsSaved = glbPreferencesExpanded.AutoAnalysisOptions
    udtAMTDefSaved = samtDef
    With glbPreferencesExpanded.AutoAnalysisOptions
        ' Make sure SearchModeCount is 1
        .AutoAnalysisSearchModeCount = 1
        
        ' Set the search mode to AUTO_SEARCH_UMC_CONGLOMERATE
        With .AutoAnalysisSearchMode(0)
            .SearchMode = AUTO_SEARCH_UMC_CONGLOMERATE
            .InternalStdSearchMode = issmFindOnlyInternalStandards
            .DBSearchMinimumHighNormalizedScore = 0
            .DBSearchMinimumHighDiscriminantScore = 0
            .DBSearchMinimumPeptideProphetProbability = 0
            .ExportResultsToDatabase = False
            .ExportUMCMembers = False
            .WriteResultsToTextFile = True
        End With
        
        ' Copy the MW search tolerance from GelUMCNETAdjDef(CallerID) to samtDef
        samtDef.MWTol = GelUMCNETAdjDef(udtWorkingParams.GelIndex).MWTol
        samtDef.TolType = GelUMCNETAdjDef(udtWorkingParams.GelIndex).MWTolType
        
        ' Use the greater of samtDef.NETTol and .AutoToleranceRefinement.DBSearchNETTol
        With .AutoToleranceRefinement
            If .DBSearchNETTol > samtDef.NETTol Then
                samtDef.NETTol = .DBSearchNETTol
            End If
        End With
        
        ' Copy samtDef to GelSearchDef()
        GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnIons = samtDef
        GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnUMCs = samtDef
        GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnPairs = samtDef
        
    End With
    
    ' 2. Call AutoAnalysisSearchDatabase
    AutoAnalysisSearchDatabase udtWorkingParams, udtAutoParams, fso, True, True
    
    ' 3. Restore samtDef and the other search definitions; also restore .AutoAnalysisOptions
    samtDef = udtAMTDefSaved
    With GelSearchDef(udtWorkingParams.GelIndex)
        .AMTSearchOnIons = samtDef
        .AMTSearchOnUMCs = samtDef
        .AMTSearchOnPairs = samtDef
    End With
    glbPreferencesExpanded.AutoAnalysisOptions = udtAutoAnalysisOptionsSaved
    
    ' 4. Call AutoAnalysisSavePictureGraphic to save a graphic of the hits
    AutoAnalysisSavePictureGraphic udtWorkingParams, udtAutoParams, fso, True
    
    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
    
    Exit Sub
    
AutoAnalysisSaveInternalStdHitsErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred in the Save Internal Standard Hits step during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

Private Sub AutoAnalysisSaveGelFile(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType)
    
    Dim blnSaveGelFile As Boolean
    
    With glbPreferencesExpanded.AutoAnalysisOptions
        ' Save the file if .SaveGelFile = True, or if there is an error and .SaveGelFileOnError = True
        
        If .SaveGelFile Then
            blnSaveGelFile = True
        ElseIf udtWorkingParams.ErrorBits <> 0 And .SaveGelFileOnError Then
            ' Do not save the gel if the only error was that the MT tags had high null counts
            ' Also, do not save the gel if the error was TOLERANCE_REFINEMENT_ERROR_BIT
            ' Also, do not save the gel if the error was SAVE_GRAPHIC_ERROR_BIT
            If udtWorkingParams.ErrorBits <> MASS_TAGS_NULL_COUNTS_HIGH_ERROR_BIT And _
               udtWorkingParams.ErrorBits <> TOLERANCE_REFINEMENT_ERROR_BIT And _
               udtWorkingParams.ErrorBits <> MASS_TAGS_NULL_COUNTS_HIGH_ERROR_BIT + TOLERANCE_REFINEMENT_ERROR_BIT And _
               udtWorkingParams.ErrorBits <> SAVE_GRAPHIC_ERROR_BIT Then
                blnSaveGelFile = True
            End If
        Else
            blnSaveGelFile = False
        End If
        
        ' However, never save if .DoNotSaveOrExport = True
        If .DoNotSaveOrExport Then
            blnSaveGelFile = False
        End If
    End With
    
    If blnSaveGelFile Then
        ' Attempt to save a .Gel file in the same directory as the source .Pek, .CSV, .mzXML, or .mzData file (or in .AlternateOutputFolderPath)
        ' Abort save if a file error, but log to log file
        With udtWorkingParams
            If glbPreferencesExpanded.ExtendedFileSaveModePreferred Then
                SaveFileAs .GelFilePath, False, False, .GelIndex, fsIncludeExtended
            Else
                SaveFileAs .GelFilePath, False, False, .GelIndex, fsNoExtended
            End If
        End With
    End If

End Sub

Private Sub AutoAnalysisSavePictureGraphic(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject, Optional blnSavingInternalStdHits As Boolean = False, Optional blnForceSave As Boolean = False)
    
    Dim strMessage As String
    Dim strGraphicFilePath As String
    Dim strGraphicExtension As String
    Dim eSaveType As pftPictureFileTypeConstants
    Dim lngError As Long
    
    Dim lngGelScanNumberMin As Long, lngGelScanNumberMax As Long
    Dim dblGelMassMin As Double, dblGelMassMax As Double
    
    Dim blnShowZoomed As Boolean

On Error GoTo SavePictureGraphicErrorHandler

    If (glbPreferencesExpanded.AutoAnalysisOptions.SavePictureGraphic And _
       Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport) Or blnForceSave Then
        ' Attempt to save a Picture of the .Gel file in the same directory as the source .Pek, .CSV, .mzXML, or .mzData file (or in .AlternateOutputFolderPath)
        
        ' Make sure we're zoomed out and that we're using Scan Numbers to label the X axis
        AutoAnalysisZoomOut udtWorkingParams, True
        GelBody(udtWorkingParams.GelIndex).SetXAxisLabelType False
        
        With glbPreferencesExpanded.AutoAnalysisOptions
            ' Note: Using LCase to avoid case conversion problems
            Select Case .SavePictureGraphicFileType
            Case pftPictureFileTypeConstants.pftPNG
                strGraphicExtension = ".png"
            Case pftPictureFileTypeConstants.pftJPG
                strGraphicExtension = ".jpg"
            Case pftPictureFileTypeConstants.pftWMF
                strGraphicExtension = ".wmf"
            Case pftPictureFileTypeConstants.pftEMF
                strGraphicExtension = ".emf"
            Case pftPictureFileTypeConstants.pftBMP
                strGraphicExtension = ".bmp"
            Case Else
                .SavePictureGraphicFileType = pftPictureFileTypeConstants.pftPNG
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Unknown Picture type specified; valid options are " & GetPictureGraphicsTypeList()
                strGraphicExtension = ".wmf"
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or PICTURE_FORMAT_WARNING_BIT
            End Select
            
            eSaveType = .SavePictureGraphicFileType
        End With

        With glbPreferencesExpanded.AutoAnalysisOptions
            ' Enlarge the window to the desired size (pixels)
            ValidateValueLng .SavePictureWidthPixels, 64, 10000, 1024
            ValidateValueLng .SavePictureHeightPixels, 64, 10000, 768
            
            ' The user may have minimized the window during auto analysis
            If GelBody(udtWorkingParams.GelIndex).WindowState <> vbNormal Then
                GelBody(udtWorkingParams.GelIndex).WindowState = vbNormal
            End If
            
            GelBody(udtWorkingParams.GelIndex).ScaleMode = vbTwips
            GelBody(udtWorkingParams.GelIndex).width = .SavePictureWidthPixels * Screen.TwipsPerPixelX
            GelBody(udtWorkingParams.GelIndex).Height = .SavePictureHeightPixels * Screen.TwipsPerPixelY
            DoEvents
        End With
        
        If blnSavingInternalStdHits Then
            lngError = 0
        Else
            ' First save the current view, which shows the data that was searched
            strGraphicFilePath = udtWorkingParams.ResultsFileNameBase & "_DataSearched" & strGraphicExtension
            strGraphicFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strGraphicFilePath)
            lngError = SaveFileAsPicture(udtWorkingParams.GelIndex, strGraphicFilePath, eSaveType)
            
            If (eSaveType = pftPictureFileTypeConstants.pftEMF Or eSaveType = pftPictureFileTypeConstants.pftWMF) And lngError = 0 Then
                ' Also save a PNG of the data
                strGraphicFilePath = udtWorkingParams.ResultsFileNameBase & "_DataSearched" & ".png"
                lngError = SaveFileAsPicture(udtWorkingParams.GelIndex, strGraphicFilePath, pftPictureFileTypeConstants.pftPNG)
            End If
            
            ' Next switch to charge-state view and save
            If lngError = 0 Then
                GelBody(udtWorkingParams.GelIndex).ShowChargeStateMap
                strGraphicFilePath = udtWorkingParams.ResultsFileNameBase & "_DataSearchedChargeView" & strGraphicExtension
                strGraphicFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strGraphicFilePath)
                lngError = SaveFileAsPicture(udtWorkingParams.GelIndex, strGraphicFilePath, eSaveType)
                AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strGraphicFilePath), "Data Searched", 1, 1
            End If
        End If
        
        If lngError = 0 Then
            ' Switch back to normal view
            GelBody(udtWorkingParams.GelIndex).ShowNormalView
        
            ' If restricted by Scan Number or by Mass, then zoom in to that region
            With glbPreferencesExpanded.AutoAnalysisFilterPrefs
                If .RestrictIsoByMass Then
                    dblGelMassMin = .RestrictIsoMassMin
                    dblGelMassMax = .RestrictIsoMassMax
                    blnShowZoomed = True
                ElseIf .RestrictCSByMass Then
                    dblGelMassMin = .RestrictCSMassMin
                    dblGelMassMax = .RestrictCSMassMax
                    blnShowZoomed = True
                End If
                
                If .RestrictScanRange Then
                    lngGelScanNumberMin = .RestrictScanRangeMin
                    lngGelScanNumberMax = .RestrictScanRangeMax
                    blnShowZoomed = True
                End If
            End With
            
            If blnShowZoomed Then
                ' Zoom in (note: showing all data, not just hits)
                ZoomGelToDimensions udtWorkingParams.GelIndex, CSng(lngGelScanNumberMin), dblGelMassMin, CSng(lngGelScanNumberMax), dblGelMassMax
                
                If Not blnSavingInternalStdHits Then
                    strGraphicFilePath = udtWorkingParams.ResultsFileNameBase & "_DataSearchedZoomed" & strGraphicExtension
                    strGraphicFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strGraphicFilePath)
                    lngError = SaveFileAsPicture(udtWorkingParams.GelIndex, strGraphicFilePath, eSaveType)
                End If
            End If
            
            ' Now change the view to only show the ions with hits
            GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 0) = True
            GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 1) = 2               ' Set to 2 to exclude unidentified
            
            ' Assign udtWorkingParams.GelIndex to frmFilter.Tag, then call .InitializeControls
            ' Since glbPreferencesExpanded.AutoAnalysisStatus.Enabled = True, will
            '  automatically unload the form after the filter is applied
            frmFilter.Tag = udtWorkingParams.GelIndex
            frmFilter.InitializeControls True
            
            If Not blnSavingInternalStdHits Then
                strGraphicFilePath = udtWorkingParams.ResultsFileNameBase & "_DataWithHits" & strGraphicExtension
                strGraphicFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strGraphicFilePath)
                lngError = SaveFileAsPicture(udtWorkingParams.GelIndex, strGraphicFilePath, eSaveType)
            End If
            
            ' Switch from Scan number view to NET view, and save
            GelBody(udtWorkingParams.GelIndex).SetXAxisLabelType True
            
            If blnSavingInternalStdHits Then
                strGraphicFilePath = udtWorkingParams.ResultsFileNameBase & "_DataWithHitsInternalStds" & strGraphicExtension
            Else
                strGraphicFilePath = udtWorkingParams.ResultsFileNameBase & "_DataWithHitsNET" & strGraphicExtension
                AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strGraphicFilePath), "Data With Matches", 1, 2
            End If
            
            strGraphicFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strGraphicFilePath)
            lngError = SaveFileAsPicture(udtWorkingParams.GelIndex, strGraphicFilePath, eSaveType)
        
            GelBody(udtWorkingParams.GelIndex).SetXAxisLabelType False
            
            If blnShowZoomed Then
                ' Make sure the scan range is zoomed out full and the mass range is zoomed out to at least the minimum mass values
                AutoAnalysisZoomOut udtWorkingParams, False
            End If
        End If
        
        If lngError <> 0 Then
            AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - Unable to save a graphic of the current view to disk (" & strGraphicFilePath & "): " & Error(lngError)
            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SAVE_GRAPHIC_ERROR_BIT
        End If
        
        ' Decrease the window to a reasonable size
        GelBody(udtWorkingParams.GelIndex).ScaleMode = vbTwips
        GelBody(udtWorkingParams.GelIndex).width = 640 * Screen.TwipsPerPixelX
        GelBody(udtWorkingParams.GelIndex).Height = 480 * Screen.TwipsPerPixelY
        DoEvents
        
    End If

    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    Exit Sub
    
SavePictureGraphicErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while saving a picture of the analysis during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

Private Sub AutoAnalysisSaveUMCsToDisk(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
        
    Dim strOutputFilePath As String
    Dim strMessage As String

On Error GoTo SaveUMCsToDiskErrorHandler
    
    If Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
        If glbPreferencesExpanded.AutoAnalysisOptions.SaveUMCStatisticsToTextFile Then
            
            ' Make sure the gel is zoomed out and remove the filters to make sure all of the LC-MS Features are shown
            AutoAnalysisZoomOut udtWorkingParams, True
            
            strOutputFilePath = udtWorkingParams.ResultsFileNameBase & "_UMCs.txt"
            
            With glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(0)
                If FolderExists(.AlternateOutputFolderPath) Then
                    strOutputFilePath = fso.BuildPath(.AlternateOutputFolderPath, strOutputFilePath)
                Else
                    strOutputFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strOutputFilePath)
                End If
            End With
            
            GelBody(udtWorkingParams.GelIndex).CopyAllUMCsInView -1, False, strOutputFilePath
        End If
    End If
    
    Exit Sub
    
SaveUMCsToDiskErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while saving LC-MS Features to disk during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    
End Sub

Private Sub AutoAnalysisSearchDatabase(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject, Optional blnToleranceRefinementSearch As Boolean = False, Optional blnInternalStdOnlySearch As Boolean = False)
    
    Dim dblGANETSlope As Double, dblGANETIntercept As Double
    Dim dblGelMassMin As Double, dblGelMassMax As Double
    
    Dim strMessage As String
    Dim strResultsFileName As String, strResultsFilePath As String
    Dim strPairsFilePath As String
    Dim strErrorDescription As String
    Dim strIniFileName As String
    
    Dim eDBSearchMode As dbsmDatabaseSearchModeConstants
    Dim eInternalStdSearchMode As issmInternalStandardSearchModeConstants
    Dim sngMTMinimumHighNormalizedScore As Single
    Dim sngMTMinimumHighDiscriminantScore As Single
    Dim sngMTMinimumPeptideProphetProbability As Single
    
    Dim lngHitCount As Long
    Dim lngExportDataToDBErrorCode As Long, lngExportDataToDiskErrorCode As Long
    Dim lngGelScanNumberMin As Long, lngGelScanNumberMax As Long
    Dim lngCrLfLoc As Long
    
    Dim intAutoSearchIndex As Integer
    Dim intSearchIndexCheck As Integer
    Dim intLastPairsBasedSearchModeIndex As Integer
    
    Dim blnExportedDataToDB As Boolean
    Dim blnIdentifiedPairsOnly As Boolean, blnShowExcludedPairs As Boolean
    Dim blnExportNonMatchingUMCsSaved As Boolean, blnAddQuantitationEntrySaved As Boolean
    
On Error GoTo SearchDatabaseErrorHandler

    If GelAnalysis(udtWorkingParams.GelIndex) Is Nothing Then
        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - The GelAnalysis() object has not been initialized; unable to search the database"
        udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
    Else
        With GelAnalysis(udtWorkingParams.GelIndex)
            dblGANETSlope = .GANET_Slope
            dblGANETIntercept = .GANET_Intercept
        End With
    
        ' If requested, Zoom in the gel to the necessary scan range to give the desired GANET range
        ' However, if .RestrictScanRange = True, then make sure the computed scan values are not outside this range
        With glbPreferencesExpanded.AutoAnalysisFilterPrefs
            If .RestrictGANETRange Then
                
                If Not GelData(udtWorkingParams.GelIndex).CustomNETsDefined And dblGANETSlope = 0 Then
                    AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - Unable to restrict to a NET range since the NET slope is 0"
                    udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or GANET_ERROR_BIT
                Else
                    ' First validate the GANET values
                    ValidateValueDbl .RestrictGANETRangeMin, -100, 100, -1
                    ValidateValueDbl .RestrictGANETRangeMax, -100, 100, 2
                    
                    lngGelScanNumberMin = GANETToScan(udtWorkingParams.GelIndex, .RestrictGANETRangeMin)
                    lngGelScanNumberMax = GANETToScan(udtWorkingParams.GelIndex, .RestrictGANETRangeMax)
                    
                    If lngGelScanNumberMin < 0 Then lngGelScanNumberMin = 0
                    If lngGelScanNumberMax < lngGelScanNumberMin Then lngGelScanNumberMax = lngGelScanNumberMin + 1
                    
                    ' Determine the current mass range
                    GetMassRangeCurrent udtWorkingParams.GelIndex, dblGelMassMin, dblGelMassMax
                    
                    If .RestrictScanRange Then
                        ' Make sure the computed scan values aren't out of range
                        ValidateValueLng lngGelScanNumberMin, .RestrictScanRangeMin, .RestrictScanRangeMax, .RestrictScanRangeMin
                        ValidateValueLng lngGelScanNumberMax, .RestrictScanRangeMin, .RestrictScanRangeMax, .RestrictScanRangeMax
                    End If
                    
                    ' Zoom to the desired dimensions
                    ZoomGelToDimensions udtWorkingParams.GelIndex, CSng(lngGelScanNumberMin), dblGelMassMin, CSng(lngGelScanNumberMax), dblGelMassMax
                    
                    AddToAnalysisHistory udtWorkingParams.GelIndex, "Limiting database search to desired NET range; min NET = " & .RestrictGANETRangeMin & "; max NET = " & .RestrictGANETRangeMax & "; min Scan = " & lngGelScanNumberMin & "; max Scan = " & lngGelScanNumberMax
                End If
            End If
        End With
        
        ' Make sure .SkipReferenced = False and .AMTSearchResultsBehavior = asrbAutoRemoveExisting
        GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnIons.SkipReferenced = False
        GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnUMCs.SkipReferenced = False
        GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnPairs.SkipReferenced = False
        samtDef.SkipReferenced = False
        
        glbPreferencesExpanded.AMTSearchResultsBehavior = asrbAutoRemoveExisting
        
        With glbPreferencesExpanded.PairSearchOptions
            intLastPairsBasedSearchModeIndex = -1
            If .AutoAnalysisRemovePairMemberHitsAfterDBSearch And Not blnToleranceRefinementSearch And Not blnInternalStdOnlySearch Then
                ' Need to find the index of the last pairs-based search mode
                
                With glbPreferencesExpanded.AutoAnalysisOptions
                    If .AutoAnalysisSearchModeCount > 0 Then
                        ' Find the last pairs-based search mode
                        intLastPairsBasedSearchModeIndex = -1
                        For intSearchIndexCheck = .AutoAnalysisSearchModeCount - 1 To 0 Step -1
                            If LookupDBSearchModeIndex(.AutoAnalysisSearchMode(intSearchIndexCheck).SearchMode) >= DB_SEARCH_MODE_PAIR_MODE_START_INDEX Then
                                ' Match found
                                intLastPairsBasedSearchModeIndex = intSearchIndexCheck
                                Exit For
                            End If
                        Next intSearchIndexCheck
                        
                        strMessage = "Error - The 'Pair member removal after database search' option is enabled, but "
                        If intLastPairsBasedSearchModeIndex < 0 Then
                            ' Add an error to the log that no pairs-based search mode is defined
                            strMessage = strMessage & "no pairs-based search modes are defined."
                            AddToAnalysisHistory udtWorkingParams.GelIndex, strMessage
                            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or PAIRS_BASED_DB_SEARCH_ERROR_BIT
                            
                        ElseIf intLastPairsBasedSearchModeIndex = .AutoAnalysisSearchModeCount - 1 Then
                            ' Add an error to the log that the .AutoAnalysisRemovePairMemberHitsAfterDBSearch option is meaningless
                            strMessage = strMessage & "no non pairs-based search modes are present after the last pairs-based search."
                            AddToAnalysisHistory udtWorkingParams.GelIndex, strMessage
                            udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or PAIRS_BASED_DB_SEARCH_ERROR_BIT
                        End If
                    End If
                End With
            End If
        End With
        
        For intAutoSearchIndex = 0 To glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount - 1
            
            With glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex)
                If blnInternalStdOnlySearch Then
                    strResultsFileName = udtWorkingParams.ResultsFileNameBase & "_InternalStdits.txt"
                Else
                    strResultsFileName = udtWorkingParams.ResultsFileNameBase & "_" & .SearchMode & ".txt"
                End If
                
                If FolderExists(.AlternateOutputFolderPath) Then
                    strResultsFilePath = fso.BuildPath(.AlternateOutputFolderPath, strResultsFileName)
                Else
                    strResultsFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strResultsFileName)
                End If
                
                eInternalStdSearchMode = .InternalStdSearchMode
                sngMTMinimumHighNormalizedScore = .DBSearchMinimumHighNormalizedScore
                sngMTMinimumHighDiscriminantScore = .DBSearchMinimumHighDiscriminantScore
                sngMTMinimumPeptideProphetProbability = .DBSearchMinimumPeptideProphetProbability
                
                ' Copy this search index's mass mods to GelSearchDef()
                GelSearchDef(udtWorkingParams.GelIndex).AMTSearchMassMods = .MassMods
            End With
            
            strIniFileName = fso.GetFileName(udtAutoParams.FilePaths.IniFilePath)
            
            ' Update the text in MD_Parameters
            GelAnalysis(udtWorkingParams.GelIndex).MD_Parameters = ConstructAnalysisParametersText(udtWorkingParams.GelIndex, glbPreferencesExpanded.AutoAnalysisOptions.UMCSearchMode, glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).SearchMode, strIniFileName)
            
            ' Remove any existing hits from the ions
            RemoveAMT udtWorkingParams.GelIndex, glScope.glSc_All
            RemoveInternalStd udtWorkingParams.GelIndex, glScope.glSc_All
            
            ' Reset the Export Data Error Codes
            lngExportDataToDiskErrorCode = 0
            lngExportDataToDBErrorCode = 0
            
            ' Disable ExportResultsToDatabase if udtAutoParams.InvalidExportPassword = True or if APP_BUILD_DISABLE_MTS = True
            If udtAutoParams.InvalidExportPassword Or APP_BUILD_DISABLE_MTS Then
                With glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex)
                    If .ExportResultsToDatabase Then
                        .ExportResultsToDatabase = False
                        .WriteResultsToTextFile = True
                    End If
                End With
            End If
            
            ' Search the Database
            eDBSearchMode = LookupDBSearchModeIndex(glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).SearchMode)
            Select Case eDBSearchMode
            Case dbsmNone
                lngHitCount = -1
            Case dbsmIndividualPeaks
                ' IndividualPeaks
                With frmSearchMT
                    .CallerID = udtWorkingParams.GelIndex
                    .Show vbModeless
                    .InitializeDBSearch
                    lngHitCount = .StartSearch(False)
                    
                    ' Possibly export the results to a text file
                    If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).WriteResultsToTextFile Then
                        ' Note: This export function lists all of the MT tags in memory,
                        '       and any hits found for the MT tags
                        lngExportDataToDiskErrorCode = .ShowOrSaveResults(strResultsFilePath, False)
                    End If
                    
                    If Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
                        ' Possibly export the results to the database
                        If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).ExportResultsToDatabase Then
                            AddToAnalysisHistory udtWorkingParams.GelIndex, "Warning - Exporting of search results from an ion-by-ion search to the database is not enabled; the search results have simply been written to the text file"
                            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or INVALID_EXPORT_OPTION_WARNING_BIT
                        End If
                    End If
                End With
                Unload frmSearchMT
            Case dbsmIndividualPeaksInUMCsWithoutNET
                ' IndividualPeaksInUMCsWithoutNET
                With frmUMCIdentification
                    .CallerID = udtWorkingParams.GelIndex
                    .Show vbModeless
                    lngHitCount = .StartSearch(False)
                    
                    If Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
                        ' Possibly export the results to a text file
                        If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).WriteResultsToTextFile Then
                            .txtExportDestination.Text = strResultsFilePath
                            lngExportDataToDiskErrorCode = .ExportText()
                        End If
                        
                        ' Possibly export the results to the database; no longer supported (September 2004)
                        If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).ExportResultsToDatabase Then
                            AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - Old Database Search Method specified; database export is no longer supported with this method"
                            'strErrorDescription = .ExportMTDB(lngExportDataToDBErrorCode, udtAutoParams.MDID)
                            blnExportedDataToDB = False
                        End If
                    End If
                End With
                Unload frmUMCIdentification
            Case dbsmIndividualPeaksInUMCsWithNET, dbsmIndividualPeaksInUMCsPaired, dbsmIndividualPeaksInUMCsUnpaired
                ' IndividualPeaksInUMCsWithNET (all, paired, or unpaired)
                ' This mode of searching is no longer supported
                ' Use the ConglomerateUMCsWithNET modes instead
                AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - Old Database Search Method specified (no longer supported); valid options are " & GetAutoAnalysisOptionsList()
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SEARCH_ERROR_BIT

            Case dbsmConglomerateUMCsWithNET, dbsmConglomerateUMCsPaired, dbsmConglomerateUMCsUnpaired, dbsmConglomerateUMCsLightPairsPlusUnpaired, dbsmExportUMCsOnly
                ' ConglomerateUMCsWithNET (all, paired, unpaired, paired plus unpaired)
                ' This code also handles ExportUMCsOnly
                With frmSearchMT_ConglomerateUMC
                    .CallerID = udtWorkingParams.GelIndex
                    .Show vbModeless
                    .InitializeSearch
                    
                    .chkUpdateGelDataWithSearchResults = vbChecked
                    .SetInternalStandardSearchMode eInternalStdSearchMode
                    
                    .SetMinimumHighNormalizedScore sngMTMinimumHighNormalizedScore
                    .SetMinimumHighDiscriminantScore sngMTMinimumHighDiscriminantScore
                    .SetMinimumPeptideProphetProbability sngMTMinimumPeptideProphetProbability
                    
                    If blnToleranceRefinementSearch Then
                        .SearchRegionShape = glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchRegionShape
                    Else
                        .SearchRegionShape = glbPreferencesExpanded.AutoAnalysisOptions.DBSearchRegionShape
                    End If
                    
                    If eDBSearchMode = dbsmConglomerateUMCsPaired Then
                        lngHitCount = .StartSearchPaired()
                    ElseIf eDBSearchMode = dbsmConglomerateUMCsUnpaired Then
                        lngHitCount = .StartSearchNonPaired()
                    ElseIf eDBSearchMode = dbsmConglomerateUMCsLightPairsPlusUnpaired Then
                        lngHitCount = .StartSearchLightPairsPlusNonPaired()
                    ElseIf eDBSearchMode = dbsmExportUMCsOnly Then
                        ' Do not search, but need to make sure .ExportUMCsWithNoMatches is True
                        blnExportNonMatchingUMCsSaved = glbPreferencesExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches
                        blnAddQuantitationEntrySaved = glbPreferencesExpanded.AutoAnalysisOptions.AddQuantitationDescriptionEntry
                        
                        ' Make sure ExportUMCs = True and AddQuantitation = False
                        glbPreferencesExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches = True
                        glbPreferencesExpanded.AutoAnalysisOptions.AddQuantitationDescriptionEntry = False
                    Else
                        ' dbsmConglomerateUMCsWithNET
                        lngHitCount = .StartSearchAll()
                    End If
                        
                    If Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
                        If Not blnToleranceRefinementSearch And _
                          (eDBSearchMode = dbsmConglomerateUMCsPaired Or eDBSearchMode = dbsmConglomerateUMCsUnpaired Or eDBSearchMode = dbsmConglomerateUMCsLightPairsPlusUnpaired) Then
                            If glbPreferencesExpanded.PairSearchOptions.AutoExcludeAmbiguous Then
                                ' Find the ambiguous pairs (those containing LC-MS Features shared across several pairs)
                                '  and mark them as Excluded
                                .ExcludeAmbiguousPairsWrapper True
                                
                                If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisSavePairsToTextFile Then
                                    strPairsFilePath = udtWorkingParams.ResultsFileNameBase & "_PairsAmbiguousExcluded.txt"
                                    strPairsFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strPairsFilePath)
                                    ReportDltLblPairsUMCWrapper udtWorkingParams.GelIndex, glPair_All, strPairsFilePath
                                End If
                                
                            End If
                        End If
                        
                        ' Possibly export the results to a text file
                        If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).WriteResultsToTextFile Then
                            lngExportDataToDiskErrorCode = .ShowOrSaveResultsByUMC(strResultsFilePath, False, glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput)
                        End If
                        
                        ' Possibly export the results to the database
                        If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).ExportResultsToDatabase Then
                            strErrorDescription = .ExportMTDBbyUMC(False, glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).ExportUMCMembers, strIniFileName, lngExportDataToDBErrorCode, udtAutoParams.MDID)
                            blnExportedDataToDB = True
                        End If
                    End If
                End With
                Unload frmSearchMT_ConglomerateUMC
            
                If eDBSearchMode = dbsmExportUMCsOnly Then
                    ' Restore .ExportUMCsWithNoMatches and .AddQuantitationDescriptionEntry to the saved values
                    glbPreferencesExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches = blnExportNonMatchingUMCsSaved
                    glbPreferencesExpanded.AutoAnalysisOptions.AddQuantitationDescriptionEntry = blnAddQuantitationEntrySaved
                End If

            Case dbsmPairsN14N15, dbsmPairsN14N15ConglomerateMass
                ' N14/N15 Pairs DB Search
                
                If eDBSearchMode = dbsmPairsN14N15 Then
                    ' Note that the dbsmPairsN14N15 mode of searching is no longer supported
                    ' However, we'll silently use dbsmPairsN14N15ConglomerateMass instead
                    AddToAnalysisHistory udtWorkingParams.GelIndex, "Warning - Old Database Search Method specified (no longer supported: " & AUTO_SEARCH_PAIRS_N14N15 & "); Will automatically switch to " & AUTO_SEARCH_PAIRS_N14N15_CONGLOMERATEMASS & "; valid options are " & GetAutoAnalysisOptionsList()
                End If
                 
                With frmSearchMTPairs
                    .CallerID = udtWorkingParams.GelIndex
                    .Show vbModeless
                        
                    ' Use Conglomerate UMC Mass
                    .InitializeSearch
                    
                    .SetMinimumHighNormalizedScore sngMTMinimumHighNormalizedScore
                    .SetMinimumHighDiscriminantScore sngMTMinimumHighDiscriminantScore
                    .SetMinimumPeptideProphetProbability sngMTMinimumPeptideProphetProbability
                    
                    If blnToleranceRefinementSearch Then
                        .SearchRegionShape = glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.DBSearchRegionShape
                    Else
                        .SearchRegionShape = glbPreferencesExpanded.AutoAnalysisOptions.DBSearchRegionShape
                    End If
                    
                    lngHitCount = .StartSearchPaired(False)
                        
                    If lngHitCount < 0 Then
                        ' Error while searching; probably a programming error
                        ' .StartSearchPaired will have already added the error message to the analysis log
                        udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SEARCH_ERROR_BIT
                    Else
                        If Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
                            If Not blnToleranceRefinementSearch Then
                                If glbPreferencesExpanded.PairSearchOptions.AutoExcludeAmbiguous Then
                                    ' Find the ambiguous pairs (those containing LC-MS Features shared across several pairs)
                                    '  and mark them as Excluded
                                    .ExcludeAmbiguousPairsWrapper True
                                    
                                    If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisSavePairsToTextFile Then
                                        strPairsFilePath = udtWorkingParams.ResultsFileNameBase & "_PairsAmbiguousExcluded.txt"
                                        strPairsFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strPairsFilePath)
                                        ReportDltLblPairsUMCWrapper udtWorkingParams.GelIndex, glPair_All, strPairsFilePath
                                    End If
                                    
                                End If
                            End If
                            
                            ' If .AutoExcludeAmbiguous = True, then do not show the pairs that were excluded because they were ambiguous
                            blnShowExcludedPairs = Not glbPreferencesExpanded.PairSearchOptions.AutoExcludeAmbiguous
                            
                            ' Possibly export the results to a text file
                            If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).WriteResultsToTextFile Then
                                ' If .AutoAnalysisSavePairsToTextFile = True, then only include the identified pairs in the following report (by setting blnIdentifiedPairsOnly = True)
                                ' Otherwise, set blnIdentifiedPairsOnly = False
                                blnIdentifiedPairsOnly = glbPreferencesExpanded.PairSearchOptions.AutoAnalysisSavePairsToTextFile
                                
                                lngExportDataToDiskErrorCode = .ShowOrSavePairsAndIDs(strResultsFilePath, False, blnIdentifiedPairsOnly, False, False, blnShowExcludedPairs, glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput)
                            End If
                            
                            ' Possibly export the results to the database
                            If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).ExportResultsToDatabase Then
                                strErrorDescription = .ExportMTDBbyUMC(False, glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(intAutoSearchIndex).ExportUMCMembers, strIniFileName, lngExportDataToDBErrorCode, udtAutoParams.MDID, blnShowExcludedPairs)
                                blnExportedDataToDB = True
                            End If
                        End If
                    End If
                End With
                Unload frmSearchMTPairs
            Case dbsmPairsICAT
                AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - ICAT Pairs-based searching Method not yet enabled"
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SEARCH_ERROR_BIT
            Case dbsmPairsPEO
                AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - PEO Pairs-based searching Method not yet enabled"
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SEARCH_ERROR_BIT
            Case Else
                AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - Unknown Database Search Method specified; valid options are " & GetAutoAnalysisOptionsList()
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SEARCH_ERROR_BIT
            End Select
            
            If lngHitCount <= 0 And eDBSearchMode <> dbsmExportUMCsOnly Then
                AddToAnalysisHistory udtWorkingParams.GelIndex, "Warning - Database search resulted in 0 hits"
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or NO_DATABASE_HITS_WARNING_BIT
            End If
            
            If lngExportDataToDiskErrorCode <> 0 Then
                AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - Unable to save search results to disk (" & strResultsFilePath & "): " & Error(lngExportDataToDBErrorCode)
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or EXPORTRESULTS_ERROR_BIT
            End If
            
            If lngExportDataToDBErrorCode <> 0 Then
                lngCrLfLoc = InStr(strErrorDescription, vbCrLf)
                If lngCrLfLoc > 0 Then
                    strErrorDescription = Mid(strErrorDescription, lngCrLfLoc + 2)
                Else
                    strErrorDescription = Error(lngExportDataToDBErrorCode)
                End If
                
                AddToAnalysisHistory udtWorkingParams.GelIndex, "Error - Unable to export search results to database: " & strErrorDescription
                udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or EXPORTRESULTS_ERROR_BIT
            End If
            
            If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsAfterDBSearch Then
                If intAutoSearchIndex = intLastPairsBasedSearchModeIndex Then
                    ' Remove the LC-MS Features that are light or heavy members of pairs that had DB hits
                    AutoAnalysisRemovePairMemberHits udtWorkingParams, udtAutoParams
                End If
            End If
            
            ' Write the latest AnalysisHistory information to the log
            AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams
            
        Next intAutoSearchIndex
        
        ' Export the GANET slope and intercept to the T_FTICR_Analysis_Description table in the DB
        If blnExportedDataToDB And Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
            strErrorDescription = ExportGANETtoMTDB(udtWorkingParams.GelIndex, GelAnalysis(udtWorkingParams.GelIndex).GANET_Slope, GelAnalysis(udtWorkingParams.GelIndex).GANET_Intercept, GelAnalysis(udtWorkingParams.GelIndex).GANET_Fit)
            If UCase(Left(strErrorDescription, 5)) = "ERROR" Then
                AddToAnalysisHistory udtWorkingParams.GelIndex, strErrorDescription
                lngExportDataToDBErrorCode = -1
            End If
        End If
        
        If glbPreferencesExpanded.AutoAnalysisOptions.WriteIDResultsByIonToTextFileAfterAutoSearches And _
           Not glbPreferencesExpanded.AutoAnalysisOptions.DoNotSaveOrExport Then
            
            strResultsFileName = udtWorkingParams.ResultsFileNameBase & "_PeaksWithMatches.txt"
            
            With glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(0)
                If FolderExists(.AlternateOutputFolderPath) Then
                    strResultsFilePath = fso.BuildPath(.AlternateOutputFolderPath, strResultsFileName)
                Else
                    strResultsFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strResultsFileName)
                End If
            End With
        
            ' Make sure the gel is zoomed out
            AutoAnalysisZoomOut udtWorkingParams, False
            
            ' Now change the view to only show the ions with hits
            GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 0) = True
            GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 1) = 2               ' Set to 2 to exclude unidentified
            
            ' Assign udtWorkingParams.GelIndex to frmFilter.Tag, then call .InitializeControls
            ' Since glbPreferencesExpanded.AutoAnalysisStatus.Enabled = True, will
            '  automatically unload the form after the filter is applied
            frmFilter.Tag = udtWorkingParams.GelIndex
            frmFilter.InitializeControls True
            
            GelBody(udtWorkingParams.GelIndex).CopyAllPointsInView -1, False, strResultsFilePath
            
        End If
    End If

AutoAnalysisSearchDBCleanup:
    ' Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    Exit Sub

SearchDatabaseErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred while matching data with the database during auto analysis: " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If
    Resume AutoAnalysisSearchDBCleanup:
    
End Sub

Private Sub AutoAnalysisToleranceRefinement(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef fso As FileSystemObject)
    
    Dim blnPerformRefinement As Boolean
    Dim dblMassCalOverride As Double
    
    Dim udtAutoAnalysisOptionsSaved As udtAutoAnalysisOptionsType
    Dim udtAMTDefSaved As SearchAMTDefinition
    Dim sngSLiCSaved As Single
    
    Dim blnSuccess As Boolean, blnValidPeakFound As Boolean
    
    Dim strErrorMessage As String, strMessage As String
    Dim strLastGoodLocation As String
    
    Dim strFilePath As String
    Dim strFilePathB As String
    Dim strCombinedFilePath As String
    Dim strWorkingCombinedFilePath As String
    Dim strUniqueSuffix As String
    
    Dim lngTimerStart As Long
    Dim lngResult As Long
    Dim blnSaveSuccessful As Boolean
    
    Dim blnPeakTooWide As Boolean, blnMassShiftTooLarge As Boolean
    Dim lngWindowWidth As Long
    Dim lngWindowHeight As Long
    Dim lngWindowWidthTwips As Long
    Dim lngWindowHeightTwips As Long
    
    Dim dblRefinedMassTol As Double
    Dim dblRefinedNETTol As Double
    Dim eRefinedTolType As glMassToleranceConstants
    
    Dim lngDefaultMaxMontageWidth As Long
    Dim lngDefaultMaxMontageHeight As Long
    
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
    
    Dim objMSAlign As frmMSAlign
    
    Dim blnSearchMassTolRefinedByMSWarp As Boolean
    Dim blnSearchNETTolRefinedByMSWarp As Boolean

On Error GoTo AutoAnalysisToleranceRefinementErrorHandler

    ' Check if Tolerance Refinement is enabled
    blnPerformRefinement = CheckAutoToleranceRefinementEnabled(glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement, udtAutoParams, udtWorkingParams, True)

    If Not blnPerformRefinement Then Exit Sub

    ' 1a. Copy settings from .AutoAnalysisOptions to udtAutoAnalysisOptionsSaved and
    '    disable saving and exporting for database searching
    '    In addition, make sure only one database search mode exists
    udtAutoAnalysisOptionsSaved = glbPreferencesExpanded.AutoAnalysisOptions
    udtAMTDefSaved = samtDef
    With glbPreferencesExpanded.AutoAnalysisOptions
        ' Make sure SearchModeCount is 1
        If .AutoAnalysisSearchModeCount <> 1 Then
            .AutoAnalysisSearchModeCount = 1
        End If
        
        ' Copy search tolerances from .AutoToleranceRefinement to samtDef and GelSearchDef()
        With .AutoToleranceRefinement
            samtDef.MWTol = .DBSearchMWTol
            samtDef.TolType = .DBSearchTolType
            samtDef.NETTol = .DBSearchNETTol
            GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnIons = samtDef
            GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnUMCs = samtDef
            GelSearchDef(udtWorkingParams.GelIndex).AMTSearchOnPairs = samtDef
        End With
        
        AutoAnalysisValidateToleranceRefinementSearchMode udtWorkingParams, .AutoAnalysisSearchMode(0), udtAutoAnalysisOptionsSaved, 1
       
        .AutoAnalysisSearchMode(0).DBSearchMinimumHighNormalizedScore = .AutoToleranceRefinement.DBSearchMinimumHighNormalizedScore
        .AutoAnalysisSearchMode(0).DBSearchMinimumHighDiscriminantScore = .AutoToleranceRefinement.DBSearchMinimumHighDiscriminantScore
        .AutoAnalysisSearchMode(0).DBSearchMinimumPeptideProphetProbability = .AutoToleranceRefinement.DBSearchMinimumPeptideProphetProbability
        
        .AutoAnalysisSearchMode(0).WriteResultsToTextFile = False
        .AutoAnalysisSearchMode(0).ExportResultsToDatabase = False
    End With
    
    ' 1b. Save the current value for .MinimumSLiC, then force .MinimumSLiC to 0
    With glbPreferencesExpanded.RefineMSDataOptions
        sngSLiCSaved = .MinimumSLiC
        .MinimumSLiC = 0
    End With
    
    If udtWorkingParams.NETDefined Then
        ' 2. Call AutoAnalysisSearchDatabase
        '    We will restore the samtDef values later in this function
        AutoAnalysisSearchDatabase udtWorkingParams, udtAutoParams, fso, True
    End If

    ' Initialize frmErrorDistribution2DLoadedData (since we need it whether or not dtWorkingParams.NETDefined is True)
    With frmErrorDistribution2DLoadedData
        .CallerID = udtWorkingParams.GelIndex
        ' Note: Must use vbModeless to prevent App from waiting for form to close
        .Show vbModeless
        .InitializeForm
    End With
    
    If udtWorkingParams.NETDefined Then
        ' 3. Save the error distributions before tolerance refinement
        AutoAnalysisToleranceRefinementSaveErrorDistributions frmErrorDistribution2DLoadedData, udtAutoParams, udtWorkingParams, fso, True, ehmErrorHistogramModeConstants.ehmBeforeRefinement
    End If
    
    ' 4. Perform mass calibration refinement
    If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibration Then
        dblMassCalOverride = glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibrationOverridePPM
        If dblMassCalOverride <> 0 Then
            ' Forcing a given mass calibration shift
            AddToAnalysisHistory udtWorkingParams.GelIndex, "Note: Forcing a mass calibration shift of " & Trim(dblMassCalOverride) & " ppm"
            blnSuccess = frmErrorDistribution2DLoadedData.ManualRefineMassCalibration(True, dblMassCalOverride)
            blnValidPeakFound = True
        Else
            
Const APPLY_ADDNL_LINEAR_MASS_ADJUSTMENT As Boolean = True

            If APP_BUILD_DISABLE_LCMSWARP Then
                With GelUMCNETAdjDef(udtWorkingParams.GelIndex)
                    If .RobustNETAdjustmentMode <> UMCRobustNETModeConstants.UMCRobustNETIterative Then
                        .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETIterative
                    End If
                End With
            End If

            If GelUMCNETAdjDef(udtWorkingParams.GelIndex).UseRobustNETAdjustment And _
               GelUMCNETAdjDef(udtWorkingParams.GelIndex).RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass And _
               Not APP_BUILD_DISABLE_LCMSWARP Then
               
                ' 4a. Perform NET and Mass Alignment using MS Warp
        
                strLastGoodLocation = "Instantiate objMSAlign"
                Set objMSAlign = New frmMSAlign
                DoEvents
                Sleep 250
                DoEvents
                
                strLastGoodLocation = "Set objMSAlign.CallerID"
                objMSAlign.CallerID = udtWorkingParams.GelIndex
                objMSAlign.RecalibratingMassDuringAutoAnalysis = True
                
                ' --------------------------------------------------------------
                ' Show the form, initialize the search, and start the alignment
                ' --------------------------------------------------------------
                DoEvents
                Sleep 250
                DoEvents
                strLastGoodLocation = "Show form objMSAlign"
                objMSAlign.Show vbModeless

                strLastGoodLocation = "InitializeSearch"
                objMSAlign.InitializeSearch
                objMSAlign.SetPlotPointSize RESIDUALS_PLOT_POINT_SIZE
                        
                strLastGoodLocation = "StartAlignment"
                objMSAlign.StartAlignment
            
                lngTimerStart = Timer()
                Do
                    Sleep 200
                    DoEvents
                    strLastGoodLocation = "Waiting while objMSAlign.MassMatchState = pscRunning (" & CStr(Timer - lngTimerStart) & " seconds elapsed)"
                Loop While objMSAlign.MassMatchState = pscRunning
              
                If objMSAlign.MassMatchState = pscComplete Then
                    strLastGoodLocation = "objMSAlign.MassMatchState now equals pscComplete"
                    
                    udtWorkingParams.NETDefined = True

                    ' -----------------------------------------
                    ' Alignment succeeded
                    ' -----------------------------------------
                    ' ToDo: Determine whether blnMassShiftTooLarge should ever be set to true
                    blnSuccess = True
                    blnMassShiftTooLarge = False
                    
                    blnValidPeakFound = True
                    blnPeakTooWide = False
               
                    ' Need to wait until objMSAlign.LocalGelUpdated = True
                    strLastGoodLocation = "Wait objMSAlign.LocalGelUpdated = True"
                    Do
                        Sleep 50
                    Loop While Not objMSAlign.LocalGelUpdated
                    
                    ' Alignment is completed, the custom NET values have been updated (again),
                    ' the slope and intercept have been defined, and the masses have been updated
                    
                    ' -----------------------------------------
                    ' Save the NET-related plots to disk
                    ' -----------------------------------------
                    
                    strLastGoodLocation = "Call AutoAnalysisSaveNETSurfaceAndResidualPlots"
                    AutoAnalysisSaveNETSurfaceAndResidualPlots objMSAlign, udtAutoParams, udtWorkingParams, fso, True
                        
                    ' -----------------------------------------
                    ' Save the mass-related plots to disk
                    ' -----------------------------------------
                    
                    ' Enlarge the window to the desired size (pixels)
                    lngWindowWidth = glbPreferencesExpanded.AutoAnalysisOptions.SavePictureWidthPixels
                    lngWindowHeight = glbPreferencesExpanded.AutoAnalysisOptions.SavePictureHeightPixels
                    ValidateValueLng lngWindowWidth, 64, 10000, 1024
                    ValidateValueLng lngWindowHeight, 64, 10000, 768
                        
                    With objMSAlign
                        ' The user may have minimized the window during auto analysis
                        strLastGoodLocation = "Assure .WindowState = vbNormal"
                        If .WindowState <> vbNormal Then
                            .WindowState = vbNormal
                        End If
                        
                        .ScaleMode = vbTwips
                        lngWindowWidthTwips = lngWindowWidth * Screen.TwipsPerPixelX * 1.3
                        lngWindowHeightTwips = lngWindowHeight * Screen.TwipsPerPixelY * 1.3
                        .width = lngWindowWidthTwips
                        .Height = lngWindowHeightTwips
                        .Top = 0
                        .Left = 0
                        DoEvents
                                                
                        ' -----------------------------------------
                        ' Shrink the window to 55% of the current height so that the two Mass residuals plots will take up less space
                        ' In addition, shrink the width to 75% of the current width
                        ' -----------------------------------------
                        .width = lngWindowWidthTwips * 0.75
                        .Height = lngWindowHeightTwips * 0.55
                        DoEvents
                        
                        lngDefaultMaxMontageWidth = .width / 23.25466
                        lngDefaultMaxMontageHeight = .Height / 29.5233
                        
                        ' -----------------------------------------
                        ' Save the uncorrected Mass vs. Scan Residuals plot
                        ' -----------------------------------------
                        strUniqueSuffix = "_" & Mid(Format(Rnd(), "0.0000000"), 3)
                        strFilePath = udtWorkingParams.ResultsFileNameBase & "_MassCalibrationResidualsVsScan" & strUniqueSuffix & ".png"
                        strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                        strFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, True)

                        strLastGoodLocation = "Call UpdatePlotViewModeToMassVsScanResiduals"
                        TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                        .UpdatePlotViewModeToMassVsScanResiduals
                        Sleep 100
                        DoEvents
                        
                        strLastGoodLocation = "Save Mass vs. Scan Residuals Plot to: " & strFilePath
                        TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                        blnSaveSuccessful = .SaveMassVsScanResidualsPlotToPNG(strFilePath)
                        
                        If blnSaveSuccessful Then
                            ' Save the Corrected Mass Residuals plot
                            strFilePathB = udtWorkingParams.ResultsFileNameBase & "_MassCalibrationResidualsVsScanCorrected" & strUniqueSuffix & ".png"
                            strFilePathB = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePathB)
                            strFilePathB = objRemoteSaveFileHandler.GetTempFilePath(strFilePathB, True)
                            
                            strLastGoodLocation = "Call UpdatePlotViewModeToMassVsScanCorrectedResiduals"
                            TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                            .UpdatePlotViewModeToMassVsScanCorrectedResiduals
                            Sleep 100
                            DoEvents
                            
                            strLastGoodLocation = "Save Warped Mass vs. Scan Residuals Plot to: " & strFilePath
                            TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                            blnSaveSuccessful = .SaveMassVsScanResidualsPlotToPNG(strFilePathB)
                            
                            If blnSaveSuccessful Then
                                ' Create a montage using the two files
                                strCombinedFilePath = udtWorkingParams.ResultsFileNameBase & "_MassCalibrationResidualsVsScan.png"
                                strCombinedFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strCombinedFilePath)
                                strWorkingCombinedFilePath = objRemoteSaveFileHandler.GetTempFilePath(strCombinedFilePath, False)

                                lngResult = CreateMontageImageFile(udtWorkingParams, strFilePath, strFilePathB, strWorkingCombinedFilePath, lngDefaultMaxMontageWidth, lngDefaultMaxMontageHeight)
                                    
                                If lngResult = 0 Then
                                    objRemoteSaveFileHandler.MoveTempFileToFinalDestination
                                    blnSaveSuccessful = True
                                Else
                                    blnSaveSuccessful = False
                                End If

                                If blnSaveSuccessful Then
                                    AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strCombinedFilePath), "Mass vs. Scan Calibration Residuals", 5, 1
                                    udtWorkingParams.TICPlotsStartRow = 6
                                End If
                            End If
                        End If
                
                
                        ' -----------------------------------------
                        ' Save the uncorrected Mass vs. m/z Residuals plot
                        ' -----------------------------------------
                        strUniqueSuffix = "_" & Mid(Format(Rnd(), "0.0000000"), 3)
                        strFilePath = udtWorkingParams.ResultsFileNameBase & "_MassCalibrationResidualsVsMZ" & strUniqueSuffix & ".png"
                        strFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePath)
                        strFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, True)
                        
                        strLastGoodLocation = "Call UpdatePlotViewModeToMassVsMZResiduals"
                        TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                        .UpdatePlotViewModeToMassVsMZResiduals
                        Sleep 100
                        DoEvents
                        
                        strLastGoodLocation = "Save Mass vs. m/z Residuals Plot to: " & strFilePath
                        TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                        blnSaveSuccessful = .SaveMassVsMZResidualsPlotToPNG(strFilePath)
                        
                        If blnSaveSuccessful Then
                            ' Save the Corrected Mass Residuals plot
                            strFilePathB = udtWorkingParams.ResultsFileNameBase & "_MassCalibrationResidualsVsMZCorrected" & strUniqueSuffix & ".png"
                            strFilePathB = fso.BuildPath(udtWorkingParams.GelOutputFolder, strFilePathB)
                            strFilePathB = objRemoteSaveFileHandler.GetTempFilePath(strFilePathB, True)

                            strLastGoodLocation = "Call UpdatePlotViewModeToMassVsMZCorrectedResiduals"
                            TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                            .UpdatePlotViewModeToMassVsMZCorrectedResiduals
                            Sleep 100
                            DoEvents
                            
                            strLastGoodLocation = "Save Warped Mass vs. m/z Residuals Plot to: " & strFilePath
                            TraceLog 5, "AutoAnalysisToleranceRefinement", strLastGoodLocation
                            blnSaveSuccessful = .SaveMassVsMZResidualsPlotToPNG(strFilePathB)
                            
                            If blnSaveSuccessful Then
                                ' Create a montage using the two files
                                strCombinedFilePath = udtWorkingParams.ResultsFileNameBase & "_MassCalibrationResidualsVsMZ.png"
                                strCombinedFilePath = fso.BuildPath(udtWorkingParams.GelOutputFolder, strCombinedFilePath)
                                strWorkingCombinedFilePath = objRemoteSaveFileHandler.GetTempFilePath(strCombinedFilePath, False)

                                lngResult = CreateMontageImageFile(udtWorkingParams, strFilePath, strFilePathB, strWorkingCombinedFilePath, lngDefaultMaxMontageWidth, lngDefaultMaxMontageHeight)
                                
                                If lngResult = 0 Then
                                    objRemoteSaveFileHandler.MoveTempFileToFinalDestination
                                    blnSaveSuccessful = True
                                Else
                                    blnSaveSuccessful = False
                                End If

                                If blnSaveSuccessful Then
                                    AddNewOutputFileForHtml udtWorkingParams, fso.GetFileName(strCombinedFilePath), "Mass vs. m/z Calibration Residuals", 5, 2
                                    udtWorkingParams.TICPlotsStartRow = 6
                                End If
                            End If
                        End If
                                
                        ' Restore the window size
                        .width = lngWindowWidthTwips
                        .Height = lngWindowHeightTwips
                        DoEvents
                    End With
                
                Else
                    ' Alignment failed
                    blnSuccess = False
                End If
               
                ' 4b. Possibly update the DB Search mass tolerance using LCMSWarp
                blnSearchMassTolRefinedByMSWarp = False
                If blnSuccess And glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance Then
                    ' Possible future task: use frmMSAlign to auto-determine the mass tolerance to use
                    '' blnSearchMassTolRefinedByMSWarp = True
                End If
                
                ' 4c. Possibly update the DB Search NET tolerance using LCMSWarp
                blnSearchNETTolRefinedByMSWarp = False
                If blnSuccess And glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance Then
                    ' Possible future task: use frmMSAlign to auto-determine the NET tolerance to use
                    '' blnSearchNETTolRefinedByMSWarp = True
                End If
                               
                               
                strLastGoodLocation = "Unload objMSAlign"
                Unload objMSAlign

                If blnSuccess Then
                    ' 4d. Repeat the search now that the masses have been updated
                    ' Necessary since MS Warp may have shifted the masses a large distance, which would make the tolerance refinement plots unlikely to contain useful information
                    strLastGoodLocation = "Call AutoAnalysisSearchDatabase"
                    AutoAnalysisSearchDatabase udtWorkingParams, udtAutoParams, fso, True
                End If

                With frmErrorDistribution2DLoadedData
                    ' 4e. Use frmErrorDistribution2DLoadedData to update the peak stats
                    ' If APPLY_ADDNL_LINEAR_MASS_ADJUSTMENT then also apply an additional linear adjustment if needed
                    .UpdateUMCStatsAndRecomputeErrors
                    blnSuccess = .RefineMassCalibrationStart(blnValidPeakFound, blnMassShiftTooLarge, blnPeakTooWide, Not APPLY_ADDNL_LINEAR_MASS_ADJUSTMENT)
                End With

                ' 5. Save the error distributions after refinement (but using the wider search tolerances)
                AutoAnalysisToleranceRefinementSaveErrorDistributions frmErrorDistribution2DLoadedData, udtAutoParams, udtWorkingParams, fso, False, ehmErrorHistogramModeConstants.ehmAfterLCMSWARP

                ' 6a. Possibly update the DB Search mass tolerance
                If blnSuccess And glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance Then
                    If blnSearchMassTolRefinedByMSWarp Then
                        blnSuccess = True
                    Else
                        blnSuccess = AutoAnalysisToleranceRefinementRefineDBSearchMass(frmErrorDistribution2DLoadedData, udtWorkingParams, strErrorMessage)
                    End If
                End If
                
                ' 6b. Possibly update the DB Search NET tolerance
                If blnSuccess And glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance Then
                    If blnSearchNETTolRefinedByMSWarp Then
                        blnSuccess = True
                    Else
                        blnSuccess = AutoAnalysisToleranceRefinementRefineDBSearchNET(frmErrorDistribution2DLoadedData, udtWorkingParams, strErrorMessage)
                    End If
                End If
                
            Else
                With frmErrorDistribution2DLoadedData
                    ' 4a. Use frmErrorDistribution2DLoadedData
                    strLastGoodLocation = "Call RefineMassCalibrationStart"
                    blnSuccess = .RefineMassCalibrationStart(blnValidPeakFound, blnMassShiftTooLarge, blnPeakTooWide, False)
                End With
            
               ' 4b. Possibly update the DB Search mass tolerance
                If blnSuccess And glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance Then
                    blnSuccess = AutoAnalysisToleranceRefinementRefineDBSearchMass(frmErrorDistribution2DLoadedData, udtWorkingParams, strErrorMessage)
                End If
                
                ' 4c. Possibly update the DB Search NET tolerance
                If blnSuccess And glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance Then
                    blnSuccess = AutoAnalysisToleranceRefinementRefineDBSearchNET(frmErrorDistribution2DLoadedData, udtWorkingParams, strErrorMessage)
                End If
                
            End If
        End If
        
        If Not blnSuccess Then
            strErrorMessage = "Tolerance Refinement error occured while refining the mass calibration"
            If blnMassShiftTooLarge Then
                strErrorMessage = strErrorMessage & "; the mass calibration adjustment value was too large"
            End If
        Else
            If Not blnValidPeakFound Then
                ' A valid peak was not found; flag this as a warning, not an error
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_PEAK_NOT_FOUND_BIT
            ElseIf blnPeakTooWide Then
                ' A peak was found, but it was too wide; flag this as a warning, not an error
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_BIT_PEAK_TOO_WIDE
            End If
        End If
    End If
    
    Unload frmErrorDistribution2DLoadedData
    
    If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance Then
        ' Cache the newly determined mass tolerance
        dblRefinedMassTol = samtDef.MWTol
        eRefinedTolType = samtDef.TolType
    End If
    
    If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance Then
        ' Cache the newly determined NET tolerance
        dblRefinedNETTol = samtDef.NETTol
    End If
    
    ' 5a. Restore samtDef
    samtDef = udtAMTDefSaved
    
    If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchMassTolerance Then
        ' Update .MWTol to the refined mass tolerance
        samtDef.MWTol = dblRefinedMassTol
        samtDef.TolType = eRefinedTolType
    End If
    
    If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineDBSearchNETTolerance Then
        ' Update .NETTol to the refined NET tolerance
        samtDef.NETTol = dblRefinedNETTol
    End If
    
    ' 5b. Copy samtDef to the other search definitions
    With GelSearchDef(udtWorkingParams.GelIndex)
        .AMTSearchOnIons = samtDef
        .AMTSearchOnUMCs = samtDef
        .AMTSearchOnPairs = samtDef
    End With
    
    ' 5c. Restore .AutoAnalysisOptions
    glbPreferencesExpanded.AutoAnalysisOptions = udtAutoAnalysisOptionsSaved
    
    ' 5d. Restore .MinimumSLiC
    glbPreferencesExpanded.RefineMSDataOptions.MinimumPeakHeight = sngSLiCSaved
    
    ' 5e. Write the latest AnalysisHistory information to the log
    AutoAnalysisAppendLatestHistoryToLog udtAutoParams, udtWorkingParams

    If Not blnSuccess Then
        AutoAnalysisLog udtAutoParams, udtWorkingParams, "Error - " & strErrorMessage
        udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or TOLERANCE_REFINEMENT_ERROR_BIT
    End If

    Exit Sub
    
AutoAnalysisToleranceRefinementErrorHandler:
    Debug.Assert False
    strMessage = "Error - An error has occurred in the tolerance refinement step during auto analysis (LastGoodLocation=" & strLastGoodLocation & "): " & Err.Description
    If udtWorkingParams.ts Is Nothing And udtAutoParams.ShowMessages Then
        MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
    Else
        AutoAnalysisLog udtAutoParams, udtWorkingParams, strMessage
    End If

End Sub

Private Sub AutoAnalysisValidateToleranceRefinementSearchMode(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef udtSearchMode As udtAutoAnalysisSearchModeOptionsType, ByRef udtAutoAnalysisOptionsSaved As udtAutoAnalysisOptionsType, ByVal intNextSearchModeIndex As Integer)
    ' Validate that udtSearchMode is appropriate for tolerance refinement
    
    Dim eDBSearchMode As dbsmDatabaseSearchModeConstants
    
    eDBSearchMode = LookupDBSearchModeIndex(udtSearchMode.SearchMode)
    Select Case eDBSearchMode
    Case dbsmIndividualPeaks
        ' This search mode is fine
    Case dbsmIndividualPeaksInUMCsWithNET, dbsmIndividualPeaksInUMCsPaired, dbsmIndividualPeaksInUMCsUnpaired
        ' These search modes are no longer supported
        ' Use AUTO_SEARCH_UMC_CONGLOMERATE instead
        udtSearchMode.SearchMode = AUTO_SEARCH_UMC_CONGLOMERATE
    Case dbsmConglomerateUMCsWithNET, dbsmConglomerateUMCsPaired, dbsmConglomerateUMCsUnpaired, dbsmConglomerateUMCsLightPairsPlusUnpaired
        ' This search mode is fine
        ' However, if InternalStdSearchMode = issmFindOnlyInternalStandards, then check if additional
        '  search modes are listed; if they are, choose the next valid one
        If udtSearchMode.InternalStdSearchMode = issmFindOnlyInternalStandards Then
            If udtAutoAnalysisOptionsSaved.AutoAnalysisSearchModeCount > intNextSearchModeIndex Then
                If udtAutoAnalysisOptionsSaved.AutoAnalysisSearchMode(intNextSearchModeIndex).InternalStdSearchMode <> issmFindOnlyInternalStandards Then
                    ' Recursively call this function to validate .AutoAnalysisSearchMode(intNextSearchModeIndex)
                    udtSearchMode = udtAutoAnalysisOptionsSaved.AutoAnalysisSearchMode(intNextSearchModeIndex)
                    AutoAnalysisValidateToleranceRefinementSearchMode udtWorkingParams, udtSearchMode, udtAutoAnalysisOptionsSaved, intNextSearchModeIndex + 1
                    AddToAnalysisHistory udtWorkingParams.GelIndex, "Note: The first search mode listed was an Internal Standard Only search (ConglomerateUMCsWithNET), which isn't ideal for Tolerance Refinement.  Since more than one search mode was specified, automatically switched to " & udtSearchMode.SearchMode
                End If
            End If
        End If
    Case dbsmPairsN14N15, dbsmPairsN14N15ConglomerateMass, dbsmPairsICAT, dbsmPairsPEO
        ' This search mode is fine, unless .AutoAnalysisRemovePairMemberHitsAfterDBSearch = True
        If glbPreferencesExpanded.PairSearchOptions.AutoAnalysisRemovePairMemberHitsAfterDBSearch Then
            ' These pairs-based search modes are too strict for tolerance refinement
            '  when removing pair member hits after the DB search
            ' Use AUTO_SEARCH_UMC_CONGLOMERATE_PAIRED instead
            udtSearchMode.SearchMode = AUTO_SEARCH_UMC_CONGLOMERATE_PAIRED
        End If
    Case dbsmIndividualPeaksInUMCsWithoutNET
        ' Cannot use this search mode for tolerance refinement since it doesn't use NET values
        ' Use AUTO_SEARCH_UMC_CONGLOMERATE instead
        udtSearchMode.SearchMode = AUTO_SEARCH_UMC_CONGLOMERATE
        udtSearchMode.InternalStdSearchMode = issmFindWithMassTags
    Case Else
        ' Other search mode; this shouldn't happen
        Debug.Assert False
        udtSearchMode.SearchMode = AUTO_SEARCH_UMC_CONGLOMERATE
    End Select
 
    If udtSearchMode.InternalStdSearchMode = issmFindOnlyInternalStandards Then
        udtSearchMode.InternalStdSearchMode = issmFindWithMassTags
    End If
    
End Sub

Private Function AutoAnalysisToleranceRefinementRefineDBSearchMass(ByRef objErrorDistribution2DLoadedData As frmErrorDistribution2DLoadedData, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef strErrorMessage As String) As Boolean
    Dim blnSuccess As Boolean
    Dim blnValidPeakFound As Boolean, blnPeakTooWide As Boolean
    
    blnSuccess = objErrorDistribution2DLoadedData.RefineDBSearchMassToleranceStart(blnValidPeakFound, blnPeakTooWide)
    If Not blnSuccess Then
        strErrorMessage = "Tolerance Refinement error occured while refining the DB search mass tolerance"
    Else
        If Not blnValidPeakFound Then
            ' A valid peak was not found; flag this as a warning, not an error
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_PEAK_NOT_FOUND_BIT
        ElseIf blnPeakTooWide Then
            ' A peak was found, but it was too wide; flag this as a warning, not an error
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_BIT_PEAK_TOO_WIDE
        End If
    End If
    
    AutoAnalysisToleranceRefinementRefineDBSearchMass = blnSuccess

End Function
                    
Private Function AutoAnalysisToleranceRefinementRefineDBSearchNET(ByRef objErrorDistribution2DLoadedData As frmErrorDistribution2DLoadedData, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef strErrorMessage As String) As Boolean
    Dim blnSuccess As Boolean
    Dim blnValidPeakFound As Boolean, blnPeakTooWide As Boolean
    
    blnSuccess = objErrorDistribution2DLoadedData.RefineDBSearchNETToleranceStart(blnValidPeakFound, blnPeakTooWide)
    If Not blnSuccess Then
        strErrorMessage = "Tolerance Refinement error occured while refining the DB search NET tolerance"
    Else
        If Not blnValidPeakFound Then
            ' A valid peak was not found; flag this as a warning, not an error
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_PEAK_NOT_FOUND_BIT
        ElseIf blnPeakTooWide Then
            ' A peak was found, but it was too wide; flag this as a warning, not an error
            udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_BIT_PEAK_TOO_WIDE
        End If
    End If
    
    AutoAnalysisToleranceRefinementRefineDBSearchNET = blnSuccess
End Function

Private Sub AutoAnalysisToleranceRefinementSaveErrorDistributions(ByRef objErrorDistribution2DLoadedData As frmErrorDistribution2DLoadedData, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByRef fso As FileSystemObject, ByVal blnUpdateCachedErrorPeakStats As Boolean, ByVal eErrorHistogramMode As ehmErrorHistogramModeConstants)

    ' Save the error distributions
    With objErrorDistribution2DLoadedData
        
        ' Make sure the tolerance refinement controls are hidden
        .ShowHideToleranceRefinementControls False
        
        If glbPreferencesExpanded.AutoAnalysisOptions.AutoToleranceRefinement.RefineMassCalibration Then
            ' Save a mass error plot (picture and .txt file) so that we have a record of the mass errors before mass calibration
            If glbPreferencesExpanded.RefineMSDataOptions.MassCalibrationTolType = gltPPM Then
                .SetPlotMode mdmMassErrorPPM
            Else
                .SetPlotMode mdmMassErrorDa
            End If
        Else
            .SetPlotMode mdmMassErrorPPM
        End If
        
        ' 3. Save the error distributions (Note that blnUpdateCachedErrorPeakStats is ignored here since blnSavingDataDuringToleranceRefinement = True
        AutoAnalysisSaveErrorDistributions udtWorkingParams, udtAutoParams, fso, True, blnUpdateCachedErrorPeakStats, eErrorHistogramMode, 0
        
        ' Make sure the tolerance refinement controls are shown
        .ShowHideToleranceRefinementControls True
        .SetPlotMode mdmMassErrorPPM
        
        ' Possibly Update the cached error peak stats
        If blnUpdateCachedErrorPeakStats Then
            .RecordMassCalPeakStatsNow
            .RecordNETTolPeakStatsNow
        End If
    End With

End Sub
 
Private Sub AutoAnalysisZoomOut(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, blnRemoveFilters As Boolean)
    
    ' If the data falls in a range less than this, then the Gel is zoomed to at least these values
    Const DEFAULT_MINIMUM_MASS = 300
    Const DEFAULT_MAXIMUM_MASS = 5000
    
    Dim lngGelScanNumberMin As Long, lngGelScanNumberMax As Long
    Dim dblGelMassMin As Double, dblGelMassMax As Double
    
    ' Make sure the scan range is zoomed out full and the mass range is zoomed out to at least DEFAULT_MINIMUM_MASS to DEFAULT_MAXIMUM_MASS
    ' Determine the minimum and maximum scan numbers possible
    GetScanRange udtWorkingParams.GelIndex, lngGelScanNumberMin, lngGelScanNumberMax, 0, 0
    
    ' Determine the current mass range
    GetMassRangeCurrent udtWorkingParams.GelIndex, dblGelMassMin, dblGelMassMax
    If dblGelMassMin > DEFAULT_MINIMUM_MASS Then dblGelMassMin = DEFAULT_MINIMUM_MASS
    If dblGelMassMax < DEFAULT_MAXIMUM_MASS Then dblGelMassMax = DEFAULT_MAXIMUM_MASS

    ZoomGelToDimensions udtWorkingParams.GelIndex, CSng(lngGelScanNumberMin), dblGelMassMin, CSng(lngGelScanNumberMax), dblGelMassMax
    
    If blnRemoveFilters Then
        ' Remove any mass filters
        With glbPreferencesExpanded.AutoAnalysisFilterPrefs
            If .RestrictIsoByMass Or .RestrictCSByMass Or .RestrictIsoByMZ Or GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 0) Then
                GelData(udtWorkingParams.GelIndex).DataFilter(fltIsoMW, 0) = False
                GelData(udtWorkingParams.GelIndex).DataFilter(fltIsoMZ, 0) = False
                GelData(udtWorkingParams.GelIndex).DataFilter(fltCSMW, 0) = False
                    
                ' Show all ions, not just the hits
                GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 0) = False
                GelData(udtWorkingParams.GelIndex).DataFilter(fltID, 1) = 0               ' Set to 0 to show all data
                    
                ' Assign udtWorkingParams.GelIndex to frmFilter.Tag, then call .InitializeControls
                ' Since glbPreferencesExpanded.AutoAnalysisStatus.Enabled = True, will
                '  automatically unload the form after the filter is applied
                frmFilter.Tag = udtWorkingParams.GelIndex
                frmFilter.InitializeControls True
            End If
        End With
    End If
    
End Sub

Public Sub AutoGenerateQCPlots(objCallingForm As Form, lngGelIndex As Long, Optional strFileNamePathBase As String = "", Optional blnShowMessages As Boolean = True)
    ' Generates the various QC plots for the given Gel
    ' If strFileNamePathBase is blank, then prompts the user for the folder and base name to use
    ' If strFileNamePathBase only contains a folder name, then auto-assigns the base name
    
    Dim strBaseFileName As String
    Dim fso As FileSystemObject
    
    Dim eResponse As VbMsgBoxResult
    Dim strErrorMessage As String
    
    Dim udtWorkingParams As udtAutoAnalysisWorkingParamsType
    Dim udtAutoParams As udtAutoAnalysisParametersType
    
    Dim dblMassErrorOverallSaved As Double
    Dim eMassUnitsSaved As glMassToleranceConstants
    
    Dim blnGenerateRobustNETPlots As Boolean
    Dim blnAMTsWereLoaded As Boolean
    Dim blnDBConnectionError As Boolean
    
    Dim blnAutoAnalysisEnabled As Boolean
    Dim blnSuccess As Boolean
    
On Error GoTo AutoGenerateQCPlotsErrorHandler
    
    If lngGelIndex < 1 Or lngGelIndex > UBound(GelBody()) Then
        ' Invalid Gel Index
        Exit Sub
    End If
    
    Set fso = New FileSystemObject
    
    ' Determine whether or not to show messages
    blnAutoAnalysisEnabled = glbPreferencesExpanded.AutoAnalysisStatus.Enabled
    udtAutoParams.ShowMessages = Not blnAutoAnalysisEnabled And blnShowMessages

    With udtWorkingParams
        .GelIndex = lngGelIndex
        .GraphicOutputFileInfoCount = 0
        ReDim .GraphicOutputFileInfo(0)
        .TICPlotsStartRow = 4
    End With
    
    ' Check whether Custom NETs exist or .UseRobustNETAdjustment is enabled
    With GelUMCNETAdjDef(udtWorkingParams.GelIndex)
        If GelData(udtWorkingParams.GelIndex).CustomNETsDefined Then
            blnGenerateRobustNETPlots = True
            glbPreferencesExpanded.AutoAnalysisOptions.SkipGANETSlopeAndInterceptComputation = False
            .UseRobustNETAdjustment = True
            If .RobustNETAdjustmentMode < UMCRobustNETModeConstants.UMCRobustNETWarpTime Then
                .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTime
            End If
        ElseIf glbPreferencesExpanded.AutoAnalysisOptions.SkipGANETSlopeAndInterceptComputation Then
            blnGenerateRobustNETPlots = False
        Else
            If .UseRobustNETAdjustment And .RobustNETAdjustmentMode >= UMCRobustNETModeConstants.UMCRobustNETWarpTime Then
                blnGenerateRobustNETPlots = True
            End If
        End If
    
        If APP_BUILD_DISABLE_LCMSWARP Then
            blnGenerateRobustNETPlots = False
        End If
    End With

    If Not blnAutoAnalysisEnabled Then
        If Len(Trim(strFileNamePathBase)) = 0 Then
            strFileNamePathBase = SelectFile(objCallingForm.hwnd, "Select base file name and output folder", "", True, fso.GetBaseName(GelData(udtWorkingParams.GelIndex).FileName), "All Files (*.*)|*.*")
        End If
        
        If udtAutoParams.ShowMessages And blnGenerateRobustNETPlots Then
            eResponse = MsgBox("Include NET alignment plots?  If yes, this then NET alignment will be repeated.", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Include NET Plots")
            If eResponse <> vbYes Then
                blnGenerateRobustNETPlots = False
            End If
        End If
    End If

    ' Extract the folder information from strFileNamePathBase
    If fso.GetParentFolderName(strFileNamePathBase) = "" Then
        strFileNamePathBase = fso.BuildPath(App.Path, strFileNamePathBase)
    End If
    
    udtAutoParams.FilePaths.OutputFolderPath = fso.GetParentFolderName(strFileNamePathBase)
    If Len(udtAutoParams.FilePaths.OutputFolderPath) = 0 Then
        udtAutoParams.FilePaths.OutputFolderPath = App.Path
    End If
    
    strBaseFileName = fso.GetBaseName(strFileNamePathBase)
    If Len(strBaseFileName) = 0 Then
        strBaseFileName = "Untitled"
    End If
    
    ' Initialize udtAutoParams and udtWorkingParams
    With udtAutoParams
        .FilePaths.InputFilePath = strBaseFileName & ".pek"
        .MTDBOverride.Enabled = False
    End With
    
    With udtWorkingParams
        .GelOutputFolder = udtAutoParams.FilePaths.OutputFolderPath
        .ResultsFileNameBase = strBaseFileName
    End With
    
    ' Override .AutoAnalysisStatus.Enabled so that error messages do not appear when generating the QC plots
    With glbPreferencesExpanded.AutoAnalysisStatus
        .Enabled = True
        .AutoAnalysisTimeStamp = Format(Now(), "yyyy.mm.dd Hh:Nn:Ss")
    End With

            
    ' Apply the default filters to the data
    AutoAnalysisFilterData udtWorkingParams, udtAutoParams
    
    If AMTCnt = 0 And udtAutoParams.ShowMessages Then
        eResponse = MsgBox("MT tags not in memory.  Load from the database?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load MT tags")
    Else
        eResponse = vbYes
    End If
    
    If eResponse = vbYes Then
        If ConfirmMassTagsAndInternalStdsLoaded(objCallingForm, udtWorkingParams.GelIndex, udtAutoParams.ShowMessages, 0, False, True, blnAMTsWereLoaded, blnDBConnectionError) Then
            ' ConstructMTStatusText(True)
        Else
            If blnDBConnectionError Then
                If udtAutoParams.ShowMessages Then
                   strErrorMessage = "Error loading MT tags: database connection error."
                End If
            Else
                strErrorMessage = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
            End If
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, "Error"
        End If
    End If
    
    If AMTCnt > 0 Then
        ' If Robust NET alignment is enabled, then perform NET adjustment to generate the output files
        ' Otherwise, do not generate the NET alignment files

        If blnGenerateRobustNETPlots Then
            ' Robust NET is enabled
            AutoAnalysisPerformNETAdjustment udtWorkingParams, udtAutoParams, fso
        End If
       
        ' Save a graphical picture of the data
        AutoAnalysisSavePictureGraphic udtWorkingParams, udtAutoParams, fso, False, True

        ' Create the mass and NET error plots
        AutoAnalysisSaveErrorDistributions udtWorkingParams, udtAutoParams, fso, False, False, ehmErrorHistogramModeConstants.ehmFinalTolerances, 2

''        ' First, save the current overall adjustment value
''        With GelSearchDef(udtWorkingParams.GelIndex).MassCalibrationInfo
''            If .AdjustmentHistoryCount > 0 Then
''                dblMassErrorOverallSaved = .OverallMassAdjustment
''                eMassUnitsSaved = .MassUnits
''                blnSuccess = MassCalibrationRevertToOriginal(udtWorkingParams.GelIndex, False, False)
''                blnSuccess = UpdateUMCStatArrays(lngGelIndex, False, objCallingForm)
''            Else
''                dblMassErrorOverallSaved = 0
''            End If
''        End With
''
''        ' Next, save a picture or text file of the mass and GANET error distribution "before refinement"
''        AutoAnalysisSaveErrorDistributions udtWorkingParams, udtAutoParams, fso, True, True, ehmErrorHistogramModeConstants.ehmBeforeRefinement, 1
''
''        If True Then
''            frmErrorDistribution2DLoadedData.CallerID = udtWorkingParams.GelIndex
''            frmErrorDistribution2DLoadedData.show vbModeless
''            frmErrorDistribution2DLoadedData.InitializeForm
''
''            frmErrorDistribution2DLoadedData.RecordMassCalPeakStatsNow
''            frmErrorDistribution2DLoadedData.RecordNETTolPeakStatsNow
''            Unload frmErrorDistribution2DLoadedData
''        End If
''
''        ' Now, reapply the overall adjustment value (if any)
''        If dblMassErrorOverallSaved <> 0 Then
''            blnSuccess = MassCalibrationApplyBulkAdjustment(udtWorkingParams.GelIndex, dblMassErrorOverallSaved, eMassUnitsSaved, False)
''
''            ' Finally, save a picture or text file of the mass and GANET error distribution "after refinement"
''            AutoAnalysisSaveErrorDistributions udtWorkingParams, udtAutoParams, fso, False, False, ehmErrorHistogramModeConstants.ehmFinalTolerances, 2
''        End If
        
    End If
        
    ' Save a picture and text file of the various chromatograms
    AutoAnalysisSaveChromatograms udtWorkingParams, udtAutoParams, fso
    
    ' 21. Create a browsable Index.html file
    AutoAnalysisGenerateHTMLBrowsingFile udtWorkingParams, udtAutoParams, fso
        
    glbPreferencesExpanded.AutoAnalysisStatus.Enabled = blnAutoAnalysisEnabled
        
    If udtAutoParams.ShowMessages Then
        MsgBox "QC Plots have been created in folder: " & udtAutoParams.FilePaths.OutputFolderPath, vbInformation + vbOKOnly, "Done"
    End If
    
    Exit Sub
   
AutoGenerateQCPlotsErrorHandler:
    If Not blnAutoAnalysisEnabled Then
        MsgBox "Error auto generating QC plots: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        LogErrors Err.Number, "LookupMatchingUMCStats"
    End If
    
    glbPreferencesExpanded.AutoAnalysisStatus.Enabled = blnAutoAnalysisEnabled
End Sub

Private Function CheckAutoToleranceRefinementEnabled(ByRef udtAutoToleranceRefinement As udtAutoToleranceRefinementType, ByRef udtAutoParams As udtAutoAnalysisParametersType, ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByVal blnLogWarnings As Boolean) As Boolean

    Dim blnPerformRefinement As Boolean

    With udtAutoToleranceRefinement
        If .RefineMassCalibration Or .RefineDBSearchMassTolerance Or .RefineDBSearchNETTolerance Then
            blnPerformRefinement = True
        Else
            blnPerformRefinement = False
        End If
        If glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchModeCount < 1 Then
            If blnLogWarnings Then
                AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Tolerance Refinement is enabled, but no database search modes are defined; skipping tolerance refinement"
                udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_PEAK_NOT_FOUND_BIT
            End If
            blnPerformRefinement = False
        Else
            If LookupDBSearchModeIndex(glbPreferencesExpanded.AutoAnalysisOptions.AutoAnalysisSearchMode(0).SearchMode) = dbsmExportUMCsOnly Then
                If blnLogWarnings Then
                    AutoAnalysisLog udtAutoParams, udtWorkingParams, "Warning - Tolerance Refinement is enabled, but the database search mode is Export UMCs Only; skipping tolerance refinement"
                    udtWorkingParams.WarningBits = udtWorkingParams.WarningBits Or TOLERANCE_REFINEMENT_WARNING_PEAK_NOT_FOUND_BIT
                End If
                blnPerformRefinement = False
            End If
        End If
    End With

    CheckAutoToleranceRefinementEnabled = blnPerformRefinement
End Function

Private Sub CheckMessageBit(ByVal lngErrorBits As Long, ByVal lngCheckBit As Long, ByRef strMessageLog As String, ByVal strErrorMessage As String)

    If (lngErrorBits And lngCheckBit) = lngCheckBit Then
        If Len(strMessageLog) > 0 Then strMessageLog = strMessageLog & vbCrLf
        strMessageLog = strMessageLog & "Code " & Trim(lngCheckBit) & ": " & strErrorMessage
    End If
    
End Sub

Private Sub CountSubFolders(strFolderStartPath As String, ByRef fso As FileSystemObject, ByRef lngFolderCount As Long, ByVal lngFolderCountAbortSearch As Long)

    Const PROGRESS_SUB_LOOP_COUNT As Long = 500

    Dim objFolder As Folder
    Dim objSubFolder As Folder

    If KeyPressAbortProcess > 1 Then Exit Sub

    Set objFolder = fso.GetFolder(strFolderStartPath)
    
    For Each objSubFolder In objFolder.SubFolders
        lngFolderCount = lngFolderCount + 1
        
        If lngFolderCount Mod 10 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngFolderCount Mod PROGRESS_SUB_LOOP_COUNT, False
            frmProgress.UpdateCurrentSubTask "Counting subfolders: " & Trim(lngFolderCount)
        
            If KeyPressAbortProcess > 1 Then Exit For
        End If
        
        If lngFolderCount > lngFolderCountAbortSearch Then Exit For
        
        CountSubFolders objSubFolder.Path, fso, lngFolderCount, lngFolderCountAbortSearch
    Next objSubFolder

End Sub

Private Function CreateMontageImageFile(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByVal strSrcFilePathA As String, ByVal strSrcFilePathB As String, ByVal strDestinationFilePath As String, Optional lngDefaultMaxWidth As Long = 644, Optional lngDefaultMaxHeight As Long = 279) As Long
    ' Note: lngDefaultMaxWidth and lngDefaultMaxHeight are used as the maximum dimensions if the Identify command fails
    
    Const MAX_MONTAGE_CREATION_RETRY_COUNT As Integer = 3
    
    Dim lngResult As Long
    Dim intRetryCount As Integer
    Dim blnSuccess As Boolean
    Dim fso As FileSystemObject
    
    ' Sleep 2500 msec before trying to create the montage file
    DoEvents
    Sleep 2500
    DoEvents
    
    Set fso = New FileSystemObject
    
    blnSuccess = False
    intRetryCount = 0
    Do While intRetryCount < MAX_MONTAGE_CREATION_RETRY_COUNT
        lngResult = CreateMontageImage(strSrcFilePathA, strSrcFilePathB, strDestinationFilePath, True, lngDefaultMaxWidth, lngDefaultMaxHeight)
        
        ' Make sure strDestinationFilePath exists
        blnSuccess = fso.FileExists(strDestinationFilePath)
        If blnSuccess And lngResult = 0 Then
            Exit Do
        Else
            ' Montage file creation failed
            ' Sleep for a random interval between 1000 and 3500 milliseconds
            Sleep Int((3500 - 1000 + 1) * Rnd() + 1000)

            If intRetryCount = 0 Then
                ' This if the first attempt
                ' Make sure the source files exist
                If Not fso.FileExists(strSrcFilePathA) Then
                    lngResult = -11
                    Exit Do
                ElseIf Not fso.FileExists(strSrcFilePathB) Then
                    lngResult = -12
                    Exit Do
                End If
            End If
            intRetryCount = intRetryCount + 1
        End If
    Loop

    If Not blnSuccess Or lngResult <> 0 Then
        AddToAnalysisHistory udtWorkingParams.GelIndex, "Error creating montage of files '" & strSrcFilePathA & "' and '" & strSrcFilePathB & "'"
        udtWorkingParams.ErrorBits = udtWorkingParams.ErrorBits Or SAVE_GRAPHIC_ERROR_BIT
        
        If gTraceLogLevel = 0 Then
            ' Note: Only delete the source files if the trace log level is 0
            On Error Resume Next
            fso.DeleteFile strSrcFilePathA, True
            fso.DeleteFile strSrcFilePathB, True
        End If
        
        If lngResult = 0 Then
            lngResult = -1
        End If
    ElseIf intRetryCount > 0 Then
        AddToAnalysisHistory udtWorkingParams.GelIndex, "Warning: had to retry " & Trim(intRetryCount) & " times when creating the montage of files '" & strSrcFilePathA & "' and '" & strSrcFilePathB & "'; however, the operation was finally successful"
    End If
    
    CreateMontageImageFile = lngResult
End Function

Public Sub GenerateAutoAnalysisHtmlFiles(Optional ByVal strFolderStartPath As String = "", Optional ByVal strFolderPathTargetMask As String = "", Optional ByVal blnOverwriteExistingIndexFiles As Boolean = False, Optional ByRef lngFoldersParsed As Long = 0, Optional ByVal intRecursionLevel As Integer = 0, Optional ByVal blnShowMessages As Boolean = True)
    ' Generates a browable hierarchy of Index.html files in strFolderStartPath
    ' If strFolderPathTargetMask is defined, then only enters a subfolder
    '  of strFolderStartPath if it is contained in strFolderPathTargetMask
    
    Const INPUT_FILE_HEADER_START As String = "Loading File; " & glCOMMENT_DATA_FILE_START
    Const INPUT_FILE_HEADER_DELIMITER As String = "):"
    Const INPUT_FILE_HEADER_END As String = "; Size"
    
    Const PROGRESS_LOOP_COUNT As Long = 50
    Const PROGRESS_SUB_LOOP_COUNT As Long = 500
    
    Const VERSION_HEADER_START As String = "(version"
    Const VERSION_HEADER_END As String = ")"
    
    Const FOLDER_COUNT_ABORT_SEARCH As Long = 1000
    
    Dim eResponse As VbMsgBoxResult
    
    Dim fso As New FileSystemObject
    Dim objFolder As Folder
    Dim objSubFolder As Folder
    Dim objFile As File
    
    Dim tsInput As TextStream
    
    Static strLastFolderStartPath As String
    Dim strBaseName As String
    Dim strTargetName As String
    Dim strMatchingName As String
    
    Dim strThisFileName As String
    Dim strFileNamesForFolder() As String
    Dim intFileIndex As Integer
    
    Dim strLogFileText As String
    
    Dim strInputFilePath As String
    Dim strVersionOverride As String
    Dim strDateOverride As String
    
    Dim strTemp As String
    Dim lngCharLoc As Long
    
    Dim strFolderPathStripped As String
    
    Dim blnProcessFolder As Boolean
    Dim blnProcessSubFolder As Boolean
    Dim blnResultsFolder As Boolean
    Dim lngFolderCount As Long
    
    Dim lngPeakMatchingTaskID As Long
    Dim strMTDBName As String
    
    Dim udtWorkingParams As udtAutoAnalysisWorkingParamsType
    Dim udtAutoParams As udtAutoAnalysisParametersType
    
    Dim strDateParts() As String
    Dim dtNow As Date
    Dim intYear As Integer, intMonth As Integer, intDay As Integer
    Dim intHour As Integer, intMinute As Integer, intSecond As Integer
    
On Error GoTo GenerateAutoAnalysisHtmlFileErrorHandler

    If Len(strFolderStartPath) = 0 Then
        If blnShowMessages Then
            strFolderStartPath = BrowseForFileOrFolder(MDIForm1.hwnd, strLastFolderStartPath, "Select Starting Folder", True)
        End If
        
        If Len(strFolderStartPath) = 0 Then Exit Sub
        strLastFolderStartPath = strFolderStartPath
        
        eResponse = MsgBox("Overwrite existing Index.html files?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Overwrite")
        If eResponse = vbYes Then
            blnOverwriteExistingIndexFiles = True
        Else
            blnOverwriteExistingIndexFiles = False
        End If
    End If
    
    If KeyPressAbortProcess > 1 Then Exit Sub

    If lngFoldersParsed = 0 Then
        frmProgress.InitializeForm "Parsing folders: 0", 0, PROGRESS_LOOP_COUNT, False, True, True, MDIForm1
        frmProgress.InitializeSubtask "Counting subfolders: 0", 0, PROGRESS_SUB_LOOP_COUNT
        
        If blnShowMessages Then
            lngFolderCount = 0
            CountSubFolders strFolderStartPath, fso, lngFolderCount, FOLDER_COUNT_ABORT_SEARCH
            If lngFolderCount >= FOLDER_COUNT_ABORT_SEARCH Then
                ' Too many subfolders to count; just assume 10,000
                lngFolderCount = 10000
            End If
        Else
            ' Do not spend time counting the number of folders to process; just assume 10,000
            lngFolderCount = 10000
        End If
        
        frmProgress.InitializeForm "Parsing folders: 0", 0, lngFolderCount, True, False, True, MDIForm1
    
    End If
    frmProgress.UpdateCurrentTask "Parsing folders: " & lngFoldersParsed
    frmProgress.UpdateCurrentSubTask CompactPathString(strFolderStartPath, 45)
    
    frmProgress.UpdateProgressBar lngFoldersParsed
    
    Set objFolder = fso.GetFolder(strFolderStartPath)
    
    ReDim strFileNamesForFolder(0)
    blnResultsFolder = False
    
    
    ' See if this folder contains "Job*_log.txt"
    strThisFileName = Dir(fso.BuildPath(objFolder.Path, "Job*_log.txt"))
    
    blnProcessFolder = False
    If Len(strThisFileName) > 0 Then
        If LCase(Left(strThisFileName, 3)) = "job" And LCase(Right(strThisFileName, 8)) = "_log.txt" Then
            ' Valid folder
            blnResultsFolder = True
            
            blnProcessFolder = True
            If Not blnOverwriteExistingIndexFiles Then
                ' See if Index.html already exists
                If fso.FileExists(fso.BuildPath(objFolder.Path, HTML_INDEX_FILE_NAME)) Then
                    blnProcessFolder = False
                End If
            End If
            
            lngPeakMatchingTaskID = 0
            
            ' Extract out the MT tag database name and the peak matching task number
            '  from the folder name, if possible
            lngCharLoc = InStr(LCase(objFolder.Name), "_pm_")
            If lngCharLoc > 0 Then
                If IsNumeric(Mid(objFolder.Name, lngCharLoc + 4)) Then
                    lngPeakMatchingTaskID = Mid(objFolder.Name, lngCharLoc + 4)
                    
                    lngCharLoc = InStr(LCase(objFolder.Name), "_job")
                    If lngCharLoc > 0 Then
                        strMTDBName = Left(objFolder.Name, lngCharLoc - 1)
                    Else
                        strMTDBName = ""
                    End If
                End If
            End If
        End If
    End If
    
    If blnProcessFolder Then
        
        ' Cache the files present in this folder
        If objFolder.Files.Count > 0 Then
            intFileIndex = 0
            
            For Each objFile In objFolder.Files
                If intFileIndex > 0 Then
                    ReDim Preserve strFileNamesForFolder(intFileIndex)
                End If
                
                strFileNamesForFolder(intFileIndex) = objFile.Name
                intFileIndex = intFileIndex + 1
            Next objFile
                                    
            strBaseName = Left(strThisFileName, Len(strThisFileName) - 7)
            
            ' Need to parse the _log.txt file to determine the OutputFolderPath and the InputFilePath
            Set tsInput = fso.OpenTextFile(fso.BuildPath(objFolder.Path, strThisFileName), ForReading)
            
            If Not tsInput.AtEndOfStream Then
                strLogFileText = tsInput.ReadAll
            End If
            tsInput.Close
            
            strInputFilePath = Trim(GenerateAutoAnalysisFindText(strLogFileText, INPUT_FILE_HEADER_START, INPUT_FILE_HEADER_END))
            
            lngCharLoc = InStr(strInputFilePath, INPUT_FILE_HEADER_DELIMITER)
            If lngCharLoc >= 1 Then
                strInputFilePath = Trim(Mid(strInputFilePath, lngCharLoc + Len(INPUT_FILE_HEADER_DELIMITER)))
            End If
            
            strVersionOverride = Trim(GenerateAutoAnalysisFindText(strLogFileText, VERSION_HEADER_START, VERSION_HEADER_END))
            
            If Len(strVersionOverride) > 0 Then
                ' Determine the date from the line containing strVersionOverride
                
                strTemp = ""
                strDateOverride = ""
                lngCharLoc = 1
                Do
                    lngCharLoc = InStr(lngCharLoc, strLogFileText, strVersionOverride)
                    If lngCharLoc > 0 Then
                        strTemp = StrReverse(Left(strLogFileText, lngCharLoc - 1))
                        lngCharLoc = lngCharLoc + 1
                    End If
                Loop While lngCharLoc > 0
                
                lngCharLoc = InStr(strTemp, vbLf & vbCr)
                If lngCharLoc = 0 Then lngCharLoc = Len(strTemp)
                strTemp = StrReverse(Left(strTemp, lngCharLoc))
                
                If Len(strTemp) > 0 Then
                    strDateOverride = GenerateAutoAnalysisFindText(vbCrLf & strTemp, vbCrLf, " > ")
                    
                    If Len(strDateOverride) > 0 Then
                        ' Convert strDateOverride to a standard date
                        strDateOverride = Replace(strDateOverride, ".", ":")
                        strDateOverride = Replace(strDateOverride, " ", ":")
                        
On Error Resume Next
                        strDateParts = Split(strDateOverride, ":")
                        
                        If UBound(strDateParts) >= 5 Then
                            intYear = strDateParts(0)
                            intMonth = strDateParts(1)
                            intDay = strDateParts(2)
                            intHour = strDateParts(3)
                            intMinute = strDateParts(4)
                            intSecond = strDateParts(5)
                        Else
                            dtNow = Now()
                            intYear = Year(dtNow)
                            intMonth = Month(dtNow)
                            intDay = Day(dtNow)
                            intHour = Hour(dtNow)
                            intMinute = Minute(dtNow)
                            intSecond = Second(dtNow)
                        End If
                        strDateOverride = Format(DateSerial(intYear, intMonth, intDay) + TimeSerial(intHour, intMinute, intSecond), "MMMM dd, yyyy at hh:nn AMPM")
                        
On Error GoTo GenerateAutoAnalysisHtmlFileErrorHandler

                    End If
                    
                End If
            End If
            
            
            With udtAutoParams
                .ShowMessages = blnShowMessages
                .FilePaths.OutputFolderPath = objFolder.Path            ' This probably doesn't need to be populated here
                .FilePaths.InputFilePath = strInputFilePath
                
                If lngPeakMatchingTaskID > 0 And Len(strMTDBName) > 0 Then
                    .MTDBOverride.Enabled = True
                    .MTDBOverride.ServerName = ""                       ' We don't have any way of looking this up
                    .MTDBOverride.MTDBName = strMTDBName
                    .MTDBOverride.PeakMatchingTaskID = lngPeakMatchingTaskID
                Else
                    .MTDBOverride.Enabled = False
                End If
            End With
            
            With udtWorkingParams
                .GelOutputFolder = objFolder.Path
                .GelIndex = 0                                           ' Set to 0 since we do not have data in memory
                .GraphicOutputFileInfoCount = 0
                ReDim .GraphicOutputFileInfo(0)
                .TICPlotsStartRow = 4
                
                ' Look for each of the expected graphic file names
                
                strTargetName = strBaseName & "DataSearchedChargeView."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Data Searched", 1, 1
                End If
                
                strTargetName = strBaseName & "DataWithHitsNET."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Data With Matches", 1, 2
                End If
                
                strTargetName = strBaseName & "NETAlignmentSurface."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "NET Alignment Surface", 4, 1
                    .TICPlotsStartRow = 5
                End If
                
                strTargetName = strBaseName & "NETAlignmentResiduals."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "NET Alignment Residuals", 4, 2
                    .TICPlotsStartRow = 5
                End If
                
                strTargetName = strBaseName & frmTICAndBPIPlots.GetChromDescription(tbcTICFromCurrentDataIntensities, False) & "_Scan."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Total Ion Chromatogram (TIC)", udtWorkingParams.TICPlotsStartRow, 1
                End If
                
                strTargetName = strBaseName & frmTICAndBPIPlots.GetChromDescription(tbcBPIFromCurrentDataIntensities, False) & "_NET."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Base Peak Intensity (BPI) Chromatogram", udtWorkingParams.TICPlotsStartRow, 2
                End If
                
                strTargetName = strBaseName & frmTICAndBPIPlots.GetChromDescription(tbcDeisotopingIntensityThresholds, False) & "_Scan."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Deisotoping Intensity Thresholds", udtWorkingParams.TICPlotsStartRow + 1, 1
                End If
                
                strTargetName = strBaseName & frmTICAndBPIPlots.GetChromDescription(tbcDeisotopingPeakCounts, False) & "_Scan."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Deisotoping Peak Counts", udtWorkingParams.TICPlotsStartRow + 1, 2
                End If
                
                strTargetName = strBaseName & "MassErrors_BeforeRefinement."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Mass Errors Before Refinement", 2, 1
                End If
                
                strTargetName = strBaseName & "MassErrors."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "Mass Errors After Refinement", 2, 2
                End If
                
                strTargetName = strBaseName & "GANETErrors_BeforeRefinement."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "NET Errors Before Refinement", 3, 1
                End If
                
                strTargetName = strBaseName & "GANETErrors."
                If GenerateAutoAnalysisFindFile(strFileNamesForFolder(), strTargetName, strMatchingName) Then
                    AddNewOutputFileForHtml udtWorkingParams, strMatchingName, "NET Errors After Refinement", 3, 2
                End If
                
            End With
            
            ' We can now generate the actual Index.html file
            AutoAnalysisGenerateHTMLBrowsingFile udtWorkingParams, udtAutoParams, fso, strVersionOverride, strDateOverride
            
        End If
    End If
    
    ' Get the list of folders in strFolderStartPath
    For Each objSubFolder In objFolder.SubFolders
        If KeyPressAbortProcess > 1 Then Exit For
        
        If Len(strFolderPathTargetMask) > 0 Then
            
            ' Construct strFolderPathStripped
            ' Remove any leading \\ or C:\ values from strFolderStartPath
            strFolderPathStripped = objSubFolder.Path
            
            If Mid(strFolderPathStripped, 2, 1) = ":" Then
                strFolderPathStripped = Mid(strFolderPathStripped, 3)
            End If
            
            Do While Left(strFolderPathStripped, 1) = "\"
                strFolderPathStripped = Mid(strFolderPathStripped, 2)
            Loop
            
            If Right(strFolderPathTargetMask, 1) = "\" Then
                strFolderPathTargetMask = Left(strFolderPathTargetMask, Len(strFolderPathTargetMask) - 1)
            End If
            
            ' See if strFolderPathStripped is present in strFolderPathTargetMask
            If InStr(LCase(strFolderPathTargetMask), LCase(strFolderPathStripped)) > 0 Then
                blnProcessSubFolder = True
            Else
                blnProcessSubFolder = False
            End If
        Else
            blnProcessSubFolder = True
        End If
        
        If blnProcessSubFolder Then
            ' Call this function for each subfolder
            lngFoldersParsed = lngFoldersParsed + 1
            GenerateAutoAnalysisHtmlFiles objSubFolder.Path, strFolderPathTargetMask, blnOverwriteExistingIndexFiles, lngFoldersParsed, intRecursionLevel + 1, blnShowMessages
        End If
    Next objSubFolder
    
    ' Create an Index.html file, with links to each of the subfolders, and a link to the parent folder
    If Not blnResultsFolder Then
        AutoAnalysisGenerateHTMLSubfolderListFile objFolder, udtWorkingParams, udtAutoParams, fso
    End If
    
    If intRecursionLevel = 0 Then
        frmProgress.HideForm
        If blnShowMessages Then
            MsgBox "Processed " & Trim(lngFoldersParsed) & " folders.", vbInformation + vbOKOnly, "Done"
        End If
    End If
    
    Exit Sub

GenerateAutoAnalysisHtmlFileErrorHandler:
    If blnShowMessages Then
        MsgBox "Error in GenerateAutoAnalysisHtmlFiles: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        LogErrors Err.Number, "GenerateAutoAnalysisHtmlFiles", Err.Description
    End If
    Debug.Assert False
    
    frmProgress.HideForm
        
End Sub

Private Function GenerateAutoAnalysisFindText(ByRef strLogFileText As String, strStart As String, strEnd As String) As String
    Dim lngCharLoc1 As Long
    Dim lngCharLoc2 As Long
    
    Dim strMatchText As String
    
    lngCharLoc1 = 1
    Do
        lngCharLoc1 = InStr(lngCharLoc1, strLogFileText, strStart)
        If lngCharLoc1 > 0 Then
            lngCharLoc2 = InStr(lngCharLoc1, strLogFileText, strEnd)
            If lngCharLoc2 > 0 Then
                lngCharLoc1 = lngCharLoc1 + Len(strStart)
                strMatchText = Mid(strLogFileText, lngCharLoc1, lngCharLoc2 - lngCharLoc1)
            End If
            lngCharLoc1 = lngCharLoc1 + 1
        End If
    Loop While lngCharLoc1 > 0

    GenerateAutoAnalysisFindText = strMatchText
    
End Function

Private Function GenerateAutoAnalysisFindFile(ByRef strFileNamesForFolder() As String, ByVal strFileNameMatch As String, ByRef strMatchingName As String) As Boolean
    ' Looks for strFileNameMatch in the files in strFolderPath
    ' Returns True if found and populates strMatchingName with the full name
    ' Returns False if not found
    
    Dim intFileNameLength As Integer
    Dim blnMatchFound As Boolean
    
    Dim intFileIndex As Integer
    
    blnMatchFound = False
    strFileNameMatch = LCase(strFileNameMatch)
    intFileNameLength = Len(strFileNameMatch)
    
    For intFileIndex = 0 To UBound(strFileNamesForFolder)
        If LCase(Left(strFileNamesForFolder(intFileIndex), intFileNameLength)) = strFileNameMatch Then
            strMatchingName = strFileNamesForFolder(intFileIndex)
            blnMatchFound = True
            Exit For
        End If
    Next intFileIndex
    
    GenerateAutoAnalysisFindFile = blnMatchFound
    
End Function

Public Sub InitializeAutoAnalysisParameters(ByRef udtAutoParams As udtAutoAnalysisParametersType)
    With udtAutoParams
        With .FilePaths
            .DatasetFolder = ""
            .ResultsFolder = ""
            .InputFilePath = ""
            .OutputFolderPath = ""
            .IniFilePath = ""
            .LogFilePath = ""
        End With
        
        .ShowMessages = False
        .DatasetID = -1
        .JobNumber = -1
        .MDID = -1
        .AutoCloseFileWhenDone = False
        .GelIndexToForce = 0
        .FullyAutomatedPRISMMode = False
        .AutoDMSAnalysisManuallyInitiated = False
        .InvalidExportPassword = False
        .ErrorBits = 0
        
        With .MTDBOverride
            .Enabled = False
            .DBSchemaVersion = 1
            .ServerName = ""
            .MTDBName = ""
            .ConnectionString = ""
            .AMTsOnly = False
            .ConfirmedOnly = False
            .LockersOnly = False
            .LimitToPMTsFromDataset = False
            .MinimumPMTQualityScore = False
            .MinimumHighNormalizedScore = 0
            .MinimumHighDiscriminantScore = 0
            .MinimumPeptideProphetProbability = 0
            .ExperimentInclusionFilter = ""
            .ExperimentExclusionFilter = ""
            .InternalStandardExplicit = ""
            .NETValueType = 0
            .MTSubsetID = -1
            .ModList = "-1"
            .PeakMatchingTaskID = -1
        End With

        .ExitViperASAP = False
        .ExitViperReason = ""
        .RestartAfterExit = False
    End With
End Sub

Public Function LookupDBSearchModeIndex(strSearchMode As String) As dbsmDatabaseSearchModeConstants
    ' Returns the numerical DatabaseSearchMode Constant for the given search mode
    ' Returns -1 if no match
    
    Dim eSearchMode As dbsmDatabaseSearchModeConstants
    
    ' Note: Using LCase to avoid case conversion problems
    Select Case LCase(strSearchMode)
    Case LCase(AUTO_SEARCH_NONE): eSearchMode = dbsmNone
    Case LCase(AUTO_SEARCH_EXPORT_UMCS_ONLY): eSearchMode = dbsmExportUMCsOnly
    Case LCase(AUTO_SEARCH_ORGANISM_MTDB): eSearchMode = dbsmIndividualPeaks
    Case LCase(AUTO_SEARCH_UMC_MTDB): eSearchMode = dbsmIndividualPeaksInUMCsWithoutNET
    Case LCase(AUTO_SEARCH_UMC_HERETIC): eSearchMode = dbsmIndividualPeaksInUMCsWithNET             ' No longer supported (June 2004)
    Case LCase(AUTO_SEARCH_UMC_CONGLOMERATE): eSearchMode = dbsmConglomerateUMCsWithNET
    Case LCase(AUTO_SEARCH_UMC_HERETIC_PAIRED): eSearchMode = dbsmIndividualPeaksInUMCsPaired       ' No longer supported (June 2004)
    Case LCase(AUTO_SEARCH_UMC_HERETIC_UNPAIRED): eSearchMode = dbsmIndividualPeaksInUMCsUnpaired   ' No longer supported (June 2004)
    Case LCase(AUTO_SEARCH_UMC_CONGLOMERATE_PAIRED): eSearchMode = dbsmConglomerateUMCsPaired
    Case LCase(AUTO_SEARCH_UMC_CONGLOMERATE_UNPAIRED): eSearchMode = dbsmConglomerateUMCsUnpaired
    Case LCase(AUTO_SEARCH_UMC_CONGLOMERATE_LIGHT_PAIRS_PLUS_UNPAIRED): eSearchMode = dbsmConglomerateUMCsLightPairsPlusUnpaired
    Case LCase(AUTO_SEARCH_PAIRS_N14N15): eSearchMode = dbsmPairsN14N15                             ' No longer supported (July 2004)
    Case LCase(AUTO_SEARCH_PAIRS_N14N15_CONGLOMERATEMASS): eSearchMode = dbsmPairsN14N15ConglomerateMass
    Case LCase(AUTO_SEARCH_PAIRS_ICAT): eSearchMode = dbsmPairsICAT
    Case LCase(AUTO_SEARCH_PAIRS_PEO): eSearchMode = dbsmPairsPEO
    Case Else
        Debug.Assert False
        eSearchMode = dbsmNone
    End Select
    
    LookupDBSearchModeIndex = eSearchMode
End Function

Private Function LookupErrorBitDescription(lngErrorBits As Long) As String
    Dim strMessage As String
    
    strMessage = ""
    
    CheckMessageBit lngErrorBits, DATAFILE_LOAD_ERROR_BIT, strMessage, "Datafile (" & KNOWN_FILE_EXTENSIONS_WITH_GEL & ") load error"
    CheckMessageBit lngErrorBits, INIFILE_LOAD_ERROR_BIT, strMessage, "Ini file load error"
    CheckMessageBit lngErrorBits, DATABASE_ERROR_BIT, strMessage, "Database MT tag retrieval error"
    CheckMessageBit lngErrorBits, UMC_COUNT_ERROR_BIT, strMessage, "LC-MS Feature search error"
    CheckMessageBit lngErrorBits, GANET_ERROR_BIT, strMessage, "NET adjustment error"
    CheckMessageBit lngErrorBits, SEARCH_ERROR_BIT, strMessage, "DB search error"
    CheckMessageBit lngErrorBits, EXPORTRESULTS_ERROR_BIT, strMessage, "Export results to DB error"
    CheckMessageBit lngErrorBits, TOLERANCE_REFINEMENT_ERROR_BIT, strMessage, "Tolerance refinement error"
    CheckMessageBit lngErrorBits, SAVE_GRAPHIC_ERROR_BIT, strMessage, "Save 2D graphic error"
    CheckMessageBit lngErrorBits, SAVE_ERROR_DISTRIBUTION_ERROR_BIT, strMessage, "Save error distribution error"
    CheckMessageBit lngErrorBits, SAVE_CHROMATOGRAM_ERROR_BIT, strMessage, "Save chromatogram error"
    CheckMessageBit lngErrorBits, MASS_TAGS_NULL_COUNTS_HIGH_ERROR_BIT, strMessage, "Number of MT tags with null mass or null NET values is abnormally high"
    
    LookupErrorBitDescription = strMessage
    
End Function

Private Sub LookupMatchingUMCStats(ByVal lngGelIndex As Long, ByRef lngUMCCount As Long, ByRef lngUMCCountWithHits As Long, ByRef lngUniqueMassTagCount As Long)
    ' Step through the LC-MS Features for lngGelIndex and count the number with hits
    ' Also keep track of the number of unique MT tag hits
    ' Using a dictionary object as a hashtable
    
    Dim htMTHitList As Dictionary
    Dim lngUMCIndex As Long
    Dim lngMatchIndex As Long
    
    Dim udtUMCList() As udtUMCMassTagMatchStats
    Dim lngUMCListCount As Long
    Dim lngUMCListCountDimmed As Long
    
On Error GoTo LookupMatchingUMCStatsErrorHandler

    ' If there isn't data in memory, then the following will generate an error and jump to the error handler
    With GelUMC(lngGelIndex)
        lngUMCCount = .UMCCnt
        lngUMCCountWithHits = 0
        lngUniqueMassTagCount = 0
        
        lngUMCListCount = 0
        lngUMCListCountDimmed = 100
        ReDim udtUMCList(lngUMCListCountDimmed)
        
        Set htMTHitList = New Dictionary
        htMTHitList.RemoveAll
        
        For lngUMCIndex = 0 To .UMCCnt - 1
            ' We don't need to cache the results for all LC-MS Features in memory; thus, reset this to zero for each UMC examined
            
            lngUMCListCount = 0
            ExtractMTHitsFromUMCMembers lngGelIndex, lngUMCIndex, False, udtUMCList(), lngUMCListCount, lngUMCListCountDimmed, False, False
            
            If lngUMCListCount > 0 Then
                If udtUMCList(0).IDIndex <> 0 Then
                    lngUMCCountWithHits = lngUMCCountWithHits + 1
                    
                    For lngMatchIndex = 0 To lngUMCListCount - 1
                        ' Note that udtUMCList(lngMatchIndex).IDIndex contains the actual MT tag ID
                        If Not htMTHitList.Exists(udtUMCList(lngMatchIndex).IDIndex) Then
                            htMTHitList.add udtUMCList(lngMatchIndex).IDIndex, 1
                        End If
                    Next lngMatchIndex
                    
                End If
            End If
        Next lngUMCIndex
        
        lngUniqueMassTagCount = htMTHitList.Count
    End With
    
    Exit Sub

LookupMatchingUMCStatsErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error examining LC-MS Features in memory to determine number with matches: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        LogErrors Err.Number, "LookupMatchingUMCStats"
    End If
    
End Sub

Private Function LookupMatchStatsForPeakMatchingTask(ByVal strServerName As String, ByVal strMTDBName As String, ByVal lngPeakMatchingTaskID As Long, ByRef lngNonUniqueHitsCount As Long, ByRef lngUMCCount As Long, ByRef lngUMCCountWithHits As Long, ByRef lngUniqueMassTagCount As Long) As Boolean
    ' Call GetPeakMatchingTaskResultStats in database strMTDBName
    ' Returns True if success, false if an error
    
    Dim cnnConnection As ADODB.Connection
    
    Dim cmdGetPMStats As New ADODB.Command
    
    Dim prmPeakMatchingTaskID As New ADODB.Parameter
    Dim prmJobNumber  As New ADODB.Parameter
    Dim prmNonUniqueHitsCount As New ADODB.Parameter
    Dim prmUMCCount As New ADODB.Parameter
    Dim prmUMCCountWithHits As New ADODB.Parameter
    Dim prmUniqueMassTagHitCount As New ADODB.Parameter
    Dim prmMessage As New ADODB.Parameter
    
    Dim strConnectionString As String
    Dim strSPName As String
    
    Dim blnSuccess As Boolean
    
On Error GoTo LookupMatchStatsForPeakMatchingTaskErrorHandler
    
    strSPName = glbPreferencesExpanded.MTSConnectionInfo.spGetPMResultStats
    
    If Len(strSPName) = 0 Then
        Debug.Assert False
        strSPName = "GetPeakMatchingTaskResultStats"
    End If
    
    strConnectionString = glbPreferencesExpanded.MTSConnectionInfo.ConnectionString
    
    ' Update the connection string sith strServerName and strMTDBName
    strConnectionString = ConstructConnectionString(strServerName, strMTDBName, strConnectionString)
    
    If Not EstablishConnection(cnnConnection, strConnectionString) Then
        LookupMatchStatsForPeakMatchingTask = False
        Exit Function
    End If
    
    ' Initialize the SP parameters
    InitializeSPCommand cmdGetPMStats, cnnConnection, strSPName
    
    Set prmPeakMatchingTaskID = cmdGetPMStats.CreateParameter("PeakMatchingTaskID", adInteger, adParamInput)
    prmPeakMatchingTaskID.Value = lngPeakMatchingTaskID
    cmdGetPMStats.Parameters.Append prmPeakMatchingTaskID

    Set prmJobNumber = cmdGetPMStats.CreateParameter("JobNumber", adInteger, adParamOutput)
    cmdGetPMStats.Parameters.Append prmJobNumber

    Set prmNonUniqueHitsCount = cmdGetPMStats.CreateParameter("NonUniqueHitsCount", adInteger, adParamOutput)
    cmdGetPMStats.Parameters.Append prmNonUniqueHitsCount

    Set prmUMCCount = cmdGetPMStats.CreateParameter("UMCCount", adInteger, adParamOutput)
    cmdGetPMStats.Parameters.Append prmUMCCount

    Set prmUMCCountWithHits = cmdGetPMStats.CreateParameter("UMCCountWithHits", adInteger, adParamOutput)
    cmdGetPMStats.Parameters.Append prmUMCCountWithHits

    Set prmUniqueMassTagHitCount = cmdGetPMStats.CreateParameter("UniqueMassTagHitCount", adInteger, adParamOutput)
    cmdGetPMStats.Parameters.Append prmUniqueMassTagHitCount

    Set prmMessage = cmdGetPMStats.CreateParameter("message", adVarChar, adParamOutput, 512)
    cmdGetPMStats.Parameters.Append prmMessage


    ' Execute the SP
    cmdGetPMStats.Execute
    
    If CLngSafe(prmJobNumber.Value) > 0 Then
        lngNonUniqueHitsCount = FixNullLng(prmNonUniqueHitsCount.Value, -1)
        lngUMCCount = FixNullLng(prmUMCCount.Value, -1)
        lngUMCCountWithHits = FixNullLng(prmUMCCountWithHits.Value, -1)
        lngUniqueMassTagCount = FixNullLng(prmUniqueMassTagHitCount.Value, -1)
        blnSuccess = True
    Else
        blnSuccess = False
    End If
    
    If lngNonUniqueHitsCount <= 0 And lngUMCCount <= 0 Then
        blnSuccess = False
    End If
    
    LookupMatchStatsForPeakMatchingTask = blnSuccess
    Exit Function

LookupMatchStatsForPeakMatchingTaskErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error connecting to database " & strMTDBName & " and calling SP " & strSPName & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        LogErrors Err.Number, "LookupMatchStatsForPeakMatchingTask"
    End If
    LookupMatchStatsForPeakMatchingTask = False
    
End Function

Private Function LookupWarningBitDescription(lngWarningBits As Long) As String
    Dim strMessage As String
    
    strMessage = ""
    
    CheckMessageBit lngWarningBits, UMC_SEARCH_ABORTED_WARNING_BIT, strMessage, "LC-MS Feature Search aborted"
    CheckMessageBit lngWarningBits, NET_ADJUSTMENT_SKIPPED_WARNING_BIT, strMessage, "NET Adjustment skipped"
    CheckMessageBit lngWarningBits, GANET_SLOPE_WARNING_BIT, strMessage, "NET slope outside expected range"
    CheckMessageBit lngWarningBits, GANET_INTERCEPT_WARNING_BIT, strMessage, "NET intercept outside expected range"
    CheckMessageBit lngWarningBits, NET_ADJUSTMENT_LOW_ID_COUNT_WARNING_BIT, strMessage, "NET Adjustment low DB match count"
    CheckMessageBit lngWarningBits, TOLERANCE_REFINEMENT_WARNING_PEAK_NOT_FOUND_BIT, strMessage, "Tolerance refinement warning: peak not found"
    CheckMessageBit lngWarningBits, TOLERANCE_REFINEMENT_WARNING_BIT_PEAK_TOO_WIDE, strMessage, "Tolerance refinement warning: peak too wide"
    CheckMessageBit lngWarningBits, PICTURE_FORMAT_WARNING_BIT, strMessage, "Invalid picture format specified"
    CheckMessageBit lngWarningBits, INVALID_EXPORT_OPTION_WARNING_BIT, strMessage, "Invalid export option specified"
    CheckMessageBit lngWarningBits, NO_DATABASE_HITS_WARNING_BIT, strMessage, "No database hits"
    CheckMessageBit lngWarningBits, PATH_FILE_ERROR_WARNING_BIT, strMessage, "Path/file warning"
    CheckMessageBit lngWarningBits, NO_PAIRS_WARNING_BIT, strMessage, "No pairs found"
    CheckMessageBit lngWarningBits, MISCELLANEOUS_MESSAGE_WARNING_BIT, strMessage, "Miscellaneous Warning (see log file)"
    
    LookupWarningBitDescription = strMessage
    
End Function

Public Function QueryUserForExportToDBPassword(Optional strMessage As String = "Please enter the password for exporting results to the database:", Optional blnUsePWForm As Boolean = True) As Boolean
    ' Returns True if the user enters the correct password
    ' This is used to prevent the average user from exporting results to the database
    
    ' Yes, this is a hard-coded password
    ' No, it's not secure
    ' It is entered as astericks if we're able to show objPWForm; we can't show objPWForm if
    '   the calling form was shown modally
    
    Dim blnSuccess As Boolean
    Dim strResponse As String
    
    Dim objPWForm As frmPassword
    Set objPWForm = New frmPassword
    
On Error GoTo QueryUserForPWErrorHandler
    
    If blnUsePWForm Then
        objPWForm.Initialize strMessage, EXPORT_TO_DB_PASSWORD
        
        objPWForm.Show
        Do
            DoEvents
            Sleep 50
        Loop While Not objPWForm.ProceedAndCloseForm()
    
        blnSuccess = objPWForm.PasswordWasValidated()
    Else
    
QueryUserUsingInputBox:
        strResponse = InputBox(strMessage, "Password", "")
        If strResponse = EXPORT_TO_DB_PASSWORD Then
            blnSuccess = True
        Else
            blnSuccess = False
        End If
    
    End If
    
QueryUserForPWCleanup:
    On Error Resume Next
    
    If blnUsePWForm Then
        If Not objPWForm Is Nothing Then
            Unload objPWForm
            Set objPWForm = Nothing
        End If
    End If
    
    QueryUserForExportToDBPassword = blnSuccess
    Exit Function
    
QueryUserForPWErrorHandler:
    ' This shouldn't happen
    Debug.Assert False
    LogErrors Err.Number, "QueryUserForExportToDBPassword"
    
    objPWForm.Visible = False
    Resume QueryUserUsingInputBox:
    
End Function

Private Sub ReplaceOutputFileForHtml(ByRef udtWorkingParams As udtAutoAnalysisWorkingParamsType, ByVal strFileName As String, ByVal strExistingDescriptionToMatch As String, ByVal strDescription As String, ByVal intTableRow As Integer, ByVal intTableColumn As Integer, Optional ByVal intWidth As Integer = 450)
    ' Look for strExistingDescriptionToMatch in .GraphicOutputFileInfo()
    ' If found, replace with strFileName and strDescription
    ' If not found, then add a new entry to .GraphicOutputFileInfo()
    
    Dim intIndex As Integer
    Dim blnMatchFound As Boolean
    
    blnMatchFound = False
    For intIndex = 0 To udtWorkingParams.GraphicOutputFileInfoCount - 1
        If udtWorkingParams.GraphicOutputFileInfo(intIndex).Description = strExistingDescriptionToMatch Then
            With udtWorkingParams.GraphicOutputFileInfo(intIndex)
                .FileName = strFileName
                .Description = strDescription
            End With
            blnMatchFound = True
            Exit For
        End If
    Next intIndex
    
    If Not blnMatchFound Then
        AddNewOutputFileForHtml udtWorkingParams, strFileName, strDescription, intTableRow, intTableColumn, intWidth
    End If
End Sub


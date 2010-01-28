Attribute VB_Name = "MonroeGlobals"
Option Explicit
 
' This portion of VIPER is by Matthew Monroe, PNNL
' Started December 23, 2002
'
' Note: The VIPER acronym means: Visual Inspection of Peak/Elution Relationships
' In that past, it also meant: Visualization, Protein/Peptide Identification, Expression Analysis Tool
' It was previously known as Contemporary 2D Displays and LaV2DG
  
Public Const SELECT_MT_DATABASE_MASSTAGS_ACCESS_CAPTION = "Select MT Database For Out"          ' Yes, this ends in "For Out" and not "For Output"
Public Const INI_FILENAME = "VIPERSettings.ini"
Public Const RECENT_DB_INI_FILENAME = "VIPERRecentDB.ini"


Public Const APP_BUILD_DATE As String = "January 13, 2010"

Public Const PRISM_AUTOMATION_CONNECTION_STRING_DEFAULT = "Provider=sqloledb;Data Source=pogo;Initial Catalog=PRISM_RPT;User ID=mtuser;Password=mt4fun"
Public Const PRISM_AUTOMATION_SP_REQUEST_TASK_DEFAULT = "RequestPeakMatchingTaskMaster"
Public Const PRISM_AUTOMATION_SP_SET_COMPLETE_DEFAULT = "SetPeakMatchingTaskCompleteMaster"
Public Const PRISM_AUTOMATION_SP_RESTART_TASK_DEFAULT = "SetPeakMatchingTaskToRestartMaster"
Public Const PRISM_AUTOMATION_SP_POST_LOG_ENTRY_DEFAULT = "PostLogEntry"

' Note: This list is used by the ExtractTimeDomainSignalFromPEK function to check for
'       processed PEK/CSV/mzXML/mzData file name extensions.  Thus, if a file is found with any of these
'       extensions (except for plain .pek) then the function will know that the PEK/CSV/mzXML/mzData file
'       has been processed.  In that case, it will look for a plain .PEK file in the same folder
'       This should be a comma delimited list
'
' This list is also used during automated analysis to automatically find the best file to load
Public Const DEFAULT_PEK_FILE_EXTENSION_ORDER As String = LCMS_FEATURES_FILE_SUFFIX & ", " & CSV_ISOS_IC_FILE_SUFFIX & ", " & CSV_ISOS_FILE_SUFFIX & ", .mzxml, .mzdata, mzxml.xml, mzdata.xml, _ic.pek, _s.pek, .pek, DeCal.pek-3, .pek-3"

Public Const KNOWN_FILE_EXTENSIONS As String = ".Pek, .CSV, .mzXML, mzxml.xml, .mzData, or mzdata.xml"
Public Const KNOWN_FILE_EXTENSIONS_WITH_GEL As String = ".Pek, .CSV, .mzXML, mzxml.xml, .mzData, mzdata.xml, or .Gel"

Public Const UMC_NET_ADJ_UMCs_IN_TOLERANCE = "LC-MS Features in tolerance"
Public Const UMC_NET_ADJ_UMCs_WITH_DB_HITS = "LC-MS Feature Count with DB hits"
Public Const UMC_NET_ADJ_ITERATION_COUNT = "Iteration Count"

Public Const SEARCH_MASS_TOL_DETERMINED = "Search mass tolerance determined"
Public Const SEARCH_NET_TOL_DETERMINED = "Search NET tolerance determined"

Public Const MASS_CALIBRATION_PEAK_STATS_START = "mass calibration peak stats"
Public Const MASS_CALIBRATION_PEAK_STATS_END = "(height, width, center, S/N)"

Public Const NET_TOL_PEAK_STATS_START = "NET tolerance peak stats"
Public Const NET_TOL_PEAK_STATS_END = "(height, width, center, S/N)"

Public Const SEPARATION_CHARACTER_TAB_STRING = "<TAB>"

Public Const UMC_INDICATOR_BIT_SPLIT_UMC = 16
Public Const UMC_INDICATOR_BIT_USED_FOR_NET_ADJ = 32                    ' LC-MS Features used for NET adjustment have this bit turned on
Public Const UMC_INDICATOR_BIT_LOWSEGMENTCOUNT_ADDITION = 64            ' LC-MS Features added due to low segment count values have this bit turned on
Public Const UMC_INDICATOR_BIT_NET_ADJ_DB_HIT = 128                     ' LC-MS Features used for NET adjustment that matched one or more MT tags in the database
Public Const UMC_INDICATOR_BIT_NET_ADJ_LOCKER_HIT = 256                 ' LC-MS Features used for Net adjustment that matched one of the Internal Standards (aka NET adjustment lockers)

Public Const GEL_DATA_STATUS_BIT_IREPORT = 2                        ' When the gel data contains IReport data, this this bit is turned on
Public Const GEL_DATA_STATUS_BIT_ISOTOPE_LABEL_TAG = 4              ' When the gel data contains Isotope LabelTag info (field .IsotopeLabel), this this bit is turned on
Public Const GEL_DATA_STATUS_BIT_ADDED_MONOPLUSMINUS4_DATA = 8      ' When the gel data contains data loaded from _Pairs_isos.csv files and MonoPlus4 or MonoMinus4 data was added, then this this bit is turned on
Public Const GEL_DATA_STATUS_BIT_IMS_DATA = 16                      ' When the gel data contains IMS data, then this bit is turned on
Public Const GEL_DATA_STATUS_BIT_LCMSFEATURES_DATA = 32             ' When the gel data contains data loaded from a _LCMSFeatures.txt file, then this bit is turned on

Public Enum natNETTypeConstants
    natGeneric = 0
    natTICNET = 1
    natGANET = 2
End Enum
  
Public Enum ddmDataDisplayModeConstants
    ddmIonsOnly = 0
    ddmUMCsOnly = 1
    ddmIonsAndUMCs = 2
End Enum

Public Enum nmNotationModeConstants
    nmDecimal = 0
    nmScientific = 1
End Enum

Public Const MENU_MODE_COUNT = 4
Public Enum mmMenuModeConstants
    mmBasic = 0
    mmDBNoPairs = 1
    mmDBWithPairs = 2
    mmFull = 3
End Enum

Public Enum fsFileSaveModeConstants
    fsUnknown = 0
    fsNoExtended = 1
    fsIncludeExtended = 2
    fsLegacy = 3
End Enum

Public Const LC_TIME_BASED_CHROM_COUNT = 8
Public Const IMS_TIME_BASED_CHROM_COUNT = 2
Public Enum tbcTICAndBPIConstants
    tbcTICFromCurrentDataIntensities = 0
    tbcBPIFromCurrentDataIntensities = 1
    tbcTICFromTimeDomain = 2
    tbcTICFromCurrentDataPointCounts = 3
    tbcTICFromRawData = 4
    tbcBPIFromRawData = 5
    tbcDeisotopingIntensityThresholds = 6
    tbcDeisotopingPeakCounts = 7
    tbcIMSTICFromCurrentDataIntensities = 8
    tbcIMSBPIFromCurrentDataIntensities = 9
End Enum

Public Enum itcIMSTICConstants
    itcTICFromCurrentDataIntensities = 0
    itcBPIFromCurrentDataIntensities = 1
End Enum

Public Const MASS_TOLERANCE_REFINEMENT_METHOD_COUNT = 5
Public Enum mtrMassToleranceRefinementConstants
    mtrExpectationMaximization = 0                      ' Use Expectation Maximization
    mtrMassErrorPlotWidthAtPctOfMax = 1                 ' Use width at .PercentageOfMaxForFindingWidth
    mtrMedianUMCMassStDev = 2
    mtrMaximumUMCMassStDev = 3
    mtrMedianUMCMassWidth = 4
    mtrMaximumUMCMassWidth = 5
End Enum

Public Enum wscWindowSizeConstants
    wscSizeForPowerpoint = 0
    wscsize640by480 = 1
    wscSize800by600 = 2
    wscSize1024by768 = 3
    wscSize1280by1024 = 4
End Enum

Public Enum punaPairsUMCNetAdjustmentConstants
    punaPairedAndUnpaired = 0
    punaUnpairedOnly = 1
    punaUnpairedPlusPairedLight = 2
    punaPairedAll = 3
    punaPairedLight = 4
    punaPairedHeavy = 5
End Enum

Public Enum stSearchTypeConstants
    stNotDefined = 0
    stStandardIndividual = 1
    stPairsN14N15 = 2
    stPairsICAT = 3
    stPairsPEO = 4
    stPairsPhIAT = 5
    stPairsLysC12C13 = 6
    stPairsPEON14N15 = 7
    stLabeledICATD0 = 8
    stLabeledICATD8 = 9
    stLabeledPEO = 10
    stLabeledPhIAT = 11
    stPairsO16O18 = 12
End Enum

Public Enum nvtNetValueTypeConstants
    nvtGANET = 0
    nvtPNET = 1
    '' nvtTheoreticalNET = 2               ' No longer supported (March 2006)
End Enum

Public Enum susgSplitUMCsScanGapConstants
    susgIgnoreScanGaps = 0
    susgSplitIfMassDifference = 1
    susgAlwaysSplit = 2
End Enum

Public Enum aewAverageERsWeightingModeConstants
    aewNoWeighting = 0
    aewAbundance = 1
    aewMemberCounts = 2
End Enum

Public Enum ifmInputFileModeConstants
    ifmPEKFile = 0
    ifmCSVFile = 1
    ifmmzXMLFile = 2                        ' i.e. InputFile.mzXml
    ifmmzXMLFileWithXMLExtension = 3        ' i.e. InputFile_mzXml.xml
    ifmmzDataFile = 4                       ' i.e. InputFile.mzData
    ifmmzDataFileWithXMLExtension = 5       ' i.e. InputFile_mzData.xml
    ifmGelFile = 6
    ifmDelimitedTextFile = 7
End Enum

Public Enum pftPictureFileTypeConstants
    pftUnknown = 0
    pftPNG = 1
    pftJPG = 2
    pftWMF = 3
    pftEMF = 4
    pftBMP = 5
End Enum

Public Enum plmPointsLoadModeConstants
    plmLoadAllPoints = 0
    plmLoadMappedPointsOnly = 1
    plmLoadOnePointPerLCMSFeature = 2
End Enum

' Up to 4 different auto searches can be performed
Public Const MAX_AUTO_SEARCH_MODE_COUNT = 4

' Auto Analysis Search Mode Constants
Public Const AUTO_SEARCH_NONE = "None"
Public Const AUTO_SEARCH_EXPORT_UMCS_ONLY = "ExportUMCsOnly"                ' No DB Search, simply export the LC-MS Features
Public Const AUTO_SEARCH_ORGANISM_MTDB = "IndividualPeaks"
Public Const AUTO_SEARCH_UMC_MTDB = "IndividualPeaksInUMCsWithoutNET"
Public Const AUTO_SEARCH_UMC_HERETIC = "IndividualPeaksInUMCsWithNET"                       ' No longer supported (June 2004)
Public Const AUTO_SEARCH_UMC_CONGLOMERATE = "ConglomerateUMCsWithNET"

Public Const AUTO_SEARCH_UMC_HERETIC_PAIRED = "IndividualPeaksInUMCsWithNETPaired"          ' No longer supported (June 2004)
Public Const AUTO_SEARCH_UMC_HERETIC_UNPAIRED = "IndividualPeaksInUMCsWithNETUnpaired"      ' No longer supported (June 2004)
Public Const AUTO_SEARCH_UMC_CONGLOMERATE_PAIRED = "ConglomerateUMCsWithNETPaired"
Public Const AUTO_SEARCH_UMC_CONGLOMERATE_UNPAIRED = "ConglomerateUMCsWithNETUnpaired"
Public Const AUTO_SEARCH_UMC_CONGLOMERATE_LIGHT_PAIRS_PLUS_UNPAIRED = "ConglomerateUMCsWithLightPairsPlusUnpaired"

Public Const AUTO_SEARCH_PAIRS_N14N15 = "DBSearchPairsN14N15"                                    ' No longer supported (July 2004)
Public Const AUTO_SEARCH_PAIRS_N14N15_CONGLOMERATEMASS = "DBSearchPairsN14N15ConglomerateMass"

Public Const AUTO_SEARCH_PAIRS_ICAT = "DBSearchPairsICAT"
Public Const AUTO_SEARCH_PAIRS_PEO = "DBSearchPairsPEO"


' Automatic pair finding constants
Public Const AUTO_FIND_PAIRS_NONE = "None"
Public Const AUTO_FIND_PAIRS_DELTA = "DeltaPairs"
Public Const AUTO_FIND_PAIRS_LABEL = "LabelPairs"

' Auto Analysis UMC Mode Constants
Public Const UMC_SEARCH_MODE_SETTING_TEXT = "LC-MS Feature Search Mode"
Public Const AUTO_ANALYSIS_UMCListType2002 = "UMCListType2002"      ' Uses frmUMCWithAutoRefine; No longer supported (July 2004)
Public Const AUTO_ANALYSIS_UMC2003 = "UMC2003"      ' Uses frmUMCSimple
Public Const AUTO_ANALYSIS_UMCIonNet = "UMCIonNet"  ' Uses frmUMCIonNet

' Add Match Making Entry: T_Match_Making_Description.MD_State values
Public Const MMD_STATE_NEW = 1
Public Const MMD_STATE_OK = 2
'Public Const MMD_STATE_UPDATED = 5

' Export UMC Matches: T_FTICR_UMC_ResultDetails.Match_State values
'Public Const MATCH_STATE_NEW = 1
Public Const MATCH_STATE_NO_HIT = 5
Public Const MATCH_STATE_HIT = 6


Public Type udtCollectionArrayType
    Value As String
    Name As String
End Type

Public Const DBSTUFF_COLLECTION_COUNT_MAX = 50
Public Type udtMTDBInfoType
    ConnectionString As String
    DBStatus As Long
    DBStuffArray(DBSTUFF_COLLECTION_COUNT_MAX) As udtCollectionArrayType        ' 0-based array, though collections are 1-based
    DBStuffArrayCount As Long
End Type

Public Const MAX_RECENT_FILE_COUNT = 8
Public Type udtRecentFileDetailsType
    ShortenedFilePath As String
    FullFilePath As String
End Type

Public Type udtRecentFilesType
    FileCount As Integer
    Files(MAX_RECENT_FILE_COUNT) As udtRecentFileDetailsType                ' 0-based array, though entryies in Ini file start at RecentFile1
End Type

Public Type udtGelAnalysisInfoType
    ValidAnalysisDataPresent As Boolean
    Analysis_Tool As String
    Created As String
    Dataset As String
    Dataset_Folder As String
    Dataset_ID As Long
    Desc_DataFolder As String
    Desc_Type As String
    Duration As Long
    Experiment As String
    GANET_Fit As Double
    GANET_Intercept As Double
    GANET_Slope As Double
    Instrument_Class As String
    Job As Long
    MD_Date As String
    MD_file As String
    MD_Parameters As String
    MD_Reference_Job As Long
    MD_State As Long
    MD_Type As Long
    MTDB As udtMTDBInfoType
    NET_Intercept As Double
    NET_Slope As Double
    NET_TICFit As Double
    Organism As String
    Organism_DB_Name As String
    Parameter_File_Name As String
    ProcessingType As Long
    Results_Folder As String
    Settings_File_Name As String
    STATE As Long
    Storage_Path As String
    Total_Scans As Long
    Vol_Client As String
    Vol_Server As String
End Type

Public Type udtDBSettingsType
    AnalysisInfo As udtGelAnalysisInfoType
    IsDeleted As Boolean            ' When true, means that none of the other values is valid
    ConnectionString As String
    DatabaseName As String          ' This is determined by parsing ConnectionString and is included to simplify keeping track of recent database connections
    DBSchemaVersion As Single
    
    AMTsOnly As Boolean
    ConfirmedOnly As Boolean            ' Only used in Schema Version 1
    LockersOnly As Boolean              ' Only used in Schema Version 1
    LimitToPMTsFromDataset As Boolean   ' Only used in Schema Version 2
    
    MinimumHighNormalizedScore As Single
    MinimumHighDiscriminantScore As Single
    MinimumPeptideProphetProbability As Single
    MinimumPMTQualityScore As Single
    
    ExperimentInclusionFilter As String
    ExperimentExclusionFilter As String
    InternalStandardExplicit As String
    
    NETValueType As Integer         ' Actually type nvtNetValueTypeConstants
    MassTagSubsetID As Long         ' Only used in Schema Version 1
    ModificationList As String
    SelectedMassTagCount As Long
End Type

Public Type udtIonToUMCLookupEntryType
    UMCCount As Long
    UMCs() As Long
End Type

' This data type is used to hold various lookup values
'   CSUMCs() and IsoUMCs() are used to hold the indices of the LC-MS Features that each of the ions are members of
'     The number of data points in CSUMCs() and IsoUMCs() should be equal to GelData().IsoLines
'   AdjacentScanNumberNext() and AdjacentScanNumberPrevious() contain the next or previous possible scan number, for all scans with scan numbers between 0 and 1 million
Public Type udtGelDataLookupIndexType
    CSUMCs() As udtIonToUMCLookupEntryType
    IsoUMCs() As udtIonToUMCLookupEntryType
    AdjacentScanNumberNext() As Long
    AdjacentScanNumberPrevious() As Long
    ScanNumberRelativeIndex() As Long
End Type

' The following structures are used to keep track of the Proteins in memory, along with
'  the loaded ions that match them
' Note that an ion can match more than one ORF
Public Type udtORFDefinitionType
    Organism As String
    MTDBName As String
    MTDBConnectionString As String
    ORFDBName As String
    ORFDBConnectionString As String
    DateDataObtained As String              ' Date the data was obtained from the database
    DataParsedCompletely As Boolean
    OtherInfo As String                     ' Reserved for future expansion
                                            ' Contains the string UMCCountLastRecordIonMatchCall=   which keeps track of the number of LC-MS Features present the last time RecordIonMatchesInORFMassTags was called
End Type

Public Type udtIonMatchType
    IonDataIndex As Long                ' Pointer to GelData().IsoData(IonDataIndex)
    MassTagIndex As Long                ' Pointer to GelOrfMassTags().Orfs().MassTags(MassTagIndex)
End Type

Public Type udtUMCMatchType
    UMCDataIndex As Long                ' Pointer to GelUMC().UMCs(UMCDataIndex)
    MassTagIndex As Long                ' Pointer to GelOrfMassTags().Orfs().MassTags(MassTagIndex)
    ClassMemberIndex As Long
End Type

Public Type udtORFDataType
    Reference As String         ' String describing ORF, for example ORF00002
    RefID As Long               ' ID of the ORF in the MassTag Database (field Ref_ID)  - I'll treat this as the "Master" ID number for each ORF
    ORFID As Long               ' ID of the ORF in the ORF Database (field Orf_ID)      - needed to perform a virtual "Join" between the data in the MT tag DB and an ORF database
    
    Description As String       ' Fasta description (or something similar)
    AlternateRef As String      ' Can be used to record an alternate ORF reference string (symbol)
    
    Classification As String    ' Would contain classification info from ORF Database, but not sure how to relate it to the T_ORF table
    ClassificationID As Long    ' ID representing Classification
    Role As String              ' From ORF Database
    SubRole As String           ' From ORF Database
    
    LocationStart As Long       ' From ORF Database
    LocationStop As Long        ' From ORF Database
    Strand As String            ' From ORF Database
    ReadingFrame As Integer     ' From ORF Database
    
    pi As Double                ' From ORF Database (Isoelectric point)
    CAI As Double               ' From ORF Database
    MassMonoisotopic As Double  ' From ORF Database or the MT tag database (equivalent values)
    MassAverage As Double       ' From ORF Database
    MolecularFormula As String  ' From ORF Database
    Sequence As String          ' From ORF Database or the MT tag database (equivalent values)
    
    TrypticFragmentCount As Long
    
    IonMatchCount As Long
    IonMatches() As udtIonMatchType
    
    UMCMatchCount As Long
    UMCMatches() As udtUMCMatchType
    
    OtherInfo As String                     ' Reserved for future expansion
End Type

Public Type udtORFListType
    Definition As udtORFDefinitionType
    ORFCount As Long
    Orfs() As udtORFDataType
End Type


' The following structures are used to keep track of the MT tags for each ORF in a MT tag database
Public Type udtORFMassTagsDefinitionType
    IncludesPMTs As Boolean
    IncludesTheoreticalTrypticMassTags As Boolean               ' Set true when includes theoretical tryptic MT tags; kept at true even if user cancels partway through computation of theoretical MT tags
    TheoreticalTrypticMassTagsSuccessfullyAdded As Boolean      ' Set true after all theoretical tryptic MT tags have been computed and loaded into GelOrfMassTags()
    DataParsedCompletely As Boolean                             ' Set true after MT tags have been loaded from the database
    OtherInfo As String                     ' Reserved for future expansion
End Type

Public Type udtORFPeptideLocationType
    ResidueStart As Long                ' Residue number in an ORF of the first residue in a peptide
    ResidueEnd As Long                  ' Residue number in an ORF of the last residue in a peptide
    TrypticFragmentName As String       ' See following comments
End Type
        ' TrypticFragmentName = "" if does not match a tryptic fragment, otherwise, it is of the
        '  form t1 for tryptic fragment 1, t2 for fragment 2, etc., up to TrypticFragmentCount
        ' It may also be of the form t1.2 if the peptide is composed of tryptic
        '  fragments 1 and 2, or t5.3 if the peptide is composed of tryptic fragments 5, 6, and 7, etc.
    
Public Type udtMassTagType
    MassTagRefID As Long
    IsTheoretical As Boolean            ' True if an unobserved, tryptic MT tag (Not that MassTagRefID will be -100000, or something like this
    IsAMT As Boolean
    IsLocker As Boolean
    Location As udtORFPeptideLocationType
    GANET As Double                             ' Average GANET value; if IsTheoretical = True, then the Predicted NET value
    Mass As Double                              ' Monoisotopic (uncharged) mass; includes the masses of any modifications
    IsModified As Boolean
    StaticModList As String
    DynamicModList As String
End Type

' No longer supported (March 2006)
''Public Type udtORFMassTagsDataType
''    RefID As Long                           ' ID of the ORF in the MassTag Database; should match GelORFData().Orfs().RefID
''    MassTagCount As Long
''    MassTags() As udtMassTagType            ' 0-based array
''    OtherInfo As String                     ' Reserved for future expansion
''End Type
''
''Public Type udtORFMassTagsListType
''    ORFCount As Long
''    Orfs() As udtORFMassTagsDataType
''    Definition As udtORFMassTagsDefinitionType
''End Type

'' Code that was used by the ORFViewer; No longer supported (March 2006)
''
''' The udtORFViewerGroupListType structure is used to hold pointers to items in GelORFData()
'''  that are to be available for display in the ORF Viewer
''' Each instance of frmORFViewer contains it's own variable of type udtORFViewerGroupListType
''' This allows multiple gels to be displayed in the ORF Viewer
''' Note that the gels need not have the the exact same Proteins, but it would make
'''  the best sense for comparison if they do
''Public Type udtORFViewerGroupItemType
''    GelIndex As Long            ' Pointer to be used to dereference x in GelORFData(x)
''    ORFIndex As Long            ' Pointer to be used to dereference x in GelORFData(GelIndex).Orfs(x)
''End Type
''
''Public Type udtORFViewerGroupDataType
''    Reference As String         ' String describing ORF, for example ORF00002
''    RefID As Long               ' ID of the ORF in the MassTag Database (field Ref_ID)
''
''    ItemCount As Long                       ' Pointers to items in GelORFData()
''    Items() As udtORFViewerGroupItemType
''End Type
''
''Public Type udtORFViewerGroupListType
''    ORFCount As Long
''    Orfs() As udtORFViewerGroupDataType         ' 0-based array
''End Type
''
''
''' The udtORFViewerGelListType structure is used to hold details on the Gels that are available
''' for display in a given ORFViewer; the list is updated each time an instance of frmORFViewer is activated
''' so that the user can choose which Gels to include in the ORF Viewer
''
''Public Type udtORFViewerOptionsType
''    PicturePixelHeight As Long
''    PicturePixelWidth As Long
''    PicturePixelSpacing As Long
''    MaxSpotSizePixels As Long
''    MinSpotSizePixels As Long
''    SwapPlottingAxes As Boolean
''
''    MassDisplayRangePPM As Double
''    NETDisplayRange As Double
''
''    MassTagMassErrorPPM As Double           ' Error/tolerance in .Mass value for the MT tag, in PPM
''    MassTagNETError As Double               ' Error/tolerance in .NET value for the MT tag
''    MassTagSpotColor As Long
''    MassTagSpotShape As Integer             ' Must use Type Integer since saving to disk; actually is sSpotsShape
''
''    LogarithmicIntensityPlotting As Boolean
''    IntensityScalar As Double                       ' Value to divide all intensities by when listing in the ListViews
''    IonToUMCPlottingIntensityRatio As Double        ' Value to multiply Ion intensities by when plotting both ions and LC-MS Features on the same graph
''
''    LoadPMTs As Boolean
''    IncludeUnobservedTrypticMassTags As Boolean
''    ShowNonTrypticMassTagsWithoutIonHits As Boolean
''    HideEmptyMassTagPictures As Boolean                 ' Hides a single MT tag in the ORF Viewer if no ions or LC-MS Features are within the view range
''
''    OnlyUseTop50PctForAveraging As Boolean          ' Only sum the intensities of the MT tags whose observed intensities are at least 50% of the maximum observed MT tag intensity (for the given ORF)
''
''    DataDisplayMode As Integer              ' Must use Type Integer since saving to disk; actually is ddmDataDisplayModeConstants
''    UseClassRepresentativeNET As Boolean
''
''    CleavageRuleID As Integer               ' See clsInSilicoDigest.InitializeCleavageRules() for a description of the rules
''
''    ShowPosition As Boolean
''    ShowTickMarkLabels As Boolean
''    ShowGridLines As Boolean
''End Type
''
''Public Type udtORFViewerGelDataType
''    Deleted As Boolean
''    FileName As String                  ' The name of the original PEK/CSV/mzXML/mzData file that was loaded
''    GelFileName As String               ' The name of the file, as saved to disk
''    IsoLines As Long
''    MTDBName As String
''    IncludeGel As Boolean               ' Whether or not to include the gel's data in the ORF viewer
''    NETAdjustmentType As Integer        ' Must use Type Integer since saving to disk; actually is natNETTypeConstants
''    IonSpotShape As Integer             ' Must use Type Integer since saving to disk; actually is sSpotsShape
''    IonSpotColor As Long
''    IonSpotColorSelected As Long
''    UMCSpotShape As Integer             ' Must use Type Integer since saving to disk; actually is sSpotsShape
''    UMCSpotColor As Long
''    UMCSpotColorSelected As Long
''    VisibleScopeOnly As Boolean         ' If True, then only uses thoses ions currently shown in the visible scope for the gel (i.e., don't show excluded ions)
''    ZOrder As Long                      ' First ZOrder has a value of 0
''End Type
''
''Public Type udtORFViewerGelListType
''    GelCount As Integer
''    Gels() As udtORFViewerGelDataType           ' 1-based array; .Gels(1) corresponds to data in GelData(1), handled by the frmORFViewer.UpdateGelDisplayList() function
''    DisplayOptions As udtORFViewerOptionsType
''End Type

''' The following is used to hold a copy of data of type udtORFViewerGelListType
''' It is used to restore the previously used options when displaying the ORF Viewer on a loaded data file that had been previously viewed with the ORF viewer
''Public Type udtORFViewerSavedGelListType
''    IsDefined As Boolean
''    SavedGelListAndOptions As udtORFViewerGelListType
''End Type

' Note: Since udtIsotopicDataType now contains MassShiftOverallPPM and MassShiftCount, these variables are only for reference
' In particular, the data in AdjustmentHistory() cannot be used to uncalibrate the data in a stepwise fashion; only the overall calibration applied is kept track of
Public Type udtMassCalibrationInfoType
    MassUnits As Integer                ' Actually type glMassToleranceConstants
    OverallMassAdjustment As Double     ' Units are given by MassUnits; once one adjustment has been applied, all subsequent adjustments must use the same units
    AdjustmentHistoryCount As Integer
    AdjustmentHistory() As Double       ' 0-based array, listing the incremental adjustments applied; the newest adjustment is in the highest index of the array
    OtherInfo As String
End Type

Public Type udtDBSearchMassModificationOptions2003dType
    DynamicMods As Boolean              ' When true, varying numbers of cysteine modifications are applied
    N15InsteadOfN14 As Boolean
    PEO As Boolean
    ICATd0 As Boolean
    ICATd8 As Boolean
    Alkylation As Boolean
    AlkylationMass As Double            ' Normally glAlkylation = 57.0215 Da, but can be customized
    OtherInfo As String
End Type

' Note: udtSearchDefinitionGroupType depends on this, so do not change without considering the consequences
Public Type udtDBSearchMassModificationOptionsType
    ' OldParameter:DynamicMods          ' 2 bytes; When true, varying numbers of residue modifications are applied; if false, then a static mod
    ModMode As Byte                     ' 1 byte; 0 = fixed (static mods); 1 = dynamic mods; 2 = decoy mods (like dynamic mods, but the NET value is varied for added AMTs)
    UnusedByte As Byte                  ' 1 byte; was previously part of DynamicMods, a boolean
    N15InsteadOfN14 As Boolean
    PEO As Boolean
    ICATd0 As Boolean
    ICATd8 As Boolean
    Alkylation As Boolean
    AlkylationMass As Double            ' Normally glAlkylation = 57.0215 Da, but can be customized
    ResidueToModify As String           ' Single letter amino acid symbol (or list of symbols) to add ResidueMassModification to; leave blank to add the mass to the entire MT tag
    ResidueMassModification As Double
    OtherInfo As String
End Type

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 2
Public Type udtSearchDefinition2002GroupType
    UMCDef As UMCDefinition2002
    AMTSearchOnIons As SearchAMTDefinition
    AMTSearchOnUMCs As SearchAMTDefinition
    AMTSearchOnPairs As SearchAMTDefinition
    AnalysisHistory() As String
    AnalysisHistoryCount As Long
End Type

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 3
Public Type udtSearchDefinition2003GroupType
    UMCDef As UMCDefinition2003a                     ' Updated for this version
    AMTSearchOnIons As SearchAMTDefinition
    AMTSearchOnUMCs As SearchAMTDefinition
    AMTSearchOnPairs As SearchAMTDefinition
    AnalysisHistory() As String
    AnalysisHistoryCount As Long
End Type

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 4
Public Type udtSearchDefinition2003bGroupType
    UMCDef As UMCDefinition2003a
    UMCIonNetDef As UMCIonNetDefinition            ' New for this version; used on frmUMCIonNet
    AMTSearchOnIons As SearchAMTDefinition
    AMTSearchOnUMCs As SearchAMTDefinition
    AMTSearchOnPairs As SearchAMTDefinition
    AnalysisHistory() As String
    AnalysisHistoryCount As Long
End Type

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 5
Public Type udtSearchDefinition2003cGroupType
    UMCDef As UMCDefinition2003a
    UMCIonNetDef As UMCIonNetDefinition
    AMTSearchOnIons As SearchAMTDefinition
    AMTSearchOnUMCs As SearchAMTDefinition
    AMTSearchOnPairs As SearchAMTDefinition
    AnalysisHistory() As String
    AnalysisHistoryCount As Long
    MassCalibrationInfo As udtMassCalibrationInfoType       ' New for this version
    OtherInfo As String                                     ' New for this version
End Type

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 6
Public Type udtSearchDefinition2003dGroupType
    UMCDef As UMCDefinition2003a
    UMCIonNetDef As UMCIonNetDefinition
    AMTSearchOnIons As SearchAMTDefinition
    AMTSearchOnUMCs As SearchAMTDefinition
    AMTSearchOnPairs As SearchAMTDefinition
    AnalysisHistory() As String                                         ' History (log) of the searches and steps performed for this gel; an entry is made whenever LC-MS Features are searched for, or modified, and whenever MT tags are searched against; entries also made for other actions
    AnalysisHistoryCount As Long
    MassCalibrationInfo As udtMassCalibrationInfoType
    AMTSearchMassMods As udtDBSearchMassModificationOptions2003dType     ' New for this version
    OtherInfo As String
End Type

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 7
Public Type udtSearchDefinition2003eGroupType
    UMCDef As UMCDefinition2003a
    UMCIonNetDef As UMCIonNetDefinition
    AMTSearchOnIons As SearchAMTDefinition
    AMTSearchOnUMCs As SearchAMTDefinition
    AMTSearchOnPairs As SearchAMTDefinition
    AnalysisHistory() As String                                     ' History (log) of the searches and steps performed for this gel; an entry is made whenever LC-MS Features are searched for, or modified, and whenever MT tags are searched against; entries also made for other actions
    AnalysisHistoryCount As Long
    MassCalibrationInfo As udtMassCalibrationInfoType
    AMTSearchMassMods As udtDBSearchMassModificationOptionsType     ' Updated for this version (added ResidueToModify and ResidueMassModification)
    OtherInfo As String
End Type

'This corresponds to FileInfoVersions(fioSearchDefinitions) version 8
Public Type udtSearchDefinitionGroupType
    UMCDef As UMCDefinition                                         ' Updated for this version
    UMCIonNetDef As UMCIonNetDefinition
    AMTSearchOnIons As SearchAMTDefinition
    AMTSearchOnUMCs As SearchAMTDefinition
    AMTSearchOnPairs As SearchAMTDefinition
    AnalysisHistory() As String                                     ' History (log) of the searches and steps performed for this gel; an entry is made whenever LC-MS Features are searched for, or modified, and whenever MT tags are searched against; entries also made for other actions
    AnalysisHistoryCount As Long
    MassCalibrationInfo As udtMassCalibrationInfoType
    AMTSearchMassMods As udtDBSearchMassModificationOptionsType
    OtherInfo As String
End Type

Public Type udtGraph2DOptionsType
    ShowPointSymbols As Boolean
    DrawLinesBetweenPoints As Boolean
    ShowGridLines As Boolean
    AutoScaleXAxis As Boolean           ' Not used on frmPairBrowser
    PointSizePixels As Long
    PointShape As Integer               ' Actually type OlectraChart2D.ShapeConstants, valid values are 1 to 11
    PointAndLineColor As Long
    LineWidthPixels As Long
    CenterYAxis As Boolean              ' Not used on frmPairBrowser or frmTICandBPIPlots
    ShowSmoothedData As Boolean
    ShowPeakEdges As Boolean
End Type

Public Type udtGraph3DOptionsType
    ContourLevelsCount As Long
    Perspective As Single
    Elevation As Single             ' XRotation
    YRotation As Single
    ZRotation As Single
    AnnotationFontSize As Long
End Type

Public Type udtSplitUMCsByAbundanceOptionsType
    MinimumDifferenceInAveragePpmMassToSplit As Double
    StdDevMultiplierForSplitting As Double                  ' If 0, then Standard Deviation is not considered when determining whether or not to split a UMC based on abundance differences
    MaximumPeakCountToSplitUMC As Long
    PeakDetectIntensityThresholdPercentageOfMaximum As Long
    PeakDetectIntensityThresholdAbsoluteMinimum As Double
    PeakWidthPointsMinimum As Long
    PeakWidthInSigma As Long
    ScanGapBehavior As susgSplitUMCsScanGapConstants
End Type

Public Type udtUMCAutoRefineOptionsType
    UMCAutoRefineRemoveCountLow As Boolean
    UMCAutoRefineRemoveCountHigh As Boolean
    UMCAutoRefineRemoveMaxLengthPctAllScans As Boolean
    
    UMCAutoRefineMinLength As Long
    UMCAutoRefineMaxLength As Long                      ' Maximum UMC size in number of members or in scan width, depending on TestLengthUsingScanRange
    UMCAutoRefineMaxLengthPctAllScans As Long           ' A percentage, between 1 and 100, that is multiplied by the total scan range (MaxScan - MinScan) to determine the maximum allowable UMC length; only used if .UMCAutoRefineRemoveMaxLengthPctAllScans = True
    UMCAutoRefinePercentMaxAbuToUseForLength As Long    ' The maximum abundance of the UMC is taken times this value; we then find the left-most and right-most points in the UMC that are greater than this value.  Those locations are used to judge the UMC width during auto refinement
    
    TestLengthUsingScanRange As Boolean
    MinMemberCountWhenUsingScanRange As Long
    
    UMCAutoRefineRemoveAbundanceLow As Boolean
    UMCAutoRefineRemoveAbundanceHigh As Boolean
    UMCAutoRefinePctLowAbundance As Double              ' Percent of low abundance peaks to remove
    UMCAutoRefinePctHighAbundance As Double             ' Percent of high abundance peaks to remove

    SplitUMCsByAbundance As Boolean
    SplitUMCOptions As udtSplitUMCsByAbundanceOptionsType
End Type

Public Type udtUMCIonNetOptionsType
    ConnectionLengthPostFilterMaxNET As Double
    UMCRepresentative As Integer                    ' UMCFROMNet_REP_ABU, UMCFROMNet_REP_FIT, UMCFROMNet_REP_FST_SCAN, UMCFROMNet_REP_LST_SCAN, or UMCFROMNet_REP_MED_SCAN
    MakeSingleMemberClasses As Boolean
End Type

Public Type udtUMCAdvancedStatsOptionsType
    ClassAbuTopXMinAbu As Double
    ClassAbuTopXMaxAbu As Double
    ClassAbuTopXMinMembers As Long              ' Maximum members to include if TopXMinAbu and TopXMaxAbu are < 0; otherwise, minimum members to include
    
    ClassMassTopXMinAbu As Double
    ClassMassTopXMaxAbu As Double
    ClassMassTopXMinMembers As Long             ' Maximum members to include if TopXMinAbu and TopXMaxAbu are < 0; otherwise, minimum members to include
End Type

Public Type udtNETAdjustmentUMCDistributionOptionsType
    RequireDispersedUMCSelection As Boolean         ' When True, makes sure the LC-MS Features used for NET adjustment
    SegmentCount As Long                            ' Number of segments to divide the Gel into for NET adjustment
    MinimumUMCsPerSegmentPctTopAbuPct As Integer    ' Percentage, ranging from 0 to 100; Defines the minimum number of LC-MS Features that must be present in each segment.  The Top Abu Pct value is taken times this percentage to get the minimum LC-MS Features to use
    ScanPctStart As Integer                         ' Percentage, ranging from 0 to 100; Total Scan Count is taken times this number to obtain the scan number of the start of the first segment
    ScanPctEnd As Integer                           ' Percentage, ranging from 0 to 100; Total Scan Count is taken times this number to obtain the scan number of the end of the last segment
End Type

Public Type udtErrorDistributionOptionsType
    MassRangePPM As Long
    MassBinSizePPM As Single                    ' Size of each bin; this should be a nice round number like 10, 5, 1, 0.5, 0.1, etc.
    GANETRange As Single
    GANETBinSize As Single                      ' Size of each bin; this should be a nice round number like 1, 0.1, 0.01, 0.005, 0.001, etc.
    
    ButterWorthFrequency As Single              ' Butterworth sampling frequency (affects smoothing); defaults to 0.15
    
    Graph2DOptions As udtGraph2DOptionsType
    Graph3DOptions As udtGraph3DOptionsType
End Type

Public Type udtTICAndBPIOptionsType
    PlotNETOnXAxis As Boolean                       ' When false, the X-Axis is Scan number
    NormalizeYAxis As Boolean                       ' If true, normalize to a maximum of 100
    SmoothUsingMovingAverage As Boolean
    MovingAverageWindowWidth As Long                ' Width in points
    TimeDomainDataMaxValue As Double                ' If non-zero, then defines the maximum allowable vlaue for the TIC based on Time Domain data
    Graph2DOptions As udtGraph2DOptionsType
    
    PointShapeSeries2 As Integer
    PointAndLineColorSeries2 As Long
    
    ClipOutliers As Boolean
    ClipOutliersFactor As Single
    
    KeepWindowOnTop As Boolean
End Type

Public Type udtFeatureBrowserOptionsType
    SortOrder As Integer                        ' Actually type epsPairSortOrderConstants or eusUMCSortOrderConstants
    SortDescending As Boolean
    AutoZoom2DPlot As Boolean
    HighlightMembers As Boolean
    PlotAllChargeStates As Boolean
    
    FixedDimensionsForAutoZoom As Boolean
    MassRangeZoom As Double
    MassRangeUnits As Integer                       ' Actually type mruMassRangeUnitsConstants
    ScanRangeZoom As Double
    ScanRangeUnits As Integer                       ' Actually type sruScanRangeUnitsConstants
    
    Graph2DOptions As udtGraph2DOptionsType         ' The point info in Graph2DOptions applies to the Light Member; following variables are for Heavy Member
    PointShapeHeavy As Integer                      ' Only used with the Pair Browser; Actually type OlectraChart2D.ShapeConstants, valid values are 1 to 11
    PointAndLineColorHeavy As Long                  ' Only used with the Pair Browser
    
    KeepWindowOnTop As Boolean
End Type

Public Type udtPairSearchOptionsType
    SearchDef As udtIsoPairsSearchDefType
    
    PairSearchMode As String                    ' DeltaPairs or LabelPairs
    
    AutoExcludeOutOfERRange As Boolean
    AutoExcludeAmbiguous As Boolean
    KeepMostConfidentAmbiguous As Boolean
    
    AutoAnalysisRemovePairMemberHitsAfterDBSearch As Boolean
    AutoAnalysisRemovePairMemberHitsRemoveHeavy As Boolean      ' If True, remove Heavy members of pairs that match; if False, remove Light members of pairs that match
    
    AutoAnalysisSavePairsToTextFile As Boolean
    AutoAnalysisSavePairsStatisticsToTextFile As Boolean
    
    NETAdjustmentPairedSearchUMCSelection As punaPairsUMCNetAdjustmentConstants
    
    '' This variable is no longer used and is effectively assumed to always be true
    ''OutlierRemovalUsesSymmetricERs As Boolean                   ' Affects both outlier removal and the weighted average value computed
    
    AutoAnalysisDeltaMassAddnlCount As Integer
    AutoAnalysisDeltaMassAddnl() As Double

End Type

Public Type udtPairMatchStatsType
    PairIndex As Long
    ExpressionRatio As Double
    ExpressionRatioStDev As Double
    ExpressionRatioChargeStateBasisCount As Integer
    ExpressionRatioMemberBasisCount As Long
    LabellingEfficiencyF As Single
    LogERCorrectedForF As Single
    LogERStandardError As Single
End Type

Public Type udtAMTStalenessOptionsType
    MaximumAgeLoadedMassTagsHours As Long       ' hours
    MaximumFractionAMTsWithNulls As Single      ' Value between 0 and 1; if value is 0.01, then allow at most 1% of the MT tags to have null masses or NET values.
                                                ' If (AMTCountWithNulls / AMTCountInDB) is >= MaximumFractionAMTsWithNulls, then try to re-load the MT tags if (Now() - AMTLoadTime) is > MinimumTimeBetweenReloadMinutes minutes
    MaximumCountAMTsWithNulls As Long           ' Alternate minimum to MaximumFractionAMTsWithNulls; useful for large databases where MaximumFractionAMTsWithNulls becomes less useful
    MinimumTimeBetweenReloadMinutes As Long     ' minutes
    
    ' The following two variables are used to keep track of when the MT tags were last
    '  loaded from the database.  Each time LoadMassTags is called, it updates MassTagStalenessOptions.AMTLoadTime
    ' When ConfirmMassTagsAndInternalStdsLoaded is called, it checks the current time against MassTagStalenessOptions.AMTLoadTime
    '  If more than MassTagStalenessOptions.MaximumAgeLoadedMassTagsHours has elapsed, then MT tags are reloaded
    AMTLoadTime As Date
    AMTCountInDB As Long                        ' Count of all MT tags in database
    AMTCountWithNulls As Long                   ' Count of number of MT tags that have a valid mass and NET value (i.e., neither is null)
    
    AMTStatsLoadTime As Date
End Type

' The following defines how the SLiC scores (aka match scores) are computed
Public Type udtMatchScoreOptionsType
    MassPPMStDev As Double                          ' Default 3
    NETStDev As Double                              ' Default 0.025
    UseAMTNETStDev As Boolean                       ' December 2005: This value is now ignored, and essentially defaults to False
    MaxSearchDistanceMultiplier As Integer          ' Default 2
    AutoDefineSLiCScoreThresholds As Boolean        ' If True, then MassPPMStDev and NETStDev are set equal to the mass and NET tolerance used during the search, divided by STDEV_SCALING_FACTOR = 2
End Type

Public Type udtGraphicExportOptionsType
    CopyEMFIncludeFilenameAndDate As Boolean
    CopyEMFIncludeTextLabels As Boolean
End Type

Public Type udtUMCMassTagMatchStats
    UMCIndex As Long                ' Index of the UMC in GelUMC().UMCs()  (Form frmSearchMT_ConglomerateUMC)
    PairIndex As Long               ' Index of the Pair                     (Form frmSearchMTPairs)
    IDIndex As Long                 ' Index of the AMT match in AMTData() or index in UMCInternalStandards.InternalStandards(); for some forms (namely frmSearchMT_ConglomerateUMC and frmSearchMTPairs) this is actually a pointer to an array that contains the actual index (mMT arrays); lastly, when copying LC-MS Features in view, this is the actual Mass_Tag_ID value
    IDIsInternalStd As Boolean      ' True if the ID is an Internal Std, False if a MT tag
    AMTMods As String               ' Mods, if any (Like PEO, ICAT, etc.); only applies to AMT's
    MemberHitCount As Long          ' The number of members of a given UMC that matched the given MT tag or Internal Standard
    SLiCScore As Double             ' SLiC Score (Spatially Localized Confidence score)
    DelSLiC As Double               ' Similar to DelCN, difference in SLiC score between top match and match with score value one less than this score
    MassDiffPPM As Double           ' Mass difference between AMT and given UMC or given point
    MultiAMTHitCount As Long        ' The number of Unique MT tag hits for each UMC; only applies to AMT's (in other words, ignores Internal Standard)
End Type

Public Type udtExclusionIonType
    IonMass As Double
    TolerancePPM As Double
    
    Charge As Integer           ' 0 means match all charges
    
    LimitScanRange As Boolean
    ScanStart As Long
    ScanEnd As Long
End Type

Public Type udtExclusionPolygonType
    VertexCount As Integer
    VertexList() As String      ' X,Y pairs defining the vertices of the polygon; X is scan number and Y is monoisotopic mass (both must be integers)
End Type

Public Type udtErrorPlottingPeakCacheType
    Center As Double
    width As Double
    Height As Integer
    SignalToNoise As Double
    SingleValidPeak As Boolean
    PeakStats As udtPeakStatsType
'    IndexCenter As Long
'    IndexBaseLeft As Long
'    IndexBaseRight As Long
'    TruePositiveArea As Long                ' Area in the peak that is above the background level
'    FalsePositiveArea As Long               ' Area in the peak that is below the background level
End Type

Public Type udtNoiseRemovalOptionsType
    SearchTolerancePPMDefault As Double
    SearchTolerancePPMAutoRemoval As Double
    
    PercentageThresholdToExcludeSlice As Single                 ' a number between 0.0 and 100.0
    PercentageThresholdToAddNeighborToSearchSlice As Single     ' a number between 0.0 and 100.0
    
    LimitMassRange As Boolean
    MassStart As Double
    MassEnd As Double
    
    LimitScanRange As Boolean
    ScanStart As Long
    ScanEnd As Long
    
    SearchScope As glScope                  ' 0 to search all, 1 to search current scope only
    RequireIdenticalCharge As Boolean

    ExclusionListCount As Long
    ExclusionList() As udtExclusionIonType
    
    ExclusionPolygonCount As Integer
    ExclusionPolygonList() As udtExclusionPolygonType
End Type

Public Type udtRefineMSDataOptionsType
    MinimumPeakHeight As Long                               ' counts/bin
    MinimumSignalToNoiseRatioForLowAbundancePeaks As Single ' signal to noise ratio; only applies to peaks with intensity <= .MinimumPeakHeight
    PercentageOfMaxForFindingWidth As Long                  ' The maximum is found, then we iterate left and right until this point is reached
    
    MassCalibrationMaximumShift As Double                   ' ppm or Da, depending on MassCalibrationTolType
    MassCalibrationTolType As glMassToleranceConstants
    
    ToleranceRefinementMethod As mtrMassToleranceRefinementConstants            ' This applies to both Mass and NET tolerance refinement
    UseMinMaxIfOutOfRange As Boolean                        ' If True, then uses MassToleranceMinimum or MassToleranceMaximum if the new tolerance defined is out-of-range
    
    MassToleranceMinimum As Double              ' ppm
    MassToleranceMaximum As Double              ' ppm
    MassToleranceAdjustmentMultiplier As Double
    
    NETToleranceMinimum As Double               ' NET
    NETToleranceMaximum As Double               ' NET
    NETToleranceAdjustmentMultiplier As Double
    
    IncludeInternalStdMatches As Boolean
    
    UseUMCClassStats As Boolean
    MinimumSLiC As Single
    MaximumAbundance As Double                          ' When UseUMCClassStats = True, then the UMC class abundances are tested against MaximumAbundance; when UseUMCClassStats = False, we're using individual data points and thus the individual data point abundances are tested against MaximumAbundance
    
    EMMassErrorPeakToleranceEstimatePPM As Single       ' Used to compute the initial variance value to send to the Expectation Maximization algorithm; this is estimated peak half width at 5 sigma
    EMNETErrorPeakToleranceEstimate As Single           ' Used to compute the initial variance value to send to the Expectation Maximization algorithm; this is estimated peak half width at 5 sigma
    EMIterationCount As Integer                         ' Number of iterations to perform for Expectation Maximization; must be >= 5
    EMPercentOfDataToExclude As Integer                 ' Value between 0 and 75; used when determining minimum and maximum values in the range of data to use for Expectation Maximization; excludes EMPercentOfDataToExclude of the data (by removing DataCount*EMPercentOfDataToExclude/100/2 values from the extremes of the array)
    
    EMMassTolRefineForceUseSingleDataPointErrors As Boolean    ' When True, then sends individual data point errors to the Expectation Maximization algorithm
    EMNETTolRefineForceUseSingleDataPointErrors As Boolean     ' When True, then sends individual data point errors to the Expectation Maximization algorithm

    ComputePairwiseMassDifferences As Boolean
    PairwiseMassDiffMinimum As Single
    PairwiseMassDiffMaximum As Single
    PairwiseMassBinSize As Single
    PairwiseMassDiffNETTolerance As Single
    PairwiseMassDiffNETOffset As Single
End Type

Public Type udtAutoToleranceRefinementType
    DBSearchMWTol As Double
    DBSearchTolType As glMassToleranceConstants
    DBSearchNETTol As Double
    DBSearchRegionShape As srsSearchRegionShapeConstants
    
    DBSearchMinimumHighNormalizedScore As Single        ' Minimum MT tag high normalized score (usually XCorr) to use when searching DB for tolerance refinement
    DBSearchMinimumHighDiscriminantScore As Single      ' Minimum MT tag high discriminant score to use when searching DB for tolerance refinement
    DBSearchMinimumPeptideProphetProbability As Single      ' Minimum MT tag peptide prophet probability to use when searching DB for tolerance refinement

    RefineMassCalibration As Boolean
    RefineMassCalibrationOverridePPM As Double          ' If this value is non-zero, and RefineMassCalibration = True, then the data will be shifted by this amount, regardless of where the peak is in the mass error plot
    RefineDBSearchMassTolerance As Boolean
    RefineDBSearchNETTolerance As Boolean
End Type

Public Type udtAutoAnalysisSearchModeOptionsType
    SearchMode As String                    ' See AUTO_SEARCH_EXPORT_UMCS_ONLY, AUTO_SEARCH_ORGANISM_MTDB, AUTO_SEARCH_UMC_MTDB, etc. for allowable strings
    AlternateOutputFolderPath As String
    WriteResultsToTextFile As Boolean
    ExportResultsToDatabase As Boolean
    ExportUMCMembers As Boolean
    PairSearchAssumeMassTagsAreLabeled As Boolean                       ' Currently not used (June 2004)
    InternalStdSearchMode As issmInternalStandardSearchModeConstants    ' Note: if APP_BUILD_DISABLE_MTS = True, then this is set to issmFindOnlyMassTags when searching the LC-MS Features
    DBSearchMinimumHighNormalizedScore As Single
    DBSearchMinimumHighDiscriminantScore As Single
    DBSearchMinimumPeptideProphetProbability As Single
    MassMods As udtDBSearchMassModificationOptionsType
End Type

Public Type udtAutoAnalysisOptionsType
    DatasetID As Long                                       ' This can be provided on the command line (in the .Par file), or in the .Ini file
    JobNumber As Long                                       ' This can be provided on the command line (in the .Par file), or in the .Ini file
    MDType As Long                                          ' Defined in T_MMD_Type_Name
    
    AutoRemoveNoiseStreaks As Boolean
    AutoRemovePolygonRegions As Boolean
    
    DoNotSaveOrExport As Boolean
    SkipFindUMCs As Boolean                        ' Only appropriate if loading data from a .Gel file or an _LCMSFeatures.txt file (however, this is not forced to false if a .Pek, .CSV, .mzXML, or .mzData file is loaded)
    SkipGANETSlopeAndInterceptComputation As Boolean
    
    DBConnectionRetryAttemptMax As Integer
    DBConnectionTimeoutSeconds As Integer
    ExportResultsFileUsesJobNumberInsteadOfDataSetName As Boolean
    
    GenerateMonoPlus4IsoLabelingFile As Boolean         ' If true, then calls IsoLabelingIDMain.exe to create the _pairs_isos.csv file using the .Dat file and the _isos.csv file
    
    SaveGelFile As Boolean
    SaveGelFileOnError As Boolean
    SavePictureGraphic As Boolean
    SavePictureGraphicFileType As pftPictureFileTypeConstants       ' 1=PNG, 2=JPG, 3=WMF, 4=EMF, 5=BMP
    SavePictureWidthPixels As Long
    SavePictureHeightPixels As Long
    
    SaveInternalStdHitsAndData As Boolean
    
    SaveErrorGraphicMass As Boolean
    SaveErrorGraphicGANET As Boolean
    SaveErrorGraphic3D As Boolean
    SaveErrorGraphicFileType As pftPictureFileTypeConstants         ' 1=PNG, 2=JPG; this also defines the format for saving TIC and BPI Plots
    SaveErrorGraphSizeWidthPixels As Long       ' This also defines the format for saving TIC and BPI Plots
    SaveErrorGraphSizeHeightPixels As Long      ' This also defines the format for saving TIC and BPI Plots

    SavePlotTIC As Boolean
    SavePlotBPI As Boolean
    SavePlotTICTimeDomain As Boolean
    SavePlotTICDataPointCounts As Boolean
    SavePlotTICDataPointCountsHitsOnly As Boolean

    SavePlotTICFromRawData As Boolean
    SavePlotBPIFromRawData As Boolean
    SavePlotDeisotopingIntensityThresholds As Boolean
    SavePlotDeisotopingPeakCounts As Boolean
    
    NETAdjustmentInitialNetTol As Double
    ''NETAdjustmentFinalNetTol As Double                ' November 2005: Unused variable; was previously used for Simulated Annealing
    NETAdjustmentMaxIterationCount As Long
    NETAdjustmentMinIDCount As Long                     ' The desired MinIDCount
    NETAdjustmentMinIDCountAbsoluteMinimum As Long      ' The optional MinIDCount to use if we're unable to get enough hits using TopAbuPct = NETAdjustmentUMCTopAbuPctMax
    NETAdjustmentMinIterationCount As Long              ' If the number of iterations is below this value during auto-analysis, a warning is entered into the analysis history and the Net Adjustment warning bit is enabled;
                                                        ' This can currently only be set in the .Ini file
    NETAdjustmentChangeThresholdStopValue As Double
    
    NETAdjustmentAutoIncrementUMCTopAbuPct As Boolean
    NETAdjustmentUMCTopAbuPctInitial As Long
    NETAdjustmentUMCTopAbuPctIncrement As Long
    NETAdjustmentUMCTopAbuPctMax As Long                ' This can currently only be set in the .Ini file
    
    ''NETAdjustmentMinimumNETMatchScore As Long         ' November 2005: Unused variable; was previously used for Simulated Annealing; minimum NET match score to obtain during Robust NET searching
    
    NETSlopeExpectedMinimum As Double                   ' Ignored if GelUMCNETAdjDef().UseRobustNETAdjustment = True
    NETSlopeExpectedMaximum As Double                   ' Ignored if GelUMCNETAdjDef().UseRobustNETAdjustment = True
    NETInterceptExpectedMinimum As Double               ' Ignored if GelUMCNETAdjDef().UseRobustNETAdjustment = True
    NETInterceptExpectedMaximum As Double               ' Ignored if GelUMCNETAdjDef().UseRobustNETAdjustment = True
    
    UMCSearchMode As String                             ' AUTO_ANALYSIS_UMC2003 or AUTO_ANALYSIS_UMCIonNet
    UMCShrinkingBoxWeightAverageMassByIntensity As Boolean      ' Only used in AUTO_ANALYSIS_UMCListType2002 (which is now obsolete - July 2004)
    UMCIonNetUsesInternalClusteringCode As Boolean      ' When true, then chkUseLCMSFeatureFinder is unchecked on frmUMCIonNet; however, this only honored during automated analysis
    
    OutputFileSeparationCharacter As String             ' Either a single character, or the word <TAB> to represent a Tab
    PEKFileExtensionPreferenceOrder As String
    WriteIDResultsByIonToTextFileAfterAutoSearches As Boolean            ' If True, then calls CopyAllPointsInView at the end of all of the auto-searches
    SaveUMCStatisticsToTextFile As Boolean
    IncludeORFNameInTextFileOutput As Boolean
    SetIsConfirmedForDBSearchMatches As Boolean
    AddQuantitationDescriptionEntry As Boolean
    ExportUMCsWithNoMatches As Boolean
    
    DBSearchRegionShape As srsSearchRegionShapeConstants
    
    UseLegacyDBForMTs As Boolean
    IgnoreNETAdjustmentFailure As Boolean               ' If True, then ignores NET Adjustment Failures that occur prior to tolerance refinement
    
    AutoToleranceRefinement As udtAutoToleranceRefinementType
    
    AutoAnalysisSearchModeCount As Integer
    AutoAnalysisSearchMode(MAX_AUTO_SEARCH_MODE_COUNT) As udtAutoAnalysisSearchModeOptionsType           ' 0-based array
End Type

' Most of these filters are dynamic filters and can be changed with data in memory
' There are a few exceptions, as noted in the comments
Public Type udtAutoAnalysisFilterPrefsType
    ExcludeDuplicates As Boolean
    ExcludeDuplicatesTolerance As Double
    
    ExcludeIsoByFit As Boolean                      ' If true, then applied when the data is loaded from the .PEK/.CSV file
    ExcludeIsoByFitMaxVal As Double                 ' Applied when the data is loaded from the PEK/CSV file
    
    ExcludeIsoSecondGuess As Boolean
    ExcludeIsoLessLikelyGuess As Boolean
    
    ExcludeCSByStdDev As Boolean
    ExcludeCSByStdDevMaxVal As Double               ' Applied when the data is loaded from the PEK/CSV file
    
    RestrictIsoByAbundance As Boolean
    RestrictIsoAbundanceMin As Double
    RestrictIsoAbundanceMax As Double
    
    RestrictIsoByMass As Boolean
    RestrictIsoMassMin As Double
    RestrictIsoMassMax As Double
    
    RestrictIsoByMZ As Boolean
    RestrictIsoMZMin As Double
    RestrictIsoMZMax As Double
    
    RestrictIsoByChargeState As Boolean
    RestrictIsoChargeStateMin As Integer
    RestrictIsoChargeStateMax As Integer
    
    RestrictCSByAbundance As Boolean
    RestrictCSAbundanceMin As Double
    RestrictCSAbundanceMax As Double
    
    RestrictCSByMass As Boolean
    RestrictCSMassMin As Double
    RestrictCSMassMax As Double
    
    RestrictScanRange As Boolean
    RestrictScanRangeMin As Long
    RestrictScanRangeMax As Long
    
    RestrictGANETRange As Boolean
    RestrictGANETRangeMin As Double
    RestrictGANETRangeMax As Double
    
    RestrictToEvenScanNumbersOnly As Boolean           ' Only one of these options can be set to True at any given time; setting both to false means no restriction
    RestrictToOddScanNumbersOnly As Boolean
    
    MaximumDataCountEnabled As Boolean                 ' This filter is only applied at the time the data is loaded into memory
    MaximumDataCountToLoad As Long

    TotalIntensityPercentageFilterEnabled As Boolean   ' This filter is only applied at the time the data is loaded into memory
    TotalIntensityPercentageFilter As Single

    AutoMapDataPointsMassTolerancePPM As Single        ' This setting is only used when we load _LCMSFeatures.txt files and we need to auto-map the data points to the features
    LCMSFeaturePointsLoadMode As plmPointsLoadModeConstants
End Type

Public Type udtAutoAnalysisCachedDataType
    Initialized As Boolean
    MassCalErrorPeakCached As udtErrorPlottingPeakCacheType
    NETTolErrorPeakCached As udtErrorPlottingPeakCacheType
End Type

Public Type udtAutoAnalyzeStateType
    Enabled As Boolean
    AutoAnalysisTimeStamp As String
    AutoRefiningUMCs As Boolean
    UseCurrentScopeOnly As Boolean
End Type

Public Type AutoQueryPRISMOptionsType
    ConnectionStringQueryDB As String
    RequestTaskSPName As String                 ' Name of the SP to call to find the next available task
    SetTaskCompleteSPName As String             ' Name of the SP to call to indicate that a task is complete (or failed)
    SetTaskToRestartSPName As String            ' Name of the SP to call to reset a task if Viper had a fatal error
    PostLogEntrySPName As String                ' Name of the SP to call to post an entry to the DB log
    
    QueryIntervalSeconds As Long
    MinimumPriorityToProcess As Integer         ' should range from 0 to MaximumPriorityToProcess
    MaximumPriorityToProcess As Integer         ' 0 to process all priorities, normally 1 to 5, though can range from 1 to 255
    PreferredDatabaseToProcess As String        ' Blank to process any database or a valid database name to prefer; also, if ExclusivelyUseThisDatabase = True, then the database to use is defined here
    ServerForPreferredDatabase As String        ' Specify the server on which this DB resides
    ExclusivelyUseThisDatabase As Boolean       ' When True, will only process peak matching tasks for PreferredDatabaseToProcess and ServerForPreferredDatabase
End Type

Public Type udtDMSConnectionInfoType
    ConnectionString As String
End Type

Public Type udtMTSConnectionInfoType
    ConnectionString As String
    
    ' Stored procedures
    spAddQuantitationDescription As String
    spGetLockers As String
    spGetMassTagMatchCount As String
    spGetMassTags As String
    spGetMassTagsSubset As String
    spGetPMResultStats As String
    spPutAnalysis As String
    ' spPutPeak As String                   ' September 2004: Unused variable
    
    spPutUMC As String
    spPutUMCMember As String
    spPutUMCMatch As String
    spPutUMCInternalStdMatch As String
    spPutUMCCSStats  As String
    
    spEditGANET As String
    spGetORFs As String
    spGetORFSeq As String
    spGetORFIDs As String
    spGetORFRecord As String
    spGetMassTagSeq As String
    spGetMassTagNames As String
    spGetInternalStandards As String
    spGetDBSchemaVersion As String
    spGetMassTagToProteinNameMap As String
    spGetMTStats As String
    
    ' Sql statements
    sqlGetMTNames As String
    ' Obsolete: sqlGetORFMassTagMap As String
    
End Type

Public Type udtPreferencesExpandedType
    MenuModeDefault As mmMenuModeConstants
    MenuModeIncludeObsolete As Boolean
    ExtendedFileSaveModePreferred As Boolean
    
    CopyPointsInViewIncludeSearchResultsChecked As Boolean
    CopyPointsInViewByUMCIncludeSearchResultsChecked As Boolean
    
    AutoAdjSize As Boolean              'if True spots on the gel are auto-sized to
                                        'make best display
    AutoSizeMultiplier As Single        ' Affects the size of the spots when auto-sizing
    
    UMCDrawType As Long
    
    UsePEKBasedERValues As Boolean          ' Can only be set in the .Ini file; when True, then stores PEK-based ER data in IsoData(i).ExpressionRatio when reading the PEK file
    UseMassTagsWithNullMass As Boolean      ' Can only be set in the .Ini file
    UseMassTagsWithNullNET As Boolean
    
    IReportAutoAddMonoPlus4AndMinus4Data As Boolean

    UseUMCConglomerateNET As Boolean        ' When true, then uses the NET of the LC-MS Features class rep, rather than using the NET of each member of the LC-MS Feature
    NetAdjustmentUsesN15AMTMasses As Boolean
    NetAdjustmentMinHighNormalizedScore As Single
    NetAdjustmentMinHighDiscriminantScore As Single
    
    AMTSearchResultsBehavior As asrbAMTSearchResultsBehaviorConstants           ' Whether or not to auto-remove existing DB search results prior to performing a new seach
    
    ICR2LSSpectrumViewZoomWindowWidthMZ As Double
    
    LastAutoAnalysisIniFilePath As String
    LastInputFileMode As ifmInputFileModeConstants
    LegacyAMTDBPath As String                   ' Replaces sAMTPath
    
    UMCAutoRefineOptions As udtUMCAutoRefineOptionsType
    UMCIonNetOptions As udtUMCIonNetOptionsType
    
    UMCAdvancedStatsOptions As udtUMCAdvancedStatsOptionsType
    
    NetAdjustmentUMCDistributionOptions As udtNETAdjustmentUMCDistributionOptionsType
    
    ErrorPlottingOptions As udtErrorDistributionOptionsType
    NoiseRemovalOptions As udtNoiseRemovalOptionsType
    RefineMSDataOptions As udtRefineMSDataOptionsType
    
    TICAndBPIPlottingOptions As udtTICAndBPIOptionsType
    
    PairBrowserPlottingOptions As udtFeatureBrowserOptionsType
    UMCBrowserPlottingOptions As udtFeatureBrowserOptionsType
    
    PairSearchOptions As udtPairSearchOptionsType
    
    MassTagStalenessOptions As udtAMTStalenessOptionsType
    
    SLiCScoreOptions As udtMatchScoreOptionsType               ' MT tag SLiC Score options
    
    GraphicExportOptions As udtGraphicExportOptionsType
    
    AutoAnalysisOptions As udtAutoAnalysisOptionsType
    AutoAnalysisFilterPrefs As udtAutoAnalysisFilterPrefsType
    AutoAnalysisCachedData As udtAutoAnalysisCachedDataType
    
    AutoAnalysisDBInfoIsValid As Boolean
    AutoAnalysisDBInfo As udtGelAnalysisInfoType
    
    AutoAnalysisStatus As udtAutoAnalyzeStateType
    
    AutoQueryPRISMOptions As AutoQueryPRISMOptionsType
    
    DMSConnectionInfo As udtDMSConnectionInfoType
    MTSConnectionInfo As udtMTSConnectionInfoType
  
End Type

Public Type udtAutoAnalysisMTDBOverrideType
    Enabled As Boolean
    ServerName As String
    MTDBName As String
    ConnectionString As String
    AMTsOnly As Boolean
    ConfirmedOnly As Boolean
    LockersOnly As Boolean
    LimitToPMTsFromDataset As Boolean
    MinimumHighNormalizedScore As Single
    MinimumHighDiscriminantScore As Single
    MinimumPeptideProphetProbability As Single
    MinimumPMTQualityScore As Single
    ExperimentInclusionFilter As String
    ExperimentExclusionFilter As String
    InternalStandardExplicit As String
    NETValueType As nvtNetValueTypeConstants
    MTSubsetID As Long
    ModList As String                                    ' Note: For all MT tags, leave blank, set to 'Any', or set to '-1'; For No Mods, use 'Not Any' in DB Schema Version 2 and Use 'Dynamic 1 and Static 1' in Schema Version 1
    DBSchemaVersion As Single
    PeakMatchingTaskID As Long
End Type

' ModList examples for DB Schema Version 2
' Items in list can be of the form:  [Not] GlobModID/Any
' For example: 1014            will filter for MT tags containing Mod 1014
'          or: 1014, 1010      will filter for MT tags containing Mod 1014 or Mod 1010
'          or: Any             will filter for any and all MT tags, regardless of mods
'          or: Not 1014        will filter for MT tags not containing Mod 1014 (including unmodified MT tags)
'          or: Not Any         will filter for MT tags without modifications
' Note that GlobModID = 1 means no modification, and thus:
'              1               will filter for MT tags without modifications (just like Not Any)
'              Not 1           will filter for MT tags with modifications
' Mods are defined in T_Mass_Correction_Factors in DMS and are accessible via MT_Main.V_DMS_Mass_Correction_Factors

Public Type udtAutoAnalysisFilePathsType
'''    VolClient As String
'''    StoragePath As String
    DatasetFolder As String     ' Folder name only
    ResultsFolder As String     ' Folder name only

    InputFilePath As String
    InputFilePathOriginal As String     ' This holds the original input file path specified; useful if the IsoLabelingIDMain.exe software was used to create a _pairs_isos.csv file from an _isos.csv file
    
    OutputFolderPath As String
    IniFilePath As String

    LogFilePath As String
    LogFilePathError As Boolean     ' If True, then an error occurred when initializing the log file
End Type


Public Type udtAutoAnalysisParametersType
    FilePaths As udtAutoAnalysisFilePathsType
    ShowMessages As Boolean
    DatasetID As Long                   ' Use -1 for none
    JobNumber As Long                   ' Use -1 for none
    MDID As Long                        ' Use -1 for none
    AutoCloseFileWhenDone As Boolean
    GelIndexToForce As Long
    MTDBOverride As udtAutoAnalysisMTDBOverrideType
    FullyAutomatedPRISMMode As Boolean                  ' When true, then initiated analysis during fully automated PRISM mode
    AutoDMSAnalysisManuallyInitiated As Boolean         ' When true, then means the user choose File->New DMS Analysis then chose to Auto-Analyze
    InvalidExportPassword As Boolean                    ' When true, then user was prompted for the Export to DB password, but entered an invalid password; thus, do not export to the DB
    ErrorBits As Long               ' Error bits are stored here
    WarningBits As Long             ' Warning bits are stored here
    ComputerName As String          ' The name of the computer running Viper
    
    ExitViperASAP As Boolean        ' If set to True, then Viper will exit auto analysis
    ExitViperReason As String       ' Reason for exiting Viper
    RestartAfterExit As Boolean     ' If True, then Viper will create a file named RestartViper.txt to indicate that Viper should be restarted
End Type

' The following is used to keep track of the UMCNetAdj definitions for each gel (used for Net Adjustment)
' It is saved to disk
Public GelUMCNETAdjDef() As NetAdjDefinition

' The following is used to keep track of the UMC searching definition, the AMT search definition, etc.
' It is saved to disk
Public GelSearchDef() As udtSearchDefinitionGroupType

' GelDataLookupArrays() allows one to quickly lookup certain values to
' determine the LC-MS Features that a given ion belongs to
Public GelDataLookupArrays() As udtGelDataLookupIndexType

' These arrays were used by the ORF Viewer; No longer supported (March 2006)
''' GelORFData contains the Proteins in memory for each gel, along with
'''  the loaded ions that match them
''' It is saved to disk
'''  The .Orfs() arrays in GelORFMassTags() are parallel arrays to those in GelORFData(),
'''  meaning GelORFData(1).Orfs(1) and GelORFMassTags(1).Orfs(1) refer to the same ORF
''Public GelORFData() As udtORFListType
''
''' GelORFMassTags contains all of the AMT's for each ORF
''' It is saved to disk
'''  The .Orfs() arrays in GelORFMassTags() are parallel arrays to those in GelORFData()
''Public GelORFMassTags() As udtORFMassTagsListType

''' This variable contains the options last used by this gel when viewed with the ORFViewer; No longer supported (March 2006)
''Public GelORFViewerSavedGelListAndOptions() As udtORFViewerSavedGelListType

''' Reserve memory for the ORFViewerLoader class immediately upon program start
''Public ORFViewerLoader As New ORFViewerLoaderClass

Public glbRecentFiles As udtRecentFilesType

Public glbPreferencesExpanded As udtPreferencesExpandedType

' This refers to the MwtWin.Dll ActiveX DLL
' It contains useful routines for handling peptide and protein sequences (determining mass, naming peptides, finding tryptic peptides, etc.)
''Public objMwtWin As MolecularWeightCalculator
''Public gMwtWinLoaded As Boolean
Public gTraceLogLevel As Integer

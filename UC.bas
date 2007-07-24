Attribute VB_Name = "Module12"
'functions dealing with unique mass classes
'-----------------------------------------------------------
'last modified: 04/16/2003 nt
'-----------------------------------------------------------
Option Explicit

''''used in ER and special AMT search
'''Public Const glUMC_LO = "[L]"
'''Public Const glUMC_HI = "[H]"

'abundance of the class
Public Enum UMCClassAbundanceConstants
    UMCAbuAvg = 0       'class average abundance
    UMCAbuSum = 1       'sum of class abundances
    UMCAbuRep = 2       'abundance of the class representative
    UMCAbuMed = 3       'median of class abundances
    UMCAbuMax = 4       'maximum of class abundances
    UMCAbuSumTopX = 5   ' sum of top X members

'''Public Const UMCClassAbundanceConstants.UMCAbuSum_SCAN_MAX = 5     'sum maximum abundance in each scan
'''Public Const UMCClassAbundanceConstants.UMCAbuSum_SCAN_AVG = 6     'sum average abundance in each scan

End Enum

'mass of the class
Public Enum UMCClassMassConstants
    UMCMassAvg = 0      'average class mass
    UMCMassRep = 1      'mass of class representative
    UMCMassMed = 2      'median
    UMCMassAvgTopX = 3      'average of top X members
    UMCMassMedTopX = 4      'median of top X members
End Enum

' Charge State Group Rep Type
Public Enum UMCChargeStateGroupConstants
    UMCCSGHighestSum = 0            ' Group with highest abundance sum
    UMCCSGMostAbuMember = 1         ' Most abundant member
    UMCCSGMostMembers = 2           ' Most members (highest member count)
End Enum

' The following Enum data is no longer used
Public Enum UMCNetConstants
    UMCNetBefore = 0
    UMCNetAt = 1
    UMCNetAfter = 2
    UMCNetFirst = 3
    UMCNetLast = 4
End Enum

Public Enum UMCRobustNETIncrementConstants
    UMCRobustNETIncrementLinear = 0
    UMCRobustNETIncrementPercentage = 1
End Enum

Public Enum UMCRobustNETModeConstants
    UMCRobustNETIterative = 0
    UMCRobustNETWarpTime = 1                ' Uses MSAlign
    UMCRobustNETWarpTimeAndMass = 2         ' Uses MSAlign
End Enum

Public Enum rmcUMCRobustNETWarpMassCalibrationType
    rmcMZRegressionRecal = 0
    rmcScanRegressionRecal = 1
    rmcHybridRecal = 2
End Enum

Public Enum UMCManageConstants
    UMCMngInitialize = 0
    UMCMngErase = 1
    UMCMngTrim = 2
    UMCMngAdd = 3
End Enum

Public Const UMC_STATISTICS1_MAX_INDEX As Integer = 18
Public Enum ustUMCStatistics1Constants
    ustClassIndex = 0
    ustClassRepMW = 1
    ustScanStart = 2
    ustScanEnd = 3
    ustClassMW = 4
    ustClassMWStDev = 5
    ustClassIntensity = 6
    ustFitAverage = 7
    ustClassRepIntensity = 8
    ustAbundanceMin = 9
    ustAbundanceMax = 10
    ustChargeMin = 11
    ustChargeMax = 12
    ustFitMin = 13
    ustFitMax = 14
    ustFitStDev = 15
    ustMassMin = 16
    ustMassMax = 17
    ustMassStDev = 18
End Enum

'definition type for the Unique Masses Classes
'This corresponds to FileInfoVersions(fioSearchDefinitions) version 2
'and to FileInfoVersions(fioGelUMC) version 1
Public Type UMCDefinition2002
    UMCType As Integer                      'type
    DefScope As Integer
    MWField As Integer                      'mass type (isfMWAvg or isfMWMono or isfMWTMA)
    TolType As Integer                      'MW width type
    Tol As Double                           'MW width
    UMCSharing As Boolean                   'members sharing
    UMCUniCS As Boolean                     'uni-charge states
    ClassAbu As Integer                     'abundance
    ClassMW As Integer                      'mass
    GapMaxCnt As Long                       'max gap count
    GapMaxSize As Long                      'max gap size
    GapMaxPct As Double                     'max gap percentage
    UMCNETType As Integer                   'NET type
    UMCMaxAbuPctBf As Double                'max abundance percentage threshold before
    UMCMaxAbuPctAf As Double                'max abundance percentage threshold after
    UMCMaxAbuEtPctBf As Double              'max abundance elution percentage before
    UMCMaxAbuEtPctAf As Double              'max abundance elution percentage after
    UMCMinCnt As Long                       'min count
    UMCMaxCnt As Long                       'max count
End Type

'definition type for the Unique Masses Classes
'This corresponds to FileInfoVersions(fioSearchDefinitions) version 3
'and to FileInfoVersions(fioGelUMC) version 2
Public Type UMCDefinition2003a
    UMCType As Integer                      'type
    DefScope As Integer
    MWField As Integer                      'mass type (isfMWAvg or isfMWMono or isfMWTMA)
    TolType As Integer                      'MW width type                          (actually type glMassToleranceConstants)
    Tol As Double                           'MW width
    UMCSharing As Boolean                   'members sharing
    UMCUniCS As Boolean                     'uni-charge states      (Require identical charge state for members of UMC; not currently implemented)
    ClassAbu As Integer                     'class abundance
    ClassMW As Integer                      'class mass
    GapMaxCnt As Long                       'max gap count
    GapMaxSize As Long                      'max gap size
    GapMaxPct As Double                     'max gap percentage     (UMCListType2002 Only)
    UMCNETType As Integer                   'NET type               (Not used; June 2003, MEM))
    UMCMaxAbuPctBf As Double                'max abundance percentage threshold before      (Not used)
    UMCMaxAbuPctAf As Double                'max abundance percentage threshold after       (Not used)
    UMCMaxAbuEtPctBf As Double              'max abundance elution percentage before        (Not used)
    UMCMaxAbuEtPctAf As Double              'max abundance elution percentage after         (Not used)
    UMCMinCnt As Long                       'min count
    UMCMaxCnt As Long                       'max count
    'use following members in class abundance calculation to get better class abundance
    InterpolateGaps As Boolean
    InterpolateMaxGapSize As Long
    InterpolationType As Long               ' Currently only one interpolation method: 0
End Type
       
'definition type for the Unique Masses Classes
'This corresponds to FileInfoVersions(fioSearchDefinitions) version 4
'and to FileInfoVersions(fioGelUMC) versions 3 and 4
Public Type UMCDefinition
    UMCType As Integer                      'type; Options include glUMC_TYPE_INTENSITY = 0; glUMC_TYPE_FIT = 1; glUMC_TYPE_MINCNT = 2; glUMC_TYPE_MAXCNT = 3; glUMC_TYPE_UNQAMT = 4; glUMC_TYPE_ISHRINKINGBOX = 5; glUMC_TYPE_FSHRINKINGBOX = 6; glUMC_TYPE_FROM_NET = 7
    DefScope As Integer
    MWField As Integer                      'mass type (isfMWAvg or isfMWMono or isfMWTMA)
    TolType As Integer                      'MW width type                          (actually type glMassToleranceConstants)
    Tol As Double                           'MW width
    UMCSharing As Boolean                   'members sharing
    UMCUniCS As Boolean                     'Unused; uni-charge states      (Require identical charge state for members of UMC)
    ClassAbu As Integer                     'class abundance        (type UMCClassAbundanceConstants)
    ClassMW As Integer                      'class mass             (type UMCClassMassConstants)
    GapMaxCnt As Long                       'max gap count          (UMC2003 only)
    GapMaxSize As Long                      'max gap size           (Applies to "Maximum size of Scan Gap" in UMC2003; applies to SplitUMCs in UMCIonNet)
    GapMaxPct As Double                     'max gap percentage     (UMCListType2002 Search mode Only)
    UMCNETType As Integer                   'NET type               (Not used; June 2003, MEM))
    UMCMaxAbuPctBf As Double                'max abundance percentage threshold before      (Not used)
    UMCMaxAbuPctAf As Double                'max abundance percentage threshold after       (Not used)
    UMCMaxAbuEtPctBf As Double              'max abundance elution percentage before        (Not used)
    UMCMaxAbuEtPctAf As Double              'max abundance elution percentage after         (Not used)
    UMCMinCnt As Long                       'min count
    UMCMaxCnt As Long                       'max count
    'use following members in class abundance calculation to get better class abundance
    InterpolateGaps As Boolean
    InterpolateMaxGapSize As Long
    InterpolationType As Long               ' Currently only one interpolation method: 0
    ' Charge state group options
    ChargeStateStatsRepType As Integer                              ' New for this version; Method to use to determine the representative charge state: Options include UMCChargeStateGroupConstants.UMCCSGHighestSum = 0, UMCChargeStateGroupConstants.UMCCSGMostAbuMember = 1, UMCChargeStateGroupConstants.UMCCSGMostMembers = 2
    UMCClassStatsUseStatsFromMostAbuChargeState As Boolean          ' New for this version
    OtherInfo As String                                             ' New for this version
End Type

' The following is used for Net Adjustment
' This corresponds to FileInfoVersions(fioUMCNetAdjDef) version 1
Public Type NetAdjDefinition2003
    MinUMCCount As Long             'minimum number of peaks in UMC
    MinScanRange As Long            'minimum number of different scans
    MaxScanPct  As Double           'maximum percentage of UMC scan range in total scans (unused)
    TopAbuPct As Double             'if > 0 use only top TopAbuPct LC-MS Features
    PeakSelection As Long           'before/at/after highest abundance; first peak; last peak
    PeakMaxAbuPct As Double         'used together with PeakMaxAbuSelection to
                                    'select actual peak
    PeakCSSelection(7) As Boolean   'acceptable charge states 1-6,6+,Any
    MWField As Long                 'which monoisotopic field to use for Net Adjustment     (Not used)
    MWTolType As Long
    MWTol As Double                 'in ppm
    NETFormula As String
    NETTol As Double                'in pct
    NETorRT As Long                 'normalized elution or retention
    UseNET As Boolean               'if True use NET in search
    UseMultiIDMaxNETDist As Boolean
    MultiIDMaxNETDist As Double     'eliminate identifications from the same peak that are too far appart in NET
    EliminateBadNET As Boolean      'if True don't use IDs with NETs out of range [0,1]
    MaxIDToUse As Long              'limits the number of IDs to use (use best)
    IterationStopType As Long       'when to stop iteration
    IterationStopValue As Double    'value to watch to stop iteration
    IterationUseMWDec As Boolean
    IterationMWDec As Double        'value to decrease MW in each step
    IterationUseNETdec As Boolean   'to use it or not to use it? that is the question!
    IterationNETDec As Double       'value to decrease NET in each step
    IterationAcceptLast As Boolean  'if True accept last iteration for NET adjustment
End Type
       
' The following is used for Net Adjustment
' This corresponds to FileInfoVersions(fioUMCNetAdjDef) version 3
Public Type NetAdjDefinition2004
    MinUMCCount As Long             'minimum number of peaks in UMC
    MinScanRange As Long            'minimum number of different scans
    MaxScanPct  As Double           'maximum percentage of UMC scan range in total scans (unused)
    TopAbuPct As Double             'if > 0 use only top TopAbuPct LC-MS Features
    PeakSelection As Long           'before/at/after highest abundance; first peak; last peak
    PeakMaxAbuPct As Double         'used together with PeakMaxAbuSelection to
                                    'select actual peak
    PeakCSSelection(7) As Boolean   'acceptable charge states 1-6,6+,Any
    MWField As Long                 'which monoisotopic field to use for Net Adjustment     (Not used)
    MWTolType As Long
    MWTol As Double                 'in ppm
    NETFormula As String
    NETTol As Double                'in pct
    NETorRT As Long                 'normalized elution or retention
    UseNET As Boolean               'if True use NET in search
    UseMultiIDMaxNETDist As Boolean
    MultiIDMaxNETDist As Double     'eliminate identifications from the same peak that are too far appart in NET
    EliminateBadNET As Boolean      'if True don't use IDs with NETs out of range [0,1]
    MaxIDToUse As Long              'limits the number of IDs to use (use best)
    IterationStopType As Long       'when to stop iteration
    IterationStopValue As Double    'value to watch to stop iteration
    IterationUseMWDec As Boolean
    IterationMWDec As Double        'value to decrease MW in each step
    IterationUseNETdec As Boolean   'to use it or not to use it? that is the question!
    IterationNETDec As Double       'value to decrease NET in each step
    IterationAcceptLast As Boolean  'if True accept last iteration for NET adjustment

    InitialSlope As Double          ' Default slope                 ' New for this version
    InitialIntercept As Double      ' Default intercept             ' New for this version
    UseNetAdjLockers As Boolean                                     ' New for this version
    UseOldNetAdjIfFailure As Boolean                                ' New for this version
    NetAdjLockerMinimumMatchCount As Integer                        ' New for this version

    OtherInfo As String                                             ' New for this version
End Type

' The following is used for Net Adjustment
' This corresponds to FileInfoVersions(fioUMCNetAdjDef) version 4
Public Type NetAdjDefinition2005a
    MinUMCCount As Long             'minimum number of peaks in UMC
    MinScanRange As Long            'minimum number of different scans
    MaxScanPct  As Double           'maximum percentage of UMC scan range in total scans (unused)
    TopAbuPct As Double             'if > 0 use only top TopAbuPct LC-MS Features
    PeakSelection As Long           'Ignored: before/at/after highest abundance; first peak; last peak; no longer used (always assume at Max abu)
    PeakMaxAbuPct As Double         'Ignored: used together with PeakMaxAbuSelection to select actual peak
    PeakCSSelection(7) As Boolean   'acceptable charge states 1-6,6+,Any
    MWTolType As Long               'Actually type glMassToleranceConstants
    MWTol As Double                 'in ppm
    NETFormula As String
    NETTol As Double                'in pct
    NETorRT As Long                 'normalized elution or retention
    UseNET As Boolean               'if True use NET in search
    UseMultiIDMaxNETDist As Boolean
    MultiIDMaxNETDist As Double     'eliminate identifications from the same peak that are too far appart in NET
    EliminateBadNET As Boolean      'if True don't use IDs with NETs out of range [0,1]
    MaxIDToUse As Long              'limits the number of IDs to use (use best)
    
    IterationStopType As Long       'when to stop iteration
    IterationStopValue As Double    'value to watch to stop iteration
    IterationUseMWDec As Boolean
    IterationMWDec As Double        'value to decrease MW in each step
    IterationUseNETdec As Boolean   'to use it or not to use it? that is the question!
    IterationNETDec As Double       'value to decrease NET in each step
    IterationAcceptLast As Boolean  'if True accept last iteration for NET adjustment

    InitialSlope As Double          ' Default slope                 ' Only used if Robust NET searching is disabled
    InitialIntercept As Double      ' Default intercept             ' Only used if Robust NET searching is disabled
    UseNetAdjLockers As Boolean
    UseOldNetAdjIfFailure As Boolean
    NetAdjLockerMinimumMatchCount As Integer

    ' The following are all New for this version
    UseRobustNETAdjustment As Boolean
    RobustNETAdjustmentMode As Integer          ' Actually type UMCRobustNETModeConstants
    
    RobustNETSlopeStart As Single
    RobustNETSlopeEnd As Single
    RobustNETSlopeIncreaseMode As Integer       ' Actually type UMCRobustNETIncrementConstants; 0 to increase in absolute steps, 1 to increase by a percentage of the current value
    RobustNETSlopeIncrement As Single
    
    RobustNETInterceptStart As Single
    RobustNETInterceptEnd As Single
    RobustNETInterceptIncrement As Single       ' The NET Intercept is always incremented in absolute steps
    
    RobustNETMassShiftPPMStart As Single
    RobustNETMassShiftPPMEnd As Single
    RobustNETMassShiftPPMIncrement As Single    ' The Mass Shift is always incremented in absolute steps
    
    ' The following are reserved for future use by simulated annealing NET adjustment
    RobustNETAnnealSteps As Long                            ' Typically 20
    RobustNETAnnealTrialsPerStep As Long                    ' Typically 250
    RobustNETAnnealMaxSwapsPerStep As Long                  ' Typically 50 (maximum good swaps at each step)
    RobustNETAnnealTemperatureReductionFactor As Single     ' Typically 0.9
    
    OtherInfo As String
End Type

' The following is used for Net Adjustment
' This corresponds to FileInfoVersions(fioUMCNetAdjDef) version 5
Public Type NetAdjDefinition2005b
    MinUMCCount As Long             'minimum number of peaks in UMC
    MinScanRange As Long            'minimum number of different scans
    MaxScanPct  As Double           'maximum percentage of UMC scan range in total scans (unused)
    TopAbuPct As Double             'if > 0 use only top TopAbuPct LC-MS Features
    PeakSelection As Long           'Ignored: before/at/after highest abundance; first peak; last peak; no longer used (always assume at Max abu)
    PeakMaxAbuPct As Double         'Ignored: used together with PeakMaxAbuSelection to select actual peak
    PeakCSSelection(7) As Boolean   'acceptable charge states 1-6,6+,Any
    MWTolType As Long               'Actually type glMassToleranceConstants
    MWTol As Double                 'in ppm
    NETFormula As String
    NETTolIterative As Double       'in pct
    NETorRT As Long                 'normalized elution or retention
    UseNET As Boolean               'if True use NET in search
    UseMultiIDMaxNETDist As Boolean
    MultiIDMaxNETDist As Double     'eliminate identifications from the same peak that are too far appart in NET
    EliminateBadNET As Boolean      'if True don't use IDs with NETs out of range [0,1]
    MaxIDToUse As Long              'limits the number of IDs to use (use best)
    
    IterationStopType As Long       'when to stop iteration
    IterationStopValue As Double    'value to watch to stop iteration
    IterationUseMWDec As Boolean
    IterationMWDec As Double        'value to decrease MW in each step
    IterationUseNETdec As Boolean   'to use it or not to use it? that is the question!
    IterationNETDec As Double       'value to decrease NET in each step
    IterationAcceptLast As Boolean  'if True accept last iteration for NET adjustment

    InitialSlope As Double          ' Default slope                 ' Only used if Robust NET searching is disabled
    InitialIntercept As Double      ' Default intercept             ' Only used if Robust NET searching is disabled
    UseNetAdjLockers As Boolean
    UseOldNetAdjIfFailure As Boolean
    NetAdjLockerMinimumMatchCount As Integer

    UseRobustNETAdjustment As Boolean
    RobustNETAdjustmentMode As Integer          ' Actually type UMCRobustNETModeConstants
    
    RobustNETSlopeStart As Single
    RobustNETSlopeEnd As Single
    RobustNETSlopeIncreaseMode As Integer       ' Actually type UMCRobustNETIncrementConstants; 0 to increase in absolute steps, 1 to increase by a percentage of the current value
    RobustNETSlopeIncrement As Single
    
    RobustNETInterceptStart As Single
    RobustNETInterceptEnd As Single
    RobustNETInterceptIncrement As Single       ' The NET Intercept is always incremented in absolute steps
    
    RobustNETMassShiftPPMStart As Single
    RobustNETMassShiftPPMEnd As Single
    RobustNETMassShiftPPMIncrement As Single    ' The Mass Shift is always incremented in absolute steps
    
   ' The following were previously reserved for simulated annealing NET adjustment, but are now unused
    AdditionalValue1 As Long
    AdditionalValue2 As Long
    AdditionalValue3 As Long
    AdditionalValue4 As Single
    
    ' The following are all New for this version
    ' The following apply to frmMSAlign
    RobustNETWarpNumberOfSections As Long                   ' Typically 100
    RobustNETWarpMaxDistortion As Long                      ' Typically 3
    RobustNETWarpContractionFactor As Long                  ' Typically 2
    RobustNETWarpMinimumPMTTagObsCount As Long              ' Typically 5
    RobustNETWarpNETTol As Single                           ' Typically 0.02
    RobustNETWarpMatchPromiscuity As Integer                ' Typically 2
    
    AdditionalValue5 As Integer                             ' No Longer Used: generic variable name
    RobustNETWarpMassCalibrationType As Integer             ' Actually type rmcUMCRobustNETWarpMassCalibrationType
    RobustNETWarpMassSplineOrder As Integer                 ' Typically 2; on file load, if greater than 10, then resets these values to defaults
    RobustNETWarpMassWindowPPM As Single                    ' Typically 50
    RobustNETWarpMassNumXSlices As Integer                  ' Typically 20
    RobustNETWarpMassNumMassDeltaBins As Integer            ' Typically 100
    RobustNETWarpMassMaxJump As Integer                     ' Typically 50
    
    AdditionalValue8 As Integer      ' 2 bytes; New for this version; use for future expansion (name can be changed in the future)
    
    OtherInfo As String
End Type

Public Type udtMSWarpOptionsType
    MassCalibrationType As Integer             ' Actually type rmcUMCRobustNETWarpMassCalibrationType
    MinimumPMTTagObsCount As Long              ' Typically 5
    MatchPromiscuity As Integer                ' Typically 2
    
    NETTol As Single                           ' Typically 0.02
    NumberOfSections As Long                   ' Typically 100
    MaxDistortion As Integer                   ' Typically 10
    ContractionFactor As Integer               ' Typically 3
    
    MassWindowPPM As Single                    ' Typically 50
    MassSplineOrder As Integer                 ' Typically 2
    MassNumXSlices As Integer                  ' Typically 20
    MassNumMassDeltaBins As Integer            ' Typically 100
    MassMaxJump As Integer                     ' Typically 50
    
    MassZScoreTolerance As Single              ' Typically 3
    MassUseLSQ As Boolean                      ' Typically True
    MassLSQOutlierZScore As Single             ' Typically 3
    MassLSQNumKnots As Integer                 ' Typically 12
        
    AdditionalValue1 As Long        ' 4 bytes; use for future expansion (name can be changed in the future)
    AdditionalValue2 As Long        ' 4 bytes
    AdditionalValue3 As Long        ' 4 bytes
    AdditionalValue4 As Long        ' 4 bytes
    AdditionalValue5 As Long        ' 4 bytes
    AdditionalValue6 As Long        ' 4 bytes
    AdditionalValue7 As Long        ' 4 bytes
    AdditionalValue8 As Long        ' 4 bytes
    AdditionalValue9 As Long        ' 4 bytes
    AdditionalValue10 As Long       ' 4 bytes
End Type

' The following is used for Net Adjustment
' This corresponds to FileInfoVersions(fioUMCNetAdjDef) version 6
Public Type NetAdjDefinition
    MinUMCCount As Long             'minimum number of peaks in UMC
    MinScanRange As Long            'minimum number of different scans
    MaxScanPct  As Double           'maximum percentage of UMC scan range in total scans (unused)
    TopAbuPct As Double             'if > 0 use only top TopAbuPct LC-MS Features
    PeakSelection As Long           'Ignored: before/at/after highest abundance; first peak; last peak; no longer used (always assume at Max abu)
    PeakMaxAbuPct As Double         'Ignored: used together with PeakMaxAbuSelection to select actual peak
    PeakCSSelection(7) As Boolean   'acceptable charge states 1-6,6+,Any
    MWTolType As Long               'Actually type glMassToleranceConstants
    MWTol As Double                 'in ppm
    NETFormula As String            'typically holds the default NET formula with default slope and intercept; see GelAnalysisInfo().GANET_Slope and GelAnalysisInfo().GANET_Intercept for the official values
    NETTolIterative As Double       'in pct
    NETorRT As Long                 'normalized elution or retention
    UseNET As Boolean               'if True use NET in search
    UseMultiIDMaxNETDist As Boolean
    MultiIDMaxNETDist As Double     'eliminate identifications from the same peak that are too far appart in NET
    EliminateBadNET As Boolean      'if True don't use IDs with NETs out of range [0,1]
    MaxIDToUse As Long              'limits the number of IDs to use (use best)
    
    IterationStopType As Long       'when to stop iteration
    IterationStopValue As Double    'value to watch to stop iteration
    IterationUseMWDec As Boolean
    IterationMWDec As Double        'value to decrease MW in each step
    IterationUseNETdec As Boolean   'to use it or not to use it? that is the question!
    IterationNETDec As Double       'value to decrease NET in each step
    IterationAcceptLast As Boolean  'if True accept last iteration for NET adjustment

    InitialSlope As Double          ' Default slope                 ' Only used if Robust NET searching is disabled
    InitialIntercept As Double      ' Default intercept             ' Only used if Robust NET searching is disabled
    UseNetAdjLockers As Boolean                 ' No longer supported: March 2006; Use the Internal Standards for NET alignment
    UseOldNetAdjIfFailure As Boolean            ' No longer supported: March 2006
    NetAdjLockerMinimumMatchCount As Integer    ' No longer supported: March 2006

    UseRobustNETAdjustment As Boolean           ' To run MSAlign, this must be set to True
    RobustNETAdjustmentMode As Integer          ' Actually type UMCRobustNETModeConstants
    
    RobustNETSlopeStart As Single
    RobustNETSlopeEnd As Single
    RobustNETSlopeIncreaseMode As Integer       ' Actually type UMCRobustNETIncrementConstants; 0 to increase in absolute steps, 1 to increase by a percentage of the current value
    RobustNETSlopeIncrement As Single
    
    RobustNETInterceptStart As Single
    RobustNETInterceptEnd As Single
    RobustNETInterceptIncrement As Single       ' The NET Intercept is always incremented in absolute steps
    
    RobustNETMassShiftPPMStart As Single
    RobustNETMassShiftPPMEnd As Single
    RobustNETMassShiftPPMIncrement As Single    ' The Mass Shift is always incremented in absolute steps
    
    ' The following are all New for this version
    AdditionalValue1 As Long        ' 4 bytes
    AdditionalValue2 As Long        ' 4 bytes
    AdditionalValue3 As Long        ' 4 bytes
    AdditionalValue4 As Long        ' 4 bytes
    AdditionalValue5 As Long        ' 4 bytes
    AdditionalValue6 As Long        ' 4 bytes
    
    ' This applies to frmMSAlign
    MSWarpOptions As udtMSWarpOptionsType
    
    OtherInfo As String
End Type

Public Type LaV2DGPoint
    MW As Double
    Scan As Double
    Abu As Double
End Type
                            
'Next two variables live as long as application lives; application starts
' with default values and keep track of all changes until the end of session
'Note that UMCIonNetDef is defined in UMCIonNet.bas, and is similar to UMCDef
Public UMCDef As UMCDefinition
Public UMCNetAdjDef As NetAdjDefinition

'module variables just for simplicity of calling procedures
'used with UMCFitIntensityCount and FindMWRange subs
Dim MWZero As Double
Dim FNZero As Long
Dim AbsTol As Double

Dim ShrinkingBox_MW_Average_Type As Integer
Dim SBValidPoints() As Long
Dim ShrunkBoxes As Long
Dim UMCsDone As Long

'UMC variables
Dim OID() As Long           'original index
Dim ODT() As Integer        'original deconvolution type
Dim OMW() As Double         'original mw (used also in Fast Unique Count)
Dim OAbu() As Double        'original abundance
Dim OFN() As Integer        'original file(scan) number
Dim UMC() As Long           'classification array; contains OID(i) if original
                            '-OID(i) if duplicate of OID(i) and 0 otherwise
                            'after the count there should be no 0

'molecular mass index
Dim IndMW() As Long         'index of original index
Dim MW() As Double          'molecular masses
'order index (intensity, fit or something else)
Dim IndOR() As Long         'index of original index
Dim ORFld() As Double       'values from ordering field
'relative FN numbers of the duplicate candidates for current MWZero
Dim IndFN() As Long
Dim RelFN() As Long

Public Function AutoRefineUMCs(ByVal lngGelIndex As Long, ByRef frmCallingForm As VB.Form) As Boolean
'---------------------------------------------------------
' This function is called by frmUMCWithAutoRefine, frmUMCSimple, and frmUMCIonNet
' Returns True if the UMCIndices were updated (via the UpdateUMCStatArrays function) during auto-refinement
' Returns False if the UMCIndices were not updated
'---------------------------------------------------------
    
    Dim blnUMCIndicesUpdated As Boolean
    
On Error GoTo AutoRefineUMCsErrorHandler

    blnUMCIndicesUpdated = False
    With glbPreferencesExpanded.UMCAutoRefineOptions
        If GelUMC(lngGelIndex).UMCCnt > 0 And _
           (.UMCAutoRefineRemoveAbundanceHigh Or .UMCAutoRefineRemoveAbundanceLow Or _
            .UMCAutoRefineRemoveCountLow Or .UMCAutoRefineRemoveCountHigh Or _
            .UMCAutoRefineRemoveMaxLengthPctAllScans) Then
            
            ' Auto refine the results
            ' Use frmVisUMC to do this
            
            ' Initialize frmVisUMC
            frmCallingForm.Status "Auto-refining LC-MS Features: Initializing"
        
            frmVisUMC.Tag = lngGelIndex
            frmVisUMC.InitializeUMCs
        
            frmVisUMC.chkRemovePairedLUMC = vbUnchecked
            frmVisUMC.chkRemovePairedHUMC = vbUnchecked
        
            glbPreferencesExpanded.AutoAnalysisStatus.AutoRefiningUMCs = True
            frmCallingForm.Status "Auto-refining LC-MS Features: Refining"
            frmVisUMC.AutoRemoveUMCsWork
        
            frmCallingForm.Status "Auto-refining LC-MS Features: Cleaning up"
            Unload frmVisUMC
            
            glbPreferencesExpanded.AutoAnalysisStatus.AutoRefiningUMCs = False
            
            ' Note that the AutoRefineUMCsWork function uses frmVisUMC to do the actual refinement
            ' When the form is unloaded, the UpdateUMCStatArrays function is called
            ' Thus, there is no need to call it again later in this function
            If GelUMC(lngGelIndex).UMCCnt > 0 Then
                blnUMCIndicesUpdated = True
            End If
        End If
    End With

    AutoRefineUMCs = blnUMCIndicesUpdated
    Exit Function

AutoRefineUMCsErrorHandler:
    Debug.Print "Error in AutoRefineUMCs: " & Err.Description
    Debug.Assert False

End Function

' Unused Procedure (February 2005)
''Public Function UMCCount(ByVal Ind As Long, _
''                         ByRef TtlCnt As Long, _
''                         ByRef frmCallingForm As VB.Form, _
''                         Optional ByVal blnOptimisticSearching As Boolean = False) As Long
'''---------------------------------------------------------
'''fills unique mass classes structure and returns number of
'''unique mass classes; -1 on any error - this is basicaly
'''unique count in the gel(different than FastUniqueCount)
'''---------------------------------------------------------
''Dim qsSort As QSDouble
''On Error GoTo err_UMCCount
''
''frmCallingForm.Status "Loading Arrays."
''UMCLoadArrays Ind, TtlCnt
''If TtlCnt > 1 Then
'''still don't have algorithms for MaxCount & MinCount
'''it will probably ask for approach different from sorting
''   Select Case UMCDef.UMCType
''   Case glUMC_TYPE_FIT, glUMC_TYPE_INTENSITY, glUMC_TYPE_ISHRINKINGBOX
''     frmCallingForm.Status "Sorting Index Arrays."
''     Set qsSort = New QSDouble
''     If Not qsSort.QSAsc(MW(), IndMW()) Then
''        MsgBox "Error sorting index arrays. Aborting Unique Count.", vbOKOnly
''        GoTo err_UMCCount
''     End If
''     Set qsSort = Nothing
''     Set qsSort = New QSDouble
''     If UMCDef.UMCType = glUMC_TYPE_FIT Then     'sort ascending
''        If Not qsSort.QSAsc(ORFld(), IndOR()) Then
''           MsgBox "Error sorting index arrays. Aborting Unique Count.", vbOKOnly
''           GoTo err_UMCCount
''        End If
''     Else        'Intensity - sort descending
''        If Not qsSort.QSDesc(ORFld(), IndOR()) Then
''           MsgBox "Error sorting index arrays. Aborting Unique Count.", vbOKOnly
''           GoTo err_UMCCount
''        End If
''     End If
''     Set qsSort = Nothing
''     'reserve space for 20,000 classes; then add 5000 as needed
''     With GelUMC(Ind)
''         .UMCCnt = 0
''         ReDim .UMCs(20000)
''         .MassCorrectionValuesDefined = False
''     End With
''
''     If UMCDef.UMCType = glUMC_TYPE_ISHRINKINGBOX Then
''        '-----------------------  Added 4/18/2002 KL ---------------------------------
''        ' Turns out to be same as above for intensity sorting, except for last call to
''        ' function UMCFitIntensityWithShrinkingBox(Ind, TtlCnt)
''        If Not UMCFitIntensityWithShrinkingBox(Ind, TtlCnt, frmCallingForm) Then GoTo err_UMCCount
''     Else
''        If Not UMCFitIntensityCount(Ind, TtlCnt, frmCallingForm, blnOptimisticSearching) Then GoTo err_UMCCount
''     End If
''   Case glUMC_TYPE_MINCNT
''     If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then MsgBox "Not implemented.", vbOKOnly
''     GoTo err_UMCCount
''   Case glUMC_TYPE_MAXCNT
''     If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then MsgBox "Not implemented.", vbOKOnly
''     GoTo err_UMCCount
''   Case glUMC_TYPE_UNQAMT
''     If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then MsgBox "Not implemented.", vbOKOnly
''     GoTo err_UMCCount
''   Case Else
''     If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then MsgBox "Not implemented.", vbOKOnly
''     GoTo err_UMCCount
''   End Select
''ElseIf TtlCnt = 1 Then
''   With GelUMC(Ind)
''     .UMCCnt = 1
''     ReDim .UMCs(0)
''     .UMCs(0).ClassCount = 1
''     .UMCs(0).ClassRepInd = OID(1)
''     .UMCs(0).ClassRepType = ODT(1)
''     .UMCs(0).ClassMInd(0) = OID(1)
''     .UMCs(0).ClassMType(0) = ODT(1)
''     .UMCs(0).ClassAbundance = OAbu(1)
''     .MassCorrectionValuesDefined = False
''   End With
''Else
''   GelUMC(Ind).UMCCnt = 0
''End If
''
''' Need to update .Def before calling CalculateClasses
''GelUMC(Ind).def = UMCDef
''
''frmCallingForm.Status "Computing UMC Statistics"
''CalculateClasses Ind, False, frmCallingForm
''
''With GelUMC(Ind)
''   'redimension classes array here
''   If .UMCCnt >= 1 Then
''      ReDim Preserve .UMCs(.UMCCnt - 1)
''   Else
''      ReDim .UMCs(0)
''   End If
''   UMCCount = .UMCCnt
''End With
''
''exit_UMCCount:
'''destroy arrays to free memory
''Erase OID
''Erase ODT
''Erase OMW
''Erase OAbu
''Erase OFN
''Erase UMC
''Erase IndMW
''Erase MW
''Erase IndOR
''Erase ORFld
''Erase IndFN
''Erase RelFN
''Exit Function
''
''err_UMCCount:
''Set qsSort = Nothing
''UMCCount = -1
''GoTo exit_UMCCount
''End Function
''
' Unused Procedure (February 2005)
''Private Sub UMCLoadArrays(ByVal Ind As Long, _
''                          ByRef TtlCnt As Long)
'''----------------------------------------------
'''load arrays neccessary for this procedure
'''----------------------------------------------
''Dim NewTtlCnt As Long
''Dim i As Long
''Dim CSCnt As Long
''Dim CSInd() As Long
''Dim ISCnt As Long
''Dim ISInd() As Long
''On Error Resume Next
''
''If TtlCnt > 0 Then
''   ReDim OID(1 To TtlCnt)
''   ReDim ODT(1 To TtlCnt)
''   ReDim OMW(1 To TtlCnt)
''   ReDim OAbu(1 To TtlCnt)
''   ReDim OFN(1 To TtlCnt)
''   ReDim ORFld(1 To TtlCnt)
''   NewTtlCnt = 0
''   With GelData(Ind)
''     CSCnt = GetCSScope(Ind, CSInd(), UMCDef.DefScope)
''     If CSCnt > 0 Then
''        For i = 1 To CSCnt
''            NewTtlCnt = NewTtlCnt + 1
''            OID(NewTtlCnt) = CSInd(i)
''            ODT(NewTtlCnt) = gldtCS
''            OMW(NewTtlCnt) = .CSData(CSInd(i)).AverageMW
''            OAbu(NewTtlCnt) = .CSData(CSInd(i)).Abundance
''            OFN(NewTtlCnt) = .CSData(CSInd(i)).ScanNumber
''            Select Case UMCDef.UMCType
''            Case glUMC_TYPE_INTENSITY
''              ORFld(NewTtlCnt) = .CSData(CSInd(i)).Abundance
''            Case glUMC_TYPE_FIT
''              ORFld(NewTtlCnt) = .CSData(CSInd(i)).MassStDev 'St.Dev. in fact
''            '------ Added 4/18/2002 KL ------
''            Case glUMC_TYPE_ISHRINKINGBOX
''              ORFld(NewTtlCnt) = .CSData(CSInd(i)).Abundance
''            '--------------------------------
''            End Select
''        Next i
''     End If
''     ISCnt = GetISScope(Ind, ISInd(), UMCDef.DefScope)
''     If ISCnt > 0 Then
''        For i = 1 To ISCnt
''            NewTtlCnt = NewTtlCnt + 1
''            OID(NewTtlCnt) = ISInd(i)
''            ODT(NewTtlCnt) = gldtIS
''            OMW(NewTtlCnt) = .IsoNum(ISInd(i), UMCDef.MWField)
''            OAbu(NewTtlCnt) = .IsoData(ISInd(i)).Abundance
''            OFN(NewTtlCnt) = .IsoData(ISInd(i)).ScanNumber
''            Select Case UMCDef.UMCType
''            Case glUMC_TYPE_INTENSITY
''              ORFld(NewTtlCnt) = .IsoData(ISInd(i)).Abundance
''            Case glUMC_TYPE_FIT
''              ORFld(NewTtlCnt) = .IsoData(ISInd(i)).Fit
''            '------  Added 4/18/2002 KL ------
''            Case glUMC_TYPE_ISHRINKINGBOX
''              ORFld(NewTtlCnt) = .IsoData(ISInd(i)).Abundance
''            '---------------------------------
''            End Select
''        Next i
''     End If
''   End With
''   TtlCnt = NewTtlCnt
''   If TtlCnt > 0 Then
''      ReDim Preserve OID(1 To TtlCnt)
''      ReDim Preserve ODT(1 To TtlCnt)
''      ReDim Preserve OMW(1 To TtlCnt)
''      ReDim Preserve OAbu(1 To TtlCnt)
''      ReDim Preserve OFN(1 To TtlCnt)
''      ReDim Preserve ORFld(1 To TtlCnt)
''      'copy OMW array to MW (new thing in VB6 - not using CopyMemory)
''      MW() = OMW()
''      'initialize index arrays
''      ReDim IndMW(1 To TtlCnt)
''      ReDim IndOR(1 To TtlCnt)
''      ReDim UMC(1 To TtlCnt)
''      For i = 1 To TtlCnt
''          IndMW(i) = i
''      Next i
''      IndOR() = IndMW()
''   Else
''      Erase OID
''      Erase ODT
''      Erase OMW
''      Erase OAbu
''      Erase OFN
''      Erase MW
''      Erase ORFld
''   End If
''End If
''End Sub
''
' Unused Procedure (February 2005)
''Private Function UMCFitIntensityCount(ByVal Ind As Long, _
''                                      ByVal TtlCnt As Long, _
''                                      ByRef frmCallingForm As VB.Form, _
''                                      Optional ByVal blnOptimisticSearching As Boolean = False) As Boolean
'''---------------------------------------------------------------------
'''do the real Unique Mass Classes count; returns True if successful
'''NOTE: this function is called only for TtlCnt>1
'''actual classes are filled in HolePattern function
'''---------------------------------------------------------------------
''Dim i As Long               'counter through sorted IndOR/ORFld arrays
''Dim RefInd As Long          'reference index in original arrays
''
''Dim MinInd As Long          'indexes in IndMW/MW arrays of range of candidates
''Dim MaxInd As Long          'to be duplicate of class representative
''Dim FNZeroInd As Long       'index in IndFN/RelFN of current representative
''
''Dim Done As Boolean
''Dim Found1stUnmarked As Boolean
''Dim TmpCnt As Long
''
''Dim qslSort As New QSLong
''Dim j As Long
''Dim mwutUC As New MWUtil  'range finder
''On Error GoTo err_UMCFitIntensityCount
''
''i = 0
''TmpCnt = 0
'''MW is sorted ascending
''If Not mwutUC.Fill(MW()) Then GoTo err_UMCFitIntensityCount
''Do Until Done
''' MonroeMod Begin
''    If i Mod 25 = 0 Then frmCallingForm.Status CStr(i & " / " & TtlCnt)
''    DoEvents
''' MonroeMod Finish
''
''    Found1stUnmarked = False
''    'find first element of UMC that is 0 (unmarked)
''    Do Until Found1stUnmarked
''       i = i + 1
''       If i > TtlCnt Then
''          Done = True
''          Exit Do
''       Else
''          RefInd = IndOR(i)
''          If UMC(RefInd) = 0 Then Found1stUnmarked = True
''       End If
''    Loop
''    If Found1stUnmarked Then    'we have next unmarked
''       TmpCnt = TmpCnt + 1
''       UMC(RefInd) = RefInd     'so it is class representative
''       MWZero = OMW(RefInd)
''       FNZero = OFN(RefInd)
''       Select Case UMCDef.TolType
''       Case gltPPM
''            AbsTol = UMCDef.Tol * MWZero * glPPM
''       Case gltABS
''            AbsTol = UMCDef.Tol
''       End Select
''       MinInd = 1
''       MaxInd = TtlCnt
''       If mwutUC.FindIndexRange(MWZero, AbsTol, MinInd, MaxInd) Then
''          ReDim IndFN(MinInd To MaxInd)
''          ReDim RelFN(MinInd To MaxInd)
''          For j = MinInd To MaxInd
''              IndFN(j) = IndMW(j)
''              RelFN(j) = OFN(IndMW(j)) - FNZero 'distance from the FNZero
''          Next j
''          'sort it based on distance from the FNZero;
''          If Not qslSort.QSAsc(RelFN(), IndFN()) Then
''             GoTo err_UMCFitIntensityCount
''          End If
''          FNZeroInd = FindRelFNZero(RefInd)
''          HolePatterns Ind, MinInd, MaxInd, FNZeroInd, RefInd, blnOptimisticSearching
''       End If
''    End If
''    If glAbortUMCProcessing Then GoTo err_UMCFitIntensityCount
''Loop
''UMCFitIntensityCount = True
''
''exit_UMCFitIntensityCount:
''Set qslSort = Nothing
''Set mwutUC = Nothing
''Exit Function
''
''err_UMCFitIntensityCount:
''UMCFitIntensityCount = False
''LogErrors Err.Number, "UMCFitIntensityCount"
''GoTo exit_UMCFitIntensityCount
''End Function
''
' Unused Procedure (February 2005)
''Private Sub HolePatterns(ByVal Ind As Long, _
''                         ByVal MinInd As Long, _
''                         ByVal MaxInd As Long, _
''                         ByVal FNZeroInd As Long, _
''                         ByVal RefInd As Long, _
''                         Optional ByVal blnOptimisticSearching As Boolean = False)
'''-------------------------------------------------------------------
'''Fills class structure; Ind is Gel index
'''RelFN array contains potential duplicates of FNZero elements; here
'''some of them will be eliminated based on their relative positions
'''RelFN array is sorted ascending; MinInd and MaxInd are bounds of it
'''and FNZeroInd is index of UMC representative
'''When blnOptimisticSearching = True, then decrements TtlHlLen and
'''HlCnt when the hole count, hole size, or hole percent exceeds the
'''search limits, thereby allowing the search to continue further
'''-------------------------------------------------------------------
''Dim HlCnt As Long       'hole counter
''Dim TtlHlLen As Long    'total hole length
''Dim HlPct As Double     'percentage of the hole in the sequence
''Dim SeqLen As Long      'sequence length
''
''Dim Delta As Long
'''because nothing guarantees symetry have to work left & right separatelly
''Dim PosL As Long, PosR As Long
''Dim LDone As Boolean, RDone As Boolean
''Dim LHole As Boolean, RHole As Boolean
''Dim TmpClassCnt As Long
''Dim ClassInd As Long
''Dim MemberInd As Long
''Dim AbuSum As Double
''Dim MWSum As Double
''Dim TmpMW() As Double                   'contains all members classes
''Dim i As Long
''
''HlCnt = 0
''SeqLen = 1  'at least FNZero is in sequence
''With GelUMC(Ind)
''    .UMCCnt = .UMCCnt + 1                               'new class
''    ClassInd = .UMCCnt - 1
''    If .UMCCnt > UBound(.UMCs) Then                     'add room for new classes
''       ReDim Preserve .UMCs(.UMCCnt + 5000)
''    End If
''    .UMCs(ClassInd).ClassRepInd = OID(RefInd)
''    .UMCs(ClassInd).ClassRepType = ODT(RefInd)
''    .UMCs(ClassInd).ClassCount = 0                      'class members count
''    'put by default class abundance at class representative abundance
''    .UMCs(ClassInd).ClassAbundance = OAbu(RefInd)
''    'and class molecular mass at class representative molecular mass
''    .UMCs(ClassInd).ClassMW = OMW(RefInd)
''End With
''TtlHlLen = 0
''PosL = FNZeroInd
''PosR = FNZeroInd
''If FNZeroInd <= MinInd Then
''   LDone = True
''Else
''   LDone = False
''End If
''If FNZeroInd >= MaxInd Then
''   RDone = True
''Else
''   RDone = False
''End If
''LHole = False
''RHole = False
''Do Until (LDone And RDone)      'work until done on both sides
''   Do Until (LDone Or LHole)    'if hole on left side jump to the right
''      PosL = PosL - 1
''      If PosL < MinInd Then     'we are done on the left
''         PosL = MinInd
''         LDone = True
''      Else
''         Delta = Abs(RelFN(PosL + 1) - RelFN(PosL))
''         SeqLen = SeqLen + Delta
''         If Delta > 1 Then      'we have hole
''            LHole = True
''            HlCnt = HlCnt + 1
''            TtlHlLen = TtlHlLen + Delta
''            HlPct = CDbl(TtlHlLen / SeqLen)
''            'if too many holes, or hole too large, or percentage of hole in sequence
''            'too high, refuse last point and mark left side as done
''            With UMCDef
''              If (HlCnt > .GapMaxCnt) Or (Delta > .GapMaxSize) Or (HlPct > .GapMaxPct) Then
''                 SeqLen = SeqLen - Delta
''                 If blnOptimisticSearching Then
''                    TtlHlLen = TtlHlLen - Delta
''                    HlCnt = HlCnt - 1
''                 End If
''                 PosL = PosL + 1
''                 LDone = True
''              End If
''            End With
''         End If
''      End If
''   Loop
''   LHole = False
''   Do Until (RDone Or RHole)
''      PosR = PosR + 1
''      If PosR > MaxInd Then
''         PosR = MaxInd
''         RDone = True
''      Else
''         Delta = Abs(RelFN(PosR) - RelFN(PosR - 1))
''         SeqLen = SeqLen + Delta
''         If Delta > 1 Then              'we have hole
''            RHole = True
''            HlCnt = HlCnt + 1
''            TtlHlLen = TtlHlLen + Delta
''            HlPct = CDbl(TtlHlLen / SeqLen)
''            With UMCDef
''              If (HlCnt > .GapMaxCnt) Or (Delta > .GapMaxSize) Or (HlPct > .GapMaxPct) Then
''                 SeqLen = SeqLen - Delta
''                 If blnOptimisticSearching Then
''                    TtlHlLen = TtlHlLen - Delta
''                    HlCnt = HlCnt - 1
''                 End If
''                 PosR = PosR - 1
''                 RDone = True
''              End If
''            End With
''         End If
''      End If
''   Loop
''   RHole = False
''Loop
'''now assign class members
''If PosL <= PosR Then
''   'maximum number of elements in class
''   TmpClassCnt = PosR - PosL + 1
''   With GelUMC(Ind).UMCs(ClassInd)
''      ReDim .ClassMInd(TmpClassCnt - 1)
''      ReDim .ClassMType(TmpClassCnt - 1)
''      ReDim TmpMW(TmpClassCnt - 1)
''      AbuSum = 0
''      MWSum = 0
''      For i = PosL To PosR
''          .ClassCount = .ClassCount + 1
''          MemberInd = .ClassCount - 1
''          .ClassMInd(MemberInd) = OID(IndFN(i))
''          .ClassMType(MemberInd) = ODT(IndFN(i))
''          AbuSum = AbuSum + OAbu(IndFN(i))
''          MWSum = MWSum + OMW(IndFN(i))
''          TmpMW(MemberInd) = OMW(IndFN(i))
''          If i <> FNZeroInd Then UMC(IndFN(i)) = -RefInd
''      Next i
''      If .ClassCount > 0 Then
''         ReDim Preserve .ClassMInd(.ClassCount - 1)
''         ReDim Preserve .ClassMType(.ClassCount - 1)
''         ReDim Preserve TmpMW(.ClassCount - 1)
''      Else
''         Erase .ClassMInd
''         Erase .ClassMType
''         Erase TmpMW
''      End If
''      Select Case UMCDef.ClassAbu
''      Case UMCClassAbundanceConstants.UMCAbuAvg
''          If .ClassCount > 0 Then
''             .ClassAbundance = AbuSum / .ClassCount
''          Else
''             .ClassAbundance = ER_CALC_ERR
''          End If
''      Case UMCClassAbundanceConstants.UMCAbuSum
''          If AbuSum > 0 And .ClassCount > 0 Then
''             .ClassAbundance = AbuSum
''          Else
''             .ClassAbundance = ER_CALC_ERR
''          End If
''      Case UMCClassAbundanceConstants.UMCAbuRep
''          'nothing it is already there
''      End Select
''      Dim MyStat As New StatDoubles
''      If MyStat.Fill(TmpMW()) Then
''         .ClassMWStD = MyStat.StDev
''      Else
''         .ClassMWStD = ER_CALC_ERR
''      End If
''      Select Case UMCDef.ClassMW
''      Case UMCClassMassConstants.UMCMassAvg
''           If MWSum > 0 And .ClassCount > 0 Then
''             .ClassMW = MWSum / .ClassCount
''           Else
''             .ClassMW = ER_CALC_ERR
''           End If
''      Case UMCClassMassConstants.UMCMassRep
''           'nothing, it is already there
''      Case UMCClassMassConstants.UMCMassMed                 'have to recalculate
''           If MWSum > 0 And .ClassCount > 0 Then
''             .ClassMW = MyStat.Median
''           Else
''             .ClassMW = ER_CALC_ERR
''           End If
''      End Select
''   End With
''End If
''End Sub
''
' Unused Procedure (February 2005)
''Private Function FindRelFNZero(ByVal RefInd As Long) As Long
'''-----------------------------------------------------------
'''returns index in RelFN of element 0 with RefInd in IndFN
'''-----------------------------------------------------------
''Dim MinInd As Long
''Dim MaxInd As Long
''Dim MidInd As Long
''Dim LInd As Long
''Dim RInd As Long
''Dim Done As Boolean
''Dim LDone As Boolean
''Dim RDone As Boolean
''
''MinInd = LBound(RelFN)
''MaxInd = UBound(RelFN)
''
''Do Until Done
''    MidInd = (MinInd + MaxInd) \ 2
''    If MidInd = MinInd Then             'MinInd and MaxInd next to each other
''       If IndFN(MinInd) = RefInd Then   'one of them have to be our FNZero
''          FindRelFNZero = MinInd
''       Else
''          FindRelFNZero = MaxInd
''       End If
''       Exit Function
''    End If
''    If RelFN(MidInd) < 0 Then
''       MinInd = MidInd
''    ElseIf RelFN(MidInd) > 0 Then
''       MaxInd = MidInd
''    Else            'we got one zero, make sure it is the right one
''       If IndFN(MidInd) = RefInd Then
''          FindRelFNZero = MidInd
''       Else         'if not right on RefInd, check on left and right of it
''          LInd = MidInd
''          Do Until LDone
''             LInd = LInd - 1
''             If LInd < MinInd Then
''                LDone = True
''             Else
''                If RelFN(LInd) <> 0 Then
''                   LDone = True
''                ElseIf IndFN(LInd) = RefInd Then
''                   FindRelFNZero = LInd
''                   Exit Function
''                End If
''             End If
''          Loop
''          RInd = MidInd
''          Do Until RDone
''             RInd = RInd + 1
''             If RInd > MaxInd Then
''                RDone = True
''             Else
''                If RelFN(RInd) <> 0 Then
''                   RDone = True
''                ElseIf IndFN(RInd) = RefInd Then
''                   FindRelFNZero = RInd
''                   Exit Function
''                End If
''             End If
''          Loop
''          'if we came here something is wrong
''          FindRelFNZero = -1
''       End If
''       Done = True
''    End If
''Loop
''End Function

Public Function GetUMCDefDesc(umcDef1 As UMCDefinition) As String
'returns formated definition of the  count
Dim sTmp As String
On Error GoTo exit_GetUMCDefDesc
With umcDef1
    Select Case .DefScope
    Case glScope.glSc_All
      sTmp = "Unique Mass Classes(UMC) on all data points." & vbCrLf
    Case glScope.glSc_Current
      sTmp = "Unique Mass Classes(UMC) on currently visible data." & vbCrLf
    End Select
    Select Case .UMCType
    Case glUMC_TYPE_INTENSITY
      sTmp = sTmp & "UMC type: Intensity" & vbCrLf
    Case glUMC_TYPE_FIT
      sTmp = sTmp & "UMC type: Fit" & vbCrLf
    Case glUMC_TYPE_MINCNT
      sTmp = sTmp & "UMC type: Minimize count" & vbCrLf
    Case glUMC_TYPE_MAXCNT
      sTmp = sTmp & "UMC type: Maximize count" & vbCrLf
    Case glUMC_TYPE_UNQAMT
      sTmp = sTmp & "UMC type: Unique AMT Hits" & vbCrLf
    Case glUMC_TYPE_ISHRINKINGBOX
      sTmp = sTmp & "UMC type: Intensity with shrinking box, "
      If ShrinkingBox_MW_Average_Type = 1 Then
        sTmp = sTmp & "Averaging: Weighted on Intensity" & vbCrLf
      Else
        sTmp = sTmp & "Averaging: Non-Weighted" & vbCrLf
      End If
    Case glUMC_TYPE_FROM_NET
      sTmp = sTmp & "UMC type: UMCIonNet Search" & vbCrLf
    End Select
    Select Case .MWField
    Case 6
      sTmp = sTmp & "Molecular mass: Average" & vbCrLf
    Case 7
      sTmp = sTmp & "Molecular mass: Monoisotopic" & vbCrLf
    Case 8
      sTmp = sTmp & "Molecular mass: Most Abundant" & vbCrLf
    End Select
    Select Case .ClassAbu
    Case UMCClassAbundanceConstants.UMCAbuAvg
      sTmp = sTmp & "Class abundance: Average of member abundances" & vbCrLf
    Case UMCClassAbundanceConstants.UMCAbuSum
      sTmp = sTmp & "Class abundance: Sum of member abundances" & vbCrLf
    Case UMCClassAbundanceConstants.UMCAbuRep
      sTmp = sTmp & "Class abundance: Abundance of class representative" & vbCrLf
    End Select
    Select Case .ClassMW
    Case UMCClassMassConstants.UMCMassAvg
      sTmp = sTmp & "Class molecular mass: Average of member masses" & vbCrLf
    Case UMCClassMassConstants.UMCMassRep
      sTmp = sTmp & "Class molecular mass: Mass of class representative" & vbCrLf
    Case UMCClassMassConstants.UMCMassMed
      sTmp = sTmp & "Class molecular mass: Median of member masses" & vbCrLf
    End Select
    Select Case .TolType
    Case gltPPM
        sTmp = sTmp & "Molecular mass tolerance: " & .Tol & "ppm" & vbCrLf
    Case gltABS
        sTmp = sTmp & "Molecular mass tolerance: " & .Tol & "Da" & vbCrLf
    End Select
    sTmp = sTmp & "Number of allowed gaps: " & .GapMaxCnt & vbCrLf
    sTmp = sTmp & "Allowed size of gap: " & .GapMaxSize & vbCrLf
    sTmp = sTmp & "Allowed percentage of gaps: " & Format$(.GapMaxPct, "Percent") & vbCrLf
End With

exit_GetUMCDefDesc:
GetUMCDefDesc = sTmp
End Function


Public Function UMCScanLocker(ByVal Ind As Long, _
                              ByVal ClassInd As Long, _
                              ByVal FN As Long) As Long
'------------------------------------------------------
'returns class member index(first found) in ISN
'structure if such exists in scan FN; -1 if not
'-----------------------------DON'T KNOW DO I NEED THIS
'looks like I do!(in locking procedures)
'------------------------------------------------------
Dim i As Long
On Error GoTo exit_UMCScanLocker
With GelUMC(Ind).UMCs(ClassInd)
  If .ClassCount > 0 Then
     For i = 1 To .ClassCount
       If .ClassMType(i) = gldtIS Then
          If GelData(Ind).IsoData(.ClassMInd(i)).ScanNumber = FN Then
             UMCScanLocker = .ClassMInd(i)
             Exit Function
          End If
       End If
     Next i
  End If
End With
exit_UMCScanLocker:
UMCScanLocker = -1
End Function

' Unused Function (March 2003)
'''Public Function GetClsInd(ByVal Ind As Long, _
'''                          ByVal ID As Long, _
'''                          ByVal IDType As Integer, _
'''                          ByRef ClsInd() As Long) As Long
''''---------------------------------------------------------
''''returns number of classes in GelUMC(Ind).UMCs to which
''''belongs element in CS/IS array ID; -1 if not found/error
''''class indexes are loaded to ClsInd array
''''---------------------------------------------------------
'''Dim i As Long, j As Long
'''Dim TmpCnt As Long
'''On Error GoTo exit_GetClsInd
'''
'''TmpCnt = 0
'''With GelUMC(Ind)
'''  If .UMCCnt > 0 Then
'''    For i = 0 To .UMCCnt - 1
'''      With .UMCs(i)
'''        If .ClassCount > 0 Then
'''          For j = 0 To .ClassCount - 1
'''            If .ClassMType(j) = IDType And .ClassMInd(j) = ID Then
'''               TmpCnt = TmpCnt + 1
'''               ClsInd(TmpCnt - 1) = i
'''               Exit For    'j For...
'''            End If
'''          Next j
'''        End If
'''      End With
'''    Next i
'''    If TmpCnt > 0 Then ReDim Preserve ClsInd(TmpCnt - 1)
'''  End If
'''End With
'''
'''exit_GetClsInd:
'''If Err Then TmpCnt = -1
'''GetClsInd = TmpCnt
'''End Function
'''
'''
'''Public Function GetClsIndInClsStat(ByVal ClsStat As Variant, _
'''                                   ByVal ClsRepInd As Long) As Long
''''-------------------------------------------------------------------
''''returns class index in ClsStat array for ClsRepInd; -1 on any error
''''----------------------------------------------------DO I NEED THIS?
'''Dim ClsStatCnt As Long
'''Dim i As Long
'''On Error GoTo exit_GetClsIndInClsStat
'''GetClsIndInClsStat = -1
'''If IsArray(ClsStat) Then
'''   ClsStatCnt = UBound(ClsStat)
'''   For i = 1 To ClsStatCnt
'''      If ClsStat(i, 1) = ClsRepInd Then
'''         GetClsIndInClsStat = i
'''         Exit Function
'''      End If
'''   Next i
'''End If
'''exit_GetClsIndInClsStat:
'''End Function

Public Function UMCStatistics1(ByVal Ind As Long, _
                               ByRef Stat() As Double) As Long
'---------------------------------------------------------------
'fills array containing statistics for UMC count; returns number
'of rows if successful, -1 on any error or if no classes found
'statistic array (0-bounded); rows represent unique classes
'last modified: 03/22/2001; nt
'last modified: 05/20/2003; mem
'---------------------------------------------------------------
'column 0 (ustClassIndex)       - class index in .UMCs
'column 1 (ustClassRepMW)       - dblMW of class representative
'column 2 (ustScanStart)        - class first scan number
'column 3 (ustScanEnd)          - class last scan number
'column 4 (ustClassMW)          - class dblMW (based on definition average or class representative)
'column 5 (ustClassMWStDev)     - dblMW standard deviation (from class dblMW)
'        Standard Deviation here is stretched in a sense that we don't
'        have to use average; rather we will use whatever will be used as
'        a class molecular mass (which might be class average or mass of the class representative)
' Monroe Note: The ustClassMWStDev value will be drastically larger than the ustMassStDev value
'              if the class mass type is not Average Mass
'column 6 (ustClassIntensity)   - class intensity (based on definition can be average,
'           sum or intensity of class representative
'column 7 (ustFitAverage)       - average class fit
'column 8 (ustClassRepIntensity)- intensity(abundance) of class representative
'column 9  (ustAbundanceMin)    - minimum abundance
'column 10 (ustAbundanceMax)    - maximum abundance
'column 11 (ustChargeMin)       - minimum charge state
'column 12 (ustChargeMax)       - maximum charge state
'column 13 (ustFitMin)          - minimum isotopic fit
'column 14 (ustFitMax)          - maximum isotopic fit
'column 15 (ustFitStDev)        - standard deviation of isotopic fit
'column 16 (ustMassMin)         - minimum mass
'column 17 (ustMassMax)         - maximum mass
'column 18 (ustMassStDev)       - standard deviation of mass

Dim i As Long
Dim j As Long
Dim ISF As Integer

Dim dblMW As Double, dblFit As Double, dblAbu As Double
Dim intCharge As Integer
Dim tmp1 As Double
Dim tmp2 As Double

Dim MassMin As Double, MassMax As Double, MassSum As Double, MassSumSq As Double, MassStDevSquared As Double
Dim AbuMin As Double, AbuMax As Double
Dim ChargeMin As Integer, ChargeMax As Integer
Dim FitMin As Double, FitMax As Double, FitSum As Double, FitSumSq As Double, FitStDevSquared As Double

On Error GoTo err_UMCStatistics1

With GelUMC(Ind)
  If .UMCCnt > 0 Then
     ReDim Stat(.UMCCnt - 1, UMC_STATISTICS1_MAX_INDEX)
     ISF = .def.MWField
     If ISF <> isfMWavg And ISF <> isfMWMono And ISF <> isfMWTMA Then
        Debug.Assert False
        ISF = isfMWMono
     End If
     
     For i = 0 To .UMCCnt - 1
        With .UMCs(i)
          If .ClassCount > 0 Then
             Stat(i, 0) = i
             'class representative(mass and intensity)
             Select Case .ClassRepType
             Case gldtCS
               Stat(i, ustClassRepMW) = GelData(Ind).CSData(.ClassRepInd).AverageMW
               Stat(i, ustClassRepIntensity) = GelData(Ind).CSData(.ClassRepInd).Abundance
             Case gldtIS
               Stat(i, ustClassRepMW) = GetIsoMass(GelData(Ind).IsoData(.ClassRepInd), ISF)
               Stat(i, ustClassRepIntensity) = GelData(Ind).IsoData(.ClassRepInd).Abundance
             End Select
             'class members are ordered on scan numbers
             'first scan number
             Select Case .ClassMType(0)
             Case gldtCS
               Stat(i, ustScanStart) = GelData(Ind).CSData(.ClassMInd(0)).ScanNumber
               MassMin = GelData(Ind).CSData(.ClassMInd(0)).AverageMW
               AbuMin = GelData(Ind).CSData(.ClassMInd(0)).Abundance
               ChargeMin = GelData(Ind).CSData(.ClassMInd(0)).Charge
               FitMin = GelData(Ind).CSData(.ClassMInd(0)).MassStDev
             Case gldtIS
               Stat(i, ustScanStart) = GelData(Ind).IsoData(.ClassMInd(0)).ScanNumber
               MassMin = GetIsoMass(GelData(Ind).IsoData(.ClassMInd(0)), ISF)
               AbuMin = GelData(Ind).IsoData(.ClassMInd(0)).Abundance
               ChargeMin = GelData(Ind).IsoData(.ClassMInd(0)).Charge
               FitMin = GelData(Ind).IsoData(.ClassMInd(0)).Fit
             End Select
             'last scan number
             Select Case .ClassMType(.ClassCount - 1)
             Case gldtCS
               Stat(i, ustScanEnd) = GelData(Ind).CSData(.ClassMInd(.ClassCount - 1)).ScanNumber
             Case gldtIS
               Stat(i, ustScanEnd) = GelData(Ind).IsoData(.ClassMInd(.ClassCount - 1)).ScanNumber
             End Select
             
             ' Reset the sum and sum of the squares variables
             MassSum = 0
             MassSumSq = 0
             FitSum = 0
             FitSumSq = 0
             
             ' Copy the minimum values to the maximum values
             MassMax = MassMin
             AbuMax = AbuMin
             ChargeMax = ChargeMin
             FitMax = FitMin
             
             For j = 0 To .ClassCount - 1
                If .ClassMType(j) = gldtCS Then
                  dblMW = GelData(Ind).CSData(.ClassMInd(j)).AverageMW
                  dblAbu = GelData(Ind).CSData(.ClassMInd(j)).Abundance
                  intCharge = GelData(Ind).CSData(.ClassMInd(j)).Charge
                  dblFit = GelData(Ind).CSData(.ClassMInd(j)).MassStDev
                Else
                  dblMW = GetIsoMass(GelData(Ind).IsoData(.ClassMInd(j)), ISF)
                  dblAbu = GelData(Ind).IsoData(.ClassMInd(j)).Abundance
                  intCharge = GelData(Ind).IsoData(.ClassMInd(j)).Charge
                  dblFit = GelData(Ind).IsoData(.ClassMInd(j)).Fit
                End If
             
                If dblMW < MassMin Then MassMin = dblMW
                If dblMW > MassMax Then MassMax = dblMW
                MassSum = MassSum + dblMW
                MassSumSq = MassSumSq + dblMW ^ 2
                
                If dblAbu < AbuMin Then AbuMin = dblAbu
                If dblAbu > AbuMax Then AbuMax = dblAbu
                
                If intCharge < ChargeMin Then ChargeMin = intCharge
                If intCharge > ChargeMax Then ChargeMax = intCharge
                
                If dblFit < FitMin Then FitMin = dblFit
                If dblFit > FitMax Then FitMax = dblFit
                FitSum = FitSum + dblFit
                FitSumSq = FitSumSq + dblFit ^ 2
             Next j
             
             ' Minimum and Maximum Mass
             Stat(i, ustMassMin) = MassMin
             Stat(i, ustMassMax) = MassMax
             
             ' Compute pseudo StDev of Mass
             Stat(i, ustClassMW) = .ClassMW       'average or class representative
             tmp1 = MassSumSq / .ClassCount
             tmp2 = .ClassMW ^ 2
             MassStDevSquared = Abs(tmp1 - tmp2)
             Stat(i, ustClassMWStDev) = Sqr(MassStDevSquared)        'dblMW standard deviation
             
             If MassMin = MassMax Or .ClassCount <= 1 Then
                Stat(i, ustMassStDev) = 0
             Else
                ' Compute the rigorous StDev of Mass
                ' StDev:
                MassStDevSquared = MassSumSq / CDbl(.ClassCount - 1) - MassSum ^ 2 / (CDbl(.ClassCount * (.ClassCount - 1)))
                If MassStDevSquared < 0 Then MassStDevSquared = 0
                Stat(i, ustMassStDev) = Sqr(MassStDevSquared)
                
'                ' StDevP:
'                MassStDevSquared = .ClassCount * MassSumSq - MassSum ^ 2
'                If MassStDevSquared < 0 Then
'                   Debug.Assert False
'                   MassStDevSquared = 0
'                End If
'                Stat(i, ustMassStDev) = Sqr(MassStDevSquared) / CDbl(.ClassCount)
             End If
             
             ' Class Intensity
             Stat(i, ustClassIntensity) = .ClassAbundance           'no need to recalculate
             
             ' Minimum and Maximum Abundance
             Stat(i, ustAbundanceMin) = AbuMin
             Stat(i, ustAbundanceMax) = AbuMax
             
             ' Minimum and Maximum Charge
             Stat(i, ustChargeMin) = ChargeMin
             Stat(i, ustChargeMax) = ChargeMax
             
             ' Average, Minimum, and Maximum Fit
             Stat(i, ustFitAverage) = FitSum / .ClassCount
             Stat(i, ustFitMin) = FitMin
             Stat(i, ustFitMax) = FitMax
             
             If .ClassCount > 1 Then
                 ' Compute StDev of Fit
                 FitStDevSquared = FitSumSq / CDbl(.ClassCount - 1) - FitSum ^ 2 / (CDbl(.ClassCount * (.ClassCount - 1)))
                 If FitStDevSquared < 0 Then FitStDevSquared = 0
                 Stat(i, ustFitStDev) = Sqr(FitStDevSquared)
             Else
                 Stat(i, ustFitStDev) = 0
             End If

'             FitStDevSquared = .ClassCount * FitSumSq - FitSum ^ 2
'             If FitStDevSquared < 0 Then
'                Debug.Assert False
'                FitStDevSquared = 0
'             End If
'             Stat(i, ustFitStDev) = Sqr(FitStDevSquared) / CDbl(.ClassCount)
             
          Else     'this should not happen
             For j = 0 To UMC_STATISTICS1_MAX_INDEX
                 Stat(i, j) = -1
             Next j
          End If
        End With
     Next i
     UMCStatistics1 = .UMCCnt
  Else
     UMCStatistics1 = -1
  End If
End With

exit_UMCStatistics1:
Exit Function

err_UMCStatistics1:
Debug.Assert False
UMCStatistics1 = -1
GoTo exit_UMCStatistics1
End Function


Public Function UMCStatistics2(ByVal Ind As Long, _
                               ByRef Stat() As Double) As Long
'---------------------------------------------------------------
'fills array containing statistics for UMC count; returns number
'of rows if successful, -1 on any error or if no classes found
'statistic array (0-bounded); rows represent unique classes
'last modified: 12/14/2001; nt
'---------------------------------------------------------------
'column 0 - class index in .UMCs
'column 1 - class average mw
'column 2 - class first scan number
'column 3 - class last scan number
'column 4 - class total intensity
'column 5 - average class fit
'column 6 - class best fit
'column 7 - index of best fit in class(index within the class)
'column 8 - class count
Dim i As Long, j As Long
Dim ISF As Integer

Dim MW As Double
Dim Abu As Double
Dim Fit As Double
Dim SumMW As Double
Dim SumFit As Double
Dim BestFit As Double
Dim BestFitInd As Long      'index of best fit in class
Dim SumAbu As Double
On Error Resume Next

With GelUMC(Ind)
  If .UMCCnt > 0 Then
     ReDim Stat(.UMCCnt - 1, 8)
     ISF = .def.MWField
     For i = 0 To .UMCCnt - 1
        With .UMCs(i)
          Stat(i, 0) = i
          Stat(i, 8) = .ClassCount
          SumMW = 0
          SumAbu = 0
          SumFit = 0
          BestFit = glHugeOverExp
          BestFitInd = -1
          If .ClassCount > 0 Then
             'class members are ordered on scan numbers
             'first scan number
             Select Case .ClassMType(0)
             Case gldtCS
               Stat(i, 2) = GelData(Ind).CSData(.ClassMInd(0)).ScanNumber
             Case gldtIS
               Stat(i, 2) = GelData(Ind).IsoData(.ClassMInd(0)).ScanNumber
             End Select
             'last scan number
             Select Case .ClassMType(.ClassCount - 1)
             Case gldtCS
               Stat(i, 3) = GelData(Ind).CSData(.ClassMInd(.ClassCount - 1)).ScanNumber
             Case gldtIS
               Stat(i, 3) = GelData(Ind).IsoData(.ClassMInd(.ClassCount - 1)).ScanNumber
             End Select
             For j = 0 To .ClassCount - 1
                If .ClassMType(j) = gldtCS Then
                    MW = GelData(Ind).CSData(.ClassMInd(j)).AverageMW
                    Abu = GelData(Ind).CSData(.ClassMInd(j)).Abundance
                    Fit = GelData(Ind).CSData(.ClassMInd(j)).MassStDev
                Else
                    MW = GetIsoMass(GelData(Ind).IsoData(.ClassMInd(j)), ISF)
                    Abu = GelData(Ind).IsoData(.ClassMInd(j)).Abundance
                    Fit = GelData(Ind).IsoData(.ClassMInd(j)).Fit
                End If
                SumMW = SumMW + MW
                SumAbu = SumAbu + Abu
                SumFit = SumFit + Fit
                If Fit < BestFit Then
                   BestFit = Fit
                   BestFitInd = j       'index within the class
                End If
             Next j
             Stat(i, 1) = SumMW / .ClassCount
             Stat(i, 4) = SumAbu
             Stat(i, 5) = SumFit / .ClassCount
             Stat(i, 6) = BestFit
             Stat(i, 7) = BestFitInd
          Else     'this should not happen
             For j = 0 To 8
                 Stat(i, j) = -1
             Next j
          End If
        End With
     Next i
     UMCStatistics2 = .UMCCnt
  Else
     UMCStatistics2 = -1
  End If
End With

exit_UMCStatistics2:
Exit Function

err_UMCStatistics2:
UMCStatistics2 = -1
GoTo exit_UMCStatistics2
End Function


Public Function GetCSScope(ByVal Ind As Long, _
                           ByRef CSInd() As Long, _
                           ByVal Scope As glScope) As Long
'---------------------------------------------------------
'fills CS data from GelData(Ind) currently in scope and
'returns its number
'---------------------------------------------------------
Dim i As Long
Dim Cnt As Long
With GelDraw(Ind)
  If .CSCount > 0 Then
     ReDim CSInd(.CSCount)
     Select Case Scope
     Case glScope.glSc_Current
       If GelBody(Ind).fgDisplay = glvDifferential Then
          For i = 1 To .CSCount
             If ((.CSID(i) > 0) And (.CSR(i) > 0) And (.CSER(i) >= 0)) Then
                Cnt = Cnt + 1
                CSInd(Cnt) = i
             End If
          Next i
       Else
          For i = 1 To .CSCount
             If ((.CSID(i) > 0) And (.CSR(i) > 0)) Then
                Cnt = Cnt + 1
                CSInd(Cnt) = i
             End If
          Next i
       End If
       If Cnt > 0 And Cnt < .CSCount Then ReDim Preserve CSInd(Cnt)
     Case glScope.glSc_All
       For i = 1 To .CSCount
          CSInd(i) = i
       Next i
       Cnt = .CSCount
     End Select
  Else
     Cnt = 0
  End If
End With
If Cnt <= 0 Then Erase CSInd
GetCSScope = Cnt
End Function


Public Function GetISScope(ByVal Ind As Long, _
                           ByRef ISInd() As Long, _
                           ByVal Scope As glScope) As Long
'----------------------------------------------------------
'fills Isotopic data from GelData(Ind) currently in scope
'Returns number of points in ISInd(), which is a 1-based array
'----------------------------------------------------------
Dim i As Long
Dim Cnt As Long
With GelDraw(Ind)
  If .IsoCount > 0 Then
     ReDim ISInd(.IsoCount)
     Select Case Scope
     Case glScope.glSc_Current
       If GelBody(Ind).fgDisplay = glvDifferential Then
          For i = 1 To .IsoCount
            If ((.IsoID(i) > 0) And (.IsoR(i) > 0) And (.IsoER(i) >= 0)) Then
               Cnt = Cnt + 1
               ISInd(Cnt) = i
            End If
          Next i
       Else
          For i = 1 To .IsoCount
            If ((.IsoID(i) > 0) And (.IsoR(i) > 0)) Then
               Cnt = Cnt + 1
               ISInd(Cnt) = i
            End If
          Next i
       End If
       If Cnt > 0 And Cnt < .IsoCount Then ReDim Preserve ISInd(Cnt)
     Case glScope.glSc_All
       For i = 1 To .IsoCount
          ISInd(i) = i
       Next i
       Cnt = .IsoCount
     End Select
  Else
     Cnt = 0
  End If
End With
If Cnt <= 0 Then Erase ISInd
GetISScope = Cnt
End Function

Public Function GetISScopeFilterByUMC(ByVal lngGelIndex As Long, _
                                      ByRef ISInd() As Long, _
                                      ByVal Scope As glScope, _
                                      ByVal UMCInd As Long) As Long
'----------------------------------------------------------
'fills Isotopic data from GelData(lngGelIndex) currently in scope
'and belonging to UMC given by UMCInd
'Returns number of points in ISInd(), which is a 1-based array
'----------------------------------------------------------
Dim i As Long
Dim Cnt As Long

Dim lngUMCMemberCount As Long
Dim lngUMCGelDataIndices() As Long

With GelUMC(lngGelIndex).UMCs(UMCInd)
    
    lngUMCMemberCount = .ClassCount
    ' Copy the .ClassMInd() array into the lngUMCGelDataIndices() array
    lngUMCGelDataIndices = .ClassMInd
    
    ShellSortLong lngUMCGelDataIndices, 0, lngUMCMemberCount - 1
End With

With GelDraw(lngGelIndex)
    If .IsoCount > 0 And lngUMCMemberCount > 0 Then
        ReDim ISInd(.IsoCount)
        Select Case Scope
        Case glScope.glSc_Current
            If GelBody(lngGelIndex).fgDisplay = glvDifferential Then
                For i = 1 To .IsoCount
                    If ((.IsoID(i) > 0) And (.IsoR(i) > 0) And (.IsoER(i) >= 0)) Then
                        If BinarySearchLng(lngUMCGelDataIndices, .IsoID(i)) >= 0 Then
                            Cnt = Cnt + 1
                            ISInd(Cnt) = i
                        End If
                    End If
                Next i
            Else
                For i = 1 To .IsoCount
                    If ((.IsoID(i) > 0) And (.IsoR(i) > 0)) Then
                        If BinarySearchLng(lngUMCGelDataIndices, .IsoID(i)) >= 0 Then
                            Cnt = Cnt + 1
                            ISInd(Cnt) = i
                        End If
                    End If
                Next i
            End If
            If Cnt > 0 And Cnt < .IsoCount Then ReDim Preserve ISInd(Cnt)
        Case glScope.glSc_All
            For i = 1 To .IsoCount
                If BinarySearchLng(lngUMCGelDataIndices, .IsoID(i)) >= 0 Then
                    ISInd(i) = i
                    Cnt = Cnt + 1
                End If
            Next i
         End Select
    Else
        Cnt = 0
    End If
End With

If Cnt <= 0 Then Erase ISInd
GetISScopeFilterByUMC = Cnt
End Function

' Unused Procedure (February 2005)
''Private Function UMCFitIntensityWithShrinkingBox(ByVal Ind As Long, _
''                                        ByVal TtlCnt As Long, _
''                                        ByRef frmCallingForm As VB.Form) As Boolean
'''-----------------------------------------------------------------------
'''do the real Unique Mass Classes count; returns True if successful
'''NOTE: this function is called only for TtlCnt>1
'''actual classes are filled in HolePatternIShrinkingBox function
'''-----------------------------------------------------------------------
''Dim i As Long               'counter through sorted IndOR/ORFld arrays
''Dim RefInd As Long          'reference index in original arrays
''
''Dim MinInd As Long          'indexes in IndMW/MW arrays of range of candidates
''Dim MaxInd As Long          'to be duplicate of class representative
''Dim FNZeroInd As Long       'index in IndFN/RelFN of current representative
''
''Dim Done As Boolean
''Dim Found1stUnmarked As Boolean
''Dim TmpCnt As Long
''
''Dim qslSort As New QSLong
''Dim j As Long
''Dim mwutUC As New MWUtil  'range finder
''
''On Error GoTo err_UMCFitIntensityCount
''
'''Input choice for averaging type for UMC base value
''ShrunkBoxes = 0
''If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''    If glbPreferencesExpanded.AutoAnalysisOptions.UMCShrinkingBoxWeightAverageMassByIntensity Then
''        ShrinkingBox_MW_Average_Type = 1
''    Else
''        ShrinkingBox_MW_Average_Type = 0
''    End If
''Else
''    Do
''        ShrinkingBox_MW_Average_Type = CInt(InputBox("0 = Average by points only" & Chr(13) & "1 = Weighted average by intensity", "Input Averaging Type", "0"))
''    Loop While Not (ShrinkingBox_MW_Average_Type = 0 Or ShrinkingBox_MW_Average_Type = 1)
''End If
''
''i = 0
''TmpCnt = 0
'''MW is sorted ascending
''If Not mwutUC.Fill(MW()) Then GoTo err_UMCFitIntensityCount
''Do Until Done
''    Found1stUnmarked = False
''    'find first element of UMC that is 0 (unmarked)
''    Do Until Found1stUnmarked
''       i = i + 1
''       If i > TtlCnt Then
''          Done = True
''          Exit Do
''       Else
''          RefInd = IndOR(i)
''          If UMC(RefInd) = 0 Then Found1stUnmarked = True
''       End If
''    Loop
''    If Found1stUnmarked Then    'we have next unmarked
''       TmpCnt = TmpCnt + 1
''       'UMC(RefInd) = RefInd     'so it is class representative
''       MWZero = OMW(RefInd)
''       FNZero = OFN(RefInd)
''       Select Case UMCDef.TolType
''       Case gltPPM
''            'AbsTol = umcDef.Tol * MWZero * glPPM
''            AbsTol = (MWZero / (1 - (UMCDef.Tol * glPPM))) * 2 * UMCDef.Tol * glPPM
''       Case gltABS
''            AbsTol = 2 * UMCDef.Tol
''       End Select
''       MinInd = 1
''       MaxInd = TtlCnt
''       If mwutUC.FindIndexRange(MWZero, AbsTol, MinInd, MaxInd) Then
''          ReDim IndFN(MinInd To MaxInd)
''          ReDim RelFN(MinInd To MaxInd)
''          ReDim SBValidPoints(MinInd To MaxInd)
''          For j = MinInd To MaxInd
''              IndFN(j) = IndMW(j)
''              RelFN(j) = OFN(IndMW(j)) - FNZero 'distance from the FNZero
''              'SBValidPoints(j) = UMC(IndMW(j))
''          Next j
''          'sort it based on distance from the FNZero;
''          If Not qslSort.QSAsc(RelFN(), IndFN()) Then
''             GoTo err_UMCFitIntensityCount
''          End If
''          For j = MinInd To MaxInd
''              SBValidPoints(j) = UMC(IndFN(j))
''          Next j
''          FNZeroInd = FindRelFNZero(RefInd)
''          HolePatternsIShrinkingBox Ind, MinInd, MaxInd, FNZeroInd, RefInd, TtlCnt, frmCallingForm
''       End If
''    End If
''    If glAbortUMCProcessing Then GoTo err_UMCFitIntensityCount
''Loop
''UMCFitIntensityWithShrinkingBox = True
''
''exit_UMCFitIntensityCount:
''Set qslSort = Nothing
''Set mwutUC = Nothing
''Exit Function
''
''err_UMCFitIntensityCount:
''UMCFitIntensityWithShrinkingBox = False
''MsgBox "Processing aborted (or an error occurred).", vbExclamation + vbOKOnly, glFGTU
''LogErrors Err.Number, "UMCFitIntensityCount"
''GoTo exit_UMCFitIntensityCount
''End Function
''
' Unused Procedure (February 2005)
''Private Function HolePatternsIShrinkingBox(ByVal Ind As Long, _
''                         ByVal MinInd As Long, _
''                         ByVal MaxInd As Long, _
''                         ByVal FNZeroInd As Long, _
''                         ByVal RefInd As Long, _
''                         ByVal TtlCnt As Long, _
''                         ByRef frmCallingForm As VB.Form) As Boolean
'''-------------------------------------------------------------------
'''Fills class structure; Ind is Gel index
'''RelFN array contains potential duplicates of FNZero elements; here
'''some of them will be eliminated based on their relative positions
'''RelFN array is sorted ascending; MinInd and MaxInd are bounds of it
'''and FNZeroInd is index of UMC representative
'''-------------------------------------------------------------------
''
''Dim HlCnt As Long       'hole counter
''Dim TtlHlLen As Long    'total hole length
''Dim HlPct As Double     'percentage of the hole in the sequence
''Dim SeqLen As Long      'sequence length
''
''Dim Delta As Long
'''because nothing guarantees symmetry have to work left & right separately
''Dim PosL As Long, PosR As Long
''Dim StartPosL As Long, StartPosR As Long
''Dim LDone As Boolean, RDone As Boolean
''Dim LHole As Boolean, RHole As Boolean
''Dim TmpClassCnt As Long
''Dim ClassInd As Long
''Dim MemberInd As Long
''Dim AbuSum As Double
''Dim MWSum As Double
''Dim i As Long
''Dim HoleStartInd() As Long
''Dim HoleEndInd() As Long
''Dim HoleDelta() As Long
''
''ReDim HoleStartInd(0 To 0)
''ReDim HoleEndInd(0 To 0)
''ReDim HoleDelta(0 To 0)
''For i = FNZeroInd - 1 To MinInd Step -1
''    If RelFN(i) = RelFN(i + 1) - 1 And SBValidPoints(i) = 0 Then
''        HoleEndInd(0) = i
''    Else
''        Exit For
''    End If
''Next i
''If HoleEndInd(0) = 0 Then
''    HoleEndInd(0) = FNZeroInd
''End If
''For i = FNZeroInd + 1 To MaxInd Step 1
''    If RelFN(i) = RelFN(i - 1) + 1 And SBValidPoints(i) = 0 Then
''        HoleStartInd(0) = i
''    Else
''        Exit For
''    End If
''Next i
''If HoleStartInd(0) = 0 Then
''    HoleStartInd(0) = FNZeroInd
''End If
''HolePatternsIShrinkingBox = False
''
''HlCnt = 0
''SeqLen = 1  'at least FNZero is in sequence
'''With GelUMC(Ind)
'''    .UMCCnt = .UMCCnt + 1                               'new class
'''    ClassInd = .UMCCnt - 1
'''    If .UMCCnt > UBound(.UMCs) Then                     'add room for new classes
'''       ReDim Preserve .UMCs(.UMCCnt + 5000)
'''    End If
'''    .UMCs(ClassInd).ClassRepInd = OID(RefInd)
'''    .UMCs(ClassInd).ClassRepType = ODT(RefInd)
'''    .UMCs(ClassInd).ClassCount = 0                      'class members count
'''    'put by default class abundance at class representative abundance
'''    .UMCs(ClassInd).ClassAbundance = OAbu(RefInd)
'''    'and class molecular mass at class representative molecular mass
'''    .UMCs(ClassInd).ClassMW = OMW(RefInd)
'''End With
''TtlHlLen = 0
''PosL = FNZeroInd
''PosR = FNZeroInd
''If FNZeroInd <= MinInd Then
''   LDone = True
''Else
''   LDone = False
''End If
''If FNZeroInd >= MaxInd Then
''   RDone = True
''Else
''   RDone = False
''End If
''LHole = False
''RHole = False
''Do Until (LDone And RDone)      'work until done on both sides
''   Do Until (LDone Or LHole)    'if hole on left side jump to the right
''      StartPosL = PosL
''      PosL = PosL - 1
''      If PosL < MinInd Then     'we are done on the left
''         PosL = MinInd
''         LDone = True
''      Else
''         Do While SBValidPoints(PosL) <> 0
''            PosL = PosL - 1
''            If PosL < MinInd Then
''                PosL = StartPosL
''                LDone = True
''                Exit Do
''            End If
''         Loop
''         Delta = Abs(RelFN(StartPosL) - RelFN(PosL))
''         SeqLen = SeqLen + Delta
''         If Delta > 1 Then      'we have hole
''            'LHole = True
''
''            'TtlHlLen = TtlHlLen + Delta
''            'HlPct = CDbl(TtlHlLen / SeqLen)
''            'if too many holes, or hole too large, or percentage of hole in sequence
''            'too high, refuse last point and mark left side as done
''            With UMCDef
''              'If (HlCnt > .GapMaxCnt) Or (Delta > .GapMaxSize) Or (HlPct > .GapMaxPct) Then
''              If (Delta > .GapMaxSize) Then
''                 SeqLen = SeqLen - Delta
''                 PosL = StartPosL
''                 LDone = True
''              Else
''                 HlCnt = HlCnt + 1
''                 ReDim Preserve HoleStartInd(0 To HlCnt)
''                 ReDim Preserve HoleEndInd(0 To HlCnt)
''                 ReDim Preserve HoleDelta(0 To HlCnt)
''                 HoleStartInd(HlCnt) = PosL
''                 HoleEndInd(HlCnt) = StartPosL
''                 HoleDelta(HlCnt) = Delta
''              End If
''            End With
''         Else 'we do not have hole
''            StartPosL = PosL
''         End If
''      End If
''   Loop
''   LHole = False
''   Do Until (RDone Or RHole)
''      StartPosR = PosR
''      PosR = PosR + 1
''      If PosR > MaxInd Then
''         PosR = MaxInd
''         RDone = True
''      Else
''         Do While SBValidPoints(PosR) <> 0
''            PosR = PosR + 1
''            If PosR > MaxInd Then
''                PosR = StartPosR
''                RDone = True
''                Exit Do
''            End If
''         Loop
''         Delta = Abs(RelFN(PosR) - RelFN(StartPosR))
''         SeqLen = SeqLen + Delta
''         If Delta > 1 Then              'we have hole
''            'RHole = True
''            'HlCnt = HlCnt + 1
''            'TtlHlLen = TtlHlLen + Delta
''            'HlPct = CDbl(TtlHlLen / SeqLen)
''            With UMCDef
''              'If (HlCnt > .GapMaxCnt) Or (Delta > .GapMaxSize) Or (HlPct > .GapMaxPct) Then
''              If (Delta > .GapMaxSize) Then
''                 SeqLen = SeqLen - Delta
''                 PosR = StartPosR
''                 RDone = True
''              Else
''                 HlCnt = HlCnt + 1
''                 ReDim Preserve HoleStartInd(0 To HlCnt)
''                 ReDim Preserve HoleEndInd(0 To HlCnt)
''                 ReDim Preserve HoleDelta(0 To HlCnt)
''                 HoleStartInd(HlCnt) = StartPosR
''                 HoleEndInd(HlCnt) = PosR
''                 HoleDelta(HlCnt) = Delta
''              End If
''            End With
''         Else 'we do not have hole
''            StartPosR = PosR
''         End If
''      End If
''   Loop
''   RHole = False
''Loop
'''--------------- check for percentage and number -------------------
''    Dim SumLeftHoles As Long
''    Dim SumRightHoles As Long
''    Dim PotentialStartInd As Long
''    Dim PotentialEndInd As Long
''    Dim BestStartInd As Long
''    Dim BestEndInd As Long
''    Dim BestNumIncludedPoints As Long
''    Dim MiddleIndexofHoles As Long
''    Dim j As Long
''    Dim TempHlCnt As Long
''    Dim TempHlPct As Double
''
''
''
''
''TtlHlLen = 0
''If HlCnt >= 1 Then
''    For i = 1 To HlCnt
''        TtlHlLen = HoleDelta(i) + TtlHlLen
''    Next i
''End If
''BestStartInd = PosL
''BestEndInd = PosR
''HlPct = CDbl(TtlHlLen / SeqLen)
''If HlPct > UMCDef.GapMaxPct Or HlCnt > UMCDef.GapMaxCnt Then
''
''    If HlCnt >= 1 Then
''        For i = 1 To HlCnt
''            If RelFN(HoleEndInd(i)) <= 0 And RelFN(HoleStartInd(i)) < 0 Then
''                MiddleIndexofHoles = i
''            End If
''        Next i
''    End If
''
''    BestNumIncludedPoints = 0
''    SumLeftHoles = 0
''    For i = 0 To MiddleIndexofHoles
''        If i = 0 Then
''            PotentialStartInd = HoleEndInd(0)
''        Else
''            SumLeftHoles = SumLeftHoles + HoleDelta(i)
''            If i <> MiddleIndexofHoles Then
''                PotentialStartInd = HoleEndInd(i + 1)
''            Else
''                PotentialStartInd = PosL
''            End If
''        End If
''        SumRightHoles = 0
''        For j = MiddleIndexofHoles To HlCnt
''            If j = MiddleIndexofHoles Then
''                PotentialEndInd = HoleStartInd(0)
''            Else
''                SumRightHoles = SumRightHoles + HoleDelta(j)
''                If j <> HlCnt Then
''                    PotentialEndInd = HoleStartInd(j + 1)
''                Else
''                    PotentialEndInd = PosR
''                End If
''            End If
''            TempHlPct = (SumRightHoles + SumLeftHoles) / ((RelFN(PotentialEndInd) - RelFN(PotentialStartInd)) + 1)
''            TempHlCnt = i + (j - MiddleIndexofHoles)
''            If TempHlPct <= UMCDef.GapMaxPct And TempHlCnt <= UMCDef.GapMaxCnt Then
''                If RelFN(PotentialEndInd) - RelFN(PotentialStartInd) - SumRightHoles - SumLeftHoles + TempHlCnt > BestNumIncludedPoints Then
''                    BestNumIncludedPoints = RelFN(PotentialEndInd) - RelFN(PotentialStartInd) - SumRightHoles - SumLeftHoles + TempHlCnt
''                    BestStartInd = PotentialStartInd
''                    BestEndInd = PotentialEndInd
''                End If
''            End If
''         Next j
''    Next i
''End If
''
'''--------------- check for mass consistency ------------------------
''Dim AverageMass As Double
''Dim NumberOfMasses As Long
''Dim WeightsSum As Double
''Dim TotalMass As Double
''
''TotalMass = 0
''WeightsSum = 0
''NumberOfMasses = 0
''For i = BestStartInd To BestEndInd
''    If SBValidPoints(i) = 0 Then
''        WeightsSum = WeightsSum + OAbu(IndFN(i))
''        NumberOfMasses = NumberOfMasses + 1
''        If ShrinkingBox_MW_Average_Type = 1 Then
''            TotalMass = TotalMass + OMW(IndFN(i)) * OAbu(IndFN(i))
''        ElseIf ShrinkingBox_MW_Average_Type = 0 Then
''            TotalMass = TotalMass + OMW(IndFN(i))
''        End If
''    End If
''Next i
''    If ShrinkingBox_MW_Average_Type = 1 Then
''        AverageMass = TotalMass / WeightsSum
''    ElseIf ShrinkingBox_MW_Average_Type = 0 Then
''        AverageMass = TotalMass / NumberOfMasses
''    End If
''
''Dim BestEliminateIndex As Long
''Dim BestEliminateDifference As Double
''Dim MWDiff As Double
''Dim EliminationNeeded As Boolean
''
''BestEliminateDifference = 0
''BestEliminateIndex = 0
''EliminationNeeded = False
''
''For i = BestStartInd To BestEndInd
''    If SBValidPoints(i) = 0 Then
''        If UMCDef.TolType = gltPPM Then
''            MWDiff = Abs((OMW(IndFN(i)) - AverageMass) / AverageMass)
''            If MWDiff >= UMCDef.Tol * glPPM Then
''                If MWDiff > BestEliminateDifference Then
''                    EliminationNeeded = True
''                    If i <> FNZeroInd Then
''                        BestEliminateDifference = MWDiff
''                        BestEliminateIndex = i
''                    End If
''                End If
''            End If
''        ElseIf UMCDef.TolType = gltABS Then
''            MWDiff = Abs(OMW(IndFN(i)) - AverageMass)
''            If MWDiff >= UMCDef.Tol Then
''                If MWDiff > BestEliminateDifference Then
''                    EliminationNeeded = True
''                    If i <> FNZeroInd Then
''                        BestEliminateDifference = MWDiff
''                        BestEliminateIndex = i
''                    End If
''                End If
''            End If
''        End If
''    End If
''Next i
''
'''If the only point that is out of mass range is fnzeroind, then take the
'''point on the start of the potential UMC out, if this can't be done, take the
'''one on the end of the potential UMC out.  Eventually, all points except the
'''index will be eliminated if they can not be eliminated on the basis of mass error.
''If EliminationNeeded = True And BestEliminateIndex = 0 Then
''    If BestStartInd < FNZeroInd Then
''        BestEliminateIndex = BestStartInd
''    ElseIf BestEndInd > FNZeroInd Then
''        BestEliminateIndex = BestEndInd
''    End If
''End If
''
''Dim HPISB As Boolean
''If BestEliminateIndex <> 0 Then
''    ShrunkBoxes = ShrunkBoxes + 1
''    SBValidPoints(BestEliminateIndex) = -1
''    HPISB = HolePatternsIShrinkingBox(Ind, BestStartInd, BestEndInd, FNZeroInd, RefInd, TtlCnt, frmCallingForm)
''    If HPISB = True Then
''        HolePatternsIShrinkingBox = True
''        Exit Function
''    Else
''        MsgBox ("Problem with recursion in function HolePatternIShrinkingBox")
''    End If
''End If
''
'''now assign class members
''If PosL <= PosR Then
''   'maximum number of elements in class
''   TmpClassCnt = PosR - PosL + 1
''   With GelUMC(Ind)
''        .UMCCnt = .UMCCnt + 1                               'new class
''        ClassInd = .UMCCnt - 1
''        If .UMCCnt > UBound(.UMCs) Then                     'add room for new classes
''           ReDim Preserve .UMCs(.UMCCnt + 5000)
''        End If
''        .UMCs(ClassInd).ClassRepInd = OID(RefInd)
''        .UMCs(ClassInd).ClassRepType = ODT(RefInd)
''        .UMCs(ClassInd).ClassCount = 0                      'class members count
''        'put by default class abundance at class representative abundance
''        .UMCs(ClassInd).ClassAbundance = OAbu(RefInd)
''        'and class molecular mass at class representative molecular mass
''        .UMCs(ClassInd).ClassMW = OMW(RefInd)
''    End With
''
''   With GelUMC(Ind).UMCs(ClassInd)
''      ReDim .ClassMInd(TmpClassCnt - 1)
''      ReDim .ClassMType(TmpClassCnt - 1)
''      AbuSum = 0
''      MWSum = 0
''      For i = BestStartInd To BestEndInd            'corrected from  For i=PosL to PosR (Kyle)
''          If SBValidPoints(i) = 0 Then
''            .ClassCount = .ClassCount + 1
''            MemberInd = .ClassCount - 1
''            .ClassMInd(MemberInd) = OID(IndFN(i))
''            .ClassMType(MemberInd) = ODT(IndFN(i))
''            AbuSum = AbuSum + OAbu(IndFN(i))
''            MWSum = MWSum + OMW(IndFN(i))
''            If i <> FNZeroInd Then
''                UMC(IndFN(i)) = -RefInd
''            Else
''                UMC(IndFN(i)) = RefInd
''            End If
''          End If
''      Next i
''      If .ClassCount > 0 Then
''         ReDim Preserve .ClassMInd(.ClassCount - 1)
''         ReDim Preserve .ClassMType(.ClassCount - 1)
''      Else
''         Erase .ClassMInd
''         Erase .ClassMType
''      End If
''      Select Case UMCDef.ClassAbu
''      Case UMCClassAbundanceConstants.UMCAbuAvg
''          If .ClassCount > 0 Then
''             .ClassAbundance = AbuSum / .ClassCount
''          Else
''             .ClassAbundance = ER_CALC_ERR
''          End If
''      Case UMCClassAbundanceConstants.UMCAbuSum
''          If AbuSum > 0 And .ClassCount > 0 Then
''             .ClassAbundance = AbuSum
''          Else
''             .ClassAbundance = ER_CALC_ERR
''          End If
''      Case UMCClassAbundanceConstants.UMCAbuRep
''          'nothing it is already there
''      End Select
''
''
''      Select Case UMCDef.ClassMW
''      Case UMCClassMassConstants.UMCMassAvg
''          If MWSum > 0 And .ClassCount > 0 Then
''             .ClassMW = MWSum / .ClassCount
''          Else
''             .ClassMW = ER_CALC_ERR
''          End If
''      Case UMCClassMassConstants.UMCMassRep
''          'nothing, it is already there
''      End Select
''   End With
''End If
''UMCsDone = GelUMC(Ind).UMCCnt
''If UMCsDone Mod 25 = 0 Then frmCallingForm.Status "LC-MS Features Done: " & Trim(UMCsDone) & " / " & Trim(TtlCnt) & " with " & Trim(ShrunkBoxes) & " shrunk boxes"
''HolePatternsIShrinkingBox = True
''End Function

Public Function UMCAverageMass(ByVal Ind As Long) As Boolean
'--------------------------------------------------------------
'sets masses of all members of a class to a class average
'returns True if successful
'--------------------------------------------------------------
Dim i As Long, j As Long
Dim DInd As Long
Dim UMCStat2() As Double
On Error GoTo err_UMCAverageMass

With GelUMC(Ind)
    If UMCStatistics2(Ind, UMCStat2()) <> .UMCCnt Then Exit Function
    For i = 0 To .UMCCnt - 1
        For j = 0 To .UMCs(i).ClassCount - 1
            DInd = .UMCs(i).ClassMInd(j)
            Select Case .UMCs(i).ClassMType(j)
            Case glCSType
                GelData(Ind).CSData(DInd).AverageMW = UMCStat2(i, 1)
            Case glIsoType
                SetIsoMass GelData(Ind).IsoData(DInd), .def.MWField, UMCStat2(i, 1)
            End Select
        Next j
    Next i
End With
UMCAverageMass = True
Exit Function

err_UMCAverageMass:
End Function

' Unused function (August 2003)
'Public Function fUMCScanRange(ByVal IndDis As Long, ByVal IndUMC As Long) As Long
''--------------------------------------------------------------------------------
''returns scan range for unique mass class IndUMC of display IndDis; -1 on error
''--------------------------------------------------------------------------------
'Dim FirstScan As Long, LastScan As Long, CurrScan As Long
'Dim i As Long
'On Error Resume Next
'FirstScan = 100000:         LastScan = -100000
'With GelUMC(IndDis).UMCs(IndUMC)
'    For i = 0 To .ClassCount - 1
'        Select Case .ClassMType(i)
'        Case glCSType
'             CurrScan = GelData(IndDis).CSData(.ClassMInd(i)).ScanNumber
'        Case glIsoType
'             CurrScan = GelData(IndDis).IsoData(.ClassMInd(i)).ScanNumber
'        End Select
'        If CurrScan < FirstScan Then FirstScan = CurrScan
'        If CurrScan > LastScan Then LastScan = CurrScan
'    Next i
'End With
'fUMCScanRange = LastScan - FirstScan + 1
'Exit Function
'
'err_fUMCScanRange:
'fUMCScanRange = -1
'End Function
'
' Unused function (February 2005)
'Public Function fUMCHiAbuInd(ByVal IndDis As Long, ByVal IndUMC As Long) As Long
''-------------------------------------------------------------------------------
''returns index in UMC of class member with highest abundance
''-------------------------------------------------------------------------------
'Dim i As Long
'Dim HiAbu As Double
'Dim HiAbuInd As Long
'Dim CurrAbu As Double
'On Error Resume Next
'HiAbuInd = -1
'With GelUMC(IndDis).UMCs(IndUMC)
'    For i = 0 To .ClassCount - 1
'        Select Case .ClassMType(i)
'        Case glCSType
'             CurrAbu = GelData(IndDis).CSData(.ClassMInd(i)).Abundance
'        Case glIsoType
'             CurrAbu = GelData(IndDis).IsoData(.ClassMInd(i)).Abundance
'        End Select
'        If CurrAbu > HiAbu Then
'           HiAbu = CurrAbu
'           HiAbuInd = i
'        End If
'    Next i
'End With
'fUMCHiAbuInd = HiAbuInd
'End Function
'
' Unused function (February 2005)
'Public Function fUMCHiAbu(ByVal IndDis As Long, ByVal IndUMC As Long) As Double
''------------------------------------------------------------------------------
''returns highest abundance in the class
''------------------------------------------------------------------------------
'Dim i As Long
'Dim HiAbu As Double
'Dim CurrAbu As Double
'On Error Resume Next
'HiAbu = -1
'With GelUMC(IndDis).UMCs(IndUMC)
'    For i = 0 To .ClassCount - 1
'        Select Case .ClassMType(i)
'        Case glCSType
'             CurrAbu = GelData(IndDis).CSData(.ClassMInd(i)).Abundance
'        Case glIsoType
'             CurrAbu = GelData(IndDis).IsoData(.ClassMInd(i)).Abundance
'        End Select
'        If CurrAbu > HiAbu Then HiAbu = CurrAbu
'    Next i
'End With
'fUMCHiAbu = HiAbu
'End Function


Public Function fUMCSpotsOnly(ByVal IndDis As Long) As Boolean
'------------------------------------------------------------------
'makes only spots belonging to the LC-MS Features visible; returns True if OK
'------------------------------------------------------------------
Dim i As Long, j As Long
On Error GoTo exit_fUMCSpotsOnly
'make all spots invisible
Call GelCSExcludeAll(IndDis)
Call GelIsoExcludeAll(IndDis)
'now go and make all UMC spots visible
With GelUMC(IndDis)
  For i = 0 To .UMCCnt - 1
      With .UMCs(i)
        For j = 0 To .ClassCount - 1
            Select Case .ClassMType(j)
            Case glCSType
              GelDraw(IndDis).CSID(.ClassMInd(j)) = Abs(GelDraw(IndDis).CSID(.ClassMInd(j)))
            Case glIsoType
              GelDraw(IndDis).IsoID(.ClassMInd(j)) = Abs(GelDraw(IndDis).IsoID(.ClassMInd(j)))
            End Select
        Next j
      End With
  Next i
End With

If False Then
    For i = 1 To GelDraw(IndDis).IsoCount
        GelDraw(IndDis).IsoID(i) = -(GelDraw(IndDis).IsoID(i))
    Next i
End If

fUMCSpotsOnly = True
exit_fUMCSpotsOnly:
End Function

Public Function ShowNetAdjUMCPoints(ByVal lngGelIndex As Long, ByVal lngUMCIndicatorBit As Long) As Boolean
'------------------------------------------------------------------
'makes only spots belonging to LC-MS Features that were used for Net Adjustment visible; returns True if OK
'------------------------------------------------------------------
Dim i As Long, j As Long
On Error GoTo exit_ShowNetAdjUMCPoints
'make all spots invisible
Call GelCSExcludeAll(lngGelIndex)
Call GelIsoExcludeAll(lngGelIndex)
'now go and make all UMC spots visible
With GelUMC(lngGelIndex)
  For i = 0 To .UMCCnt - 1
      With .UMCs(i)
        If (.ClassStatusBits And lngUMCIndicatorBit) = lngUMCIndicatorBit Then
            For j = 0 To .ClassCount - 1
                Select Case .ClassMType(j)
                Case glCSType
                  GelDraw(lngGelIndex).CSID(.ClassMInd(j)) = Abs(GelDraw(lngGelIndex).CSID(.ClassMInd(j)))
                Case glIsoType
                  GelDraw(lngGelIndex).IsoID(.ClassMInd(j)) = Abs(GelDraw(lngGelIndex).IsoID(.ClassMInd(j)))
                End Select
            Next j
        End If
      End With
  Next i
End With
ShowNetAdjUMCPoints = True
exit_ShowNetAdjUMCPoints:
End Function

Public Function ShowSplitUMCPoints(ByVal lngGelIndex As Long) As Boolean
'------------------------------------------------------------------
'makes only spots belonging to LC-MS Features that have been split visible; returns True if OK
'------------------------------------------------------------------
Dim i As Long, j As Long
On Error GoTo exit_ShowSplitUMCPoints
'make all spots invisible
Call GelCSExcludeAll(lngGelIndex)
Call GelIsoExcludeAll(lngGelIndex)
'now go and make all UMC spots visible
With GelUMC(lngGelIndex)
  For i = 0 To .UMCCnt - 1
      With .UMCs(i)
        If (.ClassStatusBits And UMC_INDICATOR_BIT_SPLIT_UMC) = UMC_INDICATOR_BIT_SPLIT_UMC Then
            For j = 0 To .ClassCount - 1
                Select Case .ClassMType(j)
                Case glCSType
                  GelDraw(lngGelIndex).CSID(.ClassMInd(j)) = Abs(GelDraw(lngGelIndex).CSID(.ClassMInd(j)))
                Case glIsoType
                  GelDraw(lngGelIndex).IsoID(.ClassMInd(j)) = Abs(GelDraw(lngGelIndex).IsoID(.ClassMInd(j)))
                End Select
            Next j
        End If
      End With
  Next i
End With
ShowSplitUMCPoints = True
exit_ShowSplitUMCPoints:
End Function

' Unused procedure
'''Public Function fUMCScore(ByVal IndDis As Long) As Boolean
'''Dim i As Long
'''On Error GoTo exit_fUMCScore
'''With GelUMC(IndDis)
'''     For i = 0 To .UMCCnt - 1
'''
'''     Next i
'''End With
'''exit_fUMCScore:
'''End Function
'''

Public Function fUMCCharacteristicPoints(Ind As Long, UMCInd As Long, _
                                         ChPoints() As LaV2DGPoint) As Boolean
'--------------------------------------------------------------------------------
'fills characteristic points of unique mass class; returns True if successful
'characteristic points are first(scan) point; last(scan) and class representative
'First point has coordinate NET, MW as first point in class(first scan; too bad
'if there are more than one point in first scan) and class abundance
'Middle point has NET coordinate of class representative and abundance of class
'representative but class mass as MW coordinate
'Third point is similar like first point with point from last scan
'--------------------------------------------------------------------------------
On Error GoTo exit_fUMCCharacteristicPoints
With GelUMC(Ind).UMCs(UMCInd)
     ReDim ChPoints(2)
'     Select Case .ClassMType(0)
'     Case glCSType
'          ChPoints(0).MW = GelData(Ind).CSData(.ClassMInd(0)).AverageMW
'          ChPoints(0).SCAN = GelData(Ind).CSData(.ClassMInd(0)).ScanNumber
'          ChPoints(0).Abu = GelData(Ind).CSData(.ClassMInd(0)).Abundance
'     Case glIsoType
'          ChPoints(0).MW = GelData(Ind).IsoData(.ClassMInd(0), GelUMC(Ind).def.MWField)
'          ChPoints(0).SCAN = GelData(Ind).IsoData(.ClassMInd(0)).ScanNumber
'          ChPoints(0).Abu = GelData(Ind).IsoData(.ClassMInd(0)).Abundance
'     End Select
'     Select Case .ClassRepType
'     Case glCSType
'          ChPoints(1).MW = GelData(Ind).CSData(.ClassMInd(0)).AverageMW
'          ChPoints(1).SCAN = GelData(Ind).CSData(.ClassRepInd).ScanNumber
'          ChPoints(1).Abu = GelData(Ind).CSData(.ClassMInd(0)).Abundance
'     Case glIsoType
'          ChPoints(1).MW = GelData(Ind).IsoData(.ClassMInd(0), GelUMC(Ind).def.MWField)
'          ChPoints(1).SCAN = GelData(Ind).IsoData(.ClassRepInd).ScanNumber
'          ChPoints(1).Abu = GelData(Ind).IsoData(.ClassMInd(0)).Abundance
'     End Select
'     Select Case .ClassMType(.ClassCount - 1)
'     Case glCSType
'          ChPoints(2).MW = GelData(Ind).CSData(.ClassMInd(0)).AverageMW
'          ChPoints(2).SCAN = GelData(Ind).CSData(.ClassMInd(.ClassCount - 1)).ScanNumber
'          ChPoints(2).Abu = GelData(Ind).CSData(.ClassMInd(0)).Abundance
'     Case glIsoType
'          ChPoints(2).MW = GelData(Ind).IsoData(.ClassMInd(0), GelUMC(Ind).def.MWField)
'          ChPoints(2).SCAN = GelData(Ind).IsoData(.ClassMInd(.ClassCount - 1)).ScanNumber
'          ChPoints(2).Abu = GelData(Ind).IsoData(.ClassMInd(0)).Abundance
'     End Select
     Select Case .ClassMType(0)
     Case glCSType
          ChPoints(0).MW = GelData(Ind).CSData(.ClassMInd(0)).AverageMW
          ChPoints(0).Scan = GelData(Ind).CSData(.ClassMInd(0)).ScanNumber
          ChPoints(0).Abu = .ClassAbundance
     Case glIsoType
          ChPoints(0).MW = GetIsoMass(GelData(Ind).IsoData(.ClassMInd(0)), GelUMC(Ind).def.MWField)
          ChPoints(0).Scan = GelData(Ind).IsoData(.ClassMInd(0)).ScanNumber
          ChPoints(0).Abu = .ClassAbundance
     End Select
     Select Case .ClassRepType
     Case glCSType
          ChPoints(1).MW = .ClassMW
          ChPoints(1).Scan = GelData(Ind).CSData(.ClassRepInd).ScanNumber
          ChPoints(1).Abu = GelData(Ind).CSData(.ClassRepInd).Abundance
     Case glIsoType
          ChPoints(1).MW = .ClassMW
          ChPoints(1).Scan = GelData(Ind).IsoData(.ClassRepInd).ScanNumber
          ChPoints(1).Abu = GelData(Ind).IsoData(.ClassRepInd).Abundance
     End Select
     Select Case .ClassMType(.ClassCount - 1)
     Case glCSType
          ChPoints(2).MW = GelData(Ind).CSData(.ClassMInd(.ClassCount - 1)).AverageMW
          ChPoints(2).Scan = GelData(Ind).CSData(.ClassMInd(.ClassCount - 1)).ScanNumber
          ChPoints(2).Abu = .ClassAbundance
     Case glIsoType
          ChPoints(2).MW = GetIsoMass(GelData(Ind).IsoData(.ClassMInd(.ClassCount - 1)), GelUMC(Ind).def.MWField)
          ChPoints(2).Scan = GelData(Ind).IsoData(.ClassMInd(.ClassCount - 1)).ScanNumber
          ChPoints(2).Abu = .ClassAbundance
     End Select
End With
fUMCCharacteristicPoints = True
exit_fUMCCharacteristicPoints:

End Function


Private Function UMCInterpolateGapsAbundance_Lin( _
                        ByVal lngGelIndex As Long, _
                        ByRef Scans() As Long, _
                        ByRef Abundances() As Double, _
                        ByVal MaxGapToInterpolate As Long) As Double
'-------------------------------------------------------------------------------------
'corrects class abundance by interpolating missing ions in the class; if lngScanGap is larger
'than the MaxGapToInterpolate then don't interpolate; return dblAbundanceSum of all abundances
'NOTE: function assumes that scans are coming in ascending order
'-------------------------------------------------------------------------------------

Const INTERPOLATE_USING_RELATIVE_SCAN_NUMBERS As Boolean = True

Dim lngindex As Long
Dim lngScanCount As Long
Dim dblAbundanceSum As Double
Dim lngScanGap As Long
Dim lngScanNumber1 As Long, lngScanNumber2 As Long
Dim lngScanIndex1 As Long, lngScanIndex2 As Long

On Error Resume Next
lngScanCount = UBound(Scans) + 1
If lngScanCount > 0 Then
    dblAbundanceSum = Abundances(0)
    If lngScanCount > 1 Then
        For lngindex = 1 To lngScanCount - 1
            dblAbundanceSum = dblAbundanceSum + Abundances(lngindex)         'add regular abundance
            
            If INTERPOLATE_USING_RELATIVE_SCAN_NUMBERS Then
                
                ' Since LTQ-FT and LTQ-Orbitrap data can have gaps between scans, need to use
                ' LookupScanNumberRelativeIndex to determine the relative index for each scan number
    
                lngScanNumber1 = Scans(lngindex - 1)
                lngScanIndex1 = LookupScanNumberRelativeIndex(lngGelIndex, lngScanNumber1)
                If lngScanIndex1 = 0 Then
                    lngScanNumber1 = LookupScanNumberClosest(lngGelIndex, lngScanNumber1)
                    lngScanIndex1 = LookupScanNumberRelativeIndex(lngGelIndex, lngScanNumber1)
                End If
            
                lngScanNumber2 = Scans(lngindex)
                lngScanIndex2 = LookupScanNumberRelativeIndex(lngGelIndex, lngScanNumber2)
                If lngScanIndex2 = 0 Then
                    lngScanNumber2 = LookupScanNumberClosest(lngGelIndex, lngScanNumber2)
                    lngScanIndex2 = LookupScanNumberRelativeIndex(lngGelIndex, lngScanNumber2)
                End If
                
                lngScanGap = lngScanIndex2 - lngScanIndex1
            Else
                lngScanGap = Scans(lngindex) - Scans(lngindex - 1)
            End If
                
            If lngScanGap > 1 Then
                ' Have to insert (lngScanGap-1) ions (for each missing scan)
                
                If lngScanGap - 1 <= MaxGapToInterpolate Then
                    dblAbundanceSum = dblAbundanceSum + (lngScanGap - 1) * (Abundances(lngindex - 1) + Abundances(lngindex)) / 2
                End If
            End If
            
        Next lngindex
    End If
    UMCInterpolateGapsAbundance_Lin = dblAbundanceSum
Else
    UMCInterpolateGapsAbundance_Lin = -1
End If
End Function

Public Function ManageClasses(ByVal Ind As Long, ByVal eManageType As UMCManageConstants) As Boolean
'------------------------------------------------------------------------------------
'prepares room for unique mass classes or reinitializes structures
'------------------------------------------------------------------------------------
Dim Cnt As Long
On Error GoTo exit_ManageClasses
With GelUMC(Ind)
    Select Case eManageType
    Case UMCManageConstants.UMCMngInitialize             ' Initially reserve space for .DataLines / 100 LC-MS Features
        Cnt = CLng(GelData(Ind).DataLines / 100)
        If Cnt < 10 Then Cnt = 10
        ReDim .UMCs(Cnt)
        .UMCCnt = 0
        .MassCorrectionValuesDefined = False
    Case UMCManageConstants.UMCMngTrim
        If .UMCCnt > 0 Then
           ReDim Preserve .UMCs(.UMCCnt - 1)
        Else
            .UMCCnt = 0
           Erase .UMCs
           .MassCorrectionValuesDefined = False
        End If
    Case UMCManageConstants.UMCMngErase
        .UMCCnt = 0
        Erase .UMCs
        .MassCorrectionValuesDefined = False
    Case UMCManageConstants.UMCMngAdd                    ' Increase space reserved by 50%
        Cnt = CLng(UBound(.UMCs) / 2)
        If Cnt < 5 Then Cnt = 5
        ReDim Preserve .UMCs(.UMCCnt + Cnt)
    End Select
End With

If eManageType = UMCManageConstants.UMCMngInitialize Or eManageType = UMCManageConstants.UMCMngErase Then
    ' Make sure no UMC-based pairs exist
    DestroyDltLblPairs Ind
End If

ManageClasses = True
exit_ManageClasses:
End Function

Public Function CalculateClasses(ByVal lngGelIndex As Long, Optional blnUseProgressForm As Boolean = False, Optional frmCallingForm As VB.Form) As Boolean
'--------------------------------------------------------------------------------
'Recalculates parameters of the unique mass classes; returns True on success
'
'Oct 2003: Expanded to find the statistics for each group of points within a UMC
'  with the same charge state; optionally, use the stats for the "most abundant"
'  charge-state based group for the UMC class mass and class abundance
'
'May 2004: Expanded to allow use of subset of members of UMC for computing mass and abundance stats, as
'            specified by glbPreferencesExpanded.UMCAdvancedStatsOptions
'          MinScan and MaxScan are still computed using the entire class
'
'--------------------------------------------------------------------------------
Const MAX_CHARGE_STATE As Integer = 10             ' Any data with a charge state over this will be grouped with the 10+ charge state
Const INITIAL_RESERVE_COUNT As Integer = 1000

Dim i As Long, j As Long
Dim lngMaxMemberIndex As Long

Dim UMCMembersMaxIndex As Long
Dim UMCMembersMW() As Double                   ' 0-based array; MW values for each of the LC-MS Features members
Dim UMCMembersAbu() As Double                  ' 0-based array; Abu values for each of the LC-MS Features members
Dim UMCMembersScan() As Long                   ' 0-based array; Scan numbers for each of the LC-MS Features members
Dim UMCMembersFit() As Double                  ' 0-based array; Fit values for each of the LC-MS Features members
Dim UMCMembersCharge() As Integer              ' 0-based array; Charge states for each of the LC-MS Features members

Dim intUMCRepresentativeType As Integer
Dim intChargeState As Integer
Dim ChargeStatePresent(0 To MAX_CHARGE_STATE) As Long       ' 0-based array; Used to keep track of which charge states are present
Dim intChargeStatesPresent As Integer

Dim lngChargeStateValueCount As Long
Dim lngChargeStateGroupRepIndex As Long      ' Index in the ChargeStateBased... arrays
Dim lngClassMIndexPointer As Long

Dim ChargeStateMaxIndex As Long
Dim ChargeStateBasedMW() As Double      ' 0-based array; Used when computing charge state based stats
Dim ChargeStateBasedAbu() As Double     ' 0-based array; Used when computing charge state based stats
Dim ChargeStateBasedScan() As Long      ' 0-based array; Used when computing charge state based stats
Dim ChargeStateBasedFit() As Double     ' 0-based array; Used when computing charge state based stats
Dim ChargeStateBasedOrgIndex() As Long     ' Original index in the member arrays (UMCMembersMW, UMCMembersAbu, etc.) of each ion in the ChargeStateBased arrays

Dim dblConglomerateMW As Double
Dim dblConglomerateMWStD As Double
Dim dblConglomerateAbu As Double

Dim dblBestValue As Double, dblCompareValue As Double
Dim intBestIndex As Integer

Dim RepMW As Double
Dim RepAbu As Double
Dim ISMWField As Integer

Dim strCaptionSaved As String
Dim blnShowProgressUsingFormCaption As Boolean

On Error GoTo err_CalculateClasses

If blnUseProgressForm Then
    frmProgress.InitializeSubtask "Updating LC-MS Feature Stats", 0, GelUMC(lngGelIndex).UMCCnt
Else
    If Not frmCallingForm Is Nothing Then
        blnShowProgressUsingFormCaption = True
        strCaptionSaved = frmCallingForm.Caption
    End If
End If

With glbPreferencesExpanded.UMCAdvancedStatsOptions
    If .ClassMassTopXMinMembers < 1 Then .ClassMassTopXMinMembers = 1
    If .ClassAbuTopXMinMembers < 1 Then .ClassAbuTopXMinMembers = 1
End With

' Initially reserve space for LC-MS Features with INITIAL_RESERVE_COUNT members
UMCMembersMaxIndex = 0
ReDim UMCMembersMW(INITIAL_RESERVE_COUNT)
ReDim UMCMembersAbu(INITIAL_RESERVE_COUNT)
ReDim UMCMembersScan(INITIAL_RESERVE_COUNT)
ReDim UMCMembersFit(INITIAL_RESERVE_COUNT)
ReDim UMCMembersCharge(INITIAL_RESERVE_COUNT)

ChargeStateMaxIndex = 0
ReDim ChargeStateBasedMW(INITIAL_RESERVE_COUNT)
ReDim ChargeStateBasedAbu(INITIAL_RESERVE_COUNT)
ReDim ChargeStateBasedScan(INITIAL_RESERVE_COUNT)
ReDim ChargeStateBasedFit(INITIAL_RESERVE_COUNT)
ReDim ChargeStateBasedOrgIndex(INITIAL_RESERVE_COUNT)

With GelUMC(lngGelIndex)
    ISMWField = .def.MWField
    If .UMCCnt > 0 Then
       For i = 0 To .UMCCnt - 1
           With .UMCs(i)
               '
               ' First, find the Min and Max scan numbers, the Min and Max MW,
               ' and populate the UMCMembersMW, UMCMembersScan, etc. arrays
               '
               ' Also determine which charge states are present
               '
               .MinScan = glHugeLong:               .MaxScan = -glHugeLong
               .MinMW = glHugeDouble:               .MaxMW = -glHugeDouble
               If .ClassCount > 0 Then
                  lngMaxMemberIndex = .ClassCount - 1
                  UMCMembersMaxIndex = lngMaxMemberIndex
                  Do While UMCMembersMaxIndex > UBound(UMCMembersMW)
                    ReDim UMCMembersMW(UBound(UMCMembersMW) * 2)
                    ReDim UMCMembersAbu(UBound(UMCMembersMW))
                    ReDim UMCMembersScan(UBound(UMCMembersMW))
                    ReDim UMCMembersFit(UBound(UMCMembersMW))
                    ReDim UMCMembersCharge(UBound(UMCMembersMW))
                  Loop
                  Erase ChargeStatePresent()       ' Reset all to 0
                  For j = 0 To lngMaxMemberIndex
                      Select Case .ClassMType(j)
                      Case glCSType
                           UMCMembersMW(j) = GelData(lngGelIndex).CSData(.ClassMInd(j)).AverageMW
                           UMCMembersAbu(j) = GelData(lngGelIndex).CSData(.ClassMInd(j)).Abundance
                           UMCMembersScan(j) = GelData(lngGelIndex).CSData(.ClassMInd(j)).ScanNumber
                           UMCMembersFit(j) = GelData(lngGelIndex).CSData(.ClassMInd(j)).MassStDev        ' Isotopic fit is not defined for charge state data; use standard deviation instead
                           UMCMembersCharge(j) = GelData(lngGelIndex).CSData(.ClassMInd(j)).Charge
                      Case glIsoType
                           UMCMembersMW(j) = GetIsoMass(GelData(lngGelIndex).IsoData(.ClassMInd(j)), ISMWField)
                           UMCMembersAbu(j) = GelData(lngGelIndex).IsoData(.ClassMInd(j)).Abundance
                           UMCMembersScan(j) = GelData(lngGelIndex).IsoData(.ClassMInd(j)).ScanNumber
                           UMCMembersFit(j) = GelData(lngGelIndex).IsoData(.ClassMInd(j)).Fit
                           UMCMembersCharge(j) = GelData(lngGelIndex).IsoData(.ClassMInd(j)).Charge
                      End Select
                      If UMCMembersMW(j) < .MinMW Then .MinMW = UMCMembersMW(j)
                      If UMCMembersMW(j) > .MaxMW Then .MaxMW = UMCMembersMW(j)
                      If UMCMembersScan(j) < .MinScan Then .MinScan = UMCMembersScan(j)
                      If UMCMembersScan(j) > .MaxScan Then .MaxScan = UMCMembersScan(j)
                      If UMCMembersCharge(j) >= 0 And UMCMembersCharge(j) <= MAX_CHARGE_STATE Then
                         ChargeStatePresent(UMCMembersCharge(j)) = ChargeStatePresent(UMCMembersCharge(j)) + 1
                      Else
                         ChargeStatePresent(MAX_CHARGE_STATE) = ChargeStatePresent(MAX_CHARGE_STATE) + 1
                      End If
                  Next j
               End If
           End With
           
           ' Determine the number of charge states present
           intChargeStatesPresent = 0
           For intChargeState = 0 To MAX_CHARGE_STATE
                If ChargeStatePresent(intChargeState) > 0 Then
                    intChargeStatesPresent = intChargeStatesPresent + 1
                End If
           Next intChargeState
           
           ' Reset the charge state based stats
           With .UMCs(i)
                .ChargeStateStatsRepInd = 0
                .ChargeStateCount = 0
                
                If intChargeStatesPresent > 0 Then
                    ReDim .ChargeStateBasedStats(intChargeStatesPresent - 1)
                Else
                    ReDim .ChargeStateBasedStats(0)
                End If
           End With
           
           If .UMCs(i).ClassCount > 0 Then
                ' Compute the charge state based stats
                ' Step through UMCMembersCharge() and copy the MW, Abu, etc. values to the
                '  appropriate place in the ChargeBasedStats arrays
                ' We always compute the charge state based stats, even if .UMCClassStatsUseStatsFromMostAbuChargeState = False,
                '  since this information is required for ER computations when finding pairs
                For intChargeState = 0 To MAX_CHARGE_STATE
                    If ChargeStatePresent(intChargeState) > 0 Then
                        lngChargeStateValueCount = 0
                        ChargeStateMaxIndex = ChargeStatePresent(intChargeState) - 1
                        Do While ChargeStateMaxIndex > UBound(ChargeStateBasedMW)
                            ReDim ChargeStateBasedMW(UBound(ChargeStateBasedMW) * 2)
                            ReDim ChargeStateBasedAbu(UBound(ChargeStateBasedMW))
                            ReDim ChargeStateBasedScan(UBound(ChargeStateBasedMW))
                            ReDim ChargeStateBasedFit(UBound(ChargeStateBasedMW))
                            ReDim ChargeStateBasedOrgIndex(UBound(ChargeStateBasedMW))
                        Loop

                        For j = 0 To lngMaxMemberIndex
                            If UMCMembersCharge(j) = intChargeState Or _
                               (intChargeState = MAX_CHARGE_STATE And (UMCMembersCharge(j) < 0 Or UMCMembersCharge(j) > MAX_CHARGE_STATE)) Then

                                ChargeStateBasedMW(lngChargeStateValueCount) = UMCMembersMW(j)
                                ChargeStateBasedAbu(lngChargeStateValueCount) = UMCMembersAbu(j)
                                ChargeStateBasedScan(lngChargeStateValueCount) = UMCMembersScan(j)
                                ChargeStateBasedFit(lngChargeStateValueCount) = UMCMembersFit(j)
                                ChargeStateBasedOrgIndex(lngChargeStateValueCount) = j

                                lngChargeStateValueCount = lngChargeStateValueCount + 1
                                Debug.Assert lngChargeStateValueCount <= ChargeStatePresent(intChargeState)
                            End If
                        Next j

                        If lngChargeStateValueCount > 0 Then

                            ' Determine the appropriate UMC Representative selection method
                            Select Case .def.UMCType
                            Case glUMC_TYPE_INTENSITY, glUMC_TYPE_ISHRINKINGBOX, glUMC_TYPE_MINCNT, glUMC_TYPE_MAXCNT, glUMC_TYPE_UNQAMT
                                intUMCRepresentativeType = UMCFROMNet_REP_ABU
                            Case glUMC_TYPE_FIT, glUMC_TYPE_FSHRINKINGBOX
                                intUMCRepresentativeType = UMCFROMNet_REP_FIT
                            Case glUMC_TYPE_FROM_NET
                                intUMCRepresentativeType = glbPreferencesExpanded.UMCIonNetOptions.UMCRepresentative
                            Case Else
                                ' Invalid option defined for .UMCType
                                Debug.Assert False
                                intUMCRepresentativeType = UMCFROMNet_REP_ABU
                            End Select

                            ' Determine the index of the representative ion for this charge state group
                            lngChargeStateGroupRepIndex = CalculateClassesFindRepIndex(ChargeStateMaxIndex, ChargeStateBasedAbu(), ChargeStateBasedFit(), intUMCRepresentativeType)

                            ' Lookup the Charge Group's Rep MW and Abu; needed for call to CalculateClassesComputeStats
                            RepMW = ChargeStateBasedMW(lngChargeStateGroupRepIndex)
                            RepAbu = ChargeStateBasedAbu(lngChargeStateGroupRepIndex)

                            ' Compute the stats for this charge state group
                            CalculateClassesComputeStats lngGelIndex, .def, ChargeStateMaxIndex, ChargeStateBasedMW(), ChargeStateBasedAbu(), ChargeStateBasedScan(), RepMW, RepAbu, dblConglomerateMW, dblConglomerateMWStD, dblConglomerateAbu

                            With .UMCs(i)
                                With .ChargeStateBasedStats(.ChargeStateCount)
                                    .Charge = intChargeState
                                    .Count = lngChargeStateValueCount
                                    .Mass = dblConglomerateMW
                                    .MassStD = dblConglomerateMWStD
                                    .Abundance = dblConglomerateAbu
                                    .GroupRepIndex = ChargeStateBasedOrgIndex(lngChargeStateGroupRepIndex)
                                End With

                                .ChargeStateCount = .ChargeStateCount + 1
                            End With
                        Else
                            ' This code shouldn't be reached
                            Debug.Assert False
                        End If
                    End If
                Next intChargeState

                If .UMCs(i).ChargeStateCount <= 1 Then
                    ' The "Best" Charge State Index must be the only index: 0
                    intBestIndex = 0
                Else
                    ' Determine .ChargeStateStatsRepInd based on the value of .ChargeStateStatsRepType
                    intBestIndex = 0
                    Select Case .def.ChargeStateStatsRepType
                    Case UMCChargeStateGroupConstants.UMCCSGHighestSum
                        ' Find the charge state group with the highest Abundance Sum
                        With .UMCs(i)
                            dblBestValue = -glHugeDouble
                            intBestIndex = 0
                            For intChargeState = 0 To .ChargeStateCount - 1
                                If .ChargeStateBasedStats(intChargeState).Abundance > dblBestValue Then
                                    dblBestValue = .ChargeStateBasedStats(intChargeState).Abundance
                                    intBestIndex = intChargeState
                                End If
                            Next intChargeState
                        End With
                    Case UMCChargeStateGroupConstants.UMCCSGMostAbuMember
                        ' Find the charge state group containing the highest intensity ion in this UMC
                        With .UMCs(i)
                            dblBestValue = -glHugeDouble
                            intBestIndex = 0
                            For intChargeState = 0 To .ChargeStateCount - 1
                                lngClassMIndexPointer = .ChargeStateBasedStats(intChargeState).GroupRepIndex
                                Select Case .ClassMType(lngClassMIndexPointer)
                                Case glCSType
                                     dblCompareValue = GelData(lngGelIndex).CSData(lngClassMIndexPointer).Abundance
                                Case glIsoType
                                     dblCompareValue = GelData(lngGelIndex).IsoData(lngClassMIndexPointer).Abundance
                                End Select
                                
                                If dblCompareValue > dblBestValue Then
                                    dblBestValue = dblCompareValue
                                    intBestIndex = intChargeState
                                End If
                            Next intChargeState
                        End With
                    Case UMCChargeStateGroupConstants.UMCCSGMostMembers
                        ' Find the charge state group with the most members
                        With .UMCs(i)
                            dblBestValue = 0
                            intBestIndex = 0
                            For intChargeState = 0 To .ChargeStateCount - 1
                                If .ChargeStateBasedStats(intChargeState).Count > dblBestValue Then
                                    dblBestValue = .ChargeStateBasedStats(intChargeState).Count
                                    intBestIndex = intChargeState
                                End If
                            Next intChargeState
                        End With
                    Case Else
                        ' Invalid type
                        Debug.Assert False
                    End Select
                End If
                .UMCs(i).ChargeStateStatsRepInd = intBestIndex

                ' Populate .ClassMW, .ClassMWStD, and .ClassAbundance
                If .def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                    With .UMCs(i)
                        .ClassMW = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Mass
                        .ClassMWStD = .ChargeStateBasedStats(.ChargeStateStatsRepInd).MassStD
                        .ClassAbundance = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Abundance
                        
                        ' Update .ClassRepInd and .ClassRepType to contain the ClassRep values for
                        '   for the class rep of the best charge state group
                        lngClassMIndexPointer = .ChargeStateBasedStats(.ChargeStateStatsRepInd).GroupRepIndex
                        .ClassRepInd = .ClassMInd(lngClassMIndexPointer)
                        .ClassRepType = .ClassMType(lngClassMIndexPointer)
                    End With
                Else
                    ' Lookup the Class Rep MW and Abu; needed for call to CalculateClassesComputeStats
                    ' Note that .ClassRepInd and .ClassRepType will have already been correctly defined
                    With .UMCs(i)
                        Select Case .ClassRepType
                        Case glCSType
                             RepMW = GelData(lngGelIndex).CSData(.ClassRepInd).AverageMW
                             RepAbu = GelData(lngGelIndex).CSData(.ClassRepInd).Abundance
                        Case glIsoType
                             RepMW = GetIsoMass(GelData(lngGelIndex).IsoData(.ClassRepInd), ISMWField)
                             RepAbu = GelData(lngGelIndex).IsoData(.ClassRepInd).Abundance
                        End Select
                    End With
                    
                    ' Compute the class stats
                    CalculateClassesComputeStats lngGelIndex, .def, UMCMembersMaxIndex, UMCMembersMW(), UMCMembersAbu(), UMCMembersScan(), RepMW, RepAbu, dblConglomerateMW, dblConglomerateMWStD, dblConglomerateAbu
                    
                    With .UMCs(i)
                        .ClassMW = dblConglomerateMW
                        .ClassMWStD = dblConglomerateMWStD
                        .ClassAbundance = dblConglomerateAbu
                    End With
                End If
                
           Else                             'something is wrong
                ' This code shouldn't be reached
                Debug.Assert False
                With .UMCs(i)
                    .ClassAbundance = -1
                    .ClassMW = -1
                    .ClassMWStD = -1
                    .MinScan = -1
                    .MaxScan = -1
                End With
           End If
           If i Mod 500 = 0 Then
              If blnShowProgressUsingFormCaption Then
                  frmCallingForm.Caption = "Updating LC-MS Feature Stats: " & Trim(i) & " / " & (.UMCCnt)
              ElseIf blnUseProgressForm Then
                  frmProgress.UpdateSubtaskProgressBar i
              End If
           End If
       Next i
    End If
End With
If blnShowProgressUsingFormCaption Then frmCallingForm.Caption = strCaptionSaved
CalculateClasses = True

Exit Function

err_CalculateClasses:
Debug.Assert False
'Resume Next
LogErrors Err.Number, "CalculateClasses"
On Error Resume Next
If blnShowProgressUsingFormCaption Then frmCallingForm.Caption = strCaptionSaved
End Function

Private Sub CalculateClassesComputeStats(ByVal lngGelIndex As Long, ByRef udtUMCDef As UMCDefinition, ByVal SrcMaxIndex As Long, ByRef MWArray() As Double, ByRef AbuArray() As Double, ByRef ScanArray() As Long, ByVal dblRepMW As Double, ByVal dblRepAbu As Double, ByRef dblConglomerateMW As Double, ByRef dblConglomerateMWStD As Double, ByRef dblConglomerateAbu As Double)
    
    ' Note: This function previously used the StatDoubles class
    ' However, this can lead to poor performance in the compiled version of the program, so we've now switched to computing the stats locally
    
    Dim blnValidData As Boolean
    Dim lngindex As Long
    
    Dim dblMWWork() As Double
    Dim dblAbuWork() As Double
    Dim lngScanWork() As Long
    
    Dim dblStatMaximum As Double
    Dim dblStatSum As Double
    Dim dblStatSumSquares As Double
    Dim dblStDevScratch As Double
    
    Dim lngHalfIndex As Long
    Dim lngDataCount As Long
    
On Error GoTo CalculateClassesComputeStatsErrorHandler

    blnValidData = False
    If udtUMCDef.ClassMW = UMCClassMassConstants.UMCMassAvgTopX Or udtUMCDef.ClassMW = UMCClassMassConstants.UMCMassMedTopX Then
        ' Note: Send MWArray() as the SrcArray here
        With glbPreferencesExpanded.UMCAdvancedStatsOptions
            blnValidData = CalculateClassesFindMemberSubsetByAbu(SrcMaxIndex, MWArray(), AbuArray(), ScanArray(), dblMWWork(), lngScanWork(), .ClassMassTopXMinAbu, .ClassMassTopXMaxAbu, .ClassMassTopXMinMembers)
        End With
    End If
    
    If Not blnValidData Then
        ' Copy MWArray to dblMWWork
        ReDim dblMWWork(0 To SrcMaxIndex)
        lngDataCount = SrcMaxIndex + 1
        
        For lngindex = 0 To SrcMaxIndex
            dblMWWork(lngindex) = MWArray(lngindex)
        Next lngindex
    Else
        lngDataCount = UBound(dblMWWork) + 1
    End If
    
    ' Compute the necessary stats
    dblStatSum = 0
    dblStatSumSquares = 0
    dblStatMaximum = dblMWWork(0)
    For lngindex = 0 To lngDataCount - 1
        If dblMWWork(lngindex) > dblStatMaximum Then
            dblStatMaximum = dblMWWork(lngindex)
        End If
        dblStatSum = dblStatSum + dblMWWork(lngindex)
        dblStatSumSquares = dblStatSumSquares + dblMWWork(lngindex) * dblMWWork(lngindex)
    Next lngindex
    
    ' Determine class mass
    If lngDataCount > 0 Then
        Select Case udtUMCDef.ClassMW
        Case UMCClassMassConstants.UMCMassAvg, UMCClassMassConstants.UMCMassAvgTopX
            dblConglomerateMW = dblStatSum / lngDataCount
        Case UMCClassMassConstants.UMCMassRep
            dblConglomerateMW = dblRepMW
        Case UMCClassMassConstants.UMCMassMed, UMCClassMassConstants.UMCMassMedTopX
            ' Compute the median
            ' First, sort dblMWWork
            ShellSortDouble dblMWWork, 0, lngDataCount - 1
            lngHalfIndex = Int(lngDataCount / 2)
            If lngDataCount Mod 2 > 0 Then               'odd number of elements
               dblConglomerateMW = dblMWWork(lngHalfIndex)
            Else                                 'even number of elements
               dblConglomerateMW = (dblMWWork(lngHalfIndex - 1) + dblMWWork(lngHalfIndex)) / 2
            End If
        Case Else
            ' Invalid option
            Debug.Assert False
            dblConglomerateMW = -1
        End Select
        If lngDataCount > 1 Then
            dblStDevScratch = dblStatSumSquares / (lngDataCount - 1) - (dblStatSum * dblStatSum) / (lngDataCount * (lngDataCount - 1))
            If dblStDevScratch > 0 Then
                dblConglomerateMWStD = Sqr(dblStDevScratch)
            Else
                dblConglomerateMWStD = 0
            End If
        Else
            dblConglomerateMWStD = 0
        End If
    Else
        dblConglomerateMW = -1
        dblConglomerateMWStD = 0
    End If
    
    blnValidData = False
    If udtUMCDef.ClassAbu = UMCClassAbundanceConstants.UMCAbuSumTopX Then
        ' Note: Send AbuArray() as the SrcArray here
        With glbPreferencesExpanded.UMCAdvancedStatsOptions
            blnValidData = CalculateClassesFindMemberSubsetByAbu(SrcMaxIndex, AbuArray(), AbuArray(), ScanArray(), dblAbuWork(), lngScanWork(), .ClassAbuTopXMinAbu, .ClassAbuTopXMaxAbu, .ClassAbuTopXMinMembers)
            
            If blnValidData Then
                ' The scan number array needs to be sorted ascending for use with UMCInterpolateGapsAbundance_Lin
                ' Sort it, and dblAbuWork parallel with it
                ShellSortLongWithParallelDouble lngScanWork, dblAbuWork, 0, UBound(lngScanWork)
            End If
            
        End With
    End If
    
    If Not blnValidData Then
        ' Copy AbuArray to dblAbuWork and ScanArray to lngScanWork
        ReDim dblAbuWork(0 To SrcMaxIndex)
        ReDim lngScanWork(0 To SrcMaxIndex)
        lngDataCount = SrcMaxIndex + 1
        
        For lngindex = 0 To SrcMaxIndex
            dblAbuWork(lngindex) = AbuArray(lngindex)
            lngScanWork(lngindex) = ScanArray(lngindex)
        Next lngindex
    Else
        lngDataCount = UBound(dblAbuWork) + 1
    End If
    
    ' Compute the necessary stats
    dblStatSum = 0
    dblStatSumSquares = 0
    dblStatMaximum = dblAbuWork(0)
    For lngindex = 0 To lngDataCount - 1
        If dblAbuWork(lngindex) > dblStatMaximum Then
            dblStatMaximum = dblAbuWork(lngindex)
        End If
        dblStatSum = dblStatSum + dblAbuWork(lngindex)
    Next lngindex
    
    ' Determine class abundance
    If lngDataCount > 0 Then
        Select Case udtUMCDef.ClassAbu
        Case UMCClassAbundanceConstants.UMCAbuAvg
            dblConglomerateAbu = dblStatSum / lngDataCount
        Case UMCClassAbundanceConstants.UMCAbuSum, UMCClassAbundanceConstants.UMCAbuSumTopX
            If udtUMCDef.InterpolateGaps Then
                ' interpolate gaps (ghost ions)
                dblConglomerateAbu = UMCInterpolateGapsAbundance_Lin(lngGelIndex, lngScanWork(), dblAbuWork(), udtUMCDef.InterpolateMaxGapSize)
            Else
                ' normal sum
                dblConglomerateAbu = dblStatSum
            End If
        Case UMCClassAbundanceConstants.UMCAbuRep
            dblConglomerateAbu = dblRepAbu
        Case UMCClassAbundanceConstants.UMCAbuMed
         ' Compute the median
            ' First, sort dblAbuWork
            ShellSortDouble dblAbuWork, 0, lngDataCount - 1
            lngHalfIndex = Int(lngDataCount / 2)
            If lngDataCount Mod 2 > 0 Then               'odd number of elements
               dblConglomerateAbu = dblAbuWork(lngHalfIndex)
            Else                                 'even number of elements
               dblConglomerateAbu = (dblAbuWork(lngHalfIndex - 1) + dblAbuWork(lngHalfIndex)) / 2
            End If
        Case UMCClassAbundanceConstants.UMCAbuMax
            dblConglomerateAbu = dblStatMaximum
        Case Else
            ' Invalid option
            Debug.Assert False
            dblConglomerateAbu = -1
        End Select
    Else
        dblConglomerateAbu = -1
    End If

    Exit Sub
    
CalculateClassesComputeStatsErrorHandler:
Debug.Assert False
'Resume Next
LogErrors Err.Number, "CalculateClassesComputeStats"

End Sub

Private Function CalculateClassesFindMemberSubsetByAbu(ByVal SrcMaxIndex As Long, ByRef SrcArray() As Double, ByRef AbuArray() As Double, ByRef ScanArray() As Long, ByRef dblTargetWork() As Double, ByRef lngScanWork() As Long, ByVal TopXMinAbu As Double, ByVal TopXMaxAbu As Double, ByVal TopXMinMembers As Long) As Boolean
    ' Fills dblTargetWork() with values from SrcArray, using items from index 0 to index SrcMaxIndex
    ' Fills lngScanWork() with the corresponding scan numbers
    ' Data chosen is a subset of the data in SrcArray, determined by examining AbuArray
    
    Dim dblAbuSorted() As Double
    Dim lngAbuPointers() As Long
    Dim lngindex As Long
    Dim lngIndexToCopy As Long          ' Pointer to location in dblAbuSorted
    Dim lngIndexHighAbundanceThreshold As Long
    
    Dim lngWorkCount As Long

On Error GoTo CalculateClassesFindMemberSubsetByAbuErrorHandler

    ' We need to sort the values in SrcArray and AbuArray by abundance
    If UBound(SrcArray) < SrcMaxIndex Or UBound(AbuArray) < SrcMaxIndex Then
        CalculateClassesFindMemberSubsetByAbu = False
        Exit Function
    End If
    
    ReDim dblAbuSorted(0 To SrcMaxIndex)
    For lngindex = 0 To SrcMaxIndex
        dblAbuSorted(lngindex) = AbuArray(lngindex)
    Next lngindex

    ' Initialize an index array
    ReDim lngAbuPointers(SrcMaxIndex)
    For lngindex = 0 To SrcMaxIndex
        lngAbuPointers(lngindex) = lngindex
    Next lngindex

    ' Sort Ascending
    ShellSortDoubleWithParallelLong dblAbuSorted, lngAbuPointers, 0, SrcMaxIndex
    'blnSuccess = objQSDouble.QSAsc(dblAbuSorted(), lngAbuPointers())

    ReDim dblTargetWork(0 To SrcMaxIndex)
    ReDim lngScanWork(0 To SrcMaxIndex)
    
    lngWorkCount = 0
    lngIndexToCopy = SrcMaxIndex
    
    ' Set TopXMinAbu to a huge number if both MinAbu and MaxAbu are <= 0
    ' This way, the number of points included in the returned array will be equal to TopXMinMembers
    If TopXMinAbu <= 0 And TopXMaxAbu <= 0 Then TopXMinAbu = 1E+308
    
    If TopXMaxAbu > 0 Then
        ' Find the first value in dblAbuSorted less than ClassMassTopXMaxAbu
        Do While lngIndexToCopy >= 0
            If dblAbuSorted(lngIndexToCopy) > TopXMaxAbu Then
                lngIndexToCopy = lngIndexToCopy - 1     ' Decrement lngIndexToCopy
            Else
                Exit Do
            End If
        Loop
        lngIndexHighAbundanceThreshold = lngIndexToCopy + 1
    End If
    
    ' Now look for matching values in dblAbuSorted, and copy their mass values into dblTargetWork
    Do While lngIndexToCopy >= 0
        If dblAbuSorted(lngIndexToCopy) >= TopXMinAbu Then
            dblTargetWork(lngWorkCount) = SrcArray(lngAbuPointers(lngIndexToCopy))
            lngScanWork(lngWorkCount) = ScanArray(lngAbuPointers(lngIndexToCopy))
            lngWorkCount = lngWorkCount + 1
            lngIndexToCopy = lngIndexToCopy - 1     ' Decrement lngIndexToCopy
        Else
            Exit Do
        End If
    Loop
    
    ' If not enough members were found to match the given range, then add members until enough are included
    Do While lngWorkCount < TopXMinMembers And lngIndexToCopy >= 0
        dblTargetWork(lngWorkCount) = SrcArray(lngAbuPointers(lngIndexToCopy))
        lngScanWork(lngWorkCount) = ScanArray(lngAbuPointers(lngIndexToCopy))
        lngWorkCount = lngWorkCount + 1
        lngIndexToCopy = lngIndexToCopy - 1         ' Decrement lngIndexToCopy
    Loop
    
    ' If we still do not have enough members, and if ClassMassTopXMaxAbu is > 0, then start adding the higher abundance values
    If lngWorkCount < TopXMinMembers And TopXMaxAbu > 0 Then
        lngIndexToCopy = lngIndexHighAbundanceThreshold
        Do While lngWorkCount < TopXMinMembers And lngIndexToCopy <= SrcMaxIndex
            dblTargetWork(lngWorkCount) = SrcArray(lngAbuPointers(lngIndexToCopy))
            lngScanWork(lngWorkCount) = ScanArray(lngAbuPointers(lngIndexToCopy))
            lngWorkCount = lngWorkCount + 1
            lngIndexToCopy = lngIndexToCopy + 1     ' Increment lngIndexToCopy
        Loop
    End If
    
    ' Shrink the arrays to contain lngWorkCount elements
    If lngWorkCount > 0 Then
        If lngWorkCount - 1 < UBound(dblTargetWork) Then
            ReDim Preserve dblTargetWork(lngWorkCount - 1)
            ReDim Preserve lngScanWork(lngWorkCount - 1)
        End If
        CalculateClassesFindMemberSubsetByAbu = True
    Else
        ' This code shouldn't be reached
        Debug.Assert False
        
        ReDim dblTargetWork(0)
        ReDim lngScanWork(0)
        
        dblTargetWork(0) = dblAbuSorted(SrcMaxIndex)
        lngWorkCount = 1
        
        CalculateClassesFindMemberSubsetByAbu = False
    End If

    Exit Function

CalculateClassesFindMemberSubsetByAbuErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "CalculateClassesFindMemberSubsetByAbu"
    CalculateClassesFindMemberSubsetByAbu = False
    
End Function

Private Function CalculateClassesFindRepIndex(ByVal SrcDataMaxIndex As Long, ByRef AbuList() As Double, ByRef FitList() As Double, ByVal intUMCRepresentativeType As Integer) As Long
    ' Determine the index of the ion in the given arrays that should be the "class representative"
    ' Returns the class rep index
    ' Note: this is the index of the ion in the given array and NOT the index of the ion in GelData()
    ' Assumes that the members of given arrays are ordered by Scan, and that the arrays are 0-based
    '
    ' Note that this Function is very similar to UMCIonNet->FindUMCClassRepIndex
    
    Dim i As Long
    Dim BestInd As Long
    Dim BestValue As Double
    
On Error GoTo CalculateClassesFindRepIndexErrorHandler

    Select Case intUMCRepresentativeType
    Case UMCFROMNet_REP_ABU
        BestInd = -1
        BestValue = -glHugeDouble
        For i = 0 To SrcDataMaxIndex
            If AbuList(i) > BestValue Then
                BestValue = AbuList(i)
                BestInd = i
            End If
        Next i
    Case UMCFROMNet_REP_FIT
        BestInd = -1
        BestValue = glHugeDouble
        For i = 0 To SrcDataMaxIndex
            If FitList(i) < BestValue Then
               BestValue = FitList(i)
               BestInd = i
            End If
        Next i
    Case UMCFROMNet_REP_FST_SCAN
        BestInd = 0
    Case UMCFROMNet_REP_LST_SCAN
        BestInd = SrcDataMaxIndex
    Case UMCFROMNet_REP_MED_SCAN
        BestInd = CLng((SrcDataMaxIndex + 1) / 2)
    End Select
    
    If BestInd < 0 Then
        ' This shouldn't happen
        Debug.Assert False
        BestInd = 0
    End If
    
    CalculateClassesFindRepIndex = BestInd
    
    Exit Function

CalculateClassesFindRepIndexErrorHandler:
    Debug.Print "Error in CalculateClassesFindRepIndex: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "UC.Bas->CalculateClassesFindRepIndex"
    CalculateClassesFindRepIndex = 0
    
End Function

Public Function GetUMCList(ByVal Ind As Long, _
                           Scan1 As Long, Scan2 As Long, _
                           MW1 As Double, MW2 As Double, _
                           ResList() As Long) As Long
'----------------------------------------------------------------------------------
'fills ResList with indexes of LC-MS Features that have at least one member in a elution/mass
'window (Scan1, Scan2) x (MW1,MW2); returns number of it, -1 on any error
'----------------------------------------------------------------------------------
Dim Cnt As Long
Dim i As Long
On Error GoTo err_GetUMCList
With GelUMC(Ind)
     ReDim ResList(.UMCCnt - 1)             'reserve enough room
     For i = 0 To .UMCCnt - 1
         With .UMCs(i)
              If (.MinScan >= Scan1) And (.MinScan <= Scan2) Then
                 If (.MinMW >= MW1) And (.MinMW <= MW2) Then
                    Cnt = Cnt + 1
                    ResList(Cnt - 1) = i
                 ElseIf (.MaxMW >= MW1) And (.MaxMW <= MW2) Then
                    Cnt = Cnt + 1
                    ResList(Cnt - 1) = i
                 End If
              ElseIf (.MaxScan >= Scan1) And (.MaxScan <= Scan2) Then
                 If (.MinMW >= MW1) And (.MinMW <= MW2) Then
                    Cnt = Cnt + 1
                    ResList(Cnt - 1) = i
                 ElseIf (.MaxMW >= MW1) And (.MaxMW <= MW2) Then
                    Cnt = Cnt + 1
                    ResList(Cnt - 1) = i
                 End If
              End If
         End With
     Next i
End With
If Cnt > 0 Then
   ReDim Preserve ResList(Cnt - 1)
Else
   Erase ResList
End If
GetUMCList = Cnt
Exit Function

err_GetUMCList:
GetUMCList = -1

End Function

Public Sub ReportUMC(ByVal lngGelIndex As Long, ByVal strUMCDefDescription As String)
'-------------------------------------------------------------------------------
'prints report in temporary file that can be saved as a semicolon delimited text
'-------------------------------------------------------------------------------
Dim FileNum As Integer
Dim FileNam As String
Dim sLine As String
Dim Stat() As Double
Dim lRows As Long
Dim i As Long
Dim RepInd As Long
Dim CSMOverZ As Double
On Error GoTo exit_cmdReport_Click
With GelData(lngGelIndex)
If GelUMC(lngGelIndex).UMCCnt > 0 Then
   lRows = UMCStatistics1(lngGelIndex, Stat())
   If lRows > 0 Then
      FileNum = FreeFile
      FileNam = GetTempFolder() & RawDataTmpFile
      Open FileNam For Output As FileNum
      'print gel file name and UMC definition as reference
      Print #FileNum, "Generated by: " & GetMyNameVersion
      Print #FileNum, strUMCDefDescription
      sLine = "UMC_ID" & glARG_SEP & "UMC_Cnt" & glARG_SEP & "UMC_MW" & glARG_SEP & "Scan1" _
            & glARG_SEP & "Scan2" & glARG_SEP & "UMC_Avg_MW" & glARG_SEP & "MW_StD" & glARG_SEP _
            & "UMC_Abu" & glARG_SEP & "UMC_Avg_Fit" & glARG_SEP & "Rep_Scan" & glARG_SEP _
            & "Rep_CS" & glARG_SEP & "Rep_MW" & glARG_SEP & "Rep_m/z" & glARG_SEP & "Rep_Abu"
      Print #FileNum, sLine
      For i = 0 To lRows - 1
          If Stat(i, 0) >= 0 Then
             sLine = Stat(i, 0) & glARG_SEP & GelUMC(lngGelIndex).UMCs(i).ClassCount _
                & glARG_SEP & GelUMC(lngGelIndex).UMCs(i).ClassMW & glARG_SEP & Stat(i, 2) _
                & glARG_SEP & Stat(i, 3) & glARG_SEP & Stat(i, 4) & glARG_SEP _
                & GelUMC(lngGelIndex).UMCs(i).ClassMWStD & glARG_SEP & Stat(i, 6) & glARG_SEP & Stat(i, 7)
             RepInd = GelUMC(lngGelIndex).UMCs(i).ClassRepInd
             Select Case GelUMC(lngGelIndex).UMCs(i).ClassRepType
             Case glCSType
                CSMOverZ = .CSData(RepInd).AverageMW / .CSData(RepInd).Charge + glMASS_CC
                sLine = sLine & glARG_SEP & .CSData(RepInd).ScanNumber & glARG_SEP _
                    & .CSData(RepInd).Charge & glARG_SEP & .CSData(RepInd).AverageMW _
                    & glARG_SEP & CSMOverZ & glARG_SEP & .CSData(RepInd).Abundance
             Case glIsoType
                sLine = sLine & glARG_SEP & .IsoData(RepInd).ScanNumber & glARG_SEP _
                    & .IsoData(RepInd).Charge & glARG_SEP & GetIsoMass(.IsoData(RepInd), GelUMC(lngGelIndex).def.MWField) _
                    & glARG_SEP & .IsoData(RepInd).MZ & glARG_SEP & .IsoData(RepInd).Abundance
             End Select
          Else
             sLine = "Error calculating statistic for class " & i
          End If
          Print #FileNum, sLine
      Next i
      Close FileNum
      DoEvents
      frmDataInfo.Tag = "UMC"
      frmDataInfo.Show vbModal
   Else
     MsgBox "Error generating report for LC-MS Features. Make sure that LC-MS Feature was at least once generated and try again.", vbOKOnly
   End If
Else
   MsgBox "No Unique Mass Classes found.", vbOKOnly
End If
End With

exit_cmdReport_Click:
End Sub


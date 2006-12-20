VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlGlobals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'last modified 02/12/2003 nt
'---------------------------------------------------------
Option Explicit

' Note: Update CSNUM_FIELD_COUNT if changing this enum
Public Enum glDocDataCSFields
    csfScan = 1             'scan number
    csfFirstCS = 2          'first charge state
    csfCSNum = 3            'number of charge states
    csfAbu = 4              'intensity(abundance)
    csfMW = 5               'molecular mass
    csfStD = 6              'standard deviation
    csfIsotopicFitRatio = 7         ' Fit from ICR-2LS based pairs finding; Legacy: matching DB mass(obsolete)
    csfIsotopicAtomCount = 8        ' Atom count from ICR-2LS based pairs finding; Legacy: error of matching DB mass
End Enum

' Note: Update CSVAR_FIELD_COUNT if changing this enum
Public Enum glDocDataCSVarFields
    csvfIsotopeLabel = 1    'AMT NET
    csvfMTDDRatio = 2       'When reading .Pek file, this is filled with isotopic ratio values determined by ICR-2LS (if they exist, aka DD ratio); Later, if pairs are searched for, then filled with Expression Ratio values by function FillUMC_ERs
    csvfMTID = 3            'MT tag ID description (could contain multiple ID's, of the form AMT:12345); can also contain NETLK: entries
End Enum

' Note: Update ISONUM_FIELD_COUNT if changing this enum
Public Enum glDocDataISFields
    isfScan = 1             'scan number
    isfCS = 2               'charge state
    isfAbu = 3              'intensity(abundance)
    isfMOverZ = 4           'm/z
    isfFit = 5              'calculated fit
    isfMWavg = 6            'average molecular mass
    isfMWMono = 7           'monoisotopic molecular mass
    isfMWTMA = 8            'the most abundant mol.mass
    isfIsotopicFitRatio = 9         ' Fit from ICR-2LS based pairs finding; Legacy: matching DB mass(obsolete)
    isfIsotopicAtomCount = 10       ' Atom count from ICR-2LS based pairs finding; Legacy: error of matching DB mass
    isfIReportMWMonoAbu = 11
    isfIReport2DaAbundance = 12
End Enum

' Note: Update ISOVAR_FIELD_COUNT if changing this enum
Public Enum glDocDataISVarFields
    ''isvfIsotopeLabel = 1    'Could be N14, N15, C12, C13, O16, or O18; Also could be blank.  In legacy .PEK files this held the NET value for the AMT matched by ICR2LS, and was thus called isvfMTNet; In legacy .PEK files this also could hold an Asterisk, indicating it was a second guess charge state entry according to ICR-2LS
    ''isvfMTDDRatio = 2       'When reading .Pek file, this is filled with isotopic ratio values determined by ICR-2LS (if they exist, aka DD ratio); Later, if pairs are searched for, then filled with Expression Ratio values by function FillUMC_ERs
    isvfMTID = 3            'MT tag ID description (could contain multiple ID's, of the form AMT:12345); can also contain NETLK: entries
End Enum

Public Enum glScope
    glSc_All = 0            'all data
    glSc_Current = 1        'current view data
End Enum

Public Enum glView
    glvNormal = 1
    glvDifferential = 2
    glvUMC = 3
End Enum

Public Enum glDistType    ' gel element types
    gldtCS = 1            ' charge state
    gldtIS = 2            ' isotopic
    gldtUMC = 3           ' unique mass class/class average
    gldtNone = -1
End Enum

Public Enum glMassToleranceConstants
    gltERR = -1       'unused
    gltPPM = 0        'in parts per million
    gltPct = 1        'in percent
    gltABS = 2        'absolute (in whatever)
    gltStd = 3        'in standard deviations
End Enum

Public Enum glPairsType
    ptNone = -1
    ptUMCDlt = 0                ' Delta mass between light and heavy members; variable number of deltas allowed
    ptUMCLbl = 2                ' Additional label (like ICAT )on light and heavy members, separated by delta mass; variable number of labels; fixed delta
    ptUMCDltLbl = 4             ' Additional label (like ICAT )on light and heavy members, separated by delta mass; variable number of labels; variable number of deltas
    ptS_Dlt = 1
    ptS_Lbl = 3
    ptS_DltLbl = 5
End Enum

Public Enum glETType
    etGenericNET = 0
    etTICFitNET = 1
    etGANET = 2
End Enum

Public Enum OverlayNETAdjustmentType
    olyNETFromMinMax = 0        'adjust based on specified min and max
    olyNETDB_TIC = 1            'adjust based on database (TIC fit)
    olyNETDB_GANET = 2          'adjust based on database (GANET)
    olyNETDisplay = 3           'adjust based on selected display
End Enum

Public Enum OverlayType
    olySolo = 0
    OlyUMC = 1
End Enum

Public Enum OverlayShape
    olyBox = 0
    olyBoxEmpty = 1
    olySpot = 2
    olySpotEmpty = 3
    olyStick = 4
    olyTriangle = 5
    olyTriangleEmpty = 6
    olyTriStar = 7
End Enum

Public Enum ScaleType
    stNone = 0
    stLinear = 1
    stLog = 2
    stLn = 3
    stExp10 = 4
    stExpe = 5
End Enum

Public Enum glAlignment
    algLeft = 0
    algCenter = 1
    algRight = 2
End Enum

' Note: MAX_FILTER_COUNT constant should be updated in MdiGraph.Bas (Module1)
'       if the following list grows to over 20 items
Public Enum glFilters
    fltDupTolerance = 1
    fltDBTolerance = 2
    fltIsoFit = 3
    fltCase2CloseResults = 4
    fltAR = 5
    fltID = 6
    fltCSAbu = 7
    fltIsoAbu = 8
    fltCSMW = 9
    fltIsoMW = 10
    fltIsoCS = 11
    fltCSStDev = 12
    fltIsoCom = 13                  'isotopic composition
    fltIsoMZ = 14                   ' filter by m/z
    fltEvenOddScanNumber = 15       ' if .DataFilter(fltEvenOddScanNumber, 2) = 1, show only odd scan numbers; if 2, show only even numbers
End Enum

Public Enum UMCDrawType
    umcdt_ActualUMC = 0
    umcdt_FullUMCRegion = 1
    umcdt_UMCIntensity = 2
End Enum

Public Enum ArrayManagementType
    amtErase = 0
    amtInitialize = 1
    amtAdd = 2
    amtTrim = 3
End Enum

Attribute VB_Name = "Module1"
'DATA STRUCTURES, LOADING DATA FROM TEXT FILES,
'Last modified 08/03/2004
'--------------------------------------------------------------------------
Option Explicit

Public Const CAL_EQUATION_1 = "m/z = A/f + B/f^2"
Public Const CAL_EQUATION_2 = "m/z=A/f+B/f^2+C/f^3"
Public Const CAL_EQUATION_3 = "m/z=A/f+B/f^2+CI/f^2"
Public Const CAL_EQUATION_4 = "m/z=A/f+B/f^2+C"
Public Const CAL_EQUATION_5 = "m/z=A/(f-B)"

' No longer supported (March 2006)
''Public Const glDBGEL_ORF = -1
''Public Const glDBGEL_AMT = -2
''Public Const glDBGEL_ERROR = -3

'following tags are applied to Comment field to keep track
'of additional parameters
Public Const glCOMMENT_DO_NOT_EDIT = "(DO NOT EDIT THIS LINE)"
Public Const glCOMMENT_CREATED = "Gel Created: "
Public Const glCOMMENT_USER = "User: "
Public Const glCOMMENT_MASS_TAG = "Mass Tag: "
Public Const glCOMMENT_DELTA = "Delta: "
Public Const glCOMMENT_DELTA_TOL = "Delta Tolerance: "
Public Const glCOMMENT_MAX_DELTAS = "Max Deltas: "
Public Const glCOMMENT_RT = "RT Base Time(sec.): "

Public Const glCOMMENT_WYS = "Created from WYS of gel: "

' No longer supported (March 2006)
''Public Const glCOMMENT_DBGEL = "Gel created from DB: "
''Public Const glCOMMENT_DBGEL_AMT = "AMT"
''Public Const glCOMMENT_DBGEL_ORF = "ORF"
Public Const glCOMMENT_MTGEL = "MT tags Display"

Public Const glCOMMENT_DATA_FILE_START = "Data File("

Public Const glLaV2DG_FREQUENCY_SHIFT = "LaV2DG Frequency Shift: "

Public Enum rfcRawFileConstants
    rfcUnknown = 0
    rfcZippedSFolders = 1
    rfcFinniganRaw = 2
End Enum

Public Enum dfmCSandIsoDataFilterModeConstants
    dfmLoadAllData = 0
    dfmLoadCSDataOnly = 1
    dfmLoadIsoDataOnly = 2
End Enum

Public Enum eosEvenOddScanFilterModeConstants
    eosLoadAllScans = 0
    eosLoadOddScansOnly = 1
    eosLoadEvenScansOnly = 2
End Enum

'global data types (this information is not saved with the .Gel file)
Public Type GelState
    Deleted As Integer
    Dirty As Integer
    
' No longer supported (March 2006)
''    DBGel As Long              '0 - not DB Gel; -1 ORF Gel
''                               '-2 AMT Gel; any positive number
''                               'is the number of Source
    UMC As Integer
    SourceDataRawFileType As rfcRawFileConstants
    FinniganRawFilePath As String
End Type

Public Type GelPrefs1999            'preferences type for glCERT1999
'SWITCHES
    IsoDataField As Integer         ' 6=Avg;7=Mono;8=Most Abundant
    Case2Results As Integer         ' 0=use both;1=eliminate asterisk(default);2=use more likely
    DRDefinition As Integer         ' 0 Repressed/Induced, 1 Induced/Repressed
    IsoICR2LSMOverZ As Boolean      ' If True use m/z when calling ICR2LS, if False use highest abundance MW
'TOLERANCES
    DBTolerance As Single           ' 0=only exact matches; -1 = all data
    DupTolerance As Single          ' 2 default; duplicate is everything closer than DupTolerance
    IsoDataFit As Single            ' 0.5 default, use only data with fit better than
'DRAWING
    MinPointFactor As Single        ' specifies the minimum point size if >0; othervise do not use it
    MaxPointFactor As Single        ' specifies the max point size in the FN coosys, MaxPointSize=MaxPointFactor*Width(FN(n+1)-FN(n))
    BorderClrSameAsInt As Boolean   ' true if border color should be same as interior color
    AbuAspectRatio As Single        ' abundance aspect ratio
'COORDINATE SYSTEM
    CooType As Integer              ' type of coordinate system
    CooOrigin As Integer            ' position of the origin of the coordinate system
    CooHOrientation As Integer      ' horizontal orientation of the coordinate system
    CooVOrientation As Integer      ' vertical orientation of the coordinate system
    CooVAxisScale As Integer        ' vertical axis scale - linear or logarithmic
End Type

Public Type GelPrefs    'preferences type for glCERT2003
'SWITCHES
    IsoDataField As Integer         ' 6=Avg;7=Mono;8=Most Abundant
    Case2Results As Integer         ' 0=use both;1=eliminate asterisk(default);2=use more likely
    DRDefinition As Integer         ' 0 Repressed/Induced, 1 Induced/Repressed
    IsoICR2LSMOverZ As Boolean      ' If True use m/z when calling ICR2LS, if False use highest abundance MW
'TOLERANCES
    DBTolerance As Double           ' 0=only exact matches; -1 = all data
    DupTolerance As Double          ' 2 default; duplicate is everything closer than DupTolerance
    IsoDataFit As Double            ' 0.5 default, use only data with fit better than
'DRAWING
    MinPointFactor As Double        ' specifies the minimum point size if >0; othervise do not use it
    MaxPointFactor As Double        ' specifies the max point size in the FN coosys, MaxPointSize=MaxPointFactor*Width(FN(n+1)-FN(n))
    BorderClrSameAsInt As Boolean   ' true if border color should be same as interior color
    AbuAspectRatio As Double        ' abundance aspect ratio
'COORDINATE SYSTEM
    CooType As Integer              ' type of coordinate system
    CooOrigin As Integer            ' position of the origin of the coordinate system
    CooHOrientation As Integer      ' horizontal orientation of the coordinate system
    CooVOrientation As Integer      ' vertical orientation of the coordinate system
    CooVAxisScale As Integer        ' vertical axis scale - linear or logarithmic
End Type

Public Type DocumentData1999            'file format for Certificate = glCERT1999
  Certificate As String
  Comment As String
  FileName As String
  Fileinfo As String
  PathtoDataFiles As String
  PathtoDatabase As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long
  MinMW As Single            'FN & PI extremes extract from DF arrays
  MaxMW As Single
  MinAbu As Single
  MaxAbu As Single
  Preferences As GelPrefs1999 'preferences that enable maximal customization of each file
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  DataFilter(1 To 8, 2) As Variant
  DFN() As String            'data file names
  DFFN() As Long             'data file file numbers
  DFPI() As Single           'data file PI number
  CSNum() As Single          'charge state numeric fields
  CSVar() As Variant         'charge state nonnumeric fields
  IsoNum() As Single         'isotopic numeric fields; 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To 10
  IsoVar() As Variant        'isotopic nonnumeric fields; 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To 3
End Type

Public Type DocumentData2000            'file format for Certificate = glCERT2000
  Certificate As String
  Comment As String
  FileName As String
  Fileinfo As String
  PathtoDataFiles As String
  PathtoDatabase As String
  MediaType As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long
  CalEquation As String
  CalArg(1 To 10) As Double  'arguments in calibration equation
  MinMW As Double            'FN & PI extremes extract from DF arrays
  MaxMW As Double
  MinAbu As Double
  MaxAbu As Double
  Preferences As GelPrefs    'preferences that enable maximal customization of each file
'this will have to change to something like Separation Coordinate System enabled
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  DataFilter(1 To 8, 2) As Variant
  DFN() As String            'data file names
  DFFN() As Long             'data file file numbers
  DFPI() As Double           'data file PI number
  DFFS() As Double           'frequency shifts
  DFIN() As Double           'intensity - MonroeMod: Storing Time Domain Signal Level here
  CSNum() As Double          'charge state numeric fields
  CSVar() As Variant         'charge state nonnumeric fields
  IsoNum() As Double         'isotopic numeric fields; 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To 10
  IsoVar() As Variant        'isotopic nonnumeric fields; 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To 3
End Type

Public Const MAX_FILTER_COUNT_2003 = 20
Public Type DocumentData2003            'file format for Certificate = glCERT2003 and glCERT2003_Modular with fioGelData = 2#
  Certificate As String
  Comment As String
  FileName As String
  Fileinfo As String
  PathtoDataFiles As String
  PathtoDatabase As String              ' Path to the legacy database used
  MediaType As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long             ' Number of isotopic data points:  Note, .IsoNum() is 1-based
  CalEquation As String
  CalArg(1 To 10) As Double  'arguments in calibration equation
  MinMW As Double            'FN & PI extremes extract from DF arrays
  MaxMW As Double
  MinAbu As Double
  MaxAbu As Double
  Preferences As GelPrefs    'preferences that enable maximal customization of each file
'this will have to change to something like Separation Coordinate System enabled
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  DataFilter(1 To MAX_FILTER_COUNT_2003, 2) As Variant             ' Second dimension is (x,0) = True to enable filter, (x,1) = filter value or filter minimum value, (x,2) = filter maximum value (if required)
  'three new filters added; MW range for Isotopic and Charge State filter
  'and charge state filter for Isotopic date; array memberes from 12 to MAX_FILTER_COUNT=20 are reserved for future use
  DFN() As String            'data file names; 1-based array
  DFFN() As Long             'data file scan numbers (aka file numbers); 1-based array; note that .ScanInfo(1).ScanNumber is the minimum scan number and DFFN(UBound(.ScanInfo)) is the maximum scan number
  DFPI() As Double           'data file PI number; 1-based array
  DFFS() As Double           'frequency shifts; 1-based array
  DFIN() As Double           'intensity - MonroeMod: Storing Time Domain Signal Level here; 1-based array
  CSNum() As Double          'charge state numeric fields               (raw, non-isotopic data, 1-based two-dimensional array, see Enum glDocDataCSFields for key to the second dimension fields)
  CSVar() As Variant         'charge state nonnumeric fields            (1-based two-dimensional array; see Enum glDocDataCSVarFields for key to the second dimension fields)
  IsoNum() As Double         'isotopic numeric fields                   (deconvoluted data, 1-based two-dimensional array, see Enum glDocDataISFields for key to the second dimension fields); 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To 10
  IsoVar() As Variant        'isotopic nonnumeric fields                (1-based two-dimensional array; see Enum glDocDataISVarFields for key to the second dimension fields); 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To 3
End Type

Public Const MAX_FILTER_COUNT_2003b = 20
Public Type DocumentData2003b            'file format for Certificate = glCERT2003_Modular with fioGelData = 3#
  Certificate As String
  Comment As String
  FileName As String            ' Holds the full path to the .Pek, or .CSV file that the data was loaded from; in previous versions this would get updated with the path to the .Gel file
  Fileinfo As String
  PathtoDataFiles As String     ' Holds the full path to the zipped s-folders for the data; this now defaults to the parent folder of the .Pek or .CSV file when loaded; in old versions this will point to a folder like C:\DMS_ICR_WorkDir2
  PathtoDatabase As String      ' Path to the legacy database used
  MediaType As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long             ' Number of isotopic data points:  Note, .IsoNum() is 1-based
  CalEquation As String
  CalArg(1 To 10) As Double  'arguments in calibration equation
  MinMW As Double            'FN & PI extremes extract from DF arrays
  MaxMW As Double
  MinAbu As Double
  MaxAbu As Double
  Preferences As GelPrefs    'preferences that enable maximal customization of each file
  'this will have to change to something like Separation Coordinate System enabled
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  
  DataStatusBits As Long     ' New for this version; DataStatusBits is used to hold various isotopic data indicator bits (thus, it is a bit bucket)
  
  DataFilter(1 To MAX_FILTER_COUNT_2003b, 2) As Variant             ' Second dimension is (x,0) = True to enable filter, (x,1) = filter value or filter minimum value, (x,2) = filter maximum value (if required)
                                                              ' Array memberes from 12 to MAX_FILTER_COUNT=20 are reserved for future use
  DFN() As String            'data file names; 1-based array
  DFFN() As Long             'data file scan numbers (aka file numbers); 1-based array; note that .ScanInfo(1).ScanNumber is the minimum scan number and DFFN(UBound(.ScanInfo)) is the maximum scan number
  DFPI() As Double           'data file PI number; 1-based array
  DFFS() As Double           'frequency shifts; 1-based array
  DFIN() As Double           'intensity - MonroeMod: Storing Time Domain Signal Level here; 1-based array
  CSNum() As Double          'charge state numeric fields               (raw, non-isotopic data, 1-based two-dimensional array, see Enum glDocDataCSFields for key to the second dimension fields)
  CSVar() As Variant         'charge state nonnumeric fields            (1-based two-dimensional array; see Enum glDocDataCSVarFields for key to the second dimension fields)
  IsoNum() As Double         'isotopic numeric fields                   (deconvoluted data, 1-based two-dimensional array, see Enum glDocDataISFields for key to the second dimension fields); 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To ISONUM_FIELD_COUNT
  IsoVar() As Variant        'isotopic nonnumeric fields                (1-based two-dimensional array; see Enum glDocDataISVarFields for key to the second dimension fields); 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To ISOVAR_FIELD_COUNT
  OtherInfo As String        ' New for this version
End Type

Public Type udtScanInfo2004Type
    ScanNumber As Long
    ElutionTime As Single
    ScanType As Integer     ' 1 = MS scan, 2 = MS/MS scan (not used in this program, but available if we wanted to track it)
End Type

Public Const MAX_FILTER_COUNT_2004 = 20
Public Type DocumentData2004         'file format for Certificate = glCERT2004_Modular with fioGelData = 4#
  Certificate As String
  Comment As String
  FileName As String            ' Holds the full path to the .Pek or .CSV file that the data was loaded from; in previous versions this would get updated with the path to the .Gel file
  Fileinfo As String
  PathtoDataFiles As String     ' Holds the full path to the folder containing the zipped s-folders or Finnigan .Raw file; this now defaults to the parent folder of the .Pek or .CSV file when loaded; in old versions this will point to a folder like C:\DMS_ICR_WorkDir2
                                ' Note that GelStatus().SourceDataRawFileType will get updated from rfcUnknown to rfcZippedSFolders or rfcFinniganRaw
  PathtoDatabase As String      ' Path to the legacy database used
  MediaType As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long             ' Number of isotopic data points:  Note, .IsoNum() is 1-based
  CalEquation As String
  CalArg(1 To 10) As Double  'arguments in calibration equation
  MinMW As Double            'FN & PI extremes extract from DF arrays
  MaxMW As Double
  MinAbu As Double
  MaxAbu As Double
  Preferences As GelPrefs    'preferences that enable maximal customization of each file
  'this will have to change to something like Separation Coordinate System enabled
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  
  DataStatusBits As Long     ' DataStatusBits is used to hold various isotopic data indicator bits (thus, it is a bit bucket)
  
  DataFilter(1 To MAX_FILTER_COUNT_2004, 2) As Variant             ' Second dimension is (x,0) = True to enable filter, (x,1) = filter value or filter minimum value, (x,2) = filter maximum value (if required)
                                                              ' Array memberes from 12 to MAX_FILTER_COUNT=20 are reserved for future use
  
  ScanInfo() As udtScanInfo2004Type    ' New for this version; lists all MS scan numbers and times, in increasing order; useful for instruments that alternate between MS scans and alternate scans; 1-based array in order to stay compatible with DFFN() since this array is parallel with .DFFN()
  
  DFN() As String            'data file names; 1-based array
  DFFN() As Long             'data file scan numbers (aka file numbers); 1-based array; note that .ScanInfo(1).ScanNumber is the minimum scan number and DFFN(UBound(.ScanInfo)) is the maximum scan number
  DFPI() As Double           'data file PI number; 1-based array
  DFFS() As Double           'frequency shifts; 1-based array
  DFIN() As Double           'intensity - MonroeMod: Storing Time Domain Signal Level here; 1-based array
  CSNum() As Double          'charge state numeric fields               (raw, non-isotopic data, 1-based two-dimensional array, see Enum glDocDataCSFields for key to the second dimension fields)
  CSVar() As Variant         'charge state nonnumeric fields            (1-based two-dimensional array; see Enum glDocDataCSVarFields for key to the second dimension fields)
  IsoNum() As Double         'isotopic numeric fields                   (deconvoluted data, 1-based two-dimensional array, see Enum glDocDataISFields for key to the second dimension fields); 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To ISONUM_FIELD_COUNT
  IsoVar() As Variant        'isotopic nonnumeric fields                (1-based two-dimensional array; see Enum glDocDataISVarFields for key to the second dimension fields); 1st Dimension is (1 To .IsoLines) and 2nd Dimension is 1 To ISOVAR_FIELD_COUNT
  OtherInfo As String
End Type

Public Type udtScanInfoType
    ScanNumber As Long          ' Scan number (used to be .DFFN())
    ElutionTime As Single       ' Elution time, in minutes; only present for certain instrument types
    ScanType As Integer         ' 1 = MS scan, 2 = MS/MS scan (not used in this program, but available if we wanted to track it)

    ScanFileName As String      ' Data file names (used to be in .DFN())
    ScanPI As Double            ' Data file PI number (used to be in .DFPI())

    NumIsotopicSignatures As Long   ' New for this version; used to generate chromatogram tbcDeisotopingPeakCounts: RawIntensity
    NumPeaks As Long                ' New for this version; used to generate chromatogram tbcDeisotopingPeakCounts: NormalizedIntensity
    TIC As Single                   ' New for this version; used to generate chromatogram tbcTICFromRawData
    BPI As Single                   ' New for this version; used to generate chromatogram tbcBPIFromRawData
    BPImz As Single                 ' New for this version
    
    TimeDomainSignal As Single              ' New for this version (used to be in .DFIN())
    
    PeakIntensityThreshold As Single        ' New for this version; Lower threshold filter, applies to isotopic peaks used in peak picking
    PeptideIntensityThreshold As Single     ' New for this version; Higher threshold filter, applies to the base peaks of the isotopic distributions
    
    FrequencyShift As Single                ' New for this version (used to be in .DFFS())
    
    CustomNET As Single             ' New for this version; custom NET value for each scan (initially 0 for all scans); was AdditionalValue1
    AdditionalValue2 As Single      ' New for this version; use for future expansion (name can be changed in the future)
    
End Type

' This is used with types DocumentData2005a and DocumentData2005b
Public Type udtIsotopicDataType2005
    ScanNumber As Long
    Charge As Integer               ' Charge for isotopic data; first charge state for CS data
    ChargeCount As Byte             ' Number of charge states; only used with CS data
    Abundance As Single
    
    MZ As Double                    ' Only used for Iso data
    Fit As Single
    
    MonoisotopicMW As Double        ' Only used with Iso data
    AverageMW As Double             ' Average MW for Iso data; Molecular mass for CS data
    MostAbundantMW As Double        ' Only used with Iso data
    MassStDev As Single             ' Only used for CS data
    
    IntensityMono As Single
    IntensityMonoPlus2 As Single
    
    FWHM As Single
    SignalToNoise As Single

    ExpressionRatio As Single
    MTID As String
End Type

Public Const MAX_FILTER_COUNT_2005 = 20
Public Type DocumentData2005a            'file format for Certificate = glCERT2004_Modular with fioGelData = 5#
  Certificate As String
  Comment As String
  FileName As String            ' Holds the full path to the .Pek or .CSV file that the data was loaded from; in previous versions this would get updated with the path to the .Gel file
  Fileinfo As String
  PathtoDataFiles As String     ' Holds the full path to the folder containing the zipped s-folders or Finnigan .Raw file; this now defaults to the parent folder of the .Pek or .CSV file when loaded; in old versions this will point to a folder like C:\DMS_ICR_WorkDir2
                                ' Note that GelStatus().SourceDataRawFileType will get updated from rfcUnknown to rfcZippedSFolders or rfcFinniganRaw
  PathtoDatabase As String      ' Path to the legacy database used
  MediaType As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long             ' Number of isotopic data points:  Note, .IsoNum() is 1-based
  CalEquation As String
  CalArg(1 To 10) As Double  'arguments in calibration equation
  MinMW As Double            'FN & PI extremes extracted from DF arrays
  MaxMW As Double
  MinAbu As Double
  MaxAbu As Double
  Preferences As GelPrefs    'preferences that enable maximal customization of each file
  
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  
  DataStatusBits As Long     ' DataStatusBits is used to hold various isotopic data indicator bits (thus, it is a bit bucket)
  
  DataFilter(1 To MAX_FILTER_COUNT_2005, 2) As Variant        ' Second dimension is (x,0) = True to enable filter, (x,1) = filter value or filter minimum value, (x,2) = filter maximum value (if required)
                                                              ' Array memberes from 12 to MAX_FILTER_COUNT=20 are reserved for future use
  
  ScanInfo() As udtScanInfoType         ' Updated for this version; lists all MS scan numbers and times, in increasing order; 1-based array for historical reasons

  CSData() As udtIsotopicDataType2005       ' New for this version
  IsoData() As udtIsotopicDataType2005      ' New for this version
  
  OtherInfo As String
End Type

Public Type DocumentData2005b            'file format for Certificate = glCERT2004_Modular with fioGelData = 6#
  Certificate As String
  Comment As String
  FileName As String            ' Holds the full path to the .Pek or .CSV file that the data was loaded from; in previous versions this would get updated with the path to the .Gel file
  Fileinfo As String
  PathtoDataFiles As String     ' Holds the full path to the folder containing the zipped s-folders or Finnigan .Raw file; this now defaults to the parent folder of the .Pek or .CSV file when loaded; in old versions this will point to a folder like C:\DMS_ICR_WorkDir2
                                ' Note that GelStatus().SourceDataRawFileType will get updated from rfcUnknown to rfcZippedSFolders or rfcFinniganRaw
  PathtoDatabase As String      ' Path to the legacy database used
  MediaType As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long             ' Number of isotopic data points:  Note, .IsoNum() is 1-based
  CalEquation As String
  CalArg(1 To 10) As Double  'arguments in calibration equation
  MinMW As Double            'FN & PI extremes extracted from DF arrays
  MaxMW As Double
  MinAbu As Double
  MaxAbu As Double
  Preferences As GelPrefs    'preferences that enable maximal customization of each file
  
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  
  DataStatusBits As Long     ' DataStatusBits is used to hold various isotopic data indicator bits (thus, it is a bit bucket)
  
  DataFilter(1 To MAX_FILTER_COUNT_2005, 2) As Variant          ' Second dimension is (x,0) = True to enable filter, (x,1) = filter value or filter minimum value, (x,2) = filter maximum value (if required)
                                                                ' Array memberes from 12 to MAX_FILTER_COUNT_2005=20 are reserved for future use
  
  CustomNETsDefined As Boolean          ' New for this version; True when a custom NET is defined
  ScanInfo() As udtScanInfoType         ' Updated for this version; lists all MS scan numbers and times, in increasing order; 1-based array for historical reasons

  CSData() As udtIsotopicDataType2005
  IsoData() As udtIsotopicDataType2005
  
  AdditionalValue1 As Long          ' New for this version; use for future expansion (name can be changed in the future)
  AdditionalValue2 As Long          ' New for this version; use for future expansion (name can be changed in the future)
  AdditionalValue3 As Single        ' New for this version; use for future expansion (name can be changed in the future)
  AdditionalValue4 As Single        ' New for this version; use for future expansion (name can be changed in the future)
  
  OtherInfo As String
End Type

' This is used with type DocumentData
Public Type udtIsotopicDataType
    ScanNumber As Long
    Charge As Integer               ' Charge for isotopic data; first charge state for CS data
    ChargeCount As Byte             ' Number of charge states; only used with CS data
    Abundance As Single
        
    MZ As Double                    ' Only used for Iso data
    Fit As Single
    
    MonoisotopicMW As Double        ' Only used with Iso data
    AverageMW As Double             ' Average MW for Iso data; Molecular mass for CS data
    MostAbundantMW As Double        ' Only used with Iso data
    MassStDev As Single             ' Only used for CS data
    
    MassShiftCount As Byte          ' New for this version; Number of mass corrections applied (maximum value is 255)
    MassShiftOverallPPM As Single   ' New for this version; overall mass correction applied to this point (in ppm); if multiple adjustments are applied, this will track the overall adjustment applied
    
    IntensityMono As Single
    IntensityMonoPlus2 As Single
    
    FWHM As Single
    SignalToNoise As Single

    ExpressionRatio As Single
    
    AdditionalValue1 As Single        ' New for this version; use for future expansion (name can be changed in the future)
    AdditionalValue2 As Single        ' New for this version; use for future expansion (name can be changed in the future)
    
    MTID As String                      ' List of MT tags and/or Internal Standards that match this data point

End Type

Public Const MAX_FILTER_COUNT = 20
Public Type DocumentData            'file format for Certificate = glCERT2004_Modular with fioGelData = 7# (current data format)
  Certificate As String
  Comment As String
  FileName As String            ' Holds the full path to the .Pek, .CSV, .mzXML, or .mzData file that the data was loaded from; in previous versions this would get updated with the path to the .Gel file
  Fileinfo As String
  PathtoDataFiles As String     ' Holds the full path to the folder containing the zipped s-folders or Finnigan .Raw file; this now defaults to the parent folder of the .Pek, .CSV, .mzXML, or .mzData file when loaded; in old versions this will point to a folder like C:\DMS_ICR_WorkDir2
                                ' Note that GelStatus().SourceDataRawFileType will get updated from rfcUnknown to rfcZippedSFolders or rfcFinniganRaw
  PathtoDatabase As String      ' Holds the path to the Legacy DB (Access DB) used for search (if appropriate)
  MediaType As String
  LinesRead As Long
  DataLines As Long
  CSLines As Long
  IsoLines As Long             ' Number of isotopic data points:  Note, .IsoNum() is 1-based
  CalEquation As String
  CalArg(1 To 10) As Double  'arguments in calibration equation
  MinMW As Double            'FN & PI extremes extracted from DF arrays
  MaxMW As Double
  MinAbu As Double
  MaxAbu As Double
  Preferences As GelPrefs    'preferences that enable maximal customization of each file
  
  pICooSysEnabled As Boolean 'True if pI numbers exists, False otherwise
  
  DataStatusBits As Long     ' DataStatusBits is used to hold various isotopic data indicator bits (thus, it is a bit bucket)
  
  DataFilter(1 To MAX_FILTER_COUNT, 2) As Variant             ' Second dimension is (x,0) = True to enable filter, (x,1) = filter value or filter minimum value, (x,2) = filter maximum value (if required)
                                                              ' Array memberes from 12 to MAX_FILTER_COUNT=20 are reserved for future use
  
  CustomNETsDefined As Boolean          ' True when a custom NET is defined
  ScanInfo() As udtScanInfoType         ' Updated for this version; lists all MS scan numbers and times, in increasing order; 1-based array for historical reasons

  CSData() As udtIsotopicDataType       ' Updated for this version
  IsoData() As udtIsotopicDataType      ' Updated for this version
  
  AdditionalValue1 As Long          ' Use for future expansion (name can be changed in the future)
  AdditionalValue2 As Long          ' Use for future expansion (name can be changed in the future)
  AdditionalValue3 As Single        ' Use for future expansion (name can be changed in the future)
  AdditionalValue4 As Single        ' Use for future expansion (name can be changed in the future)
  
  OtherInfo As String
End Type

Public Type DrawData               'for drawing we will use single precision
  CSVisible As Boolean
  CSCount As Long           'this is redundant but will help in coding
  CSID() As Long
  CSX() As Single
  CSY() As Single
  CSR() As Single
  CSER() As Single          ' Holds -1 if no ER
  CSERClr() As Long
  CSLogMW() As Single
  IsoVisible As Boolean
  IsoCount As Long          'this is redundant but will help in coding
  IsoID() As Long           ' Pointer to the index number of the spot in GelData().IsoNum(); this value is set to the equivalent negative number (i.e. 5 becomes -5) when the spot is being filtered out by any number of methods
  IsoX() As Single
  IsoY() As Single
  IsoR() As Single
  IsoER() As Single         ' Holds -1 if no ER
  IsoERClr() As Long        'has sense to keep Color and Log(MW)
  IsoLogMW() As Single      'instead of calculation it when neccessary
End Type

Public Type DrawUMC                'structure used to draw Unique Mass Classes
  DrawType As Long
  Visible As Boolean
  Count As Long
  ClassID() As Long         'used also as an flag for visibility (negative ClassID if out of scope)
  X1() As Long
  Y1() As Long
  x2() As Long
  Y2() As Long
End Type

'Unique Mass Class structure
'This corresponds to FileInfoVersions(fioGelUMC) version 1
Public Type udtUMCType2002
  ClassRepInd As Long           'index of class representative(in GelData)
  ClassRepType As Long          'type of class representative
  ClassCount As Long            'number of class elements(including representative)
  ClassMInd() As Long           'members of the UMC(including representative)         (0-based array)
  ClassMType() As Integer       'type of member
  ClassAbundance As Double      'abundance of class(based on definition)
                                'avg, sum or abu of representative
  ClassMW As Double             'average or class representative mass
  ClassMWStD As Double          'class standard deviation; -1 if not applicable
  ClassNET As Double            'normalized elution time for the class
End Type

'Unique Mass Classes structure
'This corresponds to FileInfoVersions(fioGelUMC) version 1
Public Type UMCListType2002
  def As UMCDefinition2002  'save definition of this count
  UMCCnt As Long            'number of classes
  UMCs() As udtUMCType2002     'actual classes                                           (0-based array)
End Type


'Unique Mass Class structure
'This corresponds to FileInfoVersions(fioGelUMC) version 2
Public Type udtUMCType2003a
  ClassRepInd As Long           'index of class representative(in GelData)
  ClassRepType As Long          'type of class representative
  ClassCount As Long            'number of class elements(including representative)
  ClassMInd() As Long           'members of the UMC(including representative)         (0-based array); index of the member in GelData
  ClassMType() As Integer       'type of member: gldtCS or gldtIS
  ClassAbundance As Double      'abundance of class(based on definition)
                                'avg, sum or abu of representative
  ClassMW As Double             'average or class representative mass
  ClassMWStD As Double          'class standard deviation; -1 if not applicable
  ClassStatusBits As Double     ' ClassStatusBits is used to hold various UMC Indicator Bits;
                                ' This double used to be called ClassNET, but it is was not being used, so it is now a bit bucket
                                
  'following data is neccessary for faster drawing of Unique Mass Classes
  MinScan As Long
  MaxScan As Long
  MinMW As Double
  MaxMW As Double
End Type

'Unique Mass Classes structure
'This corresponds to FileInfoVersions(fioGelUMC) version 2
Public Type UMCListType2003a
  def As UMCDefinition2003a ' save definition of this count
  UMCCnt As Long            ' number of classes
  UMCs() As udtUMCType2003a    ' actual classes                                           (0-based array)
End Type

'Unique Mass Class structure (sub type)
Public Type UMClassChargeStateBasedStatsType
  Charge As Integer
  Count As Long             ' Number of UMC members with this charge
  Mass As Double            ' MW (based on def)
  MassStD As Double         ' MW Standard Deviation (based on def)
  Abundance As Double       ' Abundance (based on def)
  GroupRepIndex As Long     ' Index of the group representative; pointer to .ClassMInd() and .ClassMType() arrays, NOT a direct pointer to GelData
  OtherInfo As String
End Type

'Unique Mass Class structure (main type)
'This corresponds to FileInfoVersions(fioGelUMC) version 3
Public Type udtUMCType2004
  ClassRepInd As Long           ' index of class representative(in GelData)
  ClassRepType As Long          ' type of class representative
  ClassCount As Long            ' number of class elements(including representative)
  ClassMInd() As Long           ' members of the UMC(including representative)         (0-based array); index of the member in GelData
  ClassMType() As Integer       ' type of member: gldtCS or gldtIS
  ClassAbundance As Double      ' abundance of class(based on definition)
                                ' avg, sum or abu of representative
  ClassMW As Double             ' average, class representative mass, or median mass (based on definition)
  ClassMWStD As Double          ' class standard deviation; -1 if not applicable
  ClassScore As Double          ' Unused; Reserve for future expansion
  ClassNET As Double            ' Unused; currently using NET of the class representative
  ClassStatusBits As Double     ' ClassStatusBits is used to hold various UMC Indicator Bits (thus, it is a bit bucket)
                                
  'Following data is useful for faster drawing of Unique Mass Classes
  'It is also used in CalcDltLblPairsERWork
  'It is populated in CalculateClasses
  MinScan As Long               ' Minimum scan number in class (across all members)
  MaxScan As Long               ' Maximum scan number in class (across all members)
  MinMW As Double
  MaxMW As Double
  
  ' The following is derived information based on the members of the UMC, grouped by charge state
  ChargeStateStatsRepInd As Integer                                 ' The charge state group that best represents this UMC (determined using UMCDef.ChargeStateStatsRepType); pointer to entry in .ChargeStateBasedStats()
  ChargeStateCount As Integer                                       ' Number of unique charge states that the members of this UMC have
  ChargeStateBasedStats() As UMClassChargeStateBasedStatsType       ' Stats on the members, grouped by charge state  (0-based array)
  
  OtherInfo As String
End Type

'Unique Mass Classes structure
'This corresponds to FileInfoVersions(fioGelUMC) version 3
Public Type UMCListType2004
  def As UMCDefinition          'save definition of this count
  UMCCnt As Long                'number of classes
  UMCs() As udtUMCType2004      'actual classes                                           (0-based array)
End Type


'Unique Mass Class structure (main type)
'This corresponds to FileInfoVersions(fioGelUMC) version 4
' Note: including class representative in class members is
'       redundant but makes drawing much faster and simpler
Public Type udtUMCType
  ClassRepInd As Long           ' index of class representative(in GelData)
  ClassRepType As Long          ' type of class representative
  ClassCount As Long            ' number of class elements(including representative)
  ClassMInd() As Long           ' members of the UMC(including representative)         (0-based array); index of the member in GelData
  ClassMType() As Integer       ' type of member: gldtCS or gldtIS
  ClassAbundance As Double      ' abundance of class(based on definition)
                                ' avg, sum or abu of representative
  ClassMW As Double             ' average, class representative mass, or median mass (based on definition)
  ClassMWStD As Double          ' class standard deviation; -1 if not applicable
  ClassMassCorrectionDa As Double   ' New for this version; Positive or negative value to add to ClassMW to fine-tune the mass
  
  ClassScore As Double          ' Unused; Reserve for future expansion
  ClassNET As Double            ' Unused; currently using NET of the class representative
  ClassStatusBits As Double     ' ClassStatusBits is used to hold various UMC Indicator Bits (thus, it is a bit bucket)
                                
  'Following data is useful for faster drawing of Unique Mass Classes
  'It is also used in CalcDltLblPairsERWork
  'It is populated in CalculateClasses
  MinScan As Long               ' Minimum scan number in class (across all members)
  MaxScan As Long               ' Maximum scan number in class (across all members)
  MinMW As Double
  MaxMW As Double
  
  ' The following is derived information based on the members of the UMC, grouped by charge state
  ChargeStateStatsRepInd As Integer                                 ' The charge state group that best represents this UMC (determined using UMCDef.ChargeStateStatsRepType); pointer to entry in .ChargeStateBasedStats()
  ChargeStateCount As Integer                                       ' Number of unique charge states that the members of this UMC have
  ChargeStateBasedStats() As UMClassChargeStateBasedStatsType       ' Stats on the members, grouped by charge state  (0-based array)
  
  AdditionalValue1 As Long          ' New for this version; use for future expansion (name can be changed in the future)
  AdditionalValue2 As Long          ' New for this version; use for future expansion (name can be changed in the future)
  AdditionalValue3 As Single        ' New for this version; use for future expansion (name can be changed in the future)
  AdditionalValue4 As Single        ' New for this version; use for future expansion (name can be changed in the future)
  
  OtherInfo As String
End Type

'Unique Mass Classes structure
'This corresponds to FileInfoVersions(fioGelUMC) version 4
Public Type UMCListType
  def As UMCDefinition          'save definition of this count
  UMCCnt As Long                'number of classes
  UMCs() As udtUMCType          'actual classes                                           (0-based array)
  
  MassCorrectionValuesDefined As Boolean        ' New for this version; True when custom mass correction values are defined for the UMCs
  
  AdditionalValue1 As Long          ' New for this version; use for future expansion (name can be changed in the future)
  AdditionalValue2 As Single        ' New for this version; use for future expansion (name can be changed in the future)
  
  OtherInfo As String
End Type


' Structure for UMCIonNet UMC Searching go here (was Public Type NET; now Public Type UMCIonNet)
Public Type UMCIonNet
    Visible As Boolean
    ThisNetDef As UMCIonNetDefinition
    NetCount As Long
    NetInd1() As Long
    NetInd2() As Long
    NetDist() As Double         'distance between nodes
    MinDist As Double           'keep distance range; could come handy
    MaxDist As Double
End Type


'Unused type (July 2004)
'Pairs for isotopic labeling analysis
'Public Type IsoPairs
'   def As PairDefinition
'   PCnt As Long       'count of pairs
'   P1() As Long       'index of first in pair
'   P2() As Long       'index of second in pair
'End Type

Public Type ScansIndex                  'used to enumerate scans;
    ScansCnt As Long                    'keeps indexes of first and last index
    Scans() As Long                 'in Data arrays for charge state and
    CSFirstInd() As Long            'isotopic data
    CSLastInd() As Long
    IsoFirstInd() As Long
    IsoLastInd() As Long
End Type

' Old structure, now split into IsoPairsDltLblType and udtIsoPairsDetailsType
Public Type IsoPairsDltLbl2003Type
   SyncWithUMC As Boolean
   DltLblType As Long
   lblType As Long                              ' August 2003: This isn't used; instead, DltLblType is used for both deltas and labels
   LblMW As Double
   DltType As Long                              ' August 2003: This isn't used; instead, DltLblType is used for both deltas and labels
   DltMW As Double
   ERCalcType As Long                               ' Actually type ectERCalcTypeConstants, though can also be glER_None = 0
   PCnt As Long             'count of pairs
   P1() As Long             'index of light member; pointer to index in GelUMC if UMC-based pairs
   P1LblCnt() As Integer    'number of PEO labels in light member
   P2() As Long             'index of heavy member; pointer to index in GelUMC if UMC-based pairs
   P2DltCnt() As Integer    'count of N deltas
   P2LblCnt() As Integer    'count of PEO labels in heavy member
   P1P2ER() As Double       'expression ratio
   PState() As Integer      '3 states: glPAIR_Exc = -1, glPAIR_Neu = 0 , glPAIR_Inc = 1
    '1 for OK pair; -1 for not OK pair; 0 for initialized
    'Pair can be declared not OK for various reasons; because delta
    'and label counts do not match database information, because ER
    'is out of required range because pair assignment conflicts with
    'other pair assignments
End Type

' Old structure
Public Type udtIsoPairsDetails2004aType
    P1 As Long                      'index of light member; pointer to index in GelUMC if UMC-based pairs
    P1LblCnt As Integer             'number of PEO labels in light member
    P2 As Long                      'index of heavy member; pointer to index in GelUMC if UMC-based pairs
    P2DltCnt As Integer             'count of N deltas
    P2LblCnt As Integer             'count of PEO labels in heavy member
    ER As Double                    'expression ratio
    ERStDev As Double               'standard deviation of expression ratio, if averaging scan-by-scan or averaging several charge states
    ERChargeStateBasisCount As Integer      'count of number of charge states averaged together
    ERMemberBasisCount As Long              'count of number of values averaged together to give ER
    STATE As Integer                '3 states: glPAIR_Exc = -1, glPAIR_Neu = 0 , glPAIR_Inc = 1
End Type

' Old structure
Public Type IsoPairsDltLbl2004aType
    SyncWithUMC As Boolean
    DltLblType As Long                           ' Actually enum glPairsType
    lblType As Long                              ' August 2003: This isn't used; instead, DltLblType is used for both deltas and labels
    LblMW As Double
    DltType As Long                              ' August 2003: This isn't used; instead, DltLblType is used for both deltas and labels
    DltMW As Double
    ERCalcType As Long                           ' Actually enum ectERCalcTypeConstants, though can also be glER_None = 0
    
    RequireMatchingChargeStatesForPairMembers As Boolean
    UseIdenticalChargesForER As Boolean
    ComputeERScanByScan As Boolean
    AverageERsAllChargeStates As Boolean
    AverageERsWeightingMode As Integer          ' Actually enum aewAverageERsWeightingModeConstants
    
    PCnt As Long                                'count of pairs
    Pairs() As udtIsoPairsDetails2004aType      ' 0-based array
       
    OtherInfo As String
End Type

Public Type udtIsoPairsDetailsType
    P1 As Long                      'index of light member; pointer to index in GelUMC if UMC-based pairs
    P1LblCnt As Integer             'number of PEO labels in light member
    P2 As Long                      'index of heavy member; pointer to index in GelUMC if UMC-based pairs
    P2DltCnt As Integer             'count of N deltas
    P2LblCnt As Integer             'count of PEO labels in heavy member
    ER As Double                    'expression ratio
    ERStDev As Double               'standard deviation of expression ratio, if averaging scan-by-scan or averaging several charge states
    ERChargeStateBasisCount As Integer      'count of number of charge states averaged together
    ERChargesUsed() As Integer      'list of charge states used to compute the ER value for this pair; 0-based array, minimum length 1; value is 0 if no charge states used
    ERMemberBasisCount As Long              'count of number of values averaged together to give ER
    STATE As Integer                '3 states: glPAIR_Exc = -1, glPAIR_Neu = 0 , glPAIR_Inc = 1
    '1 for OK pair; -1 for not OK pair; 0 for initialized
    'Pair can be declared not OK for various reasons; because delta
    'and label counts do not match database information, because ER
    'is out of required range because pair assignment conflicts with
    'other pair assignments
End Type

' Old structure
Public Type IsoPairsDltLbl2004bType
    SyncWithUMC As Boolean
    DltLblType As Long                           ' Actually enum glPairsType
    lblType As Long                              ' August 2003: This isn't used; instead, DltLblType is used for both deltas and labels
    LblMW As Double
    DltType As Long                              ' August 2003: This isn't used; instead, DltLblType is used for both deltas and labels
    DltMW As Double
    ERCalcType As Long                           ' Actually enum ectERCalcTypeConstants, though can also be glER_None = 0

    RequireMatchingChargeStatesForPairMembers As Boolean
    UseIdenticalChargesForER As Boolean
    ComputeERScanByScan As Boolean
    AverageERsAllChargeStates As Boolean
    AverageERsWeightingMode As Integer          ' Actually enum aewAverageERsWeightingModeConstants

    PCnt As Long                                'count of pairs
    Pairs() As udtIsoPairsDetailsType           ' 0-based array

    OtherInfo As String
End Type

' Old structure; used by IsoPairsDltLbl2004cType
Public Type udtIsoPairsSearchDef2004cType
    DeltaMass As Double                         ' Typically glN14N15_DELTA or glO16O18_DELTA
    DeltaMassTolerance As Double                ' on frmUMCLblPairs this is actually the label mass tolerance
    AutoCalculateDeltaMinMaxCount As Boolean
    DeltaCountMin As Long
    DeltaCountMax As Long

    LightLabelMass As Double
    HeavyLightMassDifference As Double
    LabelCountMin As Long
    LabelCountMax As Long
    MaxDifferenceInNumberOfLightHeavyLabels As Long

    RequireUMCOverlap As Boolean                    ' Require overlap at UMC edges
    RequireUMCOverlapAtApex As Boolean              ' Require overlap at peak apex

    ScanTolerance As Long                           ' Scan tolerance at UMC edges
    ScanToleranceAtApex As Long                     ' Scan tolerance between UMC apexes

    ERInclusionMin As Double
    ERInclusionMax As Double

    RequireMatchingChargeStatesForPairMembers As Boolean
    UseIdenticalChargesForER As Boolean                         ' If UseIdenticalChargesForER = True, but RequireMatchingChargeStatesForPairMembers = False, and matching charges cannot be found, then the ER is computed using the ratio of the most abundant charge state for the members of the pair
    ComputeERScanByScan As Boolean                              ' When true, then computes an ER value for pairwise between the two UMC's of a pair, stepping scan by scan, then averaging the values across all scans; if UseIdenticalChargesForER = True then does this for matching charge states; otherwise, sums all charge states together
    AverageERsAllChargeStates As Boolean                        ' When true, then use a (weighted) average to combine the ER's for all matching charge states; this option is only valid if UseIdenticalChargesForER = True
    AverageERsWeightingMode As Integer                          ' Actually enum aewAverageERsWeightingModeConstants; The weighting mode to use if AverageERsAllChargeStates = True

    ERCalcType As Integer                                       ' Actually enum ectERCalcTypeConstants, though can also be glER_None = 0; how to calculate expression ratio
End Type

' Old structure
Public Type IsoPairsDltLbl2004cType
    SyncWithUMC As Boolean                      ' True if the pairs are sync'd with the UMCs in GelUMC()
    DltLblType As Long                          ' Actually enum glPairsType; ptNone, ptUMCDlt, ptUMCLbl, etc.

    SearchDef As udtIsoPairsSearchDef2004cType

    PCnt As Long                                ' Count of pairs
    Pairs() As udtIsoPairsDetailsType           ' 0-based array

    OtherInfo As String
End Type


Public Type udtIReportAbuRatioCoefficientsType
    Multiplier As Double
    Exponent As Double
End Type

Public Type udtIReportPairOptionsType
    Enabled As Boolean
    NaturalAbundanceRatio2Coeff As udtIReportAbuRatioCoefficientsType           ' Coefficients for the M2 / M0 ratio, where ratio = Multiplier * Mass ^ Exponent
    NaturalAbundanceRatio4Coeff As udtIReportAbuRatioCoefficientsType           ' Coefficients for the M4 / M0 ratio, where ratio = Multiplier * Mass ^ Exponent
    
    MinimumFractionScansWithValidER As Single                                   ' Minimum fraction of valid scans that must have a valid ER value
    
    OtherInfo As String
End Type
        

' Old structure; used by IsoPairsDltLbl2004dType
Public Type udtIsoPairsSearchDef2004dType
    DeltaMass As Double                         ' Typically glN14N15_DELTA or glO16O18_DELTA
    DeltaMassTolerance As Double                ' on frmUMCLblPairs this is actually the label mass tolerance
    AutoCalculateDeltaMinMaxCount As Boolean
    DeltaCountMin As Long
    DeltaCountMax As Long
    
    LightLabelMass As Double
    HeavyLightMassDifference As Double
    LabelCountMin As Long
    LabelCountMax As Long
    MaxDifferenceInNumberOfLightHeavyLabels As Long

    RequireUMCOverlap As Boolean                    ' Require overlap at UMC edges
    RequireUMCOverlapAtApex As Boolean              ' Require overlap at peak apex
    
    ScanTolerance As Long                           ' Scan tolerance at UMC edges
    ScanToleranceAtApex As Long                     ' Scan tolerance between UMC apexes

    ERInclusionMin As Double
    ERInclusionMax As Double
    
    RequireMatchingChargeStatesForPairMembers As Boolean
    UseIdenticalChargesForER As Boolean                         ' If UseIdenticalChargesForER = True, but RequireMatchingChargeStatesForPairMembers = False, and matching charges cannot be found, then the ER is computed using the ratio of the most abundant charge state for the members of the pair
    ComputeERScanByScan As Boolean                              ' When true, then computes an ER value for pairwise between the two UMC's of a pair, stepping scan by scan, then averaging the values across all scans; if UseIdenticalChargesForER = True then does this for matching charge states; otherwise, sums all charge states together
    AverageERsAllChargeStates As Boolean                        ' When true, then use a (weighted) average to combine the ER's for all matching charge states; this option is only valid if UseIdenticalChargesForER = True
    AverageERsWeightingMode As Integer                          ' Actually enum aewAverageERsWeightingModeConstants; The weighting mode to use if AverageERsAllChargeStates = True
    
    ERCalcType As Integer                                       ' Actually enum ectERCalcTypeConstants, though can also be glER_None = 0; how to calculate expression ratio
    
    IReportEROptions As udtIReportPairOptionsType           ' New for this version
    
    RemoveOutlierERs As Boolean                             ' New for this version
    RemoveOutlierERsIterate As Boolean                      ' New for this version
    RemoveOutlierERsMinimumDataPointCount As Long           ' New for this version
    RemoveOutlierERsConfidenceLevel As Integer              ' New for this version; actually of type eclConfidenceLevelConstants
    
    OtherInfo As String
End Type

' Old structure
Public Type IsoPairsDltLbl2004dType
    SyncWithUMC As Boolean                      ' True if the pairs are sync'd with the UMCs in GelUMC()
    DltLblType As Long                          ' Actually enum glPairsType; ptNone, ptUMCDlt, ptUMCLbl, etc.
    
    SearchDef As udtIsoPairsSearchDef2004dType
    
    PCnt As Long                                ' Count of pairs
    Pairs() As udtIsoPairsDetailsType           ' 0-based array
    
    OtherInfo As String
End Type

' Used by IsoPairsDltLblType
Public Type udtIsoPairsSearchDefType
    DeltaMass As Double                         ' Typically glN14N15_DELTA or glO16O18_DELTA; note: see also .DeltaMass2 through .DeltaMass6 (defined below)
    DeltaMassTolerance As Double                ' on frmUMCLblPairs this is actually the label mass tolerance
    AutoCalculateDeltaMinMaxCount As Boolean
    DeltaCountMin As Long
    DeltaCountMax As Long
    DeltaStepSize As Long                       ' New for this version
    
    LightLabelMass As Double
    HeavyLightMassDifference As Double
    LabelCountMin As Long
    LabelCountMax As Long
    MaxDifferenceInNumberOfLightHeavyLabels As Long

    RequireUMCOverlap As Boolean                    ' Require overlap at UMC edges
    RequireUMCOverlapAtApex As Boolean              ' Require overlap at peak apex
    
    ScanTolerance As Long                           ' Scan tolerance at UMC edges
    ScanToleranceAtApex As Long                     ' Scan tolerance between UMC apexes

    ERInclusionMin As Double
    ERInclusionMax As Double
    
    RequireMatchingChargeStatesForPairMembers As Boolean
    UseIdenticalChargesForER As Boolean                         ' If UseIdenticalChargesForER = True, but RequireMatchingChargeStatesForPairMembers = False, and matching charges cannot be found, then the ER is computed using the ratio of the most abundant charge state for the members of the pair
    ComputeERScanByScan As Boolean                              ' When true, then computes an ER value for pairwise between the two UMC's of a pair, stepping scan by scan, then averaging the values across all scans; if UseIdenticalChargesForER = True then does this for matching charge states; otherwise, sums all charge states together
    AverageERsAllChargeStates As Boolean                        ' When true, then use a (weighted) average to combine the ER's for all matching charge states; this option is only valid if UseIdenticalChargesForER = True
    AverageERsWeightingMode As Integer                          ' Actually enum aewAverageERsWeightingModeConstants; The weighting mode to use if AverageERsAllChargeStates = True
    
    ERCalcType As Integer                                       ' Actually enum ectERCalcTypeConstants, though can also be glER_None = 0; how to calculate expression ratio
    
    IReportEROptions As udtIReportPairOptionsType
    
    RemoveOutlierERs As Boolean
    RemoveOutlierERsIterate As Boolean
    RemoveOutlierERsMinimumDataPointCount As Long
    RemoveOutlierERsConfidenceLevel As Integer              ' Actually of type eclConfidenceLevelConstants
    
    OtherInfo As String
End Type

'structure for isotopic pairs that use delta and label
'this is used for PEO N14/N15 pairs as well as ICAT pairs
'also structure works for individual as well as UMCListType pairs
'In case of individual peak pairs consideration is limited
'to Isotopic peaks
Public Type IsoPairsDltLblType
    SyncWithUMC As Boolean                      ' True if the pairs are sync'd with the UMCs in GelUMC()
    DltLblType As Long                          ' Actually enum glPairsType; ptNone, ptUMCDlt, ptUMCLbl, etc.
    
    SearchDef As udtIsoPairsSearchDefType
    
    PCnt As Long                                ' Count of pairs
    Pairs() As udtIsoPairsDetailsType           ' 0-based array
    
    OtherInfo As String
End Type

'Delta structure (used to read Expression ratios)
'this will be parallel to FN arrays in DocumentData structure
Public Type DD
   Delta As Double          'sought Delta
   Tolerance As Double      'tolerance
   TagMass As Double        '?
   MaxDeltas As Long        'max number of deltas per pair
   MinInd As Long           'minimum index (used to mark position of Scan records among all deltas)
   MaxInd As Long           'maximum index
End Type


'following structure is parallel to the GelData
Public Type LMDataWorking
    Locked As Boolean           'true if locking attempted
    MWCnt As Long               'count of distributions
    MWID() As Long              'index in CS/Iso arrays
    MWType() As Integer         'charge state or isotopic
    MWFN() As Integer           'scan number
    MWCS() As Integer           'charge state
    MWLM() As Double            'locked mass(or original if not locked)
    MWInd() As Long             'index of these arrays (they will be sorted)
    MWLckID() As Long           'locker ID (MT tag database)
    MWFreqShift() As Double     'frequency shift applied
    MWMassCorrection() As Double
End Type

Public Type LMDataResults
    CSCnt As Long
    CSLckID() As Long
    CSFreqShift() As Double
    CSMassCorrection() As Double
    IsoCnt As Long
    IsoLckID() As Long
    IsoFreqShift() As Double
    IsoMassCorrection() As Double
End Type

'Unused type (August 2003)
''next structure is used to transfer pairs identification information
'Public Type ID_UMC_Pairs
'    Cnt As Long                         'count of identifications
'    SyncWithDltLblPairs As Boolean      'True if synchronized with DltLblPairs
'    PInd() As Long                      'index of pair
'    PIDInd() As Long                    'index in MT tags array of identification
'End Type

'groups structure used with UMC editing function
Public Type GR                  'group of longs
    Count As Long
    Description As String
    Members() As Long
End Type

Public Type GR2                 'group of group of longs
    Count As Long
    Description As String
    Members() As GR
End Type

'0'th gel is reserved for loaded MT tags
Public GelBody() As New frmGraph       ' Array of child forms containing actual drawings
Public GelData() As DocumentData       ' Array of data structure, matches the GEL file structure
Public GelDraw() As DrawData           ' this should be inside the frmGraph but VB does not
                                       ' allow us to put public arrays inside the objects
Public GelStatus() As GelState         ' status of the object
Public GelUMC() As UMCListType         ' saved unique molecule mass classes
Public GelUMCIon() As UMCIonNet        ' UMCIon UMC searching definition and data

'Unused variable (July 2004)
'Public GelP() As IsoPairs                   ' saved pairs; used for isotopic labeling analysis  (seem to be unused, January 2004)

Public GelP_D_L() As IsoPairsDltLblType      ' all pairs, including delta labeled pairs

'No longer supported (March 2006)
'Public GelDB() As Database                  ' used only with DB gels

'Unused variable (August 2003)
'Public GelIDP() As ID_UMC_Pairs        ' used to communicate pairs ID info among functions



'analysis structure parallel with other Gel structures
'nothing for gels not loaded from MT tag database
Public GelAnalysis() As FTICRAnalysis
Public GelLM() As LMDataResults
Public GelUMCDraw() As DrawUMC

Public Sub Display0()
'-----------------------------------------------------------------------
'displays loaded MT tags as a 2D display
'LoadType glCSType - load them as Charge State data; glIsoType as IS
'-----------------------------------------------------------------------
Dim i As Long
Dim ScanRange As Long
'for practical reasons use only NETs between 0 and 1(not included)
Dim PracMinNET As Double, PracMaxNET As Double
Dim PracMinScan As Long, PracMaxScan As Long
On Error GoTo Display0ErrorHandler

Screen.MousePointer = vbHourglass
ScanRange = Display0MaxScan - Display0MinScan + 1
If ScanRange <= 0 Then
   MsgBox "Correct scan range boundaries on Options form and try again.", vbOKOnly, glFGTU
   Exit Sub
End If
Set GelAnalysis(0) = New FTICRAnalysis
With GelData(0)
     .Comment = glCOMMENT_CREATED & Now & vbCrLf & glCOMMENT_USER & UserName & vbCrLf & glCOMMENT_MTGEL
     Call Display0DataSource(AMTGeneration)
     .Certificate = glCERT2003_Modular
     .pICooSysEnabled = False
     .PathtoDatabase = ""
     ResetDataFilters 0, glPreferences
     
     ReDim .ScanInfo(ScanRange)
     For i = Display0MinScan To Display0MaxScan
         With .ScanInfo(i)
            .ScanNumber = i
            .ScanFileName = "MTD" & Format$(i, "00000")
            .ElutionTime = i / Display0MaxScan
        End With
     Next i
     .MaxAbu = 0:                       .MinAbu = glHugeOverExp
     .MinMW = glHugeOverExp:            .MaxMW = 0
     PracMinNET = glHugeOverExp:        PracMaxNET = -glHugeOverExp
     Select Case PresentDisplay0Type
     Case glCSType
        .CSLines = AMTCnt
        ReDim .CSData(.CSLines)
        For i = 1 To AMTCnt
            With .CSData(i)
                .ScanNumber = Display0MinScan + CLng(AMTData(i).NET * ScanRange)
                .Charge = 1
                .ChargeCount = 1
                .Abundance = 1000000
                .AverageMW = AMTData(i).MW
                .ExpressionRatio = 0                ' Legacy: stored AMTData(i).NET here
                .MTID = "MT: " & AMTData(i).ID
            End With
            If AMTData(i).MW < .MinMW Then .MinMW = AMTData(i).MW
            If AMTData(i).MW > .MaxMW Then .MaxMW = AMTData(i).MW
            If .CSData(i).Abundance < .MinAbu Then .MinAbu = .CSData(i).Abundance
            If .CSData(i).Abundance > .MaxAbu Then .MaxAbu = .CSData(i).Abundance
            If AMTData(i).NET > 0 And AMTData(i).NET < 1 Then
               If AMTData(i).NET < PracMinNET Then
                  PracMinNET = AMTData(i).NET:   PracMinScan = .CSData(i).ScanNumber
               End If
               If AMTData(i).NET > PracMaxNET Then
                  PracMaxNET = AMTData(i).NET:   PracMaxScan = .CSData(i).ScanNumber
               End If
            End If
        Next i
     Case glIsoType
        .IsoLines = AMTCnt
        ReDim .IsoData(.IsoLines)
        For i = 1 To AMTCnt
            
            With .IsoData(i)
                .ScanNumber = Display0MaxScan + CLng(AMTData(i).NET * ScanRange)
                .Charge = 1
                .MZ = AMTData(i).MW / .Charge + glMASS_CC
                .Abundance = 1000000
                .MonoisotopicMW = AMTData(i).MW
                .AverageMW = AMTData(i).MW
                .MostAbundantMW = AMTData(i).MW
                .ExpressionRatio = 0                       ' Legacy: stored AMTData(i).NET here
                .MTID = "MT: " & AMTData(i).ID
            End With
            
            If AMTData(i).MW < .MinMW Then .MinMW = AMTData(i).MW
            If AMTData(i).MW > .MaxMW Then .MaxMW = AMTData(i).MW
            If .IsoData(i).Abundance < .MinAbu Then .MinAbu = .IsoData(i).Abundance
            If .IsoData(i).Abundance > .MaxAbu Then .MaxAbu = .IsoData(i).Abundance
            If AMTData(i).NET > 0 And AMTData(i).NET < 1 Then
               If AMTData(i).NET < PracMinNET Then
                  PracMinNET = AMTData(i).NET:   PracMinScan = .IsoData(i).ScanNumber
               End If
               If AMTData(i).NET > PracMaxNET Then
                  PracMaxNET = AMTData(i).NET:   PracMaxScan = .IsoData(i).ScanNumber
               End If
            End If
        Next i
     End Select
     .DataFilter(fltCSAbu, 2) = .MaxAbu             'put initial filters on max
     .DataFilter(fltIsoAbu, 2) = .MaxAbu
     .DataFilter(fltCSMW, 2) = .MaxMW
     .DataFilter(fltIsoMW, 2) = .MaxMW
     .DataFilter(fltIsoCS, 2) = 1000                'maximum charge state
End With
If PracMinScan <> PracMaxScan Then
   With GelAnalysis(0)
        .GANET_Fit = 1
        .GANET_Slope = (PracMaxNET - PracMinNET) / (PracMaxScan - PracMinScan)
        .GANET_Intercept = PracMaxNET - .GANET_Slope * PracMaxScan
        .NET_TICFit = 1
        .NET_Slope = .GANET_Slope
        .NET_Intercept = .GANET_Intercept
   End With
End If
GelStatus(0).Dirty = True
GelBody(0).Tag = 0
GelBody(0).Caption = "--- MT tags Display --- "
GelBody(0).Show
Screen.MousePointer = vbDefault
Exit Sub

Display0ErrorHandler:
Debug.Assert False
LogErrors Err.Number, "Display0"
End Sub

Private Sub Display0DataSource(ByVal DBGeneration As Long)
If DBGeneration >= glAMT_GENERATION_MT_1 Then           'MT tag database
   GelData(0).Comment = GelData(0).Comment & vbCrLf & _
        "This display created based on data from PRISM system" & CurrMTDBInfo
Else                                'legacy database
   GelData(0).Comment = GelData(0).Comment & vbCrLf & _
        "This display created from legacy MT tag database " & CurrLegacyMTDatabase
End If
End Sub

Public Function FileNew(ByVal hwndOwner As Long, Optional ByVal strInputFilePath As String = "", Optional lngGelIndexToForce As Long = 0, Optional ByRef strErrorMessage As String = "") As Long
'---------------------------------------------------------------------------------------
'Opens the file given by strInputFilePath, or prompts the user to choose a .Pek, .CSV, .mzXML, or .mzData file
'Returns the index of the file in memory if success, 0 otherwise
'If lngGelIndexToForce is > 0 then the data will be loaded into the gel with the given index
'---------------------------------------------------------------------------------------

Dim fIndex As Long
Dim sFileName As String
Dim OpenResult As Integer
Dim blnInteractiveMode As Boolean

Dim fso As FileSystemObject
Dim objFile As File

On Error Resume Next

strErrorMessage = ""

If Len(strInputFilePath) = 0 Then
    blnInteractiveMode = True
    
    sFileName = SelectFile(hwndOwner, _
                      "Select source .Pek, .CSV, .mzXML, or .mzData file", "", False, "", _
                      "All Files (*.*)|*.*|" & _
                      "PEK Files (*.pek)|*.pek|" & _
                      "CSV Files (*.csv)|*.csv|" & _
                      "mzXML Files (*.mzXML)|*.mzXml|" & _
                      "mzXML Files (*mzXML.xml)|*mzXML.xml|" & _
                      "mzData Files (*.mzData)|*.mzData|" & _
                      "mzData Files (*mzData.xml)|*mzData.xml", _
                      glbPreferencesExpanded.LastInputFileMode + 2)
    
    If Len(sFileName) > 0 Then
        If Not FileExists(sFileName) Then
            strErrorMessage = "File not found: " & sFileName
            sFileName = ""
            MsgBox strErrorMessage, vbExclamation + vbOKOnly, "Error"
        End If
    End If
Else
    blnInteractiveMode = False
    If FileExists(strInputFilePath) Then
        sFileName = strInputFilePath
    Else
        sFileName = ""
        strErrorMessage = "File not found: " & strInputFilePath
    End If
End If

If Len(sFileName) > 0 Then ' User selected a file.
    UpdatePreferredFileExtension sFileName
   
   'Find the next available index
   If lngGelIndexToForce > 0 And lngGelIndexToForce <= UBound(GelBody()) Then
      fIndex = lngGelIndexToForce
   Else
      fIndex = FindFreeIndex()
   End If
   GelData(fIndex).Comment = glCOMMENT_CREATED & Now & vbCrLf & glCOMMENT_USER & UserName
   
   ' MonroeMod
   AddToAnalysisHistory fIndex, "New gel created (user " & UserName & ")"
   
' No longer supported (March 2006)
''   IsDBFile = (LCase(GetFileExtension(sFileName)) = ".mdb")
''   If IsDBFile Then
''      GelData(fIndex).Certificate = glCERT2000_DB
''      GelData(fIndex).PathtoDatabase = sFileName
''      frmGelFromDB.Tag = fIndex
''      frmGelFromDB.Show vbModal
''      If GelStatus(fIndex).DBGel = 0 Then   'user canceled
''         SetGelStateToDeleted fIndex
''         MDIStatus False, ""
''         Exit Function
''      End If
''   Else
    
   GelData(fIndex).Certificate = glCERT2003
   GelData(fIndex).pICooSysEnabled = False
   GelData(fIndex).PathtoDatabase = ""

   If fIndex > glMaxGels Then
      MsgBox "Command aborted. Too many open files.", vbOKOnly, glFGTU
      Exit Function
   End If
   
   With GelData(fIndex)  'save parameters for this doc
        ' Make sure sFileName contains the full path to the file
        Set fso = New FileSystemObject
        Set objFile = fso.GetFile(sFileName)
        sFileName = objFile.Path
        
        ' The full path to the .Pek, .CSV, .mzXML, or .mzData file
        .FileName = sFileName
        .Fileinfo = GetFileInfo(sFileName)
        ResetDataFilters fIndex, glPreferences
        ' MonroeMod
        AddToAnalysisHistory fIndex, "Loading File; " & .Fileinfo
   End With
   
' No longer supported (March 2006)
''   If IsDBFile Then
''     Screen.MousePointer = vbHourglass
''     OpenResult = LoadNewDBGel(sFileName, fIndex)
''   Else
   
   OpenResult = LoadNewData(sFileName, fIndex, blnInteractiveMode)
   Select Case OpenResult
   Case 0      'success
      Debug.Assert Not GelStatus(fIndex).Deleted
      GelStatus(fIndex).Dirty = True
      GelBody(fIndex).Tag = fIndex
      GelBody(fIndex).Caption = "Untitled:" & fIndex
      GelData(fIndex).PathtoDatabase = glbPreferencesExpanded.LegacyAMTDBPath
      ' MonroeMod: Need to add recent files to file menu
      GetRecentFiles
      GelBody(fIndex).Show
      FileNew = fIndex
   Case -1     'user canceled load of large data set
      MDIStatus False, "Done"
      GelStatus(fIndex).Deleted = True
      FileNew = 0
   Case -2     'data sets too large
      MDIStatus False, "Done"
      strErrorMessage = "Dataset too large"
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strErrorMessage & "; File Path = " & sFileName, vbOKOnly, glFGTU
      End If
      GelStatus(fIndex).Deleted = True
      FileNew = 0
   Case -3     'data structure problem
      MDIStatus False, "Done"
      strErrorMessage = "Scan numbers in the input file must be in ascending order."
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strErrorMessage & vbCrLf & "Open the " & GetFileExtension(sFileName) & " file with any text editor and make changes.", vbOKOnly, glFGTU
      End If
      GelStatus(fIndex).Deleted = True
      FileNew = 0
   Case -4     'no valid data
      MDIStatus False, "Done"
      strErrorMessage = "No valid data found in file"
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strErrorMessage & "; File Path = " & sFileName, vbOKOnly, glFGTU
      End If
      GelStatus(fIndex).Deleted = True
      FileNew = 0
   Case -5     ' User Cancelled load in the middle of loading (or post-load processing)
      MDIStatus False, "Done"
      strErrorMessage = "Load cancelled"
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strErrorMessage, vbOKOnly, glFGTU
      End If
      GelStatus(fIndex).Deleted = True
      FileNew = 0
   Case -6, -7
      MDIStatus False, "Done"
      strErrorMessage = "File not found"
      If OpenResult = -6 And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strErrorMessage & "; File Path = " & sFileName, vbOKOnly, glFGTU
      End If
      GelStatus(fIndex).Deleted = True
      FileNew = 0
   Case Else   'some other error
      MDIStatus False, "Done"
      strErrorMessage = "Error loading data from file; file maybe contains no data or structure does not match expected format"
      If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
          MsgBox strErrorMessage & "; File Path = " & sFileName, vbOKOnly, glFGTU
      End If
      GelStatus(fIndex).Deleted = True
      FileNew = 0
   End Select
   If GelStatus(fIndex).Deleted Then
       SetGelStateToDeleted fIndex
   End If
Else
   MDIStatus False, "Done"
   FileNew = 0
End If
frmProgress.HideForm
Screen.MousePointer = vbDefault
End Function

Public Function LoadNewData(ByVal strInputFilePath As String, ByVal lngGelIndex As Long, ByVal blnInteractiveMode As Boolean) As Integer
    '---------------------------------------------------------------------------------------
    'Returns 0 if data successfuly loaded, -1 if a user cancelled loading of large file, -2
    'if data set is too large, -3 if problems with scan numbers, -4 if no data found, -5
    'if user cancels load, -6 if file not found, -7 for file error that user was already
    'notified about, -10 for any other error
    '---------------------------------------------------------------------------------------
    Dim objLoadOptionsForm As frmFileLoadOptions
    Dim objMZXMLFileReader As clsFileIOMZXml
    Dim objmzDataFileReader As clsFileIOMZData
    
    Dim strFileExtension As String
    
    Dim udtFilterPrefs As udtAutoAnalysisFilterPrefsType
    Dim blnMSLevelFilter() As Boolean
    
    Dim eScanFilterMode As eosEvenOddScanFilterModeConstants
    Dim eDataFilterMode As dfmCSandIsoDataFilterModeConstants
    
    Dim eFileType As ifmInputFileModeConstants
    Dim intReturnCode As Integer
    
    On Error GoTo err_LoadNewData
    
    If Not DetermineFileType(strInputFilePath, eFileType) Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "The input file contains an unknown extension.  It must be " & KNOWN_FILE_EXTENSIONS, vbExclamation + vbOKCancel, glFGTU
        End If
        intReturnCode = -7
        LoadNewData = intReturnCode
        Exit Function
    End If
       
    Set objLoadOptionsForm = New frmFileLoadOptions
    objLoadOptionsForm.SetFilePath strInputFilePath
    
    ' Copy the data from .AutoAnalysisFilterPrefs
    udtFilterPrefs = glbPreferencesExpanded.AutoAnalysisFilterPrefs
    
    With udtFilterPrefs
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            ' Auto analysis
            objLoadOptionsForm.FilterOnIsoFit = .ExcludeIsoByFit
        Else
            ' Manual analysis
            If .ExcludeIsoByFit Then objLoadOptionsForm.FilterOnIsoFit = True
        End If
        
        objLoadOptionsForm.IsoFitMax = .ExcludeIsoByFitMaxVal
        
        objLoadOptionsForm.AbuFilterMin = .RestrictIsoAbundanceMin
        objLoadOptionsForm.AbuFilterMax = .RestrictIsoAbundanceMax
        
        objLoadOptionsForm.FilterOnAbundance = .RestrictIsoByAbundance
        
        objLoadOptionsForm.MaximumDataCountEnabled = .MaximumDataCountEnabled
        objLoadOptionsForm.MaximumDataCountToLoad = .MaximumDataCountToLoad
        
        If .RestrictToEvenScanNumbersOnly Or .RestrictToOddScanNumbersOnly Then
            If .RestrictToOddScanNumbersOnly Then
                objLoadOptionsForm.EvenOddScanFilterMode = eosLoadOddScansOnly
            Else
                objLoadOptionsForm.EvenOddScanFilterMode = eosLoadEvenScansOnly
            End If
        Else
            If InStr(UCase(strInputFilePath), "DREAMS") > 0 Then
                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Or blnInteractiveMode Then
                    ' This appears to be Dreams data
                    ' Inform user that loading of only Even Numbered scans will be enabled, but can be customized
                    MsgBox "The input file contains 'Dreams' in the name.  Consequently, loading of only odd-numberd data has been enabled.  You can customize this option if desired.", vbInformation + vbOKOnly, "DREAMS Data Filter"
                End If
                objLoadOptionsForm.EvenOddScanFilterMode = eosLoadOddScansOnly
            Else
                objLoadOptionsForm.EvenOddScanFilterMode = eosLoadAllScans
            End If
        End If
    End With
    
    ' Only show the LoadOptions form if interacting with the user
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Or blnInteractiveMode Then
        objLoadOptionsForm.Show vbModal
    End If
    
    If objLoadOptionsForm.LoadCancelled Then
        intReturnCode = -1
    Else
        objLoadOptionsForm.GetMSLevelFilter blnMSLevelFilter
        eScanFilterMode = eosLoadAllScans
        eDataFilterMode = dfmLoadAllData
        
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            With udtFilterPrefs
                If .RestrictToOddScanNumbersOnly Then
                    eScanFilterMode = eosEvenOddScanFilterModeConstants.eosLoadOddScansOnly
                ElseIf .RestrictToEvenScanNumbersOnly Then
                    eScanFilterMode = eosEvenOddScanFilterModeConstants.eosLoadEvenScansOnly
                Else
                    eScanFilterMode = eosEvenOddScanFilterModeConstants.eosLoadAllScans
                End If
            End With
        Else
            With udtFilterPrefs
                .ExcludeIsoByFit = objLoadOptionsForm.FilterOnIsoFit
                If .ExcludeIsoByFit Then
                    .ExcludeIsoByFitMaxVal = objLoadOptionsForm.IsoFitMax
                Else
                    .ExcludeIsoByFitMaxVal = glHugeDouble
                End If

                .RestrictIsoByAbundance = objLoadOptionsForm.FilterOnAbundance
                If objLoadOptionsForm.FilterOnAbundance Then
                    .RestrictIsoAbundanceMin = objLoadOptionsForm.AbuFilterMin
                    .RestrictIsoAbundanceMax = objLoadOptionsForm.AbuFilterMax
                End If
                .RestrictCSByAbundance = .RestrictIsoByAbundance
                .RestrictCSAbundanceMin = .RestrictIsoAbundanceMin
                .RestrictCSAbundanceMax = .RestrictIsoAbundanceMax
                
                .MaximumDataCountEnabled = objLoadOptionsForm.MaximumDataCountEnabled
                If .MaximumDataCountEnabled Then
                   .MaximumDataCountToLoad = objLoadOptionsForm.MaximumDataCountToLoad
                End If
                
                .RestrictToOddScanNumbersOnly = False
                .RestrictToEvenScanNumbersOnly = False
                
                eScanFilterMode = objLoadOptionsForm.EvenOddScanFilterMode
                Select Case eScanFilterMode
                Case eosEvenOddScanFilterModeConstants.eosLoadOddScansOnly
                   .RestrictToOddScanNumbersOnly = True
                Case eosEvenOddScanFilterModeConstants.eosLoadEvenScansOnly
                   .RestrictToEvenScanNumbersOnly = True
                End Select
                
                eDataFilterMode = objLoadOptionsForm.DataFilterMode
            End With
            
            ' Copy values from udtFilterPrefs to glbPreferencesExpanded.AutoAnalysisFilterPrefs
            With glbPreferencesExpanded.AutoAnalysisFilterPrefs
                ' Note: I'm purposely not updating .ExcludeIsoByFit and .ExcludeIsoByFitMaxVal if the user customized them
                
                .RestrictIsoByAbundance = udtFilterPrefs.RestrictIsoByAbundance
                .RestrictIsoAbundanceMin = udtFilterPrefs.RestrictIsoAbundanceMin
                .RestrictIsoAbundanceMax = udtFilterPrefs.RestrictIsoAbundanceMax
                
                .RestrictCSByAbundance = udtFilterPrefs.RestrictCSByAbundance
                .RestrictCSAbundanceMin = udtFilterPrefs.RestrictCSAbundanceMin
                .RestrictCSAbundanceMax = udtFilterPrefs.RestrictCSAbundanceMax
                
                .MaximumDataCountEnabled = udtFilterPrefs.MaximumDataCountEnabled
                .MaximumDataCountToLoad = udtFilterPrefs.MaximumDataCountToLoad
                
                .RestrictToOddScanNumbersOnly = udtFilterPrefs.RestrictToOddScanNumbersOnly
                .RestrictToEvenScanNumbersOnly = udtFilterPrefs.RestrictToEvenScanNumbersOnly
            End With

        End If
        
        Screen.MousePointer = vbHourglass
        
        ' Note: In the following calls to LoadNewPEK, LoadNewCSV, LoadNewMZXML, and LoadNewMZData, I would prefer to
        '  simply pass udtFilterPrefs into the functions, but VB6 won't let me, saying either "circular dependencies between modules"
        '  or "modules can only accept public user defined types as parameters".  Neither of these errors should be occuring, but they are
        '  and thus I'm forced to pass individual variables into the functions
        Select Case eFileType
        Case ifmInputFileModeConstants.ifmPEKFile
            With udtFilterPrefs
                intReturnCode = LoadNewPEK(strInputFilePath, lngGelIndex, .ExcludeIsoByFitMaxVal, _
                                           .RestrictIsoByAbundance, .RestrictIsoAbundanceMin, .RestrictIsoAbundanceMax, _
                                           .MaximumDataCountEnabled, .MaximumDataCountToLoad, _
                                           eScanFilterMode, eDataFilterMode)
            End With
            
        Case ifmInputFileModeConstants.ifmCSVFile
            With udtFilterPrefs
                intReturnCode = LoadNewCSV(strInputFilePath, lngGelIndex, .ExcludeIsoByFitMaxVal, _
                                           .RestrictIsoByAbundance, .RestrictIsoAbundanceMin, .RestrictIsoAbundanceMax, _
                                           .MaximumDataCountEnabled, .MaximumDataCountToLoad, _
                                           eScanFilterMode, eDataFilterMode)
            End With
            
        Case ifmInputFileModeConstants.ifmmzXMLFile, ifmInputFileModeConstants.ifmmzXMLFileWithXMLExtension
            Set objMZXMLFileReader = New clsFileIOMZXml
            With udtFilterPrefs
                intReturnCode = objMZXMLFileReader.LoadNewMZXML(strInputFilePath, lngGelIndex, .ExcludeIsoByFitMaxVal, _
                                                                .RestrictIsoByAbundance, .RestrictIsoAbundanceMin, .RestrictIsoAbundanceMax, _
                                                                .MaximumDataCountEnabled, .MaximumDataCountToLoad, _
                                                                eScanFilterMode, eDataFilterMode, blnMSLevelFilter)
            End With
            Set objMZXMLFileReader = Nothing
            
        Case ifmInputFileModeConstants.ifmmzDataFile, ifmInputFileModeConstants.ifmmzDataFileWithXMLExtension
            Set objmzDataFileReader = New clsFileIOMZData
            With udtFilterPrefs
                intReturnCode = objmzDataFileReader.LoadNewMZData(strInputFilePath, lngGelIndex, .ExcludeIsoByFitMaxVal, _
                                                                .RestrictIsoByAbundance, .RestrictIsoAbundanceMin, .RestrictIsoAbundanceMax, _
                                                                .MaximumDataCountEnabled, .MaximumDataCountToLoad, _
                                                                eScanFilterMode, eDataFilterMode, blnMSLevelFilter)
            End With
            Set objmzDataFileReader = Nothing
            
        Case Else
            intReturnCode = -7
        End Select
        
        If intReturnCode = 0 Then
            If udtFilterPrefs.ExcludeIsoByFit Then
                With GelData(lngGelIndex)
                    .Comment = .Comment & vbCrLf & strFileExtension & " file may contain more data than was loaded. Only loaded isotopic data with calculated fit better than " & udtFilterPrefs.ExcludeIsoByFitMaxVal
                    .DataFilter(fltIsoFit, 0) = True
                    .DataFilter(fltIsoFit, 1) = udtFilterPrefs.ExcludeIsoByFitMaxVal
                    AddToAnalysisHistory lngGelIndex, "File Loaded; Only isotopic data with calculated fit better than " & udtFilterPrefs.ExcludeIsoByFitMaxVal & " loaded (at user request)."
                End With
            End If
        
            If udtFilterPrefs.RestrictIsoByAbundance Then
                With GelData(lngGelIndex)
                    .Comment = .Comment & vbCrLf & strFileExtension & " file may contain more data than was loaded. Only loaded isotopic data with abundance between " & Trim(udtFilterPrefs.RestrictIsoAbundanceMin) & " and " & Trim(udtFilterPrefs.RestrictIsoAbundanceMax) & " counts"
                    .DataFilter(fltIsoAbu, 0) = True
                    .DataFilter(fltIsoAbu, 1) = udtFilterPrefs.RestrictIsoAbundanceMin
                    .DataFilter(fltIsoAbu, 2) = udtFilterPrefs.RestrictIsoAbundanceMax
                    AddToAnalysisHistory lngGelIndex, "File Loaded; Only isotopic data with abundance between " & Trim(objLoadOptionsForm.AbuFilterMin) & " and " & Trim(objLoadOptionsForm.AbuFilterMax) & " counts loaded (at user request)."
                End With
            End If
        End If
        
    End If

    Unload objLoadOptionsForm
    Set objLoadOptionsForm = Nothing
    Screen.MousePointer = vbDefault
    
    LoadNewData = intReturnCode
    
Exit Function

err_LoadNewData:
' Error during load
If Err.Number = 53 Then
    ' File not found
    LoadNewData = -6
Else
    Debug.Assert False
    LoadNewData = -10
End If
Screen.MousePointer = vbDefault

End Function

Public Sub SetGelStateToDeleted(lngGelIndex As Long)
    With GelStatus(lngGelIndex)
        .Deleted = True
        .SourceDataRawFileType = rfcUnknown
    End With
End Sub

Public Sub UpdatePreferredFileExtension(strFileName As String)
    Dim eFileType As ifmInputFileModeConstants
    
    If DetermineFileType(strFileName, eFileType) Then
        glbPreferencesExpanded.LastInputFileMode = eFileType
    End If
    
End Sub

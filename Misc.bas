Attribute VB_Name = "Module3"
'MISCELLANEOUS DEFINITIONS, DECLARATIONS AND FUNCTIONS
'Last Modified: 02/12/2003 nt
'---------------------------------------------------------------------------------------
Option Explicit
'GLOBAL VARIABLES
Public glTracking As Boolean            'True for coordinates tracking On. False for Off
' The following variable is unused (November 2003)
''Public glUpdateGel As Boolean           'True when gel needs updates after it receives focus
Public glInitFile As String
Public glWriteFreqShift As Boolean      'When True frequency shift information will be
                                        'transfered when writing PEH field from gel file
Public glAbortUMCProcessing As Boolean

'DEFAULT LOCAL PREFERENCES - loaded from Registry
Public glPreferences As GelPrefs
'ICR2LS
Public sICR2LSCommand As String         'Path to ICR2Ls exec file

' No longer supported (March 2006)
''Public sFTICR_AMTPath As String         'path to FTICR_AMT database

'COLORS
Public glBackColor As Long              'background color for graph
Public glForeColor As Long              'font color for graph
Public glUnderColor As Long             'underexpression color
Public glOverColor As Long              'overexpression color
Public glMidColor As Long               'transition color for the differential display
Public glCSColor As Long                'Charge State color
Public glIsoColor As Long               'Isotopic color
Public glSelColor As Long               'Selection Color
Public glDDRatioMax As Double           '
'MARKER SHAPE
Public glCSShape As Integer             'shape of the cs spots
Public glIsoShape As Integer            'shape of the iso spots
'COLOR ARRAYS
Public aDDColors(-50 To 50) As Long     'colors used in diff. display view
Private aDDScale(-50 To 50) As Double    'intervals used in diff. display view

'Display0 options
Public PresentDisplay0Type As Long      'load MT tags as Charge State or Isotopic
Public Display0MinScan As Long          'first scan number for Display0
Public Display0MaxScan As Long          'last scan number

' No longer supported (March 2006)
''Public Display0UseFakeAbundance As Boolean      'use abundance constructed from MT tags score

Public bSetFileNameDate As Boolean      'if True set file name and date on printed gel
Public bIncludeTextLabels As Boolean    'if True then show textual labels on printed gel
Public vWhatever As Variant             'use this for whatever you need
'''Public vMessage(10) As Variant          'use this for whatever you need

'for now charge state map is hard coded for 5 charge states 1,2,3,4,5 + all others
'but specifying different ranges in glCS1-glCS2 we can draw different colors
Public glCS1(5) As Long                   'list of charge states tracked
Public glCS2(5) As Long                   'map of charge state colors

' Note: these file formats are no no longer supported (March 2006)
Public Const glCERT1999 = "2DGelVisual1999.01" 'put in gel files to be sure
Public Const glCERT2000 = "2DGelVisual2000.01"

' This file format is still supported
Public Const glCERT2003 = "La2DDisplay.01"

' Note: these file formats are no no longer supported (March 2006)
'ORF gel certificate
Public Const glCERT2000_DB = "2DGelVisual2000.DB"
Public Const glCERT2002_MT = "2DGelVisual2002.MT"

' MonroeMod Begin
Public Const glCERT2003_Modular = "2DGelVisual2003.01"
Public Const glCERT_FileNotFound = "#FILE_NOT_FOUND#"
' MonroeMod Finish

Public Const glARG_SEP = ";"
Public Const NoHarvest = "Not found."

''Public Const ThisApp = "2DGelLand"          ' Registry App constant.
''Public Const OptionsKey = "Options"         ' Registry Key constant.
Public Const RawDataTmpFile = "Oneum.txt"   ' name of Temp file with raw data

Public Const glFGTU = "VIPER"

'initialization file name
Public Const INIT_FILE_NAME = "FAXA.init"

Public Const IdUnknown = "Unknown"

Public Enum fstFileSaveTypeConstants
    fstGel = 0          'save with extension .gel
    fstUMR = 1          'save with extension .umw
    fstPIC = 2          'save as picture
    fstTxt = 100
End Enum

Public Const glMaxGels = 10000          ' maximum number of gels
Public Const glCSType = 1               ' constants used to mark
Public Const glIsoType = 2              ' different data origins
Public Const glNoType = -1              ' used with hot spots

Public Const glMIL = 1000000

Public Const glDONT_DISPLAY = -51       'mark that something shouldn't be displayed

Public Const MW_FIELD_OFFSET = 6        ' Related to mftMWAvg, mftMWMono, and mftMWTMA

'LOGICAL COORDINATE SYSTEM CONSTANTS-SIZE
Public Const LDfX0 = 0          'logical coordinates defaults
Public Const LDfY0 = 0          '(X0,Y0)-(XE,YE) define real
Public Const LDfX1 = 300        'logical window; (X1,Y1) defines
Public Const LDfY1 = 200        'small offset from the coordinate
Public Const LDfX2 = 9700       'logical window; (X1,Y1)-(X2,Y2)
Public Const LDfY2 = 9800       'defines small offset from the
Public Const LDfXE = 10000      'coordinate axes
Public Const LDfYE = 10000
'LOGICAL COORDINATE SYSTEM CONSTANTS-INDENTS
Public Const lDfSXPercent = 0.02
Public Const lDfSYPercent = 0.03
Public Const lDfLXPercent = 0.08
Public Const lDfLYPercent = 0.07
'LOGICAL COORDINATE SYSTEM CONSTANTS-OTHER
Public Const LDfWndW As Long = (LDfXE - LDfX0) / (1 - lDfSXPercent - lDfLXPercent)
Public Const LDfWndH As Long = (LDfYE - LDfY0) / (1 - lDfSYPercent - lDfLYPercent)
Public Const LDfSX As Long = lDfSXPercent * LDfWndW
Public Const lDfSY As Long = lDfSYPercent * LDfWndH
Public Const lDfLX As Long = lDfLXPercent * LDfWndW
Public Const lDfLY As Long = lDfLYPercent * LDfWndH

Public Const lDfMinSz As Long = 25      'minimum size of a spot (overlay)
Public Const lDfMaxSz As Long = 500     'maximum size of a spot

Public Const glShapeEli = 0             ' shape of alms constants
Public Const glShapeRec = 1             ' Ellipse, Rectangle
Public Const glShapeRRe = 2             ' round rectangle
Public Const glShapeTri = 3             ' Triangle
Public Const glShapeSta = 4             ' Star
Public Const glShapeHex = 5             ' Hexagon
Public Const glShapeGradRec = 6         ' gradient filled rectangle

Public Const TrackerCaption = "Coordinates" 'used to determine are
Public Const ICR2LSCaption = "ICR-2LS"      'specific windows loaded
Public Const UMRFileInfo = "Created from the unique MW results."

Public Const glNormalDisplay = 1        ' view constants
Public Const glDifferentialDisplay = 2
Public Const glChargeStateMapDisplay = 3

Public Const glCSColorDefault = 8388863         ' RGB(128, 0, 0)
Public Const glIsoColorDefault = vbBlue

Public Const glUnderColorDefault = vbGreen
Public Const glOverColorDefault = vbRed
Public Const glMidColorDefault = vbBlack

Public Const glHugeUnderExp = 1E-38     ' exposure constants
Public Const glHugeOverExp = 3E+38      ' used to scale
Public Const glHugeOverReal = 100       ' default for nDDRatioMax
Public Const glExposureNA = -1          ' exposure indexes
Public Const glHugeLong = 2147483647
Public Const glHugeDouble = 1.79E+308

Public Const glCSOnTop = 1              ' Z-order constants
Public Const glIsoOnTop = 2

Public Const glNoAction = 0             ' actions on graph
Public Const glActionHit = 1
Public Const glActionZoom = 2

Public Const glExpand = 1               'used to navigate in the graph
Public Const glShrink = -1
Public Const glTuneLT = 0               'tune left or bottom
Public Const glTuneRB = 1               'tune right or top
Public Const glTuneMoveL = 2            'move left
Public Const glTuneMoveR = 3            'move right
Public Const glTuneMoveU = 2            'move up
Public Const glTuneMoveD = 3            'move down

'COORDINATE SYSTEM POSITION CONSTANTS
Public Const glOriginBL = 1             'top left
Public Const glOriginBR = 2             'top right
Public Const glOriginTL = 3             'bottom left
Public Const glOriginTR = 4             'bottom right

Public Const glNormal = 1               'normal orientation/definition
Public Const glReverse = 2              'reverse orientation/definition

Public Const glPICooSys = 0
Public Const glFNCooSys = 1
Public Const glNETCooSys = 2

Public Const glVAxisLin = 0
Public Const glVAxisLog = 1

Public Const glPCT = 0.01
Public Const glPPM = 0.000001

' These have been replaced by glScope.glSc_All and glScope.glSc_Current
''Public Const glSCOPE_ALL = 0            'all data
''Public Const glSCOPE_CURRENT = 1        'current view data

Public Const glUMC_TYPE_INTENSITY = 0
Public Const glUMC_TYPE_FIT = 1
Public Const glUMC_TYPE_MINCNT = 2
Public Const glUMC_TYPE_MAXCNT = 3
Public Const glUMC_TYPE_UNQAMT = 4
Public Const glUMC_TYPE_ISHRINKINGBOX = 5
Public Const glUMC_TYPE_FSHRINKINGBOX = 6
Public Const glUMC_TYPE_FROM_NET = 7            ' UMCIonNet searching

'physical constants
' Mass of the Charge Carrier (aka Hydrogen minus 1 electron = 1 proton)
Public Const glMASS_CC As Double = 1.00727649
                   
'base mass for correction
Public Const glN14_N15CorrMW = 2000

Public Const glICAT_Light = 442.2249697
Public Const glICAT_Heavy = 450.2752
Public Const glICAT_Delta = 8.051

Public Const glSPICAT_Light = 170.1055
Public Const glSPICAT_Heavy = 177.1227
Public Const glSPICAT_Delta = 7.017169

Public Const glPEO = 414.1936713     'PEO label does not have light and heavy version
Public Const glALKYLATION = 57.0215  'iodoacetamide

Public Const glPHOSPHORYLATION = "STY"
Public Const glPHOSPHORYLATION_Mass = 79.96633

Public Const glN14N15_DELTA = 0.9970356
Public Const glC12C13_DELTA = 1.0033554
Public Const glO16O18_DELTA = 4.0085
Public Const glDeuterium_DELTA = 1.006128

Public Const glERMAX_FINITE As Double = 100.8989671

Public Const glSOLO_CLR = &HC0FFFF
Public Const glUMC_CLR = &HC0FFC0

Public Const sErrLogReference = "See file UnexpErr.log for more information."

Public Const Msg_GE_0 = "This argument should be non-negative number."
Public Const Msg_GT_0 = "This argument should be positive number."
Public Const Msg_Numeric = "This argument should be numeric."
Public Const Msg_Integer = "This argument should be integer."

Public Const ic_N14 = "N14"
Public Const ic_N15 = "N15"

' Unused enum (July 2003)
''''enumeration of function GetNumberWord
'''Public Enum glNumWord
'''    nwEmpty = 0
'''    nwNumOnly = 1
'''    nwWordOnly = 2
'''    nwNumWord = 3
'''End Enum

Public Type GelRes         'this type is used to return results
    CSRes As String        'of functions on GelData
    IsoRes As String
    AllRes As String
End Type

Private Type udtSegmentStatsType
    UMCHitCountUsed As Long
    ArrayCountUnused As Long
    UnusedUMCIndices() As Long          ' 0-based array; holds the indices of the Unused LC-MS Features in this segment
End Type

Public Function GetFileNameOnly(ByVal SP As String) As String
'returns file name only from the full path string

    Dim fso As New FileSystemObject
    GetFileNameOnly = fso.GetFileName(SP)

End Function

Public Function AddDirSeparator(SP As String) As String
'adds directory separator if it's not at the end of
'string except if sP is empty string in which case
'it returns empty string
Dim sCoolString As String
sCoolString = Trim$(SP)
AddDirSeparator = ""
If Len(sCoolString) > 0 Then
   If Right$(sCoolString, 1) <> "\" Then
      AddDirSeparator = sCoolString & "\"
   Else
      AddDirSeparator = sCoolString
   End If
End If
End Function

Public Sub Initialize()
'---------------------------------------------------------
'This procedure is called when program starts
'---------------------------------------------------------
glInitFile = App.Path & "\" & INIT_FILE_NAME        ' This is the FAXA.Init file, not the .Ini settings file
glWriteFreqShift = False    'no by default
glbPreferencesExpanded.PairSearchOptions.SearchDef.ERCalcType = ectER_RAT     'ratio(by Big Kahuna decision)
InitSelectMatrices
InitJobMatrices
InitChargeStateMap
SetDefaultDefinitions
PopulateDefaultInternalStds UMCInternalStandards
' Make sure the Default Definitions have been defined before calling IniFileLoadSettings
IniFileLoadSettings glbPreferencesExpanded, UMCDef, UMCIonNetDef, UMCNetAdjDef, UMCInternalStandards, samtDef, glPreferences
GetRecentFiles
DDInitColors
SetBackForeColorObjects
SetCSIsoColorObjects
SetSelColorObjects
CreateOlyBackClrObject OlyOptions.BackColor
CreateOlyForeClrObject OlyOptions.ForeColor
SetDDRColorObjects
InitICR2LS
' Unused function (September 2006)
''InitMwtWin
ValidateDotNETDLLs
InitDisplay0
ParseCommandLine
End Sub

Public Function GetERClrInd(ByVal ER As Double) As Integer
'returns color index in aDDColors for specified ER
Dim i As Integer
On Error Resume Next
GetERClrInd = -50
i = -50
Do While ER > aDDScale(i) And i < 50
   i = i + 1
Loop
GetERClrInd = i
End Function

Public Sub DDInitColors()
Dim i As Integer
Dim DDRatioScale As Double
DDRatioScale = (glDDRatioMax - 1) / 50 'variable step for DDratio>1
For i = -50 To 50
    If i < 0 Then
       aDDScale(i) = 1 + i * 0.02
       aDDColors(i) = ((&H40404 * (-i)) And (glUnderColor Xor glMidColor) Xor glMidColor)
    Else
       If i = 0 Then
          aDDScale(i) = 0.02
          aDDColors(i) = glMidColor
       Else
          aDDScale(i) = 1 + i * DDRatioScale
          aDDColors(i) = ((&H40404 * i) And (glOverColor Xor glMidColor) Xor glMidColor)
       End If
    End If
Next i
End Sub

Public Sub MDIStatus(ByVal Visibility As Boolean, _
                     ByVal Status As String)
With MDIForm1
  .lblStatus.Caption = Status
  .lblStatus.Visible = Visibility
End With
DoEvents
End Sub

Public Function Fileinfo(Ind As Long, iInfoType As Integer) As Variant
'returns variant array containing nicely formated file informations
Dim tmp() As String
On Error GoTo fileinfo_err
ReDim tmp(35)
With GelData(Ind)
   tmp(0) = "FILE PARAMETERS"
   tmp(1) = .Fileinfo
   tmp(2) = "Path to data files: " & .PathtoDataFiles
   tmp(3) = "Path to database: " & .PathtoDatabase
   tmp(4) = "Media type: " & .MediaType
   tmp(5) = "Calibration: " & .CalEquation
   tmp(6) = "Arguments: " & Format$(.CalArg(1), "0.00000") & ", " & Format$(.CalArg(2), "0.00000") & ", " & Format$(.CalArg(3), "0.00000")
   tmp(7) = "Lines in input data file: " & Format(.LinesRead, "#,##0")
   tmp(8) = "Lines with data: " & Format(.DataLines, "#,##0")
   tmp(9) = "Charge State lines: " & Format(.CSLines, "#,##0")
   tmp(10) = "Isotopic lines: " & Format(.IsoLines, "#,##0")
   tmp(11) = "Minimum MW: " & Format(.MinMW, "###,###,##0.0000")
   tmp(12) = "Maximum MW: " & Format(.MaxMW, "###,###,##0.0000")
   tmp(13) = "Minimum abundance: " & Format(.MinAbu, "Scientific")
   tmp(14) = "Maximum abundance: " & Format(.MaxAbu, "Scientific")
   tmp(15) = "pI numbers ready: " & Format$(CLng(.pICooSysEnabled), "Yes/No")
   If iInfoType = 1 Then
      ReDim Preserve tmp(15)
      Fileinfo = tmp
      Exit Function
   End If
   tmp(16) = ""
   tmp(17) = "CURRENT FILE SETTINGS"
   Select Case .Preferences.IsoDataField
   Case 6
        tmp(18) = "Isotopic Data Field: " & "Average MW"
   Case 7
        tmp(18) = "Isotopic Fata Field: " & "Monoisotopic MW"
   Case 8
        tmp(18) = "Isotopic Data Field: " & "Most Abundant MW"
   End Select
   If .DataFilter(1, 1) < 0 Then
       tmp(19) = "Duplicate Elimination Tolerance: " & "Not set"
   Else
       tmp(19) = "Duplicate Elimination Tolerance: " & Format(.DataFilter(1, 1), "#,##0.00")
   End If
   If .DataFilter(2, 1) < 0 Then
       tmp(20) = "Database Elimination Tolerance: " & "Not set"
   Else
       tmp(20) = "Database Elimination Tolerance: " & Format(.DataFilter(2, 1), "#,##0.00")
   End If
   tmp(21) = "Calculated Fit Tolerance: " & Format(.DataFilter(3, 1), "#,##0.00")
   If iInfoType = 2 Then
      ReDim Preserve tmp(21)
      Fileinfo = tmp
      Exit Function
   End If
   tmp(22) = ""
   tmp(23) = "CURRENTLY APPLIED FILTERS"
   tmp(24) = "Duplicate elimination: " & Format(CBool(.DataFilter(1, 0)), "On/Off")
   tmp(25) = "Database bad fit elimination: " & Format(CBool(.DataFilter(2, 0)), "On/Off")
   tmp(26) = "Isotopic data bad fit elimination: " & Format(CBool(.DataFilter(3, 0)), "On/Off")
   tmp(27) = "Charge State data bad St.Dev. elimination: " & Format(CBool(.DataFilter(fltCSStDev, 0)), "On/Off")
   tmp(28) = "Case two close results: " & Format(CBool(.DataFilter(4, 0)), "On/Off")
   tmp(29) = "Comparative display elimination: " & Format(CBool(.DataFilter(5, 0)), "On/Off")
   tmp(30) = "Identity elimination: " & Format(CBool(.DataFilter(6, 0)), "On/Off")
   tmp(31) = "Charge State abundance range elimination: " & Format(CBool(.DataFilter(7, 0)), "On/Off")
   If CBool(.DataFilter(7, 0)) Then
      tmp(31) = tmp(31) & " - inclusion range [" & Format(.DataFilter(7, 1), "Scientific") _
      & "," & Format(.DataFilter(7, 2), "Scientific") & "]"
   End If
   tmp(32) = "Isotopic abundance range elimination: " & Format(CBool(.DataFilter(8, 0)), "On/Off")
   If CBool(.DataFilter(8, 0)) Then
      tmp(32) = tmp(32) & " - inclusion range [" & Format(.DataFilter(8, 1), "Scientific") _
      & "," & Format(.DataFilter(8, 2), "Scientific") & "]"
   End If
   tmp(33) = "Charge State molecular mass range elimination: " & Format(CBool(.DataFilter(fltCSMW, 0)), "On/Off")
   If CBool(.DataFilter(fltCSMW, 0)) Then
      tmp(33) = tmp(33) & " - inclusion range [" & Format(.DataFilter(fltCSMW, 1), "0.0000") _
      & "," & Format(.DataFilter(fltCSMW, 2), "0.0000") & "]"
   End If
   tmp(34) = "Isotopic molecular mass range elimination: " & Format(CBool(.DataFilter(fltIsoMW, 0)), "On/Off")
   If CBool(.DataFilter(fltIsoMW, 0)) Then
      tmp(34) = tmp(34) & " - inclusion range [" & Format(.DataFilter(fltIsoMW, 1), "Scientific") _
      & "," & Format(.DataFilter(fltIsoMW, 2), "Scientific") & "]"
   End If
   tmp(35) = "Isotopic charge state range elimination: " & Format(CBool(.DataFilter(fltIsoCS, 0)), "On/Off")
   If CBool(.DataFilter(fltIsoCS, 0)) Then
      tmp(35) = tmp(35) & " - inclusion range [" & .DataFilter(fltIsoCS, 1) _
                & "," & .DataFilter(fltIsoCS, 2) & "]"
   End If
End With
Fileinfo = tmp
Exit Function

fileinfo_err:
Fileinfo = Null
End Function

Public Function LinearNETAlignmentSelectUMCToUse(ByVal lngGelIndex As Long, _
                                                 ByRef UseUMC() As Boolean, _
                                                 ByRef lngUMCCntAddedSinceLowSegmentCount As Long, _
                                                 ByRef lngUMCSegmentCntWithLowUMCCnt As Long) As Long
'------------------------------------------------------------------
'selects unique mass classes that will be used to correct NET based
'on specified criteria; returns number selected; -1 on any error
'------------------------------------------------------------------
Dim i As Long
Dim Cnt As Long

Dim ePairedSearchUMCSelection As punaPairsUMCNetAdjustmentConstants

On Error GoTo exit_LinearNETAlignmentSelectUMCToUse
LinearNETAlignmentSelectUMCToUse = -1

For i = 0 To GelUMC(lngGelIndex).UMCCnt - 1
    With GelUMC(lngGelIndex).UMCs(i)
        .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
        .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_LOWSEGMENTCOUNT_ADDITION
        .ClassStatusBits = .ClassStatusBits And Not UMC_INDICATOR_BIT_NET_ADJ_DB_HIT
    End With
Next i

ePairedSearchUMCSelection = glbPreferencesExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection
    
If ePairedSearchUMCSelection = punaPairedAll Or _
   ePairedSearchUMCSelection = punaPairedLight Or _
   ePairedSearchUMCSelection = punaPairedHeavy Then
    If Not PairsPresent(lngGelIndex) Then
        ' No pairs are present
        ePairedSearchUMCSelection = punaPairedAndUnpaired
    End If
End If
    
If ePairedSearchUMCSelection <> punaPairedAndUnpaired Then
    Select Case ePairedSearchUMCSelection
    Case punaPairedAll, punaPairedLight, punaPairedHeavy
        ' First exclude everything
        ' Then, include LC-MS Features that are paired, depending upon ePairedSearchUMCSelection
        For i = 0 To GelUMC(lngGelIndex).UMCCnt - 1
            UseUMC(i) = False
        Next i
        
        If ePairedSearchUMCSelection = punaPairedAll Then
            ' Add back all LC-MS Features belonging to pairs
            For i = 0 To GelP_D_L(lngGelIndex).PCnt - 1
                UseUMC(GelP_D_L(lngGelIndex).Pairs(i).P1) = True
                UseUMC(GelP_D_L(lngGelIndex).Pairs(i).P2) = True
            Next i
        ElseIf ePairedSearchUMCSelection = punaPairedHeavy Then
            ' Add back LC-MS Features belonging to the heavy member of pairs
            For i = 0 To GelP_D_L(lngGelIndex).PCnt - 1
                UseUMC(GelP_D_L(lngGelIndex).Pairs(i).P2) = True
            Next i
        Else
            ' punaPairedLight
            ' Add back LC-MS Features belonging to the light member of pairs
            For i = 0 To GelP_D_L(lngGelIndex).PCnt - 1
                UseUMC(GelP_D_L(lngGelIndex).Pairs(i).P1) = True
            Next i
        End If
    Case punaUnpairedOnly
        ' Exclude LC-MS Features that are paired
        For i = 0 To GelP_D_L(lngGelIndex).PCnt - 1
            UseUMC(GelP_D_L(lngGelIndex).Pairs(i).P1) = False
            UseUMC(GelP_D_L(lngGelIndex).Pairs(i).P2) = False
        Next i
    Case punaUnpairedPlusPairedLight
        ' Exclude LC-MS Features that belong to heavy members of pairs
        For i = 0 To GelP_D_L(lngGelIndex).PCnt - 1
            UseUMC(GelP_D_L(lngGelIndex).Pairs(i).P2) = False
        Next i
    End Select
    
End If

If UMCNetAdjDef.MinUMCCount > 1 Or UMCNetAdjDef.MinScanRange > 1 Then
    ' filter-out all mass classes with insufficient membership
    '  or
    ' filter-out all mass classes with insufficient scan coverage
    For i = 0 To GelUMC(lngGelIndex).UMCCnt - 1
        If UseUMC(i) = True Then
            UseUMC(i) = LinearNETAlignmentUMCSelectionFilterCheck(lngGelIndex, i)
        End If
    Next i
End If

If UMCNetAdjDef.TopAbuPct >= 0 And UMCNetAdjDef.TopAbuPct < 100 Then
    ' Filter-out low abundant classes
    ' However, if .RequireDispersedUMCSelection = True, then make sure we have some LC-MS Features from all portions of the data
    LinearNETAlignmentSelectUMCsToUseWork lngGelIndex, _
                                          glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions.RequireDispersedUMCSelection, _
                                          UseUMC(), _
                                          lngUMCCntAddedSinceLowSegmentCount, _
                                          lngUMCSegmentCntWithLowUMCCnt
End If

Debug.Assert UBound(UseUMC) = GelUMC(lngGelIndex).UMCCnt - 1
For i = 0 To UBound(UseUMC)
    With GelUMC(lngGelIndex).UMCs(i)
        If UseUMC(i) Then
            Cnt = Cnt + 1
            ' Turn on the UMC_INDICATOR_BIT_USED_FOR_NET_ADJ bit
            .ClassStatusBits = .ClassStatusBits Or UMC_INDICATOR_BIT_USED_FOR_NET_ADJ
        End If
    End With
Next i
LinearNETAlignmentSelectUMCToUse = Cnt
Exit Function

exit_LinearNETAlignmentSelectUMCToUse:
Debug.Assert False
LogErrors Err.Number, "Misc.bas->LinearNETAlignmentSelectUMCToUse"

End Function

Private Sub LinearNETAlignmentSelectUMCsToUseWork(ByVal lngGelIndex As Long, _
                                                 ByVal blnRequireDispersed As Boolean, _
                                                 ByRef UseUMC() As Boolean, _
                                                 ByRef lngUMCCntAddedSinceLowSegmentCount As Long, _
                                                 ByRef lngUMCSegmentCntWithLowUMCCnt As Long)
                                                 
    ' If blnRequireDispersed = True, then assures that the selected LC-MS Features are representative of all parts of the data
    
    Dim Abu() As Double         ' 0-based array; abundances to sort
    Dim TmpInd() As Long        ' 0-based array; original indices in GelUMC()
    Dim qsd As New QSDouble
    
    Dim udtSegmentStats() As udtSegmentStatsType       ' 0-based array
    
    Dim UMCTopAbuPctCnt As Long
    Dim UMCMinimumCntPerSegment As Long
    
    Dim lngIndex As Long, lngSegmentIndex As Long
    Dim lngMaxUnusedUMCs As Long
    Dim lngUMCIndex As Long
    Dim lngMatchingIndices() As Long
    Dim lngMatchCount As Long
    Dim lngUMCsSelectedBeforeSegmentChecking As Long
    Dim lngUMCsSelectedAfterSegmentChecking As Long
    
    Dim lngScanMinAdjusted As Long, lngScanMaxAdjusted As Long
    Dim lngScansPerSegment As Long
    Dim lngScanCenter As Long, lngWorkingScan As Long
    
    Dim ScanMin As Long
    Dim ScanMax As Long
    Dim ScanRange As Long
    
    Dim lngSegmentBin As Long
    Dim lngSegmentCount As Long
    
    Dim ePairedSearchUMCSelection As punaPairsUMCNetAdjustmentConstants
    Dim blnAddThisUMC As Boolean

    Dim objP1IndFastSearch As FastSearchArrayLong
    Dim objP2IndFastSearch As FastSearchArrayLong

On Error GoTo SelectUMCsToUseWorkErrorHandler

    GetScanRange lngGelIndex, ScanMin, ScanMax, ScanRange

    UMCTopAbuPctCnt = CLng((UMCNetAdjDef.TopAbuPct / 100) * GelUMC(lngGelIndex).UMCCnt)
    
    ' First select the LC-MS Features to use
    ' What we do here is set the UseUMC() flag for the low abundance LC-MS Features to false
    ' We do not take the pairing preferences into account when we do this; we will
    '  consider that below if blnRequireDispersed = True
    ReDim Abu(GelUMC(lngGelIndex).UMCCnt - 1)
    ReDim TmpInd(GelUMC(lngGelIndex).UMCCnt - 1)
    For lngIndex = 0 To GelUMC(lngGelIndex).UMCCnt - 1
        With GelUMC(lngGelIndex).UMCs(lngIndex)
            Abu(lngIndex) = .ClassAbundance
            TmpInd(lngIndex) = lngIndex
        End With
    Next lngIndex
    
    If qsd.QSDesc(Abu(), TmpInd()) Then
       If UMCTopAbuPctCnt > GelUMC(lngGelIndex).UMCCnt Then UMCTopAbuPctCnt = GelUMC(lngGelIndex).UMCCnt
       If UMCTopAbuPctCnt < 0 Then UMCTopAbuPctCnt = 0
       For lngIndex = UMCTopAbuPctCnt To GelUMC(lngGelIndex).UMCCnt - 1
           UseUMC(TmpInd(lngIndex)) = False
       Next lngIndex
    End If

    lngUMCCntAddedSinceLowSegmentCount = 0
    lngUMCSegmentCntWithLowUMCCnt = 0

    If blnRequireDispersed Then
        ' Collect stats on the number of LC-MS Features used per segment
        
        With glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions
            lngSegmentCount = .SegmentCount
            If lngSegmentCount < 1 Then lngSegmentCount = 1
            
            lngScanMinAdjusted = ScanMin + (ScanMax - ScanMin) * (.ScanPctStart / 100)
            lngScanMaxAdjusted = ScanMin + (ScanMax - ScanMin) * (.ScanPctEnd / 100)
            
            If lngScanMinAdjusted < ScanMin Then lngScanMinAdjusted = ScanMin
            If lngScanMaxAdjusted > ScanMax Then lngScanMaxAdjusted = ScanMax
            
        End With
        
        ReDim udtSegmentStats(0 To lngSegmentCount - 1)
        
        ' Determine the total number of unused LC-MS Features
        lngMaxUnusedUMCs = 0
        For lngUMCIndex = 0 To GelUMC(lngGelIndex).UMCCnt - 1
            If Not UseUMC(lngUMCIndex) Then lngMaxUnusedUMCs = lngMaxUnusedUMCs + 1
        Next lngUMCIndex
        
        If lngMaxUnusedUMCs < 10 Then lngMaxUnusedUMCs = 10
        For lngSegmentIndex = 0 To lngSegmentCount - 1
            With udtSegmentStats(lngSegmentIndex)
                .ArrayCountUnused = 0
                ReDim .UnusedUMCIndices(lngMaxUnusedUMCs)
            End With
        Next lngSegmentIndex
        
        lngScansPerSegment = (lngScanMaxAdjusted - lngScanMinAdjusted) / lngSegmentCount
        
        For lngIndex = 0 To GelUMC(lngGelIndex).UMCCnt - 1
            With GelUMC(lngGelIndex).UMCs(lngIndex)
                ' Compute the center scan of this UMC
                lngScanCenter = (.MaxScan + .MinScan) / 2
            End With
            
            If lngScanCenter >= lngScanMinAdjusted And lngScanCenter <= lngScanMaxAdjusted Then
                ' Determine which segment this scan corresponds to
                
                ' First subtract lngScanMinAdjusted from lngScanCenter
                ' For example, if lngScanMinAdjusted is 100 and lngScanCenter is 250, then lngWorkingScan = 150
                lngWorkingScan = lngScanCenter - lngScanMinAdjusted
                
                ' Now, dividing lngWorkingScan by lngScansPerSegment and rounding to the nearest integer
                '  actually gives the bin
                ' For example, given lngWorkingScan = 150 and lngScansPerSegment = 1000, Bin = CLng(150/1000) = 0
                lngSegmentBin = CLng(lngWorkingScan / lngScansPerSegment)
                
                If lngSegmentBin < 0 Then lngSegmentBin = 0
                If lngSegmentBin >= lngSegmentCount Then lngSegmentBin = lngSegmentCount - 1
                
                If UseUMC(lngIndex) Then
                    udtSegmentStats(lngSegmentBin).UMCHitCountUsed = udtSegmentStats(lngSegmentBin).UMCHitCountUsed + 1
                Else
                    ' Add to the array of potential LC-MS Features that could be added if needed
                    With udtSegmentStats(lngSegmentBin)
                        .UnusedUMCIndices(.ArrayCountUnused) = lngIndex
                        .ArrayCountUnused = .ArrayCountUnused + 1
                    End With
                End If
            Else
                ' UMC is outside the desired scan range; ignore it
            End If
        Next lngIndex
        
        For lngSegmentIndex = 0 To lngSegmentCount - 1
            With udtSegmentStats(lngSegmentIndex)
                UMCMinimumCntPerSegment = (glbPreferencesExpanded.NetAdjustmentUMCDistributionOptions.MinimumUMCsPerSegmentPctTopAbuPct / 100#) * (UMCNetAdjDef.TopAbuPct / 100#) * (.UMCHitCountUsed + .ArrayCountUnused)
                lngUMCsSelectedBeforeSegmentChecking = lngUMCsSelectedBeforeSegmentChecking + .UMCHitCountUsed
                If .UMCHitCountUsed < UMCMinimumCntPerSegment Then
                    ' Hit count is too low
                    lngUMCSegmentCntWithLowUMCCnt = lngUMCSegmentCntWithLowUMCCnt + 1
                    
                    ' If .ArrayCountUnused is more than 0, then fill the Abu() and TmpInd() arrays
                    '  and sort by abundance (ascending)
                    If .ArrayCountUnused > 0 Then
                        
                        ReDim Abu(.ArrayCountUnused - 1)
                        ReDim TmpInd(.ArrayCountUnused - 1)
                        
                        For lngIndex = 0 To .ArrayCountUnused - 1
                            Abu(lngIndex) = GelUMC(lngGelIndex).UMCs(.UnusedUMCIndices(lngIndex)).ClassAbundance
                            TmpInd(lngIndex) = .UnusedUMCIndices(lngIndex)
                        Next lngIndex
                        
                        If qsd.QSAsc(Abu(), TmpInd()) Then
                            ' Add back in the necessary number of LC-MS Features, taking into
                            '   account the value of ePairedSearchUMCSelection and

                            ePairedSearchUMCSelection = glbPreferencesExpanded.PairSearchOptions.NETAdjustmentPairedSearchUMCSelection
                            
                            If ePairedSearchUMCSelection <> punaPairedAndUnpaired Then
                                ' Initialize the PairIndex lookup objects
                                If Not PairIndexLookupInitialize(lngGelIndex, objP1IndFastSearch, objP2IndFastSearch) Then
                                    ' No pairs found; pretend we're including all LC-MS Features
                                    ePairedSearchUMCSelection = punaPairedAndUnpaired
                                End If
                            End If
                            
                            lngIndex = 0
                            Do While .UMCHitCountUsed < UMCMinimumCntPerSegment And lngIndex < .ArrayCountUnused
                                
                                lngUMCIndex = TmpInd(lngIndex)
                                blnAddThisUMC = False
                                
                                Select Case ePairedSearchUMCSelection
                                Case punaPairedAndUnpaired, punaPairedAll
                                    ' Add no matter what
                                    blnAddThisUMC = True
                                Case punaPairedLight
                                    ' Only add if UMC is the light member of a pair
                                    If objP1IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        ' Match Found
                                        blnAddThisUMC = True
                                    End If
                                Case punaPairedHeavy
                                    ' Only add if UMC is the light member of a pair
                                    If objP2IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        ' Match Found
                                        blnAddThisUMC = True
                                    End If
                                Case punaUnpairedOnly
                                    ' Only add if UMC does not belong to a pair
                                    If Not objP1IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        If Not objP2IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                            ' No Match Found; use this UMC
                                            blnAddThisUMC = True
                                        End If
                                    End If
                                Case punaUnpairedPlusPairedLight
                                    ' Only add if UMC does not belong to heavy members of pairs
                                    If Not objP2IndFastSearch.FindMatchingIndices(lngUMCIndex, lngMatchingIndices(), lngMatchCount) Then
                                        ' No Match Found; use this UMC
                                        blnAddThisUMC = True
                                    End If
                                End Select
                            
                                If blnAddThisUMC Then
                                    ' Make sure the UMC passes the Minimum UMC Member count and
                                    '  Minimum Scan Range requirements, if necessary
                                    
                                    blnAddThisUMC = LinearNETAlignmentUMCSelectionFilterCheck(lngGelIndex, lngUMCIndex)
                                    
                                    If blnAddThisUMC Then
                                        UseUMC(lngUMCIndex) = True
                                        ' Turn on the LowSegmentCountAddedUMC bit
                                        GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassStatusBits = GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassStatusBits Or UMC_INDICATOR_BIT_LOWSEGMENTCOUNT_ADDITION
                                        lngUMCCntAddedSinceLowSegmentCount = lngUMCCntAddedSinceLowSegmentCount + 1
                                        .UMCHitCountUsed = .UMCHitCountUsed + 1
                                    End If
                                End If
                                
                                lngIndex = lngIndex + 1
                            Loop
                            
                        End If
                        
                    End If
                End If
                lngUMCsSelectedAfterSegmentChecking = lngUMCsSelectedAfterSegmentChecking + .UMCHitCountUsed
            End With
        Next lngSegmentIndex
    End If

Exit Sub

SelectUMCsToUseWorkErrorHandler:
Debug.Assert False
LogErrors Err.Number, "Misc.bas->SelectUMCToUseWork"

End Sub

Private Function LinearNETAlignmentUMCSelectionFilterCheck(ByVal lngGelIndex As Long, ByVal lngUMCIndex As Long) As Boolean
    
    Dim blnValidUMC As Boolean
    
    blnValidUMC = True
    With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
        ' Filter out all mass classes with insufficient membership
        If UMCNetAdjDef.MinUMCCount > 1 Then
            If .ClassCount < UMCNetAdjDef.MinUMCCount Then
                blnValidUMC = False
            End If
        End If
        
        ' Filter out all mass classes with insufficient scan coverage
        If UMCNetAdjDef.MinScanRange > 1 Then
            If (.MaxScan - .MinScan + 1) < UMCNetAdjDef.MinScanRange Then
                blnValidUMC = False
            End If
        End If
        
        ' Filter out LC-MS Features that are not of the correct charge state
        If Not LinearNETAlignmentIsOKChargeState(.ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge) Then
            blnValidUMC = False
        End If
    End With

    LinearNETAlignmentUMCSelectionFilterCheck = blnValidUMC

End Function

Private Function LinearNETAlignmentIsOKChargeState(ByVal CS As Long) As Boolean
    '--------------------------------------------------------------
    'returns True if charge state is acceptable to current criteria
    '--------------------------------------------------------------
    On Error Resume Next
    If UMCNetAdjDef.PeakCSSelection(7) Then
       LinearNETAlignmentIsOKChargeState = True
    Else
       If CS >= 7 Then
          LinearNETAlignmentIsOKChargeState = UMCNetAdjDef.PeakCSSelection(6)
       Else
          LinearNETAlignmentIsOKChargeState = UMCNetAdjDef.PeakCSSelection(CS - 1)
       End If
    End If
End Function


' MonroeMod: Added optional param for including Gel Index
'            If included, then error will be logged to the AnalysisHistory
Public Sub LogErrors(en As Long, procName As String, _
                     Optional ByVal Description As String, Optional ByVal lngGelIndex As Long = 0, Optional blnShowErrorLoggingErrorsDialog As Boolean = True)
'------------------------------------------------------------
'logs specified error with name of procedure where it occured
'if Len(Description)>0 then Description is accepted as actual
'error (so we can handle application defined errors
'------------------------------------------------------------
Dim nFN As Integer
Dim sTmp As String
Dim sPathToLog As String
On Error GoTo err_logerrors
If IsMissing(Description) Then Description = ""
sPathToLog = AddDirSeparator(App.Path) & "UnexpErr.log"
sTmp = "Error event in procedure: " & procName
If Len(Description) > 0 Then
   sTmp = sTmp & vbCrLf & "Error #:" & en & " Description:" & Description
Else
   sTmp = sTmp & vbCrLf & "Error #:" & en & " Description:" & Error(en)
End If
sTmp = "-----" & Now & "-----" & vbCrLf & sTmp & vbCrLf
nFN = FreeFile
Open sPathToLog For Append As nFN
     Print #nFN, sTmp
Close nFN

' MonroeMod Start
If lngGelIndex >= 1 Then
    On Error Resume Next
    AddToAnalysisHistory lngGelIndex, sTmp
End If

TraceLog 10, procName, "Error occurred: " & Error(en)
' MonroeMod Finish

Exit Sub
err_logerrors:
Debug.Assert False
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled And blnShowErrorLoggingErrorsDialog Then
    MsgBox "Error logging errors? Isn't that ironic? Error occurred in " & procName & vbCrLf & Err.Description, vbOKOnly, glFGTU
End If
End Sub

Public Function GoPrint(ByVal sTag As String) As Integer
On Error GoTo err_GoPrint
frmPrint.Tag = sTag
frmPrint.Show vbModal
GoPrint = 0
Exit Function
     
err_GoPrint:
GoPrint = -1
End Function

'  Unused Functions
'''Public Sub GetpICooSysRange(ByVal Ind As Long, X1 As Double, x2 As Double, Y1 As Double, Y2 As Double, ByVal XOrient As Integer, ByVal YOrient As Integer)
'''With GelData(Ind)
'''    Select Case XOrient
'''    Case glNormal
'''         X1 = .DFPI(UBound(.DFPI))
'''         x2 = .DFPI(1)
'''    Case glReverse
'''         X1 = .DFPI(1)
'''         x2 = .DFPI(UBound(.DFPI))
'''    End Select
'''    Select Case YOrient
'''    Case glNormal
'''         Y1 = .minMW
'''         Y2 = .maxMW
'''    Case glReverse
'''         Y1 = .maxMW
'''         Y2 = .minMW
'''    End Select
'''End With
'''End Sub
'''
'''
'''Public Sub GetFNCooSysRange(ByVal Ind As Long, X1 As Double, x2 As Double, Y1 As Double, Y2 As Double, ByVal XOrient As Integer, ByVal YOrient As Integer)
'''With GelData(Ind)
'''    Select Case XOrient
'''    Case glNormal
'''         X1 = .ScanInfo(UBound(.ScanInfo)).ScanNumber
'''         x2 = .ScanInfo(1).ScanNumber
'''    Case glReverse
'''         X1 = .ScanInfo(1).ScanNumber
'''         x2 = .ScanInfo(UBound(.ScanInfo)).ScanNumber
'''    End Select
'''    Select Case YOrient
'''    Case glNormal
'''         Y1 = .minMW
'''         Y2 = .maxMW
'''    Case glReverse
'''         Y1 = .maxMW
'''         Y2 = .minMW
'''    End Select
'''End With
'''End Sub

Public Sub PrintFileInfo(Ind As Long, InfoType As Integer)
Dim aInfo As Variant
Dim iLinesCount As Integer
Dim i As Integer
On Error GoTo err_printfileinfo

aInfo = Fileinfo(Ind, InfoType)
If ((IsNull(aInfo)) Or (Not IsArray(aInfo))) Then
   MsgBox "No file info available.", vbOKOnly
   Exit Sub
Else
   iLinesCount = UBound(aInfo)
   Printer.Orientation = vbPRORPortrait
   Printer.ScaleLeft = -200
   Printer.ScaleTop = -200
   Printer.CurrentX = 0
   Printer.CurrentY = 0
   Printer.Font = frmDataInfo.rtbData.Font
   Printer.Font.Size = 10
   Printer.Font.Bold = True
   Printer.Print "File Info for: " & GelBody(Ind).Caption
   Printer.Font.Bold = False
   Printer.Print " "
   For i = 0 To iLinesCount
      Printer.Print aInfo(i)
   Next i
   Printer.Print
   Printer.Print "COMMENT"
   Printer.Print GelData(Ind).Comment
   Printer.EndDoc
   Printer.ScaleLeft = 0
   Printer.ScaleTop = 0
End If
Exit Sub

err_printfileinfo:
MsgBox "Error printing file information." & vbCrLf & sErrLogReference
LogErrors Err.Number, "PrintFileInfo"
End Sub

Public Function UserName() As String
Dim iUNLen As Long
Dim sUN As String * 255
Dim iCutHere As Integer
iUNLen = Len(sUN)
If GetUserName(sUN, iUNLen) <> 0 Then
   iCutHere = InStr(sUN, Chr$(0))
   If iCutHere > 1 Then
      UserName = Left$(sUN, iCutHere - 1)
   Else
      UserName = sUN
   End If
Else
   UserName = ""
End If
End Function


Public Function FactorN(ByVal X As Long, ByVal n As Long) As Long
'----------------------------------------------------------------
'returns factor of the x closest to the N
'----------------------------------------------------------------
Dim i As Long
Dim CurrFacN As Long

CurrFacN = 1
For i = 1 To CLng(X / 2)
    If X Mod i = 0 Then
       If Abs(i - n) <= Abs(CurrFacN - n) Then
          CurrFacN = i
       End If
    End If
Next i
FactorN = CurrFacN
End Function

' Unused Function (March 2003)
'''Public Function GetKeyValue(lKey As Long, KeyName As String, SubKey As String) As String
'''Dim Res As Long, i As Long
'''Dim hKey As Long
'''Dim KeyValType As Long
'''Dim KeyValSize As Long
'''Dim tmpVal As String
'''Dim tmpVal1 As String
'''Dim tmpVal2 As String
'''
'''Res = RegOpenKeyEx(lKey, KeyName, 0, KEY_QUERY_VALUE, hKey)
'''
'''If Res <> ERROR_SUCCESS Then GoTo err_GetKeyValue
'''tmpVal = String$(1024, 0)
'''KeyValSize = 1024
'''
'''Res = RegQueryValueEx(hKey, SubKey, 0, KeyValType, tmpVal, KeyValSize)
'''If Res <> ERROR_SUCCESS Then GoTo err_GetKeyValue
'''
'''If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) And KeyValType <> REG_BINARY Then
'''   tmpVal = Left(tmpVal, KeyValSize - 1)    'Win95 Adds Null Terminated string
'''Else
'''   tmpVal = Left(tmpVal, KeyValSize)        'WinNt doesnt Null terminate
'''End If
'''
'''Select Case KeyValType
'''Case REG_SZ
'''     GetKeyValue = tmpVal
'''Case REG_BINARY
'''     tmpVal1 = ""
'''     For i = 1 To KeyValSize
'''         tmpVal2 = Hex(Asc(Mid(tmpVal, i, 1)))
'''         If Len(tmpVal2) = 1 Then tmpVal2 = "0" & tmpVal2
'''         tmpVal1 = tmpVal1 & " " & tmpVal2
'''     Next
'''     GetKeyValue = Trim(tmpVal1)
'''Case REG_DWORD
'''     tmpVal1 = ""
'''     For i = Len(tmpVal) To 1 Step -1
'''         tmpVal1 = tmpVal1 & Hex(Asc(Mid(tmpVal, i, 1)))
'''     Next
'''     GetKeyValue = Format$("&h" + tmpVal1)
'''End Select
'''Res = RegCloseKey(hKey)
'''Exit Function
'''
'''err_GetKeyValue:
'''GetKeyValue = ""
'''Res = RegCloseKey(hKey)
'''End Function

Public Function WriteRawDataFile(ByVal Ind As Long) As Boolean
Dim sFileName As String
Dim FileNum As Integer
Dim lElements As Long, L As Long
Dim sTmp As String, stmp1 As String

On Error GoTo exit_WriteRawDataFile

WriteRawDataFile = False
sFileName = GetTempFolder() & RawDataTmpFile
FileNum = FreeFile
Open sFileName For Output As #FileNum

On Error Resume Next
With GelData(Ind)
    Print #FileNum, "Data dump from gel: " & GelBody(Ind).Caption & vbCrLf
    Print #FileNum, "(to see complete information use .PEK file)" & vbCrLf
    'get data file names, indexes,pI numbers, frequency shifts, intensities
    Print #FileNum, "--Index---File #--------pI---Freq.Sh-Intensity-----File Name"
    lElements = UBound(.ScanInfo)
    If lElements > 0 Then
        For L = 1 To lElements
            sTmp = ""
            stmp1 = CStr(L)
            sTmp = sTmp & Space(7 - Len(stmp1)) & stmp1
            stmp1 = CStr(.ScanInfo(L).ScanNumber)
            sTmp = sTmp & Space(9 - Len(stmp1)) & stmp1
            stmp1 = Format$(.ScanInfo(L).ScanPI, "0.0000")
            sTmp = sTmp & Space(10 - Len(stmp1)) & stmp1
            stmp1 = Format$(.ScanInfo(L).FrequencyShift, "0.0000")
            sTmp = sTmp & Space(10 - Len(stmp1)) & stmp1
            stmp1 = Format$(.ScanInfo(L).TimeDomainSignal, "0.0000")
            sTmp = sTmp & Space(10 - Len(stmp1)) & stmp1
            sTmp = sTmp & Space(5) & .ScanInfo(L).ScanFileName
            Print #FileNum, sTmp
        Next L
    Else
        Print #FileNum, "No data found in this block." & vbCrLf
    End If
    'get charge state data
    Print #FileNum, "[Charge State Data]"
    Print #FileNum, "--Index-File #-1stCS-CSCnt-Abundance-----------MW---StDev---Exp. DB MW---Error------ER--Identity---"
    If .CSLines > 0 Then
        For L = 1 To .CSLines
            sTmp = ""
            stmp1 = CStr(L)
            sTmp = sTmp & Space(7 - Len(stmp1)) & stmp1
            stmp1 = CStr(.CSData(L).ScanNumber)
            sTmp = sTmp & Space(7 - Len(stmp1)) & stmp1
            stmp1 = CStr(.CSData(L).Charge)
            sTmp = sTmp & Space(6 - Len(stmp1)) & stmp1
            stmp1 = CStr(.CSData(L).ChargeCount)
            sTmp = sTmp & Space(6 - Len(stmp1)) & stmp1
            stmp1 = Format$(.CSData(L).Abundance, "Scientific")
            sTmp = sTmp & Space(10 - Len(stmp1)) & stmp1
            stmp1 = Format$(.CSData(L).AverageMW, "0.0000")
            sTmp = sTmp & Space(13 - Len(stmp1)) & stmp1
            stmp1 = Format$(.CSData(L).MassStDev, "Standard")
            sTmp = sTmp & Space(8 - Len(stmp1)) & stmp1
''            stmp1 = Format$(.CSData(L).IsotopicFitRatio, "0.0000")
''            sTmp = sTmp & Space(13 - Len(stmp1)) & stmp1
''            stmp1 = Format$(.CSData(L).IsotopicAtomCount, "Standard")
''            sTmp = sTmp & Space(8 - Len(stmp1)) & stmp1
''            If Not IsNull(.CSVar(L, csvfMTDDRatio)) Then
''               If IsNumeric(.CSVar(L, csvfMTDDRatio)) Then
''                  stmp1 = Format$(.CSVar(L, csvfMTDDRatio), "Standard")
''               Else
''                  stmp1 = CStr(.CSVar(L, csvfMTDDRatio))
''               End If
''               sTmp = sTmp & Space(8 - Len(stmp1)) & stmp1
''            Else
''               sTmp = sTmp & Space(8)
''            End If
            If Len(.CSData(L).MTID) > 0 Then
               stmp1 = CStr(.CSData(L).MTID)
               If Len(stmp1) > 20 Then
                  sTmp = sTmp & vbCrLf & "Identity:" & stmp1
               Else
                  sTmp = sTmp & Space(2) & stmp1
               End If
            End If
            Print #FileNum, sTmp
        Next L
        Print #FileNum, vbCrLf
    Else
        Print #FileNum, "No data found in this block." & vbCrLf
    End If
    
    'get isotopic data
    Print #FileNum, "[Isotopic Data]"
    Print #FileNum, "--Index-File #----CS-Abundance----------m/z-----Fit---Average MW--Monoiso. MW-Most Abu. MW---Exp. DB MW---Error------ER--Identity"
    
    If .IsoLines > 0 Then
        For L = 1 To .IsoLines
            sTmp = ""
            stmp1 = CStr(L)
            sTmp = sTmp & Space(7 - Len(stmp1)) & stmp1
            stmp1 = CStr(.IsoData(L).ScanNumber)
            sTmp = sTmp & Space(7 - Len(stmp1)) & stmp1
            stmp1 = CStr(.IsoData(L).Charge)
            sTmp = sTmp & Space(6 - Len(stmp1)) & stmp1
            stmp1 = Format$(.IsoData(L).Abundance, "Scientific")
            sTmp = sTmp & Space(10 - Len(stmp1)) & stmp1
            stmp1 = Format$(.IsoData(L).MZ, "0.0000")
            sTmp = sTmp & Space(13 - Len(stmp1)) & stmp1
            stmp1 = Format$(.IsoData(L).Fit, "Standard")
            sTmp = sTmp & Space(8 - Len(stmp1)) & stmp1
            stmp1 = Format$(.IsoData(L).AverageMW, "0.0000")
            sTmp = sTmp & Space(13 - Len(stmp1)) & stmp1
            stmp1 = Format$(.IsoData(L).MonoisotopicMW, "0.0000")
            sTmp = sTmp & Space(13 - Len(stmp1)) & stmp1
            stmp1 = Format$(.IsoData(L).MostAbundantMW, "0.0000")
            sTmp = sTmp & Space(13 - Len(stmp1)) & stmp1
''            stmp1 = Format$(.IsoData(L).IsotopicFitRatio, "0.0000")
''            sTmp = sTmp & Space(13 - Len(stmp1)) & stmp1
''            stmp1 = Format$(.IsoData(L).IsotopicAtomCount, "Standard")
''            sTmp = sTmp & Space(8 - Len(stmp1)) & stmp1
''            If Not IsNull(.IsoVar(L, isvfMTDDRatio)) Then
''               If IsNumeric(.IsoVar(L, isvfMTDDRatio)) Then
''                  stmp1 = Format$(.IsoVar(L, isvfMTDDRatio), "Standard")
''               Else
''                  stmp1 = CStr(.IsoVar(L, isvfMTDDRatio))
''               End If
''               sTmp = sTmp & Space(8 - Len(stmp1)) & stmp1
''            Else
               sTmp = sTmp & Space(8)
''            End If
            If Len(.IsoData(L).MTID) > 0 Then
               stmp1 = .IsoData(L).MTID
               If Len(stmp1) > 20 Then
                  sTmp = sTmp & vbCrLf & "Identity:" & stmp1
               Else
                  sTmp = sTmp & Space(2) & stmp1
               End If
            End If
            Print #FileNum, sTmp
        Next L
        Print #FileNum, vbCrLf
    Else
        Print #FileNum, "No data found in this block." & vbCrLf
    End If
End With
Print #FileNum, "End of data."

Close #FileNum
WriteRawDataFile = True

exit_WriteRawDataFile:
End Function

Public Sub PrintData1(ByVal Ind As Long)
Dim Res As Long
Dim sFileName As String
Dim hWndDsk As Long
If WriteRawDataFile(Ind) Then
   sFileName = AddDirSeparator(App.Path) & RawDataTmpFile
   hWndDsk = GetDesktopWindow()
   Res = ShellExecute(hWndDsk, "Print", sFileName, 0&, 0&, SW_HIDE)
Else
   MsgBox "Error printing data from file " & GelBody(Ind).Caption, vbOKOnly
End If
End Sub

Public Function SuggestionByName(ByVal Inspiration As String, ByVal Extension As String) As String
'returns name based on Inspiration and Extension
Dim tmpName As String
Dim DotPos As Integer
tmpName = GetFileNameOnly(Inspiration)

DotPos = InStr(1, tmpName, ".")
If DotPos > 0 Then
   SuggestionByName = Left$(tmpName, DotPos) & Extension
Else
   SuggestionByName = tmpName & "." & Extension
End If
End Function

Public Function SuggestionByIndex(ByVal Inspiration As Integer, ByVal Extension As String) As String
'returns name based on Inspiration and Extension
Dim tmpName As String
Dim DotPos As Integer
tmpName = GetFileNameOnly(GelData(Inspiration).FileName)
DotPos = InStr(1, tmpName, ".")
If DotPos > 0 Then
   SuggestionByIndex = Left$(tmpName, DotPos) & Extension
Else
   SuggestionByIndex = tmpName & "." & Extension
End If
End Function

Public Function IsArrayEmpty(aArray As Variant) As Boolean
'check if dynamic array is empty
On Error Resume Next
IsArrayEmpty = UBound(aArray)
IsArrayEmpty = Err
End Function

Public Sub GetColorAPIDlg(ByVal Ownerhwnd As Long, _
                          ThingColor As Long)
Dim ChClr As ChooseColor
Dim CustColor(15) As Long
Dim i As Integer
Dim Res As Long

For i = 0 To 15
    CustColor(i) = GetSysColor(i)
Next i

ChClr.lStructSize = Len(ChClr)
ChClr.hwndOwner = Ownerhwnd
ChClr.rgbResult = ThingColor
ChClr.hInstance = 0
ChClr.lpCustColors = VarPtr(CustColor(0))
ChClr.flags = 0
Res = ChooseColor(ChClr)
If Res = 1 Then
   ' Update ThingColor
   ThingColor = ChClr.rgbResult
Else
   ' Didn't return 1; leave ThingColor unchanged
End If

End Sub

' Unused sub; replaced with GetIsoOrCSDataByField
''Public Sub GetFieldIndexes(ByVal Ind As Long, _
''                           ByVal Field As Integer, _
''                           ByRef CSField As Integer, _
''                           ByRef IsoField As Integer)
'''sets CSField and IsoField to appropriate indexes of fields in
'''Num, Var arrays - this is like helper routine used in many functions
''End Sub

Public Function GetIsoOrCSDataByField(Ind, udtData As udtIsotopicDataType, Field As Integer, blnIsCSData As Boolean) As Double
    
    Dim lngCount As Long
    Dim strRefs() As String     ' 1-based array
    Dim dblValue As Double
    
    Select Case Field
    Case glFIELD_MW
        If blnIsCSData Then
            dblValue = udtData.AverageMW
        Else
            dblValue = GetIsoMass(udtData, GelData(Ind).Preferences.IsoDataField)
        End If
    Case glFIELD_MOVERZ
        dblValue = udtData.MZ
    Case glFIELD_CS
        dblValue = udtData.Charge
    Case glFIELD_ABU
        dblValue = udtData.Abundance
    Case glFIELD_ER
        dblValue = udtData.ExpressionRatio
    Case glFIELD_ID
        lngCount = GetAMTRefFromString1(udtData.MTID, strRefs)
         If lngCount > 0 Then
            On Error Resume Next
            dblValue = CLng(strRefs(1))
        Else
            dblValue = 0
        End If
    Case glFIELD_FIT
        dblValue = udtData.Fit
    Case Else
        dblValue = 0
    End Select
    
    GetIsoOrCSDataByField = dblValue
End Function

Public Function GetIDFromString(ByVal S As String, _
                                ByVal Mark As String, _
                                Optional Terminator As Variant) As String
'retrieves substring of s on position starting with Mark + Len(Mark)
'to first Terminator appearance; if Terminator is not provided then
'it goes after numeric values and stop with first non-numeric character
Dim MarkPos As Long
Dim IDStart As Long
Dim IDPos As Long
Dim sTmp As String
Dim sChar As String
Dim Done As Boolean
sTmp = ""
MarkPos = InStr(1, S, Mark)
If MarkPos > 0 Then
   If IsMissing(Terminator) Then
      IDPos = MarkPos + Len(Mark)
      Do Until (Done Or (IDPos > Len(S)))
         sChar = Mid$(S, IDPos, 1)
         If IsNumeric(sChar) Then
            sTmp = sTmp & sChar
            IDPos = IDPos + 1
         Else
            Done = True
         End If
      Loop
   Else
      IDStart = MarkPos + Len(Mark)
      If IDStart <= Len(S) Then
         IDPos = InStr(IDStart, S, Terminator)
         If IDPos > 0 Then sTmp = Mid$(S, IDStart, IDPos - IDStart)
      End If
   End If
End If
GetIDFromString = sTmp
End Function

Public Sub SetDefaultUMCDef(ByRef udtUMCDef As UMCDefinition)
With udtUMCDef
    .UMCType = glUMC_TYPE_INTENSITY
    .DefScope = glScope.glSc_Current         ' Current view only
    .MWField = mftMWMono
    .TolType = gltPPM
    .Tol = 12.5
    .UMCSharing = False
    .UMCUniCS = False                   ' Not yet implemented
    .ClassAbu = UMCClassAbundanceConstants.UMCAbuSum
    .ClassMW = UMCClassMassConstants.UMCMassMed
    .GapMaxCnt = 10
    .GapMaxSize = 5                   ' Applies to "Maximum size of Scan Gap" in UMC2003; applies to SplitUMCs in UMCIonNet
    .GapMaxPct = 0.8
    .UMCNETType = UMCNetConstants.UMCNetAt
    
''    .UMCMaxAbuEtPctAf = -10           'Ignored
''    .UMCMaxAbuEtPctBf = -5            'Ignored
''    .UMCMaxAbuPctAf = -10             'Ignored
''    .UMCMaxAbuPctBf = -10             'Ignored

    .OddEvenProcessingMode = oepUMCOddEvenProcessingMode.oepProcessAll
    .RequireMatchingIsotopeTag = True
    .AdditionalValue2 = 0
    .AdditionalValue3 = 0
    .AdditionalValue4 = 0
    .AdditionalValue5 = 0
    .AdditionalValue6 = 0
    .AdditionalValue7 = 0
    .AdditionalValue8 = 0
        
    .UMCMinCnt = 3
    .UMCMaxCnt = 100
    .InterpolateGaps = True
    .InterpolateMaxGapSize = 4
    .InterpolationType = 0            ' Currently only one interpolation method: 0
    
    .ChargeStateStatsRepType = UMCChargeStateGroupConstants.UMCCSGHighestSum
    .UMCClassStatsUseStatsFromMostAbuChargeState = True
End With
End Sub

Public Sub SetDefaultUMCIonNetDef(ByRef udtUMCIonNetDef As UMCIonNetDefinition)
    Dim intIndex As Integer
    
    ' These defaults were set in November 2006
    ' See SetOldDefaultUMCIonNetDef for the previous defaults
    With udtUMCIonNetDef
        .MetricType = METRIC_EUCLIDEAN
        .NETType = Net_SPIDER_66
        .NetDim = 5
        .TooDistant = 0.1
        ReDim .MetricData(.NetDim - 1)
        .MetricData(0).Use = True:  .MetricData(0).DataType = uindUMCIonNetDimConstants.uindMonoMW:   .MetricData(0).WeightFactor = 0.01:   .MetricData(0).ConstraintType = Net_CT_LT:       .MetricData(0).ConstraintValue = 10: .MetricData(0).ConstraintUnits = DATA_UNITS_MASS_PPM
        .MetricData(1).Use = False:  .MetricData(1).DataType = uindUMCIonNetDimConstants.uindAvgMW:    .MetricData(1).WeightFactor = 0.01:   .MetricData(1).ConstraintType = Net_CT_LT:       .MetricData(1).ConstraintValue = 10: .MetricData(1).ConstraintUnits = DATA_UNITS_MASS_PPM
        .MetricData(2).Use = True:  .MetricData(2).DataType = uindUMCIonNetDimConstants.uindLogAbundance:   .MetricData(2).WeightFactor = 0.1:   .MetricData(2).ConstraintType = Net_CT_None:     .MetricData(2).ConstraintValue = 0.1:   .MetricData(2).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(3).Use = True:  .MetricData(3).DataType = uindUMCIonNetDimConstants.uindGenericNET:      .MetricData(3).WeightFactor = 15:   .MetricData(3).ConstraintType = Net_CT_None:    .MetricData(3).ConstraintValue = 0.01:  .MetricData(3).ConstraintUnits = DATA_UNITS_MASS_DA
        .MetricData(4).Use = True:  .MetricData(4).DataType = uindUMCIonNetDimConstants.uindFit:       .MetricData(4).WeightFactor = 0.1:    .MetricData(4).ConstraintType = Net_CT_None:    .MetricData(4).ConstraintValue = 0.01:  .MetricData(4).ConstraintUnits = DATA_UNITS_MASS_DA
    
        .NetActualDim = 0
        For intIndex = 0 To .NetDim - 1
            If .MetricData(intIndex).Use Then
                .NetActualDim = .NetActualDim + 1
            End If
        Next intIndex
    End With
End Sub

Public Sub SetDefaultSearchAMTDef(ByRef udtAMTDef As SearchAMTDefinition, ByRef udtUMCNetAdjDef As NetAdjDefinition)
With udtAMTDef
    .SearchScope = glSc_All
    .SearchFlag = glAMT_CONFIRM_PPM     'search among all AMTs confirmed with good precision
    .MWField = mftMWMono
    .MWTol = 6
    .NETorRT = glAMT_NET
    .Formula = ConstructNETFormulaWithDefaults(udtUMCNetAdjDef)
    .TolType = gltPPM
    .NETTol = 0.025
    .MassTag = -1
    .MaxMassTags = 5
    .SkipReferenced = False
    .SaveNCnt = True
End With
End Sub

Public Sub SetDefaultUMCNETAdjDef(ByRef udtUMCNetAdjDef As NetAdjDefinition)
With udtUMCNetAdjDef
    .MaxScanPct = 10
    .MinScanRange = 3
    .MinUMCCount = 3
    .TopAbuPct = 20
    .PeakSelection = 1          ' Ignored; always select At Max = 1
    .PeakMaxAbuPct = 10         ' Ignored
    
    .MWTolType = gltPPM
    .MWTol = 10                     '10 ppm
    .InitialSlope = 0.00015
    .InitialIntercept = 0
    .NETFormula = ConstructNETFormulaWithDefaults(udtUMCNetAdjDef)
    .NETTolIterative = 0.2
    .NETorRT = glAMT_NET
    .UseNET = True
    .UseMultiIDMaxNETDist = True
    .MultiIDMaxNETDist = 0.1        '~10 pct
    .EliminateBadNET = True
    .MaxIDToUse = 5000
    .PeakCSSelection(7) = True
    .IterationStopType = 4
    .IterationStopValue = 0.0005
    .IterationUseMWDec = False
    .IterationMWDec = 2.5
    .IterationUseNETdec = True
    .IterationNETDec = 0.025
    .IterationAcceptLast = True
    
    ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
    .UseNetAdjLockers = False
    .UseOldNetAdjIfFailure = True
    .NetAdjLockerMinimumMatchCount = 3
    
    .UseRobustNETAdjustment = True
    
    If APP_BUILD_DISABLE_LCMSWARP Then
        .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETIterative
    Else
        .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass
    End If
    
    .RobustNETSlopeStart = 0.00002
    .RobustNETSlopeEnd = 0.002
    .RobustNETSlopeIncreaseMode = UMCRobustNETIncrementConstants.UMCRobustNETIncrementPercentage
    .RobustNETSlopeIncrement = 75
    
    .RobustNETInterceptStart = -0.4
    .RobustNETInterceptEnd = 0.2
    .RobustNETInterceptIncrement = 0.2
    
    .RobustNETMassShiftPPMStart = -15
    .RobustNETMassShiftPPMEnd = 15
    .RobustNETMassShiftPPMIncrement = 15
    
''    .RobustNETAnnealSteps = 20
''    .RobustNETAnnealTrialsPerStep = 250
''    .RobustNETAnnealMaxSwapsPerStep = 50
''    .RobustNETAnnealTemperatureReductionFactor = 0.9
    .AdditionalValue1 = 0
    .AdditionalValue2 = 0
    .AdditionalValue3 = 0
    .AdditionalValue4 = 0
    .AdditionalValue5 = 0
    .AdditionalValue6 = 0
    
    With .MSWarpOptions
        .MassCalibrationType = rmcUMCRobustNETWarpMassCalibrationType.rmcHybridRecal
        .MinimumPMTTagObsCount = 5
        .MatchPromiscuity = 2
        
        .NETTol = 0.02
        .NumberOfSections = 100
        .MaxDistortion = 10                 ' 1/27/2006: Changed from 3 to 10
        .ContractionFactor = 3              ' 1/27/2006: Changed from 2 to 3
    
        .MassWindowPPM = 50
        .MassSplineOrder = 2
        .MassNumXSlices = 20
        .MassNumMassDeltaBins = 100
        .MassMaxJump = 50
        
        .MassZScoreTolerance = 3
        .MassUseLSQ = True
        .MassLSQOutlierZScore = 3
        .MassLSQNumKnots = 12
        
        .AdditionalValue1 = 0
        .AdditionalValue2 = 0
        .AdditionalValue3 = 0
        .AdditionalValue4 = 0
        .AdditionalValue5 = 0
        .AdditionalValue6 = 0
        .AdditionalValue7 = 0
        .AdditionalValue8 = 0
        .AdditionalValue9 = 0
        .AdditionalValue10 = 0
    End With
    
    .OtherInfo = ""
End With

End Sub

Public Sub PopulateDefaultInternalStds(ByRef udtNetAdjLockers As udtInternalStandardsType)
    ' The default Internal Standards are the original list of NET Locker peptides
    
    With udtNetAdjLockers
''        .Count = 5
''        ReDim .InternalStandards(0 To 4)
''        DefineNetAdjLocker .InternalStandards(0), "1000", "ASHLGLAR", 823.4664, 0.215, 1, 2, 1
''        DefineNetAdjLocker .InternalStandards(1), "1001", "APRTPGGRR", 966.5471, 0.139, 1, 3, 2
''        DefineNetAdjLocker .InternalStandards(2), "1002", "pEPPGGSKVILF", 1124.6229, 0.375, 1, 2, 1
''        DefineNetAdjLocker .InternalStandards(3), "1003", "INLKALAALAKKIL", 1478.0061, 0.572, 1, 3, 2
''        DefineNetAdjLocker .InternalStandards(4), "1004", "FLPLILGKLVKGLL", 1522.0374, 0.656, 2, 3, 2
        ReDim .InternalStandards(0)
        
        .Count = 0
        .StandardsAreFromDB = False
    End With

End Sub

Private Sub DefineNetAdjLocker(ByRef udtNetAdjLocker As udtInternalStandardEntryType, strDescription As String, strPeptideSequence As String, dblMonisotopicMass As Double, dblNET As Double, intChargeMin As Integer, intChargeMax As Integer, intChargeMostAbundant As Integer)
    
    If intChargeMax < intChargeMin Then
        ' This is unexpected
        Debug.Assert False
        intChargeMax = intChargeMin
    End If
    
    With udtNetAdjLocker
        .SeqID = strDescription
        .PeptideSequence = strPeptideSequence
        .MonoisotopicMass = dblMonisotopicMass
        .NET = dblNET
        .ChargeMinimum = intChargeMin
        .ChargeMaximum = intChargeMax
        .ChargeMostAbundant = intChargeMostAbundant
    End With

End Sub

Public Sub SetDefaultDefinitions()
    SetDefaultUMCDef UMCDef
    SetDefaultUMCIonNetDef UMCIonNetDef
    
    SetDefaultUMCNETAdjDef UMCNetAdjDef
    SetDefaultSearchAMTDef samtDef, UMCNetAdjDef
    
With sorfDef
    .SearchScope = 0    'all data
    .MWField = mftMWMono
    .MWTol = 25
    .MWTolType = gltPPM
    Set .Mods = New Collection
End With
With amtlmDef
    .lmScope = 0        'all data
    .lmPropagate = glLM_PROPAGATE_NO
    .lmIsoField = 7     'monoisotopic
    .lmMultiCandidates = glLM_MULTI_INTENSITY
    .lmMultiAMTHits = glLM_AMT_MULTI_FIT
    .lmSaveResults = glLM_SAVE_ORIGINAL
End With
With PairDef
    .Case2500 = True
    .Delta = glN14N15_DELTA
    .DeltaTol = 0.02
    .MultiAMTHits = 2           'try to match with number of N
    .SearchForSingles = True
    .SaveN15Singles = True
    .StopAfterEachScan = False
    .MassLockType = glLMTYPE_PAIRS_N14
End With
With OlyOptions
    .DefType = olySolo
    .DefShape = olyStick
    .DefVisible = True
    .DefColor = vbBlue
    .DefMaxNET = 1
    .DefMinNET = 0
    .DefNETAdjustment = olyNETDB_GANET
    .DefNETTol = 0.005
    .DefUniformSize = True
    .DefBoxSizeAsSpotSize = True
    .DefCurrScopeVisible = True
    .DefWithID = True
    .BackColor = vbWhite
    .ForeColor = vbBlack
    .DefMinSize = 0.0001
    .DefMaxSize = 0.025
    .DefFontHeight = 0.025
    .DefFontWidth = 0.0065
    .DefStickWidth = 0.0001
    .DefTextHeight = 0.25            '1/4 of size of text displayed on coordinate axes
    Set .GRID = New LaAutoGrid
    .GRID.LineStyle = glsDOT
    .GRID.HorzAutoMode = gamBinsCntConst
    .GRID.HorzBinsCount = 12
    .GRID.HorzGridVisible = True
    .GRID.VertAutoMode = gamBinsCntConst
    .GRID.VertBinsCount = 8
    .GRID.VertGridVisible = True
    .Orientation = OrientMWVrtETHrz
End With
With OlyJiggyOptions
    .UseMWConstraint = True
    .MWTol = 25
    .UseNetConstraint = True
    .NETTol = 0.2
    .UseAbuConstraint = False
    .AbuTol = 1
    .JiggyScope = 0
    .JiggyType = 0
    .BaseDisplayInd = -1
End With

End Sub

Public Function GetChildCount() As Long
'-------------------------------------------------------
'returns number of currently loaded child(gel) forms
'-------------------------------------------------------
Dim i As Long
Dim nCount As Long
For i = 0 To UBound(GelStatus)
    If Not GelStatus(i).Deleted Then nCount = nCount + 1
Next i
GetChildCount = nCount
End Function

' Unused Procedure (February 2005)
'Private Sub CommonErrMsg(ByVal ErrNum As Long)
'Dim Msg As String
'Select Case ErrNum
'Case 52
'     Msg = "Error accessing specified file. Check that file exists and that network drive where file is stored is accessible."
'Case Else
'End Select
'If Len(Msg) > 0 Then
'   MsgBox Msg, vbOKOnly
'End If
'End Sub

Public Function GetTagValueFromText(ByVal sText As String, _
                                    ByVal sTag As String) As String
'retrieves the value(as string) of Tag from Text; "" if not found
'assumption is that when Tag is found the rest of that line
'will be the value except maybe warning "(DO NOT EDIT THIS LINE)"
Dim TagStartPos As Integer
Dim WarningPos As Integer
Dim EOLPos As Integer
Dim TagLine As String
Dim TagValue As String

TagStartPos = InStr(1, sText, sTag)
If TagStartPos > 0 Then
   EOLPos = InStr(TagStartPos, sText, vbCrLf)
   If EOLPos Then
      TagLine = Mid$(sText, TagStartPos, EOLPos - TagStartPos)
   Else
      TagLine = Right$(sText, Len(sText) - TagStartPos + 1)
   End If
   If Len(TagLine) > Len(sTag) Then
      TagValue = Trim$(Right$(TagLine, Len(TagLine) - Len(sTag)))
      'check is there warning inside the tagvalue
      If Len(TagValue) > 0 Then
         WarningPos = InStr(1, TagValue, glCOMMENT_DO_NOT_EDIT)
         If WarningPos > 0 Then
            TagValue = Trim$(Left$(TagValue, WarningPos - 1))
         End If
      End If
      GetTagValueFromText = TagValue
   Else
      GetTagValueFromText = ""
   End If
Else
   GetTagValueFromText = ""
End If
End Function

Public Function GetTempFolder() As String
'returns path to the system temp folder terminated
Dim Buffer As String
Dim BufferSize  As Long
BufferSize = GetTempPath(0&, Buffer)
Buffer = Space(BufferSize - 1)
GetTempPath BufferSize, Buffer
If Right$(Buffer, 1) <> "\" Then Buffer = Buffer & "\"

' MonroeMod
GetTempFolder = Buffer

End Function

Public Function GetNumFormat(ByVal NumDecPos As Integer) As String
'-----------------------------------------------------------------
'returns custom numeric format with NumDecPos decimal places
'default is 4 decimal places
'-----------------------------------------------------------------
If NumDecPos > 0 And NumDecPos < 20 Then
   GetNumFormat = "0." & String(NumDecPos, "0")
ElseIf NumDecPos = 0 Then
   GetNumFormat = "0"
Else
   GetNumFormat = "0.0000"
End If
End Function


Public Function GetMyNameVersion(Optional blnIncludeAppName As Boolean = True, Optional blnIncludeAppDate As Boolean = False) As String
Dim strVersion As String

If blnIncludeAppName Then
    strVersion = App.Title & " - "
End If
strVersion = strVersion & App.major & "." & App.minor & "." & App.Revision

If blnIncludeAppDate Then
    strVersion = strVersion & ", " & APP_BUILD_DATE
End If

GetMyNameVersion = strVersion
End Function


Public Function AACount(ByVal Seq As String, _
                        ByVal aa As String) As Long
'---------------------------------------------------------
'returns number of occurences of an amino acid in a Seq
'---------------------------------------------------------
Dim AA_ANSI As Byte        'number of amino acids to count
Dim BPC As Long            'bytes per character
Dim bSeq() As Byte         'sequence as a byte array
Dim i As Long
On Error Resume Next
BPC = LenB(aa)
AA_ANSI = CByte(Asc(aa))
bSeq = Seq
For i = 0 To UBound(bSeq) Step BPC
    If bSeq(i) = AA_ANSI Then AACount = AACount + 1
Next i
End Function

' Unused Function (July 2003)
'''Public Function ELCount(ByVal Seq As String, _
'''                        ByVal EL As String) As Long
''''--------------------------------------------------------
''''returns number of occurences of element in Seq; function
''''uses ICR-2LS methods to retrieve individual counts
''''--------------------------------------------------------
'''Dim MSeq As String         'protein sequence converted to molecular formula
'''Dim mCnt As Long           'number of elements in formula
'''Dim SeqM() As String
'''Dim SeqMCnt() As Long
'''Dim TmpCnt As Long
'''Dim i As Long
'''On Error Resume Next
'''MSeq = objICR2LS.GetMF(Seq) 'use ICR-2LS services here
''''MSeq contains formula in form something like N14 H25 C33
'''mCnt = SplitMolFormula(MSeq, SeqM(), SeqMCnt())
'''If mCnt > 0 Then
'''   For i = 0 To mCnt - 1
'''       If EL = Left$(SeqM(i), Len(EL)) Then
'''          TmpCnt = TmpCnt + SeqMCnt(i)
'''       End If
'''   Next i
'''End If
'''ELCount = TmpCnt
'''End Function

' Unused Function (July 2003)
'''Public Function SplitMolFormula(ByVal MF As String, _
'''                                MFEL() As String, _
'''                                MFELCnt() As Long) As Long
''''-------------------------------------------------------------
''''splits molecular formula MF from form N12 C25 H11 to arrays
''''MFEL that contains elements of formula, and MFELCnt with
''''their count; elements in original formula are space delimited
''''function returns number of elements
''''-------------------------------------------------------------
'''Dim Tmp() As String
'''Dim TmpCnt As Long
'''Dim Num As Double
'''Dim Wrd As String
'''Dim i As Long
'''On Error Resume Next
'''Tmp = Split(MF, Chr$(32))
'''TmpCnt = UBound(Tmp) + 1
'''If TmpCnt > 0 Then
'''   ReDim MFEL(TmpCnt - 1)
'''   ReDim MFELCnt(TmpCnt - 1)
'''   For i = 0 To TmpCnt - 1
'''       Num = 0
'''       Wrd = ""
'''       Select Case GetNumberWord(Tmp(i), Num, Wrd)
'''       Case nwEmpty, nwNumOnly  'can't do anything here
'''            MFEL(i) = ""
'''            MFELCnt(i) = -1
'''       Case nwWordOnly          'means num=1
'''            MFEL(i) = Trim$(Wrd)
'''            MFELCnt(i) = -1
'''       Case nwNumWord
'''            MFEL(i) = Trim$(Wrd)
'''            MFELCnt(i) = CLng(Num)
'''       End Select
'''   Next i
'''End If
'''SplitMolFormula = TmpCnt
'''End Function

' Unused Function (July 2003)
'''Public Function GetNumberWord(ByVal sNumWrd As String, _
'''                              ByRef dNumber As Double, _
'''                              ByRef sWord As String) As glNumWord
''''----------------------------------------------------------------
''''puts number in Number and word in Word; returns result from
''''enum. glNumWord (nwEmpty, nwNumOnly, nwWordOnly, nwNumWord)
''''----------------------------------------------------------------
'''Dim bNumWrd() As Byte
'''Dim bTmp() As Byte
'''Dim sTmp As String
'''Dim bCnt As Long
'''Dim i As Long
'''Dim BPC As Long            'bytes per character
'''Dim FirstNumPos As Long
'''On Error Resume Next
''''load everything in byte array(it will work faster)
'''bNumWrd = sNumWrd
'''BPC = LenB("A")
'''bCnt = UBound(bNumWrd) + 1
'''FirstNumPos = -1
'''If bCnt > 0 Then
'''   'go find first numeric byte after which all bytes are numeric (characters)
'''   For i = 0 To bCnt - BPC + 1 Step BPC
'''     If (bNumWrd(i) >= 48 And bNumWrd(i) <= 57) Then  'Chr$(48)="0"
'''        If FirstNumPos < 0 Then FirstNumPos = i       'Chr$(57)="9"
'''     Else                           'non-numeric byte - have to reset
'''        FirstNumPos = -1
'''     End If
'''   Next i
'''   If FirstNumPos < 0 Then                       'word only
'''      sWord = sNumWrd
'''      GetNumberWord = nwWordOnly
'''   ElseIf FirstNumPos = 0 Then                   'number only
'''      dNumber = CDbl(sNumWrd)
'''      GetNumberWord = nwNumOnly
'''   Else                                 'number word
'''      ReDim bTmp(bCnt - 1)
'''      CopyMemory bTmp(0), bNumWrd(FirstNumPos), bCnt - FirstNumPos
'''      ReDim Preserve bTmp(bCnt - FirstNumPos - 1)
'''      sTmp = CStr(bTmp)
'''      dNumber = CDbl(sTmp)
'''      'trim number from word
'''      ReDim Preserve bNumWrd(FirstNumPos - 1)
'''      sWord = bNumWrd
'''      GetNumberWord = nwNumWord
'''   End If
'''Else
'''    GetNumberWord = nwEmpty
'''End If
'''End Function


Private Sub InitChargeStateMap()
Dim i As Long
For i = 1 To 5
    glCS1(i) = i
    glCS2(i) = i
Next i
End Sub

Public Function GetChargeStateMapIndex(ByVal CS As Long) As Long
'-----------------------------------------------------------------
'returns bin index 1 to 5 to which charge state belongs; 6 if none
'-----------------------------------------------------------------
Dim i As Long
For i = 1 To 5
    If glCS1(i) <= CS And CS <= glCS2(i) Then
       GetChargeStateMapIndex = i
       Exit Function
    End If
Next i
GetChargeStateMapIndex = 6
End Function

Private Sub InitDisplay0()
PresentDisplay0Type = glCSType
Display0MinScan = 1
Display0MaxScan = 1000
End Sub

' Unused Function (March 2003)
'''Public Function FormatStringA(sS As String, lChCnt As Long, algA As glAlignment) As String
''''-----------------------------------------------------------------------------------------
''''returns aligned string sText with exactly lCharCnt characters
''''-----------------------------------------------------------------------------------------
'''Dim ExtraCnt As Long, LeftExtraCnt As Long
'''On Error Resume Next
'''ExtraCnt = lChCnt - Len(sS)
'''If ExtraCnt > 0 Then
'''   Select Case algA
'''   Case algLeft
'''        FormatStringA = sS & Space$(ExtraCnt)
'''   Case algRight
'''        FormatStringA = Space$(ExtraCnt) & sS
'''   Case algCenter
'''        LeftExtraCnt = CLng(ExtraCnt / 2)
'''        FormatStringA = Space$(LeftExtraCnt) & sS & Space$(ExtraCnt - LeftExtraCnt)
'''   End Select
'''Else                            'return original string
'''   FormatStringA = sS
'''End If
'''End Function

Public Function LongToSignedShort(ByVal dwUnsigned As Long) As Integer
'----------------------------------------------------------------------------------------
'converts unsigned Long to signed integer; this function is neccessary for some API calls
'----------------------------------------------------------------------------------------
If dwUnsigned < 32768 Then
   LongToSignedShort = CInt(dwUnsigned)
Else
   LongToSignedShort = CInt(dwUnsigned - &H10000)
End If
End Function

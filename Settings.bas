Attribute VB_Name = "Module6"
'registry settings operations; ICR-2LS management functions
'last modified 07/27/2000 nt
'----------------------------------------------------------
'there is a cheat on SelColor setting which is not part of
'any structure but is added to OtherColorsPrefs so that
'Registry will remember
Option Explicit

Public Const DEFAULT_TOLERANCE_REFINEMENT_MW_TOL As Double = 25
Public Const DEFAULT_TOLERANCE_REFINEMENT_MW_TOL_TYPE As Integer = gltPPM
Public Const DEFAULT_TOLERANCE_REFINEMENT_NET_TOL As Double = 0.05

Private Const RECENT_DB_CONNECTIONS_MAX_COUNT = 25
Private Const RECENT_DB_CONNECTIONS_SECTION_NAME = "RecentDBConnections"
Private Const RECENT_DB_CONNECTIONS_KEY_COUNT_NAME = "ConnectionCount"
Private Const RECENT_DB_CONNECTION_SUBSECTION_NAME = "Connection"
Private Const RECENT_DB_CONNECTION_INFOVERSION_NAME = "InfoVersion"
Private Const RECENT_DB_CONNECTION_INFOVERSION = 2

Private Const NET_ADJ_SECTION_OLDNAME = "UMCNetDef"
Private Const NET_ADJ_SECTION_NEWNAME = "UMCNETAdjDef"
Private Const NET_ADJ_MS_WARP_SECTION = "UMCNETAdjMSWarpDef"

'Settings Variables
Private sCooSysPref As String            'type,origin,Horientation,Vorientation
Private sICR2LSPref As String            'command line
Private sDDClrPref As String             'under,mid,over colors,DDRatioMax
Private sBackForeCSIsoClrPref As String  'back, fore, CS, Iso colors
Private sCSIsoShapePref As String        'CS & Iso data spots shape
Private sTolerancesPref As String        'duptolerance, dbtolerance, isodatafit
Private sDrawingPref As String           'shape,border color, exclude over the line, min point size
Private sSwitchPref As String            'IsoDataFrom, case 2 close results, Iso ICR 2LS m/z

' No longer supported (March 2006)
''Private sAMTPref As String               'AMT database preferences
''Private sFTICR_AMTPref As String         'FTICR_AMT database preferences

Public Const DEFAULT_MASS_BIN_SIZE_PPM As Single = 0.2
Public Const DEFAULT_GANET_BIN_SIZE As Single = 0.001

Public Const DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD As Long = 400000

Private Const ENTRY_NOT_FOUND = "<<NOT_FOUND>>"

Public Sub SaveCurrentSettingsToIniFile(lngGelIndex As Long)
    Dim strIniFilePath As String
    
    strIniFilePath = SelectFile(MDIForm1.hwnd, "Select existing Ini file or enter a new name", "", True, "*.ini", "Ini files (*.*)|*.*|All Files (*.*)|*.*")
    
    If Len(strIniFilePath) > 0 Then
        ' Update glbPreferencesExpanded.AutoAnalysisDBInfo with the DB settings for the current gel
        
        strIniFilePath = FileExtensionForce(strIniFilePath, ".ini")
        
        If Not GelAnalysis(lngGelIndex) Is Nothing Then
            With glbPreferencesExpanded
                FillGelAnalysisInfo .AutoAnalysisDBInfo, GelAnalysis(lngGelIndex)
                .AutoAnalysisDBInfoIsValid = .AutoAnalysisDBInfo.ValidAnalysisDataPresent
            End With
        End If
        
        IniFileSaveSettings glbPreferencesExpanded, UMCDef, UMCIonNetDef, UMCNetAdjDef, UMCInternalStandards, samtDef, glPreferences, strIniFilePath, True
    End If
End Sub

Public Sub ResetOptions(gp As GelPrefs)
'put options on default values
    
    ResetGelPrefs gp
    ResetICR2LSPreferences
    ResetOtherColorsPreferences
    ResetCSIsoShapePreferences
    ResetDDClrPreferences
    
    ' No longer supported (March 2006)
    ''ResetAMTPreferences
    ''ResetFTICR_AMTPreferences
    
    ResetExpandedPreferences glbPreferencesExpanded
End Sub

Public Sub ResetGelPrefs(gp As GelPrefs)
    ResetSwitchPreferences gp
    ResetTolerancesPreferences gp
    ResetDrawingPreferences gp
    ResetCooSysPreferences gp
End Sub

Private Sub ResetSwitchPreferences(gp As GelPrefs)
With gp
    .IsoDataField = isfMWMono       ' 7
    .Case2Results = 1
    .DRDefinition = glNormal
    .IsoICR2LSMOverZ = True
End With
End Sub

Private Sub ResetTolerancesPreferences(gp As GelPrefs)
With gp
    .DBTolerance = -1
    .DupTolerance = 2
    .IsoDataFit = 0.15
End With
End Sub

Private Sub ResetDrawingPreferences(gp As GelPrefs)
With gp
    .BorderClrSameAsInt = True
    .MaxPointFactor = 2
    .MinPointFactor = 0.5
    .AbuAspectRatio = 1
End With
End Sub

Private Sub ResetAutoSearchModeEntry(udtAutoSearchModeEntry As udtAutoAnalysisSearchModeOptionsType)
    With udtAutoSearchModeEntry
        .SearchMode = AUTO_SEARCH_NONE
        .AlternateOutputFolderPath = ""
        .WriteResultsToTextFile = False
        .ExportResultsToDatabase = False
        .ExportUMCMembers = False
        .PairSearchAssumeMassTagsAreLabeled = False
        
        If APP_BUILD_DISABLE_MTS Then
            .InternalStdSearchMode = issmFindOnlyMassTags
        Else
            .InternalStdSearchMode = issmFindWithMassTags
        End If
        
        .DBSearchMinimumHighNormalizedScore = 0
        .DBSearchMinimumHighDiscriminantScore = 0
        .DBSearchMinimumPeptideProphetProbability = 0
        ResetDBSearchMassMods .MassMods
    End With
End Sub

Public Sub ResetDBSearchMassMods(udtMassMods As udtDBSearchMassModificationOptionsType)
    With udtMassMods
        .DynamicMods = True
        .N15InsteadOfN14 = False
        .PEO = False
        .ICATd0 = False
        .ICATd8 = False
        .Alkylation = False
        .AlkylationMass = glALKYLATION
        .ResidueToModify = ""
        .ResidueMassModification = 0
        .OtherInfo = ""
    End With
End Sub

Private Sub ResetCooSysPreferences(gp As GelPrefs)
With gp
    .CooType = glFNCooSys
    .CooOrigin = glOriginBL
    .CooHOrientation = glNormal
    .CooVOrientation = glNormal
    .CooVAxisScale = glVAxisLin
End With
End Sub

Private Sub ResetICR2LSPreferences()
sICR2LSCommand = "C:\Program Files\ICR-2LS\icr-2ls.exe "
End Sub

Public Sub ResetDataFilters(lngGelIndex As Long, udtPreferences As GelPrefs)
    Dim i As Integer
    
On Error GoTo ResetDataFiltersErrorHandler

    With udtPreferences
        .DBTolerance = -1
        .DupTolerance = 2
        .IsoDataFit = 0.15
    End With
    
    With GelData(lngGelIndex)
       For i = 1 To MAX_FILTER_COUNT              'Do not use any filter initially
         .DataFilter(i, 0) = False
       Next i
       .Preferences = udtPreferences
       .DataFilter(fltDupTolerance, 1) = udtPreferences.DupTolerance
       .DataFilter(fltDBTolerance, 1) = udtPreferences.DBTolerance
       .DataFilter(fltIsoFit, 1) = udtPreferences.IsoDataFit
       .DataFilter(fltCase2CloseResults, 1) = udtPreferences.Case2Results
       .DataFilter(fltAR, 0) = 0
       .DataFilter(fltAR, 1) = -1
       .DataFilter(fltAR, 2) = -1
       .DataFilter(fltID, 1) = 0
       .DataFilter(fltCSAbu, 1) = 0             'min abundance
       .DataFilter(fltIsoAbu, 1) = 0
       .DataFilter(fltCSMW, 1) = 0              'min mass range
       .DataFilter(fltIsoMW, 1) = 0
       .DataFilter(fltIsoCS, 1) = 0
       .DataFilter(fltCSStDev, 1) = 1
       .DataFilter(fltIsoMZ, 1) = 0             'min m/z range
       .DataFilter(fltEvenOddScanNumber, 1) = 0     ' Use all scans
    End With
    
    Exit Sub

ResetDataFiltersErrorHandler:
    Debug.Print "Error in ResetDataFilters: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "ResetDataFilters"
    Resume Next
End Sub

Private Sub ResetDDClrPreferences()
glUnderColor = glUnderColorDefault
glMidColor = glMidColorDefault
glOverColor = glOverColorDefault
glDDRatioMax = glHugeOverReal
End Sub

Private Sub ResetOtherColorsPreferences()
glBackColor = vbWhite
glForeColor = vbBlack
glCSColor = glCSColorDefault
glIsoColor = glIsoColorDefault
glSelColor = vbRed
End Sub

Private Sub ResetCSIsoShapePreferences()
glCSShape = 0       'oval
glIsoShape = 0
End Sub

Private Function GetCooSysPrefs(udtPrefs As GelPrefs) As String
'this is used only with global preferences
With udtPrefs
    GetCooSysPrefs = CStr(.CooType) & "," & CStr(.CooOrigin) & "," _
                    & CStr(.CooHOrientation) & "," & CStr(.CooVOrientation) _
                    & "," & CStr(.CooVAxisScale)
End With
End Function

Private Function GetDDClrPrefs() As String
'this is used only with global preferences
GetDDClrPrefs = CStr(glUnderColor) & "," & CStr(glMidColor) _
                & "," & CStr(glOverColor) & "," & CStr(glDDRatioMax)
End Function

Private Function GetDrawingPrefs(udtPrefs As GelPrefs) As String
'this is used only with global preferences
With udtPrefs
    GetDrawingPrefs = CStr(.BorderClrSameAsInt) & "," _
                & CStr(.MinPointFactor) & "," & CStr(.MaxPointFactor) _
                & "," & CStr(.AbuAspectRatio)
End With
End Function

Private Function GetICR2LSPrefs() As String
GetICR2LSPrefs = sICR2LSCommand
End Function

Private Function GetOtherColorsPrefs() As String
GetOtherColorsPrefs = CStr(glBackColor) & "," & CStr(glForeColor) & "," _
                    & CStr(glCSColor) & "," & CStr(glIsoColor) & "," & CStr(glSelColor)
End Function

Private Function GetCSIsoShapePrefs() As String
GetCSIsoShapePrefs = CStr(glCSShape) & "," & CStr(glIsoShape)
End Function

Public Function GetAutoAnalysisOptionsList() As String
    GetAutoAnalysisOptionsList = AUTO_SEARCH_NONE & ", " & AUTO_SEARCH_EXPORT_UMCS_ONLY & ", " & _
                                 AUTO_SEARCH_ORGANISM_MTDB & ", " & _
                                 AUTO_SEARCH_UMC_MTDB & ", " & _
                                 AUTO_SEARCH_UMC_CONGLOMERATE & ", " & _
                                 AUTO_SEARCH_UMC_CONGLOMERATE_PAIRED & ", " & AUTO_SEARCH_UMC_CONGLOMERATE_UNPAIRED & ", " & _
                                 AUTO_SEARCH_UMC_CONGLOMERATE_LIGHT_PAIRS_PLUS_UNPAIRED & ", " & _
                                 AUTO_SEARCH_PAIRS_N14N15_CONGLOMERATEMASS & ", " & _
                                 AUTO_SEARCH_PAIRS_ICAT & ", " & AUTO_SEARCH_PAIRS_PEO

End Function

Public Function GetPictureGraphicsTypeList() As String
    GetPictureGraphicsTypeList = Trim(pftPictureFileTypeConstants.pftPNG) & "=PNG, " & Trim(pftPictureFileTypeConstants.pftJPG) & "=JPEG, " & _
                                 Trim(pftPictureFileTypeConstants.pftWMF) & "=WMF, " & Trim(pftPictureFileTypeConstants.pftEMF) & "=EMF, and " & _
                                 Trim(pftPictureFileTypeConstants.pftBMP) & "=BMP"
End Function

Public Function GetErrorGraphicsTypeList() As String
    GetErrorGraphicsTypeList = Trim(pftPictureFileTypeConstants.pftPNG) & "=PNG, " & Trim(pftPictureFileTypeConstants.pftJPG) & "=JPEG"
End Function

Private Function GetUMCSearchModeList() As String
    GetUMCSearchModeList = AUTO_ANALYSIS_UMC2003 & " and " & AUTO_ANALYSIS_UMCIonNet
End Function

Private Function GetUMCTypeList() As String

    GetUMCTypeList = Trim(glUMC_TYPE_INTENSITY) & "=Favor Higher Intensity, " & Trim(glUMC_TYPE_FIT) & "=Favor Better Fit, " & _
                     Trim(glUMC_TYPE_MINCNT) & "=Minimize Count, " & Trim(glUMC_TYPE_MAXCNT) & "=Maximize Count, " & _
                     Trim(glUMC_TYPE_UNQAMT) & "=Unique AMT, " & Trim(glUMC_TYPE_ISHRINKINGBOX) & "=Shrinking Box using Intensity, and " & _
                     Trim(glUMC_TYPE_FSHRINKINGBOX) & "=Shrinking Box using Fit"
End Function

Private Function GetTolerancesPrefs(udtPrefs As GelPrefs) As String
'this is used only with global preferences
With udtPrefs
    GetTolerancesPrefs = CStr(.DupTolerance) & "," & CStr(.DBTolerance) _
                         & "," & CStr(.IsoDataFit)
End With
End Function

Private Sub ResolveCooSysPrefs(udtPrefs As GelPrefs)
Dim aRes As Variant
On Error GoTo err_resolvecsp
aRes = ResolvePrefsString(sCooSysPref)
With udtPrefs
    .CooType = val(aRes(1))
    .CooOrigin = val(aRes(2))
    .CooHOrientation = val(aRes(3))
    .CooVOrientation = val(aRes(4))
    .CooVAxisScale = val(aRes(5))
End With
Exit Sub

err_resolvecsp:
ResetCooSysPreferences udtPrefs
End Sub

Private Sub ResolveICR2LSPrefs()
Dim aRes As Variant
On Error GoTo err_resolveicr
aRes = ResolvePrefsString(sICR2LSPref)
sICR2LSCommand = aRes(1)
Exit Sub

err_resolveicr:
ResetICR2LSPreferences
End Sub

Private Sub ResolveCSIsoShapePrefs()
Dim aRes As Variant
On Error GoTo err_resolveshape
aRes = ResolvePrefsString(sCSIsoShapePref)
glCSShape = val(aRes(1))
glIsoShape = val(aRes(2))
Exit Sub

err_resolveshape:
ResetCSIsoShapePreferences
End Sub

Private Sub ResolveDDClrPrefs()
Dim aRes As Variant
On Error GoTo err_resolveddc
aRes = ResolvePrefsString(sDDClrPref)
glUnderColor = val(aRes(1))
glMidColor = val(aRes(2))
glOverColor = val(aRes(3))
glDDRatioMax = val(aRes(4))
Exit Sub

err_resolveddc:
ResetDDClrPreferences
End Sub

Private Sub ResolveOtherColorsPrefs()
Dim aRes As Variant
On Error GoTo err_resolveocp
aRes = ResolvePrefsString(sBackForeCSIsoClrPref)
glBackColor = val(aRes(1))
glForeColor = val(aRes(2))
glCSColor = val(aRes(3))
glIsoColor = val(aRes(4))
glSelColor = val(aRes(5))
Exit Sub

err_resolveocp:
ResetOtherColorsPreferences
End Sub

Private Function ResolvePrefsString(ByVal S As String) As Variant
'returns variant array created from comma delimited string
Dim aTmp(10) As String
Dim StartPos As Integer
Dim EndPos As Integer
Dim bDone As Boolean
Dim i As Integer
On Error GoTo err_resolve
StartPos = 1
i = 0
bDone = False
Do While Not bDone
   EndPos = InStr(StartPos, S, ",")
   If EndPos > 0 Then
      i = i + 1
      aTmp(i) = Mid$(S, StartPos, EndPos - StartPos)
      StartPos = EndPos + 1
   Else
      i = i + 1
      aTmp(i) = Mid$(S, StartPos, Len(S) - StartPos + 1)
      bDone = True
   End If
Loop
ResolvePrefsString = aTmp
Exit Function

err_resolve:
ResolvePrefsString = Null
End Function

Private Sub ResolveTolerancesPrefs(udtPrefs As GelPrefs)
Dim aRes As Variant
On Error GoTo err_resolvetlr
aRes = ResolvePrefsString(sTolerancesPref)
With udtPrefs
    .DupTolerance = val(aRes(1))
    .DBTolerance = val(aRes(2))
    .IsoDataFit = val(aRes(3))
End With
Exit Sub

err_resolvetlr:
ResetTolerancesPreferences udtPrefs
End Sub

Private Sub ResolveDrawingPrefs(udtPrefs As GelPrefs)
Dim aRes As Variant
On Error GoTo err_resolvedrw
aRes = ResolvePrefsString(sDrawingPref)
With udtPrefs
    .BorderClrSameAsInt = CBool(aRes(1))
    .MinPointFactor = val(aRes(2))
    .MaxPointFactor = val(aRes(3))
    .AbuAspectRatio = val(aRes(4))
End With
Exit Sub

err_resolvedrw:
ResetDrawingPreferences udtPrefs
End Sub

Private Function GetSwitchPrefs(udtPrefs As GelPrefs) As String
'this is used only with global preferences
With udtPrefs
    GetSwitchPrefs = CStr(.IsoDataField) & "," & _
            CStr(.Case2Results) & "," & CStr(.DRDefinition) _
            & "," & CStr(.IsoICR2LSMOverZ)
End With
End Function

Private Sub ResolveSwitchPrefs(udtPrefs As GelPrefs)
Dim aRes As Variant
On Error GoTo err_resolvetlr
aRes = ResolvePrefsString(sSwitchPref)
With udtPrefs
    .IsoDataField = val(aRes(1))
    .Case2Results = val(aRes(2))
    .DRDefinition = val(aRes(3))
    .IsoICR2LSMOverZ = CBool(aRes(4))
End With
Exit Sub
err_resolvetlr:
ResetSwitchPreferences udtPrefs
End Sub

Public Function GetIniFilePath() As String
    GetIniFilePath = App.Path & "\" & INI_FILENAME
End Function

Public Function GetIniFileSetting(objIniStuff As clsIniStuff, strSection As String, strKey As String, Optional strDefault As String = "") As String
    Dim strSetting As String

    If objIniStuff.ReadValue(strSection, strKey, strDefault, strSetting) Then
        GetIniFileSetting = CStr(strSetting)
    Else
        GetIniFileSetting = strDefault
    End If

End Function

Private Function GetIniFileSettingDbl(objIniStuff As clsIniStuff, strSection As String, strKey As String, Optional dblDefault As Double = 0) As Double
    Dim strSetting As String
    Dim strDefault As String

    strDefault = Str(dblDefault)

    If objIniStuff.ReadValue(strSection, strKey, strDefault, strSetting) Then
        GetIniFileSettingDbl = CDblSafe(strSetting)
    Else
        GetIniFileSettingDbl = dblDefault
    End If

End Function

Private Function GetIniFileSettingSng(objIniStuff As clsIniStuff, strSection As String, strKey As String, Optional sngDefault As Single = 0) As Single
    Dim strSetting As String
    Dim strDefault As String

    strDefault = Str(sngDefault)

    If objIniStuff.ReadValue(strSection, strKey, strDefault, strSetting) Then
        GetIniFileSettingSng = CSngSafe(strSetting)
    Else
        GetIniFileSettingSng = sngDefault
    End If

End Function

Private Function GetIniFileSettingBln(objIniStuff As clsIniStuff, strSection As String, strKey As String, Optional blnDefault As Boolean = False) As Boolean
    Dim strSetting As String
    Dim strDefault As String

    strDefault = Str(blnDefault)

    If objIniStuff.ReadValue(strSection, strKey, strDefault, strSetting) Then
        GetIniFileSettingBln = CBoolSafe(Trim(strSetting))
    Else
        GetIniFileSettingBln = blnDefault
    End If

End Function

Private Function GetIniFileSettingInt(objIniStuff As clsIniStuff, strSection As String, strKey As String, Optional intDefault As Integer = 0) As Integer
    Dim strSetting As String
    Dim strDefault As String

    strDefault = Str(intDefault)

    If objIniStuff.ReadValue(strSection, strKey, strDefault, strSetting) Then
        GetIniFileSettingInt = CIntSafe(strSetting)
    Else
        GetIniFileSettingInt = intDefault
    End If

End Function

Public Function GetIniFileSettingLng(objIniStuff As clsIniStuff, strSection As String, strKey As String, Optional lngDefault As Long = 0) As Long
    Dim strSetting As String
    Dim strDefault As String

    strDefault = Str(lngDefault)

    If objIniStuff.ReadValue(strSection, strKey, strDefault, strSetting) Then
        GetIniFileSettingLng = CLngSafe(strSetting)
    Else
        GetIniFileSettingLng = lngDefault
    End If

End Function

Public Function IniFileReadSingleSetting(strSection As String, strKey As String, Optional strDefault As String = "", Optional strIniFilePath As String = "") As String
    ' Looks up a single setting in the .Ini file
    
    Dim IniStuff As New clsIniStuff
    Dim strValue As String
    
On Error GoTo IniFileReadSingleSettingErrorHandler

    ' Set the Ini filename
    If Len(strIniFilePath) > 0 Then
        IniStuff.FileName = strIniFilePath
    Else
        IniStuff.FileName = GetIniFilePath()
    End If
                
    strValue = GetIniFileSetting(IniStuff, strSection, strKey, strDefault)
    
    Set IniStuff = Nothing
    
    IniFileReadSingleSetting = strValue
    Exit Function

IniFileReadSingleSettingErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error loading data from the Ini file (" & IniStuff.FileName & "); Sub IniFileReadSingleSetting in Settings.Bas" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Print "Error in IniFileReadSingleSetting: " & Err.Description
        Debug.Assert False
        LogErrors Err.Number, "Settings.Bas->IniFileReadSingleSetting"
    End If
    Set IniStuff = Nothing
    IniFileReadSingleSetting = strDefault
    
End Function

Public Sub IniFileLoadSettings(ByRef udtPrefsExpanded As udtPreferencesExpandedType, ByRef udtUMCDef As UMCDefinition, ByRef udtUMCIonNetDef As UMCIonNetDefinition, ByRef udtUMCNetAdjDef As NetAdjDefinition, ByRef udtUMCInternalStandards As udtInternalStandardsType, ByRef udtAMTDef As SearchAMTDefinition, ByRef udtPrefs As GelPrefs, Optional strIniFilePath As String = "", Optional bnlAutoAnalysisFieldsOnly As Boolean = False)
    
    'Loads up settings from an .ini file
    Dim IniStuff As New clsIniStuff
    Dim intIndex As Integer, intAutoSearchModeIndex As Integer
    Dim intInternalStandardcount As Integer
    Dim intInternalStandardIndex As Integer
    Dim udtDBSettingsSingle As udtDBSettingsType
    
    Dim strKeys() As String
    Dim strValues() As String
    Dim strKeyPrefix As String, strKeyValue As String
    Dim strUMCNetAdjDefSectionName As String
    Dim strSectionName As String
    Dim dblInternalStandardMass As Double
    
    Dim blnLegacySectionName As Boolean
    
On Error GoTo LoadSettingsFileHandler

    ' Set the Ini filename
    If Len(strIniFilePath) > 0 Then
        IniStuff.FileName = strIniFilePath
    Else
        IniStuff.FileName = GetIniFilePath()
    End If
                
    ' Reset the options to defaults (in case an entry isn't present in the Ini file)
    ResetOptions udtPrefs
    ResetExpandedPreferences udtPrefsExpanded
    
    ' Turn off the maximum data count option if auto-analysis is enabled
    ' If MaximumDataCountEnabled=True is present in the .Ini file, then this will be re-enabled
    If udtPrefsExpanded.AutoAnalysisStatus.Enabled Then
        udtPrefsExpanded.AutoAnalysisFilterPrefs.MaximumDataCountEnabled = False
    End If
    
    ' Paths
    sICR2LSCommand = GetIniFileSetting(IniStuff, "Paths", "ICR2LS", sICR2LSCommand)
    
    ' udtUMCDef preferences
    With udtPrefsExpanded.AutoAnalysisOptions
        .UMCSearchMode = GetIniFileSetting(IniStuff, "UMCDef", "UMCSearchMode", .UMCSearchMode)
        If Len(.UMCSearchMode) = 0 Then .UMCSearchMode = AUTO_ANALYSIS_UMCIonNet
        .UMCShrinkingBoxWeightAverageMassByIntensity = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCShrinkingBoxWeightAverageMassByIntensity", .UMCShrinkingBoxWeightAverageMassByIntensity)
    End With
    With udtUMCDef
        .UMCType = GetIniFileSettingInt(IniStuff, "UMCDef", "UMCType", .UMCType)
        .DefScope = .DefScope   ' (Not stored in .Ini file)
        .MWField = GetIniFileSettingInt(IniStuff, "UMCDef", "MWField", .MWField)
        .TolType = GetIniFileSettingInt(IniStuff, "UMCDef", "TolType", .TolType)
        .Tol = GetIniFileSettingDbl(IniStuff, "UMCDef", "Tol", .Tol)
        .UMCSharing = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCSharing", .UMCSharing)
        .UMCUniCS = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCUniCS", .UMCUniCS)
        .ClassAbu = GetIniFileSettingInt(IniStuff, "UMCDef", "ClassAbu", .ClassAbu)
        .ClassMW = GetIniFileSettingInt(IniStuff, "UMCDef", "ClassMW", .ClassMW)
        .GapMaxCnt = GetIniFileSettingLng(IniStuff, "UMCDef", "GapMaxCnt", .GapMaxCnt)
        .GapMaxSize = GetIniFileSettingLng(IniStuff, "UMCDef", "GapMaxSize", .GapMaxSize)
        .GapMaxPct = GetIniFileSettingDbl(IniStuff, "UMCDef", "GapMaxPct", .GapMaxPct)
        .UMCNETType = GetIniFileSettingInt(IniStuff, "UMCDef", "UMCNETType", .UMCNETType)
        .InterpolateGaps = GetIniFileSettingBln(IniStuff, "UMCDef", "InterpolateGaps", .InterpolateGaps)
        .InterpolateMaxGapSize = GetIniFileSettingLng(IniStuff, "UMCDef", "InterpolateMaxGapSize", .InterpolateMaxGapSize)
        .InterpolationType = GetIniFileSettingLng(IniStuff, "UMCDef", "InterpolationType", .InterpolationType)
        .ChargeStateStatsRepType = GetIniFileSettingInt(IniStuff, "UMCDef", "ChargeStateStatsRepType", .ChargeStateStatsRepType)
        .UMCClassStatsUseStatsFromMostAbuChargeState = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCClassStatsUseStatsFromMostAbuChargeState", .UMCClassStatsUseStatsFromMostAbuChargeState)
    End With
    
    ' udtUMCDef preferences stored in udtPrefsExpanded
    With udtPrefsExpanded.UMCAutoRefineOptions
        .UMCAutoRefineRemoveCountLow = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCAutoRefineRemoveCountLow", .UMCAutoRefineRemoveCountLow)
        .UMCAutoRefineRemoveCountHigh = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCAutoRefineRemoveCountHigh", .UMCAutoRefineRemoveCountHigh)
        .UMCAutoRefineRemoveMaxLengthPctAllScans = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCAutoRefineRemoveMaxLengthPctAllScans", .UMCAutoRefineRemoveMaxLengthPctAllScans)
        
        .UMCAutoRefineMinLength = GetIniFileSettingLng(IniStuff, "UMCDef", "UMCAutoRefineMinLength", .UMCAutoRefineMinLength)
        .UMCAutoRefineMaxLength = GetIniFileSettingLng(IniStuff, "UMCDef", "UMCAutoRefineMaxLength", .UMCAutoRefineMaxLength)
        .UMCAutoRefineMaxLengthPctAllScans = GetIniFileSettingLng(IniStuff, "UMCDef", "UMCAutoRefineMaxLengthPctAllScans", .UMCAutoRefineMaxLengthPctAllScans)
        
        .UMCAutoRefinePercentMaxAbuToUseForLength = GetIniFileSettingLng(IniStuff, "UMCDef", "UMCAutoRefinePercentMaxAbuToUseForLength", .UMCAutoRefinePercentMaxAbuToUseForLength)
        .TestLengthUsingScanRange = GetIniFileSettingBln(IniStuff, "UMCDef", "TestLengthUsingScanRange", .TestLengthUsingScanRange)
        .MinMemberCountWhenUsingScanRange = GetIniFileSettingLng(IniStuff, "UMCDef", "MinMemberCountWhenUsingScanRange", .MinMemberCountWhenUsingScanRange)
        
        ' Store duplicate values in .UMCMinCnt and .UMCAutoRefineMinLength
        udtUMCDef.UMCMinCnt = .UMCAutoRefineMinLength
        udtUMCDef.UMCMaxCnt = .UMCAutoRefineMaxLength
    
        .UMCAutoRefineRemoveAbundanceLow = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCAutoRefineRemoveAbundanceLow", .UMCAutoRefineRemoveAbundanceLow)
        .UMCAutoRefineRemoveAbundanceHigh = GetIniFileSettingBln(IniStuff, "UMCDef", "UMCAutoRefineRemoveAbundanceHigh", .UMCAutoRefineRemoveAbundanceHigh)
        .UMCAutoRefinePctLowAbundance = GetIniFileSettingDbl(IniStuff, "UMCDef", "UMCAutoRefinePctLowAbundance", .UMCAutoRefinePctLowAbundance)
        .UMCAutoRefinePctHighAbundance = GetIniFileSettingDbl(IniStuff, "UMCDef", "UMCAutoRefinePctHighAbundance", .UMCAutoRefinePctHighAbundance)
        
        .SplitUMCsByAbundance = GetIniFileSettingBln(IniStuff, "UMCDef", "SplitUMCsByAbundance", .SplitUMCsByAbundance)
        With .SplitUMCOptions
            .MinimumDifferenceInAveragePpmMassToSplit = GetIniFileSettingDbl(IniStuff, "UMCDef", "MinimumDifferenceInAveragePpmMassToSplit", .MinimumDifferenceInAveragePpmMassToSplit)
            .StdDevMultiplierForSplitting = GetIniFileSettingDbl(IniStuff, "UMCDef", "StdDevMultiplierForSplitting", .StdDevMultiplierForSplitting)
            .MaximumPeakCountToSplitUMC = GetIniFileSettingLng(IniStuff, "UMCDef", "MaximumPeakCountToSplitUMC", .MaximumPeakCountToSplitUMC)
            .PeakDetectIntensityThresholdPercentageOfMaximum = GetIniFileSettingLng(IniStuff, "UMCDef", "PeakDetectIntensityThresholdPercentageOfMaximum", .PeakDetectIntensityThresholdPercentageOfMaximum)
            .PeakDetectIntensityThresholdAbsoluteMinimum = GetIniFileSettingDbl(IniStuff, "UMCDef", "PeakDetectIntensityThresholdAbsoluteMinimum", .PeakDetectIntensityThresholdAbsoluteMinimum)
            .PeakWidthPointsMinimum = GetIniFileSettingLng(IniStuff, "UMCDef", "PeakWidthPointsMinimum", .PeakWidthPointsMinimum)
            .PeakWidthInSigma = GetIniFileSettingLng(IniStuff, "UMCDef", "PeakWidthInSigma", .PeakWidthInSigma)
            .ScanGapBehavior = GetIniFileSettingInt(IniStuff, "UMCDef", "ScanGapBehavior", CInt(.ScanGapBehavior))
        End With
    End With
    
    ' UMCIonNetDef Options
    With udtUMCIonNetDef
        .NetDim = GetIniFileSettingLng(IniStuff, "UMCIonNetDef", "NetDim", .NetDim)
        .NetActualDim = GetIniFileSettingLng(IniStuff, "UMCIonNetDef", "NetActualDim", .NetActualDim)
        .MetricType = GetIniFileSettingLng(IniStuff, "UMCIonNetDef", "MetricType", .MetricType)
        .NETType = GetIniFileSettingLng(IniStuff, "UMCIonNetDef", "NETType", .NETType)
        .TooDistant = GetIniFileSettingDbl(IniStuff, "UMCIonNetDef", "TooDistant", .TooDistant)
        
        If .NetDim > UBound(.MetricData()) + 1 Then
            Debug.Assert False
            .NetDim = UBound(.MetricData()) + 1
        End If
            
        For intIndex = 0 To .NetDim - 1
            strKeyPrefix = "Dim" & Trim(intIndex + 1)
        
            strKeyValue = GetIniFileSetting(IniStuff, "UMCIonNetDef", strKeyPrefix & "Use", ENTRY_NOT_FOUND)
            If strKeyValue = ENTRY_NOT_FOUND Then
                ' Assume default settings for any dimension not found
                Exit For
            Else
                With .MetricData(intIndex)
                    .Use = CBoolSafe(strKeyValue)
                    .DataType = GetIniFileSettingLng(IniStuff, "UMCIonNetDef", strKeyPrefix & "DataType", .DataType)
                    .WeightFactor = GetIniFileSettingDbl(IniStuff, "UMCIonNetDef", strKeyPrefix & "WeightFactor", .WeightFactor)
                    .ConstraintType = GetIniFileSettingLng(IniStuff, "UMCIonNetDef", strKeyPrefix & "ConstraintType", .ConstraintType)
                    .ConstraintValue = GetIniFileSettingDbl(IniStuff, "UMCIonNetDef", strKeyPrefix & "ConstraintValue", .ConstraintValue)
                    .ConstraintUnits = GetIniFileSettingLng(IniStuff, "UMCIonNetDef", strKeyPrefix & "ConstraintUnits", .ConstraintUnits)  ' 0 = Da, 1 = ppm; Only applies to Mass-based DataTypes
                End With
            End If
        Next intIndex
        
    End With
    
    ' UMCIonNetDef Options stored in udtPrefsExpanded
    With udtPrefsExpanded.UMCIonNetOptions
        .UMCRepresentative = GetIniFileSettingInt(IniStuff, "UMCIonNetDef", "UMCRepresentative", .UMCRepresentative)
        .MakeSingleMemberClasses = GetIniFileSettingBln(IniStuff, "UMCIonNetDef", "MakeSingleMemberClasses", .MakeSingleMemberClasses)
        .ConnectionLengthPostFilterMaxNET = GetIniFileSettingDbl(IniStuff, "UMCIonNetDef", "ConnectionLengthPostFilterMaxNET", .ConnectionLengthPostFilterMaxNET)
    End With
    
    ' UMCAdvancedStatsOptions Options stored in udtPrefsExpanded
    With udtPrefsExpanded.UMCAdvancedStatsOptions
        .ClassAbuTopXMinAbu = GetIniFileSettingDbl(IniStuff, "UMCAdvancedStatsOptions", "ClassAbuTopXMinAbu", .ClassAbuTopXMinAbu)
        .ClassAbuTopXMaxAbu = GetIniFileSettingDbl(IniStuff, "UMCAdvancedStatsOptions", "ClassAbuTopXMaxAbu", .ClassAbuTopXMaxAbu)
        .ClassAbuTopXMinMembers = GetIniFileSettingLng(IniStuff, "UMCAdvancedStatsOptions", "ClassAbuTopXMinMembers", .ClassAbuTopXMinMembers)
        
        .ClassMassTopXMinAbu = GetIniFileSettingDbl(IniStuff, "UMCAdvancedStatsOptions", "ClassMassTopXMinAbu", .ClassMassTopXMinAbu)
        .ClassMassTopXMaxAbu = GetIniFileSettingDbl(IniStuff, "UMCAdvancedStatsOptions", "ClassMassTopXMaxAbu", .ClassMassTopXMaxAbu)
        .ClassMassTopXMinMembers = GetIniFileSettingLng(IniStuff, "UMCAdvancedStatsOptions", "ClassMassTopXMinMembers", .ClassMassTopXMinMembers)
    End With
    
    ' UMCNetAdjDef Options
    strUMCNetAdjDefSectionName = NET_ADJ_SECTION_NEWNAME
    
    ' See if the UMCNetDef section is present
    ' If it is, then this is an old .Ini file and we need to change strUMCNetAdjDefSectionName
    If IniStuff.ReadSection(NET_ADJ_SECTION_OLDNAME, strKeys(), strValues()) Then
        ' Only change the name if there are at least 5 entries in the section
        If UBound(strKeys()) >= 5 Then
            blnLegacySectionName = True
            strUMCNetAdjDefSectionName = NET_ADJ_SECTION_OLDNAME
        End If
    End If
    
    ' UMC Net Adjustment Options
    With udtUMCNetAdjDef
        .MinUMCCount = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "MinUMCCount", .MinUMCCount)
        .MinScanRange = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "MinScanRange", .MinScanRange)
        .MaxScanPct = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "MaxScanPct", .MaxScanPct)
        .TopAbuPct = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "TopAbuPct", .TopAbuPct)
        ' Ignored: .PeakSelection = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "PeakSelection", .PeakSelection)
        ' Ignored: .PeakMaxAbuPct = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "PeakMaxAbuPct", .PeakMaxAbuPct)
        For intIndex = 0 To UBound(.PeakCSSelection)
            .PeakCSSelection(intIndex) = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "PeakCSSelection" & Trim(intIndex), .PeakCSSelection(intIndex))
        Next intIndex
        .MWTolType = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "MWTolType", .MWTolType)
        .MWTol = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "MWTol", .MWTol)
        .NETFormula = .NETFormula                       ' (Not stored in .Ini file)
        .NETTolIterative = .NETTolIterative             ' (Not stored in .Ini file)
        .NETorRT = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETorRT", .NETorRT)
        .UseNET = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "UseNET", .UseNET)
        .UseMultiIDMaxNETDist = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "UseMultiIDMaxNETDist", .UseMultiIDMaxNETDist)
        .MultiIDMaxNETDist = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "MultiIDMaxNETDist", .MultiIDMaxNETDist)
        .EliminateBadNET = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "EliminateBadNET", .EliminateBadNET)
        .MaxIDToUse = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "MaxIDToUse", .MaxIDToUse)
        .IterationStopType = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "IterationStopType", .IterationStopType)
        .IterationStopValue = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "IterationStopValue", .IterationStopValue)
        .IterationUseMWDec = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "IterationUseMWDec", .IterationUseMWDec)
        .IterationMWDec = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "IterationMWDec", .IterationMWDec)
        .IterationUseNETdec = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "IterationUseNETdec", .IterationUseNETdec)
        .IterationNETDec = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "IterationNETDec", .IterationNETDec)
        .IterationAcceptLast = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "IterationAcceptLast", .IterationAcceptLast)
    
        .InitialSlope = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "InitialSlope", .InitialSlope)
        .InitialIntercept = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "InitialIntercept", .InitialIntercept)
        
        ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''        .UseNetAdjLockers = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "UseNetAdjLockers", .UseNetAdjLockers)
''        .UseOldNetAdjIfFailure = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "UseOldNetAdjIfFailure", .UseOldNetAdjIfFailure)
''        .NetAdjLockerMinimumMatchCount = GetIniFileSettingInt(IniStuff, strUMCNetAdjDefSectionName, "NetAdjLockerMinimumMatchCount", .NetAdjLockerMinimumMatchCount)
    
        .UseRobustNETAdjustment = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "UseRobustNETAdjustment", .UseRobustNETAdjustment)
        .RobustNETAdjustmentMode = GetIniFileSettingInt(IniStuff, strUMCNetAdjDefSectionName, "RobustNETAdjustmentMode", .RobustNETAdjustmentMode)
        
        If APP_BUILD_DISABLE_LCMSWARP Then
            If .RobustNETAdjustmentMode <> UMCRobustNETModeConstants.UMCRobustNETIterative Then
                .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETIterative
            End If
        End If
        
        .RobustNETSlopeStart = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETSlopeStart", .RobustNETSlopeStart)
        .RobustNETSlopeEnd = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETSlopeEnd", .RobustNETSlopeEnd)
        .RobustNETSlopeIncreaseMode = GetIniFileSettingInt(IniStuff, strUMCNetAdjDefSectionName, "RobustNETSlopeIncreaseMode", .RobustNETSlopeIncreaseMode)
        .RobustNETSlopeIncrement = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETSlopeIncrement", .RobustNETSlopeIncrement)
        
        .RobustNETInterceptStart = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETInterceptStart", .RobustNETInterceptStart)
        .RobustNETInterceptEnd = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETInterceptEnd", .RobustNETInterceptEnd)
        .RobustNETInterceptIncrement = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETInterceptIncrement", .RobustNETInterceptIncrement)
        
        .RobustNETMassShiftPPMStart = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETMassShiftPPMStart", .RobustNETMassShiftPPMStart)
        .RobustNETMassShiftPPMEnd = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETMassShiftPPMEnd", .RobustNETMassShiftPPMEnd)
        .RobustNETMassShiftPPMIncrement = GetIniFileSettingSng(IniStuff, strUMCNetAdjDefSectionName, "RobustNETMassShiftPPMIncrement", .RobustNETMassShiftPPMIncrement)
        
        If Not APP_BUILD_DISABLE_LCMSWARP Then
            strSectionName = NET_ADJ_MS_WARP_SECTION
            With .MSWarpOptions
                .MassCalibrationType = GetIniFileSettingInt(IniStuff, strSectionName, "MassCalibrationType", .MassCalibrationType)
                .MinimumPMTTagObsCount = GetIniFileSettingLng(IniStuff, strSectionName, "MinimumPMTTagObsCount", .MinimumPMTTagObsCount)
                .MatchPromiscuity = GetIniFileSettingInt(IniStuff, strSectionName, "MatchPromiscuity", .MatchPromiscuity)
                
                .NETTol = GetIniFileSettingSng(IniStuff, strSectionName, "NETTol", .NETTol)
                .NumberOfSections = GetIniFileSettingLng(IniStuff, strSectionName, "NumberOfSections ", .NumberOfSections)
                .MaxDistortion = GetIniFileSettingInt(IniStuff, strSectionName, "MaxDistortion", .MaxDistortion)
                .ContractionFactor = GetIniFileSettingInt(IniStuff, strSectionName, "ContractionFactor", .ContractionFactor)
                
                .MassWindowPPM = GetIniFileSettingSng(IniStuff, strSectionName, "MassWindowPPM", .MassWindowPPM)
                .MassSplineOrder = GetIniFileSettingInt(IniStuff, strSectionName, "MassSplineOrder", .MassSplineOrder)
                .MassNumXSlices = GetIniFileSettingInt(IniStuff, strSectionName, "MassNumXSlices", .MassNumXSlices)
                .MassNumMassDeltaBins = GetIniFileSettingInt(IniStuff, strSectionName, "MassNumMassDeltaBins", .MassNumMassDeltaBins)
                .MassMaxJump = GetIniFileSettingInt(IniStuff, strSectionName, "MassMaxJump", .MassMaxJump)
            
                .MassZScoreTolerance = GetIniFileSettingSng(IniStuff, strSectionName, "MassZScoreTolerance", .MassZScoreTolerance)
                .MassUseLSQ = GetIniFileSettingBln(IniStuff, strSectionName, "MassUseLSQ", .MassUseLSQ)
                .MassLSQOutlierZScore = GetIniFileSettingSng(IniStuff, strSectionName, "MassLSQOutlierZScore", .MassLSQOutlierZScore)
                .MassLSQNumKnots = GetIniFileSettingInt(IniStuff, strSectionName, "MassLSQNumKnots", .MassLSQNumKnots)
            End With
        End If
    End With
     
    With udtUMCInternalStandards
        intInternalStandardcount = GetIniFileSettingInt(IniStuff, "UMCInternalStandards", "Count", -1)
        
        If intInternalStandardcount = 0 Then
            .StandardsAreFromDB = False
            .Count = 0
            ReDim .InternalStandards(0)
        ElseIf intInternalStandardcount < 0 Then
            If .StandardsAreFromDB Then
                ' Leave the existing internal standards defined
            Else
                ' Reset the existing internal standards
                PopulateDefaultInternalStds udtUMCInternalStandards
            End If
        Else
            ' Look for and load each internal standard, only loading if present; load at most intInternalStandardcount entries (first entry is UMCInternalStandard1)
            ' First, initialize to having no internal standards
            .StandardsAreFromDB = False
            .Count = 0
            ReDim .InternalStandards(0 To intInternalStandardcount - 1)
            
            For intInternalStandardIndex = 0 To intInternalStandardcount - 1
                strSectionName = "UMCInternalStandard" & Trim(intInternalStandardIndex + 1)
                
                dblInternalStandardMass = GetIniFileSettingDbl(IniStuff, strSectionName, "MonoisotopicMass", 0)
                If dblInternalStandardMass > 0 Then
                    With .InternalStandards(.Count)
                        .SeqID = GetIniFileSetting(IniStuff, strSectionName, "SeqID", "")
                        .PeptideSequence = GetIniFileSetting(IniStuff, strSectionName, "PeptideSequence", "")
                        .MonoisotopicMass = dblInternalStandardMass
                        .NET = GetIniFileSettingDbl(IniStuff, strSectionName, "NET", 0)
                        .ChargeMinimum = GetIniFileSettingInt(IniStuff, strSectionName, "ChargeMinimum", 0)
                        .ChargeMaximum = GetIniFileSettingInt(IniStuff, strSectionName, "ChargeMaximum", 0)
                        .ChargeMostAbundant = GetIniFileSettingInt(IniStuff, strSectionName, "ChargeMostAbundant", 0)
                    End With
                    .Count = .Count + 1
                End If
            Next intInternalStandardIndex
        End If
    End With
     
     ' UMC Net Adjustment Options stored in udtPrefsExpanded
    With udtPrefsExpanded
        With .AutoAnalysisOptions
            .NETAdjustmentInitialNetTol = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentInitialNetTol", .NETAdjustmentInitialNetTol)
' November 2005: Unused variable     .NETAdjustmentFinalNetTol = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentFinalNetTol", .NETAdjustmentFinalNetTol)
            .NETAdjustmentMaxIterationCount = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentMaxIterationCount", .NETAdjustmentMaxIterationCount)
            .NETAdjustmentMinIDCount = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentMinIDCount", .NETAdjustmentMinIDCount)
            .NETAdjustmentMinIDCountAbsoluteMinimum = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentMinIDCountAbsoluteMinimum", .NETAdjustmentMinIDCountAbsoluteMinimum)
            .NETAdjustmentMinIterationCount = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentMinIterationCount", .NETAdjustmentMinIterationCount)
            .NETAdjustmentChangeThresholdStopValue = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentChangeThresholdStopValue", .NETAdjustmentChangeThresholdStopValue)
            
            .NETAdjustmentAutoIncrementUMCTopAbuPct = GetIniFileSettingBln(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentAutoIncrementUMCTopAbuPct", .NETAdjustmentAutoIncrementUMCTopAbuPct)
            .NETAdjustmentUMCTopAbuPctInitial = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentUMCTopAbuPctInitial", .NETAdjustmentUMCTopAbuPctInitial)
            .NETAdjustmentUMCTopAbuPctIncrement = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentUMCTopAbuPctIncrement", .NETAdjustmentUMCTopAbuPctIncrement)
            .NETAdjustmentUMCTopAbuPctMax = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentUMCTopAbuPctMax", .NETAdjustmentUMCTopAbuPctMax)
            
' November 2005: Unused variable    .NETAdjustmentMinimumNETMatchScore = GetIniFileSettingLng(IniStuff, strUMCNetAdjDefSectionName, "NETAdjustmentMinimumNETMatchScore", .NETAdjustmentMinimumNETMatchScore)
            
            .NETSlopeExpectedMinimum = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "NETSlopeExpectedMinimum", .NETSlopeExpectedMinimum)
            .NETSlopeExpectedMaximum = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "NETSlopeExpectedMaximum", .NETSlopeExpectedMaximum)
            .NETInterceptExpectedMinimum = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "NETInterceptExpectedMinimum", .NETInterceptExpectedMinimum)
            .NETInterceptExpectedMaximum = GetIniFileSettingDbl(IniStuff, strUMCNetAdjDefSectionName, "NETInterceptExpectedMaximum", .NETInterceptExpectedMaximum)
        End With
    End With
    
    If blnLegacySectionName Then
        ReDim strKeys(0)
        ReDim strValues(0)
        strKeys(0) = "NewSectionName"
        strValues(0) = "UMCNETAdjDef"
        IniStuff.WriteSection strUMCNetAdjDefSectionName, strKeys(), strValues()
    End If
    
    ' Search AMT preferences
    With udtAMTDef
        .SearchScope = .SearchScope     ' (Not stored in .Ini file)
        .SearchFlag = GetIniFileSettingInt(IniStuff, "SearchAMTDef", "SearchFlag", .SearchFlag)
        .MWField = GetIniFileSettingInt(IniStuff, "SearchAMTDef", "MWField", .MWField)
        .MWTol = GetIniFileSettingDbl(IniStuff, "SearchAMTDef", "MWTol", .MWTol)
        .NETorRT = GetIniFileSettingInt(IniStuff, "SearchAMTDef", "NETorRT", .NETorRT)
        .Formula = .Formula             ' (Not stored in .Ini file)
        .TolType = GetIniFileSettingInt(IniStuff, "SearchAMTDef", "TolType", .TolType)
        .NETTol = GetIniFileSettingDbl(IniStuff, "SearchAMTDef", "NETTol", .NETTol)
        .MassTag = GetIniFileSettingDbl(IniStuff, "SearchAMTDef", "MassTag", .MassTag)
        .MaxMassTags = GetIniFileSettingLng(IniStuff, "SearchAMTDef", "MaxMassTags", .MaxMassTags)
        .SkipReferenced = GetIniFileSettingBln(IniStuff, "SearchAMTDef", "SkipReferenced", .SkipReferenced)
        .SaveNCnt = GetIniFileSettingBln(IniStuff, "SearchAMTDef", "SaveNCnt", .SaveNCnt)
    End With
    
    ' Initialize the setting strings
    sCooSysPref = GetCooSysPrefs(udtPrefs)
    sDDClrPref = GetDDClrPrefs()
    sDrawingPref = GetDrawingPrefs(udtPrefs)
    sICR2LSPref = GetICR2LSPrefs()
    sBackForeCSIsoClrPref = GetOtherColorsPrefs()
    sCSIsoShapePref = GetCSIsoShapePrefs()
    sSwitchPref = GetSwitchPrefs(udtPrefs)
    sTolerancesPref = GetTolerancesPrefs(udtPrefs)
    
    ' No longer supported (March 2006)
    ''sAMTPref = GetAMTPrefs()
    ''sFTICR_AMTPref = GetFTICR_AMTPrefs()
    
    ' Preferences
    sCooSysPref = GetIniFileSetting(IniStuff, "Preferences", "CoordinateSystem", sCooSysPref)
    sDDClrPref = GetIniFileSetting(IniStuff, "Preferences", "DifferentialDisplay", sDDClrPref)
    sDrawingPref = GetIniFileSetting(IniStuff, "Preferences", "Drawing", sDrawingPref)
    sICR2LSPref = GetIniFileSetting(IniStuff, "Preferences", "ICR2LS", sICR2LSPref)
    sBackForeCSIsoClrPref = GetIniFileSetting(IniStuff, "Preferences", "ChargeStateColors", sBackForeCSIsoClrPref)
    sCSIsoShapePref = GetIniFileSetting(IniStuff, "Preferences", "ChargeStateShapes", sCSIsoShapePref)
    sSwitchPref = GetIniFileSetting(IniStuff, "Preferences", "Switches", sSwitchPref)
    sTolerancesPref = GetIniFileSetting(IniStuff, "Preferences", "Tolerances", sTolerancesPref)
    
    ' No longer supported (March 2006)
    ''sAMTPref = GetIniFileSetting(IniStuff, "Preferences", "AMTs", sAMTPref)
    ''sFTICR_AMTPref = GetIniFileSetting(IniStuff, "Preferences", "FTICRAmts", sFTICR_AMTPref)
    
    ResolveCooSysPrefs udtPrefs
    ResolveDDClrPrefs
    ResolveICR2LSPrefs
    ResolveOtherColorsPrefs
    
    ' If glCSColor is the old default (yellow) then update it to the new default (pink)
    If glCSColor = 65535 Then
        glCSColor = glCSColorDefault
    End If
    
    ResolveCSIsoShapePrefs
    ResolveDrawingPrefs udtPrefs
    ResolveSwitchPrefs udtPrefs
    ResolveTolerancesPrefs udtPrefs
    
    ' No longer supported (March 2006)
    ''ResolveAMTPrefs
    ''ResolveFTICR_AMTPrefs
    
    ' Be sure that we always start with FN or NET coordinate system
    If udtPrefs.CooType <> glFNCooSys And udtPrefs.CooType <> glNETCooSys Then
        udtPrefs.CooType = glFNCooSys
    End If

    ' Expanded preferences
    With udtPrefsExpanded
        .MenuModeDefault = GetIniFileSettingLng(IniStuff, "ExpandedPreferences", "MenuModeDefault", .MenuModeDefault)
        .MenuModeIncludeObsolete = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "MenuModeIncludeObsolete", .MenuModeIncludeObsolete)
        .ExtendedFileSaveModePreferred = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "ExtendedFileSaveModePreferred", .ExtendedFileSaveModePreferred)
        
        .AutoAdjSize = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "AutoAdjSize", .AutoAdjSize)
        .AutoSizeMultiplier = GetIniFileSettingDbl(IniStuff, "ExpandedPreferences", "AutoSizeMultiplier", CDbl(.AutoSizeMultiplier))
        .UMCDrawType = GetIniFileSettingLng(IniStuff, "ExpandedPreferences", "UMCDrawType", .UMCDrawType)
        
        .UsePEKBasedERValues = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "UsePEKBasedERValues", .UsePEKBasedERValues)
        .UseMassTagsWithNullMass = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "UseMassTagsWithNullMass", .UseMassTagsWithNullMass)
        .UseMassTagsWithNullNET = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "UseMassTagsWithNullNET", .UseMassTagsWithNullNET)
        .UseUMCConglomerateNET = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "UseUMCConglomerateNET", .UseUMCConglomerateNET)
        .NetAdjustmentUsesN15AMTMasses = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "NetAdjustmentUsesN15AMTMasses", .NetAdjustmentUsesN15AMTMasses)
        .NetAdjustmentMinHighNormalizedScore = GetIniFileSettingSng(IniStuff, "ExpandedPreferences", "NetAdjustmentMinHighNormalizedScore", .NetAdjustmentMinHighNormalizedScore)
        .NetAdjustmentMinHighDiscriminantScore = GetIniFileSettingSng(IniStuff, "ExpandedPreferences", "NetAdjustmentMinHighDiscriminantScore", .NetAdjustmentMinHighDiscriminantScore)
        
        .AMTSearchResultsBehavior = GetIniFileSettingInt(IniStuff, "ExpandedPreferences", "AMTSearchResultsBehavior", val(.AMTSearchResultsBehavior))
        .ICR2LSSpectrumViewZoomWindowWidthMZ = GetIniFileSettingDbl(IniStuff, "ExpandedPreferences", "ICR2LSSpectrumViewZoomWindowWidthMZ", .ICR2LSSpectrumViewZoomWindowWidthMZ)
        
        .LastAutoAnalysisIniFilePath = GetIniFileSetting(IniStuff, "ExpandedPreferences", "LastAutoAnalysisIniFilePath", .LastAutoAnalysisIniFilePath)
        .LastInputFileMode = GetIniFileSettingInt(IniStuff, "ExpandedPreferences", "LastInputFileMode", CInt(.LastInputFileMode))
        
        .LegacyAMTDBPath = GetIniFileSetting(IniStuff, "ExpandedPreferences", "LegacyAMTDBPath", .LegacyAMTDBPath)
    End With
    
    If Not udtPrefsExpanded.AutoAnalysisStatus.Enabled And Not bnlAutoAnalysisFieldsOnly And Not APP_BUILD_DISABLE_MTS Then
        ' Auto Query PRISM options
        ' Only update these if not currently auto analyzing
        With udtPrefsExpanded.AutoQueryPRISMOptions
            .ConnectionStringQueryDB = GetIniFileSetting(IniStuff, "AutoQueryPRISMOptions", "ConnectionStringQueryDB", .ConnectionStringQueryDB)
            .RequestTaskSPName = GetIniFileSetting(IniStuff, "AutoQueryPRISMOptions", "RequestTaskSPName", .RequestTaskSPName)
            .SetTaskCompleteSPName = GetIniFileSetting(IniStuff, "AutoQueryPRISMOptions", "SetTaskCompleteSPName", .SetTaskCompleteSPName)
            .SetTaskToRestartSPName = GetIniFileSetting(IniStuff, "AutoQueryPRISMOptions", "SetTaskToRestartSPName", .SetTaskToRestartSPName)
            .PostLogEntrySPName = GetIniFileSetting(IniStuff, "AutoQueryPRISMOptions", "PostLogEntrySPName", .PostLogEntrySPName)
            .QueryIntervalSeconds = GetIniFileSettingLng(IniStuff, "AutoQueryPRISMOptions", "QueryIntervalSeconds", .QueryIntervalSeconds)
            .MinimumPriorityToProcess = GetIniFileSettingInt(IniStuff, "AutoQueryPRISMOptions", "MinimumPriorityToProcess", .MinimumPriorityToProcess)
            .MaximumPriorityToProcess = GetIniFileSettingInt(IniStuff, "AutoQueryPRISMOptions", "MaximumPriorityToProcess", .MaximumPriorityToProcess)
            .PreferredDatabaseToProcess = GetIniFileSetting(IniStuff, "AutoQueryPRISMOptions", "PreferredDatabaseToProcess", .PreferredDatabaseToProcess)
            .ServerForPreferredDatabase = GetIniFileSetting(IniStuff, "AutoQueryPRISMOptions", "ServerForPreferredDatabase", .ServerForPreferredDatabase)
            .ExclusivelyUseThisDatabase = GetIniFileSettingBln(IniStuff, "AutoQueryPRISMOptions", "ExclusivelyUseThisDatabase", .ExclusivelyUseThisDatabase)
        End With
    End If
    
    With udtPrefsExpanded.NetAdjustmentUMCDistributionOptions
        .RequireDispersedUMCSelection = GetIniFileSettingBln(IniStuff, "NetAdjustmentUMCDistributionOptions", "RequireDispersedUMCSelection", .RequireDispersedUMCSelection)
        .SegmentCount = GetIniFileSettingLng(IniStuff, "NetAdjustmentUMCDistributionOptions", "SegmentCount", .SegmentCount)
        .MinimumUMCsPerSegmentPctTopAbuPct = GetIniFileSettingLng(IniStuff, "NetAdjustmentUMCDistributionOptions", "MinimumUMCsPerSegmentPctTopAbuPct", CLng(.MinimumUMCsPerSegmentPctTopAbuPct))
        .ScanPctStart = GetIniFileSettingLng(IniStuff, "NetAdjustmentUMCDistributionOptions", "ScanPctStart", CLng(.ScanPctStart))
        .ScanPctEnd = GetIniFileSettingLng(IniStuff, "NetAdjustmentUMCDistributionOptions", "ScanPctEnd", CLng(.ScanPctEnd))
    End With
    
    ' Error Distribution Preferences
    With udtPrefsExpanded.ErrorPlottingOptions
        .MassRangePPM = GetIniFileSettingLng(IniStuff, "ErrorPlottingOptions", "MassRangePPM", .MassRangePPM)
        .MassBinSizePPM = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptions", "MassBinSizePPM", .MassBinSizePPM)
        .GANETRange = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptions", "GANETRange", .GANETRange)
        .GANETBinSize = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptions", "GANETBinSize", .GANETBinSize)
        .ButterWorthFrequency = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptions", "ButterWorthFrequency", .ButterWorthFrequency)
        
        With .Graph2DOptions
            .ShowPointSymbols = GetIniFileSettingBln(IniStuff, "ErrorPlottingOptionsGraph2D", "ShowPointSymbols", .ShowPointSymbols)
            .DrawLinesBetweenPoints = GetIniFileSettingBln(IniStuff, "ErrorPlottingOptionsGraph2D", "DrawLinesBetweenPoints", .DrawLinesBetweenPoints)
            .ShowGridLines = GetIniFileSettingBln(IniStuff, "ErrorPlottingOptionsGraph2D", "ShowGridlines", .ShowGridLines)
            .AutoScaleXAxis = GetIniFileSettingBln(IniStuff, "ErrorPlottingOptionsGraph2D", "AutoScaleXAxis", .AutoScaleXAxis)
            .PointSizePixels = GetIniFileSettingLng(IniStuff, "ErrorPlottingOptionsGraph2D", "PointSizePixels", .PointSizePixels)
            .LineWidthPixels = GetIniFileSettingLng(IniStuff, "ErrorPlottingOptionsGraph2D", "LineWidthPixels", .LineWidthPixels)
            .CenterYAxis = GetIniFileSettingBln(IniStuff, "ErrorPlottingOptionsGraph2D", "CenterYAxis", .CenterYAxis)
            .ShowSmoothedData = GetIniFileSettingBln(IniStuff, "ErrorPlottingOptionsGraph2D", "ShowSmoothedData", .ShowSmoothedData)
            .ShowPeakEdges = GetIniFileSettingBln(IniStuff, "ErrorPlottingOptionsGraph2D", "ShowPeakEdges", .ShowPeakEdges)
        End With
        
        With .Graph3DOptions
            .ContourLevelsCount = GetIniFileSettingLng(IniStuff, "ErrorPlottingOptionsGraph3D", "ContourLevelsCount", .ContourLevelsCount)
            .Perspective = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptionsGraph3D", "Perspective", .Perspective)
            .Elevation = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptionsGraph3D", "Elevation", .Elevation)
            .YRotation = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptionsGraph3D", "YRotation", .YRotation)
            .ZRotation = GetIniFileSettingSng(IniStuff, "ErrorPlottingOptionsGraph3D", "ZRotation", .ZRotation)
            .AnnotationFontSize = GetIniFileSettingLng(IniStuff, "ErrorPlottingOptionsGraph3D", "AnnotationFontSize", .AnnotationFontSize)
        End With
    End With
    
    ' Noise removal options
    With udtPrefsExpanded.NoiseRemovalOptions
        .SearchTolerancePPMDefault = GetIniFileSettingDbl(IniStuff, "NoiseRemovalOptions", "SearchTolerancePPMDefault", .SearchTolerancePPMDefault)
        .SearchTolerancePPMAutoRemoval = GetIniFileSettingDbl(IniStuff, "NoiseRemovalOptions", "SearchTolerancePPMAutoRemoval", .SearchTolerancePPMAutoRemoval)
        
        .PercentageThresholdToExcludeSlice = GetIniFileSettingSng(IniStuff, "NoiseRemovalOptions", "PercentageThresholdToExcludeSlice", .PercentageThresholdToExcludeSlice)
        .PercentageThresholdToAddNeighborToSearchSlice = GetIniFileSettingSng(IniStuff, "NoiseRemovalOptions", "PercentageThresholdToAddNeighborToSearchSlice", .PercentageThresholdToAddNeighborToSearchSlice)
        
        .LimitMassRange = GetIniFileSettingBln(IniStuff, "NoiseRemovalOptions", "LimitMassRange", .LimitMassRange)
        .MassStart = GetIniFileSettingDbl(IniStuff, "NoiseRemovalOptions", "MassStart", .MassStart)
        .MassEnd = GetIniFileSettingDbl(IniStuff, "NoiseRemovalOptions", "MassEnd", .MassEnd)
        
        .LimitScanRange = GetIniFileSettingBln(IniStuff, "NoiseRemovalOptions", "LimitScanRange", .LimitScanRange)
        .ScanStart = GetIniFileSettingLng(IniStuff, "NoiseRemovalOptions", "ScanStart", .ScanStart)
        .ScanEnd = GetIniFileSettingLng(IniStuff, "NoiseRemovalOptions", "ScanEnd", .ScanEnd)
        
        .SearchScope = GetIniFileSettingInt(IniStuff, "NoiseRemovalOptions", "SearchScope", CInt(.SearchScope))
        .RequireIdenticalCharge = GetIniFileSettingBln(IniStuff, "NoiseRemovalOptions", "RequireIdenticalCharge", .RequireIdenticalCharge)
    End With
    
    ' Refine MS Data Options
    With udtPrefsExpanded.RefineMSDataOptions
        .MinimumPeakHeight = GetIniFileSettingLng(IniStuff, "RefineMSDataOptions", "MinimumPeakHeight", .MinimumPeakHeight)
        .MinimumSignalToNoiseRatioForLowAbundancePeaks = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "MinimumSignalToNoiseRatioForLowAbundancePeaks", .MinimumSignalToNoiseRatioForLowAbundancePeaks)
        .PercentageOfMaxForFindingWidth = GetIniFileSettingLng(IniStuff, "RefineMSDataOptions", "PercentageOfMaxForFindingWidth", .PercentageOfMaxForFindingWidth)
        .MassCalibrationMaximumShift = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "MassCalibrationMaximumShift", .MassCalibrationMaximumShift)
        .MassCalibrationTolType = GetIniFileSettingInt(IniStuff, "RefineMSDataOptions", "MassCalibrationTolType", CInt(.MassCalibrationTolType))
        
        ' Note: MassToleranceRefinementMethod was renamed to ToleranceRefinementMethod in November 2006; allowing for both names here, with ToleranceRefinementMethod taking precedence
        .ToleranceRefinementMethod = GetIniFileSettingInt(IniStuff, "RefineMSDataOptions", "MassToleranceRefinementMethod", CInt(.ToleranceRefinementMethod))
        .ToleranceRefinementMethod = GetIniFileSettingInt(IniStuff, "RefineMSDataOptions", "ToleranceRefinementMethod", CInt(.ToleranceRefinementMethod))
        
        .UseMinMaxIfOutOfRange = GetIniFileSettingBln(IniStuff, "RefineMSDataOptions", "UseMinMaxIfOutOfRange", CInt(.UseMinMaxIfOutOfRange))
        
        .MassToleranceMinimum = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "MassToleranceMinimum", .MassToleranceMinimum)
        .MassToleranceMaximum = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "MassToleranceMaximum", .MassToleranceMaximum)
        .MassToleranceAdjustmentMultiplier = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "MassToleranceAdjustmentMultiplier", .MassToleranceAdjustmentMultiplier)
        .NETToleranceMinimum = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "NETToleranceMinimum", .NETToleranceMinimum)
        .NETToleranceMaximum = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "NETToleranceMaximum", .NETToleranceMaximum)
        .NETToleranceAdjustmentMultiplier = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "NETToleranceAdjustmentMultiplier", .NETToleranceAdjustmentMultiplier)
        
        ' Parameter IncludeNetLockerMatches was renamed to IncludeNetLockerMatches; allowing for both names here, with IncludeInternalStdMatches taking precedence
        .IncludeInternalStdMatches = GetIniFileSettingBln(IniStuff, "RefineMSDataOptions", "IncludeNetLockerMatches", .IncludeInternalStdMatches)
        .IncludeInternalStdMatches = GetIniFileSettingBln(IniStuff, "RefineMSDataOptions", "IncludeInternalStdMatches", .IncludeInternalStdMatches)
        
        .UseUMCClassStats = GetIniFileSettingBln(IniStuff, "RefineMSDataOptions", "UseUMCClassStats", .UseUMCClassStats)
        .MinimumSLiC = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "MinimumSLiC", .MinimumSLiC)
        .MaximumAbundance = GetIniFileSettingDbl(IniStuff, "RefineMSDataOptions", "MaximumAbundance", .MaximumAbundance)
        
        .EMMassErrorPeakToleranceEstimatePPM = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "EMMassErrorPeakToleranceEstimatePPM", .EMMassErrorPeakToleranceEstimatePPM)
        .EMNETErrorPeakToleranceEstimate = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "EMNETErrorPeakToleranceEstimate", .EMNETErrorPeakToleranceEstimate)
        .EMIterationCount = GetIniFileSettingInt(IniStuff, "RefineMSDataOptions", "EMIterationCount", .EMIterationCount)
        .EMPercentOfDataToExclude = GetIniFileSettingInt(IniStuff, "RefineMSDataOptions", "EMPercentOfDataToExclude", .EMPercentOfDataToExclude)
        
        .EMMassTolRefineForceUseSingleDataPointErrors = GetIniFileSettingBln(IniStuff, "RefineMSDataOptions", "EMMassTolRefineForceUseSingleDataPointErrors", .EMMassTolRefineForceUseSingleDataPointErrors)
        .EMNETTolRefineForceUseSingleDataPointErrors = GetIniFileSettingBln(IniStuff, "RefineMSDataOptions", "EMNETTolRefineForceUseSingleDataPointErrors", .EMNETTolRefineForceUseSingleDataPointErrors)
    End With
    
    ' TIC and BPI Plotting Options
    With udtPrefsExpanded.TICAndBPIPlottingOptions
        .PlotNETOnXAxis = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "PlotNETOnXAxis", .PlotNETOnXAxis)
        .NormalizeYAxis = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "NormalizeYAxis", .NormalizeYAxis)
        .SmoothUsingMovingAverage = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "SmoothUsingMovingAverage", .SmoothUsingMovingAverage)
        .MovingAverageWindowWidth = GetIniFileSettingLng(IniStuff, "TICAndBPIPlottingOptions", "MovingAverageWindowWidth", .MovingAverageWindowWidth)
        .TimeDomainDataMaxValue = GetIniFileSettingDbl(IniStuff, "TICAndBPIPlottingOptions", "TimeDomainDataMaxValue", .TimeDomainDataMaxValue)
        With .Graph2DOptions
            .ShowPointSymbols = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "ShowPointSymbols", .ShowPointSymbols)
            .DrawLinesBetweenPoints = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "DrawLinesBetweenPoints", .DrawLinesBetweenPoints)
            .ShowGridLines = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "ShowGridlines", .ShowGridLines)
            .AutoScaleXAxis = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "AutoScaleXAxis", .AutoScaleXAxis)
            .PointSizePixels = GetIniFileSettingLng(IniStuff, "TICAndBPIPlottingOptions", "PointSizePixels", .PointSizePixels)
            .PointShape = GetIniFileSettingInt(IniStuff, "TICAndBPIPlottingOptions", "PointShape", .PointShape)
            .PointAndLineColor = GetIniFileSettingLng(IniStuff, "TICAndBPIPlottingOptions", "PointAndLineColor", .PointAndLineColor)
            .LineWidthPixels = GetIniFileSettingLng(IniStuff, "TICAndBPIPlottingOptions", "LineWidthPixels", .LineWidthPixels)
            .CenterYAxis = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "CenterYAxis", .CenterYAxis)
        End With
        
        .PointShapeSeries2 = GetIniFileSettingInt(IniStuff, "TICAndBPIPlottingOptions", "PointShapeSeries2", .PointShapeSeries2)
        .PointAndLineColorSeries2 = GetIniFileSettingLng(IniStuff, "TICAndBPIPlottingOptions", "PointAndLineColorSeries2", .PointAndLineColorSeries2)
        
        ' Skip this: .ClipOutliers = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "ClipOutliers", .ClipOutliers)
        .ClipOutliersFactor = GetIniFileSettingSng(IniStuff, "TICAndBPIPlottingOptions", "ClipOutliersFactor", .ClipOutliersFactor)
        
        .KeepWindowOnTop = GetIniFileSettingBln(IniStuff, "TICAndBPIPlottingOptions", "KeepWindowOnTop", .KeepWindowOnTop)
    End With
    
    ' Pair Browser Options
    With udtPrefsExpanded.PairBrowserPlottingOptions
        .SortOrder = GetIniFileSettingInt(IniStuff, "PairBrowserOptions", "SortOrder", .SortOrder)
        .SortDescending = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "SortDescending", .SortDescending)
        .AutoZoom2DPlot = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "AutoZoom2DPlot", .AutoZoom2DPlot)
        .HighlightMembers = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "HighlightMembers", .HighlightMembers)
        .PlotAllChargeStates = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "PlotAllChargeStates", .PlotAllChargeStates)
        
        .FixedDimensionsForAutoZoom = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "FixedDimensionsForAutoZoom", .FixedDimensionsForAutoZoom)
        .MassRangeZoom = GetIniFileSettingDbl(IniStuff, "PairBrowserOptions", "MassRangeZoom", .MassRangeZoom)
        .MassRangeUnits = GetIniFileSettingInt(IniStuff, "PairBrowserOptions", "MassRangeUnits", .MassRangeUnits)
        .ScanRangeZoom = GetIniFileSettingDbl(IniStuff, "PairBrowserOptions", "ScanRangeZoom", .ScanRangeZoom)
        .ScanRangeUnits = GetIniFileSettingInt(IniStuff, "PairBrowserOptions", "ScanRangeUnits", .ScanRangeUnits)
        With .Graph2DOptions
            .ShowPointSymbols = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "ShowPointSymbols", .ShowPointSymbols)
            .DrawLinesBetweenPoints = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "DrawLinesBetweenPoints", .DrawLinesBetweenPoints)
            .ShowGridLines = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "ShowGridlines", .ShowGridLines)
            .PointSizePixels = GetIniFileSettingLng(IniStuff, "PairBrowserOptions", "PointSizePixels", .PointSizePixels)
            .PointShape = GetIniFileSettingInt(IniStuff, "PairBrowserOptions", "PointShape", .PointShape)
            .PointAndLineColor = GetIniFileSettingLng(IniStuff, "PairBrowserOptions", "PointAndLineColor", .PointAndLineColor)
            .LineWidthPixels = GetIniFileSettingLng(IniStuff, "PairBrowserOptions", "LineWidthPixels", .LineWidthPixels)
        End With
        .PointShapeHeavy = GetIniFileSettingInt(IniStuff, "PairBrowserOptions", "PointShapeHeavy", .PointShapeHeavy)
        .PointAndLineColorHeavy = GetIniFileSettingLng(IniStuff, "PairBrowserOptions", "PointAndLineColorHeavy", .PointAndLineColorHeavy)
        .KeepWindowOnTop = GetIniFileSettingBln(IniStuff, "PairBrowserOptions", "KeepWindowOnTop", .KeepWindowOnTop)
    End With
    
    
    ' UMC Browser Options
    With udtPrefsExpanded.UMCBrowserPlottingOptions
        .SortOrder = GetIniFileSettingInt(IniStuff, "UMCBrowserOptions", "SortOrder", .SortOrder)
        .SortDescending = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "SortDescending", .SortDescending)
        .AutoZoom2DPlot = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "AutoZoom2DPlot", .AutoZoom2DPlot)
        .HighlightMembers = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "HighlightMembers", .HighlightMembers)
        .PlotAllChargeStates = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "PlotAllChargeStates", .PlotAllChargeStates)
        
        .FixedDimensionsForAutoZoom = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "FixedDimensionsForAutoZoom", .FixedDimensionsForAutoZoom)
        .MassRangeZoom = GetIniFileSettingDbl(IniStuff, "UMCBrowserOptions", "MassRangeZoom", .MassRangeZoom)
        .MassRangeUnits = GetIniFileSettingInt(IniStuff, "UMCBrowserOptions", "MassRangeUnits", .MassRangeUnits)
        .ScanRangeZoom = GetIniFileSettingDbl(IniStuff, "UMCBrowserOptions", "ScanRangeZoom", .ScanRangeZoom)
        .ScanRangeUnits = GetIniFileSettingInt(IniStuff, "UMCBrowserOptions", "ScanRangeUnits", .ScanRangeUnits)
        With .Graph2DOptions
            .ShowPointSymbols = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "ShowPointSymbols", .ShowPointSymbols)
            .DrawLinesBetweenPoints = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "DrawLinesBetweenPoints", .DrawLinesBetweenPoints)
            .ShowGridLines = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "ShowGridlines", .ShowGridLines)
            .PointSizePixels = GetIniFileSettingLng(IniStuff, "UMCBrowserOptions", "PointSizePixels", .PointSizePixels)
            .PointShape = GetIniFileSettingInt(IniStuff, "UMCBrowserOptions", "PointShape", .PointShape)
            .PointAndLineColor = GetIniFileSettingLng(IniStuff, "UMCBrowserOptions", "PointAndLineColor", .PointAndLineColor)
            .LineWidthPixels = GetIniFileSettingLng(IniStuff, "UMCBrowserOptions", "LineWidthPixels", .LineWidthPixels)
        End With
        .KeepWindowOnTop = GetIniFileSettingBln(IniStuff, "UMCBrowserOptions", "KeepWindowOnTop", .KeepWindowOnTop)
    End With
    
    
    ' Pair Identification and Search options
    With udtPrefsExpanded.PairSearchOptions
        With .SearchDef
            .DeltaMass = GetIniFileSettingDbl(IniStuff, "PairSearchOptions", "DeltaMass", .DeltaMass)
            .DeltaMassTolerance = GetIniFileSettingDbl(IniStuff, "PairSearchOptions", "DeltaMassTolerance", .DeltaMassTolerance)
            .AutoCalculateDeltaMinMaxCount = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AutoCalculateDeltaMinMaxCount", .AutoCalculateDeltaMinMaxCount)
            .DeltaCountMin = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "DeltaCountMin", .DeltaCountMin)
            .DeltaCountMax = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "DeltaCountMax", .DeltaCountMax)
            .DeltaStepSize = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "DeltaStepSize", .DeltaStepSize)
            
            .LightLabelMass = GetIniFileSettingDbl(IniStuff, "PairSearchOptions", "LightLabelMass", .LightLabelMass)
            .HeavyLightMassDifference = GetIniFileSettingDbl(IniStuff, "PairSearchOptions", "HeavyLightMassDifference", .HeavyLightMassDifference)
            .LabelCountMin = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "LabelCountMin", .LabelCountMin)
            .LabelCountMax = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "LabelCountMax", .LabelCountMax)
            .MaxDifferenceInNumberOfLightHeavyLabels = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "MaxDifferenceInNumberOfLightHeavyLabels", .MaxDifferenceInNumberOfLightHeavyLabels)
            
            .RequireUMCOverlap = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RequireUMCOverlap", .RequireUMCOverlap)
            .RequireUMCOverlapAtApex = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RequireUMCOverlapAtApex", .RequireUMCOverlapAtApex)
            
            .ScanTolerance = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "ScanTolerance", .ScanTolerance)
            .ScanToleranceAtApex = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "ScanToleranceAtApex", .ScanToleranceAtApex)
            
            .ERInclusionMin = GetIniFileSettingDbl(IniStuff, "PairSearchOptions", "ERInclusionMin", .ERInclusionMin)
            .ERInclusionMax = GetIniFileSettingDbl(IniStuff, "PairSearchOptions", "ERInclusionMax", .ERInclusionMax)
            
            .RequireMatchingChargeStatesForPairMembers = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RequireMatchingChargeStatesForPairMembers", .RequireMatchingChargeStatesForPairMembers)
            .UseIdenticalChargesForER = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "UseIdenticalChargesForER", .UseIdenticalChargesForER)
            .ComputeERScanByScan = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "ComputeERScanByScan", .ComputeERScanByScan)
            .AverageERsAllChargeStates = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AverageERsAllChargeStates", .AverageERsAllChargeStates)
            .AverageERsWeightingMode = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "AverageERsWeightingMode", CInt(.AverageERsWeightingMode))
            .ERCalcType = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "ERCalcType", CInt(.ERCalcType))
            
            .RemoveOutlierERs = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RemoveOutlierERs", .RemoveOutlierERs)
            .RemoveOutlierERsIterate = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RemoveOutlierERsIterate", .RemoveOutlierERsIterate)
            .RemoveOutlierERsMinimumDataPointCount = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "RemoveOutlierERsMinimumDataPointCount", .RemoveOutlierERsMinimumDataPointCount)
            .RemoveOutlierERsConfidenceLevel = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "RemoveOutlierERsConfidenceLevel", .RemoveOutlierERsConfidenceLevel)
        End With

        .PairSearchMode = GetIniFileSetting(IniStuff, "PairSearchOptions", "PairSearchMode", .PairSearchMode)
        
        .AutoExcludeOutOfERRange = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AutoExcludeOutOfERRange", .AutoExcludeOutOfERRange)
        .AutoExcludeAmbiguous = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AutoExcludeAmbiguous", .AutoExcludeAmbiguous)
        .KeepMostConfidentAmbiguous = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "KeepMostConfidentAmbiguous", .KeepMostConfidentAmbiguous)
        
        .AutoAnalysisRemovePairMemberHitsAfterDBSearch = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AutoAnalysisRemovePairMemberHitsAfterDBSearch", .AutoAnalysisRemovePairMemberHitsAfterDBSearch)
        .AutoAnalysisRemovePairMemberHitsRemoveHeavy = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AutoAnalysisRemovePairMemberHitsRemoveHeavy", .AutoAnalysisRemovePairMemberHitsRemoveHeavy)
        
        .AutoAnalysisSavePairsToTextFile = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AutoAnalysisSavePairsToTextFile", .AutoAnalysisSavePairsToTextFile)
        .AutoAnalysisSavePairsStatisticsToTextFile = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AutoAnalysisSavePairsStatisticsToTextFile", .AutoAnalysisSavePairsStatisticsToTextFile)
    
        .NETAdjustmentPairedSearchUMCSelection = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "NETAdjustmentPairedSearchUMCSelection", CInt(.NETAdjustmentPairedSearchUMCSelection))
        .OutlierRemovalUsesSymmetricERs = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "OutlierRemovalUsesSymmetricERs", .OutlierRemovalUsesSymmetricERs)
        
        .AutoAnalysisDeltaMassAddnlCount = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "AutoAnalysisDeltaMassAddnlCount", .AutoAnalysisDeltaMassAddnlCount)
        
        If .AutoAnalysisDeltaMassAddnlCount <= 0 Then
            ReDim .AutoAnalysisDeltaMassAddnl(0)
        Else
            ReDim .AutoAnalysisDeltaMassAddnl(.AutoAnalysisDeltaMassAddnlCount - 1)
        End If
        
        For intIndex = 0 To .AutoAnalysisDeltaMassAddnlCount - 1
            strKeyPrefix = "AutoAnalysisDeltaMassAddnl" & Trim(intIndex + 1)
        
            strKeyValue = GetIniFileSetting(IniStuff, "PairSearchOptions", strKeyPrefix, ENTRY_NOT_FOUND)
            If strKeyValue = ENTRY_NOT_FOUND Then
                .AutoAnalysisDeltaMassAddnlCount = intIndex
                Exit For
            Else
                If IsNumeric(strKeyValue) Then
                    .AutoAnalysisDeltaMassAddnl(intIndex) = CDbl(strKeyValue)
                Else
                    .AutoAnalysisDeltaMassAddnl(intIndex) = 0
                End If
            End If
        Next intIndex
        
    End With
    
    ' IReport Pair options
    With udtPrefsExpanded.PairSearchOptions.SearchDef.IReportEROptions
        .Enabled = GetIniFileSettingBln(IniStuff, "IReportEROptions", "Enabled", .Enabled)
        .NaturalAbundanceRatio2Coeff.Exponent = GetIniFileSettingDbl(IniStuff, "IReportEROptions", "NaturalAbundanceRatio2CoeffExponent", .NaturalAbundanceRatio2Coeff.Exponent)
        .NaturalAbundanceRatio2Coeff.Multiplier = GetIniFileSettingDbl(IniStuff, "IReportEROptions", "NaturalAbundanceRatio2CoeffMultiplier", .NaturalAbundanceRatio2Coeff.Multiplier)
        .NaturalAbundanceRatio4Coeff.Exponent = GetIniFileSettingDbl(IniStuff, "IReportEROptions", "NaturalAbundanceRatio4CoeffExponent", .NaturalAbundanceRatio4Coeff.Exponent)
        .NaturalAbundanceRatio4Coeff.Multiplier = GetIniFileSettingDbl(IniStuff, "IReportEROptions", "NaturalAbundanceRatio4CoeffMultiplier", .NaturalAbundanceRatio4Coeff.Multiplier)
        .MinimumFractionScansWithValidER = GetIniFileSettingSng(IniStuff, "IReportEROptions", "MinimumFractionScansWithValidER", .MinimumFractionScansWithValidER)
    End With
    
    ' MT tag Staleness options
    With udtPrefsExpanded.MassTagStalenessOptions
        .MaximumAgeLoadedMassTagsHours = GetIniFileSettingLng(IniStuff, "MassTagStalenessOptions", "MaximumAgeLoadedMassTagsHours", .MaximumAgeLoadedMassTagsHours)
        .MaximumFractionAMTsWithNulls = GetIniFileSettingDbl(IniStuff, "MassTagStalenessOptions", "MaximumFractionAMTsWithNulls", CDbl(.MaximumFractionAMTsWithNulls))
                                       
        .MaximumCountAMTsWithNulls = GetIniFileSettingLng(IniStuff, "MassTagStalenessOptions", "MaximumCountAMTsWithNulls", .MaximumCountAMTsWithNulls)
        .MinimumTimeBetweenReloadMinutes = GetIniFileSettingLng(IniStuff, "MassTagStalenessOptions", "MinimumTimeBetweenReloadMinutes", .MinimumTimeBetweenReloadMinutes)
    End With
    
    ' MT tag Match Score options
    With udtPrefsExpanded.SLiCScoreOptions
        .MassPPMStDev = GetIniFileSettingDbl(IniStuff, "SLiCScoreOptions", "MassPPMStDev", .MassPPMStDev)
        .NETStDev = GetIniFileSettingDbl(IniStuff, "SLiCScoreOptions", "NETStDev", .NETStDev)
        .UseAMTNETStDev = GetIniFileSettingBln(IniStuff, "SLiCScoreOptions", "UseAMTNETStDev", .UseAMTNETStDev)
        .MaxSearchDistanceMultiplier = GetIniFileSettingInt(IniStuff, "SLiCScoreOptions", "MaxSearchDistanceMultiplier", .MaxSearchDistanceMultiplier)
        
        strKeyValue = GetIniFileSetting(IniStuff, "SLiCScoreOptions", "AutoDefineSLiCScoreThresholds", ENTRY_NOT_FOUND)
        If strKeyValue = ENTRY_NOT_FOUND Then
            ' Assume default settings since entry isn't present
            .MaxSearchDistanceMultiplier = 2
            .AutoDefineSLiCScoreThresholds = True
        Else
            .AutoDefineSLiCScoreThresholds = CBool(strKeyValue)
        End If
    End With
    
    ' EditCopy Options
    With udtPrefsExpanded.GraphicExportOptions
        .CopyEMFIncludeFilenameAndDate = GetIniFileSettingBln(IniStuff, "GraphicExportOptions", "CopyEMFIncludeFilenameAndDate", .CopyEMFIncludeFilenameAndDate)
        .CopyEMFIncludeTextLabels = GetIniFileSettingBln(IniStuff, "GraphicExportOptions", "CopyEMFIncludeTextLabels", .CopyEMFIncludeTextLabels)
        
        SetEditCopyEMFOptions .CopyEMFIncludeFilenameAndDate, .CopyEMFIncludeTextLabels
    End With
    
    ' Auto Analysis Preferences
    With udtPrefsExpanded
        With .AutoAnalysisOptions.AutoToleranceRefinement
            .DBSearchMWTol = GetIniFileSettingDbl(IniStuff, "AutoToleranceRefinement", "DBSearchMWTol", .DBSearchMWTol)
            .DBSearchTolType = GetIniFileSettingInt(IniStuff, "AutoToleranceRefinement", "DBSearchTolType", CInt(.DBSearchTolType))
            .DBSearchNETTol = GetIniFileSettingDbl(IniStuff, "AutoToleranceRefinement", "DBSearchNETTol", .DBSearchNETTol)
            
            .DBSearchRegionShape = GetIniFileSettingInt(IniStuff, "AutoToleranceRefinement", "DBSearchRegionShape", CInt(.DBSearchRegionShape))
            
            .DBSearchMinimumHighNormalizedScore = GetIniFileSettingSng(IniStuff, "AutoToleranceRefinement", "DBSearchMinimumHighNormalizedScore", .DBSearchMinimumHighNormalizedScore)
            .DBSearchMinimumHighDiscriminantScore = GetIniFileSettingSng(IniStuff, "AutoToleranceRefinement", "DBSearchMinimumHighDiscriminantScore", .DBSearchMinimumHighDiscriminantScore)
            
            ' Note: if DBSearchMinimumPeptideProphetProbability is missing from the .Ini file, we're assuming a value of 0, not a value of .DBSearchMinimumPeptideProphetProbability
            ' This is done to assure backwards compatibility
            .DBSearchMinimumPeptideProphetProbability = GetIniFileSettingSng(IniStuff, "AutoToleranceRefinement", "DBSearchMinimumPeptideProphetProbability", 0)
            
            .RefineMassCalibration = GetIniFileSettingBln(IniStuff, "AutoToleranceRefinement", "RefineMassCalibration", .RefineMassCalibration)
            .RefineMassCalibrationOverridePPM = GetIniFileSettingDbl(IniStuff, "AutoToleranceRefinement", "RefineMassCalibrationOverridePPM", .RefineMassCalibrationOverridePPM)
            .RefineDBSearchMassTolerance = GetIniFileSettingBln(IniStuff, "AutoToleranceRefinement", "RefineDBSearchMassTolerance", .RefineDBSearchMassTolerance)
            .RefineDBSearchNETTolerance = GetIniFileSettingBln(IniStuff, "AutoToleranceRefinement", "RefineDBSearchNETTolerance", .RefineDBSearchNETTolerance)
        End With
        
        With .AutoAnalysisOptions
            .MDType = GetIniFileSettingLng(IniStuff, "AutoAnalysisOptions", "MDType", .MDType)
            .AutoRemoveNoiseStreaks = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "AutoRemoveNoiseStreaks", .AutoRemoveNoiseStreaks)
            .DoNotSaveOrExport = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "DoNotSaveOrExport", .DoNotSaveOrExport)
            
            .SkipFindUMCs = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SkipFindUMCs", .SkipFindUMCs)
            .SkipGANETSlopeAndInterceptComputation = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SkipGANETSlopeAndInterceptComputation", .SkipGANETSlopeAndInterceptComputation)

            .DBConnectionRetryAttemptMax = GetIniFileSettingInt(IniStuff, "AutoAnalysisOptions", "DBConnectionRetryAttemptMax", .DBConnectionRetryAttemptMax)
            .DBConnectionTimeoutSeconds = GetIniFileSettingInt(IniStuff, "AutoAnalysisOptions", "DBConnectionTimeoutSeconds", .DBConnectionTimeoutSeconds)
            .ExportResultsFileUsesJobNumberInsteadOfDataSetName = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "ExportResultsFileUsesJobNumberInsteadOfDataSetName", .ExportResultsFileUsesJobNumberInsteadOfDataSetName)
            
            .SaveGelFile = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SaveGelFile", .SaveGelFile)
            .SaveGelFileOnError = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SaveGelFileOnError", .SaveGelFileOnError)
            .SavePictureGraphic = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePictureGraphic", .SavePictureGraphic)
            .SavePictureGraphicFileType = GetIniFileSettingInt(IniStuff, "AutoAnalysisOptions", "SavePictureGraphicFileType", CInt(.SavePictureGraphicFileType))
            .SavePictureWidthPixels = GetIniFileSettingLng(IniStuff, "AutoAnalysisOptions", "SavePictureWidthPixels", .SavePictureWidthPixels)
            .SavePictureHeightPixels = GetIniFileSettingLng(IniStuff, "AutoAnalysisOptions", "SavePictureHeightPixels", .SavePictureHeightPixels)
            
            .SaveInternalStdHitsAndData = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SaveInternalStdHitsAndData", .SaveInternalStdHitsAndData)
            
            .SaveErrorGraphicMass = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SaveErrorGraphicMass", .SaveErrorGraphicMass)
            .SaveErrorGraphicGANET = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SaveErrorGraphicGANET", .SaveErrorGraphicGANET)
            .SaveErrorGraphic3D = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SaveErrorGraphic3D", .SaveErrorGraphic3D)
            .SaveErrorGraphicFileType = GetIniFileSettingInt(IniStuff, "AutoAnalysisOptions", "SaveErrorGraphicFileType", CInt(.SaveErrorGraphicFileType))
            .SaveErrorGraphSizeWidthPixels = GetIniFileSettingLng(IniStuff, "AutoAnalysisOptions", "SaveErrorGraphSizeWidthPixels", .SaveErrorGraphSizeWidthPixels)
            .SaveErrorGraphSizeHeightPixels = GetIniFileSettingLng(IniStuff, "AutoAnalysisOptions", "SaveErrorGraphSizeHeightPixels", .SaveErrorGraphSizeHeightPixels)
            
            .SavePlotTIC = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotTIC", .SavePlotTIC)
            .SavePlotBPI = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotBPI", .SavePlotBPI)
            .SavePlotTICTimeDomain = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotTICTimeDomain", .SavePlotTICTimeDomain)
            .SavePlotTICDataPointCounts = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotTICDataPointCounts", .SavePlotTICDataPointCounts)
            .SavePlotTICDataPointCountsHitsOnly = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotTICDataPointCountsHitsOnly", .SavePlotTICDataPointCountsHitsOnly)
            .SavePlotTICFromRawData = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotTICFromRawData", .SavePlotTICFromRawData)
            .SavePlotBPIFromRawData = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotBPIFromRawData", .SavePlotBPIFromRawData)
            .SavePlotDeisotopingIntensityThresholds = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotDeisotopingIntensityThresholds", .SavePlotDeisotopingIntensityThresholds)
            .SavePlotDeisotopingPeakCounts = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SavePlotDeisotopingPeakCounts", .SavePlotDeisotopingPeakCounts)
            
            .OutputFileSeparationCharacter = GetIniFileSetting(IniStuff, "AutoAnalysisOptions", "OutputFileSeparationCharacter", .OutputFileSeparationCharacter)
            .PEKFileExtensionPreferenceOrder = GetIniFileSetting(IniStuff, "AutoAnalysisOptions", "PEKFileExtensionPreferenceOrder", .PEKFileExtensionPreferenceOrder)
            
            .WriteIDResultsByIonToTextFileAfterAutoSearches = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "WriteIDResultsByIonToTextFileAfterAutoSearches", .WriteIDResultsByIonToTextFileAfterAutoSearches)
            .SaveUMCStatisticsToTextFile = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SaveUMCStatisticsToTextFile", .SaveUMCStatisticsToTextFile)
            .IncludeORFNameInTextFileOutput = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "IncludeORFNameInTextFileOutput", .IncludeORFNameInTextFileOutput)
            .SetIsConfirmedForDBSearchMatches = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "SetIsConfirmedForDBSearchMatches", .SetIsConfirmedForDBSearchMatches)
            .AddQuantitationDescriptionEntry = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "AddQuantitationDescriptionEntry", .AddQuantitationDescriptionEntry)
            .ExportUMCsWithNoMatches = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "ExportUMCsWithNoMatches", .ExportUMCsWithNoMatches)
            
            .DBSearchRegionShape = GetIniFileSettingInt(IniStuff, "AutoAnalysisOptions", "DBSearchRegionShape", CInt(.DBSearchRegionShape))
            .UseLegacyDBForMTs = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "UseLegacyDBForMTs", .UseLegacyDBForMTs)
            .IgnoreNETAdjustmentFailure = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "IgnoreNETAdjustmentFailure", .IgnoreNETAdjustmentFailure)
            
            .AutoAnalysisSearchModeCount = GetIniFileSettingInt(IniStuff, "AutoAnalysisOptions", "AutoAnalysisSearchModeCount", .AutoAnalysisSearchModeCount)
            If .AutoAnalysisSearchModeCount > MAX_AUTO_SEARCH_MODE_COUNT Then .AutoAnalysisSearchModeCount = MAX_AUTO_SEARCH_MODE_COUNT
            
            ' Load the first Search Mode, using defaults if missing
            With .AutoAnalysisSearchMode(0)
                strSectionName = "AutoAnalysisSearchMode1"
                .SearchMode = GetIniFileSetting(IniStuff, strSectionName, "SearchMode", .SearchMode)
                .AlternateOutputFolderPath = GetIniFileSetting(IniStuff, strSectionName, "AlternateOutputFolderPath", .AlternateOutputFolderPath)
                .WriteResultsToTextFile = GetIniFileSettingBln(IniStuff, strSectionName, "WriteResultsToTextFile", .WriteResultsToTextFile)
                .ExportResultsToDatabase = GetIniFileSettingBln(IniStuff, strSectionName, "ExportResultsToDatabase", .ExportResultsToDatabase)
                .ExportUMCMembers = GetIniFileSettingBln(IniStuff, strSectionName, "ExportUMCMembers", .ExportUMCMembers)
                .PairSearchAssumeMassTagsAreLabeled = GetIniFileSettingBln(IniStuff, strSectionName, "PairSearchAssumeMassTagsAreLabeled", .PairSearchAssumeMassTagsAreLabeled)
                
                ' Parameter GANETLockerSearchMode was renamed to InternalStdSearchMode; allowing for both names here, with InternalStdSearchMode taking precedence
                .InternalStdSearchMode = GetIniFileSettingInt(IniStuff, strSectionName, "GANETLockerSearchMode", CInt(.InternalStdSearchMode))
                .InternalStdSearchMode = GetIniFileSettingInt(IniStuff, strSectionName, "InternalStdSearchMode", CInt(.InternalStdSearchMode))
                
                .DBSearchMinimumHighNormalizedScore = GetIniFileSettingSng(IniStuff, strSectionName, "DBSearchMinimumHighNormalizedScore", .DBSearchMinimumHighNormalizedScore)
                .DBSearchMinimumHighDiscriminantScore = GetIniFileSettingSng(IniStuff, strSectionName, "DBSearchMinimumHighDiscriminantScore", .DBSearchMinimumHighDiscriminantScore)
                .DBSearchMinimumPeptideProphetProbability = GetIniFileSettingSng(IniStuff, strSectionName, "DBSearchMinimumPeptideProphetProbability", .DBSearchMinimumPeptideProphetProbability)

                With .MassMods
                    .DynamicMods = GetIniFileSettingBln(IniStuff, strSectionName, "DynamicMods", .DynamicMods)
                    .N15InsteadOfN14 = GetIniFileSettingBln(IniStuff, strSectionName, "N15InsteadOfN14", .N15InsteadOfN14)
                    .PEO = GetIniFileSettingBln(IniStuff, strSectionName, "PEO", .PEO)
                    .ICATd0 = GetIniFileSettingBln(IniStuff, strSectionName, "ICATd0", .ICATd0)
                    .ICATd8 = GetIniFileSettingBln(IniStuff, strSectionName, "ICATd8", .ICATd8)
                    .Alkylation = GetIniFileSettingBln(IniStuff, strSectionName, "Alkylation", .Alkylation)
                    .AlkylationMass = GetIniFileSettingDbl(IniStuff, strSectionName, "CustomMass", .AlkylationMass)
                    .ResidueToModify = GetIniFileSetting(IniStuff, strSectionName, "ResidueToModify", .ResidueToModify)
                    .ResidueMassModification = GetIniFileSettingDbl(IniStuff, strSectionName, "ResidueMassModification", .ResidueMassModification)
                End With
            End With
            
            ' Load the remaining Search Modes, but check if missing
            For intAutoSearchModeIndex = 1 To .AutoAnalysisSearchModeCount - 1
                strSectionName = "AutoAnalysisSearchMode" & Trim(intAutoSearchModeIndex + 1)
                .AutoAnalysisSearchMode(intAutoSearchModeIndex).SearchMode = GetIniFileSetting(IniStuff, strSectionName, "SearchMode", ENTRY_NOT_FOUND)
                If .AutoAnalysisSearchMode(intAutoSearchModeIndex).SearchMode = ENTRY_NOT_FOUND Then
                    .AutoAnalysisSearchModeCount = intAutoSearchModeIndex
                    Exit For
                Else
                    With .AutoAnalysisSearchMode(intAutoSearchModeIndex)
                        .AlternateOutputFolderPath = GetIniFileSetting(IniStuff, strSectionName, "AlternateOutputFolderPath", "")
                        .WriteResultsToTextFile = GetIniFileSettingBln(IniStuff, strSectionName, "WriteResultsToTextFile", False)
                        .ExportResultsToDatabase = GetIniFileSettingBln(IniStuff, strSectionName, "ExportResultsToDatabase", False)
                        .ExportUMCMembers = GetIniFileSettingBln(IniStuff, strSectionName, "ExportUMCMembers", False)
                        .PairSearchAssumeMassTagsAreLabeled = GetIniFileSettingBln(IniStuff, strSectionName, "PairSearchAssumeMassTagsAreLabeled", False)
                        
                
                        ' Parameter GANETLockerSearchMode was renamed to InternalStdSearchMode; allowing for both names here, with InternalStdSearchMode taking precedence
                        .InternalStdSearchMode = GetIniFileSettingInt(IniStuff, strSectionName, "GANETLockerSearchMode", issmFindWithMassTags)
                        .InternalStdSearchMode = GetIniFileSettingInt(IniStuff, strSectionName, "InternalStdSearchMode", CInt(.InternalStdSearchMode))
                        
                        .DBSearchMinimumHighNormalizedScore = GetIniFileSettingSng(IniStuff, strSectionName, "DBSearchMinimumHighNormalizedScore", 0)
                        .DBSearchMinimumHighDiscriminantScore = GetIniFileSettingSng(IniStuff, strSectionName, "DBSearchMinimumHighDiscriminantScore", 0)
                        .DBSearchMinimumPeptideProphetProbability = GetIniFileSettingSng(IniStuff, strSectionName, "DBSearchMinimumPeptideProphetProbability", 0)
                                        
                        With .MassMods
                            .DynamicMods = GetIniFileSettingBln(IniStuff, strSectionName, "DynamicMods", True)
                            .N15InsteadOfN14 = GetIniFileSettingBln(IniStuff, strSectionName, "N15InsteadOfN14", False)
                            .PEO = GetIniFileSettingBln(IniStuff, strSectionName, "PEO", False)
                            .ICATd0 = GetIniFileSettingBln(IniStuff, strSectionName, "ICATd0", False)
                            .ICATd8 = GetIniFileSettingBln(IniStuff, strSectionName, "ICATd8", False)
                            .Alkylation = GetIniFileSettingBln(IniStuff, strSectionName, "Alkylation", False)
                            .AlkylationMass = GetIniFileSettingDbl(IniStuff, strSectionName, "CustomMass", glALKYLATION)
                            .ResidueToModify = GetIniFileSetting(IniStuff, strSectionName, "ResidueToModify", "")
                            .ResidueMassModification = GetIniFileSettingDbl(IniStuff, strSectionName, "ResidueMassModification", 0)
                        End With
                    End With
                End If
            Next intAutoSearchModeIndex
            
            ' Blank out the remaining entries
            For intAutoSearchModeIndex = .AutoAnalysisSearchModeCount To MAX_AUTO_SEARCH_MODE_COUNT - 1
                ResetAutoSearchModeEntry .AutoAnalysisSearchMode(intAutoSearchModeIndex)
            Next intAutoSearchModeIndex
            
        End With
        
        With .AutoAnalysisFilterPrefs
            .ExcludeDuplicates = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeDuplicates", .ExcludeDuplicates)
            .ExcludeDuplicatesTolerance = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeDuplicatesTolerance", .ExcludeDuplicatesTolerance)
            
            .ExcludeIsoByFit = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeIsoByFit", .ExcludeIsoByFit)
            .ExcludeIsoByFitMaxVal = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeIsoByFitMaxVal", .ExcludeIsoByFitMaxVal)
            
            .ExcludeIsoSecondGuess = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeIsoSecondGuess", .ExcludeIsoSecondGuess)
            .ExcludeIsoLessLikelyGuess = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeIsoLessLikelyGuess", .ExcludeIsoLessLikelyGuess)
            
            .ExcludeCSByStdDev = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeCSByStdDev", .ExcludeCSByStdDev)
            .ExcludeCSByStdDevMaxVal = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "ExcludeCSByStdDevMaxVal", .ExcludeCSByStdDevMaxVal)
            
            .RestrictIsoByAbundance = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoByAbundance", .RestrictIsoByAbundance)
            .RestrictIsoAbundanceMin = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoAbundanceMin", .RestrictIsoAbundanceMin)
            .RestrictIsoAbundanceMax = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoAbundanceMax", .RestrictIsoAbundanceMax)
            
            .RestrictIsoByMass = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoByMass", .RestrictIsoByMass)
            .RestrictIsoMassMin = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoMassMin", .RestrictIsoMassMin)
            .RestrictIsoMassMax = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoMassMax", .RestrictIsoMassMax)
            
            .RestrictIsoByMZ = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoByMZ", .RestrictIsoByMZ)
            .RestrictIsoMZMin = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoMZMin", .RestrictIsoMZMin)
            .RestrictIsoMZMax = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoMZMax", .RestrictIsoMZMax)
            
            .RestrictIsoByChargeState = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoByChargeState", .RestrictIsoByChargeState)
            .RestrictIsoChargeStateMin = GetIniFileSettingInt(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoChargeStateMin", .RestrictIsoChargeStateMin)
            .RestrictIsoChargeStateMax = GetIniFileSettingInt(IniStuff, "AutoAnalysisFilterPrefs", "RestrictIsoChargeStateMax", .RestrictIsoChargeStateMax)
            
            .RestrictCSByAbundance = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictCSByAbundance", .RestrictCSByAbundance)
            .RestrictCSAbundanceMin = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictCSAbundanceMin", .RestrictCSAbundanceMin)
            .RestrictCSAbundanceMax = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictCSAbundanceMax", .RestrictCSAbundanceMax)
            
            .RestrictCSByMass = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictCSByMass", .RestrictCSByMass)
            .RestrictCSMassMin = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictCSMassMin", .RestrictCSMassMin)
            .RestrictCSMassMax = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictCSMassMax", .RestrictCSMassMax)
            
            .RestrictScanRange = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictScanRange", .RestrictScanRange)
            .RestrictScanRangeMin = GetIniFileSettingLng(IniStuff, "AutoAnalysisFilterPrefs", "RestrictScanRangeMin", .RestrictScanRangeMin)
            .RestrictScanRangeMax = GetIniFileSettingLng(IniStuff, "AutoAnalysisFilterPrefs", "RestrictScanRangeMax", .RestrictScanRangeMax)
            
            .RestrictGANETRange = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictGANETRange", .RestrictGANETRange)
            .RestrictGANETRangeMin = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictGANETRangeMin", .RestrictGANETRangeMin)
            .RestrictGANETRangeMax = GetIniFileSettingDbl(IniStuff, "AutoAnalysisFilterPrefs", "RestrictGANETRangeMax", .RestrictGANETRangeMax)
            
            .RestrictToEvenScanNumbersOnly = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictToEvenScanNumbersOnly", .RestrictToEvenScanNumbersOnly)
            .RestrictToOddScanNumbersOnly = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "RestrictToOddScanNumbersOnly", .RestrictToOddScanNumbersOnly)
                        
            ' Do not allow both of the above to be true
            ' If both are true, set both to false
            If .RestrictToEvenScanNumbersOnly And .RestrictToOddScanNumbersOnly Then
                .RestrictToEvenScanNumbersOnly = False
                .RestrictToOddScanNumbersOnly = False
            End If
            
            .MaximumDataCountEnabled = GetIniFileSettingBln(IniStuff, "AutoAnalysisFilterPrefs", "MaximumDataCountEnabled", .MaximumDataCountEnabled)
            .MaximumDataCountToLoad = GetIniFileSettingLng(IniStuff, "AutoAnalysisFilterPrefs", "MaximumDataCountToLoad", .MaximumDataCountToLoad)
        End With
        
        ' Now attempt to load the database connection info
        .AutoAnalysisDBInfoIsValid = IniFileReadSingleDBConnection(IniStuff, "AutoAnalysisDBInfo", udtDBSettingsSingle)
        
        If udtDBSettingsSingle.IsDeleted Then .AutoAnalysisDBInfoIsValid = False
        
        If .AutoAnalysisDBInfoIsValid Then
            .AutoAnalysisDBInfo = udtDBSettingsSingle.AnalysisInfo
        End If
    End With
    
    If Not bnlAutoAnalysisFieldsOnly Then
        ' Overlay options
        With OlyOptions
            .DefType = GetIniFileSettingLng(IniStuff, "OlyOptions", "DefType", .DefType)
            .DefShape = GetIniFileSettingLng(IniStuff, "OlyOptions", "DefShape", .DefShape)
            .DefColor = GetIniFileSettingLng(IniStuff, "OlyOptions", "DefColor", .DefColor)
            .DefVisible = GetIniFileSettingBln(IniStuff, "OlyOptions", "DefVisible", .DefVisible)
            .DefMinSize = GetIniFileSettingSng(IniStuff, "OlyOptions", "DefMinSize", .DefMinSize)
            .DefMaxSize = GetIniFileSettingSng(IniStuff, "OlyOptions", "DefMaxSize", .DefMaxSize)
            .DefFontWidth = GetIniFileSettingSng(IniStuff, "OlyOptions", "DefFontWidth", .DefFontWidth)
            .DefFontHeight = GetIniFileSettingSng(IniStuff, "OlyOptions", "DefFontHeight", .DefFontHeight)
            .DefTextHeight = GetIniFileSettingSng(IniStuff, "OlyOptions", "DefTextHeight", .DefTextHeight)
            
            .DefStickWidth = GetIniFileSettingDbl(IniStuff, "OlyOptions", "DefStickWidth", .DefStickWidth)
            .DefMinNET = GetIniFileSettingDbl(IniStuff, "OlyOptions", "DefMinNET", .DefMinNET)
            .DefMaxNET = GetIniFileSettingDbl(IniStuff, "OlyOptions", "DefMaxNET", .DefMaxNET)
            .DefNETAdjustment = GetIniFileSettingLng(IniStuff, "OlyOptions", "DefNETAdjustment", .DefNETAdjustment)
            .DefNETTol = GetIniFileSettingDbl(IniStuff, "OlyOptions", "DefNETTol", .DefNETTol)
            .DefUniformSize = GetIniFileSettingBln(IniStuff, "OlyOptions", "DefUniformSize", .DefUniformSize)
            .DefBoxSizeAsSpotSize = GetIniFileSettingBln(IniStuff, "OlyOptions", "DefBoxSizeAsSpotSize", .DefBoxSizeAsSpotSize)
            .DefWithID = GetIniFileSettingBln(IniStuff, "OlyOptions", "DefWithID", .DefWithID)
            .DefCurrScopeVisible = GetIniFileSettingBln(IniStuff, "OlyOptions", "DefCurrScopeVisible", .DefCurrScopeVisible)
            
            .BackColor = GetIniFileSettingLng(IniStuff, "OlyOptions", "BackColor", .BackColor)
            .ForeColor = GetIniFileSettingLng(IniStuff, "OlyOptions", "ForeColor", .ForeColor)
            .Orientation = GetIniFileSettingLng(IniStuff, "OlyOptions", "Orientation", .Orientation)
            
            If Not .GRID Is Nothing Then
                With .GRID
                    .LineStyle = GetIniFileSettingLng(IniStuff, "OlyGridOptions", "LineStyle", .LineStyle)
                    .HorzAutoMode = GetIniFileSettingLng(IniStuff, "OlyGridOptions", "HorzAutoMode", .HorzAutoMode)
                    .HorzBinsCount = GetIniFileSettingLng(IniStuff, "OlyGridOptions", "HorzBinsCount", .HorzBinsCount)
                    .HorzGridVisible = GetIniFileSettingBln(IniStuff, "OlyGridOptions", "HorzGridVisible", .HorzGridVisible)
                    .VertAutoMode = GetIniFileSettingLng(IniStuff, "OlyGridOptions", "VertAutoMode", .VertAutoMode)
                    .VertBinsCount = GetIniFileSettingLng(IniStuff, "OlyGridOptions", "VertBinsCount", .VertBinsCount)
                    .VertGridVisible = GetIniFileSettingBln(IniStuff, "OlyGridOptions", "VertGridVisible", .VertGridVisible)
                End With
            End If
        End With
        
        With OlyJiggyOptions
            .UseMWConstraint = GetIniFileSettingBln(IniStuff, "OlyJiggyOptions", "UseMWConstraint", .UseMWConstraint)
            .MWTol = GetIniFileSettingDbl(IniStuff, "OlyJiggyOptions", "MWTol", .MWTol)
            .UseNetConstraint = GetIniFileSettingBln(IniStuff, "OlyJiggyOptions", "UseNetConstraint", .UseNetConstraint)
            .NETTol = GetIniFileSettingDbl(IniStuff, "OlyJiggyOptions", "NETTol", .NETTol)
            .UseAbuConstraint = GetIniFileSettingBln(IniStuff, "OlyJiggyOptions", "UseAbuConstraint", .UseAbuConstraint)
            .AbuTol = GetIniFileSettingDbl(IniStuff, "OlyJiggyOptions", "AbuTol", .AbuTol)
            .JiggyScope = GetIniFileSettingLng(IniStuff, "OlyJiggyOptions", "JiggyScope", .JiggyScope)
            .JiggyType = GetIniFileSettingLng(IniStuff, "OlyJiggyOptions", "JiggyType", .JiggyType)
            .BaseDisplayInd = GetIniFileSettingLng(IniStuff, "OlyJiggyOptions", "BaseDisplayInd", .BaseDisplayInd)
        End With
    End If
    
    Set IniStuff = Nothing
    Exit Sub

LoadSettingsFileHandler:
    If Not udtPrefsExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error loading data from the Ini file (" & IniStuff.FileName & "); Sub IniFileLoadSettings in Settings.Bas" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Print "Error in IniFileLoadSettings: " & Err.Description
        Debug.Assert False
        LogErrors Err.Number, "Settings.Bas->IniFileLoadSettings"
    End If
    Set IniStuff = Nothing

End Sub

Public Sub IniFileSaveSettings(udtPrefsExpanded As udtPreferencesExpandedType, udtUMCDef As UMCDefinition, udtUMCIonNetDef As UMCIonNetDefinition, udtUMCNetAdjDef As NetAdjDefinition, udtInternalStandards As udtInternalStandardsType, udtAMTDef As SearchAMTDefinition, udtPrefs As GelPrefs, Optional strIniFilePath As String = "", Optional bnlAutoAnalysisFieldsOnly As Boolean = False)
    ' Saves settings to an .ini file
    ' When bnlAutoAnalysisFieldsOnly = True, then skips the settings that are not needed for Auto Analysis
    
    Dim blnSuccess As Boolean
    Dim IniStuff As New clsIniStuff
    Dim DBIniStuff As clsIniStuff
    Dim strDBIniFilePath As String
    Dim intIndex As Integer, intAutoSearchModeIndex As Integer
    Dim intTargetIndexBase As Integer
    Dim strKeys() As String, strValues() As String
    Dim strKeyPrefix As String
    Dim udtDBSettingsSingle As udtDBSettingsType
    Dim strMassTagSubsetID As String
    Dim strSectionName As String
    
On Error GoTo SaveSettingsFileHandler

    ' This Sub shouldn't need a progress bar, but without it, it takes 4 seconds to execute
    ' With the progress bar, it takes 0.5 seconds. Go figure.
    frmProgress.InitializeForm "Saving Program Settings", 0, 3, False, False, False
    
    ' Set the Ini filename
    If Len(strIniFilePath) > 0 Then
        IniStuff.FileName = strIniFilePath
    Else
        IniStuff.FileName = GetIniFilePath()
    End If
        
    ' Database Settings
    blnSuccess = IniStuff.WriteValue("Paths", "ICR2LS", sICR2LSCommand)
    
    ' If an error occurs on the first call to iniStuff.WriteValue(), then abort
    ' I don't check for errors again after this
    If Not blnSuccess Then GoTo SaveSettingsFileHandler
    
    ReDim strKeys(0 To 42)
    ReDim strValues(0 To 42)
        
    ' UMC options stored in udtPrefsExpanded.AutoAnalysisOptions
    With udtPrefsExpanded.AutoAnalysisOptions
        strKeys(0) = "UMCSearchModeList": strValues(0) = "; Options are " & GetUMCSearchModeList()
        strKeys(1) = "UMCSearchMode": strValues(1) = .UMCSearchMode
        strKeys(2) = "UMCShrinkingBoxWeightAverageMassByIntensity": strValues(2) = .UMCShrinkingBoxWeightAverageMassByIntensity
    End With
    
    ' UMC options stored in udtUMCDef
    With udtUMCDef
        strKeys(3) = "UMCTypeList": strValues(3) = "; Options are " & GetUMCTypeList()
        strKeys(4) = "UMCType": strValues(4) = .UMCType
        strKeys(5) = "MWField": strValues(5) = .MWField
        strKeys(6) = "TolType": strValues(6) = .TolType
        strKeys(7) = "Tol": strValues(7) = .Tol
        strKeys(8) = "UMCSharing": strValues(8) = .UMCSharing
        strKeys(9) = "UMCUniCS": strValues(9) = .UMCUniCS
        strKeys(10) = "ClassAbu": strValues(10) = .ClassAbu
        strKeys(11) = "ClassMW": strValues(11) = .ClassMW
        strKeys(12) = "GapMaxCnt": strValues(12) = .GapMaxCnt
        strKeys(13) = "GapMaxSize": strValues(13) = .GapMaxSize
        strKeys(14) = "GapMaxPct": strValues(14) = .GapMaxPct
        strKeys(15) = "UMCNETType": strValues(15) = .UMCNETType
        strKeys(16) = "InterpolateGaps": strValues(16) = .InterpolateGaps
        strKeys(17) = "InterpolateMaxGapSize": strValues(17) = .InterpolateMaxGapSize
        strKeys(18) = "InterpolationType": strValues(18) = .InterpolationType
        strKeys(19) = "ChargeStateStatsRepType": strValues(19) = .ChargeStateStatsRepType
        strKeys(20) = "UMCClassStatsUseStatsFromMostAbuChargeState": strValues(20) = .UMCClassStatsUseStatsFromMostAbuChargeState
    End With
    
    ' UMC options stored in udtPrefsExpanded
    With udtPrefsExpanded.UMCAutoRefineOptions
        strKeys(21) = "UMCAutoRefineRemoveCountLow": strValues(21) = .UMCAutoRefineRemoveCountLow
        strKeys(22) = "UMCAutoRefineRemoveCountHigh": strValues(22) = .UMCAutoRefineRemoveCountHigh
        strKeys(23) = "UMCAutoRefineRemoveMaxLengthPctAllScans": strValues(23) = .UMCAutoRefineRemoveMaxLengthPctAllScans
        
        strKeys(24) = "UMCAutoRefineMinLength": strValues(24) = .UMCAutoRefineMinLength
        strKeys(25) = "UMCAutoRefineMaxLength": strValues(25) = .UMCAutoRefineMaxLength
        strKeys(26) = "UMCAutoRefineMaxLengthPctAllScans": strValues(26) = .UMCAutoRefineMaxLengthPctAllScans
        
        strKeys(27) = "UMCAutoRefinePercentMaxAbuToUseForLength": strValues(27) = .UMCAutoRefinePercentMaxAbuToUseForLength
        strKeys(28) = "TestLengthUsingScanRange": strValues(28) = .TestLengthUsingScanRange
        strKeys(29) = "MinMemberCountWhenUsingScanRange": strValues(29) = .MinMemberCountWhenUsingScanRange
        strKeys(30) = "UMCAutoRefineRemoveAbundanceLow": strValues(30) = .UMCAutoRefineRemoveAbundanceLow
        strKeys(31) = "UMCAutoRefineRemoveAbundanceHigh": strValues(31) = .UMCAutoRefineRemoveAbundanceHigh
        strKeys(32) = "UMCAutoRefinePctLowAbundance": strValues(32) = .UMCAutoRefinePctLowAbundance
        strKeys(33) = "UMCAutoRefinePctHighAbundance": strValues(33) = .UMCAutoRefinePctHighAbundance
        strKeys(34) = "SplitUMCsByAbundance": strValues(34) = .SplitUMCsByAbundance
        With .SplitUMCOptions
            strKeys(35) = "MinimumDifferenceInAveragePpmMassToSplit": strValues(35) = .MinimumDifferenceInAveragePpmMassToSplit
            strKeys(36) = "StdDevMultiplierForSplitting": strValues(36) = .StdDevMultiplierForSplitting
            strKeys(37) = "MaximumPeakCountToSplitUMC": strValues(37) = .MaximumPeakCountToSplitUMC
            strKeys(38) = "PeakDetectIntensityThresholdPercentageOfMaximum": strValues(38) = .PeakDetectIntensityThresholdPercentageOfMaximum
            strKeys(39) = "PeakDetectIntensityThresholdAbsoluteMinimum": strValues(39) = .PeakDetectIntensityThresholdAbsoluteMinimum
            strKeys(40) = "PeakWidthPointsMinimum": strValues(40) = .PeakWidthPointsMinimum
            strKeys(41) = "PeakWidthInSigma": strValues(41) = .PeakWidthInSigma
            strKeys(42) = "ScanGapBehavior": strValues(42) = .ScanGapBehavior
        End With
    End With
    IniStuff.WriteSection "UMCDef", strKeys(), strValues()
    
    ' Reserve 0 to 100 to give lots of extra space
    ' Will redim below before calling IniStuff.WriteSection
    ReDim strKeys(0 To 100)
    ReDim strValues(0 To 100)
    With udtUMCIonNetDef
        strKeys(0) = "NetDim": strValues(0) = .NetDim
        strKeys(1) = "NetActualDim": strValues(1) = .NetActualDim
        strKeys(2) = "MetricType": strValues(2) = .MetricType
        strKeys(3) = "NETType": strValues(3) = .NETType
        strKeys(4) = "TooDistant": strValues(4) = .TooDistant
        For intIndex = 0 To UBound(.MetricData())
            intTargetIndexBase = 5 + intIndex * 6
            strKeyPrefix = "Dim" & Trim(intIndex + 1)
            With .MetricData(intIndex)
                strKeys(intTargetIndexBase) = strKeyPrefix & "Use": strValues(intTargetIndexBase) = .Use
                strKeys(intTargetIndexBase + 1) = strKeyPrefix & "DataType": strValues(intTargetIndexBase + 1) = .DataType
                strKeys(intTargetIndexBase + 2) = strKeyPrefix & "WeightFactor": strValues(intTargetIndexBase + 2) = .WeightFactor
                strKeys(intTargetIndexBase + 3) = strKeyPrefix & "ConstraintType": strValues(intTargetIndexBase + 3) = .ConstraintType
                strKeys(intTargetIndexBase + 4) = strKeyPrefix & "ConstraintValue": strValues(intTargetIndexBase + 4) = .ConstraintValue
                strKeys(intTargetIndexBase + 5) = strKeyPrefix & "ConstraintUnits": strValues(intTargetIndexBase + 5) = .ConstraintUnits
            End With
        Next intIndex
    
        Debug.Assert intTargetIndexBase + 5 <= 100
    End With
    
    ' UMCIso options stored in udtPrefsExpanded
    ' If .MetricData() is changed to not have 6 items, then the following + # values must be changed
    With udtPrefsExpanded.UMCIonNetOptions
        strKeys(intTargetIndexBase + 6) = "UMCRepresentative": strValues(intTargetIndexBase + 6) = .UMCRepresentative
        strKeys(intTargetIndexBase + 7) = "MakeSingleMemberClasses": strValues(intTargetIndexBase + 7) = .MakeSingleMemberClasses
        strKeys(intTargetIndexBase + 8) = "ConnectionLengthPostFilterMaxNET": strValues(intTargetIndexBase + 8) = .ConnectionLengthPostFilterMaxNET
    End With
    
    ' The following ReDim statments should match the last item saved in strKeys() above
    ReDim Preserve strKeys(0 To intTargetIndexBase + 8)
    ReDim Preserve strValues(0 To intTargetIndexBase + 8)
    IniStuff.WriteSection "UMCIonNetDef", strKeys(), strValues()
    
    ReDim strKeys(0 To 5)
    ReDim strValues(0 To 5)
    With udtPrefsExpanded.UMCAdvancedStatsOptions
        strKeys(0) = "ClassAbuTopXMinAbu": strValues(0) = .ClassAbuTopXMinAbu
        strKeys(1) = "ClassAbuTopXMaxAbu": strValues(1) = .ClassAbuTopXMaxAbu
        strKeys(2) = "ClassAbuTopXMinMembers": strValues(2) = .ClassAbuTopXMinMembers

        strKeys(3) = "ClassMassTopXMinAbu": strValues(3) = .ClassMassTopXMinAbu
        strKeys(4) = "ClassMassTopXMaxAbu": strValues(4) = .ClassMassTopXMaxAbu
        strKeys(5) = "ClassMassTopXMinMembers": strValues(5) = .ClassMassTopXMinMembers
    End With
    IniStuff.WriteSection "UMCAdvancedStatsOptions", strKeys(), strValues()
    
    
    ReDim strKeys(0 To 46)
    ReDim strValues(0 To 46)
    With udtUMCNetAdjDef
        strKeys(0) = "MinUMCCount": strValues(0) = .MinUMCCount
        strKeys(1) = "MinScanRange": strValues(1) = .MinScanRange
        strKeys(2) = "MaxScanPct": strValues(2) = .MaxScanPct
        strKeys(3) = "TopAbuPct": strValues(3) = .TopAbuPct
        ' Ignored: .PeakSelection
        ' Ignored: .PeakMaxAbuPct
        strKeys(4) = "MWTolType": strValues(4) = .MWTolType
        strKeys(5) = "MWTol": strValues(5) = .MWTol
        strKeys(6) = "NETorRT": strValues(6) = .NETorRT
        strKeys(7) = "UseNET": strValues(7) = .UseNET
        strKeys(8) = "UseMultiIDMaxNETDist": strValues(8) = .UseMultiIDMaxNETDist
        strKeys(9) = "MultiIDMaxNETDist": strValues(9) = .MultiIDMaxNETDist
        strKeys(10) = "EliminateBadNET": strValues(10) = .EliminateBadNET
        strKeys(11) = "MaxIDToUse": strValues(11) = .MaxIDToUse
        strKeys(12) = "IterationStopType": strValues(12) = .IterationStopType
        strKeys(13) = "IterationStopValue": strValues(13) = .IterationStopValue
        strKeys(14) = "IterationUseMWDec": strValues(14) = .IterationUseMWDec
        strKeys(15) = "IterationMWDec": strValues(15) = .IterationMWDec
        strKeys(16) = "IterationUseNETdec": strValues(16) = .IterationUseNETdec
        strKeys(17) = "IterationNETDec": strValues(17) = .IterationNETDec
        strKeys(18) = "IterationAcceptLast": strValues(18) = .IterationAcceptLast
        strKeys(19) = "InitialSlope": strValues(19) = .InitialSlope
        strKeys(20) = "InitialIntercept": strValues(20) = .InitialIntercept
        
        ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''        strKeys(21) = "UseNetAdjLockers": strValues(21) = .UseNetAdjLockers
''        strKeys(22) = "UseOldNetAdjIfFailure": strValues(22) = .UseOldNetAdjIfFailure
''        strKeys(23) = "NetAdjLockerMinimumMatchCount": strValues(23) = .NetAdjLockerMinimumMatchCount
        
        strKeys(21) = "UseRobustNETAdjustment": strValues(21) = .UseRobustNETAdjustment
        strKeys(22) = "RobustNETAdjustmentMode": strValues(22) = .RobustNETAdjustmentMode
        strKeys(23) = "RobustNETSlopeStart": strValues(23) = .RobustNETSlopeStart
        strKeys(24) = "RobustNETSlopeEnd": strValues(24) = .RobustNETSlopeEnd
        strKeys(25) = "RobustNETSlopeIncreaseMode": strValues(25) = .RobustNETSlopeIncreaseMode
        strKeys(26) = "RobustNETSlopeIncrement": strValues(26) = .RobustNETSlopeIncrement
        strKeys(27) = "RobustNETInterceptStart": strValues(27) = .RobustNETInterceptStart
        strKeys(28) = "RobustNETInterceptEnd": strValues(28) = .RobustNETInterceptEnd
        strKeys(29) = "RobustNETInterceptIncrement": strValues(29) = .RobustNETInterceptIncrement
        strKeys(30) = "RobustNETMassShiftPPMStart": strValues(30) = .RobustNETMassShiftPPMStart
        strKeys(31) = "RobustNETMassShiftPPMEnd": strValues(31) = .RobustNETMassShiftPPMEnd
        strKeys(32) = "RobustNETMassShiftPPMIncrement": strValues(32) = .RobustNETMassShiftPPMIncrement
    
    End With
    
    With udtPrefsExpanded
        With .AutoAnalysisOptions
            strKeys(33) = "NETAdjustmentInitialNetTol": strValues(33) = .NETAdjustmentInitialNetTol
            strKeys(34) = "NETAdjustmentMaxIterationCount": strValues(34) = .NETAdjustmentMaxIterationCount
            strKeys(35) = "NETAdjustmentMinIDCount": strValues(35) = .NETAdjustmentMinIDCount
            strKeys(36) = "NETAdjustmentMinIDCountAbsoluteMinimum": strValues(36) = .NETAdjustmentMinIDCountAbsoluteMinimum
            strKeys(37) = "NETAdjustmentMinIterationCount": strValues(37) = .NETAdjustmentMinIterationCount
            strKeys(38) = "NETAdjustmentChangeThresholdStopValue": strValues(38) = .NETAdjustmentChangeThresholdStopValue
            
            strKeys(39) = "NETAdjustmentAutoIncrementUMCTopAbuPct": strValues(39) = .NETAdjustmentAutoIncrementUMCTopAbuPct
            strKeys(40) = "NETAdjustmentUMCTopAbuPctInitial": strValues(40) = .NETAdjustmentUMCTopAbuPctInitial
            strKeys(41) = "NETAdjustmentUMCTopAbuPctIncrement": strValues(41) = .NETAdjustmentUMCTopAbuPctIncrement
            strKeys(42) = "NETAdjustmentUMCTopAbuPctMax": strValues(42) = .NETAdjustmentUMCTopAbuPctMax
            
            strKeys(43) = "NETSlopeExpectedMinimum": strValues(43) = .NETSlopeExpectedMinimum
            strKeys(44) = "NETSlopeExpectedMaximum": strValues(44) = .NETSlopeExpectedMaximum
            strKeys(45) = "NETInterceptExpectedMinimum": strValues(45) = .NETInterceptExpectedMinimum
            strKeys(46) = "NETInterceptExpectedMaximum": strValues(46) = .NETInterceptExpectedMaximum
        End With
    End With
    IniStuff.WriteSection NET_ADJ_SECTION_NEWNAME, strKeys(), strValues()

    ' Write this after writing the udtUMCNETAdjDef section
    With udtUMCNetAdjDef
        For intIndex = 0 To UBound(.PeakCSSelection)
            IniStuff.WriteValue NET_ADJ_SECTION_NEWNAME, "PeakCSSelection" & Trim(intIndex), CStr(.PeakCSSelection(intIndex))
        Next intIndex
    End With

    If Not APP_BUILD_DISABLE_LCMSWARP Then
        ReDim strKeys(0 To 15)
        ReDim strValues(0 To 15)
        With udtUMCNetAdjDef.MSWarpOptions
          
            strKeys(0) = "MassCalibrationType": strValues(0) = .MassCalibrationType
            strKeys(1) = "MinimumPMTTagObsCount": strValues(1) = .MinimumPMTTagObsCount
            strKeys(2) = "MatchPromiscuity": strValues(2) = .MatchPromiscuity
            
            strKeys(3) = "NETTol": strValues(3) = .NETTol
            strKeys(4) = "NumberOfSections": strValues(4) = .NumberOfSections
            strKeys(5) = "MaxDistortion": strValues(5) = .MaxDistortion
            strKeys(6) = "ContractionFactor": strValues(6) = .ContractionFactor
            
            strKeys(7) = "MassWindowPPM": strValues(7) = .MassWindowPPM
            strKeys(8) = "MassSplineOrder": strValues(8) = .MassSplineOrder
            strKeys(9) = "MassNumXSlices": strValues(9) = .MassNumXSlices
            strKeys(10) = "MassNumMassDeltaBins": strValues(10) = .MassNumMassDeltaBins
            strKeys(11) = "MassMaxJump": strValues(11) = .MassMaxJump
            
            strKeys(12) = "MassZScoreTolerance": strValues(12) = .MassZScoreTolerance
            strKeys(13) = "MassUseLSQ": strValues(13) = .MassUseLSQ
            strKeys(14) = "MassLSQOutlierZScore": strValues(14) = .MassLSQOutlierZScore
            strKeys(15) = "MassLSQNumKnots": strValues(15) = .MassLSQNumKnots
        End With
        IniStuff.WriteSection NET_ADJ_MS_WARP_SECTION, strKeys(), strValues()
    End If
        
        
'' Note: Uncomment the following to enable writing of the internal standards to a .Ini file
''    Dim intInternalStandardIndex As Integer
''    ReDim strKeys(0 To 0)
''    ReDim strValues(0 To 0)
''
''    ' Write the Internal Standards
''    With udtInternalStandards
''        strKeys(0) = "Count": strValues(0) = .Count
''    End With
''    IniStuff.WriteSection "UMCInternalStandards", strKeys(), strValues()
''
''    ' Write the Internal Standards
''    ' Each locker is written to its own section in the .Ini file
''    With udtInternalStandards
''        For intInternalStandardIndex = 0 To .Count - 1
''
''            With .InternalStandards(intInternalStandardIndex)
''
''                ' Write this Internal Standard
''                ReDim strKeys(0 To 6)
''                ReDim strValues(0 To 6)
''
''                strKeys(0) = "SeqID": strValues(0) = .SeqID
''                strKeys(1) = "PeptideSequence": strValues(1) = .PeptideSequence
''                strKeys(2) = "MonoisotopicMass": strValues(2) = .MonoisotopicMass
''                strKeys(3) = "NET": strValues(3) = .NET
''                strKeys(4) = "ChargeMinimum": strValues(4) = .ChargeMinimum
''                strKeys(5) = "ChargeMaximum": strValues(5) = .ChargeMaximum
''                strKeys(6) = "ChargeMostAbundant": strValues(6) = .ChargeMostAbundant
''
''                IniStuff.WriteSection "UMCInternalStandards" & Trim(intInternalStandardIndex + 1), strKeys(), strValues()
''            End With
''
''        Next intInternalStandardIndex
''    End With


    ReDim strKeys(0 To 9)
    ReDim strValues(0 To 9)
    With udtAMTDef
        strKeys(0) = "SearchFlag": strValues(0) = .SearchFlag
        strKeys(1) = "MWField": strValues(1) = .MWField
        strKeys(2) = "MWTol": strValues(2) = .MWTol
        strKeys(3) = "NETorRT": strValues(3) = .NETorRT
        strKeys(4) = "TolType": strValues(4) = .TolType
        strKeys(5) = "NETTol": strValues(5) = .NETTol
        strKeys(6) = "MassTag": strValues(6) = .MassTag
        strKeys(7) = "MaxMassTags": strValues(7) = .MaxMassTags
        strKeys(8) = "SkipReferenced": strValues(8) = .SkipReferenced
        strKeys(9) = "SaveNCnt": strValues(9) = .SaveNCnt
    End With
    IniStuff.WriteSection "SearchAMTDef", strKeys(), strValues()
   
    If Not bnlAutoAnalysisFieldsOnly Then
        ReDim strKeys(0 To 20)
        ReDim strValues(0 To 20)
        With OlyOptions
            strKeys(0) = "DefType": strValues(0) = .DefType
            strKeys(1) = "DefShape": strValues(1) = .DefShape
            strKeys(2) = "DefColor": strValues(2) = .DefColor
            strKeys(3) = "DefVisible": strValues(3) = .DefVisible
            strKeys(4) = "DefMinSize": strValues(4) = .DefMinSize
            strKeys(5) = "DefMaxSize": strValues(5) = .DefMaxSize
            strKeys(6) = "DefFontWidth": strValues(6) = .DefFontWidth
            strKeys(7) = "DefFontHeight": strValues(7) = .DefFontHeight
            strKeys(8) = "DefTextHeight": strValues(8) = .DefTextHeight
            strKeys(9) = "DefStickWidth": strValues(9) = .DefStickWidth
            strKeys(10) = "DefMinNET": strValues(10) = .DefMinNET
            strKeys(11) = "DefMaxNET": strValues(11) = .DefMaxNET
            strKeys(12) = "DefNETAdjustment": strValues(12) = .DefNETAdjustment
            strKeys(13) = "DefNETTol": strValues(13) = .DefNETTol
            strKeys(14) = "DefUniformSize": strValues(14) = .DefUniformSize
            strKeys(15) = "DefBoxSizeAsSpotSize": strValues(15) = .DefBoxSizeAsSpotSize
            strKeys(16) = "DefWithID": strValues(16) = .DefWithID
            strKeys(17) = "DefCurrScopeVisible": strValues(17) = .DefCurrScopeVisible
            strKeys(18) = "BackColor": strValues(18) = .BackColor
            strKeys(19) = "ForeColor": strValues(19) = .ForeColor
            strKeys(20) = "Orientation": strValues(20) = .Orientation
        End With
        IniStuff.WriteSection "OlyOptions", strKeys(), strValues()
        
        ReDim strKeys(0 To 6)
        ReDim strValues(0 To 6)
        With OlyOptions
            If Not .GRID Is Nothing Then
                With .GRID
                    strKeys(0) = "LineStyle": strValues(0) = .LineStyle
                    strKeys(1) = "HorzAutoMode": strValues(1) = .HorzAutoMode
                    strKeys(2) = "HorzBinsCount": strValues(2) = .HorzBinsCount
                    strKeys(3) = "HorzGridVisible": strValues(3) = .HorzGridVisible
                    strKeys(4) = "VertAutoMode": strValues(4) = .VertAutoMode
                    strKeys(5) = "VertBinsCount": strValues(5) = .VertBinsCount
                    strKeys(6) = "VertGridVisible": strValues(6) = .VertGridVisible
                End With
            End If
        End With
        IniStuff.WriteSection "OlyGridOptions", strKeys(), strValues()
        
        ReDim strKeys(0 To 8)
        ReDim strValues(0 To 8)
        With OlyJiggyOptions
            strKeys(0) = "UseMWConstraint": strValues(0) = .UseMWConstraint
            strKeys(1) = "MWTol": strValues(1) = .MWTol
            strKeys(2) = "UseNetConstraint": strValues(2) = .UseNetConstraint
            strKeys(3) = "NETTol": strValues(3) = .NETTol
            strKeys(4) = "UseAbuConstraint": strValues(4) = .UseAbuConstraint
            strKeys(5) = "AbuTol": strValues(5) = .AbuTol
            strKeys(6) = "JiggyScope": strValues(6) = .JiggyScope
            strKeys(7) = "JiggyType": strValues(7) = .JiggyType
            strKeys(8) = "BaseDisplayInd": strValues(8) = .BaseDisplayInd
        End With
        IniStuff.WriteSection "OlyJiggyOptions", strKeys(), strValues()
    End If
    frmProgress.UpdateProgressBar 1
    
    ' Initialize the setting strings
    sCooSysPref = GetCooSysPrefs(udtPrefs)
    sDDClrPref = GetDDClrPrefs()
    sDrawingPref = GetDrawingPrefs(udtPrefs)
    sICR2LSPref = GetICR2LSPrefs()
    sBackForeCSIsoClrPref = GetOtherColorsPrefs()
    sCSIsoShapePref = GetCSIsoShapePrefs()
    sSwitchPref = GetSwitchPrefs(udtPrefs)
    sTolerancesPref = GetTolerancesPrefs(udtPrefs)
    
    ' No longer supported (March 2006)
    ''sAMTPref = GetAMTPrefs()
    ''sFTICR_AMTPref = GetFTICR_AMTPrefs()

    ReDim strKeys(0 To 7)
    ReDim strValues(0 To 7)
    strKeys(0) = "CoordinateSystem": strValues(0) = sCooSysPref
    strKeys(1) = "DifferentialDisplay": strValues(1) = sDDClrPref
    strKeys(2) = "Drawing": strValues(2) = sDrawingPref
    strKeys(3) = "ICR2LS": strValues(3) = sICR2LSPref
    strKeys(4) = "ChargeStateColors": strValues(4) = sBackForeCSIsoClrPref
    strKeys(5) = "ChargeStateShapes": strValues(5) = sCSIsoShapePref
    strKeys(6) = "Switches": strValues(6) = sSwitchPref
    strKeys(7) = "Tolerances": strValues(7) = sTolerancesPref
    
    ' No longer supported (March 2006)
    ''strKeys(8) = "AMTs": strValues(8) = sAMTPref
    ''strKeys(9) = "FTICRAmts": strValues(9) = sFTICR_AMTPref
            
    IniStuff.WriteSection "Preferences", strKeys(), strValues()

    ' Write the expanded preferences
    ReDim strKeys(0 To 17)
    ReDim strValues(0 To 17)
    With udtPrefsExpanded
        strKeys(0) = "MenuModeDefault": strValues(0) = .MenuModeDefault
        strKeys(1) = "MenuModeIncludeObsolete": strValues(1) = .MenuModeIncludeObsolete
        strKeys(2) = "ExtendedFileSaveModePreferred": strValues(2) = .ExtendedFileSaveModePreferred
        strKeys(3) = "AutoAdjSize": strValues(3) = .AutoAdjSize
        strKeys(4) = "AutoSizeMultiplier": strValues(4) = .AutoSizeMultiplier
        strKeys(5) = "UMCDrawType": strValues(5) = .UMCDrawType
        strKeys(6) = "UsePEKBasedERValues": strValues(6) = .UsePEKBasedERValues
        strKeys(7) = "UseMassTagsWithNullMass": strValues(7) = .UseMassTagsWithNullMass
        strKeys(8) = "UseMassTagsWithNullNET": strValues(8) = .UseMassTagsWithNullNET
        strKeys(9) = "UseUMCConglomerateNET": strValues(9) = .UseUMCConglomerateNET
        strKeys(10) = "NetAdjustmentUsesN15AMTMasses": strValues(10) = .NetAdjustmentUsesN15AMTMasses
        strKeys(11) = "NetAdjustmentMinHighNormalizedScore": strValues(11) = .NetAdjustmentMinHighNormalizedScore
        strKeys(12) = "NetAdjustmentMinHighDiscriminantScore": strValues(12) = .NetAdjustmentMinHighDiscriminantScore
        strKeys(13) = "AMTSearchResultsBehavior": strValues(13) = .AMTSearchResultsBehavior
        strKeys(14) = "ICR2LSSpectrumViewZoomWindowWidthMZ": strValues(14) = .ICR2LSSpectrumViewZoomWindowWidthMZ
        strKeys(15) = "LastAutoAnalysisIniFilePath": strValues(15) = .LastAutoAnalysisIniFilePath
        strKeys(16) = "LastInputFileMode": strValues(16) = .LastInputFileMode
        strKeys(17) = "LegacyAMTDBPath": strValues(17) = .LegacyAMTDBPath
    End With
    IniStuff.WriteSection "ExpandedPreferences", strKeys(), strValues()
    
    If Not bnlAutoAnalysisFieldsOnly And Not APP_BUILD_DISABLE_MTS Then
        ' Auto Query PRISM options
        ReDim strKeys(0 To 10)
        ReDim strValues(0 To 10)
        With udtPrefsExpanded.AutoQueryPRISMOptions
            strKeys(0) = "ConnectionStringQueryDB": strValues(0) = .ConnectionStringQueryDB
            strKeys(1) = "RequestTaskSPName": strValues(1) = .RequestTaskSPName
            strKeys(2) = "SetTaskCompleteSPName": strValues(2) = .SetTaskCompleteSPName
            strKeys(3) = "SetTaskToRestartSPName": strValues(3) = .SetTaskToRestartSPName
            strKeys(4) = "PostLogEntrySPName": strValues(4) = .PostLogEntrySPName
            strKeys(5) = "QueryIntervalSeconds": strValues(5) = .QueryIntervalSeconds
            strKeys(6) = "MinimumPriorityToProcess": strValues(6) = .MinimumPriorityToProcess
            strKeys(7) = "MaximumPriorityToProcess": strValues(7) = .MaximumPriorityToProcess
            strKeys(8) = "PreferredDatabaseToProcess": strValues(8) = .PreferredDatabaseToProcess
            strKeys(9) = "ServerForPreferredDatabase": strValues(9) = .ServerForPreferredDatabase
            strKeys(10) = "ExclusivelyUseThisDatabase": strValues(10) = .ExclusivelyUseThisDatabase
        End With
        IniStuff.WriteSection "AutoQueryPRISMOptions", strKeys(), strValues()
    End If
    
    ' Write the NET Adjustment UMC Selection Options
    ReDim strKeys(0 To 4)
    ReDim strValues(0 To 4)
    With udtPrefsExpanded.NetAdjustmentUMCDistributionOptions
        strKeys(0) = "RequireDispersedUMCSelection": strValues(0) = .RequireDispersedUMCSelection
        strKeys(1) = "SegmentCount": strValues(1) = .SegmentCount
        strKeys(2) = "MinimumUMCsPerSegmentPctTopAbuPct": strValues(2) = .MinimumUMCsPerSegmentPctTopAbuPct
        strKeys(3) = "ScanPctStart": strValues(3) = .ScanPctStart
        strKeys(4) = "ScanPctEnd": strValues(4) = .ScanPctEnd
    End With
    IniStuff.WriteSection "NetAdjustmentUMCDistributionOptions", strKeys(), strValues()
    
    ' Write the Error Plotting Options
    ReDim strKeys(0 To 4)
    ReDim strValues(0 To 4)
    With udtPrefsExpanded.ErrorPlottingOptions
        strKeys(0) = "MassRangePPM": strValues(0) = .MassRangePPM
        strKeys(1) = "MassBinSizePPM": strValues(1) = .MassBinSizePPM
        strKeys(2) = "GANETRange": strValues(2) = .GANETRange
        strKeys(3) = "GANETBinSize": strValues(3) = .GANETBinSize
        strKeys(4) = "ButterWorthFrequency": strValues(4) = .ButterWorthFrequency
    End With
    IniStuff.WriteSection "ErrorPlottingOptions", strKeys(), strValues()
    
    ' Write the Error Plotting Options -- Graph2D
    ReDim strKeys(0 To 8)
    ReDim strValues(0 To 8)
    With udtPrefsExpanded.ErrorPlottingOptions.Graph2DOptions
        strKeys(0) = "ShowPointSymbols": strValues(0) = .ShowPointSymbols
        strKeys(1) = "DrawLinesBetweenPoints": strValues(1) = .DrawLinesBetweenPoints
        strKeys(2) = "ShowGridlines": strValues(2) = .ShowGridLines
        strKeys(3) = "AutoScaleXAxis": strValues(3) = .AutoScaleXAxis
        strKeys(4) = "PointSizePixels": strValues(4) = .PointSizePixels
        strKeys(5) = "LineWidthPixels": strValues(5) = .LineWidthPixels
        strKeys(6) = "CenterYAxis": strValues(6) = .CenterYAxis
        strKeys(7) = "ShowSmoothedData": strValues(7) = .ShowSmoothedData
        strKeys(8) = "ShowPeakEdges": strValues(8) = .ShowPeakEdges
    End With
    IniStuff.WriteSection "ErrorPlottingOptionsGraph2D", strKeys(), strValues()
    
    ' Write the Error Plotting Options -- Graph3D
    ReDim strKeys(0 To 5)
    ReDim strValues(0 To 5)
    With udtPrefsExpanded.ErrorPlottingOptions.Graph3DOptions
        strKeys(0) = "ContourLevelsCount": strValues(0) = .ContourLevelsCount
        strKeys(1) = "Perspective": strValues(1) = .Perspective
        strKeys(2) = "Elevation": strValues(2) = .Elevation
        strKeys(3) = "YRotation": strValues(3) = .YRotation
        strKeys(4) = "ZRotation": strValues(4) = .ZRotation
        strKeys(5) = "AnnotationFontSize": strValues(5) = .AnnotationFontSize
    End With
    IniStuff.WriteSection "ErrorPlottingOptionsGraph3D", strKeys(), strValues()
    
    ' Write the noise removal options
    ReDim strKeys(0 To 11)
    ReDim strValues(0 To 11)
    With udtPrefsExpanded.NoiseRemovalOptions
        strKeys(0) = "SearchTolerancePPMDefault": strValues(0) = .SearchTolerancePPMDefault
        strKeys(1) = "SearchTolerancePPMAutoRemoval": strValues(1) = .SearchTolerancePPMAutoRemoval
        strKeys(2) = "PercentageThresholdToExcludeSlice": strValues(2) = .PercentageThresholdToExcludeSlice
        strKeys(3) = "PercentageThresholdToAddNeighborToSearchSlice": strValues(3) = .PercentageThresholdToAddNeighborToSearchSlice
        strKeys(4) = "LimitMassRange": strValues(4) = .LimitMassRange
        strKeys(5) = "MassStart": strValues(5) = .MassStart
        strKeys(6) = "MassEnd": strValues(6) = .MassEnd
        strKeys(7) = "LimitScanRange": strValues(7) = .LimitScanRange
        strKeys(8) = "ScanStart": strValues(8) = .ScanStart
        strKeys(9) = "ScanEnd": strValues(9) = .ScanEnd
        strKeys(10) = "SearchScope": strValues(10) = .SearchScope
        strKeys(11) = "RequireIdenticalCharge": strValues(11) = .RequireIdenticalCharge
    End With
    IniStuff.WriteSection "NoiseRemovalOptions", strKeys(), strValues()
    frmProgress.UpdateProgressBar 2
    
    ' Write the Refine MS Data options
    ReDim strKeys(0 To 22)
    ReDim strValues(0 To 22)
    With udtPrefsExpanded.RefineMSDataOptions
        strKeys(0) = "MinimumPeakHeight": strValues(0) = .MinimumPeakHeight
        strKeys(1) = "MinimumSignalToNoiseRatioForLowAbundancePeaks": strValues(1) = .MinimumSignalToNoiseRatioForLowAbundancePeaks
        strKeys(2) = "PercentageOfMaxForFindingWidth": strValues(2) = .PercentageOfMaxForFindingWidth
        strKeys(3) = "MassCalibrationMaximumShift": strValues(3) = .MassCalibrationMaximumShift
        strKeys(4) = "MassCalibrationTolType": strValues(4) = .MassCalibrationTolType
        strKeys(5) = "ToleranceRefinementMethod": strValues(5) = .ToleranceRefinementMethod
        strKeys(6) = "UseMinMaxIfOutOfRange": strValues(6) = .UseMinMaxIfOutOfRange
        strKeys(7) = "MassToleranceMinimum": strValues(7) = .MassToleranceMinimum
        strKeys(8) = "MassToleranceMaximum": strValues(8) = .MassToleranceMaximum
        strKeys(9) = "MassToleranceAdjustmentMultiplier": strValues(9) = .MassToleranceAdjustmentMultiplier
        strKeys(10) = "NETToleranceMinimum": strValues(10) = .NETToleranceMinimum
        strKeys(11) = "NETToleranceMaximum": strValues(11) = .NETToleranceMaximum
        strKeys(12) = "NETToleranceAdjustmentMultiplier": strValues(12) = .NETToleranceAdjustmentMultiplier
        strKeys(13) = "IncludeInternalStdMatches": strValues(13) = .IncludeInternalStdMatches
        strKeys(14) = "UseUMCClassStats": strValues(14) = .UseUMCClassStats
        strKeys(15) = "MinimumSLiC": strValues(15) = .MinimumSLiC
        strKeys(16) = "MaximumAbundance": strValues(16) = .MaximumAbundance
    
        strKeys(17) = "EMMassErrorPeakToleranceEstimatePPM": strValues(17) = .EMMassErrorPeakToleranceEstimatePPM
        strKeys(18) = "EMNETErrorPeakToleranceEstimate": strValues(18) = .EMNETErrorPeakToleranceEstimate
        strKeys(19) = "EMIterationCount": strValues(19) = .EMIterationCount
        strKeys(20) = "EMPercentOfDataToExclude": strValues(20) = .EMPercentOfDataToExclude
        strKeys(21) = "EMMassTolRefineForceUseSingleDataPointErrors": strValues(21) = .EMMassTolRefineForceUseSingleDataPointErrors
        strKeys(22) = "EMNETTolRefineForceUseSingleDataPointErrors": strValues(22) = .EMNETTolRefineForceUseSingleDataPointErrors
    End With
    IniStuff.WriteSection "RefineMSDataOptions", strKeys(), strValues()
        
    ' Write the TIC Plotting Options
    ReDim strKeys(0 To 17)
    ReDim strValues(0 To 17)
    With udtPrefsExpanded.TICAndBPIPlottingOptions
        strKeys(0) = "PlotNETOnXAxis": strValues(0) = .PlotNETOnXAxis
        strKeys(1) = "NormalizeYAxis": strValues(1) = .NormalizeYAxis
        strKeys(2) = "SmoothUsingMovingAverage": strValues(2) = .SmoothUsingMovingAverage
        strKeys(3) = "MovingAverageWindowWidth": strValues(3) = .MovingAverageWindowWidth
        strKeys(4) = "TimeDomainDataMaxValue": strValues(4) = .TimeDomainDataMaxValue
        With .Graph2DOptions
            strKeys(5) = "ShowPointSymbols": strValues(5) = .ShowPointSymbols
            strKeys(6) = "DrawLinesBetweenPoints": strValues(6) = .DrawLinesBetweenPoints
            strKeys(7) = "ShowGridlines": strValues(7) = .ShowGridLines
            strKeys(8) = "AutoScaleXAxis": strValues(8) = .AutoScaleXAxis
            strKeys(9) = "PointSizePixels": strValues(9) = .PointSizePixels
            strKeys(10) = "PointShape": strValues(10) = .PointShape
            strKeys(11) = "PointAndLineColor": strValues(11) = .PointAndLineColor
            strKeys(12) = "LineWidthPixels": strValues(12) = .LineWidthPixels
        End With
        strKeys(13) = "PointShapeSeries2": strValues(13) = .PointShapeSeries2
        strKeys(14) = "PointAndLineColorSeries2": strValues(14) = .PointAndLineColorSeries2
        
        strKeys(15) = "ClipOutliers": strValues(15) = .ClipOutliers
        strKeys(16) = "ClipOutliersFactor": strValues(16) = .ClipOutliersFactor
        
        strKeys(17) = "KeepWindowOnTop": strValues(17) = .KeepWindowOnTop
    End With
    IniStuff.WriteSection "TICAndBPIPlottingOptions", strKeys(), strValues()
    
    ' Write the Pair Browser Options
    ReDim strKeys(0 To 19)
    ReDim strValues(0 To 19)
    With udtPrefsExpanded.PairBrowserPlottingOptions
        strKeys(0) = "SortOrder": strValues(0) = .SortOrder
        strKeys(1) = "SortDescending": strValues(1) = .SortDescending
        strKeys(2) = "AutoZoom2DPlot": strValues(2) = .AutoZoom2DPlot
        strKeys(3) = "HighlightMembers": strValues(3) = .HighlightMembers
        strKeys(4) = "PlotAllChargeStates": strValues(4) = .PlotAllChargeStates
        strKeys(5) = "FixedDimensionsForAutoZoom": strValues(5) = .FixedDimensionsForAutoZoom
        
        strKeys(6) = "MassRangeZoom": strValues(6) = .MassRangeZoom
        strKeys(7) = "MassRangeUnits": strValues(7) = .MassRangeUnits
        strKeys(8) = "ScanRangeZoom": strValues(8) = .ScanRangeZoom
        strKeys(9) = "ScanRangeUnits": strValues(9) = .ScanRangeUnits
        
        With .Graph2DOptions
            strKeys(10) = "ShowPointSymbols": strValues(10) = .ShowPointSymbols
            strKeys(11) = "DrawLinesBetweenPoints": strValues(11) = .DrawLinesBetweenPoints
            strKeys(12) = "ShowGridlines": strValues(12) = .ShowGridLines
            strKeys(13) = "PointSizePixels": strValues(13) = .PointSizePixels
            strKeys(14) = "PointShape": strValues(14) = .PointShape
            strKeys(15) = "PointAndLineColor": strValues(15) = .PointAndLineColor
            strKeys(16) = "LineWidthPixels": strValues(16) = .LineWidthPixels
        End With
    
        strKeys(17) = "PointShapeHeavy": strValues(17) = .PointShapeHeavy
        strKeys(18) = "PointAndLineColorHeavy": strValues(18) = .PointAndLineColorHeavy
        strKeys(19) = "KeepWindowOnTop": strValues(19) = .KeepWindowOnTop
    End With
    IniStuff.WriteSection "PairBrowserOptions", strKeys(), strValues()
    
    ' Write the UMC Browser Options
    ReDim strKeys(0 To 17)
    ReDim strValues(0 To 17)
    With udtPrefsExpanded.UMCBrowserPlottingOptions
        strKeys(0) = "SortOrder": strValues(0) = .SortOrder
        strKeys(1) = "SortDescending": strValues(1) = .SortDescending
        strKeys(2) = "AutoZoom2DPlot": strValues(2) = .AutoZoom2DPlot
        strKeys(3) = "HighlightMembers": strValues(3) = .HighlightMembers
        strKeys(4) = "PlotAllChargeStates": strValues(4) = .PlotAllChargeStates
        strKeys(5) = "FixedDimensionsForAutoZoom": strValues(5) = .FixedDimensionsForAutoZoom
        
        strKeys(6) = "MassRangeZoom": strValues(6) = .MassRangeZoom
        strKeys(7) = "MassRangeUnits": strValues(7) = .MassRangeUnits
        strKeys(8) = "ScanRangeZoom": strValues(8) = .ScanRangeZoom
        strKeys(9) = "ScanRangeUnits": strValues(9) = .ScanRangeUnits
        
        With .Graph2DOptions
            strKeys(10) = "ShowPointSymbols": strValues(10) = .ShowPointSymbols
            strKeys(11) = "DrawLinesBetweenPoints": strValues(11) = .DrawLinesBetweenPoints
            strKeys(12) = "ShowGridlines": strValues(12) = .ShowGridLines
            strKeys(13) = "PointSizePixels": strValues(13) = .PointSizePixels
            strKeys(14) = "PointShape": strValues(14) = .PointShape
            strKeys(15) = "PointAndLineColor": strValues(15) = .PointAndLineColor
            strKeys(16) = "LineWidthPixels": strValues(16) = .LineWidthPixels
        End With
    
        strKeys(17) = "KeepWindowOnTop": strValues(17) = .KeepWindowOnTop
    End With
    IniStuff.WriteSection "UMCBrowserOptions", strKeys(), strValues()
    
    ' Write the Pair Search Options
    ReDim strKeys(0 To 37)
    ReDim strValues(0 To 37)
    With udtPrefsExpanded.PairSearchOptions
        With .SearchDef
            strKeys(0) = "DeltaMass": strValues(0) = .DeltaMass
            strKeys(1) = "DeltaMassTolerance": strValues(1) = .DeltaMassTolerance
            strKeys(2) = "AutoCalculateDeltaMinMaxCount": strValues(2) = .AutoCalculateDeltaMinMaxCount
            
            strKeys(3) = "DeltaCountMin": strValues(3) = .DeltaCountMin
            strKeys(4) = "DeltaCountMax": strValues(4) = .DeltaCountMax
            strKeys(5) = "DeltaStepSize": strValues(5) = .DeltaStepSize
            
            strKeys(6) = "LightLabelMass": strValues(6) = .LightLabelMass
            strKeys(7) = "HeavyLightMassDifference": strValues(7) = .HeavyLightMassDifference
            strKeys(8) = "LabelCountMin": strValues(8) = .LabelCountMin
            strKeys(9) = "LabelCountMax": strValues(9) = .LabelCountMax
            strKeys(10) = "MaxDifferenceInNumberOfLightHeavyLabels": strValues(10) = .MaxDifferenceInNumberOfLightHeavyLabels
            
            strKeys(11) = "RequireUMCOverlap": strValues(11) = .RequireUMCOverlap
            strKeys(12) = "RequireUMCOverlapAtApex": strValues(12) = .RequireUMCOverlapAtApex
            
            strKeys(13) = "ScanTolerance": strValues(13) = .ScanTolerance
            strKeys(14) = "ScanToleranceAtApex": strValues(14) = .ScanToleranceAtApex
            
            strKeys(15) = "ERInclusionMin": strValues(15) = .ERInclusionMin
            strKeys(16) = "ERInclusionMax": strValues(16) = .ERInclusionMax
            
            strKeys(17) = "RequireMatchingChargeStatesForPairMembers": strValues(17) = .RequireMatchingChargeStatesForPairMembers
            strKeys(18) = "UseIdenticalChargesForER": strValues(18) = .UseIdenticalChargesForER
            strKeys(19) = "ComputeERScanByScan": strValues(19) = .ComputeERScanByScan
            strKeys(20) = "AverageERsAllChargeStates": strValues(20) = .AverageERsAllChargeStates
            strKeys(21) = "AverageERsWeightingMode": strValues(21) = .AverageERsWeightingMode
            strKeys(22) = "ERCalcType": strValues(22) = .ERCalcType
        
            strKeys(23) = "RemoveOutlierERs": strValues(23) = .RemoveOutlierERs
            strKeys(24) = "RemoveOutlierERsIterate": strValues(24) = .RemoveOutlierERsIterate
            strKeys(25) = "RemoveOutlierERsMinimumDataPointCount": strValues(25) = .RemoveOutlierERsMinimumDataPointCount
            strKeys(26) = "RemoveOutlierERsConfidenceLevel": strValues(26) = .RemoveOutlierERsConfidenceLevel
        End With
        
        strKeys(27) = "PairSearchMode": strValues(27) = .PairSearchMode
        
        strKeys(28) = "AutoExcludeOutOfERRange": strValues(28) = .AutoExcludeOutOfERRange
        strKeys(29) = "AutoExcludeAmbiguous": strValues(29) = .AutoExcludeAmbiguous
        strKeys(30) = "KeepMostConfidentAmbiguous": strValues(30) = .KeepMostConfidentAmbiguous
        
        strKeys(31) = "AutoAnalysisRemovePairMemberHitsAfterDBSearch": strValues(31) = .AutoAnalysisRemovePairMemberHitsAfterDBSearch
        strKeys(32) = "AutoAnalysisRemovePairMemberHitsRemoveHeavy": strValues(32) = .AutoAnalysisRemovePairMemberHitsRemoveHeavy
        
        strKeys(33) = "AutoAnalysisSavePairsToTextFile": strValues(33) = .AutoAnalysisSavePairsToTextFile
        strKeys(34) = "AutoAnalysisSavePairsStatisticsToTextFile": strValues(34) = .AutoAnalysisSavePairsStatisticsToTextFile
        
        strKeys(35) = "NETAdjustmentPairedSearchUMCSelection": strValues(35) = .NETAdjustmentPairedSearchUMCSelection
        strKeys(36) = "OutlierRemovalUsesSymmetricERs": strValues(36) = .OutlierRemovalUsesSymmetricERs
    
        strKeys(37) = "AutoAnalysisDeltaMassAddnlCount": strValues(37) = .AutoAnalysisDeltaMassAddnlCount
        
        If .AutoAnalysisDeltaMassAddnlCount > 0 Then
            intTargetIndexBase = UBound(strKeys) + 1
            ReDim Preserve strKeys(intTargetIndexBase + .AutoAnalysisDeltaMassAddnlCount - 1)
            ReDim Preserve strValues(intTargetIndexBase + .AutoAnalysisDeltaMassAddnlCount - 1)
            
            For intIndex = 0 To .AutoAnalysisDeltaMassAddnlCount - 1
                strKeys(intTargetIndexBase + intIndex) = "AutoAnalysisDeltaMassAddnl" & Trim(intIndex + 1)
                strValues(intTargetIndexBase + intIndex) = .AutoAnalysisDeltaMassAddnl(intIndex)
            Next intIndex
        Else
            intTargetIndexBase = UBound(strKeys) + 1
            ReDim Preserve strKeys(intTargetIndexBase)
            ReDim Preserve strValues(intTargetIndexBase)
            
            strKeys(intTargetIndexBase) = "AutoAnalysisDeltaMassAddnl1"
            strValues(intTargetIndexBase) = "0"
        End If
        
    End With
    IniStuff.WriteSection "PairSearchOptions", strKeys(), strValues()
        
        
    ' Write the IReport Pair options
    ReDim strKeys(0 To 5)
    ReDim strValues(0 To 5)
    With udtPrefsExpanded.PairSearchOptions.SearchDef.IReportEROptions
        strKeys(0) = "Enabled": strValues(0) = .Enabled
        strKeys(1) = "NaturalAbundanceRatio2CoeffExponent": strValues(1) = .NaturalAbundanceRatio2Coeff.Exponent
        strKeys(2) = "NaturalAbundanceRatio2CoeffMultiplier": strValues(2) = .NaturalAbundanceRatio2Coeff.Multiplier
        strKeys(3) = "NaturalAbundanceRatio4CoeffExponent": strValues(3) = .NaturalAbundanceRatio4Coeff.Exponent
        strKeys(4) = "NaturalAbundanceRatio4CoeffMultiplier": strValues(4) = .NaturalAbundanceRatio4Coeff.Multiplier
        strKeys(5) = "MinimumFractionScansWithValidER": strValues(5) = .MinimumFractionScansWithValidER
    End With
    IniStuff.WriteSection "IReportEROptions", strKeys(), strValues()
    
    
    ' Write the MT tag Staleness Options
    ReDim strKeys(0 To 3)
    ReDim strValues(0 To 3)
    With udtPrefsExpanded.MassTagStalenessOptions
        strKeys(0) = "MaximumAgeLoadedMassTagsHours": strValues(0) = .MaximumAgeLoadedMassTagsHours
        strKeys(1) = "MaximumFractionAMTsWithNulls": strValues(1) = .MaximumFractionAMTsWithNulls
        strKeys(2) = "MaximumCountAMTsWithNulls": strValues(2) = .MaximumCountAMTsWithNulls
        strKeys(3) = "MinimumTimeBetweenReloadMinutes": strValues(3) = .MinimumTimeBetweenReloadMinutes
    End With
    IniStuff.WriteSection "MassTagStalenessOptions", strKeys(), strValues()
    
    
    ' Write the Match Score options
    ReDim strKeys(0 To 4)
    ReDim strValues(0 To 4)
    With udtPrefsExpanded.SLiCScoreOptions
        strKeys(0) = "MassPPMStDev": strValues(0) = .MassPPMStDev
        strKeys(1) = "NETStDev": strValues(1) = .NETStDev
        strKeys(2) = "UseAMTNETStDev": strValues(2) = .UseAMTNETStDev
        strKeys(3) = "MaxSearchDistanceMultiplier": strValues(3) = .MaxSearchDistanceMultiplier
        strKeys(4) = "AutoDefineSLiCScoreThresholds": strValues(4) = .AutoDefineSLiCScoreThresholds
    End With
    IniStuff.WriteSection "SLiCScoreOptions", strKeys(), strValues()
    
    
    ' Write the GraphicExport options
    ReDim strKeys(0 To 1)
    ReDim strValues(0 To 1)
    With udtPrefsExpanded.GraphicExportOptions
        strKeys(0) = "CopyEMFIncludeFilenameAndDate": strValues(0) = .CopyEMFIncludeFilenameAndDate
        strKeys(1) = "CopyEMFIncludeTextLabels": strValues(1) = .CopyEMFIncludeTextLabels
    End With
    IniStuff.WriteSection "GraphicExportOptions", strKeys(), strValues()
    
    ' Write the Auto Tolerance Refinement Options
    ReDim strKeys(0 To 10)
    ReDim strValues(0 To 10)
    With udtPrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement
        strKeys(0) = "DBSearchMWTol": strValues(0) = .DBSearchMWTol
        strKeys(1) = "DBSearchTolType": strValues(1) = .DBSearchTolType
        strKeys(2) = "DBSearchNETTol": strValues(2) = .DBSearchNETTol
        strKeys(3) = "DBSearchRegionShape": strValues(3) = .DBSearchRegionShape
        strKeys(4) = "DBSearchMinimumHighNormalizedScore": strValues(4) = .DBSearchMinimumHighNormalizedScore
        strKeys(5) = "DBSearchMinimumHighDiscriminantScore": strValues(5) = .DBSearchMinimumHighDiscriminantScore
        strKeys(6) = "DBSearchMinimumPeptideProphetProbability": strValues(6) = .DBSearchMinimumPeptideProphetProbability
        strKeys(7) = "RefineMassCalibration": strValues(7) = .RefineMassCalibration
        strKeys(8) = "RefineMassCalibrationOverridePPM": strValues(8) = .RefineMassCalibrationOverridePPM
        strKeys(9) = "RefineDBSearchMassTolerance": strValues(9) = .RefineDBSearchMassTolerance
        strKeys(10) = "RefineDBSearchNETTolerance": strValues(10) = .RefineDBSearchNETTolerance
    End With
    IniStuff.WriteSection "AutoToleranceRefinement", strKeys(), strValues()
    
    
    ' Write the Auto Analysis Options
    ReDim strKeys(0 To 43)
    ReDim strValues(0 To 43)
    With udtPrefsExpanded.AutoAnalysisOptions
        strKeys(0) = "MDType": strValues(0) = "1"
        strKeys(1) = "AutoRemoveNoiseStreaks": strValues(1) = .AutoRemoveNoiseStreaks
        strKeys(2) = "DoNotSaveOrExport": strValues(2) = .DoNotSaveOrExport
        strKeys(3) = "SkipFindUMCs": strValues(3) = .SkipFindUMCs
        strKeys(4) = "SkipGANETSlopeAndInterceptComputation": strValues(4) = .SkipGANETSlopeAndInterceptComputation
        strKeys(5) = "DBConnectionRetryAttemptMax": strValues(5) = .DBConnectionRetryAttemptMax
        strKeys(6) = "DBConnectionTimeoutSeconds": strValues(6) = .DBConnectionTimeoutSeconds
        strKeys(7) = "ExportResultsFileUsesJobNumberInsteadOfDataSetName": strValues(7) = .ExportResultsFileUsesJobNumberInsteadOfDataSetName
        strKeys(8) = "SaveGelFile": strValues(8) = .SaveGelFile
        strKeys(9) = "SaveGelFileOnError": strValues(9) = .SaveGelFileOnError
        strKeys(10) = "SavePictureGraphic": strValues(10) = .SavePictureGraphic
        strKeys(11) = "SavePictureGraphicFileTypeList": strValues(11) = "; Options are " & GetPictureGraphicsTypeList()
        strKeys(12) = "SavePictureGraphicFileType": strValues(12) = .SavePictureGraphicFileType
        strKeys(13) = "SavePictureWidthPixels": strValues(13) = .SavePictureWidthPixels
        strKeys(14) = "SavePictureHeightPixels": strValues(14) = .SavePictureHeightPixels
        strKeys(15) = "SaveInternalStdHitsAndData": strValues(15) = .SaveInternalStdHitsAndData
        
        strKeys(16) = "SaveErrorGraphicMass": strValues(16) = .SaveErrorGraphicMass
        strKeys(17) = "SaveErrorGraphicGANET": strValues(17) = .SaveErrorGraphicGANET
        strKeys(18) = "SaveErrorGraphic3D": strValues(18) = .SaveErrorGraphic3D
        strKeys(19) = "SaveErrorGraphicFileTypeList": strValues(19) = "; Options are " & GetErrorGraphicsTypeList()
        strKeys(20) = "SaveErrorGraphicFileType": strValues(20) = .SaveErrorGraphicFileType
        strKeys(21) = "SaveErrorGraphSizeWidthPixels": strValues(21) = .SaveErrorGraphSizeWidthPixels
        strKeys(22) = "SaveErrorGraphSizeHeightPixels": strValues(22) = .SaveErrorGraphSizeHeightPixels
        
        strKeys(23) = "SavePlotTIC": strValues(23) = .SavePlotTIC
        strKeys(24) = "SavePlotBPI": strValues(24) = .SavePlotBPI
        strKeys(25) = "SavePlotTICTimeDomain": strValues(25) = .SavePlotTICTimeDomain
        strKeys(26) = "SavePlotTICDataPointCounts": strValues(26) = .SavePlotTICDataPointCounts
        strKeys(27) = "SavePlotTICDataPointCountsHitsOnly": strValues(27) = .SavePlotTICDataPointCountsHitsOnly
        strKeys(28) = "SavePlotTICFromRawData": strValues(28) = .SavePlotTICFromRawData
        strKeys(29) = "SavePlotBPIFromRawData": strValues(29) = .SavePlotBPIFromRawData
        strKeys(30) = "SavePlotDeisotopingIntensityThresholds": strValues(30) = .SavePlotDeisotopingIntensityThresholds
        strKeys(31) = "SavePlotDeisotopingPeakCounts": strValues(31) = .SavePlotDeisotopingPeakCounts
        
        strKeys(32) = "OutputFileSeparationCharacter": strValues(32) = .OutputFileSeparationCharacter
        strKeys(33) = "PEKFileExtensionPreferenceOrder": strValues(33) = .PEKFileExtensionPreferenceOrder
        strKeys(34) = "WriteIDResultsByIonToTextFileAfterAutoSearches": strValues(34) = .WriteIDResultsByIonToTextFileAfterAutoSearches
        strKeys(35) = "SaveUMCStatisticsToTextFile": strValues(35) = .SaveUMCStatisticsToTextFile
        strKeys(36) = "IncludeORFNameInTextFileOutput": strValues(36) = .IncludeORFNameInTextFileOutput
        strKeys(37) = "SetIsConfirmedForDBSearchMatches": strValues(37) = .SetIsConfirmedForDBSearchMatches
        strKeys(38) = "AddQuantitationDescriptionEntry": strValues(38) = .AddQuantitationDescriptionEntry
        strKeys(39) = "ExportUMCsWithNoMatches": strValues(39) = .ExportUMCsWithNoMatches
        strKeys(40) = "DBSearchRegionShape": strValues(40) = .DBSearchRegionShape
        strKeys(41) = "UseLegacyDBForMTs": strValues(41) = .UseLegacyDBForMTs
        strKeys(42) = "IgnoreNETAdjustmentFailure": strValues(42) = .IgnoreNETAdjustmentFailure
        
        If .AutoAnalysisSearchModeCount < 0 Then .AutoAnalysisSearchModeCount = 0
        If .AutoAnalysisSearchModeCount > MAX_AUTO_SEARCH_MODE_COUNT Then .AutoAnalysisSearchModeCount = MAX_AUTO_SEARCH_MODE_COUNT
        strKeys(43) = "AutoAnalysisSearchModeCount": strValues(43) = .AutoAnalysisSearchModeCount
    End With
    IniStuff.WriteSection "AutoAnalysisOptions", strKeys(), strValues()
    
    ' Write the Auto Analysis Search Mode Settings
    ' Each search mode is written to its own section in the .Ini file
    With udtPrefsExpanded.AutoAnalysisOptions
        For intAutoSearchModeIndex = 0 To .AutoAnalysisSearchModeCount - 1
            With .AutoAnalysisSearchMode(intAutoSearchModeIndex)
                
                ' Write this Auto Analysis Search Mode's settings
                ReDim strKeys(0 To 19)
                ReDim strValues(0 To 19)
                
                strKeys(0) = "SearchModeList": strValues(0) = "; Options are " & GetAutoAnalysisOptionsList()
                strKeys(1) = "SearchMode": strValues(1) = .SearchMode
                strKeys(2) = "AlternateOutputFolderPath": strValues(2) = .AlternateOutputFolderPath
                strKeys(3) = "WriteResultsToTextFile": strValues(3) = .WriteResultsToTextFile
                strKeys(4) = "ExportResultsToDatabase": strValues(4) = .ExportResultsToDatabase
                strKeys(5) = "ExportUMCMembers": strValues(5) = .ExportUMCMembers
                strKeys(6) = "PairSearchAssumeMassTagsAreLabeled": strValues(6) = .PairSearchAssumeMassTagsAreLabeled
                strKeys(7) = "InternalStdSearchMode": strValues(7) = .InternalStdSearchMode
                strKeys(8) = "DBSearchMinimumHighNormalizedScore": strValues(8) = .DBSearchMinimumHighNormalizedScore
                strKeys(9) = "DBSearchMinimumHighDiscriminantScore": strValues(9) = .DBSearchMinimumHighDiscriminantScore
                strKeys(10) = "DBSearchMinimumPeptideProphetProbability": strValues(10) = .DBSearchMinimumPeptideProphetProbability
                
                With .MassMods
                    strKeys(11) = "DynamicMods": strValues(11) = .DynamicMods
                    strKeys(12) = "N15InsteadOfN14": strValues(12) = .N15InsteadOfN14
                    strKeys(13) = "PEO": strValues(13) = .PEO
                    strKeys(14) = "ICATd0": strValues(14) = .ICATd0
                    strKeys(15) = "ICATd8": strValues(15) = .ICATd8
                    strKeys(16) = "Alkylation": strValues(16) = .Alkylation
                    strKeys(17) = "AlkylationMass": strValues(17) = .AlkylationMass
                    strKeys(18) = "ResidueToModify": strValues(18) = .ResidueToModify
                    strKeys(19) = "ResidueMassModification": strValues(19) = .ResidueMassModification
                End With
                
                IniStuff.WriteSection "AutoAnalysisSearchMode" & Trim(intAutoSearchModeIndex + 1), strKeys(), strValues()
            End With
            
        Next intAutoSearchModeIndex
    End With
    
    
    ' Write the Auto Analysis Filter Preferences
    ReDim strKeys(0 To 35)
    ReDim strValues(0 To 35)
    With udtPrefsExpanded.AutoAnalysisFilterPrefs
        strKeys(0) = "ExcludeDuplicates": strValues(0) = .ExcludeDuplicates
        strKeys(1) = "ExcludeDuplicatesTolerance": strValues(1) = .ExcludeDuplicatesTolerance
        
        strKeys(2) = "ExcludeIsoByFit": strValues(2) = .ExcludeIsoByFit
        strKeys(3) = "ExcludeIsoByFitMaxVal": strValues(3) = .ExcludeIsoByFitMaxVal
        
        strKeys(4) = "ExcludeIsoSecondGuess": strValues(4) = .ExcludeIsoSecondGuess
        strKeys(5) = "ExcludeIsoLessLikelyGuess": strValues(5) = .ExcludeIsoLessLikelyGuess
        
        strKeys(6) = "ExcludeCSByStdDev": strValues(6) = .ExcludeCSByStdDev
        strKeys(7) = "ExcludeCSByStdDevMaxVal": strValues(7) = .ExcludeCSByStdDevMaxVal
        
        strKeys(8) = "RestrictIsoByAbundance": strValues(8) = .RestrictIsoByAbundance
        strKeys(9) = "RestrictIsoAbundanceMin": strValues(9) = .RestrictIsoAbundanceMin
        strKeys(10) = "RestrictIsoAbundanceMax": strValues(10) = .RestrictIsoAbundanceMax
        
        strKeys(11) = "RestrictIsoByMass": strValues(11) = .RestrictIsoByMass
        strKeys(12) = "RestrictIsoMassMin": strValues(12) = .RestrictIsoMassMin
        strKeys(13) = "RestrictIsoMassMax": strValues(13) = .RestrictIsoMassMax
        
        strKeys(14) = "RestrictIsoByMZ": strValues(14) = .RestrictIsoByMZ
        strKeys(15) = "RestrictIsoMZMin": strValues(15) = .RestrictIsoMZMin
        strKeys(16) = "RestrictIsoMZMax": strValues(16) = .RestrictIsoMZMax
        
        strKeys(17) = "RestrictIsoByChargeState": strValues(17) = .RestrictIsoByChargeState
        strKeys(18) = "RestrictIsoChargeStateMin": strValues(18) = .RestrictIsoChargeStateMin
        strKeys(19) = "RestrictIsoChargeStateMax": strValues(19) = .RestrictIsoChargeStateMax
        
        strKeys(20) = "RestrictCSByAbundance": strValues(20) = .RestrictCSByAbundance
        strKeys(21) = "RestrictCSAbundanceMin": strValues(21) = .RestrictCSAbundanceMin
        strKeys(22) = "RestrictCSAbundanceMax": strValues(22) = .RestrictCSAbundanceMax
        
        strKeys(23) = "RestrictCSByMass": strValues(23) = .RestrictCSByMass
        strKeys(24) = "RestrictCSMassMin": strValues(24) = .RestrictCSMassMin
        strKeys(25) = "RestrictCSMassMax": strValues(25) = .RestrictCSMassMax
        
        strKeys(26) = "RestrictScanRange": strValues(26) = .RestrictScanRange
        strKeys(27) = "RestrictScanRangeMin": strValues(27) = .RestrictScanRangeMin
        strKeys(28) = "RestrictScanRangeMax": strValues(28) = .RestrictScanRangeMax
        
        strKeys(29) = "RestrictGANETRange": strValues(29) = .RestrictGANETRange
        strKeys(30) = "RestrictGANETRangeMin": strValues(30) = .RestrictGANETRangeMin
        strKeys(31) = "RestrictGANETRangeMax": strValues(31) = .RestrictGANETRangeMax
        
        strKeys(32) = "RestrictToEvenScanNumbersOnly": strValues(32) = .RestrictToEvenScanNumbersOnly
        strKeys(33) = "RestrictToOddScanNumbersOnly": strValues(33) = .RestrictToOddScanNumbersOnly
        
        ' Maximum data count filter
        strKeys(34) = "MaximumDataCountEnabled": strValues(34) = .MaximumDataCountEnabled
        strKeys(35) = "MaximumDataCountToLoad": strValues(35) = .MaximumDataCountToLoad
    End With
    IniStuff.WriteSection "AutoAnalysisFilterPrefs", strKeys(), strValues()
    
    If udtPrefsExpanded.AutoAnalysisDBInfoIsValid Then
        With udtDBSettingsSingle
            .IsDeleted = False
            
            ' Copy data from udtPrefsExpanded.AutoAnalysisDBInfo to udtDBSettingsSingle.AnalysisInfo
            .AnalysisInfo = udtPrefsExpanded.AutoAnalysisDBInfo
        
            ' Initialize the summary variables
            .ConnectionString = .AnalysisInfo.MTDB.ConnectionString
            .DatabaseName = ExtractDBNameFromConnectionString(.ConnectionString)
            
            .AMTsOnly = CBoolSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_ACCURATE_ONLY))
            .ConfirmedOnly = CBoolSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_CONFIRMED_ONLY))
            .LockersOnly = CBoolSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_LOCKERS_ONLY))
            .LimitToPMTsFromDataset = CBoolSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_LIMIT_TO_PMTS_FROM_DATASET))
            
            .MinimumHighNormalizedScore = CSngSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_MINIMUM_HIGH_NORMALIZED_SCORE))
            .MinimumHighDiscriminantScore = CSngSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE))
            .MinimumPeptideProphetProbability = CSngSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_MINIMUM_PEPTIDE_PROPHET_PROBABILITY))
            .MinimumPMTQualityScore = CSngSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_MINIMUM_PMT_QUALITY_SCORE))
            
            .ExperimentInclusionFilter = CStrSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_EXPERIMENT_INCLUSION_FILTER))
            .ExperimentExclusionFilter = CStrSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_EXPERIMENT_EXCLUSION_FILTER))
            .InternalStandardExplicit = CStrSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_INTERNAL_STANDARD_EXPLICIT))
            
            
            .NETValueType = CIntSafe(LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_NET_VALUE_TYPE))
            
            strMassTagSubsetID = LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_SUBSET, False, ENTRY_NOT_FOUND)
            If strMassTagSubsetID = ENTRY_NOT_FOUND Then
                .MassTagSubsetID = -1
            Else
                .MassTagSubsetID = CLngSafe(strMassTagSubsetID)
            End If
            .ModificationList = LookupCollectionArrayValueByName(.AnalysisInfo.MTDB.DBStuffArray(), .AnalysisInfo.MTDB.DBStuffArrayCount, NAME_INC_LIST)
            
            .SelectedMassTagCount = 0
        End With
        
        IniFileWriteSingleDBConnection IniStuff, "AutoAnalysisDBInfo", udtDBSettingsSingle, False
    Else
        ' See if a RecentDBConnection entry is present in RECENT_DB_INI_FILENAME
        ' If it is, write it as the default AutoAnalysis settings
        strDBIniFilePath = AppendToPath(App.Path, RECENT_DB_INI_FILENAME)
        If FileExists(strDBIniFilePath) Then
            Set DBIniStuff = New clsIniStuff
            DBIniStuff.FileName = strDBIniFilePath
            
            strSectionName = RECENT_DB_CONNECTIONS_SECTION_NAME & "_" & RECENT_DB_CONNECTION_SUBSECTION_NAME & Trim(0)
            If IniFileReadSingleDBConnection(DBIniStuff, strSectionName, udtDBSettingsSingle) Then
                IniFileWriteSingleDBConnection IniStuff, "AutoAnalysisDBInfo", udtDBSettingsSingle, False
            End If
            
            Set DBIniStuff = Nothing
        End If
    End If
    
    
    If Not APP_BUILD_DISABLE_MTS Then
        ' Write the DMSConnectionInfo
        ReDim strKeys(0 To 0)
        ReDim strValues(0 To 0)
        With udtPrefsExpanded.DMSConnectionInfo
            strKeys(0) = "ConnectionString": strValues(0) = .ConnectionString
        End With
        IniStuff.WriteSection "DMSConnectionInfo", strKeys(), strValues()
    
    
        ' Write the MTSConnectionInfo
        ReDim strKeys(0 To 20)
        ReDim strValues(0 To 20)
        With udtPrefsExpanded.MTSConnectionInfo
            strKeys(0) = "ConnectionString": strValues(0) = .ConnectionString
            
            strKeys(1) = "spAddQuantitationDescription": strValues(1) = .spAddQuantitationDescription
            strKeys(2) = "spGetLockers": strValues(2) = .spGetLockers
            strKeys(3) = "spGetMassTagMatchCount": strValues(3) = .spGetMassTagMatchCount
            strKeys(4) = "spGetMassTags": strValues(4) = .spGetMassTags
            strKeys(5) = "spGetMassTagsSubset": strValues(5) = .spGetMassTagsSubset
            strKeys(6) = "spGetPMResultStats": strValues(6) = .spGetPMResultStats
            strKeys(7) = "spPutAnalysis": strValues(7) = .spPutAnalysis
            strKeys(8) = "spPutUMC": strValues(8) = .spPutUMC
            strKeys(9) = "spPutUMCMember": strValues(9) = .spPutUMCMember
            strKeys(10) = "spPutUMCMatch": strValues(10) = .spPutUMCMatch
            strKeys(11) = "spPutUMCInternalStdMatch": strValues(11) = .spPutUMCInternalStdMatch
            strKeys(12) = "spEditGANET": strValues(12) = .spEditGANET
            strKeys(13) = "spGetORFs": strValues(13) = .spGetORFs
            strKeys(14) = "spGetORFSeq": strValues(14) = .spGetORFSeq
            strKeys(15) = "spGetORFIDs": strValues(15) = .spGetORFIDs
            strKeys(16) = "spGetORFRecord": strValues(16) = .spGetORFRecord
            strKeys(17) = "spGetMassTagSeq": strValues(17) = .spGetMassTagSeq
            strKeys(18) = "spGetMassTagNames": strValues(18) = .spGetMassTagNames
            strKeys(19) = "spGetInternalStandards": strValues(19) = .spGetInternalStandards
            strKeys(20) = "spGetDBSchemaVersion": strValues(20) = .spGetDBSchemaVersion
        End With
        IniStuff.WriteSection "MTSConnectionInfo", strKeys(), strValues()
    End If
    
    
    frmProgress.UpdateProgressBar 3
    
    frmProgress.HideForm
    Set IniStuff = Nothing
    Exit Sub
    
SaveSettingsFileHandler:
    If Not udtPrefsExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving data to the Ini file (" & IniStuff.FileName & "); Sub IniFileSaveSettings in Settings.Bas" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Print "Error in IniFileSaveSettings: " & Err.Description
        Debug.Assert False
        LogErrors Err.Number, "Settings.Bas->IniFileSaveSettings"
    End If
    frmProgress.HideForm
    Set IniStuff = Nothing

End Sub

Public Sub IniFileReadRecentDatabaseConnections(udtRecentDBSettings() As udtDBSettingsType, intRecentDBCount As Integer)
    ' Read the ini file to populate udtRecentDBSettings()
    
    Dim udtDBSettingsSingle As udtDBSettingsType
    Dim strSectionName As String
    Dim intConnectionIndex As Integer
    Dim intConnectionCount As Integer
    
    Dim IniStuff As New clsIniStuff
    
    IniStuff.FileName = AppendToPath(App.Path, RECENT_DB_INI_FILENAME)
    
    ' Determine the number of recent DB connections in the ini file
    intConnectionCount = GetIniFileSettingLng(IniStuff, RECENT_DB_CONNECTIONS_SECTION_NAME, RECENT_DB_CONNECTIONS_KEY_COUNT_NAME)
    
    If intConnectionCount > RECENT_DB_CONNECTIONS_MAX_COUNT Then
        intConnectionCount = RECENT_DB_CONNECTIONS_MAX_COUNT
    End If
    
    intRecentDBCount = 0
    ReDim udtRecentDBSettings(0)
    
    For intConnectionIndex = 0 To intConnectionCount - 1
        strSectionName = RECENT_DB_CONNECTIONS_SECTION_NAME & "_" & RECENT_DB_CONNECTION_SUBSECTION_NAME & Trim(intConnectionIndex)
        
        If IniFileReadSingleDBConnection(IniStuff, strSectionName, udtDBSettingsSingle) Then
            
            intRecentDBCount = intRecentDBCount + 1
            ReDim Preserve udtRecentDBSettings(intRecentDBCount - 1)
        
            ' Read the values from the ini file and store in udtRecentDBSettings(intRecentDBCount - 1)
            udtRecentDBSettings(intRecentDBCount - 1) = udtDBSettingsSingle
            
        End If
                    
    Next intConnectionIndex
                    
    Set IniStuff = Nothing
    
End Sub

Private Function IniFileReadSingleDBConnection(objIniStuff As clsIniStuff, strSectionName As String, udtDBSettingsSingle As udtDBSettingsType) As Boolean
    ' Returns True if success, False if failure
    
    Dim lngInfoVersion As Long
    Dim intDBStuffArrayIndex As Integer, lngCollectionIndex As Integer
    Dim lngArrayCount As Long
    Dim strKeys() As String
    Dim strValues() As String
    Dim blnSuccess As Boolean
    Dim strNameToAdd As String, strValueToAdd As String, strCurrentValue As String
    Dim strMassTagSubsetID As String
    
    Dim udtDefaultGelAnalysisInfo As udtGelAnalysisInfoType

    On Error GoTo IniFileReadSingleDBConnectionErrorHandler

    Dim objMTDBInfoRetriever As New MTDBInfoRetriever

    lngInfoVersion = GetIniFileSettingLng(objIniStuff, strSectionName, RECENT_DB_CONNECTION_INFOVERSION_NAME, 0)
    If lngInfoVersion = RECENT_DB_CONNECTION_INFOVERSION Then
            
        ' Note: ReadSection() returns True if success
        If objIniStuff.ReadSection(strSectionName, strKeys(), strValues()) Then
        
            lngArrayCount = UBound(strKeys())
            
            With udtDBSettingsSingle
                .ConnectionString = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "ConnectionString")
                .DatabaseName = ExtractDBNameFromConnectionString(.ConnectionString)
                .DBSchemaVersion = CSngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "DBSchemaVersion"))
                .AMTsOnly = CBoolSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "AmtsOnly"))
                .ConfirmedOnly = CBoolSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "ConfirmedOnly"))
                .LockersOnly = CBoolSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "LockersOnly"))
                .LimitToPMTsFromDataset = CBoolSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "LimitToPMTsFromDataset"))
                
                .MinimumHighNormalizedScore = CSngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MinimumHighNormalizedScore"))
                .MinimumHighDiscriminantScore = CSngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MinimumHighDiscriminantScore"))
                .MinimumPeptideProphetProbability = CSngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MinimumPeptideProphetProbability"))
                .MinimumPMTQualityScore = CSngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MinimumPMTQualityScore"))
                
                .ExperimentInclusionFilter = CStrSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "ExperimentInclusionFilter"))
                .ExperimentExclusionFilter = CStrSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "ExperimentExclusionFilter"))
                .InternalStandardExplicit = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "InternalStandardExplicit")
                
                .NETValueType = CIntSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "NETValueType"))
                
                strMassTagSubsetID = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MassTagSubsetID", False, ENTRY_NOT_FOUND)
                If strMassTagSubsetID = ENTRY_NOT_FOUND Then
                    .MassTagSubsetID = -1
                Else
                    .MassTagSubsetID = CLngSafe(strMassTagSubsetID)
                End If
                .ModificationList = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "ModificationList")
                .SelectedMassTagCount = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "SelectedMassTagCount"))
                
                With .AnalysisInfo
                    .Analysis_Tool = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "AnalysisTool")
                    .Created = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Created")
                    .Dataset = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Dataset")
                    .Dataset_Folder = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Dataset_Folder")
                    .Dataset_ID = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Dataset_ID"))
                    .Desc_DataFolder = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Desc_DataFolder")
                    .Desc_Type = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Desc_Type")
                    .Duration = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Duration"))
                    .Experiment = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Experiment")
                    .GANET_Fit = CDblSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "GANET_Fit"))
                    .GANET_Intercept = CDblSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "GANET_Intercept"))
                    .GANET_Slope = CDblSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "GANET_Slope"))
                    .Instrument_Class = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Instrument_Class")
                    .Job = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Job"))
                    .MD_Date = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MD_Date")
                    .MD_file = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MD_file")
                    .MD_Parameters = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MD_Parameters")
                    .MD_Reference_Job = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MD_Reference_Job"))
                    .MD_State = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MD_State"))
                    .MD_Type = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "MD_Type"))
                    .NET_Intercept = CDblSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "NET_Intercept"))
                    .NET_Slope = CDblSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "NET_Slope"))
                    .NET_TICFit = CDblSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "NET_TICFit"))
                    .Organism = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Organism")
                    .Organism_DB_Name = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Organism_DB_Name")
                    .Parameter_File_Name = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Parameter_File_Name")
                    .ProcessingType = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "ProcessingType"))
                    .Results_Folder = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Results_Folder")
                    .Settings_File_Name = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Settings_File_Name")
                    .STATE = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "STATE"))
                    .Storage_Path = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Storage_Path")
                    .Total_Scans = CLngSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Total_Scans"))
                    .ValidAnalysisDataPresent = CBoolSafe(LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "ValidAnalysisDataPresent"))
                    .Vol_Client = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Vol_Client")
                    .Vol_Server = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "Vol_Server")
                
                    ' Read the MTDB settings
                    strSectionName = strSectionName & "_" & "MTDB"
                    blnSuccess = False
                    If objIniStuff.ReadSection(strSectionName, strKeys(), strValues()) Then
                        lngArrayCount = UBound(strKeys())
                        
                        If lngArrayCount > 0 Then
                            With .MTDB
                                ' Read the MTDB items that go in DBStuff
                                .DBStatus = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "DBStatus")
                                .DBStuffArrayCount = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "DBStuffCount")
    
                                ' .DBStuffArrayCount is usually 33
                                If .DBStuffArrayCount > 0 Then
                                    For intDBStuffArrayIndex = 0 To .DBStuffArrayCount - 1
                                        .DBStuffArray(intDBStuffArrayIndex).Name = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "DBStuffItem" & Trim(intDBStuffArrayIndex) & "Name")
                                        .DBStuffArray(intDBStuffArrayIndex).Value = LookupParallelStringArrayItemByName(strKeys(), strValues(), lngArrayCount, "DBStuffItem" & Trim(intDBStuffArrayIndex) & "Value")
                                    Next intDBStuffArrayIndex
                                End If
                            End With
                            blnSuccess = True
                        End If
                    End If
                    
                    objMTDBInfoRetriever.InitFilePath = glInitFile
                    objMTDBInfoRetriever.GetMTDBSchema
                    
                    ' Grab the default MTDB settings
                    FillGelAnalysisInfo udtDefaultGelAnalysisInfo, objMTDBInfoRetriever.fAnalysis
                    
                    With .MTDB
                        ' Make sure all of the required MTDB info is present
                        ' None will be present if strSectionName wasn't present in the .Ini file or if lngArrayCount = 0
                        ' That's OK, we'll fill it with default values (which are stored in the FaxA.Init file
                            
                        ' Note: collections are 1-based
                        For lngCollectionIndex = 1 To objMTDBInfoRetriever.fAnalysis.MTDB.DBStuff.Count
                            strNameToAdd = objMTDBInfoRetriever.fAnalysis.MTDB.DBStuff.Item(lngCollectionIndex).Name
                            strValueToAdd = objMTDBInfoRetriever.fAnalysis.MTDB.DBStuff.Item(lngCollectionIndex).Value
                            strCurrentValue = LookupCollectionArrayValueByName(.DBStuffArray(), .DBStuffArrayCount, strNameToAdd, False, ENTRY_NOT_FOUND)
                            If strCurrentValue = ENTRY_NOT_FOUND Then
                                AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, strNameToAdd, strValueToAdd
                            End If
                        Next lngCollectionIndex
                            
                        ' Create entries for the following based upon values read above
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_SUBSET, CStr(udtDBSettingsSingle.MassTagSubsetID)                  ' Not used with DB Schema Version 2
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_INC_LIST, udtDBSettingsSingle.ModificationList
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_CONFIRMED_ONLY, CStr(udtDBSettingsSingle.ConfirmedOnly)
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_ACCURATE_ONLY, CStr(udtDBSettingsSingle.AMTsOnly)                  ' Not used with DB Schema Version 2
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_LOCKERS_ONLY, CStr(udtDBSettingsSingle.LockersOnly)                ' Not used with DB Schema Version 2
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_LIMIT_TO_PMTS_FROM_DATASET, CStr(udtDBSettingsSingle.LimitToPMTsFromDataset)
                        
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_MINIMUM_HIGH_NORMALIZED_SCORE, CStr(udtDBSettingsSingle.MinimumHighNormalizedScore)
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE, CStr(udtDBSettingsSingle.MinimumHighDiscriminantScore)
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_MINIMUM_PEPTIDE_PROPHET_PROBABILITY, CStr(udtDBSettingsSingle.MinimumPeptideProphetProbability)
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_MINIMUM_PMT_QUALITY_SCORE, CStr(udtDBSettingsSingle.MinimumPMTQualityScore)
                        
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_EXPERIMENT_INCLUSION_FILTER, CStr(udtDBSettingsSingle.ExperimentInclusionFilter)
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_EXPERIMENT_EXCLUSION_FILTER, CStr(udtDBSettingsSingle.ExperimentExclusionFilter)
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_INTERNAL_STANDARD_EXPLICIT, CStr(udtDBSettingsSingle.InternalStandardExplicit)
                        
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_NET_VALUE_TYPE, CStr(udtDBSettingsSingle.NETValueType)
                        AddOrUpdateCollectionArrayItem .DBStuffArray(), .DBStuffArrayCount, NAME_GET_DB_SCHEMA_VERSION, "GetDBSchemaVersion"
                        
                        ' .DBStuffArrayCount was 29 in August 2004
                        ' .DBStuffArrayCount is 32 in December 2005
                        ' .DBStuffArrayCount is 33 in October 2006
                        Debug.Assert .DBStuffArrayCount = 33
    
                        ' Also update .ConnectionString
                        .ConnectionString = udtDBSettingsSingle.ConnectionString
                    End With
                    blnSuccess = True
                End With
                
                If .DBSchemaVersion = 0 And Len(.ConnectionString) > 0 Then
                    .DBSchemaVersion = LookupDBSchemaVersionViaCNString(.ConnectionString)
                End If

            End With
            
        Else
            blnSuccess = False
        End If
        
    Else
        blnSuccess = False
    End If
    
    If Not objMTDBInfoRetriever Is Nothing Then Set objMTDBInfoRetriever = Nothing

    udtDBSettingsSingle.IsDeleted = Not blnSuccess
    IniFileReadSingleDBConnection = blnSuccess
    Exit Function
    
IniFileReadSingleDBConnectionErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error reading single DB connection from file " & objIniStuff.FileName & "; Sub IniFileReadSingleDBConnection in Settings.Bas" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Print "Error in IniFileReadSingleDBConnection: " & Err.Description
        Debug.Assert False
        LogErrors Err.Number, "Settings.Bas->IniFileReadSingleDBConnection"
    End If
    IniFileReadSingleDBConnection = False
        
End Function

Private Sub IniFileWriteSingleDBConnection(objIniStuff As clsIniStuff, strSectionName As String, udtDBSettingsSingle As udtDBSettingsType, Optional blnIncludeDetailedAnalysisInfo As Boolean = False, Optional blnIncludeMtdbDBStuff As Boolean = False)
    ' When blnIncludeDetailedAnalysisInfo = False then does not record any specific values for the analysis
    ' Fields which will be ignored: AnalysisTool, Created, Dataset, Dataset_Folder, Dataset_ID, Desc_DataFolder,
    '                               Desc_Type, Duration, Experiment, Instrument_Class, Job, MD_*, Organism*
    '                               Parameter_File_Name, ProcessingType, Results_Folder, Settings_File_Name, State
    '                               Storage_Path, Total_Scans, Vol_Client, and Vol_Server
    
    Const MTDB_HEADER_ITEM_COUNT = 2
    Dim strKeys() As String, strValues() As String
    Dim intDBStuffArrayIndex As Integer, intDBStuffItemCount As Integer

    ' Store the settings from udtRecentDBSettings(intIndex) in the ini file
    With udtDBSettingsSingle
    
        If blnIncludeDetailedAnalysisInfo Then
            ReDim strKeys(0 To 52)
            ReDim strValues(0 To 52)
        Else
            ReDim strKeys(0 To 21)
            ReDim strValues(0 To 21)
        End If
            
        ' Write the version number
        strKeys(0) = RECENT_DB_CONNECTION_INFOVERSION_NAME: strValues(0) = CStr(RECENT_DB_CONNECTION_INFOVERSION)
        
        ' Write the header items (summary variables)
        Debug.Assert .ConnectionString = .AnalysisInfo.MTDB.ConnectionString
        strKeys(1) = "ConnectionString": strValues(1) = .ConnectionString
        strKeys(2) = "DBSchemaVersion": strValues(2) = CStr(.DBSchemaVersion)
        strKeys(3) = "AmtsOnly": strValues(3) = CStr(.AMTsOnly)
        strKeys(4) = "ConfirmedOnly": strValues(4) = CStr(.ConfirmedOnly)
        strKeys(5) = "LockersOnly": strValues(5) = CStr(.LockersOnly)
        strKeys(6) = "LimitToPMTsFromDataset": strValues(6) = CStr(.LimitToPMTsFromDataset)
        
        strKeys(7) = "MinimumHighNormalizedScore": strValues(7) = CStr(.MinimumHighNormalizedScore)
        strKeys(8) = "MinimumHighDiscriminantScore": strValues(8) = CStr(.MinimumHighDiscriminantScore)
        strKeys(9) = "MinimumPeptideProphetProbability": strValues(9) = CStr(.MinimumPeptideProphetProbability)
        strKeys(10) = "MinimumPMTQualityScore": strValues(10) = CStr(.MinimumPMTQualityScore)
        
        strKeys(11) = "ExperimentInclusionFilter": strValues(11) = .ExperimentInclusionFilter
        strKeys(12) = "ExperimentExclusionFilter": strValues(12) = .ExperimentExclusionFilter
        strKeys(13) = "InternalStandardExplicit": strValues(13) = .InternalStandardExplicit
        
        strKeys(14) = "NETValueType": strValues(14) = CStr(.NETValueType)
        
        strKeys(15) = "MassTagSubsetID": strValues(15) = CStr(.MassTagSubsetID)
        strKeys(16) = "ModificationList": strValues(16) = .ModificationList
        
        strKeys(17) = "SelectedMassTagCount": strValues(17) = CStr(.SelectedMassTagCount)
        
        ' Now write the values in .AnalysisInfo
        With .AnalysisInfo
            
            strKeys(18) = "GANET_Fit": strValues(18) = CStr(.GANET_Fit)
            strKeys(19) = "GANET_Intercept": strValues(19) = CStr(.GANET_Intercept)
            strKeys(20) = "GANET_Slope": strValues(20) = CStr(.GANET_Slope)
            strKeys(21) = "ValidAnalysisDataPresent": strValues(21) = CStr(.ValidAnalysisDataPresent)
            
            If blnIncludeDetailedAnalysisInfo Then
                strKeys(22) = "AnalysisTool": strValues(22) = .Analysis_Tool
                strKeys(23) = "Created": strValues(23) = .Created
                strKeys(24) = "Dataset": strValues(24) = .Dataset
                strKeys(25) = "Dataset_Folder": strValues(25) = .Dataset_Folder
                strKeys(26) = "Dataset_ID": strValues(26) = .Dataset_ID
                strKeys(27) = "Desc_DataFolder": strValues(27) = .Desc_DataFolder
                strKeys(28) = "Desc_Type": strValues(28) = .Desc_Type
                strKeys(29) = "Duration": strValues(29) = .Duration
                strKeys(30) = "Experiment": strValues(30) = .Experiment
                strKeys(31) = "Instrument_Class": strValues(31) = .Instrument_Class
                strKeys(32) = "Job": strValues(32) = .Job
                strKeys(33) = "MD_Date": strValues(33) = .MD_Date
                strKeys(34) = "MD_file": strValues(34) = .MD_file
                strKeys(35) = "MD_Parameters": strValues(35) = .MD_Parameters
                strKeys(36) = "MD_Reference_Job": strValues(36) = .MD_Reference_Job
                strKeys(37) = "MD_State": strValues(37) = .MD_State
                strKeys(38) = "MD_Type": strValues(38) = .MD_Type
                strKeys(39) = "NET_Intercept": strValues(39) = .NET_Intercept
                strKeys(40) = "NET_Slope": strValues(40) = .NET_Slope
                strKeys(41) = "NET_TICFit": strValues(41) = .NET_TICFit
                strKeys(42) = "Organism": strValues(42) = .Organism
                strKeys(43) = "Organism_DB_Name": strValues(43) = .Organism_DB_Name
                strKeys(44) = "Parameter_File_Name": strValues(44) = .Parameter_File_Name
                strKeys(45) = "ProcessingType": strValues(45) = .ProcessingType
                strKeys(46) = "Results_Folder": strValues(46) = .Results_Folder
                strKeys(47) = "Settings_File_Name": strValues(47) = .Settings_File_Name
                strKeys(48) = "State": strValues(48) = .STATE
                strKeys(49) = "Storage_Path": strValues(49) = .Storage_Path
                strKeys(50) = "Total_Scans": strValues(50) = .Total_Scans
                strKeys(51) = "Vol_Client": strValues(51) = .Vol_Client
                strKeys(52) = "Vol_Server": strValues(52) = .Vol_Server
            End If
            
            objIniStuff.WriteSection strSectionName, strKeys(), strValues()
    
            If blnIncludeMtdbDBStuff Then
                strSectionName = strSectionName & "_" & "MTDB"
                With .MTDB
                    ' Write the MTDB items
                    
                    ReDim strKeys(0 To MTDB_HEADER_ITEM_COUNT + .DBStuffArrayCount * 2 - 1)
                    ReDim strValues(0 To MTDB_HEADER_ITEM_COUNT + .DBStuffArrayCount * 2 - 1)
                    
                    strKeys(0) = "DBStatus": strValues(0) = CStr(.DBStatus)
                    strKeys(1) = "DBStuffCount": strValues(1) = CStr(0)         ' Note: This will be updated below
                    
                    ' We must use a separate record of the # of items to write (intDBStuffItemCount) since
                    '  we aren't writing to disk all of the items in DBStuffArray
                    intDBStuffItemCount = 0
                    For intDBStuffArrayIndex = 0 To .DBStuffArrayCount - 1
                        
                        Select Case .DBStuffArray(intDBStuffArrayIndex).Name
                        Case NAME_SUBSET            ' Do not write to disk; it is included in the summary variables above
                        Case NAME_INC_LIST          ' Do not write to disk; it is included in the summary variables above
                        Case NAME_CONFIRMED_ONLY    ' Do not write to disk; it is included in the summary variables above
                        Case NAME_ACCURATE_ONLY     ' Do not write to disk; it is included in the summary variables above
                        Case NAME_LOCKERS_ONLY      ' Do not write to disk; it is included in the summary variables above
                        Case NAME_LIMIT_TO_PMTS_FROM_DATASET        ' Do not write to disk; it is included in the summary variables above
                        Case NAME_MINIMUM_HIGH_NORMALIZED_SCORE     ' Do not write to disk; it is included in the summary variables above
                        Case NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE   ' Do not write to disk; it is included in the summary variables above
                        Case NAME_MINIMUM_PEPTIDE_PROPHET_PROBABILITY   ' Do not write to disk; it is included in the summary variables above
                        Case NAME_MINIMUM_PMT_QUALITY_SCORE         ' Do not write to disk; it is included in the summary variables above
                        Case NAME_EXPERIMENT_INCLUSION_FILTER       ' Do not write to disk; it is included in the summary variables above
                        Case NAME_EXPERIMENT_EXCLUSION_FILTER       ' Do not write to disk; it is included in the summary variables above
                        Case NAME_INTERNAL_STANDARD_EXPLICIT        ' Do not write to disk; it is included in the summary variables above
                        Case NAME_NET_VALUE_TYPE                    ' Do not write to disk; it is included in the summary variables above
                        Case Else
                            strKeys(MTDB_HEADER_ITEM_COUNT + intDBStuffItemCount * 2) = "DBStuffItem" & Trim(intDBStuffItemCount) & "Name"
                            strValues(MTDB_HEADER_ITEM_COUNT + intDBStuffItemCount * 2) = .DBStuffArray(intDBStuffArrayIndex).Name
                            
                            strKeys(MTDB_HEADER_ITEM_COUNT + intDBStuffItemCount * 2 + 1) = "DBStuffItem" & Trim(intDBStuffItemCount) & "Value"
                            strValues(MTDB_HEADER_ITEM_COUNT + intDBStuffItemCount * 2 + 1) = .DBStuffArray(intDBStuffArrayIndex).Value
                            intDBStuffItemCount = intDBStuffItemCount + 1
                        End Select
                    Next intDBStuffArrayIndex
                    
                    If intDBStuffItemCount > 0 Then
                        Debug.Assert intDBStuffItemCount = 36
                    End If
                    
                    ReDim Preserve strKeys(0 To MTDB_HEADER_ITEM_COUNT + intDBStuffItemCount * 2 - 1)
                    ReDim Preserve strValues(0 To MTDB_HEADER_ITEM_COUNT + intDBStuffItemCount * 2 - 1)
                    
                    strKeys(1) = "DBStuffCount": strValues(1) = (intDBStuffItemCount)
                    
                    objIniStuff.WriteSection strSectionName, strKeys(), strValues()
                
                End With
            End If
        End With
    End With

End Sub

Public Sub IniFileUpdateRecentDatabaseConnectionInfo(udtNewDBSettings As udtDBSettingsType)
    ' Look for an entry in the RecentDBConnections section in the ini file containing
    '  the table stored in udtNewDBSettings
    ' If Found, update with the new settings and move to entry 0
    ' If not found, add as entry 0 and shift the other entries down by one
    
    Dim udtRecentDBSettings() As udtDBSettingsType
    
    Dim intRecentDBCount As Integer, intIndex As Integer
    Dim intMatchIndex As Integer
    Dim strSectionName As String
    Dim blnSuccess As Boolean
    
    Dim IniStuff As New clsIniStuff

On Error GoTo IniFileUpdateRecentDatabaseConnectionInfoErrorHandler

    IniFileReadRecentDatabaseConnections udtRecentDBSettings(), intRecentDBCount
    
    ' Make sure udtNewDBSettings.ConnectionString and .DatabaseName is correct
    With udtNewDBSettings
        Debug.Assert .ConnectionString = .AnalysisInfo.MTDB.ConnectionString
        If Len(.ConnectionString) > 0 Then
            Debug.Assert .DatabaseName = ExtractDBNameFromConnectionString(.ConnectionString)
        End If
    End With
    
    intMatchIndex = -1
    For intIndex = 0 To intRecentDBCount - 1
        With udtRecentDBSettings(intIndex)
            If .DatabaseName = udtNewDBSettings.DatabaseName Then
                If .ConnectionString = udtNewDBSettings.ConnectionString Then
                    intMatchIndex = intIndex
                    .AnalysisInfo = udtNewDBSettings.AnalysisInfo
                    Exit For
                End If
            End If
        End With
    Next intIndex
    
    If intMatchIndex < 0 Then
        ' Item not found; increment intRecentDBCount
        If intRecentDBCount < RECENT_DB_CONNECTIONS_MAX_COUNT Then
            intRecentDBCount = intRecentDBCount + 1
            ReDim Preserve udtRecentDBSettings(intRecentDBCount)
        End If
        
        ' Update intMatchIndex so that the following for/loop works properly
        intMatchIndex = intRecentDBCount - 1
    End If
    
    ' Shift the items up one position (unless intMatchIndex = 0, then not necessary)
    If intMatchIndex <> 0 Then
        For intIndex = intMatchIndex To 1 Step -1
            udtRecentDBSettings(intIndex) = udtRecentDBSettings(intIndex - 1)
        Next intIndex
    End If
    
    ' Copy the new settings to udtRecentDBSettings(0)
    udtRecentDBSettings(0) = udtNewDBSettings
    
    IniStuff.FileName = AppendToPath(App.Path, RECENT_DB_INI_FILENAME)
    
    blnSuccess = IniStuff.WriteValue(RECENT_DB_CONNECTIONS_SECTION_NAME, RECENT_DB_CONNECTIONS_KEY_COUNT_NAME, Trim(intRecentDBCount))
    If Not blnSuccess Then
        Debug.Assert False
        Exit Sub
    End If
    
    For intIndex = 0 To intRecentDBCount - 1
        strSectionName = RECENT_DB_CONNECTIONS_SECTION_NAME & "_" & RECENT_DB_CONNECTION_SUBSECTION_NAME & Trim(intIndex)
        
        IniFileWriteSingleDBConnection IniStuff, strSectionName, udtRecentDBSettings(intIndex), False
    Next intIndex
    
    Set IniStuff = Nothing
    Exit Sub

IniFileUpdateRecentDatabaseConnectionInfoErrorHandler:
LogErrors Err.Number, "Settings.IniFileUpdateRecentDatabaseConnectionInfo"

End Sub

Public Function IsWinLoaded(ByVal sWinName As String) As Boolean
'this function is used to determine is some form (or application) loaded
Dim hwnd As Long
IsWinLoaded = False
hwnd = FindWindow(vbNullString, sWinName)
If (hwnd > 0 And IsWindow(hwnd)) Then IsWinLoaded = True
End Function

Public Function LaunchICR2LS() As Boolean
'launches ICR-2LS application
Dim ICR2LSTaskID As Long
Dim sCommandLine As String
Dim nWinStyle As Integer
On Error Resume Next

sCommandLine = sICR2LSCommand
nWinStyle = vbNormalFocus
ICR2LSTaskID = Shell(sCommandLine, nWinStyle)
If ICR2LSTaskID = 0 Then 'somethin's wrong
   MsgBox "Error starting ICR-2LS program. Choose Options from " _
   & "the Tools menu and check path to the ICR-2LS.EXE.", vbOKOnly
   LaunchICR2LS = False
Else
   LaunchICR2LS = True
End If
End Function

' No longer supported (March 2006)
''Private Sub ResolveAMTPrefs()
'''this is now complicated more than neccessary but it
'''will be easier to adjust in case of additional AMT parameters
''Dim aRes As Variant
''On Error GoTo err_ResolveAMT
''aRes = ResolvePrefsString(sAMTPref)
''glbPreferencesExpanded.LegacyAMTDBPath = aRes(1)
''Exit Sub
''
''err_ResolveAMT:
''ResetAMTPreferences
''End Sub
''
''Private Function GetAMTPrefs() As String
''GetAMTPrefs = glbPreferencesExpanded.LegacyAMTDBPath
''End Function
''
''Private Sub ResetAMTPreferences()
''glbPreferencesExpanded.LegacyAMTDBPath = ""
''End Sub

Public Sub ResetExpandedPreferences(udtPreferencesExpanded As udtPreferencesExpandedType, Optional strSingleSectionToReset As String = "", Optional blnApplyFilterOnIsotopicFit As Boolean = True)
    ' Use strSingleSectionToReset to reset a single section
    ' Note that strSingleSectionToReset _is_ case sensitive
    
    With udtPreferencesExpanded
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "General" Then
            ' General Options
            .MenuModeDefault = mmDBWithPairs
            .MenuModeIncludeObsolete = False
            .ExtendedFileSaveModePreferred = True
            
            .CopyPointsInViewIncludeSearchResultsChecked = True
            .CopyPointsInViewByUMCIncludeSearchResultsChecked = True
            
            .AutoAdjSize = True
            .AutoSizeMultiplier = 1
            
            .UsePEKBasedERValues = False
            .UseMassTagsWithNullMass = False
            .UseMassTagsWithNullNET = False
            
            .UseUMCConglomerateNET = True
            .NetAdjustmentUsesN15AMTMasses = False
            .NetAdjustmentMinHighNormalizedScore = 2.5
            .NetAdjustmentMinHighDiscriminantScore = 0.5
            
            .UMCDrawType = umcdt_ActualUMC
            
            .AMTSearchResultsBehavior = asrbAutoRemoveExisting
            .ICR2LSSpectrumViewZoomWindowWidthMZ = 5
            
            .LastInputFileMode = ifmInputFileModeConstants.ifmPEKFile
            .LegacyAMTDBPath = ""
            
            ResetUMCAdvancedStatsOptions .UMCAdvancedStatsOptions
            ResetUMCAutoRefineOptions .UMCAutoRefineOptions
            
            With .UMCIonNetOptions
                .ConnectionLengthPostFilterMaxNET = 0.2
                
                .UMCRepresentative = UMCFROMNet_REP_ABU
                .MakeSingleMemberClasses = False
            End With
            
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                With .AutoQueryPRISMOptions
                    .ConnectionStringQueryDB = PRISM_AUTOMATION_CONNECTION_STRING_DEFAULT
                    .RequestTaskSPName = PRISM_AUTOMATION_SP_REQUEST_TASK_DEFAULT
                    .SetTaskCompleteSPName = PRISM_AUTOMATION_SP_SET_COMPLETE_DEFAULT
                    .SetTaskToRestartSPName = PRISM_AUTOMATION_SP_RESTART_TASK_DEFAULT
                    .PostLogEntrySPName = PRISM_AUTOMATION_SP_POST_LOG_ENTRY_DEFAULT
                    .QueryIntervalSeconds = 600
                    .MinimumPriorityToProcess = 0
                    .MaximumPriorityToProcess = 10
                    .PreferredDatabaseToProcess = ""
                    .ServerForPreferredDatabase = ""
                    .ExclusivelyUseThisDatabase = False
                End With
            End If
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "NetAdjustmentUMCDistributionOptions" Then
            With .NetAdjustmentUMCDistributionOptions
                .RequireDispersedUMCSelection = True
                .SegmentCount = 10
                .MinimumUMCsPerSegmentPctTopAbuPct = 100
                .ScanPctStart = 5
                .ScanPctEnd = 95
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "ErrorPlottingOptions" Then
            With .ErrorPlottingOptions
                .MassRangePPM = 40          ' Range +-; full range is .MassRangePPM x 2
                .MassBinSizePPM = DEFAULT_MASS_BIN_SIZE_PPM
                .GANETRange = 0.15
                .GANETBinSize = DEFAULT_GANET_BIN_SIZE
                .ButterWorthFrequency = 0.2
                
                With .Graph2DOptions
                    .ShowPointSymbols = False
                    .DrawLinesBetweenPoints = True
                    .ShowGridLines = True
                    .AutoScaleXAxis = False
                    .PointSizePixels = 5
                    .LineWidthPixels = 2
                    .CenterYAxis = True
                    .ShowSmoothedData = True
                    .ShowPeakEdges = True
                End With
        
                With .Graph3DOptions
                    .ContourLevelsCount = 10
                    .Perspective = 2.5
                    .Elevation = 30
                    .YRotation = 0
                    .ZRotation = 225
                    .AnnotationFontSize = 80
                End With
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "NoiseRemovalOptions" Then
            With .NoiseRemovalOptions
                .SearchTolerancePPMDefault = 3
                .SearchTolerancePPMAutoRemoval = 3
                
                .PercentageThresholdToExcludeSlice = 25
                .PercentageThresholdToAddNeighborToSearchSlice = 15
                
                .LimitScanRange = False
                .ScanStart = 1
                .ScanEnd = 1500
                
                .LimitMassRange = False
                .MassStart = 100
                .MassEnd = 6000
                
                .SearchScope = glScope.glSc_All
                .RequireIdenticalCharge = True
                
                .ExclusionListCount = 0
                ReDim .ExclusionList(0)
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "RefineMSDataOptions" Then
            With .RefineMSDataOptions
                .MinimumPeakHeight = 10                                 ' counts/bin
                .MinimumSignalToNoiseRatioForLowAbundancePeaks = 2.5    ' signal to noise ratio; only applies to peaks with intensity <= .MinimumPeakHeight
                .PercentageOfMaxForFindingWidth = 60        ' percentage of maximum; Take the Maximum times this value to find the intensity at which to estimate the peak width; only used when .ToleranceRefinementMethod = mtrMassErrorPlotWidthAtPctOfMax
                .MassCalibrationMaximumShift = 15           ' ppm
                .MassCalibrationTolType = gltPPM
                
                .ToleranceRefinementMethod = mtrExpectationMaximization
                .UseMinMaxIfOutOfRange = True               ' If True, then uses MassToleranceMinimum or MassToleranceMaximum if the new tolerance defined is out-of-range
                
                .MassToleranceMinimum = 0.75                ' ppm
                .MassToleranceMaximum = 15                  ' ppm
                .MassToleranceAdjustmentMultiplier = 1      ' Whatever tolerance adjustment is determined using .ToleranceRefinementMethod is multiplied by this value
                
                .NETToleranceMinimum = 0.0075               ' NET
                .NETToleranceMaximum = 0.2                  ' NET
                .NETToleranceAdjustmentMultiplier = 1
                .IncludeInternalStdMatches = True
                .UseUMCClassStats = True
                .MinimumSLiC = 0
                .MaximumAbundance = 0
                
                .EMMassTolRefineForceUseSingleDataPointErrors = True
                .EMNETTolRefineForceUseSingleDataPointErrors = True
                .EMMassErrorPeakToleranceEstimatePPM = 6
                .EMNETErrorPeakToleranceEstimate = 0.05
                .EMIterationCount = 32
                .EMPercentOfDataToExclude = 10
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "TICAndBPIPlottingOptions" Then
            With .TICAndBPIPlottingOptions
                .PlotNETOnXAxis = False
                .NormalizeYAxis = True
                .SmoothUsingMovingAverage = True
                .MovingAverageWindowWidth = 3              ' Points; should be an odd number
                .TimeDomainDataMaxValue = 70000
                With .Graph2DOptions
                    .ShowPointSymbols = False
                    .DrawLinesBetweenPoints = True
                    .ShowGridLines = True
                    .AutoScaleXAxis = True
                    .PointSizePixels = 5
                    .PointShape = OlectraChart2D.ShapeConstants.oc2dShapeDot
                    .PointAndLineColor = vbBlue
                    .LineWidthPixels = 1
                    ' Note: .CenterYAxis is not used with TIC and BPI plotting
                End With
                .PointShapeSeries2 = OlectraChart2D.ShapeConstants.oc2dShapeDot
                .PointAndLineColorSeries2 = RGB(192, 0, 0)
                
                .ClipOutliers = False
                .ClipOutliersFactor = 10
                
                .KeepWindowOnTop = False
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "PairBrowserOptions" Then
            With .PairBrowserPlottingOptions
                .SortOrder = 4          ' epsER = 4
                .SortDescending = False
                .AutoZoom2DPlot = True
                .HighlightMembers = True
                .PlotAllChargeStates = False
                
                .FixedDimensionsForAutoZoom = True
                .MassRangeZoom = 6
                .MassRangeUnits = 0         ' mruDa = 0
                .ScanRangeZoom = 50
                .ScanRangeUnits = 0         ' sruScan = 0
                With .Graph2DOptions
                    .ShowPointSymbols = True
                    .DrawLinesBetweenPoints = True
                    .ShowGridLines = False
                    .PointSizePixels = 5
                    .PointShape = OlectraChart2D.ShapeConstants.oc2dShapeDot
                    .PointAndLineColor = vbBlue
                    .LineWidthPixels = 1
                    ' Note: .AutoScaleXAxis and .CenterYAxis are not used with the Pair Browser
                End With
                .PointShapeHeavy = OlectraChart2D.ShapeConstants.oc2dShapeDot
                .PointAndLineColorHeavy = vbRed
                .KeepWindowOnTop = True
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "UMCBrowserOptions" Then
            With .UMCBrowserPlottingOptions
                .SortOrder = 4          ' eusAbundance = 3
                .SortDescending = True
                .AutoZoom2DPlot = True
                .HighlightMembers = True
                .PlotAllChargeStates = False
                
                .FixedDimensionsForAutoZoom = True
                .MassRangeZoom = 6
                .MassRangeUnits = 0         ' mruDa = 0
                .ScanRangeZoom = 50
                .ScanRangeUnits = 0         ' sruScan = 0
                With .Graph2DOptions
                    .ShowPointSymbols = True
                    .DrawLinesBetweenPoints = True
                    .ShowGridLines = False
                    .PointSizePixels = 5
                    .PointShape = OlectraChart2D.ShapeConstants.oc2dShapeDot
                    .PointAndLineColor = vbBlue
                    .LineWidthPixels = 1
                    ' Note: .AutoScaleXAxis and .CenterYAxis are not used with the Pair Browser
                End With
                .KeepWindowOnTop = True
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "PairSearchOptions" Then
            With .PairSearchOptions
                
                .PairSearchMode = AUTO_FIND_PAIRS_NONE
                
                With .SearchDef
                    .DeltaMass = glN14N15_DELTA
                    .DeltaMassTolerance = 0.02
                    .AutoCalculateDeltaMinMaxCount = False
                    .DeltaCountMin = 1
                    .DeltaCountMax = 100
                    .DeltaStepSize = 1
                    .LightLabelMass = glICAT_Light
                    .HeavyLightMassDifference = Round(glICAT_Heavy - glICAT_Light, 3)
                    .LabelCountMin = 1
                    .LabelCountMax = 5
                    .MaxDifferenceInNumberOfLightHeavyLabels = 1
                
                    .RequireUMCOverlap = True
                    .RequireUMCOverlapAtApex = True
                    
                    .ScanTolerance = 15
                    .ScanToleranceAtApex = 15
                    
                    .ERInclusionMin = -5
                    .ERInclusionMax = 5
                    
                    .RequireMatchingChargeStatesForPairMembers = True
                    .UseIdenticalChargesForER = True
                    .ComputeERScanByScan = True
                    .AverageERsAllChargeStates = True
                    .AverageERsWeightingMode = aewAbundance
                    
                    .ERCalcType = ectER_RAT
                
                    With .IReportEROptions
                        .Enabled = True
                        .NaturalAbundanceRatio2Coeff.Exponent = 1.9241
                        .NaturalAbundanceRatio2Coeff.Multiplier = 0.0000003                 ' 3e-7
                        .NaturalAbundanceRatio4Coeff.Exponent = 3.2684
                        .NaturalAbundanceRatio4Coeff.Multiplier = 0.000000000002            ' 2E-12
                        .MinimumFractionScansWithValidER = 0.5
                    End With
                
                    .RemoveOutlierERs = True
                    .RemoveOutlierERsIterate = True
                    .RemoveOutlierERsMinimumDataPointCount = 3
                    .RemoveOutlierERsConfidenceLevel = 0
                
                End With
                
                .AutoExcludeOutOfERRange = False
                .AutoExcludeAmbiguous = False
                .KeepMostConfidentAmbiguous = True
                
                .AutoAnalysisRemovePairMemberHitsAfterDBSearch = False
                .AutoAnalysisRemovePairMemberHitsRemoveHeavy = True
                
                .AutoAnalysisSavePairsToTextFile = True
                .AutoAnalysisSavePairsStatisticsToTextFile = True
                
                ' Note: If you change the following default, you should also change the setting in Sub DisplayDefaultSettings() of frmSearchForNETAdjustmentUMC
                .NETAdjustmentPairedSearchUMCSelection = punaUnpairedPlusPairedLight
                
                .OutlierRemovalUsesSymmetricERs = True
                
                .AutoAnalysisDeltaMassAddnlCount = 0
                ReDim .AutoAnalysisDeltaMassAddnl(0)
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "MassTagStalenessOptions" Then
            With .MassTagStalenessOptions
                .MaximumAgeLoadedMassTagsHours = 8          ' 8 hours
                .MaximumFractionAMTsWithNulls = 0.1         ' 10%
                .MaximumCountAMTsWithNulls = 2500           ' 2500 AMTs
                .MinimumTimeBetweenReloadMinutes = 30       ' 30 minutes
                
                ' Do not reset the other values here since MT tags could already be in memory
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "SLiCScoreOptions" Then
            With .SLiCScoreOptions
                .MassPPMStDev = 3                       ' 3 ppm
                .NETStDev = 0.025                       ' 0.025 NET
                .UseAMTNETStDev = False                 ' Add in the NET StDev value for each AMT; December 2005: This value is now ignored, and essentially defaults to False
                .MaxSearchDistanceMultiplier = 2        ' Minimum search distance is 2 * MassPPMStDev * STDEV_SCALING_FACTOR or 2 * MWTol, whichever is larger
                .AutoDefineSLiCScoreThresholds = True
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "GraphicExportOptions" Then
            With .GraphicExportOptions
                .CopyEMFIncludeFilenameAndDate = True
                .CopyEMFIncludeTextLabels = True
                
                SetEditCopyEMFOptions .CopyEMFIncludeFilenameAndDate, .CopyEMFIncludeTextLabels
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "AutoAnalysisFilterPrefs" Then
            With .AutoAnalysisFilterPrefs
                .ExcludeDuplicates = False
                .ExcludeDuplicatesTolerance = 2     ' equivalent to GelPrefs.DupTolerance
                
                .ExcludeIsoByFit = blnApplyFilterOnIsotopicFit
                .ExcludeIsoByFitMaxVal = 0.15       ' equivalent to GelPrefs.IsoDataFit
                
                .ExcludeIsoSecondGuess = False
                .ExcludeIsoLessLikelyGuess = False
                
                .ExcludeCSByStdDev = False
                .ExcludeCSByStdDevMaxVal = 1
                
                .RestrictIsoByAbundance = False
                .RestrictIsoAbundanceMin = 0
                .RestrictIsoAbundanceMax = 1E+15
                
                .RestrictIsoByMass = True
                .RestrictIsoMassMin = 400
                .RestrictIsoMassMax = 6000
                
                .RestrictIsoByMZ = True
                .RestrictIsoMZMin = 400
                .RestrictIsoMZMax = 3000
                
                .RestrictIsoByChargeState = True
                .RestrictIsoChargeStateMin = 1
                .RestrictIsoChargeStateMax = 6
                
                .RestrictCSByAbundance = False
                .RestrictCSAbundanceMin = 0
                .RestrictCSAbundanceMax = 1E+15
                
                .RestrictCSByMass = True
                .RestrictCSMassMin = 400
                .RestrictCSMassMax = 6000
                
                .RestrictScanRange = False
                .RestrictScanRangeMin = 0
                .RestrictScanRangeMax = 5000
                
                .RestrictGANETRange = False
                .RestrictGANETRangeMin = -1
                .RestrictGANETRangeMax = 2
                
                .RestrictToEvenScanNumbersOnly = False
                .RestrictToOddScanNumbersOnly = False
                
                .MaximumDataCountEnabled = True
                .MaximumDataCountToLoad = DEFAULT_MAXIMUM_DATA_COUNT_TO_LOAD
            End With
        End If
        
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "AutoAnalysisDBInfo" Then
            FillGelAnalysisInfo .AutoAnalysisDBInfo
        End If
                
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "AutoAnalysisOptions" Then
            With .AutoAnalysisOptions
                .DatasetID = 0
                .JobNumber = 0
                .MDType = 1
                .AutoRemoveNoiseStreaks = False
                .DoNotSaveOrExport = False
                
                .SkipFindUMCs = False
                .SkipGANETSlopeAndInterceptComputation = False
                .DBConnectionRetryAttemptMax = 5
                .DBConnectionTimeoutSeconds = 300
                .ExportResultsFileUsesJobNumberInsteadOfDataSetName = True
                
                .SaveGelFile = False
                .SaveGelFileOnError = True
                .SavePictureGraphic = True
                .SavePictureGraphicFileType = pftPictureFileTypeConstants.pftPNG
                .SavePictureWidthPixels = 1024
                .SavePictureHeightPixels = 768
                .SaveInternalStdHitsAndData = False
                
                .SaveErrorGraphicMass = True
                .SaveErrorGraphicGANET = True
                .SaveErrorGraphic3D = True
                .SaveErrorGraphicFileType = pftPictureFileTypeConstants.pftPNG
                .SaveErrorGraphSizeWidthPixels = 800
                .SaveErrorGraphSizeHeightPixels = 600
                
                .SavePlotTIC = True
                .SavePlotBPI = True
                .SavePlotTICTimeDomain = True
                .SavePlotTICDataPointCounts = True
                .SavePlotTICDataPointCountsHitsOnly = True
                .SavePlotTICFromRawData = True
                .SavePlotBPIFromRawData = True
                .SavePlotDeisotopingIntensityThresholds = True
                .SavePlotDeisotopingPeakCounts = True
                    
                ' Note: The NetAdjustment options are located later in this sub
                
                .UMCShrinkingBoxWeightAverageMassByIntensity = False        ' Only applies to UMC2003
                .UMCSearchMode = AUTO_ANALYSIS_UMCIonNet
                
                .OutputFileSeparationCharacter = SEPARATION_CHARACTER_TAB_STRING            ' Gets converted to vbTab
                .PEKFileExtensionPreferenceOrder = DEFAULT_PEK_FILE_EXTENSION_ORDER
                .WriteIDResultsByIonToTextFileAfterAutoSearches = False
                .SaveUMCStatisticsToTextFile = False
                
                .IncludeORFNameInTextFileOutput = True
                .SetIsConfirmedForDBSearchMatches = True
                .AddQuantitationDescriptionEntry = True
                .ExportUMCsWithNoMatches = False
                .DBSearchRegionShape = srsElliptical
                .UseLegacyDBForMTs = APP_BUILD_DISABLE_MTS
                .IgnoreNETAdjustmentFailure = True
                
                With .AutoToleranceRefinement
                    .DBSearchMWTol = DEFAULT_TOLERANCE_REFINEMENT_MW_TOL
                    .DBSearchTolType = gltPPM
                    .DBSearchNETTol = DEFAULT_TOLERANCE_REFINEMENT_NET_TOL
                    .DBSearchRegionShape = srsRectangular
                    .DBSearchMinimumHighNormalizedScore = 0
                    .DBSearchMinimumHighDiscriminantScore = 0.5
                    .DBSearchMinimumPeptideProphetProbability = 0.5
                    .RefineMassCalibration = True
                    .RefineMassCalibrationOverridePPM = 0
                    .RefineDBSearchMassTolerance = True
                    .RefineDBSearchNETTolerance = True
                End With
                
                Erase .AutoAnalysisSearchMode()
                
                .AutoAnalysisSearchModeCount = 1
                ResetAutoSearchModeEntry .AutoAnalysisSearchMode(0)
                
                With .AutoAnalysisSearchMode(0)
                    .SearchMode = AUTO_SEARCH_UMC_CONGLOMERATE
                    .WriteResultsToTextFile = True
                End With
            
            End With
        End If
    
        If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "AutoAnalysisOptions" Or strSingleSectionToReset = "NetAdjustmentOptions" Then
            With .AutoAnalysisOptions
                .NETAdjustmentInitialNetTol = 0.2
                ' November 2005: Unused variable     .NETAdjustmentFinalNetTol = 0.01
                .NETAdjustmentMaxIterationCount = 50
                .NETAdjustmentMinIDCount = 75
                .NETAdjustmentMinIDCountAbsoluteMinimum = 10
                .NETAdjustmentMinIterationCount = 5
                .NETAdjustmentChangeThresholdStopValue = 0.0005
                
                .NETAdjustmentAutoIncrementUMCTopAbuPct = True
                .NETAdjustmentUMCTopAbuPctInitial = 20              ' Percent (integer between 0 and 100)
                .NETAdjustmentUMCTopAbuPctIncrement = 20            ' Percent (integer between 1 and 100)
                .NETAdjustmentUMCTopAbuPctMax = 100                 ' Percent (integer between 1 and 100, >= NETAdjustmentUMCTopAbuPctInitial)
                
                ' November 2005: Unused variable     .NETAdjustmentMinimumNETMatchScore = 50
                
                .NETSlopeExpectedMinimum = 0.00001
                .NETSlopeExpectedMaximum = 0.01
                .NETInterceptExpectedMinimum = -1
                .NETInterceptExpectedMaximum = 1
            End With
        End If
    
        If Not APP_BUILD_DISABLE_MTS Then
            If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "DMSConnectionInfo" Then
                With .DMSConnectionInfo
                    .ConnectionString = "Provider=sqloledb;Data Source=gigasax;Initial Catalog=DMS5;User ID=DMSReader;Password=dms4fun"
                End With
            End If
            
            
            If Len(strSingleSectionToReset) = 0 Or strSingleSectionToReset = "MTSConnectionInfo" Then
                With .MTSConnectionInfo
                    .ConnectionString = "Provider=sqloledb;Data Source=pogo;Initial Catalog=MTS_Master;User ID=MTUser;Password=mt4fun"
                    
                    .spAddQuantitationDescription = "AddQuantitationDescription"
                    .spGetLockers = "GetLockers_02_2002"
                    .spGetMassTagMatchCount = "GetMassTagMatchCount"
                    .spGetMassTags = "GetMassTagsGANETParam"
                    .spGetMassTagsSubset = "GetMassTagsForSubset"
                    .spGetPMResultStats = "GetPeakMatchingTaskResultStats"
                    .spPutAnalysis = "AddMatchMaking"
                    
                    ' September 2004: Unused variable
                    '' .spPutPeak = "AddFTICRPeak"
                    
                    .spPutUMC = "AddFTICRUmc"
                    .spPutUMCMember = "AddFTICRUmcMember"
                    .spPutUMCMatch = "AddFTICRUmcMatch"
                    .spPutUMCInternalStdMatch = "AddFTICRUmcInternalStdMatch"
                    .spEditGANET = "EditFAD_GANET"
                    .spGetORFs = "srvGetORFs"
                    .spGetORFSeq = "srvGetORFSequenceForORFID"
                    .spGetORFIDs = "srvGetORFIDs"
                    .spGetORFRecord = "srvGetORFRecord"
                    .spGetMassTagSeq = "srvGetMassTagSequence"
                    .spGetMassTagNames = "srvGetMassTagNames"
                    .spGetInternalStandards = "GetInternalStandards"
                    .spGetDBSchemaVersion = "GetDBSchemaVersion"
                    .spGetMassTagToProteinNameMap = "GetMassTagToProteinNameMap"
                    
                    .sqlGetMTNames = "SELECT * FROM V_IFC_Mass_Tag_To_Protein_Map ORDER BY 1"               ' We really only want the first 3 columns, but they have different names in the old and new schema, and we thus will simply access the view
                    ' Obsolete: .sqlGetORFMassTagMap = "SELECT * FROM V_IFC_Mass_Tag_To_Protein_Name_Map ORDER BY 1"
                End With
            End If
        End If
    End With
    
End Sub

Public Sub ResetUMCAdvancedStatsOptions(udtAdvancedStatsOptions As udtUMCAdvancedStatsOptionsType)
    With udtAdvancedStatsOptions
        .ClassAbuTopXMinAbu = 0
        .ClassAbuTopXMaxAbu = 0             ' If 0, then no maximum abundance
        .ClassAbuTopXMinMembers = 3         ' If TopXMinAbu < 0 and TopXMaxAbu are < 0, then maximum number of members to include
                                            ' Otherwise, minimum members to be included if not enough found in range
        
        .ClassMassTopXMinAbu = 0
        .ClassMassTopXMaxAbu = 0            ' If 0, then no maximum abundance
        .ClassMassTopXMinMembers = 3        ' If TopXMinAbu < 0 and TopXMaxAbu are < 0, then maximum number of members to include
                                            ' Otherwise, minimum members to be included if not enough found in range
    End With
End Sub

Public Sub ResetUMCAutoRefineOptions(udtAutoRefineOptions As udtUMCAutoRefineOptionsType)
    
    With udtAutoRefineOptions
        .UMCAutoRefineRemoveCountLow = True
        .UMCAutoRefineMinLength = 3
        
        .UMCAutoRefineRemoveCountHigh = False
        .UMCAutoRefineMaxLength = 400
        
        .UMCAutoRefineRemoveMaxLengthPctAllScans = True
        .UMCAutoRefineMaxLengthPctAllScans = 15
        
        .UMCAutoRefinePercentMaxAbuToUseForLength = 33
        
        .TestLengthUsingScanRange = True
        .MinMemberCountWhenUsingScanRange = 3
        
        .UMCAutoRefineRemoveAbundanceLow = False
        .UMCAutoRefineRemoveAbundanceHigh = False
        .UMCAutoRefinePctLowAbundance = 30#
        .UMCAutoRefinePctHighAbundance = 30#
        
        .SplitUMCsByAbundance = True
        With .SplitUMCOptions
            .MinimumDifferenceInAveragePpmMassToSplit = 4
            .StdDevMultiplierForSplitting = 1
            .MaximumPeakCountToSplitUMC = 6
            .PeakDetectIntensityThresholdPercentageOfMaximum = 15
            .PeakDetectIntensityThresholdAbsoluteMinimum = 0
            .PeakWidthPointsMinimum = 4
            .PeakWidthInSigma = 3
            .ScanGapBehavior = susgSplitIfMassDifference
        End With
    End With

End Sub

' No longer supported (March 2006)
''Private Sub ResolveFTICR_AMTPrefs()
''Dim aRes As Variant
''On Error GoTo err_ResolveFTICR_AMT
''aRes = ResolvePrefsString(sFTICR_AMTPref)
''sFTICR_AMTPath = aRes(1)
''Exit Sub
''
''err_ResolveFTICR_AMT:
''ResetFTICR_AMTPreferences
''End Sub
''
''Private Function GetFTICR_AMTPrefs() As String
''GetFTICR_AMTPrefs = sFTICR_AMTPath
''End Function
''
''Private Sub ResetFTICR_AMTPreferences()
''sFTICR_AMTPath = ""
''End Sub

Public Function SelectLegacyMTDB(objCallingForm As Form, strCurrentFilePath As String) As String
    Dim fso As New FileSystemObject
    
    Dim strParentFolder As String
    Dim strFileName As String
    Dim strNewPath As String
    
On Error GoTo SelectLegacyAMTDBErrorHandler

    If Len(strCurrentFilePath) > 0 Then
        strParentFolder = fso.GetParentFolderName(strCurrentFilePath)
        strFileName = fso.GetFileName(strCurrentFilePath)
    End If
    
    strNewPath = SelectFile(objCallingForm.hwnd, "Select legacy MT database", strParentFolder, False, strFileName, "Access DB files (*.mdb)|*.mdb|All Files (*.*)|*.*", 1)

    If Not fso.FileExists(strNewPath) Then
        strNewPath = strCurrentFilePath
    End If
    
    SelectLegacyMTDB = strNewPath
    Exit Function

SelectLegacyAMTDBErrorHandler:
    Debug.Print "Error in SelectLegacyMTDB: " & Err.Description
    Debug.Assert False
    
    SelectLegacyMTDB = strNewPath
    
End Function


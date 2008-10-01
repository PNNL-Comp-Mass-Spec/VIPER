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

Private Const RECENT_DB_CONNECTIONS_MAX_COUNT As Integer = 25
Private Const RECENT_DB_CONNECTIONS_SECTION_NAME As String = "RecentDBConnections"
Private Const RECENT_DB_CONNECTIONS_KEY_COUNT_NAME As String = "ConnectionCount"
Private Const RECENT_DB_CONNECTION_SUBSECTION_NAME As String = "Connection"
Private Const RECENT_DB_CONNECTION_INFOVERSION_NAME As String = "InfoVersion"
Private Const RECENT_DB_CONNECTION_INFOVERSION As Integer = 2

Private Const NET_ADJ_SECTION_OLDNAME As String = "UMCNetDef"
Private Const NET_ADJ_SECTION_NEWNAME As String = "UMCNETAdjDef"
Private Const NET_ADJ_MS_WARP_SECTION As String = "UMCNETAdjMSWarpDef"

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

Private Sub AddKeyValueSetting(ByRef strKeys() As String, ByRef strValues() As String, ByRef intKeyValueCount As Integer, ByRef strKey As String, ByRef strValue As String, Optional ByVal blnResetList As Boolean)
    If blnResetList Then
        intKeyValueCount = 0
    End If
    
    If intKeyValueCount >= UBound(strKeys) Then
        ' Expand strKeys() & strValues()
        ReDim Preserve strKeys((UBound(strKeys) + 1) * 2 - 1)
        ReDim Preserve strValues(UBound(strKeys))
    End If
    
    strKeys(intKeyValueCount) = strKey
    strValues(intKeyValueCount) = strValue
    intKeyValueCount = intKeyValueCount + 1
End Sub

Private Sub AddKeyValueSettingBln(ByRef strKeys() As String, ByRef strValues() As String, ByRef intKeyValueCount As Integer, ByRef strKey As String, ByRef blnValue As Boolean, Optional ByVal blnResetList As Boolean)
    AddKeyValueSetting strKeys, strValues, intKeyValueCount, strKey, Trim(Str(blnValue)), blnResetList
End Sub

Private Sub AddKeyValueSettingByt(ByRef strKeys() As String, ByRef strValues() As String, ByRef intKeyValueCount As Integer, ByRef strKey As String, ByRef bytValue As Byte, Optional ByVal blnResetList As Boolean)
    AddKeyValueSetting strKeys, strValues, intKeyValueCount, strKey, Trim(Str(bytValue)), blnResetList
End Sub

Private Sub AddKeyValueSettingSng(ByRef strKeys() As String, ByRef strValues() As String, ByRef intKeyValueCount As Integer, ByRef strKey As String, ByRef sngValue As Single, Optional ByVal blnResetList As Boolean)
    AddKeyValueSetting strKeys, strValues, intKeyValueCount, strKey, Trim(Str(sngValue)), blnResetList
End Sub

Private Sub AddKeyValueSettingDbl(ByRef strKeys() As String, ByRef strValues() As String, ByRef intKeyValueCount As Integer, ByRef strKey As String, ByRef dblValue As Double, Optional ByVal blnResetList As Boolean)
    AddKeyValueSetting strKeys, strValues, intKeyValueCount, strKey, Trim(Str(dblValue)), blnResetList
End Sub

Private Sub AddKeyValueSettingInt(ByRef strKeys() As String, ByRef strValues() As String, ByRef intKeyValueCount As Integer, ByRef strKey As String, ByRef intValue As Integer, Optional ByVal blnResetList As Boolean)
    AddKeyValueSetting strKeys, strValues, intKeyValueCount, strKey, Trim(Str(intValue)), blnResetList
End Sub

Private Sub AddKeyValueSettingLng(ByRef strKeys() As String, ByRef strValues() As String, ByRef intKeyValueCount As Integer, ByRef strKey As String, ByRef lngValue As Long, Optional ByVal blnResetList As Boolean)
    AddKeyValueSetting strKeys, strValues, intKeyValueCount, strKey, Trim(Str(lngValue)), blnResetList
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

Private Function GetIniFileSettingByt(objIniStuff As clsIniStuff, strSection As String, strKey As String, Optional bytDefault As Byte = 0) As Byte
    Dim strSetting As String
    Dim strDefault As String
    Dim intValue As Integer
    
    strDefault = Str(bytDefault)

    If objIniStuff.ReadValue(strSection, strKey, strDefault, strSetting) Then
        intValue = CIntSafe(Trim(strSetting))
        If intValue <= 0 Then
            GetIniFileSettingByt = 0
        ElseIf intValue >= 255 Then
            GetIniFileSettingByt = 255
        Else
            GetIniFileSettingByt = CByte(intValue)
        End If
    Else
        GetIniFileSettingByt = bytDefault
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
    
    Dim strModValue As String
    
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
        .OddEvenProcessingMode = GetIniFileSettingInt(IniStuff, "UMCDef", "OddEvenProcessingMode", .OddEvenProcessingMode)
        .RequireMatchingIsotopeTag = GetIniFileSettingBln(IniStuff, "UMCDef", "RequireMatchingIsotopeTag", .RequireMatchingIsotopeTag)
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
        IniStuff.WriteSection strUMCNetAdjDefSectionName, strKeys(), strValues(), 0
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
        
        .IReportAutoAddMonoPlus4AndMinus4Data = GetIniFileSettingBln(IniStuff, "ExpandedPreferences", "IReportAutoAddMonoPlus4AndMinus4Data", .IReportAutoAddMonoPlus4AndMinus4Data)
        
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
    
        .ComputePairwiseMassDifferences = GetIniFileSettingBln(IniStuff, "RefineMSDataOptions", "ComputePairwiseMassDifferences", .ComputePairwiseMassDifferences)
        .PairwiseMassDiffMinimum = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "PairwiseMassDiffMinimum", .PairwiseMassDiffMinimum)
        .PairwiseMassDiffMaximum = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "PairwiseMassDiffMaximum", .PairwiseMassDiffMaximum)
        .PairwiseMassBinSize = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "PairwiseMassBinSize", .PairwiseMassBinSize)
        .PairwiseMassDiffNETTolerance = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "PairwiseMassDiffNETTolerance", .PairwiseMassDiffNETTolerance)
        .PairwiseMassDiffNETOffset = GetIniFileSettingSng(IniStuff, "RefineMSDataOptions", "PairwiseMassDiffNETOffset", .PairwiseMassDiffNETOffset)
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
            .DeltaMassTolType = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "DeltaMassTolType", .DeltaMassTolType)
            
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
            .ScanByScanAverageIsNotWeighted = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "ScanByScanAverageIsNotWeighted", .ScanByScanAverageIsNotWeighted)
            
            .RequireMatchingIsotopeTagLabels = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RequireMatchingIsotopeTagLabels", .RequireMatchingIsotopeTagLabels)
            
            .MonoPlusMinusThresholdForceHeavyOrLight = GetIniFileSettingByt(IniStuff, "PairSearchOptions", "MonoPlusMinusThresholdForceHeavyOrLight", .MonoPlusMinusThresholdForceHeavyOrLight)
            .IgnoreMonoPlus2AbundanceInIReportERCalc = GetIniFileSettingByt(IniStuff, "PairSearchOptions", "IgnoreMonoPlus2AbundanceInIReportERCalc", .IgnoreMonoPlus2AbundanceInIReportERCalc)
            
            .AverageERsAllChargeStates = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "AverageERsAllChargeStates", .AverageERsAllChargeStates)
            .AverageERsWeightingMode = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "AverageERsWeightingMode", CInt(.AverageERsWeightingMode))
            .ERCalcType = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "ERCalcType", CInt(.ERCalcType))
            
            .RemoveOutlierERs = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RemoveOutlierERs", .RemoveOutlierERs)
            .RemoveOutlierERsIterate = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "RemoveOutlierERsIterate", .RemoveOutlierERsIterate)
            .RemoveOutlierERsMinimumDataPointCount = GetIniFileSettingLng(IniStuff, "PairSearchOptions", "RemoveOutlierERsMinimumDataPointCount", .RemoveOutlierERsMinimumDataPointCount)
            .RemoveOutlierERsConfidenceLevel = GetIniFileSettingInt(IniStuff, "PairSearchOptions", "RemoveOutlierERsConfidenceLevel", .RemoveOutlierERsConfidenceLevel)
        
            .N15IncompleteIncorporationMode = GetIniFileSettingBln(IniStuff, "PairSearchOptions", "N15IncompleteIncorporationMode", .N15IncompleteIncorporationMode)
            .N15PercentIncorporationMinimum = GetIniFileSettingSng(IniStuff, "PairSearchOptions", "N15PercentIncorporationMinimum", .N15PercentIncorporationMinimum)
            .N15PercentIncorporationMaximum = GetIniFileSettingSng(IniStuff, "PairSearchOptions", "N15PercentIncorporationMaximum", .N15PercentIncorporationMaximum)
            .N15PercentIncorporationStep = GetIniFileSettingSng(IniStuff, "PairSearchOptions", "N15PercentIncorporationStep", .N15PercentIncorporationStep)
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
            
            .GenerateMonoPlus4IsoLabelingFile = GetIniFileSettingBln(IniStuff, "AutoAnalysisOptions", "GenerateMonoPlus4IsoLabelingFile", .GenerateMonoPlus4IsoLabelingFile)
            
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
                    ' "DynamicMods" was replaced with "ModMode" in August 2008
                    ' Preferentially use ModMode, if present
                    .ModMode = IniFileLoadSettingsGetModMode(IniStuff, strSectionName, .ModMode)
                                     
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
                            ' "DynamicMods" was replaced with "ModMode" in August 2008
                            ' Preferentially use ModMode, if present
                            .ModMode = IniFileLoadSettingsGetModMode(IniStuff, strSectionName, .ModMode)
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
            If .ExcludeIsoByFitMaxVal > 100 Then
                .ExcludeIsoByFitMaxVal = 100
            End If
            
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

Private Function IniFileLoadSettingsGetModMode(ByRef IniStuff As clsIniStuff, ByVal strSectionName As String, bytDefaultIfMissing As Byte) As Byte
                    
    Dim strModValue As String
    Dim bytNewModMode As Byte
    
    bytNewModMode = bytDefaultIfMissing
 
    strModValue = GetIniFileSetting(IniStuff, strSectionName, "ModMode", "")
    If Len(strModValue) > 0 Then
        ' ModMode value is defined; use it
        If IsNumeric(strModValue) Then
            bytNewModMode = CByte(strModValue)
        End If
    Else
        
        strModValue = GetIniFileSetting(IniStuff, strSectionName, "DynamicMods", "")
        If Len(strModValue) > 0 Then
            ' Legacy DynamicMods value was present; use it
            If CBoolSafe(strModValue) Then
                bytNewModMode = 1
            Else
                bytNewModMode = 0
            End If
        Else
            ' Could not find "DynamicMods" or "ModMode" in this section
            ' Leave .ModMode unchanged
        End If
    End If
    
    IniFileLoadSettingsGetModMode = bytNewModMode
End Function

Public Sub IniFileSaveSettings(udtPrefsExpanded As udtPreferencesExpandedType, udtUMCDef As UMCDefinition, udtUMCIonNetDef As UMCIonNetDefinition, udtUMCNetAdjDef As NetAdjDefinition, udtInternalStandards As udtInternalStandardsType, udtAMTDef As SearchAMTDefinition, udtPrefs As GelPrefs, Optional strIniFilePath As String = "", Optional bnlAutoAnalysisFieldsOnly As Boolean = False)
    ' Saves settings to an .ini file
    ' When bnlAutoAnalysisFieldsOnly = True, then skips the settings that are not needed for Auto Analysis
    
    Dim blnSuccess As Boolean
    Dim IniStuff As New clsIniStuff
    Dim DBIniStuff As clsIniStuff
    Dim strDBIniFilePath As String
    Dim intIndex As Integer, intAutoSearchModeIndex As Integer
    
    Dim iKVCount As Integer
    Dim sKeys() As String, sVals() As String
    
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
    
    ReDim sKeys(99)
    ReDim sVals(99)
    
    ' UMC options stored in udtPrefsExpanded.AutoAnalysisOptions
    With udtPrefsExpanded.AutoAnalysisOptions
        iKVCount = 0
        AddKeyValueSetting sKeys, sVals, iKVCount, "UMCSearchModeList", "; Options are " & GetUMCSearchModeList()
        AddKeyValueSetting sKeys, sVals, iKVCount, "UMCSearchMode", .UMCSearchMode
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCShrinkingBoxWeightAverageMassByIntensity", .UMCShrinkingBoxWeightAverageMassByIntensity
    End With
    
    ' UMC options stored in udtUMCDef
    With udtUMCDef
        AddKeyValueSetting sKeys, sVals, iKVCount, "UMCTypeList", "; Options are " & GetUMCTypeList()
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "UMCType", .UMCType
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MWField", .MWField
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "TolType", .TolType
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "Tol", .Tol
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCSharing", .UMCSharing
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCUniCS", .UMCUniCS
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ClassAbu", .ClassAbu
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ClassMW", .ClassMW
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "GapMaxCnt", .GapMaxCnt
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "GapMaxSize", .GapMaxSize
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "GapMaxPct", .GapMaxPct
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "UMCNETType", .UMCNETType
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "InterpolateGaps", .InterpolateGaps
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "InterpolateMaxGapSize", .InterpolateMaxGapSize
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "InterpolationType", .InterpolationType
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ChargeStateStatsRepType", .ChargeStateStatsRepType
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCClassStatsUseStatsFromMostAbuChargeState", .UMCClassStatsUseStatsFromMostAbuChargeState
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "OddEvenProcessingMode", .OddEvenProcessingMode
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RequireMatchingIsotopeTag", .RequireMatchingIsotopeTag
    End With
    
    ' UMC options stored in udtPrefsExpanded
    With udtPrefsExpanded.UMCAutoRefineOptions
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCAutoRefineRemoveCountLow", .UMCAutoRefineRemoveCountLow
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCAutoRefineRemoveCountHigh", .UMCAutoRefineRemoveCountHigh
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCAutoRefineRemoveMaxLengthPctAllScans", .UMCAutoRefineRemoveMaxLengthPctAllScans
        
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "UMCAutoRefineMinLength", .UMCAutoRefineMinLength
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "UMCAutoRefineMaxLength", .UMCAutoRefineMaxLength
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "UMCAutoRefineMaxLengthPctAllScans", .UMCAutoRefineMaxLengthPctAllScans
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "UMCAutoRefinePercentMaxAbuToUseForLength", .UMCAutoRefinePercentMaxAbuToUseForLength
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "TestLengthUsingScanRange", .TestLengthUsingScanRange
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MinMemberCountWhenUsingScanRange", .MinMemberCountWhenUsingScanRange
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCAutoRefineRemoveAbundanceLow", .UMCAutoRefineRemoveAbundanceLow
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UMCAutoRefineRemoveAbundanceHigh", .UMCAutoRefineRemoveAbundanceHigh
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "UMCAutoRefinePctLowAbundance", .UMCAutoRefinePctLowAbundance
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "UMCAutoRefinePctHighAbundance", .UMCAutoRefinePctHighAbundance
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SplitUMCsByAbundance", .SplitUMCsByAbundance
        With .SplitUMCOptions
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MinimumDifferenceInAveragePpmMassToSplit", .MinimumDifferenceInAveragePpmMassToSplit
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "StdDevMultiplierForSplitting", .StdDevMultiplierForSplitting
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "MaximumPeakCountToSplitUMC", .MaximumPeakCountToSplitUMC
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PeakDetectIntensityThresholdPercentageOfMaximum", .PeakDetectIntensityThresholdPercentageOfMaximum
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "PeakDetectIntensityThresholdAbsoluteMinimum", .PeakDetectIntensityThresholdAbsoluteMinimum
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PeakWidthPointsMinimum", .PeakWidthPointsMinimum
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PeakWidthInSigma", .PeakWidthInSigma
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "ScanGapBehavior", CInt(.ScanGapBehavior)
        End With
    End With
    IniStuff.WriteSection "UMCDef", sKeys(), sVals(), iKVCount
    
    With udtUMCIonNetDef
        iKVCount = 0
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "NetDim", .NetDim
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "NetActualDim", .NetActualDim
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MetricType", .MetricType
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETType", .NETType
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "TooDistant", .TooDistant
        For intIndex = 0 To UBound(.MetricData())
            strKeyPrefix = "Dim" & Trim(intIndex + 1)
            With .MetricData(intIndex)
                AddKeyValueSettingBln sKeys, sVals, iKVCount, strKeyPrefix & "Use", .Use
                AddKeyValueSettingLng sKeys, sVals, iKVCount, strKeyPrefix & "DataType", .DataType
                AddKeyValueSettingDbl sKeys, sVals, iKVCount, strKeyPrefix & "WeightFactor", .WeightFactor
                AddKeyValueSettingLng sKeys, sVals, iKVCount, strKeyPrefix & "ConstraintType", .ConstraintType
                AddKeyValueSettingDbl sKeys, sVals, iKVCount, strKeyPrefix & "ConstraintValue", .ConstraintValue
                AddKeyValueSettingLng sKeys, sVals, iKVCount, strKeyPrefix & "ConstraintUnits", .ConstraintUnits
            End With
        Next intIndex
    End With
    
    ' UMCIso options stored in udtPrefsExpanded
    ' If .MetricData() is changed to not have 6 items, then the following + # values must be changed
    With udtPrefsExpanded.UMCIonNetOptions
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "UMCRepresentative", .UMCRepresentative
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "MakeSingleMemberClasses", .MakeSingleMemberClasses
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ConnectionLengthPostFilterMaxNET", .ConnectionLengthPostFilterMaxNET
    End With
    
    IniStuff.WriteSection "UMCIonNetDef", sKeys(), sVals(), iKVCount
    
    With udtPrefsExpanded.UMCAdvancedStatsOptions
        iKVCount = 0
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ClassAbuTopXMinAbu", .ClassAbuTopXMinAbu
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ClassAbuTopXMaxAbu", .ClassAbuTopXMaxAbu
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "ClassAbuTopXMinMembers", .ClassAbuTopXMinMembers

        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ClassMassTopXMinAbu", .ClassMassTopXMinAbu
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ClassMassTopXMaxAbu", .ClassMassTopXMaxAbu
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "ClassMassTopXMinMembers", .ClassMassTopXMinMembers
    End With
    IniStuff.WriteSection "UMCAdvancedStatsOptions", sKeys(), sVals(), iKVCount
    
    
    With udtUMCNetAdjDef
        iKVCount = 0
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MinUMCCount", .MinUMCCount
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MinScanRange", .MinScanRange
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MaxScanPct", .MaxScanPct
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "TopAbuPct", .TopAbuPct
        ' Ignored: .PeakSelection
        ' Ignored: .PeakMaxAbuPct
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MWTolType", .MWTolType
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MWTol", .MWTol
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETorRT", .NETorRT
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseNET", .UseNET
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseMultiIDMaxNETDist", .UseMultiIDMaxNETDist
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MultiIDMaxNETDist", .MultiIDMaxNETDist
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "EliminateBadNET", .EliminateBadNET
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MaxIDToUse", .MaxIDToUse
        
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "IterationStopType", .IterationStopType
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "IterationStopValue", .IterationStopValue
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "IterationUseMWDec", .IterationUseMWDec
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "IterationMWDec", .IterationMWDec
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "IterationUseNETdec", .IterationUseNETdec
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "IterationNETDec", .IterationNETDec
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "IterationAcceptLast", .IterationAcceptLast
        
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "InitialSlope", .InitialSlope
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "InitialIntercept", .InitialIntercept
        
        ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''        AddKeyValueSetting sKeys, sVals, iKVCount, "UseNetAdjLockers", .UseNetAdjLockers
''        AddKeyValueSetting sKeys, sVals, iKVCount, "UseOldNetAdjIfFailure", .UseOldNetAdjIfFailure
''        AddKeyValueSetting sKeys, sVals, iKVCount, "NetAdjLockerMinimumMatchCount", .NetAdjLockerMinimumMatchCount
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseRobustNETAdjustment", .UseRobustNETAdjustment
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "RobustNETAdjustmentMode", .RobustNETAdjustmentMode
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETSlopeStart", .RobustNETSlopeStart
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETSlopeEnd", .RobustNETSlopeEnd
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "RobustNETSlopeIncreaseMode", .RobustNETSlopeIncreaseMode
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETSlopeIncrement", .RobustNETSlopeIncrement
        
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETInterceptStart", .RobustNETInterceptStart
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETInterceptEnd", .RobustNETInterceptEnd
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETInterceptIncrement", .RobustNETInterceptIncrement
        
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETMassShiftPPMStart", .RobustNETMassShiftPPMStart
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETMassShiftPPMEnd", .RobustNETMassShiftPPMEnd
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "RobustNETMassShiftPPMIncrement", .RobustNETMassShiftPPMIncrement
    
    End With
    
    With udtPrefsExpanded
        With .AutoAnalysisOptions
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETAdjustmentInitialNetTol", .NETAdjustmentInitialNetTol
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETAdjustmentMaxIterationCount", .NETAdjustmentMaxIterationCount
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETAdjustmentMinIDCount", .NETAdjustmentMinIDCount
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETAdjustmentMinIDCountAbsoluteMinimum", .NETAdjustmentMinIDCountAbsoluteMinimum
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETAdjustmentMinIterationCount", .NETAdjustmentMinIterationCount
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETAdjustmentChangeThresholdStopValue", .NETAdjustmentChangeThresholdStopValue
            
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "NETAdjustmentAutoIncrementUMCTopAbuPct", .NETAdjustmentAutoIncrementUMCTopAbuPct
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETAdjustmentUMCTopAbuPctInitial", .NETAdjustmentUMCTopAbuPctInitial
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETAdjustmentUMCTopAbuPctIncrement", .NETAdjustmentUMCTopAbuPctIncrement
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NETAdjustmentUMCTopAbuPctMax", .NETAdjustmentUMCTopAbuPctMax
            
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETSlopeExpectedMinimum", .NETSlopeExpectedMinimum
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETSlopeExpectedMaximum", .NETSlopeExpectedMaximum
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETInterceptExpectedMinimum", .NETInterceptExpectedMinimum
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETInterceptExpectedMaximum", .NETInterceptExpectedMaximum
        End With
    End With
    IniStuff.WriteSection NET_ADJ_SECTION_NEWNAME, sKeys(), sVals(), iKVCount

    ' Write this after writing the udtUMCNETAdjDef section
    With udtUMCNetAdjDef
        For intIndex = 0 To UBound(.PeakCSSelection)
            IniStuff.WriteValue NET_ADJ_SECTION_NEWNAME, "PeakCSSelection" & Trim(intIndex), CStr(.PeakCSSelection(intIndex))
        Next intIndex
    End With

    If Not APP_BUILD_DISABLE_LCMSWARP Then
        With udtUMCNetAdjDef.MSWarpOptions
            iKVCount = 0
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassCalibrationType", .MassCalibrationType
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "MinimumPMTTagObsCount", .MinimumPMTTagObsCount
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MatchPromiscuity", .MatchPromiscuity
            
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "NETTol", .NETTol
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "NumberOfSections", .NumberOfSections
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MaxDistortion", .MaxDistortion
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "ContractionFactor", .ContractionFactor
            
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "MassWindowPPM", .MassWindowPPM
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassSplineOrder", .MassSplineOrder
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassNumXSlices", .MassNumXSlices
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassNumMassDeltaBins", .MassNumMassDeltaBins
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassMaxJump", .MassMaxJump
            
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "MassZScoreTolerance", .MassZScoreTolerance
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "MassUseLSQ", .MassUseLSQ
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "MassLSQOutlierZScore", .MassLSQOutlierZScore
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassLSQNumKnots", .MassLSQNumKnots
        End With
        IniStuff.WriteSection NET_ADJ_MS_WARP_SECTION, sKeys(), sVals(), iKVCount
    End If
        
        
'' Note: Uncomment the following to enable writing of the internal standards to a .Ini file
''    Dim intInternalStandardIndex As Integer
''
''    ' Write the Internal Standards
''    With udtInternalStandards
''        iKVCount = 0
''        AddKeyValueSetting sKeys, sVals, iKVCount, "Count", .Count
''    End With
''    IniStuff.WriteSection "UMCInternalStandards", sKeys(), sVals(), iKVCount
''
''    ' Write the Internal Standards
''    ' Each locker is written to its own section in the .Ini file
''    With udtInternalStandards
''        For intInternalStandardIndex = 0 To .Count - 1
''
''            With .InternalStandards(intInternalStandardIndex)
''
''                ' Write this Internal Standard
''
''                iKVCount = 0
''                AddKeyValueSettingLng sKeys, sVals, iKVCount, "SeqID", .SeqID
''                AddKeyValueSetting sKeys, sVals, iKVCount, "PeptideSequence", .PeptideSequence
''                AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MonoisotopicMass", .MonoisotopicMass
''                AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NET", .NET
''                AddKeyValueSettingInt sKeys, sVals, iKVCount, "ChargeMinimum", .ChargeMinimum
''                AddKeyValueSettingInt sKeys, sVals, iKVCount, "ChargeMaximum", .ChargeMaximum
''                AddKeyValueSettingInt sKeys, sVals, iKVCount, "ChargeMostAbundant", .ChargeMostAbundant
''
''                IniStuff.WriteSection "UMCInternalStandards" & Trim(intInternalStandardIndex + 1), sKeys(), sVals(), iKVCount
''            End With
''
''        Next intInternalStandardIndex
''    End With


    With udtAMTDef
        iKVCount = 0
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "SearchFlag", .SearchFlag
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MWField", .MWField
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MWTol", .MWTol
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "NETorRT", .NETorRT
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "TolType", .TolType
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETTol", .NETTol
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassTag", .MassTag
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MaxMassTags", .MaxMassTags
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SkipReferenced", .SkipReferenced
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveNCnt", .SaveNCnt
    End With
    IniStuff.WriteSection "SearchAMTDef", sKeys(), sVals(), iKVCount
   
    If Not bnlAutoAnalysisFieldsOnly Then
        With OlyOptions
            iKVCount = 0
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "DefType", CInt(.DefType)
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "DefShape", CInt(.DefShape)
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "DefColor", .DefColor
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DefVisible", .DefVisible
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "DefMinSize", .DefMinSize
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "DefMaxSize", .DefMaxSize
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "DefFontWidth", .DefFontWidth
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "DefFontHeight", .DefFontHeight
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "DefTextHeight", .DefTextHeight
            
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DefStickWidth", .DefStickWidth
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DefMinNET", .DefMinNET
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DefMaxNET", .DefMaxNET
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "DefNETAdjustment", .DefNETAdjustment
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DefNETTol", .DefNETTol
            
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DefUniformSize", .DefUniformSize
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DefBoxSizeAsSpotSize", .DefBoxSizeAsSpotSize
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DefWithID", .DefWithID
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DefCurrScopeVisible", .DefCurrScopeVisible
            
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "BackColor", .BackColor
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "ForeColor", .ForeColor
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "Orientation", .Orientation
        End With
        IniStuff.WriteSection "OlyOptions", sKeys(), sVals(), iKVCount
        
        With OlyOptions
            If Not .GRID Is Nothing Then
                With .GRID
                    iKVCount = 0
                    AddKeyValueSettingInt sKeys, sVals, iKVCount, "LineStyle", CInt(.LineStyle)
                    AddKeyValueSettingInt sKeys, sVals, iKVCount, "HorzAutoMode", CInt(.HorzAutoMode)
                    AddKeyValueSettingLng sKeys, sVals, iKVCount, "HorzBinsCount", .HorzBinsCount
                    AddKeyValueSettingBln sKeys, sVals, iKVCount, "HorzGridVisible", .HorzGridVisible
                    AddKeyValueSettingInt sKeys, sVals, iKVCount, "VertAutoMode", CInt(.VertAutoMode)
                    AddKeyValueSettingLng sKeys, sVals, iKVCount, "VertBinsCount", .VertBinsCount
                    AddKeyValueSettingBln sKeys, sVals, iKVCount, "VertGridVisible", .VertGridVisible
                End With
            End If
        End With
        IniStuff.WriteSection "OlyGridOptions", sKeys(), sVals(), iKVCount
        
        With OlyJiggyOptions
            iKVCount = 0
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseMWConstraint", .UseMWConstraint
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MWTol", .MWTol
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseNetConstraint", .UseNetConstraint
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETTol", .NETTol
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseAbuConstraint", .UseAbuConstraint
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "AbuTol", .AbuTol
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "JiggyScope", .JiggyScope
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "JiggyType", .JiggyType
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "BaseDisplayInd", .BaseDisplayInd
        End With
        IniStuff.WriteSection "OlyJiggyOptions", sKeys(), sVals(), iKVCount
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

    iKVCount = 0
    AddKeyValueSetting sKeys, sVals, iKVCount, "CoordinateSystem", sCooSysPref
    AddKeyValueSetting sKeys, sVals, iKVCount, "DifferentialDisplay", sDDClrPref
    AddKeyValueSetting sKeys, sVals, iKVCount, "Drawing", sDrawingPref
    AddKeyValueSetting sKeys, sVals, iKVCount, "ICR2LS", sICR2LSPref
    AddKeyValueSetting sKeys, sVals, iKVCount, "ChargeStateColors", sBackForeCSIsoClrPref
    AddKeyValueSetting sKeys, sVals, iKVCount, "ChargeStateShapes", sCSIsoShapePref
    AddKeyValueSetting sKeys, sVals, iKVCount, "Switches", sSwitchPref
    AddKeyValueSetting sKeys, sVals, iKVCount, "Tolerances", sTolerancesPref
    
    ' No longer supported (March 2006)
    ''AddKeyValueSetting sKeys, sVals, iKVCount, "AMTs", sAMTPref
    ''AddKeyValueSetting sKeys, sVals, iKVCount, "FTICRAmts", sFTICR_AMTPref
            
    IniStuff.WriteSection "Preferences", sKeys(), sVals(), iKVCount

    ' Write the expanded preferences
    With udtPrefsExpanded
        iKVCount = 0
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MenuModeDefault", CInt(.MenuModeDefault)
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "MenuModeIncludeObsolete", .MenuModeIncludeObsolete
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExtendedFileSaveModePreferred", .ExtendedFileSaveModePreferred
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoAdjSize", .AutoAdjSize
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "AutoSizeMultiplier", .AutoSizeMultiplier
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "UMCDrawType", .UMCDrawType
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UsePEKBasedERValues", .UsePEKBasedERValues
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseMassTagsWithNullMass", .UseMassTagsWithNullMass
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseMassTagsWithNullNET", .UseMassTagsWithNullNET
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "IReportAutoAddMonoPlus4AndMinus4Data", .IReportAutoAddMonoPlus4AndMinus4Data
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseUMCConglomerateNET", .UseUMCConglomerateNET
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "NetAdjustmentUsesN15AMTMasses", .NetAdjustmentUsesN15AMTMasses
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "NetAdjustmentMinHighNormalizedScore", .NetAdjustmentMinHighNormalizedScore
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "NetAdjustmentMinHighDiscriminantScore", .NetAdjustmentMinHighDiscriminantScore
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "AMTSearchResultsBehavior", CInt(.AMTSearchResultsBehavior)
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ICR2LSSpectrumViewZoomWindowWidthMZ", .ICR2LSSpectrumViewZoomWindowWidthMZ
        AddKeyValueSetting sKeys, sVals, iKVCount, "LastAutoAnalysisIniFilePath", .LastAutoAnalysisIniFilePath
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "LastInputFileMode", CInt(.LastInputFileMode)
        AddKeyValueSetting sKeys, sVals, iKVCount, "LegacyAMTDBPath", .LegacyAMTDBPath
    End With
    IniStuff.WriteSection "ExpandedPreferences", sKeys(), sVals(), iKVCount
    
    If Not bnlAutoAnalysisFieldsOnly And Not APP_BUILD_DISABLE_MTS Then
        ' Auto Query PRISM options
        With udtPrefsExpanded.AutoQueryPRISMOptions
            iKVCount = 0
            AddKeyValueSetting sKeys, sVals, iKVCount, "ConnectionStringQueryDB", .ConnectionStringQueryDB
            AddKeyValueSetting sKeys, sVals, iKVCount, "RequestTaskSPName", .RequestTaskSPName
            AddKeyValueSetting sKeys, sVals, iKVCount, "SetTaskCompleteSPName", .SetTaskCompleteSPName
            AddKeyValueSetting sKeys, sVals, iKVCount, "SetTaskToRestartSPName", .SetTaskToRestartSPName
            AddKeyValueSetting sKeys, sVals, iKVCount, "PostLogEntrySPName", .PostLogEntrySPName
            
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "QueryIntervalSeconds", .QueryIntervalSeconds
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MinimumPriorityToProcess", .MinimumPriorityToProcess
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "MaximumPriorityToProcess", .MaximumPriorityToProcess
            AddKeyValueSetting sKeys, sVals, iKVCount, "PreferredDatabaseToProcess", .PreferredDatabaseToProcess
            AddKeyValueSetting sKeys, sVals, iKVCount, "ServerForPreferredDatabase", .ServerForPreferredDatabase
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExclusivelyUseThisDatabase", .ExclusivelyUseThisDatabase
        End With
        IniStuff.WriteSection "AutoQueryPRISMOptions", sKeys(), sVals(), iKVCount
    End If
    
    ' Write the NET Adjustment UMC Selection Options
    With udtPrefsExpanded.NetAdjustmentUMCDistributionOptions
        iKVCount = 0
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RequireDispersedUMCSelection", .RequireDispersedUMCSelection
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "SegmentCount", .SegmentCount
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MinimumUMCsPerSegmentPctTopAbuPct", .MinimumUMCsPerSegmentPctTopAbuPct
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ScanPctStart", .ScanPctStart
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ScanPctEnd", .ScanPctEnd
    End With
    IniStuff.WriteSection "NetAdjustmentUMCDistributionOptions", sKeys(), sVals(), iKVCount
    
    ' Write the Error Plotting Options
    With udtPrefsExpanded.ErrorPlottingOptions
        iKVCount = 0
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MassRangePPM", .MassRangePPM
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MassBinSizePPM", .MassBinSizePPM
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "GANETRange", .GANETRange
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "GANETBinSize", .GANETBinSize
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "ButterWorthFrequency", .ButterWorthFrequency
    End With
    IniStuff.WriteSection "ErrorPlottingOptions", sKeys(), sVals(), iKVCount
    
    ' Write the Error Plotting Options -- Graph2D
    With udtPrefsExpanded.ErrorPlottingOptions.Graph2DOptions
        iKVCount = 0
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowPointSymbols", .ShowPointSymbols
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "DrawLinesBetweenPoints", .DrawLinesBetweenPoints
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowGridlines", .ShowGridLines
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoScaleXAxis", .AutoScaleXAxis
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointSizePixels", .PointSizePixels
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "LineWidthPixels", .LineWidthPixels
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "CenterYAxis", .CenterYAxis
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowSmoothedData", .ShowSmoothedData
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowPeakEdges", .ShowPeakEdges
    End With
    IniStuff.WriteSection "ErrorPlottingOptionsGraph2D", sKeys(), sVals(), iKVCount
    
    ' Write the Error Plotting Options -- Graph3D
    With udtPrefsExpanded.ErrorPlottingOptions.Graph3DOptions
        iKVCount = 0
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "ContourLevelsCount", .ContourLevelsCount
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "Perspective", .Perspective
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "Elevation", .Elevation
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "YRotation", .YRotation
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "ZRotation", .ZRotation
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "AnnotationFontSize", .AnnotationFontSize
    End With
    IniStuff.WriteSection "ErrorPlottingOptionsGraph3D", sKeys(), sVals(), iKVCount
    
    ' Write the noise removal options
    With udtPrefsExpanded.NoiseRemovalOptions
        iKVCount = 0
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "SearchTolerancePPMDefault", .SearchTolerancePPMDefault
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "SearchTolerancePPMAutoRemoval", .SearchTolerancePPMAutoRemoval
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "PercentageThresholdToExcludeSlice", .PercentageThresholdToExcludeSlice
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "PercentageThresholdToAddNeighborToSearchSlice", .PercentageThresholdToAddNeighborToSearchSlice
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "LimitMassRange", .LimitMassRange
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassStart", .MassStart
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassEnd", .MassEnd
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "LimitScanRange", .LimitScanRange
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "ScanStart", .ScanStart
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "ScanEnd", .ScanEnd
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "SearchScope", CInt(.SearchScope)
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RequireIdenticalCharge", .RequireIdenticalCharge
    End With
    IniStuff.WriteSection "NoiseRemovalOptions", sKeys(), sVals(), iKVCount
    frmProgress.UpdateProgressBar 2
    
    ' Write the Refine MS Data options
    With udtPrefsExpanded.RefineMSDataOptions
        iKVCount = 0
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MinimumPeakHeight", .MinimumPeakHeight
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MinimumSignalToNoiseRatioForLowAbundancePeaks", .MinimumSignalToNoiseRatioForLowAbundancePeaks
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "PercentageOfMaxForFindingWidth", .PercentageOfMaxForFindingWidth
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassCalibrationMaximumShift", .MassCalibrationMaximumShift
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassCalibrationTolType", CInt(.MassCalibrationTolType)
        
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ToleranceRefinementMethod", CInt(.ToleranceRefinementMethod)
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseMinMaxIfOutOfRange", .UseMinMaxIfOutOfRange
        
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassToleranceMinimum", .MassToleranceMinimum
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassToleranceMaximum", .MassToleranceMaximum
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassToleranceAdjustmentMultiplier", .MassToleranceAdjustmentMultiplier
        
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETToleranceMinimum", .NETToleranceMinimum
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETToleranceMaximum", .NETToleranceMaximum
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETToleranceAdjustmentMultiplier", .NETToleranceAdjustmentMultiplier
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "IncludeInternalStdMatches", .IncludeInternalStdMatches
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseUMCClassStats", .UseUMCClassStats
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MinimumSLiC", .MinimumSLiC
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MaximumAbundance", .MaximumAbundance
    
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "EMMassErrorPeakToleranceEstimatePPM", .EMMassErrorPeakToleranceEstimatePPM
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "EMNETErrorPeakToleranceEstimate", .EMNETErrorPeakToleranceEstimate
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "EMIterationCount", .EMIterationCount
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "EMPercentOfDataToExclude", .EMPercentOfDataToExclude
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "EMMassTolRefineForceUseSingleDataPointErrors", .EMMassTolRefineForceUseSingleDataPointErrors
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "EMNETTolRefineForceUseSingleDataPointErrors", .EMNETTolRefineForceUseSingleDataPointErrors
    
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ComputePairwiseMassDifferences", .ComputePairwiseMassDifferences
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "PairwiseMassDiffMinimum", .PairwiseMassDiffMinimum
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "PairwiseMassDiffMaximum", .PairwiseMassDiffMaximum
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "PairwiseMassBinSize", .PairwiseMassBinSize
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "PairwiseMassDiffNETTolerance", .PairwiseMassDiffNETTolerance
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "PairwiseMassDiffNETOffset", .PairwiseMassDiffNETOffset
    End With
    IniStuff.WriteSection "RefineMSDataOptions", sKeys(), sVals(), iKVCount
        
    ' Write the TIC Plotting Options
    With udtPrefsExpanded.TICAndBPIPlottingOptions
        iKVCount = 0
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "PlotNETOnXAxis", .PlotNETOnXAxis
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "NormalizeYAxis", .NormalizeYAxis
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SmoothUsingMovingAverage", .SmoothUsingMovingAverage
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MovingAverageWindowWidth", .MovingAverageWindowWidth
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "TimeDomainDataMaxValue", .TimeDomainDataMaxValue
        With .Graph2DOptions
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowPointSymbols", .ShowPointSymbols
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DrawLinesBetweenPoints", .DrawLinesBetweenPoints
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowGridlines", .ShowGridLines
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoScaleXAxis", .AutoScaleXAxis
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointSizePixels", .PointSizePixels
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "PointShape", .PointShape
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointAndLineColor", .PointAndLineColor
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "LineWidthPixels", .LineWidthPixels
        End With
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "PointShapeSeries2", .PointShapeSeries2
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointAndLineColorSeries2", .PointAndLineColorSeries2
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ClipOutliers", .ClipOutliers
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "ClipOutliersFactor", .ClipOutliersFactor
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "KeepWindowOnTop", .KeepWindowOnTop
    End With
    IniStuff.WriteSection "TICAndBPIPlottingOptions", sKeys(), sVals(), iKVCount
    
    ' Write the Pair Browser Options
    With udtPrefsExpanded.PairBrowserPlottingOptions
        iKVCount = 0
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "SortOrder", .SortOrder
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SortDescending", .SortDescending
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoZoom2DPlot", .AutoZoom2DPlot
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "HighlightMembers", .HighlightMembers
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "PlotAllChargeStates", .PlotAllChargeStates
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "FixedDimensionsForAutoZoom", .FixedDimensionsForAutoZoom
        
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassRangeZoom", .MassRangeZoom
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassRangeUnits", .MassRangeUnits
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ScanRangeZoom", .ScanRangeZoom
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ScanRangeUnits", .ScanRangeUnits
        
        With .Graph2DOptions
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowPointSymbols", .ShowPointSymbols
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DrawLinesBetweenPoints", .DrawLinesBetweenPoints
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowGridlines", .ShowGridLines
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointSizePixels", .PointSizePixels
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "PointShape", .PointShape
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointAndLineColor", .PointAndLineColor
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "LineWidthPixels", .LineWidthPixels
        End With
    
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "PointShapeHeavy", .PointShapeHeavy
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointAndLineColorHeavy", .PointAndLineColorHeavy
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "KeepWindowOnTop", .KeepWindowOnTop
    End With
    IniStuff.WriteSection "PairBrowserOptions", sKeys(), sVals(), iKVCount
    
    ' Write the UMC Browser Options
    With udtPrefsExpanded.UMCBrowserPlottingOptions
        iKVCount = 0
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "SortOrder", .SortOrder
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SortDescending", .SortDescending
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoZoom2DPlot", .AutoZoom2DPlot
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "HighlightMembers", .HighlightMembers
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "PlotAllChargeStates", .PlotAllChargeStates
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "FixedDimensionsForAutoZoom", .FixedDimensionsForAutoZoom
        
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassRangeZoom", .MassRangeZoom
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MassRangeUnits", .MassRangeUnits
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ScanRangeZoom", .ScanRangeZoom
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "ScanRangeUnits", .ScanRangeUnits
        
        With .Graph2DOptions
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowPointSymbols", .ShowPointSymbols
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "DrawLinesBetweenPoints", .DrawLinesBetweenPoints
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ShowGridlines", .ShowGridLines
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointSizePixels", .PointSizePixels
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "PointShape", .PointShape
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "PointAndLineColor", .PointAndLineColor
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "LineWidthPixels", .LineWidthPixels
        End With
    
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "KeepWindowOnTop", .KeepWindowOnTop
    End With
    IniStuff.WriteSection "UMCBrowserOptions", sKeys(), sVals(), iKVCount
    
    ' Write the Pair Search Options
    With udtPrefsExpanded.PairSearchOptions
        With .SearchDef
            iKVCount = 0
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DeltaMass", .DeltaMass
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DeltaMassTolerance", .DeltaMassTolerance
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "DeltaMassTolType", .DeltaMassTolType
            
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoCalculateDeltaMinMaxCount", .AutoCalculateDeltaMinMaxCount
            
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "DeltaCountMin", .DeltaCountMin
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "DeltaCountMax", .DeltaCountMax
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "DeltaStepSize", .DeltaStepSize
            
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "LightLabelMass", .LightLabelMass
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "HeavyLightMassDifference", .HeavyLightMassDifference
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "LabelCountMin", .LabelCountMin
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "LabelCountMax", .LabelCountMax
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "MaxDifferenceInNumberOfLightHeavyLabels", .MaxDifferenceInNumberOfLightHeavyLabels
            
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "RequireUMCOverlap", .RequireUMCOverlap
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "RequireUMCOverlapAtApex", .RequireUMCOverlapAtApex
            
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "ScanTolerance", .ScanTolerance
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "ScanToleranceAtApex", .ScanToleranceAtApex
            
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ERInclusionMin", .ERInclusionMin
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ERInclusionMax", .ERInclusionMax
            
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "RequireMatchingChargeStatesForPairMembers", .RequireMatchingChargeStatesForPairMembers
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseIdenticalChargesForER", .UseIdenticalChargesForER
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ComputeERScanByScan", .ComputeERScanByScan
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ScanByScanAverageIsNotWeighted", .ScanByScanAverageIsNotWeighted
            
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "RequireMatchingIsotopeTagLabels", .RequireMatchingIsotopeTagLabels
            
            AddKeyValueSettingByt sKeys, sVals, iKVCount, "MonoPlusMinusThresholdForceHeavyOrLight", .MonoPlusMinusThresholdForceHeavyOrLight
            AddKeyValueSettingByt sKeys, sVals, iKVCount, "IgnoreMonoPlus2AbundanceInIReportERCalc", .IgnoreMonoPlus2AbundanceInIReportERCalc
            
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "AverageERsAllChargeStates", .AverageERsAllChargeStates
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "AverageERsWeightingMode", .AverageERsWeightingMode
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "ERCalcType", .ERCalcType
        
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "RemoveOutlierERs", .RemoveOutlierERs
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "RemoveOutlierERsIterate", .RemoveOutlierERsIterate
            AddKeyValueSettingLng sKeys, sVals, iKVCount, "RemoveOutlierERsMinimumDataPointCount", .RemoveOutlierERsMinimumDataPointCount
            AddKeyValueSettingInt sKeys, sVals, iKVCount, "RemoveOutlierERsConfidenceLevel", .RemoveOutlierERsConfidenceLevel
        
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "N15IncompleteIncorporationMode", .N15IncompleteIncorporationMode
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "N15PercentIncorporationMinimum", .N15PercentIncorporationMinimum
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "N15PercentIncorporationMaximum", .N15PercentIncorporationMaximum
            AddKeyValueSettingSng sKeys, sVals, iKVCount, "N15PercentIncorporationStep", .N15PercentIncorporationStep
        
        End With
        
        AddKeyValueSetting sKeys, sVals, iKVCount, "PairSearchMode", .PairSearchMode
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoExcludeOutOfERRange", .AutoExcludeOutOfERRange
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoExcludeAmbiguous", .AutoExcludeAmbiguous
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "KeepMostConfidentAmbiguous", .KeepMostConfidentAmbiguous
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoAnalysisRemovePairMemberHitsAfterDBSearch", .AutoAnalysisRemovePairMemberHitsAfterDBSearch
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoAnalysisRemovePairMemberHitsRemoveHeavy", .AutoAnalysisRemovePairMemberHitsRemoveHeavy
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoAnalysisSavePairsToTextFile", .AutoAnalysisSavePairsToTextFile
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoAnalysisSavePairsStatisticsToTextFile", .AutoAnalysisSavePairsStatisticsToTextFile
        
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "NETAdjustmentPairedSearchUMCSelection", CInt(.NETAdjustmentPairedSearchUMCSelection)
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "AutoAnalysisDeltaMassAddnlCount", .AutoAnalysisDeltaMassAddnlCount
        
        If .AutoAnalysisDeltaMassAddnlCount > 0 Then
            For intIndex = 0 To .AutoAnalysisDeltaMassAddnlCount - 1
                AddKeyValueSettingDbl sKeys, sVals, iKVCount, "AutoAnalysisDeltaMassAddnl" & Trim(intIndex + 1), .AutoAnalysisDeltaMassAddnl(intIndex)
            Next intIndex
        Else
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "AutoAnalysisDeltaMassAddnl1", "0"
        End If
        
    End With
    IniStuff.WriteSection "PairSearchOptions", sKeys(), sVals(), iKVCount
        
        
    ' Write the IReport Pair options
    With udtPrefsExpanded.PairSearchOptions.SearchDef.IReportEROptions
        iKVCount = 0
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "Enabled", .Enabled
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NaturalAbundanceRatio2CoeffExponent", .NaturalAbundanceRatio2Coeff.Exponent
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NaturalAbundanceRatio2CoeffMultiplier", .NaturalAbundanceRatio2Coeff.Multiplier
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NaturalAbundanceRatio4CoeffExponent", .NaturalAbundanceRatio4Coeff.Exponent
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NaturalAbundanceRatio4CoeffMultiplier", .NaturalAbundanceRatio4Coeff.Multiplier
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MinimumFractionScansWithValidER", .MinimumFractionScansWithValidER
    End With
    IniStuff.WriteSection "IReportEROptions", sKeys(), sVals(), iKVCount
    
    
    ' Write the MT tag Staleness Options
    With udtPrefsExpanded.MassTagStalenessOptions
        iKVCount = 0
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MaximumAgeLoadedMassTagsHours", .MaximumAgeLoadedMassTagsHours
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MaximumFractionAMTsWithNulls", .MaximumFractionAMTsWithNulls
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MaximumCountAMTsWithNulls", .MaximumCountAMTsWithNulls
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MinimumTimeBetweenReloadMinutes", .MinimumTimeBetweenReloadMinutes
    End With
    IniStuff.WriteSection "MassTagStalenessOptions", sKeys(), sVals(), iKVCount
    
    
    ' Write the Match Score options
    With udtPrefsExpanded.SLiCScoreOptions
        iKVCount = 0
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "MassPPMStDev", .MassPPMStDev
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NETStDev", .NETStDev
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseAMTNETStDev", .UseAMTNETStDev
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MaxSearchDistanceMultiplier", .MaxSearchDistanceMultiplier
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoDefineSLiCScoreThresholds", .AutoDefineSLiCScoreThresholds
    End With
    IniStuff.WriteSection "SLiCScoreOptions", sKeys(), sVals(), iKVCount
    
    
    ' Write the GraphicExport options
    With udtPrefsExpanded.GraphicExportOptions
        iKVCount = 0
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "CopyEMFIncludeFilenameAndDate", .CopyEMFIncludeFilenameAndDate
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "CopyEMFIncludeTextLabels", .CopyEMFIncludeTextLabels
    End With
    IniStuff.WriteSection "GraphicExportOptions", sKeys(), sVals(), iKVCount
    
    ' Write the Auto Tolerance Refinement Options
    With udtPrefsExpanded.AutoAnalysisOptions.AutoToleranceRefinement
        iKVCount = 0
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DBSearchMWTol", .DBSearchMWTol
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "DBSearchTolType", CInt(.DBSearchTolType)
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "DBSearchNETTol", .DBSearchNETTol
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "DBSearchRegionShape", CInt(.DBSearchRegionShape)
        
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "DBSearchMinimumHighNormalizedScore", .DBSearchMinimumHighNormalizedScore
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "DBSearchMinimumHighDiscriminantScore", .DBSearchMinimumHighDiscriminantScore
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "DBSearchMinimumPeptideProphetProbability", .DBSearchMinimumPeptideProphetProbability
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RefineMassCalibration", .RefineMassCalibration
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RefineMassCalibrationOverridePPM", .RefineMassCalibrationOverridePPM
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RefineDBSearchMassTolerance", .RefineDBSearchMassTolerance
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RefineDBSearchNETTolerance", .RefineDBSearchNETTolerance
    End With
    IniStuff.WriteSection "AutoToleranceRefinement", sKeys(), sVals(), iKVCount
    
    
    ' Write the Auto Analysis Options
    With udtPrefsExpanded.AutoAnalysisOptions
        iKVCount = 0
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "MDType", 1
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AutoRemoveNoiseStreaks", .AutoRemoveNoiseStreaks
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "DoNotSaveOrExport", .DoNotSaveOrExport
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SkipFindUMCs", .SkipFindUMCs
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SkipGANETSlopeAndInterceptComputation", .SkipGANETSlopeAndInterceptComputation
        
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "DBConnectionRetryAttemptMax", .DBConnectionRetryAttemptMax
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "DBConnectionTimeoutSeconds", .DBConnectionTimeoutSeconds
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExportResultsFileUsesJobNumberInsteadOfDataSetName", .ExportResultsFileUsesJobNumberInsteadOfDataSetName
        
        ' Note: We're always storing False for .GenerateMonoPlus4IsoLabelingFile
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "GenerateMonoPlus4IsoLabelingFile", False

        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveGelFile", .SaveGelFile
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveGelFileOnError", .SaveGelFileOnError
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePictureGraphic", .SavePictureGraphic
        AddKeyValueSetting sKeys, sVals, iKVCount, "SavePictureGraphicFileTypeList", "; Options are " & GetPictureGraphicsTypeList()
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "SavePictureGraphicFileType", CInt(.SavePictureGraphicFileType)
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "SavePictureWidthPixels", .SavePictureWidthPixels
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "SavePictureHeightPixels", .SavePictureHeightPixels
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveInternalStdHitsAndData", .SaveInternalStdHitsAndData
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveErrorGraphicMass", .SaveErrorGraphicMass
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveErrorGraphicGANET", .SaveErrorGraphicGANET
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveErrorGraphic3D", .SaveErrorGraphic3D
        AddKeyValueSetting sKeys, sVals, iKVCount, "SaveErrorGraphicFileTypeList", "; Options are " & GetErrorGraphicsTypeList()
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "SaveErrorGraphicFileType", CInt(.SaveErrorGraphicFileType)
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "SaveErrorGraphSizeWidthPixels", .SaveErrorGraphSizeWidthPixels
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "SaveErrorGraphSizeHeightPixels", .SaveErrorGraphSizeHeightPixels
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotTIC", .SavePlotTIC
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotBPI", .SavePlotBPI
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotTICTimeDomain", .SavePlotTICTimeDomain
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotTICDataPointCounts", .SavePlotTICDataPointCounts
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotTICDataPointCountsHitsOnly", .SavePlotTICDataPointCountsHitsOnly
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotTICFromRawData", .SavePlotTICFromRawData
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotBPIFromRawData", .SavePlotBPIFromRawData
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotDeisotopingIntensityThresholds", .SavePlotDeisotopingIntensityThresholds
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SavePlotDeisotopingPeakCounts", .SavePlotDeisotopingPeakCounts
        
        AddKeyValueSetting sKeys, sVals, iKVCount, "OutputFileSeparationCharacter", .OutputFileSeparationCharacter
        AddKeyValueSetting sKeys, sVals, iKVCount, "PEKFileExtensionPreferenceOrder", .PEKFileExtensionPreferenceOrder
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "WriteIDResultsByIonToTextFileAfterAutoSearches", .WriteIDResultsByIonToTextFileAfterAutoSearches
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SaveUMCStatisticsToTextFile", .SaveUMCStatisticsToTextFile
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "IncludeORFNameInTextFileOutput", .IncludeORFNameInTextFileOutput
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "SetIsConfirmedForDBSearchMatches", .SetIsConfirmedForDBSearchMatches
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AddQuantitationDescriptionEntry", .AddQuantitationDescriptionEntry
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExportUMCsWithNoMatches", .ExportUMCsWithNoMatches
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "DBSearchRegionShape", CInt(.DBSearchRegionShape)
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "UseLegacyDBForMTs", .UseLegacyDBForMTs
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "IgnoreNETAdjustmentFailure", .IgnoreNETAdjustmentFailure
        
        If .AutoAnalysisSearchModeCount < 0 Then .AutoAnalysisSearchModeCount = 0
        If .AutoAnalysisSearchModeCount > MAX_AUTO_SEARCH_MODE_COUNT Then .AutoAnalysisSearchModeCount = MAX_AUTO_SEARCH_MODE_COUNT
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "AutoAnalysisSearchModeCount", .AutoAnalysisSearchModeCount
    End With
    IniStuff.WriteSection "AutoAnalysisOptions", sKeys(), sVals(), iKVCount
    
    ' Write the Auto Analysis Search Mode Settings
    ' Each search mode is written to its own section in the .Ini file
    With udtPrefsExpanded.AutoAnalysisOptions
        For intAutoSearchModeIndex = 0 To .AutoAnalysisSearchModeCount - 1
            With .AutoAnalysisSearchMode(intAutoSearchModeIndex)
                
                ' Write this Auto Analysis Search Mode's settings
                iKVCount = 0
                AddKeyValueSetting sKeys, sVals, iKVCount, "SearchModeList", "; Options are " & GetAutoAnalysisOptionsList()
                AddKeyValueSetting sKeys, sVals, iKVCount, "SearchMode", .SearchMode
                AddKeyValueSetting sKeys, sVals, iKVCount, "AlternateOutputFolderPath", .AlternateOutputFolderPath
                AddKeyValueSettingBln sKeys, sVals, iKVCount, "WriteResultsToTextFile", .WriteResultsToTextFile
                AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExportResultsToDatabase", .ExportResultsToDatabase
                AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExportUMCMembers", .ExportUMCMembers
                AddKeyValueSettingBln sKeys, sVals, iKVCount, "PairSearchAssumeMassTagsAreLabeled", .PairSearchAssumeMassTagsAreLabeled
                AddKeyValueSettingInt sKeys, sVals, iKVCount, "InternalStdSearchMode", CInt(.InternalStdSearchMode)
                AddKeyValueSettingSng sKeys, sVals, iKVCount, "DBSearchMinimumHighNormalizedScore", .DBSearchMinimumHighNormalizedScore
                AddKeyValueSettingSng sKeys, sVals, iKVCount, "DBSearchMinimumHighDiscriminantScore", .DBSearchMinimumHighDiscriminantScore
                AddKeyValueSettingSng sKeys, sVals, iKVCount, "DBSearchMinimumPeptideProphetProbability", .DBSearchMinimumPeptideProphetProbability
                
                With .MassMods
                    AddKeyValueSettingByt sKeys, sVals, iKVCount, "ModMode", .ModMode
                    AddKeyValueSettingBln sKeys, sVals, iKVCount, "N15InsteadOfN14", .N15InsteadOfN14
                    AddKeyValueSettingBln sKeys, sVals, iKVCount, "PEO", .PEO
                    AddKeyValueSettingBln sKeys, sVals, iKVCount, "ICATd0", .ICATd0
                    AddKeyValueSettingBln sKeys, sVals, iKVCount, "ICATd8", .ICATd8
                    AddKeyValueSettingBln sKeys, sVals, iKVCount, "Alkylation", .Alkylation
                    AddKeyValueSettingDbl sKeys, sVals, iKVCount, "AlkylationMass", .AlkylationMass
                    AddKeyValueSetting sKeys, sVals, iKVCount, "ResidueToModify", .ResidueToModify
                    AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ResidueMassModification", .ResidueMassModification
                End With
                
                IniStuff.WriteSection "AutoAnalysisSearchMode" & Trim(intAutoSearchModeIndex + 1), sKeys(), sVals(), iKVCount
            End With
            
        Next intAutoSearchModeIndex
    End With
    
    
    ' Write the Auto Analysis Filter Preferences
    With udtPrefsExpanded.AutoAnalysisFilterPrefs
        iKVCount = 0
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExcludeDuplicates", .ExcludeDuplicates
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ExcludeDuplicatesTolerance", .ExcludeDuplicatesTolerance
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExcludeIsoByFit", .ExcludeIsoByFit
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ExcludeIsoByFitMaxVal", .ExcludeIsoByFitMaxVal
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExcludeIsoSecondGuess", .ExcludeIsoSecondGuess
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExcludeIsoLessLikelyGuess", .ExcludeIsoLessLikelyGuess
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ExcludeCSByStdDev", .ExcludeCSByStdDev
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "ExcludeCSByStdDevMaxVal", .ExcludeCSByStdDevMaxVal
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictIsoByAbundance", .RestrictIsoByAbundance
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictIsoAbundanceMin", .RestrictIsoAbundanceMin
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictIsoAbundanceMax", .RestrictIsoAbundanceMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictIsoByMass", .RestrictIsoByMass
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictIsoMassMin", .RestrictIsoMassMin
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictIsoMassMax", .RestrictIsoMassMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictIsoByMZ", .RestrictIsoByMZ
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictIsoMZMin", .RestrictIsoMZMin
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictIsoMZMax", .RestrictIsoMZMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictIsoByChargeState", .RestrictIsoByChargeState
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "RestrictIsoChargeStateMin", .RestrictIsoChargeStateMin
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "RestrictIsoChargeStateMax", .RestrictIsoChargeStateMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictCSByAbundance", .RestrictCSByAbundance
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictCSAbundanceMin", .RestrictCSAbundanceMin
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictCSAbundanceMax", .RestrictCSAbundanceMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictCSByMass", .RestrictCSByMass
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictCSMassMin", .RestrictCSMassMin
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictCSMassMax", .RestrictCSMassMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictScanRange", .RestrictScanRange
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "RestrictScanRangeMin", .RestrictScanRangeMin
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "RestrictScanRangeMax", .RestrictScanRangeMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictGANETRange", .RestrictGANETRange
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictGANETRangeMin", .RestrictGANETRangeMin
        AddKeyValueSettingDbl sKeys, sVals, iKVCount, "RestrictGANETRangeMax", .RestrictGANETRangeMax
        
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictToEvenScanNumbersOnly", .RestrictToEvenScanNumbersOnly
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "RestrictToOddScanNumbersOnly", .RestrictToOddScanNumbersOnly
        
        ' Maximum data count filter
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "MaximumDataCountEnabled", .MaximumDataCountEnabled
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MaximumDataCountToLoad", .MaximumDataCountToLoad
    End With
    IniStuff.WriteSection "AutoAnalysisFilterPrefs", sKeys(), sVals(), iKVCount
    
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
        With udtPrefsExpanded.DMSConnectionInfo
            iKVCount = 0
            AddKeyValueSetting sKeys, sVals, iKVCount, "ConnectionString", .ConnectionString
        End With
        IniStuff.WriteSection "DMSConnectionInfo", sKeys(), sVals(), iKVCount
    
    
        ' Write the MTSConnectionInfo
        With udtPrefsExpanded.MTSConnectionInfo
            iKVCount = 0
            AddKeyValueSetting sKeys, sVals, iKVCount, "ConnectionString", .ConnectionString
            
            AddKeyValueSetting sKeys, sVals, iKVCount, "spAddQuantitationDescription", .spAddQuantitationDescription
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetLockers", .spGetLockers
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetMassTagMatchCount", .spGetMassTagMatchCount
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetMassTags", .spGetMassTags
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetMassTagsSubset", .spGetMassTagsSubset
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetPMResultStats", .spGetPMResultStats
            AddKeyValueSetting sKeys, sVals, iKVCount, "spPutAnalysis", .spPutAnalysis
            AddKeyValueSetting sKeys, sVals, iKVCount, "spPutUMC", .spPutUMC
            AddKeyValueSetting sKeys, sVals, iKVCount, "spPutUMCMember", .spPutUMCMember
            AddKeyValueSetting sKeys, sVals, iKVCount, "spPutUMCMatch", .spPutUMCMatch
            AddKeyValueSetting sKeys, sVals, iKVCount, "spPutUMCInternalStdMatch", .spPutUMCInternalStdMatch
            AddKeyValueSetting sKeys, sVals, iKVCount, "spEditGANET", .spEditGANET
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetORFs", .spGetORFs
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetORFSeq", .spGetORFSeq
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetORFIDs", .spGetORFIDs
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetORFRecord", .spGetORFRecord
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetMassTagSeq", .spGetMassTagSeq
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetMassTagNames", .spGetMassTagNames
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetInternalStandards", .spGetInternalStandards
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetDBSchemaVersion", .spGetDBSchemaVersion
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetMassTagToProteinNameMap", .spGetMassTagToProteinNameMap
            AddKeyValueSetting sKeys, sVals, iKVCount, "spGetMTStats", .spGetMTStats
        End With
        IniStuff.WriteSection "MTSConnectionInfo", sKeys(), sVals(), iKVCount
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
    
    Dim iKVCount As Integer
    Dim sKeys() As String, sVals() As String
    
    Dim intDBStuffArrayIndex As Integer, intDBStuffItemCount As Integer

    Dim intDBStuffCountIndex As Integer

    ReDim sKeys(99)
    ReDim sVals(99)
    
    ' Store the settings from udtRecentDBSettings(intIndex) in the ini file
    With udtDBSettingsSingle

        ' Write the version number
        iKVCount = 0
        AddKeyValueSettingInt sKeys, sVals, iKVCount, RECENT_DB_CONNECTION_INFOVERSION_NAME, RECENT_DB_CONNECTION_INFOVERSION
        
        ' Write the header items (summary variables)
        Debug.Assert .ConnectionString = .AnalysisInfo.MTDB.ConnectionString
        AddKeyValueSetting sKeys, sVals, iKVCount, "ConnectionString", .ConnectionString
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "DBSchemaVersion", .DBSchemaVersion
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "AmtsOnly", .AMTsOnly
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "ConfirmedOnly", .ConfirmedOnly
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "LockersOnly", .LockersOnly
        AddKeyValueSettingBln sKeys, sVals, iKVCount, "LimitToPMTsFromDataset", .LimitToPMTsFromDataset
        
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MinimumHighNormalizedScore", .MinimumHighNormalizedScore
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MinimumHighDiscriminantScore", .MinimumHighDiscriminantScore
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MinimumPeptideProphetProbability", .MinimumPeptideProphetProbability
        AddKeyValueSettingSng sKeys, sVals, iKVCount, "MinimumPMTQualityScore", .MinimumPMTQualityScore
        
        AddKeyValueSetting sKeys, sVals, iKVCount, "ExperimentInclusionFilter", .ExperimentInclusionFilter
        AddKeyValueSetting sKeys, sVals, iKVCount, "ExperimentExclusionFilter", .ExperimentExclusionFilter
        AddKeyValueSetting sKeys, sVals, iKVCount, "InternalStandardExplicit", .InternalStandardExplicit
        
        AddKeyValueSettingInt sKeys, sVals, iKVCount, "NETValueType", .NETValueType
        
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "MassTagSubsetID", .MassTagSubsetID
        AddKeyValueSetting sKeys, sVals, iKVCount, "ModificationList", .ModificationList
        
        AddKeyValueSettingLng sKeys, sVals, iKVCount, "SelectedMassTagCount", .SelectedMassTagCount
        
        ' Now write the values in .AnalysisInfo
        With .AnalysisInfo
            
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "GANET_Fit", .GANET_Fit
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "GANET_Intercept", .GANET_Intercept
            AddKeyValueSettingDbl sKeys, sVals, iKVCount, "GANET_Slope", .GANET_Slope
            AddKeyValueSettingBln sKeys, sVals, iKVCount, "ValidAnalysisDataPresent", .ValidAnalysisDataPresent
            
            If blnIncludeDetailedAnalysisInfo Then
                AddKeyValueSetting sKeys, sVals, iKVCount, "AnalysisTool", .Analysis_Tool
                AddKeyValueSetting sKeys, sVals, iKVCount, "Created", .Created
                AddKeyValueSetting sKeys, sVals, iKVCount, "Dataset", .Dataset
                AddKeyValueSetting sKeys, sVals, iKVCount, "Dataset_Folder", .Dataset_Folder
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "Dataset_ID", .Dataset_ID
                AddKeyValueSetting sKeys, sVals, iKVCount, "Desc_DataFolder", .Desc_DataFolder
                AddKeyValueSetting sKeys, sVals, iKVCount, "Desc_Type", .Desc_Type
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "Duration", .Duration
                AddKeyValueSetting sKeys, sVals, iKVCount, "Experiment", .Experiment
                AddKeyValueSetting sKeys, sVals, iKVCount, "Instrument_Class", .Instrument_Class
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "Job", .Job
                AddKeyValueSetting sKeys, sVals, iKVCount, "MD_Date", .MD_Date
                AddKeyValueSetting sKeys, sVals, iKVCount, "MD_file", .MD_file
                AddKeyValueSetting sKeys, sVals, iKVCount, "MD_Parameters", .MD_Parameters
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "MD_Reference_Job", .MD_Reference_Job
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "MD_State", .MD_State
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "MD_Type", .MD_Type
                AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NET_Intercept", .NET_Intercept
                AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NET_Slope", .NET_Slope
                AddKeyValueSettingDbl sKeys, sVals, iKVCount, "NET_TICFit", .NET_TICFit
                AddKeyValueSetting sKeys, sVals, iKVCount, "Organism", .Organism
                AddKeyValueSetting sKeys, sVals, iKVCount, "Organism_DB_Name", .Organism_DB_Name
                AddKeyValueSetting sKeys, sVals, iKVCount, "Parameter_File_Name", .Parameter_File_Name
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "ProcessingType", .ProcessingType
                AddKeyValueSetting sKeys, sVals, iKVCount, "Results_Folder", .Results_Folder
                AddKeyValueSetting sKeys, sVals, iKVCount, "Settings_File_Name", .Settings_File_Name
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "State", .STATE
                AddKeyValueSetting sKeys, sVals, iKVCount, "Storage_Path", .Storage_Path
                AddKeyValueSettingLng sKeys, sVals, iKVCount, "Total_Scans", .Total_Scans
                AddKeyValueSetting sKeys, sVals, iKVCount, "Vol_Client", .Vol_Client
                AddKeyValueSetting sKeys, sVals, iKVCount, "Vol_Server", .Vol_Server
            End If
            
            objIniStuff.WriteSection strSectionName, sKeys(), sVals(), iKVCount
    
            If blnIncludeMtdbDBStuff Then
                strSectionName = strSectionName & "_" & "MTDB"
                With .MTDB
                    ' Write the MTDB items
                    
                    iKVCount = 0
                    AddKeyValueSettingLng sKeys, sVals, iKVCount, "DBStatus", .DBStatus
                    AddKeyValueSettingInt sKeys, sVals, iKVCount, "DBStuffCount", 0        ' Note: This value will be updated below
                    intDBStuffCountIndex = iKVCount - 1
                    
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
                            AddKeyValueSetting sKeys, sVals, iKVCount, "DBStuffItem" & Trim(intDBStuffItemCount) & "Name", .DBStuffArray(intDBStuffArrayIndex).Name
                            
                            AddKeyValueSetting sKeys, sVals, iKVCount, "DBStuffItem" & Trim(intDBStuffItemCount) & "Value", .DBStuffArray(intDBStuffArrayIndex).Value
                            intDBStuffItemCount = intDBStuffItemCount + 1
                        End Select
                    Next intDBStuffArrayIndex
                    
                    If intDBStuffItemCount > 0 Then
                        ' This assertion fails when the contents of the DBStuffArray are changed, which happens from time to time
                        ' The assertion is only here to let the programmer know that the contents have changed; it's not necessarily a problem
                        Debug.Assert intDBStuffItemCount = 36
                    End If
                    
                    ' Check this
                    Debug.Assert False
                    Debug.Assert sKeys(intDBStuffCountIndex) = "DBStuffCount"
                    sVals(intDBStuffCountIndex) = intDBStuffItemCount
                    
                    objIniStuff.WriteSection strSectionName, sKeys(), sVals(), iKVCount
                
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
        .IsoDataField = mftMWMono       ' 7
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
        .ModMode = 1                    ' Was previously: .DynamicMods = True
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

Public Sub ResetDataFilters(ByVal lngGelIndex As Long, ByRef udtPreferences As GelPrefs)
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
            
            .IReportAutoAddMonoPlus4AndMinus4Data = True
            
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
                
                .ComputePairwiseMassDifferences = False
                .PairwiseMassDiffMinimum = -100
                .PairwiseMassDiffMaximum = 100
                .PairwiseMassBinSize = 0.25
                .PairwiseMassDiffNETTolerance = 0.1
                .PairwiseMassDiffNETOffset = 0
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
                    .DeltaMassTolType = gltABS
                    
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
                    
                    .ScanTolerance = 25
                    .ScanToleranceAtApex = 25
                    
                    .ERInclusionMin = -5
                    .ERInclusionMax = 5
                    
                    .RequireMatchingChargeStatesForPairMembers = True
                    .UseIdenticalChargesForER = True
                    .ComputeERScanByScan = True
                    .ScanByScanAverageIsNotWeighted = False
                    
                    .RequireMatchingIsotopeTagLabels = False
                    
                    .MonoPlusMinusThresholdForceHeavyOrLight = 66
                    .IgnoreMonoPlus2AbundanceInIReportERCalc = 0
                    
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
                
                    .N15IncompleteIncorporationMode = False
                    .N15PercentIncorporationMinimum = 95
                    .N15PercentIncorporationMaximum = 95
                    .N15PercentIncorporationStep = 1
                    
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
                
                .GenerateMonoPlus4IsoLabelingFile = False
                
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
                    .spGetMTStats = "GetMTStatsAndPepProphetStats"
                    
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

Public Sub SaveCurrentSettingsToIniFile(ByVal lngGelIndex As Long)
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


Attribute VB_Name = "Module14"
'module dealing with ICR-2LS incorporation
'last modified 10/01/2002 nt
'-------------------------------------------------------------
Option Explicit

'the version listed here is not latest ICR-2LS version;
'rather it is the first version that provides interfaces
'used in this application

Private Const glICR2LS_VERSION_MAJOR = 2
Private Const glICR2LS_VERSION_MINOR = 25
Private Const glICR2LS_VERSION_REVISION = 102

Private Enum ieICR2LSErrorConstants
    ieStatusOK = 0
    ieStatusError = 1
    ieStatusOldVersion = 2
End Enum

Public Enum isfICR2LSScopeFunctionConstants
    isfConvert2Mass = 1
    isfConvert2Time = 2
    isfConvert2Freq = 3
    isfZoom = 4
    isfMarkCS = 5
    isfZeroFill = 6
    isfApodize = 7
    isfGetPoint = 8
    isfSetPoint = 9
    isfPickPeaks = 10
    isfPickPeaksReport = 11
End Enum

'Public objICR2LS As icr2ls.ICR2LScls
Public objICR2LS As Object

Public icrStatus2ls As Long

Private mMSSpectraCache As clsMSSpectraCache

Public Sub CleanICR2LS()
    On Error Resume Next
    If objICR2LS Is Nothing Then Exit Sub
    objICR2LS.Kill
    Set objICR2LS = Nothing
    Set mMSSpectraCache = Nothing
    'If Err Then LogErrors Err.Number, "CleanICR2LS"
End Sub

Private Function GetICR2LSVerStatus(ByVal VerStr As String) As Long
Dim StartPos As Integer
Dim EndPos As Integer
Dim VerNumStr As String

Dim lngVersionMajor As Long
Dim lngVersionMinor As Long
Dim lngRevisionNumber As Long

Dim blnValidVersion As Boolean

On Error GoTo exit_ErrorReadingVersion

If LCase(Left(VerStr, 7)) = "version" Then
    VerStr = Trim(Mid(VerStr, 8))

    EndPos = InStr(VerStr, ".")
    VerNumStr = Trim(Left(VerStr, EndPos - 1))

    If IsNumeric(VerNumStr) Then
        lngVersionMajor = CLng(VerNumStr)
            
        StartPos = EndPos + 1
        EndPos = InStr(StartPos, VerStr, ".")
        VerNumStr = Trim(Mid(VerStr, StartPos, EndPos - StartPos))
        If IsNumeric(VerNumStr) Then
            lngVersionMinor = CLng(VerNumStr)
                
            StartPos = EndPos + 1
            EndPos = InStr(StartPos, VerStr, ",")
            VerNumStr = Trim(Mid(VerStr, StartPos, EndPos - StartPos))
            If IsNumeric(VerNumStr) Then
                lngRevisionNumber = CLng(VerNumStr)
            End If
        End If
    End If
End If

blnValidVersion = True
If lngVersionMajor < glICR2LS_VERSION_MAJOR Then
    blnValidVersion = False
ElseIf lngVersionMajor = glICR2LS_VERSION_MAJOR Then
    If lngVersionMinor < glICR2LS_VERSION_MINOR Then
        blnValidVersion = False
    ElseIf lngVersionMinor = glICR2LS_VERSION_MINOR Then
        If lngRevisionNumber < glICR2LS_VERSION_REVISION Then
            blnValidVersion = False
        End If
    End If
End If

If blnValidVersion Then
    'if you came here version is current(relative to 2D Gels) or better
    GetICR2LSVerStatus = ieStatusOK
    Exit Function
Else
    MsgBox "Version of ICR-2LS on local machine older than expected. Some functions might not work? Make sure that latest version of the ICR-2LS is installed on this machine.", vbOKOnly
    GetICR2LSVerStatus = ieStatusOldVersion
    Exit Function
End If

exit_ErrorReadingVersion:
MsgBox "Error reading ICR-2LS version information. Make sure that latest version of the ICR-2LS is installed on this machine.", vbOKOnly
GetICR2LSVerStatus = ieStatusError
End Function

Public Function ICR2LSLoadSpectrumViaCache(ByVal lngGelIndex As Long, ByVal lngScanNumber As Long, ByVal blnIsMoverZ As Boolean, ByVal dblTargetMZ As Double, Optional ByRef hScope As Integer = 0, Optional dblVisibleMZMinimum As Double = 0, Optional dblVisibleMZMaximum As Double = 0)
    ' Returns True if spectrum successfully cached and loaded
    ' Otherwise, returns false
    
    Dim strPathToRawFiles As String
    Dim strRawDataFileName As String
    Dim strPathFromHistory As String
    Dim eFileType As ifmInputFileModeConstants
    
    Dim strSpectrumPathOut As String
    
    Dim lngHistoryIndexOfMatch As Long
    Dim strInputFilePath As String
    
    Dim blnSuccess As Boolean
    
    Dim lngIndex As Long
    
    Dim fso As FileSystemObject
    Dim objFolder As Folder
    
On Error GoTo ICR2LSLoadSpectrumViaCacheErrorHandler

    With GelData(lngGelIndex)
        strPathToRawFiles = .PathtoDataFiles
        
        ' For compatibility with older version where full path to raw data file was stored
        ' in gel file arrays, we have to use GetFileNameOnly function
        lngIndex = GetDFIndex(lngGelIndex, lngScanNumber)
        
        If lngIndex >= 0 Then
            strRawDataFileName = GetFileNameOnly(.ScanInfo(lngIndex).ScanFileName)
        Else
            ' Index not found; this means the user's mouse was too far away from a data point
            MsgBox "Unable to display spectrum: " & vbCrLf & "Spectrum index not found for scan " & CStr(lngScanNumber), vbExclamation + vbOKOnly, "Error"
            ICR2LSLoadSpectrumViaCache = False
            Exit Function
        End If
        
        strInputFilePath = .FileName
    End With
   
    If LCase(Right(strInputFilePath, 4)) = ".gel" Then
        ' User opened an old file that does not store the .PEK/.CSV/.mzXML/.mzData file path in .Filename
        ' Look in the AnalysisHistory for this information
        ' If it is not found, then the cache folder will be named based on the Gel file name (which isn't bad, just isn't ideal)
        strPathFromHistory = FindSettingInAnalysisHistory(lngGelIndex, glCOMMENT_DATA_FILE_START, lngHistoryIndexOfMatch, True, ":")
        If lngHistoryIndexOfMatch >= 0 Then
            If DetermineFileType(strPathFromHistory, eFileType) Then
                With GelData(lngGelIndex)
                    ' Original path to .PEK/.CSV/.mzXML/.mzData file found; update GelData().Filename
                    .FileName = strPathFromHistory
                    strInputFilePath = strPathFromHistory
                End With
                GelStatus(lngGelIndex).Dirty = True
            End If
        End If
    End If
    
    If UCase(Left(Trim(strPathToRawFiles), 7)) = "C:\DMS_" Then
        ' Most likely an invalid path to the zip files
        ' Assume the Zip files are one folder up from that file's folder
        Set fso = New FileSystemObject
        
        If fso.FolderExists(fso.GetParentFolderName(strInputFilePath)) Then
            
            Set objFolder = fso.GetFolder(fso.GetParentFolderName(strInputFilePath))
            If Not objFolder.IsRootFolder Then
                Set objFolder = objFolder.ParentFolder
            End If
            strPathToRawFiles = objFolder.Path
            
            GelData(lngGelIndex).PathtoDataFiles = strPathToRawFiles
            GelStatus(lngGelIndex).Dirty = True
        End If
        Set fso = Nothing
    End If
    
    If mMSSpectraCache.ExtractAndCacheSpectrum(strPathToRawFiles, strRawDataFileName, lngScanNumber, strInputFilePath, strSpectrumPathOut) Then
        If dblTargetMZ = 0 Then
            ICR2LSLoadScope strSpectrumPathOut, hScope
        ElseIf blnIsMoverZ Then
            If dblVisibleMZMinimum = 0 And dblVisibleMZMaximum = 0 Then
                ICR2LSLoadScopeMOverZ strSpectrumPathOut, dblTargetMZ, hScope
            Else
                ICR2LSLoadScopeMOverZ strSpectrumPathOut, dblTargetMZ, hScope, Abs(dblTargetMZ - dblVisibleMZMinimum), Abs(dblTargetMZ - dblVisibleMZMaximum)
            End If
        Else
            If dblVisibleMZMinimum = 0 And dblVisibleMZMaximum = 0 Then
                ICR2LSLoadScopeMW strSpectrumPathOut, hScope, dblTargetMZ
            Else
                ICR2LSLoadScopeMW strSpectrumPathOut, hScope, dblTargetMZ, Abs(dblTargetMZ - dblVisibleMZMinimum), Abs(dblTargetMZ - dblVisibleMZMaximum)
            End If
        End If
        blnSuccess = True
    Else
        ' Error caching spectrum
        If mMSSpectraCache.ErrorCode = scOriginalDataFolderNotFound Then
            MsgBox "Original path to zipped data files not found: " & strPathToRawFiles & vbCrLf & "Use 'Edit->Display Parameters and Paths' to enter a valid path.", vbExclamation + vbOKOnly, "Error"
        ElseIf mMSSpectraCache.ErrorCode = scZipFilesNotFoundInDataFolder Then
            Set fso = New FileSystemObject
            MsgBox "Unable to display spectrum: " & vbCrLf & mMSSpectraCache.GetErrorMessage & vbCrLf & _
                   "Currently defined data folder path is: " & strPathToRawFiles & vbCrLf & _
                   "Use 'Edit->Display Parameters and Paths' to update this path, if necessary.", vbExclamation + vbOKOnly, "Error"
            Set fso = Nothing
        Else
            MsgBox "Unable to display spectrum: " & mMSSpectraCache.GetErrorMessage, vbExclamation + vbOKOnly, "Error"
        End If
        blnSuccess = False
    End If

    ICR2LSLoadSpectrumViaCache = blnSuccess
    Exit Function

ICR2LSLoadSpectrumViaCacheErrorHandler:
    Err.Clear
    ICR2LSLoadSpectrumViaCache = False
    
End Function

Public Sub ICR2LSCloseAllScopes(blnConfirmClose As Boolean)
    Const MAX_SCOPE_COUNT = 15
    
    Dim i As Integer
    Dim eResponse As VbMsgBoxResult
    
    If objICR2LS Is Nothing Then Exit Sub
    
    If blnConfirmClose Then
        eResponse = MsgBox("Close all ICR-2LS spectra windows?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Close Scopes")
        If eResponse <> VbMsgBoxResult.vbYes Then Exit Sub
    End If
    
    On Error Resume Next
    For i = 0 To MAX_SCOPE_COUNT - 1
        objICR2LS.KillScope = i
        DoEvents
    Next i
End Sub

Public Function ICR2LSLoadFinniganSpectrum(ByVal lngGelIndex As Long, ByVal lngScanNumber As Long, ByVal dblTargetMZ As Double, Optional ByRef hScope As Integer = 0, Optional dblVisibleMZMinimum As Double = 0, Optional dblVisibleMZMaximum As Double = 0) As Boolean

    ' Returns True if spectrum successfully displayed
    ' Otherwise, returns false
    
    Dim strPathToRawFile As String
    
    Dim blnSuccess As Boolean
    
On Error GoTo ICR2LSLoadFinniganSpectrumErrorHandler

    If GelStatus(lngGelIndex).SourceDataRawFileType <> rfcFinniganRaw Then
        DetermineSourceDataRawFileType lngGelIndex, True
    End If
    
    If GelStatus(lngGelIndex).SourceDataRawFileType <> rfcFinniganRaw Then
        blnSuccess = False
    Else
        strPathToRawFile = GelStatus(lngGelIndex).FinniganRawFilePath
        
        If FileExists(strPathToRawFile) Then
            If dblVisibleMZMinimum = 0 And dblVisibleMZMaximum = 0 Then
                ICR2LSLoadFinniganScopeMOverZ strPathToRawFile, lngScanNumber, dblTargetMZ, hScope
            Else
                ICR2LSLoadFinniganScopeMOverZ strPathToRawFile, lngScanNumber, dblTargetMZ, hScope, Abs(dblTargetMZ - dblVisibleMZMinimum), Abs(dblTargetMZ - dblVisibleMZMaximum)
            End If
            If hScope >= 0 Then
                blnSuccess = True
            Else
                blnSuccess = False
            End If
        Else
            MsgBox "Finnigan .Raw file not found: " & strPathToRawFile & vbCrLf & "Use 'Edit->Display Parameters and Paths' to enter a valid path to the folder containing the file.", vbExclamation + vbOKOnly, "Error"
            blnSuccess = False
        End If
        
    End If

    ICR2LSLoadFinniganSpectrum = blnSuccess
    Exit Function

ICR2LSLoadFinniganSpectrumErrorHandler:
    ICR2LSLoadFinniganSpectrum = False

End Function

Public Sub ICR2LSLoadScope(ByVal FileName As String, ByRef hScope As Integer)
    Dim opScope As Integer
    On Error Resume Next
    
    If objICR2LS Is Nothing Then Exit Sub
    
    hScope = objICR2LS.LoadScope(FileName)
    If Err.Number <> 0 Then
        hScope = -1
        Exit Sub
    End If

    opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfConvert2Mass, hScope, 1000)
    
End Sub

Private Sub ICR2LSLoadScopeMW(ByVal FileName As String, ByRef hScope As Integer, ByVal MW As Double, Optional ByVal dblZoomWidthMZLeft As Double = 1.25, Optional ByVal dblZoomWidthMZRight As Double = 3.75)
    Dim opScope As Integer
    On Error Resume Next
    
    If objICR2LS Is Nothing Then Exit Sub
    
    hScope = objICR2LS.LoadScope(FileName)
    If Err.Number <> 0 Then
        hScope = -1
        Exit Sub
    End If
    
    opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfConvert2Mass, hScope, MW)
    'zoom only if it is good chance that this is charge state 1
    If MW < 1000 Then
        opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfZoom, hScope, MW - dblZoomWidthMZLeft, MW + dblZoomWidthMZRight)
    Else
        opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfZoom, hScope, 200, 2500)
    End If
End Sub

Private Sub ICR2LSLoadScopeMOverZ(ByVal FileName As String, ByVal MOverZ As Double, ByRef hScope As Integer, Optional ByVal dblZoomWidthMZLeft As Double = 1.25, Optional ByVal dblZoomWidthMZRight As Double = 3.75)
    Dim opScope As Integer
    On Error Resume Next
    
    If objICR2LS Is Nothing Then Exit Sub
   
    hScope = objICR2LS.LoadScope(FileName)
    If Err.Number <> 0 Then
        hScope = -1
        Exit Sub
    End If
    
    opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfConvert2Mass, hScope, MOverZ - dblZoomWidthMZLeft, MOverZ + dblZoomWidthMZRight)
    opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfZoom, hScope, MOverZ - dblZoomWidthMZLeft, MOverZ + dblZoomWidthMZRight)

End Sub

Private Sub ICR2LSLoadFinniganScopeMOverZ(ByVal RawFilePath As String, ByVal ScanNumber As Long, ByVal MOverZ As Double, ByRef hScope As Integer, Optional ByVal dblZoomWidthMZLeft As Double = 1.25, Optional ByVal dblZoomWidthMZRight As Double = 3.75)
    Dim opScope As Integer
    On Error Resume Next
    
    If objICR2LS Is Nothing Then Exit Sub
    
    If ScanNumber > 32768 Then
        MsgBox "ICR-2LS contains a programming error and cannot display scan numbers larger than 32,768.  If you need to do this, contact NavDeep Jaitly.", vbExclamation + vbOKOnly, "Error"
        hScope = -1
        Exit Sub
    End If
    
    hScope = objICR2LS.LoadLCQScan(RawFilePath, CInt(ScanNumber))
    If Err.Number <> 0 Then
        hScope = -1
        Exit Sub
    End If
    
    opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfConvert2Mass, hScope, MOverZ - dblZoomWidthMZLeft, MOverZ + dblZoomWidthMZRight)
    opScope = objICR2LS.ScopeOperation(isfICR2LSScopeFunctionConstants.isfZoom, hScope, MOverZ - dblZoomWidthMZLeft, MOverZ + dblZoomWidthMZRight)

End Sub

Public Sub InitICR2LS()
    On Error Resume Next
    Dim strMessage As String

    'Set objICR2LS = New ICR2LScls
    Set objICR2LS = CreateObject("icr2ls.ICR2LScls")
    
    If Err Then
       LogErrors Err.Number, "InitICR2LS"
       If Not CommandLineContainsAutomationCommand() Then
          ' No longer notifying user if ICR-2LS doesn't register properly; this doesn't really matter
          
          strMessage = ""
          strMessage = strMessage & "Error creating ICR-2LS object. You will be unable to view raw mass spectra. "
          strMessage = strMessage & "Make sure that latest versions of both ICR-2LS and Viper are installed on this computer. "
          strMessage = strMessage & "If inside PNNL, go to \\floyd\software\Viper\ and \\floyd\software\ICR-2LS\.  If outside PNNL, please visit http://omics.pnl.gov/software/"
          MsgBox strMessage, vbInformation + vbOKOnly, "Error"
       End If
       Err.Clear
       Set objICR2LS = Nothing
       icrStatus2ls = ieStatusError
    Else
       icrStatus2ls = GetICR2LSVerStatus(objICR2LS.ICR2LSversion)
    End If
    
    Set mMSSpectraCache = New clsMSSpectraCache
        
End Sub


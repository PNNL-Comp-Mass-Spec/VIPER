Attribute VB_Name = "Module2"
'PROCEDURES TO DEAL WITH FILES (OPENING, SAVING ETC.), MENUS
'PROCEDURES TO DEAL WITH GEL DATA (INCLUSIONS, EXCLUSIONS, ETC.)
'Last modified 03/25/2003 nt
'---------------------------------------------------------------
Option Explicit

Private Const INI_SECTION_RECENT_FILES = "RecentFiles"
Private Const INI_KEY_RECENT_FILE_COUNT = "RecentFileCount"
Private Const INI_KEY_RECENT_FILE_PREFIX = "RecentFile"

Public Function DetermineFileType(ByVal strFileNameOrPath As String, ByRef eFileType As ifmInputFileModeConstants) As Boolean
    ' Returns true if the file type is known
    ' Returns false if unknown or an error
    
    Dim fso As New FileSystemObject
    Dim strFileExtension As String
    Dim strFileName As String
    Dim blnKnownType As Boolean
    
On Error GoTo DetermineFileTypeErrorHandler

    strFileExtension = fso.GetExtensionName(strFileNameOrPath)
    If Len(strFileExtension) > 0 Then
        If Left(strFileExtension, 1) <> "." Then
            strFileExtension = "." & strFileExtension
        End If
    End If
    
    blnKnownType = True
    Select Case LCase(strFileExtension)
    Case ".pek", ".pek-3"
        eFileType = ifmInputFileModeConstants.ifmPEKFile
    Case ".csv"
        eFileType = ifmInputFileModeConstants.ifmCSVFile
    Case ".mzxml"
        eFileType = ifmInputFileModeConstants.ifmmzXMLFile
    Case ".mzdata"
        eFileType = ifmInputFileModeConstants.ifmmzDataFile
    Case ".gel"
        eFileType = ifmInputFileModeConstants.ifmGelFile
    Case ".xml"
        ' Look for mzXml and mzData in the file name
        strFileName = LCase(fso.GetBaseName(strFileNameOrPath))
        If InStr(strFileName, "mzxml") > 0 Then
            eFileType = ifmInputFileModeConstants.ifmmzXMLFileWithXMLExtension
        ElseIf InStr(strFileName, "mzdata") > 0 Then
            eFileType = ifmInputFileModeConstants.ifmmzDataFileWithXMLExtension
        Else
            blnKnownType = False
        End If
    Case ".txt"
        eFileType = ifmInputFileModeConstants.ifmDelimitedTextFile
    Case ".msalign"
        eFileType = ifmInputFileModeConstants.ifmMSAlign
    Case Else
        blnKnownType = False
    End Select
    
    DetermineFileType = blnKnownType
    Exit Function

DetermineFileTypeErrorHandler:
    Debug.Assert False
    DetermineFileType = False
    
End Function

Public Function DetermineParentFolderPath(strFileNameOrPath As String) As String
    
    Dim fso As New FileSystemObject
    Dim objFile As File
    Dim objFolder As Folder
    
    Dim blnUseParentFolder As Boolean
    Dim intIndex As Integer
    Dim intCharLoc As Integer
    Dim intAsciiValue As Integer
       
    Dim strParentFolderPath As String

On Error GoTo DetermineParentFolderPathErrorHandler

    strParentFolderPath = ""
    
    Set objFile = fso.GetFile(strFileNameOrPath)
    
    ' Initially set strParentFolderPath to the folder that objFile resides in
    strParentFolderPath = objFile.ParentFolder
    
    blnUseParentFolder = False
    intCharLoc = InStr(LCase(strParentFolderPath), "_auto")
    
    If intCharLoc > 1 Then
        blnUseParentFolder = True
        
        ' Make sure _auto is only followed by digits
        For intIndex = intCharLoc + 5 To Len(strParentFolderPath)
            intAsciiValue = Asc(Mid(strParentFolderPath, intIndex, 1))
            If intAsciiValue < 48 Or intAsciiValue > 57 Then
                blnUseParentFolder = False
                Exit For
            End If
        Next intIndex
    End If
    
    If blnUseParentFolder Then
        Set objFolder = objFile.ParentFolder
        strParentFolderPath = objFolder.ParentFolder.Path
    End If
        
    DetermineParentFolderPath = strParentFolderPath
    Exit Function
    
DetermineParentFolderPathErrorHandler:
    Debug.Print "Error in DetermineParentFolderPath: " & Err.Description
    Debug.Assert False
    
    LogErrors Err.Number, "DetermineParentFolderPath"
    DetermineParentFolderPath = strParentFolderPath
    
End Function

Public Sub FileOpenProc(ByVal hwndOwner As Long)
Dim strOpenFileName As String
Dim strFilter As String
On Error Resume Next

strFilter = "GEL Files(*.gel)" & Chr(0) & "*.gel" & Chr(0) _
            & "All files(*.*)" & Chr(0) & "*.*" & Chr(0)
strOpenFileName = OpenFileAPIDlg(hwndOwner, strFilter, 1, "Open File")
DoEvents                                    'give Windows chance to refresh(close dialog)
If Len(strOpenFileName) > 0 Then
   ReadGelFile strOpenFileName
Else
   MDIStatus False, ""
End If
End Sub

Public Function FileSaveProc(ByVal hwndOwner As Long, _
                             ByVal SuggestedFilename As String, _
                             ByRef SaveType As fstFileSaveTypeConstants, _
                             Optional ByRef PicSaveType As pftPictureFileTypeConstants = pftPNG) As String
'-------------------------------------------------------------------
'Display a Save As dialog box and return a filename; if user Cancels
'action (Cancel button or file exists and user does not want to
'overwrite existing file return empty string
'PicSaveType is important only when Save As Picture is used; then it
'returns type of picture that should be saved
'-------------------------------------------------------------------
Dim FltInd As Integer
Dim sSaveFileName As String

On Error Resume Next
FltInd = 1
Select Case SaveType
Case fstFileSaveTypeConstants.fstGel
'    sSaveFileName = SaveFileAPIDlg(hwndOwner, "Gel Files(*.gel)" & Chr(0) _
'        & "*.gel" & Chr(0), FltInd, SuggestedFilename, "Save File")
    
    ' MonroeMod Start
    sSaveFileName = SelectFile(hwndOwner, "Save File", "", True, SuggestedFilename, "Gel Files (*.gel)|*.gel")
    
    If Len(sSaveFileName) > 0 Then
        sSaveFileName = FileExtensionForce(sSaveFileName, "gel")
    End If
    FileSaveProc = sSaveFileName
    ' MonroeMod Finish
    
Case fstFileSaveTypeConstants.fstUMR
    FileSaveProc = SaveFileAPIDlg(hwndOwner, "Gel Files(*.umr)" & Chr(0) _
        & "*.umr" & Chr(0), FltInd, SuggestedFilename, "Save File")
Case fstFileSaveTypeConstants.fstPIC
    ' Default to PNG file (since nice and small, but looks good)
    FltInd = 1
    If LCase(Right(SuggestedFilename, 4)) = ".gel" Then SuggestedFilename = Left(SuggestedFilename, Len(SuggestedFilename) - 4)
    
    If InStr(SuggestedFilename, ":") > 0 Then
        ' Need to remove a colon from the filename
        SuggestedFilename = Replace(SuggestedFilename, ":", "")
    End If
    
    FileSaveProc = SaveFileAPIDlg(hwndOwner, "PNG file (*.png)" & Chr(0) & "*.png" & Chr(0) & _
                                             "JPEG file (*.jpeg)" & Chr(0) & "*.jpg" & Chr(0) & _
                                             "Windows Meta file (*.wmf)" & Chr(0) & "*.wmf" & Chr(0) & _
                                             "Enhanced Meta file (*.emf)" & Chr(0) & "*.emf" & Chr(0) & _
                                             "BMP file (*.bmp)" & Chr(0) & "*.bmp", _
                                             FltInd, SuggestedFilename, "Save Picture File")
    If Len(FileSaveProc) > 0 Then
       PicSaveType = FltInd
    Else
       PicSaveType = pftPictureFileTypeConstants.pftUnknown
    End If
Case Else
    ' Includes fstFileSaveTypeConstants.fstTxt
    FileSaveProc = SaveFileAPIDlg(hwndOwner, "Text files(*.txt)" & Chr(0) _
        & "*.txt" & Chr(0), FltInd, SuggestedFilename, "Save File")
End Select
End Function

' MonroeMod: Revised function
Private Function OnRecentFilesList(strFilePath As String) As Integer
' Looks in glbRecentFiles.Files() for the file with strFilePath
' If found, returns the index of the file in glbRecentFiles
' Otherwise, returns -1

Dim I As Integer

For I = 0 To glbRecentFiles.FileCount - 1
    If UCase(glbRecentFiles.Files(I).FullFilePath) = UCase(strFilePath) Then
       OnRecentFilesList = I
       Exit Function
    End If
Next I
OnRecentFilesList = -1

End Function

' MonreoMod: Additional parameters in function definition
Public Sub SaveFileAs(FileName As String, blnRemoveUMCData As Boolean, blnRemovePairsData As Boolean, lngGelIndex As Long, eFileSaveMode As fsFileSaveModeConstants)

If eFileSaveMode = fsLegacy Then
    BinarySaveLegacy FileName, lngGelIndex
Else
    ' MonroeMod: Use BinarySaveData to save the data
    BinarySaveData FileName, blnRemoveUMCData, blnRemovePairsData, lngGelIndex, GelData(lngGelIndex), eFileSaveMode
End If
    
End Sub

' MonroeMod: Updated function header
Public Sub UpdateFileMenu(FileName As String)
Dim intRetVal As Integer
On Error Resume Next    'not big deal if this fails
'Check if the open filename is already among the MRU files
'MonroeMod: intRetVal will be -1 if not found, >=0 if found
intRetVal = OnRecentFilesList(FileName)
WriteRecentFiles FileName, intRetVal
'Update the list of the MRU files
' Need to add a short delay here to give the Ini file time to update
Sleep 250
GetRecentFiles
End Sub

Public Function GetFileExtension(FileName As String, Optional blnIncludeLeadingPeriod As Boolean = True) As String
    ' Returns the extension of the file, optional including the leading period
    Dim fso As New FileSystemObject
    
    On Error GoTo GetFileExtensionErrorHandler
    
    If blnIncludeLeadingPeriod Then
        GetFileExtension = "." & fso.GetExtensionName(FileName)
    Else
        GetFileExtension = fso.GetExtensionName(FileName)
    End If
    
    Exit Function

GetFileExtensionErrorHandler:
    If blnIncludeLeadingPeriod Then
        GetFileExtension = "???"
    Else
        GetFileExtension = ".???"
    End If
End Function

Public Function GetFileInfo(FileName As String) As String
Dim tmp As String
Dim strExtension As String
On Error Resume Next
strExtension = GetFileExtension(FileName)
tmp = glCOMMENT_DATA_FILE_START & strExtension & " file): " & FileName
tmp = tmp & vbCrLf & "Size (" & strExtension & " file): " & Format(FileLen(FileName), "#,###,###,###") & " bytes"
tmp = tmp & vbCrLf & "Last modified (" & strExtension & " file): " & FileDateTime(FileName)
GetFileInfo = tmp
End Function

Public Sub SyncMenuCmdToolbar(ByVal bChecked As Boolean)
'synchronization of Toolbar menu command on all opened gels
Dim I As Integer
On Error Resume Next
For I = 1 To UBound(GelBody)
   If Not GelStatus(I).Deleted Then
      GelBody(I).mnuViewToolbar.Checked = bChecked
   End If
Next I
End Sub

Public Sub SyncMenuCmdTracker(ByVal bChecked As Boolean)
'synchronization of Toolbar menu command on all opened gels
Dim I As Integer
On Error Resume Next
For I = 1 To UBound(GelBody)
   If Not GelStatus(I).Deleted Then
      GelBody(I).mnuViewTracker.Checked = bChecked
   End If
Next I
End Sub

Public Sub StopTracking()
On Error Resume Next
SetTrackingLabels -1, glNoType, -1
glTracking = False
End Sub

Public Sub SetTrackingLabels(ByVal Ind As Long, ByVal DType As Integer, ByVal ID As Long)
Dim blnOutputToDebugWindow As Boolean, blnUseScientific As Boolean
Static IndexSaved As Long
Dim sFN As String, sMOverZ As String, sMW As String, sAbu As String, ser As String, sID As String
Dim sUMCIndices As String
On Error Resume Next

With GelData(Ind)
    
    Select Case DType
    Case glCSType
        sFN = .CSData(ID).ScanNumber
        sMOverZ = ""
        sMW = .CSData(ID).AverageMW
        sAbu = .CSData(ID).Abundance
        ser = GelDraw(Ind).CSER(ID)
        sID = .CSData(ID).MTID
        sUMCIndices = ConstructUMCIndexList(Ind, ID, DType)
    Case glIsoType
        sFN = .IsoData(ID).ScanNumber
        sMOverZ = .IsoData(ID).MZ
        sMW = GetIsoMass(.IsoData(ID), .Preferences.IsoDataField)
        sAbu = .IsoData(ID).Abundance
        ser = GelDraw(Ind).IsoER(ID)
        sID = .IsoData(ID).MTID
        sUMCIndices = ConstructUMCIndexList(Ind, ID, DType)
    Case Else
        sFN = ""
        sMOverZ = ""
        sMW = ""
        sAbu = ""
        ser = ""
        sID = ""
        sUMCIndices = ""
    End Select
End With

blnOutputToDebugWindow = False
blnUseScientific = (frmTracker.GetIntensityNotationMode() = nmScientific)
With frmTracker
  .lblFNTrack = sFN & Chr$(32)
  .lblMOverZTrack = Format$(sMOverZ, "##,###,#00.0000") & Chr$(32)
  .lblMWTrack = Format$(sMW, "##,###,#00.0000") & Chr$(32)
  If blnUseScientific Then
    .lblAbuTrack = Format$(sAbu, "Scientific") & Chr$(32)
  Else
    .lblAbuTrack = Trim$(sAbu) & Chr$(32)
  End If
  If Len(sMOverZ) > 0 Then
    If blnOutputToDebugWindow Then
        If ID <> IndexSaved Then
            Debug.Print "Scan " & sFN & ": " & Format$(sMOverZ, "00.0000") & "," & Format$(sMW, "00.0000") & "," & sAbu
            IndexSaved = ID
        End If
    End If
  End If
  
  If Len(ser) > 0 And IsNumeric(ser) Then
     If val(ser) < 0 Then
        .lblDRTrack = "N/A" & Chr$(32)
     Else
        If val(ser) = glHugeOverExp Then
           .lblDRTrack = Format$(ser, "Scientific") & Chr$(32)
        Else
           .lblDRTrack = Format$(ser, "Standard") & Chr$(32)
        End If
     End If
  Else
     .lblDRTrack = ""
  End If
  .lblIdentity = Chr$(32) & sID
  .txtIdentity = .lblIdentity
  .lblUMCIndex = sUMCIndices
End With
End Sub

Public Function OpenFileAPIDlg(ByVal Ownerhwnd As Long, _
                               ByVal sFilter As String, _
                               ByVal nFilterInd As Integer, _
                               ByVal sTitle As String) As String
'returns file name or zero-length string in case user canceled dialog
Dim ofDlg  As OPENFILENAME
Dim Res As Long
Dim Chr0Pos As Integer
ofDlg.lStructSize = Len(ofDlg)
ofDlg.hwndOwner = Ownerhwnd
ofDlg.hInstance = App.hInstance
ofDlg.lpstrFilter = sFilter
ofDlg.nFilterIndex = nFilterInd
ofDlg.lpstrFile = String(257, 0)
ofDlg.nMaxFile = Len(ofDlg.lpstrFile) - 1
ofDlg.lpstrFileTitle = ofDlg.lpstrFile
ofDlg.nMaxFileTitle = ofDlg.nMaxFile
ofDlg.lpstrInitialDir = CurDir
ofDlg.lpstrTitle = sTitle
ofDlg.flags = 0
Res = GetOpenFileName(ofDlg)
If Res = 0 Then
   OpenFileAPIDlg = ""
Else
   Chr0Pos = InStr(1, ofDlg.lpstrFile, Chr(0))
   If Chr0Pos > 0 Then
      OpenFileAPIDlg = Left$(ofDlg.lpstrFile, Chr0Pos - 1)
   Else
      OpenFileAPIDlg = Trim$(ofDlg.lpstrFile)
   End If
End If
End Function

Public Function SaveFileAPIDlg(ByVal Ownerhwnd As Long, _
                               ByVal sFilter As String, _
                               ByRef nFilterInd As Integer, _
                               ByVal sSuggestedFileName As String, _
                               ByVal sTitle As String) As String

Dim sfDlg  As OPENFILENAME
Dim Res As Long
Dim Chr0Pos As Integer
sfDlg.lStructSize = Len(sfDlg)
sfDlg.hwndOwner = Ownerhwnd
sfDlg.hInstance = App.hInstance
sfDlg.lpstrFilter = sFilter
sfDlg.nFilterIndex = nFilterInd
sfDlg.lpstrFile = sSuggestedFileName & String(257, 0)
sfDlg.nMaxFile = Len(sfDlg.lpstrFile) - 1
'sfDlg.lpstrDefExt = "*.gel"
sfDlg.lpstrFileTitle = sfDlg.lpstrFile
sfDlg.nMaxFileTitle = sfDlg.nMaxFile
sfDlg.lpstrInitialDir = CurDir
sfDlg.lpstrTitle = sTitle
sfDlg.flags = OFS_SAVEFILE_FLAGS    'ask if file already exists
Res = GetSaveFileName(sfDlg)
nFilterInd = sfDlg.nFilterIndex
If Res = 0 Then
   SaveFileAPIDlg = ""
Else
   Chr0Pos = InStr(1, sfDlg.lpstrFile, Chr(0))
   If Chr0Pos > 0 Then
      SaveFileAPIDlg = Left$(sfDlg.lpstrFile, Chr0Pos - 1)
   Else
      SaveFileAPIDlg = Trim$(sfDlg.lpstrFile)
   End If
End If
End Function


Public Sub PrinterSetupAPIDlg(ByVal Ownerhwnd As Long)
Dim PrtDlg As PrintDlg
Dim PrtMod As DEVMODE
Dim PrtNms As DEVNAMES
Dim lpMod As Long   'pointers to memory blocks allocated
Dim lpNms As Long   'for structures PrtMod and PrtNms
Dim Res As Long

'load default settings(relevant only) from the Printer object
PrtMod.dmDeviceName = Printer.DeviceName
PrtMod.dmSize = Len(PrtMod)
PrtMod.dmFields = DM_ORIENTATION
PrtMod.dmOrientation = DMORIENT_PORTRAIT
'load strings for default printer-see DEVNAMES structure for explanations
PrtNms.wDriverOffset = 8
PrtNms.wDeviceOffset = PrtNms.wDriverOffset + 1 + Len(Printer.DriverName)
PrtNms.wOutputOffset = PrtNms.wDeviceOffset + 1 + Len(Printer.Port)
PrtNms.wDefault = 0
PrtNms.extra = Printer.DriverName & vbNullChar & Printer.DeviceName _
             & vbNullChar & Printer.Port & vbNullChar
'now load initialization settings to the PRINTDLG structure
PrtDlg.lStructSize = Len(PrtDlg)
PrtDlg.hwndOwner = Ownerhwnd
PrtDlg.flags = PD_PRINTSETUP
'allocate memory block for DEVMODE structure inside the
'PRINTDLG structure and copy PrtMod data to allocated block
PrtDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(PrtMod))
lpMod = GlobalLock(PrtDlg.hDevMode)
CopyMemory ByVal lpMod, PrtMod, Len(PrtMod)
Res = GlobalUnlock(PrtDlg.hDevMode)
'allocate memory block for DEVNAMES structure inside the
'PRINTDLG structure and copy PrtNms data to allocated block
PrtDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(PrtNms))
lpNms = GlobalLock(PrtDlg.hDevNames)
CopyMemory ByVal lpNms, PrtNms, Len(PrtNms)
Res = GlobalUnlock(PrtDlg.hDevNames)
'open the dialog
Res = PrintDlg(PrtDlg)
'release memory
Res = GlobalFree(PrtDlg.hDevMode)
Res = GlobalFree(PrtDlg.hDevNames)
End Sub

Public Function InitDrawData(ByVal Ind As Long) As Boolean
'initialize data drawing structure; it is called only once
Dim I As Long
Dim IsoField As Integer
IsoField = GelData(Ind).Preferences.IsoDataField
On Error GoTo err_InitDrawData
With GelDraw(Ind)
    'initialize CS data
    .CSCount = GelData(Ind).CSLines
    If .CSCount > 0 Then
       ReDim .CSID(1 To .CSCount)
       ReDim .CSX(1 To .CSCount)
       ReDim .CSY(1 To .CSCount)
       ReDim .CSR(1 To .CSCount)
       ReDim .CSER(1 To .CSCount)
       ReDim .CSERClr(1 To .CSCount)
       .CSVisible = True
       For I = 1 To .CSCount
           .CSID(I) = I
       Next I
    Else
       .CSVisible = False
    End If
    'initialize Iso data
    .IsoCount = GelData(Ind).IsoLines
    If .IsoCount > 0 Then
       ReDim .IsoID(1 To .IsoCount)
       ReDim .IsoX(1 To .IsoCount)
       ReDim .IsoY(1 To .IsoCount)
       ReDim .IsoR(1 To .IsoCount)
       ReDim .IsoER(1 To .IsoCount)
       ReDim .IsoERClr(1 To .IsoCount)
       .IsoVisible = True
       For I = 1 To .IsoCount
           .IsoID(I) = I
       Next I
    Else
       .IsoVisible = False
    End If
End With
InitDrawData = True
err_InitDrawData:
End Function


Public Sub DestroyStructures(ByVal Ind As Long)
'release memory used by data structures
On Error Resume Next
With GelData(Ind)
     Erase .ScanInfo
     Erase .CSData
     Erase .IsoData
End With
With GelDraw(Ind)
  If .CSCount > 0 Then
     Erase .CSID:     Erase .CSX:      Erase .CSY
     Erase .CSR:      Erase .CSER:     Erase .CSERClr
     'this arrays might have not been used
     If Not IsArrayEmpty(.CSLogMW) Then Erase .CSLogMW
  End If
  If .IsoCount > 0 Then
     Erase .IsoID:     Erase .IsoX:     Erase .IsoY
     Erase .IsoR:      Erase .IsoER:    Erase .IsoERClr
     'this arrays might have not been used
     If Not IsArrayEmpty(.IsoLogMW) Then Erase .IsoLogMW
  End If
End With
'this might have not been used at all
With GelUMC(Ind)
   Erase .UMCs
   .UMCCnt = 0
End With
''With GelP(Ind)
''   If Not IsArrayEmpty(.P1) Then Erase .P1
''   If Not IsArrayEmpty(.P2) Then Erase .P2
''End With
Call DestroyDltLblPairs(Ind, False)

'Unused variable (March 2006); no longer supported
'If GelStatus(Ind).DBGel <> 0 Then Set GelDB(Ind) = Nothing

'analysis structure has to be preserved
If Not GelAnalysis(Ind) Is Nothing Then
   With GelAnalysis(Ind)
        .DestroyParameters
        .MTDB.DestroyDBStuff
        .MD_Reference_Job = -1
        Set .MTDB = Nothing
   End With
   Set GelAnalysis(Ind) = Nothing
End If
With GelLM(Ind)
   .CSCnt = 0:   .IsoCnt = 0
   Erase .CSFreqShift:   Erase .CSLckID:   Erase .CSMassCorrection
   Erase .IsoFreqShift:  Erase .IsoLckID:  Erase .IsoMassCorrection
End With

' Unused variable (August 2003)
'With GelIDP(Ind)
'   .Cnt = 0
'   .SyncWithDltLblPairs = False
'   Erase .PInd
'   Erase .PIDInd
'End With

' MonroeMod Begin
With GelUMCIon(Ind)
    .NETCount = 0
    ReDim .NetInd1(0)
    ReDim .NetInd2(0)
    ReDim .NetDist(0)
End With

With GelUMCNETAdjDef(Ind)
    .NETFormula = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(Ind))
    .NETTolIterative = 0.2                   '~20 pct
End With

With GelSearchDef(Ind)
    .AnalysisHistoryCount = 0
    ReDim .AnalysisHistory(0)
End With

With GelDataLookupArrays(Ind)
    ReDim .CSUMCs(0)
    ReDim .IsoUMCs(0)
End With

' No longer supported (March 2006)
''With GelORFData(Ind)
''    .ORFCount = 0
''    Erase .Orfs
''    With .Definition
''        .MTDBConnectionString = ""
''        .MTDBName = ""
''        .ORFDBConnectionString = ""
''        .ORFDBName = ""
''        .DateDataObtained = ""
''        .Organism = ""
''        .OtherInfo = ""
''    End With
''End With
''
''With GelORFMassTags(Ind)
''    .ORFCount = 0
''    Erase .Orfs
''End With
' MonroeMod Finish

'RemoveDisplayAllFromOverlay (Ind)          'remove this display from overlay
With GelUMCDraw(Ind)
   .Count = 0
   .Visible = False
   Erase .ClassID
   Erase .X1:           Erase .Y1
   Erase .x2:           Erase .Y2
End With
End Sub

Public Sub InitDrawER(ByVal Ind As Long)
'this procedure is called only when building graph
Dim I As Long
Dim dblER As Double
On Error GoTo InitDrawERErrorHandler

With GelData(Ind)
  If .CSLines > 0 Then
     For I = 1 To .CSLines
        dblER = LookupExpressionRatioValue(Ind, I, False, -1)
        If Abs(dblER) > 1E+38 Then
            If dblER > 0 Then dblER = 1E+38 Else dblER = -1E+38
        End If
            
        GelDraw(Ind).CSER(I) = dblER
     Next I
  End If
  If .IsoLines > 0 Then
     For I = 1 To .IsoLines
        dblER = LookupExpressionRatioValue(Ind, I, True, -1)
        If Abs(dblER) > 1E+38 Then
            If dblER > 0 Then dblER = 1E+38 Else dblER = -1E+38
        End If
         
        GelDraw(Ind).IsoER(I) = dblER
     Next I
  End If
End With
InitDrawERColors Ind
Exit Sub

InitDrawERErrorHandler:
Debug.Assert False
Debug.Print "Error in InitDrawER: " & Err.Description
End Sub

Public Sub InitDrawERColors(ByVal Ind As Long)
'--------------------------------------------------------------------
'this is called whenever differential display is requested
'--------------------------------------------------------------------
Dim I As Long
On Error Resume Next
With GelDraw(Ind)
  Select Case GelData(Ind).Preferences.DRDefinition
  Case glNormal
    If .CSCount > 0 Then
      For I = 1 To .CSCount
          If .CSER(I) >= 0 Then
             .CSERClr(I) = GetERClrInd(.CSER(I))
          Else
             .CSERClr(I) = glDONT_DISPLAY
          End If
      Next I
    End If
    If .IsoCount > 0 Then
      For I = 1 To .IsoCount
          If .IsoER(I) >= 0 Then
             .IsoERClr(I) = GetERClrInd(.IsoER(I))
          Else
             .IsoERClr(I) = glDONT_DISPLAY
          End If
      Next I
    End If
  Case glReverse
    If .CSCount > 0 Then
      For I = 1 To .CSCount
          If .CSER(I) >= 0 Then
             .CSERClr(I) = -GetERClrInd(.CSER(I))
          Else
             .CSERClr(I) = glDONT_DISPLAY
          End If
      Next I
    End If
    If .IsoCount > 0 Then
      For I = 1 To .IsoCount
          If .IsoER(I) >= 0 Then
             .IsoERClr(I) = -GetERClrInd(.IsoER(I))
          Else
             .IsoERClr(I) = glDONT_DISPLAY
          End If
      Next I
    End If
  End Select
End With
End Sub


Public Sub InitDrawChargeStateMap(ByVal Ind As Long)
'--------------------------------------------------------------------
'this is called whenever charge state map display is requested
'--------------------------------------------------------------------
Dim I As Long
With GelData(Ind)
    If .CSLines > 0 Then
       For I = 1 To .CSLines
           GelDraw(Ind).CSERClr(I) = 50 + GetChargeStateMapIndex(.CSData(I).Charge)
       Next I
    End If
    If .IsoLines > 0 Then
       For I = 1 To .IsoLines
           GelDraw(Ind).IsoERClr(I) = 50 + GetChargeStateMapIndex(.IsoData(I).Charge)
       Next I
    End If
End With
End Sub


Public Sub InitDrawCSLogMW(ByVal Ind As Long)
Dim MW As Double                    'can not come here if not CSCount>0
Dim I As Long
With GelDraw(Ind)
     For I = 1 To .CSCount
         MW = GelData(Ind).CSData(I).AverageMW
         If MW > 0 Then
            .CSLogMW(I) = CSng(Log(MW) / Log(10#))
         Else
            .CSLogMW(I) = -glHugeOverExp
         End If
     Next I
End With
End Sub

Public Sub InitDrawIsoLogMW(ByVal Ind As Long)
Dim MW As Double                    'can not come here if not IsoCount>0
Dim I As Long
With GelDraw(Ind)
     For I = 1 To .IsoCount
         MW = GetIsoMass(GelData(Ind).IsoData(I), GelData(Ind).Preferences.IsoDataField)
         If MW > 0 Then
            .IsoLogMW(I) = CSng(Log(MW) / Log(10#))
         Else
            .IsoLogMW(I) = -glHugeOverExp
         End If
     Next I
End With
End Sub

Public Sub GetHotSpot(ByVal Ind As Long, ByVal lx As Long, ByVal ly As Long, _
                                        HotID As Long, HotType As Integer)
'--------------------------------------------------------------------------------
'hot spot around LX,LY is defined as rectangle regardless of actual spot shape
'this is arranged for performance reasons
'--------------------------------------------------------------------------------
Dim ar As Double
Dim I As Long
ar = GelData(Ind).Preferences.AbuAspectRatio
HotType = glNoType
Select Case GelBody(Ind).fgDisplay
Case glNormalDisplay
  Select Case GelBody(Ind).fgZOrder
  Case glCSOnTop
  'look first among CS; if not there check among ISO
  'go with reverse loop to find always the spot on top
    With GelDraw(Ind)
      If .CSCount > 0 And .CSVisible Then
         For I = .CSCount To 1 Step -1
           If .CSID(I) > 0 And .CSR(I) > 0 Then 'search only among visible
              If (Abs(lx - .CSX(I)) < .CSR(I) / 2) And _
                 (Abs(ly - .CSY(I)) < .CSR(I) / (2 * ar)) Then
                 HotType = glCSType
                 HotID = I
                 Exit For
              End If
           End If
         Next I
      End If
      If HotType = glNoType Then
         If .IsoCount > 0 And .IsoVisible Then
            For I = .IsoCount To 1 Step -1
              If .IsoID(I) > 0 And .IsoR(I) > 0 Then 'search only among visible
                 If (Abs(lx - .IsoX(I)) < .IsoR(I) / 2) And _
                    (Abs(ly - .IsoY(I)) < .IsoR(I) / (2 * ar)) Then
                    HotType = glIsoType
                    HotID = I
                    Exit For
                 End If
              End If
           Next I
         End If
      End If
    End With
  Case glIsoOnTop
  'look first among Iso; if not there check among CS
    With GelDraw(Ind)
      If .IsoCount > 0 And .IsoVisible Then
         For I = .IsoCount To 1 Step -1
           If .IsoID(I) > 0 And .IsoR(I) > 0 Then 'search only among visible
              If (Abs(lx - .IsoX(I)) < .IsoR(I) / 2) And _
                 (Abs(ly - .IsoY(I)) < .IsoR(I) / (2 * ar)) Then
                 HotType = glIsoType
                 HotID = I
                 Exit For
              End If
           End If
         Next I
      End If
      If HotType = glNoType Then
         If .CSCount > 0 And .CSVisible Then
            For I = .CSCount To 1 Step -1
              If .CSID(I) > 0 And .CSR(I) > 0 Then 'search only among visible
                 If (Abs(lx - .CSX(I)) < .CSR(I) / 2) And _
                    (Abs(ly - .CSY(I)) < .CSR(I) / (2 * ar)) Then
                    HotType = glCSType
                    HotID = I
                    Exit For
                 End If
              End If
            Next I
         End If
       End If
    End With
  End Select
Case glDifferentialDisplay, glChargeStateMapDisplay
  Select Case GelBody(Ind).fgZOrder
  Case glCSOnTop
  'look first among CS; if not there check among ISO
  'go with reverse loop to find always the spot on top
    With GelDraw(Ind)
      If .CSCount > 0 And .CSVisible Then
         For I = .CSCount To 1 Step -1
           If .CSID(I) > 0 And .CSR(I) > 0 And .CSERClr(I) <> glDONT_DISPLAY Then
              If (Abs(lx - .CSX(I)) < .CSR(I) / 2) And _
                 (Abs(ly - .CSY(I)) < .CSR(I) / (2 * ar)) Then
                 HotType = glCSType
                 HotID = I
                 Exit For
              End If
           End If
         Next I
      End If
      If HotType = glNoType Then
         If .IsoCount > 0 And .IsoVisible Then
            For I = .IsoCount To 1 Step -1
              If .IsoID(I) > 0 And .IsoR(I) > 0 And .IsoERClr(I) <> glDONT_DISPLAY Then
                 If (Abs(lx - .IsoX(I)) < .IsoR(I) / 2) And _
                    (Abs(ly - .IsoY(I)) < .IsoR(I) / (2 * ar)) Then
                    HotType = glIsoType
                    HotID = I
                    Exit For
                 End If
              End If
           Next I
         End If
      End If
    End With
  Case glIsoOnTop
  'look first among Iso; if not there check among CS
    With GelDraw(Ind)
      If .IsoCount > 0 And .IsoVisible Then
         For I = .IsoCount To 1 Step -1
           If .IsoID(I) > 0 And .IsoR(I) > 0 And .IsoERClr(I) <> glDONT_DISPLAY Then
              If (Abs(lx - .IsoX(I)) < .IsoR(I) / 2) And _
                 (Abs(ly - .IsoY(I)) < .IsoR(I) / (2 * ar)) Then
                 HotType = glIsoType
                 HotID = I
                 Exit For
              End If
           End If
         Next I
      End If
      If HotType = glNoType Then
         If .CSCount > 0 And .CSVisible Then
            For I = .CSCount To 1 Step -1
              If .CSID(I) > 0 And .CSR(I) > 0 And .CSERClr(I) <> glDONT_DISPLAY Then
                 If (Abs(lx - .CSX(I)) < .CSR(I) / 2) And _
                    (Abs(ly - .CSY(I)) < .CSR(I) / (2 * ar)) Then
                    HotType = glCSType
                    HotID = I
                    Exit For
                 End If
              End If
            Next I
         End If
       End If
    End With
  End Select
End Select
End Sub

Public Sub GelCSIncludeAll(ByVal Ind As Long)
Dim I As Long
On Error Resume Next
With GelDraw(Ind)
    For I = 1 To .CSCount
        .CSID(I) = Abs(.CSID(I))
    Next I
End With
End Sub

Public Sub GelCSInvertVisible(ByVal Ind As Long)
Dim I As Long
On Error Resume Next
With GelDraw(Ind)
    For I = 1 To .CSCount
        .CSID(I) = -.CSID(I)
    Next I
End With
End Sub

Public Sub GelCSExcludeAll(ByVal Ind As Long)
Dim I As Long
On Error Resume Next
With GelDraw(Ind)
    For I = 1 To .CSCount
        .CSID(I) = -Abs(.CSID(I))
    Next I
End With
End Sub

Public Sub GelIsoIncludeAll(ByVal Ind As Long)
Dim I As Long
On Error Resume Next
With GelDraw(Ind)
    For I = 1 To .IsoCount
        .IsoID(I) = Abs(.IsoID(I))
    Next I
End With
End Sub

Public Sub GelIsoInvertVisible(ByVal Ind As Long)
Dim I As Long
On Error Resume Next
With GelDraw(Ind)
    For I = 1 To .IsoCount
        .IsoID(I) = -.IsoID(I)
    Next I
End With
End Sub

Public Sub GelIsoExcludeAll(ByVal Ind As Long)
Dim I As Long
On Error Resume Next
With GelDraw(Ind)
    For I = 1 To .IsoCount
        .IsoID(I) = -Abs(.IsoID(I))
    Next I
End With
End Sub

Public Sub GelCSExcludeAbuRange(ByVal Ind As Long)
'exclude CS data out of [MinAbu,MaxAbu] range
Dim MinAbu As Double
Dim MaxAbu As Double
Dim I As Long
With GelData(Ind)
    MinAbu = CDbl(.DataFilter(fltCSAbu, 1))
    MaxAbu = CDbl(.DataFilter(fltCSAbu, 2))
    If .CSLines > 0 Then
       For I = 1 To .CSLines
           If .CSData(I).Abundance < MinAbu Or .CSData(I).Abundance > MaxAbu Then
              GelDraw(Ind).CSID(I) = -Abs(GelDraw(Ind).CSID(I))
           End If
       Next I
    End If
End With
End Sub

Public Sub GelIsoExcludeAbuRange(ByVal Ind As Long)
'exclude Iso data out of [MinAbu,MaxAbu] range
Dim I As Long
Dim MinAbu As Double
Dim MaxAbu As Double
With GelData(Ind)
    MinAbu = CDbl(.DataFilter(fltIsoAbu, 1))
    MaxAbu = CDbl(.DataFilter(fltIsoAbu, 2))
    If .IsoLines > 0 Then
       For I = 1 To .IsoLines
           If .IsoData(I).Abundance < MinAbu Or .IsoData(I).Abundance > MaxAbu Then
              GelDraw(Ind).IsoID(I) = -Abs(GelDraw(Ind).IsoID(I))
           End If
       Next I
    End If
End With
End Sub

Public Sub GelCSExcludeMWRange(ByVal Ind As Long)
'-----------------------------------------------------------------
'exclude CS data out of [MinMW,MaxMW] range
'-----------------------------------------------------------------
Dim MinMW As Double, MaxMW As Double
Dim I As Long
On Error Resume Next
With GelData(Ind)
    MinMW = CDbl(.DataFilter(fltCSMW, 1))
    MaxMW = CDbl(.DataFilter(fltCSMW, 2))
    For I = 1 To .CSLines
        If ((.CSData(I).AverageMW < MinMW) Or (.CSData(I).AverageMW > MaxMW)) Then
           GelDraw(Ind).CSID(I) = -Abs(GelDraw(Ind).CSID(I))
        End If
    Next I
End With
End Sub

Public Sub GelIsoExcludeMWRange(ByVal Ind As Long)
'------------------------------------------------------------------
'exclude Iso data out of [MinMW,MaxMW] range
'------------------------------------------------------------------
Dim I As Long
Dim MinMW As Double, MaxMW As Double
On Error GoTo GelIsoExcludeMWRangeErrorHandler
With GelData(Ind)
    MinMW = CDbl(.DataFilter(fltIsoMW, 1))
    MaxMW = CDbl(.DataFilter(fltIsoMW, 2))
    For I = 1 To .IsoLines
        If ((GetIsoMass(.IsoData(I), .Preferences.IsoDataField) < MinMW) Or _
            (GetIsoMass(.IsoData(I), .Preferences.IsoDataField) > MaxMW)) Then
                GelDraw(Ind).IsoID(I) = -Abs(GelDraw(Ind).IsoID(I))
        End If
    Next I
End With
Exit Sub

GelIsoExcludeMWRangeErrorHandler:
Debug.Print "Error in GelIsoExcludeMWRange: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "GelIsoExcludeMWRange"
End Sub

Public Sub GelIsoExcludeCSRange(ByVal Ind As Long)
'------------------------------------------------------------------
'exclude Iso data out of [MinCS,MaxCS] range
'------------------------------------------------------------------
Dim I As Long
Dim MinCS As Double, MaxCS As Double
On Error GoTo GelIsoExcludeCSRangeErrorHandler
With GelData(Ind)
    MinCS = CDbl(.DataFilter(fltIsoCS, 1))
    MaxCS = CDbl(.DataFilter(fltIsoCS, 2))
    For I = 1 To .IsoLines
        If ((.IsoData(I).Charge < MinCS) Or (.IsoData(I).Charge > MaxCS)) Then
            GelDraw(Ind).IsoID(I) = -Abs(GelDraw(Ind).IsoID(I))
        End If
    Next I
End With
Exit Sub

GelIsoExcludeCSRangeErrorHandler:
Debug.Print "Error in GelIsoExcludeCSRange: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "GelIsoExcludeCSRange"
End Sub

Public Sub GelIsoExcludeMZRange(ByVal Ind As Long)
'------------------------------------------------------------------
'exclude Iso data out of [MinMZ,MaxMZ] range
'------------------------------------------------------------------
Dim I As Long
Dim MinMZ As Double, MaxMZ As Double
Dim intCharge As Integer
Dim TestMZ As Double
On Error GoTo GelIsoExcludeMZRangeErrorHandler
With GelData(Ind)
    MinMZ = CDbl(.DataFilter(fltIsoMZ, 1))
    MaxMZ = CDbl(.DataFilter(fltIsoMZ, 2))
    For I = 1 To .IsoLines
        intCharge = val(.IsoData(I).Charge)
        
        If intCharge > 0 Then
            TestMZ = (GetIsoMass(.IsoData(I), .Preferences.IsoDataField) + intCharge) / intCharge
            If ((TestMZ < MinMZ) Or (TestMZ > MaxMZ)) Then
                GelDraw(Ind).IsoID(I) = -Abs(GelDraw(Ind).IsoID(I))
            End If
        Else
            ' Charge is 0; error may have occurred while loading the PEK/CSV/mzXML/mzData file
            ' Or, the PEK/CSV/mzXML/mzData file could be wrong
            Debug.Assert False
        End If
    Next I
End With
Exit Sub

GelIsoExcludeMZRangeErrorHandler:
Debug.Print "Error in GelIsoExcludeMZRange: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "GelIsoExcludeMZRange"
End Sub

Public Sub GelExcludeEvenOddScans(ByVal Ind As Long)
Dim I As Long
Dim intScan As Integer
Dim intEvenOddModCompareVal As Integer

On Error GoTo GelIsoExcludeEvenOddErrorHandler
With GelData(Ind)
    If .DataFilter(fltEvenOddScanNumber, 0) And .DataFilter(fltEvenOddScanNumber, 1) > 0 Then
        If .DataFilter(fltEvenOddScanNumber, 1) = 1 Then
            intEvenOddModCompareVal = 0
        Else
            intEvenOddModCompareVal = 1
        End If
    
        For I = 1 To .CSLines
            intScan = val(.CSData(I).ScanNumber)
            
            ' Use Modulo division to check if odd or even
            If intScan Mod 2 = intEvenOddModCompareVal Then
                GelDraw(Ind).CSID(I) = -Abs(GelDraw(Ind).CSID(I))
            End If
        Next I
    
        For I = 1 To .IsoLines
            intScan = val(.IsoData(I).ScanNumber)
            
            ' Use Modulo division to check if odd or even
            If intScan Mod 2 = intEvenOddModCompareVal Then
                GelDraw(Ind).IsoID(I) = -Abs(GelDraw(Ind).IsoID(I))
            End If
        Next I
    End If
End With
Exit Sub

GelIsoExcludeEvenOddErrorHandler:
Debug.Print "Error in GelExcludeEvenOddScans: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "GelExcludeEvenOddScans"
End Sub


Public Sub GelCSExcludeER(ByVal Ind As Long)
Dim ERExclusionOption As Integer
Dim I As Long
Dim PartSum As Integer
If GelData(Ind).CSLines > 0 Then
   ERExclusionOption = CInt(GelData(Ind).DataFilter(fltAR, 0))
   If ERExclusionOption < 0 Then
      GelCSExcludeAll (Ind)
      Exit Sub
   End If

   Select Case ERExclusionOption
   Case 0              'should not happen here; do nothing
   Case 1, 5, 9, 13    'exclude data with Expression Ratio
        CSExcludeERBase Ind, 1
   Case Else
        PartSum = ERExclusionOption
        I = 4
        Do While PartSum > 0
           If PartSum >= 2 ^ I Then
              PartSum = PartSum - 2 ^ I
              CSExcludeERBase Ind, 2 ^ I
           End If
           I = I - 1
        Loop
   End Select
End If
End Sub

Private Sub CSExcludeERBase(ByVal Ind As Long, ByVal EROption As Integer)
Dim ERMin As Double         'this is never going to be called if not CSLines>0
Dim ERMax As Double
Dim I As Long
On Error Resume Next
With GelData(Ind)
   ERMin = CDbl(.DataFilter(fltAR, 1))
   ERMax = CDbl(.DataFilter(fltAR, 2))
End With
With GelDraw(Ind)
   Select Case EROption
   Case 1          'exclude with ER
        For I = 1 To .CSCount
            If .CSER(I) >= 0 Then
               .CSID(I) = -Abs(.CSID(I))
            End If
        Next I
   Case 2          'exclude without ER
        For I = 1 To .CSCount
            If .CSER(I) < 0 Then
               .CSID(I) = -Abs(.CSID(I))
            End If
        Next I
   Case 4          'exclude Huge Underexpressed
        For I = 1 To .CSCount
            If .CSER(I) = glHugeUnderExp Then
               .CSID(I) = -Abs(.CSID(I))
            End If
        Next I
   Case 8          'exclude Huge Overexpressed
        For I = 1 To .CSCount
            If .CSER(I) = glHugeOverExp Then
               .CSID(I) = -Abs(.CSID(I))
            End If
        Next I
    Case 16         'exclude by ER range
        If ERMin <= 0 And ERMax >= 0 Then
           For I = 1 To .CSCount
               If .CSER(I) > ERMax Then
                  .CSID(I) = -Abs(.CSID(I))
               End If
           Next I
        ElseIf ERMin >= 0 And ERMax < 0 Then
           For I = 1 To .CSCount
               If .CSER(I) < ERMin Then
                  .CSID(I) = -Abs(.CSID(I))
               End If
           Next I
        ElseIf ERMin < ERMax And ERMax > 0 Then
           For I = 1 To .CSCount
               If (.CSER(I) < ERMin) Or (.CSER(I) > ERMax) Then
                  .CSID(I) = -Abs(.CSID(I))
               End If
           Next I
        ElseIf ERMin = ERMax And ERMax >= 0 Then
           For I = 1 To .CSCount
               If (.CSER(I) = ERMax) Then
                  .CSID(I) = -Abs(.CSID(I))
               End If
           Next I
        ElseIf ERMin < 0 And ERMax < 0 Then
            GelCSExcludeAll Ind
        End If
    End Select
End With
End Sub


Public Sub GelIsoExcludeER(ByVal Ind As Long)
Dim ERExclusionOption As Integer
Dim I As Long
Dim PartSum As Integer

If GelData(Ind).IsoLines > 0 Then
   ERExclusionOption = CInt(GelData(Ind).DataFilter(fltAR, 0))
   If ERExclusionOption < 0 Then
      GelIsoExcludeAll (Ind)
      Exit Sub
   End If

   Select Case ERExclusionOption
   Case 0              'should not happen here; do nothing
   Case 1, 5, 9, 13    'exclude data with Expression Ratio
        IsoExcludeERBase Ind, 1
   Case Else
        PartSum = ERExclusionOption
        I = 4
        Do While PartSum > 0
           If PartSum >= 2 ^ I Then
              PartSum = PartSum - 2 ^ I
              IsoExcludeERBase Ind, 2 ^ I
           End If
           I = I - 1
        Loop
   End Select
End If
End Sub

Private Sub IsoExcludeERBase(ByVal Ind As Long, ByVal EROption As Integer)
Dim ERMin As Double             'this is never going to be called if not IsoLines>0
Dim ERMax As Double
Dim I As Long
On Error Resume Next
With GelData(Ind)
   ERMin = CDbl(.DataFilter(fltAR, 1))
   ERMax = CDbl(.DataFilter(fltAR, 2))
End With
With GelDraw(Ind)
   Select Case EROption
   Case 1          'exclude with ER
        For I = 1 To .IsoCount
            If .IsoER(I) >= 0 Then
               .IsoID(I) = -Abs(.IsoID(I))
            End If
        Next I
   Case 2          'exclude without ER
        For I = 1 To .IsoCount
            If .IsoER(I) < 0 Then
               .IsoID(I) = -Abs(.IsoID(I))
            End If
        Next I
   Case 4          'exclude Huge Underexpressed
        For I = 1 To .IsoCount
            If .IsoER(I) = glHugeUnderExp Then
               .IsoID(I) = -Abs(.IsoID(I))
            End If
        Next I
   Case 8          'exclude Huge Overexpressed
        For I = 1 To .IsoCount
            If .IsoER(I) = glHugeOverExp Then
               .IsoID(I) = -Abs(.IsoID(I))
            End If
        Next I
    Case 16         'exclude by ER range
        If ERMin <= 0 And ERMax >= 0 Then
           For I = 1 To .IsoCount
               If .IsoER(I) > ERMax Then
                  .IsoID(I) = -Abs(.IsoID(I))
               End If
           Next I
        ElseIf ERMin >= 0 And ERMax < 0 Then
           For I = 1 To .IsoCount
               If .IsoER(I) < ERMin Then
                  .IsoID(I) = -Abs(.IsoID(I))
               End If
           Next I
        ElseIf ERMin < ERMax And ERMax > 0 Then
           For I = 1 To .IsoCount
               If (.IsoER(I) < ERMin) Or (.IsoER(I) > ERMax) Then
                  .IsoID(I) = -Abs(.IsoID(I))
               End If
           Next I
        ElseIf ERMin = ERMax And ERMax >= 0 Then
           For I = 1 To .IsoCount
               If (.IsoER(I) = ERMax) Then
                  .IsoID(I) = -Abs(.IsoID(I))
               End If
           Next I
        ElseIf ERMin < 0 And ERMax < 0 Then
            GelIsoExcludeAll Ind
        End If
    End Select
End With
End Sub

Public Sub GelCSExcludeIdentity(ByVal Ind As Long)
Dim IdentityOption As Integer
Dim I As Long
With GelData(Ind)
     IdentityOption = CInt(.DataFilter(fltID, 1))
     Select Case IdentityOption
     Case 0         'should not happen
     Case 1         'exclude identified
        For I = 1 To .CSLines
            If Not AllUnidentified(.CSData(I).MTID) Then
               GelDraw(Ind).CSID(I) = -Abs(GelDraw(Ind).CSID(I))
            End If
        Next I
     Case 2         'exclude unidentified
        For I = 1 To .CSLines
            If AllUnidentified(.CSData(I).MTID) Then
               GelDraw(Ind).CSID(I) = -Abs(GelDraw(Ind).CSID(I))
            End If
        Next I
     Case Else      'exclude all
        GelIsoExcludeAll Ind
     End Select
End With
End Sub

Public Sub GelIsoExcludeIdentity(ByVal Ind As Long)
Dim IdentityOption As Integer
Dim I As Long
With GelData(Ind)
     IdentityOption = CInt(.DataFilter(fltID, 1))
     Select Case IdentityOption
     Case 0         'should not happen
     Case 1         'exclude identified
        For I = 1 To .IsoLines
            If Not AllUnidentified(.IsoData(I).MTID) Then
               GelDraw(Ind).IsoID(I) = -Abs(GelDraw(Ind).IsoID(I))
            End If
        Next I
     Case 2         'exclude unidentified
        For I = 1 To .IsoLines
            If AllUnidentified(.IsoData(I).MTID) Then
               GelDraw(Ind).IsoID(I) = -Abs(GelDraw(Ind).IsoID(I))
            End If
        Next I
     Case Else      'exclude all
        GelIsoExcludeAll Ind
     End Select
End With
End Sub

Private Function AllUnidentified(NoIDWhat As Variant) As Boolean
'returns True if NoIDWhat string is empty or each non-empty
'line contains NoHarvest string, otherwise returns False
Dim sNoIDWhat As String
Dim iStartPos As Long
Dim iEndPos As Long
Dim sLine As String
On Error GoTo err_AllUnidentified

If IsNull(NoIDWhat) Then
    sNoIDWhat = ""
Else
    sNoIDWhat = CStr(NoIDWhat)
End If

If Len(sNoIDWhat) > 0 Then
    iStartPos = 1
    iEndPos = 0
    Do While iStartPos < Len(sNoIDWhat)
       iEndPos = InStr(iStartPos, sNoIDWhat, glARG_SEP)
       If iEndPos > iStartPos Then
          sLine = Trim$(Mid$(sNoIDWhat, iStartPos, iEndPos - iStartPos))
          iStartPos = iEndPos + 1
       Else
          sLine = Trim$(Right$(sNoIDWhat, Len(sNoIDWhat) - iStartPos + 1))
          iStartPos = Len(sNoIDWhat)
       End If
       If Len(sLine) > 0 Then
          If InStr(1, sLine, NoHarvest) <= 0 Then
             AllUnidentified = False
             Exit Function
          End If
       End If
    Loop
End If

err_AllUnidentified:
AllUnidentified = True
End Function

Public Function ReadGelFile(ByVal strFileName As String, Optional ByRef lngGelIndexToForce As Long = 0) As Long
' If lngGelIndexToForce is > 0 then the data will be loaded into the gel with the given index;
'  otherwise, the next available index will be used
'
' Returns the new gel index if success, 0 if failure

Dim fIndex As Long
Dim eResponse As VbMsgBoxResult
Dim eFileSaveMode As fsFileSaveModeConstants

Dim blnSuccess As Boolean
Dim strMessage As String

On Error Resume Next

blnSuccess = False

If lngGelIndexToForce <= 0 Then
    fIndex = FindFreeIndex()
Else
    fIndex = lngGelIndexToForce
End If

If fIndex > glMaxGels Then
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Command aborted. Too many open files.", vbOKOnly, glFGTU
    End If
    ReadGelFile = 0
    Exit Function
End If

Screen.MousePointer = vbHourglass
Select Case GetGelCertificate(strFileName)
Case glCERT1999
    strMessage = "You have selected an old file format that is no longer supported.  Unable to open the file: " & strFileName
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox strMessage, vbInformation + vbOKOnly, glFGTU
    Else
        AddToAnalysisHistory fIndex, strMessage
    End If
    GoTo FailedReadGelFile

'' No longer supported (March 2006)
''    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''        eResponse = MsgBox("Old file format. Update file to the new file format?", vbYesNo, glFGTU)
''    Else
''        eResponse = vbYes
''    End If
''    If eResponse <> vbYes Then GoTo exit_ReadGelFile
''    If Not ReadGelData1999(strFileName, fIndex) Then
''       GoTo FailedReadGelFile
''    Else
''       GelStatus(fIndex).Dirty = True
''       blnSuccess = True
''    End If
    
Case glCERT2000             'still try first to read new file format!
    strMessage = "You have selected an old file format that is no longer supported.  Unable to open the file: " & strFileName
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox strMessage, vbInformation + vbOKOnly, glFGTU
    Else
        AddToAnalysisHistory fIndex, strMessage
    End If
    GoTo FailedReadGelFile

'' No longer supported (March 2006)
''
''    If Not ReadGelData2003(strFileName, fIndex, False) Then
''        If Not ReadGelData2000(strFileName, fIndex) Then
''            GoTo FailedReadGelFile
''        Else
''            GelStatus(fIndex).Dirty = True
''            blnSuccess = True
''        End If
''    End If
    
Case glCERT2003
    If Not ReadGelData2003(strFileName, fIndex, True) Then
        GoTo FailedReadGelFile
    Else
        GelStatus(fIndex).Dirty = False
        blnSuccess = True
    End If
    
Case glCERT2000_DB, glCERT2002_MT
    strMessage = "You have selected an old file format that is no longer supported.  Unable to open the file: " & strFileName
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox strMessage, vbInformation + vbOKOnly, glFGTU
    Else
        AddToAnalysisHistory fIndex, strMessage
    End If
    GoTo FailedReadGelFile

'' No longer supported (March 2006)
''
''    Dim sDBGelType As String
''    If Not ReadGelData2000(strFileName, fIndex) Then
''        GoTo FailedReadGelFile
''    Else
''        sDBGelType = GetTagValueFromText(GelData(fIndex).Comment, glCOMMENT_DBGEL)
''        If IsNumeric(sDBGelType) Then
''            GelStatus(fIndex).DBGel = CLng(sDBGelType)
''            If GelStatus(fIndex).DBGel < 0 Then
''                GelStatus(fIndex).DBGel = glDBGEL_ERROR
''                strMessage = "Error determining the type of DB originated gel."
''                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''                    MsgBox strMessage, vbOKOnly
''                End If
''                AddToAnalysisHistory fIndex, strMessage
''            End If
''        Else
''            Select Case sDBGelType
''            Case glCOMMENT_DBGEL_ORF
''                GelStatus(fIndex).DBGel = glDBGEL_ORF
''            Case glCOMMENT_DBGEL_AMT
''                GelStatus(fIndex).DBGel = glDBGEL_AMT
''            Case Else
''                GelStatus(fIndex).DBGel = glDBGEL_ERROR
''                strMessage = "Error determining the type of DB originated gel."
''                If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''                    MsgBox strMessage, vbOKOnly
''                End If
''                AddToAnalysisHistory fIndex, strMessage
''            End Select
''        End If
''
''        GelStatus(fIndex).DBGel = True
''        GelStatus(fIndex).Dirty = False
''        blnSuccess = True
''
''        If Not ConnectToFTICR_AMT(GelDB(fIndex), GelData(fIndex).PathtoDatabase, False) Then
''            strMessage = "Error accessing database file behind the gel; some functions might not be available."
''            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
''                MsgBox strMessage, vbOKOnly, "Loading ORF gel"
''            End If
''            AddToAnalysisHistory fIndex, strMessage
''        End If
''    End If
''
Case glCERT2003_Modular
    If Not BinaryLoadData(strFileName, fIndex, eFileSaveMode) Then
        GoTo FailedReadGelFile
    Else
        GelStatus(fIndex).Dirty = False
        GelBody(fIndex).mFileSaveMode = eFileSaveMode
        blnSuccess = True
    End If
    
Case glCERT_FileNotFound
    strMessage = "File not found: " & strFileName
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox strMessage, vbOKOnly, glFGTU
    Else
        AddToAnalysisHistory fIndex, strMessage
    End If
    
    RemoveFromRecentFiles strFileName
    GetRecentFiles
    GoTo FailedReadGelFile
    
Case Else
    strMessage = "Unrecognized file format.  Cannot open the file: " & strFileName
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox strMessage, vbInformation + vbOKOnly, glFGTU
    Else
        AddToAnalysisHistory fIndex, strMessage
    End If
    GoTo FailedReadGelFile
End Select

Debug.Assert Not GelStatus(fIndex).Deleted
GelBody(fIndex).Tag = fIndex
GelStatus(fIndex).GelFilePathFull = GetFilePathFull(strFileName)
GelBody(fIndex).Caption = CompactPathString(GelStatus(fIndex).GelFilePathFull, 80)


GelBody(fIndex).Show
If GelData(fIndex).Preferences.CooType = glNETCooSys Then
    GelBody(fIndex).SetXAxisLabelType True
Else
    GelBody(fIndex).SetXAxisLabelType False
End If
    
UpdateFileMenu strFileName

exit_ReadGelFile:
Screen.MousePointer = vbDefault
MDIStatus False, ""

If blnSuccess Then
    ReadGelFile = fIndex
Else
    ReadGelFile = 0
End If

Exit Function
  
FailedReadGelFile:
SetGelStateToDeleted fIndex
blnSuccess = False

GoTo exit_ReadGelFile
End Function

Private Function GetGelCertificate(ByVal FileName As String) As String
'-----------------------------------------------------------------------------------
'retrieve string on the top of file; it should contain file certificate(version tag)
'-----------------------------------------------------------------------------------
Dim nFileNum As Integer
Dim sTmp As String
On Error Resume Next
nFileNum = FreeFile

' MonroeMod: Added check to look for file; without check, a zero byte file gets created when the file doesn't exist
If FileExists(FileName) Then
    Open FileName For Random Access Read As nFileNum
    Get #nFileNum, 1, sTmp
    Close nFileNum
Else
    sTmp = glCERT_FileNotFound
End If
' MonroeMod Finish

GetGelCertificate = sTmp
End Function

Private Function ReadGelData2003(ByVal FileName As String, ByVal Ind As Long, Optional blnInformUserOnError As Boolean = True) As Boolean
Dim tmp As DocumentData2003
Dim MaxInd As Long
Dim I As Long, j As Long
Dim hfile As Long

On Error GoTo exit_ReadGelData2003

hfile = FreeFile
Open FileName For Binary Access Read As hfile
If Err Then
   MsgBox "Can't open file: " & FileName, vbOKOnly, glFGTU
   LogErrors Err.Number, "OpenFile"
   Exit Function
End If
Get #hfile, 1, tmp

'transfer old structure to the new structure
With GelData(Ind)
    .Certificate = glCERT2003
    .Comment = tmp.Comment
    .FileName = tmp.FileName
    .Fileinfo = tmp.Fileinfo
    .PathtoDataFiles = tmp.PathtoDataFiles
    .PathtoDatabase = tmp.PathtoDatabase
    .LinesRead = tmp.LinesRead
    .DataLines = tmp.DataLines
    .CSLines = tmp.CSLines
    .IsoLines = tmp.IsoLines
    .MinMW = tmp.MinMW
    .MaxMW = tmp.MaxMW
    .MinAbu = tmp.MinAbu
    .MaxAbu = tmp.MaxAbu
    .Preferences = tmp.Preferences
    .pICooSysEnabled = tmp.pICooSysEnabled
    MaxInd = UBound(tmp.DFFN)   'first index is always 1
    If MaxInd > 0 Then
       ReDim .ScanInfo(MaxInd)
       For I = 1 To MaxInd
           With .ScanInfo(I)
              .ScanNumber = tmp.DFFN(I)
              .ScanFileName = tmp.DFN(I)
              .ScanPI = tmp.DFPI(I)
           End With
       Next I
    End If
    For I = 1 To MAX_FILTER_COUNT_2003
        For j = 0 To 2
            .DataFilter(I, j) = tmp.DataFilter(I, j)
        Next j
    Next I
    
    If .CSLines > 0 Then
       ReDim .CSData(.CSLines)
       For I = 1 To .CSLines
            CopyLegacyCSToIsoData .CSData(I), tmp.CSNum, tmp.CSVar, I
       Next I
    End If
    If .IsoLines > 0 Then
       ReDim .IsoData(.IsoLines)
       For I = 1 To .IsoLines
            CopyLegacyIsoToIsoData .IsoData(I), tmp.IsoNum, tmp.IsoVar, I
       Next I
    End If
End With

AddToAnalysisHistory Ind, "Opened data file with old format; will be updated to new format when saved."
ReadGelData2003 = True
  
exit_ReadGelData2003:
Close hfile
End Function

Public Function FindFreeIndex() As Long
'---------------------------------------------------------
'return first free index in document array
'(if any is deleted(closed) take that one to fill the gap)
'---------------------------------------------------------
Dim ArrayCnt As Long, I As Long
On Error Resume Next
ArrayCnt = UBound(GelBody)
If ArrayCnt > 0 Then
   For I = 1 To ArrayCnt                   'can not use 0
       If GelStatus(I).Deleted Then
          InitializeGelDataStructures I
          FindFreeIndex = I
          Exit Function
       End If
   Next I
End If
'none deleted; increase upper bound for arrays
ArrayCnt = ArrayCnt + 1
ReDim Preserve GelBody(ArrayCnt)
ReDim Preserve GelStatus(ArrayCnt)
ReDim Preserve GelData(ArrayCnt)
ReDim Preserve GelDraw(ArrayCnt)
ReDim Preserve GelUMC(ArrayCnt)
ReDim Preserve GelUMCIon(ArrayCnt)
' No longer supported (March 2006)
''ReDim Preserve GelDB(ArrayCnt)
''ReDim Preserve GelP(ArrayCnt)
ReDim Preserve GelP_D_L(ArrayCnt)
ReDim Preserve GelAnalysis(ArrayCnt)
ReDim Preserve GelLM(ArrayCnt)
ReDim Preserve GelUMCDraw(ArrayCnt)

' Unused variable (August 2003)
'ReDim Preserve GelIDP(ArrayCnt)

' MonroeMod Start
ReDim Preserve GelUMCNETAdjDef(ArrayCnt)
ReDim Preserve GelSearchDef(ArrayCnt)
ReDim Preserve GelDataLookupArrays(ArrayCnt)

' No longer supported (March 2006)
''ReDim Preserve GelORFData(ArrayCnt)
''ReDim Preserve GelORFMassTags(ArrayCnt)
''ReDim Preserve GelORFViewerSavedGelListAndOptions(ArrayCnt)

InitializeGelDataStructures ArrayCnt
' MonroeMod Finish

FindFreeIndex = UBound(GelBody)
End Function

Public Sub FixIsosMonoPlus2Abu(ByVal lngGelIndex As Long, Optional ByVal intMatchTolerancePPMStart As Integer = 15, Optional intMatchToleranceIterations As Integer = 4, Optional ByVal dblIsoPlus2SpacingDa As Double = 2.0038)
    
    ' Processes isotopic data to determine the correct IntensityMonoPlus2 value when an IntensityMono
    ' value is defined but IntensityMonoPlus2 is 0
    '
    '
    ' This function assumes the data in GelData.IsoData() is sorted by scan number and then
    '  by mass.  Therefore, you MUST call SortIsotopicData prior to calling this function to assure
    '  that the data is sorted properly (using .Preferences.IsoDataField)

    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    
    Dim lngIndexMax As Long
    Dim lngCurrentScan As Long
    
    Dim dblIsoPlus2MassCenter As Double
    Dim dblIsoPlus2MassMin As Double
    Dim dblIsoPlus2MassMax As Double
    
    Dim lngIsoPlus2AbuDataInRange As Long
    Dim dblIsoPlus2AbuMax As Double
    Dim dblCurrentAbu As Double
    
    Dim intMatchTolerancePPM As Integer
    Dim intMatchToleranceIterationsElapsed As Integer
    
    If intMatchTolerancePPMStart < 1 Then intMatchTolerancePPMStart = 1
    If intMatchToleranceIterations < 1 Then intMatchToleranceIterations = 1
    
    lngIndexMax = GelData(lngGelIndex).IsoLines
    If lngIndexMax > 1 Then
        
        ' Step through the data and look for points with a non-zero .IntensityMono value but having IntensityMonoPlus2 = 0
        For lngIndex = 1 To lngIndexMax
            If GelData(lngGelIndex).IsoData(lngIndex).IntensityMono > 0 And GelData(lngGelIndex).IsoData(lngIndex).IntensityMonoPlus2 = 0 Then
                
                ' Match found; assure its monoisotopic mass is non-zero
                If GelData(lngGelIndex).IsoData(lngIndex).MonoisotopicMW > 0 Then
                    lngCurrentScan = GelData(lngGelIndex).IsoData(lngIndex).ScanNumber
                    
                    intMatchTolerancePPM = intMatchTolerancePPMStart
                    
                    intMatchToleranceIterationsElapsed = 0
                    Do While intMatchToleranceIterationsElapsed < intMatchToleranceIterations
                        
                        dblIsoPlus2MassCenter = GelData(lngGelIndex).IsoData(lngIndex).MonoisotopicMW + dblIsoPlus2SpacingDa
                        dblIsoPlus2MassMin = dblIsoPlus2MassCenter - intMatchTolerancePPM * (dblIsoPlus2MassCenter / 1000000#)
                        dblIsoPlus2MassMax = dblIsoPlus2MassCenter + intMatchTolerancePPM * (dblIsoPlus2MassCenter / 1000000#)
                        
                        lngIsoPlus2AbuDataInRange = 0
                        dblIsoPlus2AbuMax = 0
                        
                        ' Step through the subsequent data points, looking for those between dblIsoPlus2MassMin and dblIsoPlus2MassMax
                        ' Abort the search if a point is found weighing more than dblIsoPlus2MassMax or with a different scan number
                        For lngIndexCompare = lngIndex + 1 To lngIndexMax
                            If GelData(lngGelIndex).IsoData(lngIndexCompare).ScanNumber <> lngCurrentScan Then
                                Exit For
                            ElseIf GelData(lngGelIndex).IsoData(lngIndexCompare).MonoisotopicMW > dblIsoPlus2MassMax Then
                                Exit For
                            ElseIf GelData(lngGelIndex).IsoData(lngIndexCompare).MonoisotopicMW > dblIsoPlus2MassMin Then
                                ' Matching mass
                                dblCurrentAbu = GelData(lngGelIndex).IsoData(lngIndexCompare).IntensityMono
                                If dblCurrentAbu = 0 Then
                                    dblCurrentAbu = GelData(lngGelIndex).IsoData(lngIndexCompare).Abundance
                                End If
                                    
                                If dblCurrentAbu > dblIsoPlus2AbuMax Or lngIsoPlus2AbuDataInRange = 0 Then
                                    dblIsoPlus2AbuMax = dblCurrentAbu
                                End If
                                lngIsoPlus2AbuDataInRange = lngIsoPlus2AbuDataInRange + 1
                            End If
                        Next lngIndexCompare
                        
                        If lngIsoPlus2AbuDataInRange > 0 Then
                            ' A match was found; update the .IntensityMonoPlus2 value
                            GelData(lngGelIndex).IsoData(lngIndex).IntensityMonoPlus2 = dblIsoPlus2AbuMax
                            Exit Do
                        Else
                            intMatchTolerancePPM = intMatchTolerancePPM * 2
                        End If
                        
                        intMatchToleranceIterationsElapsed = intMatchToleranceIterationsElapsed + 1
                    Loop
                    
                End If
            End If
        Next lngIndex
    End If

End Sub

Private Sub InitializeGelDataStructures(ByVal lngGelIndex As Long)
        
    GelData(lngGelIndex).DataStatusBits = 0
    GelData(lngGelIndex).MostRecentSearchUsedSTAC = False
        
    With GelStatus(lngGelIndex)
        .Deleted = False
        .SourceDataRawFileType = rfcUnknown
    End With
        
    GelUMCNETAdjDef(lngGelIndex) = UMCNetAdjDef
    
    With GelSearchDef(lngGelIndex)
        .UMCDef = UMCDef
        .UMCIonNetDef = UMCIonNetDef
        .AMTSearchOnIons = samtDef
        .AMTSearchOnUMCs = samtDef
        .AMTSearchOnPairs = samtDef
        .AnalysisHistoryCount = 0
        
        ReDim .AnalysisHistory(0)
        With .MassCalibrationInfo
            .OverallMassAdjustment = 0
            .OtherInfo = ""
            .AdjustmentHistoryCount = 0
            ReDim .AdjustmentHistory(0)
        End With
        
        ResetDBSearchMassMods .AMTSearchMassMods
    End With
    
    ''GelORFViewerSavedGelListAndOptions(lngGelIndex).IsDefined = False
    
    With GelBody(lngGelIndex)
        .mFileSaveMode = fsUnknown
        .Tag = lngGelIndex
    
        If lngGelIndex > 1 Then
            .mnuSCopyScansIncludeEmptyScans.Checked = GelBody(1).mnuSCopyScansIncludeEmptyScans.Checked
        End If
    End With

    GelP_D_L(lngGelIndex).SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
    
    SetEditCopyEMFOptions glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeFilenameAndDate, glbPreferencesExpanded.GraphicExportOptions.CopyEMFIncludeTextLabels

End Sub

' MonroeMod: Function revised to use the .Ini file rather than the registry
Public Sub GetRecentFiles()
'Procedure returns an array of values from the application's Ini File
'Stores the files in glbRecentFiles
'Displays shortened file names on the menus, but keeps track of the full file name in glbRecentFiles
Dim I As Integer
Dim j As Integer
Dim IniStuff As New clsIniStuff
Dim strFilePath As String
Dim lngFileCount As Long
Dim blnProceed As Boolean

On Error GoTo GetRecentFilesErrorHandler

    ' Set the Ini filename
    IniStuff.FileName = GetIniFilePath()
    
    lngFileCount = GetIniFileSettingLng(IniStuff, INI_SECTION_RECENT_FILES, INI_KEY_RECENT_FILE_COUNT, 0)
    If lngFileCount > MAX_RECENT_FILE_COUNT Then
        lngFileCount = MAX_RECENT_FILE_COUNT
    End If
      
    If lngFileCount = 0 Then
        MDIForm1.mnuRecentFiles(0).Visible = False
        For j = 1 To UBound(GelBody)
          If Not GelStatus(j).Deleted Then
            GelBody(j).mnuRecentFiles(0).Visible = False
          End If
        Next j
        glbRecentFiles.FileCount = 0
    Else
        MDIForm1.mnuRecentFiles(0).Visible = True
    End If
    
    If lngFileCount > 0 Then
        'update menus on MDI form and each visible child form
        glbRecentFiles.FileCount = 0
        For I = 0 To lngFileCount - 1
            ' Need to add 1 to i since first recent file is RecentFile1
            strFilePath = GetIniFileSetting(IniStuff, INI_SECTION_RECENT_FILES, INI_KEY_RECENT_FILE_PREFIX & Trim(CStr(I + 1)), "")
            
            If Len(strFilePath) > 0 Then
                ' Add the file to glbRecentFiles.Files(), provided it's not already in the list
                With glbRecentFiles
                    blnProceed = True
                    For j = 0 To .FileCount - 1
                        If UCase(.Files(j).FullFilePath) = UCase(strFilePath) Then
                            blnProceed = False
                            Exit For
                        End If
                    Next j
                End With
                
                If blnProceed Then
                    glbRecentFiles.FileCount = glbRecentFiles.FileCount + 1
                    With glbRecentFiles.Files(glbRecentFiles.FileCount - 1)
                        .FullFilePath = strFilePath
                        .ShortenedFilePath = CompactPathString(.FullFilePath, 65)
                        
                        MDIForm1.mnuRecentFiles(glbRecentFiles.FileCount).Caption = .ShortenedFilePath
                        MDIForm1.mnuRecentFiles(glbRecentFiles.FileCount).Visible = (Len(.ShortenedFilePath) > 0)
                        
                        For j = 1 To UBound(GelBody)
                            If Not GelStatus(j).Deleted Then
                                GelBody(j).mnuRecentFiles(0).Visible = True
                                GelBody(j).mnuRecentFiles(glbRecentFiles.FileCount).Caption = .ShortenedFilePath
                                GelBody(j).mnuRecentFiles(glbRecentFiles.FileCount).Visible = True
                            End If
                        Next j
                    End With
                End If
            End If
          
        Next I
    End If
    
    ' Hide the remaining menus
    For I = glbRecentFiles.FileCount + 1 To MAX_RECENT_FILE_COUNT
        MDIForm1.mnuRecentFiles(I).Caption = ""
        MDIForm1.mnuRecentFiles(I).Visible = False
          
        For j = 1 To UBound(GelBody)
            If Not GelStatus(j).Deleted Then
                GelBody(j).mnuRecentFiles(0).Visible = False
                GelBody(j).mnuRecentFiles(I).Caption = ""
                GelBody(j).mnuRecentFiles(I).Visible = False
            End If
        Next j
    Next I
    
    Set IniStuff = Nothing

    Exit Sub

GetRecentFilesErrorHandler:
    Debug.Print "Error in sub GetRecentFiles(): " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "GetRecentFiles"
    
End Sub

' MonroeMod: Function revised to use the .Ini file rather than the registry
Private Sub WriteRecentFiles(strFilePathToAddOrUpdate As String, Optional intIndexInRecentFilesUDT As Integer = -1)
'Procedure writes the names of recently opened files to the .Ini file
'strFilePathToAddOrUpdate becomes the most recent opened, existing entries are pushed up one spot
'If intIndexInRecentFilesUDT>=0 then no new entry is written, entry for intIndexInRecentFilesUDT becomes first entry
Dim j As Integer
Dim IniStuff As New clsIniStuff
Dim blnSuccess As Boolean

On Error GoTo err_writerecentfiles

' Set the Ini filename
IniStuff.FileName = GetIniFilePath()

With glbRecentFiles
    If intIndexInRecentFilesUDT = 0 Then
        ' No shuffling is required
    ElseIf intIndexInRecentFilesUDT > 0 Then
        ' Shuffle order of entries in glbRecentFiles.Files() accordingly
        For j = intIndexInRecentFilesUDT To 1 Step -1
            .Files(j) = .Files(j - 1)
        Next j
    Else
        ' Shuffle all entries up 1 spot
        ' Add a new entry at position 0
        If .FileCount < MAX_RECENT_FILE_COUNT Then .FileCount = .FileCount + 1
        For j = .FileCount - 1 To 1 Step -1
            .Files(j) = .Files(j - 1)
        Next j
    End If
    If .FileCount = 0 Then .FileCount = 1
    .Files(0).FullFilePath = strFilePathToAddOrUpdate
    .Files(0).ShortenedFilePath = CompactPathString(strFilePathToAddOrUpdate)

    ' Now write all of the recent files to the .Ini file
    For j = 0 To .FileCount - 1
        blnSuccess = IniStuff.WriteValue(INI_SECTION_RECENT_FILES, INI_KEY_RECENT_FILE_PREFIX & Trim(j + 1), .Files(j).FullFilePath)
        If Not blnSuccess Then
            Debug.Assert False
            Exit For
        End If
    Next j
    If blnSuccess Then
        IniStuff.WriteValue INI_SECTION_RECENT_FILES, INI_KEY_RECENT_FILE_COUNT, CStr(.FileCount)
    End If
End With

Set IniStuff = Nothing
Exit Sub

err_writerecentfiles:
Debug.Print "Error in sub WriteRecentFiles(): " & Err.Description
Debug.Assert False
LogErrors Err.Number, "WriteRecentFiles"
End Sub

Private Sub RemoveFromRecentFiles(ByVal strFileNameToRemove As String)
    ' Removes FileName from the recent file list (probably since the file no longer exists)
    Dim intIndex As Integer, intIndexShift As Integer
   
    If Len(strFileNameToRemove) = 0 Then Exit Sub
   
    strFileNameToRemove = LCase(strFileNameToRemove)
    
    With glbRecentFiles
        
        intIndex = 0
        Do While intIndex < .FileCount
            If LCase(.Files(intIndex).FullFilePath) = strFileNameToRemove Then
                For intIndexShift = intIndex To .FileCount - 2
                    .Files(intIndexShift) = .Files(intIndexShift + 1)
                Next intIndexShift
                If .FileCount > 0 Then .FileCount = .FileCount - 1
            Else
                intIndex = intIndex + 1
            End If
        Loop
        
    End With

    ' Call WriteRecentFiles, sending the path of the 0th file, so that the Ini file gets updated
    WriteRecentFiles glbRecentFiles.Files(0).FullFilePath, 0
    
End Sub

Public Function RecentFileLookUpFullPath(strShortenedFilePath As String) As String
    ' Looks for strFilePath in glbRecentFiles
    ' Returns the full path if found
    ' Returns "" if not found
    
    Dim intIndex As Integer
    Dim strFullPath As String
    
    For intIndex = 0 To glbRecentFiles.FileCount - 1
        If LCase(strShortenedFilePath) = LCase(glbRecentFiles.Files(intIndex).ShortenedFilePath) Then
            strFullPath = glbRecentFiles.Files(intIndex).FullFilePath
            Exit For
        End If
    Next intIndex
    
    If Len(strFullPath) = 0 Then
        ' This is unexpected
        Debug.Assert False
        
        ' Path not found
        ' Search through list again to see if any of the FullFilePath entries match strShortenedFilePath
        For intIndex = 0 To glbRecentFiles.FileCount - 1
            If LCase(strShortenedFilePath) = LCase(glbRecentFiles.Files(intIndex).FullFilePath) Then
                strFullPath = glbRecentFiles.Files(intIndex).FullFilePath
                Exit For
            End If
        Next intIndex
    End If
    
    RecentFileLookUpFullPath = strFullPath
    
End Function

Public Function SaveFileAsPicture(ByVal lngGelIndex As Long, ByVal strFilePath As String, ByVal PicSaveType As pftPictureFileTypeConstants) As Long
    ' Returns 0 if success, the error code if an error
    
    Dim m_cDIB As New cDIBSection   'structure used to save JPG
    Dim lngErrorCode As Long
    Dim blnSuccess As Boolean
    
    Dim strEmfFilePath As String
    Dim strPNGFilePath As String
    Dim strWorkingFilePath As String
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
    
On Error GoTo SaveFileAsPictureErrorHandler

    If Len(strFilePath) > 0 Then
       ' Make sure strFilePath has the correct extension
       Select Case PicSaveType
       Case pftPictureFileTypeConstants.pftJPG
           strFilePath = FileExtensionForce(strFilePath, "jpg")
       Case pftPictureFileTypeConstants.pftWMF
           strFilePath = FileExtensionForce(strFilePath, "wmf")
       Case pftPictureFileTypeConstants.pftEMF
           strFilePath = FileExtensionForce(strFilePath, "emf")
       Case pftPictureFileTypeConstants.pftPNG
           strFilePath = FileExtensionForce(strFilePath, "png")
       Case pftPictureFileTypeConstants.pftBMP
           strFilePath = FileExtensionForce(strFilePath, "bmp")
       End Select
       
       strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
       
       Select Case PicSaveType
       Case pftPictureFileTypeConstants.pftJPG
           GelBody(lngGelIndex).picGraph.AutoRedraw = True
           GelDrawScreen lngGelIndex
           m_cDIB.LoadFromBMP GelBody(lngGelIndex).picGraph.Image
           If Not SaveJPGToFile(m_cDIB, strWorkingFilePath) Then
              lngErrorCode = -1
           End If
           GelBody(lngGelIndex).picGraph.Cls
           GelBody(lngGelIndex).picGraph.AutoRedraw = False
           blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
           
       Case pftPictureFileTypeConstants.pftWMF
           lngErrorCode = GelDrawMetafile(lngGelIndex, True, strWorkingFilePath, False)
           blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
           
       Case pftPictureFileTypeConstants.pftEMF, pftPictureFileTypeConstants.pftPNG
           If PicSaveType = pftPictureFileTypeConstants.pftPNG Then
               strEmfFilePath = Left(strWorkingFilePath, Len(strWorkingFilePath) - 4) & "_Temp" & Trim(Abs(GetTickCount())) & ".emf"
           Else
               strEmfFilePath = strWorkingFilePath
           End If
           
           lngErrorCode = GelDrawMetafile(lngGelIndex, True, strEmfFilePath, True)
           
           If PicSaveType = pftPictureFileTypeConstants.pftPNG And lngErrorCode = 0 Then
               ConvertEmfToPng strEmfFilePath, strWorkingFilePath, GelBody(lngGelIndex).width / Screen.TwipsPerPixelX, GelBody(lngGelIndex).Height / Screen.TwipsPerPixelY
           End If
           blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
           
       Case pftPictureFileTypeConstants.pftBMP
           GelBody(lngGelIndex).picGraph.AutoRedraw = True
           GelDrawScreen lngGelIndex
           SavePicture GelBody(lngGelIndex).picGraph.Image, strWorkingFilePath
           GelBody(lngGelIndex).picGraph.Cls
           GelBody(lngGelIndex).picGraph.AutoRedraw = False
           lngErrorCode = 0
           blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
           
       Case Else
           If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then MsgBox "Save picture - unknown format.", vbOKOnly
           lngErrorCode = -1
       End Select
    End If
    
    Set m_cDIB = Nothing
    
    SaveFileAsPicture = lngErrorCode
    
    Exit Function
    
SaveFileAsPictureErrorHandler:
    SaveFileAsPicture = AssureNonZero(Err.Number)
    
End Function

Public Sub SaveWYSAs(ByVal Ind As Long, ByVal eFileSaveMode As fsFileSaveModeConstants)
'------------------------------------------------------------------------------------
'this function saves only visible portion of the gel(data in scope) in a new display;
'newly created file is not opened and currently active display stays active
'------------------------------------------------------------------------------------
Dim TmpGD As DocumentData
Dim StructureSize As Long
Dim TmpCnt As Long
Dim ScopeInd() As Long
Dim lngIsoIndexNew() As Long            ' This array holds the new index values for all of the points; necessary for updating LC-MS Features
Dim I As Long, j As Long, k As Long

Dim lngUMCCountNew As Long
Dim udtUMCListSaved As UMCListType

' MonroeMod: New Variables
Dim strFilePathSaved As String, FileName As String

On Error GoTo SaveWYSAsErrorHandler

' MonroeMod: Save the file path so that the Caption can be restored after the save
strFilePathSaved = GelBody(Ind).Caption

FileName = FileSaveProc(GelBody(Ind).hwnd, "", fstFileSaveTypeConstants.fstGel)
If Len(FileName) <= 0 Then Exit Sub

' MonroeMod: Portions of this sub commented out since the
'            actual saving is accomplished using BinarySaveDAta
''FileNum = FreeFile
''Open FileName For Binary Access Write As FileNum
''If Err Then
''   MsgBox "Error creating file: " & FileName, vbOKOnly, glFGTU
''   LogErrors Err.Number, "SaveWYSAs"
''   Exit Sub
''End If
''Screen.MousePointer = 11

With GelData(Ind)
    'copy scalar values
    TmpGD.Certificate = .Certificate
    TmpGD.Comment = .Comment & vbCrLf & glCOMMENT_WYS & GelBody(Ind).Caption & vbCrLf
    TmpGD.FileName = .FileName
    TmpGD.Fileinfo = .Fileinfo
    TmpGD.PathtoDataFiles = .PathtoDataFiles
    
    TmpGD.PathtoDatabase = .PathtoDatabase
    
    TmpGD.MediaType = .MediaType
    TmpGD.LinesRead = .LinesRead
    ' TmpGD.DataLines, .CSLines, and .IsoLines are filled in below
    
    TmpGD.CalEquation = .CalEquation
    For I = LBound(.CalArg) To UBound(.CalArg)
        TmpGD.CalArg(I) = .CalArg(I)
    Next I
    
    TmpGD.Preferences = .Preferences
    TmpGD.pICooSysEnabled = .pICooSysEnabled
    TmpGD.DataStatusBits = .DataStatusBits
    
    ' Copy the data filters
    For I = 1 To MAX_FILTER_COUNT
        For j = 0 To 2
            TmpGD.DataFilter(I, j) = .DataFilter(I, j)
        Next j
    Next I
    
    TmpGD.CustomNETsDefined = .CustomNETsDefined
    
    ' Copy the scan info
    ReDim TmpGD.ScanInfo(1 To UBound(.ScanInfo))
    For I = 1 To UBound(.ScanInfo)
        TmpGD.ScanInfo(I) = .ScanInfo(I)
    Next I
    
    ' Copy the actual data; first CS data
    TmpGD.MinAbu = glHugeOverExp
    TmpGD.MaxAbu = 0
    TmpGD.MinMW = glHugeOverExp
    TmpGD.MaxMW = 0
    TmpCnt = GetScopeCS(Ind, ScopeInd())
    If TmpCnt > 0 Then
       TmpGD.CSLines = TmpCnt
       ReDim TmpGD.CSData(TmpCnt)
       For I = 1 To TmpCnt
           k = ScopeInd(I)
           TmpGD.CSData(I) = .CSData(k)
            
           If .CSData(k).Abundance < TmpGD.MinAbu Then TmpGD.MinAbu = .CSData(k).Abundance
           If .CSData(k).Abundance > TmpGD.MaxAbu Then TmpGD.MaxAbu = .CSData(k).Abundance
           
           FindMWExtremes TmpGD.CSData(I), TmpGD.MinMW, TmpGD.MaxMW, 0
       Next I
    Else
       TmpGD.CSLines = 0
    End If
    
    ' Now Isotopic data
    TmpCnt = GetScopeIso(Ind, ScopeInd())
    If TmpCnt > 0 Then
       ReDim lngIsoIndexNew(.IsoLines)
      
       TmpGD.IsoLines = TmpCnt
       ReDim TmpGD.IsoData(TmpCnt)
       For I = 1 To TmpCnt
            k = ScopeInd(I)
            lngIsoIndexNew(k) = I
            
            TmpGD.IsoData(I) = .IsoData(k)
       
            If .IsoData(k).Abundance < TmpGD.MinAbu Then TmpGD.MinAbu = .IsoData(k).Abundance
            If .IsoData(k).Abundance > TmpGD.MaxAbu Then TmpGD.MaxAbu = .IsoData(k).Abundance
        
            FindMWExtremes TmpGD.IsoData(I), TmpGD.MinMW, TmpGD.MaxMW, 0
       Next I
    Else
       TmpGD.IsoLines = 0
    End If
    TmpGD.DataLines = TmpGD.CSLines + TmpGD.IsoLines
End With

' Make a backup copy of GelUMC(Ind)
udtUMCListSaved = GelUMC(Ind)

' MonroeMod: Remove invalid LC-MS Features
With GelUMC(Ind)
    If TmpGD.DataLines = 0 Then
        .UMCCnt = 0
        ReDim .UMCs(0)
    Else
        For I = 0 To .UMCCnt - 1
            .UMCs(I).ClassRepInd = lngIsoIndexNew(.UMCs(I).ClassRepInd)
            For j = 0 To .UMCs(I).ClassCount - 1
                If .UMCs(I).ClassRepInd = 0 Or lngIsoIndexNew(.UMCs(I).ClassMInd(j)) = 0 Then
                    ' Class contains one or more invalid members; set ClassCount to 0
                    .UMCs(I).ClassCount = 0
                    .UMCs(I).ClassAbundance = 0
                    .UMCs(I).ClassMW = 0
                    .UMCs(I).ClassStatusBits = 0
                    .UMCs(I).ClassRepInd = 1
                    ReDim .UMCs(I).ClassMInd(0)
                    Exit For
                Else
                    ' Update the index for this member
                    .UMCs(I).ClassMInd(j) = lngIsoIndexNew(.UMCs(I).ClassMInd(j))
                End If
            Next j
        Next I
        
        ' Compress .UMCs, copying in place
        lngUMCCountNew = 0
        For I = 0 To .UMCCnt - 1
            If .UMCs(I).ClassCount > 0 Then
                .UMCs(lngUMCCountNew) = .UMCs(I)
                lngUMCCountNew = lngUMCCountNew + 1
            End If
        Next I
        If lngUMCCountNew < .UMCCnt Then
            .UMCCnt = lngUMCCountNew
            If .UMCCnt = 0 Then
                ReDim .UMCs(0)
            Else
                ReDim Preserve .UMCs(lngUMCCountNew - 1)
            End If
        End If
    End If
End With

' MonroeMod: Now call BinarySaveData
AddToAnalysisHistory Ind, "Only saving data in current view (current scope) to disk [i.e. Saving WYS (What-you-see)]"
AddToAnalysisHistory Ind, "Data Trimmed: Charge State Data Points = " & Trim(TmpGD.CSLines)
AddToAnalysisHistory Ind, "Data Trimmed: Isotopic (deconvoluted) Data Points = " & Trim(TmpGD.IsoLines)
BinarySaveData FileName, False, True, Ind, TmpGD, eFileSaveMode

' Restore the LC-MS Features
GelUMC(Ind) = udtUMCListSaved

' Need to restore the caption on the calling window to the original file path (BinarySaveData changed it to the new file path)
GelBody(Ind).Caption = strFilePathSaved
GelStatus(Ind).GelFilePathFull = GetFilePathFull(strFilePathSaved)

Exit Sub

SaveWYSAsErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "SaveWYSAs"
End Sub

Public Function GetScopeCS(ByVal Ind As Long, ByRef CS() As Long) As Long
'------------------------------------------------------------------------
'fills array CS with indexes in GelData(Ind).CSData that are in
'current scope of gel(index, zoom) and returns number of it(-1 if error)
'------------------------------------------------------------------------
Dim TmpCnt As Long
Dim I As Long
On Error GoTo err_GetScopeCS

With GelDraw(Ind)
  If (.CSCount > 0 And .CSVisible) Then
     ReDim CS(1 To .CSCount)
     TmpCnt = 0
     Select Case GelBody(Ind).fgDisplay
     Case glNormalDisplay
        For I = 1 To .CSCount
            If .CSID(I) > 0 And .CSR(I) > 0 Then
               TmpCnt = TmpCnt + 1
               CS(TmpCnt) = I
            End If
        Next I
     Case glDifferentialDisplay
        For I = 1 To .CSCount
            If (.CSER(I) >= 0 And .CSID(I) > 0 And .CSR(I) > 0) Then
               TmpCnt = TmpCnt + 1
               CS(TmpCnt) = I
            End If
        Next I
     End Select
     If TmpCnt > 0 Then
        ReDim Preserve CS(1 To TmpCnt)
     Else
        Erase CS
     End If
     GetScopeCS = TmpCnt
  Else
     GetScopeCS = 0
  End If
End With
Exit Function

err_GetScopeCS:
LogErrors Err.Number, "GetScopeCS"
GetScopeCS = -1
End Function

Public Function GetScopeIso(ByVal Ind As Long, ByRef Iso() As Long) As Long
'--------------------------------------------------------------------------
'fills array Iso with indexes in GelData(Ind).IsoData that are in
'current scope of gel(index, zoom) and returns number of it(-1 if error)
'--------------------------------------------------------------------------
Dim TmpCnt As Long
Dim I As Long
On Error GoTo err_GetScopeIso

With GelDraw(Ind)
  If (.IsoCount > 0 And .IsoVisible) Then
     ReDim Iso(1 To .IsoCount)
     TmpCnt = 0
     Select Case GelBody(Ind).fgDisplay
     Case glNormalDisplay
        For I = 1 To .IsoCount
            If .IsoID(I) > 0 And .IsoR(I) > 0 Then
               TmpCnt = TmpCnt + 1
               Iso(TmpCnt) = I
            End If
        Next I
     Case glDifferentialDisplay
        For I = 1 To .IsoCount
            If (.IsoER(I) >= 0 And .IsoID(I) > 0 And .IsoR(I) > 0) Then
               TmpCnt = TmpCnt + 1
               Iso(TmpCnt) = I
            End If
        Next I
     End Select
     If TmpCnt > 0 Then
        ReDim Preserve Iso(1 To TmpCnt)
     Else
        Erase Iso
     End If
     GetScopeIso = TmpCnt
  Else
     GetScopeIso = 0
  End If
End With
Exit Function

err_GetScopeIso:
LogErrors Err.Number, "GetScopeIso"
GetScopeIso = -1
End Function

Public Function InitDrawUMC(ByVal Ind As Long) As Boolean
'--------------------------------------------------------------------------------
'initialize structure for drawing Unique Mass Classes; returns True if successful
'this function has to be called every time Unique Mass Classes are calculated
'--------------------------------------------------------------------------------
Dim I As Long
On Error GoTo exit_InitDrawUMC
With GelUMCDraw(Ind)
    .Count = GelUMC(Ind).UMCCnt
    If .Count > 0 Then
        ReDim .ClassID(.Count - 1)
        ReDim .X1(.Count - 1):        ReDim .Y1(.Count - 1)
        ReDim .x2(.Count - 1):        ReDim .Y2(.Count - 1)
        For I = 0 To .Count - 1
            .ClassID(I) = I + 1       'so that we can use negative indexes for
                                      'classes that don't have to be drawn
        Next I
    Else
        Erase .ClassID
        Erase .X1:           Erase .Y1
        Erase .x2:           Erase .Y2
    End If
End With
InitDrawUMC = True
Exit Function

exit_InitDrawUMC:
Debug.Assert False
InitDrawUMC = False

End Function



Public Function GetWindowCS(ByVal Ind As Long, CSRes() As Long, _
                            Scan1 As Long, Scan2 As Long, _
                            MW1 As Double, MW2 As Double) As Long
'-------------------------------------------------------------------------
'fills array ResCS with indexes in GelData(Ind).CSData that are in a window
'(Scan1,Scan2) x (MW1,MW2) and returns number of it; -1 on any error
'-------------------------------------------------------------------------
Dim TmpCnt As Long
Dim I As Long
On Error GoTo err_GetWindowCS

With GelData(Ind)
     ReDim CSRes(.CSLines - 1)
     For I = 1 To .CSLines
         If ((.CSData(I).ScanNumber >= Scan1) And (.CSData(I).ScanNumber <= Scan2)) Then
            If ((.CSData(I).AverageMW >= MW1) And (.CSData(I).AverageMW <= MW2)) Then
               TmpCnt = TmpCnt + 1
               CSRes(TmpCnt - 1) = I
            End If
         End If
     Next I
End With
If TmpCnt > 0 Then
   ReDim Preserve CSRes(TmpCnt - 1)
Else
   Erase CSRes
End If
GetWindowCS = TmpCnt
Exit Function

err_GetWindowCS:
If Err.Number <> 9 Then LogErrors Err.Number, "GetWindowCS"
GetWindowCS = -1
End Function


Public Function GetWindowIso(ByVal Ind As Long, FMW As Integer, IsoRes() As Long, _
                             Scan1 As Long, Scan2 As Long, _
                             MW1 As Double, MW2 As Double) As Long
'-------------------------------------------------------------------------------
'fills array ResIso with indexes in GelData(Ind).IsoData that are in a window
'(Scan1,Scan2) x (MW1,MW2) and returns number of it; -1 on any error
'fMW is column from which we need to extract molecular masses
'-------------------------------------------------------------------------------
Dim TmpCnt As Long
Dim I As Long
On Error GoTo err_GetWindowIso

With GelData(Ind)
     ReDim IsoRes(.IsoLines - 1)
     For I = 1 To .IsoLines
         If ((.IsoData(I).ScanNumber >= Scan1) And (.IsoData(I).ScanNumber <= Scan2)) Then
            If ((GetIsoMass(.IsoData(I), FMW) >= MW1) And (GetIsoMass(.IsoData(I), FMW) <= MW2)) Then
               TmpCnt = TmpCnt + 1
               IsoRes(TmpCnt - 1) = I
            End If
         End If
     Next I
End With
If TmpCnt > 0 Then
   ReDim Preserve IsoRes(TmpCnt - 1)
Else
   Erase IsoRes
End If
GetWindowIso = TmpCnt
Exit Function

err_GetWindowIso:
If Err.Number <> 9 Then LogErrors Err.Number, "GetWindowIso"
GetWindowIso = -1
End Function


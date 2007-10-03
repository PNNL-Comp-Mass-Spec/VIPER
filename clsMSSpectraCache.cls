VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSSpectraCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class can be used to extract spectra from .Zip files and cache them in a local folder
' It will delete spectra from the cache if they haven't been accessed for a period of time (7 days)

Private Const LAST_ACCESSED_FILENAME As String = "Last_Accessed.txt"

Private Const MIN_CACHE_FOLDERNAME_LENGTH As Integer = 5
Private Const MAX_CACHE_FOLDERNAME_LENGTH As Integer = 30

Public Enum scSpectraCacheErrors
    scNoError = 0
    scInvalidCachePath = 1
    scOriginalDataFolderNotFound = 2
    scZipFilesNotFoundInDataFolder = 4
    scErrorCachingSpectrum = 8
    scErrorCreatingCacheFolder = 16
    scErrorDeletingCachedSpectra = 32
    scUnidentifiedError = 1024
End Enum

Private mCachePath As String
Private mAutoDeleteOldCachedSpectra As Boolean
Private mDaysToCacheSpectra As Long
Private mMaxFolderSizeMB As Long
Private mLastAutoPurgeDate As Date

Private mLastCachedFolderPath As String

Private mLocalErrorCode As scSpectraCacheErrors

Private fso As FileSystemObject

Public Property Get AutoDeleteOldCachedSpectra() As Boolean
    AutoDeleteOldCachedSpectra = mAutoDeleteOldCachedSpectra
End Property
Public Property Let AutoDeleteOldCachedSpectra(ByVal blnEnable As Boolean)
    mAutoDeleteOldCachedSpectra = blnEnable
End Property

Public Property Get CacheFolderPath() As String
    CacheFolderPath = mCachePath
End Property
Public Property Let CacheFolderPath(ByVal strFolderPath As String)
    If Len(strFolderPath) > 0 Then
        mCachePath = strFolderPath
    End If
End Property

Public Property Get DaysToCacheSpectra() As Long
    DaysToCacheSpectra = mDaysToCacheSpectra
End Property
Public Property Let DaysToCacheSpectra(ByVal lngDays As Long)
    If lngDays > 0 Then
        mDaysToCacheSpectra = lngDays
    Else
        mDaysToCacheSpectra = 1
    End If
End Property

Public Property Get ErrorCode() As scSpectraCacheErrors
    ErrorCode = mLocalErrorCode
End Property

Public Property Get MaxFolderSizeMB() As Long
    MaxFolderSizeMB = mMaxFolderSizeMB
End Property
Public Property Let MaxFolderSizeMB(lngMaxSizeMB As Long)
    If lngMaxSizeMB < 10 Then lngMaxSizeMB = 10
    mMaxFolderSizeMB = lngMaxSizeMB
End Property


Public Sub CheckFolderSpaceUsage(ByVal strFolderPath As String, Optional ByVal lngMaxFolderSizeMB As Long = 0)

    Dim objFolder As Folder
    Dim objFile As File
    
    Dim lngFileCount As Long
    Dim strFileNames() As String
    Dim lngPointerArray() As Long
    Dim FileAccessDates() As Double
    
    Dim lngIndex As Long
    
    Dim objQS As QSDouble
    
On Error GoTo CheckFolderSpaceUsageErrorHandler

    If fso.FolderExists(strFolderPath) Then
        If lngMaxFolderSizeMB < 10 Then lngMaxFolderSizeMB = mMaxFolderSizeMB
        
        Set objFolder = fso.GetFolder(strFolderPath)
        
        If objFolder.Size / 1024 / 1024 > lngMaxFolderSizeMB Then
            lngFileCount = objFolder.Files.Count
            
            If lngFileCount < 1 Then lngFileCount = 1
            ReDim strFileNames(lngFileCount - 1)
            ReDim lngPointerArray(lngFileCount - 1)
            ReDim FileAccessDates(lngFileCount - 1)
            
            lngFileCount = 0
            For Each objFile In objFolder.Files
                strFileNames(lngFileCount) = objFile.Name
                FileAccessDates(lngFileCount) = objFile.DateLastAccessed
                lngPointerArray(lngFileCount) = lngFileCount
                lngFileCount = lngFileCount + 1
            Next objFile
            
            ' Sort on FileAccessDates
            Set objQS = New QSDouble
            
            If objQS.QSAsc(FileAccessDates, lngPointerArray) Then
            
                ' Reduce the space usage to 75% of lngMaxFolderSizeMB
                For lngIndex = 0 To lngFileCount - 1
                    'Set objFile = fso.GetFile(fso.BuildPath(strFolderPath, strFileNames(lngPointerArray(lngIndex))))
                    'objFile.Size
                    
                    If LCase(fso.GetExtensionName(strFileNames(lngPointerArray(lngIndex)))) <> "txt" Then
                    
                        fso.DeleteFile fso.BuildPath(strFolderPath, strFileNames(lngPointerArray(lngIndex))), True
                        
                        If objFolder.Size / 1024 / 1024 <= lngMaxFolderSizeMB * 0.75 Then
                            Exit For
                        End If
                    End If
                    
                Next lngIndex
            Else
                Debug.Assert False
            End If
            
            Set objQS = Nothing
        End If
    End If
    
    Exit Sub
    
CheckFolderSpaceUsageErrorHandler:
    Debug.Assert False
    
End Sub

Private Sub DefineDefaultCachePath()
    
    Dim objFolder As Folder
    Dim strPath As String
    
On Error GoTo DefineCachePathErrorHandler

    Set objFolder = fso.GetFolder(fso.GetParentFolderName(App.Path))
    
    Do Until objFolder.IsRootFolder
        Set objFolder = objFolder.ParentFolder
    Loop
    
    mCachePath = fso.BuildPath(objFolder.Path, "MSSpectra")
    SetLocalErrorCode scNoError

    Exit Sub

DefineCachePathErrorHandler:
    Debug.Assert False
    
    If Not fso Is Nothing Then
        mCachePath = fso.BuildPath(fso.GetParentFolderName(App.Path), "MSSpectra")
    Else
        mCachePath = "."
    End If
    
    MsgBox "Error defining default spectra cache folder: " & Err.Description & vbCrLf & "Will use: " & mCachePath, vbExclamation + vbOKOnly, "Error"
    SetLocalErrorCode scInvalidCachePath
    
End Sub

Public Function ExtractAndCacheSpectrum(ByVal strPathToZippedSpectraFiles As String, ByVal strRawDataFileName As String, ByVal lngScanNumber As Long, ByVal strInputFilePath As String, ByRef strSpectrumPathOut As String, Optional ByVal strFilePrefix As String = "s") As Boolean

    Dim strZipFileName As String
    Dim strZipFilePath As String
    Dim strTargetFolder As String
    
    Dim strInputFileNameShortened As String
    Dim strRightChar As String
    Dim lngCharLoc As Long
    
    Dim blnSuccess As Boolean
    
    Dim strStatusFilePath As String
    Dim tsOutput As TextStream
        
    Static strLastInputFileName As String
    Static strLastInputFileNameShortened As String
    Static dtLastFileStatusUpdateDate As Date
    
    
On Error GoTo ExtractAndCacheSpectrumErrorHandler

    SetLocalErrorCode scNoError
    strSpectrumPathOut = ""

    If strInputFilePath = strLastInputFileName Then
        strInputFileNameShortened = strLastInputFileNameShortened
    Else
        ' Remove .Pek, .CSV, .mzXML, or .mzData from the Input file name and shorten, if needed
        
        strInputFileNameShortened = fso.GetBaseName(strInputFilePath)
        dtLastFileStatusUpdateDate = CDate("1/1/1900")
        
        If Len(strInputFileNameShortened) > MAX_CACHE_FOLDERNAME_LENGTH Then
            strInputFileNameShortened = Trim(Left(strInputFileNameShortened, MAX_CACHE_FOLDERNAME_LENGTH))
            
            If Len(strInputFileNameShortened) > MIN_CACHE_FOLDERNAME_LENGTH Then
                lngCharLoc = InStrRev(strInputFileNameShortened, "_")
                If lngCharLoc > MIN_CACHE_FOLDERNAME_LENGTH Then
                    strInputFileNameShortened = Left(strInputFileNameShortened, lngCharLoc - 1)
                Else
                    lngCharLoc = InStrRev(strInputFileNameShortened, "-")
                    If lngCharLoc > MIN_CACHE_FOLDERNAME_LENGTH Then
                        strInputFileNameShortened = Left(strInputFileNameShortened, lngCharLoc - 1)
                    End If
                End If
            End If
        End If
        strLastInputFileNameShortened = strInputFileNameShortened
        strLastInputFileName = strInputFilePath
    End If

    strTargetFolder = fso.BuildPath(mCachePath, strInputFileNameShortened)
    strSpectrumPathOut = fso.BuildPath(strTargetFolder, strRawDataFileName)

    If fso.FileExists(strSpectrumPathOut) Then
        ' Spectrum has already been extracted
        blnSuccess = True
    ElseIf fso.FolderExists(strPathToZippedSpectraFiles) Then
    
        ' Look for the appropriate .Zip file, based on lngScanNumber
        strZipFileName = GetZipFileNameForScan(lngScanNumber, strFilePrefix)
        strZipFilePath = fso.BuildPath(strPathToZippedSpectraFiles, strZipFileName)
        
        If fso.FileExists(strZipFilePath) Then
            
            ' Unzip strRawDataFileName from strZipFilePath to strSpectrumPathOut
            UnzipSetOptions
            
            blnSuccess = True
            If Not fso.FolderExists(strTargetFolder) Then
                On Error Resume Next
                fso.CreateFolder strTargetFolder
                If Err.Number <> 0 Then
                    SetLocalErrorCode scErrorCreatingCacheFolder
                    blnSuccess = False
                End If
            End If
            
            If blnSuccess Then
                ' Unzip strRawDataFileName from strZipFilePath

                blnSuccess = UnzipFile(strZipFilePath, strRawDataFileName, strTargetFolder)
                ' Could get Zip message using: UnzipGetZipInfo
                
                If Not blnSuccess Then
                    ' Just in case UnzipFile returned an odd error code, verify that the file doesn't exist
                    Debug.Assert Not fso.FileExists(fso.BuildPath(strTargetFolder, strRawDataFileName))
                End If
                
            End If
        Else
            SetLocalErrorCode scZipFilesNotFoundInDataFolder
        End If
    Else
        SetLocalErrorCode scOriginalDataFolderNotFound
    End If
    
    If blnSuccess Then
        ' Update the last cached folder variable
        mLastCachedFolderPath = strTargetFolder
        
        ' Update the date in the Last_Accessed.txt file if more than 5 minutes has elapsed
        If DateDiff("m", dtLastFileStatusUpdateDate, Now()) > 5 Then
            strStatusFilePath = fso.BuildPath(strTargetFolder, LAST_ACCESSED_FILENAME)
            
            On Error Resume Next
            Set tsOutput = fso.CreateTextFile(strStatusFilePath, True)
            tsOutput.WriteLine Now()
            tsOutput.Close
            
            If Err.Number = 0 Then
                dtLastFileStatusUpdateDate = Now()
                
                If mAutoDeleteOldCachedSpectra Then
                    ' Make sure the target folder size is < mMaxFolderSizeMB
                    CheckFolderSpaceUsage strTargetFolder
                End If
            Else
                Err.Clear
            End If
        
            On Error GoTo ExtractAndCacheSpectrumErrorHandler
        End If
    End If

    If DateDiff("d", mLastAutoPurgeDate, Now()) >= 3 And mAutoDeleteOldCachedSpectra Then
        ' Viper has been running for 3 days, and thus old spectra haven't had a chance
        '  to be deleted (they're normally deleted when Viper is closed)
        ' Auto-delete them now
        ' This is needed for people who always leave Viper running (if they even exist)
        RemoveOldSpectra mDaysToCacheSpectra
    End If
    
    ExtractAndCacheSpectrum = blnSuccess
    
    Exit Function

ExtractAndCacheSpectrumErrorHandler:
    SetLocalErrorCode scErrorCachingSpectrum
    ExtractAndCacheSpectrum = False
    
End Function

Public Function GetErrorMessage() As String
    ' Returns "" if no error

    If mLocalErrorCode = scNoError Then
        GetErrorMessage = ""
    ElseIf (mLocalErrorCode And scInvalidCachePath) = scInvalidCachePath Then
        GetErrorMessage = "Invalid spectrum cache path"
    ElseIf (mLocalErrorCode And scOriginalDataFolderNotFound) = scOriginalDataFolderNotFound Then
        GetErrorMessage = "Original data folder path not found"
    ElseIf (mLocalErrorCode And scZipFilesNotFoundInDataFolder) = scZipFilesNotFoundInDataFolder Then
        GetErrorMessage = "Zip files not found in original data folder; use the Proteomics Support Request Queue at http://proteomics.emsl.pnl.gov/ to request that the files be restored from the archive for the dataset."
    ElseIf (mLocalErrorCode And scErrorCachingSpectrum) = scErrorCachingSpectrum Then
        GetErrorMessage = "Error obtaining spectrum from .Zip file and caching locally"
    ElseIf (mLocalErrorCode And scErrorDeletingCachedSpectra) = scErrorDeletingCachedSpectra Then
        GetErrorMessage = "Error deleting cached spectra"
    ElseIf (mLocalErrorCode And scUnidentifiedError) = scUnidentifiedError Then
        GetErrorMessage = "Unknown error state"
    Else
        GetErrorMessage = "Unknown error state"
    End If

End Function

Public Function GetZipFileNameForScan(ByVal lngScanNumber As Long, Optional ByVal strFilePrefix As String = "s") As String
    
    Dim FolderNumber As Long
    
    If lngScanNumber Mod 100 > 0 Then
       FolderNumber = lngScanNumber \ 100 + 1
    Else
       FolderNumber = lngScanNumber \ 100
    End If
    
    GetZipFileNameForScan = strFilePrefix & Format(FolderNumber, "000") & ".zip"
    
End Function

Public Sub RemoveOldSpectra(Optional lngDaysToCache As Long = 0)

    If lngDaysToCache = 0 Then lngDaysToCache = mDaysToCacheSpectra
    
    If ValidateCacheFolder() Then
        If lngDaysToCache < 1 Then lngDaysToCache = 1
        RemoveOldSpectraWork mCachePath, lngDaysToCache
    End If

End Sub

Private Function RemoveOldSpectraWork(strFolderPath As String, lngDaysToCache As Long) As Boolean
    ' Returns True if this folder should be deleted

    Dim objFolder As Folder, objFile As File
    Dim objSubFolder As Folder
    
    Dim strStatusFilePath As String
    Dim strLineIn As String
    Dim tsTextStream As TextStream
    
    Dim blnDeleteFolder As Boolean
    
    blnDeleteFolder = False
    
On Error GoTo RemoveOldSpectraWorkErrorHandler

    strStatusFilePath = fso.BuildPath(strFolderPath, LAST_ACCESSED_FILENAME)
    If fso.FileExists(strStatusFilePath) Then
        
        Set tsTextStream = fso.OpenTextFile(strStatusFilePath)
        Do While Not tsTextStream.AtEndOfStream
            ' The first line should be a date
            ' If it isn't, keep reading until we find a line with a date
            ' If we don't find a date, then do not delete any files
            strLineIn = tsTextStream.ReadLine
            If IsDate(strLineIn) Then
                If DateDiff("d", CDate(strLineIn), Now()) > lngDaysToCache Then
                    blnDeleteFolder = True
                End If
                Exit Do
            End If
        Loop
        
    End If
    
    Set objFolder = fso.GetFolder(strFolderPath)
    
    If blnDeleteFolder Then
        RemoveOldSpectraWork = True
    Else
        For Each objSubFolder In objFolder.SubFolders
            ' Recursively call this function
            If RemoveOldSpectraWork(objSubFolder.Path, lngDaysToCache) Then
                ' Delete objSubFolder
                On Error Resume Next

                objSubFolder.Delete (True)
                If Err.Number <> 0 Then
                    MsgBox "Error deleting folder: " & objSubFolder.Path & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
                    SetLocalErrorCode scErrorDeletingCachedSpectra
                End If
                
                On Error GoTo RemoveOldSpectraWorkErrorHandler
            End If
        Next objSubFolder
        RemoveOldSpectraWork = False
    End If
    
    mLastAutoPurgeDate = Now()
    
    Exit Function
    
RemoveOldSpectraWorkErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error examining folder for old spectra: " & strFolderPath & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    SetLocalErrorCode scErrorDeletingCachedSpectra
    
End Function

Private Sub SetLocalErrorCode(ByVal eNewErrorCode As scSpectraCacheErrors, Optional ByVal blnLeaveExistingErrorCodeUnchanged As Boolean = False)

    If blnLeaveExistingErrorCodeUnchanged And mLocalErrorCode <> scSpectraCacheErrors.scNoError Then
        ' An error code is already defined; do not change it
    Else
        If eNewErrorCode = scSpectraCacheErrors.scNoError Then
            mLocalErrorCode = scSpectraCacheErrors.scNoError
        Else
            mLocalErrorCode = mLocalErrorCode Or eNewErrorCode
        End If

    End If

End Sub

Private Function ValidateCacheFolder(Optional ByVal strSubFolderName As String = "") As Boolean
    ' Returns True if mCachePath exists
    ' Will attempt to create it if it doesn't exist
    
    Dim blnExists As Boolean
    Dim strSubFolderPath As String
    
    If fso.FolderExists(mCachePath) Then
        blnExists = True
    Else
        On Error Resume Next
        fso.CreateFolder mCachePath
        blnExists = fso.FolderExists(mCachePath)
    End If
    
    If blnExists And Len(strSubFolderName) > 0 Then
        strSubFolderPath = fso.BuildPath(mCachePath, strSubFolderName)
        If Not fso.FolderExists(strSubFolderPath) Then
            On Error Resume Next
            fso.CreateFolder strSubFolderPath
            blnExists = fso.FolderExists(strSubFolderPath)
        End If
    
    End If
    
    ValidateCacheFolder = blnExists

End Function

Private Sub Class_Initialize()
    Set fso = New FileSystemObject
    
    SetLocalErrorCode scNoError
    
    DefineDefaultCachePath
    mDaysToCacheSpectra = 7
    mMaxFolderSizeMB = 100
    
    mAutoDeleteOldCachedSpectra = True
    
    mLastAutoPurgeDate = Now()
    
End Sub

Private Sub Class_Terminate()
    ' Delete old spectra from the specta cache folder
    If mAutoDeleteOldCachedSpectra Then
        RemoveOldSpectra mDaysToCacheSpectra
        
        If Len(mLastCachedFolderPath) > 0 Then
            CheckFolderSpaceUsage mLastCachedFolderPath
        End If
    End If
    
    Set fso = Nothing
End Sub

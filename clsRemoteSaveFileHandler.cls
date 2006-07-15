VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemoteSaveFileHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFilePathLocal As String
Private mFilePathRemote As String

Private fso As FileSystemObject

Private Sub Class_Initialize()
    Set fso = New FileSystemObject
End Sub

Public Sub DeleteTempFile()
    On Error Resume Next
    
    If fso.FileExists(mFilePathLocal) Then
        fso.DeleteFile mFilePathLocal, True
    End If
End Sub

Public Function GetTempFilePath(ByVal strFinalPath As String, blnForceTempFileCreation As Boolean) As String
    ' This function examines strFinalPath to see if it is on the same drive as App.Path
    ' If it is on a different drive, or if blnForceTempFileCreation = True, then this function
    '  will return a temporary file path using folder GetTemporaryDir()
    ' Otherwise, will simply return strFinalPath
    
    Dim objFile As File
    Dim strRandomSuffix As String
    
On Error GoTo GetTempFilePathErrorHandler

    ' Populate mFilePathRemote and mFilePathLocal with strFinalPath
    mFilePathRemote = strFinalPath
    mFilePathLocal = strFinalPath
    
    If LCase(fso.GetDriveName(mFilePathLocal)) <> LCase(fso.GetDriveName(App.Path)) Or blnForceTempFileCreation Then
        ' Re-Define mFilePathLocal to the temporary folder
        
        ' We could use App.Path as the temp folder, but this is cleaner
        mFilePathLocal = fso.BuildPath(GetTemporaryDir(), fso.GetFileName(mFilePathLocal))
        If fso.FileExists(mFilePathLocal) Then
            ' Existing file found; if the modfication time is more than 5 minutes old, then delete it
            ' Otherwise, uniquify mFilePathLocal
            
            Set objFile = fso.GetFile(mFilePathLocal)
            If DateDiff("n", objFile.DateLastModified, Now()) >= 5 And _
               DateDiff("n", objFile.DateCreated, Now()) >= 5 Then
               ' Delete the file
                fso.DeleteFile mFilePathLocal
            
                ' Pause 200 msec, just to be safe
                Sleep 200
            Else
                ' Uniquify mFilePathLocal
                mFilePathLocal = fso.BuildPath(fso.GetParentFolderName(mFilePathLocal), fso.GetBaseName(mFilePathLocal))
                strRandomSuffix = "_" & Mid(Format(Rnd(), "0.0000000"), 3)
                
                mFilePathLocal = mFilePathLocal & strRandomSuffix & "." & fso.GetExtensionName(strFinalPath)
            End If
            
        End If
    End If

    GetTempFilePath = mFilePathLocal
    Exit Function

GetTempFilePathErrorHandler:
    Debug.Print "Error in GetTempFilePath: " & Err.Description
    Debug.Assert False
    GetTempFilePath = ""

End Function

Public Function MoveTempFileToFinalDestination() As Boolean
    Dim blnSuccess As Boolean

On Error GoTo MoveTempFileErrorHandler
    
    If LCase(mFilePathLocal) = LCase(mFilePathRemote) Then
        ' Paths are identical; nothing to move
        blnSuccess = True
    Else
        ' Need to copy the file to the destination
        fso.CopyFile mFilePathLocal, mFilePathRemote, True
        Sleep 100
        
        ' Delete the temporary file
        On Error Resume Next
        fso.DeleteFile mFilePathLocal, True
        Sleep 100
        
        blnSuccess = True
    End If
    
    MoveTempFileToFinalDestination = blnSuccess
    Exit Function

MoveTempFileErrorHandler:
    Debug.Print "Error in MoveTempFileToFinalDestination: " & Err.Description
    Debug.Assert False
    MoveTempFileToFinalDestination = False

End Function


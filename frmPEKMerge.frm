VERSION 5.00
Begin VB.Form frmPEKMerge 
   BackColor       =   &H80000001&
   Caption         =   "Merge PEK Files"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "frmMerge"
   ScaleHeight     =   4920
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAutoDefineOutput 
      Caption         =   "Auto Define"
      Height          =   375
      Left            =   7440
      TabIndex        =   19
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSortList 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   2220
      Width           =   975
   End
   Begin VB.Timer tmrTimer 
      Left            =   8280
      Top             =   4440
   End
   Begin VB.CommandButton cmdPasteFileList 
      Caption         =   "Paste File List"
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   2220
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemoveSelected 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move &Down"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move &Up"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectOutputFile 
      Caption         =   "Select Output"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraMergeOptions 
      BackColor       =   &H80000001&
      Caption         =   "Merge Order"
      ForeColor       =   &H80000009&
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   2895
      Begin VB.OptionButton optMOrder 
         BackColor       =   &H80000001&
         Caption         =   "Interleave by scan"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton optMOrder 
         BackColor       =   &H80000001&
         Caption         =   "Sequential (renumber scans on files after first file if necessary)"
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "&Merge"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      ToolTipText     =   "0"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Source"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2220
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ListBox lstSources 
      Height          =   2010
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   7095
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "MergedData.pek"
      Top             =   270
      Width           =   5775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Progress."
      ForeColor       =   &H80000009&
      Height          =   1335
      Left            =   3120
      TabIndex        =   15
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Files:"
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblDestinationFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination File:"
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   300
      Width           =   1335
   End
End
Attribute VB_Name = "frmPEKMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 01/23/2001 nt
'this are possible combinations of Order and First
'Scan selection
'0,<0   go file by file from the list; start with indicated
'       file number as a first scan; always fill gaps
'0,1    go file by file from the list; take first scan from
'       the first file in thr list; use original numbers as
'       long as you can
'1,<0   alternate; take scan from each file start with
'       indicated file number; always fill gaps
'1,1    alternate; keep original scan numbers as long as possible
'2,any  open all original files and set scans in order they
'       they should come based on found scan numbers; if conflict
'       indicate that this merge option can not be used
Option Explicit

Private Const MERGE_ORDER_SEQUENTIAL = 0
Private Const MERGE_ORDER_INTERLEAVE = 1

Private Const PEK_FILE_EXTENSION As String = "pek"
Private Const CSV_FILE_EXTENSION As String = "csv"

Private Const DECON2LS_ISOS_FILE_SUFFIX As String = "_isos.csv"
Private Const DECON2LS_SCANS_FILE_SUFFIX As String = "_scans.csv"

Private Type udtFileInfoType
    FullPath As String
    CompactPath As String
End Type

Private mFso As FileSystemObject
Private mMergeOrder As Long

Private mCompactPathLength As Integer

Private mSourceFileListCount As Integer
Private mSourceFileList() As udtFileInfoType

Private mAbortMerge As Boolean
Private mFormResized As Boolean

Private mUsingDecon2LSFiles As Boolean

Private Sub AppendFileContents(ByRef tsIsosIn As TextStream, ByRef tsIsosOut As TextStream, _
                        ByRef objScanNumberTracker As clsScanNumberTracker, ByVal blnUseScanNumberTracker As Boolean, ByVal blnIsScansFile As Boolean, _
                        ByVal lngScanNumberAddition As Long, ByRef lngScanNumberStartCurrentFile As Long, ByRef lngScanNumberMaxCurrentFile As Long, _
                        ByVal sngScanTimeAddition As Single, ByRef sngScanTimeStartCurrentFile As Single, ByRef sngScanTimeMaxCurrentFile As Single, _
                        ByRef lngLinesRead As Long, ByRef lngBytesRead As Long, ByVal lngTotalBytes As Long, _
                        ByVal intFileNum As Integer, intIsosFileCount As Integer)
    
    Dim strLineIn As String
    Dim strRemainder As String
    Dim strRemainder2 As String
    
    Dim lngScanNum As Long
    Dim lngLinesReadCurrentFile As Long
    Dim lngScansReadCurrentFile As Long
    
    Dim sngScanTime As Single
    
    lngLinesReadCurrentFile = 0
    lngScansReadCurrentFile = 0
    strLineIn = ""
    
    If blnUseScanNumberTracker Then
        objScanNumberTracker.Reset
    End If

    Do While Not tsIsosIn.AtEndOfStream
        strLineIn = tsIsosIn.ReadLine
        
        lngLinesRead = lngLinesRead + 1
        lngLinesReadCurrentFile = lngLinesReadCurrentFile + 1
        lngBytesRead = lngBytesRead + Len(strLineIn) + 2
        
        If lngLinesReadCurrentFile = 1 Then
            ' Check for the header line
            If Not IsNumeric(Left(strLineIn, 1)) Then
                ' Header line
                If intFileNum = 1 Then
                    tsIsosOut.WriteLine strLineIn
                End If
                
                strLineIn = ""
            End If
        End If
        
        If Len(strLineIn) > 0 Then
            ' Parse out the scan number
            If ExtractFirstLng(strLineIn, lngScanNum, strRemainder) Then
                
                If lngScanNumberStartCurrentFile = 0 Then
                    lngScanNumberStartCurrentFile = lngScanNum
                    lngScanNumberMaxCurrentFile = lngScanNum
                Else
                    If lngScanNum > 0 And lngScanNum < lngScanNumberStartCurrentFile Then
                        lngScanNumberStartCurrentFile = lngScanNum
                    End If
                    
                    If lngScanNum > lngScanNumberMaxCurrentFile Then
                        lngScanNumberMaxCurrentFile = lngScanNum
                    End If
                End If

                If blnIsScansFile Then
                    ' Parse out the scan time and bump up accordingly
                    
                    If ExtractFirstSng(Mid(strRemainder, 2), sngScanTime, strRemainder2) Then
                        If sngScanTimeStartCurrentFile = 0 Then
                            sngScanTimeStartCurrentFile = sngScanTime
                            sngScanTimeMaxCurrentFile = sngScanTime
                        Else
                            If sngScanTime > 0 And sngScanTime < sngScanTimeStartCurrentFile Then
                                sngScanTimeStartCurrentFile = sngScanTime
                            End If
                            
                            If sngScanTime > sngScanTimeMaxCurrentFile Then
                                sngScanTimeMaxCurrentFile = sngScanTime
                            End If
                        End If
                        
                        strRemainder = "," & LTrim(RTrim(sngScanTime + sngScanTimeAddition)) & strRemainder2
                    End If
                End If
                
                

                tsIsosOut.WriteLine LTrim(RTrim(lngScanNum + lngScanNumberAddition)) & strRemainder
                lngScansReadCurrentFile = lngScansReadCurrentFile + 1
                
                If blnUseScanNumberTracker Then
                    objScanNumberTracker.AddScanNumberAndUpdateAverageIncrement lngScanNum
                End If
            End If
        End If
        
        If lngLinesRead Mod 1000 = 0 Then
            If mAbortMerge Then Exit Do
            lblProgress.Caption = "Merging files: " & intFileNum & " of " & intIsosFileCount & " -- " & Round(lngBytesRead / CDbl(lngTotalBytes) * 100, 0) & "% complete"
            DoEvents
        End If
    Loop
        
End Sub

Private Sub AppendSourceFile()
    Dim strFilePath As String
    Dim strExtension As String
    
    strExtension = GetPreferredFileExtension()
    
    strFilePath = SelectFile(hwnd, "Select source file", , False, "*." & strExtension, GetFilterCodesPekAndCSV(), 1)
    AppendSourceFileWork strFilePath
End Sub

Private Sub AppendSourceFileWork(ByVal strFilePath As String)
    Dim blnDuplicateFile As Boolean
    Dim intIndex As Integer
    
    Dim blnDecon2LSFile As Boolean
    Dim strBaseName As String
    Dim strFileSuffix As String
    Dim blnValidSuffix As Boolean
    
On Error GoTo AppendSourceFileWorkErrorHandler

    If Len(strFilePath) > 0 Then
        ' Only add file to mSourceFileList if not already present
        ' When processing Decon2LS CSV files, only add _isos.csv files (if the user chooses a _scans.csv file,
        '  look for and auto-add the _isos.csv files
        
        blnDecon2LSFile = IsDecon2LSCSVFile(strFilePath)
        If blnDecon2LSFile Then
            strBaseName = GetDecon2LSBaseName(strFilePath, strFileSuffix)
            
            If Len(strBaseName) > 0 Then
                If strFileSuffix = DECON2LS_ISOS_FILE_SUFFIX Then
                    ' Valid suffix; leave strFilePath unchanged
                ElseIf strFileSuffix = DECON2LS_SCANS_FILE_SUFFIX Then
                    strFilePath = mFso.BuildPath(mFso.GetParentFolderName(strFilePath), strBaseName & DECON2LS_ISOS_FILE_SUFFIX)
                Else
                    MsgBox "Unknown suffix (" & strFileSuffix & ") for file " & strFilePath
                    Exit Sub
                End If
            End If
        End If
        
        For intIndex = 0 To mSourceFileListCount - 1
            If blnDecon2LSFile Then
                If Decon2LSCSVFilesMatch(strFilePath, mSourceFileList(intIndex).FullPath) Then
                    blnDuplicateFile = True
                    Exit For
                End If
            Else
                If LCase(strFilePath) = LCase(mSourceFileList(intIndex).FullPath) Then
                    blnDuplicateFile = True
                    Exit For
                End If
            End If
        Next intIndex
        
        If Not blnDuplicateFile Then
            If UBound(mSourceFileList) <= mSourceFileListCount Then
                ReDim Preserve mSourceFileList((UBound(mSourceFileList) + 1) * 2 - 1)
            End If
            
            With mSourceFileList(mSourceFileListCount)
                .FullPath = strFilePath
                .CompactPath = CompactPathString(.FullPath, mCompactPathLength)
                
                lstSources.AddItem .CompactPath
            End With
            mSourceFileListCount = mSourceFileListCount + 1
                        
            cmdMerge.Enabled = True
            
            mUsingDecon2LSFiles = blnDecon2LSFile
        End If
    End If

    Exit Sub

AppendSourceFileWorkErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmPEKMerge->AppendSourceFileWork", Err.Description
    
End Sub

Private Sub AutoDefineOutputFileName()
    ' Look for the common text between the first and second files and define the output file name based on that
    
    Dim strFolderPath As String
    Dim strOutputFilePath As String
    Dim strComparisonFilePath As String
    Dim intIndex As Integer
    Dim strExtension As String
    
    If mSourceFileListCount > 0 Then
        strFolderPath = mFso.GetParentFolderName(mSourceFileList(0).FullPath)
        strExtension = GetPreferredFileExtension()
                
        If mSourceFileListCount <= 1 Then
            strOutputFilePath = mFso.GetBaseName(mSourceFileList(0).FullPath) & "_Merged." & strExtension
        Else
            strOutputFilePath = mFso.GetBaseName(mSourceFileList(0).FullPath)
            strComparisonFilePath = mFso.GetBaseName(mSourceFileList(1).FullPath)
            For intIndex = 1 To Len(strOutputFilePath)
                If LCase(Mid(strOutputFilePath, intIndex, 1)) <> LCase(Mid(strComparisonFilePath, intIndex, 1)) Then
                    Exit For
                End If
            Next intIndex
            
            If intIndex > 1 Then
                strOutputFilePath = Left(strOutputFilePath, intIndex - 1)
                ' Decrement intIndex until the first underscore
                intIndex = InStrRev(strOutputFilePath, "_")
                If intIndex > 2 Then
                    strOutputFilePath = Left(strOutputFilePath, intIndex - 1)
                End If
            End If
        End If
        
        strOutputFilePath = mFso.BuildPath(strFolderPath, strOutputFilePath & "_Merged." & strExtension)
        
        txtDestination = strOutputFilePath
    End If
End Sub

Private Sub ClearSourceFileList()
    mSourceFileListCount = 0
    lstSources.Clear
End Sub

Private Function Decon2LSCSVFilesMatch(ByVal strFilePath1 As String, ByVal strFilePath2 As String) As Boolean
    Dim strFileName1 As String
    Dim strFileName2 As String
    Dim blnMatch As Boolean
    
    blnMatch = False
    If IsDecon2LSCSVFile(strFilePath1) And IsDecon2LSCSVFile(strFilePath2) Then
        strFileName1 = mFso.GetFileName(strFilePath1)
        strFileName2 = mFso.GetFileName(strFilePath2)

        strFileName1 = StringTrimEnd(strFileName1, DECON2LS_ISOS_FILE_SUFFIX)
        strFileName1 = StringTrimEnd(strFileName1, DECON2LS_SCANS_FILE_SUFFIX)
        
        strFileName2 = StringTrimEnd(strFileName2, DECON2LS_ISOS_FILE_SUFFIX)
        strFileName2 = StringTrimEnd(strFileName2, DECON2LS_SCANS_FILE_SUFFIX)
        
        If LCase(strFileName1) = LCase(strFileName2) Then
            blnMatch = True
        End If
    End If
    
    Decon2LSCSVFilesMatch = blnMatch
    
End Function

Private Sub EnableDisableControls(ByVal blnEnableControls As Boolean)
    If blnEnableControls Then
        Me.MousePointer = vbDefault
    Else
        Me.MousePointer = vbHourglass
    End If
    
    cmdSelectOutputFile.Enabled = blnEnableControls
    cmdAutoDefineOutput.Enabled = blnEnableControls
    
    cmdMerge.Enabled = blnEnableControls
    cmdClose.Visible = blnEnableControls
    cmdCancel.Visible = Not blnEnableControls
    
    cmdAdd.Enabled = blnEnableControls
    cmdPasteFileList.Enabled = blnEnableControls
    cmdSortList.Enabled = blnEnableControls
    
    cmdMoveUp.Enabled = blnEnableControls
    cmdMoveDown.Enabled = blnEnableControls
    cmdRemoveSelected.Enabled = blnEnableControls
    cmdClear.Enabled = blnEnableControls
    
    fraMergeOptions.Enabled = blnEnableControls
End Sub

Private Function ExtractFirstLng(ByVal strText As String, ByRef lngNumber As Long, ByRef strRemainder As String) As Boolean
    Dim intCommaLoc As Integer
    Dim blnFound As Boolean
    
    blnFound = False
    
    intCommaLoc = InStr(strText, ",")
    If intCommaLoc > 0 Then
        If IsNumeric(Left(strText, intCommaLoc - 1)) Then
            lngNumber = CLng(Left(strText, intCommaLoc - 1))
            strRemainder = Mid(strText, intCommaLoc)
            blnFound = True
        End If
    End If

    ExtractFirstLng = blnFound
End Function

Private Function ExtractFirstSng(ByVal strText As String, ByRef sngNumber As Single, ByRef strRemainder As String) As Boolean
    Dim intCommaLoc As Integer
    Dim blnFound As Boolean
    
    blnFound = False
    
    intCommaLoc = InStr(strText, ",")
    If intCommaLoc > 0 Then
        If IsNumeric(Left(strText, intCommaLoc - 1)) Then
            sngNumber = CSng(Left(strText, intCommaLoc - 1))
            strRemainder = Mid(strText, intCommaLoc)
            blnFound = True
        End If
    End If

    ExtractFirstSng = blnFound
End Function

Private Function GetDecon2LSBaseName(ByVal strFilePath As String, ByRef strFileSuffix As String) As String
    Dim strFileName As String
    Dim strBase As String
    
    strBase = ""
    strFileSuffix = ""
    
    strFileName = mFso.GetFileName(strFilePath)
    
    If StringEndsWith(strFileName, DECON2LS_ISOS_FILE_SUFFIX) Then
        strBase = StringTrimEnd(strFileName, DECON2LS_ISOS_FILE_SUFFIX)
        strFileSuffix = DECON2LS_ISOS_FILE_SUFFIX
    ElseIf StringEndsWith(strFileName, DECON2LS_SCANS_FILE_SUFFIX) Then
        strBase = StringTrimEnd(strFileName, DECON2LS_SCANS_FILE_SUFFIX)
        strFileSuffix = DECON2LS_SCANS_FILE_SUFFIX
    Else
        strBase = mFso.GetBaseName(strFileName)
    End If
    
    GetDecon2LSBaseName = strBase
    
End Function

Private Function GetPreferredFileExtension() As String

    Dim strExtension As String
    
    If mSourceFileListCount > 0 Then
        strExtension = mFso.GetExtensionName(mSourceFileList(0).FullPath)
    Else
        If mUsingDecon2LSFiles Then
            strExtension = CSV_FILE_EXTENSION
        Else
            strExtension = PEK_FILE_EXTENSION
        End If
    End If
    
    If Len(strExtension) = 0 Then
        strExtension = PEK_FILE_EXTENSION
    End If

    GetPreferredFileExtension = strExtension
    
End Function
    
Private Function GetFilterCodes(ByVal strExtension As String) As String
    
    If Len(strExtension) = 0 Then
        strExtension = PEK_FILE_EXTENSION
    End If
    
    GetFilterCodes = "All Files (*.*)|*.*|" & UCase(strExtension) & " Files (*." & LCase(strExtension) & ")|*." & LCase(strExtension)
End Function

Private Function GetFilterCodesPekAndCSV() As String
    GetFilterCodesPekAndCSV = "All Files (*.*)|*.*|" & UCase(PEK_FILE_EXTENSION) & " Files (*." & LCase(PEK_FILE_EXTENSION) & ")|*." & LCase(PEK_FILE_EXTENSION) & "|" & UCase(CSV_FILE_EXTENSION) & " Files (*." & LCase(CSV_FILE_EXTENSION) & ")|*." & LCase(CSV_FILE_EXTENSION)
End Function


Private Sub InitializeForm()
    
    Set mFso = New FileSystemObject
    
    mMergeOrder = 0
    
    optMOrder(0).Value = True
    lblProgress.Caption = ""

    mUsingDecon2LSFiles = False

    mAbortMerge = False
    mFormResized = False
    
    mCompactPathLength = 40
    mSourceFileListCount = 0
    ReDim mSourceFileList(9)
    
    PositionControls
    RefreshSourceFileListbox True
    
    With tmrTimer
        .Interval = 500
        .Enabled = True
    End With
End Sub

Private Function IsDecon2LSCSVFile(ByVal strFileName As String) As Boolean
    Dim blnDecon2LSFile As Boolean
    
    blnDecon2LSFile = StringEndsWith(strFileName, DECON2LS_ISOS_FILE_SUFFIX)
  
    If Not blnDecon2LSFile Then
        blnDecon2LSFile = StringEndsWith(strFileName, DECON2LS_SCANS_FILE_SUFFIX)
    End If
    
    IsDecon2LSCSVFile = blnDecon2LSFile

End Function

Private Sub MergeFilesStart()

    If mSourceFileListCount < 2 Then
        MsgBox "Please define at least two input files.", vbExclamation Or vbOKOnly, "Error"
        Exit Sub
    End If
    
    If Len(txtDestination.Text) = 0 Then
        MsgBox "Please define the destination file", vbExclamation Or vbOKOnly, "Error"
        Exit Sub
    End If

    Select Case mMergeOrder
    Case MERGE_ORDER_SEQUENTIAL
        If IsDecon2LSCSVFile(mSourceFileList(0).FullPath) Then
            Call MergeCSVFiles
        Else
            Call MergePEKFiles
        End If
    Case MERGE_ORDER_INTERLEAVE
        MsgBox "Interleaved scan mode is not yet implemented.", vbExclamation + vbOKOnly, glFGTU
    Case Else
        MsgBox "Unknown merge order: " & mMergeOrder, vbExclamation + vbOKOnly, glFGTU
    End Select
End Sub

Private Function MergeCSVFiles() As Boolean
    '-----------------------------------------------------
    ' Reads CSV files in order and merges them together
    '-----------------------------------------------------
    
    Dim tsIsosIn As TextStream
    Dim tsScansIn As TextStream
    
    Dim tsIsosOut As TextStream
    Dim tsScansOut As TextStream
    
    Dim objFile As File
    
    Dim strLineIn As String
    
    Dim strOutputFile As String
    Dim strParentFolderPath As String
    
    Dim strIsosOutputFile As String
    Dim strScansOutputFile As String

    Dim i As Long, j As Long
    Dim intTargetIndex As Long
    
    Dim lngLinesRead As Long
    Dim lngLinesReadCurrentFile As Long
    Dim lngFileCountProcessed As Long
    
    Dim lngTotalBytes As Long
    Dim lngTotalBytesScans As Long
    Dim lngBytesRead As Long
    
    Dim strFilePath As String
    
    Dim lngScanNumberAddition As Long           ' Additional scan number to add to each scan number; necessary for files after the first one if the scan values in the subsequent files are less than the final scan in the first file
    Dim lngScanNumberStartCurrentFile As Long
    Dim lngScanNumberMaxCurrentFile As Long
    
    Dim sngScanTimeAddition As Single
    Dim sngScanTimeStartCurrentFile As Single
    Dim sngScanTimeMaxCurrentFile As Single

    Dim blnAddScanNumberAddition As Boolean
    
    Dim objScanNumberTracker As clsScanNumberTracker
    
    Dim intIsosFileCount As Integer
    Dim strSourceIsosFileList() As String
    Dim strSourceScansFileList() As String
    Dim blnScansFilesAllFound As Boolean
    
    Dim strFileSuffix As String
    Dim strBaseName As String
    
    Dim blnValidSuffix As Boolean
    Dim blnDuplicate As Boolean
    Dim blnProcessScansFiles As Boolean
    Dim blnUseScanNumberTracker As Boolean
    
On Error GoTo err_MergeCSVFiles
   
    mAbortMerge = False
    EnableDisableControls False
    
    intIsosFileCount = 0
    ReDim strSourceIsosFileList(mSourceFileListCount - 1)
    ReDim strSourceScansFileList(mSourceFileListCount - 1)
    blnScansFilesAllFound = False
    
    strOutputFile = Trim(txtDestination.Text)
    strParentFolderPath = mFso.GetParentFolderName(strOutputFile)
    
    strBaseName = GetDecon2LSBaseName(strOutputFile, strFileSuffix)
    strIsosOutputFile = mFso.BuildPath(strParentFolderPath, strBaseName & DECON2LS_ISOS_FILE_SUFFIX)
    strScansOutputFile = mFso.BuildPath(strParentFolderPath, strBaseName & DECON2LS_SCANS_FILE_SUFFIX)
        
    
    ' Validate that each file exists and determine the total bytes to be read
    ' Also, make sure none of the input files is the same as strIsosOutputFile or strScansOutputFile
    '
    ' We have to find every _isos.csv file
    ' If we also find all of the _scans.csv files, then we'll merge them too;
    ' Otherwise, we'll create a generic _scans.csv file
    
    lblProgress.Caption = "Finding Decon2LS CSV files to read"
    DoEvents
    
    For i = 0 To mSourceFileListCount - 1
        strFilePath = Trim(mSourceFileList(i).FullPath)
        
        If Not IsDecon2LSCSVFile(strFilePath) Then
            MsgBox "File is not a Decon2LS CSV file: " & strFilePath, vbExclamation + vbOKOnly, "Error"
            EnableDisableControls True
            MergeCSVFiles = False
            Exit Function
        End If
        
        strBaseName = GetDecon2LSBaseName(strFilePath, strFileSuffix)
        
        If Len(strBaseName) > 0 Then
            blnValidSuffix = False
            
            If strFileSuffix = DECON2LS_ISOS_FILE_SUFFIX Then
                strSourceIsosFileList(intIsosFileCount) = strFilePath
                strSourceScansFileList(intIsosFileCount) = mFso.BuildPath(strParentFolderPath, strBaseName & DECON2LS_SCANS_FILE_SUFFIX)
                blnValidSuffix = True
            ElseIf strFileSuffix = DECON2LS_SCANS_FILE_SUFFIX Then
                strSourceScansFileList(intIsosFileCount) = strFilePath
                strSourceIsosFileList(intIsosFileCount) = mFso.BuildPath(strParentFolderPath, strBaseName & DECON2LS_ISOS_FILE_SUFFIX)
                blnValidSuffix = True
            Else
                MsgBox "Unknown suffix (" & strFileSuffix & ") for file " & strFilePath
            End If
            
            If blnValidSuffix Then
                ' Make sure the isos file exists and make sure that it's not the same as the output files
                If Not mFso.FileExists(strSourceIsosFileList(intIsosFileCount)) Then
                    MsgBox "File not found: " & strSourceIsosFileList(intIsosFileCount), vbExclamation + vbOKOnly, "Error"
                    EnableDisableControls True
                    MergeCSVFiles = False
                    Exit Function
                ElseIf LCase(strIsosOutputFile) = LCase(strSourceIsosFileList(intIsosFileCount)) Or LCase(strScansOutputFile) = LCase(strSourceIsosFileList(intIsosFileCount)) Then
                    MsgBox "File " & strSourceIsosFileList(intIsosFileCount) & " is defined as both an input and an output file; this is not allowed.", vbExclamation + vbOKOnly, "Error"
                    EnableDisableControls True
                    MergeCSVFiles = False
                    Exit Function
                ElseIf LCase(strScansOutputFile) = LCase(strSourceScansFileList(intIsosFileCount)) Or LCase(strScansOutputFile) = LCase(strSourceScansFileList(intIsosFileCount)) Then
                    MsgBox "File " & strSourceScansFileList(intIsosFileCount) & " is defined as both an input and an output file; this is not allowed.", vbExclamation + vbOKOnly, "Error"
                    EnableDisableControls True
                    MergeCSVFiles = False
                    Exit Function
                End If
                
                intIsosFileCount = intIsosFileCount + 1
            End If
        End If
    Next i
    
    ' Make sure no duplicates exist in strSourceIsosFileList
    ' Using copy-in-place to compress the array
    intTargetIndex = 0
    For i = 0 To intIsosFileCount - 2
        blnDuplicate = False
        
        For j = i + 1 To intIsosFileCount - 1
            If LCase(strSourceIsosFileList(j)) = LCase(strSourceIsosFileList(i)) Then
                blnDuplicate = True
                Exit For
            End If
        Next j
        
        If Not blnDuplicate Then
            If intTargetIndex <> i Then
                strSourceIsosFileList(intTargetIndex) = strSourceIsosFileList(i)
                strSourceScansFileList(intTargetIndex) = strSourceScansFileList(i)
            End If
            intTargetIndex = intTargetIndex + 1
        End If
    Next i
    
    intIsosFileCount = intTargetIndex + 1
    
    If intIsosFileCount < 2 Then
        MsgBox "Could not find two distinct Isos files to read.", vbExclamation Or vbOKOnly, "Error"
        MergeCSVFiles = False
        Exit Function
    End If
    
    ' Determine the total number of bytes to be combined
    ' At the same time, see if the Scans files are all present
    
    
    lblProgress.Caption = "Determining total size of the input files"
    DoEvents
    
    lngTotalBytes = 0
    lngTotalBytesScans = 0
    blnProcessScansFiles = True
    
    For i = 0 To intIsosFileCount - 1
        If blnProcessScansFiles Then
            If Not mFso.FileExists(strSourceScansFileList(i)) Then
                blnProcessScansFiles = False
            Else
                Set objFile = mFso.GetFile(strSourceScansFileList(i))
                lngTotalBytesScans = lngTotalBytesScans + objFile.Size
            End If
        End If
    
        Set objFile = mFso.GetFile(strSourceIsosFileList(i))
        lngTotalBytes = lngTotalBytes + objFile.Size
    Next i
    
    If blnProcessScansFiles Then
        lngTotalBytes = lngTotalBytes + lngTotalBytesScans
    End If
    
    If lngTotalBytes < 1 Then lngTotalBytes = 1
    
    ' Create the output files
    lblProgress.Caption = "Creating output files"
    DoEvents
    
    Set tsIsosOut = mFso.OpenTextFile(strIsosOutputFile, ForWriting, True)
    If blnProcessScansFiles Then
        Set tsScansOut = mFso.OpenTextFile(strScansOutputFile, ForWriting, True)
    End If
    
    ' Open each file in strSourceIsosFileList() and strSourceScansFileList() and copy each to the output file
    
    lngFileCountProcessed = 0
    lngLinesRead = 0
    lngBytesRead = 0
    
    lngScanNumberAddition = 0
    sngScanTimeAddition = 0
    
    blnAddScanNumberAddition = False
    
    Set objScanNumberTracker = New clsScanNumberTracker
    
    lblProgress.Caption = "Merging files: 0% complete"
    
    For i = 0 To intIsosFileCount - 1
        lngScanNumberStartCurrentFile = 0
        lngScanNumberMaxCurrentFile = 0
        
        sngScanTimeStartCurrentFile = 0
        sngScanTimeMaxCurrentFile = 0
        
        Set tsIsosIn = mFso.OpenTextFile(strSourceIsosFileList(i), ForReading, True)
        
        If i = 0 Then
            blnUseScanNumberTracker = True
        Else
            blnUseScanNumberTracker = False
        End If
        
        AppendFileContents tsIsosIn, tsIsosOut, objScanNumberTracker, blnUseScanNumberTracker, False, _
                           lngScanNumberAddition, lngScanNumberStartCurrentFile, lngScanNumberMaxCurrentFile, _
                           sngScanTimeAddition, sngScanTimeStartCurrentFile, sngScanTimeMaxCurrentFile, _
                           lngLinesRead, lngBytesRead, lngTotalBytes, i + 1, intIsosFileCount
        tsIsosIn.Close
         
        blnUseScanNumberTracker = False
        If blnProcessScansFiles Then
            Set tsScansIn = mFso.OpenTextFile(strSourceScansFileList(i), ForReading, True)
            
            AppendFileContents tsScansIn, tsScansOut, objScanNumberTracker, blnUseScanNumberTracker, True, _
                               lngScanNumberAddition, lngScanNumberStartCurrentFile, lngScanNumberMaxCurrentFile, _
                               sngScanTimeAddition, sngScanTimeStartCurrentFile, sngScanTimeMaxCurrentFile, _
                               lngLinesRead, lngBytesRead, lngTotalBytes, i + 1, intIsosFileCount
                               
            tsScansIn.Close
        End If
               
        lngFileCountProcessed = lngFileCountProcessed + 1
        
        ' Bump up lngScanNumberAddition based on the maximum scan number in the current file
        If blnProcessScansFiles Then
            lngScanNumberAddition = lngScanNumberAddition + lngScanNumberMaxCurrentFile
            
            If CInt(sngScanTimeMaxCurrentFile) - sngScanTimeMaxCurrentFile > 0 Then
                If CInt(sngScanTimeMaxCurrentFile) - sngScanTimeMaxCurrentFile < 0.05 Then
                    ' Round up sngScanTimeMaxCurrentFile to the nearest integer
                    sngScanTimeMaxCurrentFile = CInt(sngScanTimeMaxCurrentFile)
                End If
                
                ' Round up, to the nearest integer, then subtract lngScanNumberStartCurrentFile
            End If
            
            sngScanTimeAddition = sngScanTimeAddition + sngScanTimeMaxCurrentFile
        Else
            lngScanNumberAddition = lngScanNumberAddition + Round(lngScanNumberMaxCurrentFile - lngScanNumberStartCurrentFile + objScanNumberTracker.AverageScanIncrement, 0)
        End If
        
        If mAbortMerge Then Exit For
    Next i
    tsIsosOut.Close
    If blnProcessScansFiles Then
        tsScansOut.Close
    End If
    
    If mAbortMerge Then
        lblProgress = "Merge aborted."
    Else
        lblProgress = "Merged " & Trim(lngFileCountProcessed) & " files into " & mFso.GetBaseName(strOutputFile) & " at " & mFso.GetParentFolderName(strOutputFile)
    End If
    
exit_MergeCSVFiles:
    EnableDisableControls True
    Set tsIsosIn = Nothing
    Set tsIsosOut = Nothing

    MergeCSVFiles = Not mAbortMerge
    Exit Function
    
err_MergeCSVFiles:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly
    mAbortMerge = True
    Resume exit_MergeCSVFiles
    
End Function

Private Function MergePEKFiles() As Boolean
    '-----------------------------------------------------
    ' Reads PEK files in order and merges them together
    '-----------------------------------------------------
    
    Const FILENAME_AKA_SCAN_NUMBER As String = "FILENAME"
    Dim intFilenameTextLength As Integer
    intFilenameTextLength = Len(FILENAME_AKA_SCAN_NUMBER)
    
    Dim tsIn As TextStream
    Dim tsOut As TextStream
    Dim objFile As File
    
    Dim strLineIn As String
    Dim strTextBeforeScanNumber As String
    Dim strTextAfterScanNumber As String
    
    Dim strOutputFile As String
    
    Dim i As Long
    Dim lngLinesRead As Long
    Dim lngFileCountProcessed As Long
    Dim lngScansReadCurrentFile As Long
    
    Dim lngTotalBytes As Long
    Dim lngBytesRead As Long
    
    Dim strInputFilePath As String
    
    Dim lngScanNumberCurrent As Long
    Dim lngScanNumberPrevious As Long
    Dim lngScanNumberAddition As Long           ' Additional scan number to add to each scan number; necessary for files after the first one if the scan values in the subsequent files are less than the final scan in the first file
    Dim lngScanNumberStartCurrentFile As Long
    
    Dim blnAddScanNumberAddition As Boolean
    
    Dim objScanNumberTracker As clsScanNumberTracker
    
On Error GoTo err_MergePEKFiles
   
    mAbortMerge = False
    EnableDisableControls False
    
    strOutputFile = Trim(txtDestination.Text)
    
    If mFso.GetExtensionName(strOutputFile) <> PEK_FILE_EXTENSION Then
        strOutputFile = mFso.BuildPath(mFso.GetParentFolderName(strOutputFile), mFso.GetBaseName(strOutputFile) & "." & PEK_FILE_EXTENSION)
    End If
     
    ' Validate that each file exists and determine the total bytes to be read
    ' Also, make sure none of the input files is the same as strOutputFile
    lblProgress.Caption = "Determining total size of the input files"
    lngTotalBytes = 0
    For i = 0 To mSourceFileListCount - 1
        strInputFilePath = Trim(mSourceFileList(i).FullPath)
        If Not mFso.FileExists(strInputFilePath) Then
            MsgBox "File not found: " & strInputFilePath, vbExclamation + vbOKOnly, "Error"
            EnableDisableControls True
            MergePEKFiles = False
            Exit Function
        ElseIf LCase(strOutputFile) = LCase(strInputFilePath) Then
            MsgBox "File " & strInputFilePath & " is defined as both an input and an output file; this is not allowed.", vbExclamation + vbOKOnly, "Error"
            EnableDisableControls True
            MergePEKFiles = False
            Exit Function
        End If
        Set objFile = mFso.GetFile(strInputFilePath)
        lngTotalBytes = lngTotalBytes + objFile.Size
    Next i
    If lngTotalBytes < 1 Then lngTotalBytes = 1
    
    ' Create the output file
    lblProgress.Caption = "Creating output file"
    Set tsOut = mFso.OpenTextFile(strOutputFile, ForWriting, True)
    
    ' Open each file in mSourceFileList() and copy each to the output file
    
    lngFileCountProcessed = 0
    lngLinesRead = 0
    lngBytesRead = 0
    
    lngScanNumberAddition = 0
    lngScanNumberCurrent = 0
    lngScanNumberPrevious = 0
    
    blnAddScanNumberAddition = False
    
    Set objScanNumberTracker = New clsScanNumberTracker
    
    lblProgress.Caption = "Merging files: 0% complete"
    
    For i = 0 To mSourceFileListCount - 1
        lngScansReadCurrentFile = 0
        strInputFilePath = mSourceFileList(i).FullPath
        Set tsIn = mFso.OpenTextFile(strInputFilePath, ForReading, True)
        
        Do While Not tsIn.AtEndOfStream
            strLineIn = tsIn.ReadLine
            
            If UCase(Left(strLineIn, intFilenameTextLength)) = FILENAME_AKA_SCAN_NUMBER Then
                ' This line contains the filename and scan number; parse it to extract the scan number
                
                ' First, store the current scan number in lngScanNumberPrevious
                lngScanNumberPrevious = lngScanNumberCurrent
                
                ' Extract the scan number from the filename line
                lngScanNumberCurrent = ExtractScanNumberFromFilenameLine(strLineIn, strTextBeforeScanNumber, strTextAfterScanNumber)
                lngScansReadCurrentFile = lngScansReadCurrentFile + 1
                
                If i = 0 Then
                    ' Reading the first file
                    ' Keep track of the delta scan values
                    objScanNumberTracker.AddScanNumberAndUpdateAverageIncrement lngScanNumberCurrent
                    If lngScansReadCurrentFile = 1 Then
                        lngScanNumberStartCurrentFile = lngScanNumberCurrent
                    End If
                Else
                    If lngScansReadCurrentFile = 1 Then
                        ' Determine whether or not we need to add lngScanNumberAddition to each scan
                        If lngScanNumberCurrent < lngScanNumberPrevious Then
                            blnAddScanNumberAddition = True
                            lngScanNumberStartCurrentFile = lngScanNumberCurrent + lngScanNumberAddition
                        Else
                            lngScanNumberStartCurrentFile = lngScanNumberCurrent
                        End If
                    End If
                    
                    If blnAddScanNumberAddition Then
                        lngScanNumberCurrent = lngScanNumberCurrent + lngScanNumberAddition
                        strLineIn = strTextBeforeScanNumber & Format(lngScanNumberCurrent, "0000") & strTextAfterScanNumber
                    End If
                End If
                
            End If
            
            tsOut.WriteLine strLineIn
            
            lngLinesRead = lngLinesRead + 1
            lngBytesRead = lngBytesRead + Len(strLineIn) + 2
            
            If lngLinesRead Mod 1000 = 0 Then
                If mAbortMerge Then Exit Do
                lblProgress.Caption = "Merging files: " & i + 1 & " of " & mSourceFileListCount & " -- " & Round(lngBytesRead / CDbl(lngTotalBytes) * 100, 0) & "% complete"
                DoEvents
            End If
        Loop
        tsIn.Close
        lngFileCountProcessed = lngFileCountProcessed + 1
        
        ' Bump up lngScanNumberAddition based on the maximum scan number in the current file
        lngScanNumberAddition = lngScanNumberAddition + Round(lngScanNumberCurrent - lngScanNumberStartCurrentFile + objScanNumberTracker.AverageScanIncrement, 0)
        If mAbortMerge Then Exit For
    Next i
    tsOut.Close
    
    If mAbortMerge Then
        lblProgress = "Merge aborted."
    Else
        lblProgress = "Merged " & Trim(lngFileCountProcessed) & " files into " & mFso.GetBaseName(strOutputFile) & " at " & mFso.GetParentFolderName(strOutputFile)
    End If
    
exit_MergePekFiles:
    EnableDisableControls True
    Set tsIn = Nothing
    Set tsOut = Nothing

    MergePEKFiles = Not mAbortMerge
    Exit Function
    
err_MergePEKFiles:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly
    mAbortMerge = True
    Resume exit_MergePekFiles
End Function

Private Sub MoveSourceItemDown()
    Dim udtCurrentItem As udtFileInfoType
    Dim udtSwap As udtFileInfoType
    
    Dim intSelectedIndex As Integer
    
    If mSourceFileListCount > 0 Then
        intSelectedIndex = lstSources.ListIndex
        If intSelectedIndex >= 0 And intSelectedIndex < mSourceFileListCount - 1 Then
            ' First move the item in mSourceFileList()
            udtCurrentItem = mSourceFileList(intSelectedIndex)
            udtSwap = mSourceFileList(intSelectedIndex + 1)
            
            mSourceFileList(intSelectedIndex + 1) = udtCurrentItem
            mSourceFileList(intSelectedIndex) = udtSwap
            
            ' Now move the item in the list and re-select it
            lstSources.List(intSelectedIndex + 1) = udtCurrentItem.CompactPath
            lstSources.List(intSelectedIndex) = udtSwap.CompactPath
            
            lstSources.Selected(intSelectedIndex + 1) = True
        End If
    End If
    
End Sub

Private Sub MoveSourceItemUp()
    Dim udtCurrentItem As udtFileInfoType
    Dim udtSwap As udtFileInfoType
    
    Dim intSelectedIndex As Integer
    
    If lstSources.ListCount > 0 Then
        intSelectedIndex = lstSources.ListIndex
        If intSelectedIndex > 0 Then
            ' First move the item in mSourceFileList()
            udtCurrentItem = mSourceFileList(intSelectedIndex)
            udtSwap = mSourceFileList(intSelectedIndex - 1)
            
            mSourceFileList(intSelectedIndex - 1) = udtCurrentItem
            mSourceFileList(intSelectedIndex) = udtSwap
            
            ' Now move the item in the list and re-select it
            lstSources.List(intSelectedIndex - 1) = udtCurrentItem.CompactPath
            lstSources.List(intSelectedIndex) = udtSwap.CompactPath
            
            lstSources.Selected(intSelectedIndex - 1) = True
        End If
    End If
End Sub

Private Sub PasteFileList()
    Dim strFilePath As String
    Dim strClipboard As String
    Dim intCharLoc As Integer
    
    If Clipboard.GetFormat(vbCFText) Then
        ClearSourceFileList
            
        strClipboard = Clipboard.GetText
        
        ' Split strClipboard on vbCrLf and populate mSourceFileListCount and lstSources
        Do
            intCharLoc = InStr(strClipboard, vbCrLf)
            If intCharLoc > 0 Then
                strFilePath = Trim(Left(strClipboard, intCharLoc - 1))
                strClipboard = Mid(strClipboard, intCharLoc + 2)
            Else
                strFilePath = Trim(strClipboard)
            End If
            
            If Len(strFilePath) > 0 Then
                AppendSourceFileWork strFilePath
            End If
        Loop While intCharLoc > 0
        
        AutoDefineOutputFileName
    Else
        MsgBox "Clipboard does not contain text data", vbInformation + vbOKOnly, "Invalid format"
    End If
    
End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    
    lngDesiredValue = Me.ScaleWidth - txtDestination.Left - cmdSelectOutputFile.width - 240
    If lngDesiredValue < 4000 Then
        lngDesiredValue = 4000
    End If
    txtDestination.width = lngDesiredValue
    
    cmdSelectOutputFile.Left = txtDestination.Left + txtDestination.width + 120
    cmdAutoDefineOutput.Left = cmdSelectOutputFile.Left
    cmdMerge.Left = cmdSelectOutputFile.Left
    cmdClose.Left = cmdSelectOutputFile.Left
    cmdCancel.Left = cmdSelectOutputFile.Left
    
    lngDesiredValue = cmdClose.Left - lblProgress.Left - 120
    If lngDesiredValue < 10 Then lngDesiredValue = 100
    lblProgress.width = lngDesiredValue
    
    lngDesiredValue = Me.ScaleWidth - lstSources.Left - cmdMoveUp.width - 240
    If lngDesiredValue < 3000 Then
        lngDesiredValue = 3000
    End If
    lstSources.width = lngDesiredValue
    
    lngDesiredValue = Me.ScaleHeight - lstSources.Top - 120
    If lngDesiredValue < 1000 Then lngDesiredValue = 1000
    lstSources.Height = lngDesiredValue
    
    cmdMoveUp.Left = lstSources.Left + lstSources.width + 120
    cmdMoveDown.Left = cmdMoveUp.Left
    cmdRemoveSelected.Left = cmdMoveUp.Left
    cmdClear.Left = cmdMoveUp.Left
      
End Sub

Private Sub RefreshSourceFileListbox(blnForceRefresh As Boolean)
    Dim intIndex As Integer
    Dim lngCompactPathLengthNew As Long
    
    lngCompactPathLengthNew = lstSources.width / 100
    If lngCompactPathLengthNew < 20 Then lngCompactPathLengthNew = 20
    
    If lngCompactPathLengthNew <> mCompactPathLength Or blnForceRefresh Then
        mCompactPathLength = lngCompactPathLengthNew
        
        For intIndex = 0 To mSourceFileListCount - 1
            With mSourceFileList(intIndex)
                .CompactPath = CompactPathString(.FullPath, mCompactPathLength)
                    
                If intIndex < lstSources.ListCount Then
                    lstSources.List(intIndex) = .CompactPath
                Else
                    ' lstSources is out of sync with mSourceFileList; this shouldn't happen
                    Debug.Assert False
                    lstSources.AddItem .CompactPath
                End If
            End With
        Next intIndex
    End If
End Sub

Private Sub RemoveSelectedItem()
    Dim intSelectedIndex  As Integer
    Dim intIndex As Integer
    
    intSelectedIndex = lstSources.ListIndex
    If intSelectedIndex >= 0 Then
        ' Remove the item from mSourceFileList
        For intIndex = intSelectedIndex To mSourceFileListCount - 2
            mSourceFileList(intIndex) = mSourceFileList(intIndex + 1)
        Next intIndex
        mSourceFileListCount = mSourceFileListCount - 1
        
        ' Remove the item from lstSources
        lstSources.RemoveItem (intSelectedIndex)
        If lstSources.ListCount > 0 Then
            If intSelectedIndex > lstSources.ListCount - 1 Then
                intSelectedIndex = lstSources.ListCount - 1
            End If
            lstSources.Selected(intSelectedIndex) = True
        End If
    End If
End Sub

Private Sub SelectOutputFile()
    Dim strFilePath As String
    Dim strExtension As String
    
    strExtension = GetPreferredFileExtension()
    
    strFilePath = SelectFile(hwnd, "Select output file", , True, "MergedData." & strExtension, GetFilterCodes(strExtension), 1)
    
    If Len(strFilePath) > 0 Then
        txtDestination.Text = strFilePath
    End If
End Sub

Private Sub SortFileList()
    Dim lngLowIndex As Long
    Dim lngHighIndex As Long
    
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim udtCompareVal As udtFileInfoType
    Dim strCompareFileName As String
    Dim blnItemsSwapped As Boolean
    
On Error GoTo ShellSortLongErrorHandler

    ' sort array[lngLowIndex..lngHighIndex]
    If mSourceFileListCount <= 1 Then Exit Sub
    
    ' compute largest increment
    lngLowIndex = 0
    lngHighIndex = mSourceFileListCount - 1
    lngIncrement = 1
    If (mSourceFileListCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < mSourceFileListCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            udtCompareVal = mSourceFileList(lngIndex)
            strCompareFileName = mFso.GetFileName(udtCompareVal.FullPath)
            
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If mFso.GetFileName(mSourceFileList(lngIndexCompare).FullPath) <= strCompareFileName Then Exit For
                mSourceFileList(lngIndexCompare + lngIncrement) = mSourceFileList(lngIndexCompare)
                blnItemsSwapped = True
            Next lngIndexCompare
            mSourceFileList(lngIndexCompare + lngIncrement) = udtCompareVal
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop
    
    If blnItemsSwapped Then
        RefreshSourceFileListbox True
    End If
    
    Exit Sub

ShellSortLongErrorHandler:
    Debug.Assert False
End Sub

Private Function StringEndsWith(ByVal strText As String, ByVal strComparisonText As String) As Boolean
  
    Dim blnMatchFound As Boolean
  
    blnMatchFound = False
    If Len(strText) >= Len(strComparisonText) Then
        If LCase(Right(strText, Len(strComparisonText))) = LCase(strComparisonText) Then
            blnMatchFound = True
        End If
    End If
    
    StringEndsWith = blnMatchFound
    
End Function

Private Function StringTrimEnd(ByVal strText As String, ByVal strTextToTrim As String) As String
    Dim intTrimLength As Integer
    Dim strTrimmedText As String
    
    intTrimLength = Len(strTextToTrim)
    
    strTrimmedText = strText
    If Len(strTrimmedText) >= intTrimLength Then
        If LCase(Right(strTrimmedText, intTrimLength)) = LCase(strTextToTrim) Then
            strTrimmedText = Left(strTrimmedText, Len(strTrimmedText) - intTrimLength)
        End If
    End If
    
    StringTrimEnd = strTrimmedText
End Function

Private Sub cmdAdd_Click()
    AppendSourceFile
End Sub

Private Sub cmdAutoDefineOutput_Click()
    AutoDefineOutputFileName
End Sub

Private Sub cmdCancel_Click()
    mAbortMerge = True
End Sub

Private Sub cmdClear_Click()
    Dim eResponse As VbMsgBoxResult
    
    If mSourceFileListCount > 0 Then
        eResponse = MsgBox("Are you sure you want to clear the source file list?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Clear list")
        If eResponse = vbYes Then
            ClearSourceFileList
            cmdMerge.Enabled = False
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMerge_Click()
    MergeFilesStart
End Sub

Private Sub cmdMoveDown_Click()
    MoveSourceItemDown
End Sub

Private Sub cmdMoveUp_Click()
    MoveSourceItemUp
End Sub

Private Sub cmdPasteFileList_Click()
    PasteFileList
End Sub

Private Sub cmdRemoveSelected_Click()
    RemoveSelectedItem
End Sub

Private Sub cmdSelectOutputFile_Click()
    SelectOutputFile
End Sub

Private Sub cmdSortList_Click()
    SortFileList
End Sub

Private Sub Form_Load()
    InitializeForm
End Sub

Private Sub Form_Resize()
    PositionControls
    mFormResized = True
End Sub

Private Sub optMOrder_Click(Index As Integer)
    mMergeOrder = Index
End Sub

Private Sub tmrTimer_Timer()
    Static mLastUpdateTickCount As Long
    
    If mFormResized Then
        mFormResized = False
        mLastUpdateTickCount = GetTickCount()
        RefreshSourceFileListbox False
    End If
End Sub

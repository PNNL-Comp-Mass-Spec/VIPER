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

Private Type udtFileInfoType
    FullPath As String
    CompactPath As String
End Type

Private mMergeOrder As Long

Private mCompactPathLength As Integer

Private mSourceFileListCount As Integer
Private mSourceFileList() As udtFileInfoType

Private mAbortMerge As Boolean
Private mFormResized As Boolean

Private Sub AppendSourceFile()
    Dim strFilePath As String
    
    strFilePath = SelectFile(hwnd, "Select source file", , False, "*.pek", "All Files (*.*)|*.*|PEK Files (*.pek)|*.pek", 1)
    AppendSourceFileWork strFilePath
End Sub

Private Sub AppendSourceFileWork(strFilePath As String)
    Dim blnDuplicateFile As Boolean
    Dim intIndex As Integer
    
    If Len(strFilePath) > 0 Then
        ' Only add file to mSourceFileList if not already present
        For intIndex = 0 To mSourceFileListCount - 1
            If LCase(strFilePath) = LCase(mSourceFileList(intIndex).FullPath) Then
                blnDuplicateFile = True
                Exit For
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
        End If
    End If

End Sub

Private Sub AutoDefineOutputFileName()
    ' Look for the common text between the first and second files and define the output file name based on that
    
    Dim strFolderPath As String
    Dim strOutputFilePath As String
    Dim strComparisonFilePath As String
    Dim intIndex As Integer
    
    Dim fso As FileSystemObject
    
    If mSourceFileListCount > 0 Then
        Set fso = New FileSystemObject
        strFolderPath = fso.GetParentFolderName(mSourceFileList(0).FullPath)
        
        If mSourceFileListCount <= 1 Then
            strOutputFilePath = fso.GetBaseName(mSourceFileList(0).FullPath) & "_Merged.pek"
        Else
            strOutputFilePath = fso.GetBaseName(mSourceFileList(0).FullPath)
            strComparisonFilePath = fso.GetBaseName(mSourceFileList(1).FullPath)
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
        
        strOutputFilePath = fso.BuildPath(strFolderPath, strOutputFilePath & "_Merged.pek")
        
        txtDestination = strOutputFilePath
    End If
End Sub

Private Sub ClearSourceFileList()
    mSourceFileListCount = 0
    lstSources.Clear
End Sub

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

Private Sub InitializeForm()
    mMergeOrder = 0
    
    optMOrder(0).Value = True
    lblProgress.Caption = ""

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

Private Function MergeFiles() As Boolean
    '-----------------------------------------------------
    ' Reads files in order and merge them together
    '-----------------------------------------------------
    
    Const FILENAME_AKA_SCAN_NUMBER As String = "FILENAME"
    Dim intFilenameTextLength As Integer
    intFilenameTextLength = Len(FILENAME_AKA_SCAN_NUMBER)
    
    Dim fso As FileSystemObject
    
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
    
On Error GoTo err_MergeFiles
    
    If mSourceFileListCount < 2 Then
        MsgBox "Please define at least two input files.", vbExclamation Or vbOKOnly, "Error"
        MergeFiles = False
        Exit Function
    End If
    
    strOutputFile = Trim(txtDestination.Text)
    If Len(strOutputFile) = 0 Then
        MsgBox "Please define the destination file", vbExclamation Or vbOKOnly, "Error"
        MergeFiles = False
        Exit Function
    End If
   
    mAbortMerge = False
    EnableDisableControls False
    
    Set fso = New FileSystemObject
    
    ' Validate that each file exists and determine the total bytes to be read
    ' Also, make sure none of the input files is the same as strOutputFile
    lblProgress.Caption = "Determining total size of the input files"
    lngTotalBytes = 0
    For i = 0 To mSourceFileListCount - 1
        strInputFilePath = Trim(mSourceFileList(i).FullPath)
        If Not fso.FileExists(strInputFilePath) Then
            MsgBox "File not found: " & strInputFilePath, vbExclamation + vbOKOnly, "Error"
            EnableDisableControls True
            MergeFiles = False
            Exit Function
        ElseIf LCase(strOutputFile) = LCase(strInputFilePath) Then
            MsgBox "File " & strInputFilePath & " is defined as both an input and an output file; this is not allowed.", vbExclamation + vbOKOnly, "Error"
            EnableDisableControls True
            MergeFiles = False
            Exit Function
        End If
        Set objFile = fso.GetFile(strInputFilePath)
        lngTotalBytes = lngTotalBytes + objFile.Size
    Next i
    If lngTotalBytes < 1 Then lngTotalBytes = 1
    
    ' Create the output file
    lblProgress.Caption = "Creating output file"
    Set tsOut = fso.OpenTextFile(strOutputFile, ForWriting, True)
    
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
        Set tsIn = fso.OpenTextFile(strInputFilePath, ForReading, True)
        
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
        lblProgress = "Merged " & Trim(lngFileCountProcessed) & " files into " & fso.GetBaseName(strOutputFile) & " at " & fso.GetParentFolderName(strOutputFile)
    End If
    
exit_MergeFiles:
    EnableDisableControls True
    Set tsIn = Nothing
    Set tsOut = Nothing

    MergeFiles = Not mAbortMerge
    Exit Function
    
err_MergeFiles:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly
    mAbortMerge = True
    Resume exit_MergeFiles
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
    strFilePath = SelectFile(hwnd, "Select output file", , True, "MergedData.pek", "All Files (*.*)|*.*|PEK Files (*.pek)|*.pek", 1)
    
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
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
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
            strCompareFileName = fso.GetFileName(udtCompareVal.FullPath)
            
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If fso.GetFileName(mSourceFileList(lngIndexCompare).FullPath) <= strCompareFileName Then Exit For
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
    Select Case mMergeOrder
    Case MERGE_ORDER_SEQUENTIAL
        Call MergeFiles
    Case MERGE_ORDER_INTERLEAVE
        MsgBox "Interleaved scan mode is not yet implemented.", vbExclamation + vbOKOnly, glFGTU
    Case Else
        MsgBox "Unknown merge order: " & mMergeOrder, vbExclamation + vbOKOnly, glFGTU
    End Select
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

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPEKSplit 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PEK File Split Function"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmPEKSplit.frx":0000
   LinkTopic       =   "frmMerge"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSplitType 
      BackColor       =   &H80000001&
      Caption         =   "Split Type"
      ForeColor       =   &H80000009&
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   4335
      Begin VB.TextBox txtOptST 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   12
         Text            =   "1000"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtOptST 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   11
         Text            =   "1000"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtOptST 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   10
         Text            =   "2"
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optST 
         BackColor       =   &H80000001&
         Caption         =   "Start new file after "
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1020
         Width           =   1695
      End
      Begin VB.OptionButton optST 
         BackColor       =   &H80000001&
         Caption         =   "Start new file after every scan multiple of"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   660
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton optST 
         BackColor       =   &H80000001&
         Caption         =   "&Alternate modulo"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "scans."
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   1020
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   90
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CmnDlg1 
      Left            =   2880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "PEK"
      Filter          =   "PEK Files (*.pek)|*.pek|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "&Split"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtInFile 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File to split:"
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPEKSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Splits PEK file based on specified criteria.
'Splits are written in the same folder as original if that fails (rw
'permission) then ask user for location to write
'NOTE: File numbers used to split PEK are ordinals not actual scan
'      numbers found in the file
'NOTE: Assumption is that scan blocks are separated with empty line
'-------------------------------------------------------------------
'last modified 05/31/2002 nt
'-------------------------------------------------------------------

Option Explicit

Const DEF_ALT = 2
Const DEF_SC1 = 1000
Const DEF_SC2 = 1000

Const SPLIT_ALT = 0
Const SPLIT_SC1 = 1
Const SPLIT_SC2 = 2

Dim SplitType As Long

Dim AltVal As Long
Dim Sc1Val As Long
Dim Sc2Val As Long

Dim InFName As String           'input file name
Dim BInFName As String          'base input file name
Dim OutFBaseName As String      'output file base name

Dim fs As New FileSystemObject

Private Sub cmdBrowse_Click()
On Error Resume Next
CmnDlg1.CancelError = True
CmnDlg1.ShowOpen
If Err Then Exit Sub
If Len(CmnDlg1.FileName) > 0 Then
   InFName = CmnDlg1.FileName
   txtInFile.Text = InFName
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSplit_Click()
If Not fs.FileExists(InFName) Then
   MsgBox "Input file not found. Mission impossible.", vbOKOnly, glFGTU
   Exit Sub
End If
Select Case SplitType
Case SPLIT_ALT
    Call SplitAlternate
Case SPLIT_SC1
    Call SplitSC1
Case SPLIT_SC2
    'Call SplitSC2
    MsgBox "This option is not yet implemented.", vbOKOnly, glFGTU
End Select
End Sub

Private Sub Form_Load()
AltVal = DEF_ALT
txtOptST(0).Text = AltVal
Sc1Val = DEF_SC1
txtOptST(1).Text = Sc1Val
Sc2Val = DEF_SC2
txtOptST(2).Text = Sc2Val
SplitType = SPLIT_SC1
End Sub

Private Sub optST_Click(Index As Integer)
SplitType = Index
End Sub

Private Sub txtInFile_LostFocus()
InFName = Trim$(txtInFile.Text)
End Sub

Private Sub UpdateStatus(ByVal StatusMsg As String)
lblStatus.Caption = StatusMsg
DoEvents
End Sub

Private Sub txtOptST_LostFocus(Index As Integer)
On Error GoTo err_txtOptST
Select Case Index
Case SPLIT_ALT
   AltVal = CLng(txtOptST(Index).Text)
   If AltVal < 2 Or AltVal > 32 Then
      MsgBox "Alternating split is possible for 2 to 32 files.", vbOKOnly, glFGTU
      txtOptST(Index).SetFocus
   End If
Case SPLIT_SC1
   Sc1Val = CLng(txtOptST(Index).Text)
Case SPLIT_SC2
   Sc1Val = CLng(txtOptST(Index).Text)
End Select
Exit Sub

err_txtOptST:
MsgBox "This argument has to be positive integer.", vbOKOnly, glFGTU
txtOptST(Index).SetFocus
End Sub

Private Sub SplitAlternate()
'-------------------------------------------------
'splits PEK by alternating between k output files
'-------------------------------------------------
Dim Ln As String
Dim tsIn As TextStream
Dim tsOut() As TextStream
Dim Res As Long
Dim FCnt As Long
Dim CurrOutStreamInd As Long
Dim OverwriteIfExists As Boolean
Dim i As Long
On Error GoTo err_SplitAlternate

'first try to establish output streams to write with the split
OverwriteIfExists = False
If Not SetBaseFName(False) Then Exit Sub    'try first to write at the same folder
Set tsIn = fs.OpenTextFile(InFName, ForReading, False)
ReDim tsOut(AltVal - 1)
For i = 0 To AltVal - 1
    Set tsOut(i) = fs.CreateTextFile(OutFBaseName & i & ".spl", OverwriteIfExists)
Next i
'then go and write blocks of data to open streams
Do Until tsIn.AtEndOfStream
   Ln = Trim$(tsIn.ReadLine)
   If Len(Ln) > 0 Then      'just rewrite line in currently open data stream
      If FCnt = 0 Then              'first line; direct to first output stream
         CurrOutStreamInd = 0       'first output file
         FCnt = 1                   'first scan data block
         UpdateStatus "Writing file " & BInFName & "_" & CurrOutStreamInd & ".spl"
      End If
      tsOut(CurrOutStreamInd).WriteLine Ln
   Else
      FCnt = FCnt + 1
      CurrOutStreamInd = FCnt Mod AltVal
      UpdateStatus "Writing file " & BInFName & "_" & CurrOutStreamInd & ".spl"
      tsOut(CurrOutStreamInd).WriteLine Ln
   End If
Loop
tsIn.Close
For i = 0 To AltVal - 1
    tsOut(i).Close
Next i

exit_SplitAlternate:        'exit clean
UpdateStatus ""
Set tsIn = Nothing
For i = 0 To AltVal - 1
    Set tsOut(i) = Nothing
Next i
Erase tsOut
Exit Sub

err_SplitAlternate:
UpdateStatus "Hmmm, lets see ..."
Select Case Err.Number
Case 58         'split file already exists; ask for permission to overwrite
    Res = MsgBox("Split file found on specified target destination. Overwrite existing files?", vbYesNoCancel, glFGTU)
    If Res = vbYes Then
       OverwriteIfExists = True
       Resume
    End If
Case 70         'maybe permission is denied for write operation
                'NOTE: this error will be(if) triggered from inside the loop
    Res = MsgBox("Access permission denied for selected destination. Do you want split to be written in different place?", vbYesNoCancel, glFGTU)
    If Res = vbYes Then
       If SetBaseFName(True) Then Resume
    End If
Case Else
    LogErrors Err.Number, "frmPEKSplit.SplitAlternate"
End Select
GoTo exit_SplitAlternate
End Sub


Public Sub SplitSC1()
'------------------------------------------------------------------
'read input file line by line and write it to the output file
'when line with file number multiple of Sc1Val comes close current
'out file and start new (File number is not actual Scan Number)
'------------------------------------------------------------------
Dim OutFCnt As Long
Dim OutFName As String
Dim Ln As String
Dim tsIn As TextStream
Dim tsOut As TextStream
Dim Res As Long
Dim FCnt As Long
Dim OverwriteIfExists As Boolean
On Error GoTo err_SplitSC1

OverwriteIfExists = False
If Not SetBaseFName(False) Then Exit Sub    'try first to write at the same folder
Set tsIn = fs.OpenTextFile(InFName, ForReading, False)

Do Until tsIn.AtEndOfStream
   Ln = Trim$(tsIn.ReadLine)
   If Len(Ln) > 0 Then      'just rewrite line in currently open data stream
      If FCnt = 0 Then      'first line; open first output stream
         OutFCnt = 1        'first output file
         FCnt = 1           'first scan data block
         OutFName = OutFBaseName & OutFCnt & ".spl"
         UpdateStatus "Writing file " & BInFName & "_" & OutFCnt & ".spl"
         Set tsOut = fs.CreateTextFile(OutFName, OverwriteIfExists)
         tsOut.WriteLine Ln
      Else
         tsOut.WriteLine Ln
      End If
   Else
      FCnt = FCnt + 1
      If (FCnt Mod Sc1Val = 0) Then 'close stream; open new stream; write line and continue
          tsOut.Close
          DoEvents
          OutFCnt = OutFCnt + 1
          OutFName = OutFBaseName & OutFCnt & ".spl"
          UpdateStatus "Writing file " & BInFName & "_" & OutFCnt & ".spl"
          Set tsOut = fs.CreateTextFile(OutFName, OverwriteIfExists)
          tsOut.WriteLine Ln
      Else    'just continue
          tsOut.WriteLine Ln
      End If
   End If
Loop
tsIn.Close
tsOut.Close

exit_SplitSC1:
UpdateStatus ""
Set tsIn = Nothing
Set tsOut = Nothing
Exit Sub

err_SplitSC1:
UpdateStatus "Hmmm, lets see...."
Select Case Err.Number
Case 58         'split file already exists; ask for permission to overwrite
    Res = MsgBox("Split file found on specified target destination. Overwrite existing files?", vbYesNoCancel, glFGTU)
    If Res = vbYes Then
       OverwriteIfExists = True
       Resume
    End If
Case 70         'maybe permission is denied for write operation
    Res = MsgBox("Access permission denied for selected destination. Do you want split to be written in different place?", vbYesNoCancel, glFGTU)
    If Res = vbYes Then
       If SetBaseFName(True) Then                       'if user selects new place
          OutFName = OutFBaseName & OutFCnt & ".spl"    'assemble new file name
          UpdateStatus "Writing file " & BInFName & "_" & OutFCnt & ".spl"
          Resume                                        'and continue
       End If
    End If
Case Else
    LogErrors Err.Number, "frmPEKSplit.SplitSC1"
End Select
GoTo exit_SplitSC1
End Sub

Private Function SetBaseFName(ByVal AskForDirPath As Boolean) As Boolean
'-----------------------------------------------------------------------
'sets base file name and returns True if OK; If AskForDirPath is set
'user is asked to browse for folder; otherwise it uses folder of InFName
'-----------------------------------------------------------------------
Dim DirPath As String       'directory path
Dim bi As BROWSEINFO
Dim pidl As Long
Dim pos As Long
On Error GoTo exit_SetBaseFName

With fs
    BInFName = .GetBaseName(InFName)
    If AskForDirPath Then       'ask user to browse for folder
       bi.hwndOwner = Me.hwnd
       bi.pidlRoot = 0&
       bi.lpszTitle = "Browse to the Data Folder"
       bi.ulFlags = BIF_RETURNONLYFSDIRS
       pidl = SHBrowseForFolder(bi)
       DirPath = Space$(MAX_PATH)
       If SHGetPathFromIDList(ByVal pidl, ByVal DirPath) Then
          pos = InStr(DirPath, Chr$(0))
          DirPath = Left$(DirPath, pos - 1)
       End If
       Call CoTaskMemFree(pidl)
       If Len(DirPath) <= 0 Then GoTo exit_SetBaseFName
    Else
       DirPath = .GetParentFolderName(InFName)
    End If
    OutFBaseName = .BuildPath(DirPath, BInFName & "_")
End With
SetBaseFName = True

exit_SetBaseFName:
End Function


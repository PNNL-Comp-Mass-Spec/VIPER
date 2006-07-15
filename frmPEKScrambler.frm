VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPEKScrambler 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PEK Scrambler"
   ClientHeight    =   3045
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Do It"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame fraMass 
      BackColor       =   &H80000001&
      Caption         =   "Molecular Mass"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3855
      Begin VB.OptionButton optMWChange 
         BackColor       =   &H80000001&
         Caption         =   "Random change (ppm)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1500
         Width           =   2895
      End
      Begin VB.TextBox txtMWChangeVal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Text            =   "1"
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton optMWChange 
         BackColor       =   &H80000001&
         Caption         =   "Random change (Da)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   2895
      End
      Begin VB.OptionButton optMWChange 
         BackColor       =   &H80000001&
         Caption         =   "Add to each molecular mass (ppm)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   900
         Width           =   2895
      End
      Begin VB.OptionButton optMWChange 
         BackColor       =   &H80000001&
         Caption         =   "Add to each molecular mass (Da)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optMWChange 
         BackColor       =   &H80000001&
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   180
      Width           =   5655
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   540
      Width           =   5655
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CmnDlg1 
      Left            =   7080
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Target:"
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "frmPEKScrambler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'created: 02/14/2003 nt
'last modified: 02/14/2003 nt
'------------------------------------------------------------------
Dim Scrambling As Boolean

Const DATA_DELI_ASC = 9

Const CHG_NONE = 0

Const CHG_MW_ADD_Da = 1
Const CHG_MW_ADD_ppm = 2
Const CHG_MW_RND_Da = 3
Const CHG_MW_RND_ppm = 4

Const PREV_LN_TYPE_OTHER_NOT_DATA = 0
Const PREV_LN_TYPE_CS_HEADER = 1
Const PREV_LN_TYPE_IS_HEADER = 2

Dim fso As New FileSystemObject

Dim MWChangeType As Long
Dim MWChangeVal As Double

Option Explicit

Private Sub cmdBrowse_Click()
On Error Resume Next
If Not Scrambling Then
   CmnDlg1.DialogTitle = "Browse to PEK File"
   CmnDlg1.Filter = "| files (*.pek*)|*.pek*|All files (*.*)|*.*"
   CmnDlg1.ShowOpen
   If Err Then Exit Sub
   If Len(CmnDlg1.FileName) > 0 Then
      txtSource.Text = CmnDlg1.FileName
      txtTarget.Text = txtSource.Text & "S"
   End If
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDoIt_Click()         'THIS PROCEDURE IS I THINK OK
Dim SourceFile As String
Dim TargetFile As String
Dim tsIn As TextStream
Dim tsOut As TextStream
Dim sLnIn As String
Dim sLnOut As String
Dim LnType As Long
Dim sLnSplit() As String
Dim Res As Long
On Error Resume Next

'sum here all criteria to see if any needs to be applied
If MWChangeType <= 0 Then
    MsgBox "What should be changed; nothing is selected.", vbOKOnly, glFGTU
    Exit Sub
Else
    If IsNumeric(txtMWChangeVal.Text) Then
       MWChangeVal = CDbl(txtMWChangeVal.Text)
    Else
       MsgBox "MW change value should be numeric.", vbOKOnly, glFGTU
       txtMWChangeVal.SetFocus
       Exit Sub
    End If
End If
SourceFile = Trim$(txtSource.Text)
If Len(SourceFile) <= 0 Then
    MsgBox "Source file not specified.", vbOKOnly
    txtSource.SetFocus
    Exit Sub
End If
TargetFile = Trim$(txtTarget.Text)
If Len(TargetFile) <= 0 Then
    MsgBox "Target file not specified.", vbOKOnly
    txtTarget.SetFocus
    Exit Sub
End If
Me.MousePointer = vbHourglass
Scrambling = True
Set tsIn = fso.OpenTextFile(SourceFile, ForReading)
Set tsOut = fso.CreateTextFile(TargetFile, True)

Do Until tsIn.AtEndOfStream
   sLnIn = tsIn.ReadLine
   Select Case Left$(sLnIn, 8)
   Case t8DATA_CS
        LnType = PREV_LN_TYPE_CS_HEADER
        sLnOut = sLnIn
   Case t8DATA_ISO
        LnType = PREV_LN_TYPE_IS_HEADER
        sLnOut = sLnIn
   Case Else
        sLnSplit = Split(sLnIn, Chr(DATA_DELI_ASC))
        Select Case LnType
        Case PREV_LN_TYPE_CS_HEADER
             If UBound(sLnSplit) > 5 Then        'enough for data line
                sLnOut = ScrambleLineOfDataCS(sLnSplit)
                If Len(sLnOut) <= 0 Then
                   sLnOut = sLnIn
                   LnType = PREV_LN_TYPE_OTHER_NOT_DATA
                End If
             Else
                sLnOut = sLnIn
                LnType = PREV_LN_TYPE_OTHER_NOT_DATA
             End If
        Case PREV_LN_TYPE_IS_HEADER
             If UBound(sLnSplit) > 5 Then        'enough for data line
                sLnOut = ScrambleLineOfDataIso(sLnSplit)
                If Len(sLnOut) <= 0 Then
                   sLnOut = sLnIn
                   LnType = PREV_LN_TYPE_OTHER_NOT_DATA
                End If
             Else
                sLnOut = sLnIn
                LnType = PREV_LN_TYPE_OTHER_NOT_DATA
             End If
        Case Else
             sLnOut = sLnIn
             LnType = PREV_LN_TYPE_OTHER_NOT_DATA
        End Select
   End Select
   tsOut.WriteLine sLnOut
Loop

tsOut.Close
tsIn.Close
Scrambling = False
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Randomize
MWChangeType = 1
End Sub

Private Sub optMWChange_Click(Index As Integer)
MWChangeType = Index
End Sub

Private Function ScrambleLineOfDataCS(lSplitLine() As String) As String
Dim i As Long
Dim tmp As String
Dim ThisVal As Double
On Error GoTo err_ScrambleLineOfDataCS
ThisVal = CDbl(lSplitLine(3))
Select Case MWChangeType        'mass column is 4
Case CHG_MW_ADD_Da
     lSplitLine(3) = Format$(ThisVal + MWChangeVal, "0.0000")
Case CHG_MW_ADD_ppm
     lSplitLine(3) = Format$(ThisVal * (1 + MWChangeVal * glPPM), "0.0000")
Case CHG_MW_RND_Da
     lSplitLine(3) = Format$(ThisVal + Rnd() * MWChangeVal, "0.0000")
Case CHG_MW_RND_ppm
     lSplitLine(3) = Format$(ThisVal * (1 + Rnd() * MWChangeVal * glPPM), "0.0000")
End Select
For i = 0 To UBound(lSplitLine)
    tmp = tmp & lSplitLine(i) & vbTab
Next i
ScrambleLineOfDataCS = Left$(tmp, Len(tmp) - 1)
Exit Function

err_ScrambleLineOfDataCS:
ScrambleLineOfDataCS = ""
End Function

Private Function ScrambleLineOfDataIso(lSplitLine() As String) As String
Dim i As Long
Dim tmp As String
Dim ThisVal As Double
On Error GoTo err_ScrambleLineOfDataIso
Select Case MWChangeType        'mass columns are 5,6,7
Case CHG_MW_ADD_Da
     ThisVal = CDbl(lSplitLine(4))
     lSplitLine(4) = Format$(ThisVal + MWChangeVal, "0.0000")
     ThisVal = CDbl(lSplitLine(5))
     lSplitLine(5) = Format$(ThisVal + MWChangeVal, "0.0000")
     ThisVal = CDbl(lSplitLine(6))
     lSplitLine(6) = Format$(ThisVal + MWChangeVal, "0.0000")
Case CHG_MW_ADD_ppm
     ThisVal = CDbl(lSplitLine(4))
     lSplitLine(4) = Format$(ThisVal * (1 + MWChangeVal * glPPM), "0.0000")
     ThisVal = CDbl(lSplitLine(5))
     lSplitLine(5) = Format$(ThisVal * (1 + MWChangeVal * glPPM), "0.0000")
     ThisVal = CDbl(lSplitLine(6))
     lSplitLine(6) = Format$(ThisVal * (1 + MWChangeVal * glPPM), "0.0000")
Case CHG_MW_RND_Da
     ThisVal = CDbl(lSplitLine(4))
     lSplitLine(4) = Format$(ThisVal + Rnd() * MWChangeVal, "0.0000")
     ThisVal = CDbl(lSplitLine(5))
     lSplitLine(5) = Format$(ThisVal + Rnd() * MWChangeVal, "0.0000")
     ThisVal = CDbl(lSplitLine(6))
     lSplitLine(6) = Format$(ThisVal + Rnd() * MWChangeVal, "0.0000")
Case CHG_MW_RND_ppm
     ThisVal = CDbl(lSplitLine(4))
     lSplitLine(4) = Format$(ThisVal * (1 + Rnd() * MWChangeVal * glPPM), "0.0000")
     ThisVal = CDbl(lSplitLine(5))
     lSplitLine(5) = Format$(ThisVal * (1 + Rnd() * MWChangeVal * glPPM), "0.0000")
     ThisVal = CDbl(lSplitLine(6))
     lSplitLine(6) = Format$(ThisVal * (1 + Rnd() * MWChangeVal * glPPM), "0.0000")
End Select
For i = 0 To UBound(lSplitLine)
    tmp = tmp & lSplitLine(i) & vbTab
Next i
ScrambleLineOfDataIso = Left$(tmp, Len(tmp) - 1)
Exit Function

err_ScrambleLineOfDataIso:
ScrambleLineOfDataIso = ""
End Function


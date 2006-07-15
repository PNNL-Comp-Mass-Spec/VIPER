VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProtoViewFASTA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Protein Based Analysis - FASTA"
   ClientHeight    =   7170
   ClientLeft      =   2760
   ClientTop       =   4035
   ClientWidth     =   7350
   Icon            =   "frmProtoViewFASTA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5640
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstPeptides 
      Height          =   4545
      Left            =   4920
      TabIndex        =   10
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CheckBox chkSeqOnly 
      Caption         =   "Show seq. only"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "If checked only sequence of protein will be shown"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox txtORFNavPos 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Text            =   "0"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   ">>"
      Height          =   315
      Index           =   3
      Left            =   2860
      TabIndex        =   7
      ToolTipText     =   "Last"
      Top             =   6480
      Width           =   425
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   ">"
      Height          =   315
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Next"
      Top             =   6480
      Width           =   425
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   "<"
      Height          =   315
      Index           =   1
      Left            =   580
      TabIndex        =   5
      ToolTipText     =   "Last"
      Top             =   6480
      Width           =   425
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   "<<"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "First"
      Top             =   6480
      Width           =   425
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cleavage Rules"
      Height          =   1200
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtDigRules 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmProtoViewFASTA.frx":030A
         Top             =   240
         Width           =   2055
      End
   End
   Begin RichTextLib.RichTextBox rtbORFSeq 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10821
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmProtoViewFASTA.frx":032F
   End
   Begin VB.Label Label1 
      Caption         =   "Peptides"
      Height          =   225
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   4695
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Save Digestion"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmProtoViewFASTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE: proteins are loaded from FASTA file
'created: 08/12/2002 nt
'last modified: 08/12/2002 nt
'---------------------------------------------------------------
Option Explicit

Const NAV_FIRST = 0
Const NAV_PREVIOUS = 1
Const NAV_NEXT = 2
Const NAV_LAST = 3
Const NAV_STAY = 21

Dim CallerID As Long

Dim CurrORFInd As Long          'index in RefID array

Dim PeptCnt As Long             'count of peptides for current protein
Dim PeptID() As Long
Dim PeptMW() As Double

Dim FAReader As FastaReader

Private Sub chkSeqOnly_Click()
Call cmdORFNavigate_Click(NAV_STAY)
End Sub

Private Sub cmdORFNavigate_Click(Index As Integer)
If FAReader Is Nothing Then Exit Sub
Select Case Index
Case NAV_FIRST
     CurrORFInd = 1
Case NAV_PREVIOUS
     If CurrORFInd > 1 Then CurrORFInd = CurrORFInd - 1
Case NAV_NEXT
     If CurrORFInd < FAReader.RecordCount Then CurrORFInd = CurrORFInd + 1
Case NAV_LAST
     CurrORFInd = FAReader.RecordCount
Case Else                       'don't go anywhere
End Select
txtORFNavPos.Text = " ORF: " & CurrORFInd & "/" & FAReader.RecordCount
If chkSeqOnly.Value = vbChecked Then
   rtbORFSeq.Text = FAReader.GetFASTARecordSeq(CurrORFInd)
Else
   rtbORFSeq.Text = FAReader.GetFASTARecordDesc(CurrORFInd) & vbCrLf & FAReader.GetFASTARecordSeq(CurrORFInd)
End If
PeptCnt = 0
ReDim PeptID(1 To 1000)      'will be plenty for most proteins
ReDim PeptMW(1 To 1000)
lstPeptides.Clear
End Sub

Private Sub Form_Activate()
'------------------------------------------------------------
'load MT tags and ORF databases date if neccessary
'------------------------------------------------------------
On Error Resume Next
CallerID = Me.Tag
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub


Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFOpen_Click()
'----------------------------------------------------------------------
'enumerate FASTA file records (instead of usual file loading)
'----------------------------------------------------------------------
Dim Resp As Long
On Error Resume Next
If Not FAReader Is Nothing Then
   If FAReader.RecordCount > 0 Then
      Resp = MsgBox("Records from file: " & FAReader.FASTAFile & " already loaded. Continue?", vbYesNo, glFGTU)
      If Resp <> vbYes Then Exit Sub
   End If
End If
If Err Then Err.Clear
cd1.CancelError = True
cd1.DialogTitle = "Select FASTA File to Load ..."
cd1.Filter = "FASTA files (*.FASTA)|*.FASTA|Text files (*.txt)|*.txt|All files (*.*)|*.*"
cd1.ShowOpen
If Err Then Exit Sub
If Len(Trim$(cd1.FileName)) > 0 Then
   Me.MousePointer = vbHourglass
   UpdateStatus "Enumerating ORF records ..."
   Set FAReader = New FastaReader
   FAReader.FASTAFile = cd1.FileName
   If FAReader.EnumerateFASTARecords() Then Call cmdORFNavigate_Click(NAV_FIRST)
   UpdateStatus "ORF count: " & FAReader.RecordCount
   Me.MousePointer = vbDefault
End If
End Sub

VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProtoView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Protein Based Analysis"
   ClientHeight    =   7170
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstPeptides 
      Height          =   3375
      Left            =   4920
      TabIndex        =   15
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CheckBox chkSeqOnly 
      Caption         =   "Show seq. only"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      ToolTipText     =   "If checked only sequence of protein will be shown"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox txtORFNavPos 
      Height          =   315
      Left            =   960
      TabIndex        =   13
      Text            =   "0"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   ">>"
      Height          =   315
      Index           =   3
      Left            =   2860
      TabIndex        =   12
      ToolTipText     =   "Last"
      Top             =   6480
      Width           =   425
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   ">"
      Height          =   315
      Index           =   2
      Left            =   2400
      TabIndex        =   11
      ToolTipText     =   "Next"
      Top             =   6480
      Width           =   425
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   "<"
      Height          =   315
      Index           =   1
      Left            =   580
      TabIndex        =   10
      ToolTipText     =   "Last"
      Top             =   6480
      Width           =   425
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   "<<"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "First"
      Top             =   6480
      Width           =   425
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cleavage Rules"
      Height          =   1200
      Left            =   4920
      TabIndex        =   6
      Top             =   1380
      Width           =   2295
      Begin VB.TextBox txtDigRules 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmProtoView.frx":0000
         Top             =   240
         Width           =   2055
      End
   End
   Begin RichTextLib.RichTextBox rtbORFSeq 
      Height          =   6135
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10821
      _Version        =   393217
      TextRTF         =   $"frmProtoView.frx":0025
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mass Tags Selection"
      Height          =   1200
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optInclude 
         Caption         =   "Local digestion"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton optInclude 
         Caption         =   "Loaded mass tags only"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optInclude 
         Caption         =   "All mass tags for ORF"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Peptides"
      Height          =   225
      Left            =   4920
      TabIndex        =   16
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   4695
   End
End
Attribute VB_Name = "frmProtoView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE: proteins are loaded from database; it would be relatively
'simple to extend this functionality to FASTA file  but hey ....
'created: 08/01/2002 nt
'last modified: 08/05/2002 nt
'---------------------------------------------------------------
Option Explicit

Const CLR_HIGHLIGHT = vbRed
Const CLR_SEQUENCE = vbBlue

Const NAV_FIRST = 0
Const NAV_PREVIOUS = 1
Const NAV_NEXT = 2
Const NAV_LAST = 3
Const NAV_STAY = 21

Const OPT_PEPT_DB_ALL = 0
Const OPT_PEPT_DB_LOADED = 1
Const OPT_PEPT_DIGESTION = 2

Dim CallerID As Long
Dim bLoading As Boolean

Dim CurrORFInd As Long          'index in RefID array
Dim CurrORFSeq As String        'current ORF(Ref) sequence

Dim PeptOption As Long          'option that determines which peptides are loaded

Dim PeptCnt As Long             'count of peptides for current protein
Dim PeptID() As Long
Dim PeptMW() As Double

Dim AMTIDCol As Collection      'used for search of loaded mass tags based on ID


Private Sub chkSeqOnly_Click()
Call cmdORFNavigate_Click(NAV_STAY)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdORFNavigate_Click(Index As Integer)
Select Case Index
Case NAV_FIRST
     CurrORFInd = 1
Case NAV_PREVIOUS
     If CurrORFInd > 1 Then CurrORFInd = CurrORFInd - 1
Case NAV_NEXT
     If CurrORFInd < ORFCnt Then CurrORFInd = CurrORFInd + 1
Case NAV_LAST
     CurrORFInd = ORFCnt
Case Else                       'don't go anywhere
End Select
txtORFNavPos.Text = " ORF: " & CurrORFInd & "/" & ORFCnt
If chkSeqOnly.value = vbChecked Then
   rtbORFSeq.Text = GetORFSequence(CallerID, ORFID(CurrORFInd))
Else
   rtbORFSeq.Text = GetORFRecord(CallerID, ORFID(CurrORFInd))
End If
PeptCnt = 0
ReDim PeptID(1 To 1000)      'will be plenty for most proteins
ReDim PeptMW(1 To 1000)
lstPeptides.Clear
Select Case PeptOption
Case OPT_PEPT_DB_ALL
     Call GetPeptidesDBAll
Case OPT_PEPT_DB_LOADED
     Call GetPeptidesDBLoaded
Case OPT_PEPT_DIGESTION
     Call GetPeptidesDigestion
End Select
End Sub

Private Sub Form_Activate()
'------------------------------------------------------------
'load mass tags and ORF databases date if neccessary
'------------------------------------------------------------
On Error Resume Next
CallerID = Me.Tag
Me.MousePointer = vbHourglass
If bLoading Then
   UpdateStatus "Loading ORF list ..."
   LoadORFs CallerID
   UpdateStatus "Number of loaded ORFs: " & ORFCnt
   UpdateStatus "Loading mass tag - ORF mappings ..."
   LoadMassTag_ORFMapping CallerID
   UpdateStatus "ORFs: " & ORFCnt & " Mass Tags: " & AMTCnt & " Mappings: " & MapCnt
   bLoading = False
   Call cmdORFNavigate_Click(0)
End If
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Long
bLoading = True
If AMTCnt > 0 Then
   PeptOption = OPT_PEPT_DB_LOADED         'select ORF peptides among loaded mass tags
   'add mass tags indexes in collection with ID as an key so that we can access them fast
   Set AMTIDCol = New Collection
   For i = 1 To AMTCnt
       AMTIDCol.Add i, AMTID(i)
   Next i
Else                                       'nothing is loaded so that is not an option
   optInclude(OPT_PEPT_DIGESTION).value = True
End If
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub optInclude_Click(Index As Integer)
PeptOption = Index
End Sub

Private Sub GetPeptidesDBAll()

End Sub

Private Sub GetPeptidesDBLoaded()
Dim i As Long
On Error GoTo err_GetPeptidesDBLoaded

'find all peptides among loaded mass tags mapped with current ORF
For i = 1 To MapCnt
    If ORFIDMap(i) = ORFID(CurrORFInd) Then
       PeptCnt = PeptCnt + 1
       PeptID(PeptCnt) = MTIDMap(i)
       PeptMW(PeptCnt) = AMTMW(AMTIDCol.Item(CStr(MTIDMap(i))))
       lstPeptides.AddItem Format$(PeptCnt, "@@@@@@@") & " " & Format$(PeptID(PeptCnt), "@@@@@@@@@@") & " " & Format$(PeptMW(PeptCnt), "0.0000") & "Da"
    End If
Next i
If PeptCnt > 0 Then
   If PeptCnt < UBound(PeptID) Then
      ReDim Preserve PeptID(1 To PeptCnt)
      ReDim Preserve PeptMW(1 To PeptCnt)
   End If
Else
   Erase PeptID
   Erase PeptMW
End If
Exit Sub

err_GetPeptidesDBLoaded:
If Err.Number = 9 Then          'add more room in peptide arrays and continue
   Err.Clear
   ReDim Preserve PeptID(1 To PeptCnt + 1000)
   ReDim Preserve PeptMW(1 To PeptCnt + 1000)
   Resume
End If
End Sub

Private Sub GetPeptidesDigestion()

End Sub


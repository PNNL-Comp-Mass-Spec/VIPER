VERSION 5.00
Begin VB.Form frmORFSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proteins Search"
   ClientHeight    =   5385
   ClientLeft      =   2760
   ClientTop       =   4035
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearchScope 
      Caption         =   "Search Scope"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton optSearchScope 
         Caption         =   "Current view"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optSearchScope 
         Caption         =   "All data points"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.Frame fraModifications 
      Caption         =   "Modifications"
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame fraMWTolerance 
      Caption         =   "MW Tolerance"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
      Begin VB.OptionButton optMWTolType 
         Caption         =   "Da"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optMWTolType 
         Caption         =   "ppm"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "25"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "MW Tolerance"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame fraMWType 
      Caption         =   "MW Type"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
      Begin VB.OptionButton optMWType 
         Caption         =   "The Most Abundant"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   900
         Width           =   1815
      End
      Begin VB.OptionButton optMWType 
         Caption         =   "Monoisotopic"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optMWType 
         Caption         =   "Average"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Menu mnuF 
      Caption         =   "Function"
      Begin VB.Menu mnuFSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuFReport 
         Caption         =   "Report"
      End
      Begin VB.Menu mnuFSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuO 
      Caption         =   "ORFs"
      Begin VB.Menu mnuOLoadLegacyDB 
         Caption         =   "Load From Legacy DB"
      End
      Begin VB.Menu mnuOLoadFASTA 
         Caption         =   "Load From FASTA File"
      End
      Begin VB.Menu mnuOLoadMTDB 
         Caption         =   "Load From MT Tag DB"
      End
      Begin VB.Menu mnuOSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOStatus 
         Caption         =   "Status"
      End
   End
End
Attribute VB_Name = "frmORFSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'search for intact proteins
'created: 10/04/2002 nt
'last modified: 10/04/2002 nt
'-----------------------------------------------------------
Option Explicit

Dim CallerID As Long
Dim bLoading As Boolean

Dim objORF As New dbORF
Dim spName As String            'name of stored procedure to retrieve Proteins
Dim cnString As String          'connection string for MT tag database

Private Sub Form_Activate()
If bLoading Then
   CallerID = Me.Tag
   bLoading = False
   Call mnuOLoadMTDB_Click
End If
End Sub

Private Sub Form_Load()
bLoading = True
With sorfDef
    txtMWTol.Text = .MWTol
    optMWType(.MWField - MW_FIELD_OFFSET).Value = True
    If .MWTolType = gltPPM Then
       optMWTolType(0).Value = True
    Else
       optMWTolType(1).Value = True
    End If
    optSearchScope(.SearchScope).Value = True
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objORF = Nothing
End Sub

Private Sub mnuF_Click()
Call PickParameters
End Sub

' MonroeMod
Private Sub mnuFClose_Click()
    Unload Me
End Sub

Private Sub mnuO_Click()
Call PickParameters
End Sub

Private Sub mnuOLoadMTDB_Click()
'------------------------------------------------------------------------------------
'load Proteins (ORFs) from MT tag database; CallerID display has to be associated with the DB
'for this to work; also it has to know name of procedure that will retrieve Proteins
'------------------------------------------------------------------------------------
Dim Resp As Long
On Error Resume Next
If Not GelAnalysis(CallerID) Is Nothing Then
   If objORF.ORFCount > 0 Then
      Resp = MsgBox("Protein object already contains data. Reload anyway?", vbYesNo, glFGTU)
      If Resp <> vbYes Then Exit Sub
   End If
   cnString = GelAnalysis(CallerID).MTDB.cn.ConnectionString
   spName = glbPreferencesExpanded.MTSConnectionInfo.spGetORFs
   If Len(cnString) > 0 And Len(spName) > 0 Then
      UpdateStatus "Loading Protein data ..."
      If objORF.LoadORFsFromORFDB(cnString, spName) Then
         UpdateStatus "Loaded Proteins: " & objORF.ORFCount
      Else
         UpdateStatus "Error loading Protein data."
      End If
   Else
      MsgBox "Missing parameters necessary to load protein data.", vbOKOnly, glFGTU
   End If
End If
End Sub

Private Sub optMWTolType_Click(Index As Integer)
If Index = 0 Then
    sorfDef.MWTolType = gltPPM
Else
    sorfDef.MWTolType = gltABS
End If
End Sub

Private Sub optMWType_Click(Index As Integer)
sorfDef.MWField = Index + 6
End Sub

Private Sub optSearchScope_Click(Index As Integer)
sorfDef.SearchScope = Index
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   sorfDef.MWTol = CDbl(Abs(txtMWTol.Text))
Else
   MsgBox "This argument should be non-negative number.", vbOKOnly, glFGTU
   txtMWTol.SetFocus
End If
End Sub

Private Sub PickParameters()
Call txtMWTol_LostFocus
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

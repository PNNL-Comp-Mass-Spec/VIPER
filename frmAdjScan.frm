VERSION 5.00
Begin VB.Form frmAdjScan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjacent Scans"
   ClientHeight    =   3915
   ClientLeft      =   2760
   ClientTop       =   4035
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VIPER.LaDist LaDist1 
      Height          =   3855
      Left            =   2040
      TabIndex        =   7
      Top             =   60
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6800
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   ">>"
      Height          =   255
      Index           =   3
      Left            =   1520
      TabIndex        =   12
      ToolTipText     =   "Last"
      Top             =   1200
      Width           =   450
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   ">"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      ToolTipText     =   "Next"
      Top             =   1200
      Width           =   450
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "<"
      Height          =   255
      Index           =   1
      Left            =   560
      TabIndex        =   10
      ToolTipText     =   "Previous"
      Top             =   1200
      Width           =   450
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "<<"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "First"
      Top             =   1200
      Width           =   450
   End
   Begin VB.TextBox txtInfo 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CheckBox chk1_1Link 
      Caption         =   "1 - 1 Link"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtMMA 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "25"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtScan2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtScan1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Scan 2:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Scan 1:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "MMA Tol.(ppm):"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Menu mnuF 
      Caption         =   "&Function"
      Begin VB.Menu mnuFCompare 
         Caption         =   "C&ompare"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFNext 
         Caption         =   "&Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFRandom 
         Caption         =   "&Random"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFLoop 
         Caption         =   "&Loop"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuR 
      Caption         =   "&Report"
   End
End
Attribute VB_Name = "frmAdjScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'created: 08/22/2002 nt
'last modified: 08/24/2002 nt
'-------------------------------------------------------------
Option Explicit

Const NAV_FIRST = 0
Const NAV_PREVIOUS = 1
Const NAV_NEXT = 2
Const NAV_LAST = 3
Const NAV_REC_NUMBER = 4
Const NAV_STAY = 21

Dim CallerID As Long
Dim bLoading As Boolean

Dim Scan1 As Long
Dim Scan2 As Long

Dim MMA As Double
Dim Link1To1 As Boolean

Dim ScInd As ScansIndex             'used to enumerate data by scans

Dim Cnt1 As Long
Dim MW1() As Double

Dim Cnt2 As Long
Dim MW2() As Double

Dim HitsCnt As Long
Dim Hits() As Double

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case NAV_FIRST
     Scan1 = ScInd.Scans(0)
     Scan2 = ScInd.Scans(1)
Case NAV_PREVIOUS
     If Scan1 > ScInd.Scans(0) Then Scan1 = Scan1 - 1
     If Scan2 > ScInd.Scans(0) Then Scan2 = Scan2 - 1
Case NAV_NEXT
     If Scan1 < ScInd.Scans(ScInd.ScansCnt - 1) Then Scan1 = Scan1 + 1
     If Scan2 < ScInd.Scans(ScInd.ScansCnt - 1) Then Scan2 = Scan2 + 1
Case NAV_LAST
     Scan1 = ScInd.Scans(ScInd.ScansCnt - 2)
     Scan2 = ScInd.Scans(ScInd.ScansCnt - 1)
End Select
txtScan1.Text = Scan1
txtScan2.Text = Scan2
Call mnuFCompare_Click
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
UpdateStatus "Enumerating scans ...", False
Call EnumerateScans
UpdateStatus "Scan range: " & ScInd.Scans(0) & " - " & ScInd.Scans(ScInd.ScansCnt - 1), False
Call cmdNav_Click(NAV_FIRST)
End Sub

Private Sub Form_Load()
bLoading = True
Link1To1 = (chk1_1Link.value = vbChecked)
MMA = CDbl(txtMMA.Text)
LaDist1.DFGraphType = GraphBar
End Sub


Private Function CompareAll() As Boolean
'-----------------------------------------------------------------
'compares scan1 and scan 2
'-----------------------------------------------------------------
Dim Dummy() As Long
Dim CurrPPMDiff As Double
Dim MW1Sorter As New QSDouble
Dim MW2Sorter As New QSDouble
Dim k As Long, L As Long
On Error GoTo err_CompareAll

If Cnt1 > 0 And Cnt2 > 0 Then
   'sort ascending masses from both scans so it is easier to search
   If Not MW1Sorter.QSAsc(MW1(), Dummy()) Then
       UpdateStatus "Error sorting sequences from scan " & Scan1, True
       Exit Function
   End If
   If Not MW2Sorter.QSAsc(MW2(), Dummy()) Then
       UpdateStatus "Error sorting sequences from scan " & Scan2, True
       Exit Function
   End If
   HitsCnt = 0
   ReDim Hits(1000)
   For k = 0 To Cnt2 - 1
       For L = 0 To Cnt1 - 1
           CurrPPMDiff = (MW2(k) - MW1(L)) / (MW2(k) * glPPM)
           If Abs(CurrPPMDiff) <= MMA Then     'we have match within range of interest
              HitsCnt = HitsCnt + 1
              Hits(HitsCnt - 1) = CurrPPMDiff
           End If
       Next L
   Next k
   UpdateStatus "Number of matches: " & HitsCnt, True
   If HitsCnt > 0 Then
      ReDim Preserve Hits(HitsCnt - 1)
      LaDist1.DataFill Hits, -MMA, MMA, 2 * MMA + 2, BinsUni
   Else
      Erase Hits
      LaDist1.Clear
   End If
Else
   UpdateStatus "No base for alignment", True
End If
Exit Function

err_CompareAll:
If Err.Number = 9 Then
   ReDim Preserve Hits(HitsCnt + 1000)
   Resume
End If
End Function


Private Function GetScanMasses(ByVal ScanOrder As Long) As Boolean
Dim i As Long
Dim ScanInd As Long
On Error GoTo err_GetScanMasses
Select Case ScanOrder
Case 1
   Cnt1 = 0
   ReDim MW1(1000)
   ScanInd = FindScanIndFast(Scan1, 0, ScInd.ScansCnt - 1)
   If ScanInd < 0 Then Exit Function
   With GelData(CallerID)
        For i = ScInd.CSFirstInd(ScanInd) To ScInd.CSLastInd(ScanInd)
            Cnt1 = Cnt1 + 1
            MW1(Cnt1 - 1) = .CSData(i).AverageMW
        Next i
        For i = ScInd.IsoFirstInd(ScanInd) To ScInd.IsoLastInd(ScanInd)
            Cnt1 = Cnt1 + 1
            MW1(Cnt1 - 1) = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
        Next i
   End With
   If Cnt1 > 0 Then
      ReDim Preserve MW1(Cnt1 - 1)
   Else
      Erase MW1
   End If
Case 2
   Cnt2 = 0
   ReDim MW2(1000)
   ScanInd = FindScanIndFast(Scan2, 0, ScInd.ScansCnt - 1)
   If ScanInd < 0 Then Exit Function
   With GelData(CallerID)
        For i = ScInd.CSFirstInd(ScanInd) To ScInd.CSLastInd(ScanInd)
            Cnt2 = Cnt2 + 1
            MW2(Cnt2 - 1) = .CSData(i).AverageMW
        Next i
        For i = ScInd.IsoFirstInd(ScanInd) To ScInd.IsoLastInd(ScanInd)
            Cnt2 = Cnt2 + 1
            MW2(Cnt2 - 1) = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
        Next i
   End With
   If Cnt2 > 0 Then
      ReDim Preserve MW2(Cnt2 - 1)
   Else
      Erase MW2
   End If
End Select
GetScanMasses = True
Exit Function

err_GetScanMasses:
If Err.Number = 9 Then
   Select Case ScanOrder
   Case 1
        ReDim Preserve MW1(Cnt1 + 500)
   Case 2
        ReDim Preserve MW2(Cnt2 + 500)
   End Select
   Resume
End If
End Function


Private Sub mnuF_Click()
PickParameters
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFCompare_Click()
GetScanMasses 1
UpdateStatus "Scan: " & Scan1 & " Peaks: " & Cnt1 & vbCrLf, False
GetScanMasses 2
UpdateStatus "Scan: " & Scan2 & " Peaks: " & Cnt2 & vbCrLf, True
Call CompareAll
End Sub


Private Sub mnuR_Click()
PickParameters
End Sub

Private Sub txtMMA_LostFocus()
If IsNumeric(txtMMA.Text) Then
   MMA = Abs(CDbl(txtMMA.Text))
Else
   MsgBox "This argument should be positive number.", vbOKOnly, glFGTU
   txtMMA.SetFocus
End If
End Sub


Private Sub txtScan1_LostFocus()
If IsNumeric(txtScan1.Text) Then
   Scan1 = Abs(CLng(txtScan1.Text))
Else
   MsgBox "This argument should be positive integer.", vbOKOnly, glFGTU
   txtScan1.SetFocus
End If
End Sub


Private Sub txtScan2_LostFocus()
If IsNumeric(txtScan2.Text) Then
   Scan2 = Abs(CLng(txtScan2.Text))
Else
   MsgBox "This argument should be positive integer.", vbOKOnly, glFGTU
   txtScan2.SetFocus
End If
End Sub


Private Sub UpdateStatus(ByVal Msg As String, ByVal bAppend As Boolean)
If bAppend Then
   txtInfo.SelStart = Len(txtInfo.Text)
   txtInfo.SelText = Msg & vbCrLf
Else
   txtInfo.Text = Msg
End If
DoEvents
End Sub

Public Sub EnumerateScans()
Dim i As Long
Dim CurrScan As Long
Dim CurrScanInd As Long
With GelData(CallerID)
     ScInd.ScansCnt = UBound(.ScanInfo)
     ReDim ScInd.Scans(ScInd.ScansCnt - 1)
     ReDim ScInd.CSFirstInd(ScInd.ScansCnt - 1)
     ReDim ScInd.CSLastInd(ScInd.ScansCnt - 1)
     ReDim ScInd.IsoFirstInd(ScInd.ScansCnt - 1)
     ReDim ScInd.IsoLastInd(ScInd.ScansCnt - 1)
     For i = 1 To ScInd.ScansCnt
         ScInd.Scans(i - 1) = .ScanInfo(i).ScanNumber
         ScInd.CSLastInd(i - 1) = -1            'so that we can say if there
         ScInd.IsoLastInd(i - 1) = -1           'was nothing in scan
     Next
     CurrScan = -1
     CurrScanInd = -1
     For i = 1 To .CSLines
         If .CSData(i).ScanNumber > CurrScan Then
            CurrScan = .CSData(i).ScanNumber
            CurrScanInd = FindScanIndFast(CurrScan, 0, ScInd.ScansCnt - 1)
            If CurrScanInd >= 0 Then
               ScInd.CSFirstInd(CurrScanInd) = i
               ScInd.CSLastInd(CurrScanInd) = i
            End If
         Else
            If CurrScanInd >= 0 Then ScInd.CSLastInd(CurrScanInd) = i
         End If
     Next i
     CurrScan = -1
     CurrScanInd = -1
     For i = 1 To .IsoLines
         If .IsoData(i).ScanNumber > CurrScan Then
            CurrScan = .IsoData(i).ScanNumber
            CurrScanInd = FindScanIndFast(CurrScan, 0, ScInd.ScansCnt - 1)
            If CurrScanInd >= 0 Then
               ScInd.IsoFirstInd(CurrScanInd) = i
               ScInd.IsoLastInd(CurrScanInd) = i
            End If
         Else
            If CurrScanInd >= 0 Then ScInd.IsoLastInd(CurrScanInd) = i
         End If
     Next i
End With
End Sub

Private Function FindScanIndFast(ScanToFind As Long, _
                                 ByVal MinInd As Long, _
                                 ByVal MaxInd As Long) As Long
'-------------------------------------------------------------
'returns index of ScanToFind in array ScInd.Scans
'-------------------------------------------------------------
Dim MidInd As Long

If ScInd.Scans(MinInd) = ScanToFind Then
   FindScanIndFast = MinInd
   Exit Function
End If
If ScInd.Scans(MaxInd) = ScanToFind Then
   FindScanIndFast = MaxInd
   Exit Function
End If
MidInd = (MinInd + MaxInd) \ 2
If MidInd = MinInd Then         'Min and Max next to each other we didn't find scan
   FindScanIndFast = -1
   Exit Function
End If
If ScInd.Scans(MidInd) = ScanToFind Then
   FindScanIndFast = MidInd
   Exit Function
End If
If ScInd.Scans(MidInd) > ScanToFind Then
   FindScanIndFast = FindScanIndFast(ScanToFind, MinInd, MidInd)
ElseIf ScInd.Scans(MidInd) < ScanToFind Then
   FindScanIndFast = FindScanIndFast(ScanToFind, MidInd, MaxInd)
End If
End Function


Private Sub PickParameters()
Call txtScan1_LostFocus
Call txtScan2_LostFocus
Call txtMMA_LostFocus
End Sub


VERSION 5.00
Begin VB.Form frmFindDenseAreas 
   BackColor       =   &H80000005&
   Caption         =   "Dense Areas Of 2D Display"
   ClientHeight    =   7140
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   4275
   Icon            =   "frmFindDenseAreas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraList 
      BackColor       =   &H80000005&
      Caption         =   "List"
      Height          =   3375
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   4095
      Begin VB.CommandButton cmdView 
         Caption         =   "List Only"
         Height          =   375
         Left            =   3120
         TabIndex        =   30
         ToolTipText     =   "List only view"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3120
         TabIndex        =   29
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdZoomIn 
         Caption         =   "&Zoom In"
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Cl&ear"
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.ListBox lstTMDA 
         BackColor       =   &H80000018&
         Height          =   2985
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox chkAbundanceWeightedCount 
      BackColor       =   &H80000005&
      Caption         =   "Abundance Weighted Count"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Frame fraScope 
      BackColor       =   &H80000005&
      Caption         =   "Scope"
      Height          =   855
      Left            =   3120
      TabIndex        =   18
      Top             =   1320
      Width           =   1095
      Begin VB.OptionButton optScope 
         BackColor       =   &H80000005&
         Caption         =   "&Current"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optScope 
         BackColor       =   &H80000005&
         Caption         =   "&All"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox txtListCnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   16
      Text            =   "100"
      Top             =   2400
      Width           =   495
   End
   Begin VB.Frame fraETRange 
      BackColor       =   &H80000005&
      Caption         =   "Elution Range"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2895
      Begin VB.TextBox txtETOverlap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Text            =   "0.5"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtETRange 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblETOverlapInScans 
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lblETRangeInScans 
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numbers represent % of total duration. (Numbers) are scan count."
         Height          =   465
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2685
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Max elution overlap:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Elution range :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraMWRange 
      BackColor       =   &H80000005&
      Caption         =   "Molecular Mass Range"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton optMWType 
         BackColor       =   &H80000005&
         Caption         =   "&Da"
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   17
         Top             =   780
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optMWType 
         BackColor       =   &H80000005&
         Caption         =   "p&ct"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   540
         Width           =   615
      End
      Begin VB.OptionButton optMWType 
         BackColor       =   &H80000005&
         Caption         =   "&ppm"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtMWOverlap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "0.01"
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtMWRange 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "1"
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Max overlap:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Range width:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List top"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   0
      Top             =   2460
      Width           =   615
   End
End
Attribute VB_Name = "frmFindDenseAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this function could be used to locate dense areas of the
'2D display
'--------------------------------------------------------
'created: 07/01/2002 nt
'last modified: 07/02/2002 nt
'--------------------------------------------------------
Option Explicit

Const lst_Erase = 0
Const lst_Redim = 1

Const ListLimit = 2500      'limit to 2500

Const ViewListOnly = "List Only"
Const ViewEverything = "View All"

Const ReportList = "List"
Const ReportHood = "Hood"

Dim ListTopPosViewAll As Long
Dim ListTopPosListOnly As Long
Dim HViewAll As Long
Dim HViewListOnly As Long

Dim CallerID As Long

Dim MWRange As Double
Dim MWOverlap As Double
Dim MWType As Long

Dim ETRange As Double
Dim ETOverlap As Double

Dim Scope As Long

Dim AbuWeight As Boolean

Dim ListRealCnt As Long     'making list a little bit longer than allowed
Dim ListMaxCnt As Long      'will save us some troubles

'this is actual list of most dense areas
Dim PeakInd() As Long               'index in Num arrays
Dim PeakType() As Long              'Charge State or Isotopic
Dim PeakMW() As Double              'molecular mass
Dim PeakScan() As Long              'scan
Dim PeakMWHalfWidth() As Double     'half width of mass neighborhood
Dim PeakCnt() As Double             'count of peaks in the hood

Dim OID() As Long               'original data  (1 based arrays)
Dim ODT() As Long
Dim OMW() As Double
Dim OAbu() As Double
Dim OFN() As Long
Dim OCnt As Long
Dim LoadedScope As Long

Dim MW() As Double
Dim IndMW() As Long

Dim MWRangeFinder As MWUtil

'ET tolerance is always expressed as percentage of total scan range
Dim FirstScan As Long
Dim LastScan As Long
Dim ScanRange As Long
Dim ScanHalfWidth As Long
Dim ScanOverlapWidth As Long

Dim fso As New FileSystemObject
Dim ts As TextStream
Dim fname As String

Private Sub chkAbundanceWeightedCount_Click()
AbuWeight = (chkAbundanceWeightedCount.value = vbChecked)
End Sub

Private Sub cmdClear_Click()
'---------------------------------------------------
'clear list and erase arrays
'---------------------------------------------------
lstTMDA.Clear
ManageList lst_Erase
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Me.MousePointer = vbHourglass
If Scope <> glSc_All Or LoadedScope <> glSc_All Then
   'only time when we don't have to load is when request is for
   'all data points and all is already loaded
   UpdateStatus "Loading data ..."
   If Not LoadData() Then
      MsgBox "Error loading data.", vbOKOnly, glFGTU
      GoTo exit_cmdFind
   End If
   UpdateStatus "Sorting ..."
   If Not SortDataMW() Then
      MsgBox "Error sorting data on molecular masses.", vbOKOnly, glFGTU
      GoTo exit_cmdFind
   End If
End If
ManageList lst_Redim
Set MWRangeFinder = New MWUtil
If Not MWRangeFinder.Fill(MW) Then
   MsgBox "Error filling data structures.", vbOKOnly, glFGTU
   GoTo exit_cmdFind
End If
UpdateStatus "Looking for dense areas ..."
If FindDense() Then
   Call FillList
Else
   MsgBox "Error searching for dense spots.", vbOKOnly, glFGTU
End If

exit_cmdFind:
UpdateStatus ""
Me.MousePointer = vbDefault
End Sub

Private Sub cmdReport_Click()
Dim i As Long
On Error Resume Next
If PeakCnt(0) > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   WriteReportHeader ReportList
   i = 0
   Do While i < ListMaxCnt
       If PeakCnt(i) > 0 Then
          ts.WriteLine PeakMW(i) & glARG_SEP & PeakMWHalfWidth(i) & glARG_SEP _
            & PeakScan(i) & glARG_SEP & ScanHalfWidth & glARG_SEP & PeakCnt(i)
       End If
       i = i + 1
   Loop
   ts.Close
   Set ts = Nothing
   Me.MousePointer = vbDefault
   UpdateStatus ""
   frmDataInfo.Tag = "Misc"
   frmDataInfo.Show vbModal
   Exit Sub
End If
MsgBox "No dense spots found.", vbOKOnly, glFGTU
End Sub

Private Sub cmdView_Click()
Select Case cmdView.Caption
Case ViewListOnly
     cmdView.Caption = ViewEverything
     fraList.Top = ListTopPosListOnly
     Me.Height = HViewListOnly
Case ViewEverything
     cmdView.Caption = ViewListOnly
     fraList.Top = ListTopPosViewAll
     Me.Height = HViewAll
End Select
End Sub

Private Sub cmdZoomIn_Click()
Dim CurrListInd As Long
CurrListInd = lstTMDA.ListIndex
If CurrListInd >= 0 Then
   GelBody(CallerID).csMyCooSys.ZoomInR _
        PeakScan(CurrListInd) - ScanHalfWidth, _
        PeakMW(CurrListInd) - PeakMWHalfWidth(CurrListInd), _
        PeakScan(CurrListInd) + ScanHalfWidth, _
        PeakMW(CurrListInd) + PeakMWHalfWidth(CurrListInd)
Else
   MsgBox "No list item selected.", vbOKOnly, glFGTU
End If
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
GetScanRange CallerID, FirstScan, LastScan, ScanRange
Call GetETScanCounts
End Sub

Private Sub Form_Load()
'temporary file name (reports)
fname = GetTempFolder() & RawDataTmpFile

cmdView.Caption = ViewListOnly
ListTopPosViewAll = fraList.Top
ListTopPosListOnly = fraMWRange.Top
HViewAll = Me.Height
HViewListOnly = Me.Height - (ListTopPosViewAll - ListTopPosListOnly) * Screen.TwipsPerPixelY
'Debug.Print HViewAll, HViewListOnly

LoadedScope = -1

MWRange = CDbl(txtMWRange.Text)
MWOverlap = CDbl(txtMWOverlap.Text)

ETRange = CDbl(txtETRange.Text)
ETOverlap = CDbl(txtETOverlap.Text)

ListMaxCnt = CLng(txtListCnt.Text)

If optMWType(gltPPM).value Then MWType = gltPPM
If optMWType(gltPct).value Then MWType = gltPct
If optMWType(gltABS).value Then MWType = gltABS

If optScope(glSc_All).value Then Scope = glSc_All
If optScope(glSc_Current).value Then Scope = glSc_Current

AbuWeight = (chkAbundanceWeightedCount.value = vbChecked)
End Sub

Private Sub optMWType_Click(Index As Integer)
MWType = Index
End Sub

Private Sub optScope_Click(Index As Integer)
Scope = Index
End Sub

Private Sub txtETOverlap_LostFocus()
Dim tmp As String
On Error Resume Next
tmp = Trim$(txtETOverlap.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 And tmp <= 100 Then
      ETOverlap = CDbl(tmp)
      Call GetETScanCounts
      Exit Sub
   End If
End If
MsgBox "This argument should be a number between 0 and 100.", vbOKOnly
txtETOverlap.SetFocus
End Sub

Private Sub ManageList(ByVal DoWhat As Long)
Select Case DoWhat
Case lst_Erase
     Erase PeakInd()
     Erase PeakType()
     Erase PeakMW()
     Erase PeakScan()
     Erase PeakCnt()
     Erase PeakMWHalfWidth
Case lst_Redim
     ListRealCnt = ListMaxCnt + 100
     ReDim PeakInd(ListRealCnt - 1)
     ReDim PeakType(ListRealCnt - 1)
     ReDim PeakMW(ListRealCnt - 1)
     ReDim PeakScan(ListRealCnt - 1)
     ReDim PeakCnt(ListRealCnt - 1)
     ReDim PeakMWHalfWidth(ListRealCnt - 1)
End Select
End Sub


Private Sub txtETRange_LostFocus()
Dim tmp As String
On Error Resume Next
tmp = Trim$(txtETRange.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 Then
      ETRange = CDbl(tmp)
      Call GetETScanCounts
      Exit Sub
   End If
End If
MsgBox "This argument should be a non-negative number.", vbOKOnly
txtETRange.SetFocus
End Sub


Private Sub txtListCnt_LostFocus()
Dim tmp As String
On Error Resume Next
tmp = Trim$(txtListCnt.Text)
If IsNumeric(tmp) Then
   If tmp > 0 And tmp <= ListLimit Then
      ListMaxCnt = CLng(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be positive integer up to " & ListLimit & ".", vbOKOnly
txtListCnt.SetFocus
End Sub


Private Sub txtMWOverlap_LostFocus()
Dim tmp As String
On Error Resume Next
tmp = Trim$(txtMWOverlap.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 Then
      MWOverlap = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be a non-negative number.", vbOKOnly
txtMWOverlap.SetFocus
End Sub


Private Sub txtMWRange_LostFocus()
Dim tmp As String
On Error Resume Next
tmp = Trim$(txtMWRange.Text)
If IsNumeric(tmp) Then
   If tmp > 0 Then
      MWRange = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be a non-negative number.", vbOKOnly
txtMWRange.SetFocus
End Sub



Private Function LoadData() As Boolean
'----------------------------------------------
'load arrays neccessary for this procedure
'----------------------------------------------
Dim i As Long
Dim CSCnt As Long
Dim CSInd() As Long
Dim ISCnt As Long
Dim ISInd() As Long
On Error GoTo err_LoadData

With GelData(CallerID)
     ReDim OID(1 To .CSLines + .IsoLines)
     ReDim ODT(1 To .CSLines + .IsoLines)
     ReDim OMW(1 To .CSLines + .IsoLines)
     ReDim OAbu(1 To .CSLines + .IsoLines)
     ReDim OFN(1 To .CSLines + .IsoLines)
     OCnt = 0
     CSCnt = GetCSScope(CallerID, CSInd(), Scope)
     If CSCnt > 0 Then
        For i = 1 To CSCnt
            OCnt = OCnt + 1
            OID(OCnt) = CSInd(i)
            ODT(OCnt) = gldtCS
            OMW(OCnt) = .CSData(CSInd(i)).AverageMW
            OAbu(OCnt) = .CSData(CSInd(i)).Abundance
            OFN(OCnt) = .CSData(CSInd(i)).ScanNumber
        Next i
     End If
     ISCnt = GetISScope(CallerID, ISInd(), Scope)
     If ISCnt > 0 Then
        For i = 1 To ISCnt
            OCnt = OCnt + 1
            OID(OCnt) = ISInd(i)
            ODT(OCnt) = gldtIS
            OMW(OCnt) = GetIsoMass(.IsoData(ISInd(i)), .Preferences.IsoDataField)
            OAbu(OCnt) = .IsoData(ISInd(i)).Abundance
            OFN(OCnt) = .IsoData(ISInd(i)).ScanNumber
        Next i
     End If
End With
If OCnt > 0 Then
   ReDim Preserve OID(1 To OCnt)
   ReDim Preserve ODT(1 To OCnt)
   ReDim Preserve OMW(1 To OCnt)
   ReDim Preserve OAbu(1 To OCnt)
   ReDim Preserve OFN(1 To OCnt)
   'copy OMW array to MW (new thing in VB6 - not using CopyMemory)
   MW() = OMW()
   'initialize index arrays
   ReDim IndMW(1 To OCnt)
   For i = 1 To OCnt
       IndMW(i) = i
   Next i
Else
   Erase OID
   Erase ODT
   Erase OMW
   Erase OAbu
   Erase OFN
   Erase MW
   Erase IndMW
End If
LoadedScope = Scope
LoadData = True
Exit Function

err_LoadData:
OCnt = 0
Erase OID
Erase ODT
Erase OMW
Erase OAbu
Erase OFN
Erase MW
Erase IndMW
End Function

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub FillList()
'-------------------------------------------------
'fill list with current results
'-------------------------------------------------
Dim i As Long
i = 0
'Do While i < ListMaxCnt And PeakCnt(i) > 0
Do While i < ListMaxCnt
   lstTMDA.AddItem "MW: " & Format$(PeakMW(i), "0.0000") & " Scan: " & PeakScan(i) & " Cnt: " & Format$(PeakCnt(i), "0")
   i = i + 1
Loop
End Sub


Private Sub AddToListOrWhat(ByVal Ind As Long, ByVal CurrCnt As Double)
'----------------------------------------------------------------------
'add peak to the "the most dense list" if qualifies; peak is inserted
'on correct spot based on count and overlap requirements
'NOTE: Ind here is index in Oarrays; indexing should be taken into
'      account on previous level
'----------------------------------------------------------------------
Dim i As Long
Dim j As Long
On Local Error GoTo err_AddToListOrWhat
i = ListMaxCnt - 1
Do While i >= 0
   If CurrCnt >= PeakCnt(i) Then
      i = i - 1
   Else
      Exit Do
   End If
Loop
'we have i positioned on first in list that will stay in front of current
'if current overlap more than allowed with any in the list before itself we
'do not have to do anything since current is not going in the list
'On the other side any in the list bellow i will be trown out if overlaps
'more than allowed with the current
If i >= 0 Then
   If i < ListMaxCnt - 1 Then   'if not than current does not qualify for list
      'have to put it in the list so that I can use function IsTooMuchOverlap
      InsertInListOnPos ListRealCnt - 1, Ind, CurrCnt
      'if any in list above overlap with current current is going nowhere
      For j = 0 To i
          If IsTooMuchOverlap(j, ListRealCnt - 1) Then Exit Sub
      Next j
      InsertInListOnPos i + 1, Ind, CurrCnt
      'delete any in the list below current if it overlaps too much
      For j = i + 2 To ListRealCnt - 1
          If IsTooMuchOverlap(i + 1, j) Then RemoveFromList j
      Next j
   End If
Else                            'current goes to the top position
   InsertInListOnPos 0, Ind, CurrCnt
   'eliminate everything that overlaps more than allowed with it
   For j = 1 To ListRealCnt - 1
       If IsTooMuchOverlap(0, j) Then RemoveFromList j
   Next j
End If

err_AddToListOrWhat:        'don't add if any error
End Sub

Private Function SortDataMW() As Boolean
'---------------------------------------------------
'sorts data in MW and IndMW arrays on molecular mass
'---------------------------------------------------
Dim objSort As New QSDouble
If Not objSort.QSAsc(MW(), IndMW()) Then
   SortDataMW = False
Else
   SortDataMW = True
End If
End Function


Private Function FindDense() As Boolean
'------------------------------------------------------------------
'fills arrays indicating dense areas; returns true if everything OK
'NOTE: log10(peak abundance) might be used used to score density
'------------------------------------------------------------------
Dim MWRangeFinder As New MWUtil
Dim MinInd As Long
Dim MaxInd As Long
Dim AbsMWRange As Double
Dim i As Long, j As Long
Dim CurrCnt As Long
On Error GoTo err_FindDense

If Not MWRangeFinder.Fill(MW) Then GoTo err_FindDense
For i = 1 To OCnt
    ' MonroeMod
    If i Mod 25 = 0 Then UpdateStatus Trim(i) & " / " & Trim(OCnt)
    CurrCnt = 0
    Select Case MWType
    Case gltPPM
         AbsMWRange = OMW(IndMW(i)) * MWRange * glPPM
    Case gltPct
         AbsMWRange = OMW(IndMW(i)) * MWRange * glPCT
    Case gltABS
         AbsMWRange = MWRange
    End Select
    MinInd = 0
    MaxInd = 0
    MWRangeFinder.FindIndexRange OMW(IndMW(i)), AbsMWRange, MinInd, MaxInd
    If MinInd <= MaxInd Then
       'for any peak close enough add log10 of its peak intensity
       For j = MinInd To MaxInd
           If Abs(OFN(IndMW(j)) - OFN(IndMW(i))) <= ScanHalfWidth Then
              If AbuWeight Then
                 CurrCnt = CurrCnt + Log(OAbu(IndMW(j))) / Log(10#)
              Else
                 CurrCnt = CurrCnt + 1
              End If
           End If
       Next j
    End If
    If CurrCnt > 0 Then AddToListOrWhat IndMW(i), CurrCnt
Next i
FindDense = True
Exit Function

err_FindDense:
End Function


Private Sub RemoveFromList(ByVal Ind As Long)
'--------------------------------------------
'removes element with index Ind from the list
'--------------------------------------------
Dim i As Long
For i = Ind To ListRealCnt - 2
    PeakInd(i) = PeakInd(i + 1)
    PeakType(i) = PeakInd(i + 1)
    PeakMW(i) = PeakMW(i + 1)
    PeakScan(i) = PeakScan(i + 1)
    PeakMWHalfWidth(i) = PeakMWHalfWidth(i + 1)
    PeakCnt(i) = PeakCnt(i + 1)
Next i
PeakInd(ListRealCnt - 1) = 0            'this is not neccessary but...
PeakType(ListRealCnt - 1) = 0
PeakMW(ListRealCnt - 1) = 0
PeakScan(ListRealCnt - 1) = 0
PeakMWHalfWidth(ListRealCnt - 1) = 0
PeakCnt(ListRealCnt - 1) = 0
End Sub


Private Sub InsertInListOnPos(ByVal pos As Long, _
                              ByVal Ind As Long, _
                              ByVal CurrCnt As Double)
'-----------------------------------------------------------
'inserts element from O arrays with index Ind in Peaks array
'on position Pos (everything is pushed down after Pos)
'-----------------------------------------------------------
Dim i As Long
For i = ListRealCnt - 1 To pos + 1 Step -1
    PeakInd(i) = PeakInd(i - 1)
    PeakType(i) = PeakInd(i - 1)
    PeakMW(i) = PeakMW(i - 1)
    PeakScan(i) = PeakScan(i - 1)
    PeakMWHalfWidth(i) = PeakMWHalfWidth(i - 1)
    PeakCnt(i) = PeakCnt(i - 1)
Next i
PeakInd(pos) = OID(Ind)
PeakType(pos) = ODT(Ind)
PeakMW(pos) = OMW(Ind)
PeakScan(pos) = OFN(Ind)
Select Case MWType
Case gltPPM
     PeakMWHalfWidth(pos) = PeakMW(pos) * MWRange * glPPM
Case gltPct
     PeakMWHalfWidth(pos) = PeakMW(pos) * MWRange * glPCT
Case gltABS
     PeakMWHalfWidth(pos) = MWRange
End Select
PeakCnt(pos) = CurrCnt
End Sub


Public Function IsTooMuchOverlap(ByVal i As Long, _
                                 ByVal j As Long) As Boolean
'--------------------------------------------------------------------
'returns true if peak array members i and j overlap more than allowed
'--------------------------------------------------------------------
Dim AbsMWDiff As Double
Dim AbsScanDiff As Double
Dim MaxMWOverlap As Double      'allowed mol.mass overlap
Dim ThisMWOverlap
Dim ThisScanOverlap As Double
If PeakCnt(i) <= 0 Or PeakCnt(j) <= 0 Then Exit Function
AbsScanDiff = Abs(PeakScan(i) - PeakScan(j))
ThisScanOverlap = (2 * ScanHalfWidth - AbsScanDiff)
If ThisScanOverlap <= 0 Then Exit Function                   'no overlap
If ThisScanOverlap < ScanOverlapWidth Then Exit Function     'not too much overlap
AbsMWDiff = Abs(PeakMW(i) - PeakMW(j))
Select Case MWType
Case gltPPM      'take middle between two points
       MaxMWOverlap = ((PeakMW(i) + PeakMW(j)) / 2) * MWOverlap * glPPM
Case gltPct
       MaxMWOverlap = ((PeakMW(i) + PeakMW(j)) / 2) * MWOverlap * glPCT
Case gltABS
       MaxMWOverlap = MWOverlap
End Select
ThisMWOverlap = (PeakMWHalfWidth(i) + PeakMWHalfWidth(j)) - AbsMWDiff
If ThisMWOverlap <= 0 Then Exit Function                    'no overlap
If ThisMWOverlap < MaxMWOverlap Then Exit Function          'not too much overlap
IsTooMuchOverlap = True
End Function


Public Sub GetETScanCounts()
ScanHalfWidth = CLng(ScanRange * ETRange / 100)     'this does not change
lblETRangeInScans.Caption = "(" & ScanHalfWidth & ")"
ScanOverlapWidth = CLng(ScanRange * ETOverlap / 100)
lblETOverlapInScans.Caption = "(" & ScanOverlapWidth & ")"
End Sub

Public Sub WriteReportHeader(ByVal ReportType As String)
ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
ts.WriteLine "Gel File: " & GelBody(CallerID).Caption
ts.WriteLine "Total distributions: " & GelData(CallerID).DataLines
ts.WriteLine
ts.WriteLine "Reporting the most dense areas - " & ReportType
If AbuWeight Then
   ts.WriteLine "Count: Abundance weighted (log10)"
Else
   ts.WriteLine "Count: Normal"
End If
ts.WriteLine
Select Case ReportType
Case ReportList
    ts.WriteLine "MW" & glARG_SEP & "MW Range(Da)" & glARG_SEP & "Scan" _
                & glARG_SEP & "Scan Range(Scans)" & glARG_SEP & "Count"
Case ReportHood
    ts.WriteLine "Peak Ind" & glARG_SEP & "Scan" & glARG_SEP & "CS" _
            & glARG_SEP & "MW" & glARG_SEP & "Abu" & glARG_SEP & "Fit"
End Select
End Sub

VERSION 5.00
Begin VB.Form frmIntCalLM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internal Calibration Lock Mass"
   ClientHeight    =   3135
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   FillStyle       =   0  'Solid
   Icon            =   "frmIntCalLM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   4800
      TabIndex        =   17
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdResetFreqShift 
      Caption         =   "Reset Freq.Shifts"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtLastFN 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      ToolTipText     =   "Double-click for default value"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtFirstFN 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Double-click for default value"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "C&ontinue"
      Height          =   315
      Left            =   3360
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox chkUseUnlocked 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Use scans not locked"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "If checked only scans that has not been locked will be used"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox chkStopAfterScan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop &after each scan"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report"
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdLockMass 
      Caption         =   "&Lock Mass"
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtSmallReport 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   2055
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmIntCalLM.frx":030A
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox txtCalMWErr 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "25"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtCalMW 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "531.02"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   15
      Top             =   1560
      Width           =   165
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scans"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max. Error (ppm)"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calibration Mass (Da)"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmIntCalLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'internal calibration lock mass
'Behavior: works with .ScanInfo().FrequencyShift data
' in GelData structure
'Loads .FrequencyShift to FNS() array of local frequency shifts
'No change takes place in gel file until explicit save
'is issued (save when closing form or with Save button)
'Saving takes effect only in scans with count of matches
'positive and frequency shift <> 0
'Reseting frequency shifts works only on current scans
'------------------------------------------------------------
'created: 09/04/2001 nt
'last modified: 09/06/2001 nt
'------------------------------------------------------------
Option Explicit

'values of next 3 constants are not the same as in frmUMCLockMass
Const LM_FN_ERROR = -2
Const LM_FN_NOATTEMPT = -1
Const LM_FN_NOCANDIDATES = 0

Const REPORT_PROCESSING = 1
Const REPORT_RESULTS = 2

Dim CallerID As Long

Dim m_CalMW As Double           'calibration molecular mass
Dim m_CalMWErr As Double        'allowed error (ppm)
Dim m_AbsMWErr As Double        'absolute error(Da)

Dim FirstFN As Long             'for processing
Dim LastFN As Long
Dim MinFN As Long               'in gel file
Dim MaxFN As Long
Dim CurrFN As Long              'current scan (not processed)

'contains count and indexes of masses of all distributions from
'current scan matching calibrant mass
Dim CurrLckeeCnt As Long
Dim CurrLckee() As Long

Dim StopAfterEachScan As Boolean
'Scan is considered to be locked if it has non-zero frequency
'shift stored in FrequencyShift array of GelData structure
Dim UseUnlockedScans As Boolean
Dim WasLockMassUsed As Boolean
Dim StopProcessing As Boolean

'frequency shifts array(indexed from MinFN to MaxFN)
Dim FNS() As Double
Dim FNMW() As Double        'molecular mass that locks this scan
Dim FNMWCnt() As Long       'count of matching distributions
'scan number statistic arrays(indexed from MinFN to MaxFN)
Dim FNCnt() As Long      'total number of points in a scan
Dim FNIndMin() As Long   'first index of mw in a scan
Dim FNIndMax() As Long   'last index of mw in a scan

'Molecular Mass arrays
Dim MWID() As Long          'index in GelData array
Dim MWFN() As Integer       'scan numbers(ordered parallel with MWInd)
Dim MWLM() As Double        'initially original mass; after lock mass function
                            'new mass (or original mass where not locked)
Dim MWInd() As Long         'indexes in this array; used for sort with
                            'QSInteger object; indirect access to real data
Dim MWType() As Integer     'deconvolution type
Dim MWCnt As Long           'total count of data points

Dim Cal As Object           'calibration recalculation object

Private Sub chkStopAfterScan_Click()
StopAfterEachScan = (chkStopAfterScan.value = vbChecked)
End Sub

Private Sub chkUseUnlocked_Click()
UseUnlockedScans = (chkUseUnlocked.value = vbChecked)
End Sub

Private Sub cmdClose_Click()
Dim Res As Long
If WasLockMassUsed Then
   Res = MsgBox("Replace original masses with locked masses?", vbYesNoCancel, glFGTU)
   Select Case Res
   Case vbYes          'save an unload
     SaveResults
     Unload Me
   Case vbNo           'unload without saving
     Unload Me
   Case vbCancel       'do nothing
   End Select
Else
   Unload Me
End If
End Sub

Private Sub cmdContinue_Click()
'------------------------------------------------------------
'this approach allows to search one part of the file with one
'calibrant and next with another
'------------------------------------------------------------
StopProcessing = False
DoEvents
If CurrFN < LastFN Then
   If GetSearchParameters() Then
      LockMass CurrFN + 1, LastFN
   End If
End If
End Sub

Private Sub cmdLockMass_Click()
If GetSearchParameters() Then
   If FillArrays() Then
      Me.MousePointer = vbHourglass
      LockMass FirstFN, LastFN
      Me.MousePointer = vbDefault
   Else
      ReportSomething "Error initializing data structures. Can not perform lock mass function."
      EnableCommands False
   End If
End If
End Sub

Private Sub cmdReport_Click()
Dim FileNum As Integer
Dim FileNam As String
Dim sLine As String
Dim i As Long
On Error Resume Next

If WasLockMassUsed Then
  If MinFN <= MaxFN Then
     Me.MousePointer = vbHourglass
     ReportSomething "Generating report..."
     FileNum = FreeFile
     FileNam = GetTempFolder() & RawDataTmpFile
     Open FileNam For Output As FileNum
     'print gel file name and Search definition as reference
     Print #FileNum, "Gel File: " & GelBody(CallerID).Caption
     Print #FileNum, "Lock mass on internal calibrant"
     sLine = "Scan #" & vbTab & "Count" & vbTab & "Cal MW" _
           & vbTab & "  Lock.Cnt  " & vbTab & "Freq.Shift"
     Print #FileNum, sLine
     For i = MinFN To MaxFN
       sLine = i & vbTab & FNCnt(i) & vbTab
       Select Case FNMWCnt(i)
       Case Is > 0
          sLine = sLine & FNMW(i) & vbTab & FNMWCnt(i) & vbTab & FNS(i)
       Case LM_FN_NOCANDIDATES
         sLine = sLine & "No lock mass candidates found in this scan."
       Case LM_FN_ERROR
         sLine = sLine & "Error locking masses in this scan."
       Case LM_FN_NOATTEMPT
         sLine = sLine & "No lock mass attempt made in this scan."
       End Select
       Print #FileNum, sLine
     Next i
     Close FileNum
     ReportSomething Null
     frmDataInfo.Tag = "INCALLM"
     frmDataInfo.Show vbModal
  Else
     MsgBox "No scans found.", vbOKOnly
  End If
  Me.MousePointer = vbDefault
Else
  MsgBox "No mass locking applied yet.", vbOKCancel, "2DGelLand - Lock Mass"
End If
End Sub

Private Sub cmdResetFreqShift_Click()
'------------------------------------------------
'resets frequency shift array for specified scans
'------------------------------------------------
Dim i As Long
On Error Resume Next
For i = FirstFN To LastFN
    FNS(i) = 0
    FNMW(i) = 0
    FNMWCnt(i) = LM_FN_NOATTEMPT
Next i
End Sub

Private Sub cmdSave_Click()
SaveResults
End Sub

Private Sub cmdStop_Click()
StopProcessing = True
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
With GelData(CallerID)
   Select Case UCase(.CalEquation)
   Case UCase(CAL_EQUATION_1)
     Set Cal = New CalEq1
     Cal.A = .CalArg(1)
     Cal.B = .CalArg(2)
     If Cal.EquationOK Then
        EnableCommands True
     Else
        MsgBox "Invalid parameters in calibration equation.", vbOKOnly, glFGTU
        EnableCommands False
     End If
   Case Else     'later we might use other calibration equations
     MsgBox "Calibration equation not found. Lock mass function not possible.", vbOKOnly, glFGTU
     EnableCommands False
   End Select
   'load all scans
   MinFN = .ScanInfo(1).ScanNumber
   MaxFN = .ScanInfo(UBound(.ScanInfo)).ScanNumber
   FirstFN = MinFN
   LastFN = MaxFN
   txtFirstFN.Text = FirstFN
   txtLastFN.Text = LastFN
End With
End Sub

Private Sub Form_Load()
m_CalMW = CDbl(txtCalMW.Text)
m_CalMWErr = CDbl(txtCalMWErr.Text)
End Sub

Private Sub txtCalMW_LostFocus()
On Error Resume Next
If IsNumeric(txtCalMW.Text) Then
   m_CalMW = CDbl(txtCalMW.Text)
Else
   MsgBox "This argument has to be positive number.", vbOKOnly, glFGTU
   txtCalMW.SetFocus
End If
End Sub


Private Sub txtCalMWErr_LostFocus()
On Error Resume Next
If IsNumeric(txtCalMWErr.Text) Then
   m_CalMWErr = CDbl(txtCalMWErr.Text)
Else
   MsgBox "This argument has to be positive number.", vbOKOnly, glFGTU
   txtCalMWErr.SetFocus
End If
End Sub

Private Sub ReportSomething(ByVal ReportWhat As Variant)
'-------------------------------------------------------
'do small report so that user knows what is going on
'send Null to clear report; text to display text; or
'long constant to report predetermined processing info
'-------------------------------------------------------
Dim RepText As String
If IsNumeric(ReportWhat) Then
    Select Case CLng(ReportWhat)
    Case REPORT_PROCESSING
         RepText = "Processing scan: " & CurrFN & " ... " & vbCrLf _
            & "Total distributions: " & FNCnt(CurrFN)
    Case REPORT_RESULTS
         RepText = "Results for scan: " & CurrFN
    End Select
ElseIf IsNull(ReportWhat) Then
    RepText = ""
Else
    RepText = CStr(ReportWhat)
End If
txtSmallReport.Text = RepText
DoEvents
End Sub

Private Function FillArrays() As Boolean
'-------------------------------------------------
'returns True if data structures are successfully
'initialized and filled; False on any error
'-------------------------------------------------
Dim CurrFN As Long, PrevFN As Long
Dim i As Long
Dim FNSorter As New QSInteger
On Error Resume Next

MWCnt = 0
ReportSomething "Loading data structures..."
With GelData(CallerID)
    If .CSLines + .IsoLines > 0 Then
       ReDim MWID(1 To .CSLines + .IsoLines)
       ReDim MWFN(1 To .CSLines + .IsoLines)
       ReDim MWLM(1 To .CSLines + .IsoLines)
       ReDim MWType(1 To .CSLines + .IsoLines)
       ReDim MWInd(1 To .CSLines + .IsoLines)
    Else
       ReportSomething Null
       Exit Function
    End If
    If .CSLines > 0 Then
       For i = 1 To .CSLines
           MWCnt = MWCnt + 1
           MWID(MWCnt) = i
           MWType(MWCnt) = glCSType
           MWFN(MWCnt) = .CSData(i).ScanNumber
           MWLM(MWCnt) = .CSData(i).AverageMW
           MWInd(MWCnt) = MWCnt      'index for indirect access
       Next
    End If
    If .IsoLines > 0 Then
       For i = 1 To .IsoLines
           MWCnt = MWCnt + 1
           MWID(MWCnt) = i
           MWType(MWCnt) = glIsoType
           MWFN(MWCnt) = .IsoData(i).ScanNumber
           MWLM(MWCnt) = GetIsoMass(.IsoData(i), amtlmDef.lmIsoField)
           MWInd(MWCnt) = MWCnt      'index for indirect access
       Next
    End If
    If (MWCnt > 0) Then        'initialize FN-arrays
        ReDim FNLM(MinFN To MaxFN)
        ReDim FNS(MinFN To MaxFN)
        ReDim FNMW(MinFN To MaxFN)
        ReDim FNMWCnt(MinFN To MaxFN)
        ReDim FNCnt(MinFN To MaxFN)
        ReDim FNIndMin(MinFN To MaxFN)
        ReDim FNIndMax(MinFN To MaxFN)
    Else
        ReportSomething Null
        Exit Function
    End If
    'load stored frequency shifts
    For i = MinFN To MaxFN
        FNS(i) = .ScanInfo(GetDFIndex(CallerID, i)).FrequencyShift
    Next i
End With

'order MW on FN and count original data
ReportSomething "Analyzing original data..."
If Not FNSorter.QSAsc(MWFN(), MWInd()) Then
   ReportSomething Null
   Exit Function
Else
   PrevFN = MinFN
   CurrFN = -1
   For i = 1 To MWCnt
       If MWFN(i) <> CurrFN Then
          'note last reference index of previous scan (except for first element)
          If PrevFN = CurrFN Then FNIndMax(PrevFN) = i - 1
          CurrFN = MWFN(i)
          PrevFN = CurrFN
          FNIndMin(CurrFN) = i      'note first reference index of this scan
       End If
       FNCnt(CurrFN) = FNCnt(CurrFN) + 1
   Next i
   FNIndMax(MaxFN) = MWCnt          'note last reference index of the last scan
End If
ReportSomething Null
FillArrays = True
End Function


Private Sub LockMass(ByVal FirstScan As Long, ByVal LastScan As Long)
'--------------------------------------------------------------------
'does lock mass and; FirstScan is the scan to start processing with
'LastScan the one to finish with; keeps CurrFN on last processed scan
'--------------------------------------------------------------------
On Error Resume Next

WasLockMassUsed = True
For CurrFN = FirstScan To LastScan
    DoEvents
    If StopProcessing Then Exit For
    If UseUnlockedScans And FNS(CurrFN) <> 0 Then
       FNMWCnt(CurrFN) = LM_FN_NOATTEMPT
    Else
       ReportSomething REPORT_PROCESSING
       'processing here
       FNMW(CurrFN) = m_CalMW          'remember locker mass
       'do selection of lockees
       ReportSomething txtSmallReport.Text & vbCrLf & LMSelect()
       FNMWCnt(CurrFN) = CurrLckeeCnt
       ReportSomething txtSmallReport.Text & vbCrLf & LMAvgFreqShift()
    End If
    If CurrFN < LastScan Then       'no need to stop if last scan
       If StopAfterEachScan Then Exit For
    End If
    DoEvents
Next
If StopProcessing Then
   ReportSomething txtSmallReport.Text & vbCrLf & "Processing paused. Press Continue to continue from scan: " & CurrFN + 1 & " or Mass Lock to start it over."
Else                    'end of scans or checked "Stop after each scan"
   If CurrFN > LastScan Then
      ReportSomething txtSmallReport.Text & vbCrLf & "End of scans."
   Else
      ReportSomething txtSmallReport.Text & vbCrLf & "Click Continue to process next scan."
   End If
End If
End Sub

Private Sub EnableCommands(ByVal Arg As Boolean)
cmdLockMass.Enabled = Arg
cmdStop.Enabled = Arg
cmdContinue.Enabled = Arg
cmdReport.Enabled = Arg
End Sub


Private Function GetSearchParameters() As Boolean
'----------------------------------------------------------
'retrieves absolute error and makes sure numbers make sense
'----------------------------------------------------------
If m_CalMW > 0 And m_CalMWErr >= 0 Then
   m_AbsMWErr = m_CalMW * m_CalMWErr * glPPM
   GetSearchParameters = True
   EnableCommands True
Else
   ReportSomething "Calibration MW: " & m_CalMW & " Da" & vbCrLf & "Tolerance: " & m_CalMWErr & " ppm" & vbCrLf & "Could not agree with the choice. Sorry." & vbCrLf & "Your Server"
   GetSearchParameters = False
   EnableCommands False
End If
End Function

Private Sub txtFirstFN_DblClick()
txtFirstFN.Text = MinFN
End Sub

Private Sub txtFirstFN_LostFocus()
If IsNumeric(txtFirstFN.Text) Then
   FirstFN = CLng(txtFirstFN.Text)
   If FirstFN >= MinFN And FirstFN <= MaxFN Then Exit Sub
End If
MsgBox "This argument should be number between " & MinFN & " and " & MaxFN & "."
txtFirstFN.SetFocus
End Sub

Private Sub txtLastFN_DblClick()
txtLastFN.Text = MaxFN
End Sub

Private Sub txtLastFN_LostFocus()
If IsNumeric(txtLastFN.Text) Then
   LastFN = CLng(txtLastFN.Text)
   If LastFN >= MinFN And LastFN <= MaxFN Then Exit Sub
End If
MsgBox "This argument should be number between " & MinFN & " and " & MaxFN & "."
txtLastFN.SetFocus
End Sub

Private Function LMSelect() As String
'----------------------------------------------------
'does selection of lockees masses for scan CurrFN and
'returns message indicating status of operation.
'----------------------------------------------------
Dim i As Long
Dim tmp As String
On Error GoTo err_LMSelect
CurrLckeeCnt = 0
ReDim CurrLckee(100)
If FNIndMin(CurrFN) <= FNIndMax(CurrFN) And FNIndMax(CurrFN) > 0 Then
    For i = FNIndMin(CurrFN) To FNIndMax(CurrFN)
        If Abs(MWLM(MWInd(i)) - m_CalMW) <= m_AbsMWErr Then
            CurrLckeeCnt = CurrLckeeCnt + 1
            CurrLckee(CurrLckeeCnt - 1) = MWInd(i)
            tmp = tmp & Format$(MWLM(MWInd(i)), "0.0000") & ";"
        End If
    Next i
End If
If Len(tmp) > 0 Then tmp = Left$(tmp, Len(tmp) - 1)
LMSelect = "Total locking candidates: " & CurrLckeeCnt & vbCrLf & Trim$(tmp)
Exit Function

err_LMSelect:
Select Case Err.Number
Case 9          'make sure error is not on some other array place
     ReDim Preserve CurrLckee(CurrLckeeCnt + 100)
     Resume
Case Else
     LMSelect = "Error: " & Err.Number & " - " & Err.Description & " (during lockers selection)."
     FNMWCnt(CurrFN) = LM_FN_ERROR
     chkStopAfterScan.value = vbChecked
     DoEvents
End Select
End Function

Private Function LMAvgFreqShift() As String
'------------------------------------------------------
'does actual calculation of frequency shifts and makes
'official frequency shift the average of all shifts
'also recalculates molecular masses based on this shift
'------------------------------------------------------
Dim i As Long, j As Long
Dim TtlFS As Double
Dim CS As Double
Dim MOverZ As Double
Dim FreqE As Double         'experimental mass frequency
Dim FreqT As Double         'theoretical(calibration) mass frequency
On Error GoTo err_LMAvgFreqShift
With GelData(CallerID)
   If CurrLckeeCnt > 0 Then
      TtlFS = 0
      For i = 0 To CurrLckeeCnt - 1
          Select Case MWType(CurrLckee(i))
          Case glCSType    'can always go with 1st charge state
            CS = .CSData(MWID(CurrLckee(i))).Charge
            If CS > 0 Then
               'calculate frequency for experimental mw
               MOverZ = MWLM(CurrLckee(i)) / CS + glMASS_CC
               FreqE = Cal.CyclotFreq(MOverZ)
               'calculate frequency for theoretical mw
               MOverZ = m_CalMW / CS + glMASS_CC
               FreqT = Cal.CyclotFreq(MOverZ)
               'we have now frequency shift
               TtlFS = TtlFS + (FreqT - FreqE)
            End If
          Case glIsoType
            CS = .IsoData(MWID(CurrLckee(i))).Charge
            If CS > 0 Then
               'calculate frequency for experimental mw
               MOverZ = MWLM(CurrLckee(i)) / CS + glMASS_CC
               FreqE = Cal.CyclotFreq(MOverZ)
               'calculate frequency for theoretical mw
               MOverZ = m_CalMW / CS + glMASS_CC
               FreqT = Cal.CyclotFreq(MOverZ)
               'we have now frequency shift
               TtlFS = TtlFS + (FreqT - FreqE)
            End If
          End Select
      Next i
      FNS(CurrFN) = TtlFS / CurrLckeeCnt
      'now recalculate all masses
      For j = FNIndMin(CurrFN) To FNIndMax(CurrFN)
        Select Case MWType(MWInd(j))
        Case glCSType
           CS = 1
           If CS > 0 Then
              MOverZ = MWLM(MWInd(j)) / CS + glMASS_CC
              FreqE = Cal.CyclotFreq(MOverZ) + FNS(CurrFN)
              MOverZ = Cal.MOverZ(FreqE)
              MWLM(MWInd(j)) = CS * (MOverZ - glMASS_CC)
           End If
        Case glIsoType
           CS = .IsoData(MWID(MWInd(j))).Charge
           If CS > 0 Then
              MOverZ = MWLM(MWInd(j)) / CS + glMASS_CC
              FreqE = Cal.CyclotFreq(MOverZ) + FNS(CurrFN)
              MOverZ = Cal.MOverZ(FreqE)
              MWLM(MWInd(j)) = CS * (MOverZ - glMASS_CC)
           End If
        End Select
      Next j
      LMAvgFreqShift = "Avg.Freq.Shift: " & FNS(CurrFN)
   Else                    'no candidates; frequency shift is 0
      FNS(CurrFN) = 0
      LMAvgFreqShift = "Avg.Freq.Shift: 0 (no matches)"
   End If
End With
Exit Function

err_LMAvgFreqShift:
LMAvgFreqShift = "Error: " & Err.Number & " - " & Err.Description & " (during lockers selection)."
FNMWCnt(CurrFN) = LM_FN_ERROR
chkStopAfterScan.value = vbChecked
DoEvents
End Function

Private Sub SaveResults()
Dim i As Long, j As Long
Dim FNInd As Long
On Error Resume Next
ReportSomething "Saving results..."
Me.MousePointer = vbHourglass
DoEvents
With GelData(CallerID)
    'write frequency shifts and data changes
    For j = MinFN To MaxFN
        'we do not have neccesserilly all them in gel
        FNInd = GetDFIndex(CallerID, j)
        If FNInd > 0 Then
           If FNMWCnt(j) > 0 Then       'only if it was locked
              .ScanInfo(FNInd).FrequencyShift = FNS(j)
              For i = FNIndMin(j) To FNIndMax(j)
                  Select Case MWType(MWInd(i))
                  Case glCSType
                    .CSData(MWID(MWInd(i))).AverageMW = MWLM(MWInd(i))
                  Case glIsoType
                    SetIsoMass .IsoData(MWID(MWInd(i))), amtlmDef.lmIsoField, MWLM(MWInd(i))
                  End Select
              Next i
           End If
        End If
    Next j
    .Comment = .Comment & vbCrLf & Now() & vbCrLf & "Lock mass function based on internal calibration applied."
End With
GelStatus(CallerID).Dirty = True
ReportSomething ""
WasLockMassUsed = False         'since changes were saved
Me.MousePointer = vbDefault
End Sub

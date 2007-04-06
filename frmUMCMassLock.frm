VERSION 5.00
Begin VB.Form frmUMCLockMass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MT Lock Mass Function"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmUMCMassLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMassTag 
      Height          =   285
      Left            =   2760
      TabIndex        =   20
      Top             =   2840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelectLockers 
      Caption         =   "&Select Lockers"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame fraResults 
      Caption         =   "Save Results"
      Height          =   1575
      Left            =   2400
      TabIndex        =   9
      ToolTipText     =   "This option is not implemented until dialog is unloaded."
      Top             =   1200
      Width           =   2055
      Begin VB.OptionButton optLMResults 
         Caption         =   "Save in &New Gel"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Creates duplicates of the gel with new masses and without MT references"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optLMResults 
         Caption         =   "Save in &Original Gel"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Replace masses in original gel and cleans MT references"
         Top             =   720
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optLMResults 
         Caption         =   "&Don't Save"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Re&port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "&Lock Mass"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame fraClsScore 
      Caption         =   "Score Classes on"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
      Begin VB.OptionButton optScore 
         Caption         =   "&Number of MT Hits"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optScore 
         Caption         =   "A&MT Fit"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton optScore 
         Caption         =   "C&alculated Fit(Asc.)"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optScore 
         Caption         =   "In&tensity(Desc.)"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   2295
      ScaleWidth      =   5055
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblMassTag 
      Caption         =   "MT Tag:"
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "This function always attempts to lock as many as possible data points."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   800
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "This lock mass function is based on Unique Mass Classes."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   $"frmUMCMassLock.frx":000C
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "frmUMCLockMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'locking masses on AMT search results based on the current UMC;
'last modified 07/18/2002 nt
'--------------------------------------------------------------
Option Explicit

Dim CallerID As Long

Const FN_LM_NOATTEMPT = 0
Const FN_LM_NOCANDIDATES = -1
Const FN_LM_ERROR = -2

'used in special case
Dim MassTag As Double

Dim Selected As Boolean
Dim Locked As Boolean
Dim bCancel As Boolean
'scan number limits for current scope
Dim MinFN As Integer
Dim MaxFN As Integer
'statistic array that describes UMC
Dim ClsStat() As Double
Dim ClsCnt As Long          'count of UMC
Dim ClsScore() As Double    'score for the class based on user selection
                            'score based on intensity(average), fit(average)
                            'best AMT hit or highest number of AMT hits
Dim ClsAMT() As String      'AMT that best describes class
Dim ClsHits() As Long       'class hits of best AMT

'variable used to score LC-MS Features
Dim ScoreOption As Integer      '0 - score on Intensity(average class Intensity
                                '1 - score on Fit(average class Fit)
                                '2 - score on class best AMT fit
                                '3 - score on class most AMT hits

Dim FNCls() As Long             'class that will lock this scan(indexed from MinFN to MaxFN)
Dim FNClsLM() As Long           'Isotopic point that will lock this scan
'lock masses array-index in MWInd array of mass that locked scan,
'-1 if no lock mass in scan, -2 if error while locking this scan
'0 if no lock mass attempted on this scan
Dim FNLM() As Long

'frequency shifts array(indexed from MinFN to MaxFN)
Dim FNS() As Double
'AMT ID on which masses are locked
Dim FNAMT() As String
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

Private Sub cmdCancel_Click()
bCancel = True
End Sub

Private Sub cmdClose_Click()
Dim Respond
Dim i As Long
Dim j As Integer
Dim FNIndex As Integer
If Locked Then
  Select Case amtlmDef.lmSaveResults
  Case glLM_SAVE_NOT
  Case glLM_SAVE_ORIGINAL
    Respond = MsgBox("You selected to save results in original gel. This will replace original molecular masses and remove all MT references in the currentlly selected scope.", vbOKCancel)
    If Respond = vbOK Then
       lblStatus.Caption = "Saving results..."
       Me.MousePointer = vbHourglass
       DoEvents
       RemoveAMT CallerID, amtlmDef.lmScope
       With GelData(CallerID)
         For i = 1 To MWCnt
           Select Case MWType(i)
           Case glCSType
             .CSData(MWID(i)).AverageMW = MWLM(i)
           Case glIsoType
             SetIsoMass .IsoData(MWID(i)), amtlmDef.lmIsoField, MWLM(i)
           End Select
         Next i
         'write frequency shifts and mass lockers
         For j = MinFN To MaxFN
           'we do not have neccesserilly all them in gel
           FNIndex = GetDFIndex(CallerID, j)
           If FNIndex > 0 Then
              .ScanInfo(FNIndex).FrequencyShift = FNS(j)
              If FNLM(j) > 0 Then
                 Select Case MWType(j)
                 Case glCSType   'will not happen but
                   .CSData(MWID(MWInd(FNLM(j)))).MTID = glMASS_LOCKER_MARK & glARG_SEP
                 Case glIsoType
                   .IsoData(MWID(MWInd(FNLM(j)))).MTID = glMASS_LOCKER_MARK & glARG_SEP
                 End Select
              End If
           End If
         Next j
         .Comment = .Comment & vbCrLf & Now() & vbCrLf & "Lock mass function applied."
       End With
       GelStatus(CallerID).Dirty = True
       lblStatus.Caption = ""
       Me.MousePointer = vbDefault
    Else
       Exit Sub
    End If
  Case glLM_SAVE_NEW
  End Select
End If
Unload Me
End Sub

Private Sub cmdLock_Click()
Dim LastOKScan As Integer
Dim LockDone As Boolean
Dim Respond
If Selected Then
   LastOKScan = MinFN - 1
'also mark that we at least attempted to lock masses
   Locked = True
   Do Until LockDone
      LastOKScan = LockMass(LastOKScan + 1)
      If LastOKScan = MaxFN Then
         LockDone = True
      Else
         If bCancel Then
            lblStatus.Caption = ""
            bCancel = False
            Exit Sub
         End If
         FNLM(LastOKScan + 1) = FN_LM_ERROR
         Respond = MsgBox("Error locking masses in scan " & LastOKScan + 1 & " in scan range [" & MinFN & "," & MaxFN & "]. Continue with next scan?", vbYesNo)
         If Respond = vbYes Then
            LastOKScan = LastOKScan + 1
            If LastOKScan = MaxFN Then LockDone = True
         Else
            LockDone = True
         End If
      End If
    Loop
Else
   MsgBox "Lock masses have to be selected first. Use Select Lockers function.", vbOKOnly
End If
End Sub

Private Sub cmdReport_Click()
Dim FileNum As Integer
Dim FileNam As String
Dim sLine As String
Dim i As Long
On Error Resume Next

If Locked Then
  If MinFN <= MaxFN Then
     Me.MousePointer = vbHourglass
     lblStatus.Caption = "Generating report..."
     FileNum = FreeFile
     FileNam = GetTempFolder() & RawDataTmpFile
     Open FileNam For Output As FileNum
     'print gel file name and Search definition as reference
     Print #FileNum, "Gel File: " & GelBody(CallerID).Caption
     Print #FileNum, "MT Database: " & GelData(CallerID).PathtoDatabase
     Print #FileNum, "Lock mass un Unique Mass Classes"
     Print #FileNum, GetUMCDefDesc(GelUMC(CallerID).def)
     sLine = "Scan #" & vbTab & "Count" & vbTab & "MT" _
           & vbTab & "  Lock MW  " & vbTab & "Freq.Shift"
     Print #FileNum, sLine
     For i = MinFN To MaxFN
       sLine = i & vbTab & FNCnt(i) & vbTab
       Select Case FNLM(i)
       Case FN_LM_NOCANDIDATES
         sLine = sLine & "No lock mass candidates found in this scan."
       Case FN_LM_ERROR
         sLine = sLine & "Error locking masses in this scan."
       Case FN_LM_NOATTEMPT
         sLine = sLine & "No lock mass attempt made in this scan."
       Case Else
         sLine = sLine & FNAMT(i) & vbTab & Format(MWLM(MWInd(FNLM(i))), "0.00000") _
               & vbTab & Format$(FNS(i), "0.00000")
       End Select
       Print #FileNum, sLine
     Next i
     Close FileNum
     lblStatus.Caption = ""
     frmDataInfo.Tag = "AMTLM"
     frmDataInfo.Show vbModal
  Else
     MsgBox "No scans found.", vbOKOnly
  End If
  Me.MousePointer = vbDefault
Else
  MsgBox "No mass locking applied yet.", vbOKCancel, "2DGelLand - Lock Mass"
End If
End Sub

Private Sub cmdSelectLockers_Click()
Dim i As Long
Dim j As Long
Dim BestScore As Double
Dim BestScoreInd As Long
Dim ScanLocker As Long
Dim BestScanLocker As Long
Dim Resp

Resp = MsgBox("Before using this function, gel needs to be broken to the Unique Mass Classes and searched against the MT database. Continue with lock mass selection?", vbYesNo)
If Resp <> vbYes Then Exit Sub
If FillArrays() Then
   Me.MousePointer = vbHourglass
   lblStatus.Caption = "Calculating LC-MS Feature parameters..."
   DoEvents
   ClsCnt = UMCStatistics1(CallerID, ClsStat())
   If GelUMC(CallerID).UMCCnt > 0 Then
     'pick best AMT choice, and score for each class
     lblStatus.Caption = "Selecting MTs representing LC-MS Features..."
     DoEvents
     ReDim ClsAMT(GelUMC(CallerID).UMCCnt - 1)
     ReDim ClsHits(GelUMC(CallerID).UMCCnt - 1)
     ReDim ClsScore(GelUMC(CallerID).UMCCnt - 1)
     For i = 0 To GelUMC(CallerID).UMCCnt - 1
       lblStatus.Caption = "Class " & (i + 1) & "/" & GelUMC(CallerID).UMCCnt
       DoEvents
       UMCBestAMT i, ClsAMT(i), ClsHits(i)
       ClsScore(i) = UMCScore(i)
     Next i
     'pick the best class for each scan
     lblStatus.Caption = "Selecting the best LC-MS Feature for each scan..."
     DoEvents
     For i = MinFN To MaxFN
       BestScore = 0
       BestScoreInd = -1
       BestScanLocker = -1
       For j = 0 To GelUMC(CallerID).UMCCnt - 1
         If (GelUMC(CallerID).UMCs(j).MinScan <= i) And (i <= GelUMC(CallerID).UMCs(j).MaxScan) _
            And (Len(ClsAMT(j)) > 0) Then
            'class covers this scan and it is also AMT hit; because
            'Charge State data can not be "mass lockers" we need to check
            'that there is a Isotopic point in this class in this scan
            If ClsScore(j) > BestScore Then
               ScanLocker = UMCScanLocker(CallerID, j, i)
               If ScanLocker > 0 Then
                  BestScoreInd = j
                  BestScore = ClsScore(j)
                  BestScanLocker = ScanLocker
               End If
            End If
         End If
       Next j
       FNCls(i) = BestScoreInd
       FNClsLM(i) = BestScanLocker
     Next i
     'pick the best "lock mass" for each scan(from the best class for that scan)
     lblStatus.Caption = "Selecting lock masses..."
     DoEvents
     For i = MinFN To MaxFN
       If FNCls(i) > 0 Then
          FNLM(i) = ScanLockerMWInd(i, FNClsLM(i))
          FNAMT(i) = ClsAMT(FNCls(i))
       Else
          FNLM(i) = FN_LM_NOCANDIDATES
       End If
     Next i
     'put lock masses to the gel selection(Isotopic)
     With GelBody(CallerID).GelSel
         For i = MinFN To MaxFN
             If FNLM(i) > 0 Then
                .AddToIsoSelection MWID(MWInd(FNLM(i)))
             End If
         Next i
     End With
     lblStatus.Caption = ""
     Selected = True
     Me.MousePointer = vbDefault
   Else
     MsgBox "No LC-MS Feature found; lock mass selection failed.", vbOKOnly, glFGTU
   End If
Else
   MsgBox "Error initializing data structures. Can not perform lock mass function.", vbOKOnly, glFGTU
End If
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
If GelUMC(CallerID).def.MWField > 0 Then
   With GelData(CallerID)
        Select Case UCase(.CalEquation)
        Case UCase(CAL_EQUATION_1)
            Set Cal = New CalEq1
            Cal.A = .CalArg(1)
            Cal.B = .CalArg(2)
            If Cal.EquationOK Then
               EnableCommands True
            Else
               MsgBox "Invalid parameters in calibration equation.", vbOKOnly
               EnableCommands False
            End If
        Case UCase(CAL_EQUATION_4)
            Set Cal = New CalEq4
            Cal.A = .CalArg(1)
            Cal.B = .CalArg(2)
            If Cal.EquationOK Then
               EnableCommands True
            Else
               MsgBox "Invalid parameters in calibration equation.", vbOKOnly
               EnableCommands False
            End If
        Case UCase(CAL_EQUATION_5)
            Set Cal = New CalEq5
            Cal.A = .CalArg(1)
            Cal.B = .CalArg(2)
            If Not Cal.EquationOK Then
               lblStatus.Caption = "Invalid parameters in calibration equation."
               Exit Sub
            End If
        Case UCase(CAL_EQUATION_2), UCase(CAL_EQUATION_3)
            If .CalArg(3) = 0 Then    'same case as CAL_EQUATION_1
               Set Cal = New CalEq1
               Cal.A = .CalArg(1)
               Cal.B = .CalArg(2)
               If Not Cal.EquationOK Then
                  lblStatus.Caption = "Invalid parameters in calibration equation."
                  Exit Sub
               End If
            Else
               lblStatus.Caption = "Calibration equation form not implemented."
               Exit Sub
            End If
        Case Else     'later we might use other calibration equations
            MsgBox "Calibration equation not found. Lock mass function not possible.", vbOKOnly
            EnableCommands False
        End Select
        'load all scans
        MinFN = .ScanInfo(1).ScanNumber
        MaxFN = .ScanInfo(UBound(.ScanInfo)).ScanNumber
   End With
Else
   MsgBox "No Unique Mass Classes found. Make sure that UNC count is performed at least once.", vbOKOnly
End If
End Sub

Private Sub Form_Load()
If IsWinLoaded(TrackerCaption) Then Unload frmTracker
Load frmSearchAMT  'this will load AMT if not already loaded
Unload frmSearchAMT
If AMTCnt > 0 Then  'everything OK
   EnableCommands True
Else                'something wrong; disable functions
   EnableCommands False
End If
txtMassTag.Text = ""
optLMResults(amtlmDef.lmSaveResults).Value = True
End Sub

Private Sub optLMResults_Click(Index As Integer)
'this is used both here and on AMTLockMass
amtlmDef.lmSaveResults = Index
End Sub

Private Function FillArrays() As Boolean
'returns True if data structures are successfully
'initialized and filled; False on any error
Dim CurrFN As Long, PrevFN As Long
Dim i As Long
Dim FNSorter As New QSInteger

MWCnt = 0
lblStatus.Caption = "Loading data structures..."
DoEvents
With GelData(CallerID)
    If .CSLines + .IsoLines > 0 Then
       ReDim MWID(1 To .CSLines + .IsoLines)
       ReDim MWFN(1 To .CSLines + .IsoLines)
       ReDim MWLM(1 To .CSLines + .IsoLines)
       ReDim MWType(1 To .CSLines + .IsoLines)
       ReDim MWInd(1 To .CSLines + .IsoLines)
    Else
       lblStatus.Caption = ""
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
        ReDim FNAMT(MinFN To MaxFN)
        ReDim FNCnt(MinFN To MaxFN)
        ReDim FNIndMin(MinFN To MaxFN)
        ReDim FNIndMax(MinFN To MaxFN)
        ReDim FNCls(MinFN To MaxFN)
        ReDim FNClsLM(MinFN To MaxFN)
    Else
        lblStatus.Caption = ""
        Exit Function
    End If
End With
'order MW on FN and count original data
lblStatus.Caption = "Analyzing original data..."
DoEvents
If Not FNSorter.QSAsc(MWFN(), MWInd()) Then
   lblStatus.Caption = ""
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
lblStatus.Caption = ""
DoEvents
FillArrays = True
End Function

Private Function LockMass(ByVal Scan As Integer) As Integer
'returns last scan with successfully locked masses (note that
'successful includes lock mass not performed because of no candidates)
Dim i As Integer
Dim j As Long
Dim LM As Double
Dim MOverZ As Double
Dim Freq As Double
Dim FreqAMT As Double
Dim CS As Double
On Error GoTo exit_LockMass

With GelData(CallerID)
  Me.MousePointer = vbHourglass
  For i = Scan To MaxFN
    lblStatus.Caption = "Locking masses in scan " & i & " [" & MinFN & "," & MaxFN & "]"
    DoEvents
    If FNIndMin(i) <= FNIndMax(i) And FNIndMax(i) > 0 Then
       If FNLM(i) > 0 Then     'lock masses
          Select Case MWType(MWInd(FNLM(i)))
          Case glCSType
          'Charge State data can not be lock mass???
          Case glIsoType
              CS = .IsoData(MWID(MWInd(FNLM(i)))).Charge
              If CS > 0 Then
                'calculate frequency for experimental mw
                MOverZ = MWLM(MWInd(FNLM(i))) / CS + glMASS_CC
                Freq = Cal.CyclotFreq(MOverZ)
                'calculate frequency for theoretical mw
                LM = GetAMTMWByID(FNAMT(i)) + MassTag
                If LM > 0 Then
                   MOverZ = LM / CS + glMASS_CC
                   FreqAMT = Cal.CyclotFreq(MOverZ)
                   'we have now frequency shift
                   FNS(i) = FreqAMT - Freq
                   'recalculate all other masses
                   For j = FNIndMin(i) To FNIndMax(i)
                      Select Case MWType(MWInd(j))
                      Case glCSType
                      'Charge State data can not be lock mass and is basically
                      'average over multiple charge states but can be
                      'recalculated(shifted) based on any charge state
                        CS = 1
                        If CS > 0 Then
                           MOverZ = MWLM(MWInd(j)) / CS + glMASS_CC
                           Freq = Cal.CyclotFreq(MOverZ) + FNS(i)
                           MOverZ = Cal.MOverZ(Freq)
                           MWLM(MWInd(j)) = CS * (MOverZ - glMASS_CC)
                        End If
                      Case glIsoType
                        CS = .IsoData(MWID(MWInd(j))).Charge
                        If CS > 0 Then
                           MOverZ = MWLM(MWInd(j)) / CS + glMASS_CC
                           Freq = Cal.CyclotFreq(MOverZ) + FNS(i)
                           MOverZ = Cal.MOverZ(Freq)
                           MWLM(MWInd(j)) = CS * (MOverZ - glMASS_CC)
                        End If
                      End Select
                   Next j
                End If
              End If
          End Select
       End If
    Else
       FNLM(i) = FN_LM_NOCANDIDATES
    End If
    If bCancel Then Exit For
  Next i
End With

exit_LockMass:
LockMass = i - 1
lblStatus.Caption = ""
Me.MousePointer = vbDefault
End Function

Public Function GetAMTMWByID(ByVal ID As String) As Double
'returns AMT MW for AMT; not optimized to tell the truth
Dim i As Long
GetAMTMWByID = -1
For i = 1 To UBound(AMTData)
    If AMTData(i).ID = ID Then
       GetAMTMWByID = AMTData(i).MW
       Exit Function
    End If
Next i
End Function

Private Sub optScore_Click(Index As Integer)
ScoreOption = Index
End Sub

Private Sub UMCBestAMT(ByVal ClsInd As Long, _
                       ByRef ID As String, _
                       ByRef Hits As Long)
'returns in ID best AMT for class with index ClsInd
'or empty string for any error or if such AMT does not exists
'it returns in Hits number of hits for best AMT
Dim AMTs() As String
Dim AMTCnt
Dim Cnt As Long
Dim UnqAMTCnt As Long
Dim Identity As String
Dim IsInList As Boolean
Dim AMTList() As String
Dim AMTErr() As Double
Dim AMTHits() As Long
Dim sAMT As String
Dim sErr As String
Dim BestError As Double
Dim MostHits As Long
Dim BestAMT As String
Dim i As Long, j As Long, k As Long
On Error GoTo exit_UMCBestAMT

ID = ""
Hits = 0
UnqAMTCnt = 0
Cnt = 0
With GelUMC(CallerID).UMCs(ClsInd)
   For i = 0 To .ClassCount - 1     'for each class member extract all AMT listed
     Identity = ""
     Select Case .ClassMType(i)
     Case glCSType
        Identity = GelData(CallerID).CSData(.ClassMInd(i)).MTID
     Case glIsoType
        Identity = GelData(CallerID).IsoData(.ClassMInd(i)).MTID
     End Select
     
     AMTCnt = GetAMTRefFromString2(Identity, AMTs())
     If AMTCnt > 0 Then
        UnqAMTCnt = UnqAMTCnt + AMTCnt
        ReDim Preserve AMTList(1 To UnqAMTCnt)
        ReDim Preserve AMTErr(1 To UnqAMTCnt)
        ReDim Preserve AMTHits(1 To UnqAMTCnt)
        For j = 1 To AMTCnt
            IsInList = False
            If Cnt > 0 Then   'First is always new
               sAMT = GetIDFromString(AMTs(j), AMTMark, AMTIDEnd)
               sErr = GetMWErrFromString(AMTs(j))
               If IsNumeric(sErr) Then
                  For k = 1 To Cnt
                      If sAMT = AMTList(k) Then
                         AMTHits(k) = AMTHits(k) + 1
                         AMTErr(k) = AMTErr(k) + Abs(CDbl(sErr))
                         IsInList = True
                      End If
                   Next k
               Else
                  IsInList = True   'it is not but it is error so just ignore
               End If
            End If
            If Not IsInList Then
               Cnt = Cnt + 1
               sAMT = GetIDFromString(AMTs(j), AMTMark, AMTIDEnd)
               sErr = GetMWErrFromString(AMTs(j))
               If IsNumeric(sErr) Then
                  AMTList(Cnt) = sAMT
                  AMTErr(Cnt) = Abs(CDbl(sErr))
                  AMTHits(Cnt) = 1
               End If
            End If
        Next j
        UnqAMTCnt = Cnt
     End If
   Next i
End With
'now go and pick the best
If UnqAMTCnt > 0 Then
   ReDim Preserve AMTList(1 To UnqAMTCnt)
   ReDim Preserve AMTErr(1 To UnqAMTCnt)
   ReDim Preserve AMTHits(1 To UnqAMTCnt)
   MostHits = 0
   BestError = glHugeOverExp
   BestAMT = ""
   For i = 1 To UnqAMTCnt
       If (AMTHits(i) >= MostHits) And (AMTHits(i) > 0) Then
          If AMTHits(i) > MostHits Then      'new most hits
             AMTErr(i) = AMTErr(i) / AMTHits(i)
             BestError = AMTErr(i)
             MostHits = AMTHits(i)
             BestAMT = AMTList(i)
          Else    'same number of hits; check the error
             AMTErr(i) = AMTErr(i) / AMTHits(i)
             If AMTErr(i) < BestError Then
                BestError = AMTErr(i)
                MostHits = AMTHits(i)
                BestAMT = AMTList(i)
             End If
          End If
       End If
   Next i
   ID = BestAMT
   Hits = MostHits
End If

exit_UMCBestAMT:
End Sub

Private Function UMCScore(ByVal ClsInd As Long) As Long
'returns score best for class with index ClsInd
'all scores are ascending meaning more is better
On Error GoTo exit_UMCScore
UMCScore = -1
Select Case ScoreOption
Case 0  'intensity
     UMCScore = ClsStat(ClsInd, ustClassIntensity)           'average class intensity
Case 1  'fit
     If ClsStat(ClsInd, ustFitAverage) > 0 Then
        UMCScore = 1 / ClsStat(ClsInd, ustFitAverage)    'inverse of average class fit
     Else
        UMCScore = glHugeOverExp
     End If
Case 2  'AMT fit
Case 3  'number of AMT hits
     UMCScore = ClsHits(ClsInd)
End Select

exit_UMCScore:
End Function

Private Function ScanLockerMWInd(ByVal FN As Integer, _
                                 ByVal LockerInd As Long) As Long
'returns index of best scan locker for scan FN in the MWInd array
'LockerInd is index in GelData().IsoData array
Dim i As Long
For i = FNIndMin(FN) To FNIndMax(FN)
    If MWType(MWInd(i)) = glIsoType Then
       If MWID(MWInd(i)) = LockerInd Then
          ScanLockerMWInd = i
          Exit Function
       End If
    End If
Next i
ScanLockerMWInd = FN_LM_ERROR
End Function

Private Sub EnableCommands(ByVal Arg As Boolean)
cmdSelectLockers.Enabled = Arg
cmdLock.Enabled = Arg
cmdReport.Enabled = Arg
cmdCancel.Enabled = Arg
End Sub

Private Sub txtMassTag_LostFocus()
If IsNumeric(txtMassTag.Text) Then
   MassTag = CDbl(txtMassTag.Text)
Else
   MassTag = 0
   txtMassTag.Text = ""
End If
End Sub

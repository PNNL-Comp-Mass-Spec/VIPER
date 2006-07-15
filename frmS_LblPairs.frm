VERSION 5.00
Begin VB.Form frmS_LblPairs 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Individual Labeled Pairing Analysis"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4785
   Icon            =   "frmS_LblPairs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMaxLblDiff 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   19
      Text            =   "1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtDelta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   17
      Text            =   "8.05"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtERMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Text            =   "4"
      Top             =   2745
      Width           =   855
   End
   Begin VB.TextBox txtERMin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Text            =   "0.25"
      Top             =   2745
      Width           =   855
   End
   Begin VB.TextBox txtMinLbl 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   "1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtScanTol 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      Text            =   "1"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtMaxLbl 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Text            =   "5"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtPairTol 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "0.02"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtLabel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "442.2249697"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max. difference in number of labels in Lt/Hv:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   120
      X2              =   4680
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Heavy/Light Delta :"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   16
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ER Inclusion Range:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   14
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Labels:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Tolerance:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   7
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max Labels:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pair MW Tol.:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label (Lt.):"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmS_LblPairs.frx":030A
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "&Function"
      Begin VB.Menu mnuFFindPairs 
         Caption         =   "&Find Pairs"
      End
      Begin VB.Menu mnuFMarkAmbPairs 
         Caption         =   "Mark &Ambiguous Pairs"
      End
      Begin VB.Menu mnuFMarkBadERPairs 
         Caption         =   "&Mark Pairs Out Of ER Range"
      End
      Begin VB.Menu mnuFClearAllPairs 
         Caption         =   "Clear All &Pairs"
      End
      Begin VB.Menu mnuFDelExcPairs 
         Caption         =   "Delete &Excluded Pairs"
      End
      Begin VB.Menu mnuFDelER 
         Caption         =   "&Clear Pairs ER"
      End
      Begin VB.Menu mnuFERRecalculation 
         Caption         =   "&Recalculate ER"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuRPairsAll 
         Caption         =   "Pairs &All"
      End
      Begin VB.Menu mnuRPairsIncluded 
         Caption         =   "Pairs I&ncluded Only"
      End
      Begin VB.Menu mnuRPairsExcluded 
         Caption         =   "Pairs &Excluded Only"
      End
      Begin VB.Menu mnuRERStat 
         Caption         =   "ER &Statistics (Text)"
      End
      Begin VB.Menu mnuRERStatGraph 
         Caption         =   "ER Statistics (Graph)"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmS_LblPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 05/28/2002 nt
'----------------------------------------------------------
'This form is derived from frmUMCLblPairs
'----------------------------------------------------------
Option Explicit

Const MAXPAIRS As Long = 10000000

Dim CallerID As Long
Dim bLoading As Boolean

Dim Label As Double
Dim Delta As Double
Dim PairTol As Double
Dim MinLbl As Long
Dim MaxLbl As Long
Dim MaxLblDiff As Long      'maximum difference between
                            'number of labels in l/h member
                            'if 0 number of labels must be the same
Dim ScanTol As Long
Dim ERMin As Double
Dim ERMax As Double


Dim MWCnt As Long
Dim MW() As Double        'parallel with Iso array in GelData
Dim MWScan() As Long


'ER statistic depends on type of ER calculation but it always uses 1000 bins
'for ratio                  nonequidistant nodes from 0 to 50
'for logarithmic ratio      equidistant nodes from -50 to 50
'for symmetric ratio         equidistant nodes from -50 to 50
Dim ERBin() As Double       'ER nodes
Dim ERBinAll() As Long      'bin count - all pairs
Dim ERBinInc() As Long      'bin count - included pairs
Dim ERBinExc() As Long      'bin count - excluded pairs
Dim ERAllS As ERStatHelper
Dim ERIncS As ERStatHelper
Dim ERExcS As ERStatHelper


Private Sub Form_Activate()
If bLoading Then
   bLoading = False
   CallerID = Me.Tag
   UpdateStatus "Loading data structures ..."
   Call LoadData
   UpdateStatus ""
End If
End Sub


Private Sub Form_Load()
'set defaults
bLoading = True
Label = txtLabel.Text
Delta = txtDelta.Text
PairTol = txtPairTol.Text
MinLbl = txtMinLbl.Text
MaxLbl = txtMaxLbl.Text
MaxLblDiff = txtMaxLblDiff.Text
ScanTol = txtScanTol.Text
ERMin = txtERMin.Text
ERMax = txtERMax.Text
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim eResponse As VbMsgBoxResult
If GelP_D_L(CallerID).PCnt > 0 Then
    eResponse = MsgBox("Fill comparative display structures with calculated expression ratios?", vbYesNoCancel, glFGTU)
    Select Case eResponse
    Case vbYes
        Me.MousePointer = vbHourglass
        UpdateStatus "Filling comparative display structures..."
        Call FillSolo_ERs(CallerID)
        GelBody(CallerID).ResetGraph True, False, GelBody(CallerID).fgDisplay
        Me.MousePointer = vbDefault
    Case vbNo           'do nothing, just continue unload
    Case vbCancel
        Cancel = True
    End Select
End If
End Sub

Private Sub mnuFClearAllPairs_Click()
ClearAllPairs
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFDelER_Click()
'-----------------------------------------------------
'this will resdimension ER structure and leave type as
'glER_NONE as warning that calculation is not done
'-----------------------------------------------------
InitDltLblPairsER CallerID
End Sub

Private Sub mnuFDelExcPairs_Click()
UpdateStatus "Deleting excluded pairs ..."
Me.MousePointer = vbHourglass
UpdateStatus DeleteExcludedPairs(CallerID)
Me.MousePointer = vbDefault
End Sub

Private Sub mnuFERRecalculation_Click()
CalcDltLblPairsER_Solo CallerID
End Sub

Private Sub mnuFFindPairs_Click()
'--------------------------------
'find all potential pairs
'--------------------------------
Dim Respond
On Error GoTo exit_cmdFindPairs

If GelP_D_L(CallerID).PCnt > 0 Then
'something is already in pairs structure; give user chance to change their mind
   Respond = MsgBox("Pairs structure already contains pairs. Selected procedure will clear all existing pairs. Continue?", vbOKCancel, glFGTU)
   If Respond <> vbOK Then Exit Sub
   ClearAllPairs
End If
Me.MousePointer = vbHourglass
If MWCnt > 0 Then
   UpdateStatus "Finding pair classes ..."
   FindPairs
Else
   MsgBox "No isotopic peaks found.", vbOKOnly, glFGTU
End If

exit_cmdFindPairs:
UpdateStatus ""
Me.MousePointer = vbDefault
End Sub

Private Sub mnuFMarkAmbPairs_Click()
Dim strMessage As String
strMessage = PairsSearchMarkAmbiguous(Me, CallerID, False)
UpdateStatus strMessage
End Sub

Private Sub mnuFMarkBadERPairs_Click()
'-----------------------------------------------------------------
'mark pair as bad if expression ratio is out of ER inclusion range
'-----------------------------------------------------------------
Dim strMessage As String
strMessage = PairsSearchMarkBadER(ERMin, ERMax, CallerID, False)
UpdateStatus strMessage
End Sub

Private Sub mnuFunction_Click()
Call PickParameters
End Sub

Private Sub mnuReport_Click()
Call PickParameters
End Sub

Private Sub mnuRERStat_Click()
UpdateStatus "Generating report ..."
Me.MousePointer = vbHourglass
If GenerateERStat Then
   ReportERStat CallerID, ERBin(), ERBinAll(), ERBinInc(), _
                ERBinExc(), ERAllS, ERIncS, ERExcS
End If
Me.MousePointer = vbDefault
UpdateStatus ""
End Sub

Private Sub mnuRPairsAll_Click()
UpdateStatus "Generating report ..."
Me.MousePointer = vbHourglass
ReportDltLblPairs_S CallerID, 0
Me.MousePointer = vbDefault
UpdateStatus ""
End Sub


Private Sub mnuRPairsExcluded_Click()
UpdateStatus "Generating report ..."
Me.MousePointer = vbHourglass
ReportDltLblPairs_S CallerID, glPAIR_Exc
Me.MousePointer = vbDefault
UpdateStatus ""
End Sub

Private Sub mnuRPairsIncluded_Click()
UpdateStatus "Generating report ..."
Me.MousePointer = vbHourglass
ReportDltLblPairs_S CallerID, glPAIR_Inc
Me.MousePointer = vbDefault
UpdateStatus ""
End Sub


Private Sub txtDelta_LostFocus()
On Error GoTo err_Delta
Delta = CDbl(txtDelta.Text)
If Delta > 0 Then Exit Sub
err_Delta:
MsgBox "Delta should be positive number.", vbOKOnly, glFGTU
txtDelta.SetFocus
End Sub

Private Sub txtERMax_LostFocus()
On Error GoTo err_ERMax
ERMax = CDbl(txtERMax.Text)
Exit Sub
err_ERMax:
MsgBox "Maximum of ER range should be a number.", vbOKOnly, glFGTU
txtERMax.SetFocus
End Sub

Private Sub txtERMin_LostFocus()
On Error GoTo err_ERMin
ERMin = CDbl(txtERMin.Text)
Exit Sub
err_ERMin:
MsgBox "Minimum of ER range should be a number.", vbOKOnly, glFGTU
txtERMin.SetFocus
End Sub


Private Sub txtLabel_LostFocus()
On Error GoTo err_Label
Label = CDbl(txtLabel.Text)
If Label >= 0 Then Exit Sub
err_Label:
MsgBox "Label mass should be non-negative number.", vbOKOnly, glFGTU
txtLabel.SetFocus
End Sub

Private Sub txtMaxLbl_LostFocus()
On Error GoTo err_MaxLbl
MaxLbl = CLng(txtMaxLbl.Text)
If MaxLbl > 0 Then Exit Sub
err_MaxLbl:
MsgBox "Maximum number of labels should be positive integer.", vbOKOnly, glFGTU
txtMaxLbl.SetFocus
End Sub

Private Sub txtMaxLblDiff_LostFocus()
On Error GoTo err_MaxLblDiff
MaxLblDiff = CLng(txtMaxLblDiff.Text)
If MaxLblDiff >= 0 And MaxLblDiff <= MaxLbl Then Exit Sub
err_MaxLblDiff:
MsgBox "Difference in labels count for light and heavy pair memebers shouldn't exceed maximum number of labels.", vbOKOnly, glFGTU
txtMaxLblDiff.SetFocus
End Sub

Private Sub txtMinLbl_LostFocus()
On Error GoTo err_MinLbl
MinLbl = CLng(txtMinLbl.Text)
If MinLbl > 0 Then Exit Sub
err_MinLbl:
MsgBox "Minimum number of labels should be positive integer.", vbOKOnly, glFGTU
txtMinLbl.SetFocus
End Sub

Private Sub txtPairTol_LostFocus()
On Error GoTo err_PairTol
PairTol = CDbl(txtPairTol.Text)
If PairTol > 0 Then Exit Sub
err_PairTol:
MsgBox "Pairs tolerance should be positive number.", vbOKOnly, glFGTU
txtPairTol.SetFocus
End Sub

Private Sub txtScanTol_LostFocus()
On Error GoTo err_ScanTol
ScanTol = CLng(txtScanTol.Text)
If ScanTol >= 0 Then Exit Sub
err_ScanTol:
MsgBox "Scan tolerance should be non-negative integer.", vbOKOnly, glFGTU
txtScanTol.SetFocus
End Sub

Private Sub UpdateStatus(ByVal Status As String)
'-----------------------------------------------
'set status label; entertain user so it doesn't
'freak out before function finishes
'-----------------------------------------------
lblStatus.Caption = Status
DoEvents
End Sub

Private Sub ClearAllPairs()
    DestroyDltLblPairs CallerID
End Sub

' Unused Procedure (February 2005)
''Private Sub CleanPairsERs()
'''-------------------------------------------
'''this function resets ER in Pairs structure;
'''underlying gel does not change
'''-------------------------------------------
''Dim i As Long
''With GelP_D_L(CallerID)
''    For i = 0 To .PCnt - 1
''        With .Pairs(i)
''            .ER = ER_CALC_ERR
''            .ERStDev = 0
''            .ERChargeStateBasisCount = 0
''            ReDim .ERChargesUsed(0)
''            .ERMemberBasisCount = 0
''        End With
''    Next i
''End With
''End Sub


Private Sub FindPairs()
'-----------------------------------------------------
'actual pairing function; finds and put into structure
'all potential pairs based on numerical criteria
'-----------------------------------------------------
Dim i As Long, j As Long, k As Long, L As Long
Dim MWDiff As Double
On Error GoTo err_FindPairs

If InitDltLblPairs(CallerID) Then     'this will reserve 40000 pairs to start with
  With GelP_D_L(CallerID)
    .DltLblType = ptS_Lbl
    .PCnt = 0
    .SearchDef.DeltaMass = 0
    .SearchDef.LightLabelMass = Label
    '----------------------------------------------------------
    'this is a little wicked situation in which heavy and light
    'members do not have to have same number of labels attached
    'and that can cause light member to be heavier than heavy
    '----------------------------------------------------------
    For i = 1 To MWCnt
      UpdateStatus "Peak: " & (i) & " MW: " & Format$(MW(i), ".0000")
      For j = 1 To MWCnt            'light/heavy here is asymmetric so we have to try all
                                    'combinations ((i,j) and (j,i) could be pairs with
                                    'different combination of label numbers)
        If i <> j Then              'not interested in trivial pairs
           MWDiff = MW(j) - MW(i)
           If MaxLblDiff > 0 Then       'label count could differ
              For k = MinLbl To MaxLbl
                  If (MW(j) - k * (Label + Delta) > 0) Then   'don't consider impossible pairs
                     For L = MinLbl To MaxLbl
                         If (MW(i) - L * Label > 0) Then     'don't consider impossible pairs
                            If (Abs(k - L) <= MaxLblDiff And (k + L) > 0) Then
                               If Abs(MWDiff - ((k - L) * Label + k * Delta)) <= PairTol Then
                               'pair if scan condition OK
                                  If ((Abs(MWScan(j) - MWScan(i)) <= ScanTol)) Then
                                     .PCnt = .PCnt + 1
                                     With .Pairs(.PCnt - 1)
                                        .P1 = i
                                        .P2 = j
                                        .P1LblCnt = L
                                        .P2LblCnt = k
                                     End With
                                  End If
                               End If
                            End If
                         End If
                     Next L
                  End If
              Next k
           ElseIf MaxLblDiff = 0 Then   'label count has to be the same
              For k = MinLbl To MaxLbl
                If (MW(i) - k * Label > 0) Then       'don't consider impossible pairs
                  If (MW(j) - k * (Label + Delta) > 0) Then
                    If Abs(MWDiff - k * Delta) <= PairTol Then
                      'pair if scan conditions OK
                      If ((Abs(MWScan(j) - MWScan(i)) <= ScanTol)) Then
                         .PCnt = .PCnt + 1
                         With .Pairs(.PCnt - 1)
                           .P1 = i
                           .P2 = j
                           .P1LblCnt = k
                           .P2LblCnt = k
                         End With
                      End If
                    End If
                  End If
                End If
              Next k
           End If
         End If
      Next j
    Next i
  End With
  'calculate expression ratios here
  CalcDltLblPairsER_Solo CallerID
    
    'MonroeMod
    AddToAnalysisHistory CallerID, "Searched for Labeled pairs (using individual spectra); Pair Count = " & Trim(GelP_D_L(CallerID).PCnt) & "; Label = " & Trim(Label) & " Da; Heavy/Light Delta = " & Trim(Delta) & " Da; Pair Tolerance = " & Trim(PairTol) & " Da; Scan Tolerance = " & Trim(ScanTol) & "; ER Inclusion Range = " & Trim(ERMin) & " to " & Trim(ERMax)
Else
   MsgBox "Unable to reserve space for pairs structures.", vbOKOnly, glFGTU
End If

exit_FindPairs:
If GelP_D_L(CallerID).PCnt > 0 Then
  Call TrimDltLblPairs(CallerID)
Else
  DestroyDltLblPairs CallerID, False
End If
MsgBox "Total number of unique pairs (uncleaned): " & GelP_D_L(CallerID).PCnt
Exit Sub

err_FindPairs:
Select Case Err.Number
Case 9
'increase array size and resume
  If UBound(GelP_D_L(CallerID).Pairs) > MAXPAIRS Then
     MsgBox "Number of detected pairs too high.(max. number of pairs " & MAXPAIRS & ")", vbOKOnly
     Resume exit_FindPairs
  Else
     If AddDltLblPairs(CallerID, 10000) Then Resume
  End If
Case Else
  MsgBox "Error establishing delta-labeled pairs " & vbCrLf _
    & "Error: " & Err.Number & ", " & Err.Description, vbOKOnly, glFGTU
End Select
End Sub


Private Function GenerateERStat() As Boolean
'-----------------------------------------------------
'do actual ER statistics for all currently included
'and currently excluded pairs
'-----------------------------------------------------
Dim i As Long
Dim BinInd As Long
Dim Done As Boolean
On Error Resume Next

With GelP_D_L(CallerID)
  If .PCnt > 0 Then
     ReDim ERBin(1000)
     ReDim ERBinAll(1000)
     ReDim ERBinInc(1000)
     ReDim ERBinExc(1000)
     ERAllS.ERCnt = 0
     ERAllS.ERBadL = 0
     ERAllS.ERBadR = 0
     ERIncS.ERCnt = 0
     ERIncS.ERBadL = 0
     ERIncS.ERBadR = 0
     ERExcS.ERCnt = 0
     ERExcS.ERBadL = 0
     ERExcS.ERBadR = 0
     
     Select Case GelP_D_L(CallerID).SearchDef.ERCalcType
     Case ectER_RAT                       'cover range from 0 to 50
        ERBin(500) = 1
        For i = 1 To 500
            ERBin(500 + i) = 1 + i * 0.1
            ERBin(500 - i) = 1 / ERBin(500 + i)
        Next i
     Case ectER_LOG                       'cover range from -10 to 10
        ERBin(500) = 0
        For i = 1 To 500
            ERBin(500 - i) = -i * 0.02
            ERBin(500 + i) = i * 0.02
        Next i
     Case ectER_ALT                       'cover range from -50 to 50
        ERBin(500) = 0
        For i = 1 To 500
            ERBin(500 - i) = -i * 0.1
            ERBin(500 + i) = i * 0.1
        Next i
     End Select
        
     For i = 0 To .PCnt - 1
           'find bin for this expression ratio
           BinInd = -1
           Done = False
           Do Until Done
              If .Pairs(i).ER < ERBin(BinInd + 1) Then
                 Done = True
              Else
                 BinInd = BinInd + 1
                 If BinInd >= 1000 Then Done = True
              End If
           Loop
           'add counts
           ERAllS.ERCnt = ERAllS.ERCnt + 1
           If .Pairs(i).STATE = glPAIR_Inc Then
              ERIncS.ERCnt = ERIncS.ERCnt + 1
           ElseIf .Pairs(i).STATE = glPAIR_Exc Then
              ERExcS.ERCnt = ERExcS.ERCnt + 1
           End If
           Select Case BinInd
           Case Is < 0
                ERAllS.ERBadL = ERAllS.ERBadL + 1
                If .Pairs(i).STATE = glPAIR_Inc Then
                    ERIncS.ERBadL = ERIncS.ERBadL + 1
                ElseIf .Pairs(i).STATE = glPAIR_Exc Then
                    ERExcS.ERBadL = ERExcS.ERBadL + 1
                End If
           Case Is > 1000
                ERAllS.ERBadR = ERAllS.ERBadR + 1
                If .Pairs(i).STATE = glPAIR_Inc Then
                    ERIncS.ERBadR = ERIncS.ERBadR + 1
                ElseIf .Pairs(i).STATE = glPAIR_Exc Then
                    ERExcS.ERBadR = ERExcS.ERBadR + 1
                End If
           Case Else            'some of our cases
                ERBinAll(BinInd) = ERBinAll(BinInd) + 1
                If .Pairs(i).STATE = glPAIR_Inc Then
                   ERBinInc(BinInd) = ERBinInc(BinInd) + 1
                ElseIf .Pairs(i).STATE = glPAIR_Exc Then
                   ERBinExc(BinInd) = ERBinExc(BinInd) + 1
                End If
           End Select
       Next i
       GenerateERStat = True
    Else
        MsgBox "No pairs found. Find Pairs function should be used first."
    End If
End With
End Function


Public Sub LoadData()
'--------------------------------------------------
'loads data to temporary structures for ease of use
'--------------------------------------------------
Dim i As Long
On Error GoTo err_LoadData
UpdateStatus "Loading data structures ..."
With GelData(CallerID)
     MWCnt = .IsoLines
     If MWCnt > 0 Then
        ReDim MW(MWCnt)
        ReDim MWScan(MWCnt)
        For i = 1 To .IsoLines
            MW(i) = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
            MWScan(i) = .IsoData(i).ScanNumber
        Next i
     Else
        Erase MW
        Erase MWScan
     End If
End With
UpdateStatus ""
Exit Sub

err_LoadData:
UpdateStatus "Error loading data structures."
End Sub


Private Sub PickParameters()
Call txtDelta_LostFocus
Call txtPairTol_LostFocus
Call txtERMin_LostFocus
Call txtERMax_LostFocus
Call txtLabel_LostFocus
Call txtMaxLbl_LostFocus
Call txtMaxLblDiff_LostFocus
Call txtMinLbl_LostFocus
Call txtScanTol_LostFocus
End Sub


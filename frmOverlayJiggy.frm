VERSION 5.00
Begin VB.Form frmOverlayJiggy 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gettin' Jiggy"
   ClientHeight    =   4080
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAcceptChanges 
      Caption         =   "Accept"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearReport 
      Caption         =   "Clear Report"
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtReport 
      Height          =   3855
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Type"
      Height          =   1695
      Left            =   3480
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
      Begin VB.OptionButton optJiggyType 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Shift"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   19
         ToolTipText     =   "Looking for systematic shift"
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optJiggyType 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Simple LSM"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   17
         ToolTipText     =   "Least Square Method on selected points"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Go Jiggy"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame fraJiggyBase 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Display Selection"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
      Begin VB.OptionButton optJiggyScope 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Use visible overlaid display"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optJiggyScope 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Use all overlaid displays"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbBaseDisplay 
         Height          =   315
         ItemData        =   "frmOverlayJiggy.frx":0000
         Left            =   120
         List            =   "frmOverlayJiggy.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Base display"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraTolerances 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tolerances For Matches In Different Displays"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtAbuTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   7
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Text            =   "0.05"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Text            =   "25"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkUseAbuTol 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Use abundance constraint with tolerance (abundance scale used here is logarithmic)"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1040
         Width           =   3375
      End
      Begin VB.CheckBox chkUseNETTol 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Use NET constraint with tolerance"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   2895
      End
      Begin VB.CheckBox chkUseMWTol 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Use MW constraint with tolerance (ppm)"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "frmOverlayJiggy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'gettin' jiggy function allows fine tunning when overlaying one display over another
'created: 10/08/2002 nt
'last modified: 02/03/2003 nt
'-----------------------------------------------------------------------------------
Option Explicit

Const JIGGY_SCOPE_ALL = 0
Const JIGGY_SCOPE_VISIBLE = 1

Const JIGGY_TYPE_SimpleLSM = 0
Const JIGGY_TYPE_SHIFT = 1

Dim TmpOly() As OverlayStructure         'give user chance to change his/her min

'BaseNET and BaseAbu arrays need to be addresed indirectly with BaseNET(BaseInd(i))
Dim BaseMW() As Double            'MW values in base display
Dim BaseNET() As Double           'NET values in base display
Dim BaseAbu() As Double           'logarithm of abundance values in base display
Dim BaseInd() As Long             'original index in BaseDisplayInd

Dim DispMW() As Double             'data from display that has to be jigged
Dim DispNET() As Double
Dim DispAbu() As Double
Dim DispInd() As Long

Dim MatchCnt As Long            'count of matches
Dim MatchDispInd() As Long       'index of spot in Dis OlyCoo structure
Dim MatchBaseInd() As Long      'index of spot in BaseDisplayInd OlyCoo structure

Dim AdjSlp As Double
Dim AdjInt As Double
Dim AdjAvgDev As Double


'shift is looked within the [0,1] range;
'search starts from specific bin width and increases it size until one bin contains
'more hits than all others combined or it reaches AdjShiftBinWidthMaxValue
'the shift is then declared to be the average of all shifts in the best bin
Dim AdjShift As Double
Dim AdjShiftBinWidth As Double
Dim AdjShiftBinMaxWidth As Double
Dim AdjShiftChangeFactor As Double
Dim AdjShiftBinFreq() As Long
Dim AdjShiftBinAvg() As Double

Dim MWFastSearch As MWUtil

Dim NETShiftBins As New BinDoubles

Private Sub chkUseAbuTol_Click()
OlyJiggyOptions.UseAbuConstraint = (chkUseAbuTol.value = vbChecked)
End Sub

Private Sub chkUseMWTol_Click()
OlyJiggyOptions.UseMWConstraint = (chkUseMWTol.value = vbChecked)
End Sub

Private Sub chkUseNETTol_Click()
OlyJiggyOptions.UseNetConstraint = (chkUseNETTol.value = vbChecked)
End Sub

Private Sub cmbBaseDisplay_Click()
OlyJiggyOptions.BaseDisplayInd = cmbBaseDisplay.ListIndex
End Sub

Private Sub cmdAcceptChanges_Click()
Oly = TmpOly
End Sub

Private Sub cmdClearReport_Click()
txtReport.Text = ""
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDoIt_Click()
'-------------------------------------------------------------------------------
'controls getting jiggy processing
'-------------------------------------------------------------------------------
Dim i As Long
Dim Resp As Long
On Error Resume Next
Me.MousePointer = vbHourglass
If OlyJiggyOptions.BaseDisplayInd >= 0 Then
   UpdateReport "---- NET Adjustment Calculation" & vbCrLf, False
   If PrepareBaseDisplaySearch() Then
      Select Case OlyJiggyOptions.JiggyScope
      Case JIGGY_SCOPE_ALL
        For i = 0 To OlyCnt - 1
            If i <> OlyJiggyOptions.BaseDisplayInd Then
               UpdateReport "Gettin' jiggy on " & Oly(i).Name & vbCrLf, True
               Select Case OlyJiggyOptions.JiggyType
               Case JIGGY_TYPE_SimpleLSM
                 Call GettinJiggy_SimpleLSM(i)
               Case JIGGY_TYPE_SHIFT
                 Call GettinJiggy_Shift(i)
               Case Else
                 UpdateReport "Unrecognized processing type ..." & vbCrLf, True
               End Select
            End If
        Next i
      Case JIGGY_SCOPE_VISIBLE
        For i = 0 To OlyCnt - 1
            If Oly(i).Visible Then
               If i <> OlyJiggyOptions.BaseDisplayInd Then
                  UpdateReport "Gettin' jiggy on " & Oly(i).Name & vbCrLf, True
                  Select Case OlyJiggyOptions.JiggyType
                  Case JIGGY_TYPE_SimpleLSM
                    Call GettinJiggy_SimpleLSM(i)
                  Case JIGGY_TYPE_SHIFT
                    Call GettinJiggy_Shift(i)
                  Case Else
                    UpdateReport "Unrecognized processing type ..." & vbCrLf, True
                  End Select
               End If
            End If
        Next i
      End Select
      Resp = MsgBox("For changes to take effect click on Accept button.", vbOKOnly, glFGTU)
   Else
      MsgBox "Error preparing fast search procedures.", vbOKOnly, glFGTU
   End If
   DestroyStructures
Else
   MsgBox "Select base display and try again.", vbOKOnly, glFGTU
End If
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'load current OlyJiggyOptions settings
Dim i As Long
With OlyJiggyOptions
    If .UseMWConstraint Then
       chkUseMWTol.value = vbChecked
    Else
       chkUseMWTol.value = vbUnchecked
    End If
    If .UseNetConstraint Then
       chkUseNETTol.value = vbChecked
    Else
       chkUseNETTol.value = vbUnchecked
    End If
    If .UseAbuConstraint Then
       chkUseAbuTol.value = vbChecked
    Else
       chkUseAbuTol.value = vbUnchecked
    End If
    txtMWTol.Text = .MWTol
    txtNETTol.Text = .NETTol
    txtAbuTol.Text = .AbuTol
    optJiggyScope(.JiggyScope).value = True
    optJiggyType(.JiggyType).value = True
    For i = 0 To OlyCnt - 1
        cmbBaseDisplay.AddItem Oly(i).Name
    Next i
    If .BaseDisplayInd >= 0 Then cmbBaseDisplay.ListIndex = .BaseDisplayInd
    'create temporary structure that will enable us to undo the changes
    ReDim TmpOly(OlyCnt - 1)
    TmpOly = Oly
End With
End Sub

Private Sub optJiggyScope_Click(Index As Integer)
OlyJiggyOptions.JiggyScope = Index
End Sub

Private Sub optJiggyType_Click(Index As Integer)
OlyJiggyOptions.JiggyType = Index
End Sub

Private Sub txtAbuTol_LostFocus()
On Error Resume Next
If IsNumeric(txtAbuTol.Text) Then
   OlyJiggyOptions.AbuTol = CDbl(txtAbuTol.Text)
Else
   MsgBox "This argument should be numeric value.", vbOKOnly, glFGTU
   txtAbuTol.SetFocus
End If
End Sub

Private Sub txtMWTol_LostFocus()
On Error Resume Next
If IsNumeric(txtMWTol.Text) Then
   OlyJiggyOptions.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "This argument should be numeric value.", vbOKOnly, glFGTU
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNETTol_LostFocus()
On Error Resume Next
If IsNumeric(txtNETTol.Text) Then
   OlyJiggyOptions.NETTol = CDbl(txtNETTol.Text)
Else
   MsgBox "This argument should be numeric value.", vbOKOnly, glFGTU
   txtNETTol.SetFocus
End If
End Sub

Private Sub UpdateReport(ByVal Msg As String, ByVal bAppend As Boolean)
'----------------------------------------------------------------------------
'adds Msg to report text box if bAppend True or replaces it with Msg if False
'----------------------------------------------------------------------------
If bAppend Then
   txtReport.SelStart = Len(txtReport.Text)
   txtReport.SelText = Msg
Else
   txtReport.Text = Msg
End If
DoEvents
End Sub

Public Function GettinJiggy_SimpleLSM(ByVal DInd As Long) As Boolean
'---------------------------------------------------------------------
'find all matches for DInd display in OlyJiggyOptions.BaseDisplayInd and
'optimize linear adjustments based on least square methods
'---------------------------------------------------------------------
Dim i As Long, j As Long
Dim AbsTol As Double
Dim Ind1 As Long, Ind2 As Long
On Error GoTo err_GettinJiggy_SimpleLSM
UpdateReport "Adjusting " & Oly(DInd).Name & "(" & Oly(DInd).DisplayCaption & ")" & vbCrLf, True
MatchCnt = 0
ReDim MatchDispInd(5000)
ReDim MatchBaseInd(5000)
If PrepareDisplayToSearch(DInd) Then
   For i = 0 To UBound(DispInd)
       Ind1 = 0
       Ind2 = -1
       AbsTol = DispMW(i) * OlyJiggyOptions.MWTol * glPPM
       If MWFastSearch.FindIndexRange(DispMW(i), AbsTol, Ind1, Ind2) Then
          If Ind1 <= Ind2 Then
             For j = Ind1 To Ind2
                 If IsJiggyMatch(i, j) Then
                    MatchCnt = MatchCnt + 1
                    MatchDispInd(MatchCnt - 1) = i
                    MatchBaseInd(MatchCnt - 1) = j
                 End If
             Next j
          End If
       End If
   Next i
   If MatchCnt > 1 Then
      ReDim Preserve MatchDispInd(MatchCnt - 1)
      ReDim Preserve MatchBaseInd(MatchCnt - 1)
      UpdateReport "NET adjustment calculation based on " & MatchCnt & " matches." & vbCrLf, True
      If CalcSlopeIntercept() Then
         'calculate new absolute Slope and Intercept
         With TmpOly(DInd)
             UpdateReport "---- Original NET adjustment: " & vbCrLf, True
             UpdateReport "---- ---- Slope: " & .NETSlope & vbCrLf, True
             UpdateReport "---- ---- Intercept: " & .NETIntercept & vbCrLf, True
             UpdateReport "---- ---- Avg.Dev: " & .NETFit & vbCrLf, True
             .NETAdjustment = olyNETDisplay
             .NETDisplayInd = OlyJiggyOptions.BaseDisplayInd
             UpdateReport "---- Relative NET adjustment: " & vbCrLf, True
             UpdateReport "---- ---- Slope: " & AdjSlp & vbCrLf, True
             UpdateReport "---- ---- Intercept: " & AdjInt & vbCrLf, True
             UpdateReport "---- ---- Avg.Dev: " & AdjAvgDev & vbCrLf, True
             .NETAdjustment = olyNETDisplay
             .NETDisplayInd = OlyJiggyOptions.BaseDisplayInd
             .NETSlope = AdjSlp * .NETSlope
             .NETIntercept = AdjSlp * .NETIntercept + AdjInt
             .NETFit = AdjAvgDev
             UpdateReport "---- New absolute NET adjustment: " & vbCrLf, True
             UpdateReport "---- ---- Slope: " & .NETSlope & vbCrLf, True
             UpdateReport "---- ---- Intercept: " & .NETIntercept & vbCrLf, True
             UpdateReport "---- ---- Avg.Dev: " & .NETFit & vbCrLf, True
         End With
      Else
      End If
   Else             'even if there is one match it is not enough to derive adjustment
      Erase MatchDispInd
      Erase MatchBaseInd
      UpdateReport "No matching spots found; adjustment not possible." & vbCrLf, True
   End If
Else
   UpdateReport "Error preparing data for display " & Oly(DInd).DisplayCaption & vbCrLf, True
End If
Exit Function
err_GettinJiggy_SimpleLSM:
Select Case Err.Number
Case 9
   ReDim Preserve MatchDispInd(MatchCnt + 2000)
   ReDim Preserve MatchBaseInd(MatchCnt + 2000)
   Resume
Case Else
   UpdateReport "Error searching for matching spots.", True
End Select
End Function



Public Function GettinJiggy_Shift(ByVal DInd As Long) As Boolean
'---------------------------------------------------------------------
'find all matches for DInd display in OlyJiggyOptions.BaseDisplayInd
'optimize linear adjustments based on systematic shift statistics
'---------------------------------------------------------------------
Dim i As Long, j As Long
Dim AbsTol As Double
Dim Ind1 As Long, Ind2 As Long
Dim MatchNETShift() As Double
Dim CurrHiBinPct As Double            'highest bin percent of total count
Dim Done As Boolean
On Error GoTo err_GettinJiggy_Shift
UpdateReport "Adjusting " & Oly(DInd).Name & "(" & Oly(DInd).DisplayCaption & ")" & vbCrLf, True
MatchCnt = 0
ReDim MatchDispInd(5000)
ReDim MatchBaseInd(5000)
If PrepareDisplayToSearch(DInd) Then
   For i = 0 To UBound(DispInd)
       Ind1 = 0
       Ind2 = -1
       AbsTol = DispMW(i) * OlyJiggyOptions.MWTol * glPPM
       If MWFastSearch.FindIndexRange(DispMW(i), AbsTol, Ind1, Ind2) Then
          If Ind1 <= Ind2 Then
             For j = Ind1 To Ind2
                 If IsJiggyMatch(i, j) Then
                    MatchCnt = MatchCnt + 1
                    MatchDispInd(MatchCnt - 1) = i
                    MatchBaseInd(MatchCnt - 1) = j
                 End If
             Next j
          End If
       End If
   Next i
   On Error Resume Next
   If MatchCnt > 1 Then
      ReDim Preserve MatchDispInd(MatchCnt - 1)
      ReDim Preserve MatchBaseInd(MatchCnt - 1)
      ReDim MatchNETShift(MatchCnt - 1)
      For i = 0 To MatchCnt - 1
          MatchNETShift(i) = BaseNET(BaseInd(MatchBaseInd(i))) - DispNET(MatchDispInd(i))
      Next i
      AdjShiftBinMaxWidth = 0.1
      AdjShiftBinWidth = 0.001
      AdjShiftChangeFactor = 5
      UpdateReport "NET adjustment calculation based on " & MatchCnt & " matches." & vbCrLf, True
      UpdateReport "NET shift count starts with bin width: " & AdjShiftBinWidth & vbCrLf, True
      UpdateReport "NET shift count ends with bin width: " & AdjShiftBinMaxWidth & vbCrLf, True
      UpdateReport "---- Original NET adjustment: " & vbCrLf, True
      UpdateReport "---- ---- Slope: " & TmpOly(DInd).NETSlope & vbCrLf, True
      UpdateReport "---- ---- Intercept: " & TmpOly(DInd).NETIntercept & vbCrLf, True
      UpdateReport "---- ---- Avg.Dev: " & TmpOly(DInd).NETFit & vbCrLf, True
      NETShiftBins.Fill MatchNETShift()         'fill differences into bins object
      NETShiftBins.MinValue = -1
      NETShiftBins.MaxValue = 1
      Do Until Done
         NETShiftBins.BinWidth = AdjShiftBinWidth
         NETShiftBins.CalculateBins
         'display current results
         UpdateReport "---- Top five counts for bin width " & NETShiftBins.BinWidth & vbCrLf, True
         For i = 0 To 4
             UpdateReport "---- ---- " & NETShiftBins.GetBinRange(i) & " " & NETShiftBins.GetBinCount(i) & vbCrLf, True
         Next i
         CurrHiBinPct = NETShiftBins.GetBinPercent(0)  'use returned value to scale increase
                                                       'in shift bin width(more we have in highest
                                                       'bin the less we change in next step)
         'display latest results
         If CurrHiBinPct >= 0.05 Then
            Done = True
         Else
            'maximum increase is ten-folds and then goes down as percentage in highest
            'bin increases; if half of total number is reached it is acceptable to stop
            AdjShiftBinWidth = AdjShiftBinWidth * (1 - CurrHiBinPct) * 10
            If AdjShiftBinWidth > AdjShiftBinMaxWidth Then Done = True
         End If
      Loop
      UpdateReport "---- Calculated NET shift " & NETShiftBins.GetBinAverage(0) & vbCrLf, True
      TmpOly(DInd).NETIntercept = TmpOly(DInd).NETIntercept + NETShiftBins.GetBinAverage(0)
   Else             'even if there is one match it is not enough to derive adjustment
      Erase MatchDispInd
      Erase MatchBaseInd
      UpdateReport "No matching spots found; adjustment not possible." & vbCrLf, True
   End If
Else
   UpdateReport "Error preparing data for display " & Oly(DInd).DisplayCaption & vbCrLf, True
End If
Exit Function
err_GettinJiggy_Shift:
Select Case Err.Number
Case 9
   ReDim Preserve MatchDispInd(MatchCnt + 2000)
   ReDim Preserve MatchBaseInd(MatchCnt + 2000)
   Resume
Case Else
   UpdateReport "Error searching for matching spots.", True
End Select
End Function




Private Function PrepareBaseDisplaySearch() As Boolean
'---------------------------------------------------------------------
'loads structures used for faster search within base display
'---------------------------------------------------------------------
Dim i As Long, Cnt As Long
Dim qsd As New QSDouble
Dim IsoMWField As Integer
Dim ChP() As LaV2DGPoint
IsoMWField = GelData(Oly(OlyJiggyOptions.BaseDisplayInd).DisplayInd).Preferences.IsoDataField
With OlyCoo(OlyJiggyOptions.BaseDisplayInd)
    ReDim BaseMW(.DataCnt - 1)
    ReDim BaseNET(.DataCnt - 1)
    ReDim BaseAbu(.DataCnt - 1)
    ReDim BaseInd(.DataCnt - 1)
End With
With Oly(OlyJiggyOptions.BaseDisplayInd)
    Select Case .Type
    Case olySolo
        For i = 1 To OlyCoo(OlyJiggyOptions.BaseDisplayInd).CSCnt
            Cnt = Cnt + 1
            BaseInd(Cnt - 1) = Cnt - 1
            BaseMW(Cnt - 1) = GelData(.DisplayInd).CSData(i).AverageMW
            BaseNET(Cnt - 1) = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
            BaseAbu(Cnt - 1) = Log(GelData(.DisplayInd).CSData(i).Abundance) / Log(10#)
        Next i
        For i = 1 To OlyCoo(OlyJiggyOptions.BaseDisplayInd).IsoCnt
            Cnt = Cnt + 1
            BaseInd(Cnt - 1) = Cnt - 1
            BaseMW(Cnt - 1) = GetIsoMass(GelData(.DisplayInd).IsoData(i), IsoMWField)
            BaseNET(Cnt - 1) = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
            BaseAbu(Cnt - 1) = Log(GelData(.DisplayInd).IsoData(i).Abundance) / Log(10#)
        Next i
    Case OlyUMC
        For i = 0 To GelUMC(.DisplayInd).UMCCnt - 1
            Cnt = Cnt + 1
            BaseInd(Cnt - 1) = Cnt - 1
            If fUMCCharacteristicPoints(.DisplayInd, i, ChP()) Then
                BaseMW(Cnt - 1) = GelUMC(.DisplayInd).UMCs(i).ClassMW
                BaseNET(Cnt - 1) = .NETSlope * ChP(1).Scan + .NETIntercept
                BaseAbu(Cnt - 1) = GelUMC(.DisplayInd).UMCs(i).ClassAbundance
            End If
        Next i
    End Select
End With
If qsd.QSAsc(BaseMW(), BaseInd()) Then
   Set MWFastSearch = New MWUtil
   If MWFastSearch.Fill(BaseMW()) Then PrepareBaseDisplaySearch = True
End If
End Function
    

Private Function PrepareDisplayToSearch(ByVal DInd As Long) As Boolean
'-----------------------------------------------------------------------
'loads data from display that has to be adjusted; no need to sort here
'-----------------------------------------------------------------------
Dim i As Long, Cnt As Long
Dim IsoMWField As Integer
Dim ChP() As LaV2DGPoint
On Error GoTo exit_PrepareDisplayToSearch
IsoMWField = GelData(Oly(DInd).DisplayInd).Preferences.IsoDataField
With OlyCoo(DInd)
    ReDim DispMW(.DataCnt - 1)
    ReDim DispNET(.DataCnt - 1)
    ReDim DispAbu(.DataCnt - 1)
    ReDim DispInd(.DataCnt - 1)
End With
With Oly(DInd)
   Select Case .Type
   Case olySolo
        For i = 1 To OlyCoo(DInd).CSCnt
            Cnt = Cnt + 1
            DispInd(Cnt - 1) = Cnt - 1
            DispMW(Cnt - 1) = GelData(.DisplayInd).CSData(i).AverageMW
            DispNET(Cnt - 1) = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
            DispAbu(Cnt - 1) = Log(GelData(.DisplayInd).CSData(i).Abundance) / Log(10#)
        Next i
        For i = 1 To OlyCoo(DInd).IsoCnt
            Cnt = Cnt + 1
            DispInd(Cnt - 1) = Cnt - 1
            DispMW(Cnt - 1) = GetIsoMass(GelData(.DisplayInd).IsoData(i), IsoMWField)
            DispNET(Cnt - 1) = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
            DispAbu(Cnt - 1) = Log(GelData(.DisplayInd).IsoData(i).Abundance) / Log(10#)
        Next i
    Case OlyUMC
        For i = 0 To GelUMC(.DisplayInd).UMCCnt - 1
            Cnt = Cnt + 1
            DispInd(Cnt - 1) = Cnt - 1
            If fUMCCharacteristicPoints(.DisplayInd, i, ChP()) Then
               DispMW(Cnt - 1) = GelUMC(.DisplayInd).UMCs(i).ClassMW
               'use class representative for class NET; this probably can be improved by
               'using first, representative and last scan and adjusting it to first,
               'representative and last point of the matching class; however that would
               'not work for adjusting Solo and UMC displays together
               DispNET(Cnt - 1) = .NETSlope * ChP(1).Scan + .NETIntercept
               DispAbu(Cnt - 1) = GelUMC(.DisplayInd).UMCs(i).ClassAbundance
            End If
        Next i
    End Select
End With
PrepareDisplayToSearch = True
exit_PrepareDisplayToSearch:
End Function


Private Function CalcSlopeIntercept() As Boolean
'---------------------------------------------------------------------------
'least square method to lay best straight line through set of points (xi,yi)
'---------------------------------------------------------------------------
Dim SumY As Double, SumX As Double
Dim SumXY As Double, SumXX As Double
Dim i As Long
SumY = 0:           SumX = 0
SumXY = 0:          SumXX = 0
For i = 0 To MatchCnt - 1
    SumX = SumX + DispNET(MatchDispInd(i))
    SumY = SumY + BaseNET(BaseInd(MatchBaseInd(i)))
    SumXY = SumXY + DispNET(MatchDispInd(i)) * BaseNET(BaseInd(MatchBaseInd(i)))
    SumXX = SumXX + DispNET(MatchDispInd(i)) ^ 2
Next i
AdjSlp = (MatchCnt * SumXY - SumX * SumY) / (MatchCnt * SumXX - SumX * SumX)
AdjInt = (SumY - AdjSlp * SumX) / MatchCnt
CalcSlopeIntercept = True
exit_CalcSlopeIntercept:
End Function


Public Sub DestroyStructures()
Erase BaseMW
Erase BaseNET
Erase BaseAbu
Erase BaseInd

Erase DispMW
Erase DispNET
Erase DispAbu
Erase DispInd

MatchCnt = 0
Erase MatchDispInd
Erase MatchBaseInd

Set MWFastSearch = Nothing
End Sub

Private Function IsJiggyMatch(i As Long, j As Long) As Boolean
'-----------------------------------------------------------------------------
'i is index in Dis arrays, j is index in sorted Base arrays (BaseInd & BaseMW)
'function returns True if i from Dis array and j from Base array make match
'based on OlyJiggyOptions criteria; match is rejected on first unfulfilled criteria
'-----------------------------------------------------------------------------
If OlyJiggyOptions.UseNetConstraint Then
   If Abs(DispNET(i) - BaseNET(BaseInd(j))) > OlyJiggyOptions.NETTol Then Exit Function
End If
If OlyJiggyOptions.UseAbuConstraint Then
   If Abs(DispAbu(i) - BaseAbu(BaseInd(j))) > OlyJiggyOptions.AbuTol Then Exit Function
End If
IsJiggyMatch = True
End Function


VERSION 5.00
Begin VB.Form frmERAnalysis 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expression Ratio Analysis"
   ClientHeight    =   2505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5280
   Icon            =   "frmERAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtERMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      Text            =   "-5"
      Top             =   1620
      Width           =   495
   End
   Begin VB.TextBox txtERMin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Text            =   "-5"
      Top             =   1620
      Width           =   495
   End
   Begin VB.CheckBox chkExcludeERRange 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exclude pairs with ER out of range"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1620
      Width           =   2775
   End
   Begin VB.TextBox txtBinWidth 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Text            =   "0.1"
      Top             =   1140
      Width           =   615
   End
   Begin VB.TextBox txtMaxER 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "5"
      Top             =   1140
      Width           =   615
   End
   Begin VB.CommandButton cmdERStat 
      Caption         =   "&Statistics"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblERType 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ER Calculation Type:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblPCnt 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblPType 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pairs Type:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Pairs:"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   9
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bin Width"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Max. ER"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   735
   End
End
Attribute VB_Name = "frmERAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'calculation of ER values for Delta Label pairs
'----------------------------------------------
'created: 10/12/2001 nt
'last modified: 10/17/2001 nt
'----------------------------------------------
Option Explicit
Dim bLoading As Boolean
Dim CallerID As Long

Dim TmpERType As Long

Dim StatERMin As Double     'statistics is always done on symmetric range
Dim StatERMax As Double
Dim StatBinWidth As Double
Dim StatBinCnt As Long
Dim HalfBinCnt As Long
Dim StatBinInd() As Long    'bin index number
Dim StatBinMin() As Double  'ER bin minimal value
Dim StatBinMax() As Double  'ER bin minimal value
Dim StatBinHits() As Long   'count of hits in bin
Dim StatBadERCnt As Long

Dim StatTop10() As Long     'indexes and counts of top 10 bins

Dim fso As New FileSystemObject

' Unused Function (May 2003)
'''Private Sub cmdCalcER_Click()
''''----------------------------------------------
''''calculate expression ratios for existing pairs
''''----------------------------------------------
'''Dim UserRes As Long
'''Dim i As Long
'''Dim LtAbu As Double
'''Dim HvAbu As Double
'''On Error Resume Next
'''
'''Me.MousePointer = vbHourglass
'''With GelP_D_L(CallerID)
'''    If .SearchDef.ERCalcType <> glER_NONE Then
'''       UserRes = MsgBox("Overwrite existing expression ratios?", vbYesNo, glFGTU)
'''       If UserRes <> vbYes Then Exit Sub
'''    End If
'''    InitDltLblPairsER CallerID
'''    .SearchDef.ERCalcType = TmpERType
'''    Select Case .SearchDef.ERCalcType
'''    Case glER_SOLO_RAT
'''        For i = 0 To .PCnt - 1
'''            LtAbu = GelData(CallerID).IsoData(.Pairs(i).P1).Abundance
'''            HvAbu = GelData(CallerID).IsoData(.Pairs(i).P2).Abundance
'''            .pairs(i).er = ComputeRatioER(LtAbu, HvAbu)
'''            .Pairs(i).ERMemberBasisCount = 1
'''        Next i
'''    Case glER_SOLO_LOG
'''        For i = 0 To .PCnt - 1
'''            LtAbu = GelData(CallerID).IsoData(.Pairs(i).P1).Abundance
'''            HvAbu = GelData(CallerID).IsoData(.Pairs(i).P2).Abundance
'''            .pairs(i).er = ComputeLogER(LtAbu, HvAbu)
'''            .Pairs(i).ERMemberBasisCount = 1
'''        Next i
'''    Case glER_SOLO_ALT
'''        For i = 0 To .PCnt - 1
'''            LtAbu = GelData(CallerID).IsoData(.Pairs(i).P1).Abundance
'''            HvAbu = GelData(CallerID).IsoData(.Pairs(i).P2).Abundance
'''            .pairs(i).er = ComputeAltER(LtAbu, HvAbu)
'''            .Pairs(i).ERMemberBasisCount = 1
'''        Next i
'''    Case Else
'''        MsgBox "Unknown ER calculation type", vbOKOnly, glFGTU
'''    End Select
'''End With
'''Me.MousePointer = vbDefault
'''End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdERStat_Click()
'-------------------------------------------------------
'calculates and reports statistics for expression ratios
'-------------------------------------------------------
If GenerateERStatistic() Then
   ReportERStatistic
Else
   MsgBox "Error generating ER statistics. Make sure bin width is non-zero.", vbOKOnly, glFGTU
End If
End Sub


Private Sub Form_Activate()
If bLoading Then
   CallerID = Me.Tag
   With GelP_D_L(CallerID)
      Select Case .DltLblType
      Case ptUMCDlt, ptUMCLbl, ptUMCDltLbl
           Me.BackColor = glUMC_CLR
      Case ptS_Dlt, ptS_Lbl, ptS_DltLbl
           Me.BackColor = glSOLO_CLR
      End Select
      lblPCnt.Caption = .PCnt
      lblERType.Caption = GetERDesc(.SearchDef.ERCalcType)
      lblPType.Caption = GetPairsTypeDesc(CallerID)
   End With
   bLoading = False
End If
End Sub

Private Sub Form_Load()
bLoading = True
TmpERType = glER_SOLO_ALT
StatERMax = CDbl(txtMaxER.Text)
StatBinWidth = CDbl(txtBinWidth.Text)
End Sub

Private Sub txtBinWidth_LostFocus()
On Error Resume Next
StatBinWidth = CDbl(txtBinWidth.Text)
If Err Or StatBinWidth <= 0 Then
   MsgBox "This parameter should be positive number.", vbOKOnly, glFGTU
   txtBinWidth.SetFocus
End If
End Sub

Private Sub txtMaxER_LostFocus()
On Error Resume Next
StatERMax = CDbl(txtMaxER.Text)
If Err Or StatERMax <= 0 Then
   MsgBox "This parameter should be positive number.", vbOKOnly, glFGTU
   txtMaxER.SetFocus
End If
End Sub

Private Function GenerateERStatistic() As Boolean
'------------------------------------------------------
'actual statistics calculation; returns True on success
'------------------------------------------------------
Dim i As Long
Dim CurrBin As Long
On Error GoTo err_GenerateERStatistic
With GelP_D_L(CallerID)
    'calculate number of bins and reserve space
    Select Case .SearchDef.ERCalcType
    Case glER_SOLO_RAT
       HalfBinCnt = CLng((StatERMax - 1) / StatBinWidth) + 1
    Case glER_SOLO_LOG, glER_SOLO_ALT
       HalfBinCnt = CLng(StatERMax / StatBinWidth) + 1
    End Select
    StatBinCnt = 2 * HalfBinCnt
    ReDim StatBinInd(StatBinCnt - 1)
    ReDim StatBinMin(StatBinCnt - 1)
    ReDim StatBinMax(StatBinCnt - 1)
    ReDim StatBinHits(StatBinCnt - 1)
    ReDim StatTop10(9, 1)                 'top 10 counts
    StatBadERCnt = 0
    'create bins depending on the ER calculation type
    Select Case .SearchDef.ERCalcType
    Case glER_SOLO_RAT  'symmetric around 1; not equidistant
      StatERMin = 1 / StatERMax
      For i = 0 To HalfBinCnt - 1
          StatBinInd(HalfBinCnt + i) = HalfBinCnt + i
          StatBinInd(HalfBinCnt - i - 1) = HalfBinCnt - i - 1
          StatBinMin(HalfBinCnt + i) = 1 + i * StatBinWidth
          StatBinMax(HalfBinCnt + i) = StatBinMin(HalfBinCnt + i) + StatBinWidth
          StatBinMin(HalfBinCnt - i - 1) = 1 / StatBinMax(HalfBinCnt + i)
          StatBinMax(HalfBinCnt - i - 1) = 1 / StatBinMin(HalfBinCnt + i)
      Next i
      StatBinMin(0) = 0
      StatBinMax(StatBinCnt - 1) = -ER_CALC_ERR
    Case glER_SOLO_LOG, glER_SOLO_ALT   'symmetric around 0
      StatERMin = -StatERMax
      For i = 0 To StatBinCnt - 1
          StatBinInd(i) = i
          StatBinMin(i) = StatERMin + (i - 1) * StatBinWidth
          StatBinMax(i) = StatBinMin(i) + StatBinWidth
      Next i
      StatBinMin(0) = ER_CALC_ERR    'some huge negative number
      StatBinMax(StatBinCnt - 1) = -StatBinMin(0) 'some huge positive number
    End Select
    'do count for each bin
    Select Case .SearchDef.ERCalcType
    Case glER_SOLO_RAT  'symmetric around 1; not equidistant
         For i = 0 To .PCnt - 1
            CurrBin = GetCurrentBin_Rat(.Pairs(i).ER)
            If CurrBin >= 0 Then
               StatBinHits(CurrBin) = StatBinHits(CurrBin) + 1
            Else
               StatBadERCnt = StatBadERCnt + 1
            End If
         Next i
    Case glER_SOLO_LOG, glER_SOLO_ALT   'symmetric around 0
         For i = 0 To .PCnt - 1
            CurrBin = GetCurrentBin_Log_Alt(.Pairs(i).ER)
            If CurrBin >= 0 Then
               StatBinHits(CurrBin) = StatBinHits(CurrBin) + 1
            Else
               StatBadERCnt = StatBadERCnt + 1
            End If
         Next i
    End Select
End With
GenerateERStatistic = True
Exit Function

err_GenerateERStatistic:
LogErrors Err.Number, "frmDltLblSER.GenerateERStatistic"
If Err.Number <> 11 Then Resume Next    'division by zero
End Function

Private Sub ReportERStatistic()
'------------------------------------------------------
'report ER statistics in Data Info form
'------------------------------------------------------
Dim tsTmpStat As TextStream
Dim TmpFName As String
Dim i As Long
Dim sLine As String, sHeader As String
On Error GoTo err_ReportERStatistic
TmpFName = GetTempFolder() & RawDataTmpFile
Set tsTmpStat = fso.OpenTextFile(TmpFName, ForWriting, True)
With tsTmpStat
   sHeader = "Bin #:" & glARG_SEP & "[Min:" & glARG_SEP & "Max:>" & glARG_SEP & "Count"
   .WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
   .WriteLine "Gel File: " & GelBody(CallerID).Caption
   .WriteLine "Individual distributions ER statistics"
   Select Case GelP_D_L(CallerID).SearchDef.ERCalcType
   Case glER_SOLO_RAT
        .WriteLine "ER Type: Light/Heavy Ratio; AbuLight/AbuHeavy"
   Case glER_SOLO_LOG
        .WriteLine "ER Type: Logarithmic Light/Heavy Ratio; Ln(AbuLight/AbuHeavy)"
   Case glER_SOLO_ALT
        .WriteLine "ER Type: 0-Centered symmetric Light/Heavy Ratio; (AbuL/AbuH)-1 for AbuL>=AbuH; 1-(AbuH/AbuL) for AbuL<AbuH"
   End Select
   
   .WriteLine "Min ER: " & StatERMin & " - Max ER: " & StatERMax
   .WriteLine "Bin Width: " & StatBinWidth
   .WriteLine "Bad ER Count: " & StatBadERCnt
   .WriteLine
   
   If SortHighScores() Then
      If StatBinCnt > 10 Then
         .WriteLine "Top 10 - Why Lady Sings Gaudeamus Igitur?"
         .WriteLine sHeader
         For i = 0 To 9
           sLine = i & glARG_SEP & StatBinMin(StatBinInd(i)) & glARG_SEP _
                & StatBinMax(StatBinInd(i)) & glARG_SEP & StatBinHits(StatBinInd(i))
           .WriteLine sLine
         Next i
      End If
   Else
      If StatBinCnt > 10 Then .WriteLine "High scores function failed."
   End If
   .WriteLine
   .WriteLine sHeader
   For i = 0 To StatBinCnt - 1
     sLine = StatBinInd(i) & glARG_SEP & StatBinMin(i) & glARG_SEP _
             & StatBinMax(i) & glARG_SEP & StatBinHits(i)
     .WriteLine sLine
   Next i
End With
tsTmpStat.Close
frmDataInfo.Tag = "ER_STAT"
frmDataInfo.Show vbModal
Exit Sub

err_ReportERStatistic:
End Sub

Private Function GetCurrentBin_Rat(ByVal ER As Double) As Long
'----------------------------------------------------------------
'returns current bin for expression ratio statistics; -1 on error
'----------------------------------------------------------------
On Error Resume Next
GetCurrentBin_Rat = -1
If ER <= ER_CALC_ERR Then Exit Function
If ER >= StatERMax Then     'last
   GetCurrentBin_Rat = GelP_D_L(CallerID).PCnt - 1
ElseIf ER < StatERMin Then  'first
   GetCurrentBin_Rat = 0
ElseIf ER >= 1 Then
   GetCurrentBin_Rat = HalfBinCnt + Int((ER - 1) / StatBinWidth)
Else
   GetCurrentBin_Rat = HalfBinCnt - Int((1 / ER - 1) / StatBinWidth) - 1
End If
End Function

Private Function GetCurrentBin_Log_Alt(ByVal ER As Double) As Long
'-----------------------------------------------------------------
'returns current bin for expression ratio statistics; -1 on error
'-----------------------------------------------------------------
On Error Resume Next
GetCurrentBin_Log_Alt = -1
If ER <= ER_CALC_ERR Then Exit Function
If ER >= StatERMax Then     'last
   GetCurrentBin_Log_Alt = GelP_D_L(CallerID).PCnt - 1
ElseIf ER < StatERMin Then  'first
   GetCurrentBin_Log_Alt = 0
ElseIf ER >= 0 Then
   GetCurrentBin_Log_Alt = HalfBinCnt + Int(ER / StatBinWidth)
Else
   GetCurrentBin_Log_Alt = HalfBinCnt - Int(-ER / StatBinWidth) - 1
End If
End Function


Private Function SortHighScores() As Boolean
'-----------------------------------------------
'sort high scores so that we can list Top10 bins
'-----------------------------------------------
Dim qsHS As New QSLong
Dim tmpHS() As Long
On Error Resume Next
tmpHS() = StatBinHits()
If qsHS.QSDesc(tmpHS(), StatBinInd()) Then SortHighScores = True
End Function

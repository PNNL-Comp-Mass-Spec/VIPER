VERSION 5.00
Begin VB.Form frmAbuEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Abundance Editor"
   ClientHeight    =   2100
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOptions 
      Caption         =   "Calculation"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtFactor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Text            =   "10E+6"
         Top             =   1160
         Width           =   975
      End
      Begin VB.OptionButton optAbuCalculation 
         Caption         =   "Multiply with "
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtMaxAbu 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtMinAbu 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optAbuCalculation 
         Caption         =   "Logarithm"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optAbuCalculation 
         Caption         =   "Map to range"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Max."
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Min."
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calc&ulate"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'created: 01/29/2003 nt
'last modified: 01/29/2003 nt
'-----------------------------------------------------
Option Explicit

Const CALC_MAP_TO_RANGE = 0
Const CALC_LOG = 1
Const CALC_MULTIPLY = 2

Dim bLoading As Boolean
Dim bCalcPressed As Boolean

Dim CallerID As Long

Dim CSCnt As Long
Dim ISCnt As Long

Dim CSAbu() As Double
Dim ISAbu() As Double

Dim CurrMinAbu As Double
Dim CurrMaxAbu As Double

Dim CalcType As Long

Dim mMinAbu As Double
Dim mMaxAbu As Double
Dim mFactor As Double

Private Sub cmdCalc_Click()
Dim TmpMinAbu As Double
Dim TmpMaxAbu As Double
Dim AbuScale As Double
Dim i As Long
On Error GoTo err_cmdCalc_Click
Status "Recalculating..."
bCalcPressed = True
TmpMinAbu = 1E+308
TmpMaxAbu = -1E+308
Select Case CalcType
Case CALC_MAP_TO_RANGE
    mMinAbu = CDbl(txtMinAbu.Text)
    mMaxAbu = CDbl(txtMaxAbu.Text)
    AbuScale = (mMaxAbu - mMinAbu) / (CurrMaxAbu - CurrMinAbu)
    For i = 0 To CSCnt - 1
        CSAbu(i) = AbuScale * (CSAbu(i) - CurrMinAbu) + mMinAbu
        If CSAbu(i) < TmpMinAbu Then TmpMinAbu = CSAbu(i)
        If CSAbu(i) > TmpMaxAbu Then TmpMaxAbu = CSAbu(i)
    Next i
    For i = 0 To ISCnt - 1
        ISAbu(i) = AbuScale * (ISAbu(i) - CurrMinAbu) + mMinAbu
        If ISAbu(i) < TmpMinAbu Then TmpMinAbu = ISAbu(i)
        If ISAbu(i) > TmpMaxAbu Then TmpMaxAbu = ISAbu(i)
    Next i
Case CALC_LOG
    For i = 0 To CSCnt - 1
        CSAbu(i) = Log(CSAbu(i)) / Log(10)
        If CSAbu(i) < TmpMinAbu Then TmpMinAbu = CSAbu(i)
        If CSAbu(i) > TmpMaxAbu Then TmpMaxAbu = CSAbu(i)
    Next i
    For i = 0 To ISCnt - 1
        ISAbu(i) = Log(ISAbu(i)) / Log(10)
        If ISAbu(i) < TmpMinAbu Then TmpMinAbu = ISAbu(i)
        If ISAbu(i) > TmpMaxAbu Then TmpMaxAbu = ISAbu(i)
    Next i
Case CALC_MULTIPLY
    mFactor = CDbl(txtFactor.Text)
    For i = 0 To CSCnt - 1
        CSAbu(i) = CSAbu(i) * mFactor
        If CSAbu(i) < TmpMinAbu Then TmpMinAbu = CSAbu(i)
        If CSAbu(i) > TmpMaxAbu Then TmpMaxAbu = CSAbu(i)
    Next i
    For i = 0 To ISCnt - 1
        ISAbu(i) = ISAbu(i) * mFactor
        If ISAbu(i) < TmpMinAbu Then TmpMinAbu = ISAbu(i)
        If ISAbu(i) > TmpMaxAbu Then TmpMaxAbu = ISAbu(i)
    Next i
End Select
CurrMinAbu = TmpMinAbu
CurrMaxAbu = TmpMaxAbu
txtMinAbu.Text = MyFormat(CurrMinAbu)
txtMaxAbu.Text = MyFormat(CurrMaxAbu)
Status "Done."
Exit Sub

err_cmdCalc_Click:
Status ""
MsgBox "Error recalculating abundances.", vbOKOnly, glFGTU
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdReset_Click()
Call LoadAbundances
Status "Reset."
bCalcPressed = False
End Sub

Private Sub Form_Activate()
If bLoading Then
   CallerID = Me.Tag
   Call LoadAbundances
   optAbuCalculation(CALC_MAP_TO_RANGE).value = True
End If
End Sub

Private Sub Form_Load()
bLoading = True
End Sub

Private Function LoadAbundances() As Boolean
'--------------------------------------------------------------------
'loads abundances from the CallerID display
'--------------------------------------------------------------------
Dim i As Long
On Error Resume Next
Status "Loading abundances..."
With GelData(CallerID)
    txtMinAbu.Text = MyFormat(.MinAbu)
    txtMaxAbu.Text = MyFormat(.MaxAbu)
    CurrMinAbu = .MinAbu
    CurrMaxAbu = .MaxAbu
    If .CSLines > 0 Then
       CSCnt = .CSLines
       ReDim CSAbu(CSCnt - 1)
       For i = 1 To .CSLines
           CSAbu(i - 1) = .CSData(i).Abundance
       Next i
    End If
    If .IsoLines > 0 Then
       ISCnt = .IsoLines
       ReDim ISAbu(ISCnt - 1)
       For i = 1 To .IsoLines
           ISAbu(i - 1) = .IsoData(i).Abundance
       Next i
    End If
End With
Status ""
End Function


Private Function SaveAbundances() As Boolean
'--------------------------------------------------------------------
'loads abundances from the CallerID display
'--------------------------------------------------------------------
Dim i As Long
Dim strOldAbundanceRange As String
On Error Resume Next
Status "Saving abundances..."
With GelData(CallerID)
    strOldAbundanceRange = "Old Min abundance = " & Trim(.MinAbu) & "; Old Max abundance = " & Trim(.MaxAbu)
    .MinAbu = CurrMinAbu
    .MaxAbu = CurrMaxAbu
    If .CSLines > 0 Then
       For i = 1 To .CSLines
           .CSData(i).Abundance = CSAbu(i - 1)
       Next i
    End If
    If .IsoLines > 0 Then
       For i = 1 To .IsoLines
          .IsoData(i).Abundance = ISAbu(i - 1)
       Next i
    End If
    AddToAnalysisHistory CallerID, "Abundances adjusted: Min Abundance = " & Trim(.MinAbu) & "; Max Abundance = " & Trim(.MaxAbu) & "; " & strOldAbundanceRange
End With
Status ""
End Function


Private Sub Status(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Res As Long
If bCalcPressed Then
   Res = MsgBox("Replace abundances in original display?", vbYesNoCancel, glFGTU)
   Select Case Res
   Case vbYes
        Call SaveAbundances
   Case vbNo
   Case Else
        Cancel = True
   End Select
End If
End Sub

Private Sub optAbuCalculation_Click(Index As Integer)
CalcType = Index
End Sub

Private Sub txtFactor_LostFocus()
On Error Resume Next
If txtFactor.Text >= 0 Then Exit Sub
MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
txtFactor.SetFocus
End Sub

Private Sub txtMaxAbu_LostFocus()
On Error Resume Next
If txtMaxAbu.Text >= 0 Then Exit Sub
MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
txtMaxAbu.SetFocus
End Sub

Private Sub txtMinAbu_LostFocus()
On Error Resume Next
If txtMinAbu.Text >= 0 Then Exit Sub
MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
txtMinAbu.SetFocus
End Sub

Private Function MyFormat(Number As Double) As String
If Abs(Number) > 1000 Then
   MyFormat = Format(Number, "Scientific")
Else
   MyFormat = Format(Number, "0.0000")
End If
End Function

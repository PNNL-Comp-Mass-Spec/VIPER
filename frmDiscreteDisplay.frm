VERSION 5.00
Begin VB.Form frmDiscreteDisplay 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Discrete Display Function"
   ClientHeight    =   2070
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3060
   Icon            =   "frmDiscreteDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optDiscreteMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ceiling"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton optDiscreteMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Floor"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton optDiscreteMethod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rounding"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txtDecDig 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "2"
      Top             =   180
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discretization method"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal digits"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmDiscreteDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this function discretize display to the certain number of decimal places
'created: 05/01/2002 nt
'last modified: 05/01/2002 nt
'------------------------------------------------------------------------
Option Explicit

Const MAX_DEC_DIG = 5

Const DM_ROUNDING = 0
Const DM_FLOOR = 1
Const DM_CEILING = 2

Dim CallerID As Long
Dim bLoading As Boolean

Dim DecDigCnt As Long
Dim DiscMethod As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Res As Long
Dim Factor As Long
Dim i As Long
On Error GoTo err_OK
Res = MsgBox("This function will change molecular masses in current display. Continue?", vbYesNo, glFGTU)
If Res <> vbYes Then Exit Sub
UpdateStatus "....."
Me.MousePointer = vbHourglass
Factor = 10 ^ DecDigCnt
With GelData(CallerID)
     Select Case DiscMethod
     Case DM_ROUNDING
          For i = 1 To .CSLines
              .CSData(i).AverageMW = Round(.CSData(i).AverageMW, DecDigCnt)
          Next i
          For i = 1 To .IsoLines
              .IsoData(i).AverageMW = Round(.IsoData(i).AverageMW, DecDigCnt)
              .IsoData(i).MonoisotopicMW = Round(.IsoData(i).MonoisotopicMW, DecDigCnt)
              .IsoData(i).MostAbundantMW = Round(.IsoData(i).MostAbundantMW, DecDigCnt)
          Next i
          AddToAnalysisHistory CallerID, "Masses rounded to discrete values; number of decimal places = " & Trim(DecDigCnt)
     Case DM_FLOOR
          For i = 1 To .CSLines
              .CSData(i).AverageMW = Int(.CSData(i).AverageMW * Factor) / Factor
          Next i
          For i = 1 To .IsoLines
              .IsoData(i).AverageMW = Int(.IsoData(i).AverageMW * Factor) / Factor
              .IsoData(i).MonoisotopicMW = Int(.IsoData(i).MonoisotopicMW * Factor) / Factor
              .IsoData(i).MostAbundantMW = Int(.IsoData(i).MostAbundantMW * Factor) / Factor
          Next i
          AddToAnalysisHistory CallerID, "Masses rounded down to discrete values (floor); number of decimal places = " & Trim(DecDigCnt)
     Case DM_CEILING
          For i = 1 To .CSLines
              .CSData(i).AverageMW = Abs(Int(.CSData(i).AverageMW / (-Factor))) * Factor
          Next i
          For i = 1 To .IsoLines
              .IsoData(i).AverageMW = Abs(Int(.IsoData(i).AverageMW * (-Factor))) / Factor
              .IsoData(i).MonoisotopicMW = Abs(Int(.IsoData(i).MonoisotopicMW * (-Factor))) / Factor
              .IsoData(i).MostAbundantMW = Abs(Int(.IsoData(i).MostAbundantMW * (-Factor))) / Factor
          Next i
          AddToAnalysisHistory CallerID, "Masses rounded up to discrete values (ceiling); number of decimal places = " & Trim(DecDigCnt)
     End Select
End With
Me.MousePointer = vbDefault
UpdateStatus "Done."
Exit Sub

err_OK:
MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, glFGTU
End Sub

Private Sub Form_Activate()
If bLoading Then
   CallerID = Me.Tag
   bLoading = False
End If
End Sub

Private Sub Form_Load()
bLoading = True
DecDigCnt = CLng(txtDecDig.Text)
DiscMethod = DM_ROUNDING
End Sub

Private Sub optDiscreteMethod_Click(Index As Integer)
DiscMethod = Index
End Sub

Private Sub txtDecDig_LostFocus()
Dim tmp As String
tmp = Trim$(txtDecDig.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 And tmp <= MAX_DEC_DIG Then
      DecDigCnt = CLng(tmp)
      Exit Sub
   End If
End If
MsgBox "Number of decimal digits should be integer between 0 and " & MAX_DEC_DIG & ".", vbOKOnly, glFGTU
txtDecDig.SetFocus
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

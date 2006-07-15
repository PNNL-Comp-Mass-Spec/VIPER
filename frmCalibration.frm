VERSION 5.00
Begin VB.Form frmCalibration 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calibration"
   ClientHeight    =   2760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Calibration"
      Height          =   2055
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtNewC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Text            =   "0.0"
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox txtNewB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Text            =   "0.0"
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtNewA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   "0.0"
         Top             =   900
         Width           =   1815
      End
      Begin VB.ComboBox cmbNewCalFn 
         Height          =   315
         ItemData        =   "frmCalibration.frx":0000
         Left            =   120
         List            =   "frmCalibration.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblArg1 
         BackStyle       =   0  'Transparent
         Caption         =   "C:"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblArg1 
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblArg1 
         BackStyle       =   0  'Transparent
         Caption         =   "A:"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Calibration"
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtCurrCalFn 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Text            =   "O tempora! O mores!"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtCurrArg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   7
         Text            =   "0.0"
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox txtCurrArg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Text            =   "0.0"
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtCurrArg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Text            =   "0.0"
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblArg1 
         BackStyle       =   0  'Transparent
         Caption         =   "C:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblArg1 
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblArg1 
         BackStyle       =   0  'Transparent
         Caption         =   "A:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdReCal 
      Caption         =   "ReCal"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Messangero e non importante!"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2340
      Width           =   3735
   End
End
Attribute VB_Name = "frmCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this function allows change of calibration function
'and recalculation of all data points
'last modified: 06/18/2002 nt
'----------------------------------------------------
Option Explicit

'''Const CAL_OLD = 0
'''Const CAL_NEW = 1

Dim CallerID As Long
Dim bLoading As Long

Dim CurrCal As Object       'recalculation objects
Dim NewCal As Object

Dim NewA As Double
Dim NewB As Double
Dim NewC As Double
Dim NewCalFN As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdReCal_Click()
Dim eResponse As VbMsgBoxResult
eResponse = MsgBox("This function will recalculate all masses in the file based on specified selection of new calibration function. Continue?", vbYesNoCancel)
Select Case eResponse
Case vbNo       'abort and unload
     Unload Me
Case vbYes      'recalculate
     If PrepareCurrCal() And PrepareNewCal() Then
        If Recalculate() Then
           Call SetNewAsCurrent
           GelStatus(CallerID).Dirty = True
           lblStatus.Caption = "Done."
           MsgBox "Data has been recalculated.", vbOKOnly, "Done"
        Else
           MsgBox "Error recalculating molecular masses.", vbOKOnly
        End If
     Else
        MsgBox "Karamba, something is wrong.", vbOKOnly, glFGTU
     End If
Case vbCancel   'do nothing
End Select
End Sub

Private Sub Form_Activate()
If bLoading Then
   CallerID = Me.Tag
   Call SetCurrCalSettings
   bLoading = False
End If
End Sub

Private Sub Form_Load()
lblStatus.Caption = ""
cmbNewCalFn.Clear
cmbNewCalFn.AddItem CAL_EQUATION_1
cmbNewCalFn.AddItem CAL_EQUATION_2
cmbNewCalFn.AddItem CAL_EQUATION_3
cmbNewCalFn.AddItem CAL_EQUATION_4
cmbNewCalFn.AddItem CAL_EQUATION_5
bLoading = True
End Sub


Public Function PrepareCurrCal() As Boolean
'--------------------------------------------------------------
'prepares current calibration calculator and returns True if OK
'--------------------------------------------------------------
On Error Resume Next
With GelData(CallerID)
  Select Case UCase(.CalEquation)
  Case UCase(CAL_EQUATION_1)
       Set CurrCal = New CalEq1
       CurrCal.A = .CalArg(1)
       CurrCal.B = .CalArg(2)
       If Not CurrCal.EquationOK Then lblStatus.Caption = "Invalid parameters in current calibration equation."
  Case UCase(CAL_EQUATION_5)
       Set CurrCal = New CalEq5
       CurrCal.A = .CalArg(1)
       CurrCal.B = .CalArg(2)
       If Not CurrCal.EquationOK Then lblStatus.Caption = "Invalid parameters in current calibration equation."
  Case UCase(CAL_EQUATION_2), UCase(CAL_EQUATION_3), UCase(CAL_EQUATION_4)
       If .CalArg(3) = 0 Then    'same case as CAL_EQUATION_1
          Set CurrCal = New CalEq1
          CurrCal.A = .CalArg(1)
          CurrCal.B = .CalArg(2)
          If Not CurrCal.EquationOK Then lblStatus.Caption = "Invalid parameters in current calibration equation."
       Else
          lblStatus.Caption = "Calibration equation form not implemented."
       End If
  Case Else     'later we might use other calibration equations
       lblStatus.Caption = "Calibration equation not found. Lock mass function not possible."
  End Select
End With
PrepareCurrCal = CurrCal.EquationOK
End Function


Public Function PrepareNewCal() As Boolean
'----------------------------------------------------------
'prepares new calibration calculator and returns True if OK
'----------------------------------------------------------
On Error Resume Next
NewCalFN = cmbNewCalFn.Text
NewA = CDbl(txtNewA.Text)
NewB = CDbl(txtNewB.Text)
NewC = CDbl(txtNewC.Text)
Select Case UCase(NewCalFN)
Case UCase(CAL_EQUATION_1)
     Set NewCal = New CalEq1
     NewCal.A = NewA
     NewCal.B = NewB
     If Not NewCal.EquationOK Then lblStatus.Caption = "Invalid parameters in new calibration equation."
Case UCase(CAL_EQUATION_5)
     Set NewCal = New CalEq5
     NewCal.A = NewA
     NewCal.B = NewB
     If Not NewCal.EquationOK Then lblStatus.Caption = "Invalid parameters in new calibration equation."
Case UCase(CAL_EQUATION_2), UCase(CAL_EQUATION_3), UCase(CAL_EQUATION_4)
     If NewC = 0 Then    'same case as CAL_EQUATION_1
        Set NewCal = New CalEq1
        NewCal.A = NewA
        NewCal.B = NewB
        If Not NewCal.EquationOK Then lblStatus.Caption = "Invalid parameters in new calibration equation."
     Else
        lblStatus.Caption = "Calibration equation form not implemented."
     End If
Case Else     'later we might use other calibration equations
       lblStatus.Caption = "Calibration equation not found."
End Select
PrepareNewCal = NewCal.EquationOK
End Function


Public Sub SetNewAsCurrent()
'-------------------------------------------------------------
'sets new calibration equation as current and redraws settings
'-------------------------------------------------------------
With GelData(CallerID)
    .CalEquation = NewCalFN
    .CalArg(1) = NewA
    .CalArg(2) = NewB
    .CalArg(3) = NewC
End With
Call SetCurrCalSettings
End Sub

Public Sub SetCurrCalSettings()
Dim i As Long
With GelData(CallerID)
     txtCurrCalFn.Text = .CalEquation
     For i = 1 To 3
         txtCurrArg(i - 1).Text = .CalArg(i)
     Next i
End With
End Sub

Public Function Recalculate() As Boolean
'---------------------------------------
'actual recalculation is done here
'---------------------------------------
Dim i As Long
Dim MOverZ As Double
Dim CS As Double
Dim Freq As Double
On Error GoTo err_Recalculate

lblStatus.Caption = "Recalculating masses"
Me.MousePointer = vbHourglass
With GelData(CallerID)

    For i = 1 To .CSLines
    Next i

    lblStatus.Caption = "Recalculating masses: 0 / " & Trim(.IsoLines)
    For i = 1 To .IsoLines
        If i Mod 1000 = 0 Then
            lblStatus.Caption = "Recalculating masses: " & Trim(i) & " / " & Trim(.IsoLines)
            DoEvents
        End If
        
        CS = .IsoData(i).Charge
        MOverZ = .IsoData(i).MonoisotopicMW / CS + glMASS_CC
        Freq = CurrCal.CyclotFreq(MOverZ)
        MOverZ = NewCal.MOverZ(Freq)
        .IsoData(i).MZ = MOverZ
        .IsoData(i).MonoisotopicMW = (MOverZ - glMASS_CC) * CS
    Next i
    
    AddToAnalysisHistory CallerID, "Updated calibration equation and recalculated all isotopic masses; Old equation " & CurrCal.CalDescription & "; New equation " & NewCal.CalDescription & "; Number of data points recalculated = " & .IsoLines
End With
Recalculate = True
exit_Recalculate:
Me.MousePointer = vbDefault
Exit Function

err_Recalculate:
LogErrors Err.Number, "frmCalibration.Recalculate"
Resume exit_Recalculate
End Function


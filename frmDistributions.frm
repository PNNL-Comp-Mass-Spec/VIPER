VERSION 5.00
Begin VB.Form frmDistributions 
   Caption         =   "Distributions"
   ClientHeight    =   4455
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5040
   Icon            =   "frmDistributions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VIPER.LaDist LaDist1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6588
   End
   Begin VB.CommandButton cmdSendToText 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Save As Text"
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cmbParameter 
      Height          =   315
      ItemData        =   "frmDistributions.frx":030A
      Left            =   720
      List            =   "frmDistributions.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Show"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmDistributions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Function to display distributions of different display parameters
'created: 04/07/2002 nt
'last modified: 04/16/2003 nt
'-----------------------------------------------------------------
Option Explicit

Const D_MW = 0
Const D_UMC_MW = 1
Const D_ABU = 2
Const D_UMC_ABU = 3
Const D_ER = 4
Const D_CHARGE_STATE = 5
Const D_FIT = 6
Const D_Scan = 7
Const D_UMC_CNT = 8
Const D_PAIRS_DELTA = 9
Const D_PAIRS_LABEL = 10
Const D_MOverZ = 11
Const D_StDev = 12
Const D_Distances = 13
Const D_UMC_MASS_RANGE_ABS = 14
Const D_UMC_MASS_RANGE_PPM = 15
Const D_UMC_SCAN_RANGE = 16

Dim bLoading As Boolean
Dim CallerID As Long

Dim Parameter As Long

Private Sub ResizeForm()
    Dim lngDesiredValue As Long
    
    With LaDist1
        .Left = 120
        .Top = 600
        
        lngDesiredValue = Me.ScaleWidth - .Left - 120
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .width = lngDesiredValue
        
        lngDesiredValue = Me.ScaleHeight - .Top - 120
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .Height = lngDesiredValue
        
    End With
    
    
End Sub

Private Sub cmbParameter_Click()
Parameter = cmbParameter.ListIndex
Call GoForIt
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSendToText_Click()
Dim TmpFileName As String
TmpFileName = GetTempFolder & RawDataTmpFile
If LaDist1.IsFree Then
   MsgBox "Distribution function not found.", vbOKOnly, glFGTU
Else
   LaDist1.WriteDFToFile TmpFileName
   DoEvents
   frmDataInfo.Tag = "DF"
   frmDataInfo.Show vbModal
End If
End Sub

Private Sub Form_Activate()
If bLoading Then
   CallerID = Me.Tag
   bLoading = False
   cmbParameter.ListIndex = 0
Else
   LaDist1.DFRefresh
End If
End Sub


Private Sub Form_Load()
bLoading = True
With cmbParameter
    .Clear
    .AddItem "Molecular Mass"
    .AddItem "Molecular Mass UMC"
    .AddItem "Abundance"
    .AddItem "Abundance UMC"
    .AddItem "Expression Ratio"
    .AddItem "Charge States"
    .AddItem "Isotopic Fit"
    .AddItem "Scan"
    .AddItem "UMC Count"
    .AddItem "Pairs Deltas"
    .AddItem "Pairs Labels"
    .AddItem "m/z"
    .AddItem "St.Dev. CS"
End With

    Me.width = 5490
    Me.Height = 5265
End Sub


Private Sub GoForIt()
Dim i As Long
Dim Data() As Double
Dim TmpCnt As Long
Dim MinValue As Double, MaxValue As Double

If bLoading Then Exit Sub
Select Case Parameter
Case D_MW
  With GelData(CallerID)
     TmpCnt = .CSLines + .IsoLines
     If TmpCnt > 0 Then
        ReDim Data(TmpCnt - 1)
        For i = 1 To .CSLines
            Data(i - 1) = .CSData(i).AverageMW
        Next i
        For i = 1 To .IsoLines
            Data(.CSLines + i - 1) = .IsoData(i).MonoisotopicMW
        Next i
        LaDist1.HLabel = "MW"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 0, .MaxMW, 4096, BinsUni
     Else
        MsgBox "No data found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_UMC_MW
  With GelUMC(CallerID)
     If .UMCCnt > 0 Then
        ReDim Data(.UMCCnt - 1)
        For i = 0 To .UMCCnt - 1
            Data(i) = .UMCs(i).ClassMW
        Next i
        LaDist1.HLabel = "MW-UMC"
        LaDist1.HNumFmt = "0.0000"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, GelData(CallerID).MinMW, GelData(CallerID).MaxMW, 4096, BinsUni
     Else
        MsgBox "Unique mass classes not found.", vbOKOnly, glFGTU
     End If
  End With
Case D_ABU
  With GelData(CallerID)
     TmpCnt = .CSLines + .IsoLines
     If TmpCnt > 0 Then
        ReDim Data(TmpCnt - 1)
        For i = 1 To .CSLines
            Data(i - 1) = .CSData(i).Abundance
        Next i
        For i = 1 To .IsoLines
            Data(.CSLines + i - 1) = .IsoData(i).Abundance
        Next i
        LaDist1.HLabel = "Abundance"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 5, 10, 512, BinsLog
     Else
        MsgBox "No data found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_UMC_ABU
  With GelUMC(CallerID)
     If .UMCCnt > 0 Then
        ReDim Data(.UMCCnt - 1)
        For i = 0 To .UMCCnt - 1
            Data(i) = .UMCs(i).ClassAbundance
        Next i
        LaDist1.HLabel = "Abu-UMC"
        LaDist1.HNumFmt = "0.0000"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 5, 10, 512, BinsLog
     Else
        MsgBox "Unique mass classes not found.", vbOKOnly, glFGTU
     End If
  End With
Case D_ER
  If GelP_D_L(CallerID).PCnt > 0 Then
     LaDist1.HLabel = "ER"
     LaDist1.HNumFmt = "0.00"
     LaDist1.VLabel = "Frequency"
          
     With GelP_D_L(CallerID)
        ReDim Data(.PCnt - 1)
        For i = 0 To .PCnt - 1
           Data(i) = .Pairs(i).ER
        Next i
     End With
     
     LaDist1.DataFill Data(), 0, 10, 256, BinsRat
  Else
     MsgBox "No expression ratio data found.", vbOKOnly, glFGTU
  End If
Case D_CHARGE_STATE
  With GelData(CallerID)
     TmpCnt = .CSLines + .IsoLines
     If TmpCnt > 0 Then
        ReDim Data(TmpCnt - 1)
        For i = 1 To .CSLines
            Data(i - 1) = .CSData(i).Charge
        Next i
        For i = 1 To .IsoLines
            Data(.CSLines + i - 1) = .IsoData(i).Charge
        Next i
        LaDist1.HLabel = "CS"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 0, 100, 100, BinsUni
     Else
        MsgBox "No data found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_FIT
  With GelData(CallerID)
     If .IsoLines > 0 Then
        ReDim Data(.IsoLines - 1)
        For i = 1 To .IsoLines
            Data(i - 1) = .IsoData(i).Fit
        Next i
        LaDist1.HLabel = "Fit"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 0, 1, 32, BinsUni
     Else
        MsgBox "No data found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_Scan
  With GelData(CallerID)
     TmpCnt = .CSLines + .IsoLines
     If TmpCnt > 0 Then
        ReDim Data(TmpCnt - 1)
        For i = 1 To .CSLines
            Data(i - 1) = .CSData(i).ScanNumber
        Next i
        For i = 1 To .IsoLines
            Data(.CSLines + i - 1) = .IsoData(i).ScanNumber
        Next i
        LaDist1.HLabel = "Scan"
        LaDist1.HNumFmt = "0"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 0, .ScanInfo(UBound(.ScanInfo)).ScanNumber, .ScanInfo(UBound(.ScanInfo)).ScanNumber + 1, BinsUni
     Else
        MsgBox "No data found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_UMC_CNT
  With GelUMC(CallerID)
     If .UMCCnt > 0 Then
        ReDim Data(.UMCCnt - 1)
        For i = 0 To .UMCCnt - 1
            Data(i) = .UMCs(i).ClassCount
        Next i
        LaDist1.HLabel = "UMC-Count"
        LaDist1.HNumFmt = "0"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 0, 100, 100, BinsUni
     Else
        MsgBox "Unique mass classes not found.", vbOKOnly, glFGTU
     End If
  End With
Case D_PAIRS_DELTA
  With GelP_D_L(CallerID)
    If .PCnt > 0 Then
       Select Case .DltLblType
       Case ptUMCDlt, ptUMCDltLbl, ptS_Dlt = 1, ptS_DltLbl
          LaDist1.HLabel = "Delta Count"
          LaDist1.HNumFmt = "0.00"
          LaDist1.VLabel = "Frequency"
          ReDim Data(.PCnt - 1)
          For i = 0 To .PCnt - 1
            Data(i) = .Pairs(i).P2DltCnt
          Next i
          LaDist1.DataFill Data, 0, 200, 201, BinsUni
       Case Else
          MsgBox "Current pair type does not allows for this count.", vbOKOnly, glFGTU
       End Select
    Else
       MsgBox "No pairs data found.", vbOKOnly, glFGTU
    End If
  End With
Case D_PAIRS_LABEL
  With GelP_D_L(CallerID)
    If .PCnt > 0 Then
       Select Case .DltLblType
       Case ptUMCDlt, ptUMCDltLbl, ptS_Dlt = 1, ptS_DltLbl
          LaDist1.HLabel = "Label Count"
          LaDist1.HNumFmt = "0.00"
          LaDist1.VLabel = "Frequency"
          ReDim Data(.PCnt - 1)
          For i = 0 To .PCnt - 1
            Data(i) = .Pairs(i).P2LblCnt
          Next i
          LaDist1.DataFill Data, 0, 50, 51, BinsUni
        Case Else
          MsgBox "Current pair type does not allows for this count.", vbOKOnly, glFGTU
        End Select
    Else
        MsgBox "No pairs data found.", vbOKOnly, glFGTU
    End If
  End With
Case D_MOverZ
  With GelData(CallerID)
     TmpCnt = .CSLines + .IsoLines
     If TmpCnt > 0 Then
        ReDim Data(TmpCnt - 1)
        For i = 1 To .CSLines
            Data(i - 1) = .CSData(i).AverageMW / .CSData(i).Charge
        Next i
        For i = 1 To .IsoLines
            Data(.CSLines + i - 1) = .IsoData(i).MZ
        Next i
        LaDist1.HLabel = "m/z"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 0, 5000, 1024, BinsUni
     Else
        MsgBox "No data found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_StDev
  With GelData(CallerID)
     If .CSLines > 0 Then
        ReDim Data(.CSLines - 1)
        For i = 1 To .CSLines
            Data(i - 1) = .CSData(i).MassStDev
        Next i
        LaDist1.HLabel = "StDev"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, 0, 100, 256, BinsUni
     Else
        MsgBox "No Charge State data found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_Distances
  With GelUMCIon(CallerID)
     If .NetCount > 0 Then
        ReDim Data(.NetCount - 1)
        For i = 0 To .NetCount - 1
            Data(i) = .NetDist(i)
        Next i
        LaDist1.HLabel = "Distances"
        LaDist1.HNumFmt = "0.00E-00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, .MinDist, .MaxDist, 16, BinsUni
     Else
        MsgBox "Metric not found in current display.", vbOKOnly, glFGTU
     End If
  End With
Case D_UMC_MASS_RANGE_ABS
  With GelUMC(CallerID)
     If .UMCCnt > 0 Then
        MinValue = glHugeDouble:    MaxValue = -glHugeDouble
        ReDim Data(.UMCCnt - 1)
        For i = 0 To .UMCCnt - 1
            Data(i) = Abs(.UMCs(i).MaxMW - .UMCs(i).MinMW)
            If Data(i) < MinValue Then MinValue = Data(i)
            If Data(i) > MaxValue Then MaxValue = Data(i)
        Next i
        LaDist1.HLabel = "UMC MW Range(Da)"
        LaDist1.HNumFmt = "0.00E-00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, MinValue, MaxValue, 128, BinsUni
     Else
        MsgBox "Unique mass classes not found.", vbOKOnly, glFGTU
     End If
  End With
Case D_UMC_MASS_RANGE_PPM
  With GelUMC(CallerID)
     If .UMCCnt > 0 Then
        MinValue = glHugeDouble:    MaxValue = -glHugeDouble
        ReDim Data(.UMCCnt - 1)
        For i = 0 To .UMCCnt - 1
            Data(i) = (Abs(.UMCs(i).MaxMW - .UMCs(i).MinMW) / .UMCs(i).ClassMW) * 1000000
            If Data(i) < MinValue Then MinValue = Data(i)
            If Data(i) > MaxValue Then MaxValue = Data(i)
        Next i
        LaDist1.HLabel = "UMC MW Range(ppm)"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, MinValue, MaxValue, 128, BinsUni
     Else
        MsgBox "Unique mass classes not found.", vbOKOnly, glFGTU
     End If
  End With
Case D_UMC_SCAN_RANGE
  With GelUMC(CallerID)
     If .UMCCnt > 0 Then
        MinValue = glHugeLong:    MaxValue = -glHugeLong
        ReDim Data(.UMCCnt - 1)
        For i = 0 To .UMCCnt - 1
            Data(i) = .UMCs(i).MaxScan - .UMCs(i).MinScan + 1
            If Data(i) < MinValue Then MinValue = Data(i)
            If Data(i) > MaxValue Then MaxValue = Data(i)
        Next i
        LaDist1.HLabel = "UMC MW Range(ppm)"
        LaDist1.HNumFmt = "0.00"
        LaDist1.VLabel = "Frequency"
        LaDist1.DataFill Data, MinValue, MaxValue, (MaxValue - MinValue + 1), BinsUni
     Else
        MsgBox "Unique mass classes not found.", vbOKOnly, glFGTU
     End If
  End With
End Select
End Sub

Private Sub Form_Paint()
If Not bLoading Then LaDist1.DFRefresh
End Sub

Private Sub Form_Resize()
    ResizeForm
End Sub

Private Sub LaDist1_ArgChange()
LaDist1.DFRefresh
End Sub


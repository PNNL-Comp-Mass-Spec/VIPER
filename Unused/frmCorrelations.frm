VERSION 5.00
Begin VB.Form frmCorrelations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Correlations"
   ClientHeight    =   5310
   ClientLeft      =   2760
   ClientTop       =   4035
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   ShowInTaskbar   =   0   'False
   Begin VIPER.LaSpots LaSpots1 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8070
   End
   Begin VB.ComboBox cmbCorrelationType 
      Height          =   315
      ItemData        =   "frmCorrelations.frx":0000
      Left            =   120
      List            =   "frmCorrelations.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuF 
      Caption         =   "&Function"
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVFixedWindow 
         Caption         =   "&Fixed Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVVariableWindow 
         Caption         =   "Variable Window"
      End
   End
End
Attribute VB_Name = "frmCorrelations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim CallerID As Long

Const CORR_PAIRS_LABU_HABU = 0
Const CORR_PAIRS_LCNT_HCNT = 1

Dim Cnt As Long
' MonroeMod: Changed to String
Dim XYID() As String
Dim x() As Double
Dim y() As Double
' MonroeMod
Dim Intensity() As Double

Dim CorrType As Long

Private Sub cmbCorrelationType_Click()
CorrType = cmbCorrelationType.ListIndex
Select Case CorrType
Case CORR_PAIRS_LABU_HABU
     If GelP_D_L(CallerID).PCnt > 0 Then
        If FillDataPairsAbuCorr() Then
           LaSpots1.SetFixedWindow 6, 10, 6, 10
           LaSpots1.VNumFmt = "0.00"
           LaSpots1.HNumFmt = "0.00"
           LaSpots1.AddSpotsMany XYID(), x(), y(), Intensity()
        Else
           MsgBox "Error retrieving pairs data!", vbOKOnly, glFGTU
        End If
     Else
        MsgBox "Pairs not found! Apply some of pairing functions first!", vbOKOnly, glFGTU
     End If
Case CORR_PAIRS_LCNT_HCNT
     If GelP_D_L(CallerID).PCnt > 0 Then
        If FillDataPairsCnt() Then
           LaSpots1.SetFixedWindow 0, 1.1, 0, 1.1
           LaSpots1.VNumFmt = "0.00"
           LaSpots1.HNumFmt = "0.00"
           LaSpots1.AddSpotsMany XYID(), x(), y(), Intensity()
        Else
           MsgBox "Error retrieving pairs data!", vbOKOnly, glFGTU
        End If
     Else
        MsgBox "Pairs not found! Apply some of pairing functions first!", vbOKOnly, glFGTU
     End If
End Select
End Sub

Private Function FillDataPairsAbuCorr() As Boolean
Dim i As Long
Dim Ln10 As Double
On Error GoTo err_FillDataPairsAbuCorr
Ln10 = Log(10#)
With GelP_D_L(CallerID)
    Cnt = .PCnt
    ReDim XYID(Cnt - 1)
    ReDim x(Cnt - 1)
    ReDim y(Cnt - 1)
' MonroeMod
    ReDim Intensity(Cnt - 1)
    Select Case .DltLblType
    Case ptUMCDlt, ptUMCLbl, ptUMCDltLbl        'UMC pairs
        For i = 0 To .PCnt - 1
            XYID(i) = i
            x(i) = Log(GelUMC(CallerID).UMCs(.Pairs(i).P1).ClassAbundance) / Ln10
            y(i) = Log(GelUMC(CallerID).UMCs(.Pairs(i).P2).ClassAbundance) / Ln10
        Next i
    Case ptS_Dlt, ptS_Lbl, ptS_DltLbl           'Solo pairs
        For i = 0 To .PCnt - 1
            XYID(i) = i
            x(i) = Log(GelData(CallerID).IsoNum(.Pairs(i).P1, isfAbu)) / Ln10
            y(i) = Log(GelData(CallerID).IsoNum(.Pairs(i).P2, isfAbu)) / Ln10
        Next i
    End Select
End With
FillDataPairsAbuCorr = True
err_FillDataPairsAbuCorr:
End Function


Private Function FillDataPairsCnt() As Boolean
Dim i As Long
On Error GoTo err_FillDataPairsCnt
With GelP_D_L(CallerID)
    Cnt = .PCnt
    ReDim XYID(Cnt - 1)
    ReDim x(Cnt - 1)
    ReDim y(Cnt - 1)
' MonroeMod
    ReDim Intensity(Cnt - 1)
    Select Case .DltLblType
    Case ptUMCDlt, ptUMCLbl, ptUMCDltLbl        'UMC pairs
        For i = 0 To .PCnt - 1
            XYID(i) = i
            x(i) = 1 / GelUMC(CallerID).UMCs(.Pairs(i).P1).ClassCount
            y(i) = 1 / GelUMC(CallerID).UMCs(.Pairs(i).P2).ClassCount
        Next i
    Case ptS_Dlt, ptS_Lbl, ptS_DltLbl           'Solo pairs
        For i = 0 To .PCnt - 1
            XYID(i) = i
            x(i) = 1
            y(i) = 1
        Next i
    End Select
End With
FillDataPairsCnt = True
err_FillDataPairsCnt:
End Function


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
End Sub

Private Sub Form_Load()
' MonroeMod
LaSpots1.MaxSpotSizeLogicalCoords = 75
With cmbCorrelationType
    .Clear
    .AddItem "Pairs -Intensities"
    .AddItem "Pairs - UMC Counts"
    .AddItem "Peaks -Fit - Intensity"
End With

End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuVFixedWindow_Click()
mnuVFixedWindow.Checked = True
mnuVVariableWindow.Checked = False
LaSpots1.ViewWindow = sFixedWindow
End Sub

Private Sub mnuVVariableWindow_Click()
mnuVFixedWindow.Checked = False
mnuVVariableWindow.Checked = True
LaSpots1.ViewWindow = sVariableWindow
End Sub


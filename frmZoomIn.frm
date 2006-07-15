VERSION 5.00
Begin VB.Form frmZoomIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zoom In"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2340
   Icon            =   "frmZoomIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSynchronize 
      Caption         =   "&Synchronize all displays"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Check to apply the same range to all gels."
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtYMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtYMin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtXMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtXMin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lbl1 
      Caption         =   "MW Max"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1215
      Width           =   1000
   End
   Begin VB.Label lbl1 
      Caption         =   "MW Min"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   855
      Width           =   1000
   End
   Begin VB.Label lbl1 
      Caption         =   "File # Max"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   495
      Width           =   1000
   End
   Begin VB.Label lbl1 
      Caption         =   "File # Min"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1000
   End
End
Attribute VB_Name = "frmZoomIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Last modified 09/20/2002 nt
'-----------------------------------------------------------------------------
Option Explicit

Private Const NumbersOnly = "This argument must be a number!"
Private Const NonNegativeOnly = "Range arguments can not be negative numbers!"
Private Const MinMaxDifferent = "Min and Max arguments can not be the same!"
Private Const MinLTMax = "Minimum value has to be less than maximum value!"

Dim CallerID As Long

Dim X1 As Double, x2 As Double
Dim Y1 As Double, Y2 As Double

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

Dim lngStartIndex As Long, lngEndIndex As Long
Dim i As Long
Dim FN1 As Long, FN2 As Long
Dim minNET As Double, maxNET As Double

On Error GoTo ZoomInErrorHandler

If ValidationOK Then
   If CallerID = OlyCallerID Then
      If frmGraphOverlay.MyCooSys.csYScale = glVAxisLog Then
         Y1 = CDbl(Log(Y1) / Log(10#)):     Y2 = CDbl(Log(Y2) / Log(10#))
      End If
      frmGraphOverlay.MyCooSys.ZoomInR CSng(X1), CSng(Y1), CSng(x2), CSng(Y2)
   Else
     
      If cChkBox(chkSynchronize) Then
        lngStartIndex = 1
        lngEndIndex = UBound(GelBody)
      Else
        lngStartIndex = CallerID
        lngEndIndex = CallerID
      End If
      
      Select Case GelBody(CallerID).csMyCooSys.csType
      Case glFNCooSys
        If cChkBox(chkSynchronize) Then
            ' Convert from scan number to NET for this gel
            ' Then, convert from NET to scan number for all other gels, and zoom in,
            ' thus synchronizing
            minNET = ScanToGANET(CallerID, CLng(X1))
            maxNET = ScanToGANET(CallerID, CLng(x2))
            
            For i = lngStartIndex To lngEndIndex
                If i = CallerID Then
                    FN1 = CLng(X1)
                    FN2 = CLng(x2)
                Else
                    If Not GelStatus(i).Deleted Then
                        FN1 = GANETToScan(i, minNET)
                        FN2 = GANETToScan(i, maxNET)
                    End If
                End If
                
                ZoomGelToDimensions i, CSng(FN1), Y1, CSng(FN2), Y2
            Next i
        Else
            ZoomGelToDimensions CallerID, CSng(X1), Y1, CSng(x2), Y2
        End If
      Case glNETCooSys
        ' Convert from NET to scan number and callZ oomGelToDimensions for each gel
        For i = lngStartIndex To lngEndIndex
            If Not GelStatus(i).Deleted Then
                X1 = GANETToScan(i, X1)
                x2 = GANETToScan(i, x2)
                ZoomGelToDimensions i, CSng(X1), Y1, CSng(x2), Y2
            End If
        Next i
      Case glPICooSys
        For i = lngStartIndex To lngEndIndex
            If Not GelStatus(i).Deleted And GelData(i).pICooSysEnabled Then
                ZoomGelToDimensions CallerID, CSng(x2), Y1, CSng(X1), Y2
            End If
        Next i
      End Select
   End If
   Unload Me
End If
Exit Sub

ZoomInErrorHandler:
Debug.Print "Error in frmZoomIn.cmdOK"
Debug.Assert False
Resume Next

End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
If GetChildCount() > 1 And CallerID >= 0 Then
   chkSynchronize.Enabled = True
Else
   chkSynchronize.Enabled = False
End If

If CallerID = OlyCallerID Then
   lbl1(0).Caption = "Min NET"
   lbl1(1).Caption = "Max NET"
   With frmGraphOverlay.MyCooSys
        txtXMin.Text = Round(.CurrRXMin, 4)
        txtXMax.Text = Round(.CurrRXMax, 4)
        If .csYScale = glVAxisLin Then
           txtYMin.Text = Format$(.CurrRYMin, "0.0000")
           txtYMax.Text = Format$(.CurrRYMax, "0.0000")
        Else
           txtYMin.Text = Format$(10 ^ .CurrRYMin, "0.0000")
           txtYMax.Text = Format$(10 ^ .CurrRYMax, "0.0000")
        End If
   End With
Else
   With GelBody(CallerID).csMyCooSys
      Select Case .csType
      Case glPICooSys
         lbl1(0).Caption = "Min pI"
         lbl1(1).Caption = "Max pI"
         txtXMin.Text = GelData(CallerID).ScanInfo(GetDFIndex(CallerID, .CurrRXMax)).ScanPI
         txtXMax.Text = GelData(CallerID).ScanInfo(GetDFIndex(CallerID, .CurrRXMin)).ScanPI
      Case glFNCooSys
         lbl1(0).Caption = "Scan # Min"
         lbl1(1).Caption = "Scan # Max"
         txtXMin.Text = .CurrRXMin
         txtXMax.Text = .CurrRXMax
      Case glNETCooSys
         lbl1(0).Caption = "NET Min"
         lbl1(1).Caption = "NET Max"
         txtXMin.Text = Round(ScanToGANET(CallerID, .CurrRXMin), 4)
         txtXMax.Text = Round(ScanToGANET(CallerID, .CurrRXMax), 4)
      End Select
      If .csYScale = glVAxisLin Then
         txtYMin.Text = Format$(.CurrRYMin, "0.0000")
         txtYMax.Text = Format$(.CurrRYMax, "0.0000")
      Else
         txtYMin.Text = Format$(10 ^ .CurrRYMin, "0.0000")
         txtYMax.Text = Format$(10 ^ .CurrRYMax, "0.0000")
      End If
   End With
End If
SelectText txtXMin
txtXMin.SetFocus
End Sub

Private Sub SelectText(t As TextBox)
t.SelStart = 0
t.SelLength = Len(t.Text)
End Sub

Private Sub txtXMax_GotFocus()
SelectText txtXMax
End Sub

Private Sub txtXMax_LostFocus()
If Not IsNumeric(txtXMax.Text) Then
   MsgBox NumbersOnly, vbOKOnly, glFGTU
   txtXMax.SetFocus
End If
End Sub

Private Sub txtXMin_GotFocus()
SelectText txtXMin
End Sub

Private Sub txtXMin_LostFocus()
If Not IsNumeric(txtXMin.Text) Then
   MsgBox NumbersOnly, vbOKOnly, glFGTU
   txtXMin.SetFocus
End If
End Sub

Private Sub txtYMax_GotFocus()
SelectText txtYMax
End Sub

Private Sub txtYMax_LostFocus()
If Not IsNumeric(txtYMax.Text) Then
   MsgBox NumbersOnly, vbOKOnly, glFGTU
   txtYMax.SetFocus
End If
End Sub

Private Sub txtYMin_GotFocus()
SelectText txtYMin
End Sub

Private Sub txtYMin_LostFocus()
If Not IsNumeric(txtYMin.Text) Then
   MsgBox NumbersOnly, vbOKOnly, glFGTU
   txtYMin.SetFocus
End If
End Sub

Private Function ValidationOK() As Boolean
X1 = CDbl(txtXMin.Text)
x2 = CDbl(txtXMax.Text)
If CallerID <> OlyCallerID And CallerID >= 0 Then
    If X1 < 0 And Not GelBody(CallerID).csMyCooSys.csType = glNETCooSys Then
       MsgBox NonNegativeOnly, vbOKOnly, glFGTU
       txtXMin.SetFocus
       ValidationOK = False
       Exit Function
    End If
    If x2 < 0 And Not GelBody(CallerID).csMyCooSys.csType = glNETCooSys Then
       MsgBox NonNegativeOnly, vbOKOnly, glFGTU
       txtXMax.SetFocus
       ValidationOK = False
       Exit Function
    End If
End If

If X1 = x2 Then
   MsgBox MinMaxDifferent, vbOKOnly, glFGTU
   txtXMin.SetFocus
   ValidationOK = False
   Exit Function
End If
If x2 < X1 Then
   MsgBox MinLTMax, vbOKOnly, glFGTU
   txtXMin.SetFocus
   ValidationOK = False
   Exit Function
End If
Y1 = CDbl(txtYMin.Text)
If Y1 < 0 Then
   MsgBox NonNegativeOnly, vbOKOnly, glFGTU
   txtYMin.SetFocus
   ValidationOK = False
   Exit Function
End If
Y2 = CDbl(txtYMax.Text)
If Y2 < 0 Then
   MsgBox NonNegativeOnly, vbOKOnly, glFGTU
   txtYMax.SetFocus
   ValidationOK = False
   Exit Function
End If
If Y1 = Y2 Then
   MsgBox MinMaxDifferent, vbOKOnly, glFGTU
   txtYMin.SetFocus
   ValidationOK = False
   Exit Function
End If
If Y2 < Y1 Then
   MsgBox MinMaxDifferent, vbOKOnly, glFGTU
   txtYMin.SetFocus
   ValidationOK = False
   Exit Function
End If
ValidationOK = True
End Function

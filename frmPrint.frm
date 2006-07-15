VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSetFND 
      Caption         =   "In&clude file name and date with gel graph"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame fraWhat 
      Caption         =   "Print What"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1815
      Begin VB.CheckBox chkPrintWhat 
         Caption         =   "Gel &data set"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   540
         Width           =   1215
      End
      Begin VB.CheckBox chkPrintWhat 
         Caption         =   "Gel &file info"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkPrintWhat 
         Caption         =   "Gel &graph"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.Frame fraRange 
      Caption         =   "Scope"
      Height          =   1215
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1815
      Begin VB.OptionButton optRange 
         Caption         =   "&All open gels"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optRange 
         Caption         =   "A&ctive gel only"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblPrt 
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   4245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Printer:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 07/16/1999
Option Explicit

Dim lPrtRange As Long     '0 current, 1 all
Dim lPrtQuality As Long
Dim CallerID As Long  'index of the calling gel
Dim iCallType As Integer  '0 call from gel, 1 call from data, 2 call from data info

Private Sub chkSetFND_Click()
If chkSetFND.value = vbChecked Then
   bSetFileNameDate = True
Else
   bSetFileNameDate = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Long
Select Case lPrtRange
Case 0  'print current only
     If chkPrintWhat(0).value = vbChecked Then GelDrawPrinter CallerID
     If chkPrintWhat(1).value = vbChecked Then PrintData1 CallerID
     If chkPrintWhat(2).value = vbChecked Then PrintFileInfo CallerID, 3
Case 1  'print all
    For i = 1 To UBound(GelBody)
        If Not GelStatus(i).Deleted Then
           If chkPrintWhat(0).value = vbChecked Then GelDrawPrinter i
           If chkPrintWhat(1).value = vbChecked Then PrintData1 i
           If chkPrintWhat(2).value = vbChecked Then PrintFileInfo i, 3
        End If
    Next i
End Select
Unload Me
End Sub

Private Sub Form_Activate()
If IsNumeric(Me.Tag) Then
   iCallType = 0
   CallerID = Me.Tag
Else
   CallerID = Left(Me.Tag, Len(Me.Tag) - 1)
   If Right(Me.Tag, 1) = "D" Then 'file info
      iCallType = 1
   Else
      iCallType = 2
   End If
End If
Select Case iCallType
Case 0  'no restrictions
Case 1  'data
     optRange(1).Enabled = False
     chkPrintWhat(0).value = vbUnchecked
     chkPrintWhat(1).value = vbChecked
     chkPrintWhat(0).Enabled = False
     chkPrintWhat(2).Enabled = False
Case 2  'file info
     optRange(1).Enabled = False
     chkPrintWhat(0).value = vbUnchecked
     chkPrintWhat(2).value = vbChecked
     chkPrintWhat(0).Enabled = False
     chkPrintWhat(1).Enabled = False
End Select
lPrtRange = 0
lPrtQuality = -4 'high
bSetFileNameDate = True
bIncludeTextLabels = True

End Sub

Private Sub Form_Load()
lblPrt.Caption = Printer.DeviceName
End Sub

Private Sub optRange_Click(Index As Integer)
lPrtRange = Index
End Sub

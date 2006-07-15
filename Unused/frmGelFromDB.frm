VERSION 5.00
Begin VB.Form frmGelFromDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Gel From DB"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmGelFromDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox cmbSources 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGelFromDB.frx":000C
      Left            =   480
      List            =   "frmGelFromDB.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin VB.OptionButton optGelFomWhat 
      Caption         =   "Create Gel From ICR &Source"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.OptionButton optGelFomWhat 
      Caption         =   "Create Gel From &AMT Table"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.OptionButton optGelFomWhat 
      Caption         =   "Create Gel From O&RF Table"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   2535
   End
End
Attribute VB_Name = "frmGelFromDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 08/14/2000 nt
Option Explicit
Dim CallerID As Long

Dim Src() As String
Dim SrcID() As Long
Dim SrcCnt As Long

Private Sub cmdCancel_Click()
GelStatus(CallerID).DBGel = 0
If Not GelDB(CallerID) Is Nothing Then
   GelDB(CallerID).Close
   Set GelDB(CallerID) = Nothing
End If
Unload Me
End Sub

Private Sub cmdOK_Click()
If optGelFomWhat(0).value Then
   GelStatus(CallerID).DBGel = glDBGEL_ORF
ElseIf optGelFomWhat(1).value Then
   GelStatus(CallerID).DBGel = glDBGEL_AMT
ElseIf optGelFomWhat(2).value Then
   If cmbSources.ListIndex >= 0 Then
      GelStatus(CallerID).DBGel = SrcID(cmbSources.ListIndex)
   Else
      MsgBox "Select the source of gel data.", vbOKOnly
   End If
End If
Unload Me
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
'open database here
If Not ConnectToFTICR_AMT(GelDB(CallerID), GelData(CallerID).PathtoDatabase, False) Then
   MsgBox "Error accessing database file behind the gel; some functions might not be available.", vbOKOnly, "Loading ORF gel"
   cmdOK.Enabled = False
   Exit Sub
End If
Select Case LoadSources(CallerID, Src(), SrcID())
Case 0 'OK
   FillSourcesCombo
Case 1 'Source table not found
   MsgBox "Table FTSources missing in database " & GelDB(CallerID).Name & ".", vbOKOnly
   optGelFomWhat(2).Enabled = False
Case 2 'Source table empty
   MsgBox "No source found in database " & GelDB(CallerID).Name & ".", vbOKOnly
   optGelFomWhat(2).Enabled = False
Case Else
   optGelFomWhat(2).Enabled = False
End Select
End Sub

Private Sub FillSourcesCombo()
Dim i As Long
SrcCnt = UBound(Src) + 1
If SrcCnt > 0 Then
   For i = 0 To SrcCnt - 1
       cmbSources.AddItem Src(i), i
   Next i
End If
End Sub

Private Sub optGelFomWhat_Click(Index As Integer)
Select Case Index
Case 0, 1
     cmbSources.Enabled = False
Case 2
     cmbSources.Enabled = True
End Select
End Sub

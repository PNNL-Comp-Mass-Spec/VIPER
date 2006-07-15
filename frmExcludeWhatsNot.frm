VERSION 5.00
Begin VB.Form frmExcludeWhatsNot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exclusion - List"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmExcludeWhatsNot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtMMA 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "1"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TxtList 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdExclude 
      Caption         =   "Exclude"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Da from list elements."
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Exclude MW not within"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmExcludeWhatsNot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'simple exclusion function that lets user see different mass ranges
'last modified: 07/22/2002 nt
'------------------------------------------------------------------
Option Explicit

Dim CallerID As Long

Dim IncList() As Double
Dim ListCnt As Long

Private Sub FillList()
Dim FirstPos As Long
Dim LastPos As Long
Dim Done As Boolean
ReDim IncList(100)
ListCnt = 0

FirstPos = 1
LastPos = 0

Do Until Done
   LastPos = InStr(FirstPos, TxtList.Text, ";")
   If LastPos > 0 Then
      ListCnt = ListCnt + 1
      IncList(ListCnt) = CDbl(Mid$(TxtList.Text, FirstPos, LastPos - FirstPos))
      FirstPos = LastPos + 1
   Else
      If FirstPos < Len(TxtList) Then
         ListCnt = ListCnt + 1
         IncList(ListCnt) = CDbl(Right$(TxtList.Text, Len(TxtList.Text) - FirstPos + 1))
      End If
      Done = True
   End If
Loop
If ListCnt > 0 Then
   ReDim Preserve IncList(ListCnt)
Else
   Erase IncList
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExclude_Click()
Dim i As Long, j As Long
Dim IsoF As Integer
Dim AbsErr As Double
Dim BadCount As Long
IsoF = GelData(CallerID).Preferences.IsoDataField
FillList
If ListCnt > 0 Then
   AbsErr = txtMMA.Text
   With GelData(CallerID)
     If .CSLines > 0 Then
       For j = 1 To .CSLines
           BadCount = 0
           For i = 1 To ListCnt
             If Abs(.CSData(j).AverageMW - IncList(i)) > AbsErr Then
                BadCount = BadCount + 1
             End If
           Next i
           If BadCount = ListCnt Then
              GelDraw(CallerID).CSID(j) = -Abs(GelDraw(CallerID).CSID(j))
           End If
       Next j
     End If
     If .IsoLines > 0 Then
       For j = 1 To .IsoLines
           BadCount = 0
           For i = 1 To ListCnt
             If Abs(GetIsoMass(.IsoData(j), IsoF) - IncList(i)) > AbsErr Then
                BadCount = BadCount + 1
             End If
           Next i
           If BadCount = ListCnt Then
              GelDraw(CallerID).IsoID(j) = -Abs(GelDraw(CallerID).IsoID(j))
           Else
              DoEvents
           End If
        Next j
     End If
   End With
End If
End Sub


Private Sub cmdHelp_Click()
MsgBox "List semicolon delimited list of masses that you want to keep. Press Exclude button. Data out of specified ranges will be filtered out.", vbOKOnly, glFGTU
End Sub


Private Sub Form_Activate()
CallerID = Me.Tag
End Sub


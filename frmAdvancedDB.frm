VERSION 5.00
Begin VB.Form frmAdvancedDB 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced DB Settings"
   ClientHeight    =   3915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5325
   Icon            =   "frmAdvancedDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNewConnectString 
      Caption         =   "&New"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditConnectString 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtThingees 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox txtConnectString 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5160
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000004&
      X1              =   840
      X2              =   2400
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000004&
      X1              =   240
      X2              =   2400
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000004&
      X1              =   4320
      X2              =   4920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   3240
      X2              =   5040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   2280
      X2              =   5040
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   840
      X2              =   4920
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   2520
      X2              =   2520
      Y1              =   240
      Y2              =   3720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Access Properties"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Connect String"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmAdvancedDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'advanced database settings dialog; there is no problems with the
'modality of this form(it is controled from frmNewAnalysis instance
'last modified: 08/22/2001 nt
'-----------------------------------------------------
Option Explicit

Dim Loading As Boolean                    'loading indicator flag
Dim StuffChanged As Boolean               'indicates if stuff text changed

'public properties
Public AcceptChanges As Boolean
Public ThisCN As ADODB.Connection         'working connection
Public ThisStuff As New Collection        'working stuff

Private Sub cmdCancel_Click()
'----------------------------------------------------------
'keep old (do nothing if user Cancels changes)
'----------------------------------------------------------
AcceptChanges = False
Me.Hide
End Sub

Private Sub cmdEditConnectString_Click()
Dim dl As New DataLinks
If dl.PromptEdit(ThisCN) Then    'User selected OK
   txtConnectString.Text = ThisCN.ConnectionString
End If
End Sub

Private Sub cmdNewConnectString_Click()
'----------------------------------------------------
'asks for new connection string for selected database
'----------------------------------------------------
Dim dl As New DataLinks
Dim NewCn As ADODB.Connection
Set NewCn = dl.PromptNew()
If Not NewCn Is Nothing Then
   Set ThisCN = NewCn
   txtConnectString.Text = ThisCN.ConnectionString
End If
End Sub

Private Sub cmdOK_Click()
'---------------------------------------------
'mark that changes should be accepted and hide
'---------------------------------------------
'maybe user typed things; check it out
If StuffChanged Then StuffRebuild
ThisCN.ConnectionString = Trim$(txtConnectString)
AcceptChanges = True
Me.Hide
End Sub

Private Sub Form_Activate()
'-------------------------------------------
'loads properties from ThisCN and ThisStuff
'-------------------------------------------
Dim i As Long
Dim Thingees As String
On Error Resume Next
txtConnectString.Text = ThisCN.ConnectionString
If Loading Then
   For i = 1 To ThisStuff.Count
       Thingees = Thingees & ThisStuff.Item(i).Name & MyGl.INIT_Value & ThisStuff.Item(i).Value & vbCrLf
   Next
   txtThingees.Text = Thingees
   StuffChanged = False
   Loading = False
End If
End Sub

Private Sub Form_Load()
Loading = True
End Sub

Private Sub txtThingees_Change()
StuffChanged = True
End Sub

Private Sub StuffRebuild()
'---------------------------------------------------------
'rebuilds ThisStuff collection from the Thingees text box
'---------------------------------------------------------
Dim i As Long
On Error Resume Next
Dim PairsCount As Long
Dim MyNames() As String
Dim MyValues() As String
Dim nv As NameValue

'remove whatever is in it
EmptyCollection ThisStuff
'rebuild from text box
PairsCount = GetNamesValues(txtThingees.Text, MyNames(), MyValues())
If PairsCount > 0 Then
   Set ThisStuff = New Collection
   For i = 1 To PairsCount
       If Len(MyNames(i - 1)) > 0 Then
          Set nv = New NameValue
          nv.Name = MyNames(i - 1)
          nv.Value = MyValues(i - 1)
          ThisStuff.Add nv, MyNames(i - 1)
       End If
   Next i
End If
End Sub

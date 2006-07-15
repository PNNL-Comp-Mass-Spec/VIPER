VERSION 5.00
Begin VB.Form frmChart3DOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graph Options"
   ClientHeight    =   3375
   ClientLeft      =   2325
   ClientTop       =   1950
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset to &Defaults"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdCloseOptions 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOKOptions 
      Caption         =   "&Change"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame frameOptions 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.TextBox inputFontSize 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox inputPerspective 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox inputElevation 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox inputRotation 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label labelFontSize 
         Caption         =   "&Label Size:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label labelPerspective 
         Caption         =   "&Perspective:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label labelElevation 
         Caption         =   "&Elevation:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label labelRotation 
         Caption         =   "&Rotation:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmChart3DOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public goClose As Boolean

Private Sub cmdCloseOptions_Click()
goClose = True
Me.Hide
End Sub

Private Sub cmdOKOptions_Click()
Me.Hide
End Sub

Private Sub cmdReset_Click()
inputRotation = "225"
inputElevation = "30"
inputPerspective = "2.5"
inputFontSize = "80"
End Sub

Private Sub Form_Activate()
goClose = False
End Sub

Private Sub Form_Load()
    Me.width = 2535
    Me.Height = 3900
    
End Sub

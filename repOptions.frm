VERSION 5.00
Begin VB.Form frmErrorDistribution3DReportOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Options"
   ClientHeight    =   4170
   ClientLeft      =   6525
   ClientTop       =   3495
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Include in Report"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox inputComments 
         Height          =   495
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2520
         Width           =   3255
      End
      Begin VB.CheckBox chkCumCountsCol 
         Caption         =   "Cumulative Counts (Columns)"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox chkCumCountsTable 
         Caption         =   "Cumulative Counts (Table)"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkCountsCol 
         Caption         =   "Counts (Columns)"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkCountsTable 
         Caption         =   "Counts (Table)"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkHeader 
         Caption         =   "Include Header Information"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Label labelComments 
         Caption         =   "Comments:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmErrorDistribution3DReportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public roCancel As Boolean

Private Sub cmdCancel_Click()
roCancel = True
Me.Hide
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
roCancel = False
End Sub

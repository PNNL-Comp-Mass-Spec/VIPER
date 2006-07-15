VERSION 5.00
Begin VB.Form frmComment 
   Caption         =   "Comment"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtComment 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 02/11/2003 nt
Option Explicit

' This form can be used to edit the gel comment
' Alternatively, it can simply be used to display any long text the programmer wishes to display
' For the second mode, set frmComment.Tag to a value < 0

Private CallerID As Long
Private OldComment As String

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    
    With txtComment
        .Left = 0
        .Top = 0
        .width = Me.ScaleWidth
        lngDesiredValue = Me.ScaleHeight - cmdOK.Height - 360
        If lngDesiredValue < 1000 Then lngDesiredValue = 1000
        .Height = lngDesiredValue
    End With
    
    With cmdCancel
        .Top = txtComment.Top + txtComment.Height + 120
        lngDesiredValue = Me.ScaleWidth - cmdCancel.width - cmdOK.width - 240
        If lngDesiredValue < 120 Then lngDesiredValue = 120
        
        .Left = lngDesiredValue
    
        cmdOK.Top = .Top
        cmdOK.Left = .Left + .width + 120
    End With

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If CallerID >= 1 Then
    GelData(CallerID).Comment = Trim(txtComment.Text)
    If OldComment <> GelData(CallerID).Comment Then
       GelStatus(CallerID).Dirty = True
    End If
End If
Unload Me
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
If CallerID >= 1 Then
    OldComment = GelData(CallerID).Comment
    txtComment.Text = OldComment
End If
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

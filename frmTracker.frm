VERSION 5.00
Begin VB.Form frmTracker 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Coordinates"
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780.278
   ScaleMode       =   0  'User
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtIdentity 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   790
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   12
      Tag             =   "4"
      Top             =   1019
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label lblUMCIndex 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Tag             =   "2"
      ToolTipText     =   "UMC Index"
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UMC Index"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "Drag and drop form elsewhere"
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label lblIdentity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Tag             =   "3"
      Top             =   510
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Expr. Ratio"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Drag and drop form elsewhere"
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Intensity"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Double click to move to default location"
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MW"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Tag             =   "1"
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "m/z"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Tag             =   "1"
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scan Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Double click to hide"
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label lblDRTrack 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Expression ratio"
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label lblAbuTrack 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Intensity"
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label lblMWTrack 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Molecular mass (Dalton)"
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label lblMOverZTrack 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "pI number"
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label lblFNTrack 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Scan number"
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label lblFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Tag             =   "4"
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'always on top form to track coordinates of the gel spots
'--------------------------------------------------------
'last modified: 03/29/2002 nt
Option Explicit

Private Const LABEL_INDEX_MZ = 1
Private Const LABEL_INDEX_SCAN_NUMBER = 2
Private Const LABEL_INDEX_MW = 3
Private Const LABEL_INDEX_INTENSITY = 4
Private Const LABEL_INDEX_ER = 5
Private Const LABEL_INDEX_UMC = 6

Dim IdentityState      '0 normal;              1 no labels
                       '2 text box disabled;   3 text box enabled
Dim OrigIdentityTop
Dim OrigFNTop
Dim FineLineWidth

Dim MouseRight As Boolean

Private mIntensityNotationMode As nmNotationModeConstants
Private mUMCIndexExpanded As Boolean

' Unused code
'Private Sub Form_Activate()
''no need to update gel when loses focus over this form
'glUpdateGel = False
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'unload if Ctrl+T combination
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
If (CtrlDown And KeyCode = vbKeyT) Then
    Unload Me
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim Res As Long
Dim lTop As Long
Dim lLeft As Long
Dim lFlags

lTop = GetSystemMetrics(SM_CYMENU) \ 2
lLeft = GetSystemMetrics(SM_CXSCREEN) - Me.width / Screen.TwipsPerPixelX - 120
lFlags = SWP_NOSIZE
Res = SetWindowPos(Me.hwnd, HWND_TOPMOST, lLeft, lTop, 0, 0, lFlags)
glTracking = True
IdentityState = 0
OrigIdentityTop = lblIdentity.Top
OrigFNTop = lblFNTrack.Top
FineLineWidth = 2 * lblFNTrack.Height - (lblIdentity.Top + lblIdentity.Height - OrigFNTop)

' Set some ToolTips
lblIdentity.ToolTipText = "Double-click to expand.  Double right-click to expand further."
Label2(2).ToolTipText = "Double-click to close."
Label2(4).ToolTipText = "Double-click to toggle scientific and decimal notation"

End Sub

Private Sub Form_Unload(Cancel As Integer)
glTracking = False
SyncMenuCmdTracker False
End Sub

Public Function GetIntensityNotationMode() As nmNotationModeConstants
    ' Used to return the current intensity notation mode (scientific or decimal)
    GetIntensityNotationMode = mIntensityNotationMode
End Function

Private Sub Label2_DblClick(Index As Integer)
Select Case Index
Case LABEL_INDEX_SCAN_NUMBER  'scan number
    Unload Me
Case LABEL_INDEX_INTENSITY  'intensity
    Form_Load
    ToggleIntensityNotation
Case LABEL_INDEX_UMC  'UMC Index
    ToggleUMCIndexWidth
End Select
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Res As Long
'this allows to drag and drop form
If Button = vbLeftButton Then
    ReleaseCapture
    Res = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub lblIdentity_DblClick()
On Error Resume Next

If txtIdentity.Visible Then    'only possible in this event when IdentityState= 2
   IdentityState = 3
Else
   If MouseRight Then
      IdentityState = 2
   ElseIf IdentityState = 0 Then
      IdentityState = 1
   ElseIf IdentityState = 1 Then
      IdentityState = 0
   End If
End If
TrackerState
End Sub

Private Sub lblIdentity_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Button
Case vbLeftButton
     MouseRight = False
Case vbRightButton
     MouseRight = True
End Select
End Sub

Private Sub txtIdentity_DblClick()
'only possible when IdentityState=3
IdentityState = 0
TrackerState
End Sub

Private Sub ToggleIntensityNotation()
    If mIntensityNotationMode = nmScientific Then
        mIntensityNotationMode = nmDecimal
    Else
        mIntensityNotationMode = nmScientific
    End If
End Sub

Private Sub ToggleUMCIndexWidth()
    mUMCIndexExpanded = Not mUMCIndexExpanded
    
    With Label2(LABEL_INDEX_UMC)
        If mUMCIndexExpanded Then
            .width = frmTracker.ScaleWidth - Label2(LABEL_INDEX_INTENSITY).Left
        Else
            .width = 1095
        End If
        .Left = frmTracker.ScaleWidth - .width
    
        lblUMCIndex.Left = .Left
        lblUMCIndex.width = .width
    End With
    
End Sub

Public Sub TrackerState()
Dim c As Control
Select Case IdentityState
Case 0
   txtIdentity.Visible = False
   For Each c In Me.Controls
       Select Case c.Tag
       Case 1       'first row
            c.Visible = True
       Case 2       'second row
            c.Top = OrigFNTop
       Case 3       'third row
            c.Top = OrigIdentityTop
            c.Height = Label2(1).Height
       End Select
   Next
   lblIdentity.ToolTipText = "Double-click to expand.  Double right-click to expand further."
Case 1
   txtIdentity.Visible = False
   For Each c In Me.Controls
       Select Case c.Tag
       Case 1       'first row
            c.Visible = False
       Case 2       'second row
            c.Top = Label2(1).Top
       Case 3       'third row
            c.Top = OrigFNTop
            c.Height = 2 * Label2(1).Height - FineLineWidth
       End Select
   Next
   lblIdentity.ToolTipText = "Double-click to restore.  Double right-click to expand further."
Case 2
   txtIdentity.Visible = True
   txtIdentity.Top = 0
   txtIdentity.Enabled = False
   lblIdentity.ToolTipText = "Double-click to allow scrolling. Ctrl+T to hide tracker, then Ctrl+T restore."
Case 3
   txtIdentity.Visible = True
   txtIdentity.Top = 0
   txtIdentity.Enabled = True
   lblIdentity.ToolTipText = "Double-click to restore."
End Select
End Sub


VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOverlayUMCManager 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Overlay Manager - UMC"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4260
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   4260
      ScaleWidth      =   4695
      TabIndex        =   14
      Top             =   480
      Width           =   4695
      Begin VB.Frame fraGrid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grid"
         Height          =   2775
         Left            =   2400
         TabIndex        =   35
         Top             =   120
         Width           =   2175
         Begin VB.OptionButton optHGridAutoMode 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Const."
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   53
            Top             =   2430
            Width           =   855
         End
         Begin VB.OptionButton optHGridAutoMode 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Const."
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   52
            Top             =   2070
            Width           =   855
         End
         Begin VB.OptionButton optVGridAutoMode 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Const."
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   51
            ToolTipText     =   "Constant bin count"
            Top             =   1380
            Width           =   855
         End
         Begin VB.OptionButton optVGridAutoMode 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Const."
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   50
            ToolTipText     =   "Constant width"
            Top             =   1020
            Width           =   855
         End
         Begin VB.TextBox txtHGridBinsCnt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            TabIndex        =   49
            Text            =   "0"
            ToolTipText     =   "Count of horizontal grid bins"
            Top             =   2400
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtVGridBinsCnt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            TabIndex        =   48
            Text            =   "0"
            ToolTipText     =   "Count of vertical grid bins"
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtVGridWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            TabIndex        =   42
            Text            =   "0"
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtHGridWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            TabIndex        =   40
            Text            =   "0"
            Top             =   2040
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cmbGridlineStyle 
            Height          =   315
            ItemData        =   "frmOverlayUMCManager.frx":0000
            Left            =   720
            List            =   "frmOverlayUMCManager.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkShowHGrid 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show"
            Height          =   195
            Left            =   1200
            TabIndex        =   37
            Top             =   1800
            Width           =   735
         End
         Begin VB.CheckBox chkShowVGrid 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show"
            Height          =   195
            Left            =   1200
            TabIndex        =   36
            Top             =   720
            Width           =   735
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   120
            X2              =   2040
            Y1              =   1690
            Y2              =   1690
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   120
            X2              =   2040
            Y1              =   630
            Y2              =   630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Count"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   47
            Top             =   2440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Count"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   46
            Top             =   1360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Horizontal"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   45
            Top             =   1800
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Vertical"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   43
            Top             =   1020
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   41
            Top             =   2100
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Line"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   39
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.TextBox txtNETStickWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         Text            =   "0"
         Top             =   1740
         Width           =   615
      End
      Begin VB.Frame fraOrientation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Orientation"
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   2175
         Begin VB.OptionButton optOrientation 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Horz.MW/Vert.NET"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "Horizontal axis is MW; vertical NET"
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optOrientation 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Horz.NET/Vert.MW"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Horizontal axis is NET; vertical MW"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtMaxSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "0"
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtMinSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Text            =   "0"
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NET Stick Wd.(%)"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Min.sz. (% Log.)"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max.sz. (% Log.)"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblForeClr 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         ToolTipText     =   "Double click to change"
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ForeColor"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblBackClr 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         ToolTipText     =   "Double click to change"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4260
      Index           =   1
      Left            =   240
      ScaleHeight     =   4260
      ScaleWidth      =   4695
      TabIndex        =   7
      Top             =   480
      Width           =   4695
      Begin VB.TextBox txtComment 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   3000
         Width           =   4455
      End
      Begin VB.CheckBox chkShowText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Text"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Frame fraChangeName 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2130
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   2505
         Begin VB.TextBox txtNewName 
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Top             =   150
            Width           =   1695
         End
         Begin VB.CommandButton cmdAcceptNameChange 
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1860
            Picture         =   "frmOverlayUMCManager.frx":0046
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   150
            Width           =   255
         End
         Begin VB.CommandButton cmdCancelNameChange 
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2140
            Picture         =   "frmOverlayUMCManager.frx":01D0
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Cancel name change"
            Top             =   150
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdChangeName 
         Caption         =   "Change Name"
         Height          =   315
         Left            =   2280
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkUniSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Uniform size"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cmbShape 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1360
         Width           =   1695
      End
      Begin VB.ListBox lstOverlays 
         Height          =   2400
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Double-click to change name"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblType 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2280
         TabIndex        =   72
         ToolTipText     =   "type"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblClr 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         ToolTipText     =   "Double click to change"
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Shape"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Overlaid displays (in z-order)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3060
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "&Create"
      Height          =   315
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdCommands 
      Caption         =   "&Close"
      Height          =   315
      Index           =   0
      Left            =   4400
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4260
      Index           =   0
      Left            =   240
      ScaleHeight     =   4260
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   480
      Width           =   4695
      Begin VB.Frame fraMiscOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Miscellaneous"
         Height          =   1335
         Left            =   120
         TabIndex        =   68
         Top             =   2760
         Width           =   4455
         Begin VB.CheckBox chkDefCurrVisible 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Make current display visible"
            Height          =   255
            Left            =   360
            TabIndex        =   71
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkDefCreateWithID 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Create with ID"
            Height          =   255
            Left            =   360
            TabIndex        =   70
            ToolTipText     =   "If checked ID from display will be used"
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox chkDefUniSize 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Uniform size"
            Height          =   255
            Left            =   360
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraColorsShapes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Colors/Shapes"
         Height          =   1815
         Left            =   2640
         TabIndex        =   63
         Top             =   840
         Width           =   1935
         Begin VB.ComboBox cmbDefShape 
            Height          =   315
            ItemData        =   "frmOverlayUMCManager.frx":035A
            Left            =   120
            List            =   "frmOverlayUMCManager.frx":035C
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Default shape:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblDefClr 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1200
            TabIndex        =   65
            ToolTipText     =   "Double click to change"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Default color:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   64
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame fraNETAdjustment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NET Adjustment"
         Height          =   1815
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   2415
         Begin VB.TextBox txtMinNET 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            TabIndex        =   60
            Text            =   "0"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtMaxNET 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            TabIndex        =   59
            Text            =   "1"
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox cmbDefNETAdjustment 
            Height          =   315
            ItemData        =   "frmOverlayUMCManager.frx":035E
            Left            =   240
            List            =   "frmOverlayUMCManager.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Overlay min NET:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   62
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Overlay max NET:"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   61
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Type:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   57
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UMC"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   55
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Individual peaks"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   54
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox cmbDisplayList 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   400
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Create new overlay from:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8281
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Creator"
            Key             =   "C"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Editor"
            Key             =   "E"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc"
            Key             =   "M"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOverlayUMCManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'overlays creator and editor function
'created: 08/28/2002 nt
'last modified: 12/12/2002 nt
'-----------------------------------------------------------------
Option Explicit

Const GRID_STYLE_SOLID = 0
Const GRID_STYLE_DASH = 1
Const GRID_STYLE_DOT = 2
Const GRID_STYLE_DASH_DOT = 3
Const GRID_STYLE_DASH_DOT_DOT = 4

Const TAB_CREATOR = 1
Const TAB_EDITOR = 2
Const TAB_MISC = 3

Const CMD_CLOSE = 0
Const CMD_CREATE = 1
Const CMD_DELETE = 2

Const DB_NOT_LOADED = "DB Display Not Loaded"

Dim ActiveTab As Long

Dim MyZOrder As New ZOrder          'z-order is 0-based
Dim CurrZPos As Long                'current z-order position(selected from the list
Dim CurrOlyInd As Long           'and coresponding oly UMC index

' MonroeFix
'Private Sub chkCreateWithID_Click()
'OlyOptions.DefWithID = (chkCreateWithID.value = vbChecked)
'End Sub

Private Sub chkDefUniSize_Click()
OlyOptions.DefUniformSize = (chkDefUniSize.value = vbChecked)
End Sub

Private Sub chkShowText_Click()
On Error Resume Next
Oly(CurrOlyInd).ShowText = (chkShowText.value = vbChecked)
End Sub

Private Sub chkUniSize_Click()
On Error Resume Next
Oly(CurrOlyInd).UniformSize = (chkUniSize.value = vbChecked)
End Sub

Private Sub cmbDefNETAdjustment_Click()
OlyOptions.DefNETAdjustment = cmbDefNETAdjustment.ListIndex
End Sub

Private Sub cmbDefShape_Click()
OlyOptions.DefShape = cmbDefShape.ListIndex
End Sub

Private Sub cmbGridlineStyle_Click()
Select Case cmbGridlineStyle.ListIndex
Case glsSOLID
     OlyOptions.GRID.LineStyle = PS_SOLID
Case glsDASH
     OlyOptions.GRID.LineStyle = PS_DASH
Case glsDOT
     OlyOptions.GRID.LineStyle = PS_DOT
Case glsDASHDOT
     OlyOptions.GRID.LineStyle = PS_DASHDOT
Case glsDASHDOTDOT
     OlyOptions.GRID.LineStyle = PS_DASHDOTDOT
Case Else
     Exit Sub
End Select
CreateOlyForeClrObject OlyOptions.ForeColor
End Sub


Private Sub cmbShape_Click()
If CurrZPos >= 0 Then Oly(CurrOlyInd).Type = cmbShape.ListIndex
End Sub


Private Sub cmdAcceptNameChange_Click()
'--------------------------------------------------------------
'accept new name(if acceptable) and hide
'--------------------------------------------------------------
On Error Resume Next
If Len(Trim$(txtNewName.Text)) > 0 Then
   Oly(CurrOlyInd).Name = Trim$(txtNewName.Text)
   Call FillOverlaidList
   fraChangeName.Visible = False
Else
   MsgBox "Type new name!", vbOKOnly, glFGTU
   txtNewName.SetFocus
End If
End Sub

Private Sub cmdCancelNameChange_Click()
fraChangeName.Visible = False
End Sub

Private Sub cmdChangeName_Click()
fraChangeName.Visible = True
End Sub

Private Sub cmdCommands_Click(Index As Integer)
Dim CurrDisplayInd As Long
Dim Resp As Long
On Error Resume Next
Select Case Index
Case CMD_CLOSE
     Unload Me
Case CMD_CREATE
     CurrDisplayInd = cmbDisplayList.ListIndex
     If CurrDisplayInd >= 0 Then
        If MyZOrder.GetZOrder(CurrDisplayInd) >= 0 Then
           Resp = MsgBox("Selected display is already overlaid! Do you want to create another? (Hint: No)", vbYesNo, glFGTU)
           If Resp <> vbYes Then Exit Sub
        End If
        Me.MousePointer = vbHourglass
        If AddDisplayToOverlay(CurrDisplayInd) Then
           MyZOrder.AddInZOrder CurrDisplayInd, OlyCnt - 1
           FillOverlaidList
        Else
           MsgBox "Error creating overlay!", vbOKOnly, glFGTU
        End If
        Me.MousePointer = vbDefault
     Else
        MsgBox "Select display to overlay!", vbOKOnly
     End If
Case CMD_DELETE
     If CurrZPos >= 0 Then
        If RemoveZOrderPositionFromOverlay(CurrZPos) Then
           MyZOrder.RemoveFromZOrder CurrZPos
           FillOverlaidList
        Else
           MsgBox "Error removing overlay from the list!", vbOKOnly, glFGTU
        End If
     Else
        MsgBox "Select display overlay you want to remove!", vbOKOnly, glFGTU
     End If
     Call ClearEditorControls
End Select
End Sub

Private Sub Form_Load()
CurrZPos = -1
FillTypeCombos
FillDisplaysCombo
FillNETAdjustmentCombo
cmbDefShape.ListIndex = OlyOptions.DefShape
cmbDefNETAdjustment.ListIndex = OlyOptions.DefNETAdjustment
optType(OlyOptions.DefType).value = True
lblDefClr.BackColor = OlyOptions.DefColor
lblBackClr.BackColor = OlyOptions.BackColor
lblForeClr.BackColor = OlyOptions.ForeColor
txtMinNET.Text = OlyOptions.DefMinNET
txtMaxNET.Text = OlyOptions.DefMaxNET
txtMinSize.Text = OlyOptions.MinSize
txtMaxSize.Text = OlyOptions.MaxSize
' MonroeFix txtStickWidth.Text = OlyOptions.StickWidth
If OlyOptions.DefUniformSize Then
   chkDefUniSize.value = vbChecked
Else
   chkDefUniSize.value = vbUnchecked
End If
If OlyOptions.GRID.VertGridVisible Then
   chkShowVGrid.value = vbChecked
Else
   chkShowVGrid.value = vbUnchecked
End If
If OlyOptions.GRID.HorzGridVisible Then
   chkShowHGrid.value = vbChecked
Else
   chkShowHGrid.value = vbUnchecked
End If
Select Case OlyOptions.GRID.LineStyle
Case PS_SOLID
     cmbGridlineStyle.ListIndex = glsSOLID
Case PS_DOT
     cmbGridlineStyle.ListIndex = glsDOT
Case PS_DASHDOT
     cmbGridlineStyle.ListIndex = glsDASHDOT
Case PS_DASHDOTDOT
     cmbGridlineStyle.ListIndex = glsDASHDOTDOT
Case Else            'default is PS_DASH
     cmbGridlineStyle.ListIndex = glsDASH
End Select
optOrientation(OlyOptions.Orientation).value = True
If OlyCnt > 0 Then
   If InitZOrderFromOly() Then FillOverlaidList
End If
' MonroeFix
'If OlyOptions.DefWithID Then
'   chkCreateWithID.value = vbChecked
'Else
'   chkCreateWithID.value = vbUnchecked
'End If
optHGridAutoMode(OlyOptions.GRID.HorzAutoMode).value = True
optVGridAutoMode(OlyOptions.GRID.VertAutoMode).value = True
txtHGridBinsCnt.Text = OlyOptions.GRID.HorzBinsCount
txtVGridBinsCnt.Text = OlyOptions.GRID.VertBinsCount
End Sub


Private Sub lblBackClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblBackClr_DblClick
TmpClr = lblBackClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblBackClr.BackColor = TmpClr
OlyOptions.BackColor = TmpClr
exit_lblBackClr_DblClick:
End Sub

Private Sub lblClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblClr_DblClick
TmpClr = lblClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblClr.BackColor = TmpClr
If CurrZPos >= 0 Then
   Oly(CurrOlyInd).Color = TmpClr
   AddEditOlyClr CurrOlyInd, Oly(CurrOlyInd).Color  'will not execute if color dialog canceled
End If
exit_lblClr_DblClick:
End Sub

Private Sub lblDefClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblDefClr_DblClick
TmpClr = lblDefClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblDefClr.BackColor = TmpClr
OlyOptions.DefColor = TmpClr
exit_lblDefClr_DblClick:
End Sub


Private Sub lblForeClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblForeClr_DblClick
TmpClr = lblForeClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblForeClr.BackColor = TmpClr
OlyOptions.ForeColor = TmpClr
exit_lblForeClr_DblClick:
End Sub

Private Sub lstOverlays_Click()
On Error Resume Next
CurrZPos = lstOverlays.ListIndex
CurrOlyInd = GetOlyIndFromZOrder(CurrZPos)
If CurrZPos >= 0 Then
   cmbShape.ListIndex = Oly(CurrOlyInd).Shape
   lblClr.BackColor = Oly(CurrOlyInd).Color
   Select Case Oly(CurrOlyInd).Type
   Case olySolo
        lblType.Caption = "Type: individual spots"
   Case olyUMC
        lblType.Caption = "Type: unique mass classes"
   End Select
   txtComment.Text = Oly(CurrOlyInd).Comment
   txtNewName.Text = Oly(CurrOlyInd).Name
   If Oly(CurrOlyInd).UniformSize Then
      chkUniSize.value = vbChecked
   Else
      chkUniSize.value = vbUnchecked
   End If
   If Oly(CurrOlyInd).ShowText Then
      chkShowText.value = vbChecked
   Else
      chkShowText.value = vbUnchecked
   End If
End If
End Sub

Private Sub optHGridAutoMode_Click(Index As Integer)
Select Case Index
Case gamWidthConst
     txtHGridWidth.Enabled = True
     txtHGridBinsCnt.Enabled = False
     OlyOptions.GRID.VertAutoMode = Index
' MonroeFix
Case gamBinsCntConst
     txtHGridWidth.Enabled = False
     txtHGridBinsCnt.Enabled = True
     OlyOptions.GRID.VertAutoMode = Index
End Select
End Sub

Private Sub optOrientation_Click(Index As Integer)
OlyOptions.Orientation = Index
End Sub

Private Sub optType_Click(Index As Integer)
OlyOptions.DefType = Index
End Sub

Private Sub optVGridAutoMode_Click(Index As Integer)
Select Case Index
Case gamWidthConst
     txtVGridWidth.Enabled = True
     txtVGridBinsCnt.Enabled = False
     OlyOptions.GRID.HorzAutoMode = Index
' MonroeFix
Case gamBinsCntConst
     txtVGridWidth.Enabled = False
     txtVGridBinsCnt.Enabled = True
     OlyOptions.GRID.HorzAutoMode = Index
End Select
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer
ActiveTab = TabStrip1.SelectedItem.Index
For i = 0 To TabStrip1.Tabs.Count - 1
    If i = TabStrip1.SelectedItem.Index - 1 Then
       Picture1(i).Left = 240
       Picture1(i).Enabled = True
    Else
       Picture1(i).Left = -20000
       Picture1(i).Enabled = False
    End If
Next i
Select Case ActiveTab
Case TAB_CREATOR
     cmdCommands(CMD_CREATE).Enabled = True
     cmdCommands(CMD_DELETE).Enabled = False
Case TAB_EDITOR
     cmdCommands(CMD_CREATE).Enabled = False
     cmdCommands(CMD_DELETE).Enabled = True
Case TAB_MISC
     cmdCommands(CMD_CREATE).Enabled = False
     cmdCommands(CMD_DELETE).Enabled = False
End Select
End Sub

Private Sub txtHGridBinsCnt_LostFocus()
If IsNumeric(txtHGridBinsCnt.Text) Then
   OlyOptions.GRID.HorzBinsCount = Abs(CLng(txtHGridBinsCnt.Text))
   txtHGridBinsCnt.Text = OlyOptions.GRID.HorzBinsCount
   txtHGridWidth.Text = OlyOptions.GRID.HorzWidth
Else
   MsgBox "This argument should be positive number!", vbOKOnly, glFGTU
   txtHGridBinsCnt.SetFocus
End If
End Sub

Private Sub txtMaxNET_LostFocus()
If IsNumeric(txtMaxNET.Text) Then
   OlyOptions.DefMaxNET = CDbl(txtMaxNET.Text)
Else
   MsgBox "This argument should be numeric!", vbOKOnly, glFGTU
   txtMaxNET.SetFocus
End If
End Sub


Private Sub txtMaxSize_LostFocus()
If IsNumeric(txtMaxSize.Text) Then
   OlyOptions.MaxSize = Abs(CSng(txtMaxSize.Text))
   txtMaxSize.Text = OlyOptions.MaxSize
Else
   MsgBox "This argument should integer 1-50!", vbOKOnly, glFGTU
   txtMaxSize.SetFocus
End If
End Sub

Private Sub txtMinNET_LostFocus()
If IsNumeric(txtMinNET.Text) Then
   OlyOptions.DefMinNET = CDbl(txtMinNET.Text)
Else
   MsgBox "This argument should be numeric!", vbOKOnly, glFGTU
   txtMinNET.SetFocus
End If
End Sub


Private Sub txtMinSize_LostFocus()
If IsNumeric(txtMinSize.Text) Then
   OlyOptions.MinSize = Abs(CSng(txtMinSize.Text))
   txtMinSize.Text = OlyOptions.MinSize
Else
   MsgBox "This argument should integer 1-50!", vbOKOnly, glFGTU
   txtMinSize.SetFocus
End If
End Sub

Private Sub txtHGridWidth_LostFocus()
If IsNumeric(txtHGridWidth.Text) Then
   OlyOptions.GRID.HorzWidth = Abs(CSng(txtHGridWidth.Text))
   txtHGridWidth.Text = OlyOptions.GRID.HorzWidth
   txtHGridBinsCnt.Text = OlyOptions.GRID.HorzBinsCount
Else
   MsgBox "This argument should be positive number!", vbOKOnly, glFGTU
   txtHGridWidth.SetFocus
End If
End Sub


Private Sub txtVGridBinsCnt_LostFocus()
If IsNumeric(txtVGridBinsCnt.Text) Then
   OlyOptions.GRID.VertBinsCount = Abs(CLng(txtVGridBinsCnt.Text))
   txtVGridBinsCnt.Text = OlyOptions.GRID.VertBinsCount
   txtVGridWidth.Text = OlyOptions.GRID.VertWidth
Else
   MsgBox "This argument should be positive number!", vbOKOnly, glFGTU
   txtVGridBinsCnt.SetFocus
End If
End Sub

Private Sub txtVGridWidth_LostFocus()
If IsNumeric(txtVGridWidth.Text) Then
   OlyOptions.GRID.VertWidth = Abs(CSng(txtVGridWidth.Text))
   txtVGridWidth.Text = OlyOptions.GRID.VertWidth
   txtVGridBinsCnt.Text = OlyOptions.GRID.VertBinsCount
Else
   MsgBox "This argument should be positive number!", vbOKOnly, glFGTU
   txtVGridWidth.SetFocus
End If
End Sub

' MonroeFix
'Private Sub txtStickWidth_LostFocus()
''--------------------------------------------------------------------------
''this is not very kosher(does not change width for all) but could be useful
''--------------------------------------------------------------------------
'On Error Resume Next
'If IsNumeric(txtStickWidth.Text) Then
'   OlyOptions.StickWidth = Abs(CSng(txtStickWidth.Text))
'   txtStickWidth.Text = OlyOptions.StickWidth
'   AddEditOlyClr CurrOlyInd, Oly(CurrOlyInd).Color
'Else
'   MsgBox "This argument should be positive number!", vbOKOnly, glFGTU
'   txtStickWidth.SetFocus
'End If
'End Sub


Private Sub FillDisplaysCombo()
Dim i As Long
cmbDisplayList.Clear
For i = 0 To UBound(GelBody)
    cmbDisplayList.AddItem GelBody(i).Caption, i
Next i
End Sub


Private Sub FillOverlaidList()
'------------------------------------------------
'fill list of overlaid displays in z-order
'------------------------------------------------
Dim i As Long
Dim ZOrderOlyInd As Long
lstOverlays.Clear
For i = 0 To OlyCnt - 1
    ZOrderOlyInd = GetOlyIndFromZOrder(i)
    If ZOrderOlyInd >= 0 Then lstOverlays.AddItem Oly(ZOrderOlyInd).Name, i
Next i
End Sub


Private Function InitZOrderFromOly() As Boolean
'------------------------------------------------------------------------------
'put overlaids in current z-order; this reverse indexing could probably be used
'in most cases; sometimes we will have to sort it though
'------------------------------------------------------------------------------
Dim i As Long
Dim TmpZOrder() As Long
On Error Resume Next
ReDim TmpZOrder(OlyCnt - 1)
For i = 0 To OlyCnt - 1
    TmpZOrder(Oly(i).ZOrder) = Oly(i).DisplayInd
Next i
If MyZOrder.AddInZOrderAll(TmpZOrder) Then InitZOrderFromOly = True
End Function


Private Function UpdateZOrderInOly() As Boolean
'---------------------------------------------------------------------
'updates z-order member of overlay structures based on current z-order
'---------------------------------------------------------------------
Dim i As Long
On Error Resume Next
For i = 0 To OlyCnt - 1
    Oly(i).ZOrder = MyZOrder.GetZOrder(i)
Next i
End Function

Private Sub FillTypeCombos()
cmbShape.AddItem "Box", olyBox
cmbShape.AddItem "Empty Box", olyBoxEmpty
cmbShape.AddItem "Spot", olySpot
cmbShape.AddItem "Empty spot", olySpotEmpty
cmbShape.AddItem "NET Stick", olyStick
cmbShape.AddItem "Triangle", olyTriangle
cmbShape.AddItem "Empty triangle", olyTriangleEmpty
cmbShape.AddItem "Tri-star", olyTriStar
cmbDefShape.AddItem "Box", olyBox
cmbDefShape.AddItem "Empty Box", olyBoxEmpty
cmbDefShape.AddItem "Spot", olySpot
cmbDefShape.AddItem "Empty spot", olySpotEmpty
cmbDefShape.AddItem "NET Stick", olyStick
cmbDefShape.AddItem "Triangle", olyTriangle
cmbDefShape.AddItem "Empty triangle", olyTriangleEmpty
cmbDefShape.AddItem "Tri-star", olyTriStar
End Sub

Private Sub FillNETAdjustmentCombo()
cmbDefNETAdjustment.AddItem "MinNET, MaxNET", olyNETFromMinMax
cmbDefNETAdjustment.AddItem "DB - TIC Fit", olyNETDB_TIC
cmbDefNETAdjustment.AddItem "DB - GANET", olyNETDB_GANET
cmbDefNETAdjustment.AddItem "Selected Display", olyNETDisplay
End Sub

Private Sub ClearEditorControls()
cmbShape.ListIndex = -1
lblClr.BackColor = OlyOptions.DefColor
lblType.Caption = ""
txtComment.Text = ""
End Sub

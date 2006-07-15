VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPEKFilter 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PEK File Data Filter"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmPEKFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraChargeState 
      BackColor       =   &H00000000&
      Caption         =   "Charge State"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   96
      Top             =   7560
      Width           =   6255
      Begin VB.TextBox txtCSIncExc 
         Height          =   615
         Left            =   3960
         MultiLine       =   -1  'True
         TabIndex        =   99
         ToolTipText     =   "Type semicolon delimited list of charge states; leave blank not to use"
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optCSIncExc 
         BackColor       =   &H00000000&
         Caption         =   "Exclude charge state(s)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   98
         Top             =   600
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optCSIncExc 
         BackColor       =   &H00000000&
         Caption         =   "Include only charge state(s)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   97
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "&Filter"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   7080
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CmnDlg1 
      Left            =   7080
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Caption         =   "St.Dev.(Charge State), Fit(Isotopic)"
      ForeColor       =   &H80000009&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   6255
      Begin VB.TextBox txtIncIsoFit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   101
         Text            =   "0.25"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtIncCSStDev 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   100
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkFit 
         BackColor       =   &H80000001&
         Caption         =   "Include Isotopic Distributions with Calculated Fit <="
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   4095
      End
      Begin VB.CheckBox chkStDev 
         BackColor       =   &H80000001&
         Caption         =   "Include Charge State Distributions with St.Dev <="
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Width           =   3855
      End
   End
   Begin VB.Frame fraScan 
      BackColor       =   &H80000001&
      Caption         =   "Text Found In Scans And Scan Number Patterns"
      ForeColor       =   &H80000009&
      Height          =   4575
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   7695
      Begin VB.CheckBox chkUsePattern 
         BackColor       =   &H80000001&
         Caption         =   "&Use pattern as filter"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   360
         TabIndex        =   95
         Top             =   4080
         Width           =   2415
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   315
         Left            =   2880
         TabIndex        =   32
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   315
         Left            =   2880
         TabIndex        =   31
         ToolTipText     =   "Apply list of patterns"
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   315
         Left            =   2880
         TabIndex        =   30
         ToolTipText     =   "Remove formula from list"
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdAddToList 
         Caption         =   "Add To List"
         Height          =   375
         Left            =   840
         TabIndex        =   29
         ToolTipText     =   "Add current formula to list of patterns"
         Top             =   2640
         Width           =   975
      End
      Begin VB.ListBox lstPatterns 
         Height          =   1035
         Left            =   3960
         TabIndex        =   28
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtFormula 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   27
         Text            =   "2*N"
         Top             =   2280
         Width           =   1440
      End
      Begin VB.TextBox txtLastScan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         TabIndex        =   26
         Text            =   "99999"
         Top             =   2685
         Width           =   600
      End
      Begin VB.TextBox txtFirstScan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         TabIndex        =   25
         Text            =   "1"
         Top             =   2205
         Width           =   600
      End
      Begin VB.CommandButton cmdExcludeAllPattern 
         Caption         =   "&Exclude All"
         Height          =   375
         Left            =   5280
         TabIndex        =   24
         Top             =   4020
         Width           =   975
      End
      Begin VB.CommandButton cmdIncludeAllPattern 
         Caption         =   "Include &All"
         Height          =   375
         Left            =   6480
         TabIndex        =   20
         ToolTipText     =   "Clear pattern"
         Top             =   4020
         Width           =   975
      End
      Begin VB.OptionButton optScansIncExc 
         BackColor       =   &H80000001&
         Caption         =   "Exclude scans not containing text"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   17
         Top             =   300
         Width           =   2775
      End
      Begin VB.OptionButton optScansIncExc 
         BackColor       =   &H80000001&
         Caption         =   "Exclude scans containing text"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.TextBox txtScanIncExc 
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Leave blank to ship this filter"
         Top             =   600
         Width           =   7215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   7560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblIncExcClr 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Included"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   94
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblIncExcClr 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Excluded"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   93
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   59
         Left            =   7080
         TabIndex        =   92
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   58
         Left            =   6720
         TabIndex        =   91
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   57
         Left            =   6360
         TabIndex        =   90
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   56
         Left            =   6000
         TabIndex        =   89
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   55
         Left            =   5640
         TabIndex        =   88
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   54
         Left            =   5280
         TabIndex        =   87
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   53
         Left            =   4920
         TabIndex        =   86
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   52
         Left            =   4560
         TabIndex        =   85
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   51
         Left            =   4200
         TabIndex        =   84
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   50
         Left            =   3840
         TabIndex        =   83
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   49
         Left            =   3480
         TabIndex        =   82
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   48
         Left            =   3120
         TabIndex        =   81
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   47
         Left            =   2760
         TabIndex        =   80
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   46
         Left            =   2400
         TabIndex        =   79
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   45
         Left            =   2040
         TabIndex        =   78
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   44
         Left            =   1680
         TabIndex        =   77
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   43
         Left            =   1320
         TabIndex        =   76
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   42
         Left            =   960
         TabIndex        =   75
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   41
         Left            =   600
         TabIndex        =   74
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   40
         Left            =   240
         TabIndex        =   73
         Top             =   3615
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   39
         Left            =   7080
         TabIndex        =   72
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   38
         Left            =   6720
         TabIndex        =   71
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   37
         Left            =   6360
         TabIndex        =   70
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   36
         Left            =   6000
         TabIndex        =   69
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   35
         Left            =   5640
         TabIndex        =   68
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   34
         Left            =   5280
         TabIndex        =   67
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   33
         Left            =   4920
         TabIndex        =   66
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   32
         Left            =   4560
         TabIndex        =   65
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   31
         Left            =   4200
         TabIndex        =   64
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   30
         Left            =   3840
         TabIndex        =   63
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   29
         Left            =   3480
         TabIndex        =   62
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   28
         Left            =   3120
         TabIndex        =   61
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   27
         Left            =   2760
         TabIndex        =   60
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   26
         Left            =   2400
         TabIndex        =   59
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   25
         Left            =   2040
         TabIndex        =   58
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   24
         Left            =   1680
         TabIndex        =   57
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   23
         Left            =   1320
         TabIndex        =   56
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   22
         Left            =   960
         TabIndex        =   55
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   21
         Left            =   600
         TabIndex        =   54
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   53
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   19
         Left            =   7080
         TabIndex        =   52
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   18
         Left            =   6720
         TabIndex        =   51
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   17
         Left            =   6360
         TabIndex        =   50
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   49
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   15
         Left            =   5640
         TabIndex        =   48
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   14
         Left            =   5280
         TabIndex        =   47
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   13
         Left            =   4920
         TabIndex        =   46
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   12
         Left            =   4560
         TabIndex        =   45
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   44
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   43
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   9
         Left            =   3480
         TabIndex        =   42
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   41
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   40
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   39
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   38
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   37
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   36
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   35
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   34
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label lblPattern 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   3135
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Scan:"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   23
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "First Scan:"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   22
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Patterns are established on actual scan numbers and not on order in which scans appear in the file."
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   7215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Scan Exclusion Pattern Formula"
         ForeColor       =   &H80000009&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   480
      Width           =   5655
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   5655
   End
   Begin VB.Frame fraDeconvolution 
      BackColor       =   &H80000001&
      Caption         =   "Deconvolution Type"
      ForeColor       =   &H80000009&
      Height          =   800
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7695
      Begin VB.OptionButton optDeconvolution 
         BackColor       =   &H80000001&
         Caption         =   "Isotopic Distributions Only"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   2
         Left            =   5280
         TabIndex        =   3
         Top             =   300
         Width           =   2175
      End
      Begin VB.OptionButton optDeconvolution 
         BackColor       =   &H80000001&
         Caption         =   "Charge State Distributions Only"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Top             =   300
         Width           =   2535
      End
      Begin VB.OptionButton optDeconvolution 
         BackColor       =   &H80000001&
         Caption         =   "Include All"
         ForeColor       =   &H80000009&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   8520
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Target:"
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPEKFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------
'filtering PEK files
'created: 09/19/2001 nt
'last modified: 11/05/2002 nt
'---------------------------------------------------
Option Explicit

Const EXCLUDE = 0
Const INCLUDE = 1

Const FLT_EXCLUDE_FOUND = 0
Const FLT_EXCLUDE_NOT_FOUND = 1
Const PATTERN_CNT = 60

Const MY_DELI = ";"
Const DATA_DELI_ASC = 9

Dim fso As New FileSystemObject

Dim TmpDir As String

Dim Filtering As Boolean
Dim TmpFileCnt As Long
Dim TmpFileNames() As String

Dim fltDataType As Long
Dim fltScanText As Long
Dim fltUsePattern As Boolean    'default False

Dim patFirstScan As Long
Dim patLastScan As Long

Dim IncExcClr(1) As Long

Dim MyPatterns As ExprEvaluator
Dim Pattern() As Integer
Dim VarVals() As Long

Dim CSCnt As Long               'count of charges states in the list
Dim CSList() As Long            'if list is empty filter is not used
Dim CSIncExc As Long

Dim IncCSStDev As Double        'filter based on isotopic fit and standard deviation
Dim IncIsoFit As Double         'of deconvoluted peaks

Private Sub chkUsePattern_Click()
fltUsePattern = (chkUsePattern.value = vbChecked)
End Sub

Private Sub cmdAddToList_Click()
'-----------------------------------------
'adds formula to the list of patterns
'-----------------------------------------
Dim Formula As String
Formula = Trim$(txtFormula.Text)
If Len(Formula) > 0 Then lstPatterns.AddItem Formula
End Sub

Private Sub cmdApply_Click()
'-----------------------------------------
'designs pattern based on lists of formula
'-----------------------------------------
Dim i As Long
Dim CurrFormula As String
On Error Resume Next
Call SetPattern(INCLUDE)
Call ShowPattern
With lstPatterns
    If .ListCount > 0 Then
        Me.MousePointer = vbHourglass
        For i = 0 To .ListCount - 1
            CurrFormula = .List(i)
            CalculatePattern CurrFormula
        Next i
        Me.MousePointer = vbDefault
    Else
        MsgBox "No patterns specified. Enter some formula F(N) and try again", vbOKOnly
    End If
End With
Call ShowPattern
End Sub

Private Sub cmdBrowse_Click()
On Error Resume Next
If Not Filtering Then
   CmnDlg1.DialogTitle = "Browse to PEK File"
   CmnDlg1.Filter = "Pek files (*.pek*)|*.pek*|All files (*.*)|*.*"
   CmnDlg1.ShowOpen
   If Err Then Exit Sub
   If Len(CmnDlg1.FileName) > 0 Then
      txtSource.Text = CmnDlg1.FileName
      txtTarget.Text = txtSource.Text & "F"
   End If
End If
End Sub

Private Sub cmdClear_Click()
lstPatterns.Clear
End Sub

Private Sub cmdClose_Click()
If Not Filtering Then Unload Me
End Sub

Private Sub cmdExcludeAllPattern_Click()
Call SetPattern(EXCLUDE)
Call ShowPattern
End Sub

Private Sub cmdFilter_Click()
Dim SourceFile As String
Dim TargetFile As String
Dim Res As Long
On Error Resume Next
SourceFile = Trim$(txtSource.Text)
If Len(SourceFile) <= 0 Then
    MsgBox "Source file not specified.", vbOKOnly
    txtSource.SetFocus
End If
TargetFile = Trim$(txtTarget.Text)
If Len(TargetFile) <= 0 Then
    MsgBox "Target file not specified.", vbOKOnly
    txtTarget.SetFocus
End If

Me.MousePointer = vbHourglass
ReDim TmpFileNames(10)
TmpFileNames(0) = SourceFile
Filtering = True
TmpFileCnt = 0
SetStatus "Filtering distribution types ..."
FilterDeconvolutionType
SetStatus "Filtering on scan patterns ..."
FilterScanTextPattern
If CSCnt > 0 Then                       'here should come ... OR any other criteria
   SetStatus "Filtering data ..."       'that would indicate need to filter on data
   FilterData
End If
SetStatus ""
If TmpFileCnt > 0 Then
    'set last file as target file and delete the rest
    fso.MoveFile TmpFileNames(TmpFileCnt), TargetFile
    If Err.Number = 58 Then         'TargetFile file exists
       Res = MsgBox("Specified file already exists. Overwrite?", vbYesNo)
       If Res = vbYes Then
          fso.DeleteFile TargetFile
          DoEvents
          fso.MoveFile TmpFileNames(TmpFileCnt), TargetFile
       Else
          fso.DeleteFile TmpFileNames(TmpFileCnt)
       End If
    End If
    DeleteTemporaryFiles
Else
    MsgBox "No filters specified. Target file not created.", vbOKOnly
End If
Filtering = False
Me.MousePointer = vbDefault
End Sub

Public Sub DeleteTemporaryFiles()
'---------------------------------------------------
'move and rename last temporary file to desired form
'and delete all other temorary files (if any)
'---------------------------------------------------
Dim i As Long
For i = 1 To TmpFileCnt - 1
    fso.DeleteFile TmpFileNames(i)
Next i
End Sub

Public Sub FilterDeconvolutionType()
'not implemented for now
'Select Case fltDataType
'Case 0
'Case 1
'    TmpFileCnt = TmpFileCnt + 1
'Case 2
'    TmpFileCnt = TmpFileCnt + 1
'End Select
End Sub

Public Sub FilterFit()
'not implemented for now
End Sub

Public Sub FilterScanTextPattern()
'--------------------------------------------------
'filters out scans depending on settings and text
'--------------------------------------------------
Dim txtToFind As String
Dim tsIn As TextStream
Dim tsOut As TextStream
Dim Done As Boolean
Dim tmp As String
Dim EndOfScan As Boolean
Dim ContainsTextToFind As Boolean
Dim ScanNumber As Long
Dim IncludeScan As Boolean
Dim sL As String
On Error GoTo err_FilterScanText
txtToFind = Trim$(txtScanIncExc.Text)
If Len(txtToFind) > 0 Then
   With fso
     TmpFileCnt = TmpFileCnt + 1
     TmpFileNames(TmpFileCnt) = TmpDir & .GetTempName
     Debug.Print TmpFileNames(TmpFileCnt)
     Set tsIn = .OpenTextFile(TmpFileNames(TmpFileCnt - 1), ForReading)
     Set tsOut = .CreateTextFile(TmpFileNames(TmpFileCnt))
     Do Until Done
        Do Until EndOfScan
           If tsIn.AtEndOfStream Then
              EndOfScan = True
              Done = True
           Else
              sL = Trim$(tsIn.ReadLine)
              'always pick scan numbers since pattern might be in use
              If Left$(sL, Len(PEK_D_FILENAME)) = PEK_D_FILENAME Then
                 ScanNumber = GetScanNumber(sL)
              End If
              If Len(sL) > 0 Then
                 tmp = tmp & sL & vbCrLf
              Else
                 EndOfScan = True
              End If
           End If
        Loop
        If Len(tmp) > 0 Then
           IncludeScan = False
           ContainsTextToFind = (InStr(1, tmp, txtToFind) > 0)
           Select Case fltScanText
           Case FLT_EXCLUDE_FOUND
                If Not ContainsTextToFind Then IncludeScan = True
           Case FLT_EXCLUDE_NOT_FOUND
                If ContainsTextToFind Then IncludeScan = True
           End Select
           If fltUsePattern Then
              'everything out of FirstScan,LastScan stays the same(included or excluded)
              If ScanNumber >= patFirstScan And ScanNumber <= patLastScan Then
                 'do not include if text filter already excluded this scan
                 IncludeScan = IncludeScan And (Pattern(ScanNumber) = INCLUDE)
              End If
           End If
           If IncludeScan Then
              tsOut.Write tmp
              tsOut.WriteBlankLines 1
           End If
        End If
        If Not Done Then
           tmp = ""
           EndOfScan = False
        End If
     Loop
     tsIn.Close
     tsOut.Close
   End With
End If
Exit Sub

err_FilterScanText:
End Sub

Private Sub cmdIncludeAllPattern_Click()
Call SetPattern(INCLUDE)
Call ShowPattern
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
With lstPatterns
    .RemoveItem .ItemData(.ListIndex)
End With
End Sub

Private Sub Form_Load()
'------------------------------------------------------
'set defaults settings and prepare various functions
'------------------------------------------------------
Dim i As Long
On Error Resume Next
Me.Move 200, 200
fltDataType = 0
fltScanText = 0
patFirstScan = 1
patLastScan = 99999
CSIncExc = 1            'exclusion
For i = 0 To 1
    IncExcClr(i) = lblIncExcClr(i).BackColor
Next i
Call SetPattern(INCLUDE)
ShowPattern
With fso
    TmpDir = .GetAbsolutePathName(.GetSpecialFolder(TemporaryFolder)) & "\"
    Debug.Print TmpDir
End With
For i = 0 To PATTERN_CNT - 1
    lblPattern(i).Caption = i + 1
Next i
'initialize pattern calculator
Set MyPatterns = New ExprEvaluator
MyPatterns.Vars.add 1, "N"
ReDim VarVals(1 To 1)

'pick other default values
IncCSStDev = CDbl(txtIncCSStDev.Text)
IncIsoFit = CDbl(txtIncIsoFit.Text)
End Sub

Private Sub lblPattern_Click(Index As Integer)
If Pattern(Index + 1) = 1 Then
   Pattern(Index + 1) = 0
Else
   Pattern(Index + 1) = 1
End If
lblPattern(Index).BackColor = IncExcClr(Pattern(Index + 1))
End Sub

Private Sub optCSIncExc_Click(Index As Integer)
CSIncExc = Index
End Sub

Private Sub optDeconvolution_Click(Index As Integer)
fltDataType = Index
End Sub

Private Sub optScansIncExc_Click(Index As Integer)
fltScanText = Index
End Sub

Private Sub SetStatus(ByVal Status As String)
lblStatus.Caption = Status
DoEvents
End Sub

Private Sub txtCSIncExc_LostFocus()
'-----------------------------------------------
'resolve list of charge states to use
'-----------------------------------------------
Dim sCS As String
Dim CSSplit() As String
Dim i As Long
On Error Resume Next
CSCnt = 0
sCS = Trim$(txtCSIncExc.Text)
If Len(sCS) > 0 Then
   CSSplit = Split(sCS, MY_DELI)
   ReDim CSList(UBound(CSSplit))
   For i = 0 To UBound(CSSplit)
       If IsNumeric(CSSplit(i)) Then
          CSCnt = CSCnt + 1
          CSList(CSCnt - 1) = CLng(CSSplit(i))
       End If
   Next i
End If
If CSCnt > 0 Then
   ReDim Preserve CSList(CSCnt - 1)
Else
   Erase CSList
End If
End Sub

Private Sub txtFirstScan_LostFocus()
On Error Resume Next
patFirstScan = CLng(txtFirstScan.Text)
If Err Then
   MsgBox "This argument has to be positive integer.", vbOKOnly
   txtFirstScan.SetFocus
End If
End Sub

Private Sub txtIncCSStDev_LostFocus()
If IsNumeric(txtIncCSStDev.Text) Then
   IncCSStDev = CDbl(txtIncCSStDev.Text)
Else
   MsgBox "This value should be numeric.", vbOKOnly, glFGTU
   txtIncCSStDev.SetFocus
End If
End Sub

Private Sub txtIncIsoFit_LostFocus()
If IsNumeric(txtIncIsoFit.Text) Then
   IncIsoFit = CDbl(txtIncIsoFit.Text)
Else
   MsgBox "This value should be numeric.", vbOKOnly, glFGTU
   txtIncIsoFit.SetFocus
End If
End Sub

Private Sub txtLastScan_LostFocus()
On Error Resume Next
patLastScan = CLng(txtLastScan.Text)
If Err Then
   MsgBox "This argument has to be positive integer.", vbOKOnly
   txtLastScan.SetFocus
End If
End Sub

Private Sub SetPattern(ByVal IncExc As Long)
'------------------------------------------
'sets or clears pattern array
'------------------------------------------
Dim i As Long
ReDim Pattern(99999)
If IncExc = 1 Then
   For i = 1 To 99999
      Pattern(i) = 1
   Next i
End If
End Sub

Private Sub ShowPattern()
'------------------------------------------
'displays pattern on check boxes
'------------------------------------------
Dim i As Long
For i = 1 To PATTERN_CNT
    lblPattern(i - 1).BackColor = IncExcClr(Pattern(i))
Next i
End Sub

Public Sub CalculatePattern(ByVal Formula As String)
'---------------------------------------------------------
'calculates pattern based on formula; ignores bad formulas
'old pattern is preserved so that patterns can be merged
'---------------------------------------------------------
Dim i As Long
Dim Ind As Long
Dim Done As Boolean
With MyPatterns
    .Expr = Formula
    If .IsExprValid Then
       i = 1
       Do Until Done
          VarVals(1) = i
          Ind = CLng(.ExprVal(VarVals()))
          If Ind <= 99999 Then
             Pattern(Ind) = EXCLUDE
             i = i + 1
          Else
             Done = True
          End If
       Loop
    Else
       Debug.Print "Invalid formula: " & Formula
    End If
End With
End Sub


Private Function GetTokenValue(ByVal sLine As String, _
                              ByVal sToken As String) As String
'--------------------------------------------------------------
'returns part of the sLine after sToken
'--------------------------------------------------------------
Dim TokenStart As Long
Dim TokenLen As Long
TokenLen = Len(sToken)
If TokenLen > 0 Then
   TokenStart = InStr(1, sLine, sToken)
   If TokenStart > 0 Then
      GetTokenValue = Right$(sLine, Len(sLine) - (TokenStart + TokenLen) + 1)
   End If
End If
End Function


Private Function GetScanNumber(ByVal sText As String) As Long
'-----------------------------------------------------------
'returns scan number from sText; this is hard coded for
'different file types
'FTICR scans file name ends with 5 character scan number
'used as file name extension
'QToF scans file name ends with anything:ScanNumber:number
'-----------------------------------------------------------
Dim tmp As String
'try first with FTICR scan names
tmp = Right$(sText, 5)
If IsNumeric(tmp) Then
   GetScanNumber = CLng(tmp)
   Exit Function
End If
End Function


Private Function IsInCSList(ByVal CS As Long) As Boolean
'-------------------------------------------------------
'returns True if charge state CS is in the CSList; False
'otherwise(or on any error)
'-------------------------------------------------------
Dim i As Long
On Error GoTo exit_IsInCSList
For i = 0 To CSCnt - 1
    If CSList(i) = CS Then
       IsInCSList = True
       GoTo exit_IsInCSList
    End If
Next i
exit_IsInCSList:
End Function


Private Sub FilterData()
'--------------------------------------------------
'filters out data lines based on various settings
'--------------------------------------------------
Dim tsIn As TextStream
Dim tsOut As TextStream
Dim Done As Boolean
Dim sL As String
Dim sLSplit() As String
Dim CurrCS As Long
Dim CurrDbl As Double
Dim IncLine As Boolean
Dim DataType As Long            'charge state or isotopic
On Error Resume Next

With fso
   TmpFileCnt = TmpFileCnt + 1
   TmpFileNames(TmpFileCnt) = TmpDir & .GetTempName
   Set tsIn = .OpenTextFile(TmpFileNames(TmpFileCnt - 1), ForReading)
   Set tsOut = .CreateTextFile(TmpFileNames(TmpFileCnt))
   Do Until Done
      IncLine = True
      sL = Trim$(tsIn.ReadLine)
      Select Case Left$(sL, 8)
      Case t8DATA_CS
           DataType = glCSType
      Case t8DATA_ISO
           DataType = glIsoType
      End Select
      sLSplit = Split(sL, Chr(DATA_DELI_ASC))
      If UBound(sLSplit) > 5 Then             'good enough for data line
         Select Case DataType
         Case glCSType
            If CSCnt > 0 Then                       'apply charge state filter
               CurrCS = -1
               If IsNumeric(sLSplit(0)) Then
                  CurrCS = CLng(sLSplit(0))
               Else                                 'data line could start with asterisk
                  CurrCS = CLng(sLSplit(1))
               End If
               If CurrCS >= 0 Then
                  Select Case CSIncExc
                  Case 0                            'exclude if not in list
                       If Not IsInCSList(CurrCS) Then IncLine = False
                  Case 1                            'exclude if in list
                       If IsInCSList(CurrCS) Then IncLine = False
                  End Select
               End If
            End If
            If chkStDev.value = vbChecked Then      'apply standard deviation filter
               If IsNumeric(sLSplit(0)) Then
                  CurrDbl = CDbl(sLSplit(4))
               Else                                 'data line could start with asterisk
                  CurrDbl = CDbl(sLSplit(5))
               End If
               If CurrDbl > IncCSStDev Then IncLine = False
            End If
         Case glIsoType
            If CSCnt > 0 Then
               CurrCS = -1
               If IsNumeric(sLSplit(0)) Then
                  CurrCS = CLng(sLSplit(0))
               Else                                 'data line could start with asterisk
                  CurrCS = CLng(sLSplit(1))
               End If
               If CurrCS >= 0 Then
                  Select Case CSIncExc
                  Case 0                            'exclude if not in list
                       If Not IsInCSList(CurrCS) Then IncLine = False
                  Case 1                            'exclude if in list
                       If IsInCSList(CurrCS) Then IncLine = False
                  End Select
               End If
            End If
            If chkFit.value = vbChecked Then        'apply isotopic fit filter
               If IsNumeric(sLSplit(0)) Then
                  CurrDbl = CDbl(sLSplit(3))
               Else                                 'data line could start with asterisk
                  CurrDbl = CDbl(sLSplit(4))
               End If
               If CurrDbl > IncIsoFit Then IncLine = False
            End If
         End Select
      End If
      If IncLine Then tsOut.WriteLine sL
      If tsIn.AtEndOfStream Then Done = True
   Loop
   tsIn.Close
   tsOut.Close
End With
exit_FilterData:
End Sub

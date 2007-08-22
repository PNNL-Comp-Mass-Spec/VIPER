VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   6315
   ClientLeft      =   2580
   ClientTop       =   1515
   ClientWidth     =   6600
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin TabDlg.SSTab tbsOptions 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Data"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picOpt(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Graph"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picOpt(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Miscellaneous"
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picOpt(2)"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picOpt 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   4695
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   4695
         ScaleWidth      =   6135
         TabIndex        =   86
         Top             =   480
         Width           =   6135
         Begin VB.Frame fraICR2LS 
            Caption         =   "ICR-2LS"
            Height          =   2055
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   5895
            Begin VB.Frame fraIsoICR2LSCall 
               Caption         =   "Determine range in ICR-2LS calls for Isotopic data based on"
               Height          =   735
               Left            =   240
               TabIndex        =   91
               Top             =   1080
               Width           =   5295
               Begin VB.OptionButton optIsoICR2LS 
                  Caption         =   "the &most abundant mass"
                  Height          =   255
                  Index           =   1
                  Left            =   2520
                  TabIndex        =   93
                  Top             =   280
                  Width           =   2535
               End
               Begin VB.OptionButton optIsoICR2LS 
                  Caption         =   "m/&z value"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   92
                  Top             =   280
                  Value           =   -1  'True
                  Width           =   1455
               End
            End
            Begin VB.CommandButton cmdBrowseForICR2LS 
               Caption         =   "&Browse"
               Height          =   375
               Left            =   4680
               TabIndex        =   90
               Top             =   540
               Width           =   855
            End
            Begin VB.TextBox txtICR2LSExePath 
               Height          =   285
               Left            =   240
               TabIndex        =   89
               Top             =   600
               Width           =   4215
            End
            Begin VB.Label lblICR2LSExePath 
               Caption         =   "ICR-2LS.Exe Path"
               Height          =   255
               Left            =   240
               TabIndex        =   88
               Top             =   255
               Width           =   4575
            End
         End
      End
      Begin VB.PictureBox picOpt 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   5055
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   5055
         ScaleWidth      =   6105
         TabIndex        =   33
         Top             =   360
         Width           =   6100
         Begin VB.CheckBox chkBorderColor 
            Caption         =   "&Marker border color same as interior"
            Height          =   420
            Left            =   2040
            TabIndex        =   82
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox txtDDRatioMax 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5400
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   4680
            Width           =   585
         End
         Begin VB.TextBox txtMinPointSize 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3000
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   3960
            Width           =   585
         End
         Begin VB.Frame fraHOrientation 
            Caption         =   "Hor. Orientation"
            Height          =   960
            Left            =   1680
            TabIndex        =   62
            Top             =   2880
            Width           =   1455
            Begin VB.OptionButton optHorOrientation 
               Caption         =   "pI-positive"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   63
               ToolTipText     =   "pI increases (FN decreases) from left to right"
               Top             =   280
               Width           =   1095
            End
            Begin VB.OptionButton optHorOrientation 
               Caption         =   "pI-negative"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   64
               ToolTipText     =   "pI increases (FN decreases) from right to left"
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.Frame fraVOrientation 
            Caption         =   "Ver. Orientation"
            Height          =   960
            Left            =   120
            TabIndex        =   59
            Top             =   2880
            Width           =   1455
            Begin VB.OptionButton optVerOrientation 
               Caption         =   "&Positive"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   60
               ToolTipText     =   "Molecular mass increases from bottom to top"
               Top             =   280
               Width           =   1095
            End
            Begin VB.OptionButton optVerOrientation 
               Caption         =   "&Negative"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   61
               ToolTipText     =   "Molecular mass increases from top to bottom"
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.Frame fraType 
            Caption         =   "Chart Type"
            Height          =   600
            Left            =   120
            TabIndex        =   52
            Top             =   2160
            Width           =   3000
            Begin VB.OptionButton optType 
               Caption         =   "NET-type"
               Height          =   195
               Index           =   2
               Left            =   1950
               TabIndex        =   55
               ToolTipText     =   "Horizontal axis represents NET"
               Top             =   280
               Width           =   975
            End
            Begin VB.OptionButton optType 
               Caption         =   "&pI-type"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   53
               ToolTipText     =   "Horizontal axis represents pI numbers"
               Top             =   280
               Width           =   855
            End
            Begin VB.OptionButton optType 
               Caption         =   "&FN-type"
               Height          =   195
               Index           =   1
               Left            =   1000
               TabIndex        =   54
               ToolTipText     =   "Horizontal axis represents scan numbers"
               Top             =   280
               Width           =   975
            End
         End
         Begin VB.TextBox txtMaxPointFactor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5400
            MultiLine       =   -1  'True
            TabIndex        =   74
            Top             =   3960
            Width           =   585
         End
         Begin VB.TextBox txtAbuAspectRatio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            MultiLine       =   -1  'True
            TabIndex        =   78
            Top             =   4680
            Width           =   585
         End
         Begin VB.Frame fraVAxis 
            Caption         =   "Vertical Axis"
            Height          =   600
            Left            =   3240
            TabIndex        =   56
            Top             =   2160
            Width           =   2775
            Begin VB.OptionButton optVAxis 
               Caption         =   "&Linear"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   57
               Top             =   280
               Width           =   975
            End
            Begin VB.OptionButton optVAxis 
               Caption         =   "Lo&garithmic"
               Height          =   195
               Index           =   1
               Left            =   1400
               TabIndex        =   58
               Top             =   280
               Width           =   1110
            End
         End
         Begin VB.Frame fraCooSysCross 
            Caption         =   "Coordinate Axes Crossing"
            Height          =   975
            Left            =   3240
            TabIndex        =   65
            Top             =   2880
            Width           =   2775
            Begin VB.OptionButton optCooSysCross 
               Caption         =   "&Bottom Left"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   66
               Top             =   280
               Width           =   1215
            End
            Begin VB.OptionButton optCooSysCross 
               Caption         =   "Bottom &Right"
               Height          =   255
               Index           =   1
               Left            =   1400
               TabIndex        =   68
               Top             =   280
               Width           =   1320
            End
            Begin VB.OptionButton optCooSysCross 
               Caption         =   "&Top Left"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   67
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optCooSysCross 
               Caption         =   "Top Rig&ht"
               Height          =   255
               Index           =   3
               Left            =   1400
               TabIndex        =   69
               Top             =   600
               Width           =   1200
            End
         End
         Begin VB.Frame fraSelection 
            Caption         =   "Selection"
            Height          =   1095
            Left            =   3960
            TabIndex        =   44
            Top             =   0
            Width           =   2055
            Begin VB.OptionButton optSelection 
               Caption         =   "&Color"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   45
               Top             =   300
               Width           =   735
            End
            Begin VB.OptionButton optSelection 
               Caption         =   "&Flag"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   46
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "&Selection Color"
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   47
               Top             =   660
               Width           =   1065
            End
            Begin VB.Label lblClr 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   4
               Left            =   1440
               TabIndex        =   48
               ToolTipText     =   "Double-click to change"
               Top             =   640
               Width           =   315
            End
         End
         Begin VB.CheckBox chkAutoSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto-size"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            ToolTipText     =   "If Auto Size is checked Min and Max size factors are adjusted automatically for optimal viewing."
            Top             =   4040
            Width           =   1095
         End
         Begin VB.TextBox txtAutoSizeMultiplier 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3000
            MultiLine       =   -1  'True
            TabIndex        =   76
            Top             =   4320
            Width           =   585
         End
         Begin VB.Label label1 
            Caption         =   "M&ax Ind./ Supp. for ER over"
            Height          =   195
            Index           =   3
            Left            =   3240
            TabIndex        =   79
            ToolTipText     =   "ERs over this value will be displayed as Overexpressed"
            Top             =   4740
            Width           =   2055
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Ma&x pt. size factor"
            Height          =   195
            Index           =   0
            Left            =   3840
            TabIndex        =   73
            ToolTipText     =   "Controls the maximum overlap(horizontal) of the neighboring spots."
            Top             =   4020
            Width           =   1290
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Mi&n pt. size factor"
            Height          =   195
            Index           =   5
            Left            =   1560
            TabIndex        =   71
            ToolTipText     =   "Controls the minimum size of the data spot."
            Top             =   4020
            Width           =   1245
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Data point size aspect ratio"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   77
            ToolTipText     =   "Horizontal/vertical aspect ratio for data point size."
            Top             =   4740
            Width           =   1920
         End
         Begin VB.Label lblCSShape 
            Caption         =   "Charge State Marker Shape"
            Height          =   375
            Left            =   2040
            TabIndex        =   42
            ToolTipText     =   "Double-click to change"
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblIsoShape 
            Caption         =   "Isotopic Marker Shape"
            Height          =   375
            Left            =   2040
            TabIndex        =   43
            ToolTipText     =   "Double-click to change"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblUnderColor 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suppressed"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Double-click to change"
            Top             =   1550
            Width           =   840
         End
         Begin VB.Label lblOverColor 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suppressed"
            Height          =   195
            Left            =   5160
            TabIndex        =   51
            ToolTipText     =   "Double-click to change"
            Top             =   1550
            Width           =   840
         End
         Begin VB.Label lblNormalColor 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Normal"
            Height          =   195
            Left            =   2805
            TabIndex        =   50
            ToolTipText     =   "Double-click to change"
            Top             =   1550
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Background Color"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Foreground Color"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Charge State Color"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   825
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Isotopic Color"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Top             =   1185
            Width           =   960
         End
         Begin VB.Label lblClr 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   35
            ToolTipText     =   "Left-click to change"
            Top             =   120
            Width           =   315
         End
         Begin VB.Label lblClr 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   37
            ToolTipText     =   "Left-click to change"
            Top             =   480
            Width           =   315
         End
         Begin VB.Label lblClr 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   39
            ToolTipText     =   "Left-click to change"
            Top             =   825
            Width           =   315
         End
         Begin VB.Label lblClr 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   41
            ToolTipText     =   "Left-click to change"
            Top             =   1155
            Width           =   315
         End
         Begin VB.Image imCSMarker 
            Height          =   240
            Index           =   0
            Left            =   3480
            Picture         =   "frmOptions.frx":0060
            Stretch         =   -1  'True
            ToolTipText     =   "Left-click to change"
            Top             =   165
            Width           =   345
         End
         Begin VB.Image imCSMarker 
            Height          =   240
            Index           =   1
            Left            =   3480
            Picture         =   "frmOptions.frx":036A
            Stretch         =   -1  'True
            Top             =   165
            Width           =   345
         End
         Begin VB.Image imCSMarker 
            Height          =   240
            Index           =   2
            Left            =   3480
            Picture         =   "frmOptions.frx":0674
            Stretch         =   -1  'True
            Top             =   165
            Width           =   345
         End
         Begin VB.Image imCSMarker 
            Height          =   240
            Index           =   3
            Left            =   3480
            Picture         =   "frmOptions.frx":097E
            Stretch         =   -1  'True
            Top             =   165
            Width           =   345
         End
         Begin VB.Image imCSMarker 
            Height          =   240
            Index           =   4
            Left            =   3480
            Picture         =   "frmOptions.frx":0C88
            Stretch         =   -1  'True
            Top             =   165
            Width           =   345
         End
         Begin VB.Image imCSMarker 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   5
            Left            =   3480
            Picture         =   "frmOptions.frx":0F92
            Stretch         =   -1  'True
            Top             =   165
            Width           =   345
         End
         Begin VB.Image imIsoMarker 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   3480
            Picture         =   "frmOptions.frx":129C
            Stretch         =   -1  'True
            ToolTipText     =   "Left-click to change"
            Top             =   640
            Width           =   345
         End
         Begin VB.Image imIsoMarker 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   3480
            Picture         =   "frmOptions.frx":15A6
            Stretch         =   -1  'True
            Top             =   640
            Width           =   345
         End
         Begin VB.Image imIsoMarker 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   3480
            Picture         =   "frmOptions.frx":18B0
            Stretch         =   -1  'True
            Top             =   640
            Width           =   345
         End
         Begin VB.Image imIsoMarker 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   3
            Left            =   3480
            Picture         =   "frmOptions.frx":1BBA
            Stretch         =   -1  'True
            Top             =   640
            Width           =   345
         End
         Begin VB.Image imIsoMarker 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   4
            Left            =   3480
            Picture         =   "frmOptions.frx":1EC4
            Stretch         =   -1  'True
            Top             =   640
            Width           =   345
         End
         Begin VB.Image imIsoMarker 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   5
            Left            =   3480
            Picture         =   "frmOptions.frx":21CE
            Stretch         =   -1  'True
            Top             =   640
            Width           =   345
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Auto Size Multiplier"
            Height          =   195
            Index           =   7
            Left            =   1560
            TabIndex        =   75
            ToolTipText     =   "Controls the minimum size of the data spot."
            Top             =   4380
            Width           =   1335
         End
      End
      Begin VB.PictureBox picOpt 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   4695
         Index           =   0
         Left            =   120
         ScaleHeight     =   4695
         ScaleWidth      =   6135
         TabIndex        =   1
         Top             =   360
         Width           =   6135
         Begin VB.Frame fraMTDisplay 
            Caption         =   "MT tags Display"
            Height          =   1455
            Left            =   120
            TabIndex        =   10
            Top             =   3120
            Width           =   2535
            Begin VB.TextBox txtMTScanMax 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1860
               TabIndex        =   15
               Text            =   "1000"
               Top             =   675
               Width           =   600
            End
            Begin VB.TextBox txtMTScanMin 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1160
               TabIndex        =   14
               Text            =   "1"
               Top             =   675
               Width           =   600
            End
            Begin VB.OptionButton optMTDisplayType 
               Caption         =   "Isotopic"
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   12
               Top             =   280
               Width           =   855
            End
            Begin VB.OptionButton optMTDisplayType 
               Caption         =   "CS"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   11
               Top             =   280
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Scan Range:"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame fraGelToPEK 
            Caption         =   "Gel To PEK Transfer"
            Height          =   855
            Left            =   2760
            TabIndex        =   31
            Top             =   3720
            Width           =   3255
            Begin VB.CheckBox chkWriteFreqShift 
               Caption         =   "Write frequency shifts to PEK file"
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   360
               Width           =   2655
            End
         End
         Begin VB.Frame fraDDRatioOrientation 
            Caption         =   "Expression Ratio (ER) Definition"
            Height          =   1695
            Left            =   2760
            TabIndex        =   23
            Top             =   1920
            Width           =   3255
            Begin VB.Frame fraERDef 
               Caption         =   "Definition"
               Height          =   560
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   3015
               Begin VB.OptionButton optDDROrientation 
                  Caption         =   "&Ind. - Suppr."
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   26
                  ToolTipText     =   "Induction - Suppression"
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.OptionButton optDDROrientation 
                  Caption         =   "&Suppr. - Ind."
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   25
                  ToolTipText     =   "Suppression - Induction"
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.Frame fraERCalc 
               Caption         =   "Calculation"
               Height          =   735
               Left            =   120
               TabIndex        =   27
               Top             =   840
               Width           =   3015
               Begin VB.OptionButton optERCalc 
                  Caption         =   "Symmetric Ratio"
                  Height          =   195
                  Index           =   2
                  Left            =   720
                  TabIndex        =   30
                  ToolTipText     =   "Zero Symetric Ratio (Abundance of Light member/ Abundance of Heavy Member)"
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.OptionButton optERCalc 
                  Caption         =   "Log. Ratio"
                  Height          =   195
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   29
                  ToolTipText     =   "Logarithmic Ratio (Abundance of Light member/ Abundance of Heavy Member)"
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.OptionButton optERCalc 
                  Caption         =   "Ratio (L/H)"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   28
                  ToolTipText     =   "Ratio (Abundance of Light member/ Abundance of Heavy Member)"
                  Top             =   240
                  Width           =   1215
               End
            End
         End
         Begin VB.Frame fraTolerances 
            Caption         =   "Tolerances (defaults)"
            Height          =   1575
            Left            =   2760
            TabIndex        =   16
            Top             =   240
            Width           =   3255
            Begin VB.TextBox txtIsoDataFit 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               MultiLine       =   -1  'True
               TabIndex        =   22
               Text            =   "frmOptions.frx":24D8
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtDupTolerance 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               MultiLine       =   -1  'True
               TabIndex        =   20
               Text            =   "frmOptions.frx":24DE
               Top             =   690
               Width           =   615
            End
            Begin VB.TextBox txtDBTolerance 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               MultiLine       =   -1  'True
               TabIndex        =   18
               Top             =   300
               Width           =   615
            End
            Begin VB.Label label1 
               Caption         =   "Use Data with Fit Better Than"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   21
               Top             =   1100
               Width           =   2175
            End
            Begin VB.Label label1 
               Caption         =   "Duplicate Elimination Tolerance"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   740
               Width           =   2415
            End
            Begin VB.Label label1 
               Caption         =   "Database Matching Tolerance"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   2175
            End
         End
         Begin VB.Frame frCloseGuesses 
            Caption         =   "Case of  Two Close Results"
            Height          =   1280
            Left            =   120
            TabIndex        =   6
            Top             =   1680
            Width           =   2535
            Begin VB.OptionButton optTwoCloseResults 
               Caption         =   "Use &more likely"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   9
               ToolTipText     =   "Use more likely according to other results"
               Top             =   900
               Width           =   1935
            End
            Begin VB.OptionButton optTwoCloseResults 
               Caption         =   "Use better fit (&calculation)"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   8
               ToolTipText     =   "Include better fit from calculation"
               Top             =   600
               Value           =   -1  'True
               Width           =   2175
            End
            Begin VB.OptionButton optTwoCloseResults 
               Caption         =   "Use &both"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   7
               ToolTipText     =   "Include both data points"
               Top             =   300
               Width           =   1935
            End
         End
         Begin VB.Frame frIsoDataFrom 
            Caption         =   "Use Isotopic MW From"
            Height          =   1280
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2535
            Begin VB.OptionButton optIsoDataFrom 
               Caption         =   "Most &abundant"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   5
               Top             =   900
               Width           =   1815
            End
            Begin VB.OptionButton optIsoDataFrom 
               Caption         =   "&Average field"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   4
               Top             =   600
               Width           =   1815
            End
            Begin VB.OptionButton optIsoDataFrom 
               Caption         =   "&Monoisotopic field"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   3
               Top             =   300
               Value           =   -1  'True
               Width           =   1815
            End
         End
      End
   End
   Begin VB.CheckBox chkActiveOnly 
      Caption         =   "&Apply to active gel only"
      Height          =   375
      Left            =   1560
      TabIndex        =   83
      ToolTipText     =   "If checked the changes apply only to the active gel (color settings are always used in all gels.)"
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   360
      TabIndex        =   81
      ToolTipText     =   "Reset options to default values"
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   85
      ToolTipText     =   "Reject options changes and close"
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   84
      ToolTipText     =   "Save options settings and close"
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Last modified 11/08/2001 nt
'----------------------------------------------------------
Option Explicit

Dim WorkPref As GelPrefs
Dim ActiveInd As Long
Dim OldDRDefinition As Integer   'no need to restore this Old, just to see if changed
Dim OldICR2LSCommandLine As String
Dim OldCSColor As Long
Dim OldIsoColor As Long
Dim OldBackColor As Long
Dim OldUnderColor As Long
Dim OldOverColor As Long
Dim OldMidColor As Long
Dim OldForeColor As Long
Dim OldCSShape As Integer
Dim OldIsoShape As Integer
Dim OldDDRatioMax As Double
Dim OldSelColor As Long
Dim OldFTICR_AMTPath As String
Dim OldWriteFreqShift As Boolean
Dim OldAutoAdjSize As Boolean
Dim OldAutoSizeMultiplier As Single
Dim OldERCalcType As Long
Dim OldMTType As Long
Dim OldMTScanMin As Long
Dim OldMTScanMax As Long

Private Sub chkAutoSize_Click()
glbPreferencesExpanded.AutoAdjSize = (chkAutoSize.Value = vbChecked)
End Sub

Private Sub chkBorderColor_Click()
If chkBorderColor.Value = vbChecked Then
    WorkPref.BorderClrSameAsInt = True
Else
    WorkPref.BorderClrSameAsInt = False
End If
End Sub

Private Sub chkWriteFreqShift_Click()
glWriteFreqShift = (chkWriteFreqShift.Value = vbChecked)
End Sub

Private Sub cmdBrowseForICR2LS_Click()
Dim PathToICR2LS As String

On Error Resume Next
PathToICR2LS = SelectFile(Me.hwnd, "Find ICR-2LS.EXE", txtICR2LSExePath.Text, False, "ICR-2LS.exe", "Executable files (*.exe)|*.exe|All Files (*.*)|*.*", 1)

If Len(PathToICR2LS) > 0 Then txtICR2LSExePath.Text = PathToICR2LS
End Sub

'No longer supported on this form (March 2006)
''Private Sub cmdBrowseForLegacyDB_Click()
''    Dim strNewFilePath As String
''
''    strNewFilePath = SelectLegacyMTDB(Me, txtLegacyDBPath.Text)
''    If Len(strNewFilePath) > 0 Then
''        txtLegacyDBPath.Text = strNewFilePath
''    End If
''End Sub

' No longer supported (March 2006)
''Private Sub cmdBrowseFTICR_AMT_Click()
''Dim PathToFTICR_AMTDatabase As String
''On Error Resume Next
''PathToFTICR_AMTDatabase = SelectFile(Me.hwnd, "Select FTICR_AMT database", , False, txtFTICR_AMTPath.Text, "Access DB files (*.mdb)|*.mdb|All Files (*.*)|*.*", 1)
''If Len(PathToFTICR_AMTDatabase) > 0 Then txtFTICR_AMTPath.Text = PathToFTICR_AMTDatabase
''End Sub

Private Sub cmdCancel_Click()
RestoreOldSettings
DDInitColors
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo err_OK
If ActiveInd < 0 Then
   glPreferences = WorkPref
   If ColorsShapesChanged Then vWhatever = 4  'indication to MDI that something changed
Else
   GelData(ActiveInd).Preferences = WorkPref
   If chkActiveOnly.Value = vbChecked Then
      If ColorsShapesChanged Then
         vWhatever = 2
      Else
         vWhatever = 1
      End If
   Else     'set new global preferences if changed for all gels
      glPreferences = WorkPref
      vWhatever = 3
   End If
End If
If OldDRDefinition <> WorkPref.DRDefinition Then vWhatever = -vWhatever
If OldDDRatioMax <> glDDRatioMax Then vWhatever = -vWhatever

Unload Me

err_OK:
Exit Sub
End Sub

Private Sub cmdReset_Click()
On Error GoTo err_Reset
ResetOptions WorkPref
DDInitColors
Settings

err_Reset:
Exit Sub
End Sub

Private Sub Form_Activate()
vWhatever = 0
If Me.Tag = "MDI" Then
   chkActiveOnly.Value = vbUnchecked
   chkActiveOnly.Enabled = False
   optType(0).Enabled = False
   ActiveInd = -1
Else
   chkActiveOnly.Value = vbChecked
   ActiveInd = val(Me.Tag)
   If Not GelData(ActiveInd).pICooSysEnabled Then optType(0).Enabled = False
End If
SaveOldSettings
If ActiveInd < 0 Then
   WorkPref = glPreferences
Else
   WorkPref = GelData(ActiveInd).Preferences
End If
Settings
End Sub

Private Sub Form_Load()
    tbsOptions.Tab = 0
    PositionControls
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub imCSMarker_Click(Index As Integer)
glCSShape = (glCSShape + 1) Mod 6
CSMarkerChange
End Sub

Private Sub imIsoMarker_Click(Index As Integer)
glIsoShape = (glIsoShape + 1) Mod 6
IsoMarkerChange
End Sub

Private Sub lblClr_Click(Index As Integer)
lblClr_DblClick Index
End Sub

Private Sub lblClr_DblClick(Index As Integer)
On Error GoTo exit_lblClr
Select Case Index
Case 0
     Call GetColorAPIDlg(Me.hwnd, glBackColor)
     lblClr(0).BackColor = glBackColor
Case 1
     Call GetColorAPIDlg(Me.hwnd, glForeColor)
     lblClr(1).BackColor = glForeColor
Case 2
     Call GetColorAPIDlg(Me.hwnd, glCSColor)
     lblClr(2).BackColor = glCSColor
Case 3
     Call GetColorAPIDlg(Me.hwnd, glIsoColor)
     lblClr(3).BackColor = glIsoColor
Case 4
     Call GetColorAPIDlg(Me.hwnd, glSelColor)
     lblClr(4).BackColor = glSelColor
End Select

exit_lblClr:
End Sub

Private Sub lblCSShape_Click()
lblCSShape_DblClick
End Sub

Private Sub lblCSShape_DblClick()
glCSShape = (glCSShape + 1) Mod 6
CSMarkerChange
End Sub

Private Sub lblIsoShape_Click()
lblIsoShape_DblClick
End Sub

Private Sub lblIsoShape_DblClick()
glIsoShape = (glIsoShape + 1) Mod 6
IsoMarkerChange
End Sub

Private Sub lblNormalColor_Click()
lblNormalColor_DblClick
End Sub

Private Sub lblNormalColor_DblClick()
On Error Resume Next
Call GetColorAPIDlg(Me.hwnd, glMidColor)
Call DDInitColors
Iris
End Sub

Private Sub lblOverColor_Click()
lblOverColor_DblClick
End Sub

Private Sub lblOverColor_DblClick()
On Error Resume Next
Call GetColorAPIDlg(Me.hwnd, glOverColor)
Call DDInitColors
Iris
End Sub

Private Sub lblUnderColor_Click()
lblUnderColor_DblClick
End Sub

Private Sub lblUnderColor_DblClick()
On Error Resume Next
Call GetColorAPIDlg(Me.hwnd, glUnderColor)
Call DDInitColors
Iris
End Sub

Private Sub optCooSysCross_Click(Index As Integer)
WorkPref.CooOrigin = Index + 1
End Sub

Private Sub optDDROrientation_Click(Index As Integer)
Select Case Index
Case 0
     WorkPref.DRDefinition = glNormal
     lblUnderColor.Caption = "Suppressed"
     lblOverColor.Caption = "Induced"
Case 1
     WorkPref.DRDefinition = glReverse
     lblUnderColor.Caption = "Induced"
     lblOverColor.Caption = "Suppressed"
End Select
Iris
End Sub

Private Sub optERCalc_Click(Index As Integer)
    glbPreferencesExpanded.PairSearchOptions.SearchDef.ERCalcType = Index
End Sub

Private Sub optHorOrientation_Click(Index As Integer)
WorkPref.CooHOrientation = 3 - (Index + 1)
End Sub

Private Sub optIsoDataFrom_Click(Index As Integer)
Select Case Index
Case 0
     WorkPref.IsoDataField = isfMWMono      ' 7
Case 1
     WorkPref.IsoDataField = isfMWAvg       ' 6
Case 2
     WorkPref.IsoDataField = isfMWTMA       ' 8
End Select
End Sub

Private Sub optIsoICR2LS_Click(Index As Integer)
If Index = 0 Then
   WorkPref.IsoICR2LSMOverZ = True
Else
   WorkPref.IsoICR2LSMOverZ = False
End If
End Sub

Private Sub optSelection_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
    If glSelColor = -1 Then
       Call GetColorAPIDlg(Me.hwnd, glSelColor)
    End If
    lblClr(4).BackColor = glSelColor
Case 1
    glSelColor = -1
End Select
End Sub

Private Sub optTwoCloseResults_Click(Index As Integer)
WorkPref.Case2Results = Index
End Sub

Private Sub optType_Click(Index As Integer)
WorkPref.CooType = Index
End Sub

Private Sub optVAxis_Click(Index As Integer)
WorkPref.CooVAxisScale = Index
End Sub

Private Sub optVerOrientation_Click(Index As Integer)
WorkPref.CooVOrientation = Index + 1
End Sub


Private Sub picOpt_Click(Index As Integer)
    If Index = 1 Then
        MsgBox "Please click on the word Suppressed, Normal, or Induced to define the expressino ratio colors.  Also, note that you can swap the induced and suppressed colors by changing the Expression Ratio Definition on the Data tab.", vbInformation + vbOKOnly, glFGTU
    End If
End Sub

Private Sub tbsOptions_Click(PreviousTab As Integer)
    PositionControls
End Sub

Private Sub txtAbuAspectRatio_LostFocus()
If (Not IsNumeric(txtAbuAspectRatio.Text) Or (val(txtAbuAspectRatio.Text) <= 0)) Then
   MsgBox "This argument should be number > 0.", vbOKOnly
   TryAgain txtAbuAspectRatio
Else
   WorkPref.AbuAspectRatio = val(txtAbuAspectRatio.Text)
End If
End Sub

'No longer supported on this form (March 2006)
''Private Sub txtLegacyDBPath_Change()
''glbPreferencesExpanded.LegacyAMTDBPath = Trim(txtLegacyDBPath.Text)
''End Sub

Private Sub txtAutoSizeMultiplier_LostFocus()
If Not IsNumeric(txtMinPointSize.Text) Then
   MsgBox "This argument should be a number from the range [.001 to 1000].", vbOKOnly
   txtMinPointSize.Text = 1
   txtMinPointSize.SetFocus
Else
   ValidateTextboxValueDbl txtAutoSizeMultiplier, 0.001, 1000, 1
   glbPreferencesExpanded.AutoSizeMultiplier = val(txtAutoSizeMultiplier.Text)
End If

End Sub

Private Sub txtDBTolerance_LostFocus()
If Not IsNumeric(txtDBTolerance.Text) Then
   If Len(Trim(txtDBTolerance.Text)) = 0 Then
      WorkPref.DBTolerance = -1
   Else
      MsgBox "This argument should be number or left blank.", vbOKOnly
      TryAgain txtDBTolerance
   End If
Else
   WorkPref.DBTolerance = val(txtDBTolerance.Text)
End If
End Sub

Private Sub txtDDRatioMax_LostFocus()
If (Not IsNumeric(txtDDRatioMax.Text) Or (val(txtDDRatioMax.Text) <= 0)) Then
   MsgBox "This argument should be number > 0.", vbOKOnly
   TryAgain txtDDRatioMax
Else
   glDDRatioMax = val(txtDDRatioMax.Text)
   Call DDInitColors
End If
End Sub

Private Sub txtDupTolerance_LostFocus()
If Not IsNumeric(txtDupTolerance.Text) Then
   If Len(Trim(txtDupTolerance.Text)) = 0 Then
      WorkPref.DupTolerance = -1
   Else
      MsgBox "This argument should be number or left blank.", vbOKOnly
      TryAgain txtDupTolerance
   End If
Else
   WorkPref.DupTolerance = val(txtDupTolerance.Text)
End If
End Sub

' No longer supported (March 2006)
''Private Sub txtFTICR_AMTPath_Change()
''sFTICR_AMTPath = Trim$(txtFTICR_AMTPath.Text)
''End Sub

Private Sub txtICR2LSExePath_LostFocus()
sICR2LSCommand = Trim$(txtICR2LSExePath.Text)
End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    
On Error Resume Next

    If Not Me.WindowState = vbMinimized Then
        lngDesiredValue = Me.width - tbsOptions.Left - 240
        If lngDesiredValue < 6125 Then lngDesiredValue = 6125
        
        tbsOptions.width = lngDesiredValue
        
        If picOpt(2).Left > 0 Then
            picOpt(2).width = lngDesiredValue - picOpt(2).Left - 120
            
            fraICR2LS.width = tbsOptions.width - fraICR2LS.Left - 360
            
            cmdBrowseForICR2LS.Left = fraICR2LS.width - cmdBrowseForICR2LS.width - 120
            
            txtICR2LSExePath.width = cmdBrowseForICR2LS.Left - txtICR2LSExePath.Left - 120
        
''            fraAMTLocation.width = fraICR2LS.width
''            cmdBrowseForLegacyDB.Left = cmdBrowseForICR2LS.Left
''            txtLegacyDBPath.width = txtICR2LSExePath.width
        
        End If
    End If
    
End Sub
    
Private Sub Settings()
On Error Resume Next
With WorkPref
'SWITCHES
    Select Case .IsoDataField
    Case 6
         optIsoDataFrom(1).Value = True
    Case 8
         optIsoDataFrom(2).Value = True
    Case 7
         optIsoDataFrom(0).Value = True
    End Select
    
    If Not (.Case2Results = 0 Or .Case2Results = 1) Then
        .Case2Results = 2
    End If
    
    optTwoCloseResults(.Case2Results).Value = True
    
    If .IsoICR2LSMOverZ Then
       optIsoICR2LS(0).Value = True
    Else
       optIsoICR2LS(1).Value = True
    End If
'TOLERANCES
    If .DBTolerance = -1 Then
       txtDBTolerance.Text = ""
    Else
       txtDBTolerance.Text = .DBTolerance
    End If
    
    If .DupTolerance = -1 Then
       txtDupTolerance.Text = ""
    Else
       txtDupTolerance.Text = .DupTolerance
    End If
    
    If .IsoDataFit = -1 Then
       txtIsoDataFit.Text = ""
    Else
       txtIsoDataFit.Text = .IsoDataFit
    End If
'ICR2LS
    txtICR2LSExePath.Text = sICR2LSCommand
    
    If IsWinLoaded(ICR2LSCaption) Then 'no changes if ICR-2LS loaded
       txtICR2LSExePath.Enabled = False
       lblICR2LSExePath.Caption = "ICR-2LS already resident in memory; path not updatable"
       cmdBrowseForICR2LS.Enabled = False
    Else
       txtICR2LSExePath.Enabled = True
       lblICR2LSExePath.Caption = "ICR-2LS.Exe Path"
       cmdBrowseForICR2LS.Enabled = True
    End If
    
'AMT & FTICR_AMT databases
    'No longer supported on this form (March 2006)
    ''txtLegacyDBPath.Text = glbPreferencesExpanded.LegacyAMTDBPath
    
    ' No longer supported (March 2006)
    ''txtFTICR_AMTPath.Text = sFTICR_AMTPath
    
'COORDINATE SYSTEM
    optType(.CooType).Value = True
    optVAxis(.CooVAxisScale).Value = True
    Select Case .CooHOrientation
    Case glNormal
        optHorOrientation(1).Value = True
    Case glReverse
        optHorOrientation(0).Value = True
    End Select
    optVerOrientation(.CooVOrientation - 1).Value = True
    optCooSysCross(.CooOrigin - 1).Value = True
'DRAWING AND COLORS
    txtDDRatioMax.Text = glDDRatioMax
    txtMaxPointFactor.Text = .MaxPointFactor
    txtMinPointSize.Text = .MinPointFactor
    txtAbuAspectRatio.Text = .AbuAspectRatio
    SetCheckBox chkAutoSize, glbPreferencesExpanded.AutoAdjSize
    txtAutoSizeMultiplier = glbPreferencesExpanded.AutoSizeMultiplier
    SetCheckBox chkBorderColor, .BorderClrSameAsInt
    OldDRDefinition = .DRDefinition
    optDDROrientation(.DRDefinition - 1).Value = True
    
    lblClr(0).BackColor = glBackColor
    lblClr(1).BackColor = glForeColor
    lblClr(2).BackColor = glCSColor
    lblClr(3).BackColor = glIsoColor

    CSMarkerChange
    IsoMarkerChange
    Iris
End With
'selection
If glSelColor > 0 Then
   optSelection(0).Value = True
   lblClr(4).BackColor = glSelColor
Else
   optSelection(1).Value = True
End If
'writing PEH from gel
SetCheckBox chkWriteFreqShift, glWriteFreqShift
optERCalc(glbPreferencesExpanded.PairSearchOptions.SearchDef.ERCalcType).Value = True

End Sub

Private Sub SaveOldSettings()
OldICR2LSCommandLine = sICR2LSCommand
OldCSColor = glCSColor
OldIsoColor = glIsoColor
OldBackColor = glBackColor
OldUnderColor = glUnderColor
OldOverColor = glOverColor
OldMidColor = glMidColor
OldCSShape = glCSShape
OldIsoShape = glIsoShape
OldDDRatioMax = glDDRatioMax
OldSelColor = glSelColor

'No longer supported on this form (March 2006)
''OldAMTPath = glbPreferencesExpanded.LegacyAMTDBPath

' No longer supported (March 2006)
''OldFTICR_AMTPath = sFTICR_AMTPath

OldWriteFreqShift = glWriteFreqShift
OldAutoAdjSize = glbPreferencesExpanded.AutoAdjSize
OldAutoSizeMultiplier = glbPreferencesExpanded.AutoSizeMultiplier
OldERCalcType = glbPreferencesExpanded.PairSearchOptions.SearchDef.ERCalcType
OldMTType = PresentDisplay0Type
OldMTScanMin = Display0MaxScan
OldMTScanMax = Display0MinScan
End Sub

Private Sub RestoreOldSettings()
sICR2LSCommand = OldICR2LSCommandLine
glCSColor = OldCSColor
glIsoColor = OldIsoColor
glBackColor = OldBackColor
glForeColor = OldForeColor
glUnderColor = OldUnderColor
glOverColor = OldOverColor
glMidColor = OldMidColor
glCSShape = OldCSShape
glIsoShape = OldIsoShape
glDDRatioMax = OldDDRatioMax
glSelColor = OldSelColor

' No longer supported (March 2006)
''glbPreferencesExpanded.LegacyAMTDBPath = OldAMTPath
''sFTICR_AMTPath = OldFTICR_AMTPath

glWriteFreqShift = OldWriteFreqShift
glbPreferencesExpanded.AutoAdjSize = OldAutoAdjSize
glbPreferencesExpanded.AutoSizeMultiplier = OldAutoSizeMultiplier
glbPreferencesExpanded.PairSearchOptions.SearchDef.ERCalcType = OldERCalcType
PresentDisplay0Type = OldMTType
Display0MaxScan = OldMTScanMin
Display0MinScan = OldMTScanMax
End Sub

Public Sub Iris()
Dim i As Integer
Dim sLegTop As Double
Dim sLegHeight As Double
Dim sUnitWidth As Double
Dim sCenterX As Double
On Error Resume Next
sLegTop = lblUnderColor.Top + lblUnderColor.Height + 20
sLegHeight = 195
sUnitWidth = picOpt(1).ScaleWidth / 100
sCenterX = picOpt(1).ScaleWidth / 2
With picOpt(1)
    .DrawStyle = vbTransparent
    .DrawMode = vbCopyPen
    .FillStyle = vbFSSolid
End With
Select Case WorkPref.DRDefinition
Case glNormal
    For i = -50 To 50
        picOpt(1).ForeColor = aDDColors(i)
        picOpt(1).FillColor = aDDColors(i)
        picOpt(1).Line (sCenterX + i * sUnitWidth, sLegTop)-(sCenterX + (i + 1) * sUnitWidth, sLegTop + sLegHeight), , B
    Next i
Case glReverse
    For i = -50 To 50
        picOpt(1).ForeColor = aDDColors(-i)
        picOpt(1).FillColor = aDDColors(-i)
        picOpt(1).Line (sCenterX + i * sUnitWidth, sLegTop)-(sCenterX + (i + 1) * sUnitWidth, sLegTop + sLegHeight), , B
    Next i
End Select
End Sub

Private Sub txtIsoDataFit_LostFocus()
If Not IsNumeric(txtIsoDataFit.Text) Then
   If Len(Trim(txtIsoDataFit.Text)) = 0 Then
      WorkPref.IsoDataFit = -1
   Else
      MsgBox "This argument should be number or left blank.", vbOKOnly
      TryAgain txtIsoDataFit
   End If
Else
   WorkPref.IsoDataFit = val(txtIsoDataFit.Text)
End If
End Sub

Private Sub txtMaxPointFactor_LostFocus()
If (Not IsNumeric(txtMaxPointFactor.Text) Or (val(txtMaxPointFactor.Text) <= 0)) Then
   MsgBox "This argument should be number > 0.", vbOKOnly
   TryAgain txtMaxPointFactor
Else
   WorkPref.MaxPointFactor = val(txtMaxPointFactor.Text)
End If
End Sub

Private Sub txtMinPointSize_LostFocus()
If Not IsNumeric(txtMinPointSize.Text) Then
   MsgBox "This argument should be number from the range [0,1].", vbOKOnly
   txtMinPointSize.SetFocus
Else
   WorkPref.MinPointFactor = Abs(val(txtMinPointSize.Text))
   txtMinPointSize.Text = WorkPref.MinPointFactor
End If
End Sub

Private Function ColorsShapesChanged() As Boolean
ColorsShapesChanged = True
If OldCSColor <> glCSColor Then Exit Function
If OldIsoColor <> glIsoColor Then Exit Function
If OldBackColor <> glBackColor Then Exit Function
If OldUnderColor <> glUnderColor Then Exit Function
If OldOverColor <> glOverColor Then Exit Function
If OldMidColor <> glMidColor Then Exit Function
If OldCSShape <> glCSShape Then Exit Function
If OldIsoShape <> glIsoShape Then Exit Function
If OldSelColor <> glSelColor Then Exit Function
ColorsShapesChanged = False
End Function

Private Sub TryAgain(ctlTxt As TextBox)
ctlTxt.SetFocus
ctlTxt.Text = ""
End Sub

Private Sub CSMarkerChange()
Dim i As Integer
For i = 0 To imCSMarker.UBound
    If i = glCSShape Then
       imCSMarker(i).Visible = True
    Else
       imCSMarker(i).Visible = False
    End If
Next i
End Sub

Private Sub IsoMarkerChange()
Dim i As Integer
For i = 0 To imIsoMarker.UBound
    If i = glIsoShape Then
       imIsoMarker(i).Visible = True
    Else
       imIsoMarker(i).Visible = False
    End If
Next i
End Sub

Private Sub txtMTScanMax_LostFocus()
If Not IsNumeric(txtMTScanMax.Text) Then
   MsgBox "This argument should be positive integer.", vbOKOnly
   txtMTScanMax.SetFocus
Else
   Display0MaxScan = CLng(txtMTScanMax.Text)
End If
End Sub

Private Sub txtMTScanMin_LostFocus()
If Not IsNumeric(txtMTScanMin.Text) Then
   MsgBox "This argument should be positive integer.", vbOKOnly
   txtMTScanMin.SetFocus
Else
   Display0MinScan = CLng(txtMTScanMin.Text)
End If
End Sub

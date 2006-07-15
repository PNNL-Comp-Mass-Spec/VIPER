VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmGraphOverlayPrintMetric 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Metric Settings"
   ClientHeight    =   4815
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3825
      ScaleWidth      =   5865
      TabIndex        =   65
      Top             =   480
      Width           =   5895
      Begin VB.Frame fraTextBoxes 
         Caption         =   "Text/Comments/Legend"
         Height          =   3375
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   5415
         Begin VB.TextBox txtTextFontHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   90
            Text            =   "0"
            Top             =   1950
            Width           =   615
         End
         Begin VB.ComboBox cmbTextName 
            Height          =   315
            ItemData        =   "frmGraphOverlayPrintMetric.frx":0000
            Left            =   240
            List            =   "frmGraphOverlayPrintMetric.frx":0002
            TabIndex        =   79
            Top             =   480
            Width           =   1935
         End
         Begin VB.ListBox lstTextForGraph 
            Height          =   1425
            Left            =   240
            TabIndex        =   78
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddText 
            Caption         =   "&Add"
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   77
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmdRemoveText 
            Caption         =   "&Remove"
            Height          =   315
            Left            =   240
            TabIndex        =   76
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtTextText 
            Height          =   735
            Left            =   2400
            MultiLine       =   -1  'True
            TabIndex        =   75
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtTextPosX1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   74
            Text            =   "0"
            Top             =   1290
            Width           =   615
         End
         Begin VB.TextBox txtTextPosY1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            TabIndex        =   73
            Text            =   "0"
            Top             =   1290
            Width           =   615
         End
         Begin VB.CommandButton cmdClearText 
            Caption         =   "&Clear"
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   72
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtTextWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   71
            Text            =   "0"
            Top             =   1620
            Width           =   615
         End
         Begin VB.TextBox txtTextHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            TabIndex        =   70
            Text            =   "0"
            Top             =   1620
            Width           =   615
         End
         Begin VB.CheckBox chkTextShowBorder 
            Caption         =   "Show Border"
            Height          =   255
            Left            =   2520
            TabIndex        =   69
            Top             =   2580
            Width           =   1335
         End
         Begin VB.CheckBox chkTextItalic 
            Caption         =   "Italic"
            Height          =   255
            Left            =   2520
            TabIndex        =   68
            Top             =   2280
            Width           =   735
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "&Accept"
            Height          =   315
            Left            =   2400
            TabIndex        =   67
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Font Height (mm)"
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   89
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Name"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   88
            Top             =   240
            Width           =   615
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   2280
            X2              =   2280
            Y1              =   240
            Y2              =   3240
         End
         Begin VB.Label Label4 
            Caption         =   "Text"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   87
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Position X,Y (mm)"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   86
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Active Text"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   85
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Width, Height (mm)"
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   84
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Back Color"
            Height          =   210
            Index           =   5
            Left            =   3960
            TabIndex        =   83
            Top             =   2580
            Width           =   855
         End
         Begin VB.Label lblTextBackColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   82
            ToolTipText     =   "Double-click to change"
            Top             =   2565
            Width           =   255
         End
         Begin VB.Label lblTextForeColor 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4920
            TabIndex        =   81
            ToolTipText     =   "Double-click to change"
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Fore Color"
            Height          =   210
            Index           =   6
            Left            =   3960
            TabIndex        =   80
            Top             =   2295
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdFitMetricTo 
      Caption         =   "Fit Range"
      Height          =   315
      Left            =   2280
      TabIndex        =   53
      ToolTipText     =   "Fit range to current metric and page"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdFitPage 
      Caption         =   "Fit Page"
      Height          =   315
      Left            =   1200
      TabIndex        =   52
      ToolTipText     =   "Fit page to current metric and range"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdFitMetric 
      Caption         =   "Fit Metric"
      Height          =   315
      Left            =   120
      TabIndex        =   51
      ToolTipText     =   "Fit metric to current page and range"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   3960
      TabIndex        =   26
      Top             =   4440
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3825
      ScaleWidth      =   5865
      TabIndex        =   3
      ToolTipText     =   "Fit page to current metric and range"
      Top             =   480
      Width           =   5895
      Begin VB.Frame fraCurrentPage 
         Caption         =   "Current Page(m)"
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   1695
         Begin VB.TextBox txtPageHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   43
            Text            =   "0"
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtPageWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   42
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Height"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Width"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame fraCurrentRange 
         Caption         =   "Current Range"
         Height          =   1575
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1695
         Begin VB.TextBox txtMaxET 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   38
            Text            =   "0"
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtMinET 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   37
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtMaxMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   36
            Text            =   "0"
            Top             =   1140
            Width           =   735
         End
         Begin VB.TextBox txtMinMW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   35
            Text            =   "0"
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Max ET"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Min ET"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Max MW"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Min MW"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   900
            Width           =   735
         End
      End
      Begin VB.Frame fraMetric 
         Caption         =   "Metric(mm if not otherwise indicated)"
         Height          =   3615
         Left            =   1920
         TabIndex        =   7
         Top             =   120
         Width           =   3855
         Begin VB.TextBox txtOriginVert 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            LinkTimeout     =   0
            TabIndex        =   64
            Text            =   "75"
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtOriginHorz 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            LinkTimeout     =   0
            TabIndex        =   63
            Text            =   "100"
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtETDecPlaces 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   60
            Text            =   "3"
            Top             =   3060
            Width           =   495
         End
         Begin VB.TextBox txtMWDecPlaces 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   59
            Text            =   "3"
            Top             =   2760
            Width           =   495
         End
         Begin VB.ComboBox cmbSizeType 
            Height          =   315
            ItemData        =   "frmGraphOverlayPrintMetric.frx":0004
            Left            =   840
            List            =   "frmGraphOverlayPrintMetric.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtGridET_R 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   50
            Text            =   "0.01"
            Top             =   3060
            Width           =   615
         End
         Begin VB.TextBox txtGridMW_R 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   49
            Text            =   "20"
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtMaxSpotHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   46
            Text            =   "5"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtFontLabelsHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   44
            Text            =   "50"
            Top             =   2460
            Width           =   495
         End
         Begin VB.TextBox txtDaMetric 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   16
            Text            =   "100"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtETMetric 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   15
            Text            =   "1000"
            Top             =   540
            Width           =   615
         End
         Begin VB.TextBox txtSpotWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Text            =   "1"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtSpotHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   13
            Text            =   "1"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtStickWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   12
            Text            =   "0.5"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtFontWidthRatio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Text            =   "0.2"
            ToolTipText     =   "Set to 0 to let device chose based on aspect ratio"
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtFontHeight 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Text            =   "4"
            Top             =   1860
            Width           =   495
         End
         Begin VB.TextBox txtHorzMargin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            LinkTimeout     =   0
            TabIndex        =   9
            Text            =   "10"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtVertMargin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Text            =   "5"
            ToolTipText     =   "Vertical offset from the page edge"
            Top             =   540
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Origin(vert.)"
            Height          =   255
            Index           =   17
            Left            =   2040
            TabIndex        =   62
            ToolTipText     =   "Horizontal offset from the page egde"
            Top             =   2200
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Origin(horz.)"
            Height          =   255
            Index           =   16
            Left            =   2040
            TabIndex        =   61
            ToolTipText     =   "Horizontal offset from the page egde"
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "ET dec. places"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   58
            Top             =   3100
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "MW dec. places"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Top             =   2800
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Font lbl. height"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   2505
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Spot size"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   54
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Grid ET(NET)"
            Height          =   255
            Index           =   9
            Left            =   2040
            TabIndex        =   48
            Top             =   3100
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Grid MW(Da)"
            Height          =   255
            Index           =   7
            Left            =   2040
            TabIndex        =   47
            Top             =   2800
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Max spot ht."
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   45
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "1Da"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "1ET"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Spot width"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Spot height"
            Height          =   255
            Index           =   6
            Left            =   2040
            TabIndex        =   22
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Stick width"
            Height          =   255
            Index           =   8
            Left            =   2040
            TabIndex        =   21
            Top             =   1620
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Font wd/ht ratio"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   20
            Top             =   2200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Font height"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   19
            Top             =   1905
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Horz.margin"
            Height          =   255
            Index           =   12
            Left            =   2040
            TabIndex        =   18
            ToolTipText     =   "Horizontal offset from the page egde"
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Vert.margin"
            Height          =   255
            Index           =   13
            Left            =   2040
            TabIndex        =   17
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame fraOrientation 
         Caption         =   "Orientation"
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   1695
         Begin VB.OptionButton optOrientation 
            Caption         =   "ET Horizontal"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   6
            ToolTipText     =   "ET Horizontal / MW Vertical"
            Top             =   240
            Width           =   1515
         End
         Begin VB.OptionButton optOrientation 
            Caption         =   "ET Vertical"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   5
            ToolTipText     =   "MW Horz / ET Vert"
            Top             =   540
            Value           =   -1  'True
            Width           =   1515
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   0
      Left            =   120
      ScaleHeight     =   3825
      ScaleWidth      =   5865
      TabIndex        =   2
      Top             =   480
      Width           =   5895
      Begin VB.ComboBox cmbPrinters 
         Height          =   315
         ItemData        =   "frmGraphOverlayPrintMetric.frx":0008
         Left            =   120
         List            =   "frmGraphOverlayPrintMetric.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrinterProperties 
         Caption         =   "Properties"
         Height          =   315
         Left            =   3960
         TabIndex        =   27
         Top             =   220
         Width           =   1215
      End
      Begin VB.Label lblPrinterInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Printer Info"
         Height          =   855
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   4575
         WordWrap        =   -1  'True
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7435
      TabWidthStyle   =   2
      Style           =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Printer"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Metric"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc."
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5040
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
End
Attribute VB_Name = "frmGraphOverlayPrintMetric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'definition of printing in metric mode
'created: 01/02/2003 nt
'last modified: 01/20/2003 nt
'-------------------------------------------------------------------
Option Explicit

Const WIDTH_OFFSET_PCT = 0.5         '50% of label font width
Const HEIGHT_OFFSET_PCT = 0.5        '50% of label font height
Const NAME_OFFSET_PCT = 0.25         '50% of name font height

Const TXT_LEGEND = "Legend"
Const TXT_DATE = "Date"
Const TXT_DATETIME = "Date/Time"
Const TXT_PNNL = "PNNL"
Const TXT_SCAN_METRIC = "Scan Metric"       'scans per milimeter info

Const PNNL = "Pacific Northwest National Laboratory"

Const TXT_OFFSET_H = 1               'milimetar
Const TXT_OFFSET_V = 1

'public properties
Dim mMinMW As Double
Dim mmaxmw As Double
Dim mMinET As Double
Dim mMaxET As Double

Dim mMinMW_mm As Double
Dim mmaxmw_mm As Double
Dim mMinET_mm As Double
Dim mMaxET_mm As Double

Dim mDa_mm As Double
Dim mET_mm As Double

Dim mSpotWidth As Double
Dim mSpotHeight As Double

Dim mStickWidth As Double

Dim mFontWidthRatio As Double           'ratio of font width and height
Dim mFontHeight As Double               'font height in mm
Dim mFontLabelHeight As Double          'font height for labels in mm
                                        'use same ratio as name fonts

Dim mOriginHorz As Double               'coordinates of the origin
Dim mOriginVert As Double               'ori

'NOTE: origin is always in the bottom left corner and includes when
'calculating page size horizontal and vertical margins

Dim mHorzMargin As Double               'NOT SURE I WILL USE THIS
Dim mVertMargin As Double


Dim mOrientation As Long

Dim mPageWidth As Double
Dim mPageHeight As Double

Dim mMaxSpotHeight As Double            'use only if size is not uniform

Dim mGridMW_R As Double               'labels in Da       put label(and grid) on each
Dim mGridET_R As Double               'labels in ET       position divisible with number

Dim mMWDecPlaces As Long
Dim mETDecPlaces As Long

Dim mSizeScaleType As ScaleType

Dim mMinAbu As Double
Dim mMaxAbu As Double
Dim mAdjAbuRange As Double            'depends on scaling type

'logical coordinates in HI_METRIC coordinate system(calculated; read-only)
'all client drawing should be done with these properties
Dim mMinETL As Long
Dim mMaxETL As Long
Dim mMinMWL As Long
Dim mmaxmwL As Long
Dim mOriginHorzL As Long
Dim mOriginVertL As Long
Dim mMWRangeL As Long
Dim mETRangeL As Long
Dim mSpotWidthL As Long
Dim mSpotHeightL As Long
Dim mMaxSpotHeightL As Long
Dim mStickWidthL As Long
Dim mFontHeightL As Long
Dim mFontLabelHeightL As Long
Dim mFontWidthL As Long
Dim mFontLabelWidthL As Long
Dim mNameOffsetL As Long
Dim mHorzOffsetL As Long
Dim mVertOffsetL As Long

Public IsCancel As Long
Public TextBoxes As Collection          'collection of text boxes to be displayed with graph

Dim CurrTextBoxName As String
Dim CurrTextBoxInd As Long

Private Sub cmbPrinters_Click()
Set Printer = Printers(cmbPrinters.ListIndex)
Call SetPrinterInfo
End Sub

Private Sub cmbSizeType_Click()
mSizeScaleType = cmbSizeType.ListIndex      'no need to invoke property Let
Call SetAdjAbuRange
End Sub

Private Sub cmdAccept_Click()
If Len(CurrTextBoxName) > 0 Then
   GetTextProperties CurrTextBoxName
Else
   MsgBox "Select text box to edit from the list.", vbOKOnly, glFGTU
End If
End Sub

Private Sub cmdAddText_Click(Index As Integer)
Dim NewName  As String
Dim tb As TextBoxGraph
Dim i As Long
On Error Resume Next
NewName = cmbTextName.Text
Set tb = TextBoxes(NewName)
If tb Is Nothing Then               'add new text box to collection
   If Len(NewName) > 0 Then
      Set tb = New TextBoxGraph
      tb.Name = NewName
      Select Case NewName
      Case TXT_LEGEND
           tb.AddTextLine "1 Da = " & mDa_mm & " mm"
           tb.AddTextLine "1 ET = " & mET_mm & " mm"
           tb.ShowBorder = True
           tb.FontHeightL = 4 * mFontHeightL
           tb.FontWidthL = 4 * mFontWidthL
           tb.FontItalic = True
      Case TXT_DATE
           tb.AddTextLine "Date: " & date
           tb.FontHeightL = mFontLabelHeightL
           tb.FontWidthL = mFontLabelWidthL
      Case TXT_DATETIME
           tb.AddTextLine "Date/Time: " & Now
           tb.FontHeightL = mFontLabelHeightL
           tb.FontWidthL = mFontLabelWidthL
      Case TXT_PNNL
           tb.AddTextLine PNNL
           tb.ShowBorder = False
           tb.FontHeightL = mFontLabelHeightL
           tb.FontWidthL = mFontLabelWidthL
      Case TXT_SCAN_METRIC
           For i = 0 To OlyCnt - 1
               If Oly(i).DisplayInd <> 0 Then
                  tb.AddTextLine GetScanMetric(i)
               End If
           Next i
           tb.FontHeightL = 4 * mFontHeightL
           tb.FontWidthL = 4 * mFontWidthL
      Case Else
           tb.FontHeightL = 4 * mFontHeightL
           tb.FontWidthL = 4 * mFontWidthL
      End Select
      TextBoxes.add tb, tb.Name
      lstTextForGraph.AddItem tb.Name
      lstTextForGraph.ListIndex = lstTextForGraph.NewIndex      'this will set some properties
      GetTextProperties (tb.Name)
   Else
      MsgBox "Text box must have a name. Select from the list or type new name.", vbOKOnly, glFGTU
   End If
Else                                'name already exists
   MsgBox "Text box named " & NewName & "already exists.", vbOKOnly, glFGTU
End If
End Sub

Private Sub cmdCancel_Click()
IsCancel = True
Me.Hide
End Sub

Private Sub cmdClearText_Click(Index As Integer)
On Error Resume Next
lstTextForGraph.Clear
Do While TextBoxes.Count > 0
   TextBoxes.Remove (1)
Loop
Set TextBoxes = New Collection
txtTextText.Text = ""
End Sub

Private Sub cmdFitPage_Click()
'-----------------------------------------------------------------------
'calculate page size based on current range and metric settings
'-----------------------------------------------------------------------
On Error Resume Next
Select Case mOrientation
Case OrientMWVrtETHrz
     PageWidth = (mMaxET - mMinET) * mET_mm + mHorzMargin + mOriginHorz
     PageHeight = (mmaxmw - mMinMW) * mDa_mm + mVertMargin + mOriginVert
Case OrientMWHrzETVrt
     PageWidth = (mmaxmw - mMinMW) * mDa_mm + mHorzMargin + mOriginHorz
     PageHeight = (mMaxET - mMinET) * mET_mm + mVertMargin + mOriginVert
End Select
End Sub

Private Sub cmdPrint_Click()
IsCancel = False
Me.Hide
End Sub

Private Sub cmdPrinterProperties_Click()
Dim PrtDefs As PRINTER_DEFAULTS
Dim hPrt As Long
Dim Res As Long
On Error Resume Next
PrtDefs.DesiredAccess = PRINTER_ACCESS_USE
PrtDefs.pDevMode = 0        'use defaults
PrtDefs.pDatatype = vbNullString
Res = OpenPrinter(Printer.DeviceName, hPrt, PrtDefs)
If hPrt = 0 Then
   MsgBox "Can not open selected printer.", vbOKOnly, glFGTU
Else
   Call PrinterProperties(Me.hwnd, hPrt)
   Call ClosePrinter(hPrt)
End If
End Sub

Private Sub cmdRemoveText_Click()
On Error Resume Next
If Len(CurrTextBoxName) > 0 Then
   lstTextForGraph.RemoveItem CurrTextBoxInd
   TextBoxes.Remove CurrTextBoxName
   CurrTextBoxName = ""
Else
   MsgBox "Select item from the list.", vbOKOnly, glFGTU
End If
End Sub

Private Sub Form_Load()
Dim DefPrtInd As Long
Dim i As Long
On Error Resume Next
Set TextBoxes = New Collection
Call FillTextNameCombo
Call FillSizeTypeCombo
'set defaults for public properties
mmPerDa = CDbl(txtDaMetric.Text)        'these are property let and will trigger
mmPerET = CDbl(txtETMetric.Text)        'setting of some read-only properties
SpotWidth = CDbl(txtSpotWidth.Text)
SpotHeight = CDbl(txtSpotHeight.Text)
StickWidth = CDbl(txtStickWidth.Text)
HorzMargin = CDbl(txtHorzMargin.Text)
VertMargin = CDbl(txtVertMargin.Text)
FontHeight = CDbl(txtFontHeight.Text)
FontLabelHeight = CDbl(txtFontLabelsHeight.Text)
FontWidthRatio = CDbl(txtFontWidthRatio.Text)
MaxSpotHeight = CDbl(txtMaxSpotHeight.Text)
GridMW_R = CDbl(txtGridMW_R.Text)
GridET_R = CDbl(txtGridET_R.Text)
MWDecPlaces = CLng(txtMWDecPlaces.Text)
ETDecPlaces = CLng(txtETDecPlaces.Text)
OriginHorz = CLng(txtOriginHorz.Text)
OriginVert = CLng(txtOriginVert.Text)
cmbSizeType.ListIndex = 0
If optOrientation(OrientMWVrtETHrz).value Then
   mOrientation = OrientMWVrtETHrz
Else
   mOrientation = OrientMWHrzETVrt
End If

Printer.TrackDefault = True
Printer.ScaleMode = vbMillimeters
DefPrtInd = -1
For i = 0 To Printers.Count - 1
    cmbPrinters.AddItem Printers(i).DeviceName
    If Printers(i).DeviceName = Printer.DeviceName Then DefPrtInd = i
Next i
If DefPrtInd >= 0 Then cmbPrinters.ListIndex = DefPrtInd
Call cmdFitPage_Click
End Sub


Private Sub lblTextBackColor_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblTextBackColor_DblClick
TmpClr = lblTextBackColor.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblTextBackColor.BackColor = TmpClr
exit_lblTextBackColor_DblClick:
End Sub


Private Sub lblTextForeColor_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblTextForeColor_DblClick
TmpClr = lblTextForeColor.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblTextForeColor.BackColor = TmpClr
exit_lblTextForeColor_DblClick:
End Sub

Private Sub lstTextForGraph_Click()
CurrTextBoxInd = lstTextForGraph.ListIndex
CurrTextBoxName = lstTextForGraph.Text
SetTextProperties CurrTextBoxName
End Sub

Private Sub optOrientation_Click(Index As Integer)
mOrientation = Index
Call cmdFitPage_Click
End Sub

Private Sub TabStrip1_Click()
Dim i As Integer
For i = 0 To TabStrip1.Tabs.Count - 1
    If i = TabStrip1.SelectedItem.Index - 1 Then
       Picture1(i).Left = 120
       Picture1(i).Enabled = True
    Else
       Picture1(i).Left = -20000
       Picture1(i).Enabled = False
    End If
Next i
End Sub

Private Sub SetPrinterInfo()
Dim tmp As String
On Error Resume Next
tmp = tmp & "Name: "
tmp = tmp & Printer.DeviceName
tmp = tmp & vbCrLf

tmp = tmp & "Driver: "
tmp = tmp & Printer.DriverName
tmp = tmp & vbCrLf
     
tmp = tmp & "Quality: "
tmp = tmp & Printer.PrintQuality
tmp = tmp & vbCrLf

tmp = tmp & "Paper size: "
tmp = tmp & Printer.PaperSize
tmp = tmp & vbCrLf

lblPrinterInfo.Caption = tmp
End Sub

'---------properties listing------------------------
Public Property Get MinMW() As Double
MinMW = mMinMW
End Property

Public Property Let MinMW(ByVal dNewValue As Double)
mMinMW = dNewValue
mMinMW_mm = mMinMW * mDa_mm
mMinMWL = mMinMW_mm * HI_MM
mMWRangeL = (mmaxmw_mm - mMinMW_mm) * HI_MM
txtMinMW.Text = Format$(mMinMW, "0.00")
End Property

Public Property Get MaxMW() As Double
MaxMW = mmaxmw
End Property

Public Property Let MaxMW(ByVal dNewValue As Double)
mmaxmw = dNewValue
mmaxmw_mm = mmaxmw * mDa_mm
mmaxmwL = mmaxmw_mm * HI_MM
mMWRangeL = (mmaxmw_mm - mMinMW_mm) * HI_MM
txtMaxMW.Text = Format$(mmaxmw, "0.00")
End Property

Public Property Get MinET() As Double
MinET = mMinET
End Property

Public Property Let MinET(ByVal dNewValue As Double)
mMinET = dNewValue
mMinET_mm = mMinET * mET_mm
mMinETL = mMinET_mm * HI_MM
mETRangeL = (mMaxET_mm - mMinET_mm) * HI_MM
txtMinET.Text = Format$(mMinET, "0.00")
End Property

Public Property Get MaxET() As Double
MaxET = mMaxET
End Property

Public Property Let MaxET(ByVal dNewValue As Double)
mMaxET = dNewValue
mMaxET_mm = mMaxET * mET_mm
mMaxETL = mMaxET_mm * HI_MM
mETRangeL = (mMaxET_mm - mMinET_mm) * HI_MM
txtMaxET.Text = Format$(mMaxET, "0.00")
End Property

Public Property Get Orientation() As Long
Orientation = mOrientation
End Property

Public Property Let Orientation(ByVal lNewValue As Long)
optOrientation(lNewValue).value = True
End Property

Public Property Get PageWidth() As Double
PageWidth = mPageWidth
End Property

Public Property Let PageWidth(ByVal dNewValue As Double)
mPageWidth = dNewValue
txtPageWidth.Text = Format$(mPageWidth / 1000, "0.00")
End Property

Public Property Get PageHeight() As Double
PageHeight = mPageHeight
End Property

Public Property Let PageHeight(ByVal dNewValue As Double)
mPageHeight = dNewValue
txtPageHeight.Text = Format$(mPageHeight / 1000, "0.00")
End Property

Public Property Get mmPerDa() As Double
mmPerDa = mDa_mm
End Property

Public Property Let mmPerDa(ByVal dNewValue As Double)
mDa_mm = dNewValue
mMinMW_mm = mMinMW * mDa_mm
mmaxmw_mm = mmaxmw * mDa_mm
mMinMWL = mMinMW_mm * HI_MM
mmaxmwL = mmaxmw_mm * HI_MM
mMWRangeL = (mmaxmw_mm - mMinMW_mm) * HI_MM
txtDaMetric.Text = Format$(mDa_mm, "0.00")
End Property

Public Property Get mmPerET() As Double
mmPerET = mET_mm
End Property

Public Property Let mmPerET(ByVal dNewValue As Double)
mET_mm = dNewValue
mMinET_mm = mMinET * mET_mm
mMaxET_mm = mMaxET * mET_mm
mMinETL = mMinET_mm * HI_MM
mMaxETL = mMaxET_mm * HI_MM
mETRangeL = (mMaxET_mm - mMinET_mm) * HI_MM
txtETMetric.Text = Format$(mET_mm, "0.00")
End Property

Public Property Get HorzMargin() As Double
HorzMargin = mHorzMargin
End Property

Public Property Let HorzMargin(ByVal dNewValue As Double)
mHorzMargin = dNewValue
txtHorzMargin.Text = Format$(mHorzMargin, "0.00")
End Property

Public Property Get VertMargin() As Double
VertMargin = mVertMargin
End Property

Public Property Let VertMargin(ByVal dNewValue As Double)
mVertMargin = dNewValue
txtVertMargin.Text = Format$(mVertMargin, "0.00")
End Property

Public Property Get StickWidth() As Double
StickWidth = mStickWidth
End Property

Public Property Let StickWidth(ByVal dNewValue As Double)
mStickWidth = dNewValue
mStickWidthL = mStickWidth * HI_MM
txtStickWidth.Text = Format$(mStickWidth, "0.00")
End Property

Public Property Get SpotWidth() As Double
SpotWidth = mSpotWidth
End Property

Public Property Let SpotWidth(ByVal dNewValue As Double)
mSpotWidth = dNewValue
mSpotWidthL = mSpotWidth * HI_MM
txtSpotWidth.Text = Format$(mSpotWidth, "0.00")
End Property

Public Property Get SpotHeight() As Double
SpotHeight = mSpotHeight
End Property

Public Property Let SpotHeight(ByVal dNewValue As Double)
mSpotHeight = dNewValue
mSpotHeightL = mSpotHeight * HI_MM
txtSpotHeight.Text = Format$(mSpotHeight, "0.00")
End Property

Public Property Get FontWidthRatio() As Double
FontWidthRatio = mFontWidthRatio
End Property

Public Property Let FontWidthRatio(ByVal dNewValue As Double)
mFontWidthRatio = dNewValue
mFontWidthL = mFontHeightL * mFontWidthRatio
mFontLabelWidthL = mFontLabelHeightL * mFontWidthRatio
txtFontWidthRatio.Text = Format$(mFontWidthRatio, "0.00")
End Property

Public Property Get FontHeight() As Double
FontHeight = mFontHeight
End Property

Public Property Let FontHeight(ByVal dNewValue As Double)
mFontHeight = dNewValue
mFontHeightL = mFontHeight * HI_MM
mFontWidthL = mFontHeightL * mFontWidthRatio
mNameOffsetL = mFontHeight * NAME_OFFSET_PCT * HI_MM
txtFontHeight.Text = Format$(mFontHeight, "0.00")
End Property

Public Property Get FontLabelHeight() As Double
FontLabelHeight = mFontLabelHeight
End Property

Public Property Let FontLabelHeight(ByVal dNewValue As Double)
mFontLabelHeight = dNewValue
mFontLabelHeightL = mFontLabelHeight * HI_MM
mFontLabelWidthL = mFontLabelHeightL * mFontWidthRatio
mVertOffsetL = mFontLabelHeight * HEIGHT_OFFSET_PCT * HI_MM
mHorzOffsetL = mFontLabelHeight * WIDTH_OFFSET_PCT * HI_MM
txtFontLabelsHeight.Text = Format$(mFontLabelHeight, "0.00")
End Property

Public Property Get MaxSpotHeight() As Double
MaxSpotHeight = mMaxSpotHeight
End Property

Public Property Let MaxSpotHeight(ByVal dNewValue As Double)
mMaxSpotHeight = dNewValue
mMaxSpotHeightL = mMaxSpotHeight * HI_MM
txtMaxSpotHeight.Text = mMaxSpotHeight
End Property

Public Property Get ScaleSizeType() As ScaleType
ScaleSizeType = mSizeScaleType
End Property

Public Property Let ScaleSizeType(ByVal lNewValue As ScaleType)
mSizeScaleType = lNewValue
Call SetAdjAbuRange
cmbSizeType.ListIndex = mSizeScaleType
End Property

Public Property Get GridMW_R() As Double
GridMW_R = mGridMW_R
End Property

Public Property Let GridMW_R(ByVal dNewValue As Double)
mGridMW_R = dNewValue
txtGridMW_R.Text = mGridMW_R
End Property

Public Property Get GridET_R() As Double
GridET_R = mGridET_R
End Property

Public Property Let GridET_R(ByVal dNewValue As Double)
mGridET_R = dNewValue
txtGridET_R.Text = mGridET_R
End Property

Public Property Get MWDecPlaces() As Long
MWDecPlaces = mMWDecPlaces
End Property

Public Property Let MWDecPlaces(ByVal lNewValue As Long)
mMWDecPlaces = lNewValue
txtMWDecPlaces.Text = mMWDecPlaces
End Property

Public Property Get ETDecPlaces() As Long
ETDecPlaces = mETDecPlaces
End Property

Public Property Let ETDecPlaces(ByVal lNewValue As Long)
mETDecPlaces = lNewValue
txtETDecPlaces.Text = mETDecPlaces
End Property

Public Property Get OriginHorz() As Double
OriginHorz = mOriginHorz
End Property

Public Property Let OriginHorz(ByVal dNewValue As Double)
mOriginHorz = dNewValue
mOriginHorzL = mOriginHorz * HI_MM
txtOriginHorz.Text = mOriginHorz
End Property

Public Property Get OriginVert() As Double
OriginVert = mOriginVert
End Property

Public Property Let OriginVert(ByVal dNewValue As Double)
mOriginVert = dNewValue
mOriginVertL = mOriginVert * HI_MM
txtOriginVert.Text = mOriginVert
End Property


'next two properties do not have graphical interface to be set but are read-write
Public Property Get MinAbu() As Double
MinAbu = mMinAbu
End Property

Public Property Let MinAbu(ByVal dNewValue As Double)
On Error Resume Next
mMinAbu = dNewValue
Call SetAdjAbuRange
End Property

Public Property Get MaxAbu() As Double
MaxAbu = mMaxAbu
End Property

Public Property Let MaxAbu(ByVal dNewValue As Double)
On Error Resume Next
mMaxAbu = dNewValue
Call SetAdjAbuRange
End Property

'read-only properties
Public Property Get MinET_mm() As Double
MinET_mm = mMinET_mm
End Property

Public Property Get MaxET_mm() As Double
MaxET_mm = mMaxET_mm
End Property

Public Property Get MinMW_mm() As Double
MinMW_mm = mMinMW_mm
End Property

Public Property Get maxmw_mm() As Double
maxmw_mm = mmaxmw_mm
End Property

Public Property Get MWRangeL() As Long
MWRangeL = mMWRangeL
End Property

Public Property Get ETRangeL() As Long
ETRangeL = mETRangeL
End Property

Public Property Get OriginHorzL() As Long
OriginHorzL = mOriginHorzL
End Property

Public Property Get OriginVertL() As Long
OriginVertL = mOriginVertL
End Property

Public Property Get MinETL() As Long
MinETL = mMinETL
End Property

Public Property Get MaxETL() As Long
MaxETL = mMaxETL
End Property

Public Property Get MinMWL() As Long
MinMWL = mMinMWL
End Property

Public Property Get maxmwL() As Long
maxmwL = mmaxmwL
End Property

Public Property Get SpotWidthL() As Long
mSpotWidthL = mSpotWidthL
End Property

Public Property Get SpotHeightL() As Long
SpotHeightL = mSpotHeightL
End Property

Public Property Get MaxSpotHeightL() As Long
MaxSpotHeightL = mMaxSpotHeightL
End Property

Public Property Get StickWidthL() As Long
StickWidthL = mStickWidthL
End Property

Public Property Get FontHeightL() As Long
FontHeightL = mFontHeightL
End Property

Public Property Get FontLabelHeightL() As Long
FontLabelHeightL = mFontLabelHeightL
End Property

Public Property Get FontWidthL() As Long
FontWidthL = mFontWidthL
End Property

Public Property Get FontLabelWidthL() As Long
FontLabelWidthL = mFontLabelWidthL
End Property

Public Property Get NameOffsetL() As Long
NameOffsetL = mNameOffsetL
End Property

Public Property Get HorzOffsetL() As Long
HorzOffsetL = mHorzOffsetL
End Property

Public Property Get VertOffsetL() As Long
VertOffsetL = mVertOffsetL
End Property

'-----------------------------------------------end properties-------------

'--------------------------------------------------------------------------
'changes in settings affects also public properties

Private Sub txtFontLabelsHeight_LostFocus()
On Error Resume Next
If IsNumeric(txtFontLabelsHeight.Text) Then
   FontLabelHeight = Abs(CDbl(txtFontLabelsHeight.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtFontLabelsHeight.SetFocus
End If
End Sub


Private Sub txtDaMetric_LostFocus()
On Error Resume Next
If IsNumeric(txtDaMetric.Text) Then
   mmPerDa = Abs(CDbl(txtDaMetric.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtDaMetric.SetFocus
End If
End Sub

Private Sub txtETMetric_LostFocus()
On Error Resume Next
If IsNumeric(txtETMetric.Text) Then
   mmPerET = Abs(CDbl(txtETMetric.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtETMetric.SetFocus
End If
End Sub

Private Sub txtHorzMargin_LostFocus()
On Error Resume Next
If IsNumeric(txtHorzMargin.Text) Then
   HorzMargin = Abs(CDbl(txtHorzMargin.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtHorzMargin.SetFocus
End If
End Sub


Private Sub txtGridET_R_LostFocus()
On Error Resume Next
If IsNumeric(txtGridET_R.Text) Then
   GridET_R = Abs(CDbl(txtGridET_R.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtGridET_R.SetFocus
End If
End Sub

Private Sub txtGridMW_R_LostFocus()
On Error Resume Next
If IsNumeric(txtGridMW_R.Text) Then
   GridMW_R = Abs(CDbl(txtGridMW_R.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtGridMW_R.SetFocus
End If
End Sub

Private Sub txtMaxET_LostFocus()
On Error Resume Next
If IsNumeric(txtMaxET.Text) Then
   MaxET = CDbl(txtMaxET.Text)
Else
   MsgBox Msg_Numeric, vbOKOnly, glFGTU
   txtMaxET.SetFocus
End If
End Sub

Private Sub txtmaxmw_LostFocus()
On Error Resume Next
If IsNumeric(txtMaxMW.Text) Then
   MaxMW = Abs(CDbl(txtMaxMW.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtMaxMW.SetFocus
End If
End Sub


Private Sub txtMaxSpotHeight_LostFocus()
On Error Resume Next
If IsNumeric(txtMaxSpotHeight.Text) Then
   MaxSpotHeight = Abs(CDbl(txtMaxSpotHeight.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtMaxSpotHeight.SetFocus
End If
End Sub


Private Sub txtMinET_LostFocus()
On Error Resume Next
If IsNumeric(txtMinET.Text) Then
   MinET = CDbl(txtMinET.Text)
Else
   MsgBox Msg_Numeric, vbOKOnly, glFGTU
   txtMinET.SetFocus
End If
End Sub


Private Sub txtMinMW_LostFocus()
On Error Resume Next
If IsNumeric(txtMinMW.Text) Then
   MinMW = Abs(CDbl(txtMinMW.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtMinMW.SetFocus
End If
End Sub


Private Sub txtFontHeight_LostFocus()
On Error Resume Next
If IsNumeric(txtFontHeight.Text) Then
   FontHeight = Abs(CDbl(txtFontHeight.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtFontHeight.SetFocus
End If
End Sub


Private Sub txtFontWidthRatio_LostFocus()
On Error Resume Next
If IsNumeric(txtFontWidthRatio.Text) Then
   FontWidthRatio = Abs(CDbl(txtFontWidthRatio.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtFontWidthRatio.SetFocus
End If
End Sub

Private Sub txtMWDecPlaces_LostFocus()
On Error Resume Next
If IsNumeric(txtMWDecPlaces.Text) Then
   MWDecPlaces = Abs(CLng(txtMWDecPlaces.Text))
Else
   MsgBox Msg_Integer, vbOKOnly, glFGTU
   txtMWDecPlaces.SetFocus
End If
End Sub

Private Sub txtETDecPlaces_LostFocus()
On Error Resume Next
If IsNumeric(txtETDecPlaces.Text) Then
   ETDecPlaces = Abs(CLng(txtETDecPlaces.Text))
Else
   MsgBox Msg_Integer, vbOKOnly, glFGTU
   txtETDecPlaces.SetFocus
End If
End Sub

Private Sub txtOriginHorz_LostFocus()
On Error Resume Next
If IsNumeric(txtOriginHorz.Text) Then
   OriginHorz = Abs(CDbl(txtOriginHorz.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtOriginHorz.SetFocus
End If
End Sub


Private Sub txtOriginVert_LostFocus()
On Error Resume Next
If IsNumeric(txtOriginVert.Text) Then
   OriginVert = Abs(CDbl(txtOriginVert.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtOriginVert.SetFocus
End If
End Sub

Private Sub txtPageHeight_LostFocus()
On Error Resume Next
If IsNumeric(txtPageHeight.Text) Then
   PageHeight = Abs(CDbl(txtPageHeight.Text) * 1000)
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtPageHeight.SetFocus
End If
End Sub


Private Sub txtPageWidth_LostFocus()
On Error Resume Next
If IsNumeric(txtPageWidth.Text) Then
   PageWidth = Abs(CDbl(txtPageWidth.Text) * 1000)
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtPageWidth.SetFocus
End If
End Sub


Private Sub txtSpotHeight_LostFocus()
On Error Resume Next
If IsNumeric(txtSpotHeight.Text) Then
   SpotHeight = Abs(CDbl(txtSpotHeight.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtSpotHeight.SetFocus
End If
End Sub


Private Sub txtSpotWidth_LostFocus()
On Error Resume Next
If IsNumeric(txtSpotWidth.Text) Then
   SpotWidth = Abs(CDbl(txtSpotWidth.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtSpotWidth.SetFocus
End If
End Sub


Private Sub txtStickWidth_LostFocus()
On Error Resume Next
If IsNumeric(txtStickWidth.Text) Then
   StickWidth = Abs(CDbl(txtStickWidth.Text))
Else
   MsgBox Msg_GT_0, vbOKOnly, glFGTU
   txtStickWidth.SetFocus
End If
End Sub


Private Sub txtVertMargin_LostFocus()
On Error Resume Next
If IsNumeric(txtVertMargin.Text) Then
   VertMargin = Abs(CDbl(txtVertMargin.Text))
Else
   MsgBox Msg_GE_0, vbOKOnly, glFGTU
   txtVertMargin.SetFocus
End If
End Sub

Private Sub FillSizeTypeCombo()
cmbSizeType.AddItem "None"
cmbSizeType.AddItem "Linear"
cmbSizeType.AddItem "Log."
End Sub

Public Function IsInScope(ValNET As Double, ValMW As Double) As Boolean
'----------------------------------------------------------------------
'returns True if ValNET, ValMW fit within current scope
'----------------------------------------------------------------------
If ValNET >= mMinET Then
   If ValNET <= mMaxET Then
      If ValMW >= mMinMW Then
         If ValMW <= mmaxmw Then IsInScope = True
      End If
   End If
End If
End Function


Public Function GetSizeL(ValAbu As Double) As Single
'----------------------------------------------------------------------
'returns appropriate size scaled between mSpotHeight and mMaxSpotHeight
'----------------------------------------------------------------------
Select Case mSizeScaleType
Case stLinear
     GetSizeL = mSpotHeightL + ((ValAbu - mMinAbu) / mAdjAbuRange) * (mMaxSpotHeightL - mSpotHeightL)
Case stLog
     GetSizeL = mSpotHeightL + ((Log(ValAbu) / Log(10) - Log(mMinAbu) / Log(10)) / mAdjAbuRange) * (mMaxSpotHeightL - mSpotHeightL)
Case Else
     GetSizeL = mSpotHeightL
End Select
End Function


Private Function SetAdjAbuRange() As Boolean
On Error Resume Next
Select Case mSizeScaleType
Case stLinear
     mAdjAbuRange = mMaxAbu - mMinAbu
Case stLog
     mAdjAbuRange = Log(mMaxAbu) / Log(10) - Log(mMinAbu) / Log(10)
Case Else
     mAdjAbuRange = -1
End Select
End Function


'text and comments procedures and settings-----------------------------------------
Private Sub FillTextNameCombo()
cmbTextName.AddItem TXT_PNNL
cmbTextName.AddItem TXT_LEGEND
cmbTextName.AddItem TXT_DATE
cmbTextName.AddItem TXT_DATETIME
cmbTextName.AddItem TXT_SCAN_METRIC
End Sub


Private Sub SetTextProperties(TextName As String)
'-----------------------------------------------------------------------------
'retrieves text properties for the currently selected text box
'-----------------------------------------------------------------------------
Dim tb As New TextBoxGraph
On Error Resume Next
Set tb = TextBoxes(TextName)
If Not tb Is Nothing Then
   txtTextText.Text = tb.GetText
   txtTextPosX1.Text = CLng(tb.lX1 / HI_MM)     'set text boxes in milimeters
   txtTextPosY1.Text = CLng(tb.lY1 / HI_MM)
   txtTextWidth.Text = CLng((tb.lX2 - tb.lX1) / HI_MM)
   txtTextHeight.Text = CLng((tb.lY2 - tb.lY1) / HI_MM)
   txtTextFontHeight.Text = CLng(tb.FontHeightL / HI_MM)
   If tb.FontItalic Then
      chkTextItalic.value = vbChecked
   Else
      chkTextItalic.value = vbUnchecked
   End If
   If tb.ShowBorder Then
      chkTextShowBorder.value = vbChecked
   Else
      chkTextShowBorder.value = vbUnchecked
   End If
   lblTextBackColor.BackColor = tb.BackColor
   lblTextForeColor.BackColor = tb.ForeColor
Else
End If
End Sub


Private Sub GetTextProperties(TextName As String)
'------------------------------------------------
'sets text properties (all at once)
'------------------------------------------------
Dim tb As New TextBoxGraph
Dim i As Long
Dim Ln() As String
On Error Resume Next
Set tb = TextBoxes(TextName)
If Not tb Is Nothing Then
   tb.ClearText
   Ln = Split(txtTextText.Text, vbCrLf)
   For i = 0 To UBound(Ln)
       tb.AddTextLine Ln(i)
   Next i
   tb.lX1 = CLng(txtTextPosX1.Text * HI_MM)           'user entry is in milimeters
   tb.lY1 = CLng(txtTextPosY1.Text * HI_MM)
   tb.lX2 = tb.lX1 + CLng(txtTextWidth.Text * HI_MM)
   tb.lY2 = tb.lY1 + CLng(txtTextHeight.Text * HI_MM)
   tb.FontItalic = (chkTextItalic.value = vbChecked)
   tb.ShowBorder = (chkTextShowBorder.value = vbChecked)
   tb.BackColor = lblTextBackColor.BackColor
   tb.ForeColor = lblTextForeColor.BackColor
   'properties that are not interfaced (hard coded)
   tb.lIndentX = 1 * HI_MM
   tb.lIndentY = 1 * HI_MM
   tb.FontHeightL = CLng(txtTextFontHeight.Text * HI_MM)
   tb.FontWidthL = CLng(tb.FontHeightL * mFontWidthRatio)
End If
End Sub


Private Function GetScanMetric(ByVal OlyInd As Long) As String
'-------------------------------------------------------------------------
'returns information about number of scans per milimeter on metric display
'-------------------------------------------------------------------------
Dim ScansPer_mm As Double
Dim FirstScan As Long, LastScan As Long, ScanRange As Long
On Error Resume Next
With Oly(OlyInd)
    GetScanRange .DisplayInd, FirstScan, LastScan, ScanRange
    ScansPer_mm = ScanRange / (mET_mm * (.maxNET - .minNET))
    If Err Then
       GetScanMetric = .Name & " - Scans/mm = Error"
    Else
       GetScanMetric = .Name & " - Scans/mm = " & Format$(ScansPer_mm, "0.0")
    End If
End With
End Function

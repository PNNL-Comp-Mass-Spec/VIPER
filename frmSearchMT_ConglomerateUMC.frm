VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Begin VB.Form frmSearchMT_ConglomerateUMC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Search MT tag DB - Single LC-MS Feature Mass"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   14025
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSTACMatchStats 
      Height          =   525
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   70
      Top             =   7100
      Width           =   7455
   End
   Begin VB.Frame fraSTACPlotOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "STAC Plot Options"
      Height          =   1455
      Left            =   11880
      TabIndex        =   64
      Top             =   1320
      Width           =   2055
      Begin VB.CheckBox chkPlotUPFilteredFDR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&UP Filtered FDR"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkSTACPlotY2Gridlines 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FDR Gridlines"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkSTACPlotY1Gridlines 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Matches Gridlines"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkSTACPlotXGridlines 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vertical Gridlines"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdZoomOutSTACPlot 
      Caption         =   "&Zoom Out"
      Height          =   375
      Left            =   11880
      TabIndex        =   69
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopySTACSTats 
      Caption         =   "Copy Stats"
      Height          =   375
      Left            =   11880
      TabIndex        =   63
      Top             =   360
      Width           =   1335
   End
   Begin CWUIControlsLib.CWGraph ctlSTACStats 
      Height          =   3255
      Left            =   7800
      TabIndex        =   62
      Top             =   3480
      Width           =   4095
      _Version        =   393218
      _ExtentX        =   7223
      _ExtentY        =   5741
      _StockProps     =   71
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   393218
      Graph_0         =   1
      ClassName_1     =   "CCWGraphFrame"
      opts_1          =   62
      C[0]_1          =   16777215
      C[1]_1          =   16777215
      Event_1         =   2
      ClassName_2     =   "CCWGFPlotEvent"
      Owner_2         =   1
      Plots_1         =   3
      ClassName_3     =   "CCWDataPlots"
      Array_3         =   3
      Editor_3        =   4
      ClassName_4     =   "CCWGFPlotArrayEditor"
      Owner_4         =   1
      Array[0]_3      =   5
      ClassName_5     =   "CCWDataPlot"
      opts_5          =   4194367
      Name_5          =   "Matches"
      C[0]_5          =   16711680
      C[1]_5          =   0
      C[2]_5          =   16711680
      C[3]_5          =   16776960
      Event_5         =   2
      X_5             =   6
      ClassName_6     =   "CCWAxis"
      opts_6          =   575
      Name_6          =   "STAC Threshold"
      Orientation_6   =   2944
      format_6        =   7
      ClassName_7     =   "CCWFormat"
      Scale_6         =   8
      ClassName_8     =   "CCWScale"
      opts_8          =   90112
      rMin_8          =   43
      rMax_8          =   210
      dMax_8          =   1
      discInterval_8  =   1
      Radial_6        =   0
      Enum_6          =   9
      ClassName_9     =   "CCWEnum"
      Editor_9        =   10
      ClassName_10    =   "CCWEnumArrayEditor"
      Owner_10        =   6
      Font_6          =   0
      tickopts_6      =   2711
      major_6         =   0.5
      minor_6         =   0.25
      Caption_6       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   62
      C[0]_11         =   -2147483640
      Image_11        =   12
      ClassName_12    =   "CCWTextImage"
      font_12         =   0
      Animator_11     =   0
      Blinker_11      =   0
      Y_5             =   13
      ClassName_13    =   "CCWAxis"
      opts_13         =   1599
      Name_13         =   "Matches"
      C[3]_13         =   8421504
      Orientation_13  =   2067
      format_13       =   14
      ClassName_14    =   "CCWFormat"
      Format_14       =   "."
      Scale_13        =   15
      ClassName_15    =   "CCWScale"
      opts_15         =   122880
      rMin_15         =   28
      rMax_15         =   187
      dMax_15         =   10
      discInterval_15 =   1
      Radial_13       =   0
      Enum_13         =   16
      ClassName_16    =   "CCWEnum"
      Editor_16       =   17
      ClassName_17    =   "CCWEnumArrayEditor"
      Owner_17        =   13
      Font_13         =   0
      tickopts_13     =   2743
      major_13        =   2
      minor_13        =   1
      Caption_13      =   18
      ClassName_18    =   "CCWDrawObj"
      opts_18         =   62
      C[0]_18         =   16711680
      Image_18        =   19
      ClassName_19    =   "CCWTextImage"
      szText_19       =   "Matches"
      font_19         =   0
      Animator_18     =   0
      Blinker_18      =   0
      PointStyle_5    =   2
      LineStyle_5     =   1
      LineWidth_5     =   2
      BasePlot_5      =   0
      DefaultXInc_5   =   1
      DefaultPlotPerRow_5=   -1  'True
      Array[1]_3      =   20
      ClassName_20    =   "CCWDataPlot"
      opts_20         =   4194367
      Name_20         =   "UpFilteredMatches"
      C[0]_20         =   41984
      C[1]_20         =   0
      C[2]_20         =   16711680
      C[3]_20         =   16776960
      Event_20        =   2
      X_20            =   6
      Y_20            =   13
      PointStyle_20   =   16
      LineStyle_20    =   1
      LineWidth_20    =   2
      BasePlot_20     =   0
      DefaultXInc_20  =   1
      DefaultPlotPerRow_20=   -1  'True
      Array[2]_3      =   21
      ClassName_21    =   "CCWDataPlot"
      opts_21         =   4194367
      Name_21         =   "FDR"
      C[0]_21         =   255
      C[1]_21         =   0
      C[2]_21         =   16711680
      C[3]_21         =   16776960
      Event_21        =   2
      X_21            =   6
      Y_21            =   22
      ClassName_22    =   "CCWAxis"
      opts_22         =   575
      Name_22         =   "FDR"
      Orientation_22  =   2172
      format_22       =   23
      ClassName_23    =   "CCWFormat"
      Format_23       =   "%*100"
      Scale_22        =   24
      ClassName_24    =   "CCWScale"
      opts_24         =   90112
      rMin_24         =   28
      rMax_24         =   187
      dMax_24         =   1
      discInterval_24 =   1
      Radial_22       =   0
      Enum_22         =   25
      ClassName_25    =   "CCWEnum"
      Editor_25       =   26
      ClassName_26    =   "CCWEnumArrayEditor"
      Owner_26        =   22
      Font_22         =   0
      tickopts_22     =   2711
      major_22        =   0.2
      minor_22        =   0.1
      Caption_22      =   27
      ClassName_27    =   "CCWDrawObj"
      opts_27         =   62
      C[0]_27         =   255
      Image_27        =   28
      ClassName_28    =   "CCWTextImage"
      szText_28       =   "FDR"
      style_28        =   74772712
      font_28         =   0
      Animator_27     =   0
      Blinker_27      =   0
      PointStyle_21   =   10
      LineStyle_21    =   1
      LineWidth_21    =   2
      BasePlot_21     =   0
      DefaultXInc_21  =   1
      DefaultPlotPerRow_21=   -1  'True
      Axes_1          =   29
      ClassName_29    =   "CCWAxes"
      Array_29        =   3
      Editor_29       =   30
      ClassName_30    =   "CCWGFAxisArrayEditor"
      Owner_30        =   1
      Array[0]_29     =   6
      Array[1]_29     =   13
      Array[2]_29     =   22
      DefaultPlot_1   =   31
      ClassName_31    =   "CCWDataPlot"
      opts_31         =   4194367
      Name_31         =   "[Template]"
      C[0]_31         =   65280
      C[1]_31         =   255
      C[2]_31         =   16711680
      C[3]_31         =   16776960
      Event_31        =   2
      X_31            =   6
      Y_31            =   13
      PointStyle_31   =   16
      LineWidth_31    =   1
      BasePlot_31     =   0
      DefaultXInc_31  =   1
      DefaultPlotPerRow_31=   -1  'True
      Cursors_1       =   32
      ClassName_32    =   "CCWCursors"
      Editor_32       =   33
      ClassName_33    =   "CCWGFCursorArrayEditor"
      Owner_33        =   1
      TrackMode_1     =   10
      GraphBackground_1=   0
      GraphFrame_1    =   34
      ClassName_34    =   "CCWDrawObj"
      opts_34         =   62
      C[0]_34         =   16777215
      C[1]_34         =   16777215
      Image_34        =   35
      ClassName_35    =   "CCWPictImage"
      opts_35         =   1280
      Rows_35         =   1
      Cols_35         =   1
      F_35            =   16777215
      B_35            =   16777215
      ColorReplaceWith_35=   8421504
      ColorReplace_35 =   8421504
      Tolerance_35    =   2
      Animator_34     =   0
      Blinker_34      =   0
      PlotFrame_1     =   36
      ClassName_36    =   "CCWDrawObj"
      opts_36         =   62
      C[0]_36         =   16777215
      C[1]_36         =   16777215
      Image_36        =   37
      ClassName_37    =   "CCWPictImage"
      opts_37         =   1280
      Rows_37         =   1
      Cols_37         =   1
      Pict_37         =   1
      F_37            =   16777215
      B_37            =   16777215
      ColorReplaceWith_37=   8421504
      ColorReplace_37 =   8421504
      Tolerance_37    =   2
      Animator_36     =   0
      Blinker_36      =   0
      Caption_1       =   38
      ClassName_38    =   "CCWDrawObj"
      opts_38         =   62
      C[0]_38         =   -2147483640
      Image_38        =   39
      ClassName_39    =   "CCWTextImage"
      szText_39       =   "STAC Stats"
      font_39         =   0
      Animator_38     =   0
      Blinker_38      =   0
      DefaultXInc_1   =   1
      DefaultPlotPerRow_1=   -1  'True
      Bindings_1      =   40
      ClassName_40    =   "CCWBindingHolderArray"
      Editor_40       =   41
      ClassName_41    =   "CCWBindingHolderArrayEditor"
      Owner_41        =   1
      Annotations_1   =   42
      ClassName_42    =   "CCWAnnotations"
      Editor_42       =   43
      ClassName_43    =   "CCWAnnotationArrayEditor"
      Owner_43        =   1
      AnnotationTemplate_1=   44
      ClassName_44    =   "CCWAnnotation"
      opts_44         =   63
      Name_44         =   "[Template]"
      Plot_44         =   45
      ClassName_45    =   "CCWDataPlot"
      opts_45         =   4194367
      Name_45         =   "[Template]"
      C[0]_45         =   65280
      C[1]_45         =   255
      C[2]_45         =   16711680
      C[3]_45         =   16776960
      Event_45        =   2
      X_45            =   6
      Y_45            =   13
      LineStyle_45    =   1
      LineWidth_45    =   1
      BasePlot_45     =   0
      DefaultXInc_45  =   1
      DefaultPlotPerRow_45=   -1  'True
      Text_44         =   "[Template]"
      TextXPoint_44   =   6.7
      TextYPoint_44   =   6.7
      TextColor_44    =   16777215
      TextFont_44     =   46
      ClassName_46    =   "CCWFont"
      bFont_46        =   -1  'True
      BeginProperty Font_46 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShapeXPoints_44 =   47
      ClassName_47    =   "CDataBuffer"
      Type_47         =   5
      m_cDims;_47     =   1
      m_cElts_47      =   1
      Element[0]_47   =   3.3
      ShapeYPoints_44 =   48
      ClassName_48    =   "CDataBuffer"
      Type_48         =   5
      m_cDims;_48     =   1
      m_cElts_48      =   1
      Element[0]_48   =   3.3
      ShapeFillColor_44=   16777215
      ShapeLineColor_44=   16777215
      ShapeLineWidth_44=   1
      ShapeLineStyle_44=   1
      ShapePointStyle_44=   10
      ShapeImage_44   =   49
      ClassName_49    =   "CCWDrawObj"
      opts_49         =   62
      Image_49        =   50
      ClassName_50    =   "CCWPictImage"
      opts_50         =   1280
      Rows_50         =   1
      Cols_50         =   1
      Pict_50         =   7
      F_50            =   -2147483633
      B_50            =   -2147483633
      ColorReplaceWith_50=   8421504
      ColorReplace_50 =   8421504
      Tolerance_50    =   2
      Animator_49     =   0
      Blinker_49      =   0
      ArrowVisible_44 =   -1  'True
      ArrowColor_44   =   16777215
      ArrowWidth_44   =   1
      ArrowLineStyle_44=   1
      ArrowHeadStyle_44=   1
   End
   Begin MSComctlLib.ListView lvwSTACStats 
      Height          =   3015
      Left            =   7800
      TabIndex        =   61
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox chkSTACUsesPriorProbability 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use Prior Prob"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   360
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkUseSTAC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use STAC"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetDefaultsForToleranceRefinement 
      Caption         =   "Set to Tolerance Refinement Defaults"
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtDBSearchMinimumPeptideProphetProbability 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   17
      Text            =   "0"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdSetDefaults 
      Caption         =   "Set to Defaults"
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtUniqueMatchStats 
      Height          =   525
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   59
      Top             =   6525
      Width           =   7455
   End
   Begin VB.TextBox txtDBSearchMinimumHighDiscriminantScore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Text            =   "0"
      Top             =   2340
      Width           =   615
   End
   Begin VB.ComboBox cboAMTSearchResultsBehavior 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtDBSearchMinimumHighNormalizedScore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.ComboBox cboInternalStdSearchMode 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkUpdateGelDataWithSearchResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update data in current file with results of search"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearchAllUMCs 
      Caption         =   "Search All LC-MS Features"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemoveAMTMatchesFromUMCs 
      Caption         =   "Remove existing MT matches from features"
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Remove MT reference for current gel"
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame fraMods 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modifications"
      Height          =   1575
      Left            =   0
      TabIndex        =   35
      Top             =   4560
      Width           =   7575
      Begin VB.TextBox txtDecoySearchNETWobble 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         TabIndex        =   52
         Text            =   "0.1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox cboResidueToModify 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   43
         ToolTipText     =   $"frmSearchMT_ConglomerateUMC.frx":0000
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtResidueToModifyMass 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   45
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame fraOptionFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Index           =   47
         Left            =   6600
         TabIndex        =   53
         Top             =   360
         Width           =   795
         Begin VB.OptionButton optN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N14"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Value           =   -1  'True
            Width           =   700
         End
         Begin VB.OptionButton optN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N15"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   525
            Width           =   700
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "N Type:"
            Height          =   255
            Index           =   103
            Left            =   120
            TabIndex        =   54
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.Frame fraOptionFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   750
         Index           =   49
         Left            =   4480
         TabIndex        =   46
         Top             =   240
         Width           =   1920
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Decoy"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   50
            ToolTipText     =   $"frmSearchMT_ConglomerateUMC.frx":00D7
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dynamic"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   $"frmSearchMT_ConglomerateUMC.frx":0174
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fixed"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   48
            ToolTipText     =   "Changes the mass of all loaded AMTs, adding the value specified by the modification mass"
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Mod Type:"
            Height          =   255
            Index           =   100
            Left            =   120
            TabIndex        =   47
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.TextBox txtAlkylationMWCorrection 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   41
         Text            =   "57.0215"
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkAlkylation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alkylation"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         ToolTipText     =   "Check to add the alkylation mass correction below to all MT Tag masses (added to each cys residue)"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkICATHv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICAT d8"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkICATLt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ICAT d0"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkPEO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PEO"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDecoySearchNETWobble 
         BackStyle       =   0  'Transparent
         Caption         =   "Decoy NET Wobble"
         Height          =   375
         Left            =   4560
         TabIndex        =   51
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Residue to modify:"
         Height          =   255
         Left            =   3000
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mass (Da):"
         Height          =   255
         Left            =   3000
         TabIndex        =   44
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6480
         X2              =   6480
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4440
         X2              =   4440
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1320
         X2              =   1320
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alkylation mass:"
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame fraNET 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NET  Calculation"
      Height          =   1455
      Left            =   0
      TabIndex        =   20
      Top             =   3000
      Width           =   5175
      Begin VB.CheckBox chkDisableCustomNETs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disable NETs from Warping"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   600
         Width           =   2400
      End
      Begin VB.CheckBox chkUseUMCConglomerateNET 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Class NET of LC-MS Features"
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         ToolTipText     =   $"frmSearchMT_ConglomerateUMC.frx":0206
         Top             =   180
         Width           =   1965
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   28
         Text            =   "0.1"
         Top             =   1020
         Width           =   615
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pred. NET for MT Tags"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Use Predicted NET values for the MT tags"
         Top             =   480
         Width           =   2500
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Obs. NET for MT Tags"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Use NET calculated from all peptide observations for each MT tag"
         Top             =   240
         Value           =   -1  'True
         Width           =   2500
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   1020
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "NET T&olerance"
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label lblNETFormula 
         BackStyle       =   0  'Transparent
         Caption         =   "Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   810
         Width           =   2415
      End
   End
   Begin VB.Frame fraMWTolerance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1455
      Left            =   5280
      TabIndex        =   29
      Top             =   3000
      Width           =   2175
      Begin VB.ComboBox cboSearchRegionShape 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   32
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   160
         TabIndex        =   31
         Text            =   "10"
         Top             =   640
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerance"
         Height          =   255
         Left            =   160
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblSTACStats 
      BackStyle       =   0  'Transparent
      Caption         =   "STAC Search Stats"
      Height          =   255
      Left            =   7800
      TabIndex        =   60
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Peptide Prophet Probability"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2655
      Width           =   2865
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum PMT Discriminant Score"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2360
      Width           =   2505
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum PMT XCorr"
      Height          =   255
      Index           =   134
      Left            =   120
      TabIndex        =   12
      Top             =   2060
      Width           =   2145
   End
   Begin VB.Label lblInternalStdSearchMode 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Standard Search Mode:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1725
      Width           =   2415
   End
   Begin VB.Label lblUMCMassMode 
      BackStyle       =   0  'Transparent
      Caption         =   "LC-MS Feature Mass = ??"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblETType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Generic NET"
      Height          =   255
      Left            =   5280
      TabIndex        =   58
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   6240
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuF 
      Caption         =   "&Function"
      Begin VB.Menu mnuFSearchAll 
         Caption         =   "Search &All LC-MS Features (UMCs)"
      End
      Begin VB.Menu mnuFSearchPaired 
         Caption         =   "Search Paired LC-MS Features (skips excluded pairs)"
      End
      Begin VB.Menu mnuFSearchPairedPlusNonPaired 
         Caption         =   "Search Light Members of Pairs &Plus Non-paired LC-MS Features (skips excluded)"
      End
      Begin VB.Menu mnuFSearchNonPaired 
         Caption         =   "Search &Non-paired LC-MS Features"
      End
      Begin VB.Menu mnuFSearchN14LabeledFeatures 
         Caption         =   "Search N14-labeled Features"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExcludeAmbiguous 
         Caption         =   "Exclude Ambiguous Pairs (all pairs)"
      End
      Begin VB.Menu mnuFExcludeAmbiguousHitsOnly 
         Caption         =   "Exclude Ambiguous Pairs (only those with hits)"
      End
      Begin VB.Menu mnuFResetExclusionFlags 
         Caption         =   "Reset Exclusion Flags for All Pairs"
      End
      Begin VB.Menu mnuFDeleteExcludedPairs 
         Caption         =   "Delete Excluded Pairs"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFReportByUMC 
         Caption         =   "Report Results by &UMCs (LC-MS Features)..."
      End
      Begin VB.Menu mnuFReportByIon 
         Caption         =   "Report Results by &Ions..."
      End
      Begin VB.Menu mnuFReportIncludeORFs 
         Caption         =   "Include Proteins (ORFs) in Report"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFSepExportToDatabase 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExportResultsToDBbyUMC 
         Caption         =   "Export Results to MT Tag DB (by LC-MS Feature)"
      End
      Begin VB.Menu mnuFExportDetailedMemberInformation 
         Caption         =   "Export detailed member information for each LC-MS Feature"
      End
      Begin VB.Menu mnuFSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFMassCalAndToleranceRefinement 
         Caption         =   "&Mass Calibration and Tolerance Refinement"
      End
      Begin VB.Menu mnuFSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopySTACStats 
         Caption         =   "&Copy STAC Stats"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditCopySTACPlot 
         Caption         =   "&Copy STAC Plot to Clipboard"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEditSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSaveSTACPlot 
         Caption         =   "Save STAC Plot"
         Begin VB.Menu mnuEditSaveSTACPlotAsEMF 
            Caption         =   "as &EMF"
         End
         Begin VB.Menu mnuEditSaveSTACPlotAsPNG 
            Caption         =   "as &PNG"
         End
      End
      Begin VB.Menu mnuEditSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSetToDefaults 
         Caption         =   "Set to &Defaults"
      End
   End
   Begin VB.Menu mnuMT 
      Caption         =   "&MT Tags"
      Begin VB.Menu mnuMTLoadMT 
         Caption         =   "Load MT Tag DB"
      End
      Begin VB.Menu mnuMTLoadLegacy 
         Caption         =   "Load Legacy MT DB"
      End
      Begin VB.Menu mnuMTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTStatus 
         Caption         =   "MT Tags Status"
      End
   End
   Begin VB.Menu mnuETHeader 
      Caption         =   "&Elution Time"
      Begin VB.Menu mnuET 
         Caption         =   "&Generic NET"
         Index           =   0
      End
      Begin VB.Menu mnuET 
         Caption         =   "&TIC Fit NET"
         Index           =   1
      End
      Begin VB.Menu mnuET 
         Caption         =   "G&ANET"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSearchMT_ConglomerateUMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is UMC identification - pairs are here just to distinguish
'which UMC to include in search
'---------------------------------------------------------------
'Elution is not corrected for N15 versions of peptides (???)
'When looking for N14; LC-MS Features that are heavy members of pairs only
'are not search; neither are LC-MS Features light only pair members when
'N15 search is performed
'---------------------------------------------------------------
'created: 10/10/2002 nt
'last modified: 10/17/2002 nt
'---------------------------------------------------------------
Option Explicit

Private Const NET_PRECISION As Integer = 5

Const MOD_TKN_NONE As String = "none"
Const MOD_TKN_PEO As String = "PEO"
Const MOD_TKN_ICAT_D0 As String = "ICAT_D0"
Const MOD_TKN_ICAT_D8 As String = "ICAT_D8"
Const MOD_TKN_ALK As String = "ALK"
Const MOD_TKN_N14 As String = "N14"
Const MOD_TKN_N15 As String = "N15"
Const MOD_TKN_RES_MOD As String = "RES_MOD"
Const MOD_TKN_MT_MOD As String = "MT_MOD"

Const SEARCH_N14 As Integer = 0
Const SEARCH_N15 As Integer = 1

Const MODS_FIXED As Integer = 0
Const MODS_DYNAMIC As Integer = 1
Const MODS_DECOY As Integer = 2

Const SEARCH_ALL As Integer = 0
Const SEARCH_PAIRED As Integer = 1
Const SEARCH_NON_PAIRED As Integer = 2
Const SEARCH_PAIRED_PLUS_NON_PAIRED As Integer = 3

'if called with any positive number add that many points
Const MNG_RESET As Integer = 0
Const MNG_ERASE As Integer = -1
Const MNG_TRIM As Integer = -2
Const MNG_ADD_START_SIZE As Integer = -3

Const MNG_START_SIZE As Long = 500

Const NET_WOBBLE_SEED As Long = 1000

Private Const STAC_APP_NAME As String = "STAC.exe"
Private Const TEMP_FILE_FLAG As String = "FILE"
Private Const TEMP_FOLDER_FLAG As String = "FOLDER"

'in this case CallerID is a public property
Public CallerID As Long

Private bLoading As Boolean

Private OldSearchFlag As Long

'for faster search mass array will be sorted; therefore all other arrays
'has to be addressed indirectly (mMTNET(mMTInd(i))
Private mMTCnt                  'count of masses to search
Private mMTInd() As Long        'index(unique key)              ' 0-based array
Private mMTOrInd() As Long      'index of original MT tag (in AMT array)
Private mMTMWN14() As Double    'mass to look for N14
Private mMTMWN15() As Double    'mass to look for N15
Private mMTNET() As Double      'NET value
Private mMTMods() As String     'modification description

Private MWFastSearch As MWUtil

Private mInternalStdIndexPointers() As Long             ' Pointer to entry in UMCInternalStandards.InternalStandards()
Private InternalStdFastSearch As MWUtil

Private AlkMWCorrection As Double
Private N14N15 As Long                  ' SEARCH_N14 or SEARCH_N15
Private SearchType As Long              ' SEARCH_ALL, SEARCH_PAIRED, SEARCH_NON_PAIRED, or SEARCH_PAIRED_PLUS_NON_PAIRED
Private mMTListContains16O18OMods As Boolean            ' Set to True when the user enters a full-peptide modification of 4.0085 Da (+/- 0.01 Da)
Private mSearchRegionShape As srsSearchRegionShapeConstants

Private LastSearchTypeN14N15 As Long
Private NTypeStr As String

'following arrays are parallel to the LC-MS Features in GelUMC()
Private ClsCnt As Long              'this is not actually neccessary except
Private ClsStat() As Double         'to create nice reports; necessary to use this since we report the Min/Max Charge stats and Average Fit stats
Private eClsPaired() As umcpUMCPairMembershipConstants      ' Keeps track of whether UMC is member of 1 or more pairs

                                
'mUMCMatchStats contains all possible identifications for all LC-MS Features with scores
'as count of each identification hits within the UMC
Private mMatchStatsCount As Long                                'count of UMC-ID matches
Private mUMCMatchStats() As udtUMCMassTagMatchStats             ' 0-based array
Private mSearchUsedCustomNETs As Boolean

' The following hold match stats for each individual UMC
Private mCurrIDCnt As Long
Private mCurrIDMatches() As udtUMCMassTagRawMatches         ' 0-based array

' The following is populated after the search finishes
' It tracks the number of UMCs matched, plus the unique number of AMTs matched at different FDR levels
Private mSearchSummaryStats As udtSearchSummaryStatsType

Private mSTACTempFolderPath As String
Private mSTACSessionID As String
Private mTempFilesToDelete As Dictionary
    
Private mSTACAMTFilePath As String
Private mSTACUMCFilePath As String
Private mMaxPlottedFDR As Double

'Expression Evaluator variables for elution time calculation
Private MyExprEva As ExprEvaluator
Private VarVals() As Long
Private MinFN As Long
Private MaxFN As Long

Private ExpAnalysisSPName As String             ' Stored procedure AddMatchMaking
''Private ExpPeakSPName As String               ' Stored procedure AddFTICRPeak; Unused variable
Private ExpUmcSPName As String                  ' Stored procedure AddFTICRUmc
Private ExpUMCMemberSPName As String            ' Stored procedure AddFTICRUmcMember
Private ExpUmcMatchSPName As String             ' Stored procedure AddFTICRUmcMatch
Private ExpUmcInternalStdMatchSPName As String  ' Stored procedure AddFTICRUmcInternalStdMatch
Private ExpUMCCSStats As String                 ' Stored procedure AddFTICRUmcCSStats
Private ExpStoreSTACStats As String             ' Stored procedure AddMatchMakingFDR
Private ExpQuantitationDescription As String    ' Stored procedure AddQuantitationDescription

Private mUMCCountSkippedSinceRefPresent As Long
Private mUsingDefaultGANET As Boolean
Private eInternalStdSearchMode As issmInternalStandardSearchModeConstants
Private mMTMinimumHighNormalizedScore As Single
Private mMTMinimumHighDiscriminantScore As Single
Private mMTMinimumPeptideProphetProbability As Single

Private mMDTypeSaved As Long

Private mKeyPressAbortProcess As Integer

Private objMTDBNameLookupClass As mtdbMTNames

Private Enum eSearchModeConstants
    eSearchModeAll = 0
    eSearchModePaired = 1
    eSearchModePairedPlusUnpaired = 2
    eSearchModeNonPaired = 3
End Enum
'

Public Property Get PlotUPFilteredFDR() As Boolean
    PlotUPFilteredFDR = cChkBox(chkPlotUPFilteredFDR)
End Property
Public Property Let PlotUPFilteredFDR(Value As Boolean)
    SetCheckBox chkPlotUPFilteredFDR, Value
End Property

Public Property Get SearchRegionShape() As srsSearchRegionShapeConstants
    SearchRegionShape = mSearchRegionShape
End Property
Public Property Let SearchRegionShape(Value As srsSearchRegionShapeConstants)
    cboSearchRegionShape.ListIndex = Value
    mSearchRegionShape = Value
End Property

Public Property Get UpdateGelDataWithSearchResults() As Boolean
    UpdateGelDataWithSearchResults = cChkBox(chkUpdateGelDataWithSearchResults)
End Property
Public Property Let UpdateGelDataWithSearchResults(ByVal Value As Boolean)
    SetCheckBox chkUpdateGelDataWithSearchResults, Value
End Property

Public Property Get STACUsesPriorProbability() As Boolean
    STACUsesPriorProbability = cChkBox(chkSTACUsesPriorProbability)
End Property
Public Property Let STACUsesPriorProbability(ByVal Value As Boolean)
    If cChkBox(chkSTACUsesPriorProbability) <> Value Then
        SetCheckBox chkSTACUsesPriorProbability, Value
    End If
    glbPreferencesExpanded.STACUsesPriorProbability = Value
End Property

Public Property Get UseSTAC() As Boolean
    UseSTAC = cChkBox(chkUseSTAC)
End Property
Public Property Let UseSTAC(ByVal Value As Boolean)
    If cChkBox(chkUseSTAC) <> Value Then
        SetCheckBox chkUseSTAC, Value
    End If
    glbPreferencesExpanded.UseSTAC = Value
End Property

Private Function AddCurrIDsToAllIDs(ClsInd As Long) As Boolean
    '---------------------------------------------------------------------------
    'returns True if successful; adds current identifications to list of all IDs
    '---------------------------------------------------------------------------
    Dim lngIndex As Long, lngTargetIndex As Long
    Dim lngAMTHitCount As Long
    
    On Error GoTo err_AddCurrIDsToAllIDs
    mMatchStatsCount = mMatchStatsCount + mCurrIDCnt
    
    If mMatchStatsCount > UBound(mUMCMatchStats) Then
        ' Reserve more room in mUMCMatchStats
        ReDim Preserve mUMCMatchStats(UBound(mUMCMatchStats) * 2)
    End If
    
    ' Count the number of non Internal Standard matches
    lngAMTHitCount = 0
    For lngIndex = 0 To mCurrIDCnt - 1
        If Not mCurrIDMatches(lngIndex).IDIsInternalStd Then
            lngAMTHitCount = lngAMTHitCount + 1
        End If
    Next lngIndex
    
    For lngIndex = 0 To mCurrIDCnt - 1
        lngTargetIndex = (mMatchStatsCount - mCurrIDCnt) + lngIndex
        With mUMCMatchStats(lngTargetIndex)
            .UMCIndex = ClsInd
            .IDIndex = mCurrIDMatches(lngIndex).IDInd
            .MemberHitCount = mCurrIDMatches(lngIndex).MatchingMemberCount
            .StacOrSLiC = mCurrIDMatches(lngIndex).StacOrSLiC
            .DelScore = mCurrIDMatches(lngIndex).DelScore
            .UniquenessProbability = mCurrIDMatches(lngIndex).UniquenessProbability
            .FDRThreshold = 1
            .IDIsInternalStd = mCurrIDMatches(lngIndex).IDIsInternalStd
            .MultiAMTHitCount = lngAMTHitCount
        End With
    Next lngIndex
    
    AddCurrIDsToAllIDs = True
    
    Exit Function
    
err_AddCurrIDsToAllIDs:
    ' Error updating UMC with new matches
    Debug.Assert False
    
End Function

Public Sub AutoSizeForm(Optional ByVal blnSizeForSTACPlotSave As Boolean = False)
    
    If blnSizeForSTACPlotSave Then
        Me.width = 15000
        Me.Height = 10000
    Else
        If Me.UseSTAC And STACStatsCount > 0 Then
            If Me.width < 14150 Then
                Me.width = 14150
            End If
            
            If Me.Height < 9000 Then
                Me.Height = 9000
            End If
        Else
            Me.width = 7800
            Me.Height = 8000
            txtSTACMatchStats.Visible = False
        End If
    End If
    
End Sub

Private Sub CheckNETEquationStatus()
    If RobustNETValuesEnabled() Then
        mUsingDefaultGANET = True
    Else
        If Not GelAnalysis(CallerID) Is Nothing Then
            If txtNETFormula.Text = ConstructNETFormula(GelAnalysis(CallerID).GANET_Slope, GelAnalysis(CallerID).GANET_Intercept) _
               And InStr(UCase(txtNETFormula), "MINFN") = 0 Then
                mUsingDefaultGANET = True
            Else
                mUsingDefaultGANET = False
            End If
        Else
            mUsingDefaultGANET = False
        End If
    End If
End Sub

Private Function CheckVsMinimum(ByVal lngValue As Long, Optional ByVal lngMinimum As Long = 0) As Long
    If lngValue < lngMinimum Then lngValue = lngMinimum
    CheckVsMinimum = lngValue
End Function

Private Function ComputePeptideLevelSTACFDR() As Boolean

    Dim lngIndex As Long
    
    Dim lngPointerArray() As Long
    Dim dblSTACScores() As Double
    Dim dblFDR() As Double
    
    Dim dblRunningSum As Double
    Dim lngRunningCount As Long
    
    Dim objQSDbl As QSDouble
    Dim objQSLong As QSLongWithDouble
    
    Dim blnSuccess As Boolean
    
On Error GoTo ComputePeptideLevelSTACFDRErrorHandler
    
    If mMatchStatsCount > 0 Then
        
        ReDim lngPointerArray(mMatchStatsCount - 1)
        ReDim dblSTACScores(mMatchStatsCount - 1)
        ReDim dblFDR(mMatchStatsCount - 1)
        
        ' Populate two parallel arrays
        
        For lngIndex = 0 To mMatchStatsCount - 1
            lngPointerArray(lngIndex) = lngIndex
            dblSTACScores(lngIndex) = mUMCMatchStats(lngIndex).StacOrSLiC
        Next lngIndex
        
        ' Sort dblSTACScores() and sort lngPointerArray() in parallel
        
        Set objQSDbl = New QSDouble
        blnSuccess = objQSDbl.QSDesc(dblSTACScores, lngPointerArray)
                
        ' Now step through the data and compute the FDR values
        ' FDR = (CountToThisPoint - RunningSTACSum) / CountToThisPoint
        
        dblRunningSum = 0
        For lngIndex = 0 To mMatchStatsCount - 1
            dblRunningSum = dblRunningSum + dblSTACScores(lngIndex)
            lngRunningCount = lngIndex + 1
            dblFDR(lngIndex) = (lngRunningCount - dblRunningSum) / lngRunningCount
        Next lngIndex
        
        ' Now step through the data and assign the same FDR for contiguous data with the same STAC scores
        For lngIndex = mMatchStatsCount - 1 To 1 Step -1
            If dblSTACScores(lngIndex) = dblSTACScores(lngIndex - 1) Then
                dblFDR(lngIndex - 1) = dblFDR(lngIndex)
            End If
        Next lngIndex
        
        ' Now re-sort by lngPointerArray()
        
        Set objQSLong = New QSLongWithDouble
        blnSuccess = objQSLong.QSAsc(lngPointerArray, dblFDR)
        
        ' Finally, store the FDR values
        ' We're also counting the number of features that pass each threshold
        For lngIndex = 0 To mMatchStatsCount - 1
            mUMCMatchStats(lngIndex).FDRThreshold = dblFDR(lngIndex)
        Next lngIndex
        
        ' Count the number of features passing each FDR threshold
        For lngIndex = 0 To mMatchStatsCount - 1
        Next lngIndex
    End If
    
    ComputePeptideLevelSTACFDR = True
    Exit Function

ComputePeptideLevelSTACFDRErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "ComputePeptideLevelSTACFDR"
    
End Function

Private Function ConvertScanToNET(lngScanNumber As Long) As Double

    If mUsingDefaultGANET Then
        ConvertScanToNET = ScanToGANET(CallerID, lngScanNumber)
    Else
        ConvertScanToNET = Elution(lngScanNumber, MinFN, MaxFN)
    End If

End Function

Private Function CountMassTagsInUMCMatchStats() As Long
    ' Returns the number of items in mUMCMatchStats() with .IDIsInternalStd = False
    
    Dim lngMassTagHitCount As Long
    Dim lngIndex As Long
    
    lngMassTagHitCount = 0
    For lngIndex = 0 To mMatchStatsCount - 1
        If Not mUMCMatchStats(lngIndex).IDIsInternalStd Then lngMassTagHitCount = lngMassTagHitCount + 1
    Next lngIndex
    
    CountMassTagsInUMCMatchStats = lngMassTagHitCount

End Function

Public Function DeleteExcludedPairsWrapper()
    Dim strMessage As String
    strMessage = DeleteExcludedPairs(CallerID)
    AddToAnalysisHistory CallerID, strMessage
    
    UpdateUMCsPairingStatusNow
End Function

Private Sub DeleteTempFiles()

    Dim lngIndex As Long
    Dim objKeys() As Variant
    
    Dim fso As New FileSystemObject
    
    On Error GoTo DeleteTempFilesErrorHandler
    
    If mTempFilesToDelete.Count > 0 Then
        objKeys = mTempFilesToDelete.Keys
        For lngIndex = 0 To mTempFilesToDelete.Count - 1
            If CStr(mTempFilesToDelete.Item(objKeys(lngIndex))) = TEMP_FOLDER_FLAG Then
                ' This is a folder
                If fso.FolderExists(objKeys(lngIndex)) Then
                    fso.DeleteFolder objKeys(lngIndex), True
                End If
            Else
                ' This is a file
                If fso.FileExists(objKeys(lngIndex)) Then
                    fso.DeleteFile objKeys(lngIndex)
                End If
            End If
        Next lngIndex
    End If
    
    Exit Sub

DeleteTempFilesErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub DestroyIDStructures()
    On Error Resume Next
    mMatchStatsCount = 0
    Erase mUMCMatchStats
    Call ManageCurrID(MNG_ERASE)
End Sub

Private Sub DestroySearchStructures()
    On Error Resume Next
    mMTCnt = 0
    Erase mMTInd
    Erase mMTOrInd
    Erase mMTMWN14
    Erase mMTMWN15
    Erase mMTNET
    Erase mMTMods
    Erase mInternalStdIndexPointers
    Set MWFastSearch = Nothing
    Set InternalStdFastSearch = Nothing
End Sub

Private Sub DisplayCurrentSearchTolerances()
    With samtDef
        txtMWTol.Text = .MWTol
    
        SetTolType .TolType
        
        'NETTol is used both for NET and RT
        If .NETTol >= 0 Then
           txtNETTol.Text = .NETTol
           txtNETTol_Validate False
        Else
           txtNETTol.Text = ""
        End If
    End With
End Sub

Private Function DisplayHitSummary(strSearchScope As String) As String

    Dim strMessage As String
    Dim strStats As String
    Dim strSTACStats As String
    
    Dim strSearchItems As String
    Dim strModMassDescription As String
    
    Dim sngUMCMatchPercentage As Single
    
    strMessage = "Hits: " & LongToStringWithCommas(mMatchStatsCount)
    Select Case eInternalStdSearchMode
    Case issmFindWithMassTags
        strSearchItems = "MT tags and/or Internal Stds"
    Case issmFindOnlyInternalStandards
        strSearchItems = "Internal Stds"
    Case Else
        ' Includes issmFindOnlyMassTags
        strSearchItems = "MT tags"
    End Select
    strMessage = strMessage & " " & strSearchItems
    
    ' Determine the unique number of LC-MS Features with matches, the unique MT tag count, and the unique Internal Standard Count
    
    If mUMCCountSkippedSinceRefPresent > 0 Then
        strMessage = strMessage & " (" & Trim(mUMCCountSkippedSinceRefPresent) & " LC-MS Features skipped)"
    End If
    
    UpdateStatus strMessage
    
    GelSearchDef(CallerID).AMTSearchOnUMCs = samtDef
    
    AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched " & strSearchScope & " LC-MS Features for " & strSearchItems & " (searched by LC-MS Feature conglomerate mass, " & lblUMCMassMode & "; however, all members of a LC-MS Feature are assigned all matches found for the UMC)", mMatchStatsCount, mMTMinimumHighNormalizedScore, mMTMinimumHighDiscriminantScore, mMTMinimumPeptideProphetProbability, samtDef, True, mSearchUsedCustomNETs)
    
    strModMassDescription = ConstructMassTagModMassDescription(GelSearchDef(CallerID).AMTSearchMassMods)
    If Len(strModMassDescription) > 0 Then
        AddToAnalysisHistory CallerID, strModMassDescription
    End If

    ' Re-compute the match stats
    GenerateUniqueMatchStats
    
    If GelUMC(CallerID).UMCCnt > 0 Then
        sngUMCMatchPercentage = mSearchSummaryStats.UMCCountWithHits / CSng(GelUMC(CallerID).UMCCnt) * 100#
    Else
        sngUMCMatchPercentage = 0
    End If
    
    strStats = "LC-MS Features with match = " & LongToStringWithCommas(mSearchSummaryStats.UMCCountWithHits) & " (" & Trim(Round(sngUMCMatchPercentage, 0)) & "%)"
    If eInternalStdSearchMode <> issmFindOnlyInternalStandards Then
        strStats = strStats & "; Unique MT tags matched = " & LongToStringWithCommas(mSearchSummaryStats.UniqueMTCount) & " / " & LongToStringWithCommas(mMTCnt)
        If mMTCnt > AMTCnt Then
            strStats = strStats & " (" & LongToStringWithCommas(AMTCnt) & " source MT tags)"
        End If
    End If
    
    If eInternalStdSearchMode <> issmFindOnlyMassTags Then
        strStats = strStats & "; Unique Int Stds = " & LongToStringWithCommas(mSearchSummaryStats.UniqueIntStdCount) & " / " & LongToStringWithCommas(UMCInternalStandards.Count)
    End If
    
    strSTACStats = "Unique AMT count, 1% FDR: " & LongToStringWithCommas(mSearchSummaryStats.UniqueMTCount1PctFDR) & ";  " & _
                   "5% FDR: " & LongToStringWithCommas(mSearchSummaryStats.UniqueMTCount5PctFDR) & ";  " & _
                   "10% FDR: " & LongToStringWithCommas(mSearchSummaryStats.UniqueMTCount10PctFDR) & ";  " & _
                   "25% FDR: " & LongToStringWithCommas(mSearchSummaryStats.UniqueMTCount25PctFDR)
    
    txtUniqueMatchStats.Text = strStats
    
    If Not txtSTACMatchStats.Visible Then
        If GelData(CallerID).MostRecentSearchUsedSTAC Then
            txtSTACMatchStats.Visible = True
        End If
    End If
    
    txtSTACMatchStats.Text = strSTACStats
    
    AddToAnalysisHistory CallerID, "Match stats: " & strStats

    DisplayHitSummary = strMessage

End Function

Private Function Elution(FN As Long, MinFN As Long, MaxFN As Long) As Double
'---------------------------------------------------
'this function does not care are we using NET or RT
'---------------------------------------------------
VarVals(1) = FN
VarVals(2) = MinFN
VarVals(3) = MaxFN
Elution = MyExprEva.ExprVal(VarVals())
End Function

Private Sub EnableDisableControls()
    If optDBSearchModType(2).Value = True Then
        txtDecoySearchNETWobble.Enabled = True
    Else
        txtDecoySearchNETWobble.Enabled = False
    End If
    
    If Me.UseSTAC Then
        chkSTACUsesPriorProbability.Enabled = True
        txtSTACMatchStats.Visible = True
    Else
        chkSTACUsesPriorProbability.Enabled = False
        txtSTACMatchStats.Visible = False
    End If
    
    AutoSizeForm
End Sub

Private Sub EnableDisableNETFormulaControls()
    Dim i As Integer
    
    txtNETFormula.Enabled = Not RobustNETValuesEnabled()
    lblNETFormula.Enabled = txtNETFormula.Enabled
    mnuETHeader.Enabled = txtNETFormula.Enabled
    
    If RobustNETValuesEnabled() Then
        lblETType.Caption = "Using Custom NETs"
    Else
        For i = mnuET.LBound To mnuET.UBound
            If mnuET(i).Checked Then
               lblETType.Caption = "ET: " & mnuET(i).Caption
               SetETMode val(i)
            End If
        Next i
    End If
End Sub

Public Sub ExcludeAmbiguousPairsWrapper(blnOnlyExaminePairsWithHits As Boolean)
    Dim strMessage As String
    
    If blnOnlyExaminePairsWithHits Then
        strMessage = PairsSearchMarkAmbiguousPairsWithHitsOnly(Me, CallerID)
    Else
        strMessage = PairsSearchMarkAmbiguous(Me, CallerID, True)
    End If
    
    UpdateUMCsPairingStatusNow
    UpdateStatus strMessage
End Sub

Public Function ExportMTDBbyUMC(Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional blnExportUMCMembers As Boolean = False, Optional strIniFileName As String = "", Optional ByRef lngErrorNumber As Long, Optional ByRef lngMDID As Long) As String
'--------------------------------------------------------------------------------
' This function exports data to both T_FTICR_Peak_Results and T_FTICR_UMC_Results (plus T_FTICR_UMC_ResultDetails)
' Optionally returns the error number in lngErrorNumber
' Optionally returns the MD_ID value in lngMDID
'--------------------------------------------------------------------------------
    
    Dim strStatus As String
    Dim eResponse As VbMsgBoxResult
    Dim blnAddQuantitationEntry As Boolean
    Dim blnExportUMCsWithNoMatches As Boolean
    
    lngMDID = -1
    mKeyPressAbortProcess = 0
    cmdSearchAllUMCs.Visible = False
    cmdRemoveAMTMatchesFromUMCs.Visible = False
        
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        With glbPreferencesExpanded.AutoAnalysisOptions
            blnAddQuantitationEntry = .AddQuantitationDescriptionEntry
            blnExportUMCsWithNoMatches = .ExportUMCsWithNoMatches
        End With
    Else
        eResponse = MsgBox("Export LC-MS Features that do not have any database matches?", vbQuestion + vbYesNo + vbDefaultButton2, "Export Non-Matching LC-MS Features")
        blnExportUMCsWithNoMatches = (eResponse = vbYes)
    End If
    
    ' September 2004: Unsupported code
    ''strStatus = ExportMTDBbyUMCToPeakResultsTable(lngMDID, blnUpdateGANETForAnalysisInDB, lngErrorNumber)
    
    ' Note: The following function call will create a new entry in T_Match_Making_Description
    strStatus = ExportMTDBbyUMCToUMCResultsTable(lngMDID, blnUpdateGANETForAnalysisInDB, blnExportUMCMembers, lngErrorNumber, blnAddQuantitationEntry, blnExportUMCsWithNoMatches, strIniFileName)
    
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True
    ExportMTDBbyUMC = strStatus
    
End Function

' September 2004: Unused Function
''Private Function ExportMTDBbyUMCToPeakResultsTable(ByRef lngMDID As Long, Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByRef lngErrorNumber As Long) As String
'''---------------------------------------------------
'''this is simple but long procedure of exporting data
'''results to Organism MT tag database associated with gel
'''
'''We're currently writing the results to the T_Match_Making_Description table and T_FTICR_Peak_Results
'''These tables are designed to hold search results from an ion-by-ion search (either using all ions or using UMC ions only)
'''Since this form uses a UMC by UMC search, and we assign all matches for a UMC to all ions for the UMC, we'll
'''  only export the search results for the class representative ion for each UMC (typically the most abundant ion)
'''
'''Returns a status message
'''lngErrorNumber will contain the error number, if an error occurs
'''lngMDID contains the new MMD_ID value
'''---------------------------------------------------
''Const MASS_PRECISION = 6
''Const FIT_PRECISION = 3
''
''Dim mgInd As Long
''Dim lngUMCIndexOriginal As Long, lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long
''Dim ExpCnt As Long
''
''Dim strCaptionSaved As String
''Dim strExportStatus As String
''
''Dim lngPairIndex As Long
''
''Dim objP1IndFastSearch As FastSearchArrayLong
''Dim objP2IndFastSearch As FastSearchArrayLong
''Dim blnPairsPresent As Boolean
''
''Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
''Dim udtPairMatchStats() As udtPairMatchStatsType
''
'''ADO objects for stored procedure adding Match Making row
''Dim cnNew As New ADODB.Connection
''
'''ADO objects for stored procedure that adds FTICR peak rows
''Dim cmdPutNewPeak As New ADODB.Command
''Dim prmMMDID As New ADODB.Parameter
''Dim prmFTICRID As New ADODB.Parameter
''Dim prmFTICRType As New ADODB.Parameter
''Dim prmScanNumber As New ADODB.Parameter
''Dim prmChargeState As New ADODB.Parameter
''Dim prmMonoisotopicMass As New ADODB.Parameter
''Dim prmAbundance As New ADODB.Parameter
''Dim prmFit As New ADODB.Parameter
''Dim prmExpressionRatio As New ADODB.Parameter
''Dim prmLckID As New ADODB.Parameter
''Dim prmFreqShift As New ADODB.Parameter
''Dim prmMassCorrection As New ADODB.Parameter
''Dim prmMassTagID As New ADODB.Parameter
''Dim prmResType As New ADODB.Parameter
''Dim prmHitsCount As New ADODB.Parameter
''Dim prmUMCInd As New ADODB.Parameter
''Dim prmUMCFirstScan As New ADODB.Parameter
''Dim prmUMCLastScan As New ADODB.Parameter
''Dim prmUMCCount As New ADODB.Parameter
''Dim prmUMCAbundance As New ADODB.Parameter
''Dim prmUMCBestFit As New ADODB.Parameter
''Dim prmUMCAvgMW As New ADODB.Parameter
''Dim prmPairInd As New ADODB.Parameter
''
''On Error GoTo err_ExportMTDBbyUMC
''
''strCaptionSaved = Me.Caption
''
''' Connect to the database
''Me.Caption = "Connecting to the database"
''If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
''    Debug.Assert False
''    lngErrorNumber = -1
''    Me.Caption = strCaptionSaved
''    ExportMTDBbyUMCToPeakResultsTable = "Error: Unable to establish a connection to the database"
''    Exit Function
''End If
''
'''first write new analysis in T_Match_Making_Description table
''' Note that we're using CountMassTagsInUMCMatchStats() to determine the number of items in mUMCMatchStats that are MT tags
''AddEntryToMatchMakingDescriptionTable cnNew, lngMDID, ExpAnalysisSPName, CallerID, CountMassTagsInUMCMatchStats(), GelData(CallerID).CustomNETsDefined, True, strIniFileName
''AddToAnalysisHistory CallerID, "Exported UMC Identification results (single UMC mass) to Peak Results table in database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
''AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file
''
'''nothing to export
''If mMatchStatsCount <= 0 Then
''    cnNew.Close
''    Me.Caption = strCaptionSaved
''    Exit Function
''End If
''
''' Initialize the SP
''InitializeSPCommand cmdPutNewPeak, cnNew, ExpPeakSPName
''
''Set prmMMDID = cmdPutNewPeak.CreateParameter("MMDID", adInteger, adParamInput, , lngMDID)
''cmdPutNewPeak.Parameters.Append prmMMDID
''Set prmFTICRID = cmdPutNewPeak.CreateParameter("FTICRID", adVarChar, adParamInput, 50, Null)
''cmdPutNewPeak.Parameters.Append prmFTICRID
''Set prmFTICRType = cmdPutNewPeak.CreateParameter("FTICRType", adTinyInt, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFTICRType
''Set prmScanNumber = cmdPutNewPeak.CreateParameter("ScanNumber", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmScanNumber
''Set prmChargeState = cmdPutNewPeak.CreateParameter("ChargeState", adSmallInt, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmChargeState
''Set prmMonoisotopicMass = cmdPutNewPeak.CreateParameter("MonoisotopicMass", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMonoisotopicMass
''Set prmAbundance = cmdPutNewPeak.CreateParameter("Abundance", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmAbundance
''Set prmFit = cmdPutNewPeak.CreateParameter("Fit", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFit
''Set prmExpressionRatio = cmdPutNewPeak.CreateParameter("ExpressionRatio", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmExpressionRatio
''Set prmLckID = cmdPutNewPeak.CreateParameter("LckID", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmLckID
''Set prmFreqShift = cmdPutNewPeak.CreateParameter("FreqShift", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFreqShift
''Set prmMassCorrection = cmdPutNewPeak.CreateParameter("MassCorrection", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMassCorrection
''Set prmMassTagID = cmdPutNewPeak.CreateParameter("MassTagID", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMassTagID
''Set prmResType = cmdPutNewPeak.CreateParameter("Type", adInteger, adParamInput, , FPR_Type_Standard)
''cmdPutNewPeak.Parameters.Append prmResType
''Set prmHitsCount = cmdPutNewPeak.CreateParameter("HitCount", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmHitsCount
''Set prmUMCInd = cmdPutNewPeak.CreateParameter("UMCInd", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCInd
''Set prmUMCFirstScan = cmdPutNewPeak.CreateParameter("UMCFirstScan", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCFirstScan
''Set prmUMCLastScan = cmdPutNewPeak.CreateParameter("UMCLastScan", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCLastScan
''Set prmUMCCount = cmdPutNewPeak.CreateParameter("UMCCount", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCCount
''Set prmUMCAbundance = cmdPutNewPeak.CreateParameter("UMCAbundance", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCAbundance
''Set prmUMCBestFit = cmdPutNewPeak.CreateParameter("UMCBestFit", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCBestFit
''Set prmUMCAvgMW = cmdPutNewPeak.CreateParameter("UMCAvgMW", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCAvgMW
''Set prmPairInd = cmdPutNewPeak.CreateParameter("PairInd", adInteger, adParamInput, , -1)
''cmdPutNewPeak.Parameters.Append prmPairInd
''
''' Initialize the PairIndex lookup objects
''blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)
''
''Me.Caption = "Exporting peaks to DB: 0 / " & Trim(mMatchStatsCount)
''
'''now export data
''ExpCnt = 0
''With GelData(CallerID)
''    ' Step through the UMC hits and export information on each hit
''    ' Since the target table is an ion-based table, will use the index and info of the class representative ion
''    For mgInd = 0 To mMatchStatsCount - 1
''        If mgInd Mod 25 = 0 Then
''            Me.Caption = "Exporting peaks to DB: " & Trim(mgInd) & " / " & Trim(mMatchStatsCount)
''            DoEvents
''        End If
''
''        If Not mUMCMatchStats(mgInd).IDIsInternalStd Then
''            ' Only export to T_FTICR_Peak_Results if this is a MT tag hit
''            lngUMCIndexOriginal = mUMCMatchStats(mgInd).UMCIndex
''
''            With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
''                prmFTICRID.value = .ClassRepInd
''                prmFTICRType.value = .ClassRepType
''
''                Select Case .ClassRepType
''                Case glCSType
''                    prmScanNumber.value = GelData(CallerID).CSData(.ClassRepInd).ScanNumber
''                    prmChargeState.value = GelData(CallerID).CSData(.ClassRepInd).Charge
''                    prmMonoisotopicMass.value = Round(.ClassMW, MASS_PRECISION)                 ' Mass of the class rep would be: GelData(CallerID).CSData(.ClassRepInd).AverageMW
''                    prmAbundance.value = .ClassAbundance                                        ' Abundance of the class rep would be: GelData(CallerID).CSData(.ClassRepInd).Abundance
''                    prmFit.value = GelData(CallerID).CSData(.ClassRepInd).MassStDev     'standard deviation
''                    If GelLM(CallerID).CSCnt > 0 Then
''                      prmLckID.value = GelLM(CallerID).CSLckID(.ClassRepInd)
''                      prmFreqShift.value = GelLM(CallerID).CSFreqShift(.ClassRepInd)
''                      prmMassCorrection.value = GelLM(CallerID).CSMassCorrection(.ClassRepInd)
''                    End If
''                Case glIsoType
''                    prmScanNumber.value = GelData(CallerID).IsoData(.ClassRepInd).ScanNumber
''                    prmChargeState.value = GelData(CallerID).IsoData(.ClassRepInd).Charge
''                    prmMonoisotopicMass.value = Round(.ClassMW, MASS_PRECISION)                 ' Mass of the class rep would be: GelData(CallerID).IsoData(.ClassRepInd).MonoisotopicMW
''                    prmAbundance.value = .ClassAbundance                                        ' Abundance of the class rep would be: GelData(CallerID).IsoData(.ClassRepInd).Abundance
''                    prmFit.value = GelData(CallerID).IsoData(.ClassRepInd).Fit
''                    If GelLM(CallerID).IsoCnt > 0 Then
''                      prmLckID.value = GelLM(CallerID).IsoLckID(.ClassRepInd)
''                      prmFreqShift.value = GelLM(CallerID).IsoFreqShift(.ClassRepInd)
''                      prmMassCorrection.value = GelLM(CallerID).IsoMassCorrection(.ClassRepInd)
''                    End If
''                End Select
''
''                ' Note: The multi-hit count value for the UMC is the same as that for the class representative, and can thus be placed in prmHitsCount
''                prmHitsCount.value = mUMCMatchStats(mgInd).MultiAMTHitCount
''                prmUMCInd.value = mUMCMatchStats(mgInd).UMCIndex
''                prmUMCFirstScan.value = .MinScan
''                prmUMCLastScan.value = .MaxScan
''                prmUMCCount.value = .ClassCount
''                prmUMCAbundance.value = .ClassAbundance
''                prmUMCBestFit.value = Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), FIT_PRECISION)
''                prmUMCAvgMW.value = Round(.ClassMW, MASS_PRECISION)              ' This is usually the median mass of the class, not the average mass
''
''            End With
''
''            lngMassTagIndexPointer = mMTInd(mUMCMatchStats(mgInd).IDIndex)
''            lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
''
''            prmMassTagID.value = AMTData(lngMassTagIndexOriginal).ID
''
''            lngPairIndex = -1
''            lngPairMatchCount = 0
''            ReDim udtPairMatchStats(0)
''            InitializePairMatchStats udtPairMatchStats(0)
''            If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
''                 blnReturnAllPairInstances = True
''                 blnFavorHeavy = (LastSearchTypeN14N15 = SEARCH_N15)
''                 lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, _
''                                                 objP1IndFastSearch, objP2IndFastSearch, _
''                                                 blnReturnAllPairInstances, blnFavorHeavy, _
''                                                 lngPairMatchCount, udtPairMatchStats())
''            End If
''
''            ' If pairs exist, then we need to output an entry for each pair that this UMC is a member of
''            If lngPairMatchCount > 0 Then
''                For lngPairMatchIndex = 0 To lngPairMatchCount - 1
''                    With udtPairMatchStats(lngPairMatchIndex)
''                        prmPairInd.value = .PairIndex
''                        prmExpressionRatio.value = .ExpressionRatio
''
''                        cmdPutNewPeak.Execute
''                        ExpCnt = ExpCnt + 1
''                    End With
''                Next lngPairMatchIndex
''            Else
''                With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
''                    prmExpressionRatio.value = LookupExpressionRatioValue(CallerID, .ClassRepInd, (.ClassRepType = glIsoType))
''                    prmPairInd.value = -1
''
''                    cmdPutNewPeak.Execute
''                    ExpCnt = ExpCnt + 1
''                End With
''            End If
''        End If
''    Next mgInd
''End With
''
''' MonroeMod
''AddToAnalysisHistory CallerID, "Export to Peak Results table details: UMC Peaks Match Count = " & ExpCnt
''
''Me.Caption = strCaptionSaved
''
''strExportStatus = ExpCnt & " associations between MT tags and LC-MS Features exported to peak results table."
''Set cmdPutNewPeak.ActiveConnection = Nothing
''cnNew.Close
''
''If blnUpdateGANETForAnalysisInDB Then
''    ' Export the the GANET Slope, Intercept, and Fit to the database
''    With GelAnalysis(CallerID)
''        strExportStatus = strExportStatus & vbCrLf & ExportGANETtoMTDB(CallerID, .GANET_Slope, .GANET_Intercept, .GANET_Fit)
''    End With
''End If
''
''Set objP1IndFastSearch = Nothing
''Set objP2IndFastSearch = Nothing
''
''ExportMTDBbyUMCToPeakResultsTable = strExportStatus
''lngErrorNumber = 0
''Exit Function
''
''err_ExportMTDBbyUMC:
''ExportMTDBbyUMCToPeakResultsTable = "Error: " & Err.Number & vbCrLf & Err.Description
''lngErrorNumber = Err.Number
''On Error Resume Next
''If Not cnNew Is Nothing Then cnNew.Close
''Me.Caption = strCaptionSaved
''Set objP1IndFastSearch = Nothing
''Set objP2IndFastSearch = Nothing
''
''End Function

Private Function ExportMTDBbyUMCToUMCResultsTable(ByRef lngMDID As Long, Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByVal blnExportUMCMembers As Boolean = False, Optional ByRef lngErrorNumber As Long, Optional ByVal blnAddQuantitationDescriptionEntry As Boolean = True, Optional ByVal blnExportUMCsWithNoMatches As Boolean = True, Optional ByVal strIniFileName As String = "") As String
    '---------------------------------------------------
    'This function will export data to the T_FTICR_UMC_Results table, T_FTICR_UMC_ResultDetails table,
    '  and T_FTICR_UMC_InternalStandardDetails table
    '
    'It will create a new entry in T_Match_Making_Description
    'If blnAddQuantitationDescriptionEntry = True, then calls ExportMTDBAddQuantitationDescriptionEntry
    '  to create a new entry in T_Quantitation_Description and T_Quantitation_MDIDs
    '
    'Returns a status message
    'lngErrorNumber will contain the error number, if an error occurs
    '---------------------------------------------------
    Dim lngPointer As Long, lngUMCIndex As Long
    Dim lngUMCIndexCompare As Long
    Dim lngUMCIndexOriginal As Long
    Dim lngUMCIndexOriginalLastStored As Long
    
    Dim lngUMCIndexOriginalPairOther As Long
    Dim lngPeakFPRType As Long
    
    Dim lngPairIndex As Long
    
    Dim objP1IndFastSearch As FastSearchArrayLong
    Dim objP2IndFastSearch As FastSearchArrayLong
    Dim blnPairsPresent As Boolean
    Dim blnReturnAllPairInstances As Boolean
    Dim blnFavorHeavy As Boolean
    
    Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
    Dim udtPairMatchStats() As udtPairMatchStatsType
    Dim lngUMCResultsIDReturn() As Long
    Dim lngMatchHitCount As Long
    
    Dim blnContinueCompare As Boolean
    
    Dim lngInternalStdMatchCount As Long
    Dim MassTagExpCnt As Long
    Dim InternalStdExpCnt As Long
    Dim strCaptionSaved As String
    Dim strExportStatus As String
    
    'ADO objects for stored procedure adding Match Making row
    Dim cnNew As New ADODB.Connection
    
    Dim sngDBSchemaVersion As Single
    
    'ADO objects for stored procedure that adds FTICR UMC rows
    Dim cmdPutNewUMC As New ADODB.Command
    Dim udtPutUMCParams As udtPutUMCParamsListType
        
    'ADO objects for stored procedure that adds FTICR UMC member rows
    Dim cmdPutNewUMCMember As New ADODB.Command
    Dim udtPutUMCMemberParams As udtPutUMCMemberParamsListType
        
    'ADO objects for stored procedure adding FTICR UMC Details
    Dim cmdPutNewUMCMatch As New ADODB.Command
    Dim udtPutUMCMatchParams As udtPutUMCMatchParamsListType
    
    'ADO objects for stored procedure adding FTICR UMC Internal Standard Details
    Dim cmdPutNewUMCInternalStdMatch As New ADODB.Command
    Dim udtPutUMCInternalStdMatchParams As udtPutUMCInternalStdMatchParamsListType
    
    'ADO objects for stored procedure adding FTICR UMC CS Stats
    Dim cmdPutNewUMCCSStats As New ADODB.Command
    Dim udtPutUMCCSStatsParams As udtPutUMCCSStatsParamsListType
    
    Dim blnUMCMatchFound() As Boolean       ' 0-based array, used to keep track of whether or not the UMC matched any MT tags or Internal Standards
    Dim blnSetStateToOK As Boolean
    Dim blnOverrideMassNETTolerance As Boolean
    
    On Error GoTo err_ExportMTDBbyUMC
    
    strCaptionSaved = Me.Caption
    
    ' Connect to the database
    Me.Caption = "Connecting to the database"
    If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
        Debug.Assert False
        lngErrorNumber = -1
        Me.Caption = strCaptionSaved
        ExportMTDBbyUMCToUMCResultsTable = "Error: Unable to establish a connection to the database"
        Exit Function
    End If
    
    ' Lookup the DB Schema Version
    sngDBSchemaVersion = LookupDBSchemaVersion(cnNew)
    
    If sngDBSchemaVersion < 2 Then
        ' Force UMC Member export to be false
        blnExportUMCMembers = False
    End If
    
    ' First write new analysis in T_Match_Making_Description table
    blnSetStateToOK = True
    blnOverrideMassNETTolerance = GelData(CallerID).MostRecentSearchUsedSTAC
    lngMatchHitCount = CountMassTagsInUMCMatchStats()
    
    lngErrorNumber = AddEntryToMatchMakingDescriptionTableEx(cnNew, lngMDID, ExpAnalysisSPName, CallerID, _
                                                             lngMatchHitCount, mSearchUsedCustomNETs, _
                                                             blnSetStateToOK, strIniFileName, _
                                                             blnOverrideMassNETTolerance, _
                                                             mSearchSummaryStats.MassToleranceFromSTAC, _
                                                             mSearchSummaryStats.NETToleranceFromSTAC, _
                                                             mSearchSummaryStats.UniqueMTCount1PctFDR, _
                                                             mSearchSummaryStats.UniqueMTCount5PctFDR, _
                                                             mSearchSummaryStats.UniqueMTCount10PctFDR, _
                                                             mSearchSummaryStats.UniqueMTCount25PctFDR, _
                                                             mSearchSummaryStats.UniqueMTCount50PctFDR)
    
    If lngErrorNumber <> 0 Then
        Debug.Assert False
        GoTo err_Cleanup
    End If
    
    If mMatchStatsCount > 0 Or blnExportUMCsWithNoMatches Then
        ' MonroeMod
        AddToAnalysisHistory CallerID, "Exported LC-MS Feature Identification results (single UMC mass) to UMC Results table in database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
        AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file
    End If
    
    'nothing to export
    If mMatchStatsCount <= 0 And Not blnExportUMCsWithNoMatches Then
        cnNew.Close
        Me.Caption = strCaptionSaved
        Exit Function
    End If
    
    ' Initialize cmdPutNewUMC and all of the params in udtPutUMCParams
    ExportMTDBInitializePutNewUMCParams cnNew, cmdPutNewUMC, udtPutUMCParams, lngMDID, ExpUmcSPName
    
    ' Initialize the variables for accessing the AddFTICRUmcMember SP
    ExportMTDBInitializePutNewUMCMemberParams cnNew, cmdPutNewUMCMember, udtPutUMCMemberParams, ExpUMCMemberSPName
    
    ' Initialize the variables for accessing the AddFTICRUmcMatch SP
    ExportMTDBInitializePutUMCMatchParams cnNew, cmdPutNewUMCMatch, udtPutUMCMatchParams, ExpUmcMatchSPName
    
    ' Initialize the variables for accessing the AddFTICRUmcInternalStdMatch SP
    ExportMTDBInitializePutUMCInternalStdMatchParams cnNew, cmdPutNewUMCInternalStdMatch, udtPutUMCInternalStdMatchParams, ExpUmcInternalStdMatchSPName
    
    ' Initialize the variables for accessing the AddFTICRUmcCSStats SP
    ExportMTDBInitializePutUMCCSStatsParams cnNew, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, ExpUMCCSStats
    
    
    ' Initialize the PairIndex lookup objects
    blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)
    
    Select Case LastSearchTypeN14N15
    Case SEARCH_N14
         NTypeStr = MOD_TKN_N14
    Case SEARCH_N15
         NTypeStr = MOD_TKN_N15
    End Select
    
    
    If GelData(CallerID).MostRecentSearchUsedSTAC Then
        ' Populate T_Match_Making_FDR
        Me.Caption = "Exporting STAC Stats to DB"
        ExportMTDBStoreSTACStats cnNew, lngMDID
    End If
     
    
    Me.Caption = "Exporting LC-MS Features to DB: 0 / " & Trim(mMatchStatsCount)
    
    'now export data
    MassTagExpCnt = 0
    InternalStdExpCnt = 0

    ' Step through the UMC hits and export information on each hit
    ' mUMCMatchStats() will contain multiple entries for each UMC if the UMC matched multiple MT tags
    ' Additionally, if the UMC matched an Internal Standard, then that will also be included in mUMCMatchStats()
    ' However, we only want to write one entry for each UMC to T_FTICR_UMC_Results
    ' Thus, we need to keep track of whether or not an entry has been made to T_FTICR_UMC_Results
    ' Luckily, results are stored to mUMCMatchStats() in order of UMC Index
    
    ' We need to keep track of which LC-MS Features are exported to the results table
    ReDim blnUMCMatchFound(GelUMC(CallerID).UMCCnt)
    
    lngUMCIndexOriginalLastStored = -1
    
    For lngPointer = 0 To mMatchStatsCount - 1
        If lngPointer Mod 25 = 0 Then
            Me.Caption = "Exporting LC-MS Features to DB: " & Trim(lngPointer) & " / " & Trim(mMatchStatsCount)
            DoEvents
            If mKeyPressAbortProcess = 2 Then Exit For
        End If
        
        lngUMCIndexOriginal = mUMCMatchStats(lngPointer).UMCIndex
        If lngUMCIndexOriginal <> lngUMCIndexOriginalLastStored Then
            ' Add a new row to T_FTICR_UMC_Results
            ' Note: If we searched only paired LC-MS Features, then record both members of the pairs and set lngPeakFPRType to FPR_Type_N14_N15_L
            '       Additionally, record the pair index in the database and record the opposite pair member
            
            ' Need to perform a look-ahead to determine the number of Internal Standard matches for this UMC Index
            lngInternalStdMatchCount = 0
            lngUMCIndexCompare = lngPointer
            blnContinueCompare = True
            Do
                If mUMCMatchStats(lngUMCIndexCompare).IDIsInternalStd Then
                    lngInternalStdMatchCount = lngInternalStdMatchCount + 1
                End If
                lngUMCIndexCompare = lngUMCIndexCompare + 1
                If lngUMCIndexCompare < mMatchStatsCount Then
                    blnContinueCompare = (mUMCMatchStats(lngUMCIndexCompare).UMCIndex = lngUMCIndexOriginal)
                Else
                    blnContinueCompare = False
                End If
            Loop While blnContinueCompare
            
            lngPairIndex = -1
            lngPairMatchCount = 0
            ReDim udtPairMatchStats(0)
            InitializePairMatchStats udtPairMatchStats(0)
            If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
                blnReturnAllPairInstances = True
                blnFavorHeavy = (LastSearchTypeN14N15 = SEARCH_N15)
                
                lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, _
                                                     objP1IndFastSearch, objP2IndFastSearch, _
                                                     blnReturnAllPairInstances, blnFavorHeavy, _
                                                     lngPairMatchCount, udtPairMatchStats())
            End If
          
            ' If pairs exist, then we need to output an entry for each pair that this UMC is a member of
            If lngPairMatchCount > 0 Then
                ReDim lngUMCResultsIDReturn(lngPairMatchCount - 1)
                
                For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                    ' Lookup whether this UMC is the light or heavy member in the pair
                    With GelP_D_L(CallerID).Pairs(udtPairMatchStats(lngPairMatchIndex).PairIndex)
                        If .p1 = lngUMCIndexOriginal Then
                            lngPeakFPRType = FPR_Type_N14_N15_L      ' Light member of pair
                        Else
                            lngPeakFPRType = FPR_Type_N14_N15_H      ' Heavy member of pair
                        End If
                    End With
                    
                    ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, blnExportUMCMembers, CallerID, lngUMCIndexOriginal, mUMCMatchStats(lngPointer).MultiAMTHitCount, ClsStat(), udtPairMatchStats(lngPairMatchIndex), lngPeakFPRType, lngInternalStdMatchCount
                    blnUMCMatchFound(lngUMCIndexOriginal) = True
        
                    ' Populate array with return value
                    lngUMCResultsIDReturn(lngPairMatchIndex) = FixNullLng(udtPutUMCParams.UMCResultsIDReturn.Value)
            
                    ' Add the other member of the pair too (typically the heavy member)
                    ' Need to determine the UMC index for the other member of the pair
                    With GelP_D_L(CallerID).Pairs(udtPairMatchStats(lngPairMatchIndex).PairIndex)
                        If .p1 = lngUMCIndexOriginal Then
                            ' Already saved the light member, now save the heavy member
                            lngUMCIndexOriginalPairOther = .p2
                            lngPeakFPRType = FPR_Type_N14_N15_H
                        Else
                            ' Already saved the heavy member, now save the light member
                            lngUMCIndexOriginalPairOther = .p1
                            lngPeakFPRType = FPR_Type_N14_N15_L
                        End If
                        
                        ' Always export the other member of the pair, even if it has already been exported
                        ' Note that we do not record any MT tag hits for the other member of the pair
                        ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, blnExportUMCMembers, CallerID, lngUMCIndexOriginalPairOther, 0, ClsStat(), udtPairMatchStats(lngPairMatchIndex), lngPeakFPRType, 0
                        blnUMCMatchFound(lngUMCIndexOriginalPairOther) = True
                        
                    End With
                    
                Next lngPairMatchIndex
            Else
                lngPeakFPRType = FPR_Type_Standard
            
                ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, blnExportUMCMembers, CallerID, lngUMCIndexOriginal, mUMCMatchStats(lngPointer).MultiAMTHitCount, ClsStat(), udtPairMatchStats(0), lngPeakFPRType, lngInternalStdMatchCount
                blnUMCMatchFound(lngUMCIndexOriginal) = True
        
                udtPutUMCMatchParams.UMCResultsID.Value = FixNullLng(udtPutUMCParams.UMCResultsIDReturn.Value)
                udtPutUMCInternalStdMatchParams.UMCResultsID.Value = udtPutUMCMatchParams.UMCResultsID.Value
                
            End If
        End If
        
        ' Now write the MT tag match
        If lngPairMatchCount > 0 Then
            For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                udtPutUMCMatchParams.UMCResultsID.Value = lngUMCResultsIDReturn(lngPairMatchIndex)
                udtPutUMCInternalStdMatchParams.UMCResultsID.Value = lngUMCResultsIDReturn(lngPairMatchIndex)
                
                ExportMTDBbyUMCToUMCResultDetailsTable lngPointer, udtPutUMCInternalStdMatchParams, cmdPutNewUMCInternalStdMatch, udtPutUMCMatchParams, cmdPutNewUMCMatch
            Next lngPairMatchIndex
        Else
            ExportMTDBbyUMCToUMCResultDetailsTable lngPointer, udtPutUMCInternalStdMatchParams, cmdPutNewUMCInternalStdMatch, udtPutUMCMatchParams, cmdPutNewUMCMatch
        End If
            
        If mUMCMatchStats(lngPointer).IDIsInternalStd Then
            ' Increment this if we export an Internal Standard
            InternalStdExpCnt = InternalStdExpCnt + 1
        Else
            ' Increment this if we export a MT tag
            MassTagExpCnt = MassTagExpCnt + 1
        End If
        
        ' Update lngUMCIndexOriginalLastStored
        lngUMCIndexOriginalLastStored = lngUMCIndexOriginal
        
    Next lngPointer

    If blnExportUMCsWithNoMatches And mKeyPressAbortProcess < 2 Then
        ' Also export the LC-MS Features that do not have any hits
        ' If SearchType = SEARCH_PAIRED or SEARCH_NON_PAIRED then only export paired or unpaired LC-MS Features without matches
        
        With GelUMC(CallerID)
            For lngUMCIndex = 0 To .UMCCnt - 1
                If lngUMCIndex Mod 25 = 0 Then
                    Me.Caption = "Exporting non-matching LC-MS Features: " & Trim(lngUMCIndex) & " / " & Trim(.UMCCnt)
                    DoEvents
                    If mKeyPressAbortProcess = 2 Then Exit For
                End If
                
                If Not blnUMCMatchFound(lngUMCIndex) Then
                    ' No match was found
                    If SearchType = SEARCH_ALL Or _
                       SearchType = SEARCH_PAIRED_PLUS_NON_PAIRED Or _
                      (SearchType = SEARCH_PAIRED And eClsPaired(lngUMCIndex) <> umcpNone) Or _
                      (SearchType = SEARCH_NON_PAIRED And eClsPaired(lngUMCIndex) = umcpNone) Then
                    
                        ' Export to the database
                        lngPairIndex = -1
                        lngPairMatchCount = 0
                        ReDim udtPairMatchStats(0)
                        InitializePairMatchStats udtPairMatchStats(0)
                        If eClsPaired(lngUMCIndex) <> umcpNone And blnPairsPresent Then
                            blnReturnAllPairInstances = True
                            blnFavorHeavy = (LastSearchTypeN14N15 = SEARCH_N15)
                            
                            lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndex, _
                                                                 objP1IndFastSearch, objP2IndFastSearch, _
                                                                 blnReturnAllPairInstances, blnFavorHeavy, _
                                                                 lngPairMatchCount, udtPairMatchStats())
                        End If
                            
                        ' If pairs exist, then we need to output an entry for each pair that this UMC is a member of
                        If lngPairMatchCount > 0 Then
                            For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                                ' Lookup whether this UMC is the light or heavy member in the pair
                                With GelP_D_L(CallerID).Pairs(udtPairMatchStats(lngPairMatchIndex).PairIndex)
                                    If .p1 = lngUMCIndex Then
                                        lngPeakFPRType = FPR_Type_N14_N15_L      ' Light member of pair
                                    Else
                                        lngPeakFPRType = FPR_Type_N14_N15_H      ' Heavy member of pair
                                    End If
                                End With
                                        
                                ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, blnExportUMCMembers, CallerID, lngUMCIndex, 0, ClsStat(), udtPairMatchStats(lngPairMatchIndex), lngPeakFPRType, 0
                            Next lngPairMatchIndex
                        Else
                            lngPeakFPRType = FPR_Type_Standard
                        
                            ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, blnExportUMCMembers, CallerID, lngUMCIndex, 0, ClsStat(), udtPairMatchStats(0), lngPeakFPRType, 0
                        End If
                            
                    End If
                End If
            Next lngUMCIndex
        End With
    End If

    ' MonroeMod
    AddToAnalysisHistory CallerID, "Export to LC-MS Feature Results table details: MT tags Match Count = " & MassTagExpCnt & "; Internal Std Match Count = " & InternalStdExpCnt
    
    Me.Caption = strCaptionSaved
    
    strExportStatus = MassTagExpCnt & " associations between MT tags and LC-MS Features exported (" & Trim(InternalStdExpCnt) & " Internal Standards)."
    Set cmdPutNewUMC.ActiveConnection = Nothing
    Set cmdPutNewUMCMatch.ActiveConnection = Nothing
    cnNew.Close
    
    If blnUpdateGANETForAnalysisInDB Then
        ' Export the the GANET Slope, Intercept, and Fit to the database
        With GelAnalysis(CallerID)
            strExportStatus = strExportStatus & vbCrLf & ExportGANETtoMTDB(CallerID, .GANET_Slope, .GANET_Intercept, .GANET_Fit)
        End With
    End If
    
    If blnAddQuantitationDescriptionEntry Then
        If lngErrorNumber = 0 And lngMDID >= 0 And (MassTagExpCnt > 0 Or InternalStdExpCnt > 0) Then
            ExportMTDBAddQuantitationDescriptionEntry Me, CallerID, ExpQuantitationDescription, lngMDID, lngErrorNumber, strIniFileName, 1, 1, 1, Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
        End If
    End If
    
    ExportMTDBbyUMCToUMCResultsTable = strExportStatus
    lngErrorNumber = 0
    Set objP1IndFastSearch = Nothing
    Set objP2IndFastSearch = Nothing

Exit Function

err_ExportMTDBbyUMC:
Debug.Assert False
LogErrors Err.Number, "ExportMTDBbyUMCToUMCResultsTable"
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "Error exporting matches to the LC-MS Feature results table: " & Err.Description, vbExclamation + vbOKOnly, glFGTU
End If

err_Cleanup:
On Error Resume Next
If Not cnNew Is Nothing Then cnNew.Close
Me.Caption = strCaptionSaved
Set objP1IndFastSearch = Nothing
Set objP2IndFastSearch = Nothing

If Err.Number <> 0 Then lngErrorNumber = Err.Number
ExportMTDBbyUMCToUMCResultsTable = "Error: " & lngErrorNumber & vbCrLf & Err.Description

End Function

Private Function ExportMTDBbyUMCToUMCResultDetailsTable(lngPointer As Long, ByRef udtPutUMCInternalStdMatchParams As udtPutUMCInternalStdMatchParamsListType, ByRef cmdPutNewUMCInternalStdMatch As ADODB.Command, ByRef udtPutUMCMatchParams As udtPutUMCMatchParamsListType, cmdPutNewUMCMatch As ADODB.Command)

    Dim lngInternalStdIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long

    Dim strMassMods As String

    If mUMCMatchStats(lngPointer).IDIsInternalStd Then
    
        ' Write an entry to T_FTICR_UMC_InternalStdDetails
        lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(lngPointer).IDIndex)

        udtPutUMCInternalStdMatchParams.SeqID.Value = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).SeqID
        udtPutUMCInternalStdMatchParams.MatchingMemberCount.Value = mUMCMatchStats(lngPointer).MemberHitCount
        udtPutUMCInternalStdMatchParams.MatchScore.Value = mUMCMatchStats(lngPointer).StacOrSLiC
        udtPutUMCInternalStdMatchParams.ExpectedNET.Value = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).NET
        udtPutUMCInternalStdMatchParams.DelMatchScore.Value = mUMCMatchStats(lngPointer).DelScore
        udtPutUMCInternalStdMatchParams.UniquenessProbability.Value = CSng(mUMCMatchStats(lngPointer).UniquenessProbability)
        udtPutUMCInternalStdMatchParams.FDRThreshold.Value = CSng(mUMCMatchStats(lngPointer).FDRThreshold)
        
        cmdPutNewUMCInternalStdMatch.Execute
        
    Else
        ' Write an entry to T_FTICR_UMC_ResultDetails
        
        lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngPointer).IDIndex)
        lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
    
        udtPutUMCMatchParams.MassTagID.Value = AMTData(lngMassTagIndexOriginal).ID
        udtPutUMCMatchParams.MatchingMemberCount.Value = mUMCMatchStats(lngPointer).MemberHitCount
        udtPutUMCMatchParams.MatchScore.Value = mUMCMatchStats(lngPointer).StacOrSLiC
        
        strMassMods = NTypeStr
        If Len(mMTMods(lngMassTagIndexPointer)) > 0 Then
            strMassMods = strMassMods & " " & Trim(mMTMods(lngMassTagIndexPointer))
            If NTypeStr = MOD_TKN_N14 Then
                udtPutUMCMatchParams.MassTagModMass.Value = CSng(mMTMWN14(mUMCMatchStats(lngPointer).IDIndex) - AMTData(lngMassTagIndexOriginal).MW)
            Else
                udtPutUMCMatchParams.MassTagModMass.Value = CSng(mMTMWN15(mUMCMatchStats(lngPointer).IDIndex) - AMTData(lngMassTagIndexOriginal).MW)
            End If
        Else
            If NTypeStr = MOD_TKN_N14 Then
                udtPutUMCMatchParams.MassTagModMass.Value = 0
            Else
                udtPutUMCMatchParams.MassTagModMass.Value = CSng(glN14N15_DELTA * AMTData(lngMassTagIndexOriginal).CNT_N)
            End If
        End If
        
        If Len(strMassMods) > PUT_UMC_MATCH_MAX_MODSTRING_LENGTH Then strMassMods = Left(strMassMods, PUT_UMC_MATCH_MAX_MODSTRING_LENGTH)
        udtPutUMCMatchParams.MassTagMods.Value = strMassMods
        
        udtPutUMCMatchParams.DelMatchScore.Value = mUMCMatchStats(lngPointer).DelScore
        udtPutUMCMatchParams.UniquenessProbability.Value = CSng(mUMCMatchStats(lngPointer).UniquenessProbability)
        udtPutUMCMatchParams.FDRThreshold.Value = CSng(mUMCMatchStats(lngPointer).FDRThreshold)
        
        cmdPutNewUMCMatch.Execute
    
    End If

End Function

 
Private Sub ExportMTDBStoreSTACStats(ByRef cnNew As ADODB.Connection, ByVal lngMDID As Long)

    ' Adds new rows to the T_Match_Making_FDR table
 
    Dim cmdStoreSTACStats As New ADODB.Command
    Dim udtStoreSTACStatsParams As udtStoreSTACStatsParamsListType
    
    Dim lngIndex As Long
    
     ' Initialize the variables for accessing the AddMatchMakingFDR SP
    ExportMTDBInitializeStoreSTACStats cnNew, cmdStoreSTACStats, udtStoreSTACStatsParams, ExpStoreSTACStats


On Error GoTo AddMatchMakingFDRRowErrorHandler


    udtStoreSTACStatsParams.MDID = lngMDID
    
    For lngIndex = 0 To STACStatsCount - 1
            
        ' Note: Dividing the FDR values by 100 prior to storing in the DB
        
        With STACStats(lngIndex)
            udtStoreSTACStatsParams.STACCutoff = .STACCuttoff
            udtStoreSTACStatsParams.Matches = .Matches
            udtStoreSTACStatsParams.Errors = .Errors
            udtStoreSTACStatsParams.FDR = .FDR / 100#
            
            udtStoreSTACStatsParams.UPFilteredMatches = .UP_Filtered_Matches
            udtStoreSTACStatsParams.UPFilteredErrors = .UP_Filtered_Errors
            udtStoreSTACStatsParams.UPFilteredFDR = .UP_Filtered_FDR / 100#
            
        End With
        
        cmdStoreSTACStats.Execute
        
    Next lngIndex

    Exit Sub

AddMatchMakingFDRRowErrorHandler:
    ' Error populating or executing cmdStoreSTACStats
    
    Debug.Assert False
    
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->ExportMTDBStoreSTACStats"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error storing STAC results in the database: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub GenerateUniqueMatchStats()
    ' Determine the number of LC-MS Features with at least one match,
    ' the unique number of MT tags matched, and the unique number of Internal Standards matched
    
    Dim htMTHitList1Pct As Dictionary
    Dim htMTHitList5Pct As Dictionary
    Dim htMTHitList10Pct As Dictionary
    Dim htMTHitList25Pct As Dictionary
    Dim htMTHitList50Pct As Dictionary
    
    Dim blnUMCHasMatch() As Boolean
    Dim blnPMTTagMatched() As Boolean
    Dim blnInternalStdMatched() As Boolean
    
    Dim lngIndex As Long
    Dim lngUMCIndexOriginal As Long
    Dim lngInternalStdIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long
    Dim lngMassTagIndexOriginal As Long
     
On Error GoTo GenerateUniqueMatchStatsErrorHandler

    With mSearchSummaryStats
        .UMCCountWithHits = 0
        .UniqueMTCount = 0
        .UniqueIntStdCount = 0
        .UniqueMTCount1PctFDR = 0
        .UniqueMTCount5PctFDR = 0
        .UniqueMTCount10PctFDR = 0
        .UniqueMTCount25PctFDR = 0
        .UniqueMTCount50PctFDR = 0
        
        ' Don't clear .MassToleranceFromSTAC or .NETToleranceFromSTAC; they've already been populated
    End With
    
    ReDim blnUMCHasMatch(GelUMC(CallerID).UMCCnt - 1) As Boolean
    
    ReDim blnPMTTagMatched(AMTCnt) As Boolean
    If UMCInternalStandards.Count > 0 Then
        ReDim blnInternalStdMatched(UMCInternalStandards.Count - 1) As Boolean
    Else
        ReDim blnInternalStdMatched(0)
    End If
    
     ' The following are populated with features passing a given FDR Threshold
    Set htMTHitList1Pct = New Dictionary
    Set htMTHitList5Pct = New Dictionary
    Set htMTHitList10Pct = New Dictionary
    Set htMTHitList25Pct = New Dictionary
    Set htMTHitList50Pct = New Dictionary
    
    htMTHitList1Pct.RemoveAll
    htMTHitList5Pct.RemoveAll
    htMTHitList10Pct.RemoveAll
    htMTHitList25Pct.RemoveAll
    htMTHitList50Pct.RemoveAll
    
    For lngIndex = 0 To mMatchStatsCount - 1
        lngUMCIndexOriginal = mUMCMatchStats(lngIndex).UMCIndex
        If lngUMCIndexOriginal < GelUMC(CallerID).UMCCnt Then
            blnUMCHasMatch(lngUMCIndexOriginal) = True
        Else
            ' Invalid UMC index
            Debug.Assert False
        End If
        
        If mUMCMatchStats(lngIndex).IDIsInternalStd Then
            lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(lngIndex).IDIndex)
            If lngInternalStdIndexOriginal < UMCInternalStandards.Count Then
                blnInternalStdMatched(lngInternalStdIndexOriginal) = True
            Else
                ' Invalid Internal Standard index
                Debug.Assert False
            End If
        Else
            lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngIndex).IDIndex)
            lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            If lngMassTagIndexOriginal <= AMTCnt Then
                blnPMTTagMatched(lngMassTagIndexOriginal) = True
            Else
                ' Invalid MT tag index
                Debug.Assert False
            End If
       
            If mUMCMatchStats(lngIndex).FDRThreshold <= 0.5 Then
                If Not htMTHitList50Pct.Exists(lngMassTagIndexOriginal) Then
                    htMTHitList50Pct.add lngMassTagIndexOriginal, 1
                End If
                
                If mUMCMatchStats(lngIndex).FDRThreshold <= 0.25 Then
                If Not htMTHitList25Pct.Exists(lngMassTagIndexOriginal) Then
                    htMTHitList25Pct.add lngMassTagIndexOriginal, 1
                End If
                        
                    If mUMCMatchStats(lngIndex).FDRThreshold <= 0.1 Then
                        If Not htMTHitList10Pct.Exists(lngMassTagIndexOriginal) Then
                            htMTHitList10Pct.add lngMassTagIndexOriginal, 1
                        End If
                        
                        If mUMCMatchStats(lngIndex).FDRThreshold <= 0.05 Then
                            If Not htMTHitList5Pct.Exists(lngMassTagIndexOriginal) Then
                                htMTHitList5Pct.add lngMassTagIndexOriginal, 1
                            End If
                               
                            If mUMCMatchStats(lngIndex).FDRThreshold <= 0.01 Then
                                If Not htMTHitList1Pct.Exists(lngMassTagIndexOriginal) Then
                                    htMTHitList1Pct.add lngMassTagIndexOriginal, 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
        End If
    Next lngIndex
    
    With mSearchSummaryStats
        .UMCCountWithHits = 0
        For lngIndex = 0 To UBound(blnUMCHasMatch)
            If blnUMCHasMatch(lngIndex) Then .UMCCountWithHits = .UMCCountWithHits + 1
        Next lngIndex
        
        .UniqueMTCount = 0
        For lngIndex = 0 To UBound(blnPMTTagMatched)
            If blnPMTTagMatched(lngIndex) Then .UniqueMTCount = .UniqueMTCount + 1
        Next lngIndex
        
        .UniqueIntStdCount = 0
        If UMCInternalStandards.Count > 0 Then
            For lngIndex = 0 To UBound(blnInternalStdMatched)
                If blnInternalStdMatched(lngIndex) Then .UniqueIntStdCount = .UniqueIntStdCount + 1
            Next lngIndex
        End If
        
        .UniqueMTCount1PctFDR = htMTHitList1Pct.Count
        .UniqueMTCount5PctFDR = htMTHitList5Pct.Count
        .UniqueMTCount10PctFDR = htMTHitList10Pct.Count
        .UniqueMTCount25PctFDR = htMTHitList25Pct.Count
        .UniqueMTCount50PctFDR = htMTHitList50Pct.Count
    End With
    
    Exit Sub

GenerateUniqueMatchStatsErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "GenerateUniqueMatchStats"
    
End Sub

Private Function GetDBSearchModeType() As Byte
    If optDBSearchModType(MODS_DECOY).Value Then
        GetDBSearchModeType = 2
    ElseIf optDBSearchModType(MODS_DYNAMIC).Value Then
        GetDBSearchModeType = 1
    Else
        ' Assume mode MODS_FIXED mods
        GetDBSearchModeType = 0
    End If
End Function

Private Function GetTempReportFilePath() As String
    
    Dim strTempReportFilePath As String
    
    If mSTACSessionID = "" Then
        mSTACSessionID = "_" & CLng(Timer()) & "_" & CLng(Rnd(1) * 100000)
    End If

    strTempReportFilePath = GetTempFolder() & "VIPER_Report " & mSTACSessionID
    
    If Not mTempFilesToDelete.Exists(strTempReportFilePath) Then
        mTempFilesToDelete.add strTempReportFilePath, TEMP_FILE_FLAG
    End If

    GetTempReportFilePath = strTempReportFilePath
    
End Function

Public Sub GetSummaryStats(ByRef UMCCountWithHits As Long, _
                           ByRef UniqueMTCount As Long, _
                           ByRef UniqueIntStdCount As Long, _
                           ByRef UniqueMTCount1PctFDR As Long, _
                           ByRef UniqueMTCount5PctFDR As Long, _
                           ByRef UniqueMTCount10PctFDR As Long, _
                           ByRef UniqueMTCount25PctFDR As Long, _
                           ByRef UniqueMTCount50PctFDR As Long, _
                           ByRef MassToleranceFromSTAC As Double, _
                           ByRef NETToleranceFromSTAC As Double)
                            
    With mSearchSummaryStats
        UMCCountWithHits = .UMCCountWithHits
        UniqueMTCount = .UniqueMTCount
        UniqueIntStdCount = .UniqueIntStdCount
        UniqueMTCount1PctFDR = .UniqueMTCount1PctFDR
        UniqueMTCount5PctFDR = .UniqueMTCount5PctFDR
        UniqueMTCount10PctFDR = .UniqueMTCount10PctFDR
        UniqueMTCount25PctFDR = .UniqueMTCount25PctFDR
        UniqueMTCount50PctFDR = .UniqueMTCount50PctFDR
        MassToleranceFromSTAC = .MassToleranceFromSTAC
        NETToleranceFromSTAC = .NETToleranceFromSTAC
    End With
    
End Sub

Private Function GetTokenValue(ByVal S As String, ByVal t As String) As Long
'---------------------------------------------------------------------------
'returns value next to token T in string of type Token1/Value1 Token2/Value2
'-1 if not found or on any error
'---------------------------------------------------------------------------
Dim SSplit() As String
Dim MSplit() As String
Dim i As Long
On Error GoTo exit_GetTokenValue
GetTokenValue = -1
SSplit = Split(S, " ")
For i = 0 To UBound(SSplit)
    If Len(SSplit(i)) > 0 Then
        If InStr(SSplit(i), "/") > 0 Then
            MSplit = Split(SSplit(i), "/")
            If Trim$(MSplit(0)) = t Then
               If IsNumeric(MSplit(1)) Then
                  GetTokenValue = CLng(MSplit(1))
                  Exit Function
               End If
            End If
        End If
    End If
Next i
Exit Function

exit_GetTokenValue:
Debug.Assert False

End Function

Private Function GetWobbledNET(ByVal dblNET As Double, ByVal dblNETWobbleDistance As Double) As Double
    Static PosWobble As Long
    Static NegWobble As Long

    If Rnd() < 0.5 Then
        GetWobbledNET = dblNET - dblNETWobbleDistance
        NegWobble = NegWobble + 1
    Else
        GetWobbledNET = dblNET + dblNETWobbleDistance
        PosWobble = PosWobble + 1
    End If

End Function

Private Function InitExprEvaluator(ByVal sExpr As String) As Boolean
'-------------------------------------------------------------------
'initializes expression evaluator for elution time
'-------------------------------------------------------------------
On Error Resume Next
Set MyExprEva = New ExprEvaluator
With MyExprEva
    .Vars.add 1, "FN"
    .Vars.add 2, "MinFN"
    .Vars.add 3, "MaxFN"
    .Expr = sExpr
    InitExprEvaluator = .IsExprValid
    ReDim VarVals(1 To 3)
End With
End Function

Private Function InitializeORFInfo(blnForceDataReload As Boolean) As Boolean
    ' Initializes objMTDBNameLookupClass
    ' Returns True if success, False if failure
    ' If the class has already been initialized, then does nothing, unless blnForceDataReload = True
    
    Dim blnSuccess As Boolean
    
    If Not objMTDBNameLookupClass Is Nothing Then
        If Not blnForceDataReload Then
            If objMTDBNameLookupClass.DataStatus = dsLoaded Then
                InitializeORFInfo = True
                Exit Function
            End If
        End If
        
        objMTDBNameLookupClass.DeleteData
        Set objMTDBNameLookupClass = Nothing
    End If
    
    Set objMTDBNameLookupClass = New mtdbMTNames
    
    With objMTDBNameLookupClass
        'loading protein names
        UpdateStatus "Loading Protein info"
        
        If Not GelAnalysis(CallerID) Is Nothing Then
            If Len(GelAnalysis(CallerID).MTDB.cn.ConnectionString) > 0 And Not APP_BUILD_DISABLE_MTS Then
                Me.MousePointer = vbHourglass
                .DBConnectionString = GelAnalysis(CallerID).MTDB.cn.ConnectionString
                .RetrieveSQL = glbPreferencesExpanded.MTSConnectionInfo.sqlGetMTNames
                If .FillData(Me) Then
                   If .DataStatus = dsLoaded Then
                        blnSuccess = True
                    End If
                End If
                Me.MousePointer = vbDefault
            End If
        End If
    End With
    
    InitializeORFInfo = blnSuccess
End Function

Public Sub InitializeSearch()
'------------------------------------------------------------------------------------
'load MT tag database data if neccessary
'if CallerID is associated with MT tag database load that db if not already loaded
'if CallerID is not associated with MT tag database load legacy database
'------------------------------------------------------------------------------------
Dim eResponse As VbMsgBoxResult

On Error Resume Next
Me.MousePointer = vbHourglass
If bLoading Then
    ' Update lblUMCMassMode to reflect the mass mode used to identify the LC-MS Features
    Select Case GelUMC(CallerID).def.ClassMW
    Case UMCClassMassConstants.UMCMassAvg
        lblUMCMassMode = "LC-MS Feature Mass = Average of the masses of the UMC members"
    Case UMCClassMassConstants.UMCMassRep
        lblUMCMassMode = "LC-MS Feature Mass = Mass of the UMC Class Representative"
    Case UMCClassMassConstants.UMCMassMed
        lblUMCMassMode = "LC-MS Feature Mass = Median of the masses of the UMC members"
    Case UMCMassAvgTopX
        lblUMCMassMode = "LC-MS Feature Mass = Average of top X members of the UMC"
    Case UMCMassMedTopX
        lblUMCMassMode = "LC-MS Feature Mass = Median of top X members of the UMC"
    Case Else
        lblUMCMassMode = "LC-MS Feature Mass = ?? Unable to determine; is it a new mass mode?"
    End Select
    
    If GelAnalysis(CallerID) Is Nothing Then
        If AMTCnt > 0 Then    'something is loaded
          If (Len(CurrMTDatabase) > 0 Or Len(CurrLegacyMTDatabase) > 0) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                'MT tag data; we dont know is it appropriate; warn user
                WarnUserUnknownMassTags CallerID
            End If
            lblMTStatus.Caption = ConstructMTStatusText(True)
          
            ' Initialize the MT search object
            If Not CreateNewMTSearchObject() Then
                lblMTStatus.Caption = "Error creating search object."
            Else
               ' Error initializing MT search object
            End If
       
       Else                  'nothing is loaded
            If Len(GelData(CallerID).PathtoDatabase) > 0 And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                If APP_BUILD_DISABLE_MTS Then
                    eResponse = vbYes
                Else
                    eResponse = MsgBox("Current display is not associated with any MT tag database.  Do you want to load the MT tags from the defined legacy MT tag database?" & vbCrLf & GelData(CallerID).PathtoDatabase, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load Legacy MT tags")
                End If
            Else
                eResponse = vbNo
            End If
            
            If eResponse = vbYes Then
                LoadLegacyMassTags
            Else
                WarnUserNotConnectedToDB CallerID, True
                lblMTStatus.Caption = "No MT tags loaded"
            End If
        End If
    Else         'have to have MT tag database loaded
        Call LoadMTDB
    End If
    UpdateStatus "Generating LC-MS Feature statistics ..."
    ClsCnt = UMCStatistics1(CallerID, ClsStat())
    UpdateStatus "Pairs Count: " & GelP_D_L(CallerID).PCnt
    
    chkDisableCustomNETs.Enabled = GelData(CallerID).CustomNETsDefined
    If APP_BUILD_DISABLE_LCMSWARP Then
        chkDisableCustomNETs.Visible = chkDisableCustomNETs.Enabled
    End If
    
    EnableDisableNETFormulaControls
    
    SetETMode etGANET
   
    UpdateStatus "LC-MS Features pairing status ..."
    UpdateUMCsPairingStatusNow
    UpdateStatus "Ready"
    
    'memorize number of scans (to be used with elution)
    MinFN = GelData(CallerID).ScanInfo(1).ScanNumber
    MaxFN = GelData(CallerID).ScanInfo(UBound(GelData(CallerID).ScanInfo)).ScanNumber
    bLoading = False
End If
Me.MousePointer = vbDefault
End Sub

Private Sub InitializeSTACStatsListView()

    lvwSTACStats.View = lvwReport

    lvwSTACStats.ColumnHeaders.Clear
    
    lvwSTACStats.ColumnHeaders.add , , "Cutoff", 800
    
    lvwSTACStats.ColumnHeaders.add , , "Matches", 900
    lvwSTACStats.ColumnHeaders.add , , "Errors", 900
    lvwSTACStats.ColumnHeaders.add , , "FDR %", 800
    
    lvwSTACStats.ColumnHeaders.add , , "Matches, UP>0.5", 1100
    lvwSTACStats.ColumnHeaders.add , , "Errors, UP>0.5", 1100
    lvwSTACStats.ColumnHeaders.add , , "FDR %, UP>0.5", 1000
    
End Sub

Private Function IsValidMatch(dblCurrMW As Double, AbsMWErr As Double, CurrScan As Long, dblMatchNET As Double, dblMatchMass As Double) As Boolean
    ' Checks if dblCurrMW is within tolerance of the given MT tag
    ' Also checks if the NET equivalent of CurrScan is within tolerance of the NET value for the given MT tag
    ' Returns True if both are within tolerance, false otherwise
    
    Dim InvalidMatch As Boolean
    
    ' If dblCurrMW is not within AbsMWErr of dblMatchMass then this match is inherited
    If Abs(dblCurrMW - dblMatchMass) > AbsMWErr Then
        InvalidMatch = True
    Else
        ' If CurrScan is not within .NETTol of dblMatchNET then this match is inherited
        If samtDef.NETTol >= 0 Then
            If Abs(ConvertScanToNET(CurrScan) - dblMatchNET) > samtDef.NETTol Then
                InvalidMatch = True
            End If
        End If
    End If
    
    IsValidMatch = Not InvalidMatch
End Function

Private Sub LoadLegacyMassTags()

    '------------------------------------------------------------
    'load/reload MT tags
    '------------------------------------------------------------
    Dim eResponse As VbMsgBoxResult
    On Error Resume Next
    'ask user if it wants to replace legitimate MT tag DB with legacy DB
    If Not GelAnalysis(CallerID) Is Nothing And Not APP_BUILD_DISABLE_MTS Then
       eResponse = MsgBox("Current display is associated with MT tag database." & vbCrLf _
                    & "Are you sure you want to use a legacy database for search?", vbYesNoCancel, glFGTU)
       If eResponse <> vbYes Then Exit Sub
    End If
    Me.MousePointer = vbHourglass
    If Len(GelData(CallerID).PathtoDatabase) > 0 Then
       If ConnectToLegacyAMTDB(Me, CallerID, False, True, False) Then
          If CreateNewMTSearchObject() Then
             lblMTStatus.Caption = "Loaded; MT tag count: " & LongToStringWithCommas(AMTCnt)
          Else
             lblMTStatus.Caption = "Error creating search object."
          End If
       Else
          lblMTStatus.Caption = "Error loading MT tags."
       End If
    Else
        WarnUserInvalidLegacyDBPath
    End If
    Me.MousePointer = vbDefault

End Sub

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean
    Dim strMessage As String
    
    Static blnWorking As Boolean
    
    If blnWorking Then Exit Sub
    blnWorking = True
    
    cmdSearchAllUMCs.Enabled = False
    
    Dim blnLoadMTStats As Boolean
    blnLoadMTStats = False
    
    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, blnLoadMTStats, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTStatus.Caption = ConstructMTStatusText(True)
    
        If Not CreateNewMTSearchObject() Then
           lblMTStatus.Caption = "Error creating search object."
        End If
    Else
        If blnDBConnectionError Then
            strMessage = "Error loading MT tags: database connection error."
        Else
            If Not GelAnalysis(CallerID) Is Nothing Then
                If Len(GelAnalysis(CallerID).MTDB.cn.ConnectionString) > 0 And Not APP_BUILD_DISABLE_MTS Then
                    strMessage = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
                ElseIf Len(GelData(CallerID).PathtoDatabase) > 0 Then
                    strMessage = "Error loading MT tags from Legacy DB"
                Else
                    strMessage = "Error loading MT tags: MT tag database not defined"
                End If
            Else
                strMessage = "Error loading MT tags: MT tag database not defined"
            End If
        End If
    
        lblMTStatus.Caption = strMessage
    End If
    
    cmdSearchAllUMCs.Enabled = True
    blnWorking = False
    
End Sub

Private Function LoadSTACLogFile(ByRef fso As FileSystemObject, _
                                 ByVal strSTACLogFilePath As String) As Boolean

    Dim ts As TextStream
    
    Dim lngErrorCount As Long
    
    Dim strLineIn As String
    Dim strMessage As String
   Dim dblValue As Double
   
On Error GoTo LoadSTACLogFileErrorHandler

    ' Read the log messages from STAC
     
    If fso.FileExists(strSTACLogFilePath) Then
        
        Set ts = fso.OpenTextFile(strSTACLogFilePath, ForReading, False)
    
        lngErrorCount = 0
    
        Do While Not ts.AtEndOfStream
            strLineIn = ts.ReadLine
            
            If StartsWith(strLineIn, "Error") Then
                Debug.Assert False
                    
                If lngErrorCount = 0 Then
                    AddToAnalysisHistory CallerID, "STAC log file contains error messages:"
                End If
                
                AddToAnalysisHistory CallerID, "  " & strLineIn
                
                lngErrorCount = lngErrorCount + 1
                
            ElseIf StartsWith(strLineIn, "ToleranceNET") Then
                If LoadSTACLogFileGetValue(strLineIn, dblValue) Then
                    mSearchSummaryStats.NETToleranceFromSTAC = dblValue
                End If
                
            ElseIf StartsWith(strLineIn, "ToleranceMassPPM") Then
                If LoadSTACLogFileGetValue(strLineIn, dblValue) Then
                    mSearchSummaryStats.MassToleranceFromSTAC = dblValue
                End If
                
            End If
            
        Loop
    
        ts.Close
        
    End If
    
    LoadSTACLogFile = True
    Exit Function

LoadSTACLogFileErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->LoadSTACLogFile"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error loading STAC results: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    LoadSTACLogFile = False
    
End Function

Private Function LoadSTACLogFileGetValue(ByVal strText As String, ByRef dblValue As Double) As Boolean
    Dim strSplitLine() As String
    
On Error GoTo LoadSTACLogFileGetValueErrorHandler
    strSplitLine = Split(strText, ":")
    
    If UBound(strSplitLine) > 0 Then
        If IsNumeric(strSplitLine(1)) Then
            dblValue = CDbl(strSplitLine(1))
            LoadSTACLogFileGetValue = True
        Else
            LoadSTACLogFileGetValue = False
        End If
    Else
        LoadSTACLogFileGetValue = False
    End If
    
    Exit Function
    
LoadSTACLogFileGetValueErrorHandler:
    Debug.Assert False
    LoadSTACLogFileGetValue = False

End Function

Private Function LoadSTACResults(ByRef fso As FileSystemObject, _
                                 ByVal strSTACDataFilePath As String) As Boolean

    Dim ts As TextStream
    
    Dim lngLinesRead As Long
    
    Dim strLineIn As String
    Dim strMessage As String
    
    Dim strSplitLine() As String
   
    Dim lngUMCIndexSaved As Long
    
    Dim blnValidData As Boolean
    
    Dim lngMassTagID As Long
    Dim lngUMCIndex As Long
    Dim dblSTACScore As Double
    Dim dblMassError As Double
    Dim dblNETError As Double
    Dim dblUP As Double
    
    Dim lngMemberIndex As Long
    Dim lngInternalStdIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long
    Dim lngMassTagIndexOriginal As Long
                        
    Dim dblCurrMW As Double
    Dim dblMatchMass As Double
    Dim dblMatchNET As Double
    Dim dblNETTol As Double
    Dim dblNETDifference As Double
    
    Dim MWTolAbsBroad As Double
    Dim NETTolBroad As Double
    Dim MWTolAbsFinal As Double
    Dim NETTolFinal As Double
    
On Error GoTo LoadSTACResultsErrorHandler

    If Not ManageCurrID(MNG_RESET) Then
        ' Report Memory Management Error
        Debug.Assert False
        UpdateStatus "Error managing memory."
    End If
    
    Set ts = fso.OpenTextFile(strSTACDataFilePath, ForReading, False)
   
    lngUMCIndexSaved = -1
    
    Do While Not ts.AtEndOfStream
        strLineIn = ts.ReadLine
        
        If Len(strLineIn) > 0 Then
            strSplitLine = Split(strLineIn, ",")
             
            If UBound(strSplitLine) >= 5 Then
                blnValidData = LoadSTACResultsParseLine(strSplitLine, _
                                                        lngMassTagID, _
                                                        lngUMCIndex, _
                                                        dblSTACScore, _
                                                        dblMassError, _
                                                        dblNETError, _
                                                        dblUP)
                
                If blnValidData Then
                                   
                    If lngUMCIndex <> lngUMCIndexSaved Then
                    
                        ' Process the data for the previous feature
                        
                        LoadSTACResultsUpdateUMC lngUMCIndexSaved
                    
                        lngUMCIndexSaved = lngUMCIndex
                
                        mCurrIDCnt = 0
                        
                        ' Define the tolerances to use for the current UMC
                        SearchAMTDefineTolerances CallerID, lngUMCIndex, samtDef, GelUMC(CallerID).UMCs(lngUMCIndex).ClassMW, MWTolAbsBroad, NETTolBroad, MWTolAbsFinal, NETTolFinal

                    End If
                
                    
                    If mCurrIDCnt > UBound(mCurrIDMatches) Then ManageCurrID (MNG_ADD_START_SIZE)
                                            
                    If lngMassTagID >= mMTCnt Then
                        ' Note: Need to subtract mMTCnt from .IDInd to get the correct location in mInternalStdIndexPointers()
                        mCurrIDMatches(mCurrIDCnt).IDInd = lngMassTagID - mMTCnt
                        mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = True
                    Else
                        mCurrIDMatches(mCurrIDCnt).IDInd = lngMassTagID
                        mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = False
                    End If
                    
                    mCurrIDMatches(mCurrIDCnt).MatchingMemberCount = 0
                    mCurrIDMatches(mCurrIDCnt).StacOrSLiC = dblSTACScore
                    mCurrIDMatches(mCurrIDCnt).MassErr = dblMassError
                    mCurrIDMatches(mCurrIDCnt).NETErr = dblNETError
                    mCurrIDMatches(mCurrIDCnt).UniquenessProbability = dblUP

                    If mCurrIDMatches(mCurrIDCnt).IDIsInternalStd Then
                        lngInternalStdIndexOriginal = mInternalStdIndexPointers(mCurrIDMatches(mCurrIDCnt).IDInd)
                        dblMatchMass = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).MonoisotopicMass
                        dblMatchNET = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).NET
                    Else
                        mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = False
                        
                        lngMassTagIndexPointer = mMTInd(mCurrIDMatches(mCurrIDCnt).IDInd)
                        lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
                        
                        If LastSearchTypeN14N15 = SEARCH_N14 Then
                            ' N14
                            dblMatchMass = mMTMWN14(mCurrIDMatches(mCurrIDCnt).IDInd)
                        Else
                            ' N15
                            dblMatchMass = mMTMWN15(mCurrIDMatches(mCurrIDCnt).IDInd)
                        End If
                        dblMatchNET = AMTData(lngMassTagIndexOriginal).NET
                        
                        ' Unless something wacky is going on, the difference in mass between the UMC and the MT Tag should be less than 1 Da
                        Debug.Assert Abs(GelUMC(CallerID).UMCs(lngUMCIndex).ClassMW - dblMatchMass) < 1
                    End If
                    
                    mCurrIDCnt = mCurrIDCnt + 1
  
  
                    With GelUMC(CallerID).UMCs(lngUMCIndex)
                        ' Compare the mass of this AMT Tag to each member of the UMC
                        For lngMemberIndex = 0 To .ClassCount - 1
                            If SearchUMCTestNET(CInt(.ClassMType(lngMemberIndex)), .ClassMInd(lngMemberIndex), dblMatchNET, NETTolFinal, dblNETDifference) Then
                                
                                Select Case .ClassMType(lngMemberIndex)
                                Case glCSType
                                     dblCurrMW = GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).AverageMW
                                Case glIsoType
                                     dblCurrMW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)), samtDef.MWField)
                                End Select
                                
                                ' See if the member is within mass tolerance
                                If Abs(dblMatchMass - dblCurrMW) <= MWTolAbsFinal Then
                                    ' Yes, within both mass and NET tolerance; increment mCurrIDMatches().MatchingMemberCount
                                    mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount = mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount + 1
                                End If
                             End If
                        Next lngMemberIndex
                    End With

                   
                End If
            End If
        End If
    
        If lngLinesRead Mod 250 = 0 Then
            UpdateStatus "Reading STAC Results: " & lngLinesRead
        End If
        
        lngLinesRead = lngLinesRead + 1
    Loop

    ' Process the last UMC
    LoadSTACResultsUpdateUMC lngUMCIndexSaved
    
    ts.Close
  
    LoadSTACResults = True
    Exit Function

LoadSTACResultsErrorHandler:
    Debug.Assert False
    'Resume
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->LoadSTACResults"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error loading STAC results: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    LoadSTACResults = False
    
End Function

Private Function LoadSTACStats(ByRef fso As FileSystemObject, _
                               ByVal strSTACFDRFilePath As String) As Boolean

    Dim ts As TextStream
    
    Dim lngLinesRead As Long
    
    Dim strLineIn As String
    Dim strMessage As String
    
    Dim strSplitLine() As String
    
    Dim dblCutoff As Double
    
    Dim lngMatches As Long
    Dim dblErrors As Double
    Dim dblFDR As Double

    Dim lngUPFilteredMatches As Long
    Dim dblUPFilteredErrors As Double
    Dim dblUPFilteredFDR As Double
    
    Dim blnValidData As Boolean
    
    Dim lstNewItem As MSComctlLib.ListItem
    
On Error GoTo LoadSTACStatsErrorHandler
    
    UpdateStatus "Reading STAC FDR stats"
    
    lvwSTACStats.ListItems.Clear
    
    STACStatsCount = 0
    ReDim STACStats(50)
    
    Set ts = fso.OpenTextFile(strSTACFDRFilePath, ForReading, False)
    
    Do While Not ts.AtEndOfStream
        strLineIn = ts.ReadLine
        
        If Len(strLineIn) > 0 Then
            strSplitLine = Split(strLineIn, ",")
             
            If UBound(strSplitLine) >= 3 Then
                blnValidData = LoadSTACStatsParseLine(strSplitLine, _
                                                      dblCutoff, _
                                                      lngMatches, _
                                                      dblErrors, _
                                                      dblFDR, _
                                                      lngUPFilteredMatches, _
                                                      dblUPFilteredErrors, _
                                                      dblUPFilteredFDR)
                
                If blnValidData Then
                    With STACStats(STACStatsCount)
                        .STACCuttoff = dblCutoff
                        .Matches = lngMatches
                        .Errors = dblErrors
                        .FDR = dblFDR
                        .UP_Filtered_Matches = lngUPFilteredMatches
                        .UP_Filtered_Errors = dblUPFilteredErrors
                        .UP_Filtered_FDR = dblUPFilteredFDR
                    End With
                    
                    Set lstNewItem = lvwSTACStats.ListItems.add(, , Round(dblCutoff, 2))
                            
                    lstNewItem.SubItems(1) = lngMatches
                    lstNewItem.SubItems(2) = Round(dblErrors, 1)
                    lstNewItem.SubItems(3) = Round(dblFDR, 2)

                    lstNewItem.SubItems(4) = lngUPFilteredMatches
                    lstNewItem.SubItems(5) = Round(dblUPFilteredErrors, 1)
                    lstNewItem.SubItems(6) = Round(dblUPFilteredFDR, 2)

                    STACStatsCount = STACStatsCount + 1
                End If
            End If
        End If
    
        lngLinesRead = lngLinesRead + 1
    Loop
    
    ts.Close
  
    LoadSTACStats = True
    Exit Function

LoadSTACStatsErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->LoadSTACStats"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error loading STAC results: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    LoadSTACStats = False
    
End Function

Private Function LoadSTACResultsParseLine(ByRef strSplitLine() As String, _
                                          ByRef lngMassTagID As Long, _
                                          ByRef lngUMCIndex As Long, _
                                          ByRef dblSTACScore As Double, _
                                          ByRef dblMassError As Double, _
                                          ByRef dblNETError As Double, _
                                          ByRef dblUP As Double) As Boolean
    
    Dim blnValidData As Boolean
    
On Error GoTo LoadSTACResultsParseLineErrorHandler
                
    If IsNumeric(strSplitLine(0)) Then
        ' Parse this line
        ' Column order is:
        ' MassTagID, FeatureID, STAC_Score, Mass_Error, NET_Error, UP
        
        lngMassTagID = CLng(strSplitLine(0))
        lngUMCIndex = CLng(strSplitLine(1))
        dblSTACScore = CDbl(strSplitLine(2))
        dblMassError = CDbl(strSplitLine(3))
        dblNETError = CDbl(strSplitLine(4))
        dblUP = CDbl(strSplitLine(5))           ' Uniqueness probability (like SLiC)
        
        blnValidData = True
        
    Else
        blnValidData = False
    End If
    
    LoadSTACResultsParseLine = blnValidData
    Exit Function
    
LoadSTACResultsParseLineErrorHandler:
    Debug.Assert False
    LoadSTACResultsParseLine = False
        
End Function

Private Function LoadSTACStatsParseLine(ByRef strSplitLine() As String, _
                                        ByRef dblCutoff As Double, _
                                        ByRef lngMatches As Long, _
                                        ByRef dblErrors As Double, _
                                        ByRef dblFDR As Double, _
                                        ByRef lngUPFilteredMatches As Long, _
                                        ByRef dblUPFilteredErrors As Double, _
                                        ByRef dblUPFilteredFDR As Double) As Boolean
    
    Dim blnValidData As Boolean
    
On Error GoTo LoadSTACStatsParseLineErrorHandler
                
    If IsNumeric(strSplitLine(0)) Then
        ' Parse this line
        ' Column order is:
        ' STAC Cutoff, Matches, Errors, FDR (%)
        
        dblCutoff = CDbl(strSplitLine(0))
        lngMatches = CLng(strSplitLine(1))
        dblErrors = CDbl(strSplitLine(2))
        dblFDR = CDbl(strSplitLine(3))
          
        lngUPFilteredMatches = CLng(strSplitLine(4))
        dblUPFilteredErrors = CDbl(strSplitLine(5))
        dblUPFilteredFDR = CDbl(strSplitLine(6))
        
        blnValidData = True
    
    Else
        blnValidData = False
    End If
    
    LoadSTACStatsParseLine = blnValidData
    Exit Function
    
LoadSTACStatsParseLineErrorHandler:
    Debug.Assert False
    LoadSTACStatsParseLine = False
        
End Function

Private Sub LoadSTACResultsUpdateUMC(ByVal lngUMCIndex As Long)

    Dim lngIndex As Long
    Dim lngMassTagIndexPointer As Long

    Dim blnUsingPrecomputedSLiCScores As Boolean
    Dim blnFilterUsingFinalTolerances As Boolean
    
    If mCurrIDCnt > 0 Then
        ' Populate .IDIndexOriginal
        For lngIndex = 0 To mCurrIDCnt - 1
            If mCurrIDMatches(lngIndex).IDIsInternalStd Then
                lngMassTagIndexPointer = mInternalStdIndexPointers(mCurrIDMatches(lngIndex).IDInd)
                mCurrIDMatches(lngIndex).IDIndexOriginal = lngMassTagIndexPointer
            Else
                lngMassTagIndexPointer = mMTInd(mCurrIDMatches(lngIndex).IDInd)
                mCurrIDMatches(lngIndex).IDIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            End If
        Next lngIndex
        
        blnUsingPrecomputedSLiCScores = True
        blnFilterUsingFinalTolerances = False
        
        ' Compute the DelSLiCScores using the already-loaded UP scores
        SearchAMTComputeSLiCScores mCurrIDCnt, mCurrIDMatches, 0, 0, 0, srsSearchRegionShapeConstants.srsRectangular, blnUsingPrecomputedSLiCScores, blnFilterUsingFinalTolerances

        If mCurrIDCnt > 0 Then
            Call AddCurrIDsToAllIDs(lngUMCIndex)
        End If
    End If
                        
End Sub

Private Function ManageCurrID(ByVal ManageType As Long) As Boolean
On Error GoTo exit_ManageCurrID
Select Case ManageType
Case MNG_ERASE
     mCurrIDCnt = 0
     Erase mCurrIDMatches
Case MNG_TRIM
     If mCurrIDCnt > 0 Then
        ReDim Preserve mCurrIDMatches(mCurrIDCnt - 1)
     Else
        ManageCurrID = ManageCurrID(MNG_ERASE)
     End If
Case MNG_RESET
     mCurrIDCnt = 0
     ReDim mCurrIDMatches(MNG_START_SIZE)
Case MNG_ADD_START_SIZE
     ReDim Preserve mCurrIDMatches(mCurrIDCnt + MNG_START_SIZE)
Case Else
     If ManageType > 0 Then
        ReDim Preserve mCurrIDMatches(mCurrIDCnt + ManageType)
     End If
End Select
ManageCurrID = True
exit_ManageCurrID:
End Function

Private Function PerformSearch(ByVal eSearchMode As eSearchModeConstants) As Long
    ' Returns the number of hits
    Dim strMessage As String
    Dim blnCustomNETsAreAvailable As Boolean
    Dim strSearchModeDescription As String
    Dim eResponse As VbMsgBoxResult
    
    Dim blnWarnSearchTol As Boolean
    Dim strSearchTolText As String
    
On Error GoTo PerformSearchErrorHandler
    
    If Me.UseSTAC Then
    
        Select Case samtDef.TolType
        Case gltPPM
             strSearchTolText = samtDef.MWTol & " ppm"
        Case gltABS
             strSearchTolText = samtDef.MWTol & " Da"
        Case Else
            Debug.Assert False
            strSearchTolText = "??"
        End Select
    
        blnWarnSearchTol = False
        If samtDef.TolType = gltPPM Then
            If samtDef.MWTol < DEFAULT_MW_TOL Then blnWarnSearchTol = True
        Else
            If MassToPPM(samtDef.MWTol, 1000) < DEFAULT_MW_TOL Then blnWarnSearchTol = True
        End If
            
        If blnWarnSearchTol Then
            strMessage = "Warning: Mass tolerance of " & strSearchTolText & " is less than the suggested minimum when using STAC (" & DEFAULT_MW_TOL & " ppm)."
            
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            Else
                eResponse = MsgBox("You are strongly encouraged to use a mass tolerance of " & DEFAULT_MW_TOL & " ppm or more when using STAC.  Auto-update it now?", vbQuestion + vbYesNoCancel, glFGTU)
                If eResponse = vbCancel Then
                    PerformSearch = 0
                    Exit Function
                ElseIf eResponse = vbYes Then
                    
                    samtDef.TolType = gltPPM
                    SetTolType samtDef.TolType
                    
                    txtMWTol.Text = DEFAULT_MW_TOL
                    samtDef.MWTol = DEFAULT_MW_TOL
                Else
                    AddToAnalysisHistory CallerID, strMessage
                End If
            End If
        End If
    
        If samtDef.NETTol < DEFAULT_NET_TOL Then
            strMessage = "Warning: NET tolerance of " & samtDef.NETTol & " is less than the suggested minimum when using STAC (" & DEFAULT_NET_TOL & ")."
            
            eResponse = MsgBox("You are strongly encouraged to use a NET tolerance of " & DEFAULT_NET_TOL & " or more when using STAC.  Auto-update it now?", vbQuestion + vbYesNoCancel, glFGTU)
            If eResponse = vbCancel Then
                PerformSearch = 0
                Exit Function
            ElseIf eResponse = vbYes Then
                txtNETTol.Text = DEFAULT_NET_TOL
                samtDef.NETTol = DEFAULT_NET_TOL
            Else
                AddToAnalysisHistory CallerID, strMessage
            End If
        End If
    
    End If
    
    mKeyPressAbortProcess = 0
    cmdSearchAllUMCs.Visible = False
    cmdRemoveAMTMatchesFromUMCs.Visible = False
    DoEvents
    
    If mMatchStatsCount > 0 Then    'something already identified
       Call DestroyIDStructures
    End If

    ' Initialize mUMCMatchStats
    ReDim mUMCMatchStats(100)
    
    With mSearchSummaryStats
        .UMCCountWithHits = 0
        .UniqueMTCount = 0
        .UniqueIntStdCount = 0
        .UniqueMTCount1PctFDR = 0
        .UniqueMTCount5PctFDR = 0
        .UniqueMTCount10PctFDR = 0
        .UniqueMTCount25PctFDR = 0
        .UniqueMTCount50PctFDR = 0
        .MassToleranceFromSTAC = 0
        .NETToleranceFromSTAC = 0
    End With

    Select Case eSearchMode
        Case eSearchModeAll
            SearchType = SEARCH_ALL
            strSearchModeDescription = "all"

        Case eSearchModePaired
            SearchType = SEARCH_PAIRED
            strSearchModeDescription = "paired"

        Case eSearchModePairedPlusUnpaired
            SearchType = SEARCH_PAIRED_PLUS_NON_PAIRED
            strSearchModeDescription = "light pairs plus non-paired"

        Case eSearchModeNonPaired
            SearchType = SEARCH_NON_PAIRED
            strSearchModeDescription = "non-paired"

        Case Else
            ' Unknown search mode
            LogErrors 0, "ExportMTDBbyUMCToUMCResultsTable", "Unknown value for eSearchMode: " & eSearchMode
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "Unknown value for eSearchMode: " & eSearchMode, vbExclamation + vbOKOnly, glFGTU
            End If

            cmdSearchAllUMCs.Visible = True
            cmdRemoveAMTMatchesFromUMCs.Visible = True
            PerformSearch = 0
            
            Exit Function
    End Select

    samtDef.SearchScope = glScope.glSc_All
    mSearchRegionShape = cboSearchRegionShape.ListIndex
    
    ' Note: PrepareMTArrays will update mSearchUsedCustomNETs based on .CustomNETsDefined
    blnCustomNETsAreAvailable = GelData(CallerID).CustomNETsDefined
    If cChkBox(chkDisableCustomNETs) Then
        GelData(CallerID).CustomNETsDefined = False
    End If
    
    CheckNETEquationStatus
    eInternalStdSearchMode = cboInternalStdSearchMode.ListIndex
    
    Select Case glbPreferencesExpanded.AMTSearchResultsBehavior
    Case asrbAutoRemoveExisting
        RemoveAMTMatchesFromUMCs False
        samtDef.SkipReferenced = False
    Case asrbKeepExisting
        samtDef.SkipReferenced = False
    Case asrbKeepExistingAndSkip
        samtDef.SkipReferenced = True
    Case Else
        Debug.Assert False
        RemoveAMTMatchesFromUMCs False
        samtDef.SkipReferenced = False
    End Select
    
    If PrepareMTArrays() Then
        mUMCCountSkippedSinceRefPresent = 0
        txtUniqueMatchStats.Text = ""
        txtSTACMatchStats.Text = ""

        STACStatsCount = 0
        lvwSTACStats.ListItems.Clear
        ctlSTACStats.ClearData
    
        If Me.UseSTAC Then
            ' Search the UMCs using STAC
            GelData(CallerID).MostRecentSearchUsedSTAC = True
            SearchUMCsUsingSTAC eSearchMode
        Else
            ' Search the UMCs using VIPER
            GelData(CallerID).MostRecentSearchUsedSTAC = False
            SearchUMCsUsingVIPER eSearchMode
        End If

        LastSearchTypeN14N15 = N14N15

        Select Case eSearchMode
            Case eSearchModePaired, eSearchModePairedPlusUnpaired
                If GelAnalysis(CallerID).MD_Type = stNotDefined Or GelAnalysis(CallerID).MD_Type = stStandardIndividual Then
                    ' Only update MD_Type if it is currently stStandardIndividual
                    GelAnalysis(CallerID).MD_Type = stPairsO16O18
                End If

            Case Else
                ' Includes eSearchModeAll and eSearchModeNonPaired

                With GelSearchDef(CallerID).AMTSearchMassMods
                    If .PEO Then
                        GelAnalysis(CallerID).MD_Type = stLabeledPEO
                    ElseIf .ICATd0 Then
                        GelAnalysis(CallerID).MD_Type = stLabeledICATD0
                    ElseIf .ICATd8 Then
                        GelAnalysis(CallerID).MD_Type = stLabeledICATD8
                    Else
                        GelAnalysis(CallerID).MD_Type = stStandardIndividual
                    End If
                End With
        End Select
    
        If mKeyPressAbortProcess <= 1 Then

            strMessage = DisplayHitSummary(strSearchModeDescription)
            
            If Me.UpdateGelDataWithSearchResults Then
                ' Store the search results in the gel data
                If mMatchStatsCount > 0 Then RecordSearchResultsInData
                UpdateStatus strMessage
            End If
        Else
            UpdateStatus "Search aborted."
        End If
    Else
       UpdateStatus "Error searching for matches"
    End If
    
    AutoSizeForm
    
    GelData(CallerID).CustomNETsDefined = blnCustomNETsAreAvailable
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True
    DoEvents
    
    PerformSearch = mMatchStatsCount
    Exit Function

PerformSearchErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->PerformSearch"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured during the search: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True
    PerformSearch = 0

End Function

Private Sub PickParameters()
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
Call txtAlkylationMWCorrection_LostFocus
Call txtNETFormula_LostFocus
End Sub

Private Sub PopulateComboBoxes()
    Dim intIndex As Integer
    
On Error GoTo PopulateComboBoxesErrorHandler

    With cboResidueToModify
        .Clear
        .AddItem "Full MT"
        For intIndex = 0 To 25
            .AddItem Chr(vbKeyA + intIndex)
        Next intIndex
        .AddItem glPHOSPHORYLATION
        .ListIndex = 0
    End With
    
    With cboInternalStdSearchMode
        .Clear
        .AddItem "Search only MT tags", issmFindOnlyMassTags
        .AddItem "Search MT tags & Int Stds", issmFindWithMassTags
        .AddItem "Search only Internal Stds", issmFindOnlyInternalStandards
        
        If APP_BUILD_DISABLE_MTS Then
            .ListIndex = issmFindOnlyMassTags
        Else
            .ListIndex = issmFindWithMassTags
        End If
    End With
    
    With cboAMTSearchResultsBehavior
        .Clear
        .AddItem "Auto remove existing results prior to search", asrbAMTSearchResultsBehaviorConstants.asrbAutoRemoveExisting
        .AddItem "Keep existing results; do not skip LC-MS Features", asrbAMTSearchResultsBehaviorConstants.asrbKeepExisting
        .AddItem "Keep existing results; skip LC-MS Features with results", asrbAMTSearchResultsBehaviorConstants.asrbKeepExistingAndSkip
        .ListIndex = asrbAutoRemoveExisting
    End With
    
    With cboSearchRegionShape
        .Clear
        .AddItem "Elliptical search region"
        .AddItem "Rectangular search region"
        .ListIndex = srsSearchRegionShapeConstants.srsElliptical
    End With
    
    Exit Sub
    
PopulateComboBoxesErrorHandler:
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->PopulateComboBoxes"
End Sub

Private Sub PositionControls()
        
    Dim lngNewValue As Long
    
    ctlSTACStats.Height = CheckVsMinimum(Me.ScaleHeight - ctlSTACStats.Top - 75, 2500)
    ctlSTACStats.width = CheckVsMinimum(Me.ScaleWidth - ctlSTACStats.Left - 75, 2500)

    fraSTACPlotOptions.Left = CheckVsMinimum(Me.ScaleWidth - fraSTACPlotOptions.width - 25, 8500)
    
    cmdCopySTACSTats.Left = fraSTACPlotOptions.Left
    cmdZoomOutSTACPlot.Left = fraSTACPlotOptions.Left
    
    lvwSTACStats.width = CheckVsMinimum(fraSTACPlotOptions.Left - lvwSTACStats.Left - 50, 500)
    
End Sub

Private Function PossiblyQuotePath(ByVal strPath As String) As String
    If InStr(strPath, " ") > 0 Then
        PossiblyQuotePath = """" & strPath & """"
    Else
        PossiblyQuotePath = strPath
    End If
End Function

Private Function PrepareMTArrays() As Boolean
    '---------------------------------------------------------------
    'prepares masses from loaded MT tags based on specified
    'modifications; returns True if succesful, False on any error
    '---------------------------------------------------------------
    Dim i As Long, j As Long
    Dim TmpCnt As Long
    Dim CysCnt As Long                 'Cysteine count in peptide
    Dim CysLeft As Long                'Cysteine left for modification use
    Dim CysUsedPEO As Long             'Cysteine already used in calculation for PEO
    Dim CysUsedICAT_D0 As Long         'Cysteine already used in calculation for ICAT_D0
    Dim CysUsedICAT_D8 As Long         'Cysteine already used in calculation for ICAT_D8
    
    Dim strResiduesToModify As String   ' One or more residues to modify (single letter amino acid symbols)
    Dim dblResidueModMass As Double
    Dim ResidueOccurrenceCount As Integer
    Dim strResModToken As String
    Dim blnAddMassTag As Boolean
    
    Dim dblNETWobbleDistance As Double
    
    On Error GoTo err_PrepareMTArrays
    
    ' Update GelSearchDef(CallerID).AMTSearchMassMods with the current settings
    With GelSearchDef(CallerID).AMTSearchMassMods
        .PEO = cChkBox(chkPEO)
        .ICATd0 = cChkBox(chkICATLt)
        .ICATd8 = cChkBox(chkICATHv)
        .Alkylation = cChkBox(chkAlkylation)
        .AlkylationMass = CDblSafe(txtAlkylationMWCorrection)
        If cboResidueToModify.ListIndex > 0 Then
            .ResidueToModify = cboResidueToModify
        Else
            .ResidueToModify = ""
        End If
        
        .ResidueMassModification = CDblSafe(txtResidueToModifyMass)
        txtResidueToModifyMass = Round(.ResidueMassModification, 5)
        
        strResiduesToModify = .ResidueToModify
        dblResidueModMass = .ResidueMassModification
        
        .N15InsteadOfN14 = optN(SEARCH_N15).Value
        
        ' Superseded by .ModMode in August 2008
        '.DynamicMods = optDBSearchModType(MODS_DYNAMIC).Value
        
        .ModMode = GetDBSearchModeType()
    End With
    
    ' Check whether the user is using a 4.0085 Da, full peptide modification, which would indicate 16O/18O modification
    If Abs(dblResidueModMass - glO16O18_DELTA) <= 0.01 And Len(strResiduesToModify) = 0 Then
        mMTListContains16O18OMods = True
    Else
        mMTListContains16O18OMods = False
    End If
    
    
    If IsNumeric(txtDBSearchMinimumHighNormalizedScore.Text) Then
        mMTMinimumHighNormalizedScore = CSngSafe(txtDBSearchMinimumHighNormalizedScore.Text)
    Else
        mMTMinimumHighNormalizedScore = 0
    End If
        
    If IsNumeric(txtDBSearchMinimumHighDiscriminantScore.Text) Then
        mMTMinimumHighDiscriminantScore = CSngSafe(txtDBSearchMinimumHighDiscriminantScore.Text)
    Else
        mMTMinimumHighDiscriminantScore = 0
    End If
    
    If IsNumeric(txtDBSearchMinimumPeptideProphetProbability.Text) Then
        mMTMinimumPeptideProphetProbability = CSngSafe(txtDBSearchMinimumPeptideProphetProbability.Text)
    Else
        mMTMinimumPeptideProphetProbability = 0
    End If
    
    If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
        If mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
            ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighDiscriminantScore, also taking into account HighNormalizedScore
            ValidateMTMinimumDiscriminantAndPepProphet AMTData(), 1, AMTCnt, mMTMinimumHighDiscriminantScore, mMTMinimumPeptideProphetProbability, mMTMinimumHighNormalizedScore, 2
        Else
            ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighNormalizedScore
            ValidateMTMinimimumHighNormalizedScore AMTData(), 1, AMTCnt, mMTMinimumHighNormalizedScore, 2
        End If
    End If
    
    ' Record the current state of .CustomNETsDefined
    ' If chkDisableCustomNETs is checked, then this will have temporarily been set to False
    mSearchUsedCustomNETs = GelData(CallerID).CustomNETsDefined
    
    If Not IsNumeric(txtDecoySearchNETWobble.Text) Then
        txtDecoySearchNETWobble.Text = 0.1
    End If
    dblNETWobbleDistance = CSngSafe(txtDecoySearchNETWobble.Text)
    
    If AMTCnt <= 0 Then
        mMTCnt = 0
    Else
       UpdateStatus "Preparing arrays for search..."
       
       If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
            ' Call Rnd() with a negative number before calling Randomize() lngRandomNumberSeed in order to
            '  guarantee that we get the same order of random numbers each time
            Call Rnd(-1)
            
           Randomize NET_WOBBLE_SEED
       End If
             
       'initially reserve space for AMTCnt peptides
       ReDim mMTInd(AMTCnt - 1)
       ReDim mMTOrInd(AMTCnt - 1)
       ReDim mMTMWN14(AMTCnt - 1)
       ReDim mMTMWN15(AMTCnt - 1)
       ReDim mMTNET(AMTCnt - 1)
       ReDim mMTMods(AMTCnt - 1)
       mMTCnt = 0
       For i = 1 To AMTCnt
            If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
                If AMTData(i).HighNormalizedScore >= mMTMinimumHighNormalizedScore And _
                   AMTData(i).HighDiscriminantScore >= mMTMinimumHighDiscriminantScore And _
                   AMTData(i).PeptideProphetProbability >= mMTMinimumPeptideProphetProbability Then
                    blnAddMassTag = True
                Else
                    blnAddMassTag = False
                End If
            Else
                blnAddMassTag = True
            End If
            
            If blnAddMassTag Then
                mMTCnt = mMTCnt + 1
                mMTInd(mMTCnt - 1) = mMTCnt - 1
                mMTOrInd(mMTCnt - 1) = i             'index; not the ID
                mMTMWN14(mMTCnt - 1) = AMTData(i).MW
                mMTMWN15(mMTCnt - 1) = AMTData(i).MW + glN14N15_DELTA * AMTData(i).CNT_N       ' N15 is always fixed
                Select Case samtDef.NETorRT
                Case glAMT_NET
                     mMTNET(mMTCnt - 1) = AMTData(i).NET
                Case glAMT_RT_or_PNET
                     mMTNET(mMTCnt - 1) = AMTData(i).PNET
                End Select
                mMTMods(mMTCnt - 1) = ""
            End If
       Next i
       
       If chkPEO.Value = vbChecked Then         'correct based on cys number for PEO label
          UpdateStatus "Adding PEO labeled peptides..."
          TmpCnt = mMTCnt
          For i = 0 To TmpCnt - 1
              CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
              If CysCnt > 0 Then
                 If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 1 Or _
                    GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                    
                    ' Dynamic Mods
                    For j = 1 To CysCnt
                        mMTCnt = mMTCnt + 1
                        mMTInd(mMTCnt - 1) = mMTCnt - 1
                        mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                        mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * glPEO
                        mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                        If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                            mMTNET(mMTCnt - 1) = GetWobbledNET(mMTNET(i), dblNETWobbleDistance)
                        Else
                            mMTNET(mMTCnt - 1) = mMTNET(i)
                        End If
                        
                        mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_PEO & "/" & j
                    Next j
                 Else
                    ' Static Mods
                    ' Simply update the stats for this MT tag
                    mMTMWN14(i) = mMTMWN14(i) + CysCnt * glPEO
                    mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTMods(i) = mMTMods(i) & " " & MOD_TKN_PEO & "/" & CysCnt
                 End If
              End If
          Next i
       End If
       
       'yeah, yeah I know that same cysteine can not be labeled with PEO and ICAT at the same
       'time but who cares anyway I can fix this here easily
       If chkICATLt.Value = vbChecked Then         'correct based on cys number for ICAT label
          UpdateStatus "Adding D0 ICAT labeled peptides..."
          TmpCnt = mMTCnt
          For i = 0 To TmpCnt - 1
              CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
              CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
              If CysUsedPEO < 0 Then CysUsedPEO = 0
              CysLeft = CysCnt - CysUsedPEO
              If CysLeft > 0 Then
                 If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 1 Or _
                    GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                    
                    ' Dynamic Mods
                    For j = 1 To CysLeft
                        mMTCnt = mMTCnt + 1
                        mMTInd(mMTCnt - 1) = mMTCnt - 1
                        mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                        mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * glICAT_Light
                        mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                        
                        If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                            mMTNET(mMTCnt - 1) = GetWobbledNET(mMTNET(i), dblNETWobbleDistance)
                        Else
                            mMTNET(mMTCnt - 1) = mMTNET(i)
                        End If
                        
                        mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D0 & "/" & j
                    Next j
                 Else
                    ' Static Mods
                    ' Simply update the stats for this MT tag
                    ' However, if use also has ICAT_d0 enabled, we need to duplicate this
                    '  MT tag first
                    If chkICATHv.Value = vbChecked Then
                        mMTCnt = mMTCnt + 1
                        mMTInd(mMTCnt - 1) = mMTCnt - 1
                        mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                        mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + CysLeft * glICAT_Heavy
                        mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                        mMTNET(mMTCnt - 1) = mMTNET(i)
                        mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & CysLeft
                    End If
                    
                    ' Now update this MT tag to have ICAT_d0 on all the cysteines
                    mMTMWN14(i) = mMTMWN14(i) + CysLeft * glICAT_Light
                    mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ICAT_D0 & "/" & CysLeft
                 End If
              End If
          Next i
       End If
       
       If chkICATHv.Value = vbChecked Then         'correct based on cys number for ICAT label
          UpdateStatus "Adding D8 ICAT labeled peptides..."
          TmpCnt = mMTCnt
          For i = 0 To TmpCnt - 1
              CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
              CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
              If CysUsedPEO < 0 Then CysUsedPEO = 0
              CysUsedICAT_D0 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D0)
              If CysUsedICAT_D0 < 0 Then CysUsedICAT_D0 = 0
              CysUsedICAT_D8 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D8)
              If CysUsedICAT_D8 < 0 Then CysUsedICAT_D8 = 0
              CysLeft = CysCnt - CysUsedPEO - CysUsedICAT_D0 - CysUsedICAT_D8
              If CysLeft > 0 Then
                 If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 1 Or _
                    GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                    
                    ' Dynamic Mods
                    For j = 1 To CysLeft
                        mMTCnt = mMTCnt + 1
                        mMTInd(mMTCnt - 1) = mMTCnt - 1
                        mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                        mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * glICAT_Heavy
                        mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                        
                        If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                            mMTNET(mMTCnt - 1) = GetWobbledNET(mMTNET(i), dblNETWobbleDistance)
                        Else
                            mMTNET(mMTCnt - 1) = mMTNET(i)
                        End If
                        
                        mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & j
                    Next j
                 Else
                    If chkICATLt.Value = vbChecked Then
                        ' We shouldn't have reached this code since all of the cysteines should
                        '  have been assigned ICAT_d0 or ICAT_d8
                        Debug.Assert False
                    Else
                        ' Static Mods
                        ' Simply update the stats for this MT tag
                        mMTMWN14(i) = mMTMWN14(i) + CysLeft * glICAT_Heavy
                        mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                        mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ICAT_D8 & "/" & CysLeft
                    End If
                 End If
              End If
          Next i
       End If
       
       If chkAlkylation.Value = vbChecked Then         'correct based on cys number for alkylation label
          UpdateStatus "Adding alkylated peptides..."
          TmpCnt = mMTCnt
          For i = 0 To TmpCnt - 1
              CysCnt = AMTData(mMTOrInd(i)).CNT_Cys
              CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
              If CysUsedPEO < 0 Then CysUsedPEO = 0
              CysUsedICAT_D0 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D0)
              If CysUsedICAT_D0 < 0 Then CysUsedICAT_D0 = 0
              CysUsedICAT_D8 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D8)
              If CysUsedICAT_D8 < 0 Then CysUsedICAT_D8 = 0
              CysLeft = CysCnt - CysUsedPEO - CysUsedICAT_D0 - CysUsedICAT_D8
              If CysLeft > 0 Then
                 If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 1 Or _
                    GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                    
                    ' Dynamic Mods
                    For j = 1 To CysLeft
                        mMTCnt = mMTCnt + 1
                        mMTInd(mMTCnt - 1) = mMTCnt - 1
                        mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                        mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * AlkMWCorrection
                        mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                        
                        If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                            mMTNET(mMTCnt - 1) = GetWobbledNET(mMTNET(i), dblNETWobbleDistance)
                        Else
                            mMTNET(mMTCnt - 1) = mMTNET(i)
                        End If
                        
                        mMTMods(mMTCnt - 1) = mMTMods(i) & " " & MOD_TKN_ALK & "/" & j
                    Next j
                 Else
                    ' Static Mods
                    ' Simply update the stats for this MT tag
                    mMTMWN14(i) = mMTMWN14(i) + CysLeft * AlkMWCorrection
                    mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                    mMTMods(i) = mMTMods(i) & " " & MOD_TKN_ALK & "/" & CysLeft
                 End If
              End If
          Next i
       End If
       
       If dblResidueModMass <> 0 Or GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
          UpdateStatus "Adding modified residue mass peptides..."
          TmpCnt = mMTCnt
          For i = 0 To TmpCnt - 1
                
            If Len(strResiduesToModify) > 0 Then
              ResidueOccurrenceCount = LookupResidueOccurrence(mMTOrInd(i), strResiduesToModify)
              
              If InStr(strResiduesToModify, "C") > 0 Then
                CysUsedPEO = GetTokenValue(mMTMods(i), MOD_TKN_PEO)
                If CysUsedPEO < 0 Then CysUsedPEO = 0
                CysUsedICAT_D0 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D0)
                If CysUsedICAT_D0 < 0 Then CysUsedICAT_D0 = 0
                CysUsedICAT_D8 = GetTokenValue(mMTMods(i), MOD_TKN_ICAT_D8)
                If CysUsedICAT_D8 < 0 Then CysUsedICAT_D8 = 0
                ResidueOccurrenceCount = ResidueOccurrenceCount - CysUsedPEO - CysUsedICAT_D0 - CysUsedICAT_D8
              End If
              strResModToken = MOD_TKN_RES_MOD
            Else
              ' Add dblResidueModMass once to the entire MT tag
              ' Accomplish this by setting ResidueOccurrenceCount to 1
              ResidueOccurrenceCount = 1
              strResModToken = MOD_TKN_MT_MOD
            End If
            
            If ResidueOccurrenceCount > 0 Then
               If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 1 Or _
                  GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                    
                  ' Dynamic Mods
                  For j = 1 To ResidueOccurrenceCount
                      mMTCnt = mMTCnt + 1
                      mMTInd(mMTCnt - 1) = mMTCnt - 1
                      mMTOrInd(mMTCnt - 1) = mMTOrInd(i)
                      mMTMWN14(mMTCnt - 1) = mMTMWN14(i) + j * dblResidueModMass
                      mMTMWN15(mMTCnt - 1) = mMTMWN14(mMTCnt - 1) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                      
                      If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                          mMTNET(mMTCnt - 1) = GetWobbledNET(mMTNET(i), dblNETWobbleDistance)
                      Else
                          mMTNET(mMTCnt - 1) = mMTNET(i)
                      End If
                        
                      mMTMods(mMTCnt - 1) = mMTMods(i) & " " & strResModToken & "/" & strResiduesToModify & j
                  Next j
               Else
                  ' Static Mods
                  ' Simply update the stats for this MT tag
                  mMTMWN14(i) = mMTMWN14(i) + ResidueOccurrenceCount * dblResidueModMass
                  mMTMWN15(i) = mMTMWN14(i) + glN14N15_DELTA * AMTData(mMTOrInd(i)).CNT_N
                  mMTMods(i) = mMTMods(i) & " " & strResModToken & "/" & strResiduesToModify & ResidueOccurrenceCount
               End If
            End If
          Next i
       End If
       
       If mMTCnt > 0 Then
          UpdateStatus "Preparing fast search structures..."
          ReDim Preserve mMTInd(mMTCnt - 1)
          ReDim Preserve mMTOrInd(mMTCnt - 1)
          ReDim Preserve mMTMWN14(mMTCnt - 1)
          ReDim Preserve mMTMWN15(mMTCnt - 1)
          ReDim Preserve mMTNET(mMTCnt - 1)
          ReDim Preserve mMTMods(mMTCnt - 1)
          Select Case N14N15
          Case SEARCH_N14
               If Not PrepareSearchN14() Then
                  Debug.Assert False
                  Call DestroySearchStructures
                  Exit Function
               End If
          Case SEARCH_N15
               If Not PrepareSearchN15() Then
                  Debug.Assert False
                  Call DestroySearchStructures
                  Exit Function
               End If
          End Select
       End If
    End If
    
    If Not PrepareSearchInternalStandards() Then
         Debug.Assert False
         Call DestroySearchStructures
         Exit Function
    End If
    
    PrepareMTArrays = True
    Exit Function
    
err_PrepareMTArrays:
    Select Case Err.Number
    Case 9                      'add space in chunks of 10000
       ReDim Preserve mMTInd(mMTCnt + 10000)
       ReDim Preserve mMTOrInd(mMTCnt + 10000)
       ReDim Preserve mMTMWN14(mMTCnt + 10000)
       ReDim Preserve mMTMWN15(mMTCnt + 10000)
       ReDim Preserve mMTNET(mMTCnt + 10000)
       ReDim Preserve mMTMods(mMTCnt + 10000)
       Resume
    Case Else
       Debug.Assert False
       Call DestroySearchStructures
    End Select
End Function

Private Function PrepareSearchInternalStandards() As Boolean
Dim intIndex As Integer
Dim dblInternalStdMasses() As Double
Dim qsd As New QSDouble
Dim blnSuccess As Boolean

On Error GoTo PrepareSearchInternalStandardsErrorHandler

blnSuccess = False
With UMCInternalStandards
    If .Count > 0 Then
        UpdateStatus "Preparing fast Internal Standard search..."
        ReDim dblInternalStdMasses(.Count - 1)
        ReDim mInternalStdIndexPointers(.Count - 1)
        
        For intIndex = 0 To .Count - 1
            dblInternalStdMasses(intIndex) = .InternalStandards(intIndex).MonoisotopicMass
            mInternalStdIndexPointers(intIndex) = intIndex
        Next intIndex
   
        If qsd.QSAsc(dblInternalStdMasses, mInternalStdIndexPointers) Then
            Set InternalStdFastSearch = New MWUtil
            If InternalStdFastSearch.Fill(dblInternalStdMasses()) Then
                blnSuccess = True
            End If
        End If
    Else
        ReDim mInternalStdIndexPointers(0)
        blnSuccess = True
    End If
End With

PrepareSearchInternalStandards = blnSuccess
Exit Function

PrepareSearchInternalStandardsErrorHandler:
Debug.Assert False
LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.PrepareSearchInternalStandards"
PrepareSearchInternalStandards = False

End Function

Private Function PrepareSearchN14() As Boolean
'---------------------------------------------------------------
'prepare search of N14 peptide (use loaded peptides masses)
'---------------------------------------------------------------
On Error Resume Next
If mMTCnt > 0 Then
   UpdateStatus "Preparing fast N14 search..."
   ' Dim qsd As New QSDouble
   ' Old: If qsd.QSAsc(mMTMWN14(), mMTInd()) Then
   If ShellSortDoubleWithParallelLong(mMTMWN14(), mMTInd(), 0, UBound(mMTMWN14)) Then
      Set MWFastSearch = New MWUtil
      If MWFastSearch.Fill(mMTMWN14()) Then PrepareSearchN14 = True
   End If
End If
End Function

Private Function PrepareSearchN15() As Boolean
'---------------------------------------------------------------
'prepare search of N15 peptide (use number of N to correct mass)
'---------------------------------------------------------------
On Error Resume Next
If mMTCnt > 0 Then
   UpdateStatus "Preparing fast N15 search..."
   ' Dim qsd As New QSDouble
   ' Old: If qsd.QSAsc(mMTMWN15(), mMTInd()) Then
   If ShellSortDoubleWithParallelLong(mMTMWN15(), mMTInd(), 0, UBound(mMTMWN15)) Then
      Set MWFastSearch = New MWUtil
      If MWFastSearch.Fill(mMTMWN15()) Then PrepareSearchN15 = True
   End If
End If
End Function

Private Sub RecordSearchResultsInData()
    ' Step through mUMCMatchStats() and add the ID's for each UMC to all of the members of each UMC
    
    Dim lngIndex As Long, lngMemberIndex As Long
    Dim lngUMCIndexOriginal As Long
    Dim lngMassTagIndexPointer As Long                  'absolute index in mMT... arrays
    Dim lngMassTagIndexOriginal As Long                 'absolute index in AMT... arrays
    Dim lngInternalStdIndexOriginal As Long
    Dim lngIonIndexOriginal As Long
    Dim blnAddRef As Boolean
    Dim lngIonCountUpdated As Long
    
    Dim AMTorInternalStdRef As String
    Dim dblMatchMass As Double, dblMatchNET As Double
    Dim dblCurrMW As Double, AbsMWErr As Double
    Dim dblStacOrSLiC As Double
    Dim dblDelScore As Double
    Dim dblUPScore As Double
    
    Dim CurrScan As Long
     
    'always reinitialize statistics arrays
    InitAMTStat
    
    KeyPressAbortProcess = 0
    
    CheckNETEquationStatus

On Error GoTo RecordSearchResultsInDataErrorHandler

    With GelData(CallerID)
        For lngIndex = 0 To mMatchStatsCount - 1
            If lngIndex Mod 50 = 0 Then
                UpdateStatus "Storing results: " & LongToStringWithCommas(lngIndex) & " / " & LongToStringWithCommas(mMatchStatsCount)
                If KeyPressAbortProcess > 1 Then Exit For
            End If
            
            lngUMCIndexOriginal = mUMCMatchStats(lngIndex).UMCIndex
            
            If mUMCMatchStats(lngIndex).IDIsInternalStd Then
                lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(lngIndex).IDIndex)
                dblMatchMass = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).MonoisotopicMass
                dblMatchNET = UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal).NET
            Else
                lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngIndex).IDIndex)
                lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
                
                If LastSearchTypeN14N15 = SEARCH_N14 Then
                    ' N14
                    dblMatchMass = mMTMWN14(mUMCMatchStats(lngIndex).IDIndex)
                Else
                    ' N15
                    dblMatchMass = mMTMWN15(mUMCMatchStats(lngIndex).IDIndex)
                End If
                dblMatchNET = AMTData(lngMassTagIndexOriginal).NET
            End If
            
            dblStacOrSLiC = mUMCMatchStats(lngIndex).StacOrSLiC
            dblDelScore = mUMCMatchStats(lngIndex).DelScore
            dblUPScore = mUMCMatchStats(lngIndex).UniquenessProbability
            
            ' Record the search results in each of the members of the UMC
            For lngMemberIndex = 0 To GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassCount - 1
                lngIonIndexOriginal = GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMInd(lngMemberIndex)
                blnAddRef = False
                
                Select Case GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMType(lngMemberIndex)
                Case glCSType
                    dblCurrMW = .CSData(lngIonIndexOriginal).AverageMW
                    CurrScan = .CSData(lngIonIndexOriginal).ScanNumber
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = dblCurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select
                    
                    If mUMCMatchStats(lngIndex).IDIsInternalStd Then
                        AMTorInternalStdRef = ConstructInternalStdReference(.CSData(lngIonIndexOriginal).AverageMW, ConvertScanToNET(CLng(.CSData(lngIonIndexOriginal).ScanNumber)), lngInternalStdIndexOriginal, dblStacOrSLiC, dblDelScore, dblUPScore)
                    Else
                        AMTorInternalStdRef = ConstructAMTReference(.CSData(lngIonIndexOriginal).AverageMW, ConvertScanToNET(CLng(.CSData(lngIonIndexOriginal).ScanNumber)), 0, lngMassTagIndexOriginal, dblMatchMass, dblStacOrSLiC, dblDelScore, dblUPScore)
                    End If
                    
                    If Len(.CSData(lngIonIndexOriginal).MTID) = 0 Then
                        blnAddRef = True
                    ElseIf InStr(.CSData(lngIonIndexOriginal).MTID, AMTorInternalStdRef) <= 0 Then
                        blnAddRef = True
                    End If
                    
                    If blnAddRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        ' If this specific data point is not within tolerance, then mark it as "Inherited"
                        If Not IsValidMatch(dblCurrMW, AbsMWErr, CurrScan, dblMatchNET, dblMatchMass) Then
                            AMTorInternalStdRef = Trim(AMTorInternalStdRef)
                            If Right(AMTorInternalStdRef, 1) = glARG_SEP Then
                                AMTorInternalStdRef = Left(AMTorInternalStdRef, Len(AMTorInternalStdRef) - 1)
                            End If
                            AMTorInternalStdRef = AMTorInternalStdRef & AMTMatchInheritedMark
                        End If
                        
                        InsertBefore .CSData(lngIonIndexOriginal).MTID, AMTorInternalStdRef
                    End If
                Case glIsoType
                    dblCurrMW = GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField)
                    CurrScan = .IsoData(lngIonIndexOriginal).ScanNumber
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = dblCurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select

                    If mUMCMatchStats(lngIndex).IDIsInternalStd Then
                        AMTorInternalStdRef = ConstructInternalStdReference(GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField), ConvertScanToNET(.IsoData(lngIonIndexOriginal).ScanNumber), lngInternalStdIndexOriginal, dblStacOrSLiC, dblDelScore, dblUPScore)
                    Else
                        AMTorInternalStdRef = ConstructAMTReference(GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField), ConvertScanToNET(.IsoData(lngIonIndexOriginal).ScanNumber), 0, lngMassTagIndexOriginal, dblMatchMass, dblStacOrSLiC, dblDelScore, dblUPScore)
                    End If
                    
                    ' Only add AMTorInternalStdRef if .MTID does not contain it
                    ' First perform a quick check to see if .MTID is empty
                    ' If it's not empty, then use InStr to see if .MTID contains AMTorInternalStdRef (a relatively slow operation)
                    If Len(.IsoData(lngIonIndexOriginal).MTID) = 0 Then
                        blnAddRef = True
                    ElseIf InStr(.IsoData(lngIonIndexOriginal).MTID, AMTorInternalStdRef) <= 0 Then
                        blnAddRef = True
                    End If
                    
                    If blnAddRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        If Not IsValidMatch(dblCurrMW, AbsMWErr, CurrScan, dblMatchNET, dblMatchMass) Then
                            AMTorInternalStdRef = Trim(AMTorInternalStdRef)
                            If Right(AMTorInternalStdRef, 1) = glARG_SEP Then
                                AMTorInternalStdRef = Left(AMTorInternalStdRef, Len(AMTorInternalStdRef) - 1) & AMTMatchInheritedMark & glARG_SEP
                            Else
                                AMTorInternalStdRef = AMTorInternalStdRef & AMTMatchInheritedMark & glARG_SEP
                            End If
                        End If
                        
                        InsertBefore .IsoData(lngIonIndexOriginal).MTID, AMTorInternalStdRef
                    End If
                End Select
            Next lngMemberIndex
        Next lngIndex
    End With
    
    If KeyPressAbortProcess <= 1 Then
        AddToAnalysisHistory CallerID, "Stored search results in ions; recorded all MT tag hits for each LC-MS Feature in all members of the UMC; total ions updated = " & Trim(lngIonCountUpdated)
    End If
    
    Exit Sub

RecordSearchResultsInDataErrorHandler:
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->RecordSearchResultsInData"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured while storing the search results in the data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
End Sub

Private Sub RemoveAMTMatchesFromUMCs(blnQueryUser As Boolean)

    Dim eResponse As VbMsgBoxResult
    
    If blnQueryUser And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Remove MT tag references from the data in the current gel?", vbQuestion + vbYesNoCancel + vbDefaultButton1)
    Else
        eResponse = vbYes
    End If
    
    If eResponse = vbYes Then
        TraceLog 5, "RemoveAMTMatchesFromUMCs", "Calling RemoveAMT"
        RemoveAMT CallerID, glScope.glSc_All
        If eInternalStdSearchMode <> issmFindOnlyMassTags Or cboInternalStdSearchMode.ListIndex <> issmFindOnlyMassTags Then
            TraceLog 5, "RemoveInternalStdFromUMCs", "Calling RemoveInternalStd"
            RemoveInternalStd CallerID, glScope.glSc_All
        End If
        
        TraceLog 5, "RemoveAMTMatchesFromUMCs", "Setting GelStatus(CallerID).Dirty = True"
        
        GelStatus(CallerID).Dirty = True
        If blnQueryUser Then AddToAnalysisHistory CallerID, "Deleted MT tag search results from ions"
        
        TraceLog 5, "RemoveAMTMatchesFromUMCs", "MT tag references removed"
        UpdateStatus "MT tag references removed."
    End If


End Sub

Private Function RobustNETValuesEnabled() As Boolean
    If GelData(CallerID).CustomNETsDefined And Not cChkBox(chkDisableCustomNETs) Then
        RobustNETValuesEnabled = True
    Else
        RobustNETValuesEnabled = False
    End If
End Function

Public Function SaveSTACPlotToClipboardOrEMF(ByVal strFilePath As String, Optional ByVal blnSaveHQ As Boolean = False) As Boolean
    SaveSTACPlotToClipboardOrEMF = SavePlotToClipboardOrEMF( _
                ctlSTACStats, _
                strFilePath, blnSaveHQ)
End Function

Public Function SaveSTACPlotToFile(ByVal ePicfileType As pftPictureFileTypeConstants) As Boolean
   SaveSTACPlotToFile = SavePlotToFile( _
                ctlSTACStats, _
                ePicfileType)
End Function

Private Function SavePlotToClipboardOrEMF(ByRef ctlPlot As CWGraph, ByVal strFilePath As String, ByVal blnSaveHQ As Boolean) As Boolean
    ' If strFilePath is empty then copies to the clipboard
    ' Otherwise, saves to strFilePath
        
On Error GoTo SaveSTACPlotErrorHandler

    TraceLog 5, "SavePlotToClipboardOrEMF", "Save " & ctlPlot.Name & " to: " & strFilePath

    If Len(strFilePath) > 0 Then
        ' Note: The .ControlImageEx function is only available in Measurement Studio v6.0 after you
        '  download and install the patch from http://digital.ni.com/softlib.nsf/websearch/2AAC97491D073A6C86256EEF005374CE?opendocument&node=132060_US
        ' After updating, the c:\windows\system32\cwui.ocx file should be 2,335,240 bytes with date 7/24/2004 2:20 am
        ' Also, make sure the installer does not install an out-of-date cwui.ocx file in the c:\program files\viper folder
        If blnSaveHQ Then
            SavePicture ctlPlot.ControlImageEx(400, 400), strFilePath
        Else
            SavePicture ctlPlot.ControlImageEx(ctlPlot.width / 15, ctlPlot.Height / 15), strFilePath
        End If
    Else
        Clipboard.Clear
        Clipboard.SetData ctlPlot.ControlImage, vbCFMetafile
    End If

    SavePlotToClipboardOrEMF = True
    Exit Function
    
SaveSTACPlotErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        If Len(strFilePath) > 0 Then
            MsgBox "Error saving residuals plot from " & ctlPlot.Name & " to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
        Else
            MsgBox "Error copying residuals plot from " & ctlPlot.Name & " to clipboard: " & Err.Description, vbExclamation + vbOKOnly, "Error"
        End If
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.SavePlotToClipboardOrEMF"
    End If
 
    SavePlotToClipboardOrEMF = False

End Function

Private Function SavePlotToFile(ByRef ctlPlot As CWGraph, ByVal ePicfileType As pftPictureFileTypeConstants) As Boolean
    Const SAVE_HQ As Boolean = False
    Dim strFilePath As String
    Dim blnSuccess As Boolean
    
On Error GoTo SavePlotToFileErrorHandler

    Select Case ePicfileType
    Case pftPictureFileTypeConstants.pftEMF, pftPictureFileTypeConstants.pftWMF
        ' Saving EMF file
        strFilePath = SelectFile(Me.hwnd, "Save picture as EMF ...", , True, "*.emf", "EMF Files (*.emf)|*.emf|All Files (*.*)|*.*", 1)

        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "emf")
            blnSuccess = SavePlotToClipboardOrEMF(ctlPlot, strFilePath, SAVE_HQ)
        End If
    Case Else
        ' Includes pftPictureFileTypeConstants.pftPNG
        ' Saving PNG file
        strFilePath = SelectFile(Me.hwnd, "Save picture as PNG ...", , True, "*.png", "PNG Files (*.png)|*.png|All Files (*.*)|*.*", 1)
        
        If Len(strFilePath) > 0 Then
            strFilePath = FileExtensionForce(strFilePath, "png")
            blnSuccess = SaveSTACPlotToPNG(strFilePath)
        End If
    End Select

    SavePlotToFile = blnSuccess
  Exit Function
    
SavePlotToFileErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving residuals plot to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.SavePlotToFile"
    End If
    SavePlotToFile = False
End Function

Public Function SaveSTACPlotToPNG(ByVal strFilePath As String) As Boolean
    Dim strEmfFilePath As String, strWorkingFilePath As String
    Dim blnSuccess As Boolean
    Dim lngReturn As Long
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
    
On Error GoTo SaveSTACPlotToPNGErrorHandler

    strFilePath = FileExtensionForce(strFilePath, "png")
    strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
    strEmfFilePath = FileExtensionForce(strWorkingFilePath, "emf")
    
    blnSuccess = SavePlotToClipboardOrEMF(ctlSTACStats, strEmfFilePath, True)
      
    If blnSuccess Then
        lngReturn = ConvertEmfToPng(strEmfFilePath, strWorkingFilePath, ctlSTACStats.width / Screen.TwipsPerPixelX, ctlSTACStats.Height / Screen.TwipsPerPixelY)
        If lngReturn = 0 Then
            blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
        Else
            objRemoteSaveFileHandler.DeleteTempFile
            blnSuccess = False
        End If
    End If
    
    SaveSTACPlotToPNG = blnSuccess
    Exit Function

SaveSTACPlotToPNGErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error saving residuals plot from " & ctlSTACStats.Name & " to " & strFilePath & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
    Else
        Debug.Assert False
        LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.SaveSTACPlotToPNG"
    End If
    SaveSTACPlotToPNG = False
    
End Function



Private Function SearchUMCSingleMass(ByVal ClassInd As Long) As Long
'-----------------------------------------------------------------------------
'returns number of hits found for UMC with index ClassInd; -1 in case of error;
'  -2 if skipped since hit already present
'-----------------------------------------------------------------------------

Dim MWTolAbsBroad As Double     ' MWTol used to compute the MatchScore
Dim NETTolBroad As Double       ' NETTol used to compute the MatchScore

Dim MWTolAbsFinal As Double     ' Final MWErr required
Dim NETTolFinal As Double

Dim dblClassMass As Double

Dim blnProceed As Boolean

Dim lngIndex As Long
Dim lngMassTagIndexPointer As Long

' MassTagHitCount holds the number of matching MT tags, excluding Internal Standards
Dim MassTagHitCount As Long

Dim blnUsingPrecomputedSLiCScores As Boolean
Dim blnFilterUsingFinalTolerances As Boolean

On Error GoTo err_SearchUMCSingleMass

If ManageCurrID(MNG_RESET) Then
    If SearchType = SEARCH_PAIRED Or SearchType = SEARCH_PAIRED_PLUS_NON_PAIRED Then
        Select Case N14N15
        Case SEARCH_N14     'don't search if this class is found only as heavy member
            If eClsPaired(ClassInd) = umcpHeavyUnique Or _
               eClsPaired(ClassInd) = umcpHeavyMultiple Then
                SearchUMCSingleMass = 0
                Exit Function
            End If
        Case SEARCH_N15     'don't search if this class is found only as light member
            If eClsPaired(ClassInd) = umcpLightUnique Or _
               eClsPaired(ClassInd) = umcpLightMultiple Then
                SearchUMCSingleMass = 0
                Exit Function
            End If
        End Select
    End If
    
    ' Define the tolerances
    SearchAMTDefineTolerances CallerID, ClassInd, samtDef, dblClassMass, MWTolAbsBroad, NETTolBroad, MWTolAbsFinal, NETTolFinal
    
    With GelUMC(CallerID)
        blnProceed = True
        If samtDef.SkipReferenced Then
            ' Skip this UMC if one or more of its members have an AMT match defined
            blnProceed = Not IsAMTReferencedByUMC(.UMCs(ClassInd), CallerID)
        End If
    End With
        
    If blnProceed Then
        If eInternalStdSearchMode <> issmFindOnlyInternalStandards Then
            ' Search for the MT tags using broad tolerances
            SearchUMCSingleMassAMT GelUMC(CallerID).UMCs(ClassInd), MWTolAbsBroad, NETTolBroad
        End If
        ' MassTagHitCount holds the number of matching MT tags, excluding Internal Standards
        MassTagHitCount = mCurrIDCnt
    Else
        ' Skipped UMC since already has a match
        MassTagHitCount = -2
    End If
    
    If eInternalStdSearchMode <> issmFindOnlyMassTags Then
        ' Search for Internal Standards using broad tolerances
        SearchUMCSingleMassInternalStd GelUMC(CallerID).UMCs(ClassInd), MWTolAbsBroad, NETTolBroad
    End If
     
    If mCurrIDCnt > 0 Then
        ' Populate .IDIndexOriginal
        For lngIndex = 0 To mCurrIDCnt - 1
            If mCurrIDMatches(lngIndex).IDIsInternalStd Then
                lngMassTagIndexPointer = mInternalStdIndexPointers(mCurrIDMatches(lngIndex).IDInd)
                mCurrIDMatches(lngIndex).IDIndexOriginal = lngMassTagIndexPointer
            Else
                lngMassTagIndexPointer = mMTInd(mCurrIDMatches(lngIndex).IDInd)
                mCurrIDMatches(lngIndex).IDIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            End If
        Next lngIndex
        
        blnUsingPrecomputedSLiCScores = False
        blnFilterUsingFinalTolerances = True
        
        ' Next compute the Match Scores
        SearchAMTComputeSLiCScores mCurrIDCnt, mCurrIDMatches, dblClassMass, MWTolAbsFinal, NETTolFinal, mSearchRegionShape, blnUsingPrecomputedSLiCScores, blnFilterUsingFinalTolerances
        
        If mCurrIDCnt > 0 Then
            Call AddCurrIDsToAllIDs(ClassInd)
        End If
    End If
     
Else
    UpdateStatus "Error managing memory."
    MassTagHitCount = -1
End If

SearchUMCSingleMass = MassTagHitCount

Exit Function

err_SearchUMCSingleMass:
    Debug.Assert False
    SearchUMCSingleMass = -1
End Function

Private Sub SearchUMCSingleMassAMT(ByRef udtTestUMC As udtUMCType, ByVal dblMWTol As Double, ByVal dblNETTol As Double)
    ' Compare this LC-MS Feature's mass, NET, and charge with the MT tags

    Dim FastSearchMatchInd As Long
    Dim MatchInd1 As Long, MatchInd2 As Long
    
    Dim dblMassTagMass As Double
    Dim dblMassTagNET As Double

    ' Only need to call MWFastSearch once, sending it udtTestUMC.ClassMW
    MatchInd1 = 0
    MatchInd2 = -1
    If MWFastSearch.FindIndexRange(udtTestUMC.ClassMW, dblMWTol, MatchInd1, MatchInd2) Then
        If MatchInd1 <= MatchInd2 Then
            ' One or more MT tags is within dblMWTol of the median UMC mass
            
            ' Now test each MT tag with dblMWTol and dblNETTol and record the matches
            For FastSearchMatchInd = MatchInd1 To MatchInd2
                
                dblMassTagMass = MWFastSearch.GetMWByIndex(FastSearchMatchInd)
                dblMassTagNET = mMTNET(mMTInd(FastSearchMatchInd))
                
                SearchUMCSingleMassValidate FastSearchMatchInd, dblMWTol, dblNETTol, udtTestUMC, dblMassTagMass, dblMassTagNET, False
            
            Next FastSearchMatchInd
        End If
    End If

End Sub

Private Sub SearchUMCSingleMassInternalStd(ByRef udtTestUMC As udtUMCType, ByVal dblMWTol As Double, ByVal dblNETTol As Double)
    ' Compare this LC-MS Feature's mass, NET, and charge with the Internal Standard in UMCInternalStandards

    Dim FastSearchMatchInd As Long
    Dim MatchInd1 As Long, MatchInd2 As Long
    Dim udtInternalStd As udtInternalStandardEntryType
    
    If UMCInternalStandards.Count <= 0 Then Exit Sub
    
    ' Only need to call InternalStdFastSearch once, sending it udtTestUMC.ClassMW
    MatchInd1 = 0
    MatchInd2 = -1
    If InternalStdFastSearch.FindIndexRange(udtTestUMC.ClassMW, dblMWTol, MatchInd1, MatchInd2) Then
        If MatchInd1 <= MatchInd2 Then
            ' One or more Internal Standard is within dblMWTol of the median UMC mass
            
            ' Now test each MT tag with dblMWTol and dblNETTol and record the matches
            For FastSearchMatchInd = MatchInd1 To MatchInd2
                
                udtInternalStd = UMCInternalStandards.InternalStandards(mInternalStdIndexPointers(FastSearchMatchInd))
                   
                If SearchUMCTestCharge(udtTestUMC.ClassRepType, udtTestUMC.ClassRepInd, udtInternalStd) Then
                    SearchUMCSingleMassValidate FastSearchMatchInd, dblMWTol, dblNETTol, udtTestUMC, udtInternalStd.MonoisotopicMass, udtInternalStd.NET, True
                End If
                
            Next FastSearchMatchInd
        End If
    End If

End Sub

Private Sub SearchUMCSingleMassValidate(ByVal FastSearchMatchInd As Long, ByVal dblMWTol As Double, ByVal dblNETTol As Double, ByRef udtTestUMC As udtUMCType, ByVal dblMassTagMass As Double, ByVal dblMassTagNET As Double, ByVal blnIsInternalStdMatch As Boolean)
    ' Note: This sub is called by both SearchUMCSingleMassAMT and SearchUMCSingleMassInternalStd
    
    ' Check if the match is within NET and mass tolerance
    ' If it is, increment mCurrIDMatches().MatchingMemberCount
    
    ' Note that since we used udtTestUMC.ClassMW in the call to FindIndexRange(), not all members
    '  of the class will necessarily have a matching mass
    
    ' Additionally, it is possible that the conglomerate class mass will match a MT tag, but none
    ' of the members will match.  An example of this is a UMC with two members, weighing 500.0 and 502.0 Da
    ' The median mass is 501.0 Da.  If the dblMWTol = 0.1, then the median will match a MT tag of 501 Da,
    '  but none of the members will match.  In this case, we'll record the match,
    '  but place a 0 in mCurrIDMatches().MatchingMemberCount

    
    Dim blnFirstMatchFound As Boolean
    Dim lngMemberIndex As Long
    
    Dim dblCurrMW As Double
    Dim dblNETDifference As Double
    
    With udtTestUMC
        ' See if each MassTag is within the NET tolerance of any of the members of the class
        ' Alternatively, if .UseUMCConglomerateNET = True, then use the NET value of the class representative
        
        blnFirstMatchFound = False
        If glbPreferencesExpanded.UseUMCConglomerateNET Then
            If SearchUMCTestNET(.ClassRepType, .ClassRepInd, dblMassTagNET, dblNETTol, dblNETDifference) Then
                
                ' Either: AMT Matches this LC-MS Feature's median mass and Class Rep NET
                ' Or:     Internal Standard Matches this LC-MS Feature's median mass, Class Rep NET, and charge
                ' Thus:   Add to mCurrIDMatches()
                
                If mCurrIDCnt > UBound(mCurrIDMatches) Then ManageCurrID (MNG_ADD_START_SIZE)
                
                mCurrIDMatches(mCurrIDCnt).IDInd = FastSearchMatchInd
                mCurrIDMatches(mCurrIDCnt).MatchingMemberCount = 0
                mCurrIDMatches(mCurrIDCnt).StacOrSLiC = -1             ' Set this to -1 for now
                mCurrIDMatches(mCurrIDCnt).MassErr = .ClassMW - dblMassTagMass
                mCurrIDMatches(mCurrIDCnt).NETErr = dblNETDifference
                mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = blnIsInternalStdMatch
                mCurrIDMatches(mCurrIDCnt).UniquenessProbability = 0

                mCurrIDCnt = mCurrIDCnt + 1
                
                blnFirstMatchFound = True
            End If
        End If
        
        If blnFirstMatchFound Or Not glbPreferencesExpanded.UseUMCConglomerateNET Then
            For lngMemberIndex = 0 To .ClassCount - 1
                If SearchUMCTestNET(CInt(.ClassMType(lngMemberIndex)), .ClassMInd(lngMemberIndex), dblMassTagNET, dblNETTol, dblNETDifference) Then
                    
                    Select Case .ClassMType(lngMemberIndex)
                    Case glCSType
                         dblCurrMW = GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).AverageMW
                    Case glIsoType
                         dblCurrMW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)), samtDef.MWField)
                    End Select
                    
                    If Not blnFirstMatchFound Then
                        ' We haven't had a match for this index yet; add to mCurrIDMatches()
                        
                        If mCurrIDCnt > UBound(mCurrIDMatches) Then ManageCurrID (MNG_ADD_START_SIZE)
                        
                        mCurrIDMatches(mCurrIDCnt).IDInd = FastSearchMatchInd
                        mCurrIDMatches(mCurrIDCnt).MatchingMemberCount = 0
                        mCurrIDMatches(mCurrIDCnt).StacOrSLiC = -1    ' Set this to -1 for now
                        mCurrIDMatches(mCurrIDCnt).MassErr = .ClassMW - dblMassTagMass
                        mCurrIDMatches(mCurrIDCnt).NETErr = dblNETDifference
                        mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = blnIsInternalStdMatch
                        mCurrIDMatches(mCurrIDCnt).UniquenessProbability = 0
                        
                        mCurrIDCnt = mCurrIDCnt + 1
                        
                        blnFirstMatchFound = True
                    End If

                    ' See if the member is within mass tolerance
                    If Abs(dblMassTagMass - dblCurrMW) <= dblMWTol Then
                        ' Yes, within both mass and NET tolerance; increment mCurrIDMatches().MatchingMemberCount
                        mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount = mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount + 1
                    End If
                End If
            Next lngMemberIndex
        End If
    End With

End Sub

Private Function SearchUMCsUsingSTAC(ByVal eSearchMode As eSearchModeConstants) As Boolean
    
    Const DEFAULT_MAXIMUM_PROCESSING_TIME_MINUTES As Single = 60
    
    Const APP_MONITOR_INTERVAL_MSEC As Integer = 100
    Const STATUS_UPDATE_INTERVAL_MSEC As Integer = 500
    
    Dim fso As New FileSystemObject
    Dim blnSuccess As Boolean
    
    Dim strSTACProgramPath As String
    Dim strSTACOuputFolder As String
    Dim strArguments As String
    Dim strCurDirSaved As String
    
    Dim strPPMTol As String
    Dim strMessage As String
    Dim strStatusBase As String
    
    Dim objProgRunner As clsProgRunner
    Dim dtProcessingStartTime As Date
    
    Dim lngIteration As Long
    Dim lngStatusUpdateIterationCount As Long
    
    Dim sngProcessingTimeSeconds As Single
    Dim sngMaxProcessingTimeMinutes As Single
    Dim blnAbortProcessing As Boolean
    
    On Error GoTo SearchUMCsUsingSTACErrorHandler
    
    sngMaxProcessingTimeMinutes = DEFAULT_MAXIMUM_PROCESSING_TIME_MINUTES
    If sngMaxProcessingTimeMinutes < 1 Then sngMaxProcessingTimeMinutes = 1
    If sngMaxProcessingTimeMinutes > 300 Then sngMaxProcessingTimeMinutes = 300
    
    
    ' Write out the AMTs and UMCs that we're searching against
    blnSuccess = SearchUMCsUsingSTACExportData(fso, eSearchMode)
    
    If blnSuccess Then
        ' Make sure the STAC .exe exists
        strSTACProgramPath = fso.BuildPath(App.Path, STAC_APP_NAME)
        If Not fso.FileExists(strSTACProgramPath) Then
            strMessage = STAC_APP_NAME & " app not found, unable to continue: " & vbCrLf & strSTACProgramPath
            If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                AddToAnalysisHistory CallerID, strMessage
            Else
                MsgBox strMessage, vbExclamation + vbOKOnly, "File Not Found"
            End If
            
            SearchUMCsUsingSTAC = False
            Exit Function
        End If
        
        
        ' Create the STAC Output folder
        ' We're including mSTACSessionID in the name in case multiple copies of VIPER are running
        strSTACOuputFolder = fso.BuildPath(mSTACTempFolderPath, "STAC_Output" & mSTACSessionID)
        
        If Not mTempFilesToDelete.Exists(strSTACOuputFolder) Then
            mTempFilesToDelete.add strSTACOuputFolder, TEMP_FOLDER_FLAG
        End If

        If Not fso.FolderExists(strSTACOuputFolder) Then
            fso.CreateFolder strSTACOuputFolder
        End If
        
        ' Construct the command line arguments
        strArguments = ""
        strArguments = strArguments & " -m " & PossiblyQuotePath(mSTACAMTFilePath)
        strArguments = strArguments & " -u " & PossiblyQuotePath(mSTACUMCFilePath)
        strArguments = strArguments & " -odir " & PossiblyQuotePath(strSTACOuputFolder)
        
        If cChkBox(chkSTACUsesPriorProbability) Then
            strArguments = strArguments & " -useP T"
        Else
            strArguments = strArguments & " -useP F"
        End If
        
        
        Select Case samtDef.TolType
        Case gltPPM
            strPPMTol = samtDef.MWTol
        Case gltABS
            ' User specified tolerance in Daltons
            ' Convert to the PPM amount that would apply to a 1000 Da peptide
            strPPMTol = samtDef.MWTol / (1000 / 1000000#)
        Case Else
           Debug.Assert False
           strPPMTol = 15
        End Select
        
        strArguments = strArguments & " -ppm " & strPPMTol
        strArguments = strArguments & " -NET " & samtDef.NETTol
                        
        strMessage = "Calling " & STAC_APP_NAME & " using" & strArguments
        AddToAnalysisHistory CallerID, strMessage
        
        strStatusBase = "Calling " & STAC_APP_NAME & " to search the LC-MS Features"
        UpdateStatus strStatusBase
        
        Set objProgRunner = New clsProgRunner
        dtProcessingStartTime = Now()
        
        If objProgRunner.StartProgram(strSTACProgramPath, strArguments, vbNormalNoFocus) Then
        
            lngStatusUpdateIterationCount = CInt(STATUS_UPDATE_INTERVAL_MSEC / CSng(APP_MONITOR_INTERVAL_MSEC))
            If lngStatusUpdateIterationCount < 1 Then lngStatusUpdateIterationCount = 1
            
            Do While objProgRunner.AppRunning
                Sleep APP_MONITOR_INTERVAL_MSEC
                
                sngProcessingTimeSeconds = (Now - dtProcessingStartTime) * 86400#
                If sngProcessingTimeSeconds / 60# >= sngMaxProcessingTimeMinutes Then
                    blnAbortProcessing = True
                    strMessage = "Peak Matching using STAC aborted because over " & Trim(sngMaxProcessingTimeMinutes) & " minutes has elapsed."
                ElseIf mKeyPressAbortProcess = 2 Then
                    blnAbortProcessing = True
                    strMessage = "Peak Matching using STAC was manually aborted by the user after " & Trim(sngProcessingTimeSeconds) & " seconds of processing."
                End If
                
                If blnAbortProcessing Then
                    objProgRunner.AbortProcessing
                    DoEvents
                    
                    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                       MsgBox strMessage, vbOKOnly, glFGTU
                    Else
                       Debug.Assert False
                       LogErrors Err.Number, "frmSearchMT->SearchUMCsUsingSTAC", strMessage
                       AddToAnalysisHistory CallerID, strMessage
                    End If
                    
                    UpdateStatus strMessage
                    Exit Do
                End If
                
                If lngIteration Mod lngStatusUpdateIterationCount = 0 Then
                    UpdateStatus strStatusBase & ": " & Round(sngProcessingTimeSeconds, 1) & " seconds elapsed"
                End If
                
                DoEvents
                lngIteration = lngIteration + 1
                            
            Loop
           
            blnSuccess = Not blnAbortProcessing
            
            If blnSuccess Then
                ' Load the results
                blnSuccess = SearchUMCsUsingSTACLoadResults(fso, strSTACOuputFolder, eSearchMode)
            End If
        End If

    End If
    
    SearchUMCsUsingSTAC = blnSuccess
    Exit Function

SearchUMCsUsingSTACErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->SearchUMCsUsingSTAC"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured during search using STAC: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    SearchUMCsUsingSTAC = False
    
End Function

Private Function SearchUMCsUsingSTACExportData(ByRef fso As FileSystemObject, _
                                               ByVal eSearchMode As eSearchModeConstants) As Boolean
    ' Write out the AMTs and UMCs that we're searching against
    
    Dim ts As TextStream
    
    Dim i As Long
    Dim dblMass As Double
    Dim dblNET As Double
    
    Dim lngInternalStdID As Long
    Dim sngInternalStdPepProphetProbability As Single
    Dim lngInternalStdNETCount As Integer
    Dim lngScanClassRep As Long

    Dim blnSearchThisUMC As Boolean
                  
    On Error GoTo SearchUMCsUsingSTACExportDataErrorHandler
    
    ' Get the temp folder path
    mSTACTempFolderPath = GetTemporaryDir
    
    ' Generate a Uniquifier in case two copies of VIPER are running at once
    If mSTACSessionID = "" Then
        mSTACSessionID = "_" & CLng(Timer()) & "_" & CLng(Rnd(1) * 100000)
    End If
    
    mSTACAMTFilePath = fso.BuildPath(mSTACTempFolderPath, "STAC_AMT_DB" & mSTACSessionID & ".txt")
    mSTACUMCFilePath = fso.BuildPath(mSTACTempFolderPath, "STAC_UMCs" & mSTACSessionID & ".txt")
    
    If Not mTempFilesToDelete.Exists(mSTACAMTFilePath) Then
        mTempFilesToDelete.add mSTACAMTFilePath, TEMP_FILE_FLAG
    End If
        
    If Not mTempFilesToDelete.Exists(mSTACUMCFilePath) Then
        mTempFilesToDelete.add mSTACUMCFilePath, TEMP_FILE_FLAG
    End If
        
    ' Write out the AMTs in the mMT arrays
    
    Set ts = fso.OpenTextFile(mSTACAMTFilePath, ForWriting, True)

    ' Write the header line
    ts.WriteLine ("Mass_Tag_ID" & vbTab & _
                  "Monoisotopic_Mass" & vbTab & _
                  "Avg_GANET" & vbTab & _
                  "High_Peptide_Prophet_Probability" & vbTab & _
                  "Cnt_GANET")

    If eInternalStdSearchMode <> issmFindOnlyInternalStandards Then
        ' Write out the AMTs
        For i = 0 To mMTCnt - 1
            
            If N14N15 = SEARCH_N15 Then
                ' Write out the N15-based mass
                dblMass = mMTMWN15(mMTInd(i))
            Else
                ' Write out the N14-based mass
                dblMass = mMTMWN14(mMTInd(i))
            End If
    
            
            ' Note that the Mass_Tag_ID column will not have the real mass tag ID values
            ' Instead, it contains the index value in mMTInd()
            ' This is necessary in case we're using dynamic mods
    
            ' Also, depending on the mass tag DB, .PeptideProphetProbability might contain
            ' probability values based on MSGF instead of Peptide Prophet
            With AMTData(mMTOrInd(mMTInd(i)))
                ts.WriteLine (mMTInd(i) & vbTab & _
                              dblMass & vbTab & _
                              mMTNET(mMTInd(i)) & vbTab & _
                              .PeptideProphetProbability & vbTab & _
                              .NETCount)
                
            End With
        Next i
    End If
    
    If eInternalStdSearchMode <> issmFindOnlyMassTags Then
        ' Write out the internal standards
        ' The value written to the Mass_Tag_ID column is mMTCnt plus i
        ' That way, when we read the results, any results with an ID value >= mMTCnt must be internal standards
        
        For i = 0 To UMCInternalStandards.Count - 1
            dblMass = UMCInternalStandards.InternalStandards(mInternalStdIndexPointers(i)).MonoisotopicMass
            dblNET = UMCInternalStandards.InternalStandards(mInternalStdIndexPointers(i)).NET

            ' Using fixed values for probability and Cnt_GANET
            sngInternalStdPepProphetProbability = 0.99
            lngInternalStdNETCount = 100

            ts.WriteLine (CStr(mMTCnt + i) & vbTab & _
                          dblMass & vbTab & _
                          dblNET & vbTab & _
                          sngInternalStdPepProphetProbability & vbTab & _
                          lngInternalStdNETCount)
        Next i
    End If
    
    ts.Close

    
    ' Write out the UMCs
   
    Set ts = fso.OpenTextFile(mSTACUMCFilePath, ForWriting, True)
   
    ' Write out the header line
    ts.WriteLine ("UMCIndex" & vbTab & "NETClassRep" & vbTab & "UMCMonoMW")
    
    For i = 0 To ClsCnt - 1
         
         Select Case eSearchMode
             Case eSearchModeAll, eSearchModePairedPlusUnpaired
                blnSearchThisUMC = True
                
             Case eSearchModePaired
                 If eClsPaired(i) <> umcpNone Then
                     blnSearchThisUMC = True
                 Else
                     blnSearchThisUMC = False
                 End If
                 
             Case eSearchModeNonPaired
                 If eClsPaired(i) = umcpNone Then
                     blnSearchThisUMC = True
                 Else
                     blnSearchThisUMC = False
                 End If
                 
             Case Else
                 blnSearchThisUMC = True
         End Select
         
         If blnSearchThisUMC Then
             If SearchType = SEARCH_PAIRED Or SearchType = SEARCH_PAIRED_PLUS_NON_PAIRED Then
                 Select Case N14N15
                 Case SEARCH_N14     'don't search if this class is found only as heavy member
                     If eClsPaired(i) = umcpHeavyUnique Or _
                        eClsPaired(i) = umcpHeavyMultiple Then
                         blnSearchThisUMC = False
                     End If
                 Case SEARCH_N15     'don't search if this class is found only as light member
                     If eClsPaired(i) = umcpLightUnique Or _
                        eClsPaired(i) = umcpLightMultiple Then
                         blnSearchThisUMC = False
                     End If
                 End Select
             End If
     
             If blnSearchThisUMC Then
                 With GelUMC(CallerID)
                     If samtDef.SkipReferenced Then
                         ' Skip this UMC if one or more of its members have an AMT match defined
                         blnSearchThisUMC = Not IsAMTReferencedByUMC(.UMCs(i), CallerID)
                     End If
                 End With
             End If
     
             If blnSearchThisUMC Then
                                     
                  lngScanClassRep = -1
                  Select Case GelUMC(CallerID).UMCs(i).ClassRepType
                  Case glCSType
                      lngScanClassRep = GelData(CallerID).CSData(GelUMC(CallerID).UMCs(i).ClassRepInd).ScanNumber
                  Case glIsoType
                      lngScanClassRep = GelData(CallerID).IsoData(GelUMC(CallerID).UMCs(i).ClassRepInd).ScanNumber
                  End Select
                  
                  If lngScanClassRep >= 0 Then
                    dblNET = ConvertScanToNET(lngScanClassRep)
                    
                    ' GelUMC(CallerID).UMCs(i).ClassNET will likely be non-zero
                    ' But, if it's not, we could compare it to dblNET
                    If GelUMC(CallerID).UMCs(i).ClassNET <> 0 Then
                         Debug.Assert Math.Abs(GelUMC(CallerID).UMCs(i).ClassNET - dblNET) < 0.01
                    End If
                    
                    
                    ts.WriteLine (i & vbTab & _
                               dblNET & vbTab & _
                               GelUMC(CallerID).UMCs(i).ClassMW)
                  End If
             End If
         End If
    Next i
     
    ts.Close
    
    
    SearchUMCsUsingSTACExportData = True
    Exit Function

SearchUMCsUsingSTACExportDataErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->SearchUMCsUsingSTACExportData"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured while exporting data for STAC to use: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    SearchUMCsUsingSTACExportData = False
    
End Function

Private Function SearchUMCsUsingSTACLoadResults(ByRef fso As FileSystemObject, _
                                                ByVal strSTACOuputFolder As String, _
                                                ByVal eSearchMode As eSearchModeConstants) As Boolean

    Dim ts As TextStream

    Dim strSTACLogFilePath As String
    Dim strSTACDataFilePath As String
    Dim strSTACFDRFilePath As String
    
    Dim strMessage As String
    
    Dim blnLoadResults As Boolean
    Dim blnSuccess As Boolean
    
On Error GoTo SearchUMCsUsingSTACLoadResultsErrorHandler
    
    ' Find the matches for each UMC (data in the STAC result file is already sorted by FeatureID)
    
    strSTACLogFilePath = "STAC_UMCs" & mSTACSessionID & "_Log.txt"
    strSTACDataFilePath = "STAC_UMCs" & mSTACSessionID & "_STAC.csv"
    strSTACFDRFilePath = "STAC_UMCs" & mSTACSessionID & "_FDR.csv"
    
    strSTACLogFilePath = fso.BuildPath(strSTACOuputFolder, strSTACLogFilePath)
    strSTACDataFilePath = fso.BuildPath(strSTACOuputFolder, strSTACDataFilePath)
    strSTACFDRFilePath = fso.BuildPath(strSTACOuputFolder, strSTACFDRFilePath)
    
    If Not fso.FileExists(strSTACDataFilePath) Then
        strMessage = "STAC results file not found: " & strSTACDataFilePath
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox strMessage, vbExclamation + vbOKOnly, "Error"
        End If
        blnLoadResults = False
    Else
        blnLoadResults = True
    End If
    
    ' Look for the STAC log file; load and parse the contents
    blnSuccess = LoadSTACLogFile(fso, strSTACLogFilePath)

    If blnLoadResults Then
        ' Load the STAC peak matching results
        blnSuccess = LoadSTACResults(fso, strSTACDataFilePath)
            
        ' Load the STAC FDR stats
        blnSuccess = LoadSTACStats(fso, strSTACFDRFilePath)
        
        ' Update the STAC FDR Plot
        If blnSuccess Then
            UpdateSTACPlot
            AutoSizeForm
        End If
    
        ' Compute the peptide-level FDR values
        blnSuccess = ComputePeptideLevelSTACFDR()
    Else
        blnSuccess = False
    End If
  
    SearchUMCsUsingSTACLoadResults = blnSuccess
    Exit Function

SearchUMCsUsingSTACLoadResultsErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC->SearchUMCsUsingSTACLoadResults"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured while reading the results from STAC: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    SearchUMCsUsingSTACLoadResults = False
    
End Function

Private Sub SearchUMCsUsingVIPER(ByVal eSearchMode As eSearchModeConstants)
    Dim i As Long
    Dim lngHitCount As Long
    Dim blnSearchThisUMC As Boolean
    
    For i = 0 To ClsCnt - 1
        If i Mod 25 = 0 Then
           UpdateStatus "Searching: " & Trim(i) & " / " & Trim(ClsCnt)
           If mKeyPressAbortProcess > 1 Then Exit For
        End If

        Select Case eSearchMode
            Case eSearchModeAll, eSearchModePairedPlusUnpaired
               blnSearchThisUMC = True
               
            Case eSearchModePaired
                If eClsPaired(i) <> umcpNone Then
                    blnSearchThisUMC = True
                Else
                    blnSearchThisUMC = False
                End If
                
            Case eSearchModeNonPaired
                If eClsPaired(i) = umcpNone Then
                    blnSearchThisUMC = True
                Else
                    blnSearchThisUMC = False
                End If
                
            Case Else
                blnSearchThisUMC = True
        End Select
        
        If blnSearchThisUMC Then
            lngHitCount = SearchUMCSingleMass(i)
            If lngHitCount = -2 Then mUMCCountSkippedSinceRefPresent = mUMCCountSkippedSinceRefPresent + 1
        End If
    Next i
 
End Sub

Private Function SearchUMCTestCharge(eMemberType As glDistType, lngMemberIndex As Long, udtInternalStd As udtInternalStandardEntryType) As Boolean
    ' Make sure at least one of the charges for this Net Adj Locker is present in the UMC

    Dim blnValidHit As Boolean
    Dim intCharge As Integer
    
    Select Case eMemberType
    Case glCSType
        intCharge = GelData(CallerID).CSData(lngMemberIndex).Charge
    Case glIsoType
        intCharge = GelData(CallerID).IsoData(lngMemberIndex).Charge
    End Select

    ' Make sure at least one of the charges for this Net Adj Locker is present in the UMC
    If intCharge >= udtInternalStd.ChargeMinimum And _
       intCharge <= udtInternalStd.ChargeMaximum Then
       ' Valid Charge
       blnValidHit = True
    Else
        blnValidHit = False
    End If
    
    SearchUMCTestCharge = blnValidHit

End Function

Private Function SearchUMCTestNET(eMemberType As glDistType, lngMemberIndex As Long, dblAMTNet As Double, ByVal dblNETTol As Double, ByRef dblNETDifference As Double) As Boolean
    
    Dim lngScan As Long
    Dim blnNETMatch As Boolean
    
    Select Case eMemberType
    Case glCSType
        lngScan = GelData(CallerID).CSData(lngMemberIndex).ScanNumber
    Case glIsoType
        lngScan = GelData(CallerID).IsoData(lngMemberIndex).ScanNumber
    End Select
    
    blnNETMatch = False
    dblNETDifference = ConvertScanToNET(lngScan) - dblAMTNet
    If dblNETTol > 0 Then
        If Abs(dblNETDifference) <= dblNETTol Then
            blnNETMatch = True
        End If
    Else
        ' NETTol = 0; assume a match
        blnNETMatch = True
    End If

    SearchUMCTestNET = blnNETMatch
    
End Function

Public Sub SetAlkylationMWCorrection(ByVal dblMass As Double)
    txtAlkylationMWCorrection = dblMass
    AlkMWCorrection = dblMass
End Sub

Private Sub SetDBSearchModType(ByVal bytModMode As Byte)
    If bytModMode = 2 Then
        optDBSearchModType(MODS_DECOY).Value = True
    ElseIf bytModMode = 1 Then
        optDBSearchModType(MODS_DYNAMIC).Value = True
    Else
        ' Assumed fixed
        optDBSearchModType(MODS_FIXED).Value = True
    End If

    GelSearchDef(CallerID).AMTSearchMassMods.ModMode = GetDBSearchModeType()
    
End Sub

Public Sub SetDBSearchNType(blnUseN15 As Boolean)
    If blnUseN15 Then
        optN(1).Value = True
        N14N15 = SEARCH_N15
    Else
        optN(0).Value = True
        N14N15 = SEARCH_N14
    End If
End Sub

Private Sub SetDefaultOptions(ByVal blnUseToleranceRefinementSettings As Boolean)

    Dim udtSearchDef As SearchAMTDefinition
    SetDefaultSearchAMTDef udtSearchDef, UMCNetAdjDef
    
    Me.UpdateGelDataWithSearchResults = True
    
    If blnUseToleranceRefinementSettings Then
        Me.UseSTAC = False
    
        With udtSearchDef
            .MWTol = DEFAULT_TOLERANCE_REFINEMENT_MW_TOL
            .TolType = DEFAULT_TOLERANCE_REFINEMENT_MW_TOL_TYPE
            .NETTol = DEFAULT_TOLERANCE_REFINEMENT_NET_TOL
        End With
    
    Else
        Me.UseSTAC = True
    
        With udtSearchDef
            .MWTol = DEFAULT_MW_TOL
            .TolType = DEFAULT_TOL_TYPE
            .NETTol = DEFAULT_NET_TOL
        End With
    End If
    
    Me.STACUsesPriorProbability = True
    
    cboAMTSearchResultsBehavior.ListIndex = asrbAMTSearchResultsBehaviorConstants.asrbAutoRemoveExisting
    
    If blnUseToleranceRefinementSettings Then
        cboSearchRegionShape.ListIndex = srsSearchRegionShapeConstants.srsRectangular
    Else
        cboSearchRegionShape.ListIndex = srsSearchRegionShapeConstants.srsElliptical
    End If
    
    If APP_BUILD_DISABLE_MTS Then
        cboInternalStdSearchMode.ListIndex = issmInternalStandardSearchModeConstants.issmFindOnlyMassTags
    Else
        cboInternalStdSearchMode.ListIndex = issmInternalStandardSearchModeConstants.issmFindWithMassTags
    End If
    
    txtDBSearchMinimumHighNormalizedScore.Text = 0
    txtDBSearchMinimumHighDiscriminantScore.Text = 0
    txtDBSearchMinimumPeptideProphetProbability.Text = 0
    
    optNETorRT(udtSearchDef.NETorRT).Value = True
    SetCheckBox chkUseUMCConglomerateNET, True
    SetCheckBox chkDisableCustomNETs, False
    
    SetTolType udtSearchDef.TolType
    txtMWTol.Text = udtSearchDef.MWTol
    
    txtNETTol = udtSearchDef.NETTol
    
    SetCheckBox chkPEO, False
    SetCheckBox chkICATLt, False
    SetCheckBox chkICATHv, False
    SetCheckBox chkAlkylation, False
    txtAlkylationMWCorrection = 57.0215
    
    cboResidueToModify.ListIndex = 0
    txtResidueToModifyMass.Text = 0
    
    optDBSearchModType(MODS_DYNAMIC).Value = True
    optN(0).Value = True
    
    SetETMode etGANET

    PickParameters
    
    SetCheckBox chkPlotUPFilteredFDR, True
    
    SetCheckBox chkSTACPlotXGridlines, True
    SetCheckBox chkSTACPlotY1Gridlines, False
    SetCheckBox chkSTACPlotY2Gridlines, True
    
End Sub

Private Sub SetETMode(eETModeDesired As glETType)
    Dim i As Long
    Dim eETModeToUse As glETType

On Error Resume Next

    If RobustNETValuesEnabled() Then
        lblETType.Caption = "Using Custom NETs"
    Else
        If GelAnalysis(CallerID) Is Nothing Then
            eETModeToUse = etGenericNET
        Else
            eETModeToUse = eETModeDesired
        End If
        
        Select Case eETModeToUse
        Case etGenericNET
            If eETModeDesired <> etGenericNET Then
                txtNETFormula.Text = GelUMCNETAdjDef(CallerID).NETFormula
            Else
                txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
            End If
        Case etTICFitNET
          With GelAnalysis(CallerID)
            If .NET_Slope <> 0 Then
                txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
            Else
                txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
            End If
          End With
          If Err Then
             MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
             Exit Sub
          End If
        Case etGANET
          With GelAnalysis(CallerID)
            If .GANET_Slope <> 0 Then
               txtNETFormula.Text = ConstructNETFormula(.GANET_Slope, .GANET_Intercept)
            Else
               txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
            End If
          End With
          If Err Then
             MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
             Exit Sub
          End If
        End Select
        For i = mnuET.LBound To mnuET.UBound
            If i = eETModeDesired Then
               mnuET(i).Checked = True
               lblETType.Caption = "ET: " & mnuET(i).Caption
            Else
               mnuET(i).Checked = False
            End If
        Next i
        Call txtNETFormula_LostFocus        'make sure expression evaluator is
                                            'initialized for this formula
    End If

End Sub

Public Sub SetInternalStandardSearchMode(eInternalStdSearchMode As issmInternalStandardSearchModeConstants)
    On Error Resume Next
    
    If APP_BUILD_DISABLE_MTS Then
        eInternalStdSearchMode = issmInternalStandardSearchModeConstants.issmFindOnlyMassTags
    End If
    
    cboInternalStdSearchMode.ListIndex = eInternalStdSearchMode
    If cboInternalStdSearchMode.ListIndex < 0 Then cboInternalStdSearchMode.ListIndex = 0
End Sub

Public Sub SetMinimumHighDiscriminantScore(sngMinimumHighDiscriminantScore As Single)
    txtDBSearchMinimumHighDiscriminantScore = sngMinimumHighDiscriminantScore
End Sub

Public Sub SetMinimumHighNormalizedScore(sngMinimumHighNormalizedScore As Single)
    txtDBSearchMinimumHighNormalizedScore = sngMinimumHighNormalizedScore
End Sub

Public Sub SetMinimumPeptideProphetProbability(sngMinimumPeptideProphetProbability As Single)
    txtDBSearchMinimumPeptideProphetProbability = sngMinimumPeptideProphetProbability
End Sub

Private Sub SetTolType(ByVal eTolType As Integer)

    Select Case eTolType
    Case gltPPM
        optTolType(0).Value = True
    Case gltABS
        optTolType(1).Value = True
    Case Else
        Debug.Assert False
        optTolType(0).Value = True
    End Select

End Sub

Private Sub ShowErrorDistribution2DForm()
    
    frmErrorDistribution2DLoadedData.CallerID = CallerID
    frmErrorDistribution2DLoadedData.Show vbModal
    
    ' Make sure the search tolerances displayed match those now in memory (in case the user performed tolerance refinement)
    DisplayCurrentSearchTolerances
End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuFSepExportToDatabase.Visible = blnVisible
    mnuFExportResultsToDBbyUMC.Visible = blnVisible
    mnuFExportDetailedMemberInformation.Visible = blnVisible

    mnuMTLoadMT.Visible = blnVisible
    
    cboInternalStdSearchMode.Visible = blnVisible
    lblInternalStdSearchMode.Visible = blnVisible
    
End Sub

Private Function ShowOrSaveResultsByIon(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True, Optional ByVal blnIncludeORFInfo As Boolean = True) As Long
    '---------------------------------------------------
    'report results, listing by data point (by ion)
    ' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
    ' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
    ' If blnIncludeORFInfo = True, then attempts to connect to the database and retrieve the ORF information for each MT tag
    '
    ' Returns 0 if no error, the error number if an error
    '---------------------------------------------------
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim strBaseMatchInfo As String
    Dim strLineOut As String
    Dim fname As String
    
    ' Note: AMTRefs() is 1-based
    Dim AMTRefs() As String
    Dim AMTRefsCnt As Long
    Dim i As Long
    Dim lngExportCount As Long
    Dim strSepChar As String
    Dim dblIonMass As Double
    
    Dim lngAMTID() As Long      ' AMT ID's, copied from the globaly array AMTData()
    Dim lngIndex As Long
    
    ' The following is used to lookup the mass of each MT tag, given the MT tag ID
    ' It is initialized using AMTData()
    Dim objAMTIDFastSearch As New FastSearchArrayLong
    
    ' Since AMT masses can be modified (e.g. alkylation), we must use the Pointer determined above
    '   to search mMTOrInd() to determine the correct match
    ' We'll use objMTOrIndFastSearch, initializing using mMTOrInd()
    Dim objMTOrIndFastSearch As New FastSearchArrayLong
    
    ' In order to add to the confusion, we must actually lookup the mMTOrInd() value in mMTInd()
    ' This requires a 3rd FastSearch object, initialized using mMTInd()
    Dim objMTIndFastSearch As New FastSearchArrayLong
    
    ' This last FastSearch object is used to lookup an ORF name
    Dim objORFNameFastSearch As New FastSearchArrayLong
    Dim blnSuccess As Boolean
    
    On Error GoTo err_ShowOrSaveResultsByIon
    
    If blnIncludeORFInfo Then
        UpdateStatus "Sorting Protein lookup arrays"
        If MTtoORFMapCount = 0 Then
            blnIncludeORFInfo = InitializeORFInfo(False)
        Else
            ' We can use MTIDMap(), ORFIDMap(), and ORFRefNames() to get the ORF name
            blnSuccess = objORFNameFastSearch.Fill(MTIDMap())
            Debug.Assert blnSuccess
        End If
    End If
    
    Select Case LastSearchTypeN14N15
    Case SEARCH_N14
         NTypeStr = MOD_TKN_N14
    Case SEARCH_N15
         NTypeStr = MOD_TKN_N15
    End Select
    
    UpdateStatus "Sorting MT lookup arrays"
    mKeyPressAbortProcess = 0
    
    ' Construct the MT tag ID lookup arrays
    ' We need to copy the AMT ID's from AMTData() to lngAMTID() since AMTData().ID is a String array that actually simply holds numbers
    If AMTCnt > 0 Then
        ReDim lngAMTID(1 To AMTCnt)
        For lngIndex = 1 To AMTCnt
            lngAMTID(lngIndex) = AMTData(lngIndex).ID
        Next lngIndex
    Else
        ReDim lngAMTID(1 To 1)
    End If
    
    blnSuccess = objAMTIDFastSearch.Fill(lngAMTID())
    Debug.Assert blnSuccess
    
    blnSuccess = objMTOrIndFastSearch.Fill(mMTOrInd())
    Debug.Assert blnSuccess
    
    blnSuccess = objMTIndFastSearch.Fill(mMTInd())
    Debug.Assert blnSuccess
    
    Me.MousePointer = vbHourglass
    
    mKeyPressAbortProcess = 0
    cmdSearchAllUMCs.Visible = False
    cmdRemoveAMTMatchesFromUMCs.Visible = False
    
    'temporary file for results output
    fname = GetTempReportFilePath()
    If Len(strOutputFilePath) > 0 Then fname = strOutputFilePath
    Set ts = fso.OpenTextFile(fname, ForWriting, True)
    
    strSepChar = LookupDefaultSeparationCharacter()
    
    strLineOut = "Index" & strSepChar & "Scan" & strSepChar & "ChargeState" & strSepChar & "MonoMW" & strSepChar & "Abundance" & strSepChar
    strLineOut = strLineOut & "Fit" & strSepChar & "ER" & strSepChar & "LockerID" & strSepChar & "FreqShift" & strSepChar & "MassCorrection" & strSepChar & "MultiMassTagHitCount" & strSepChar
    strLineOut = strLineOut & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagMods" & strSepChar & "Peptide_Warning_PossiblyIncorrect"
    
    If blnIncludeORFInfo Then
        strLineOut = strLineOut & strSepChar & "MultiORFCount" & strSepChar & "ORFName"
    End If
    
    ts.WriteLine strLineOut

    With GelData(CallerID)
      If .CSLines > 0 Then ts.WriteLine "Charge State Data Block"
      For i = 1 To .CSLines
          If i Mod 500 = 0 Then
            UpdateStatus "Preparing results: " & Trim(i) & " / " & Trim(.CSLines)
            If mKeyPressAbortProcess > 1 Then Exit For
          End If
          If Not IsNull(.CSData(i).MTID) Then
             If IsAMTReferenced(.CSData(i).MTID) Then
                AMTRefsCnt = GetAMTRefFromString2(.CSData(i).MTID, AMTRefs())
                If AMTRefsCnt > 0 Then
                'for Charge State standard deviation is used on place of Fit
                    dblIonMass = .CSData(i).AverageMW
                    strBaseMatchInfo = i & strSepChar & .CSData(i).ScanNumber & strSepChar _
                       & .CSData(i).Charge & strSepChar & .CSData(i).AverageMW & strSepChar _
                       & .CSData(i).Abundance & strSepChar & .CSData(i).MassStDev & strSepChar
                    strBaseMatchInfo = strBaseMatchInfo & LookupExpressionRatioValue(CallerID, i, False)
                    If GelLM(CallerID).CSCnt > 0 Then   'we have mass correction
                       strBaseMatchInfo = strBaseMatchInfo & strSepChar & GelLM(CallerID).CSLckID(i) & strSepChar _
                            & GelLM(CallerID).CSFreqShift(i) & strSepChar _
                            & GelLM(CallerID).CSMassCorrection(i)
                    Else
                       strBaseMatchInfo = strBaseMatchInfo & strSepChar & strSepChar & strSepChar
                    End If
                
                    WriteAMTMatchesForIon ts, strBaseMatchInfo, dblIonMass, AMTRefs(), AMTRefsCnt, objAMTIDFastSearch, objMTOrIndFastSearch, objMTIndFastSearch, lngExportCount, blnIncludeORFInfo, objORFNameFastSearch, strSepChar
                End If
             End If
          End If
      Next i
      If .IsoLines > 0 Then ts.WriteLine "Isotopic Data Block"
      For i = 1 To .IsoLines
          If i Mod 500 = 0 Then
            UpdateStatus "Preparing results: " & Trim(i) & " / " & Trim(.IsoLines)
            If mKeyPressAbortProcess > 1 Then Exit For
          End If
          If Not IsNull(.IsoData(i).MTID) Then
             If IsAMTReferenced(.IsoData(i).MTID) Then
                AMTRefsCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTRefs())
                If AMTRefsCnt > 0 Then
                    dblIonMass = .IsoData(i).MonoisotopicMW
                    strBaseMatchInfo = i & strSepChar & .IsoData(i).ScanNumber & strSepChar _
                       & .IsoData(i).Charge & strSepChar & .IsoData(i).MonoisotopicMW & strSepChar _
                       & .IsoData(i).Abundance & strSepChar & .IsoData(i).Fit & strSepChar
                    strBaseMatchInfo = strBaseMatchInfo & LookupExpressionRatioValue(CallerID, i, True)
                    If GelLM(CallerID).IsoCnt > 0 Then
                       strBaseMatchInfo = strBaseMatchInfo & strSepChar & GelLM(CallerID).IsoLckID(i) & strSepChar _
                             & GelLM(CallerID).IsoFreqShift(i) & strSepChar _
                             & GelLM(CallerID).IsoMassCorrection(i)
                    Else
                       strBaseMatchInfo = strBaseMatchInfo & strSepChar & strSepChar & strSepChar
                    End If
                    
                    WriteAMTMatchesForIon ts, strBaseMatchInfo, dblIonMass, AMTRefs(), AMTRefsCnt, objAMTIDFastSearch, objMTOrIndFastSearch, objMTIndFastSearch, lngExportCount, blnIncludeORFInfo, objORFNameFastSearch, strSepChar
                End If
             End If
          End If
      Next i
    End With
    ts.Close
    
    UpdateStatus ""
    
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True
            
    If blnDisplayResults Then
       frmDataInfo.Tag = "EXP"
       frmDataInfo.SourceFilePath = fname
       frmDataInfo.Show vbModal
       frmDataInfo.SourceFilePath = ""
    Else
        ' MonroeMod
        AddToAnalysisHistory CallerID, "Exported " & lngExportCount & " search results to text file: " & fname
    End If
    ShowOrSaveResultsByIon = 0
    
ShowOrSaveResultsCleanup:
    
    Set ts = Nothing
    Set fso = Nothing
    
    Set objAMTIDFastSearch = Nothing
    Set objMTOrIndFastSearch = Nothing
    Set objMTIndFastSearch = Nothing
    Set objORFNameFastSearch = Nothing
    
    Me.MousePointer = vbDefault
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True
    
    Exit Function
    
err_ShowOrSaveResultsByIon:
    Debug.Assert False
    ShowOrSaveResultsByIon = Err.Number
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.ShowOrSaveResultsByIon"
    Resume ShowOrSaveResultsCleanup

End Function

Public Function ShowOrSaveSTACStats(ByVal blnCopyToClipboard As Boolean, _
                                    Optional strOutputFilePath As String = "", _
                                    Optional blnDisplayResults As Boolean = True) As Long
                                    
    '-------------------------------------
    ' Show the STAC Stats
    '
    ' If blnCopyToClipboard = True, then ignores strOutputFilePath and blnDisplayResults
    ' Returns 0 if no error, the error number if an error
    '-------------------------------------
    
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim fname As String
    
    Dim strSepChar As String
    Dim strLineOut As String
    Dim strClipboardText As String
    
    Dim lngIndex As Long
     
On Error GoTo ShowOrSaveSTACStatsErrorHandler
    
    If blnCopyToClipboard Then
        blnDisplayResults = False
    Else
        'temporary file for results output
        fname = GetTempReportFilePath()
        
        If Len(strOutputFilePath) > 0 Then fname = strOutputFilePath
        Set ts = fso.OpenTextFile(fname, ForWriting, True)
    End If
    
    strSepChar = LookupDefaultSeparationCharacter()
    strClipboardText = ""
    
    ' Write the header line
    strLineOut = "STAC Cutoff" & strSepChar & "Matches" & strSepChar & "Errors" & strSepChar & "FDR" & strSepChar
    strLineOut = strLineOut & "Matches, UP>0.5" & strSepChar & "Errors, UP>0.5" & strSepChar & "FDR, UP>0.5"

    If blnCopyToClipboard Then
        strClipboardText = strClipboardText & strLineOut & vbCrLf
    Else
        ts.WriteLine strLineOut
    End If
        
    For lngIndex = 0 To STACStatsCount - 1
       
        With STACStats(lngIndex)
            strLineOut = Round(.STACCuttoff, 2) & strSepChar & .Matches & strSepChar & Round(.Errors, 1) & strSepChar & Format(.FDR / 100#, "0.000%") & strSepChar
            strLineOut = strLineOut & .UP_Filtered_Matches & strSepChar & Round(.UP_Filtered_Errors, 1) & strSepChar & Format(.UP_Filtered_FDR / 100#, "0.000%")
        End With
        
        If blnCopyToClipboard Then
            strClipboardText = strClipboardText & strLineOut & vbCrLf
        Else
            ts.WriteLine strLineOut
        End If
        
    Next lngIndex
    
    If blnCopyToClipboard Then
        On Error Resume Next
        
        Clipboard.Clear
        Clipboard.SetText strClipboardText, vbCFText
    Else
    
        ts.Close
        
        If Len(strOutputFilePath) > 0 Then
            AddToAnalysisHistory CallerID, "Saved STAC Stats to disk: " & strOutputFilePath
        End If
        
        If blnDisplayResults Then
             frmDataInfo.Tag = "STAC_Stats"
             frmDataInfo.SourceFilePath = fname
             frmDataInfo.Show vbModal
             frmDataInfo.SourceFilePath = ""
        End If
         
         
        Set ts = Nothing
        Set fso = Nothing
    End If

Exit Function

ShowOrSaveSTACStatsErrorHandler:
    Debug.Assert False
    
    ShowOrSaveSTACStats = Err.Number
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.ShowOrSaveSTACStats"
    Set fso = Nothing
End Function

Public Function ShowOrSaveResultsByUMC(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True, Optional ByVal blnIncludeORFInfo As Boolean = True) As Long
    '-------------------------------------
    ' Report identified unique mass classes
    ' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
    ' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
    ' If blnIncludeORFInfo = True, then attempts to connect to the database and retrieve the ORF information for each MT tag
    '
    ' Returns 0 if no error, the error number if an error
    '-------------------------------------
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim fname As String
    
    Dim strLineOut As String
    Dim strLineOutMiddle As String
    Dim strLineOutEnd As String
    Dim strLineOutEndAddnl As String
    
    Dim strMinMaxCharges As String
    Dim mgInd As Long
    Dim lngUMCIndexOriginal As Long                     'absolute index of UMC
    Dim lngMassTagIndexPointer As Long                  'absolute index in mMT... arrays
    Dim lngMassTagIndexOriginal As Long                 'absolute index in AMT... arrays
    Dim lngInternalStdIndexOriginal As Long
    Dim strSepChar As String
    
    Dim dblMatchMass As Double
    Dim dblMatchNET As Double
    Dim strMatchID As String
    Dim strInternalStdDescription As String
    Dim strPeptideSequence As String
    Dim sngPeptideProphetProbability As Single
    
    Dim dblMassErrorPPM As Double
    Dim lngScanClassRep As Long
    Dim lngScanIndex As Long
    
    Dim dblGANETClassRep As Double
    Dim dblGANETError As Double
    Dim objORFNameFastSearch As New FastSearchArrayLong
    Dim blnSuccess As Boolean
    
    Dim lngPairIndex As Long
    Dim blnPairsPresent As Boolean
    Dim blnCorrectedIReportEREnabled As Boolean
        
    Dim objP1IndFastSearch As FastSearchArrayLong
    Dim objP2IndFastSearch As FastSearchArrayLong
    Dim lngPairMatchCount As Long, lngPairMatchIndex As Long
    Dim udtPairMatchStats() As udtPairMatchStatsType
    
    Dim blnReturnAllPairInstances As Boolean
    Dim blnFavorHeavy As Boolean
    
    Dim lngPeakFPRType As Long
    
On Error GoTo ShowOrSaveResultsByUMCErrorHandler
    
    If blnIncludeORFInfo Then
        UpdateStatus "Sorting Protein lookup arrays"
        If MTtoORFMapCount = 0 Then
            blnIncludeORFInfo = InitializeORFInfo(False)
        Else
            ' We can use MTIDMap(), ORFIDMap(), and ORFRefNames() to get the ORF name
            blnSuccess = objORFNameFastSearch.Fill(MTIDMap())
            Debug.Assert blnSuccess
        End If
    End If
    
    UpdateStatus "Preparing results: 0 / " & Trim(mMatchStatsCount)
    
    Me.MousePointer = vbHourglass
    
    mKeyPressAbortProcess = 0
    cmdSearchAllUMCs.Visible = False
    cmdRemoveAMTMatchesFromUMCs.Visible = False
    
    'temporary file for results output
    fname = GetTempReportFilePath()
    
    If Len(strOutputFilePath) > 0 Then fname = strOutputFilePath
    Set ts = fso.OpenTextFile(fname, ForWriting, True)
    
    Select Case LastSearchTypeN14N15
    Case SEARCH_N14
         NTypeStr = MOD_TKN_N14
    Case SEARCH_N15
         NTypeStr = MOD_TKN_N15
    End Select
    
    ' Initialize the PairIndex lookup objects
    blnPairsPresent = PairIndexLookupInitialize(CallerID, objP1IndFastSearch, objP2IndFastSearch)
    
    strSepChar = LookupDefaultSeparationCharacter()
    
    ' UMCIndex; ScanStart; ScanEnd; ScanClassRep; GANETClassRep; UMCMonoMW; UMCMWStDev; UMCMWMin; UMCMWMax; UMCAbundance; ClassStatsChargeBasis; ChargeStateMin; ChargeStateMax; UMCMZForChargeBasis; UMCMemberCount; UMCMemberCountUsedForAbu; UMCAverageFit; PairIndex; PairMemberType; ExpressionRatio; MultiMassTagHitCount; MassTagID; MassTagMonoMW; MassTagMods; MemberCountMatchingMassTag; MassErrorPPM; GANETError; SLiC_Score; Del_SLiC; Uniqueness_Probability; FDR_Threshold, IsInternalStdMatch; PeptideProphetProbability; TIC_from_Raw_Data; Deisotoping_Peak_Count
    strLineOut = "UMCIndex" & strSepChar & "ScanStart" & strSepChar & "ScanEnd" & strSepChar & "ScanClassRep" & strSepChar & "NETClassRep" & strSepChar & "UMCMonoMW" & strSepChar & "UMCMWStDev" & strSepChar & "UMCMWMin" & strSepChar & "UMCMWMax" & strSepChar & "UMCAbundance" & strSepChar
    strLineOut = strLineOut & "ClassStatsChargeBasis" & strSepChar & "ChargeStateMin" & strSepChar & "ChargeStateMax" & strSepChar & "UMCMZForChargeBasis" & strSepChar & "UMCMemberCount" & strSepChar & "UMCMemberCountUsedForAbu" & strSepChar & "UMCAverageFit" & strSepChar & "PairIndex" & strSepChar & "PairMemberType" & strSepChar
    strLineOut = strLineOut & "ExpressionRatio" & strSepChar & "ExpressionRatioStDev" & strSepChar & "ExpressionRatioChargeStateBasisCount" & strSepChar & "ExpressionRatioMemberBasisCount" & strSepChar
    strLineOut = strLineOut & "MultiMassTagHitCount" & strSepChar
    strLineOut = strLineOut & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagMods" & strSepChar & "MemberCountMatchingMassTag" & strSepChar & "MassErrorPPM" & strSepChar & "NETError" & strSepChar
    
    If GelData(CallerID).MostRecentSearchUsedSTAC Then
        strLineOut = strLineOut & "STAC_Score" & strSepChar & "Del_STAC" & strSepChar & "Uniqueness Probability" & strSepChar & "FDR Threshold" & strSepChar
    Else
        strLineOut = strLineOut & "SLiC_Score" & strSepChar & "Del_SLiC" & strSepChar
    End If
         
    strLineOut = strLineOut & "IsInternalStdMatch" & strSepChar & "PeptideProphetProbability" & strSepChar & "Peptide" & strSepChar
    
    strLineOut = strLineOut & "TIC_from_Raw_Data" & strSepChar & "Deisotoping_Peak_Count"
    
    With GelP_D_L(CallerID)
        If blnPairsPresent And .SearchDef.IReportEROptions.Enabled And .SearchDef.ComputeERScanByScan Then
            strLineOut = strLineOut & strSepChar & "Labelling Efficiency F" & strSepChar & "Log2(ER) Corrected for F" & strSepChar & "Log2(ER) Corrected Standard Error"
            blnCorrectedIReportEREnabled = True
        End If
    End With
    
    If blnIncludeORFInfo Then strLineOut = strLineOut & strSepChar & "MultiORFCount" & strSepChar & "ORFName"
    
    ts.WriteLine strLineOut

    For mgInd = 0 To mMatchStatsCount - 1
        lngUMCIndexOriginal = mUMCMatchStats(mgInd).UMCIndex
        
        If mUMCMatchStats(mgInd).IDIsInternalStd Then
            lngInternalStdIndexOriginal = mInternalStdIndexPointers(mUMCMatchStats(mgInd).IDIndex)
            With UMCInternalStandards.InternalStandards(lngInternalStdIndexOriginal)
                dblMatchMass = .MonoisotopicMass
                dblMatchNET = .NET
                strMatchID = .SeqID
                strInternalStdDescription = .Description
                strPeptideSequence = .PeptideSequence
            End With
            sngPeptideProphetProbability = 0
        Else
            lngMassTagIndexPointer = mMTInd(mUMCMatchStats(mgInd).IDIndex)
            lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
            
            If LastSearchTypeN14N15 = SEARCH_N14 Then
                ' N14
                dblMatchMass = mMTMWN14(mUMCMatchStats(mgInd).IDIndex)
            Else
                ' N15
                dblMatchMass = mMTMWN15(mUMCMatchStats(mgInd).IDIndex)
            End If
        
            dblMatchNET = AMTData(lngMassTagIndexOriginal).NET
            ' Future: dblMatchNETStDev = AMTData(lngMassTagIndexOriginal).NETStDev
            strMatchID = Trim(AMTData(lngMassTagIndexOriginal).ID)
            
            sngPeptideProphetProbability = AMTData(lngMassTagIndexOriginal).PeptideProphetProbability
            strPeptideSequence = AMTData(lngMassTagIndexOriginal).Sequence
        End If
    
        GetUMCClassRepScanAndNET CallerID, lngUMCIndexOriginal, lngScanClassRep, dblGANETClassRep
        
        With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
            strLineOut = lngUMCIndexOriginal & strSepChar & .MinScan & strSepChar & .MaxScan & strSepChar & lngScanClassRep & strSepChar & Format(dblGANETClassRep, "0.0000") & strSepChar & Round(.ClassMW, 6) & strSepChar
            strLineOut = strLineOut & Round(.ClassMWStD, 6) & strSepChar & .MinMW & strSepChar & .MaxMW & strSepChar & .ClassAbundance & strSepChar
            
            strMinMaxCharges = ClsStat(lngUMCIndexOriginal, ustChargeMin) & strSepChar & ClsStat(lngUMCIndexOriginal, ustChargeMax) & strSepChar
            
            ' Record ClassStatsChargeBasis, ChargeMin, ChargeMax, UMCMZForChargeBasis, UMCMemberCount, and UMCMemberCountUsedForAbu
            If GelUMC(CallerID).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge & strSepChar
                strLineOut = strLineOut & strMinMaxCharges
                strLineOut = strLineOut & Round(MonoMassToMZ(.ClassMW, .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge), 6) & strSepChar
            Else
                strLineOut = strLineOut & 0 & strSepChar
                strLineOut = strLineOut & strMinMaxCharges
                strLineOut = strLineOut & Round(MonoMassToMZ(.ClassMW, GelData(CallerID).IsoData(.ClassRepInd).Charge), 6) & strSepChar
            End If
            
            strLineOut = strLineOut & .ClassCount & strSepChar
            
            ' Record UMCMemberCountUsedForAbu
            If GelUMC(CallerID).def.UMCClassStatsUseStatsFromMostAbuChargeState Then
                strLineOut = strLineOut & .ChargeStateBasedStats(.ChargeStateStatsRepInd).Count & strSepChar
            Else
                strLineOut = strLineOut & .ClassCount & strSepChar
            End If
            
        End With
        
        strLineOut = strLineOut & Round(ClsStat(lngUMCIndexOriginal, ustFitAverage), 3) & strSepChar
        
        ' Now start populating strLineOutEnd
        strLineOutEnd = ""
        
        If mUMCMatchStats(mgInd).IDIsInternalStd Then
            strLineOutEnd = strLineOutEnd & "0" & strSepChar
        Else
            strLineOutEnd = strLineOutEnd & mUMCMatchStats(mgInd).MultiAMTHitCount & strSepChar
        End If
    
        With GelUMC(CallerID).UMCs(lngUMCIndexOriginal)
            dblMassErrorPPM = MassToPPM(.ClassMW - dblMatchMass, .ClassMW)
            dblGANETError = dblGANETClassRep - dblMatchNET
        End With
        
        strLineOutEnd = strLineOutEnd & strMatchID & strSepChar & Round(dblMatchMass, 6) & strSepChar
        
        If Not mUMCMatchStats(mgInd).IDIsInternalStd Then
            strLineOutEnd = strLineOutEnd & NTypeStr
            If Len(mMTMods(lngMassTagIndexPointer)) > 0 Then
                strLineOutEnd = strLineOutEnd & " " & mMTMods(lngMassTagIndexPointer)
            End If
        End If
        
        strLineOutEnd = strLineOutEnd & strSepChar & mUMCMatchStats(mgInd).MemberHitCount & strSepChar & Round(dblMassErrorPPM, 4) & strSepChar & Round(dblGANETError, NET_PRECISION)
        strLineOutEnd = strLineOutEnd & strSepChar & Round(mUMCMatchStats(mgInd).StacOrSLiC, 4)
        strLineOutEnd = strLineOutEnd & strSepChar & Round(mUMCMatchStats(mgInd).DelScore, 4)
        
        If GelData(CallerID).MostRecentSearchUsedSTAC Then
            strLineOutEnd = strLineOutEnd & strSepChar & Round(mUMCMatchStats(mgInd).UniquenessProbability, 4)
            strLineOutEnd = strLineOutEnd & strSepChar & Format(mUMCMatchStats(mgInd).FDRThreshold, "0.000%")
        End If

        strLineOutEnd = strLineOutEnd & strSepChar & mUMCMatchStats(mgInd).IDIsInternalStd
        strLineOutEnd = strLineOutEnd & strSepChar & Round(sngPeptideProphetProbability, 5)
        strLineOutEnd = strLineOutEnd & strSepChar & strPeptideSequence
        
        lngScanIndex = LookupScanNumberRelativeIndex(CallerID, lngScanClassRep)
        If lngScanIndex = 0 Then
            lngScanClassRep = LookupScanNumberClosest(CallerID, lngScanClassRep)
            lngScanIndex = LookupScanNumberRelativeIndex(CallerID, lngScanClassRep)
        End If
        
        With GelData(CallerID).ScanInfo(lngScanIndex)
            strLineOutEnd = strLineOutEnd & strSepChar & Round(.TIC, 1)
            strLineOutEnd = strLineOutEnd & strSepChar & .NumDeisotoped
        End With
        
        lngPairIndex = -1
        lngPairMatchCount = 0
        ReDim udtPairMatchStats(0)
        InitializePairMatchStats udtPairMatchStats(0)
        If eClsPaired(lngUMCIndexOriginal) <> umcpNone And blnPairsPresent Then
            blnReturnAllPairInstances = True
            blnFavorHeavy = (LastSearchTypeN14N15 = SEARCH_N15)
            
            lngPairIndex = PairIndexLookupSearch(CallerID, lngUMCIndexOriginal, _
                                                 objP1IndFastSearch, objP2IndFastSearch, _
                                                 blnReturnAllPairInstances, blnFavorHeavy, _
                                                 lngPairMatchCount, udtPairMatchStats())
        End If
        
        strLineOutEndAddnl = ""
        If lngPairMatchCount > 0 Then
            For lngPairMatchIndex = 0 To lngPairMatchCount - 1
                ' Lookup whether this UMC is the light or heavy member in the pair
                With GelP_D_L(CallerID).Pairs(udtPairMatchStats(lngPairMatchIndex).PairIndex)
                    If .p1 = lngUMCIndexOriginal Then
                        lngPeakFPRType = FPR_Type_N14_N15_L      ' Light member of pair
                    Else
                        lngPeakFPRType = FPR_Type_N14_N15_H      ' Heavy member of pair
                    End If
                End With
                
                With udtPairMatchStats(lngPairMatchIndex)
                    strLineOutMiddle = Trim(.PairIndex) & strSepChar & Trim(lngPeakFPRType) & strSepChar & Trim(.ExpressionRatio) & strSepChar & Trim(.ExpressionRatioStDev) & strSepChar & Trim(.ExpressionRatioChargeStateBasisCount) & strSepChar & Trim(.ExpressionRatioMemberBasisCount) & strSepChar
                    
                    If blnCorrectedIReportEREnabled Then
                        strLineOutEndAddnl = strSepChar & Round(.LabellingEfficiencyF, 4) & strSepChar & .LogERCorrectedForF & strSepChar & .LogERStandardError
                    End If
                    
                    If Not blnIncludeORFInfo Then
                        ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd & strLineOutEndAddnl
                    Else
                        If mUMCMatchStats(mgInd).IDIsInternalStd Then
                            ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd & strLineOutEndAddnl & strSepChar & "1" & strSepChar & strInternalStdDescription
                        Else
                            WriteORFResults ts, strLineOut & strLineOutMiddle & strLineOutEnd & strLineOutEndAddnl, AMTData(lngMassTagIndexOriginal).ID, objORFNameFastSearch, strSepChar
                        End If
                    End If
                    
                End With
            Next lngPairMatchIndex
        Else
            ' No pair, and thus no expression ratio values
            strLineOutMiddle = Trim(-1) & strSepChar & Trim(-1) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar & Trim(0) & strSepChar
            
            If blnCorrectedIReportEREnabled Then
                strLineOutEndAddnl = strSepChar & "0" & strSepChar & "0" & strSepChar & "0"
            End If
            
            If Not blnIncludeORFInfo Then
                ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd & strLineOutEndAddnl
            Else
                If mUMCMatchStats(mgInd).IDIsInternalStd Then
                    ts.WriteLine strLineOut & strLineOutMiddle & strLineOutEnd & strLineOutEndAddnl & strSepChar & "1" & strSepChar & strInternalStdDescription
                Else
                    WriteORFResults ts, strLineOut & strLineOutMiddle & strLineOutEnd & strLineOutEndAddnl, AMTData(lngMassTagIndexOriginal).ID, objORFNameFastSearch, strSepChar
                End If
            End If
            
        End If
        
        If mgInd Mod 25 = 0 Then
            UpdateStatus "Preparing results: " & Trim(mgInd) & " / " & Trim(mMatchStatsCount)
            If mKeyPressAbortProcess > 1 Then Exit For
        End If
    Next mgInd
    ts.Close
    
    If Len(strOutputFilePath) > 0 Then
        AddToAnalysisHistory CallerID, "Saved search results to disk: " & strOutputFilePath
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus ""
    
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True
    
    If blnDisplayResults Then
         frmDataInfo.Tag = "UMC_MTID"
         frmDataInfo.SourceFilePath = fname
         frmDataInfo.Show vbModal
         frmDataInfo.SourceFilePath = ""
    End If
    
    Set ts = Nothing
    Set fso = Nothing
    Set objORFNameFastSearch = Nothing
Exit Function

ShowOrSaveResultsByUMCErrorHandler:
    Debug.Assert False
    
    ShowOrSaveResultsByUMC = Err.Number
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.ShowOrSaveResultsByUMC"
    Set fso = Nothing
    
    Me.MousePointer = vbDefault
    cmdSearchAllUMCs.Visible = True
    cmdRemoveAMTMatchesFromUMCs.Visible = True

End Function

Private Sub StartExportResultsToDBbyUMC()
    Dim eResponse As VbMsgBoxResult
    Dim strStatus As String
    Dim strUMCSearchMode As String
    
On Error GoTo ExportResultsToDBErrorHandler
    
    If mMatchStatsCount = 0 And Not glbPreferencesExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches Then
        MsgBox "Search results not found in memory.", vbInformation + vbOKOnly, "Nothing to Export"
    Else
        eResponse = MsgBox("Proceed with exporting of the search results to the database?  This is an advanced feature that should normally only be performed during VIPER Automated PRISM Analysis Mode.  If you continue, you will be prompted for a password.", vbQuestion + vbYesNo + vbDefaultButton1, "Export Results")
        If eResponse = vbYes Then
            If QueryUserForExportToDBPassword(, False) Then
                ' Update the text in MD_Parameters
                strUMCSearchMode = FindSettingInAnalysisHistory(CallerID, UMC_SEARCH_MODE_SETTING_TEXT, , True, ":", ";")
                If Right(strUMCSearchMode, 1) = ")" Then strUMCSearchMode = Left(strUMCSearchMode, Len(strUMCSearchMode) - 1)
                GelAnalysis(CallerID).MD_Parameters = ConstructAnalysisParametersText(CallerID, strUMCSearchMode, AUTO_SEARCH_UMC_CONGLOMERATE)
                
                strStatus = ExportMTDBbyUMC(True, mnuFExportDetailedMemberInformation.Checked)
                MsgBox strStatus, vbInformation + vbOKOnly, glFGTU
            Else
                MsgBox "Invalid password, export aborted.", vbExclamation Or vbOKOnly, "Invalid"
            End If
        End If
    End If
    
    Exit Sub
    
ExportResultsToDBErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.StartExportResultsToDBbyUMC"
    Resume Next

End Sub

Public Function StartSearchAll() As Long
    ' Returns the number of hits
    StartSearchAll = PerformSearch(eSearchModeAll)
End Function

Public Function StartSearchPaired() As Long
    ' Returns the number of hits
    StartSearchPaired = PerformSearch(eSearchModePaired)
End Function

Public Function StartSearchLightPairsPlusNonPaired() As Long
    ' Returns the number of hits
    StartSearchLightPairsPlusNonPaired = PerformSearch(eSearchModePairedPlusUnpaired)
End Function

Public Function StartSearchNonPaired() As Long
    ' Returns the number of hits
    StartSearchNonPaired = PerformSearch(eSearchModeNonPaired)
End Function

Private Function StartsWith(ByVal strText As String, ByVal strTextToFind As String) As Boolean

    Dim blnMatched As Boolean
    Dim lngStrLen As Long
    
On Error GoTo StartsWithErrorHandler
    lngStrLen = Len(strTextToFind)
    
    If Len(strText) >= lngStrLen Then
        If UCase(Left(strText, lngStrLen)) = UCase(strTextToFind) Then
            blnMatched = True
        End If
    End If

    StartsWith = blnMatched
    Exit Function

StartsWithErrorHandler:
    Debug.Assert False
    StartsWith = False
    
End Function
Private Sub UpdateUMCsPairingStatusNow()
    Dim blnSuccess As Boolean
    blnSuccess = UpdateUMCsPairingStatus(CallerID, eClsPaired())
End Sub

Private Sub UpdateSTACPlot()

    ' 2D array of doubles with X values in the first dimension and Y values in the second dimension
    Dim varMatches() As Variant
    Dim varMatchesFiltered() As Variant
    Dim varFDR() As Variant
    
    Dim lngIndex As Long
    Dim lngTargetIndex As Long
    Dim lngEndIndex As Long
    
    Dim blnUPFilteredFDR As Boolean
    
On Error GoTo UpdateSTACPlotErrorHandler
    
    ctlSTACStats.ClearData
    
    If STACStatsCount <= 0 Then
        Exit Sub
    End If
    
    blnUPFilteredFDR = Me.PlotUPFilteredFDR
    
    ' Find the first non-zero Matches entry in STACStats()
    For lngIndex = 0 To STACStatsCount - 1
        lngEndIndex = lngIndex
        If STACStats(lngIndex).Matches > 0 Then Exit For
    Next lngIndex
    
    If lngEndIndex >= STACStatsCount - 1 Then
        lngEndIndex = 0
    End If
    
    ReDim varMatches(1, STACStatsCount - (lngEndIndex + 1))
    ReDim varMatchesFiltered(1, STACStatsCount - (lngEndIndex + 1))
    ReDim varFDR(1, STACStatsCount - (lngEndIndex + 1))
    mMaxPlottedFDR = 0
    
    For lngIndex = STACStatsCount - 1 To lngEndIndex Step -1
        lngTargetIndex = STACStatsCount - 1 - lngIndex
        
        varMatches(0, lngTargetIndex) = STACStats(lngIndex).STACCuttoff
        varMatches(1, lngTargetIndex) = STACStats(lngIndex).Matches
        
        varMatchesFiltered(0, lngTargetIndex) = STACStats(lngIndex).STACCuttoff
        varMatchesFiltered(1, lngTargetIndex) = STACStats(lngIndex).UP_Filtered_Matches

        varFDR(0, lngTargetIndex) = STACStats(lngIndex).STACCuttoff
        If blnUPFilteredFDR Then
            varFDR(1, lngTargetIndex) = STACStats(lngIndex).UP_Filtered_FDR / 100#
        Else
            varFDR(1, lngTargetIndex) = STACStats(lngIndex).FDR / 100#
        End If
        
        If varFDR(1, lngTargetIndex) = 0 And lngTargetIndex > 0 Then
            ' When the reported FDR is 0, use the previous value
            Debug.Assert False
            varFDR(1, lngTargetIndex) = varFDR(1, lngTargetIndex - 1)
        End If
        
        If varFDR(1, lngTargetIndex) > mMaxPlottedFDR Then
            mMaxPlottedFDR = varFDR(1, lngTargetIndex)
        End If
        
    Next lngIndex
    
    ctlSTACStats.Plots(1).PlotXY varMatches
    ctlSTACStats.Plots(2).PlotXY varMatchesFiltered
    ctlSTACStats.Plots(3).PlotXY varFDR
    
    UpdateSTACPlotLayout
        
    ZoomOutSTACPlot
    
    Exit Sub
    
UpdateSTACPlotErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.UpdateSTACPlot"
        
End Sub

Private Sub UpdateSTACPlotLayout()
    
    Dim strCaption As String
    
    strCaption = "STAC Trends -- Red=FDR, Blue=Matches, Green=UP Filtered Matches"
    
    ctlSTACStats.Caption = strCaption
    
    ctlSTACStats.Axes(1).Caption = "STAC Threshold"
    ctlSTACStats.Axes(2).Caption = "Matches"
    
    If cChkBox(chkPlotUPFilteredFDR) Then
        ctlSTACStats.Axes(3).Caption = "FDR, UP > 0.5"
    Else
        ctlSTACStats.Axes(3).Caption = "FDR"
    End If
    
  
    ctlSTACStats.Axes(1).Ticks.MajorGrid = cChkBox(chkSTACPlotXGridlines)
    ctlSTACStats.Axes(1).Ticks.MajorGridColor = vbBlack
    
    ' Do not show the gridlines for the left y axis (matches)
    ctlSTACStats.Axes(2).Ticks.MajorGrid = cChkBox(chkSTACPlotY1Gridlines)
    ctlSTACStats.Axes(2).Ticks.MajorGridColor = vbBlack
        
    ' Show the gridlines for the right y axis (FDR)
    ctlSTACStats.Axes(3).Ticks.MajorGrid = cChkBox(chkSTACPlotY2Gridlines)
    ctlSTACStats.Axes(3).Ticks.MajorGridColor = vbBlack
    
    ' Matches
    With ctlSTACStats.Plots(1)
        .LineStyle = cwLineSolid
        .LineWidth = 2
        .PointStyle = cwPointNone
    End With
    
    ' Matches, UP > 0.5
    With ctlSTACStats.Plots(2)
        .LineStyle = cwLineSolid
        .LineWidth = 2
        .PointStyle = cwPointNone
    End With
    
    ' FDR
    With ctlSTACStats.Plots(3)
        .LineStyle = cwLineSolid
        .LineWidth = 2
        .PointStyle = cwPointNone
    End With

End Sub

Private Sub UpdateStatus(ByVal Msg As String)
    lblStatus.Caption = Msg
    DoEvents
End Sub

Private Sub WriteAMTMatchesForIon(ts As TextStream, strLineOutPrefix As String, dblIonMass As Double, AMTRefs() As String, AMTRefsCnt As Long, objAMTIDFastSearch As FastSearchArrayLong, objMTOrIndFastSearch As FastSearchArrayLong, objMTIndFastSearch As FastSearchArrayLong, ByRef lngExportCount, blnIncludeORFInfo As Boolean, objORFNameFastSearch As FastSearchArrayLong, Optional strSepChar As String = glARG_SEP)
    ' Note: AMTRefs() is 1-based
    
    Dim strBaseMatchInfo As String
    Dim strLineOut As String
    Dim lngAMTRefIndex As Long
    Dim lngMassTagID As Long
    
    Dim lngOriginalAMTIndex As Long             ' Index of the AMT in AMTData().MW, etc.
    Dim lngMTOrIndIndexOriginal As Long         ' Index of the AMT in mMTOrInd()
    Dim lngMassTagIndexOriginal As Long         ' Index of teh AMT in AMTData()
    
    Dim lngMatchingIndices() As Long            ' Used with both objAMTIDFastSearch and objMTOrIndFastSearch
    Dim lngMatchCount As Long
    
    Dim lngMTIndMatchingIndices() As Long       ' Index of the AMT in mMTInd()
    Dim lngMTIndMatchCount As Long
    
    Dim lngPointerIndex As Long
    
    Dim dblAMTMass As Double
    Dim dblBestAMTMass As Double, dblBestAMTMassDiff As Double
    Dim strBestAMTMods As String, strBestSequence As String
   
    ' AMTRefsCnt is the number of AMTs that this ion matched (aka MultiMassTagHitCount)
    strBaseMatchInfo = strLineOutPrefix & strSepChar & AMTRefsCnt
    For lngAMTRefIndex = 1 To AMTRefsCnt         'extract MT tag ID
        lngMassTagID = CLng(GetIDFromString(AMTRefs(lngAMTRefIndex), AMTMark, AMTIDEnd))
        
        strLineOut = strBaseMatchInfo & strSepChar & lngMassTagID & strSepChar
        
        If objAMTIDFastSearch.FindMatchingIndices(lngMassTagID, lngMatchingIndices(), lngMatchCount) Then
            ' Match Found
            
            lngOriginalAMTIndex = lngMatchingIndices(0)
            
            ' Now look for lngOriginalAMTIndex in lngMTOrInd()
            ' It could actually be present several times if the mass modifications were
            '  defined as dynamic (rather than static)
            If objMTOrIndFastSearch.FindMatchingIndices(lngOriginalAMTIndex, lngMatchingIndices(), lngMatchCount) Then
                ' Match Found
                
                dblBestAMTMass = 0
                strBestAMTMods = ""
                For lngPointerIndex = 0 To lngMatchCount - 1
                    lngMTOrIndIndexOriginal = lngMatchingIndices(lngPointerIndex)
                    
                    ' Now look for lngMTOrIndIndexOriginal in mMTInd()
                    If objMTIndFastSearch.FindMatchingIndices(lngMTOrIndIndexOriginal, lngMTIndMatchingIndices(), lngMTIndMatchCount) Then
                        ' Match found
                        
                        lngMassTagIndexOriginal = mMTOrInd(lngMTIndMatchingIndices(0))
            
                        If LastSearchTypeN14N15 = SEARCH_N14 Then
                            ' N14
                            dblAMTMass = mMTMWN14(lngMTIndMatchingIndices(0))
                        Else
                            ' N15
                            dblAMTMass = mMTMWN15(lngMTIndMatchingIndices(0))
                        End If
                        
                        If dblBestAMTMass = 0 Then
                            dblBestAMTMass = dblAMTMass
                            dblBestAMTMassDiff = Abs(dblAMTMass - dblIonMass)
                            strBestAMTMods = mMTMods(lngMTOrIndIndexOriginal)
                            
                            If lngMassTagIndexOriginal <= AMTCnt Then
                                strBestSequence = AMTData(lngMassTagIndexOriginal).Sequence
                            Else
                                ' Invalid MT tag index
                                Debug.Assert False
                            End If
                           
                        Else
                            If Abs(dblAMTMass - dblIonMass) < dblBestAMTMassDiff Then
                                dblBestAMTMass = dblAMTMass
                                dblBestAMTMassDiff = Abs(dblAMTMass - dblIonMass)
                                strBestAMTMods = mMTMods(lngMTOrIndIndexOriginal)
                                
                                If lngMassTagIndexOriginal <= AMTCnt Then
                                    strBestSequence = AMTData(lngMassTagIndexOriginal).Sequence
                                Else
                                    ' Invalid MT tag index
                                    Debug.Assert False
                                End If
                            End If
                        End If
                    End If
                Next lngPointerIndex
                
                dblAMTMass = dblBestAMTMass
                If dblBestAMTMass <> 0 Then
                    Debug.Assert Abs(dblAMTMass - dblIonMass) < 0.5
                    Debug.Assert Abs(dblAMTMass - AMTData(lngOriginalAMTIndex).MW) < 0.0001 Or dblAMTMass > AMTData(lngOriginalAMTIndex).MW
                End If
            Else
                dblAMTMass = 0
            End If
        Else
            dblAMTMass = 0
        End If
        
        strLineOut = strLineOut & Round(dblAMTMass, 6) & strSepChar & NTypeStr
        If Len(strBestAMTMods) > 0 Then
            strLineOut = strLineOut & " " & strBestAMTMods
        End If
        strLineOut = strLineOut & strSepChar & strBestSequence & strSepChar
        
        If Not blnIncludeORFInfo Then
            ts.WriteLine strLineOut
        Else
            WriteORFResults ts, strLineOut, lngMassTagID, objORFNameFastSearch, strSepChar
        End If
        
        lngExportCount = lngExportCount + 1
    Next lngAMTRefIndex

End Sub

Private Sub WriteORFResults(ts As TextStream, strLineOutPrefix As String, lngMassTagID As Long, objORFNameFastSearch As FastSearchArrayLong, Optional strSepChar As String = glARG_SEP)
    
    Dim ORFNames() As String            ' 0-based array
    Dim lngORFNamesCount As Long
    Dim lngORFNameIndex As Long

    If MTtoORFMapCount = 0 Then
        lngORFNamesCount = LookupORFNamesForMTIDusingMTDBNamer(objMTDBNameLookupClass, lngMassTagID, ORFNames())
    Else
        lngORFNamesCount = LookupORFNamesForMTIDusingMTtoORFMapOptimized(lngMassTagID, ORFNames(), objORFNameFastSearch)
    End If
    
    If lngORFNamesCount > 0 Then
        For lngORFNameIndex = 0 To lngORFNamesCount - 1
            ts.WriteLine strLineOutPrefix & strSepChar & lngORFNamesCount & strSepChar & ORFNames(lngORFNameIndex)
        Next lngORFNameIndex
    Else
        ts.WriteLine strLineOutPrefix & strSepChar & "0" & strSepChar & "UnknownORF"
    End If

End Sub

Private Sub ZoomOutSTACPlot()
    
    Static intRecursionDepth As Integer
    
On Error GoTo ZoomOutSTACPlotErrorHandler
    
    Dim varYData As Variant
    Dim lngIndex As Long
    
    If False Then
        Debug.Print Me.width, Me.Height
        Me.AutoSizeForm True
    End If
    
    If False Then
        intRecursionDepth = intRecursionDepth + 1
        If intRecursionDepth <= 1 Then
            UpdateSTACPlot
        End If
        intRecursionDepth = intRecursionDepth - 1
    End If
        
    ' Set the range of the X axis to be 0 to 1
    With ctlSTACStats.Plots(1).XAxis
        .AutoScale = False
        .Minimum = 0
        .Maximum = 1
    End With
    
    ' Auto-scale the left Y-Axis
    With ctlSTACStats.Plots(1).YAxis
        .AutoScale = True
        .AutoScaleNow
    End With
    
    With ctlSTACStats.Plots(3).YAxis
        .AutoScale = False
        .Minimum = 0
        
        ' Scale the right Y-Axis based on mMaxPlottedFDR
        If mMaxPlottedFDR <= 0.1 Then
            .Maximum = 0.1
        ElseIf mMaxPlottedFDR <= 0.2 Then
            .Maximum = 0.2
        ElseIf mMaxPlottedFDR <= 0.3 Then
            .Maximum = 0.3
        ElseIf mMaxPlottedFDR <= 0.5 Then
            .Maximum = 0.5
        ElseIf mMaxPlottedFDR <= 0.75 Then
            .Maximum = 0.75
        Else
            .Maximum = 1
        End If
    End With

    Exit Sub

ZoomOutSTACPlotErrorHandler:
    Debug.Assert False

End Sub

Private Sub cboAMTSearchResultsBehavior_Click()
    On Error Resume Next
    If Not bLoading Then
        glbPreferencesExpanded.AMTSearchResultsBehavior = cboAMTSearchResultsBehavior.ListIndex
    End If
End Sub

Private Sub cboResidueToModify_Click()
    If cboResidueToModify.List(cboResidueToModify.ListIndex) = glPHOSPHORYLATION Then
        txtResidueToModifyMass = Trim(glPHOSPHORYLATION_Mass)
    Else
        ' For safety reasons, reset txtResidueToModifyMass to "0"
        txtResidueToModifyMass = "0"
    End If
End Sub

Private Sub chkAlkylation_Click()
    If cChkBox(chkAlkylation) And CDblSafe(txtAlkylationMWCorrection) <= 0 Then
        txtAlkylationMWCorrection = glALKYLATION
        AlkMWCorrection = glALKYLATION
    End If
End Sub

Private Sub chkDisableCustomNETs_Click()
    EnableDisableNETFormulaControls
End Sub

Private Sub chkPlotUPFilteredFDR_Click()
    UpdateSTACPlot
End Sub

Private Sub chkSTACPlotXGridlines_Click()
    UpdateSTACPlotLayout
End Sub

Private Sub chkSTACPlotY1Gridlines_Click()
    UpdateSTACPlotLayout
End Sub

Private Sub chkSTACPlotY2Gridlines_Click()
    UpdateSTACPlotLayout
End Sub

Private Sub chkSTACUsesPriorProbability_Click()
    Me.STACUsesPriorProbability = cChkBox(chkSTACUsesPriorProbability)
End Sub

Private Sub chkUseSTAC_Click()
    Me.UseSTAC = cChkBox(chkUseSTAC)
    EnableDisableControls
End Sub

Private Sub chkUseUMCConglomerateNET_Click()
    glbPreferencesExpanded.UseUMCConglomerateNET = cChkBox(chkUseUMCConglomerateNET)
End Sub

Private Sub cmdCancel_Click()
    mKeyPressAbortProcess = 2
    KeyPressAbortProcess = 2
End Sub

Private Sub cmdCopySTACSTats_Click()
    ShowOrSaveSTACStats True
End Sub

Private Sub cmdRemoveAMTMatchesFromUMCs_Click()
    RemoveAMTMatchesFromUMCs True
End Sub

Private Sub cmdSearchAllUMCs_Click()
    StartSearchAll
End Sub

Private Sub cmdSetDefaults_Click()
    SetDefaultOptions False
End Sub

Private Sub cmdSetDefaultsForToleranceRefinement_Click()
    SetDefaultOptions True
End Sub

Private Sub cmdZoomOutSTACPlot_Click()
    ZoomOutSTACPlot
End Sub

Private Sub Form_Activate()
    InitializeSearch
End Sub

Private Sub Form_Load()
    '----------------------------------------------------
    'load search settings and initializes controls
    '----------------------------------------------------
    
    Dim intIndex As Integer
    
    On Error GoTo FormLoadErrorHandler
    
    bLoading = True
    If IsWinLoaded(TrackerCaption) Then Unload frmTracker
    If CallerID >= 1 And CallerID <= UBound(GelBody) Then samtDef = GelSearchDef(CallerID).AMTSearchOnUMCs
    
    If APP_BUILD_DISABLE_LCMSWARP Then
        chkDisableCustomNETs.Visible = False
    End If
    
    mSTACSessionID = ""
    Set mTempFilesToDelete = New Dictionary
    
    ShowHidePNNLMenus
    
    EnableDisableControls
    
    'set current Search Definition values
    DisplayCurrentSearchTolerances
    
    With samtDef
        If glbPreferencesExpanded.AMTSearchResultsBehavior = asrbKeepExistingAndSkip Then
            .SkipReferenced = True
        Else
            .SkipReferenced = False
        End If
        
        optNETorRT(.NETorRT).Value = True
        
        'save old value and set search on "search all"
        OldSearchFlag = .SearchFlag
        .SearchFlag = 0         'search all
        
        mnuET(etGANET).Checked = True
    End With
    
    Me.UseSTAC = glbPreferencesExpanded.UseSTAC
    Me.STACUsesPriorProbability = glbPreferencesExpanded.STACUsesPriorProbability
    
    With GelSearchDef(CallerID).AMTSearchMassMods
        SetCheckBox chkPEO, .PEO
        SetCheckBox chkICATLt, .ICATd0
        SetCheckBox chkICATHv, .ICATd8
        SetCheckBox chkAlkylation, .Alkylation
        txtAlkylationMWCorrection = .AlkylationMass
        
        PopulateComboBoxes
        
        cboResidueToModify.ListIndex = 0
        If Len(.ResidueToModify) >= 1 Then
            For intIndex = 0 To cboResidueToModify.ListCount - 1
                If UCase(cboResidueToModify.List(intIndex)) = UCase(.ResidueToModify) Then
                    cboResidueToModify.ListIndex = intIndex
                    Exit For
                End If
            Next intIndex
        End If
        txtResidueToModifyMass = Round(.ResidueMassModification, 5)
        
        SetAlkylationMWCorrection .AlkylationMass
        SetDBSearchModType .ModMode
        
        SetDBSearchNType .N15InsteadOfN14
    End With
    
    With glbPreferencesExpanded
        cboAMTSearchResultsBehavior.ListIndex = .AMTSearchResultsBehavior
        SetCheckBox chkUseUMCConglomerateNET, .UseUMCConglomerateNET
    End With
    
    With glbPreferencesExpanded.MTSConnectionInfo
        ExpAnalysisSPName = .spPutAnalysis
        'ExpPeakSPName = .spPutPeak
        ExpUmcSPName = .spPutUMC
        ExpUMCMemberSPName = .spPutUMCMember
        ExpUmcMatchSPName = .spPutUMCMatch
        ExpUmcInternalStdMatchSPName = .spPutUMCInternalStdMatch
        ExpUMCCSStats = .spPutUMCCSStats
        ExpStoreSTACStats = .spPutSTACStats
        ExpQuantitationDescription = .spAddQuantitationDescription
    End With
    
    If Not GelAnalysis(CallerID) Is Nothing Then
        mMDTypeSaved = GelAnalysis(CallerID).MD_Type
    Else
        mMDTypeSaved = stNotDefined
    End If
    
    If Len(ExpUmcSPName) = 0 Then
        ExpUmcSPName = "AddFTICRUmc"
    End If
    Debug.Assert ExpUmcSPName = "AddFTICRUmc"
    
    If Len(ExpUmcMatchSPName) = 0 Then
        ExpUmcMatchSPName = "AddFTICRUmcMatch"
    End If
    Debug.Assert ExpUmcMatchSPName = "AddFTICRUmcMatch"
    
    If Len(ExpUmcInternalStdMatchSPName) = 0 Then
        ExpUmcInternalStdMatchSPName = "AddFTICRUmcInternalStdMatch"
    End If
    Debug.Assert ExpUmcInternalStdMatchSPName = "AddFTICRUmcInternalStdMatch"
    
    If Len(ExpUMCCSStats) = 0 Then
        ExpUMCCSStats = "AddFTICRUmcCSStats"
    End If
    Debug.Assert ExpUMCCSStats = "AddFTICRUmcCSStats"
    
    If Len(ExpStoreSTACStats) = 0 Then
        ExpStoreSTACStats = "AddMatchMakingFDR"
    End If
    Debug.Assert ExpStoreSTACStats = "AddMatchMakingFDR"
    
    If Len(ExpQuantitationDescription) = 0 Then
        ExpQuantitationDescription = "AddQuantitationDescription"
    End If
    Debug.Assert ExpQuantitationDescription = "AddQuantitationDescription"
    
    If Len(ExpAnalysisSPName) = 0 Then
        ExpAnalysisSPName = "AddMatchMaking"
    End If
    Debug.Assert ExpAnalysisSPName = "AddMatchMaking"
    
    ' September 2004: Unused Variable
    ''If Len(ExpPeakSPName) = 0 Then
    ''    ExpPeakSPName = "AddFTICRPeak"
    ''End If
    ''Debug.Assert ExpPeakSPName = "AddFTICRPeak"
    
    ' Possibly add a checkmark to the mnuFReportIncludeORFs menu
    mnuFReportIncludeORFs.Checked = glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput
    
    InitializeSTACStatsListView
    
    Exit Sub

FormLoadErrorHandler:
    LogErrors Err.Number, "frmSearchMT_ConglomerateUMC.Form_Load"
    Resume Next

End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Restore .SearchFlag using the saved value
    samtDef.SearchFlag = OldSearchFlag
    
    If Not objMTDBNameLookupClass Is Nothing Then
        objMTDBNameLookupClass.DeleteData
        Set objMTDBNameLookupClass = Nothing
    End If
    
    ' Restore .MD_Type from mMDTypeSaved
    If Not GelAnalysis(CallerID) Is Nothing Then
        GelAnalysis(CallerID).MD_Type = mMDTypeSaved
    End If
    
    DeleteTempFiles
End Sub

Private Sub mnuEditCopySTACPlot_Click()
    SaveSTACPlotToClipboardOrEMF ""
End Sub

Private Sub mnuEditCopySTACStats_Click()
    ShowOrSaveSTACStats True
End Sub

Private Sub mnuEditSaveSTACPlotAsEMF_Click()
    SaveSTACPlotToFile pftEMF
End Sub

Private Sub mnuEditSaveSTACPlotAsPNG_Click()
    SaveSTACPlotToFile pftPNG
End Sub

Private Sub mnuEditSetToDefaults_Click()
    SetDefaultOptions False
End Sub

Private Sub mnuET_Click(Index As Integer)
    SetETMode (Index)
End Sub

Private Sub mnuETHeader_Click()
Call PickParameters
End Sub

Private Sub mnuF_Click()
Call PickParameters
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFDeleteExcludedPairs_Click()
    Me.DeleteExcludedPairsWrapper
End Sub

Private Sub mnuFExcludeAmbiguous_Click()
    Me.ExcludeAmbiguousPairsWrapper False
End Sub

Private Sub mnuFExcludeAmbiguousHitsOnly_Click()
    Me.ExcludeAmbiguousPairsWrapper True
End Sub

Private Sub mnuFExportDetailedMemberInformation_Click()
    mnuFExportDetailedMemberInformation.Checked = Not mnuFExportDetailedMemberInformation.Checked
End Sub

Private Sub mnuFExportResultsToDBbyUMC_Click()
    StartExportResultsToDBbyUMC
End Sub

Private Sub mnuFMassCalAndToleranceRefinement_Click()
    ShowErrorDistribution2DForm
End Sub

Private Sub mnuFReportByIon_Click()
    ShowOrSaveResultsByIon "", True, mnuFReportIncludeORFs.Checked
End Sub

Private Sub mnuFReportByUMC_Click()
    ShowOrSaveResultsByUMC "", True, mnuFReportIncludeORFs.Checked
End Sub

Private Sub mnuFReportIncludeORFs_Click()
    mnuFReportIncludeORFs.Checked = Not mnuFReportIncludeORFs.Checked
    glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput = mnuFReportIncludeORFs.Checked
End Sub

Private Sub mnuFResetExclusionFlags_Click()
Dim strMessage As String
strMessage = PairsResetExclusionFlag(CallerID)
UpdateUMCsPairingStatusNow
UpdateStatus strMessage
End Sub

Private Sub mnuFSearchAll_Click()
StartSearchAll
End Sub

Private Sub mnuFSearchN14LabeledFeatures_Click()
    '''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''
    '''    ToDo: Implement this search
    '''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''
    Debug.Assert False
    
    ' Step 0: Clear any existing pairs
    ' Step 1: Match UMCs to the DB, though only use N14 UMCs
    ' Step 2: For each match, compute the average peptide mass, then compare to the average mass value of the N15 UMCs
    '         However, if the N15 UMCs have monoisotopic mass values, then compute the mono peptide mass
    '         For matching UMCs, check their scan boundaries using the same code as is on the PairSearch form
    '         If a match, then create a new pair, compute the ER (not scan-by-scan, but use charge-state rules from the Pair Search form), and store the peptide hit
    '
    
End Sub

Private Sub mnuFSearchNonPaired_Click()
StartSearchNonPaired
End Sub

Private Sub mnuFSearchPaired_Click()
StartSearchPaired
End Sub

Private Sub mnuFSearchPairedPlusNonPaired_Click()
StartSearchLightPairsPlusNonPaired
End Sub

Private Sub mnuMT_Click()
Call PickParameters
End Sub

Private Sub mnuMTLoadLegacy_Click()
    LoadLegacyMassTags
End Sub

Private Sub mnuMTLoadMT_Click()
'------------------------------------------------------------
'load/reload MT tags
'------------------------------------------------------------
If Not GelAnalysis(CallerID) Is Nothing Then
   Call LoadMTDB(True)
Else
   WarnUserNotConnectedToDB CallerID, True
   lblMTStatus.Caption = "No MT tags loaded"
End If
End Sub

Private Sub mnuMTStatus_Click()
'----------------------------------------------
'displays short MT tags statistics, it might
'help with determining problems with MT tags
'----------------------------------------------
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly
Me.MousePointer = vbDefault
End Sub

Private Sub optDBSearchModType_Click(Index As Integer)
    EnableDisableControls
End Sub

Private Sub optN_Click(Index As Integer)
N14N15 = Index
End Sub

Private Sub optNETorRT_Click(Index As Integer)
samtDef.NETorRT = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   samtDef.TolType = gltPPM
Else
   samtDef.TolType = gltABS
End If
End Sub

Private Sub txtAlkylationMWCorrection_LostFocus()
If IsNumeric(txtAlkylationMWCorrection.Text) Then
   AlkMWCorrection = CDbl(txtAlkylationMWCorrection.Text)
Else
   txtAlkylationMWCorrection.Text = glALKYLATION
   AlkMWCorrection = glALKYLATION
End If
End Sub

Private Sub txtDBSearchMinimumHighDiscriminantScore_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumHighDiscriminantScore, 0, 1, 0
End Sub

Private Sub txtDBSearchMinimumHighNormalizedScore_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumHighNormalizedScore, 0, 100000, 0
End Sub

Private Sub txtDBSearchMinimumPeptideProphetProbability_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumPeptideProphetProbability, 0, 1, 0
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   samtDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "Molecular Mass Tolerance should be numeric value.", vbOKOnly
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNETFormula_LostFocus()
'------------------------------------------------
'initialize new expression evaluator
'------------------------------------------------
    If Not InitExprEvaluator(txtNETFormula.Text) Then
       MsgBox "Error in elution calculation formula.", vbOKOnly, glFGTU
       txtNETFormula.SetFocus
    Else
       samtDef.Formula = txtNETFormula.Text
    End If
End Sub

Private Sub txtNETTol_LostFocus()
If IsNumeric(txtNETTol.Text) Then
   samtDef.NETTol = CDbl(txtNETTol.Text)
Else
   If Len(Trim(txtNETTol.Text)) > 0 Then
      MsgBox "NET Tolerance should be number between 0 and 1.", vbOKOnly
      txtNETTol.SetFocus
   Else
      samtDef.NETTol = -1   'do not consider NET when searching
   End If
End If
End Sub

Private Sub txtNETTol_Validate(Cancel As Boolean)
    TextBoxLimitNumberLength txtNETTol, 12
End Sub

Private Sub txtResidueToModifyMass_LostFocus()
    ValidateTextboxValueDbl txtResidueToModifyMass, -10000, 10000, 0
End Sub

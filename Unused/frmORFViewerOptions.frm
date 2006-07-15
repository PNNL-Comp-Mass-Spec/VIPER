VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmORFViewerOptions 
   Caption         =   "ORF Viewer Options"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabOptions 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Plotting Options"
      TabPicture(0)   =   "frmORFViewerOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraPlotOptions"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mass Tag Options"
      TabPicture(1)   =   "frmORFViewerOptions.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraViewOptions"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ORF Intensity Options"
      TabPicture(2)   =   "frmORFViewerOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraORFIntensityOptions"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraORFIntensityOptions 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   62
         Top             =   360
         Width           =   5535
         Begin VB.CheckBox chkOnlyUseTop50PctForAveraging 
            Caption         =   "Only use Top 50% of mass tags (by intensity) for ion counts and ORF intensity average"
            Height          =   495
            Left            =   240
            TabIndex        =   63
            Top             =   240
            Value           =   1  'Checked
            Width           =   3615
         End
      End
      Begin VB.Frame fraViewOptions 
         Height          =   2535
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   5535
         Begin VB.CheckBox chkHideEmptyMassTagPics 
            Caption         =   "Hide empty mass tag zoomed views"
            Height          =   495
            Left            =   3480
            TabIndex        =   64
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkShowNonTrypticMassTagsWithoutIonHits 
            Caption         =   "Show non-tryptic mass tags without ion hits"
            Height          =   495
            Left            =   3480
            TabIndex        =   45
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox chkAddUnobservedTrypticMassTags 
            Caption         =   "Add unobserved tryptic mass tags (theoretical NET values)"
            Height          =   735
            Left            =   3480
            TabIndex        =   44
            Top             =   720
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.ComboBox cboMassTagShape 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   990
            Width           =   1695
         End
         Begin VB.CheckBox chkLoadPMTs 
            Caption         =   "Load PMT's in addition to AMT's"
            Height          =   375
            Left            =   3480
            TabIndex        =   43
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtMassTagMassError 
            Height          =   285
            Left            =   2640
            TabIndex        =   34
            Text            =   "15"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtMassTagNETError 
            Height          =   285
            Left            =   2640
            TabIndex        =   36
            Text            =   "0.15"
            Top             =   600
            Width           =   735
         End
         Begin VB.ComboBox cboCleavageRuleType 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   2040
            Width           =   3135
         End
         Begin VB.Label lblMassTagColorSelection 
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   40
            ToolTipText     =   "Double click to change"
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lblMassTagColorLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Tag Color"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1350
            Width           =   1455
         End
         Begin VB.Label lblMassTagShape 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass Tag Shape"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label lblMassTagMassError 
            Caption         =   "Mass Tag Mass Error (PPM)"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   2295
         End
         Begin VB.Label lblMassTagNETError 
            Caption         =   "Mass Tag NET Error"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   630
            Width           =   2295
         End
         Begin VB.Label lblCleavageRuleType 
            Caption         =   "Cleavage rule to group mass tags by"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1800
            Width           =   2895
         End
      End
      Begin VB.Frame fraPlotOptions 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   5535
         Begin VB.TextBox txtIonToUMCScalingRatio 
            Height          =   285
            Left            =   2280
            TabIndex        =   25
            Text            =   "2"
            ToolTipText     =   "Amount to scale the ion intensities by when plotting both ions and UMC's on the same plot"
            Top             =   1275
            Width           =   615
         End
         Begin VB.TextBox txtIntensityScalarForListView 
            Height          =   285
            Left            =   2280
            TabIndex        =   23
            Text            =   "1000000"
            ToolTipText     =   "Amount to divide displayed intensities by in the ORF list and Mass Tag list"
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox chkSwapAxes 
            Caption         =   "Swap Axes (NET on Y, Mass on X)"
            Height          =   255
            Left            =   2400
            TabIndex        =   31
            Top             =   2160
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.TextBox txtMinSpotSize 
            Height          =   285
            Left            =   4200
            TabIndex        =   21
            Text            =   "1"
            ToolTipText     =   "Minimum spot size (pixels)"
            Top             =   555
            Width           =   615
         End
         Begin VB.TextBox txtMaxSpotSize 
            Height          =   285
            Left            =   4200
            TabIndex        =   19
            Text            =   "500"
            ToolTipText     =   "Maximum spot size (pixels)"
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox chkShowGridlines 
            Caption         =   "Show Gridlines"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkShowPosition 
            Caption         =   "Show Position"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkShowTickMarkLabels 
            Caption         =   "Tick mark labels"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtPictureHeight 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Text            =   "250"
            ToolTipText     =   "Picture height (pixels)"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtPictureWidth 
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Text            =   "300"
            ToolTipText     =   "Picture width (pixels)"
            Top             =   555
            Width           =   615
         End
         Begin VB.CheckBox chkLogarithmicIntensity 
            Caption         =   "Logarithmic Intensities"
            Height          =   255
            Left            =   2400
            TabIndex        =   29
            Top             =   1680
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkUseUMCClassRepresentativeNET 
            Caption         =   "Use NET of Class Rep. for UMC's"
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            ToolTipText     =   "When checked, will center the UMC spot around the NET of the class representative ion for a UMC.  Results in asymmetric triangles."
            Top             =   1920
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.Label lblIonToUMCScalingRatio 
            Caption         =   "Ion to UMC Scaling Ratio"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1305
            Width           =   2175
         End
         Begin VB.Label lblIntensityScalarForListView 
            Caption         =   "ListView Intensity Scalar"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   990
            Width           =   2175
         End
         Begin VB.Label lblMinSpotSize 
            Caption         =   "Min Spot Size"
            Height          =   255
            Left            =   2880
            TabIndex        =   20
            Top             =   585
            Width           =   1575
         End
         Begin VB.Label lblSpotSize 
            Caption         =   "Max Spot Size"
            Height          =   255
            Left            =   2880
            TabIndex        =   18
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label lblPictureHeight 
            Caption         =   "Picture Height"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label lblPictureWidth 
            Caption         =   "Picture Width"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   585
            Width           =   1575
         End
      End
   End
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   9480
      TabIndex        =   58
      Top             =   2520
      Width           =   975
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   0
         TabIndex        =   60
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   0
         TabIndex        =   61
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame fraDrawingOptionsForSelectedGel 
      Caption         =   "Drawing options for selected display"
      Height          =   2655
      Left            =   6120
      TabIndex        =   46
      Top             =   2400
      Width           =   3135
      Begin VB.ComboBox cboUMCShape 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkVisibleScopeOnly 
         Caption         =   "Only show ions in visible scope"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2280
         Width           =   2535
      End
      Begin VB.ComboBox cboNETAdjustmentType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   1770
         Width           =   1695
      End
      Begin VB.ComboBox cboIonShape 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblUMCColorSelection 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   54
         ToolTipText     =   "Double click to change"
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblUMCColorLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "UMC Color"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblUMCShape 
         BackStyle       =   0  'Transparent
         Caption         =   "UMC Shape"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   990
         Width           =   975
      End
      Begin VB.Label lblNETAdjustmentType 
         BackStyle       =   0  'Transparent
         Caption         =   "NET mode"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblIonShape 
         BackStyle       =   0  'Transparent
         Caption         =   "Ion Shape"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lbIonColorLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Ion Color"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblIonColorSelection 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   50
         ToolTipText     =   "Double click to change"
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame fraGelsToShow 
      Caption         =   "Displays whose ORFs are to be included in the master ORF list"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.Frame fraZOrder 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   9600
         TabIndex        =   8
         Top             =   360
         Width           =   735
         Begin VB.CommandButton cmdZOrderDown 
            Caption         =   "Down"
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdZOrderUp 
            Caption         =   "Up"
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblZOrder 
            Alignment       =   2  'Center
            Caption         =   "Shift Drawing Order"
            Height          =   615
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame fraMoveGels 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   4560
         TabIndex        =   2
         Top             =   330
         Width           =   375
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "< <"
            Height          =   375
            Left            =   0
            TabIndex        =   6
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "> >"
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<"
            Height          =   375
            Left            =   0
            TabIndex        =   5
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   " >"
            Height          =   375
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.ListBox lstGelsInUse 
         Height          =   1620
         Left            =   5160
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
      Begin VB.ListBox lstAvailableGels 
         Height          =   1620
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmORFViewerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OK_BUTTON_CAPTION = "&OK"
Private Const CLOSE_BUTTON_CAPTION = "Cl&ose"

Private mIndexByZOrder() As Long            ' 0-based array: Holds the Indices of the gels at each Z-Order (first ZOrder is has a value of 0)
Private mGelListChanged As Boolean
Private mCallingFormID As Long

Private mLastSelectedSpotShape As sSpotsShape

Private mInitializingControls As Boolean            ' True when initializing the controls; prevents UpdateGeneralOptions() from executing
Private mInitializingSelectedGelControls As Boolean         ' True when initializing the controls for the selected gel; prevents UpdateOptionsForSelectedGels() from executing

Private Sub AddSelectedGels(blnAddAll As Boolean)
    ' Make sure the selected Gels are present in lstGelsInUse
    
    Dim intListIndex As Integer, lngGelIndex As Long
    Dim blnListChanged As Boolean
    Dim lngColonLoc As Long
    
    If blnAddAll Then
        For intListIndex = 0 To lstAvailableGels.ListCount - 1
            lstAvailableGels.Selected(intListIndex) = True
        Next intListIndex
    End If
    
    ' Note: lstAvailableGels is 0-based while gOrfViewerOptionsCurrentGelList.Gels() is 1-based
    For intListIndex = 0 To lstAvailableGels.ListCount - 1
        If lstAvailableGels.Selected(intListIndex) Then
            lngColonLoc = InStr(lstAvailableGels.List(intListIndex), ":")
            
            If lngColonLoc > 0 Then
                lngGelIndex = CLngSafe(Left(lstAvailableGels.List(intListIndex), lngColonLoc - 1))
                
                With gOrfViewerOptionsCurrentGelList.Gels(lngGelIndex)
                    If Not .IncludeGel Then
                        .IncludeGel = True
                        .IonSpotColor = GetUnusedORFViewerSpotColor(.UMCSpotColor, lngGelIndex)
                        .IonSpotColorSelected = RGB(0, 128, 0)
                        .IonSpotShape = mLastSelectedSpotShape
                        
                        .UMCSpotColorSelected = .IonSpotColorSelected
                        .UMCSpotShape = sTriangleWithExtents
                        
                        blnListChanged = True
                    End If
                End With
            Else
                ' Colon not found This shouldn't happen
                Debug.Assert False
            End If
        End If
    Next intListIndex
            
    If blnAddAll Then
        For intListIndex = 0 To lstAvailableGels.ListCount - 1
            lstAvailableGels.Selected(intListIndex) = False
        Next intListIndex
    End If
            
    If blnListChanged Then
        UpdateGelInUseList
        SetGelInListChanged
    End If
    
End Sub

Private Sub ApplyChanges()
    ' Apply changes
    
    SetOKButtonCaption True
    
    gOrfViewerOptionsSavedGelList = gOrfViewerOptionsCurrentGelList
    ORFViewerLoader.UpdateORFViewerForm mCallingFormID, mGelListChanged
    
    SetGelInListChanged False

End Sub

Public Sub InitializeGeneralOptions()
    ' Initializes controls using gOrfViewerOptionsCurrentGelList
    
    mInitializingControls = True
    With gOrfViewerOptionsCurrentGelList.DisplayOptions
        
        txtMassTagMassError = .MassTagMassErrorPPM
        txtMassTagNETError = .MassTagNETError
        
        cboMassTagShape.ListIndex = .MassTagSpotShape
        lblMassTagColorSelection.BackColor = .MassTagSpotColor
        
        cboCleavageRuleType.ListIndex = .CleavageRuleID
        SetCheckBox chkLoadPMTs, .LoadPMTs
        
        txtPictureHeight = .PicturePixelHeight
        txtPictureWidth = .PicturePixelWidth
        
        txtMaxSpotSize = .MaxSpotSizePixels
        txtMinSpotSize = .MinSpotSizePixels
        
        txtIntensityScalarForListView = .IntensityScalar
        txtIonToUMCScalingRatio = .IonToUMCPlottingIntensityRatio
        
        SetCheckBox chkShowPosition, .ShowPosition
        SetCheckBox chkShowTickMarkLabels, .ShowTickMarkLabels
        SetCheckBox chkShowGridlines, .ShowGridLines
        
        SetCheckBox chkOnlyUseTop50PctForAveraging, .OnlyUseTop50PctForAveraging
        
        SetCheckBox chkLogarithmicIntensity, .LogarithmicIntensityPlotting
        SetCheckBox chkUseUMCClassRepresentativeNET, .UseClassRepresentativeNET
        SetCheckBox chkSwapAxes, .SwapPlottingAxes
        
        SetCheckBox chkLoadPMTs, .LoadPMTs
        SetCheckBox chkAddUnobservedTrypticMassTags, .IncludeUnobservedTrypticMassTags
        
        chkShowNonTrypticMassTagsWithoutIonHits.Enabled = True
        SetCheckBox chkShowNonTrypticMassTagsWithoutIonHits, .ShowNonTrypticMassTagsWithoutIonHits
        SetCheckBox chkHideEmptyMassTagPics, .HideEmptyMassTagPictures
        chkShowNonTrypticMassTagsWithoutIonHits.Enabled = Not .HideEmptyMassTagPictures
        
    End With
    mInitializingControls = False
End Sub

Private Sub PopulateComboBoxes()
    Dim intIndex As Integer
    
    With cboIonShape
        .Clear
        .AddItem "Circle", sCircle
        .AddItem "Rectangle", sRectangle
        .AddItem "Round Rectangle", sRoundRectangle
        .AddItem "Star", sStar
        .AddItem "Empty Rectangle", sEmptyRectangle
        .AddItem "Triangle", sTriangleWithExtents
        .ListIndex = sCircle
    
        cboUMCShape.Clear
        cboMassTagShape.Clear
        For intIndex = 0 To .ListCount - 1
            cboUMCShape.AddItem .List(intIndex)
            cboMassTagShape.AddItem .List(intIndex)
        Next intIndex
        cboUMCShape.ListIndex = sTriangleWithExtents
        cboMassTagShape.ListIndex = sEmptyRectangle
        
    End With
    
    With cboNETAdjustmentType
        .Clear
        .AddItem "Generic NET", natGeneric
        .AddItem "TIC NET", natTICNET
        .AddItem "GANET", natGANET
        .ListIndex = natGANET
    End With
    
    Dim objInSilicoDigest As New clsInSilicoDigest
    
    cboCleavageRuleType.Clear
    For intIndex = 0 To objInSilicoDigest.CleaveageRuleCount - 1
        cboCleavageRuleType.AddItem objInSilicoDigest.GetCleaveageRuleName(CInt(intIndex)) & " (" & objInSilicoDigest.GetCleaveageRuleResiduesSymbols(CInt(intIndex)) & ")"
    Next intIndex
   
End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    
    With tabOptions
        .Left = 120
        lngDesiredValue = Me.ScaleHeight - .Height - 120
        If lngDesiredValue < 2500 Then lngDesiredValue = 2500
        .Top = lngDesiredValue
    End With
    
    With fraDrawingOptionsForSelectedGel
        .Left = tabOptions.Left + tabOptions.width + 120
        .Top = tabOptions.Top + 120
    End With
    
    With fraControls
        .Left = fraDrawingOptionsForSelectedGel.Left + fraDrawingOptionsForSelectedGel.width + 120
        .Top = tabOptions.Top + 60
    End With
    
    With fraGelsToShow
        .Left = 120
        .Top = 120
        
        lngDesiredValue = Me.ScaleWidth - 240
        If lngDesiredValue < 2500 Then lngDesiredValue = 2500
        .width = lngDesiredValue
        
        .Height = tabOptions.Top - .Top - 120
    End With
    
    With fraMoveGels
        .Top = 330
        .Left = fraGelsToShow.width / 2 - .width / 2
    End With
    
    With fraZOrder
        .Top = 360
        .Left = fraGelsToShow.width - .width - 120
    End With
    
    With lstAvailableGels
        .Left = 120
        .Top = 360
        .width = fraMoveGels.Left - 240
        .Height = fraGelsToShow.Height - .Top - 60
    End With
    
    With lstGelsInUse
        .Top = lstAvailableGels.Top
        .Left = fraMoveGels.Left + fraMoveGels.width + 120
        lngDesiredValue = fraGelsToShow.width - .Left - fraZOrder.width - 240
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .width = lngDesiredValue
        .Height = lstAvailableGels.Height
    End With
    
End Sub

Private Sub RemoveSelectedGels(blnRemoveAll As Boolean)
    ' Make sure the selected Gels are not present in lstGelsInUse
    
    Dim intListIndex As Integer
    Dim blnListChanged As Boolean
    
    If blnRemoveAll Then
        For intListIndex = 0 To lstGelsInUse.ListCount - 1
            lstGelsInUse.Selected(intListIndex) = True
        Next intListIndex
    End If
    
    For intListIndex = 0 To lstGelsInUse.ListCount - 1
        If lstGelsInUse.Selected(intListIndex) Or blnRemoveAll Then
            With gOrfViewerOptionsCurrentGelList.Gels(mIndexByZOrder(intListIndex))
                .IncludeGel = False
                blnListChanged = True
            End With
        End If
    Next intListIndex
            
    If blnListChanged Then
        UpdateGelInUseList
        SetGelInListChanged
    End If
    
End Sub

Private Sub SelectCustomColor(lblThisLabel As Label)
    Dim lngTemporaryColor As Long
    
    lngTemporaryColor = lblThisLabel.BackColor
    Call GetColorAPIDlg(Me.hwnd, lngTemporaryColor)
    If lngTemporaryColor >= 0 Then
        lblThisLabel.BackColor = lngTemporaryColor
    End If
End Sub

Public Sub SetCallingFormID(lngFormID As Long)
    mCallingFormID = lngFormID
End Sub

Public Sub SetGelInListChanged(Optional blnListChanged As Boolean = True)
    mGelListChanged = blnListChanged
    
    If blnListChanged Then
        SetOKButtonCaption False
    End If
End Sub

Private Sub SetOKButtonCaption(blnShowClose As Boolean)
    If blnShowClose Then
        cmdOK.Caption = CLOSE_BUTTON_CAPTION
    Else
        cmdOK.Caption = OK_BUTTON_CAPTION
    End If
    
End Sub

Private Sub ShowSettingsForSelectedGel()

    mInitializingSelectedGelControls = True
    
    With gOrfViewerOptionsCurrentGelList.Gels(mIndexByZOrder(lstGelsInUse.ListIndex))
        lblIonColorSelection.BackColor = .IonSpotColor
        cboIonShape.ListIndex = .IonSpotShape
        
        lblUMCColorSelection.BackColor = .UMCSpotColor
        cboUMCShape.ListIndex = .UMCSpotShape
        
        If .VisibleScopeOnly Then
            chkVisibleScopeOnly.value = vbChecked
        Else
            chkVisibleScopeOnly.value = vbUnchecked
        End If
        cboNETAdjustmentType.ListIndex = .NETAdjustmentType
    End With

    mInitializingSelectedGelControls = False
End Sub

Private Sub ShuffleGelZOrder(blnMoveTowardsTop As Boolean)
    Dim intIndex As Integer
    Dim lngGelIndex As Long
    Dim lngGelIndexToSwapWith As Long
    Dim intZOrderSwap As Integer
    
    ' Find the index of the first selected Gel
    
    For intIndex = 0 To lstGelsInUse.ListCount - 1
        If lstGelsInUse.Selected(intIndex) Then
            lngGelIndex = mIndexByZOrder(intIndex)
            
            With gOrfViewerOptionsCurrentGelList
                lngGelIndexToSwapWith = -1
                If blnMoveTowardsTop Then
                    If intIndex > 0 Then
                        ' Swap with item directly above in list
                        lngGelIndexToSwapWith = mIndexByZOrder(intIndex - 1)
                    End If
                Else
                    If intIndex < lstGelsInUse.ListCount - 1 Then
                        ' Swap with item directly below in list
                        lngGelIndexToSwapWith = mIndexByZOrder(intIndex + 1)
                    End If
                End If
                If lngGelIndexToSwapWith > 0 Then
                    intZOrderSwap = .Gels(lngGelIndex).ZOrder
                    .Gels(lngGelIndex).ZOrder = .Gels(lngGelIndexToSwapWith).ZOrder
                    .Gels(lngGelIndexToSwapWith).ZOrder = intZOrderSwap
                    
                    UpdateGelInUseList
                    SetGelInListChanged
                End If
            End With
            Exit For
        End If
    Next intIndex
End Sub

Public Sub UpdateGelInUseList()
    Dim lngGelIndex As Long, intZOrderIndex As Integer
    Dim lngLastZOrder As Long
    Dim blnItemAdded As Boolean, blnMatched As Boolean
    
    With gOrfViewerOptionsCurrentGelList
        ' Add the Gels to the list, in the correct Z-Order (0 to highest)
        
        ReDim mIndexByZOrder(.GelCount)
        
        lstGelsInUse.Clear
        
        lngLastZOrder = -1
        Do
            blnItemAdded = False
            For lngGelIndex = 1 To .GelCount
                With .Gels(lngGelIndex)
                    If .IncludeGel Then
                        ' See if already in mIndexByZorder()
                        blnMatched = False
                        For intZOrderIndex = 0 To lstGelsInUse.ListCount - 1
                            If mIndexByZOrder(intZOrderIndex) = lngGelIndex Then
                                blnMatched = True
                                Exit For
                            End If
                        Next intZOrderIndex
                        
                        If Not blnMatched Then
                            If .ZOrder >= lngLastZOrder Then
                                If .ZOrder = lngLastZOrder Then
                                    ' Can't have duplicate z-orders
                                    .ZOrder = lngLastZOrder + 1
                                End If
                                
                                lstGelsInUse.AddItem StripFullPath(.GelFileName)
                                mIndexByZOrder(lstGelsInUse.ListCount - 1) = lngGelIndex
                                blnItemAdded = True
                            End If
                        End If
                    End If
                End With
                ' Can't use an Exit For above because that would jump out of a 'With Block'
                If blnItemAdded Then Exit For
            Next lngGelIndex
            
        Loop While blnItemAdded
        
    End With
    
    If lstGelsInUse.ListCount > 0 Then
        lstGelsInUse.ListIndex = 0
        lstGelsInUse.Selected(0) = True
        ShowSettingsForSelectedGel
    End If
End Sub

Private Sub UpdateOptionsForSelectedGels()
    Dim intListIndex As Integer
    Dim blnVisibleScopeOnly As Boolean
    
    If mInitializingSelectedGelControls Then Exit Sub
    
    For intListIndex = 0 To lstGelsInUse.ListCount - 1
        If lstGelsInUse.Selected(intListIndex) Then
           With gOrfViewerOptionsCurrentGelList.Gels(mIndexByZOrder(intListIndex))
                If .NETAdjustmentType <> cboNETAdjustmentType.ListIndex Then
                    .NETAdjustmentType = cboNETAdjustmentType.ListIndex
                    SetOKButtonCaption False
                End If
                If .IonSpotColor <> lblIonColorSelection.BackColor Then
                    .IonSpotColor = lblIonColorSelection.BackColor
                    SetOKButtonCaption False
                End If
                If .IonSpotShape <> cboIonShape.ListIndex Then
                    .IonSpotShape = cboIonShape.ListIndex
                    SetOKButtonCaption False
                End If
                
                If .UMCSpotColor <> lblUMCColorSelection.BackColor Then
                    .UMCSpotColor = lblUMCColorSelection.BackColor
                    SetOKButtonCaption False
                End If
                If .UMCSpotShape <> cboUMCShape.ListIndex Then
                    .UMCSpotShape = cboUMCShape.ListIndex
                    SetOKButtonCaption False
                End If
                
                blnVisibleScopeOnly = cChkBox(chkVisibleScopeOnly)
                If .VisibleScopeOnly <> blnVisibleScopeOnly Then
                    .VisibleScopeOnly = blnVisibleScopeOnly
                    SetOKButtonCaption False
                End If
           End With
        End If
    Next intListIndex
End Sub

Private Sub UpdateGeneralOptions()
    If mInitializingControls Then Exit Sub
    
    With gOrfViewerOptionsCurrentGelList.DisplayOptions
        .MassTagMassErrorPPM = txtMassTagMassError
        ValidateValueDbl .MassTagMassErrorPPM, 0, 10000, DEFAULT_ORF_MASS_TAG_MASS_ERROR_PPM
        
        .MassTagNETError = txtMassTagNETError
        ValidateValueDbl .MassTagNETError, 0, 2, DEFAULT_ORF_MASS_TAG_NET_ERROR
        
        .PicturePixelHeight = CLngSafe(txtPictureHeight)
        ValidateValueLng .PicturePixelHeight, 25, 32000, DEFAULT_ORF_PICTURE_HEIGHT
        
        .PicturePixelWidth = CLngSafe(txtPictureWidth)
        ValidateValueLng .PicturePixelWidth, 25, 32000, DEFAULT_ORF_PICTURE_WIDTH
        
        .MaxSpotSizePixels = CLngSafe(txtMaxSpotSize)
        ValidateValueLng .MaxSpotSizePixels, 1, 10000, DEFAULT_ORF_MAX_SPOT_SIZE_PIXELS
        
        .MinSpotSizePixels = CLngSafe(txtMinSpotSize)
        ValidateValueLng .MinSpotSizePixels, 1, 10000, DEFAULT_ORF_MIN_SPOT_SIZE_PIXELS
        
        .IntensityScalar = CDblSafe(txtIntensityScalarForListView)
        ValidateValueDbl .IntensityScalar, 1, 1E+30, DEFAULT_ORF_LISTVIEW_INTENSITY_SCALAR
        
        .IonToUMCPlottingIntensityRatio = CDblSafe(txtIonToUMCScalingRatio)
        ValidateValueDbl .IonToUMCPlottingIntensityRatio, 0.001, 1000, DEFAULT_ORF_PICTURE_ION_TO_UMC_INTENSITY_SCALING_RATIO
        
        .MassTagSpotColor = lblMassTagColorSelection.BackColor
        .MassTagSpotShape = cboMassTagShape.ListIndex
        
        .ShowPosition = cChkBox(chkShowPosition)
        .ShowTickMarkLabels = cChkBox(chkShowTickMarkLabels)
        .ShowGridLines = cChkBox(chkShowGridlines)
        
        .ShowNonTrypticMassTagsWithoutIonHits = cChkBox(chkShowNonTrypticMassTagsWithoutIonHits)
        
        .OnlyUseTop50PctForAveraging = cChkBox(chkOnlyUseTop50PctForAveraging)
        
        .LogarithmicIntensityPlotting = cChkBox(chkLogarithmicIntensity)
        .UseClassRepresentativeNET = cChkBox(chkUseUMCClassRepresentativeNET)
        .SwapPlottingAxes = cChkBox(chkSwapAxes)
        
        .LoadPMTs = cChkBox(chkLoadPMTs)
        .IncludeUnobservedTrypticMassTags = cChkBox(chkAddUnobservedTrypticMassTags)
        
        .CleavageRuleID = cboCleavageRuleType.ListIndex
        
        .HideEmptyMassTagPictures = cChkBox(chkHideEmptyMassTagPics)
        
        ' Disable chkShowNonTrypticMassTagsWithoutIonHits when .HideEmptyMassTagPictures = True
        chkShowNonTrypticMassTagsWithoutIonHits.Enabled = Not .HideEmptyMassTagPictures
            
    End With
    
    SetOKButtonCaption False

End Sub

Private Sub cboCleavageRuleType_Click()
    UpdateGeneralOptions
End Sub

Private Sub cboMassTagShape_Click()
    If cboMassTagShape.ListIndex >= 0 Then
        UpdateGeneralOptions
    End If
End Sub

Private Sub cboNETAdjustmentType_Click()
    UpdateGeneralOptions
End Sub

Private Sub cboIonShape_Click()
    If cboIonShape.ListIndex >= 0 Then
        mLastSelectedSpotShape = cboIonShape.ListIndex
        UpdateOptionsForSelectedGels
    End If
End Sub

Private Sub cboUMCShape_Click()
    If cboUMCShape.ListIndex >= 0 Then
        UpdateOptionsForSelectedGels
    End If
End Sub

Private Sub chkAddUnobservedTrypticMassTags_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkHideEmptyMassTagPics_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkOnlyUseTop50PctForAveraging_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkShowNonTrypticMassTagsWithoutIonHits_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkUseUMCClassRepresentativeNET_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkLoadPMTs_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkLogarithmicIntensity_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkShowGridlines_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkShowPosition_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkShowTickMarkLabels_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkSwapAxes_Click()
    UpdateGeneralOptions
End Sub

Private Sub chkVisibleScopeOnly_Click()
    UpdateOptionsForSelectedGels
    
    ' Since user has selected to a different basis for the ions,
    '  will need to update the IonMass or UMCMass pointer arrays
    ' Setting mGelListChanged to True will accomplish this
    SetGelInListChanged True
End Sub

Private Sub cmdAdd_Click()
    AddSelectedGels False
End Sub

Private Sub cmdAddAll_Click()
    AddSelectedGels True
End Sub

Private Sub cmdApply_Click()
    ApplyChanges
End Sub

Private Sub cmdCancel_Click()
        
    gOrfViewerOptionsCurrentGelList = gOrfViewerOptionsSavedGelList
    ApplyChanges
    
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()
    If cmdOK.Caption = OK_BUTTON_CAPTION Then
        Me.Hide
        ApplyChanges
    End If
    
    Me.Hide
End Sub

Private Sub cmdRemove_Click()
    RemoveSelectedGels False
End Sub

Private Sub cmdRemoveAll_Click()
    RemoveSelectedGels True
End Sub

Private Sub cmdZOrderDown_Click()
    ShuffleGelZOrder False
End Sub

Private Sub cmdZOrderUp_Click()
    ShuffleGelZOrder True
End Sub

Private Sub Form_Load()
    PopulateComboBoxes
    
    mLastSelectedSpotShape = sCircle
    
    SetOKButtonCaption False
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub lblIonColorSelection_Click()
    SelectCustomColor lblIonColorSelection
    UpdateOptionsForSelectedGels
End Sub

Private Sub lblMassTagColorSelection_Click()
    SelectCustomColor lblMassTagColorSelection
    UpdateGeneralOptions
End Sub

Private Sub lblUMCColorSelection_Click()
    SelectCustomColor lblUMCColorSelection
    UpdateOptionsForSelectedGels
End Sub

Private Sub lstAvailableGels_DblClick()
    ' Automatically call AddSelectedGels
    AddSelectedGels False
End Sub

Private Sub lstGelsInUse_Click()
    ShowSettingsForSelectedGel
End Sub

Private Sub txtMinSpotSize_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtMinSpotSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMinSpotSize, KeyAscii, True, False
End Sub

Private Sub txtMinSpotSize_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtMinSpotSize, 1, 1000, DEFAULT_ORF_MIN_SPOT_SIZE_PIXELS
End Sub

Private Sub txtIntensityScalarForListView_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtIntensityScalarForListView_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIntensityScalarForListView, KeyAscii, True, False
End Sub

Private Sub txtIntensityScalarForListView_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtIntensityScalarForListView, 1, 1E+30, DEFAULT_ORF_LISTVIEW_INTENSITY_SCALAR
End Sub

Private Sub txtIonToUMCScalingRatio_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtIonToUMCScalingRatio_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIonToUMCScalingRatio, KeyAscii, True, True, False
End Sub

Private Sub txtIonToUMCScalingRatio_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtIonToUMCScalingRatio, 0.001, 1000, DEFAULT_ORF_PICTURE_ION_TO_UMC_INTENSITY_SCALING_RATIO
End Sub

Private Sub txtMassTagMassError_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtMassTagMassError_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMassTagMassError, KeyAscii, True, True, False
End Sub

Private Sub txtMassTagMassError_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtMassTagMassError, 0, 10000, DEFAULT_ORF_MASS_TAG_MASS_ERROR_PPM
End Sub

Private Sub txtMassTagNETError_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtMassTagNETError_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMassTagNETError, KeyAscii, True, True, False
End Sub

Private Sub txtMassTagNETError_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtMassTagNETError, 0, 2, DEFAULT_ORF_MASS_TAG_NET_ERROR
End Sub

Private Sub txtPictureHeight_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtPictureHeight_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPictureHeight, KeyAscii, True, False
End Sub

Private Sub txtPictureHeight_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtPictureHeight, 25, 32000, DEFAULT_ORF_PICTURE_HEIGHT
End Sub

Private Sub txtPictureWidth_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtPictureWidth_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtPictureWidth, KeyAscii, True, False
End Sub

Private Sub txtPictureWidth_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtPictureWidth, 25, 32000, DEFAULT_ORF_PICTURE_WIDTH
End Sub

Private Sub txtMaxSpotSize_Change()
    UpdateGeneralOptions
End Sub

Private Sub txtMaxSpotSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMaxSpotSize, KeyAscii, True, False
End Sub

Private Sub txtMaxSpotSize_Validate(Cancel As Boolean)
    ValidateTextboxValueLng txtMaxSpotSize, 1, 1000, DEFAULT_ORF_MAX_SPOT_SIZE_PIXELS
End Sub

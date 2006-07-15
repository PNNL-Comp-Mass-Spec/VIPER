VERSION 5.00
Begin VB.UserControl SelToolBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   ScaleHeight     =   2190
   ScaleWidth      =   3165
   ToolboxBitmap   =   "SelToolBox.ctx":0000
   Begin VB.Frame fraSelBox 
      Caption         =   "Selection Tools"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdRange 
         Caption         =   "Rng"
         Height          =   375
         Left            =   2540
         TabIndex        =   19
         ToolTipText     =   "Maximum value in range"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdAvg 
         Caption         =   "&Avg"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         ToolTipText     =   "Average"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdStDev 
         Caption         =   "&St D"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Standard Deviation"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMin 
         Caption         =   "Mi&n"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Minimum value in range"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMax 
         Caption         =   "Ma&x"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         ToolTipText     =   "Maximum value in range"
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optField 
         Caption         =   "&ER"
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Expression Ratio"
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optField 
         Caption         =   "&Fit"
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Calculated Fit"
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton optField 
         Caption         =   "In&t"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Intensity"
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optField 
         Caption         =   "&MW"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Molecular Mass"
         Top             =   600
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.Label lblAllRes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblIsoRes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblCSRes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblResTitleAll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   680
         TabIndex        =   15
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label lblResTitleIso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   680
         TabIndex        =   14
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label lblResTitleCS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   680
         TabIndex        =   13
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected CS:"
         Height          =   255
         Index           =   0
         Left            =   680
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblSelIsoCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblSelCSCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Iso:"
         Height          =   255
         Index           =   1
         Left            =   680
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "SelToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this is just shell for Selection Statistics toolbox
'last modified 08/27/2000 nt
Option Explicit

Const MyWidth = 3165
Const MyHeight = 2190

Public Enum DblClickPosition
    DblClickTL = 0
    DblClickBR = 1
End Enum

Dim mField As Integer
Dim mMoPoDbCl As DblClickPosition

Public Event ClickAvg()
Public Event ClickStD()
Public Event ClickMin()
Public Event ClickMax()
Public Event ClickRange()
Public Event DblClick(ByVal DblClickType As DblClickPosition)
Public Event Click()
Public Event MouseDown()
Public Event MouseUp()

Public Property Get BackColor() As OLE_COLOR
    BackColor = fraSelBox.BackColor
End Property

Public Property Let BackColor(ByVal lClr As OLE_COLOR)
    fraSelBox.BackColor = lClr
    PropertyChanged "BackColor"
End Property

Public Property Get SelCSCount() As Long
'this is not neccessary but it might come handy
    SelCSCount = lblSelCSCnt.Caption
End Property

Public Property Let SelCSCount(ByVal SelCSCnt As Long)
    lblSelCSCnt.Caption = SelCSCnt
    PropertyChanged "SelCSCount"
End Property

Public Property Get SelIsoCount() As Long
'this is not neccessary but it might come handy
    SelIsoCount = lblSelIsoCnt.Caption
End Property

Public Property Let SelIsoCount(ByVal SelIsoCnt As Long)
    lblSelIsoCnt.Caption = SelIsoCnt
    PropertyChanged "SelIsoCount"
End Property

Private Sub cmdAvg_Click()
    Select Case mField
    Case 0      'MW
         lblResTitleCS.Caption = "Avg.MW(CS):"
         lblResTitleIso.Caption = "Avg.MW(Iso):"
         lblResTitleAll.Caption = "Avg.MW(All):"
    Case 1      'Intensity
         lblResTitleCS.Caption = "Avg.Int.(CS):"
         lblResTitleIso.Caption = "Avg.Int.(Iso):"
         lblResTitleAll.Caption = "Avg.Int.(All):"
    Case 2      'Fit
         lblResTitleCS.Caption = "Avg.StDev.(CS):"
         lblResTitleIso.Caption = "Avg.Fit(Iso):"
         lblResTitleAll.Caption = "Avg.Fit(All):"
    Case 3      'ER
         lblResTitleCS.Caption = "Avg.ER(CS):"
         lblResTitleIso.Caption = "Avg.ER(Iso):"
         lblResTitleAll.Caption = "Avg.ER(All):"
    End Select
    RaiseEvent ClickAvg
End Sub

Private Sub cmdMax_Click()
    Select Case mField
    Case 0      'MW
         lblResTitleCS.Caption = "Max MW(CS):"
         lblResTitleIso.Caption = "Max MW(Iso):"
         lblResTitleAll.Caption = "Max MW(All):"
    Case 1      'Intensity
         lblResTitleCS.Caption = "Max Int.(CS):"
         lblResTitleIso.Caption = "Max Int.(Iso):"
         lblResTitleAll.Caption = "Max Int.(All):"
    Case 2      'Fit
         lblResTitleCS.Caption = "Max StDev.(CS):"
         lblResTitleIso.Caption = "Max Fit(Iso):"
         lblResTitleAll.Caption = "Max Fit(All):"
    Case 3      'ER
         lblResTitleCS.Caption = "Max ER(CS):"
         lblResTitleIso.Caption = "Max ER(Iso):"
         lblResTitleAll.Caption = "Max ER(All):"
    End Select
    RaiseEvent ClickMax
End Sub

Private Sub cmdMin_Click()
    Select Case mField
    Case 0      'MW
         lblResTitleCS.Caption = "Min MW(CS):"
         lblResTitleIso.Caption = "Min MW(Iso):"
         lblResTitleAll.Caption = "Min MW(All):"
    Case 1      'Intensity
         lblResTitleCS.Caption = "Min Int.(CS):"
         lblResTitleIso.Caption = "Min Int.(Iso):"
         lblResTitleAll.Caption = "Min Int.(All):"
    Case 2      'Fit
         lblResTitleCS.Caption = "Min StDev.(CS):"
         lblResTitleIso.Caption = "Min Fit(Iso):"
         lblResTitleAll.Caption = "Min Fit(All):"
    Case 3      'ER
         lblResTitleCS.Caption = "Min ER(CS):"
         lblResTitleIso.Caption = "Min ER(Iso):"
         lblResTitleAll.Caption = "Min ER(All):"
    End Select
    RaiseEvent ClickMin
End Sub

Private Sub cmdRange_Click()
    Select Case mField
    Case 0      'MW
         lblResTitleCS.Caption = "Range MW(CS):"
         lblResTitleIso.Caption = "Range MW(Iso):"
         lblResTitleAll.Caption = "Range MW(All):"
    Case 1      'Intensity
         lblResTitleCS.Caption = "Range Int.(CS):"
         lblResTitleIso.Caption = "Range Int.(Iso):"
         lblResTitleAll.Caption = "Range Int.(All):"
    Case 2      'Fit
         lblResTitleCS.Caption = "Range StDev.(CS):"
         lblResTitleIso.Caption = "Range Fit(Iso):"
         lblResTitleAll.Caption = "Range Fit(All):"
    Case 3      'ER
         lblResTitleCS.Caption = "Range ER(CS):"
         lblResTitleIso.Caption = "Range ER(Iso):"
         lblResTitleAll.Caption = "Range ER(All):"
    End Select
    RaiseEvent ClickRange
End Sub

Private Sub cmdStDev_Click()
    Select Case mField
    Case 0      'MW
         lblResTitleCS.Caption = "StDev.MW(CS):"
         lblResTitleIso.Caption = "StDev.MW(Iso):"
         lblResTitleAll.Caption = "StDev.MW(All):"
    Case 1      'Intensity
         lblResTitleCS.Caption = "StDev.Int.(CS):"
         lblResTitleIso.Caption = "StDev.Int.(Iso):"
         lblResTitleAll.Caption = "StDev.Int.(All):"
    Case 2      'Fit
         lblResTitleCS.Caption = "StDev.StDev.(CS):"
         lblResTitleIso.Caption = "StDev.Fit(Iso):"
         lblResTitleAll.Caption = "StDev.Fit(All):"
    Case 3      'ER
         lblResTitleCS.Caption = "StDev.ER(CS):"
         lblResTitleIso.Caption = "StDev.ER(Iso):"
         lblResTitleAll.Caption = "StDev.ER(All):"
    End Select
    RaiseEvent ClickStD
End Sub

Private Sub fraSelBox_Click()
    RaiseEvent Click
End Sub

Private Sub fraSelBox_DblClick()
    RaiseEvent DblClick(mMoPoDbCl)
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub fraSelBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown
End Sub

Private Sub fraSelBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < cmdAvg.Left And Y < optField(0).Top Then
       mMoPoDbCl = DblClickTL
    Else
       mMoPoDbCl = DblClickBR
    End If
End Sub

Private Sub fraSelBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp
End Sub

Private Sub optField_Click(Index As Integer)
    mField = Index
    ClearResults
End Sub

Private Sub UserControl_Initialize()
    mField = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    BackColor = PropBag.ReadProperty("BackColor")
End Sub

Private Sub UserControl_Resize()
UserControl.width = MyWidth
UserControl.Height = MyHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    PropBag.WriteProperty "BackColor", BackColor
End Sub

Public Property Get DataField() As Integer
    DataField = mField
End Property

Public Property Get ResCS() As String
    ResCS = lblCSRes.Caption
End Property

Public Property Let ResCS(ByVal Result As String)
    lblCSRes.Caption = Result
    lblCSRes.Visible = True
    lblResTitleCS.Visible = True
    PropertyChanged "ResCS"
End Property

Public Property Get ResIso() As String
    ResIso = lblIsoRes.Caption
End Property

Public Property Let ResIso(ByVal Result As String)
    lblIsoRes.Caption = Result
    lblIsoRes.Visible = True
    lblResTitleIso.Visible = True
    PropertyChanged "ResIso"
End Property

Public Property Get ResAll() As String
    ResAll = lblAllRes.Caption
End Property

Public Property Let ResAll(ByVal Result As String)
    lblAllRes.Caption = Result
    lblAllRes.Visible = True
    lblResTitleAll.Visible = True
    PropertyChanged "ResAll"
End Property

Public Sub ClearResults()
    lblCSRes.Visible = False
    lblIsoRes.Visible = False
    lblAllRes.Visible = False
    lblResTitleCS.Visible = False
    lblResTitleIso.Visible = False
    lblResTitleAll.Visible = False
End Sub

Public Sub SetResults(ByVal ResC As String, ByVal ResI As String, ByVal ResA As String)
'sets all three results in one procedure
    ResCS = ResC
    ResIso = ResI
    ResAll = ResA
End Sub

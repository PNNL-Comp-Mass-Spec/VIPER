VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraReturnOptions 
      Caption         =   "Return Options"
      Height          =   975
      Left            =   3480
      TabIndex        =   14
      Top             =   1560
      Width           =   1815
      Begin VB.CheckBox chkResSelect 
         Caption         =   "Select Fi&ndings"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkResZoom 
         Caption         =   "&Zoom to Results"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkFindSelection 
      Caption         =   "Find &Selection from active gel"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Frame fraTolerance 
      Caption         =   "Tolerance"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
      Begin VB.OptionButton optMMA 
         Caption         =   "&units"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   17
         ToolTipText     =   "absolute tolerance in units"
         Top             =   680
         Width           =   855
      End
      Begin VB.OptionButton optMMA 
         Caption         =   "per&cent"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   12
         ToolTipText     =   "relative tolerance in ""percent"""
         Top             =   440
         Width           =   915
      End
      Begin VB.OptionButton optMMA 
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   11
         ToolTipText     =   "relative tolerance in ""parts per million"""
         Top             =   200
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtTolerance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "10"
         Top             =   400
         Width           =   375
      End
   End
   Begin VB.Frame fraScope 
      Caption         =   "Look in"
      Height          =   975
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
      Begin VB.OptionButton optScope 
         Caption         =   "&All open gels"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optScope 
         Caption         =   "A&ctive gel only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   400
      Left            =   4080
      TabIndex        =   5
      Top             =   100
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   400
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cmbFindWhere 
      Height          =   315
      ItemData        =   "frmFind.frx":000C
      Left            =   1200
      List            =   "frmFind.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtFindWhat 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   640
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Find Where:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   700
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 05/03/2000 nt
Option Explicit

Dim CallerID As Long

Const RangeMark = ".."
Const ListMark = ";"

Dim FindWhere As Integer
Dim FindWhat As String
Dim FindArgType As Integer
Dim FindTolType As Integer
Dim FindTolerance As Double
Dim sngArg() As Double
Dim strArg() As String
Dim ArgCnt As Integer

Private Sub FillComboBoxes()
    With cmbFindWhere
        .Clear
        .AddItem "Molecular Mass"
        .AddItem "Expression Ratio"
        .AddItem "m/z"
        .AddItem "Charge State"
        .AddItem "Intensity"
        .AddItem "Identification (PMT Match)"
        .ListIndex = 0
    End With
End Sub

Private Sub chkFindSelection_Click()
If chkFindSelection.value = vbChecked Then
   txtFindWhat.Text = ""
   txtFindWhat.Enabled = False
   optScope(1).value = True
   optScope(0).Enabled = False
   FindArgType = glFIND_SELECTION
Else
   txtFindWhat.Enabled = True
   optScope(0).Enabled = True
End If
End Sub

Private Sub cmbFindWhere_Click()
FindWhere = cmbFindWhere.ListIndex
FindWhereTolerance
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Select Case FindWhere
Case glFIELD_MW, glFIELD_ABU, glFIELD_CS, glFIELD_ER
    If FindArgType = glFIND_SELECTION Then
       If GetSelectionFieldNumeric(CallerID, FindWhere, sngArg()) <= 0 Then
          MsgBox "No selection found in active gel.", vbOKOnly
          Exit Sub
       End If
       'look in all open gels except CallerID
       FindInAllNumeric False
    Else
       If Len(FindWhat) > 0 Then
          ParseArgumentsNumbers
       Else
          MsgBox "No arguments specified.", vbOKOnly
          Exit Sub
       End If
       If optScope(0).value Then   'look in active gel only
          FindInActiveNumeric
       Else     'look in all active gels(including CallerID)
          FindInAllNumeric True
       End If
    End If
Case glFIELD_ID  'Identity (text field)
    If FindArgType = glFIND_SELECTION Then
       If GetSelectionFieldMatchingIDs(CallerID, FindWhere, strArg()) <= 0 Then
          MsgBox "No selection found in active gel.", vbOKOnly
          Exit Sub
       End If
       'look in all open gels except CallerID
       FindInAllText False
    Else
       If Len(FindWhat) > 0 Then
          ParseArgumentsStrings
       Else
          MsgBox "No arguments specified.", vbOKOnly
          Exit Sub
       End If
       If optScope(0).value Then   'look in active gel only
          FindInActiveText
       Else     'look in all active gels(including CallerID)
          FindInAllText True
       End If
    End If
End Select
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag
End Sub

Private Sub Form_Load()
FillComboBoxes
Me.Move 100, 100
FindWhere = glFIELD_MW
FindTolType = gltPPM
FindTolerance = 10
If GetChildCount() > 1 Then
   optScope(1).Enabled = True
   chkFindSelection.Enabled = True
End If
End Sub

Private Sub FindWhereTolerance()
Select Case FindWhere
Case glFIELD_ID
    fraTolerance.Enabled = False
Case Else
    fraTolerance.Enabled = True
End Select
End Sub

Private Sub optMMA_Click(Index As Integer)
FindTolType = Index
End Sub

Private Sub txtFindWhat_LostFocus()
FindWhat = Trim$(txtFindWhat.Text)
End Sub

Private Sub ParseArgumentsStrings()
Dim CurrPos As Long
Dim EndPos As Long
Dim CurrStr As String

ReDim strArg(1 To 500)
ArgCnt = 0
If Len(FindWhat) > 0 Then
   CurrPos = InStr(1, FindWhat, RangeMark)
   If CurrPos > 0 Then
      CurrStr = Trim$(Left$(FindWhat, CurrPos - 1))
      If IsNumeric(CurrStr) Then
         strArg(1) = Trim$(CurrStr)
      Else
         GoTo exit_BadArguments
      End If
      CurrStr = Trim$(Right$(FindWhat, Len(FindWhat) - CurrPos - 1))
      If IsNumeric(CurrStr) Then
         sngArg(2) = Trim$(CurrStr)
      Else
         GoTo exit_BadArguments
      End If
      ArgCnt = 2
      FindArgType = glFIND_RANGE
   Else
      CurrPos = 1
      EndPos = 0
      Do While CurrPos <= Len(FindWhat)
         EndPos = InStr(CurrPos, FindWhat, ListMark)
         If EndPos > CurrPos Then
            CurrStr = Trim$(Mid$(FindWhat, CurrPos, EndPos - CurrPos))
            If Len(CurrStr) > 0 Then
               ArgCnt = ArgCnt + 1
               If ArgCnt <= 500 Then strArg(ArgCnt) = CurrStr
            End If
            CurrPos = EndPos + 1
         Else
            CurrStr = Trim$(Right$(FindWhat, Len(FindWhat) - CurrPos + 1))
            If Len(CurrStr) > 0 Then
               ArgCnt = ArgCnt + 1
               If ArgCnt <= 500 Then strArg(ArgCnt) = CurrStr
            End If
            CurrPos = Len(FindWhat) + 1
         End If
         If ArgCnt = 1 Then
            FindArgType = glFIND_VALUE
         Else
            FindArgType = glFIND_LIST
         End If
      Loop
   End If
   If ArgCnt > 0 Then
      ReDim Preserve strArg(1 To ArgCnt)
   Else
      GoTo exit_BadArguments
   End If
End If
Exit Sub

exit_BadArguments:
MsgBox "No valid arguments specified.", vbOKOnly
FindArgType = glFIND_NOTHING
txtFindWhat.SetFocus
End Sub

Private Sub ParseArgumentsNumbers()
Dim CurrPos As Long
Dim EndPos As Long
Dim CurrStr As String
Dim sngTmp  As Double

ReDim sngArg(1 To 500)
ArgCnt = 0
If Len(FindWhat) > 0 Then
   CurrPos = InStr(1, FindWhat, RangeMark)
   If CurrPos > 0 Then
      CurrStr = Trim$(Left$(FindWhat, CurrPos - 1))
      If IsNumeric(CurrStr) Then
         sngArg(1) = CDbl(CurrStr)
      ElseIf Len(CurrStr) = 0 Then
         sngArg(1) = 0
      Else
         GoTo exit_BadArguments
      End If
      CurrStr = Trim$(Right$(FindWhat, Len(FindWhat) - CurrPos - 1))
      If IsNumeric(CurrStr) Then
         sngArg(2) = CDbl(CurrStr)
      ElseIf Len(CurrStr) = 0 Then
         sngArg(2) = glHugeOverExp
      Else
         GoTo exit_BadArguments
      End If
      If sngArg(1) > sngArg(2) Then     'swap them and continue
         sngTmp = sngArg(1)
         sngArg(1) = sngArg(2)
         sngArg(2) = sngTmp
      End If
      ArgCnt = 2
      FindArgType = glFIND_RANGE
   Else
      CurrPos = 1
      EndPos = 0
      Do While CurrPos <= Len(FindWhat)
         EndPos = InStr(CurrPos, FindWhat, ListMark)
         If EndPos > CurrPos Then
            CurrStr = Trim$(Mid$(FindWhat, CurrPos, EndPos - CurrPos))
            If IsNumeric(CurrStr) Then
               ArgCnt = ArgCnt + 1
               If ArgCnt <= 500 Then sngArg(ArgCnt) = CDbl(CurrStr)
            End If
            CurrPos = EndPos + 1
         Else
            CurrStr = Trim$(Right$(FindWhat, Len(FindWhat) - CurrPos + 1))
            If IsNumeric(CurrStr) Then
               ArgCnt = ArgCnt + 1
               If ArgCnt <= 500 Then sngArg(ArgCnt) = CDbl(CurrStr)
            End If
            CurrPos = Len(FindWhat) + 1
         End If
         If ArgCnt = 1 Then
            FindArgType = glFIND_VALUE
         Else
            FindArgType = glFIND_LIST
         End If
      Loop
   End If
   If ArgCnt > 0 Then
      ReDim Preserve sngArg(1 To ArgCnt)
   Else
      MsgBox "No valid arguments specified.", vbOKOnly
      FindArgType = glFIND_NOTHING
      txtFindWhat.SetFocus
   End If
End If
Exit Sub

exit_BadArguments:
MsgBox "Arguments should be numeric.", vbOKOnly
FindArgType = glFIND_NOTHING
txtFindWhat.SetFocus
End Sub

Private Sub txtTolerance_LostFocus()
If IsNumeric(txtTolerance.Text) Then
   FindTolerance = CDbl(txtTolerance.Text)
Else
   MsgBox "Tolerance should be numeric value.", vbOKOnly
   txtTolerance.SetFocus
End If
End Sub


Private Sub FindInAllNumeric(ByVal DoMe As Boolean)
Dim i As Integer
Dim Results As String
Dim CurrCnt As Long
On Error Resume Next
Results = "Find Results:" & vbCrLf
If DoMe Then
   For i = 1 To UBound(GelData)
       If Not GelStatus(i).Deleted Then
          CurrCnt = FindNumeric(i, sngArg, FindWhere, FindTolerance, _
             FindTolType, FindArgType, chkResZoom.value, chkResSelect.value)
          Results = Results & GelBody(i).Caption & ": " & CurrCnt & vbCrLf
       End If
   Next i
Else
   For i = 1 To UBound(GelData)
       If Not GelStatus(i).Deleted And i <> CallerID Then
          CurrCnt = FindNumeric(i, sngArg, FindWhere, FindTolerance, _
             FindTolType, FindArgType, chkResZoom.value, chkResSelect.value)
          Results = Results & GelBody(i).Caption & ": " & CurrCnt & vbCrLf
       End If
   Next i
End If
MsgBox Results, vbOKOnly
End Sub

Private Sub FindInActiveNumeric()
Dim CurrCnt As Long
On Error Resume Next
CurrCnt = FindNumeric(CallerID, sngArg, FindWhere, FindTolerance, _
    FindTolType, FindArgType, chkResZoom.value, chkResSelect.value)
If CurrCnt > 0 Then
    'print result only if results will not be selected
   If chkResSelect.value <> vbChecked Then
      MsgBox CurrCnt & " data points found.", vbOKOnly
   End If
Else
   MsgBox "No data points satisfying Find criteria found.", vbOKOnly
End If
End Sub

Private Sub FindInActiveText()
Dim CurrCnt As Long
On Error Resume Next
CurrCnt = FindMatchingIDs(CallerID, strArg, FindWhere, FindArgType, _
                    chkResZoom.value, chkResSelect.value)
If CurrCnt > 0 Then
   'print result only if results will not be selected
   If chkResSelect.value <> vbChecked Then
      MsgBox CurrCnt & " data points found.", vbOKOnly
   End If
Else
   MsgBox "No data points satisfying Find criteria found.", vbOKOnly
End If
End Sub

Private Sub FindInAllText(ByVal DoMe As Boolean)
Dim i As Integer
Dim Results As String
Dim CurrCnt As Long
On Error Resume Next
Results = "Find Results:" & vbCrLf
If DoMe Then
   For i = 1 To UBound(GelData)
       If Not GelStatus(i).Deleted Then
          CurrCnt = FindMatchingIDs(i, strArg, FindWhere, FindArgType, _
                             chkResZoom.value, chkResSelect.value)
          Results = Results & GelBody(i).Caption & ": " & CurrCnt & vbCrLf
       End If
   Next i
Else
   For i = 1 To UBound(GelData)
       If Not GelStatus(i).Deleted And i <> CallerID Then
          CurrCnt = FindMatchingIDs(i, strArg, FindWhere, FindArgType, _
                             chkResZoom.value, chkResSelect.value)
          Results = Results & GelBody(i).Caption & ": " & CurrCnt & vbCrLf
       End If
   Next i
End If
MsgBox Results, vbOKOnly
End Sub

Private Sub FindInAllER(ByVal DoMe As Boolean)
Dim i As Integer
Dim Results As String
Dim CurrCnt As Long
On Error Resume Next
Results = "Find Results:" & vbCrLf
If DoMe Then
   For i = 1 To UBound(GelData)
       If Not GelStatus(i).Deleted Then
          CurrCnt = FindER(i, sngArg, FindTolerance, FindTolType, _
                FindArgType, chkResZoom.value, chkResSelect.value)
          Results = Results & GelBody(i).Caption & ": " & CurrCnt & vbCrLf
       End If
   Next i
Else
   For i = 1 To UBound(GelData)
       If Not GelStatus(i).Deleted And i <> CallerID Then
          CurrCnt = FindER(i, sngArg, FindTolerance, FindTolType, _
                FindArgType, chkResZoom.value, chkResSelect.value)
          Results = Results & GelBody(i).Caption & ": " & CurrCnt & vbCrLf
       End If
   Next i
End If
MsgBox Results, vbOKOnly
End Sub


Private Sub FindInActiveER()
Dim CurrCnt As Long
On Error Resume Next
CurrCnt = FindER(CallerID, sngArg, FindTolerance, FindTolType, _
            FindArgType, chkResZoom.value, chkResSelect.value)
If CurrCnt > 0 Then
   'print result only if results will not be selected
   If chkResSelect.value <> vbChecked Then
      MsgBox CurrCnt & " data points found.", vbOKOnly
   End If
Else
   MsgBox "No data points satisfying Find criteria found.", vbOKOnly
End If
End Sub


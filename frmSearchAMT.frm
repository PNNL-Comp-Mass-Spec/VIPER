VERSION 5.00
Begin VB.Form frmSearchAMT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search MT Database"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCheckMassTags 
      Caption         =   "C&heck MT Tags"
      Height          =   375
      Left            =   4920
      TabIndex        =   44
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CheckBox chkSaveNCnt 
      Caption         =   "Store N-atoms count with MT Tag reference"
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   6120
      Width           =   4215
   End
   Begin VB.Frame fraMassTag 
      Caption         =   "MT tag"
      Height          =   1335
      Left            =   4920
      TabIndex        =   38
      Top             =   2160
      Width           =   1455
      Begin VB.TextBox txtMTNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   41
         Text            =   "5"
         Top             =   920
         Width           =   735
      End
      Begin VB.TextBox txtMassTag 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Max."
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label6 
         Caption         =   "MT tag"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame fraAMTIncluded 
      Caption         =   "MT Tags Included in Search"
      Height          =   2055
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   4695
      Begin VB.TextBox txtFragMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3520
         TabIndex        =   32
         Text            =   "3"
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkAMTIncluded 
         Caption         =   "MT tags marked as N14/N15 pairs confirmed"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CheckBox chkAMTIncluded 
         Caption         =   "MT tags marked as high mass accuracy and NET condition"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   4525
      End
      Begin VB.CheckBox chkAMTIncluded 
         Caption         =   "MT tags marked as high mass accuracy"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox chkAMTIncluded 
         Caption         =   "All potential MT Tags(PMTs) found in database"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4335
      End
      Begin VB.CheckBox chkAMTIncluded 
         Caption         =   "MT tags marked as MS/MS confirmed (Min"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   1600
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "fragments)"
         Height          =   255
         Left            =   3860
         TabIndex        =   33
         Top             =   1695
         Width           =   800
      End
   End
   Begin VB.Frame fraNET 
      Caption         =   "NET/RT Calculation"
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   4695
      Begin VB.OptionButton optNETorRT 
         Caption         =   "Use RT"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   36
         ToolTipText     =   "Use Retention Time"
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optNETorRT 
         Caption         =   "Use NET"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "Use Normalized Elution Time"
         Top             =   440
         Width           =   975
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   540
         Width           =   3255
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Text            =   "0.1"
         Top             =   920
         Width           =   615
      End
      Begin VB.Label lblSec 
         Caption         =   "sec."
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "&Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   285
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "T&olerance"
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSearchResults 
      Caption         =   "S&earch Results"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "Display search statistic for MT database"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "D&elete MT Ref."
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "Remove MT reference for current gel"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CheckBox chkSkipAMTReferenced 
      Caption         =   "S&kip records already MT referenced"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5760
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Frame fraMWTolerance 
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1215
      Left            =   4200
      TabIndex        =   16
      Top             =   840
      Width           =   2175
      Begin VB.OptionButton optTolType 
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   160
         TabIndex        =   18
         Text            =   "10"
         Top             =   640
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Tolerance"
         Height          =   255
         Left            =   160
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraMWField 
      Caption         =   "Molecular Mass Field"
      Height          =   1215
      Left            =   2160
      TabIndex        =   12
      Top             =   840
      Width           =   1935
      Begin VB.OptionButton optMWField 
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   80
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   80
         TabIndex        =   14
         Top             =   540
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optMWField 
         Caption         =   "A&verage"
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraSearchScope 
      Caption         =   "Search Scope"
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1935
      Begin VB.OptionButton optSearchScope 
         Caption         =   "C&urrent View"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optSearchScope 
         Caption         =   "&All Data Points"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      ToolTipText     =   "Refresh connection to the MT database"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblAMTStatus 
      Caption         =   "Loading..."
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblAMTPath 
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "MTDB Path:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Path to the MT Tag database"
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmSearchAMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 10/01/2001 nt
Option Explicit

Dim CallerID As Long
Dim MinFrag As Integer
Dim Settings As Boolean

Private Sub chkAMTIncluded_Click(Index As Integer)
If Not Settings Then
   If chkAMTIncluded(Index).Value = vbChecked Then
      If Index = 0 Then
         samtDef.SearchFlag = glAMT_CONFIRM_NO
      Else
         Select Case Index
         Case 0
            samtDef.SearchFlag = samtDef.SearchFlag + glAMT_CONFIRM_NO
         Case 1
            samtDef.SearchFlag = samtDef.SearchFlag + glAMT_CONFIRM_PPM
         Case 2
            samtDef.SearchFlag = samtDef.SearchFlag + glAMT_CONFIRM_PPM_NET
         Case 3
            samtDef.SearchFlag = samtDef.SearchFlag + glAMT_CONFIRM_N14_N15
         Case 4
            Select Case MinFrag
            Case 1
                samtDef.SearchFlag = samtDef.SearchFlag + glAMT_CONFIRM_MSMS1
            Case 2
                samtDef.SearchFlag = samtDef.SearchFlag + glAMT_CONFIRM_MSMS2
            Case 3
                samtDef.SearchFlag = samtDef.SearchFlag + glAMT_CONFIRM_MSMS3PLUS
            End Select
         End Select
      End If
   Else
      If Index > 0 Then
         Select Case Index
         Case 0
            samtDef.SearchFlag = samtDef.SearchFlag - glAMT_CONFIRM_NO
         Case 1
            samtDef.SearchFlag = samtDef.SearchFlag - glAMT_CONFIRM_PPM
         Case 2
            samtDef.SearchFlag = samtDef.SearchFlag - glAMT_CONFIRM_PPM_NET
         Case 3
            samtDef.SearchFlag = samtDef.SearchFlag - glAMT_CONFIRM_N14_N15
         Case 4
            Select Case MinFrag
            Case 1
                samtDef.SearchFlag = samtDef.SearchFlag - glAMT_CONFIRM_MSMS1
            Case 2
                samtDef.SearchFlag = samtDef.SearchFlag - glAMT_CONFIRM_MSMS2
            Case 3
                samtDef.SearchFlag = samtDef.SearchFlag - glAMT_CONFIRM_MSMS3PLUS
            End Select
         End Select
      End If
   End If
   IncludeAMTs
End If
End Sub

Private Sub chkSaveNCnt_Click()
If chkSaveNCnt.Value = vbChecked Then
   samtDef.SaveNCnt = True
Else
   samtDef.SaveNCnt = False
End If
End Sub

Private Sub chkSkipAMTReferenced_Click()
If chkSkipAMTReferenced.Value = vbChecked Then
   samtDef.SkipReferenced = True
Else
   samtDef.SkipReferenced = False
End If
End Sub

Private Sub cmdCheckMassTags_Click()
'----------------------------------------------
'displays short MT tags statistics, it might
'help with determining problems with MT tags
'----------------------------------------------
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly
Me.MousePointer = vbDefault
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRefresh_Click()
If Not dbAMT Is Nothing Then Set dbAMT = Nothing
If ConnectToLegacyAMTDB(Me, CallerID, False, True, False) Then
   lblAMTStatus.Caption = GetConnectedStatus(True)
Else
   lblAMTStatus.Caption = GetConnectedStatus(False)
   cmdSearch.Enabled = False
End If
End Sub

Private Sub cmdRemove_Click()
Dim Respond
Respond = MsgBox("Remove MT tag references from the current gel?", vbYesNo + vbDefaultButton2)
If Respond <> vbYes Then Exit Sub
RemoveAMT CallerID, samtDef.SearchScope
GelStatus(CallerID).Dirty = True
AddToAnalysisHistory CallerID, "Deleted MT tag search results from ions"
End Sub

Private Sub cmdSearch_Click()
Dim HitsCnt As Long
On Error Resume Next

'always reinitialize statistics arrays
InitAMTStat
samtDef.Formula = Trim$(txtNETFormula.Text)
Me.MousePointer = vbHourglass
If samtDef.MassTag > 0 Then
   'HitsCnt = SearchAMTWithTag(CallerID, samtDef.Formula)
   HitsCnt = SearchAMTWithTag1(CallerID, samtDef.Formula)
Else
' MonroeMod
   HitsCnt = SearchAMT(CallerID, samtDef.Formula, Me)
End If
GelStatus(CallerID).Dirty = True

'MonroeMod
GelSearchDef(CallerID).AMTSearchOnIons = samtDef
AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched ions for MT tags", HitsCnt, 0, 0, 0, samtDef, False, GelData(CallerID).CustomNETsDefined)

Me.MousePointer = vbDefault
Select Case HitsCnt
Case -1
   MsgBox "Error searching MT database.", vbOKOnly
Case -2
   MsgBox "Error in NET/RT calculation formula.", vbOKOnly
   txtNETFormula.SetFocus
Case Else
   MsgBox "MT records hits: " & HitsCnt & "(non-unique)", vbOKOnly
End Select
End Sub

Private Sub cmdSearchResults_Click()
Dim FileNum As Integer
Dim FileNam As String
Dim sLine As String
Dim AvgErrDa As String
Dim AvgErrppm As String
Dim AvgErrNET As String
Dim i As Long
On Error GoTo exit_cmdSearchResults_Click

If AMTCnt > 0 Then  'global AMT count
   Me.MousePointer = vbHourglass
   FileNum = FreeFile()
   FileNam = GetTempFolder() & RawDataTmpFile
   Open FileNam For Output As FileNum
   'print gel file name and Search definition as reference
   Print #FileNum, "Generated by: " & GetMyNameVersion() & " on " & Now()
   Print #FileNum, "Gel File: " & GelBody(CallerID).Caption
   Print #FileNum, GetAMTSearchDefDesc()
   Select Case samtDef.NETorRT
   Case glAMT_NET
      sLine = "MT ID" & glARG_SEP & "MT MW" & glARG_SEP & "MT NET" & glARG_SEP _
           & "Hits" & glARG_SEP & "MT MW Avg Error(Da)" & glARG_SEP & _
           "MT MW Avg Error(ppm)" & glARG_SEP & "MT NET Avg Error" & _
           glARG_SEP & "MT Range Min" & glARG_SEP & "MT Range Max"
   Case glAMT_RT_or_PNET
      sLine = "MT ID" & glARG_SEP & "MT MW" & glARG_SEP & "MT RT" & glARG_SEP _
           & "Hits" & glARG_SEP & "MT MW Avg Error(Da)" & glARG_SEP & _
           "MT MW Avg Error(ppm)" & glARG_SEP & "MT RT Avg Error" & _
           glARG_SEP & "MT Range Min" & glARG_SEP & "MT Range Max"
   End Select
   Print #FileNum, sLine
   For i = 1 To AMTCnt
       Select Case samtDef.NETorRT
       Case glAMT_NET
         sLine = Trim(AMTData(i).ID) & glARG_SEP & AMTData(i).MW & glARG_SEP & AMTData(i).NET & glARG_SEP & AMTHits(i)
       Case glAMT_RT_or_PNET
         sLine = Trim(AMTData(i).ID) & glARG_SEP & AMTData(i).MW & glARG_SEP & AMTData(i).PNET & glARG_SEP & AMTHits(i)
       End Select
       If AMTHits(i) > 0 Then
          AvgErrDa = Str(AMTMWErr(i) / AMTHits(i))
          AvgErrppm = Str(AvgErrDa / (AMTData(i).MW * glPPM))
          AvgErrNET = Str(AMTNETErr(i) / AMTHits(i))
          sLine = sLine & glARG_SEP & AvgErrDa & glARG_SEP & AvgErrppm & glARG_SEP _
              & AvgErrNET & glARG_SEP & AMTNETMin(i) & glARG_SEP & AMTNETMax(i)
       End If
       Print #FileNum, sLine
   Next i
   Close FileNum
   DoEvents
   frmDataInfo.Tag = "AMT"
   DoEvents
   frmDataInfo.Show vbModal
Else
   MsgBox "No loaded MT data found.", vbOKOnly
End If

exit_cmdSearchResults_Click:
On Error Resume Next
Close FileNum
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
CallerID = Me.Tag

' MonroeMod
If CallerID >= 1 And CallerID <= UBound(GelBody) Then samtDef = GelSearchDef(CallerID).AMTSearchOnIons

' MonroeMod: This code was in Form_Load, but was moved here so that the search def
'            gets updated when the form is activated
'set current Search Definition values
With samtDef
    txtMWTol.Text = .MWTol
    txtNETFormula.Text = .Formula
    optSearchScope(.SearchScope).Value = True
    optMWField(.MWField - MW_FIELD_OFFSET).Value = True
    Select Case .TolType
    Case gltPPM
      optTolType(0).Value = True
    Case gltABS
      optTolType(1).Value = True
    Case Else
      Debug.Assert False
    End Select
    'check do we have retention time in this database
    If AMTGeneration < dbgGeneration0800 Then .NETorRT = glAMT_NET
    optNETorRT(.NETorRT).Value = True
    If .NETorRT = glAMT_NET Then
       lblSec.Visible = False
    Else
       lblSec.Visible = True
    End If
    If AMTGeneration < dbgGeneration0800 Then optNETorRT(glAMT_RT_or_PNET).Enabled = False
    'NETTol is used both for NET and RT
    If .NETTol >= 0 Then
       txtNETTol.Text = .NETTol
    Else
       txtNETTol.Text = ""
    End If
    If .MassTag > 0 Then
       txtMassTag.Text = .MassTag
    Else
       txtMassTag.Text = ""
    End If
    txtMTNum.Text = .MaxMassTags
    If .SkipReferenced Then
       chkSkipAMTReferenced.Value = vbChecked
    Else
       chkSkipAMTReferenced.Value = vbUnchecked
    End If
    If AMTGeneration < dbgGeneration0900 Then    'disable N atoms
       .SaveNCnt = False                            'count option
       chkSaveNCnt.Enabled = False                  'if db too old
    End If
    If .SaveNCnt Then
       chkSaveNCnt.Value = vbChecked
    Else
       chkSaveNCnt.Value = vbUnchecked
    End If
End With
SetMinFrag
IncludeAMTs

End Sub

Private Sub Form_Load()
If IsWinLoaded(TrackerCaption) Then Unload frmTracker

lblAMTPath.Caption = GelData(CallerID).PathtoDatabase
If dbAMT Is Nothing Then
   If ConnectToLegacyAMTDB(Me, CallerID, True, True, False) Then
      lblAMTStatus.Caption = GetConnectedStatus(True)
   Else
      lblAMTStatus.Caption = GetConnectedStatus(False)
      cmdSearch.Enabled = False
   End If
Else
   If GelData(CallerID).PathtoDatabase <> dbAMT.Name Then   'close current AMT and
      Set dbAMT = Nothing           'connect to new if appropriate
      If Len(GelData(CallerID).PathtoDatabase) > 0 Then
         If ConnectToLegacyAMTDB(Me, CallerID, True, True, False) Then
            lblAMTStatus.Caption = GetConnectedStatus(True)
         Else
            lblAMTStatus.Caption = GetConnectedStatus(False)
            cmdSearch.Enabled = False
         End If
      End If
   End If
   lblAMTStatus.Caption = GetConnectedStatus(True)
End If

' MonroeMod: The code that was here is now in Form_Activate

End Sub

Private Function GetConnectedStatus(ByVal AMTConnected As Boolean) As String
Dim sTmp As String
If AMTConnected Then
   sTmp = "Connected; MT Count: " & AMTCnt & "; Protein Count: " & ORFCnt
Else
   sTmp = "Not connected;"
End If
GetConnectedStatus = sTmp
End Function

Private Sub Form_Unload(Cancel As Integer)
'free some memory
DestroyAMTStat
End Sub

Private Sub optMWField_Click(Index As Integer)
samtDef.MWField = 6 + Index
End Sub

Private Sub optNETorRT_Click(Index As Integer)
samtDef.NETorRT = Index
lblSec.Visible = CBool(Index)
End Sub

Private Sub optSearchScope_Click(Index As Integer)
samtDef.SearchScope = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   samtDef.TolType = gltPPM
Else
   samtDef.TolType = gltABS
End If
End Sub

Private Sub txtFragMin_LostFocus()
If IsNumeric(txtFragMin.Text) Then
   Select Case txtFragMin.Text
   Case 1, 2, 3
        With samtDef
          If .SearchFlag >= glAMT_CONFIRM_MSMS1 Then
            If .SearchFlag >= glAMT_CONFIRM_MSMS2 Then
              If .SearchFlag >= glAMT_CONFIRM_MSMS3PLUS Then
                 .SearchFlag = .SearchFlag - glAMT_CONFIRM_MSMS3PLUS
              Else
                 .SearchFlag = .SearchFlag - glAMT_CONFIRM_MSMS2
              End If
            Else
              .SearchFlag = .SearchFlag - glAMT_CONFIRM_MSMS1
            End If
          End If
          MinFrag = txtFragMin.Text
          If chkAMTIncluded(4).Value = vbChecked Then
            Select Case MinFrag
            Case 1
                 .SearchFlag = .SearchFlag + glAMT_CONFIRM_MSMS1
            Case 2
                 .SearchFlag = .SearchFlag + glAMT_CONFIRM_MSMS2
            Case 3
                 .SearchFlag = .SearchFlag + glAMT_CONFIRM_MSMS3PLUS
            End Select
          End If
        End With
   Case Else
        MsgBox "This parameter can be only 1,2 or 3.", vbOKOnly
        txtFragMin.SetFocus
   End Select
Else
   MsgBox "This parameter can be only 1,2 or 3.", vbOKOnly
   txtFragMin.SetFocus
End If
End Sub

Private Sub txtMassTag_LostFocus()
Dim dTmp As Double
If IsNumeric(txtMassTag.Text) Then
   dTmp = CDbl(txtMassTag.Text)
   If dTmp > 0 Then
      samtDef.MassTag = dTmp
   Else
      MsgBox "MT tag ID should be positive number.", vbOKOnly
      txtMassTag.Text = ""
      txtMassTag.SetFocus
   End If
Else
   If Len(Trim(txtMassTag.Text)) > 0 Then
      MsgBox "MT tag ID should be positive number.", vbOKOnly
      txtMassTag.SetFocus
   Else
      samtDef.MassTag = -1
   End If
End If
End Sub

Private Sub txtMTNum_LostFocus()
On Error Resume Next
samtDef.MaxMassTags = CLng(Abs(txtMTNum.Text))
If Err Then
   MsgBox "Maximum number of MT tags should be positive integer.", vbOKOnly
   txtMTNum.SetFocus
End If
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   samtDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "Molecular Mass Tolerance should be numeric value.", vbOKOnly
   txtMWTol.SetFocus
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

Private Sub IncludeAMTs()
Dim TmpCalc As Integer
Settings = True

Select Case samtDef.SearchFlag
Case glAMT_CONFIRM_NO
     chkAMTIncluded(0).Value = vbChecked
Case Else
     chkAMTIncluded(0).Value = vbUnchecked
End Select

If samtDef.SearchFlag Mod 2 > 0 Then
     chkAMTIncluded(1).Value = vbChecked
Else
     chkAMTIncluded(1).Value = vbUnchecked
End If

If samtDef.SearchFlag >= glAMT_CONFIRM_PPM_NET Then
   TmpCalc = samtDef.SearchFlag Mod (2 * glAMT_CONFIRM_PPM_NET)
   If TmpCalc > 0 Then
      Select Case TmpCalc
      Case glAMT_CONFIRM_PPM
        chkAMTIncluded(2).Value = vbUnchecked
      Case Else
        chkAMTIncluded(2).Value = vbChecked
      End Select
   Else
      chkAMTIncluded(2).Value = vbUnchecked
   End If
Else
   chkAMTIncluded(2).Value = vbUnchecked
End If

If samtDef.SearchFlag >= glAMT_CONFIRM_N14_N15 Then
   TmpCalc = samtDef.SearchFlag Mod (2 * glAMT_CONFIRM_N14_N15)
   If TmpCalc > 0 Then
      Select Case TmpCalc
      Case glAMT_CONFIRM_PPM, glAMT_CONFIRM_PPM_NET, _
           glAMT_CONFIRM_PPM + glAMT_CONFIRM_PPM_NET
        chkAMTIncluded(3).Value = vbUnchecked
      Case Else
        chkAMTIncluded(3).Value = vbChecked
      End Select
   Else
      chkAMTIncluded(3).Value = vbUnchecked
   End If
Else
   chkAMTIncluded(3).Value = vbUnchecked
End If

If samtDef.SearchFlag >= glAMT_CONFIRM_MSMS1 Then
   SetMinFrag
   chkAMTIncluded(4).Value = vbChecked
Else
   chkAMTIncluded(4).Value = vbUnchecked
End If
Settings = False
End Sub

Private Sub SetMinFrag()
If samtDef.SearchFlag >= glAMT_CONFIRM_MSMS1 Then
   If samtDef.SearchFlag >= glAMT_CONFIRM_MSMS2 Then
      If samtDef.SearchFlag >= glAMT_CONFIRM_MSMS3PLUS Then
         MinFrag = 3
      Else
         MinFrag = 2
      End If
   Else
      MinFrag = 1
   End If
Else
   MinFrag = 3
End If
txtFragMin.Text = MinFrag
End Sub

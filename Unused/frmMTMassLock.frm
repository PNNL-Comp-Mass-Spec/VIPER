VERSION 5.00
Begin VB.Form frmMTLockMass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRISM Lock Mass Function"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "frmMTMassLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelectLockersSet 
      Caption         =   "Lockers S&et"
      Height          =   375
      Left            =   1440
      TabIndex        =   38
      ToolTipText     =   "Select set of lockers to use with lock procedure"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdListLockers 
      Caption         =   "List L&oaded"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraLockMass 
      Caption         =   "Lock Mass"
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   7215
      Begin VB.CommandButton cmdReport 
         Caption         =   "Re&port"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   27
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "&Lock Mass"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   26
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelectLockers 
         Caption         =   "Select L&ockers "
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Frame fraFreqShiftCalculation 
         Caption         =   "Frequency Shift Calculation"
         Height          =   1335
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   4215
         Begin VB.ComboBox cmbBestClassScore 
            Height          =   315
            ItemData        =   "frmMTMassLock.frx":000C
            Left            =   2400
            List            =   "frmMTMassLock.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   660
            Width           =   1695
         End
         Begin VB.OptionButton optFSCalculation 
            Caption         =   "Best Locker Score"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "Best Locker Score From Mass Tag Database"
            Top             =   960
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optFSCalculation 
            Caption         =   "Best Class Score On"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optFSCalculation 
            Caption         =   "Best Agreement"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   21
            ToolTipText     =   "Best Agreement Among Frequency Shifts"
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton optFSCalculation 
            Caption         =   "Average Frequency Shift"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            ToolTipText     =   "Average Frequency Shifts For All Lockers"
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame fraGrouping 
         Caption         =   "Grouping"
         Height          =   1335
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton optGrouping 
            Caption         =   "&Scan Segments"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton optGrouping 
            Caption         =   "&Unique Mass Classes"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optGrouping 
            Caption         =   "In&dividual Scans"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Label lblCalEquation 
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   2415
      End
   End
   Begin VB.Frame fraLockersSearch 
      Caption         =   "Search for Lockers Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7215
      Begin VB.CommandButton cmdgaNETElutionFormula 
         Caption         =   "Set &gaNET"
         Height          =   375
         Left            =   3480
         TabIndex        =   39
         ToolTipText     =   "Brings NET formula from the Mass Tags database"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDBElutionFormula 
         Caption         =   "Set TIC &ET"
         Height          =   375
         Left            =   4560
         TabIndex        =   31
         ToolTipText     =   "Brings NET formula from the Mass Tags database"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbElutionFormula 
         Height          =   315
         ItemData        =   "frmMTMassLock.frx":0010
         Left            =   1680
         List            =   "frmMTMassLock.frx":0012
         TabIndex        =   28
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmdSearchResults 
         Caption         =   "Search &Results"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbET 
         Height          =   315
         ItemData        =   "frmMTMassLock.frx":0014
         Left            =   240
         List            =   "frmMTMassLock.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Text            =   "0.2"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtMMA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   "25"
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Elution Type"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Elution Calculation Formula"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Elution Tolerance"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "MMA(ppm)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdLoadLockers 
      Caption         =   "Load &Lockers"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   2295
      ScaleWidth      =   5055
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TIC Intercept:"
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   37
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TIC Slope:"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   36
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TIC Fit:"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   35
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblTICIntercept 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6120
      TabIndex        =   34
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTICSlope 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4920
      TabIndex        =   33
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTICFit 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4080
      TabIndex        =   32
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   5655
   End
End
Attribute VB_Name = "frmMTLockMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'locking masses on Mass Tags database Lockers;
'created: 06/27/2001 nt
'last modified: 07/09/2002 nt
'---------------------------------------------------------
Option Explicit

'grouping constants
Const GR_SOLO = 0
Const GR_UMC = 1
Const GR_SEGMENTS = 2

'elution type constants
Const NET_COL = 0
Const RET_COL = 1
Const OET_COL = 2

'frequency shift calculation constants
Const FS_AVERAGE = 0
Const FS_BEST_WITH_REST = 1
Const FS_BEST_UMC_SCORE = 2
Const FS_BEST_LCK_SCORE = 3

'class scoring types (other than best locker score)
Const UMC_SCORE_ABUNDANCE = 0       'highest intensity
'''Const UMC_SCORE_FIT = 1             'best fit(lowest)
'''Const UMC_SCORE_LCK_MMA = 2         'lockers MMA(this one does not have a lot of sense)
'''Const UMC_SCORE_LCK_CNT = 3         'lockers count in class

'lockers retrieval command (stored procedure name)
Dim GetLockersCommand As String

'assumption is that lockers information will be loaded with
'fields in this order and sorted ascending on monoisotopic mass
Dim LckCnt As Long              'count of loaded lockers
Dim LckID() As Long             'field 0
Dim LckSeq() As String          'field 1
Dim LckName() As String         'field 2
Dim LckMW() As Double           'field 3
Dim LckET() As Double           'fields 4,5,6
Dim LckScore() As Double        'field 7




'statistic for each locker of matches made;
Dim LckHits() As Long
Dim LckMWErr() As Double     'sum of absolute values of absolute errors
'following arrays are used for both NET and RT calculation
Dim LckETErr() As Double    'sum of absolute ET errors (direction could help)
Dim LckETMin() As Double    'min of ET range
Dim LckETMax() As Double    'max of ET range

Dim CallerID As Long
Dim bLoading As Boolean
'parameters of locking; will become public properties later
Public ElutionFormula As String        'bounded to cmbElutionFormula
Public ElutionType As Long             'bounded to cmbET
Public FreqShiftCalcType As Long       'bounded to optFSCalculation
Public ClassBestScoreType As Long      'bounded to cmbClassBestScore
Public GroupingType As Long            'bounded to optGrouping

Public MMA As Double                   'mass measurement accuracy
Public ETTol As Double                 'elution tolerance

Const FN_LM_NOCANDIDATES = -1       'no lockers for scan
Const FN_LM_ERROR = -2              'error locking scan

'scan number limits for current scope
Dim MinFN As Integer
Dim MaxFN As Integer

'statistic variant array that describes UMC
'changes here are due to changes in UMC structure
Dim ClsStat() As Double
Dim ClsCnt As Long          'count of UMC
Dim ClsScore() As Double    'score for the class based on user selection
                            'score based on intensity(average), fit(average)
                            'best locker hit or highest number of locker hits
Dim ClsLck() As Long        'LockerID that best describes class
Dim ClsHits() As Long       'class hits of best AMT

Dim FNCls() As Long         'class that will lock scan(indexed from MinFN to MaxFN)
Dim FNClsLM() As Long       'Isotopic point that will lock this scan
'lock masses array-index in MWInd array of mass that locked scan,
'-1 if no lock mass in scan, -2 if error while locking this scan
'0 if no lock mass attempted on this scan
Dim FNLM() As Long
'frequency shifts array(indexed from MinFN to MaxFN)
Dim FNS() As Double
'ID of locker locking this scan; 0 if none; -1 if other method
Dim FNLckID() As Long
'scan number statistic arrays(indexed from MinFN to MaxFN)
Dim FNCnt() As Long      'total number of points in a scan
Dim FNIndMin() As Long   'first index of mw in a scan
Dim FNIndMax() As Long   'last index of mw in a scan

Dim Cal As Object           'calibration recalculation object

'object used to fast locate index ranges in AMTData().MW
Dim mwutSearch As MWUtil
'counts number of hits to the locker data (non-unique)
Dim HitsCount As Long
'Expression Evaluator variables
Dim MyExprEva As ExprEvaluator
Dim VarVals() As Long
Dim LMWork As LMDataWorking

Private Sub cmbBestClassScore_Click()
ClassBestScoreType = cmbBestClassScore.ItemData(cmbBestClassScore.ListIndex)
End Sub

Private Sub cmbElutionFormula_LostFocus()
ElutionFormula = cmbElutionFormula.Text
End Sub

Private Sub cmbET_Change()
ElutionType = cmbET.ItemData(cmbET.ListIndex)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDBElutionFormula_Click()
With GelAnalysis(CallerID)
  If .NET_Intercept <> 0 Or .NET_Slope <> 0 Then
     cmbElutionFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
  Else
     MsgBox "Invalid parameters of database elution formula!", vbOKOnly, glFGTU
  End If
End With
End Sub

Private Sub cmdgaNETElutionFormula_Click()
With GelAnalysis(CallerID)
  If .GANET_Intercept <> 0 Or .GANET_Slope <> 0 Then
     cmbElutionFormula.Text = ConstructNETFormula(.GANET_Slope, .GANET_Intercept)
  Else
     MsgBox "Invalid parameters of gaNET elution formula!", vbOKOnly, glFGTU
  End If
End With
End Sub

Private Sub cmdListLockers_Click()
'-----------------------------------------
'displays list of lockers
'-----------------------------------------
Dim FileNum As Integer
Dim FileNam As String
Dim sLine As String
Dim Response As Long
Dim i As Long
On Error GoTo exit_ListLockers

If LckCnt > 0 Then
   If LckCnt > 1000 Then
      Response = MsgBox("There are " & LckCnt & " locker records! Continue with display?", vbYesNo)
      If Response <> vbYes Then Exit Sub
   End If
   Me.MousePointer = vbHourglass
   FileNum = FreeFile()
   FileNam = GetTempFolder() & RawDataTmpFile
   Open FileNam For Output As FileNum
   'print gel file name and Search definition as reference
   Print #FileNum, "Generated by: " & GetMyNameVersion() & " on " & Now()
   Print #FileNum, "Organism Database: " & GelAnalysis(CallerID).Organism_DB_Name
   sLine = "ID" & glARG_SEP & "Peptide" & glARG_SEP & "Name" & glARG_SEP _
           & "MW" & glARG_SEP & "NET" & glARG_SEP & "RET" & glARG_SEP _
           & "OET" & glARG_SEP & "Score"
   Print #FileNum, sLine
   For i = 1 To LckCnt
       sLine = LckID(i) & glARG_SEP & LckSeq(i) & glARG_SEP & LckName(i) _
               & glARG_SEP & LckMW(i) & glARG_SEP & LckET(i, 0) & glARG_SEP _
               & LckET(i, 1) & glARG_SEP & LckET(i, 2) & glARG_SEP & LckScore(i)
       Print #FileNum, sLine
   Next i
   Close FileNum
   DoEvents
   frmDataInfo.Tag = "LOCKERS"
   DoEvents
   frmDataInfo.Show vbModal
Else
   MsgBox "No lockers records loaded!", vbOKOnly
End If

exit_ListLockers:
On Error Resume Next
Close FileNum
Me.MousePointer = vbDefault
End Sub

Private Sub cmdLoadLockers_Click()
'-------------------------------------------------
'executes command that retrieves lockers data; and
'loads rows to temporary arrays
'-------------------------------------------------
Dim LckNET() As Double          'field 4
Dim LckRET() As Double          'field 5
Dim LckOET() As Double          'field 6
Dim cnNew As New ADODB.Connection
Dim rsLockers As New ADODB.Recordset
Dim cmdGetLockers As New ADODB.Command
Dim prmArg As New ADODB.Parameter
Dim ErrCnt As Long                          'list only first 10 errors
Dim i As Long
Dim tmpLckTypeID As String
Dim tmpLckMinScore As String
Dim tmpLckCallerID As String
Dim ArgLine As String
On Error Resume Next
'retrieve settings which lockers should be loaded
With GelAnalysis(CallerID).MTDB
   tmpLckTypeID = CLng(.DBStuff(NAME_LOCKERS_TYPE).value)
   If Err Then
      tmpLckTypeID = "1"        'all
      Err.Clear
   End If
   tmpLckMinScore = .DBStuff(NAME_LOCKERS_MIN_SCORE).value
   If Err Then
      tmpLckMinScore = "-1E+308"
      Err.Clear
   End If
   tmpLckCallerID = .DBStuff(NAME_LOCKERS_CALLER_ID).value
   If Err Then
      tmpLckCallerID = ""
      Err.Clear
   End If
End With
'pass all argument as delimited string of arguments
ArgLine = tmpLckTypeID & DELI & tmpLckMinScore & DELI & tmpLckCallerID

On Error GoTo err_cmdLoadLockers
'reserve space for 1000 lockers; increase in chunks of 200 after that
ReDim LckID(1 To 1000)
ReDim LckSeq(1 To 1000)
ReDim LckName(1 To 1000)
ReDim LckMW(1 To 1000)
ReDim LckNET(1 To 1000)
ReDim LckRET(1 To 1000)
ReDim LckOET(1 To 1000)
ReDim LckScore(1 To 1000)

Screen.MousePointer = vbHourglass
LckCnt = 0

If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
    Debug.Assert False
    Exit Sub
End If

'create and tune command object to retrieve lockers
' Initialize the SP
InitializeSPCommand cmdGetLockers, cnNew, GetLockersCommand

Set prmArg = cmdGetLockers.CreateParameter("Args", adVarChar, adParamInput, 1024, ArgLine)
cmdGetLockers.Parameters.Append prmArg

'procedure returns error number or 0 if OK
Set rsLockers = cmdGetLockers.Execute
With rsLockers
    'load lockers data
    Do Until .EOF
       LckCnt = LckCnt + 1
       LckID(LckCnt) = .Fields(0).value
       LckSeq(LckCnt) = .Fields(1).value
       LckName(LckCnt) = .Fields(2).value
       LckMW(LckCnt) = .Fields(3).value
       LckNET(LckCnt) = .Fields(4).value
       LckRET(LckCnt) = .Fields(5).value
       LckOET(LckCnt) = .Fields(6).value
       LckScore(LckCnt) = .Fields(7).value
       .MoveNext
    Loop
End With
rsLockers.Close
Set rsLockers = Nothing
lblStatus.Caption = "Number of loaded lockers: " & LckCnt

'clean things and exit
exit_cmdLoadLockers:
On Error Resume Next
Set cmdGetLockers.ActiveConnection = Nothing
Set cmdGetLockers = Nothing
cnNew.Close
Set cnNew = Nothing

If LckCnt > 0 Then
   If LckCnt - 1 < UBound(LckID) Then
      ReDim Preserve LckID(1 To LckCnt)
      ReDim Preserve LckSeq(1 To LckCnt)
      ReDim Preserve LckName(1 To LckCnt)
      ReDim Preserve LckMW(1 To LckCnt)
      ReDim Preserve LckScore(1 To LckCnt)
      ReDim LckET(1 To LckCnt, 2)
   End If
   For i = 0 To LckCnt - 1
      LckET(i, NET_COL) = LckNET(i)
      LckET(i, RET_COL) = LckRET(i)
      LckET(i, OET_COL) = LckOET(i)
   Next i
   'enable commands for listing lockers and search
   cmdListLockers.Enabled = True
   cmdSearch.Enabled = True
Else
   Erase LckID
   Erase LckSeq
   Erase LckName
   Erase LckMW
   Erase LckScore
End If
Screen.MousePointer = vbDefault
Exit Sub

err_cmdLoadLockers:
Select Case Err.Number
Case 9                       'need more room for lockers
    ReDim Preserve LckID(1 To LckCnt + 200)
    ReDim Preserve LckSeq(1 To LckCnt + 200)
    ReDim Preserve LckName(1 To LckCnt + 200)
    ReDim Preserve LckMW(1 To LckCnt + 200)
    ReDim Preserve LckNET(1 To LckCnt + 200)
    ReDim Preserve LckRET(1 To LckCnt + 200)
    ReDim Preserve LckOET(1 To LckCnt + 200)
    ReDim Preserve LckScore(1 To LckCnt + 200)
    Resume
Case 13, 94                  'Type Mismatch or Invalid Use of Null
    Resume Next
Case Else
    ErrCnt = ErrCnt + 1
    lblStatus.Caption = "Error retrieving lockers information from the database!"
    If ErrCnt < 10 Then
       LogErrors Err.Number, "frmMTLockMass.cmdLoadLockers"
       Resume Next
    End If
End Select
GoTo exit_cmdLoadLockers
LckCnt = -1
End Sub

Private Sub cmdLock_Click()
'-------------------------------------------------------
'initiates lock of all masses based on frequency shifts
'(or calibration equation in unsupported segmented lock)
'-------------------------------------------------------
Me.MousePointer = vbHourglass
Select Case GroupingType
Case GR_SOLO, GR_UMC
    LMWork.Locked = FSLockMass()
    lblStatus.Caption = ""
    cmdReport.Enabled = True
Case GR_SEGMENTS            'lock on segment equation; not supported for now
    MsgBox "Locking on segmented calibration equation is not supported for now", vbOKOnly, glFGTU
End Select
Me.MousePointer = vbDefault
End Sub

Private Sub cmdReport_Click()
'-----------------------------------------------------
'generates text file with mass lock report
'-----------------------------------------------------
Dim FileNum As Integer
Dim FileNam As String
Dim sLine As String
Dim i As Long
On Error Resume Next

With LMWork
  If .Locked Then
    If MinFN <= MaxFN Then
       Me.MousePointer = vbHourglass
       lblStatus.Caption = "Generating report..."
       FileNum = FreeFile
       FileNam = GetTempFolder() & RawDataTmpFile
       Open FileNam For Output As FileNum
       If Not GelAnalysis(CallerID) Is Nothing Then
          Print #FileNum, "File: " & GelAnalysis(CallerID).MD_file
       Else
          Print #FileNum, "File: " & GelBody(CallerID).Caption
       End If
       Print #FileNum, String(60, "-")
       'should print search definition here
       sLine = "Scan #" & vbTab & "Count" & vbTab & "Locker" _
           & vbTab & "  Lock MW  " & vbTab & "Freq.Shift"
       Print #FileNum, sLine
       For i = MinFN To MaxFN
         sLine = i & vbTab & FNCnt(i) & vbTab
         Select Case FNLM(i)
         Case FN_LM_NOCANDIDATES
           sLine = sLine & "No lock mass candidates found in this scan!"
         Case FN_LM_ERROR
           sLine = sLine & "Error locking masses in this scan!"
         Case Else
           sLine = sLine & FNLckID(i) & vbTab & Format(.MWLM(.MWInd(FNLM(i))), "0.00000") _
               & vbTab & Format$(FNS(i), "0.00000")
         End Select
         Print #FileNum, sLine
       Next i
       Close FileNum
       lblStatus.Caption = ""
       frmDataInfo.Caption = GelAnalysis(CallerID).Desc_DataFolder & "LockInfo.txt"
       frmDataInfo.Tag = "MTLM"
       frmDataInfo.Show vbModal
    Else
       MsgBox "No scans found!", vbOKOnly
    End If
    Me.MousePointer = vbDefault
  Else
    MsgBox "No mass locking applied yet!", vbOKCancel, "2DGelLand - Lock Mass"
  End If
End With
End Sub

Private Sub cmdSearch_Click()
'---------------------------------------------------------------
'initialize statistics arrays, clean gel from all
'lockers information and look for current lockers
'---------------------------------------------------------------
Dim TtlHitsCnt As Long
Dim Resp As Long
On Error Resume Next

Resp = MsgBox("This procedure will clean all identification data from the gel! Continue?", vbYesNo)
If Resp <> vbYes Then Exit Sub
Me.MousePointer = vbHourglass
lblStatus.Caption = "Cleaning present lockers information!"
DoEvents
CleanIDData CallerID
lblStatus.Caption = "Initializing statistic structure!"
DoEvents
InitLckStat
lblStatus.Caption = "Matching gel distributions with locker set!"
DoEvents
TtlHitsCnt = MatchLockers
If TtlHitsCnt >= 0 Then
    MsgBox "Total number of distributions matched with lockers: " & TtlHitsCnt, vbOKOnly
    If TtlHitsCnt > 5000 Then MsgBox "That's a lot, dude!", vbOKOnly
    cmdSearchResults.Enabled = True
    cmdSelectLockers.Enabled = True
Else
    MsgBox "Error occured while matching with lockers!", vbOKOnly
End If
lblStatus.Caption = ""
Me.MousePointer = vbDefault
End Sub

Private Sub cmdSearchResults_Click()
'------------------------------------------------------
'prints lockers matching results to temporary text file
'and displays it in the different form
'------------------------------------------------------
Dim FileNum As Integer
Dim FileNam As String
Dim sLine As String
Dim AvgErrDa As String
Dim AvgErrppm As String
Dim AvgErrET As String
Dim i As Long
On Error GoTo exit_cmdSearchResults

If LckCnt > 0 Then
   Me.MousePointer = vbHourglass
   FileNum = FreeFile()
   FileNam = GetTempFolder() & RawDataTmpFile
   Open FileNam For Output As FileNum
   'print gel file name and Search definition as reference
   Print #FileNum, "Generated by: " & GetMyNameVersion() & " on " & Now()
   Print #FileNum, "Gel File: " & GelAnalysis(CallerID).MD_file
   Print #FileNum, GetLockersMatchDefinition()
   sLine = "Lck ID" & glARG_SEP & "Lck MW" & glARG_SEP & "Lck ET" & glARG_SEP _
           & "Hits" & glARG_SEP & "Lck MW Avg Error(Da)" & glARG_SEP & _
           "Lck MW Avg Error(ppm)" & glARG_SEP & "Lck ET Avg Error" & _
           glARG_SEP & "Lck Range Min" & glARG_SEP & "Lck Range Max"
   Print #FileNum, sLine
   For i = 1 To LckCnt
     sLine = LckID(i) & glARG_SEP & LckMW(i) & glARG_SEP & LckET(i, ElutionType) & glARG_SEP & LckHits(i)
     If LckHits(i) > 0 Then
        AvgErrDa = Str(LckMWErr(i) / LckHits(i))
        AvgErrppm = Str(AvgErrDa / (LckMW(i) * glPPM))
        AvgErrET = Str(LckETErr(i) / LckHits(i))
        sLine = sLine & glARG_SEP & AvgErrDa & glARG_SEP & AvgErrppm & glARG_SEP _
              & AvgErrET & glARG_SEP & LckETMin(i) & glARG_SEP & LckETMax(i)
     End If
     Print #FileNum, sLine
   Next i
   Close FileNum
   DoEvents
   frmDataInfo.Tag = "LCK"
   DoEvents
   frmDataInfo.Show vbModal
Else
   MsgBox "No loaded lockers found!", vbOKOnly
End If

exit_cmdSearchResults:
MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbOKOnly
On Error Resume Next
Close FileNum
Me.MousePointer = vbDefault
End Sub

Private Sub cmdSelectLockers_Click()
'---------------------------------------------------
'initiates locker selection procedures based on user
'selected options
'---------------------------------------------------
On Error Resume Next
If Not FillArrays() Then Exit Sub
Select Case GroupingType
Case GR_SOLO
    Select Case FreqShiftCalcType
    Case FS_AVERAGE
    Case FS_BEST_WITH_REST
    Case FS_BEST_UMC_SCORE          'should not happen
    Case FS_BEST_LCK_SCORE
    End Select
Case GR_UMC
    lblStatus.Caption = "Calculating UMC parameters..."
    DoEvents
    Me.MousePointer = vbHourglass
    ClsCnt = UMCStatistics1(CallerID, ClsStat())
    If ClsCnt > 0 Then
       SelectLockers_UMC
       cmdLock.Enabled = True
    Else
       MsgBox "Error calculating UMC statistic! Make sure that gel was broken to Unique Mass Classes!", vbOKOnly
    End If
    Me.MousePointer = vbDefault
Case GR_SEGMENTS
    Select Case FreqShiftCalcType
    Case FS_AVERAGE
    Case FS_BEST_WITH_REST
    Case FS_BEST_UMC_SCORE
    Case FS_BEST_LCK_SCORE
    End Select
End Select
End Sub

Private Sub cmdSelectLockersSet_Click()
If GelAnalysis(CallerID) Is Nothing Then
   MsgBox "Current display is not associated with mass tags databases.  Cannot save the lockers.", vbOKOnly, glFGTU
Else
   GelAnalysis(CallerID).MTDB.SelectLockers
End If
End Sub

Private Sub Form_Activate()
'----------------------------------------------------------------
'load neccessary information and enable commands if everything OK
'----------------------------------------------------------------
On Error Resume Next
CallerID = Me.Tag
'retrieve name of stored procedure that will retrieve lockers information
GetLockersCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetLockers
If Len(GetLockersCommand) > 0 And bLoading Then
   With GelData(CallerID)
      Select Case UCase(.CalEquation)
      Case UCase(CAL_EQUATION_1)
          lblCalEquation.Caption = CAL_EQUATION_1
          Set Cal = New CalEq1
          Cal.A = .CalArg(1)
          Cal.B = .CalArg(2)
          If Not Cal.EquationOK Then
             lblStatus.Caption = "Invalid parameters in calibration equation!"
             Exit Sub
          End If
      Case UCase(CAL_EQUATION_4)
          lblCalEquation.Caption = CAL_EQUATION_4
          Set Cal = New CalEq4
          Cal.A = .CalArg(1)
          Cal.B = .CalArg(2)
          If Not Cal.EquationOK Then
             lblStatus.Caption = "Invalid parameters in calibration equation!"
             Exit Sub
          End If
      Case UCase(CAL_EQUATION_5)
          lblCalEquation.Caption = CAL_EQUATION_5
          Set Cal = New CalEq5
          Cal.A = .CalArg(1)
          Cal.B = .CalArg(2)
          If Not Cal.EquationOK Then
             lblStatus.Caption = "Invalid parameters in calibration equation!"
             Exit Sub
          End If
      Case UCase(CAL_EQUATION_2), UCase(CAL_EQUATION_3)
          If .CalArg(3) = 0 Then    'same case as CAL_EQUATION_1
             lblCalEquation.Caption = CAL_EQUATION_1
             Set Cal = New CalEq1
             Cal.A = .CalArg(1)
             Cal.B = .CalArg(2)
             If Not Cal.EquationOK Then
                lblStatus.Caption = "Invalid parameters in calibration equation!"
                Exit Sub
             End If
          Else
            lblStatus.Caption = "Calibration equation form not implemented!"
            Exit Sub
          End If
      Case Else     'later we might use other calibration equations
          lblStatus.Caption = "Calibration equation not found! Lock mass function not possible!"
          Exit Sub
      End Select
      'load all scans
      MinFN = .ScanInfo(1).ScanNumber
      MaxFN = .ScanInfo(UBound(.ScanInfo)).ScanNumber
      bLoading = False
   End With
   cmdLoadLockers.Enabled = True
Else        'can not proceede without this information
   lblStatus.Caption = "Lockers information not found!"
End If
With GelAnalysis(CallerID)
    'bring elution formula from database
    lblTICFit.Caption = Format$(.NET_TICFit, "0.0000")
    lblTICSlope.Caption = Format$(.NET_Slope, "0.0000000000")
    lblTICIntercept.Caption = Format$(.NET_Intercept, "0.0000000000")
    If .NET_Intercept <> 0 Or .NET_Slope <> 0 Then
       cmbElutionFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
    Else
       cmbElutionFormula.Text = samtDef.Formula
    End If
    ElutionFormula = cmbElutionFormula.Text
End With
End Sub

Private Sub Form_Load()
'--------------------------------------------
'set default values on controls and variables
'--------------------------------------------
bLoading = True
With cmbET
    .AddItem "Normalized"
    .AddItem "Retention"
    .AddItem "Order"
    .ListIndex = NET_COL
End With

With cmbBestClassScore
    .AddItem "NIntensity"
    .AddItem "NCalculated Fit"
    .AddItem "NLockers MMA Fit"
    .AddItem "NLockers Hit"
    .ListIndex = UMC_SCORE_ABUNDANCE
End With

With cmbElutionFormula
    .AddItem "(FN-MinFN)/(MaxFN-MinFN)"
    .AddItem "FN/MaxFN"
    .AddItem "FN*NET_Slope+NET_Intercept"
End With

GroupingType = GR_UMC
FreqShiftCalcType = FS_BEST_LCK_SCORE
ClassBestScoreType = UMC_SCORE_ABUNDANCE
ElutionType = NET_COL
MMA = CDbl(txtMMA.Text)
ETTol = CDbl(txtETTol.Text)
InitExprEvaluator ElutionFormula
End Sub

Private Function FillArrays() As Boolean
'------------------------------------------------
'returns True if data structures are successfully
'initialized and filled; False on any error
'------------------------------------------------
Dim CurrFN As Long, PrevFN As Long
Dim i As Long
Dim tmp As Long
Dim FNSorter As New QSInteger

LMWork.MWCnt = 0
lblStatus.Caption = "Loading data structures..."
DoEvents
With GelData(CallerID)
    If .CSLines + .IsoLines > 0 Then
       ReDim LMWork.MWID(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWFN(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWCS(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWLM(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWType(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWInd(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWFreqShift(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWLckID(1 To .CSLines + .IsoLines)
       ReDim LMWork.MWMassCorrection(1 To .CSLines + .IsoLines)
    Else
       lblStatus.Caption = "Distributions must be hiding, none found!"
       Exit Function
    End If

    For i = 1 To .CSLines
       LMWork.MWCnt = LMWork.MWCnt + 1
       tmp = LMWork.MWCnt
       LMWork.MWID(tmp) = i
       LMWork.MWType(tmp) = glCSType
       LMWork.MWFN(tmp) = .CSData(i).ScanNumber
       LMWork.MWCS(tmp) = .CSData(i).Charge
       LMWork.MWLM(tmp) = .CSData(i).AverageMW
       LMWork.MWInd(tmp) = tmp    'index for indirect access
    Next i
    For i = 1 To .IsoLines
       LMWork.MWCnt = LMWork.MWCnt + 1
       tmp = LMWork.MWCnt
       LMWork.MWID(tmp) = i
       LMWork.MWType(tmp) = glIsoType
       LMWork.MWFN(tmp) = .IsoData(i).ScanNumber
       LMWork.MWCS(tmp) = .IsoData(i).Charge
       LMWork.MWLM(tmp) = GetIsoMass(.IsoData(i), amtlmDef.lmIsoField)
       LMWork.MWInd(tmp) = tmp      'index for indirect access
    Next i
    If (LMWork.MWCnt > 0) Then        'initialize FN-arrays
        ReDim FNLM(MinFN To MaxFN)
        ReDim FNS(MinFN To MaxFN)
        ReDim FNLckID(MinFN To MaxFN)
        ReDim FNCnt(MinFN To MaxFN)
        ReDim FNIndMin(MinFN To MaxFN)
        ReDim FNIndMax(MinFN To MaxFN)
        ReDim FNCls(MinFN To MaxFN)
        ReDim FNClsLM(MinFN To MaxFN)
    Else
        lblStatus.Caption = ""
        Exit Function
    End If
End With
'order MW on FN and count original data
lblStatus.Caption = "Analyzing original data..."
DoEvents
With LMWork
  If Not FNSorter.QSAsc(.MWFN(), .MWInd()) Then
     lblStatus.Caption = "Error sorting data! Errare Humanum Est!"
     Exit Function
  Else
     PrevFN = MinFN
     CurrFN = -1
     For i = 1 To .MWCnt
       If .MWFN(i) <> CurrFN Then
          'note last reference index of previous scan (except for first element)
          If PrevFN = CurrFN Then FNIndMax(PrevFN) = i - 1
          CurrFN = .MWFN(i)
          PrevFN = CurrFN
          FNIndMin(CurrFN) = i      'note first reference index of this scan
       End If
       FNCnt(CurrFN) = FNCnt(CurrFN) + 1
     Next i
     FNIndMax(MaxFN) = .MWCnt          'note last reference index of the last scan
  End If
End With
lblStatus.Caption = ""
DoEvents
FillArrays = True
End Function

Private Function FSLockMass() As Boolean
'----------------------------------------------------------
'locks masses in scans from calculated scan frequency shift
'returns true if no error occured; false otherwise
'----------------------------------------------------------
Dim i As Integer
Dim j As Long
Dim MOverZ As Double
Dim Freq As Double
Dim CS As Double
Dim LM As Double        'locked mass
'On Error GoTo err_FSLockMass

With LMWork
  Me.MousePointer = vbHourglass
  For i = MinFN To MaxFN
    lblStatus.Caption = "Locking masses in scan " & i & " [" & MinFN & "," & MaxFN & "]"
    DoEvents
    If FNIndMin(i) <= FNIndMax(i) And FNIndMax(i) > 0 Then
       If FNS(i) <> 0 Then     'there is a frequency shift
          For j = FNIndMin(i) To FNIndMax(i)
            CS = .MWCS(.MWInd(j))
            If CS > 0 Then
               MOverZ = .MWLM(.MWInd(j)) / CS + glMASS_CC
               Freq = Cal.CyclotFreq(MOverZ) + FNS(i)
               MOverZ = Cal.MOverZ(Freq)
               LM = CS * (MOverZ - glMASS_CC)
               .MWMassCorrection(.MWInd(j)) = LM - .MWLM(.MWInd(j))
               .MWLM(.MWInd(j)) = LM
               .MWLckID(.MWInd(j)) = FNLckID(i)
               .MWFreqShift(.MWInd(j)) = FNS(i)
            End If
          Next j
       End If
    Else
       FNLM(i) = FN_LM_NOCANDIDATES
    End If
  Next i
End With
FSLockMass = True
Exit Function

err_FSLockMass:
MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbOKOnly
End Function

'''Public Function GetAMTMWByID(ByVal ID As String) As Double
''''---------------------------------------------------------
''''returns AMT MW for AMT; not optimized to tell the truth
''''---------------------------------------------------------
'''Dim i As Long
'''GetAMTMWByID = -1
'''For i = 1 To UBound(AMTData)
'''    If AMTData(i).ID = ID Then
'''       GetAMTMWByID = AMTData(i).MW
'''       Exit Function
'''    End If
'''Next i
'''End Function

Private Sub UMCBestLck(ByVal ClsInd As Long, _
                       ByRef ID As Long, _
                       ByRef Hits As Long)
'-----------------------------------------------------------
'returns in ID best LckID for class with index
'ClsInd (in ClsStat array); empty string if none
'in Hits procedure returns number of class member that hit
'best LckID; based on this procedure each class is assigned
'one(or none)lockers ID
'-----------------------------------------------------------
Dim Lcks() As String
Dim LckCnt
Dim Cnt As Long
Dim UnqLckCnt As Long
Dim IDString As String
'-----------------------------------------------------------
'we need to keep list of unique lockers hit in each class
'so that we can select the one with most sense. Selection
'goes like this; lockers hit most time in class is selected
'as best; if there are 2 or more with same number of hits
'the one with best elution time matching is selected; if ET
'measurement does not exists (or if it does not provide
'answer, the one with best mass agreement with the class is
'selected(optimistic instrument evaluation)
'-----------------------------------------------------------
Dim IsInList As Boolean
Dim LckList() As String                'unique lockers list
Dim LckMWErr() As Double              'mass error
Dim LckETErr() As Double              'elution time error
Dim LckHits() As Long                 'number of hits
Dim sLck As String
Dim sMWErr As String
Dim sETErr As String
'next 2 variables are not actually best errors; rather they are
'errors of lockers currently considered to be the best choice
Dim BestMWErr As Double, BestETErr As Double
Dim MostHits As Long
Dim BestLck As String
Dim i As Long, j As Long, k As Long
On Error GoTo exit_UMCBestLck
ID = -1
Hits = 0
UnqLckCnt = 0
With GelUMC(CallerID).UMCs(ClsInd)
    Cnt = 0
    For i = 0 To .ClassCount - 1
        'for each class member extract all Lck listed
        IDString = ""
        Select Case .ClassMType(i)
        Case glCSType
               IDString = GelData(CallerID).CSData(.ClassMInd(i)).MTID
        Case glIsoType
               IDString = GelData(CallerID).IsoData(.ClassMInd(i)).MTID
        End Select
        LckCnt = GetTagRefFromString(LCK_MARK, IDString, Lcks())
        'NOTE: Lcks array is 0-based
        If LckCnt > 0 Then
           UnqLckCnt = UnqLckCnt + LckCnt
           ReDim Preserve LckList(1 To UnqLckCnt)
           ReDim Preserve LckMWErr(1 To UnqLckCnt)
           ReDim Preserve LckETErr(1 To UnqLckCnt)
           ReDim Preserve LckHits(1 To UnqLckCnt)
           For j = 1 To LckCnt
             IsInList = False
             If Cnt > 0 Then   'First is always new
               sLck = GetIDFromString(Lcks(j - 1), LCK_MARK, LckIDEnd)
               sMWErr = GetMWErrFromString(Lcks(j - 1))
               sETErr = GetETErrFromString(Lcks(j - 1))
               If IsNumeric(sMWErr) Then
                  For k = 1 To Cnt
                      If sLck = LckList(k) Then
                         LckHits(k) = LckHits(k) + 1
                         LckMWErr(k) = LckMWErr(k) + Abs(CDbl(sMWErr))
                         If IsNumeric(sETErr) Then LckETErr(k) = LckETErr(k) + CDbl(sETErr)
                         IsInList = True
                      End If
                   Next k
               Else
                  IsInList = True   'it is not but it is error so just ignore
               End If
             End If
             If Not IsInList Then
               Cnt = Cnt + 1
               sLck = GetIDFromString(Lcks(j - 1), LCK_MARK, LckIDEnd)
               sMWErr = GetMWErrFromString(Lcks(j - 1))
               sETErr = GetETErrFromString(Lcks(j - 1))
               If IsNumeric(sMWErr) Then
                  LckList(Cnt) = sLck
                  LckMWErr(Cnt) = Abs(CDbl(sMWErr))
                  If IsNumeric(sETErr) Then LckETErr(Cnt) = CDbl(sETErr)
                  LckHits(Cnt) = 1
               End If
             End If
           Next j
           UnqLckCnt = Cnt
        End If
    Next i
End With
'now go and pick the best as described above
If UnqLckCnt > 0 Then
   ReDim Preserve LckList(1 To UnqLckCnt)
   ReDim Preserve LckMWErr(1 To UnqLckCnt)
   ReDim Preserve LckHits(1 To UnqLckCnt)
   MostHits = 0
   BestMWErr = glHugeOverExp
   BestETErr = glHugeOverExp
   BestLck = -1
   For i = 1 To UnqLckCnt
       If (LckHits(i) >= MostHits) And (LckHits(i) > 0) Then
          If LckHits(i) > MostHits Then      'new most hits
             LckMWErr(i) = LckMWErr(i) / LckHits(i)
             LckETErr(i) = LckETErr(i) / LckHits(i)
             BestMWErr = LckMWErr(i)
             BestETErr = LckETErr(i)
             MostHits = LckHits(i)
             BestLck = LckList(i)
          Else    'same number of hits; check the errors(first elution)
             LckETErr(i) = LckETErr(i) / LckHits(i)
             LckMWErr(i) = LckMWErr(i) / LckHits(i)
             If LckETErr(i) < BestETErr Then     'select as best
                BestMWErr = LckMWErr(i)
                BestETErr = LckETErr(i)
                MostHits = LckHits(i)
                BestLck = LckList(i)
             Else    'then if elution same(this includes missing elution) mass
                If (LckETErr(i) = BestETErr) And (LckMWErr(i) < BestMWErr) Then
                  BestMWErr = LckMWErr(i)
                  BestETErr = LckETErr(i)
                  MostHits = LckHits(i)
                  BestLck = LckList(i)
                End If
             End If
          End If
       End If
   Next i
   ID = CLng(BestLck)
   Hits = MostHits
End If
exit_UMCBestLck:
End Sub

Private Function UMCScore(ByVal ClsInd As Long) As Double
'-----------------------------------------------------------------
'returns score best for class with index ClsInd (in ClsStat array)
'all scores are ascending meaning more is better
'-----------------------------------------------------------------
'On Error GoTo exit_UMCScore
UMCScore = -1
Select Case FreqShiftCalcType
Case FS_BEST_LCK_SCORE                           'score of the class best locker
    UMCScore = GetLckScoreByID(ClsLck(ClsInd))
Case FS_AVERAGE, FS_BEST_WITH_REST               'irelevant
    UMCScore = 1
Case FS_BEST_UMC_SCORE
    Select Case ClassBestScoreType
    Case UMC_SCORE_ABUNDANCE                     'abundance
         UMCScore = ClsStat(ClsInd, 6)           'average class abundance
    Case 1  'fit
         If ClsStat(ClsInd, 7) > 0 Then
            UMCScore = 1 / ClsStat(ClsInd, 7)    'inverse of average class fit
         Else
            UMCScore = glHugeOverExp
         End If
    Case 2  'Lockers fit
    Case 3  'number of AMT hits
         UMCScore = ClsHits(ClsInd)
    End Select
End Select
exit_UMCScore:
End Function

Private Function ScanLockerMWInd(ByVal FN As Integer, _
                                 ByVal LockerInd As Long) As Long
'----------------------------------------------------------------
'returns index of best scan locker for scan FN in the MWInd array
'LockerInd is index in GelData().IsoData array
'----------------------------------------------------------------
Dim i As Long
On Error Resume Next
With LMWork
  For i = FNIndMin(FN) To FNIndMax(FN)
    If .MWType(.MWInd(i)) = glIsoType Then
       If .MWID(.MWInd(i)) = LockerInd Then
          ScanLockerMWInd = i
          Exit Function
       End If
    End If
  Next i
End With
ScanLockerMWInd = FN_LM_ERROR
End Function

' Unused Procedure (February 2005)
''Private Sub EnableCommands(ByVal Arg As Boolean)
''cmdSelectLockers.Enabled = Arg
''cmdLock.Enabled = Arg
''cmdReport.Enabled = Arg
''End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Respond As Long
Dim i As Long
Dim j As Integer
Dim FNIndex As Integer
On Error GoTo err_Form_Unload

If Not LMWork.Locked Then Exit Sub

'copy results to better data structure (parallel with GelData)
lblStatus.Caption = "Saving results to alternate structure..."
DoEvents
With GelLM(CallerID)
    .CSCnt = GelData(CallerID).CSLines
    If .CSCnt > 0 Then
        ReDim .CSLckID(1 To .CSCnt)
        ReDim .CSFreqShift(1 To .CSCnt)
        ReDim .CSMassCorrection(1 To .CSCnt)
    End If
    .IsoCnt = GelData(CallerID).IsoLines
    If .IsoCnt > 0 Then
        ReDim .IsoLckID(1 To .IsoCnt)
        ReDim .IsoFreqShift(1 To .IsoCnt)
        ReDim .IsoMassCorrection(1 To .IsoCnt)
    End If
    For i = 1 To LMWork.MWCnt
        Select Case LMWork.MWType(i)
        Case glCSType
            .CSLckID(LMWork.MWID(i)) = LMWork.MWLckID(i)
            .CSFreqShift(LMWork.MWID(i)) = LMWork.MWFreqShift(i)
            .CSMassCorrection(LMWork.MWID(i)) = LMWork.MWMassCorrection(i)
        Case glIsoType
            .IsoLckID(LMWork.MWID(i)) = LMWork.MWLckID(i)
            .IsoFreqShift(LMWork.MWID(i)) = LMWork.MWFreqShift(i)
            .IsoMassCorrection(LMWork.MWID(i)) = LMWork.MWMassCorrection(i)
        End Select
    Next i
End With
Respond = MsgBox("Replace data in original gel?", vbYesNoCancel)
Select Case Respond
Case vbYes          'save
    Me.MousePointer = vbHourglass
    With GelData(CallerID)
         For i = 1 To LMWork.MWCnt
           Select Case LMWork.MWType(i)
           Case glCSType
             .CSData(LMWork.MWID(i)).AverageMW = LMWork.MWLM(i)
           Case glIsoType
             SetIsoMass .IsoData(LMWork.MWID(i)), amtlmDef.lmIsoField, LMWork.MWLM(i)
           End Select
         Next i
         'write frequency shifts
         For j = MinFN To MaxFN
           FNIndex = GetDFIndex(CallerID, j)
           If FNIndex > 0 Then
              .ScanInfo(FNIndex).FrequencyShift = FNS(j)
           End If
         Next j
         .Comment = .Comment & vbCrLf & Now() & vbCrLf & "Lock mass function applied!"
       End With
       GelStatus(CallerID).Dirty = True
    Me.MousePointer = vbDefault
Case vbNo           'do nothing
Case vbCancel       'do not unload
    Cancel = True
End Select
Exit Sub

err_Form_Unload:
If Err.Number <> 9 Then
   MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, glFGTU
End If
End Sub

Private Sub optFSCalculation_Click(Index As Integer)
FreqShiftCalcType = Index
End Sub

Private Sub optGrouping_Click(Index As Integer)
GroupingType = Index
Select Case GroupingType
Case GR_SOLO
    If optFSCalculation(FS_BEST_UMC_SCORE).value Then
       optFSCalculation(FS_BEST_LCK_SCORE).value = True
    End If
    optFSCalculation(FS_BEST_UMC_SCORE).Enabled = False
Case GR_UMC
    optFSCalculation(FS_BEST_UMC_SCORE).Enabled = True
Case GR_SEGMENTS
End Select
End Sub


Private Sub txtETTol_LostFocus()
'-----------------------------------------
'empty means do not use it so set it to -1
'-----------------------------------------
Dim sETTol As String
sETTol = Trim$(txtETTol.Text)
If Len(sETTol) > 0 Then
   If IsNumeric(sETTol) Then
      ETTol = CDbl(sETTol)
   Else
      MsgBox "Elution time tolerance should be positive number!", vbOKOnly
      txtETTol.SetFocus
   End If
Else
   ETTol = -1
End If
End Sub


Private Function MatchLockers() As Long
'------------------------------------------------------
'matches lockers and loaded distributions based on
'molecular mass and normalized elution time
'To optimize for speed arrays loaded from the AMT table
'are actually searched rather than database tables
'------------------------------------------------------
Dim MinScan As Long
Dim MaxScan As Long
Dim ScanWidth As Long
Dim LckRef As String
Dim i As Long

If Not InitExprEvaluator(ElutionFormula) Then
   MatchLockers = -2
   Exit Function
End If
With GelData(CallerID)
   HitsCount = 0
   Set mwutSearch = New MWUtil
   If Not mwutSearch.Fill(LckMW()) Then GoTo err_MatchLockers
   GetScanRange CallerID, MinScan, MaxScan, ScanWidth
   If ScanWidth <= 0 And ETTol = 0 Then GoTo err_MatchLockers 'can not do it
   For i = 1 To .CSLines
     lblStatus.Visible = Not lblStatus.Visible
     DoEvents
     If ETTol >= 0 Then
        LckRef = GetLckReferenceMWET(.CSData(i).AverageMW, ElTi(.CSData(i).ScanNumber, MinScan, MaxScan))
     Else
        LckRef = GetLckReferenceMW(.CSData(i).AverageMW, ElTi(.CSData(i).ScanNumber, MinScan, MaxScan))
     End If
     InsertBefore .CSData(i).MTID, LckRef
   Next i
   For i = 1 To .IsoLines
     lblStatus.Visible = Not lblStatus.Visible
     DoEvents
     If ETTol >= 0 Then
        LckRef = GetLckReferenceMWET(.IsoData(i).MonoisotopicMW, ElTi(.IsoData(i).ScanNumber, MinScan, MaxScan))
     Else
        LckRef = GetLckReferenceMW(.IsoData(i).MonoisotopicMW, ElTi(.IsoData(i).ScanNumber, MinScan, MaxScan))
     End If
     InsertBefore .IsoData(i).MTID, LckRef
   Next i
End With
MatchLockers = HitsCount
   
exit_MatchLockers:
lblStatus.Visible = True
DoEvents
Set mwutSearch = Nothing
Exit Function

err_MatchLockers:
MatchLockers = -1
GoTo exit_MatchLockers
End Function

Private Function GetLckReferenceMWET(ByVal MW As Double, _
                                     ByVal ET As Double, _
                                     Optional ByVal blnStoreAbsoluteValueOfError As Boolean = False) As String
'-----------------------------------------------------------------
'returns lockers reference string based on MW, ET
'-----------------------------------------------------------------
Dim LckRef As String
Dim MWTolRef As Double
Dim sMWTolRef As String
Dim ETTolRef As Double
Dim sETTolRef As String
Dim FirstInd As Long
Dim LastInd As Long
Dim AbsTol As Double
Dim i As Long
On Error GoTo exit_GetLckReferenceMWET

AbsTol = MW * MMA * glPPM
If mwutSearch.FindIndexRange(MW, AbsTol, FirstInd, LastInd) Then
   For i = FirstInd To LastInd
     If Abs(ET - LckET(i, ElutionType)) <= ETTol Then
        HitsCount = HitsCount + 1
        If blnStoreAbsoluteValueOfError Then
            MWTolRef = Abs(MW - LckMW(i))
        Else
            MWTolRef = MW - LckMW(i)
        End If
        sMWTolRef = MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd
        ETTolRef = (LckET(i, ElutionType) - ET)
        sETTolRef = ETErrMark & Format$(ETTolRef, "0.000") & ETErrEnd
        'put Lck ID and actual errors
        LckRef = LckRef & LCK_MARK & LckID(i) & sMWTolRef & sETTolRef
        LckRef = LckRef & glARG_SEP & Chr$(32)
        'do statistics
        LckHits(i) = LckHits(i) + 1
        LckMWErr(i) = LckMWErr(i) + MWTolRef
        LckETErr(i) = LckETErr(i) + ETTolRef
        If ET < LckETMin(i) Then LckETMin(i) = ET
        If ET > LckETMax(i) Then LckETMax(i) = ET
     End If
   Next i
End If
exit_GetLckReferenceMWET:
GetLckReferenceMWET = LckRef
End Function


Private Function GetLckReferenceMW(ByVal MW As Double, _
                                   ByVal ET As Double, _
                                   Optional ByVal blnStoreAbsoluteValueOfError As Boolean = False) As String
'---------------------------------------------------------------
'returns lockers reference string based on MW; ET is here used
'only to generate statistic and not as a search criteria
'---------------------------------------------------------------
Dim LckRef As String
Dim MWTolRef As Double
Dim sMWTolRef As String
Dim FirstInd As Long
Dim LastInd As Long
Dim AbsTol As Double
Dim i As Long
On Error GoTo exit_GetLckReferenceMW

AbsTol = MW * MMA * glPPM
If mwutSearch.FindIndexRange(MW, AbsTol, FirstInd, LastInd) Then
   For i = FirstInd To LastInd
       HitsCount = HitsCount + 1
       If blnStoreAbsoluteValueOfError Then
           MWTolRef = Abs(MW - LckMW(i))
       Else
           MWTolRef = MW - LckMW(i)
       End If
       sMWTolRef = MWErrMark & Format$(MWTolRef / (MW * glPPM), "0.00") & MWErrEnd
       'put locker ID and actual errors
       LckRef = LckRef & LCK_MARK & LckID(i) & sMWTolRef & glARG_SEP & Chr$(32)
       'do statistics
       LckHits(i) = LckHits(i) + 1
       LckMWErr(i) = LckMWErr(i) + MWTolRef
       LckETErr(i) = LckETErr(i) + (LckET(i, ElutionType) - ET)
       If ET < LckETMin(i) Then LckETMin(i) = ET
       If ET > LckETMax(i) Then LckETMax(i) = ET
   Next i
End If
exit_GetLckReferenceMW:
GetLckReferenceMW = LckRef
End Function

Public Sub InitLckStat()
'-------------------------------------------------
'redimensions and initialize statistic arrays
'-------------------------------------------------
Dim i As Long
If LckCnt > 0 Then
   ReDim LckHits(1 To LckCnt)
   ReDim LckMWErr(1 To LckCnt)
   ReDim LckETErr(1 To LckCnt)
   ReDim LckETMin(1 To LckCnt)
   ReDim LckETMax(1 To LckCnt)
   'only last 2 arrays need special initialization
   For i = 1 To LckCnt
       LckETMin(i) = glHugeOverExp
       LckETMax(i) = -1
   Next i
End If
End Sub

Private Sub txtMMA_LostFocus()
On Error Resume Next
MMA = CDbl(Trim$(txtMMA.Text))
'If Err Then
'   MsgBox "MMA should be positive number!", vbOKOnly
'   txtMMA.SetFocus
'End If
End Sub


Private Function InitExprEvaluator(ByVal sExpr As String) As Boolean
'-------------------------------------------------------------------
'initialize expression evaluator to be used with elution formula
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


Private Function ElTi(FN As Long, MinScan As Long, MaxScan As Long)
'-----------------------------------------
'returns expression value (elution time)
'-----------------------------------------
VarVals(1) = FN
VarVals(2) = MinScan
VarVals(3) = MaxScan
ElTi = MyExprEva.ExprVal(VarVals())
End Function


Private Function GetLockersMatchDefinition()
'----------------------------------------------
'returns locking match definition
'----------------------------------------------
Dim tmp As String
tmp = "Org. DB Name(Lockers)=" & GelAnalysis(CallerID).Organism_DB_Name & vbCrLf
tmp = tmp & "MMA(Lockers)=" & MMA & vbCrLf
tmp = tmp & "Elution type(Lockers)=" & ElutionType & vbCrLf
tmp = tmp & "Elution tolerance(Lockers)=" & ETTol & vbCrLf
tmp = tmp & "Elution formula(Lockers)=" & ElutionFormula & vbCrLf
GetLockersMatchDefinition = tmp
End Function

Private Sub SelectLockers_UMC()
'-----------------------------------------------------
'here comes actual workload for selecting lockers when
'unique mass class strategy is used
'NOTE: This procedure can not be called if ClsCnt<=0
'-----------------------------------------------------
Dim i As Long, j As Long
Dim MOverZ As Double
Dim Freq As Double
Dim FreqLM As Double
Dim LM As Double
Dim CS As Double
Dim BestScore As Double
Dim BestScoreInd As Long
Dim ScanLocker As Long
Dim BestScanLocker As Long

On Error Resume Next
'pick best locker choice, and score for each class
lblStatus.Caption = "Selecting AMTs representing UMCs..."
DoEvents
ReDim ClsLck(ClsCnt - 1)
ReDim ClsHits(ClsCnt - 1)
ReDim ClsScore(ClsCnt - 1)
For i = 0 To ClsCnt - 1
    lblStatus.Caption = "Class " & (i + 1) & "/" & ClsCnt
    DoEvents
    UMCBestLck ClsStat(i, 0), ClsLck(i), ClsHits(i)
    ClsScore(i) = UMCScore(i)
Next i
Select Case FreqShiftCalcType
Case FS_AVERAGE
Case FS_BEST_WITH_REST
Case FS_BEST_UMC_SCORE, FS_BEST_LCK_SCORE
    'pick the best class for each scan
    lblStatus.Caption = "Selecting the best UMC for each scan..."
    DoEvents
    For i = MinFN To MaxFN
        lblStatus.Visible = Not lblStatus.Visible
        DoEvents
        BestScore = -glHugeOverExp
        BestScoreInd = -1
        BestScanLocker = -1
        For j = 0 To ClsCnt - 1
            If (ClsStat(j, 2) <= i) And (i <= ClsStat(j, 3)) And (ClsLck(j) > 0) Then
              'class covers this scan and it is also lockers hit; because
              'Charge State data can not be "mass lockers" we need to check
              'that there is a Isotopic point in this class in this scan
              If ClsScore(j) > BestScore Then
                 ScanLocker = UMCScanLocker(CallerID, ClsStat(j, 0), i)
                 If ScanLocker > 0 Then
                    BestScoreInd = j
                    BestScore = ClsScore(j)
                    BestScanLocker = ScanLocker
                 End If
              End If
            End If
        Next j
        FNCls(i) = BestScoreInd
        FNClsLM(i) = BestScanLocker
    Next i
    lblStatus.Visible = True
    DoEvents
    'pick the best "lock mass" for each scan(from the best class for that scan)
    lblStatus.Caption = "Selecting lock masses..."
    DoEvents
    'calculate also frequency shift here
    With LMWork
       For i = MinFN To MaxFN
         If FNCls(i) > 0 Then              'find index to which locker matches
            FNLM(i) = ScanLockerMWInd(i, FNClsLM(i))
            FNLckID(i) = ClsLck(FNCls(i))
         Else
            FNLM(i) = FN_LM_NOCANDIDATES
         End If
         If FNLM(i) > 0 Then
            Select Case .MWType(.MWInd(FNLM(i)))
            Case glCSType          'Charge State data can not be lock mass???
            Case glIsoType
               CS = .MWCS(.MWInd(FNLM(i)))
               If CS > 0 Then      'calculate frequency for experimental mw
                  MOverZ = .MWLM(.MWInd(FNLM(i))) / CS + glMASS_CC
                  Freq = Cal.CyclotFreq(MOverZ)
                  'calculate frequency for theoretical mw
                  LM = GetLckMWByID(FNLckID(i))
                  If LM > 0 Then
                     MOverZ = LM / CS + glMASS_CC
                     FreqLM = Cal.CyclotFreq(MOverZ)
                     'we have now frequency shift
                     FNS(i) = FreqLM - Freq
                  End If
               End If
            End Select
         End If
       Next i
    End With
    'put lock masses to the gel selection(Isotopic)
    With GelBody(CallerID).GelSel
        For i = MinFN To MaxFN
            If FNLM(i) > 0 Then
                .AddToIsoSelection LMWork.MWID(LMWork.MWInd(FNLM(i)))
            End If
        Next i
    End With
    lblStatus.Caption = ""
    cmdLock.Enabled = True
End Select
End Sub

Private Function GetLckScoreByID(ByVal ID As Long) As Double
'-----------------------------------------------------------
'returns lockers score; locker set is small array so no need
'to complicate here
'-----------------------------------------------------------
Dim i As Long
GetLckScoreByID = -1             'matches Unknown in the database
If ID <= 0 Then Exit Function    'no lockers with ID<0
For i = 1 To LckCnt
    If LckID(i) = ID Then
       GetLckScoreByID = LckScore(i)
       Exit Function
    End If
Next i
End Function

Private Function GetLckMWByID(ByVal ID As Long) As Double
'-----------------------------------------------------------
'returns lockers score; locker set is small array so no need
'to complicate here
'-----------------------------------------------------------
Dim i As Long
If ID <= 0 Then Exit Function 'no lockers with ID<0
For i = 1 To LckCnt
    If LckID(i) = ID Then
       GetLckMWByID = LckMW(i)
       Exit Function
    End If
Next i
End Function


VERSION 5.00
Begin VB.Form frmNewDummyAnalysis 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select MT Database For Analysis"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmNewDummyAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelectMassTags 
      Caption         =   "Select Mass Tags"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      ToolTipText     =   "Select mass tags to search with this analysis"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   4800
      Width           =   975
   End
   Begin VB.Frame fraStage 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Organism Mass Tag Database"
      Height          =   4575
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtSearchForDB 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   3330
         Width           =   2655
      End
      Begin VB.CommandButton cmdSearchForDB 
         Caption         =   "&Search"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   3300
         Width           =   1095
      End
      Begin VB.CheckBox chkShowUnusedDBs 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Show Unused Databases"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CheckBox chkShowFrozenDBs 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Show Frozen Databases"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.ListBox lstOrgMTDBNames 
         Height          =   2205
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search for:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lblMTDBDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "No directory data found; Server might be down!"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   5895
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Organism Mass Tag Database"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmNewDummyAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
'form that controls initialization of dummy FTICR analysis
'the only purpose of this analysis is to provide connection
'with the Mass Tag database so that out-of-system data
'files could be linked and compared with Mass Tag databases
'----------------------------------------------------------
'last modified: 12/11/2001 nt
'modified from frmNewAnalysis
'----------------------------------------------------------
Option Explicit

Const SECTION_MTDB = "[Organism Mass Tag Database]"

Dim MyInit As New InitFile         'object that deals with initialization file

Public InitFileName As String                  'full path to init file
Public fAnalysis As FTICRAnalysis              'FTICR Analysis object
Attribute fAnalysis.VB_VarHelpID = -1

Public Event AnalysisDialogClose()

Private MTDBInd As Long                     'index of currently selected db
Private MTDBCnt As Long                     'count of mass tag databases
Private MTDBInfo() As udtMTDBInfoType

Private MTDBCntVisible As Long
Private MTDBNameListPointers() As Long          ' Used to display database names sorted properly

Private Sub EditAddName(ByRef objCol As Collection, ByVal PairName As String, ByVal NewValue As String)
'-------------------------------------------------------------------------
'modifies value of name value pair; if pair does not exist adds it
'-------------------------------------------------------------------------
Dim nv As NameValue
On Error Resume Next
objCol.Item(PairName).Value = NewValue
If Err Then
   Set nv = New NameValue
   nv.Name = PairName
   nv.Value = NewValue
   objCol.Add nv, nv.Name
End If
End Sub

Private Sub HighlightDBByName(ByVal strTextToFind As String, ByVal intIndexStart)
    
    Dim i As Integer
    Dim intCharLoc As Integer
    
    If Len(strTextToFind) > 0 And lstOrgMTDBNames.ListCount > 0 Then
        ' Step through lstOrgMTDBNames and find the first to contain strTextToFind (starting at index intIndexStart)
            
        strTextToFind = LCase(strTextToFind)
        
        If intIndexStart < 0 Then
            intIndexStart = lstOrgMTDBNames.ListCount - 1
        End If

        i = intIndexStart
        Do
            i = i + 1
            If i > lstOrgMTDBNames.ListCount - 1 Then
                i = 0
            End If

            intCharLoc = InStr(LCase(lstOrgMTDBNames.List(i)), strTextToFind)
            
            If intCharLoc > 0 Then
                lstOrgMTDBNames.ListIndex = i
                Exit Do
            End If

        Loop While i <> intIndexStart
        
    End If
    
End Sub

Private Sub PopulateDatabaseCombobox()
    Dim i As Long
    Dim blnShowFrozenDBs As Boolean
    Dim blnShowUnusedDBs As Boolean     ' Forced to False if blnShowFrozenDBs = False
    Dim strDatabaseNameSaved As String
    
On Error GoTo PopulateDatabaseComboboxErrorHandler

    If chkShowFrozenDBs.Value = vbChecked Then
        blnShowFrozenDBs = True
        blnShowUnusedDBs = (chkShowUnusedDBs.Value = vbChecked)
    Else
        blnShowFrozenDBs = False
        blnShowUnusedDBs = False
    End If
    
    If lstOrgMTDBNames.ListIndex >= 0 Then
        strDatabaseNameSaved = lstOrgMTDBNames.List(lstOrgMTDBNames.ListIndex)
    End If
    
    lstOrgMTDBNames.Clear
    
    MTDBCnt = UBound(MTDBInfo) + 1
    If MTDBCnt > 0 Then
        SortMTDBNameList MTDBInfo(), MTDBNameListPointers(), 0, MTDBCnt - 1, blnShowFrozenDBs, blnShowUnusedDBs
        
        MTDBCntVisible = UBound(MTDBNameListPointers) + 1
        
        For i = 0 To MTDBCntVisible - 1
            lstOrgMTDBNames.AddItem MTDBInfo(MTDBNameListPointers(i)).Name
            If MTDBInfo(MTDBNameListPointers(i)).Name = strDatabaseNameSaved Then
                lstOrgMTDBNames.ListIndex = lstOrgMTDBNames.ListCount - 1
            End If
        Next i
        
        If lstOrgMTDBNames.ListIndex < 0 And lstOrgMTDBNames.ListCount > 0 Then
            lstOrgMTDBNames.ListIndex = 0
        End If
    
    End If
    
    Exit Sub

PopulateDatabaseComboboxErrorHandler:
Debug.Assert False

End Sub

Private Sub chkShowFrozenDBs_Click()
    EnableDisableControls
    PopulateDatabaseCombobox
End Sub

Private Sub chkShowUnusedDBs_Click()
    PopulateDatabaseCombobox
End Sub

Private Sub cmdCancel_Click()
Set fAnalysis = Nothing
Me.Hide
RaiseEvent AnalysisDialogClose
End Sub

Private Sub cmdOK_Click()
Me.Hide
RaiseEvent AnalysisDialogClose
End Sub

Private Sub cmdSearchForDB_Click()
    HighlightDBByName txtSearchForDB, lstOrgMTDBNames.ListIndex
End Sub

Private Sub cmdSelectMassTags_Click()
fAnalysis.MTDB.SelectMassTags InitFileName
End Sub

Private Sub Form_Load()
'-----------------------------------------------------------
'loads information from initialization file
'-----------------------------------------------------------
Dim Res As Long
On Error Resume Next

EnableDisableControls

'allow form to present itself
Me.Visible = True
Me.Show
DoEvents

Set fAnalysis = New FTICRAnalysis
fAnalysis.ProcessingType = fptDummy
MTDBInd = -1
' 12/12/2004 mem - Switched from using MT_Main to MTS_Master to retrieve DB info
Res = GetMTSMasterDirectoryData(InitFileName, MTDBInfo)
If Res <> 0 Then
   Set fAnalysis = Nothing
   MsgBox "Error loading Mass Tag databases information!" & vbCrLf & "Error: " & Res & " - " & Error(Res), vbOKOnly
   Me.Hide
   RaiseEvent AnalysisDialogClose
Else
   GetMTDBSchema
   PopulateDatabaseCombobox
End If
End Sub

Private Sub EnableDisableControls()
    chkShowUnusedDBs.Enabled = (chkShowFrozenDBs.Value = vbChecked)
End Sub

Private Sub GetMTDBSchema()
'-----------------------------------------------------
'retrieves schema information from initialization file
'this is same for all Mass Tags databases
'-----------------------------------------------------
Dim Arg() As String
Dim ArgCnt As Long
Dim MTName As String, MTValue As String
Dim MTValuePos As Long
Dim nv As NameValue
Dim i As Long
On Error Resume Next

ArgCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_MTDB_Schema, Arg())
If ArgCnt > 0 Then
   For i = 0 To ArgCnt - 1
       MTValuePos = InStr(1, Arg(i), MyGl.INIT_Value)
       If MTValuePos > 0 Then
          MTName = Trim(Left$(Arg(i), MTValuePos - 1))
          MTValue = Trim$(Right$(Arg(i), Len(Arg(i)) - MTValuePos))
       Else     'everything is a name
          MTName = Trim(Arg(i))
          MTValue = ""
       End If
       If Len(MTName) > 0 Then
          Set nv = New NameValue
          nv.Name = MTName
          nv.Value = MTValue
          fAnalysis.MTDB.DBStuff.Add nv, nv.Name
       End If   'do nothing if name is missing
   Next i
End If

Const NAME_MINIMUM_PMT_QUALITY_SCORE As String = "MinimumPMTQualityScore"
EditAddName fAnalysis.MTDB.DBStuff, NAME_MINIMUM_PMT_QUALITY_SCORE, "1"

End Sub

Private Function InitDBConnection() As Boolean
'--------------------------------------------------------
'initializes MT database connection from InitFileName
'returns True if everything cool; False otherwise
'--------------------------------------------------------
On Error GoTo err_InitDBConnection

lblMTDBDesc.Caption = ""
If MTDBInd < 0 Then Exit Function
lblMTDBDesc.Caption = MTDBInfo(MTDBInd).Description & vbCrLf & "State: " & MTDBInfo(MTDBInd).DBState & vbCrLf & "Server: " & MTDBInfo(MTDBInd).Server
With fAnalysis.MTDB.cn
   If .State <> adStateClosed Then .Close
   .ConnectionString = MTDBInfo(MTDBInd).CnStr
End With
fAnalysis.Job = -1      'this is dummy analysis
InitDBConnection = True
Exit Function

err_InitDBConnection:
End Function

Private Sub lstOrgMTDBNames_Click()
'-------------------------------------------------
'initialize database connection with new selection
'of Mass Tags database
'-------------------------------------------------
On Error Resume Next
MTDBInd = MTDBNameListPointers(lstOrgMTDBNames.ListIndex)
If Not InitDBConnection() Then
   lblMTDBDesc.Caption = "Error retrieving database connection information!"
End If
End Sub

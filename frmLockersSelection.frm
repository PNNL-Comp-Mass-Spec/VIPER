VERSION 5.00
Begin VB.Form frmLockersSelection 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lockers Selection"
   ClientHeight    =   2625
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3045
   Icon            =   "frmLockersSelection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCallerID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1620
      Width           =   855
   End
   Begin VB.TextBox txtMinScore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "0"
      Top             =   1020
      Width           =   855
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   80
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      Index           =   5
      X1              =   120
      X2              =   2940
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      Index           =   4
      X1              =   120
      X2              =   2940
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   2940
      Y1              =   1465
      Y2              =   1465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   2940
      Y1              =   2065
      Y2              =   2065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Use only lockers established by:"
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore lockers with score less than:"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   2940
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      Index           =   0
      X1              =   120
      X2              =   2940
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmLockersSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this dialog is used to specify which lockers should be used
'for search
'Criteria on which selection is based could be
'LckType    - ID of type to use (1 for all)
'MinScore   - minimum score of locker to accept
'CallerID   - ID of scientist who called them lockers
'---------------------------------------------------------------
'created: 02/13/2002 nt
'last modified: 02/14/2002 nt
'---------------------------------------------------------------
Option Explicit

'names of database properties(Names in DBStuff list)
'containing relevant lockers selection information
Const NAME_LOCKERS_TYPE = "Locker Type ID"
Const NAME_LOCKERS_MIN_SCORE = "Locker Min Score"
Const NAME_LOCKERS_CALLER_ID = "Locker Caller ID"

Const VERY_LOW_SCORE As Double = -1E+307

'name of pair containing lockers info access query
Const INIT_Fill_Locker_Types = "sql_GET_Locker_Types"

Public MyStuff As Collection
Public MyConnString As String       'connection string for database
Public MyCancel As Boolean

Dim Lck_Type_ID As String           'lockers type
Dim Lck_Min_Score As String         'minimum score for locker
Dim Lck_Caller_ID As String         'caller ID whose lockers we want to use
                                    'empty if it does not matter

Dim bLoading As Boolean             'True until the first activation of form

'locker types IDs and names
Dim LckTypeCnt As Long
Dim LckTypeID() As Long
Dim LckTypeName() As String

Public Event DialogClosed()     'public event raised when this dialog is closed


Private Sub InitSettings()
'-------------------------------------------------------
'loads current settings
'-------------------------------------------------------
Dim Res As Long
Dim i As Long
On Error Resume Next

cmbType.Clear
Res = GetLockerTypes(MyConnString, MyStuff.Item(INIT_Fill_Locker_Types).Value, _
                     LckTypeID(), LckTypeName())
If Res = 0 Then
   LckTypeCnt = UBound(LckTypeID) + 1
   If LckTypeCnt > 0 Then
      For i = 0 To LckTypeCnt - 1
          cmbType.AddItem LckTypeName(i), i
      Next i
   End If
End If
'position current item in the list to the current locker type
Lck_Type_ID = MyStuff.Item(NAME_LOCKERS_TYPE).Value
If Len(Lck_Type_ID) > 0 Then
   For i = 0 To LckTypeCnt - 1
       If CStr(LckTypeID(i)) = Lck_Type_ID Then
          cmbType.ListIndex = i
          Exit For
       End If
   Next i
End If

Lck_Min_Score = MyStuff.Item(NAME_LOCKERS_MIN_SCORE).Value
If Len(Lck_Min_Score) <= 0 Then Lck_Min_Score = CStr(VERY_LOW_SCORE)
txtMinScore.Text = Lck_Min_Score

Lck_Caller_ID = MyStuff.Item(NAME_LOCKERS_CALLER_ID).Value
txtCallerID.Text = Lck_Caller_ID
End Sub

Private Sub cmbType_Click()
Dim Ind As Long
On Error Resume Next
Ind = cmbType.ListIndex
If Ind >= 0 Then Lck_Type_ID = CStr(LckTypeID(Ind))
End Sub

Private Sub cmdCancel_Click()
MyCancel = True
Me.Hide
RaiseEvent DialogClosed
End Sub

Private Sub cmdOK_Click()
'-----------------------------------------------
'accept settings
'-----------------------------------------------
On Error Resume Next
EditAddName NAME_LOCKERS_TYPE, Lck_Type_ID
EditAddName NAME_LOCKERS_MIN_SCORE, Lck_Min_Score
EditAddName NAME_LOCKERS_CALLER_ID, Lck_Caller_ID
Me.Hide
RaiseEvent DialogClosed
End Sub


Private Sub EditAddName(ByVal PairName As String, ByVal NewValue As String)
'-------------------------------------------------------------------------
'modifies value of name value pair; if pair does not exist adds it
'-------------------------------------------------------------------------
Dim nv As NameValue
On Error Resume Next
MyStuff.Item(PairName).Value = NewValue
If Err Then
   Set nv = New NameValue
   nv.Name = PairName
   nv.Value = NewValue
   MyStuff.Add nv, nv.Name
End If
End Sub

Private Sub Form_Activate()
DoEvents
If bLoading Then
   InitSettings
   bLoading = False
End If
End Sub

Private Sub Form_Load()
bLoading = True
End Sub

Private Sub txtCallerID_LostFocus()
Lck_Caller_ID = Trim$(txtCallerID.Text)
End Sub

Private Sub txtMinScore_LostFocus()
Dim TmpMinScore As String
TmpMinScore = Trim$(txtMinScore.Text)
If IsNumeric(TmpMinScore) Then
   Lck_Min_Score = TmpMinScore
Else
   If Len(TmpMinScore) > 0 Then     'can not accept
      MsgBox "This argument should be numeric! Leave it blank to ignore this parametar!", vbOKOnly, MyName
      txtMinScore.SetFocus
   Else         'blank; put some realy negative number to ignore
      Lck_Min_Score = CStr(VERY_LOW_SCORE)
   End If
End If
End Sub

VERSION 5.00
Begin VB.Form frmFilterIsoCom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter Isotopic Composition"
   ClientHeight    =   2130
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Info"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame fraIsoCom 
      Caption         =   "Isotopic Composition"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtIsoComList 
         Height          =   615
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optIsoComFilterType 
         Caption         =   "Include only data labeled as"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optIsoComFilterType 
         Caption         =   "Exclude data labeled as"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkUseIsoComFilter 
         Caption         =   "Use Isotopic Composition Filter"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmFilterIsoCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'created: 02/28/2003 nt
'last modified: 03/01/2003 nt
'-------------------------------------------------------------
'if user selects OK do filtering, Cancel do nothing

'this filter is different from the rest of filters;
'element 0 is the same as in other filters; element 1 is 0
'or 1; element 2 is the list;

'if element 1 is 0 the list is exclusive; if it is 1 the list
'is inclusive (elements of list specify what needs not to be
'filtered out)

'Exclusive list is aggresive (it excludes no matter what is the
'prior inclusive state of the data point); inclusive list is
'non-aggresive(it does not include things just based on this
'criteria - excluded points stay excluded)
'--------------------------------------------------------------
Option Explicit

Const LCL_N14 = "N14"
Const LCL_N15 = "N15"
Const LCL_O16 = "O16"
Const LCL_O18 = "O18"
Const LCL_C12 = "C12"
Const LCL_C13 = "C13"
Const LCL_Natural = "Natural"

Const EXCLUSIVE = 0
Const INCLUSIVE = 1


Const ListDeli = ";"

Dim bLoading As Boolean
Dim CallerID As Long


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Dim Msg As String
Msg = "List should be semicolon delimited! Use terms N14, N15, O16, O18, C12, C13!"
MsgBox Msg, vbOKOnly, glFGTU
End Sub

Private Sub cmdOK_Click()
Dim i As Long
Dim ThisList() As String
On Error Resume Next
Me.MousePointer = vbHourglass
With GelData(CallerID)
    If chkUseIsoComFilter.value = vbChecked Then
       .DataFilter(fltIsoCom, 0) = True
    Else
       .DataFilter(fltIsoCom, 0) = False
    End If
    .DataFilter(fltIsoCom, 1) = -1
    For i = 0 To optIsoComFilterType.UBound
        If optIsoComFilterType(i).value Then .DataFilter(fltIsoCom, 1) = i
    Next i
    If .DataFilter(fltIsoCom, 0) Then
       If .DataFilter(fltIsoCom, 1) >= 0 Then
          If GetList(txtIsoComList.Text, ThisList) Then
             .DataFilter(fltIsoCom, 2) = txtIsoComList.Text
             Select Case .DataFilter(fltIsoCom, 1)
             Case EXCLUSIVE
                  FilterIsoComExclusive CallerID, ThisList
             Case INCLUSIVE
                  FilterIsoComInclusive CallerID, ThisList
             End Select
             GelStatus(CallerID).Dirty = True
          Else
             MsgBox "To use this filter specify semicolon delimited list of charge states to include or exclude!", vbOKOnly, glFGTU
          End If
       Else
          MsgBox "Ambiguous settings! Select exclusive/inclusive option!", vbOKOnly, glFGTU
       End If
    End If
End With
Me.MousePointer = vbDefault
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
If bLoading Then
   CallerID = Me.Tag
   With GelData(CallerID)
        If CBool(.DataFilter(fltIsoCom, 0)) Then
           chkUseIsoComFilter.value = vbChecked
        Else
           chkUseIsoComFilter.value = vbUnchecked
        End If
        optIsoComFilterType(CLng(.DataFilter(fltIsoCom, 1))).value = True
        txtIsoComList.Text = .DataFilter(fltIsoCom, 2)
   End With
   bLoading = False
End If
End Sub

Private Sub Form_Load()
bLoading = True
End Sub

Private Function GetList(ThisString As String, ThisList() As String) As Boolean
'----------------------------------------------------------------------------
'splits list to array and returns True if at least one element in list
'----------------------------------------------------------------------------
On Error Resume Next
ThisList = Split(ThisString, ListDeli)
If UBound(ThisList) >= 0 Then GetList = True
End Function


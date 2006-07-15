VERSION 5.00
Begin VB.Form frmPath 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Path To File Not Found"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   350
      Left            =   4440
      TabIndex        =   4
      Top             =   990
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   4440
      TabIndex        =   3
      Top             =   550
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   350
      Left            =   4440
      TabIndex        =   2
      ToolTipText     =   "Browse to file"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmPath.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 04/20/2000 nt
Option Explicit
Dim OrigPath As String

Public pthCancel As Boolean     'public property

Private Sub cmdBrowse_Click()
'Browsing for folder is work of Brad Martinez (copied from VBnet/Randy Birch)
Dim bi As BROWSEINFO
Dim pidl As Long
Dim FolderPath As String
Dim pos As Integer

bi.hwndOwner = Me.hwnd
bi.pidlRoot = 0&
bi.lpszTitle = "Browse to the Data Folder"
bi.ulFlags = BIF_RETURNONLYFSDIRS
pidl = SHBrowseForFolder(bi)
FolderPath = Space$(MAX_PATH)
If SHGetPathFromIDList(ByVal pidl, ByVal FolderPath) Then
   pos = InStr(FolderPath, Chr$(0))
   txtPath.Text = AddDirSeparator(Left(FolderPath, pos - 1))
   GelData(Me.Tag).PathtoDataFiles = txtPath.Text
End If
Call CoTaskMemFree(pidl)
End Sub

Private Sub cmdCancel_Click()
pthCancel = True
GelData(Me.Tag).PathtoDataFiles = OrigPath
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Activate()
pthCancel = False
txtPath.Text = GelData(Me.Tag).PathtoDataFiles
OrigPath = GelData(Me.Tag).PathtoDataFiles
End Sub

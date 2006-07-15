VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   325
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Caption         =   "Please enter the password"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mInfo As String

Private mPassword As String

Private mPasswordValidated As Boolean
Private mProceedAndCloseForm As Boolean

Private Sub CheckPassword()
    If txtPassword = mPassword Then
        mPasswordValidated = True
    Else
        mPasswordValidated = False
    End If
End Sub

Public Sub Initialize(strMessage As String, strPassword As String)
    mInfo = strMessage
    lblInfo = mInfo
    mPassword = strPassword
    
    mProceedAndCloseForm = False
    mPasswordValidated = False
    txtPassword = ""
    
    ' Set form on top
    Me.ScaleMode = vbTwips
    WindowStayOnTop Me.hwnd, True, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)
End Sub

Public Function PasswordWasValidated() As Boolean
    PasswordWasValidated = mPasswordValidated
End Function

Public Function ProceedAndCloseForm() As Boolean
    ProceedAndCloseForm = mProceedAndCloseForm
End Function

Private Sub cmdCancel_Click()
    mPasswordValidated = False
    mProceedAndCloseForm = True
End Sub

Private Sub cmdOK_Click()
    CheckPassword
    If mPasswordValidated Then
        mProceedAndCloseForm = True
    Else
        lblInfo = "Invalid password.  Please try again or click cancel"
        txtPassword = ""
        txtPassword.SetFocus
    End If
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowUpperThird, 4200, 3500, False
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If lblInfo <> mInfo Then lblInfo = mInfo
End Sub

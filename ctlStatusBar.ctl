VERSION 5.00
Begin VB.UserControl ctlStatusBar 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   ScaleHeight     =   630
   ScaleWidth      =   3210
   Begin VB.Timer tmrTimer 
      Left            =   2760
      Top             =   120
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "ctlStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const UPDATE_INTERVAL_MSEC As Long = 250
Const DIM_CHUNK_SIZE As Long = 10

Private mLastUpdateTime As Date
Private mHoldTimeSec As Long
Private mHoldTimeDays As Double

Private mMessageSepChar As String

Private mMessageStackCountDimmed As Long
Private mMessageStackCount As Long
Private mMessageStack() As String

Public Property Let Message(strNewText As String)
    AddMessageText strNewText
End Property

Public Property Get Message() As String
    Message = lblStatus.Caption
End Property

Public Property Let HoldTime(lngSeconds As Long)
    If lngSeconds < 0 Then lngSeconds = 0
    mHoldTimeSec = lngSeconds
    mHoldTimeDays = mHoldTimeSec / 60 / 60 / 24
End Property
Public Property Get HoldTime() As Long
    HoldTime = mHoldTimeSec
End Property

Public Property Let MessageSepChar(strSepChar As String)
    mMessageSepChar = strSepChar
End Property
Public Property Get MessageSepChar() As String
    MessageSepChar = mMessageSepChar
End Property

Public Function AddMessageText(strNewText As String)
    mMessageStackCount = mMessageStackCount + 1
    If mMessageStackCount > mMessageStackCountDimmed Then
        mMessageStackCountDimmed = mMessageStackCountDimmed + DIM_CHUNK_SIZE
        ReDim Preserve mMessageStack(mMessageStackCountDimmed)
    End If
    mMessageStack(mMessageStackCount - 1) = strNewText
    
    DisplayMessage
    
End Function

Private Sub DisplayMessage()
    Dim strMessage As String
    Dim lngIndex As Long
    
    strMessage = ""
    For lngIndex = 0 To mMessageStackCount - 1
        strMessage = strMessage & mMessageStack(lngIndex)
        If lngIndex < mMessageStackCount - 1 Then
            strMessage = strMessage & mMessageSepChar
        End If
    Next lngIndex
    
    lblStatus.Caption = strMessage
    mLastUpdateTime = Now()
    
    If mMessageStackCount = 0 Then
        tmrTimer.Enabled = False
    Else
        tmrTimer.Enabled = True
    End If
    
End Sub

Private Sub ManageMessageStack()
    Dim lngIndex As Long
    
    If mMessageStackCount > 0 Then
        If Now() - mLastUpdateTime > mHoldTimeDays Then
            ' Shuffle the messages down the stack and update the control
            For lngIndex = 0 To mMessageStackCount - 2
                mMessageStack(lngIndex) = mMessageStack(lngIndex + 1)
            Next lngIndex
            mMessageStackCount = mMessageStackCount - 1
            
            DisplayMessage
        End If
    End If
End Sub

Public Sub ReplaceFullMessage(strNewMessageText As String)
    mMessageStackCount = 0
    
    If Len(strNewMessageText) > 0 Then
        AddMessageText strNewMessageText
    Else
        lblStatus.Caption = ""
        tmrTimer.Enabled = False
    End If
End Sub

Private Sub tmrTimer_Timer()
    ManageMessageStack
End Sub

Private Sub UserControl_Initialize()
    mMessageStackCountDimmed = DIM_CHUNK_SIZE
    ReDim mMessageStack(mMessageStackCountDimmed)
    
    mMessageSepChar = "; "
    
    tmrTimer.Interval = UPDATE_INTERVAL_MSEC
    tmrTimer.Enabled = False
    
    mLastUpdateTime = Now()
    
    ReplaceFullMessage ""
End Sub

Private Sub UserControl_Resize()
    On Error GoTo ResizeControlErrorHandler
    With lblStatus
        .Left = 0
        .Top = 0
        .width = UserControl.width
        .Height = UserControl.Height
    End With
    Exit Sub

ResizeControlErrorHandler:
    Debug.Print "Error in ctlStatusBar->UserControl_Resize: " & Err.Description
    Resume Next
End Sub

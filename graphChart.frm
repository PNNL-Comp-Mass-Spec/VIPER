VERSION 5.00
Object = "{3D931500-4601-11CF-80B2-0020AF19EE14}#5.0#0"; "olch3x32.ocx"
Begin VB.Form frmChart3D 
   Caption         =   "Graphing Chart"
   ClientHeight    =   10770
   ClientLeft      =   2205
   ClientTop       =   450
   ClientWidth     =   13920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   13920
   Begin VB.Frame fraControls 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   9240
      Width           =   5175
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   540
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdZoomOut 
         Caption         =   "&Zoom Out (Ctrl+A)"
         Height          =   540
         Left            =   75
         TabIndex        =   2
         Top             =   120
         Width           =   945
      End
      Begin VB.Frame fraMouseAction 
         Caption         =   "Mouse Action"
         Height          =   645
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Width           =   2700
         Begin VB.ComboBox cboMouseAction 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   2535
         End
      End
   End
   Begin OlectraChart3D.Chart3D Chart3D1 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _Version        =   327680
      _Revision       =   3
      _ExtentX        =   21828
      _ExtentY        =   16325
      _StockProps     =   0
      ControlProperties=   "graphChart.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu menuSaveAsJPG 
         Caption         =   "Save Graph as JPG..."
      End
      Begin VB.Menu menuSaveAsPNG 
         Caption         =   "Save Graph as PNG..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuPrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuEditOptions 
         Caption         =   "Edit Graph Options..."
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmChart3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Windows messages, taken out of the WINAPI.TXT file
'
' Constants for dealing with mouse events
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSELAST = &H209

' Flags set when one of the mouse events is triggered
Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2

' Keyboard events for when a key is pressed/released
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

' Flags set when a keyboard event is triggered
Private Const MK_ALT = &H20
Private Const MK_CONTROL = &H8
Private Const MK_SHIFT = &H4

' The Virtual Key codes
Private Const VK_ESCAPE = &H1B  'The <Esc> key (ASCII Character 27)
Private Const VK_SHIFT = &H10   'The <Shift> key
Private Const VK_CONTROL = &H11 'The <Ctrl> key

Private Enum maMouseActionConstants
    maZoom = 0
    maMove = 1
    maScale = 2
    maRotate = 3
End Enum

Private frmGraphOptions As New frmChart3DOptions
Private mTemporaryActionMode As Boolean
'

Private Sub ComboBoxSetAction()
    Select Case cboMouseAction.ListIndex
    Case maMove
        EnableActionMove
    Case maScale
        EnableActionScale
    Case maRotate
        EnableActionRotate
    Case Else
        ' Includes case maZoom
        EnableActionZoom
    End Select
End Sub

Public Sub EnableActionMove()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints inaccessible
    
    Chart3D1.ActionMaps.RemoveAll
    
    With Chart3D1.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc3dActionModifyStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc3dActionTranslate
        .add WM_LBUTTONUP, 0, 0, oc3dActionModifyEnd
    End With

End Sub

Public Sub EnableActionNone()
    'Clear any previous ActionMaps.
    'Make the rotation constraints inaccessible
    
    Chart3D1.ActionMaps.RemoveAll

End Sub

Public Sub EnableActionRotate()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints accessible
    
    Chart3D1.ActionMaps.RemoveAll
    
    With Chart3D1.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc3dActionModifyStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc3dActionRotate
        .add WM_LBUTTONUP, 0, 0, oc3dActionModifyEnd
    End With
    
End Sub

Public Sub EnableActionScale()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints inaccessible
        
    Chart3D1.ActionMaps.RemoveAll
    
    With Chart3D1.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc3dActionModifyStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc3dActionScale
        .add WM_LBUTTONUP, 0, 0, oc3dActionModifyEnd
    End With
    
End Sub

Public Sub EnableActionZoom()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints inaccessible
    
    Chart3D1.ActionMaps.RemoveAll
    
    With Chart3D1.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc3dActionZoomStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc3dActionZoomUpdate
        .add WM_LBUTTONUP, 0, 0, oc3dActionZoomEnd
        .add WM_KEYDOWN, MK_LBUTTON, VK_ESCAPE, oc3dActionZoomCancel
        
        .add WM_RBUTTONUP, 0, 0, oc3dActionReset
    End With
    
End Sub

Public Function Redraw()
'Alter the current graph's view
On Error GoTo errormsg

MousePointer = vbHourglass
With Chart3D1
    .ChartArea.View3D.ZRotation = Val(frmGraphOptions.inputRotation)
    .ChartArea.View3D.XRotation = Val(frmGraphOptions.inputElevation)
    If Val(frmGraphOptions.inputPerspective) > 0 Then
        .ChartArea.View3D.Perspective = Val(frmGraphOptions.inputPerspective)
    End If
    If Val(frmGraphOptions.inputFontSize) >= 10 Then
        .ChartArea.Axes(1).AnnotationFont.Size = Val(frmGraphOptions.inputFontSize)
    End If
    
    SynchronizeChartOptions
End With
MousePointer = vbDefault

Exit Function

errormsg:
MsgBox "You have entered an invalid number. Please stay within the range.", vbOKOnly + vbExclamation, "Invalid Input."
MousePointer = vbDefault

End Function

Public Sub ResetGraph()
Chart3D1.CallAction oc3dActionReset, 0, 0
End Sub

Public Sub ResetToDefaults()
    With Chart3D1
        .IsBatched = True
        With .ChartArea
            
            .Viewport.HorizontalShift = 0
            .Viewport.ScaleFactor = 1
            .Viewport.VerticalShift = 0
            
            With .View3D
                .XRotation = 30
                .YRotation = 0
                .ZRotation = 225
                .Perspective = 2.5
            End With
        End With
        .IsBatched = False
    End With
    
    SynchronizeChartOptions
    
End Sub

Public Sub SaveGraphPicture(blnSaveAsPNG As Boolean, Optional strFilepath As String = "")
'Save the chart as a JPG or PNG file

Dim strPictureFormat As String
Dim strPictureExtension As String

On Error GoTo SaveGraphPictureErrorHandler

If blnSaveAsPNG Then
    strPictureFormat = "PNG"
    strPictureExtension = ".png"
Else
    strPictureFormat = "JPG"
    strPictureExtension = ".jpg"
End If

If Len(strFilepath) = 0 Then
    strFilepath = SelectFile(Me.hwnd, "Enter filename", "", True, "MassDiffs" & strPictureExtension, strPictureFormat & " Files (*." & strPictureExtension & ")|*." & strPictureExtension & "|All Files (*.*)|*.*")
End If

If Len(strFilepath) > 0 Then
    strFilepath = FileExtensionForce(strFilepath, strPictureExtension)

    If blnSaveAsPNG Then
        If Not Chart3D1.SaveImageAsPng(strFilepath, False) Then
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "Could not save file for some reason.", vbOKOnly + vbExclamation, "Error Saving File."
            End If
        End If
    Else
        If Not Chart3D1.SaveImageAsJpeg(strFilepath, 90, False, True, False) Then
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "Could not save file for some reason.", vbOKOnly + vbExclamation, "Error Saving File."
            End If
        End If
    End If
End If

Exit Sub
SaveGraphPictureErrorHandler:

End Sub

Private Sub SynchronizeChartOptions()
    ' Update .Graph3DOptions
    With glbPreferencesExpanded.ErrorPlottingOptions.Graph3DOptions
        .ZRotation = Chart3D1.ChartArea.View3D.ZRotation
        .Elevation = Chart3D1.ChartArea.View3D.XRotation
        .Perspective = Chart3D1.ChartArea.View3D.Perspective
        .AnnotationFontSize = Chart3D1.ChartArea.Axes(1).AnnotationFont.Size
    End With

End Sub

Private Sub cboMouseAction_Click()
    ComboBoxSetAction
End Sub

Private Sub cmdReset_Click()
    ResetToDefaults
End Sub

Private Sub cmdZoomOut_Click()
    ResetGraph
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' Note: KeyPreview must be enabled on this form for this code to execute

'    Debug.Print "Keycode = " & KeyCode & " and shift = " & Shift
    If Shift = 2 Then
        ' Ctrl pressed
        If KeyCode = 65 Then
            ' Ctrl+A pressed; Reset graph
            ResetGraph
        End If
    ElseIf KeyCode = vbKeySpace Then
        mTemporaryActionMode = True
        EnableActionRotate
    End If
End Sub

Private Sub form_KeyUp(KeyCode As Integer, Shift As Integer)
    If mTemporaryActionMode Then
        ComboBoxSetAction
        mTemporaryActionMode = False
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 600 * Screen.TwipsPerPixelX
    Me.width = 800 * Screen.TwipsPerPixelY
    
    ' Populate the combo box
    cboMouseAction.AddItem "Zoom"
    cboMouseAction.AddItem "Move"
    cboMouseAction.AddItem "Scale"
    cboMouseAction.AddItem "Rotate (space+drag)"
    cboMouseAction.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Dim lngDesiredValue As Long
    
    On Error Resume Next
    With Chart3D1
        .Left = 0
        .Top = 0
        .width = Me.ScaleWidth
        lngDesiredValue = Me.ScaleHeight - fraControls.Height - 180
        If lngDesiredValue < 1000 Then lngDesiredValue = 1000
        .Height = lngDesiredValue
        
        fraControls.Top = .Top + .Height + 120
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmGraphOptions
End Sub


Private Sub menuPrint_Click()
'Print the chart
MousePointer = vbHourglass
If MsgBox("Proceed with printing the graph?", vbQuestion + vbYesNoCancel, "Print") = vbYes Then
    If Chart3D1.PrintChart(oc3dFormatBitmap, oc3dScaleToFit, 0, 0, 0, 0) = True Then
        MsgBox "Chart successfully printed.", vbInformation + vbOKOnly
    End If
End If
MousePointer = vbDefault
End Sub

Private Sub menuSaveAsJPG_Click()
    SaveGraphPicture False, ""
End Sub

Private Sub menuSaveAsPNG_Click()
    SaveGraphPicture True, ""
End Sub

Private Sub mnuEditOptions_Click()
'Show graphOptions form to edit the graph's viewing angles
frmGraphOptions.inputRotation = Chart3D1.ChartArea.View3D.ZRotation
frmGraphOptions.inputElevation = Chart3D1.ChartArea.View3D.XRotation
frmGraphOptions.inputPerspective = Chart3D1.ChartArea.View3D.Perspective
frmGraphOptions.inputFontSize = Chart3D1.ChartArea.Axes(1).AnnotationFont.Size

Do Until frmGraphOptions.goClose
    frmGraphOptions.Show vbModal
    DoEvents
    If Not frmGraphOptions.goClose Then Redraw
Loop

frmGraphOptions.goClose = False
End Sub

Private Sub mnuExit_Click()
Me.Hide
Unload frmGraphOptions
End Sub


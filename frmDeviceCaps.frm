VERSION 5.00
Begin VB.Form frmDeviceCaps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Device Caps"
   ClientHeight    =   3120
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDeviceCaps 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "&Printer"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdScreen 
      Caption         =   "&Screen"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDeviceCaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'implements interface to device caps and displays short info about
'screen and printer devices
'created: 12/09/2002 nt
'last modified: 12/09/2002 nt
'-----------------------------------------------------------
Option Explicit

Public ParentDC As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrinter_Click()
txtDeviceCaps.Text = LoadDeviceInfo(Printer.hDC)
End Sub

Private Sub cmdScreen_Click()
txtDeviceCaps.Text = LoadDeviceInfo(ParentDC)
End Sub

Private Function LoadDeviceInfo(hDC As Long) As String
Dim Tmp As String
Dim lDF As Long
On Error GoTo err_LoadDeviceInfo

lDF = GetDeviceCaps(hDC, TECHNOLOGY)
If lDF And DT_RASPRINTER Then Tmp = "Raster Printer"
If lDF And DT_PLOTTER Then Tmp = "Plotter"
If lDF And DT_RASDISPLAY Then Tmp = "Raster Display"
If lDF And DT_RASCAMERA Then Tmp = "Raster Camera"
If Len(Tmp) = 0 Then Tmp = "Other technology"
Tmp = Tmp & vbCrLf
Tmp = Tmp & "X,Y dimensions in millimeters: " & Str$(GetDeviceCaps(hDC, HORZSIZE)) _
    & ", " & Str$(GetDeviceCaps(hDC, VERTSIZE)) & vbCrLf
Tmp = Tmp & "X,Y dimensions in pixels: " & Str$(GetDeviceCaps(hDC, HORZRES)) _
    & ", " & Str$(GetDeviceCaps(hDC, VERTRES)) & vbCrLf
Tmp = Tmp & "X,Y pixels/logical inch: " & Str$(GetDeviceCaps(hDC, LOGPIXELSX)) _
    & ", " & Str$(GetDeviceCaps(hDC, LOGPIXELSY)) & vbCrLf
Tmp = Tmp & "Bits/Pixels: " & Str$(GetDeviceCaps(hDC, BITSPIXEL)) & vbCrLf
Tmp = Tmp & "Color planes: " & Str$(GetDeviceCaps(hDC, PLANES)) & vbCrLf
Tmp = Tmp & "Color table entries: " & Str$(GetDeviceCaps(hDC, NUMCOLORS)) & vbCrLf
Tmp = Tmp & "Aspect X,Y,XY: " & Str$(GetDeviceCaps(hDC, ASPECTX)) _
    & ", " & Str$(GetDeviceCaps(hDC, ASPECTY)) & ", " & Str$(GetDeviceCaps(hDC, ASPECTXY)) & vbCrLf

Tmp = Tmp & vbCrLf & "Device capabilities:" & vbCrLf
lDF = GetDeviceCaps(hDC, RASTERCAPS)
If lDF And RC_BANDING Then Tmp = Tmp & "Banding" & vbCrLf
If lDF And RC_BIGFONT Then Tmp = Tmp & "Font>64K" & vbCrLf
If lDF And RC_BITBLT Then Tmp = Tmp & "BitBlt" & vbCrLf
If lDF And RC_BITMAP64 Then Tmp = Tmp & "Bitmaps>64K" & vbCrLf
If lDF And RC_DI_BITMAP Then Tmp = Tmp & "Device Independent Bitmaps" & vbCrLf
If lDF And RC_DIBTODEV Then Tmp = Tmp & "DIB to device" & vbCrLf
If lDF And RC_FLOODFILL Then Tmp = Tmp & "Flood fill" & vbCrLf
If lDF And RC_SCALING Then Tmp = Tmp & "Scaling" & vbCrLf
If lDF And RC_STRETCHBLT Then Tmp = Tmp & "StretchBlt" & vbCrLf
If lDF And RC_STRETCHDIB Then Tmp = Tmp & "StretchDIB" & vbCrLf
LoadDeviceInfo = Tmp
Exit Function

err_LoadDeviceInfo:
LoadDeviceInfo = "Error retrieving device information!" & vbCrLf _
                & "Error: " & Err.Number & "; " & Err.Description
End Function


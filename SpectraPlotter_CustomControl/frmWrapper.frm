VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Spectra Plotter Test"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin SpectraPlotterProj.ctlSpectraPlotter ctlSpectraPlotter1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15055
      _ExtentY        =   10610
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ctlSpectraPlotter1.SetLabelXAxis "m/z"
    ctlSpectraPlotter1.SetLabelYAxis "Intensity"
End Sub

Private Sub Form_Resize()
    Dim lngDesiredValue As Long
    
    lngDesiredValue = Me.ScaleWidth - 250
    If lngDesiredValue < 0 Then lngDesiredValue = 0
    ctlSpectraPlotter1.Width = lngDesiredValue
    
    lngDesiredValue = Me.ScaleHeight - 250
    If lngDesiredValue < 0 Then lngDesiredValue = 0
    ctlSpectraPlotter1.Height = lngDesiredValue
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

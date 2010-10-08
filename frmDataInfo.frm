VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmDataInfo 
   Caption         =   "Data Info"
   ClientHeight    =   3615
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8265
   Icon            =   "frmDataInfo.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbData 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDataInfo.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveText 
         Caption         =   "Save As &Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Selected"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy &All"
      End
   End
End
Attribute VB_Name = "frmDataInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'displays PEK/CSV/mzXML/mzData file                      (Tag property is gel index)
'gel data structured as in gel file     (Tag property is empty)
'file info                              (Tag property is -(gel index))
'mass tags settings(MTDB DB Stuff)      (Tag property is 1000000+(gel index)
'AMT search results                     (Tag property is "AMT")
'UMC report                             (Tag property is "UMC")
'Lock Mass Function reports             (Tag property is "AMTLM")
'ORF statistics report                  (Tag property is "ORF")
'N14/N15 pairs report                   (Tag property is "N14_N15")
'Lockers report                         (Tag property is "LOCKERS")
'Lockers matching results               (Tag property is "LCK")
'N14/N15 pairs with label               (tag property is "Dlt_Lbl")
'Cys-labeled pairs                      (tag property is "Cys_Lbl")
'ER statistics                          (tag property is "ER_STAT")
'Attention List                         (tag property is "AL")
'distribution function                  (tag property is "DF")
'NET adjustments                        (tag property is "AdjNET")
'Miscellaneous reports                  (tag property is "Misc")
'All reports are written in temporary text file that is destroyed
'when form unloads (can be saved)

'last modified: 07/02/2002 nt
'---------------------------------------------------------------
Option Explicit

Private mSourceFilePath As String

Private mTmpFile As String

Private mNameFromCaption As String   'another name to suggest name

Private CallerID

Public Property Let SourceFilePath(ByVal value As String)
    mSourceFilePath = value
End Property

Public Property Get SourceFilePath() As String
    SourceFilePath = mSourceFilePath
End Property

Private Sub CopyAll()
    
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText rtbData.Text, vbCFText
    
End Sub

Private Sub CopySelected()
    
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Mid(rtbData.Text, rtbData.SelStart, rtbData.SelLength), vbCFText
    
End Sub

Private Sub Form_Activate()
On Error Resume Next
mNameFromCaption = Trim$(Me.Caption)
CallerID = Me.Tag
If Len(CallerID) > 0 Then
   If IsNumeric(CallerID) Then
      If CallerID > 0 Then
         If CallerID >= 1000000 Then        'display MTDB stuff
            Me.Caption = "Data Info - MT tag database Settings"
            rtbData.Text = GelAnalysis(CallerID - 1000000).MTDB.GetDBStuff
         Else                               'display PEK/CSV/mzXML/mzData file
            Me.Caption = "Data Info - PEK/CSV/mzXML/mzData File"
            rtbData.loadFile GelData(CallerID).FileName, rtfText
         End If
      ElseIf CallerID < 0 Then
         Me.Caption = "File Info"
         rtbData.Text = GetInfoString()
      Else
         rtbData.Text = GetNoInfo()
      End If
   Else
      Select Case Me.Tag
      Case "AMT"
        Me.Caption = "AMT Search Results"
      Case "AMTLM"
        Me.Caption = "AMT Lock Mass Report"
      Case "ORF"
        Me.Caption = "ORF Statistics report"
      Case "UMC"
        Me.Caption = "UMC Report"
      Case "N14_N15"
        Me.Caption = "N14/N15 Pairs Report"
      Case "MTLM"
        Me.Caption = "MT tag DB Lockers Report"
      Case "LOCKERS"
        Me.Caption = "Lockers Matching Results"
      Case "INCALLM"
        Me.Caption = "Internal Calibration Lock Mass Report"
      Case "Dlt_Lbl"
        Me.Caption = "Delta Label Pairs"
      Case "Cys_Lbl"
        Me.Caption = "Cys-Labeled Pairs"
      Case "ER_STAT"
        Me.Caption = "Expression Ratio Statistics"
      Case "AL"
        Me.Caption = "Attention List"
      Case "DF"
        Me.Caption = "Distribution Function"
      Case "AdjNET"
        Me.Caption = "NET Adjustment Report"
      Case "Misc"
        Me.Caption = "Miscellaneous Reports"
      Case "UMC_MTID"
        Me.Caption = "UMC - MT tags Association"
      Case "ANALYSIS_HISTORY"
        Me.Caption = "Analysis History Log"
      Case "STAC_Stats"
        Me.Caption = "STAC Match Stats"
      End Select
      If Me.SourceFilePath = "" Then
          mTmpFile = GetTempFolder() & RawDataTmpFile
      Else
          mTmpFile = Me.SourceFilePath
      End If
      rtbData.loadFile mTmpFile, rtfText
      mnuFileSaveText.Enabled = True
   End If
Else
   Me.Caption = "Data Info - GEL File"
   mTmpFile = GetTempFolder() & RawDataTmpFile
   rtbData.loadFile mTmpFile, rtfText
   mnuFileSaveText.Enabled = True
End If
End Sub

Private Sub Form_Resize()
With rtbData
  .width = Me.ScaleWidth
  .Height = Me.ScaleHeight
  If .width < 8000 Then
      .RightMargin = 5 * .width / 4   'do not want wordwrap
  Else
    .RightMargin = .width - 240
  End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If (Not IsNumeric(CallerID)) Then  'don't kill PEK/CSV/mzXML/mzData file
   Kill mTmpFile
End If
End Sub

Private Sub mnuCopy_Click()
CopySelected
End Sub

Private Sub mnuCopyAll_Click()
CopyAll
End Sub

Private Sub mnuFileClose_Click()
Unload Me
End Sub

Private Sub mnuFilePrint_Click()
On Error GoTo err_mnuFilePrint
Printer.Print ""
rtbData.SelPrint (Printer.hDC)
Printer.EndDoc
Exit Sub

err_mnuFilePrint:
MsgBox "Error: " & Err.Number & "; " & Err.Description, vbOKOnly, "2DGelLand"
End Sub
         
Private Sub mnuFileSaveText_Click()
Dim sFN As String
Dim sSuggestedFN As String
On Error GoTo err_mnuFileSaveText
If Len(mNameFromCaption) > 0 Then
   sSuggestedFN = mNameFromCaption
Else
   sSuggestedFN = SuggestionByName(Me.Tag, "txt")
End If
'sFN = SaveFileAPIDlg(Me.hWnd, "Text file (*.txt)" & Chr(0) & "*.txt" _
'            & Chr(0), 1, sSuggestedFN, "Save Text File")

sFN = SelectFile(Me.hwnd, "Save Text File", "", True, sSuggestedFN, , 2)

DoEvents
If Len(sFN) > 0 Then rtbData.SaveFile sFN, rtfText
Exit Sub

err_mnuFileSaveText:
MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Virtual 2D Gels"
End Sub

Private Function GetInfoString() As String
Dim sTmp As String
Dim i As Long
Dim MaxInd As Long
Dim aInfo As Variant
aInfo = Fileinfo(-CallerID, 3)
If IsNull(aInfo) Or Not IsArray(aInfo) Then
   GetInfoString = GetNoInfo()
Else
   MaxInd = UBound(aInfo)
   If MaxInd <= 0 Then
      GetInfoString = GetNoInfo()
   Else
      sTmp = aInfo(0)
      For i = 1 To MaxInd
          sTmp = sTmp & vbCrLf & aInfo(i)
      Next i
      If Len(GelData(-CallerID).Comment) > 0 Then sTmp = sTmp & vbCrLf & vbCrLf & "COMENT" & vbCrLf & GelData(-CallerID).Comment
      GetInfoString = sTmp
   End If
End If
End Function

Private Function GetNoInfo() As String
Dim sTmp As String
sTmp = "No information available!!!"
sTmp = sTmp & vbCrLf & "File corrupted or unappropriate file format."
sTmp = sTmp & vbCrLf & "Close file and try to open again."
GetNoInfo = sTmp
End Function

Private Sub rtbData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuEdit
    End If
End Sub

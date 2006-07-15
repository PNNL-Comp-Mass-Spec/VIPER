VERSION 5.00
Begin VB.UserControl LaDist 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   ScaleHeight     =   3795
   ScaleWidth      =   4845
   Begin VB.ComboBox cmbBinType 
      Height          =   315
      ItemData        =   "LaDist.ctx":0000
      Left            =   3480
      List            =   "LaDist.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox txtResolution 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Text            =   "369"
      Top             =   60
      Width           =   615
   End
   Begin VB.TextBox txtMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   60
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "0"
      Top             =   60
      Width           =   615
   End
   Begin VB.PictureBox picD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   440
      Width           =   4815
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Res."
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Range"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuDF 
      Caption         =   "Distribution"
      Visible         =   0   'False
      Begin VB.Menu mnuDFGraphType 
         Caption         =   "&Line"
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuDFGraphType 
         Caption         =   "&Bar"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuDFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDFCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "LaDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Customized from DistGraph (nt) control
'function distribution user control
'--------------------------------------------------------
'last modified: 07/23/2002 nt
'--------------------------------------------------------
Option Explicit

Const MSG_TITLE = "2D Display - Distribution Function"

Const MAXDATACOUNT = 10000000       'maximum data points
Const MAXRESOLUTION = 32768         'maximum resolution
'DEFINITION: Data resolution represents number of bins used
'to interpret data; for example data resolution of 2 means
'two bins between min and max value and two bins for out of
'range values; therefore data resolution 2 requires 3 bins
'borders and 4 frequency array members

'For ratio bins Min=1/Max therefore only Max and Resolution
'are needed to produce bins

'Out of range points are not drawn - they are indicated in
'a status bar

'''Const DEFMAXFREQ = 100

Public Enum eDFBinsType
    BinsUni = 0
    BinsLog = 1
    BinsRat = 2
    BinsInt = 3     'not implemented
End Enum

Public Enum eDFGraphType
    GraphLine = 0
    GraphBar = 1
End Enum

Public VLabel As String     'label on vertical axis
Public VNumFmt As String    'numerical format for vertical axis
Public HLabel As String     'label on horizontal axis
Public HNumFmt As String    'numerical format for horizontal axis

Public DFGraphType As eDFGraphType          'bar or line
Public DFBinsType As eDFBinsType            'bins type

Dim mMin As Double            'minimum of data range
Dim mMax As Double            'maximum of data range
Dim mRes As Long              'data resolution
Dim mDelta As Double          'basically bins width (in case of ratio it is
                              'bin width right of 1 - not uniform left of 1)
                              
Dim mDataCnt As Long
Dim mData() As Double         'data that has to be drawn

Dim mBins() As Double         'actual bin borders
Dim mFreq() As Long           'frequencies

Dim mFreqMax As Long          'maximum frequency (other than out of range bins)
Dim mFreqMaxInd As Long       'index of maximum frequency in mFreq

Dim mScaleH As Double         'scales used to draw
Dim mScaleV As Double         'on logical window

Dim mIsFree As Boolean        'read only properety; True if control
                              'is free for assignment; False if it is
                              'filled with data

'coordinate system-viewport coordinates
Dim VPX0 As Long
Dim VPY0 As Long
Dim VPXE As Long
Dim VPYE As Long

'actual drawing points
Dim mgH() As Long
Dim mgV() As Long

'control of changes in arguments
Dim mMinChg As Boolean
Dim mMaxChg As Boolean
Dim mResChg As Boolean
Dim mBinChg As Boolean

Public Event ArgChange()    'raised when control lost focus
                            'and one of the arguments change

Public Function DataFill(ByRef Data() As Double, ByVal MinVal As Double, _
                         ByVal MaxVal As Double, ByVal Res As Long, _
                         ByVal BinsType As eDFBinsType) As Boolean
'-------------------------------------------------------------------------
'fills arrays and attempts to draw distribution
'-------------------------------------------------------------------------
On Error GoTo err_DataFill
If UBound(Data) > MAXDATACOUNT Then
   UpdateStatus "Too many data points!"
   Exit Function
End If
If Not mIsFree Then Clear
mData = Data
mDataCnt = UBound(mData) + 1
mMin = MinVal
txtMin.Text = MinVal
mMax = MaxVal
txtMax.Text = MaxVal
mRes = Res
txtResolution.Text = Res
DFBinsType = BinsType
cmbBinType.ListIndex = DFBinsType   'this will trigger refresh if change
DFRefresh
mIsFree = False
DataFill = True
Exit Function

err_DataFill:
UpdateStatus "Error loading data!"
End Function

Public Sub Clear()
'---------------------------------------------------
'erase arrays and clean picture; declare object free
'---------------------------------------------------
On Error Resume Next
Erase mData
Erase mBins
Erase mFreq
mDataCnt = 0
picD.Cls
mIsFree = True
End Sub


Public Sub Draw()
'-----------------------------------------------------
'draws distribution function on picture device context
'-----------------------------------------------------
Dim OldDC As Long
Dim Res As Long
OldDC = SaveDC(picD.hDC)
DGCooSys picD.hDC, picD.ScaleWidth, picD.ScaleHeight
DGDrawCooSys picD.hDC
WriteLabels picD.hDC
Select Case DFGraphType
Case GraphLine
     DrawDFLine picD.hDC
Case GraphBar
     DrawDFBar picD.hDC
End Select
Res = RestoreDC(picD.hDC, OldDC)
End Sub


Private Sub DGCooSys(ByVal hDC As Long, ByVal dcw As Long, ByVal dch As Long)
'----------------------------------------------------------------------------
'establishes coordinate system on device context
'----------------------------------------------------------------------------
Dim Res As Long
Dim ptPoint As POINTAPI
Dim szSize As Size
On Error Resume Next
VPX0 = 25
VPY0 = dch - 25
VPXE = dcw - 40
VPYE = 40 - dch
Res = SetMapMode(hDC, MM_ANISOTROPIC)
'logical window
Res = SetWindowOrgEx(hDC, LDfX0, LDfY0, ptPoint)
Res = SetWindowExtEx(hDC, LDfXE, LDfYE, szSize)
'viewport
Res = SetViewportOrgEx(hDC, VPX0, VPY0, ptPoint)
Res = SetViewportExtEx(hDC, VPXE, VPYE, szSize)
End Sub

Public Sub DGDrawCooSys(ByVal hDC As Long)
'-----------------------------------------
'draws coordinate system on device context
'-----------------------------------------
Dim ptPoint As POINTAPI
Dim Res As Long
On Error Resume Next
'horizontal
Res = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
Res = LineTo(hDC, LDfX0 + LDfXE, LDfY0)
'vertical
Res = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
Res = LineTo(hDC, LDfX0, LDfY0 + LDfYE)
End Sub

Private Sub PositionControls()
    Const MIN_PIC_SIZE = 1000
    Dim lngDesiredValue As Long
    
    With picD
        .Left = 0
        .Top = 440
        
        lngDesiredValue = UserControl.ScaleHeight - lblStatus.Height - .Top - 240
        If lngDesiredValue < MIN_PIC_SIZE Then lngDesiredValue = MIN_PIC_SIZE
        .Height = lngDesiredValue
    
        lngDesiredValue = UserControl.ScaleWidth - .Left - 120
        If lngDesiredValue < MIN_PIC_SIZE Then lngDesiredValue = MIN_PIC_SIZE
        .width = lngDesiredValue
    End With
    
    With lblStatus
        .Top = picD.Top + picD.Height + 60
        .Left = 120
        .width = picD.width
    End With
End Sub

Private Sub cmbBinType_Change()
mBinChg = True
End Sub

Private Sub cmbBinType_Click()
DFBinsType = cmbBinType.ListIndex
If mBinChg Then
   DFRefresh
   mBinChg = False
End If
End Sub

Private Sub mnuDFCopy_Click()
Call CopyFD
End Sub

Private Sub mnuDFGraphType_Click(Index As Integer)
mnuDFGraphType(Index).Checked = True
mnuDFGraphType((Index + 1) Mod 2).Checked = False
DFGraphType = Index
picD.Cls
Call Draw
End Sub

Private Sub picD_KeyDown(KeyCode As Integer, Shift As Integer)
'copy as metafile on system clipboard
Dim CtrlDown As Boolean
CtrlDown = (Shift And vbCtrlMask) > 0
If KeyCode = vbKeyC And CtrlDown Then CopyFD
End Sub

Private Sub picD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then UserControl.PopupMenu mnuDF
End Sub

Private Sub txtMax_Change()
mMaxChg = True
End Sub

Private Sub txtMax_GotFocus()
mMaxChg = False
End Sub

Private Sub txtMax_LostFocus()
If mMaxChg Then
   If IsNumeric(txtMax.Text) Then
      mMax = CDbl(txtMax.Text)
      Exit Sub
   End If
   MsgBox "This argument should be a number!", vbOKOnly, MSG_TITLE
   txtMax.SetFocus
End If
End Sub

Private Sub txtMin_Change()
mMinChg = True
End Sub

Private Sub txtMin_GotFocus()
mMinChg = False
End Sub

Private Sub txtMin_LostFocus()
If mMinChg Then
   If IsNumeric(txtMin.Text) Then
      mMin = txtMin.Text
      Exit Sub
   End If
   MsgBox "This argument should be a number!", vbOKOnly, MSG_TITLE
   txtMin.SetFocus
End If
End Sub

Private Sub txtResolution_Change()
mResChg = True
End Sub

Private Sub txtResolution_GotFocus()
mResChg = False
End Sub

Private Sub txtResolution_LostFocus()
Dim TmpRes As String
If mResChg Then
   TmpRes = txtResolution.Text
   If IsNumeric(TmpRes) Then
      If TmpRes > 0 Then
         If TmpRes <= MAXRESOLUTION Then
            mRes = CLng(TmpRes)
            Exit Sub
         End If
      End If
    End If
    MsgBox "Resolution should be positive integer up to " & MAXRESOLUTION & "!", vbOKOnly, MSG_TITLE
    txtResolution.SetFocus
End If
End Sub

Private Sub UserControl_ExitFocus()
If (mMinChg Or mMaxChg Or mResChg) Then RaiseEvent ArgChange
End Sub

Private Sub UserControl_Initialize()
VLabel = "?"
HLabel = "?"
mIsFree = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
DFBinsType = PropBag.ReadProperty("DFBinsType", BinsUni)
DFGraphType = PropBag.ReadProperty("DFGraphType", GraphLine)
End Sub

Private Sub UserControl_Resize()
    PositionControls
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty "DFBinsType", DFBinsType, BinsUni
PropBag.WriteProperty "DFGraphType", DFGraphType, GraphLine
End Sub


Private Sub CalculateGraphPoints()
Dim i As Long
On Error Resume Next
Select Case DFBinsType
Case BinsUni, BinsLog
    mScaleV = LDfYE / mFreqMax
    mScaleH = LDfXE / (mMax - mMin)

    ReDim mgH(mRes - 1)
    ReDim mgV(mRes - 1)
    For i = 1 To mRes
        mgH(i) = CLng((i * mDelta) * mScaleH)
        mgV(i) = CLng(mFreq(i) * mScaleV)
    Next i
Case BinsRat    'have to draw with indexes 'cause bins are not uniform
    mScaleV = LDfYE / mFreqMax
    mScaleH = LDfXE / mRes
    ReDim mgH(mRes - 1)
    ReDim mgV(mRes - 1)
    For i = 1 To mRes
        mgH(i) = CLng(i * mScaleH)
        mgV(i) = CLng(mFreq(i) * mScaleV)
    Next i
End Select
End Sub

Public Sub DrawDFLine(ByVal hDC As Long)
'---------------------------------------
'draws lines to device context
'---------------------------------------
Dim ptPoint As POINTAPI
Dim Res As Long
Dim i As Long
On Error Resume Next
Res = MoveToEx(hDC, mgH(0), mgV(0), ptPoint)
For i = 0 To mRes - 1
    Res = LineTo(hDC, mgH(i), mgV(i))
Next i
End Sub

Public Sub DrawDFBar(ByVal hDC As Long)
'--------------------------------------
'draws bars to device context
'--------------------------------------
Dim ptPoint As POINTAPI
Dim Res As Long
Dim i As Long
On Error Resume Next
For i = 0 To mRes - 1
    Res = MoveToEx(hDC, mgH(i), 0, ptPoint)
    Res = LineTo(hDC, mgH(i), mgV(i))
Next i
End Sub

Public Property Get MaxDataValue() As Double
MaxDataValue = mMax
End Property

Public Property Get MinDataValue() As Double
MinDataValue = mMin
End Property

Public Property Get Resolution() As Long
Resolution = mRes
End Property

Public Property Get IsFree() As Boolean
IsFree = mIsFree
End Property

Private Sub WriteLabels(ByVal hDC As Long)
Dim lFont As Long
Dim lfLogFont As LOGFONT
Dim lOldFont As Long
Dim lNewFont As Long
Dim Res As Long
On Error Resume Next
'get the font from the picture box control (Arial Narrow)
lFont = SelectObject(picD.hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lFont, Len(lfLogFont), lfLogFont)
Res = SelectObject(picD.hDC, lFont)

'create new logical font
lfLogFont.lfWidth = 180
lfLogFont.lfHeight = 900
lNewFont = CreateFontIndirect(lfLogFont)

'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
'draw coordinate axes labels
Res = TextOut(hDC, 9600, -800, HLabel, Len(HLabel))
Res = TextOut(hDC, 50, 10700, VLabel, Len(VLabel))
'draw numeric markers on vertical axes
WriteVMarkers hDC
WriteHMarkers hDC
WriteHighPeakInfo hDC
'restore old font to hdc
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
End Sub


Public Sub CopyFD()
Dim Res As Long

Dim OldDC As Long
Dim hRefDC As Long
Dim emfDC As Long       'metafile device context
Dim emfHandle As Long   'metafile handle

Dim iWidthMM As Long
Dim iHeightMM As Long
Dim iWidthPels As Long
Dim iHeightPels As Long
Dim iMMPerPelX As Double
Dim iMMPerPelY As Double
Dim rcRef As Rect           'reference rectangle
On Error Resume Next


hRefDC = picD.hDC
iWidthMM = GetDeviceCaps(hRefDC, HORZSIZE)
iHeightMM = GetDeviceCaps(hRefDC, VERTSIZE)
iWidthPels = GetDeviceCaps(hRefDC, HORZRES)
iHeightPels = GetDeviceCaps(hRefDC, VERTRES)

iMMPerPelX = (iWidthMM * 100) / iWidthPels
iMMPerPelY = (iHeightMM * 100) / iHeightPels

rcRef.Top = 0
rcRef.Left = 0
rcRef.Bottom = picD.ScaleHeight
rcRef.Right = picD.ScaleWidth
'convert to himetric units
rcRef.Left = rcRef.Left * iMMPerPelX
rcRef.Top = rcRef.Top * iMMPerPelY
rcRef.Right = rcRef.Right * iMMPerPelX
rcRef.Bottom = rcRef.Bottom * iMMPerPelY

emfDC = CreateEnhMetaFile(hRefDC, vbNullString, rcRef, vbNullString)
OldDC = SaveDC(emfDC)

DGCooSys emfDC, picD.ScaleWidth, picD.ScaleHeight
DGDrawCooSys emfDC
WriteLabels emfDC
Select Case DFGraphType
Case GraphLine
     DrawDFLine emfDC
Case GraphBar
     DrawDFBar emfDC
End Select
Res = RestoreDC(emfDC, OldDC)
emfHandle = CloseEnhMetaFile(emfDC)

Res = OpenClipboard(picD.hwnd)
Res = EmptyClipboard()
Res = SetClipboardData(CF_ENHMETAFILE, emfHandle)
Res = CloseClipboard
End Sub

Private Sub WriteVMarkers(ByVal hDC As Long)
'--------------------------------------------------
'maximum number of vertical labels is 12, minimum 5
'--------------------------------------------------
Dim LblCnt As Integer
Dim VDeltaD As Double
Dim VDelta As Long
Dim LDelta As Long
Dim Done As Boolean
Dim VLbl As String
Dim szLbl As Size
Dim Res As Long
Dim i As Long
LblCnt = 5
Do Until Done
   VDeltaD = mFreqMax / (LblCnt - 1)
   VDelta = CLng(VDeltaD)
   If Abs(VDeltaD - VDelta) < 0.001 Then
      Done = True
      LDelta = LDfYE / (LblCnt - 1)
   Else
      LblCnt = LblCnt + 1
      If LblCnt > 12 Then
         LblCnt = 2
         VDelta = mFreqMax
         LDelta = LDfYE
         Done = True
      End If
   End If
Loop
For i = 0 To LblCnt - 1
    VLbl = Format$(CStr(i * VDelta), VNumFmt)
    Res = GetTextExtentPoint32(hDC, VLbl, Len(VLbl), szLbl)
    Res = TextOut(hDC, -szLbl.cx - 100, i * LDelta + 400, VLbl, Len(VLbl))
Next i
End Sub


Private Sub WriteHMarkers(ByVal hDC As Long)
Dim HDelta As Long
Dim LDelta As Long
Dim HLbl As String
Dim szLbl As Size
Dim Res As Long
Dim i As Long
LDelta = LDfXE / 4
HDelta = Int(mRes / 4)
For i = 0 To 4
   HLbl = Format(mBins(i * HDelta), HNumFmt)
   Res = GetTextExtentPoint32(hDC, HLbl, Len(HLbl), szLbl)
   Res = TextOut(hDC, i * LDelta - szLbl.cy \ 2, -200, HLbl, Len(HLbl))
Next i
End Sub


Private Sub WriteHighPeakInfo(ByVal hDC As Long)
Dim Lbl As String
Dim szLbl As Size
Dim DecPla As Long        'number of decimal places neccessary to
                          'see the difference between bin borders
Dim DecPlaFmt As String   'actual format to apply
Dim Res As Long
On Error Resume Next
DecPla = Abs(Int(Log(mBins(mFreqMaxInd) - mBins(mFreqMaxInd - 1)) / Log(10)))
If DecPla < 1 Then DecPla = 2
If DecPla > 6 Then DecPla = 6
DecPlaFmt = "0." & String(DecPla, "0")
Lbl = "High point: [" & Format(mBins(mFreqMaxInd - 1), DecPlaFmt) & ", " _
      & Format(mBins(mFreqMaxInd), DecPlaFmt) & "> - " & mFreqMax
Res = GetTextExtentPoint32(hDC, Lbl, Len(Lbl), szLbl)
Res = TextOut(hDC, 9200 - szLbl.cx, 10700, Lbl, Len(Lbl))
End Sub


Private Function CreateBins()
'----------------------------------------------------------
'create bins array, based on min-max value and type of bins
'----------------------------------------------------------
Dim i As Long
Dim HalfRes As Long
On Error GoTo err_CreateBins
ReDim mBins(mRes)
ReDim mFreq(mRes + 1)
Select Case DFBinsType
Case BinsUni, BinsLog
     mDelta = (mMax - mMin) / mRes
     For i = 0 To mRes
         mBins(i) = mMin + i * mDelta
     Next i
Case BinsRat
     mMin = 1 / mMax
     txtMin.Text = mMin
     HalfRes = Int(mRes / 2)
     If mRes Mod 2 > 0 Then         '1 is in the middle of bin
        mDelta = (mMax - 1) / mRes  'this is kind of half Delta
        For i = 0 To HalfRes
            mBins(HalfRes + i + 1) = 1 + (2 * i + 1) * mDelta
            mBins(HalfRes - i) = 1 / mBins(HalfRes + i + 1)
        Next i
        mDelta = 2 * mDelta
     Else                       '1 is border of bin
        mDelta = (mMax - 1) / HalfRes
        mBins(HalfRes) = 1
        For i = 1 To HalfRes
            mBins(HalfRes + i) = 1 + i * mDelta
            mBins(HalfRes - i) = 1 / mBins(HalfRes + i)
        Next i
     End If
End Select
CreateBins = True
Exit Function

err_CreateBins:
End Function

Private Function CalculateFrequencies()
Dim i As Long
Dim CurrData As Double
Dim CurrBin As Long
Dim HalfRes As Long
Dim HalfDelta As Double
On Error GoTo err_CalculateFrequencies
mFreqMax = 0
mFreqMaxInd = -1
Select Case DFBinsType
Case BinsUni
     For i = 0 To mDataCnt - 1
         If mData(i) < mMin Then
            CurrBin = 0
         ElseIf mData(i) > mMax Then
            CurrBin = mRes + 1
         Else
            CurrBin = Int((mData(i) - mMin) / mDelta)
         End If
         mFreq(CurrBin) = mFreq(CurrBin) + 1
         If mFreq(CurrBin) > mFreqMax Then          'need to keep track of max. frequency
            If CurrBin > 0 And CurrBin < mRes + 1 Then  'except out of range
               mFreqMax = mFreq(CurrBin)
               mFreqMaxInd = CurrBin
            End If
         End If
     Next i
Case BinsLog
     For i = 0 To mDataCnt - 1
         CurrData = Log(mData(i)) / Log(10#)
         If CurrData < mMin Then
            CurrBin = 0
         ElseIf CurrData >= mMax Then
            CurrBin = mRes + 1
         Else
            CurrBin = Int((CurrData - mMin) / mDelta)
         End If
         mFreq(CurrBin) = mFreq(CurrBin) + 1
         If mFreq(CurrBin) > mFreqMax Then          'need to keep track of max. frequency
            If CurrBin > 0 And CurrBin < mRes + 1 Then  'except out of range
               mFreqMax = mFreq(CurrBin)
               mFreqMaxInd = CurrBin
            End If
         End If
     Next i
Case BinsRat
     HalfRes = Int(mRes / 2)
     If mRes Mod 2 = 0 Then
        For i = 0 To mDataCnt - 1
            If mData(i) < mMin Then
                CurrBin = 0
            ElseIf mData(i) > mMax Then
                CurrBin = mRes + 1
            Else
                If mData(i) >= 1 Then
                   CurrBin = HalfRes + Int((mData(i) - 1) / mDelta) + 1
                Else
                   CurrBin = HalfRes - Int((1 / mData(i) - 1) / mDelta)
                End If
            End If
            mFreq(CurrBin) = mFreq(CurrBin) + 1
            If mFreq(CurrBin) > mFreqMax Then          'need to keep track of max. frequency
               If CurrBin > 0 And CurrBin < mRes + 1 Then  'except out of range
                  mFreqMax = mFreq(CurrBin)
                  mFreqMaxInd = CurrBin
               End If
            End If
         Next i
     Else
        HalfDelta = mDelta / 2
        For i = 0 To mDataCnt - 1
            If mData(i) < mMin Then
                CurrBin = 0
            ElseIf mData(i) > mMax Then
                CurrBin = mRes + 1
            Else
                If mData(i) >= 1 Then
                   If mData(i) - 1 <= HalfDelta Then
                      CurrBin = HalfRes + 1
                   Else
                      CurrBin = HalfRes + Int((mData(i) - (1 + HalfDelta)) / mDelta) + 2
                   End If
                Else
                   CurrData = 1 / mData(i)
                   If 1 / mData(i) - 1 <= HalfDelta Then
                      CurrBin = HalfRes + 1
                   Else
                      CurrBin = HalfRes - Int((CurrData - (1 + HalfDelta)) / mDelta)
                   End If
                End If
            End If
            mFreq(CurrBin) = mFreq(CurrBin) + 1
            If mFreq(CurrBin) > mFreqMax Then          'need to keep track of max. frequency
               If CurrBin > 0 And CurrBin < mRes + 1 Then  'except out of range
                  mFreqMax = mFreq(CurrBin)
                  mFreqMaxInd = CurrBin
               End If
            End If
         Next i
     End If
End Select
CalculateFrequencies = True
Exit Function

err_CalculateFrequencies:
End Function

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub


Public Sub DFRefresh()
'----------------------------------------
'rebuild bins and recalculate frequencies
'----------------------------------------
If CreateBins() Then
   If CalculateFrequencies() Then
      picD.Cls
      Call CalculateGraphPoints
      Call Draw
   End If
   UpdateStatus "Out of range count: " & mFreq(0) & "," & mFreq(mRes + 1)
   Exit Sub
End If
UpdateStatus "Error refreshing graph!"
End Sub


Public Function WriteDFToFile(ByVal FileName As String) As Boolean
'-----------------------------------------------------------------
'write bins results to text semicolon delimited file
'returns True if successful, False if not or any error occurs
'NOTE: bins are appended to a file; create file if not found
'-----------------------------------------------------------------
Dim i As Long
Dim hfile As Integer
On Error GoTo exit_WriteDFToFile

hfile = FreeFile
Open FileName For Append As hfile
Print #hfile, "Presenting - " & HLabel & ", " & VLabel
Print #hfile, "Range - " & mMin & ", " & mMax
Print #hfile, "Resolution - " & mRes
Select Case DFBinsType
Case BinsUni
    Print #hfile, "Bins type - uniform"
Case BinsLog
    Print #hfile, "Bins type - logarithmic"
Case BinsRat
    Print #hfile, "Bins type - ratio"
Case BinsInt
    Print #hfile, "Bins type - integer"
End Select
Print #hfile, vbCrLf
Print #hfile, "Bin Min;Bin Max;Frequency"
Print #hfile, "" & ";" & mMin & ";" & mFreq(0)
For i = 0 To mRes - 1
    Print #hfile, mBins(i) & ";" & mBins(i + 1) & ";" & mFreq(i + 1)
Next i
Print #hfile, mMax & ";" & "" & ";" & mFreq(mRes + 1)
Close hfile
WriteDFToFile = True

exit_WriteDFToFile:
End Function

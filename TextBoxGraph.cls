VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextBoxGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'implements properties and drawing of a text box for graph
'all coordinates(including font size properties) are specified in logical units
'created: 01/19/2003 nt
'last modified: 01/19/2003 nt
'------------------------------------------------------------------------------
Option Explicit

Dim mName As String              'name of text box to be used in lists and displays
Dim mTextLnCount As Long         'count of number of lines to draw
Dim mText() As String            'text to draw

Dim mX1L As Long                'logical coordinates of corner where text starts
Dim mY1L As Long

Dim mX2L As Long                'logical coordinates of corner where text ends
Dim mY2L As Long

Dim mIndentXL As Long
Dim mIndentYL As Long

Dim mForeColor As Long
Dim mBackColor As Long

Dim mShowBorder As Boolean

Dim mFontWidthL As Long
Dim mFontHeightL As Long
Dim mFontWeightL As Long
Dim mFontEscapement As Long
Dim mFontItalic As Boolean
Dim mFontUnderlined As Boolean      'not used for now
Dim mFontTypeFace As String         'this has to be zero terminated string(if used)

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Name(ByVal sNewValue As String)
mName = sNewValue
End Property

Public Property Get TextLnCount() As String         'read-only property
TextLnCount = mTextLnCount
End Property

Public Property Get lX1() As Long
lX1 = mX1L
End Property

Public Property Let lX1(ByVal lNewValue As Long)
mX1L = lNewValue
End Property

Public Property Get lY1() As Long
lY1 = mY1L
End Property

Public Property Let lY1(ByVal lNewValue As Long)
mY1L = lNewValue
End Property

Public Property Get lX2() As Long
lX2 = mX2L
End Property

Public Property Let lX2(ByVal lNewValue As Long)
mX2L = lNewValue
End Property

Public Property Get lY2() As Long
lY2 = mY2L
End Property

Public Property Let lY2(ByVal lNewValue As Long)
mY2L = lNewValue
End Property

Public Property Get lIndentX() As Long
lIndentX = mIndentXL
End Property

Public Property Let lIndentX(ByVal lNewValue As Long)
mIndentXL = lNewValue
End Property

Public Property Get lIndentY() As Long
lIndentY = mIndentYL
End Property

Public Property Let lIndentY(ByVal lNewValue As Long)
mIndentYL = lNewValue
End Property

Public Property Get BackColor() As Long
BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal lNewValue As Long)
mBackColor = lNewValue
End Property

Public Property Get ForeColor() As Long
ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal lNewValue As Long)
mForeColor = lNewValue
End Property

Public Property Get ShowBorder() As Boolean
ShowBorder = mShowBorder
End Property

Public Property Let ShowBorder(ByVal bNewValue As Boolean)
mShowBorder = bNewValue
End Property

Public Property Get FontWidthL() As Long
FontWidthL = mFontWidthL
End Property

Public Property Let FontWidthL(ByVal lNewValue As Long)
mFontWidthL = lNewValue
End Property

Public Property Get FontHeightL() As Long
FontHeightL = mFontHeightL
End Property

Public Property Let FontHeightL(ByVal lNewValue As Long)
mFontHeightL = lNewValue
End Property

Public Property Get FontWeightL() As Long
FontWeightL = mFontWeightL
End Property

Public Property Let FontWeightL(ByVal lNewValue As Long)
mFontWeightL = lNewValue
End Property

Public Property Get FontItalic() As Boolean
FontItalic = mFontItalic
End Property

Public Property Let FontItalic(ByVal bNewValue As Boolean)
mFontItalic = bNewValue
End Property

Public Sub AddTextLine(ByVal sNewLine As String)
mTextLnCount = mTextLnCount + 1
ReDim Preserve mText(mTextLnCount - 1)
mText(mTextLnCount - 1) = sNewLine
End Sub

Public Sub ClearText()
Erase mText
mTextLnCount = 0
End Sub

Public Function GetText() As String
Dim i As Long
Dim Tmp As String
On Error Resume Next
For i = 0 To mTextLnCount - 1
    Tmp = Tmp & mText(i) & vbCrLf
Next i
GetText = Left$(Tmp, Len(Tmp) - Len(vbCrLf))
End Function


Public Sub Draw(ByVal hDC As Long)
'-----------------------------------------------------------------
'draws text on specified device context(does not handle word wraps
'but ensures integrity of device context with clipping region
'-----------------------------------------------------------------
Dim i As Long, Res As Long
Dim lOldPen As Long, lNewPen As Long
Dim lOldBrush As Long, lNewBrush As Long
Dim lNewFont As Long, lOldFont As Long
Dim lfFnt As LOGFONT
'Dim hCR As Long               'clipping region so we don't draw over the border
Dim lTextPosX As Long
Dim lTextPosY As Long
Dim ptClipRect() As POINTAPI
On Error Resume Next

lNewPen = CreatePen(PS_SOLID, 1, mForeColor)
lNewBrush = CreateSolidBrush(mBackColor)
lOldPen = SelectObject(hDC, lNewPen)
lOldBrush = SelectObject(hDC, lNewBrush)
'retrieve structure for currently selected font
lOldFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lOldFont, Len(lfFnt), lfFnt)
Res = SelectObject(hDC, lOldFont)
'set font properties
lfFnt.lfHeight = mFontHeightL
lfFnt.lfWidth = mFontWidthL
lfFnt.lfItalic = mFontItalic
lfFnt.lfUnderline = mFontUnderlined
lNewFont = CreateFontIndirect(lfFnt)
Res = SetBkMode(hDC, TRANSPARENT)
Res = SetTextColor(hDC, mForeColor)
lOldFont = SelectObject(hDC, lNewFont)
'if required draw box around the text
If mShowBorder Then Res = Rectangle(hDC, mX1L, mY1L, mX2L, mY2L)
'set clipping region within the box
'ReDim ptClipRect(1)
'ptClipRect(0).x = mX1L + mIndentXL
'ptClipRect(0).y = mY1L + mIndentYL
'ptClipRect(1).x = mX2L - mIndentXL
'ptClipRect(1).y = mY2L - mIndentYL
'Res = LPtoDP(hDC, ptClipRect(0), 2)
'hCR = CreateRectRgn(ptClipRect(0).x, ptClipRect(0).y, ptClipRect(1).x, ptClipRect(1).y)
'Res = SelectClipRgn(hDC, hCR)
'now when everything is ready do the actual drawing
lTextPosX = mX1L + mIndentXL
lTextPosY = mY1L + mIndentYL
For i = 0 To mTextLnCount - 1
    Res = TextOut(hDC, lTextPosX, lTextPosY, mText(i), Len(mText(i)))
    lTextPosY = lTextPosX - mFontHeightL
Next i
Res = SelectObject(hDC, lOldPen)
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
'Res = SelectClipRgn(hDC, 0&)                'remove clipping region from the dc
DeleteObject (lNewFont)
DeleteObject (lNewPen)
DeleteObject (lNewBrush)
'DeleteObject (hCR)
End Sub


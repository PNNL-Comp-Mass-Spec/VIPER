Attribute VB_Name = "Module21"
'overlay drawing functions
'created: 09/07/2002 nt
'last modified: 09/07/2002 nt
'--------------------------------------------------------------------------
Option Explicit

Public hOlyBackClrBrush As Long
Public hOlyForeClrPen As Long
Public hOlyGridClrPen As Long

Public hOlyClrPen() As Long             'we want ready drawing tools
Public hOlyClrPenStick() As Long        'we want ready drawing tools
Public hOlyClrBrush() As Long

'--------------------------------------------------------------------------
'First block of functions relating with drawing objects and its maintenance
'--------------------------------------------------------------------------
Public Sub CreateOlyBackClrObject(ByVal BackColor As Long)
On Error Resume Next
If hOlyBackClrBrush <> 0 Then DeleteObject (hOlyBackClrBrush)
hOlyBackClrBrush = CreateSolidBrush(BackColor)
End Sub

Public Sub CreateOlyForeClrObject(ByVal ForeColor As Long)
On Error Resume Next
If hOlyForeClrPen <> 0 Then DeleteObject (hOlyForeClrPen)
hOlyForeClrPen = CreatePen(PS_SOLID, 1, ForeColor)

If hOlyGridClrPen <> 0 Then DeleteObject (hOlyGridClrPen)
hOlyGridClrPen = CreatePen(OlyOptions.GridLineStyle, 1, ForeColor)
End Sub

Public Sub DestroyOlyDrawingObjects()
On Error Resume Next
Dim i As Integer
If hOlyBackClrBrush <> 0 Then DeleteObject (hOlyBackClrBrush)
If hOlyForeClrPen <> 0 Then DeleteObject (hOlyForeClrPen)
If hOlyGridClrPen <> 0 Then DeleteObject (hOlyGridClrPen)
For i = 0 To OlyCnt - 1
    If hOlyClrBrush(i) <> 0 Then DeleteObject (hOlyClrBrush(i))
    If hOlyClrPen(i) <> 0 Then DeleteObject (hOlyClrPen(i))
    If hOlyClrPenStick(i) <> 0 Then DeleteObject (hOlyClrPenStick(i))
Next i
End Sub

Public Function AddEditOlyClr(ByVal OlyInd As Long, ByVal OlyColor As Long) As Boolean
'---------------------------------------------------------------------------------------
'if call to OlyClr object arrays fail we need to initialize; otherwise just change color
'---------------------------------------------------------------------------------------
On Error GoTo err_AddEditOlyClr
If hOlyClrBrush(OlyInd) <> 0 Then DeleteObject (hOlyClrBrush(OlyInd))
hOlyClrBrush(OlyInd) = CreateSolidBrush(OlyColor)
If hOlyClrPen(OlyInd) <> 0 Then DeleteObject (hOlyClrPen(OlyInd))
hOlyClrPen(OlyInd) = CreatePen(PS_SOLID, 1, OlyColor)
If hOlyClrPenStick(OlyInd) <> 0 Then DeleteObject (hOlyClrPenStick(OlyInd))
hOlyClrPenStick(OlyInd) = CreatePen(PS_SOLID, OlyOptions.NETStickWidth, OlyColor)
AddEditOlyClr = True
Exit Function

err_AddEditOlyClr:
If Err.Number = 9 Then
   ReDim Preserve hOlyClrPen(OlyInd)
   ReDim Preserve hOlyClrBrush(OlyInd)
   ReDim Preserve hOlyClrPenStick(OlyInd)
   Resume
End If
End Function

Public Function RemoveOlyClr(ByVal OlyInd As Long) As Boolean
Dim OlyClrCnt As Long
Dim i As Long
On Error Resume Next
OlyClrCnt = UBound(hOlyClrBrush) - 1
If hOlyClrBrush(OlyInd) <> 0 Then DeleteObject (hOlyClrBrush(OlyInd))
If hOlyClrPen(OlyInd) <> 0 Then DeleteObject (hOlyClrPen(OlyInd))
If hOlyClrPenStick(OlyInd) <> 0 Then DeleteObject (hOlyClrPenStick(OlyInd))
For i = OlyInd To OlyClrCnt - 1
    hOlyClrBrush(i) = hOlyClrBrush(i + 1)
    hOlyClrPen(i) = hOlyClrPen(i + 1)
    hOlyClrPenStick(i) = hOlyClrPenStick(i + 1)
Next i
ReDim Preserve hOlyClrPen(OlyClrCnt - 1)
ReDim Preserve hOlyClrPenStick(OlyClrCnt - 1)
ReDim Preserve hOlyClrBrush(OlyClrCnt - 1)
End Function
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------


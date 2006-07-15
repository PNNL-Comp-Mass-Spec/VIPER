VERSION 5.00
Begin VB.UserControl ctl2DHeatMap 
   BackColor       =   &H80FF0014&
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8385
   ScaleHeight     =   6045
   ScaleWidth      =   8385
   Begin VB.PictureBox pctSurface 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   240
      ScaleHeight     =   5415
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   120
      Width           =   7185
   End
   Begin VB.Timer tmrRefreshTimer 
      Interval        =   250
      Left            =   7800
      Top             =   120
   End
End
Attribute VB_Name = "ctl2DHeatMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const HORZSIZE As Long = 4 'Horizontal size in millimetres
Private Const VERTSIZE As Long = 6 'Vertical size in millimetres
Private Const HORZRES As Long = 8  'Horizontal width in pixels
Private Const VERTRES As Long = 10 'Vertical width in pixels

Private Const STRETCH_ANDSCANS As Long = 1
Private Const STRETCH_ORSCANS As Long = 2
Private Const STRETCH_DELETESCANS As Long = 3
Private Const STRETCH_HALFTONE As Long = 4



Private Declare Function CreateEnhMetaFile Lib "gdi32" _
   Alias "CreateEnhMetaFileA" _
  (ByVal hdcRef As Long, _
   ByVal lpFileName As String, _
   ByRef lpRect As Rect, _
   ByVal lpDescription As String) As Long

Private Declare Function CloseEnhMetaFile Lib "gdi32" _
  (ByVal hDC As Long) As Long

Private Declare Function DeleteEnhMetaFile Lib "gdi32" _
  (ByVal hEmf As Long) As Long

Private Declare Function PlayEnhMetaFile Lib "gdi32" _
   (ByVal hDC As Long, _
    ByVal hEmf As Long, _
    ByRef lpRect As Any) As Long

Private Declare Function BitBlt Lib "gdi32" _
   (ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hDC As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function GetClientRect Lib "user32" _
   (ByVal hwnd As Long, _
    ByRef lpRect As Rect) As Long

Private Declare Function GetWindowRect Lib "user32" _
  (ByVal hwnd As Long, _
   lpRect As Rect) As Long

Private Declare Function ReleaseDC Lib "user32" _
   (ByVal hwnd As Long, _
    ByVal hDC As Long) As Long

Private Declare Function SetStretchBltMode Lib "gdi32" _
   (ByVal hDC As Long, _
    ByVal nStretchMode As Long) As Long
    

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal xOffset As Long, ByVal yOffset As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Enum EDrawTextFormat
   DT_BOTTOM = &H8
   DT_CALCRECT = &H400
   DT_CENTER = &H1
   DT_EXPANDTABS = &H40
   DT_EXTERNALLEADING = &H200
   DT_INTERNAL = &H1000
   DT_LEFT = &H0
   DT_NOCLIP = &H100
   DT_NOPREFIX = &H800
   DT_RIGHT = &H2
   DT_SINGLELINE = &H20
   DT_TABSTOP = &H80
   DT_TOP = &H0
   DT_VCENTER = &H4
   DT_WORDBREAK = &H10
   DT_EDITCONTROL = &H2000&
   DT_PATH_ELLIPSIS = &H4000&
   DT_END_ELLIPSIS = &H8000&
   DT_MODIFYSTRING = &H10000
   DT_RTLREADING = &H20000
   DT_WORD_ELLIPSIS = &H40000
End Enum
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private m_font As IFont
Private m_tR As Rect


Const mint_num_colors = 64
Private DataZ() As Double
Private num_pts_x As Integer
Private num_pts_y As Integer
Private MinZVal As Double
Private MaxZVal As Double
Private Grade As Double
Private Colors() As Long
Private mLines() As Long
Private mZScoreMode As Boolean
Private num_pts As Long
Private marker_size As Integer
Private x_margin As Integer
Private y_margin As Integer
Private xlabel As String
Private ylabel As String
Private Color_Brushes() As Long
Private mlng_min_scan As Long
Private mlng_max_scan As Long
Private mdbl_min_net As Double
Private mdbl_max_net As Double
Private line_color As Long
Private net_line_color As Long
Private show_line As Boolean
Private NET_Slope As Double
Private NET_Intercept As Double
Private show_net_line As Boolean
Private axes_pen As Long
Private mNeedToRefresh As Boolean

Public Property Get ShowLine() As Boolean
    ShowLine = show_line
End Property
Public Property Let ShowLine(blnValue As Boolean)
    show_line = blnValue
End Property

Public Property Get ShowNetLine() As Boolean
    ShowNetLine = show_net_line
End Property
Public Property Let ShowNetLine(blnValue As Boolean)
    show_net_line = blnValue
End Property

Public Property Get ZScoreMode() As Boolean
    ZScoreMode = mZScoreMode
End Property
Public Property Let ZScoreMode(blnValue As Boolean)
    mZScoreMode = blnValue
End Property

Public Sub ClearData()
    ReDim DataZ(0, 0)
End Sub

Private Sub CreateColors()
    Dim i As Integer
    Dim R As Integer
    Dim g As Integer
    Dim B As Integer
    ReDim Colors(mint_num_colors, 2)
    ReDim Color_Brushes(mint_num_colors)
    For i = 0 To mint_num_colors
        If (i < 24) Then
            R = (255 * i * 1# / 24)
            g = 0
            B = 0
        ElseIf (i < 48) Then
            R = 255
            g = (255 * (i - 24) * 1# / 24)
            B = 0
        Else
            R = 255
            g = 255
            B = (255 * (i - 48) * 1# / 16)
        End If
        Colors(i, 0) = R
        Colors(i, 1) = g
        Colors(i, 2) = B
        Color_Brushes(i) = CreateSolidBrush(RGB(R, g, B))
        line_color = RGB(40, 40, 255)
        net_line_color = RGB(0, 200, 0)
        axes_pen = CreatePen(0, 2, RGB(0, 0, 0))
    Next
End Sub

Private Function DcToEmf2(ByVal hDcIn As Long, _
                          inArea As Rect, _
                          sOutputFile As String) As Long
    
   Dim rc As Rect
   Dim MetaDC As Long
   Dim OldMode As Long
   Dim hsize As Long
   Dim vsize As Long
   Dim hres As Long
   Dim vres As Long

  'Convert the area from pixels to .01mm's
  'Rectangle coordinates must be normalised
   hsize = GetDeviceCaps(hDcIn, HORZSIZE) * 100
   vsize = GetDeviceCaps(hDcIn, VERTSIZE) * 100
   hres = GetDeviceCaps(hDcIn, HORZRES)
   vres = GetDeviceCaps(hDcIn, VERTRES)
   
   With rc
      .Left = (inArea.Left * hsize) / hres
      .Top = (inArea.Top * vsize) / vres
      .Right = (inArea.Right * hsize) / hres
      .Bottom = (inArea.Bottom * vsize) / vres
   End With
    
  'Create a new MetaDC and output file
   MetaDC = CreateEnhMetaFile(hDcIn, sOutputFile, rc, vbNullString)
        
   If (MetaDC) Then
        
     'Draw the image to the MetaDC
     'Set STRETCH_HALFTONE stretch mode here for higher quality
      OldMode = SetStretchBltMode(MetaDC, STRETCH_HALFTONE)
        
      Call BitBlt(MetaDC, _
                  0, 0, _
                 (inArea.Right - inArea.Left), _
                 (inArea.Bottom - inArea.Top), _
                  hDcIn, _
                  inArea.Left, _
                  inArea.Top, _
                  vbSrcCopy)
            
     'restore the saved dc mode
      Call SetStretchBltMode(MetaDC, OldMode)

     'delete the MetaDC and return the
     'EMF object's handle
      DcToEmf2 = CloseEnhMetaFile(MetaDC)
        
   End If
   
End Function

Public Sub Draw2EMF(file_name As String)
     Dim hEmf As Long
     Dim rc As Rect
    
    'Obtain a handle to a Windows
    'enhanced metafile of the desktop
    '(or to the client area of another
    'form or window specified by hwnd),
    'and optionally display the result
    'in a picturebox using metafile APIs,
    'then clean up
  
On Error GoTo Draw2EMFErrorHandler
   
    hEmf = WindowClientToEMF(file_name)
    Call DeleteEnhMetaFile(hEmf)

    Exit Sub

Draw2EMFErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "ctl2DHeatMap->Draw2EMF"

End Sub

Public Sub Draw2EMF2Clipboard()
   Dim hEmf As Long
   Dim rc As Rect
   Dim Res As Long
   Dim hTmpDc As Long

  'obtain the display context (DC)
  'to the window passed

On Error GoTo Draw2EMF2ClipboardErrorHandler

   hTmpDc = pctSurface.hDC
   
   If hTmpDc <> 0 Then
   
     'get the size of the client
     'area of the passed handle
     rc.Left = 0
     rc.Top = 0
     rc.Bottom = pctSurface.Height / Screen.TwipsPerPixelX
     rc.Right = pctSurface.width / Screen.TwipsPerPixelY
 '     If GetClientRect(UserControl.hwnd, rc) <> 0 Then
       
        'pass the DC, rectangle and filename
        'to create the file, returning the
        'handle to the memory metafile
        
        hEmf = DcToEmf2(hTmpDc, rc, vbNullString)
         
        Res = OpenClipboard(UserControl.hwnd)
        Res = EmptyClipboard()
        Res = SetClipboardData(CF_ENHMETAFILE, hEmf)
        Res = CloseClipboard
        'release the temporary DC
        Call ReleaseDC(UserControl.hwnd, hTmpDc)
        Call DeleteEnhMetaFile(hEmf)
'      End If
   End If
    
    mNeedToRefresh = True
    
    Exit Sub
    
Draw2EMF2ClipboardErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "ctl2DHeatMap->Draw2EMF2Clipboard"

End Sub

Private Sub DrawNetLine(hDC_pic As Long)
    Dim y_coord As Long, y_val As Double
    Dim y_coord_next As Long, y_val_next As Double
    Dim x_coord As Long, x_val As Long
    Dim x_coord_next As Long, x_val_next As Long
    Dim Wide As Long, High As Long
    Dim net_pen As Long
    Dim lOldBrush As Long
    Dim tJunk As POINTAPI

    If NET_Slope = 0 Then Exit Sub
    
On Error GoTo DrawNetLineErrorHandler

    Wide = pctSurface.width - y_margin * 2
    High = pctSurface.Height - x_margin * 2
        
    y_val = mlng_min_scan * NET_Slope + NET_Intercept
    x_val = mlng_min_scan
    If y_val < mdbl_min_net Then
        y_val = mdbl_min_net
        x_val = (y_val - NET_Intercept) / NET_Slope
    End If
    
    y_val_next = mlng_max_scan * NET_Slope + NET_Intercept
    x_val_next = mlng_max_scan
    If y_val_next > mdbl_max_net Then
        y_val_next = mdbl_max_net
        x_val_next = (y_val_next - NET_Intercept) / NET_Slope
    End If
    
    x_coord = (Wide * CDbl(x_val - mlng_min_scan)) / (mlng_max_scan - mlng_min_scan) + y_margin
    x_coord_next = (Wide * CDbl(x_val_next - mlng_min_scan)) / (mlng_max_scan - mlng_min_scan) + y_margin
    
    y_coord = (High * CDbl(mdbl_max_net - y_val)) / (mdbl_max_net - mdbl_min_net) + x_margin
    y_coord_next = (High * CDbl(mdbl_max_net - y_val_next)) / (mdbl_max_net - mdbl_min_net) + x_margin
    
    'UserControl.Line (x_coord, y_coord)-Step(x_coord_next - x_coord, y_coord_next - y_coord), net_line_color

    net_pen = CreatePen(0, 2, net_line_color)
    lOldBrush = SelectObject(hDC_pic, net_pen)
    
    MoveToEx hDC_pic, x_coord \ Screen.TwipsPerPixelX, y_coord \ Screen.TwipsPerPixelY, tJunk
    LineTo hDC_pic, x_coord_next \ Screen.TwipsPerPixelX, y_coord_next \ Screen.TwipsPerPixelY

    Exit Sub
    
DrawNetLineErrorHandler:
    Debug.Assert False
End Sub

Private Sub DrawLine(hDC_pict As Long)
    Dim j As Long
    Dim x_val As Long, y_val As Long
    Dim x_val_next As Long, y_val_next As Long
    Dim x_step As Long, y_step As Long
    Dim Wide As Long, High As Long
    Dim transform_pen As Long
    Dim lOldBrush As Long
    Dim x_from As Long, y_from As Long
    Dim x_to As Long, y_to As Long
    Dim tJunk As POINTAPI
    
    If num_pts_x = 0 Or num_pts_y = 0 Then Exit Sub
    Wide = pctSurface.width - y_margin * 2
    High = pctSurface.Height - x_margin * 2
    x_step = Wide / (num_pts_x + 1)
    y_step = High / (num_pts_y + 1)

    
    transform_pen = CreatePen(0, 2, line_color)
    lOldBrush = SelectObject(hDC_pict, transform_pen)
    
    For j = 0 To num_pts - 1
        x_val = mLines(j, 0)
        y_val = mLines(j, 1)
        x_val_next = mLines(j + 1, 0)
        y_val_next = mLines(j + 1, 1)
        x_from = ((Wide * CDbl(x_val)) / (num_pts_x + 1) + y_margin) \ Screen.TwipsPerPixelX
        x_to = x_from + (x_step * (x_val_next - x_val)) \ Screen.TwipsPerPixelX
        
        y_from = ((High * CDbl(num_pts_y - y_val - 1)) / (num_pts_y + 1) + x_margin) \ Screen.TwipsPerPixelY
        y_to = y_from - ((y_val_next - y_val) * y_step) \ Screen.TwipsPerPixelY
        
        'UserControl.Line ((Wide * x_val) / (num_pts_x + 1) + y_margin, (High * (num_pts_y - y_val - 1)) / (num_pts_y + 1) + x_margin)-Step(x_step * (x_val_next - x_val), -1 * (y_val_next - y_val) * y_step), line_color
    
        MoveToEx hDC_pict, x_from, y_from, tJunk
        LineTo hDC_pict, x_to, y_to
    Next

End Sub

Private Sub DrawAxes(hDC_pic As Long)
    Dim Wide As Long, High As Long
    Dim tJunk As POINTAPI
    Dim lOldBrush As Long
    Wide = pctSurface.width - y_margin * 2
    High = pctSurface.Height - x_margin * 2
    
    lOldBrush = SelectObject(hDC_pic, axes_pen)
    MoveToEx hDC_pic, y_margin \ Screen.TwipsPerPixelX, (x_margin + High) \ Screen.TwipsPerPixelY, tJunk
    LineTo hDC_pic, (y_margin + Wide + y_margin / 2) \ Screen.TwipsPerPixelX, (x_margin + High) \ Screen.TwipsPerPixelY
    
    MoveToEx hDC_pic, y_margin \ Screen.TwipsPerPixelX, (x_margin + High) \ Screen.TwipsPerPixelY, tJunk
    LineTo hDC_pic, y_margin \ Screen.TwipsPerPixelX, (x_margin / 2) \ Screen.TwipsPerPixelY

End Sub

Private Sub DrawTicks(hDC_pic As Long)
    Dim num_divisions As Integer
    Dim i As Integer
    Dim x_val As Long, y_val As Double
    Dim x_coordinate As Long
    Dim y_coordinate As Long
    Dim retVal As Long
    Dim tR As Rect
    Dim Wide As Long
    Dim High As Long
    Dim x_step As Long, y_step As Double
    Dim tJunk As POINTAPI
    Dim tick_label_margin As Long
    Dim num_y_divisions As Integer
    Dim num_x_divisions As Integer

On Error GoTo DrawTicksErrorHandler
    y_step = 0.1
    tick_label_margin = 100
    num_divisions = 5
    If mdbl_min_net = mdbl_max_net Then Exit Sub
    
    Wide = pctSurface.width - y_margin * 2
    High = pctSurface.Height - x_margin * 2
    
    x_step = CLng(Log10(mlng_max_scan - mlng_min_scan) - 0.5)
    x_step = 10 ^ CLng(x_step)
    num_divisions = (mlng_max_scan - mlng_min_scan) / x_step
    If num_divisions < 3 Then
        num_divisions = 5 * num_divisions
        x_step = x_step / 5
    ElseIf num_divisions < 5 Then
        num_divisions = 2 * num_divisions
        x_step = x_step / 2
    End If
        
    x_val = CLng(mlng_min_scan / x_step) * x_step
    If x_val < mlng_min_scan Then x_val = x_val + x_step
    While x_val < mlng_max_scan
        x_coordinate = (Wide * CDbl(x_val - mlng_min_scan)) / (mlng_max_scan - mlng_min_scan) + y_margin
        y_coordinate = x_margin + High
        
        tR.Left = (x_coordinate - UserControl.TextWidth(CStr(x_val)) / 2) \ Screen.TwipsPerPixelX
        tR.Top = y_coordinate \ Screen.TwipsPerPixelY + tick_label_margin \ Screen.TwipsPerPixelY
        tR.Right = (x_coordinate + UserControl.TextWidth(CStr(x_val)) / 2) \ Screen.TwipsPerPixelX
        tR.Bottom = (y_coordinate + UserControl.TextHeight(CStr(x_val))) \ Screen.TwipsPerPixelY + tick_label_margin \ Screen.TwipsPerPixelY
        
        DrawText hDC_pic, CStr(x_val), -1, tR, DT_LEFT Or DT_SINGLELINE
        MoveToEx hDC_pic, x_coordinate \ Screen.TwipsPerPixelX, y_coordinate \ Screen.TwipsPerPixelY, tJunk
        LineTo hDC_pic, x_coordinate \ Screen.TwipsPerPixelX, y_coordinate \ Screen.TwipsPerPixelY + tick_label_margin \ Screen.TwipsPerPixelY
        
        x_val = x_val + x_step
    Wend
    
    
    tR.Left = (Wide / 2 + y_margin - UserControl.TextWidth(xlabel) / 2) \ Screen.TwipsPerPixelX
    tR.Top = (High + x_margin + x_margin / 2) \ Screen.TwipsPerPixelY
    tR.Right = (Wide / 2 + y_margin + UserControl.TextWidth(xlabel) / 2) \ Screen.TwipsPerPixelX
    tR.Bottom = (High + 2 * x_margin) \ Screen.TwipsPerPixelY
    DrawText hDC_pic, xlabel, -1, tR, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    
    y_val = CInt(mdbl_min_net / y_step) * y_step
    If y_val < mdbl_min_net Then y_val = y_val + y_step
    
    While y_val < mdbl_max_net
        y_coordinate = (High * CDbl(mdbl_max_net - y_val)) / (mdbl_max_net - mdbl_min_net) + x_margin
        x_coordinate = y_margin
        tR.Left = (x_coordinate - UserControl.TextWidth(Format(CStr(y_val), "0.00"))) \ Screen.TwipsPerPixelX - tick_label_margin \ Screen.TwipsPerPixelY
        tR.Top = (y_coordinate - UserControl.TextHeight(Format(CStr(y_val), "0.00")) / 2) \ Screen.TwipsPerPixelY
        tR.Right = (x_coordinate) \ Screen.TwipsPerPixelX - tick_label_margin \ Screen.TwipsPerPixelY
        tR.Bottom = (y_coordinate + UserControl.TextHeight(Format(CStr(y_val), "0.00")) / 2) \ Screen.TwipsPerPixelY
        
        DrawText hDC_pic, Format(CStr(y_val), "0.00"), -1, tR, DT_LEFT Or DT_SINGLELINE
        MoveToEx hDC_pic, x_coordinate \ Screen.TwipsPerPixelX, y_coordinate \ Screen.TwipsPerPixelY, tJunk
        LineTo hDC_pic, x_coordinate \ Screen.TwipsPerPixelX - tick_label_margin \ Screen.TwipsPerPixelX, y_coordinate \ Screen.TwipsPerPixelY
        y_val = y_val + y_step
    Wend
    
    tR.Left = 0
    tR.Top = (High / 2 + x_margin) \ Screen.TwipsPerPixelY
    tR.Right = (UserControl.TextWidth(ylabel)) \ Screen.TwipsPerPixelX
    tR.Bottom = (High / 2 + x_margin + TextWidth(ylabel)) \ Screen.TwipsPerPixelY
    DrawText hDC_pic, ylabel, -1, tR, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    
    Exit Sub

DrawTicksErrorHandler:
    Debug.Assert False
    
End Sub

Private Sub DrawSurface(hDC_pic As Long)
    Dim Wide As Long
    Dim High As Long
    Dim x_step As Long, y_step As Long
    Dim x_val As Long, y_val As Long
    Dim x_val_next As Long, y_val_next As Long
    Dim clr_index As Integer
    Dim lNewBrush As Long
    Dim lOldBrush As Long
    Dim retVal As Long
    If num_pts_x = 0 Then Exit Sub
    Dim tR As Rect
    Dim i As Long
    Dim j As Long
    Dim hBr As Long
    
    Wide = pctSurface.width - y_margin * 2
    High = pctSurface.Height - x_margin * 2
    x_step = Wide / (num_pts_x + 1)
    y_step = High / (num_pts_y + 1)
    
    num_pts_x = UBound(DataZ, 1)
    num_pts_y = UBound(DataZ, 2)
    If num_pts_x = 0 Or num_pts_y = 0 Then Exit Sub
    x_val = y_margin
    x_val_next = y_margin + x_step
    For i = 0 To num_pts_x
        y_val = x_margin + High
        y_val_next = y_val - y_step
        For j = 0 To num_pts_y
            clr_index = GetColorIndex(DataZ(i, j))
            If clr_index < 0 Then clr_index = 0
            lNewBrush = Color_Brushes(clr_index)
            lOldBrush = SelectObject(hDC_pic, lNewBrush)
            
            tR.Left = x_val \ Screen.TwipsPerPixelX
            tR.Top = y_val_next \ Screen.TwipsPerPixelY
            tR.Right = x_val_next \ Screen.TwipsPerPixelX
            tR.Bottom = y_val \ Screen.TwipsPerPixelY
            FillRect hDC_pic, tR, lNewBrush
            
            y_val = y_val_next
            y_val_next = y_val_next - y_step
        Next
        x_val = x_val_next
        x_val_next = x_val_next + x_step
    Next

End Sub

Private Function GetColorIndex(val As Double) As Long
    Dim clr_index As Long
    If val <= MinZVal Then
        GetColorIndex = 0
        Exit Function
    End If
    clr_index = mint_num_colors * (val - MinZVal) / Grade
    If clr_index < 0 Then
        clr_index = 0
    ElseIf clr_index >= mint_num_colors Then
        clr_index = mint_num_colors - 1
    End If
    GetColorIndex = clr_index
    
End Function

Public Sub Refresh()
    RefreshPlotNow
End Sub

Public Sub Plot3DSimpleSurface(scores As Variant)
    Dim sum_val As Double
    Dim diff_val As Double
    Dim sum_square_val As Double
    Dim std_dev As Double
    Dim mean_val As Double
    Dim i As Long
    Dim j As Long
    
    
On Error GoTo Plot3DSimpleSurfaceErrorHandler
       
    num_pts_x = UBound(scores, 1)
    num_pts_y = UBound(scores, 2)
    ReDim DataZ(num_pts_x, num_pts_y)
    MinZVal = scores(0, 0)
    MaxZVal = scores(0, 0)
    For i = 0 To num_pts_x
        For j = 0 To num_pts_y
            DataZ(i, j) = scores(i, j)
            If DataZ(i, j) > MaxZVal Then MaxZVal = DataZ(i, j)
            If DataZ(i, j) < MinZVal Then MinZVal = DataZ(i, j)
        Next
    Next
    
    If mZScoreMode Then
        MinZVal = 0
        MaxZVal = 0
        For i = 0 To num_pts_x
            sum_val = 0
            sum_square_val = 0
            For j = 0 To num_pts_y
                sum_val = sum_val + scores(i, j)
            Next
            mean_val = sum_val / (num_pts_y + 1)
            For j = 0 To num_pts_y
                diff_val = scores(i, j) - mean_val
                sum_square_val = sum_square_val + diff_val * diff_val
            Next
            
            std_dev = Sqr(sum_square_val / num_pts_y)
            If std_dev = 0 Then
                For j = 0 To num_pts_y
                    DataZ(i, j) = 0
                Next
            Else
                For j = 0 To num_pts_y
                    DataZ(i, j) = (DataZ(i, j) - mean_val) / std_dev
                    If DataZ(i, j) > MaxZVal Then MaxZVal = DataZ(i, j)
                    If DataZ(i, j) < MinZVal Then MinZVal = DataZ(i, j)
                Next
            End If
            
        Next
    End If
    
    If MaxZVal = MinZVal Then
        Grade = 0
    Else
        Grade = MaxZVal - MinZVal
    End If
    Exit Sub
    
Plot3DSimpleSurfaceErrorHandler:
    Debug.Assert False
    
End Sub

Public Sub PlotLine(LineCoordinates As Variant)
    Dim j As Long
    
On Error GoTo Plot3DLineErrorHandler
     
    num_pts = UBound(LineCoordinates, 1)
    ReDim mLines(num_pts, 1)
    For j = 0 To num_pts
        mLines(j, 0) = LineCoordinates(j, 0)
        mLines(j, 1) = LineCoordinates(j, 1)
    Next
    Exit Sub

Plot3DLineErrorHandler:
    Debug.Assert False

End Sub

Public Sub RefreshPlotNow()
    Dim hDC_temp As Long

On Error GoTo RefreshPlotNowErrorHandler

    mNeedToRefresh = False

    pctSurface.Cls
    hDC_temp = pctSurface.hDC

    If num_pts_x = 0 Or num_pts_y = 0 Then Exit Sub
    DrawSurface (hDC_temp)
    
    ' Draw the NET line first
    If show_net_line Then DrawNetLine (hDC_temp)
    
    ' Now draw the transform line on top of the NET line
    If show_line Then DrawLine (hDC_temp)
    DrawAxes (hDC_temp)
    DrawTicks (hDC_temp)

  
    Exit Sub

RefreshPlotNowErrorHandler:
    Debug.Assert False
    'LogErrors Err.Number, "ctl2DHeatMap->Refresh"
    
End Sub

Private Sub ResizeControl()

On Error GoTo ResizeControlErrorHandler

    With pctSurface
        .Left = 0
        .Top = 0
        .width = UserControl.width
        .Height = UserControl.Height

        x_margin = .Height / 10
        y_margin = .width / 10
    End With

    Exit Sub
    
ResizeControlErrorHandler:
    Debug.Assert False

End Sub

Public Sub SetBounds(min_scan As Long, max_scan As Long, min_net As Double, max_net As Double)
    mlng_min_scan = min_scan
    mlng_max_scan = max_scan
    mdbl_min_net = min_net
    mdbl_max_net = max_net
End Sub

Public Sub SetNetSlopeAndIntercept(slope As Double, intercept As Double)
    NET_Slope = slope
    NET_Intercept = intercept
End Sub

Private Function WindowClientToEMF(sOutputFile As String) As Long
    
   Dim rc As Rect

   If pctSurface.hDC <> 0 Then
   
     'get the size of the client
     'area of the passed handle
      If GetClientRect(pctSurface.hwnd, rc) <> 0 Then
       
        'pass the DC, rectangle and filename
        'to create the file, returning the
        'handle to the memory metafile
         WindowClientToEMF = DcToEmf2(pctSurface.hDC, rc, sOutputFile)
         
        'release the temporary DC
         Call ReleaseDC(pctSurface.hwnd, pctSurface.hDC)
      
      End If
   End If
    
End Function

Private Sub tmrRefreshTimer_Timer()
    If mNeedToRefresh Then
        RefreshPlotNow
    End If
End Sub

Private Sub UserControl_Initialize()
    num_pts_x = 0
    num_pts_y = 0
    MinZVal = 0
    MaxZVal = 0
    num_pts = 0
    marker_size = 10
    mZScoreMode = True
    Call CreateColors
    x_margin = 300
    y_margin = 300
    xlabel = "MS Scan"
    ylabel = "NET"
    mlng_min_scan = 0
    mlng_max_scan = 0
    mdbl_min_net = 0
    mdbl_max_net = 0
    show_net_line = True

    ResizeControl
End Sub

Private Sub UserControl_Paint()
    ' Send request to paint plot
    mNeedToRefresh = True
End Sub

Public Sub UserControl_Resize()
    ResizeControl
End Sub

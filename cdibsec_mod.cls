VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type


Private Enum ERGBCompression
  BI_RGB = 0&
  BI_RLE4 = 2&
  BI_RLE8 = 1&
  DIB_RGB_COLORS = 0 '  color table in RGBs
End Enum


Private Type BITMAPINFOHEADER '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As ERGBCompression
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type


Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type


Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type


' From winuser.h
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ENHMETAFILE = 3

Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000


' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBMP As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As Long, ByRef lpbi As BITMAPINFO, ByVal uUsage As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInstance As Long, ByVal Name As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long


' Handle to the current DIBSection:
Private m_hDib As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_bmi As BITMAPINFO


Public Function CreateDIB( _
  ByVal lhDC As Long, _
  ByVal lWidth As Long, _
  ByVal lHeight As Long, _
  ByVal lChannels As Long, _
  ByRef hDib As Long _
  ) As Boolean
   
  With m_bmi.bmiHeader
    .biSize = Len(m_bmi.bmiHeader)
    .biWidth = lWidth
    .biHeight = lHeight
    .biPlanes = 1
    If lChannels = 3 Then
      .biBitCount = 24
    Else
      .biBitCount = 32
    End If
    .biCompression = BI_RGB
    .biSizeImage = BytesPerScanLine * .biHeight
  End With
  
  'The m_lPtr is passed in byref.. so that it returns the the pointer to the bitmapinfo bits
  'the m_lptr is then stored as a reference to the uncompressed image data
  'the m_lptr is filled with image data when the ijlread method is invoked.
  hDib = CreateDIBSection(lhDC, m_bmi, DIB_RGB_COLORS, m_lPtr, 0, 0)
  
  CreateDIB = (hDib <> 0)

End Function


Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lChannels As Long) As Boolean
  
  CleanUp
  
  m_hDC = CreateCompatibleDC(0)
  
  If (m_hDC <> 0) Then
    If (CreateDIB(m_hDC, lWidth, lHeight, lChannels, m_hDib)) Then
      m_hBmpOld = SelectObject(m_hDC, m_hDib)
      Create = True
    Else
      Call DeleteObject(m_hDC)
      m_hDC = 0
    End If
  End If

End Function


Public Function Load(ByVal Name As String) As Boolean
  Dim hBMP As Long
  Dim pName As Long
  Dim aName As String

  Load = False

  CleanUp

  m_hDC = CreateCompatibleDC(0)
  If m_hDC = 0 Then
    Exit Function
  End If

  aName = StrConv(Name, vbFromUnicode)
  pName = StrPtr(aName)

  hBMP = LoadImage(0, pName, IMAGE_BITMAP, 0, 0, (LR_CREATEDIBSECTION Or LR_LOADFROMFILE))
  If hBMP = 0 Then
    Call DeleteObject(m_hDC)
    m_hDC = 0
    MsgBox "Can't load BMP image"
    Exit Function
  End If

  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

  ' get image sizes
  Call GetDIBits(m_hDC, hBMP, 0, 0, 0, m_bmi, DIB_RGB_COLORS)

  ' make 24 bpp dib section
  m_bmi.bmiHeader.biBitCount = 24
  m_bmi.bmiHeader.biCompression = BI_RGB
  m_bmi.bmiHeader.biClrUsed = 0
  m_bmi.bmiHeader.biClrImportant = 0
  
  m_hDib = CreateDIBSection(m_hDC, m_bmi, DIB_RGB_COLORS, m_lPtr, 0, 0)
  If m_hDib = 0 Then
    Call DeleteObject(hBMP)
    Call DeleteObject(m_hDC)
    m_hDC = 0
    Exit Function
  End If

  m_hBmpOld = SelectObject(m_hDC, m_hDib)

  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

  ' get image data in 24 bpp format (convert if need)
  Call GetDIBits(m_hDC, hBMP, 0, m_bmi.bmiHeader.biHeight, m_lPtr, m_bmi, DIB_RGB_COLORS)

  Call DeleteObject(hBMP)

  Load = True

End Function


Public Property Get BytesPerScanLine() As Long
  ' Scans must align on dword boundaries:
  BytesPerScanLine = (m_bmi.bmiHeader.biWidth * (m_bmi.bmiHeader.biBitCount / 8) + 3) And &HFFFFFFFC
End Property


Public Property Get dib_width() As Long
  dib_width = m_bmi.bmiHeader.biWidth
End Property


Public Property Get dib_height() As Long
  dib_height = m_bmi.bmiHeader.biHeight
End Property


Public Property Get dib_channels() As Long
  dib_channels = m_bmi.bmiHeader.biBitCount / 8
End Property


Public Sub PaintPicture( _
  ByVal lhDC As Long, _
  Optional ByVal lDestLeft As Long = 0, _
  Optional ByVal lDestTop As Long = 0, _
  Optional ByVal lDestWidth As Long = -1, _
  Optional ByVal lDestHeight As Long = -1, _
  Optional ByVal lSrcLeft As Long = 0, _
  Optional ByVal lSrcTop As Long = 0, _
  Optional ByVal eRop As RasterOpConstants = vbSrcCopy)

  If (lDestWidth < 0) Then lDestWidth = m_bmi.bmiHeader.biWidth
  If (lDestHeight < 0) Then lDestHeight = m_bmi.bmiHeader.biHeight

  Call BitBlt(lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop)

End Sub


Public Property Get hDC() As Long
  hDC = m_hDC
End Property


Public Property Get hDib() As Long
  hDib = m_hDib
End Property


Public Property Get DIBSectionBitsPtr() As Long
  DIBSectionBitsPtr = m_lPtr
End Property


Public Sub CleanUp()
  
  If (m_hDC <> 0) Then
    If (m_hDib <> 0) Then
      Call SelectObject(m_hDC, m_hBmpOld)
      Call DeleteObject(m_hDib)
    End If
    Call DeleteObject(m_hDC)
  End If
  
  m_hDC = 0
  m_hDib = 0
  m_hBmpOld = 0
  m_lPtr = 0

  m_bmi.bmiColors.rgbBlue = 0
  m_bmi.bmiColors.rgbGreen = 0
  m_bmi.bmiColors.rgbRed = 0
  m_bmi.bmiColors.rgbReserved = 0
  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)
  m_bmi.bmiHeader.biWidth = 0
  m_bmi.bmiHeader.biHeight = 0
  m_bmi.bmiHeader.biPlanes = 0
  m_bmi.bmiHeader.biBitCount = 0
  m_bmi.bmiHeader.biClrUsed = 0
  m_bmi.bmiHeader.biClrImportant = 0
  m_bmi.bmiHeader.biCompression = 0

End Sub


Private Sub Class_Terminate()
  CleanUp
End Sub


Public Function LoadFromBMP(ByVal hBMP As Long) As Boolean
  Dim Res As Long
  'On Error Resume Next
  CleanUp
  m_hDC = CreateCompatibleDC(0)
  If m_hDC = 0 Then Exit Function
  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

  ' get image sizes
  Call GetDIBits(m_hDC, hBMP, 0, 0, 0, m_bmi, DIB_RGB_COLORS)

  ' make 24 bpp dib section
  m_bmi.bmiHeader.biBitCount = 24
  m_bmi.bmiHeader.biCompression = BI_RGB
  m_bmi.bmiHeader.biClrUsed = 0
  m_bmi.bmiHeader.biClrImportant = 0
  
  m_hDib = CreateDIBSection(m_hDC, m_bmi, DIB_RGB_COLORS, m_lPtr, 0, 0)
  If m_hDib = 0 Then
    Call DeleteObject(hBMP)
    Call DeleteObject(m_hDC)
    m_hDC = 0
    Exit Function
  End If

  m_hBmpOld = SelectObject(m_hDC, m_hDib)
  
  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

  ' get image data in 24 bpp format (convert if need)
  Call GetDIBits(m_hDC, hBMP, 0, m_bmi.bmiHeader.biHeight, m_lPtr, m_bmi, DIB_RGB_COLORS)

  Call DeleteObject(hBMP)
  LoadFromBMP = True
End Function


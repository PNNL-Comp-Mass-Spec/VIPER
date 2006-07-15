VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LaAutoGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'implements visibility, style and calculation modes of 2D grid
'created: 12/12/2002 nt
'last modified: 12/21/2002 nt
'-------------------------------------------------------------
Option Explicit

Public Enum GridLineStyle
    glsSOLID = 0
    glsDASH = 1
    glsDOT = 2
    glsDASHDOT = 3
    glsDASHDOTDOT = 4
End Enum

Public Enum GridAutoMode
    gamWidthConst = 0           'constant width of grid
    gamBinsCntConst = 1         'constant number of bins
End Enum

Public HorzGridVisible As Boolean

Public VertGridVisible As Boolean

Public LineStyle As GridLineStyle

Dim mHorzAutoMode As GridAutoMode

Dim mVertAutoMode As GridAutoMode

'coordinates of the rectangle that has to be gridlined
Dim mXMin As Single
Dim mXMax As Single
Dim mYMin As Single
Dim mYMax As Single

Dim mHorzWidth As Single
Dim mHorzBinsCnt As Long

Dim mVertWidth As Single
Dim mVertBinsCnt As Long


Public Sub UpdateGrid(NewXMin As Single, NewXMax As Single, _
                      NewYMin As Single, NewYMax As Single)
On Error Resume Next
mXMin = NewXMin:        mXMax = NewXMax
mYMin = NewYMin:        mYMax = NewYMax
Call CalculateGridParameters
End Sub

Public Property Get HorzAutoMode() As GridAutoMode
HorzAutoMode = mHorzAutoMode
End Property

Public Property Let HorzAutoMode(ByVal NewMode As GridAutoMode)
mHorzAutoMode = NewMode
Call CalculateGridParameters
End Property

Public Property Get VertAutoMode() As GridAutoMode
VertAutoMode = mVertAutoMode
End Property

Public Property Let VertAutoMode(ByVal NewMode As GridAutoMode)
mVertAutoMode = NewMode
Call CalculateGridParameters
End Property

Public Property Get HorzWidth() As Single
HorzWidth = mHorzWidth
End Property

Public Property Let HorzWidth(ByVal NewWidth As Single)
mHorzWidth = NewWidth
Call CalculateGridParameters
End Property

Public Property Get VertWidth() As Single
VertWidth = mVertWidth
End Property

Public Property Let VertWidth(ByVal NewWidth As Single)
mVertWidth = NewWidth
Call CalculateGridParameters
End Property

Public Property Get HorzBinsCount() As Long
HorzBinsCount = mHorzBinsCnt
End Property

Public Property Let HorzBinsCount(ByVal NewCount As Long)
mHorzBinsCnt = NewCount
Call CalculateGridParameters
End Property

Public Property Get VertBinsCount() As Long
VertBinsCount = mVertBinsCnt
End Property

Public Property Let VertBinsCount(ByVal NewCount As Long)
mVertBinsCnt = NewCount
Call CalculateGridParameters
End Property

Private Sub CalculateGridParameters()
'-------------------------------------------------------------------------------
'to keep it simple calculate always both horizontal and vertical grid parameters
'-------------------------------------------------------------------------------
On Error Resume Next
Select Case mHorzAutoMode
Case gamWidthConst
     mHorzBinsCnt = CLng((mXMax - mXMin) / mHorzWidth)
Case gamBinsCntConst
     mHorzWidth = CSng((mXMax - mXMin) / mHorzBinsCnt)
End Select

Select Case mVertAutoMode
Case gamWidthConst
     mVertBinsCnt = CLng((mYMax - mYMin) / mVertWidth)
Case gamBinsCntConst
     mVertWidth = CSng((mYMax - mYMin) / mVertBinsCnt)
End Select
End Sub


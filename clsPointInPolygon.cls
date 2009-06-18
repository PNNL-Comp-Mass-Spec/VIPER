VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPointInPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtPointType
    X As Long
    Y As Long
End Type

Private Type udtPointDblType
    X As Long
    Y As Double
End Type

Dim mPolygonVertexCount As Integer
Dim mPolygonDef() As udtPointType

Dim mErrorMessage As String

Public Property Get ErrorMessage()
    ErrorMessage = mErrorMessage
End Property


Private Sub Class_Initialize()
    mErrorMessage = ""
    
    ' Define an empty triangle, centered at 0,0
    Dim lngTriangleX(2) As Long
    Dim lngTriangleY(2) As Long
    
    lngTriangleX(0) = 0
    lngTriangleX(1) = 0
    lngTriangleX(2) = 0
    
    lngTriangleY(0) = 0
    lngTriangleY(1) = 0
    lngTriangleY(2) = 0
    
    SetPolygon 3, lngTriangleX, lngTriangleY
End Sub

' Defines the polygon for which points will be tested
' lngXVertexPoints() and lngYVertexPoints() should contain the X,Y points that define the polygon
' They should range from 0 to intVertexCount-1
Public Function SetPolygon(ByVal intVertexCount As Integer, ByRef lngXVertexPoints() As Long, ByRef lngYVertexPoints() As Long) As Boolean

    Dim intIndex As Integer

On Error GoTo SetPolygonErrorHandler

    If intVertexCount <= 0 Then
        SetPolygon = False
        Exit Function
    End If
    
    mPolygonVertexCount = intVertexCount
    ReDim mPolygonDef(mPolygonVertexCount - 1)

    For intIndex = 0 To intVertexCount - 1
        mPolygonDef(intIndex).X = lngXVertexPoints(intIndex)
        mPolygonDef(intIndex).Y = lngYVertexPoints(intIndex)
    Next intIndex
    
    SetPolygon = True
    
    Exit Function

SetPolygonErrorHandler:
    mErrorMessage = "Error in clsPointInPolygon.SetPolygon: " & Err.Description
    Debug.Assert False
    
    SetPolygon = False

End Function

' Note: Before calling this function, Call SetPolygon() or SetPolygonUDT() to define the polygon
Public Function TestPointInPolygon(ByVal lngX As Long, ByVal lngY As Long, Optional blnUseAlgorithmB As Boolean = False) As Boolean
    
    Dim udtPoint As udtPointType
    udtPoint.X = lngX
    udtPoint.Y = lngY
    
    If blnUseAlgorithmB Then
        TestPointInPolygon = TestPointInsidePolygonB(udtPoint)
    Else
        TestPointInPolygon = TestPointInsidePolygonA(udtPoint)
    End If
    
End Function

' Note: Before calling this function, Call SetPolygon() or SetPolygonUDT() to define the polygon
Public Function TestPointInPolygonDbl(ByVal dblX As Double, ByVal dblY As Double) As Boolean
    
    Dim udtPoint As udtPointDblType
    udtPoint.X = dblX
    udtPoint.Y = dblY
    
    TestPointInPolygonDbl = TestPointInsidePolygonDbl(udtPoint)
    
End Function

Private Function TestPointInsidePolygonA(ByRef udtPoint As udtPointType) As Boolean
    ' mPolygonDef() includes the points that makeup the polygon
    ' A four sided polygon (e.g. square or trapezoid) should have mPolygonVertexCount = 4, and should define them using
    '  mPolygonDef(0), mPolygonDef(1), mPolygonDef(2), and mPolygonDef(3)
    '
    ' Code from http://local.wasp.uwa.edu.au/~pbourke/geometry/insidepoly/
    ' Written by Paul Bourke in November 1987
    
    Dim intCounter As Long
    Dim i As Long
    
    Dim xinters As Double
    
    Dim p1 As udtPointType
    Dim p2 As udtPointType
    
On Error GoTo TestPointInsidePolygonErrorHandler

    intCounter = 0
    p1 = mPolygonDef(0)
  
    For i = 1 To mPolygonVertexCount
        p2 = mPolygonDef(i Mod mPolygonVertexCount)
        If udtPoint.Y > MinTwoLongs(p1.Y, p2.Y) Then
            If udtPoint.Y <= MaxTwoLongs(p1.Y, p2.Y) Then
                If udtPoint.X <= MaxTwoLongs(p1.X, p2.X) Then
                    If p1.Y <> p2.Y Then
                        xinters = (udtPoint.Y - p1.Y) * (p2.X - p1.X) / (p2.Y - p1.Y) + p1.X
                        If p1.X = p2.X Or udtPoint.X <= xinters Then
                            intCounter = intCounter + 1
                        End If
                    End If
                End If
            End If
        End If
        p1 = p2
    Next i
    
    If (intCounter Mod 2 = 0) Then
        TestPointInsidePolygonA = False
    Else
        TestPointInsidePolygonA = True
    End If

    Exit Function

TestPointInsidePolygonErrorHandler:
    Debug.Print "Error in TestPointInsidePolygonA: " & Err.Description

End Function

Private Function TestPointInsidePolygonB(ByRef udtPoint As udtPointType) As Boolean
    
    ' Code from http://local.wasp.uwa.edu.au/~pbourke/geometry/insidepoly/
    'Algorithm by Randolph Franklin
    
    Dim intCounter As Long
    
    Dim i As Integer
    Dim j As Integer
    Dim blnInsideFlag As Boolean

On Error GoTo TestPointInsidePolygonErrorHandlerB

    blnInsideFlag = False
      
    j = mPolygonVertexCount - 1
    
    For i = 0 To mPolygonVertexCount - 1
        If ((mPolygonDef(i).Y <= udtPoint.Y) And (udtPoint.Y < mPolygonDef(j).Y)) Or _
           ((mPolygonDef(j).Y <= udtPoint.Y) And (udtPoint.Y < mPolygonDef(i).Y)) Then
            
            If udtPoint.X < (mPolygonDef(j).X - mPolygonDef(i).X) * _
                            (udtPoint.Y - mPolygonDef(i).Y) / _
                            (mPolygonDef(j).Y - mPolygonDef(i).Y) + _
                            mPolygonDef(i).X Then
                blnInsideFlag = Not blnInsideFlag
            End If
            
        End If
                
        j = i
    Next i
    
      
    TestPointInsidePolygonB = blnInsideFlag
    Exit Function

TestPointInsidePolygonErrorHandlerB:
    Debug.Print "Error in TestPointInsidePolygonB: " & Err.Description

End Function

Private Function TestPointInsidePolygonDbl(ByRef udtPoint As udtPointDblType) As Boolean
    
    ' Code from http://local.wasp.uwa.edu.au/~pbourke/geometry/insidepoly/
    'Algorithm by Randolph Franklin
    
    Dim intCounter As Long
    
    Dim i As Integer
    Dim j As Integer
    Dim blnInsideFlag As Boolean

On Error GoTo TestPointInsidePolygonDblErrorHandler

    blnInsideFlag = False
      
    j = mPolygonVertexCount - 1
    
    For i = 0 To mPolygonVertexCount - 1
        If ((mPolygonDef(i).Y <= udtPoint.Y) And (udtPoint.Y < mPolygonDef(j).Y)) Or _
           ((mPolygonDef(j).Y <= udtPoint.Y) And (udtPoint.Y < mPolygonDef(i).Y)) Then
            
            If udtPoint.X < (mPolygonDef(j).X - mPolygonDef(i).X) * _
                            (udtPoint.Y - mPolygonDef(i).Y) / _
                            (mPolygonDef(j).Y - mPolygonDef(i).Y) + _
                            mPolygonDef(i).X Then
                blnInsideFlag = Not blnInsideFlag
            End If
            
        End If
                
        j = i
    Next i
    
      
    TestPointInsidePolygonDbl = blnInsideFlag
    Exit Function

TestPointInsidePolygonDblErrorHandler:
    Debug.Print "Error in TestPointInsidePolygonDbl: " & Err.Description

End Function


Private Function MinTwoLongs(ByVal lngX As Long, ByVal lngY As Long) As Long
    If lngX < lngY Then
        MinTwoLongs = lngX
    Else
        MinTwoLongs = lngY
    End If
End Function

Private Function MaxTwoLongs(ByVal lngX As Long, ByVal lngY As Long) As Long
    If lngX > lngY Then
        MaxTwoLongs = lngX
    Else
        MaxTwoLongs = lngY
    End If
End Function

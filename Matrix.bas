Attribute VB_Name = "Module8"
'last modified 03/16/2000 nt
'all matrices are saved as variant arrays for easier passing
'-----------------------------------------------------------
Option Explicit

Public vAxXSelectMatrix As Variant
Public vAxYSelectMatrix As Variant

Public vHXJobMatrix As Variant
Public vHYJobMatrix As Variant
Public vVXJobMatrix As Variant
Public vVYJobMatrix As Variant
Public vVYBJobMatrix As Variant
Public vLgndJobMatrix As Variant
Public vYHNumJobMatrix As Variant
Public vXVNumJobMatrix As Variant
Public vFNDJobMatrix As Variant
Public vLgndCSMatrix As Variant

Public Sub InitSelectMatrices()

Dim aMatrix(1 To 4, 1 To 4) As Long

aMatrix(1, 1) = 1
aMatrix(1, 2) = 2
aMatrix(1, 3) = 1
aMatrix(1, 4) = 2
aMatrix(2, 1) = 4
aMatrix(2, 2) = 3
aMatrix(2, 3) = 4
aMatrix(2, 4) = 3
aMatrix(3, 1) = 1
aMatrix(3, 2) = 2
aMatrix(3, 3) = 1
aMatrix(3, 4) = 2
aMatrix(4, 1) = 4
aMatrix(4, 2) = 3
aMatrix(4, 3) = 4
aMatrix(4, 4) = 3
vAxXSelectMatrix = aMatrix

aMatrix(1, 1) = 1
aMatrix(1, 2) = 1
aMatrix(1, 3) = 2
aMatrix(1, 4) = 2
aMatrix(2, 1) = 1
aMatrix(2, 2) = 1
aMatrix(2, 3) = 2
aMatrix(2, 4) = 2
aMatrix(3, 1) = 3
aMatrix(3, 2) = 3
aMatrix(3, 3) = 4
aMatrix(3, 4) = 4
aMatrix(4, 1) = 3
aMatrix(4, 2) = 3
aMatrix(4, 3) = 4
aMatrix(4, 4) = 4
vAxYSelectMatrix = aMatrix

End Sub


Public Function GetJobVector(ByVal Ind As Long, ByVal SelectMatrix As Variant, ByVal JobMatrix As Variant) As Variant
'returns vector (variant array) appropriate for this job and coosys selection
Dim JobColumn As Integer
Dim I As Integer
Dim aRes(1 To 4) As Long
On Error GoTo err_GetJobVector

With GelBody(Ind).csMyCooSys
    JobColumn = SelectMatrix(.csOriginXY, .csOrigin)
End With
For I = 1 To 4
    aRes(I) = JobMatrix(I, JobColumn)
Next I
GetJobVector = aRes
Exit Function

err_GetJobVector:
GetJobVector = Null
End Function

Public Sub InitJobMatrices()
Dim aMatrix(1 To 4, 1 To 4) As Long

'X coordinates of the Horizontal axis label matrix
aMatrix(1, 1) = 0
aMatrix(2, 1) = 1
aMatrix(3, 1) = 1
aMatrix(4, 1) = -1
aMatrix(1, 2) = 1
aMatrix(2, 2) = 0
aMatrix(3, 2) = -1
aMatrix(4, 2) = 1
aMatrix(1, 3) = 0
aMatrix(2, 3) = 1
aMatrix(3, 3) = 1
aMatrix(4, 3) = 0
aMatrix(1, 4) = 1
aMatrix(2, 4) = 0
aMatrix(3, 4) = -1
aMatrix(4, 4) = 0
vHXJobMatrix = aMatrix
vFNDJobMatrix = TransposeCol(vHXJobMatrix)

'X coordinates of the Vertical axis label matrix
aMatrix(1, 1) = 1
aMatrix(2, 1) = 0
aMatrix(3, 1) = 1
aMatrix(4, 1) = 0
aMatrix(1, 2) = 0
aMatrix(2, 2) = 1
aMatrix(3, 2) = -1
aMatrix(4, 2) = 0
aMatrix(1, 3) = 1
aMatrix(2, 3) = 0
aMatrix(3, 3) = 1
aMatrix(4, 3) = 1
aMatrix(1, 4) = 0
aMatrix(2, 4) = 1
aMatrix(3, 4) = -1
aMatrix(4, 4) = -1
vVXJobMatrix = aMatrix

'Y coordinates of the Horizontal axis label matrix
aMatrix(1, 1) = 1
aMatrix(2, 1) = 0
aMatrix(3, 1) = -1
aMatrix(4, 1) = 0
aMatrix(1, 2) = 0
aMatrix(2, 2) = 1
aMatrix(3, 2) = 1
aMatrix(4, 2) = 0
aMatrix(1, 3) = 0
aMatrix(2, 3) = 1
aMatrix(3, 3) = 1
aMatrix(4, 3) = 1
aMatrix(1, 4) = 1
aMatrix(2, 4) = 0
aMatrix(3, 4) = -1
aMatrix(4, 4) = -1
vHYJobMatrix = aMatrix

'Y coordinates of the Vertical axis label matrix
aMatrix(1, 1) = 0
aMatrix(2, 1) = 1
aMatrix(3, 1) = 1
aMatrix(4, 1) = 1
aMatrix(1, 2) = 1
aMatrix(2, 2) = 0
aMatrix(3, 2) = -1
aMatrix(4, 2) = -1
aMatrix(1, 3) = 1
aMatrix(2, 3) = 0
aMatrix(3, 3) = -1
aMatrix(4, 3) = 0
aMatrix(1, 4) = 0
aMatrix(2, 4) = 1
aMatrix(3, 4) = 1
aMatrix(4, 4) = 0
vVYJobMatrix = aMatrix
vVYBJobMatrix = TransposeCol(vVYJobMatrix)

'legend matrix is a bit different, there are only 2 cases
'and no select matrix, it's used directly with ScalarProduct1
aMatrix(1, 1) = 1
aMatrix(2, 1) = 1
aMatrix(3, 1) = 1
aMatrix(4, 1) = 1
aMatrix(1, 2) = 1
aMatrix(2, 2) = -1
aMatrix(3, 2) = -1
aMatrix(4, 2) = -1
aMatrix(1, 3) = 0
aMatrix(2, 3) = 0
aMatrix(3, 3) = 0
aMatrix(4, 3) = 0
aMatrix(1, 4) = 0
aMatrix(2, 4) = 0
aMatrix(3, 4) = 0
aMatrix(4, 4) = 0
vLgndJobMatrix = aMatrix

'X coordinates of the Numbers on the Horizontal axis
aMatrix(1, 1) = 1
aMatrix(2, 1) = -1
aMatrix(3, 1) = -1
aMatrix(4, 1) = 0
aMatrix(1, 2) = 1
aMatrix(2, 2) = 1
aMatrix(3, 2) = 1
aMatrix(4, 2) = 0
aMatrix(1, 3) = 1
aMatrix(2, 3) = 1
aMatrix(3, 3) = 1
aMatrix(4, 3) = 1
aMatrix(1, 4) = 1
aMatrix(2, 4) = -1
aMatrix(3, 4) = -1
aMatrix(4, 4) = -1
vYHNumJobMatrix = aMatrix

'X coordinates of the Numbers on the Vertical axis
aMatrix(1, 1) = 1
aMatrix(2, 1) = -1
aMatrix(3, 1) = -1
aMatrix(4, 1) = -1
aMatrix(1, 2) = 1
aMatrix(2, 2) = 1
aMatrix(3, 2) = 1
aMatrix(4, 2) = 1
aMatrix(1, 3) = 1
aMatrix(2, 3) = -1
aMatrix(3, 3) = -1
aMatrix(4, 3) = 0
aMatrix(1, 4) = 1
aMatrix(2, 4) = 1
aMatrix(3, 4) = 1
aMatrix(4, 4) = 0
vXVNumJobMatrix = aMatrix

'CS legend matrix
aMatrix(1, 1) = 0
aMatrix(2, 1) = 1
aMatrix(3, 1) = -1
aMatrix(4, 1) = -1
aMatrix(1, 2) = 1
aMatrix(2, 2) = 0
aMatrix(3, 2) = 1
aMatrix(4, 2) = 1
aMatrix(1, 3) = 0
aMatrix(2, 3) = 1
aMatrix(3, 3) = 1
aMatrix(4, 3) = 1
aMatrix(1, 4) = 1
aMatrix(2, 4) = 0
aMatrix(3, 4) = -1
aMatrix(4, 4) = -1
vLgndCSMatrix = aMatrix

End Sub

Public Function SP(ByVal vVec1 As Variant, ByVal X1 As Long, ByVal x2 As Long, ByVal X3 As Long, ByVal x4 As Long) As Long
'calculates scalar product of two vectors
Dim aVec2(1 To 4) As Long
Dim I  As Integer
Dim tmp As Long

aVec2(1) = X1
aVec2(2) = x2
aVec2(3) = X3
aVec2(4) = x4
tmp = 0
For I = 1 To 4
    tmp = tmp + vVec1(I) * aVec2(I)
Next I
SP = tmp
End Function


Private Function TransposeCol(ByVal vMatrix As Variant) As Variant
'returns variant arrray with transposed columns
Dim aMatrix(1 To 4, 1 To 4) As Long
Dim I As Integer, j As Integer
For j = 1 To 4
    For I = 1 To 4
        aMatrix(I, 4 - j + 1) = vMatrix(I, j)
    Next I
Next j
TransposeCol = aMatrix
End Function

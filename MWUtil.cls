VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MWUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MW range class - locator of range of indexes in array
'used with unique counts and searching MT tag databases
'array MW should be sorted asc. before being passed to Fill
'function; otherwise function FindIndexRange will not work
'correctly and can even result in infinite loop!
'----------------------------------------------------------
'last modified 05/17/2000 nt
Option Explicit

Dim clsMW() As Double  'array of molecular masses

Private LoIndex As Long
Private HiIndex As Long

Private MWToFind As Double
Private MWAbsTol As Double

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal ByteCount As Long)

Public Function Fill(ThisMW() As Double) As Boolean
Dim ArrSize As Long
Dim Cnt As Long

On Error GoTo err_Fill
LoIndex = LBound(ThisMW)
HiIndex = UBound(ThisMW)
Cnt = HiIndex - LoIndex + 1
If Cnt > 0 Then
   ReDim clsMW(LoIndex To HiIndex)
   ArrSize = Cnt * Len(ThisMW(LoIndex))
   CopyMemory clsMW(LoIndex), ThisMW(LoIndex), ArrSize
   Fill = True
   Exit Function
End If

err_Fill:
Fill = False
End Function

Private Sub Class_Terminate()
On Error Resume Next
Erase clsMW
End Sub

Public Function FindIndexRange(ByVal sMW As Double, _
                               ByVal sTol As Double, _
                               ByRef FirstInd As Long, _
                               ByRef LastInd As Long) As Boolean
'---------------------------------------------------------------
'fills First and Last index of range in which MW array elements
'match sMW with absolute tolerance sTol; returns True if such
'indexes exist, False on any error or if indexes don't exist
'---------------------------------------------------------------
On Error GoTo exit_FindIndexRange
MWToFind = sMW
MWAbsTol = sTol
FirstInd = LoIndex
LastInd = HiIndex
FindMWRange FirstInd, LastInd
If FirstInd > LastInd Then
   FindIndexRange = False
Else
   FindIndexRange = True
End If
exit_FindIndexRange:
End Function


Private Sub FindMWRange(MinInd As Long, MaxInd As Long)
'-------------------------------------------------------------------
'recursive procedure to find all elements(range of elements indexes)
'of MW array that are closer to MWToFind than MWAbsTol
'if found then MinInd <= MaxInd; if not found then MinInd > MaxInd
'-------------------------------------------------------------------
Dim MidInd As Long
Dim LeftDone As Boolean
Dim RightDone As Boolean
Dim LeftIndex As Long
Dim RightIndex As Long

MidInd = (MinInd + MaxInd) \ 2
If MidInd = MinInd Then  'Min and Max next to each other
   If Abs(MWToFind - clsMW(MinInd)) > MWAbsTol Then MinInd = MaxInd
   If Abs(MWToFind - clsMW(MaxInd)) > MWAbsTol Then MaxInd = MidInd
   Exit Sub         'we are done
End If
If clsMW(MidInd) > MWToFind + MWAbsTol Then        'we are out of range on right
   MaxInd = MidInd
   FindMWRange MinInd, MaxInd
ElseIf clsMW(MidInd) < MWToFind - MWAbsTol Then    'we are out of range on left
   MinInd = MidInd
   FindMWRange MinInd, MaxInd
Else                                        'we are inside the range; go
   LeftIndex = MidInd                            'figure out borders
   Do Until LeftDone
      LeftIndex = LeftIndex - 1
      If LeftIndex < MinInd Then
         LeftDone = True
      Else
         If Abs(MWToFind - clsMW(LeftIndex)) > MWAbsTol Then LeftDone = True
      End If
   Loop
   RightIndex = MidInd
   Do Until RightDone
      RightIndex = RightIndex + 1
      If RightIndex > MaxInd Then
         RightDone = True
      Else
         If Abs(MWToFind - clsMW(RightIndex)) > MWAbsTol Then RightDone = True
      End If
   Loop
   MinInd = LeftIndex + 1
   MaxInd = RightIndex - 1
End If
End Sub

Public Function GetMWByIndex(lngIndex As Long) As Double
    On Error Resume Next
    If lngIndex >= LoIndex And lngIndex <= HiIndex Then
        GetMWByIndex = clsMW(lngIndex)
    Else
        GetMWByIndex = 0
    End If
End Function


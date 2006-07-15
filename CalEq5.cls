VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalEq5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'last modified 10/15/2001 nt
'class for Mass to Frequency to Mass calculations
'based on calibration equation m/z = A/(f - B)
'------------------------------------------------
'created: 10/15/2001 nt
'last modified: 10/15/2001 nt
'------------------------------------------------
Option Explicit

Dim mA As Double
Dim mB As Double

Dim mAOK As Boolean
Dim mBOK As Boolean

Public Property Get a() As Double
    a = mA
End Property

Public Property Let a(ByVal ArgA As Double)
    mA = ArgA
    mAOK = True
End Property

Public Property Get B() As Double
    B = mB
End Property

Public Property Let B(ByVal ArgB As Double)
    mB = ArgB
    mBOK = True
End Property

Public Property Get C() As Double
    ' C isn't used in this class
    C = 0
End Property

Public Property Let C(ByVal ArgC As Double)
    ' C isn't used in this class
End Property

Public Property Get CalDescription() As String
    CalDescription = Me.Equation & " with A = " & Trim(Me.a) & " and B = " & Trim(Me.B)
End Property

Public Function MagnetFreq(ByVal MOverZ As Double) As Double
'-----------------------------------------------------------
'this calibration equation form does not predict for
'magnetron frequency
'-----------------------------------------------------------
MagnetFreq = -1
End Function

Public Function CyclotFreq(ByVal MOverZ As Double) As Double
'-----------------------------------------------------------
'returns cyclotron frequency for m/z
'-----------------------------------------------------------
Dim CF As Double
On Error GoTo exit_CyclotFreq:
CF = -1
If mAOK And mBOK And MOverZ >= 0 Then
   CF = mB + (mA / MOverZ)
End If

exit_CyclotFreq:
CyclotFreq = CF
End Function


Public Function MOverZ(ByVal Freq As Double) As Double
'-----------------------------------------------------
'returns m/z for specified frequency
'-----------------------------------------------------
If Freq > 0 Then
   MOverZ = mA / (Freq - mB)
Else
   MOverZ = -1
End If
End Function

Public Property Get Equation() As Variant
Equation = CAL_EQUATION_5
End Property

Public Property Get EquationOK() As Boolean
EquationOK = mAOK And mBOK
End Property


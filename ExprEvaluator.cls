VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExprEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'last modified 06/16/2000 nt
Option Explicit
'evaluator object of simple numeric expressions
'Properties:
'Expr - read/write - expression
'IsExprValid - read - True if expression is valid in Vars grammar
'Public function
'ExprVal - returns value of expression; takes as arguments variables values
'operators precendence constants
Const eePREC_NONE = 8
Const eePREC_UNARY = 7   ' Not actually used.
Const eePREC_POWER = 6
Const eePREC_TIMES = 5
Const eePREC_DIV = 4
Const eePREC_INT_DIV = 3
Const eePREC_MOD = 2
Const eePREC_PLUS = 1

Dim mExpr As String
Dim mValidExpr As Boolean

Public Vars As Collection           'collection with enumerated variable names
Dim vals() As Double

Private Sub Class_Initialize()
Set Vars = New Collection
End Sub

Private Function ExprEva(ByVal Expr As String) As Boolean
'returns True if Expr is legal expression in Vars and
'math operators defined grammar; it does not calculate anything
Dim ParCnt As Integer
Dim pos As Integer
Dim ExprLen As Integer
Dim ch As String
Dim LExpr As String
Dim RExpr As String
Dim value As String
Dim BestPos As Integer
Dim BestPrec As Integer

Expr = Trim$(Expr)
ExprLen = Len(Expr)
If ExprLen = 0 Then Exit Function  'return False
BestPrec = eePREC_NONE
' Find the operator with the lowest precedence.
' Look for places where there are no open parentheses.
For pos = 1 To ExprLen
    ' Examine the next character.
    ch = Mid$(Expr, pos, 1)
    If ch = " " Then       ' Just skip spaces.
    ElseIf ch = "(" Then   ' Increase the open parentheses count.
       ParCnt = ParCnt + 1
    ElseIf ch = ")" Then   ' Decrease the open parentheses count.
       ParCnt = ParCnt - 1
       ' If parens < 0, too many ')'s.
       If ParCnt < 0 Then Exit Function    'return False
    ElseIf ParCnt = 0 Then
       ' See if this is an operator.
       If ch = "^" Or ch = "*" Or ch = "/" Or ch = "\" Or ch = "%" Or ch = "+" Or ch = "-" Then                 ' An operator after an operator is unary.
          Select Case ch
          Case "^"
             If BestPrec >= eePREC_POWER Then
                BestPrec = eePREC_POWER
                BestPos = pos
             End If
          Case "*", "/"
             If BestPrec >= eePREC_TIMES Then
                BestPrec = eePREC_TIMES
                BestPos = pos
             End If
          Case "\"
             If BestPrec >= eePREC_INT_DIV Then
                BestPrec = eePREC_INT_DIV
                BestPos = pos
             End If
          Case "%"
             If BestPrec >= eePREC_MOD Then
                BestPrec = eePREC_MOD
                BestPos = pos
             End If
          Case "+", "-" ' Ignore unary operators if on 1st position
             If (pos > 1) And (BestPrec >= eePREC_PLUS) Then
                BestPrec = eePREC_PLUS
                BestPos = pos
             End If
          End Select
       End If
    End If
Next pos
' If the parenteses count is not zero, there's a ')' missing
If ParCnt <> 0 Then Exit Function   'return False
' Hopefully we have the operator
If BestPrec < eePREC_NONE Then
   LExpr = Left$(Expr, BestPos - 1)
   RExpr = Right$(Expr, ExprLen - BestPos)
   Select Case Mid$(Expr, BestPos, 1)
   Case "^", "*", "/", "\", "%", "+", "-"
        ExprEva = ExprEva(LExpr) And ExprEva(RExpr)
   End Select
   Exit Function
End If
' If we do not yet have an operator, there are several possibilities:
' 1. expr is (expr2) for some expr2
' 2. expr is -expr2 or +expr2 for some expr2
' 3. expr is Fun(expr2) for a function Fun
' 4. expr is a variable name
' 5. It's a numeric constant
' Look for (expr2)
If Left$(Expr, 1) = "(" And Right$(Expr, 1) = ")" Then
   ' Remove the parentheses.
   ExprEva = ExprEva(Mid$(Expr, 2, ExprLen - 2))
   Exit Function
End If
' Look for -expr2
If Left$(Expr, 1) = "-" Then
   ExprEva = -ExprEva(Right$(Expr, ExprLen - 1))
   Exit Function
End If
' Look for +expr2
If Left$(Expr, 1) = "+" Then
   ExprEva = ExprEva(Right$(Expr, ExprLen - 1))
   Exit Function
End If
' Look for Fun(expr2) (some basic functions)
If ExprLen > 5 And Right$(Expr, 1) = ")" Then
   LExpr = LCase$(Left$(Expr, 4))
   RExpr = Mid$(Expr, 5, ExprLen - 5)
   Select Case LExpr
   Case "sin(", "cos(", "tan(", "sqr("
        ExprEva = ExprEva(RExpr)
        Exit Function
   End Select
End If
' See if it's a variable name
On Error Resume Next
value = Vars.Item(Expr)
If Err.Number = 0 Then
   ExprEva = True
   Exit Function
End If
' It must be a numeric constant
ExprEva = IsNumeric(Expr)
End Function

Public Property Get Expr() As String
Expr = mExpr
End Property

Public Property Let Expr(ByVal sExpr As String)
mExpr = Trim$(sExpr)
mValidExpr = ExprEva(mExpr)
End Property

Public Property Get IsExprValid() As Boolean
'read only property
IsExprValid = mValidExpr
End Property

Public Function ExprVal(Args() As Long) As Double
Dim i As Integer
Dim ArgCnt
On Error GoTo err_ExprVal
ArgCnt = UBound(Args)
If ArgCnt <> Vars.Count Or ArgCnt <= 0 Then
   Err.Raise 9, "Expression Evaluator"
Else
   ReDim vals(1 To ArgCnt)
   For i = 1 To ArgCnt
       vals(i) = Args(i)
   Next i
   ExprVal = ExprEvaVal(mExpr)
End If
Exit Function

err_ExprVal:
Err.Raise Err.Number, "Expression Evaluator"
End Function

Private Function ExprEvaVal(ByVal Expr As String) As Double
'this function returns value of the expression
'Vals() passes variable values in the same order as they
'are added to collection Vars; it raises error on any error
Dim ParCnt As Integer
Dim pos As Integer
Dim ExprLen As Integer
Dim ch As String
Dim LExpr As String
Dim RExpr As String
Dim value As String
Dim BestPos As Integer
Dim BestPrec As Integer
On Error GoTo err_ExprEvaVal

Expr = Trim$(Expr)
ExprLen = Len(Expr)
If ExprLen = 0 Then Exit Function
' So far we have nothing.
BestPrec = eePREC_NONE
' Find the operator with the lowest precedence.
' Look for places where there are no open parentheses.
For pos = 1 To ExprLen
    ' Examine the next character.
    ch = Mid$(Expr, pos, 1)
    If ch = " " Then       ' Just skip spaces.
    ElseIf ch = "(" Then       ' Increase the open parentheses count.
       ParCnt = ParCnt + 1
    ElseIf ch = ")" Then       ' Decrease the open parentheses count.
       ParCnt = ParCnt - 1
       ' If parens < 0, too many ')'s, but because all functions are first evaluated this will not happen
    ElseIf ParCnt = 0 Then     ' See if this is an operator.
       If ch = "^" Or ch = "*" Or ch = "/" Or ch = "\" Or ch = "%" Or ch = "+" Or ch = "-" Then
         Select Case ch
         Case "^"
            If BestPrec >= eePREC_POWER Then
               BestPrec = eePREC_POWER
               BestPos = pos
            End If
         Case "*", "/"
            If BestPrec >= eePREC_TIMES Then
               BestPrec = eePREC_TIMES
               BestPos = pos
            End If
         Case "\"
            If BestPrec >= eePREC_INT_DIV Then
               BestPrec = eePREC_INT_DIV
               BestPos = pos
            End If
         Case "%"
            If BestPrec >= eePREC_MOD Then
               BestPrec = eePREC_MOD
               BestPos = pos
            End If
         Case "+", "-"  'ignore unary operators at the 1st position
            If (pos > 1) And (BestPrec >= eePREC_PLUS) Then
               BestPrec = eePREC_PLUS
               BestPos = pos
            End If
         End Select
       End If
    End If
Next pos
' No need to check for parenteses if already evaluated
' Hopefully we have the operator.
If BestPrec < eePREC_NONE Then
   LExpr = Left$(Expr, BestPos - 1)
   RExpr = Right$(Expr, ExprLen - BestPos)
   Select Case Mid$(Expr, BestPos, 1)
   Case "^"
        ExprEvaVal = ExprEvaVal(LExpr) ^ ExprEvaVal(RExpr)
   Case "*"
        ExprEvaVal = ExprEvaVal(LExpr) * ExprEvaVal(RExpr)
   Case "/"
        ExprEvaVal = ExprEvaVal(LExpr) / ExprEvaVal(RExpr)
   Case "\"
        ExprEvaVal = ExprEvaVal(LExpr) \ ExprEvaVal(RExpr)
   Case "%"
        ExprEvaVal = ExprEvaVal(LExpr) Mod ExprEvaVal(RExpr)
   Case "+"
        ExprEvaVal = ExprEvaVal(LExpr) + ExprEvaVal(RExpr)
   Case "-"
        ExprEvaVal = ExprEvaVal(LExpr) - ExprEvaVal(RExpr)
   End Select
   Exit Function
End If
' If we do not yet have an operator, there
' are several possibilities:
' 1. expr is (expr2) for some expr2.
' 2. expr is -expr2 or +expr2 for some expr2.
' 3. expr is Fun(expr2) for a function Fun.
' 4. expr is a primitive.
' 5. It's a numeric constant
' Look for (expr2).
If Left$(Expr, 1) = "(" And Right$(Expr, 1) = ")" Then
   ExprEvaVal = ExprEvaVal(Mid$(Expr, 2, ExprLen - 2))
   Exit Function
End If
' Look for -expr2.
If Left$(Expr, 1) = "-" Then
   ExprEvaVal = -ExprEvaVal(Right$(Expr, ExprLen - 1))
   Exit Function
End If
' Look for +expr2.
If Left$(Expr, 1) = "+" Then
   ExprEvaVal = ExprEvaVal(Right$(Expr, ExprLen - 1))
   Exit Function
End If
' Look for Fun(expr2).
If ExprLen > 5 And Right$(Expr, 1) = ")" Then
   LExpr = LCase$(Left$(Expr, 4))
   RExpr = Mid$(Expr, 5, ExprLen - 5)
   Select Case LExpr
   Case "sin("
        ExprEvaVal = Sin(ExprEvaVal(RExpr))
        Exit Function
   Case "cos("
        ExprEvaVal = Cos(ExprEvaVal(RExpr))
        Exit Function
   Case "tan("
        ExprEvaVal = Tan(ExprEvaVal(RExpr))
        Exit Function
   Case "sqr("
        ExprEvaVal = Sqr(ExprEvaVal(RExpr))
        Exit Function
   End Select
End If
' See if it's a variable name
On Error Resume Next
value = vals(Vars.Item(Expr))
If Err.Number = 0 Then
   ExprEvaVal = CDbl(value)
   Exit Function
End If
' It must be a constant
ExprEvaVal = CDbl(Expr)
Exit Function

err_ExprEvaVal:
Err.Raise Err.Number, "Expression Evaluator"
End Function

Private Sub Class_Terminate()
On Error Resume Next
Set Vars = Nothing
Erase vals
End Sub

Attribute VB_Name = "Operator"
'==============================
'Ver:1.0.0
'Created:2021/6/5
'Update:2021/6/5
'Require:m-IsType
'==============================
'Inc    ++  自增
'Dec    --  自减
'Asgn   =   赋值
'AAsgn  +=  加赋值
'SAsgn  -=  减赋值
'MAsgn  *=  乘赋值
'DAsgn  /=  除赋值
'Vor    ||  值或
'Des    []  解构
Option Explicit
Public Function opr(symbol As String, ByRef var1, Optional ByRef var2, Optional ByRef var3, Optional ByRef var4, _
Optional ByRef var5, Optional ByRef var6, Optional ByRef var7, Optional ByRef var8, Optional ByRef var9, _
Optional ByRef var10)
    Select Case symbol
        Case "++": opr = Inc(var1)
        Case "--": opr = Dec(var1)
        Case "=": Asgn var1, var2
        Case "+=": opr = AAsgn(var1, var2)
        Case "-=": opr = SAsgn(var1, var2)
        Case "*=": opr = MAsgn(var1, var2)
        Case "/=": opr = DAsgn(var1, var2)
        Case "||": Asgn opr, Vor(var1, var2)
        Case "[]"
            Call Des(var1, var2, var3, var4, _
            var5, var6, var7, var8, var9, var10)
        Case Else
            Err.Raise 1, Description:="Operator not support :" & symbol
    End Select
End Function

'increment
Public Function Inc(ByRef var1)
    var1 = var1 + 1: Inc = var1
End Function

'decrement
Public Function Dec(ByRef var1)
    var1 = var1 - 1: Dec = var1
End Function

'assignment
Public Sub Asgn(ByRef var1, ByRef val) 
    If IsObject(val) Then Set var1 = val _
    Else var1 = val
End Sub

'value or
Public Function Vor(ByRef var1, ByRef var2) 
'var1 isValide Then return var1   Else return var2
   If isType.isValide(var1) Then Asgn Vor, var1 _
   Else Asgn Vor, var2
End Function

Public Function AAsgn(ByRef var1, val)
    var1 = var1 + val: AAsgn = var1
End Function

Public Function SAsgn(ByRef var1, val)
    var1 = var1 - val: SAsgn = var1
End Function

Public Function MAsgn(ByRef var1, val)
    var1 = var1 * val: MAsgn = var1
End Function

Public Function DAsgn(ByRef var1, val)
    var1 = var1 / val: DAsgn = var1
End Function

'arrayDestructuring    
Public Sub Des(ByRef arr, ByRef var1, Optional ByRef var2, Optional ByRef var3, Optional ByRef var4, _
Optional ByRef var5, Optional ByRef var6, Optional ByRef var7, Optional ByRef var8, Optional ByRef var9)
   Dim lb As Long
   lb = LBound(arr)
   'array length < var count Then throw error 
    Asgn var1, arr(lb + 0)
    If IsMissing(var2) Then Exit Sub
    Asgn var2, arr(lb + 1)
    If IsMissing(var3) Then Exit Sub
    Asgn var3, arr(lb + 2)
    If IsMissing(var4) Then Exit Sub
    Asgn var4, arr(lb + 3)
    If IsMissing(var5) Then Exit Sub
    Asgn var5, arr(lb + 4)
    If IsMissing(var6) Then Exit Sub
    Asgn var6, arr(lb + 5)
    If IsMissing(var7) Then Exit Sub
    Asgn var7, arr(lb + 6)
    If IsMissing(var8) Then Exit Sub
    Asgn var8, arr(lb + 7)
    If IsMissing(var9) Then Exit Sub
    Asgn var9, arr(lb + 8)
End Sub




Attribute VB_Name = "IsType"
'========================
'Ver:1.0.0
'Created:2021/6/5
'Update:2021/6/5
'Require:m-Utils
'========================
'vb native function
'IsArray
'IsEmpty
'IsNull
Option Explicit
public Function isValide(ByRef param) As Boolean
    isValide = True 'init true
    Select Case TypeName(param)
        Case "Empty", "Null", "Nothing": isValide = False
        Case "Boolean": isValide = param
        Case "String": If param = "" Then isValide = False
        Case Else
            If IsArray(param) Then 'EmptyArray
                If isEmptyArray(param) Then isValide = False
            ElseIf isNumber(param) Then 'number 0
                If param = 0 Then isValide = False
            End If
    End Select
End Function

public Function isNothing(ByRef param) As Boolean
    If TypeName(param) = "Nothing" Then isEmpty = True
End Function

public Function isInt(ByRef param) As Boolean
    If Utils.arrayIndexOf( _
    Array("Long","LongLong","Byte","Integer"), _
    TypeName(param)) > -1 Then isInt = True
End Function

public Function isDecimal(ByRef param) As Boolean
    If Utils.arrayIndexOf( _
    Array("Single", "Double"), _
    TypeName(param)) > -1 Then isDecimal = True
End Function

public Function isNumber(ByRef param) As Boolean
    If isInt(param) Then
        isNumber = True
    ElseIf isDecimal(param) Then
        isNumber = True
    End If
End Function

public Function isEmptyArray(ByRef arr) As Boolean
    If UBound(arr) = -1 Then isEmptyArray = True
End Function



Attribute VB_Name = "Utils"
Option Explicit
Public Function arrayIndexOf(ByRef arr, ByRef val) As Long 
    arrayIndexOf = -1
    Dim i as Long 
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            arrayIndexOf = i
            Exit Function
        End If
    Next
End Function
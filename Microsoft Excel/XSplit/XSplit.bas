Attribute VB_Name = "XSplit"
Function XSplit(rng As String, delimiter As String, index As Integer)
On Error Resume Next

    If rng = "" Or delimiter = "" Or index = Null Then
        Exit Function
    End If
    
    Dim arr As Variant
    arr = Split(rng, delimiter)
    
    Dim count As Integer
    count = UBound(arr) + 1
    
    Dim indexClear As Integer
    indexClear = CInt(index)
    
    If indexClear > count Or indexClear <= 0 Then
        XSplit = "Out of range"
    Else
        XSplit = arr(indexClear - 1)
    End If
End Function

Attribute VB_Name = "Module1"
Public Function Regex(rng As Range, strPattern As String, Optional returnFirstItem As Boolean = False)
On Error Resume Next

    If (IsNull(rng) Or rng = "") Or (IsNull(strPattern) Or strPattern = "") Then
        Regex = "NULL"
    End If

    Dim rg As Object
    Set rg = CreateObject("VBScript.RegExp")
    With rg
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = strPattern
    End With
    
    Dim result As Boolean
    result = rg.Test(rng)
    
    If result Then
        If returnFirstItem Then
            Dim matches As Object
            Set matches = rg.Execute(rng)
            
            For Each Mch In matches
                Regex = Mch.Value
                Exit For
            Next
        Else
            Regex = True
        End If
    Else
        Regex = False
    End If

End Function


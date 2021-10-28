Attribute VB_Name = "Number2Word"
Option Explicit

Public H As Variant
Public S As Variant
Public D2 As Variant
Public D1 As Variant
Public Y As Variant

Function Number2Word(myNumber As Range) As String
On Error GoTo Mehrdad

    If myNumber Like "*E*" Or myNumber Like "*e*" Then
        Number2Word = "This number is very large"
        Exit Function
    End If

    H = Array("", "Â“«—", "„Ì·ÌÊ‰", "„Ì·Ì«—œ", "»Ì·ÌÊ‰", "»Ì·Ì«—œ", " —Ì·ÌÊ‰", " —Ì·Ì«—œ")
    S = Array("", "’œ", "œÊÌ” ", "”Ì’œ", "çÂ«—’œ", "Å«‰’œ", "‘‘’œ", "Â› ’œ", "Â‘ ’œ", "‰Â’œ")
    D2 = Array("", "", "»Ì” ", "”Ì", "çÂ·", "Å‰Ã«Â", "‘’ ", "Â› «œ", "Â‘ «œ", "‰Êœ")
    D1 = Array("œÂ", "Ì«“œÂ", "œÊ«“œÂ", "”Ì“œÂ", "çÂ«—œÂ", "Å«‰“œÂ", "‘«‰“œÂ", "Â›œÂ", "ÂÃœÂ", "‰Ê“œÂ")
    Y = Array("’›—", "Ìò", "œÊ", "”Â", "çÂ«—", "Å‰Ã", "‘‘", "Â› ", "Â‘ ", "‰Â")

    Dim xLen As Integer
    Dim xLen3Digit As Integer
    Dim xNumber As String
    Dim xResult As String
    
    xNumber = Trim(myNumber)
    xLen = Len(xNumber)
    xLen3Digit = Application.WorksheetFunction.RoundUp(xLen / 3, 0)

    Dim xArray3Digit(), xArray3String(), xArray3StringInverse() As String
    ReDim xArray3Digit(xLen3Digit)
    ReDim xArray3String(xLen3Digit)
    ReDim xArray3StringInverse(xLen3Digit)
    
    If xLen Mod 3 <> 0 Then
        Dim countZero As Integer
        countZero = 3 - (xLen Mod 3)
        xNumber = Space(countZero) & myNumber
        xNumber = Replace(xNumber, " ", "0")
    End If
    
    xLen = Len(xNumber)
    
    Dim x3Digit As String
    Dim i, xStart, xEnd As Integer
    For i = 0 To xLen3Digit - 1
        xStart = xLen - ((i * 3) + 3)
        
        x3Digit = Mid(xNumber, xStart + 1, 3)
        xArray3Digit(i) = x3Digit
        xArray3StringInverse(i) = Mini3Digit2String(x3Digit)
    Next i
    
    Dim j As Integer
    j = 0
    For i = xLen3Digit - 1 To 0 Step -1
        If xArray3StringInverse(i) <> "" Then
            xArray3String(j) = xArray3StringInverse(i) & " " & H(i)
        End If
        j = j + 1
    Next i
    
    xResult = Join(xArray3String, " Ê ")
    xResult = Replace(xResult, "  ", " ")
    xResult = Left(xResult, Len(xResult) - 2)
    
    For i = 0 To xLen3Digit
        xResult = Replace(xResult, "Ê Ê", "Ê")
    Next i
    Number2Word = xResult

Mehrdad:
    Exit Function
End Function

Function Mini3Digit2String(my3Digit As String) As String
On Error GoTo Mehrdad

    Dim x As Integer
    Dim xNumber As Integer
    Dim xResult(2) As String
    Dim xReturn As String
    Dim dg1, dg2, dg3 As String
    
    dg1 = Left(my3Digit, 1)
    dg2 = Mid(my3Digit, 2, 1)
    dg3 = Right(my3Digit, 1)
    
    'Check digit 1
    If dg1 <> "0" Then
        xResult(0) = S(CInt(dg1))
    End If
    
    'Check digit 2
    If dg2 <> 0 Then
        If dg2 = "1" Then
            xResult(1) = D1(CInt(dg2 & dg3) - 10)
        Else
            xResult(1) = D2(CInt(dg2))
        End If
    End If
    
    'Check digit 3
    If dg3 <> "0" And dg2 <> "1" Then
        xResult(2) = Y(CInt(dg3))
    End If
    
    
    xReturn = ""
    If xResult(0) <> "" Then
        xReturn = xResult(0)
    End If
    
    If xResult(1) <> "" Then
        If xReturn = "" Then
            xReturn = xResult(1)
        Else
            xReturn = xReturn & " Ê " & xResult(1)
        End If
    End If
    
    If xResult(2) <> "" Then
        If xReturn = "" Then
            xReturn = xResult(2)
        Else
            xReturn = xReturn & " Ê " & xResult(2)
        End If
    End If
    
    Mini3Digit2String = xReturn

Mehrdad:
    Exit Function
End Function

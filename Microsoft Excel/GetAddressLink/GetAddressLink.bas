Attribute VB_Name = "Module1"
Function GetAddressLink(rng As Range) As String
On Error Resume Next
     GetAddressLink = rng.Hyperlinks(1).Address
End Function

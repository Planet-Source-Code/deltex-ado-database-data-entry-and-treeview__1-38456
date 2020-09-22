Attribute VB_Name = "basFunc"
Option Explicit

Function getCode(ByVal the_Description As String) As Double
    Dim xCode As String, strTester As String
    Dim i As Integer, noChar As Integer
    
    noChar = Len(Trim(the_Description)) ' count number of characters
    
    For i = 1 To noChar
        strTester = Mid(the_Description, i, 1)  ' get the tester
        
        If strTester = ")" Then     ' compare with current character
            Exit For
        Else
            If strTester <> ")" Then    ' if not a parenthesis, concatenate with previous value
                xCode = xCode & strTester
            End If
        End If
    Next
    
    xCode = Mid(xCode, 2)
    
    getCode = Val(xCode)
End Function

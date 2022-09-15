Attribute VB_Name = "Test_StrArrayDiff_Word"
Option Explicit

'
' Copyright (c) 2022 Koki Takeyama
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included
' in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'

'
' --- Test ---
'

Sub Test_StrArrayDiff_Char_1()
    Test_StrArrayDiff_Char_Core _
        "The quick brown fox jumps over the lazy dog", _
        "The quick brown fox jumped over the lazy dogs"
End Sub

Sub Test_StrArrayDiff_Word_1()
    Test_StrArrayDiff_Word_Core _
        "The quick brown fox jumps over the lazy dog", _
        "The quick brown fox jumped over the lazy dogs"
End Sub

'
' --- Test Core ---
'

Sub Test_StrArrayDiff_Char_Core(Str1 As String, Str2 As String)
    Dim Len1 As Long
    Dim Len2 As Long
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    Dim StrArray1() As String
    Dim StrArray2() As String
    If Len1 > 0 Then
        ReDim StrArray1(0 To Len1 - 1)
    End If
    If Len2 > 0 Then
        ReDim StrArray2(0 To Len2 - 1)
    End If
    
    Dim Index1 As Long
    For Index1 = 0 To Len1 - 1
        StrArray1(Index1) = Mid(Str1, Index1 + 1, 1)
    Next
    
    Dim Index2 As Long
    For Index2 = 0 To Len2 - 1
        StrArray2(Index2) = Mid(Str2, Index2 + 1, 1)
    Next
    
    Debug_Print "=========="
    Debug_Print "Str1: " & Str1
    Debug_Print "Str2: " & Str2
    Debug_Print "ED: " & CStr(EditDistance(StrArray1, StrArray2))
    Debug_Print "LCS: " & LongestCommonSubsequence(StrArray1, StrArray2)
    
    Dim SES As String
    SES = ShortestEditScript(StrArray1, StrArray2)
    Debug_Print "SES: " & SES
    
    Dim SES1 As String
    Dim SES2 As String
    Dim SES3 As String
    Index1 = 1
    Index2 = 1
    Dim Index3 As Long
    For Index3 = 1 To Len(SES)
        Select Case Mid(SES, Index3, 1)
        Case "-"
           SES1 = SES1 & "-" & Mid(Str1, Index1, 1)
           SES3 = SES3 & "-" & Mid(Str1, Index1, 1)
           Index1 = Index1 + 1
        Case "+"
           SES2 = SES2 & "+" & Mid(Str2, Index2, 1)
           SES3 = SES3 & "+" & Mid(Str2, Index2, 1)
           Index2 = Index2 + 1
        Case " "
           SES1 = SES1 & " " & Mid(Str1, Index1, 1)
           SES2 = SES2 & " " & Mid(Str2, Index2, 1)
           SES3 = SES3 & " " & Mid(Str1, Index1, 1)
           Index1 = Index1 + 1
           Index2 = Index2 + 1
        End Select
    Next
    Debug_Print "SES1: " & SES1
    Debug_Print "SES2: " & SES2
    Debug_Print "SES3: " & SES3
End Sub

Sub Test_StrArrayDiff_Word_Core(Str1 As String, Str2 As String)
    Dim Str1Words As Object
    Dim Str2Words As Object
    Set Str1Words = RegExp_Execute(Str1, "(\w+)\W*", False, True, False)
    Set Str2Words = RegExp_Execute(Str2, "(\w+)\W*", False, True, False)
    
    Dim Len1 As Long
    Dim Len2 As Long
    Len1 = Str1Words.Count
    Len2 = Str2Words.Count
    
    Dim StrArray1() As String
    Dim StrArray2() As String
    If Len1 > 0 Then
        ReDim StrArray1(0 To Len1 - 1)
    End If
    If Len2 > 0 Then
        ReDim StrArray2(0 To Len2 - 1)
    End If
    
    Dim Index1 As Long
    For Index1 = 0 To Len1 - 1
        StrArray1(Index1) = Str1Words.Item(Index1).SubMatches.Item(0) & " "
    Next
    
    Dim Index2 As Long
    For Index2 = 0 To Len2 - 1
        StrArray2(Index2) = Str2Words.Item(Index2).SubMatches.Item(0) & " "
    Next
    
    Debug_Print "=========="
    Debug_Print "Str1: " & Str1
    Debug_Print "Str2: " & Str2
    Debug_Print "ED: " & CStr(EditDistance(StrArray1, StrArray2))
    Debug_Print "LCS: " & LongestCommonSubsequence(StrArray1, StrArray2)
    
    Dim SES As String
    SES = ShortestEditScript(StrArray1, StrArray2)
    Debug_Print "SES: " & SES
    
    Dim SES1 As String
    Dim SES2 As String
    Dim SES3 As String
    Index1 = 1
    Index2 = 1
    Dim Index3 As Long
    For Index3 = 1 To Len(SES)
        Select Case Mid(SES, Index3, 1)
        Case "-"
           SES1 = SES1 & "-" & Str1Words.Item(Index1 - 1)
           SES3 = SES3 & "-" & Str1Words.Item(Index1 - 1)
           Index1 = Index1 + 1
        Case "+"
           SES2 = SES2 & "+" & Str2Words.Item(Index2 - 1)
           SES3 = SES3 & "+" & Str2Words.Item(Index2 - 1)
           Index2 = Index2 + 1
        Case " "
           SES1 = SES1 & " " & Str1Words.Item(Index1 - 1)
           SES2 = SES2 & " " & Str2Words.Item(Index2 - 1)
           SES3 = SES3 & " " & Str1Words.Item(Index1 - 1)
           Index1 = Index1 + 1
           Index2 = Index2 + 1
        End Select
    Next
    Debug_Print "SES1: " & SES1
    Debug_Print "SES2: " & SES2
    Debug_Print "SES3: " & SES3
End Sub

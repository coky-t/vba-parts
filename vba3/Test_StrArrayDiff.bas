Attribute VB_Name = "Test_StrArrayDiff"
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

Sub Test_StrArrayDiff_1()
    Test_StrArrayDiff_Core "abcdef", "dacfea"
End Sub

Sub Test_StrArrayDiff_2()
    Test_StrArrayDiff_Core "kitten", "sitting"
End Sub

Sub Test_StrArrayDiff_3()
    Test_StrArrayDiff_Core "", "dacfea"
End Sub

Sub Test_StrArrayDiff_4()
    Test_StrArrayDiff_Core "abcdef", ""
End Sub

Sub Test_StrArrayDiff_5()
    Test_StrArrayDiff_Core "", ""
End Sub

'
' --- Test Core ---
'

Sub Test_StrArrayDiff_Core(Str1, Str2)
    Dim Len1
    Dim Len2
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    Dim StrArray1()
    Dim StrArray2()
    If Len1 > 0 Then
        ReDim StrArray1(Len1 - 1)
    End If
    If Len2 > 0 Then
        ReDim StrArray2(Len2 - 1)
    End If
    
    Dim Index1
    For Index1 = 0 To Len1 - 1
        StrArray1(Index1) = Mid(Str1, Index1 + 1, 1)
    Next
    
    Dim Index2
    For Index2 = 0 To Len2 - 1
        StrArray2(Index2) = Mid(Str2, Index2 + 1, 1)
    Next
    
    Debug_Print "=========="
    Debug_Print "Str1: " & Str1
    Debug_Print "Str2: " & Str2
    Debug_Print "ED: " & CStr(EditDistance(StrArray1, StrArray2))
    Debug_Print "LCS: " & LongestCommonSubsequence(StrArray1, StrArray2)
    
    Dim SES
    SES = ShortestEditScript(StrArray1, StrArray2)
    Debug_Print "SES: " & SES
    
    Dim SES1
    Dim SES2
    Dim SES3
    Index1 = 1
    Index2 = 1
    Dim Index3
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

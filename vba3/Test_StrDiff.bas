Attribute VB_Name = "Test_StrDiff"
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

Sub Test_StrDiff_1()
    Test_StrDiff_Core "abcdef", "dacfea"
End Sub

Sub Test_StrDiff_2()
    Test_StrDiff_Core "kitten", "sitting"
End Sub

Sub Test_StrDiff_3()
    Test_StrDiff_Core "", "dacfea"
End Sub

Sub Test_StrDiff_4()
    Test_StrDiff_Core "abcdef", ""
End Sub

Sub Test_StrDiff_5()
    Test_StrDiff_Core "", ""
End Sub

'
' --- Test Core ---
'

Sub Test_StrDiff_Core(Str1, Str2)
    Debug_Print "=========="
    Debug_Print "Str1: " & Str1
    Debug_Print "Str2: " & Str2
    Debug_Print "ED: " & CStr(EditDistance(Str1, Str2))
    Debug_Print "LCS: " & LongestCommonSubsequence(Str1, Str2)
    Debug_Print "SES: " & ShortestEditScript(Str1, Str2)
End Sub

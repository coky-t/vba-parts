Attribute VB_Name = "MDictionary"
Option Explicit

'
' Copyright (c) 2024 Koki Takeyama
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
' Microsoft Scripting Runtime
' - Scripting.Dictionary
'
' Dictionary object
' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object
'

Public Function StrArray_Unique(StrArray As Variant, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Variant
    If Not IsArray(StrArray) Then
        StrArray_Unique = StrArray
        Exit Function
    End If
    
    Dim StrDic As Object
    Set StrDic = CreateObject("Scripting.Dictionary")
    StrDic.CompareMode = CompareMode
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(StrArray)
    UB = UBound(StrArray)
    
    Dim Index As Long
    For Index = LB To UB
        If Not StrDic.Exists(StrArray(Index)) Then
            StrDic.Add StrArray(Index), StrArray(Index)
        End If
    Next
    
    StrArray_Unique = StrDic.Keys
End Function

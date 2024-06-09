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
    
    Dim StrDic As Scripting.Dictionary
    Set StrDic = New Scripting.Dictionary
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

Public Function StrMatrix_Unique(StrMatrix As Variant, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Variant
    If Not IsArray(StrMatrix) Then
        StrMatrix_Unique = StrMatrix
        Exit Function
    End If
    
    Dim StrDic As Scripting.Dictionary
    Set StrDic = New Scripting.Dictionary
    StrDic.CompareMode = CompareMode
    
    ' Step1. StrMatrix to StrDic
    
    Dim LB1 As Long
    Dim UB1 As Long
    Dim LB2 As Long
    Dim UB2 As Long
    LB1 = LBound(StrMatrix, 1)
    UB1 = UBound(StrMatrix, 1)
    LB2 = LBound(StrMatrix, 2)
    UB2 = UBound(StrMatrix, 2)
    
    Dim StrArray() As String
    ReDim StrArray(LB2 To UB2)
    
    Dim Index1 As Long
    For Index1 = LB1 To UB1
        Dim Index2 As Long
        For Index2 = LB2 To UB2
            StrArray(Index2) = StrMatrix(Index1, Index2)
        Next
        
        Dim StrTemp As String
        StrTemp = Join(StrArray, vbTab)
        If Not StrDic.Exists(StrTemp) Then
            StrDic.Add StrTemp, Index1
        End If
    Next
    
    ' Strp2. StrMatrix to StrMatrixNew
    
    Dim Count As Long
    Count = StrDic.Count
    
    Dim LB1New As Long
    Dim UB1New As Long
    LB1New = LB1
    UB1New = LB1 + Count - 1
    
    Dim StrMatrixNew() As String
    ReDim StrMatrixNew(LB1New To UB1New, LB2 To UB2)
    
    Dim Items As Variant
    Items = StrDic.Items
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(Items)
    UB = UBound(Items)
    
    Dim Index As Long
    For Index = LB To UB
        Index1 = CLng(Items(Index))
        Dim Index1New As Long
        Index1New = LB1 + Index - LB
        For Index2 = LB2 To UB2
            StrMatrixNew(Index1New, Index2) = StrMatrix(Index1, Index2)
        Next
    Next
    
    StrMatrix_Unique = StrMatrixNew
End Function

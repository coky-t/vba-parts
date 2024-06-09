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

Public Function StrArray_Unique(StrArray, CompareMode)
    If Not IsArray(StrArray) Then
        StrArray_Unique = StrArray
        Exit Function
    End If
    
    Dim StrDic
    Set StrDic = CreateObject("Scripting.Dictionary")
    StrDic.CompareMode = CompareMode
    
    Dim LB
    Dim UB
    LB = LBound(StrArray)
    UB = UBound(StrArray)
    
    Dim Index
    For Index = LB To UB
        If Not StrDic.Exists(StrArray(Index)) Then
            StrDic.Add StrArray(Index), StrArray(Index)
        End If
    Next
    
    StrArray_Unique = StrDic.Keys
End Function

Public Function StrMatrix_Unique(StrMatrix, CompareMode)
    If Not IsArray(StrMatrix) Then
        StrMatrix_Unique = StrMatrix
        Exit Function
    End If
    
    Dim StrDic
    Set StrDic = CreateObject("Scripting.Dictionary")
    StrDic.CompareMode = CompareMode
    
    ' Step1. StrMatrix to StrDic
    
    Dim LB1
    Dim UB1
    Dim LB2
    Dim UB2
    LB1 = LBound(StrMatrix, 1)
    UB1 = UBound(StrMatrix, 1)
    LB2 = LBound(StrMatrix, 2)
    UB2 = UBound(StrMatrix, 2)
    
    Dim StrArray()
    ReDim StrArray(0 To UB2 - LB2)
    
    Dim Index1
    For Index1 = LB1 To UB1
        Dim Index2
        For Index2 = LB2 To UB2
            StrArray(Index2 - LB2) = StrMatrix(Index1, Index2)
        Next
        
        Dim StrTemp
        StrTemp = Join(StrArray, vbTab)
        If Not StrDic.Exists(StrTemp) Then
            StrDic.Add StrTemp, Index1
        End If
    Next
    
    ' Strp2. StrMatrix to StrMatrixNew
    
    Dim Count
    Count = StrDic.Count
    
    Dim LB1New
    Dim UB1New
    LB1New = LB1
    UB1New = LB1 + Count - 1
    
    Dim StrMatrixNew()
    ReDim StrMatrixNew(0 To UB1New - LB1New, 0 To UB2 - LB2)
    
    Dim Items As Variant
    Items = StrDic.Items
    
    Dim LB
    Dim UB
    LB = LBound(Items)
    UB = UBound(Items)
    
    Dim Index
    For Index = LB To UB
        Index1 = CLng(Items(Index))
        Dim Index1New
        Index1New = LB1 + Index - LB
        For Index2 = LB2 To UB2
            StrMatrixNew(Index1New - LB1, Index2 - LB2) = StrMatrix(Index1, Index2)
        Next
    Next
    
    StrMatrix_Unique = StrMatrixNew
End Function

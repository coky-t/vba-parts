Attribute VB_Name = "Test_MDictionary"
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
' --- Test ---
'

Public Sub Test_StrArray_Unique()
    Test_StrArray_Unique_Core Array("aaa", "bbb", "ccc", "ddd", "aaa")
End Sub

Public Sub Test_StrMatrix_Unique()
    Dim StrMatrix
    ReDim StrMatrix(0 To 4, 0 To 1)
    
    StrMatrix(0, 0) = "aaa"
    StrMatrix(0, 1) = "aaa"
    StrMatrix(1, 0) = "aaa"
    StrMatrix(1, 1) = "bbb"
    StrMatrix(2, 0) = "bbb"
    StrMatrix(2, 1) = "aaa"
    StrMatrix(3, 0) = "ccc"
    StrMatrix(3, 1) = "ccc"
    StrMatrix(4, 0) = "aaa"
    StrMatrix(4, 1) = "aaa"
    
    Test_StrMatrix_Unique_Core StrMatrix
End Sub

'
' --- Test Core ---
'

Private Sub Test_StrArray_Unique_Core(StrArray)
    Debug_Print "---"
    Debug_Print "Input: " & Join(StrArray, ", ")
    Debug_Print "Output: " & Join(StrArray_Unique(StrArray, vbBinaryCompare), ", ")
End Sub

Private Sub Test_StrMatrix_Unique_Core(StrMatrix)
    Debug_Print "---"
    Debug_Print "Input: "
    Debug_Print_StrMatrix StrMatrix
    Debug_Print "Output: "
    Debug_Print_StrMatrix StrMatrix_Unique(StrMatrix, vbBinaryCompare)
End Sub

Private Sub Debug_Print_StrMatrix(StrMatrix)
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
        Debug_Print Join(StrArray, ", ")
    Next
End Sub

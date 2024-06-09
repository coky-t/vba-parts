Attribute VB_Name = "X_StrMatrix"
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

Public Sub Selection_Unique(Selection_ As Object)
    Select Case TypeName(Selection_)
    Case "Range"
        Range_Unique Selection_
    Case Else
        ' nop
    End Select
End Sub

Public Sub Range_Unique(Cells As Range)
    If Cells Is Nothing Then Exit Sub
    
    Dim CurrentSheet As Worksheet
    Set CurrentSheet = Cells.Parent
    
    Dim Items
    Items = Cells.Value
    
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(Items, 1)
    UB1 = UBound(Items, 1)
    
    If LB1 = UB1 Then Exit Sub
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(Items, 2)
    UB2 = UBound(Items, 2)
    
    Dim ItemsNew
    ItemsNew = StrMatrix_Unique(Items)
    
    Dim LB1New As Long
    Dim UB1New As Long
    LB1New = LBound(ItemsNew, 1)
    UB1New = UBound(ItemsNew, 1)
    
    Dim TopLeftCell As Range
    Set TopLeftCell = Cells.Item(1, 1)
    
    Dim RightBottomCellNew As Range
    Set RightBottomCellNew = TopLeftCell.Offset(UB1New - LB1New, UB2 - LB2)
    
    Dim NewRange As Range
    Set NewRange = CurrentSheet.Range(TopLeftCell, RightBottomCellNew)
    
    Cells.Clear
    NewRange.Value = ItemsNew
End Sub

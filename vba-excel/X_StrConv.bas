Attribute VB_Name = "X_StrConv"
Option Explicit

'
' Copyright (c) 2021 Koki Takeyama
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

Public Sub Selection_StrConv( _
    Selection_ As Object, _
    conversion As Integer, _
    Optional bShapes As Boolean)
    
    Select Case TypeName(Selection_)
    Case "Workbook"
        Book_StrConv Selection_, conversion, bShapes
    Case "Worksheet"
        Sheet_StrConv Selection_, conversion, bShapes
    Case "Range"
        Range_StrConv Selection_, conversion
    Case "Rectangle", "DrawingObjects"
        ShapeRange_StrConv Selection_.ShapeRange, conversion
    Case Else
        ' nop
    End Select
End Sub

Public Sub Book_StrConv( _
    Selection_ As Object, _
    conversion As Integer, _
    Optional bShapes As Boolean)
    
    If Book Is Nothing Then Exit Sub
    
    Dim Sheet As Worksheet
    For Each Sheet In Book.Worksheets
        Sheet_StrConv Sheet, conversion, bShapes
    Next
End Sub

Public Sub Sheet_StrConv( _
    Selection_ As Object, _
    conversion As Integer, _
    Optional bShapes As Boolean)
    
    ' Cells
    
    Dim Cells As Range
    
    With Sheet
        Set Cells = _
            .Range(.Cells(1, 1), .Cells.SpecialCells(xlCellTypeLastCell))
    End With
    
    Range_StrConv Cells, conversion
    
    ' Shapes
    
    If bShapes Then
        Dim Shape_ As Shape
        For Each Shape_ In Sheet.Shapes
            Shape_StrConv Shape_, conversion
        Next
    End If
End Sub

Public Sub Range_StrConv(Cells As Range, conversion As Integer)
    If Cells Is Nothing Then Exit Sub
    
    Dim Sheet As Worksheet
    Set Sheet = Cells.Parent
    
    Dim LastRow As Long
    Dim LastCol As Long
    With Sheet.Cells.SpecialCells(xlCellTypeLastCell)
        LastRow = .Row
        LastCol = .Column
    End With
    
    Dim Cell As Range
    For Each Cell In Cells
        If Cell.Row <= LastRow Then
        If Cell.Column <= LastCol Then
            Cell_StrConv Cell, conversion
        End If
        End If
    Next
End Sub

Public Sub Cell_StrConv(Cell As Range, conversion As Integer)
    If Cell Is Nothing Then Exit Sub
    If IsError(Cell.Value) Then Exit Sub
    If Cell.Value = "" Then Exit Sub
    
    On Error Resume Next
    Cell.Value = StrConv(Cell.Value, conversion)
    On Error GoTo 0
End Sub

Public Sub ShapeRange_StrConv(ShapeRange_ As ShapeRange, conversion As Integer)
    If ShapeRange_ Is Nothing Then Exit Sub
    
    Dim Shape_ As Shape
    For Each Shape_ In ShapeRange_
        Shape_StrConv Shape_, conversion
    Next
End Sub

Public Sub ShapeGroup_StrConv(ShapeGroup As Shape, conversion As Integer)
    If ShapeGroup Is Nothing Then Exit Sub
    
    Dim Shape_ As Shape
    For Each Shape_ In ShapeGroup.GroupItems
        Shape_StrConv Shape_, conversion
    Next
End Sub

Public Sub ShapeCanvas_StrConv(ShapeCanvas As Shape, conversion As Integer)
    If ShapeCanvas Is Nothing Then Exit Sub
    
    Dim Shape_ As Shape
    For Each Shape_ In ShapeCanvas.CanvasItems
        Shape_StrConv Shape_, conversion
    Next
End Sub

Public Sub Shape_StrConv(Shape_ As Shape, conversion As Integer)
    If Shape_ Is Nothing Then Exit Sub
    
    Dim Text As String
    Text = Shape_GetText(Shape_)
    If Not Text = "" Then
        Shape_LetText Shape_, StrConv(Text, conversion)
    End If
    
    Const msoGroup As Long = 6
    Const msoCanvas As Long = 20
    
    Select Case Shape_.Type
    Case msoGroup
        ShapeGroup_StrConv Shape_, conversion
    Case msoCanvas
        ShapeCanvas_StrConv Shape_, conversion
    End Select
End Sub

Public Function Shape_GetText(Shape_ As Shape) As String
    If Shape_ Is Nothing Then Exit Function
    
    On Error Resume Next
    Shape_GetText = Shape_.TextFrame.Characters.Text
    On Error GoTo 0
End Function

Public Sub Shape_LetText(Shape_ As Shape, Text As String)
    If Shape_ Is Nothing Then Exit Sub
    
    On Error Resume Next
    Shape_.TextFrame.Characters.Text = Text
    On Error GoTo 0
End Sub

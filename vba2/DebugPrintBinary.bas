Attribute VB_Name = "DebugPrintBinary"
Option Explicit

'
' Copyright (c) 2020 Koki Takeyama
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

Public Sub Debug_Print_Binary(Binary() As Byte)
    Dim Text As String
    Dim Index1 As Long
    Dim Index2 As Long
    For Index1 = LBound(Binary) To UBound(Binary) Step 16
        For Index2 = Index1 To MinL(Index1 + 15, UBound(Binary))
            Text = Text & Right("0" & Hex(Binary(Index2)), 2) & " "
        Next
        Text = Text & vbNewLine
    Next
    
    Debug_Print "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
    Debug_Print Text
    Debug_Print "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
End Sub

Private Function MinL(Value1 As Long, Value2 As Long) As Long
    If Value1 < Value2 Then
        MinL = Value1
    Else
        MinL = Value2
    End If
End Function

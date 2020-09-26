Attribute VB_Name = "Test_MXMLDOM"
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

'
' --- Test ---
'

Public Sub Test_Base64()
    Test_Base64_Core GetTestBinary()
End Sub

Public Sub Test_Hex()
    Test_Hex_Core GetTestBinary()
End Sub

'
' --- Test Core ---
'

Public Sub Test_Base64_Core(Binary() As Byte)
    Dim Base64Text As String
    Base64Text = GetBase64TextFromBinary(Binary)
    Debug_Print Base64Text
    
    Binary = GetBinaryFromBase64Text(Base64Text)
    Debug_Print_Binary Binary
End Sub

Public Sub Test_Hex_Core(Binary() As Byte)
    Dim HexText As String
    HexText = GetHexTextFromBinary(Binary)
    Debug_Print HexText
    
    Binary = GetBinaryFromHexText(HexText)
    Debug_Print_Binary Binary
End Sub

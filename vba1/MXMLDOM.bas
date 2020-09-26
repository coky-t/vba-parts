Attribute VB_Name = "MXMLDOM"
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
' Microsoft XML, vX.X
' - MSXML2.DOMDocument
'

'
' --- MSXML2.DOMDocument ---
'

'
' GetXMLDOM
' - Returns a MSXML2.DOMDocument object.
'

Public Function GetXMLDOM() As MSXML2.DOMDocument60
    Static XMLDOM As MSXML2.DOMDocument60
    If XMLDOM Is Nothing Then
        Set XMLDOM = New MSXML2.DOMDocument60
    End If
    Set GetXMLDOM = XMLDOM
End Function

'
' GetBinBase64
' - Returns the IXMLDOMElement object with bin.base64 datatype.
'

Public Function GetBinBase64() As MSXML2.IXMLDOMElement
    Static BinBase64 As MSXML2.IXMLDOMElement
    If BinBase64 Is Nothing Then
        Set BinBase64 = GetXMLDOM().createElement("tmp")
        BinBase64.DataType = "bin.base64"
    End If
    Set GetBinBase64 = BinBase64
End Function


'
' GetBinHex
' - Returns the IXMLDOMElement object with bin.hex datatype.
'

Public Function GetBinHex() As MSXML2.IXMLDOMElement
    Static BinHex As MSXML2.IXMLDOMElement
    If BinHex Is Nothing Then
        Set BinHex = GetXMLDOM().createElement("tmp")
        BinHex.DataType = "bin.hex"
    End If
    Set GetBinHex = BinHex
End Function


'
' --- Base64 ---
'

'
' GetBase64TextFromBinary
' - Return the base64-encoded data.
'

'
' Binary:
'   Required. A Variant that contains an array of bytes.
'

Public Function GetBase64TextFromBinary(Binary() As Byte) As String
    On Error Resume Next
    
    With GetBinBase64()
        .nodeTypedValue = Binary
        GetBase64TextFromBinary = .Text
    End With
End Function

'
' GetBinaryFromBase64Text
' - Return the resulting data.
'

'
' Base64Text:
'   Required. A String that contains a base64-encoded data.
'

Public Function GetBinaryFromBase64Text(Base64Text As String) As Variant
    On Error Resume Next
    
    With GetBinBase64()
        .Text = Base64Text
        GetBinaryFromBase64Text = .nodeTypedValue
    End With
End Function

'
' --- Hex ---
'

'
' GetHexTextFromBinary
' - Return the hex-text data.
'

'
' Binary:
'   Required. A Variant that contains an array of bytes.
'

Public Function GetHexTextFromBinary(Binary() As Byte) As String
    On Error Resume Next
    
    With GetBinHex()
        .nodeTypedValue = Binary
        GetHexTextFromBinary = .Text
    End With
End Function

'
' GetBinaryFromHexText
' - Return the resulting data.
'

'
' HexText:
'   Required. A String that contains a hex-text data.
'

Public Function GetBinaryFromHexText(HexText As String) As Variant
    On Error Resume Next
    
    With GetBinHex()
        .Text = HexText
        GetBinaryFromHexText = .nodeTypedValue
    End With
End Function

Attribute VB_Name = "MsgPack_Common"
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

''
'' MessagePack for VBA - Common
''

''
'' MessagePack for VBA - Common - Serialization Helper
''

Public Function GetBytesHelper1( _
    FormatValue As Byte, SrcBytes() As Byte) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(SrcBytes)
    UB = UBound(SrcBytes)
    
    Dim Length As Long
    Length = UB - LB + 1
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To Length)
    Bytes(0) = FormatValue
    
    BitConverter.CopyBytes Bytes, 1, SrcBytes, LB, Length
    
    GetBytesHelper1 = Bytes
End Function

Public Function GetBytesHelper2A( _
    FormatValue As Byte, SrcByte1 As Byte, SrcBytes2) As Byte()
    'FormatValue As Byte, SrcByte1 As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(SrcBytes2)
    UB = UBound(SrcBytes2)
    
    Dim Length As Long
    Length = UB - LB + 1
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To 1 + Length)
    Bytes(0) = FormatValue
    Bytes(1) = SrcByte1
    
    BitConverter.CopyBytes Bytes, 2, SrcBytes2, LB, Length
    
    GetBytesHelper2A = Bytes
End Function

Public Function GetBytesHelper2B( _
    FormatValue As Byte, SrcBytes1, SrcBytes2) As Byte()
    'FormatValue As Byte, SrcBytes1() As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(SrcBytes1)
    UB1 = UBound(SrcBytes1)
    
    Dim Length1 As Long
    Length1 = UB1 - LB1 + 1
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(SrcBytes2)
    UB2 = UBound(SrcBytes2)
    
    Dim Length2 As Long
    Length2 = UB2 - LB2 + 1
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To Length1 + Length2)
    Bytes(0) = FormatValue
    
    BitConverter.CopyBytes Bytes, 1, SrcBytes1, LB1, Length1
    BitConverter.CopyBytes Bytes, 1 + Length1, SrcBytes2, LB2, Length2
    
    GetBytesHelper2B = Bytes
End Function

Public Function GetBytesHelper3A(FormatValue As Byte, _
    SrcByte1 As Byte, SrcByte2 As Byte, SrcBytes3) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(SrcBytes3)
    UB = UBound(SrcBytes3)
    
    Dim Length As Long
    Length = UB - LB + 1
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To 1 + 1 + Length)
    Bytes(0) = FormatValue
    Bytes(1) = SrcByte1
    Bytes(2) = SrcByte2
    
    BitConverter.CopyBytes Bytes, 3, SrcBytes3, LB, Length
    
    GetBytesHelper3A = Bytes
End Function

Public Function GetBytesHelper3B(FormatValue As Byte, _
    SrcBytes1, SrcByte2 As Byte, SrcBytes3) As Byte()
    
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(SrcBytes1)
    UB1 = UBound(SrcBytes1)
    
    Dim Length1 As Long
    Length1 = UB1 - LB1 + 1
    
    Dim LB3 As Long
    Dim UB3 As Long
    LB3 = LBound(SrcBytes3)
    UB3 = UBound(SrcBytes3)
    
    Dim Length3 As Long
    Length3 = UB3 - LB3 + 1
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To Length1 + 1 + Length3)
    Bytes(0) = FormatValue
    
    BitConverter.CopyBytes Bytes, 1, SrcBytes1, LB1, Length1
    
    Bytes(Length1 + 1) = SrcByte2
    
    BitConverter.CopyBytes Bytes, 1 + Length1 + 1, SrcBytes3, LB3, Length3
    
    GetBytesHelper3B = Bytes
End Function

''
'' MessagePack for VBA - Common - Byte Array Helper
''

Public Function AddBytes(DstBytes() As Byte, SrcBytes() As Byte) As Byte()
    Dim DstLB As Long
    Dim DstUB As Long
    DstLB = LBound(DstBytes)
    DstUB = UBound(DstBytes)
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    Dim SrcLength As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    SrcLength = SrcUB - SrcLB + 1
    
    ReDim Preserve DstBytes(DstLB To DstUB + SrcLength)
    BitConverter.CopyBytes DstBytes, DstUB + 1, SrcBytes, SrcLB, SrcLength
    
    AddBytes = DstBytes
End Function

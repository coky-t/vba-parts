Attribute VB_Name = "ByteArray"
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

Public Function GetByteArrayLEFromInteger(Value As Integer) As Variant
    Dim ByteArrayLE(1) As Byte
    ByteArrayLE(0) = Value And &HFF
    ByteArrayLE(1) = RightShiftInteger(Value, 8) And &HFF
    GetByteArrayLEFromInteger = ByteArrayLE
End Function

Public Function GetByteArrayBEFromInteger(Value As Integer) As Variant
    Dim ByteArrayBE(1) As Byte
    ByteArrayBE(0) = RightShiftInteger(Value, 8) And &HFF
    ByteArrayBE(1) = Value And &HFF
    GetByteArrayBEFromInteger = ByteArrayBE
End Function

Public Function GetByteArrayLEFromLong(Value As Long) As Variant
    Dim ByteArrayLE(3) As Byte
    ByteArrayLE(0) = Value And &HFF
    ByteArrayLE(1) = RightShiftLong(Value, 8) And &HFF
    ByteArrayLE(2) = RightShiftLong(Value, 16) And &HFF
    ByteArrayLE(3) = RightShiftLong(Value, 24) And &HFF
    GetByteArrayLEFromLong = ByteArrayLE
End Function

Public Function GetByteArrayBEFromLong(Value As Long) As Variant
    Dim ByteArrayBE(3) As Byte
    ByteArrayBE(0) = RightShiftLong(Value, 24) And &HFF
    ByteArrayBE(1) = RightShiftLong(Value, 16) And &HFF
    ByteArrayBE(2) = RightShiftLong(Value, 8) And &HFF
    ByteArrayBE(3) = Value And &HFF
    GetByteArrayBEFromLong = ByteArrayBE
End Function

#If Win64 Then
Public Function GetByteArrayLEFromLongLong(Value As LongLong) As Variant
    Dim ByteArrayLE(7) As Byte
    ByteArrayLE(0) = CByte(Value And &HFF^)
    ByteArrayLE(1) = CByte(RightShiftLongLong(Value, 8) And &HFF^)
    ByteArrayLE(2) = CByte(RightShiftLongLong(Value, 16) And &HFF^)
    ByteArrayLE(3) = CByte(RightShiftLongLong(Value, 24) And &HFF^)
    ByteArrayLE(4) = CByte(RightShiftLongLong(Value, 32) And &HFF^)
    ByteArrayLE(5) = CByte(RightShiftLongLong(Value, 40) And &HFF^)
    ByteArrayLE(6) = CByte(RightShiftLongLong(Value, 48) And &HFF^)
    ByteArrayLE(7) = CByte(RightShiftLongLong(Value, 56) And &HFF^)
    GetByteArrayLEFromLongLong = ByteArrayLE
End Function

Public Function GetByteArrayBEFromLongLong(Value As LongLong) As Variant
    Dim ByteArrayBE(7) As Byte
    ByteArrayBE(0) = CByte(RightShiftLongLong(Value, 56) And &HFF^)
    ByteArrayBE(1) = CByte(RightShiftLongLong(Value, 48) And &HFF^)
    ByteArrayBE(2) = CByte(RightShiftLongLong(Value, 40) And &HFF^)
    ByteArrayBE(3) = CByte(RightShiftLongLong(Value, 32) And &HFF^)
    ByteArrayBE(4) = CByte(RightShiftLongLong(Value, 24) And &HFF^)
    ByteArrayBE(5) = CByte(RightShiftLongLong(Value, 16) And &HFF^)
    ByteArrayBE(6) = CByte(RightShiftLongLong(Value, 8) And &HFF^)
    ByteArrayBE(7) = CByte(Value And &HFF^)
    GetByteArrayBEFromLongLong = ByteArrayBE
End Function
#End If

Public Function GetIntegerFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As Integer
    
    GetIntegerFromByteArrayLE = LE(Pos) Or LeftShiftInteger(LE(Pos + 1), 8)
End Function

Public Function GetIntegerFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As Integer
    
    GetIntegerFromByteArrayBE = BE(Pos + 1) Or LeftShiftInteger(BE(Pos), 8)
End Function

Public Function GetLongFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As Long
    
    GetLongFromByteArrayLE = LE(Pos) Or _
        LeftShiftLong(LE(Pos + 1), 8) Or _
        LeftShiftLong(LE(Pos + 2), 16) Or _
        LeftShiftLong(LE(Pos + 3), 24)
End Function

Public Function GetLongFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As Long
    
    GetLongFromByteArrayBE = BE(Pos + 3) Or _
        LeftShiftLong(BE(Pos + 2), 8) Or _
        LeftShiftLong(BE(Pos + 1), 16) Or _
        LeftShiftLong(BE(Pos), 24)
End Function

#If Win64 Then
Public Function GetLongLongFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As LongLong
    
    GetLongLongFromByteArrayLE = LE(Pos) Or _
        LeftShiftLongLong(LE(Pos + 1), 8) Or _
        LeftShiftLongLong(LE(Pos + 2), 16) Or _
        LeftShiftLongLong(LE(Pos + 3), 24) Or _
        LeftShiftLongLong(LE(Pos + 4), 32) Or _
        LeftShiftLongLong(LE(Pos + 5), 40) Or _
        LeftShiftLongLong(LE(Pos + 6), 48) Or _
        LeftShiftLongLong(LE(Pos + 7), 56)
End Function

Public Function GetLongLongFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As LongLong
    
    GetLongLongFromByteArrayBE = BE(Pos + 7) Or _
        LeftShiftLongLong(BE(Pos + 6), 8) Or _
        LeftShiftLongLong(BE(Pos + 5), 16) Or _
        LeftShiftLongLong(BE(Pos + 4), 24) Or _
        LeftShiftLongLong(BE(Pos + 3), 32) Or _
        LeftShiftLongLong(BE(Pos + 2), 40) Or _
        LeftShiftLongLong(BE(Pos + 1), 48) Or _
        LeftShiftLongLong(BE(Pos), 56)
End Function
#End If

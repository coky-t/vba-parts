Attribute VB_Name = "ByteArrayX"
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

Private Type IntegerType
    Value As Integer
End Type

Private Type LongType
    Value As Long
End Type

#If Win64 Then
Private Type LongLongType
    Value As LongLong
End Type
#End If

Private Type SingleType
    Value As Single
End Type

Private Type DoubleType
    Value As Double
End Type

Private Type ByteArray2
    Items(1) As Byte
End Type

Private Type ByteArray4
    Items(3) As Byte
End Type

Private Type ByteArray8
    Items(7) As Byte
End Type

Public Function GetByteArrayLEFromInteger(Value As Integer) As Variant
    Dim I As IntegerType
    I.Value = Value
    
    Dim B2 As ByteArray2
    LSet B2 = I
    
    Dim ByteArrayLE(1) As Byte
    ByteArrayLE(0) = B2.Items(0)
    ByteArrayLE(1) = B2.Items(1)
    
    GetByteArrayLEFromInteger = ByteArrayLE
End Function

Public Function GetByteArrayBEFromInteger(Value As Integer) As Variant
    Dim I As IntegerType
    I.Value = Value
    
    Dim B2 As ByteArray2
    LSet B2 = I
    
    Dim ByteArrayBE(1) As Byte
    ByteArrayBE(0) = B2.Items(1)
    ByteArrayBE(1) = B2.Items(0)
    
    GetByteArrayBEFromInteger = ByteArrayBE
End Function

Public Function GetByteArrayLEFromLong(Value As Long) As Variant
    Dim L As LongType
    L.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = L
    
    Dim ByteArrayLE(3) As Byte
    ByteArrayLE(0) = B4.Items(0)
    ByteArrayLE(1) = B4.Items(1)
    ByteArrayLE(2) = B4.Items(2)
    ByteArrayLE(3) = B4.Items(3)
    
    GetByteArrayLEFromLong = ByteArrayLE
End Function

Public Function GetByteArrayBEFromLong(Value As Long) As Variant
    Dim L As LongType
    L.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = L
    
    Dim ByteArrayBE(3) As Byte
    ByteArrayBE(0) = B4.Items(3)
    ByteArrayBE(1) = B4.Items(2)
    ByteArrayBE(2) = B4.Items(1)
    ByteArrayBE(3) = B4.Items(0)
    
    GetByteArrayBEFromLong = ByteArrayBE
End Function

#If Win64 Then
Public Function GetByteArrayLEFromLongLong(Value As LongLong) As Variant
    Dim LL As LongLongType
    LL.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = LL
    
    Dim ByteArrayLE(7) As Byte
    ByteArrayLE(0) = B8.Items(0)
    ByteArrayLE(1) = B8.Items(1)
    ByteArrayLE(2) = B8.Items(2)
    ByteArrayLE(3) = B8.Items(3)
    ByteArrayLE(4) = B8.Items(4)
    ByteArrayLE(5) = B8.Items(5)
    ByteArrayLE(6) = B8.Items(6)
    ByteArrayLE(7) = B8.Items(7)
    
    GetByteArrayLEFromLongLong = ByteArrayLE
End Function

Public Function GetByteArrayBEFromLongLong(Value As LongLong) As Variant
    Dim LL As LongLongType
    LL.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = LL
    
    Dim ByteArrayBE(7) As Byte
    ByteArrayBE(0) = B8.Items(7)
    ByteArrayBE(1) = B8.Items(6)
    ByteArrayBE(2) = B8.Items(5)
    ByteArrayBE(3) = B8.Items(4)
    ByteArrayBE(4) = B8.Items(3)
    ByteArrayBE(5) = B8.Items(2)
    ByteArrayBE(6) = B8.Items(1)
    ByteArrayBE(7) = B8.Items(0)
    
    GetByteArrayBEFromLongLong = ByteArrayBE
End Function
#End If

Public Function GetByteArrayLEFromSingle(Value As Single) As Variant
    Dim S As SingleType
    S.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = S
    
    Dim ByteArrayLE(3) As Byte
    ByteArrayLE(0) = B4.Items(0)
    ByteArrayLE(1) = B4.Items(1)
    ByteArrayLE(2) = B4.Items(2)
    ByteArrayLE(3) = B4.Items(3)
    
    GetByteArrayLEFromSingle = ByteArrayLE
End Function

Public Function GetByteArrayBEFromSingle(Value As Single) As Variant
    Dim S As SingleType
    S.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = S
    
    Dim ByteArrayBE(3) As Byte
    ByteArrayBE(0) = B4.Items(3)
    ByteArrayBE(1) = B4.Items(2)
    ByteArrayBE(2) = B4.Items(1)
    ByteArrayBE(3) = B4.Items(0)
    
    GetByteArrayBEFromSingle = ByteArrayBE
End Function

Public Function GetByteArrayLEFromDouble(Value As Double) As Variant
    Dim D As DoubleType
    D.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = D
    
    Dim ByteArrayLE(7) As Byte
    ByteArrayLE(0) = B8.Items(0)
    ByteArrayLE(1) = B8.Items(1)
    ByteArrayLE(2) = B8.Items(2)
    ByteArrayLE(3) = B8.Items(3)
    ByteArrayLE(4) = B8.Items(4)
    ByteArrayLE(5) = B8.Items(5)
    ByteArrayLE(6) = B8.Items(6)
    ByteArrayLE(7) = B8.Items(7)
    
    GetByteArrayLEFromDouble = ByteArrayLE
End Function

Public Function GetByteArrayBEFromDouble(Value As Double) As Variant
    Dim D As DoubleType
    D.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = D
    
    Dim ByteArrayBE(7) As Byte
    ByteArrayBE(0) = B8.Items(7)
    ByteArrayBE(1) = B8.Items(6)
    ByteArrayBE(2) = B8.Items(5)
    ByteArrayBE(3) = B8.Items(4)
    ByteArrayBE(4) = B8.Items(3)
    ByteArrayBE(5) = B8.Items(2)
    ByteArrayBE(6) = B8.Items(1)
    ByteArrayBE(7) = B8.Items(0)
    
    GetByteArrayBEFromDouble = ByteArrayBE
End Function

Public Function GetIntegerFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As Integer
    
    Dim B2 As ByteArray2
    B2.Items(0) = LE(Pos)
    B2.Items(1) = LE(Pos + 1)
    
    Dim I As IntegerType
    LSet I = B2
    
    GetIntegerFromByteArrayLE = I.Value
End Function

Public Function GetIntegerFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As Integer
    
    Dim B2 As ByteArray2
    B2.Items(0) = BE(Pos + 1)
    B2.Items(1) = BE(Pos)
    
    Dim I As IntegerType
    LSet I = B2
    
    GetIntegerFromByteArrayBE = I.Value
End Function

Public Function GetLongFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As Long
    
    Dim B4 As ByteArray4
    B4.Items(0) = LE(Pos)
    B4.Items(1) = LE(Pos + 1)
    B4.Items(2) = LE(Pos + 2)
    B4.Items(3) = LE(Pos + 3)
    
    Dim L As LongType
    LSet L = B4
    
    GetLongFromByteArrayLE = L.Value
End Function

Public Function GetLongFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As Long
    
    Dim B4 As ByteArray4
    B4.Items(0) = BE(Pos + 3)
    B4.Items(1) = BE(Pos + 2)
    B4.Items(2) = BE(Pos + 1)
    B4.Items(3) = BE(Pos + 0)
    
    Dim L As LongType
    LSet L = B4
    
    GetLongFromByteArrayBE = L.Value
End Function

#If Win64 Then
Public Function GetLongLongFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As LongLong
    
    Dim B8 As ByteArray8
    B8.Items(0) = LE(Pos)
    B8.Items(1) = LE(Pos + 1)
    B8.Items(2) = LE(Pos + 2)
    B8.Items(3) = LE(Pos + 3)
    B8.Items(4) = LE(Pos + 4)
    B8.Items(5) = LE(Pos + 5)
    B8.Items(6) = LE(Pos + 6)
    B8.Items(7) = LE(Pos + 7)
    
    Dim LL As LongLongType
    LSet LL = B8
    
    GetLongLongFromByteArrayLE = LL.Value
End Function

Public Function GetLongLongFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As LongLong
    
    Dim B8 As ByteArray8
    B8.Items(0) = BE(Pos + 7)
    B8.Items(1) = BE(Pos + 6)
    B8.Items(2) = BE(Pos + 5)
    B8.Items(3) = BE(Pos + 4)
    B8.Items(4) = BE(Pos + 3)
    B8.Items(5) = BE(Pos + 2)
    B8.Items(6) = BE(Pos + 1)
    B8.Items(7) = BE(Pos)
    
    Dim LL As LongLongType
    LSet LL = B8
    
    GetLongLongFromByteArrayBE = LL.Value
End Function
#End If

Public Function GetSingleFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As Single
    
    Dim B4 As ByteArray4
    B4.Items(0) = LE(Pos)
    B4.Items(1) = LE(Pos + 1)
    B4.Items(2) = LE(Pos + 2)
    B4.Items(3) = LE(Pos + 3)
    
    Dim S As SingleType
    LSet S = B4
    
    GetSingleFromByteArrayLE = S.Value
End Function

Public Function GetSingleFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As Single
    
    Dim B4 As ByteArray4
    B4.Items(0) = BE(Pos + 3)
    B4.Items(1) = BE(Pos + 2)
    B4.Items(2) = BE(Pos + 1)
    B4.Items(3) = BE(Pos + 0)
    
    Dim S As SingleType
    LSet S = B4
    
    GetSingleFromByteArrayBE = S.Value
End Function

Public Function GetDoubleFromByteArrayLE( _
    LE() As Byte, Optional Pos As Long) As Double
    
    Dim B8 As ByteArray8
    B8.Items(0) = LE(Pos)
    B8.Items(1) = LE(Pos + 1)
    B8.Items(2) = LE(Pos + 2)
    B8.Items(3) = LE(Pos + 3)
    B8.Items(4) = LE(Pos + 4)
    B8.Items(5) = LE(Pos + 5)
    B8.Items(6) = LE(Pos + 6)
    B8.Items(7) = LE(Pos + 7)
    
    Dim D As DoubleType
    LSet D = B8
    
    GetDoubleFromByteArrayLE = D.Value
End Function

Public Function GetDoubleFromByteArrayBE( _
    BE() As Byte, Optional Pos As Long) As Double
    
    Dim B8 As ByteArray8
    B8.Items(0) = BE(Pos + 7)
    B8.Items(1) = BE(Pos + 6)
    B8.Items(2) = BE(Pos + 5)
    B8.Items(3) = BE(Pos + 4)
    B8.Items(4) = BE(Pos + 3)
    B8.Items(5) = BE(Pos + 2)
    B8.Items(6) = BE(Pos + 1)
    B8.Items(7) = BE(Pos)
    
    Dim D As DoubleType
    LSet D = B8
    
    GetDoubleFromByteArrayBE = D.Value
End Function

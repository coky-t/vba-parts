Attribute VB_Name = "CBOR_7_Float"
Option Explicit

'
' Copyright (c) 2022 Koki Takeyama
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
'' CBOR for VBA
''

'
' Conditional
'

' Integer
#If Win64 Then
#Const USE_LONGLONG = True
#End If

'
' Declare
'

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As LongPtr)
#Else
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As Long)
#End If

'
' Types
'

Private Type IntegerT
    Value As Integer
End Type

Private Type LongT
    Value As Long
End Type

#If Win64 And USE_LONGLONG Then
Private Type LongLongT
    Value As LongLong
End Type
#End If

Private Type SingleT
    Value As Single
End Type

Private Type DoubleT
    Value As Double
End Type

Private Type CurrencyT
    Value As Currency
End Type

Private Type DateT
    Value As Date
End Type

Private Type Bytes2T
    Bytes(0 To 1) As Byte
End Type

Private Type Bytes4T
    Bytes(0 To 3) As Byte
End Type

Private Type Bytes8T
    Bytes(0 To 7) As Byte
End Type

''
'' CBOR for VBA - Encoding
''

Public Function GetCborBytes(Value) As Byte()
    Select Case VarType(Value)
    
    ' 4
    Case vbSingle
        GetCborBytes = GetCborBytesFromSingle(Value)
        
    ' 5
    Case vbDouble
        GetCborBytes = GetCborBytesFromDouble(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 4. Single
'
Private Function GetCborBytesFromSingle(Value) As Byte()
    GetCborBytesFromSingle = GetCborBytesFromFloat32((Value))
End Function

'
' 5. Double
'
Private Function GetCborBytesFromDouble(Value) As Byte()
    GetCborBytesFromDouble = GetCborBytesFromFloat64((Value))
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 7: simple/float
'

' 0xf9 | half-precision float (two-byte IEEE 754)
'Private Function GetCborBytesFromFloat16(ByVal Value As Single) As Byte()
'    GetCborBytesFromFloat16 = _
'        GetCborBytes1(&HF9, GetBytesFromFloat16(Value, True))
'End Function

' 0xfa | single-precision float (four-byte IEEE 754)
Private Function GetCborBytesFromFloat32(ByVal Value As Single) As Byte()
    GetCborBytesFromFloat32 = _
        GetCborBytes1(&HFA, GetBytesFromFloat32(Value, True))
End Function

' 0xfb | double-precision float (eight-byte IEEE 754)
Private Function GetCborBytesFromFloat64(ByVal Value As Double) As Byte()
    GetCborBytesFromFloat64 = _
        GetCborBytes1(&HFB, GetBytesFromFloat64(Value, True))
End Function

''
'' CBOR for VBA - Encoding - Formatter
''

Private Function GetCborBytes0(HeaderValue As Byte) As Byte()
    Dim CborBytes(0) As Byte
    CborBytes(0) = HeaderValue
    GetCborBytes0 = CborBytes
End Function

Private Function GetCborBytes1( _
    HeaderValue As Byte, SrcBytes) As Byte()
    'HeaderValue As Byte, SrcBytes() As Byte) As Byte()
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    
    Dim SrcLen As Long
    SrcLen = SrcUB - SrcLB + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen)
    CborBytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes, SrcLB, SrcLen
    
    GetCborBytes1 = CborBytes
End Function

Private Function GetCborBytes2( _
    HeaderValue As Byte, SrcBytes1, SrcBytes2) As Byte()
    'HeaderValue As Byte, SrcBytes1() As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen1 + SrcLen2)
    CborBytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    CopyBytes CborBytes, 1 + SrcLen1, SrcBytes2, SrcLB2, SrcLen2
    
    GetCborBytes2 = CborBytes
End Function

Private Function GetCborBytes3(HeaderValue As Byte, _
    SrcBytes1, SrcBytes2, SrcBytes3) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim SrcLB3 As Long
    Dim SrcUB3 As Long
    SrcLB3 = LBound(SrcBytes3)
    SrcUB3 = UBound(SrcBytes3)
    
    Dim SrcLen3 As Long
    SrcLen3 = SrcUB3 - SrcLB3 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen1 + SrcLen2 + SrcLen3)
    Bytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    
    CopyBytes CborBytes, 1 + SrcLen1, SrcBytes2, SrcLB2, SrcLen2
    
    CopyBytes CborBytes, 1 + SrcLen1 + SrcLen2, SrcBytes3, SrcLB3, SrcLen3
    
    GetCborBytes3 = CborBytes
End Function

''
'' CBOR for VBA - Encoding - Bytes Operator
''

Private Sub AddBytes(DstBytes() As Byte, SrcBytes() As Byte)
    Dim DstLB As Long
    Dim DstUB As Long
    DstLB = LBound(DstBytes)
    DstUB = UBound(DstBytes)
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    Dim SrcLen As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    SrcLen = SrcUB - SrcLB + 1
    
    ReDim Preserve DstBytes(DstLB To DstUB + SrcLen)
    CopyBytes DstBytes, DstUB + 1, SrcBytes, SrcLB, SrcLen
End Sub

Private Sub CopyBytes( _
    DstBytes() As Byte, DstIndex As Long, _
    SrcBytes, SrcIndex As Long, ByVal Length As Long)
    'SrcBytes() As Byte, SrcIndex As Long, ByVal Length As Long)
    
    Dim Offset As Long
    For Offset = 0 To Length - 1
        DstBytes(DstIndex + Offset) = SrcBytes(SrcIndex + Offset)
    Next
End Sub

Private Sub ReverseBytes( _
    ByRef Bytes() As Byte, _
    Optional Index As Long, _
    Optional ByVal Length As Long)
    
    Dim UB As Long
    
    If Length = 0 Then
        UB = UBound(Bytes)
        Length = UB - Index + 1
    Else
        UB = Index + Length - 1
    End If
    
    Dim Offset As Long
    For Offset = 0 To (Length \ 2) - 1
        Dim Temp As Byte
        Temp = Bytes(Index + Offset)
        Bytes(Index + Offset) = Bytes(UB - Offset)
        Bytes(UB - Offset) = Temp
    Next
End Sub

''
'' CBOR for VBA - Encoding - Converter
''

'
' 0xf9. Float16 - an IEEE 754 half precision floating point number
'

'Private Function GetBytesFromFloat16( _
'    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
'
'    GetBytesFromFloat16 = GetBytesFromSingle(Value, BigEndian)
'End Function

'
' 0xfa. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromFloat32( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat32 = GetBytesFromSingle(Value, BigEndian)
End Function

'
' 0xfb. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromFloat64( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat64 = GetBytesFromDouble(Value, BigEndian)
End Function

'
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromSingle( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    Dim S As SingleT
    S.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = S
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromSingle = B4.Bytes
End Function

'
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromDouble( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    Dim D As DoubleT
    D.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = D
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromDouble = B8.Bytes
End Function

''
'' CBOR for VBA - Decoding
''

Public Function GetCborLength( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    Dim ItemCount As Long
    Dim ItemLength As Long
    
    Select Case CborBytes(Index)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    'Case &HF9
    '    GetCborLength = 1 + 2
        
    ' 0xfa | single-precision float (four-byte IEEE 754)
    Case &HFA
        GetCborLength = 1 + 4
        
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HFB
        GetCborLength = 1 + 8
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function IsCborObject( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    Select Case CborBytes(Index)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    ' 0xfa | single-precision float (four-byte IEEE 754)
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HF9 To &HFB
        IsCborObject = False
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetValue(CborBytes() As Byte, Optional Index As Long) As Variant
    Select Case CborBytes(Index)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    'Case &HF9
    '    GetValue = GetFloat16FromCborBytes(CborBytes, Index)
        
    ' 0xfa | single-precision float (four-byte IEEE 754)
    Case &HFA
        GetValue = GetFloat32FromCborBytes(CborBytes, Index)
        
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HFB
        GetValue = GetFloat64FromCborBytes(CborBytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' major type 7: simple/float
'

' 0xf9 | half-precision float (two-byte IEEE 754)
'Private Function GetFloat16FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Single
'
'    GetFloat16FromCborBytes = GetFloat16FromBytes(CborBytes, Index + 1, True)
'End Function

' 0xfa | single-precision float (four-byte IEEE 754)
Private Function GetFloat32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Single
    
    GetFloat32FromCborBytes = GetFloat32FromBytes(CborBytes, Index + 1, True)
End Function

' 0xfb | double-precision float (eight-byte IEEE 754)
Private Function GetFloat64FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Double
    
    GetFloat64FromCborBytes = GetFloat64FromBytes(CborBytes, Index + 1, True)
End Function

''
'' CBOR for VBA - Decoding - Converter
''

'
' 0xf9. Float16 - an IEEE 754 half precision floating point number
'

'Private Function GetFloat16FromBytes(Bytes() As Byte, _
'    Optional Index As Long, Optional BigEndian As Boolean) As Single
'
'    GetFloat16FromBytes = GetSingleFromBytes(Bytes, Index, BigEndian)
'End Function

'
' 0xfa. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetFloat32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    GetFloat32FromBytes = GetSingleFromBytes(Bytes, Index, BigEndian)
End Function

'
' 0xfb. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetFloat64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    GetFloat64FromBytes = GetDoubleFromBytes(Bytes, Index, BigEndian)
End Function

'
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetSingleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim S As SingleT
    LSet S = B4
    
    GetSingleFromBytes = S.Value
End Function

'
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetDoubleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim D As DoubleT
    LSet D = B8
    
    GetDoubleFromBytes = D.Value
End Function

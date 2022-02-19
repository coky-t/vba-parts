Attribute VB_Name = "CBOR_01_Int"
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
    
    ' 2
    Case vbInteger
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 3
    Case vbLong
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 14 - temporaly work around
    Case vbDecimal
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 17
    Case vbByte
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 20
    #If Win64 And USE_LONGLONG Then
    Case vbLongLong
        GetCborBytes = GetCborBytesFromInt(Value)
    #End If
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 2. Integer
' 3. Long
' 17. Byte
' 20. LongLong
'
Private Function GetCborBytesFromInt(Value) As Byte()
    If Value >= 0 Then
        GetCborBytesFromInt = GetCborBytesFromPosInt(Value)
    Else
        GetCborBytesFromInt = GetCborBytesFromNegInt(Value)
    End If
End Function

Private Function GetCborBytesFromPosInt(Value) As Byte()
    Select Case Value
    
    Case 0 To 23 '&H17
        GetCborBytesFromPosInt = GetCborBytesFromPosFixInt((Value))
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromPosInt = GetCborBytesFromPosInt8((Value))
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromPosInt = GetCborBytesFromPosInt16((Value))
        
    #If Win64 And USE_LONGLONG Then
    Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
        GetCborBytesFromPosInt = GetCborBytesFromPosInt32((Value))
        
    Case Else
        GetCborBytesFromPosInt = GetCborBytesFromPosInt64((Value))
        
    #Else
    Case Else
        GetCborBytesFromPosInt = GetCborBytesFromPosInt32((Value))
        
    #End If
        
    End Select
End Function

Private Function GetCborBytesFromNegInt(Value) As Byte()
    Select Case Value
    
    Case -24 To -1
        GetCborBytesFromNegInt = GetCborBytesFromNegFixInt((Value))
        
    Case -256 To -25
        GetCborBytesFromNegInt = GetCborBytesFromNegInt8((Value))
        
    Case -65536 To -257
        GetCborBytesFromNegInt = GetCborBytesFromNegInt16((Value))
        
    #If Win64 And USE_LONGLONG Then
    Case -4294967296^ To -65537
        GetCborBytesFromNegInt = GetCborBytesFromNegInt32((Value))
        
    Case Else
        GetCborBytesFromNegInt = GetCborBytesFromNegInt64((Value))
        
    #Else
    Case Else
        GetCborBytesFromNegInt = GetCborBytesFromNegInt32((Value))
        
    #End If
        
    End Select
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 0: positive integer
'

' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)

Private Function GetCborBytesFromPosFixInt(ByVal Value As Byte) As Byte()
    Debug.Assert (Value <= &H17)
    GetCborBytesFromPosFixInt = GetCborBytes0(Value)
End Function

' 0x18 | unsigned integer (one-byte uint8_t follows)

Private Function GetCborBytesFromPosInt8(ByVal Value As Byte) As Byte()
    GetCborBytesFromPosInt8 = _
        GetCborBytes1(&H18, GetBytesFromUInt8(Value))
End Function

' 0x19 | unsigned integer (two-byte uint16_t follows)

Private Function GetCborBytesFromPosInt16(ByVal Value As Long) As Byte()
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    GetCborBytesFromPosInt16 = _
        GetCborBytes1(&H19, GetBytesFromUInt16(Value, True))
End Function

' 0x1a | unsigned integer (four-byte uint32_t follows)

Private Function GetCborBytesFromPosInt32(ByVal Value) As Byte()
    #If Win64 Then
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    #Else
    Debug.Assert (Value >= 0)
    #End If
    GetCborBytesFromPosInt32 = _
        GetCborBytes1(&H1A, GetBytesFromUInt32(Value, True))
End Function

' 0x1b | unsigned integer (eight-byte uint64_t follows)

Private Function GetCborBytesFromPosInt64(ByVal Value) As Byte()
    Debug.Assert (Value >= 0)
    GetCborBytesFromPosInt64 = _
        GetCborBytes1(&H1B, GetBytesFromUInt64(Value, True))
End Function

'
' major type 1: negative integer
'

' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)

Private Function GetCborBytesFromNegFixInt(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -24) And (Value <= -1))
    GetCborBytesFromNegFixInt = GetCborBytes0(&H20 Or CByte(Abs(Value + 1)))
End Function

' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)

Private Function GetCborBytesFromNegInt8(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -256) And (Value <= -1))
    GetCborBytesFromNegInt8 = _
        GetCborBytes1(&H38, GetBytesFromUInt8(CByte(Abs(Value + 1))))
End Function

' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)

Private Function GetCborBytesFromNegInt16(ByVal Value As Long) As Byte()
    Debug.Assert ((Value >= -65536) And (Value <= -1))
    GetCborBytesFromNegInt16 = _
        GetCborBytes1(&H39, GetBytesFromUInt16(CLng(Abs(Value + 1)), True))
End Function

' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)

#If Win64 And USE_LONGLONG Then

Private Function GetCborBytesFromNegInt32(ByVal Value As LongLong) As Byte()
    Debug.Assert ((Value >= -4294967296^) And (Value <= -1))
    GetCborBytesFromNegInt32 = _
        GetCborBytes1(&H3A, GetBytesFromUInt32(CLngLng(Abs(Value + 1)), True))
End Function

#Else

Private Function GetCborBytesFromNegInt32(ByVal Value) As Byte()
    Debug.Assert (Value <= -1)
    GetCborBytesFromNegInt32 = _
        GetCborBytes1(&H3A, GetBytesFromUInt32(Abs(Value + 1), True))
End Function

#End If

' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)

Private Function GetCborBytesFromNegInt64(ByVal Value) As Byte()
    Debug.Assert (Value <= -1)
    GetCborBytesFromNegInt64 = _
        GetCborBytes1(&H3B, GetBytesFromUInt64(Abs(Value + 1), True))
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
' 0x18. UInt8 - a 8-bit unsigned integer
'

Private Function GetBytesFromUInt8(ByVal Value As Byte) As Byte()
    Dim Bytes(0) As Byte
    Bytes(0) = Value
    GetBytesFromUInt8 = Bytes
End Function

'
' 0x19. UInt16 - a 16-bit unsigned integer
'

Private Function GetBytesFromUInt16( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    
    Dim Bytes4() As Byte
    Bytes4 = GetBytesFromLong(Value, BigEndian)
    
    Dim Bytes(0 To 1) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes4, 2, 2
    Else
        CopyBytes Bytes, 0, Bytes4, 0, 2
    End If
    
    GetBytesFromUInt16 = Bytes
End Function

'
' 0x1a. UInt32 - a 32-bit unsigned integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetBytesFromUInt32( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    
    Dim Bytes8() As Byte
    Bytes8 = GetBytesFromLongLong(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes8, 4, 4
    Else
        CopyBytes Bytes, 0, Bytes8, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#Else

Private Function GetBytesFromUInt32( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("4294967295")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 10, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#End If

'
' 0x1b. UInt64 - a 64-bit unsigned integer
'

Private Function GetBytesFromUInt64( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("18446744073709551615")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 7) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 6, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 8
    End If
    
    GetBytesFromUInt64 = Bytes
End Function

'
' 2. Integer - a 16-bit signed integer
'

Private Function GetBytesFromInteger( _
    ByVal Value As Integer, Optional BigEndian As Boolean) As Byte()
    
    Dim I As IntegerT
    I.Value = Value
    
    Dim B2 As Bytes2T
    LSet B2 = I
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    GetBytesFromInteger = B2.Bytes
End Function

'
' 3. Long - a 32-bit signed integer
'

Private Function GetBytesFromLong( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Dim L As LongT
    L.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = L
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromLong = B4.Bytes
End Function

'
' 14. Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Private Function GetBytesFromDecimal( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    CopyMemory ByVal VarPtr(BytesRaw(0)), ByVal VarPtr(Value), 16
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    CopyBytes BytesX, 0, BytesRaw, 2, 14
    
    ' Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim Bytes(0 To 13) As Byte
    
    ' data low bytes
    CopyBytes Bytes, 0, BytesX, 6, 8
    
    ' data high bytes
    CopyBytes Bytes, 8, BytesX, 2, 4
    
    ' scale
    Bytes(12) = BytesX(0)
    
    ' sign
    Bytes(13) = BytesX(1)
    
    If BigEndian Then
        ReverseBytes Bytes
    End If
    
    GetBytesFromDecimal = Bytes
End Function

'
' 20. LongLong - a 64-bit signed integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetBytesFromLongLong( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Dim LL As LongLongT
    LL.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = LL
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromLongLong = B8.Bytes
End Function

#End If

''
'' CBOR for VBA - Decoding
''

Public Function GetCborLength( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    Dim ItemCount As Long
    Dim ItemLength As Long
    
    Select Case CborBytes(Index)
    
    '
    ' major type 0: positive integer
    '
    
    ' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)
    Case &H0 To &H17
        GetCborLength = 1
        
    ' 0x18 | unsigned integer (one-byte uint8_t follows)
    Case &H18
        GetCborLength = 1 + 1
        
    ' 0x19 | unsigned integer (two-byte uint16_t follows)
    Case &H19
        GetCborLength = 1 + 2
        
    ' 0x1a | unsigned integer (four-byte uint32_t follows)
    Case &H1A
        GetCborLength = 1 + 4
        
    ' 0x1b | unsigned integer (eight-byte uint64_t follows)
    Case &H1B
        GetCborLength = 1 + 8
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    Case &H20 To &H37
        GetCborLength = 1
        
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    Case &H38
        GetCborLength = 1 + 1
        
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    Case &H39
        GetCborLength = 1 + 2
        
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    Case &H3A
        GetCborLength = 1 + 4
        
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H3B
        GetCborLength = 1 + 8
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function IsCborObject( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    Select Case CborBytes(Index)
    
    '
    ' major type 0: positive integer
    '
    
    ' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)
    ' 0x18 | unsigned integer (one-byte uint8_t follows)
    ' 0x19 | unsigned integer (two-byte uint16_t follows)
    ' 0x1a | unsigned integer (four-byte uint32_t follows)
    ' 0x1b | unsigned integer (eight-byte uint64_t follows)
    Case &H0 To &H1B
        IsCborObject = False
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H20 To &H3B
        IsCborObject = False
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetValue( _
    CborBytes() As Byte, Optional Index As Long) As Variant
    
    Select Case CborBytes(Index)
    
    '
    ' major type 0: positive integer
    '
    
    ' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)
    Case &H0 To &H17
        GetValue = GetPosFixIntFromCborBytes(CborBytes, Index)
        
    ' 0x18 | unsigned integer (one-byte uint8_t follows)
    Case &H18
        GetValue = GetPosInt8FromCborBytes(CborBytes, Index)
        
    ' 0x19 | unsigned integer (two-byte uint16_t follows)
    Case &H19
        GetValue = GetPosInt16FromCborBytes(CborBytes, Index)
        
    ' 0x1a | unsigned integer (four-byte uint32_t follows)
    Case &H1A
        GetValue = GetPosInt32FromCborBytes(CborBytes, Index)
        
    ' 0x1b | unsigned integer (eight-byte uint64_t follows)
    Case &H1B
        GetValue = GetPosInt64FromCborBytes(CborBytes, Index)
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    Case &H20 To &H37
        GetValue = GetNegFixIntFromCborBytes(CborBytes, Index)
        
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    Case &H38
        GetValue = GetNegInt8FromCborBytes(CborBytes, Index)
        
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    Case &H39
        GetValue = GetNegInt16FromCborBytes(CborBytes, Index)
        
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    Case &H3A
        GetValue = GetNegInt32FromCborBytes(CborBytes, Index)
        
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H3B
        GetValue = GetNegInt64FromCborBytes(CborBytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' major type 0: positive integer
'

' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)

Private Function GetPosFixIntFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte
    
    GetPosFixIntFromCborBytes = CborBytes(Index)
End Function

' 0x18 | unsigned integer (one-byte uint8_t follows)

Private Function GetPosInt8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte
    
    GetPosInt8FromCborBytes = CborBytes(Index + 1)
End Function

' 0x19 | unsigned integer (two-byte uint16_t follows)

Private Function GetPosInt16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    GetPosInt16FromCborBytes = GetUInt16FromBytes(CborBytes, Index + 1, True)
End Function

' 0x1a | unsigned integer (four-byte uint32_t follows)

Private Function GetPosInt32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetPosInt32FromCborBytes = GetUInt32FromBytes(CborBytes, Index + 1, True)
End Function

' 0x1b | unsigned integer (eight-byte uint64_t follows)

Private Function GetPosInt64FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetPosInt64FromCborBytes = GetUInt64FromBytes(CborBytes, Index + 1, True)
End Function

'
' major type 1: negative integer
'

' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)

Private Function GetNegFixIntFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Integer
    
    GetNegFixIntFromCborBytes = -1 - (CborBytes(Index) And &H1F)
End Function

' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)

Private Function GetNegInt8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Integer
    
    GetNegInt8FromCborBytes = -1 - CborBytes(Index + 1)
End Function

' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)

Private Function GetNegInt16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    GetNegInt16FromCborBytes = _
        -1 - GetUInt16FromBytes(CborBytes, Index + 1, True)
End Function

' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)

#If Win64 And USE_LONGLONG Then

Private Function GetNegInt32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As LongLong
    
    GetNegInt32FromCborBytes = _
        -1 - GetUInt32FromBytes(CborBytes, Index + 1, True)
End Function

#Else

Private Function GetNegInt32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetNegInt32FromCborBytes = _
        -1 - GetUInt32FromBytes(CborBytes, Index + 1, True)
End Function

#End If

' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)

Private Function GetNegInt64FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetNegInt64FromCborBytes = _
        -1 - GetUInt64FromBytes(CborBytes, Index + 1, True)
End Function

''
'' CBOR for VBA - Decoding - Converter
''

'
' 0x19. UInt16 - a 16-bit unsigned integer
'

Private Function GetUInt16FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim Bytes4(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes4, 2, Bytes, Index, 2
    Else
        CopyBytes Bytes4, 0, Bytes, Index, 2
    End If
    
    GetUInt16FromBytes = GetLongFromBytes(Bytes4, 0, BigEndian)
End Function

'
' 0x1a. UInt32 - a 32-bit unsigned integer
'

#If USE_LONGLONG Then

Private Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim Bytes8(0 To 7) As Byte
    If BigEndian Then
        CopyBytes Bytes8, 4, Bytes, Index, 4
    Else
        CopyBytes Bytes8, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetLongLongFromBytes(Bytes8, 0, BigEndian)
End Function

#Else

Private Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 10, Bytes, Index, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

#End If

'
' 0x1b. UInt64 - a 64-bit unsigned integer
'

Private Function GetUInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 6, Bytes, Index, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 8
    End If
    
    GetUInt64FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

'
' 2. Integer - a 16-bit signed integer
'

Private Function GetIntegerFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    Dim B2 As Bytes2T
    CopyBytes B2.Bytes, 0, Bytes, Index, 2
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    Dim I As IntegerT
    LSet I = B2
    
    GetIntegerFromBytes = I.Value
End Function

'
' 3. Long - a 32-bit signed integer
'

Private Function GetLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim L As LongT
    LSet L = B4
    
    GetLongFromBytes = L.Value
End Function

'
' 14. Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Private Function GetDecimalFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    ' BytesXX = Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim BytesXX(0 To 13) As Byte
    CopyBytes BytesXX, 0, Bytes, Index, 14
    
    If BigEndian Then
        ReverseBytes BytesXX
    End If
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    
    ' scale
    BytesX(0) = BytesXX(12)
    
    ' sign
    BytesX(1) = BytesXX(13)
    
    ' data high bytes
    CopyBytes BytesX, 2, BytesXX, 8, 4
    
    ' data low bytes
    CopyBytes BytesX, 6, BytesXX, 0, 8
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    BytesRaw(0) = 14
    BytesRaw(1) = 0
    CopyBytes BytesRaw, 2, BytesX, Index, 14
    
    Dim Value As Variant
    CopyMemory ByVal VarPtr(Value), ByVal VarPtr(BytesRaw(0)), 16
    
    GetDecimalFromBytes = Value
End Function

'
' 20. LongLong - a 64-bit signed integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetLongLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim LL As LongLongT
    LSet LL = B8
    
    GetLongLongFromBytes = LL.Value
End Function

#End If

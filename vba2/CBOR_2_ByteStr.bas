Attribute VB_Name = "CBOR_2_ByteStr"
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

' Array
#Const USE_COLLECTION = True

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
    
    ' 8209. (17 + 8192)
    Case vbByte + vbArray
        GetCborBytes = GetCborBytesFromByteArray(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 8209. (17 + 8192) Byte Array
'

Private Function GetCborBytesFromByteArray(Value) As Byte()
    Dim Length As Long
    
    On Error Resume Next
    Length = UBound(Value) - LBound(Value) + 1
    On Error GoTo 0
    
    Select Case Length
    
    Case 0
        GetCborBytesFromByteArray = GetCborBytes0(&H40)
        
    Case 1 To 23 '&H17
        GetCborBytesFromByteArray = GetCborBytesFromFixBin(Value, Length)
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromByteArray = GetCborBytesFromBin8(Value, Length)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromByteArray = GetCborBytesFromBin16(Value, Length)
        
    '#If Win64 And USE_LONGLONG Then
    'Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
    '    GetCborBytesFromByteArray = GetCborBytesFromBin32(Value, Length)
    '
    'Case Else
    '    GetCborBytesFromByteArray = GetCborBytesFromBin64(Value, Length)
    '
    '#Else
    Case Else
        GetCborBytesFromByteArray = GetCborBytesFromBin32(Value, Length)
        
    '#End If
        
    End Select
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 2: byte string
'

' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
Private Function GetCborBytesFromFixBin( _
    BinBytes, ByVal BinLength As Byte) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert ((BinLength > 0) And (BinLength <= &H17))
    
    GetCborBytesFromFixBin = GetCborBytes1(&H40 Or BinLength, BinBytes)
End Function

' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
Private Function GetCborBytesFromBin8( _
    BinBytes, ByVal BinLength As Byte) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetCborBytesFromBin8 = _
        GetCborBytes2(&H58, GetBytesFromUInt8(BinLength), BinBytes)
End Function

' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
Private Function GetCborBytesFromBin16( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetCborBytesFromBin16 = _
        GetCborBytes2(&H59, GetBytesFromUInt16(BinLength, True), BinBytes)
End Function

' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
Private Function GetCborBytesFromBin32( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetCborBytesFromBin32 = _
        GetCborBytes2(&H5A, GetBytesFromUInt32(BinLength, True), BinBytes)
End Function

' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
'Private Function GetCborBytesFromBin64( _
'    BinBytes, ByVal BinLength As LongLong) As Byte()
'    'BinBytes() As Byte, ByVal BinLength As LongLong) As Byte()
'
'    Debug.Assert (BinLength > 0)
'
'    GetCborBytesFromBin64 = _
'        GetCborBytes2(&H5B, GetBytesFromUInt64(BinLength, True), BinBytes)
'End Function

' 0x5f | byte string, byte strings follow, terminated by "break"
'Private Function GetCborBytesFromBinBreak( _
'    BinBytes, ByVal BinLength As Long) As Byte()
'    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
'
'    Debug.Assert (BinLength > 0)
'
'    GetCborBytesFromBinBreak = _
'        GetCborBytes2(&H5F, BinBytes, GetBytesFromUInt8(&HFF))
'End Function

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
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    Case &H40 To &H57
        ItemLength = (CborBytes(Index) And &H1F)
        GetCborLength = 1 + ItemLength
        
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    Case &H58
        ItemLength = CborBytes(Index + 1)
        GetCborLength = 1 + 1 + ItemLength
        
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    Case &H59
        ItemLength = GetUInt16FromBytes(CborBytes, Index + 1, True)
        GetCborLength = 1 + 2 + ItemLength
        
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    Case &H5A
        ItemLength = GetUInt32FromBytes(CborBytes, Index + 1, True)
        GetCborLength = 1 + 4 + ItemLength
        
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H5B
    '    ItemLength = GetUInt64FromBytes(CborBytes, Index + 1, True)
    '    GetCborLength = 1 + 8 + ItemLength
        
    ' 0x5f | byte string, byte strings follow, terminated by "break"
    'Case &H5F
    '    ItemLength = _
    '        GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
    '    GetCborLength = 1 + ItemLength + 1
        
    End Select
End Function

Public Function IsCborObject( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    Select Case CborBytes(Index)
    
    '
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x5f | byte string, byte strings follow, terminated by "break"
    Case &H40 To &H5B, &H5F
        IsCborObject = False
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetValue(CborBytes() As Byte, Optional Index As Long) As Variant
    Select Case CborBytes(Index)
    
    '
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    Case &H40 To &H57
        GetValue = GetFixBinFromCborBytes(CborBytes, Index)
        
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    Case &H58
        GetValue = GetBin8FromCborBytes(CborBytes, Index)
        
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    Case &H59
        GetValue = GetBin16FromCborBytes(CborBytes, Index)
        
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    Case &H5A
        GetValue = GetBin32FromCborBytes(CborBytes, Index)
        
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H5B
    '    GetValue = GetBin64FromCborBytes(CborBytes, Index)
        
    ' 0x5f | byte string, byte strings follow, terminated by "break"
    'Case &H5F
    '    GetValue = GetBinBreakFromCborBytes(CborBytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' major type 2: byte string
'

' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
Private Function GetFixBinFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Byte
    Length = CborBytes(Index) And &H1F
    If Length = 0 Then
        GetFixBinFromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1, Length
    
    GetFixBinFromCborBytes = BinBytes
End Function

' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
Private Function GetBin8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Byte
    Length = CborBytes(Index + 1)
    If Length = 0 Then
        GetBin8FromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 1, Length
    
    GetBin8FromCborBytes = BinBytes
End Function

' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
Private Function GetBin16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = GetUInt16FromBytes(CborBytes, Index + 1, True)
    If Length = 0 Then
        GetBin16FromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 2, Length
    
    GetBin16FromCborBytes = BinBytes
End Function

' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
Private Function GetBin32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    If Length = 0 Then
        GetBin32FromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 4, Length
    
    GetBin32FromCborBytes = BinBytes
End Function

' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
'Private Function GetBin64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Byte()
'
'    Dim BinBytes() As Byte
'
'    Dim Length As LongLong
'    Length = CLngLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
'    If Length = 0 Then
'        GetBin64FromCborBytes = BinBytes
'        Exit Function
'    End If
'
'    ReDim BinBytes(0 To Length - 1)
'    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 8, Length
'
'    GetBin64FromCborBytes = BinBytes
'End Function

' 0x5f | byte string, byte strings follow, terminated by "break"
'Private Function GetBinBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Byte()
'
'    Dim BinBytes() As Byte
'
'    Dim Length As Long
'    Length = GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
'    If Length = 0 Then
'        GetBinBreakFromCborBytes = BinBytes
'        Exit Function
'    End If
'
'    ReDim BinBytes(0 To Length - 1)
'    CopyBytes BinBytes, 0, CborBytes, Index + 1, Length
'
'    GetBinBreakFromCborBytes = BinBytes
'End Function

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

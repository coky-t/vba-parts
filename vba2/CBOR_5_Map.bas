Attribute VB_Name = "CBOR_5_Map"
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
    
    ' 2
    Case vbInteger
        GetCborBytes = CBOR_01_Int.GetCborBytes(Value)
        
    ' 3
    Case vbLong
        GetCborBytes = CBOR_01_Int.GetCborBytes(Value)
        
    ' 14 - temporaly work around
    Case vbDecimal
        GetCborBytes = CBOR_01_Int.GetCborBytes(Value)
        
    ' 17
    Case vbByte
        GetCborBytes = CBOR_01_Int.GetCborBytes(Value)
        
    ' 20
    #If Win64 And USE_LONGLONG Then
    Case vbLongLong
        GetCborBytes = CBOR_01_Int.GetCborBytes(Value)
    #End If
        
    ' 8209. (17 + 8192)
    Case vbByte + vbArray
        GetCborBytes = CBOR_2_ByteStr.GetCborBytes(Value)
        
    ' 8
    Case vbString
        GetCborBytes = CBOR_3_TextStr.GetCborBytes(Value)
        
    ' 0
    Case vbEmpty
        GetCborBytes = CBOR_7_Simple.GetCborBytes(Value)
        
    ' 1
    Case vbNull
        GetCborBytes = CBOR_7_Simple.GetCborBytes(Value)
        
    ' 11
    Case vbBoolean
        GetCborBytes = CBOR_7_Simple.GetCborBytes(Value)
        
    ' 4
    Case vbSingle
        GetCborBytes = CBOR_7_Float.GetCborBytes(Value)
        
    ' 5
    Case vbDouble
        GetCborBytes = CBOR_7_Float.GetCborBytes(Value)
        
    ' ---
    
   ' 9
    Case vbObject
        GetCborBytes = GetCborBytesFromObject(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 9. Object
'

Private Function GetCborBytesFromObject(Value) As Byte()
    If Value Is Nothing Then
        GetCborBytesFromObject = GetCborBytes0(&HF6) 'Null
        Exit Function
    End If
    
    Select Case TypeName(Value)
    Case "Collection"
        GetCborBytesFromObject = CBOR_4_Array.GetCborBytes(Value)
        
    Case "Dictionary"
        GetCborBytesFromObject = GetCborBytesFromDictionary(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 9. Object - Dictionary
'

Private Function GetCborBytesFromDictionary(Value) As Byte()
    Select Case Value.Count
    
    Case 0
        GetCborBytesFromDictionary = GetCborBytes0(&HA0)
        
    Case 1 To 23 ' &H17
        GetCborBytesFromDictionary = GetCborBytesFromFixMap(Value)
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromDictionary = GetCborBytesFromMap8(Value)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromDictionary = GetCborBytesFromMap16(Value)
        
    '#If Win64 And USE_LONGLONG Then
    'Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
    '    GetCborBytesFromDictionary = GetCborBytesFromMap32(Value)
    '
    'Case Else
    '    GetCborBytesFromDictionary = GetCborBytesFromMap64(Value)
    '
    '#Else
    Case Else
        GetCborBytesFromDictionary = GetCborBytesFromMap32(Value)
        
    '#End If
        
    End Select
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 5: map
'

' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
Private Function GetCborBytesFromFixMap(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    Debug.Assert ((Value.Count > 0) And (Value.Count <= &H17))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HA0 Or Value.Count
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromFixMap = CborBytes
End Function

' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
Private Function GetCborBytesFromMap8(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    Debug.Assert ((Value.Count > 0) And (Value.Count <= &HFF))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HB8
    AddBytes CborBytes, GetBytesFromUInt8(Value.Count)
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromMap8 = CborBytes
End Function

' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
Private Function GetCborBytesFromMap16(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    Debug.Assert ((Value.Count > 0) And (Value.Count <= &HFFFF&))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HB9
    AddBytes CborBytes, GetBytesFromUInt16(Value.Count, True)
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromMap16 = CborBytes
End Function

' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
Private Function GetCborBytesFromMap32(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HBA
    AddBytes CborBytes, GetBytesFromUInt32(Value.Count, True)
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromMap32 = CborBytes
End Function

' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
'Private Function GetCborBytesFromMap64( _
'    Value, Optional Optimize As Boolean) As Byte()
'
'    Dim CborBytes() As Byte
'    ReDim CborBytes(0)
'    CborBytes(0) = &HBB
'    AddBytes CborBytes, GetBytesFromUInt64(Value.Count, True)
'
'    AddCborBytesFromMap CborBytes, Value
'
'    GetCborBytesFromMap64 = CborBytes
'End Function

' 0xbf | map, pairs of data items follow, terminated by "break"
'Private Function GetCborBytesFromMapBreak(Value) As Byte()
'    Debug.Assert (TypeName(Value) = "Dictionary")
'
'    Dim CborBytes() As Byte
'    ReDim CborBytes(0)
'    CborBytes(0) = &HBF
'    AddCborBytesFromMap CborBytes, Value
'
'    AddBytes CborBytes, GetBytesFromUInt8(&HFF)
'
'    GetCborBytesFromMapBreak = CborBytes
'End Function

''
'' CBOR for VBA - Encoding - Map Helper
''

Private Sub AddCborBytesFromMap(CborBytes() As Byte, Value)
    Debug.Assert (TypeName(Value) = "Dictionary")
    
    Dim Keys
    Keys = Value.Keys
    
    Dim Index As Long
    For Index = LBound(Keys) To UBound(Keys)
        AddBytes CborBytes, GetCborBytes(Keys(Index))
        AddBytes CborBytes, GetCborBytes(Value.Item(Keys(Index)))
    Next
End Sub

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
    ' 0x18 | unsigned integer (one-byte uint8_t follows)
    ' 0x19 | unsigned integer (two-byte uint16_t follows)
    ' 0x1a | unsigned integer (four-byte uint32_t follows)
    ' 0x1b | unsigned integer (eight-byte uint64_t follows)
    Case 0 To &H1B
        GetCborLength = CBOR_01_Int.GetCborLength(CborBytes, Index)
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H20 To &H3B
        GetCborLength = CBOR_01_Int.GetCborLength(CborBytes, Index)
        
    '
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    Case &H40 To &H5A
        GetCborLength = CBOR_2_ByteStr.GetCborLength(CborBytes, Index)
        
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x5f | byte string, byte strings follow, terminated by "break"
        
    '
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    Case &H60 To &H7A
        GetCborLength = CBOR_3_TextStr.GetCborLength(CborBytes, Index)
        
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    
    '
    ' major type 4: array
    '
    
    ' 0x80..0x97 | array (0x00..0x17 data items follow)
    ' 0x98 | array (one-byte uint8_t for n, and then n data items follow)
    ' 0x99 | array (two-byte uint16_t for n, and then n data items follow)
    ' 0x9a | array (four-byte uint32_t for n, and then n data items follow)
    Case &H80 To &H9A
        GetCborLength = CBOR_4_Array.GetCborLength(CborBytes, Index)
        
    ' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)
    ' 0x9f | array, data items follow, terminated by "break"
        
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    ' 0xf5 | true
    ' 0xf6 | null
    ' 0xf7 | undefined
    Case &HF4 To &HF7
        GetCborLength = 1
        
    ' 0xf9 | half-precision float (two-byte IEEE 754)
        
    ' 0xfa | single-precision float (four-byte IEEE 754)
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HFA, &HFB
        GetCborLength = CBOR_7_Float.GetCborLength(CborBytes, Index)
        
    ' 0xff | "break" stop code
    
    ' ----
    
    '
    ' major type 5: map
    '
    
    ' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
    Case &HA0 To &HB7
        ItemCount = CborBytes(Index) And &H1F
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1)
        GetCborLength = 1 + ItemLength
        
    ' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
    Case &HB8
        ItemCount = CborBytes(Index + 1)
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 1)
        GetCborLength = 1 + 1 + ItemLength
        
    ' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
    Case &HB9
        ItemCount = GetUInt16FromBytes(CborBytes, Index + 1, True)
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 2)
        GetCborLength = 1 + 2 + ItemLength
        
    ' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
    Case &HBA
        ItemCount = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 4)
        GetCborLength = 1 + 4 + ItemLength
        
    ' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
    'Case &HBB
    '    ItemCount = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
    '    ItemLength = _
    '        GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 8)
    '    GetCborLength = 1 + 8 + ItemLength
        
    ' 0xbf | map, pairs of data items follow, terminated by "break"
    'Case &HBF
    '    ' to do
    '    ItemLength = 0
    '    GetCborLength = 1 + ItemLength + 1
        
    End Select
End Function

Private Function GetCborLengthFromItemCborBytes(ByVal ItemCount As Long, _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Dim Count As Long
    For Count = 1 To ItemCount
        Length = Length + GetCborLength(CborBytes, Index + Length)
    Next
    
    GetCborLengthFromItemCborBytes = Length
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
        
    '
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    Case &H60 To &H7B, &H7F
        IsCborObject = False
        
    '
    ' major type 4: array
    '
    
    ' 0x80..0x97 | array (0x00..0x17 data items follow)
    ' 0x98 | array (one-byte uint8_t for n, and then n data items follow)
    ' 0x99 | array (two-byte uint16_t for n, and then n data items follow)
    ' 0x9a | array (four-byte uint32_t for n, and then n data items follow)
    ' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)
    ' 0x9f | array, data items follow, terminated by "break"
    Case &H80 To &H9B ', &H9F
        #If USE_COLLECTION Then
        IsCborObject = True
        #Else
        IsCborObject = False
        #End If
        
    '
    ' major type 5: map
    '
    
    ' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
    ' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
    ' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
    ' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
    ' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
    ' 0xbf | map, pairs of data items follow, terminated by "break"
    Case &HA0 To &HBB ', &HBF
        IsCborObject = True
        
    '
    ' major type 6: tag
    '
    
    ' 0xc0 | text-based date/time (data item follows; see Section 3.4.1)
    ' 0xc1 | epoch-based date/time (data item follows; see Section 3.4.2)
    ' 0xc2 | unsigned bignum (data item "byte string" follows)
    ' 0xc3 | negative bignum (data item "byte string" follows)
    ' 0xc4 | decimal Fraction (data item "array" follows; see Section 3.4.4)
    ' 0xc5 | bigfloat (data item "array" follows; see Section 3.4.4)
    ' 0xc6..0xd4 | (tag)
    ' 0xd5..0xd7 | expected conversion (data item follows; see Section 3.4.5.2)
    ' 0xd8..0xdb | (more tags; 1/2/4/8 bytes of tag number and then a data item follow)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    ' 0xf5 | true
    ' 0xf6 | null
    ' 0xf7 | undefined
    Case &HF4 To &HF7 ', &HFF
        IsCborObject = False
        
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    ' 0xfa | single-precision float (four-byte IEEE 754)
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HF9 To &HFB
        IsCborObject = False
        
    ' 0xff | "break" stop code
    
    End Select
End Function

Public Function GetValue(CborBytes() As Byte, Optional Index As Long) As Variant
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
        GetValue = CBOR_01_Int.GetValue(CborBytes, Index)
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H20 To &H3B
        GetValue = CBOR_01_Int.GetValue(CborBytes, Index)
        
    '
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x5f | byte string, byte strings follow, terminated by "break"
    Case &H40 To &H5A ', &H5B, &H5F
        GetValue = CBOR_2_ByteStr.GetValue(CborBytes, Index)
        
    '
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    Case &H60 To &H7A ', &H7B, &H7F
        GetValue = CBOR_3_TextStr.GetValue(CborBytes, Index)
        
    '
    ' major type 4: array
    '
    
    ' 0x80..0x97 | array (0x00..0x17 data items follow)
    ' 0x98 | array (one-byte uint8_t for n, and then n data items follow)
    ' 0x99 | array (two-byte uint16_t for n, and then n data items follow)
    ' 0x9a | array (four-byte uint32_t for n, and then n data items follow)
    ' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)
    ' 0x9f | array, data items follow, terminated by "break"
    Case &H80 To &H9A ', &H9B, &H9F
        GetValue = CBOR_4_Array.GetValue(CborBytes, Index)
        
    '
    ' major type 5: map
    '
    
    ' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
    ' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
    ' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
    ' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
    ' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
    ' 0xbf | map, pairs of data items follow, terminated by "break"
    
    '
    ' major type 6: tag
    '
    
    ' 0xc0 | text-based date/time (data item follows; see Section 3.4.1)
    ' 0xc1 | epoch-based date/time (data item follows; see Section 3.4.2)
    ' 0xc2 | unsigned bignum (data item "byte string" follows)
    ' 0xc3 | negative bignum (data item "byte string" follows)
    ' 0xc4 | decimal Fraction (data item "array" follows; see Section 3.4.4)
    ' 0xc5 | bigfloat (data item "array" follows; see Section 3.4.4)
    ' 0xc6..0xd4 | (tag)
    ' 0xd5..0xd7 | expected conversion (data item follows; see Section 3.4.5.2)
    ' 0xd8..0xdb | (more tags; 1/2/4/8 bytes of tag number and then a data item follow)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    ' 0xf5 | true
    ' 0xf6 | null
    ' 0xf7 | undefined
    Case &HF4 To &HF7
        GetValue = CBOR_7_Simple.GetValue(CborBytes, Index)
        
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    ' 0xfa | single-precision float (four-byte IEEE 754)
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HF9 To &HFB
        GetValue = CBOR_7_Float.GetValue(CborBytes, Index)
        
    ' 0xff | "break" stop code
    
    ' ----
    
    '
    ' major type 5: map
    '
    
    ' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
    Case &HA0 To &HB7
        Set GetValue = GetFixMapFromCborBytes(CborBytes, Index)
        
    ' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
    Case &HB8
        Set GetValue = GetMap8FromCborBytes(CborBytes, Index)
        
    ' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
    Case &HB9
        Set GetValue = GetMap16FromCborBytes(CborBytes, Index)
        
    ' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
    Case &HBA
        Set GetValue = GetMap32FromCborBytes(CborBytes, Index)
        
    ' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
    'Case &HBB
    '    Set GetValue = GetMap64FromCborBytes(CborBytes, Index)
        
    ' 0xbf | map, pairs of data items follow, terminated by "break"
    'Case &HBF
    '    Set GetValue = GetMapBreakFromCborBytes(CborBytes, Index)
        
    End Select
End Function

'
' major type 5: map
'

' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
Private Function GetFixMapFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index) And &H1F
    
    Set GetFixMapFromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1, ItemCount)
End Function

' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
Private Function GetMap8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index + 1)
    
    Set GetMap8FromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1 + 1, ItemCount)
End Function

' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
Private Function GetMap16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = GetUInt16FromBytes(CborBytes, Index + 1, True)
    
    Set GetMap16FromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1 + 2, ItemCount)
End Function

' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
Private Function GetMap32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    
    Set GetMap32FromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1 + 4, ItemCount)
End Function

' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
'Private Function GetMap64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Object
'
'    Dim ItemCount As Long
'    ItemCount = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
'
'    Set GetMap64FromCborBytes = _
'        GetMapFromCborBytes(CborBytes, Index + 1 + 8, ItemCount)
'End Function

' 0xbf | map, pairs of data items follow, terminated by "break"
'Private Function GetMapBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Object
'
'    Dim ItemCount As Long
'    ItemCount = 0 ' to do
'
'    Set GetMapBreakFromCborBytes = _
'        GetMapFromCborBytes(CborBytes, Index + 1, ItemCount)
'End Function

''
'' CBOR for VBA - Decoding - Map Helper
''

Private Function GetMapFromCborBytes( _
    CborBytes() As Byte, Index As Long, ItemCount As Long) As Object
    
    Dim Map As Object
    Set Map = CreateObject("Scripting.Dictionary")
    
    Dim KeyOffset As Long
    Dim ValueOffset As Long
    
    Dim Count As Long
    For Count = 1 To ItemCount
        ValueOffset = KeyOffset + GetCborLength(CborBytes, Index + KeyOffset)
        
        Map.Add _
            GetValue(CborBytes, Index + KeyOffset), _
            GetValue(CborBytes, Index + ValueOffset)
        
        KeyOffset = ValueOffset + GetCborLength(CborBytes, Index + ValueOffset)
    Next
    
    Set GetMapFromCborBytes = Map
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

Attribute VB_Name = "CBOR_3_TextStr"
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
    
    ' 8
    Case vbString
        GetCborBytes = GetCborBytesFromString((Value))
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 8. String
'

Private Function GetCborBytesFromString(ByVal Value As String) As Byte()
    If CStr(Value) = "" Then
        GetCborBytesFromString = GetCborBytes0(&H60)
        Exit Function
    End If
    
    Dim StrBytes() As Byte
    StrBytes = GetBytesFromString(Value)
    
    Dim StrLength As Long
    StrLength = UBound(StrBytes) - LBound(StrBytes) + 1
    
    Select Case StrLength
    
    Case 1 To 23 '&H17
        GetCborBytesFromString = GetCborBytesFromFixStr(StrBytes, StrLength)
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromString = GetCborBytesFromStr8(StrBytes, StrLength)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromString = GetCborBytesFromStr16(StrBytes, StrLength)
        
    '#If Win64 And USE_LONGLONG Then
    'Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
    '    GetCborBytesFromString = GetCborBytesFromStr32(StrBytes, StrLength)
    '
    'Case Else
    '    GetCborBytesFromString = GetCborBytesFromStr64(StrBytes, StrLength)
    '
    '#Else
    Case Else
        GetCborBytesFromString = GetCborBytesFromStr32(StrBytes, StrLength)
        
    '#End If
        
    End Select
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 3: text string
'

' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
Private Function GetCborBytesFromFixStr( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert ((StrLength > 0) And (StrLength <= &H17))
    
    GetCborBytesFromFixStr = GetCborBytes1(&H60 Or StrLength, StrBytes)
End Function

' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
Private Function GetCborBytesFromStr8( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetCborBytesFromStr8 = _
        GetCborBytes2(&H78, GetBytesFromUInt8(StrLength), StrBytes)
End Function

' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
Private Function GetCborBytesFromStr16( _
    StrBytes() As Byte, ByVal StrLength As Long) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetCborBytesFromStr16 = _
        GetCborBytes2(&H79, GetBytesFromUInt16(StrLength, True), StrBytes)
End Function

' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
Private Function GetCborBytesFromStr32( _
    StrBytes() As Byte, ByVal StrLength) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetCborBytesFromStr32 = _
        GetCborBytes2(&H7A, GetBytesFromUInt32(StrLength, True), StrBytes)
End Function

' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
'Private Function GetCborBytesFromStr64( _
'    StrBytes() As Byte, ByVal StrLength) As Byte()
'
'    Debug.Assert (StrLength > 0)
'
'    GetCborBytesFromStr64 = _
'        GetCborBytes2(&H7B, GetBytesFromUInt64(StrLength, True), StrBytes)
'End Function

' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
'Private Function GetCborBytesFromStrBreak( _
'    StrBytes() As Byte, ByVal StrLength) As Byte()
'
'    Debug.Assert (StrLength > 0)
'
'    GetCborBytesFromStrBreak = _
'        GetCborBytes2(&H7F, StrBytes, GetBytesFromUInt8(&HFF))
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
' 8. String
'

Private Function GetBytesFromString( _
    ByVal Value As String, Optional Charset As String = "utf-8") As Byte()
    
    Debug.Assert (Value <> "")
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        .WriteText Value
        
        .Position = 0
        .Type = 1 'ADODB.adTypeBinary
        If Charset = "utf-8" Then
            .Position = 3 ' avoid BOM
        End If
        GetBytesFromString = .Read
        
        .Close
    End With
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
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    Case &H60 To &H77
        ItemLength = (CborBytes(Index) And &H1F)
        GetCborLength = 1 + ItemLength
        
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    Case &H78
        ItemLength = CborBytes(Index + 1)
        GetCborLength = 1 + 1 + ItemLength
        
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    Case &H79
        ItemLength = GetUInt16FromBytes(CborBytes, Index + 1, True)
        GetCborLength = 1 + 2 + ItemLength
        
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    Case &H7A
        ItemLength = GetUInt32FromBytes(CborBytes, Index + 1, True)
        GetCborLength = 1 + 4 + ItemLength
        
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H7B
    '    ItemLength = GetUInt64FromBytes(CborBytes, Index + 1, True)
    '    GetCborLength = 1 + 8 + ItemLength
        
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    'Case &H7F
    '    ItemLength = _
    '        GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
    '    GetCborLength = 1 + ItemLength + 1
        
    End Select
End Function

Public Function IsCborObject( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    Select Case CborBytes(Index)
    
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
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetValue(CborBytes() As Byte, Optional Index As Long) As Variant
    Select Case CborBytes(Index)
    
    '
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    Case &H60 To &H77
        GetValue = GetFixStrFromCborBytes(CborBytes, Index)
        
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    Case &H78
        GetValue = GetStr8FromCborBytes(CborBytes, Index)
        
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    Case &H79
        GetValue = GetStr16FromCborBytes(CborBytes, Index)
        
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    Case &H7A
        GetValue = GetStr32FromCborBytes(CborBytes, Index)
        
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H7B
    '    GetValue = GetStr64FromCborBytes(CborBytes, Index)
        
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    'Case &H7F
    '    GetValue = GetStrBreakFromCborBytes(CborBytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' major type 3: text string
'

' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
Private Function GetFixStrFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Byte
    Length = (CborBytes(Index) And &H1F)
    If Length = 0 Then
        GetFixStrFromCborBytes = ""
        Exit Function
    End If
    
    GetFixStrFromCborBytes = GetStringFromBytes(CborBytes, Index + 1, Length)
End Function

' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
Private Function GetStr8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Byte
    Length = CborBytes(Index + 1)
    If Length = 0 Then
        GetStr8FromCborBytes = ""
        Exit Function
    End If
    
    GetStr8FromCborBytes = _
        GetStringFromBytes(CborBytes, Index + 1 + 1, Length)
End Function

' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
Private Function GetStr16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Long
    Length = GetUInt16FromBytes(CborBytes, Index + 1, True)
    If Length = 0 Then
        GetStr16FromCborBytes = ""
        Exit Function
    End If
    
    GetStr16FromCborBytes = _
        GetStringFromBytes(CborBytes, Index + 1 + 2, Length)
End Function

' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
Private Function GetStr32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Long
    Length = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    If Length = 0 Then
        GetStr32FromCborBytes = ""
        Exit Function
    End If
    
    GetStr32FromCborBytes = _
        GetStringFromBytes(CborBytes, Index + 1 + 4, Length)
End Function

' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
'Private Function GetStr64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As String
'
'    Dim Length As LongLong
'    Length = CLngLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
'    If Length = 0 Then
'        GetStr64FromCborBytes = ""
'        Exit Function
'    End If
'
'    GetStr64FromCborBytes = _
'        GetStringFromBytes(CborBytes, Index + 1 + 8, Length)
'End Function

' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
'Private Function GetStrBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As String
'
'    Dim Length As Long
'    Length = GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
'    If Length = 0 Then
'        GetStrBreakFromCborBytes = ""
'        Exit Function
'    End If
'
'    GetStrBreakFromCborBytes = _
'        GetStringFromBytes(CborBytes, Index + 1, Length)
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
' 8. String
'

Private Function GetStringFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional ByVal Length As Long, _
    Optional Charset As String = "utf-8") As String
    
    If Length = 0 Then
        Length = UBound(Bytes) - Index + 1
    End If
    
    Dim Bytes_() As Byte
    ReDim Bytes_(0 To Length - 1)
    CopyBytes Bytes_, 0, Bytes, Index, Length
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 1 'ADODB.adTypeBinary
        .Write Bytes_
        
        .Position = 0
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        GetStringFromBytes = .ReadText
        
        .Close
    End With
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

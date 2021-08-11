Attribute VB_Name = "MsgPack_Ext"
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
'' MessagePack for VBA - Extension
''

''
'' MessagePack for VBA - Extension - Serialization
''

Public Function IsVBAExt(Value) As Boolean
    Select Case VarType(Value)
    Case vbByte + vbArray
        IsVBAExt = True
        
    Case Else
        IsVBAExt = False
        
    End Select
End Function

Public Function GetBytesFromExt(ExtType As Byte, Value) As Byte()
    Debug.Assert IsVBAExt(Value)
    
    Dim Length As Long
    Length = GetLengthFromExt(Value)
    
    If Length = 0 Then
        Dim Bytes(0 To 2) As Byte
        Bytes(0) = &HC7
        Bytes(1) = &H0
        Bytes(2) = ExtType
        GetBytesFromExt = Bytes
        Exit Function
    End If
    
    Select Case Length
    
    'fixext 1        | 11010100               | 0xd4
    Case &H1
        GetBytesFromExt = GetBytesFromExtBytes_FixExt1(ExtType, Value)
        Exit Function
        
    'fixext 2        | 11010101               | 0xd5
    Case &H2
        GetBytesFromExt = GetBytesFromExtBytes_FixExt2(ExtType, Value)
        Exit Function
        
    'fixext 4        | 11010110               | 0xd6
    Case &H4
        GetBytesFromExt = GetBytesFromExtBytes_FixExt4(ExtType, Value)
        Exit Function
        
    'fixext 8        | 11010111               | 0xd7
    Case &H8
        GetBytesFromExt = GetBytesFromExtBytes_FixExt8(ExtType, Value)
        Exit Function
        
    'fixext 16       | 11011000               | 0xd8
    Case &H10
        GetBytesFromExt = GetBytesFromExtBytes_FixExt16(ExtType, Value)
        Exit Function
        
    Case Else
        ' nop - continue
        
    End Select
    
    Select Case Length
    
    'ext 8           | 11000111               | 0xc7
    Case &H0 To &HFF
        GetBytesFromExt = GetBytesFromExtBytes_Ext8(ExtType, Value, Length)
        
    'ext 16          | 11001000               | 0xc8
    Case &H100 To &HFFFF&
        GetBytesFromExt = GetBytesFromExtBytes_Ext16(ExtType, Value, Length)
        
    'ext 32          | 11001001               | 0xc9
    Case Else
        GetBytesFromExt = GetBytesFromExtBytes_Ext32(ExtType, Value, Length)
        
    End Select
End Function

Private Function GetLengthFromExt(Value) As Long
    On Error Resume Next
    GetLengthFromExt = UBound(Value) - LBound(Value) + 1
    On Error GoTo 0
End Function

'ext 8           | 11000111               | 0xc7
'ext 8 stores an integer and a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+--------+========+
'|  0xc7  |XXXXXXXX|  type  |  data  |
'+--------+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_Ext8( _
    ExtType As Byte, BinBytes, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetBytesFromExtBytes_Ext8 = _
        MsgPack_Common.GetBytesHelper3A(&HC7, BinLength, ExtType, BinBytes)
End Function

'ext 16          | 11001000               | 0xc8
'ext 16 stores an integer and a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+--------+========+
'|  0xc8  |YYYYYYYY|YYYYYYYY|  type  |  data  |
'+--------+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_Ext16( _
    ExtType As Byte, BinBytes, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetBytesFromExtBytes_Ext16 = _
        MsgPack_Common.GetBytesHelper3B(&HC8, _
            BitConverter.GetBytesFromUInt16(BinLength, True), _
            ExtType, BinBytes)
End Function

'ext 32          | 11001001               | 0xc9
'ext 32 stores an integer and a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+--------+========+
'|  0xc9  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  type  |  data  |
'+--------+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a big-endian 32-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_Ext32( _
    ExtType As Byte, BinBytes, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetBytesFromExtBytes_Ext32 = _
        MsgPack_Common.GetBytesHelper3B(&HC9, _
            BitConverter.GetBytesFromUInt32(BinLength, True), _
            ExtType, BinBytes)
End Function

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_FixExt1( _
    ExtType As Byte, BinBytes) As Byte()
    
    GetBytesFromExtBytes_FixExt1 = _
        MsgPack_Common.GetBytesHelper2A(&HD4, ExtType, BinBytes)
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_FixExt2( _
    ExtType As Byte, BinBytes) As Byte()
    
    GetBytesFromExtBytes_FixExt2 = _
        MsgPack_Common.GetBytesHelper2A(&HD5, ExtType, BinBytes)
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_FixExt4( _
    ExtType As Byte, BinBytes) As Byte()
    
    GetBytesFromExtBytes_FixExt4 = _
        MsgPack_Common.GetBytesHelper2A(&HD6, ExtType, BinBytes)
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_FixExt8( _
    ExtType As Byte, BinBytes) As Byte()
    
    GetBytesFromExtBytes_FixExt8 = _
        MsgPack_Common.GetBytesHelper2A(&HD7, ExtType, BinBytes)
End Function

'fixext 16       | 11011000               | 0xd8
'fixext 16 stores an integer and a byte array whose length is 16 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd8  |  type  |                                  data
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                              data (cont.)                              |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetBytesFromExtBytes_FixExt16( _
    ExtType As Byte, BinBytes) As Byte()
    
    GetBytesFromExtBytes_FixExt16 = _
        MsgPack_Common.GetBytesHelper2A(&HD8, ExtType, BinBytes)
End Function

''
'' MessagePack for VBA - Extension - Deserialization
''

Public Function IsMPExt(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'ext 8           | 11000111               | 0xc7
    'ext 16          | 11001000               | 0xc8
    'ext 32          | 11001001               | 0xc9
    'fixext 1        | 11010100               | 0xd4
    'fixext 2        | 11010101               | 0xd5
    'fixext 4        | 11010110               | 0xd6
    'fixext 8        | 11010111               | 0xd7
    'fixext 16       | 11011000               | 0xd8
    Case &HC7 To &HC9, &HD4 To &HD8
        IsMPExt = True
        
    Case Else
        IsMPExt = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Select Case Bytes(Index)
    
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        Length = Bytes(Index + 1)
        GetLengthFromBytes = 1 + 1 + 1 + Length
        
    'ext 16          | 11001000               | 0xc8
    Case &HC8
        Length = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
        GetLengthFromBytes = 1 + 2 + 1 + Length
        
    'ext 32          | 11001001               | 0xc9
    Case &HC9
        Length = BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True)
        GetLengthFromBytes = 1 + 4 + 1 + Length
        
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetLengthFromBytes = 1 + 1 + 1
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetLengthFromBytes = 1 + 1 + 2
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetLengthFromBytes = 1 + 1 + 4
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetLengthFromBytes = 1 + 1 + 8
        
    'fixext 16       | 11011000               | 0xd8
    Case &HD8
        GetLengthFromBytes = 1 + 1 + 16
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetExtFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Select Case Bytes(Index)
    
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        GetExtFromBytes = GetExtFromBytes_Ext8(Bytes, Index)
        
    'ext 16          | 11001000               | 0xc8
    Case &HC8
        GetExtFromBytes = GetExtFromBytes_Ext16(Bytes, Index)
        
    'ext 32          | 11001001               | 0xc9
    Case &HC9
        GetExtFromBytes = GetExtFromBytes_Ext32(Bytes, Index)
        
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetExtFromBytes = GetExtFromBytes_FixExt1(Bytes, Index)
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetExtFromBytes = GetExtFromBytes_FixExt2(Bytes, Index)
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetExtFromBytes = GetExtFromBytes_FixExt4(Bytes, Index)
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetExtFromBytes = GetExtFromBytes_FixExt8(Bytes, Index)
        
    'fixext 16       | 11011000               | 0xd8
    Case &HD8
        GetExtFromBytes = GetExtFromBytes_FixExt16(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'ext 8           | 11000111               | 0xc7
'ext 8 stores an integer and a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+--------+========+
'|  0xc7  |XXXXXXXX|  type  |  data  |
'+--------+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_Ext8( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HC7)
    
    Dim ExtBytes() As Byte
    
    Dim Length As Byte
    Length = Bytes(Index + 1)
    If Length = 0 Then
        'ReDim ExtBytes(0)
        'ExtBytes(0) = Bytes(Index + 2) ' type
        GetExtFromBytes_Ext8 = ExtBytes
        Exit Function
    End If
    
    'ReDim ExtBytes(0 To Length)
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 1, 1 + Length
    ReDim ExtBytes(0 To Length - 1)
    BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 1 + 1, Length
    
    GetExtFromBytes_Ext8 = ExtBytes
End Function

'ext 16          | 11001000               | 0xc8
'ext 16 stores an integer and a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+--------+========+
'|  0xc8  |YYYYYYYY|YYYYYYYY|  type  |  data  |
'+--------+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_Ext16( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HC8)
    
    Dim ExtBytes() As Byte
    
    Dim Length As Long
    Length = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
    If Length = 0 Then
        'ReDim ExtBytes(0)
        'ExtBytes(0) = Bytes(Index + 1 + 2) ' type
        GetExtFromBytes_Ext16 = ExtBytes
        Exit Function
    End If
    
    'ReDim ExtBytes(0 To Length)
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 2, 1 + Length
    ReDim ExtBytes(0 To Length - 1)
    BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 2 + 1, Length
    
    GetExtFromBytes_Ext16 = ExtBytes
End Function

'ext 32          | 11001001               | 0xc9
'ext 32 stores an integer and a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+--------+========+
'|  0xc9  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  type  |  data  |
'+--------+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a big-endian 32-bit unsigned integer which represents N
'* N is a length of data
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_Ext32( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HC9)
    
    Dim ExtBytes() As Byte
    
    Dim Length As Long
    Length = CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
    If Length = 0 Then
        'ReDim ExtBytes(0)
        'ExtBytes(0) = Bytes(Index + 1 + 4) ' type
        GetExtFromBytes_Ext32 = ExtBytes
        Exit Function
    End If
    
    'ReDim ExtBytes(0 To Length)
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 4, 1 + Length
    ReDim ExtBytes(0 To Length - 1)
    BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 4 + 1, Length
    
    GetExtFromBytes_Ext32 = ExtBytes
End Function

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_FixExt1( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HD4)
    
    'Dim ExtBytes(0 To 1) As Byte
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1, 1 + 1
    Dim ExtBytes(0) As Byte
    ExtBytes(0) = Bytes(Index + 1 + 1)
    
    GetExtFromBytes_FixExt1 = ExtBytes
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_FixExt2( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HD5)
    
    'Dim ExtBytes(0 To 2) As Byte
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1, 1 + 2
    Dim ExtBytes(0 To 1) As Byte
    BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 1, 2
    
    GetExtFromBytes_FixExt2 = ExtBytes
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_FixExt4( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HD6)
    
    'Dim ExtBytes(0 To 4) As Byte
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1, 1 + 4
    Dim ExtBytes(0 To 3) As Byte
    BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 1, 4
    
    GetExtFromBytes_FixExt4 = ExtBytes
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_FixExt8( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HD7)
    
    'Dim ExtBytes(0 To 8) As Byte
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1, 1 + 8
    Dim ExtBytes(0 To 7) As Byte
    BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 1, 8
    
    GetExtFromBytes_FixExt8 = ExtBytes
End Function

'fixext 16       | 11011000               | 0xd8
'fixext 16 stores an integer and a byte array whose length is 16 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd8  |  type  |                                  data
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                              data (cont.)                              |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtFromBytes_FixExt16( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HD8)
    
    'Dim ExtBytes(0 To 16) As Byte
    'BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1, 1 + 16
    Dim ExtBytes(0 To 15) As Byte
    BitConverter.CopyBytes ExtBytes, 0, Bytes, Index + 1 + 1, 16
    
    GetExtFromBytes_FixExt16 = ExtBytes
End Function

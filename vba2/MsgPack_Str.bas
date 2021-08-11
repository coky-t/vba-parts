Attribute VB_Name = "MsgPack_Str"
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
'' MessagePack for VBA - String
''

''
'' MessagePack for VBA - String - Serialization
''

Public Function IsVBAStr(Value) As Boolean
    Select Case VarType(Value)
    Case vbString
        IsVBAStr = True
        
    Case Else
        IsVBAStr = False
        
    End Select
End Function

Public Function GetBytesFromStr(Value) As Byte()
    Debug.Assert IsVBAStr(Value)
    
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    If CStr(Value) = "" Then
        Dim Bytes(0) As Byte
        Bytes(0) = &HA0
        GetBytesFromStr = Bytes
        Exit Function
    End If
    
    Dim StrBytes() As Byte
    StrBytes = BitConverter.GetBytesFromString(Value)
    
    Dim StrLength
    StrLength = UBound(StrBytes) - LBound(StrBytes) + 1
    
    Select Case StrLength
    
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &H1 To &H1F
        GetBytesFromStr = GetBytesFromStrBytes_FixStr(StrBytes, StrLength)
        
    'str 8           | 11011001               | 0xd9
    Case &H20 To &HFF
        GetBytesFromStr = GetBytesFromStrBytes_Str8(StrBytes, StrLength)
        
    'str 16          | 11011010               | 0xda
    Case &H100 To &HFFFF&
        GetBytesFromStr = GetBytesFromStrBytes_Str16(StrBytes, StrLength)
        
    'str 32          | 11011011               | 0xdb
    Case Else
        GetBytesFromStr = GetBytesFromStrBytes_Str32(StrBytes, StrLength)
        
    End Select
End Function

'fixstr          | 101xxxxx               | 0xa0 - 0xbf
'fixstr stores a byte array whose length is upto 31 bytes:
'+--------+========+
'|101XXXXX|  data  |
'+--------+========+
'* XXXXX is a 5-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetBytesFromStrBytes_FixStr( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetBytesFromStrBytes_FixStr = _
        MsgPack_Common.GetBytesHelper1(&HA0 Or StrLength, StrBytes)
End Function

'str 8           | 11011001               | 0xd9
'str 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xd9  |YYYYYYYY|  data  |
'+--------+--------+========+
'* YYYYYYYY is a 8-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetBytesFromStrBytes_Str8( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetBytesFromStrBytes_Str8 = _
        MsgPack_Common.GetBytesHelper2A(&HD9, StrLength, StrBytes)
End Function

'str 16          | 11011010               | 0xda
'str 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xda  |ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetBytesFromStrBytes_Str16( _
    StrBytes() As Byte, ByVal StrLength As Long) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetBytesFromStrBytes_Str16 = _
        MsgPack_Common.GetBytesHelper2B(&HDA, _
            BitConverter.GetBytesFromUInt16(StrLength, True), StrBytes)
End Function

'str 32          | 11011011               | 0xdb
'str 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xdb  |AAAAAAAA|AAAAAAAA|AAAAAAAA|AAAAAAAA|  data  |
'+--------+--------+--------+--------+--------+========+
'* AAAAAAAA_AAAAAAAA_AAAAAAAA_AAAAAAAA is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetBytesFromStrBytes_Str32( _
    StrBytes() As Byte, ByVal StrLength) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetBytesFromStrBytes_Str32 = _
        MsgPack_Common.GetBytesHelper2B(&HDB, _
            BitConverter.GetBytesFromUInt32(StrLength, True), StrBytes)
End Function

''
'' MessagePack for VBA - String - Deserialization
''

Public Function IsMPStr(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    'str 8           | 11011001               | 0xd9
    'str 16          | 11011010               | 0xda
    'str 32          | 11011011               | 0xdb
    Case &HA0 To &HBF, &HD9 To &HDB
        IsMPStr = True
        
    Case Else
        IsMPStr = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Select Case Bytes(Index)
    
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &HA0 To &HBF
        Length = (Bytes(Index) And &H1F)
        GetLengthFromBytes = 1 + Length
        
    'str 8           | 11011001               | 0xd9
    Case &HD9
        Length = Bytes(Index + 1)
        GetLengthFromBytes = 1 + 1 + Length
        
    'str 16          | 11011010               | 0xda
    Case &HDA
        Length = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
        GetLengthFromBytes = 1 + 2 + Length
        
    'str 32          | 11011011               | 0xdb
    Case &HDB
        Length = CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
        GetLengthFromBytes = 1 + 4 + Length
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetStrFromBytes( _
    Bytes() As Byte, Optional Index As Long) As String
    
    Select Case Bytes(Index)
    
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &HA0 To &HBF
        GetStrFromBytes = GetStrFromBytes_FixStr(Bytes, Index)
        
    'str 8           | 11011001               | 0xd9
    Case &HD9
        GetStrFromBytes = GetStrFromBytes_Str8(Bytes, Index)
        
    'str 16          | 11011010               | 0xda
    Case &HDA
        GetStrFromBytes = GetStrFromBytes_Str16(Bytes, Index)
        
    'str 32          | 11011011               | 0xdb
    Case &HDB
        GetStrFromBytes = GetStrFromBytes_Str32(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'fixstr          | 101xxxxx               | 0xa0 - 0xbf
'fixstr stores a byte array whose length is upto 31 bytes:
'+--------+========+
'|101XXXXX|  data  |
'+--------+========+
'* XXXXX is a 5-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetStrFromBytes_FixStr( _
    Bytes() As Byte, Optional Index As Long) As String
    
    Debug.Assert ((Bytes(Index) And &HE0) = &HA0)
    
    Dim Length As Byte
    Length = (Bytes(Index) And &H1F)
    If Length = 0 Then
        GetStrFromBytes_FixStr = ""
        Exit Function
    End If
    
    GetStrFromBytes_FixStr = _
        BitConverter.GetStringFromBytes(Bytes, Index + 1, Length)
End Function

'str 8           | 11011001               | 0xd9
'str 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xd9  |YYYYYYYY|  data  |
'+--------+--------+========+
'* YYYYYYYY is a 8-bit unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetStrFromBytes_Str8( _
    Bytes() As Byte, Optional Index As Long) As String
    
    Debug.Assert (Bytes(Index) = &HD9)
    
    Dim Length As Byte
    Length = Bytes(Index + 1)
    If Length = 0 Then
        GetStrFromBytes_Str8 = ""
        Exit Function
    End If
    
    GetStrFromBytes_Str8 = _
        BitConverter.GetStringFromBytes(Bytes, Index + 1 + 1, Length)
End Function

'str 16          | 11011010               | 0xda
'str 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xda  |ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetStrFromBytes_Str16( _
    Bytes() As Byte, Optional Index As Long) As String
    
    Debug.Assert (Bytes(Index) = &HDA)
    
    Dim Length As Long
    Length = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
    If Length = 0 Then
        GetStrFromBytes_Str16 = ""
        Exit Function
    End If
    
    GetStrFromBytes_Str16 = _
        BitConverter.GetStringFromBytes(Bytes, Index + 1 + 2, Length)
End Function

'str 32          | 11011011               | 0xdb
'str 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xdb  |AAAAAAAA|AAAAAAAA|AAAAAAAA|AAAAAAAA|  data  |
'+--------+--------+--------+--------+--------+========+
'* AAAAAAAA_AAAAAAAA_AAAAAAAA_AAAAAAAA is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data
'* **String** extending Raw type represents a UTF-8 string
Private Function GetStrFromBytes_Str32( _
    Bytes() As Byte, Optional Index As Long) As String
    
    Debug.Assert (Bytes(Index) = &HDB)
    
    Dim Length As Long
    Length = CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
    If Length = 0 Then
        GetStrFromBytes_Str32 = ""
        Exit Function
    End If
    
    GetStrFromBytes_Str32 = _
        BitConverter.GetStringFromBytes(Bytes, Index + 1 + 4, Length)
End Function

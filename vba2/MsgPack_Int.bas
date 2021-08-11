Attribute VB_Name = "MsgPack_Int"
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
'' MessagePack for VBA - Integer
''

''
'' MessagePack for VBA - Integer - Serialization
''

Public Function IsVBAInt(Value) As Boolean
    Select Case VarType(Value)
    Case vbByte, vbInteger, vbLong
        IsVBAInt = True
        
    #If Win64 Then
    Case vbLongLong
        IsVBAInt = True
    #End If
        
    Case Else
        IsVBAInt = False
        
    End Select
End Function

Public Function GetBytesFromInt(Value) As Byte()
    Debug.Assert IsVBAInt(Value)
    
    Select Case Value
    
    #If Win64 Then
    Case -2147483648^ To -32769
        GetBytesFromInt = GetBytesFromInt32(Value)
    #End If
        
    Case -32768 To -129
        GetBytesFromInt = GetBytesFromInt16(Value)
        
    Case -128 To -33
        GetBytesFromInt = GetBytesFromInt8(Value)
        
    Case -32 To -1
        GetBytesFromInt = GetBytesFromNegativeFixInt(Value)
        
    Case 0 To 127 '&H7F
        GetBytesFromInt = GetBytesFromPositiveFixInt(Value)
        
    Case 128 To 255 '&H80 To &HFF
        GetBytesFromInt = GetBytesFromUInt8(Value)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetBytesFromInt = GetBytesFromUInt16(Value)
        
    #If Win64 Then
    Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
        GetBytesFromInt = GetBytesFromUInt32(Value)
    #End If
        
    Case Else
    #If Win64 Then
        GetBytesFromInt = GetBytesFromInt64(Value)
    #Else
        GetBytesFromInt = GetBytesFromInt32(Value)
    #End If
        
    End Select
End Function

'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
'positive fixint stores 7-bit positive integer
'+--------+
'|0XXXXXXX|
'+--------+
'* 0XXXXXXX is 8-bit unsigned integer
Public Function GetBytesFromPositiveFixInt(ByVal Value As Byte) As Byte()
    Debug.Assert (Value <= &H7F)
    
    Dim Bytes(0) As Byte
    Bytes(0) = Value
    
    GetBytesFromPositiveFixInt = Bytes
End Function

'uint 8          | 11001100               | 0xcc
'uint 8 stores a 8-bit unsigned integer
'+--------+--------+
'|  0xcc  |ZZZZZZZZ|
'+--------+--------+
Public Function GetBytesFromUInt8(ByVal Value As Byte) As Byte()
    Dim Bytes(0 To 1) As Byte
    Bytes(0) = &HCC
    Bytes(1) = Value
    
    GetBytesFromUInt8 = Bytes
End Function

'uint 16         | 11001101               | 0xcd
'uint 16 stores a 16-bit big-endian unsigned integer
'+--------+--------+--------+
'|  0xcd  |ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+
Public Function GetBytesFromUInt16(ByVal Value As Long) As Byte()
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    
    GetBytesFromUInt16 = _
        MsgPack_Common.GetBytesHelper1(&HCD, _
            BitConverter.GetBytesFromUInt16(Value, True))
End Function

'uint 32         | 11001110               | 0xce
'uint 32 stores a 32-bit big-endian unsigned integer
'+--------+--------+--------+--------+--------+
'|  0xce  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+
#If Win64 Then
Public Function GetBytesFromUInt32(ByVal Value As LongLong) As Byte()
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    
    GetBytesFromUInt32 = _
        MsgPack_Common.GetBytesHelper1(&HCE, _
            BitConverter.GetBytesFromUInt32(Value, True))
End Function
#End If

'uint 64         | 11001111               | 0xcf
'uint 64 stores a 64-bit big-endian unsigned integer
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xcf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
#If Win64 Then
Public Function GetBytesFromUInt64(ByVal Value As Variant) As Byte()
    Debug.Assert (Value >= 0)
    
    GetBytesFromUInt64 = _
        MsgPack_Common.GetBytesHelper1(&HCF, _
            BitConverter.GetBytesFromUInt64(Value, True))
End Function
#End If

'int 8           | 11010000               | 0xd0
'int 8 stores a 8-bit signed integer
'+--------+--------+
'|  0xd0  |ZZZZZZZZ|
'+--------+--------+
Public Function GetBytesFromInt8(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -128) And (Value <= &H7F))
    
    Dim Bytes(0 To 1) As Byte
    Bytes(0) = &HD0
    Bytes(1) = Value And &HFF
    
    GetBytesFromInt8 = Bytes
End Function

'int 16          | 11010001               | 0xd1
'int 16 stores a 16-bit big-endian signed integer
'+--------+--------+--------+
'|  0xd1  |ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+
Public Function GetBytesFromInt16(ByVal Value As Integer) As Byte()
    GetBytesFromInt16 = _
        MsgPack_Common.GetBytesHelper1(&HD1, _
            BitConverter.GetBytesFromInt16(Value, True))
End Function

'int 32          | 11010010               | 0xd2
'int 32 stores a 32-bit big-endian signed integer
'+--------+--------+--------+--------+--------+
'|  0xd2  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+
Public Function GetBytesFromInt32(ByVal Value As Long) As Byte()
    GetBytesFromInt32 = _
        MsgPack_Common.GetBytesHelper1(&HD2, _
            BitConverter.GetBytesFromInt32(Value, True))
End Function

'int 64          | 11010011               | 0xd3
'int 64 stores a 64-bit big-endian signed integer
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd3  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
#If Win64 Then
Public Function GetBytesFromInt64(ByVal Value As LongLong) As Byte()
    GetBytesFromInt64 = _
        MsgPack_Common.GetBytesHelper1(&HD3, _
            BitConverter.GetBytesFromInt64(Value, True))
End Function
#End If

'negative fixint | 111xxxxx               | 0xe0 - 0xff
'negative fixint stores 5-bit negative integer
'+--------+
'|111YYYYY|
'+--------+
'* 111YYYYY is 8-bit signed integer
Public Function GetBytesFromNegativeFixInt(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -32) And (Value < 0))
    
    Dim Bytes(0) As Byte
    Bytes(0) = Value And &HFF
    
    GetBytesFromNegativeFixInt = Bytes
End Function

''
'' MessagePack for VBA - Integer - Deserialization
''

Public Function IsMPInt(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
    Case &H0 To &H7F
        IsMPInt = True
        
    'uint 8          | 11001100               | 0xcc
    'uint 16         | 11001101               | 0xcd
    'uint 32         | 11001110               | 0xce
    'uint 64         | 11001111               | 0xcf
    'int 8           | 11010000               | 0xd0
    'int 16          | 11010001               | 0xd1
    'int 32          | 11010010               | 0xd2
    'int 64          | 11010011               | 0xd3
    Case &HCC To &HD3
        IsMPInt = True
        
    'negative fixint | 111xxxxx               | 0xe0 - 0xff
    Case &HE0 To &HFF
        IsMPInt = True
        
    Case Else
        IsMPInt = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Select Case Bytes(Index)
    
    'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
    Case &H0 To &H7F
        GetLengthFromBytes = 1 + 0
        
    'uint 8          | 11001100               | 0xcc
    Case &HCC
        GetLengthFromBytes = 1 + 1
        
    'uint 16         | 11001101               | 0xcd
    Case &HCD
        GetLengthFromBytes = 1 + 2
        
    'uint 32         | 11001110               | 0xce
    Case &HCE
        GetLengthFromBytes = 1 + 4
        
    'uint 64         | 11001111               | 0xcf
    Case &HCF
        GetLengthFromBytes = 1 + 8
        
    'int 8           | 11010000               | 0xd0
    Case &HD0
        GetLengthFromBytes = 1 + 1
        
    'int 16          | 11010001               | 0xd1
    Case &HD1
        GetLengthFromBytes = 1 + 2
        
    'int 32          | 11010010               | 0xd2
    Case &HD2
        GetLengthFromBytes = 1 + 4
        
    'int 64          | 11010011               | 0xd3
    Case &HD3
        GetLengthFromBytes = 1 + 8
        
    'negative fixint | 111xxxxx               | 0xe0 - 0xff
    Case &HE0 To &HFF
        GetLengthFromBytes = 1 + 0
        
    Case Else
        Err.Raise 13 ' type mismatch
        
    End Select
End Function

Public Function GetIntFromBytes(Bytes() As Byte, Optional Index As Long)
    Select Case Bytes(Index)
    
    'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
    'positive fixint stores 7-bit positive integer
    '+--------+
    '|0XXXXXXX|
    '+--------+
    '* 0XXXXXXX is 8-bit unsigned integer
    Case &H0 To &H7F
        GetIntFromBytes = Bytes(Index)
        
    'uint 8          | 11001100               | 0xcc
    'uint 8 stores a 8-bit unsigned integer
    '+--------+--------+
    '|  0xcc  |ZZZZZZZZ|
    '+--------+--------+
    Case &HCC
        GetIntFromBytes = Bytes(Index + 1)
        
    'uint 16         | 11001101               | 0xcd
    'uint 16 stores a 16-bit big-endian unsigned integer
    '+--------+--------+--------+
    '|  0xcd  |ZZZZZZZZ|ZZZZZZZZ|
    '+--------+--------+--------+
    Case &HCD
        GetIntFromBytes = _
            BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
        
    'uint 32         | 11001110               | 0xce
    'uint 32 stores a 32-bit big-endian unsigned integer
    '+--------+--------+--------+--------+--------+
    '|  0xce  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
    '+--------+--------+--------+--------+--------+
    Case &HCE
        GetIntFromBytes = _
            BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True)
        
    'uint 64         | 11001111               | 0xcf
    'uint 64 stores a 64-bit big-endian unsigned integer
    '+--------+--------+--------+--------+--------+--------+--------+--------+--------+
    '|  0xcf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
    '+--------+--------+--------+--------+--------+--------+--------+--------+--------+
    Case &HCF
        GetIntFromBytes = _
            BitConverter.GetUInt64FromBytes(Bytes, Index + 1, True)
        
    'int 8           | 11010000               | 0xd0
    'int 8 stores a 8-bit signed integer
    '+--------+--------+
    '|  0xd0  |ZZZZZZZZ|
    '+--------+--------+
    Case &HD0
        GetIntFromBytes = _
            BitConverter.GetInt8FromBytes(Bytes, Index + 1, True)
        
    'int 16          | 11010001               | 0xd1
    'int 16 stores a 16-bit big-endian signed integer
    '+--------+--------+--------+
    '|  0xd1  |ZZZZZZZZ|ZZZZZZZZ|
    '+--------+--------+--------+
    Case &HD1
        GetIntFromBytes = _
            BitConverter.GetInt16FromBytes(Bytes, Index + 1, True)
        
    'int 32          | 11010010               | 0xd2
    'int 32 stores a 32-bit big-endian signed integer
    '+--------+--------+--------+--------+--------+
    '|  0xd2  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
    '+--------+--------+--------+--------+--------+
    Case &HD2
        GetIntFromBytes = _
            BitConverter.GetInt32FromBytes(Bytes, Index + 1, True)
        
    'int 64          | 11010011               | 0xd3
    'int 64 stores a 64-bit big-endian signed integer
    '+--------+--------+--------+--------+--------+--------+--------+--------+--------+
    '|  0xd3  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|
    '+--------+--------+--------+--------+--------+--------+--------+--------+--------+
    Case &HD3
        GetIntFromBytes = _
            BitConverter.GetInt64FromBytes(Bytes, Index + 1, True)
        
    'negative fixint | 111xxxxx               | 0xe0 - 0xff
    'negative fixint stores 5-bit negative integer
    '+--------+
    '|111YYYYY|
    '+--------+
    '* 111YYYYY is 8-bit signed integer
    Case &HE0 To &HFF
        GetIntFromBytes = _
            BitConverter.GetInt8FromBytes(Bytes, Index, True)
        
    Case Else
        Err.Raise 13 ' type mismatch
        
    End Select
End Function

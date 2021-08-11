Attribute VB_Name = "MsgPack_Bin"
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
'' MessagePack for VBA - Binary
''

''
'' MessagePack for VBA - Binary - Serialization
''

Public Function IsVBABin(Value) As Boolean
    Select Case VarType(Value)
    Case vbByte + vbArray
        IsVBABin = True
        
    Case Else
        IsVBABin = False
        
    End Select
End Function

Public Function GetBytesFromBin(Value) As Byte()
    Debug.Assert IsVBABin(Value)
    
    Dim Length As Long
    Length = GetLengthFromBin(Value)
    
    If Length = 0 Then
        Dim Bytes(0 To 1) As Byte
        Bytes(0) = &HC4
        Bytes(1) = &H0
        GetBytesFromBin = Bytes
        Exit Function
    End If
    
    Select Case Length
    
    'bin 8           | 11000100               | 0xc4
    Case &H1 To &HFF
        GetBytesFromBin = GetBytesFromBinBytes_Bin8(Value, Length)
        
    'bin 16          | 11000101               | 0xc5
    Case &H100 To &HFFFF&
        GetBytesFromBin = GetBytesFromBinBytes_Bin16(Value, Length)
        
    'bin 32          | 11000110               | 0xc6
    Case Else
        GetBytesFromBin = GetBytesFromBinBytes_Bin32(Value, Length)
        
    End Select
End Function

Private Function GetLengthFromBin(Value) As Long
    On Error Resume Next
    GetLengthFromBin = UBound(Value) - LBound(Value) + 1
    On Error GoTo 0
End Function

'bin 8           | 11000100               | 0xc4
'bin 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xc4  |XXXXXXXX|  data  |
'+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is the length of data
Private Function GetBytesFromBinBytes_Bin8( _
    BinBytes, ByVal BinLength As Byte) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetBytesFromBinBytes_Bin8 = _
        MsgPack_Common.GetBytesHelper2A(&HC4, BinLength, BinBytes)
End Function

'bin 16          | 11000101               | 0xc5
'bin 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xc5  |YYYYYYYY|YYYYYYYY|  data  |
'+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data
Private Function GetBytesFromBinBytes_Bin16( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetBytesFromBinBytes_Bin16 = _
        MsgPack_Common.GetBytesHelper2B(&HC5, _
            BitConverter.GetBytesFromUInt16(BinLength, True), BinBytes)
End Function

'bin 32          | 11000110               | 0xc6
'bin 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xc6  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data
Private Function GetBytesFromBinBytes_Bin32( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetBytesFromBinBytes_Bin32 = _
        MsgPack_Common.GetBytesHelper2B(&HC6, _
            BitConverter.GetBytesFromUInt32(BinLength, True), BinBytes)
End Function

''
'' MessagePack for VBA - Binary - Deserialization
''

Public Function IsMPBin(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'bin 8           | 11000100               | 0xc4
    'bin 16          | 11000101               | 0xc5
    'bin 32          | 11000110               | 0xc6
    Case &HC4 To &HC6
        IsMPBin = True
        
    Case Else
        IsMPBin = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Select Case Bytes(Index)
    
    'bin 8           | 11000100               | 0xc4
    Case &HC4
        Length = Bytes(Index + 1)
        GetLengthFromBytes = 1 + 1 + Length
        
    'bin 16          | 11000101               | 0xc5
    Case &HC5
        Length = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
        GetLengthFromBytes = 1 + 2 + Length
        
    'bin 32          | 11000110               | 0xc6
    Case &HC6
        Length = CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
        GetLengthFromBytes = 1 + 4 + Length
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetBinFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Select Case Bytes(Index)
    
    'bin 8           | 11000100               | 0xc4
    Case &HC4
        GetBinFromBytes = GetBinFromBytes_Bin8(Bytes, Index)
        
    'bin 16          | 11000101               | 0xc5
    Case &HC5
        GetBinFromBytes = GetBinFromBytes_Bin16(Bytes, Index)
        
    'bin 32          | 11000110               | 0xc6
    Case &HC6
        GetBinFromBytes = GetBinFromBytes_Bin32(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'bin 8           | 11000100               | 0xc4
'bin 8 stores a byte array whose length is upto (2^8)-1 bytes:
'+--------+--------+========+
'|  0xc4  |XXXXXXXX|  data  |
'+--------+--------+========+
'* XXXXXXXX is a 8-bit unsigned integer which represents N
'* N is the length of data
Private Function GetBinFromBytes_Bin8( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HC4)
    
    Dim BinBytes() As Byte
    
    Dim Length As Byte
    Length = Bytes(Index + 1)
    If Length = 0 Then
        GetBinFromBytes_Bin8 = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    BitConverter.CopyBytes BinBytes, 0, Bytes, Index + 1 + 1, Length
    
    GetBinFromBytes_Bin8 = BinBytes
End Function

'bin 16          | 11000101               | 0xc5
'bin 16 stores a byte array whose length is upto (2^16)-1 bytes:
'+--------+--------+--------+========+
'|  0xc5  |YYYYYYYY|YYYYYYYY|  data  |
'+--------+--------+--------+========+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the length of data
Private Function GetBinFromBytes_Bin16( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HC5)
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
    If Length = 0 Then
        GetBinFromBytes_Bin16 = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    BitConverter.CopyBytes BinBytes, 0, Bytes, Index + 1 + 2, Length
    
    GetBinFromBytes_Bin16 = BinBytes
End Function

'bin 32          | 11000110               | 0xc6
'bin 32 stores a byte array whose length is upto (2^32)-1 bytes:
'+--------+--------+--------+--------+--------+========+
'|  0xc6  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|  data  |
'+--------+--------+--------+--------+--------+========+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the length of data
Private Function GetBinFromBytes_Bin32( _
    Bytes() As Byte, Optional Index As Long) As Byte()
    
    Debug.Assert (Bytes(Index) = &HC6)
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
    If Length = 0 Then
        GetBinFromBytes_Bin32 = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    BitConverter.CopyBytes BinBytes, 0, Bytes, Index + 1 + 4, Length
    
    GetBinFromBytes_Bin32 = BinBytes
End Function

Attribute VB_Name = "MsgPack_Ext_Time"
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
'' MessagePack for VBA - Extension - Timestamp
''

Private Property Get mpTimestamp() As Long
    mpTimestamp = &HFF
End Property

''
'' MessagePack for VBA - Extension - Timestamp - Serialization
''

Public Function IsVBAExtTime(Value) As Boolean
    Select Case VarType(Value)
    Case vbDate
        IsVBAExtTime = True
        
    Case Else
        IsVBAExtTime = False
        
    End Select
End Function

Public Function GetBytesFromExtTime(Value) As Byte()
    Debug.Assert IsVBAExtTime(Value)
    
    If Value >= DateSerial(1970, 1, 1) Then
        If Value < DateSerial(2106, 2, 7) + TimeSerial(6, 28, 16) Then
            GetBytesFromExtTime = GetBytesFromExtTime_FixExt4(Value)
            Exit Function
        ElseIf Value < DateSerial(2514, 5, 30) + TimeSerial(1, 53, 4) Then
            GetBytesFromExtTime = GetBytesFromExtTime_FixExt8(Value)
            Exit Function
        End If
    End If
    
    GetBytesFromExtTime = GetBytesFromExtTime_Ext8(Value)
End Function

'fixext 4        | 11010110               | 0xd6
'timestamp 32 stores the number of seconds that have elapsed since 1970-01-01 00:00:00 UTC
'in an 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |   -1   |   seconds in 32-bit unsigned int  |
'+--------+--------+--------+--------+--------+--------+
'* Timestamp 32 format can represent a timestamp in [1970-01-01 00:00:00 UTC, 2106-02-07 06:28:16 UTC) range. Nanoseconds part is 0.
Private Function GetBytesFromExtTime_FixExt4(ByVal DateTime As Date) As Byte()
    Debug.Assert (DateTime >= DateSerial(1970, 1, 1))
    Debug.Assert (DateTime < DateSerial(2106, 2, 7) + TimeSerial(6, 28, 16))
    
    Dim Seconds As Double
    Seconds = DateDiff("s", DateSerial(1970, 1, 1), DateTime)
    
    GetBytesFromExtTime_FixExt4 = _
        MsgPack_Common.GetBytesHelper2A(&HD6, mpTimestamp, _
            BitConverter.GetBytesFromUInt32(Seconds, True))
End Function

'fixext 8        | 11010111               | 0xd7
'timestamp 64 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 32-bit unsigned integers:
'+--------+--------+--------+--------+--------+------|-+--------+--------+--------+--------+
'|  0xd7  |   -1   | nanosec. in 30-bit unsigned int |   seconds in 34-bit unsigned int    |
'+--------+--------+--------+--------+--------+------^-+--------+--------+--------+--------+
'* Timestamp 64 format can represent a timestamp in [1970-01-01 00:00:00.000000000 UTC, 2514-05-30 01:53:04.000000000 UTC) range.
Private Function GetBytesFromExtTime_FixExt8(ByVal DateTime As Date) As Byte()
    Debug.Assert (DateTime >= DateSerial(1970, 1, 1))
    Debug.Assert (DateTime < DateSerial(2514, 5, 30) + TimeSerial(1, 53, 4))
    
    Dim Seconds As Double
    Seconds = DateDiff("s", DateSerial(1970, 1, 1), DateTime)
    
    GetBytesFromExtTime_FixExt8 = _
        MsgPack_Common.GetBytesHelper2A(&HD7, mpTimestamp, _
            BitConverter.GetBytesFromUInt64(CDec(Seconds), True))
End Function

'ext 8           | 11000111               | 0xc7
'timestamp 96 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 64-bit signed integer and 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+--------+
'|  0xc7  |   12   |   -1   |nanoseconds in 32-bit unsigned int |
'+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                    seconds in 64-bit signed int                        |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* Timestamp 96 format can represent a timestamp in [-292277022657-01-27 08:29:52 UTC, 292277026596-12-04 15:30:08.000000000 UTC) range.
'* In timestamp 64 and timestamp 96 formats, nanoseconds must not be larger than 999999999.
Private Function GetBytesFromExtTime_Ext8(ByVal DateTime As Date) As Byte()
    Dim Seconds As Double
    Seconds = DateDiff("s", DateSerial(1970, 1, 1), DateTime)
    
    Dim SecBytes8() As Byte
    SecBytes8 = BitConverter.GetBytesFromInt64(Seconds, True)
    
    Dim SecBytes12(0 To 11) As Byte
    BitConverter.CopyBytes SecBytes12, 4, SecBytes8, 0, 8
    
    GetBytesFromExtTime_Ext8 = _
        MsgPack_Common.GetBytesHelper3A(&HC7, 12, mpTimestamp, SecBytes12)
End Function

''
'' MessagePack for VBA - Extension - Timestamp - Deserialization
''

Public Function IsMPExtTime(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'ext 16          | 11001000               | 0xc8
    'ext 32          | 11001001               | 0xc9
    'fixext 1        | 11010100               | 0xd4
    'fixext 2        | 11010101               | 0xd5
    'fixext 16       | 11011000               | 0xd8
    Case &HC8, &HC9, &HD4, &HD5, &HD8
        IsMPExtTime = False
        
    'ext 8           | 11000111               | 0xc7
    'fixext 4        | 11010110               | 0xd6
    'fixext 8        | 11010111               | 0xd7
    Case &HC7, &HD6, &HD7
        IsMPExtTime = (Bytes(Index + 1) = mpTimestamp)
        
    Case Else
        IsMPExtTime = False
        
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
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetLengthFromBytes = 1 + 1 + 4
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetLengthFromBytes = 1 + 1 + 8
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetExtTimeFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Date
    
    Select Case Bytes(Index)
    
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        GetExtTimeFromBytes = GetExtTimeFromBytes_Ext8(Bytes, Index)
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetExtTimeFromBytes = GetExtTimeFromBytes_FixExt4(Bytes, Index)
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetExtTimeFromBytes = GetExtTimeFromBytes_FixExt8(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'fixext 4        | 11010110               | 0xd6
'timestamp 32 stores the number of seconds that have elapsed since 1970-01-01 00:00:00 UTC
'in an 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |   -1   |   seconds in 32-bit unsigned int  |
'+--------+--------+--------+--------+--------+--------+
'* Timestamp 32 format can represent a timestamp in [1970-01-01 00:00:00 UTC, 2106-02-07 06:28:16 UTC) range. Nanoseconds part is 0.
Private Function GetExtTimeFromBytes_FixExt4( _
    Bytes() As Byte, Optional Index As Long) As Date
    
    Debug.Assert (Bytes(Index) = &HD6)
    Debug.Assert (Bytes(Index + 1) = mpTimestamp)
    
    Dim Seconds As Double
    Seconds = BitConverter.GetUInt32FromBytes(Bytes, Index + 1 + 1, True)
    
    'GetExtTimeFromBytes_FixExt4 = _
        DateAdd("s", Seconds, DateSerial(1970, 1, 1))
    
    Dim Days As Double
    Days = Int(Seconds / (CDbl(24) * 60 * 60))
    Seconds = Seconds - Days * (CDbl(24) * 60 * 60)
    
    GetExtTimeFromBytes_FixExt4 = _
        DateAdd("s", Seconds, DateAdd("d", Days, DateSerial(1970, 1, 1)))
End Function

'fixext 8        | 11010111               | 0xd7
'timestamp 64 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 32-bit unsigned integers:
'+--------+--------+--------+--------+--------+------|-+--------+--------+--------+--------+
'|  0xd7  |   -1   | nanosec. in 30-bit unsigned int |   seconds in 34-bit unsigned int    |
'+--------+--------+--------+--------+--------+------^-+--------+--------+--------+--------+
'* Timestamp 64 format can represent a timestamp in [1970-01-01 00:00:00.000000000 UTC, 2514-05-30 01:53:04.000000000 UTC) range.
Private Function GetExtTimeFromBytes_FixExt8( _
    Bytes() As Byte, Optional Index As Long) As Date
    
    Debug.Assert (Bytes(Index) = &HD7)
    Debug.Assert (Bytes(Index + 1) = mpTimestamp)
    
    Dim SecBytes(0 To 7) As Byte
    BitConverter.CopyBytes SecBytes, 0, Bytes, Index + 1 + 1, 8
    SecBytes(0) = 0
    SecBytes(1) = 0
    SecBytes(2) = 0
    SecBytes(3) = SecBytes(3) And &H3
    
    Dim Seconds As Double
    Seconds = _
        BitConverter.GetUInt64FromBytes(SecBytes, 0, True)
    
    'GetExtTimeFromBytes_FixExt8 = _
        DateAdd("s", Seconds, DateSerial(1970, 1, 1))
    
    Dim Days As Double
    Days = Int(Seconds / (CDbl(24) * 60 * 60))
    Seconds = Seconds - Days * (CDbl(24) * 60 * 60)
    
    GetExtTimeFromBytes_FixExt8 = _
        DateAdd("s", Seconds, DateAdd("d", Days, DateSerial(1970, 1, 1)))
End Function

'ext 8           | 11000111               | 0xc7
'timestamp 96 stores the number of seconds and nanoseconds that have elapsed since 1970-01-01 00:00:00 UTC
'in 64-bit signed integer and 32-bit unsigned integer:
'+--------+--------+--------+--------+--------+--------+--------+
'|  0xc7  |   12   |   -1   |nanoseconds in 32-bit unsigned int |
'+--------+--------+--------+--------+--------+--------+--------+
'+--------+--------+--------+--------+--------+--------+--------+--------+
'                    seconds in 64-bit signed int                        |
'+--------+--------+--------+--------+--------+--------+--------+--------+
'* Timestamp 96 format can represent a timestamp in [-292277022657-01-27 08:29:52 UTC, 292277026596-12-04 15:30:08.000000000 UTC) range.
'* In timestamp 64 and timestamp 96 formats, nanoseconds must not be larger than 999999999.

Private Function GetExtTimeFromBytes_Ext8( _
    Bytes() As Byte, Optional Index As Long) As Date
    
    Debug.Assert (Bytes(Index) = &HC7)
    Debug.Assert (Bytes(Index + 1) = 12)
    Debug.Assert (Bytes(Index + 2) = mpTimestamp)
    
    Dim Seconds As Double
    Seconds = _
        BitConverter.GetInt64FromBytes(Bytes, Index + 1 + 1 + 1 + 4, True)
    
    'GetExtTimeFromBytes_Ext8 = _
        DateAdd("s", Seconds, DateSerial(1970, 1, 1))
    
    Dim Days As Double
    Days = Int(Seconds / (CDbl(24) * 60 * 60))
    Seconds = Seconds - Days * (CDbl(24) * 60 * 60)
    
    GetExtTimeFromBytes_Ext8 = _
        DateAdd("s", Seconds, DateAdd("d", Days, DateSerial(1970, 1, 1)))
End Function

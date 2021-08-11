Attribute VB_Name = "MsgPack_Map"
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
'' MessagePack for VBA - Map
''

''
'' MessagePack for VBA - Map - Serialization
''

Public Function IsVBAMap(Value) As Boolean
    Select Case VarType(Value)
    Case vbObject
        IsVBAMap = (TypeName(Value) = "Dictionary")
        
    Case Else
        IsVBAMap = False
        
    End Select
End Function

Public Function GetBytesFromMap(Value) As Byte()
    Debug.Assert IsVBAMap(Value)
    
    Select Case Value.Count
    
    Case 0
        Dim Bytes(0) As Byte
        Bytes(0) = &H80
        GetBytesFromMap = Bytes
    
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H1 To &HF
        GetBytesFromMap = GetBytesFromMap_FixMap(Value)
        
    'map 16          | 11011110               | 0xde
    Case &H10 To &HFFFF&
        GetBytesFromMap = GetBytesFromMap_Map16(Value)
        
    'map 32          | 11011111               | 0xdf
    Case Else
        GetBytesFromMap = GetBytesFromMap_Map32(Value)
        
    End Select
End Function

'fixmap          | 1000xxxx               | 0x80 - 0x8f
'fixmap stores a map whose length is upto 15 elements
'+--------+~~~~~~~~~~~~~~~~~+
'|1000XXXX|   N*2 objects   |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value
Private Function GetBytesFromMap_FixMap(Value) As Byte()
    Debug.Assert IsVBAMap(Value)
    Debug.Assert (Value.Count <= &HF)
    
    Dim Bytes() As Byte
    ReDim Bytes(0)
    Bytes(0) = &H80 Or Value.Count
    
    Dim Keys
    Keys = Value.Keys
    
    Dim Index As Long
    For Index = LBound(Keys) To UBound(Keys)
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Keys(Index))
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Value.Item(Keys(Index)))
    Next
    
    GetBytesFromMap_FixMap = Bytes
End Function

'map 16          | 11011110               | 0xde
'map 16 stores a map whose length is upto (2^16)-1 elements
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xde  |YYYYYYYY|YYYYYYYY|   N*2 objects   |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value
Private Function GetBytesFromMap_Map16(Value) As Byte()
    Debug.Assert IsVBAMap(Value)
    Debug.Assert (Value.Count <= &HFFFF&)
    
    Dim Bytes() As Byte
    ReDim Bytes(0)
    Bytes(0) = &HDE
    MsgPack_Common.AddBytes Bytes, _
        BitConverter.GetBytesFromUInt16(Value.Count, True)
    
    Dim Keys
    Keys = Value.Keys
    
    Dim Index As Long
    For Index = LBound(Keys) To UBound(Keys)
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Keys(Index))
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Value.Item(Keys(Index)))
    Next
    
    GetBytesFromMap_Map16 = Bytes
End Function

'map 32          | 11011111               | 0xdf
'map 32 stores a map whose length is upto (2^32)-1 elements
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|   N*2 objects   |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value
Private Function GetBytesFromMap_Map32(Value) As Byte()
    Debug.Assert IsVBAMap(Value)
    
    Dim Bytes() As Byte
    ReDim Bytes(0)
    Bytes(0) = &HDF
    MsgPack_Common.AddBytes Bytes, _
        BitConverter.GetBytesFromUInt32(Value.Count, True)
    
    Dim Keys
    Keys = Value.Keys
    
    Dim Index As Long
    For Index = LBound(Keys) To UBound(Keys)
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Keys(Index))
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Value.Item(Keys(Index)))
    Next
    
    GetBytesFromMap_Map32 = Bytes
End Function

''
'' MessagePack for VBA - Map - Deserialization
''

Public Function IsMPMap(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    'map 16          | 11011110               | 0xde
    'map 32          | 11011111               | 0xdf
    Case &H80 To &H8F, &HDE, &HDF
        IsMPMap = True
        
    Case Else
        IsMPMap = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim ElementCount As Long
    Dim ElementLength As Long
    
    Select Case Bytes(Index)
    
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H80 To &H8F
        ElementCount = Bytes(Index) And &HF
        ElementLength = _
            GetLengthFromMapElementBytes(ElementCount, Bytes, Index + 1)
        GetLengthFromBytes = 1 + Length
        
    'map 16          | 11011110               | 0xde
    Case &HDE
        ElementCount = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
        ElementLength = _
            GetLengthFromMapElementBytes(ElementCount, Bytes, Index + 1 + 2)
        GetLengthFromBytes = 1 + 2 + Length
        
    'map 32          | 11011111               | 0xdf
    Case &HDF
        ElementCount = _
            CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
        ElementLength = _
            GetLengthFromMapElementBytes(ElementCount, Bytes, Index + 1 + 4)
        GetLengthFromBytes = 1 + 4 + Length
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Private Function GetLengthFromMapElementBytes(ByVal ElementCount As Long, _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Dim Count As Long
    For Count = 1 To ElementCount * 2
        Length = Length + MsgPack.GetLength(Bytes, Index + Length)
    Next
    
    GetLengthFromMapElementBytes = Length
End Function

Public Function GetMapFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Object
    
    Select Case Bytes(Index)
    
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H80 To &H8F
        Set GetMapFromBytes = GetMapFromBytes_FixMap(Bytes, Index)
        
    'map 16          | 11011110               | 0xde
    Case &HDE
        Set GetMapFromBytes = GetMapFromBytes_Map16(Bytes, Index)
        
    'map 32          | 11011111               | 0xdf
    Case &HDF
        Set GetMapFromBytes = GetMapFromBytes_Map32(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'fixmap          | 1000xxxx               | 0x80 - 0x8f
'fixmap stores a map whose length is upto 15 elements
'+--------+~~~~~~~~~~~~~~~~~+
'|1000XXXX|   N*2 objects   |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value
Private Function GetMapFromBytes_FixMap( _
    Bytes() As Byte, Optional Index As Long) As Object
    
    Debug.Assert ((Bytes(Index) And &HF0) = &H80)
    
    Dim ElementCount As Long
    ElementCount = Bytes(Index) And &HF
    
    Set GetMapFromBytes_FixMap = _
        GetMapFromBytes_Helper(Bytes, Index + 1, ElementCount)
End Function

'map 16          | 11011110               | 0xde
'map 16 stores a map whose length is upto (2^16)-1 elements
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xde  |YYYYYYYY|YYYYYYYY|   N*2 objects   |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value
Public Function GetMapFromBytes_Map16( _
    Bytes() As Byte, Optional Index As Long) As Object
    
    Debug.Assert (Bytes(Index) = &HDE)
    
    Dim ElementCount As Long
    ElementCount = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
    
    Set GetMapFromBytes_Map16 = _
        GetMapFromBytes_Helper(Bytes, Index + 1 + 2, ElementCount)
End Function

'map 32          | 11011111               | 0xdf
'map 32 stores a map whose length is upto (2^32)-1 elements
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdf  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|   N*2 objects   |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of a map
'* odd elements in objects are keys of a map
'* the next element of a key is its associated value
Public Function GetMapFromBytes_Map32( _
    Bytes() As Byte, Optional Index As Long) As Object
    
    Debug.Assert (Bytes(Index) = &HDF)
    
    Dim ElementCount As Long
    ElementCount = _
        CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
    
    Set GetMapFromBytes_Map32 = _
        GetMapFromBytes_Helper(Bytes, Index + 1 + 4, ElementCount)
End Function

''
'' MessagePack for VBA - Map - Deserialization - Helper
''

Private Function GetMapFromBytes_Helper( _
    Bytes() As Byte, Index As Long, ElementCount As Long) As Object
    
    Dim Map As Object
    Set Map = CreateObject("Scripting.Dictionary")
    
    Dim KeyOffset As Long
    Dim ValueOffset As Long
    
    Dim Count As Long
    For Count = 1 To ElementCount
        ValueOffset = KeyOffset + MsgPack.GetLength(Bytes, Index + KeyOffset)
        
        Map.Add _
            MsgPack.GetValue(Bytes, Index + KeyOffset), _
            MsgPack.GetValue(Bytes, Index + ValueOffset)
        
        KeyOffset = ValueOffset + MsgPack.GetLength(Bytes, Index + ValueOffset)
    Next
    
    Set GetMapFromBytes_Helper = Map
End Function

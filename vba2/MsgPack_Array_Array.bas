Attribute VB_Name = "MsgPack_Array_Array"
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
'' MessagePack for VBA - Array
''

''
'' MessagePack for VBA - Array - Serialization
''

Public Function IsVBAArray(Value) As Boolean
    IsVBAArray = IsArray(Value)
End Function

Public Function GetBytesFromArray(Value) As Byte()
    Debug.Assert IsVBAArray(Value)
    
    Dim Length As Long
    Length = GetLengthFromArray(Value)
    
    Select Case Length
    
    Case 0
        Dim Bytes(0) As Byte
        Bytes(0) = &H90
        GetBytesFromArray = Bytes
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H1 To &HF
        GetBytesFromArray = GetBytesFromArray_FixArray(Value)
        
    'array 16        | 11011100               | 0xdc
    Case &H10 To &HFFFF&
        GetBytesFromArray = GetBytesFromArray_Array16(Value)
        
    'array 32        | 11011101               | 0xdd
    Case Else
        GetBytesFromArray = GetBytesFromArray_Array32(Value)
        
    End Select
End Function

Private Function GetLengthFromArray(Value) As Long
    On Error Resume Next
    GetLengthFromArray = UBound(Value) - LBound(Value) + 1
    On Error GoTo 0
End Function

'fixarray        | 1001xxxx               | 0x90 - 0x9f
'fixarray stores an array whose length is upto 15 elements:
'+--------+~~~~~~~~~~~~~~~~~+
'|1001XXXX|    N objects    |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of an array
Private Function GetBytesFromArray_FixArray(Value) As Byte()
    Debug.Assert IsVBAArray(Value)
    
    Dim Count As Long
    Count = UBound(Value) - LBound(Value) + 1
    
    Debug.Assert (Count <= &HF)
    
    Dim Bytes() As Byte
    ReDim Bytes(0)
    Bytes(0) = &H90 Or Count
    
    Dim Index As Long
    For Index = LBound(Value) To UBound(Value)
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Value(Index))
    Next
    
    GetBytesFromArray_FixArray = Bytes
End Function

'array 16        | 11011100               | 0xdc
'array 16 stores an array whose length is upto (2^16)-1 elements:
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdc  |YYYYYYYY|YYYYYYYY|    N objects    |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of an array
Private Function GetBytesFromArray_Array16(Value) As Byte()
    Debug.Assert IsVBAArray(Value)
    
    Dim Count As Long
    Count = UBound(Value) - LBound(Value) + 1
    
    Debug.Assert (Count <= &HFFFF&)
    
    Dim Bytes() As Byte
    ReDim Bytes(0)
    Bytes(0) = &HDC
    MsgPack_Common.AddBytes Bytes, _
        BitConverter.GetBytesFromUInt16(Count, True)
    
    Dim Index As Long
    For Index = LBound(Value) To UBound(Value)
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Value(Index))
    Next
    
    GetBytesFromArray_Array16 = Bytes
End Function

'array 32        | 11011101               | 0xdd
'array 32 stores an array whose length is upto (2^32)-1 elements:
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdd  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|    N objects    |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of an array
Private Function GetBytesFromArray_Array32(Value) As Byte()
    Debug.Assert IsVBAArray(Value)
    
    Dim Count As Long
    Count = UBound(Value) - LBound(Value) + 1
    
    Dim Bytes() As Byte
    ReDim Bytes(0)
    Bytes(0) = &HDD
    MsgPack_Common.AddBytes Bytes, _
        BitConverter.GetBytesFromUInt32(Count, True)
    
    Dim Index As Long
    For Index = LBound(Value) To UBound(Value)
        MsgPack_Common.AddBytes Bytes, _
            MsgPack.GetBytes(Value(Index))
    Next
    
    GetBytesFromArray_Array32 = Bytes
End Function

''
'' MessagePack for VBA - Array - Deserialization
''

Public Function IsMPArray(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H90 To &H9F
        IsMPArray = True
        
    'array 16        | 11011100               | 0xdc
    'array 32        | 11011101               | 0xdd
    Case &HDC, &HDD
        IsMPArray = True
        
    Case Else
        IsMPArray = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim ElementCount As Long
    Dim ElementLength As Long
    
    Select Case Bytes(Index)
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H90 To &H9F
        ElementCount = Bytes(Index) And &HF
        ElementLength = _
            GetLengthFromArrayElementBytes(ElementCount, Bytes, Index + 1)
        GetLengthFromBytes = 1 + Length
        
    'array 16        | 11011100               | 0xdc
    Case &HDC
        ElementCount = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
        ElementLength = _
            GetLengthFromArrayElementBytes(ElementCount, Bytes, Index + 1 + 2)
        GetLengthFromBytes = 1 + 2 + Length
        
    'array 32        | 11011101               | 0xdd
    Case &HDD
        ElementCount = _
            CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
        ElementLength = _
            GetLengthFromArrayElementBytes(ElementCount, Bytes, Index + 1 + 4)
        GetLengthFromBytes = 1 + 4 + Length
        
    Case Else
        Err.Raise 13 ' type mismatch
        
    End Select
End Function

Private Function GetLengthFromArrayElementBytes(ByVal ElementCount As Long, _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Dim Count As Long
    For Count = 1 To ElementCount
        Length = Length + MsgPack.GetLength(Bytes, Index + Length)
    Next
    
    GetLengthFromArrayElementBytes = Length
End Function

Public Function GetArrayFromBytes(Bytes() As Byte, Optional Index As Long)
    Select Case Bytes(Index)
    
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H90 To &H9F
        GetArrayFromBytes = GetArrayFromBytes_FixArray(Bytes, Index)
        
    'array 16        | 11011100               | 0xdc
    Case &HDC
        GetArrayFromBytes = GetArrayFromBytes_Array16(Bytes, Index)
        
    'array 32        | 11011101               | 0xdd
    Case &HDD
        GetArrayFromBytes = GetArrayFromBytes_Array32(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' type mismatch
        
    End Select
End Function

'fixarray        | 1001xxxx               | 0x90 - 0x9f
'fixarray stores an array whose length is upto 15 elements:
'+--------+~~~~~~~~~~~~~~~~~+
'|1001XXXX|    N objects    |
'+--------+~~~~~~~~~~~~~~~~~+
'* XXXX is a 4-bit unsigned integer which represents N
'* N is the size of an array
Private Function GetArrayFromBytes_FixArray( _
    Bytes() As Byte, Optional Index As Long)
    
    Debug.Assert ((Bytes(Index) And &HF0) = &H90)
    
    Dim ElementCount As Long
    ElementCount = Bytes(Index) And &HF
    
    GetArrayFromBytes_FixArray = _
        GetArrayFromBytes_Helper(Bytes, Index + 1, ElementCount)
End Function

'array 16        | 11011100               | 0xdc
'array 16 stores an array whose length is upto (2^16)-1 elements:
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdc  |YYYYYYYY|YYYYYYYY|    N objects    |
'+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* YYYYYYYY_YYYYYYYY is a 16-bit big-endian unsigned integer which represents N
'* N is the size of an array
Private Function GetArrayFromBytes_Array16( _
    Bytes() As Byte, Optional Index As Long)
    
    Debug.Assert (Bytes(Index) = &HDC)
    
    Dim ElementCount As Long
    ElementCount = BitConverter.GetUInt16FromBytes(Bytes, Index + 1, True)
    
    GetArrayFromBytes_Array16 = _
        GetArrayFromBytes_Helper(Bytes, Index + 1 + 2, ElementCount)
End Function

'array 32        | 11011101               | 0xdd
'array 32 stores an array whose length is upto (2^32)-1 elements:
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'|  0xdd  |ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|ZZZZZZZZ|    N objects    |
'+--------+--------+--------+--------+--------+~~~~~~~~~~~~~~~~~+
'* ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ_ZZZZZZZZ is a 32-bit big-endian unsigned integer which represents N
'* N is the size of an array
Private Function GetArrayFromBytes_Array32( _
    Bytes() As Byte, Optional Index As Long)
    
    Debug.Assert (Bytes(Index) = &HDD)
    
    Dim ElementCount As Long
    ElementCount = _
        CLng(BitConverter.GetUInt32FromBytes(Bytes, Index + 1, True))
    
    GetArrayFromBytes_Array32 = _
        GetArrayFromBytes_Helper(Bytes, Index + 1 + 4, ElementCount)
End Function

''
'' MessagePack for VBA - Array - Deserialization - Helper
''

Private Function GetArrayFromBytes_Helper( _
    Bytes() As Byte, Index As Long, ElementCount As Long)
    
    Dim Array_()
    
    If ElementCount = 0 Then
        GetArrayFromBytes_Helper = Array_
        Exit Function
    End If
    
    ReDim Array_(0 To ElementCount - 1)
    
    Dim Offset As Long
    Dim Count As Long
    For Count = 0 To ElementCount - 1
        If IsObject(MsgPack.GetValue(Bytes, Index + Offset)) Then
            Set Array_(Count) = MsgPack.GetValue(Bytes, Index + Offset)
        Else
            Array_(Count) = MsgPack.GetValue(Bytes, Index + Offset)
        End If
        
        Offset = Offset + MsgPack.GetLength(Bytes, Index + Offset)
    Next
    
    GetArrayFromBytes_Helper = Array_
End Function

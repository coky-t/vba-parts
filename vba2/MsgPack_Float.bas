Attribute VB_Name = "MsgPack_Float"
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
'' MessagePack for VBA - Float
''

''
'' MessagePack for VBA - Float - Serialization
''

Public Function IsVBAFloat(Value) As Boolean
    Select Case VarType(Value)
    Case vbSingle, vbDouble
        IsVBAFloat = True
        
    Case Else
        IsVBAFloat = False
        
    End Select
End Function

Public Function GetBytesFromFloat(Value) As Byte()
    Debug.Assert IsVBAFloat(Value)
    
    Select Case VarType(Value)
    
    Case vbSingle
        GetBytesFromFloat = GetBytesFromFloat32(Value)
        
    Case vbDouble
        GetBytesFromFloat = GetBytesFromFloat64(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'float 32        | 11001010               | 0xca
'float 32 stores a floating point number in IEEE 754 single precision floating point number format:
'+--------+--------+--------+--------+--------+
'|  0xca  |XXXXXXXX|XXXXXXXX|XXXXXXXX|XXXXXXXX|
'+--------+--------+--------+--------+--------+
'* XXXXXXXX_XXXXXXXX_XXXXXXXX_XXXXXXXX is a big-endian IEEE 754 single precision floating point number.
'  Extension of precision from single-precision to double-precision does not lose precision.
Public Function GetBytesFromFloat32(ByVal Value As Single) As Byte()
    GetBytesFromFloat32 = _
        MsgPack_Common.GetBytesHelper1(&HCA, _
            BitConverter.GetBytesFromFloat32(Value, True))
End Function

'float 64        | 11001011               | 0xcb
'float 64 stores a floating point number in IEEE 754 double precision floating point number format:
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xcb  |YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY is a big-endian
'  IEEE 754 double precision floating point number
Public Function GetBytesFromFloat64(ByVal Value As Double) As Byte()
    GetBytesFromFloat64 = _
        MsgPack_Common.GetBytesHelper1(&HCB, _
            BitConverter.GetBytesFromFloat64(Value, True))
End Function

''
'' MessagePack for VBA - Float - Deserialization
''

Public Function IsMPFloat(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'float 32        | 11001010               | 0xca
    'float 64        | 11001011               | 0xcb
    Case &HCA, &HCB
        IsMPFloat = True
        
    Case Else
        IsMPFloat = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Select Case Bytes(Index)
    
    'float 32        | 11001010               | 0xca
    Case &HCA
        GetLengthFromBytes = 1 + 4
        
    'float 64        | 11001011               | 0xcb
    Case &HCB
        GetLengthFromBytes = 1 + 8
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetFloatFromBytes(Bytes() As Byte, Optional Index As Long)
    Select Case Bytes(Index)
    
    'float 32        | 11001010               | 0xca
    'float 32 stores a floating point number in IEEE 754 single precision floating point number format:
    '+--------+--------+--------+--------+--------+
    '|  0xca  |XXXXXXXX|XXXXXXXX|XXXXXXXX|XXXXXXXX|
    '+--------+--------+--------+--------+--------+
    '* XXXXXXXX_XXXXXXXX_XXXXXXXX_XXXXXXXX is a big-endian IEEE 754 single precision floating point number.
    '  Extension of precision from single-precision to double-precision does not lose precision.
    Case &HCA
        GetFloatFromBytes = _
            BitConverter.GetFloat32FromBytes(Bytes, Index + 1, True)
        
    'float 64        | 11001011               | 0xcb
    'float 64 stores a floating point number in IEEE 754 double precision floating point number format:
    '+--------+--------+--------+--------+--------+--------+--------+--------+--------+
    '|  0xcb  |YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|YYYYYYYY|
    '+--------+--------+--------+--------+--------+--------+--------+--------+--------+
    '* YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY_YYYYYYYY is a big-endian
    '  IEEE 754 double precision floating point number
    Case &HCB
        GetFloatFromBytes = _
            BitConverter.GetFloat64FromBytes(Bytes, Index + 1, True)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

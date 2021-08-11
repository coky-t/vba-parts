Attribute VB_Name = "MsgPack_Ext_Dec"
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
'' MessagePack for VBA - Extension - Decimal
''

Private Property Get mpDecimal() As Long
    mpDecimal = vbDecimal
End Property

''
'' MessagePack for VBA - Extension - Decimal - Serialization
''

Public Function IsVBAExtDec(Value) As Boolean
    Select Case VarType(Value)
    Case vbDecimal
        IsVBAExtDec = True
        
    Case Else
        IsVBAExtDec = False
        
    End Select
End Function

Public Function GetBytesFromExtDec(Value) As Byte()
    Debug.Assert IsVBAExtDec(Value)
    
    Dim DecBytesBE() As Byte
    DecBytesBE = BitConverter.GetBytesFromDecimal(Value, True)
    
    Dim DecBytesBETemp() As Byte
    
    If DecBytesBE(0) = 0 And DecBytesBE(1) = 0 Then ' Positive, No scaling
        If DecBytesBE(2) = 0 And DecBytesBE(3) = 0 And _
            DecBytesBE(4) = 0 And DecBytesBE(5) = 0 Then ' High Bytes
        
            If DecBytesBE(6) = 0 And DecBytesBE(7) = 0 And _
                DecBytesBE(8) = 0 And DecBytesBE(9) = 0 Then ' Middle Bytes
                
                If DecBytesBE(10) = 0 And DecBytesBE(11) = 0 Then
                    If DecBytesBE(12) = 0 Then
                        ReDim DecBytesBETemp(0)
                        DecBytesBETemp(0) = DecBytesBE(13)
                    Else
                        ReDim DecBytesBETemp(0 To 1)
                        BitConverter.CopyBytes _
                            DecBytesBETemp, 0, DecBytesBE, 12, 2
                    End If
                Else
                    ReDim DecBytesBETemp(0 To 3)
                    BitConverter.CopyBytes _
                        DecBytesBETemp, 0, DecBytesBE, 10, 4
                End If
            Else
                ReDim DecBytesBETemp(0 To 7)
                BitConverter.CopyBytes DecBytesBETemp, 0, DecBytesBE, 6, 8
            End If
        Else
            ReDim DecBytesBETemp(0 To 11)
            BitConverter.CopyBytes DecBytesBETemp, 0, DecBytesBE, 2, 12
        End If
    Else
        DecBytesBETemp = DecBytesBE
    End If
    
    GetBytesFromExtDec = _
        MsgPack_Ext.GetBytesFromExt(mpDecimal, DecBytesBETemp)
End Function

''
'' MessagePack for VBA - Extension - Decimal - Deserialization
''

Public Function IsMPExtDec(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'ext 16          | 11001000               | 0xc8
    'ext 32          | 11001001               | 0xc9
    'fixext 16       | 11011000               | 0xd8
    Case &HC8, &HC9, &HD8
        IsMPExtDec = False
        
    'ext 8           | 11000111               | 0xc7
    'fixext 1        | 11010100               | 0xd4
    'fixext 2        | 11010101               | 0xd5
    'fixext 4        | 11010110               | 0xd6
    'fixext 8        | 11010111               | 0xd7
    Case &HC7, &HD4 To &HD7
        IsMPExtDec = (Bytes(Index + 1) = mpDecimal)
        
    Case Else
        IsMPExtDec = False
        
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
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetExtDecFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Variant
    
    Select Case Bytes(Index)
    
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        GetExtDecFromBytes = GetExtDecFromBytes_Ext8(Bytes, Index)
        
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetExtDecFromBytes = GetExtDecFromBytes_FixExt1(Bytes, Index)
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetExtDecFromBytes = GetExtDecFromBytes_FixExt2(Bytes, Index)
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetExtDecFromBytes = GetExtDecFromBytes_FixExt4(Bytes, Index)
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetExtDecFromBytes = GetExtDecFromBytes_FixExt8(Bytes, Index)
        
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
Private Function GetExtDecFromBytes_Ext8( _
    Bytes() As Byte, Optional Index As Long) As Variant
    
    Debug.Assert (Bytes(Index) = &HC7)
    Debug.Assert (Bytes(Index + 2) = mpDecimal)
    
    Dim Length As Byte
    Length = Bytes(Index + 1)
    If Length = 0 Then
        GetExtDecFromBytes_Ext8 = CDec(0)
        Exit Function
    End If
    
    Dim DecBytesBE() As Byte
    ReDim DecBytesBE(0 To 13)
    BitConverter.CopyBytes _
        DecBytesBE, 14 - Length, Bytes, Index + 1 + 1 + 1, Length
    
    GetExtDecFromBytes_Ext8 = _
        BitConverter.GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtDecFromBytes_FixExt1( _
    Bytes() As Byte, Optional Index As Long) As Variant
    
    Debug.Assert (Bytes(Index) = &HD4)
    Debug.Assert (Bytes(Index + 1) = mpDecimal)
    
    Dim DecBytesBE(0 To 13) As Byte
    DecBytesBE(13) = Bytes(Index + 1 + 1)
    
    GetExtDecFromBytes_FixExt1 = _
        BitConverter.GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtDecFromBytes_FixExt2( _
    Bytes() As Byte, Optional Index As Long) As Variant
    
    Debug.Assert (Bytes(Index) = &HD5)
    Debug.Assert (Bytes(Index + 1) = mpDecimal)
    
    Dim DecBytesBE(0 To 13) As Byte
    BitConverter.CopyBytes DecBytesBE, 12, Bytes, Index + 1 + 1, 2
    
    GetExtDecFromBytes_FixExt2 = _
        BitConverter.GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtDecFromBytes_FixExt4( _
    Bytes() As Byte, Optional Index As Long) As Variant
    
    Debug.Assert (Bytes(Index) = &HD6)
    Debug.Assert (Bytes(Index + 1) = mpDecimal)
    
    Dim DecBytesBE(0 To 13) As Byte
    BitConverter.CopyBytes DecBytesBE, 10, Bytes, Index + 1 + 1, 4
    
    GetExtDecFromBytes_FixExt4 = _
        BitConverter.GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtDecFromBytes_FixExt8( _
    Bytes() As Byte, Optional Index As Long) As Variant
    
    Debug.Assert (Bytes(Index) = &HD7)
    Debug.Assert (Bytes(Index + 1) = mpDecimal)
    
    Dim DecBytesBE(0 To 13) As Byte
    BitConverter.CopyBytes DecBytesBE, 6, Bytes, Index + 1 + 1, 8
    
    GetExtDecFromBytes_FixExt8 = _
        BitConverter.GetDecimalFromBytes(DecBytesBE, 0, True)
End Function

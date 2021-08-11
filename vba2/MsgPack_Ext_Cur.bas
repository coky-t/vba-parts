Attribute VB_Name = "MsgPack_Ext_Cur"
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
'' MessagePack for VBA - Extension - Currency
''

Private Property Get mpCurrency() As Long
    mpCurrency = vbCurrency
End Property

''
'' MessagePack for VBA - Extension - Currency - Serialization
''

Public Function IsVBAExtCur(Value) As Boolean
    Select Case VarType(Value)
    Case vbCurrency
        IsVBAExtCur = True
        
    Case Else
        IsVBAExtCur = False
        
    End Select
End Function

Public Function GetBytesFromExtCur(Value) As Byte()
    Debug.Assert IsVBAExtCur(Value)
    
    Dim CurBytesBE() As Byte
    CurBytesBE = BitConverter.GetBytesFromCurrency(Value, True)
    
    Dim CurBytesBETemp() As Byte
    
    If CurBytesBE(0) = 0 And CurBytesBE(1) = 0 And _
        CurBytesBE(2) = 0 And CurBytesBE(3) = 0 Then
        
        If CurBytesBE(4) = 0 And CurBytesBE(5) = 0 Then
            If CurBytesBE(6) = 0 Then
                ReDim CurBytesBETemp(0)
                CurBytesBETemp(0) = CurBytesBE(7)
            Else
                ReDim CurBytesBETemp(0 To 1)
                BitConverter.CopyBytes CurBytesBETemp, 0, CurBytesBE, 6, 2
            End If
        Else
            ReDim CurBytesBETemp(0 To 3)
            BitConverter.CopyBytes CurBytesBETemp, 0, CurBytesBE, 4, 4
        End If
    Else
        CurBytesBETemp = CurBytesBE
    End If
    
    GetBytesFromExtCur = _
        MsgPack_Ext.GetBytesFromExt(mpCurrency, CurBytesBETemp)
End Function

''
'' MessagePack for VBA - Extension - Currency - Deserialization
''

Public Function IsMPExtCur(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'ext 8           | 11000111               | 0xc7
    'ext 16          | 11001000               | 0xc8
    'ext 32          | 11001001               | 0xc9
    'fixext 16       | 11011000               | 0xd8
    Case &HC7 To &HC9, &HD8
        IsMPExtCur = False
        
    'fixext 1        | 11010100               | 0xd4
    'fixext 2        | 11010101               | 0xd5
    'fixext 4        | 11010110               | 0xd6
    'fixext 8        | 11010111               | 0xd7
    Case &HD4 To &HD7
        IsMPExtCur = (Bytes(Index + 1) = mpCurrency)
        
    Case Else
        IsMPExtCur = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Select Case Bytes(Index)
    
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

Public Function GetExtCurFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Currency
    
    Select Case Bytes(Index)
    
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetExtCurFromBytes = GetExtCurFromBytes_FixExt1(Bytes, Index)
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetExtCurFromBytes = GetExtCurFromBytes_FixExt2(Bytes, Index)
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetExtCurFromBytes = GetExtCurFromBytes_FixExt4(Bytes, Index)
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetExtCurFromBytes = GetExtCurFromBytes_FixExt8(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'fixext 1        | 11010100               | 0xd4
'fixext 1 stores an integer and a byte array whose length is 1 byte
'+--------+--------+--------+
'|  0xd4  |  type  |  data  |
'+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtCurFromBytes_FixExt1( _
    Bytes() As Byte, Optional Index As Long) As Currency
    
    Debug.Assert (Bytes(Index) = &HD4)
    Debug.Assert (Bytes(Index + 1) = mpCurrency)
    
    Dim CurBytesBE(0 To 7) As Byte
    CurBytesBE(7) = Bytes(Index + 1 + 1)
    
    GetExtCurFromBytes_FixExt1 = _
        BitConverter.GetCurrencyFromBytes(CurBytesBE, 0, True)
End Function

'fixext 2        | 11010101               | 0xd5
'fixext 2 stores an integer and a byte array whose length is 2 bytes
'+--------+--------+--------+--------+
'|  0xd5  |  type  |       data      |
'+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtCurFromBytes_FixExt2( _
    Bytes() As Byte, Optional Index As Long) As Currency
    
    Debug.Assert (Bytes(Index) = &HD5)
    Debug.Assert (Bytes(Index + 1) = mpCurrency)
    
    Dim CurBytesBE(0 To 7) As Byte
    BitConverter.CopyBytes CurBytesBE, 6, Bytes, Index + 1 + 1, 2
    
    GetExtCurFromBytes_FixExt2 = _
        BitConverter.GetCurrencyFromBytes(CurBytesBE, 0, True)
End Function

'fixext 4        | 11010110               | 0xd6
'fixext 4 stores an integer and a byte array whose length is 4 bytes
'+--------+--------+--------+--------+--------+--------+
'|  0xd6  |  type  |                data               |
'+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtCurFromBytes_FixExt4( _
    Bytes() As Byte, Optional Index As Long) As Currency
    
    Debug.Assert (Bytes(Index) = &HD6)
    Debug.Assert (Bytes(Index + 1) = mpCurrency)
    
    Dim CurBytesBE(0 To 7) As Byte
    BitConverter.CopyBytes CurBytesBE, 4, Bytes, Index + 1 + 1, 4
    
    GetExtCurFromBytes_FixExt4 = _
        BitConverter.GetCurrencyFromBytes(CurBytesBE, 0, True)
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtCurFromBytes_FixExt8( _
    Bytes() As Byte, Optional Index As Long) As Currency
    
    Debug.Assert (Bytes(Index) = &HD7)
    Debug.Assert (Bytes(Index + 1) = mpCurrency)
    
    GetExtCurFromBytes_FixExt8 = _
        BitConverter.GetCurrencyFromBytes(Bytes, Index + 1 + 1, True)
End Function

Attribute VB_Name = "MsgPack_Ext_Date"
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
'' MessagePack for VBA - Extension - Date
''

Private Property Get mpDate() As Long
    mpDate = vbDate
End Property

''
'' MessagePack for VBA - Extension - Date - Serialization
''

Public Function IsVBAExtDate(Value) As Boolean
    Select Case VarType(Value)
    Case vbDate
        IsVBAExtDate = True
        
    Case Else
        IsVBAExtDate = False
        
    End Select
End Function

Public Function GetBytesFromExtDate(Value) As Byte()
    Debug.Assert IsVBAExtDate(Value)
    
    GetBytesFromExtDate = _
        MsgPack_Ext.GetBytesFromExt(mpDate, _
            BitConverter.GetBytesFromDate(Value, True))
End Function

''
'' MessagePack for VBA - Extension - Date - Deserialization
''

Public Function IsMPExtDate(Bytes() As Byte, Optional Index As Long) As Boolean
    Select Case Bytes(Index)
    
    'ext 8           | 11000111               | 0xc7
    'ext 16          | 11001000               | 0xc8
    'ext 32          | 11001001               | 0xc9
    'fixext 1        | 11010100               | 0xd4
    'fixext 2        | 11010101               | 0xd5
    'fixext 4        | 11010110               | 0xd6
    'fixext 16       | 11011000               | 0xd8
    Case &HC7 To &HC9, &HD4 To &HD6, &HD8
        IsMPExtDate = False
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        IsMPExtDate = (Bytes(Index + 1) = mpDate)
        
    Case Else
        IsMPExtDate = False
        
    End Select
End Function

Public Function GetLengthFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Select Case Bytes(Index)
    
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetLengthFromBytes = 1 + 1 + 8
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetExtDateFromBytes( _
    Bytes() As Byte, Optional Index As Long) As Date
    
    Select Case Bytes(Index)
    
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetExtDateFromBytes = GetExtDateFromBytes_FixExt8(Bytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'fixext 8        | 11010111               | 0xd7
'fixext 8 stores an integer and a byte array whose length is 8 bytes
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'|  0xd7  |  type  |                                  data                                 |
'+--------+--------+--------+--------+--------+--------+--------+--------+--------+--------+
'* type is a signed 8-bit signed integer
'* type < 0 is reserved for future extension including 2-byte type information
Private Function GetExtDateFromBytes_FixExt8( _
    Bytes() As Byte, Optional Index As Long) As Date
    
    Debug.Assert (Bytes(Index) = &HD7)
    Debug.Assert (Bytes(Index + 1) = mpDate)
    
    GetExtDateFromBytes_FixExt8 = _
        BitConverter.GetDateFromBytes(Bytes, Index + 1 + 1, True)
End Function

Attribute VB_Name = "MsgPack"
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
'' MessagePack for VBA
''

''
'' MessagePack for VBA - Serialization
''

Public Function GetBytes(Value) As Byte()
    Dim Bytes() As Byte
    
    Select Case VarType(Value)
    
    Case vbEmpty, vbNull
        ReDim Bytes(0) As Byte
        Bytes(0) = &HC0
        GetBytes = Bytes
        
    Case vbInteger, vbLong
        GetBytes = MsgPack_Int.GetBytesFromInt(Value)
        
    Case vbSingle, vbDouble
        GetBytes = MsgPack_Float.GetBytesFromFloat(Value)
        
    Case vbCurrency
        GetBytes = MsgPack_Ext_Cur.GetBytesFromExtCur(Value)
        
    Case vbDate
        GetBytes = MsgPack_Ext_Date.GetBytesFromExtDate(Value)
        
    Case vbString
        GetBytes = MsgPack_Str.GetBytesFromStr(Value)
        
    Case vbObject
        Select Case TypeName(Value)
        Case "Dictionary"
            GetBytes = MsgPack_Map.GetBytesFromMap(Value)
            
        Case "Collection"
            GetBytes = MsgPack_Array.GetBytesFromArray(Value)
            
        Case Else
            Err.Raise 13 ' unmatched type
            
        End Select
        
    Case vbError
        ReDim Bytes(0) As Byte
        Bytes(0) = &HC0
        GetBytes = Bytes
        
    Case vbBoolean
        ReDim Bytes(0) As Byte
        Bytes(0) = IIf(Value, &HC3, &HC2)
        GetBytes = Bytes
        
    Case vbVariant
        ' to do
        Err.Raise 13 ' unmatched type
        
    Case vbDecimal
        GetBytes = MsgPack_Ext_Dec.GetBytesFromExtDec(Value)
        
    Case vbByte
        GetBytes = MsgPack_Int.GetBytesFromInt(Value)
        
    #If Win64 Then
    Case vbLongLong
        GetBytes = MsgPack_Int.GetBytesFromInt(Value)
    #End If
        
    Case vbArray + vbByte
        GetBytes = MsgPack_Bin.GetBytesFromBin(Value)
        
    Case Else
        If IsArray(Value) Then
            GetBytes = MsgPack_Array.GetBytesFromArray(Value)
        Else
            Err.Raise 13 ' unmatched type
        End If
        
    End Select
End Function

''
'' MessagePack for VBA - Deserialization
''

Public Function GetLength(Bytes() As Byte, Optional Index As Long) As Long
    Select Case Bytes(Index)
    
    'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
    Case &H0 To &H7F
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H80 To &H8F
        GetLength = MsgPack_Map.GetLengthFromBytes(Bytes, Index)
        
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H90 To &H9F
        GetLength = MsgPack_Array.GetLengthFromBytes(Bytes, Index)
        
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &HA0 To &HBF
        GetLength = MsgPack_Str.GetLengthFromBytes(Bytes, Index)
        
    'nil             | 11000000               | 0xc0
    Case &HC0
        GetLength = 1
        
    '(never used)    | 11000001               | 0xc1
    Case &HC1
        GetLength = 1
        
    'false           | 11000010               | 0xc2
    Case &HC2
        GetLength = 1
        
    'true            | 11000011               | 0xc3
    Case &HC3
        GetLength = 1
        
    'bin 8           | 11000100               | 0xc4
    Case &HC4
        GetLength = MsgPack_Bin.GetLengthFromBytes(Bytes, Index)
        
    'bin 16          | 11000101               | 0xc5
    Case &HC5
        GetLength = MsgPack_Bin.GetLengthFromBytes(Bytes, Index)
        
    'bin 32          | 11000110               | 0xc6
    Case &HC6
        GetLength = MsgPack_Bin.GetLengthFromBytes(Bytes, Index)
        
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'ext 16          | 11001000               | 0xc8
    Case &HC8
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'ext 32          | 11001001               | 0xc9
    Case &HC9
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'float 32        | 11001010               | 0xca
    Case &HCA
        GetLength = MsgPack_Float.GetLengthFromBytes(Bytes, Index)
        
    'float 64        | 11001011               | 0xcb
    Case &HCB
        GetLength = MsgPack_Float.GetLengthFromBytes(Bytes, Index)
        
    'uint 8          | 11001100               | 0xcc
    Case &HCC
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'uint 16         | 11001101               | 0xcd
    Case &HCD
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'uint 32         | 11001110               | 0xce
    Case &HCE
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'uint 64         | 11001111               | 0xcf
    Case &HCF
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'int 8           | 11010000               | 0xd0
    Case &HD0
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'int 16          | 11010001               | 0xd1
    Case &HD1
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'int 32          | 11010010               | 0xd2
    Case &HD2
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'int 64          | 11010011               | 0xd3
    Case &HD3
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'fixext 16       | 11011000               | 0xd8
    Case &HD8
        GetLength = MsgPack_Ext.GetLengthFromBytes(Bytes, Index)
        
    'str 8           | 11011001               | 0xd9
    Case &HD9
        GetLength = MsgPack_Str.GetLengthFromBytes(Bytes, Index)
        
    'str 16          | 11011010               | 0xda
    Case &HDA
        GetLength = MsgPack_Str.GetLengthFromBytes(Bytes, Index)
        
    'str 32          | 11011011               | 0xdb
    Case &HDB
        GetLength = MsgPack_Str.GetLengthFromBytes(Bytes, Index)
        
    'array 16        | 11011100               | 0xdc
    Case &HDC
        GetLength = MsgPack_Array.GetLengthFromBytes(Bytes, Index)
        
    'array 32        | 11011101               | 0xdd
    Case &HDD
        GetLength = MsgPack_Array.GetLengthFromBytes(Bytes, Index)
        
    'map 16          | 11011110               | 0xde
    Case &HDE
        GetLength = MsgPack_Map.GetLengthFromBytes(Bytes, Index)
        
    'map 32          | 11011111               | 0xdf
    Case &HDF
        GetLength = MsgPack_Map.GetLengthFromBytes(Bytes, Index)
        
    'negative fixint | 111xxxxx               | 0xe0 - 0xff
    Case &HE0 To &HFF
        GetLength = MsgPack_Int.GetLengthFromBytes(Bytes, Index)
        
    End Select
End Function

Public Function GetValue(Bytes() As Byte, Optional Index As Long) As Variant
    Select Case Bytes(Index)
    
    'positive fixint | 0xxxxxxx               | 0x00 - 0x7f
    Case &H0 To &H7F
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'fixmap          | 1000xxxx               | 0x80 - 0x8f
    Case &H80 To &H8F
        GetValue = MsgPack_Map.GetMapFromBytes(Bytes, Index)
        
    'fixarray        | 1001xxxx               | 0x90 - 0x9f
    Case &H90 To &H9F
        GetValue = MsgPack_Array.GetArrayFromBytes(Bytes, Index)
        
    'fixstr          | 101xxxxx               | 0xa0 - 0xbf
    Case &HA0 To &HBF
        GetValue = MsgPack_Str.GetStrFromBytes(Bytes, Index)
        
    'nil             | 11000000               | 0xc0
    Case &HC0
        GetValue = Null
        
    '(never used)    | 11000001               | 0xc1
    Case &HC1
        GetValue = Empty
        
    'false           | 11000010               | 0xc2
    Case &HC2
        GetValue = False
        
    'true            | 11000011               | 0xc3
    Case &HC3
        GetValue = True
        
    'bin 8           | 11000100               | 0xc4
    Case &HC4
        GetValue = MsgPack_Bin.GetBinFromBytes(Bytes, Index)
        
    'bin 16          | 11000101               | 0xc5
    Case &HC5
        GetValue = MsgPack_Bin.GetBinFromBytes(Bytes, Index)
        
    'bin 32          | 11000110               | 0xc6
    Case &HC6
        GetValue = MsgPack_Bin.GetBinFromBytes(Bytes, Index)
        
    'ext 8           | 11000111               | 0xc7
    Case &HC7
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'ext 16          | 11001000               | 0xc8
    Case &HC8
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'ext 32          | 11001001               | 0xc9
    Case &HC9
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'float 32        | 11001010               | 0xca
    Case &HCA
        GetValue = MsgPack_Float.GetFloatFromBytes(Bytes, Index)
        
    'float 64        | 11001011               | 0xcb
    Case &HCB
        GetValue = MsgPack_Float.GetFloatFromBytes(Bytes, Index)
        
    'uint 8          | 11001100               | 0xcc
    Case &HCC
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'uint 16         | 11001101               | 0xcd
    Case &HCD
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'uint 32         | 11001110               | 0xce
    Case &HCE
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'uint 64         | 11001111               | 0xcf
    Case &HCF
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'int 8           | 11010000               | 0xd0
    Case &HD0
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'int 16          | 11010001               | 0xd1
    Case &HD1
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'int 32          | 11010010               | 0xd2
    Case &HD2
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'int 64          | 11010011               | 0xd3
    Case &HD3
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    'fixext 1        | 11010100               | 0xd4
    Case &HD4
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'fixext 2        | 11010101               | 0xd5
    Case &HD5
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'fixext 4        | 11010110               | 0xd6
    Case &HD6
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'fixext 8        | 11010111               | 0xd7
    Case &HD7
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'fixext 16       | 11011000               | 0xd8
    Case &HD8
        GetValue = MsgPack_Ext.GetExtFromBytes(Bytes, Index)
        
    'str 8           | 11011001               | 0xd9
    Case &HD9
        GetValue = MsgPack_Str.GetStrFromBytes(Bytes, Index)
        
    'str 16          | 11011010               | 0xda
    Case &HDA
        GetValue = MsgPack_Str.GetStrFromBytes(Bytes, Index)
        
    'str 32          | 11011011               | 0xdb
    Case &HDB
        GetValue = MsgPack_Str.GetStrFromBytes(Bytes, Index)
        
    'array 16        | 11011100               | 0xdc
    Case &HDC
        GetValue = MsgPack_Array.GetArrayFromBytes(Bytes, Index)
        
    'array 32        | 11011101               | 0xdd
    Case &HDD
        GetValue = MsgPack_Array.GetArrayFromBytes(Bytes, Index)
        
    'map 16          | 11011110               | 0xde
    Case &HDE
        GetValue = MsgPack_Map.GetMapFromBytes(Bytes, Index)
        
    'map 32          | 11011111               | 0xdf
    Case &HDF
        GetValue = MsgPack_Map.GetMapFromBytes(Bytes, Index)
        
    'negative fixint | 111xxxxx               | 0xe0 - 0xff
    Case &HE0 To &HFF
        GetValue = MsgPack_Int.GetIntFromBytes(Bytes, Index)
        
    End Select
End Function

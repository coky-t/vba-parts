Attribute VB_Name = "BitConverter"
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
'' BitConverter
''

'
' Conditional
'

#If Win64 Then
#Const USE_LONGLONG = True
#End If

'
' Declare
'

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As LongPtr)
#Else
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As Long)
#End If

'
' Types
'

Private Type IntegerT
    Value As Integer
End Type

Private Type LongT
    Value As Long
End Type

#If USE_LONGLONG Then
Private Type LongLongT
    Value As LongLong
End Type
#End If

Private Type SingleT
    Value As Single
End Type

Private Type DoubleT
    Value As Double
End Type

Private Type CurrencyT
    Value As Currency
End Type

Private Type DateT
    Value As Date
End Type

Private Type Bytes2T
    Bytes(0 To 1) As Byte
End Type

Private Type Bytes4T
    Bytes(0 To 3) As Byte
End Type

Private Type Bytes8T
    Bytes(0 To 7) As Byte
End Type

''
'' BitConverter - Test - Counter
''

Private m_Count As Long
Private m_Success As Long
Private m_Fail As Long

''
'' BitConverter - GetBytesFromXXX
''

'
' UInt16 - a 16-bit unsigned integer
'

Public Function GetBytesFromUInt16( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    
    Dim Bytes4() As Byte
    Bytes4 = GetBytesFromLong(Value, BigEndian)
    
    Dim Bytes(0 To 1) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes4, 2, 2
    Else
        CopyBytes Bytes, 0, Bytes4, 0, 2
    End If
    
    GetBytesFromUInt16 = Bytes
End Function

'
' UInt32 - a 32-bit unsigned integer
'

#If USE_LONGLONG Then

Public Function GetBytesFromUInt32( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    
    Dim Bytes8() As Byte
    Bytes8 = GetBytesFromLongLong(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes8, 4, 4
    Else
        CopyBytes Bytes, 0, Bytes8, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#Else

Public Function GetBytesFromUInt32( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("4294967295")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 10, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#End If

'
' UInt64 - a 64-bit unsigned integer
'

Public Function GetBytesFromUInt64( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("18446744073709551615")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 7) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 6, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 8
    End If
    
    GetBytesFromUInt64 = Bytes
End Function

'
' Int8 - a 8-bit signed integer
'

Public Function GetBytesFromInt8( _
    ByVal Value As Integer) As Byte()
    
    Debug.Assert ((Value >= -128) And (Value <= &H7F))
    
    Dim Bytes(0) As Byte
    Bytes(0) = Value And &HFF
    
    GetBytesFromInt8 = Bytes
End Function

'
' Int16 - a 16-bit signed integer
'

Public Function GetBytesFromInt16( _
    ByVal Value As Integer, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromInt16 = GetBytesFromInteger(Value, BigEndian)
End Function

'
' Int32 - a 32-bit signed integer
'

Public Function GetBytesFromInt32( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromInt32 = GetBytesFromLong(Value, BigEndian)
End Function

'
' Int64 - a 64-bit signed integer
'

#If USE_LONGLONG Then

Public Function GetBytesFromInt64( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromInt64 = GetBytesFromLongLong(Value, BigEndian)
End Function

#Else

Public Function GetBytesFromInt64( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert _
        ((Value >= CDec("-9223372036854775808")) And _
        (Value <= CDec("9223372036854775807")))
    
    Dim Bytes14() As Byte
    Dim Bytes(0 To 7) As Byte
    Dim Offset As Long
    
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        If Value < 0 Then
            Bytes14 = GetBytesFromDecimal(Value + 1, BigEndian)
            For Offset = 0 To 7
                Bytes(0 + Offset) = Not Bytes14(6 + Offset)
            Next
        Else
            Bytes14 = GetBytesFromDecimal(Value, BigEndian)
            CopyBytes Bytes, 0, Bytes14, 6, 8
        End If
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        If Value < 0 Then
            Bytes14 = GetBytesFromDecimal(Value + 1, BigEndian)
            For Offset = 0 To 7
                Bytes(0 + Offset) = Not Bytes14(0 + Offset)
            Next
        Else
            Bytes14 = GetBytesFromDecimal(Value, BigEndian)
            CopyBytes Bytes, 0, Bytes14, 0, 8
        End If
    End If
    
    GetBytesFromInt64 = Bytes
End Function

#End If

'
' Float32 - an IEEE 754 single precision floating point number
'

Public Function GetBytesFromFloat32( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat32 = GetBytesFromSingle(Value, BigEndian)
End Function

'
' Float64 - an IEEE 754 double precision floating point number
'

Public Function GetBytesFromFloat64( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat64 = GetBytesFromDouble(Value, BigEndian)
End Function

'
' Integer - a 16-bit signed integer
'

Public Function GetBytesFromInteger( _
    ByVal Value As Integer, Optional BigEndian As Boolean) As Byte()
    
    Dim I As IntegerT
    I.Value = Value
    
    Dim B2 As Bytes2T
    LSet B2 = I
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    GetBytesFromInteger = B2.Bytes
End Function

'
' Long - a 32-bit signed integer
'

Public Function GetBytesFromLong( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Dim L As LongT
    L.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = L
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromLong = B4.Bytes
End Function

'
' LongLong - a 64-bit signed integer
'

#If USE_LONGLONG Then
Public Function GetBytesFromLongLong( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Dim LL As LongLongT
    LL.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = LL
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromLongLong = B8.Bytes
End Function
#End If

'
' Single - an IEEE 754 single precision floating point number
'

Public Function GetBytesFromSingle( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    Dim S As SingleT
    S.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = S
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromSingle = B4.Bytes
End Function

'
' Double - an IEEE 754 double precision floating point number
'

Public Function GetBytesFromDouble( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    Dim D As DoubleT
    D.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = D
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromDouble = B8.Bytes
End Function

'
' Currency - a 64-bit number
'in an integer format, scaled by 10,000 to give a fixed-point number
'with 15 digits to the left of the decimal point and 4 digits to the right.
'

Public Function GetBytesFromCurrency( _
    ByVal Value As Currency, Optional BigEndian As Boolean) As Byte()
    
    Dim C As CurrencyT
    C.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = C
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromCurrency = B8.Bytes
End Function

'
' Date - an IEEE 754 double precision floating point number
'that represent dates ranging from 1 January 100, to 31 December 9999,
'and times from 0:00:00 to 23:59:59.
'
'When other numeric types are converted to Date,
'values to the left of the decimal represent date information,
'while values to the right of the decimal represent time.
'Midnight is 0 and midday is 0.5.
'

Public Function GetBytesFromDate( _
    ByVal Value As Date, Optional BigEndian As Boolean) As Byte()
    
    Dim D As DateT
    D.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = D
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromDate = B8.Bytes
End Function

'
' Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Public Function GetBytesFromDecimal( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    CopyMemory ByVal VarPtr(BytesRaw(0)), ByVal VarPtr(Value), 16
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    CopyBytes BytesX, 0, BytesRaw, 2, 14
    
    ' Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim Bytes(0 To 13) As Byte
    
    ' data low bytes
    CopyBytes Bytes, 0, BytesX, 6, 8
    
    ' data high bytes
    CopyBytes Bytes, 8, BytesX, 2, 4
    
    ' scale
    Bytes(12) = BytesX(0)
    
    ' sign
    Bytes(13) = BytesX(1)
    
    If BigEndian Then
        ReverseBytes Bytes
    End If
    
    GetBytesFromDecimal = Bytes
End Function

'
' String
'

Public Function GetBytesFromString( _
    ByVal Value As String, Optional Charset As String = "utf-8") As Byte()
    
    Debug.Assert (Value <> "")
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        .WriteText Value
        
        .Position = 0
        .Type = 1 'ADODB.adTypeBinary
        If Charset = "utf-8" Then
            .Position = 3 ' avoid BOM
        End If
        GetBytesFromString = .Read
        
        .Close
    End With
End Function

'
' Hex String
'

Public Function GetBytesFromHexString(ByVal Value As String) As Byte()
    Dim Value_ As String
    Dim Index As Long
    For Index = 1 To Len(Value)
        Select Case Mid(Value, Index, 1)
        Case "0" To "9", "A" To "F", "a" To "f"
            Value_ = Value_ & Mid(Value, Index, 1)
        End Select
    Next
    
    Dim Length As Long
    Length = Len(Value_) \ 2
    
    Dim Bytes() As Byte
    
    If Length = 0 Then
        GetBytesFromHexString = Bytes
        Exit Function
    End If
    
    ReDim Bytes(0 To Length - 1)
    
    'Dim Index As Long
    For Index = 0 To Length - 1
        Bytes(Index) = CByte("&H" & Mid(Value_, 1 + Index * 2, 2))
    Next
    
    GetBytesFromHexString = Bytes
End Function

''
'' BitConverter - GetXXXFromBytes
''

'
' UInt16 - a 16-bit unsigned integer
'

Public Function GetUInt16FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim Bytes4(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes4, 2, Bytes, Index, 2
    Else
        CopyBytes Bytes4, 0, Bytes, Index, 2
    End If
    
    GetUInt16FromBytes = GetLongFromBytes(Bytes4, 0, BigEndian)
End Function

'
' UInt32 - a 32-bit unsigned integer
'

#If USE_LONGLONG Then

Public Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim Bytes8(0 To 7) As Byte
    If BigEndian Then
        CopyBytes Bytes8, 4, Bytes, Index, 4
    Else
        CopyBytes Bytes8, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetLongLongFromBytes(Bytes8, 0, BigEndian)
End Function

#Else

Public Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 10, Bytes, Index, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

#End If

'
' UInt64 - a 64-bit unsigned integer
'

Public Function GetUInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 6, Bytes, Index, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 8
    End If
    
    GetUInt64FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

'
' Int8 - a 8-bit signed integer
'

Public Function GetInt8FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    Dim Bytes2LE(0 To 1) As Byte
    Bytes2LE(0) = Bytes(Index)
    If (Bytes(Index) And &H80) = &H80 Then
        Bytes2LE(1) = &HFF
    Else
        Bytes2LE(1) = 0
    End If
    
    GetInt8FromBytes = GetIntegerFromBytes(Bytes2LE)
End Function

'
' Int16 - a 16-bit signed integer
'

Public Function GetInt16FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    GetInt16FromBytes = GetIntegerFromBytes(Bytes, Index, BigEndian)
End Function

'
' Int32 - a 32-bit signed integer
'

Public Function GetInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    GetInt32FromBytes = GetLongFromBytes(Bytes, Index, BigEndian)
End Function

'
' Int64 - a 64-bit signed integer
'

#If USE_LONGLONG Then

Public Function GetInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    GetInt64FromBytes = GetLongLongFromBytes(Bytes, Index, BigEndian)
End Function

#Else

Public Function GetInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Offset As Long
    Dim Value As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        If (Bytes(Index) And &H80) = &H80 Then
            Bytes14(0) = &H80
            For Offset = 0 To 7
                Bytes14(6 + Offset) = Not Bytes(Index + Offset)
            Next
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian) - 1
        Else
            CopyBytes Bytes14, 6, Bytes, Index, 8
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian)
        End If
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        If (Bytes(Index + 7) And &H80) = &H80 Then
            For Offset = 0 To 7
                Bytes14(0 + Offset) = Not Bytes(Index + Offset)
            Next
            Bytes14(13) = &H80
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian) - 1
        Else
            CopyBytes Bytes14, 0, Bytes, Index, 8
            Value = GetDecimalFromBytes(Bytes14, 0, BigEndian)
        End If
    End If
    
    GetInt64FromBytes = Value
End Function

#End If

'
' Float32 - an IEEE 754 single precision floating point number
'

Public Function GetFloat32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    GetFloat32FromBytes = GetSingleFromBytes(Bytes, Index, BigEndian)
End Function

'
' Float64 - an IEEE 754 double precision floating point number
'

Public Function GetFloat64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    GetFloat64FromBytes = GetDoubleFromBytes(Bytes, Index, BigEndian)
End Function

'
' Integer - a 16-bit signed integer
'

Public Function GetIntegerFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    Dim B2 As Bytes2T
    CopyBytes B2.Bytes, 0, Bytes, Index, 2
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    Dim I As IntegerT
    LSet I = B2
    
    GetIntegerFromBytes = I.Value
End Function

'
' Long - a 32-bit signed integer
'

Public Function GetLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim L As LongT
    LSet L = B4
    
    GetLongFromBytes = L.Value
End Function

'
' LongLong - a 64-bit signed integer
'

#If USE_LONGLONG Then
Public Function GetLongLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim LL As LongLongT
    LSet LL = B8
    
    GetLongLongFromBytes = LL.Value
End Function
#End If

'
' Single - an IEEE 754 single precision floating point number
'

Public Function GetSingleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim S As SingleT
    LSet S = B4
    
    GetSingleFromBytes = S.Value
End Function

'
' Double - an IEEE 754 double precision floating point number
'

Public Function GetDoubleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim D As DoubleT
    LSet D = B8
    
    GetDoubleFromBytes = D.Value
End Function

'
' Currency - a 64-bit number
'in an integer format, scaled by 10,000 to give a fixed-point number
'with 15 digits to the left of the decimal point and 4 digits to the right.
'

Public Function GetCurrencyFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Currency
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim C As CurrencyT
    LSet C = B8
    
    GetCurrencyFromBytes = C.Value
End Function

'
' Date - an IEEE 754 double precision floating point number
'that represent dates ranging from 1 January 100, to 31 December 9999,
'and times from 0:00:00 to 23:59:59.
'
'When other numeric types are converted to Date,
'values to the left of the decimal represent date information,
'while values to the right of the decimal represent time.
'Midnight is 0 and midday is 0.5.
'

Public Function GetDateFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Date
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim D As DateT
    LSet D = B8
    
    GetDateFromBytes = D.Value
End Function

'
' Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Public Function GetDecimalFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    ' BytesXX = Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim BytesXX(0 To 13) As Byte
    CopyBytes BytesXX, 0, Bytes, Index, 14
    
    If BigEndian Then
        ReverseBytes BytesXX
    End If
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    
    ' scale
    BytesX(0) = BytesXX(12)
    
    ' sign
    BytesX(1) = BytesXX(13)
    
    ' data high bytes
    CopyBytes BytesX, 2, BytesXX, 8, 4
    
    ' data low bytes
    CopyBytes BytesX, 6, BytesXX, 0, 8
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    BytesRaw(0) = 14
    BytesRaw(1) = 0
    CopyBytes BytesRaw, 2, BytesX, Index, 14
    
    Dim Value As Variant
    CopyMemory ByVal VarPtr(Value), ByVal VarPtr(BytesRaw(0)), 16
    
    GetDecimalFromBytes = Value
End Function

'
' String
'

Public Function GetStringFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional ByVal Length As Long, _
    Optional Charset As String = "utf-8") As String
    
    If Length = 0 Then
        Length = UBound(Bytes) - Index + 1
    End If
    
    Dim Bytes_() As Byte
    ReDim Bytes_(0 To Length - 1)
    CopyBytes Bytes_, 0, Bytes, Index, Length
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 1 'ADODB.adTypeBinary
        .Write Bytes_
        
        .Position = 0
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        GetStringFromBytes = .ReadText
        
        .Close
    End With
End Function

'
' Hex String
'

'Public Function GetHexStringFromBytes(Bytes() As Byte,
Public Function GetHexStringFromBytes(Bytes, _
    Optional Index As Long, Optional Length As Long, _
    Optional Separator As String) As String
    
    If Length = 0 Then
        On Error Resume Next
        Length = UBound(Bytes) - Index + 1
        On Error GoTo 0
    End If
    If Length = 0 Then
        GetHexStringFromBytes = ""
        Exit Function
    End If
    
    Dim HexString As String
    HexString = Right("0" & Hex(Bytes(Index)), 2)
    
    Dim Offset As Long
    For Offset = 1 To Length - 1
        HexString = _
            HexString & Separator & Right("0" & Hex(Bytes(Index + Offset)), 2)
    Next
    
    GetHexStringFromBytes = HexString
End Function

''
'' BitConverter - Byte Array Helper
''

Public Sub CopyBytes( _
    DstBytes() As Byte, DstIndex As Long, _
    SrcBytes, SrcIndex As Long, ByVal Length As Long)
    'SrcBytes() As Byte, SrcIndex As Long, ByVal Length As Long)
    
    Dim Offset As Long
    For Offset = 0 To Length - 1
        DstBytes(DstIndex + Offset) = SrcBytes(SrcIndex + Offset)
    Next
End Sub

Private Sub ReverseBytes( _
    ByRef Bytes() As Byte, _
    Optional Index As Long, _
    Optional ByVal Length As Long)
    
    Dim UB As Long
    
    If Length = 0 Then
        UB = UBound(Bytes)
        Length = UB - Index + 1
    Else
        UB = Index + Length - 1
    End If
    
    Dim Offset As Long
    For Offset = 0 To (Length \ 2) - 1
        Dim Temp As Byte
        Temp = Bytes(Index + Offset)
        Bytes(Index + Offset) = Bytes(UB - Offset)
        Bytes(UB - Offset) = Temp
    Next
End Sub

''
'' BitConverter - Test - Counter
''

Private Property Get Test_Count() As Long
    Test_Count = m_Count
End Property

Public Sub Test_Initialize()
    m_Count = 0
    m_Success = 0
    m_Fail = 0
End Sub

Private Sub Test_Countup(bSuccess As Boolean)
    m_Count = m_Count + 1
    If bSuccess Then
        m_Success = m_Success + 1
    Else
        m_Fail = m_Fail + 1
    End If
End Sub

Public Sub Test_Terminate()
    Debug.Print _
        "Count: " & CStr(m_Count) & ", " & _
        "Success: " & CStr(m_Success) & ", " & _
        "Fail: " & CStr(m_Fail)
End Sub

''
'' BitConverter - Test - Debug.Print
''

Public Sub DebugPrint_GetBytes( _
    Source, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    Dim bSuccess As Boolean
    bSuccess = CompareBytes(OutputBytes, ExpectedBytes)
    
    Test_Countup bSuccess
    
    Dim OutputBytesStr As String
    OutputBytesStr = _
        BitConverter.GetHexStringFromBytes(OutputBytes, , , " ")
    
    Dim ExpectedBytesStr As String
    ExpectedBytesStr = _
        BitConverter.GetHexStringFromBytes(ExpectedBytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & Source & _
        " Output: " & OutputBytesStr & _
        " Expect: " & ExpectedBytesStr
End Sub

Public Sub DebugPrint_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue, Output, Expect)
    
    Dim bSuccess As Boolean
    bSuccess = (OutputValue = ExpectedValue)
    
    Test_Countup bSuccess
    
    Dim HexString As String
    HexString = BitConverter.GetHexStringFromBytes(Bytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & HexString & _
        " Output: " & Output & _
        " Expect: " & Expect
End Sub

''
'' BitConverter - Test - Byte Array Helper
''

Private Function CompareBytes(Bytes1() As Byte, Bytes2() As Byte) As Boolean
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(Bytes1)
    UB1 = UBound(Bytes1)
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(Bytes2)
    UB2 = UBound(Bytes2)
    
    If (UB1 - LB1 + 1) <> (UB2 - LB2 + 1) Then Exit Function
    
    Dim Index As Long
    For Index = 0 To UB1 - LB1
        If Bytes1(LB1 + Index) <> Bytes2(LB2 + Index) Then Exit Function
    Next
    
    CompareBytes = True
End Function

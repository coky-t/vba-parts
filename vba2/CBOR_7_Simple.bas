Attribute VB_Name = "CBOR_7_Simple"
Option Explicit

'
' Copyright (c) 2022 Koki Takeyama
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
'' CBOR for VBA
''

'
' Conditional
'

' Integer
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

#If Win64 And USE_LONGLONG Then
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
'' CBOR for VBA - Encoding
''

Public Function GetCborBytes(Value) As Byte()
    Select Case VarType(Value)
    
    ' 0
    Case vbEmpty
        GetCborBytes = GetCborBytesFromEmpty
        
    ' 1
    Case vbNull
        GetCborBytes = GetCborBytesFromNull
        
    ' 11
    Case vbBoolean
        GetCborBytes = GetCborBytesFromBoolean(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 0. Empty
'

Private Function GetCborBytesFromEmpty() As Byte()
    GetCborBytesFromEmpty = GetCborBytesFromUndefined
End Function

'
' 1. Null
'

'Private Function GetCborBytesFromNull() As Byte()
'    GetCborBytesFromNull = GetCborBytesFromNull
'End Function

'
' 11. Boolean
'

Private Function GetCborBytesFromBoolean(Value) As Byte()
    If Value Then
        GetCborBytesFromBoolean = GetCborBytesFromTrue
    Else
        GetCborBytesFromBoolean = GetCborBytesFromFalse
    End If
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 7: simple/float
'

' 0xf4 | false
Private Function GetCborBytesFromFalse() As Byte()
    GetCborBytesFromFalse = GetCborBytes0(&HF4)
End Function

' 0xf5 | true
Private Function GetCborBytesFromTrue() As Byte()
    GetCborBytesFromTrue = GetCborBytes0(&HF5)
End Function

' 0xf6 | null
Private Function GetCborBytesFromNull() As Byte()
    GetCborBytesFromNull = GetCborBytes0(&HF6)
End Function

' 0xf7 | undefined
Private Function GetCborBytesFromUndefined() As Byte()
    GetCborBytesFromUndefined = GetCborBytes0(&HF7)
End Function

''
'' CBOR for VBA - Encoding - Formatter
''

Private Function GetCborBytes0(HeaderValue As Byte) As Byte()
    Dim CborBytes(0) As Byte
    CborBytes(0) = HeaderValue
    GetCborBytes0 = CborBytes
End Function

Private Function GetCborBytes1A( _
    HeaderValue As Byte, SrcByte As Byte) As Byte()
    Dim CborBytes(0 To 1) As Byte
    CborBytes(0) = HeaderValue
    CborBytes(1) = SrcByte
    GetCborBytes1A = CborBytes
End Function

Private Function GetCborBytes1B( _
    HeaderValue As Byte, SrcBytes() As Byte) As Byte()
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    
    Dim SrcLen As Long
    SrcLen = SrcUB - SrcLB + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen)
    CborBytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes, SrcLB, SrcLen
    
    GetCborBytes1B = CborBytes
End Function

Private Function GetCborBytes2A( _
    HeaderValue As Byte, SrcByte1 As Byte, SrcBytes2) As Byte()
    'HeaderValue As Byte, SrcByte1 As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To 1 + SrcLen2)
    CborBytes(0) = HeaderValue
    CborBytes(1) = SrcByte1
    
    CopyBytes CborBytes, 2, SrcBytes2, SrcLB2, SrcLen2
    
    GetCborBytes2A = CborBytes
End Function

Private Function GetCborBytes2B( _
    HeaderValue As Byte, SrcBytes1, SrcBytes2) As Byte()
    'HeaderValue As Byte, SrcBytes1() As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen1 + SrcLen2)
    CborBytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    CopyBytes CborBytes, 1 + SrcLen1, SrcBytes2, SrcLB2, SrcLen2
    
    GetCborBytes2B = CborBytes
End Function

Private Function GetCborBytes3A(HeaderValue As Byte, _
    SrcByte1 As Byte, SrcByte2 As Byte, SrcBytes3) As Byte()
    
    Dim SrcLB3 As Long
    Dim SrcUB3 As Long
    SrcLB3 = LBound(SrcBytes3)
    SrcUB3 = UBound(SrcBytes3)
    
    Dim SrcLen3 As Long
    SrcLen3 = SrcUB3 - SrcLB3 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To 1 + 1 + SrcLen3)
    CborBytes(0) = HeaderValue
    CborBytes(1) = SrcByte1
    CborBytes(2) = SrcByte2
    
    CopyBytes CborBytes, 3, SrcBytes3, SrcLB3, SrcLen3
    
    GetCborBytes3A = CborBytes
End Function

Private Function GetCborBytes3B(HeaderValue As Byte, _
    SrcBytes1, SrcByte2 As Byte, SrcBytes3) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB3 As Long
    Dim SrcUB3 As Long
    SrcLB3 = LBound(SrcBytes3)
    SrcUB3 = UBound(SrcBytes3)
    
    Dim SrcLen3 As Long
    SrcLen3 = SrcUB3 - SrcLB3 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen1 + 1 + SrcLen3)
    Bytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    
    CborBytes(SrcLen1 + 1) = SrcByte2
    
    CopyBytes CborBytes, 1 + SrcLen1 + 1, SrcBytes3, SrcLB3, SrcLen3
    
    GetCborBytes3B = CborBytes
End Function

''
'' CBOR for VBA - Encoding - Bytes Operator
''

Private Sub AddBytes(DstBytes() As Byte, SrcBytes() As Byte)
    Dim DstLB As Long
    Dim DstUB As Long
    DstLB = LBound(DstBytes)
    DstUB = UBound(DstBytes)
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    Dim SrcLen As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    SrcLen = SrcUB - SrcLB + 1
    
    ReDim Preserve DstBytes(DstLB To DstUB + SrcLen)
    CopyBytes DstBytes, DstUB + 1, SrcBytes, SrcLB, SrcLen
End Sub

Private Sub CopyBytes( _
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
'' CBOR for VBA - Decoding
''

Private Function GetCborLength( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    Dim ItemCount As Long
    Dim ItemLength As Long
    
    Select Case CborBytes(Index)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    Case &HF4
        GetCborLength = 1
        
    ' 0xf5 | true
    Case &HF5
        GetCborLength = 1
        
    ' 0xf6 | null
    Case &HF6
        GetCborLength = 1
        
    ' 0xf7 | undefined
    Case &HF7
        GetCborLength = 1
        
    ' 0xff | "break" stop code
    'Case &HFF
    '    GetCborLength = 1
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function IsCborObject( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    Select Case CborBytes(Index)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    ' 0xf5 | true
    ' 0xf6 | null
    ' 0xf7 | undefined
    ' 0xff | "break" stop code
    Case &HF4 To &HF7 ', &HFF
        IsCborObject = False
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetValue(CborBytes() As Byte, Optional Index As Long) As Variant
    Select Case CborBytes(Index)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    Case &HF4
        GetValue = GetFalseFromCborBytes(CborBytes, Index)
        
    ' 0xf5 | true
    Case &HF5
        GetValue = GetTrueFromCborBytes(CborBytes, Index)
        
    ' 0xf6 | null
    Case &HF6
        GetValue = GetNullFromCborBytes(CborBytes, Index)
        
    ' 0xf7 | undefined
    Case &HF7
        GetValue = GetUndefinedFromCborBytes(CborBytes, Index)
        
    ' 0xff | "break" stop code
    'Case &HFF
    '    ' to do
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' major type 7: simple/float
'

' 0xf4 | false
Private Function GetFalseFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    GetFalseFromCborBytes = False
End Function

' 0xf5 | true
Private Function GetTrueFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    GetTrueFromCborBytes = True
End Function

' 0xf6 | null
Private Function GetNullFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetNullFromCborBytes = Null
End Function

' 0xf7 | undefined
Private Function GetUndefinedFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetUndefinedFromCborBytes = Empty
End Function

' 0xff | "break" stop code
' to do

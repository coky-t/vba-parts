Attribute VB_Name = "CBOR_7_Float"
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

' Single
#Const SINGLE_TO_FLOAT32 = True

' Double
#Const DOUBLE_TO_FLOAT64 = True

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
    
    ' 4
    Case vbSingle
        GetCborBytes = GetCborBytesFromSingle(Value)
        
    ' 5
    Case vbDouble
        GetCborBytes = GetCborBytesFromDouble(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 4. Single
'
Private Function GetCborBytesFromSingle(Value) As Byte()
#If SINGLE_TO_FLOAT32 Then
    GetCborBytesFromSingle = GetCborBytesFromFloat32((Value))
#Else
    
    ' --- single precision floating point number ---
    
    ' Float32: sign=1bit, exponent=8bits (bias=127), fraction=23bits
    
    ' SngBytesBE
    ' +--------+--------+--------+--------+
    ' |sxxxxxxx xfffffff ffffffff ffffffff|
    ' +--------+--------+--------+--------+
    
    ' SngSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' SngExpBits
    ' +--------+
    ' |xxxxxxxx|
    ' +--------+
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-fffffff ffffffff ffffffff|
    ' +--------+--------+--------+
    
    Dim SngBytesBE() As Byte
    SngBytesBE = GetBytesFromSingle(Value, True)
    
    Dim SngSgnBit As Byte ' 1 bit
    Dim SngExpBits As Byte ' 8 bits
    Dim SngFrcBits(0 To 2) As Byte ' 23 bits
    
    SngSgnBit = SngBytesBE(0) \ &H80
    SngExpBits = (SngBytesBE(0) And &H7F) * 2 + (SngBytesBE(1) \ &H80)
    SngFrcBits(0) = SngBytesBE(1) And &H7F
    SngFrcBits(1) = SngBytesBE(2)
    SngFrcBits(2) = SngBytesBE(3)
    
    Dim SngExp As Integer
    SngExp = CInt(SngExpBits) - 127
    
    ' --- check ---
    
    Select Case SngExpBits
    Case 0
        If (SngFrcBits(0) = 0) And _
            (SngFrcBits(1) = 0) And _
            (SngFrcBits(2) = 0) Then
            ' Single:Zero to Half:Zero - nop - continue
        Else
            ' Single:SubNormal, Half:OutOfRange
            'GetCborBytesFromSingle = GetCborBytesFromFloat32((Value))
            GetCborBytesFromSingle = GetCborBytes1(&HFA, SngBytesBE)
            Exit Function
        End If
    Case &HFF
        ' Single:Infinity to Half:Infinity - nop - continue
        ' Single:NaN to Half:NaN - nop - continue
    Case Else
        Select Case SngExp
        Case -15
            ' Single:Normal to Half:SubNormal - nop - continue
        Case -14 To 15
            ' Single:Normal to Half:Normal - nop - continue
        Case Else
            ' Single:Normal, Half:OutOfRange
            'GetCborBytesFromSingle = GetCborBytesFromFloat32((Value))
            GetCborBytesFromSingle = GetCborBytes1(&HFA, SngBytesBE)
            Exit Function
        End Select
    End Select
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-ffggggg ggghhhhh hhhhhhhh|
    ' +--------+--------+--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff gggggggg|
    ' +--------+--------+
    
    If Not (SngFrcBits(2) = 0) Then
        'GetCborBytesFromSingle = GetCborBytesFromFloat32((Value))
        GetCborBytesFromSingle = GetCborBytes1(&HFA, SngBytesBE)
        Exit Function
    End If
    
    If Not ((SngFrcBits(1) And &H1F) = 0) Then
        'GetCborBytesFromSingle = GetCborBytesFromFloat32((Value))
        GetCborBytesFromSingle = GetCborBytes1(&HFA, SngBytesBE)
        Exit Function
    End If
    
    'GetCborBytesFromSingle = GetCborBytesFromFloat16((Value))
    'Exit Function
    
    ' --- half precision floating point number ---
    
    ' Float16: sign=1bit, exponent=5bits (bias=15), fraction=10bits
    
    Dim HlfSgnBit As Byte ' 1 bit
    Dim HlfExpBits As Byte ' 5 bits
    Dim HlfFrcBits(0 To 1) As Byte ' 10 bits
    
    HlfSgnBit = SngSgnBit
    
    Select Case SngExpBits
    Case &H0
        HlfExpBits = &H0
    Case &HFF
        HlfExpBits = &H1F
    Case Else
        HlfExpBits = CByte(SngExp + 15)
    End Select
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-ffggggg ggghhhhh hhhhhhhh|
    ' +--------+--------+--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff gggggggg|
    ' +--------+--------+
    
    HlfFrcBits(0) = SngFrcBits(0) \ &H20
    HlfFrcBits(1) = (SngFrcBits(0) And &H1F) * &H8 + (SngFrcBits(1) \ &H20)
    
    ' HlfSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' HlfExpBits
    ' +--------+
    ' |---xxxxx|
    ' +--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff ffffffff|
    ' +--------+--------+
    
    ' HlfBytesBE
    ' +--------+--------+
    ' |sxxxxxff ffffffff|
    ' +--------+--------+
    
    Dim HlfBytesBE(0 To 1) As Byte
    HlfBytesBE(0) = HlfSgnBit * &H80 + (HlfExpBits * &H4) + HlfFrcBits(0)
    HlfBytesBE(1) = HlfFrcBits(1)
    
    GetCborBytesFromSingle = GetCborBytes1(&HF9, HlfBytesBE)
    
#End If
End Function

'
' 5. Double
'
Private Function GetCborBytesFromDouble(Value) As Byte()
#If DOUBLE_TO_FLOAT64 Then
    GetCborBytesFromDouble = GetCborBytesFromFloat64((Value))
#Else
    
    ' --- double precision floating point number ---
    
    ' Float64: sign=1bit, exponent=11bits (bias=1023), fraction=52bits
    
    ' DblBytesBE
    ' +--------+--------+--------+--------+--------+--------+--------+--------+
    ' |sxxxxxxx xxxxffff ffffffff ffffffff ffffffff ffffffff ffffffff ffffffff|
    ' +--------+--------+--------+--------+--------+--------+--------+--------+
    
    ' DblSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' DblExpBits
    ' +--------+--------+
    ' |-----xxx xxxxxxxx|
    ' +--------+--------+
    
    ' DblFrcBits
    ' +--------+--------+--------+--------+--------+--------+--------+
    ' |----ffff ffffffff ffffffff ffffffff ffffffff ffffffff ffffffff|
    ' +--------+--------+--------+--------+--------+--------+--------+
    
    Dim DblBytesBE() As Byte
    DblBytesBE = GetBytesFromDouble(Value, True)
    
    Dim DblSgnBit As Byte ' 1 bit
    Dim DblExpBits(0 To 1) As Byte ' 11 bits
    Dim DblFrcBits(0 To 6) As Byte ' 52 bits
    
    DblSgnBit = DblBytesBE(0) \ &H80
    DblExpBits(0) = (DblBytesBE(0) And &H7F) \ &H10
    DblExpBits(1) = (DblBytesBE(0) And &HF) * &H10 + (DblBytesBE(1) \ &H10)
    DblFrcBits(0) = DblBytesBE(1) And &HF
    DblFrcBits(1) = DblBytesBE(2)
    DblFrcBits(2) = DblBytesBE(3)
    DblFrcBits(3) = DblBytesBE(4)
    DblFrcBits(4) = DblBytesBE(5)
    DblFrcBits(5) = DblBytesBE(6)
    DblFrcBits(6) = DblBytesBE(7)
    
    Dim DblExp As Integer
    DblExp = GetIntegerFromBytes(DblExpBits, 0, True) - 1023
    
    ' --- check ---
    
    Select Case DblExpBits(0) * &H100 + DblExpBits(1)
    Case 0
        If (DblFrcBits(0) = 0) And _
            (DblFrcBits(1) = 0) And _
            (DblFrcBits(2) = 0) And _
            (DblFrcBits(3) = 0) And _
            (DblFrcBits(4) = 0) And _
            (DblFrcBits(5) = 0) And _
            (DblFrcBits(6) = 0) Then
            ' Double:Zero to Single:Zero - nop - continue
        Else
            ' Double:SubNormal, Single:OutOfRange
            'GetCborBytesFromDouble = GetCborBytesFromFloat64((Value))
            GetCborBytesFromDouble = GetCborBytes1(&HFB, DblBytesBE)
            Exit Function
        End If
    Case &H7FF
        ' Double:Infinity to Single:Infinity - nop - continue
        ' Double:NaN to Single:NaN - nop - continue
    Case Else
        Select Case DblExp
        Case -127
            ' Double:Normal to Single:SubNormal - nop - continue
        Case -126 To 127
            ' Double:Normal to Single:Normal - nop - continue
        Case Else
            ' Double:Normal, Single:OutOfRange
            'GetCborBytesFromDouble = GetCborBytesFromFloat64((Value))
            GetCborBytesFromDouble = GetCborBytes1(&HFB, DblBytesBE)
            Exit Function
        End Select
    End Select
    
    ' DblFrcBits
    ' +--------+--------+--------+--------+--------+--------+--------+
    ' |----ffff fffggggg ggghhhhh hhhiiiii iiiiiiii iiiiiiii iiiiiiii|
    ' +--------+--------+--------+--------+--------+--------+--------+
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-fffffff gggggggg hhhhhhhh|
    ' +--------+--------+--------+
    
    If Not (DblFrcBits(6) = 0) Or _
        Not (DblFrcBits(5) = 0) Or _
        Not (DblFrcBits(4) = 0) Or _
        Not ((DblFrcBits(3) And &H1F) = 0) Then
        'GetCborBytesFromDouble = GetCborBytesFromFloat64((Value))
        GetCborBytesFromDouble = GetCborBytes1(&HFB, DblBytesBE)
        Exit Function
    End If
    
    'GetCborBytesFromDouble = GetCborBytesFromFloat32(CSng(Value))
    'Exit Function
    
    ' --- single precision floating point number ---
    
    ' Float32: sign=1bit, exponent=8bits (bias=127), fraction=23bits
    
    Dim SngExp As Integer
    
    SngExp = DblExp
    
    Dim SngSgnBit As Byte ' 1 bit
    Dim SngExpBits As Byte ' 8 bits
    Dim SngFrcBits(0 To 2) As Byte ' 23 bits
    
    SngSgnBit = DblSgnBit
    
    Select Case DblExpBits(0) * &H100 + DblExpBits(1)
    Case &H0
        SngExpBits = &H0
    Case &H7FF
        SngExpBits = &HFF
    Case Else
        SngExpBits = CByte(SngExp + 127)
    End Select
    
    ' DblFrcBits
    ' +--------+--------+--------+--------+--------+--------+--------+
    ' |----ffff fffggggg ggghhhhh hhhiiiii iiiiiiii iiiiiiii iiiiiiii|
    ' +--------+--------+--------+--------+--------+--------+--------+
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-fffffff gggggggg hhhhhhhh|
    ' +--------+--------+--------+
    
    SngFrcBits(0) = (DblFrcBits(0) And &HF) * &H8 + (DblFrcBits(1) \ &H20)
    SngFrcBits(1) = (DblFrcBits(1) And &H1F) * &H8 + (DblFrcBits(2) \ &H20)
    SngFrcBits(2) = (DblFrcBits(2) And &H1F) * &H8 + (DblFrcBits(3) \ &H20)
    
    ' SngSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' SngExpBits
    ' +--------+
    ' |xxxxxxxx|
    ' +--------+
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-fffffff ffffffff ffffffff|
    ' +--------+--------+--------+
    
    ' SngBytesBE
    ' +--------+--------+--------+--------+
    ' |sxxxxxxx xfffffff ffffffff ffffffff|
    ' +--------+--------+--------+--------+
    
    Dim SngBytesBE(0 To 3) As Byte
    SngBytesBE(0) = SngSgnBit * &H80 + (SngExpBits \ &H2)
    SngBytesBE(1) = (SngExpBits And &H1) * &H80 + (SngFrcBits(0) And &H7F)
    SngBytesBE(2) = SngFrcBits(1)
    SngBytesBE(3) = SngFrcBits(2)
    
    ' --- check ---
    
    Select Case SngExpBits
    Case 0
        If (SngFrcBits(0) = 0) And _
            (SngFrcBits(1) = 0) And _
            (SngFrcBits(2) = 0) Then
            ' Single:Zero to Half:Zero - nop - continue
        Else
            ' Single:SubNormal, Half:OutOfRange
            'GetCborBytesFromDouble = GetCborBytesFromFloat32(CSng(Value))
            GetCborBytesFromDouble = GetCborBytes1(&HFA, SngBytesBE)
            Exit Function
        End If
    Case &HFF
        ' Single:Infinity to Half:Infinity - nop - continue
        ' Single:NaN to Half:NaN - nop - continue
    Case Else
        Select Case SngExp
        Case -15
            ' Single:Normal to Half:SubNormal - nop - continue
        Case -14 To 15
            ' Single:Normal to Half:Normal - nop - continue
        Case Else
            ' Single:Normal, Half:OutOfRange
            'GetCborBytesFromDouble = GetCborBytesFromFloat32(CSng(Value))
            GetCborBytesFromDouble = GetCborBytes1(&HFA, SngBytesBE)
            Exit Function
        End Select
    End Select
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-ffggggg ggghhhhh hhhhhhhh|
    ' +--------+--------+--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff gggggggg|
    ' +--------+--------+
    
    If Not (SngFrcBits(2) = 0) Then
        'GetCborBytesFromDouble = GetCborBytesFromFloat32(CSng(Value))
        GetCborBytesFromDouble = GetCborBytes1(&HFA, SngBytesBE)
        Exit Function
    End If
    
    If Not ((SngFrcBits(1) And &H1F) = 0) Then
        'GetCborBytesFromDouble = GetCborBytesFromFloat32(CSng(Value))
        GetCborBytesFromDouble = GetCborBytes1(&HFA, SngBytesBE)
        Exit Function
    End If
    
    'GetCborBytesFromDouble = GetCborBytesFromFloat16(CSng(Value))
    'Exit Function
    
    ' --- half precision floating point number ---
    
    ' Float16: sign=1bit, exponent=5bits (bias=15), fraction=10bits
    
    Dim HlfSgnBit As Byte ' 1 bit
    Dim HlfExpBits As Byte ' 5 bits
    Dim HlfFrcBits(0 To 1) As Byte ' 10 bits
    
    HlfSgnBit = SngSgnBit
    
    Select Case SngExpBits
    Case &H0
        HlfExpBits = &H0
    Case &HFF
        HlfExpBits = &H1F
    Case Else
        HlfExpBits = CByte(SngExp + 15)
    End Select
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-ffggggg ggghhhhh hhhhhhhh|
    ' +--------+--------+--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff gggggggg|
    ' +--------+--------+
    
    HlfFrcBits(0) = SngFrcBits(0) \ &H20
    HlfFrcBits(1) = (SngFrcBits(0) And &H1F) * &H8 + (SngFrcBits(1) \ &H20)
    
    ' HlfSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' HlfExpBits
    ' +--------+
    ' |---xxxxx|
    ' +--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff ffffffff|
    ' +--------+--------+
    
    ' HlfBytesBE
    ' +--------+--------+
    ' |sxxxxxff ffffffff|
    ' +--------+--------+
    
    Dim HlfBytesBE(0 To 1) As Byte
    HlfBytesBE(0) = HlfSgnBit * &H80 + (HlfExpBits * &H4) + HlfFrcBits(0)
    HlfBytesBE(1) = HlfFrcBits(1)
    
    GetCborBytesFromDouble = GetCborBytes1(&HF9, HlfBytesBE)
    
#End If
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 7: simple/float
'

' 0xf9 | half-precision float (two-byte IEEE 754)
Private Function GetCborBytesFromFloat16(ByVal Value As Single) As Byte()
    GetCborBytesFromFloat16 = _
        GetCborBytes1(&HF9, GetBytesFromFloat16(Value, True))
End Function

' 0xfa | single-precision float (four-byte IEEE 754)
Private Function GetCborBytesFromFloat32(ByVal Value As Single) As Byte()
    GetCborBytesFromFloat32 = _
        GetCborBytes1(&HFA, GetBytesFromFloat32(Value, True))
End Function

' 0xfb | double-precision float (eight-byte IEEE 754)
Private Function GetCborBytesFromFloat64(ByVal Value As Double) As Byte()
    GetCborBytesFromFloat64 = _
        GetCborBytes1(&HFB, GetBytesFromFloat64(Value, True))
End Function

''
'' CBOR for VBA - Encoding - Formatter
''

Private Function GetCborBytes0(HeaderValue As Byte) As Byte()
    Dim CborBytes(0) As Byte
    CborBytes(0) = HeaderValue
    GetCborBytes0 = CborBytes
End Function

Private Function GetCborBytes1( _
    HeaderValue As Byte, SrcBytes) As Byte()
    'HeaderValue As Byte, SrcBytes() As Byte) As Byte()
    
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
    
    GetCborBytes1 = CborBytes
End Function

Private Function GetCborBytes2( _
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
    
    GetCborBytes2 = CborBytes
End Function

Private Function GetCborBytes3(HeaderValue As Byte, _
    SrcBytes1, SrcBytes2, SrcBytes3) As Byte()
    
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
    
    Dim SrcLB3 As Long
    Dim SrcUB3 As Long
    SrcLB3 = LBound(SrcBytes3)
    SrcUB3 = UBound(SrcBytes3)
    
    Dim SrcLen3 As Long
    SrcLen3 = SrcUB3 - SrcLB3 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen1 + SrcLen2 + SrcLen3)
    Bytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    
    CopyBytes CborBytes, 1 + SrcLen1, SrcBytes2, SrcLB2, SrcLen2
    
    CopyBytes CborBytes, 1 + SrcLen1 + SrcLen2, SrcBytes3, SrcLB3, SrcLen3
    
    GetCborBytes3 = CborBytes
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
'' CBOR for VBA - Encoding - Converter
''

'
' 0xf9. Float16 - an IEEE 754 half precision floating point number
'

Private Function GetBytesFromFloat16( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    ' --- single precision floating point number ---
    
    ' Float32: sign=1bit, exponent=8bits (bias=127), fraction=23bits
    
    ' SngBytesBE
    ' +--------+--------+--------+--------+
    ' |sxxxxxxx xfffffff ffffffff ffffffff|
    ' +--------+--------+--------+--------+
    
    ' SngSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' SngExpBits
    ' +--------+
    ' |xxxxxxxx|
    ' +--------+
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-fffffff ffffffff ffffffff|
    ' +--------+--------+--------+
    
    Dim SngBytesBE() As Byte
    SngBytesBE = GetBytesFromSingle(Value, True)
    
    Dim SngSgnBit As Boolean ' 1 bit
    Dim SngExpBits As Byte ' 8 bits
    Dim SngFrcBits(0 To 2) As Byte ' 23 bits
    
    SngSgnBit = SngBytesBE(0) \ &H80
    SngExpBits = (SngBytesBE(0) And &H7F) * 2 + (SngBytesBE(1) \ &H80)
    SngFrcBits(0) = SngBytesBE(1) And &H7F
    SngFrcBits(1) = SngBytesBE(2)
    SngFrcBits(2) = SngBytesBE(3)
    
    Dim SngExp As Integer
    SngExp = CInt(SngExpBits) - 127
    
    ' --- check ---
    
    Select Case SngExpBits
    Case 0
        If (SngFrcBits(0) = 0) And _
            (SngFrcBits(1) = 0) And _
            (SngFrcBits(2) = 0) Then
            ' Single:Zero to Half:Zero - nop - continue
        Else
            ' Single:SubNormal, Half:OutOfRange
            Err.Raise 6 ' overflow
        End If
    Case &HFF
        ' Single:Infinity to Half:Infinity - nop - continue
        ' Single:NaN to Half:NaN - nop - continue
    Case Else
        Select Case SngExp
        Case -15
            ' Single:Normal to Half:SubNormal - nop - continue
        Case -14 To 15
            ' Single:Normal to Half:Normal - nop - continue
        Case Else
            ' Single:Normal, Half:OutOfRange
            Err.Raise 6 ' overflow
        End Select
    End Select
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-ffggggg ggghhhhh hhhhhhhh|
    ' +--------+--------+--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff gggggggg|
    ' +--------+--------+
    
    If Not (SngFrcBits(2) = 0) Then
        Err.Raise 6 ' overflow
    End If
    
    If Not ((SngFrcBits(1) And &H1F) = 0) Then
        Err.Raise 6 ' overflow
    End If
    
    ' --- half precision floating point number ---
    
    ' Float16: sign=1bit, exponent=5bits (bias=15), fraction=10bits
    
    Dim HlfSgnBit As Boolean ' 1 bit
    Dim HlfExpBits As Byte ' 5 bits
    Dim HlfFrcBits(0 To 1) As Byte ' 10 bits
    
    HlfSgnBit = SngSgnBit
    
    Select Case SngExpBits
    Case &H0
        HlfExpBits = &H0
    Case &HFF
        HlfExpBits = &H1F
    Case Else
        HlfExpBits = CByte(SngExp + 15)
    End Select
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-ffggggg ggghhhhh hhhhhhhh|
    ' +--------+--------+--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff gggggggg|
    ' +--------+--------+
    
    HlfFrcBits(0) = SngFrcBits(0) \ &H20
    HlfFrcBits(1) = (SngFrcBits(0) And &H1F) * &H8 + (SngFrcBits(1) \ &H20)
    
    ' HlfSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' HlfExpBits
    ' +--------+
    ' |---xxxxx|
    ' +--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff ffffffff|
    ' +--------+--------+
    
    ' HlfBytesBE
    ' +--------+--------+
    ' |sxxxxxff ffffffff|
    ' +--------+--------+
    
    Dim HlfBytesBE(0 To 1) As Byte
    HlfBytesBE(0) = HlfSgnBit * &H80 + (HlfExpBits * &H4) + HlfFrcBits(0)
    HlfBytesBE(1) = HlfFrcBits(1)
    
    GetBytesFromFloat16 = HlfBytesBE
End Function

'
' 0xfa. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromFloat32( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat32 = GetBytesFromSingle(Value, BigEndian)
End Function

'
' 0xfb. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromFloat64( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat64 = GetBytesFromDouble(Value, BigEndian)
End Function

'
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromSingle( _
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
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromDouble( _
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

''
'' CBOR for VBA - Decoding
''

Public Function GetCborLength( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    Dim ItemCount As Long
    Dim ItemLength As Long
    
    Select Case CborBytes(Index)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    Case &HF9
        GetCborLength = 1 + 2
        
    ' 0xfa | single-precision float (four-byte IEEE 754)
    Case &HFA
        GetCborLength = 1 + 4
        
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HFB
        GetCborLength = 1 + 8
        
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
    
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    ' 0xfa | single-precision float (four-byte IEEE 754)
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HF9 To &HFB
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
    
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    Case &HF9
        GetValue = GetFloat16FromCborBytes(CborBytes, Index)
        
    ' 0xfa | single-precision float (four-byte IEEE 754)
    Case &HFA
        GetValue = GetFloat32FromCborBytes(CborBytes, Index)
        
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HFB
        GetValue = GetFloat64FromCborBytes(CborBytes, Index)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' major type 7: simple/float
'

' 0xf9 | half-precision float (two-byte IEEE 754)
Private Function GetFloat16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Single

    GetFloat16FromCborBytes = GetFloat16FromBytes(CborBytes, Index + 1, True)
End Function

' 0xfa | single-precision float (four-byte IEEE 754)
Private Function GetFloat32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Single
    
    GetFloat32FromCborBytes = GetFloat32FromBytes(CborBytes, Index + 1, True)
End Function

' 0xfb | double-precision float (eight-byte IEEE 754)
Private Function GetFloat64FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Double
    
    GetFloat64FromCborBytes = GetFloat64FromBytes(CborBytes, Index + 1, True)
End Function

''
'' CBOR for VBA - Decoding - Converter
''

'
' 0xf9. Float16 - an IEEE 754 half precision floating point number
'

Private Function GetFloat16FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single

    Dim HlfBytesBE(0 To 1) As Byte
    If BigEndian Then
        HlfBytesBE(0) = Bytes(Index)
        HlfBytesBE(1) = Bytes(Index + 1)
    Else
        HlfBytesBE(0) = Bytes(Index + 1)
        HlfBytesBE(1) = Bytes(Index)
    End If
    
    ' --- half precision floating point number ---
    
    ' Float16: sign=1bit, exponent=5bits (bias=15), fraction=10bits
    
    ' HlfBytesBE
    ' +--------+--------+
    ' |sxxxxxff ffffffff|
    ' +--------+--------+
    
    ' HlfSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' HlfExpBits
    ' +--------+
    ' |---xxxxx|
    ' +--------+
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff ffffffff|
    ' +--------+--------+
    
    Dim HlfSgnBit As Byte ' 1 bit
    Dim HlfExpBits As Byte ' 5 bits
    Dim HlfFrcBits(0 To 1) As Byte ' 10 bits
    
    HlfSgnBit = HlfBytesBE(0) \ &H80
    HlfExpBits = (HlfBytesBE(0) And &H7F) \ &H4
    HlfFrcBits(0) = HlfBytesBE(0) And &H3
    HlfFrcBits(1) = HlfBytesBE(1)
    
    Dim HlfExp As Integer
    HlfExp = CInt(HlfExpBits) - 15
    
    ' --- single precision floating point number ---
    
    ' Float32: sign=1bit, exponent=8bits (bias=127), fraction=23bits
    
    Dim SngSgnBit As Byte ' 1 bit
    Dim SngExpBits As Byte ' 8 bits
    Dim SngFrcBits(0 To 2) As Byte ' 23 bits
    
    SngSgnBit = HlfSgnBit
    
    Select Case HlfExpBits
    Case &H0
        If (HlfFrcBits(0) = 0) And (HlfFrcBits(1) = 0) Then
            ' Half:Zero to Single:Zero
            SngExpBits = 0
        Else
            ' Half:SubNormal to Single:Normal
            SngExpBits = CByte(-15 + 127)
        End If
    Case &H1F
        ' Half:Infinity to Single:Infinity
        ' Half:NaN to Single:NaN
        SngExpBits = &HFF
    Case Else
        ' Half:Normal to Single:Normal
        SngExpBits = CByte(HlfExp + 127)
    End Select
    
    ' HlfFrcBits
    ' +--------+--------+
    ' |------ff fffffggg|
    ' +--------+--------+
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-fffffff ggg00000 00000000|
    ' +--------+--------+--------+
    
    SngFrcBits(0) = (HlfFrcBits(0) And &H3) * &H20 + (HlfFrcBits(1) \ &H8)
    SngFrcBits(1) = (HlfFrcBits(1) And &H7) * &H20
    SngFrcBits(2) = 0
    
    ' SngSgnBit
    ' +--------+
    ' |-------s|
    ' +--------+
    
    ' SngExpBits
    ' +--------+
    ' |xxxxxxxx|
    ' +--------+
    
    ' SngFrcBits
    ' +--------+--------+--------+
    ' |-fffffff ffffffff ffffffff|
    ' +--------+--------+--------+
    
    ' SngBytesBE
    ' +--------+--------+--------+--------+
    ' |sxxxxxxx xfffffff ffffffff ffffffff|
    ' +--------+--------+--------+--------+
    
    Dim SngBytesBE(0 To 3) As Byte
    SngBytesBE(0) = SngSgnBit * &H80 + (SngExpBits \ &H2)
    SngBytesBE(1) = (SngExpBits And &H1) * &H80 + SngFrcBits(0)
    SngBytesBE(2) = SngFrcBits(1)
    SngBytesBE(3) = SngFrcBits(2)
    
    GetFloat16FromBytes = GetSingleFromBytes(SngBytesBE, 0, True)
End Function

'
' 0xfa. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetFloat32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    GetFloat32FromBytes = GetSingleFromBytes(Bytes, Index, BigEndian)
End Function

'
' 0xfb. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetFloat64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    GetFloat64FromBytes = GetDoubleFromBytes(Bytes, Index, BigEndian)
End Function

'
' 2. Integer - a 16-bit signed integer
'

Private Function GetIntegerFromBytes(Bytes() As Byte, _
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
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetSingleFromBytes(Bytes() As Byte, _
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
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetDoubleFromBytes(Bytes() As Byte, _
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

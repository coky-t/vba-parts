Attribute VB_Name = "ByteStringX"
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

Private Type IntegerType
    Value As Integer
End Type

Private Type LongType
    Value As Long
End Type

#If Win64 Then
Private Type LongLongType
    Value As LongLong
End Type
#End If

Private Type SingleType
    Value As Single
End Type

Private Type DoubleType
    Value As Double
End Type

Private Type ByteArray2
    Items(1) As Byte
End Type

Private Type ByteArray4
    Items(3) As Byte
End Type

Private Type ByteArray8
    Items(7) As Byte
End Type

Public Function GetStringBFromByte(Value As Byte) As String
    GetStringBFromByte = ChrB(Value)
End Function

Public Function GetStringB_LEFromInteger(Value As Integer) As String
    Dim I As IntegerType
    I.Value = Value
    
    Dim B2 As ByteArray2
    LSet B2 = I
    
    GetStringB_LEFromInteger = _
        ChrB(B2.Items(0)) & _
        ChrB(B2.Items(1))
End Function

Public Function GetStringB_BEFromInteger(Value As Integer) As String
    Dim I As IntegerType
    I.Value = Value
    
    Dim B2 As ByteArray2
    LSet B2 = I
    
    GetStringB_BEFromInteger = _
        ChrB(B2.Items(1)) & _
        ChrB(B2.Items(0))
End Function

Public Function GetStringB_LEFromLong(Value As Long) As String
    Dim L As LongType
    L.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = L
    
    GetStringB_LEFromLong = _
        ChrB(B4.Items(0)) & _
        ChrB(B4.Items(1)) & _
        ChrB(B4.Items(2)) & _
        ChrB(B4.Items(3))
End Function

Public Function GetStringB_BEFromLong(Value As Long) As String
    Dim L As LongType
    L.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = L
    
    GetStringB_BEFromLong = _
        ChrB(B4.Items(3)) & _
        ChrB(B4.Items(2)) & _
        ChrB(B4.Items(1)) & _
        ChrB(B4.Items(0))
End Function

#If Win64 Then
Public Function GetStringB_LEFromLongLong(Value As LongLong) As String
    Dim LL As LongLongType
    LL.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = LL
    
    GetStringB_LEFromLongLong = _
        ChrB(B8.Items(0)) & _
        ChrB(B8.Items(1)) & _
        ChrB(B8.Items(2)) & _
        ChrB(B8.Items(3)) & _
        ChrB(B8.Items(4)) & _
        ChrB(B8.Items(5)) & _
        ChrB(B8.Items(6)) & _
        ChrB(B8.Items(7))
End Function

Public Function GetStringB_BEFromLongLong(Value As LongLong) As String
    Dim LL As LongLongType
    LL.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = LL
    
    GetStringB_BEFromLongLong = _
        ChrB(B8.Items(7)) & _
        ChrB(B8.Items(6)) & _
        ChrB(B8.Items(5)) & _
        ChrB(B8.Items(4)) & _
        ChrB(B8.Items(3)) & _
        ChrB(B8.Items(2)) & _
        ChrB(B8.Items(1)) & _
        ChrB(B8.Items(0))
End Function
#End If

Public Function GetStringB_LEFromSingle(Value As Single) As String
    Dim S As SingleType
    S.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = S
    
    GetStringB_LEFromSingle = _
        ChrB(B4.Items(0)) & _
        ChrB(B4.Items(1)) & _
        ChrB(B4.Items(2)) & _
        ChrB(B4.Items(3))
End Function

Public Function GetStringB_BEFromSingle(Value As Single) As String
    Dim S As SingleType
    S.Value = Value
    
    Dim B4 As ByteArray4
    LSet B4 = S
    
    GetStringB_BEFromSingle = _
        ChrB(B4.Items(3)) & _
        ChrB(B4.Items(2)) & _
        ChrB(B4.Items(1)) & _
        ChrB(B4.Items(0))
End Function

Public Function GetStringB_LEFromDouble(Value As Double) As String
    Dim D As DoubleType
    D.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = D
    
    GetStringB_LEFromDouble = _
        ChrB(B8.Items(0)) & _
        ChrB(B8.Items(1)) & _
        ChrB(B8.Items(2)) & _
        ChrB(B8.Items(3)) & _
        ChrB(B8.Items(4)) & _
        ChrB(B8.Items(5)) & _
        ChrB(B8.Items(6)) & _
        ChrB(B8.Items(7))
End Function

Public Function GetStringB_BEFromDouble(Value As Double) As String
    Dim D As DoubleType
    D.Value = Value
    
    Dim B8 As ByteArray8
    LSet B8 = D
    
    GetStringB_BEFromDouble = _
        ChrB(B8.Items(7)) & _
        ChrB(B8.Items(6)) & _
        ChrB(B8.Items(5)) & _
        ChrB(B8.Items(4)) & _
        ChrB(B8.Items(3)) & _
        ChrB(B8.Items(2)) & _
        ChrB(B8.Items(1)) & _
        ChrB(B8.Items(0))
End Function

Public Function GetByteFromStringB( _
    StrB As String, Optional Pos As Long = 1) As Byte
    
    GetByteFromStringB = AscB(MidB(StrB, Pos, 1))
End Function

Public Function GetIntegerFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As Integer
    
    Dim B2 As ByteArray2
    B2.Items(0) = AscB(MidB(StrB_LE, Pos, 1))
    B2.Items(1) = AscB(MidB(StrB_LE, Pos + 1, 1))
    
    Dim I As IntegerType
    LSet I = B2
    
    GetIntegerFromStringB_LE = I.Value
End Function

Public Function GetIntegerFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As Integer
    
    Dim B2 As ByteArray2
    B2.Items(0) = AscB(MidB(StrB_BE, Pos + 1, 1))
    B2.Items(1) = AscB(MidB(StrB_BE, Pos, 1))
    
    Dim I As IntegerType
    LSet I = B2
    
    GetIntegerFromStringB_BE = I.Value
End Function

Public Function GetLongFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As Long
    
    Dim B4 As ByteArray4
    B4.Items(0) = AscB(MidB(StrB_LE, Pos, 1))
    B4.Items(1) = AscB(MidB(StrB_LE, Pos + 1, 1))
    B4.Items(2) = AscB(MidB(StrB_LE, Pos + 2, 1))
    B4.Items(3) = AscB(MidB(StrB_LE, Pos + 3, 1))
    
    Dim L As LongType
    LSet L = B4
    
    GetLongFromStringB_LE = L.Value
End Function

Public Function GetLongFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As Long
    
    Dim B4 As ByteArray4
    B4.Items(0) = AscB(MidB(StrB_BE, Pos + 3, 1))
    B4.Items(1) = AscB(MidB(StrB_BE, Pos + 2, 1))
    B4.Items(2) = AscB(MidB(StrB_BE, Pos + 1, 1))
    B4.Items(3) = AscB(MidB(StrB_BE, Pos, 1))
    
    Dim L As LongType
    LSet L = B4
    
    GetLongFromStringB_BE = L.Value
End Function

#If Win64 Then
Public Function GetLongLongFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As LongLong
    
    Dim B8 As ByteArray8
    B8.Items(0) = AscB(MidB(StrB_LE, Pos, 1))
    B8.Items(1) = AscB(MidB(StrB_LE, Pos + 1, 1))
    B8.Items(2) = AscB(MidB(StrB_LE, Pos + 2, 1))
    B8.Items(3) = AscB(MidB(StrB_LE, Pos + 3, 1))
    B8.Items(4) = AscB(MidB(StrB_LE, Pos + 4, 1))
    B8.Items(5) = AscB(MidB(StrB_LE, Pos + 5, 1))
    B8.Items(6) = AscB(MidB(StrB_LE, Pos + 6, 1))
    B8.Items(7) = AscB(MidB(StrB_LE, Pos + 7, 1))
    
    Dim LL As LongLongType
    LSet LL = B8
    
    GetLongLongFromStringB_LE = LL.Value
End Function

Public Function GetLongLongFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As LongLong
    
    Dim B8 As ByteArray8
    B8.Items(0) = AscB(MidB(StrB_BE, Pos + 7, 1))
    B8.Items(1) = AscB(MidB(StrB_BE, Pos + 6, 1))
    B8.Items(2) = AscB(MidB(StrB_BE, Pos + 5, 1))
    B8.Items(3) = AscB(MidB(StrB_BE, Pos + 4, 1))
    B8.Items(4) = AscB(MidB(StrB_BE, Pos + 3, 1))
    B8.Items(5) = AscB(MidB(StrB_BE, Pos + 2, 1))
    B8.Items(6) = AscB(MidB(StrB_BE, Pos + 1, 1))
    B8.Items(7) = AscB(MidB(StrB_BE, Pos, 1))
    
    Dim LL As LongLongType
    LSet LL = B8
    
    GetLongLongFromStringB_BE = LL.Value
End Function
#End If

Public Function GetSingleFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As Single
    
    Dim B4 As ByteArray4
    B4.Items(0) = AscB(MidB(StrB_LE, Pos, 1))
    B4.Items(1) = AscB(MidB(StrB_LE, Pos + 1, 1))
    B4.Items(2) = AscB(MidB(StrB_LE, Pos + 2, 1))
    B4.Items(3) = AscB(MidB(StrB_LE, Pos + 3, 1))
    
    Dim S As SingleType
    LSet S = B4
    
    GetSingleFromStringB_LE = S.Value
End Function

Public Function GetSingleFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As Single
    
    Dim B4 As ByteArray4
    B4.Items(0) = AscB(MidB(StrB_BE, Pos + 3, 1))
    B4.Items(1) = AscB(MidB(StrB_BE, Pos + 2, 1))
    B4.Items(2) = AscB(MidB(StrB_BE, Pos + 1, 1))
    B4.Items(3) = AscB(MidB(StrB_BE, Pos, 1))
    
    Dim S As SingleType
    LSet S = B4
    
    GetSingleFromStringB_BE = S.Value
End Function

Public Function GetDoubleFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As Double
    
    Dim B8 As ByteArray8
    B8.Items(0) = AscB(MidB(StrB_LE, Pos, 1))
    B8.Items(1) = AscB(MidB(StrB_LE, Pos + 1, 1))
    B8.Items(2) = AscB(MidB(StrB_LE, Pos + 2, 1))
    B8.Items(3) = AscB(MidB(StrB_LE, Pos + 3, 1))
    B8.Items(4) = AscB(MidB(StrB_LE, Pos + 4, 1))
    B8.Items(5) = AscB(MidB(StrB_LE, Pos + 5, 1))
    B8.Items(6) = AscB(MidB(StrB_LE, Pos + 6, 1))
    B8.Items(7) = AscB(MidB(StrB_LE, Pos + 7, 1))
    
    Dim D As DoubleType
    LSet D = B8
    
    GetDoubleFromStringB_LE = D.Value
End Function

Public Function GetDoubleFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As Double
    
    Dim B8 As ByteArray8
    B8.Items(0) = AscB(MidB(StrB_BE, Pos + 7, 1))
    B8.Items(1) = AscB(MidB(StrB_BE, Pos + 6, 1))
    B8.Items(2) = AscB(MidB(StrB_BE, Pos + 5, 1))
    B8.Items(3) = AscB(MidB(StrB_BE, Pos + 4, 1))
    B8.Items(4) = AscB(MidB(StrB_BE, Pos + 3, 1))
    B8.Items(5) = AscB(MidB(StrB_BE, Pos + 2, 1))
    B8.Items(6) = AscB(MidB(StrB_BE, Pos + 1, 1))
    B8.Items(7) = AscB(MidB(StrB_BE, Pos, 1))
    
    Dim D As DoubleType
    LSet D = B8
    
    GetDoubleFromStringB_BE = D.Value
End Function

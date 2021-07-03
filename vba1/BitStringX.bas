Attribute VB_Name = "BitStringX"
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

Private Type LongType
    Value As Long
End Type

Private Type LongLongType
#If Win64 Then
    Value As LongLong
#Else
    Values(1) As Long
#End If
End Type

Private Type SingleType
    Value As Single
End Type

Private Type DoubleType
    Value As Double
End Type

Public Function Bin(ByVal Value)
    If IsNull(Value) Then
        Bin = Null
        Exit Function
    End If
    
    If IsEmpty(Value) Then
        Bin = Empty
        Exit Function
    End If
    
    Select Case TypeName(Value)
    Case "Byte"
        Bin = GetBinStringFromByte(Value)
    Case "Integer"
        Bin = GetBinStringFromInteger(Value)
    Case "Long"
        Bin = GetBinStringFromLong(Value)
#If Win64 Then
    Case "LongLong"
        Bin = GetBinStringFromLongLong(Value)
#End If
    Case "Single"
        Bin = GetBinStringFromSingle(Value)
    Case "Double"
        Bin = GetBinStringFromDouble(Value)
    End Select
End Function

Private Function BinCore(ByVal Value) As String
    Dim BinStr As String
    Do
        BinStr = IIf((Value Mod 2) = 0, "0", "1") & BinStr
        Value = Value \ 2
    Loop Until Value = 0
    BinCore = BinStr
End Function

Public Function Zeros(ByVal Count As Long) As String
    Dim ZerosStr As String
    Dim Index As Long
    For Index = 1 To Count
        ZerosStr = ZerosStr & "0"
    Next
    Zeros = ZerosStr
End Function

Public Function GetBinStringFromByte( _
    ByVal Value As Byte, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    BinString = BinCore(Value)
    
    If ZeroPadding Then
        BinString = Right(Zeros(7) & BinString, 8)
    End If
    
    GetBinStringFromByte = BinString
End Function

Public Function GetBinStringFromByteArrayLE( _
    Values() As Byte, _
    Optional ZeroPadding As Boolean) As String
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(Values)
    UB = UBound(Values)
    
    Dim BinString As String
    BinString = GetBinStringFromByte(Values(UB), ZeroPadding)
    
    Dim Index As Long
    For Index = UB - 1 To LB Step -1
        If BinString = "0" Then
            BinString = GetBinStringFromByte(Values(Index), ZeroPadding)
        Else
            BinString = BinString & GetBinStringFromByte(Values(Index), True)
        End If
    Next
    
    GetBinStringFromByteArrayLE = BinString
End Function

Public Function GetBinStringFromInteger( _
    ByVal Value As Integer, _
    Optional ZeroPadding As Boolean) As String
    
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromInteger(Value)
    
    GetBinStringFromInteger = _
        GetBinStringFromByteArrayLE(ByteArray, ZeroPadding)
End Function

Public Function GetBinStringFromLong( _
    ByVal Value As Long, _
    Optional ZeroPadding As Boolean) As String
    
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromLong(Value)
    
    GetBinStringFromLong = _
        GetBinStringFromByteArrayLE(ByteArray, ZeroPadding)
End Function

#If Win64 Then
Public Function GetBinStringFromLongLong( _
    ByVal Value As LongLong, _
    Optional ZeroPadding As Boolean) As String
    
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromLongLong(Value)
    
    GetBinStringFromLongLong = _
        GetBinStringFromByteArrayLE(ByteArray, ZeroPadding)
End Function
#End If

Public Function GetBinStringFromSingle( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromSingle(Value)
    
    GetBinStringFromSingle = _
        GetBinStringFromByteArrayLE(ByteArray, ZeroPadding)
End Function

Public Function GetBinStringFromDouble( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromDouble(Value)
    
    GetBinStringFromDouble = _
        GetBinStringFromByteArrayLE(ByteArray, ZeroPadding)
End Function

Public Function GetOctStringFromByte( _
    ByVal Value As Byte, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetOctStringFromByte = Right(Zeros(2) & Oct(Value), 3)
    Else
        GetOctStringFromByte = Oct(Value)
    End If
End Function

Public Function GetOctStringFromInteger( _
    ByVal Value As Integer, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetOctStringFromInteger = Right(Zeros(5) & Oct(Value), 6)
    Else
        GetOctStringFromInteger = Oct(Value)
    End If
End Function

Public Function GetOctStringFromLong( _
    ByVal Value As Long, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetOctStringFromLong = Right(Zeros(10) & Oct(Value), 11)
    Else
        GetOctStringFromLong = Oct(Value)
    End If
End Function

#If Win64 Then
Public Function GetOctStringFromLongLong( _
    ByVal Value As LongLong, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetOctStringFromLongLong = Right(Zeros(21) & Oct(Value), 22)
    Else
        GetOctStringFromLongLong = Oct(Value)
    End If
End Function
#End If

Public Function GetOctStringFromSingle( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim S As SingleType
    S.Value = Value
    
    Dim L As LongType
    LSet L = S
    
    GetOctStringFromSingle = GetOctStringFromLong(L.Value, ZeroPadding)
End Function

Public Function GetOctStringFromDouble( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim D As DoubleType
    D.Value = Value
    
    Dim LL As LongLongType
    LSet LL = D
    
#If Win64 Then
    GetOctStringFromDouble = GetOctStringFromLongLong(LL.Value, ZeroPadding)
#Else
    Dim Temp As String
    Temp = GetOctStringFromLong(LL.Values(1), ZeroPadding)
    If Temp = "0" Then
        Temp = GetOctStringFromLong(LL.Values(0), ZeroPadding)
    Else
        Temp = Temp & GetOctStringFromLong(LL.Values(0), True)
    End If
    GetOctStringFromDouble = Temp
#End If
End Function

Public Function GetHexStringFromByte( _
    ByVal Value As Byte, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetHexStringFromByte = Right(Zeros(1) & Hex(Value), 2)
    Else
        GetHexStringFromByte = Hex(Value)
    End If
End Function

Public Function GetHexStringFromInteger( _
    ByVal Value As Integer, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetHexStringFromInteger = Right(Zeros(3) & Hex(Value), 4)
    Else
        GetHexStringFromInteger = Hex(Value)
    End If
End Function

Public Function GetHexStringFromLong( _
    ByVal Value As Long, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetHexStringFromLong = Right(Zeros(7) & Hex(Value), 8)
    Else
        GetHexStringFromLong = Hex(Value)
    End If
End Function

#If Win64 Then
Public Function GetHexStringFromLongLong( _
    ByVal Value As LongLong, _
    Optional ZeroPadding As Boolean) As String
    
    If ZeroPadding Then
        GetHexStringFromLongLong = Right(Zeros(15) & Hex(Value), 16)
    Else
        GetHexStringFromLongLong = Hex(Value)
    End If
End Function
#End If

Public Function GetHexStringFromSingle( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim S As SingleType
    S.Value = Value
    
    Dim L As LongType
    LSet L = S
    
    GetHexStringFromSingle = GetHexStringFromLong(L.Value, ZeroPadding)
End Function

Public Function GetHexStringFromDouble( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim D As DoubleType
    D.Value = Value
    
    Dim LL As LongLongType
    LSet LL = D
    
#If Win64 Then
    GetHexStringFromDouble = GetHexStringFromLongLong(LL.Value, ZeroPadding)
#Else
    Dim Temp As String
    Temp = GetHexStringFromLong(LL.Values(1), ZeroPadding)
    If Temp = "0" Then
        Temp = GetHexStringFromLong(LL.Values(0), ZeroPadding)
    Else
        Temp = Temp & GetHexStringFromLong(LL.Values(0), True)
    End If
    GetHexStringFromDouble = Temp
#End If
End Function

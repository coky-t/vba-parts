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

Public Function GetBinStringFromBinString(BinString As String) As String
    Dim Temp As String
    Dim Index As Long
    For Index = 1 To Len(BinString)
        Select Case Mid(BinString, Index, 1)
        Case "0"
            Temp = Temp & "0"
        Case "1"
            Temp = Temp & "1"
        Case Else
            ' nop
        End Select
    Next
    GetBinStringFromBinString = Temp
End Function

Public Function GetByteFromBinString(BinString As String) As Byte
    Dim Temp As String
    Temp = Right(Zeros(8) & GetBinStringFromBinString(BinString), 8)
    
    Dim Value As Byte
    Dim Index As Long
    For Index = 0 To 7
        If Mid(Temp, 8 - Index, 1) = "1" Then
            Value = Value + 2 ^ Index
        End If
    Next
    GetByteFromBinString = Value
End Function

Public Function GetByteArrayLEFromBinString( _
    BinString As String, _
    ByteCount As Long) As Variant
    
    Dim Temp As String
    Temp = GetBinStringFromBinString(BinString)
    Temp = Right(Zeros(8 * ByteCount) & Temp, 8 * ByteCount)
    
    Dim LE() As Byte
    ReDim LE(ByteCount - 1)
    
    Dim Index As Long
    For Index = 0 To ByteCount - 1
        LE(Index) = _
            GetByteFromBinString( _
                Mid(Temp, 1 + (ByteCount - 1 - Index) * 8, 8))
    Next
    
    GetByteArrayLEFromBinString = LE
End Function

Public Function GetIntegerFromBinString(BinString As String) As Integer
    Dim LE() As Byte
    LE = GetByteArrayLEFromBinString(BinString, 2)
    
    GetIntegerFromBinString = GetIntegerFromByteArrayLE(LE)
End Function

Public Function GetLongFromBinString(BinString As String) As Long
    Dim LE() As Byte
    LE = GetByteArrayLEFromBinString(BinString, 4)
    
    GetLongFromBinString = GetLongFromByteArrayLE(LE)
End Function

#If Win64 Then
Public Function GetLongLongFromBinString(BinString As String) As LongLong
    Dim LE() As Byte
    LE = GetByteArrayLEFromBinString(BinString, 8)
    
    GetLongLongFromBinString = GetLongLongFromByteArrayLE(LE)
End Function
#End If

Public Function GetSingleFromBinString(BinString As String) As Single
    Dim LE() As Byte
    LE = GetByteArrayLEFromBinString(BinString, 4)
    
    GetSingleFromBinString = GetSingleFromByteArrayLE(LE)
End Function

Public Function GetDoubleFromBinString(BinString As String) As Double
    Dim LE() As Byte
    LE = GetByteArrayLEFromBinString(BinString, 8)
    
    GetDoubleFromBinString = GetDoubleFromByteArrayLE(LE)
End Function

Public Function GetBinStringFromOctString(OctString As String) As String
    Dim BinString As String
    Dim Index As Long
    For Index = 1 To Len(OctString)
        Select Case Mid(OctString, Index, 1)
        Case "0"
            BinString = BinString & "000"
        Case "1"
            BinString = BinString & "001"
        Case "2"
            BinString = BinString & "010"
        Case "3"
            BinString = BinString & "011"
        Case "4"
            BinString = BinString & "100"
        Case "5"
            BinString = BinString & "101"
        Case "6"
            BinString = BinString & "110"
        Case "7"
            BinString = BinString & "111"
        Case Else
            ' nop
        End Select
    Next
    GetBinStringFromOctString = BinString
End Function

Public Function GetByteFromOctString(OctString As String) As Byte
    Dim BinString As String
    BinString = GetBinStringFromOctString(OctString)
    
    GetByteFromOctString = GetByteFromBinString(BinString)
End Function

Public Function GetIntegerFromOctString(OctString As String) As Integer
    Dim BinString As String
    BinString = GetBinStringFromOctString(OctString)
    
    GetIntegerFromOctString = GetIntegerFromBinString(BinString)
End Function

Public Function GetLongFromOctString(OctString As String) As Long
    Dim BinString As String
    BinString = GetBinStringFromOctString(OctString)
    
    GetLongFromOctString = GetLongFromBinString(BinString)
End Function

#If Win64 Then
Public Function GetLongLongFromOctString(OctString As String) As LongLong
    Dim BinString As String
    BinString = GetBinStringFromOctString(OctString)
    
    GetLongLongFromOctString = GetLongLongFromBinString(BinString)
End Function
#End If

Public Function GetSingleFromOctString(OctString As String) As Single
    Dim BinString As String
    BinString = GetBinStringFromOctString(OctString)
    
    GetSingleFromOctString = GetSingleFromBinString(BinString)
End Function

Public Function GetDoubleFromOctString(OctString As String) As Double
    Dim BinString As String
    BinString = GetBinStringFromOctString(OctString)
    
    GetDoubleFromOctString = GetDoubleFromBinString(BinString)
End Function

Public Function GetBinStringFromHexString(HexString As String) As String
    Dim BinString As String
    Dim Index As Long
    For Index = 1 To Len(HexString)
        Select Case UCase(Mid(HexString, Index, 1))
        Case "0"
            BinString = BinString & "0000"
        Case "1"
            BinString = BinString & "0001"
        Case "2"
            BinString = BinString & "0010"
        Case "3"
            BinString = BinString & "0011"
        Case "4"
            BinString = BinString & "0100"
        Case "5"
            BinString = BinString & "0101"
        Case "6"
            BinString = BinString & "0110"
        Case "7"
            BinString = BinString & "0111"
        Case "8"
            BinString = BinString & "1000"
        Case "9"
            BinString = BinString & "1001"
        Case "A"
            BinString = BinString & "1010"
        Case "B"
            BinString = BinString & "1011"
        Case "C"
            BinString = BinString & "1100"
        Case "D"
            BinString = BinString & "1101"
        Case "E"
            BinString = BinString & "1110"
        Case "F"
            BinString = BinString & "1111"
        Case Else
            ' nop
        End Select
    Next
    GetBinStringFromHexString = BinString
End Function

Public Function GetByteFromHexString(HexString As String) As Byte
    Dim BinString As String
    BinString = GetBinStringFromHexString(HexString)
    
    GetByteFromHexString = GetByteFromBinString(BinString)
End Function

Public Function GetIntegerFromHexString(HexString As String) As Integer
    Dim BinString As String
    BinString = GetBinStringFromHexString(HexString)
    
    GetIntegerFromHexString = GetIntegerFromBinString(BinString)
End Function

Public Function GetLongFromHexString(HexString As String) As Long
    Dim BinString As String
    BinString = GetBinStringFromHexString(HexString)
    
    GetLongFromHexString = GetLongFromBinString(BinString)
End Function

#If Win64 Then
Public Function GetLongLongFromHexString(HexString As String) As LongLong
    Dim BinString As String
    BinString = GetBinStringFromHexString(HexString)
    
    GetLongLongFromHexString = GetLongLongFromBinString(BinString)
End Function
#End If

Public Function GetSingleFromHexString(HexString As String) As Single
    Dim BinString As String
    BinString = GetBinStringFromHexString(HexString)
    
    GetSingleFromHexString = GetSingleFromBinString(BinString)
End Function

Public Function GetDoubleFromHexString(HexString As String) As Double
    Dim BinString As String
    BinString = GetBinStringFromHexString(HexString)
    
    GetDoubleFromHexString = GetDoubleFromBinString(BinString)
End Function

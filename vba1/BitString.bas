Attribute VB_Name = "BitString"
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

Public Function Ones(ByVal Count As Long) As String
    Dim OnesStr As String
    Dim Index As Long
    For Index = 1 To Count
        OnesStr = OnesStr & "1"
    Next
    Ones = OnesStr
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

Public Function GetBinStringFromInteger( _
    ByVal Value As Integer, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    If (Value And &H8000) = &H8000 Then
        BinString = "1" & Right(Zeros(14) & BinCore(Value And &H7FFF), 15)
    Else
        BinString = BinCore(Value)
        
        If ZeroPadding Then
            BinString = Right(Zeros(15) & BinString, 16)
        End If
    End If
    
    GetBinStringFromInteger = BinString
End Function

Public Function GetBinStringFromLong( _
    ByVal Value As Long, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    If (Value And &H80000000) = &H80000000 Then
        BinString = "1" & Right(Zeros(30) & BinCore(Value And &H7FFFFFFF), 31)
    Else
        BinString = BinCore(Value)
        
        If ZeroPadding Then
            BinString = Right(Zeros(31) & BinString, 32)
        End If
    End If
    
    GetBinStringFromLong = BinString
End Function

#If Win64 Then
Public Function GetBinStringFromLongLong( _
    ByVal Value As LongLong, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    'If (Value And &H8000000000000000) = &H8000000000000000 Then
    '    BinString = "1" & _
    '        Right(Zeros(62) & BinCore(Value And &H7FFFFFFFFFFFFFFF), 63)
    If Value < 0 Then
        Dim NotValue As LongLong
        NotValue = Not Value
        
        Do
            BinString = IIf((NotValue Mod 2) = 0, "1", "0") & BinString
            NotValue = NotValue \ 2
        Loop Until NotValue = 0
        
        BinString = Right(Ones(63) & BinString, 64)
    Else
        BinString = BinCore(Value)
        
        If ZeroPadding Then
            BinString = Right(Zeros(63) & BinString, 64)
        End If
    End If
    
    GetBinStringFromLongLong = BinString
End Function
#End If

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

Public Function GetStringB_LEFromBinString( _
    BinString As String, _
    ByteCount As Long) As String
    
    Dim Temp As String
    Temp = GetBinStringFromBinString(BinString)
    Temp = Right(Zeros(8 * ByteCount) & Temp, 8 * ByteCount)
    
    Dim StringB_LE As String
    
    Dim Index As Long
    For Index = 0 To ByteCount - 1
        StringB_LE = StringB_LE & _
            ChrB(GetByteFromBinString( _
                Mid(Temp, 1 + (ByteCount - 1 - Index) * 8, 8)))
    Next
    
    GetStringB_LEFromBinString = StringB_LE
End Function

Public Function GetIntegerFromBinString(BinString As String) As Integer
    Dim StringB_LE As String
    StringB_LE = GetStringB_LEFromBinString(BinString, 2)
    
    GetIntegerFromBinString = GetIntegerFromStringB_LE(StringB_LE)
End Function

Public Function GetLongFromBinString(BinString As String) As Long
    Dim StringB_LE As String
    StringB_LE = GetStringB_LEFromBinString(BinString, 4)
    
    GetLongFromBinString = GetLongFromStringB_LE(StringB_LE)
End Function

#If Win64 Then
Public Function GetLongLongFromBinString(BinString As String) As LongLong
    Dim StringB_LE As String
    StringB_LE = GetStringB_LEFromBinString(BinString, 8)
    
    GetLongLongFromBinString = GetLongLongFromStringB_LE(StringB_LE)
End Function
#End If

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

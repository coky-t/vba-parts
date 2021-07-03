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

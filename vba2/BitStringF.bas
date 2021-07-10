Attribute VB_Name = "BitStringF"
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

Private Function BinCore(ByVal Value) As String
    Dim BinStr As String
    Do
        BinStr = IIf((Value Mod 2) = 0, "0", "1") & BinStr
        Value = Value \ 2
    Loop Until Value = 0
    BinCore = BinStr
End Function

Private Function GetBinStringFromSingleLarge( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim AbsValue As Single
    AbsValue = Abs(Value)
    If AbsValue <= CSng(&H7FFFFFFF) Then Exit Function
    
    Dim SignValue As Boolean
    Dim ExpValue As Integer
    SignValue = (Value < 0!)
    ExpValue = 0
    
    Dim TempValue As Single
    TempValue = AbsValue
    Do
        TempValue = TempValue / 2
        ExpValue = ExpValue + 1
    Loop Until TempValue < CSng(&H7FFFFFFF)
    
    Dim FlacBinStringTemp As String
    FlacBinStringTemp = BinCore(CLng(TempValue))
    ExpValue = ExpValue + Len(FlacBinStringTemp) - 1
    If ExpValue > 127 Then
        If SignValue Then
            ' Negative Normal Maximum
            GetBinStringFromSingleLarge = "11111111011111111111111111111111"
        Else
            ' Positive Normal Maximum
            GetBinStringFromSingleLarge = "01111111011111111111111111111111"
        End If
        Exit Function
    End If
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = IIf(SignValue, "1", "0")
    ExpBinString = GetBinStringFromByte(CByte(ExpValue + 127), True)
    FlacBinString = Left(Mid(FlacBinStringTemp, 2) & Zeros(23), 23)
    
    GetBinStringFromSingleLarge = SignBinString & ExpBinString & FlacBinString
End Function

Private Function GetBinStringFromSingleSmall( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim AbsValue As Single
    AbsValue = Abs(Value)
    If AbsValue >= 1! Then Exit Function
    
    Dim SignValue As Boolean
    Dim ExpValue As Integer
    SignValue = (Value < 0!)
    ExpValue = 0
    
    Dim ExpNormal As Boolean
    Dim TempValue As Single
    TempValue = AbsValue
    Do
        TempValue = TempValue * 2
        ExpValue = ExpValue - 1
        If TempValue >= 1! Then
            TempValue = TempValue - 1!
            ExpNormal = True
            Exit Do
        End If
    Loop Until ExpValue = -126
    If Not ExpNormal Then ExpValue = -127
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = IIf(SignValue, "1", "0")
    ExpBinString = GetBinStringFromByte(CByte(ExpValue + 127), True)
    
    Dim Index As Integer
    For Index = 1 To 23
        TempValue = TempValue * 2
        If TempValue >= 1! Then
            TempValue = TempValue - 1!
            FlacBinString = FlacBinString & "1"
        Else
            FlacBinString = FlacBinString & "0"
        End If
    Next
    
    GetBinStringFromSingleSmall = SignBinString & ExpBinString & FlacBinString
End Function

Private Function GetBinStringFromSingleNormal( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim AbsValue As Single
    AbsValue = Abs(Value)
    If AbsValue < 1! Then Exit Function
    If AbsValue > CSng(&H7FFFFFFF) Then Exit Function
    
    Dim SignValue As Boolean
    Dim ExpValue As Integer
    SignValue = (Value < 0!)
    ExpValue = 0
    
    Dim TempValue As Single
    TempValue = AbsValue
    
    Dim FlacBinStringTemp As String
    FlacBinStringTemp = BinCore(CLng(TempValue))
    ExpValue = ExpValue + Len(FlacBinStringTemp) - 1
    
    TempValue = TempValue - CSng(CLng(TempValue))
    
    Dim Index As Integer
    For Index = Len(FlacBinStringTemp) + 1 To 23
        TempValue = TempValue * 2
        If TempValue >= 1! Then
            TempValue = TempValue - 1!
            FlacBinStringTemp = FlacBinStringTemp & "1"
        Else
            FlacBinStringTemp = FlacBinStringTemp & "0"
        End If
    Next
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = IIf(SignValue, "1", "0")
    ExpBinString = GetBinStringFromByte(CByte(ExpValue + 127), True)
    FlacBinString = Left(Mid(FlacBinStringTemp, 2) & Zeros(23), 23)
    
    GetBinStringFromSingleNormal = _
        SignBinString & ExpBinString & FlacBinString
End Function

Public Function GetBinStringFromSingle( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Select Case CStr(Value)
    Case "0"
        GetBinStringFromSingle = "00000000000000000000000000000000"
        Exit Function
    Case "inf"
        GetBinStringFromSingle = "01111111100000000000000000000000"
        Exit Function
    Case "nan"
        GetBinStringFromSingle = "01111111111111111111111111111111"
        Exit Function
    Case "-0"
        GetBinStringFromSingle = "10000000000000000000000000000000"
        Exit Function
    Case "-inf"
        GetBinStringFromSingle = "11111111100000000000000000000000"
        Exit Function
    Case "-nan", "-nan(ind)"
        GetBinStringFromSingle = "11111111111111111111111111111111"
        Exit Function
    End Select
    
    Dim AbsValue As Single
    AbsValue = Abs(Value)
    
    Dim SignValue As Boolean
    Dim ExpValue As Byte
    Dim FlacValue As Long
    SignValue = (Value < 0!)
    
    If AbsValue > CSng(&H7FFFFFFF) Then
        GetBinStringFromSingle = _
            GetBinStringFromSingleLarge(Value, ZeroPadding)
        Exit Function
    ElseIf AbsValue < 1! Then
        GetBinStringFromSingle = _
            GetBinStringFromSingleSmall(Value, ZeroPadding)
        Exit Function
    End If
    
    GetBinStringFromSingle = GetBinStringFromSingleNormal(Value, ZeroPadding)
End Function

Private Function GetBinStringFromDoubleLarge( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim AbsValue As Double
    AbsValue = Abs(Value)
#If Win64 Then
    If AbsValue <= CDbl(2 ^ 63 - 1) Then Exit Function
#Else
    If (AbsValue / CDbl(2 ^ 30)) <= CDbl(2 ^ 30 - 1) Then Exit Function
#End If
    
    Dim SignValue As Boolean
    Dim ExpValue As Integer
    SignValue = (Value < 0#)
    ExpValue = 0
    
    Dim TempValue As Double
    TempValue = AbsValue
    Do
        TempValue = TempValue / 2
        ExpValue = ExpValue + 1
#If Win64 Then
    Loop Until TempValue < CDbl(2 ^ 63 - 1)
#Else
    Loop Until (TempValue / CDbl(2 ^ 30)) < CDbl(2 ^ 30 - 1)
#End If
    
#If Win64 Then
    Dim FlacBinStringTemp As String
    FlacBinStringTemp = BinCore(CLngLng(TempValue))
#Else
    Dim TempHighValue As Long
    Dim TempLowValue As Long
    TempHighValue = TempValue / CDbl(2 ^ 30)
    TempLowValue = TempValue - CLng(TempValue / CDbl(2 ^ 30)) * (2 ^ 30)
    
    Dim FlacHighBinStringTemp As String
    Dim FlacLowBinStringTemp As String
    FlacHighBinStringTemp = BinCore(CLng(TempHighValue))
    FlacLowBinStringTemp = Right(Zeros(30) & BinCore(CLng(TempLowValue)), 30)
    
    Dim FlacBinStringTemp As String
    FlacBinStringTemp = FlacHighBinStringTemp & FlacLowBinStringTemp
#End If
    ExpValue = ExpValue + Len(FlacBinStringTemp) - 1
    If ExpValue > 1023 Then
        ' Positive Normal Maximum
        If SignValue Then
            GetBinStringFromDoubleLarge = "111111111110" & _
                "1111111111111111111111111111111111111111111111111111"
        Else
            GetBinStringFromDoubleLarge = "011111111110" & _
                "1111111111111111111111111111111111111111111111111111"
        End If
        Exit Function
    End If
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = IIf(SignValue, "1", "0")
    ExpBinString = Right(Zeros(11) & BinCore(CInt(ExpValue + 1023)), 11)
    FlacBinString = Left(Mid(FlacBinStringTemp, 2) & Zeros(52), 52)
    
    GetBinStringFromDoubleLarge = SignBinString & ExpBinString & FlacBinString
End Function

Private Function GetBinStringFromDoubleSmall( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim AbsValue As Double
    AbsValue = Abs(Value)
    If AbsValue >= 1# Then Exit Function
    
    Dim SignValue As Boolean
    Dim ExpValue As Integer
    SignValue = (Value < 0#)
    ExpValue = 0
    
    Dim ExpNormal As Boolean
    Dim TempValue As Double
    TempValue = AbsValue
    Do
        TempValue = TempValue * 2
        ExpValue = ExpValue - 1
        If TempValue >= 1# Then
            TempValue = TempValue - 1#
            ExpNormal = True
            Exit Do
        End If
    Loop Until ExpValue = -1022
    If Not ExpNormal Then ExpValue = -1023
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = IIf(SignValue, "1", "0")
    ExpBinString = Right(Zeros(11) & BinCore(CInt(ExpValue + 1023)), 11)
    
    Dim Index As Integer
    For Index = 1 To 52
        TempValue = TempValue * 2
        If TempValue >= 1# Then
            TempValue = TempValue - 1#
            FlacBinString = FlacBinString & "1"
        Else
            FlacBinString = FlacBinString & "0"
        End If
    Next
    
    GetBinStringFromDoubleSmall = SignBinString & ExpBinString & FlacBinString
End Function

Private Function GetBinStringFromDoubleNormal( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim AbsValue As Double
    AbsValue = Abs(Value)
    If AbsValue < 1# Then Exit Function
#If Win64 Then
    If AbsValue > CDbl(2 ^ 63 - 1) Then Exit Function
#Else
    If (AbsValue / CDbl(2 ^ 30)) > CDbl(2 ^ 30 - 1) Then Exit Function
#End If
    
    Dim SignValue As Boolean
    Dim ExpValue As Integer
    SignValue = (Value < 0#)
    ExpValue = 0
    
    Dim TempValue As Double
    TempValue = AbsValue
    
#If Win64 Then
    Dim FlacBinStringTemp As String
    FlacBinStringTemp = BinCore(CLngLng(TempValue))
#Else
    Dim TempHighValue As Long
    Dim TempLowValue As Long
    TempHighValue = TempValue / CDbl(2 ^ 30)
    TempLowValue = TempValue - CLng(TempValue / CDbl(2 ^ 30)) * (2 ^ 30)
    
    Dim FlacHighBinStringTemp As String
    Dim FlacLowBinStringTemp As String
    If TempHighValue = 0 Then
        FlacHighBinStringTemp = ""
        FlacLowBinStringTemp = BinCore(CLng(TempLowValue))
    Else
        FlacHighBinStringTemp = BinCore(CLng(TempHighValue))
        FlacLowBinStringTemp = _
            Right(Zeros(30) & BinCore(CLng(TempLowValue)), 30)
    End If
    
    Dim FlacBinStringTemp As String
    FlacBinStringTemp = FlacHighBinStringTemp & FlacLowBinStringTemp
#End If
    ExpValue = ExpValue + Len(FlacBinStringTemp) - 1
    
#If Win64 Then
    TempValue = TempValue - CDbl(CLngLng(TempValue))
#Else
    TempValue = _
        TempValue - CDbl(TempHighValue) * CDbl(2 ^ 30) - CDbl(TempLowValue)
#End If
    
    Dim Index As Integer
    For Index = Len(FlacBinStringTemp) + 1 To 52
        TempValue = TempValue * 2
        If TempValue >= 1# Then
            TempValue = TempValue - 1#
            FlacBinStringTemp = FlacBinStringTemp & "1"
        Else
            FlacBinStringTemp = FlacBinStringTemp & "0"
        End If
    Next
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = IIf(SignValue, "1", "0")
    ExpBinString = Right(Zeros(11) & BinCore(CInt(ExpValue + 1023)), 11)
    FlacBinString = Left(Mid(FlacBinStringTemp, 2) & Zeros(52), 52)
    
    GetBinStringFromDoubleNormal = _
        SignBinString & ExpBinString & FlacBinString
End Function

Public Function GetBinStringFromDouble( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Select Case CStr(Value)
    Case "0"
        GetBinStringFromDouble = _
            "0000000000000000000000000000000000000000000000000000000000000000"
        Exit Function
    Case "inf"
        GetBinStringFromDouble = _
            "0111111111110000000000000000000000000000000000000000000000000000"
        Exit Function
    Case "nan"
        GetBinStringFromDouble = _
            "0111111111111111111111111111111111111111111111111111111111111111"
        Exit Function
    Case "-0"
        GetBinStringFromDouble = _
            "1000000000000000000000000000000000000000000000000000000000000000"
        Exit Function
    Case "-inf"
        GetBinStringFromDouble = _
            "1111111111110000000000000000000000000000000000000000000000000000"
        Exit Function
    Case "-nan", "-nan(ind)"
        GetBinStringFromDouble = _
            "1111111111111111111111111111111111111111111111111111111111111111"
        Exit Function
    End Select
    
    Dim AbsValue As Double
    AbsValue = Abs(Value)
    
    Dim SignValue As Boolean
    Dim ExpValue As Byte
    Dim FlacValue As Long
    SignValue = (Value < 0#)
    
#If Win64 Then
    If AbsValue > CDbl(2 ^ 63 - 1) Then
#Else
    If (AbsValue / CDbl(2 ^ 30)) > CDbl(2 ^ 30 - 1) Then
#End If
        GetBinStringFromDouble = _
            GetBinStringFromDoubleLarge(Value, ZeroPadding)
        Exit Function
    ElseIf AbsValue < 1# Then
        GetBinStringFromDouble = _
            GetBinStringFromDoubleSmall(Value, ZeroPadding)
        Exit Function
    End If
    
    GetBinStringFromDouble = GetBinStringFromDoubleNormal(Value, ZeroPadding)
End Function

Public Function GetOctStringFromBinString(BinString As String) As String
    Dim BinStringTemp As String
    BinStringTemp = GetBinStringFromBinString(BinString)
    BinStringTemp = _
        Right(Zeros(3 - 1) & BinStringTemp, _
            ((Len(BinStringTemp) + 3 - 1) \ 3) * 3)
    
    Dim OctString As String
    Dim Index As Long
    For Index = 1 To Len(BinStringTemp) Step 3
        Select Case Mid(BinStringTemp, Index, 3)
        Case "000"
            OctString = OctString & "0"
        Case "001"
            OctString = OctString & "1"
        Case "010"
            OctString = OctString & "2"
        Case "011"
            OctString = OctString & "3"
        Case "100"
            OctString = OctString & "4"
        Case "101"
            OctString = OctString & "5"
        Case "110"
            OctString = OctString & "6"
        Case "111"
            OctString = OctString & "7"
        Case Else
            ' nop
        End Select
    Next
    GetOctStringFromBinString = OctString
End Function

Public Function GetOctStringFromSingle( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    BinString = GetBinStringFromSingle(Value, ZeroPadding)
    
    GetOctStringFromSingle = GetOctStringFromBinString(BinString)
End Function

Public Function GetOctStringFromDouble( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    BinString = GetBinStringFromDouble(Value, ZeroPadding)
    
    GetOctStringFromDouble = GetOctStringFromBinString(BinString)
End Function

Public Function GetHexStringFromBinString(BinString As String) As String
    Dim BinStringTemp As String
    BinStringTemp = GetBinStringFromBinString(BinString)
    BinStringTemp = _
        Right(Zeros(4 - 1) & BinStringTemp, _
            ((Len(BinStringTemp) + 4 - 1) \ 4) * 4)
    
    Dim HexString As String
    Dim Index As Long
    For Index = 1 To Len(BinStringTemp) Step 4
        Select Case Mid(BinStringTemp, Index, 4)
        Case "0000"
            HexString = HexString & "0"
        Case "0001"
            HexString = HexString & "1"
        Case "0010"
            HexString = HexString & "2"
        Case "0011"
            HexString = HexString & "3"
        Case "0100"
            HexString = HexString & "4"
        Case "0101"
            HexString = HexString & "5"
        Case "0110"
            HexString = HexString & "6"
        Case "0111"
            HexString = HexString & "7"
        Case "1000"
            HexString = HexString & "8"
        Case "1001"
            HexString = HexString & "9"
        Case "1010"
            HexString = HexString & "A"
        Case "1011"
            HexString = HexString & "B"
        Case "1100"
            HexString = HexString & "C"
        Case "1101"
            HexString = HexString & "D"
        Case "1110"
            HexString = HexString & "E"
        Case "1111"
            HexString = HexString & "F"
        Case Else
            ' nop
        End Select
    Next
    GetHexStringFromBinString = HexString
End Function

Public Function GetHexStringFromSingle( _
    ByVal Value As Single, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    BinString = GetBinStringFromSingle(Value, ZeroPadding)
    
    GetHexStringFromSingle = GetHexStringFromBinString(BinString)
End Function

Public Function GetHexStringFromDouble( _
    ByVal Value As Double, _
    Optional ZeroPadding As Boolean) As String
    
    Dim BinString As String
    BinString = GetBinStringFromDouble(Value, ZeroPadding)
    
    GetHexStringFromDouble = GetHexStringFromBinString(BinString)
End Function

Public Function GetSingleFromBinString(BinString As String) As Single
    Dim TempBinString As String
    TempBinString = GetBinStringFromBinString(BinString)
    TempBinString = Right(Zeros(32) & TempBinString, 32)
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = Left(TempBinString, 1)
    ExpBinString = Mid(TempBinString, 2, 8)
    FlacBinString = Right(TempBinString, 23)
    
    Dim SignBitValue As Boolean
    Dim ExpBitsValue As Byte
    Dim FlacBitsValue As Long
    SignBitValue = (SignBinString = "1")
    ExpBitsValue = GetByteFromBinString(ExpBinString)
    FlacBitsValue = GetLongFromBinString(FlacBinString)
    
    ' Zero
    If (ExpBitsValue = 0) And (FlacBitsValue = 0) Then
        If SignBitValue Then
            GetSingleFromBinString = -0!
        Else
            GetSingleFromBinString = 0!
        End If
        Exit Function
    End If
    
    ' SubNormal
    If ExpBitsValue = 0 Then ' FlacBitsValue <> 0
        If SignBitValue Then
            GetSingleFromBinString = _
                -(CSng(FlacBitsValue) * (2 ^ (-23))) * (2 ^ (-126))
        Else
            GetSingleFromBinString = _
                (CSng(FlacBitsValue) * (2 ^ (-23))) * (2 ^ (-126))
        End If
        Exit Function
    End If
    
    ' Normal
    If ExpBitsValue < &HFF Then ' And (ExpBitsValue > 0)
        If SignBitValue Then
            GetSingleFromBinString = _
                -(1! + CSng(FlacBitsValue) * (2 ^ (-23))) * _
                (2 ^ (ExpBitsValue - 127))
        Else
            GetSingleFromBinString = _
                (1! + CSng(FlacBitsValue) * (2 ^ (-23))) * _
                (2 ^ (ExpBitsValue - 127))
        End If
        Exit Function
    End If
    
    ' Infinity
    If (ExpBitsValue = &HFF) And (FlacBitsValue = 0) Then
        If SignBitValue Then
            On Error Resume Next
            GetSingleFromBinString = -1! / 0!
            On Error GoTo 0
        Else
            On Error Resume Next
            GetSingleFromBinString = 1! / 0!
            On Error GoTo 0
        End If
        Exit Function
    End If
    
    ' NaN
    'If (ExpBitsValue = &HFF) And (FlacBitsValue <> 0) Then
        If SignBitValue Then
            On Error Resume Next
            GetSingleFromBinString = -0! / 0!
            On Error GoTo 0
        Else
            On Error Resume Next
            GetSingleFromBinString = 0! / 0!
            On Error GoTo 0
        End If
        Exit Function
    'End If
End Function

Public Function GetDoubleFromBinString(BinString As String) As Double
#If Win64 Then
    Dim TempBinString As String
    TempBinString = GetBinStringFromBinString(BinString)
    TempBinString = Right(Zeros(64) & TempBinString, 64)
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacBinString As String
    SignBinString = Left(TempBinString, 1)
    ExpBinString = Mid(TempBinString, 2, 11)
    FlacBinString = Right(TempBinString, 52)
    
    Dim SignBitValue As Boolean
    Dim ExpBitsValue As Integer
    Dim FlacBitsValue As LongLong
    SignBitValue = (SignBinString = "1")
    ExpBitsValue = GetIntegerFromBinString(ExpBinString)
    FlacBitsValue = GetLongLongFromBinString(FlacBinString)
    
    ' Zero
    If (ExpBitsValue = 0) And (FlacBitsValue = 0) Then
        If SignBitValue Then
            GetDoubleFromBinString = -0!
        Else
            GetDoubleFromBinString = 0!
        End If
        Exit Function
    End If
    
    ' SubNormal
    If ExpBitsValue = 0 Then ' FlacBitsValue <> 0
        If SignBitValue Then
            GetDoubleFromBinString = _
                -(CDbl(FlacBitsValue) * (2 ^ (-52))) * (2 ^ (-1022))
        Else
            GetDoubleFromBinString = _
                (CDbl(FlacBitsValue) * (2 ^ (-52))) * (2 ^ (-1022))
        End If
        Exit Function
    End If
    
    ' Normal
    If ExpBitsValue < &H7FF Then ' And (ExpBitsValue > 0)
        If SignBitValue Then
            GetDoubleFromBinString = _
                -(1# + CDbl(FlacBitsValue) * (2 ^ (-52))) * _
                (2 ^ (ExpBitsValue - 1023))
        Else
            GetDoubleFromBinString = _
                (1# + CDbl(FlacBitsValue) * (2 ^ (-52))) * _
                (2 ^ (ExpBitsValue - 1023))
        End If
        Exit Function
    End If
    
    ' Infinity
    If (ExpBitsValue = &H7FF) And (FlacBitsValue = 0) Then
        If SignBitValue Then
            On Error Resume Next
            GetDoubleFromBinString = -1# / 0#
            On Error GoTo 0
        Else
            On Error Resume Next
            GetDoubleFromBinString = 1# / 0#
            On Error GoTo 0
        End If
        Exit Function
    End If
    
    ' NaN
    'If (ExpBitsValue = &H7FF) And (FlacBitsValue <> 0) Then
        If SignBitValue Then
            On Error Resume Next
            GetDoubleFromBinString = -0# / 0#
            On Error GoTo 0
        Else
            On Error Resume Next
            GetDoubleFromBinString = 0# / 0#
            On Error GoTo 0
        End If
        Exit Function
    'End If
#Else
    Dim TempBinString As String
    TempBinString = GetBinStringFromBinString(BinString)
    TempBinString = Right(Zeros(64) & TempBinString, 64)
    
    Dim SignBinString As String
    Dim ExpBinString As String
    Dim FlacHighBinString As String
    Dim FlacLowBinString As String
    SignBinString = Left(TempBinString, 1)
    ExpBinString = Mid(TempBinString, 2, 11)
    FlacHighBinString = Mid(TempBinString, 13, 26)
    FlacLowBinString = Right(TempBinString, 26)
    
    Dim SignBitValue As Boolean
    Dim ExpBitsValue As Integer
    Dim FlacHighBitsValue As Long
    Dim FlacLowBitsValue As Long
    SignBitValue = (SignBinString = "1")
    ExpBitsValue = GetIntegerFromBinString(ExpBinString)
    FlacHighBitsValue = GetLongFromBinString(FlacHighBinString)
    FlacLowBitsValue = GetLongFromBinString(FlacLowBinString)
    
    ' Zero
    If (ExpBitsValue = 0) And _
        (FlacHighBitsValue = 0) And (FlacLowBitsValue = 0) Then
        
        If SignBitValue Then
            GetDoubleFromBinString = -0#
        Else
            GetDoubleFromBinString = 0#
        End If
        Exit Function
    End If
    
    ' SubNormal
    If ExpBitsValue = 0 Then ' FlacHighBitsValue <> 0 Or FlacLowBitsValue <> 0
        If SignBitValue Then
            GetDoubleFromBinString = _
                -(CDbl(FlacLowBitsValue) * (2 ^ (-52)) + _
                CDbl(FlacHighBitsValue) * (2 ^ (-26))) * (2 ^ (-1022))
        Else
            GetDoubleFromBinString = _
                (CDbl(FlacLowBitsValue) * (2 ^ (-52)) + _
                CDbl(FlacHighBitsValue) * (2 ^ (-26))) * (2 ^ (-1022))
        End If
        Exit Function
    End If
    
    ' Normal
    If ExpBitsValue < &H7FF Then ' And (ExpBitsValue > 0)
        If SignBitValue Then
            GetDoubleFromBinString = _
                -(1# + CDbl(FlacLowBitsValue) * (2 ^ (-52)) + _
                CDbl(FlacHighBitsValue) * (2 ^ (-26))) * _
                (2 ^ (ExpBitsValue - 1023))
        Else
            GetDoubleFromBinString = _
                (1# + CDbl(FlacLowBitsValue) * (2 ^ (-52)) + _
                CDbl(FlacHighBitsValue) * (2 ^ (-26))) * _
                (2 ^ (ExpBitsValue - 1023))
        End If
        Exit Function
    End If
    
    ' Infinity
    If (ExpBitsValue = &H7FF) And _
        (FlacHighBitsValue = 0) And (FlacLowBitsValue = 0) Then
        If SignBitValue Then
            On Error Resume Next
            GetDoubleFromBinString = -1# / 0#
            On Error GoTo 0
        Else
            On Error Resume Next
            GetDoubleFromBinString = 1# / 0#
            On Error GoTo 0
        End If
        Exit Function
    End If
    
    ' NaN
    'If (ExpBitsValue = &H7FF) And _
    '    ((FlacHighBitsValue <> 0) Or (FlacLowBitsValue <> 0)) Then
        If SignBitValue Then
            On Error Resume Next
            GetDoubleFromBinString = -0# / 0#
            On Error GoTo 0
        Else
            On Error Resume Next
            GetDoubleFromBinString = 0# / 0#
            On Error GoTo 0
        End If
        Exit Function
    'End If
#End If
End Function

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

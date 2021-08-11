Attribute VB_Name = "Test_MsgPack_Str"
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

'
' MessagePack for VBA - String - Test
'

Public Sub Test_MsgPack_Str()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Str_FixStr_TestCases
    Test_MsgPack_Str_Str8_TestCases
    Test_MsgPack_Str_Str16_TestCases
    Test_MsgPack_Str_Str32_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Str_FixStr()
    BitConverter.Test_Initialize
    Test_MsgPack_Str_FixStr_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Str_Str8()
    BitConverter.Test_Initialize
    Test_MsgPack_Str_Str8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Str_Str16()
    BitConverter.Test_Initialize
    Test_MsgPack_Str_Str16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Str_Str32()
    BitConverter.Test_Initialize
    Test_MsgPack_Str_Str32_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - String - Test Cases
'

Public Sub Test_MsgPack_Str_FixStr_TestCases()
    Debug.Print "Target: FixStr"
    
    Test_MsgPack_Str_Core "A0", ""
    Test_MsgPack_Str_Core "A1 61", "a"
    Test_MsgPack_Str_Core "A3 E3 81 82", ChrW(&H3042)
    Test_MsgPack_Str_Core _
        "BF 61 62 63 64 65 66 67 68 69 6A 6B 6C 6D 6E 6F" & _
        "70 71 72 73 74 75 76 77 78 79 7A 41 42 43 44 45", _
        "abcdefghijklmnopqrstuvwxyzABCDE"
End Sub

Public Sub Test_MsgPack_Str_Str8_TestCases()
    Debug.Print "Target: Str8"
    
    Test_MsgPack_Str_Core _
        "D9 20 61 62 63 64 65 66 67 68 69 6A 6B 6C 6D 6E 6F" & _
        "70 71 72 73 74 75 76 77 78 79 7A 41 42 43 44 45 46", _
        "abcdefghijklmnopqrstuvwxyzABCDEF"
    Test_MsgPack_Str_Core2 "D9 FF", &HFF
End Sub

Public Sub Test_MsgPack_Str_Str16_TestCases()
    Debug.Print "Target: Str16"
    
    Test_MsgPack_Str_Core2 "DA 01 00", &H100
    Test_MsgPack_Str_Core2 "DA FF FF", &HFFFF&
End Sub

Public Sub Test_MsgPack_Str_Str32_TestCases()
    Debug.Print "Target: Str32"
    
    Test_MsgPack_Str_Core2 "DB 00 01 00 00", &H10000
End Sub

'
' MessagePack for VBA - String - Test Core
'

Public Sub Test_MsgPack_Str_Core(HexStr As String, ExpectedValue As String)
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytesFromHexString(HexStr)
    
    Dim OutputValue As String
    OutputValue = MsgPack_Str.GetStrFromBytes(Bytes)
    
    DebugPrint_MsgPack_Str_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Str.GetBytesFromStr(OutputValue)
    
    DebugPrint_MsgPack_Str_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Public Sub Test_MsgPack_Str_Core2(HeadHex As String, DataLength As Long)
    Dim HeadBytes() As Byte
    HeadBytes = BitConverter.GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestBytes(HeadBytes, DataLength)
    
    Dim ExpectedValue As String
    ExpectedValue = GetTestStr(DataLength)
    
    Dim OutputValue As String
    OutputValue = MsgPack_Str.GetStrFromBytes(Bytes)
    
    DebugPrint_MsgPack_Str_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Str.GetBytesFromStr(OutputValue)
    
    DebugPrint_MsgPack_Str_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Private Function GetTestStr(Length As Long) As String
    Dim TestStr As String
    
    Dim Index As Long
    For Index = 1 To Length
        TestStr = TestStr & Hex(Index Mod 16)
    Next
    
    GetTestStr = TestStr
End Function

Private Function GetTestBytes( _
    HeadBytes() As Byte, BodyLength As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(HeadLength + BodyLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To BodyLength
        TestBytes(HeadLength + Index - 1) = Asc(Hex(Index Mod 16))
    Next
    
    GetTestBytes = TestBytes
End Function

'
' MessagePack for VBA - String - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Str_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        Value, OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Str_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        OutputValue, ExpectedValue
End Sub

Attribute VB_Name = "Test_MsgPack_Bin"
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
' MessagePack for VBA - Binary - Test
'

Public Sub Test_MsgPack_Bin()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Bin_Bin8_TestCases
    Test_MsgPack_Bin_Bin16_TestCases
    Test_MsgPack_Bin_Bin32_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Bin_Bin8()
    BitConverter.Test_Initialize
    Test_MsgPack_Bin_Bin8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Bin_Bin16()
    BitConverter.Test_Initialize
    Test_MsgPack_Bin_Bin16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Bin_Bin32()
    BitConverter.Test_Initialize
    Test_MsgPack_Bin_Bin32_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Binary - Test Cases
'

Public Sub Test_MsgPack_Bin_Bin8_TestCases()
    Debug.Print "Target: Bin8"
    
    Test_MsgPack_Bin_Core "C4 00", ""
    Test_MsgPack_Bin_Core2 "C4 01", &H1
    Test_MsgPack_Bin_Core2 "C4 FF", &HFF
End Sub

Public Sub Test_MsgPack_Bin_Bin16_TestCases()
    Debug.Print "Target: Bin16"
    
    Test_MsgPack_Bin_Core2 "C5 01 00", &H100
    Test_MsgPack_Bin_Core2 "C5 FF FF", &HFFFF&
End Sub

Public Sub Test_MsgPack_Bin_Bin32_TestCases()
    Debug.Print "Target: Bin32"
    
    Test_MsgPack_Bin_Core2 "C6 00 01 00 00", &H10000
End Sub

'
' MessagePack for VBA - Binary - Test Core
'

Public Sub Test_MsgPack_Bin_Core(HexStr As String, ExpectedHexStr As String)
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytesFromHexString(HexStr)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = BitConverter.GetBytesFromHexString(ExpectedHexStr)
    
    Dim OutputValue() As Byte
    OutputValue = MsgPack_Bin.GetBinFromBytes(Bytes)
    
    DebugPrint_MsgPack_Bin_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Bin.GetBytesFromBin(OutputValue)
    
    DebugPrint_MsgPack_Bin_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Public Sub Test_MsgPack_Bin_Core2(HeadHex As String, DataLength As Long)
    Dim HeadBytes() As Byte
    HeadBytes = BitConverter.GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestBytes(HeadBytes, DataLength)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = GetTestValue(DataLength)
    
    Dim OutputValue() As Byte
    OutputValue = MsgPack_Bin.GetBinFromBytes(Bytes)
    
    DebugPrint_MsgPack_Bin_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Bin.GetBytesFromBin(OutputValue)
    
    DebugPrint_MsgPack_Bin_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Private Function GetTestValue(Length As Long) As Byte()
    Dim TestValue() As Byte
    ReDim TestValue(0 To Length - 1)
    
    Dim Index As Long
    For Index = 1 To Length
        TestValue(Index - 1) = Index Mod 256
    Next
    
    GetTestValue = TestValue
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
        TestBytes(HeadLength + Index - 1) = Index Mod 256
    Next
    
    GetTestBytes = TestBytes
End Function

'
' MessagePack for VBA - Binary - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Bin_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    Dim HexString As String
    HexString = BitConverter.GetHexStringFromBytes(Value, , , " ")
    
    BitConverter.DebugPrint_GetBytes _
        HexString, OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Bin_GetValue( _
    Bytes() As Byte, OutputValue() As Byte, ExpectedValue() As Byte)
    
    Dim OutputHexString As String
    OutputHexString = BitConverter.GetHexStringFromBytes(OutputValue, , , " ")
    
    Dim ExpectedHexString As String
    ExpectedHexString = _
        BitConverter.GetHexStringFromBytes(ExpectedValue, , , " ")
    
    BitConverter.DebugPrint_GetValue Bytes, _
        OutputHexString, ExpectedHexString, _
        OutputHexString, ExpectedHexString
End Sub

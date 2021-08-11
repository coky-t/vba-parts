Attribute VB_Name = "Test_MsgPack_Ext"
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
' MessagePack for VBA - Extension - Test
'

Public Sub Test_MsgPack_Ext()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Ext_FixExt1_TestCases
    Test_MsgPack_Ext_FixExt2_TestCases
    Test_MsgPack_Ext_FixExt4_TestCases
    Test_MsgPack_Ext_FixExt8_TestCases
    Test_MsgPack_Ext_FixExt16_TestCases
    Test_MsgPack_Ext_Ext8_TestCases
    Test_MsgPack_Ext_Ext16_TestCases
    Test_MsgPack_Ext_Ext32_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_FixExt1()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_FixExt1_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_FixExt2()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_FixExt2_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_FixExt4()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_FixExt4_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_FixExt8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_FixExt8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_FixExt16()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_FixExt16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Ext8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Ext8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Ext16()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Ext16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Ext32()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Ext32_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Extension - Test Cases
'

Public Sub Test_MsgPack_Ext_FixExt1_TestCases()
    Debug.Print "Target: FixExt1"
    
    Test_MsgPack_Ext_Core "D4 01 00", &H1, "00"
    Test_MsgPack_Ext_Core "D4 01 01", &H1, "01"
    Test_MsgPack_Ext_Core "D4 01 FF", &H1, "FF"
End Sub

Public Sub Test_MsgPack_Ext_FixExt2_TestCases()
    Debug.Print "Target: FixExt2"
    
    Test_MsgPack_Ext_Core "D5 01 00 00", &H1, "00 00"
    Test_MsgPack_Ext_Core "D5 01 00 01", &H1, "00 01"
    Test_MsgPack_Ext_Core "D5 01 01 00", &H1, "01 00"
    Test_MsgPack_Ext_Core "D5 01 FF FF", &H1, "FF FF"
End Sub

Public Sub Test_MsgPack_Ext_FixExt4_TestCases()
    Debug.Print "Target: FixExt4"
    
    Test_MsgPack_Ext_Core "D6 01 00 00 00 00", &H1, "00 00 00 00"
    Test_MsgPack_Ext_Core "D6 01 00 00 00 01", &H1, "00 00 00 01"
    Test_MsgPack_Ext_Core "D6 01 00 00 01 00", &H1, "00 00 01 00"
    Test_MsgPack_Ext_Core "D6 01 00 01 00 00", &H1, "00 01 00 00"
    Test_MsgPack_Ext_Core "D6 01 01 00 00 00", &H1, "01 00 00 00"
    Test_MsgPack_Ext_Core "D6 01 FF FF FF FF", &H1, "FF FF FF FF"
End Sub

Public Sub Test_MsgPack_Ext_FixExt8_TestCases()
    Debug.Print "Target: FixExt8"
    
    Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00"
    Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 00 00 01", &H1, _
        "00 00 00 00 00 00 00 01"
    Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 00 01 00", &H1, _
        "00 00 00 00 00 00 01 00"
    Test_MsgPack_Ext_Core "D7 01 00 00 00 00 00 01 00 00", &H1, _
        "00 00 00 00 00 01 00 00"
    Test_MsgPack_Ext_Core "D7 01 00 00 00 00 01 00 00 00", &H1, _
        "00 00 00 00 01 00 00 00"
    Test_MsgPack_Ext_Core "D7 01 00 00 00 01 00 00 00 00", &H1, _
        "00 00 00 01 00 00 00 00"
    Test_MsgPack_Ext_Core "D7 01 00 00 01 00 00 00 00 00", &H1, _
        "00 00 01 00 00 00 00 00"
    Test_MsgPack_Ext_Core "D7 01 00 01 00 00 00 00 00 00", &H1, _
        "00 01 00 00 00 00 00 00"
    Test_MsgPack_Ext_Core "D7 01 01 00 00 00 00 00 00 00", &H1, _
        "01 00 00 00 00 00 00 00"
    Test_MsgPack_Ext_Core "D7 01 FF FF FF FF FF FF FF FF", &H1, _
        "FF FF FF FF FF FF FF FF"
End Sub

Public Sub Test_MsgPack_Ext_FixExt16_TestCases()
    Debug.Print "Target: FixExt16"
    
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00"
    Test_MsgPack_Ext_Core _
        "D8 01 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00 00", &H1, _
        "00 00 00 00 00 00 00 00 01 00 00 00 00 00 00 00"
    Test_MsgPack_Ext_Core _
        "D8 01 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF", &H1, _
        "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
End Sub

Public Sub Test_MsgPack_Ext_Ext8_TestCases()
    Debug.Print "Target: Ext8"
    
    Test_MsgPack_Ext_Core "C7 00 01", &H1, ""
    Test_MsgPack_Ext_Core2 "C7 03 01", &H3
    Test_MsgPack_Ext_Core2 "C7 05 01", &H5
    Test_MsgPack_Ext_Core2 "C7 07 01", &H7
    Test_MsgPack_Ext_Core2 "C7 09 01", &H9
    Test_MsgPack_Ext_Core2 "C7 0F 01", &HF
    Test_MsgPack_Ext_Core2 "C7 11 01", &H11
    Test_MsgPack_Ext_Core2 "C7 FF 01", &HFF
End Sub

Public Sub Test_MsgPack_Ext_Ext16_TestCases()
    Debug.Print "Target: Ext16"
    
    Test_MsgPack_Ext_Core2 "C8 01 00 01", &H100
    Test_MsgPack_Ext_Core2 "C8 FF FF 01", &HFFFF&
End Sub

Public Sub Test_MsgPack_Ext_Ext32_TestCases()
    Debug.Print "Target: Ext32"
    
    Test_MsgPack_Ext_Core2 "C9 00 01 00 00 01", &H10000
End Sub

'
' MessagePack for VBA - Extension - Test Core
'

Public Sub Test_MsgPack_Ext_Core( _
    HexStr As String, ExtType As Byte, ExpectedHexStr As String)
    
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytesFromHexString(HexStr)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = BitConverter.GetBytesFromHexString(ExpectedHexStr)
    
    Dim OutputValue() As Byte
    OutputValue = MsgPack_Ext.GetExtFromBytes(Bytes)
    
    DebugPrint_MsgPack_Ext_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Ext.GetBytesFromExt(ExtType, OutputValue)
    
    DebugPrint_MsgPack_Ext_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Public Sub Test_MsgPack_Ext_Core2(HeadHex As String, DataLength As Long)
    Dim HeadBytes() As Byte
    HeadBytes = BitConverter.GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestBytes(HeadBytes, DataLength)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = GetTestValue(DataLength)
    
    Dim OutputValue() As Byte
    OutputValue = MsgPack_Ext.GetExtFromBytes(Bytes)
    
    DebugPrint_MsgPack_Ext_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = _
        MsgPack_Ext.GetBytesFromExt(HeadBytes(UBound(HeadBytes)), OutputValue)
    
    DebugPrint_MsgPack_Ext_GetBytes OutputValue, OutputBytes, Bytes
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
' MessagePack for VBA - Extension - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Ext_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    Dim HexString As String
    HexString = BitConverter.GetHexStringFromBytes(Value, , , " ")
    
    BitConverter.DebugPrint_GetBytes _
        HexString, OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Ext_GetValue( _
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

Attribute VB_Name = "Test_MsgPack_Ext_Dec"
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
' MessagePack for VBA - Extension - Decimal - Test
'

Public Sub Test_MsgPack_Ext_Dec()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Ext_Dec_FixExt1_TestCases
    Test_MsgPack_Ext_Dec_FixExt2_TestCases
    Test_MsgPack_Ext_Dec_FixExt4_TestCases
    Test_MsgPack_Ext_Dec_FixExt8_TestCases
    Test_MsgPack_Ext_Dec_Ext8_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Dec_FixExt1()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Dec_FixExt1_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Dec_FixExt2()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Dec_FixExt2_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Dec_FixExt4()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Dec_FixExt4_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Dec_FixExt8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Dec_FixExt8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Dec_Ext8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Dec_FixExt8_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Extension - Decimal - Test Cases
'

Public Sub Test_MsgPack_Ext_Dec_FixExt1_TestCases()
    Debug.Print "Target: Decimal FixExt1"
    
    Test_MsgPack_Ext_Dec_Core "D4 0E 00", CDec(0)
    Test_MsgPack_Ext_Dec_Core "D4 0E 01", CDec("1")
    Test_MsgPack_Ext_Dec_Core "D4 0E FF", CDec("255")
End Sub

Public Sub Test_MsgPack_Ext_Dec_FixExt2_TestCases()
    Debug.Print "Target: Decimal FixExt2"
    
    Test_MsgPack_Ext_Dec_Core "D5 0E 01 00", CDec("256")
    Test_MsgPack_Ext_Dec_Core "D5 0E FF FF", CDec("65535")
End Sub

Public Sub Test_MsgPack_Ext_Dec_FixExt4_TestCases()
    Debug.Print "Target: Decimal FixExt4"
    
    Test_MsgPack_Ext_Dec_Core "D6 0E 00 01 00 00", CDec("65536")
    Test_MsgPack_Ext_Dec_Core "D6 0E FF FF FF FF", CDec("4294967295")
End Sub

Public Sub Test_MsgPack_Ext_Dec_FixExt8_TestCases()
    Debug.Print "Target: Decimal FixExt8"
    
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E 00 00 00 01 00 00 00 00", CDec("4294967296")
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E 7F FF FF FF FF FF FF FF", CDec("9223372036854775807")
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E 80 00 00 00 00 00 00 00", CDec("9223372036854775808")
    Test_MsgPack_Ext_Dec_Core _
        "D7 0E FF FF FF FF FF FF FF FF", CDec("18446744073709551615")
    
End Sub

Public Sub Test_MsgPack_Ext_Dec_Ext8_TestCases()
    Debug.Print "Target: Decimal FixExt8"
    
    Test_MsgPack_Ext_Dec_Core _
        "C7 0C 0E 00 00 00 01 00 00 00 00 00 00 00 00", CDec("18446744073709551616")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0C 0E FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("79228162514264337593543950335")
    
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 00 1C 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("0.0000000000000000000000000001")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 00 1C FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("7.9228162514264337593543950335")
    
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 00 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("-1")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 00 FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("-79228162514264337593543950335")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 1C 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("-0.0000000000000000000000000001")
    Test_MsgPack_Ext_Dec_Core _
        "C7 0E 0E 80 1C FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("-7.9228162514264337593543950335")
End Sub

'
' MessagePack for VBA - Extension - Decimal - Test Core
'

Public Sub Test_MsgPack_Ext_Dec_Core( _
    HexBE As String, ExpectedValue As Variant)
    
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Variant
    OutputValue = MsgPack_Ext_Dec.GetExtDecFromBytes(BytesBE)
    
    DebugPrint_MsgPack_Ext_Dec_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Ext_Dec.GetBytesFromExtDec(OutputValue)
    
    DebugPrint_MsgPack_Ext_Dec_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

'
' MessagePack for VBA - Extension - Decimal - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Ext_Dec_GetBytes( _
    Value As Variant, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        CStr(Value), OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Ext_Dec_GetValue( _
    Bytes() As Byte, OutputValue As Variant, ExpectedValue As Variant)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

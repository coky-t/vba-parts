Attribute VB_Name = "Test_MsgPack_Float"
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
' MessagePack for VBA - Float - Test
'

Public Sub Test_MsgPack_Float()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Float_Float32_TestCases
    Test_MsgPack_Float_Float64_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Float_Float32()
    BitConverter.Test_Initialize
    Test_MsgPack_Float_Float32_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Float_Float64()
    BitConverter.Test_Initialize
    Test_MsgPack_Float_Float64_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Float - Test Cases
'

Public Sub Test_MsgPack_Float_Float32_TestCases()
    Debug.Print "Target: Float32"
    
    Test_MsgPack_Float_Float32_Core "41 46 00 00", 12.375!
    Test_MsgPack_Float_Float32_Core "3F 80 00 00", 1!
    Test_MsgPack_Float_Float32_Core "3F 00 00 00", 0.5
    Test_MsgPack_Float_Float32_Core "3E C0 00 00", 0.375
    Test_MsgPack_Float_Float32_Core "3E 80 00 00", 0.25
    Test_MsgPack_Float_Float32_Core "BF 80 00 00", -1!
    
    Test_MsgPack_Float_Core "CA 41 46 00 00", 12.375!
    Test_MsgPack_Float_Core "CA 3F 80 00 00", 1!
    Test_MsgPack_Float_Core "CA 3F 00 00 00", 0.5
    Test_MsgPack_Float_Core "CA 3E C0 00 00", 0.375
    Test_MsgPack_Float_Core "CA 3E 80 00 00", 0.25
    Test_MsgPack_Float_Core "CA BF 80 00 00", -1!
    
    ' Positive Zero
    Test_MsgPack_Float_Float32_Core "00 00 00 00", 0!
    Test_MsgPack_Float_Core "CA 00 00 00 00", 0!
    
    ' Positive SubNormal Minimum
    Test_MsgPack_Float_Float32_Core "00 00 00 01", 1.401298E-45
    Test_MsgPack_Float_Core "CA 00 00 00 01", 1.401298E-45
    
    ' Positive SubNormal Maximum
    Test_MsgPack_Float_Float32_Core "00 7F FF FF", 1.175494E-38
    Test_MsgPack_Float_Core "CA 00 7F FF FF", 1.175494E-38
    
    ' Positive Normal Minimum
    Test_MsgPack_Float_Float32_Core "00 80 00 00", 1.175494E-38
    Test_MsgPack_Float_Core "CA 00 80 00 00", 1.175494E-38
    
    ' Positive Normal Maximum
    Test_MsgPack_Float_Float32_Core "7F 7F FF FF", 3.402823E+38
    Test_MsgPack_Float_Core "CA 7F 7F FF FF", 3.402823E+38
    
    ' Positive Infinity
    Test_MsgPack_Float_Float32_Core "7F 80 00 00", "inf"
    Test_MsgPack_Float_Core "CA 7F 80 00 00", "inf"
    
    ' Positive NaN
    Test_MsgPack_Float_Float32_Core "7F FF FF FF", "nan"
    Test_MsgPack_Float_Core "CA 7F FF FF FF", "nan"
    
    ' Negative Zero
    Test_MsgPack_Float_Float32_Core "80 00 00 00", -0!
    Test_MsgPack_Float_Core "CA 80 00 00 00", -0!
    
    ' Negative SubNormal Minimum
    Test_MsgPack_Float_Float32_Core "80 00 00 01", -1.401298E-45
    Test_MsgPack_Float_Core "CA 80 00 00 01", -1.401298E-45
    
    ' Negative SubNormal Maximum
    Test_MsgPack_Float_Float32_Core "80 7F FF FF", -1.175494E-38
    Test_MsgPack_Float_Core "CA 80 7F FF FF", -1.175494E-38
    
    ' Negative Normal Minimum
    Test_MsgPack_Float_Float32_Core "80 80 00 00", -1.175494E-38
    Test_MsgPack_Float_Core "CA 80 80 00 00", -1.175494E-38
    
    ' Negative Normal Maximum
    Test_MsgPack_Float_Float32_Core "FF 7F FF FF", -3.402823E+38
    Test_MsgPack_Float_Core "CA FF 7F FF FF", -3.402823E+38
    
    ' Negative Infinity
    Test_MsgPack_Float_Float32_Core "FF 80 00 00", "-inf"
    Test_MsgPack_Float_Core "CA FF 80 00 00", "-inf"
    
    ' Negative NaN
    Test_MsgPack_Float_Float32_Core "FF FF FF FF", "-nan"
    Test_MsgPack_Float_Core "CA FF FF FF FF", "-nan"
End Sub

Public Sub Test_MsgPack_Float_Float64_TestCases()
    Debug.Print "Target: Float64"
    
    Test_MsgPack_Float_Float64_Core "40 28 C0 00 00 00 00 00", 12.375
    Test_MsgPack_Float_Float64_Core "3F F0 00 00 00 00 00 00", 1#
    Test_MsgPack_Float_Float64_Core "3F E0 00 00 00 00 00 00", 0.5
    Test_MsgPack_Float_Float64_Core "3F D8 00 00 00 00 00 00", 0.375
    Test_MsgPack_Float_Float64_Core "3F D0 00 00 00 00 00 00", 0.25
    Test_MsgPack_Float_Float64_Core "3F B9 99 99 99 99 99 9A", 0.1
    Test_MsgPack_Float_Float64_Core "3F D5 55 55 55 55 55 55", 1# / 3#
    Test_MsgPack_Float_Float64_Core "BF F0 00 00 00 00 00 00", -1#
    
    Test_MsgPack_Float_Core "CB 40 28 C0 00 00 00 00 00", 12.375
    Test_MsgPack_Float_Core "CB 3F F0 00 00 00 00 00 00", 1#
    Test_MsgPack_Float_Core "CB 3F E0 00 00 00 00 00 00", 0.5
    Test_MsgPack_Float_Core "CB 3F D8 00 00 00 00 00 00", 0.375
    Test_MsgPack_Float_Core "CB 3F D0 00 00 00 00 00 00", 0.25
    Test_MsgPack_Float_Core "CB 3F B9 99 99 99 99 99 9A", 0.1
    Test_MsgPack_Float_Core "CB 3F D5 55 55 55 55 55 55", 1# / 3#
    Test_MsgPack_Float_Core "CB BF F0 00 00 00 00 00 00", -1#
    
    ' Positive Zero
    Test_MsgPack_Float_Float64_Core "00 00 00 00 00 00 00 00", 0#
    Test_MsgPack_Float_Core "CB 00 00 00 00 00 00 00 00", 0#
    
    ' Positive SubNormal Minimum
    Test_MsgPack_Float_Float64_Core "00 00 00 00 00 00 00 01", _
        4.94065645841247E-324
    Test_MsgPack_Float_Core "CB 00 00 00 00 00 00 00 01", _
        4.94065645841247E-324
    
    ' Positive SubNormal Maximum
    Test_MsgPack_Float_Float64_Core "00 0F FF FF FF FF FF FF", _
        2.2250738585072E-308
    Test_MsgPack_Float_Core "CB 00 0F FF FF FF FF FF FF", _
        2.2250738585072E-308
    
    ' Positive Normal Minimum
    Test_MsgPack_Float_Float64_Core "00 10 00 00 00 00 00 00", _
        2.2250738585072E-308
    Test_MsgPack_Float_Core "CB 00 10 00 00 00 00 00 00", _
        2.2250738585072E-308
    
    ' Positive Normal Maximum
    Test_MsgPack_Float_Float64_Core "7F EF FF FF FF FF FF FF", _
        "1.79769313486232E+308"
    Test_MsgPack_Float_Core "CB 7F EF FF FF FF FF FF FF", _
        "1.79769313486232E+308"
    
    ' Positive Infinity
    Test_MsgPack_Float_Float64_Core "7F F0 00 00 00 00 00 00", "inf"
    Test_MsgPack_Float_Core "CB 7F F0 00 00 00 00 00 00", "inf"
    
    ' Positive NaN
    Test_MsgPack_Float_Float64_Core "7F FF FF FF FF FF FF FF", "nan"
    Test_MsgPack_Float_Core "CB 7F FF FF FF FF FF FF FF", "nan"
    
    ' Negative Zero
    Test_MsgPack_Float_Float64_Core "80 00 00 00 00 00 00 00", -0#
    Test_MsgPack_Float_Core "CB 80 00 00 00 00 00 00 00", -0#
    
    ' Negative SubNormal Minimum
    Test_MsgPack_Float_Float64_Core "80 00 00 00 00 00 00 01", _
        -4.94065645841247E-324
    Test_MsgPack_Float_Core "CB 80 00 00 00 00 00 00 01", _
        -4.94065645841247E-324
    
    ' Negative SubNormal Maximum
    Test_MsgPack_Float_Float64_Core "80 0F FF FF FF FF FF FF", _
        -2.2250738585072E-308
    Test_MsgPack_Float_Core "CB 80 0F FF FF FF FF FF FF", _
        -2.2250738585072E-308
    
    ' Negative Normal Minimum
    Test_MsgPack_Float_Float64_Core "80 10 00 00 00 00 00 00", _
        -2.2250738585072E-308
    Test_MsgPack_Float_Core "CB 80 10 00 00 00 00 00 00", _
        -2.2250738585072E-308
    
    ' Negative Normal Maximum
    Test_MsgPack_Float_Float64_Core "FF EF FF FF FF FF FF FF", _
        "-1.79769313486232E+308"
    Test_MsgPack_Float_Core "CB FF EF FF FF FF FF FF FF", _
        "-1.79769313486232E+308"
    
    ' Negative Infinity
    Test_MsgPack_Float_Float64_Core "FF F0 00 00 00 00 00 00", "-inf"
    Test_MsgPack_Float_Core "CB FF F0 00 00 00 00 00 00", "-inf"
    
    ' Negative NaN
    Test_MsgPack_Float_Float64_Core "FF FF FF FF FF FF FF FF", "-nan"
    Test_MsgPack_Float_Core "CB FF FF FF FF FF FF FF FF", "-nan"
End Sub

'
' MessagePack for VBA - Float - Test Core
'

Public Sub Test_MsgPack_Float_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Float.GetFloatFromBytes(BytesBE)
    
    DebugPrint_Float_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Float.GetBytesFromFloat(OutputValue)
    
    DebugPrint_Float_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Float_Float32_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetFloat32FromBytes(BytesBE, 0, True)
    
    DebugPrint_Float_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromFloat32(OutputValue, True)
    
    DebugPrint_Float_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Float_Float64_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetFloat64FromBytes(BytesBE, 0, True)
    
    DebugPrint_Float_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromFloat64(OutputValue, True)
    
    DebugPrint_Float_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

'
' MessagePack for VBA - Float - Test - Debug.Print
'

Private Sub DebugPrint_Float_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        CStr(Value), OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_Float_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    BitConverter.DebugPrint_GetValue Bytes, _
        CStr(OutputValue), CStr(ExpectedValue), _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

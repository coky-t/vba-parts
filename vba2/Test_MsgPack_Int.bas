Attribute VB_Name = "Test_MsgPack_Int"
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
' MessagePack for VBA - Integer - Test
'

Public Sub Test_MsgPack_Int()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Int_PositiveFixInt_TestCases
    Test_MsgPack_Int_NegativeFixInt_TestCases
    Test_MsgPack_Int_UInt8_TestCases
    Test_MsgPack_Int_UInt16_TestCases
    Test_MsgPack_Int_UInt32_TestCases
    Test_MsgPack_Int_UInt64_TestCases
    Test_MsgPack_Int_Int8_TestCases
    Test_MsgPack_Int_Int16_TestCases
    Test_MsgPack_Int_Int32_TestCases
    Test_MsgPack_Int_Int64_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_PositiveFixInt()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_PositiveFixInt_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_NegativeFixInt()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_NegativeFixInt_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_UInt8()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_UInt8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_UInt16()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_UInt16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_UInt32()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_UInt32_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_UInt64()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_UInt64_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_Int8()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_Int8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_Int16()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_Int16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_Int32()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_Int32_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Int_Int64()
    BitConverter.Test_Initialize
    Test_MsgPack_Int_Int64_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Integer - Test Cases
'

Public Sub Test_MsgPack_Int_PositiveFixInt_TestCases()
    Debug.Print "Target: PositiveFixInt"
    
    Test_MsgPack_Int_PositiveFixInt_Core "00", &H0
    Test_MsgPack_Int_PositiveFixInt_Core "7F", &H7F
    
    Test_MsgPack_Int_Core "00", &H0
    Test_MsgPack_Int_Core "7F", &H7F
End Sub

Public Sub Test_MsgPack_Int_NegativeFixInt_TestCases()
    Debug.Print "Target: NegativeFixInt"
    
    Test_MsgPack_Int_NegativeFixInt_Core "E0", -32
    Test_MsgPack_Int_NegativeFixInt_Core "FF", -1
    
    Test_MsgPack_Int_Core "E0", -32
    Test_MsgPack_Int_Core "FF", -1
End Sub

Public Sub Test_MsgPack_Int_UInt8_TestCases()
    Debug.Print "Target: UInt8"
    
    Test_MsgPack_Int_UInt8_Core "CC 00", &H0
    Test_MsgPack_Int_UInt8_Core "CC 01", &H1
    Test_MsgPack_Int_UInt8_Core "CC 7F", &H7F
    Test_MsgPack_Int_UInt8_Core "CC 80", &H80
    Test_MsgPack_Int_UInt8_Core "CC FF", &HFF
    
    Test_MsgPack_Int_Core "CC 80", &H80
    Test_MsgPack_Int_Core "CC FF", &HFF
End Sub

Public Sub Test_MsgPack_Int_UInt16_TestCases()
    Debug.Print "Target: UInt16"
    
    Test_MsgPack_Int_UInt16_Core "CD 00 00", &H0
    Test_MsgPack_Int_UInt16_Core "CD 00 01", &H1
    Test_MsgPack_Int_UInt16_Core "CD 00 FF", &HFF
    Test_MsgPack_Int_UInt16_Core "CD 01 00", &H100
    Test_MsgPack_Int_UInt16_Core "CD FF FF", &HFFFF&
    
    Test_MsgPack_Int_Core "CD 01 00", &H100
    Test_MsgPack_Int_Core "CD FF FF", &HFFFF&
End Sub

Public Sub Test_MsgPack_Int_UInt32_TestCases()
    Debug.Print "Target: UInt32"
    
    Test_MsgPack_Int_UInt32_Core "CE 00 00 00 00", 0
    Test_MsgPack_Int_UInt32_Core "CE 00 00 00 01", 1
    Test_MsgPack_Int_UInt32_Core "CE 00 00 FF FF", &HFFFF&
    Test_MsgPack_Int_UInt32_Core "CE 00 01 00 00", &H10000
    Test_MsgPack_Int_UInt32_Core "CE FF FF FF FF", CDec("4294967295")
    
    Test_MsgPack_Int_Core "CE 00 01 00 00", &H10000
    Test_MsgPack_Int_Core "CE 7F FF FF FF", &H7FFFFFFF
#If Win64 Then
    Test_MsgPack_Int_Core "CE FF FF FF FF", &HFFFFFFFF^
#End If
End Sub

Public Sub Test_MsgPack_Int_UInt64_TestCases()
    Debug.Print "Target: UInt64"

    Test_MsgPack_Int_UInt64_Core "CF 00 00 00 00 00 00 00 00", 0
    Test_MsgPack_Int_UInt64_Core "CF 00 00 00 00 00 00 00 01", 1
    Test_MsgPack_Int_UInt64_Core "CF 00 00 00 00 FF FF FF FF", _
        CDec("4294967295")
    Test_MsgPack_Int_UInt64_Core "CF 00 00 00 01 00 00 00 00", _
        CDec("&H100000000")
    Test_MsgPack_Int_UInt64_Core "CF FF FF FF FF FF FF FF FF", _
        CDec("18446744073709551615")
End Sub

Public Sub Test_MsgPack_Int_Int8_TestCases()
    Debug.Print "Target: Int8"
    Test_MsgPack_Int_Int8_Core "D0 00", 0
    Test_MsgPack_Int_Int8_Core "D0 01", 1
    Test_MsgPack_Int_Int8_Core "D0 7F", &H7F
    Test_MsgPack_Int_Int8_Core "D0 80", -128
    Test_MsgPack_Int_Int8_Core "D0 FF", -1
    
    Test_MsgPack_Int_Core "D0 DF", -33
    Test_MsgPack_Int_Core "D0 80", -128
End Sub

Public Sub Test_MsgPack_Int_Int16_TestCases()
    Debug.Print "Target: Int16"
    
    Test_MsgPack_Int_Int16_Core "D1 00 00", 0
    Test_MsgPack_Int_Int16_Core "D1 00 01", 1
    Test_MsgPack_Int_Int16_Core "D1 00 FF", &HFF
    Test_MsgPack_Int_Int16_Core "D1 01 00", &H100
    Test_MsgPack_Int_Int16_Core "D1 7F FF", &H7FFF
    Test_MsgPack_Int_Int16_Core "D1 80 00", CInt(-32768)
    Test_MsgPack_Int_Int16_Core "D1 FF FF", -1
    
    Test_MsgPack_Int_Core "D1 FF 7F", -129
    Test_MsgPack_Int_Core "D1 80 00", CInt(-32768)
End Sub

Public Sub Test_MsgPack_Int_Int32_TestCases()
    Debug.Print "Target: Int32"
    
    Test_MsgPack_Int_Int32_Core "D2 00 00 00 00", 0
    Test_MsgPack_Int_Int32_Core "D2 00 00 00 01", 1
    Test_MsgPack_Int_Int32_Core "D2 00 00 FF FF", &HFFFF&
    Test_MsgPack_Int_Int32_Core "D2 00 01 00 00", &H10000
    Test_MsgPack_Int_Int32_Core "D2 7F FF FF FF", &H7FFFFFFF
    Test_MsgPack_Int_Int32_Core "D2 80 00 00 00", CLng("-2147483648")
    Test_MsgPack_Int_Int32_Core "D2 FF FF FF FF", -1
    
#If Win64 Then
    Test_MsgPack_Int_Core "D2 FF FF 7F FF", CLng("-32769")
    Test_MsgPack_Int_Core "D2 80 00 00 00", CLng("-2147483648")
#Else
    Test_MsgPack_Int_Core "D2 00 01 00 00", &H10000
    Test_MsgPack_Int_Core "D2 7F FF FF FF", &H7FFFFFFF
#End If
End Sub

Public Sub Test_MsgPack_Int_Int64_TestCases()
    Debug.Print "Target: Int64"

    Test_MsgPack_Int_Int64_Core "D3 00 00 00 00 00 00 00 00", 0
    Test_MsgPack_Int_Int64_Core "D3 00 00 00 00 00 00 00 01", 1
    Test_MsgPack_Int_Int64_Core "D3 00 00 00 00 FF FF FF FF", _
        CDec("4294967295")
    Test_MsgPack_Int_Int64_Core "D3 00 00 00 01 00 00 00 00", _
        CDec("&H100000000")
    Test_MsgPack_Int_Int64_Core "D3 7F FF FF FF FF FF FF FF", _
        CDec("&H7FFFFFFFFFFFFFFF")
    Test_MsgPack_Int_Int64_Core "D3 80 00 00 00 00 00 00 00", _
        CDec("-9223372036854775808")
    Test_MsgPack_Int_Int64_Core "D3 FF FF FF FF FF FF FF FF", -1
    
#If Win64 Then
    Test_MsgPack_Int_Core "D3 00 00 00 01 00 00 00 00", _
        CDec("&H100000000")
    Test_MsgPack_Int_Core "D3 7F FF FF FF FF FF FF FF", _
        CDec("&H7FFFFFFFFFFFFFFF")
    Test_MsgPack_Int_Core "D3 80 00 00 00 00 00 00 00", _
        CDec("-9223372036854775808")
    Test_MsgPack_Int_Core "D3 FF FF FF FF 7F FF FF FF", _
        CDec("-2147483649")
#End If
End Sub

'
' MessagePack for VBA - Integer - Test Core
'

Public Sub Test_MsgPack_Int_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromInt(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_PositiveFixInt_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromPositiveFixInt(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_UInt8_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromUInt8(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_UInt16_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromUInt16(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_UInt32_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromUInt32(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_UInt64_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromUInt64(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_Int8_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromInt8(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_Int16_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromInt16(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_Int32_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromInt32(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_Int64_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromInt64(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_MsgPack_Int_NegativeFixInt_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = MsgPack_Int.GetIntFromBytes(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Int.GetBytesFromNegativeFixInt(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

'
' MessagePack for VBA - Integer - Test - Debug.Print
'

Private Sub DebugPrint_Int_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    If VarType(Value) = vbDecimal Then
        BitConverter.DebugPrint_GetBytes _
            CStr(Value), OutputBytes, ExpectedBytes
    Else
        BitConverter.DebugPrint_GetBytes _
            CStr(Value) & " (" & Hex(Value) & ")", OutputBytes, ExpectedBytes
    End If
End Sub

Private Sub DebugPrint_Int_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    If (VarType(OutputValue) = vbDecimal) Or _
        (VarType(ExpectedValue) = vbDecimal) Then
        
        BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
            CStr(OutputValue), CStr(ExpectedValue)
    Else
        BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
            CStr(OutputValue) & " (" & Hex(OutputValue) & ")", _
            CStr(ExpectedValue) & " (" & Hex(ExpectedValue) & ")"
    End If
End Sub

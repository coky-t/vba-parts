Attribute VB_Name = "Test_BitConverter"
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

#If Win64 Then
#Const USE_LONGLONG = True
#End If

'
' BitConverter - Test
'

Public Sub Test_BitConv()
    BitConverter.Test_Initialize
    
    Test_BitConv_UInt16_TestCases
    Test_BitConv_UInt32_TestCases
    Test_BitConv_UInt64_TestCases
    Test_BitConv_Int8_TestCases
    Test_BitConv_Int16_TestCases
    Test_BitConv_Int32_TestCases
    Test_BitConv_Int64_TestCases
    Test_BitConv_Integer_TestCases
    Test_BitConv_Long_TestCases
#If USE_LONGLONG Then
    Test_BitConv_LongLong_TestCases
#End If
    Test_BitConv_Single_TestCases
    Test_BitConv_Double_TestCases
    Test_BitConv_Currency_TestCases
    Test_BitConv_Date_TestCases
    Test_BitConv_Decimal_TestCases
    Test_BitConv_String_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_UInt16()
    BitConverter.Test_Initialize
    Test_BitConv_UInt16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_UInt32()
    BitConverter.Test_Initialize
    Test_BitConv_UInt32_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_UInt64()
    BitConverter.Test_Initialize
    Test_BitConv_UInt64_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Int8()
    BitConverter.Test_Initialize
    Test_BitConv_Int8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Int16()
    BitConverter.Test_Initialize
    Test_BitConv_Int16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Int32()
    BitConverter.Test_Initialize
    Test_BitConv_Int32_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Int64()
    BitConverter.Test_Initialize
    Test_BitConv_Int64_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Integer()
    BitConverter.Test_Initialize
    Test_BitConv_Integer_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Long()
    BitConverter.Test_Initialize
    Test_BitConv_Long_TestCases
    BitConverter.Test_Terminate
End Sub

#If USE_LONGLONG Then
Public Sub Test_BitConv_LongLong()
    BitConverter.Test_Initialize
    Test_BitConv_LongLong_TestCases
    BitConverter.Test_Terminate
End Sub
#End If

Public Sub Test_BitConv_Single()
    BitConverter.Test_Initialize
    Test_BitConv_Single_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Double()
    BitConverter.Test_Initialize
    Test_BitConv_Double_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Currency()
    BitConverter.Test_Initialize
    Test_BitConv_Currency_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Date()
    BitConverter.Test_Initialize
    Test_BitConv_Date_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_Decimal()
    BitConverter.Test_Initialize
    Test_BitConv_Decimal_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_BitConv_String()
    BitConverter.Test_Initialize
    Test_BitConv_String_TestCases
    BitConverter.Test_Terminate
End Sub

'
' BitConverter - Test Cases
'

Public Sub Test_BitConv_UInt16_TestCases()
    Debug.Print "Target: UInt16"
    Test_BitConv_UInt16_Core "00 00", 0
    Test_BitConv_UInt16_Core "00 01", 1
    Test_BitConv_UInt16_Core "00 FF", &HFF
    Test_BitConv_UInt16_Core "01 00", &H100
    Test_BitConv_UInt16_Core "FF FF", 65535
End Sub

Public Sub Test_BitConv_UInt32_TestCases()
    Debug.Print "Target: UInt32"
    Test_BitConv_UInt32_Core "00 00 00 00", 0
    Test_BitConv_UInt32_Core "00 00 00 01", 1
    Test_BitConv_UInt32_Core "00 00 FF FF", &HFFFF&
    Test_BitConv_UInt32_Core "00 01 00 00", &H10000
    Test_BitConv_UInt32_Core "FF FF FF FF", CDec("4294967295")
End Sub

Public Sub Test_BitConv_UInt64_TestCases()
    Debug.Print "Target: UInt64"
    Test_BitConv_UInt64_Core "00 00 00 00 00 00 00 00", 0
    Test_BitConv_UInt64_Core "00 00 00 00 00 00 00 01", 1
    Test_BitConv_UInt64_Core "00 00 00 00 FF FF FF FF", CDec("4294967295")
    Test_BitConv_UInt64_Core "00 00 00 01 00 00 00 00", CDec("&H100000000")
    Test_BitConv_UInt64_Core "FF FF FF FF FF FF FF FF", _
        CDec("18446744073709551615")
End Sub

Public Sub Test_BitConv_Int8_TestCases()
    Debug.Print "Target: Int8"
    Test_BitConv_Int8_Core "00", 0
    Test_BitConv_Int8_Core "01", 1
    Test_BitConv_Int8_Core "7F", &H7F
    Test_BitConv_Int8_Core "80", -128
    Test_BitConv_Int8_Core "FF", -1
End Sub

Public Sub Test_BitConv_Int16_TestCases()
    Debug.Print "Target: Int16"
    Test_BitConv_Int16_Core "00 00", 0
    Test_BitConv_Int16_Core "00 01", 1
    Test_BitConv_Int16_Core "00 FF", &HFF
    Test_BitConv_Int16_Core "01 00", &H100
    Test_BitConv_Int16_Core "7F FF", &H7FFF
    Test_BitConv_Int16_Core "80 00", CInt(-32768)
    Test_BitConv_Int16_Core "FF FF", -1
End Sub

Public Sub Test_BitConv_Int32_TestCases()
    Debug.Print "Target: Int32"
    Test_BitConv_Int32_Core "00 00 00 00", 0
    Test_BitConv_Int32_Core "00 00 00 01", 1
    Test_BitConv_Int32_Core "00 00 FF FF", &HFFFF&
    Test_BitConv_Int32_Core "00 01 00 00", &H10000
    Test_BitConv_Int32_Core "7F FF FF FF", &H7FFFFFFF
    Test_BitConv_Int32_Core "80 00 00 00", CLng("-2147483648")
    Test_BitConv_Int32_Core "FF FF FF FF", -1
End Sub

Public Sub Test_BitConv_Int64_TestCases()
    Debug.Print "Target: Int64"
    Test_BitConv_Int64_Core "00 00 00 00 00 00 00 00", 0
    Test_BitConv_Int64_Core "00 00 00 00 00 00 00 01", 1
    Test_BitConv_Int64_Core "00 00 00 00 FF FF FF FF", CDec("4294967295")
    Test_BitConv_Int64_Core "00 00 00 01 00 00 00 00", CDec("&H100000000")
    Test_BitConv_Int64_Core "7F FF FF FF FF FF FF FF", _
        CDec("&H7FFFFFFFFFFFFFFF")
    Test_BitConv_Int64_Core "80 00 00 00 00 00 00 00", _
        CDec("-9223372036854775808")
    Test_BitConv_Int64_Core "FF FF FF FF FF FF FF FF", -1
End Sub

Public Sub Test_BitConv_Integer_TestCases()
    Debug.Print "Target: Integer"
    Test_BitConv_Integer_Core "00 00", 0
    Test_BitConv_Integer_Core "00 01", 1
    Test_BitConv_Integer_Core "00 FF", &HFF
    Test_BitConv_Integer_Core "01 00", &H100
    Test_BitConv_Integer_Core "7F FF", &H7FFF
    Test_BitConv_Integer_Core "80 00", CInt(-32768)
    Test_BitConv_Integer_Core "FF 7F", -129
    Test_BitConv_Integer_Core "FF FF", -1
End Sub

Public Sub Test_BitConv_Long_TestCases()
    Debug.Print "Target: Long"
    Test_BitConv_Long_Core "00 00 00 00", 0
    Test_BitConv_Long_Core "00 00 00 01", 1
    Test_BitConv_Long_Core "00 00 FF FF", &HFFFF&
    Test_BitConv_Long_Core "00 01 00 00", &H10000
    Test_BitConv_Long_Core "7F FF FF FF", &H7FFFFFFF
    Test_BitConv_Long_Core "80 00 00 00", CLng("-2147483648")
    Test_BitConv_Long_Core "FF FF FF FF", -1
End Sub

#If USE_LONGLONG Then
Public Sub Test_BitConv_LongLong_TestCases()
    Debug.Print "Target: LongLong"
    Test_BitConv_LongLong_Core "00 00 00 00 00 00 00 00", 0
    Test_BitConv_LongLong_Core "00 00 00 00 00 00 00 01", 1
    Test_BitConv_LongLong_Core "00 00 00 00 FF FF FF FF", _
        CLngLng("4294967295")
    Test_BitConv_LongLong_Core "00 00 00 01 00 00 00 00", _
        CLngLng("&H100000000")
    Test_BitConv_LongLong_Core "7F FF FF FF FF FF FF FF", _
        CLngLng("9223372036854775807")
    Test_BitConv_LongLong_Core "80 00 00 00 00 00 00 00", _
        CLngLng("-9223372036854775808")
    Test_BitConv_LongLong_Core "FF FF FF FF FF FF FF FF", CLngLng(-1)
End Sub
#End If

Public Sub Test_BitConv_Single_TestCases()
    Debug.Print "Target: Single"
    
    Test_BitConv_Single_Core "41 46 00 00", 12.375!
    Test_BitConv_Single_Core "3F 80 00 00", 1!
    Test_BitConv_Single_Core "3F 00 00 00", 0.5
    Test_BitConv_Single_Core "3E C0 00 00", 0.375
    Test_BitConv_Single_Core "3E 80 00 00", 0.25
    Test_BitConv_Single_Core "BF 80 00 00", -1!
    
    ' Positive Zero
    Test_BitConv_Single_Core "00 00 00 00", 0!
    
    ' Positive SubNormal Minimum
    Test_BitConv_Single_Core "00 00 00 01", 1.401298E-45
    
    ' Positive SubNormal Maximum
    Test_BitConv_Single_Core "00 7F FF FF", 1.175494E-38
    
    ' Positive Normal Minimum
    Test_BitConv_Single_Core "00 80 00 00", 1.175494E-38
    
    ' Positive Normal Maximum
    Test_BitConv_Single_Core "7F 7F FF FF", 3.402823E+38
    
    ' Positive Infinity
    Test_BitConv_Single_Core "7F 80 00 00", "inf"
    
    ' Positive NaN
    Test_BitConv_Single_Core "7F FF FF FF", "nan"
    
    ' Negative Zero
    Test_BitConv_Single_Core "80 00 00 00", -0!
    
    ' Negative SubNormal Minimum
    Test_BitConv_Single_Core "80 00 00 01", -1.401298E-45
    
    ' Negative SubNormal Maximum
    Test_BitConv_Single_Core "80 7F FF FF", -1.175494E-38
    
    ' Negative Normal Minimum
    Test_BitConv_Single_Core "80 80 00 00", -1.175494E-38
    
    ' Negative Normal Maximum
    Test_BitConv_Single_Core "FF 7F FF FF", -3.402823E+38
    
    ' Negative Infinity
    Test_BitConv_Single_Core "FF 80 00 00", "-inf"
    
    ' Negative NaN
    Test_BitConv_Single_Core "FF FF FF FF", "-nan"
End Sub

Public Sub Test_BitConv_Double_TestCases()
    Debug.Print "Target: Double"
    
    Test_BitConv_Double_Core "40 28 C0 00 00 00 00 00", 12.375
    Test_BitConv_Double_Core "3F F0 00 00 00 00 00 00", 1#
    Test_BitConv_Double_Core "3F E0 00 00 00 00 00 00", 0.5
    Test_BitConv_Double_Core "3F D8 00 00 00 00 00 00", 0.375
    Test_BitConv_Double_Core "3F D0 00 00 00 00 00 00", 0.25
    Test_BitConv_Double_Core "3F B9 99 99 99 99 99 9A", 0.1
    Test_BitConv_Double_Core "3F D5 55 55 55 55 55 55", 1# / 3#
    Test_BitConv_Double_Core "BF F0 00 00 00 00 00 00", -1#
    
    ' Positive Zero
    Test_BitConv_Double_Core "00 00 00 00 00 00 00 00", 0#
    
    ' Positive SubNormal Minimum
    Test_BitConv_Double_Core "00 00 00 00 00 00 00 01", 4.94065645841247E-324
    
    ' Positive SubNormal Maximum
    Test_BitConv_Double_Core "00 0F FF FF FF FF FF FF", 2.2250738585072E-308
    
    ' Positive Normal Minimum
    Test_BitConv_Double_Core "00 10 00 00 00 00 00 00", 2.2250738585072E-308
    
    ' Positive Normal Maximum
    Test_BitConv_Double_Core "7F EF FF FF FF FF FF FF", "1.79769313486232E+308"
    
    ' Positive Infinity
    Test_BitConv_Double_Core "7F F0 00 00 00 00 00 00", "inf"
    
    ' Positive NaN
    Test_BitConv_Double_Core "7F FF FF FF FF FF FF FF", "nan"
    
    ' Negative Zero
    Test_BitConv_Double_Core "80 00 00 00 00 00 00 00", -0#
    
    ' Negative SubNormal Minimum
    Test_BitConv_Double_Core "80 00 00 00 00 00 00 01", -4.94065645841247E-324
    
    ' Negative SubNormal Maximum
    Test_BitConv_Double_Core "80 0F FF FF FF FF FF FF", -2.2250738585072E-308
    
    ' Negative Normal Minimum
    Test_BitConv_Double_Core "80 10 00 00 00 00 00 00", -2.2250738585072E-308
    
    ' Negative Normal Maximum
    Test_BitConv_Double_Core "FF EF FF FF FF FF FF FF", "-1.79769313486232E+308"
    
    ' Negative Infinity
    Test_BitConv_Double_Core "FF F0 00 00 00 00 00 00", "-inf"
    
    ' Negative NaN
    Test_BitConv_Double_Core "FF FF FF FF FF FF FF FF", "-nan"
End Sub

Public Sub Test_BitConv_Currency_TestCases()
    Debug.Print "Target: Currency"
    Test_BitConv_Currency_Core _
        "FF FF FF FF FF FF FF FF", CCur("-0.0001")
    Test_BitConv_Currency_Core _
        "FF FF FF FF FF FF D8 F0", CCur("-1")
    Test_BitConv_Currency_Core _
        "80 00 00 00 00 00 00 00", CCur("-922337203685477.5808")
    Test_BitConv_Currency_Core _
        "00 00 00 00 00 00 27 10", CCur("1")
    Test_BitConv_Currency_Core _
        "00 00 00 00 00 00 00 00", 0@
    Test_BitConv_Currency_Core _
        "7F FF FF FF FF FF FF FF", _
        CCur("922337203685477.5807")
End Sub

Public Sub Test_BitConv_Date_TestCases()
    Debug.Print "Target: Date"
    Test_BitConv_Date_Core "C1 24 10 34 00 00 00 00", DateSerial(100, 1, 1)
    Test_BitConv_Date_Core "00 00 00 00 00 00 00 00", DateSerial(1899, 12, 30)
    Test_BitConv_Date_Core "41 46 92 40 80 00 00 00", DateSerial(9999, 12, 31)
    
    Test_BitConv_Date_Core "00 00 00 00 00 00 00 00", TimeSerial(0, 0, 0)
    Test_BitConv_Date_Core "3F E0 00 00 00 00 00 00", TimeSerial(12, 0, 0)
    Test_BitConv_Date_Core "3F EF FF E7 BA 37 5F 32", TimeSerial(23, 59, 59)
    
    Test_BitConv_Date_Core _
        "41 46 92 40 FF FF 9E E9", _
        DateSerial(9999, 12, 31) + TimeSerial(23, 59, 59)
End Sub

Public Sub Test_BitConv_Decimal_TestCases()
    Debug.Print "Target: Decimal"
    Test_BitConv_Decimal_Core _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 00", _
        CDec(0)
    Test_BitConv_Decimal_Core _
        "00 00 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec(1)
    Test_BitConv_Decimal_Core _
        "80 00 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec(-1)
    Test_BitConv_Decimal_Core _
        "00 00 00 00 00 00 FF FF FF FF FF FF FF FF", _
        CDec("18446744073709551615")
    Test_BitConv_Decimal_Core _
        "80 00 00 00 00 00 FF FF FF FF FF FF FF FF", _
        CDec("-18446744073709551615")
    Test_BitConv_Decimal_Core _
        "00 00 00 00 00 01 00 00 00 00 00 00 00 00", _
        CDec("18446744073709551616")
    Test_BitConv_Decimal_Core _
        "80 00 00 00 00 01 00 00 00 00 00 00 00 00", _
        CDec("-18446744073709551616")
    
    ' With a scale of 0 (no decimal places), the largest possible value
    Test_BitConv_Decimal_Core _
        "00 00 FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("79228162514264337593543950335")
    Test_BitConv_Decimal_Core _
        "80 00 FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("-79228162514264337593543950335")
    
    ' With a scale of 28 decimal places, the smallest, non-zero value
    Test_BitConv_Decimal_Core _
        "00 1C 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("0.0000000000000000000000000001")
    Test_BitConv_Decimal_Core _
        "80 1C 00 00 00 00 00 00 00 00 00 00 00 01", _
        CDec("-0.0000000000000000000000000001")
    
    ' With a scale of 28 decimal places, the largest value
    Test_BitConv_Decimal_Core _
        "00 1C FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("7.9228162514264337593543950335")
    Test_BitConv_Decimal_Core _
        "80 1C FF FF FF FF FF FF FF FF FF FF FF FF", _
        CDec("-7.9228162514264337593543950335")
End Sub

Public Sub Test_BitConv_String_TestCases()
    Debug.Print "Target: String"
    Test_BitConv_String_Core Hex(Asc("a")), "a"
    Test_BitConv_String_Core "E3 81 82", ChrW(&H3042)
End Sub

'
' BitConverter - Test Core
'

Public Sub Test_BitConv_UInt16_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetUInt16FromBytes(BytesBE, 0, True)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromUInt16(OutputValue, True)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_UInt32_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetUInt32FromBytes(BytesBE, 0, True)
    
#If USE_LONGLONG Then
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
#Else
    DebugPrint_Dec_GetValue BytesBE, OutputValue, ExpectedValue
#End If
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromUInt32(OutputValue, True)
    
#If USE_LONGLONG Then
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
#Else
    DebugPrint_Dec_GetBytes OutputValue, OutputBytesBE, BytesBE
#End If
End Sub

Public Sub Test_BitConv_UInt64_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetUInt64FromBytes(BytesBE, 0, True)
    
    DebugPrint_Dec_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromUInt64(OutputValue, True)
    
    DebugPrint_Dec_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Int8_Core(Hex, ExpectedValue)
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytesFromHexString(Hex)
    
    Dim OutputValue
    OutputValue = BitConverter.GetInt8FromBytes(Bytes, 0, True)
    
    DebugPrint_Int_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = BitConverter.GetBytesFromInt8(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Public Sub Test_BitConv_Int16_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetInt16FromBytes(BytesBE, 0, True)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromInt16(OutputValue, True)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Int32_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetInt32FromBytes(BytesBE, 0, True)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromInt32(OutputValue, True)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Int64_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetInt64FromBytes(BytesBE, 0, True)
    
#If USE_LONGLONG Then
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
#Else
    DebugPrint_Dec_GetValue BytesBE, OutputValue, ExpectedValue
#End If
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromInt64(OutputValue, True)
    
#If USE_LONGLONG Then
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
#Else
    DebugPrint_Dec_GetBytes OutputValue, OutputBytesBE, BytesBE
#End If
End Sub

Public Sub Test_BitConv_Integer_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetIntegerFromBytes(BytesBE, 0, True)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromInteger(OutputValue, True)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Long_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetLongFromBytes(BytesBE, 0, True)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromLong(OutputValue, True)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

#If USE_LONGLONG Then
Public Sub Test_BitConv_LongLong_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetLongLongFromBytes(BytesBE, 0, True)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromLongLong(OutputValue, True)
    
    DebugPrint_Int_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub
#End If

Public Sub Test_BitConv_Single_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetSingleFromBytes(BytesBE, 0, True)
    
    DebugPrint_Float_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromSingle(OutputValue, True)
    
    DebugPrint_Float_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Double_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetDoubleFromBytes(BytesBE, 0, True)
    
    DebugPrint_Float_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromDouble(OutputValue, True)
    
    DebugPrint_Float_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Currency_Core(HexBE, ExpectedValue As Currency)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Currency
    OutputValue = BitConverter.GetCurrencyFromBytes(BytesBE, 0, True)
    
    DebugPrint_Cur_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromCurrency(OutputValue, True)
    
    DebugPrint_Cur_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Date_Core(HexBE, ExpectedValue As Date)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Date
    OutputValue = BitConverter.GetDateFromBytes(BytesBE, 0, True)
    
    DebugPrint_Date_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromDate(OutputValue, True)
    
    DebugPrint_Date_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_Decimal_Core(HexBE, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = BitConverter.GetDecimalFromBytes(BytesBE, 0, True)
    
    DebugPrint_Dec_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = BitConverter.GetBytesFromDecimal(OutputValue, True)
    
    DebugPrint_Dec_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

Public Sub Test_BitConv_String_Core(HexStr, ExpectedValue)
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytesFromHexString(HexStr)
    
    Dim OutputValue
    OutputValue = BitConverter.GetStringFromBytes(Bytes)
    
    DebugPrint_Str_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = BitConverter.GetBytesFromString(OutputValue)
    
    DebugPrint_Str_GetBytes OutputValue, OutputBytes, Bytes
End Sub

'
' BitConverter - Debug.Print
'

Private Sub DebugPrint_Int_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        CStr(Value) & " (" & Hex(Value) & ")", OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_Int_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        CStr(OutputValue) & " (" & Hex(OutputValue) & ")", _
        CStr(ExpectedValue) & " (" & Hex(ExpectedValue) & ")"
End Sub

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

Private Sub DebugPrint_Cur_GetBytes( _
    Value As Currency, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        CStr(Value), OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_Cur_GetValue( _
    Bytes() As Byte, OutputValue As Currency, ExpectedValue As Currency)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

Private Sub DebugPrint_Date_GetBytes( _
    Value As Date, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        FormatDateTime(Value, vbLongDate) & " " & _
        FormatDateTime(Value, vbLongTime), _
        OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_Date_GetValue( _
    Bytes() As Byte, OutputValue As Date, ExpectedValue As Date)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        FormatDateTime(OutputValue, vbLongDate) & " " & _
        FormatDateTime(OutputValue, vbLongTime), _
        FormatDateTime(ExpectedValue, vbLongDate) & " " & _
        FormatDateTime(ExpectedValue, vbLongTime)
End Sub

Private Sub DebugPrint_Dec_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        CStr(Value), OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_Dec_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

Private Sub DebugPrint_Str_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        Value, OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_Str_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        OutputValue, ExpectedValue
End Sub

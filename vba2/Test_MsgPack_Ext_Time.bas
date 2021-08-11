Attribute VB_Name = "Test_MsgPack_Ext_Time"
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
' MessagePack for VBA - Extension - Timestamp - Test
'

Public Sub Test_MsgPack_Ext_Time()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Ext_Time_FixExt4_TestCases
    Test_MsgPack_Ext_Time_FixExt8_TestCases
    Test_MsgPack_Ext_Time_Ext8_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Time_FixExt4()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Time_FixExt4_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Time_FixExt8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Time_FixExt8_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Time_Ext8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Time_Ext8_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Extension - Timestamp - Test Cases
'

Public Sub Test_MsgPack_Ext_Time_FixExt4_TestCases()
    Debug.Print "Target: Timestamp FixExt4"
    
    Test_MsgPack_Ext_Time_Core _
        "D6 FF 00 00 00 00", DateSerial(1970, 1, 1)
    Test_MsgPack_Ext_Time_Core _
        "D6 FF 7F FF FF FF", DateSerial(2038, 1, 19) + TimeSerial(3, 14, 7)
    Test_MsgPack_Ext_Time_Core _
        "D6 FF FF FF FF FF", DateSerial(2106, 2, 7) + TimeSerial(6, 28, 15)
End Sub

Public Sub Test_MsgPack_Ext_Time_FixExt8_TestCases()
    Debug.Print "Target: Timestamp FixExt8"
    
    Test_MsgPack_Ext_Time_Core _
        "D7 FF 00 00 00 01 00 00 00 00", _
        DateSerial(2106, 2, 7) + TimeSerial(6, 28, 16)
    Test_MsgPack_Ext_Time_Core _
        "D7 FF 00 00 00 03 FF FF FF FF", _
        DateSerial(2514, 5, 30) + TimeSerial(1, 53, 3)
End Sub

Public Sub Test_MsgPack_Ext_Time_Ext8_TestCases()
    Debug.Print "Target: Timestamp Ext8"
    
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 FF FF FF F2 42 A4 97 80", _
        DateSerial(100, 1, 1)
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 FF FF FF FF FF FF FF FF", _
        DateSerial(1969, 12, 31) + TimeSerial(23, 59, 59)
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 00 00 00 04 00 00 00 00", _
        DateSerial(2514, 5, 30) + TimeSerial(1, 53, 4)
    Test_MsgPack_Ext_Time_Core _
        "C7 0C FF 00 00 00 00 00 00 00 3A FF F4 41 7F", _
        DateSerial(9999, 12, 31) + TimeSerial(23, 59, 59)
End Sub

'
' MessagePack for VBA - Extension - Timestamp - Test Core
'

Public Sub Test_MsgPack_Ext_Time_Core( _
    HexBE As String, ExpectedValue As Date)
    
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Date
    OutputValue = MsgPack_Ext_Time.GetExtTimeFromBytes(BytesBE)
    
    DebugPrint_MsgPack_Ext_Time_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Ext_Time.GetBytesFromExtTime(OutputValue)
    
    DebugPrint_MsgPack_Ext_Time_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

'
' MessagePack for VBA - Extension - Timestamp - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Ext_Time_GetBytes( _
    Value As Date, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        FormatDateTime(Value, vbLongDate) & " " & _
        FormatDateTime(Value, vbLongTime), _
        OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Ext_Time_GetValue( _
    Bytes() As Byte, OutputValue As Date, ExpectedValue As Date)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        FormatDateTime(OutputValue, vbLongDate) & " " & _
        FormatDateTime(OutputValue, vbLongTime), _
        FormatDateTime(ExpectedValue, vbLongDate) & " " & _
        FormatDateTime(ExpectedValue, vbLongTime)
End Sub

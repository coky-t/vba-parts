Attribute VB_Name = "Test_MsgPack_Ext_Date"
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
' MessagePack for VBA - Extension - Date - Test
'

Public Sub Test_MsgPack_Ext_Date()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Ext_Date_FixExt8_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Date_FixExt8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Date_FixExt8_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Extension - Date - Test Cases
'

Public Sub Test_MsgPack_Ext_Date_FixExt8_TestCases()
    Debug.Print "Target: Date FixExt8"
    
    Test_MsgPack_Ext_Date_Core _
        "D7 07 C1 24 10 34 00 00 00 00", DateSerial(100, 1, 1)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 00 00 00 00 00 00 00 00", DateSerial(1899, 12, 30)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 41 46 92 40 80 00 00 00", DateSerial(9999, 12, 31)
    
    Test_MsgPack_Ext_Date_Core _
        "D7 07 00 00 00 00 00 00 00 00", TimeSerial(0, 0, 0)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 3F E0 00 00 00 00 00 00", TimeSerial(12, 0, 0)
    Test_MsgPack_Ext_Date_Core _
        "D7 07 3F EF FF E7 BA 37 5F 32", TimeSerial(23, 59, 59)
    
    Test_MsgPack_Ext_Date_Core _
        "D7 07 41 46 92 40 FF FF 9E E9", _
        DateSerial(9999, 12, 31) + TimeSerial(23, 59, 59)
End Sub

'
' MessagePack for VBA - Extension - Date - Test Core
'

Public Sub Test_MsgPack_Ext_Date_Core( _
    HexBE As String, ExpectedValue As Date)
    
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Date
    OutputValue = MsgPack_Ext_Date.GetExtDateFromBytes(BytesBE)
    
    DebugPrint_MsgPack_Ext_Date_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Ext_Date.GetBytesFromExtDate(OutputValue)
    
    DebugPrint_MsgPack_Ext_Date_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

'
' MessagePack for VBA - Extension - Date - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Ext_Date_GetBytes( _
    Value As Date, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        FormatDateTime(Value, vbLongDate) & " " & _
        FormatDateTime(Value, vbLongTime), _
        OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Ext_Date_GetValue( _
    Bytes() As Byte, OutputValue As Date, ExpectedValue As Date)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        FormatDateTime(OutputValue, vbLongDate) & " " & _
        FormatDateTime(OutputValue, vbLongTime), _
        FormatDateTime(ExpectedValue, vbLongDate) & " " & _
        FormatDateTime(ExpectedValue, vbLongTime)
End Sub

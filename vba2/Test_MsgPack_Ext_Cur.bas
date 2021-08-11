Attribute VB_Name = "Test_MsgPack_Ext_Cur"
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
' MessagePack for VBA - Extension - Currency - Test
'

Public Sub Test_MsgPack_Ext_Cur()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Ext_Cur_FixExt1_TestCases
    Test_MsgPack_Ext_Cur_FixExt2_TestCases
    Test_MsgPack_Ext_Cur_FixExt4_TestCases
    Test_MsgPack_Ext_Cur_FixExt8_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Cur_FixExt1()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Cur_FixExt1_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Cur_FixExt2()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Cur_FixExt2_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Cur_FixExt4()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Cur_FixExt4_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Ext_Cur_FixExt8()
    BitConverter.Test_Initialize
    Test_MsgPack_Ext_Cur_FixExt8_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Extension - Currency - Test Cases
'

Public Sub Test_MsgPack_Ext_Cur_FixExt1_TestCases()
    Debug.Print "Target: Currency FixExt1"
    
    Test_MsgPack_Ext_Cur_Core "D4 06 00", 0@
    Test_MsgPack_Ext_Cur_Core "D4 06 01", CCur("0.0001")
    Test_MsgPack_Ext_Cur_Core "D4 06 FF", CCur("0.0255")
End Sub

Public Sub Test_MsgPack_Ext_Cur_FixExt2_TestCases()
    Debug.Print "Target: Currency FixExt2"
    
    Test_MsgPack_Ext_Cur_Core "D5 06 01 00", CCur("0.0256")
    Test_MsgPack_Ext_Cur_Core "D5 06 FF FF", CCur("6.5535")
    
    Test_MsgPack_Ext_Cur_Core "D5 06 27 10", CCur("1")
End Sub

Public Sub Test_MsgPack_Ext_Cur_FixExt4_TestCases()
    Debug.Print "Target: Currency FixExt4"
    
    Test_MsgPack_Ext_Cur_Core "D6 06 00 01 00 00", CCur("6.5536")
    Test_MsgPack_Ext_Cur_Core "D6 06 FF FF FF FF", CCur("429496.7295")
End Sub

Public Sub Test_MsgPack_Ext_Cur_FixExt8_TestCases()
    Debug.Print "Target: Currency FixExt8"
    
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 00 00 00 01 00 00 00 00", CCur("429496.7296")
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 7F FF FF FF FF FF FF FF", CCur("922337203685477.5807")
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 80 00 00 00 00 00 00 00", CCur("-922337203685477.5808")
    Test_MsgPack_Ext_Cur_Core _
        "D7 06 FF FF FF FF FF FF FF FF", CCur("-0.0001")
    
    Test_MsgPack_Ext_Cur_Core "D7 06 FF FF FF FF FF FF D8 F0", CCur("-1")
End Sub

'
' MessagePack for VBA - Extension - Currency - Test Core
'

Public Sub Test_MsgPack_Ext_Cur_Core( _
    HexBE As String, ExpectedValue As Currency)
    
    Dim BytesBE() As Byte
    BytesBE = BitConverter.GetBytesFromHexString(HexBE)
    
    Dim OutputValue As Currency
    OutputValue = MsgPack_Ext_Cur.GetExtCurFromBytes(BytesBE)
    
    DebugPrint_MsgPack_Ext_Cur_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputBytesBE() As Byte
    OutputBytesBE = MsgPack_Ext_Cur.GetBytesFromExtCur(OutputValue)
    
    DebugPrint_MsgPack_Ext_Cur_GetBytes OutputValue, OutputBytesBE, BytesBE
End Sub

'
' MessagePack for VBA - Extension - Currency - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Ext_Cur_GetBytes( _
    Value As Currency, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        CStr(Value), OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Ext_Cur_GetValue( _
    Bytes() As Byte, OutputValue As Currency, ExpectedValue As Currency)
    
    BitConverter.DebugPrint_GetValue Bytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

Attribute VB_Name = "Test_MsgPack_Map"
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
' MessagePack for VBA - Map - Test
'

Public Sub Test_MsgPack_Map()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Map_FixMap_TestCases
    Test_MsgPack_Map_Map16_TestCases
    Test_MsgPack_Map_Map32_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Map_FixMap()
    BitConverter.Test_Initialize
    Test_MsgPack_Map_FixMap_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Map_Map16()
    BitConverter.Test_Initialize
    Test_MsgPack_Map_Map16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Map_Map32()
    BitConverter.Test_Initialize
    Test_MsgPack_Map_Map32_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Map - Test Cases
'

Public Sub Test_MsgPack_Map_FixMap_TestCases()
    Debug.Print "Target: FixMap"
    
    Test_MsgPack_Map_Core "80"
    Test_MsgPack_Map_Core "81 A1 61 00"
    Test_MsgPack_Map_Core2 "8F", &HF
End Sub

Public Sub Test_MsgPack_Map_Map16_TestCases()
    Debug.Print "Target: Map16"
    
    Test_MsgPack_Map_Core2 "DE 00 10", &H10
    Test_MsgPack_Map_Core2 "DE 01 00", &H100
    'Test_MsgPack_Map_Core2 "DE FF FF", &HFFFF&
End Sub

Public Sub Test_MsgPack_Map_Map32_TestCases()
    Debug.Print "Target: Map32"
    
    'Test_MsgPack_Map_Core2 "DF 00 01 00 00", &H10000
End Sub

'
' MessagePack for VBA - Map - Test Core
'

Public Sub Test_MsgPack_Map_Core(HexStr As String)
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytesFromHexString(HexStr)
    
    Dim ExpectedDummy As Object
    Set ExpectedDummy = CreateObject("Scripting.Dictionary")
    
    Dim OutputValue As Object
    Set OutputValue = MsgPack_Map.GetMapFromBytes(Bytes)
    
    DebugPrint_MsgPack_Map_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Map.GetBytesFromMap(OutputValue)
    
    DebugPrint_MsgPack_Map_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Public Sub Test_MsgPack_Map_Core2(HeadHex As String, ElementCount As Long)
    Dim HeadBytes() As Byte
    HeadBytes = BitConverter.GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestMapBytes(HeadBytes, ElementCount)
    
    Dim ExpectedDummy As Object
    Set ExpectedDummy = CreateObject("Scripting.Dictionary")
    
    Dim OutputValue As Object
    Set OutputValue = MsgPack_Map.GetMapFromBytes(Bytes)
    
    DebugPrint_MsgPack_Map_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Map.GetBytesFromMap(OutputValue)
    
    DebugPrint_MsgPack_Map_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Private Function GetTestMapBytes( _
    HeadBytes() As Byte, ElementCount As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(0 To HeadLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To ElementCount
        MsgPack_Common.AddBytes TestBytes, _
            MsgPack_Str.GetBytesFromStr("key-" & CStr(Index))
        MsgPack_Common.AddBytes TestBytes, _
            MsgPack_Str.GetBytesFromStr("value-" & CStr(Index))
    Next
    
    GetTestMapBytes = TestBytes
End Function

'
' MessagePack for VBA - Map - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Map_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        "(" & TypeName(Value) & ")", OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Map_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    Dim OutputDummy As String
    OutputDummy = "(" & TypeName(OutputValue) & ")"
    
    Dim ExpectedDummy As String
    ExpectedDummy = "(" & TypeName(ExpectedValue) & ")"
    
    BitConverter.DebugPrint_GetValue Bytes, _
        OutputDummy, ExpectedDummy, _
        OutputDummy, ExpectedDummy
End Sub

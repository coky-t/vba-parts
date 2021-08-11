Attribute VB_Name = "Test_MsgPack_Array_Collection"
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
' MessagePack for VBA - Array(Collection) - Test
'

Public Sub Test_MsgPack_Array()
    BitConverter.Test_Initialize
    
    Test_MsgPack_Array_FixArray_TestCases
    Test_MsgPack_Array_Array16_TestCases
    Test_MsgPack_Array_Array32_TestCases
    
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Array_FixArray()
    BitConverter.Test_Initialize
    Test_MsgPack_Array_FixArray_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Array_Array16()
    BitConverter.Test_Initialize
    Test_MsgPack_Array_Array16_TestCases
    BitConverter.Test_Terminate
End Sub

Public Sub Test_MsgPack_Array_Array32()
    BitConverter.Test_Initialize
    Test_MsgPack_Array_Array32_TestCases
    BitConverter.Test_Terminate
End Sub

'
' MessagePack for VBA - Array(Collection) - Test Cases
'

Public Sub Test_MsgPack_Array_FixArray_TestCases()
    Debug.Print "Target: FixArray"
    
    Test_MsgPack_Array_Core "90"
    Test_MsgPack_Array_Core "91 A1 61"
    Test_MsgPack_Array_Core2 "9F", &HF
End Sub

Public Sub Test_MsgPack_Array_Array16_TestCases()
    Debug.Print "Target: Array16"
    
    Test_MsgPack_Array_Core2 "DC 00 10", &H10
    Test_MsgPack_Array_Core2 "DC 01 00", &H100
    'Test_MsgPack_Array_Core2 "DC FF FF", &HFFFF&
End Sub

Public Sub Test_MsgPack_Array_Array32_TestCases()
    Debug.Print "Target: Array32"
    
    'Test_MsgPack_Array_Core2 "DD 00 01 00 00", &H10000
End Sub

'
' MessagePack for VBA - Array(Collection) - Test Core
'

Public Sub Test_MsgPack_Array_Core(HexStr As String)
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytesFromHexString(HexStr)
    
    Dim ExpectedDummy As Collection
    Set ExpectedDummy = New Collection
    
    Dim OutputValue As Collection
    Set OutputValue = MsgPack_Array_Collection.GetArrayFromBytes(Bytes)
    
    DebugPrint_MsgPack_Array_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Array_Collection.GetBytesFromArray(OutputValue)
    
    DebugPrint_MsgPack_Array_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Public Sub Test_MsgPack_Array_Core2(HeadHex As String, ElementCount As Long)
    Dim HeadBytes() As Byte
    HeadBytes = BitConverter.GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestArrayBytes(HeadBytes, ElementCount)
    
    Dim ExpectedDummy As Collection
    Set ExpectedDummy = New Collection
    
    Dim OutputValue As Collection
    Set OutputValue = MsgPack_Array_Collection.GetArrayFromBytes(Bytes)
    
    DebugPrint_MsgPack_Array_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputBytes() As Byte
    OutputBytes = MsgPack_Array_Collection.GetBytesFromArray(OutputValue)
    
    DebugPrint_MsgPack_Array_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Private Function GetTestArrayBytes( _
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
            MsgPack_Str.GetBytesFromStr("value-" & CStr(Index))
    Next
    
    GetTestArrayBytes = TestBytes
End Function

'
' MessagePack for VBA - Array(Collection) - Test - Debug.Print
'

Private Sub DebugPrint_MsgPack_Array_GetBytes( _
    Value, OutputBytes() As Byte, ExpectedBytes() As Byte)
    
    BitConverter.DebugPrint_GetBytes _
        "(" & TypeName(Value) & ")", OutputBytes, ExpectedBytes
End Sub

Private Sub DebugPrint_MsgPack_Array_GetValue( _
    Bytes() As Byte, OutputValue, ExpectedValue)
    
    Dim OutputDummy As String
    OutputDummy = "(" & TypeName(OutputValue) & ")"
    
    Dim ExpectedDummy As String
    ExpectedDummy = "(" & TypeName(ExpectedValue) & ")"
    
    BitConverter.DebugPrint_GetValue Bytes, _
        OutputDummy, ExpectedDummy, _
        OutputDummy, ExpectedDummy
End Sub

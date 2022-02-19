Attribute VB_Name = "Test_CBOR_7_Simple"
Option Explicit

'
' Copyright (c) 2022 Koki Takeyama
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

''
'' CBOR for VBA - Test
''

' Test Counter
Private m_Test_Count As Long
Private m_Test_Success As Long
Private m_Test_Fail As Long

#If Win64 Then
#Const USE_LONGLONG = True
#End If

Public Sub Test_Cbor()
    Test_Initialize
    
    Test_Cbor_False_TestCases
    Test_Cbor_True_TestCases
    Test_Cbor_Null_TestCases
    Test_Cbor_Undefined_TestCases
    
    Test_Terminate
End Sub

'
' CBOR for VBA - Test Cases
'

Private Sub Test_Cbor_False_TestCases()
    Debug.Print "Target: False"
    
    Test_Cbor_Bool_Core "F4", False
End Sub

Private Sub Test_Cbor_True_TestCases()
    Debug.Print "Target: True"
    
    Test_Cbor_Bool_Core "F5", True
End Sub

Private Sub Test_Cbor_Null_TestCases()
    Debug.Print "Target: Null"
    Test_Cbor_Null_Core "F6", Null
End Sub

Private Sub Test_Cbor_Undefined_TestCases()
    Debug.Print "Target: Undefined"
    Test_Cbor_Undefined_Core "F7", Empty
End Sub

'
' CBOR for VBA - Test Core
'

Private Sub Test_Cbor_Bool_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR_7_Simple.GetValue(BytesBE)
    
    DebugPrint_Bool_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR_7_Simple.GetCborBytes(OutputValue)
    
    DebugPrint_Bool_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

Private Sub Test_Cbor_Null_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR_7_Simple.GetValue(BytesBE)
    
    DebugPrint_Null_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR_7_Simple.GetCborBytes(OutputValue)
    
    DebugPrint_Null_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

Private Sub Test_Cbor_Undefined_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR_7_Simple.GetValue(BytesBE)
    
    DebugPrint_Undefined_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR_7_Simple.GetCborBytes(OutputValue)
    
    DebugPrint_Undefined_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

'
' CBOR for VBA - Test - Debug.Print - Boolean
'

Private Sub DebugPrint_Bool_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes CStr(Value), OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Bool_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

'
' CBOR for VBA - Test - Debug.Print - Null
'

Private Sub DebugPrint_Null_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes _
        IIf(IsNull(Value), "Null", "not Null"), OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Null_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, _
        IIf(IsNull(OutputValue), "Null", "not Null"), _
        IIf(IsNull(ExpectedValue), "Null", "not Null"), _
        IIf(IsNull(OutputValue), "Null", "not Null"), _
        IIf(IsNull(ExpectedValue), "Null", "not Null")
End Sub

'
' CBOR for VBA - Test - Debug.Print - Undefined
'

Private Sub DebugPrint_Undefined_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes _
        IIf(IsEmpty(Value), "Empty", "not Empty"), _
        OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Undefined_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, _
        IIf(IsEmpty(OutputValue), "Empty", "not Empty"), _
        IIf(IsEmpty(ExpectedValue), "Empty", "not Empty"), _
        IIf(IsEmpty(OutputValue), "Empty", "not Empty"), _
        IIf(IsEmpty(ExpectedValue), "Empty", "not Empty")
End Sub

''
'' CBOR for VBA - Test Counter
''

Private Property Get Test_Count() As Long
    Test_Count = m_Test_Count
End Property

Private Sub Test_Initialize()
    m_Test_Count = 0
    m_Test_Success = 0
    m_Test_Fail = 0
End Sub

Private Sub Test_Countup(bSuccess As Boolean)
    m_Test_Count = m_Test_Count + 1
    If bSuccess Then
        m_Test_Success = m_Test_Success + 1
    Else
        m_Test_Fail = m_Test_Fail + 1
    End If
End Sub

Private Sub Test_Terminate()
    Debug.Print _
        "Count: " & CStr(m_Test_Count) & ", " & _
        "Success: " & CStr(m_Test_Success) & ", " & _
        "Fail: " & CStr(m_Test_Fail)
End Sub

''
'' CBOR for VBA - Test - Debug.Print
''

Private Sub DebugPrint_GetCborBytes( _
    Source, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    Dim bSuccess As Boolean
    bSuccess = CompareBytes(OutputCBBytes, ExpectedCBBytes)
    
    Test_Countup bSuccess
    
    Dim OutputCBBytesStr As String
    OutputCBBytesStr = GetHexStringFromBytes(OutputCBBytes, , , " ")
    
    Dim ExpectedCBBytesStr As String
    ExpectedCBBytesStr = GetHexStringFromBytes(ExpectedCBBytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & Source & _
        " Output: " & OutputCBBytesStr & _
        " Expect: " & ExpectedCBBytesStr
End Sub

Private Sub DebugPrint_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue, Output, Expect)
    
    Dim bSuccess As Boolean
    bSuccess = (OutputValue = ExpectedValue)
    
    Test_Countup bSuccess
    
    Dim CBBytesStr As String
    CBBytesStr = GetHexStringFromBytes(CBBytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & CBBytesStr & _
        " Output: " & Output & _
        " Expect: " & Expect
End Sub

''
'' CBOR for VBA - Test - Byte Array Helper
''

Private Function CompareBytes(Bytes1() As Byte, Bytes2() As Byte) As Boolean
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(Bytes1)
    UB1 = UBound(Bytes1)
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(Bytes2)
    UB2 = UBound(Bytes2)
    
    If (UB1 - LB1 + 1) <> (UB2 - LB2 + 1) Then Exit Function
    
    Dim Index As Long
    For Index = 0 To UB1 - LB1
        If Bytes1(LB1 + Index) <> Bytes2(LB2 + Index) Then Exit Function
    Next
    
    CompareBytes = True
End Function

''
'' CBOR for VBA - Test - Hex String
''

Private Function GetBytesFromHexString(ByVal Value As String) As Byte()
    Dim Value_ As String
    Dim Index As Long
    For Index = 1 To Len(Value)
        Select Case Mid(Value, Index, 1)
        Case "0" To "9", "A" To "F", "a" To "f"
            Value_ = Value_ & Mid(Value, Index, 1)
        End Select
    Next
    
    Dim Length As Long
    Length = Len(Value_) \ 2
    
    Dim Bytes() As Byte
    
    If Length = 0 Then
        GetBytesFromHexString = Bytes
        Exit Function
    End If
    
    ReDim Bytes(0 To Length - 1)
    
    'Dim Index As Long
    For Index = 0 To Length - 1
        Bytes(Index) = CByte("&H" & Mid(Value_, 1 + Index * 2, 2))
    Next
    
    GetBytesFromHexString = Bytes
End Function

'Private Function GetHexStringFromBytes(Bytes() As Byte,
Private Function GetHexStringFromBytes(Bytes, _
    Optional Index As Long, Optional Length As Long, _
    Optional Separator As String) As String
    
    If Length = 0 Then
        On Error Resume Next
        Length = UBound(Bytes) - Index + 1
        On Error GoTo 0
    End If
    If Length = 0 Then
        GetHexStringFromBytes = ""
        Exit Function
    End If
    
    Dim HexString As String
    HexString = Right("0" & Hex(Bytes(Index)), 2)
    
    Dim Offset As Long
    For Offset = 1 To Length - 1
        HexString = _
            HexString & Separator & Right("0" & Hex(Bytes(Index + Offset)), 2)
    Next
    
    GetHexStringFromBytes = HexString
End Function

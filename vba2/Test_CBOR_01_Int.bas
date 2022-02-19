Attribute VB_Name = "Test_CBOR_01_Int"
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
    
    Test_Cbor_PosFixInt_TestCases
    Test_Cbor_PosInt8_TestCases
    Test_Cbor_PosInt16_TestCases
    Test_Cbor_PosInt32_TestCases
    Test_Cbor_PosInt64_TestCases
    
    Test_Cbor_NegFixInt_TestCases
    Test_Cbor_NegInt8_TestCases
    Test_Cbor_NegInt16_TestCases
    Test_Cbor_NegInt32_TestCases
    Test_Cbor_NegInt64_TestCases
    
    Test_Terminate
End Sub

'
' CBOR for VBA - Test Cases
'

Private Sub Test_Cbor_PosFixInt_TestCases()
    Debug.Print "Target: PosFixInt"
    
    Test_Cbor_Int_Core "00", &H0
    Test_Cbor_Int_Core "17", &H17
End Sub


Private Sub Test_Cbor_PosInt8_TestCases()
    Debug.Print "Target: PosInt8"
    
    Test_Cbor_Int_Core "18 18", &H18
    Test_Cbor_Int_Core "18 7F", &H7F
    Test_Cbor_Int_Core "18 80", &H80
    Test_Cbor_Int_Core "18 FF", &HFF
End Sub

Private Sub Test_Cbor_PosInt16_TestCases()
    Debug.Print "Target: PosInt16"
    
    Test_Cbor_Int_Core "19 01 00", &H100
    Test_Cbor_Int_Core "19 7F FF", &H7FFF
    Test_Cbor_Int_Core "19 80 00", &H8000&
    Test_Cbor_Int_Core "19 FF FF", &HFFFF&
End Sub

Private Sub Test_Cbor_PosInt32_TestCases()
    Debug.Print "Target: PosInt32"
    
    Test_Cbor_Int_Core "1A 00 01 00 00", &H10000
    Test_Cbor_Int_Core "1A 7F FF FF FF", &H7FFFFFFF
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "1A 80 00 00 00", &H80000000^
    Test_Cbor_Int_Core "1A FF FF FF FF", &HFFFFFFFF^
#Else
    Test_Cbor_Int_Core "1A 80 00 00 00", CDec("&H80000000")
    Test_Cbor_Int_Core "1A FF FF FF FF", CDec("&HFFFFFFFF")
#End If
End Sub

Private Sub Test_Cbor_PosInt64_TestCases()
    Debug.Print "Target: PosInt64"
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "1B 00 00 00 01 00 00 00 00", _
        CLngLng("4294967296")
    Test_Cbor_Int_Core "1B 7F FF FF FF FF FF FF FF", _
        CLngLng("9223372036854775807")
#Else
    Test_Cbor_Int_Core "1B 00 00 00 01 00 00 00 00", _
        CDec("4294967296")
    Test_Cbor_Int_Core "1B 7F FF FF FF FF FF FF FF", _
        CDec("9223372036854775807")
#End If
    Test_Cbor_Int_Core "1B 80 00 00 00 00 00 00 00", _
        CDec("9223372036854775808")
    Test_Cbor_Int_Core "1B FF FF FF FF FF FF FF FF", _
        CDec("18446744073709551615")
End Sub

Private Sub Test_Cbor_NegFixInt_TestCases()
    Debug.Print "Target: NegFixInt"
    
    Test_Cbor_Int_Core "20", -1
    Test_Cbor_Int_Core "37", -24
End Sub

Private Sub Test_Cbor_NegInt8_TestCases()
    Debug.Print "Target: NegInt8"
    
    Test_Cbor_Int_Core "38 18", -25
    Test_Cbor_Int_Core "38 7F", -128
    Test_Cbor_Int_Core "38 80", -129
    Test_Cbor_Int_Core "38 FF", -256
End Sub

Private Sub Test_Cbor_NegInt16_TestCases()
    Debug.Print "Target: NegInt16"
    
    Test_Cbor_Int_Core "39 01 00", -257
    Test_Cbor_Int_Core "39 7F FF", CInt(-32768)
    Test_Cbor_Int_Core "39 80 00", CLng(-32769)
    Test_Cbor_Int_Core "39 FF FF", -65536
End Sub

Private Sub Test_Cbor_NegInt32_TestCases()
    Debug.Print "Target: NegInt32"
    
    Test_Cbor_Int_Core "3A 00 01 00 00", -65537
    Test_Cbor_Int_Core "3A 7F FF FF FF", CLng("-2147483648")
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "3A 80 00 00 00", CLngLng("-2147483649")
    Test_Cbor_Int_Core "3A FF FF FF FF", CLngLng("-4294967296")
#Else
    Test_Cbor_Int_Core "3A 80 00 00 00", CDec("-2147483649")
    Test_Cbor_Int_Core "3A FF FF FF FF", CDec("-4294967296")
#End If
End Sub

Private Sub Test_Cbor_NegInt64_TestCases()
    Debug.Print "Target: NegInt64"
    
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "3B 00 00 00 01 00 00 00 00", _
        CLngLng("-4294967297")
    Test_Cbor_Int_Core "3B 7F FF FF FF FF FF FF FF", _
        CLngLng("-9223372036854775808")
#Else
    Test_Cbor_Int_Core "3B 00 00 00 01 00 00 00 00", _
        CDec("-4294967297")
    Test_Cbor_Int_Core "3B 7F FF FF FF FF FF FF FF", _
        CDec("-9223372036854775808")
#End If
    Test_Cbor_Int_Core "3B 80 00 00 00 00 00 00 00", _
        CDec("-9223372036854775809")
    Test_Cbor_Int_Core "3B FF FF FF FF FF FF FF FF", _
        CDec("-18446744073709551616")
End Sub

'
' CBOR for VBA - Test Core
'

Private Sub Test_Cbor_Int_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR_01_Int.GetValue(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR_01_Int.GetCborBytes(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

'
' CBOR for VBA - Test - Debug.Print - Integer
'

Private Sub DebugPrint_Int_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    If VarType(Value) = vbDecimal Then
        DebugPrint_GetCborBytes CStr(Value), OutputCBBytes, ExpectedCBBytes
    Else
        DebugPrint_GetCborBytes _
            CStr(Value) & " (" & Hex(Value) & ")", _
            OutputCBBytes, ExpectedCBBytes
    End If
End Sub

Private Sub DebugPrint_Int_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    If (VarType(OutputValue) = vbDecimal) Or _
        (VarType(ExpectedValue) = vbDecimal) Then
        
        DebugPrint_GetValue CBBytes, OutputValue, ExpectedValue, _
            CStr(OutputValue), CStr(ExpectedValue)
    Else
        DebugPrint_GetValue CBBytes, OutputValue, ExpectedValue, _
            CStr(OutputValue) & " (" & Hex(OutputValue) & ")", _
            CStr(ExpectedValue) & " (" & Hex(ExpectedValue) & ")"
    End If
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

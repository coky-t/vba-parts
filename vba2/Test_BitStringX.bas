Attribute VB_Name = "Test_BitStringX"
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
' --- Test ---
'

Public Sub Test_GetBinStringFromByte()
    Dim Index As Long
    For Index = 0 To 7
        Test_GetBinStringFromByte_Core 2 ^ Index
    Next
    Test_GetBinStringFromByte_Core &HF0
    Test_GetBinStringFromByte_Core &HFF
End Sub

Public Sub Test_GetBinStringFromInteger()
    Dim Index As Long
    For Index = 0 To 14
        Test_GetBinStringFromInteger_Core 2 ^ Index
    Next
    For Index = 0 To 14
        Test_GetBinStringFromInteger_Core -(2 ^ Index)
    Next
    Test_GetBinStringFromInteger_Core (-(2 ^ 14)) * 2
End Sub

Public Sub Test_GetBinStringFromLong()
    Dim Index As Long
    For Index = 0 To 30
        Test_GetBinStringFromLong_Core 2 ^ Index
    Next
    For Index = 0 To 30
        Test_GetBinStringFromLong_Core -(2 ^ Index)
    Next
    Test_GetBinStringFromLong_Core (-(2 ^ 30)) * 2
End Sub

#If Win64 Then
Public Sub Test_GetBinStringFromLongLong()
    Dim Index As Long
    For Index = 0 To 62
        Test_GetBinStringFromLongLong_Core 2 ^ Index
    Next
    For Index = 0 To 62
        Test_GetBinStringFromLongLong_Core -(2 ^ Index)
    Next
    Test_GetBinStringFromLongLong_Core (-(2 ^ 62)) * 2
End Sub
#End If

Public Sub Test_GetBinStringFromSingle()
    Test_GetBinStringFromSingle_Core 0!
    Test_GetBinStringFromSingle_Core 1.401298E-45
    Test_GetBinStringFromSingle_Core 1.175494E-38
    Test_GetBinStringFromSingle_Core 3.402823E+38
    
    Test_GetBinStringFromSingle_Core -0!
    Test_GetBinStringFromSingle_Core -1.401298E-45
    Test_GetBinStringFromSingle_Core -1.175494E-38
    Test_GetBinStringFromSingle_Core -3.402823E+38
    
    Test_GetBinStringFromSingle_Core 1!
    Test_GetBinStringFromSingle_Core -1!
    Test_GetBinStringFromSingle_Core 0.5
    Test_GetBinStringFromSingle_Core 0.1
    Test_GetBinStringFromSingle_Core 1! / 3!
End Sub

Public Sub Test_GetBinStringFromDouble()
    Test_GetBinStringFromDouble_Core 0#
    Test_GetBinStringFromDouble_Core 4.94065645841247E-324
    Test_GetBinStringFromDouble_Core 2.2250738585072E-308
    Test_GetBinStringFromDouble_Core 1.7976931348623E+308
    
    Test_GetBinStringFromDouble_Core -0#
    Test_GetBinStringFromDouble_Core -4.94065645841247E-324
    Test_GetBinStringFromDouble_Core -2.2250738585072E-308
    Test_GetBinStringFromDouble_Core -1.7976931348623E+308
    
    Test_GetBinStringFromDouble_Core 1#
    Test_GetBinStringFromDouble_Core -1#
    Test_GetBinStringFromDouble_Core 0.5
    Test_GetBinStringFromDouble_Core 0.1
    Test_GetBinStringFromDouble_Core 1# / 3#
End Sub

Public Sub Test_GetSingleFromBinString()
    ' Positive Zero
    Test_GetSingleFromBinString_Core "00000000000000000000000000000000"
    
    ' Positive SubNormal Minimum
    Test_GetSingleFromBinString_Core "00000000000000000000000000000001"
    
    ' Positive SubNormal Maximum
    Test_GetSingleFromBinString_Core "00000000011111111111111111111111"
    
    ' Positive Normal Minimum
    Test_GetSingleFromBinString_Core "00000000100000000000000000000000"
    
    ' Positive Normal Maximum
    Test_GetSingleFromBinString_Core "01111111011111111111111111111111"
    
    ' Positive Infinity
    Test_GetSingleFromBinString_Core "01111111100000000000000000000000"
    
    ' Positive NaN
    Test_GetSingleFromBinString_Core "01111111111111111111111111111111"
    
    ' Negative Zero
    Test_GetSingleFromBinString_Core "10000000000000000000000000000000"
    
    ' Negative SubNormal Minimum
    Test_GetSingleFromBinString_Core "10000000000000000000000000000001"
    
    ' Negative SubNormal Maximum
    Test_GetSingleFromBinString_Core "10000000011111111111111111111111"
    
    ' Negative Normal Minimum
    Test_GetSingleFromBinString_Core "10000000100000000000000000000000"
    
    ' Negative Normal Maximum
    Test_GetSingleFromBinString_Core "11111111011111111111111111111111"
    
    ' Negative Infinity
    Test_GetSingleFromBinString_Core "11111111100000000000000000000000"
    
    ' Negative NaN
    Test_GetSingleFromBinString_Core "11111111111111111111111111111111"
End Sub

Public Sub Test_GetDoubleFromBinString()
    ' Positive Zero
    Test_GetDoubleFromBinString_Core _
        "0000000000000000000000000000000000000000000000000000000000000000"
    
    ' Positive SubNormal Minimum
    Test_GetDoubleFromBinString_Core _
        "0000000000000000000000000000000000000000000000000000000000000001"
    
    ' Positive SubNormal Maximum
    Test_GetDoubleFromBinString_Core _
        "0000000000001111111111111111111111111111111111111111111111111111"
    
    ' Positive Normal Minimum
    Test_GetDoubleFromBinString_Core _
        "0000000000010000000000000000000000000000000000000000000000000000"
    
    ' Positive Normal Maximum
    Test_GetDoubleFromBinString_Core _
        "0111111111101111111111111111111111111111111111111111111111111111"
    
    ' Positive Infinity
    Test_GetDoubleFromBinString_Core _
        "0111111111110000000000000000000000000000000000000000000000000000"
    
    ' Positive NaN
    Test_GetDoubleFromBinString_Core _
        "0111111111111111111111111111111111111111111111111111111111111111"
    
    ' Negative Zero
    Test_GetDoubleFromBinString_Core _
        "1000000000000000000000000000000000000000000000000000000000000000"
    
    ' Negative SubNormal Minimum
    Test_GetDoubleFromBinString_Core _
        "1000000000000000000000000000000000000000000000000000000000000001"
    
    ' Negative SubNormal Maximum
    Test_GetDoubleFromBinString_Core _
        "1000000000001111111111111111111111111111111111111111111111111111"
    
    ' Negative Normal Minimum
    Test_GetDoubleFromBinString_Core _
        "1000000000010000000000000000000000000000000000000000000000000000"
    
    ' Negative Normal Maximum
    Test_GetDoubleFromBinString_Core _
        "1111111111101111111111111111111111111111111111111111111111111111"
    
    ' Negative Infinity
    Test_GetDoubleFromBinString_Core _
        "1111111111110000000000000000000000000000000000000000000000000000"
    
    ' Negative NaN
    Test_GetDoubleFromBinString_Core _
        "1111111111111111111111111111111111111111111111111111111111111111"
End Sub

Public Sub Test_GetOctStringFromByte()
    Dim Index As Long
    For Index = 0 To 7
        Test_GetOctStringFromByte_Core 2 ^ Index
    Next
    Test_GetOctStringFromByte_Core &HF0
    Test_GetOctStringFromByte_Core &HFF
End Sub

Public Sub Test_GetOctStringFromInteger()
    Dim Index As Long
    For Index = 0 To 14
        Test_GetOctStringFromInteger_Core 2 ^ Index
    Next
    For Index = 0 To 14
        Test_GetOctStringFromInteger_Core -(2 ^ Index)
    Next
    Test_GetOctStringFromInteger_Core (-(2 ^ 14)) * 2
End Sub

Public Sub Test_GetOctStringFromLong()
    Dim Index As Long
    For Index = 0 To 30
        Test_GetOctStringFromLong_Core 2 ^ Index
    Next
    For Index = 0 To 30
        Test_GetOctStringFromLong_Core -(2 ^ Index)
    Next
    Test_GetOctStringFromLong_Core (-(2 ^ 30)) * 2
End Sub

#If Win64 Then
Public Sub Test_GetOctStringFromLongLong()
    Dim Index As Long
    For Index = 0 To 62
        Test_GetOctStringFromLongLong_Core 2 ^ Index
    Next
    For Index = 0 To 62
        Test_GetOctStringFromLongLong_Core -(2 ^ Index)
    Next
    Test_GetOctStringFromLongLong_Core (-(2 ^ 62)) * 2
End Sub
#End If

Public Sub Test_GetOctStringFromSingle()
    Test_GetOctStringFromSingle_Core 0!
    Test_GetOctStringFromSingle_Core 1.401298E-45
    Test_GetOctStringFromSingle_Core 1.175494E-38
    Test_GetOctStringFromSingle_Core 3.402823E+38
    
    Test_GetOctStringFromSingle_Core -0!
    Test_GetOctStringFromSingle_Core -1.401298E-45
    Test_GetOctStringFromSingle_Core -1.175494E-38
    Test_GetOctStringFromSingle_Core -3.402823E+38
    
    Test_GetOctStringFromSingle_Core 1!
    Test_GetOctStringFromSingle_Core -1!
    Test_GetOctStringFromSingle_Core 0.5
    Test_GetOctStringFromSingle_Core 0.1
    Test_GetOctStringFromSingle_Core 1! / 3!
End Sub

Public Sub Test_GetOctStringFromDouble()
    Test_GetOctStringFromDouble_Core 0#
    Test_GetOctStringFromDouble_Core 4.94065645841247E-324
    Test_GetOctStringFromDouble_Core 2.2250738585072E-308
    Test_GetOctStringFromDouble_Core 1.7976931348623E+308
    
    Test_GetOctStringFromDouble_Core -0#
    Test_GetOctStringFromDouble_Core -4.94065645841247E-324
    Test_GetOctStringFromDouble_Core -2.2250738585072E-308
    Test_GetOctStringFromDouble_Core -1.7976931348623E+308
    
    Test_GetOctStringFromDouble_Core 1#
    Test_GetOctStringFromDouble_Core -1#
    Test_GetOctStringFromDouble_Core 0.5
    Test_GetOctStringFromDouble_Core 0.1
    Test_GetOctStringFromDouble_Core 1# / 3#
End Sub

Public Sub Test_GetSingleFromOctString()
    ' Positive Zero
    Test_GetSingleFromOctString_Core "00000000000"
    
    ' Positive SubNormal Minimum
    Test_GetSingleFromOctString_Core "00000000001"
    
    ' Positive SubNormal Maximum
    Test_GetSingleFromOctString_Core "00037777777"
    
    ' Positive Normal Minimum
    Test_GetSingleFromOctString_Core "00040000000"
    
    ' Positive Normal Maximum
    Test_GetSingleFromOctString_Core "17737777777"
    
    ' Positive Infinity
    Test_GetSingleFromOctString_Core "17740000000"
    
    ' Positive NaN
    Test_GetSingleFromOctString_Core "17777777777"
    
    ' Negative Zero
    Test_GetSingleFromOctString_Core "20000000000"
    
    ' Negative SubNormal Minimum
    Test_GetSingleFromOctString_Core "20000000001"
    
    ' Negative SubNormal Maximum
    Test_GetSingleFromOctString_Core "20037777777"
    
    ' Negative Normal Minimum
    Test_GetSingleFromOctString_Core "20040000000"
    
    ' Negative Normal Maximum
    Test_GetSingleFromOctString_Core "37737777777"
    
    ' Negative Infinity
    Test_GetSingleFromOctString_Core "37740000000"
    
    ' Negative NaN
    Test_GetSingleFromOctString_Core "37777777777"
End Sub

Public Sub Test_GetDoubleFromOctString()
    ' Positive Zero
    Test_GetDoubleFromOctString_Core "0000000000000000000000"
    
    ' Positive SubNormal Minimum
    Test_GetDoubleFromOctString_Core "0000000000000000000001"
    
    ' Positive SubNormal Maximum
    Test_GetDoubleFromOctString_Core "0000177777777777777777"
    
    ' Positive Normal Minimum
    Test_GetDoubleFromOctString_Core "0000200000000000000000"
    
    ' Positive Normal Maximum
    Test_GetDoubleFromOctString_Core "0777577777777777777777"
    
    ' Positive Infinity
    Test_GetDoubleFromOctString_Core "0777600000000000000000"
    
    ' Positive NaN
    Test_GetDoubleFromOctString_Core "0777777777777777777777"
    
    ' Negative Zero
    Test_GetDoubleFromOctString_Core "1000000000000000000000"
    
    ' Negative SubNormal Minimum
    Test_GetDoubleFromOctString_Core "1000000000000000000001"
    
    ' Negative SubNormal Maximum
    Test_GetDoubleFromOctString_Core "1000177777777777777777"
    
    ' Negative Normal Minimum
    Test_GetDoubleFromOctString_Core "1000200000000000000000"
    
    ' Negative Normal Maximum
    Test_GetDoubleFromOctString_Core "1777577777777777777777"
    
    ' Negative Infinity
    Test_GetDoubleFromOctString_Core "1777600000000000000000"
    
    ' Negative NaN
    Test_GetDoubleFromOctString_Core "1777777777777777777777"
End Sub

Public Sub Test_GetHexStringFromByte()
    Dim Index As Long
    For Index = 0 To 7
        Test_GetHexStringFromByte_Core 2 ^ Index
    Next
    Test_GetHexStringFromByte_Core &HF0
    Test_GetHexStringFromByte_Core &HFF
End Sub

Public Sub Test_GetHexStringFromInteger()
    Dim Index As Long
    For Index = 0 To 14
        Test_GetHexStringFromInteger_Core 2 ^ Index
    Next
    For Index = 0 To 14
        Test_GetHexStringFromInteger_Core -(2 ^ Index)
    Next
    Test_GetHexStringFromInteger_Core (-(2 ^ 14)) * 2
End Sub

Public Sub Test_GetHexStringFromLong()
    Dim Index As Long
    For Index = 0 To 30
        Test_GetHexStringFromLong_Core 2 ^ Index
    Next
    For Index = 0 To 30
        Test_GetHexStringFromLong_Core -(2 ^ Index)
    Next
    Test_GetHexStringFromLong_Core (-(2 ^ 30)) * 2
End Sub

#If Win64 Then
Public Sub Test_GetHexStringFromLongLong()
    Dim Index As Long
    For Index = 0 To 62
        Test_GetHexStringFromLongLong_Core 2 ^ Index
    Next
    For Index = 0 To 62
        Test_GetHexStringFromLongLong_Core -(2 ^ Index)
    Next
    Test_GetHexStringFromLongLong_Core (-(2 ^ 62)) * 2
End Sub
#End If

Public Sub Test_GetHexStringFromSingle()
    Test_GetHexStringFromSingle_Core 0!
    Test_GetHexStringFromSingle_Core 1.401298E-45
    Test_GetHexStringFromSingle_Core 1.175494E-38
    Test_GetHexStringFromSingle_Core 3.402823E+38
    
    Test_GetHexStringFromSingle_Core -0!
    Test_GetHexStringFromSingle_Core -1.401298E-45
    Test_GetHexStringFromSingle_Core -1.175494E-38
    Test_GetHexStringFromSingle_Core -3.402823E+38
    
    Test_GetHexStringFromSingle_Core 1!
    Test_GetHexStringFromSingle_Core -1!
    Test_GetHexStringFromSingle_Core 0.5
    Test_GetHexStringFromSingle_Core 0.1
    Test_GetHexStringFromSingle_Core 1! / 3!
End Sub

Public Sub Test_GetHexStringFromDouble()
    Test_GetHexStringFromDouble_Core 0#
    Test_GetHexStringFromDouble_Core 4.94065645841247E-324
    Test_GetHexStringFromDouble_Core 2.2250738585072E-308
    Test_GetHexStringFromDouble_Core 1.7976931348623E+308
    
    Test_GetHexStringFromDouble_Core -0#
    Test_GetHexStringFromDouble_Core -4.94065645841247E-324
    Test_GetHexStringFromDouble_Core -2.2250738585072E-308
    Test_GetHexStringFromDouble_Core -1.7976931348623E+308
    
    Test_GetHexStringFromDouble_Core 1#
    Test_GetHexStringFromDouble_Core -1#
    Test_GetHexStringFromDouble_Core 0.5
    Test_GetHexStringFromDouble_Core 0.1
    Test_GetHexStringFromDouble_Core 1# / 3#
End Sub

Public Sub Test_GetSingleFromHexString()
    ' Positive Zero
    Test_GetSingleFromHexString_Core "00000000"
    
    ' Positive SubNormal Minimum
    Test_GetSingleFromHexString_Core "00000001"
    
    ' Positive SubNormal Maximum
    Test_GetSingleFromHexString_Core "007FFFFF"
    
    ' Positive Normal Minimum
    Test_GetSingleFromHexString_Core "00800000"
    
    ' Positive Normal Maximum
    Test_GetSingleFromHexString_Core "7F7FFFFF"
    
    ' Positive Infinity
    Test_GetSingleFromHexString_Core "7F800000"
    
    ' Positive NaN
    Test_GetSingleFromHexString_Core "7FFFFFFF"
    
    ' Negative Zero
    Test_GetSingleFromHexString_Core "80000000"
    
    ' Negative SubNormal Minimum
    Test_GetSingleFromHexString_Core "80000001"
    
    ' Negative SubNormal Maximum
    Test_GetSingleFromHexString_Core "807FFFFF"
    
    ' Negative Normal Minimum
    Test_GetSingleFromHexString_Core "80800000"
    
    ' Negative Normal Maximum
    Test_GetSingleFromHexString_Core "FF7FFFFF"
    
    ' Negative Infinity
    Test_GetSingleFromHexString_Core "FF800000"
    
    ' Negative NaN
    Test_GetSingleFromHexString_Core "FFFFFFFF"
End Sub

Public Sub Test_GetDoubleFromHexString()
    ' Positive Zero
    Test_GetDoubleFromHexString_Core "0000000000000000"
    
    ' Positive SubNormal Minimum
    Test_GetDoubleFromHexString_Core "0000000000000001"
    
    ' Positive SubNormal Maximum
    Test_GetDoubleFromHexString_Core "000FFFFFFFFFFFFF"
    
    ' Positive Normal Minimum
    Test_GetDoubleFromHexString_Core "0010000000000000"
    
    ' Positive Normal Maximum
    Test_GetDoubleFromHexString_Core "7FEFFFFFFFFFFFFF"
    
    ' Positive Infinity
    Test_GetDoubleFromHexString_Core "7FF0000000000000"
    
    ' Positive NaN
    Test_GetDoubleFromHexString_Core "7FFFFFFFFFFFFFFF"
    
    ' Negative Zero
    Test_GetDoubleFromHexString_Core "8000000000000000"
    
    ' Negative SubNormal Minimum
    Test_GetDoubleFromHexString_Core "8000000000000001"
    
    ' Negative SubNormal Maximum
    Test_GetDoubleFromHexString_Core "800FFFFFFFFFFFFF"
    
    ' Negative Normal Minimum
    Test_GetDoubleFromHexString_Core "8010000000000000"
    
    ' Negative Normal Maximum
    Test_GetDoubleFromHexString_Core "FFEFFFFFFFFFFFFF"
    
    ' Negative Infinity
    Test_GetDoubleFromHexString_Core "FFF0000000000000"
    
    ' Negative NaN
    Test_GetDoubleFromHexString_Core "FFFFFFFFFFFFFFFF"
End Sub

'
' --- Test Core ---
'

Public Sub Test_GetBinStringFromByte_Core(ByVal Value As Byte)
    Dim BinString As String
    BinString = GetBinStringFromByte(Value, True)
    
    Dim Result As Byte
    Result = GetByteFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetBinStringFromInteger_Core(ByVal Value As Integer)
    Dim BinString As String
    BinString = GetBinStringFromInteger(Value, True)
    
    Dim Result As Integer
    Result = GetIntegerFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetBinStringFromLong_Core(ByVal Value As Long)
    Dim BinString As String
    BinString = GetBinStringFromLong(Value, True)
    
    Dim Result As Long
    Result = GetLongFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetBinStringFromLongLong_Core(ByVal Value As LongLong)
    Dim BinString As String
    BinString = GetBinStringFromLongLong(Value, True)
    
    Dim Result As LongLong
    Result = GetLongLongFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

Public Sub Test_GetBinStringFromSingle_Core(ByVal Value As Single)
    Dim BinString As String
    BinString = GetBinStringFromSingle(Value, True)
    
    Dim Result As Single
    Result = GetSingleFromBinString(BinString)
    
    Debug_Print CStr(Value) & " = " & _
        BinString & " = " & CStr(Result)
End Sub

Public Sub Test_GetBinStringFromDouble_Core(ByVal Value As Double)
    Dim BinString As String
    BinString = GetBinStringFromDouble(Value, True)
    
    Dim Result As Double
    Result = GetDoubleFromBinString(BinString)
    
    Debug_Print CStr(Value) & " = " & _
        BinString & " = " & CStr(Result)
End Sub

Public Sub Test_GetSingleFromBinString_Core(BinString As String)
    Dim Value As Single
    Value = GetSingleFromBinString(BinString)
    
    Dim Result As String
    Result = GetBinStringFromSingle(Value, True)
    
    Debug_Print BinString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetDoubleFromBinString_Core(BinString As String)
    Dim Value As Double
    Value = GetDoubleFromBinString(BinString)
    
    Dim Result As String
    Result = GetBinStringFromDouble(Value, True)
    
    Debug_Print BinString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetOctStringFromByte_Core(ByVal Value As Byte)
    Dim OctString As String
    OctString = GetOctStringFromByte(Value, True)
    
    Dim Result As Byte
    Result = GetByteFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetOctStringFromInteger_Core(ByVal Value As Integer)
    Dim OctString As String
    OctString = GetOctStringFromInteger(Value, True)
    
    Dim Result As Integer
    Result = GetIntegerFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetOctStringFromLong_Core(ByVal Value As Long)
    Dim OctString As String
    OctString = GetOctStringFromLong(Value, True)
    
    Dim Result As Long
    Result = GetLongFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetOctStringFromLongLong_Core(ByVal Value As LongLong)
    Dim OctString As String
    OctString = GetOctStringFromLongLong(Value, True)
    
    Dim Result As LongLong
    Result = GetLongLongFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

Public Sub Test_GetOctStringFromSingle_Core(ByVal Value As Single)
    Dim OctString As String
    OctString = GetOctStringFromSingle(Value, True)
    
    Dim Result As Single
    Result = GetSingleFromOctString(OctString)
    
    Debug_Print CStr(Value) & " = " & _
        OctString & " = " & CStr(Result)
End Sub

Public Sub Test_GetOctStringFromDouble_Core(ByVal Value As Double)
    Dim OctString As String
    OctString = GetOctStringFromDouble(Value, True)
    
    Dim Result As Double
    Result = GetDoubleFromOctString(OctString)
    
    Debug_Print CStr(Value) & " = " & _
        OctString & " = " & CStr(Result)
End Sub

Public Sub Test_GetSingleFromOctString_Core(OctString As String)
    Dim Value As Single
    Value = GetSingleFromOctString(OctString)
    
    Dim Result As String
    Result = GetOctStringFromSingle(Value, True)
    
    Debug_Print OctString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetDoubleFromOctString_Core(OctString As String)
    Dim Value As Double
    Value = GetDoubleFromOctString(OctString)
    
    Dim Result As String
    Result = GetOctStringFromDouble(Value, True)
    
    Debug_Print OctString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetHexStringFromByte_Core(ByVal Value As Byte)
    Dim HexString As String
    HexString = GetHexStringFromByte(Value, True)
    
    Dim Result As Byte
    Result = GetByteFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetHexStringFromInteger_Core(ByVal Value As Integer)
    Dim HexString As String
    HexString = GetHexStringFromInteger(Value, True)
    
    Dim Result As Integer
    Result = GetIntegerFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetHexStringFromLong_Core(ByVal Value As Long)
    Dim HexString As String
    HexString = GetHexStringFromLong(Value, True)
    
    Dim Result As Long
    Result = GetLongFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetHexStringFromLongLong_Core(ByVal Value As LongLong)
    Dim HexString As String
    HexString = GetHexStringFromLongLong(Value, True)
    
    Dim Result As LongLong
    Result = GetLongLongFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

Public Sub Test_GetHexStringFromSingle_Core(ByVal Value As Single)
    Dim HexString As String
    HexString = GetHexStringFromSingle(Value, True)
    
    Dim Result As Single
    Result = GetSingleFromHexString(HexString)
    
    Debug_Print CStr(Value) & " = " & _
        HexString & " = " & CStr(Result)
End Sub

Public Sub Test_GetHexStringFromDouble_Core(ByVal Value As Double)
    Dim HexString As String
    HexString = GetHexStringFromDouble(Value, True)
    
    Dim Result As Double
    Result = GetDoubleFromHexString(HexString)
    
    Debug_Print CStr(Value) & " = " & _
        HexString & " = " & CStr(Result)
End Sub

Public Sub Test_GetSingleFromHexString_Core(HexString As String)
    Dim Value As Single
    Value = GetSingleFromHexString(HexString)
    
    Dim Result As String
    Result = GetHexStringFromSingle(Value, True)
    
    Debug_Print HexString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetDoubleFromHexString_Core(HexString As String)
    Dim Value As Double
    Value = GetDoubleFromHexString(HexString)
    
    Dim Result As String
    Result = GetHexStringFromDouble(Value, True)
    
    Debug_Print HexString & " = " & CStr(Value) & " = " & Result
End Sub

Attribute VB_Name = "Test_BitStringF"
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
    Test_GetBinStringFromSingle_Core 12.375!
    Test_GetBinStringFromSingle_Core 0.5!
    Test_GetBinStringFromSingle_Core 0.375!
    Test_GetBinStringFromSingle_Core 0.25!
    Test_GetBinStringFromSingle_Core 0.1!
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
    
    'Test_GetSingleFromBinString_Core "00000000000000000000000000000010"
    'Test_GetSingleFromBinString_Core "00000000000000000000000000000100"
    'Test_GetSingleFromBinString_Core "00000000000000000000000000001000"
    'Test_GetSingleFromBinString_Core "00000000000000000000000000001111"
    'Test_GetSingleFromBinString_Core "00000000000000000000000011110000"
    'Test_GetSingleFromBinString_Core "00000000000000000000111100000000"
    'Test_GetSingleFromBinString_Core "00000000000000001111000000000000"
    'Test_GetSingleFromBinString_Core "00000000000011110000000000000000"
    'Test_GetSingleFromBinString_Core "00000000011100000000000000000000"
    
    ' Positive SubNormal Maximum
    Test_GetSingleFromBinString_Core "00000000011111111111111111111111"
    
    ' Positive Normal Minimum
    Test_GetSingleFromBinString_Core "00000000100000000000000000000000"
    
    'Test_GetSingleFromBinString_Core "01111000011111111111111111111111"
    'Test_GetSingleFromBinString_Core "01111100000000000000000000000000"
    'Test_GetSingleFromBinString_Core "01111110011111111111111111111111"
    
    ' Positive Normal Maximum
    Test_GetSingleFromBinString_Core "01111111011111111111111111111111"
    
    '' Positive Infinity
    'Test_GetSingleFromBinString_Core "01111111100000000000000000000000"
    
    '' Positive NaN
    'Test_GetSingleFromBinString_Core "01111111111111111111111111111111"
    
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
    
    '' Negative Infinity
    'Test_GetSingleFromBinString_Core "11111111100000000000000000000000"
    
    '' Negative NaN
    'Test_GetSingleFromBinString_Core "11111111111111111111111111111111"
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
    
    '' Positive Infinity
    'Test_GetDoubleFromBinString_Core _
    '    "0111111111110000000000000000000000000000000000000000000000000000"
    
    '' Positive NaN
    'Test_GetDoubleFromBinString_Core _
    '    "0111111111111111111111111111111111111111111111111111111111111111"
    
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
    
    '' Negative Infinity
    'Test_GetDoubleFromBinString_Core _
    '    "1111111111110000000000000000000000000000000000000000000000000000"
    
    '' Negative NaN
    'Test_GetDoubleFromBinString_Core _
    '    "1111111111111111111111111111111111111111111111111111111111111111"
End Sub

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
    
    '' Positive Infinity
    'Test_GetSingleFromOctString_Core "17740000000"
    
    '' Positive NaN
    'Test_GetSingleFromOctString_Core "17777777777"
    
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
    
    '' Negative Infinity
    'Test_GetSingleFromOctString_Core "37740000000"
    
    '' Negative NaN
    'Test_GetSingleFromOctString_Core "37777777777"
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
    
    '' Positive Infinity
    'Test_GetDoubleFromOctString_Core "0777600000000000000000"
    
    '' Positive NaN
    'Test_GetDoubleFromOctString_Core "0777777777777777777777"
    
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
    
    '' Negative Infinity
    'Test_GetDoubleFromOctString_Core "1777600000000000000000"
    
    '' Negative NaN
    'Test_GetDoubleFromOctString_Core "1777777777777777777777"
End Sub

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
    
    '' Positive Infinity
    'Test_GetSingleFromHexString_Core "7F800000"
    
    '' Positive NaN
    'Test_GetSingleFromHexString_Core "7FFFFFFF"
    
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
    
    '' Negative Infinity
    'Test_GetSingleFromHexString_Core "FF800000"
    
    '' Negative NaN
    'Test_GetSingleFromHexString_Core "FFFFFFFF"
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
    
    '' Positive Infinity
    'Test_GetDoubleFromHexString_Core "7FF0000000000000"
    
    '' Positive NaN
    'Test_GetDoubleFromHexString_Core "7FFFFFFFFFFFFFFF"
    
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
    
    '' Negative Infinity
    'Test_GetDoubleFromHexString_Core "FFF0000000000000"
    
    '' Negative NaN
    'Test_GetDoubleFromHexString_Core "FFFFFFFFFFFFFFFF"
End Sub

'
' --- Test Core ---
'

Public Sub Test_GetBinStringFromSingle_Core(ByVal Value)
    Dim BinString
    BinString = GetBinStringFromSingle(Value, True)
    
    Dim Result
    Result = GetSingleFromBinString(BinString)
    
    Debug_Print CStr(Value) & " = " & _
        BinString & " = " & CStr(Result)
End Sub

Public Sub Test_GetBinStringFromDouble_Core(ByVal Value)
    Dim BinString
    BinString = GetBinStringFromDouble(Value, True)
    
    Dim Result
    Result = GetDoubleFromBinString(BinString)
    
    Debug_Print CStr(Value) & " = " & _
        BinString & " = " & CStr(Result)
End Sub

Public Sub Test_GetSingleFromBinString_Core(BinString)
    Dim Value
    Value = GetSingleFromBinString(BinString)
    
    Dim Result
    Result = GetBinStringFromSingle(Value, True)
    
    Debug_Print BinString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetDoubleFromBinString_Core(BinString)
    Dim Value
    Value = GetDoubleFromBinString(BinString)
    
    Dim Result
    Result = GetBinStringFromDouble(Value, True)
    
    Debug_Print BinString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetOctStringFromSingle_Core(ByVal Value)
    Dim OctString
    OctString = GetOctStringFromSingle(Value, True)
    
    Dim Result
    Result = GetSingleFromOctString(OctString)
    
    Debug_Print CStr(Value) & " = " & _
        OctString & " = " & CStr(Result)
End Sub

Public Sub Test_GetOctStringFromDouble_Core(ByVal Value)
    Dim OctString
    OctString = GetOctStringFromDouble(Value, True)
    
    Dim Result
    Result = GetDoubleFromOctString(OctString)
    
    Debug_Print CStr(Value) & " = " & _
        OctString & " = " & CStr(Result)
End Sub

Public Sub Test_GetSingleFromOctString_Core(OctString)
    Dim Value
    Value = GetSingleFromOctString(OctString)
    
    Dim Result
    Result = GetOctStringFromSingle(Value, True)
    
    Debug_Print OctString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetDoubleFromOctString_Core(OctString)
    Dim Value
    Value = GetDoubleFromOctString(OctString)
    
    Dim Result
    Result = GetOctStringFromDouble(Value, True)
    
    Debug_Print OctString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetHexStringFromSingle_Core(ByVal Value)
    Dim HexString
    HexString = GetHexStringFromSingle(Value, True)
    
    Dim Result
    Result = GetSingleFromHexString(HexString)
    
    Debug_Print CStr(Value) & " = " & _
        HexString & " = " & CStr(Result)
End Sub

Public Sub Test_GetHexStringFromDouble_Core(ByVal Value)
    Dim HexString
    HexString = GetHexStringFromDouble(Value, True)
    
    Dim Result
    Result = GetDoubleFromHexString(HexString)
    
    Debug_Print CStr(Value) & " = " & _
        HexString & " = " & CStr(Result)
End Sub

Public Sub Test_GetSingleFromHexString_Core(HexString)
    Dim Value
    Value = GetSingleFromHexString(HexString)
    
    Dim Result
    Result = GetHexStringFromSingle(Value, True)
    
    Debug_Print HexString & " = " & CStr(Value) & " = " & Result
End Sub

Public Sub Test_GetDoubleFromHexString_Core(HexString)
    Dim Value
    Value = GetDoubleFromHexString(HexString)
    
    Dim Result
    Result = GetHexStringFromDouble(Value, True)
    
    Debug_Print HexString & " = " & CStr(Value) & " = " & Result
End Sub

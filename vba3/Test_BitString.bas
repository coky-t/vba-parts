Attribute VB_Name = "Test_BitString"
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
    Dim Index
    For Index = 0 To 7
        Test_GetBinStringFromByte_Core 2 ^ Index
    Next
    Test_GetBinStringFromByte_Core &HF0
    Test_GetBinStringFromByte_Core &HFF
End Sub

Public Sub Test_GetBinStringFromInteger()
    Dim Index
    For Index = 0 To 14
        Test_GetBinStringFromInteger_Core CInt(2 ^ Index)
    Next
    For Index = 0 To 14
        Test_GetBinStringFromInteger_Core CInt(-(2 ^ Index))
    Next
    Test_GetBinStringFromInteger_Core CInt((-(2 ^ 14)) * 2)
End Sub

Public Sub Test_GetBinStringFromLong()
    Dim Index
    For Index = 0 To 30
        Test_GetBinStringFromLong_Core CLng(2 ^ Index)
    Next
    For Index = 0 To 30
        Test_GetBinStringFromLong_Core CLng(-(2 ^ Index))
    Next
    Test_GetBinStringFromLong_Core CLng((-(2 ^ 30)) * 2)
End Sub

#If Win64 Then
Public Sub Test_GetBinStringFromLongLong()
    Dim Index
    For Index = 0 To 62
        Test_GetBinStringFromLongLong_Core CLngLng(2 ^ Index)
    Next
    For Index = 0 To 62
        Test_GetBinStringFromLongLong_Core CLngLng(-(2 ^ Index))
    Next
    Test_GetBinStringFromLongLong_Core CLngLng((-(2 ^ 62)) * 2)
End Sub
#End If

Public Sub Test_GetOctStringFromByte()
    Dim Index
    For Index = 0 To 7
        Test_GetOctStringFromByte_Core 2 ^ Index
    Next
    Test_GetOctStringFromByte_Core &HF0
    Test_GetOctStringFromByte_Core &HFF
End Sub

Public Sub Test_GetOctStringFromInteger()
    Dim Index
    For Index = 0 To 14
        Test_GetOctStringFromInteger_Core CInt(2 ^ Index)
    Next
    For Index = 0 To 14
        Test_GetOctStringFromInteger_Core CInt(-(2 ^ Index))
    Next
    Test_GetOctStringFromInteger_Core CInt((-(2 ^ 14)) * 2)
End Sub

Public Sub Test_GetOctStringFromLong()
    Dim Index
    For Index = 0 To 30
        Test_GetOctStringFromLong_Core CLng(2 ^ Index)
    Next
    For Index = 0 To 30
        Test_GetOctStringFromLong_Core CLng(-(2 ^ Index))
    Next
    Test_GetOctStringFromLong_Core CLng((-(2 ^ 30)) * 2)
End Sub

#If Win64 Then
Public Sub Test_GetOctStringFromLongLong()
    Dim Index
    For Index = 0 To 62
        Test_GetOctStringFromLongLong_Core CLngLng(2 ^ Index)
    Next
    For Index = 0 To 62
        Test_GetOctStringFromLongLong_Core CLngLng(-(2 ^ Index))
    Next
    Test_GetOctStringFromLongLong_Core CLngLng((-(2 ^ 62)) * 2)
End Sub
#End If

Public Sub Test_GetHexStringFromByte()
    Dim Index
    For Index = 0 To 7
        Test_GetHexStringFromByte_Core 2 ^ Index
    Next
    Test_GetHexStringFromByte_Core &HF0
    Test_GetHexStringFromByte_Core &HFF
End Sub

Public Sub Test_GetHexStringFromInteger()
    Dim Index
    For Index = 0 To 14
        Test_GetHexStringFromInteger_Core CInt(2 ^ Index)
    Next
    For Index = 0 To 14
        Test_GetHexStringFromInteger_Core CInt(-(2 ^ Index))
    Next
    Test_GetHexStringFromInteger_Core CInt((-(2 ^ 14)) * 2)
End Sub

Public Sub Test_GetHexStringFromLong()
    Dim Index
    For Index = 0 To 30
        Test_GetHexStringFromLong_Core CLng(2 ^ Index)
    Next
    For Index = 0 To 30
        Test_GetHexStringFromLong_Core CLng(-(2 ^ Index))
    Next
    Test_GetHexStringFromLong_Core CLng((-(2 ^ 30)) * 2)
End Sub

#If Win64 Then
Public Sub Test_GetHexStringFromLongLong()
    Dim Index
    For Index = 0 To 62
        Test_GetHexStringFromLongLong_Core CLngLng(2 ^ Index)
    Next
    For Index = 0 To 62
        Test_GetHexStringFromLongLong_Core CLngLng(-(2 ^ Index))
    Next
    Test_GetHexStringFromLongLong_Core CLngLng((-(2 ^ 62)) * 2)
End Sub
#End If

'
' --- Test Core ---
'

Public Sub Test_GetBinStringFromByte_Core(ByVal Value)
    Dim BinString
    BinString = GetBinStringFromByte(Value, True)
    
    Dim Result
    Result = GetByteFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetBinStringFromInteger_Core(ByVal Value)
    Dim BinString
    BinString = GetBinStringFromInteger(Value, True)
    
    Dim Result
    Result = GetIntegerFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetBinStringFromLong_Core(ByVal Value)
    Dim BinString
    BinString = GetBinStringFromLong(Value, True)
    
    Dim Result
    Result = GetLongFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetBinStringFromLongLong_Core(ByVal Value)
    Dim BinString
    BinString = GetBinStringFromLongLong(Value, True)
    
    Dim Result
    Result = GetLongLongFromBinString(BinString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        BinString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

Public Sub Test_GetOctStringFromByte_Core(ByVal Value)
    Dim OctString
    OctString = GetOctStringFromByte(Value, True)
    
    Dim Result
    Result = GetByteFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetOctStringFromInteger_Core(ByVal Value)
    Dim OctString
    OctString = GetOctStringFromInteger(Value, True)
    
    Dim Result
    Result = GetIntegerFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetOctStringFromLong_Core(ByVal Value)
    Dim OctString
    OctString = GetOctStringFromLong(Value, True)
    
    Dim Result
    Result = GetLongFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetOctStringFromLongLong_Core(ByVal Value)
    Dim OctString
    OctString = GetOctStringFromLongLong(Value, True)
    
    Dim Result
    Result = GetLongLongFromOctString(OctString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        OctString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

Public Sub Test_GetHexStringFromByte_Core(ByVal Value)
    Dim HexString
    HexString = GetHexStringFromByte(Value, True)
    
    Dim Result
    Result = GetByteFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetHexStringFromInteger_Core(ByVal Value)
    Dim HexString
    HexString = GetHexStringFromInteger(Value, True)
    
    Dim Result
    Result = GetIntegerFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetHexStringFromLong_Core(ByVal Value)
    Dim HexString
    HexString = GetHexStringFromLong(Value, True)
    
    Dim Result
    Result = GetLongFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetHexStringFromLongLong_Core(ByVal Value)
    Dim HexString
    HexString = GetHexStringFromLongLong(Value, True)
    
    Dim Result
    Result = GetLongLongFromHexString(HexString)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        HexString & " = " & CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

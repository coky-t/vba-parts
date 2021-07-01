Attribute VB_Name = "Test_ByteArrayX"
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

Public Sub Test_GetByteArrayLEFromInteger()
    Test_GetByteArrayLEFromInteger_Core &H0
    Test_GetByteArrayLEFromInteger_Core &H1
    Test_GetByteArrayLEFromInteger_Core &H8
    Test_GetByteArrayLEFromInteger_Core &H10
    Test_GetByteArrayLEFromInteger_Core &H80
    Test_GetByteArrayLEFromInteger_Core &H100
    Test_GetByteArrayLEFromInteger_Core &H800
    Test_GetByteArrayLEFromInteger_Core &H1000
    Test_GetByteArrayLEFromInteger_Core &H8000
    Test_GetByteArrayLEFromInteger_Core &HF000
    Test_GetByteArrayLEFromInteger_Core &HFF00
    Test_GetByteArrayLEFromInteger_Core &HFFF0
    Test_GetByteArrayLEFromInteger_Core &HFFFF
End Sub

Public Sub Test_GetByteArrayBEFromInteger()
    Test_GetByteArrayBEFromInteger_Core &H0
    Test_GetByteArrayBEFromInteger_Core &H1
    Test_GetByteArrayBEFromInteger_Core &H8
    Test_GetByteArrayBEFromInteger_Core &H10
    Test_GetByteArrayBEFromInteger_Core &H80
    Test_GetByteArrayBEFromInteger_Core &H100
    Test_GetByteArrayBEFromInteger_Core &H800
    Test_GetByteArrayBEFromInteger_Core &H1000
    Test_GetByteArrayBEFromInteger_Core &H8000
    Test_GetByteArrayBEFromInteger_Core &HF000
    Test_GetByteArrayBEFromInteger_Core &HFF00
    Test_GetByteArrayBEFromInteger_Core &HFFF0
    Test_GetByteArrayBEFromInteger_Core &HFFFF
End Sub

Public Sub Test_GetByteArrayLEFromLong()
    Test_GetByteArrayLEFromLong_Core &H0
    Test_GetByteArrayLEFromLong_Core &H1
    Test_GetByteArrayLEFromLong_Core &H8
    Test_GetByteArrayLEFromLong_Core &H10
    Test_GetByteArrayLEFromLong_Core &H80
    Test_GetByteArrayLEFromLong_Core &H100
    Test_GetByteArrayLEFromLong_Core &H800
    Test_GetByteArrayLEFromLong_Core &H1000
    Test_GetByteArrayLEFromLong_Core &H8000&
    Test_GetByteArrayLEFromLong_Core &H10000
    Test_GetByteArrayLEFromLong_Core &H80000
    Test_GetByteArrayLEFromLong_Core &H100000
    Test_GetByteArrayLEFromLong_Core &H800000
    Test_GetByteArrayLEFromLong_Core &H1000000
    Test_GetByteArrayLEFromLong_Core &H8000000
    Test_GetByteArrayLEFromLong_Core &H10000000
    Test_GetByteArrayLEFromLong_Core &H80000000
    Test_GetByteArrayLEFromLong_Core &HF0000000
    Test_GetByteArrayLEFromLong_Core &HFF000000
    Test_GetByteArrayLEFromLong_Core &HFFF00000
    Test_GetByteArrayLEFromLong_Core &HFFFF0000
    Test_GetByteArrayLEFromLong_Core &HFFFFF000
    Test_GetByteArrayLEFromLong_Core &HFFFFFF00
    Test_GetByteArrayLEFromLong_Core &HFFFFFFF0
    Test_GetByteArrayLEFromLong_Core &HFFFFFFFF
End Sub

Public Sub Test_GetByteArrayBEFromLong()
    Test_GetByteArrayBEFromLong_Core &H0
    Test_GetByteArrayBEFromLong_Core &H1
    Test_GetByteArrayBEFromLong_Core &H8
    Test_GetByteArrayBEFromLong_Core &H10
    Test_GetByteArrayBEFromLong_Core &H80
    Test_GetByteArrayBEFromLong_Core &H100
    Test_GetByteArrayBEFromLong_Core &H800
    Test_GetByteArrayBEFromLong_Core &H1000
    Test_GetByteArrayBEFromLong_Core &H8000&
    Test_GetByteArrayBEFromLong_Core &H10000
    Test_GetByteArrayBEFromLong_Core &H80000
    Test_GetByteArrayBEFromLong_Core &H100000
    Test_GetByteArrayBEFromLong_Core &H800000
    Test_GetByteArrayBEFromLong_Core &H1000000
    Test_GetByteArrayBEFromLong_Core &H8000000
    Test_GetByteArrayBEFromLong_Core &H10000000
    Test_GetByteArrayBEFromLong_Core &H80000000
    Test_GetByteArrayBEFromLong_Core &HF0000000
    Test_GetByteArrayBEFromLong_Core &HFF000000
    Test_GetByteArrayBEFromLong_Core &HFFF00000
    Test_GetByteArrayBEFromLong_Core &HFFFF0000
    Test_GetByteArrayBEFromLong_Core &HFFFFF000
    Test_GetByteArrayBEFromLong_Core &HFFFFFF00
    Test_GetByteArrayBEFromLong_Core &HFFFFFFF0
    Test_GetByteArrayBEFromLong_Core &HFFFFFFFF
End Sub

#If Win64 Then
Public Sub Test_GetByteArrayLEFromLongLong()
    Dim Value As LongLong
    Dim Index As Integer
    For Index = 0 To 62
        Test_GetByteArrayLEFromLongLong_Core CLngLng(2 ^ Index)
    Next
    For Index = 0 To 62
        Test_GetByteArrayLEFromLongLong_Core -CLngLng(2 ^ Index)
    Next
    Test_GetByteArrayLEFromLongLong_Core -CLngLng(2 ^ 62) * 2
End Sub

Public Sub Test_GetByteArrayBEFromLongLong()
    Dim Value As LongLong
    Dim Index As Integer
    For Index = 0 To 62
        Test_GetByteArrayBEFromLongLong_Core CLngLng(2 ^ Index)
    Next
    For Index = 0 To 62
        Test_GetByteArrayBEFromLongLong_Core -CLngLng(2 ^ Index)
    Next
    Test_GetByteArrayBEFromLongLong_Core -CLngLng(2 ^ 62) * 2
End Sub
#End If

Public Sub Test_GetByteArrayLEFromSingle()
    Test_GetByteArrayLEFromSingle_Core 0!
    Test_GetByteArrayLEFromSingle_Core 1!
    Test_GetByteArrayLEFromSingle_Core 0.5!
    Test_GetByteArrayLEFromSingle_Core 0.1!
End Sub

Public Sub Test_GetByteArrayBEFromSingle()
    Test_GetByteArrayBEFromSingle_Core 0!
    Test_GetByteArrayBEFromSingle_Core 1!
    Test_GetByteArrayBEFromSingle_Core 0.5!
    Test_GetByteArrayBEFromSingle_Core 0.1!
End Sub

Public Sub Test_GetByteArrayLEFromDouble()
    Test_GetByteArrayLEFromDouble_Core 0#
    Test_GetByteArrayLEFromDouble_Core 1#
    Test_GetByteArrayLEFromDouble_Core 0.5
    Test_GetByteArrayLEFromDouble_Core 0.1
End Sub

Public Sub Test_GetByteArrayBEFromDouble()
    Test_GetByteArrayBEFromDouble_Core 0#
    Test_GetByteArrayBEFromDouble_Core 1#
    Test_GetByteArrayBEFromDouble_Core 0.5
    Test_GetByteArrayBEFromDouble_Core 0.1
End Sub

'
' --- Test Core ---
'

Public Sub Test_GetByteArrayLEFromInteger_Core(ByVal Value As Integer)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromInteger(Value)
    
    Dim Result As Integer
    Result = GetIntegerFromByteArrayLE(ByteArray)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetByteArrayBEFromInteger_Core(ByVal Value As Integer)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayBEFromInteger(Value)
    
    Dim Result As Integer
    Result = GetIntegerFromByteArrayBE(ByteArray)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetByteArrayLEFromLong_Core(ByVal Value As Long)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromLong(Value)
    
    Dim Result As Long
    Result = GetLongFromByteArrayLE(ByteArray)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetByteArrayBEFromLong_Core(ByVal Value As Long)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayBEFromLong(Value)
    
    Dim Result As Long
    Result = GetLongFromByteArrayBE(ByteArray)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetByteArrayLEFromLongLong_Core(ByVal Value As LongLong)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromLongLong(Value)
    
    Dim Result As LongLong
    Result = GetLongLongFromByteArrayLE(ByteArray)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetByteArrayBEFromLongLong_Core(ByVal Value As LongLong)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayBEFromLongLong(Value)
    
    Dim Result As LongLong
    Result = GetLongLongFromByteArrayBE(ByteArray)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

Public Sub Test_GetByteArrayLEFromSingle_Core(ByVal Value As Single)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromSingle(Value)
    
    Dim Result As Single
    Result = GetSingleFromByteArrayLE(ByteArray)
    
    Debug_Print CStr(Value) & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result)
End Sub

Public Sub Test_GetByteArrayBEFromSingle_Core(ByVal Value As Single)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayBEFromSingle(Value)
    
    Dim Result As Single
    Result = GetSingleFromByteArrayBE(ByteArray)
    
    Debug_Print CStr(Value) & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result)
End Sub

Public Sub Test_GetByteArrayLEFromDouble_Core(ByVal Value As Double)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayLEFromDouble(Value)
    
    Dim Result As Double
    Result = GetDoubleFromByteArrayLE(ByteArray)
    
    Debug_Print CStr(Value) & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result)
End Sub

Public Sub Test_GetByteArrayBEFromDouble_Core(ByVal Value As Double)
    Dim ByteArray() As Byte
    ByteArray = GetByteArrayBEFromDouble(Value)
    
    Dim Result As Double
    Result = GetDoubleFromByteArrayBE(ByteArray)
    
    Debug_Print CStr(Value) & " = " & _
        GetByteArrayString(ByteArray) & " = " & _
        CStr(Result)
End Sub

Public Function GetByteArrayString(ByteArray() As Byte) As String
    Dim ByteArrayString As String
    ByteArrayString = Right("0" & Hex(ByteArray(LBound(ByteArray))), 2)
    
    Dim Index As Long
    For Index = LBound(ByteArray) + 1 To UBound(ByteArray)
        ByteArrayString = ByteArrayString & " " & _
            Right("0" & Hex(ByteArray(Index)), 2)
    Next
    
    GetByteArrayString = ByteArrayString
End Function

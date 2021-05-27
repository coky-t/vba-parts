Attribute VB_Name = "Test_ByteArray"
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

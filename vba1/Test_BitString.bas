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

'
' --- Test Core ---
'

Public Sub Test_GetBinStringFromByte_Core(ByVal Value As Byte)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetBinStringFromByte(Value, True)
End Sub

Public Sub Test_GetBinStringFromInteger_Core(ByVal Value As Integer)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetBinStringFromInteger(Value, True)
End Sub

Public Sub Test_GetBinStringFromLong_Core(ByVal Value As Long)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetBinStringFromLong(Value, True)
End Sub

#If Win64 Then
Public Sub Test_GetBinStringFromLongLong_Core(ByVal Value As LongLong)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetBinStringFromLongLong(Value, True)
End Sub
#End If

Public Sub Test_GetOctStringFromByte_Core(ByVal Value As Byte)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetOctStringFromByte(Value, True)
End Sub

Public Sub Test_GetOctStringFromInteger_Core(ByVal Value As Integer)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetOctStringFromInteger(Value, True)
End Sub

Public Sub Test_GetOctStringFromLong_Core(ByVal Value As Long)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetOctStringFromLong(Value, True)
End Sub

#If Win64 Then
Public Sub Test_GetOctStringFromLongLong_Core(ByVal Value As LongLong)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetOctStringFromLongLong(Value, True)
End Sub
#End If

Public Sub Test_GetHexStringFromByte_Core(ByVal Value As Byte)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetHexStringFromByte(Value, True)
End Sub

Public Sub Test_GetHexStringFromInteger_Core(ByVal Value As Integer)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetHexStringFromInteger(Value, True)
End Sub

Public Sub Test_GetHexStringFromLong_Core(ByVal Value As Long)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetHexStringFromLong(Value, True)
End Sub

#If Win64 Then
Public Sub Test_GetHexStringFromLongLong_Core(ByVal Value As LongLong)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetHexStringFromLongLong(Value, True)
End Sub
#End If

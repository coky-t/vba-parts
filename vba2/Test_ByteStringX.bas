Attribute VB_Name = "Test_ByteStringX"
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

Public Sub Test_GetStringB_LEFromInteger()
    Dim Index As Long
    For Index = 0 To 14
        Test_GetStringB_LEFromInteger_Core 2 ^ Index
    Next
    For Index = 0 To 14
        Test_GetStringB_LEFromInteger_Core -(2 ^ Index)
    Next
    Test_GetStringB_LEFromInteger_Core (-(2 ^ 14)) * 2
End Sub

Public Sub Test_GetStringB_BEFromInteger()
    Dim Index As Long
    For Index = 0 To 14
        Test_GetStringB_BEFromInteger_Core 2 ^ Index
    Next
    For Index = 0 To 14
        Test_GetStringB_BEFromInteger_Core -(2 ^ Index)
    Next
    Test_GetStringB_BEFromInteger_Core (-(2 ^ 14)) * 2
End Sub

Public Sub Test_GetStringB_LEFromLong()
    Dim Index As Long
    For Index = 0 To 30
        Test_GetStringB_LEFromLong_Core 2 ^ Index
    Next
    For Index = 0 To 30
        Test_GetStringB_LEFromLong_Core -(2 ^ Index)
    Next
    Test_GetStringB_LEFromLong_Core (-(2 ^ 30)) * 2
End Sub

Public Sub Test_GetStringB_BEFromLong()
    Dim Index As Long
    For Index = 0 To 30
        Test_GetStringB_BEFromLong_Core 2 ^ Index
    Next
    For Index = 0 To 30
        Test_GetStringB_BEFromLong_Core -(2 ^ Index)
    Next
    Test_GetStringB_BEFromLong_Core (-(2 ^ 30)) * 2
End Sub

#If Win64 Then
Public Sub Test_GetStringB_LEFromLongLong()
    Dim Value As LongLong
    Dim Index As Integer
    For Index = 0 To 62
        Test_GetStringB_LEFromLongLong_Core CLngLng(2 ^ Index)
    Next
    For Index = 0 To 62
        Test_GetStringB_LEFromLongLong_Core -CLngLng(2 ^ Index)
    Next
    Test_GetStringB_LEFromLongLong_Core -CLngLng(2 ^ 62) * 2
End Sub

Public Sub Test_GetStringB_BEFromLongLong()
    Dim Value As LongLong
    Dim Index As Integer
    For Index = 0 To 62
        Test_GetStringB_BEFromLongLong_Core CLngLng(2 ^ Index)
    Next
    For Index = 0 To 62
        Test_GetStringB_BEFromLongLong_Core -CLngLng(2 ^ Index)
    Next
    Test_GetStringB_BEFromLongLong_Core -CLngLng(2 ^ 62) * 2
End Sub
#End If

Public Sub Test_GetStringB_LEFromSingle()
    Test_GetStringB_LEFromSingle_Core 0!
    Test_GetStringB_LEFromSingle_Core 1!
    Test_GetStringB_LEFromSingle_Core 0.5!
    Test_GetStringB_LEFromSingle_Core 0.1!
End Sub

Public Sub Test_GetStringB_BEFromSingle()
    Test_GetStringB_BEFromSingle_Core 0!
    Test_GetStringB_BEFromSingle_Core 1!
    Test_GetStringB_BEFromSingle_Core 0.5!
    Test_GetStringB_BEFromSingle_Core 0.1!
End Sub

Public Sub Test_GetStringB_LEFromDouble()
    Test_GetStringB_LEFromDouble_Core 0#
    Test_GetStringB_LEFromDouble_Core 1#
    Test_GetStringB_LEFromDouble_Core 0.5
    Test_GetStringB_LEFromDouble_Core 0.1
End Sub

Public Sub Test_GetStringB_BEFromDouble()
    Test_GetStringB_BEFromDouble_Core 0#
    Test_GetStringB_BEFromDouble_Core 1#
    Test_GetStringB_BEFromDouble_Core 0.5
    Test_GetStringB_BEFromDouble_Core 0.1
End Sub

'
' --- Test Core ---
'

Public Sub Test_GetStringB_LEFromInteger_Core(ByVal Value As Integer)
    Dim StrB As String
    StrB = GetStringB_LEFromInteger(Value)
    
    Dim Result As Integer
    Result = GetIntegerFromStringB_LE(StrB)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetStringB_BEFromInteger_Core(ByVal Value As Integer)
    Dim StrB As String
    StrB = GetStringB_BEFromInteger(Value)
    
    Dim Result As Integer
    Result = GetIntegerFromStringB_BE(StrB)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetStringB_LEFromLong_Core(ByVal Value As Long)
    Dim StrB As String
    StrB = GetStringB_LEFromLong(Value)
    
    Dim Result As Long
    Result = GetLongFromStringB_LE(StrB)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetStringB_BEFromLong_Core(ByVal Value As Long)
    Dim StrB As String
    StrB = GetStringB_BEFromLong(Value)
    
    Dim Result As Long
    Result = GetLongFromStringB_BE(StrB)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

#If Win64 Then
Public Sub Test_GetStringB_LEFromLongLong_Core(ByVal Value As LongLong)
    Dim StrB As String
    StrB = GetStringB_LEFromLongLong(Value)
    
    Dim Result As LongLong
    Result = GetLongLongFromStringB_LE(StrB)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub

Public Sub Test_GetStringB_BEFromLongLong_Core(ByVal Value As LongLong)
    Dim StrB As String
    StrB = GetStringB_BEFromLongLong(Value)
    
    Dim Result As LongLong
    Result = GetLongLongFromStringB_BE(StrB)
    
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result) & "(" & Hex(Result) & ")"
End Sub
#End If

Public Sub Test_GetStringB_LEFromSingle_Core(ByVal Value As Single)
    Dim StrB As String
    StrB = GetStringB_LEFromSingle(Value)
    
    Dim Result As Single
    Result = GetSingleFromStringB_LE(StrB)
    
    Debug_Print CStr(Value) & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result)
End Sub

Public Sub Test_GetStringB_BEFromSingle_Core(ByVal Value As Single)
    Dim StrB As String
    StrB = GetStringB_BEFromSingle(Value)
    
    Dim Result As Single
    Result = GetSingleFromStringB_BE(StrB)
    
    Debug_Print CStr(Value) & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result)
End Sub

Public Sub Test_GetStringB_LEFromDouble_Core(ByVal Value As Double)
    Dim StrB As String
    StrB = GetStringB_LEFromDouble(Value)
    
    Dim Result As Double
    Result = GetDoubleFromStringB_LE(StrB)
    
    Debug_Print CStr(Value) & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result)
End Sub

Public Sub Test_GetStringB_BEFromDouble_Core(ByVal Value As Double)
    Dim StrB As String
    StrB = GetStringB_BEFromDouble(Value)
    
    Dim Result As Double
    Result = GetDoubleFromStringB_BE(StrB)
    
    Debug_Print CStr(Value) & " = " & _
        GetDebugStringFromStrB(StrB) & " = " & _
        CStr(Result)
End Sub

Public Function GetDebugStringFromStrB(StrB As String) As String
    Dim DebugString As String
    DebugString = Right("0" & Hex(AscB(MidB(StrB, 1, 1))), 2)
    
    Dim Index As Long
    For Index = 2 To LenB(StrB)
        DebugString = DebugString & " " & _
            Right("0" & Hex(AscB(MidB(StrB, Index, 1))), 2)
    Next
    
    GetDebugStringFromStrB = DebugString
End Function

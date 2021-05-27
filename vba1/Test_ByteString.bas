Attribute VB_Name = "Test_ByteString"
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
    Test_GetStringB_LEFromInteger_Core &H0
    Test_GetStringB_LEFromInteger_Core &H1
    Test_GetStringB_LEFromInteger_Core &H8
    Test_GetStringB_LEFromInteger_Core &H10
    Test_GetStringB_LEFromInteger_Core &H80
    Test_GetStringB_LEFromInteger_Core &H100
    Test_GetStringB_LEFromInteger_Core &H800
    Test_GetStringB_LEFromInteger_Core &H1000
    Test_GetStringB_LEFromInteger_Core &H8000
    Test_GetStringB_LEFromInteger_Core &HF000
    Test_GetStringB_LEFromInteger_Core &HFF00
    Test_GetStringB_LEFromInteger_Core &HFFF0
    Test_GetStringB_LEFromInteger_Core &HFFFF
End Sub

Public Sub Test_GetStringB_BEFromInteger()
    Test_GetStringB_BEFromInteger_Core &H0
    Test_GetStringB_BEFromInteger_Core &H1
    Test_GetStringB_BEFromInteger_Core &H8
    Test_GetStringB_BEFromInteger_Core &H10
    Test_GetStringB_BEFromInteger_Core &H80
    Test_GetStringB_BEFromInteger_Core &H100
    Test_GetStringB_BEFromInteger_Core &H800
    Test_GetStringB_BEFromInteger_Core &H1000
    Test_GetStringB_BEFromInteger_Core &H8000
    Test_GetStringB_BEFromInteger_Core &HF000
    Test_GetStringB_BEFromInteger_Core &HFF00
    Test_GetStringB_BEFromInteger_Core &HFFF0
    Test_GetStringB_BEFromInteger_Core &HFFFF
End Sub

Public Sub Test_GetStringB_LEFromLong()
    Test_GetStringB_LEFromLong_Core &H0
    Test_GetStringB_LEFromLong_Core &H1
    Test_GetStringB_LEFromLong_Core &H8
    Test_GetStringB_LEFromLong_Core &H10
    Test_GetStringB_LEFromLong_Core &H80
    Test_GetStringB_LEFromLong_Core &H100
    Test_GetStringB_LEFromLong_Core &H800
    Test_GetStringB_LEFromLong_Core &H1000
    Test_GetStringB_LEFromLong_Core &H8000&
    Test_GetStringB_LEFromLong_Core &H10000
    Test_GetStringB_LEFromLong_Core &H80000
    Test_GetStringB_LEFromLong_Core &H100000
    Test_GetStringB_LEFromLong_Core &H800000
    Test_GetStringB_LEFromLong_Core &H1000000
    Test_GetStringB_LEFromLong_Core &H8000000
    Test_GetStringB_LEFromLong_Core &H10000000
    Test_GetStringB_LEFromLong_Core &H80000000
    Test_GetStringB_LEFromLong_Core &HF0000000
    Test_GetStringB_LEFromLong_Core &HFF000000
    Test_GetStringB_LEFromLong_Core &HFFF00000
    Test_GetStringB_LEFromLong_Core &HFFFF0000
    Test_GetStringB_LEFromLong_Core &HFFFFF000
    Test_GetStringB_LEFromLong_Core &HFFFFFF00
    Test_GetStringB_LEFromLong_Core &HFFFFFFF0
    Test_GetStringB_LEFromLong_Core &HFFFFFFFF
End Sub

Public Sub Test_GetStringB_BEFromLong()
    Test_GetStringB_BEFromLong_Core &H0
    Test_GetStringB_BEFromLong_Core &H1
    Test_GetStringB_BEFromLong_Core &H8
    Test_GetStringB_BEFromLong_Core &H10
    Test_GetStringB_BEFromLong_Core &H80
    Test_GetStringB_BEFromLong_Core &H100
    Test_GetStringB_BEFromLong_Core &H800
    Test_GetStringB_BEFromLong_Core &H1000
    Test_GetStringB_BEFromLong_Core &H8000&
    Test_GetStringB_BEFromLong_Core &H10000
    Test_GetStringB_BEFromLong_Core &H80000
    Test_GetStringB_BEFromLong_Core &H100000
    Test_GetStringB_BEFromLong_Core &H800000
    Test_GetStringB_BEFromLong_Core &H1000000
    Test_GetStringB_BEFromLong_Core &H8000000
    Test_GetStringB_BEFromLong_Core &H10000000
    Test_GetStringB_BEFromLong_Core &H80000000
    Test_GetStringB_BEFromLong_Core &HF0000000
    Test_GetStringB_BEFromLong_Core &HFF000000
    Test_GetStringB_BEFromLong_Core &HFFF00000
    Test_GetStringB_BEFromLong_Core &HFFFF0000
    Test_GetStringB_BEFromLong_Core &HFFFFF000
    Test_GetStringB_BEFromLong_Core &HFFFFFF00
    Test_GetStringB_BEFromLong_Core &HFFFFFFF0
    Test_GetStringB_BEFromLong_Core &HFFFFFFFF
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

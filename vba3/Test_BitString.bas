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
    Test_GetBinStringFromByte_Core &H0
    Test_GetBinStringFromByte_Core &H1
    Test_GetBinStringFromByte_Core &H2
    Test_GetBinStringFromByte_Core &H4
    Test_GetBinStringFromByte_Core &H8
    Test_GetBinStringFromByte_Core &H10
    Test_GetBinStringFromByte_Core &H20
    Test_GetBinStringFromByte_Core &H40
    Test_GetBinStringFromByte_Core &H80
    Test_GetBinStringFromByte_Core &HF0
    Test_GetBinStringFromByte_Core &HFF
End Sub

Public Sub Test_GetBinStringFromInteger()
    Test_GetBinStringFromInteger_Core &H0
    Test_GetBinStringFromInteger_Core &H1
    Test_GetBinStringFromInteger_Core &H8
    Test_GetBinStringFromInteger_Core &H10
    Test_GetBinStringFromInteger_Core &H80
    Test_GetBinStringFromInteger_Core &H100
    Test_GetBinStringFromInteger_Core &H800
    Test_GetBinStringFromInteger_Core &H1000
    Test_GetBinStringFromInteger_Core &H8000
    Test_GetBinStringFromInteger_Core &HF000
    Test_GetBinStringFromInteger_Core &HFF00
    Test_GetBinStringFromInteger_Core &HFFF0
    Test_GetBinStringFromInteger_Core &HFFFF
End Sub

Public Sub Test_GetBinStringFromLong()
    Test_GetBinStringFromLong_Core &H0
    Test_GetBinStringFromLong_Core &H1
    Test_GetBinStringFromLong_Core &H8
    Test_GetBinStringFromLong_Core &H10
    Test_GetBinStringFromLong_Core &H80
    Test_GetBinStringFromLong_Core &H100
    Test_GetBinStringFromLong_Core &H800
    Test_GetBinStringFromLong_Core &H1000&
    Test_GetBinStringFromLong_Core &H8000&
    Test_GetBinStringFromLong_Core &H10000
    Test_GetBinStringFromLong_Core &H80000
    Test_GetBinStringFromLong_Core &H100000
    Test_GetBinStringFromLong_Core &H800000
    Test_GetBinStringFromLong_Core &H1000000
    Test_GetBinStringFromLong_Core &H8000000
    Test_GetBinStringFromLong_Core &H10000000
    Test_GetBinStringFromLong_Core &H80000000
    Test_GetBinStringFromLong_Core &HF0000000
    Test_GetBinStringFromLong_Core &HFF000000
    Test_GetBinStringFromLong_Core &HFFF00000
    Test_GetBinStringFromLong_Core &HFFFF0000
    Test_GetBinStringFromLong_Core &HFFFFF000
    Test_GetBinStringFromLong_Core &HFFFFFF00
    Test_GetBinStringFromLong_Core &HFFFFFFF0
    Test_GetBinStringFromLong_Core &HFFFFFFFF
End Sub

Public Sub Test_GetOctStringFromByte()
    Test_GetOctStringFromByte_Core &H0
    Test_GetOctStringFromByte_Core &H1
    Test_GetOctStringFromByte_Core &H2
    Test_GetOctStringFromByte_Core &H4
    Test_GetOctStringFromByte_Core &H8
    Test_GetOctStringFromByte_Core &H10
    Test_GetOctStringFromByte_Core &H20
    Test_GetOctStringFromByte_Core &H40
    Test_GetOctStringFromByte_Core &H80
    Test_GetOctStringFromByte_Core &HF0
    Test_GetOctStringFromByte_Core &HFF
End Sub

Public Sub Test_GetOctStringFromInteger()
    Test_GetOctStringFromInteger_Core &H0
    Test_GetOctStringFromInteger_Core &H1
    Test_GetOctStringFromInteger_Core &H8
    Test_GetOctStringFromInteger_Core &H10
    Test_GetOctStringFromInteger_Core &H80
    Test_GetOctStringFromInteger_Core &H100
    Test_GetOctStringFromInteger_Core &H800
    Test_GetOctStringFromInteger_Core &H1000
    Test_GetOctStringFromInteger_Core &H8000
    Test_GetOctStringFromInteger_Core &HF000
    Test_GetOctStringFromInteger_Core &HFF00
    Test_GetOctStringFromInteger_Core &HFFF0
    Test_GetOctStringFromInteger_Core &HFFFF
End Sub

Public Sub Test_GetOctStringFromLong()
    Test_GetOctStringFromLong_Core &H0
    Test_GetOctStringFromLong_Core &H1
    Test_GetOctStringFromLong_Core &H8
    Test_GetOctStringFromLong_Core &H10
    Test_GetOctStringFromLong_Core &H80
    Test_GetOctStringFromLong_Core &H100
    Test_GetOctStringFromLong_Core &H800
    Test_GetOctStringFromLong_Core &H1000&
    Test_GetOctStringFromLong_Core &H8000&
    Test_GetOctStringFromLong_Core &H10000
    Test_GetOctStringFromLong_Core &H80000
    Test_GetOctStringFromLong_Core &H100000
    Test_GetOctStringFromLong_Core &H800000
    Test_GetOctStringFromLong_Core &H1000000
    Test_GetOctStringFromLong_Core &H8000000
    Test_GetOctStringFromLong_Core &H10000000
    Test_GetOctStringFromLong_Core &H80000000
    Test_GetOctStringFromLong_Core &HF0000000
    Test_GetOctStringFromLong_Core &HFF000000
    Test_GetOctStringFromLong_Core &HFFF00000
    Test_GetOctStringFromLong_Core &HFFFF0000
    Test_GetOctStringFromLong_Core &HFFFFF000
    Test_GetOctStringFromLong_Core &HFFFFFF00
    Test_GetOctStringFromLong_Core &HFFFFFFF0
    Test_GetOctStringFromLong_Core &HFFFFFFFF
End Sub

Public Sub Test_GetHexStringFromByte()
    Test_GetHexStringFromByte_Core &H0
    Test_GetHexStringFromByte_Core &H1
    Test_GetHexStringFromByte_Core &H2
    Test_GetHexStringFromByte_Core &H4
    Test_GetHexStringFromByte_Core &H8
    Test_GetHexStringFromByte_Core &H10
    Test_GetHexStringFromByte_Core &H20
    Test_GetHexStringFromByte_Core &H40
    Test_GetHexStringFromByte_Core &H80
    Test_GetHexStringFromByte_Core &HF0
    Test_GetHexStringFromByte_Core &HFF
End Sub

Public Sub Test_GetHexStringFromInteger()
    Test_GetHexStringFromInteger_Core &H0
    Test_GetHexStringFromInteger_Core &H1
    Test_GetHexStringFromInteger_Core &H8
    Test_GetHexStringFromInteger_Core &H10
    Test_GetHexStringFromInteger_Core &H80
    Test_GetHexStringFromInteger_Core &H100
    Test_GetHexStringFromInteger_Core &H800
    Test_GetHexStringFromInteger_Core &H1000
    Test_GetHexStringFromInteger_Core &H8000
    Test_GetHexStringFromInteger_Core &HF000
    Test_GetHexStringFromInteger_Core &HFF00
    Test_GetHexStringFromInteger_Core &HFFF0
    Test_GetHexStringFromInteger_Core &HFFFF
End Sub

Public Sub Test_GetHexStringFromLong()
    Test_GetHexStringFromLong_Core &H0
    Test_GetHexStringFromLong_Core &H1
    Test_GetHexStringFromLong_Core &H8
    Test_GetHexStringFromLong_Core &H10
    Test_GetHexStringFromLong_Core &H80
    Test_GetHexStringFromLong_Core &H100
    Test_GetHexStringFromLong_Core &H800
    Test_GetHexStringFromLong_Core &H1000&
    Test_GetHexStringFromLong_Core &H8000&
    Test_GetHexStringFromLong_Core &H10000
    Test_GetHexStringFromLong_Core &H80000
    Test_GetHexStringFromLong_Core &H100000
    Test_GetHexStringFromLong_Core &H800000
    Test_GetHexStringFromLong_Core &H1000000
    Test_GetHexStringFromLong_Core &H8000000
    Test_GetHexStringFromLong_Core &H10000000
    Test_GetHexStringFromLong_Core &H80000000
    Test_GetHexStringFromLong_Core &HF0000000
    Test_GetHexStringFromLong_Core &HFF000000
    Test_GetHexStringFromLong_Core &HFFF00000
    Test_GetHexStringFromLong_Core &HFFFF0000
    Test_GetHexStringFromLong_Core &HFFFFF000
    Test_GetHexStringFromLong_Core &HFFFFFF00
    Test_GetHexStringFromLong_Core &HFFFFFFF0
    Test_GetHexStringFromLong_Core &HFFFFFFFF
End Sub

'
' --- Test Core ---
'

Public Sub Test_GetBinStringFromByte_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetBinStringFromByte(Value, True)
End Sub

Public Sub Test_GetBinStringFromInteger_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetBinStringFromInteger(Value, True)
End Sub

Public Sub Test_GetBinStringFromLong_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetBinStringFromLong(Value, True)
End Sub

Public Sub Test_GetOctStringFromByte_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetOctStringFromByte(Value, True)
End Sub

Public Sub Test_GetOctStringFromInteger_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetOctStringFromInteger(Value, True)
End Sub

Public Sub Test_GetOctStringFromLong_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetOctStringFromLong(Value, True)
End Sub

Public Sub Test_GetHexStringFromByte_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetHexStringFromByte(Value, True)
End Sub

Public Sub Test_GetHexStringFromInteger_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetHexStringFromInteger(Value, True)
End Sub

Public Sub Test_GetHexStringFromLong_Core(ByVal Value)
    Debug_Print CStr(Value) & "(" & Hex(Value) & ")" & " = " & _
        GetHexStringFromLong(Value, True)
End Sub

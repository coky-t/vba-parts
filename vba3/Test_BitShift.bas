Attribute VB_Name = "Test_BitShift"
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

Public Sub Test_LeftShiftByte01()
    Test_LeftShiftByteX &H1
End Sub

Public Sub Test_LeftShiftByteFF()
    Test_LeftShiftByteX &HFF
End Sub

Public Sub Test_LeftShiftInteger0001()
    Test_LeftShiftIntegerX &H1
End Sub

Public Sub Test_LeftShiftIntegerFFFF()
    Test_LeftShiftIntegerX &HFFFF
End Sub

Public Sub Test_LeftShiftLong00000001()
    Test_LeftShiftLongX &H1
End Sub

Public Sub Test_LeftShiftLongFFFFFFFF()
    Test_LeftShiftLongX &HFFFFFFFF
End Sub

#If Win64 Then
Public Sub Test_LeftShiftLongLong0000000000000001()
    Test_LeftShiftLongLongX &H1
End Sub

Public Sub Test_LeftShiftLongLongFFFFFFFFFFFFFFFF()
    Test_LeftShiftLongLongX (-1)
End Sub
#End If

Public Sub Test_RightArithmeticShiftByte80()
    Test_RightArithmeticShiftByteX &H80
End Sub

Public Sub Test_RightArithmeticShiftByte7F()
    Test_RightArithmeticShiftByteX &H7F
End Sub

Public Sub Test_RightArithmeticShiftByteBF()
    Test_RightArithmeticShiftByteX &HBF
End Sub

Public Sub Test_RightArithmeticShiftInteger8000()
    Test_RightArithmeticShiftIntegerX &H8000
End Sub

Public Sub Test_RightArithmeticShiftInteger7FFF()
    Test_RightArithmeticShiftIntegerX &H7FFF
End Sub

Public Sub Test_RightArithmeticShiftIntegerBFFF()
    Test_RightArithmeticShiftIntegerX &HBFFF
End Sub

Public Sub Test_RightArithmeticShiftLong80000000()
    Test_RightArithmeticShiftLongX &H80000000
End Sub

Public Sub Test_RightArithmeticShiftLong7FFFFFFF()
    Test_RightArithmeticShiftLongX &H7FFFFFFF
End Sub

Public Sub Test_RightArithmeticShiftLongBFFFFFFF()
    Test_RightArithmeticShiftLongX &HBFFFFFFF
End Sub

#If Win64 Then
Public Sub Test_RightArithmeticShiftLongLong7FFFFFFFFFFFFFFF()
    Dim Value
    Dim Index
    For Index = 0 To 62
        Value = Value Or CLngLng(2 ^ Index)
    Next
    Test_RightArithmeticShiftLongLongX Value
End Sub

Public Sub Test_RightArithmeticShiftLongLongBFFFFFFFFFFFFFFF()
    Test_RightArithmeticShiftLongLongX Not CLngLng(2 ^ 62)
End Sub
#End If

Public Sub Test_RightShiftByte80()
    Test_RightShiftByteX &H80
End Sub

Public Sub Test_RightShiftByteFF()
    Test_RightShiftByteX &HFF
End Sub

Public Sub Test_RightShiftInteger8000()
    Test_RightShiftIntegerX &H8000
End Sub

Public Sub Test_RightShiftIntegerFFFF()
    Test_RightShiftIntegerX &HFFFF
End Sub

Public Sub Test_RightShiftLong80000000()
    Test_RightShiftLongX &H80000000
End Sub

Public Sub Test_RightShiftLongFFFFFFFF()
    Test_RightShiftLongX &HFFFFFFFF
End Sub

#If Win64 Then
Public Sub Test_RightShiftLongLong8000000000000000()
    Test_RightShiftLongLongX ((-CLngLng(2 ^ 62)) * 2)
End Sub

Public Sub Test_RightShiftLongLongFFFFFFFFFFFFFFFF()
    Test_RightShiftLongLongX -1
End Sub
#End If

Public Sub Test_LeftRotateByte0F()
    Test_LeftRotateByteX &HF
End Sub

Public Sub Test_LeftRotateByteF0()
    Test_LeftRotateByteX &HF0
End Sub

Public Sub Test_LeftRotateInteger00FF()
    Test_LeftRotateIntegerX &HFF
End Sub

Public Sub Test_LeftRotateIntegerFF00()
    Test_LeftRotateIntegerX &HFF00
End Sub

Public Sub Test_LeftRotateLong0000FFFF()
    Test_LeftRotateLongX &HFFFF&
End Sub

Public Sub Test_LeftRotateLongFFFF0000()
    Test_LeftRotateLongX &HFFFF0000
End Sub

#If Win64 Then
Public Sub Test_LeftRotateLongLong00000000FFFFFFFF()
    Test_LeftRotateLongLongX &HFFFFFFFF^
End Sub

Public Sub Test_LeftRotateLongLongFFFFFFFF00000000()
    Test_LeftRotateLongLongX Not &HFFFFFFFF^
End Sub
#End If

Public Sub Test_RightRotateByte0F()
    Test_RightRotateByteX &HF
End Sub

Public Sub Test_RightRotateByteF0()
    Test_RightRotateByteX &HF0
End Sub

Public Sub Test_RightRotateInteger00FF()
    Test_RightRotateIntegerX &HFF
End Sub

Public Sub Test_RightRotateIntegerFF00()
    Test_RightRotateIntegerX &HFF00
End Sub

Public Sub Test_RightRotateLong0000FFFF()
    Test_RightRotateLongX &HFFFF&
End Sub

Public Sub Test_RightRotateLongFFFF0000()
    Test_RightRotateLongX &HFFFF0000
End Sub

#If Win64 Then
Public Sub Test_RightRotateLongLong00000000FFFFFFFF()
    Test_RightRotateLongLongX &HFFFFFFFF^
End Sub

Public Sub Test_RightRotateLongFFFFFFFF00000000()
    Test_RightRotateLongLongX Not &HFFFFFFFF^
End Sub
#End If

'
' --- Test X ---
'

Public Sub Test_LeftShiftByteX(ByVal Value)
    Dim Count
    For Count = -1 To 8
        Test_LeftShiftByte_Core Value, Count
    Next
End Sub

Public Sub Test_LeftShiftIntegerX(ByVal Value)
    Dim Count
    For Count = -1 To 16
        Test_LeftShiftInteger_Core Value, Count
    Next
End Sub

Public Sub Test_LeftShiftLongX(ByVal Value)
    Dim Count
    For Count = -1 To 32
        Test_LeftShiftLong_Core Value, Count
    Next
End Sub

#If Win64 Then
Public Sub Test_LeftShiftLongLongX(ByVal Value)
    Dim Count
    For Count = -1 To 64
        Test_LeftShiftLongLong_Core Value, Count
    Next
End Sub
#End If

Public Sub Test_RightArithmeticShiftByteX(ByVal Value)
    Dim Count
    For Count = -1 To 8
        Test_RightArithmeticShiftByte_Core Value, Count
    Next
End Sub

Public Sub Test_RightArithmeticShiftIntegerX(ByVal Value)
    Dim Count
    For Count = -1 To 16
        Test_RightArithmeticShiftInteger_Core Value, Count
    Next
End Sub

Public Sub Test_RightArithmeticShiftLongX(ByVal Value)
    Dim Count
    For Count = -1 To 32
        Test_RightArithmeticShiftLong_Core Value, Count
    Next
End Sub

#If Win64 Then
Public Sub Test_RightArithmeticShiftLongLongX(ByVal Value)
    Dim Count
    For Count = -1 To 64
        Test_RightArithmeticShiftLongLong_Core Value, Count
    Next
End Sub
#End If

Public Sub Test_RightShiftByteX(ByVal Value)
    Dim Count
    For Count = -1 To 8
        Test_RightShiftByte_Core Value, Count
    Next
End Sub

Public Sub Test_RightShiftIntegerX(ByVal Value)
    Dim Count
    For Count = -1 To 16
        Test_RightShiftInteger_Core Value, Count
    Next
End Sub

Public Sub Test_RightShiftLongX(ByVal Value)
    Dim Count
    For Count = -1 To 32
        Test_RightShiftLong_Core Value, Count
    Next
End Sub

#If Win64 Then
Public Sub Test_RightShiftLongLongX(ByVal Value)
    Dim Count
    For Count = -1 To 64
        Test_RightShiftLongLong_Core Value, Count
    Next
End Sub
#End If

Public Sub Test_LeftRotateByteX(ByVal Value)
    Dim Count
    For Count = -1 To 8
        Test_LeftRotateByte_Core Value, Count
    Next
End Sub

Public Sub Test_LeftRotateIntegerX(ByVal Value)
    Dim Count
    For Count = -1 To 16
        Test_LeftRotateInteger_Core Value, Count
    Next
End Sub

Public Sub Test_LeftRotateLongX(ByVal Value)
    Dim Count
    For Count = -1 To 32
        Test_LeftRotateLong_Core Value, Count
    Next
End Sub

#If Win64 Then
Public Sub Test_LeftRotateLongLongX(ByVal Value)
    Dim Count
    For Count = -1 To 64
        Test_LeftRotateLongLong_Core Value, Count
    Next
End Sub
#End If

Public Sub Test_RightRotateByteX(ByVal Value)
    Dim Count
    For Count = -1 To 8
        Test_RightRotateByte_Core Value, Count
    Next
End Sub

Public Sub Test_RightRotateIntegerX(ByVal Value)
    Dim Count
    For Count = -1 To 16
        Test_RightRotateInteger_Core Value, Count
    Next
End Sub

Public Sub Test_RightRotateLongX(ByVal Value)
    Dim Count
    For Count = -1 To 32
        Test_RightRotateLong_Core Value, Count
    Next
End Sub

#If Win64 Then
Public Sub Test_RightRotateLongLongX(ByVal Value)
    Dim Count
    For Count = -1 To 64
        Test_RightRotateLongLong_Core Value, Count
    Next
End Sub
#End If

'
' --- Test Core ---
'

Public Sub Test_LeftShiftByte_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftShiftByte(Value, Count)
    DebugPrintBinOpByte Value, "<<", Count, Result
End Sub

Public Sub Test_LeftShiftInteger_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftShiftInteger(Value, Count)
    DebugPrintBinOpInteger Value, "<<", Count, Result
End Sub

Public Sub Test_LeftShiftLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftShiftLong(Value, Count)
    DebugPrintBinOpLong Value, "<<", Count, Result
End Sub

#If Win64 Then
Public Sub Test_LeftShiftLongLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftShiftLongLong(Value, Count)
    DebugPrintBinOpLongLong Value, "<<", Count, Result
End Sub
#End If

Public Sub Test_RightArithmeticShiftByte_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightArithmeticShiftByte(Value, Count)
    DebugPrintBinOpByte Value, ">>", Count, Result
End Sub

Public Sub Test_RightArithmeticShiftInteger_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightArithmeticShiftInteger(Value, Count)
    DebugPrintBinOpInteger Value, ">>", Count, Result
End Sub

Public Sub Test_RightArithmeticShiftLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightArithmeticShiftLong(Value, Count)
    DebugPrintBinOpLong Value, ">>", Count, Result
End Sub

#If Win64 Then
Public Sub Test_RightArithmeticShiftLongLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightArithmeticShiftLongLong(Value, Count)
    DebugPrintBinOpLongLong Value, ">>", Count, Result
End Sub
#End If

Public Sub Test_RightShiftByte_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightShiftByte(Value, Count)
    DebugPrintBinOpByte Value, ">>", Count, Result
End Sub

Public Sub Test_RightShiftInteger_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightShiftInteger(Value, Count)
    DebugPrintBinOpInteger Value, ">>", Count, Result
End Sub

Public Sub Test_RightShiftLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightShiftLong(Value, Count)
    DebugPrintBinOpLong Value, ">>", Count, Result
End Sub

#If Win64 Then
Public Sub Test_RightShiftLongLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightShiftLongLong(Value, Count)
    DebugPrintBinOpLongLong Value, ">>", Count, Result
End Sub
#End If

Public Sub Test_LeftRotateByte_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftRotateByte(Value, Count)
    DebugPrintBinOpByte Value, "lrot", Count, Result
End Sub

Public Sub Test_LeftRotateInteger_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftRotateInteger(Value, Count)
    DebugPrintBinOpInteger Value, "lrot", Count, Result
End Sub

Public Sub Test_LeftRotateLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftRotateLong(Value, Count)
    DebugPrintBinOpLong Value, "lrot", Count, Result
End Sub

#If Win64 Then
Public Sub Test_LeftRotateLongLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = LeftRotateLongLong(Value, Count)
    DebugPrintBinOpLongLong Value, "lrot", Count, Result
End Sub
#End If

Public Sub Test_RightRotateByte_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightRotateByte(Value, Count)
    DebugPrintBinOpByte Value, "rrot", Count, Result
End Sub

Public Sub Test_RightRotateInteger_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightRotateInteger(Value, Count)
    DebugPrintBinOpInteger Value, "rrot", Count, Result
End Sub

Public Sub Test_RightRotateLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightRotateLong(Value, Count)
    DebugPrintBinOpLong Value, "rrot", Count, Result
End Sub

#If Win64 Then
Public Sub Test_RightRotateLongLong_Core( _
    ByVal Value, ByVal Count)
    
    Dim Result
    Result = RightRotateLongLong(Value, Count)
    DebugPrintBinOpLongLong Value, "rrot", Count, Result
End Sub
#End If

Public Sub DebugPrintBinOpByte( _
    Value, Op, Count, Result)
    
    Debug_Print _
        GetBinStringFromByte(Value, True) & " " & _
        Op & " " & CStr(Count) & " = " & _
        GetBinStringFromByte(Result, True) & " " & CStr(Result)
End Sub

Public Sub DebugPrintBinOpInteger( _
    Value, Op, Count, Result)
    
    Debug_Print _
        GetBinStringFromInteger(Value, True) & " " & _
        Op & " " & CStr(Count) & " = " & _
        GetBinStringFromInteger(Result, True) & " " & CStr(Result)
End Sub

Public Sub DebugPrintBinOpLong( _
    Value, Op, Count, Result)
    
    Debug_Print _
        GetBinStringFromLong(Value, True) & " " & _
        Op & " " & CStr(Count) & " = " & _
        GetBinStringFromLong(Result, True) & " " & CStr(Result)
End Sub

#If Win64 Then
Public Sub DebugPrintBinOpLongLong( _
    Value, Op, Count, Result)
    
    Debug_Print _
        GetBinStringFromLongLong(Value, True) & " " & _
        Op & " " & CStr(Count) & " = " & _
        GetBinStringFromLongLong(Result, True) & " " & CStr(Result)
End Sub
#End If

Attribute VB_Name = "BitShift"
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
' === Bitwise Operations ===
'

'
' Left Arithmetic Shift
' Left Logical Shift
'

Public Function LeftShiftByte( _
    ByVal Value As Byte, _
    ByVal Count As Integer) As Byte
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 8) + 8
    ElseIf Count >= 8 Then
        Cnt = Count Mod 8
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftShiftByte = Value
        Exit Function
    End If
    
    Dim BitMask
    BitMask = Array(&H7F, &H3F, &H1F, &HF, &H7, &H3, &H1)
    
    LeftShiftByte = (Value And BitMask(Cnt - 1)) * 2 ^ Cnt
End Function

Public Function LeftShiftInteger( _
    ByVal Value As Integer, _
    ByVal Count As Integer) As Integer
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 16) + 16
    ElseIf Count >= 16 Then
        Cnt = Count Mod 16
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftShiftInteger = Value
        Exit Function
    End If
    
    Dim BitMask1
    BitMask1 = Array(&H3FFF, &H1FFF, _
        &HFFF, &H7FF, &H3FF, &H1FF, _
        &HFF, &H7F, &H3F, &H1F, _
        &HF, &H7, &H3, &H1, &H0)
        
    Dim Temp As Integer
    Temp = (Value And BitMask1(Cnt - 1)) * 2 ^ Cnt
    
    Dim BitMask2
    BitMask2 = Array(&H4000, &H2000, _
        &H1000, &H800, &H400, &H200, _
        &H100, &H80, &H40, &H20, _
        &H10, &H8, &H4, &H2, &H1)
        
    If (Value And BitMask2(Cnt - 1)) = BitMask2(Cnt - 1) Then
        Temp = Temp Or &H8000
    End If
    
    LeftShiftInteger = Temp
End Function

Public Function LeftShiftLong( _
    ByVal Value As Long, _
    ByVal Count As Integer) As Long
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 32) + 32
    ElseIf Count >= 32 Then
        Cnt = Count Mod 32
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftShiftLong = Value
        Exit Function
    End If
    
    Dim BitMask1
    BitMask1 = Array(&H3FFFFFFF, &H1FFFFFFF, _
        &HFFFFFFF, &H7FFFFFF, &H3FFFFFF, &H1FFFFFF, _
        &HFFFFFF, &H7FFFFF, &H3FFFFF, &H1FFFFF, _
        &HFFFFF, &H7FFFF, &H3FFFF, &H1FFFF, _
        &HFFFF&, &H7FFF, &H3FFF, &H1FFF, _
        &HFFF, &H7FF, &H3FF, &H1FF, _
        &HFF, &H7F, &H3F, &H1F, _
        &HF, &H7, &H3, &H1, &H0)
        
    Dim Temp As Long
    Temp = (Value And BitMask1(Cnt - 1)) * 2 ^ Cnt
    
    Dim BitMask2
    BitMask2 = Array(&H40000000, &H20000000, _
        &H10000000, &H8000000, &H4000000, &H2000000, _
        &H1000000, &H800000, &H400000, &H200000, _
        &H100000, &H80000, &H40000, &H20000, _
        &H10000, &H8000&, &H4000, &H2000, _
        &H1000, &H800, &H400, &H200, _
        &H100, &H80, &H40, &H20, _
        &H10, &H8, &H4, &H2, &H1)
        
    If (Value And BitMask2(Cnt - 1)) = BitMask2(Cnt - 1) Then
        Temp = Temp Or &H80000000
    End If
    
    LeftShiftLong = Temp
End Function

#If Win64 Then
Public Function LeftShiftLongLong( _
    ByVal Value As LongLong, _
    ByVal Count As Integer) As LongLong
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 64) + 64
    ElseIf Count >= 64 Then
        Cnt = Count Mod 64
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftShiftLongLong = Value
        Exit Function
    End If
    
    Dim BitMask1 As LongLong
    Dim Index1 As Integer
    For Index1 = 0 To 64 - Cnt - 1 - 1
        BitMask1 = BitMask1 Or (2 ^ Index1)
    Next
        
    Dim Temp As LongLong
    If Cnt < 63 Then
        Temp = CLngLng(Value And BitMask1) * CLngLng(2 ^ Cnt)
    End If
    
    Dim BitMask2 As LongLong
    BitMask2 = 2 ^ (64 - Cnt - 1)
    
    If (Value And BitMask2) = BitMask2 Then
        Temp = Temp Or ((-(2 ^ 62)) * 2)
    End If
    
    LeftShiftLongLong = Temp
End Function
#End If

'
' Right Arithmetic Shift
'

Public Function RightArithmeticShiftByte( _
    ByVal Value As Byte, _
    ByVal Count As Integer) As Byte
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 8) + 8
    ElseIf Count >= 8 Then
        Cnt = Count Mod 8
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightArithmeticShiftByte = Value
        Exit Function
    End If
    
    Dim Temp As Byte
    If Cnt < 7 Then
        Temp = Value \ 2 ^ Cnt
    Else
        Temp = 0
    End If
    
    If (Value And &H80) = &H80 Then
        Dim BitPattern
        BitPattern = Array(&HC0, &HE0, &HF0, &HF8, &HFC, &HFE, &HFF)
        
        Temp = Temp Or BitPattern(Cnt - 1)
    End If
    
    RightArithmeticShiftByte = Temp
End Function

Public Function RightArithmeticShiftInteger( _
    ByVal Value As Integer, _
    ByVal Count As Integer) As Integer
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 16) + 16
    ElseIf Count >= 16 Then
        Cnt = Count Mod 16
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightArithmeticShiftInteger = Value
        Exit Function
    End If
    
    Dim Temp As Integer
    If Cnt < 15 Then
        Temp = (Value And &H7FFF) \ 2 ^ Cnt
    Else
        Temp = 0
    End If
    
    If (Value And &H8000) = &H8000 Then
        Dim BitPattern
        BitPattern = Array(&HC000, &HE000, &HF000, _
            &HF800, &HFC00, &HFE00, &HFF00, _
            &HFF80, &HFFC0, &HFFE0, &HFFF0, _
            &HFFF8, &HFFFC, &HFFFE, &HFFFF)
        
        Temp = Temp Or BitPattern(Cnt - 1)
    End If
    
    RightArithmeticShiftInteger = Temp
End Function

Public Function RightArithmeticShiftLong( _
    ByVal Value As Long, _
    ByVal Count As Integer) As Long
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 32) + 32
    ElseIf Count >= 32 Then
        Cnt = Count Mod 32
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightArithmeticShiftLong = Value
        Exit Function
    End If
    
    Dim Temp As Long
    If Cnt < 31 Then
        Temp = (Value And &H7FFFFFFF) \ 2 ^ Cnt
    Else
        Temp = 0
    End If
    
    If (Value And &H80000000) = &H80000000 Then
        Dim BitPattern
        BitPattern = Array(&HC0000000, &HE0000000, &HF0000000, _
            &HF8000000, &HFC000000, &HFE000000, &HFF000000, _
            &HFF800000, &HFFC00000, &HFFE00000, &HFFF00000, _
            &HFFF80000, &HFFFC0000, &HFFFE0000, &HFFFF0000, _
            &HFFFF8000, &HFFFFC000, &HFFFFE000, &HFFFFF000, _
            &HFFFFF800, &HFFFFFC00, &HFFFFFE00, &HFFFFFF00, _
            &HFFFFFF80, &HFFFFFFC0, &HFFFFFFE0, &HFFFFFFF0, _
            &HFFFFFFF8, &HFFFFFFFC, &HFFFFFFFE, &HFFFFFFFF)
        
        Temp = Temp Or BitPattern(Cnt - 1)
    End If
    
    RightArithmeticShiftLong = Temp
End Function

#If Win64 Then
Public Function RightArithmeticShiftLongLong( _
    ByVal Value As LongLong, _
    ByVal Count As Integer) As LongLong
    
    If Value = 0 Then
        RightArithmeticShiftLongLong = 0
        Exit Function
    End If
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 64) + 64
    ElseIf Count >= 64 Then
        Cnt = Count Mod 64
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightArithmeticShiftLongLong = Value
        Exit Function
    End If
    
    If Value > 0 Then
        If Cnt < 63 Then
            RightArithmeticShiftLongLong = Value \ 2 ^ Cnt
            Exit Function
        Else ' Cnt = 63
            RightArithmeticShiftLongLong = 0
            Exit Function
        End If
    Else ' Value < 0
        If Cnt < 63 Then
            RightArithmeticShiftLongLong = Not ((Not Value) \ 2 ^ Cnt)
            Exit Function
        Else ' Cnt = 63
            RightArithmeticShiftLongLong = -1
            Exit Function
        End If
    End If
End Function
#End If

'
' Right Logical Shift
'

Public Function RightShiftByte( _
    ByVal Value As Byte, _
    ByVal Count As Integer) As Byte
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 8) + 8
    ElseIf Count >= 8 Then
        Cnt = Count Mod 8
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightShiftByte = Value
        Exit Function
    End If
    
    RightShiftByte = Value \ 2 ^ Cnt
End Function

Public Function RightShiftInteger( _
    ByVal Value As Integer, _
    ByVal Count As Integer) As Integer
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 16) + 16
    ElseIf Count >= 16 Then
        Cnt = Count Mod 16
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightShiftInteger = Value
        Exit Function
    End If
    
    Dim Temp As Integer
    If Cnt < 15 Then
        Temp = (Value And &H7FFF) \ 2 ^ Cnt
    Else
        Temp = 0
    End If
    
    If (Value And &H8000) = &H8000 Then
        Dim BitPattern
        BitPattern = Array(&H4000, &H2000, &H1000, _
            &H800, &H400, &H200, &H100, _
            &H80, &H40, &H20, &H10, _
            &H8, &H4, &H2, &H1)
        
        Temp = Temp Or BitPattern(Cnt - 1)
    End If
    
    RightShiftInteger = Temp
End Function

Public Function RightShiftLong( _
    ByVal Value As Long, _
    ByVal Count As Integer) As Long
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 32) + 32
    ElseIf Count >= 32 Then
        Cnt = Count Mod 32
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightShiftLong = Value
        Exit Function
    End If
    
    Dim Temp As Long
    If Cnt < 31 Then
        Temp = (Value And &H7FFFFFFF) \ 2 ^ Cnt
    Else
        Temp = 0
    End If
    
    If (Value And &H80000000) = &H80000000 Then
        Dim BitPattern
        BitPattern = Array(&H40000000, &H20000000, &H10000000, _
            &H8000000, &H4000000, &H2000000, &H1000000, _
            &H800000, &H400000, &H200000, &H100000, _
            &H80000, &H40000, &H20000, &H10000, _
            &H8000&, &H4000, &H2000, &H1000, _
            &H800, &H400, &H200, &H100, _
            &H80, &H40, &H20, &H10, _
            &H8, &H4, &H2, &H1)
        
        Temp = Temp Or BitPattern(Cnt - 1)
    End If
    
    RightShiftLong = Temp
End Function

#If Win64 Then
Public Function RightShiftLongLong( _
    ByVal Value As LongLong, _
    ByVal Count As Integer) As LongLong
    
    If Value = 0 Then
        RightShiftLongLong = 0
        Exit Function
    End If
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 64) + 64
    ElseIf Count >= 64 Then
        Cnt = Count Mod 64
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightShiftLongLong = Value
        Exit Function
    End If
    
    If Value > 0 Then
        If Cnt < 63 Then
            RightShiftLongLong = Value \ 2 ^ Cnt
            Exit Function
        Else ' Cnt = 63
            RightShiftLongLong = 0
            Exit Function
        End If
    Else ' Value < 0
        If Cnt < 63 Then
            RightShiftLongLong = _
                (Not ((Not Value) \ 2 ^ Cnt)) And _
                (Not ((-CLngLng(2 ^ (64 - Cnt - 1))) * 2))
            Exit Function
        Else ' Cnt = 63
            RightShiftLongLong = 1
            Exit Function
        End If
    End If
End Function
#End If

'
' Left Circular Shift (Left Rotate)
'

Public Function LeftRotateByte( _
    ByVal Value As Byte, _
    ByVal Count As Integer) As Byte
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 8) + 8
    ElseIf Count >= 8 Then
        Cnt = Count Mod 8
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftRotateByte = Value
        Exit Function
    End If
    
    LeftRotateByte = _
        LeftShiftByte(Value, Cnt) Or RightShiftByte(Value, 8 - Cnt)
End Function

Public Function LeftRotateInteger( _
    ByVal Value As Integer, _
    ByVal Count As Integer) As Integer
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 16) + 16
    ElseIf Count >= 16 Then
        Cnt = Count Mod 16
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftRotateInteger = Value
        Exit Function
    End If
    
    LeftRotateInteger = _
        LeftShiftInteger(Value, Cnt) Or RightShiftInteger(Value, 16 - Cnt)
End Function

Public Function LeftRotateLong( _
    ByVal Value As Long, _
    ByVal Count As Integer) As Long
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 32) + 32
    ElseIf Count >= 32 Then
        Cnt = Count Mod 32
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftRotateLong = Value
        Exit Function
    End If
    
    LeftRotateLong = _
        LeftShiftLong(Value, Cnt) Or RightShiftLong(Value, 32 - Cnt)
End Function

#If Win64 Then
Public Function LeftRotateLongLong( _
    ByVal Value As LongLong, _
    ByVal Count As Integer) As LongLong
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 64) + 64
    ElseIf Count >= 64 Then
        Cnt = Count Mod 64
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        LeftRotateLongLong = Value
        Exit Function
    End If
    
    LeftRotateLongLong = _
        LeftShiftLongLong(Value, Cnt) Or RightShiftLongLong(Value, 64 - Cnt)
End Function
#End If

'
' Right Circular Shift (Right Rotate)
'

Public Function RightRotateByte( _
    ByVal Value As Byte, _
    ByVal Count As Integer) As Byte
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 8) + 8
    ElseIf Count >= 8 Then
        Cnt = Count Mod 8
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightRotateByte = Value
        Exit Function
    End If
    
    RightRotateByte = _
        RightShiftByte(Value, Cnt) Or LeftShiftByte(Value, 8 - Cnt)
End Function

Public Function RightRotateInteger( _
    ByVal Value As Integer, _
    ByVal Count As Integer) As Integer
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 16) + 16
    ElseIf Count >= 16 Then
        Cnt = Count Mod 16
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightRotateInteger = Value
        Exit Function
    End If
    
    RightRotateInteger = _
        RightShiftInteger(Value, Cnt) Or LeftShiftInteger(Value, 16 - Cnt)
End Function

Public Function RightRotateLong( _
    ByVal Value As Long, _
    ByVal Count As Integer) As Long
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 32) + 32
    ElseIf Count >= 32 Then
        Cnt = Count Mod 32
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightRotateLong = Value
        Exit Function
    End If
    
    RightRotateLong = _
        RightShiftLong(Value, Cnt) Or LeftShiftLong(Value, 32 - Cnt)
End Function

#If Win64 Then
Public Function RightRotateLongLong( _
    ByVal Value As LongLong, _
    ByVal Count As Integer) As LongLong
    
    Dim Cnt As Integer
    If Count < 0 Then
        Cnt = (Count Mod 64) + 64
    ElseIf Count >= 64 Then
        Cnt = Count Mod 64
    Else
        Cnt = Count
    End If
    
    If Cnt = 0 Then
        RightRotateLongLong = Value
        Exit Function
    End If
    
    RightRotateLongLong = _
        RightShiftLongLong(Value, Cnt) Or LeftShiftLongLong(Value, 64 - Cnt)
End Function
#End If

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
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        LeftShiftByte = RightShiftByte(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        LeftShiftByte = Value
        Exit Function
    ElseIf Count > 7 Then
        LeftShiftByte = 0
        Exit Function
    End If
    
    Dim BitMask
    BitMask = Array(&H7F, &H3F, &H1F, &HF, &H7, &H3, &H1)
    
    LeftShiftByte = (Value And BitMask(Count - 1)) * 2 ^ Count
End Function

Public Function LeftShiftInteger( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        LeftShiftInteger = RightShiftInteger(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        LeftShiftInteger = Value
        Exit Function
    ElseIf Count > 15 Then
        LeftShiftInteger = 0
        Exit Function
    End If
    
    Dim BitMask1
    BitMask1 = Array(&H3FFF, &H1FFF, &HFFF, &H7FF, &H3FF, &H1FF, _
        &HFF, &H7F, &H3F, &H1F, &HF, &H7, &H3, &H1, &H0)
        
    Dim LeftShiftIntegerTemp
    LeftShiftIntegerTemp = (Value And BitMask1(Count - 1)) * 2 ^ Count
    
    Dim BitMask2
    BitMask2 = Array(&H4000, &H2000, &H1000, &H800, &H400, &H200, _
        &H100, &H80, &H40, &H20, &H10, &H8, &H4, &H2, &H1)
        
    If (Value And BitMask2(Count - 1)) = BitMask2(Count - 1) Then
        LeftShiftIntegerTemp = LeftShiftIntegerTemp Or &H8000
    End If
    
    LeftShiftInteger = LeftShiftIntegerTemp
End Function

Public Function LeftShiftLong( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        LeftShiftLong = RightShiftLong(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        LeftShiftLong = Value
        Exit Function
    ElseIf Count > 31 Then
        LeftShiftLong = 0
        Exit Function
    End If
    
    Dim BitMask1
    BitMask1 = Array(&H3FFFFFFF, &H1FFFFFFF, _
        &HFFFFFFF, &H7FFFFFF, &H3FFFFFF, &H1FFFFFF, _
        &HFFFFFF, &H7FFFFF, &H3FFFFF, &H1FFFFF, _
        &HFFFFF, &H7FFFF, &H3FFFF, &H1FFFF, _
        &HFFFF&, &H7FFF, &H3FFF, &H1FFF, &HFFF, &H7FF, &H3FF, &H1FF, _
        &HFF, &H7F, &H3F, &H1F, &HF, &H7, &H3, &H1, &H0)
        
    Dim LeftShiftLongTemp
    LeftShiftLongTemp = (Value And BitMask1(Count - 1)) * 2 ^ Count
    
    Dim BitMask2
    BitMask2 = Array(&H40000000, &H20000000, _
        &H10000000, &H8000000, &H4000000, &H2000000, _
        &H1000000, &H800000, &H400000, &H200000, _
        &H100000, &H80000, &H40000, &H20000, _
        &H10000, &H8000&, &H4000, &H2000, &H1000, &H800, &H400, &H200, _
        &H100, &H80, &H40, &H20, &H10, &H8, &H4, &H2, &H1)
        
    If (Value And BitMask2(Count - 1)) = BitMask2(Count - 1) Then
        LeftShiftLongTemp = LeftShiftLongTemp Or &H80000000
    End If
    
    LeftShiftLong = LeftShiftLongTemp
End Function

'
' Right Arithmetic Shift - To Do
' Right Logical Shift
'

Public Function RightShiftByte( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        RightShiftByte = LeftShiftByte(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        RightShiftByte = Value
        Exit Function
    ElseIf Count > 7 Then
        RightShiftByte = 0
        Exit Function
    End If
    
    RightShiftByte = Value \ 2 ^ Count
End Function

Public Function RightShiftInteger( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        RightShiftInteger = LeftShiftInteger(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        RightShiftInteger = Value
        Exit Function
    ElseIf Count > 15 Then
        RightShiftInteger = 0
        Exit Function
    End If
    
    Dim RightShiftIntegerTemp
    If Count < 15 Then
        RightShiftIntegerTemp = (Value And &H7FFF) \ 2 ^ Count
    End If
    
    Dim BitPattern
    BitPattern = Array(&H4000, &H2000, &H1000, &H800, &H400, &H200, _
        &H100, &H80, &H40, &H20, &H10, &H8, &H4, &H2, &H1)
        
    If (Value And &H8000) = &H8000 Then
        RightShiftIntegerTemp = RightShiftIntegerTemp Or BitPattern(Count - 1)
    End If
    
    RightShiftInteger = RightShiftIntegerTemp
End Function

Public Function RightShiftLong( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        RightShiftLong = LeftShiftInteger(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        RightShiftLong = Value
        Exit Function
    ElseIf Count > 31 Then
        RightShiftLong = 0
        Exit Function
    End If
    
    Dim RightShiftLongTemp
    If Count < 31 Then
        RightShiftLongTemp = (Value And &H7FFFFFFF) \ 2 ^ Count
    End If
    
    Dim BitPattern
    BitPattern = Array(&H40000000, &H20000000, _
        &H10000000, &H8000000, &H4000000, &H2000000, _
        &H1000000, &H800000, &H400000, &H200000, _
        &H100000, &H80000, &H40000, &H20000, _
        &H10000, &H8000&, &H4000, &H2000, &H1000, &H800, &H400, &H200, _
        &H100, &H80, &H40, &H20, &H10, &H8, &H4, &H2, &H1)
        
    If (Value And &H80000000) = &H80000000 Then
        RightShiftLongTemp = RightShiftLongTemp Or BitPattern(Count - 1)
    End If
    
    RightShiftLong = RightShiftLongTemp
End Function

'
' Left Circular Shift (Left Rotate)
'

Public Function LeftRotateByte( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        LeftRotateByte = RightRotateByte(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        LeftRotateByte = Value
        Exit Function
    ElseIf Count > 7 Then
        LeftRotateByte = LeftRotateByte(Value, Count Mod 8)
        Exit Function
    End If
    
    LeftRotateByte = _
        LeftShiftByte(Value, Count) Or RightShiftByte(Value, 8 - Count)
End Function

Public Function LeftRotateInteger( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        LeftRotateInteger = RightRotateInteger(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        LeftRotateInteger = Value
        Exit Function
    ElseIf Count > 15 Then
        LeftRotateInteger = LeftRotateInteger(Value, Count Mod 16)
        Exit Function
    End If
    
    LeftRotateInteger = _
        LeftShiftInteger(Value, Count) Or RightShiftInteger(Value, 16 - Count)
End Function

Public Function LeftRotateLong( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        LeftRotateLong = RightRotateLong(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        LeftRotateLong = Value
        Exit Function
    ElseIf Count > 31 Then
        LeftRotateLong = LeftRotateLong(Value, Count Mod 32)
        Exit Function
    End If
    
    LeftRotateLong = _
        LeftShiftLong(Value, Count) Or RightShiftLong(Value, 32 - Count)
End Function

'
' Right Circular Shift (Right Rotate)
'

Public Function RightRotateByte( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        RightRotateByte = LeftRotateByte(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        RightRotateByte = Value
        Exit Function
    ElseIf Count > 7 Then
        RightRotateByte = RightRotateByte(Value, Count Mod 8)
        Exit Function
    End If
    
    RightRotateByte = _
        RightShiftByte(Value, Count) Or LeftShiftByte(Value, 8 - Count)
End Function

Public Function RightRotateInteger( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        RightRotateInteger = LeftRotateInteger(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        RightRotateInteger = Value
        Exit Function
    ElseIf Count > 15 Then
        RightRotateInteger = RightRotateInteger(Value, Count Mod 16)
        Exit Function
    End If
    
    RightRotateInteger = _
        RightShiftInteger(Value, Count) Or LeftShiftInteger(Value, 16 - Count)
End Function

Public Function RightRotateLong( _
    ByVal Value, _
    ByVal Count)
    
    If Count < 0 Then
        RightRotateLong = LeftRotateLong(Value, Abs(Count))
        Exit Function
    ElseIf Count = 0 Then
        RightRotateLong = Value
        Exit Function
    ElseIf Count > 31 Then
        RightRotateLong = RightRotateLong(Value, Count Mod 32)
        Exit Function
    End If
    
    RightRotateLong = _
        RightShiftLong(Value, Count) Or LeftShiftLong(Value, 32 - Count)
End Function

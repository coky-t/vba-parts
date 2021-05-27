Attribute VB_Name = "ByteString"
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

Public Function GetStringBFromByte(Value)
    GetStringBFromByte = ChrB(Value)
End Function

Public Function GetStringB_LEFromInteger(Value)
    GetStringB_LEFromInteger = _
        ChrB(Value And &HFF) & _
        ChrB(RightShiftInteger(Value, 8) And &HFF)
End Function

Public Function GetStringB_BEFromInteger(Value)
    GetStringB_BEFromInteger = _
        ChrB(RightShiftInteger(Value, 8) And &HFF) & _
        ChrB(Value And &HFF)
End Function

Public Function GetStringB_LEFromLong(Value)
    GetStringB_LEFromLong = _
        ChrB(Value And &HFF) & _
        ChrB(RightShiftLong(Value, 8) And &HFF) & _
        ChrB(RightShiftLong(Value, 16) And &HFF) & _
        ChrB(RightShiftLong(Value, 24) And &HFF)
End Function

Public Function GetStringB_BEFromLong(Value)
    GetStringB_BEFromLong = _
        ChrB(RightShiftLong(Value, 24) And &HFF) & _
        ChrB(RightShiftLong(Value, 16) And &HFF) & _
        ChrB(RightShiftLong(Value, 8) And &HFF) & _
        ChrB(Value And &HFF)
End Function

Public Function GetByteFromStringB( _
    StrB, Pos)
    
    GetByteFromStringB = AscB(MidB(StrB, Pos, 1))
End Function

Public Function GetIntegerFromStringB_LE( _
    StrB_LE, Pos)
    
    GetIntegerFromStringB_LE = CInt( _
        AscB(MidB(StrB_LE, Pos, 1)) Or _
        LeftShiftInteger(AscB(MidB(StrB_LE, Pos + 1, 1)), 8))
End Function

Public Function GetIntegerFromStringB_BE( _
    StrB_BE, Pos)
    
    GetIntegerFromStringB_BE = CInt( _
        AscB(MidB(StrB_BE, Pos + 1, 1)) Or _
        LeftShiftInteger(AscB(MidB(StrB_BE, Pos, 1)), 8))
End Function

Public Function GetLongFromStringB_LE( _
    StrB_LE, Pos)
    
    GetLongFromStringB_LE = AscB(MidB(StrB_LE, Pos, 1)) Or _
        LeftShiftLong(AscB(MidB(StrB_LE, Pos + 1, 1)), 8) Or _
        LeftShiftLong(AscB(MidB(StrB_LE, Pos + 2, 1)), 16) Or _
        LeftShiftLong(AscB(MidB(StrB_LE, Pos + 3, 1)), 24)
End Function

Public Function GetLongFromStringB_BE( _
    StrB_BE, Pos)
    
    GetLongFromStringB_BE = AscB(MidB(StrB_BE, Pos + 3, 1)) Or _
        LeftShiftLong(AscB(MidB(StrB_BE, Pos + 2, 1)), 8) Or _
        LeftShiftLong(AscB(MidB(StrB_BE, Pos + 1, 1)), 16) Or _
        LeftShiftLong(AscB(MidB(StrB_BE, Pos, 1)), 24)
End Function

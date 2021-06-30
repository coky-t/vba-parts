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

Public Function GetStringBFromByte(Value As Byte) As String
    GetStringBFromByte = ChrB(Value)
End Function

Public Function GetStringB_LEFromInteger(Value As Integer) As String
    GetStringB_LEFromInteger = _
        ChrB(Value And &HFF) & _
        ChrB(RightShiftInteger(Value, 8) And &HFF)
End Function

Public Function GetStringB_BEFromInteger(Value As Integer) As String
    GetStringB_BEFromInteger = _
        ChrB(RightShiftInteger(Value, 8) And &HFF) & _
        ChrB(Value And &HFF)
End Function

Public Function GetStringB_LEFromLong(Value As Long) As String
    GetStringB_LEFromLong = _
        ChrB(Value And &HFF) & _
        ChrB(RightShiftLong(Value, 8) And &HFF) & _
        ChrB(RightShiftLong(Value, 16) And &HFF) & _
        ChrB(RightShiftLong(Value, 24) And &HFF)
End Function

Public Function GetStringB_BEFromLong(Value As Long) As String
    GetStringB_BEFromLong = _
        ChrB(RightShiftLong(Value, 24) And &HFF) & _
        ChrB(RightShiftLong(Value, 16) And &HFF) & _
        ChrB(RightShiftLong(Value, 8) And &HFF) & _
        ChrB(Value And &HFF)
End Function

#If Win64 Then
Public Function GetStringB_LEFromLongLong(Value As LongLong) As String
    GetStringB_LEFromLongLong = _
        ChrB(CByte(Value And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 8) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 16) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 24) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 32) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 40) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 48) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 56) And &HFF))
End Function

Public Function GetStringB_BEFromLongLong(Value As LongLong) As String
    GetStringB_BEFromLongLong = _
        ChrB(CByte(RightShiftLongLong(Value, 56) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 48) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 40) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 32) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 24) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 16) And &HFF)) & _
        ChrB(CByte(RightShiftLongLong(Value, 8) And &HFF)) & _
        ChrB(CByte(Value And &HFF))
End Function
#End If

Public Function GetByteFromStringB( _
    StrB As String, Optional Pos As Long = 1) As Byte
    
    GetByteFromStringB = AscB(MidB(StrB, Pos, 1))
End Function

Public Function GetIntegerFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As Integer
    
    GetIntegerFromStringB_LE = _
        AscB(MidB(StrB_LE, Pos, 1)) Or _
        LeftShiftInteger(AscB(MidB(StrB_LE, Pos + 1, 1)), 8)
End Function

Public Function GetIntegerFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As Integer
    
    GetIntegerFromStringB_BE = _
        AscB(MidB(StrB_BE, Pos + 1, 1)) Or _
        LeftShiftInteger(AscB(MidB(StrB_BE, Pos, 1)), 8)
End Function

Public Function GetLongFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As Long
    
    GetLongFromStringB_LE = AscB(MidB(StrB_LE, Pos, 1)) Or _
        LeftShiftLong(AscB(MidB(StrB_LE, Pos + 1, 1)), 8) Or _
        LeftShiftLong(AscB(MidB(StrB_LE, Pos + 2, 1)), 16) Or _
        LeftShiftLong(AscB(MidB(StrB_LE, Pos + 3, 1)), 24)
End Function

Public Function GetLongFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As Long
    
    GetLongFromStringB_BE = AscB(MidB(StrB_BE, Pos + 3, 1)) Or _
        LeftShiftLong(AscB(MidB(StrB_BE, Pos + 2, 1)), 8) Or _
        LeftShiftLong(AscB(MidB(StrB_BE, Pos + 1, 1)), 16) Or _
        LeftShiftLong(AscB(MidB(StrB_BE, Pos, 1)), 24)
End Function

#If Win64 Then
Public Function GetLongLongFromStringB_LE( _
    StrB_LE As String, Optional Pos As Long = 1) As LongLong
    
    GetLongLongFromStringB_LE = AscB(MidB(StrB_LE, Pos, 1)) Or _
        LeftShiftLongLong(AscB(MidB(StrB_LE, Pos + 1, 1)), 8) Or _
        LeftShiftLongLong(AscB(MidB(StrB_LE, Pos + 2, 1)), 16) Or _
        LeftShiftLongLong(AscB(MidB(StrB_LE, Pos + 3, 1)), 24) Or _
        LeftShiftLongLong(AscB(MidB(StrB_LE, Pos + 4, 1)), 32) Or _
        LeftShiftLongLong(AscB(MidB(StrB_LE, Pos + 5, 1)), 40) Or _
        LeftShiftLongLong(AscB(MidB(StrB_LE, Pos + 6, 1)), 48) Or _
        LeftShiftLongLong(AscB(MidB(StrB_LE, Pos + 7, 1)), 56)
End Function

Public Function GetLongLongFromStringB_BE( _
    StrB_BE As String, Optional Pos As Long = 1) As LongLong
    
    GetLongLongFromStringB_BE = AscB(MidB(StrB_BE, Pos + 7, 1)) Or _
        LeftShiftLongLong(AscB(MidB(StrB_BE, Pos + 6, 1)), 8) Or _
        LeftShiftLongLong(AscB(MidB(StrB_BE, Pos + 5, 1)), 16) Or _
        LeftShiftLongLong(AscB(MidB(StrB_BE, Pos + 4, 1)), 24) Or _
        LeftShiftLongLong(AscB(MidB(StrB_BE, Pos + 3, 1)), 32) Or _
        LeftShiftLongLong(AscB(MidB(StrB_BE, Pos + 2, 1)), 40) Or _
        LeftShiftLongLong(AscB(MidB(StrB_BE, Pos + 1, 1)), 48) Or _
        LeftShiftLongLong(AscB(MidB(StrB_BE, Pos, 1)), 56)
End Function
#End If

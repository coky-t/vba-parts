Attribute VB_Name = "Test_MADODBStream"
Option Explicit

'
' Copyright (c) 2020 Koki Takeyama
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

Public Sub Test_TextFileW()
    Test_TextFileW_Core _
        "testw.txt", _
        "WriteTextFileW" & vbNewLine, _
        "AppendTextFileW" & vbNewLine
End Sub

Public Sub Test_TextFileA()
    Test_TextFileA_Core _
        "testa.txt", _
        "WriteTextFileA" & vbNewLine, _
        "AppendTextFileA" & vbNewLine
End Sub

Public Sub Test_TextFileUTF8()
    Test_TextFileUTF8_Core _
        "testutf8.txt", _
        "WriteTextFileUTF8" & vbNewLine, _
        "AppendTextFileUTF8" & vbNewLine
End Sub

Public Sub Test_TextFileUTF8_withoutBOM()
    Test_TextFileUTF8_withoutBOM_Core _
        "testutf8-no-bom.txt", _
        "WriteTextFileUTF8 (w/o BOM)" & vbNewLine, _
        "AppendTextFileUTF8 (w/o BOM)" & vbNewLine
End Sub

Public Sub Test_BinaryFile()
    Test_BinaryFile_Core _
        "testb.dat", _
        GetTestStringB(), _
        GetTestStringB()
End Sub

Public Sub Test_GetBinaryGetTextW()
    Test_GetBinaryGetTextW_Core "abcdefghijklmnopqrstuvwxyz"
End Sub

Public Sub Test_GetBinaryGetTextA()
    Test_GetBinaryGetTextA_Core "abcdefghijklmnopqrstuvwxyz"
End Sub

Public Sub Test_GetBinaryGetTextUTF8()
    Test_GetBinaryGetTextUTF8_Core "abcdefghijklmnopqrstuvwxyz"
End Sub

'
' --- Test Core ---
'

Public Sub Test_TextFileW_Core( _
    FileName, _
    Text1, _
    Text2)
    
    Dim Text
    
    WriteTextFileW FileName, Text1
    Text = ReadTextFileW(FileName)
    Debug_Print Text
    
    AppendTextFileW FileName, Text2
    Text = ReadTextFileW(FileName)
    Debug_Print Text
End Sub

Public Sub Test_TextFileA_Core( _
    FileName, _
    Text1, _
    Text2)
    
    Dim Text
    
    WriteTextFileA FileName, Text1
    Text = ReadTextFileA(FileName)
    Debug_Print Text
    
    AppendTextFileA FileName, Text2
    Text = ReadTextFileA(FileName)
    Debug_Print Text
End Sub

Public Sub Test_TextFileUTF8_Core( _
    FileName, _
    Text1, _
    Text2)
    
    Dim Text
    
    WriteTextFileUTF8 FileName, Text1, True
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
    
    AppendTextFileUTF8 FileName, Text2, True
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
End Sub

Public Sub Test_TextFileUTF8_withoutBOM_Core( _
    FileName, _
    Text1, _
    Text2)
    
    Dim Text
    
    WriteTextFileUTF8 FileName, Text1, False
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
    
    AppendTextFileUTF8 FileName, Text2, False
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
End Sub

Public Sub Test_BinaryFile_Core( _
    FileName, _
    StringB1, _
    StringB2)
    
    Dim Binary
    
    WriteBinaryFileFromStringB FileName, StringB1
    Binary = ReadBinaryFile(FileName, 0)
    Debug_Print_Binary Binary
    
    AppendBinaryFileFromStringB FileName, StringB1
    Binary = ReadBinaryFile(FileName, 0)
    Debug_Print_Binary Binary
End Sub

Public Sub Test_GetBinaryGetTextW_Core(Text1)
    Dim Binary
    Binary = GetBinaryFromTextW(Text1)
    Debug_Print_Binary Binary
    
    Dim Text
    Text = GetTextWFromBinary(Binary)
    Debug_Print Text
End Sub

Public Sub Test_GetBinaryGetTextA_Core(Text1)
    Dim Binary
    Binary = GetBinaryFromTextA(Text1)
    Debug_Print_Binary Binary
    
    Dim Text
    Text = GetTextAFromBinary(Binary)
    Debug_Print Text
End Sub

Public Sub Test_GetBinaryGetTextUTF8_Core(Text1)
    Dim Binary
    Binary = GetBinaryFromTextUTF8(Text1)
    Debug_Print_Binary Binary
    
    Dim Text
    Text = GetTextUTF8FromBinary(Binary)
    Debug_Print Text
End Sub

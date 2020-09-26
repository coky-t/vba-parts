Attribute VB_Name = "MFile"
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
' Microsoft Scripting Runtime
' - Scripting.File
'

' Scripting.Tristate
Public Const Scripting_TristateUseDefault = -2
Public Const Scripting_TristateTrue = -1
Public Const Scripting_TristateFalse = 0

' Scripting.IOMode
Public Const Scripting_ForReading = 1
Public Const Scripting_ForWriting = 2
Public Const Scripting_ForAppending = 8

'
' === TextFile ===
'

'
' ReadTextFileW
' - Reads an entire file and returns the resulting string (Unicode).
'
' ReadTextFileA
' - Reads an entire file and returns the resulting string (ASCII).
'

'
' FileObject:
'   Required. The name of a File object.
'

Public Function ReadTextFileW(FileObject)
    ReadTextFileW = ReadTextFile(FileObject, Scripting_TristateTrue)
End Function

Public Function ReadTextFileA(FileObject)
    ReadTextFileA = ReadTextFile(FileObject, Scripting_TristateFalse)
End Function

Private Function ReadTextFile( _
    FileObject, _
    Format)
    
    If FileObject Is Nothing Then Exit Function
    If FileObject.Size = 0 Then Exit Function
    
    ReadTextFile = OpenAsTextStreamAndReadAll(FileObject, Format)
End Function

'
' WriteTextFileW
' - Writes a specified string (Unicode) to a file.
'
' WriteTextFileA
' - Writes a specified string (ASCII) to a file.
'
' AppendTextFileW
' - Writes a specified string (Unicode) to the end of a file.
'
' AppendTextFileA
' - Writes a specified string (ASCII) to the end of a file.
'

'
' FileObject:
'   Required. The name of a File object.
'
' Text:
'   Required. The text you want to write to the file.
'

Public Sub WriteTextFileW(FileObject, Text)
    WriteTextFile _
        FileObject, _
        Text, _
        Scripting_ForWriting, _
        Scripting_TristateTrue
End Sub

Public Sub WriteTextFileA(FileObject, Text)
    WriteTextFile _
        FileObject, _
        Text, _
        Scripting_ForWriting, _
        Scripting_TristateFalse
End Sub

Public Sub AppendTextFileW(FileObject, Text)
    WriteTextFile _
        FileObject, _
        Text, _
        Scripting_ForAppending, _
        Scripting_TristateTrue
End Sub

Public Sub AppendTextFileA(FileObject, Text)
    WriteTextFile _
        FileObject, _
        Text, _
        Scripting_ForAppending, _
        Scripting_TristateFalse
End Sub

Private Sub WriteTextFile( _
    FileObject, _
    Text, _
    IOMode, _
    Format)
    
    If FileObject Is Nothing Then Exit Sub
    If (FileObject.Attributes And 1) = 1 Then Exit Sub 'ReadOnly
    
    If IOMode = Scripting_ForReading Then Exit Sub
    
    OpenAsTextStreamAndWrite FileObject, Text, IOMode, Format
End Sub

'
' --- TextFile ---
'

'
' OpenAsTextStreamAndReadAll
' - Reads an entire file and returns the resulting string.
'

'
' FileObject:
'   Required. The name of a File object.
'
' Format:
'   Optional. One of three Tristate values used to indicate the format of
'   the opened file. If omitted, the file is opened as ASCII.
'   TristateUseDefault(-2): Opens the file by using the system default.
'   TristateTrue(-1): Opens the file as Unicode.
'   TristateFalse(0): Opens the file as ASCII.
'

Public Function OpenAsTextStreamAndReadAll( _
    FileObject, _
    Format)
    
    On Error Resume Next
    
    With FileObject.OpenAsTextStream(, Format)
        OpenAsTextStreamAndReadAll = .ReadAll
        .Close
    End With
End Function

'
' OpenAsTextStreamAndWrite
' - Writes a specified string to a file.
'

'
' FileObject:
'   Required. The name of a File object.
'
' Text:
'   Required. The text you want to write to the file.
'
' IOMode:
'   Optional. Indicates input/output mode.
'   Can be one of two constants: ForWriting(2), or ForAppending(8).
'
' Format:
'   Optional. One of three Tristate values used to indicate the format of
'   the opened file. If omitted, the file is opened as ASCII.
'   TristateUseDefault(-2): Opens the file by using the system default.
'   TristateTrue(-1): Opens the file as Unicode.
'   TristateFalse(0): Opens the file as ASCII.
'

Public Sub OpenAsTextStreamAndWrite( _
    FileObject, _
    Text, _
    IOMode, _
    Format)
    
    On Error Resume Next
    
    With FileObject.OpenAsTextStream(IOMode, Format)
        .Write (Text)
        .Close
    End With
End Sub

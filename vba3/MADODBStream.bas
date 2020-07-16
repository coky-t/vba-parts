Attribute VB_Name = "MADODBStream"
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
' Microsoft ActiveX Data Objects X.X Library
' - ADODB.Stream
'

Private ADODBStream

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
' ReadTextFileUTF8
' - Reads an entire file and returns the resulting string (UTF-8).
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'

Public Function ReadTextFileW(FileName)
    ReadTextFileW = ReadTextFile(FileName, "unicode")
End Function

Public Function ReadTextFileA(FileName)
    ReadTextFileA = ReadTextFile(FileName, "iso-8859-1")
End Function

Public Function ReadTextFileUTF8(FileName)
    ReadTextFileUTF8 = ReadTextFile(FileName, "utf-8")
End Function

Private Function ReadTextFile( _
    FileName, _
    Charset)
    
    If FileName = "" Then Exit Function
    
    ReadTextFile = LoadFromFileAndReadText(FileName, Charset)
End Function

'
' WriteTextFileW
' - Writes a specified string (Unicode) to a file.
'
' WriteTextFileA
' - Writes a specified string (ASCII) to a file.
'
' WriteTextFileUTF8
' - Writes a specified string (UTF-8) to a file.
'
' AppendTextFileW
' - Writes a specified string (Unicode) to the end of a file.
'
' AppendTextFileA
' - Writes a specified string (ASCII) to the end of a file.
'
' AppendTextFileUTF8
' - Writes a specified string (UTF-8) to the end of a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Text:
'   Required. A String value that contains the text in characters to be
'   written.
'

Public Sub WriteTextFileW(FileName, Text)
    WriteTextFile FileName, Text, 0, "unicode"
End Sub

Public Sub WriteTextFileA(FileName, Text)
    WriteTextFile FileName, Text, 0, "iso-8859-1"
End Sub

Public Sub WriteTextFileUTF8( _
    FileName, _
    Text, _
    BOM)
    
    WriteTextFile FileName, Text, 0, "utf-8"
    
    If Not BOM Then
        Dim Binary
        Binary = ReadBinaryFile(FileName, 3)
        WriteBinaryFile FileName, Binary, 0
    End If
End Sub

Public Sub AppendTextFileW(FileName, Text)
    WriteTextFile FileName, Text, -1, "unicode"
End Sub

Public Sub AppendTextFileA(FileName, Text)
    WriteTextFile FileName, Text, -1, "iso-8859-1"
End Sub

Public Sub AppendTextFileUTF8( _
    FileName, _
    Text, _
    BOM)
    
    WriteTextFile FileName, Text, -1, "utf-8"
    
    If Not BOM Then
        Dim Binary
        Binary = ReadBinaryFile(FileName, 3)
        WriteBinaryFile FileName, Binary, 0
    End If
End Sub

Private Sub WriteTextFile( _
    FileName, _
    Text, _
    Position, _
    Charset)
    
    If FileName = "" Then Exit Sub
    
    WriteTextAndSaveToFile FileName, Text, Position, Charset
End Sub

'
' --- ADODB.Stream ---
'

'
' GetADODBStream
' - Returns a ADODB.Stream object.
'

Public Function GetADODBStream()
    'Static ADODBStream
    If IsEmpty(ADODBStream) Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    Set GetADODBStream = ADODBStream
End Function

'
' --- TextFile ---
'

'
' LoadFromFileAndReadText
' - Reads an entire file and returns the resulting string.
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'
' Charset:
'   Optional. A String value that specifies the character set into
'   which the contents of the Stream will be translated.
'   The default value is Unicode.
'   Allowed values are typical strings passed over the interface as
'   Internet character set names (for example, "iso-8859-1", "Windows-1252",
'   and so on).
'   For a list of the character set names that are known by a system,
'   see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset
'   in the Windows Registry.
'

Public Function LoadFromFileAndReadText( _
    FileName, _
    Charset)
    
    On Error Resume Next
    
    With GetADODBStream()
        .Type = 2 'ADODB.adTypeText
        If Charset <> "" Then .Charset = Charset
        .Open
        .LoadFromFile FileName
        LoadFromFileAndReadText = .ReadText
        .Close
    End With
End Function

'
' WriteTextAndSaveToFile
' - Writes a specified string to a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Text:
'   Required. A String value that contains the text in characters to be
'   written.
'
' Position:
'   Optional. Sets a Long value that specifies the offset, in number of
'   bytes, of the current position from the beginning of the stream.
'   The default is 0, which represents the first byte in the stream.
'
' Charset:
'   Optional. A String value that specifies the character set into
'   which the contents of the Stream will be translated.
'   The default value is Unicode.
'   Allowed values are typical strings passed over the interface as
'   Internet character set names (for example, "iso-8859-1", "Windows-1252",
'   and so on).
'   For a list of the character set names that are known by a system,
'   see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset
'   in the Windows Registry.
'

Public Sub WriteTextAndSaveToFile( _
    FileName, _
    Text, _
    Position, _
    Charset)
    
    On Error Resume Next
    
    With GetADODBStream()
        .Type = 2 'ADODB.adTypeText
        If Charset <> "" Then .Charset = Charset
        .Open
        If Position = 0 Then
            ' nop
        Else
            .LoadFromFile FileName
            If Position > 0 Then
                .Position = Position
                .SetEOS
            Else 'If Position < 0 Then
                .Position = .Size
            End If
        End If
        .WriteText Text
        .SaveToFile FileName, 2 'ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

'
' === BinaryFile ===
'

'
' ReadBinaryFile
' - Reads an entire file and returns the resulting data.
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'
' Position:
'   Optional. Sets a Long value that specifies the offset, in number of
'   bytes, of the current position from the beginning of the stream.
'   The default is 0, which represents the first byte in the stream.
'

Public Function ReadBinaryFile( _
    FileName, _
    Position)
    
    If FileName = "" Then Exit Function
    
    ReadBinaryFile = LoadFromFileAndRead(FileName, Position)
End Function

'
' WriteBinaryFile
' - Writes a binary data to a file.
'
' AppendBinaryFile
' - Writes a binary data to the end of a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Binary:
'   Required. A Variant that contains an array of bytes to be written.
'
' Position:
'   Optional. Sets a Long value that specifies the offset, in number of
'   bytes, of the current position from the beginning of the stream.
'   The default is 0, which represents the first byte in the stream.
'

Public Sub WriteBinaryFile( _
    FileName, _
    Binary, _
    Position)
    
    If FileName = "" Then Exit Sub
    
    WriteAndSaveToFile FileName, Binary, Position
End Sub

Public Sub AppendBinaryFile(FileName, Binary)
    WriteBinaryFile FileName, Binary, -1
End Sub

Public Sub WriteBinaryFileFromStringB(FileName, StringB)
    If FileName = "" Then Exit Sub
    
    Dim StringWB
    StringWB = GetStringWBFromStringB(StringB)
    
    WriteTextFileA FileName, StringWB
End Sub

Public Sub AppendBinaryFileFromStringB(FileName, StringB)
    If FileName = "" Then Exit Sub
    
    Dim StringWB
    StringWB = GetStringWBFromStringB(StringB)
    
    AppendTextFileA FileName, StringWB
End Sub

Private Function GetStringWBFromStringB(StringB)
    Dim StringWB
    Dim Index
    For Index = 1 To LenB(StringB)
        StringWB = StringWB & ChrW(AscB(MidB(StringB, Index, 1)))
    Next
    GetStringWBFromStringB = StringWB
End Function

'
' --- BinaryFile ---
'

'
' LoadFromFileAndRead
' - Reads an entire file and returns the resulting data.
'

'
' FileName:
'   Required. A String value that contains the name of a file.
'   FileName can contain any valid path and name in UNC format.
'
' Position:
'   Optional. Sets a Long value that specifies the offset, in number of
'   bytes, of the current position from the beginning of the stream.
'   The default is 0, which represents the first byte in the stream.
'

Public Function LoadFromFileAndRead( _
    FileName, _
    Position)
    
    On Error Resume Next
    
    With GetADODBStream()
        .Type = 1 'ADODB.adTypeBinary
        .Open
        .LoadFromFile FileName
        If Position > 0 Then .Position = Position
        LoadFromFileAndRead = .Read
        .Close
    End With
End Function

'
' WriteAndSaveToFile
' - Writes a binary data to a file.
'

'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Binary:
'   Required. A Variant that contains an array of bytes to be written.
'
' Position:
'   Optional. Sets a Long value that specifies the offset, in number of
'   bytes, of the current position from the beginning of the stream.
'   The default is 0, which represents the first byte in the stream.
'

Public Sub WriteAndSaveToFile( _
    FileName, _
    Binary, _
    Position)
    
    On Error Resume Next
    
    With GetADODBStream()
        .Type = 1 'ADODB.adTypeBinary
        .Open
        If Position = 0 Then
            ' nop
        Else
            .LoadFromFile FileName
            If Position > 0 Then
                .Position = Position
                .SetEOS
            Else 'If Position < 0 Then
                .Position = .Size
            End If
        End If
        .Write Binary
        .SaveToFile FileName, 2 'ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

'
' --- Text / Binary ---
'

'
' GetTextFromBinary
' - Return a string value that contains the text in characters.
'

'
' Binary:
'   Required. A Variant that contains an array of bytes.
'
' Charset:
'   Required. A String value that specifies the character set into
'   which the contents of the Stream will be translated.
'   The default value is Unicode.
'   Allowed values are typical strings passed over the interface as
'   Internet character set names (for example, "iso-8859-1", "Windows-1252",
'   and so on).
'   For a list of the character set names that are known by a system,
'   see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset
'   in the Windows Registry.
'

Public Function GetTextFromBinary(Binary, Charset)
    On Error Resume Next
    
    With GetADODBStream()
        .Open
        
        .Type = 1 'ADODB.adTypeBinary
        .Write Binary
        
        .Position = 0
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        GetTextFromBinary = .ReadText
        
        .Close
    End With
End Function

'
' GetBinaryFromText
' - Return a variant that contains an array of bytes.
'

'
' Text:
'   Required. A String value that contains the text in characters.
'
' Charset:
'   Required. A String value that specifies the character set into
'   which the contents of the Stream will be translated.
'   The default value is Unicode.
'   Allowed values are typical strings passed over the interface as
'   Internet character set names (for example, "iso-8859-1", "Windows-1252",
'   and so on).
'   For a list of the character set names that are known by a system,
'   see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset
'   in the Windows Registry.
'

Public Function GetBinaryFromText(Text, Charset)
    On Error Resume Next
    
    With GetADODBStream()
        .Open
        
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        .WriteText Text
        
        .Position = 0
        .Type = 1 'ADODB.adTypeBinary
        Select Case Charset
        Case "unicode"
            .Position = 2
        Case "utf-8"
            .Position = 3
        End Select
        GetBinaryFromText = .Read
        
        .Close
    End With
End Function

'
' --- Test ---
'

Private Sub Test_TextFileW()
    Dim FileName
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileW" & vbNewLine
    WriteTextFileW FileName, Text
    Text = ReadTextFileW(FileName)
    Debug_Print Text
    
    Text = "AppendTextFileW" & vbNewLine
    AppendTextFileW FileName, Text
    Text = ReadTextFileW(FileName)
    Debug_Print Text
End Sub

Private Sub Test_TextFileA()
    Dim FileName
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileA" & vbNewLine
    WriteTextFileA FileName, Text
    Text = ReadTextFileA(FileName)
    Debug_Print Text
    
    Text = "AppendTextFileA" & vbNewLine
    AppendTextFileA FileName, Text
    Text = ReadTextFileA(FileName)
    Debug_Print Text
End Sub

Private Sub Test_TextFileUTF8()
    Dim FileName
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileUTF8" & vbNewLine
    WriteTextFileUTF8 FileName, Text, True
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
    
    Text = "AppendTextFileUTF8" & vbNewLine
    AppendTextFileUTF8 FileName, Text, True
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
End Sub

Private Sub Test_TextFileUTF8_withoutBOM()
    Dim FileName
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileUTF8 (w/o BOM)" & vbNewLine
    WriteTextFileUTF8 FileName, Text, False
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
    
    Text = "AppendTextFileUTF8 (w/o BOM)" & vbNewLine
    AppendTextFileUTF8 FileName, Text, False
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
End Sub

Private Sub Test_BinaryFile()
    Dim FileName
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim StringB
    Dim Binary
    
    StringB = GetTestStringB()
    WriteBinaryFileFromStringB FileName, StringB
    Binary = ReadBinaryFile(FileName, 0)
    Debug_Print_Binary Binary
    
    StringB = GetTestStringB()
    AppendBinaryFileFromStringB FileName, StringB
    Binary = ReadBinaryFile(FileName, 0)
    Debug_Print_Binary Binary
End Sub

Private Function GetTestStringB()
    Dim StringB
    Dim Index
    For Index = 0 To 255
        StringB = StringB & ChrB(Index)
    Next
    GetTestStringB = StringB
End Function

Private Sub Test_GetBinaryGetTextA()
    Test_GetBinaryGetTextT "iso-8859-1"
End Sub

Private Sub Test_GetBinaryGetTextW()
    Test_GetBinaryGetTextT "unicode"
End Sub

Private Sub Test_GetBinaryGetTextUTF8()
    Test_GetBinaryGetTextT "utf-8"
End Sub

Private Sub Test_GetBinaryGetTextT(Charset)
    Dim Text0
    Text0 = "abcdefghijklmnopqrstuvwxyz"
    
    Dim Binary
    Binary = GetBinaryFromText(Text0, Charset)
    Debug_Print_Binary Binary
    
    Dim Text
    Text = GetTextFromBinary(Binary, Charset)
    Debug_Print Text
End Sub

Private Sub Debug_Print_Binary(Binary)
    Dim Text
    Dim Index1
    Dim Index2
    For Index1 = LBound(Binary) To UBound(Binary) Step 16
        For Index2 = Index1 To MinL(Index1 + 15, UBound(Binary))
            Text = Text & Right("0" & Hex(Binary(Index2)), 2) & " "
        Next
        Text = Text & vbNewLine
    Next
    
    Debug_Print "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
    Debug_Print Text
    Debug_Print "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
End Sub

Private Function MinL(Value1, Value2)
    If Value1 < Value2 Then
        MinL = Value1
    Else
        MinL = Value2
    End If
End Function

Private Function GetSaveAsFileName()
    Dim SaveAsFileName
    SaveAsFileName = Application.GetSaveAsFileName()
    If SaveAsFileName = False Then Exit Function
    GetSaveAsFileName = CStr(SaveAsFileName)
    'GetSaveAsFileName = InputBox("GetSaveAsFileName")
End Function

Private Sub Debug_Print(Str)
    Debug.Print Str
End Sub

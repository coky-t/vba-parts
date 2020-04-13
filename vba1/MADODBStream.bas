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

'
' --- ADODB.Stream ---
'

'
' GetADODBStream
' - Returns a ADODB.Stream object.
'

'
' ADODBStream:
'   Optional. The name of a ADODB.Stream object.
'

Public Function GetADODBStream( _
    Optional ADODBStream As ADODB.Stream) As ADODB.Stream
    
    If ADODBStream Is Nothing Then
        Set GetADODBStream = New ADODB.Stream
    Else
        Set GetADODBStream = ADODBStream
    End If
End Function

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
' ADODBStream:
'   Optional. The name of a ADODB.Stream object.
'

Public Function ReadTextFileW( _
    FileName As String, _
    Optional ADODBStream As ADODB.Stream) As String
    
    ReadTextFileW = ReadTextFileT(FileName, "unicode", ADODBStream)
End Function

Public Function ReadTextFileA( _
    FileName As String, _
    Optional ADODBStream As ADODB.Stream) As String
    
    ReadTextFileA = ReadTextFileT(FileName, "iso-8859-1", ADODBStream)
End Function

Public Function ReadTextFileUTF8( _
    FileName As String, _
    Optional ADODBStream As ADODB.Stream) As String
    
    ReadTextFileUTF8 = ReadTextFileT(FileName, "utf-8", ADODBStream)
End Function

Public Function ReadTextFileT( _
    FileName As String, _
    Optional Charset As String, _
    Optional ADODBStream As ADODB.Stream) As String
    
    ReadTextFileT = _
        ReadTextFile(GetADODBStream(ADODBStream), FileName, Charset)
End Function

Private Function ReadTextFile( _
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Optional Charset As String) As String
    
    If ADODBStream Is Nothing Then Exit Function
    
    If FileName = "" Then Exit Function
    
    ReadTextFile = LoadFromFileAndReadText(ADODBStream, FileName, Charset)
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
' ADODBStream:
'   Optional. The name of a ADODB.Stream object.
'

Public Sub WriteTextFileW( _
    FileName As String, _
    Text As String, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteTextFileT FileName, Text, 0, "unicode", ADODBStream
End Sub

Public Sub WriteTextFileA( _
    FileName As String, _
    Text As String, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteTextFileT FileName, Text, 0, "iso-8859-1", ADODBStream
End Sub

Public Sub WriteTextFileUTF8( _
    FileName As String, _
    Text As String, _
    Optional BOM As Boolean = True, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteTextFileT FileName, Text, 0, "utf-8", ADODBStream
    
    If Not BOM Then
        Dim Data() As Byte
        Data = ReadBinaryFile(FileName, 3, ADODBStream)
        WriteBinaryFile FileName, Data, ADODBStream
    End If
End Sub

Public Sub AppendTextFileW( _
    FileName As String, _
    Text As String, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteTextFileT FileName, Text, -1, "unicode", ADODBStream
End Sub

Public Sub AppendTextFileA( _
    FileName As String, _
    Text As String, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteTextFileT FileName, Text, -1, "iso-8859-1", ADODBStream
End Sub

Public Sub AppendTextFileUTF8( _
    FileName As String, _
    Text As String, _
    Optional BOM As Boolean = True, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteTextFileT FileName, Text, -1, "utf-8", ADODBStream
    
    If Not BOM Then
        Dim Data() As Byte
        Data = ReadBinaryFile(FileName, 3, ADODBStream)
        WriteBinaryFile FileName, Data, ADODBStream
    End If
End Sub

Public Sub WriteTextFileT( _
    FileName As String, _
    Text As String, _
    Optional Position As Long, _
    Optional Charset As String, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteTextFile _
        GetADODBStream(ADODBStream), FileName, Text, Position, Charset
End Sub

Private Sub WriteTextFile( _
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Text As String, _
    Optional Position As Long, _
    Optional Charset As String)
    
    If ADODBStream Is Nothing Then Exit Sub
    
    If FileName = "" Then Exit Sub
    
    WriteTextAndSaveToFile ADODBStream, FileName, Text, Position, Charset
End Sub

'
' --- TextFile ---
'

'
' LoadFromFileAndReadText
' - Reads an entire file and returns the resulting string.
'

'
' ADODBStream:
'   Required. The name of a ADODB.Stream object.
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
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Optional Charset As String) As String
    
    On Error Resume Next
    
    With ADODBStream
        .Type = adTypeText
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
' ADODBStream:
'   Required. The name of a ADODB.Stream object.
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
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Text As String, _
    Optional Position As Long, _
    Optional Charset As String)
    
    On Error Resume Next
    
    With ADODBStream
        .Type = adTypeText
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
        .SaveToFile FileName, ADODB.adSaveCreateOverWrite
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
' ADODBStream:
'   Optional. The name of a ADODB.Stream object.
'

Public Function ReadBinaryFile( _
    FileName As String, _
    Optional Position As Long, _
    Optional ADODBStream As ADODB.Stream) As Variant
    
    ReadBinaryFile = _
        ReadBinaryFileT(GetADODBStream(ADODBStream), FileName, Position)
End Function

Private Function ReadBinaryFileT( _
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Optional Position As Long) As Variant
    
    If ADODBStream Is Nothing Then Exit Function
    
    If FileName = "" Then Exit Function
    
    ReadBinaryFileT = LoadFromFileAndRead(ADODBStream, FileName, Position)
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
' Buffer:
'   Required. A Variant that contains an array of bytes to be written.
'
' ADODBStream:
'   Optional. The name of a ADODB.Stream object.
'

Public Sub WriteBinaryFile( _
    FileName As String, _
    Buffer() As Byte, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteBinaryFileT GetADODBStream(ADODBStream), FileName, Buffer, 0
End Sub

Public Sub AppendBinaryFile( _
    FileName As String, _
    Buffer() As Byte, _
    Optional ADODBStream As ADODB.Stream)
    
    WriteBinaryFileT GetADODBStream(ADODBStream), FileName, Buffer, -1
End Sub

Private Sub WriteBinaryFileT( _
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Buffer() As Byte, _
    Optional Position As Long)
    
    If ADODBStream Is Nothing Then Exit Sub
    
    If FileName = "" Then Exit Sub
    
    WriteAndSaveToFile ADODBStream, FileName, Buffer, Position
End Sub

'
' --- BinaryFile ---
'

'
' LoadFromFileAndRead
' - Reads an entire file and returns the resulting data.
'

'
' ADODBStream:
'   Required. The name of a ADODB.Stream object.
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
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Optional Position As Long) As Variant
    
    On Error Resume Next
    
    With ADODBStream
        .Type = ADODB.adTypeBinary
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
' ADODBStream:
'   Required. The name of a ADODB.Stream object.
'
' FileName:
'   Required. A String value that contains the fully-qualified name of
'   the file to which the contents of the Stream will be saved.
'   You can save to any valid local location, or any location you have
'   access to via a UNC value.
'
' Buffer:
'   Required. A Variant that contains an array of bytes to be written.
'
' Position:
'   Optional. Sets a Long value that specifies the offset, in number of
'   bytes, of the current position from the beginning of the stream.
'   The default is 0, which represents the first byte in the stream.
'

Public Sub WriteAndSaveToFile( _
    ADODBStream As ADODB.Stream, _
    FileName As String, _
    Buffer() As Byte, _
    Optional Position As Long)
    
    On Error Resume Next
    
    With ADODBStream
        .Type = ADODB.adTypeBinary
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
        .Write Buffer
        .SaveToFile FileName, ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

'
' --- Test ---
'

Private Sub Test_TextFileW()
    Dim FileName As String
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text As String
    
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
    Dim FileName As String
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text As String
    
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
    Dim FileName As String
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text As String
    
    Text = "WriteTextFileUTF8" & vbNewLine
    WriteTextFileUTF8 FileName, Text
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
    
    Text = "AppendTextFileUTF8" & vbNewLine
    AppendTextFileUTF8 FileName, Text
    Text = ReadTextFileUTF8(FileName)
    Debug_Print Text
End Sub

Private Sub Test_TextFileUTF8_withoutBOM()
    Dim FileName As String
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text As String
    
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
    Dim FileName As String
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Buffer(0 To 255) As Byte
    Dim Index As Integer
    For Index = 0 To 255
        Buffer(Index) = Index
    Next
    
    WriteBinaryFile FileName, Buffer
    
    Dim Data() As Byte
    Data = ReadBinaryFile(FileName)
    
    Dim Text As String
    Dim Index1 As Integer
    For Index1 = LBound(Data) To UBound(Data) Step 16
        Dim Index2 As Integer
        For Index2 = Index1 To Index1 + 15
            Text = Text & Right("0" & Hex(Data(Index2)), 2) & " "
        Next
        Text = Text & vbNewLine
    Next
    
    Debug_Print Text
    
    AppendBinaryFile FileName, Buffer
    Data = ReadBinaryFile(FileName)
    
    Text = ""
    For Index1 = LBound(Data) To UBound(Data) Step 16
        For Index2 = Index1 To Index1 + 15
            Text = Text & Right("0" & Hex(Data(Index2)), 2) & " "
        Next
        Text = Text & vbNewLine
    Next
    
    Debug_Print "---"
    Debug_Print Text
End Sub

Private Function GetSaveAsFileName() As String
    Dim SaveAsFileName As Variant
    SaveAsFileName = Application.GetSaveAsFileName()
    If SaveAsFileName = False Then Exit Function
    GetSaveAsFileName = CStr(SaveAsFileName)
    'GetSaveAsFileName = InputBox("GetSaveAsFileName")
End Function

Private Sub Debug_Print(Str As String)
    Debug.Print Str
End Sub

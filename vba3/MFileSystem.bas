Attribute VB_Name = "MFileSystem"
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
' - Scripting.FileSystemObject
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
' --- FileSystemObject ---
'

'
' GetFileSystemObject
' - Returns a FileSystemObject object.
'

'
' FileSystemObject:
'   Optional. The name of a FileSystemObject object.
'

Public Function GetFileSystemObject( _
    FileSystemObject)
    
    If FileSystemObject Is Nothing Then
        Set GetFileSystemObject = CreateObject("Scripting.FileSystemObject")
    Else
        Set GetFileSystemObject = FileSystemObject
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

'
' FileName:
'   Required. String expression that identifies the file to open.
'
' FileSystemObject:
'   Optional. The name of a FileSystemObject object.
'

Public Function ReadTextFileW( _
    FileName, _
    FileSystemObject)
    
    ReadTextFileW = _
        ReadTextFileT(FileName, Scripting_TristateTrue, FileSystemObject)
End Function

Public Function ReadTextFileA( _
    FileName, _
    FileSystemObject)
    
    ReadTextFileA = _
        ReadTextFileT(FileName, Scripting_TristateFalse, FileSystemObject)
End Function

Private Function ReadTextFileT( _
    FileName, _
    Format, _
    FileSystemObject)
    
    ReadTextFileT = _
        ReadTextFile(GetFileSystemObject(FileSystemObject), FileName, Format)
End Function

Private Function ReadTextFile( _
    FileSystemObject, _
    FileName, _
    Format)
    
    If FileSystemObject Is Nothing Then Exit Function
    
    If FileName = "" Then Exit Function
    If Not FileSystemObject.FileExists(FileName) Then Exit Function
    
    ReadTextFile = OpenTextFileAndReadAll(FileSystemObject, FileName, Format)
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
' FileName:
'   Required. String expression that identifies the file to create.
'
' Text:
'   Required. The text you want to write to the file.
'
' FileSystemObject:
'   Optional. The name of a FileSystemObject object.
'

Public Sub WriteTextFileW( _
    FileName, _
    Text, _
    FileSystemObject)
    
    WriteTextFileT _
        FileName, _
        Text, _
        Scripting_ForWriting, _
        Scripting_TristateTrue, _
        FileSystemObject
End Sub

Public Sub WriteTextFileA( _
    FileName, _
    Text, _
    FileSystemObject)
    
    WriteTextFileT _
        FileName, _
        Text, _
        Scripting_ForWriting, _
        Scripting_TristateFalse, _
        FileSystemObject
End Sub

Public Sub AppendTextFileW( _
    FileName, _
    Text, _
    FileSystemObject)
    
    WriteTextFileT _
        FileName, _
        Text, _
        Scripting_ForAppending, _
        Scripting_TristateTrue, _
        FileSystemObject
End Sub

Public Sub AppendTextFileA( _
    FileName, _
    Text, _
    FileSystemObject)
    
    WriteTextFileT _
        FileName, _
        Text, _
        Scripting_ForAppending, _
        Scripting_TristateFalse, _
        FileSystemObject
End Sub

Private Sub WriteTextFileT( _
    FileName, _
    Text, _
    IOMode, _
    Format, _
    FileSystemObject)
    
    WriteTextFile _
        GetFileSystemObject(FileSystemObject), _
        FileName, _
        Text, _
        IOMode, _
        Format
End Sub

Private Sub WriteTextFile( _
    FileSystemObject, _
    FileName, _
    Text, _
    IOMode, _
    Format)
    
    If FileSystemObject Is Nothing Then Exit Sub
    
    If FileName = "" Then Exit Sub
    If FileSystemObject.FolderExists(FileName) Then Exit Sub
    
    If IOMode = Scripting_ForReading Then Exit Sub
    
    MakeFolder _
        FileSystemObject, _
        GetParentFolderName(FileSystemObject, FileName)
    
    If IOMode = Scripting_ForWriting Then
        CreateTextFileAndWrite _
            FileSystemObject, _
            FileName, _
            Text, _
            (Format = Scripting_TristateTrue)
        Exit Sub
    End If
    
    OpenTextFileAndWrite FileSystemObject, FileName, Text, IOMode, Format
End Sub

'
' --- TextFile ---
'

'
' OpenTextFileAndReadAll
' - Reads an entire file and returns the resulting string.
'

'
' FileSystemObject:
'   Required. The name of a FileSystemObject object.
'
' FileName:
'   Required. String expression that identifies the file to open.
'
' Format:
'   Optional. One of three Tristate values used to indicate the format of
'   the opened file. If omitted, the file is opened as ASCII.
'   TristateUseDefault(-2): Opens the file by using the system default.
'   TristateTrue(-1): Opens the file as Unicode.
'   TristateFalse(0): Opens the file as ASCII.
'

Public Function OpenTextFileAndReadAll( _
    FileSystemObject, _
    FileName, _
    Format)
    
    On Error Resume Next
    
    With FileSystemObject.OpenTextFile(FileName, , , Format)
        OpenTextFileAndReadAll = .ReadAll
        .Close
    End With
End Function

'
' OpenTextFileAndWrite
' - Writes a specified string to a file.
'

'
' FileSystemObject:
'   Required. The name of a FileSystemObject object.
'
' FileName:
'   Required. String expression that identifies the file to create.
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

Public Sub OpenTextFileAndWrite( _
    FileSystemObject, _
    FileName, _
    Text, _
    IOMode, _
    Format)
    
    On Error Resume Next
    
    With FileSystemObject.OpenTextFile(FileName, IOMode, True, Format)
        .Write (Text)
        .Close
    End With
End Sub

'
' CreateTextFileAndWrite
' - Writes a specified string to a file.
'

'
' FileSystemObject:
'   Required. The name of a FileSystemObject object.
'
' FileName:
'   Required. String expression that identifies the file to create.
'
' Text:
'   Required. The text you want to write to the file.
'
' Unicode:
'   Optional. Boolean value that indicates whether the file is created
'   as a Unicode or ASCII file.
'   The value is True if the file is created as a Unicode file;
'   False if it's created as an ASCII file.
'   If omitted, an ASCII file is assumed.
'

Public Sub CreateTextFileAndWrite( _
    FileSystemObject, _
    FileName, _
    Text, _
    Unicode)
    
    On Error Resume Next
    
    With FileSystemObject.CreateTextFile(FileName, True, Unicode)
        .Write (Text)
        .Close
    End With
End Sub

'
' === Folder / Directory ===
'

'
' MakeDirectory
' - Creates a directory.
'

'
' DirName:
'   Required. String expression that identifies the directory to create.
'
' FileSystemObject:
'   Optional. The name of a FileSystemObject object.
'

Public Sub MakeDirectory( _
    DirName, _
    FileSystemObject)
    
    MakeFolder GetFileSystemObject(FileSystemObject), DirName
End Sub

Private Sub MakeFolder( _
    FileSystemObject, _
    FolderName)
    
    If FileSystemObject Is Nothing Then Exit Sub
    
    If FolderName = "" Then Exit Sub
    If FileSystemObject.FolderExists(FolderName) Then Exit Sub
    
    Dim FolderPath
    FolderPath = FileSystemObject.GetAbsolutePathName(FolderName)
    If FolderPath = "" Then Exit Sub
    
    Dim DriveName
    DriveName = FileSystemObject.GetDriveName(FolderPath)
    If Not DriveName = "" Then
        If Not FileSystemObject.DriveExists(DriveName) Then Exit Sub
    End If
    
    CreateFolder FileSystemObject, FolderPath
End Sub

'
' --- Folder / Directory ---
'

'
' CreateFolder
' - Creates a folder (recursively).
'

'
' FileSystemObject:
'   Required. The name of a FileSystemObject object.
'
' FolderPath:
'   Required. String expression that identifies the folder to create.
'

Public Sub CreateFolder( _
    FileSystemObject, _
    FolderPath)
    
    On Error Resume Next
    
    If FolderPath = "" Then Exit Sub
    
    With FileSystemObject
        If .FolderExists(FolderPath) Then Exit Sub
        
        CreateFolder FileSystemObject, .GetParentFolderName(FolderPath)
        .CreateFolder FolderPath
    End With
End Sub

'
' GetParentFolderName
' - Returns a string containing the name of the parent folder
'   of the last component in a specified path.
'

'
' FileSystemObject:
'   Required. The name of a FileSystemObject object.
'
' Path:
'   Required. String expression that identifies the folder.
'

Public Function GetParentFolderName( _
    FileSystemObject, _
    Path)
    
    On Error Resume Next
    
    With FileSystemObject
        GetParentFolderName = .GetParentFolderName(.GetAbsolutePathName(Path))
    End With
End Function

'
' --- Drive ---
'

'
' GetDriveName
' - Returns a string containing the name of the drive for a specified path.
'

'
' FileSystemObject:
'   Required. The name of a FileSystemObject object.
'
' Path:
'   Required. The path specification for the component whose drive name is
'   to be returned.
'

Public Function GetDriveName( _
    FileSystemObject, _
    Path)
    
    On Error Resume Next
    
    With FileSystemObject
        GetDriveName = .GetDriveName(.GetAbsolutePathName(Path))
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
    WriteTextFileW FileName, Text, Nothing
    Text = ReadTextFileW(FileName, Nothing)
    Debug_Print Text
    
    Text = "AppendTextFileW" & vbNewLine
    AppendTextFileW FileName, Text, Nothing
    Text = ReadTextFileW(FileName, Nothing)
    Debug_Print Text
End Sub

Private Sub Test_TextFileA()
    Dim FileName
    FileName = GetSaveAsFileName()
    If FileName = "" Then Exit Sub
    
    Dim Text
    
    Text = "WriteTextFileA" & vbNewLine
    WriteTextFileA FileName, Text, Nothing
    Text = ReadTextFileA(FileName, Nothing)
    Debug_Print Text
    
    Text = "AppendTextFileA" & vbNewLine
    AppendTextFileA FileName, Text, Nothing
    Text = ReadTextFileA(FileName, Nothing)
    Debug_Print Text
End Sub

Private Function GetSaveAsFileName()
    GetSaveAsFileName = Application.GetSaveAsFileName()
    'GetSaveAsFileName = InputBox("SaveAsFileName")
End Function

Private Sub Debug_Print(Str)
    Debug.Print Str
End Sub


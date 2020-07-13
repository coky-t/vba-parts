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

Public Function ReadTextFileW(FileName As String) As String
    ReadTextFileW = ReadTextFile(FileName, Scripting.TristateTrue)
End Function

Public Function ReadTextFileA(FileName As String) As String
    ReadTextFileA = ReadTextFile(FileName, Scripting.TristateFalse)
End Function

Private Function ReadTextFile( _
    FileName As String, _
    Optional Format As Scripting.Tristate) As String
    
    If FileName = "" Then Exit Function
    If Not GetFileSystemObject().FileExists(FileName) Then Exit Function
    
    ReadTextFile = OpenTextFileAndReadAll(FileName, Format)
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

Public Sub WriteTextFileW(FileName As String, Text As String)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting.ForWriting, _
        Scripting.TristateTrue
End Sub

Public Sub WriteTextFileA(FileName As String, Text As String)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting.ForWriting, _
        Scripting.TristateFalse
End Sub

Public Sub AppendTextFileW(FileName As String, Text As String)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting.ForAppending, _
        Scripting.TristateTrue
End Sub

Public Sub AppendTextFileA(FileName As String, Text As String)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting.ForAppending, _
        Scripting.TristateFalse
End Sub

Private Sub WriteTextFile( _
    FileName As String, _
    Text As String, _
    Optional IOMode As Scripting.IOMode = Scripting.ForWriting, _
    Optional Format As Scripting.Tristate)
    
    If FileName = "" Then Exit Sub
    If GetFileSystemObject().FolderExists(FileName) Then Exit Sub
    
    If IOMode = Scripting.ForReading Then Exit Sub
    
    MakeDirectory GetParentFolderName(FileName)
    
    If IOMode = Scripting.ForWriting Then
        CreateTextFileAndWrite _
            FileName, _
            Text, _
            (Format = Scripting.TristateTrue)
        Exit Sub
    End If
    
    OpenTextFileAndWrite FileName, Text, IOMode, Format
End Sub

'
' --- FileSystemObject ---
'

'
' GetFileSystemObject
' - Returns a FileSystemObject object.
'

Public Function GetFileSystemObject() As Scripting.FileSystemObject
    Static FileSystemObject As Scripting.FileSystemObject
    If FileSystemObject Is Nothing Then
        Set FileSystemObject = New Scripting.FileSystemObject
    End If
    Set GetFileSystemObject = FileSystemObject
End Function

'
' --- TextFile ---
'

'
' OpenTextFileAndReadAll
' - Reads an entire file and returns the resulting string.
'

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
    FileName As String, _
    Optional Format As Scripting.Tristate) As String
    
    On Error Resume Next
    
    With GetFileSystemObject()
        With .OpenTextFile(FileName, Format:=Format)
            OpenTextFileAndReadAll = .ReadAll
            .Close
        End With
    End With
End Function

'
' OpenTextFileAndWrite
' - Writes a specified string to a file.
'

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
    FileName As String, _
    Text As String, _
    Optional IOMode As Scripting.IOMode = Scripting.ForWriting, _
    Optional Format As Scripting.Tristate)
    
    On Error Resume Next
    
    With GetFileSystemObject()
        With .OpenTextFile(FileName, IOMode, True, Format)
            .Write (Text)
            .Close
        End With
    End With
End Sub

'
' CreateTextFileAndWrite
' - Writes a specified string to a file.
'

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
    FileName As String, _
    Text As String, _
    Optional Unicode As Boolean)
    
    On Error Resume Next
    
    With GetFileSystemObject()
        With .CreateTextFile(FileName, True, Unicode)
            .Write (Text)
            .Close
        End With
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

Public Sub MakeDirectory(DirName As String)
    Dim FileSystemObject As Scripting.FileSystemObject
    Set FileSystemObject = GetFileSystemObject()
    
    If FileSystemObject Is Nothing Then Exit Sub
    
    If DirName = "" Then Exit Sub
    If FileSystemObject.FolderExists(DirName) Then Exit Sub
    
    Dim FolderPath As String
    FolderPath = FileSystemObject.GetAbsolutePathName(DirName)
    If FolderPath = "" Then Exit Sub
    
    Dim DriveName As String
    DriveName = FileSystemObject.GetDriveName(FolderPath)
    If Not DriveName = "" Then
        If Not FileSystemObject.DriveExists(DriveName) Then Exit Sub
    End If
    
    CreateFolder FolderPath
End Sub

'
' --- Folder / Directory ---
'

'
' CreateFolder
' - Creates a folder (recursively).
'

'
' FolderPath:
'   Required. String expression that identifies the folder to create.
'

Public Sub CreateFolder(FolderPath As String)
    On Error Resume Next
    
    If FolderPath = "" Then Exit Sub
    
    With GetFileSystemObject()
        If .FolderExists(FolderPath) Then Exit Sub
        
        CreateFolder .GetParentFolderName(FolderPath)
        .CreateFolder FolderPath
    End With
End Sub

'
' GetParentFolderName
' - Returns a string containing the name of the parent folder
'   of the last component in a specified path.
'

'
' Path:
'   Required. String expression that identifies the folder.
'

Public Function GetParentFolderName(Path As String) As String
    On Error Resume Next
    
    With GetFileSystemObject()
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
' Path:
'   Required. The path specification for the component whose drive name is
'   to be returned.
'

Public Function GetDriveName(Path As String) As String
    On Error Resume Next
    
    With GetFileSystemObject()
        GetDriveName = .GetDriveName(.GetAbsolutePathName(Path))
    End With
End Function

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

Private Function GetSaveAsFileName() As String
    Dim SaveAsFileName As Variant
    SaveAsFileName = Application.GetSaveAsFileName()
    If SaveAsFileName = False Then Exit Function
    GetSaveAsFileName = CStr(SaveAsFileName)
    'GetSaveAsFileName = InputBox("SaveAsFileName")
End Function

Private Sub Debug_Print(Str As String)
    Debug.Print Str
End Sub

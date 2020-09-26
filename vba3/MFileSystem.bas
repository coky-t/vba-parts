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

Private FileSystemObject

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

Public Function ReadTextFileW(FileName)
    ReadTextFileW = ReadTextFile(FileName, Scripting_TristateTrue)
End Function

Public Function ReadTextFileA(FileName)
    ReadTextFileA = ReadTextFile(FileName, Scripting_TristateFalse)
End Function

Private Function ReadTextFile( _
    FileName, _
    Format)
    
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

Public Sub WriteTextFileW(FileName, Text)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting_ForWriting, _
        Scripting_TristateTrue
End Sub

Public Sub WriteTextFileA(FileName, Text)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting_ForWriting, _
        Scripting_TristateFalse
End Sub

Public Sub AppendTextFileW(FileName, Text)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting_ForAppending, _
        Scripting_TristateTrue
End Sub

Public Sub AppendTextFileA(FileName, Text)
    WriteTextFile _
        FileName, _
        Text, _
        Scripting_ForAppending, _
        Scripting_TristateFalse
End Sub

Private Sub WriteTextFile( _
    FileName, _
    Text, _
    IOMode, _
    Format)
    
    If FileName = "" Then Exit Sub
    If GetFileSystemObject().FolderExists(FileName) Then Exit Sub
    
    If IOMode = Scripting_ForReading Then Exit Sub
    
    MakeDirectory GetParentFolderName(FileName)
    
    If IOMode = Scripting_ForWriting Then
        CreateTextFileAndWrite _
            FileName, _
            Text, _
            (Format = Scripting_TristateTrue)
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

Public Function GetFileSystemObject()
    'Static FileSystemObject
    If IsEmpty(FileSystemObject) Then
        Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
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
    FileName, _
    Format)
    
    On Error Resume Next
    
    With GetFileSystemObject()
        With .OpenTextFile(FileName, , , Format)
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
    FileName, _
    Text, _
    IOMode, _
    Format)
    
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
    FileName, _
    Text, _
    Unicode)
    
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

Public Sub MakeDirectory(DirName)
    Dim FileSystemObject
    Set FileSystemObject = GetFileSystemObject()
    
    If FileSystemObject Is Nothing Then Exit Sub
    
    If DirName = "" Then Exit Sub
    If FileSystemObject.FolderExists(DirName) Then Exit Sub
    
    Dim FolderPath
    FolderPath = FileSystemObject.GetAbsolutePathName(DirName)
    If FolderPath = "" Then Exit Sub
    
    Dim DriveName
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

Public Sub CreateFolder(FolderPath)
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

Public Function GetParentFolderName(Path)
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

Public Function GetDriveName(Path)
    On Error Resume Next
    
    With GetFileSystemObject()
        GetDriveName = .GetDriveName(.GetAbsolutePathName(Path))
    End With
End Function

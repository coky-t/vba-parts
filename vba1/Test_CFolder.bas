Attribute VB_Name = "Test_CFolder"
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

Public Sub Test_CFolder_Folder()
    Test_CFolder_Folder_Core "C:\work"
End Sub

Public Sub Test_CFolder_Files()
    Test_CFolder_Files_Core "C:\work"
End Sub

Public Sub Test_CFolder_FilesAll()
    Test_CFolder_FilesAll_Core "C:\work"
End Sub

Public Sub Test_CFolder_Path()
    Test_CFolder_Path_Core "C:\work"
End Sub

Public Sub Test_CFolder_SubFolders()
    Test_CFolder_SubFolders_Core "C:\work"
End Sub

Public Sub Test_CFolder_SubFoldersAll()
    Test_CFolder_SubFoldersAll_Core "C:\work"
End Sub

Public Sub Test_CFolder_TempFolder()
    Test_CFolder_TempFolder_Core
End Sub

'
' --- Test Core ---
'

Public Sub Test_CFolder_Folder_Core(FolderName As String)
    Dim Folder As Scripting.Folder
    With New CFolder
        .Path = FolderName
        Set Folder = .Folder
    End With
    With New CFolder
        Set .Folder = Folder
        
        Debug_Print "-----"
        Debug_Print "Attributes: " & .Attributes
        Debug_Print "DateCreated: " & .DateCreated
        Debug_Print "DateLastAccessed: " & .DateLastAccessed
        Debug_Print "DateLastModified: " & .DateLastModified
        Debug_Print "Drive.Path: " & .Drive.Path
        Debug_Print "DriveName: " & .DriveName
        Debug_Print "IsRootFolder: " & .IsRootFolder
        Debug_Print "Name: " & .Name
        Debug_Print "ParentFolder.Path: " & .ParentFolder.Path
        Debug_Print "ParentFolderName: " & .ParentFolderName
        Debug_Print "ShortName: " & .ShortName
        Debug_Print "ShortPath: " & .ShortPath
        Debug_Print "Size: " & .Size
        Debug_Print "TypeName: " & .TypeName
    End With
End Sub

Public Sub Test_CFolder_Files_Core(FolderName As String)
    With New CFolder
        .Path = FolderName
        
        Debug_Print "-----"
        Debug_Print .Path
        Debug_Print "-----"
        Dim File As Scripting.File
        For Each File In .Files
            Debug_Print File.Path
        Next
    End With
End Sub

Public Sub Test_CFolder_FilesAll_Core(FolderName As String)
    With New CFolder
        .Path = FolderName
        
        Debug_Print "-----"
        Debug_Print .Path
        Debug_Print "-----"
        Dim File As Scripting.File
        For Each File In .FilesAll
            Debug_Print File.Path
        Next
    End With
End Sub

Public Sub Test_CFolder_Path_Core(FolderName As String)
    With New CFolder
        .Path = FolderName
        
        Debug_Print "-----"
        Debug_Print "Attributes: " & .Attributes
        Debug_Print "DateCreated: " & .DateCreated
        Debug_Print "DateLastAccessed: " & .DateLastAccessed
        Debug_Print "DateLastModified: " & .DateLastModified
        Debug_Print "Drive.Path: " & .Drive.Path
        Debug_Print "DriveName: " & .DriveName
        Debug_Print "IsRootFolder: " & .IsRootFolder
        Debug_Print "Name: " & .Name
        Debug_Print "ParentFolder.Path: " & .ParentFolder.Path
        Debug_Print "ParentFolderName: " & .ParentFolderName
        Debug_Print "ShortName: " & .ShortName
        Debug_Print "ShortPath: " & .ShortPath
        Debug_Print "Size: " & .Size
        Debug_Print "TypeName: " & .TypeName
    End With
End Sub

Public Sub Test_CFolder_SubFolders_Core(FolderName As String)
    With New CFolder
        .Path = FolderName
        
        Debug_Print "-----"
        Debug_Print .Path
        Debug_Print "-----"
        Dim Folder As Scripting.Folder
        For Each Folder In .SubFolders
            Debug_Print Folder.Path
        Next
    End With
End Sub

Public Sub Test_CFolder_SubFoldersAll_Core(FolderName As String)
    With New CFolder
        .Path = FolderName
        
        Debug_Print "-----"
        Debug_Print .Path
        Debug_Print "-----"
        Dim Folder As Scripting.Folder
        For Each Folder In .SubFoldersAll
            Debug_Print Folder.Path
        Next
    End With
End Sub

Public Sub Test_CFolder_TempFolder_Core()
    With New CFolder
        .GetTempFolderName
        
        Debug_Print "-----"
        Debug_Print .Path
        Debug_Print "-----"
    End With
End Sub

Attribute VB_Name = "Test_MFolder"
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

Public Sub Test_GetFolders()
    Test_GetFolders_Core "C:\work"
End Sub

Public Sub Test_GetFiles()
    Test_GetFiles_Core "C:\work"
End Sub

'
' --- Test Core ---
'

Public Sub Test_GetFolders_Core(FolderName)
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim FolderObject
    Set FolderObject = FSO.GetFolder(FolderName)
    If FolderObject Is Nothing Then Exit Sub
    
    Dim Folders
    Set Folders = GetFolders(FolderObject)
    Dim FolderTemp
    For Each FolderTemp In Folders
        Debug_Print FolderTemp.Path
    Next
End Sub

Public Sub Test_GetFiles_Core(FolderName)
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim FolderObject
    Set FolderObject = FSO.GetFolder(FolderName)
    If FolderObject Is Nothing Then Exit Sub
    
    Dim Files
    Set Files = GetFiles(FolderObject)
    Dim FileTemp
    For Each FileTemp In Files
        Debug_Print FileTemp.Path
    Next
End Sub

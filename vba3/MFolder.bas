Attribute VB_Name = "MFolder"
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
' - Scripting.Folder
'

'
' --- Folders ---
'

'
' GetFolders
' - Return collection of all Folder objects contained within a Folder object.
'

'
' FolderObject:
'   Required. The name of a Folder object.
'

Public Function GetFolders(FolderObject)
    If FolderObject Is Nothing Then Exit Function
    
    Dim Folders
    Set Folders = New Collection
    CollectFolders Folders, FolderObject
    Set GetFolders = Folders
End Function

Private Sub CollectFolders( _
    ByRef Folders, _
    FolderObject)
    
    If Folders Is Nothing Then Exit Sub
    If FolderObject Is Nothing Then Exit Sub
    
    If Not FolderObject.SubFolders Is Nothing Then
        If FolderObject.SubFolders.Count > 0 Then
            Dim SubFolder
            For Each SubFolder In FolderObject.SubFolders
                CollectFolders Folders, SubFolder
            Next
        End If
    End If
    
    Folders.Add FolderObject
End Sub

'
' --- Files ---
'

'
' GetFiles
' - Returns collection of all File objects contained within a Folder object.
'

'
' FolderObject:
'   Required. The name of a Folder object.
'

Public Function GetFiles(FolderObject)
    If FolderObject Is Nothing Then Exit Function
    
    Dim Files
    Set Files = New Collection
    CollectFiles Files, FolderObject
    Set GetFiles = Files
End Function

Private Sub CollectFiles( _
    ByRef Files, _
    FolderObject)
    
    If Files Is Nothing Then Exit Sub
    If FolderObject Is Nothing Then Exit Sub
    
    If Not FolderObject.SubFolders Is Nothing Then
        If FolderObject.SubFolders.Count > 0 Then
            Dim SubFolder
            For Each SubFolder In FolderObject.SubFolders
                CollectFiles Files, SubFolder
            Next
        End If
    End If
    
    If Not FolderObject.Files Is Nothing Then
        If FolderObject.Files.Count > 0 Then
            Dim FileObject
            For Each FileObject In FolderObject.Files
                Files.Add FileObject
            Next
        End If
    End If
End Sub

'
' --- Test ---
'

Private Sub Test_GetFolders()
    Dim FolderName
    FolderName = GetFolderName()
    If FolderName = "" Then Exit Sub
    
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

Private Sub Test_GetFiles()
    Dim FolderName
    FolderName = GetFolderName()
    If FolderName = "" Then Exit Sub
    
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

Private Function GetFolderName()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then
            GetFolderName = .SelectedItems(1)
        End If
    End With
    'GetFolderName = InputBox("FolderName")
End Function

Private Sub Debug_Print(Str)
    Debug.Print Str
End Sub


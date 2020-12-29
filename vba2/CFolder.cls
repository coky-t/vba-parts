VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' - Scripting.Folder
'

Private m_FileSystemObject As Object

Private m_Folder As Object
Private m_Path As String

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Set m_FileSystemObject = Nothing
    Reset
End Sub

'
' --- Private Method ---
'

'
' Reset
'

Private Sub Reset()
    Set m_Folder = Nothing
    m_Path = ""
End Sub

'
' --- Private Properties ---
'

'
' FileSystemObject
' - Returns a FileSystemObject object.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/filesystemobject-object
'

Private Property Get FileSystemObject() As Object
    If m_FileSystemObject Is Nothing Then
        Set m_FileSystemObject = CreateObject("Scripting.FileSystemObject")
    End If
    Set FileSystemObject = m_FileSystemObject
End Property

'
' --- Public Properties ---
'

'
' AbsolutePathName
' - Returns a complete and unambiguous path.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getabsolutepathname-method
'

Public Property Get AbsolutePathName() As String
    AbsolutePathName = FileSystemObject.GetAbsolutePathName(Path)
End Property

'
' Attributes
' - Sets or returns the attributes of folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/attributes-property
' https://docs.microsoft.com/en-us/office/vba/Language/Reference/user-interface-help/getattr-function
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/setattr-statement
'

Public Property Get Attributes() As Integer
    If Not m_Folder Is Nothing Then
        Attributes = m_Folder.Attributes
    ElseIf Not m_Path = "" Then
        Attributes = GetAttr(m_Path)
    End If
End Property

Public Property Let Attributes(Attributes_ As Integer)
    If Not m_File Is Nothing Then
        m_Folder.Attributes = Attributes_
    ElseIf Not m_Path = "" Then
        SetAttr m_Path, Attributes
    End If
End Property

'
' DateCreated
' - Returns the date and time that the specified folder was created.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datecreated-property
'

Public Property Get DateCreated() As Date
    DateCreated = Folder.DateCreated
End Property

'
' DateLastAccessed
' - Returns the date and time that the specified folder was last accessed.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datelastaccessed-property
'

Public Property Get DateLastAccessed() As Date
    DateLastAccessed = Folder.DateLastAccessed
End Property

'
' DateLastModified
' - Returns the date and time that the specified folder was last modified.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datelastmodified-property
'

Public Property Get DateLastModified() As Date
    DateLastModified = Folder.DateLastModified
End Property

'
' Drive
' - Returns a Drive object corresponding to the drive in a specified path.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getdrive-method
'

Public Property Get Drive() As Object
    Set Drive = FileSystemObject.GetDrive(DriveName)
End Property

'
' DriveName
' - Returns the drive letter of the drive on which the specified folder resides.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/drive-property
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getdrivename-method
'

Public Property Get DriveName() As String
    If Not m_Folder Is Nothing Then
        DriveName = m_Folder.Drive
    ElseIf Not m_Path = "" Then
        DriveName = FileSystemObject.GetDriveName(m_Path)
    End If
End Property

'
' Files
' - Returns a Files collection consisting of all File objects
'   contained in the specified folder, including those with hidden
'   and system file attributes set.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/files-property
'

Public Property Get Files() As Object
    Set Files = Folder.Files
End Property

Public Property Get FilesAll() As Collection
    Dim Files_ As New Collection
    CollectFiles Files_, Folder
    Set FilesAll = Files_
End Property

Private Sub CollectFiles( _
    ByRef Files_ As Collection, Folder_ As Object)
    
    If Files_ Is Nothing Then Exit Sub
    If Folder_ Is Nothing Then Exit Sub
    
    If Not Folder_.SubFolders Is Nothing Then
        If Folder_.SubFolders.Count > 0 Then
            Dim SubFolder As Object
            For Each SubFolder In Folder_.SubFolders
                CollectFiles Files_, SubFolder
            Next
        End If
    End If
    
    If Not Folder_.Files Is Nothing Then
        If Folder_.Files.Count > 0 Then
            Dim File_ As Object
            For Each File_ In Folder_.Files
                Files_.Add File_
            Next
        End If
    End If
End Sub

'
' Folder
' - Sets or returns a Folder object corresponding to the folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
'

Public Property Get Folder() As Object
    If Not m_Folder Is Nothing Then
        Set Folder = m_Folder
    ElseIf Not m_Path = "" Then
        Set Folder = FileSystemObject.GetFolder(m_Path)
    End If
End Property

Public Property Set Folder(Folder_ As Object)
    Reset
    Set m_Folder = Folder_
End Property

'
' IsRootFolder
' - Returns True if a folder is the root folder and False if not.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/isrootfolder-property
'

Public Property Get IsRootFolder() As Boolean
    IsRootFolder = Folder.IsRootFolder
End Property

'
' Name
' - Sets or returns the name of a specified folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/name-property-filesystemobject-object
'

Public Property Get Name() As String
    Name = Folder.Name
End Property

Public Property Let Name(Name_ As String)
    Reset
    m_Path = FileSystemObject.GetAbsolutePathName(Name_)
End Property

'
' ParentFolder
' - Returns the folder object for the parent of the specified folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/parentfolder-property
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getparentfoldername-method
'

Public Property Get ParentFolder() As Object
    If Not m_Folder Is Nothing Then
        Set ParentFolder = m_Folder.ParentFolder
    ElseIf Not m_Path = "" Then
        With FileSystemObject
            Set ParentFolder = .GetFolder(.GetParentFolderName(m_Path))
        End With
    End If
End Property

'
' ParentFolderName
' - Returns a string containing the name of the parent folder
'   of the last component in a specified path.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/parentfolder-property
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/path-property-filesystemobject-object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getparentfoldername-method
'

Public Property Get ParentFolderName() As String
    If Not m_Folder Is Nothing Then
        ParentFolderName = m_Folder.ParentFolder.Path
    ElseIf Not m_Path = "" Then
        ParentFolderName = FileSystemObject.GetParentFolderName(m_Path)
    End If
End Property

'
' Path
' - Sets or returns the path for a specified folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/path-property-filesystemobject-object
'

Public Property Get Path() As String
    If Not m_Folder Is Nothing Then
        Path = m_Folder.Path
    ElseIf Not m_Path = "" Then
        Path = m_Path
    End If
End Property

Public Property Let Path(Path_ As String)
    Reset
    m_Path = FileSystemObject.GetAbsolutePathName(Path_)
End Property

'
' ShortName
' - Returns the short name used by programs that require
'   the earlier 8.3 naming convention.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shortname-property
'

Public Property Get ShortName() As String
    ShortName = Folder.ShortName
End Property

'
' ShortPath
' - Returns the short path used by programs that require
'   the earlier 8.3 folder naming convention.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shortpath-property
'

Public Property Get ShortPath() As String
    ShortPath = Folder.ShortPath
End Property

'
' Size
' - returns the size, in bytes, of all files and subfolders
'   contained in the folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/size-property-filesystemobject-object
'

Public Property Get Size() As Long
    Size = Folder.Size
End Property

'
' SubFolders
' - Returns a Folders collection consisting of all folders
'   contained in a specified folder, including those with Hidden
'   and System file attributes set.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/subfolders-property
'

Public Property Get SubFolders() As Object
    Set SubFolders = Folder.SubFolders
End Property

Public Property Get SubFoldersAll() As Collection
    Dim Folders_ As New Collection
    CollectFolders Folders_, Folder
    Set SubFoldersAll = Folders_
End Property

Private Sub CollectFolders( _
    ByRef Folders_ As Collection, _
    Folder_ As Object)
    
    If Folders_ Is Nothing Then Exit Sub
    If Folder_ Is Nothing Then Exit Sub
    
    If Not Folder_.SubFolders Is Nothing Then
        If Folder_.SubFolders.Count > 0 Then
            Dim SubFolder_ As Object
            For Each SubFolder_ In Folder_.SubFolders
                CollectFolders Folders_, SubFolder_
            Next
        End If
    End If
    
    Folders_.Add Folder_
End Sub

'
' TypeName
' - Returns information about the type of a folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/type-property-filesystemobject-object
'

Public Property Get TypeName() As String
    TypeName = Folder.Type
End Property

'
' --- Public Methods ---
'

'
' Build
' - Combines a folder path and the name of a folder and returns the combination
'   with valid path separators.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/buildpath-method
'

Public Sub Build(ParentFolderName As String, FolderName As String)
    Reset
    m_Path = FileSystemObject.BuildPath(ParentFolderName, FolderName)
End Sub

'
' Copy
' - Copies a specified folder from one location to another.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copy-method-visual-basic-for-applications
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copyfolder-method
'

Public Sub Copy( _
    Destination As String, Optional OverWrite As Boolean = True)
    
    If Not m_Folder Is Nothing Then
        m_Folder.Copy Destination, OverWrite
    ElseIf Not m_Path = "" Then
        FileSystemObject.CopyFolder m_Path, Destination, OverWrite
    End If
End Sub

'
' Create
' - Creates a folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createfolder-method
'

Public Sub Create(Optional Recursive As Boolean)
    If Recursive Then
        CreateFolder Path
    Else
        FileSystemObject.CreateFolder Path
    End If
End Sub

Private Sub CreateFolder(Path_ As String)
    If Path_ = "" Then Exit Sub
    
    With FileSystemObject
        If .FolderExists(Path_) Then Exit Sub
        
        CreateFolder .GetParentFolderName(Path_)
        .CreateFolder Path_
    End With
End Sub

'
' Delete
' - Deletes a specified folder.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/delete-method-visual-basic-for-applications
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletefolder-method
'

Public Sub Delete(Optional Force As Boolean)
    If Not m_Folder Is Nothing Then
        m_File.Delete Force
    ElseIf Not m_Path = "" Then
        FileSystemObject.DeleteFolder m_Path, Force
    End If
End Sub

'
' Exists
' - Returns True if a specified folder exists; False if it does not.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folderexists-method
'

Public Function Exists() As Boolean
    Exists = FileSystemObject.FolderExists(Path)
End Function

'
' GetFolderName
' - Displays the standard Open dialog box and gets a folder name.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/api/excel.application.filedialog
'

Public Sub GetFolderName(Optional Title As String)
    With Application.FileDialog(Office.msoFileDialogFolderPicker)
        If Title <> "" Then .Title = Title
        If .Show = -1 Then
            Reset
            m_Path = .SelectedItems(1)
        End If
    End With
End Sub

'
' GetSpecialFolderName
' - Returns the special folder specified.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getspecialfolder-method
'

Public Sub GetSpecialFolderName( _
    SpecialFolder As Integer)
    
    Reset
    m_Path = FileSystemObject.GetSpecialFolder(SpecialFolder)
End Sub

Public Sub GetSystemFolderName()
    GetSpecialFolderName 1 ' SystemFolder
End Sub

Public Sub GetTempFolderName()
    GetSpecialFolderName 2 ' TemporaryFolder
End Sub

Public Sub GetWindowsFoldername()
    GetSpecialFolderName 0 ' WindowsFolder
End Sub

'
' Move
' - Moves a specified folder from one location to another.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/move-method-filesystemobject-object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/movefolder-method
'

Public Sub Move(Destination As String)
    If Not m_Folder Is Nothing Then
        m_Folder.Move Destination
    ElseIf Not m_Path = "" Then
        FileSystemObject.MoveFolder m_Path, Destination
        m_Path = FileSystemObject.GetAbsolutePathName(Destination)
    End If
End Sub

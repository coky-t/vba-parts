VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFile"
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
' - Scripting.File
'

'
' Microsoft ActiveX Data Objects X.X Library
' - ADODB.Stream
'

Private m_ADODBStream As ADODB.Stream
Private m_FileSystemObject As Scripting.FileSystemObject

Private m_File As Scripting.File
Private m_Path As String

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Set m_ADODBStream = Nothing
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
    Set m_File = Nothing
    m_Path = ""
End Sub

'
' --- Private Properties ---
'

'
' ADODBStream
' - Returns a ADODB.Stream object.
'
' Reference:
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/stream-object-ado
'

Private Property Get ADODBStream() As ADODB.Stream
    If m_ADODBStream Is Nothing Then
        Set m_ADODBStream = New ADODB.Stream
    End If
    Set ADODBStream = m_ADODBStream
End Property

'
' FileSystemObject
' - Returns a FileSystemObject object.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/filesystemobject-object
'

Private Property Get FileSystemObject() As Scripting.FileSystemObject
    If m_FileSystemObject Is Nothing Then
        Set m_FileSystemObject = New Scripting.FileSystemObject
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
' - Sets or returns the attributes of file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/attributes-property
' https://docs.microsoft.com/en-us/office/vba/Language/Reference/user-interface-help/getattr-function
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/setattr-statement
'

Public Property Get Attributes() As Integer
    If Not m_File Is Nothing Then
        Attributes = m_File.Attributes
    ElseIf Not m_Path = "" Then
        Attributes = GetAttr(m_Path)
    End If
End Property

Public Property Let Attributes(Attributes_ As Integer)
    If Not m_File Is Nothing Then
        m_File.Attributes = Attributes_
    ElseIf Not m_Path = "" Then
        SetAttr m_Path, Attributes
    End If
End Property

'
' BaseName
' - Returns a string containing the base name of the last component,
'   less any file extension, in a path.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getbasename-method
'

Public Property Get BaseName() As String
    BaseName = FileSystemObject.GetBaseName(Path)
End Property

'
' Binary
' - Reads an entire file and returns the resulting data.
' - Writes a binary data to a file.
'

Public Property Get Binary() As Variant
    Binary = ReadBinary()
End Property

Public Property Let Binary(Binary_ As Variant)
    WriteBinary Binary_
End Property

'
' DateCreated
' - Returns the date and time that the specified file was created.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datecreated-property
'

Public Property Get DateCreated() As Date
    DateCreated = File.DateCreated
End Property

'
' DateLastAccessed
' - Returns the date and time that the specified file was last accessed.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datelastaccessed-property
'

Public Property Get DateLastAccessed() As Date
    DateLastAccessed = File.DateLastAccessed
End Property

'
' DateLastModified
' - Returns the date and time that the specified file was last modified.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datelastmodified-property
'

Public Property Get DateLastModified() As Date
    DateLastModified = File.DateLastModified
End Property

'
' Drive
' - Returns a Drive object corresponding to the drive in a specified path.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getdrive-method
'

Public Property Get Drive() As Scripting.Drive
    Set Drive = FileSystemObject.GetDrive(DriveName)
End Property

'
' DriveName
' - Returns the drive letter of the drive on which the specified file resides.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/drive-property
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getdrivename-method
'

Public Property Get DriveName() As String
    If Not m_File Is Nothing Then
        DriveName = m_File.Drive
    ElseIf Not m_Path = "" Then
        DriveName = FileSystemObject.GetDriveName(m_Path)
    End If
End Property

'
' ExtensionName
' - Returns a string containing the extension name for the last component
'   in a path.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getextensionname-method
'

Public Property Get ExtensionName() As String
    ExtensionName = FileSystemObject.GetExtensionName(Path)
End Property

'
' File
' - Sets or returns a File object corresponding to the file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfile-method
'

Public Property Get File() As Scripting.File
    If Not m_File Is Nothing Then
        Set File = m_File
    ElseIf Not m_Path = "" Then
        Set File = FileSystemObject.GetFile(m_Path)
    End If
End Property

Public Property Set File(File_ As Scripting.File)
    Reset
    Set m_File = File_
End Property

'
' Name
' - Sets or returns the name of a specified file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/name-property-filesystemobject-object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfilename-method-visual-basic-for-applications
'

Public Property Get Name() As String
    If Not m_File Is Nothing Then
        Name = m_File.Name
    ElseIf Not m_Path = "" Then
        Name = FileSystemObject.GetFileName(m_Path)
    End If
End Property

Public Property Let Name(Name_ As String)
    Reset
    m_Path = FileSystemObject.GetAbsolutePathName(Name_)
End Property

'
' ParentFolder
' - Returns the folder object for the parent of the specified file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/parentfolder-property
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getparentfoldername-method
'

Public Property Get ParentFolder() As Scripting.Folder
    If Not m_File Is Nothing Then
        Set ParentFolder = m_File.ParentFolder
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
    If Not m_File Is Nothing Then
        ParentFolderName = m_File.ParentFolder.Path
    ElseIf Not m_Path = "" Then
        ParentFolderName = FileSystemObject.GetParentFolderName(m_Path)
    End If
End Property

'
' Path
' - Sets or returns the path for a specified file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/path-property-filesystemobject-object
'

Public Property Get Path() As String
    If Not m_File Is Nothing Then
        Path = m_File.Path
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
    ShortName = File.ShortName
End Property

'
' ShortPath
' - Returns the short path used by programs that require
'   the earlier 8.3 file naming convention.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/shortpath-property
'

Public Property Get ShortPath() As String
    ShortPath = File.ShortPath
End Property

'
' Size
' - Returns the size, in bytes, of the specified file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/size-property-filesystemobject-object
'

Public Property Get Size() As Long
    Size = File.Size
End Property

'
' Text
' - Reads an entire file and returns the resulting string.
' - Writes a specified string to a file.
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

Public Property Get Text(Charset As String) As String
    Text = ReadText(Charset)
End Property

Public Property Let Text(Charset As String, Text_ As String)
    WriteText Text_, Charset
End Property

'
' TextA
' - Reads an entire file and returns the resulting string (ASCII).
' - Writes a specified string (ASCII) to a file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/readall-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/write-method
'

Public Property Get TextA() As String
    If Not m_File Is Nothing Then
        With m_File.OpenAsTextStream()
            TextA = .ReadAll
            .Close
        End With
    ElseIf Not m_Path = "" Then
        With FileSystemObject.OpenTextFile(m_Path)
            TextA = .ReadAll
            .Close
        End With
    End If
End Property

Public Property Let TextA(TextA_ As String)
    If Not m_File Is Nothing Then
        With m_File.OpenAsTextStream(Scripting.ForWriting)
            .Write TextA_
            .Close
        End With
    ElseIf Not m_Path = "" Then
        With FileSystemObject.CreateTextFile(m_Path, True)
            .Write TextA_
            .Close
        End With
    End If
End Property

'
' TextB
' - Reads an entire file and returns the resulting string (Binary).
' - Writes a specified string (Binary) to a file.
'

Public Property Get TextB() As String
    TextB = Binary
End Property

Public Property Let TextB(TextB_ As String)
    Dim TextWB As String
    TextWB = GetTextWBFromTextB(TextB_)
    Text("iso-8859-1") = TextWB
End Property

Private Function GetTextWBFromTextB(TextB As String) As String
    Dim TextWB As String
    Dim Index As Long
    For Index = 1 To LenB(TextB)
        TextWB = TextWB & ChrW(AscB(MidB(TextB, Index, 1)))
    Next
    GetTextWBFromTextB = TextWB
End Function

'
' TextUTF8
' - Reads an entire file and returns the resulting string (UTF-8).
' - Writes a specified string (UTF-8) to a file.
'

Public Property Get TextUTF8() As String
    TextUTF8 = Text("utf-8")
End Property

Public Property Let TextUTF8(TextUTF8_ As String)
    Text("utf-8") = TextUTF8_
    
    ' Remove BOM
    Binary = ReadBinary(3)
End Property

'
' TextW
' - Reads an entire file and returns the resulting string (Unicode).
' - Writes a specified string (Unicode) to a file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/readall-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/write-method
'

Public Property Get TextW() As String
    If Not m_File Is Nothing Then
        With m_File.OpenAsTextStream(, Scripting.TristateTrue)
            TextW = .ReadAll
            .Close
        End With
    ElseIf Not m_Path = "" Then
        With FileSystemObject.OpenTextFile(m_Path, , , Scripting.TristateTrue)
            TextW = .ReadAll
            .Close
        End With
    End If
End Property

Public Property Let TextW(TextW_ As String)
    If Not m_File Is Nothing Then
        With m_File.OpenAsTextStream( _
            Scripting.ForWriting, Scripting.TristateTrue)
            
            .Write TextW_
            .Close
        End With
    ElseIf Not m_Path = "" Then
        With FileSystemObject.CreateTextFile(m_Path, True, True)
            .Write TextW_
            .Close
        End With
    End If
End Property

'
' TypeName
' - Returns information about the type of a file.
'   For example, for files ending in .TXT, "Text Document" is returned.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/type-property-filesystemobject-object
'

Public Property Get TypeName() As String
    TypeName = File.Type
End Property

'
' --- Public Methods ---
'

'
' AppendBinary
' - Writes a binary data to the end of a file.
'

Public Sub AppendBinary(Binary_ As Variant)
    WriteBinary Binary_, -1
End Sub

'
' AppendText
' - Writes a specified string to the end of a file.
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
' Reference:
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/charset-property-ado
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/writetext-method-ado
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/savetofile-method-ado
'

Public Sub AppendText(Text_ As String, Optional Charset As String)
    With ADODBStream
        .Type = ADODB.adTypeText
        If Charset <> "" Then .Charset = Charset
        .Open
        .LoadFromFile Path
        .Position = .Size
        .WriteText Text_
        .SaveToFile Path, ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

'
' AppendTextA
' - Writes a specified string (ASCII) to the end of a file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/openastextstream-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/write-method
'

Public Sub AppendTextA(TextA_ As String)
    If Not m_File Is Nothing Then
        With m_File.OpenAsTextStream(Scripting.ForAppending)
            .Write TextA_
            .Close
        End With
    ElseIf Not m_Path = "" Then
        With FileSystemObject.OpenTextFile( _
            m_Path, Scripting.ForAppending, True)
            
            .Write TextA_
            .Close
        End With
    End If
End Sub

'
' AppendTextB
' - Writes a specified string (Binary) to the end of a file.
'

Public Sub AppendTextB(TextB_ As String)
    Dim TextWB As String
    TextWB = GetTextWBFromTextB(TextB_)
    AppendText TextWB, "iso-8859-1"
End Sub

'
' AppendTextUTF8
' - Writes a specified string (UTF-8) to the end of a file.
'

Public Sub AppendTextUTF8( _
    TextUTF8_ As String, Optional BOM As Boolean = True)
    
    AppendText TextUTF8_, "utf-8"
    
    If Not BOM Then
        Binary = ReadBinary(3)
    End If
End Sub

'
' AppendTextW
' - Writes a specified string (Unicode) to the end of a file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/openastextstream-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/write-method
'

Public Sub AppendTextW(TextW_ As String)
    If Not m_File Is Nothing Then
        With m_File.OpenAsTextStream( _
            Scripting.ForAppending, Scripting.TristateTrue)
            
            .Write TextW_
            .Close
        End With
    ElseIf Not m_Path = "" Then
        With FileSystemObject.OpenTextFile( _
            m_Path, Scripting.ForAppending, True, Scripting.TristateTrue)
            
            .Write TextW_
            .Close
        End With
    End If
End Sub

'
' Build
' - Combines a folder path and the name of a file and returns the combination
'   with valid path separators.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/buildpath-method
'

Public Sub Build(ParentFolderName As String, FileName As String)
    Reset
    m_Path = FileSystemObject.BuildPath(ParentFolderName, FileName)
End Sub

'
' Copy
' - Copies a specified file from one location to another.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copy-method-visual-basic-for-applications
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copyfile-method
'

Public Sub Copy( _
    Destination As String, Optional OverWrite As Boolean = True)
    
    If Not m_File Is Nothing Then
        m_File.Copy Destination, OverWrite
    ElseIf Not m_Path = "" Then
        FileSystemObject.CopyFile m_Path, Destination, OverWrite
    End If
End Sub

'
' CreateTextFile
' - Creates a specified file name and returns a TextStream object
' that can be used to read from or write to the file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createtextfile-method
'

Public Function CreateTextFile( _
    Optional OverWrite As Boolean = True, _
    Optional Unicode As Boolean) As Scripting.TextStream
    
    Set CreateTextFile = _
        FileSystemObject.CreateTextFile(Path, OverWrite, Unicode)
End Function

'
' Delete
' - Deletes a specified file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/delete-method-visual-basic-for-applications
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletefile-method
'

Public Sub Delete(Optional Force As Boolean)
    If Not m_File Is Nothing Then
        m_File.Delete Force
    ElseIf Not m_Path = "" Then
        FileSystemObject.DeleteFile m_Path, Force
    End If
End Sub

'
' Exists
' - Returns True if a specified file exists; False if it does not.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fileexists-method
'

Public Function Exists() As Boolean
    Exists = FileSystemObject.FileExists(Path)
End Function

'
' GetOpenFileName
' - Displays the standard Open dialog box and gets a file name
'   from the user without actually opening any files.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/api/excel.application.getopenfilename
'

Public Sub GetOpenFileName( _
    Optional FileFilter As String, _
    Optional FilterIndex As Integer, _
    Optional Title As String)
    
    Dim OpenFileName As Variant
    OpenFileName = Application.GetOpenFileName(FileFilter, FilterIndex, Title)
    If OpenFileName = False Then Exit Sub
    
    Reset
    m_Path = OpenFileName
End Sub

'
' GetSaveAsFileName
' - Displays the standard Save As dialog box and gets a file name
'   from the user without actually saving any files.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/api/excel.application.getsaveasfilename
'

Public Sub GetSaveAsFileName( _
    Optional InitialFileName As String, _
    Optional FileFilter As String, _
    Optional FilterIndex As Integer, _
    Optional Title As String)
    
    Dim SaveAsFileName As Variant
    SaveAsFileName = _
        Application.GetSaveAsFileName( _
            InitialFileName, FileFilter, FilterIndex, Title)
    If SaveAsFileName = False Then Exit Sub
    
    Reset
    m_Path = SaveAsFileName
End Sub

'
' Move
' - Moves a specified file from one location to another.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/move-method-filesystemobject-object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/movefile-method
'

Public Sub Move(Destination As String)
    If Not m_File Is Nothing Then
        m_File.Move Destination
    ElseIf Not m_Path = "" Then
        FileSystemObject.MoveFile m_Path, Destination
        m_Path = FileSystemObject.GetAbsolutePathName(Destination)
    End If
End Sub

'
' OpenAsTextStream
' - Opens a specified file and returns a TextStream object
'   that can be used to read from, write to, or append to the file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/openastextstream-method
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
'

Public Function OpenAsTextStream( _
    Optional IOMode As Scripting.IOMode = Scripting.ForReading, _
    Optional Format As Scripting.Tristate) As Scripting.TextStream
    
    If Not m_File Is Nothing Then
        Set OpenAsTextStream = m_File.OpenAsTextStream(IOMode, Format)
    ElseIf Not m_Path = "" Then
        Set OpenAsTextStream = _
            FileSystemObject.OpenTextFile(m_Path, IOMode, True, Format)
    End If
End Function

'
' OpenTextFile
' - Opens a specified file and returns a TextStream object
'   that can be used to read from, write to, or append to the file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
'

Public Function OpenTextFile( _
    Optional IOMode As Scripting.IOMode = Scripting.ForReading, _
    Optional Create As Boolean, _
    Optional Format As Scripting.Tristate) As Scripting.TextStream
    
    Set OpenTextFile = _
        FileSystemObject.OpenTextFile(Path, IOMode, Create, Format)
End Function

'
' ReadBinary
' - Reads a specified number of bytes from a file and
'   returns the resulting data.
'
' Reference:
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/loadfromfile-method-ado
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/read-method-ado
'

Public Function ReadBinary( _
    Optional Position As Long, Optional NumBytes As Long) As Variant
    
    With ADODBStream
        .Type = ADODB.adTypeBinary
        .Open
        .LoadFromFile Path
        If Position > 0 Then .Position = Position
        If NumBytes > 0 Then
            ReadBinary = .Read(NumBytes)
        Else
            ReadBinary = .Read
        End If
        .Close
    End With
End Function

'
' ReadText
' - Reads an entire file and returns the resulting string.
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
' Reference:
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/charset-property-ado
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/loadfromfile-method-ado
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/readtext-method-ado
'

Public Function ReadText( _
    Optional Charset As String, Optional NumChars As Long) As String
    
    With ADODBStream
        .Type = ADODB.adTypeText
        If Charset <> "" Then .Charset = Charset
        .Open
        .LoadFromFile Path
        If NumChars > 0 Then
            ReadText = .ReadText(NumChars)
        Else
            ReadText = .ReadText
        End If
        .Close
    End With
End Function

'
' WriteBinary
' - Writes a binary data to a file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/write-method-ado
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/savetofile-method-ado
'

Public Sub WriteBinary(Binary_ As Variant, Optional Position As Long)
    With ADODBStream
        .Type = ADODB.adTypeBinary
        .Open
        If Position = 0 Then
            ' nop
        Else
            .LoadFromFile Path
            If Position > 0 Then
                .Position = Position
                .SetEOS
            Else 'If Position < 0 Then
                .Position = .Size
            End If
        End If
        .Write Binary_
        .SaveToFile Path, ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

'
' WriteText
' - Writes a specified string to a file.
'
' Reference:
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/writetext-method-ado
' https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/savetofile-method-ado
'

Public Sub WriteText(Text_ As String, Optional Charset As String)
    With ADODBStream
        .Type = ADODB.adTypeText
        If Charset <> "" Then .Charset = Charset
        .Open
        .WriteText Text_
        .SaveToFile Path, ADODB.adSaveCreateOverWrite
        .Close
    End With
End Sub

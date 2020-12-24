Attribute VB_Name = "Test_CFile"
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

Public Sub Test_CFile_Binary()
    Test_CFile_Binary_Core _
        "testb.dat", _
        GetTestBinary(), _
        GetTestBinary()
End Sub

Public Sub Test_CFile_File()
    Test_CFile_File_Core "testb.dat"
End Sub

Public Sub Test_CFile_Path()
    Test_CFile_Path_Core "testb.dat"
End Sub

Public Sub Test_CFile_TextA()
    Test_CFile_TextA_Core _
        "testa.txt", _
        "WriteTextFileA" & vbNewLine, _
        "AppendTextFileA" & vbNewLine
End Sub

Public Sub Test_CFile_TextB()
    Test_CFile_TextB_Core _
        "testb.txt", _
        GetTestStringB(), _
        GetTestStringB()
End Sub

Public Sub Test_CFile_TextUTF8()
    Test_CFile_TextUTF8_Core _
        "testutf8.txt", _
        "WriteTextFileUTF8" & vbNewLine, _
        "AppendTextFileUTF8" & vbNewLine
End Sub

Public Sub Test_CFile_TextW()
    Test_CFile_TextW_Core _
        "testw.txt", _
        "WriteTextFileW" & vbNewLine, _
        "AppendTextFileW" & vbNewLine
End Sub

'
' --- Test Core ---
'

Public Sub Test_CFile_Binary_Core( _
    FileName, _
    Binary1, _
    Binary2)
    
    With New CFile
        .Path = FileName
        .Binary = Binary1
        Debug_Print_Binary .Binary
        
        .AppendBinary Binary2
        Debug_Print_Binary .Binary
    End With
End Sub

Public Sub Test_CFile_File_Core(FileName)
    Dim File
    With New CFile
        .Path = FileName
        Set File = .File
    End With
    With New CFile
        Set .File = File
        Debug_Print "Attributes: " & .Attributes
        ' BaseName
        Debug_Print "DateCreated: " & .DateCreated
        Debug_Print "DateLastAccessed: " & .DateLastAccessed
        Debug_Print "DateLastModified: " & .DateLastModified
        Debug_Print "Drive.Path: " & .Drive.Path
        Debug_Print "DriveName: " & .DriveName
        ' ExtensionName
        Debug_Print "Name: " & .Name
        Debug_Print "ParentFolder.Path: " & .ParentFolder.Path
        Debug_Print "ParentFolderName: " & .ParentFolderName
        Debug_Print "ShortName: " & .ShortName
        Debug_Print "ShortPath: " & .ShortPath
        Debug_Print "Size: " & .Size
        Debug_Print "TypeName: " & .TypeName
    End With
End Sub

Public Sub Test_CFile_Path_Core(FileName)
    With New CFile
        .Path = FileName
        Debug_Print "Attributes: " & .Attributes
        Debug_Print "BaseName: " & .BaseName
        ' DateCreated
        ' DateLastAccessed
        ' DateLastModified
        Debug_Print "Drive.Path: " & .Drive.Path
        Debug_Print "DriveName: " & .DriveName
        Debug_Print "ExtensionName: " & .ExtensionName
        Debug_Print "Name: " & .Name
        Debug_Print "ParentFolder.Path: " & .ParentFolder.Path
        Debug_Print "ParentFolderName: " & .ParentFolderName
        ' ShortName
        ' ShortPath
        ' Size
        ' TypeName
    End With
End Sub

Public Sub Test_CFile_TextA_Core( _
    FileName, _
    Text1, _
    Text2)
    
    With New CFile
        .Path = FileName
        .TextA = Text1
        Debug_Print .TextA
        
        .AppendTextA Text2
        Debug_Print .TextA
    End With
End Sub

Public Sub Test_CFile_TextB_Core( _
    FileName, _
    Text1, _
    Text2)
    
    With New CFile
        .Path = FileName
        .TextB = Text1
        Debug_Print_Binary .TextB
        
        .AppendTextB Text2
        Debug_Print_Binary .TextB
    End With
End Sub

Public Sub Test_CFile_TextUTF8_Core( _
    FileName, _
    Text1, _
    Text2)
    
    With New CFile
        .Path = FileName
        .TextUTF8 = Text1
        Debug_Print .TextUTF8
        
        .AppendTextUTF8 Text2, False
        Debug_Print .TextUTF8
    End With
End Sub

Public Sub Test_CFile_TextW_Core( _
    FileName, _
    Text1, _
    Text2)
    
    With New CFile
        .Path = FileName
        .TextW = Text1
        Debug_Print .TextW
        
        .AppendTextW Text2
        Debug_Print .TextW
    End With
End Sub


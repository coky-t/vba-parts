Attribute VB_Name = "Test_SpdxTextFile"
Option Explicit

'
' Copyright (c) 2022 Koki Takeyama
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

Sub Test_SaveSpdxTextFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\spdx-text.txt"
    
    ' https://github.com/spdx/license-list-data/tree/v3.18/text
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_SaveSpdxTextFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\spdx-template.txt"
    
    ' https://github.com/spdx/license-list-data/tree/v3.18/template
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTextLinesFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\spdx-text-lines.txt"
    
    ' https://github.com/spdx/license-list-data/tree/v3.18/text
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_SaveSpdxTextLinesFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateLinesFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\spdx-template-lines.txt"
    
    ' https://github.com/spdx/license-list-data/tree/v3.18/template
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateLinesFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

'
' --- Test Core ---
'

Sub Test_SaveSpdxTextFile_Core( _
    OutputFilePath, DirPath)
    
    Dim OutputText
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText
        FileText = ReadTextFileUTF8(File.Path)
        OutputText = OutputText & _
            "<pre name=""" & _
            Left(File.Name, Len(File.Name) - Len(".txt")) & _
            """>" & ReplaceChars(FileText) & "</pre>" & vbCrLf
    Next
    
    WriteTextFileUTF8 OutputFilePath, OutputText, False
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateFile_Core( _
    OutputFilePath, DirPath)
    
    Dim OutputText
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText
        FileText = ReadTextFileUTF8(File.Path)
        OutputText = OutputText & _
            "<pre name=""" & _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & _
            """>" & ReplaceChars(FileText) & "</pre>" & vbCrLf
    Next
    
    WriteTextFileUTF8 OutputFilePath, OutputText, False
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTextLinesFile_Core( _
    OutputFilePath, DirPath)
    
    Dim OutputText
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText
        FileText = ReadTextFileUTF8(File.Path)
        
        Dim Lines
        Lines = Split(Replace(FileText, vbCrLf, vbLf), vbLf)
        
        Dim LB
        Dim UB
        LB = LBound(Lines)
        UB = UBound(Lines)
        
        Dim Index
        Dim Count
        Count = 1
        For Index = LB To UB
            If Lines(Index) <> "" Then
                OutputText = OutputText & _
                    "<pre name=""" & _
                    Left(File.Name, Len(File.Name) - Len(".txt")) & _
                    "_" & Right("00" & CStr(Count), 3) & _
                    """>" & ReplaceChars(Lines(Index)) & "</pre>" & vbCrLf
                Count = Count + 1
            End If
        Next
    Next
    
    WriteTextFileUTF8 OutputFilePath, OutputText, False
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateLinesFile_Core( _
    OutputFilePath, DirPath)
    
    Dim OutputText
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText
        FileText = ReadTextFileUTF8(File.Path)
        
        Dim Lines
        Lines = Split(Replace(FileText, vbCrLf, vbLf), vbLf)
        
        Dim LB
        Dim UB
        LB = LBound(Lines)
        UB = UBound(Lines)
        
        Dim Index
        Dim Count
        Count = 1
        For Index = LB To UB
            If Lines(Index) <> "" Then
                OutputText = OutputText & _
                    "<pre name=""" & _
                    Left(File.Name, Len(File.Name) - Len(".template.txt")) & _
                    "_" & Right("00" & CStr(Count), 3) & _
                    """>" & ReplaceChars(Lines(Index)) & "</pre>" & vbCrLf
                Count = Count + 1
            End If
        Next
    Next
    
    WriteTextFileUTF8 OutputFilePath, OutputText, False
    Debug_Print "... Done."
End Sub

Function ReplaceChars(Str)
    Dim Temp
    Temp = Str
    
    Temp = Replace(Temp, "&", "&amp;")
    Temp = Replace(Temp, ">", "&gt;")
    Temp = Replace(Temp, "<", "&lt;")
    'Temp = Replace(Temp, vbCrLf, "<br>")
    'Temp = Replace(Temp, vbLf, "<br>")
    
    ReplaceChars = Temp
End Function

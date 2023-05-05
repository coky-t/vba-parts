Attribute VB_Name = "Test_SpdxTextFile"
Option Explicit

'
' Copyright (c) 2022,2023 Koki Takeyama
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
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/text
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_SaveSpdxTextFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\spdx-template.txt"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTextLinesFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\spdx-text-lines.txt"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/text
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_SaveSpdxTextLinesFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateLinesFile()
    Dim OutputFilePath
    OutputFilePath = "C:\work\data\spdx-template-lines.txt"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateLinesFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateToTextFiles()
    Dim OutputDirPath
    OutputDirPath = "C:\work\data\spdx-license-template-to-text"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateToTextFiles_Core _
        OutputDirPath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateToTextFilesEx()
    Dim OutputDirPath
    OutputDirPath = "C:\work\data\spdx-license-template-to-text-ex"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateToTextFilesEx_Core _
        OutputDirPath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateToFontFiles()
    Dim OutputDirPath
    OutputDirPath = "C:\work\data\spdx-license-template-to-font"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateToFontFiles_Core _
        OutputDirPath, SpdxTextDirPath
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
    
    WriteTextFileUTF8 OutputFilePath, OutputText, True
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
    
    WriteTextFileUTF8 OutputFilePath, OutputText, True
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
    
    WriteTextFileUTF8 OutputFilePath, OutputText, True
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
    
    WriteTextFileUTF8 OutputFilePath, OutputText, True
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

Sub Test_SaveSpdxTemplateToTextFiles_Core( _
    OutputDirPath, DirPath)
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim InputFilePath
        InputFilePath = File.Path
        
        Dim OutputFileName
        OutputFileName = _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & ".txt"
        
        Dim OutputFilePath
        OutputFilePath = _
            GetFileSystemObject().BuildPath(OutputDirPath, OutputFileName)
        
        Test_SaveSpdxTemplateToTextFile_Core OutputFilePath, InputFilePath
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateToTextFile_Core( _
    OutputFilePath, InputFilePath)
    
    Dim InputText
    InputText = ReadTextFileUTF8(InputFilePath)
    
    Dim OutputText
    OutputText = GetPlainText(InputText)
    
    WriteTextFileUTF8 OutputFilePath, OutputText, True
End Sub

Sub Test_SaveSpdxTemplateToTextFilesEx_Core( _
    OutputDirPath, DirPath)
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim InputFilePath
        InputFilePath = File.Path
        
        Dim OutputFileName
        OutputFileName = _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & ".txt"
        
        Dim OutputFilePath
        OutputFilePath = _
            GetFileSystemObject().BuildPath(OutputDirPath, OutputFileName)
        
        Test_SaveSpdxTemplateToTextFileEx_Core OutputFilePath, InputFilePath
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateToTextFileEx_Core( _
    OutputFilePath, InputFilePath)
    
    Dim InputText
    InputText = ReadTextFileUTF8(InputFilePath)
    
    Dim OutputText
    OutputText = GetPlainTextEx(InputText)
    
    WriteTextFileUTF8 OutputFilePath, OutputText, True
End Sub

Sub Test_SaveSpdxTemplateToFontFiles_Core( _
    OutputDirPath, DirPath)
    
    Dim OKCount
    Dim NGCount
    OKCount = 0
    NGCount = 0
    
    Dim Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File
    For Each File In Folder.Files
        'Debug_Print File.Name
        
        Dim InputFilePath
        InputFilePath = File.Path
        
        Dim OutputFileName
        OutputFileName = _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & ".txt"
        
        Dim OutputFilePath
        OutputFilePath = _
            GetFileSystemObject().BuildPath(OutputDirPath, OutputFileName)
        
        Dim Result
        Result = _
            Test_SaveSpdxTemplateToFontFile_Core(OutputFilePath, InputFilePath)
        
        If Result Then
            Debug_Print File.Name & ": OK"
            OKCount = OKCount + 1
        Else
            Debug_Print File.Name & ": NG"
            NGCount = NGCount + 1
        End If
    Next
    
    Debug_Print "... Done."
    Debug_Print "OK: " & CStr(OKCount) & ", NG: " & CStr(NGCount)
End Sub

Function Test_SaveSpdxTemplateToFontFile_Core( _
    OutputFilePath, InputFilePath) As Boolean
    
    Dim InputText
    InputText = ReadTextFileUTF8(InputFilePath)
    
    Dim OutputText
    OutputText = GetFontText(InputText)
    
    WriteTextFileUTF8 OutputFilePath, OutputText, True
    
    Dim OutputTextTemp
    OutputTextTemp = GetPlainTextEx(InputText)
    
    Test_SaveSpdxTemplateToFontFile_Core = _
        (Len(OutputText) = Len(OutputTextTemp))
End Function

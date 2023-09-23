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
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\spdx-text.txt"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/text
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_SaveSpdxTextFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateFile()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\spdx-template.txt"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTextLinesFile()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\spdx-text-lines.txt"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/text
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_SaveSpdxTextLinesFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateLinesFile()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\spdx-template-lines.txt"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateLinesFile_Core _
        OutputFilePath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateToTextFiles()
    Dim OutputDirPath As String
    OutputDirPath = "C:\work\data\spdx-license-template-to-text"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateToTextFiles_Core _
        OutputDirPath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateToTextFilesEx()
    Dim OutputDirPath As String
    OutputDirPath = "C:\work\data\spdx-license-template-to-text-ex"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateToTextFilesEx_Core _
        OutputDirPath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateToFontFiles()
    Dim OutputDirPath As String
    OutputDirPath = "C:\work\data\spdx-license-template-to-font"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateToFontFiles_Core _
        OutputDirPath, SpdxTextDirPath
End Sub

Sub Test_SaveSpdxTemplateToPatternFiles()
    Dim OutputDirPath As String
    OutputDirPath = "C:\work\data\spdx-license-template-to-pattern"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/template
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-template"
    
    Test_SaveSpdxTemplateToPatternFiles_Core _
        OutputDirPath, SpdxTextDirPath
End Sub

Sub Test_CheckSpdxPatternFiles1()
    Dim SpdxPatternDirPath As String
    SpdxPatternDirPath = "C:\work\data\spdx-license-template-to-pattern"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/text
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_CheckSpdxPatternFiles1_Core _
        SpdxPatternDirPath, SpdxTextDirPath
End Sub

Sub Test_CheckSpdxPatternFiles2()
    Dim SpdxPatternDirPath As String
    SpdxPatternDirPath = "C:\work\data\spdx-license-template-to-pattern"
    
    ' https://github.com/spdx/license-list-data/tree/vX.XX/text
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Test_CheckSpdxPatternFiles2_Core _
        SpdxPatternDirPath, SpdxTextDirPath
End Sub

'
' --- Test Core ---
'

Sub Test_SaveSpdxTextFile_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim OutputText As String
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText As String
        FileText = ReadTextFileUTF8(File.Path)
        OutputText = OutputText & _
            "<pre name=""" & _
            Left(File.Name, Len(File.Name) - Len(".txt")) & _
            """>" & ReplaceChars(FileText) & "</pre>" & vbCrLf
    Next
    
    WriteTextFileUTF8 OutputFilePath, OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateFile_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim OutputText As String
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText As String
        FileText = ReadTextFileUTF8(File.Path)
        OutputText = OutputText & _
            "<pre name=""" & _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & _
            """>" & ReplaceChars(FileText) & "</pre>" & vbCrLf
    Next
    
    WriteTextFileUTF8 OutputFilePath, OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTextLinesFile_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim OutputText As String
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText As String
        FileText = ReadTextFileUTF8(File.Path)
        
        Dim Lines() As String
        Lines = Split(Replace(FileText, vbCrLf, vbLf), vbLf)
        
        Dim LB As Long
        Dim UB As Long
        LB = LBound(Lines)
        UB = UBound(Lines)
        
        Dim Index As Long
        Dim Count As Long
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
    
    WriteTextFileUTF8 OutputFilePath, OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateLinesFile_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim OutputText As String
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim FileText As String
        FileText = ReadTextFileUTF8(File.Path)
        
        Dim Lines() As String
        Lines = Split(Replace(FileText, vbCrLf, vbLf), vbLf)
        
        Dim LB As Long
        Dim UB As Long
        LB = LBound(Lines)
        UB = UBound(Lines)
        
        Dim Index As Long
        Dim Count As Long
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
    
    WriteTextFileUTF8 OutputFilePath, OutputText
    Debug_Print "... Done."
End Sub

Function ReplaceChars(Str As String) As String
    Dim Temp As String
    Temp = Str
    
    Temp = Replace(Temp, "&", "&amp;")
    Temp = Replace(Temp, ">", "&gt;")
    Temp = Replace(Temp, "<", "&lt;")
    'Temp = Replace(Temp, vbCrLf, "<br>")
    'Temp = Replace(Temp, vbLf, "<br>")
    
    ReplaceChars = Temp
End Function

Sub Test_SaveSpdxTemplateToTextFiles_Core( _
    OutputDirPath As String, DirPath As String)
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim InputFilePath As String
        InputFilePath = File.Path
        
        Dim OutputFileName As String
        OutputFileName = _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & ".txt"
        
        Dim OutputFilePath As String
        OutputFilePath = _
            GetFileSystemObject().BuildPath(OutputDirPath, OutputFileName)
        
        Test_SaveSpdxTemplateToTextFile_Core OutputFilePath, InputFilePath
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateToTextFile_Core( _
    OutputFilePath As String, InputFilePath As String)
    
    Dim InputText As String
    InputText = ReadTextFileUTF8(InputFilePath)
    
    Dim OutputText As String
    OutputText = GetPlainText(InputText)
    
    WriteTextFileUTF8 OutputFilePath, OutputText
End Sub

Sub Test_SaveSpdxTemplateToTextFilesEx_Core( _
    OutputDirPath As String, DirPath As String)
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim InputFilePath As String
        InputFilePath = File.Path
        
        Dim OutputFileName As String
        OutputFileName = _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & ".txt"
        
        Dim OutputFilePath As String
        OutputFilePath = _
            GetFileSystemObject().BuildPath(OutputDirPath, OutputFileName)
        
        Test_SaveSpdxTemplateToTextFileEx_Core OutputFilePath, InputFilePath
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateToTextFileEx_Core( _
    OutputFilePath As String, InputFilePath As String)
    
    Dim InputText As String
    InputText = ReadTextFileUTF8(InputFilePath)
    
    Dim OutputText As String
    OutputText = GetPlainTextEx(InputText)
    
    WriteTextFileUTF8 OutputFilePath, OutputText
End Sub

Sub Test_SaveSpdxTemplateToFontFiles_Core( _
    OutputDirPath As String, DirPath As String)
    
    Dim OKCount As Long
    Dim NGCount As Long
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        'Debug_Print File.Name
        
        Dim InputFilePath As String
        InputFilePath = File.Path
        
        Dim OutputFileName As String
        OutputFileName = _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & ".txt"
        
        Dim OutputFilePath As String
        OutputFilePath = _
            GetFileSystemObject().BuildPath(OutputDirPath, OutputFileName)
        
        Dim Result As Boolean
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
    OutputFilePath As String, InputFilePath As String) As Boolean
    
    Dim InputText As String
    InputText = ReadTextFileUTF8(InputFilePath)
    
    Dim OutputText As String
    OutputText = GetFontText(InputText)
    
    WriteTextFileUTF8 OutputFilePath, OutputText
    
    Dim OutputTextTemp As String
    OutputTextTemp = GetPlainTextEx(InputText)
    
    Test_SaveSpdxTemplateToFontFile_Core = _
        (Len(OutputText) = Len(OutputTextTemp))
End Function

Sub Test_SaveSpdxTemplateToPatternFiles_Core( _
    OutputDirPath As String, DirPath As String)
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Debug_Print File.Name
        
        Dim InputFilePath As String
        InputFilePath = File.Path
        
        Dim OutputFileName As String
        OutputFileName = _
            Left(File.Name, Len(File.Name) - Len(".template.txt")) & ".txt"
        
        Dim OutputFilePath As String
        OutputFilePath = _
            GetFileSystemObject().BuildPath(OutputDirPath, OutputFileName)
        
        Test_SaveSpdxTemplateToPatternFile_Core OutputFilePath, InputFilePath
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_SaveSpdxTemplateToPatternFile_Core( _
    OutputFilePath As String, InputFilePath As String)
    
    Dim InputText As String
    InputText = ReadTextFileUTF8(InputFilePath)
    
    Dim OutputText As String
    OutputText = GetMatchingText(InputText)
    
    WriteTextFileUTF8 OutputFilePath, OutputText
End Sub

Sub Test_CheckSpdxPatternFiles1_Core( _
    SpdxPatternDirPath As String, SpdxTextDirPath As String)
    
    Dim SpdxPatternFolder As Scripting.Folder
    Set SpdxPatternFolder = _
        GetFileSystemObject().GetFolder(SpdxPatternDirPath)
    
    Dim SpdxPatternFile As Scripting.File
    For Each SpdxPatternFile In SpdxPatternFolder.Files
        'Debug_Print SpdxPatternFile.Name
        
        Dim SpdxTextFileName As String
        SpdxTextFileName = SpdxPatternFile.Name
        
        Dim SpdxTextFilePath As String
        SpdxTextFilePath = _
            GetFileSystemObject().BuildPath(SpdxTextDirPath, SpdxTextFileName)
        
        Test_CheckSpdxPatternFile1_Core _
            SpdxPatternFile.Name, SpdxPatternFile.Path, SpdxTextFilePath
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_CheckSpdxPatternFile1_Core( _
    SpdxPatternFileName As String, SpdxPatternFilePath As String, _
    SpdxTextFilePath As String)
    
    Dim IgnoredFile As Boolean
    
    Select Case SpdxPatternFileName
    Case "CC-BY-SA-2.1-JP.txt", "DL-DE-BY-2.0.txt", "gSOAP-1.3b.txt", _
        "LPPL-1.3a.txt", "LPPL-1.3c.txt", _
        "MulanPSL-1.0.txt", "MulanPSL-2.0.txt", "NOSL.txt", _
        "Python-2.0.txt", "UCL-1.0.txt"
        IgnoredFile = True
    Case Else
        IgnoredFile = False
    End Select
    
    If IgnoredFile Then
        Debug_Print SpdxPatternFileName & ": Ignored"
        Exit Sub
    End If
        
    If Not GetFileSystemObject().FileExists(SpdxTextFilePath) Then
        Debug_Print SpdxPatternFileName & ": Not Exist"
        Exit Sub
    End If
    
    Dim SpdxPattern As String
    SpdxPattern = ReadTextFileUTF8(SpdxPatternFilePath)
    
    Dim SpdxText As String
    SpdxText = ReadTextFileUTF8(SpdxTextFilePath)
    
    If RegExp_Test(SpdxText, SpdxPattern, True, False) Then
        'Debug_Print SpdxPatternFileName & ": OK"
    Else
        Debug_Print SpdxPatternFileName & ": NG"
    End If
End Sub

Sub Test_CheckSpdxPatternFiles2_Core( _
    SpdxPatternDirPath As String, SpdxTextDirPath As String)
    
    Dim SpdxTextFolder As Scripting.Folder
    Set SpdxTextFolder = _
        GetFileSystemObject().GetFolder(SpdxTextDirPath)
    
    Dim SpdxText() As String
    ReDim SpdxText(SpdxTextFolder.Files.Count - 1, 1)
    
    Dim SpdxTextIndex As Long
    
    Dim SpdxTextFile As Scripting.File
    For Each SpdxTextFile In SpdxTextFolder.Files
        Dim SpdxTextTemp As String
        SpdxTextTemp = ReadTextFileUTF8(SpdxTextFile.Path)
        
        SpdxText(SpdxTextIndex, 0) = SpdxTextFile.Name
        SpdxText(SpdxTextIndex, 1) = SpdxTextTemp
        SpdxTextIndex = SpdxTextIndex + 1
    Next
    
    ' ---
    
    Dim SpdxPatternFolder As Scripting.Folder
    Set SpdxPatternFolder = _
        GetFileSystemObject().GetFolder(SpdxPatternDirPath)
    
    Dim SpdxPatternFile As Scripting.File
    For Each SpdxPatternFile In SpdxPatternFolder.Files
        'Debug_Print SpdxPatternFile.Name
        
        Dim SpdxTextFileName As String
        SpdxTextFileName = SpdxPatternFile.Name
        
        Dim SpdxTextFilePath As String
        SpdxTextFilePath = _
            GetFileSystemObject().BuildPath(SpdxTextDirPath, SpdxTextFileName)
        
        Dim SpdxPattern As String
        SpdxPattern = ReadTextFileUTF8(SpdxPatternFile.Path)
        
        Test_CheckSpdxPatternFile2_Text _
            SpdxPatternFile.Name, SpdxPattern, _
            SpdxText
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_CheckSpdxPatternFile2_Text( _
    SpdxPatternFileName As String, SpdxPattern As String, _
    SpdxText() As String)
    
    Select Case SpdxPatternFileName
    Case "CC-BY-SA-2.1-JP.txt", "DL-DE-BY-2.0.txt", "gSOAP-1.3b.txt", _
        "LPPL-1.3a.txt", "LPPL-1.3c.txt", _
        "MulanPSL-1.0.txt", "MulanPSL-2.0.txt", "NOSL.txt", _
        "Python-2.0.txt", "UCL-1.0.txt"
        Exit Sub
    End Select
    
    If Len(SpdxPattern) > 2000 Then
        'Debug_Print SpdxPatternFileName & ": Large"
        Exit Sub
    End If
    
    Dim SpdxTextIndex As Long
    For SpdxTextIndex = LBound(SpdxText) To UBound(SpdxText)
        Test_CheckSpdxPatternFile2_Text_Core _
            SpdxPatternFileName, SpdxPattern, _
            SpdxText(SpdxTextIndex, 0), SpdxText(SpdxTextIndex, 1)
    Next
End Sub

Sub Test_CheckSpdxPatternFile2_Text_Core( _
    SpdxPatternFileName As String, SpdxPattern As String, _
    SpdxTextFileName As String, SpdxText As String)
    
    Select Case SpdxTextFileName
    Case "CC-BY-SA-2.1-JP.txt", "DL-DE-BY-2.0.txt", "gSOAP-1.3b.txt", _
        "LPPL-1.3a.txt", "LPPL-1.3c.txt", _
        "MulanPSL-1.0.txt", "MulanPSL-2.0.txt", "NOSL.txt", _
        "Python-2.0.txt", "UCL-1.0.txt"
        Exit Sub
    Case "OGDL-Taiwan-1.0.txt", _
        "AFL-2.0.txt", "AFL-2.1.txt", "AFL-3.0.txt", _
        "MPL-2.0.txt", "MPL-2.0-no-copyleft-exception.txt"
        Exit Sub
    End Select
    
    If SpdxPatternFileName = SpdxTextFileName Then
        Exit Sub
    End If
    
    If Len(SpdxText) > 2000 Then
        'Debug_Print SpdxPatternFileName & ": " & SpdxTextFileName & ": Large"
        Exit Sub
    End If
    
    If RegExp_Test(SpdxText, SpdxPattern, True, False) Then
        Debug_Print SpdxPatternFileName & ": " & SpdxTextFileName & ": OK"
    Else
        'Debug_Print SpdxPatternFileName & ": " & SpdxTextFileName & ": NG"
    End If
End Sub

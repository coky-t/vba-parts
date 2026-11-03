Attribute VB_Name = "Test_ScanCodeTextFileEx"
Option Explicit

'
' Copyright (c) 2026 Koki Takeyama
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

Sub Test_SaveScanCodeTextFileEx()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\scancode\scancode-licenses.txt"
    ' CounterXXX: xxx
    '
    ' ### XXX
    '
    ' | Key |
    ' | --- |
    ' | xxx |
    '
    ' ### All
    '
    ' | Key | IsException | IsDeprecated | IsGeneric | IsSpdx |
    ' |-----|-------------|--------------|-----------|--------|
    ' | xxx | xxx         | xxx          | xxx       | xxx    |
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    ' https://github.com/aboutcode-org/scancode-toolkit/tree/develop/src/licensedcode/data/licenses
    Dim ScanCodeTextDirPath As String
    ScanCodeTextDirPath = "C:\work\data\scancode\scancode-licenses"
    
    Test_SaveScanCodeTextFileEx_Core _
        OutputFilePath, ScanCodeTextDirPath
End Sub

Sub Test_SaveScanCodeTextFileEx2()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\scancode\scancode-licenses-ex2.txt"
    ' CounterXXX: xxx
    '
    ' ### Normal, Exception, Generic, Generic Exception
    '
    ' | Short Name | Key |
    ' |------------|-----|
    ' | [xxx](xxx) | xxx |
    '
    ' ### Normal (Deprecated), Exception (Deprecated), Spdx (Deprecated), Spdx Exception (Deprecated), Generic (Deprecated), Generic Exception (Deprecated)
    '
    ' | Short Name | Key | Replaced By |
    ' |------------|-----|-------------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)  |
    '
    ' ### Spdx, Spdx Exception
    '
    ' | Short Name | Key | SPDX Identifier |
    ' |------------|-----|-----------------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)      |
    '
    ' ### All
    '
    ' | Short Name | Key | IsException | IsDeprecated | IsGeneric | IsSpdx |
    ' |------------|-----|-------------|--------------|-----------|--------|
    ' | [xxx](xxx) | xxx | xxx         | xxx          | xxx       | xxx    |
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    ' https://github.com/aboutcode-org/scancode-toolkit/tree/develop/src/licensedcode/data/licenses
    Dim ScanCodeTextDirPath As String
    ScanCodeTextDirPath = "C:\work\data\scancode\scancode-licenses"
    
    Test_SaveScanCodeTextFileEx2_Core _
        OutputFilePath, ScanCodeTextDirPath
End Sub

Sub Test_SaveScanCodeTextFileEx3()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\scancode\scancode-licenses-ex3.txt"
    ' CounterXXX: xxx
    '
    ' ### Normal, Exception, Generic, Generic Exception
    '
    ' | Short Name | Key |
    ' |------------|-----|
    ' | [xxx](xxx) | xxx |
    '
    ' ### Normal (Deprecated), Exception (Deprecated), Generic (Deprecated), Generic Exception (Deprecated)
    '
    ' | Short Name | Key | Replaced By |
    ' |------------|-----|-------------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)  |
    '
    ' ### Spdx, Spdx Exception
    '
    ' | Short Name | Key | SPDX Identifier |
    ' |------------|-----|-----------------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)      |
    '
    ' ### Spdx (Deprecated), Spdx Exception (Deprecated)
    '
    ' | Short Name | Key | Replaced By | SPDX Identifier |
    ' |------------|-----|-------------|-----------------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)  | [xxx](xxx)      |
    '
    ' ### All
    '
    ' | Short Name | Key | IsException | IsDeprecated | IsGeneric | IsSpdx |
    ' |------------|-----|-------------|--------------|-----------|--------|
    ' | [xxx](xxx) | xxx | xxx         | xxx          | xxx       | xxx    |
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    ' https://github.com/aboutcode-org/scancode-toolkit/tree/develop/src/licensedcode/data/licenses
    Dim ScanCodeTextDirPath As String
    ScanCodeTextDirPath = "C:\work\data\scancode\scancode-licenses"
    
    Test_SaveScanCodeTextFileEx3_Core _
        OutputFilePath, ScanCodeTextDirPath
End Sub

Sub Test_SaveScanCodeTextFileEx4()
    Dim OutputDirPath As String
    OutputDirPath = "C:\work\data\scancode\scancode-licenses-ex4"
    ' name
    ' name-pattern
    ' text
    ' text-pattern
    ' text-template
    ' url
    ' url-pattern
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    ' https://github.com/aboutcode-org/scancode-toolkit/tree/develop/src/licensedcode/data/licenses
    Dim ScanCodeTextDirPath As String
    ScanCodeTextDirPath = "C:\work\data\scancode\scancode-licenses"
    
    Test_SaveScanCodeTextFileEx4_Core _
        OutputDirPath, ScanCodeTextDirPath
End Sub

Sub Test_SaveScanCodeTextFileEx5()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\scancode\scancode-licenses-ex5.txt"
    ' CounterXXX: xxx
    '
    ' ### Normal, Exception, Generic, Generic Exception
    '
    ' | Short Name | Key | Ignorable | Min. Cov. |
    ' |------------|-----|-----------|-----------|
    ' | [xxx](xxx) | xxx | xxx       | xxx       |
    '
    ' ### Normal (Deprecated), Exception (Deprecated), Generic (Deprecated), Generic Exception (Deprecated)
    '
    ' | Short Name | Key | Replaced By | Ignorable | Min. Cov. |
    ' |------------|-----|-------------|-----------|-----------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)  | xxx       | xxx       |
    '
    ' ### Spdx, Spdx Exception
    '
    ' | Short Name | Key | SPDX Identifier | Ignorable | Min. Cov. |
    ' |------------|-----|-----------------|-----------|-----------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)      | xxx       | xxx       |
    '
    ' ### Spdx (Deprecated), Spdx Exception (Deprecated)
    '
    ' | Short Name | Key | Replaced By | SPDX Identifier | Ignorable | Min. Cov. |
    ' |------------|-----|-------------|-----------------|-----------|-----------|
    ' | [xxx](xxx) | xxx | [xxx](xxx)  | [xxx](xxx)      | xxx       | xxx       |
    '
    ' ### All
    '
    ' | Short Name | Key | IsException | IsDeprecated | IsGeneric | IsSpdx | Ignorable | Min. Cov. |
    ' |------------|-----|-------------|--------------|-----------|--------|-----------|-----------|
    ' | [xxx](xxx) | xxx | xxx         | xxx          | xxx       | xxx    | xxx       | xxx       |
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    ' https://github.com/aboutcode-org/scancode-toolkit/tree/develop/src/licensedcode/data/licenses
    Dim ScanCodeTextDirPath As String
    ScanCodeTextDirPath = "C:\work\data\scancode\scancode-licenses"
    
    Test_SaveScanCodeTextFileEx5_Core _
        OutputFilePath, ScanCodeTextDirPath
End Sub

Sub Test_SaveScanCodeTextFileEx6()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\scancode\scancode-licenses-ex6.txt"
    ' CounterXXX: xxx
    '
    ' ### Normal, Exception, Generic, Generic Exception
    '
    ' | Short Name | Key | Release |
    ' |------------|-----|---------|
    ' | [xxx](xxx) | xxx | xxx     |
    '
    ' ### Normal (Deprecated), Exception (Deprecated), Generic (Deprecated), Generic Exception (Deprecated)
    '
    ' | Short Name | Key | Release | Deprecated | Replaced By |
    ' |------------|-----|---------| ---------- |-------------|
    ' | [xxx](xxx) | xxx | xxx     | xxx        | [xxx](xxx)  |
    '
    ' ### Spdx, Spdx Exception
    '
    ' | Short Name | Key | Release | SPDX Identifier |
    ' |------------|-----|---------|-----------------|
    ' | [xxx](xxx) | xxx | xxx     | [xxx](xxx)      |
    '
    ' ### Spdx (Deprecated), Spdx Exception (Deprecated)
    '
    ' | Short Name | Key | Release | Deprecated | Replaced By | SPDX Identifier |
    ' |------------|-----|---------| ---------- |-------------|-----------------|
    ' | [xxx](xxx) | xxx | xxx     | xxx        | [xxx](xxx)  | [xxx](xxx)      |
    '
    ' ### All
    '
    ' | Short Name | Key | Release | IsException | IsDeprecated | IsGeneric | IsSpdx |
    ' |------------|-----|---------|-------------|--------------|-----------|--------|
    ' | [xxx](xxx) | xxx | xxx     | xxx         | xxx          | xxx       | xxx    |
    
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    ' https://github.com/aboutcode-org/scancode-toolkit/tree/develop/src/licensedcode/data/licenses
    Dim ScanCodeTextDirPath As String
    ScanCodeTextDirPath = "C:\work\data\scancode" ' "\scancode-licenses-x.x.x"
    
    Test_SaveScanCodeTextFileEx6_Core _
        OutputFilePath, ScanCodeTextDirPath
End Sub

Sub Test_SaveScanCodeTextFileEx7()
    Dim OutputFilePath As String
    OutputFilePath = "C:\work\data\scancode\scancode-licenses-ex7.txt"
    ' ### XXX, XXX (Deprecated)
    '
    ' | Version | Count | Count (Dep.) |
    ' |---------|-------|--------------|
    '
    ' ### All
    '
    ' | Version | Total |   | Nml. | Nml. (Dep.) | Nml. Exc. | Nml. Exc. (Dep.) | Spdx | Spdx (Dep.) | Spdx Exc. | Spdx Exc. (Dep.) | Gen. | Gen. (Dep.) | Gen. Exc. | Gen. Exc. (Dep.) |
    ' |---------|-------|---|------|-------------|-----------|------------------|------|-------------|-----------|------------------|------|-------------|-----------|------------------|
    ' | x.x.x   | xxx   |   | xxx  | xxx         | xxx       | xxx              | xxx  | xxx         | xxx       | xxx              | xxx  | xxx         |  xxx      | xxx              |
    
    
    ' https://github.com/aboutcode-org/scancode-licensedb/tree/main/docs
    ' https://github.com/aboutcode-org/scancode-toolkit/tree/develop/src/licensedcode/data/licenses
    Dim ScanCodeTextDirPath As String
    ScanCodeTextDirPath = "C:\work\data\scancode" ' "\scancode-licenses-x.x.x"
    
    Test_SaveScanCodeTextFileEx7_Core _
        OutputFilePath, ScanCodeTextDirPath
End Sub

'
' --- Test Core ---
'

Sub Test_SaveScanCodeTextFileEx_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim CounterNormal As Long
    Dim CounterNormalDeprecated As Long
    Dim CounterException As Long
    Dim CounterExceptionDeprecated As Long
    Dim CounterSpdx As Long
    Dim CounterSpdxDeprecated As Long
    Dim CounterSpdxException As Long
    Dim CounterSpdxExceptionDeprecated As Long
    Dim CounterGeneric As Long
    Dim CounterGenericDeprecated As Long
    Dim CounterGenericException As Long
    Dim CounterGenericExceptionDeprecated As Long
    
    Dim CounterAll As Long
    
    Dim OutputTextNormal As String
    Dim OutputTextNormalDeprecated As String
    Dim OutputTextException As String
    Dim OutputTextExceptionDeprecated As String
    Dim OutputTextSpdx As String
    Dim OutputTextSpdxDeprecated As String
    Dim OutputTextSpdxException As String
    Dim OutputTextSpdxExceptionDeprecated As String
    Dim OutputTextGeneric As String
    Dim OutputTextGenericDeprecated As String
    Dim OutputTextGenericException As String
    Dim OutputTextGenericExceptionDeprecated As String
    
    OutputTextNormal = _
        "### Normal" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextNormalDeprecated = _
        "### Normal (Deprecated)" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextException = _
        "### Exception" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextExceptionDeprecated = _
        "### Exception (Deprecated)" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextSpdx = _
        "### Spdx" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextSpdxDeprecated = _
        "### Spdx (Deprecated)" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextSpdxException = _
        "### Spdx Exception" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextSpdxExceptionDeprecated = _
        "### Spdx Exception (Deprecated)" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextGeneric = _
        "### Generic" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextGenericDeprecated = _
        "### Generic (Deprecated)" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextGenericException = _
        "### Generic Exception" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    OutputTextGenericExceptionDeprecated = _
        "### Generic Exception (Deprecated)" & vbLf & vbLf & _
        "| Key |" & vbLf & _
        "|-----|" & vbLf
        '"| Key | Dummy |" & vbLf & _
        '"|-----|-------|" & vbLf
        
    Dim OutputTextAll As String
    OutputTextAll = _
        "### All" & vbLf & vbLf & _
        "| Key | IsException | IsDeprecated | IsGeneric | IsSpdx |" & vbLf & _
        "|-----|-------------|--------------|-----------|--------|" & vbLf
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim Key As String
            Dim IsException As Boolean
            Dim IsDeprecated As Boolean
            Dim IsGeneric As Boolean
            Dim IsSpdx As Boolean
            
            Key = GetItem(FileText, "key:")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            CounterAll = CounterAll + 1
            
            OutputTextAll = OutputTextAll & "|"
            OutputTextAll = OutputTextAll & " " & Key
            OutputTextAll = OutputTextAll & " |"
            
            If IsException Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsDeprecated Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsGeneric Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsSpdx Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            OutputTextAll = OutputTextAll & vbLf
            
            '
            
            If IsGeneric Then
                If IsException Then
                    If IsDeprecated Then
                        CounterGenericExceptionDeprecated = CounterGenericExceptionDeprecated + 1
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & "|"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Key
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " |"
                        'OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & vbLf
                        
                    Else
                        CounterGenericException = CounterGenericException + 1
                        OutputTextGenericException = OutputTextGenericException & "|"
                        OutputTextGenericException = OutputTextGenericException & " " & Key
                        OutputTextGenericException = OutputTextGenericException & " |"
                        'OutputTextGenericException = OutputTextGenericException & " |"
                        OutputTextGenericException = OutputTextGenericException & vbLf
                        
                    End If
                Else
                    If IsDeprecated Then
                        CounterGenericDeprecated = CounterGenericDeprecated + 1
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & "|"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Key
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " |"
                        'OutputTextGenericDeprecated = OutputTextGenericDeprecated & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & vbLf
                        
                    Else
                        CounterGeneric = CounterGeneric + 1
                        OutputTextGeneric = OutputTextGeneric & "|"
                        OutputTextGeneric = OutputTextGeneric & " " & Key
                        OutputTextGeneric = OutputTextGeneric & " |"
                        'OutputTextGeneric = OutputTextGeneric & " |"
                        OutputTextGeneric = OutputTextGeneric & vbLf
                        
                    End If
                End If
            ElseIf IsException Then
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxExceptionDeprecated = CounterSpdxExceptionDeprecated + 1
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & "|"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Key
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " |"
                        'OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & vbLf
                        
                    Else
                        CounterExceptionDeprecated = CounterExceptionDeprecated + 1
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & "|"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Key
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " |"
                        'OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdxException = CounterSpdxException + 1
                        OutputTextSpdxException = OutputTextSpdxException & "|"
                        OutputTextSpdxException = OutputTextSpdxException & " " & Key
                        OutputTextSpdxException = OutputTextSpdxException & " |"
                        'OutputTextSpdxException = OutputTextSpdxException & " |"
                        OutputTextSpdxException = OutputTextSpdxException & vbLf
                        
                    Else
                        CounterException = CounterException + 1
                        OutputTextException = OutputTextException & "|"
                        OutputTextException = OutputTextException & " " & Key
                        OutputTextException = OutputTextException & " |"
                        'OutputTextException = OutputTextException & " |"
                        OutputTextException = OutputTextException & vbLf
                        
                    End If
                End If
            Else
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxDeprecated = CounterSpdxDeprecated + 1
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & "|"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Key
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " |"
                        'OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & vbLf
                        
                    Else
                        CounterNormalDeprecated = CounterNormalDeprecated + 1
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & "|"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Key
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " |"
                        'OutputTextNormalDeprecated = OutputTextNormalDeprecated & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdx = CounterSpdx + 1
                        OutputTextSpdx = OutputTextSpdx & "|"
                        OutputTextSpdx = OutputTextSpdx & " " & Key
                        OutputTextSpdx = OutputTextSpdx & " |"
                        'OutputTextSpdx = OutputTextSpdx & " |"
                        OutputTextSpdx = OutputTextSpdx & vbLf
                        
                    Else
                        CounterNormal = CounterNormal + 1
                        OutputTextNormal = OutputTextNormal & "|"
                        OutputTextNormal = OutputTextNormal & " " & Key
                        OutputTextNormal = OutputTextNormal & " |"
                        'OutputTextNormal = OutputTextNormal & " |"
                        OutputTextNormal = OutputTextNormal & vbLf
                        
                    End If
                End If
            End If
    
        End If
    Next
    
    Dim OutputCounter As String
    OutputCounter = _
        "CounterNormal: " & CStr(CounterNormal) & vbLf & _
        "CounterNormalDeprecated: " & CStr(CounterNormalDeprecated) & vbLf & vbLf & _
        "CounterException: " & CStr(CounterException) & vbLf & _
        "CounterExceptionDeprecated: " & CStr(CounterExceptionDeprecated) & vbLf & vbLf & _
        "CounterSpdx: " & CStr(CounterSpdx) & vbLf & _
        "CounterSpdxDeprecated: " & CStr(CounterSpdxDeprecated) & vbLf & vbLf & _
        "CounterSpdxException: " & CStr(CounterSpdxException) & vbLf & _
        "CounterSpdxExceptionDeprecated: " & CStr(CounterSpdxExceptionDeprecated) & vbLf & vbLf & _
        "CounterGeneric: " & CStr(CounterGeneric) & vbLf & _
        "CounterGenericDeprecated: " & CStr(CounterGenericDeprecated) & vbLf & _
        "CounterGenericException: " & CStr(CounterGenericException) & vbLf & _
        "CounterGenericExceptionDeprecated: " & CStr(CounterGenericExceptionDeprecated) & vbLf & vbLf & _
        "CounterAll: " & CStr(CounterAll) & vbLf & vbLf
    
    Dim OutputText As String
    OutputText = _
        OutputTextNormal & vbLf & _
        OutputTextNormalDeprecated & vbLf & _
        OutputTextException & vbLf & _
        OutputTextExceptionDeprecated & vbLf & _
        OutputTextSpdx & vbLf & _
        OutputTextSpdxDeprecated & vbLf & _
        OutputTextSpdxException & vbLf & _
        OutputTextSpdxExceptionDeprecated & vbLf & _
        OutputTextGeneric & vbLf & _
        OutputTextGenericDeprecated & vbLf & _
        OutputTextGenericException & vbLf & _
        OutputTextGenericExceptionDeprecated & vbLf & _
        OutputTextAll
    
    WriteTextFileUTF8 OutputFilePath, OutputCounter & OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveScanCodeTextFileEx2_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim CounterNormal As Long
    Dim CounterNormalDeprecated As Long
    Dim CounterException As Long
    Dim CounterExceptionDeprecated As Long
    Dim CounterSpdx As Long
    Dim CounterSpdxDeprecated As Long
    Dim CounterSpdxException As Long
    Dim CounterSpdxExceptionDeprecated As Long
    Dim CounterGeneric As Long
    Dim CounterGenericDeprecated As Long
    Dim CounterGenericException As Long
    Dim CounterGenericExceptionDeprecated As Long
    
    Dim CounterAll As Long
    
    Dim OutputTextNormal As String
    Dim OutputTextNormalDeprecated As String
    Dim OutputTextException As String
    Dim OutputTextExceptionDeprecated As String
    Dim OutputTextSpdx As String
    Dim OutputTextSpdxDeprecated As String
    Dim OutputTextSpdxException As String
    Dim OutputTextSpdxExceptionDeprecated As String
    Dim OutputTextGeneric As String
    Dim OutputTextGenericDeprecated As String
    Dim OutputTextGenericException As String
    Dim OutputTextGenericExceptionDeprecated As String
    
    OutputTextNormal = _
        "### Normal" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextNormalDeprecated = _
        "### Normal (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextException = _
        "### Exception" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextExceptionDeprecated = _
        "### Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextSpdx = _
        "### Spdx" & vbLf & vbLf & _
        "| Short Name | Key | SPDX Identifier |" & vbLf & _
        "|------------|-----|-----------------|" & vbLf
        
    OutputTextSpdxDeprecated = _
        "### Spdx (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextSpdxException = _
        "### Spdx Exception" & vbLf & vbLf & _
        "| Short Name | Key | SPDX Identifier |" & vbLf & _
        "|------------|-----|-----------------|" & vbLf
        
    OutputTextSpdxExceptionDeprecated = _
        "### Spdx Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextGeneric = _
        "### Generic" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextGenericDeprecated = _
        "### Generic (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextGenericException = _
        "### Generic Exception" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextGenericExceptionDeprecated = _
        "### Generic Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    Dim OutputTextAll As String
    OutputTextAll = _
        "### All" & vbLf & vbLf & _
        "| Short Name | Key | IsException | IsDeprecated | IsGeneric | IsSpdx |" & vbLf & _
        "|------------|-----|-------------|--------------|-----------|--------|" & vbLf
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim Key As String
            Dim ShortName As String
            Dim SpdxLicenseKey As String
            Dim ReplacedBy As String
            Dim IsException As Boolean
            Dim IsDeprecated As Boolean
            Dim IsGeneric As Boolean
            Dim IsSpdx As Boolean
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            CounterAll = CounterAll + 1
            
            OutputTextAll = OutputTextAll & "|"
            OutputTextAll = OutputTextAll & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
            OutputTextAll = OutputTextAll & " " & Key & " |"
            
            If IsException Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsDeprecated Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsGeneric Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsSpdx Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            OutputTextAll = OutputTextAll & vbLf
            
            '
            
            If IsGeneric Then
                If IsException Then
                    If IsDeprecated Then
                        CounterGenericExceptionDeprecated = CounterGenericExceptionDeprecated + 1
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & "|"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Key & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & vbLf
                        
                    Else
                        CounterGenericException = CounterGenericException + 1
                        OutputTextGenericException = OutputTextGenericException & "|"
                        OutputTextGenericException = OutputTextGenericException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericException = OutputTextGenericException & " " & Key & " |"
                        OutputTextGenericException = OutputTextGenericException & vbLf
                        
                    End If
                Else
                    If IsDeprecated Then
                        CounterGenericDeprecated = CounterGenericDeprecated + 1
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & "|"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Key & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & vbLf
                        
                    Else
                        CounterGeneric = CounterGeneric + 1
                        OutputTextGeneric = OutputTextGeneric & "|"
                        OutputTextGeneric = OutputTextGeneric & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGeneric = OutputTextGeneric & " " & Key & " |"
                        OutputTextGeneric = OutputTextGeneric & vbLf
                        
                    End If
                End If
            ElseIf IsException Then
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxExceptionDeprecated = CounterSpdxExceptionDeprecated + 1
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & "|"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Key & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & vbLf
                        
                    Else
                        CounterExceptionDeprecated = CounterExceptionDeprecated + 1
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & "|"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Key & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdxException = CounterSpdxException + 1
                        OutputTextSpdxException = OutputTextSpdxException & "|"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & Key & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        OutputTextSpdxException = OutputTextSpdxException & vbLf
                        
                    Else
                        CounterException = CounterException + 1
                        OutputTextException = OutputTextException & "|"
                        OutputTextException = OutputTextException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextException = OutputTextException & " " & Key & " |"
                        OutputTextException = OutputTextException & vbLf
                        
                    End If
                End If
            Else
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxDeprecated = CounterSpdxDeprecated + 1
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & "|"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Key & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & vbLf
                        
                    Else
                        CounterNormalDeprecated = CounterNormalDeprecated + 1
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & "|"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Key & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdx = CounterSpdx + 1
                        OutputTextSpdx = OutputTextSpdx & "|"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & Key & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        OutputTextSpdx = OutputTextSpdx & vbLf
                        
                    Else
                        CounterNormal = CounterNormal + 1
                        OutputTextNormal = OutputTextNormal & "|"
                        OutputTextNormal = OutputTextNormal & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormal = OutputTextNormal & " " & Key & " |"
                        OutputTextNormal = OutputTextNormal & vbLf
                        
                    End If
                End If
            End If
    
        End If
    Next
    
    Dim OutputCounter As String
    OutputCounter = _
        "CounterNormal: " & CStr(CounterNormal) & vbLf & _
        "CounterNormalDeprecated: " & CStr(CounterNormalDeprecated) & vbLf & vbLf & _
        "CounterException: " & CStr(CounterException) & vbLf & _
        "CounterExceptionDeprecated: " & CStr(CounterExceptionDeprecated) & vbLf & vbLf & _
        "CounterSpdx: " & CStr(CounterSpdx) & vbLf & _
        "CounterSpdxDeprecated: " & CStr(CounterSpdxDeprecated) & vbLf & vbLf & _
        "CounterSpdxException: " & CStr(CounterSpdxException) & vbLf & _
        "CounterSpdxExceptionDeprecated: " & CStr(CounterSpdxExceptionDeprecated) & vbLf & vbLf & _
        "CounterGeneric: " & CStr(CounterGeneric) & vbLf & _
        "CounterGenericDeprecated: " & CStr(CounterGenericDeprecated) & vbLf & _
        "CounterGenericException: " & CStr(CounterGenericException) & vbLf & _
        "CounterGenericExceptionDeprecated: " & CStr(CounterGenericExceptionDeprecated) & vbLf & vbLf & _
        "CounterAll: " & CStr(CounterAll) & vbLf & vbLf
    
    Dim OutputText As String
    OutputText = _
        OutputTextNormal & vbLf & _
        OutputTextNormalDeprecated & vbLf & _
        OutputTextException & vbLf & _
        OutputTextExceptionDeprecated & vbLf & _
        OutputTextSpdx & vbLf & _
        OutputTextSpdxDeprecated & vbLf & _
        OutputTextSpdxException & vbLf & _
        OutputTextSpdxExceptionDeprecated & vbLf & _
        OutputTextGeneric & vbLf & _
        OutputTextGenericDeprecated & vbLf & _
        OutputTextGenericException & vbLf & _
        OutputTextGenericExceptionDeprecated & vbLf & _
        OutputTextAll
    
    WriteTextFileUTF8 OutputFilePath, OutputCounter & OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveScanCodeTextFileEx3_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim CounterNormal As Long
    Dim CounterNormalDeprecated As Long
    Dim CounterException As Long
    Dim CounterExceptionDeprecated As Long
    Dim CounterSpdx As Long
    Dim CounterSpdxDeprecated As Long
    Dim CounterSpdxException As Long
    Dim CounterSpdxExceptionDeprecated As Long
    Dim CounterGeneric As Long
    Dim CounterGenericDeprecated As Long
    Dim CounterGenericException As Long
    Dim CounterGenericExceptionDeprecated As Long
    
    Dim CounterAll As Long
    
    Dim OutputTextNormal As String
    Dim OutputTextNormalDeprecated As String
    Dim OutputTextException As String
    Dim OutputTextExceptionDeprecated As String
    Dim OutputTextSpdx As String
    Dim OutputTextSpdxDeprecated As String
    Dim OutputTextSpdxException As String
    Dim OutputTextSpdxExceptionDeprecated As String
    Dim OutputTextGeneric As String
    Dim OutputTextGenericDeprecated As String
    Dim OutputTextGenericException As String
    Dim OutputTextGenericExceptionDeprecated As String
    
    OutputTextNormal = _
        "### Normal" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextNormalDeprecated = _
        "### Normal (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextException = _
        "### Exception" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextExceptionDeprecated = _
        "### Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextSpdx = _
        "### Spdx" & vbLf & vbLf & _
        "| Short Name | Key | SPDX Identifier |" & vbLf & _
        "|------------|-----|-----------------|" & vbLf
        
    OutputTextSpdxDeprecated = _
        "### Spdx (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | SPDX Identifier |" & vbLf & _
        "|------------|-----|-------------|-----------------|" & vbLf
        
    OutputTextSpdxException = _
        "### Spdx Exception" & vbLf & vbLf & _
        "| Short Name | Key | SPDX Identifier |" & vbLf & _
        "|------------|-----|-----------------|" & vbLf
        
    OutputTextSpdxExceptionDeprecated = _
        "### Spdx Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | SPDX Identifier |" & vbLf & _
        "|------------|-----|-------------|-----------------|" & vbLf
        
    OutputTextGeneric = _
        "### Generic" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextGenericDeprecated = _
        "### Generic (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    OutputTextGenericException = _
        "### Generic Exception" & vbLf & vbLf & _
        "| Short Name | Key |" & vbLf & _
        "|------------|-----|" & vbLf
        
    OutputTextGenericExceptionDeprecated = _
        "### Generic Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By |" & vbLf & _
        "|------------|-----|-------------|" & vbLf
        
    Dim OutputTextAll As String
    OutputTextAll = _
        "### All" & vbLf & vbLf & _
        "| Short Name | Key | IsException | IsDeprecated | IsGeneric | IsSpdx |" & vbLf & _
        "|------------|-----|-------------|--------------|-----------|--------|" & vbLf
    
    Dim LicenseDic As Object
    Set LicenseDic = CreateObject("Scripting.Dictionary")
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim Key As String
            Dim ShortName As String
            Dim SpdxLicenseKey As String
            Dim ReplacedBy As String
            Dim IsException As Boolean
            Dim IsDeprecated As Boolean
            Dim IsGeneric As Boolean
            Dim IsSpdx As Boolean
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            Dim LicenseInfo As ScanCodeLicense
            Set LicenseInfo = New ScanCodeLicense
            With LicenseInfo
                .Key = Key
                .ShortName = ShortName
                .SpdxLicenseKey = SpdxLicenseKey
                .ReplacedBy = ReplacedBy
                .IsException = IsException
                .IsDeprecated = IsDeprecated
                .IsGeneric = IsGeneric
                .IsSpdx = IsSpdx
            End With
            
            LicenseDic.Add Key, LicenseInfo
        End If
    Next
    
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            'Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            'Dim Key As String
            'Dim ShortName As String
            'Dim SpdxLicenseKey As String
            'Dim ReplacedBy As String
            'Dim IsException As Boolean
            'Dim IsDeprecated As Boolean
            'Dim IsGeneric As Boolean
            'Dim IsSpdx As Boolean
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            CounterAll = CounterAll + 1
            
            OutputTextAll = OutputTextAll & "|"
            OutputTextAll = OutputTextAll & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
            OutputTextAll = OutputTextAll & " " & Key & " |"
            
            If IsException Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsDeprecated Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsGeneric Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsSpdx Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            OutputTextAll = OutputTextAll & vbLf
            
            '
            
            If IsGeneric Then
                If IsException Then
                    If IsDeprecated Then
                        CounterGenericExceptionDeprecated = CounterGenericExceptionDeprecated + 1
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & "|"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Key & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & vbLf
                        
                    Else
                        CounterGenericException = CounterGenericException + 1
                        OutputTextGenericException = OutputTextGenericException & "|"
                        OutputTextGenericException = OutputTextGenericException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericException = OutputTextGenericException & " " & Key & " |"
                        OutputTextGenericException = OutputTextGenericException & vbLf
                        
                    End If
                Else
                    If IsDeprecated Then
                        CounterGenericDeprecated = CounterGenericDeprecated + 1
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & "|"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Key & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & vbLf
                        
                    Else
                        CounterGeneric = CounterGeneric + 1
                        OutputTextGeneric = OutputTextGeneric & "|"
                        OutputTextGeneric = OutputTextGeneric & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGeneric = OutputTextGeneric & " " & Key & " |"
                        OutputTextGeneric = OutputTextGeneric & vbLf
                        
                    End If
                End If
            ElseIf IsException Then
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxExceptionDeprecated = CounterSpdxExceptionDeprecated + 1
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & "|"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Key & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetSpdxLinksFromReplacedBy(ReplacedBy, LicenseDic) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & vbLf
                        
                    Else
                        CounterExceptionDeprecated = CounterExceptionDeprecated + 1
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & "|"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Key & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdxException = CounterSpdxException + 1
                        OutputTextSpdxException = OutputTextSpdxException & "|"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & Key & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        OutputTextSpdxException = OutputTextSpdxException & vbLf
                        
                    Else
                        CounterException = CounterException + 1
                        OutputTextException = OutputTextException & "|"
                        OutputTextException = OutputTextException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextException = OutputTextException & " " & Key & " |"
                        OutputTextException = OutputTextException & vbLf
                        
                    End If
                End If
            Else
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxDeprecated = CounterSpdxDeprecated + 1
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & "|"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Key & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetSpdxLinksFromReplacedBy(ReplacedBy, LicenseDic) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & vbLf
                        
                    Else
                        CounterNormalDeprecated = CounterNormalDeprecated + 1
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & "|"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Key & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdx = CounterSpdx + 1
                        OutputTextSpdx = OutputTextSpdx & "|"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & Key & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        OutputTextSpdx = OutputTextSpdx & vbLf
                        
                    Else
                        CounterNormal = CounterNormal + 1
                        OutputTextNormal = OutputTextNormal & "|"
                        OutputTextNormal = OutputTextNormal & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormal = OutputTextNormal & " " & Key & " |"
                        OutputTextNormal = OutputTextNormal & vbLf
                        
                    End If
                End If
            End If
    
        End If
    Next
    
    Dim OutputCounter As String
    OutputCounter = _
        "CounterNormal: " & CStr(CounterNormal) & vbLf & _
        "CounterNormalDeprecated: " & CStr(CounterNormalDeprecated) & vbLf & vbLf & _
        "CounterException: " & CStr(CounterException) & vbLf & _
        "CounterExceptionDeprecated: " & CStr(CounterExceptionDeprecated) & vbLf & vbLf & _
        "CounterSpdx: " & CStr(CounterSpdx) & vbLf & _
        "CounterSpdxDeprecated: " & CStr(CounterSpdxDeprecated) & vbLf & vbLf & _
        "CounterSpdxException: " & CStr(CounterSpdxException) & vbLf & _
        "CounterSpdxExceptionDeprecated: " & CStr(CounterSpdxExceptionDeprecated) & vbLf & vbLf & _
        "CounterGeneric: " & CStr(CounterGeneric) & vbLf & _
        "CounterGenericDeprecated: " & CStr(CounterGenericDeprecated) & vbLf & _
        "CounterGenericException: " & CStr(CounterGenericException) & vbLf & _
        "CounterGenericExceptionDeprecated: " & CStr(CounterGenericExceptionDeprecated) & vbLf & vbLf & _
        "CounterAll: " & CStr(CounterAll) & vbLf & vbLf
    
    Dim OutputText As String
    OutputText = _
        OutputTextNormal & vbLf & _
        OutputTextNormalDeprecated & vbLf & _
        OutputTextException & vbLf & _
        OutputTextExceptionDeprecated & vbLf & _
        OutputTextSpdx & vbLf & _
        OutputTextSpdxDeprecated & vbLf & _
        OutputTextSpdxException & vbLf & _
        OutputTextSpdxExceptionDeprecated & vbLf & _
        OutputTextGeneric & vbLf & _
        OutputTextGenericDeprecated & vbLf & _
        OutputTextGenericException & vbLf & _
        OutputTextGenericExceptionDeprecated & vbLf & _
        OutputTextAll
    
    WriteTextFileUTF8 OutputFilePath, OutputCounter & OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveScanCodeTextFileEx4_Core( _
    OutputDirPath As String, InputDirPath As String)
    
    Dim OutputNameDirPath As String
    OutputNameDirPath = OutputDirPath & "\name"
    
    Dim OutputNamePatternDirPath As String
    OutputNamePatternDirPath = OutputDirPath & "\name-pattern"
    
    Dim OutputTextDirPath As String
    OutputTextDirPath = OutputDirPath & "\text"
    
    Dim OutputTextTemplateDirPath As String
    OutputTextTemplateDirPath = OutputDirPath & "\text-template"
    
    Dim OutputTextPatternDirPath As String
    OutputTextPatternDirPath = OutputDirPath & "\text-pattern"
    
    Dim OutputUrlDirPath As String
    OutputUrlDirPath = OutputDirPath & "\url"
    
    Dim OutputUrlPatternDirPath As String
    OutputUrlPatternDirPath = OutputDirPath & "\url-pattern"
    
    MakeDirectory OutputNameDirPath
    MakeDirectory OutputNamePatternDirPath
    MakeDirectory OutputTextDirPath
    MakeDirectory OutputTextTemplateDirPath
    MakeDirectory OutputTextPatternDirPath
    MakeDirectory OutputUrlDirPath
    MakeDirectory OutputUrlPatternDirPath
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(InputDirPath)
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim Key As String
            Dim ShortName As String
            Dim Name As String
            Dim SpdxLicenseKey As String
            Dim ReplacedBy As String
            Dim IsException As Boolean
            Dim IsDeprecated As Boolean
            Dim IsGeneric As Boolean
            Dim IsSpdx As Boolean
            Dim LicenseText As String
            
            'Dim bIgnorableCopyrights As Boolean ' ignorable_copyrights
            'Dim bIgnorableHolders As Boolean ' ignorable_holders
            'Dim bIgnorableAuthors As Boolean ' ignorable_authors
            'Dim bIgnorableUrls As Boolean ' ignorable_urls
            'Dim bIgnorableEmails As Boolean ' ignorable_emails
            
            'Dim MinimumCoverage As String ' minimum_coverage
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            Name = GetItem(FileText, vbLf & "name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            LicenseText = GetLicenseText(FileText)
            
            'bIgnorableCopyrights = TagExists(FileText, "ignorable_copyrights:")
            'bIgnorableHolders = TagExists(FileText, "ignorable_holders:")
            'bIgnorableAuthors = TagExists(FileText, "ignorable_authors:")
            'bIgnorableUrls = TagExists(FileText, "ignorable_urls:")
            'bIgnorableEmails = TagExists(FileText, "ignorable_emails:")
            
            'MinimumCoverage = GetItem(FileText, "minimum_coverage:")
            
            Dim bOutput As Boolean
            bOutput = True
            If IsSpdx Then bOutput = False
            If IsDeprecated Then bOutput = False
            
            If bOutput Then
                Dim NamesText As String
                NamesText = ""
                
                If IsException Then
                    NamesText = NamesText & "SPDX-License-Identifier: XXX WITH " & SpdxLicenseKey & vbLf
                    NamesText = NamesText & "SPDX-Exception-Identifier: " & SpdxLicenseKey & vbLf
                Else
                    NamesText = NamesText & "SPDX-License-Identifier: " & SpdxLicenseKey & vbLf
                End If
                
                If ShortName <> "" Then
                    If IsException Or (InStr(ShortName, "License") > 0) Then
                        NamesText = NamesText & ShortName & vbLf
                    Else
                        NamesText = NamesText & "License: " & ShortName & vbLf
                    End If
                End If
                
                If (Name <> "") And (Name <> ShortName) Then
                    If IsException Or (InStr(Name, "License") > 0) Then
                        NamesText = NamesText & Name & vbLf
                    Else
                        NamesText = NamesText & "License: " & Name & vbLf
                    End If
                End If
                
                WriteTextFileUTF8 OutputNameDirPath & "\" & Key & ".name.txt", NamesText, False
                
                ' ---
                
                Dim NamesPatternText As String
                NamesPatternText = ""
                
                If IsException Then
                    NamesPatternText = NamesPatternText & "SPDX-License-Identifier\W*(?:[A-Za-z0-9\+\-\.]+\W*(?:AND\W*|OR\W*|WITH\W*))*" & Replace(SpdxLicenseKey, ".", "\.") & vbLf
                    NamesPatternText = NamesPatternText & "SPDX-Exception-Identifier: " & Replace(SpdxLicenseKey, ".", "\.") & vbLf
                Else
                    NamesPatternText = NamesPatternText & "SPDX-License-Identifier\W*(?:[A-Za-z0-9\+\-\.]+\W*(?:AND\W*|OR\W*|WITH\W*))*" & Replace(SpdxLicenseKey, ".", "\.") & vbLf
                End If
                
                If ShortName <> "" Then
                    If IsException Or (InStr(ShortName, "License") > 0) Then
                        NamesPatternText = NamesPatternText & ShortName & vbLf
                    Else
                        NamesPatternText = NamesPatternText & "License: " & ShortName & vbLf
                    End If
                End If
                
                If (Name <> "") And (Name <> ShortName) Then
                    If IsException Or (InStr(Name, "License") > 0) Then
                        NamesPatternText = NamesPatternText & Name & vbLf
                    Else
                        NamesPatternText = NamesPatternText & "License: " & Name & vbLf
                    End If
                End If
                
                NamesPatternText = NamesPatternText & vbLf & vbLf & _
                    "examples:" & vbLf & _
                    "(?:MODULE_LICENSE|__license__)\W*" & vbLf & _
                    "Licen[sc]e(?:\W*(magnet.{1,100})?|d\W*(?:as|under)\W*(?:the\W*)?(?:terms\W*(?:and\W*conditions\W*)?of\W*)?)" & vbLf & _
                    vbLf & _
                    "(?<!\w)" & vbLf & _
                    "(?!\w)" & vbLf & _
                    "Licen[sc]e" & vbLf & _
                    "(?:\W*Licen[sc]e|(?!\w))" & vbLf & _
                    vbLf & _
                    "(?:v(?:er(?:sion|\.)?|\.)?\W*)?" & vbLf & _
                    "(?:\.0)?" & vbLf & _
                    "(?:\.0|(?!\.[1-9]))" & vbLf & _
                    vbLf & _
                    "(?<!\w) xxx \W*(?:v(?:er(?:sion|\.)?|\.)?\W*)? xxx \W*Licen[sc]e" & vbLf & _
                    "(?<!\w) xxx \W*Licen[sc]e\W*(?:v(?:er(?:sion|\.)?|\.)?\W*)? xxx" & vbLf & _
                    "v(?:er(?:sion|\.)?|\.)?\W* xxx \W*of\W*(?:the\W*)? xxx \W*Licen[sc]e" & vbLf
                
                WriteTextFileUTF8 OutputNamePatternDirPath & "\" & Key & ".name.pattern.txt", NamesPatternText, False
                
                ' ---
                
                If LicenseText <> "" Then
                    WriteTextFileUTF8 OutputTextDirPath & "\" & Key & ".txt", LicenseText, False
                End If
                
                ' ---
                
                Dim LicenseTextTemplateText As String
                LicenseTextTemplateText = FileText
                
                If LicenseText <> "" Then
                    LicenseTextTemplateText = LicenseTextTemplateText & vbLf & vbLf & _
                        "examples:" & vbLf & _
                        "<<beginOptional>>xxx<<endOptional>>" & vbLf & _
                        "<<var;name=""bullet"";original=""1."";match="".{0,20}"">>" & vbLf & _
                        "<<var;name=""copyright"";original=""xxx"";match=""[\w\W]{1,80}"">>" & vbLf & _
                        "<<var;name=""holder"";original=""xxx"";match=""[\w\W]{1,80}"">>" & vbLf & _
                        "<<var;name=""author"";original=""xxx"";match=""[\w\W]{1,80}"">>" & vbLf & _
                        "<<var;name=""url"";original=""xxx"";match=""[\w\W]{1,80}"">>" & vbLf & _
                        "<<var;name=""email"";original=""xxx"";match=""[\w\W]{1,80}"">>" & vbLf
                        
                    WriteTextFileUTF8 OutputTextTemplateDirPath & "\" & Key & ".template.txt", LicenseTextTemplateText, False
                End If
                
                ' ---
                
                Dim LicenseTextPatternText As String
                LicenseTextPatternText = LicenseText
                
                If LicenseTextPatternText <> "" Then
                    LicenseTextPatternText = LicenseTextPatternText & vbLf & vbLf & _
                        "tool:" & vbLf & _
                        "https://coky-t.github.io/js-parts/js.html/Test_MRegExp_SpdxLicenseText.js.html" & vbLf
                    
                    WriteTextFileUTF8 OutputTextPatternDirPath & "\" & Key & ".pattern.txt", LicenseTextPatternText, False
                End If
                
                ' ---
                
                Dim UrlText As String
                UrlText = "https://scancode-licensedb.aboutcode.org/" & Key & ".html" & vbLf
                
                WriteTextFileUTF8 OutputUrlDirPath & "\" & Key & ".url.txt", UrlText, False
                
                ' ---
                
                Dim UrlPatternText As String
                UrlPatternText = "https?://scancode-licensedb\.aboutcode\.org/" & Replace(Key, ".", "\.") & "\.html" & vbLf
                
                WriteTextFileUTF8 OutputUrlPatternDirPath & "\" & Key & ".url.pattern.txt", UrlPatternText, False
            End If
        End If
    Next
    
    Debug_Print "... Done."
End Sub

Sub Test_SaveScanCodeTextFileEx5_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim CounterNormal As Long
    Dim CounterNormalDeprecated As Long
    Dim CounterException As Long
    Dim CounterExceptionDeprecated As Long
    Dim CounterSpdx As Long
    Dim CounterSpdxDeprecated As Long
    Dim CounterSpdxException As Long
    Dim CounterSpdxExceptionDeprecated As Long
    Dim CounterGeneric As Long
    Dim CounterGenericDeprecated As Long
    Dim CounterGenericException As Long
    Dim CounterGenericExceptionDeprecated As Long
    
    Dim CounterAll As Long
    
    Dim OutputTextNormal As String
    Dim OutputTextNormalDeprecated As String
    Dim OutputTextException As String
    Dim OutputTextExceptionDeprecated As String
    Dim OutputTextSpdx As String
    Dim OutputTextSpdxDeprecated As String
    Dim OutputTextSpdxException As String
    Dim OutputTextSpdxExceptionDeprecated As String
    Dim OutputTextGeneric As String
    Dim OutputTextGenericDeprecated As String
    Dim OutputTextGenericException As String
    Dim OutputTextGenericExceptionDeprecated As String
    
    OutputTextNormal = _
        "### Normal" & vbLf & vbLf & _
        "| Short Name | Key | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-----------|-----------|" & vbLf
        
    OutputTextNormalDeprecated = _
        "### Normal (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-------------|-----------|-----------|" & vbLf
        
    OutputTextException = _
        "### Exception" & vbLf & vbLf & _
        "| Short Name | Key | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-----------|-----------|" & vbLf
        
    OutputTextExceptionDeprecated = _
        "### Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-------------|-----------|-----------|" & vbLf
        
    OutputTextSpdx = _
        "### Spdx" & vbLf & vbLf & _
        "| Short Name | Key | SPDX Identifier | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-----------------|-----------|-----------|" & vbLf
        
    OutputTextSpdxDeprecated = _
        "### Spdx (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | SPDX Identifier | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-------------|-----------------|-----------|-----------|" & vbLf
        
    OutputTextSpdxException = _
        "### Spdx Exception" & vbLf & vbLf & _
        "| Short Name | Key | SPDX Identifier | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-----------------|-----------|-----------|" & vbLf
        
    OutputTextSpdxExceptionDeprecated = _
        "### Spdx Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | SPDX Identifier | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-------------|-----------------|-----------|-----------|" & vbLf
        
    OutputTextGeneric = _
        "### Generic" & vbLf & vbLf & _
        "| Short Name | Key | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-----------|-----------|" & vbLf
        
    OutputTextGenericDeprecated = _
        "### Generic (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-------------|-----------|-----------|" & vbLf
        
    OutputTextGenericException = _
        "### Generic Exception" & vbLf & vbLf & _
        "| Short Name | Key | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-----------|-----------|" & vbLf
        
    OutputTextGenericExceptionDeprecated = _
        "### Generic Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Replaced By | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-------------|-----------|-----------|" & vbLf
        
    Dim OutputTextAll As String
    OutputTextAll = _
        "### All" & vbLf & vbLf & _
        "| Short Name | Key | IsException | IsDeprecated | IsGeneric | IsSpdx | Ignorable | Min. Cov. |" & vbLf & _
        "|------------|-----|-------------|--------------|-----------|--------|-----------|-----------|" & vbLf
    
    Dim LicenseDic As Object
    Set LicenseDic = CreateObject("Scripting.Dictionary")
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim Key As String
            Dim ShortName As String
            Dim SpdxLicenseKey As String
            Dim ReplacedBy As String
            Dim IsException As Boolean
            Dim IsDeprecated As Boolean
            Dim IsGeneric As Boolean
            Dim IsSpdx As Boolean
            
            ' ignorable_copyrights
            ' ignorable_holders
            ' ignorable_authors
            ' ignorable_urls
            ' ignorable_emails
            
            ' minimum_coverage
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            Dim LicenseInfo As ScanCodeLicense
            Set LicenseInfo = New ScanCodeLicense
            With LicenseInfo
                .Key = Key
                .ShortName = ShortName
                .SpdxLicenseKey = SpdxLicenseKey
                .ReplacedBy = ReplacedBy
                .IsException = IsException
                .IsDeprecated = IsDeprecated
                .IsGeneric = IsGeneric
                .IsSpdx = IsSpdx
            End With
            
            LicenseDic.Add Key, LicenseInfo
        End If
    Next
    
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            'Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            'Dim Key As String
            'Dim ShortName As String
            'Dim SpdxLicenseKey As String
            'Dim ReplacedBy As String
            'Dim IsException As Boolean
            'Dim IsDeprecated As Boolean
            'Dim IsGeneric As Boolean
            'Dim IsSpdx As Boolean
            
            Dim bIgnorableCopyrights As Boolean ' ignorable_copyrights
            Dim bIgnorableHolders As Boolean ' ignorable_holders
            Dim bIgnorableAuthors As Boolean ' ignorable_authors
            Dim bIgnorableUrls As Boolean ' ignorable_urls
            Dim bIgnorableEmails As Boolean ' ignorable_emails
            
            Dim Ignorables As String
            
            Dim MinimumCoverage As String ' minimum_coverage
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            bIgnorableCopyrights = TagExists(FileText, "ignorable_copyrights:")
            bIgnorableHolders = TagExists(FileText, "ignorable_holders:")
            bIgnorableAuthors = TagExists(FileText, "ignorable_authors:")
            bIgnorableUrls = TagExists(FileText, "ignorable_urls:")
            bIgnorableEmails = TagExists(FileText, "ignorable_emails:")
            
            Ignorables = ""
            If bIgnorableCopyrights Then
                If Ignorables <> "" Then
                    Ignorables = Ignorables & ", "
                End If
                Ignorables = Ignorables & "copyrights"
            End If
            If bIgnorableHolders Then
                If Ignorables <> "" Then
                    Ignorables = Ignorables & ", "
                End If
                Ignorables = Ignorables & "holders"
            End If
            If bIgnorableAuthors Then
                If Ignorables <> "" Then
                    Ignorables = Ignorables & ", "
                End If
                Ignorables = Ignorables & "authors"
            End If
            If bIgnorableUrls Then
                If Ignorables <> "" Then
                    Ignorables = Ignorables & ", "
                End If
                Ignorables = Ignorables & "urls"
            End If
            If bIgnorableEmails Then
                If Ignorables <> "" Then
                    Ignorables = Ignorables & ", "
                End If
                Ignorables = Ignorables & "emails"
            End If
            
            MinimumCoverage = GetItem(FileText, "minimum_coverage:")
            
            CounterAll = CounterAll + 1
            
            OutputTextAll = OutputTextAll & "|"
            OutputTextAll = OutputTextAll & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
            OutputTextAll = OutputTextAll & " " & Key & " |"
            
            If IsException Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsDeprecated Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsGeneric Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsSpdx Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If Ignorables <> "" Then
                OutputTextAll = OutputTextAll & " " & Ignorables
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If MinimumCoverage <> "" Then
                OutputTextAll = OutputTextAll & " " & MinimumCoverage
            End If
            OutputTextAll = OutputTextAll & " |"
            
            OutputTextAll = OutputTextAll & vbLf
            
            '
            
            If IsGeneric Then
                If IsException Then
                    If IsDeprecated Then
                        CounterGenericExceptionDeprecated = CounterGenericExceptionDeprecated + 1
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & "|"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Key & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Ignorables
                        End If
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & MinimumCoverage
                        End If
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " |"
                        
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & vbLf
                        
                    Else
                        CounterGenericException = CounterGenericException + 1
                        OutputTextGenericException = OutputTextGenericException & "|"
                        OutputTextGenericException = OutputTextGenericException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericException = OutputTextGenericException & " " & Key & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextGenericException = OutputTextGenericException & " " & Ignorables
                        End If
                        OutputTextGenericException = OutputTextGenericException & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextGenericException = OutputTextGenericException & " " & MinimumCoverage
                        End If
                        OutputTextGenericException = OutputTextGenericException & " |"
                        
                        OutputTextGenericException = OutputTextGenericException & vbLf
                        
                    End If
                Else
                    If IsDeprecated Then
                        CounterGenericDeprecated = CounterGenericDeprecated + 1
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & "|"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Key & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Ignorables
                        End If
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & MinimumCoverage
                        End If
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " |"
                        
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & vbLf
                        
                    Else
                        CounterGeneric = CounterGeneric + 1
                        OutputTextGeneric = OutputTextGeneric & "|"
                        OutputTextGeneric = OutputTextGeneric & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGeneric = OutputTextGeneric & " " & Key & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextGeneric = OutputTextGeneric & " " & Ignorables
                        End If
                        OutputTextGeneric = OutputTextGeneric & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextGeneric = OutputTextGeneric & " " & MinimumCoverage
                        End If
                        OutputTextGeneric = OutputTextGeneric & " |"
                        
                        OutputTextGeneric = OutputTextGeneric & vbLf
                        
                    End If
                End If
            ElseIf IsException Then
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxExceptionDeprecated = CounterSpdxExceptionDeprecated + 1
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & "|"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Key & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetSpdxLinksFromReplacedBy(ReplacedBy, LicenseDic) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Ignorables
                        End If
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & MinimumCoverage
                        End If
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " |"
                        
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & vbLf
                        
                    Else
                        CounterExceptionDeprecated = CounterExceptionDeprecated + 1
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & "|"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Key & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Ignorables
                        End If
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & MinimumCoverage
                        End If
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " |"
                        
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdxException = CounterSpdxException + 1
                        OutputTextSpdxException = OutputTextSpdxException & "|"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & Key & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextSpdxException = OutputTextSpdxException & " " & Ignorables
                        End If
                        OutputTextSpdxException = OutputTextSpdxException & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextSpdxException = OutputTextSpdxException & " " & MinimumCoverage
                        End If
                        OutputTextSpdxException = OutputTextSpdxException & " |"
                        
                        OutputTextSpdxException = OutputTextSpdxException & vbLf
                        
                    Else
                        CounterException = CounterException + 1
                        OutputTextException = OutputTextException & "|"
                        OutputTextException = OutputTextException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextException = OutputTextException & " " & Key & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextException = OutputTextException & " " & Ignorables
                        End If
                        OutputTextException = OutputTextException & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextException = OutputTextException & " " & MinimumCoverage
                        End If
                        OutputTextException = OutputTextException & " |"
                        
                        OutputTextException = OutputTextException & vbLf
                        
                    End If
                End If
            Else
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxDeprecated = CounterSpdxDeprecated + 1
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & "|"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Key & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetSpdxLinksFromReplacedBy(ReplacedBy, LicenseDic) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Ignorables
                        End If
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & MinimumCoverage
                        End If
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " |"
                        
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & vbLf
                        
                    Else
                        CounterNormalDeprecated = CounterNormalDeprecated + 1
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & "|"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Key & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Ignorables
                        End If
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & MinimumCoverage
                        End If
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " |"
                        
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdx = CounterSpdx + 1
                        OutputTextSpdx = OutputTextSpdx & "|"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & Key & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextSpdx = OutputTextSpdx & " " & Ignorables
                        End If
                        OutputTextSpdx = OutputTextSpdx & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextSpdx = OutputTextSpdx & " " & MinimumCoverage
                        End If
                        OutputTextSpdx = OutputTextSpdx & " |"
                        
                        OutputTextSpdx = OutputTextSpdx & vbLf
                        
                    Else
                        CounterNormal = CounterNormal + 1
                        OutputTextNormal = OutputTextNormal & "|"
                        OutputTextNormal = OutputTextNormal & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormal = OutputTextNormal & " " & Key & " |"
                        
                        If Ignorables <> "" Then
                            OutputTextNormal = OutputTextNormal & " " & Ignorables
                        End If
                        OutputTextNormal = OutputTextNormal & " |"
                        
                        If MinimumCoverage <> "" Then
                            OutputTextNormal = OutputTextNormal & " " & MinimumCoverage
                        End If
                        OutputTextNormal = OutputTextNormal & " |"
                        
                        OutputTextNormal = OutputTextNormal & vbLf
                        
                    End If
                End If
            End If
    
        End If
    Next
    
    Dim OutputCounter As String
    OutputCounter = _
        "CounterNormal: " & CStr(CounterNormal) & vbLf & _
        "CounterNormalDeprecated: " & CStr(CounterNormalDeprecated) & vbLf & vbLf & _
        "CounterException: " & CStr(CounterException) & vbLf & _
        "CounterExceptionDeprecated: " & CStr(CounterExceptionDeprecated) & vbLf & vbLf & _
        "CounterSpdx: " & CStr(CounterSpdx) & vbLf & _
        "CounterSpdxDeprecated: " & CStr(CounterSpdxDeprecated) & vbLf & vbLf & _
        "CounterSpdxException: " & CStr(CounterSpdxException) & vbLf & _
        "CounterSpdxExceptionDeprecated: " & CStr(CounterSpdxExceptionDeprecated) & vbLf & vbLf & _
        "CounterGeneric: " & CStr(CounterGeneric) & vbLf & _
        "CounterGenericDeprecated: " & CStr(CounterGenericDeprecated) & vbLf & _
        "CounterGenericException: " & CStr(CounterGenericException) & vbLf & _
        "CounterGenericExceptionDeprecated: " & CStr(CounterGenericExceptionDeprecated) & vbLf & vbLf & _
        "CounterAll: " & CStr(CounterAll) & vbLf & vbLf
    
    Dim OutputText As String
    OutputText = _
        OutputTextNormal & vbLf & _
        OutputTextNormalDeprecated & vbLf & _
        OutputTextException & vbLf & _
        OutputTextExceptionDeprecated & vbLf & _
        OutputTextSpdx & vbLf & _
        OutputTextSpdxDeprecated & vbLf & _
        OutputTextSpdxException & vbLf & _
        OutputTextSpdxExceptionDeprecated & vbLf & _
        OutputTextGeneric & vbLf & _
        OutputTextGenericDeprecated & vbLf & _
        OutputTextGenericException & vbLf & _
        OutputTextGenericExceptionDeprecated & vbLf & _
        OutputTextAll
    
    WriteTextFileUTF8 OutputFilePath, OutputCounter & OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveScanCodeTextFileEx6_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim KeyToReleaseDic As Object
    Set KeyToReleaseDic = CreateObject("Scripting.Dictionary")
    
    Dim KeyToDeprecatedDic As Object
    Set KeyToDeprecatedDic = CreateObject("Scripting.Dictionary")
    
    Dim LatestVersion As String
    LatestVersion = ScanCodeToolkitVersions_GetKeyToReleaseDic(KeyToReleaseDic, KeyToDeprecatedDic, DirPath)
    
    ' ---
    
    Const DirName_Prefix As String = "scancode-licenses-"
    
    Dim LatestVersionName As String
    LatestVersionName = Replace(LatestVersion, "v", "")
    
    Dim LatestVersionDirName As String
    LatestVersionDirName = DirName_Prefix & LatestVersionName
    
    Dim LatestVersionDirPath As String
    LatestVersionDirPath = GetFileSystemObject().BuildPath(DirPath, LatestVersionDirName)
    
    ' ---
    
    Dim CounterNormal As Long
    Dim CounterNormalDeprecated As Long
    Dim CounterException As Long
    Dim CounterExceptionDeprecated As Long
    Dim CounterSpdx As Long
    Dim CounterSpdxDeprecated As Long
    Dim CounterSpdxException As Long
    Dim CounterSpdxExceptionDeprecated As Long
    Dim CounterGeneric As Long
    Dim CounterGenericDeprecated As Long
    Dim CounterGenericException As Long
    Dim CounterGenericExceptionDeprecated As Long
    
    Dim CounterAll As Long
    
    Dim OutputTextNormal As String
    Dim OutputTextNormalDeprecated As String
    Dim OutputTextException As String
    Dim OutputTextExceptionDeprecated As String
    Dim OutputTextSpdx As String
    Dim OutputTextSpdxDeprecated As String
    Dim OutputTextSpdxException As String
    Dim OutputTextSpdxExceptionDeprecated As String
    Dim OutputTextGeneric As String
    Dim OutputTextGenericDeprecated As String
    Dim OutputTextGenericException As String
    Dim OutputTextGenericExceptionDeprecated As String
    
    OutputTextNormal = _
        "### Normal" & vbLf & vbLf & _
        "| Short Name | Key | Release |" & vbLf & _
        "|------------|-----|---------|" & vbLf
        
    OutputTextNormalDeprecated = _
        "### Normal (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Release | Deprecated | Replaced By |" & vbLf & _
        "|------------|-----|---------|------------|-------------|" & vbLf
        
    OutputTextException = _
        "### Exception" & vbLf & vbLf & _
        "| Short Name | Key | Release |" & vbLf & _
        "|------------|-----|---------|" & vbLf
        
    OutputTextExceptionDeprecated = _
        "### Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Release | Deprecated | Replaced By |" & vbLf & _
        "|------------|-----|---------|------------|-------------|" & vbLf
        
    OutputTextSpdx = _
        "### Spdx" & vbLf & vbLf & _
        "| Short Name | Key | Release | SPDX Identifier |" & vbLf & _
        "|------------|-----|---------|-----------------|" & vbLf
        
    OutputTextSpdxDeprecated = _
        "### Spdx (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Release | Deprecated | Replaced By | SPDX Identifier |" & vbLf & _
        "|------------|-----|---------|------------|-------------|-----------------|" & vbLf
        
    OutputTextSpdxException = _
        "### Spdx Exception" & vbLf & vbLf & _
        "| Short Name | Key | Release | SPDX Identifier |" & vbLf & _
        "|------------|-----|---------|-----------------|" & vbLf
        
    OutputTextSpdxExceptionDeprecated = _
        "### Spdx Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Release | Deprecated | Replaced By | SPDX Identifier |" & vbLf & _
        "|------------|-----|---------|------------|-------------|-----------------|" & vbLf
        
    OutputTextGeneric = _
        "### Generic" & vbLf & vbLf & _
        "| Short Name | Key | Release |" & vbLf & _
        "|------------|-----|---------|" & vbLf
        
    OutputTextGenericDeprecated = _
        "### Generic (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Release | Deprecated | Replaced By |" & vbLf & _
        "|------------|-----|---------|------------|-------------|" & vbLf
        
    OutputTextGenericException = _
        "### Generic Exception" & vbLf & vbLf & _
        "| Short Name | Key | Release |" & vbLf & _
        "|------------|-----|---------|" & vbLf
        
    OutputTextGenericExceptionDeprecated = _
        "### Generic Exception (Deprecated)" & vbLf & vbLf & _
        "| Short Name | Key | Release | Deprecated | Replaced By |" & vbLf & _
        "|------------|-----|---------|------------|-------------|" & vbLf
        
    Dim OutputTextAll As String
    OutputTextAll = _
        "### All" & vbLf & vbLf & _
        "| Short Name | Key | Release | IsException | IsDeprecated | IsGeneric | IsSpdx |" & vbLf & _
        "|------------|-----|---------|-------------|--------------|-----------|--------|" & vbLf
    
    Dim LicenseDic As Object
    Set LicenseDic = CreateObject("Scripting.Dictionary")
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(LatestVersionDirPath)
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim Key As String
            Dim ShortName As String
            Dim SpdxLicenseKey As String
            Dim ReplacedBy As String
            Dim IsException As Boolean
            Dim IsDeprecated As Boolean
            Dim IsGeneric As Boolean
            Dim IsSpdx As Boolean
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            Dim LicenseInfo As ScanCodeLicense
            Set LicenseInfo = New ScanCodeLicense
            With LicenseInfo
                .Key = Key
                .ShortName = ShortName
                .SpdxLicenseKey = SpdxLicenseKey
                .ReplacedBy = ReplacedBy
                .IsException = IsException
                .IsDeprecated = IsDeprecated
                .IsGeneric = IsGeneric
                .IsSpdx = IsSpdx
            End With
            
            LicenseDic.Add Key, LicenseInfo
        End If
    Next
    
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = "LICENSE" Then ' since 32.0.0rc1
        'If GetFileSystemObject().GetExtensionName(File.Path) = "yml" Then ' until 31.2.6
            Debug_Print File.Name
            
            'Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            'Dim Key As String
            'Dim ShortName As String
            'Dim SpdxLicenseKey As String
            'Dim ReplacedBy As String
            'Dim IsException As Boolean
            'Dim IsDeprecated As Boolean
            'Dim IsGeneric As Boolean
            'Dim IsSpdx As Boolean
            
            Key = GetItem(FileText, "key:")
            ShortName = GetItem(FileText, "short_name:")
            SpdxLicenseKey = GetItem(FileText, "spdx_license_key:")
            ReplacedBy = GetItem(FileText, "replaced_by:" & vbLf & "    - ")
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            Dim Release As String
            Release = ""
            If KeyToReleaseDic.Exists(Key) Then
                Release = KeyToReleaseDic.Item(Key)
            End If
            
            Dim Deprecated As String
            Deprecated = ""
            If KeyToDeprecatedDic.Exists(Key) Then
                Deprecated = KeyToDeprecatedDic.Item(Key)
            End If
            
            CounterAll = CounterAll + 1
            
            OutputTextAll = OutputTextAll & "|"
            OutputTextAll = OutputTextAll & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
            OutputTextAll = OutputTextAll & " " & Key & " |"
            OutputTextAll = OutputTextAll & " " & Release & " |"
            
            If IsException Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsDeprecated Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsGeneric Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            If IsSpdx Then
                OutputTextAll = OutputTextAll & " Yes"
            End If
            OutputTextAll = OutputTextAll & " |"
            
            OutputTextAll = OutputTextAll & vbLf
            
            '
            
            If IsGeneric Then
                If IsException Then
                    If IsDeprecated Then
                        CounterGenericExceptionDeprecated = CounterGenericExceptionDeprecated + 1
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & "|"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Key & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Release & " |"
                        
                        If Deprecated <> "" Then
                            OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & Deprecated
                        End If
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " |"
                        
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextGenericExceptionDeprecated = OutputTextGenericExceptionDeprecated & vbLf
                        
                    Else
                        CounterGenericException = CounterGenericException + 1
                        OutputTextGenericException = OutputTextGenericException & "|"
                        OutputTextGenericException = OutputTextGenericException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericException = OutputTextGenericException & " " & Key & " |"
                        OutputTextGenericException = OutputTextGenericException & " " & Release & " |"
                        OutputTextGenericException = OutputTextGenericException & vbLf
                        
                    End If
                Else
                    If IsDeprecated Then
                        CounterGenericDeprecated = CounterGenericDeprecated + 1
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & "|"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Key & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Release & " |"
                        
                        If Deprecated <> "" Then
                            OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & Deprecated
                        End If
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " |"
                        
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextGenericDeprecated = OutputTextGenericDeprecated & vbLf
                        
                    Else
                        CounterGeneric = CounterGeneric + 1
                        OutputTextGeneric = OutputTextGeneric & "|"
                        OutputTextGeneric = OutputTextGeneric & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextGeneric = OutputTextGeneric & " " & Key & " |"
                        OutputTextGeneric = OutputTextGeneric & " " & Release & " |"
                        OutputTextGeneric = OutputTextGeneric & vbLf
                        
                    End If
                End If
            ElseIf IsException Then
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxExceptionDeprecated = CounterSpdxExceptionDeprecated + 1
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & "|"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Key & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Release & " |"
                        
                        If Deprecated <> "" Then
                            OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & Deprecated
                        End If
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " |"
                        
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & " " & GetSpdxLinksFromReplacedBy(ReplacedBy, LicenseDic) & " |"
                        OutputTextSpdxExceptionDeprecated = OutputTextSpdxExceptionDeprecated & vbLf
                        
                    Else
                        CounterExceptionDeprecated = CounterExceptionDeprecated + 1
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & "|"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Key & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Release & " |"
                        
                        If Deprecated <> "" Then
                            OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & Deprecated
                        End If
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " |"
                        
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextExceptionDeprecated = OutputTextExceptionDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdxException = CounterSpdxException + 1
                        OutputTextSpdxException = OutputTextSpdxException & "|"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & Key & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & Release & " |"
                        OutputTextSpdxException = OutputTextSpdxException & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        OutputTextSpdxException = OutputTextSpdxException & vbLf
                        
                    Else
                        CounterException = CounterException + 1
                        OutputTextException = OutputTextException & "|"
                        OutputTextException = OutputTextException & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextException = OutputTextException & " " & Key & " |"
                        OutputTextException = OutputTextException & " " & Release & " |"
                        OutputTextException = OutputTextException & vbLf
                        
                    End If
                End If
            Else
                If IsDeprecated Then
                    If IsSpdx Then
                        CounterSpdxDeprecated = CounterSpdxDeprecated + 1
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & "|"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Key & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Release & " |"
                        
                        If Deprecated <> "" Then
                            OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & Deprecated
                        End If
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " |"
                        
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & ReplacedBy & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & " " & GetSpdxLinksFromReplacedBy(ReplacedBy, LicenseDic) & " |"
                        OutputTextSpdxDeprecated = OutputTextSpdxDeprecated & vbLf
                        
                    Else
                        CounterNormalDeprecated = CounterNormalDeprecated + 1
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & "|"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Key & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Release & " |"
                        
                        If Deprecated <> "" Then
                            OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & Deprecated
                        End If
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & " " & GetLinksFromReplacedBy(ReplacedBy) & " |"
                        OutputTextNormalDeprecated = OutputTextNormalDeprecated & vbLf
                        
                    End If
                Else
                    If IsSpdx Then
                        CounterSpdx = CounterSpdx + 1
                        OutputTextSpdx = OutputTextSpdx & "|"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & Key & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & Release & " |"
                        OutputTextSpdx = OutputTextSpdx & " " & GetLinkFromSpdxLicenseKey(SpdxLicenseKey) & " |"
                        OutputTextSpdx = OutputTextSpdx & vbLf
                        
                    Else
                        CounterNormal = CounterNormal + 1
                        OutputTextNormal = OutputTextNormal & "|"
                        OutputTextNormal = OutputTextNormal & " " & GetLinkFromShortNameAndKey(ShortName, Key) & " |"
                        OutputTextNormal = OutputTextNormal & " " & Key & " |"
                        OutputTextNormal = OutputTextNormal & " " & Release & " |"
                        OutputTextNormal = OutputTextNormal & vbLf
                        
                    End If
                End If
            End If
    
        End If
    Next
    
    Dim OutputCounter As String
    OutputCounter = _
        "CounterNormal: " & CStr(CounterNormal) & vbLf & _
        "CounterNormalDeprecated: " & CStr(CounterNormalDeprecated) & vbLf & vbLf & _
        "CounterException: " & CStr(CounterException) & vbLf & _
        "CounterExceptionDeprecated: " & CStr(CounterExceptionDeprecated) & vbLf & vbLf & _
        "CounterSpdx: " & CStr(CounterSpdx) & vbLf & _
        "CounterSpdxDeprecated: " & CStr(CounterSpdxDeprecated) & vbLf & vbLf & _
        "CounterSpdxException: " & CStr(CounterSpdxException) & vbLf & _
        "CounterSpdxExceptionDeprecated: " & CStr(CounterSpdxExceptionDeprecated) & vbLf & vbLf & _
        "CounterGeneric: " & CStr(CounterGeneric) & vbLf & _
        "CounterGenericDeprecated: " & CStr(CounterGenericDeprecated) & vbLf & _
        "CounterGenericException: " & CStr(CounterGenericException) & vbLf & _
        "CounterGenericExceptionDeprecated: " & CStr(CounterGenericExceptionDeprecated) & vbLf & vbLf & _
        "CounterAll: " & CStr(CounterAll) & vbLf & vbLf
    
    Dim OutputText As String
    OutputText = _
        OutputTextNormal & vbLf & _
        OutputTextNormalDeprecated & vbLf & _
        OutputTextException & vbLf & _
        OutputTextExceptionDeprecated & vbLf & _
        OutputTextSpdx & vbLf & _
        OutputTextSpdxDeprecated & vbLf & _
        OutputTextSpdxException & vbLf & _
        OutputTextSpdxExceptionDeprecated & vbLf & _
        OutputTextGeneric & vbLf & _
        OutputTextGenericDeprecated & vbLf & _
        OutputTextGenericException & vbLf & _
        OutputTextGenericExceptionDeprecated & vbLf & _
        OutputTextAll
    
    WriteTextFileUTF8 OutputFilePath, OutputCounter & OutputText
    Debug_Print "... Done."
End Sub

Sub Test_SaveScanCodeTextFileEx7_Core( _
    OutputFilePath As String, DirPath As String)
    
    Dim OutputCounterNormal As String
    Dim OutputCounterException As String
    Dim OutputCounterSpdx As String
    Dim OutputCounterSpdxException As String
    Dim OutputCounterGeneric As String
    'Dim OutputCounterGenericException As String
    
    Dim OutputCounterAll As String
    
    OutputCounterNormal = _
        "### Normal, Normal (Deprecated)" & vbLf & vbLf & _
        "| Version | Count | Count (Dep.) |" & vbLf & _
        "|---------|-------|--------------|" & vbLf
        
    OutputCounterException = _
        "### Exception, Exception (Deprecated)" & vbLf & vbLf & _
        "| Version | Count | Count (Dep.) |" & vbLf & _
        "|---------|-------|--------------|" & vbLf
        
    OutputCounterSpdx = _
        "### Spdx, Spdx (Deprecated)" & vbLf & vbLf & _
        "| Version | Count | Count (Dep.) |" & vbLf & _
        "|---------|-------|--------------|" & vbLf
        
    OutputCounterSpdxException = _
        "### Spdx Exception, Spdx Exception (Deprecated)" & vbLf & vbLf & _
        "| Version | Count | Count (Dep.) |" & vbLf & _
        "|---------|-------|--------------|" & vbLf
        
    OutputCounterGeneric = _
        "### Generic, Generic (Deprecated)" & vbLf & vbLf & _
        "| Version | Count | Count (Dep.) | Count (Exc.) | Count (Exc. Dep.) |" & vbLf & _
        "|---------|-------|--------------|--------------|-------------------|" & vbLf
        '"| Version | Count | Count (Dep.) |" & vbLf & _
        '"|---------|-------|--------------|" & vbLf
        
    'OutputCounterGenericException = _
    '    "### Generic Exception, Generic Exception (Deprecated)" & vbLf & vbLf & _
    '    "| Version | Count | Count (Dep.) |" & vbLf & _
    '    "|---------|-------|--------------|" & vbLf
        
    OutputCounterAll = _
        "### All" & vbLf & vbLf & _
        "| Version | Total |   | Nml. | Nml. (Dep.) | Nml. Exc. | Nml. Exc. (Dep.) | Spdx | Spdx (Dep.) | Spdx Exc. | Spdx Exc. (Dep.) | Gen. | Gen. (Dep.) | Gen. Exc. | Gen. Exc. (Dep.) |" & vbLf & _
        "|---------|-------|---|------|-------------|-----------|------------------|------|-------------|-----------|------------------|------|-------------|-----------|------------------|" & vbLf
    
    ScanCodeToolkitVersions_GetCounter _
        OutputCounterNormal, _
        OutputCounterException, _
        OutputCounterSpdx, _
        OutputCounterSpdxException, _
        OutputCounterGeneric, _
        OutputCounterAll, _
        DirPath
        'OutputCounterGenericException, _

    Dim OutputCounter As String
    OutputCounter = _
        OutputCounterNormal & vbLf & _
        OutputCounterException & vbLf & _
        OutputCounterSpdx & vbLf & _
        OutputCounterSpdxException & vbLf & _
        OutputCounterGeneric & vbLf & _
        OutputCounterAll
        'OutputCounterGenericException & vbLf & _

    WriteTextFileUTF8 OutputFilePath, OutputCounter
    Debug_Print "... Done."
End Sub

Sub ScanCodeToolkitVersions_GetCounter( _
    ByRef OutputCounterNormal As String, _
    ByRef OutputCounterException As String, _
    ByRef OutputCounterSpdx As String, _
    ByRef OutputCounterSpdxException As String, _
    ByRef OutputCounterGeneric As String, _
    ByRef OutputCounterAll As String, _
    DirPath As String)
    'ByRef OutputCounterGenericException As String, _

    Const ScanCodeToolkitVersions_FileName1 As String = "scancode-toolkit-versions1.txt"
    Const ScanCodeToolkitVersions_FileName2 As String = "scancode-toolkit-versions2.txt"
    Const ScanCodeToolkitVersions_FileName3 As String = "scancode-toolkit-versions3.txt"
    
    Dim ScanCodeToolkitVersions_FilePath1 As String
    Dim ScanCodeToolkitVersions_FilePath2 As String
    Dim ScanCodeToolkitVersions_FilePath3 As String
    ScanCodeToolkitVersions_FilePath1 = GetFileSystemObject().BuildPath(DirPath, ScanCodeToolkitVersions_FileName1)
    ScanCodeToolkitVersions_FilePath2 = GetFileSystemObject().BuildPath(DirPath, ScanCodeToolkitVersions_FileName2)
    ScanCodeToolkitVersions_FilePath3 = GetFileSystemObject().BuildPath(DirPath, ScanCodeToolkitVersions_FileName3)
    
    Dim ScanCodeToolkitVersions_Text1 As String
    Dim ScanCodeToolkitVersions_Text2 As String
    Dim ScanCodeToolkitVersions_Text3 As String
    ScanCodeToolkitVersions_Text1 = ReadTextFileUTF8(ScanCodeToolkitVersions_FilePath1)
    ScanCodeToolkitVersions_Text2 = ReadTextFileUTF8(ScanCodeToolkitVersions_FilePath2)
    ScanCodeToolkitVersions_Text3 = ReadTextFileUTF8(ScanCodeToolkitVersions_FilePath3)
    
    Dim ScanCodeToolkitVersions1 As Variant
    Dim ScanCodeToolkitVersions2 As Variant
    Dim ScanCodeToolkitVersions3 As Variant
    ScanCodeToolkitVersions1 = Split(ScanCodeToolkitVersions_Text1, vbLf)
    ScanCodeToolkitVersions2 = Split(ScanCodeToolkitVersions_Text2, vbLf)
    ScanCodeToolkitVersions3 = Split(ScanCodeToolkitVersions_Text3, vbLf)
    
    Dim LB3 As Long
    Dim UB3 As Long
    LB3 = LBound(ScanCodeToolkitVersions3)
    UB3 = UBound(ScanCodeToolkitVersions3)
    
    Dim Index3 As Long
    For Index3 = LB3 To UB3
        Debug_Print CStr(ScanCodeToolkitVersions3(Index3))
        ScanCodeToolkitVersions_GetCounter_Core _
            OutputCounterNormal, _
            OutputCounterException, _
            OutputCounterSpdx, _
            OutputCounterSpdxException, _
            OutputCounterGeneric, _
            OutputCounterAll, _
            DirPath, _
            CStr(ScanCodeToolkitVersions3(Index3)), _
            "LICENSE", _
            True
            'OutputCounterGenericException, _

    Next
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(ScanCodeToolkitVersions2)
    UB2 = UBound(ScanCodeToolkitVersions2)
    
    Dim Index2 As Long
    For Index2 = LB2 To UB2
        Debug_Print CStr(ScanCodeToolkitVersions2(Index2))
        ScanCodeToolkitVersions_GetCounter_Core _
            OutputCounterNormal, _
            OutputCounterException, _
            OutputCounterSpdx, _
            OutputCounterSpdxException, _
            OutputCounterGeneric, _
            OutputCounterAll, _
            DirPath, _
            CStr(ScanCodeToolkitVersions2(Index2)), _
            "yml", _
            True
            'OutputCounterGenericException, _

    Next
    
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(ScanCodeToolkitVersions1)
    UB1 = UBound(ScanCodeToolkitVersions1)
    
    Dim Index1 As Long
    For Index1 = LB1 To UB1
        Debug_Print CStr(ScanCodeToolkitVersions1(Index1))
        ScanCodeToolkitVersions_GetCounter_Core _
            OutputCounterNormal, _
            OutputCounterException, _
            OutputCounterSpdx, _
            OutputCounterSpdxException, _
            OutputCounterGeneric, _
            OutputCounterAll, _
            DirPath, _
            CStr(ScanCodeToolkitVersions1(Index1)), _
            "yml", _
            False
            'OutputCounterGenericException, _

    Next
End Sub

Sub ScanCodeToolkitVersions_GetCounter_Core( _
    ByRef OutputCounterNormal As String, _
    ByRef OutputCounterException As String, _
    ByRef OutputCounterSpdx As String, _
    ByRef OutputCounterSpdxException As String, _
    ByRef OutputCounterGeneric As String, _
    ByRef OutputCounterAll As String, _
    DirPath As String, _
    Version As String, _
    Extension As String, _
    bSpdx As Boolean)
    'ByRef OutputCounterGenericException As String, _

    If Version = "" Then Exit Sub
    
    Const DirName_Prefix As String = "scancode-licenses-"
    
    Dim VersionName As String
    VersionName = Replace(Version, "v", "")
    
    Dim VersionDirName As String
    VersionDirName = DirName_Prefix & VersionName
    
    Dim VersionDirPath As String
    VersionDirPath = GetFileSystemObject().BuildPath(DirPath, VersionDirName)
    
    'Debug_Print VersionDirPath
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(VersionDirPath)
    
    Dim CounterNormal As Long
    Dim CounterNormalDeprecated As Long
    Dim CounterException As Long
    Dim CounterExceptionDeprecated As Long
    Dim CounterSpdx As Long
    Dim CounterSpdxDeprecated As Long
    Dim CounterSpdxException As Long
    Dim CounterSpdxExceptionDeprecated As Long
    Dim CounterGeneric As Long
    Dim CounterGenericDeprecated As Long
    Dim CounterGenericException As Long
    Dim CounterGenericExceptionDeprecated As Long
    
    Dim CounterAll As Long
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = Extension Then
            'Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim IsException As Boolean
            Dim IsDeprecated As Boolean
            Dim IsGeneric As Boolean
            Dim IsSpdx As Boolean
            
            IsException = (InStr(FileText, "is_exception: yes") > 0)
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            IsGeneric = (InStr(FileText, "is_generic: yes") > 0)
            IsSpdx = Not (InStr(FileText, "spdx_license_key: LicenseRef-scancode-") > 0)
            
            
            CounterAll = CounterAll + 1
            
            If IsGeneric Then
                If IsException Then
                    If IsDeprecated Then
                        CounterGenericExceptionDeprecated = CounterGenericExceptionDeprecated + 1
                        
                    Else
                        CounterGenericException = CounterGenericException + 1
                        
                    End If
                Else
                    If IsDeprecated Then
                        CounterGenericDeprecated = CounterGenericDeprecated + 1
                        
                    Else
                        CounterGeneric = CounterGeneric + 1
                        
                    End If
                End If
            ElseIf IsException Then
                If IsDeprecated Then
                    If IsSpdx And bSpdx Then
                        CounterSpdxExceptionDeprecated = CounterSpdxExceptionDeprecated + 1
                        
                    Else
                        CounterExceptionDeprecated = CounterExceptionDeprecated + 1
                        
                    End If
                Else
                    If IsSpdx And bSpdx Then
                        CounterSpdxException = CounterSpdxException + 1
                        
                    Else
                        CounterException = CounterException + 1
                        
                    End If
                End If
            Else
                If IsDeprecated Then
                    If IsSpdx And bSpdx Then
                        CounterSpdxDeprecated = CounterSpdxDeprecated + 1
                        
                    Else
                        CounterNormalDeprecated = CounterNormalDeprecated + 1
                        
                    End If
                Else
                    If IsSpdx And bSpdx Then
                        CounterSpdx = CounterSpdx + 1
                        
                    Else
                        CounterNormal = CounterNormal + 1
                        
                    End If
                End If
            End If
        End If
    Next
    
    OutputCounterNormal = OutputCounterNormal & "| " & VersionName & " | " & CStr(CounterNormal) & " | " & CStr(CounterNormalDeprecated) & " |" & vbLf
    OutputCounterException = OutputCounterException & "| " & VersionName & " | " & CStr(CounterException) & " | " & CStr(CounterExceptionDeprecated) & " |" & vbLf
    OutputCounterSpdx = OutputCounterSpdx & "| " & VersionName & " | " & CStr(CounterSpdx) & " | " & CStr(CounterSpdxDeprecated) & " |" & vbLf
    OutputCounterSpdxException = OutputCounterSpdxException & "| " & VersionName & " | " & CStr(CounterSpdxException) & " | " & CStr(CounterSpdxExceptionDeprecated) & " |" & vbLf
    'OutputCounterGeneric = OutputCounterGeneric & "| " & VersionName & " | " & CStr(CounterGeneric) & " | " & CStr(CounterGenericDeprecated) & " |" & vbLf
    OutputCounterGeneric = OutputCounterGeneric & "| " & VersionName & " | " & CStr(CounterGeneric) & " | " & CStr(CounterGenericDeprecated) & " | " & CStr(CounterGenericException) & " | " & CStr(CounterGenericExceptionDeprecated) & " |" & vbLf
    'OutputCounterGenericException = OutputCounterGenericException & "| " & VersionName & " | " & CStr(CounterGenericException) & " | " & CStr(CounterGenericExceptionDeprecated) & " |" & vbLf
    
    OutputCounterAll = OutputCounterAll & "| " & VersionName & " | " & CStr(CounterAll) & " | | " & _
        CStr(CounterNormal) & " | " & CStr(CounterNormalDeprecated) & " | " & _
        CStr(CounterException) & " | " & CStr(CounterExceptionDeprecated) & " | " & _
        CStr(CounterSpdx) & " | " & CStr(CounterSpdxDeprecated) & " | " & _
        CStr(CounterSpdxException) & " | " & CStr(CounterSpdxExceptionDeprecated) & " | " & _
        CStr(CounterGeneric) & " | " & CStr(CounterGenericDeprecated) & " | " & _
        CStr(CounterGenericException) & " | " & CStr(CounterGenericExceptionDeprecated) & " |" & vbLf
End Sub

Function ScanCodeToolkitVersions_GetKeyToReleaseDic( _
    ByRef KeyToReleaseDic As Object, ByRef KeyToDeprecatedDic As Object, DirPath As String) As String
    
    Const ScanCodeToolkitVersions_FileName1 As String = "scancode-toolkit-versions1.txt"
    Const ScanCodeToolkitVersions_FileName2 As String = "scancode-toolkit-versions2.txt"
    Const ScanCodeToolkitVersions_FileName3 As String = "scancode-toolkit-versions3.txt"
    
    Dim ScanCodeToolkitVersions_FilePath1 As String
    Dim ScanCodeToolkitVersions_FilePath2 As String
    Dim ScanCodeToolkitVersions_FilePath3 As String
    ScanCodeToolkitVersions_FilePath1 = GetFileSystemObject().BuildPath(DirPath, ScanCodeToolkitVersions_FileName1)
    ScanCodeToolkitVersions_FilePath2 = GetFileSystemObject().BuildPath(DirPath, ScanCodeToolkitVersions_FileName2)
    ScanCodeToolkitVersions_FilePath3 = GetFileSystemObject().BuildPath(DirPath, ScanCodeToolkitVersions_FileName3)
    
    Dim ScanCodeToolkitVersions_Text1 As String
    Dim ScanCodeToolkitVersions_Text2 As String
    Dim ScanCodeToolkitVersions_Text3 As String
    ScanCodeToolkitVersions_Text1 = ReadTextFileUTF8(ScanCodeToolkitVersions_FilePath1)
    ScanCodeToolkitVersions_Text2 = ReadTextFileUTF8(ScanCodeToolkitVersions_FilePath2)
    ScanCodeToolkitVersions_Text3 = ReadTextFileUTF8(ScanCodeToolkitVersions_FilePath3)
    
    Dim ScanCodeToolkitVersions1 As Variant
    Dim ScanCodeToolkitVersions2 As Variant
    Dim ScanCodeToolkitVersions3 As Variant
    ScanCodeToolkitVersions1 = Split(ScanCodeToolkitVersions_Text1, vbLf)
    ScanCodeToolkitVersions2 = Split(ScanCodeToolkitVersions_Text2, vbLf)
    ScanCodeToolkitVersions3 = Split(ScanCodeToolkitVersions_Text3, vbLf)
    
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(ScanCodeToolkitVersions1)
    UB1 = UBound(ScanCodeToolkitVersions1)
    
    Dim Index1 As Long
    For Index1 = UB1 To LB1 Step -1
        Debug_Print CStr(ScanCodeToolkitVersions1(Index1))
        ScanCodeToolkitVersions_GetKeyToReleaseDic_Core KeyToReleaseDic, KeyToDeprecatedDic, DirPath, CStr(ScanCodeToolkitVersions1(Index1)), "yml"
    Next
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(ScanCodeToolkitVersions2)
    UB2 = UBound(ScanCodeToolkitVersions2)
    
    Dim Index2 As Long
    For Index2 = UB2 To LB2 Step -1
        Debug_Print CStr(ScanCodeToolkitVersions2(Index2))
        ScanCodeToolkitVersions_GetKeyToReleaseDic_Core KeyToReleaseDic, KeyToDeprecatedDic, DirPath, CStr(ScanCodeToolkitVersions2(Index2)), "yml"
    Next
    
    Dim LB3 As Long
    Dim UB3 As Long
    LB3 = LBound(ScanCodeToolkitVersions3)
    UB3 = UBound(ScanCodeToolkitVersions3)
    
    Dim Index3 As Long
    For Index3 = UB3 To LB3 Step -1
        Debug_Print CStr(ScanCodeToolkitVersions3(Index3))
        ScanCodeToolkitVersions_GetKeyToReleaseDic_Core KeyToReleaseDic, KeyToDeprecatedDic, DirPath, CStr(ScanCodeToolkitVersions3(Index3)), "LICENSE"
    Next
    
    Dim LatestVersion As String
    LatestVersion = Replace(ScanCodeToolkitVersions3(LB3), "v", "")
    
    ScanCodeToolkitVersions_GetKeyToReleaseDic = LatestVersion
End Function

Sub ScanCodeToolkitVersions_GetKeyToReleaseDic_Core( _
    ByRef KeyToReleaseDic As Object, ByRef KeyToDeprecatedDic As Object, DirPath As String, Version As String, Extension As String)
    
    If Version = "" Then Exit Sub
    
    Const DirName_Prefix As String = "scancode-licenses-"
    
    Dim VersionName As String
    VersionName = Replace(Version, "v", "")
    
    Dim VersionDirName As String
    VersionDirName = DirName_Prefix & VersionName
    
    Dim VersionDirPath As String
    VersionDirPath = GetFileSystemObject().BuildPath(DirPath, VersionDirName)
    
    'Debug_Print VersionDirPath
    
    Dim Folder As Object
    Set Folder = GetFileSystemObject().GetFolder(VersionDirPath)
    
    Dim File As Object
    For Each File In Folder.Files
        If GetFileSystemObject().GetExtensionName(File.Path) = Extension Then
            'Debug_Print File.Name
            
            Dim FileText As String
            FileText = ReadTextFileUTF8(File.Path)
            'Debug_Print FileText
            
            Dim Key As String
            Dim IsDeprecated As Boolean
            Key = GetItem(FileText, "key:")
            IsDeprecated = (InStr(FileText, "is_deprecated: yes") > 0)
            
            '=== KeyToReleaseDic ===
            If Not KeyToReleaseDic.Exists(Key) Then
                KeyToReleaseDic.Add Key, VersionName
            End If
            
            If IsDeprecated Then
                '=== KeyToDeprecatedDic ===
                If Not KeyToDeprecatedDic.Exists(Key) Then
                    KeyToDeprecatedDic.Add Key, VersionName
                End If
            End If
        End If
    Next
End Sub

Private Function TagExists(FileText, Tag As String) As Boolean
    TagExists = (InStr(FileText, Tag) > 0)
End Function

Private Function GetItem(FileText, Tag As String) As String
    Dim Item As String
    
    Dim TagIndex As Long
    TagIndex = InStr(FileText, Tag)
    
    If Not (TagIndex > 0) Then Exit Function
    
    Dim LfIndex As Long
    LfIndex = InStr(TagIndex + Len(Tag), FileText, vbLf)
    
    If Not (LfIndex > 0) Then Exit Function
    
    Item = Trim(Mid(FileText, TagIndex + Len(Tag), LfIndex - (TagIndex + Len(Tag))))
    
    GetItem = Item
End Function

Private Function GetLicenseText(FileText) As String
    Dim Tag As String
    Tag = "---" & vbLf & vbLf
    
    Dim TagIndex As Long
    TagIndex = InStr(FileText, Tag)
    
    If Not (TagIndex > 0) Then Exit Function
    
    Dim LicenseText As String
    LicenseText = Trim(Right(FileText, Len(FileText) - (TagIndex + Len(Tag) - 1)))
    
    GetLicenseText = LicenseText
End Function

Private Function GetLinkFromShortNameAndKey(ShortName As String, Key As String) As String
    If ShortName = "" Then Exit Function
    If Key = "" Then Exit Function
    GetLinkFromShortNameAndKey = "[" & ShortName & "]" & _
        "(https://scancode-licensedb.aboutcode.org/" & Key & ".html)"
End Function

Private Function GetLinksFromReplacedBy(ReplacedBy As String) As String
    If ReplacedBy = "" Then Exit Function
    
    Dim ReplacedByArray As Variant
    ReplacedByArray = Split(ReplacedBy, " ")
    
    Dim Link1 As String
    Dim Link2 As String
    Dim Link3 As String
    
    Select Case UBound(ReplacedByArray) - LBound(ReplacedByArray)
    Case 2
        Select Case ReplacedByArray(1)
        Case "AND", "OR", "WITH"
            Link1 = "[" & ReplacedByArray(0) & "]" & _
                "(https://scancode-licensedb.aboutcode.org/" & _
                ReplacedByArray(0) & ".html)"
            
            Link2 = "[" & ReplacedByArray(2) & "]" & _
                "(https://scancode-licensedb.aboutcode.org/" & _
                ReplacedByArray(2) & ".html)"
            
            GetLinksFromReplacedBy = _
                Link1 & " " & ReplacedByArray(1) & " " & Link2
            Exit Function
        End Select
    Case 4
        Select Case ReplacedByArray(1)
        Case "AND", "OR", "WITH"
        Select Case ReplacedByArray(3)
        Case "AND", "OR", "WITH"
            Link1 = "[" & ReplacedByArray(0) & "]" & _
                "(https://scancode-licensedb.aboutcode.org/" & _
                ReplacedByArray(0) & ".html)"
            
            Link2 = "[" & ReplacedByArray(2) & "]" & _
                "(https://scancode-licensedb.aboutcode.org/" & _
                ReplacedByArray(2) & ".html)"
            
            Link3 = "[" & ReplacedByArray(4) & "]" & _
                "(https://scancode-licensedb.aboutcode.org/" & _
                ReplacedByArray(4) & ".html)"
            
            GetLinksFromReplacedBy = _
                Link1 & " " & ReplacedByArray(1) & " " & Link2 & _
                " " & ReplacedByArray(3) & " " & Link3
            Exit Function
        End Select
        End Select
    End Select
    
    GetLinksFromReplacedBy = "[" & ReplacedBy & "]" & _
        "(https://scancode-licensedb.aboutcode.org/" & ReplacedBy & ".html)"
End Function

Private Function GetSpdxLinksFromReplacedBy(ReplacedBy As String, LicenseDic As Object) As String
    If ReplacedBy = "" Then Exit Function
    
    Dim ReplacedByArray As Variant
    ReplacedByArray = Split(ReplacedBy, " ")
    
    Dim SpdxLicenseKey1 As String
    Dim SpdxLicenseKey2 As String
    Dim SpdxLicenseKey3 As String
    
    Dim Link1 As String
    Dim Link2 As String
    Dim Link3 As String
    
    Select Case UBound(ReplacedByArray) - LBound(ReplacedByArray)
    Case 2
        Select Case ReplacedByArray(1)
        Case "AND", "OR", "WITH"
            SpdxLicenseKey1 = GetSpdxLicenseKeyFromKey(CStr(ReplacedByArray(0)), LicenseDic)
            SpdxLicenseKey2 = GetSpdxLicenseKeyFromKey(CStr(ReplacedByArray(2)), LicenseDic)
            
            Link1 = GetLinkFromSpdxLicenseKey(SpdxLicenseKey1)
            Link2 = GetLinkFromSpdxLicenseKey(SpdxLicenseKey2)
            
            GetSpdxLinksFromReplacedBy = _
                Link1 & " " & ReplacedByArray(1) & " " & Link2
            Exit Function
        End Select
    Case 4
        Select Case ReplacedByArray(1)
        Case "AND", "OR", "WITH"
        Select Case ReplacedByArray(3)
        Case "AND", "OR", "WITH"
            SpdxLicenseKey1 = GetSpdxLicenseKeyFromKey(CStr(ReplacedByArray(0)), LicenseDic)
            SpdxLicenseKey2 = GetSpdxLicenseKeyFromKey(CStr(ReplacedByArray(2)), LicenseDic)
            SpdxLicenseKey3 = GetSpdxLicenseKeyFromKey(CStr(ReplacedByArray(4)), LicenseDic)
            
            Link1 = GetLinkFromSpdxLicenseKey(SpdxLicenseKey1)
            Link2 = GetLinkFromSpdxLicenseKey(SpdxLicenseKey2)
            Link3 = GetLinkFromSpdxLicenseKey(SpdxLicenseKey3)
            
            GetSpdxLinksFromReplacedBy = _
                Link1 & " " & ReplacedByArray(1) & " " & Link2 & _
                " " & ReplacedByArray(3) & " " & Link3
            Exit Function
        End Select
        End Select
    End Select
    
    Dim SpdxLicenseKey As String
    SpdxLicenseKey = GetSpdxLicenseKeyFromKey(ReplacedBy, LicenseDic)
    
    GetSpdxLinksFromReplacedBy = GetLinkFromSpdxLicenseKey(SpdxLicenseKey)
End Function

Private Function GetSpdxLicenseKeyFromKey(Key As String, LicenseDic As Object) As String
    If Key = "" Then Exit Function
    If LicenseDic Is Nothing Then Exit Function
    If LicenseDic.Count = 0 Then Exit Function
    If Not LicenseDic.Exists(Key) Then Exit Function
    
    Dim LicenseInfo As ScanCodeLicense
    Set LicenseInfo = LicenseDic.Item(Key)
    
    GetSpdxLicenseKeyFromKey = LicenseInfo.SpdxLicenseKey
End Function

Private Function GetLinkFromSpdxLicenseKey(SpdxLicenseKey As String) As String
    If SpdxLicenseKey = "" Then Exit Function
    
    If Left(SpdxLicenseKey, Len("LicenseRef-scancode-")) = "LicenseRef-scancode-" Then
        Dim Key As String
        Key = Right(SpdxLicenseKey, Len(SpdxLicenseKey) - Len("LicenseRef-scancode-"))
        GetLinkFromSpdxLicenseKey = "[(ScanCode) " & Key & "]" & _
            "(https://scancode-licensedb.aboutcode.org/" & Key & ".html)"
    Else
        GetLinkFromSpdxLicenseKey = "[" & SpdxLicenseKey & "]" & _
            "(https://spdx.org/licenses/" & SpdxLicenseKey & ".html)"
    End If
End Function

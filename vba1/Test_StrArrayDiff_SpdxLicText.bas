Attribute VB_Name = "Test_StrArrayDiff_SpdxLicText"
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

Sub Test_StrArrayDiff_SpdxLicenseText_EditDistance()
    ' copy files of
    ' https://github.com/spdx/license-list-data/tree/master/text
    ' to C:\work\data\spdx-license-text.
    
    Dim TargetFilePath As String
    TargetFilePath = "C:\work\data\spdx-license-text\MIT.txt"
    
    Dim SpdxTextDirPath As String
    SpdxTextDirPath = "C:\work\data\spdx-license-text"
    
    Dim TimerObj As CProgressTimer
    Set TimerObj = New CProgressTimer
    
    Test_StrArrayDiff_Files_EditDistance_Core _
        TargetFilePath, SpdxTextDirPath
End Sub

'
' --- Test Core ---
'

Sub Test_StrArrayDiff_Files_EditDistance_Core( _
    TargetFilePath As String, DirPath As String)
    
    Dim TargetText As String
    TargetText = ReadTextFileA(TargetFilePath)
    
    Dim Folder As Scripting.Folder
    Set Folder = GetFileSystemObject().GetFolder(DirPath)
    
    Dim File As Scripting.File
    For Each File In Folder.Files
        Dim FileText As String
        FileText = ReadTextFileA(File.Path)
        Test_StrArrayDiff_File_EditDistance_Core _
            TargetText, FileText, File.Name
    Next
End Sub

Sub Test_StrArrayDiff_File_EditDistance_Core( _
    Str1 As String, Str2 As String, Str2Name As String)
    
    Dim Str1Words As VBScript_RegExp_55.MatchCollection
    Dim Str2Words As VBScript_RegExp_55.MatchCollection
    Set Str1Words = RegExp_Execute(Str1, "(\w+)\W*", False, True, False)
    Set Str2Words = RegExp_Execute(Str2, "(\w+)\W*", False, True, False)
    
    Dim Len1 As Long
    Dim Len2 As Long
    Len1 = Str1Words.Count
    Len2 = Str2Words.Count
    
    Dim StrArray1() As String
    Dim StrArray2() As String
    If Len1 > 0 Then
        ReDim StrArray1(0 To Len1 - 1)
    End If
    If Len2 > 0 Then
        ReDim StrArray2(0 To Len2 - 1)
    End If
    
    Dim Index1 As Long
    For Index1 = 0 To Len1 - 1
        'StrArray1(Index1) = Str1Words.Item(Index1).SubMatches.Item(0) & " "
        StrArray1(Index1) = LCase(Trim(Str1Words.Item(Index1).Value))
    Next
    
    Dim Index2 As Long
    For Index2 = 0 To Len2 - 1
        'StrArray2(Index2) = Str2Words.Item(Index2).SubMatches.Item(0) & " "
        StrArray2(Index2) = LCase(Trim(Str2Words.Item(Index2).Value))
    Next
    
    Dim ED As Long
    ED = EditDistance(StrArray1, StrArray2)
    
    Dim LCD As Long
    LCD = Len1 + Len2 - ED
    
    Debug_Print Str2Name & _
        " ED: " & CStr(ED) & " LCD: " & CStr(LCD) & "/" & CStr(Len1 + Len2)
End Sub

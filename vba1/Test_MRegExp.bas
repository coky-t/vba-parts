Attribute VB_Name = "Test_MRegExp"
Option Explicit

'
' Copyright (c) 2020,2022,2024 Koki Takeyama
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

Public Sub Test_RegExp_Test()
    Test_RegExp_Test_Core "abc 123 xyz #$%", "[a-z]+", True, True
End Sub

Public Sub Test_RegExp_Replace()
    Test_RegExp_Replace_Core _
        "abc 123 xyz #$%", "xxx", "[a-z]+", True, True, True
End Sub

Public Sub Test_RegExp_Execute()
    Test_RegExp_Execute_Core "abc 123 xyz #$%", "([a-z]+)", True, True, True
End Sub

Public Sub Test_RegExp_MatchedValue()
    Test_RegExp_MatchedValue_Core "abc 123 xyz #$%", "([a-z]+)", True, True
End Sub

Public Sub Test_RegExp_ExecuteEx()
    Test_RegExp_ExecuteEx_Core _
        "abc" & vbCrLf & "123" & vbCrLf & "xyz" & vbCrLf & "#$%", _
        "([a-z]+)", _
        True, True, True, vbCrLf
End Sub

'
' --- Test Core ---
'

Public Sub Test_RegExp_Test_Core( _
    SourceString As String, _
    Pattern As String, _
    IgnoreCase As Boolean, _
    MultiLine As Boolean)
    
    Dim Result As Boolean
    Result = RegExp_Test(SourceString, Pattern, IgnoreCase, MultiLine)
    
    Debug_Print "=== RegExp_Test ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "Test - result: " & CStr(Result)
End Sub

Public Sub Test_RegExp_Replace_Core( _
    SourceString As String, _
    ReplaceString As String, _
    Pattern As String, _
    IgnoreCase As Boolean, _
    GlobalMatch As Boolean, _
    MultiLine As Boolean)
    
    Dim Result As String
    Result = _
        RegExp_Replace( _
            SourceString, _
            ReplaceString, _
            Pattern, _
            IgnoreCase, _
            GlobalMatch, _
            MultiLine)
    
    Debug_Print "=== RegExp_Replace ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ReplaceString: " & ReplaceString
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "Replace - result: " & Result
End Sub

Public Sub Test_RegExp_Execute_Core( _
    SourceString As String, _
    Pattern As String, _
    IgnoreCase As Boolean, _
    GlobalMatch As Boolean, _
    MultiLine As Boolean)
    
    Dim Matches As VBScript_RegExp_55.MatchCollection
    Set Matches = _
        RegExp_Execute( _
            SourceString, Pattern, IgnoreCase, GlobalMatch, MultiLine)
    
    Debug_Print "=== RegExp_Execute ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "--- Execute ---"
    
    Debug_Print_Matches Matches
End Sub

Public Sub Test_RegExp_MatchedValue_Core( _
    SourceString As String, _
    Pattern As String, _
    IgnoreCase As Boolean, _
    MultiLine As Boolean)
    
    Dim Result As String
    Result = _
        RegExp_MatchedValue( _
            SourceString, _
            Pattern, _
            IgnoreCase, _
            MultiLine)
    
    Debug_Print "=== RegExp_MatchedValue ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "MatchedValue - result: " & Result
End Sub

Public Sub Test_RegExp_ExecuteEx_Core( _
    SourceString As String, _
    Pattern As String, _
    IgnoreCase As Boolean, _
    GlobalMatch As Boolean, _
    MultiLine As Boolean, _
    LineSeparator As String)
    
    Dim Matches As VBScript_RegExp_55.MatchCollection
    Set Matches = _
        RegExp_Execute( _
            SourceString, Pattern, IgnoreCase, GlobalMatch, MultiLine)
    
    Debug_Print "=== RegExp_Execute ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "--- Execute ---"
    
    Debug_Print_Matches Matches
    
    If Matches Is Nothing Then Exit Sub
    If Matches.Count = 0 Then Exit Sub
    
    Debug_Print "--- LineNumber ---"
    
    Dim Match As VBScript_RegExp_55.Match
    For Each Match In Matches
        Test_RegExp_LineNumber_Core _
            SourceString, Match.FirstIndex, LineSeparator
    Next
End Sub

Public Sub Test_RegExp_LineNumber_Core( _
    SourceString As String, _
    Index As Long, _
    LineSeparator As String)
    
    Dim LineNumber As Long
    LineNumber = RegExp_LineNumber(SourceString, Index, LineSeparator)
    
    Debug_Print "Index: " & CStr(Index) & ", LineNumber: " & CStr(LineNumber)
End Sub

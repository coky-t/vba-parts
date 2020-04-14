Attribute VB_Name = "MRegExp"
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
' Microsoft VBScript Regular Expression 5.5
' - VBScript_RegExp_55.RegExp
'

'
' --- RegExp ---
'

'
' GetRegExp
' - Returns a RegExp object.
'

'
' RegExpObject:
'   Optional. The name of a RegExp object.
'

Public Function GetRegExp( _
    Optional RegExpObject As Object) _
    As Object
    
    If RegExpObject Is Nothing Then
        Set GetRegExp = CreateObject("VBScript.RegExp")
    Else
        Set GetRegExp = RegExpObject
    End If
End Function

'
' === RegExp ===
'

'
' RegExp_Execute
' - Executes a regular expression search against a specified string.
'
' RegExp_Replace
' - Replaces text found in a regular expression search.
'
' RegExp_Test
' - Executes a regular expression search against a specified string
'   and returns a Boolean value that indicates if a pattern match was found.
'

'
' SourceString:
'   Required. The text string upon which the regular expression is executed.
'
' ReplaceString:
'   Required. The replacement text string.
'
' Pattern:
'   Required. Regular string expression being searched for.
'   https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/f97kw5ka(v=vs.84)
'
' IgnoreCase:
'   Optional. The value is False if the search is case-sensitive,
'   True if it is not. Default is False.
'
' GlobalMatch:
'   Optional. The value is True if the search applies to the entire string,
'   False if it does not. Default is False.
'
' MultiLine:
'   Optional. The value is False if the search is single-line mode,
'   True if it is multi-line mode. Default is False.
'
' RegExpObject:
'   Optional. The name of a RegExp object.
'

Public Function RegExp_Execute( _
    SourceString As String, _
    Pattern As String, _
    Optional IgnoreCase As Boolean, _
    Optional GlobalMatch As Boolean, _
    Optional MultiLine As Boolean, _
    Optional RegExpObject As Object) _
    As Object
    
    On Error Resume Next
    
    With GetRegExp(RegExpObject)
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .Global = GlobalMatch
        .MultiLine = MultiLine
        Set RegExp_Execute = .Execute(SourceString)
    End With
End Function

Public Function RegExp_Replace( _
    SourceString As String, _
    ReplaceString As String, _
    Pattern As String, _
    Optional IgnoreCase As Boolean, _
    Optional GlobalMatch As Boolean, _
    Optional MultiLine As Boolean, _
    Optional RegExpObject As Object) _
    As String
    
    On Error Resume Next
    
    With GetRegExp(RegExpObject)
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .Global = GlobalMatch
        .MultiLine = MultiLine
        RegExp_Replace = .Replace(SourceString, ReplaceString)
    End With
End Function

Public Function RegExp_Test( _
    SourceString As String, _
    Pattern As String, _
    Optional IgnoreCase As Boolean, _
    Optional MultiLine As Boolean, _
    Optional RegExpObject As Object) _
    As Boolean
    
    On Error Resume Next
    
    With GetRegExp(RegExpObject)
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .MultiLine = MultiLine
        RegExp_Test = .Test(SourceString)
    End With
End Function

'
' --- Test ---
'

Private Sub Test_RegExp_Test()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim Pattern As String
    Pattern = InputBox("Pattern:")
    If Pattern = "" Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
    Dim Result As Boolean
    Result = RegExp_Test(SourceString, Pattern, IgnoreCase, MultiLine)
    
    Debug_Print "=== RegExp_Test ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "Test - result: " & CStr(Result)
End Sub

Private Sub Test_RegExp_Replace()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim ReplaceString As String
    ReplaceString = InputBox("ReplaceString:")
    If ReplaceString = "" Then Exit Sub
    
    Dim Pattern As String
    Pattern = InputBox("Pattern:")
    If Pattern = "" Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim GlobalMatch As Boolean
    GlobalMatch = (MsgBox("GlobalMatch", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
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

Private Sub Test_RegExp_Execute()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim Pattern As String
    Pattern = InputBox("Pattern:")
    If Pattern = "" Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim GlobalMatch As Boolean
    GlobalMatch = (MsgBox("GlobalMatch", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
    Dim Matches As Object
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
    
    Test_RegExp_Execute_Matches Matches
End Sub

Private Sub Test_RegExp_Execute_Matches( _
    Matches As Object)
    
    If Matches Is Nothing Then
        Debug_Print "Matches: Nothing"
        Exit Sub
    ElseIf Matches.Count = 0 Then
        Debug_Print "Matches: No item"
        Exit Sub
    Else
        Debug_Print "Matches.Count: " & CStr(Matches.Count)
    End If
    
    Dim Match As Object
    For Each Match In Matches
        Test_RegExp_Execute_Match Match
    Next
End Sub

Private Sub Test_RegExp_Execute_Match(Match As Object)
    Debug_Print "---"
    Debug_Print "FirstIndex: " & CStr(Match.FirstIndex)
    Debug_Print "Length: " & CStr(Match.Length)
    Debug_Print "Value: " & Match.Value
    Test_RegExp_Execute_SubMatches Match.SubMatches
End Sub

Private Sub Test_RegExp_Execute_SubMatches( _
    SubMatches As Object)
    
    If SubMatches Is Nothing Then
        Debug_Print "SubMatches: Nothing"
        Exit Sub
    ElseIf SubMatches.Count = 0 Then
        Debug_Print "SubMatches: No item"
        Exit Sub
    Else
        Debug_Print "SubMatches.Count: " & CStr(SubMatches.Count)
    End If
    
    Dim Index As Long
    Dim SubMatch As String
    For Index = 0 To SubMatches.Count - 1
        SubMatch = SubMatches.Item(Index)
        Debug_Print "... " & SubMatch
    Next
End Sub

Private Sub Debug_Print(Str As String)
    Debug.Print Str
End Sub

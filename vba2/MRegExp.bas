Attribute VB_Name = "MRegExp"
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

Public Function GetRegExp() As Object
    Static RegExpObject As Object
    If RegExpObject Is Nothing Then
        Set RegExpObject = CreateObject("VBScript.RegExp")
    End If
    Set GetRegExp = RegExpObject
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

Public Function RegExp_Execute( _
    SourceString As String, _
    Pattern As String, _
    Optional IgnoreCase As Boolean, _
    Optional GlobalMatch As Boolean, _
    Optional MultiLine As Boolean) _
    As Object
    
    On Error Resume Next
    
    With GetRegExp()
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
    Optional MultiLine As Boolean) _
    As String
    
    On Error Resume Next
    
    With GetRegExp()
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
    Optional MultiLine As Boolean) _
    As Boolean
    
    On Error Resume Next
    
    With GetRegExp()
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .MultiLine = MultiLine
        RegExp_Test = .Test(SourceString)
    End With
End Function

Public Function RegExp_MatchedValue( _
    SourceString As String, _
    Pattern As String, _
    Optional IgnoreCase As Boolean, _
    Optional MultiLine As Boolean) _
    As String
    
    Dim Matches As Object
    Set Matches = _
        RegExp_Execute( _
            SourceString, Pattern, IgnoreCase, False, MultiLine)
    
    If Matches Is Nothing Then
        Exit Function
    ElseIf Matches.Count = 0 Then
        Exit Function
    End If
    
    RegExp_MatchedValue = Matches.Item(0).Value
End Function

Public Function RegExp_Matches_Count( _
    SourceString As String, _
    Pattern As String, _
    Optional IgnoreCase As Boolean, _
    Optional MultiLine As Boolean) _
    As Long
    
    Dim Matches As Object
    Set Matches = _
        RegExp_Execute( _
            SourceString, Pattern, IgnoreCase, True, MultiLine)
    
    If Matches Is Nothing Then
        Exit Function
    ElseIf Matches.Count = 0 Then
        Exit Function
    End If
    
    RegExp_Matches_Count = Matches.Count
End Function

Public Function RegExp_LineNumber( _
    SourceString As String, _
    Index As Long, _
    Optional LineSeparator As String = vbCrLf) _
    As Long
    
    Dim SourceStr As String
    SourceStr = Left(SourceString, Index + 1)
    
    Dim LineSep As String
    Select Case LineSeparator
    Case vbCr
        LineSep = "\r"
    Case vbLf
        LineSep = "\n"
    Case Else 'vbCrLf
        LineSep = "\r\n"
    End Select
    
    RegExp_LineNumber = RegExp_Matches_Count(SourceStr, LineSep) + 1
End Function

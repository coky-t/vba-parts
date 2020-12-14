Attribute VB_Name = "MRegExpEx"
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
' === RegExpEx ===
'

'
' RegExp_ParamsList_Execute
' RegExp_Params_Execute
' - Executes a regular expression search against a specified string.
'
' RegExp_ParamsList_Replace
' RegExp_Params_Replace
' - Replaces text found in a regular expression search.
'
' RegExp_ParamsList_Test
' RegExp_Params_Test
' - Executes a regular expression search against a specified string
'   and returns a Boolean value that indicates if a pattern match was found.
'

'
' SourceString:
'   Required. The text string upon which the regular expression is executed.
'
' ParametersListString:
'   For Execute, Test:
'     PatternName(Tab)Pattern(Tab)IgnoreCase(Tab)GlobalMatch(Tab)MultiLine(Newline)
'   For Replace:
'     ReplaceString(Tab)Pattern(Tab)IgnoreCase(Tab)GlobalMatch(Tab)MultiLine(Newline)
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

Public Function RegExp_ParamsList_Execute( _
    SourceString As String, _
    ParametersListString As String) _
    As Collection
    
    If SourceString = "" Then Exit Function
    If ParametersListString = "" Then Exit Function
    
    Dim MatchesCollection As Collection
    Set MatchesCollection = New Collection
    
    Dim ParamsList As Variant
    ParamsList = Split(ParametersListString, vbNewLine)
    
    Dim Index As Long
    For Index = LBound(ParamsList) To UBound(ParamsList)
        Dim Params As String
        Params = CStr(ParamsList(Index))
        
        Dim RegExpMatches As CRegExpMatches
        Set RegExpMatches = RegExp_Params_Execute(SourceString, Params)
        
        If Not RegExpMatches Is Nothing Then
            MatchesCollection.Add RegExpMatches
        End If
    Next
    
    Set RegExp_ParamsList_Execute = MatchesCollection
End Function

Public Function RegExp_Params_Execute( _
    SourceString As String, _
    ParametersString As String) _
    As CRegExpMatches
    
    If SourceString = "" Then Exit Function
    If ParametersString = "" Then Exit Function
    
    Dim Params As Variant
    Params = Split(ParametersString, vbTab)
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(Params)
    UB = UBound(Params)
    
    Dim PatternName As String
    Dim Pattern As String
    Dim IgnoreCase As Boolean
    Dim GlobalMatch As Boolean
    Dim MultiLine As Boolean
    
    PatternName = CStr(Params(LB))
    If LB + 1 <= UB Then Pattern = CStr(Params(LB + 1))
    If LB + 2 <= UB Then IgnoreCase = CBool(Params(LB + 2))
    If LB + 3 <= UB Then GlobalMatch = CBool(Params(LB + 3))
    If LB + 4 <= UB Then MultiLine = CBool(Params(LB + 4))
    
    Dim RegExpMatches As CRegExpMatches
    Set RegExpMatches = New CRegExpMatches
    With RegExpMatches
        .PatternName = PatternName
        Set .Matches = MRegExp.RegExp_Execute( _
            SourceString, _
            Pattern, _
            IgnoreCase, _
            GlobalMatch, _
            MultiLine)
    End With
    
    Set RegExp_Params_Execute = RegExpMatches
End Function

Public Function RegExp_ParamsList_Replace( _
    SourceString As String, _
    ParametersListString As String) _
    As String
    
    RegExp_ParamsList_Replace = SourceString
    
    If SourceString = "" Then Exit Function
    If ParametersListString = "" Then Exit Function
    
    Dim ResultString As String
    ResultString = SourceString
    
    Dim ParamsList As Variant
    ParamsList = Split(ParametersListString, vbNewLine)
    
    Dim Index As Long
    For Index = LBound(ParamsList) To UBound(ParamsList)
        Dim Params As String
        Params = CStr(ParamsList(Index))
        ResultString = RegExp_Params_Replace(ResultString, Params)
    Next
    
    RegExp_ParamsList_Replace = ResultString
End Function

Public Function RegExp_Params_Replace( _
    SourceString As String, _
    ParametersString As String) _
    As String
    
    RegExp_Params_Replace = SourceString
    
    If SourceString = "" Then Exit Function
    If ParametersString = "" Then Exit Function
    
    Dim Params As Variant
    Params = Split(ParametersString, vbTab)
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(Params)
    UB = UBound(Params)
    
    Dim ReplaceString As String
    Dim Pattern As String
    Dim IgnoreCase As Boolean
    Dim GlobalMatch As Boolean
    Dim MultiLine As Boolean
    
    ReplaceString = CStr(Params(LB))
    If LB + 1 <= UB Then Pattern = CStr(Params(LB + 1))
    If LB + 2 <= UB Then IgnoreCase = CBool(Params(LB + 2))
    If LB + 3 <= UB Then GlobalMatch = CBool(Params(LB + 3))
    If LB + 4 <= UB Then MultiLine = CBool(Params(LB + 4))
    
    RegExp_Params_Replace = _
        MRegExp.RegExp_Replace( _
            SourceString, _
            ReplaceString, _
            Pattern, _
            IgnoreCase, _
            GlobalMatch, _
            MultiLine)
End Function

Public Function RegExp_ParamsList_Test( _
    SourceString As String, _
    ParametersListString As String) _
    As String
    
    If SourceString = "" Then Exit Function
    If ParametersListString = "" Then Exit Function
    
    Dim ResultString As String
    
    Dim ParamsList As Variant
    ParamsList = Split(ParametersListString, vbNewLine)
    
    Dim Index As Long
    For Index = LBound(ParamsList) To UBound(ParamsList)
        Dim Params As String
        Params = CStr(ParamsList(Index))
        
        Dim Result As String
        Result = RegExp_Params_Test(SourceString, Params)
        
        If Result <> "" Then
            ResultString = ResultString & Result & vbNewLine
        End If
    Next
    
    RegExp_ParamsList_Test = ResultString
End Function

Public Function RegExp_Params_Test( _
    SourceString As String, _
    ParametersString As String) _
    As String
    
    If SourceString = "" Then Exit Function
    If ParametersString = "" Then Exit Function
    
    Dim ResultString As String
    
    Dim Params As Variant
    Params = Split(ParametersString, vbTab)
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(Params)
    UB = UBound(Params)
    
    Dim PatternName As String
    Dim Pattern As String
    Dim IgnoreCase As Boolean
    'Dim GlobalMatch As Boolean
    Dim MultiLine As Boolean
    
    PatternName = CStr(Params(LB))
    If LB + 1 <= UB Then Pattern = CStr(Params(LB + 1))
    If LB + 2 <= UB Then IgnoreCase = CBool(Params(LB + 2))
    'If LB + 3 <= UB Then GlobalMatch = CBool(Params(LB + 3))
    If LB + 4 <= UB Then MultiLine = CBool(Params(LB + 4))
    
    RegExp_Params_Test = PatternName & vbTab & _
        CStr(MRegExp.RegExp_Test( _
            SourceString, _
            Pattern, _
            IgnoreCase, _
            MultiLine))
End Function

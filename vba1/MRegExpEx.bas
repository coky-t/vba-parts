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
'     (Title)(Tab)Pattern(Tab)IgnoreCase(Tab)GlobalMatch(Tab)MultiLine(Newline)
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
    
    Dim Title As String
    Dim Pattern As String
    Dim IgnoreCase As Boolean
    Dim GlobalMatch As Boolean
    Dim MultiLine As Boolean
    
    Title = CStr(Params(LB))
    If LB + 1 <= UB Then Pattern = CStr(Params(LB + 1))
    If LB + 2 <= UB Then IgnoreCase = CBool(Params(LB + 2))
    If LB + 3 <= UB Then GlobalMatch = CBool(Params(LB + 3))
    If LB + 4 <= UB Then MultiLine = CBool(Params(LB + 4))
    
    Dim RegExpMatches As CRegExpMatches
    Set RegExpMatches = New CRegExpMatches
    With RegExpMatches
        .Title = Title
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
    
    Dim Title As String
    Dim Pattern As String
    Dim IgnoreCase As Boolean
    'Dim GlobalMatch As Boolean
    Dim MultiLine As Boolean
    
    Title = CStr(Params(LB))
    If LB + 1 <= UB Then Pattern = CStr(Params(LB + 1))
    If LB + 2 <= UB Then IgnoreCase = CBool(Params(LB + 2))
    'If LB + 3 <= UB Then GlobalMatch = CBool(Params(LB + 3))
    If LB + 4 <= UB Then MultiLine = CBool(Params(LB + 4))
    
    RegExp_Params_Test = Title & vbTab & _
        CStr(MRegExp.RegExp_Test( _
            SourceString, _
            Pattern, _
            IgnoreCase, _
            MultiLine))
End Function

'
' --- Test ---
'

Private Sub Test_RegExp_ParamsList_Test()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim PatternCollection As Collection
    Set PatternCollection = New Collection
    Do While True
        Dim Pattern As String
        Pattern = InputBox("Pattern:")
        If Pattern = "" Then Exit Do
        
        PatternCollection.Add Pattern
    Loop
    If PatternCollection.Count = 0 Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
    Dim ParamsList As String
    Dim Index As Long
    For Index = 1 To PatternCollection.Count
        ParamsList = ParamsList & _
            "Pattern" & CStr(Index) & vbTab & _
            PatternCollection.Item(Index) & vbTab & _
            CStr(IgnoreCase) & vbTab & _
            "False" & vbTab & _
            CStr(MultiLine) & vbNewLine
    Next
    
    Dim Result As String
    Result = RegExp_ParamsList_Test(SourceString, ParamsList)
    
    Debug_Print "=== RegExp_ParamsList_Test ==="
    Debug_Print "SourceString: " & SourceString
    For Index = 1 To PatternCollection.Count
        Debug_Print _
            "Pattern" & CStr(Index) & ": " & PatternCollection.Item(Index)
    Next
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "Test - result: "
    Debug_Print Result
End Sub

Private Sub Test_RegExp_Params_Test()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim Title As String
    Title = InputBox("Title:")
    If Title = "" Then Exit Sub
    
    Dim Pattern As String
    Pattern = InputBox("Pattern:")
    If Pattern = "" Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
    Dim Params As String
    Params = _
        Title & vbTab & _
        Pattern & vbTab & _
        CStr(IgnoreCase) & vbTab & _
        "False" & vbTab & _
        CStr(MultiLine)
    
    Dim Result As String
    Result = RegExp_Params_Test(SourceString, Params)
    
    Debug_Print "=== RegExp_Params_Test ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Title: " & Title
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "Test - result: " & Result
End Sub

Private Sub Test_RegExp_ParamsList_Replace()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim PatternCollection As Collection
    Set PatternCollection = New Collection
    Do While True
        Dim ReplaceString As String
        ReplaceString = InputBox("ReplaceString:")
        If ReplaceString = "" Then Exit Do
        
        Dim Pattern As String
        Pattern = InputBox("Pattern:")
        If Pattern = "" Then Exit Do
        
        PatternCollection.Add ReplaceString & vbTab & Pattern
    Loop
    If PatternCollection.Count = 0 Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim GlobalMatch As Boolean
    GlobalMatch = (MsgBox("GlobalMatch", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
    Dim ParamsList As String
    Dim Index As Long
    For Index = 1 To PatternCollection.Count
        ParamsList = ParamsList & _
            PatternCollection.Item(Index) & vbTab & _
            CStr(IgnoreCase) & vbTab & _
            CStr(GlobalMatch) & vbTab & _
            CStr(MultiLine) & vbNewLine
    Next
    
    Dim Result As String
    Result = RegExp_ParamsList_Replace(SourceString, ParamsList)
    
    Debug_Print "=== RegExp_ParamsList_Replace ==="
    Debug_Print "SourceString: " & SourceString
    For Index = 1 To PatternCollection.Count
        Debug_Print _
            "ReplaceString and Pattern " & CStr(Index) & ": " & _
            PatternCollection.Item(Index)
    Next
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "Replace - result: " & Result
End Sub

Private Sub Test_RegExp_Params_Replace()
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
    
    Dim Params As String
    Params = _
        ReplaceString & vbTab & _
        Pattern & vbTab & _
        CStr(IgnoreCase) & vbTab & _
        CStr(GlobalMatch) & vbTab & _
        CStr(MultiLine)
    
    Dim Result As String
    Result = RegExp_Params_Replace(SourceString, Params)
    
    Debug_Print "=== RegExp_Params_Replace ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ReplaceString: " & ReplaceString
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "Replace - result: " & Result
End Sub

Private Sub Test_RegExp_ParamsList_Execute()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim PatternCollection As Collection
    Set PatternCollection = New Collection
    Do While True
        Dim Pattern As String
        Pattern = InputBox("Pattern:")
        If Pattern = "" Then Exit Do
        
        PatternCollection.Add Pattern
    Loop
    If PatternCollection.Count = 0 Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim GlobalMatch As Boolean
    GlobalMatch = (MsgBox("GlobalMatch", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
    Dim ParamsList As String
    Dim Index As Long
    For Index = 1 To PatternCollection.Count
        ParamsList = ParamsList & _
            "Pattern" & CStr(Index) & vbTab & _
            PatternCollection.Item(Index) & vbTab & _
            CStr(IgnoreCase) & vbTab & _
            CStr(GlobalMatch) & vbTab & _
            CStr(MultiLine) & vbNewLine
    Next
    
    Dim REMCollection As Collection
    Set REMCollection = RegExp_ParamsList_Execute(SourceString, ParamsList)
    
    Debug_Print "=== RegExp_ParamsList_Execute ==="
    Debug_Print "SourceString: " & SourceString
    For Index = 1 To PatternCollection.Count
        Debug_Print _
            "Pattern" & CStr(Index) & ": " & PatternCollection.Item(Index)
    Next
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "--- Execute ---"
    
    Debug_Print_RegExpMatchesCollection REMCollection
End Sub

Private Sub Test_RegExp_Params_Execute()
    Dim SourceString As String
    SourceString = InputBox("SourceString:")
    If SourceString = "" Then Exit Sub
    
    Dim Title As String
    Title = InputBox("Title:")
    If Title = "" Then Exit Sub
    
    Dim Pattern As String
    Pattern = InputBox("Pattern:")
    If Pattern = "" Then Exit Sub
    
    Dim IgnoreCase As Boolean
    IgnoreCase = (MsgBox("IgnoreCase", vbYesNo) = vbYes)
    
    Dim GlobalMatch As Boolean
    GlobalMatch = (MsgBox("GlobalMatch", vbYesNo) = vbYes)
    
    Dim MultiLine As Boolean
    MultiLine = (MsgBox("MultiLine", vbYesNo) = vbYes)
    
    Dim Params As String
    Params = _
        Title & vbTab & _
        Pattern & vbTab & _
        CStr(IgnoreCase) & vbTab & _
        CStr(GlobalMatch) & vbTab & _
        CStr(MultiLine)
    
    Dim RegExpMatches As CRegExpMatches
    Set RegExpMatches = RegExp_Params_Execute(SourceString, Params)
    
    Debug_Print "=== RegExp_Params_Execute ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Title: " & Title
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "--- Execute ---"
    
    Debug_Print_RegExpMatches RegExpMatches
End Sub

Private Sub Debug_Print_RegExpMatchesCollection( _
    RegExpMatchesCollection As Collection)
    
    If RegExpMatchesCollection Is Nothing Then
        Debug_Print "RegExpMatchesCollection: Nothing"
    ElseIf RegExpMatchesCollection.Count = 0 Then
        Debug_Print "RegExpMatchesCollection: No item"
    Else
        Dim RegExpMatches As CRegExpMatches
        For Each RegExpMatches In RegExpMatchesCollection
            Debug_Print_RegExpMatches RegExpMatches
            Debug_Print "---"
        Next
    End If
End Sub

Private Sub Debug_Print_RegExpMatches(RegExpMatches As CRegExpMatches)
    If RegExpMatches Is Nothing Then
        Debug_Print "RegExpMatches: Nothing"
    Else
        Debug_Print "Title: " & RegExpMatches.Title
        Debug_Print_Matches RegExpMatches.Matches
    End If
End Sub

Private Sub Debug_Print_Matches( _
    Matches As VBScript_RegExp_55.MatchCollection)
    
    If Matches Is Nothing Then
        Debug_Print "Matches: Nothing"
        Exit Sub
    ElseIf Matches.Count = 0 Then
        Debug_Print "Matches: No item"
        Exit Sub
    Else
        Debug_Print "Matches.Count: " & CStr(Matches.Count)
    End If
    
    Dim Match As VBScript_RegExp_55.Match
    For Each Match In Matches
        Debug_Print_Match Match
    Next
End Sub

Private Sub Debug_Print_Match(Match As VBScript_RegExp_55.Match)
    Debug_Print "---"
    Debug_Print "FirstIndex: " & CStr(Match.FirstIndex)
    Debug_Print "Length: " & CStr(Match.Length)
    Debug_Print "Value: " & Match.Value
    Debug_Print_SubMatches Match.SubMatches
End Sub

Private Sub Debug_Print_SubMatches( _
    SubMatches As VBScript_RegExp_55.SubMatches)
    
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

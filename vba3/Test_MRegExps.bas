Attribute VB_Name = "Test_MRegExps"
Option Explicit

'
' Copyright (c) 2020,2021 Koki Takeyama
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

Public Sub Test_CRegExps_Test()
    Test_CRegExps_Test_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine & _
        "num" & vbTab & _
            "[0-9]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine
End Sub

Public Sub Test_CRegExp_Test()
    Test_CRegExp_Test_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True"
End Sub

Public Sub Test_CRegExps_Replace()
    Test_CRegExps_Replace_Core _
        "abc 123 xyz #$%", _
        "xxx" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine & _
        "999" & vbTab & _
            "[0-9]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine
End Sub

Public Sub Test_CRegExp_Replace()
    Test_CRegExp_Replace_Core _
        "abc 123 xyz #$%", _
        "xxx" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True"
End Sub

Public Sub Test_CRegExps_Execute()
    Test_CRegExps_Execute_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "([a-z]+)" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine & _
        "num" & vbTab & _
            "[0-9]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine
End Sub

Public Sub Test_CRegExp_Execute()
    Test_CRegExp_Execute_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "([a-z]+)" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True"
End Sub

Public Sub Test_CRegExp_GetCRegExpMatches()
    Test_CRegExp_GetCRegExpMatches_Core _
        "abc 123 xyz #$%", "alpha", "([a-z]+)", True, True, True
End Sub

Public Sub Test_CRegExp_GetCRegExpMatch()
    Test_CRegExp_GetCRegExpMatch_Core _
        "abc 123 xyz #$%", "alpha", "([a-z]+)", True, False, False
End Sub

'
' --- Test Core ---
'

Public Sub Test_CRegExps_Test_Core( _
    SourceString, _
    ParamsList)
    
    Dim CRegExps_
    Set CRegExps_ = GetCRegExps(ParamsList)
    
    Dim Result
    Result = CRegExps_Test(CRegExps_, SourceString)
    
    Debug_Print "=== CRegExps_Test ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ParamsList: "
    Debug_Print ParamsList
    Debug_Print "Test - result: "
    Debug_Print Result
End Sub

Public Sub Test_CRegExp_Test_Core( _
    SourceString, _
    Params)
    
    Dim CRegExp_
    Set CRegExp_ = GetCRegExp(Params)
    
    Dim Result
    Result = CRegExp_Test(CRegExp_, SourceString)
    
    Debug_Print "=== CRegExp_Test ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Params: " & Params
    Debug_Print "Test - result: " & Result
End Sub

Public Sub Test_CRegExps_Replace_Core( _
    SourceString, _
    ParamsList)
    
    Dim CRegExps_
    Set CRegExps_ = GetCRegExps(ParamsList)
    
    Dim Result
    Result = CRegExps_Replace(CRegExps_, SourceString)
    
    Debug_Print "=== CRegExps_Replace ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ParamsList: "
    Debug_Print ParamsList
    Debug_Print "Replace - result: " & Result
End Sub

Public Sub Test_CRegExp_Replace_Core( _
    SourceString, _
    Params)
    
    Dim CRegExp_
    Set CRegExp_ = GetCRegExp(Params)
    
    Dim Result
    Result = CRegExp_Replace(CRegExp_, SourceString)
    
    Debug_Print "=== CRegExp_Replace ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Params: " & Params
    Debug_Print "Replace - result: " & Result
End Sub

Public Sub Test_CRegExps_Execute_Core( _
    SourceString, _
    ParamsList)
    
    Dim CRegExps_
    Set CRegExps_ = GetCRegExps(ParamsList)
    
    Dim REMCollection
    Set REMCollection = CRegExps_Execute(CRegExps_, SourceString)
    
    Debug_Print "=== CRegExps_Execute ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ParamsList: "
    Debug_Print ParamsList
    Debug_Print "--- Execute ---"
    
    Debug_Print_RegExpMatchesCollection REMCollection
End Sub

Public Sub Test_CRegExp_Execute_Core( _
    SourceString, _
    Params)
    
    Dim CRegExp_
    Set CRegExp_ = GetCRegExp(Params)
    
    Dim RegExpMatches
    Set RegExpMatches = CRegExp_Execute(CRegExp_, SourceString)
    
    Debug_Print "=== CRegExp_Execute ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Params: " & Params
    Debug_Print "--- Execute ---"
    
    Debug_Print_RegExpMatches RegExpMatches
End Sub

Public Sub Test_CRegExp_GetCRegExpMatches_Core( _
    SourceString, _
    PatternName, _
    Pattern, _
    IgnoreCase, _
    GlobalMatch, _
    MultiLine)
    
    Dim CRegExp_
    Set CRegExp_ = New CRegExp
    With CRegExp_
        .PatternName = PatternName
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .GlobalMatch = GlobalMatch
        .MultiLine = MultiLine
    End With
    
    Dim CRegExpMatches_
    Set CRegExpMatches_ = CRegExp_.GetCRegExpMatches(SourceString)
    
    Debug_Print "=== RegExp_GetCRegExpMatches ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "PatternName: " & PatternName
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "--- GetCRegExpMatches ---"
    
    Debug_Print "PatternName: " & CRegExpMatches_.PatternName
    Debug_Print_Matches CRegExpMatches_.Matches
End Sub

Public Sub Test_CRegExp_GetCRegExpMatch_Core( _
    SourceString, _
    PatternName, _
    Pattern, _
    IgnoreCase, _
    GlobalMatch, _
    MultiLine)
    
    Dim CRegExp_
    Set CRegExp_ = New CRegExp
    With CRegExp_
        .PatternName = PatternName
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .GlobalMatch = GlobalMatch
        .MultiLine = MultiLine
    End With
    
    Dim CRegExpMatch_
    Set CRegExpMatch_ = CRegExp_.GetCRegExpMatch(SourceString)
    
    Debug_Print "=== RegExp_GetCRegExpMatch ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "PatternName: " & PatternName
    Debug_Print "Pattern: " & Pattern
    Debug_Print "IgnoreCase: " & CStr(IgnoreCase)
    Debug_Print "GlobalMatch: " & CStr(GlobalMatch)
    Debug_Print "MultiLine: " & CStr(MultiLine)
    Debug_Print "--- GetCRegExpMatch ---"
    
    Debug_Print "PatternName: " & CRegExpMatch_.PatternName
    Debug_Print_Match CRegExpMatch_.Match
End Sub

Public Sub Debug_Print_RegExpMatchesCollection( _
    RegExpMatchesCollection)
    
    If RegExpMatchesCollection Is Nothing Then
        Debug_Print "RegExpMatchesCollection: Nothing"
    ElseIf RegExpMatchesCollection.Count = 0 Then
        Debug_Print "RegExpMatchesCollection: No item"
    Else
        Dim RegExpMatches
        For Each RegExpMatches In RegExpMatchesCollection
            Debug_Print_RegExpMatches RegExpMatches
            Debug_Print "---"
        Next
    End If
End Sub

Public Sub Debug_Print_RegExpMatches(RegExpMatches)
    If RegExpMatches Is Nothing Then
        Debug_Print "RegExpMatches: Nothing"
    Else
        Debug_Print "PatternName: " & RegExpMatches.PatternName
        Debug_Print_Matches RegExpMatches.Matches
    End If
End Sub

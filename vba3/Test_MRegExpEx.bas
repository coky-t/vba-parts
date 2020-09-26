Attribute VB_Name = "Test_MRegExpEx"
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
' --- Test ---
'

Public Sub Test_RegExp_ParamsList_Test()
    Test_RegExp_ParamsList_Test_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine & _
        "num" & vbTab & _
            "[0-9]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine
End Sub

Public Sub Test_RegExp_Params_Test()
    Test_RegExp_Params_Test_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True"
End Sub

Public Sub Test_RegExp_ParamsList_Replace()
    Test_RegExp_ParamsList_Replace_Core _
        "abc 123 xyz #$%", _
        "xxx" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine & _
        "999" & vbTab & _
            "[0-9]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine
End Sub

Public Sub Test_RegExp_Params_Replace()
    Test_RegExp_Params_Replace_Core _
        "abc 123 xyz #$%", _
        "xxx" & vbTab & _
            "[a-z]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True"
End Sub

Public Sub Test_RegExp_ParamsList_Execute()
    Test_RegExp_ParamsList_Execute_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "([a-z]+)" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine & _
        "num" & vbTab & _
            "[0-9]+" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True" & vbNewLine
End Sub

Public Sub Test_RegExp_Params_Execute()
    Test_RegExp_Params_Execute_Core _
        "abc 123 xyz #$%", _
        "alpha" & vbTab & _
            "([a-z]+)" & vbTab & _
            "True" & vbTab & "True" & vbTab & "True"
End Sub

'
' --- Test Core ---
'

Public Sub Test_RegExp_ParamsList_Test_Core( _
    SourceString, _
    ParamsList)
    
    Dim Result
    Result = RegExp_ParamsList_Test(SourceString, ParamsList)
    
    Debug_Print "=== RegExp_ParamsList_Test ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ParamsList: "
    Debug_Print ParamsList
    Debug_Print "Test - result: "
    Debug_Print Result
End Sub

Public Sub Test_RegExp_Params_Test_Core( _
    SourceString, _
    Params)
    
    Dim Result
    Result = RegExp_Params_Test(SourceString, Params)
    
    Debug_Print "=== RegExp_Params_Test ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Params: " & Params
    Debug_Print "Test - result: " & Result
End Sub

Public Sub Test_RegExp_ParamsList_Replace_Core( _
    SourceString, _
    ParamsList)
    
    Dim Result
    Result = RegExp_ParamsList_Replace(SourceString, ParamsList)
    
    Debug_Print "=== RegExp_ParamsList_Replace ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ParamsList: "
    Debug_Print ParamsList
    Debug_Print "Replace - result: " & Result
End Sub

Public Sub Test_RegExp_Params_Replace_Core( _
    SourceString, _
    Params)
    
    Dim Result
    Result = RegExp_Params_Replace(SourceString, Params)
    
    Debug_Print "=== RegExp_Params_Replace ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Params: " & Params
    Debug_Print "Replace - result: " & Result
End Sub

Public Sub Test_RegExp_ParamsList_Execute_Core( _
    SourceString, _
    ParamsList)
    
    Dim REMCollection
    Set REMCollection = RegExp_ParamsList_Execute(SourceString, ParamsList)
    
    Debug_Print "=== RegExp_ParamsList_Execute ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "ParamsList: "
    Debug_Print ParamsList
    Debug_Print "--- Execute ---"
    
    Debug_Print_RegExpMatchesCollection REMCollection
End Sub

Public Sub Test_RegExp_Params_Execute_Core( _
    SourceString, _
    Params)
    
    Dim RegExpMatches
    Set RegExpMatches = RegExp_Params_Execute(SourceString, Params)
    
    Debug_Print "=== RegExp_Params_Execute ==="
    Debug_Print "SourceString: " & SourceString
    Debug_Print "Params: " & Params
    Debug_Print "--- Execute ---"
    
    Debug_Print_RegExpMatches RegExpMatches
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
        Debug_Print "Title: " & RegExpMatches.Title
        Debug_Print_Matches RegExpMatches.Matches
    End If
End Sub

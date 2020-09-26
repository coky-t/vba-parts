Attribute VB_Name = "DebugPrintMatches"
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

Public Sub Debug_Print_Matches( _
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
        Debug_Print_Match Match
    Next
End Sub

Public Sub Debug_Print_Match(Match As Object)
    Debug_Print "---"
    Debug_Print "FirstIndex: " & CStr(Match.FirstIndex)
    Debug_Print "Length: " & CStr(Match.Length)
    Debug_Print "Value: " & Match.Value
    Debug_Print_SubMatches Match.SubMatches
End Sub

Public Sub Debug_Print_SubMatches( _
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

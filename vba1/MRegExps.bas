Attribute VB_Name = "MRegExps"
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
' --- CRegExps ---
'

Public Function GetCRegExps(ParamsListString As String) As Collection
    If ParamsListString = "" Then Exit Function
    
    Dim CRegExps As Collection
    Set CRegExps = New Collection
    
    Dim ParamsList As Variant
    ParamsList = Split(ParamsListString, vbNewLine)
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(ParamsList)
    UB = UBound(ParamsList)
    
    Dim Index As Long
    For Index = LB To UB
        Dim Params As String
        Params = ParamsList(Index)
        
        Dim CRegExp_ As CRegExp
        Set CRegExp_ = GetCRegExp(Params)
        
        If Not CRegExp_ Is Nothing Then
            CRegExps.Add CRegExp_
        End If
    Next
    
    Set GetCRegExps = CRegExps
End Function

Public Function GetCRegExp(ParamsString As String) As CRegExp
    If ParamsString = "" Then Exit Function
    
    Dim Params As Variant
    Params = Split(ParamsString, vbTab)
    
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
    
    Dim CRegExp_ As CRegExp
    Set CRegExp_ = New CRegExp
    With CRegExp_
        .PatternName = PatternName
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .GlobalMatch = GlobalMatch
        .MultiLine = MultiLine
    End With
    
    Set GetCRegExp = CRegExp_
End Function

Public Sub CRegExps_LetOptionals( _
    ByRef CRegExps As Collection, _
    Optional IgnoreCase As Boolean, _
    Optional GlobalMatch As Boolean, _
    Optional MultiLine As Boolean)
    
    If CRegExps Is Nothing Then Exit Sub
    If CRegExps.Count = 0 Then Exit Sub
    
    Dim CRegExp_ As CRegExp
    For Each CRegExp_ In CRegExps
        With CRegExp_
            .IgnoreCase = IgnoreCase
            .GlobalMatch = GlobalMatch
            .MultiLine = MultiLine
        End With
    Next
End Sub

Public Function CRegExps_Execute( _
    CRegExps As Collection, _
    SourceString As String) _
    As Collection
    
    If CRegExps Is Nothing Then Exit Function
    If CRegExps.Count = 0 Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim MatchesCollection As Collection
    Set MatchesCollection = New Collection
    
    Dim CRegExp_ As CRegExp
    For Each CRegExp_ In CRegExps
        Dim RegExpMatches As CRegExpMatches
        Set RegExpMatches = CRegExp_Execute(CRegExp_, SourceString)
        If Not RegExpMatches Is Nothing Then
            MatchesCollection.Add RegExpMatches
        End If
    Next
    
    Set CRegExps_Execute = MatchesCollection
End Function

Public Function CRegExp_Execute(CRegExp_ As CRegExp, SourceString As String) _
    As CRegExpMatches
    
    If CRegExp_ Is Nothing Then Exit Function
    If CRegExp_.PatternName = "" Then Exit Function
    If CRegExp_.Pattern = "" Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim Matches As VBScript_RegExp_55.MatchCollection
    Set Matches = CRegExp_.Execute(SourceString)
    If Matches Is Nothing Then Exit Function
    If Matches.Count = 0 Then Exit Function
    
    Dim RegExpMatches As CRegExpMatches
    Set RegExpMatches = New CRegExpMatches
    With RegExpMatches
        .PatternName = CRegExp_.PatternName
        Set .Matches = Matches
    End With
    
    Set CRegExp_Execute = RegExpMatches
End Function

Public Function CRegExps_Replace( _
    CRegExps As Collection, _
    SourceString As String) _
    As String
    
    CRegExps_Replace = SourceString
    
    If CRegExps Is Nothing Then Exit Function
    If CRegExps.Count = 0 Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim ResultString As String
    ResultString = SourceString
    
    Dim CRegExp_ As CRegExp
    For Each CRegExp_ In CRegExps
        ResultString = CRegExp_Replace(CRegExp_, ResultString)
    Next
    
    CRegExps_Replace = ResultString
End Function

Public Function CRegExp_Replace(CRegExp_ As CRegExp, SourceString As String) _
    As String
    
    CRegExp_Replace = SourceString
    
    If CRegExp_ Is Nothing Then Exit Function
    If CRegExp_.PatternName = "" Then Exit Function
    If CRegExp_.Pattern = "" Then Exit Function
    If SourceString = "" Then Exit Function
    
    CRegExp_Replace = CRegExp_.Replace(SourceString, CRegExp_.PatternName)
End Function

Public Function CRegExps_Test( _
    CRegExps As Collection, _
    SourceString As String) _
    As String
    
    If CRegExps Is Nothing Then Exit Function
    If CRegExps.Count = 0 Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim ResultString As String
    
    Dim CRegExp_ As CRegExp
    For Each CRegExp_ In CRegExps
        Dim Result As String
        Result = CRegExp_Test(CRegExp_, SourceString)
        
        If Result <> "" Then
            ResultString = ResultString & Result & vbNewLine
        End If
    Next
    
    CRegExps_Test = ResultString
End Function

Public Function CRegExp_Test(CRegExp_ As CRegExp, SourceString As String) _
    As String
    
    If CRegExp_ Is Nothing Then Exit Function
    If CRegExp_.PatternName = "" Then Exit Function
    If CRegExp_.Pattern = "" Then Exit Function
    If SourceString = "" Then Exit Function
    
    CRegExp_Test = _
        CRegExp_.PatternName & vbTab & CStr(CRegExp_.Test(SourceString))
End Function

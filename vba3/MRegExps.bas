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

Public Function GetCRegExps(ParamsListString)
    If ParamsListString = "" Then Exit Function
    
    Dim CRegExps
    Set CRegExps = New Collection
    
    Dim ParamsList
    ParamsList = Split(ParamsListString, vbNewLine)
    
    Dim LB
    Dim UB
    LB = LBound(ParamsList)
    UB = UBound(ParamsList)
    
    Dim Index
    For Index = LB To UB
        Dim Params
        Params = ParamsList(Index)
        
        If Params <> "" Then
            Dim CRegExp_
            Set CRegExp_ = GetCRegExp(Params)
            
            If Not IsEmpty(CRegExp_) Then
                CRegExps.Add CRegExp_
            End If
        End If
    Next
    
    Set GetCRegExps = CRegExps
End Function

Public Function GetCRegExp(ParamsString)
    If ParamsString = "" Then Exit Function
    
    Dim Params
    Params = Split(ParamsString, vbTab)
    
    Dim LB
    Dim UB
    LB = LBound(Params)
    UB = UBound(Params)
    
    Dim PatternName
    Dim Pattern
    Dim IgnoreCase
    Dim GlobalMatch
    Dim MultiLine
    
    PatternName = CStr(Params(LB))
    If LB + 1 <= UB Then Pattern = CStr(Params(LB + 1))
    If LB + 2 <= UB Then IgnoreCase = CBool(Params(LB + 2))
    If LB + 3 <= UB Then GlobalMatch = CBool(Params(LB + 3))
    If LB + 4 <= UB Then MultiLine = CBool(Params(LB + 4))
    
    Dim CRegExp_
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
    ByRef CRegExps, _
    IgnoreCase, _
    GlobalMatch, _
    MultiLine)
    
    If CRegExps Is Nothing Then Exit Sub
    If CRegExps.Count = 0 Then Exit Sub
    
    Dim CRegExp_
    For Each CRegExp_ In CRegExps
        With CRegExp_
            .IgnoreCase = IgnoreCase
            .GlobalMatch = GlobalMatch
            .MultiLine = MultiLine
        End With
    Next
End Sub

Public Function CRegExps_Execute( _
    CRegExps, _
    SourceString) _
   
    
    If CRegExps Is Nothing Then Exit Function
    If CRegExps.Count = 0 Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim MatchesCollection
    Set MatchesCollection = New Collection
    
    Dim CRegExp_
    For Each CRegExp_ In CRegExps
        Dim RegExpMatches
        Set RegExpMatches = CRegExp_Execute(CRegExp_, SourceString)
        If Not RegExpMatches Is Nothing Then
            MatchesCollection.Add RegExpMatches
        End If
    Next
    
    Set CRegExps_Execute = MatchesCollection
End Function

Public Function CRegExp_Execute(CRegExp_, SourceString)
    If CRegExp_ Is Nothing Then Exit Function
    If CRegExp_.PatternName = "" Then Exit Function
    If CRegExp_.Pattern = "" Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim Matches
    Set Matches = CRegExp_.Execute(SourceString)
    If Matches Is Nothing Then Exit Function
    If Matches.Count = 0 Then Exit Function
    
    Dim RegExpMatches
    Set RegExpMatches = New CRegExpMatches
    With RegExpMatches
        .PatternName = CRegExp_.PatternName
        Set .Matches = Matches
    End With
    
    Set CRegExp_Execute = RegExpMatches
End Function

Public Function CRegExps_Replace( _
    CRegExps, _
    SourceString)
    
    CRegExps_Replace = SourceString
    
    If CRegExps Is Nothing Then Exit Function
    If CRegExps.Count = 0 Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim ResultString
    ResultString = SourceString
    
    Dim CRegExp_
    For Each CRegExp_ In CRegExps
        ResultString = CRegExp_Replace(CRegExp_, ResultString)
    Next
    
    CRegExps_Replace = ResultString
End Function

Public Function CRegExp_Replace(CRegExp_, SourceString)
    CRegExp_Replace = SourceString
    
    If CRegExp_ Is Nothing Then Exit Function
    If CRegExp_.PatternName = "" Then Exit Function
    If CRegExp_.Pattern = "" Then Exit Function
    If SourceString = "" Then Exit Function
    
    CRegExp_Replace = CRegExp_.Replace(SourceString, CRegExp_.PatternName)
End Function

Public Function CRegExps_Test( _
    CRegExps, _
    SourceString)
    
    If CRegExps Is Nothing Then Exit Function
    If CRegExps.Count = 0 Then Exit Function
    If SourceString = "" Then Exit Function
    
    Dim ResultString
    
    Dim CRegExp_
    For Each CRegExp_ In CRegExps
        Dim Result
        Result = CRegExp_Test(CRegExp_, SourceString)
        
        If Result <> "" Then
            ResultString = ResultString & Result & vbNewLine
        End If
    Next
    
    CRegExps_Test = ResultString
End Function

Public Function CRegExp_Test(CRegExp_, SourceString)
    If CRegExp_ Is Nothing Then Exit Function
    If CRegExp_.PatternName = "" Then Exit Function
    If CRegExp_.Pattern = "" Then Exit Function
    If SourceString = "" Then Exit Function
    
    CRegExp_Test = _
        CRegExp_.PatternName & vbTab & CStr(CRegExp_.Test(SourceString))
End Function

Attribute VB_Name = "SpdxLicenseTemplateText"
Option Explicit

'
' Copyright (c) 2020,2023 Koki Takeyama
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
' SPDX License List Matching Guidelines, v2.3
' https://spdx.github.io/spdx-spec/v2.3/license-matching-guidelines-and-templates/
'

Public Function GetMatchingLines(TemplateText As String) As String
    
    ' TemplateTextArray
    
    Dim TemplateTextArray As Variant
    TemplateTextArray = Split(Replace(TemplateText, vbCrLf, vbLf), vbLf)
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(TemplateTextArray)
    UB = UBound(TemplateTextArray)
    
    ' ResultArray
    
    Dim ResultArray() As String
    ReDim ResultArray(LB To UB)
    
    Dim Index As Long
    For Index = LB To UB
        ResultArray(Index) = GetMatchingText(Trim(TemplateTextArray(Index)))
    Next
    
    ' MatchingLines
    
    Dim MatchingLines As String
    MatchingLines = Join(ResultArray, vbNewLine) & vbNewLine
    
    GetMatchingLines = MatchingLines
End Function

Public Function GetMatchingText(TemplateText As String) As String
    If TemplateText = "" Then Exit Function
    
    ' B.15.3 Legacy Text Template format
    Const Pattern As String = "(?:" & _
        "<<var;name=""([^""]+)"";original=""(.+)"";match=""(.+)"">>" & "|" & _
        "<<beginOptional>>(.+)<<endOptional>>" & ")"
    
    If RegExp_Test(TemplateText, Pattern, True) Then
        Dim Matches As VBScript_RegExp_55.MatchCollection
        Set Matches = RegExp_Execute(TemplateText, Pattern, True, False)
        
        Dim Match As VBScript_RegExp_55.Match
        Set Match = Matches.Item(0)
        
        Dim PreviousPattern As String
        If Match.FirstIndex > 0 Then
            Dim PreviousText As String
            PreviousText = Left(TemplateText, Match.FirstIndex)
            PreviousPattern = GetMatchingPattern(PreviousText)
        End If
        
        Dim MiddlePattern As String
        If Match.SubMatches.Item(0) <> "" Then
            ' B.3.4 Guideline: replaceable text
            ' B.8 Bullets and numbering
            ' B.11 Copyright notice
            ' <<var;name="([^"]+)";original="(.+)";match="(.+)">>
            
            'Dim VarName As String
            'Dim VarOriginal As String
            Dim VarMatch As String
            'VarName = Match.SubMatches.Item(0)
            'VarOriginal = Match.SubMatches.Item(1)
            VarMatch = Match.SubMatches.Item(2)
            
            MiddlePattern = VarMatch
            
        Else
            ' B.3.5 Guideline: omittable text
            ' B.12 License name or title
            ' B.13 Extraneous text at the end of a license
            ' "<<beginOptional>>(.+)<<endOptional>>"
            
            Dim OptText As String
            OptText = Match.SubMatches.Item(3)
            
            Dim OptPattern As String
            OptPattern = GetMatchingPattern(OptText)
            
            MiddlePattern = "(?:" & OptPattern & ")?"
            
        End If
        
        Dim PostPattern As String
        If Match.FirstIndex + Match.Length < Len(TemplateText) Then
            Dim PostText As String
            PostText = _
                Right( _
                    TemplateText, _
                    Len(TemplateText) - (Match.FirstIndex + Match.Length))
            PostPattern = GetMatchingText(PostText)
        End If
        
        GetMatchingText = PreviousPattern & MiddlePattern & PostPattern
        
    Else
        GetMatchingText = GetMatchingPattern(TemplateText)
        
    End If
End Function

Public Function GetPlainText(TemplateText As String) As String
    If TemplateText = "" Then Exit Function
    
    Dim PlainText As String
    PlainText = TemplateText
    
    PlainText = Replace(PlainText, "<<beginOptional>>", "")
    PlainText = Replace(PlainText, "<<endOptional>>", "")
    PlainText = _
        RegExp_Replace( _
            PlainText, _
            "", "<<var;name=""([^""]+)"";original=""", _
            False, True, False)
    PlainText = _
        RegExp_Replace( _
            PlainText, _
            "", """;match=""([^""]+)"">>", _
            False, True, False)
    
    GetPlainText = PlainText
End Function

Public Function GetPlainTextEx(TemplateText As String) As String
    If TemplateText = "" Then Exit Function
    
    ' B.15.3 Legacy Text Template format
    'Const Pattern As String = "(?:" & _
    '    "<<var;name=""([^""]+)"";original=""(.+)"";match=""(.+)"">>" & "|" & _
    '    "<<beginOptional>>(.+)<<endOptional>>" & ")"
    Const Pattern As String = "(?:" & _
        "<<var;name=""([^""]+)"";original=""" & "|" & _
        "<<beginOptional>>" & ")"
    
    If Not RegExp_Test(TemplateText, Pattern, True) Then
        GetPlainTextEx = TemplateText
        Exit Function
    End If
    
    Dim Matches As VBScript_RegExp_55.MatchCollection
    Set Matches = RegExp_Execute(TemplateText, Pattern, True, False)
    
    Dim Match As VBScript_RegExp_55.Match
    Set Match = Matches.Item(0)
    
    Dim PreviousText As String
    If Match.FirstIndex > 0 Then
        PreviousText = Left(TemplateText, Match.FirstIndex)
    End If
    
    Dim MiddleText As String
    Dim MiddleTextTemp As String
    Dim PostText As String
    Dim PostTextTemp As String
    
    If Match.SubMatches.Item(0) <> "" Then
        ' B.3.4 Guideline: replaceable text
        ' B.8 Bullets and numbering
        ' B.11 Copyright notice
        ' <<var;name="([^"]+)";original="(.+)";match="(.+)">>
        
        'Dim VarName As String
        'Dim VarOriginal As String
        'Dim VarMatch As String
        'VarName = Match.SubMatches.Item(0)
        'VarOriginal = Match.SubMatches.Item(1)
        'VarMatch = Match.SubMatches.Item(2)
        
        'MiddleText = VarOriginal
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetPlainTextEx(PostTextTemp)
        'End If
        
        Dim VarMatchPos As Long
        Dim VarEndPos As Long
        
        VarMatchPos = _
            InStr( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText, _
                """;match=""")
        
        If VarMatchPos > 0 Then
            VarEndPos = _
                InStr( _
                    VarMatchPos + Len(""";match="""), _
                    TemplateText, _
                    """>>")
        End If
        
        If VarEndPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    VarMatchPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = MiddleTextTemp
            
            If VarEndPos - 1 + Len(""">>") < Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - (VarEndPos - 1 + Len(""">>")))
                PostText = GetPlainTextEx(PostTextTemp)
            End If
            
        End If
        
    Else
        ' B.3.5 Guideline: omittable text
        ' B.12 License name or title
        ' B.13 Extraneous text at the end of a license
        ' "<<beginOptional>>(.+)<<endOptional>>"
        
        'MiddleTextTemp = Match.SubMatches.Item(3)
        'MiddleText = GetPlainTextEx(MiddleTextTemp)
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetPlainTextEx(PostTextTemp)
        'End If
        
        Dim EndOptionalPos As Long
        'EndOptionalPos = _
        '    InStr( _
        '        Match.FirstIndex + Match.Length + 1, _
        '        TemplateText, _
        '        "<<endOptional>>")
        EndOptionalPos = _
            GetEndOptionalPos( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText)
        If EndOptionalPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    EndOptionalPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = GetPlainTextEx(MiddleTextTemp)
            
            If EndOptionalPos - 1 + Len("<<endOptional>>") < _
                Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - _
                            (EndOptionalPos - 1 + Len("<<endOptional>>")))
                PostText = GetPlainTextEx(PostTextTemp)
            End If
            
        End If
        
    End If
    
    GetPlainTextEx = PreviousText & MiddleText & PostText
End Function

Private Function GetEndOptionalPos(StartPos As Long, TemplateText As String) _
    As Long
    
    Dim BeginOptionalPos As Long
    BeginOptionalPos = InStr(StartPos, TemplateText, "<<beginOptional>>")
    
    Dim EndOptionalPos As Long
    EndOptionalPos = InStr(StartPos, TemplateText, "<<endOptional>>")
    
    If BeginOptionalPos = 0 Then
        GetEndOptionalPos = EndOptionalPos
        Exit Function
    End If
    
    If BeginOptionalPos > EndOptionalPos Then
        GetEndOptionalPos = EndOptionalPos
        Exit Function
    End If
    
    Dim EndOptionalPosTemp As Long
    EndOptionalPosTemp = _
        GetEndOptionalPos( _
            BeginOptionalPos + Len("<<beginOptional>>"), _
            TemplateText)
    
    EndOptionalPos = _
        GetEndOptionalPos( _
            EndOptionalPosTemp + Len("<<endOptional>>"), _
            TemplateText)
    
    GetEndOptionalPos = EndOptionalPos
End Function

Public Function GetFontText(TemplateText As String) As String
    If TemplateText = "" Then Exit Function
    
    ' B.15.3 Legacy Text Template format
    'Const Pattern As String = "(?:" & _
    '    "<<var;name=""([^""]+)"";original=""(.+)"";match=""(.+)"">>" & "|" & _
    '    "<<beginOptional>>(.+)<<endOptional>>" & ")"
    Const Pattern As String = "(?:" & _
        "<<var;name=""([^""]+)"";original=""" & "|" & _
        "<<beginOptional>>" & ")"
    
    If Not RegExp_Test(TemplateText, Pattern, True) Then
        GetFontText = Space(Len(TemplateText))
        Exit Function
    End If
    
    Dim Matches As VBScript_RegExp_55.MatchCollection
    Set Matches = RegExp_Execute(TemplateText, Pattern, True, False)
    
    Dim Match As VBScript_RegExp_55.Match
    Set Match = Matches.Item(0)
    
    Dim PreviousText As String
    If Match.FirstIndex > 0 Then
        PreviousText = Space(Match.FirstIndex)
    End If
    
    Dim MiddleText As String
    Dim MiddleTextTemp As String
    Dim PostText As String
    Dim PostTextTemp As String
    
    If Match.SubMatches.Item(0) <> "" Then
        ' B.3.4 Guideline: replaceable text
        ' B.8 Bullets and numbering
        ' B.11 Copyright notice
        ' <<var;name="([^"]+)";original="(.+)";match="(.+)">>
        
        'Dim VarName As String
        'Dim VarOriginal As String
        'Dim VarMatch As String
        'VarName = Match.SubMatches.Item(0)
        'VarOriginal = Match.SubMatches.Item(1)
        'VarMatch = Match.SubMatches.Item(2)
        
        'MiddleText = StrRepeat("r", Len(VarOriginal))
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetFontText(PostTextTemp)
        'End If
        
        Dim VarMatchPos As Long
        Dim VarEndPos As Long
        
        VarMatchPos = _
            InStr( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText, _
                """;match=""")
        
        If VarMatchPos > 0 Then
            VarEndPos = _
                InStr( _
                    VarMatchPos + Len(""";match="""), _
                    TemplateText, _
                    """>>")
        End If
        
        If VarEndPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    VarMatchPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = StrRepeat("r", Len(MiddleTextTemp))
            
            If VarEndPos - 1 + Len(""">>") < Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - (VarEndPos - 1 + Len(""">>")))
                PostText = GetFontText(PostTextTemp)
            End If
            
        End If
        
    Else
        ' B.3.5 Guideline: omittable text
        ' B.12 License name or title
        ' B.13 Extraneous text at the end of a license
        ' "<<beginOptional>>(.+)<<endOptional>>"
        
        'MiddleTextTemp = Match.SubMatches.Item(3)
        'MiddleText = StrRepeat("b", Len(GetFontText(MiddleTextTemp)))
        
        'If Match.FirstIndex + Match.Length < Len(TemplateText) Then
        '    PostTextTemp = _
        '        Right( _
        '            TemplateText, _
        '            Len(TemplateText) - (Match.FirstIndex + Match.Length))
        '    PostText = GetFontText(PostTextTemp)
        'End If
        
        Dim EndOptionalPos As Long
        'EndOptionalPos = _
        '    InStr( _
        '        Match.FirstIndex + Match.Length + 1, _
        '        TemplateText, _
        '        "<<endOptional>>")
        EndOptionalPos = _
            GetEndOptionalPos( _
                Match.FirstIndex + Match.Length + 1, _
                TemplateText)
        If EndOptionalPos > 0 Then
            MiddleTextTemp = _
                Mid( _
                    TemplateText, _
                    Match.FirstIndex + Match.Length + 1, _
                    EndOptionalPos - 1 - (Match.FirstIndex + Match.Length))
            MiddleText = StrRepeat("b", Len(GetFontText(MiddleTextTemp)))
            
            If EndOptionalPos - 1 + Len("<<endOptional>>") < _
                Len(TemplateText) Then
                PostTextTemp = _
                    Right( _
                        TemplateText, _
                        Len(TemplateText) - _
                            (EndOptionalPos - 1 + Len("<<endOptional>>")))
                PostText = GetFontText(PostTextTemp)
            End If
            
        End If
        
    End If
    
    GetFontText = PreviousText & MiddleText & PostText
End Function

Private Function StrRepeat(Text As String, Length As Long) As String
    If Text = "" Then Exit Function
    If Length <= 0 Then Exit Function
    
    Dim Result As String
    
    Dim Index As Long
    For Index = 0 To Length - 1
        Result = Result & Text
    Next
    
    StrRepeat = Result
End Function

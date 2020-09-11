Attribute VB_Name = "SpdxLicenseText"
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
' SPDX License List Matching Guidelines, v2.1
' https://spdx.org/spdx-license-list/matching-guidelines
'

Public Function GetMatchingPattern(LicenseText As String) As String
    Dim TempString As String
    TempString = LCase(LicenseText)
    
    ' Escape Special Characters
    TempString = RegExpReplace(TempString, "\(", "\(")
    TempString = RegExpReplace(TempString, "\)", "\)")
    TempString = RegExpReplace(TempString, "\[", "\[")
    TempString = RegExpReplace(TempString, "\]", "\]")
    
    ' 8. Varietal Word Spelling
    TempString = RegExpReplaceWords(TempString)
    
    ' 3. Whitespace
    ' 6. Code Comment Indicators
    TempString = RegExpReplace(TempString, "\W+", "\s+")
    
    ' 5.1.1 Guideline: Punctuation
    TempString = RegExpReplace(TempString, "\.", "\.")
    
    ' 5.1.2 Guideline: Hyphens, Dashes
    ' https://en.wikipedia.org/wiki/Dash
    ' https://en.wikipedia.org/wiki/Hyphen
    TempString = RegExpReplace(TempString, "\W+", "-")
    
    ' 5.1.3 Guideline: Quotes
    ' https://en.wikipedia.org/wiki/Quotation_mark
    TempString = RegExpReplace(TempString, "\W+", "['""]")
    
    ' 13. HTTP Protocol
    TempString = RegExpReplace(TempString, "https?://", "https?://")
    
    GetMatchingPattern = TempString
End Function

'
' 8. Varietal Word Spelling
'
' | Word1 | Word2 | MatchingPattern |
' | --- | --- | --- |
' | acknowledgement | acknowledgment | acknowledge?ment |
' | analog | analogue | analog(?:ue)? |
' | analyze | analyse | analy[zs]e |
' | artifact | artefact | art[ie]fact |
' | authorization | authorisation | authori[zs]ation |
' | authorized | authorised | authori[zs]ed |
' | caliber | calibre | calib(?:er|re) |
' | canceled | cancelled | cancell?ed |
' | capitalizations | capitalisations | capitali[zs]ations |
' | catalog | catalogue | catalog(?:ue)? |
' | categorize | categorise | categori[zs]e |
' | center | centre | cent(?:er|re) |
' | copyright holder | copyright owner | copyright\W+(?:hold|own)er |
' | emphasized | emphasised | emphasi[zs]ed |
' | favor | favour | favou?r |
' | favorite | favourite | favou?rite |
' | fulfill | fulfil | fulfill? |
' | fulfillment | fulfilment | fulfill?ment |
' | Initialize | initialise | initiali[zs]e |
' | judgement | judgment | judge?ment |
' | labeling | labelling | labell?ing |
' | labor | labour | labou?r |
' | license | licence | licen[sc]e |
' | maximize | maximise | maximi[zs]e |
' | modeled | modelled | modell?ed |
' | modeling | modelling | modell?ing |
' | noncommercial | non-commercial | non-?commercial |
' | offense | offence | offen[sc]e |
' | optimize | optimise | optimi[zs]e |
' | organization | organisation | organi[zs]ation |
' | organize | organise | organi[zs]e |
' | percent | per cent | per\s*cent |
' | practice | practise | practi[cs]e |
' | program | programme | program(?:me)? |
' | realize | realise | reali[zs]e |
' | Recognize | recognise | recogni[zs]e |
' | signaling | signalling | signall?ing |
' | sublicense | sub-license | sub(?: |-)?licen[sc]e |
' | sub-license | sub license | sub(?: |-)?licen[sc]e |
' | sublicense | sub license | sub(?: |-)?licen[sc]e |
' | utilization | utilisation | utili[zs]ation |
' | while | whilst | whil(?:e|st) |
' | wilfull | wilful | wilfull? |
'

Private Function RegExpReplaceWords(SourceString As String) As String
    Dim ResultString As String
    ResultString = SourceString
    
    Dim PatternAndReplaceStringArray As Variant
    PatternAndReplaceStringArray = Array( _
        "sub\W*licen[sc]e", _
        "acknowledge?ment", "analog(?:ue)?", "analy[zs]e", _
        "art[ie]fact", "authori[zs]ation", "authori[zs]ed", _
        "calib(?:er|re)", "cancell?ed", "capitali[zs]ations", _
        "catalog(?:ue)?", "categori[zs]e", "cent(?:er|re)", _
        "copyright\W+(?:hold|own)er", "emphasi[zs]ed", _
        "favou?r", "favou?rite", "fulfill?", _
        "fulfill?ment", "initiali[zs]e", "judge?ment", _
        "labell?ing", "labou?r", "licen[sc]e", _
        "maximi[zs]e", "modell?ed", "modell?ing", _
        "non\W*commercial", "offen[sc]e", "optimi[zs]e", _
        "organi[zs]ation", "organi[zs]e", _
        "per\s*cent", "practi[cs]e", "program(?:me)?", _
        "reali[zs]e", "recogni[zs]e", "signall?ing", _
        "utili[zs]ation", "whil(?:e|st)", "wilfull?")
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(PatternAndReplaceStringArray)
    UB = UBound(PatternAndReplaceStringArray)
    
    Dim Index As Long
    For Index = LB To UB
        Dim PatternAndReplaceString As String
        PatternAndReplaceString = CStr(PatternAndReplaceStringArray(Index))
        
        ResultString = _
            RegExpReplace( _
                ResultString, _
                PatternAndReplaceString, _
                PatternAndReplaceString)
    Next
    
    RegExpReplaceWords = ResultString
End Function

Private Function RegExpReplace( _
    SourceString As String, _
    ReplaceString As String, _
    Pattern As String) As String
    
    On Error Resume Next
    
    With GetRegExp()
        .Pattern = Pattern
        .IgnoreCase = True ' 4. Capitalization
        .Global = True
        .MultiLine = False
        RegExpReplace = .Replace(SourceString, ReplaceString)
    End With
End Function

'
' Microsoft VBScript Regular Expression 5.5
' - VBScript_RegExp_55.RegExp
'

Private Function GetRegExp() As VBScript_RegExp_55.RegExp
    Static RegExpObject As VBScript_RegExp_55.RegExp
    If RegExpObject Is Nothing Then
        Set RegExpObject = New VBScript_RegExp_55.RegExp
    End If
    Set GetRegExp = RegExpObject
End Function

'
' --- Test ---
'

Private Sub Test_GetMatchingPattern()
    Dim LicenseText As String
    LicenseText = InputBox("LicenseText")
    If LicenseText = "" Then Exit Sub
    
    Dim MatchingPattern As String
    MatchingPattern = GetMatchingPattern(LicenseText)
    
    Debug_Print "--- LicenseText ---"
    Debug_Print LicenseText
    Debug_Print "--- MatchingPattern ---"
    Debug_Print MatchingPattern
End Sub

Private Sub Test_GetMatchingPattern_Apache20()
    Dim StandardLicenseHeader As String
    StandardLicenseHeader = _
"Licensed under the Apache License, Version 2.0 (the ""License""); " & _
"you may not use this file except in compliance with the License. " & _
"You may obtain a copy of the License at " & _
"http://www.apache.org/licenses/LICENSE-2.0 " & _
"Unless required by applicable law or agreed to in writing, software " & _
"distributed under the License is distributed on an ""AS IS"" BASIS, " & _
"WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. " & _
"See the License for the specific language governing permissions and " & _
"limitations under the License."
    
    Dim MatchingPattern As String
    MatchingPattern = GetMatchingPattern(StandardLicenseHeader)
    
    Debug_Print "--- StandardLicenseHeader ---"
    Debug_Print StandardLicenseHeader
    Debug_Print "--- MatchingPattern ---"
    Debug_Print MatchingPattern
End Sub

Private Sub Test_GetMatchingPattern_MIT()
    Dim LicenseText As String
    LicenseText = _
"Permission is hereby granted, free of charge, to any person obtaining a copy " & _
"of this software and associated documentation files (the ""Software""), to deal " & _
"in the Software without restriction, including without limitation the rights " & _
"to use, copy, modify, merge, publish, distribute, sublicense, and/or sell " & _
"copies of the Software, and to permit persons to whom the Software is furnished " & _
"to do so, subject to the following conditions: " & _
"The above copyright notice and this permission notice (including the next " & _
"paragraph) shall be included in all copies or substantial portions of the " & _
"Software. " & _
"THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR " & _
"IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS " & _
"FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS " & _
"OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, " & _
"WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF " & _
"OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."
    
    Dim MatchingPattern As String
    MatchingPattern = GetMatchingPattern(LicenseText)
    
    Debug_Print "--- LicenseText ---"
    Debug_Print LicenseText
    Debug_Print "--- MatchingPattern ---"
    Debug_Print MatchingPattern
End Sub

Private Sub Test_GetMatchingPattern_Chars()
    Test_GetMatchingPattern_Word "("
    Test_GetMatchingPattern_Word ")"
    Test_GetMatchingPattern_Word "["
    Test_GetMatchingPattern_Word "]"
    Test_GetMatchingPattern_Word " "
    Test_GetMatchingPattern_Word "."
    Test_GetMatchingPattern_Word "-"
    Test_GetMatchingPattern_Word "'"
    Test_GetMatchingPattern_Word """"
    Test_GetMatchingPattern_Word "http://"
    Test_GetMatchingPattern_Word "https://"
End Sub

Private Sub Test_GetMatchingPattern_Words()
    Test_GetMatchingPattern_Word "acknowledgement"
    Test_GetMatchingPattern_Word "acknowledgment"
    Test_GetMatchingPattern_Word "analog"
    Test_GetMatchingPattern_Word "analogue"
    Test_GetMatchingPattern_Word "analyze"
    Test_GetMatchingPattern_Word "analyse"
    Test_GetMatchingPattern_Word "artifact"
    Test_GetMatchingPattern_Word "artefact"
    Test_GetMatchingPattern_Word "authorization"
    Test_GetMatchingPattern_Word "authorisation"
    Test_GetMatchingPattern_Word "authorized"
    Test_GetMatchingPattern_Word "authorised"
    Test_GetMatchingPattern_Word "caliber"
    Test_GetMatchingPattern_Word "calibre"
    Test_GetMatchingPattern_Word "canceled"
    Test_GetMatchingPattern_Word "cancelled"
    Test_GetMatchingPattern_Word "capitalizations"
    Test_GetMatchingPattern_Word "capitalisations"
    Test_GetMatchingPattern_Word "catalog"
    Test_GetMatchingPattern_Word "catalogue"
    Test_GetMatchingPattern_Word "categorize"
    Test_GetMatchingPattern_Word "categorise"
    Test_GetMatchingPattern_Word "center"
    Test_GetMatchingPattern_Word "centre"
    Test_GetMatchingPattern_Word "copyright holder"
    Test_GetMatchingPattern_Word "copyright owner"
    Test_GetMatchingPattern_Word "emphasized"
    Test_GetMatchingPattern_Word "emphasised"
    Test_GetMatchingPattern_Word "favor"
    Test_GetMatchingPattern_Word "favour"
    Test_GetMatchingPattern_Word "favorite"
    Test_GetMatchingPattern_Word "favourite"
    Test_GetMatchingPattern_Word "fulfill"
    Test_GetMatchingPattern_Word "fulfil"
    Test_GetMatchingPattern_Word "fulfillment"
    Test_GetMatchingPattern_Word "fulfilment"
    Test_GetMatchingPattern_Word "Initialize"
    Test_GetMatchingPattern_Word "initialise"
    Test_GetMatchingPattern_Word "judgement"
    Test_GetMatchingPattern_Word "judgment"
    Test_GetMatchingPattern_Word "labeling"
    Test_GetMatchingPattern_Word "labelling"
    Test_GetMatchingPattern_Word "labor"
    Test_GetMatchingPattern_Word "labour"
    Test_GetMatchingPattern_Word "license"
    Test_GetMatchingPattern_Word "licence"
    Test_GetMatchingPattern_Word "maximize"
    Test_GetMatchingPattern_Word "maximise"
    Test_GetMatchingPattern_Word "modeled"
    Test_GetMatchingPattern_Word "modelled"
    Test_GetMatchingPattern_Word "modeling"
    Test_GetMatchingPattern_Word "modelling"
    Test_GetMatchingPattern_Word "noncommercial"
    Test_GetMatchingPattern_Word "non-commercial"
    Test_GetMatchingPattern_Word "offense"
    Test_GetMatchingPattern_Word "offence"
    Test_GetMatchingPattern_Word "optimize"
    Test_GetMatchingPattern_Word "optimise"
    Test_GetMatchingPattern_Word "organization"
    Test_GetMatchingPattern_Word "organisation"
    Test_GetMatchingPattern_Word "organize"
    Test_GetMatchingPattern_Word "organise"
    Test_GetMatchingPattern_Word "percent"
    Test_GetMatchingPattern_Word "per cent"
    Test_GetMatchingPattern_Word "practice"
    Test_GetMatchingPattern_Word "practise"
    Test_GetMatchingPattern_Word "program"
    Test_GetMatchingPattern_Word "programme"
    Test_GetMatchingPattern_Word "realize"
    Test_GetMatchingPattern_Word "realise"
    Test_GetMatchingPattern_Word "recognize"
    Test_GetMatchingPattern_Word "recognise"
    Test_GetMatchingPattern_Word "signaling"
    Test_GetMatchingPattern_Word "signalling"
    Test_GetMatchingPattern_Word "sublicense"
    Test_GetMatchingPattern_Word "sub-license"
    Test_GetMatchingPattern_Word "sub license"
    Test_GetMatchingPattern_Word "utilization"
    Test_GetMatchingPattern_Word "utilisation"
    Test_GetMatchingPattern_Word "while"
    Test_GetMatchingPattern_Word "whilst"
    Test_GetMatchingPattern_Word "wilfull"
    Test_GetMatchingPattern_Word "wilful"
End Sub

Private Sub Test_GetMatchingPattern_Word(Word As String)
    Dim MatchingPattern As String
    MatchingPattern = GetMatchingPattern(Word)
    Debug_Print Word & " : " & MatchingPattern
End Sub

Private Sub Debug_Print(Str As String)
    Debug.Print Str
End Sub

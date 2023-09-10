Attribute VB_Name = "Test_SpdxLicenseText"
Option Explicit

'
' Copyright (c) 2020,2022,2023 Koki Takeyama
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

Public Sub Test_GetMatchingPattern_Apache20()
    Test_GetMatchingPattern_Core _
"Licensed under the Apache License, Version 2.0 (the ""License""); " & _
"you may not use this file except in compliance with the License. " & _
"You may obtain a copy of the License at " & _
"http://www.apache.org/licenses/LICENSE-2.0 " & _
"Unless required by applicable law or agreed to in writing, software " & _
"distributed under the License is distributed on an ""AS IS"" BASIS, " & _
"WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. " & _
"See the License for the specific language governing permissions and " & _
"limitations under the License."
End Sub

Public Sub Test_GetMatchingPattern_MIT()
    Test_GetMatchingPattern_Core _
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
End Sub

Public Sub Test_GetMatchingPattern_Chars()
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

Public Sub Test_GetMatchingPattern_Words()
    Test_GetMatchingPattern_Word "acknowledgement"
    Test_GetMatchingPattern_Word "acknowledgment"
    Test_GetMatchingPattern_Word "analog"
    Test_GetMatchingPattern_Word "analogue"
    Test_GetMatchingPattern_Word "and"
    Test_GetMatchingPattern_Word "&"
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
    Test_GetMatchingPattern_Word "merchantability"
    Test_GetMatchingPattern_Word "merchantibility"
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

'
' --- Test Core ---
'

Public Sub Test_GetMatchingPattern_Core(Text)
    Dim MatchingPattern
    MatchingPattern = GetMatchingPattern(Text)
    
    Dim SimpleMatchingPattern
    SimpleMatchingPattern = GetSimpleMatchingPattern(Text)
    
    Debug_Print "--- Text ---"
    Debug_Print Text
    Debug_Print "--- MatchingPattern ---"
    Debug_Print MatchingPattern
    Debug_Print "--- SimpleMatchingPattern ---"
    Debug_Print SimpleMatchingPattern
End Sub

Public Sub Test_GetMatchingPattern_Word(Word)
    Dim MatchingPattern
    MatchingPattern = GetMatchingPattern(Word)
    
    Dim SimpleMatchingPattern
    SimpleMatchingPattern = GetSimpleMatchingPattern(Word)
    
    Debug_Print Word & " : " & MatchingPattern & " : " & SimpleMatchingPattern
End Sub

Attribute VB_Name = "Test_SpdxLicenseTemplateText"
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

Public Sub Test_GetMatchingLines_Apache20()
    Test_GetMatchingLines_Core _
"Copyright <<var;name=""copyright"";original=""[yyyy] [name of copyright owner]"";match="".+"">>" & vbLf & vbLf & _
"Licensed under the Apache License, Version 2.0 (the ""License""); " & vbLf & vbLf & _
"you may not use this file except in compliance with the License. " & vbLf & vbLf & _
"You may obtain a copy of the License at " & vbLf & vbLf & _
"http://www.apache.org/licenses/LICENSE-2.0 " & vbLf & vbLf & _
"Unless required by applicable law or agreed to in writing, software " & vbLf & vbLf & _
"distributed under the License is distributed on an ""AS IS"" BASIS, " & vbLf & vbLf & _
"WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. " & vbLf & vbLf & _
"See the License for the specific language governing permissions and " & vbLf & vbLf & _
"limitations under the License."
End Sub

Public Sub Test_GetMatchingLines_MIT()
    Test_GetMatchingLines_Core _
"<<beginOptional>> MIT License<<endOptional>> " & _
"<<var;name=""copyright"";original=""Copyright (c) <year> <copyright holders>"";match="".{0,1000}"">>" & vbLf & vbLf & _
"Permission is hereby granted, free of charge, to any person obtaining a copy " & _
"of <<var;name=""software"";original=""this software and associated documentation files"";match=""this software and associated documentation files|this source file"">> (the ""Software""), to deal " & _
"in the Software without restriction, including without limitation the rights " & _
"to use, copy, modify, merge, publish, distribute, sublicense, and/or sell " & _
"copies of the Software, and to permit persons to whom the Software is furnished " & _
"to do so, subject to the following conditions: " & vbLf & vbLf & _
"The above copyright notice and this permission notice<<beginOptional>> (including the next paragraph)<<endOptional>>" & _
" shall be included in all copies or substantial portions of the " & _
"Software. " & vbLf & vbLf & _
"THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR " & _
"IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS " & _
"FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL <<var;name=""copyrightHolder"";original=""THE AUTHORS OR COPYRIGHT HOLDERS"";match="".+"">>" & _
" BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, " & _
"WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF " & _
"OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."
End Sub

'
' --- Test Core ---
'

Public Sub Test_GetMatchingLines_Core(Text)
    Dim MatchingLines
    MatchingLines = GetMatchingLines(Text)
    
    Debug_Print "--- Text ---"
    Debug_Print Text
    Debug_Print "--- MatchingLines ---"
    Debug_Print MatchingLines
End Sub

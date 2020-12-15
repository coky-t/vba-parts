VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private m_PatternName
Private m_RegExp

Private Sub Class_Initialize()
    Set m_RegExp = CreateObject("VBScript.RegExp")
End Sub

Private Sub Class_Terminate()
    Set m_RegExp = Nothing
End Sub

Public Property Get PatternName()
    PatternName = m_PatternName
End Property

Public Property Let PatternName(PatternName_)
    m_PatternName = PatternName_
End Property

'
' Pattern:
'   Required. Regular string expression being searched for.
'   https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/f97kw5ka(v=vs.84)
'

Public Property Get Pattern()
    If m_RegExp Is Nothing Then Exit Property
    Pattern = m_RegExp.Pattern
End Property

Public Property Let Pattern(Pattern_)
    If m_RegExp Is Nothing Then Exit Property
    m_RegExp.Pattern = Pattern_
End Property

'
' IgnoreCase:
'   Optional. The value is False if the search is case-sensitive,
'   True if it is not. Default is False.
'

Public Property Get IgnoreCase()
    If m_RegExp Is Nothing Then Exit Property
    IgnoreCase = m_RegExp.IgnoreCase
End Property

Public Property Let IgnoreCase(IgnoreCase_)
    If m_RegExp Is Nothing Then Exit Property
    m_RegExp.IgnoreCase = IgnoreCase_
End Property

'
' GlobalMatch:
'   Optional. The value is True if the search applies to the entire string,
'   False if it does not. Default is False.
'

Public Property Get GlobalMatch()
    If m_RegExp Is Nothing Then Exit Property
    GlobalMatch = m_RegExp.Global
End Property

Public Property Let GlobalMatch(GlobalMatch_)
    If m_RegExp Is Nothing Then Exit Property
    m_RegExp.Global = GlobalMatch_
End Property

'
' MultiLine:
'   Optional. The value is False if the search is single-line mode,
'   True if it is multi-line mode. Default is False.
'

Public Property Get MultiLine()
    If m_RegExp Is Nothing Then Exit Property
    MultiLine = m_RegExp.MultiLine
End Property

Public Property Let MultiLine(MultiLine_)
    If m_RegExp Is Nothing Then Exit Property
    m_RegExp.MultiLine = MultiLine_
End Property

'
' Execute
' - Executes a regular expression search against a specified string.
'
' Replace
' - Replaces text found in a regular expression search.
'
' Test
' - Executes a regular expression search against a specified string
'   and returns a Boolean value that indicates if a pattern match was found.
'

'
' SourceString:
'   Required. The text string upon which the regular expression is executed.
'
' ReplaceString:
'   Required. The replacement text string.
'

Public Function Execute(SourceString)
    If m_RegExp Is Nothing Then Exit Function
    Set Execute = m_RegExp.Execute(SourceString)
End Function

Public Function Replace(SourceString, ReplaceString)
    Replace = SourceString
    If m_RegExp Is Nothing Then Exit Function
    Replace = m_RegExp.Replace(SourceString, ReplaceString)
End Function

Public Function Test(SourceString)
    If m_RegExp Is Nothing Then Exit Function
    Test = m_RegExp.Test(SourceString)
End Function

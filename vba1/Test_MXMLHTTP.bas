Attribute VB_Name = "Test_MXMLHTTP"
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

Public Sub Test_HttpGet()
    Test_HttpGet_Core "https://www.google.com"
End Sub

Public Sub Test_HttpPost()
    Test_HttpPost_Core "https://www.yahoo.com", ""
End Sub

'
' --- Test Core ---
'

Public Sub Test_HttpGet_Core(Url As String)
    Dim Text As String
    Text = HttpGet(Url)
    Debug_Print Text
End Sub

Public Sub Test_HttpPost_Core(Url As String, Body As String)
    Dim Text As String
    Text = HttpPost(Url, Body)
    Debug_Print Text
End Sub

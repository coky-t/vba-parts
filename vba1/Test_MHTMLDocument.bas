Attribute VB_Name = "Test_MHTMLDocument"
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

Public Sub Test_ParseJsonText1()
    Test_ParseJsonText_Core "{a:1,b:2}"
End Sub

Public Sub Test_ParseJsonText2()
    Test_ParseJsonText_Core "{""key1"":""value1"",""key2"":""value2""}"
End Sub

Public Sub Test_ParseJsonText3()
    Test_ParseJsonText_Core "[10,11,12]"
End Sub

Public Sub Test_ParseJsonText4()
    Test_ParseJsonText_Core "[""a"",""b"",""c""]"
End Sub

Public Sub Test_ParseJsonText5()
    Test_ParseJsonText_Core _
    "{key1:1,""key2"":""value2"",key3:{key3_1:3},key4:[""a"",""b"",""c""]}"
End Sub

Public Sub Test_ParseJsonText6()
    Test_ParseJsonText_Core "[1,""value2"",{key3_1:3},[""a"",""b"",""c""]]"
End Sub

'
' --- Test Core ---
'

Public Sub Test_ParseJsonText_Core(JsonText As String)
    Debug_Print "==="
    Debug_Print JsonText
    Debug_Print "==="
    
    Dim JsonObject As Object
    Set JsonObject = ParseJsonText(JsonText)
    
    Debug_Print_JsonObject JsonObject
End Sub

Public Sub Debug_Print_JsonObject(JsonObject As Object)
    Dim Keys As Object
    Set Keys = GetJsonKeys(JsonObject)
    
    Dim KeysLength As Long
    KeysLength = GetJsonKeysLength(Keys)
    
    Dim Index As Long
    Dim Key As Variant
    Dim Value As Variant
    For Index = 0 To KeysLength - 1
        Key = GetJsonKeysItem(Keys, Index)
        If IsJsonItemObject(JsonObject, Key) Then
            Debug_Print Key & " ---"
            Debug_Print_JsonObject GetJsonItemObject(JsonObject, Key)
            Debug_Print Key & " ---"
        Else
            Value = CStr(GetJsonItemValue(JsonObject, Key))
            Debug_Print Key & ": " & Value
        End If
    Next
End Sub

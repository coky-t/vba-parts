Attribute VB_Name = "MHTMLDocument"
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
' Microsoft HTML Object Library
' - MSHTML.HTMLDocument
' - htmlfile
'

'
' --- MSHTML.HTMLDocument ---
'

'
' GetHTMLDocumentForJson
' - Returns a MSHTML.HTMLDocument object.
'

'
' HTMLDocument:
'   Optional. The name of a MSHTML.HTMLDocument object.
'

Public Function GetHTMLDocumentForJson( _
    Optional HTMLDocument As MSHTML.HTMLDocument) As MSHTML.HTMLDocument
    
    If HTMLDocument Is Nothing Then
        Dim HTMLDoc As Object
        Set HTMLDoc = New MSHTML.HTMLDocument
        With HTMLDoc
            .write _
                "<script>document.ParseJsonText=function (JsonText) { " & _
                "return eval('(' + JsonText + ')'); }</script>"
            .write _
                "<script>document.GetKeys=function (JsonObj) { " & _
                "var keys = []; " & _
                "for (var key in JsonObj) { keys.push(key); } " & _
                "return keys; }</script>"
        End With
        Set GetHTMLDocumentForJson = HTMLDoc
    Else
        Set GetHTMLDocumentForJson = HTMLDocument
    End If
End Function

'
' === Json ===
'

'
' ParseJsonText
' - Returns a JSON object.
'

'
' JsonText:
'   Required. String expression that identifies JSON data.
'
' HTMLDocument:
'   Optional. The name of a MSHTML.HTMLDocument object.
'

Public Function ParseJsonText( _
    JsonText As String, _
    Optional HTMLDocument As MSHTML.HTMLDocument) As Object
    
    On Error Resume Next
    
    Set ParseJsonText = _
        CallByName( _
            GetHTMLDocumentForJson(HTMLDocument), _
            "ParseJsonText", _
            VbMethod, _
            JsonText)
End Function

'
' GetJsonKeys
' - Returns an array containing all existing keys in a JSON object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'
' HTMLDocument:
'   Optional. The name of a MSHTML.HTMLDocument object.
'

Public Function GetJsonKeys( _
    JsonObject As Object, _
    Optional HTMLDocument As MSHTML.HTMLDocument) As Object
    
    On Error Resume Next
    
    Set GetJsonKeys = _
        CallByName( _
            GetHTMLDocumentForJson(HTMLDocument), _
            "GetKeys", _
            VbMethod, _
            JsonObject)
End Function

'
' GetJsonKeysLength
' - Returns the number of elements in a JSON keys array.
'

'
' JsonKeys:
'   Required. The name of a JSON keys array object
'

Public Function GetJsonKeysLength(JsonKeys As Object) As Long
    GetJsonKeysLength = CallByName(JsonKeys, "length", VbGet)
End Function

'
' GetJsonKeysItem
' - Returns an item of JSON  object.
'

'
' JsonKeys:
'   Required. The name of a JSON keys array object
'
' Index:
'   Required. Index associated with the item being retrieved.
'

Public Function GetJsonKeysItem(JsonKeys As Object, Index As Long) As Variant
    GetJsonKeysItem = CallByName(JsonKeys, Index, VbGet)
End Function

'
' IsJsonItemObject
' - Returns a Boolean value indicating whether an item of JSON object
'   represents an object variable.
'

'
' JsonObject:
'   Required. The name of a JSON object
'
' Key:
'   Required. Key associated with the item being retrieved.
'

Public Function IsJsonItemObject( _
    JsonObject As Object, _
    Key As Variant) As Boolean
    
    IsJsonItemObject = IsObject(CallByName(JsonObject, Key, VbGet))
End Function

'
' GetJsonItemValue
' - Returns an item of JSON object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'
' Key:
'   Required. Key associated with the item being retrieved.
'

Public Function GetJsonItemValue( _
    JsonObject As Object, _
    Key As Variant) As Variant
    
    GetJsonItemValue = CallByName(JsonObject, Key, VbGet)
End Function

'
' GetJsonItemObject
' - Returns an item of JSON object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'
' Key:
'   Required. Key associated with the item being retrieved.
'

Public Function GetJsonItemObject( _
    JsonObject As Object, _
    Key As Variant) As Object
    
    Set GetJsonItemObject = CallByName(JsonObject, Key, VbGet)
End Function

'
' --- Test ---
'

Private Sub Test_ParseJsonText()
    Dim JsonText As String
    JsonText = InputBox("JsonText")
    If JsonText = "" Then Exit Sub
    
    Debug_Print_ParseJsonText JsonText
End Sub

Private Sub Test_ParseJsonText1()
    Debug_Print_ParseJsonText "{a:1,b:2}"
End Sub

Private Sub Test_ParseJsonText2()
    Debug_Print_ParseJsonText "{""key1"":""value1"",""key2"":""value2""}"
End Sub

Private Sub Test_ParseJsonText3()
    Debug_Print_ParseJsonText "[10,11,12]"
End Sub

Private Sub Test_ParseJsonText4()
    Debug_Print_ParseJsonText "[""a"",""b"",""c""]"
End Sub

Private Sub Test_ParseJsonText5()
    Debug_Print_ParseJsonText "{key1:1,""key2"":""value2"",key3:{key3_1:3},key4:[""a"",""b"",""c""]}"
End Sub

Private Sub Test_ParseJsonText6()
    Debug_Print_ParseJsonText "[1,""value2"",{key3_1:3},[""a"",""b"",""c""]]"
End Sub

Private Sub Debug_Print_ParseJsonText(JsonText As String)
    Debug_Print "==="
    Debug_Print JsonText
    Debug_Print "==="
    
    Dim JsonObject As Object
    Set JsonObject = ParseJsonText(JsonText)
    
    Debug_Print_JsonObject JsonObject
End Sub

Private Sub Debug_Print_JsonObject(JsonObject As Object)
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

Private Sub Debug_Print(Str As String)
    Debug.Print Str
End Sub

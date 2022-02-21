Attribute VB_Name = "MHTMLDocument"
Option Explicit

'
' Copyright (c) 2020,2022 Koki Takeyama
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

#Const UseCallByName = True

'
' --- MSHTML.HTMLDocument ---
'

'
' GetHTMLDocumentForJson
' - Returns a MSHTML.HTMLDocument object.
'

Public Function GetHTMLDocumentForJson() As MSHTML.HTMLDocument
    Static HTMLDocument As Object 'MSHTML.HTMLDocument
    If HTMLDocument Is Nothing Then
        Set HTMLDocument = New MSHTML.HTMLDocument
        With HTMLDocument
            .write _
                "<script>document.ParseJsonText=function (JsonText) { " & _
                "return eval('(' + JsonText + ')'); }</script>"
            .write _
                "<script>document.GetKeys=function (JsonObj) { " & _
                "var keys = []; " & _
                "for (var key in JsonObj) { keys.push(key); } " & _
                "return keys; }</script>"
#If UseCallByName Then
            ' nop
#Else
            .write _
                "<script>document.GetItem=function (obj, i) { " & _
                "return obj[i]; }</script>"
            .write _
                "<script>document.GetLength=function (obj) { " & _
                "return obj.length; }</script>"
#End If
            .write _
                "<script>document.IsJsonArray=function (obj) { " & _
                "return Object.prototype.toString.call(obj) === " & _
                "'[object Array]'; }</script>"
            .write _
                "<script>document.GetJsonText=function (obj) { " & _
                "return document.parentWindow.JSON.stringify(obj); }</script>"
        End With
    End If
    Set GetHTMLDocumentForJson = HTMLDocument
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

Public Function ParseJsonText(JsonText As String) As Object
    On Error Resume Next
    
#If UseCallByName Then
    Set ParseJsonText = _
        CallByName( _
            GetHTMLDocumentForJson(), _
            "ParseJsonText", _
            VbMethod, _
            JsonText)
#Else
    Set ParseJsonText = GetHTMLDocumentForJson().ParseJsonText(JsonText)
#End If
End Function

'
' GetJsonKeys
' - Returns an array containing all existing keys in a JSON object.
'

'
' JsonObject:
'   Required. The name of a JSON object
'

Public Function GetJsonKeys(JsonObject As Object) As Object
    On Error Resume Next
    
#If UseCallByName Then
    Set GetJsonKeys = _
        CallByName( _
            GetHTMLDocumentForJson(), _
            "GetKeys", _
            VbMethod, _
            JsonObject)
#Else
    Set GetJsonKeys = GetHTMLDocumentForJson().GetKeys(JsonObject)
#End If
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
#If UseCallByName Then
    GetJsonKeysLength = CallByName(JsonKeys, "length", VbGet)
#Else
    GetJsonKeysLength = GetHTMLDocumentForJson().GetLength(JsonKeys)
    'GetJsonKeysLength = JsonKeys.Length
#End If
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
#If UseCallByName Then
    GetJsonKeysItem = CallByName(JsonKeys, Index, VbGet)
#Else
    GetJsonKeysItem = GetHTMLDocumentForJson().GetItem(JsonKeys, Index)
#End If
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
    
#If UseCallByName Then
    IsJsonItemObject = IsObject(CallByName(JsonObject, Key, VbGet))
#Else
    IsJsonItemObject = _
        IsObject(GetHTMLDocumentForJson().GetItem(JsonObject, Key))
#End If
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
    
#If UseCallByName Then
    GetJsonItemValue = CallByName(JsonObject, Key, VbGet)
#Else
    GetJsonItemValue = GetHTMLDocumentForJson().GetItem(JsonObject, Key)
#End If
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
    
#If UseCallByName Then
    Set GetJsonItemObject = CallByName(JsonObject, Key, VbGet)
#Else
    Set GetJsonItemObject = GetHTMLDocumentForJson().GetItem(JsonObject, Key)
#End If
End Function

'
' IsJsonArray
' - Returns whether JSON object is Array.
'

'
' JsonObject:
'   Required. The name of a JSON object
'

Public Function IsJsonArray(JsonObject As Object) As Boolean
#If UseCallByName Then
    IsJsonArray = _
        CallByName( _
            GetHTMLDocumentForJson(), _
            "IsJsonArray", _
            VbMethod, _
            JsonObject)
#Else
    IsJsonArray = GetHTMLDocumentForJson().IsJsonArray(JsonObject)
#End If
End Function

'
' GetJsonText
' - Returns string expression that identifies JSON data.
'

'
' JsonObject:
'   Required. The name of a JSON object
'

Public Function GetJsonText(JsonObject As Object) As String
#If UseCallByName Then
    GetJsonText = _
        CallByName( _
            GetHTMLDocumentForJson(), _
            "GetJsonText", _
            VbMethod, _
            JsonObject)
#Else
    GetJsonText = GetHTMLDocumentForJson().GetJsonText(JsonObject)
#End If
End Function

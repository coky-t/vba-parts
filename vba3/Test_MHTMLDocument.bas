Attribute VB_Name = "Test_MHTMLDocument"
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

Public Sub Test_ParseJsonText_Core(JsonText)
    Test_ParseJsonText_JsonObject JsonText
    Test_ParseJsonText_VbaObject1 JsonText
    Test_ParseJsonText_VbaObject2 JsonText
End Sub

Public Sub Test_ParseJsonText_JsonObject(JsonText)
    Debug_Print "==="
    Debug_Print JsonText
    Debug_Print "==="
    
    Dim JsonObject
    Set JsonObject = ParseJsonText(JsonText)
    
    If IsJsonArray(JsonObject) Then
        Debug_Print "=== Array ==="
    Else
        Debug_Print "=== Object ==="
    End If
    Debug_Print_JsonObject JsonObject
    
    Debug_Print "==="
    Debug_Print GetJsonText(JsonObject)
    Debug_Print "==="
End Sub

Public Sub Test_ParseJsonText_VbaObject1(JsonText)
    Debug_Print "==="
    Debug_Print JsonText
    Debug_Print "==="
    
    Dim VbaObject
    Set VbaObject = GetVbaObjectFromJsonText(JsonText)
    
    Debug_Print "=== " & TypeName(VbaObject) & " ==="
    Debug_Print_VbaObject VbaObject
    
    Debug_Print "==="
    Debug_Print GetJsonTextFromVbaObject(VbaObject)
    Debug_Print "==="
End Sub

Public Sub Test_ParseJsonText_VbaObject2(JsonText)
    Debug_Print "==="
    Debug_Print JsonText
    Debug_Print "==="
    
    Dim VbaObject
    Set VbaObject = GetVbaObjectFromJsonText(JsonText)
    
    Debug_Print "=== " & TypeName(VbaObject) & " ==="
    Debug_Print_VbaObject VbaObject
    
    Dim JsonObject
    Set JsonObject = GetJsonObjectFromVbaObject(VbaObject)
    
    If IsJsonArray(JsonObject) Then
        Debug_Print "=== Array ==="
    Else
        Debug_Print "=== Object ==="
    End If
    Debug_Print_JsonObject JsonObject
    
    Debug_Print "==="
    Debug_Print GetJsonText(JsonObject)
    Debug_Print "==="
End Sub

Public Sub Debug_Print_JsonObject(JsonObject)
    Dim Keys
    Set Keys = GetJsonKeys(JsonObject)
    
    Dim KeysLength
    KeysLength = GetJsonKeysLength(Keys)
    
    Dim Index
    Dim Key
    Dim Value
    For Index = 0 To KeysLength - 1
        Key = GetJsonKeysItem(Keys, Index)
        If IsJsonItemObject(JsonObject, Key) Then
            Dim JsonItemObject
            Set JsonItemObject = GetJsonItemObject(JsonObject, Key)
            If IsJsonArray(JsonItemObject) Then
                Debug_Print Key & " --- Array ---"
            Else
                Debug_Print Key & " --- Object ---"
            End If
            Debug_Print_JsonObject JsonItemObject
            Debug_Print Key & " ---"
        Else
            Value = CStr(GetJsonItemValue(JsonObject, Key))
            Debug_Print Key & ": " & Value
        End If
    Next
End Sub

Private Sub Debug_Print_VbaObject(VbaObject)
    Select Case TypeName(VbaObject)
    
    Case "Collection"
        Debug_Print_VbaCollection VbaObject
        
    Case "Dictionary"
        Debug_Print_VbaDictionary VbaObject
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Sub

Private Sub Debug_Print_VbaCollection(VbaCollection)
    Dim Index
    For Index = 1 To VbaCollection.Count
        If IsObject(VbaCollection.Item(Index)) Then
            Debug_Print _
                CStr(Index) & " --- " & _
                TypeName(VbaCollection.Item(Index)) & " ---"
            Debug_Print_VbaObject VbaCollection.Item(Index)
            Debug_Print CStr(Index) & " ---"
        Else
            Debug_Print CStr(Index) & ": " & CStr(VbaCollection.Item(Index))
        End If
    Next
End Sub

Private Sub Debug_Print_VbaDictionary(VbaDictionary)
    Dim VbaDicKeys
    VbaDicKeys = VbaDictionary.Keys
    
    Dim Index
    For Index = LBound(VbaDicKeys) To UBound(VbaDicKeys)
        Dim Key
        Key = VbaDicKeys(Index)
        If IsObject(VbaDictionary.Item(Key)) Then
            Debug_Print _
                Key & " --- " & _
                TypeName(VbaDictionary.Item(Key)) & " ---"
            Debug_Print_VbaObject VbaDictionary.Item(Key)
            Debug_Print Key & " ---"
        Else
            Debug_Print Key & ": " & CStr(VbaDictionary.Item(Key))
        End If
    Next
End Sub

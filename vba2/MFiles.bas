Attribute VB_Name = "MFiles"
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
' Microsoft Scripting Runtime
' - Scripting.File
' - Scripting.Dictionary
'

'
' --- Files ---
'

Public Function FilterFiles( _
    Files As Collection, _
    Optional IgnoredExtNames As String, _
    Optional SizeLimit As Long) As Collection
    
    Set FilterFiles = Files
    
    If Files Is Nothing Then Exit Function
    If Files.Count = 0 Then Exit Function
    If IgnoredExtNames = "" And SizeLimit = 0 Then Exit Function
    
    Dim IgnoredExtNameDic As Object
    Set IgnoredExtNameDic = GetIgnoredExtNameDic(IgnoredExtNames)
    
    Dim Files_ As Collection
    Set Files_ = New Collection
    
    Dim File As Object
    For Each File In Files
        If Not IsIgnoredFile(File, IgnoredExtNameDic, SizeLimit) Then
            Files_.Add File
        End If
    Next
    
    Set FilterFiles = Files_
End Function

Private Function GetIgnoredExtNameDic(IgnoredExtNames As String) _
    As Object
    
    If IgnoredExtNames = "" Then Exit Function
    
    Dim IgnoredExtNameDic As Object
    Set IgnoredExtNameDic = CreateObject("Scripting.Dictionary")
    
    Dim IgnoredExtNameArray As Variant
    IgnoredExtNameArray = Split(IgnoredExtNames, vbNewLine)
    
    Dim Index As Long
    For Index = LBound(IgnoredExtNameArray) To UBound(IgnoredExtNameArray)
        Dim IgnoredExtName As String
        IgnoredExtName = IgnoredExtNameArray(Index)
        
        If IgnoredExtName <> "" Then
            If Not IgnoredExtNameDic.Exists(IgnoredExtName) Then
                IgnoredExtNameDic.Add IgnoredExtName, IgnoredExtName
            End If
        End If
    Next
    
    Set GetIgnoredExtNameDic = IgnoredExtNameDic
End Function

Private Function IsIgnoredFile( _
    File As Object, _
    Optional IgnoredExtNameDic As Object, _
    Optional SizeLimit As Long) As Boolean
    
    IsIgnoredFile = _
        IsIgnoredFile_ExtNames(File, IgnoredExtNameDic) Or _
        IsIgnoredFile_SizeLimit(File, SizeLimit)
End Function

Private Function IsIgnoredFile_ExtNames( _
    File As Object, _
    IgnoredExtNameDic As Object) As Boolean
    
    If File Is Nothing Then Exit Function
    If IgnoredExtNameDic Is Nothing Then Exit Function
    If IgnoredExtNameDic.Count = 0 Then Exit Function
    
    Dim ExtName As String
    ExtName = GetFileSystemObject().GetExtensionName(File.Path)
    
    IsIgnoredFile_ExtNames = IgnoredExtNameDic.Exists(ExtName)
End Function

Private Function IsIgnoredFile_SizeLimit( _
    File As Object, _
    SizeLimit As Long) As Boolean
    
    If File Is Nothing Then Exit Function
    If SizeLimit <= 0 Then Exit Function
    
    IsIgnoredFile_SizeLimit = (File.Size > SizeLimit)
End Function

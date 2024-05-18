Attribute VB_Name = "MADODBRecordset"
Option Explicit

'
' Copyright (c) 2024 Koki Takeyama
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
' Microsoft ActiveX Data Objects X.X Library
' - ADODB.Recordset
'

Public Function StrArray_Sort(StrArray, Order)
    If Not IsArray(StrArray) Then
        StrArray_Sort = StrArray
        Exit Function
    End If
    
    Const adLongVarChar = 201
    
    Dim RecSet
    Set RecSet = CreateObject("ADODB.Recordset")
    With RecSet
        .Fields.Append "Key", adLongVarChar, 4096
        .Open
    End With
    
    Dim LB
    Dim UB
    LB = LBound(StrArray)
    UB = UBound(StrArray)
    
    Dim Index
    For Index = LB To UB
        RecSet.AddNew
        RecSet("Key").Value = StrArray(Index)
        RecSet.Update
    Next
    
    RecSet.Sort = "Key " & Order
    
    Dim Keys()
    ReDim Keys(UB)
    
    For Index = LB To UB
        Keys(Index) = RecSet("Key").Value
        RecSet.MoveNext
    Next
    
    RecSet.Close
    
    StrArray_Sort = Keys
End Function

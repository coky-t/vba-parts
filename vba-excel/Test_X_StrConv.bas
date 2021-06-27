Attribute VB_Name = "Test_X_StrConv"
Option Explicit

'
' Copyright (c) 2021 Koki Takeyama
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
' Test
'

' ActiveWorkbook

Public Sub Test_ActiveWorkbook_StrConv_UpperCase()
    Selection_StrConv ActiveWorkbook, vbUpperCase
End Sub

Public Sub Test_ActiveWorkbook_StrConv_LowerCase()
    Selection_StrConv ActiveWorkbook, vbLowerCase
End Sub

Public Sub Test_ActiveWorkbook_StrConv_ProperCase()
    Selection_StrConv ActiveWorkbook, vbProperCase
End Sub

Public Sub Test_ActiveWorkbook_StrConv_Wide()
    Selection_StrConv ActiveWorkbook, vbWide
End Sub

Public Sub Test_ActiveWorkbook_StrConv_Narrow()
    Selection_StrConv ActiveWorkbook, vbNarrow
End Sub

Public Sub Test_ActiveWorkbook_StrConv_Katakana()
    Selection_StrConv ActiveWorkbook, vbKatakana
End Sub

Public Sub Test_ActiveWorkbook_StrConv_Hiragana()
    Selection_StrConv ActiveWorkbook, vbHiragana
End Sub

' ActiveSheet

Public Sub Test_ActiveSheet_StrConv_UpperCase()
    Selection_StrConv ActiveSheet, vbUpperCase
End Sub

Public Sub Test_ActiveSheet_StrConv_LowerCase()
    Selection_StrConv ActiveSheet, vbLowerCase
End Sub

Public Sub Test_ActiveSheet_StrConv_ProperCase()
    Selection_StrConv ActiveSheet, vbProperCase
End Sub

Public Sub Test_ActiveSheet_StrConv_Wide()
    Selection_StrConv ActiveSheet, vbWide
End Sub

Public Sub Test_ActiveSheet_StrConv_Narrow()
    Selection_StrConv ActiveSheet, vbNarrow
End Sub

Public Sub Test_ActiveSheet_StrConv_Katakana()
    Selection_StrConv ActiveSheet, vbKatakana
End Sub

Public Sub Test_ActiveSheet_StrConv_Hiragana()
    Selection_StrConv ActiveSheet, vbHiragana
End Sub

' Selection

Public Sub Test_Selection_StrConv_UpperCase()
    Selection_StrConv Selection, vbUpperCase
End Sub

Public Sub Test_Selection_StrConv_LowerCase()
    Selection_StrConv Selection, vbLowerCase
End Sub

Public Sub Test_Selection_StrConv_ProperCase()
    Selection_StrConv Selection, vbProperCase
End Sub

Public Sub Test_Selection_StrConv_Wide()
    Selection_StrConv Selection, vbWide
End Sub

Public Sub Test_Selection_StrConv_Narrow()
    Selection_StrConv Selection, vbNarrow
End Sub

Public Sub Test_Selection_StrConv_Katakana()
    Selection_StrConv Selection, vbKatakana
End Sub

Public Sub Test_Selection_StrConv_Hiragana()
    Selection_StrConv Selection, vbHiragana
End Sub

'
' Debug
'

Public Sub Debug_Print_TypeName_Selection()
    Debug.Print TypeName(Selection)
End Sub

Public Sub Debug_Print_Selection_Count()
    Debug.Print Selection.Count
End Sub

Private Sub Debug_Print_TypeName_ActiveWorkbook()
    Debug.Print TypeName(ActiveWorkbook)
End Sub

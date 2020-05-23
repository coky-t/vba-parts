Attribute VB_Name = "MExcelApp"
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
' Microsoft Excel X.X Object Library
' - Excel.Application
'

'
' --- Excel.Application ---
'

'
' GetExcelApplication
' - Returns a Excel.Application object.
'

'
' ExcelApplication:
'   Optional. The name of a Excel.Application object.
'

Public Function GetExcelApplication( _
    ExcelApplication)
    
    If ExcelApplication Is Nothing Then
        Set GetExcelApplication = CreateObject("Excel.Application")
    Else
        Set GetExcelApplication = ExcelApplication
    End If
End Function

'
' --- File DialogBox ---
'

'
' GetOpenFileName
' - Displays the standard Open dialog box and gets a file name from the user
'   without actually opening any files.
'

'
' ExcelApplication:
'   Optional. The name of a Excel.Application object.
'
' FileFilter:
'   Optional. A string specifying file filtering criteria.
'
' FilterIndex:
'   Optional. Specifies the index numbers of the default file filtering
'   criteria, from 1 to the number of filters specified in FileFilter.
'   If this argument is omitted or greater than the number of filters
'   present, the first file filter is used.
'
' Title:
'   Optional. Specifies the title of the dialog box.
'   If this argument is omitted, the default title is used.
'

Public Function GetOpenFileName( _
    ExcelApplication, _
    FileFilter, _
    FilterIndex, _
    Title)
    
    On Error Resume Next
    
    Dim OpenFileName
    OpenFileName = _
        GetExcelApplication(ExcelApplication) _
        .GetOpenFileName(FileFilter, FilterIndex, Title)
    If OpenFileName = False Then
        ' nop
    Else
        GetOpenFileName = CStr(OpenFileName)
    End If
End Function

'
' GetSaveAsFileName
' - Displays the standard Save As dialog box and gets a file name from
'   the user without actually saving any files.
'

'
' ExcelApplication:
'   Optional. The name of a Excel.Application object.
'
' InitialFileName:
'   Optional. Specifies the suggested file name.
'
' FileFilter:
'   Optional. A string specifying file filtering criteria.
'
' FilterIndex:
'   Optional. Specifies the index numbers of the default file filtering
'   criteria, from 1 to the number of filters specified in FileFilter.
'   If this argument is omitted or greater than the number of filters
'   present, the first file filter is used.
'
' Title:
'   Optional. Specifies the title of the dialog box.
'   If this argument is omitted, the default title is used.
'

Public Function GetSaveAsFileName( _
    ExcelApplication, _
    InitialFileName, _
    FileFilter, _
    FilterIndex, _
    Title)
    
    On Error Resume Next
    
    Dim SaveAsFileName
    SaveAsFileName = _
        GetExcelApplication(ExcelApplication) _
        .GetSaveAsFileName(InitialFileName, FileFilter, FilterIndex, Title)
    If SaveAsFileName = False Then
        ' nop
    Else
        GetSaveAsFileName = CStr(SaveAsFileName)
    End If
End Function

'
' --- Test ---
'

Private Sub Test_GetOpenFileName()
    Dim FileName
    FileName = GetOpenFileName(Nothing, "", 0, "")
    Debug_Print FileName
End Sub

Private Sub Test_GetSaveAsFileName()
    Dim FileName
    FileName = GetSaveAsFileName(Nothing, "", "", 0, "")
    Debug_Print FileName
End Sub

Private Sub Debug_Print(Str)
    Debug.Print Str
End Sub

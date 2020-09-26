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

#Const UseCurrentInstance = True

Private ExcelApplication

'
' --- Excel.Application ---
'

'
' GetExcelApplication
' - Returns a Excel.Application object.
'

Public Function GetExcelApplication()
    'Static ExcelApplication
    If IsEmpty(ExcelApplication) Then
#If UseCurrentInstance Then
        Set ExcelApplication = Excel.Application
#Else
        Set ExcelApplication = CreateObject("Excel.Application")
#End If
    End If
    Set GetExcelApplication = ExcelApplication
End Function

'
' --- File Dialog Box ---
'

'
' GetOpenFileName
' - Displays the standard Open dialog box and gets a file name from the user
'   without actually opening any files.
'

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
    FileFilter, _
    FilterIndex, _
    Title)
    
    On Error Resume Next
    
    Dim OpenFileName
    OpenFileName = _
        GetExcelApplication().GetOpenFileName(FileFilter, FilterIndex, Title)
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
    InitialFileName, _
    FileFilter, _
    FilterIndex, _
    Title)
    
    On Error Resume Next
    
    Dim SaveAsFileName
    SaveAsFileName = _
        GetExcelApplication() _
        .GetSaveAsFileName(InitialFileName, FileFilter, FilterIndex, Title)
    If SaveAsFileName = False Then
        ' nop
    Else
        GetSaveAsFileName = CStr(SaveAsFileName)
    End If
End Function

'
' --- Folder Dialog Box ---
'

'
' GetFolderName
' - Displays the standard Open dialog box and gets a folder name.
'

'
' Title:
'   Optional. Specifies the title of the dialog box.
'   If this argument is omitted, the default title is used.
'

Public Function GetFolderName(Title)
    On Error Resume Next
    
    With GetExcelApplication()
        With .FileDialog(4) 'msoFileDialogFolderPicker
            If Title <> "" Then .Title = Title
            If .Show = -1 Then GetFolderName = CStr(.SelectedItems(1))
        End With
    End With
End Function

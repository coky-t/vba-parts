VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProgressDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Counter As Long
Public CounterEnd As Long

Public IndicatorDoneSymbol As String
Public IndicatorNotYetSymbol As String
Public IndicatorEnd As Long

Public Title As String
Public Comment As String

Private DisplayStatusBar As Boolean

Private Sub Class_Initialize()
    Reset
    StatusBar_Initialize
End Sub

Private Sub Class_Terminate()
    StatusBar_Terminate
End Sub

Private Sub Reset()
    Counter = 0
    CounterEnd = 100
    
    IndicatorDoneSymbol = "###"
    IndicatorNotYetSymbol = "___"
    IndicatorEnd = 10
    
    Title = ""
    Comment = ""
End Sub

Private Sub StatusBar_Initialize()
    DisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
End Sub

Private Sub StatusBar_Terminate()
    Application.StatusBar = False
    Application.DisplayStatusBar = DisplayStatusBar
End Sub

Public Sub Increment()
    Counter = Counter + 1
    Display GetIndicator, GetPercent
End Sub

Private Function GetIndicator() As String
    Dim Indicator As String
    
    Dim IndicatorDoneCount As Long
    IndicatorDoneCount = (Counter * IndicatorEnd) \ CounterEnd
    
    Dim Index As Long
    For Index = 1 To IndicatorDoneCount
        Indicator = Indicator & IndicatorDoneSymbol
    Next
    For Index = IndicatorDoneCount + 1 To IndicatorEnd
        Indicator = Indicator & IndicatorNotYetSymbol
    Next
    
    GetIndicator = Indicator
End Function

Private Function GetPercent() As String
    GetPercent = Space(1) & CStr((Counter * 100) \ CounterEnd) & "%"
End Function

Public Sub Display(Indicator As String, Percent As String)
    Application.StatusBar = Title & Indicator & Percent & Comment
    DoEvents
End Sub

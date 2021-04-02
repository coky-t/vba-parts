Attribute VB_Name = "Test_CProgressDisplay"
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
' --- Test ---
'

Public Sub Test_CProgressDisplay1()
    Dim PDisp As CProgressDisplay
    Set PDisp = New CProgressDisplay
    PDisp.CounterEnd = 10
    
    Dim Index As Long
    For Index = 1 To 10
        PDisp.Increment
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Public Sub Test_CProgressDisplay2()
    Dim PDisp As CProgressDisplay
    Set PDisp = New CProgressDisplay
    PDisp.CounterEnd = 10
    PDisp.IndicatorNotYetSymbol = ""
    
    Dim Index As Long
    For Index = 1 To 10
        PDisp.Increment
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Public Sub Test_CProgressDisplay3()
    Dim PDisp As CProgressDisplay
    Set PDisp = New CProgressDisplay
    PDisp.CounterEnd = 20
    PDisp.IndicatorDoneSymbol = "+"
    PDisp.IndicatorNotYetSymbol = "-"
    
    Dim Index As Long
    For Index = 1 To 20
        PDisp.Increment
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Public Sub Test_CProgressDisplay4()
    Dim PDisp As CProgressDisplay
    Set PDisp = New CProgressDisplay
    PDisp.CounterEnd = 10
    PDisp.IndicatorDoneSymbol = "+"
    PDisp.IndicatorNotYetSymbol = "-"
    PDisp.IndicatorEnd = 20
    
    Dim Index As Long
    For Index = 1 To 10
        PDisp.Increment
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Public Sub Test_CProgressDisplay5()
    Dim PDisp As CProgressDisplay
    Set PDisp = New CProgressDisplay
    PDisp.CounterEnd = 30
    PDisp.IndicatorDoneSymbol = "#"
    PDisp.IndicatorNotYetSymbol = "="
    PDisp.IndicatorEnd = 20
    PDisp.Title = "Test: "
    
    Dim Index As Long
    For Index = 1 To 30
        PDisp.Increment
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Public Sub Test_CProgressDisplay6()
    Dim PDisp As CProgressDisplay
    Set PDisp = New CProgressDisplay
    PDisp.CounterEnd = 30
    PDisp.IndicatorDoneSymbol = "#"
    PDisp.IndicatorNotYetSymbol = "="
    PDisp.IndicatorEnd = 20
    PDisp.Title = "Test: "
    
    Dim Index As Long
    For Index = 1 To 30
        PDisp.Comment = Space(1) & CStr(Index) & "/30"
        PDisp.Increment
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProgressTimer"
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

Private m_StartTime

Private Sub Class_Initialize()
    m_StartTime = Now
End Sub

Private Sub Class_Terminate()
    Debug.Print "Elapsed time: " & FormatTimeDiff(m_StartTime, Now)
End Sub

Public Sub Reset()
    m_StartTime = Now
End Sub

Private Function FormatTimeDiff(StartTime, EndTime)
    Dim SecondDiff
    SecondDiff = DateDiff("s", StartTime, EndTime)
    
    Dim MinuteDiff
    MinuteDiff = SecondDiff \ 60
    SecondDiff = SecondDiff Mod 60
    
    Dim HourDiff
    HourDiff = MinuteDiff \ 60
    MinuteDiff = MinuteDiff Mod 60
    
    Dim TimeDiffString
    TimeDiffString = _
        Right("0" & HourDiff, 2) & ":" & _
        Right("0" & MinuteDiff, 2) & ":" & _
        Right("0" & SecondDiff, 2)
    
    FormatTimeDiff = TimeDiffString
End Function

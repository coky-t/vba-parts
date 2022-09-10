Attribute VB_Name = "StrDiff2"
Option Explicit

'
' Copyright (c) 2022 Koki Takeyama
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
' === String Difference 2 - O(ND) Implementation ===
'

Function EditDistance(Str1 As String, Str2 As String) As Long
    Dim Len1 As Long
    Dim Len2 As Long
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    If Len1 = 0 Then
        EditDistance = Len2
        Exit Function
    End If
    If Len2 = 0 Then
        EditDistance = Len1
        Exit Function
    End If
    
    Dim MaxCost As Long
    MaxCost = Len1 + Len2
    
    Dim Index2() As Long
    ReDim Index2(-MaxCost To MaxCost)
    
    Dim Index0 As Long
    
    Dim MinIndex0 As Long
    Dim MaxIndex0 As Long
    
    Dim TempIndex1 As Long
    Dim TempIndex2 As Long
    
    Dim Cost As Long
    For Cost = 0 To MaxCost
        MinIndex0 = -Cost
        MaxIndex0 = Cost
        
        For Index0 = MinIndex0 To MaxIndex0 Step 2
            If Cost = 0 Then
                TempIndex2 = 0
                
            ElseIf Index0 = MinIndex0 Then
                TempIndex2 = Index2(Index0 + 1)
                
            ElseIf Index0 = MaxIndex0 Then
                TempIndex2 = Index2(Index0 - 1) + 1
                
            Else
                TempIndex2 = _
                    LongMax2(Index2(Index0 + 1), Index2(Index0 - 1) + 1)
                    
            End If
            
            TempIndex1 = TempIndex2 - Index0
            Do While TempIndex1 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex1 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempIndex1 = TempIndex1 + 1
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            If TempIndex1 >= Len1 And TempIndex2 >= Len2 Then
                EditDistance = Cost
                Exit Function
            End If
            Index2(Index0) = TempIndex2
        Next
    Next
End Function

Function LongestCommonSubsequence(Str1 As String, Str2 As String) As String
    Dim Len1 As Long
    Dim Len2 As Long
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    If Len1 = 0 Then
        LongestCommonSubsequence = ""
        Exit Function
    End If
    If Len2 = 0 Then
        LongestCommonSubsequence = ""
        Exit Function
    End If
    
    Dim MaxCost As Long
    MaxCost = Len1 + Len2
    
    Dim Index2() As Long
    ReDim Index2(-MaxCost To MaxCost)
    
    Dim LCS() As String
    ReDim LCS(-MaxCost To MaxCost)
    
    Dim Index0 As Long
    
    Dim MinIndex0 As Long
    Dim MaxIndex0 As Long
    
    Dim TempIndex1 As Long
    Dim TempIndex2 As Long
    
    Dim TempLCS As String
    
    Dim Cost As Long
    For Cost = 0 To MaxCost
        MinIndex0 = -Cost
        MaxIndex0 = Cost
        
        For Index0 = MinIndex0 To MaxIndex0 Step 2
            If Cost = 0 Then
                TempIndex2 = 0
                TempLCS = ""
                
            ElseIf Index0 = MinIndex0 Then
                TempIndex2 = Index2(Index0 + 1)
                TempLCS = LCS(Index0 + 1)
                
            ElseIf Index0 = MaxIndex0 Then
                TempIndex2 = Index2(Index0 - 1) + 1
                TempLCS = LCS(Index0 - 1)
                
            ElseIf Index2(Index0 + 1) > Index2(Index0 - 1) + 1 Then
                TempIndex2 = Index2(Index0 + 1)
                TempLCS = LCS(Index0 + 1)
                
            Else
                TempIndex2 = Index2(Index0 - 1) + 1
                TempLCS = LCS(Index0 - 1)
                    
            End If
            
            TempIndex1 = TempIndex2 - Index0
            Do While TempIndex1 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex1 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempLCS = TempLCS & Mid(Str1, TempIndex1 + 1, 1)
                    
                    TempIndex1 = TempIndex1 + 1
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            If TempIndex1 >= Len1 And TempIndex2 >= Len2 Then
                LongestCommonSubsequence = TempLCS
                Exit Function
            End If
            Index2(Index0) = TempIndex2
            LCS(Index0) = TempLCS
        Next
    Next
End Function

Function ShortestEditScript(Str1 As String, Str2 As String) As String
    Dim Len1 As Long
    Dim Len2 As Long
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    If (Len1 = 0) And (Len2 = 0) Then
        ShortestEditScript = ""
        Exit Function
        
    ElseIf Len2 = 0 Then
        Dim SCSTemp1 As String
        Dim Index1Temp As Long
        For Index1Temp = 1 To Len1
            SCSTemp1 = SCSTemp1 & "-" & Mid(Str1, Index1Temp, 1)
        Next
        ShortestEditScript = SCSTemp1
        Exit Function
        
    ElseIf Len1 = 0 Then
        Dim SCSTemp2 As String
        Dim Index2Temp As Long
        For Index2Temp = 1 To Len2
            SCSTemp2 = SCSTemp2 & "+" & Mid(Str2, Index2Temp, 1)
        Next
        ShortestEditScript = SCSTemp2
        Exit Function
        
    End If
    
    Dim MaxCost As Long
    MaxCost = Len1 + Len2
    
    Dim Index2() As Long
    ReDim Index2(-MaxCost To MaxCost)
    
    Dim SES() As String
    ReDim SES(-MaxCost To MaxCost)
    
    Dim Index0 As Long
    
    Dim MinIndex0 As Long
    Dim MaxIndex0 As Long
    
    Dim TempIndex1 As Long
    Dim TempIndex2 As Long
    
    Dim TempSES As String
    
    Dim Cost As Long
    For Cost = 0 To MaxCost
        MinIndex0 = -Cost
        MaxIndex0 = Cost
        
        For Index0 = MinIndex0 To MaxIndex0 Step 2
            If Cost = 0 Then
                TempIndex2 = 0
                TempSES = ""
                
            ElseIf Index0 = MinIndex0 Then
                TempIndex2 = Index2(Index0 + 1)
                TempSES = _
                    SES(Index0 + 1) & "-" & Mid(Str1, TempIndex2 - Index0, 1)
                
            ElseIf Index0 = MaxIndex0 Then
                TempIndex2 = Index2(Index0 - 1) + 1
                TempSES = SES(Index0 - 1) & "+" & Mid(Str2, TempIndex2, 1)
                
            ElseIf Index2(Index0 + 1) > Index2(Index0 - 1) + 1 Then
                TempIndex2 = Index2(Index0 + 1)
                TempSES = _
                    SES(Index0 + 1) & "-" & Mid(Str1, TempIndex2 - Index0, 1)
                
            Else
                TempIndex2 = Index2(Index0 - 1) + 1
                TempSES = SES(Index0 - 1) & "+" & Mid(Str2, TempIndex2, 1)
                    
            End If
            
            TempIndex1 = TempIndex2 - Index0
            Do While TempIndex1 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex1 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempSES = TempSES & " " & Mid(Str1, TempIndex1 + 1, 1)
                    
                    TempIndex1 = TempIndex1 + 1
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            If TempIndex1 >= Len1 And TempIndex2 >= Len2 Then
                ShortestEditScript = TempSES
                Exit Function
            End If
            Index2(Index0) = TempIndex2
            SES(Index0) = TempSES
        Next
    Next
End Function

Private Function LongMax2(Lng1 As Long, Lng2 As Long) As Long
    LongMax2 = IIf(Lng1 > Lng2, Lng1, Lng2)
End Function

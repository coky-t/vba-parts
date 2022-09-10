Attribute VB_Name = "StrDiff3"
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
' === String Difference 3 - O(NP) Implementation ===
'

Function EditDistance(Str1, Str2)
    If Len(Str1) < Len(Str2) Then
        EditDistance = EditDistanceCore(Str1, Str2)
    Else
        EditDistance = EditDistanceCore(Str2, Str1)
    End If
End Function

Private Function EditDistanceCore(Str1, Str2)
    Dim Len1
    Dim Len2
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    If Len1 = 0 Then
        EditDistanceCore = Len2
        Exit Function
    End If
    If Len2 = 0 Then
        EditDistanceCore = Len1
        Exit Function
    End If
    
    Dim Index2()
    ReDim Index2(Len1 + Len2)
    
    Dim LenDiff
    LenDiff = Len2 - Len1
    
    Dim Index0
    
    Dim TempIndex1
    Dim TempIndex2
    
    For TempIndex1 = 0 To Len1
        For Index0 = -TempIndex1 To LenDiff - 1
            If TempIndex1 = 0 Then
                If Index0 = 0 Then
                    TempIndex2 = 0
                Else
                    TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                End If
                
            ElseIf Index0 = -TempIndex1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                
            Else
                TempIndex2 = _
                    LongMax2(Index2(Len1 + Index0 + 1), _
                    Index2(Len1 + Index0 - 1) + 1)
                    
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            Index2(Len1 + Index0) = TempIndex2
        Next
        
        For Index0 = LenDiff + TempIndex1 To LenDiff Step -1
            If Index0 = LenDiff + TempIndex1 Then
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                
            Else
                TempIndex2 = _
                    LongMax2(Index2(Len1 + Index0 + 1), _
                    Index2(Len1 + Index0 - 1) + 1)
                    
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            Index2(Len1 + Index0) = TempIndex2
        Next
        
        For Index0 = LenDiff To LenDiff
            If TempIndex1 = 0 Then
                If Index0 = 0 Then
                    TempIndex2 = 0
                Else
                    TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                End If
                
            Else
                TempIndex2 = _
                    LongMax2(Index2(Len1 + Index0 + 1), _
                    Index2(Len1 + Index0 - 1) + 1)
                    
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            If TempIndex2 = Len2 Then
                EditDistanceCore = LenDiff + 2 * TempIndex1
                Exit Function
            End If
            
            Index2(Len1 + Index0) = TempIndex2
        Next
    Next
End Function

Function LongestCommonSubsequence(Str1, Str2)
    If Len(Str1) < Len(Str2) Then
        LongestCommonSubsequence = LongestCommonSubsequenceCore(Str1, Str2)
    Else
        LongestCommonSubsequence = LongestCommonSubsequenceCore(Str2, Str1)
    End If
End Function

Private Function LongestCommonSubsequenceCore(Str1, Str2)
    Dim Len1
    Dim Len2
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    If Len1 = 0 Then
        LongestCommonSubsequenceCore = ""
        Exit Function
    End If
    If Len2 = 0 Then
        LongestCommonSubsequenceCore = ""
        Exit Function
    End If
    
    Dim Index2()
    ReDim Index2(Len1 + Len2)
    
    Dim LCS()
    ReDim LCS(Len1 + Len2)
    
    Dim LenDiff
    LenDiff = Len2 - Len1
    
    Dim Index0
    
    Dim TempIndex1
    Dim TempIndex2
    
    Dim TempLCS
    
    For TempIndex1 = 0 To Len1
        For Index0 = -TempIndex1 To LenDiff - 1
            If TempIndex1 = 0 Then
                If Index0 = 0 Then
                    TempIndex2 = 0
                    TempLCS = ""
                Else
                    TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                    TempLCS = LCS(Len1 + Index0 - 1)
                End If
                
            ElseIf Index0 = -TempIndex1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempLCS = LCS(Len1 + Index0 + 1)
                
            ElseIf Index2(Len1 + Index0 + 1) > _
                Index2(Len1 + Index0 - 1) + 1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempLCS = LCS(Len1 + Index0 + 1)
                
            Else
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempLCS = LCS(Len1 + Index0 - 1)
                
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempLCS = TempLCS & Mid(Str2, TempIndex2 + 1, 1)
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            Index2(Len1 + Index0) = TempIndex2
            LCS(Len1 + Index0) = TempLCS
        Next
        
        For Index0 = LenDiff + TempIndex1 To LenDiff Step -1
            If Index0 = LenDiff + TempIndex1 Then
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempLCS = LCS(Len1 + Index0 - 1)
                
            ElseIf Index2(Len1 + Index0 + 1) > _
                Index2(Len1 + Index0 - 1) + 1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempLCS = LCS(Len1 + Index0 + 1)
                
            Else
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempLCS = LCS(Len1 + Index0 - 1)
                
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempLCS = TempLCS & Mid(Str2, TempIndex2 + 1, 1)
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            Index2(Len1 + Index0) = TempIndex2
            LCS(Len1 + Index0) = TempLCS
        Next
        
        For Index0 = LenDiff To LenDiff
            If TempIndex1 = 0 Then
                If Index0 = 0 Then
                    TempIndex2 = 0
                    TempLCS = ""
                Else
                    TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                    TempLCS = LCS(Len1 + Index0 - 1)
                End If
                
            ElseIf Index2(Len1 + Index0 + 1) > _
                Index2(Len1 + Index0 - 1) + 1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempLCS = LCS(Len1 + Index0 + 1)
                
            Else
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempLCS = LCS(Len1 + Index0 - 1)
                
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempLCS = TempLCS & Mid(Str2, TempIndex2 + 1, 1)
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            If TempIndex2 = Len2 Then
                LongestCommonSubsequenceCore = TempLCS
                Exit Function
            End If
            
            Index2(Len1 + Index0) = TempIndex2
            LCS(Len1 + Index0) = TempLCS
        Next
    Next
End Function

Function ShortestEditScript(Str1, Str2)
    If Len(Str1) < Len(Str2) Then
        ShortestEditScript = ShortestEditScriptCore(Str1, Str2, "-", "+")
    Else
        ShortestEditScript = ShortestEditScriptCore(Str2, Str1, "+", "-")
    End If
End Function

Private Function ShortestEditScriptCore(Str1, Str2, EditChar1, EditChar2)
    Dim Len1
    Dim Len2
    Len1 = Len(Str1)
    Len2 = Len(Str2)
    
    If (Len1 = 0) And (Len2 = 0) Then
        ShortestEditScriptCore = ""
        Exit Function
        
    ElseIf Len2 = 0 Then
        Dim SCSTemp1
        Dim Index1Temp
        For Index1Temp = 1 To Len1
            SCSTemp1 = SCSTemp1 & EditChar1 & Mid(Str1, Index1Temp, 1)
        Next
        ShortestEditScriptCore = SCSTemp1
        Exit Function
        
    ElseIf Len1 = 0 Then
        Dim SCSTemp2
        Dim Index2Temp
        For Index2Temp = 1 To Len2
            SCSTemp2 = SCSTemp2 & EditChar2 & Mid(Str2, Index2Temp, 1)
        Next
        ShortestEditScriptCore = SCSTemp2
        Exit Function
        
    End If
    
    Dim Index2()
    ReDim Index2(Len1 + Len2)
    
    Dim SES()
    ReDim SES(Len1 + Len2)
    
    Dim LenDiff
    LenDiff = Len2 - Len1
    
    Dim Index0
    
    Dim TempIndex1
    Dim TempIndex2
    
    Dim TempSES
    
    For TempIndex1 = 0 To Len1
        For Index0 = -TempIndex1 To LenDiff - 1
            If TempIndex1 = 0 Then
                If Index0 = 0 Then
                    TempIndex2 = 0
                    TempSES = ""
                Else
                    TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                    TempSES = _
                        SES(Len1 + Index0 - 1) & EditChar2 & _
                        Mid(Str2, TempIndex2, 1)
                End If
                
            ElseIf Index0 = -TempIndex1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempSES = _
                    SES(Len1 + Index0 + 1) & EditChar1 & _
                    Mid(Str1, TempIndex2 - Index0, 1)
                
            ElseIf Index2(Len1 + Index0 + 1) > _
                Index2(Len1 + Index0 - 1) + 1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempSES = _
                    SES(Len1 + Index0 + 1) & EditChar1 & _
                    Mid(Str1, TempIndex2 - Index0, 1)
                
            Else
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempSES = _
                    SES(Len1 + Index0 - 1) & EditChar2 & _
                    Mid(Str2, TempIndex2, 1)
                
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempSES = TempSES & " " & Mid(Str2, TempIndex2 + 1, 1)
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            Index2(Len1 + Index0) = TempIndex2
            SES(Len1 + Index0) = TempSES
        Next
        
        For Index0 = LenDiff + TempIndex1 To LenDiff Step -1
            If Index0 = LenDiff + TempIndex1 Then
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempSES = _
                    SES(Len1 + Index0 - 1) & EditChar2 & _
                    Mid(Str2, TempIndex2, 1)
                
            ElseIf Index2(Len1 + Index0 + 1) > _
                Index2(Len1 + Index0 - 1) + 1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempSES = _
                    SES(Len1 + Index0 + 1) & EditChar1 & _
                    Mid(Str1, TempIndex2 - Index0, 1)
                
            Else
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempSES = _
                    SES(Len1 + Index0 - 1) & EditChar2 & _
                    Mid(Str2, TempIndex2, 1)
                
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempSES = TempSES & " " & Mid(Str2, TempIndex2 + 1, 1)
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            Index2(Len1 + Index0) = TempIndex2
            SES(Len1 + Index0) = TempSES
        Next
        
        For Index0 = LenDiff To LenDiff
            If TempIndex1 = 0 Then
                If Index0 = 0 Then
                    TempIndex2 = 0
                    TempSES = ""
                Else
                    TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                    TempSES = _
                        SES(Len1 + Index0 - 1) & EditChar2 & _
                        Mid(Str2, TempIndex2, 1)
                End If
                
            ElseIf Index2(Len1 + Index0 + 1) > _
                Index2(Len1 + Index0 - 1) + 1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempSES = _
                    SES(Len1 + Index0 + 1) & EditChar1 & _
                    Mid(Str1, TempIndex2 - Index0, 1)
                
            Else
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempSES = _
                    SES(Len1 + Index0 - 1) & EditChar2 & _
                    Mid(Str2, TempIndex2, 1)
                
            End If
            
            Do While TempIndex2 - Index0 < Len1 And TempIndex2 < Len2
                If Mid(Str1, TempIndex2 - Index0 + 1, 1) = _
                    Mid(Str2, TempIndex2 + 1, 1) Then
                    TempSES = TempSES & " " & Mid(Str2, TempIndex2 + 1, 1)
                    TempIndex2 = TempIndex2 + 1
                Else
                    Exit Do
                End If
            Loop
            
            If TempIndex2 = Len2 Then
                ShortestEditScriptCore = TempSES
                Exit Function
            End If
            
            Index2(Len1 + Index0) = TempIndex2
            SES(Len1 + Index0) = TempSES
        Next
    Next
End Function

Private Function LongMax2(Lng1, Lng2)
    LongMax2 = IIf(Lng1 > Lng2, Lng1, Lng2)
End Function

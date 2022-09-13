Attribute VB_Name = "StrArrayDiff2X"
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
' === String Array Difference 2 - O(ND) Implementation ===
'

'
' Modified - MinIndex0, MaxIndex0
'

Function EditDistance(Str1(), Str2())
    Dim LB1
    Dim UB1
    Dim Len1
    If Not IsError(Str1) Then
        LB1 = LBound(Str1)
        UB1 = UBound(Str1)
        Len1 = UB1 - LB1 + 1
    End If
    
    Dim LB2
    Dim UB2
    Dim Len2
    If Not IsError(Str2) Then
        LB2 = LBound(Str2)
        UB2 = UBound(Str2)
        Len2 = UB2 - LB2 + 1
    End If
    
    If Len1 = 0 Then
        EditDistance = Len2
        Exit Function
    End If
    If Len2 = 0 Then
        EditDistance = Len1
        Exit Function
    End If
    
    Dim MaxCost
    MaxCost = Len1 + Len2
    
    Dim Index2()
    ReDim Index2(Len1 + Len2)
    
    Dim Index0
    
    Dim MinIndex0
    Dim MaxIndex0
    
    Dim TempIndex1
    Dim TempIndex2
    
    Dim Cost
    For Cost = 0 To MaxCost
        MinIndex0 = -Cost
        MaxIndex0 = Cost
        
        If Cost > Len1 Then MinIndex0 = -Len1 + (Cost - Len1)
        If Cost > Len2 Then MaxIndex0 = Len2 - (Cost - Len2)
        
        For Index0 = MinIndex0 To MaxIndex0 Step 2
            If Cost = 0 Then
                TempIndex2 = 0
                
            ElseIf Index0 = MinIndex0 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                
            ElseIf Index0 = MaxIndex0 Then
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                
            Else
                TempIndex2 = _
                    LongMax2( _
                        Index2(Len1 + Index0 + 1), _
                        Index2(Len1 + Index0 - 1) + 1)
                    
            End If
            
            TempIndex1 = TempIndex2 - Index0
            Do While TempIndex1 < Len1 And TempIndex2 < Len2
                If Str1(LB1 + TempIndex1) = Str2(LB2 + TempIndex2) Then
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
            Index2(Len1 + Index0) = TempIndex2
        Next
    Next
End Function

Function LongestCommonSubsequence(Str1(), Str2())
    Dim LB1
    Dim UB1
    Dim Len1
    If Not IsError(Str1) Then
        LB1 = LBound(Str1)
        UB1 = UBound(Str1)
        Len1 = UB1 - LB1 + 1
    End If
    
    Dim LB2
    Dim UB2
    Dim Len2
    If Not IsError(Str2) Then
        LB2 = LBound(Str2)
        UB2 = UBound(Str2)
        Len2 = UB2 - LB2 + 1
    End If
    
    If Len1 = 0 Then
        LongestCommonSubsequence = ""
        Exit Function
    End If
    If Len2 = 0 Then
        LongestCommonSubsequence = ""
        Exit Function
    End If
    
    Dim MaxCost
    MaxCost = Len1 + Len2
    
    Dim Index2()
    ReDim Index2(Len1 + Len2)
    
    Dim LCS()
    ReDim LCS(Len1 + Len2)
    
    Dim Index0
    
    Dim MinIndex0
    Dim MaxIndex0
    
    Dim TempIndex1
    Dim TempIndex2
    
    Dim TempLCS
    
    Dim Cost
    For Cost = 0 To MaxCost
        MinIndex0 = -Cost
        MaxIndex0 = Cost
        
        If Cost > Len1 Then MinIndex0 = -Len1 + (Cost - Len1)
        If Cost > Len2 Then MaxIndex0 = Len2 - (Cost - Len2)
        
        For Index0 = MinIndex0 To MaxIndex0 Step 2
            If Cost = 0 Then
                TempIndex2 = 0
                TempLCS = ""
                
            ElseIf Index0 = MinIndex0 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempLCS = LCS(Len1 + Index0 + 1)
                
            ElseIf Index0 = MaxIndex0 Then
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
            
            TempIndex1 = TempIndex2 - Index0
            Do While TempIndex1 < Len1 And TempIndex2 < Len2
                If Str1(LB1 + TempIndex1) = Str2(LB2 + TempIndex2) Then
                    TempLCS = TempLCS & Str1(LB1 + TempIndex1)
                    
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
            Index2(Len1 + Index0) = TempIndex2
            LCS(Len1 + Index0) = TempLCS
        Next
    Next
End Function

Function ShortestEditScript(Str1(), Str2())
    Dim LB1
    Dim UB1
    Dim Len1
    If Not IsError(Str1) Then
        LB1 = LBound(Str1)
        UB1 = UBound(Str1)
        Len1 = UB1 - LB1 + 1
    End If
    
    Dim LB2
    Dim UB2
    Dim Len2
    If Not IsError(Str2) Then
        LB2 = LBound(Str2)
        UB2 = UBound(Str2)
        Len2 = UB2 - LB2 + 1
    End If
    
    If (Len1 = 0) And (Len2 = 0) Then
        ShortestEditScript = ""
        Exit Function
        
    ElseIf Len2 = 0 Then
        Dim SCSTemp1
        Dim Index1Temp
        For Index1Temp = 1 To Len1
            SCSTemp1 = SCSTemp1 & "-"
        Next
        ShortestEditScript = SCSTemp1
        Exit Function
        
    ElseIf Len1 = 0 Then
        Dim SCSTemp2
        Dim Index2Temp
        For Index2Temp = 1 To Len2
            SCSTemp2 = SCSTemp2 & "+"
        Next
        ShortestEditScript = SCSTemp2
        Exit Function
        
    End If
    
    Dim MaxCost
    MaxCost = Len1 + Len2
    
    Dim Index2()
    ReDim Index2(Len1 + Len2)
    
    Dim SES()
    ReDim SES(Len1 + Len2)
    
    Dim Index0
    
    Dim MinIndex0
    Dim MaxIndex0
    
    Dim TempIndex1
    Dim TempIndex2
    
    Dim TempSES
    
    Dim Cost
    For Cost = 0 To MaxCost
        MinIndex0 = -Cost
        MaxIndex0 = Cost
        
        If Cost > Len1 Then MinIndex0 = -Len1 + (Cost - Len1)
        If Cost > Len2 Then MaxIndex0 = Len2 - (Cost - Len2)
        
        For Index0 = MinIndex0 To MaxIndex0 Step 2
            If Cost = 0 Then
                TempIndex2 = 0
                TempSES = ""
                
            ElseIf Index0 = MinIndex0 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempSES = SES(Len1 + Index0 + 1) & "-"
                
            ElseIf Index0 = MaxIndex0 Then
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempSES = SES(Len1 + Index0 - 1) & "+"
                
            ElseIf Index2(Len1 + Index0 + 1) > _
                Index2(Len1 + Index0 - 1) + 1 Then
                TempIndex2 = Index2(Len1 + Index0 + 1)
                TempSES = SES(Len1 + Index0 + 1) & "-"
                
            Else
                TempIndex2 = Index2(Len1 + Index0 - 1) + 1
                TempSES = SES(Len1 + Index0 - 1) & "+"
                
            End If
            
            TempIndex1 = TempIndex2 - Index0
            Do While TempIndex1 < Len1 And TempIndex2 < Len2
                If Str1(LB1 + TempIndex1) = Str2(LB2 + TempIndex2) Then
                    TempSES = TempSES & " "
                    
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
            Index2(Len1 + Index0) = TempIndex2
            SES(Len1 + Index0) = TempSES
        Next
    Next
End Function

Private Function LongMax2(Lng1, Lng2)
    LongMax2 = IIf(Lng1 > Lng2, Lng1, Lng2)
End Function

Private Function IsError(Str())
On Error Resume Next
    Dim Len_Str
    Len_Str = UBound(Str) - LBound(Str) + 1
    IsError = (Len_Str <= 0)
End Function

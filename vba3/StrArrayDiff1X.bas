Attribute VB_Name = "StrArrayDiff1X"
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
' === String Array Difference 1 - Simple Implementation ===
'

'
' Modified - Cost: 2 dimension to 1 dimension
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
    
    If Len1 > Len2 Then
        EditDistance = EditDistanceCore(Str1, Str2)
    Else
        EditDistance = EditDistanceCore(Str2, Str1)
    End If
End Function

Private Function EditDistanceCore(Str1(), Str2())
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
        EditDistanceCore = Len2
        Exit Function
    End If
    If Len2 = 0 Then
        EditDistanceCore = Len1
        Exit Function
    End If
    
    Dim Cost()
    ReDim Cost(Len2)
    
    Dim Index1
    Dim Index2
    
    For Index2 = 0 To Len2
        Cost(Index2) = Index2
    Next
    
    Dim DiagonalCost
    Dim TempDiagonalCost
    
    For Index1 = 1 To Len1
        Cost(0) = Index1
        DiagonalCost = Index1 - 1
        For Index2 = 1 To Len2
            TempDiagonalCost = Cost(Index2)
            If Str1(LB1 + Index1 - 1) = Str2(LB2 + Index2 - 1) Then
                Cost(Index2) = _
                    LongMin3(Cost(Index2) + 1, Cost(Index2 - 1) + 1, _
                    DiagonalCost)
            Else
                Cost(Index2) = _
                    LongMin2(Cost(Index2) + 1, Cost(Index2 - 1) + 1)
            End If
            DiagonalCost = TempDiagonalCost
        Next
    Next
    
    EditDistanceCore = Cost(Len2)
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
    
    If Len1 > Len2 Then
        LongestCommonSubsequence = LongestCommonSubsequenceCore(Str1, Str2)
    Else
        LongestCommonSubsequence = LongestCommonSubsequenceCore(Str2, Str1)
    End If
End Function

Private Function LongestCommonSubsequenceCore( _
    Str1(), Str2())
    
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
        LongestCommonSubsequenceCore = ""
        Exit Function
    End If
    If Len2 = 0 Then
        LongestCommonSubsequenceCore = ""
        Exit Function
    End If
    
    Dim Cost()
    ReDim Cost(Len2)
    
    Dim LCS()
    ReDim LCS(Len2)
    
    Dim Index1
    Dim Index2
    
    For Index2 = 0 To Len2
        Cost(Index2) = Index2
    Next
    
    Dim DiagonalCost
    Dim TempDiagonalCost
    
    Dim TempCost1
    Dim TempCost2
    Dim TempCost3
    
    Dim DiagonalLCS
    Dim TempDiagonalLCS
    
    Dim TempLCS1
    Dim TempLCS2
    Dim TempLCS3
    
    For Index1 = 1 To Len1
        Cost(0) = Index1
        DiagonalCost = Index1 - 1
        DiagonalLCS = ""
        
        For Index2 = 1 To Len2
            TempDiagonalCost = Cost(Index2)
            
            TempCost1 = Cost(Index2) + 1
            TempCost2 = Cost(Index2 - 1) + 1
            
            TempDiagonalLCS = LCS(Index2)
            
            TempLCS1 = LCS(Index2)
            TempLCS2 = LCS(Index2 - 1)
            
        If Str1(LB1 + Index1 - 1) = Str2(LB2 + Index2 - 1) Then
                TempCost3 = DiagonalCost
                
                TempLCS3 = DiagonalLCS & Str1(LB1 + Index1 - 1)
                
                If TempCost1 < TempCost2 Then
                    If TempCost1 < TempCost3 Then
                        Cost(Index2) = TempCost1
                        LCS(Index2) = TempLCS1
                    Else
                        Cost(Index2) = TempCost3
                        LCS(Index2) = TempLCS3
                    End If
                Else
                    If TempCost2 < TempCost3 Then
                        Cost(Index2) = TempCost2
                        LCS(Index2) = TempLCS2
                    Else
                        Cost(Index2) = TempCost3
                        LCS(Index2) = TempLCS3
                    End If
                End If
                
            Else
                If TempCost1 < TempCost2 Then
                    Cost(Index2) = TempCost1
                    LCS(Index2) = TempLCS1
                Else
                    Cost(Index2) = TempCost2
                    LCS(Index2) = TempLCS2
                End If
                
            End If
            
            DiagonalCost = TempDiagonalCost
            DiagonalLCS = TempDiagonalLCS
        Next
    Next
    
    LongestCommonSubsequenceCore = LCS(Len2)
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
    
    If Len1 > Len2 Then
        ShortestEditScript = ShortestEditScriptCore(Str1, Str2, "-", "+")
    Else
        ShortestEditScript = ShortestEditScriptCore(Str2, Str1, "+", "-")
    End If
End Function

Private Function ShortestEditScriptCore( _
    Str1(), Str2(), _
    EditChar1, EditChar2)
    
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
        ShortestEditScriptCore = ""
        Exit Function
        
    ElseIf Len2 = 0 Then
        Dim SCSTemp1
        Dim Index1Temp
        For Index1Temp = 1 To Len1
            SCSTemp1 = SCSTemp1 & EditChar1
        Next
        ShortestEditScriptCore = SCSTemp1
        Exit Function
        
    ElseIf Len1 = 0 Then
        Dim SCSTemp2
        Dim Index2Temp
        For Index2Temp = 1 To Len2
            SCSTemp2 = SCSTemp2 & EditChar2
        Next
        ShortestEditScriptCore = SCSTemp2
        Exit Function
        
    End If
    
    Dim Cost()
    ReDim Cost(Len2)
    
    Dim SES()
    ReDim SES(Len2)
    
    Dim Index1
    Dim Index2
    
    Cost(0) = 0
    SES(0) = ""
    For Index2 = 1 To Len2
        Cost(Index2) = Index2
        SES(Index2) = SES(Index2 - 1) & EditChar2
    Next
    
    Dim DiagonalCost
    Dim TempDiagonalCost
    
    Dim TempCost1
    Dim TempCost2
    Dim TempCost3
    
    Dim DiagonalSES
    Dim TempDiagonalSES
    
    Dim TempSES1
    Dim TempSES2
    Dim TempSES3
    
    For Index1 = 1 To Len1
        Cost(0) = Index1
        DiagonalCost = Index1 - 1
        
        SES(0) = ""
        For Index1Temp = 1 To Index1
            SES(0) = SES(0) & EditChar1
        Next
        DiagonalSES = ""
        For Index1Temp = 1 To Index1 - 1
            DiagonalSES = DiagonalSES & EditChar1
        Next
        
        For Index2 = 1 To Len2
            TempDiagonalCost = Cost(Index2)
            
            TempCost1 = Cost(Index2) + 1
            TempCost2 = Cost(Index2 - 1) + 1
            
            TempDiagonalSES = SES(Index2)
            
            TempSES1 = SES(Index2) & EditChar1
            TempSES2 = SES(Index2 - 1) & EditChar2
            
            If Str1(LB1 + Index1 - 1) = Str2(LB2 + Index2 - 1) Then
                TempCost3 = DiagonalCost
                
                TempSES3 = DiagonalSES & " "
                
                If TempCost1 < TempCost2 Then
                    If TempCost1 < TempCost3 Then
                        Cost(Index2) = TempCost1
                        SES(Index2) = TempSES1
                    Else
                        Cost(Index2) = TempCost3
                        SES(Index2) = TempSES3
                    End If
                Else
                    If TempCost2 < TempCost3 Then
                        Cost(Index2) = TempCost2
                        SES(Index2) = TempSES2
                    Else
                        Cost(Index2) = TempCost3
                        SES(Index2) = TempSES3
                    End If
                End If
                
            Else
                If TempCost1 < TempCost2 Then
                    Cost(Index2) = TempCost1
                    SES(Index2) = TempSES1
                Else
                    Cost(Index2) = TempCost2
                    SES(Index2) = TempSES2
                End If
                
            End If
            
            DiagonalCost = TempDiagonalCost
            DiagonalSES = TempDiagonalSES
        Next
    Next
    
    ShortestEditScriptCore = SES(Len2)
End Function

Private Function LongMin3(Lng1, Lng2, Lng3)
    LongMin3 = LongMin2(LongMin2(Lng1, Lng2), Lng3)
End Function

Private Function LongMin2(Lng1, Lng2)
    LongMin2 = IIf(Lng1 < Lng2, Lng1, Lng2)
End Function

Private Function IsError(Str())
On Error Resume Next
    Dim Len_Str
    Len_Str = UBound(Str) - LBound(Str) + 1
    IsError = (Len_Str <= 0)
End Function

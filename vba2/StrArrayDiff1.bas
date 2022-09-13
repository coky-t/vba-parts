Attribute VB_Name = "StrArrayDiff1"
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

Function EditDistance(Str1() As String, Str2() As String) As Long
    Dim LB1 As Long
    Dim UB1 As Long
    Dim Len1 As Long
    If Not IsError(Str1) Then
        LB1 = LBound(Str1)
        UB1 = UBound(Str1)
        Len1 = UB1 - LB1 + 1
    End If
    
    Dim LB2 As Long
    Dim UB2 As Long
    Dim Len2 As Long
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
    
    Dim Cost() As Long
    ReDim Cost(Len1, Len2)
    
    Dim Index1 As Long
    Dim Index2 As Long
    
    For Index1 = 0 To Len1
        Cost(Index1, 0) = Index1
    Next
    For Index2 = 0 To Len2
        Cost(0, Index2) = Index2
    Next
    
    For Index1 = 1 To Len1
    For Index2 = 1 To Len2
        If Str1(LB1 + Index1 - 1) = Str2(LB2 + Index2 - 1) Then
            Cost(Index1, Index2) = _
                LongMin3( _
                    Cost(Index1 - 1, Index2) + 1, _
                    Cost(Index1, Index2 - 1) + 1, _
                    Cost(Index1 - 1, Index2 - 1))
        Else
            Cost(Index1, Index2) = _
                LongMin2( _
                    Cost(Index1 - 1, Index2) + 1, _
                    Cost(Index1, Index2 - 1) + 1)
        End If
    Next
    Next
    
    EditDistance = Cost(Len1, Len2)
End Function

Function LongestCommonSubsequence(Str1() As String, Str2() As String) As String
    Dim LB1 As Long
    Dim UB1 As Long
    Dim Len1 As Long
    If Not IsError(Str1) Then
        LB1 = LBound(Str1)
        UB1 = UBound(Str1)
        Len1 = UB1 - LB1 + 1
    End If
    
    Dim LB2 As Long
    Dim UB2 As Long
    Dim Len2 As Long
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
    
    Dim Cost() As Long
    ReDim Cost(Len1, Len2)
    
    Dim LCS() As String
    ReDim LCS(Len1, Len2)
    
    Dim Index1 As Long
    Dim Index2 As Long
    
    For Index1 = 0 To Len1
        Cost(Index1, 0) = Index1
    Next
    For Index2 = 0 To Len2
        Cost(0, Index2) = Index2
    Next
    
    Dim TempCost1 As Long
    Dim TempCost2 As Long
    Dim TempCost3 As Long
    
    Dim TempLCS1 As String
    Dim TempLCS2 As String
    Dim TempLCS3 As String
    
    For Index1 = 1 To Len1
    For Index2 = 1 To Len2
        TempCost1 = Cost(Index1 - 1, Index2) + 1
        TempCost2 = Cost(Index1, Index2 - 1) + 1
        
        TempLCS1 = LCS(Index1 - 1, Index2)
        TempLCS2 = LCS(Index1, Index2 - 1)
        
        If Str1(LB1 + Index1 - 1) = Str2(LB2 + Index2 - 1) Then
            TempCost3 = Cost(Index1 - 1, Index2 - 1)
            
            TempLCS3 = LCS(Index1 - 1, Index2 - 1) & Str1(LB1 + Index1 - 1)
            
            If TempCost1 < TempCost2 Then
                If TempCost1 < TempCost3 Then
                    Cost(Index1, Index2) = TempCost1
                    LCS(Index1, Index2) = TempLCS1
                Else
                    Cost(Index1, Index2) = TempCost3
                    LCS(Index1, Index2) = TempLCS3
                End If
            Else
                If TempCost2 < TempCost3 Then
                    Cost(Index1, Index2) = TempCost2
                    LCS(Index1, Index2) = TempLCS2
                Else
                    Cost(Index1, Index2) = TempCost3
                    LCS(Index1, Index2) = TempLCS3
                End If
            End If
            
        Else
            If TempCost1 < TempCost2 Then
                Cost(Index1, Index2) = TempCost1
                LCS(Index1, Index2) = TempLCS1
            Else
                Cost(Index1, Index2) = TempCost2
                LCS(Index1, Index2) = TempLCS2
            End If
            
        End If
    Next
    Next
    
    LongestCommonSubsequence = LCS(Len1, Len2)
End Function

Function ShortestEditScript(Str1() As String, Str2() As String) As String
    Dim LB1 As Long
    Dim UB1 As Long
    Dim Len1 As Long
    If Not IsError(Str1) Then
        LB1 = LBound(Str1)
        UB1 = UBound(Str1)
        Len1 = UB1 - LB1 + 1
    End If
    
    Dim LB2 As Long
    Dim UB2 As Long
    Dim Len2 As Long
    If Not IsError(Str2) Then
        LB2 = LBound(Str2)
        UB2 = UBound(Str2)
        Len2 = UB2 - LB2 + 1
    End If
    
    If (Len1 = 0) And (Len2 = 0) Then
        ShortestEditScript = ""
        Exit Function
        
    ElseIf Len2 = 0 Then
        Dim SCSTemp1 As String
        Dim Index1Temp As Long
        For Index1Temp = 1 To Len1
            SCSTemp1 = SCSTemp1 & "-"
        Next
        ShortestEditScript = SCSTemp1
        Exit Function
        
    ElseIf Len1 = 0 Then
        Dim SCSTemp2 As String
        Dim Index2Temp As Long
        For Index2Temp = 1 To Len2
            SCSTemp2 = SCSTemp2 & "+"
        Next
        ShortestEditScript = SCSTemp2
        Exit Function
        
    End If
    
    Dim Cost() As Long
    ReDim Cost(Len1, Len2)
    
    Dim SES() As String
    ReDim SES(Len1, Len2)
    
    Dim Index1 As Long
    Dim Index2 As Long
    
    Cost(0, 0) = 0
    SES(0, 0) = ""
    For Index1 = 1 To Len1
        Cost(Index1, 0) = Index1
        SES(Index1, 0) = SES(Index1 - 1, 0) & "-"
    Next
    For Index2 = 1 To Len2
        Cost(0, Index2) = Index2
        SES(0, Index2) = SES(0, Index2 - 1) & "+"
    Next
    
    Dim TempCost1 As Long
    Dim TempCost2 As Long
    Dim TempCost3 As Long
    
    Dim TempSES1 As String
    Dim TempSES2 As String
    Dim TempSES3 As String
    
    For Index1 = 1 To Len1
    For Index2 = 1 To Len2
        TempCost1 = Cost(Index1 - 1, Index2) + 1
        TempCost2 = Cost(Index1, Index2 - 1) + 1
        
        TempSES1 = SES(Index1 - 1, Index2) & "-"
        TempSES2 = SES(Index1, Index2 - 1) & "+"
        
        If Str1(LB1 + Index1 - 1) = Str2(LB2 + Index2 - 1) Then
            TempCost3 = Cost(Index1 - 1, Index2 - 1)
            
            TempSES3 = SES(Index1 - 1, Index2 - 1) & " "
            
            If TempCost1 < TempCost2 Then
                If TempCost1 < TempCost3 Then
                    Cost(Index1, Index2) = TempCost1
                    SES(Index1, Index2) = TempSES1
                Else
                    Cost(Index1, Index2) = TempCost3
                    SES(Index1, Index2) = TempSES3
                End If
            Else
                If TempCost2 < TempCost3 Then
                    Cost(Index1, Index2) = TempCost2
                    SES(Index1, Index2) = TempSES2
                Else
                    Cost(Index1, Index2) = TempCost3
                    SES(Index1, Index2) = TempSES3
                End If
            End If
            
        Else
            If TempCost1 < TempCost2 Then
                Cost(Index1, Index2) = TempCost1
                SES(Index1, Index2) = TempSES1
            Else
                Cost(Index1, Index2) = TempCost2
                SES(Index1, Index2) = TempSES2
            End If
            
        End If
    Next
    Next
    
    ShortestEditScript = SES(Len1, Len2)
End Function

Private Function LongMin3(Lng1 As Long, Lng2 As Long, Lng3 As Long) As Long
    LongMin3 = LongMin2(LongMin2(Lng1, Lng2), Lng3)
End Function

Private Function LongMin2(Lng1 As Long, Lng2 As Long) As Long
    LongMin2 = IIf(Lng1 < Lng2, Lng1, Lng2)
End Function

Private Function IsError(Str() As String) As Boolean
On Error Resume Next
    Dim Len_Str As Long
    Len_Str = UBound(Str) - LBound(Str) + 1
    IsError = (Len_Str <= 0)
End Function

Attribute VB_Name = "減免判定2号"
Sub 減免判定2号()
    Dim monthlyPastIncome As Long '前年中の世帯平均所得'
    Dim monthlyCurrentIncome As Long '今年の世帯平均所得（見込）'
    Dim decreaseRate As Variant '減少率'
   
    Dim outputTerm2 As Range '２号減免期間出力先'
    Dim applyMonth2 As Long '収入減少開始月'
    Dim p As Long '過年度所得計算(繰返処理用)'
    Dim c As Long '現年度所得計算(繰返処理用)'
    Dim additionP As Long '前年度旧ただし書き所得カウンター'
    Dim additionC As Long '今年度旧ただし書き所得カウンター'
    
    
    '減少率の算出はここから↓'
    
    monthlyPastIncome = 0
    monthlyCurrentIncome = 0
    
    '過年度旧ただし書き所得の算出'
    For p = 32 To Range("I32").End(xlDown).Row
        additionP = Application.WorksheetFunction.Sum(Range(Cells(p, 11), Cells(p, 13)))
        If additionP >= 0 Then
            monthlyPastIncome = monthlyPastIncome + additionP
        ElseIf additionP < 0 Then
            monthlyPastIncome = monthlyPastIncome + 0
        End If
    Next p
    
    
    '現年度旧ただし書き所得の算出'
    For c = 44 To Range("I44").End(xlDown).Row
        additionC = Application.WorksheetFunction.Sum(Range(Cells(c, 11), Cells(c, 13)))
        If additionC >= 0 Then
            monthlyCurrentIncome = monthlyCurrentIncome + additionC
        ElseIf additionC < 0 Then
            monthlyCurrentIncome = monthlyCurrentIncome + 0
        End If
    Next c
    
    On Error Resume Next
    decreaseRate = 1 - Application.WorksheetFunction.Round(monthlyCurrentIncome / monthlyPastIncome, 5)
    
    Set outputDecreaseRate = Range("J55")
    Set outputExemptionRate = Range("N55")
    Set outputTerm2 = Range("L55")
    outputDecreaseRate.Value = decreaseRate
    applyMonth2 = Range("C13").Value
    
    'Debug.Print (monthlyPastIncome)
    'Debug.Print (monthlyCurrentIncome)
    'Debug.Print (decreaseRate) '
    
    If decreaseRate >= 0.3 And decreaseRate < 0.4 Then
        outputExemptionRate.Value = "30%"
    
    ElseIf decreaseRate >= 0.4 And decreaseRate < 0.5 Then
        outputExemptionRate.Value = "40%"
    
    ElseIf decreaseRate >= 0.5 And decreaseRate < 0.6 Then
        outputExemptionRate.Value = "50%"
    
    ElseIf decreaseRate >= 0.6 And decreaseRate < 0.7 Then
        outputExemptionRate.Value = "60%"
    
    ElseIf decreaseRate >= 0.7 And decreaseRate < 0.8 Then
        outputExemptionRate.Value = "70%"
    
    ElseIf decreaseRate >= 0.8 And decreaseRate < 0.9 Then
        outputExemptionRate.Value = "80%"
    
    ElseIf decreaseRate >= 0.9 And decreaseRate < 1 Then
        outputExemptionRate.Value = "90%"
    
    ElseIf decreaseRate >= 1 Then
        outputExemptionRate.Value = "100%"
    
    Else
        outputExemptionRate.Value = "減免不可"
    
    End If
    
    '減免適用期間の出力処理'
    If applyMonth = 6 Then
    
        Select Case applyMonth2
            Case Is = 4
                outputTerm2.Value = "４〜"
            Case Is = 5
                outputTerm2.Value = "５〜"
            Case Is = 6
                outputTerm2.Value = "６〜"
            Case Is = 7
                outputTerm2.Value = "７〜"
            Case Is = 8
                outputTerm2.Value = "８〜"
            Case Is = 9
                outputTerm2.Value = "９〜"
        Case Else
            outputTerm2.Value = "減免対象期間外です。"
            MsgBox "所得減少月が選択されていない可能性があります。"
        End Select
        
        ElseIf applyMonth = 7 Then
            outputTerm2.Value = "７〜"
        
        ElseIf applyMonth = 8 Then
            outputTerm2.Value = "８〜"
        
        ElseIf applyMonth = 9 Then
            outputTerm2.Value = "９〜"
        
        ElseIf applyMonth = 10 Then
            outputTerm2.Value = "10〜"
        
        ElseIf applyMonth = 11 Then
            outputTerm2.Value = "11〜"
        
        ElseIf applyMonth = 12 Then
            outputTerm2.Value = "12〜"
        
        ElseIf applyMonth = 1 Then
            outputTerm2.Value = "１〜"
        
        ElseIf applyMonth = 2 Then
            outputTerm2.Value = "２〜"
        
        ElseIf applyMonth = 3 Then
            outputTerm2.Value = "３〜"
        
        ElseIf applyMonth = 4 Then
            outputTerm2.Value = "３〜"
    
    Else
        outputTerm2.Value = "減免対象期間外です。"
        MsgBox "所得減少月が選択されていない可能性があります。"
        
    End If
    
End Sub


Attribute VB_Name = "Œ¸–Æ”»’è2†"
Sub Œ¸–Æ”»’è2†()
    Dim monthlyPastIncome As Long '‘O”N’†‚Ì¢‘Ñ•½‹ÏŠ“¾'
    Dim monthlyCurrentIncome As Long '¡”N‚Ì¢‘Ñ•½‹ÏŠ“¾iŒ©j'
    Dim decreaseRate As Variant 'Œ¸­—¦'
   
    Dim outputTerm2 As Range '‚Q†Œ¸–ÆŠúŠÔo—Íæ'
    Dim applyMonth2 As Long 'û“üŒ¸­ŠJnŒ'
    Dim p As Long '‰ß”N“xŠ“¾ŒvZ(ŒJ•Ôˆ——p)'
    Dim c As Long 'Œ»”N“xŠ“¾ŒvZ(ŒJ•Ôˆ——p)'
    Dim additionP As Long '‘O”N“x‹Œ‚½‚¾‚µ‘‚«Š“¾ƒJƒEƒ“ƒ^['
    Dim additionC As Long '¡”N“x‹Œ‚½‚¾‚µ‘‚«Š“¾ƒJƒEƒ“ƒ^['
    
    
    'Œ¸­—¦‚ÌZo‚Í‚±‚±‚©‚ç«'
    
    monthlyPastIncome = 0
    monthlyCurrentIncome = 0
    
    '‰ß”N“x‹Œ‚½‚¾‚µ‘‚«Š“¾‚ÌZo'
    For p = 32 To Range("I32").End(xlDown).Row
        additionP = Application.WorksheetFunction.Sum(Range(Cells(p, 11), Cells(p, 13)))
        If additionP >= 0 Then
            monthlyPastIncome = monthlyPastIncome + additionP
        ElseIf additionP < 0 Then
            monthlyPastIncome = monthlyPastIncome + 0
        End If
    Next p
    
    
    'Œ»”N“x‹Œ‚½‚¾‚µ‘‚«Š“¾‚ÌZo'
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
        outputExemptionRate.Value = "Œ¸–Æ•s‰Â"
    
    End If
    
    'Œ¸–Æ“K—pŠúŠÔ‚Ìo—Íˆ—'
    If applyMonth = 6 Then
    
        Select Case applyMonth2
            Case Is = 4
                outputTerm2.Value = "‚S`"
            Case Is = 5
                outputTerm2.Value = "‚T`"
            Case Is = 6
                outputTerm2.Value = "‚U`"
            Case Is = 7
                outputTerm2.Value = "‚V`"
            Case Is = 8
                outputTerm2.Value = "‚W`"
            Case Is = 9
                outputTerm2.Value = "‚X`"
        Case Else
            outputTerm2.Value = "Œ¸–Æ‘ÎÛŠúŠÔŠO‚Å‚·B"
            MsgBox "Š“¾Œ¸­Œ‚ª‘I‘ğ‚³‚ê‚Ä‚¢‚È‚¢‰Â”\«‚ª‚ ‚è‚Ü‚·B"
        End Select
        
        ElseIf applyMonth = 7 Then
            outputTerm2.Value = "‚V`"
        
        ElseIf applyMonth = 8 Then
            outputTerm2.Value = "‚W`"
        
        ElseIf applyMonth = 9 Then
            outputTerm2.Value = "‚X`"
        
        ElseIf applyMonth = 10 Then
            outputTerm2.Value = "10`"
        
        ElseIf applyMonth = 11 Then
            outputTerm2.Value = "11`"
        
        ElseIf applyMonth = 12 Then
            outputTerm2.Value = "12`"
        
        ElseIf applyMonth = 1 Then
            outputTerm2.Value = "‚P`"
        
        ElseIf applyMonth = 2 Then
            outputTerm2.Value = "‚Q`"
        
        ElseIf applyMonth = 3 Then
            outputTerm2.Value = "‚R`"
        
        ElseIf applyMonth = 4 Then
            outputTerm2.Value = "‚R`"
    
    Else
        outputTerm2.Value = "Œ¸–Æ‘ÎÛŠúŠÔŠO‚Å‚·B"
        MsgBox "Š“¾Œ¸­Œ‚ª‘I‘ğ‚³‚ê‚Ä‚¢‚È‚¢‰Â”\«‚ª‚ ‚è‚Ü‚·B"
        
    End If
    
End Sub


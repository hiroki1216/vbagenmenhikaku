Attribute VB_Name = "‹‹—^Š“¾Zo"
Sub ‹‹—^Š“¾Zo()
    
'‹‹—^Š“¾Zo‚Í‚±‚±‚©‚ç'
    Select Case annualSumS
    
        Case Is < 550999
            'debug.Print (0)'
            annualIncomeS = 0
            monthlyIncomeSInput.Value = annualIncomeS
        
        Case 551000 To 1618999
            'Debug.Print (annualSumS - 550000) '
            annualIncomeS = annualSumS - 550000
        
        Case 1619000 To 1619999
            'Debug.Print (1069000)'
            annualIncomeS = 1069000
        
        Case 1620000 To 1621999
            'Debug.Print (1070000) '
            annualIncomeS = 1070000
            
        Case 1622000 To 1623999
            'Debug.Print (1072000) '
            annualIncomeS = 1072000
        
        Case 1624000 To 1627999
            'Debug.Print (1074000)'
            annualIncomeS = 1074000
        
        Case 1628000 To 1799999
            'Debug.Print (annualSumS * 0.6 + 100000) '
            annualIncomeS = annualSumS * 0.6 + 100000
        
        Case 1800000 To 3599999
            'Debug.Print (annualSumS * 0.7 - 80000) '
            annualIncomeS = annualSumS * 0.7 - 80000
        
        Case 3600000 To 6599999
            'Debug.Print (annualSumS * 0.8 - 440000) '
            annualIncomeS = annualSumS * 0.8 - 440000
        
        Case 6600000 To 8499999
            'Debug.Print (annualSumS * 0.9 - 1100000)'
            annualIncomeS = annualSumS * 0.9 - 1100000
        
        Case Is > 8500000
            'Debug.Print (annualSumS - 1950000)'
            annualIncomeS = annualSumS - 1950000
        
        Case Else
            'Debug.Print ("–³Œø‚È’l‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚·B")'
            annualIncomeS = 0
            monthlyIncomeSInput.Value = annualIncomeS
            MsgBox "‹‹—^Š“¾‚ÌŒvZ‚É¸”s‚µ‚Ü‚µ‚½B"
    End Select
End Sub

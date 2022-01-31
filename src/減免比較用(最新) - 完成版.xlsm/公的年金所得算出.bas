Attribute VB_Name = "Œö“I”N‹àŠ“¾Zo"
Sub Œö“I”N‹àŠ“¾Zo()
    '‡@65Î–¢–”N‹àû“ü130–œ‰~ˆÈ‰º'
     
     '‡@-1‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=1000–œ‰~'
     If age < 65 And annualSumP <= 1300000 And annualTotalIncome <= 10000000 Then
     'Debug.Print (600000)'
     
         If annualSumP - 600000 > 0 Then
             annualIncomeP = (annualSumP - 600000)
         Else
             annualIncomeP = 0
         End If
     
     '‡@-2 1000–œ‰~<‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=2000–œ‰~'
     ElseIf age < 65 And annualSumP <= 1300000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (500000)'
         
         If annualSumP - 500000 > 0 Then
             annualIncomeP = (annualSumP - 500000)
         Else
             annualIncomeP = 0
         End If
     
     '‡@-3 ‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)>2000–œ‰~'
     ElseIf age < 65 And annualSumP <= 1300000 And annualTotalIncome > 20000000 Then
         'Debug.Print (400000) '
         
         If annualSumP - 400000 > 0 Then
             annualIncomeP = (annualSumP - 400000)
         Else
             annualIncomeP = 0
         End If
     
     
     '‡A65Î–¢– ”N‹àû“ü130–œ‰~’´410–œ‰~ˆÈ‰º'
     
     '‡A-1‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=1000–œ‰~'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.25 + 275000) '
         pDeduction = annualSumP * 0.25 + 275000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡A-2 1000–œ‰~<‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=2000–œ‰~'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 175000) '
         pDeduction = annualSumP * 0.25 + 175000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡A-3 ‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)>2000–œ‰~'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 75000)'
         pDeduction = annualSumP * 0.25 + 75000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡B‹¤’Ê ”N‹àû“ü410–œ‰~’´770–œ‰~ˆÈ‰º'
     
     '‡B-1‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=1000–œ‰~'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 685000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡B-2 1000–œ‰~<‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=2000–œ‰~'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 585000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡B-3 ‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)>2000–œ‰~'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 485000
         annualIncomeP = (annualSumP - pDeduction)
     
     
     '‡C‹¤’Ê ”N‹àû“ü770–œ‰~’´1000–œ‰~ˆÈ‰º'
     
     '‡C-1‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=1000–œ‰~'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.05 + 1455000)'
         pDeduction = annualSumP * 0.05 + 1455000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡C-2 1000–œ‰~<‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=2000–œ‰~'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.05 + 1355000)'
         pDeduction = annualSumP * 0.05 + 1355000
         annualIncomeP = (annualSumP - pDeduction)
         
     '‡C-3 ‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)>2000–œ‰~'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.05 + 1255000)'
         pDeduction = annualSumP * 0.05 + 1255000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡D‹¤’Ê ”N‹àû“ü1000–œ‰~’´'
     
     '‡D-1‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=1000–œ‰~'
     ElseIf annualSumP > 10000000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (1955000)'
         pDeduction = 1955000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡D-2 1000–œ‰~<‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=2000–œ‰~'
     ElseIf annualSumP > 10000000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (1855000)'
         pDeduction = 1855000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡D-3 ‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)>2000–œ‰~'
     ElseIf annualSumP > 10000000 And annualTotalIncome > 20000000 Then
         'Debug.Print (1755000) '
         pDeduction = 1755000
         annualIncomeP = (annualSumP - pDeduction)
     
     
     
     
     
     '‡@65ÎˆÈã”N‹àû“ü330–œ‰~ˆÈ‰º'
     
     '‡@-1‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=1000–œ‰~'
     ElseIf age >= 65 And annualSumP <= 3300000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (1100000)'
         pDeduction = 1100000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '‡@-2 1000–œ‰~<‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=2000–œ‰~'
     ElseIf age >= 65 And annualSumP <= 3300000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (1000000)'
         pDeduction = 1000000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '‡@-3 ‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)>2000–œ‰~'
     ElseIf age >= 65 And annualSumP <= 330000 And annualTotalIncome > 20000000 Then
         'Debug.Print (900000)'
         pDeduction = 900000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '‡A65Î–¢– ”N‹àû“ü330–œ‰~’´410–œ‰~ˆÈ‰º'
     
     '‡A-1‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=1000–œ‰~'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.25 + 275000) '
         pDeduction = annualSumP * 0.25 + 275000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡A-2 1000–œ‰~<‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)<=2000–œ‰~'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 175000)'
         pDeduction = annualSumP * 0.25 + 175000
         annualIncomeP = (annualSumP - pDeduction)
     
     '‡A-3 ‡ŒvŠ“¾‹àŠz(”N‹à‚ğœ‚­)>2000–œ‰~'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 75000) '
         pDeduction = annualSumP * 0.25 + 75000
         annualIncomeP = (annualSumP - pDeduction)
     
     Else
         'Debug.Print ("–¢İ’è‚Ì’l‚ª“ü—Í‚³‚ê‚Ü‚µ‚½")'
         annualIncomeP = -9999999
         MsgBox "”N‹àŠ“¾‚ÌŒvZ‚É¸”s‚µ‚Ü‚µ‚½B"
     
     End If
End Sub

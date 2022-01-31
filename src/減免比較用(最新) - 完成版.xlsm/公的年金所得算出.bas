Attribute VB_Name = "公的年金所得算出"
Sub 公的年金所得算出()
    '�@65歳未満年金収入130万円以下'
     
     '�@-1合計所得金額(年金を除く)<=1000万円'
     If age < 65 And annualSumP <= 1300000 And annualTotalIncome <= 10000000 Then
     'Debug.Print (600000)'
     
         If annualSumP - 600000 > 0 Then
             annualIncomeP = (annualSumP - 600000)
         Else
             annualIncomeP = 0
         End If
     
     '�@-2 1000万円<合計所得金額(年金を除く)<=2000万円'
     ElseIf age < 65 And annualSumP <= 1300000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (500000)'
         
         If annualSumP - 500000 > 0 Then
             annualIncomeP = (annualSumP - 500000)
         Else
             annualIncomeP = 0
         End If
     
     '�@-3 合計所得金額(年金を除く)>2000万円'
     ElseIf age < 65 And annualSumP <= 1300000 And annualTotalIncome > 20000000 Then
         'Debug.Print (400000) '
         
         If annualSumP - 400000 > 0 Then
             annualIncomeP = (annualSumP - 400000)
         Else
             annualIncomeP = 0
         End If
     
     
     '�A65歳未満 年金収入130万円超410万円以下'
     
     '�A-1合計所得金額(年金を除く)<=1000万円'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.25 + 275000) '
         pDeduction = annualSumP * 0.25 + 275000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-2 1000万円<合計所得金額(年金を除く)<=2000万円'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 175000) '
         pDeduction = annualSumP * 0.25 + 175000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-3 合計所得金額(年金を除く)>2000万円'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 75000)'
         pDeduction = annualSumP * 0.25 + 75000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�B共通 年金収入410万円超770万円以下'
     
     '�B-1合計所得金額(年金を除く)<=1000万円'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 685000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�B-2 1000万円<合計所得金額(年金を除く)<=2000万円'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 585000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�B-3 合計所得金額(年金を除く)>2000万円'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 485000
         annualIncomeP = (annualSumP - pDeduction)
     
     
     '�C共通 年金収入770万円超1000万円以下'
     
     '�C-1合計所得金額(年金を除く)<=1000万円'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.05 + 1455000)'
         pDeduction = annualSumP * 0.05 + 1455000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�C-2 1000万円<合計所得金額(年金を除く)<=2000万円'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.05 + 1355000)'
         pDeduction = annualSumP * 0.05 + 1355000
         annualIncomeP = (annualSumP - pDeduction)
         
     '�C-3 合計所得金額(年金を除く)>2000万円'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.05 + 1255000)'
         pDeduction = annualSumP * 0.05 + 1255000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�D共通 年金収入1000万円超'
     
     '�D-1合計所得金額(年金を除く)<=1000万円'
     ElseIf annualSumP > 10000000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (1955000)'
         pDeduction = 1955000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�D-2 1000万円<合計所得金額(年金を除く)<=2000万円'
     ElseIf annualSumP > 10000000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (1855000)'
         pDeduction = 1855000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�D-3 合計所得金額(年金を除く)>2000万円'
     ElseIf annualSumP > 10000000 And annualTotalIncome > 20000000 Then
         'Debug.Print (1755000) '
         pDeduction = 1755000
         annualIncomeP = (annualSumP - pDeduction)
     
     
     
     
     
     '�@65歳以上年金収入330万円以下'
     
     '�@-1合計所得金額(年金を除く)<=1000万円'
     ElseIf age >= 65 And annualSumP <= 3300000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (1100000)'
         pDeduction = 1100000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '�@-2 1000万円<合計所得金額(年金を除く)<=2000万円'
     ElseIf age >= 65 And annualSumP <= 3300000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (1000000)'
         pDeduction = 1000000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '�@-3 合計所得金額(年金を除く)>2000万円'
     ElseIf age >= 65 And annualSumP <= 330000 And annualTotalIncome > 20000000 Then
         'Debug.Print (900000)'
         pDeduction = 900000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '�A65歳未満 年金収入330万円超410万円以下'
     
     '�A-1合計所得金額(年金を除く)<=1000万円'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.25 + 275000) '
         pDeduction = annualSumP * 0.25 + 275000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-2 1000万円<合計所得金額(年金を除く)<=2000万円'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 175000)'
         pDeduction = annualSumP * 0.25 + 175000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-3 合計所得金額(年金を除く)>2000万円'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 75000) '
         pDeduction = annualSumP * 0.25 + 75000
         annualIncomeP = (annualSumP - pDeduction)
     
     Else
         'Debug.Print ("未設定の値が入力されました")'
         annualIncomeP = -9999999
         MsgBox "年金所得の計算に失敗しました。"
     
     End If
End Sub

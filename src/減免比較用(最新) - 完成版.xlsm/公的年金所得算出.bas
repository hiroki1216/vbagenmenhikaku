Attribute VB_Name = "���I�N�������Z�o"
Sub ���I�N�������Z�o()
    '�@65�Ζ����N������130���~�ȉ�'
     
     '�@-1���v�������z(�N��������)<=1000���~'
     If age < 65 And annualSumP <= 1300000 And annualTotalIncome <= 10000000 Then
     'Debug.Print (600000)'
     
         If annualSumP - 600000 > 0 Then
             annualIncomeP = (annualSumP - 600000)
         Else
             annualIncomeP = 0
         End If
     
     '�@-2 1000���~<���v�������z(�N��������)<=2000���~'
     ElseIf age < 65 And annualSumP <= 1300000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (500000)'
         
         If annualSumP - 500000 > 0 Then
             annualIncomeP = (annualSumP - 500000)
         Else
             annualIncomeP = 0
         End If
     
     '�@-3 ���v�������z(�N��������)>2000���~'
     ElseIf age < 65 And annualSumP <= 1300000 And annualTotalIncome > 20000000 Then
         'Debug.Print (400000) '
         
         If annualSumP - 400000 > 0 Then
             annualIncomeP = (annualSumP - 400000)
         Else
             annualIncomeP = 0
         End If
     
     
     '�A65�Ζ��� �N������130���~��410���~�ȉ�'
     
     '�A-1���v�������z(�N��������)<=1000���~'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.25 + 275000) '
         pDeduction = annualSumP * 0.25 + 275000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-2 1000���~<���v�������z(�N��������)<=2000���~'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 175000) '
         pDeduction = annualSumP * 0.25 + 175000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-3 ���v�������z(�N��������)>2000���~'
     ElseIf age < 65 And annualSumP > 1300000 And annualSumP <= 4100000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 75000)'
         pDeduction = annualSumP * 0.25 + 75000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�B���� �N������410���~��770���~�ȉ�'
     
     '�B-1���v�������z(�N��������)<=1000���~'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 685000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�B-2 1000���~<���v�������z(�N��������)<=2000���~'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 585000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�B-3 ���v�������z(�N��������)>2000���~'
     ElseIf annualSumP > 4100000 And annualSumP <= 7700000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.15 + 685000)'
         pDeduction = annualSumP * 0.15 + 485000
         annualIncomeP = (annualSumP - pDeduction)
     
     
     '�C���� �N������770���~��1000���~�ȉ�'
     
     '�C-1���v�������z(�N��������)<=1000���~'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.05 + 1455000)'
         pDeduction = annualSumP * 0.05 + 1455000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�C-2 1000���~<���v�������z(�N��������)<=2000���~'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.05 + 1355000)'
         pDeduction = annualSumP * 0.05 + 1355000
         annualIncomeP = (annualSumP - pDeduction)
         
     '�C-3 ���v�������z(�N��������)>2000���~'
     ElseIf annualSumP > 7700000 And annualSumP <= 10000000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.05 + 1255000)'
         pDeduction = annualSumP * 0.05 + 1255000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�D���� �N������1000���~��'
     
     '�D-1���v�������z(�N��������)<=1000���~'
     ElseIf annualSumP > 10000000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (1955000)'
         pDeduction = 1955000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�D-2 1000���~<���v�������z(�N��������)<=2000���~'
     ElseIf annualSumP > 10000000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (1855000)'
         pDeduction = 1855000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�D-3 ���v�������z(�N��������)>2000���~'
     ElseIf annualSumP > 10000000 And annualTotalIncome > 20000000 Then
         'Debug.Print (1755000) '
         pDeduction = 1755000
         annualIncomeP = (annualSumP - pDeduction)
     
     
     
     
     
     '�@65�Έȏ�N������330���~�ȉ�'
     
     '�@-1���v�������z(�N��������)<=1000���~'
     ElseIf age >= 65 And annualSumP <= 3300000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (1100000)'
         pDeduction = 1100000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '�@-2 1000���~<���v�������z(�N��������)<=2000���~'
     ElseIf age >= 65 And annualSumP <= 3300000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (1000000)'
         pDeduction = 1000000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '�@-3 ���v�������z(�N��������)>2000���~'
     ElseIf age >= 65 And annualSumP <= 330000 And annualTotalIncome > 20000000 Then
         'Debug.Print (900000)'
         pDeduction = 900000
         If annualSumP - pDeduction > 0 Then
             annualIncomeP = (annualSumP - pDeduction)
         Else
             annualIncomeP = 0
         End If
     
     '�A65�Ζ��� �N������330���~��410���~�ȉ�'
     
     '�A-1���v�������z(�N��������)<=1000���~'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome <= 10000000 Then
         'Debug.Print (annualSumP * 0.25 + 275000) '
         pDeduction = annualSumP * 0.25 + 275000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-2 1000���~<���v�������z(�N��������)<=2000���~'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome > 10000000 And annualTotalIncome <= 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 175000)'
         pDeduction = annualSumP * 0.25 + 175000
         annualIncomeP = (annualSumP - pDeduction)
     
     '�A-3 ���v�������z(�N��������)>2000���~'
     ElseIf age >= 65 And annualSumP > 3300000 And annualSumP <= 4100000 And annualTotalIncome > 20000000 Then
         'Debug.Print (annualSumP * 0.25 + 75000) '
         pDeduction = annualSumP * 0.25 + 75000
         annualIncomeP = (annualSumP - pDeduction)
     
     Else
         'Debug.Print ("���ݒ�̒l�����͂���܂���")'
         annualIncomeP = -9999999
         MsgBox "�N�������̌v�Z�Ɏ��s���܂����B"
     
     End If
End Sub

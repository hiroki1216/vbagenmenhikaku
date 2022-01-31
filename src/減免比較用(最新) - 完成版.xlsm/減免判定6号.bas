Attribute VB_Name = "���Ɣ���6��"
Sub ���Ɣ���6��()
    Set outputExePrice = Range("S54") '6�����Ɗz�o�̓Z���I�u�W�F�N�g�̎擾'
    Set outputExeRate = Range("S55") '6�����Ɨ��o�̓Z���I�u�W�F�N�g�̎擾'
    
    If Range("G3").Value <> 0 Then
        MsgBox "���ю傪�񎩔��I���Ǝ҂̂���6�����ƕs�ł��B"
        outputExeRate.Value = 0
        outputExePrice.Value = "���ƕs��(�傪�񎩔�)"
    
    Else
    
        Dim outputKoronakyuuhu As Long '�O�N���R���i�֌W���t���z'
        Dim outputKoronaPaymentPa As Long '�O�N���R���i�e������(�O�N���R���i�֌W���t���z�T����)'
        Dim outputKoronaPaymentPa2 As Long '�O�N���R���i�e������(�O�N���R���i�֌W���t���z�T���O)'
        Dim outputKoronaPaymentC As Long '���N���̃R���i�e�������ݎ���'
        
        Dim outputKoronaIncomeS As Long '�O�N���̃R���i�e������'
        Dim outputTotalIncome As Long '�O�N���̍��v����'
        Dim outputNonKoronaIncome As Long '�O�N���̃R���i�e���O����'
        Dim outputObjectPrice As Long '���ƑΏەی���'
        Dim annualCost As Variant '�Z��ی����z�f
        Dim familyTotalIncome As Long '�O�N���̐��э��v����'
        Dim influencePayment As String '�R���i�e��������'
        
        Dim tbl1 As Range 'v-lookup�����p�͈�(�O�N���̃R���i�e�������j'
        Dim tbl2 As Range 'v-lookup�����p�͈�(�O�N���̃R���i�e�������j'
        
        Dim key As Long 'v-lookup�����p�L�['
        Dim decreaseRateKorona As Variant '�R���i�e������������'
        Dim message As String '�G���[���b�Z�[�W�f
        
        outputKoronaPaymentC = Range("Q7").Value
        key = Range("B3").Value
        influencePayment = Range("Q2").Value
        annualCost = Range("P31").Value
        familyTotalIncome = Range("R31").Value
        
        
        '�e�������ŏ������򏈗��f
        
        If influencePayment = "���^����" Then
        
            Set tbl = Worksheets("�ėp���o�i�����E���Y�j�������(�ڍ�)").Range("C:AZ")
            
            '�O�N�̉e������(�R���i�֌W���t���T����)�̎擾'
            On Error Resume Next
            outputKoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 13, False) - Range("C23").Value
            If Err.Number <> 0 Then
                outputKoronaPaymentPa = -999999
                MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            '�O�N�̉e������(�R���i�֌W���t���T���O)�̎擾'
            On Error Resume Next
            outputKoronaPaymentPa2 = Application.WorksheetFunction.VLookup(key, tbl, 13, False)
            If Err.Number <> 0 Then
                outputKoronaPaymentPa2 = -999999
                MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            
            '�O�N�̉e�����^�����̎擾'
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 26, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            
            '�O�N�̍��v�����̎擾'
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "�O�N�̍��v�����̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            
            '�O�N�̉e���O�����̎擾'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            '�e�������̌�����'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '���ƑΏەی����z�̎Z�o'
            On Error Resume Next
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
            
            Dim KoronaPaymentPa As Long '�O�N�R���i�e������(���^�N�������ȊO)'
            Dim inputResult As Long '���͌���'
        
        ElseIf influencePayment = "�c�Ǝ���" Then
        
            Set tbl = Worksheets("�ėp���o�i�����E���Y�j�������(�ڍ�)").Range("C:AZ")
            
            KoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) '�O�N�R���i�e���c�Ǝ����̎擾'
            
            
            '�O�N���̎����̐\���̗L���ŏ������򂱂����火'
            
            If KoronaPaymentPa > 0 Then
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T����)'
                On Error Resume Next
                outputKoronaPaymentPa = KoronaPaymentPa - Range("C23").Value
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa = -999999
                    MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
                End If
                On Error GoTo 0
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T���O)'
                On Error Resume Next
                outputKoronaPaymentPa2 = KoronaPaymentPa
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa2 = -999999
                    MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
                End If
                On Error GoTo 0
            ElseIf KoronaPaymentPa = 0 Then
            
                inputResult = Application.InputBox("�O�N�c�Ǝ�����0�ł��B" & vbCrLf & "�O�N�c�Ǝ�������͂��ĉ������B")
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T����)'
                outputKoronaPaymentPa = inputResult - Range("C23").Value
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T���O)'
                outputKoronaPaymentPa2 = inputResult
            End If
            
            '�O�N���̎����̐\���̗L���ŏ������򂱂��܂Ł�'
            
            
            '�O�N�̉e�������̎擾'
            
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 28, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            '�O�N�̍��v�����̎擾'
            
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "�O�N�̍��v�����̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            '�O�N�̉e���O�����̎擾'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            '�e�������̌�����'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '���ƑΏەی����z�̎Z�o'
            On Error Resume Next
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
        
        ElseIf influencePayment = "���̑�����" Then
        
            Set tbl = Worksheets("�ėp���o�i�����E���Y�j�������(�ڍ�)").Range("C:AZ")
            
            KoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) '�O�N�R���i�e�����̑������̎擾'
            
            
            '�O�N���̎����̐\���̗L���ŏ������򂱂����火'
            
            If KoronaPaymentPa > 0 Then
                
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T����)'
                On Error Resume Next
                outputKoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) - Range("C23").Value
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa = -999999
                    MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
                End If
                On Error GoTo 0
                
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T���O)'
                On Error Resume Next
                outputKoronaPaymentPa2 = Application.WorksheetFunction.VLookup(key, tbl, 17, False)
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa2 = -999999
                    MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
                End If
                On Error GoTo 0
            
            ElseIf KoronaPaymentPa = 0 Then
            
                inputResult = Application.InputBox("�O�N���̑�������0�ł��B" & vbCrLf & "�O�N���̑���������͂��ĉ������B")
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T����)'
                outputKoronaPaymentPa = inputResult - Range("C23").Value
                '�O�N�̉e�������̎擾(�R���i�֌W���t���T���O)'
                outputKoronaPaymentPa2 = inputResult
            
            End If
            
            '�O�N���̎����̐\���̗L���ŏ������򂱂��܂Ł�'
            
            
            '�O�N�̉e�������̎擾'
            
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 36, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "�O�N�̉e�������̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            '�O�N�̍��v�����̎擾'
            
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "�O�N�̍��v�����̎擾�Ɏ��s���܂����B"
            End If
            On Error GoTo 0
            
            '�O�N�̉e���O�����̎擾'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            '�e�������̌�����'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '���ƑΏەی����z�̎Z�o'
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
        
        
        
        Else
            MsgBox "���ю�̃R���i�e���������I������Ă��܂���B"
            MsgBox "���Ɣ���𒆎~���܂��B"
            Exit Sub
        
        End If
        
        
        '�e�o�͏���'

        Range("S7").Value = decreaseRateKorona '�������o��'
        Range("R7").Value = outputKoronaPaymentPa '�R���i�e���O�N������(�R���i�֌W���t���T����)�̏o��'
        Range("R9").Value = outputKoronaPaymentPa2 '�R���i�e���O�N������(�R���i�֌W���t���T���O)�̏o��'
        Range("Q14").Value = outputKoronaIncomeS '�R���i�e���O�N�������̏o�͂P'
        Range("Q31").Value = outputKoronaIncomeS '�R���i�e���O�N�������̏o��2'
        Range("R14").Value = outputNonKoronaIncome '�R���i�e���O�O�N�������̏o��'
        Range("P34").Value = outputObjectPrice '�U�����ƑΏەی����z�̏o��'
        
        
        
        '���Ɗz�v�Z�͂�������'
        
        Dim nushiTotalIncome As Long '���ю�̑O�N�����v����'
        Dim exemptionRate6 As Variant '�R���i���Ɨ�'
        
        
        nushiTotalIncome = Range("S14").Value
        
        
        
        '���Ɨ��̌v�Z'
        If decreaseRateKorona >= 0.3 And outputNonKoronaIncome < 4000000 Then
        
            Select Case nushiTotalIncome
                Case 0 To 3000000
                    exemptionRate6 = 1
                    Range("Q42").Interior.ColorIndex = 15
                
                Case 3000001 To 4000000
                    exemptionRate6 = 0.8
                    Range("Q43").Interior.ColorIndex = 15
                
                Case 4000001 To 5500000
                    exemptionRate6 = 0.6
                    Range("Q44").Interior.ColorIndex = 15
                
                Case 5500001 To 7500000
                    exemptionRate6 = 0.4
                    Range("Q45").Interior.ColorIndex = 15
                
                Case 7500001 To 10000000
                    exemptionRate6 = 0.2
                    Range("Q46").Interior.ColorIndex = 15
                
                Case Else
                    exemptionRate6 = 0
            
            End Select
            
            outputExePrice.Value = outputObjectPrice * exemptionRate6
            outputExeRate.Value = Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5) * exemptionRate6
        
        ElseIf decreaseRateKorona < 0.3 And outputNonKoronaIncome < 4000000 Then
            outputExePrice.Value = outputObjectPrice * 0
            outputExeRate.Value = "���ƕs��(<30%)"
            MsgBox "������������30��������邽�߁A�U�����Ƃ��ł��܂���B"
        
        ElseIf outputNonKoronaIncome > 4000000 Then
            outputExePrice.Value = outputObjectPrice * 0
            outputExeRate.Value = "���ƕs��(>400��)"
            MsgBox "�R���i�e���O������400���~�𒴂��邽�߁A�U�����Ƃ��ł��܂���B"
        
        End If
    End If
End Sub

Attribute VB_Name = "�}�j�t�F�X�g���W���[��"
Sub CommandButton2_Click()
    Dim i, s, sumS, sumP, annualIncomeO As Long
    Dim annualIncomeOInput As Range '�P����������̂��̑��������z
    Dim counter As Long '�o�͗p�s�ԍ�'
    Dim counter2 As Long '�N��s�ԍ�'
    Dim specialDeduction As Long '�������z�����T��'
    Dim sumIncomePS As Long '���^����+�N������(�������z�����T���z�Z�o�p)'
    Dim specialDeductionInput As Range '�������z�����T���o�͗p'
    Dim nushiSection As String '���ю�敪'
    Dim koronaSection As String '���ю�R���i�e���̗L��'
    Dim res As String '���b�Z�[�W�{�b�N�X�̃��X�|���X����'
    
    
    nushiSection = Range("C1").Value '���ю�敪�̎擾'
    koronaSection = Range("C25").Value '�R���i�e���̗L���̎擾'
    specialDeduction = 0 '�������z�����T���z�̏�����
    
    '�[�����ю��6�����Ɣ��肵�Ă������̊m�F
    If koronaSection = "�L" And nushiSection = "�[�����ю�" Then
        res = MsgBox("�[�����ю吢�тł����A�R���i�e���y�L�z�ƂȂ��Ă��܂��B" & vbCrLf & "���̂܂܏����𑱍s���Ă�낵���ł���?", vbYesNo)
        If res = vbNo Then
            MsgBox "�����𒆎~���܂����B"
            Exit Sub
        Else
            GoTo MainProcessing
        End If 'If res = vbNo Then'
    Else
        GoTo MainProcessing
    End If
    
    
MainProcessing:
    counter = 44 '���N�̏����̏o�͗p'
    counter2 = 3 '�ΏێҔN��擾�p'
   
    For i = 7 To 25 Step 3
        Set monthlyIncomeSInput = Cells(counter, 11) '���N�̋��^�����o�͐�'
        Set annualIncomePInput = Cells(counter, 12) '���N�̔N�������o�͐�'
        Set annualIncomeOInput = Cells(counter, 13) '���N�̂��̑������o�͐�

        If Cells(i, 9).Value = "" Then
            Exit For
        Else
            '2�����Ɨp�R�������^�����A�N�������A���̑������̍��v�z���Z�o
            sumS = Application.WorksheetFunction.Sum(Range(Cells(i, 11), Cells(i + 2, 11))) '���͋��^����(3������)�̍��v�z�̎Z�o'
            sumP = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i + 2, 12))) '���͔N������(3������)�̍��v�z�̎Z�o'
            sumO = Application.WorksheetFunction.Sum(Range(Cells(i, 13), Cells(i + 2, 13))) '���͂��̑�����(3������)�̍��v�z�̎Z�o'
            age = Cells(counter2, 4).Value
            
            annualSumS = sumS * 4 '�N�ԋ��^�����̎Z�o'
            annualSumP = sumP * 4 '�N�ԔN�������̎Z�o'
            annualIncomeO = sumO * 4 '�N�Ԃ��̑������̎Z�o'
            
             '���Ԃ��̑��������z�̏o��'
            annualIncomeOInput.Value = annualIncomeO / 12
            
            '���^�����Z�o�͂�������'
            Call ���^�����Z�o.���^�����Z�o
            
            '���I�N�����ɌW��G�����ȊO�̍��v�����̎Z�o'
            annualTotalIncome = annualIncomeS + annualIncomeO
            
            '�N�������T���z�̎Z�o�͂�������'
            Call ���I�N�������Z�o.���I�N�������Z�o
        End If
        
        
        '�������z�����T���z�̎Z�o����
        Call �������z�����T���z�Z�o.�������z�����T���z�Z�o
        specialDeduction = specialDeduction + addDeduction '�������z�����T���Ώێ҂���������ꍇ�́A���Z����B
        
        '���ԔN�������̏o��'
        annualIncomePInput.Value = annualIncomeP / 12
         
        '���ԋ��^����(�������z�����T����)�̏o��'
        monthlyIncomeSInput.Value = (annualIncomeS - addDeduction) / 12
         
         '�J�E���^�[�����Z
        counter = counter + 1
        counter2 = counter2 + 1
       
    Next i
    
    
    '�������z�����T���z���V�[�g�֏o��'
    Set specialDeductionInput = Range("M56") '�������z�����T���o�͐�'
    specialDeductionInput.Value = specialDeduction
    
    
    
    '���Ɣ��菈���͂�������'
    
    '�������z�Z�o����'
    Call ���Ɨp�������z�Z�o.���Ɨp�������z�Z�o
    
    '5�����Ɣ��菈��'
    Call ���Ɣ���5��.���Ɣ���5��
    
    
    '2�����Ɣ��菈��'
    Call ���Ɣ���2��.���Ɣ���2��
    
    '6�����Ɣ��菈��'
    
    If Range("C25").Value = "��" Then
        GoTo ExitLabel
        
    '6�����ƃX�L�b�v���x��'
ExitLabel:
        MsgBox "�Z���Ɣ��茋��" & vbCrLf & vbCrLf & "�T�����Ɨ�:" & output5.Value & vbCrLf & "�Q�����Ɨ�:" & outputExemptionRate.Value & vbCrLf & "�U�����Ɨ�:" & 0 & vbCrLf & "�U�����Ɗz:" & 0
        Worksheets("���Ɣ���p").PrintOut
    
    Else
        
        Call ���Ɣ���6��.���Ɣ���6��
        
    End If
        
        
    MsgBox "�Z���Ɣ��茋��" & vbCrLf & vbCrLf & "�T�����Ɨ�:" & output5.Value & vbCrLf & "�Q�����Ɨ�:" & outputExemptionRate.Value & vbCrLf & "�U�����Ɨ�:" & outputExeRate.Value & vbCrLf & "�U�����Ɗz:" & outputExePrice.Value
    Worksheets("���Ɣ���p").PrintOut
    
End Sub


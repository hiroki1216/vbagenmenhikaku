Attribute VB_Name = "�������z�����T���z�Z�o"
Sub �������z�����T���z�Z�o()
    Dim saDeduction As Long '���^������(���10���~)'
    Dim penDeduction As Long '�N��������(���10���~)'
    
    '�������z�����T�����Z�p�̏�����'
    addDeduction = 0
    
    '�@���^�������̎Z�o(���10���~)'
    Select Case annualIncomeS
        Case 0 To 100000
            saDeduction = annualIncomeS
        
        Case Is > 100000
            saDeduction = 100000
        
        Case Else
            saDeduction = 0
    
    End Select
    
    '�A�N���������̎Z�o(���10���~)'
    Select Case annualIncomeP
    
        Case 0 To 100000
            penDeduction = annualIncomeP
        
        Case Is > 100000
            penDeduction = 100000
        
        Case Else
            penDeduction = 0
    
    End Select
    
    
    
    '�B�������z�����T���z�̎Z�o(���10���~)'
    If saDeduction + penDeduction - 100000 > 0 Then
        addDeduction = saDeduction + penDeduction - 100000
    Else
        addDeduction = 0
    End If
       
End Sub

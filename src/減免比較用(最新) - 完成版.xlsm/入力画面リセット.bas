Attribute VB_Name = "���͉�ʃ��Z�b�g"
Sub CommandButton3_Click()
    '�����ԍ��̃��Z�b�g����'
    Dim atenaNum As Range
    Dim i As Long '�e���ڃ��Z�b�g�p'
    Set atenaNum = Range("B3:B9")
    atenaNum.ClearContents
    atenaNum.Interior.Color = RGB(245, 245, 245)
    
    
    '�e���ڃ��Z�b�g����'
    For i = 13 To 19 Step 2
        Cells(i, 3).ClearContents
    Next i
    
    '�R���i�֌W���t�����Z�b�g'
    Range("C23").ClearContents
    
    '�R���i�e���̗L�����Z�b�g'
    Range("C25").ClearContents
    
    '�T�����Ɠ��͓��e�̃��Z�b�g����'
    Dim fiveGInput As Range
    Set fiveGInput = Range("D29:F49")
    fiveGInput.ClearContents
    
    '�T�����Ɨ��̃��Z�b�g����'
    Range("G55").ClearContents
    
    
    '�Q�����Ɠ��͓��e�̃��Z�b�g����'
    Dim twoGInput1 As Range
    Set twoGInput1 = Range("J7:M27")
    twoGInput1.ClearContents
    
    '�Q�����ƂւT�����Ɠ��͓��e�̍đ��'
    Dim t As Long '�Q�����ƃ��Z�b�g�p1'
    Dim s As Long '�Q�����ƃ��Z�b�g�p2'
    s = 29
    For t = 7 To 27
        Cells(t, 10).Value = "=C" & s
        Cells(t, 11).Value = "=D" & s
        Cells(t, 12).Value = "=E" & s
        Cells(t, 13).Value = "=F" & s
        s = s + 1
    Next t
    
    '�Q�����ƍ��N���������Z�b�g'
    Dim twoGInput2 As Range
    Set twoGInput2 = Range("K44:M50")
    twoGInput2.ClearContents
    
    '�Q�����ƌ��������Z�b�g'
    Range("J55").ClearContents
    
    '�Q�����Ɨ����Z�b�g'
    Range("N55").ClearContents
    
    '�U��'
    '�U���O�N�������E���������Z�b�g'
    Dim rate1 As Range
    Set rate1 = Range("R7:S7")
    rate1.ClearContents
    
    '�U���O�N���e���O�E�e���������Z�b�g'
    Dim rate2 As Range
    Set rate2 = Range("Q14:R14")
    rate2.ClearContents
    
    '�U���O�N���������Z�b�g'
    Range("Q31").ClearContents
    
    '�U���O�N�����ƑΏەی���'
    Range("P34").ClearContents
    
    '���Ɗz���Z�b�g'
    Range("S54").ClearContents
    
    '���Ɨ����Z�b�g'
    Range("S55").ClearContents
    
    '�ߘa�Q�N�������Z�b�g'
    Range("R9").ClearContents
    
    '�T�����Ɣ����z���Z�b�g'
    Range("C55").ClearContents
    
    '�U�����ƌ��Ɨ��w�i�F�̃��Z�b�g'
    Range("Q42:Q46").Interior.ColorIndex = 0
    
    '��۔ԍ��̃��Z�b�g'
    Range("B1").ClearContents
    
    '���������v�z�̃��Z�b�g'
    Range("C56").ClearContents
    Range("J56").ClearContents
    
    '����(�n�_)�̃��Z�b�g'
    Range("E55").ClearContents
    Range("L55").ClearContents
    
    '�������z�����T���z�̃��Z�b�g'
    Range("M56").ClearContents
    
    

End Sub

Attribute VB_Name = "�����ی����擾"
Sub CommandButton1_Click()
    Dim wb1 As Workbook '���Ɣ���p�u�b�N'
    Dim wb2 As Workbook '�T�����Ɣ���p�u�b�N'
    Dim fileName As String '�o�͐�̃t�@�C����'
    
    Dim i As Long '���Ɣ���p�s�ԍ�'
    Dim s As Long '�o�͐�v�Z�V�[�g�p�s�ԍ�'
    Dim a As Long '�N��擾�p�s�ԍ�'
    
    Dim counta As Long '���ۉ����Ґ��̎擾'
    Dim housingAllowance As Long '�Z��}���z'
    Dim educateAllowance  As Long '����}���z'
    Dim markup  As Long '����}�����Z�z'
    Dim age As Long '����}���z�v�Z�p�N��'
    
    
    fileName = "�ߘa�R�N�x(�����ی��z�v�Z�c�[��).xls"
    
    '���ۉ����Ґ��̎擾'
    
    counta = Range("B2").End(xlDown).Row - 2
        
    Set wb1 = ThisWorkbook
    Set wb2 = Workbooks.Open("\\J16sv009\���L\���N�ی���\��p�t�H���_\�y���ߘa�R�N�x ���ƌv�Z�z\" & fileName)
    
    s = 6
    For i = 3 To Range("B2").End(xlDown).Row
        If Range("B4").Value = "" Then
            i = 3
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 5).Value = wb1.Worksheets("���Ɣ���p").Cells(i, 4).Value
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 7).Value = "�P���n�|�P"
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 10).Value = "�Y��"
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 12).Value = "����"
            Exit For
        Else
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 5).Value = wb1.Worksheets("���Ɣ���p").Cells(i, 4).Value
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 7).Value = "�P���n�|�P"
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 10).Value = "�Y��"
            wb2.Worksheets("�v�Z�V�[�g").Cells(s, 12).Value = "����"
        End If
        s = s + 1
    Next i
    
    '�Z��}���z�̎Z�o'
    Select Case counta
        Case Is = 1
            housingAllowance = 39000
        
        Case Is = 2
            housingAllowance = 47000
        
        Case 3 To 5
            housingAllowance = 51000
        
        Case Is = 6
            housingAllowance = 55000
        
        Case Is = 7
            housingAllowance = 61000
        
        Case Else
            housingAllowance = 61000
    End Select
    
    '����}���z�̎Z�o'
    educateAllowance = 0
    For a = 3 To Range("B2").End(xlDown).Row
        If Range("B4").Value = "" Then
            a = 3
            age = wb1.Worksheets("���Ɣ���p").Cells(a, 4).Value
            educateAllowance = educateAllowance + 0
            Exit For
        
        Else
            age = wb1.Worksheets("���Ɣ���p").Cells(a, 4).Value
        End If
        
        Select Case age
        Case 6 To 7
            markup = 7050
            educateAllowance = educateAllowance + markup
            
        Case 8
            educateAllowance = 0
            MsgBox "�W�΂̉����҂����܂��B" & vbCrLf & "���w�Q�N���������͏��w�R�N���ł��B" & vbCrLf & "����}���z���v�Z�̂����A����͂��Ă��������B"
            Exit For
        
        Case 9 To 11
            markup = 7150
            educateAllowance = educateAllowance + markup
            
        Case 12
            educateAllowance = 0
            MsgBox "12�΂̉����҂����܂��B" & vbCrLf & "���w6�N���������͒��w1�N���ł��B" & vbCrLf & "����}���z���v�Z�̂����A����͂��Ă��������B"
            Exit For
        
        Case 13 To 14
            markup = 10690
            educateAllowance = educateAllowance + markup
            
        Case 15
            educateAllowance = 0
            MsgBox "15�΂̉����҂����܂��B" & vbCrLf & "���w�R�N���������͍��Z1�N���ł��B" & vbCrLf & "����}���z���v�Z�̂����A����͂��Ă��������B"
            Exit For
        
        Case Else
            markup = 0
            educateAllowance = educateAllowance + markup

        End Select
    
    Next a
    
    
    '�Z��}���z�̏o��'
    wb2.Worksheets("�v�Z�V�[�g").Cells(26, 14).Value = housingAllowance
    
    '����}���z�̏o��'
    wb2.Worksheets("�v�Z�V�[�g").Cells(26, 19).Value = educateAllowance


End Sub

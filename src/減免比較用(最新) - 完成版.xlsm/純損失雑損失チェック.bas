Attribute VB_Name = "�������G�����`�F�b�N"
Sub �������G�����`�F�b�N()
    Dim tblJZ As Range '�������E�G�����͈̔͗p'
    Dim keyJZ As Long '�������E�G�����̌����p'
    Dim J As Long '�������p'
    Dim Z As Long '�G�����p'
    Dim JZ As Long '�������E�G�����J�Ԃ��p'
    Dim last As Long '�J�ԉ񐔐���p
    
    Set tblJZ = Worksheets("�����E���Y���ꗗ").Range("C:AV") '�������E�G�����͈̔�'
    last = Range("B3").End(xlDown).Row '�J�ԉ�
    
    For JZ = 3 To last
    
        keyJZ = Cells(JZ, 2).Value '�������E�G�����̌����p�͈�'
        On Error Resume Next
        J = Application.WorksheetFunction.VLookup(keyJZ, tblJZ, 44, False) '�������̎擾'
        On Error Resume Next
        Z = Application.WorksheetFunction.VLookup(keyJZ, tblJZ, 46, False) '�G�����̎擾'
        
        '�������P�~�ȏォ����
        If J > 0 Then
            MsgBox "���������P�~�ȏ�̑Ώێ҂����܂��B�U�����Ƃ̍Čv�Z���s���Ă��������B" & vbCrLf & "A3�`A9�ŐԐF�̂����Z�����������Ώێ҂ł��B", vbExclamation
            Cells(JZ, 2).Interior.ColorIndex = 3
            
            outputExeRate.Value = "�Čv�Z�v(������)"
        '�G�����P�~�ȏォ����
        ElseIf Z > 0 Then
            MsgBox "�G�������P�~�ȏ�̑Ώێ҂����܂��B�U�����Ƃ̍Čv�Z���s���Ă��������B" & vbCrLf & "A3�`A9�ŉ��F�̂����Z�����G�����Ώێ҂ł��B", vbExclamation
            Cells(JZ, 2).Interior.ColorIndex = 6
            
            outputExeRate.Value = "�Čv�Z�v(�G����)"
        
        End If
    
    Next JZ

    MsgBox "�Z���Ɣ��茋��" & vbCrLf & vbCrLf & "�T�����Ɨ�:" & output5.Value & vbCrLf & "�Q�����Ɨ�:" & outputExemptionRate.Value & vbCrLf & "�U�����Ɨ�:" & outputExeRate.Value & vbCrLf & "�U�����Ɗz:" & outputExePrice.Value
    Worksheets("���Ɣ���p").PrintOut
   
End Sub



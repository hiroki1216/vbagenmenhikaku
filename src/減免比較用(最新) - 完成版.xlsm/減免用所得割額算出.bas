Attribute VB_Name = "���Ɨp�������z�Z�o"
Sub ���Ɨp�������z�Z�o()
    Dim objectPrice25 As Long '2�E5�����Ɗz�Z�o�p���������v�z'
    Dim tbl25 As Range 'v-lookup���������v�z�擾�p�͈�'
    Dim key25 As Long  'v-lookup�����p�L�['
    Dim sIryou As Long '��Õ��������z�̎擾'
    Dim sShiennkinn  As Long '�x�������������z�̎擾'
    Dim sKaigo As Long '��앪�������z�̎擾
    
    Set tbl25 = Worksheets("���ۏ��ꗗ").Range("C:DN")
    key25 = Range("B1").Value
    
    '��Õ��������z�̎擾'
    On Error Resume Next
        sIryou = Application.WorksheetFunction.VLookup(key25, tbl25, 34, False)
        If Err.Number <> 0 Then
            sIryou = -999999
            MsgBox "��Õ��������z�̎擾�Ɏ��s���܂����B"
        End If
    On Error GoTo 0
    
    '�x�������������z�̎擾'
    On Error Resume Next
        sShiennkinn = Application.WorksheetFunction.VLookup(key25, tbl25, 62, False)
        If Err.Number <> 0 Then
            sShiennkinn = -999999
            MsgBox "�x�������������z�̎擾�Ɏ��s���܂����B"
        End If
    On Error GoTo 0
    
    '��앪�������z�̎擾'
    On Error Resume Next
        sKaigo = Application.WorksheetFunction.VLookup(key25, tbl25, 90, False)
        If Err.Number <> 0 Then
            sKaigo = -999999
            MsgBox "��앪�������z�̎擾�Ɏ��s���܂����B"
        End If
    On Error GoTo 0
    
    '���������v�z�̎擾'
    objectPrice25 = sIryou + sShiennkinn + sKaigo
    
    '���������v�z�̏o��'
    Range("C56").Value = objectPrice25
    Range("J56").Value = objectPrice25
End Sub

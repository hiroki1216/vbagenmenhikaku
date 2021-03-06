Attribute VB_Name = "減免用所得割額算出"
Sub 減免用所得割額算出()
    Dim objectPrice25 As Long '2・5号減免額算出用所得割合計額'
    Dim tbl25 As Range 'v-lookup所得割合計額取得用範囲'
    Dim key25 As Long  'v-lookup検索用キー'
    Dim sIryou As Long '医療分所得割額の取得'
    Dim sShiennkinn  As Long '支援金分所得割額の取得'
    Dim sKaigo As Long '介護分所得割額の取得
    
    Set tbl25 = Worksheets("賦課情報一覧").Range("C:DN")
    key25 = Range("B1").Value
    
    '医療分所得割額の取得'
    On Error Resume Next
        sIryou = Application.WorksheetFunction.VLookup(key25, tbl25, 34, False)
        If Err.Number <> 0 Then
            sIryou = -999999
            MsgBox "医療分所得割額の取得に失敗しました。"
        End If
    On Error GoTo 0
    
    '支援金分所得割額の取得'
    On Error Resume Next
        sShiennkinn = Application.WorksheetFunction.VLookup(key25, tbl25, 62, False)
        If Err.Number <> 0 Then
            sShiennkinn = -999999
            MsgBox "支援金分所得割額の取得に失敗しました。"
        End If
    On Error GoTo 0
    
    '介護分所得割額の取得'
    On Error Resume Next
        sKaigo = Application.WorksheetFunction.VLookup(key25, tbl25, 90, False)
        If Err.Number <> 0 Then
            sKaigo = -999999
            MsgBox "介護分所得割額の取得に失敗しました。"
        End If
    On Error GoTo 0
    
    '所得割合計額の取得'
    objectPrice25 = sIryou + sShiennkinn + sKaigo
    
    '所得割合計額の出力'
    Range("C56").Value = objectPrice25
    Range("J56").Value = objectPrice25
End Sub

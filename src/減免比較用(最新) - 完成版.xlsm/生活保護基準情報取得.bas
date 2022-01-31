Attribute VB_Name = "生活保護基準情報取得"
Sub CommandButton1_Click()
    Dim wb1 As Workbook '減免判定用ブック'
    Dim wb2 As Workbook '５号減免判定用ブック'
    Dim fileName As String '出力先のファイル名'
    
    Dim i As Long '減免判定用行番号'
    Dim s As Long '出力先計算シート用行番号'
    Dim a As Long '年齢取得用行番号'
    
    Dim counta As Long '国保加入者数の取得'
    Dim housingAllowance As Long '住宅扶助額'
    Dim educateAllowance  As Long '教育扶助額'
    Dim markup  As Long '教育扶助加算額'
    Dim age As Long '教育扶助額計算用年齢'
    
    
    fileName = "令和３年度(生活保護基準額計算ツール).xls"
    
    '国保加入者数の取得'
    
    counta = Range("B2").End(xlDown).Row - 2
        
    Set wb1 = ThisWorkbook
    Set wb2 = Workbooks.Open("\\J16sv009\共有\健康保険課\専用フォルダ\【★令和３年度 減免計算】\" & fileName)
    
    s = 6
    For i = 3 To Range("B2").End(xlDown).Row
        If Range("B4").Value = "" Then
            i = 3
            wb2.Worksheets("計算シート").Cells(s, 5).Value = wb1.Worksheets("減免判定用").Cells(i, 4).Value
            wb2.Worksheets("計算シート").Cells(s, 7).Value = "１級地－１"
            wb2.Worksheets("計算シート").Cells(s, 10).Value = "Ⅵ区"
            wb2.Worksheets("計算シート").Cells(s, 12).Value = "居宅"
            Exit For
        Else
            wb2.Worksheets("計算シート").Cells(s, 5).Value = wb1.Worksheets("減免判定用").Cells(i, 4).Value
            wb2.Worksheets("計算シート").Cells(s, 7).Value = "１級地－１"
            wb2.Worksheets("計算シート").Cells(s, 10).Value = "Ⅵ区"
            wb2.Worksheets("計算シート").Cells(s, 12).Value = "居宅"
        End If
        s = s + 1
    Next i
    
    '住宅扶助額の算出'
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
    
    '教育扶助額の算出'
    educateAllowance = 0
    For a = 3 To Range("B2").End(xlDown).Row
        If Range("B4").Value = "" Then
            a = 3
            age = wb1.Worksheets("減免判定用").Cells(a, 4).Value
            educateAllowance = educateAllowance + 0
            Exit For
        
        Else
            age = wb1.Worksheets("減免判定用").Cells(a, 4).Value
        End If
        
        Select Case age
        Case 6 To 7
            markup = 7050
            educateAllowance = educateAllowance + markup
            
        Case 8
            educateAllowance = 0
            MsgBox "８歳の加入者がいます。" & vbCrLf & "小学２年生もしくは小学３年生です。" & vbCrLf & "教育扶助額を計算のうえ、手入力してください。"
            Exit For
        
        Case 9 To 11
            markup = 7150
            educateAllowance = educateAllowance + markup
            
        Case 12
            educateAllowance = 0
            MsgBox "12歳の加入者がいます。" & vbCrLf & "小学6年生もしくは中学1年生です。" & vbCrLf & "教育扶助額を計算のうえ、手入力してください。"
            Exit For
        
        Case 13 To 14
            markup = 10690
            educateAllowance = educateAllowance + markup
            
        Case 15
            educateAllowance = 0
            MsgBox "15歳の加入者がいます。" & vbCrLf & "中学３年生もしくは高校1年生です。" & vbCrLf & "教育扶助額を計算のうえ、手入力してください。"
            Exit For
        
        Case Else
            markup = 0
            educateAllowance = educateAllowance + markup

        End Select
    
    Next a
    
    
    '住宅扶助額の出力'
    wb2.Worksheets("計算シート").Cells(26, 14).Value = housingAllowance
    
    '教育扶助額の出力'
    wb2.Worksheets("計算シート").Cells(26, 19).Value = educateAllowance


End Sub

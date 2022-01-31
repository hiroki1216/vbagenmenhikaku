Attribute VB_Name = "マニフェストモジュール"
Sub CommandButton2_Click()
    Dim i, s, sumS, sumP, annualIncomeO As Long
    Dim annualIncomeOInput As Range '１ヶ月当たりのその他所得金額
    Dim counter As Long '出力用行番号'
    Dim counter2 As Long '年齢行番号'
    Dim specialDeduction As Long '所得金額調整控除'
    Dim sumIncomePS As Long '給与所得+年金所得(所得金額調整控除額算出用)'
    Dim specialDeductionInput As Range '所得金額調整控除出力用'
    Dim nushiSection As String '世帯主区分'
    Dim koronaSection As String '世帯主コロナ影響の有無'
    Dim res As String 'メッセージボックスのレスポンス結果'
    
    
    nushiSection = Range("C1").Value '世帯主区分の取得'
    koronaSection = Range("C25").Value 'コロナ影響の有無の取得'
    specialDeduction = 0 '所得金額調整控除額の初期化
    
    '擬制世帯主で6号減免判定していいかの確認
    If koronaSection = "有" And nushiSection = "擬制世帯主" Then
        res = MsgBox("擬制世帯主世帯ですが、コロナ影響【有】となっています。" & vbCrLf & "このまま処理を続行してよろしいですか?", vbYesNo)
        If res = vbNo Then
            MsgBox "処理を中止しました。"
            Exit Sub
        Else
            GoTo MainProcessing
        End If 'If res = vbNo Then'
    Else
        GoTo MainProcessing
    End If
    
    
MainProcessing:
    counter = 44 '今年の所得の出力用'
    counter2 = 3 '対象者年齢取得用'
   
    For i = 7 To 25 Step 3
        Set monthlyIncomeSInput = Cells(counter, 11) '今年の給与所得出力先'
        Set annualIncomePInput = Cells(counter, 12) '今年の年金所得出力先'
        Set annualIncomeOInput = Cells(counter, 13) '今年のその他所得出力先

        If Cells(i, 9).Value = "" Then
            Exit For
        Else
            '2号減免用３ヶ月給与収入、年金収入、その他所得の合計額を算出
            sumS = Application.WorksheetFunction.Sum(Range(Cells(i, 11), Cells(i + 2, 11))) '入力給与収入(3ヶ月分)の合計額の算出'
            sumP = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i + 2, 12))) '入力年金収入(3ヶ月分)の合計額の算出'
            sumO = Application.WorksheetFunction.Sum(Range(Cells(i, 13), Cells(i + 2, 13))) '入力その他所得(3ヶ月分)の合計額の算出'
            age = Cells(counter2, 4).Value
            
            annualSumS = sumS * 4 '年間給与収入の算出'
            annualSumP = sumP * 4 '年間年金収入の算出'
            annualIncomeO = sumO * 4 '年間その他所得の算出'
            
             '月間その他所得金額の出力'
            annualIncomeOInput.Value = annualIncomeO / 12
            
            '給与所得算出はここから'
            Call 給与所得算出.給与所得算出
            
            '公的年金等に係る雑所得以外の合計所得の算出'
            annualTotalIncome = annualIncomeS + annualIncomeO
            
            '年金所得控除額の算出はここから'
            Call 公的年金所得算出.公的年金所得算出
        End If
        
        
        '所得金額調整控除額の算出処理
        Call 所得金額調整控除額算出.所得金額調整控除額算出
        specialDeduction = specialDeduction + addDeduction '所得金額調整控除対象者が複数いる場合は、加算する。
        
        '月間年金所得の出力'
        annualIncomePInput.Value = annualIncomeP / 12
         
        '月間給与所得(所得金額調整控除後)の出力'
        monthlyIncomeSInput.Value = (annualIncomeS - addDeduction) / 12
         
         'カウンターを加算
        counter = counter + 1
        counter2 = counter2 + 1
       
    Next i
    
    
    '所得金額調整控除額をシートへ出力'
    Set specialDeductionInput = Range("M56") '所得金額調整控除出力先'
    specialDeductionInput.Value = specialDeduction
    
    
    
    '減免判定処理はここから'
    
    '所得割額算出処理'
    Call 減免用所得割額算出.減免用所得割額算出
    
    '5号減免判定処理'
    Call 減免判定5号.減免判定5号
    
    
    '2号減免判定処理'
    Call 減免判定2号.減免判定2号
    
    '6号減免判定処理'
    
    If Range("C25").Value = "無" Then
        GoTo ExitLabel
        
    '6号減免スキップラベル'
ExitLabel:
        MsgBox "〇減免判定結果" & vbCrLf & vbCrLf & "５号減免率:" & output5.Value & vbCrLf & "２号減免率:" & outputExemptionRate.Value & vbCrLf & "６号減免率:" & 0 & vbCrLf & "６号減免額:" & 0
        Worksheets("減免判定用").PrintOut
    
    Else
        
        Call 減免判定6号.減免判定6号
        
    End If
        
        
    MsgBox "〇減免判定結果" & vbCrLf & vbCrLf & "５号減免率:" & output5.Value & vbCrLf & "２号減免率:" & outputExemptionRate.Value & vbCrLf & "６号減免率:" & outputExeRate.Value & vbCrLf & "６号減免額:" & outputExePrice.Value
    Worksheets("減免判定用").PrintOut
    
End Sub


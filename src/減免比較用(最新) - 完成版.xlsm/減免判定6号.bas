Attribute VB_Name = "減免判定6号"
Sub 減免判定6号()
    Set outputExePrice = Range("S54") '6号減免額出力セルオブジェクトの取得'
    Set outputExeRate = Range("S55") '6号減免率出力セルオブジェクトの取得'
    
    If Range("G3").Value <> 0 Then
        MsgBox "世帯主が非自発的失業者のため6号減免不可です。"
        outputExeRate.Value = 0
        outputExePrice.Value = "減免不可(主が非自発)"
    
    Else
    
        Dim outputKoronakyuuhu As Long '前年中コロナ関係給付金額'
        Dim outputKoronaPaymentPa As Long '前年中コロナ影響収入(前年中コロナ関係給付金額控除後)'
        Dim outputKoronaPaymentPa2 As Long '前年中コロナ影響収入(前年中コロナ関係給付金額控除前)'
        Dim outputKoronaPaymentC As Long '今年中のコロナ影響見込み収入'
        
        Dim outputKoronaIncomeS As Long '前年中のコロナ影響所得'
        Dim outputTotalIncome As Long '前年中の合計所得'
        Dim outputNonKoronaIncome As Long '前年中のコロナ影響外所得'
        Dim outputObjectPrice As Long '減免対象保険料'
        Dim annualCost As Variant '算定保険料額’
        Dim familyTotalIncome As Long '前年中の世帯合計所得'
        Dim influencePayment As String 'コロナ影響収入名'
        
        Dim tbl1 As Range 'v-lookup検索用範囲(前年中のコロナ影響収入）'
        Dim tbl2 As Range 'v-lookup検索用範囲(前年中のコロナ影響所得）'
        
        Dim key As Long 'v-lookup検索用キー'
        Dim decreaseRateKorona As Variant 'コロナ影響収入減少率'
        Dim message As String 'エラーメッセージ’
        
        outputKoronaPaymentC = Range("Q7").Value
        key = Range("B3").Value
        influencePayment = Range("Q2").Value
        annualCost = Range("P31").Value
        familyTotalIncome = Range("R31").Value
        
        
        '影響収入で条件分岐処理’
        
        If influencePayment = "給与収入" Then
        
            Set tbl = Worksheets("汎用抽出（所得・資産）所得情報(詳細)").Range("C:AZ")
            
            '前年の影響収入(コロナ関係給付金控除後)の取得'
            On Error Resume Next
            outputKoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 13, False) - Range("C23").Value
            If Err.Number <> 0 Then
                outputKoronaPaymentPa = -999999
                MsgBox "前年の影響収入の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            '前年の影響収入(コロナ関係給付金控除前)の取得'
            On Error Resume Next
            outputKoronaPaymentPa2 = Application.WorksheetFunction.VLookup(key, tbl, 13, False)
            If Err.Number <> 0 Then
                outputKoronaPaymentPa2 = -999999
                MsgBox "前年の影響収入の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            
            '前年の影響給与所得の取得'
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 26, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "前年の影響所得の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            
            '前年の合計所得の取得'
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "前年の合計所得の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            
            '前年の影響外所得の取得'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            '影響収入の減少率'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '減免対象保険料額の算出'
            On Error Resume Next
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
            
            Dim KoronaPaymentPa As Long '前年コロナ影響収入(給与年金収入以外)'
            Dim inputResult As Long '入力結果'
        
        ElseIf influencePayment = "営業収入" Then
        
            Set tbl = Worksheets("汎用抽出（所得・資産）所得情報(詳細)").Range("C:AZ")
            
            KoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) '前年コロナ影響営業収入の取得'
            
            
            '前年中の収入の申告の有無で条件分岐ここから↓'
            
            If KoronaPaymentPa > 0 Then
                '前年の影響収入の取得(コロナ関係給付金控除後)'
                On Error Resume Next
                outputKoronaPaymentPa = KoronaPaymentPa - Range("C23").Value
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa = -999999
                    MsgBox "前年の影響収入の取得に失敗しました。"
                End If
                On Error GoTo 0
                '前年の影響収入の取得(コロナ関係給付金控除前)'
                On Error Resume Next
                outputKoronaPaymentPa2 = KoronaPaymentPa
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa2 = -999999
                    MsgBox "前年の影響収入の取得に失敗しました。"
                End If
                On Error GoTo 0
            ElseIf KoronaPaymentPa = 0 Then
            
                inputResult = Application.InputBox("前年営業収入が0です。" & vbCrLf & "前年営業収入を入力して下さい。")
                '前年の影響収入の取得(コロナ関係給付金控除後)'
                outputKoronaPaymentPa = inputResult - Range("C23").Value
                '前年の影響収入の取得(コロナ関係給付金控除前)'
                outputKoronaPaymentPa2 = inputResult
            End If
            
            '前年中の収入の申告の有無で条件分岐ここまで↑'
            
            
            '前年の影響所得の取得'
            
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 28, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "前年の影響所得の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            '前年の合計所得の取得'
            
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "前年の合計所得の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            '前年の影響外所得の取得'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            '影響収入の減少率'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '減免対象保険料額の算出'
            On Error Resume Next
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
        
        ElseIf influencePayment = "その他収入" Then
        
            Set tbl = Worksheets("汎用抽出（所得・資産）所得情報(詳細)").Range("C:AZ")
            
            KoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) '前年コロナ影響その他収入の取得'
            
            
            '前年中の収入の申告の有無で条件分岐ここから↓'
            
            If KoronaPaymentPa > 0 Then
                
                '前年の影響収入の取得(コロナ関係給付金控除後)'
                On Error Resume Next
                outputKoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) - Range("C23").Value
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa = -999999
                    MsgBox "前年の影響収入の取得に失敗しました。"
                End If
                On Error GoTo 0
                
                '前年の影響収入の取得(コロナ関係給付金控除前)'
                On Error Resume Next
                outputKoronaPaymentPa2 = Application.WorksheetFunction.VLookup(key, tbl, 17, False)
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa2 = -999999
                    MsgBox "前年の影響収入の取得に失敗しました。"
                End If
                On Error GoTo 0
            
            ElseIf KoronaPaymentPa = 0 Then
            
                inputResult = Application.InputBox("前年その他収入が0です。" & vbCrLf & "前年その他収入を入力して下さい。")
                '前年の影響収入の取得(コロナ関係給付金控除後)'
                outputKoronaPaymentPa = inputResult - Range("C23").Value
                '前年の影響収入の取得(コロナ関係給付金控除前)'
                outputKoronaPaymentPa2 = inputResult
            
            End If
            
            '前年中の収入の申告の有無で条件分岐ここまで↑'
            
            
            '前年の影響所得の取得'
            
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 36, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "前年の影響所得の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            '前年の合計所得の取得'
            
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "前年の合計所得の取得に失敗しました。"
            End If
            On Error GoTo 0
            
            '前年の影響外所得の取得'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            '影響収入の減少率'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '減免対象保険料額の算出'
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
        
        
        
        Else
            MsgBox "世帯主のコロナ影響収入が選択されていません。"
            MsgBox "減免判定を中止します。"
            Exit Sub
        
        End If
        
        
        '各出力処理'

        Range("S7").Value = decreaseRateKorona '減少率出力'
        Range("R7").Value = outputKoronaPaymentPa 'コロナ影響前年中収入(コロナ関係給付金控除後)の出力'
        Range("R9").Value = outputKoronaPaymentPa2 'コロナ影響前年中収入(コロナ関係給付金控除前)の出力'
        Range("Q14").Value = outputKoronaIncomeS 'コロナ影響前年中所得の出力１'
        Range("Q31").Value = outputKoronaIncomeS 'コロナ影響前年中所得の出力2'
        Range("R14").Value = outputNonKoronaIncome 'コロナ影響外前年中所得の出力'
        Range("P34").Value = outputObjectPrice '６号減免対象保険料額の出力'
        
        
        
        '減免額計算はここから'
        
        Dim nushiTotalIncome As Long '世帯主の前年中合計所得'
        Dim exemptionRate6 As Variant 'コロナ減免率'
        
        
        nushiTotalIncome = Range("S14").Value
        
        
        
        '減免率の計算'
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
            outputExeRate.Value = "減免不可(<30%)"
            MsgBox "収入減少率が30％を下回るため、６号減免ができません。"
        
        ElseIf outputNonKoronaIncome > 4000000 Then
            outputExePrice.Value = outputObjectPrice * 0
            outputExeRate.Value = "減免不可(>400万)"
            MsgBox "コロナ影響外所得が400万円を超えるため、６号減免ができません。"
        
        End If
    End If
End Sub

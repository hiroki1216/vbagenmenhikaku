Attribute VB_Name = "入力画面リセット"
Sub CommandButton3_Click()
    '宛名番号のリセット処理'
    Dim atenaNum As Range
    Dim i As Long '各項目リセット用'
    Set atenaNum = Range("B3:B9")
    atenaNum.ClearContents
    atenaNum.Interior.Color = RGB(245, 245, 245)
    
    
    '各項目リセット処理'
    For i = 13 To 19 Step 2
        Cells(i, 3).ClearContents
    Next i
    
    'コロナ関係給付金リセット'
    Range("C23").ClearContents
    
    'コロナ影響の有無リセット'
    Range("C25").ClearContents
    
    '５号減免入力内容のリセット処理'
    Dim fiveGInput As Range
    Set fiveGInput = Range("D29:F49")
    fiveGInput.ClearContents
    
    '５号減免率のリセット処理'
    Range("G55").ClearContents
    
    
    '２号減免入力内容のリセット処理'
    Dim twoGInput1 As Range
    Set twoGInput1 = Range("J7:M27")
    twoGInput1.ClearContents
    
    '２号減免へ５号減免入力内容の再代入'
    Dim t As Long '２号減免リセット用1'
    Dim s As Long '２号減免リセット用2'
    s = 29
    For t = 7 To 27
        Cells(t, 10).Value = "=C" & s
        Cells(t, 11).Value = "=D" & s
        Cells(t, 12).Value = "=E" & s
        Cells(t, 13).Value = "=F" & s
        s = s + 1
    Next t
    
    '２号減免今年中所得リセット'
    Dim twoGInput2 As Range
    Set twoGInput2 = Range("K44:M50")
    twoGInput2.ClearContents
    
    '２号減免減少率リセット'
    Range("J55").ClearContents
    
    '２号減免率リセット'
    Range("N55").ClearContents
    
    '６号'
    '６号前年中収入・減少率リセット'
    Dim rate1 As Range
    Set rate1 = Range("R7:S7")
    rate1.ClearContents
    
    '６号前年中影響外・影響所得リセット'
    Dim rate2 As Range
    Set rate2 = Range("Q14:R14")
    rate2.ClearContents
    
    '６号前年中所得リセット'
    Range("Q31").ClearContents
    
    '６号前年中減免対象保険料'
    Range("P34").ClearContents
    
    '減免額リセット'
    Range("S54").ClearContents
    
    '減免率リセット'
    Range("S55").ClearContents
    
    '令和２年収入リセット'
    Range("R9").ClearContents
    
    '５号減免判定基準額リセット'
    Range("C55").ClearContents
    
    '６号減免減免率背景色のリセット'
    Range("Q42:Q46").Interior.ColorIndex = 0
    
    '被保番号のリセット'
    Range("B1").ClearContents
    
    '所得割合計額のリセット'
    Range("C56").ClearContents
    Range("J56").ClearContents
    
    '期間(始点)のリセット'
    Range("E55").ClearContents
    Range("L55").ClearContents
    
    '所得金額調整控除額のリセット'
    Range("M56").ClearContents
    
    

End Sub

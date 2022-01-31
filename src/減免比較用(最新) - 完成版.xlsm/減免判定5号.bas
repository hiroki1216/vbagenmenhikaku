Attribute VB_Name = "Œ¸–Æ”»’è5†"
Sub Œ¸–Æ”»’è5†()
    Dim stdPrice As Long '‚T†Œ¸–ÆŠî€Šz'
    Dim avePrice As Long '’¼‹ß‚Rƒ–ŒŽ•½‹ÏŽû“üŠz'
    
    
    stdPrice = Range("C55").Value
    avePrice = Range("C52").Value
    Set output5 = Range("G55")
    Set outputTerm = Range("E55")
    
    
    applyMonth = Range("C11").Value
    
    'Œ¸–Æ“K—pŠúŠÔ‚Ìo—Íˆ—'
    Select Case applyMonth
        Case Is = 6
            outputTerm.Value = "‚S`"
        Case Is = 7
            outputTerm.Value = "‚T`"
        Case Is = 8
            outputTerm.Value = "‚U`"
        Case Is = 9
            outputTerm.Value = "‚V`"
        Case Is = 10
            outputTerm.Value = "‚W`"
        Case Is = 11
            outputTerm.Value = "10`"
        Case Is = 12
            outputTerm.Value = "11`"
        Case Is = 1
            outputTerm.Value = "12`"
        Case Is = 2
            outputTerm.Value = "‚P`"
        Case Is = 3
            outputTerm.Value = "‚Q`"
        Case Is = 4
            outputTerm.Value = "‚R`"
        Case Else
            outputTerm.Value = "Œ¸–Æ‘ÎÛŠúŠÔŠO‚Å‚·B"
    End Select
    
    'Œ¸–Æ—¦‚ÌŽZo'
    If stdPrice = 0 Then
        output5.Value = "Œ¸–Æ•s‰Â"
    
    ElseIf stdPrice >= avePrice And stdPrice > 0 Then
        output5.Value = "30%"
    
    Else
        output5.Value = "Œ¸–Æ•s‰Â"
    End If
End Sub

Attribute VB_Name = "Š“¾‹àŠz’²®TœŠzŽZo"
Sub Š“¾‹àŠz’²®TœŠzŽZo()
    Dim saDeduction As Long '‹‹—^Š“¾•ª(ãŒÀ10–œ‰~)'
    Dim penDeduction As Long '”N‹àŠ“¾•ª(ãŒÀ10–œ‰~)'
    
    'Š“¾‹àŠz’²®Tœ‰ÁŽZ—p‚Ì‰Šú‰»'
    addDeduction = 0
    
    '‡@‹‹—^Š“¾•ª‚ÌŽZo(ãŒÀ10–œ‰~)'
    Select Case annualIncomeS
        Case 0 To 100000
            saDeduction = annualIncomeS
        
        Case Is > 100000
            saDeduction = 100000
        
        Case Else
            saDeduction = 0
    
    End Select
    
    '‡A”N‹àŠ“¾•ª‚ÌŽZo(ãŒÀ10–œ‰~)'
    Select Case annualIncomeP
    
        Case 0 To 100000
            penDeduction = annualIncomeP
        
        Case Is > 100000
            penDeduction = 100000
        
        Case Else
            penDeduction = 0
    
    End Select
    
    
    
    '‡BŠ“¾‹àŠz’²®TœŠz‚ÌŽZo(ãŒÀ10–œ‰~)'
    If saDeduction + penDeduction - 100000 > 0 Then
        addDeduction = saDeduction + penDeduction - 100000
    Else
        addDeduction = 0
    End If
       
End Sub

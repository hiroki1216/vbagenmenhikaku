Attribute VB_Name = "¸Æ»è6"
Sub ¸Æ»è6()
    Set outputExePrice = Range("S54") '6¸ÆzoÍZIuWFNgÌæ¾'
    Set outputExeRate = Range("S55") '6¸Æ¦oÍZIuWFNgÌæ¾'
    
    If Range("G3").Value <> 0 Then
        MsgBox "¢Ñåªñ©­I¸ÆÒÌ½ß6¸ÆsÂÅ·B"
        outputExeRate.Value = 0
        outputExePrice.Value = "¸ÆsÂ(åªñ©­)"
    
    Else
    
        Dim outputKoronakyuuhu As Long 'ONRiÖWtàz'
        Dim outputKoronaPaymentPa As Long 'ONRie¿ûü(ONRiÖWtàzTã)'
        Dim outputKoronaPaymentPa2 As Long 'ONRie¿ûü(ONRiÖWtàzTO)'
        Dim outputKoronaPaymentC As Long '¡NÌRie¿©Ýûü'
        
        Dim outputKoronaIncomeS As Long 'ONÌRie¿¾'
        Dim outputTotalIncome As Long 'ONÌv¾'
        Dim outputNonKoronaIncome As Long 'ONÌRie¿O¾'
        Dim outputObjectPrice As Long '¸ÆÎÛÛ¯¿'
        Dim annualCost As Variant 'ZèÛ¯¿zf
        Dim familyTotalIncome As Long 'ONÌ¢Ñv¾'
        Dim influencePayment As String 'Rie¿ûü¼'
        
        Dim tbl1 As Range 'v-lookupõpÍÍ(ONÌRie¿ûüj'
        Dim tbl2 As Range 'v-lookupõpÍÍ(ONÌRie¿¾j'
        
        Dim key As Long 'v-lookupõpL['
        Dim decreaseRateKorona As Variant 'Rie¿ûü¸­¦'
        Dim message As String 'G[bZ[Wf
        
        outputKoronaPaymentC = Range("Q7").Value
        key = Range("B3").Value
        influencePayment = Range("Q2").Value
        annualCost = Range("P31").Value
        familyTotalIncome = Range("R31").Value
        
        
        'e¿ûüÅðªòf
        
        If influencePayment = "^ûü" Then
        
            Set tbl = Worksheets("Äpoi¾EYj¾îñ(Ú×)").Range("C:AZ")
            
            'ONÌe¿ûü(RiÖWtàTã)Ìæ¾'
            On Error Resume Next
            outputKoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 13, False) - Range("C23").Value
            If Err.Number <> 0 Then
                outputKoronaPaymentPa = -999999
                MsgBox "ONÌe¿ûüÌæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            'ONÌe¿ûü(RiÖWtàTO)Ìæ¾'
            On Error Resume Next
            outputKoronaPaymentPa2 = Application.WorksheetFunction.VLookup(key, tbl, 13, False)
            If Err.Number <> 0 Then
                outputKoronaPaymentPa2 = -999999
                MsgBox "ONÌe¿ûüÌæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            
            'ONÌe¿^¾Ìæ¾'
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 26, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "ONÌe¿¾Ìæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            
            'ONÌv¾Ìæ¾'
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "ONÌv¾Ìæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            
            'ONÌe¿O¾Ìæ¾'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            'e¿ûüÌ¸­¦'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '¸ÆÎÛÛ¯¿zÌZo'
            On Error Resume Next
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
            
            Dim KoronaPaymentPa As Long 'ONRie¿ûü(^NàûüÈO)'
            Dim inputResult As Long 'üÍÊ'
        
        ElseIf influencePayment = "cÆûü" Then
        
            Set tbl = Worksheets("Äpoi¾EYj¾îñ(Ú×)").Range("C:AZ")
            
            KoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) 'ONRie¿cÆûüÌæ¾'
            
            
            'ONÌûüÌ\ÌL³Åðªò±±©ç«'
            
            If KoronaPaymentPa > 0 Then
                'ONÌe¿ûüÌæ¾(RiÖWtàTã)'
                On Error Resume Next
                outputKoronaPaymentPa = KoronaPaymentPa - Range("C23").Value
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa = -999999
                    MsgBox "ONÌe¿ûüÌæ¾É¸sµÜµ½B"
                End If
                On Error GoTo 0
                'ONÌe¿ûüÌæ¾(RiÖWtàTO)'
                On Error Resume Next
                outputKoronaPaymentPa2 = KoronaPaymentPa
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa2 = -999999
                    MsgBox "ONÌe¿ûüÌæ¾É¸sµÜµ½B"
                End If
                On Error GoTo 0
            ElseIf KoronaPaymentPa = 0 Then
            
                inputResult = Application.InputBox("ONcÆûüª0Å·B" & vbCrLf & "ONcÆûüðüÍµÄº³¢B")
                'ONÌe¿ûüÌæ¾(RiÖWtàTã)'
                outputKoronaPaymentPa = inputResult - Range("C23").Value
                'ONÌe¿ûüÌæ¾(RiÖWtàTO)'
                outputKoronaPaymentPa2 = inputResult
            End If
            
            'ONÌûüÌ\ÌL³Åðªò±±ÜÅª'
            
            
            'ONÌe¿¾Ìæ¾'
            
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 28, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "ONÌe¿¾Ìæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            'ONÌv¾Ìæ¾'
            
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "ONÌv¾Ìæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            'ONÌe¿O¾Ìæ¾'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            'e¿ûüÌ¸­¦'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '¸ÆÎÛÛ¯¿zÌZo'
            On Error Resume Next
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
        
        ElseIf influencePayment = "»Ì¼ûü" Then
        
            Set tbl = Worksheets("Äpoi¾EYj¾îñ(Ú×)").Range("C:AZ")
            
            KoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) 'ONRie¿»Ì¼ûüÌæ¾'
            
            
            'ONÌûüÌ\ÌL³Åðªò±±©ç«'
            
            If KoronaPaymentPa > 0 Then
                
                'ONÌe¿ûüÌæ¾(RiÖWtàTã)'
                On Error Resume Next
                outputKoronaPaymentPa = Application.WorksheetFunction.VLookup(key, tbl, 17, False) - Range("C23").Value
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa = -999999
                    MsgBox "ONÌe¿ûüÌæ¾É¸sµÜµ½B"
                End If
                On Error GoTo 0
                
                'ONÌe¿ûüÌæ¾(RiÖWtàTO)'
                On Error Resume Next
                outputKoronaPaymentPa2 = Application.WorksheetFunction.VLookup(key, tbl, 17, False)
                If Err.Number <> 0 Then
                    outputKoronaPaymentPa2 = -999999
                    MsgBox "ONÌe¿ûüÌæ¾É¸sµÜµ½B"
                End If
                On Error GoTo 0
            
            ElseIf KoronaPaymentPa = 0 Then
            
                inputResult = Application.InputBox("ON»Ì¼ûüª0Å·B" & vbCrLf & "ON»Ì¼ûüðüÍµÄº³¢B")
                'ONÌe¿ûüÌæ¾(RiÖWtàTã)'
                outputKoronaPaymentPa = inputResult - Range("C23").Value
                'ONÌe¿ûüÌæ¾(RiÖWtàTO)'
                outputKoronaPaymentPa2 = inputResult
            
            End If
            
            'ONÌûüÌ\ÌL³Åðªò±±ÜÅª'
            
            
            'ONÌe¿¾Ìæ¾'
            
            On Error Resume Next
            outputKoronaIncomeS = Application.WorksheetFunction.VLookup(key, tbl, 36, False)
            If Err.Number <> 0 Then
                outputKoronaIncomeS = -999999
                MsgBox "ONÌe¿¾Ìæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            'ONÌv¾Ìæ¾'
            
            On Error Resume Next
            outputTotalIncome = Application.WorksheetFunction.VLookup(key, tbl, 50, False)
            If Err.Number <> 0 Then
                outputTotalIncome = -999999
                MsgBox "ONÌv¾Ìæ¾É¸sµÜµ½B"
            End If
            On Error GoTo 0
            
            'ONÌe¿O¾Ìæ¾'
            outputNonKoronaIncome = outputTotalIncome - outputKoronaIncomeS
            
            
            'e¿ûüÌ¸­¦'
            'On Error GoTo Errlabel'
            decreaseRateKorona = 1 - Application.WorksheetFunction.Round(outputKoronaPaymentC / outputKoronaPaymentPa, 5)
            
            
            '¸ÆÎÛÛ¯¿zÌZo'
            outputObjectPrice = annualCost * Application.WorksheetFunction.Round(outputKoronaIncomeS / familyTotalIncome, 5)
        
        
        
        Else
            MsgBox "¢ÑåÌRie¿ûüªIð³êÄ¢Ü¹ñB"
            MsgBox "¸Æ»èð~µÜ·B"
            Exit Sub
        
        End If
        
        
        'eoÍ'

        Range("S7").Value = decreaseRateKorona '¸­¦oÍ'
        Range("R7").Value = outputKoronaPaymentPa 'Rie¿ONûü(RiÖWtàTã)ÌoÍ'
        Range("R9").Value = outputKoronaPaymentPa2 'Rie¿ONûü(RiÖWtàTO)ÌoÍ'
        Range("Q14").Value = outputKoronaIncomeS 'Rie¿ON¾ÌoÍP'
        Range("Q31").Value = outputKoronaIncomeS 'Rie¿ON¾ÌoÍ2'
        Range("R14").Value = outputNonKoronaIncome 'Rie¿OON¾ÌoÍ'
        Range("P34").Value = outputObjectPrice 'U¸ÆÎÛÛ¯¿zÌoÍ'
        
        
        
        '¸ÆzvZÍ±±©ç'
        
        Dim nushiTotalIncome As Long '¢ÑåÌONv¾'
        Dim exemptionRate6 As Variant 'Ri¸Æ¦'
        
        
        nushiTotalIncome = Range("S14").Value
        
        
        
        '¸Æ¦ÌvZ'
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
            outputExeRate.Value = "¸ÆsÂ(<30%)"
            MsgBox "ûü¸­¦ª30ðºñé½ßAU¸ÆªÅ«Ü¹ñB"
        
        ElseIf outputNonKoronaIncome > 4000000 Then
            outputExePrice.Value = outputObjectPrice * 0
            outputExeRate.Value = "¸ÆsÂ(>400)"
            MsgBox "Rie¿O¾ª400~ð´¦é½ßAU¸ÆªÅ«Ü¹ñB"
        
        End If
    End If
End Sub

Attribute VB_Name = "Œ¸–Æ—pŠ“¾Š„ŠzZo"
Sub Œ¸–Æ—pŠ“¾Š„ŠzZo()
    Dim objectPrice25 As Long '2E5†Œ¸–ÆŠzZo—pŠ“¾Š„‡ŒvŠz'
    Dim tbl25 As Range 'v-lookupŠ“¾Š„‡ŒvŠzæ“¾—p”ÍˆÍ'
    Dim key25 As Long  'v-lookupŒŸõ—pƒL['
    Dim sIryou As Long 'ˆã—Ã•ªŠ“¾Š„Šz‚Ìæ“¾'
    Dim sShiennkinn  As Long 'x‰‡‹à•ªŠ“¾Š„Šz‚Ìæ“¾'
    Dim sKaigo As Long '‰îŒì•ªŠ“¾Š„Šz‚Ìæ“¾
    
    Set tbl25 = Worksheets("•Š‰Ûî•ñˆê——").Range("C:DN")
    key25 = Range("B1").Value
    
    'ˆã—Ã•ªŠ“¾Š„Šz‚Ìæ“¾'
    On Error Resume Next
        sIryou = Application.WorksheetFunction.VLookup(key25, tbl25, 34, False)
        If Err.Number <> 0 Then
            sIryou = -999999
            MsgBox "ˆã—Ã•ªŠ“¾Š„Šz‚Ìæ“¾‚É¸”s‚µ‚Ü‚µ‚½B"
        End If
    On Error GoTo 0
    
    'x‰‡‹à•ªŠ“¾Š„Šz‚Ìæ“¾'
    On Error Resume Next
        sShiennkinn = Application.WorksheetFunction.VLookup(key25, tbl25, 62, False)
        If Err.Number <> 0 Then
            sShiennkinn = -999999
            MsgBox "x‰‡‹à•ªŠ“¾Š„Šz‚Ìæ“¾‚É¸”s‚µ‚Ü‚µ‚½B"
        End If
    On Error GoTo 0
    
    '‰îŒì•ªŠ“¾Š„Šz‚Ìæ“¾'
    On Error Resume Next
        sKaigo = Application.WorksheetFunction.VLookup(key25, tbl25, 90, False)
        If Err.Number <> 0 Then
            sKaigo = -999999
            MsgBox "‰îŒì•ªŠ“¾Š„Šz‚Ìæ“¾‚É¸”s‚µ‚Ü‚µ‚½B"
        End If
    On Error GoTo 0
    
    'Š“¾Š„‡ŒvŠz‚Ìæ“¾'
    objectPrice25 = sIryou + sShiennkinn + sKaigo
    
    'Š“¾Š„‡ŒvŠz‚Ìo—Í'
    Range("C56").Value = objectPrice25
    Range("J56").Value = objectPrice25
End Sub

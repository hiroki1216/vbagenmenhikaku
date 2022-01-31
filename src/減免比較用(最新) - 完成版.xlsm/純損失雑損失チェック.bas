Attribute VB_Name = "純損失雑損失チェック"
Sub 純損失雑損失チェック()
    Dim tblJZ As Range '純損失・雑損失の範囲用'
    Dim keyJZ As Long '純損失・雑損失の検索用'
    Dim J As Long '純損失用'
    Dim Z As Long '雑損失用'
    Dim JZ As Long '純損失・雑損失繰返し用'
    Dim last As Long '繰返回数制御用
    
    Set tblJZ = Worksheets("所得・資産情報一覧").Range("C:AV") '純損失・雑損失の範囲'
    last = Range("B3").End(xlDown).Row '繰返回数
    
    For JZ = 3 To last
    
        keyJZ = Cells(JZ, 2).Value '純損失・雑損失の検索用範囲'
        On Error Resume Next
        J = Application.WorksheetFunction.VLookup(keyJZ, tblJZ, 44, False) '純損失の取得'
        On Error Resume Next
        Z = Application.WorksheetFunction.VLookup(keyJZ, tblJZ, 46, False) '雑損失の取得'
        
        '純損失１円以上か判定
        If J > 0 Then
            MsgBox "純損失が１円以上の対象者がいます。６号減免の再計算を行ってください。" & vbCrLf & "A3～A9で赤色のついたセルが純損失対象者です。", vbExclamation
            Cells(JZ, 2).Interior.ColorIndex = 3
            
            outputExeRate.Value = "再計算要(純損失)"
        '雑損失１円以上か判定
        ElseIf Z > 0 Then
            MsgBox "雑損失が１円以上の対象者がいます。６号減免の再計算を行ってください。" & vbCrLf & "A3～A9で黄色のついたセルが雑損失対象者です。", vbExclamation
            Cells(JZ, 2).Interior.ColorIndex = 6
            
            outputExeRate.Value = "再計算要(雑損失)"
        
        End If
    
    Next JZ

    MsgBox "〇減免判定結果" & vbCrLf & vbCrLf & "５号減免率:" & output5.Value & vbCrLf & "２号減免率:" & outputExemptionRate.Value & vbCrLf & "６号減免率:" & outputExeRate.Value & vbCrLf & "６号減免額:" & outputExePrice.Value
    Worksheets("減免判定用").PrintOut
   
End Sub



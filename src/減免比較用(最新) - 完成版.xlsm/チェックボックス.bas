Attribute VB_Name = "チェックボックス"
Sub CheckBox1_Click()
    Dim check_result1 As Object
    Set check_result1 = Range("H7")
    
        If check_result1.Value = True Then
            Cells(7, 10).Value = Cells(13, 3).Value
            Cells(8, 10).Value = Cells(13, 3).Value + 1
            Cells(9, 10).Value = Cells(13, 3).Value + 2
            
            Cells(7, 11).Value = ""
            Cells(8, 11).Value = ""
            Cells(9, 11).Value = ""
            
            Cells(7, 12).Value = Cells(29, 5).Value
            Cells(8, 12).Value = Cells(30, 5).Value
            Cells(9, 12).Value = Cells(31, 5).Value
            
            Cells(7, 13).Value = ""
            Cells(8, 13).Value = ""
            Cells(9, 13).Value = ""
        
        Else
            Cells(7, 10).Value = Cells(29, 3).Value
            Cells(8, 10).Value = Cells(30, 3).Value
            Cells(9, 10).Value = Cells(31, 3).Value
            
            Cells(7, 11).Value = Cells(29, 4).Value
            Cells(8, 11).Value = Cells(30, 4).Value
            Cells(9, 11).Value = Cells(31, 4).Value
            
            Cells(7, 12).Value = Cells(29, 5).Value
            Cells(8, 12).Value = Cells(30, 5).Value
            Cells(9, 12).Value = Cells(31, 5).Value
            
            Cells(7, 13).Value = Cells(29, 6).Value
            Cells(8, 13).Value = Cells(30, 6).Value
            Cells(9, 13).Value = Cells(31, 6).Value
            
        End If
End Sub

Sub CheckBox2_Click()
        Dim check_result2 As Object
        Set check_result2 = Range("H10")
   
        If check_result2.Value = True Then
            Cells(10, 10).Value = Cells(13, 3).Value
            Cells(11, 10).Value = Cells(13, 3).Value + 1
            Cells(12, 10).Value = Cells(13, 3).Value + 2
            
            Cells(10, 11).Value = ""
            Cells(11, 11).Value = ""
            Cells(12, 11).Value = ""
            
            Cells(10, 12).Value = Cells(32, 5).Value
            Cells(11, 12).Value = Cells(33, 5).Value
            Cells(12, 12).Value = Cells(34, 5).Value
            
            Cells(10, 13).Value = ""
            Cells(11, 13).Value = ""
            Cells(12, 13).Value = ""
        Else
            Cells(10, 10).Value = Cells(32, 3).Value
            Cells(11, 10).Value = Cells(33, 3).Value
            Cells(12, 10).Value = Cells(34, 3).Value
            
            Cells(10, 11).Value = Cells(32, 4).Value
            Cells(11, 11).Value = Cells(33, 4).Value
            Cells(12, 11).Value = Cells(34, 4).Value
            
            Cells(10, 12).Value = Cells(32, 5).Value
            Cells(11, 12).Value = Cells(33, 5).Value
            Cells(12, 12).Value = Cells(34, 5).Value
            
            Cells(10, 13).Value = Cells(32, 6).Value
            Cells(11, 13).Value = Cells(33, 6).Value
            Cells(12, 13).Value = Cells(34, 6).Value
        End If
End Sub
Sub CheckBox3_Click()
        Dim check_result3 As Object
        Set check_result3 = Range("H13")
        If check_result3.Value = True Then
            Cells(13, 10).Value = Cells(13, 3).Value
            Cells(14, 10).Value = Cells(13, 3).Value + 1
            Cells(15, 10).Value = Cells(13, 3).Value + 2
            
            Cells(13, 11).Value = ""
            Cells(14, 11).Value = ""
            Cells(15, 11).Value = ""
            
            Cells(13, 12).Value = Cells(35, 5).Value
            Cells(14, 12).Value = Cells(36, 5).Value
            Cells(15, 12).Value = Cells(37, 5).Value
            
            Cells(13, 13).Value = ""
            Cells(14, 13).Value = ""
            Cells(15, 13).Value = ""
        Else
            Cells(13, 10).Value = Cells(35, 3).Value
            Cells(14, 10).Value = Cells(36, 3).Value
            Cells(15, 10).Value = Cells(37, 3).Value
            
            Cells(13, 11).Value = Cells(35, 4).Value
            Cells(14, 11).Value = Cells(36, 4).Value
            Cells(15, 11).Value = Cells(37, 4).Value
            
            Cells(13, 12).Value = Cells(35, 5).Value
            Cells(14, 12).Value = Cells(36, 5).Value
            Cells(15, 12).Value = Cells(37, 5).Value
            
            Cells(13, 13).Value = Cells(35, 6).Value
            Cells(14, 13).Value = Cells(36, 6).Value
            Cells(15, 13).Value = Cells(37, 6).Value
        End If
End Sub

Sub CheckBox4_Click()
        Dim check_result4 As Object
        Set check_result4 = Range("H16")
        If check_result4.Value = True Then
            Cells(16, 10).Value = Cells(13, 3).Value
            Cells(17, 10).Value = Cells(13, 3).Value + 1
            Cells(18, 10).Value = Cells(13, 3).Value + 2
            
            Cells(16, 11).Value = ""
            Cells(17, 11).Value = ""
            Cells(18, 11).Value = ""
            
            Cells(16, 12).Value = Cells(38, 5).Value
            Cells(17, 12).Value = Cells(39, 5).Value
            Cells(18, 12).Value = Cells(40, 5).Value
            
            Cells(16, 13).Value = ""
            Cells(17, 13).Value = ""
            Cells(18, 13).Value = ""
        Else
            Cells(16, 10).Value = Cells(38, 3).Value
            Cells(17, 10).Value = Cells(39, 3).Value
            Cells(18, 10).Value = Cells(40, 3).Value
            
            Cells(16, 11).Value = Cells(38, 4).Value
            Cells(17, 11).Value = Cells(39, 4).Value
            Cells(18, 11).Value = Cells(40, 4).Value
            
            Cells(16, 12).Value = Cells(38, 5).Value
            Cells(17, 12).Value = Cells(39, 5).Value
            Cells(18, 12).Value = Cells(40, 5).Value
            
            Cells(16, 13).Value = Cells(38, 6).Value
            Cells(17, 13).Value = Cells(39, 6).Value
            Cells(18, 13).Value = Cells(40, 6).Value
        End If
End Sub
Sub CheckBox5_Click()
        Dim check_result5 As Object
        Set check_result5 = Range("H19")
    
        If check_result5.Value = True Then
            Cells(19, 10).Value = Cells(13, 3).Value
            Cells(20, 10).Value = Cells(13, 3).Value + 1
            Cells(21, 10).Value = Cells(13, 3).Value + 2
            
            Cells(19, 11).Value = ""
            Cells(20, 11).Value = ""
            Cells(21, 11).Value = ""
            
            Cells(19, 12).Value = Cells(41, 5).Value
            Cells(20, 12).Value = Cells(42, 5).Value
            Cells(21, 12).Value = Cells(43, 5).Value
            
            Cells(19, 13).Value = ""
            Cells(20, 13).Value = ""
            Cells(21, 13).Value = ""
        
        Else
            Cells(19, 10).Value = Cells(41, 3).Value
            Cells(20, 10).Value = Cells(42, 3).Value
            Cells(21, 10).Value = Cells(43, 3).Value
            
            Cells(19, 11).Value = Cells(41, 4).Value
            Cells(20, 11).Value = Cells(42, 4).Value
            Cells(21, 11).Value = Cells(43, 4).Value
            
            Cells(19, 12).Value = Cells(41, 5).Value
            Cells(20, 12).Value = Cells(42, 5).Value
            Cells(21, 12).Value = Cells(43, 5).Value
            
            Cells(19, 13).Value = Cells(41, 6).Value
            Cells(20, 13).Value = Cells(42, 6).Value
            Cells(21, 13).Value = Cells(43, 6).Value
            
        End If
End Sub

Sub CheckBox6_Click()
        Dim check_result6 As Object
        Set check_result6 = Range("H22")
        If check_result6.Value = True Then
            Cells(22, 10).Value = Cells(13, 3).Value
            Cells(23, 10).Value = Cells(13, 3).Value + 1
            Cells(24, 10).Value = Cells(13, 3).Value + 2
            
            Cells(22, 11).Value = ""
            Cells(23, 11).Value = ""
            Cells(24, 11).Value = ""
            
            Cells(22, 12).Value = Cells(44, 5).Value
            Cells(23, 12).Value = Cells(45, 5).Value
            Cells(24, 12).Value = Cells(46, 5).Value
            
            Cells(22, 13).Value = ""
            Cells(23, 13).Value = ""
            Cells(24, 13).Value = ""
        
        Else
            Cells(22, 10).Value = Cells(44, 3).Value
            Cells(23, 10).Value = Cells(45, 3).Value
            Cells(24, 10).Value = Cells(46, 3).Value
            
            Cells(22, 11).Value = Cells(44, 4).Value
            Cells(23, 11).Value = Cells(45, 4).Value
            Cells(24, 11).Value = Cells(46, 4).Value
            
            Cells(22, 12).Value = Cells(44, 5).Value
            Cells(23, 12).Value = Cells(45, 5).Value
            Cells(24, 12).Value = Cells(46, 5).Value
            
            Cells(22, 13).Value = Cells(44, 6).Value
            Cells(23, 13).Value = Cells(45, 6).Value
            Cells(24, 13).Value = Cells(46, 6).Value
            
        End If
End Sub

Sub CheckBox7_Click()
        Dim check_result7 As Object
        Set check_result7 = Range("H25")
        If check_result7.Value = True Then
            Cells(25, 10).Value = Cells(13, 3).Value
            Cells(26, 10).Value = Cells(13, 3).Value + 1
            Cells(27, 10).Value = Cells(13, 3).Value + 2
            
            Cells(25, 11).Value = ""
            Cells(26, 11).Value = ""
            Cells(27, 11).Value = ""
            
            Cells(25, 12).Value = Cells(47, 5).Value
            Cells(26, 12).Value = Cells(48, 5).Value
            Cells(27, 12).Value = Cells(49, 5).Value
            
            Cells(25, 13).Value = ""
            Cells(26, 13).Value = ""
            Cells(27, 13).Value = ""
        Else
            Cells(25, 10).Value = Cells(47, 3).Value
            Cells(26, 10).Value = Cells(48, 3).Value
            Cells(27, 10).Value = Cells(49, 3).Value
            
            Cells(25, 11).Value = Cells(47, 4).Value
            Cells(26, 11).Value = Cells(48, 4).Value
            Cells(27, 11).Value = Cells(49, 4).Value
            
            Cells(25, 12).Value = Cells(47, 5).Value
            Cells(26, 12).Value = Cells(48, 5).Value
            Cells(27, 12).Value = Cells(49, 5).Value
            
            Cells(25, 13).Value = Cells(47, 6).Value
            Cells(26, 13).Value = Cells(48, 6).Value
            Cells(27, 13).Value = Cells(49, 6).Value
        End If
End Sub

Private Sub Command1_Click()
    If Len(Text1.Text) <> 8 Then
        MsgBox "输入的用户名不是8个字符", , "重新输入"
    End If
    
    If Len(Text2.Text) <> 6 Then
        MsgBox "你输入的密码中不是6个字符", , "重新输入"
    End If
    
    If IsNumeric(Text2.Text) = False Then
        MsgBox "你输入的密码中含有非数字的字符", , "重新输入"
    End If
    
    Dim a As String
    
    If Len(Text1.Text) = 8 And Len(Text2.Text) = 6 And IsNumeric(Text2.Text) = True Then
        a = InputBox("转账金额", "输入转账金额")
        Print a
    End If
End Sub
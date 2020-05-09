Select Case  strC
Case "a" To "z"，"A" To "Z"
    Print  strC + "是字母字符"
Case "0" To "9"
    Print  strC + "是数字字符"
Case Else
    Print  strC + "其他字符"
End Select



Dim x As Integer
x = 1
Select Case x
Case 1
    Print "x is 1"
Case 2
    Print "x is 2"
Case 3
    Print "x is 3"
Case Else
    Print "else"
End Select
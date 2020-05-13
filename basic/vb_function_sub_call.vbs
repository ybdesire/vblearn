
Private Function findMax(a() As Integer) As Integer
    For i = 0 To UBound(a) Step 1
        Print a(i)
    Next i
    findMax = a(0)
End Function



Private Function add(x%, y As Integer) As Integer
    add = x + y
End Function



Private Sub test(x%)
    Print x
End Sub



Private Sub Form_Click()
    Call test(666)
    Print add(5, 6)
    Dim m As Integer
    Dim a(5) As Integer
    a(0) = 5
    a(1) = 4
    a(2) = 3
    a(3) = 2
    a(4) = 1
    a(5) = 0
    
    m = findMax(a)
    Print "find " & m
    
End Sub









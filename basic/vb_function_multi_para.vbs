
Private Function addd(x%, y%, z%)
    addd = x + y + z + 1
End Function


Private Sub Form_Click()
    Dim x As Integer
    x = addd(1, 2, 3)
    Print x
    
End Sub

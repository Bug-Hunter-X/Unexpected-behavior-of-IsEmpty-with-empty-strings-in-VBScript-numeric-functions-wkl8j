Function f(a, b)
  If Len(a) = 0 Then Exit Function
  If Len(b) = 0 Then Exit Function
  f = a + b
End Function

MsgBox f(1, "")
'Now it correctly returns an error or exits the function.
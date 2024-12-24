Function f(a)
  If a = 1 Then
    f = 1
  ElseIf a = 2 Then
    f = 2
  Else
    f = f(a - 1) + f(a - 2)
  End If
end Function
MsgBox f(5)
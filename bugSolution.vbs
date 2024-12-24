Function f_iterative(a)
  Dim fib(a)
  fib(1) = 1
  fib(2) = 2

  If a <= 0 Then
    f_iterative = 0
  ElseIf a = 1 Then
    f_iterative = 1
  ElseIf a = 2 Then
    f_iterative = 2
  Else
    For i = 3 To a
      fib(i) = fib(i - 1) + fib(i - 2)
    Next
    f_iterative = fib(a)
  End If
End Function
MsgBox f_iterative(5)
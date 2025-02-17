Function f(a)
  If IsEmpty(a) Then Exit Function
  If IsNumeric(a) Then
    f = a * 2
  ElseIf IsDate(a) Then
    f = Day(a)
  Else
    f = Len(a)
  End If
End Function

MsgBox f(10) 'Output: 20
MsgBox f(#1/1/2024#) 'Output: 1
MsgBox f("hello") 'Output: 5
MsgBox f(Empty) 'Output: Empty
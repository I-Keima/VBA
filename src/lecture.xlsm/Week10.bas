Attribute VB_Name = "Week10"

Option Explicit
Option Base 1

Sub week10_1(Optional h As Double = 0.1, Optional n As Integer = 1)
  Call worksheets("sheet1").Activate
  
  Dim ans as variant
  Dim t, col1, col2, col3, t_min, t_max as double
  t_max = 2.4
  t_min = 0.0
  ans = create_matrix(Int((t_max - t_min) / h) + 3, 4)
  ans(1) = Array("t", "数値解", "解析解", "誤差")
  
  Dim i as integer
  ans(2) = Array(0, 0, 0, 0)
  For i = 3 To UBound(ans)
    t = (i - 2) * h
    col1 = ans(i - 1)(2) + h * (1 - ans(i - 1)(2) ^ 2)
    col2 = week10_1_x(t)
    col3 = col1 - col2
    ans(i) = Array(t, col1, col2, col3)
  Next i

  Call print_matrix(1, 1 + 4 * (n - 1), ans)
End Sub
  
  '実習課題１をやった後実習課題２のために関数化したため、
  'week10_1プロシージャを実行するためのプロシージャが必要になった。
Sub week10_1_do()
  Call week10_1()
End Sub

Function week10_1_x(t)
  week10_1_x = (Exp(t) - Exp(-t)) / (Exp(t) + Exp(-t))
End Function

Sub week10_2()
  Dim i As Integer
  Dim h_list As Variant
  h_list = Array(0.05, 0.1, 0.2, 0.4)
  For i = 1 To UBound(h_list)
    Call week10_1(CDbl(h_list(i)), i + 1)
  Next i
End Sub


Sub week10_3()
  Call worksheets("sheet2").Activate

  Dim ans as variant
  Dim t, col1, col2, col3, t_min, t_max, x_hat, h as double
  h = 0.1
  t_max = 2.4
  t_min = 0.0
  ans = create_matrix(Int((t_max - t_min) / h) + 3, 4)
  ans(1) = Array("t", "数値解", "解析解", "誤差")

  Dim i as integer
  ans(2) = Array(0, 0, 0, 0)
  For i = 3 To UBound(ans)
    t = (i - 2) * h
    x_hat = ans(i - 1)(2) + h * (1 - ans(i - 1)(2) ^ 2)
    col1 = ans(i - 1)(2) + h * (1 - (ans(i - 1)(2)) ^ 2 + 1 - x_hat ^ 2) / 2
    col2 = week10_1_x(t)
    col3 = col1 - col2
    ans(i) = Array(t, col1, col2, col3)
  Next i

  Call print_matrix(1, 1, ans)
End Sub


Sub week10_4()
  Call worksheets("sheet3").Activate

  Dim ans as variant
  Dim t, col1, col2, col3, t_min, t_max, x_hat, h as double
  Dim f_1, f_2, f_3, f_4, t_k, x_k as double
  h = 0.1
  t_max = 2.4
  t_min = 0.0
  ans = create_matrix(Int((t_max - t_min) / h) + 3, 4)
  ans(1) = Array("t", "数値解", "解析解", "誤差")
‘h
  Dim i as integer
  ans(2) = Array(0, 0, 0, 0)
  For i = 3 To UBound(ans)
    t = (i - 2) * h
    t_k = (i - 3) * h
    x_k = ans(i - 1)(2)
    f_1 = 1 - x_k ^ 2
    f_2 = 1 - (x_k + f_1 * h / 2) ^ 2
    f_3 = 1 - (x_k + f_2 * h / 2) ^ 2
    f_4 = 1 - (x_k + h * f_3)
    col1 = x_k + h * (f_1 + 2 * f_2 + 2 * f_3 + f_4) / 6
    col2 = week10_1_x(t)
    col3 = col1 - col2
    ans(i) = Array(t, col1, col2, col3)
  Next i

  Call print_matrix(1, 1, ans)
End Sub
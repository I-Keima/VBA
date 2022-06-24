Attribute VB_Name = "Week11"

Option Explicit
Option Base 1

Function f_x(x)
  f_x = Exp(x)
  '発展課題(week11_4)を実行する時は上行をコメントアウトし、下行を用いる
  'f_x = Exp(-x ^ 2 / 2)
  '(1)用
  'f_x = x ^ 3 / 5 - 3 * x ^ 2 / 5 - 2 * x / 5 + 3
End Function


Function midpoint_rule(h, Optional x_min As double = 0, Optional x_max As double = 1)
  Dim large_n As Integer 
  Dim integration_value, x_k, x_k1 As Double

  large_n = Int((x_max - x_min) / h)
  
  Dim k As Integer
  integration_value = 0
  x_k1 = x_min
  For k = 0 To large_n - 1
    x_k = x_k1
    x_k1 = x_min + h * (k + 1)
    integration_value = integration_value + h * f_x((x_k + x_k1)/2)
  Next k

  midpoint_rule = integration_value
End Function


Sub week11_1()
  Dim h_list, ans As Variant
  h_list = Array(1/2, 1/4, 1/8)
  ans = h_list
  
  Dim i As Integer
  For i = 1 To UBound(h_list)
    ans(i) = midpoint_rule(h_list(i))
  Next i 

  Dim head As Variant
  head = Array("中点則", "h=1/2", "h=1/4", "h=1/8")
  Call print_array(1, 1, head)
  Call print_array(2, 2, ans)
End Sub


Function trapezoidal_rule(h, Optional x_min As double = 0, Optional x_max As double = 1)
  Dim large_n As Integer
  Dim integration_value, x_k, x_k1 As Double

  large_n = Int((x_max - x_min) / h)
  
  Dim k As Integer
  integration_value = 0
  x_k1 = x_min
  For k = 0 To large_n - 1
    x_k = x_k1
    x_k1 = x_min + h * (k + 1)
    integration_value = integration_value + h * (f_x(x_k) + f_x(x_k1)) / 2
  Next k

  trapezoidal_rule = integration_value
End Function


Sub week11_2()
  Dim h_list, ans As Variant
  h_list = Array(1/2, 1/4, 1/8)
  ans = h_list

  Dim i As Integer
  For i = 1 To UBound(h_list)
    ans(i) = trapezoidal_rule(h_list(i))
  Next i

  Dim head As Variant
  head = Array("台形則", "h=1/2", "h=1/4", "h=1/8")
  Call print_array(4, 1, head)
  Call print_array(5, 2, ans)
End Sub


Function simpsons_rule(h, Optional x_min As Double = 0, Optional x_max As Double = 1)
  Dim m_f_2h, t_f_2h As Double

  m_f_2h = midpoint_rule(2 * h)
  t_f_2h = trapezoidal_rule(2 * h)

  simpsons_rule = (2 * m_f_2h + t_f_2h) / 3
End Function


Sub week11_3()
  Dim h_list, ans As Variant
  h_list = Array(1/2, 1/4, 1/8)
  ans = h_list

  Dim i As Integer
  For i = 1 To UBound(h_list)
    ans(i) = simpsons_rule(h_list(i))
  Next i

  Dim head As Variant
  head = Array("シンプソン則", "h=1/2", "h=1/4", "h=1/8")
  Call print_array(7, 1, head)
  Call print_array(8, 2, ans)
End Sub


Sub week11_4()
  Dim epsilon, h, abs_f, x_min, x, integration_value, error, true_value As Double
  Dim large_n, i, k As Integer
  Dim ans As Variant
  ans = create_matrix(4,3)
  ans(1) = Array("", "近似値", "真値", "誤差")
  epsilon = 0.002
  h = 0.5
  i = 0
  x_min = 0
  true_value = 1.2533141

  abs_f = 1
  While abs_f >= epsilon
    i = i + 1
    x = x_min + h * i
    abs_f = Abs(f_x(x)) + Abs(f_x(x + h))
  Wend
  large_n = i


  integration_value = midpoint_rule(h, CDbl(x_min), CDbl(large_n))
  error = Abs(integration_value - true_value)
  ans(2) = Array("中点則", integration_value, true_value, error)

  integration_value = trapezoidal_rule(h, CDbl(x_min), CDbl(large_n))
  error = Abs(integration_value - true_value)
  ans(3) = Array("台形則", integration_value, true_value, error)

  integration_value = simpsons_rule(h, CDbl(x_min), CDbl(large_n))
  error = Abs(integration_value - true_value)
  ans(4) = Array("シンプソン則", integration_value, true_value, error)

  Call print_matrix(10, 1, ans)
End Sub
  
Sub week11_kiso()
  Dim h As double: h = 1/4
  Cells(16,1) = midpoint_rule(h, 0, 4)
  Cells(16, 2) = trapezoidal_rule(h, 0, 4)
  Cells(16, 3) = simpsons_rule(h, 0, 4)
End Sub

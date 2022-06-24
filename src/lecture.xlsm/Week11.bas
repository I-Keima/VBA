Attribute VB_Name = "Week11"

Option Explicit
Option Base 1

Function f_x(x)
  f_x = Exp(x)
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
  head = Array("区分求積法", "h=1/2", "h=1/4", "h=1/8")
  Call print_matrix(head, 1, 1)
  Call print_matrix(ans, 2, 2)
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
  Call print_matrix(head, 4, 1)
  Call print_matrix(ans, 5, 2)
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
  Call print_matrix(head, 7, 1)
  Call print_matrix(ans, 8, 2)
End Sub


Attribute VB_Name = "Week9"

Option Explicit
Option Base 1

'実習課題１におけるf_1を計算する関数
Function week9_1_f_1(a, h)
  week9_1_f_1 = (Sin(a + h) - Sin(a)) / h
End Function

'実習課題１
Sub week9_1()
  Call worksheets("sheet1").Activate

  Dim ans as variant
  Dim n as integer
  Dim col1, col2, col3, a, h, diff_f as double
  ans = create_matrix(51, 4)
  ans(1) = Array("n", "前進差分", "打ち切り誤差", "f1(a,h)-f1(a,2h)")
  diff_f = Cos(0.3 * 4 * Atn(1))

  Dim i as integer
  a = 0.3 * 4 * Atn(1) 
  For i = 2 To UBound(ans)
    n = i - 1
    h = 2 ^ (-n)
    col1 = week9_1_f_1(a, h)
    col2 = col1 - diff_f
    col3 = col1 - week9_1_f_1(a, 2 * h)
    ans(i) = Array(n, col1, col2, col3)
  Next i

  Call print_matrix(1,1,ans)
End Sub

Function week9_2_f_3(a, h)
  week9_2_f_3 = (Sin(a + h) - Sin(a - h)) / (2 * h)
End Function


Sub week9_2()
  Call worksheets("sheet1").Activate

  Dim ans as variant
  Dim n as integer
  Dim col1, col2, col3, a, h, diff_f as double
  ans = create_matrix(51, 4)
  ans(1) = Array("n", "前進差分", "打ち切り誤差", "f1(a,h)-f1(a,2h)")
  diff_f = Cos(0.3 * 4 * Atn(1))

  Dim i as integer
  a = 0.3 * 4 * Atn(1) 
  For i = 2 To UBound(ans)
    n = i - 1
    h = 2 ^ (-n)
    col1 = week9_2_f_3(a, h)
    col2 = col1 - diff_f
    col3 = col1 - week9_2_f_3(a, 2 * h)
    ans(i) = Array(n, col1, col2, col3)
  Next i

  Call print_matrix(1,6,ans)
End Sub


Function week9_4_f(x)
  week9_4_f = 3 - 4 / (1 + Exp(x))
End Function

Sub week9_4()
  Dim diff_f, x_0, x_1, precision, h as Double
  h = 2 ^ (-50)
  x_0 = 1.7
  precision = 10 ^ (-5)

  x_1 = x_0
  While Abs(week9_4_f(x_1)) >= precision
    x_0 = x_1  
  '中心差分による微分係数
    diff_f = (week9_4_f(x_0 + h) - week9_4_f(x_0 - h)) / (2 * h)
    x_1 = x_0 - week9_4_f(x_0) / diff_f
  Wend

  Call MsgBox("解：" & x_1 & "、f(x)=" & week9_4_f(x_1))
End Sub


Function week9_5_f_1(a, h)
  week9_5_f_1 = (week9_1_f_1(a, h) - (1/2) * week9_1_f_1(a, 2 * h)) / (1 - (1/2))
End Function


Sub week9_5()
  Call worksheets("sheet1").Activate

  Dim ans as variant
  Dim n as integer
  Dim col1, col2, col3, a, h, diff_f as double
  ans = create_matrix(51, 4)
  ans(1) = Array("n", "前進差分", "打ち切り誤差", "f1(a,h)-f1(a,2h)")
  diff_f = Cos(0.3 * 4 * Atn(1))

  Dim i as integer
  a = 0.3 * 4 * Atn(1) 
  For i = 2 To UBound(ans)
    n = i - 1
    h = 2 ^ (-n)
    col1 = week9_5_f_1(a, h)
    col2 = col1 - diff_f
    col3 = col1 - week9_5_f_1(a, 2 * h)
    ans(i) = Array(n, col1, col2, col3)
  Next i

  Call print_matrix(1,12,ans)
End Sub



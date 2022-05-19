Attribute VB_Name = "Week5"

Option Explicit
Option Base 1 '配列の添え字の最小値を1に設定

Sub Kadai17()
  Dim tmp, i As Integer
  tmp = 0
  For i = 20 To 100
    tmp = tmp + i
  Next i
  MsgBox tmp
  Dim m, n As Integer
  m = 0 + 0 + 4
  n = 74 * (2 + 1)
  tmp = 0
  For i = m To n
    tmp = tmp + i
  Next i
  MsgBox tmp
End Sub

Sub Kadai18
  Dim i, j As Integer, arr As Variant: arr = Array(2)
  Dim is_prime As Boolean 
  Dim s As Integer: s = 2
  For i = 3 To 100
    is_prime = True
    For j = 1 To UBound(arr)
      If i Mod arr(j) = 0 Then
        is_prime = false
      End If
    Next j
    If is_prime Then
      ReDim Preserve arr(UBound(arr) + 1)
      arr(UBound(arr)) = i
      s = s + i
    End If
  Next i
  Call printVec(1, 2, arr)
  MsgBox s
  s = 0
  For i = 5 To 9
    s = s + arr(i)
  Next i
  MsgBox s
End Sub

Sub Kadai19()
  Dim n As Integer, a() As Double
  n = 11
  ReDim a(n) as Double

  Dim i As Integer
  For i = 1 To n
    a(i) = i
  Next i

  For i = 1 To n
    Cells(i+4, 3) = a(i)
  Next i
End Sub

Function multiply_scalar_vector(s1 As Double, m1 As Variant) As Variant
  Dim i As Integer
  Dim ans As Variant: ans = m1
  For i = 1 To UBound(m1)
    ans(i) = s1 * m1(i)
  Next i
  multiply_scalar_vector = ans
End Function
  
Function addition_vector(v1 As Variant, v2 As Variant) As Variant
  Dim i, n As Integer: n = UBound(v1)
  Dim ans As Variant: ans = v1
  For i = 1 To n
    ans(i) = v1(i) + v2(i)
  Next i
  addition_vector = ans
End Function
    

Sub Kadai20()
  Dim i, n As Integer
  n = 74
  Dim a, b As Variant: ReDim a(n): ReDim b(n)
  For i = 1 To n
    a(i) = i
    b(i) = i + 1
  Next i
  Dim ans As Variant
  Dim k, h As Integer
  k = 3: h = 4
  ans = addition_vector(multiply_scalar_vector(CDbl(k), a), multiply_scalar_vector(CDbl(h), b))
  Range("D1:D" & n) = ans

  a = Range("E1:E6")
  b = Range("F1:F6")
  k = 2: h = 1
  ans = addition_vector(multiply_scalar_vector(CDbl(k), a), multiply_scalar_vector(CDbl(h), b))
  Range("G1:G6") = ans
End Sub


Function get_last_row(row As Integer, col As Integer) As Integer
  Dim i As Integer, last As Boolean: last = False
  i = row
  While not last
    If Cells(i, col) = "" Then
      last = True
    Else
      i = i + 1
    End If
  Wend
  get_last_row = i
End Function

Sub Kadai21()
  Dim v1, v2 As Variant
  v1 = Range("E1:E" & get_last_row(1, 5))
  v2 = Range("F1:F" & get_last_row(1, 6))
  
  Dim i As Integer, dot_p As Double
  dot_p = 0
  For i = 1 To Ubound(v1)
    dot_p = dot_p + v1(i) * v2(i)
  Next i
  MsgBox dot_p
End Sub


Function t_matrix(m1 As Variant) As Variant
  Dim ans As Variant: ReDim ans(UBound(m1,2), UBound(m1,1)) As Double
  Dim i, j As Integer
  For i = 1 To UBound(m1, 1)
    For j = 1 To UBound(m1, 2)
      ans(j, i) = m1(i, j)
    Next j
  Next i
  t_matrix = ans
End Function


Sub Kadai22()
  Call WorkSheets("Sheet2").Activate
  Call printMatrix(1, 5, t_matrix(getMatrix(1, 1, 3, 3)))
End Sub

Function multiply_scalar_matrix(s1 As Double, m1 As Variant) As Variant
  Dim i, j As Integer, ans As Variant
  ans = m1
  For i = 1 To UBound(m1, 1)
    For j = 1 To UBound(m1, 2)
      ans(i, j) = s1 * m1(i, j)
    Next j
  Next i
  multiply_scalar_matrix = ans
End Function

Function addition_matrix(m1 As Variant, m2 As Variant) As Variant
  Dim i, j As Integer
  Dim ans As Variant
  ans = m1
  For i = 1 To UBound(m1, 1)
    For j = 1 To UBound(m1, 2)
      ans(i, j) = m1(i, j) + m2(i, j)
    Next j
  Next i
  addition_matrix = ans
End Function

Function cross_matrix(m1 As Variant, m2 As Variant) As Variant
	Dim i As Integer, j as integer, k as integer, l as integer, m as integer, n as integer
	l = UBound(m1, 1): m = UBound(m1, 2): n = UBound(m2, 2)
	Dim arr() As Double: ReDim arr(l, n)
	For i = 1 To l
		For j = 1 To n
			For k = 1 To m
				arr(i, j) = arr(i, j) + m1(i, k) * m2(k, j)
			Next k
		Next j
	Next i
	cross_matrix = arr
End Function


Sub Kadai23()
  Dim a, b As Variant
  Dim ans As Variant
  
  a = getMatrix(1, 1, 3, 3)
  b = getMatrix(5, 1, 3, 3)
  
  ans = addition_matrix(multiply_scalar_matrix(3, a), multiply_scalar_matrix(5, b))
  Call printMatrix(11, 1, ans)

  Call printMatrix(17, 1, cross_matrix(a, b))
  Call printMatrix(17, 5, cross_matrix(a, b))

  Call printMatrix(23, 1, cross_matrix(t_matrix(a), a))
  Call printMatrix(23, 5, cross_matrix(a, t_matrix(a)))
End Sub

Sub Kadai24()
  'これまでに実装した通りである
  '転置：t_matrix(m1)
  '和：addition_matrix(m1, m2)
  '積：cross_matrix(m1, m2)
End Sub

Sub kadai1()
  Dim i, n, row, col As Integer
  Dim s_x, s_x2, ans As Double

  row = 1: col = 1
  n = get_last_row(CInt(row), CInt(col))
  Dim y_vec() As Double: ReDim y_vec(n)
  For i = 1 To n
    y_vec(i) = Cells(row + i - 1, col)
    s_x = s_x + y_vec(i)
    s_x2 = s_x2 + y_vec(i) ^ 2
  Next i
  ans = s_x2 - s_x ^ 2 / n
  Cells(row, col + 1) = "課題1で求めた平方和"
  Cells(row + 1, col + 1) = ans
End Sub

Sub kadai2()
  Dim i, j, n, row, col As Integer
  Dim ans As Double
  
  row = 1: col = 1
  n = get_last_row(CInt(row), CInt(col))
  Dim arr() As Double: ReDim arr(n)
  Dim y_vec() As Double: ReDim y_vec(n)
  Dim n_matrix() As Double: ReDim n_matrix(n, n)

  For i = 1 To n
    y_vec(i) = Cells(row + i - 1, col)
    For j = 1 To n
      If i = j Then
        n_matrix(i, j) = 1 - 1 / n
      Else
        n_matrix(i, j) = 0 - 1 / n
      End If
    Next 
  Next i

  For j = 1 To n
    For i =  1 To n
      arr(j) = arr(j) + y_vec(i) * n_matrix(i, j)
    Next i
  Next j

  For i = 1 To n
    ans = ans + arr(i) * y_vec(i)
  Next i

  Cells(row + 3, col + 1) = "課題2で求めた平方和"
  Cells(row + 4, col + 1) = ans
End Sub

    


Attribute VB_Name = "Common"

Option Explicit
Option Base 1 '配列の添え字の最小値を1に設定


Function read_matrix

Function matrix_plus(m1, m2)
  Dim ans As Variant: ans = m1
  For i = 1 To UBound(m1)
    For j = 1 To Uboud(m1(1))
      ans(i)(j) = m1(i)(j) + m2(i)(j)
    Next j
  Next i
  matrixPlus = ans
End Function

Function forwardElimination_matrix(matrix_a As Variant, matrix_b As Variant) As Variant
  ' 小数の計算時の誤差を防ぐために最も値の絶対値が大きいピボットを選択し行の入れ替えを行う
  Dim i As Integer, k As Integer, j As Integer, n As Integer, pivot As Double
  n = UBound(matrix_a, 1)
  Dim a_old, a_new, b_old, b_new As Variant
  a_old = matrix_a
  b_old = matrix_b
  For k = 1 To n
    pivot = Abs(a_old(k, k))
    For i = k + 1 To n
      For j = i To n
        matrix_a(i, j) = matrix_a(i, j) - matrix_a(i, k) * matrix_a(k, j) / matrix_a(k, k)
        b(i) = b(i) - matrix_a(i, k) * b(k) / matrix_a(k, k)
      Next j
    Next i
  Next k
	' 返り値をVariant型にすることで配列の中に配列をいれて、二つの要素を返すことが出来る
  forwardElimination = Array(matrix_a, b)
End Function
Attribute VB_Name = "Week4"

Option Explicit
Option Base 1 '配列の添え字の最小値を1に設定

Function forwardElimination(matrix_a As Variant, b As Variant) As Variant
  Dim i As Integer, k As Integer, j As Integer, n As Integer
  n = UBound(matrix_a, 1)
  Dim a_new, a_old, b_new, b_old As Variant
  a_new = matrix_a: b_new = b
  For k = 1 To n
    a_old = a_new: b_old = b_new
    For i = k + 1 To n
      For j = k To n
        a_new(i, j) = a_old(i, j) - a_old(i, k) * a_old(k, j) / a_old(k, k)
        b_new(i) = b_old(i) - a_old(i, k) * b_old(k) / a_old(k, k) 
      Next j
    Next i
  Next k
  ' 返り値をVariant型にすることで配列の中に配列をいれて、二つの要素を返すことが出来る
  forwardElimination = Array(a_new, b_new)
End Function

Sub kadai12()
  Dim arr As Variant
  arr = forwardElimination(getMatrix(1,1,3,3), createRandNdVec(3))
  Call printMatrix(1, 5, arr(1))
End Sub

Function backwardSubstitutution(matrix_u As Variant, y as Variant) As Variant
  Dim i As Integer, j As Integer, x() As Double, n As Integer, s As Double
  n = UBound(matrix_u, 1)
  ReDim x(n)
  x(n) = y(n) / matrix_u(n, n)
  For i = n - 1 To 1 Step -1
    s = 0
    For j = i + 1 To n
      s = s + matrix_u(i, j) * x(j)
    Next j
    x(i) = (y(i) - s) / matrix_u(i, i)
  Next i
  backwardSubstitutution = x
End Function

Sub kadai13()
  Dim matrix_u As Variant, e_n As Variant, ans As Variant
  matrix_u = createOnTriangleMatrix(21)
  e_n = create1NdVec(21)
  ans = backwardSubstitutution(matrix_u, e_n)
  Call printVec(2, 1, ans)
End Sub

Function gaussElimination(matrix_a As Variant, b As Variant) As Variant
  Dim j As integer, x As Variant, n As Integer
  n = UBound(matrix_a, 1)
  Dim ans As Variant
  ans = forwardElimination(matrix_a, b)
  x = backwardSubstitutution(ans(1), ans(2))
  gaussElimination = x
End Function

Sub kadai14()
  Dim matrix_u As Variant, e_n As Variant, ans As Variant
  matrix_u = createOnTriangleMatrix(5)
  e_n = create1NdVec(5)
  ans = gaussElimination(matrix_u, e_n)
  Call printVec(2, 3, ans)
  Call printVec(2, 5, e_n)
  Call printVec(2, 6, matrixVectorProduct(matrix_u, ans))
End Sub

Function identityMatrix(n As Integer) As Double()
  Dim i As Integer, j As Integer, arr() As Double: ReDim arr(n, n)
  For i = 1 To n
    For j = 1 To n
      if i = j Then
        arr(i, j) = 1
      Else
        arr(i, j) = 0
      End If
    Next j
  Next i
  identityMatrix = arr
End Function

Function gaussInverseMatrix(matrix_a As Variant) As Variant
  Dim i As Integer, k As Integer, arr As Variant, l As Integer, n As Integer, x As Variant, inverse As Variant
  inverse = matrix_a
  n = UBound(matrix_a, 1)
  Dim delta As Variant: delta = Array(): ReDim delta(n)
  For i = 1 To n
    For k = 1 To n
      If k = i Then
        delta(k) = 1
      Else
        delta(k) = 0
      End If
    Next k
    arr = forwardElimination(matrix_a, delta)
    x = backwardSubstitutution(arr(1), arr(2))
    For k = 1 To n
      inverse(k, i) = x(k)
    Next k
  Next i
  gaussInverseMatrix = inverse 
End Function

Sub kadai15()
  Dim matrix_a As Variant
  matrix_a = getMatrix(1,1,3,3)
  Call printMatrix(1,1, gaussInverseMatrix(matrix_a))
  Call printMatrix(1,7, matrixProduct(gaussInverseMatrix(matrix_a), matrix_a))
End Sub

Sub kadai16()
  '学籍番号7421004
  Dim a As Integer, b As Integer, c As Integer
  a = 0: b = 0: c = 4
  Dim matrix_a() As Double: ReDim matrix_a(4, 4)
  'matrix_a As Variantなどにして、=Array(Array(),Array())のようにすると
  '実質的な二次元配列になるが、要素の指定（インデックスの指定方法が）
  'matrix_a(1)(1)のようになり、普段のものと異なってしまうため今回は使うのを断念
  '大幅な修正が必要にるが上記のVariant的な書き方に統一したほうがスムーズかもしれない
  matrix_a(1,1) = a - 2: matrix_a(1,2) = b: matrix_a(1,3) = c: matrix_a(1,4) = a+b+c
  matrix_a(2,1) = c: matrix_a(2,2) = a - 3: matrix_a(2,3) = b: matrix_a(2,4) = a+b+c
  matrix_a(3,1) = b: matrix_a(3,2) = c: matrix_a(3,3) = a - 4: matrix_a(3,4) = a+b+c
  matrix_a(4,1) = 1: matrix_a(4,2) = 1: matrix_a(4,3) = 1: matrix_a(4,4) = 1
  Call printVec(2, 1, gaussElimination(matrix_a, Array(6, 8, -2, 0.5)))
  Call printMatrix(2, 3, gaussInverseMatrix(matrix_a))
End Sub

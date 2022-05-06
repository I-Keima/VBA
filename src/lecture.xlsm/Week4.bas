Attribute VB_Name = "Week4"

Option Explicit
Option Base 1 '配列の添え字の最小値を1に設定

Function forwardElimination(matrix_a As Variant, b As Variant) As Variant
  Dim i As Integer, k As Integer, j As Integer, n As Integer
  n = UBound(matrix_a, 1)
  For k = 1 To n - 1
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

Sub kadai12()
  Dim arr As Variant
  arr = forwardElimination(createOnTriangleMatrix(5), createRandNdVec(5))
  Call printMatrix(1, 1, arr(1))
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


	

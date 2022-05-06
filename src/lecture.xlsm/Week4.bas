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
  forwardElimination = Array(matrix_a, b)
End Function

Sub kadai12()
  Dim arr As Variant
  arr = forwardElimination(createOnTriangleMatrix(5), createRandNdVec(5))
  Call printMatrix(1, 1, arr(1))
End Sub

Function backwardSubstitutution(matrix_u As Variant, y as Variant) As Variant
	
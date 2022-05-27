Attribute VB_Name = "Common"

Option Explicit
Option Base 1 '配列の添え字の最小値を1に設定


Function create_matrix(row_size, col_size)
  '任意サイズのMatrixオブジェクト（Arrayの入れ子構造）を作成する
  Dim i As Integer
  Dim ans, row As Variant

  ans = Array()
  ReDim ans(row_size)
  For i = 1 To row_size
    row = Array()
    ReDim row(col_size)
    ans(i) = row
  Next i

  create_matrix = ans
End Function


Function t_matrix(m)
  '行列を転置する
  Dim ans As Variant
  ans = create_matrix(UBound(m(1)), UBound(m))

  Dim i, j As Integer
  For i = 1 To UBound(ans)
    For j = 1 To UBound(ans(1))
      ans(i)(j) = m(j)(i)
    Next j
  Next i
  t_matrix = ans
End Function


Function read_matrix(row_1, col_1, Optional ByVal row_size As Integer = 0, Optional ByVal col_size As Integer = 0)
  'シートからMatrixオブジェクトを取得する
  '第3,4引数の行列のサイズを指定しなかった場合は自動で空白のセルを区切りとしてサイズを取得
  Dim i, j As Integer
  Dim ans As Variant
  Dim last As Boolean
  
  '行のサイズ取得(変数を再利用するためrow_sizeは値渡し)
  If row_size = 0 Then
    last = False
    i = 0
    While not last
      If Cells(row_1 + i, col_1) = "" Then
        last = True
      Else
        i = i + 1
      End If
    Wend
    row_size = i
  End If

  '列サイズの取得
  If col_size = 0 Then
    last = False
    j = 0
    While not last
      If Cells(row_1, col_1 + j) = "" Then
        last = True
      Else
        j = j + 1
      End If
    Wend
    col_size = j
    End If

  ans = create_matrix(row_size, col_size)
  For i = 1 To row_size
    For j = 1 To col_size
      ans(i)(j) = Cells(row_1 + i - 1, col_1 + i - 1)
    Next j
  Next i
  read_matrix = ans

End Function

  
Function print_matrix(row, col, m)
  Dim i, j As Integer
  For i = 1 To UBound(m)
    For j = 1 To UBound(m(1))
      Cells(row + i - 1, col + j - 1) = m(i)(j)
    Next j
  Next i
End Function


Function matrix_plus(m1, m2)
  Dim i, j As Integer
  Dim ans As Variant: ans = m1

  For i = 1 To UBound(m1)
    For j = 1 To Uboud(m1(1))
      ans(i)(j) = m1(i)(j) + m2(i)(j)
    Next j
  Next i

  matrixPlus = ans
End Function


Function forward_elimination_matrix(matrix_a As Variant, matrix_b As Variant) As Variant
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
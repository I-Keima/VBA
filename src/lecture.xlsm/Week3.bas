Attribute VB_Name = "Week3"

Option Explicit
Option Base 1 '配列の添え字の最小値を1に設定


Function getVec(k As Integer, l As Integer, n As Integer) As Double()
	Dim arr() As Double, i As Integer: ReDim arr(n)
	'For文は分かりやすさのため配列の添え字と揃うようにする
	For i = 1 To n
		arr(i) = Cells(k + i - 1, l)
	Next i
	getVec = arr
End Function

Function getMatrix(k As Integer, l As Integer, m As Integer, n As Integer) As Double()
	Dim arr() As Double, i As Integer, j As Integer: ReDim arr(m, n)
	For i = 1 To m
		For j = 1 To n
			arr(i, j) = Cells(k + i - 1, j + l - 1)
		Next j
	Next i
	getMatrix = arr
End Function

Function printVec(k As Integer, l As Integer, arr As Variant) As Variant
	Dim i As Integer, n As Integer
	n = UBound(arr)
	For i = 1 To n
		Cells(i + k - 1, l) = arr(i)
	Next i
End Function

Function printMatrix(k As Integer, l As Integer, arr As Variant) As Variant
	Dim i As Integer, j As Integer, m As Integer, n As Integer
	m = UBound(arr, 1)
	n = Ubound(arr, 2)
	For i = 1 To m
		For j = 1 To n
			Cells(i + k - 1 , j + l - 1) = arr(i, j)
		Next j 
	Next i
End Function

Sub kadai1and3()
	Dim arr As Variant
	Cells(1, 1) = "課題1用1次元配列"
	arr = getVec(2, 1, 5)
	Cells(1, 2) = "課題3出力結果"
	Call printVec(2, 2, arr)
End Sub

Sub kadai2ans4()
	Dim arr As Variant
	Cells(1, 4) = "課題2用2次元配列"
	arr = getMatrix(2, 4, 5, 2)
	Cells(1, 6) = "課題4出力結果"
	Call printMatrix(2, 6, arr)
End Sub

Function create1NdVec(n as Integer) as Double()
	Dim arr() As Double: Redim arr(n)
	Dim i As Integer
	For i = 1 To n
		arr(i) = 1
	Next i
	create1NdVec = arr
End Function

Function createRandNdVec(n As Integer) as Double()
	Dim arr() As Double: ReDim arr(n)
	Dim i As Integer
	For i = 1 To n
		arr(i) = Int(6 * Rnd + 1)
	Next i
	createRandNdVec = arr
End Function

Sub kadai5()
	Call printVec(2, 7, create1NdVec(5))
	Call printVec(2, 8, createRandNdVec(5))
End Sub

Function createOnTriangleMatrix(n As Integer) As Double()
	Dim arr() As Double: Redim arr(n, n)
	Dim i As Integer, j As Integer
	For i = 1 To n
		For j = 1 To n
			if i <= j Then
				arr(i, j) = j - i + 1
			Else
				arr(i, j) = 0
			End If
		Next j
	Next i
	createOnTriangleMatrix = arr
End Function

Sub kadai6()
	Call printMatrix(2, 10, createOnTriangleMatrix(5))
End Sub

Function dotProd(x As Variant, y As Variant) As Double
	Dim n As Integer, i As Integer, dot_prod As Double
	n = UBound(x)
	if UBound(y) <> n Then
		Call MsgBox("Invalid vectors")
		Exit Function
	End If
	For i = 1 To n
		dot_prod = dot_prod + x(i) * y(i)
	Next i
	dotProd = dot_prod
End Function

Sub kadai7()
	Dim arr1 As Variant, arr2 As Variant
	arr1 = getVec(2, 12, 5)
	arr2 = getVec(2, 13, 5)
	Call MsgBox(dotProd(arr1, arr2))
End Sub

Function matrixVectorProduct(matrix_a As Variant, x As Variant) As Double()
	Dim n as integer, m as integer, i as integer, j as integer, arr() As Double
	n = UBound(x)
	m = Ubound(matrix_a, 1)
	ReDim arr(m)
	For i = 1 To m
		For j = 1 To n
			arr(i) = arr(i) + matrix_a(i, j) * x(j)
		Next j
	Next i
	matrixVectorProduct = arr
End Function

Sub kadai8()
	Call printMatrix(8, 1, matrixVectorProduct(createOnTriangleMatrix(5), createRandNdVec(5)))
End Sub

Function matrixProduct(matrix1 As Variant, matrix2 As Variant) As Double()
	Dim i As Integer, j as integer, k as integer, l as integer, m as integer, n as integer
	l = UBound(matrix1, 1): m = UBound(matrix1, 2): n = UBound(matrix2, 2)
	Dim arr() As Double: ReDim arr(l, n)
	For i = 1 To l
		For j = 1 To n
			For k = 1 To m
				arr(i, j) = arr(i, j) + matrix1(i, k) * matrix2(k, j)
			Next k
		Next j
	Next i
	matrixProduct = arr
End Function

Sub kadai9()
		Call printMatrix(8, 1, matrixVectorProduct(createOnTriangleMatrix(5), createOnTriangleMatrix(5)))
End Sub

Sub kadai10()
	Dim n As Double
	n = dotProd(create1NdVec(2022), createRandNdVec(2022))/dotProd(create1NdVec(2022), create1NdVec(2022))
	MsgBox n 
End Sub

Sub kadai11()
	Call printVec(1, 1, matrixVectorProduct(createOnTriangleMatrix(2022), create1NdVec(2022)))
	MsgBox Cells((7421004 Mod 2022) + 1, 1)
End Sub
'遭遇したエラー
	'当初printMatrix()を用いていたが、”n = Ubound(arr, 2)”の部分でインデクスが範囲外というエラーが発生した
	'これは与えられた行列を計算する際にmatrixVectorProductを用いて行列とベクトルの積で計算したことで、
	'printMatrix()の引数に与えられるべき行列がベクトル（1次元の配列）となり、存在しない次元を参照したことでエラーとなった。
	'与えられた問題文より行列の積の計算結果は必ずベクトルとなるためprintVec()を用いて取り扱う行列の次元を揃えた。
	
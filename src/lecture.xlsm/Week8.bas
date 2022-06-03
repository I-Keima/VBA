Attribute VB_Name = "Week8"

Option Explicit
Option Base 1

'標本平均の期待値は1/2、分散は1/12n 

Function rnd_uniform_distribution(n As Long) As Variant
  Dim arr() As Double: ReDim arr(n)
  Dim i As Long
  For i = 1 To n
    arr(i) = Rnd
  Next i
  rnd_uniform_distribution = arr
End Function

Function sample_average(arr As Variant) As Double
  Dim i, n, total As Long
  n = UBound(arr)
  total = 0
  For i = 1 To n
    total = total + arr(i)
  Next i
  sample_average = total / n
End Function

Sub week8_1()
  'n：1, 2, 5, 10
  'm：2000
  'nは別配列に格納
  Dim arr, sample As Variant
  
  Dim i, n, m, m_max As Integer
  'mは任意の2000以上の整数
  m_max = 2000
  Dim n_list As Variant
  '任意のnの値を設定
  n_list = Array(1, 2, 5, 10)
  arr = create_matrix(UBound(n_list), m_max)

  For i = 1 To UBound(n_list)
    For m = 1 To m_max
      sample = rnd_uniform_distribution(n_list(i))
      arr(i)(m) = sample_average(sample)
    Next m
  Next i

  Call print_matrix(1,1,arr)
End Sub




  


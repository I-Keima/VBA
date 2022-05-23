Attribute VB_Name = "Week6"

Option Explicit
Option Base 1

Sub kadai1()
  Dim arr() As Double: ReDim arr(1000)
  Dim i As Integer
  For i = 1 To 1000
    arr(i) = Rnd
  Next i
  Call printVec(1,1,arr)
End Sub


Sub kadai2()
  Dim arr, section_list, ans As Variant
  Dim i, j, k As Integer
  Dim l_limit, u_limit, frequency As Double

  '区間の幅を0.05とする
  ans = Array()
  ReDim ans(21)
  ans(1) = Array("番号", "下限", "上限", "度数")

  For i = 1 To 20 
    l_limit = (i - 1) * 0.05 
    u_limit = i * 0.05
    '各行の第２列に下限、第３列に上限を入れる
    section_list = Array(i, l_limit, u_limit, 0)
    ans(i + 1) = section_list
  Next i

  arr = getVec(1,1,1000)
  For i = 1 To 1000
    k = Int(arr(i) / 0.05) + 1
    section_list = ans(k + 1)
    section_list(UBound(section_list)) = section_list(UBound(section_list)) + 1
    ans(k + 1) = section_list
  Next i

  For i = 1 To 21
    section_list = ans(i)
    For j = 1 To UBound(section_list)
      Cells(i, 2 + j) = ans(i)(j)
    Next j
  Next i

  'ワークシートで作業仕様とするとパソコンがクラッシュするためVBAでワークシート関数を用いて代用....
  With ActiveSheet.Shapes.AddChart.Chart

    .ChartType = xlColumnClustered
    .SetSourceData Range(Cells(2, 6), Cells(21, 6))

  End With
End Sub


Function binary_distribution(n As Double, p As Double, m As Integer)
  '２項分布の確率変数(成功回数)をm行の一次元配列で返す
  Dim i, j As Integer, r As Double
  Dim ans() As Double: ReDIM ans(m)
  'mは繰り返し数

  For i = 1 To m
    ans(i) = 0
    For j = 1 To n
      r = Rnd
      If r < p Then
        ans(i) = ans(i) + 1
      End If
    Next j 
  Next i
  
  binary_distribution = ans
End Function

Sub kadai3_1()
  Dim ans As Variant
  ans = binary_distribution(1, 0.4, 1)
  MsgBox ans(1)
End Sub

Sub kadai3_2()
  Dim ans As Variant
  ans = binary_distribution(20, 0.4, 1)
  MsgBox ans(1)
End Sub

Function kadai3_3()
  Dim ans As Variant
  ans = binary_distribution(29, 0.33, 1204)
  kadai3_3 = ans
End Function

Sub kadai3_4()
  Call Worksheets("sheet2").Activate
  Dim arr, section_list, ans As Variant
  Dim i, j, k As Integer
  Dim l_limit, u_limit, frequency As Double

  '区間の幅を1とする
  Dim width As Double: width = 1
  ans = Array()
  ReDim ans(29 / width + 1)
  ans(1) = Array("番号", "下限", "上限", "度数")

  For i = 1 To UBound(ans) - 1 
    l_limit = (i - 1) * width
    u_limit = i * width
    '各行の第２列に下限、第３列に上限を入れる
    section_list = Array(i, l_limit, u_limit, 0)
    ans(i + 1) = section_list
  Next i

  arr = kadai3_3()
  For i = 1 To UBound(arr)
    k = Int(arr(i) / width) + 1
    section_list = ans(k + 1)
    section_list(UBound(section_list)) = section_list(UBound(section_list)) + 1
    ans(k + 1) = section_list
  Next i

  For i = 1 To UBound(ans)
    section_list = ans(i)
    For j = 1 To UBound(section_list)
      Cells(i, j) = ans(i)(j)
    Next j
  Next i

  'ワークシートで作業仕様とするとパソコンがクラッシュするためVBAでワークシート関数を用いて代用....
  With ActiveSheet.Shapes.AddChart.Chart

    .ChartType = xlColumnClustered
    .SetSourceData Range(Cells(2, 4), Cells(UBound(ans) + 1, 4))

  End With

End Sub

Sub kadai4()
  Dim u, v, x, y As Double
  Dim ans() As Double: ReDim ans(2000)

  Dim i As Integer
  For i = 1 To 2000
    x = Sin(4 * Ath(1) * 2 * Rnd) * (-2 * Log(Rnd)) ^ (1/2) 
    ans(i) = x
  Next i

  Call WorkSheets("sheet3").Activate
  Call printVec(1,1,ans)
End Sub
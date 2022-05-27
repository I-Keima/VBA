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
  ReDim ans(29 / width + 2)
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
  Dim i, j, k As Integer
  Dim u, v, x, y As Double
  Dim ans() As Double: ReDim ans(2000)

  For i = 1 To 2000
    x = Sin(4 * Atn(1) * 2 * Rnd) * (-2 * Log(Rnd)) ^ (1/2) 
    ans(i) = x
  Next i

  Call WorkSheets("sheet3").Activate
  '一次元配列をシートに出力(行列の一回目で作成した関数)
  Call printVec(1,1,ans)

  '度数表の作成
  Dim frequency_table, section_list As Variant
  '確率変数は±3の範囲まででとる
  '区間の幅を1とする
  Dim width As Double: width = 0.1
  frequency_table = Array()
  '-3以下と3以上、ヘッダー分
  ReDim frequency_table(6 / width + 2 + 1)
  frequency_table(1) = Array("番号", "下限(以上)", "上限(未満)", "度数")
  frequency_table(2) = Array(1, "-", -3, 0)
  frequency_table(UBound(frequency_table)) = Array(UBound(frequency_table), 3, "-", 0)

  Dim l_limit, u_limit, frequency As Double
  For i = 1 To UBound(frequency_table) - 3 
    l_limit = (i - 1) * width - 3
    u_limit = i * width - 3
    '各行の第２列に下限、第３列に上限を入れる
    section_list = Array(i + 1, l_limit, u_limit, 0)
    frequency_table(i + 2) = section_list
  Next i

  Dim row As Integer
  For i = 1 To UBound(ans)
    'Intは小数点以下の切り捨て、下限値-3がwidth*0に値するよう調整
    'ただし添え字は1スタートなので+1
    '±3より大きいものだけ個別に判定
    k = Int((ans(i) + 3) / width) + 1
    If k <= 0 Then
      row = 2
    ElseIf k >= 61 Then
      row = UBound(frequency_table)
    Else
      row = k + 2
    End If
    section_list = frequency_table(row)
    section_list(UBound(section_list)) = section_list(UBound(section_list)) + 1
    frequency_table(row) = section_list
  Next i

  For i = 1 To UBound(frequency_table)
    section_list = frequency_table(i)
    For j = 1 To UBound(section_list)
      Cells(i, j + 2) = frequency_table(i)(j)
    Next j
  Next i

  'ワークシートで作業仕様とするとパソコンがクラッシュするためVBAでワークシート関数を用いて代用....
  With ActiveSheet.Shapes.AddChart.Chart

    .ChartType = xlColumnClustered
    .SetSourceData Range(Cells(2, 6), Cells(UBound(frequency_table) + 2, 6))

  End With
End Sub

'改善点：端点の取り扱いが難しく、どうしても各々の度数表の範囲に合わせた
'”定数"(29や6など)がコードに入ってしまい、
'コードの再利用が面倒になってしまっている点。
'またシートを利用する際にそのセルの座標に関した定数も多く入ってしまった点
'またヘッダーや表の添え字（番号）をつける処理は別関数を作ったほうがいいと感じた。
'正規分布の範囲は±3ではなく±4のほうがよりヒストグラムの端のほうが奇麗になったのではと感じた。
'このような変更を加えにくくなるため、やはり定数をなるべく含めないコードを書きたいと思った。
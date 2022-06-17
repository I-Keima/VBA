Attribute VB_Name = "Week8"

Option Explicit
Option Base 1


Function rnd_uniform_distribution(n As Long) As Variant
  Dim arr() As Double: ReDim arr(n)
  Dim i As Long
  For i = 1 To n
    arr(i) = Rnd
  Next i
  rnd_uniform_distribution = arr
End Function

Function sample_average(arr As Variant) As Double
  Dim i, n As Long
  Dim total As Double
  n = UBound(arr)
  total = 0
  For i = 1 To n
    total = total + arr(i)
  Next i
  sample_average = total / n
End Function

Sub week8_1()
  Call WorkSheets("Sheet1").Activate
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

  '一様分布の乱数を発生させ、その標本平均を求める
  For i = 1 To UBound(n_list)
    For m = 1 To m_max
      sample = rnd_uniform_distribution(CLng(n_list(i)))
      arr(i)(m) = sample_average(sample)
    Next m
  Next i

  'n*mの標本平均のデータarrをセルに表示
  Call print_matrix(1, 1, arr)
  
  '標準化を行う
  Dim standard_arr As Variant: standard_arr = arr
  For i = 1 To UBound(n_list)
    '標本平均の期待値は1/2、分散は1/12n 
    n = n_list(i)
    standard_arr(i) = standardization(arr(i), 0.5, 1 / 12 / n, CLng(n))
  Next i

  '標準化したデータを表示
  Call print_matrix(UBound(n_list) + 2, 1, standard_arr)

  '度数表をもとめ表示する
  Dim frequency_table As Variant
  For i = 1 to UBound(n_list)
    '区間の幅は0.2とする
    frequency_table = make_frequency_table(standard_arr(i), 0.5 * n_list(i) ^ (1/2) * 12 * n / 5, 0.1 * i)
    'nの値の個数+3行、1 + (i - 1) * 5列目を起点とし度数表をセルに表示
    Call print_matrix(3 + UBound(n_list) * 2, 1 + (i - 1) * 5, frequency_table)
  Next i
  
End Sub


Function standardization(arr As Variant, mu As Double, sigma As Double, n As Long) As Variant
  Dim i, m As Long
  m = UBound(arr)
  Dim z() As Double: ReDim z(m)

  For i = 1 To m
    z(i) = (arr(i) - mu) * (n) ^ (1/2) / sigma
  Next i
  
  standardization = z

End Function


Function make_frequency_table(arr As Variant, Optional range As Double = 4, Optional width As Double = 0.2)
  'ただし分布は標準正規分布とする
  'rangeは度数を取る範囲の片側（半分）
  '度数表の作成
  Dim frequency_table, section_list As Variant
  '確率変数は±rangeの範囲まででとる
  frequency_table = Array()
  '-4以下と4以上、ヘッダー分
  ReDim frequency_table(2 * range / width + 2 + 1)
  frequency_table(1) = Array("番号", "下限(以上)", "上限(未満)", "度数")
  frequency_table(2) = Array(1, "-", -range, 0)
  frequency_table(UBound(frequency_table)) = Array(UBound(frequency_table), range, "-", 0)

  Dim i, k As Integer
  Dim l_limit, u_limit, frequency As Double
  For i = 1 To UBound(frequency_table) - 3 
    l_limit = (i - 1) * width - range
    u_limit = i * width - range
    '各行の第２列に下限、第３列に上限を入れる
    section_list = Array(i + 1, l_limit, u_limit, 0)
    frequency_table(i + 2) = section_list
  Next i

  Dim row As Integer
  For i = 1 To UBound(arr)
    'Intは小数点以下の切り捨て、下限値-rangeがwidth*0に値するよう調整
    'ただし添え字は1スタートなので+1
    '±rangeより大きいものだけ個別に判定
    k = Int((arr(i) + range) / width) + 1
    If k <= 0 Then
      row = 2
    ElseIf k >= 2 * range / width + 1 Then
      row = UBound(frequency_table)
    Else
      row = k + 2
    End If
    section_list = frequency_table(row)
    section_list(UBound(section_list)) = section_list(UBound(section_list)) + 1
    frequency_table(row) = section_list
  Next i 

  make_frequency_table = frequency_table

End Function


Function botu_ni_sita_yatu()
    '度数表の作成
  Dim section_list, ans As Variant
  Dim i, j, k As Integer
  Dim l_limit, u_limit, frequency As Double
  
  ans = Array(): ReDim ans(1 / width + 1)
  ans(1) = Array("番号", "下限", "上限", "度数")

  For i = 1 To UBound(ans) - 1 
    l_limit = (i - 1) * width
    u_limit = i * width
    '各行の第２列に下限、第３列に上限を入れる
    section_list = Array(i, l_limit, u_limit, 0)
    ans(i + 1) = section_list
  Next i

  For i = 1 To UBound(arr)
    k = Int(arr(i) / width) + 1
    section_list = ans(k + 1)
    section_list(UBound(section_list)) = section_list(UBound(section_list)) + 1
    ans(k + 1) = section_list
  Next i
End Function


Sub week8_ouyou()
  Call Worksheets("Sheet2").Activate
  'n：1, 2, 5, 10
  'm：2000
  'nは別配列に格納
  Dim arr, sample As Variant
  
  Dim i, n, m, m_max As Integer
  'mは任意の2000以上の整数
  m_max = 2000
  Dim n_list As Variant
  '任意のnの値を設定
  n_list = Array(30)
  arr = create_matrix(UBound(n_list), m_max)

  '一様分布の乱数を発生させ、その標本平均を求める
  For i = 1 To UBound(n_list)
    For m = 1 To m_max
      sample = binary_distribution(10, 0.4, CInt(n_list(i)))
      arr(i)(m) = sample_average(sample)
    Next m
  Next i

  'n*mの標本平均のデータarrをセルに表示
  Call print_matrix(1, 1, arr)
  
  '標準化を行う
  Dim standard_arr As Variant: standard_arr = arr
  For i = 1 To UBound(n_list)
    '標本平均の期待値は1/2、分散は1/12n 
    n = n_list(i)
    standard_arr(i) = standardization(arr(i), 0.5, 1 / 12 / n, CLng(n))
  Next i

  '標準化したデータを表示
  Call print_matrix(UBound(n_list) + 2, 1, standard_arr)

  '度数表をもとめ表示する
  Dim frequency_table As Variant
  For i = 1 to UBound(n_list)
    '区間の幅は0.2とする
    frequency_table = make_frequency_table(standard_arr(i), 0.5 * n_list(i) ^ (1/2) * 12 * n / 2)
    'nの値の個数+3行、1 + (i - 1) * 5列目を起点とし度数表をセルに表示
    Call print_matrix(3 + UBound(n_list) * 2, 1 + (i - 1) * 5, frequency_table)
  Next i
  
End Sub

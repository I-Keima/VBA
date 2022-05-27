Attribute VB_Name = "Week7"

Option Explicit
Option Base 1

Sub law_of_large_number()
  Dim n, i As Integer
  Dim x_average() As Double: ReDim x_average(1200)
  Dim bernoulli_distribution As Variant

  For n = 1 To 1200
    bernoulli_distribution = binary_distribution(1, 0.35, n)
    For i = 1 To UBound(bernoulli_distribution)
      x_average(n) = x_average(n) + bernoulli_distribution(i)
    Next i
    x_average(n) = x_average(n) / n
  Next n

End Sub

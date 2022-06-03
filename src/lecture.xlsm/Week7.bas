Attribute VB_Name = "Week7"

Option Explicit
Option Base 1

Function law_of_large_number(n_max As Integer, Optional p As Double = 0.5)
  Dim i, n As Integer
  Dim x_average() As Double: ReDim x_average(n_max)
  Dim bernoulli_distribution As Variant

  For n = 1 To n_max
    bernoulli_distribution = binary_distribution(1, CDbl(p), CInt(n))
    For i = 1 To UBound(bernoulli_distribution)
      x_average(n) = x_average(n) + bernoulli_distribution(i)
    Next i
    x_average(n) = x_average(n) / n
  Next n
  law_of_large_number = x_average
End Function


Sub kadai7_1()
  Dim i As Integer
  Dim x_average AS Variant: x_average = law_of_large_number(100)
  Dim num As Variant: ReDim num(5000)
  Dim logN As Variant: ReDim logN(5000)
  For i = 1 To 5000
    num(i) = i
    logN(i) = Log(i)
  Next i

 
  Call printVec(2,1,num)
  Call printVec(2,2,logN)
  Call printVec(2,3,x_average)

  '再度繰り返す
  x_average = law_of_large_number(500)
  Call printVec(2,4,x_average)

  x_average = law_of_large_number(1000)
  Call printVec(2,5,x_average)

  x_average = law_of_large_number(5000)
  Call printVec(2,6,x_average)


  '再度繰り返す
  x_average = law_of_large_number(100)
  Call printVec(2,7,x_average)

  x_average = law_of_large_number(500)
  Call printVec(2,8,x_average)

  x_average = law_of_large_number(1000)
  Call printVec(2,9,x_average)

  x_average = law_of_large_number(5000)
  Call printVec(2,10,x_average)
End Sub


Sub kadai7_1_7()
  Dim i As Integer
  Dim x_average AS Variant: x_average = law_of_large_number(100, 0.6)
  Call printVec(2,12,x_average)

  '再度繰り返す
  x_average = law_of_large_number(500)
  Call printVec(2,13,x_average)

  x_average = law_of_large_number(1000)
  Call printVec(2,14,x_average)

  x_average = law_of_large_number(5000)
  Call printVec(2,15,x_average)
  
  '再度繰り返す
  x_average = law_of_large_number(100)
  Call printVec(2,16,x_average)

  x_average = law_of_large_number(500)
  Call printVec(2,17,x_average)

  x_average = law_of_large_number(1000)
  Call printVec(2,18,x_average)

  x_average = law_of_large_number(5000)
  Call printVec(2,19,x_average)
End Sub

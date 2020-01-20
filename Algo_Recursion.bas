Attribute VB_Name = "Module8"
Option Explicit

Sub demo_Recursion()
    Debug.Print fac(8)
End Sub

Public Function fac(num As Integer) As Double
       If num < 2 Then
          fac = 1
       Else
          fac = num * fac(num - 1)
       End If
       
       Debug.Print "num = " & num & ", fac = " & fac
End Function

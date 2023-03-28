Attribute VB_Name = "Ä£¿é1"
Function MT(a)
MT = 0.0000000885 * a ^ 6 - 0.0000072339 * a ^ 5 + 0.000182 * a ^ 4 - 0.001137 * a ^ 3 + 0.00075 * a ^ 2 + 0.227 * a + 0.00125
End Function

Function MN(a)
MN = -0.00005 * a ^ 6 + 0.00497 * a ^ 5 - 0.20639 * a ^ 4 + 4.2153 * a ^ 3 - 42.97273 * a ^ 2 + 308.90978 * a + 30.13818
End Function

Function MPout(a, b)
MPout = a * b * 1000 / 9550
End Function

Function MU(a)
MU = -0.0000011609 * a ^ 6 + 0.00012 * a ^ 5 - 0.004879 * a ^ 4 + 0.098734 * a ^ 3 - 1.012609 * a ^ 2 + 7.233075 * a + 0.716773
End Function


Function MPin(a, b)
MPin = a * b
End Function

Function MA(a)
MA = -0.0018 * a ^ 3 + 0.1319 * a ^ 2 + 0.0157 * a + 1.7703
End Function

Function GPL(a, T, d, e, f, n, l, m, z, p, u, h, b)
f = a * (2 * T / d) * (e + f) / 2 * n / 60 * 2 * Cos(l)
c = 1.177 * m ^ 0.0214 * z ^ 0.0106 * p * d * (u / p ^ 2 * d ^ 2 * n / 60) * 0.34 * (n * d / 9.8) ^ 0.61 * (h) ^ 0.35 * (b) ^ 1.21
GPL = f + c
End Function

Function GN(a, b)
GN = a / b
End Function

Function GT(a, b, c)
GT = 9.55 * (c - b) / a

End Function
Function Eff(a)
Number = a
Select Case Number
Case 0 To 25
Eff = 0.84
Case 25.01 To 44
Eff = 0.86
Case 44.01 To 82
Eff = 0.88
Case Else
Debug.Print "Too big voltage"
End Select
End Function


Function PC(a, b)
PC = a * 3600 / b

End Function
Function PL(a, b)
PL = a * 3600 / b

End Function
Function Percentage(a, b)
Percentage = a / b

End Function

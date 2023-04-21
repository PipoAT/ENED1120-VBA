Attribute VB_Name = "Module1"
Sub Spring_Motion()

Dim A, K, M, t As Double

A = ActiveSheet.Cells(3, 3)
K = ActiveSheet.Cells(4, 3)
M = ActiveSheet.Cells(5, 3)
t = ActiveSheet.Cells(6, 3)

pt = A * Cos(Sqr(K / M) * t)
vt = -(A * Sqr(K / M)) * Sin(Sqr(K / M) * t)
at = -((A * K) / M) * Cos(Sqr(K / M) * t)
KEt = (1 / 2) * M * (vt ^ 2)
PEt = (1 / 2) * K * (pt ^ 2)


pt = Application.WorksheetFunction.Round(pt, 2)
vt = Application.WorksheetFunction.Round(vt, 1)
at = Application.WorksheetFunction.Round(at, 2)

KEt = Application.WorksheetFunction.Round(KEt, 1)
PEt = Application.WorksheetFunction.Round(PEt, 1)



ActiveSheet.Cells(9, 3) = pt
ActiveSheet.Cells(10, 3) = vt
ActiveSheet.Cells(11, 3) = at
ActiveSheet.Cells(12, 3) = KEt
ActiveSheet.Cells(13, 3) = PEt

End Sub


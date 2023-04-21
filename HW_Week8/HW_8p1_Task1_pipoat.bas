Attribute VB_Name = "Module1"
Sub Resultant()

Dim F1, F2, A1, A2, W, L As Double

F1 = ActiveSheet.Cells(2, 3)
F2 = ActiveSheet.Cells(3, 3)
A1 = ActiveSheet.Cells(4, 3)
A2 = ActiveSheet.Cells(5, 3)
W = ActiveSheet.Cells(6, 3)
L = ActiveSheet.Cells(7, 3)

A1 = (A1 * WorksheetFunction.Pi()) / 180
A2 = (A2 * WorksheetFunction.Pi()) / 180

FRX = (F1 * Cos(A1)) - (F2 * Cos(A2))
FRY = (F1 * Sin(A1)) - (F2 * Sin(A2)) - W

XR = (-(F2 * Sin(A2)) * L - W * (L / 2)) / FRY
FR = Sqr(FRX ^ 2 + FRY ^ 2)
AR = Atn(FRY / FRX)
AR = (AR * 180) / WorksheetFunction.Pi()

AR = Application.WorksheetFunction.Round(AR, 2)
FR = Application.WorksheetFunction.Round(FR, 2)
XR = Application.WorksheetFunction.Round(XR, 2)


ActiveSheet.Cells(9, 3) = FR
ActiveSheet.Cells(10, 3) = AR
ActiveSheet.Cells(11, 3) = XR

End Sub

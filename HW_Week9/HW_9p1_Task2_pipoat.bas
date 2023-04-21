Attribute VB_Name = "Module1"
Sub Analysis()

Dim gavg, sdev, day1, day2, day3, day4, day5 As Double
Dim cat, reaction, result As String



gavg = ActiveSheet.Cells(2, 3).Value
sdev = ActiveSheet.Cells(3, 3).Value
day1 = ActiveSheet.Cells(6, 3).Value
day2 = ActiveSheet.Cells(7, 3).Value
day3 = ActiveSheet.Cells(8, 3).Value
day4 = ActiveSheet.Cells(9, 3).Value
day5 = ActiveSheet.Cells(10, 3).Value

avg = (day1 + day2 + day3 + day4 + day5) / 5
ActiveSheet.Cells(11, 3).Value = avg

If avg > (gavg + sdev) Then
    cat = "More"
    ActiveSheet.Cells(12, 3).Value = cat
ElseIf avg < (gavg - sdev) Then
    cat = "Less"
    ActiveSheet.Cells(12, 3).Value = cat
Else
    cat = ""
    ActiveSheet.Cells(12, 3).Value = cat
End If


reactions = ActiveSheet.Cells(14, 3).Value

If ActiveSheet.Cells(12, 3).Value = "More" And reactions = "H" Or reactions = "N" Or reactions = "H,N" Then
    result = "Severe"
    ActiveSheet.Cells(16, 3).Value = result
ElseIf ActiveSheet.Cells(12, 3).Value = "" Or ActiveSheet.Cells(12, 3).Value = "Less" And reactions = "H" Or reactions = "N" Or reactions = "H,N" Then
    result = "Mild"
    ActiveSheet.Cells(16, 3).Value = result
ElseIf ActiveSheet.Cells(12, 3).Value = "Less" And ActiveSheet.Cells(12, 3).Value = "Less" And reactions = "" Then

    result = "Helpful"
    ActiveSheet.Cells(16, 3).Value = result
Else
    result = ""
    ActiveSheet.Cells(16, 3).Value = result
End If




End Sub

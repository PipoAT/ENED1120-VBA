Attribute VB_Name = "Module1"
Sub Registration()

Dim year As Integer
Dim lbs As Double
Dim weightclass As String
Dim fee As String


year = ActiveSheet.Cells(2, 3).Value
lbs = ActiveSheet.Cells(3, 3).Value

If year <= 2000 Then
    If lbs < 2700 Then
        weightclass = "Class 1"
        fee = "$26.50"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    ElseIf lbs <= 3800 And lbs >= 2700 Then
        weightclass = "Class 2"
        fee = "$35.50"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    ElseIf lbs > 3800 Then
        weightclass = "Class 3"
        fee = "$56.50"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    End If
ElseIf year >= 2001 And year <= 2010 Then
    If lbs < 2700 Then
        weightclass = "Class 4"
        fee = "$35.00"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    ElseIf lbs <= 3800 And lbs >= 2700 Then
        weightclass = "Class 5"
        fee = "$45.50"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    ElseIf lbs > 3800 Then
        weightclass = "Class 6"
        fee = "$62.50"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    End If
ElseIf year >= 2011 Then
    If lbs < 3500 Then
        weightclass = "Class 7"
        fee = "S49.50"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    ElseIf lbs >= 3500 Then
        weightclass = "Class 8"
        fee = "$62.50"
        ActiveSheet.Cells(5, 3).Value = weightclass
        ActiveSheet.Cells(6, 3).Value = fee
    End If
End If

        


End Sub

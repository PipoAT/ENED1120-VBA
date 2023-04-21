Attribute VB_Name = "Module2"
Sub CityAnalysis()

Dim oldpop As Double
Dim newpop As Double
Dim Growth As Double
Dim i, c, j As Integer
Dim city As String


c = 0
j = 1

For i = 2 To 1188 Step 1
    oldpop = ActiveSheet.Cells(i, 5).Value
    newpop = ActiveSheet.Cells(i, 4).Value
    Growth = (newpop - oldpop) / oldpop
    ActiveSheet.Cells(i, 6).Value = Growth
    If ActiveSheet.Cells(i, 6).Value < 0 Then
        c = c + 1
    End If
Next

ActiveSheet.Cells(1, 9).Value = c

city = ActiveSheet.Cells(2, 9).Value

Do Until ActiveSheet.Cells(j, 2) = city Or j = 1189
    j = j + 1
    ActiveSheet.Cells(3, 9) = "City is not in database"
    ActiveSheet.Cells(4, 9) = ""
Loop

If ActiveSheet.Cells(j, 2).Value = city Then
    ActiveSheet.Cells(3, 9).Value = ActiveSheet.Cells(j, 3).Value
    ActiveSheet.Cells(4, 9).Value = ActiveSheet.Cells(j, 6).Value
    
End If




End Sub

    


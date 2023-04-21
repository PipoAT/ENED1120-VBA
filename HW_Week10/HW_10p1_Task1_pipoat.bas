Attribute VB_Name = "Module2"
Sub MaxPower()

Dim RL As Double
Dim Vs As Double
Dim Rs As Double
Dim RLM As Double
Dim PLM As Double
Dim r, m, n As Double

m = 0
n = 0
Vs = ActiveSheet.Cells(2, 6).Value
Rs = ActiveSheet.Cells(3, 6).Value

For r = 3 To 122 Step 1
    RL = ActiveSheet.Cells(r, 2).Value
    I = Vs / (RL + Rs)
    ActiveSheet.Cells(r, 3).Value = I
Next

For r = 3 To 122 Step 1
    m = ActiveSheet.Cells(r, 2).Value
    If RLM < m Then
        RLM = m
    End If
    ActiveSheet.Cells(4, 6).Value = RLM
    n = ActiveSheet.Cells(r, 3).Value
    If PLM < n Then
        PLM = n
    End If
    ActiveSheet.Cells(5, 6).Value = PLM
Next

End Sub

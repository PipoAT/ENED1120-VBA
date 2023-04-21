Attribute VB_Name = "Module1"
Sub Analysis()

Dim V() As Double
Dim k As Integer
Dim A As Double
Dim RMS As Double
Dim alphamin As Double
Dim RMSmin As Double

k = 0
alphamin = 999999
RMSmin = 999999

While ActiveSheet.Cells(k + 8, 3).Value <> ""
    ReDim Preserve V(k) As Double
    V(k) = ActiveSheet.Cells(k + 8, 3).Value
    k = k + 1
Wend

For A = 0.001 To 0.2 Step 0.001
    RMS = Expo(V, A)
    If RMS < RMSmin Then
        RMSmin = RMS
        alphamin = A
    End If
Next

ActiveSheet.Cells(4, 3).Value = RMSmin
ActiveSheet.Cells(5, 3).Value = alphamin

    

End Sub

Function Expo(Ay() As Double, A As Double) As Double

Dim E() As Double
Dim k As Integer
Dim RMS As Double
Dim total As Double
Dim size As Double

total = 0
size = 0

For k = 0 To UBound(Ay)
    ReDim Preserve E(k) As Double
    
    If k = 0 Then
        E(k) = Ay(k)
    Else
        E(k) = E(k - 1) * (1 - A) + A * Ay(k - 1)
    End If
    
    size = size + 1
Next

For k = 0 To UBound(Ay)
    total = total + (E(k) - Ay(k)) * (E(k) - Ay(k))
Next

RMS = Sqr((1 / (size - 1)) * total)

Expo = RMS

End Function

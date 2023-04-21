Attribute VB_Name = "Module1"
Function Medical(temp As Double, heartrate As Double) As String

Dim coa As String


If temp > 103 Or heartrate > 110 Then
    coa = "Send to Emergency Room"
ElseIf temp > 100 Or heartrate > 95 Then
    coa = "Request a Doctor Visit"
ElseIf temp > 99 Or heartrate > 85 Then
    coa = "Give prescribed Medication"
Else
    coa = "Normal, no action"
End If

Medical = coa


End Function


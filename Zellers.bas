Attribute VB_Name = "Zellers"


Public Function ZellersAlgorithm(TheDay As Integer, TheMonth As Integer, TheYear As Integer)

Dim C As Integer
Dim Y As Integer

If TheMonth < 3 Then
    TheMonth = TheMonth + 12
    TheYear = TheYear - 1
End If

C = Int(TheYear / 100)
Y = TheYear - (100 * C)
Debug.Print C; "  "; Y
     
        TheAns = Int((2.6 * TheMonth) - 5.39) + Int(Y / 4) + Int(C / 4) + TheDay + Y - (2 * C)
        ZellersAlgorithm = (TheAns) - (7 * Int(TheAns / 7)) 'mod 7 can cause errors


End Function



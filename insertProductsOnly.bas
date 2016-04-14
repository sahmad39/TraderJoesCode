Attribute VB_Name = "Module3"


Sub ProdArray()
Worksheets("Sheet1").Activate
Range("A2").Activate
Do
    If ActiveCell.Value = "" Then Exit Do
        Select Case ActiveCell
            Case Is = ActiveCell
                InsideArry ActiveCell, ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 3)
       End Select
    Worksheets("Sheet1").Activate
    ActiveCell.Offset(1, 0).Activate
Loop

End Sub


Sub InsideArry(n, m, o)
    Worksheets("Sheet2").Activate
    Range("L8").Activate
    Do
        If ActiveCell = "" Then Exit Do
        Select Case ActiveCell
            Case Is = n
                If IsEmpty(ActiveCell.Offset(0, -2).Value) = True Then
                    ActiveCell.Offset(0, -2).Value = m
                    ActiveCell.Offset(0, -3).Value = o
                Else
                    ActiveCell.Offset(1, 0).EntireRow.Insert
                    ActiveCell.Offset(1, 0).Value = ActiveCell.Value
                    ActiveCell.Offset(1, -2).Value = m
                    ActiveCell.Offset(1, -3).Value = o
                    ActiveCell.Value = 0
                    ActiveCell.Offset(2, 0).Activate
                End If
            End Select
        ActiveCell.Offset(1, 0).Activate
    Loop
End Sub


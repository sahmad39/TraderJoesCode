Attribute VB_Name = "Module1"


Sub ProdArray()
Worksheets("Sheet1").Activate
Range("A2").Activate
Do
    If ActiveCell.Value = "" And ActiveCell.Offset(1, 0) = "" And ActiveCell.Offset(2, 0) = "" Then Exit Do
        Select Case ActiveCell
            Case Is = ActiveCell
                InsideArry ActiveCell, ActiveCell.Offset(0, 1), ActiveCell.Offset(0, 3)
       End Select
    Worksheets("Sheet1").Activate
    ActiveCell.Offset(1, 0).Activate
Loop
Worksheets("Sheet2").Activate
Range("J8").Activate
Do
    If ActiveCell.Offset(0, 2).Value = 55555 Then Exit Do
    If ActiveCell.Value = 55555 Then Exit Do
    If ActiveCell.Value = "" Then
    ActiveCell.EntireRow.Delete
    ActiveCell.Activate
    Else
        ActiveCell.Offset(1, 0).Activate
    
    End If
Loop

Columns(["L"]).EntireColumn.Delete

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

Sub cycle()

    Dim wk As Worksheet
    For Each wk In ThisWorkbook.Worksheets
        Dim total As Double

        ' get the row number of the last row with data
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
            For I = 2 To RowCount
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                'print ticker symbol
                Range("I" & 2 + j).Value = Cells(I, 1).Value
                'print total
                Range("J" & 2 + j).Value = total
                'reset total
                total = 0
        
                'move next row
                j = j + 1
            Else
                total = total + Cells(I, 7).Value
            End If
        Next I
    
Next wk

End Sub


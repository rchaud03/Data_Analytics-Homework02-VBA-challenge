Sub stock_sheets():

'Declaring variables for the sheets
Dim sheet As Worksheet

'For each sheet in the workbook, call the stock_watch subroutine

For Each sheet In ActiveWorkbook.Worksheets
    Call stock_watch(sheet)
Next

End Sub


Sub stock_watch(sheet As Worksheet)

'Determine how many rows contain data in a worksheet that contains data in the column "A"
    row_count = Cells(Rows.Count, "A").End(xlUp).Row
    Row = 0
    volume = 0
    Opening = 0
    Closing = 0

'For each row in row_count, if cell1 of said row is equal to cell 1 of the next

    For i = 1 To row_count
        If Cells(i, 1) = Cells(i + 1, 1) Then
            volume = volume + Cells(i + 1, 7)
        Else
            Row = Row + 1
            Cells(Row + 1, 9) = Cells(i + 1, 1)
            volume = volume + Cells(i + 1, 7)
            Cells(Row, 12) = volume
            volume = 0
            Opening = Opening + Cells(i + 1, 3)
            Cells(Row + 1, 13) = Opening
            Opening = 0

            If i > 1 Then
                Closing = Closing + Cells(i, 6)
                Cells(Row, 14) = Closing
                Closing = 0
            End If
        End If
    Next i

    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"


    Cells(Row + 1, 13).ClearContents


    For i = 2 To Row
        Cells(i, 11) = Cells(i, 14) / Cells(i, 13) - 1
        Cells(i, 10) = Cells(i, 14) - Cells(i, 13)
        If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        ElseIf Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i


    Range("M:N").Delete


End Sub

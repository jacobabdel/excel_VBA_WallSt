Attribute VB_Name = "Module3"
Sub wallstreetCalc()
'Begin calculations
Dim evalRange As Variant
Dim evalRow As Range
Dim openDate As Double
Dim closeDate As Double
    openDate = 0
    closeDate = 0
Dim year As String
    year = Left(Cells(2, 2), 4)
Dim stockVol As Double
Dim openVal As Double
Dim closeVal As Double
    stockVol = 0
    openVal = 0
    closeVal = 0
Dim j As Double
Dim i As Long

'iterates for each unique ticker value
For i = tickerCount To 2 Step -1
     With allRowsRange
        Range(allRowsRange).AutoFilter Field:=1, Criteria1:=Cells(i, 9)
        Set evalRange = Range(allRowsRange).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
            'within rows related to ticker, find rows: last, with first & last date of year
            With evalRange
                lastRow = evalRange.Rows.Count + evalRange.Row - 1
                openDate = Application.WorksheetFunction.Min(Range(Cells(evalRange.Row, 2), Cells(lastRow, 2)))
                closeDate = Application.WorksheetFunction.Max(Range(Cells(evalRange.Row, 2), Cells(lastRow, 2)))
            End With
        'within rows related to ticker, calculate stock volume and match values to dates found above
        For Each evalRow In evalRange
            stockVol = stockVol + Cells(evalRow.Row, 7)
            If Cells(evalRow.Row, 2) = openDate Then
                openVal = Cells(evalRow.Row, 3)
            ElseIf Cells(evalRow.Row, 2) = closeDate Then
                closeVal = Cells(evalRow.Row, 6)
            End If
        Next
    End With
    
    'calculates and displays values based on above
    Cells(i, 10) = stockVol
    Cells(i, 11) = closeVal - openVal
    
    If openVal = 0 Then
        Cells(i, 12) = 0 'avoid error when dividing by zero
    Else
        Cells(i, 12) = (closeVal - openVal) / openVal
    End If
    Cells(i, 12).NumberFormat = ("00.00%")
    
    'reset variables for next  i iteration
    lastRow = 0
    openDate = 0
    closeDate = 0
    stockVol = 0
    openVal = 0
    closeVal = 0
    
Next i

End Sub

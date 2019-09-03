Attribute VB_Name = "Module4"
Sub wallstreetGreats()
'Find the greats
Dim greatIncrease As Double
Dim greatDecrease As Double
Dim greatVolume As Double

    'calculates greats
    greatIncrease = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(tickerCount, 12)))
    greatDecrease = Application.WorksheetFunction.Min(Range(Cells(2, 12), Cells(tickerCount, 12)))
    greatVolume = Application.WorksheetFunction.Max(Range(Cells(2, 10), Cells(tickerCount, 10)))
    
    'displays greats values
    Cells(2, 16) = greatIncrease
    Cells(3, 16) = greatDecrease
    Cells(4, 16) = greatVolume
    Columns(16).AutoFit
    
'match values found above with related ticker
Dim searchRange As Range
Dim searchRow As Range
Dim valMatch As Variant

    Set searchRange = Range(Cells(2, 12), Cells(tickerCount, 12))
    With searchRange
        For Each searchRow In searchRange
            If Cells(searchRow.Row, 12) = greatIncrease Then
                Cells(2, 15) = Cells(searchRow.Row, 9)
            ElseIf Cells(searchRow.Row, 12) = greatDecrease Then
                Cells(3, 15) = Cells(searchRow.Row, 9)
            ElseIf Cells(searchRow.Row, 10) = greatVolume Then
                Cells(4, 15) = Cells(searchRow.Row, 9)
            End If
        Next
    End With

End Sub


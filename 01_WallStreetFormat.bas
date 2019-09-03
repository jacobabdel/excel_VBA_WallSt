Attribute VB_Name = "Module2"
Sub wallstreetFormat()
'label and size needed columns
Columns(8).ColumnWidth = 5
Columns(9).AutoFit
Columns(13).ColumnWidth = 5
Cells(1, 10) = "Total Stock Volume"
Columns(10).AutoFit
Cells(1, 11) = "Yearly Change"
Columns(11).AutoFit
Cells(1, 12) = "Percent Change"
Columns(12).AutoFit
Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Greatest Total Volume"
Columns(14).AutoFit
Cells(1, 15) = "Ticker"
Columns(15).AutoFit
Cells(1, 16) = "Value"

'Set Conditional Formatting Rules for Yearly Change column
With Range(Cells(2, 11), Cells(tickerCount, 11))
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 3
        .TintAndShade = 0
    End With
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 4
        .TintAndShade = 0
    End With
End With

Cells(2, 16).NumberFormat = ("00.00%")
Cells(3, 16).NumberFormat = ("00.00%")
Cells(4, 16).NumberFormat = ("000,000")

End Sub



Attribute VB_Name = "Module1"
'=====START HERE=====
Public allRowsRange As Variant
Public tickerCount As Long
Public lastRow As Double

Sub wallstreetMain()

Application.ScreenUpdating = False

Dim sourceWorkbook As Workbook
    Set sourceWorkbook = ThisWorkbook
Dim xsheet As Worksheet

For Each xsheet In sourceWorkbook.Worksheets
    xsheet.Select
    
        'Determine the last row of initial data
        With xsheet
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        End With
    
    allRowsRange = Range(Cells(1, 1), Cells(lastRow, 1)).Address
    'copy all unique values in tickers to new column
    xsheet.Range(allRowsRange).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=xsheet.Range("I1"), Unique:=True
        
        'Match tickerCount value to total of unique values found
        With xsheet
            tickerCount = .Cells(.Rows.Count, 9).End(xlUp).Row
        End With
    
    wallstreetFormat 'sets general formatting
    wallstreetCalc 'calculates amounts
    wallstreetGreats 'finds greats
    
    xsheet.AutoFilter.ShowAllData 'clears all filters

Next xsheet

Application.ScreenUpdating = True
End Sub


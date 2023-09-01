Attribute VB_Name = "Module1"
Sub Stock_OneYear_Data()

Dim Stock As String
Dim YearlyChange As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim PercentChange As Double
Dim StockVolume As Double

Dim TableRow As Integer
Dim LastRow As Double

Dim Year As String

Dim Max As Double
Dim Min As Double
Dim MaxVolume As Double

Dim ws As Worksheet

'Loop for every sheet in the workbook from census_data_2016-2019_pt2
For Each ws In ActiveWorkbook.Worksheets
    Year = ws.Name
    'Resetting these values for independent comparability between worksheets
    Max = 0
    Min = 0
    MaxVolume = 0

    'Formatting the first table
    For j = 9 To 12
        If j = 9 Then
            ws.Cells(1, j).Value = "Ticker"
        ElseIf j = 10 Then
            ws.Cells(1, j).Value = "Yearly Change"
        ElseIf j = 11 Then
            ws.Cells(1, j).Value = "Percent Change"
        Else
            ws.Cells(1, j).Value = "Total Stock Volume"
        End If
    Next j
    
    'Formatting the second table
    For k = 15 To 17
        If k = 15 Then
            ws.Cells(2, k).Value = "Greatest % Increase"
            ws.Cells(3, k).Value = "Greatest % Decrease"
            ws.Cells(4, k).Value = "Greatest Total Volume"
        ElseIf k = 16 Then
            ws.Cells(1, k).Value = "Ticker"
        Else
            ws.Cells(1, k).Value = "Value"
        End If
    Next k
    
    'row counter idea from credit_charges exercise
    TableRow = 2
    
    'Row counter from census_data_2016-2019_pt2
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    For i = 2 To LastRow
        'Checking for the opening value of a stock within one year
        If ws.Cells(i, 2).Value = Year & "0102" Then
            OpeningPrice = ws.Cells(i, 3).Value
        End If
        
        'Structure of if statement from credit_charges exercise
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            Stock = ws.Cells(i, 1).Value
            ClosingPrice = ws.Cells(i, 6).Value
            
            YearlyChange = ClosingPrice - OpeningPrice
            PercentChange = YearlyChange / OpeningPrice
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            ws.Range("I" & TableRow).Value = Stock
            ws.Range("J" & TableRow).Value = YearlyChange
            
            'Color Index from grader exercise
            If YearlyChange < 0 Then
                ws.Range("J" & TableRow).Interior.ColorIndex = 3
            Else
                ws.Range("J" & TableRow).Interior.ColorIndex = 4
            End If
            
            'FormatPercent function from Stack Overflow
            ws.Range("K" & TableRow).Value = FormatPercent(PercentChange, 2, vbTrue, vbFalse, vbFalse)
            
            If PercentChange < 0 Then
                ws.Range("K" & TableRow).Interior.ColorIndex = 3
            Else
                ws.Range("K" & TableRow).Interior.ColorIndex = 4
            End If
            
            'Searching for greatest % increase/decrease
            If ws.Range("K" & TableRow).Value > Max Then
                Max = ws.Range("K" & TableRow).Value
                ws.Range("Q2").Value = FormatPercent(Max, 2, vbTrue, vbFalse, vbFalse)
                ws.Range("P2").Value = Stock
            ElseIf ws.Range("K" & TableRow).Value < Min Then
                Min = ws.Range("K" & TableRow).Value
                ws.Range("Q3").Value = FormatPercent(Min, 2, vbTrue, vbFalse, vbFalse)
                ws.Range("P3").Value = Stock
            End If
            
            ws.Range("L" & TableRow).Value = StockVolume
            
            'Searching for greatest total volume
            If ws.Range("L" & TableRow).Value > MaxVolume Then
                MaxVolume = ws.Range("L" & TableRow).Value
                ws.Range("Q4").Value = MaxVolume
                ws.Range("P4").Value = Stock
            End If
            
            TableRow = TableRow + 1
            StockVolume = 0
        Else
            StockVolume = StockVolume + ws.Cells(i, 7).Value
        End If
    
    Next i

Next ws

End Sub

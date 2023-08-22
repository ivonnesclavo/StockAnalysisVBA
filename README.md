# StockAnalysisVBA
Module 2 Challenge
Sub StockAnalysisOnAllWorksheets()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            StockAnalysis ws
        End If
    Next ws
    
End Sub

Sub StockAnalysis(ws As Worksheet)

    ' Page Setup
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Define variables
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    ' Initialize variables
    SummaryRow = 2
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    YearOpen = ws.Cells(2, 3).Value
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    
    ' Loop through rows
    For i = 2 To LastRow
        ' Check if we are still in the same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Ticker column
            Ticker = ws.Cells(i, 1).Value
            
            'Percent Change Column
            YearClose = ws.Cells(i, 6).Value
            YearlyChange = YearClose - YearOpen

            If YearOpen = 0 Then
                PercentChange = 0
            Else
                PercentChange = YearlyChange / YearOpen
            End If
            
            'Total Stock Volume Column
            TotalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(SummaryRow, 7), ws.Cells(i, 7)))
            
            ' Print results
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
            ws.Cells(SummaryRow, 11).Value = PercentChange
            ws.Cells(SummaryRow, 12).NumberFormat = "0"
            ws.Cells(SummaryRow, 12).Value = TotalVolume
            
             'Add Conditional Formatting
            If YearlyChange >= 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            End If
            
            'Summary
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                GreatestIncreaseTicker = Ticker
            ElseIf PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                GreatestDecreaseTicker = Ticker
            End If
            
            If TotalVolume > GreatestVolume Then
                GreatestVolume = TotalVolume
                GreatestVolumeTicker = Ticker
            End If
            
            ' Move to the next summary row
            SummaryRow = SummaryRow + 1
            
            ' Reset yearOpen for the next ticker
            YearOpen = ws.Cells(i + 1, 3).Value
            End If
            Next i
    
            'Print Summary
            ws.Cells(2, 16).Value = GreatestIncreaseTicker
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(2, 17).Value = GreatestIncrease
    
            ws.Cells(3, 16).Value = GreatestDecreaseTicker
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).Value = GreatestDecrease
    
            ws.Cells(4, 16).Value = GreatestVolumeTicker
            ws.Cells(4, 17).NumberFormat = "0.00%"
            ws.Cells(4, 17).Value = GreatestVolume
    
End Sub

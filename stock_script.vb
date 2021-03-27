Sub stocksumm()

' Create tab count variable.
Dim tab_count As Integer

' Count number of tabs in workbook.
tab_count = ActiveWorkbook.Worksheets.Count
' Test that tab_count worked.
' MsgBox tab_count

' Create variables to hold stock data.
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim price_diff As Double
Dim percent_price_diff As Double
Dim volume As Double

' Create counter for the summary table row.
Dim summary_row As Integer

' Loop through all tabs in workbook.
For i = 1 To tab_count

    ' Reset summary table counter to 2 for each new tab.
    summary_row = 2

    ' Set initial value of stock data variables to 0 for each new tab.
    open_price = 0
    close_price = 0
    price_diff = 0
    percent_price_diff = 0
    volume = 0

    ' Find last row in each tab.
    Dim last_row As Long
    last_row = ActiveWorkbook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through the table of stock data in a given tab.
    For Row = 2 To last_row

        ' Collect data for the last row of each ticker.
        If ActiveWorkbook.Worksheets(i).Cells(Row, 1) <> ActiveWorkbook.Worksheets(i).Cells(Row + 1, 1) Then
            
            ' Collect data for the last row of each ticker symbol.
            ticker = ActiveWorkbook.Worksheets(i).Cells(Row, 1)
            close_price = ActiveWorkbook.Worksheets(i).Cells(Row, 6)
            
            ' Now that last row of ticker has been reached, calculate summary statistics.
            price_diff = close_price - open_price

            ' Eliminate division by zero error.
            If open_price <> 0 Then
                percent_price_diff = (close_price - open_price) / open_price
            End If

            volume = volume + ActiveWorkbook.Worksheets(i).Cells(Row, 7)

            ' Add summary table column titles.
            ActiveWorkbook.Worksheets(i).Cells(1, 9) = "Ticker"
            ActiveWorkbook.Worksheets(i).Cells(1, 10) = "Absolute Yearly Change"
            ActiveWorkbook.Worksheets(i).Cells(1, 11) = "Percent Yearly Change"
            ActiveWorkbook.Worksheets(i).Cells(1, 12) = "Total Yearly Stock Volume"
            
            ' Record summary statistics for each ticker in the summary table.
            ActiveWorkbook.Worksheets(i).Cells(summary_row, 9) = ticker
            ActiveWorkbook.Worksheets(i).Cells(summary_row, 10) = price_diff
            ActiveWorkbook.Worksheets(i).Cells(summary_row, 11) = percent_price_diff
            ActiveWorkbook.Worksheets(i).Cells(summary_row, 12) = volume
            
            ' Assign green fill to tickers whose price has increased, red fill to decreases.
            If ActiveWorkbook.Worksheets(i).Cells(summary_row, 10) < 0 Then
                ActiveWorkbook.Worksheets(i).Cells(summary_row, 10).Interior.ColorIndex = 3
            Else
                ActiveWorkbook.Worksheets(i).Cells(summary_row, 10).Interior.ColorIndex = 4
            End If
            
            ' Format percentage change column as percentage.
            ActiveWorkbook.Worksheets(i).Cells(summary_row, 11).NumberFormat = "0.00%"

            ' Reset running volume count to zero.
            volume = 0

            ' Move to next summary table row.
            summary_row = summary_row + 1

        ' Collect data for the first row of each ticker.
        ElseIf ActiveWorkbook.Worksheets(i).Cells(Row, 1) <> ActiveWorkbook.Worksheets(i).Cells(Row - 1, 1) Then
            volume = volume + ActiveWorkbook.Worksheets(i).Cells(Row, 7)
            open_price = ActiveWorkbook.Worksheets(i).Cells(Row, 3)

        ' Collect data for middle rows of each ticker.
        Else
            volume = volume + ActiveWorkbook.Worksheets(i).Cells(Row, 7)

        End If

    Next Row

    ' Create summary table of maximums.
    ActiveWorkbook.Worksheets(i).Cells(1, 16) = "Ticker"
    ActiveWorkbook.Worksheets(i).Cells(1, 17) = "Value"
    ActiveWorkbook.Worksheets(i).Cells(2, 15) = "Greatest % Increase"
    ActiveWorkbook.Worksheets(i).Cells(3, 15) = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(i).Cells(4, 15) = "Greatest Total Volume"

    ' Find requested max/min values in the reference summary table.
    max_percent = ActiveWorkbook.Worksheets(i).Application.WorksheetFunction.Max(Range("k:k"))
    min_percent = ActiveWorkbook.Worksheets(i).Application.WorksheetFunction.Min(Range("k:k"))
    max_volume = ActiveWorkbook.Worksheets(i).Application.WorksheetFunction.Max(Range("l:l"))

    ' Pull tickers into summary table of maximums by finding the max/mins in the reference summary table.
    For j = 2 To 290
        If ActiveWorkbook.Worksheets(i).Cells(j, 11) = max_percent Then
            ActiveWorkbook.Worksheets(i).Cells(2, 16) = ActiveWorkbook.Worksheets(i).Cells(j, 9).Value
        ElseIf ActiveWorkbook.Worksheets(i).Cells(j, 11) = min_percent Then
            ActiveWorkbook.Worksheets(i).Cells(3, 16) = ActiveWorkbook.Worksheets(i).Cells(j, 9).Value
        Else
        End If

        If ActiveWorkbook.Worksheets(i).Cells(j, 12) = max_volume Then
            ActiveWorkbook.Worksheets(i).Cells(4, 16) = ActiveWorkbook.Worksheets(i).Cells(j, 9).Value
        Else
        End If

    ' Put max/min values in summary table of maximums.
    ActiveWorkbook.Worksheets(i).Cells(2, 17) = max_percent
    ActiveWorkbook.Worksheets(i).Cells(3, 17) = min_percent
    ActiveWorkbook.Worksheets(i).Cells(4, 17) = max_volume
    
    ' Format max/min percentage cells as percentages.
    ActiveWorkbook.Worksheets(i).Cells(2, 17).NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(i).Cells(3, 17).NumberFormat = "0.00%"
    
    Next j

Next i

End Sub


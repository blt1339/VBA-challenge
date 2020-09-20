Attribute VB_Name = "Module1"
Sub ThroughWorksheets():

    ' Loop through all worksheets and run stocks
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
    
        ' Freeze the top row for this worksheet
        ws.Activate
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        ' Change the zoom to 92% for this worksheet
        ActiveWindow.Zoom = 92
        
        ' Run stocks on this worksheet
        Call Stocks(ws)
    Next
    ' Make the first worksheet have the focus
    Worksheets(1).Activate
    Application.ScreenUpdating = True
End Sub
Sub Stocks(Sheet)
    ' Dimension the main loop variables
    Dim Ticker As String
    Dim OldTicker As String
    
    ' Dimension variables for collect information
    ' for each ticker as we loop
    Dim OpenYear As Double
    Dim CloseYear As Double
    Dim StockVolume As Double
    Dim TotalStockVolume As Double
    
    
    ' Dimension variables to determine the greatest information
    Dim GreatestCheckTicker As String
    Dim GreatestCheckPercentChange As Double
    Dim GreatestCheckTotalVolume As Double

    Dim GreatestIncreaseTicker As String
    Dim GreatestIncreaseNumber As Double
    Dim GreatestDecreaseTicker As String
    Dim GreatestDecreaseNumber As Double
    Dim GreatestTotalVolumeTicker As String
    Dim GreatestTotalVolumeNumber As Double
    
    ' Dimension variables for the number of rows
    ' for the main data and the output row for the
    ' summary information
    Dim NumRows As Long
    Dim OutputRow As Long
   
    ' Create and / or adjust headings and labels
    Sheet.Rows("1").VerticalAlignment = xlCenter
    Sheet.Rows("1").HorizontalAlignment = xlCenter
    
    Sheet.Columns("A:G").ColumnWidth = 8.5
    
    Sheet.Columns("I").ColumnWidth = 8.5
    Sheet.Range("I1").Value = "Ticker"
    
    Sheet.Columns("J").ColumnWidth = 13.5
    Sheet.Range("J1").Value = "Yearly Change"
    
    Sheet.Columns("K").ColumnWidth = 13.5
    Sheet.Range("K1").Value = "Percent Change"
    
    Sheet.Columns("L").ColumnWidth = 16.5
    Sheet.Range("L1").Value = "Total Stock Volume"
    
    Sheet.Columns("O").ColumnWidth = 20
    Sheet.Range("O2").Value = "Greatest % Increase"
    Sheet.Range("O3").Value = "Greatest % Decrease"
    Sheet.Range("O4").Value = "Greatest Total Volume"
    
    Sheet.Columns("P").ColumnWidth = 8.5
    Sheet.Range("P1").Value = "Ticker"
    
    Sheet.Columns("Q").ColumnWidth = 16.5
    Sheet.Range("Q1").Value = "Value"
  

   
    ' Get the rows of this worksheet
    NumRows = Sheet.Range("A1", Sheet.Range("A1").End(xlDown)).Rows.Count
    
   
    ' Initialize loop variables
    OldTicker = "old"
    OutputRow = 2
    OpenYear = 0
    CloseYear = 0
    TotalStockVolume = 0

    ' Sort the data by ticker by date
    ' Once this is done we can get the beginning price from first
    ' row for the ticker and the closing price from the last row for the ticker
    With Sheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SetRange Sheet.Range("A1:G" & NumRows)
        .Header = xlYes
        .Apply
    End With
    
    ' Loop through detail rows processing the data for each Ticker
    For Row = 2 To NumRows
        Ticker = Sheet.Cells(Row, 1).Value
        If Ticker <> OldTicker Then
            
            ' If we are not on the first row of data than
            ' output info for previous Ticker
            If OldTicker <> "old" Then
                CloseYear = CloseValue
                
                ' Output the line for the previous Ticker
                Call BuildOutputRow(Sheet, OutputRow, OldTicker, OpenYear, CloseYear, TotalStockVolume)
                    
                 ' Update the Greatest information if needed
                 Call CheckForGreatestChange(Sheet, OutputRow)
                 
                 ' Increase the OutputRow to be ready for next Ticker
                OutputRow = OutputRow + 1
           End If
            
            ' Start processing for the first row of a new Ticker
            OldTicker = Ticker
            OpenValue = Sheet.Cells(Row, 3).Value
            OpenYear = OpenValue
            TotalStockVolume = 0
        End If
                
        ' Grab the close value and the StockVolume and add to TotalStockVolume
        CloseValue = Sheet.Cells(Row, 6).Value
        StockVolume = Sheet.Cells(Row, 7).Value

        TotalStockVolume = TotalStockVolume + StockVolume
    Next Row
    
    ' Output the last row
    CloseYear = CloseValue
    Call BuildOutputRow(Sheet, OutputRow, OldTicker, OpenYear, CloseYear, TotalStockVolume)
                     
    ' Update the Greatest information if needed
    Call CheckForGreatestChange(Sheet, OutputRow)
End Sub
Sub BuildOutputRow(Sheet, OutputRow, RowTicker, RowOpenYear, RowCloseYear, RowTotalStockVolume):
    ' Output Row
    Sheet.Cells(OutputRow, 9).Value = RowTicker
    Sheet.Cells(OutputRow, 10).Value = RowCloseYear - RowOpenYear
    
    ' Add conditional formatting
    With Sheet.Cells(OutputRow, 10).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        .Interior.Color = RGB(57, 255, 20)
    End With
                
    With Sheet.Cells(OutputRow, 10).FormatConditions.Add(xlCellValue, xlLess, "=0")
        .Interior.Color = RGB(255, 0, 0)
    End With
                
    ' Calculate the Percent Change being careful not to divide by 0
    If RowOpenYear = 0 Then
        Sheet.Cells(OutputRow, 11).Value = 0
    Else
        Sheet.Cells(OutputRow, 11).Value = (RowCloseYear - RowOpenYear) / RowOpenYear
    End If
    Sheet.Cells(OutputRow, 11).NumberFormat = "0.00%"
    Sheet.Cells(OutputRow, 12).Value = RowTotalStockVolume
End Sub
Sub CheckForGreatestChange(Sheet, OutputRow)
    
    ' Grab all of the data we need to review to see if we need to update
    GreatestCheckTicker = Sheet.Cells(OutputRow, 9).Value
    GreatestCheckPercentChange = Sheet.Cells(OutputRow, 11).Value
    GreatestCheckTotalVolume = Sheet.Cells(OutputRow, 12).Value
    
    GreatestIncreaseTicker = Sheet.Range("P2").Value
    GreatestIncreaseNumber = Sheet.Range("Q2").Value
    
    GreatestDecreaseTicker = Sheet.Range("P3").Value
    GreatestDecreaseNumber = Sheet.Range("Q3").Value
    
    GreatestTotalVolumeTicker = Sheet.Range("P4").Value
    GreatestTotalVolumeNumber = Sheet.Range("Q4").Value
    
    ' See if we need to change the Greatest Increase information
    If GreatestCheckPercentChange > 0 Then
        If GreatestIncreaseTicker = "" Then
            GreatestIncreaseTicker = GreatestCheckTicker
            GreatestIncreaseNumber = GreatestCheckPercentChange
        Else
            If GreatestCheckPercentChange > GreatestIncreaseNumber Then
                GreatestIncreaseTicker = GreatestCheckTicker
                GreatestIncreaseNumber = GreatestCheckPercentChange
            End If
        End If
    End If
    
    ' See if we need to change the Greatest Decrease information
    If GreatestCheckPercentChange < 0 Then
        If GreatestDecreaseTicker = "" Then
            GreatestDecreaseTicker = GreatestCheckTicker
            GreatestDecreaseNumber = GreatestCheckPercentChange
        Else
            If GreatestCheckPercentChange < GreatestDecreaseNumber Then
                GreatestDecreaseTicker = GreatestCheckTicker
                GreatestDecreaseNumber = GreatestCheckPercentChange
            End If
        End If
    End If
 
    ' See if we need to change the Greatest Total Volume information
    If GreatestTotalVolumeTicker = "" Then
        GreatestTotalVolumeTicker = GreatestCheckTicker
        GreatestTotalVolumeNumber = GreatestCheckTotalVolume
    Else
        If GreatestCheckTotalVolume > GreatestTotalVolumeNumber Then
            GreatestTotalVolumeTicker = GreatestCheckTicker
            GreatestTotalVolumeNumber = GreatestCheckTotalVolume
        End If
    End If

    ' Update the worksheet with current greatest information
    Sheet.Range("P2").Value = GreatestIncreaseTicker
    Sheet.Range("Q2").NumberFormat = "0.00%"
    Sheet.Range("Q2").Value = GreatestIncreaseNumber
    
    Sheet.Range("P3").Value = GreatestDecreaseTicker
    Sheet.Range("Q3").NumberFormat = "0.00%"
    Sheet.Range("Q3").Value = GreatestDecreaseNumber
    
    Sheet.Range("P4").Value = GreatestTotalVolumeTicker
    Sheet.Range("Q4").Value = GreatestTotalVolumeNumber
 End Sub

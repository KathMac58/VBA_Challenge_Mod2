Code:
Sub VBAStockData()

'Set Variables and Data Types

Dim lastrow As Double
Dim ClosingPrice As Double
Dim OpeningPrice As Double

'Set Variables and Data Types for new columns
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim MinValue As Double
Dim MaxValue As Double
Dim MaxVol As Double
Dim MinMax As Range
Dim ws As Worksheet

'Apply code to each worksheet/year in workbook
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Set summary table locations and column headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Manage results in summary table; start writing at row 2 since row 1 will be the column headers
Dim Summary_Table As Integer
Summary_Table = 2

'Look at data in cell rows 2 through 753001; instead of typing last row number, tell program to count number of rows in column
lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

'Set opening price. Remaining opening ticker prices will be determined in conditional loop below.
OpeningPrice = Cells(2, 3).Value

'Create loop to look through each of the rows by ticker name

    For i = 2 To lastrow

    'Sort ticker names, remove duplicates, and list in Col I
    
    'In Column A, check ticker of each row, once it changes...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                           
            'Set ticker name and print name to Col I
            Ticker = Cells(i, 1).Value
            Range("I" & Summary_Table).Value = Ticker
            
            'Add total volume for each ticker trade then print total volume to Col L
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            Range("L" & Summary_Table).Value = TotalStockVolume
              
            'Pull closing price for each ticker name, calculate yearly change by subtracting set Open Price from Closing Price, then print calculation to Col J
            ClosingPrice = Cells(i, 6).Value
            YearlyChange = (ClosingPrice - OpeningPrice)
            Range("J" & Summary_Table).Value = YearlyChange
                        
            'Calculate percent change of each ticker symbol for the year (Yearly change from above divided by Opening Price), change to percentage format, and print in Col K
            PercentChange = YearlyChange / OpeningPrice
            Range("K" & Summary_Table).Value = PercentChange
            Range("K" & Summary_Table).NumberFormat = "0.00%"
        
            'Reset opening price for next ticker symbol
            OpeningPrice = Cells(i + 1, 3)
        
            'Reset ticker total
            TotalStockVolume = 0
    
            'Add one to summary table row/re-set row counter so answers appear in rows and not calculated/listed in 1 cell
            Summary_Table = Summary_Table + 1
            
        
        'If the cell following a row has the same ticker, add to the total
        
        Else
     
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    
        End If
    
    Next i

'Reset last row to look at Col 10 instead of 1

lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 10).End(xlUp).Row
Debug.Print lastrow


'Create loop to look through each row of column J and format values < 0 to be red and those > 0 to be green
    For i = 2 To lastrow

        If Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
                
        Else
            Cells(i, 10).Interior.ColorIndex = 4
                
        End If
        
    Next i


'Find the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume

Set MinMax = Range("K:K")

    'Find the Greatest % Decrease - Pull the min value, change it to a % format, and enter it in Cell(3,17)
    MinValue = Application.WorksheetFunction.Min(MinMax)
    Cells(3, 17).Value = MinValue
    Cells(3, 17).NumberFormat = "0.00%"
    
    'Loop through ticker to find associated ticker symbol for Min value
    For i = 2 To lastrow
    
        'Check if the ticker matches the Min Value
        If Cells(i, 11).Value = MinValue Then
        
        'Retrieve the ticker associated with the Min Value and enter the symbol into Cell(3, 16)
        Cells(3, 16).Value = Cells(i, 9).Value
    
        End If
        
    Next i
        
    'Find the max value in Col K, change it to a percent format, and enter it in Cell(2, 17)
    MaxValue = Application.WorksheetFunction.Max(MinMax)
    Cells(2, 17).Value = MaxValue
    Cells(2, 17).NumberFormat = "0.00%"
    
    'Loop through ticker to find associated ticker symbol for Max value
    For i = 2 To lastrow
    
        'Check if the ticker matches the Max Value
        If Cells(i, 11).Value = MaxValue Then
        
        'Retrieve the ticker associated with the Max Value and enter the symbol into Cell(2, 16)
        Cells(2, 16).Value = Cells(i, 9).Value
    
        End If
        
    Next i

Set MinMax = Range("L:L")
    
    'Find the greatest total volume and enter it in Cell(4, 17)
    MaxVol = Application.WorksheetFunction.Max(MinMax)
    Cells(4, 17).Value = MaxVol
           
    'Loop through ticker to find associated ticker symbol for greatest volume
    For i = 2 To lastrow
    
        'Check if the ticker matches the greatest volume
        If Cells(i, 12).Value = MaxVol Then
        
        'Retrieve the ticker associated with the greatest volume and enter the symbol into Cell(4, 16)
        Cells(4, 16).Value = Cells(i, 9).Value
    
        End If
        
    Next i

Next ws
 
End Sub

The goal of this project was to use VBA scripting to analyze the data of different stocks.  The dataset includes three years of data, each on a different worksheet that the code will work on reguardless of column length.  Included in this repo is this README which contains a description of the project and the VBA code, a .vbs file with the code that can be run, and a folder with the Excel Macro-Enabled Worksheet containing the data that the code was run on, which unfortunately can not be viewed on github and has to be downloaded to view. Because the file cannot be viewd here, a before and an after screenshot will be included at the end of this README. 

What the code does:

- Grab the total amount of volume each stock had for the year and display it with the corresponding ticker symbol.

- Grab the yearly change from the stock open to the year close and display that number and the percent change.

- Check to make sure there is not a divide by 0 error.

- Locate and display the stock with the greatest % increase, the stock with the greatest % decrease and stock with the greatest total volume for the year.

- Repeate for each worksheet

- Also used conditional formating to color a positive yearly change in green and a negative change in red.

Before and After Screenshots:

Before:
![Alt text](Before_screen.png "After code has run")

After:
![Alt text](After_screen.png "After code has run")


Code:
Sub StockCalc():
    'Set variables to loop run the code on each worksheet
    Dim WSCount As Integer
    Dim x As Worksheet
    Dim Workbook As Workbook
        
    'Set the number of worksheets in a variable to use
    WSCount = ActiveWorkbook.Worksheets.Count
    'MsgBox (WSCount)
    For Each x In Worksheets
    
       x.Activate
        'Set Variable for for the Ticker Symbols
        Dim TickerSym As String
        
        'Set Variable for the total volume of the ticker
        Dim TickerTotal As Double
        TickerTotal = 0
        
        'Create a variable to set the position of the row
        Dim Position As Integer
        Position = 2
        
        
        Dim openV As Double
        Dim closeV As Double
        Dim YearCh As Double
        Dim PerCh As Double
        
        'Set titles on the spreadsheet
        Range("I1").Value = "Ticker Symbol"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Volume"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        'Set the value of LastRow to equal the last row in the sheet
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set the value for where a stock opens
        openV = Range("C2").Value
        'Use a For loop to loop through each row
        For i = 2 To LastRow
            
            'Check if the Ticker Symbol in two cells is differnet
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'MsgBox (Cells(i, 1).Value)
                
                'Set value for where a stock closes
                closeV = Cells(i, 6).Value
                'MsgBox (Str(closeV) + Str(openV))
                
                'Put the value of stock change in the spreadsheet
                YearCh = closeV - openV
                Cells(Position, 10).Value = YearCh
                
                If openV = 0 Then
                    Cells(Position, 11).Value = "0"
                Else
                    PerCh = YearCh / openV
                    Cells(Position, 11).Value = PerCh
                End If
                
                
                'Set the Ticker Symbol
                TickerSym = Cells(i, 1).Value
                
                'Place the Ticker Symbol
                Range("I" & Position).Value = TickerSym
                
                'Add to the value of the Stocks total volume
                TickerTotal = TickerTotal + Cells(i, 7).Value
                
                'Set the total volume onto the Spreadsheet
                Range("L" & Position).Value = TickerTotal
                
                'Increment the Row where the variables will be placed
                Position = Position + 1
                
                'Reset the TickerTotal
                TickerTotal = 0
            
                'Change the open value to that of the next stock sym
                openV = Cells(i + 1, 3).Value
                
            'If the Ticker Symbols are the same
            Else
                TickerTotal = TickerTotal + Cells(i, 7).Value
                
            End If
                   
        Next i
        
        
        'Range("K:K").NumberFormat = "0.00"
        
        'Set variables to find the geatest Increase, Decrease, and Volume
        Dim great As Double
        great = 0
        Dim greatTic As String
        Dim Dec As Double
        Dec = 0
        Dim DecTic As String
        Dim HighVol As Double
        HighVol = 0
        Dim HVTic As String
        
        For j = 2 To (Position - 1)
        
            'Gives the cell a color based on it's value
            If Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
                
            Else
                Cells(j, 10).Interior.ColorIndex = 4
                
            End If
            
            'Styles the cells for % and 2 decimal places
            Cells(j, 11).Style = "Percent"
            Cells(j, 11).NumberFormat = "0.00%"
        
            'Check which stock has the greatest % increase
            If Cells(j, 11).Value > great Then
                great = Cells(j, 11).Value
                greatTic = Cells(j, 9).Value
            End If
                
            'Check which stock has the greatest % decrease
            If Cells(j, 11).Value < Dec Then
                Dec = Cells(j, 11).Value
                DecTic = Cells(j, 9).Value
                
            End If
            
            'Check which stock has the greatest volume
            If Cells(j, 12).Value > HighVol Then
                HighVol = Cells(j, 12).Value
                HVTic = Cells(j, 9).Value
            End If
            
        
        Next j
        
        'Places the values in the spreadsheet
        Range("O2").Value = greatTic
        Range("O3").Value = DecTic
        Range("P2").Value = great
        Range("P3").Value = Dec
        Range("P2:P3").Style = "Percent"
        Range("O4").Value = HVTic
        Range("P4").Value = HighVol
   
   Next
End Sub


Attribute VB_Name = "Module1"
Sub SummaryTableStocks()

    'Declare Current as worksheet object variable
    Dim ws As Worksheet
    
    ' Loop through all of the worksheets in teh active workbook.
    For Each ws In Worksheets
    
        'Activate the current worksheet
        ws.Activate

      'Assign Values to Summary Table headings
       ws.Range("I1").Value = "Ticker"
       ws.Range("I1").Font.Bold = True
       
       ws.Range("J1").Value = "Yearly Change"
       ws.Range("J1").Font.Bold = True
       ws.Range("J1").EntireColumn.AutoFit
       
       ws.Range("K1").Value = "Percent Change"
       ws.Range("K1").Font.Bold = True
       ws.Range("K1").EntireColumn.AutoFit
       
       ws.Range("L1").Value = "Volume"
       ws.Range("L1").Font.Bold = True
    
       'Set variables for the challenge: Ticker, yearly change, percent change, TotalVolume, Open_Value and Close_Value
       Dim Ticker As String
       Dim Yearly_Change As Double
       Dim Percent_Change As Double
       Dim TotalVolume As LongLong
       Dim Open_Value As Double
       Dim Close_Value As Double
       
      'Create TotaVolume counter
       TotalVolume = 0
       
       'Set location of Summary Table
       Dim Summary_Table_Row As Integer
       Summary_Table_Row = 2
       
       'Declare last row
       lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
       'Loop through tickers
       For Row = 2 To lastrow
       
           'Retrieve Open value only if Open value has not been set for this ticker
           If ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
            Open_Value = ws.Cells(Row, 3).Value
            TotalVolume = ws.Cells(Row, 7).Value
                
          'Check if we are still within the same ticker, if it is not then...
           ElseIf ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
        
               'Retrieve ticker
               Ticker = ws.Cells(Row, 1).Value
            
              'Retrieve Close_Value
               Close_Value = ws.Cells(Row, 6).Value
               
               'Calculate Yearly_Change
               Yearly_Change = Close_Value - Open_Value
               
               'Calculate percent change
               Percent_Change = Yearly_Change / Open_Value
           
               'Add up Volumes
               TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
           
               'Populate summary table with Ticker
               ws.Range("I" & Summary_Table_Row).Value = Ticker
           
               'Populate summary Table with Yearly_Change
               ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
               
                   'format table to highlight Yearly Change below and above 0 values
                   If Yearly_Change > 0 Then
                   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                   
                       Else
                       ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                   End If
           
               'Populate summary table with Percent_Change
               ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                   
                   'format table to highlight Percent Changes below and above 0 values
                   If Percent_Change > 0 Then
                   ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                   
                       Else
                       ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                   End If
                        
               'Format percent change with %
               ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
           
               'Populate TotalVolume
               ws.Range("L" & Summary_Table_Row).Value = TotalVolume
               ws.Range("L" & Summary_Table_Row).NumberFormat = "General"
               ws.Range("L" & Summary_Table_Row).EntireColumn.AutoFit
           
               'Add one row to the summary table
               Summary_Table_Row = Summary_Table_Row + 1
           
               'Reset the TotalVolume
               TotalVolume = 0
               Close_Value = 0
               
           'If the cell immediately following a row is the same Ticker ...
           Else
                   
          'Keep adding up the Volumes
           TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
                   
           End If
           
       Next Row
       
       'Assign Col headers to second table
       ws.Range("P1").Value = "Ticker"
       ws.Range("P1").Font.Bold = True
       
       ws.Range("Q1").Value = "Value"
       ws.Range("Q1").Font.Bold = True
       
       'Assing Row names to rows of second table
       ws.Range("O2").Value = "Greatest % Increase"
       ws.Range("O2").Font.Bold = True
       ws.Range("O2").EntireColumn.AutoFit
       
       ws.Range("O3").Value = "Greatest % Decrease"
       ws.Range("O3").Font.Bold = True
       ws.Range("O3").EntireColumn.AutoFit
       
       ws.Range("O4").Value = "Greatest Total Volume"
       ws.Range("O4").Font.Bold = True
       ws.Range("O4").EntireColumn.AutoFit
       
       'Declare variable and assign value
       Dim Greatest_Inc As Double
       Dim Greatest_Dec As Double
       Dim GreatestTL As LongLong
       
       'Declare last row
       lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       'Loop thru percent change column
       For Row = 2 To lastrow
       
           If ws.Cells(Row + 1, 11).Value > ws.Cells(Row, 11).Value And ws.Cells(Row + 1, 11).Value > Greatest_Inc Then
               Greatest_Inc = ws.Cells(Row + 1, 11).Value
               Ticker = ws.Cells(Row + 1, 9).Value
               
               'Populate Table
               ws.Range("P2").Value = Ticker
               ws.Range("Q2").Value = Greatest_Inc
               ws.Range("Q2").NumberFormat = "0.00%"
               
               'Find and retrieve Greatest Decrease
               
               ElseIf ws.Cells(Row + 1, 11).Value < ws.Cells(Row, 11).Value And ws.Cells(Row + 1, 11).Value < Greatest_Dec Then
               Greatest_Dec = ws.Cells(Row + 1, 11).Value
               Ticker = ws.Cells(Row + 1, 9).Value
               
               'Populate second table with greatest decrease %
               ws.Range("P3").Value = Ticker
               ws.Range("Q3").Value = Greatest_Dec
               ws.Range("Q3").NumberFormat = "0.00%"
               
               'If the above is not correct then maintain value of Greatest_Inc...
               Else
               Greatest_Inc = Greatest_Inc
               Greatest_Dec = Greatest_Dec
          
           End If
           
           'Find greatest total volume
                   
           If ws.Cells(Row + 1, 12).Value > ws.Cells(Row, 12).Value And ws.Cells(Row + 1, 12) > GreatestTL Then
            GreatestTL = ws.Cells(Row + 1, 12).Value
            Ticker = ws.Cells(Row + 1, 9).Value
                    
           'Populate second table
           ws.Range("P4").Value = Ticker
           ws.Range("Q4").Value = GreatestTL
           ws.Range("Q4").NumberFormat = "General"
           ws.Range("Q4").EntireColumn.AutoFit
           
           'If the above is not true, carry the last found GreatestTL
               Else
               GreatestTL = GreatestTL
               
           End If
    
       Next Row
       
    Yearly_Change = 0
    Percent_Change = 0
    TotalVolume = 0
    Open_Value = 0
    Close_Value = 0
    Greatest_Inc = 0
    Greatest_Dec = 0
    GreatestTL = 0

Next ws

End Sub








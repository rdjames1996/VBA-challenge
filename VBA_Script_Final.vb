Sub Stock_data_Final()

'Define all variables. This includes a Worksheet variable, because there are multiple tabs in the excel booklet.

Dim Ticker As String
Dim year_open As Double
Dim year_close As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double
Dim start_data As Integer
Dim ws As Worksheet

For Each ws In Worksheets


    'these are the column headers.
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Start_data = row the recording of data will start in the summary table,
    'ticker_open = the opening value of the ticker at the beginning of the year
    'Total_Stock_Volume = 0 because it needs to be reset for every ticker
    start_data = 2
    ticker_open = 1
    Total_Stock_Volume = 0

        'Create this variable so the loop automatically identifies the last row, where it is supposed to stop.
        Dim lastrow As Long
        lastrow = Range("A2").End(xlDown).Row
        
       'Loop through each ticker to determine the yearly change, percent change, and total stock volume.
        For i = 2 To lastrow

            'If the ticker symbol changes, then...

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            
            'ticker = new ticker string in summary table
            Ticker = ws.Cells(i, 1).Value

            'Make sure loop is recording year opening value of this ticker
            ticker_open = ticker_open + 1

            year_open = ws.Cells(ticker_open, 3).Value
            year_close = ws.Cells(i, 6).Value

            'This loop will find the total stock volume

            For j = ticker_open To i

                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j

            
            ' This if statement determines the percent change data
            If year_open = 0 Then

                Percent_Change = year_close

            Else
            
                ' percent change = (year close - year open) / year open.
                Yearly_Change = year_close - year_open

                Percent_Change = Yearly_Change / year_open

            End If
    

            
            ' Creating the Summary table. data will track starting on row 2
            ws.Cells(start_data, 9).Value = Ticker
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change
            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = Total_Stock_Volume

            
            start_data = start_data + 1

           
            ' Make Sure Each value is set at zero before recording in summary table
            
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
            ticker_open = i

        End If


    Next i
    
    Dim lastrowcolor As Long
    lastrowcolor = Range("J2").End(xlDown).Row
    'Conditional Formatting. Color Index 4 = Green, Color Index 3 = Red.'
    For j = 2 To lastrowcolor
        If ws.Cells(j, 10) > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
        
        Else
        
        ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    
    
Next ws

    
    
End Sub

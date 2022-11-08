Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

' declare worksheet variable to loop through each worksheet and start the for loop..
Dim ws As Worksheet
For Each ws In Worksheets

    ' set an initial variable to hold Tickers
    Dim Ticker_Symbol As String
    
    ' set a variable to store value of lastrow in the worksheet
    Dim LastRow As Double
    
    ' set a variable to store row number for output table so as to maintain a counter to fill output table sequentially, initiliaze it as '2' to fill data in output table
    Dim OutputRow As Double
    OutputRow = 2
    
    ' set a variable to store row number when ticker symbol changes and initialise it as '2' as first data point is in row 2
    Dim NewTickerRow As Double
    NewTickerRow = 2
    
    ' set new variable to store total of stock volume
    Dim StockVolume As Double
    ' StockVolume = ws.Cells(2, 7).Value
    StockVolume = 0
    
    ' enter new column heads for Ticker, Yearly Change, Percent Change, Total Stock Volume
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' calcluate Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
        ' loop through all stock data
        For i = 2 To LastRow
        
            ' Check if we are still within the same ticker, if we are not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' print ticker name in the output table under column Ticker
                Ticker_Symbol = ws.Cells(i, 1).Value
                ws.Cells(OutputRow, 9).Value = Ticker_Symbol
                
                ' calculate yearly change and percent change
                ws.Cells(OutputRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(NewTickerRow, 3).Value
                ws.Cells(OutputRow, 11).Value = (ws.Cells(i, 6).Value / ws.Cells(NewTickerRow, 3).Value) - 1
                
                ' calculate stockvolume and print in output table
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                ws.Cells(OutputRow, 12).Value = StockVolume
                
                ' add one to both row counters
                OutputRow = OutputRow + 1
                NewTickerRow = i + 1
                
                StockVolume = 0 ' reset stock volume to zero before condition is met again
            
            Else: StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
       
       
     ' Conditional formatting routine for output table
     
     ' declare a new variable to store last row number for output table
     Dim LastRow_OutputTable As Integer
     
     ' calculate last row of output table
     LastRow_OutputTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
               ' Loop through output table values
               For j = 2 To LastRow_OutputTable
                   
                      ' check for values of column Yearly Change.......
                        If ws.Cells(j, 10).Value > 0 Then
                       
                               ws.Cells(j, 10).Interior.ColorIndex = 4 'color cell to green if yearly change is positive
                           
                        Else: ws.Cells(j, 10).Interior.ColorIndex = 3 'color cell to  red  if yearly change is negative
                       
                       End If
                       
               Next j
               
               ' Loop through output table values
               For k = 2 To LastRow_OutputTable
                   
                       ws.Cells(k, 11).NumberFormat = "0.00%"   ' format as percentage
                       
                       ' check for values of column Percent Change.......
                       If ws.Cells(k, 11).Value > 0 Then
                       
                       ws.Cells(k, 11).Interior.ColorIndex = 4 'color cell to green if perecnt change is positive
                       
                       Else: ws.Cells(k, 11).Interior.ColorIndex = 3  'color cell to  red  if precent change is negative
                       
                       End If
                       
               Next k
            
   ' routine to implement bonus to find greatest Greatest % Increase, Greatest % Decrease and Greatest Total Volume
   
   'give headers
   ws.Range("p1").Value = "Ticker"
   ws.Range("q1").Value = "Value"
   ws.Range("o2").Value = "Greatest % Increase"
   ws.Range("o3").Value = "Greatest % Decrease"
   ws.Range("o4").Value = "Greatest Total Volume"
    
   Dim GTV As LongLong 'variable to store value of greatest total volume
   Dim GPI As Double 'variable to store value of greatest percent increase
   Dim GPD As Double 'variable to store value of greatest percent decrease
   
   GPI = 0 '  initialise  as zero
   GPD = 0 '  initialise as zero
   GTV = 0 ' initialise GTV as zero
   
       ' Loop through output table to find greatest volume, greatest % chnages
        For m = 2 To LastRow_OutputTable
             
             If ws.Cells(m, 12).Value > GTV Then ' compare every cell in this column with GTV
             
                GTV = ws.Cells(m, 12).Value ' update GTV as current cell if condition met
                ws.Range("q4").Value = GTV ' print GTV in the table in the corrresponding cell
                ws.Range("p4").Value = ws.Cells(m, 9).Value ' print corresponding ticker
                
             End If
             
             If ws.Cells(m, 11).Value > GPI Then ' compare every cell in this column with GPI
             
                GPI = ws.Cells(m, 11).Value ' update GPI as current cell if condition met
                ws.Range("q2").Value = FormatPercent(GPI) ' print GPI in the table in the corrresponding cell and format as a percentage
                ws.Range("p2").Value = ws.Cells(m, 9).Value ' print corresponding ticker
                
             
             End If
             
             If ws.Cells(m, 11).Value < GPD Then ' compare every cell in this column with GPD
             
                GPD = ws.Cells(m, 11).Value ' update GPD as current cell if condition met
                ws.Range("q3").Value = FormatPercent(GPD) ' print GPD in the table in the corrresponding cell and format as a percentage
                ws.Range("p3").Value = ws.Cells(m, 9).Value ' print corresponding ticker
                
        
             End If

             
        Next m
              
Next ws


End Sub

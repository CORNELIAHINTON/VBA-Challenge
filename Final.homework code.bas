Attribute VB_Name = "Module1"
Sub stockdata()

    'Loop through all worksheets in the workbook
    For Each ws In Worksheets
        
        'variable to hold the worksheet name
        Dim worksheetname As String
        
        worksheetname = ws.Name 'to store the name of the worksheet
        
        'added ticker to column i
        ws.Range("I1").EntireColumn.Insert
        
        'added the ticker name to the desired column
        ws.Range("I1").Value = "Ticker"
        
        'added yearly change to column j
        ws.Range("J1").EntireColumn.Insert
        
        'added the yearly change name to the desired column
        ws.Range("J1").Value = "Yearly Change"
        
        'add percent change name to column k
        ws.Range("K1").EntireColumn.Insert
        
        'add percent change name to the desired column
        ws.Range("K1").Value = "Percent Change"
        
        'add total stock volume name to column l
        ws.Range("L1").EntireColumn.Insert
        
        'add total stock volume name to the desired column
        ws.Range("L1").Value = "Total Stock Volume"
        
         'summary table for tickers
        Dim SummaryTable As Integer
        SummaryTable = 2
        
        'variable for ticker
        Dim tickername As String
        
         'variable for yearly change
        Dim yearlychange As Double
        yearlychange = 0
        
        'variable for percent change
        Dim percentchange As Double
        percentchange = 0
        
        'variable for total stock volume
        Dim totalstockvolume As LongLong
        totalstockvolume = 0
        
        'get the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'variable for open price
        Dim openprice As Double
        
        'variable for close price
        Dim closeprice As Double
        
        
        'check through all the tickers
        For Row = 2 To lastrow
        
         'calculate total stock volume
        totalstockvolume = totalstockvolume + ws.Range("G" & Row).Value
        
        If (Cells(Row, 1).Value <> Cells(Row - 1, 1).Value) Then
            'set open price
            openprice = ws.Range("C" & Row).Value
            
        End If
        
        'check to see if we are still on the same ticker name
        'if not do the following
        If (Cells(Row, 1).Value <> Cells(Row + 1, 1).Value) Then
             
             'set close price
             closeprice = ws.Range("F" & Row).Value
             
            'set(reset) the ticker name
            tickername = ws.Range("A" & Row).Value
                
            'add the ticker name to the ticker column in the summary table
             ws.Range("I" & SummaryTable).Value = tickername
             
             'calculate the yearly change from opening price to closing price
                yearlychange = closeprice - openprice
                
                'add yearly change total into column J
                    ws.Range("J" & SummaryTable).Value = yearlychange
                    
                    'highlight positive change in green or negative change in red
                    If yearlychange >= 0 Then
                        ws.Range("J" & SummaryTable).Interior.Color = vbGreen
                        
                        Else
                            ws.Range("J" & SummaryTable).Interior.Color = vbRed
                            
                    End If
                    
                'reset yearly change total for next ticker
                
             'calculate the percent change from opening price to closing price
             If openprice <> 0 Then
             percentchange = (closeprice - openprice) / openprice
             
             Else
             percentchange = closeprice
             
             End If
             
                'add percent change total into column K
                ws.Range("K" & SummaryTable).Value = percentchange
                    
                'reset percent change total for next ticker
            
                'add total stock volume total into column L
                ws.Range("L" & SummaryTable).Value = totalstockvolume
                
                'reset total stock volume for next ticker
                totalstockvolume = 0
                
                'adding next ticker value to summary table
               SummaryTable = SummaryTable + 1
               
             
            
            End If


         Next Row

    Next ws
    
    
End Sub



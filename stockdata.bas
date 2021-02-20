Attribute VB_Name = "Module1"
Sub stockdata()

Dim tickername As String

Dim totalvolume As LongLong

Dim closingprice, openingprice, yearlychange As Double

Dim percentchange As Variant

Dim summary_table_row As Integer
'summary_table_row = 2

Dim min, max As Double
Dim maxvol As LongLong
Dim tag As String
       
    '  loop through all sheets
    For Each ws In Worksheets
    
        summary_table_row = 2
        closingprice = 0
        yearlychange = 0
        percentchange = 0
        max = 0
        min = 0
        maxvol = 0
        
        '  find the last row of the sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        '  opening price of the first ticker
        openingprice = ws.Cells(2, 3).Value
                      
        '  summary table header
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        '  autofit and number formatting
        ws.Range("J:J").NumberFormat = "0.00"
        ws.Range("K:K").NumberFormat = "0.00%"
        
        '  loop through all ticker
        For I = 2 To lastrow
                        
            '  check if we're still in the same ticker
            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                
                '  set the ticker
                 tickername = ws.Cells(I, 1).Value
                                             
                '  get closing price at the end of the year
                closingprice = ws.Cells(I, 6).Value
                                
                '  calculate yearly change
                yearlychange = closingprice - openingprice
                
                '  calculate percent change
                percentchange = (closingprice - openingprice) / closingprice
                
                '  add to the total volume
                totalvolume = totalvolume + ws.Cells(I, 7).Value
                            
                '  print the ticker name to the summary table
                ws.Range("I" & summary_table_row).Value = tickername
                  
                '  print yearly change
                ws.Range("J" & summary_table_row).Value = yearlychange
                
                '  print percent change
                ws.Range("K" & summary_table_row).Value = percentchange
                  
                 '  print total stock volume
                 ws.Range("L" & summary_table_row).Value = totalvolume
                  
                '  add one row to the summary table row
                summary_table_row = summary_table_row + 1
                
                '  reset the total stock volume
                totalvolume = 0
                
                 '  get opening price at the beginning of the year for the next ticker
                openingprice = ws.Cells(I + 1, 3).Value
                
                percentchange = 0
                
            Else
                                    
                '  add to the total volume
                totalvolume = totalvolume + ws.Cells(I, 7).Value
                                                     
            End If
            
        Next I
        
            '  autofit
            ws.Range("L:L").Columns.AutoFit
                
            '  conditional formatting
            
            lastrowsummary = ws.Cells(Rows.Count, 10).End(xlUp).Row
                        
            '  loop through all rows of summary table
            For I = 2 To lastrowsummary
        
                '  negative numbers should be red
                If ws.Cells(I, 10).Value < 0 Then
                    ws.Cells(I, 10).Interior.ColorIndex = 3
                    
                Else
                    ws.Cells(I, 10).Interior.ColorIndex = 4
                
                End If
            
            Next I
    
    
'-------------------------------------------------------------------------------
'CHALLENGES
'-------------------------------------------------------------------------------
    

        '  summary table 2
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("O:O").Columns.AutoFit
        
        For J = 2 To lastrowsummary
        
            '  find highest % increase
            If ws.Cells(J, 11).Value > max Then
                max = ws.Cells(J, 11).Value
                tag = ws.Cells(J, 9).Value
                
                '  print greatest % increase to summary table 2
                ws.Range("P2") = tag
                ws.Range("Q2") = max
                
            End If
            
            '  find greatest % decrease
            If ws.Cells(J, 11).Value < min Then
                min = ws.Cells(J, 11).Value
                tag = ws.Cells(J, 9).Value
                
                 '  print greatest % decrease to summary table 2
                ws.Range("P3") = tag
                ws.Range("Q3") = min
            
            End If
               
            '  find highest total volume
            If ws.Cells(J, 12) > maxvol Then
                maxvol = ws.Cells(J, 12).Value
                tag = ws.Cells(J, 9).Value
            
                '  print highest total volume to summary table 2
                ws.Range("P4") = tag
                ws.Range("Q4") = maxvol
            
            End If
            
        Next J
        
        'format % inc and % dec in summary table 2
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q:Q").Columns.AutoFit
              
    Next ws
        

End Sub


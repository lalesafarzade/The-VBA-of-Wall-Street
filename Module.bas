Attribute VB_Name = "Module1"
Sub stock_volume()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In ThisWorkbook.Worksheets

        ' Determine the Last Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        
        ' Add the proper word to the First row Header
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest%Increase"
        ws.Cells(3, 15).Value = "Greatest%Decrease"
        ws.Cells(4, 15).Value = "GreatestTotalVolume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
        
        ' Set an initial variable for holding the ticker name
        Dim tick As String
        Dim yearly_change As Double
        Dim percent_change As Double
        
        
        

        ' Set an initial variable for holding the total volume
        V_Total = 0
        
        

        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        For i = 2 To lastrow
        
        ' determin the first ticker open value
            If i = 2 Then
                ws.Cells(2, 19).Value = ws.Cells(i, 3).Value
            End If
            
                
            

        ' Check if we are still within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
 
                ' Set the ticker name
                tick = ws.Cells(i, 1).Value

                ' Add to the V_Total
                V_Total = V_Total + ws.Cells(i, 7).Value


                ' Print the ticker in the Summary Table
                ws.Cells(Summary_Table_Row, 10).Value = tick

                ' Print the total_volume to the Summary Table
                ws.Cells(Summary_Table_Row, 13).Value = V_Total
                
              
                
                ' print the open and close values of all tickers in temporary columns
                ws.Cells(Summary_Table_Row, 18).Value = ws.Cells(i, 6).Value
                ws.Cells(Summary_Table_Row + 1, 19).Value = ws.Cells(i + 1, 3).Value
                
                ' print the yearly change: close-open values
                
                ws.Cells(Summary_Table_Row, 11).Value = ws.Cells(Summary_Table_Row, 18).Value - ws.Cells(Summary_Table_Row, 19).Value
                
                ' conditional formating: positive-->green, negative-->red
                
                If ws.Cells(Summary_Table_Row, 11).Value < 0 Then
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
                End If
                
                ' printing percent change if the open values are not zero
                If ws.Cells(Summary_Table_Row, 19).Value <> 0 Then
                    percent_change = Round(ws.Cells(Summary_Table_Row, 11).Value / ws.Cells(Summary_Table_Row, 19).Value, 4)
                Else
                    percent_change = 0
                
                End If

                ws.Cells(Summary_Table_Row, 12).Value = Format(percent_change, "percent")
                
                ' Add one to the summary tabl
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the volume_Total
                V_Total = 0
                
                

            ' If the cell immediately following a row is the same ticker...
            Else

                ' Add to the volume_Total
                V_Total = V_Total + ws.Cells(i, 7).Value

            End If

        Next i
        
        
        ' Determine the Last Row of summary table
        last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Set initial cells for holding the max_volume,max_val,min_val
        max_val = ws.Cells(2, 12).Value
        max_tick = ws.Cells(2, 10).Value
        min_val = ws.Cells(2, 12).Value
        min_tick = ws.Cells(2, 10).Value
        max_vol = ws.Cells(2, 13).Value
        max_vol_tick = ws.Cells(2, 10).Value

        'loop for finding max_val and min_val of percent change
        For ctr = 2 To last_row
            x = ws.Cells(ctr, 12).Value
            t = ws.Cells(ctr, 10).Value
            If x > max_val Then
                max_val = x
                max_tick = t
            ElseIf x < min_val Then
                min_val = x
                min_tick = t
            End If
        Next
        
        'loop for finding max of total_vol
        For ctr = 2 To last_row
            y = ws.Cells(ctr, 13).Value
            v = ws.Cells(ctr, 10).Value
            If y > max_vol Then
                max_vol = y
                max_vol_tick = v
            End If
        Next
        

        ws.Cells(2, 17).Value = Format(max_val, "percent")
        ws.Cells(3, 17).Value = Format(min_val, "percent")
        ws.Cells(4, 17).Value = max_vol
        ws.Cells(2, 16).Value = max_tick
        ws.Cells(3, 16).Value = min_tick
        ws.Cells(4, 16).Value = max_vol_tick
        

        'Adjust column width automatically
        ws.Columns("A:Z").AutoFit
        
        'deleting temporary columns
        ws.Columns("r:s").Delete
    Next ws
    MsgBox ("Done")
End Sub



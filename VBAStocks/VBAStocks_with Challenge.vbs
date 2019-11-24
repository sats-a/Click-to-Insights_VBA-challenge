Sub ticker_stats()

'Set all Dimensions

' Set an initial variable for holding the ticker name
Dim ticker As String

' Define initial variable for holding the total stock Volume for each Ticker
Dim stock_vol As Double

' Keep track of the location for each Ticker row in the Result Table
Dim result_row As Integer

' Define the index
Dim i As Long
Dim j As Long

' Define Dimensions for all the values to be calculated
Dim close_val As Double
Dim open_val As Double
Dim year_change As Double
Dim percent As Double

' Define LastRow for Data
Dim LastRow As Long

'Define Dimensions for Challenge to get %increase, %decrease & greatest total volume
Dim max_p As Double
Dim min_p As Double
Dim max_v As Double
Dim placeholder_max As Double
Dim placeholder_min As Double
Dim placeholder_vol As Double
Dim ticker_max As String
Dim ticker_min As String
Dim ticker_vol As String


' Set an initial value for holding the total stock Volume for each Ticker
stock_vol = 0

'Set initial values
stock_vol = 0
close_val = 0
open_val = 0
percentage = 0



' Loop through all sheets
    For Each ws In Worksheets

            result_row = 2
        
            ' Determine the Last Row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' Determine the Titles for the Ticker Result Table
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest total volume"
        
        
        

                'Loop through all rows
                For i = 2 To LastRow
                    
                    'Compare next row to current and if not equal then
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                            ' Set the Ticker Name
                            ticker = ws.Cells(i, 1).Value

                            ' Add to the Stock Value Total
                            stock_vol = stock_vol + ws.Cells(i, 7).Value
                            close_val = ws.Cells(i, 6).Value

                            'Calculate Yearly Change
                            year_change = close_val - open_val

                            ' Print the Ticker name in the Ticker Result Table
                            ws.Range("I" & result_row).Value = ticker

                            ' Print the Total Stock Volume to the Ticker Result Table
                            ws.Range("L" & result_row).Value = stock_vol
      
                                    'Make sure open value is not 0 to avoid division error and calculate percentage
                                    If open_val <> 0 Then
                                            percentage = (year_change / open_val)
                                    Else
                                    End If
                            
                            'Print the values and get percentage
                            ws.Range("J" & result_row).Value = year_change
                            ws.Range("K" & result_row).Value = percentage
                            ws.Range("K" & result_row).NumberFormat = "0.00%"
                            
                                    'Assign Colors
                                    If year_change > 0 Then
                                        ws.Range("J" & result_row).Interior.ColorIndex = 4
                                    ElseIf year_change < 0 Then
                                        ws.Range("J" & result_row).Interior.ColorIndex = 3
                                    Else
                                        ws.Range("J" & result_row).Interior.ColorIndex = 0
                                    End If
                        

                            ' Add one to the Ticker Result Table row
                            result_row = result_row + 1

                            'Reset the variable for new Ticker
                            stock_vol = 0
                            close_val = 0
                            open_val = 0
                            percentage = 0


                    Else

                            ' Add to the Stock Total if Ticker is still same
                             stock_vol = stock_vol + ws.Cells(i, 7).Value
                            
                    End If
            
            
                    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                            open_val = ws.Cells(i, 3).Value
            
                    End If

                Next i
                
            'This is the total number of rows of the result table - reducing by one to avoid the last empty cell
            result_row = result_row - 1
        
            'Assigning all placeholders initial value of the first row from the final Result table
            placeholder_max = ws.Range("K2").Value
            placeholder_min = ws.Range("K2").Value
            placeholder_vol = ws.Range("L2").Value
            ticker_max = ws.Range("I2").Value
            ticker_min = ws.Range("I2").Value
            ticker_vol = ws.Range("I2").Value


                'Loop through all values of Result row - Starting from row 3 as the initial value is assigned to placeholders
                For j = 3 To result_row
                    
                    'Calculating the Greatest % increase Value and corresponding Ticker
                    If ws.Range("K" & j).Value > placeholder_max Then
                        max_p = ws.Range("K" & j).Value
                        ticker_max = ws.Range("I" & j).Value
                        placeholder_max = ws.Range("K" & j).Value
                    Else
                        max_p = placeholder_max
                    End If
                    
                    'Calculating the Greatest % Decrease Value and corresponding Ticker
                    If ws.Range("K" & j).Value < placeholder_min Then
                        min_p = ws.Range("K" & j).Value
                        ticker_min = ws.Range("I" & j).Value
                        placeholder_min = ws.Range("K" & j).Value
                    Else
                        min_p = placeholder_min
                    End If
                    
                    'Calculating the Greatest total volume and corresponding Ticker
                    If ws.Range("L" & j).Value > placeholder_vol Then
                        max_v = ws.Range("L" & j).Value
                        ticker_vol = ws.Range("I" & j).Value
                        placeholder_vol = ws.Range("L" & j).Value
                    Else
                        max_v = placeholder_vol
                    End If
                    
                     
                Next j

            'Assigning Values to the the Challenge Results and converting to percentage
            ws.Range("P2") = ticker_max
            ws.Range("Q2") = max_p
            ws.Range("P3") = ticker_min
            ws.Range("Q3") = min_p
            ws.Range("P4") = ticker_vol
            ws.Range("Q4") = max_v
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
   'Next worksheet
   Next ws

End Sub




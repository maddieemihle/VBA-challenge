Attribute VB_Name = "Module1"
Sub Stock_Analysis():

    ' Establish current/first worksheet and loop through all worksheets
    Dim CurrentWs As Worksheet
    For Each CurrentWs In Worksheets
    Dim LastRow As Long
    Dim i As Long
        
' ------------------------------------------
' Define all variables
' ------------------------------------------
    
    ' Define variable for Ticker
    Dim ticker As String
        
    ' Define variable for Total Stock Volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
        
    ' Define initial variable for opening and closing price, quarterly change, and percent change
    Dim Open_Price As Double
    Open_Price = 0
    Dim Close_Price As Double
    Close_Price = 0
    Dim Quarterly_Change As Double
    Quarterly_Change = 0
    Dim Percent_Change As Double
    Percent_Change = 0
        
    ' Define variables for greatest % increase, greatest % decrease, and greatest total volume
    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    Dim Greatest_Decrease As Double
    Greatest_Decrease = 0
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0
        
    ' Establish variables for tickers of greatest % increase, greatest % decrease, and greatest total volume
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Total_Volume_Ticker As String
        
    ' -----------------------------------------
    '  Set initial variable for tracking the location of date's open and close price
    Dim Price_Row As Long
    Price_Row = 2
    
    ' Set total stock volume as 0
    Total_Stock_Volume = 0
    
    ' Establish location for tickers within summary row table
    Dim summary_row_table As Integer
    summary_row_table = 2
      
    ' -----------------------------------------
    ' Assign column headers for ticker, quarterly change, percent change, and total stock volume
    CurrentWs.Range("I1").Value = "Ticker"
    CurrentWs.Range("J1").Value = "Quarterly Change"
    CurrentWs.Range("K1").Value = "Percent Change"
    CurrentWs.Range("L1").Value = "Total Stock Volume"
    
    ' Assign column headers for greatest % increase, greatest % decrease, greatest total volume, ticker, and value
    CurrentWs.Range("O2").Value = "Greatest % Increase"
    CurrentWs.Range("O3").Value = "Greatest % Decrease"
    CurrentWs.Range("O4").Value = "Greatest Total Volume"
    CurrentWs.Range("P1").Value = "Ticker"
    CurrentWs.Range("Q1").Value = "Value"
    
' -------------------------------------------
' Perform functions to loop columns, establish ticker, quarterly change, percent change, and total stock volume.
' -------------------------------------------

      ' Retrieve last row in worksheets
        LastRow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
      
      ' For each ticker, summerize and loop for quarterly change, percent change, and total stock volume
        For i = 2 To LastRow:
      
            ' Define if ticker is equal of not equal to previous, get ticker symbol, and initate
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
               ticker = CurrentWs.Cells(i, 1).Value
               
               ' Add the total stock volume
               Total_Stock_Volume = Total_Stock_Volume + CurrentWs.Range("G" & i).Value
               
               ' Print out the ticker, total stock volume, quarterly change, and percent change in table
               CurrentWs.Range("I" & summary_row_table).Value = ticker
               CurrentWs.Range("L" & summary_row_table).Value = Total_Stock_Volume
               
               ' Perform calculation for quarterly change and percent change
               Open_Price = CurrentWs.Range("C" & Price_Row).Value
               Close_Price = CurrentWs.Range("F" & i).Value
               Quarterly_Change = Close_Price - Open_Price
               
                ' If unable to divide by 0
                  If Open_Price = 0 Then
                      Percent_Change = 0
                     Else
                         Percent_Change = Quarterly_Change / Open_Price
                  End If
                 
                 ' Print the values of quarterly change and percent change
                  CurrentWs.Range("J" & summary_row_table).Value = Quarterly_Change
                  CurrentWs.Range("K" & summary_row_table).Value = Percent_Change
                  CurrentWs.Range("K" & summary_row_table).NumberFormat = "0.00%"
                  
                        ' Use Conditional Formatting for quarterly change
                        If (Quarterly_Change > 0) Then
                            CurrentWs.Range("J" & summary_row_table).Interior.ColorIndex = 4
                        ElseIf (Quarterly_Change < 0) Then
                            CurrentWs.Range("J" & summary_row_table).Interior.ColorIndex = 3
                        ElseIf (Quarterly_Change = 0) Then
                            CurrentWs.Range("J" & summary_row_table).Interior.ColorIndex = 2
                        End If
                        
                        ' Use Conditional Formatting for percent change
                        If (Percent_Change > 0) Then
                            CurrentWs.Range("K" & summary_row_table).Interior.ColorIndex = 4
                        ElseIf (Percent_Change < 0) Then
                            CurrentWs.Range("K" & summary_row_table).Interior.ColorIndex = 3
                        ElseIf (Percent_Change = 0) Then
                            CurrentWs.Range("K" & summary_row_table).Interior.ColorIndex = 2
                        End If
                        
                  ' Add the value "1" to the summary row table
                  summary_row_table = summary_row_table + 1
                  Price_Row = i + 1
               
                  ' Reset the total stock volume
                  Total_Stock_Volume = 0
            Else
              Total_Stock_Volume = Total_Stock_Volume + CurrentWs.Range("G" & i).Value
                 
            End If
                      
        Next i
        
' ----------------------------------------
' Establish, calculate and define the greatest increase, decrease, and total volume
' ----------------------------------------

        ' Set the values in table for greatest % increase, greatest % decrease, and greatest total volume
        Greatest_Increase = CurrentWs.Range("K2").Value
        Greatest_Decrease = CurrentWs.Range("K2").Value
        Greatest_Total_Volume = CurrentWs.Range("L2").Value
        
        ' Define the last row of the ticker column
        Lastrow_Ticker = CurrentWs.Cells(Rows.Count, "I").End(xlUp).Row
        
        ' Calculate the greatest % increase, greatest % decrease, and the greatest total volume and loop through each row
         For r = 2 To Lastrow_Ticker
               If CurrentWs.Range("K" & r + 1).Value > Greatest_Increase Then
                  Greatest_Increase = CurrentWs.Range("K" & r + 1).Value
                  Greatest_Increase_Ticker = CurrentWs.Range("I" & r + 1).Value
               ElseIf CurrentWs.Range("K" & r + 1).Value < Greatest_Decrease Then
                  Greatest_Decrease = CurrentWs.Range("K" & r + 1).Value
                  Greatest_Decrease_Ticker = CurrentWs.Range("I" & r + 1).Value
                ElseIf CurrentWs.Range("L" & r + 1).Value > Greatest_Total_Volume Then
                  Greatest_Total_Volume = CurrentWs.Range("L" & r + 1).Value
                  Greatest_Total_Volume_Ticker = CurrentWs.Range("I" & r + 1).Value
                End If
            Next r
            
            ' Print greatest % increase, greatest % decrease, greatest total volume values
            CurrentWs.Range("P2").Value = Greatest_Increase_Ticker
            CurrentWs.Range("P3").Value = Greatest_Decrease_Ticker
            CurrentWs.Range("P4").Value = Greatest_Total_Volume_Ticker
            CurrentWs.Range("Q2").Value = Greatest_Increase
            CurrentWs.Range("Q3").Value = Greatest_Decrease
            CurrentWs.Range("Q4").Value = Greatest_Total_Volume
            CurrentWs.Range("Q2:Q3").NumberFormat = "0.00%"
            
    Next CurrentWs
    
End Sub

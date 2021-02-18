Sub Stockmarket_Data()
 
  ' LOOP THROUGH ALL SHEETS
  Dim ws As Worksheet
  For Each ws In Worksheets
    ' --------------------------------------------
    ' DEFINE THE VARIABLES
    ' --------------------------------------------
    ' Set an initial variable for holding the Ticker Symbol
    Dim Ticker_Symbol As String

    ' Set an initial variable for holding Opening Stock Price & Opening Stock Value
    Dim Opening_Price As Double
    Opening_Price = ws.Cells(2, 3).Value
  
    ' Set an initial variable for holding the Closing Stock Price & Closing Stock Value
    Dim Closing_Price As Double
    Closing_Price = 0

    ' Set an initial variable for holding the Yearly Change
    Dim Yearly_Change As Double

    ' Set an initial variable for holding the Percent Change
    Dim Percent_Change As Double

    ' Set an initial variable for holding the Total Volume of the Stock
    Dim Total_Stock_Volume As Double

    ' Set the Total Volume to zero
    Total_Stock_Volume = 0

    ' Set the Summary Table Row
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

    ' Set an intial variable for holding the Greatest Increase and Greatest Decrease and intitial values
    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    Dim Greatest_Decrease As Double
    Greatest_Decrease = 0
        
    ' Set an Intial Variable and value for Greatest Total Volume
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0 

    ' --------------------------------------------
    ' INSERT THE HEADERS
    ' --------------------------------------------
    ' Add the Headers
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    ' --------------------------------------------
    ' INSERT THE DATA
    ' --------------------------------------------
    ' Determine the Last Row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Loop through all of the Stock Transactions
    For i = 2 To LastRow

      ' Set Ticker Symbol
      Ticker_Symbol = ws.Cells(i, 1).Value

      ' Check if we are still within the same stock ticker symbol, if it is not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
        ' Print the Ticker Sumbol in the Summary Table
        ws.Cells(Summary_Table_Row, 9).Value = Ticker_Symbol

        ' Set the Closing Price
        Closing_Price = ws.Cells(i, 6).Value

        ' Add the Yearly Change
        Yearly_Change = Closing_Price - Opening_Price
        ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change

         ' Add Percent Change
        If Closing_Price = 0 Then
          Percent_Change = 0
        ws.Cells(Summary_Table_Row, 11).Value = Percent_Change  

        ElseIf (Opeing_Price = 0 And Closing_Price <> 0) Then
            Percent_Change = 1
            
        Else
          Percent_Change = Yearly_Change / Opening_Price
          ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
          ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
          
        End If

        ' Add to the Total Stock Volume amount
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume

        ' Set the cell colors
        If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
          ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
      
        Else
          ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3

        End If

        If Percent_Change > Greatest_Increase Then
          Greatest_Increase = Percent_Change
        End If

        If Percent_Change < Greatest_Decrease Then
          Greatest_Decrease = Percent_Change
        End If 

        If Total_Stock_Volume > Greatest_Total_Volume Then
          Greatest_Total_Volume = Total_Stock_Volume
        End If
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1

        ' Reset the Total Stock Volume
        Total_Stock_Volume = 0

        ' Reset the Opening Price
        Opening_Price = ws.Cells(i + 1, 3).Value
          
      ' If Cells have the same ticker symbol
      Else

        ' Total the Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
      End If

    Next i
    
    'Print the data to the desired cells with the proper formatting
    ws.Cells(2, 16).Value = Greatest_Increase
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = Greatest_Decrease
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = Greatest_Total_Volume

    ' Resize Columns
    ws.Columns("I:Q").EntireColumn.AutoFit
    
    ' Exit For

  'Do this For loop for all 3 worksheets  
  Next ws
      

End Sub

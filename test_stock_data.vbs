Sub testing()

' LOOP THROUGH ALL SHEETS

Dim ws As Worksheet


For Each ws In Worksheets

    ' DECLARE VARIABLES FOR TICKER, FILE NAME, OPEN PRICE, CLOSE PRICE
    
       
    Dim LastRow As Long
    
    Dim StockWorkSheet As String
    
    ' Determine the Last Row
    ' LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    


    ' VARIABLE FOR WORKSHEET

    StockWorkSheet = ws.Name
    MsgBox StockWorkSheet

    ' GRAB TICKER FROM BETWEEN DELIMITERS

    ' Set an initial variable for holding the ticker symbol
    Dim Ticker As String

    ' Set an initial variable for holding the yearly change per ticker
    Dim Yearly_Change As Double
    Yearly_Change = 0

   
    ' Declare total stock volume as variable for holding Total Stock Volume
    Dim Total_StockVolume As Double
    Total_StockVolume = 0
  
    ' Keep track of the location for each ticker symbol in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    ' Declare DfinRow as difference between opening and closing stock in each row
    Dim DFinRow As Double
    
  
    ' Declare Percent_Change as variable for holding yearly Percent_Change
    Dim Percent_Change As Double
    Percent_Change = 0
  
    ' Declare PerCalc as variable for percentage calculated between closing and opening stock per ticker symbol
    Dim PerCalc As Double
    
    Dim YrOp As Double
    Dim YrCl As Double
    'Dim DivErr As Double
       
    
    Dim i As Long
    
  
    ' Loop through all opening and closing stock per ticker symbol
    For i = 2 To LastRow

      ' Check if we are still within the same ticker symbol, if it is not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker Symbol
      Ticker = ws.Cells(i, 1).Value

      ' Add DfinRow to the Yearly_Change
      YrOp = ws.Cells(i + 1, 3).Value
      YrCl = ws.Cells(i + 1, 6).Value
      DFinRow = YrCl - YrOp
      Yearly_Change = Yearly_Change + DFinRow
       
      ' Add PerCalc to Percent_Change
      'DivErr = 0
      'PerCalc = FormatPercent(DFinRow / YrOp)
      'Percent_Change = Percent_Change + PerCalc
      
      ' Add to the Total_Stock_Volume
      Total_StockVolume = Total_StockVolume + ws.Cells(i, 7).Value

      ' Print Ticker_Symbol in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Ticker

      ' Print the Yearly_Change to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
      
      ' Print Percent_Change in the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Percent_Change
      
      ' Print Total_Stock_Volume to the Summary Table
      ws.Range("M" & Summary_Table_Row).Value = Total_StockVolume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Yearly_Change
      Yearly_Change = 0
      
      ' Reset the Percent_Change
      Percent_Change = 0
      
      ' Reset Total_Stock_Volume
      Total_StockVolume = 0
      
      
      

     ' If the cell immediately following a row is the Ticker_Symbol...
    Else
    
     Yearly_Change = Yearly_Change + DFinRow

      ' Add to the Yearly_Change
      ' Yearly_Change = Yearly_Change + DFinRow
      
      ' Add PerCalc to Percent_Change
      Percent_Change = Percent_Change + PerCalc
      
      'Add to the Total_Stock_Volume
      Total_StockVolume = Total_StockVolume + ws.Cells(i, 7).Value
      
      End If
      
    Next i
    
 Next ws


End Sub







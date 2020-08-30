Sub single_sheet_test()

Dim ws As Worksheet

' Loop through each sheet in Worksheets

For Each ws In Worksheets


   ' Set an initial variable for holding Ticker
   Dim Ticker As String

   ' Set Variable for last Row
   Dim LastRow As Long
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

   ' Set an initial variable for holding the total stock volume
   Dim total_stock_volume As Double
   total_stock_volume = 0

   ' Keep track of the location for total stock volume in the summary table
   Dim Summary_Table_Row As Long
   Summary_Table_Row = 2
  
   

   ' Declare and Set new variables for open, close price, difference in price, percent change and yearly change

   Dim YrOp As Double
   YrOp = 0
   Dim YrCl As Double
   YrCl = 0
   Dim DfInRow As Double
   DfInRow = 0
   Dim Percent_Change As Double
   Percent_Change = 0

   Dim i As Long

   ' Insert summary table in all worksheets

    If Summary_Table_Row Then



     ' Insert summary table in current worksheet

     ws.Range("J1").Value = "Ticker"
     ws.Range("K1").Value = "Yearly Change"
     ws.Range("L1").Value = "Percent Change"
     ws.Range("M1").Value = "Total Stock Volume"

    Else

        ' reset summary table for the other sheets

        Summary_Table_Row = True    

    End If

     ' Set value of initial open stock price
        
     YrOp = ws.Cells(2, 3).Value



    ' Loop through all stock
    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


             ' Set the Ticker Position
             Ticker = ws.Cells(i, 1).Value

             ' Calculate DfInRow and Percent Change

             YrCl = ws.Cells(i, 6).Value
             DfInRow = YrCl - YrOp

             ' condition if YrOp is not zero

            If YrOp <> 0 Then

                 Percent_Change = (DfInRow / YrOp) * 100

            End If  
    
       
      
             ' Add difference to total stock volume

             total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    
      
             ' Print the Ticker symbol in the summary table

             ws.Range("J" & Summary_Table_Row).Value = Ticker

             ' Print the yearly change in the summary table

             ws.Range("K" & Summary_Table_Row).Value = DfInRow

 
             ' Format yearly change cells to reflect positive(green) and negative(red) change

            If (DfInRow > 0) Then

                 ' Green for positive change

                 ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4

            ElseIf (DfInRow <= 0) Then

                 ' Red for zero value and negative change

                 ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3

            End If

             ' Insert percent change into summary table

             ws.Range("L" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")

             ' Insert total stock volume into summary table

             ws.Range("M" & Summary_Table_Row).Value = total_stock_volume
      
             ' Add 1 to the summary table row count

             Summary_Table_Row = Summary_Table_Row + 1

             ' Reset DfInRow to start with new cells

             DfInRow = 0

             ' Reset YrCl to calculate new DfInRow

             YrCl = 0

             ' Assign value to next open price

             YrOp = ws.Cells(i + 1, 3).Value

        Else
             ' Add to total stock volume

             total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        End If
            
      
    Next i


Next ws

End Sub



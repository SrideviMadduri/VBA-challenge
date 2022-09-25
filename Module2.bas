Attribute VB_Name = "Module2"
Sub stocks():

For Each ws In Worksheets

Dim WorksheetName As String


'Set an initial variable for holding the ticker symbol,
'total stock value per stock,closing price,Opening price
'Yearly Change ,Percentage change.
Dim Ticker_Symbol As String
Dim StockVolume_Total As Double
Dim Closing_Price As Double
Dim Opening_Price As Double
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Grt_Per_Increase_Ticker As String
Dim Grt_Per_Decrease_Ticker As String
Dim Grt_TotalVolume_Ticker As String
Dim Grt_Per_Increase_Value As Double
Dim Grt_Per_Decrease_Value As Double
Dim Grt_TotalVolume_Value As Double


'Set coloumn names for Ticker_Symbol, StockVolume_Total ,Opening_Price
'Closing_Price, Yearly_Change ,Percentage_Change, Value, Ticker.
 ws.Cells(1, 13).Value = "Ticker_Symbol"
 ws.Cells(1, 14).Value = "StockVolume_Total"
 ws.Cells(1, 15).Value = "Closing_Price"
 ws.Cells(1, 16).Value = "Opening_Price"
 ws.Cells(1, 17).Value = "Yearly_Change"
 ws.Cells(1, 18).Value = "Percentage_Change"
 ws.Cells(1, 21).Value = "Value"
 ws.Cells(1, 22).Value = "Ticker"
 
 
'Set initial variable

     StockVolume_Total = 0
     Grt_Per_Increase_Ticker = ""
     Grt_Per_Decrease_Ticker = ""
     Grt_TotalVolume_Ticker = ""
     Grt_Per_Increase_Value = 0
     Grt_Per_Decrease_Value = 100
     Grt_TotalVolume_Value = 0

Dim Ticker_Table_Row As Integer
Ticker_Table_Row = 2
Dim lRow As Long

'Determine the last row
lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Grabbed the WorksheetName
WorksheetName = ws.Name

'loop through all Ticker Symbols
    For I = 2 To lRow
        
        'Check if we are still in same ticker symbol. if not...
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

            'Set ticker symbol
            Ticker_Symbol = ws.Cells(I, 1).Value

            'Set Closing price
            Closing_Price = ws.Cells(I, 6).Value
            
            'Print the closing price in the worksheet
            ws.Range("O" & Ticker_Table_Row).Value = Closing_Price
            
            
            
            'return the stock with the Greatest % increase,
            'Greatest % decrease, and Greatest total volume
            
            If Percentage_Change > Grt_Per_Increase_Value Then
            
             Grt_Per_Increase_Value = Percentage_Change
             Grt_Per_Increase_Ticker = Ticker_Symbol
             
             
             End If
             
             If Percentage_Change < Grt_Per_Decrease_Value Then
             
             Grt_Per_Decrease_Value = Percentage_Change
             Grt_Per_Decrease_Ticker = Ticker_Symbol
             
              End If
              
              
            
            Yearly_Change = Closing_Price - Opening_Price
            ws.Range("Q" & Ticker_Table_Row).Value = Yearly_Change
            
         'Check if opening_price is 0 to avoid division by 0
        If ws.Range("P" & Ticker_Table_Row).Value = 0 Then
 
             'Set percentage_Change =0
             Percentage_Change = 0
        
        'Calculate Percentage change
        Else
            
            Percentage_Change = Yearly_Change / Opening_Price
            'Percentage_Change = ws.Range("Q" & Ticker_Table_Row).Value / ws.Range("P" & Ticker_Table_Row).Value

            'Percentage_Change = Format(Percentage_Change, "0.00%")
            'Print percentage change in worksheet
            ws.Range("R" & Ticker_Table_Row).Value = Percentage_Change

            'Changing the format to percentage
            ws.Range("R" & Ticker_Table_Row).Value = Format(ws.Range("R" & Ticker_Table_Row).Value, "0.00%")
        End If

         'Check if yearly_change is positive
        If ws.Range("Q" & Ticker_Table_Row).Value >= 0 Then
            
            'Highlight positive change in Green
            ws.Range("Q" & Ticker_Table_Row).Interior.Color = vbGreen
        
        'Yearly_change is negative
        Else
            
            'Highlight negative change to Red
            ws.Range("Q" & Ticker_Table_Row).Interior.Color = vbRed

        End If
            
            'Add the total stock volume
            StockVolume_Total = StockVolume_Total + ws.Cells(I, 7).Value

        If StockVolume_Total > Grt_TotalVolume_Value Then
             
                Grt_TotalVolume_Value = StockVolume_Total
                Grt_TotalVolume_Ticker = Ticker_Symbol
                
        End If
               
            'Print the Ticker Symbol in the worksheet
            ws.Range("M" & Ticker_Table_Row).Value = Ticker_Symbol
            
            

            'Print the Total stock volume per Ticker
            ws.Range("N" & Ticker_Table_Row).Value = StockVolume_Total


            'Add one to ticker symbol row
            Ticker_Table_Row = Ticker_Table_Row + 1

            'Reset the stock volume total
            StockVolume_Total = 0
            
        'If cell immediatly following a row has same ticker symbol
        Else
            'Add to the total Stock volume
            StockVolume_Total = StockVolume_Total + ws.Cells(I, 7).Value

        End If
        
        'Check if current and previous ticker are not same
        If ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then

            'Set opening price
            Opening_Price = ws.Cells(I, 3).Value

            'Print opening price in the worksheet
            ws.Range("P" & Ticker_Table_Row).Value = Opening_Price



        End If
        
        


Next I

    'Print
    ws.Range("U" & 2).Value = Grt_Per_Increase_Value
    ws.Range("V" & 2).Value = Grt_Per_Increase_Ticker

    ws.Range("U" & 3).Value = Grt_Per_Decrease_Value
    ws.Range("V" & 3).Value = Grt_Per_Decrease_Ticker

    ws.Range("U" & 4).Value = Grt_TotalVolume_Value
    ws.Range("V" & 4).Value = Grt_TotalVolume_Ticker

    ws.Range("T" & 2).Value = "Greatest % Increase"
    ws.Range("T" & 3).Value = "Greatest % Decrease"
    ws.Range("T" & 4).Value = "Greatest Total Volume"


Next ws

End Sub




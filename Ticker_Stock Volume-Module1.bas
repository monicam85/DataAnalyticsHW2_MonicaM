Attribute VB_Name = "Module1"
' Monica Martin Nov. 12, 2017
' Calculate and store Ticker Symbol and Total Stock Volume
Sub testdata()

'Set variable for holding Ticker Symbols
Dim Ticker As String

'Set initial variable for holding the Total Stock Volume

Dim Total_Stock_Volume As Double

Total_Stock_Volume = 0

'Helps with Dim Ticker_Row As Integer
Ticker_Row = 2


Cells(1, 9).Value = "Ticker"
      
Cells(1, 10).Value = "Total_Stock_Volume"

'Loop all Ticker symbols

For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


      ' Set the Ticker Name
      Ticker = Cells(i, 1).Value
      
      'Add the Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
      'Print the Ticker symbol to Column I (Starting at row2)
      Range("I" & Ticker_Row).Value = Ticker
      
      'Print the Ticker symbol to Column J (Starting at row2)
      Range("J" & Ticker_Row).Value = Total_Stock_Volume
      
      'Add one to the Ticker Row
      
      Ticker_Row = Ticker_Row + 1
      
      'Reset Stock Volume
      Total_Stock_Volume = 0
    
      
      Else
      
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      


    End If
    
   
    Next i

End Sub

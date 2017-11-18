Attribute VB_Name = "Module2"
' Monica Martin Nov. 12, 2017
' Calculate and store Ticker Symbol and Total Stock Volume
Sub Yearly_Percentage_StockVolume()

'Set variable for holding Ticker Symbols
Dim Ticker As String

'Set initial variable for holding the Total Stock Volume

Dim Total_Stock_Volume As Double

Total_Stock_Volume = 0

'Set Variable for Yearly Change
Dim sYear As Long
Dim Max_Year As Long
Dim Min_Year As Long


'Helps with Dim Ticker_Row As Integer
Ticker_Row = 2


Cells(1, 9).Value = "Ticker"
      
Cells(1, 12).Value = "Total_Stock_Volume"

Max_Year = Cells(2, 2).Value
Min_Year = Cells(2, 2).Value
sYear = Cells(2, 2).Value
Open_Value = Cells(2, 3).Value
Close_Value = Cells(2, 6).Value
Ticker = Cells(2, 1).Value
Cells(2, 9).Value = Ticker

'Loop all Ticker symbols

For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


        ' Set the Ticker Name
            Ticker = Cells(i + 1, 1).Value
      
        'Add the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
          
        'Set Percent_Change Value
        
          
        'Print the Ticker symbol to Column I (Starting at row2)
            Range("I" & Ticker_Row + 1).Value = Ticker
          
        'Print the Total_Stock_Volume to Column L (Starting at row2)
            Range("L" & Ticker_Row).Value = Total_Stock_Volume
          
        'Print the Yearly_Change to Column J (Starting at row2)
    
          
        'Print the Percent_Change to Column K (Starting at row2)
    
          
        'Add one to the Ticker Row
          
            Ticker_Row = Ticker_Row + 1
          
        'Reset Stock Volume
            Total_Stock_Volume = 0
          
            Range("J" & Ticker_Row).NumberFormat = "0.000000000"
        
      
      Else
      
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
        'Set Yearly Change Value
      
            sYear = Cells(i + 1, 2).Value
  
            If Max_Year < sYear Then
      
                Max_Year = sYear
      
                'Range("H" & Ticker_Row).Value = Max_Year
        
                Close_Value = Range("F" & i + 1).Value
        
            End If
        
        
        
    
            If Min_Year > sYear Then
      
                Min_Year = sYear
      
                'Range("M" & Ticker_Row).Value = Min_Year
        
                Open_Value = Range("C" & i + 1).Value
            
            End If
        End If
        
        If i >= 262 Then
          t = a
        End If
Next i

End Sub

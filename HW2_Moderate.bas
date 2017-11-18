Attribute VB_Name = "HW2_Moderate"
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
Dim ws As Worksheet

      
For Each ws In ActiveWorkbook.Worksheets
'Set the woorksheet
ws.Activate
'Clear output cells
ActiveSheet.Range("I:Q").Clear

'Helps with Dim Ticker_Row As Integer
Ticker_Row = 2


Cells(1, 9).Value = "Ticker"
      
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 11).Value = "Percet Change"

Cells(1, 10).Value = "Yearly Change"

Max_Year = Cells(2, 2).Value
Min_Year = Cells(2, 2).Value
sYear = Cells(2, 2).Value
open_value = Cells(2, 3).Value
close_value = Cells(2, 6).Value
Ticker = Cells(2, 1).Value
Cells(2, 9).Value = Ticker

'Loop all Ticker symbols

For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


        ' Set the Ticker Name
            Ticker = Cells(i + 1, 1).Value
      
        'Add the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                   
        'Print the Ticker symbol to Column I (Starting at row2)
            Range("I" & Ticker_Row + 1).Value = Ticker
          
        'Print the Total_Stock_Volume to Column L (Starting at row2)
            Range("L" & Ticker_Row).Value = Total_Stock_Volume
                   
        'Add one to the Ticker Row
          
            Ticker_Row = Ticker_Row + 1
          
        'Reset Stock Volume
            Total_Stock_Volume = 0
            Min_Year = Cells(Ticker_Row, 2).Value
            Max_Year = Cells(Ticker_Row, 2).Value
            sYear = Cells(Ticker_Row, 2).Value
          
            'Range("J" & Ticker_Row).NumberFormat = "0.000000000"
        
      
      Else
      
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
        'Set Yearly Change Value
      
            sYear = Cells(i + 1, 2).Value
  
            If Max_Year < sYear Then
      
                Max_Year = sYear
      
                'Range("H" & Ticker_Row).Value = Max_Year
        
                close_value = Range("F" & i + 1).Value
        
            End If
        
        
        
    
            If Min_Year > sYear Then
      
                Min_Year = sYear
      
                'Range("M" & Ticker_Row).Value = Min_Year
        
                open_value = Range("C" & i + 1).Value
            
            End If
            
                'Check for divide by 0
            If open_value > 0 Then
            
                Yearly_Change = close_value - open_value
        
                Percet_Diff = (open_value - close_value) / open_value
                
                Range("K" & Ticker_Row).Value = Percet_Diff
            
                Range("J" & Ticker_Row).Value = Yearly_Change
        
            Else
        
                Percet_Diff = 1
                
                Yearly_Change = close_value - open_value
            
                Range("K" & Ticker_Row).Value = Percet_Diff
                
                Range("J" & Ticker_Row).Value = Yearly_Change
            
          
             End If
        
        End If
        
       
Next i

Call Formatting

Next ws
End Sub

Sub Formatting()
    ActiveSheet.Range("J1:J" & Range("J" & Rows.Count).End(xlUp).Row).Select
    Selection.NumberFormat = "0.000000000"
    Selection.FormatConditions.Add xlCellValue, xlLess, "=0"
    Selection.FormatConditions.Add xlCellValue, xlLess, ">0"
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    
    With Selection.FormatConditions(2).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    
    Range("K:K").NumberFormat = "0.00%"
End Sub
